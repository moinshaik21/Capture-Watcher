"""
Control API for watcher_v5.py.

The API is the only thing that is "always on". Pressing Start in the web page makes this
API launch watcher_v5.py as a child process with `--since <chosen time>`, so filling begins
from that moment. Stop terminates it. watcher_v5.py itself is NOT modified.

Env vars:
  API_KEY          (required) secret the web page must send in the X-API-Key header
  WATCHER_TZ       time zone the script runs in, default Asia/Kolkata. The script writes
                   "local" times into Excel, and a cloud server is UTC by default.
  ALLOWED_ORIGINS  comma separated, e.g. https://your-site.netlify.app   (default: *)
  RENDER_EXTERNAL_URL / SELF_URL   public URL of this API, used for the keep-awake ping
"""
import collections
import json
import os
import re
import secrets
import subprocess
import sys
import threading
import time
import urllib.request
from contextlib import asynccontextmanager
from datetime import datetime, timezone
from pathlib import Path
from typing import Optional

from fastapi import Depends, FastAPI, Header, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

BASE = Path(__file__).resolve().parent
DATA = BASE / "data"                 # token_cache.bin, state.json, errors.log live here
DATA.mkdir(exist_ok=True)
SCRIPT = BASE / "watcher_v5.py"

API_KEY = os.environ.get("API_KEY", "")
WATCHER_TZ = os.environ.get("WATCHER_TZ", "Asia/Kolkata")
SELF_URL = (os.environ.get("RENDER_EXTERNAL_URL") or os.environ.get("SELF_URL") or "").rstrip("/")
ORIGINS = [o.strip().rstrip("/") for o in os.environ.get("ALLOWED_ORIGINS", "*").split(",") if o.strip()]
KEEPALIVE_SECONDS = 600              # free hosts sleep after ~15 min without inbound traffic

DEVICE_URL = re.compile(r"https://\S*devicelogin\S*")
DEVICE_CODE = re.compile(r"enter the code\s+([A-Z0-9]+)", re.I)
TIMESTAMPED = re.compile(r"^\[\d{4}-\d{2}-\d{2} ")


def iso(dt):
    return dt.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ") if dt else None


class Watcher:
    def __init__(self):
        self.lock = threading.RLock()
        self.proc = None
        self.logs = collections.deque(maxlen=1500)
        self.seq = 0
        self.started_at = None
        self.since = None
        self.login = None            # {"url":..., "code":...} while Microsoft sign-in is pending
        self.exit_code = None
        self._user_stop = None

    # ---- logging -------------------------------------------------------
    def push(self, text):
        with self.lock:
            self.seq += 1
            self.logs.append((self.seq, text))

    def snapshot(self, after):
        with self.lock:
            return [(n, t) for n, t in self.logs if n > after], self.seq

    # ---- process -------------------------------------------------------
    def is_running(self):
        with self.lock:
            return bool(self.proc and self.proc.poll() is None)

    def _reader(self, proc):
        for line in proc.stdout:
            line = line.rstrip("\n")
            if not line:
                continue
            self.push(line)
            with self.lock:
                if DEVICE_URL.search(line) and DEVICE_CODE.search(line):
                    self.login = {"url": DEVICE_URL.search(line).group(0),
                                  "code": DEVICE_CODE.search(line).group(1)}
                elif self.login and TIMESTAMPED.match(line):
                    self.login = None        # the script moved on -> sign-in finished
        code = proc.wait()
        with self.lock:
            by_user = self._user_stop is proc
            if self.proc is proc:
                self.proc, self.login = None, None
                self.exit_code = None if by_user else code
        self.push("[backend] watcher stopped" if by_user else f"[backend] watcher exited with code {code}")

    def start(self, since_iso):
        with self.lock:
            if self.is_running():
                return False
            env = {**os.environ, "TZ": WATCHER_TZ, "PYTHONUNBUFFERED": "1"}
            try:
                off = subprocess.run(
                    [sys.executable, "-c",
                     "from datetime import datetime;print(datetime.now().astimezone().strftime('UTC%z'))"],
                    env=env, capture_output=True, text=True, timeout=10).stdout.strip()
            except Exception:
                off = "?"
            self.push(f"[backend] starting watcher | mails received after {since_iso} | "
                      f"time zone {WATCHER_TZ} ({off})")
            proc = subprocess.Popen(
                [sys.executable, "-u", str(SCRIPT), "--since", since_iso],
                cwd=DATA, env=env, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                text=True, bufsize=1)
            self.proc, self.since, self.exit_code, self.login = proc, since_iso, None, None
            self.started_at = datetime.now(timezone.utc)
            threading.Thread(target=self._reader, args=(proc,), daemon=True).start()
            return True

    def stop(self):
        with self.lock:
            proc = self.proc
            if not proc or proc.poll() is not None:
                return False
            self._user_stop = proc
        proc.terminate()
        try:
            proc.wait(timeout=10)
        except subprocess.TimeoutExpired:
            proc.kill()
            proc.wait()
        return True

    def status(self):
        with self.lock:
            running = self.is_running()
            return {
                "running": running,
                "started_at": iso(self.started_at) if running else None,
                "since": self.since if running else None,
                "checkpoint": read_checkpoint(),
                "login": self.login if running else None,
                "last_exit_code": self.exit_code,
                "server_time": iso(datetime.now(timezone.utc)),
            }


def read_checkpoint():
    """The script keeps the receive-time of the last mail it handled in state.json ('since')."""
    try:
        return json.loads((DATA / "state.json").read_text()).get("since")
    except Exception:
        return None


watcher = Watcher()


def keepalive_loop():
    """Only while the watcher runs: ping our own public URL so the free host stays awake.
    When the watcher is stopped nothing pings and the host is free to sleep."""
    while True:
        time.sleep(KEEPALIVE_SECONDS)
        if SELF_URL and watcher.is_running():
            try:
                urllib.request.urlopen(SELF_URL + "/health", timeout=20).read()
            except Exception as e:
                watcher.push(f"[backend] keep-awake ping failed: {e}")


@asynccontextmanager
async def lifespan(_app):
    threading.Thread(target=keepalive_loop, daemon=True).start()
    yield
    watcher.stop()


app = FastAPI(lifespan=lifespan)
app.add_middleware(CORSMiddleware, allow_origins=ORIGINS, allow_methods=["*"], allow_headers=["*"])


def auth(x_api_key: str = Header(default="")):
    if not API_KEY:
        raise HTTPException(503, "API_KEY is not set on the server")
    if not secrets.compare_digest(x_api_key.encode(), API_KEY.encode()):
        raise HTTPException(401, "Wrong access key")


class StartBody(BaseModel):
    since: Optional[str] = None       # ISO date-time; empty = now


def normalize_since(value):
    if not value:
        dt = datetime.now(timezone.utc)
    else:
        try:
            dt = datetime.fromisoformat(value.strip().replace("Z", "+00:00"))
        except ValueError:
            raise HTTPException(400, "'since' must be an ISO date-time, e.g. 2026-10-08T09:30:00Z")
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
    return iso(dt)


@app.get("/")
def root():
    return {"service": "capture-watcher-api"}


@app.get("/health")
def health():
    return {"ok": True}


@app.get("/api/status", dependencies=[Depends(auth)])
def get_status():
    return watcher.status()


@app.post("/api/start", dependencies=[Depends(auth)])
def start(body: StartBody):
    since = normalize_since(body.since)
    started = watcher.start(since)
    return {"started": started, "already_running": not started, **watcher.status()}


@app.post("/api/stop", dependencies=[Depends(auth)])
def stop():
    stopped = watcher.stop()
    return {"stopped": stopped, **watcher.status()}


@app.get("/api/logs", dependencies=[Depends(auth)])
def get_logs(after: int = 0):
    lines, seq = watcher.snapshot(after)
    return {"lines": [{"n": n, "t": t} for n, t in lines], "seq": seq}

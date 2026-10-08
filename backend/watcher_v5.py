"""
Watches Outlook folder 'capturesAutomate' and writes 360-video captures
into the Excel sheet 'Detail Visual projects' (columns D:H) using Microsoft Graph.

Columns written:
  D = Projects        (from email subject)
  E = Structure Name  (from email table, e.g. Level 1)
  F = Date of Capture (from email table, Capture Time)
  G = Upload email time (email received time, local time)
  H = Capture path    (from email table)

Behaviour:
  * Every capture in a mail goes in its own row, one after another, below the last used row.
  * If captures already exist above, one blank row is left before the new mail's rows.
  * Colours alternate grey / blue: each mail group gets the opposite colour of the group above
    (set COLOR_BY = "row" to alternate every single row instead).
  * EVERY time a capture mail arrives the cycle is: find the real last filled row -> leave 1 blank row
    -> write the rows -> verify -> colour them (only rows that really hold data) -> remember the row.
  * Overlap protection: the last row we wrote is remembered in state.json ("last_row"). New rows are
    never written at or above it, even if Excel returns a stale / shorter used range. The target
    block is re-checked to be empty right before writing. A workbook session is used so reads see
    the latest writes.
  * Retries on throttling / locks / network errors, re-reads the sheet right before writing,
    verifies after writing, skips duplicates, and never lets one bad mail block the others.

Install:  pip install msal requests beautifulsoup4 python-dateutil
Run:      python watcher.py            (python watcher.py --check = read-only check of the last row)
"""
import base64
import json
import os
import random
import re
import sys
import time
from datetime import datetime, timezone
import msal
import requests
from bs4 import BeautifulSoup
from dateutil import parser as dtparser
from dateutil import tz

# ------------------------- CONFIG (edit these) -------------------------
CLIENT_ID = "77f0e163-94be-464c-a492-3fec90130b4d"
TENANT_ID = "69e91f03-f1cf-486d-8567-7d4914edffb2"      # or "organizations"
EXCEL_SHARE_URL = "https://constructn-my.sharepoint.com/:x:/g/personal/manoj_penmudi_track3d_ai/IQB2raqU0BBdSZEIJ_57fkayAW8l30HL8DayDQe4V7M2dq4?e=ZWJxJl"
SHEET_NAME = "Detail Visual projects"
MAIL_FOLDER_NAME = "capturesAutomate"
POLL_SECONDS = 60
ONLY_360 = True              # only rows whose Capture Type is a 360 VIDEO ("360 Video")
# 1) FROM: only mails sent from these addresses / domains are used.
#    An entry with "@" = that exact address (as shown in the mail: admin@track3d.ai);
#    an entry without "@" = a whole domain, e.g. "track3d.ai".
ALLOWED_SENDERS = ["admin@track3d.ai"]
# 2) TO: the To line must contain ONLY Tech ops (no other receiver). Matched against the
#    recipient name OR address, ignoring case/spaces/punctuation. Add the real address if you know it,
#    e.g. ["Tech ops", "techops@yourcompany.com"]. Empty list = do not check the To line.
#    (Cc is ignored.)
ALLOWED_TO = ["Tech ops","techops@track3d.ai"]
PALETTE = ["#D9D9D9", "#BDD7EE","#61A4F2", "#C379ED", "#E88BC7"]   # grey, blue - used alternately
COLOR_BY = "mail"                  # "mail" = one colour per mail group, "row" = alternate every row
COLOR_COLUMNS = ("A", "V")         # columns to colour on the filled rows, e.g. ("A", "H") for a wider band
HEADER_ROWS = 1                    # rows at the top that are headers (no blank gap after them)
GAP_ROWS = 1                       # blank rows left between the old data and a new mail's rows
MAX_MAIL_FAILURES = 3        # after this many failed attempts a mail is skipped + logged
WRITE_RETRIES = 5
SCAN_CHUNK = 1000            # rows read per step while scanning upwards for the last used row
LOOKBACK_ROWS = 500          # how many rows from the bottom are scanned for duplicates
BLANK_LOOKAHEAD = 5          # blank rows allowed between filled rows when looking below our last row
SHEET_CHECK_SECONDS = 30     # how often the sheet is checked for rows that changed
SNAPSHOT_ROWS = 2000         # how many rows from the bottom are watched for changes
STALE_RECHECKS = 3           # re-reads (2s apart) when the sheet shows FEWER rows than we already wrote
# -----------------------------------------------------------------------

SCOPES = ["Mail.Read", "Files.ReadWrite.All"]
GRAPH = "https://graph.microsoft.com/v1.0"
CACHE_FILE = "token_cache.bin"
STATE_FILE = "state.json"
ERROR_LOG = "errors.log"

LOCAL_TZ = tz.tzlocal()
RETRY_STATUS = {408, 409, 423, 429, 500, 502, 503, 504}


def log(msg):
    print(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {msg}", flush=True)


def log_error(msg):
    log("ERROR: " + msg)
    with open(ERROR_LOG, "a", encoding="utf-8") as f:
        f.write(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {msg}\n")


# ------------------------------ AUTH -----------------------------------
_cache = msal.SerializableTokenCache()
_app = None


def _get_app():
    global _app
    if _app is None:
        if os.path.exists(CACHE_FILE):
            _cache.deserialize(open(CACHE_FILE).read())
        _app = msal.PublicClientApplication(
            CLIENT_ID, authority=f"https://login.microsoftonline.com/{TENANT_ID}",
            token_cache=_cache)
    return _app


def get_token():
    app = _get_app()
    result = None
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
    if not result:
        flow = app.initiate_device_flow(scopes=SCOPES)
        print(flow["message"])          # open the link, enter the code, sign in
        result = app.acquire_token_by_device_flow(flow)
    if "access_token" not in result:
        raise RuntimeError(result.get("error_description", result))
    if _cache.has_state_changed:
        open(CACHE_FILE, "w").write(_cache.serialize())
    return result["access_token"]


_sess = {"id": None, "used": 0.0}


def ensure_session(wb):
    """Workbook session = consistent reads right after writes (no stale used-range / old rows)."""
    if _sess["id"] and time.time() - _sess["used"] < 240:
        return
    _sess["id"] = None
    try:
        r = g("POST", f"{wb}/createSession", json={"persistChanges": True})
        _sess["id"], _sess["used"] = r.get("id"), time.time()
    except Exception as e:
        log(f"  (workbook session not available, continuing without: {e})")


class GraphError(RuntimeError):
    def __init__(self, status, url, text):
        super().__init__(f"{status} {url}\n{text}")
        self.status = status


def g(method, url, retries=6, **kw):
    """Graph call with token refresh, throttling/lock retry and backoff."""
    extra = kw.pop("headers", {})
    last = None
    for attempt in range(retries):
        headers = {"Authorization": f"Bearer {get_token()}"}
        headers.update(extra)
        use_sess = bool(_sess["id"]) and "/workbook" in url and not url.endswith("/createSession")
        if use_sess:
            headers["workbook-session-id"] = _sess["id"]
            _sess["used"] = time.time()
        try:
            r = requests.request(method, url, headers=headers, timeout=60, **kw)
        except (requests.ConnectionError, requests.Timeout) as e:
            last = e
            time.sleep(min(2 ** attempt, 30))
            continue
        if r.status_code == 401:                       # token expired -> loop gets a fresh one
            last = GraphError(401, url, r.text)
            continue
        if r.status_code in RETRY_STATUS:              # throttled / file locked / busy
            wait = int(r.headers.get("Retry-After", min(2 ** attempt, 30)))
            last = GraphError(r.status_code, url, r.text)
            log(f"  Graph {r.status_code}, retrying in {wait}s")
            time.sleep(wait)
            continue
        if r.status_code >= 400:
            if use_sess and r.status_code in (400, 404) and "session" in r.text.lower():
                _sess["id"] = None                     # expired session -> a new one is made next time
            raise GraphError(r.status_code, url, r.text)
        return r.json() if r.text else {}
    raise last if last else RuntimeError("Graph call failed")


# ------------------------------ STATE ----------------------------------
def load_state():
    if os.path.exists(STATE_FILE):
        s = json.load(open(STATE_FILE))
        s.setdefault("fail", {})
        s.setdefault("pending_colour", [])
        s.setdefault("last_row", 0)          # last sheet row WE wrote (overlap guard)
        s.setdefault("last_key", None)       # [project, path] of that row
        return s
    # first run: only look at mails from now on (don't back-fill thousands of mails)
    return {"since": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
            "done": [], "fail": {}, "pending_colour": [],
            "last_row": 0, "last_key": None}


def save_state(state):
    state["done"] = state["done"][-2000:]
    tmp = STATE_FILE + ".tmp"
    json.dump(state, open(tmp, "w"), indent=1)
    os.replace(tmp, STATE_FILE)          # atomic: a crash can't corrupt state.json


# ------------------------------ OUTLOOK --------------------------------
def find_folder_id():
    url = f"{GRAPH}/me/mailFolders?$top=100"
    while url:
        data = g("GET", url)
        for f in data["value"]:
            if f["displayName"].lower() == MAIL_FOLDER_NAME.lower():
                return f["id"]
        url = data.get("@odata.nextLink")
    raise RuntimeError(f"Mail folder '{MAIL_FOLDER_NAME}' not found "
                       f"(sub-folders are not searched - check the name)")


def new_messages(folder_id, since):
    url = (f"{GRAPH}/me/mailFolders/{folder_id}/messages"
           f"?$filter=receivedDateTime ge {since}"
           f"&$orderby=receivedDateTime asc"
           f"&$select=subject,receivedDateTime,internetMessageId,body,from,sender,toRecipients&$top=50")
    out = []
    while url:
        data = g("GET", url, headers={"Prefer": 'outlook.body-content-type="html"'})
        out += data["value"]
        url = data.get("@odata.nextLink")
    return out


REQUIRED_COLS = ("structure name", "capture time", "capture path")


def _norm(text):
    return re.sub(r"[^a-z0-9@]", "", (text or "").lower())


def sender_allowed(msg):
    """FROM must be an allowed domain (or exact address)."""
    if not ALLOWED_SENDERS:
        return True
    ea = ((msg.get("from") or msg.get("sender") or {}).get("emailAddress") or {})
    addr = (ea.get("address") or "").strip().lower()
    for entry in ALLOWED_SENDERS:
        e = entry.strip().lower()
        if "@" in e:
            if addr == e:
                return True
        elif addr.endswith("@" + e) or addr.endswith("." + e):
            return True
    return False


def to_allowed(msg):
    """TO must contain only Tech ops: at least one recipient and every recipient is Tech ops."""
    if not ALLOWED_TO:
        return True
    recips = msg.get("toRecipients") or []
    if not recips:
        return False
    for r in recips:
        ea = r.get("emailAddress") or {}
        who = _norm(ea.get("name", "")) + "|" + _norm(ea.get("address", ""))
        if not any(_norm(a) and _norm(a) in who for a in ALLOWED_TO):
            return False
    return True


def skip_reason(msg):
    if not sender_allowed(msg):
        ea = ((msg.get("from") or {}).get("emailAddress") or {})
        return f"sender {ea.get('address')} is not from {ALLOWED_SENDERS}"
    if not to_allowed(msg):
        tos = ", ".join(((r.get("emailAddress") or {}).get("address") or "?")
                        for r in (msg.get("toRecipients") or [])) or "(empty)"
        return f"To line [{tos}] is not only {ALLOWED_TO}"
    return None


def _clean(text):
    """Collapse line breaks / double spaces: 'Capture\n Time' -> 'Capture Time'."""
    return re.sub(r"\s+", " ", (text or "")).strip()


def is_360_video(ctype):
    t = _clean(ctype).lower()
    return "360" in t and "video" in t


def parse_email(msg):
    """Return list of rows [D, E, F, G, H] - one per 360 capture in the mail."""
    subject = (msg.get("subject") or "").strip()
    m = re.search(r"captures of\s+(.+)$", subject, re.I)
    project = m.group(1).strip() if m else subject

    received = dtparser.isoparse(msg["receivedDateTime"]).astimezone(LOCAL_TZ)
    soup = BeautifulSoup((msg.get("body") or {}).get("content") or "", "html.parser")

    rows, skipped = [], 0
    for table in soup.find_all("table"):
        first_tr = table.find("tr")
        if not first_tr:
            continue
        header = [_clean(c.get_text(" ", strip=True)).lower() for c in first_tr.find_all(["th", "td"])]
        if "capture path" not in header:
            continue
        missing = [c for c in REQUIRED_COLS if c not in header]
        if missing:
            log_error(f"'{subject}': table is missing columns {missing}")
            continue
        ix = {name: i for i, name in enumerate(header)}
        for tr in table.find_all("tr")[1:]:
            try:
                cells = [_clean(c.get_text(" ", strip=True)) for c in tr.find_all(["td", "th"])]
                if len(cells) < len(header):
                    continue
                ctype = cells[ix["capture type"]] if "capture type" in ix else "360"
                if ONLY_360 and not is_360_video(ctype):
                    continue
                path = cells[ix["capture path"]].strip()
                if not path:
                    continue
                cap_time = dtparser.parse(cells[ix["capture time"]])
                if cap_time.tzinfo:
                    cap_time = cap_time.astimezone(LOCAL_TZ).replace(tzinfo=None)
                rows.append([
                    project,
                    cells[ix["structure name"]],
                    excel_serial(cap_time.replace(hour=0, minute=0, second=0, microsecond=0)),
                    excel_serial(received.replace(tzinfo=None)),
                    path,
                ])
            except Exception as e:                      # one bad row must not kill the mail
                skipped += 1
                log_error(f"'{subject}': skipped a row ({e}) -> {tr.get_text(' | ', strip=True)[:150]}")
    if skipped:
        log(f"  {skipped} row(s) skipped, see {ERROR_LOG}")
    return rows


def excel_serial(dt):
    delta = dt - datetime(1899, 12, 30)
    return delta.days + delta.seconds / 86400


# ------------------------------ EXCEL ----------------------------------
def resolve_workbook():
    enc = "u!" + base64.urlsafe_b64encode(EXCEL_SHARE_URL.encode()).decode().rstrip("=")
    item = g("GET", f"{GRAPH}/shares/{enc}/driveItem?$select=id,parentReference")
    wb = f"{GRAPH}/drives/{item['parentReference']['driveId']}/items/{item['id']}/workbook"
    try:                                               # fail early with a clear message
        g("GET", f"{wb}/worksheets('{SHEET_NAME}')?$select=name")
    except GraphError as e:
        if e.status == 404:
            raise RuntimeError(f"Sheet '{SHEET_NAME}' not found in the workbook")
        raise
    return wb


def ws_url(wb):
    return f"{wb}/worksheets('{SHEET_NAME}')"


def _scan_last_filled(ws):
    """Bottom-up scan of D:H for the REAL last used row (stray cells below the data are ignored)."""
    used = g("GET", f"{ws}/usedRange(valuesOnly=true)?$select=address")
    top = int(re.findall(r"(\d+)$", used["address"])[0])
    last_filled, hi = 0, top + 1
    while hi >= 1 and not last_filled:
        lo = max(1, hi - SCAN_CHUNK + 1)
        vals = g("GET", f"{ws}/range(address='D{lo}:H{hi}')?$select=values")["values"]
        for i in range(len(vals) - 1, -1, -1):
            if any(str(c).strip() for c in vals[i]):     # row is used if any of D:H has a value
                last_filled = lo + i
                break
        hi = lo - 1
    return last_filled


def _row_holds_key(ws, row, key):
    if not key or not row:
        return False
    cur = g("GET", f"{ws}/range(address='D{row}:H{row}')?$select=values")["values"][0]
    return str(cur[0]).strip() == key[0] and str(cur[4]).strip() == key[1]


def tail_snapshot(wb, state=None):
    """Find the REAL last filled row in D:H (re-checked on every call = every time a capture comes).
    Overlap guard: if the scan says the data ends ABOVE the row we last wrote, but that row still
    holds our capture, the scan was stale -> we trust our row and never write above it.
    Returns (next_row, existing {(project, path)}, last_filled_row, preview_text)."""
    ws = ws_url(wb)
    last_filled = _scan_last_filled(ws)

    hw = int((state or {}).get("last_row") or 0)
    hw_key = (state or {}).get("last_key")
    if hw and last_filled < hw:
        if _row_holds_key(ws, hw, hw_key):
            log(f"  scan found last row {last_filled} but row {hw} holds our last capture "
                f"-> stale read, re-scanning")
            for _ in range(STALE_RECHECKS):
                time.sleep(2)
                last_filled = max(last_filled, _scan_last_filled(ws))
                if last_filled >= hw:
                    break
            last_filled = max(last_filled, hw)           # never go back above our own last row
        else:
            log(f"  row {hw} no longer holds our last capture (rows removed/edited manually) "
                f"-> following the sheet, last row = {last_filled}")

    keys, preview = set(), "sheet is empty"
    if last_filled:
        k_lo = max(1, last_filled - LOOKBACK_ROWS + 1)
        vals = g("GET", f"{ws}/range(address='D{k_lo}:H{last_filled}')?$select=values")["values"]
        for r in vals:
            if str(r[0]).strip():
                keys.add((str(r[0]).strip(), str(r[4]).strip()))
        last = vals[-1]
        preview = f"{last[0]} | {last[1]} | {str(last[4])[-45:]}"
    return last_filled + 1, keys, last_filled, preview


def write_rows(wb, rows, state):
    """Append rows at the first free row. Returns (first_row, last_row) or None if all duplicates."""
    ws = ws_url(wb)
    ensure_session(wb)
    fmt = ["General", "General", "m/d/yyyy", "m/d/yyyy h:mm", "General"]
    for attempt in range(1, WRITE_RETRIES + 1):
        first, existing, last_filled, preview = tail_snapshot(wb, state)
        if last_filled > HEADER_ROWS:          # captures already above -> leave a blank row
            first += GAP_ROWS
        log(f"  PRE-CHECK: last used row = {last_filled} [{preview}], new rows will start at row {first}")
        todo = [r for r in rows if (r[0].strip(), r[4].strip()) not in existing]
        if len(todo) < len(rows):
            log(f"  {len(rows) - len(todo)} duplicate capture(s) skipped")
        if not todo:
            return None
        last = first + len(todo) - 1
        addr = f"D{first}:H{last}"

        # 1) gap row(s) + target block must still be empty (someone may have typed there meanwhile)
        check_from = last_filled + 1 if last_filled > HEADER_ROWS else first
        check_addr = f"D{check_from}:H{last}"
        cur = g("GET", f"{ws}/range(address='{check_addr}')?$select=values")["values"]
        if any(str(c).strip() for row in cur for c in row):
            log(f"  {check_addr} got occupied, re-scanning (attempt {attempt})")
            time.sleep(1 + random.random())
            continue

        # 2) write
        g("PATCH", f"{ws}/range(address='{addr}')",
          json={"values": todo, "numberFormat": [fmt] * len(todo)})

        # 3) read back and verify nothing else overwrote us (re-read a few times: reads can lag)
        ok = False
        for _ in range(3):
            back = g("GET", f"{ws}/range(address='{addr}')?$select=values")["values"]
            ok = all(str(b[0]).strip() == t[0].strip() and str(b[4]).strip() == t[4].strip()
                     for b, t in zip(back, todo))
            if ok:
                break
            time.sleep(2)
        if ok:
            log(f"  wrote {len(todo)} row(s) to {addr}")
            # remember our last row -> next mail can never be written at/above it
            state["last_row"] = last
            state["last_key"] = [todo[-1][0].strip(), todo[-1][4].strip()]
            save_state(state)
            # the blank separator row must stay blank AND uncoloured
            if first > last_filled + 1 and last_filled > HEADER_ROWS:
                try:
                    g("POST", f"{ws}/range(address='{COLOR_COLUMNS[0]}{last_filled + 1}:"
                              f"{COLOR_COLUMNS[1]}{first - 1}')/format/fill/clear")
                except Exception as e:
                    log(f"  could not clear fill on blank row: {e}")
            return first, last, last_filled
        log_error(f"{addr}: read-back mismatch (concurrent edit?) - retrying (attempt {attempt})")
        time.sleep(1 + random.random())
    raise RuntimeError("Could not write rows safely after several attempts")


def colour_rows(wb, first, last, prev_row):
    """Alternate grey / blue. The colour of the last filled row above decides what comes next,
    so the pattern survives restarts. COLOR_BY="mail": one colour per mail group,
    COLOR_BY="row": colours alternate on every row."""
    ws = ws_url(wb)
    ensure_session(wb)
    C1, C2 = COLOR_COLUMNS

    # 0) colour ONLY rows that are really filled (re-read a few times in case the read lags)
    filled = []
    for _ in range(STALE_RECHECKS + 1):
        vals = g("GET", f"{ws}/range(address='D{first}:H{last}')?$select=values")["values"]
        filled = [first + i for i, r in enumerate(vals) if any(str(c).strip() for c in r)]
        if filled:
            break
        time.sleep(2)
    if not filled:
        log_error(f"rows {first}-{last} are empty, nothing to colour (rows moved/deleted?) - job dropped")
        return
    runs, run = [], [filled[0]]                         # contiguous blocks of filled rows
    for r in filled[1:]:
        if r == run[-1] + 1:
            run.append(r)
        else:
            runs.append(run)
            run = [r]
    runs.append(run)
    first = filled[0]

    prev = None
    if prev_row > HEADER_ROWS:
        try:
            prev = (g("GET", f"{ws}/range(address='D{prev_row}')/format/fill").get("color") or "").lower()
        except Exception:
            prev = None
    pal = [c.lower() for c in PALETTE]
    # above is palette[0] -> start with palette[1]; anything else (white/none/other) -> palette[0]
    start_idx = 1 if prev == pal[0] else 0
    if prev == pal[1]:
        start_idx = 0

    if COLOR_BY == "row":
        for i, r in enumerate(filled):
            colour = PALETTE[(start_idx + i) % len(PALETTE)]
            g("PATCH", f"{ws}/range(address='{C1}{r}:{C2}{r}')/format/fill", json={"color": colour})
        expected = PALETTE[start_idx]
        log(f"  coloured {len(filled)} filled row(s) alternating by row")
    else:
        colour = PALETTE[start_idx]
        for rn in runs:                                  # blank rows are never coloured
            g("PATCH", f"{ws}/range(address='{C1}{rn[0]}:{C2}{rn[-1]}')/format/fill",
              json={"color": colour})
            log(f"  coloured {C1}{rn[0]}:{C2}{rn[-1]} {colour}")
        expected = colour

    # verify: first filled row must really carry the expected colour
    got = (g("GET", f"{ws}/range(address='{C1}{first}')/format/fill").get("color") or "").lower()
    if got != expected.lower():
        raise RuntimeError(f"colour check failed on row {first}: expected {expected}, got {got}")


# --------------------------- SHEET CHANGE CHECK ------------------------
_snap = {"rows": None, "top": 1, "bottom": 0, "rebase": False}


def _col_letter(i):
    n, out = i + 1, ""
    while n:
        n, rem = divmod(n - 1, 26)
        out = chr(65 + rem) + out
    return out


def _col_index(letter):
    n = 0
    for ch in letter.upper():
        n = n * 26 + ord(ch) - 64
    return n - 1


def _read_snapshot(ws):
    """Read the watched block (COLOR_COLUMNS wide, last SNAPSHOT_ROWS rows) -> (rows, top, bottom)."""
    C1, C2 = COLOR_COLUMNS
    used = g("GET", f"{ws}/usedRange(valuesOnly=true)?$select=address")
    bottom = int(re.findall(r"(\d+)$", used["address"])[0])
    top = max(1, bottom - SNAPSHOT_ROWS + 1)
    rows, lo = {}, top
    while lo <= bottom:
        hi = min(bottom, lo + SCAN_CHUNK - 1)
        vals = g("GET", f"{ws}/range(address='{C1}{lo}:{C2}{hi}')?$select=values")["values"]
        for i, r in enumerate(vals):
            t = tuple(str(c).strip() for c in r)
            if any(t):
                rows[lo + i] = t
        lo = hi + 1
    return rows, top, bottom


def lookahead_below(data_rows, ref):
    """Walk down from row `ref`, tolerating up to BLANK_LOOKAHEAD blank rows in a row.
    Returns (last_filled_row_reached, [filled rows found below ref])."""
    have, found, blanks, r = set(data_rows), [], 0, ref
    last = ref
    while blanks <= BLANK_LOOKAHEAD:
        r += 1
        if r in have:
            found.append(r)
            last, blanks = r, 0
        else:
            blanks += 1
    return last, found


def check_sheet_changes(wb, state=None):
    """Every SHEET_CHECK_SECONDS: compare the sheet with the previous check and print the ROW
    NUMBER of every row that was added / edited / cleared, plus the next row to be filled."""
    ws = ws_url(wb)
    ensure_session(wb)
    rows, top, bottom = _read_snapshot(ws)
    c0 = _col_index(COLOR_COLUMNS[0])
    d_lo, d_hi = _col_index("D") - c0, _col_index("H") - c0
    data_rows = [r for r, t in rows.items() if any(t[d_lo:d_hi + 1])]
    last_data = max(data_rows) if data_rows else 0
    next_row = last_data + 1 + (GAP_ROWS if last_data > HEADER_ROWS else 0)

    ref = int((state or {}).get("last_row") or 0)
    if ref and ref < last_data:
        reach, found = lookahead_below(data_rows, ref)
        if found:
            gap = found[0] - ref - 1
            log(f"SHEET CHECK: filled row(s) below our last written row {ref}: "
                f"{found[:10]}{' ...' if len(found) > 10 else ''} "
                f"(first one after {gap} blank row(s)); filled chain ends at row {reach}")
        if last_data > reach:
            log(f"SHEET CHECK: note - more data further below at row {last_data} "
                f"(more than {BLANK_LOOKAHEAD} blank rows after row {reach})")

    old = _snap["rows"]
    if old is None or _snap["rebase"]:
        _snap.update(rows=rows, top=top, bottom=bottom, rebase=False)
        log(f"SHEET CHECK: watching rows {top}-{bottom} | last filled row = {last_data} "
            f"| next row to fill = {next_row}")
        return
    changed = []
    for r in range(max(top, _snap["top"]), max(bottom, _snap["bottom"]) + 1):
        a, b = old.get(r, ()), rows.get(r, ())
        if a != b:
            width = max(len(a), len(b))
            a = a or ("",) * width
            b = b or ("",) * width
            cols = [_col_letter(c0 + i) for i in range(width) if a[i] != b[i]]
            kind = "NEW" if not any(a) else ("CLEARED" if not any(b) else "UPDATED")
            changed.append((r, kind, cols))
    _snap.update(rows=rows, top=top, bottom=bottom)
    if changed:
        for r, kind, cols in changed[:30]:
            log(f"SHEET CHECK: row {r} {kind} (columns {', '.join(cols)})")
        if len(changed) > 30:
            log(f"SHEET CHECK: ... and {len(changed) - 30} more changed rows")
        log(f"SHEET CHECK: last filled row = {last_data} | next row to fill = {next_row}")
    else:
        log(f"SHEET CHECK: no change | last filled row = {last_data} | next row to fill = {next_row}")


# ------------------------------- MAIN ----------------------------------
def run_colour_jobs(wb, state):
    """Colouring is MANDATORY: every written block is stored as a pending job (state.json) and is
    retried until the colour is applied AND verified. Jobs run in order so alternation stays right."""
    while state["pending_colour"]:
        ensure_session(wb)
        job = state["pending_colour"][0]
        try:
            colour_rows(wb, job["first"], job["last"], job["prev_row"])
        except Exception as e:
            job["tries"] = job.get("tries", 0) + 1
            save_state(state)
            log_error(f"colouring rows {job['first']}-{job['last']} not done yet "
                      f"(try {job['tries']}): {e}")
            return False
        state["pending_colour"].pop(0)
        save_state(state)
    return True


def process_message(wb, msg, state):
    reason = skip_reason(msg)
    if reason:
        log(f"Skipped: {msg.get('subject')} -> {reason}")
        return
    rows = parse_email(msg)
    log(f"Mail: {msg.get('subject')} -> {len(rows)} 360 row(s)")
    if not rows:
        return
    result = write_rows(wb, rows, state)
    if result:
        first, last, prev_row = result
        _snap["rebase"] = True                       # our own write: don't report it as an outside edit
        # store the colour job BEFORE colouring so a crash/failure can never lose it
        state["pending_colour"].append({"first": first, "last": last, "prev_row": prev_row})
        save_state(state)
        run_colour_jobs(wb, state)


def precheck(wb, state=None):
    """Read-only check: shows where the next captures would be written. Run: python watcher.py --check"""
    ensure_session(wb)
    first, existing, last_filled, preview = tail_snapshot(wb, state)
    start = first + GAP_ROWS if last_filled > HEADER_ROWS else first
    colour = None
    if last_filled > HEADER_ROWS:
        colour = g("GET", f"{ws_url(wb)}/range(address='D{last_filled}')/format/fill").get("color")
    log(f"CHECK: last used row = {last_filled} [{preview}] (fill {colour}); "
        f"blank row = {first if start != first else '-'}; next captures start at row {start}")
    return start


def main():
    get_token()
    folder_id = find_folder_id()
    wb = resolve_workbook()
    state = load_state()
    if "--since" in sys.argv:                       # e.g. --since 2026-10-06T18:00:00Z
        state["since"] = sys.argv[sys.argv.index("--since") + 1]
        save_state(state)
    precheck(wb, state)
    if "--check" in sys.argv:
        return
    log(f"Watching folder: {MAIL_FOLDER_NAME} | since {state['since']}")

    next_mail = next_sheet = 0.0
    while True:
        now = time.time()
        if now >= next_mail:
            mail_cycle(wb, folder_id, state)
            next_mail = time.time() + POLL_SECONDS
        if time.time() >= next_sheet:
            try:
                check_sheet_changes(wb, state)
            except Exception as e:
                log_error(f"sheet check failed: {e}")
            next_sheet = time.time() + SHEET_CHECK_SECONDS
        time.sleep(1)


def mail_cycle(wb, folder_id, state):
    """One mail poll (unchanged logic)."""
    try:
        # unfinished colouring from an earlier cycle must be completed first
        if not run_colour_jobs(wb, state):
            return
        for msg in new_messages(folder_id, state["since"]):
            mid = msg["internetMessageId"]
            if mid in state["done"]:
                continue
            try:
                process_message(wb, msg, state)
                state["done"].append(mid)
                state["fail"].pop(mid, None)
            except Exception as e:
                n = state["fail"].get(mid, 0) + 1
                state["fail"][mid] = n
                log_error(f"mail '{msg.get('subject')}' failed ({n}/{MAX_MAIL_FAILURES}): {e}")
                if n >= MAX_MAIL_FAILURES:       # give up on this one, don't block the queue
                    state["done"].append(mid)
                    log_error(f"giving up on '{msg.get('subject')}' - add it manually")
                else:
                    break                        # keep order; retry this mail next poll
            state["since"] = msg["receivedDateTime"]
            save_state(state)
            if state["pending_colour"]:          # colour not applied yet -> stop, retry next cycle
                break
    except Exception as e:
        log_error(f"poll cycle failed: {e}")


if __name__ == "__main__":
    main()

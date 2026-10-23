# Capture Watcher

A lightweight web-based control panel for managing an automated email-to-spreadsheet watcher. The application provides **Start**, **Stop**, and live log monitoring through a Netlify-hosted frontend, with a DRF backend deployed on Render.

## Architecture

```text
┌──────────────────────────┐
│     Netlify Frontend     │
│     frontend/index.html  │
│                          │
│  Start | Stop | Logs     │
└────────────┬─────────────┘
             │ HTTPS
             ▼
┌──────────────────────────┐
│     Render Backend       │
│       backend/app.py     │
│                          │
│  API Authentication      │
│  Watcher Process Control │
└────────────┬─────────────┘
             │
             ▼
┌──────────────────────────┐
│      watcher_v5.py       │
│                          │
│  Email Monitoring        │
│  Spreadsheet Updates     │
└──────────────────────────┘
```

## Key Features

- **Remote Control:** Start and stop the watcher through a browser-based interface.
- **Configurable Start Time:** Monitor emails received from the current time or a user-selected historical start time.
- **Time-Based Filtering:** Passes the selected timestamp to `watcher_v5.py` using the `--since` argument.
- **Live Logs:** View watcher activity and execution status from the frontend.
- **API Authentication:** Protect backend endpoints using an API key.
- **Cross-Origin Protection:** Restrict frontend access through the `ALLOWED_ORIGINS` configuration.
- **Cloud Deployment:** Host the frontend on Netlify and the backend on Render.
- **Original Script Preserved:** Uses the existing `watcher_v5.py` without modifying its implementation.

## Technology Stack

| Component | Technology |
|---|---|
| Frontend | HTML, CSS, JavaScript |
| Frontend Hosting | Netlify |
| Backend API | Python, DRF |
| ASGI Server | Uvicorn |
| Backend Hosting | Render |
| Watcher | `watcher_v5.py` |
| Version Control | GitHub |

## Project Structure

```text
capture-watcher/
├── frontend/
│   └── index.html
├── backend/
│   ├── app.py
│   ├── requirements.txt
│   └── ...
├── watcher_v5.py
├── render.yaml
└── README.md
```

*Note: Adjust the structure above to match the actual location of your watcher script and deployment files.*

## Prerequisites

Before deployment, ensure you have:

- A GitHub account and a private repository.
- A Netlify account for frontend hosting.
- A Render account for backend hosting.
- The existing `watcher_v5.py` script and its required dependencies.
- Access to the mailbox and spreadsheet used by the watcher.
- The necessary authentication and permissions for accessing the mailbox and spreadsheet.

## 1. Push the Project to GitHub

1. Create a **private GitHub repository**.
2. Add the project files and commit your changes.
3. Push the project to the repository.
4. Verify that credentials, access tokens, secrets, and other sensitive configuration values are not committed to Git.

**Security note:** The watcher may contain tenant IDs, client IDs, SharePoint URLs, or other configuration details. Keep the repository private and move sensitive values into environment variables wherever possible. Never commit client secrets, access tokens, passwords, or API keys.

## 2. Deploy the Backend on Render

### Option A: Deploy using a Blueprint

1. Sign in to [Render](https://render.com/).
2. Select **New → Blueprint**.
3. Connect your private GitHub repository.
4. Allow Render to read the `render.yaml` configuration.
5. Review the service configuration and deploy.

### Option B: Create a Web Service manually

Configure the service as follows:

| Setting | Value |
|---|---|
| Root Directory | `backend` |
| Build Command | `pip install -r requirements.txt` |
| Start Command | `uvicorn app:app --host 0.0.0.0 --port $PORT` |
| Instance Plan | Free |

Ensure the watcher script and all required files are accessible from the backend's execution environment.

### Configure Environment Variables

Add the following variables in the Render dashboard.

| Variable | Description |
|---|---|
| `API_KEY` | A strong, randomly generated secret used to authenticate API requests. |
| `WATCHER_TZ` | The timezone used for timestamp handling, such as `Asia/Kolkata`. |
| `ALLOWED_ORIGINS` | The exact URL of your deployed Netlify frontend, such as `https://my-watcher.netlify.app`. |

Use the actual deployed frontend URL for `ALLOWED_ORIGINS`. If the backend supports a specific timezone format or has a documented fallback, follow its implementation when configuring `WATCHER_TZ`.

After deployment, copy the Render service URL. For example:

`https://capture-watcher-api.onrender.com`

Replace this example with your actual backend URL.

## 3. Deploy the Frontend on Netlify

1. Open [Netlify](https://app.netlify.com/).
2. Deploy the `frontend` directory using Netlify's manual deployment option, or connect the GitHub repository.
3. If using Git-based deployment, configure the base directory as `frontend`.
4. Wait for the deployment to complete.
5. Open your deployed website.

### Connect the Backend

In the frontend's connection settings:

1. Enter the Render backend URL.
2. Enter the `API_KEY` configured in Render.
3. Save the settings and connect to the backend.

The frontend stores these settings in the browser, according to the application's current implementation.

**Important:** A browser-stored API key is accessible to users who can access that browser session and inspect the frontend. Use this approach only for a trusted, restricted-access control page. For stronger security, use server-side authentication or another access-control mechanism rather than relying solely on a shared API key embedded in browser-accessible settings.

## 4. Start and Manage the Watcher

Once the frontend and backend are connected, follow these steps.

### Step 1: Configure the Start Time

Choose one of the available options:

- **From now:** Start monitoring from the selected current timestamp.
- **From a set time:** Choose a specific start timestamp.

### Step 2: Start Watching

Click **Start Watching**.

The frontend sends the start request to the backend. If the Render service is sleeping, the request may initially take longer while the service wakes up.

The backend then launches `watcher_v5.py` with the selected timestamp through the `--since` argument.

Only emails that satisfy the script's configured date and filtering logic are eligible for processing.

### Step 3: Complete Authentication

During the first run, or after a server restart when authentication must be renewed:

1. Check the sign-in card displayed by the application.
2. Open the provided authentication link.
3. Enter the displayed code.
4. Sign in using the authorized mailbox account.
5. Allow the watcher to continue once authentication is complete.

Authentication behavior depends on the existing watcher and its configured authentication flow.

### Step 4: Monitor Activity

Use the log panel to follow execution messages, processing activity, and errors reported by the application.

### Step 5: Stop Watching

Click **Stop** when monitoring is no longer required.

The backend requests the watcher process to stop. Confirm the status in the frontend logs before assuming the process has terminated.

## 5. Start-Time and Email Processing Behavior

The selected start timestamp is passed to the watcher using this command-line pattern:

```bash
python watcher_v5.py --since "<selected_timestamp>"
```

The exact timestamp format depends on the backend implementation.

The watcher is responsible for applying the timestamp filter and updating the target spreadsheet. The frontend provides the control interface, while the backend manages the watcher process.

**Note:** Starting from a historical timestamp may process older emails that meet the watcher's filtering rules. Verify the script's duplicate-handling and spreadsheet-update logic before using historical start times.

## 6. Troubleshooting

| Issue | Recommended Action |
|---|---|
| Backend is unavailable | Check the Render deployment status and service logs. |
| First request is slow | Allow time for the Render free service to wake up. |
| API authentication fails | Verify that the frontend API key matches the Render `API_KEY`. |
| Browser reports a CORS error | Check that `ALLOWED_ORIGINS` contains the exact Netlify site URL. |
| Watcher fails to start | Review backend logs, the Python executable, script path, and installed dependencies. |
| Mailbox authentication is required | Complete the authentication flow displayed by the frontend. |
| Incorrect email timestamps | Verify timezone configuration and timestamp conversion in the backend and watcher. |
| Spreadsheet is not updated | Check mailbox access, spreadsheet permissions, watcher logs, and the script's filtering rules. |
| Stop does not terminate the watcher | Inspect the process-management implementation and confirm that child processes are handled correctly. |

## 7. Operational and Security Considerations

- Keep the GitHub repository private if it contains internal project configuration.
- Never expose client secrets, refresh tokens, or passwords in frontend code.
- Use HTTPS for communication between Netlify and Render.
- Restrict API access and validate all incoming requests on the backend.
- Configure appropriate process cleanup and error handling for watcher startup and shutdown.
- Check Render's filesystem and service lifecycle limitations before relying on local authentication caches or generated files.
- The Render free service may sleep when idle, and its available runtime resources and execution limits may affect long-running watcher processes.
- Verify that the watcher remains active for the expected duration under the selected hosting plan.

## 8. Summary

Capture Watcher provides a simple browser-based interface for controlling an existing automated email-monitoring workflow.

By combining a Netlify frontend with a DRF backend on Render, it enables users to configure a monitoring start time, launch or stop the watcher, complete mailbox authentication, and inspect execution logs without manually starting the Python script from a terminal.

The original `watcher_v5.py` remains responsible for the core email-processing and spreadsheet-update logic.

---

**Project:** Capture Watcher  
**Frontend:** Netlify  
**Backend:** Render + DRF  
**Watcher Script:** `watcher_v5.py`

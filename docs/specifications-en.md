# Functional, Technical Specifications & Screen Guide — SyncContactsOG

- **Current Version** : `1.0.11`
- **Last Updated** : `2026-11-22`
- **Target Platform** : Windows 10 / Windows 11 (x64)
- **Tech Stack** : Python 3.12, CustomTkinter, PyWin32 (MAPI/COM), Microsoft Graph API, Google People API, SQLite, OpenPyXL, PyInstaller

---

## 1. Global Objective & Scope

### 1.1 Purpose
**SyncContactsOG** is a portable Windows desktop application designed to reliably synchronize, unify, and archive address books between **Microsoft Outlook** (Classic desktop client or Microsoft 365 cloud account) and **Google Contacts**.

### 1.2 Scope
- **In-Scope**:
  - One-way and two-way synchronization of personal contacts.
  - High data fidelity preservation (structured names, `FileAs`, multiple phone numbers, postal addresses, rich RTF notes, categories/labels, and high-resolution contact photos).
  - Browsable local archiving (SQLite database with binary photos + `.xlsx` Excel workbook).
  - Interactive execution (GUI) and background execution (Windows System Tray / Windows startup).
- **Out-of-Scope**:
  - Synchronization of calendars, tasks, or emails.
  - Simultaneous multi-account synchronization within a single running instance.

---

## 2. Actors & Execution Modes

| Execution Mode | Trigger | Expected Behavior |
| :--- | :--- | :--- |
| **Interactive (GUI)** | Manual launch of `SyncContactsOG.exe` | Displays the main window, allowing configuration, simulation (*Dry-Run*), and manual sync execution. |
| **Background (Tray)** | Closing window (when enabled) or `--minimized` CLI flag | Minimizes the app to the Windows System Tray and runs scheduled background syncs at configured intervals. |

---

## 3. Functional Requirements

- **FR-01 — Dual-Source Connectivity**:
  - *Outlook*: Supports local **Outlook Classic (COM/MAPI)** and cloud **Microsoft 365 (Graph API OAuth2)** connectors.
  - *Google*: Connects via **Google People API (OAuth2)** with automatic silent token refresh.
- **FR-02 — Synchronization Modes**:
  - `OUTLOOK_TO_GOOGLE_ONLY`: One-way propagation from Outlook to Google (production-validated mode).
  - `GOOGLE_TO_OUTLOOK_ONLY`: One-way propagation from Google to Outlook.
  - `MERGE_OUTLOOK_WINS`: Two-way sync where Outlook overwrites Google on conflicting changes.
  - `MERGE_GOOGLE_WINS`: Two-way sync where Google overwrites Outlook on conflicting changes.
  - `MERGE_PROMPT`: Two-way sync opening an interactive conflict resolution dialog for each conflict.
- **FR-03 — Advanced Field Preservation**:
  - Optional enforcement of `FileAs` formatting (`Name SURNAME` or `SURNAME, Name`).
  - Protection of rich-text (RTF) Outlook notes against plain-text overwrites from Google unless explicitly forced.
  - Two-way contact photo sync using SHA-256 hash comparison.
  - Bidirectional mapping between Outlook categories and Google Contact labels (groups).
- **FR-04 — Deletion Safety & Simulation**:
  - **Dry-Run Mode**: Runs full diff analysis and preview logging without modifying any data.
  - Automatic safety abort if deletion count exceeds `max_deletions_threshold` or if a source unexpectedly returns 0 contacts when the mapping database is non-empty.
  - Automatic JSON backup of deleted contacts into `backups/` prior to deletion.
- **FR-05 — Resilient Local Archiving (`.db` + `.xlsx`)**:
  - Automatically created inside `<exe_folder>\contacts\` when write permissions are available.
  - Interactive folder picker fallback if the executable directory is read-only.
  - Free disk space check prior to writing: if disk space is insufficient, existing archive files are safely renamed with a `_NOT_UP_TO_DATE` suffix.

---

## 4. Functional & Technical Assumptions

### 4.1 Functional Assumptions
- Every synchronized contact is persistently paired via the `(outlook_entry_id, google_resource_name)` tuple.
- Incremental (*Delta*) sync runs by default; a *Full Sync* runs manually or automatically whenever the last full sync is older than 7 days.
- The user has a configured local Outlook profile (for COM mode) or HTTPS network access to `graph.microsoft.com` and `people.googleapis.com`.

### 4.2 Technical Assumptions
- **Zero-Pollution Portable Architecture**: `SyncContactsOG.exe` is a self-contained `--onefile` binary. No technical files are written into the executable's directory.
- **Technical Data Isolation**: Sync state (`sync_state.db`), logs (`sync_details.log`), and deletion backups (`backups/`) reside in `%LOCALAPPDATA%\SyncContactsOG\`.
- **Credential Security**: OAuth client secrets (`credentials.json`) are embedded into the binary at build time, and user OAuth tokens are encrypted inside **Windows Credential Manager**. App preferences are persisted in the Windows Registry (`HKCU\Software\OutlookGoogleSyncPro`).
- **Single-Instance Enforcement**: A system-wide Windows Mutex (`Local\SyncContactsOG_SingleInstance_Mutex`) prevents multiple instances or versions from running simultaneously.

---

## 5. Google Authentication (`credentials.json` & OAuth2 Token)

### 5.1 Role of the Two Authentication Artifacts
1. **OAuth Client Configuration (`credentials.json`)**:
   - Holds the Google Cloud Console `client_id` and `client_secret` authorized for **Google People API** (`https://www.googleapis.com/auth/contacts` scope).
   - **Standard Build (`SyncContactsOG.exe`)**: `credentials.json` is securely bundled inside the executable at compile time. End users do not need to supply or manage any `.json` file.
   - **No-Key Build (`SyncContactsOG_NoKey.exe`)**: Designed for custom Google Cloud projects; prompts the user to select a Desktop App `credentials.json` file on first login and stores it securely.
2. **Access & Refresh Token (`token`)**:
   - Generated automatically after the user grants consent in their web browser when clicking `Test Connection` or `Sync Now`.
   - **Secure Storage**: The OAuth2 token payload (including `refresh_token`) is stored encrypted inside **Windows Credential Manager** under a dedicated key bound to `SyncContactsOG` and the connected Google account. No plaintext `token.json` file is ever written to disk.
   - **Silent Renewal**: Unless revoked from Google Account security settings, expired access tokens are refreshed automatically in the background without user intervention.

### 5.2 Google Login Window & Workflow
1. Enter your Gmail or Google Workspace address in the **Google Account** field and click **`Test Connection`** (or start a sync directly).
2. If no valid token exists for that email address, the app opens your **default web browser** to the official Google sign-in page (`accounts.google.com`) while listening on a temporary local callback port (`localhost`).
3. In the browser: select the matching Google account, approve access to Google Contacts (*Google People API*), and close the browser tab once the confirmation message appears.
4. Back in **SyncContactsOG**: the email field locks to read-only mode to prevent accidental edits, and the connection status turns to **`✔ Connected`**.
5. **Switching or Resetting Accounts**: Click the **`✕`** button on the right side of the Google email field to unlock the input, clear the current session, and bind a different Google account.

### 5.3 Pre-Build Credential Resolution & Validation Chain (`build.ps1`)
When building a standard release (`SyncContactsOG.exe`), `build.ps1` locates and validates `credentials.json` online against Google's OAuth2 servers (`validate_creds.py`) in the following priority order, automatically skipping any corrupted or revoked file:
1. `credentials.json` placed next to the main executable (`$ProjectDir\credentials.json`).
2. Local backup (`C:\Temp\credentials_backup.json`).
3. Extracted from the existing main executable (`SyncContactsOG.exe`).
4. Extracted from previous releases (`releases\SyncContactsOG_v*.exe`, newest to oldest).

---

## 6. Architecture & Execution Flow

1. **Pre-flight Checks**: Single-instance Mutex check, email validation, and archive folder write-permission verification.
2. **Collection & Canonicalization**: Reads Outlook and Google contacts, normalizes every record into a canonical schema, and computes its SHA-256 hash.
3. **Diff Calculation (Delta / Full)**: Compares current hashes against `sync_state.db` (`contact_map`) to classify creations, updates, conflicts, and deletions.
4. **Execution & Live Logging**: Applies API changes (with exponential backoff on HTTP 429 rate limits) and streams color-coded logs to the right console.
5. **Post-Processing**: Commits updates to `sync_state.db`, checks free disk space, and atomically writes `contacts_archive.db` and `contacts_archive.xlsx`.

---

## 7. Screen Specifications & Detailed User Guide

### 7.1 Main Window Overview
The main window uses a fixed two-column layout topped by a header bar:
- **Header Bar**: Displays the application title, version (`vX.Y.Z`), language selector (`FR` / `EN`), and theme switcher (`Dark` / `Light`). Changes persist immediately.
- **Left Column (Controls & Settings)**: Contains 5 functional cards (*Accounts & Profiles*, *Sync Mode*, *Options & Safety*, *Windows Integration*, *Exclusive Contacts*).
- **Right Column (Execution Console)**: Displays real-time progress and detailed operation logs.

### 7.2 Card-by-Card & Modal Screen Details

#### A. Card 1 — Accounts & Profiles
- **Outlook Account**: Dropdown auto-detecting local Outlook profiles.
- **Connector Selector (Radio Buttons)**:
  - *Outlook Classic COM*: Connects directly to the locally installed Outlook desktop client and displays the live contact count in parentheses.
  - *Microsoft 365 Graph API*: Connects to Microsoft 365 cloud via browser OAuth2 and displays the cloud contact count.
- **Google Account**: Input box for Gmail/Google Workspace (locks to read-only once authenticated, with a `✕` reset button).
- **`Test Connection` Button**: Triggers the OAuth2 flow described in Section 5 and displays `✔ Connected`.
- **Auto-Sync (hours)**: Numeric spinbox (`0` to `168`). `0` disables scheduled sync (manual only); any positive value $N$ runs automatic sync every $N$ hours.

#### B. Card 2 — Sync Mode
- **5 Sync Modes (Radio Buttons with Fixed-Width Status Pills)**:
  - Each row displays the mode name, explanatory subtitle, and a uniform status pill (`✔ Tested` or `⚠ Untested`) with a hover tooltip.
- **`Sync Now` Primary Button**: Starts the synchronization worker thread (disabled while running to prevent duplicate triggers).

#### C. Card 3 — Options & Safety
- **Use FileAs field**: Enables synchronization of Outlook's `FileAs` display field.
- **Enforce FileAs format**: Forces `Name SURNAME` or `SURNAME, Name` formatting on synced contacts.
- **Force overwrite RTF notes**: Disabled by default (and greyed out in *Outlook → Google* mode) to protect rich-text Outlook notes from plain-text overwrites.
- **Sync photos**: Enables SHA-256 comparison and two-way profile picture sync.
- **Sync deletions**: Propagates contact deletions across platforms, guarded by **Max deletions threshold** to block accidental mass deletions.
- **Visual confirmation before deletion**: Prompts a confirmation modal before each deletion.
- **Force Full Compare**: Bypasses incremental cache and compares all fields across every contact.
- **Local Archive (`.db` + `.xlsx`)**: Generates and updates the SQLite photo database and Excel workbook inside `contacts\`.
- **Dry-Run (Simulation Mode)**: Runs full analysis and prints the preview summary **without modifying any contact** in Outlook or Google.

#### D. Cards 4a & 4b — Windows Integration & Exclusive Contacts
- **Minimize to System Tray**: Clicking the window close button (`✕`) hides the app into the Windows notification area instead of exiting. Double-clicking the tray icon restores the window; right-clicking provides *Open*, *Sync Now*, and *Exit*.
- **Run at Windows Startup**: Registers `SyncContactsOG.exe --minimized` under `HKCU\Software\Microsoft\Windows\CurrentVersion\Run` for silent startup in the System Tray.
- **Exclusive Contacts & Force Mirror**: Displays counts of contacts existing on only one platform. In one-way modes, *Force Mirror* strictly mirrors the source onto the target by removing extra target contacts.

#### E. Right Panel — Execution Console
- **Search & Filter Bar**: Instantly filters log entries by keyword or severity (`Information`, `Warning`, `Error`).

#### F. Modal Dialogs (Secondary Screens)
1. **Conflict Resolution Dialog (`MERGE_PROMPT`)**: Appears when a contact was modified on both sides since the last sync. Displays Outlook vs. Google values side by side and lets you keep *Outlook*, *Google*, or *Skip*.
2. **Deletion Confirmation Dialog**: Displays the contact details prior to deletion and asks for confirmation (*Yes* / *No*).
3. **Archive Folder Picker (Read-Only Fallback)**: If `SyncContactsOG.exe` runs from a read-only folder while Local Archive is enabled, prompts the user to pick a writable destination folder.

### 7.3 Typical User Workflows
1. **First Launch & Safe Dry-Run**:
   - Launch `SyncContactsOG.exe` $\rightarrow$ verify detected Outlook profile $\rightarrow$ enter Google email and click *Test Connection* $\rightarrow$ select *Outlook → Google (One-Way)* $\rightarrow$ enable *Dry-Run (Simulation Mode)* $\rightarrow$ click *Sync Now* and review the preview summary in the right console.
2. **Live Sync & Background Automation**:
   - Uncheck *Dry-Run* $\rightarrow$ click *Sync Now* $\rightarrow$ set *Auto-Sync (hours)* to `4` $\rightarrow$ enable *Minimize to System Tray* and *Run at Windows Startup*.
3. **Browsing Local Archives**:
   - Open `contacts\` next to the executable to view `contacts_archive.xlsx` in Excel or query `contacts_archive.db`.

---

## 8. Non-Functional Requirements & Security

- **API Resilience**: Automatic handling of transient network errors and Google People API rate limits using exponential backoff.
- **Transactional Integrity**: SQLite archive updates are written to a temporary file before atomic replacement (`shutil.move`) to prevent corruption.
- **Bilingual UI**: Dynamic French (`FR`) and English (`EN`) interface and tooltips, with Dark and Light themes.

---

## 9. Sync Mode Validation Matrix (Current Status)

| Synchronization Mode | Internal Code | Validation Status | UI Badge |
| :--- | :--- | :--- | :--- |
| Outlook → Google (One-Way) | `OUTLOOK_TO_GOOGLE_ONLY` | Validated and tested in production | `✔ Tested` |
| Google → Outlook (One-Way) | `GOOGLE_TO_OUTLOOK_ONLY` | Pending field validation | `⚠ Untested` |
| Two-Way (Outlook priority) | `MERGE_OUTLOOK_WINS` | Pending field validation | `⚠ Untested` |
| Two-Way (Google priority) | `MERGE_GOOGLE_WINS` | Pending field validation | `⚠ Untested` |
| Two-Way (Prompt on conflict) | `MERGE_PROMPT` | Pending field validation | `⚠ Untested` |
# Architecture & Engineering Best Practices (Blueprint v1.1.1)

This document defines the architectural, security, UI/UX, build, and GitHub release standards established for **SyncContactsOG v1.1.1**. It is designed both as the technical reference for this repository and as a reusable blueprint for future portable Windows desktop applications built with Python.

---

## 1. Standard GitHub Repository Structure

A clean separation between application code, developer utilities, documentation, user runtime data, and release artifacts ensures maintainability and prevents accidental secret leaks.

    ProjectRoot/
    ├── .gitignore                      # Strictly excludes data/, contacts/, secrets, caches, binaries
    ├── CHANGELOG.md                    # Single source of truth for version history (Keep a Changelog)
    ├── LICENSE                         # Open-source license (e.g., GPL-3.0)
    ├── README.md                       # English entry point with architecture overview & doc links
    ├── requirements.txt                # Pinned minimum Python dependencies
    ├── source/                         # Pure application source code (compiled by PyInstaller)
    │   ├── SyncContactsOG.py           # Main entry point & UI/Sync engine
    │   ├── SyncContactsOG.ico          # Application icon embedded in EXE & Toplevel dialogs
    │   └── build.ps1                   # Isolated build & packaging pipeline
    ├── tools/                          # Standalone developer & documentation automation scripts
    │   ├── generate_specs_from_code.py # Extracts constants/modes from source to update specs
    │   ├── convert_md_2_html.py        # Renders Markdown docs into styled offline HTML files
    │   └── validate_credentialsJSON.py # CLI validator for Google OAuth credentials.json
    ├── docs/                           # Lowercase documentation assets (Markdown + generated HTML)
    │   ├── best-practices.md           # This engineering blueprint (Markdown source)
    │   ├── best-practices.html         # Auto-generated HTML for offline viewing / GitHub Pages
    │   ├── specifications-en.md        # Functional & technical specifications (English)
    │   ├── specifications-en.html      # Auto-generated HTML (English)
    │   ├── specifications-fr.md        # Functional & technical specifications (French)
    │   ├── specifications-fr.html      # Auto-generated HTML (French)
    │   └── changelog.html              # Auto-generated HTML from root CHANGELOG.md
    ├── data/                           # [GITIGNORED] Portable OAuth client config (token stored in HKCU DPAPI)
    │   └── credentials.json            # User's Google Cloud OAuth Desktop client secrets (token is in HKCU DPAPI)
    ├── contacts/                       # [GITIGNORED] Portable local backup snapshots
    │   ├── contacts.db                 # SQLite relational snapshot & delta sync cache
    │   └── contacts.xlsx               # Human-readable Excel export
    └── releases/                       # [GITIGNORED] Local release staging archives
        └── v1.1.1/
            ├── SyncContactsOG-v1.1.1-win64.zip
            ├── SyncContactsOG-v1.1.1-source.zip
            ├── SHA256SUMS.txt
            └── RELEASE_NOTES.md

### Naming Conventions
- **Repository Root**: Uppercase for standard GitHub meta files (`README.md`, `CHANGELOG.md`, `LICENSE`) so they stand out at the top of the file tree.
- **`docs/` Folder**: Lowercase kebab-case (`best-practices.md`, `specifications-en.md`) for clean URLs on GitHub and GitHub Pages. Markdown (`.md`) is the single source of truth; HTML (`.html`) files are generated automatically for offline distribution.
- **`tools/` vs `source/`**: Keep build, validation, and documentation generators in `tools/` so PyInstaller only bundles runtime dependencies from `source/`.

---

## 2. Portable Application Architecture

To run without administrator rights, MSI installers, or registry modifications:

1. **Bundled Assets (`_MEIPASS`)**: Read-only assets packaged inside the executable (such as `.ico` icons) must be resolved via `getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))`.
2. **Persistent User Data (`sys.executable` directory)**: Writable runtime folders (`.\data\`, `.\contacts\`) and configuration files (`config.json`) must always resolve relative to `os.path.dirname(sys.executable)` when frozen (`getattr(sys, "frozen", False)`), or `os.path.dirname(os.path.abspath(__file__))` when run from source.
3. **Zero OS Keyring Dependency**: Never rely on Windows Credential Manager (`keyring`) for portable apps; moving the folder to another PC would break authentication. Store validated JSON credentials and tokens inside `.\data\`.

---

## 3. OAuth 2.0 & API Security Practices (Google People API v1)

1. **Pre-Flight JSON Schema Validation**: Before initiating any OAuth flow, parse `.\data\credentials.json` and verify that it contains a valid `"installed"` or `"web"` object with non-empty `"client_id"` and `"client_secret"` keys.
2. **Encrypted Windows Registry Token Storage (`HKCU\Software\SyncContactsOG` + DPAPI)**:
   - Encrypt the serialized token JSON using Windows DPAPI (`win32crypt.CryptProtectData`) bound to the current Windows user session **and** an application-specific entropy (`SyncContactsOG::OAuth2::DPAPI::v1`), storing the resulting `REG_BINARY` blob in `HKEY_CURRENT_USER\Software\SyncContactsOG\GoogleOAuthTokenDPAPI` (no administrator rights required).
   - Load and decrypt via `CryptUnprotectData`, refresh silently via `creds.refresh(Request())`, or trigger `InstalledAppFlow.run_local_server(port=0)` on first sign-in.
3. **Minimal Least-Privilege Scope & Verification**:
   - Request only `https://www.googleapis.com/auth/contacts`.
   - Verify connectivity by calling `service.people().connections().list(resourceName="people/me", pageSize=1, personFields="names").execute()` rather than `people().get(resourceName="people/me")`, which requires an extra `profile` scope and triggers `HttpError 403`.
4. **Repository Secret Hygiene**:
   - `.gitignore` must explicitly block `data/`, `contacts/`, `config.json`, `credentials*.json`, `*.db`, and `*.xlsx`.
   - Release ZIP archives must include only neutral `README.txt` placeholders inside empty `data/` and `contacts/` folders.

---

## 4. Desktop UI/UX & Threading Standards (CustomTkinter)

1. **Non-Blocking Background Workers**: All network calls (OAuth browser flow, Google People API pagination) and COM operations (Outlook contact scanning) must run inside daemon `threading.Thread` workers. UI updates (progress bars, log insertion, button state changes) must be marshaled back to the main thread via `self.after(0, callback)`.
2. **High-DPI & Grid Layout Discipline**:
   - Never assign `weight=1` to fixed-height header or progress bar rows in `grid_rowconfigure()`. Always set `weight=0` on status/progress containers and `weight=1` exclusively on expandable log or data viewers.
   - Avoid hardcoded pixel heights on multi-line text containers; use `wraplength` and natural padding so text never clips at 125% or 150% Windows DPI scaling.
3. **Actionable Visual Feedback**:
   - Highlight missing prerequisites (e.g., missing `credentials.json` or unconnected Google account) using primary action colors (blue `#1a73e8`), and transition buttons to neutral styling once verified alongside a clear status indicator (`✓` / `✗`).
   - Keep modal configuration dialogs open after importing credentials so the user can observe live connection diagnostics before closing.
4. **Dynamic Localization (EN / FR)**: Centralize all UI labels in a translation dictionary and bind language toggles to a refresh method that updates widget text in real time without requiring an application restart.

---

## 5. Data Integrity & Synchronization Safeguards

- **Incremental Delta Caching**: Compute deterministic hashes of contact fields and photos in SQLite (`.\contacts\contacts.db`) so subsequent runs only process modified records, while offering a *Force Full Compare* override.
- **Pre-Write Snapshots**: Automatically export both relational (`contacts.db`) and tabular (`contacts.xlsx`) backups before applying destructive changes.
- **Deletion Guardrails**: Enforce a configurable deletion threshold alert and an interactive confirmation prompt (`Prompt Deletion`), plus a zero-contact safety abort if a source unexpectedly returns `0` records when linked contacts exist.
- **Dry-Run Simulation Mode**: Support full comparison and logging without executing write/delete mutations on Outlook or Google.

---

## 6. Build, Packaging & GitHub Release Pipeline (`source/build.ps1`)

1. **Syntax & Static Gate**: Run `py_compile` on `source/SyncContactsOG.py` before invoking PyInstaller to fail fast on syntax errors.
2. **Out-of-Tree Compilation**: Configure PyInstaller `--workpath`, `--distpath`, and `--specpath` inside `C:\Temp\Build_<App>_v<Version>\`. Compiling outside OneDrive/cloud-synced folders prevents file-locking errors (`PermissionDenied`) and keeps the working tree clean.
3. **Standardized GitHub Release Assets**:
   - `SyncContactsOG-vX.Y.Z-win64.zip`: Portable executable, `README.md`, `CHANGELOG.md`, `LICENSE`, `docs/`, and clean `data/` + `contacts/` folders.
   - `SyncContactsOG-vX.Y.Z-source.zip`: Clean `source/`, `tools/`, `docs/`, `requirements.txt`, `.gitignore`, `README.md`, `CHANGELOG.md`, and `LICENSE`.
   - `SHA256SUMS.txt`: Lowercase SHA-256 digests for all release archives.

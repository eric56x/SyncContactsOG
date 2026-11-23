# SyncContactsOG (v1.1.1)

**SyncContactsOG** is a modern, 100% portable Windows desktop application for synchronizing contacts between **Microsoft Outlook (Classic COM)** and **Google Contacts (Google People API v1)**.

Inspired by **GO Contact Sync Mod**, this iteration provides native **Google People API v1** support, portable OAuth client configuration (`.\data\credentials.json`) with Windows DPAPI-encrypted user registry token storage (`HKCU\Software\SyncContactsOG`), local SQLite/Excel snapshots (`.\contacts\`), and a bilingual (**EN / FR**) Light/Dark UI.

## Documentation
- [Changelog](CHANGELOG.md) ([HTML](docs/changelog.html))
- [Specifications (English)](docs/specifications-en.md) ([HTML](docs/specifications-en.html))
- [Spécifications (Français)](docs/specifications-fr.md) ([HTML](docs/specifications-fr.html))
- [Architecture & Best Practices](docs/best-practices.md) ([HTML](docs/best-practices.html))

## Quick Start
1. Extract `SyncContactsOG-v1.1.1-win64.zip` and run `SyncContactsOG.exe`.
2. Click **Configure Google Connection** (*Configurer la Connexion Google* in FR).
3. Import your Google Cloud OAuth Desktop `credentials.json` into `.\data\` and click **Connect Google Account**.
4. Select your Outlook profile and sync mode, then click **Run Now**.

## Build from Source

    pip install -r requirements.txt
    .\source\build.ps1

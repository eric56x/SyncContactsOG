import ctypes
try:
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID('WePoint.SyncContactsOG.App.1')
except Exception:
    pass
import os
import re
import sys
import json
import time
import base64
import shutil
import winreg
import hashlib
import sqlite3
import tempfile
import threading
import urllib.request
from datetime import datetime
import tkinter as tk
from tkinter import messagebox
import customtkinter as ctk

import pystray
from PIL import Image, ImageDraw
import openpyxl

import pythoncom
import win32com.client
import win32cred
from google.auth.transport.requests import Request
from google.auth.exceptions import RefreshError
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

__version__ = "1.1.1"

# ==============================================================================
# GESTION DES CHEMINS (EXECUTABLE PARTAGEABLE SANS FICHIERS JSON EXTERNES)
# ==============================================================================
if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
    BUNDLE_DIR = sys._MEIPASS
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    BUNDLE_DIR = BASE_DIR

# Ressources embarquées dans l'EXE (via --add-data)
CREDENTIALS_FILE = os.path.join(BUNDLE_DIR, "credentials.json")
ICON_FILE = os.path.join(BUNDLE_DIR, "SyncContactsOG.ico")

# Stockage technique dans %LOCALAPPDATA%\SyncContactsOG (zero pollution a cote de l'EXE)
APP_DATA_DIR = os.path.join(
    os.environ.get("LOCALAPPDATA") or os.path.expanduser("~"),
    "SyncContactsOG"
)
os.makedirs(APP_DATA_DIR, exist_ok=True)

# Migration silencieuse des anciens fichiers techniques s'ils etaient a cote de l'EXE
for _legacy_name in ("sync_state.db", "sync_details.log", "backups"):
    _src_leg = os.path.join(BASE_DIR, _legacy_name)
    _dst_leg = os.path.join(APP_DATA_DIR, _legacy_name)
    if os.path.exists(_src_leg) and not os.path.exists(_dst_leg):
        try:
            shutil.move(_src_leg, _dst_leg)
        except Exception:
            pass

DB_FILE = os.path.join(APP_DATA_DIR, "sync_state.db")
LOG_FILE = os.path.join(APP_DATA_DIR, "sync_details.log")
BACKUP_DIR = os.path.join(APP_DATA_DIR, "backups")


def is_dir_writable(dir_path: str) -> bool:
    try:
        os.makedirs(dir_path, exist_ok=True)
        probe = os.path.join(dir_path, f".write_test_{os.getpid()}.tmp")
        with open(probe, "w", encoding="utf-8") as tf:
            tf.write("ok")
        os.remove(probe)
        return True
    except Exception:
        return False


def has_enough_disk_space(target_dir: str, required_bytes: int) -> bool:
    try:
        check_path = target_dir if os.path.exists(target_dir) else os.path.dirname(os.path.abspath(target_dir))
        usage = shutil.disk_usage(check_path)
        return usage.free >= required_bytes
    except Exception:
        return True


def mark_archive_not_up_to_date(archive_dir: str) -> list:
    renamed = []
    pairs = [
        ("contacts.db", "contacts_NOT_UP_TO_DATE.db"),
        ("contacts.xlsx", "contacts_NOT_UP_TO_DATE.xlsx"),
    ]
    for cur_name, stale_name in pairs:
        cur_p = os.path.join(archive_dir, cur_name)
        stale_p = os.path.join(archive_dir, stale_name)
        if os.path.exists(cur_p):
            try:
                if os.path.exists(stale_p):
                    os.remove(stale_p)
                os.replace(cur_p, stale_p)
                renamed.append(stale_name)
            except Exception:
                pass
    return renamed


def restore_stale_archive_if_needed(archive_dir: str):
    pairs = [
        ("contacts_NOT_UP_TO_DATE.db", "contacts.db"),
        ("contacts_NOT_UP_TO_DATE.xlsx", "contacts.xlsx"),
    ]
    for stale_name, cur_name in pairs:
        stale_p = os.path.join(archive_dir, stale_name)
        cur_p = os.path.join(archive_dir, cur_name)
        if os.path.exists(stale_p) and not os.path.exists(cur_p):
            try:
                os.replace(stale_p, cur_p)
            except Exception:
                pass

# Coffre-fort Windows & Registre utilisateur (zéro fichier JSON sur le disque)
CRED_TARGET_NAME = "OutlookGoogleSyncPro_GoogleOAuth"
GRAPH_CRED_TARGET_NAME = "OutlookGoogleSyncPro_GraphOAuth"
REG_SETTINGS_PATH = r"Software\OutlookGoogleSyncPro"
APP_REG_NAME = "SyncContactsOG"

SCOPES = ["https://www.googleapis.com/auth/contacts"]


def detect_default_outlook_email() -> str:
    """Détecte automatiquement l'adresse e-mail du profil Outlook par défaut sur le PC."""
    pythoncom.CoInitialize()
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        if namespace.Accounts.Count > 0:
            smtp = getattr(namespace.Accounts.Item(1), "SmtpAddress", "")
            if smtp and "@" in smtp:
                return smtp.strip()
        for store in namespace.Stores:
            if "@" in store.DisplayName:
                return store.DisplayName.strip()
    except Exception:
        pass
    finally:
        pythoncom.CoUninitialize()
    return ""


DEFAULT_CONFIG = {
    "language": "FR",
    "theme": "Dark",
    "outlook_email": "",
    "outlook_connector": "COM",
    "google_email": "",
    "sync_mode": "OUTLOOK_TO_GOOGLE_ONLY",
    "use_file_as": True,
    "force_file_as_format": False,
    "file_as_format_style": "Name SURNAME",
    "force_rtf": True,
    "sync_photos": True,
    "sync_deletions": True,
    "prompt_deletion": True,
    "force_full_compare": False,
    "extract_local_db": False,
    "last_archive_date": "",
    "last_full_archive_date": "",
    "dry_run": False,
    "auto_sync_hours": 6,
    "max_deletions_threshold": 10,
    "minimize_to_tray": True,
    "start_with_windows": False
}


# ==============================================================================
# COFFRE-FORT WINDOWS (WINDOWS CREDENTIAL MANAGER) POUR LES JETONS OAUTH
# ==============================================================================
def save_secret_to_vault(target_name: str, secret_str: str):
    cred = {
        "Type": win32cred.CRED_TYPE_GENERIC,
        "TargetName": target_name,
        "UserName": "SyncContactsOGUser",
        "CredentialBlob": secret_str,
        "Persist": win32cred.CRED_PERSIST_LOCAL_MACHINE,
    }
    win32cred.CredWrite(cred, 0)


def load_secret_from_vault(target_name: str):
    try:
        cred = win32cred.CredRead(target_name, win32cred.CRED_TYPE_GENERIC, 0)
        blob = cred["CredentialBlob"]
        if isinstance(blob, bytes):
            return blob.decode("utf-16-le").rstrip("\x00")
        return blob
    except Exception:
        return None


def delete_secret_from_vault(target_name: str):
    try:
        win32cred.CredDelete(target_name, win32cred.CRED_TYPE_GENERIC, 0)
    except Exception:
        pass


# ==============================================================================
# DICTIONNAIRE BILINGUE (FRANCAIS / ANGLAIS) + INFOBULLES
# ==============================================================================
I18N = {
    "FR": {
        "app_title": "SyncContactsOG",
        "theme_light": "Clair",
        "theme_dark": "Sombre",
        "sec_accounts": "Comptes & Profils",
        "lbl_ol_email": "Compte Outlook (Store) :",
        "lbl_g_email": "Compte Google :",
        "lbl_auto_sync": "Auto-Sync toutes les (heures, 0=off) :",
        "tip_auto_sync": "Lance automatiquement la synchronisation à intervalle régulier tant que l'application est active (0 pour désactiver).",
        "sec_mode": "Mode de Synchronisation",
        "mode_prompt": "Merge Prompt",
        "mode_prompt_sub": "Demander en cas de conflit",
        "tip_mode_prompt": "Bidirectionnel : affiche une fenêtre de comparaison pour choisir le gagnant si un contact a été modifié des deux côtés.",
        "mode_ol_wins": "Merge Outlook Wins",
        "mode_ol_wins_sub": "Bidirectionnel - Outlook prioritaire",
        "tip_mode_ol_wins": "Bidirectionnel : synchronise dans les deux sens, mais Outlook écrase Google en cas de modification simultanée.",
        "mode_g_wins": "Merge Google Wins",
        "mode_g_wins_sub": "Bidirectionnel - Google prioritaire",
        "tip_mode_g_wins": "Bidirectionnel : synchronise dans les deux sens, mais Google écrase Outlook en cas de modification simultanée.",
        "mode_ol_to_g": "Outlook To Google Only",
        "mode_ol_to_g_sub": "Unidirectionnel",
        "tip_mode_ol_to_g": "Unidirectionnel strict : Outlook est la source de vérité. Aucune modification n'est jamais écrite dans Outlook.",
        "mode_g_to_ol": "Google To Outlook Only",
        "mode_g_to_ol_sub": "Unidirectionnel",
        "tip_mode_g_to_ol": "Unidirectionnel strict : Google est la source de vérité. Aucune modification n'est jamais écrite dans Google.",
        "sec_opts": "Options & Intégrité des données",
        "opt_photos": "Synchroniser les photos des contacts",
        "tip_opt_photos": "Synchronise les portraits haute définition entre Outlook et Google Contacts.",
        "opt_file_as": "Utiliser « Classer sous » Outlook pour le titre Google",
        "tip_opt_file_as": "Utilise le champ « Classer sous » d'Outlook comme nom d'affichage principal dans Google Contacts.",
        "opt_force_fa_fmt": "Forcer le format Classer sous :",
        "tip_opt_force_fa_fmt": "Normalise automatiquement le champ Classer sous (FileAs) des contacts sources selon le format choisi (Prénom NOM ou NOM, Prénom).",
        "opt_force_rtf": "Force RTF Contacts (Écrase les notes RTF Outlook)",
        "tip_opt_force_rtf": "Autorise le remplacement des notes Outlook mises en forme (RTF) par du texte brut venant de Google.",
        "opt_del": "Synchroniser les suppressions",
        "tip_opt_del": "Si un contact déjà synchronisé est supprimé d'un côté, il sera également supprimé de l'autre côté.",
        "opt_prompt_del": "Confirmer avant toute suppression",
        "tip_opt_prompt_del": "Demande une confirmation visuelle avant d'exécuter toute suppression unitaire de contact.",
        "opt_force_full": "Forcer une synchro complète (sinon delta par défaut, full >7j)",
        "tip_opt_force_full": "Si coché : Full Sync Outlook + Google + Archive (.db et .xlsx). Sinon : synchro incrémentale rapide par défaut et bascule automatique en Full Sync si la dernière Full Sync date de > 7 jours.",
        "opt_extract_db": "Conserver une Archive Locale (.db avec photos + .xlsx)",
        "tip_opt_extract_db": "Maintient contacts.db (avec photos) et contacts.xlsx : mise à jour delta 1x/jour et vérification complète si la dernière réussie date de plus de 7 jours.",
        "opt_dry_run": "Mode Simulation (Dry-Run : aucune écriture)",
        "tip_opt_dry_run": "Simule toute l'opération et affiche le journal détaillé sans modifier aucun contact dans Outlook ni dans Google.",
        "sec_win": "Intégration Windows",
        "sec_excl": "Contacts exclusifs",
        "btn_test_g": "Tester la connexion",
        "lbl_g_ok": "✔ Connecté",
        "lbl_excl_ol": "Contacts exclusifs Outlook : {n}",
        "lbl_excl_g": "Contacts exclusifs Google : {n}",
        "btn_clean": "Nettoyer",
        "sub_ready": "Sous-étape : prêt",
        "win_min_tray": "Réduire dans le System Tray",
        "tip_win_min_tray": "Envoie la fenêtre dans la zone de notification en bas à droite lorsqu'on clique sur Réduire (_).",
        "win_startup": "Démarrer avec Windows",
        "tip_win_startup": "Lance automatiquement l'application au démarrage de votre session Windows (repliée dans le System Tray).",
        "lbl_max_del": "Alerte si suppressions >",
        "tip_max_del": "Met en pause et demande une validation globale si le nombre de suppressions planifiées dépasse ce seuil.",
        "btn_start": "▶  Lancer maintenant",
        "sec_log": "Sync Details & Log",
        "btn_clear_log": "Effacer l'écran",
        "status_ready": "Prêt.",
        "conflict_title": "Merge Prompt - Conflit détecté",
        "conflict_header": "Le contact suivant a été modifié des deux côtés :",
        "btn_skip": "Ignorer (Skip)",
        "confirm_del_title": "Prompt Deletion",
        "confirm_del_msg": "Le contact \"{disp}\" a été supprimé de l'autre côté.\nVoulez-vous le supprimer de {side} ?",
        "mass_del_title": "⚠️ Alerte : Seuil de suppressions dépassé",
        "mass_del_msg": "{count} suppressions anticipées détectées (seuil autorisé : {limit}).\n\nExemples :\n{preview}\n\nVoulez-vous approuver et exécuter ces {count} suppressions ?",
        "err_zero_contacts": "Sécurité déclenchée : la source {side} a renvoyé 0 contact alors que {linked} contacts sont liés. Opération interrompue."
    },
    "EN": {
        "app_title": "SyncContactsOG",
        "theme_light": "Light",
        "theme_dark": "Dark",
        "sec_accounts": "Accounts & Profiles",
        "lbl_ol_email": "Outlook Account (Store):",
        "lbl_g_email": "Google Account:",
        "lbl_auto_sync": "Auto-Sync every (hours, 0=off):",
        "tip_auto_sync": "Automatically triggers synchronization at the specified interval while running (0 to disable).",
        "sec_mode": "Synchronization Mode",
        "mode_prompt": "Merge Prompt",
        "mode_prompt_sub": "Ask on conflict",
        "tip_mode_prompt": "Two-way sync: prompts you with a side-by-side comparison whenever a contact was modified on both sides.",
        "mode_ol_wins": "Merge Outlook Wins",
        "mode_ol_wins_sub": "Two-way - Outlook priority",
        "tip_mode_ol_wins": "Two-way sync: syncs both ways, automatically resolving conflicts in favor of Outlook.",
        "mode_g_wins": "Merge Google Wins",
        "mode_g_wins_sub": "Two-way - Google priority",
        "tip_mode_g_wins": "Two-way sync: syncs both ways, automatically resolving conflicts in favor of Google.",
        "mode_ol_to_g": "Outlook To Google Only",
        "mode_ol_to_g_sub": "One-way",
        "tip_mode_ol_to_g": "Strict one-way sync: Outlook is the master source. Google Contacts never overwrites Outlook.",
        "mode_g_to_ol": "Google To Outlook Only",
        "mode_g_to_ol_sub": "One-way",
        "tip_mode_g_to_ol": "Strict one-way sync: Google is the master source. Outlook never overwrites Google.",
        "sec_opts": "Options & Data Integrity",
        "opt_photos": "Sync Contact Photos",
        "tip_opt_photos": "Syncs high-resolution contact photos between Outlook and Google Contacts.",
        "opt_file_as": "Use Outlook FileAs for Google Title",
        "tip_opt_file_as": "Uses Outlook's 'File As' field as the primary display name in Google Contacts.",
        "opt_force_fa_fmt": "Force FileAs format to:",
        "tip_opt_force_fa_fmt": "Automatically formats the FileAs field on source contacts using the selected pattern (Name SURNAME or SURNAME, Name).",
        "opt_force_rtf": "Force RTF Contacts (Overwrite Outlook RTF notes)",
        "tip_opt_force_rtf": "Allows Google plain-text notes to overwrite Outlook rich-text (RTF) formatted notes.",
        "opt_del": "Sync Deletion (Propagate deletions)",
        "tip_opt_del": "Propagates deletions: if a previously synced contact is deleted on one side, it will be deleted on the other.",
        "opt_prompt_del": "Prompt Deletion (Confirm before deleting)",
        "tip_opt_prompt_del": "Asks for confirmation before deleting any single contact on either side.",
        "opt_force_full": "Force Full Compare (vs. incremental by default and full >7d)",
        "tip_opt_force_full": "Si coché : Full Sync Outlook + Google + Archive (.db et .xlsx). Sinon : synchro incrémentale rapide par défaut et bascule automatique en Full Sync si la dernière Full Sync date de > 7 jours.",
        "opt_extract_db": "Keep a Local Archive (.db with photos + .xlsx)",
        "tip_opt_extract_db": "Maintains contacts.db (with photos) and contacts.xlsx: delta update once/day and full refresh if last complete run is older than 7 days.",
        "opt_dry_run": "Simulation Mode (Dry-Run: no writes)",
        "tip_opt_dry_run": "Simulates the entire run and logs every action without modifying any contact in Outlook or Google.",
        "sec_win": "Windows Integration",
        "sec_excl": "Exclusive contacts",
        "btn_test_g": "Test Connection",
        "lbl_g_ok": "✔ Connected",
        "lbl_excl_ol": "Outlook exclusive contacts: {n}",
        "lbl_excl_g": "Google exclusive contacts: {n}",
        "btn_clean": "Clean",
        "sub_ready": "Sub-step: ready",
        "win_min_tray": "Minimize to System Tray",
        "tip_win_min_tray": "Sends the window to the bottom-right notification area when clicking Minimize (_).",
        "win_startup": "Start with Windows",
        "tip_win_startup": "Automatically starts the application at Windows login (minimized in the System Tray).",
        "lbl_max_del": "Alert if deletions >",
        "tip_max_del": "Pauses and asks for global approval if planned deletions exceed this threshold.",
        "btn_start": "▶  Run Now",
        "sec_log": "Sync Details & Log",
        "btn_clear_log": "Clear Log",
        "status_ready": "Ready.",
        "conflict_title": "Merge Prompt - Conflict Detected",
        "conflict_header": "This contact was modified on both sides:",
        "btn_skip": "Skip",
        "confirm_del_title": "Prompt Deletion",
        "confirm_del_msg": "Contact \"{disp}\" was deleted on the other side.\nDo you want to delete it from {side}?",
        "mass_del_title": "⚠️ Alert: Deletion Threshold Exceeded",
        "mass_del_msg": "{count} planned deletions detected (configured limit: {limit}).\n\nExamples:\n{preview}\n\nDo you approve executing all {count} deletions?",
        "err_zero_contacts": "Safety lock triggered: {side} returned 0 contacts while {linked} contacts are linked. Run aborted."
    }
}


# ==============================================================================
# GESTION DU DEMARRAGE WINDOWS (REGISTRE HKCU - SANS DROITS ADMIN)
# ==============================================================================
def set_windows_startup(enabled: bool):
    key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
    if getattr(sys, "frozen", False):
        cmd = f'"{sys.executable}" --minimized'
    else:
        pythonw_exe = os.path.join(os.path.dirname(sys.executable), "pythonw.exe")
        if not os.path.exists(pythonw_exe):
            pythonw_exe = sys.executable
        cmd = f'"{pythonw_exe}" "{os.path.abspath(__file__)}" --minimized'

    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_SET_VALUE) as key:
            if enabled:
                winreg.SetValueEx(key, APP_REG_NAME, 0, winreg.REG_SZ, cmd)
            else:
                try:
                    winreg.DeleteValue(key, APP_REG_NAME)
                except FileNotFoundError:
                    pass
    except Exception:
        pass


# ==============================================================================
# WIDGET COMPACT A FLECHES VERTICALES
# ==============================================================================
class CompactVerticalSpinbox(ctk.CTkFrame):



    def __init__(self, master, variable, min_val=0, max_val=500, command=None, **kwargs):
        super().__init__(master, width=106, height=32, fg_color="transparent")
        self.grid_propagate(False)
        self.grid_columnconfigure(0, minsize=28)
        self.grid_columnconfigure(1, minsize=46)
        self.grid_columnconfigure(2, minsize=28)
        self.variable = variable
        self.min_val = min_val
        self.max_val = max_val
        self.command = command
        btn_down = ctk.CTkButton(
            self, text="-", width=26, height=26,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color=("#d5d9df", "#3a3f47"), hover_color=("#c3c9d1", "#4b525c"),
            text_color=("#1f2328", "#f0f3f6"), command=self._decrement
        )
        btn_down.grid(row=0, column=0, padx=(0, 2), pady=2)
        self.entry = ctk.CTkEntry(self, textvariable=self.variable, width=42, height=26, justify="center")
        self.entry.grid(row=0, column=1, padx=2, pady=2)
        btn_up = ctk.CTkButton(
            self, text="+", width=26, height=26,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color=("#d5d9df", "#3a3f47"), hover_color=("#c3c9d1", "#4b525c"),
            text_color=("#1f2328", "#f0f3f6"), command=self._increment
        )
        btn_up.grid(row=0, column=2, padx=(2, 0), pady=2)

    def _get_int(self):
        try:
            return int(self.variable.get())
        except ValueError:
            return self.min_val

    def _increment(self):
        val = min(self.max_val, self._get_int() + 1)
        self.variable.set(str(val))
        if self.command:
            self.command()

    def _decrement(self):
        val = max(self.min_val, self._get_int() - 1)
        self.variable.set(str(val))
        if self.command:
            self.command()

    def _validate(self, _event=None):
        val = max(self.min_val, min(self.max_val, self._get_int()))
        self.variable.set(str(val))
        if self.command:
            self.command()


# ==============================================================================
# BULLE D'AIDE AU SURVOL (TOOLTIP)
# ==============================================================================
class HoverToolTip:
    def __init__(self, widget, text_getter):
        self.widget = widget
        self.text_getter = text_getter
        self.tip_window = None
        self.after_id = None
        widget.bind("<Enter>", self._schedule)
        widget.bind("<Leave>", self._hide)
        widget.bind("<ButtonPress>", self._hide)

    def _schedule(self, _event=None):
        self._cancel()
        self.after_id = self.widget.after(350, self._show)

    def _cancel(self):
        if self.after_id:
            self.widget.after_cancel(self.after_id)
            self.after_id = None

    def _show(self):
        text = self.text_getter()
        if not text or self.tip_window:
            return
        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 6

        self.tip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_attributes("-topmost", True)
        tw.wm_geometry(f"+{x}+{y}")

        frame = tk.Frame(tw, background="#2b313d", borderwidth=1, relief="solid")
        frame.pack()
        lbl = tk.Label(
            frame, text=text, justify="left", wraplength=380,
            background="#2b313d", foreground="#f1f5f9",
            font=("Segoe UI", 9), padx=10, pady=6
        )
        lbl.pack()

    def _hide(self, _event=None):
        self._cancel()
        if self.tip_window:
            self.tip_window.destroy()
            self.tip_window = None


# ==============================================================================
# UTILITAIRES SQLITE, NORMALISATION & SONDAGE OUTLOOK
# ==============================================================================
def init_db():
    conn = sqlite3.connect(DB_FILE)
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS contact_map (
            outlook_entry_id TEXT PRIMARY KEY,
            google_resource_name TEXT UNIQUE NOT NULL,
            outlook_hash TEXT NOT NULL,
            google_hash TEXT NOT NULL,
            outlook_photo_hash TEXT,
            google_photo_hash TEXT,
            last_synced TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS deleted_contacts_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            deleted_at TEXT NOT NULL,
            target_system TEXT NOT NULL,
            display_name TEXT NOT NULL,
            outlook_entry_id TEXT,
            google_resource_name TEXT,
            contact_data_json TEXT NOT NULL,
            photo_blob BLOB
        );
    """)
    conn.commit()
    return conn


def normalize_phone(phone_str):
    if not phone_str:
        return ""
    digits = re.sub(r"\D", "", phone_str)
    return digits[-9:] if len(digits) >= 9 else digits


def compute_hash(canonical_dict):
    serialized = json.dumps(canonical_dict, sort_keys=True, ensure_ascii=False)
    return hashlib.sha256(serialized.encode("utf-8")).hexdigest()


def is_outlook_body_rich_rtf(item):
    try:
        rtf_bytes = item.RTFBody
        if not rtf_bytes:
            return False
        rtf_text = bytes(rtf_bytes).decode("latin1", errors="ignore")
        rich_markers = (r"\b ", r"\i ", r"\ul", r"\highlight", r"\cf1", r"\cf2", r"\trowd", r"\bullet")
        return any(marker in rtf_text for marker in rich_markers)
    except Exception:
        return False


def count_outlook_com_contacts(outlook_email: str):
    if not outlook_email:
        return None
    pythoncom.CoInitialize()
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        for store in namespace.Stores:
            if outlook_email.lower() in store.DisplayName.lower():
                folder = store.GetDefaultFolder(10)
                count = sum(1 for item in folder.Items if item.Class == 40)
                return count
        return None
    except Exception:
        return None
    finally:
        pythoncom.CoUninitialize()


def count_outlook_graph_contacts():
    raw_graph = load_secret_from_vault(GRAPH_CRED_TARGET_NAME)
    if not raw_graph:
        return None
    try:
        token = json.loads(raw_graph).get("access_token")
        if not token:
            return None
        req = urllib.request.Request(
            "https://graph.microsoft.com/v1.0/me/contacts/$count",
            headers={"Authorization": f"Bearer {token}", "ConsistencyLevel": "eventual"}
        )
        with urllib.request.urlopen(req, timeout=8) as resp:
            return int(resp.read().decode("utf-8"))
    except Exception:
        return None


def execute_google_call_with_backoff(api_callable, max_retries=6):
    delay = 3.0
    for attempt in range(max_retries):
        try:
            return api_callable.execute()
        except (ConnectionResetError, ConnectionAbortedError, OSError) as net_err:
            if attempt < 4:
                time.sleep(1.5 * (2 ** attempt))
                continue
            raise
        except HttpError as err:
            if err.resp.status == 429 and attempt < max_retries - 1:
                time.sleep(62.0)
            elif err.resp.status in (500, 502, 503) and attempt < max_retries - 1:
                time.sleep(delay)
                delay *= 2
            else:
                raise
        except Exception:
            if attempt < max_retries - 1:
                time.sleep(delay)
                delay *= 2
            else:
                raise


# ==============================================================================
# CONVERSION DES CONTACTS (FORMAT PIVOT)
# ==============================================================================
def format_target_file_as(first_name, last_name, style):
    fn = (first_name or "").strip()
    ln = (last_name or "").strip()
    if not fn and not ln:
        return ""
    if not fn:
        return ln.upper()
    if not ln:
        return fn
    if style == "SURNAME, Name":
        return f"{ln.upper()}, {fn}"
    return f"{fn} {ln.upper()}"

def outlook_to_canonical(item, use_file_as=True):
    emails = []
    for attr, label in [("Email1Address", "work"), ("Email2Address", "home"), ("Email3Address", "other")]:
        val = getattr(item, attr, "")
        if val and "@" in val:
            emails.append({"value": val.strip(), "type": label})

    phones = []
    phone_map = [
        ("MobileTelephoneNumber", "mobile"),
        ("BusinessTelephoneNumber", "work"),
        ("Business2TelephoneNumber", "work"),
        ("HomeTelephoneNumber", "home"),
        ("Home2TelephoneNumber", "home"),
        ("BusinessFaxNumber", "workFax"),
        ("HomeFaxNumber", "homeFax"),
        ("OtherTelephoneNumber", "other"),
        ("AssistantTelephoneNumber", "assistant")
    ]
    for attr, p_type in phone_map:
        val = getattr(item, attr, "")
        if val and val.strip():
            phones.append({"value": val.strip(), "type": p_type})

    addresses = []
    for prefix, a_type in [("Business", "work"), ("Home", "home"), ("Other", "other")]:
        street = getattr(item, f"{prefix}AddressStreet", "").strip()
        city = getattr(item, f"{prefix}AddressCity", "").strip()
        state = getattr(item, f"{prefix}AddressState", "").strip()
        postal = getattr(item, f"{prefix}AddressPostalCode", "").strip()
        country = getattr(item, f"{prefix}AddressCountry", "").strip()
        if any([street, city, state, postal, country]):
            addresses.append({
                "type": a_type, "street": street, "city": city,
                "state": state, "postal": postal, "country": country
            })

    bday_str = ""
    try:
        bday = item.Birthday
        if bday and bday.year != 4501:
            bday_str = f"{bday.year:04d}-{bday.month:02d}-{bday.day:02d}"
    except Exception:
        pass

    custom = {}
    for idx in range(1, 5):
        u_val = getattr(item, f"User{idx}", "").strip()
        if u_val:
            custom[f"Outlook_User{idx}"] = u_val

    categories = sorted([c.strip() for c in (item.Categories or "").split(",") if c.strip()])
    file_as = (item.FileAs or "").strip() if use_file_as else ""

    return {
        "file_as": file_as,
        "first_name": (item.FirstName or "").strip(),
        "middle_name": (item.MiddleName or "").strip(),
        "last_name": (item.LastName or "").strip(),
        "prefix": (item.Title or "").strip(),
        "suffix": (item.Suffix or "").strip(),
        "company": (item.CompanyName or "").strip(),
        "job_title": (item.JobTitle or "").strip(),
        "department": (item.Department or "").strip(),
        "emails": emails,
        "phones": phones,
        "addresses": addresses,
        "birthday": bday_str,
        "website": (item.WebPage or "").strip(),
        "notes": (item.Body or "").replace("\r\n", "\n").strip(),
        "custom": custom,
        "categories": categories
    }


def google_to_canonical(person, id_to_label_name, use_file_as=True):
    names = person.get("names", [{}])[0]
    orgs = person.get("organizations", [{}])[0]

    emails = [
        {"value": e.get("value", "").strip(), "type": e.get("type", "other")}
        for e in person.get("emailAddresses", []) if e.get("value")
    ]
    phones = [
        {"value": p.get("value", "").strip(), "type": p.get("type", "other")}
        for p in person.get("phoneNumbers", []) if p.get("value")
    ]

    addresses = []
    for a in person.get("addresses", []):
        street = a.get("streetAddress", "").strip()
        city = a.get("city", "").strip()
        state = a.get("region", "").strip()
        postal = a.get("postalCode", "").strip()
        country = a.get("country", "").strip()
        if any([street, city, state, postal, country]):
            addresses.append({
                "type": a.get("type", "other"), "street": street, "city": city,
                "state": state, "postal": postal, "country": country
            })

    bday_str = ""
    bdays = person.get("birthdays", [])
    if bdays and "date" in bdays[0]:
        d = bdays[0]["date"]
        if d.get("year") and d.get("month") and d.get("day"):
            bday_str = f"{d['year']:04d}-{d['month']:02d}-{d['day']:02d}"

    urls = person.get("urls", [])
    website = urls[0].get("value", "").strip() if urls else ""

    bios = person.get("biographies", [])
    notes = bios[0].get("value", "").replace("\r\n", "\n").strip() if bios else ""

    custom = {
        u["key"]: u["value"].strip()
        for u in person.get("userDefined", []) if u.get("key") and u.get("value")
    }

    categories = []
    for m in person.get("memberships", []):
        grp_id = m.get("contactGroupMembership", {}).get("contactGroupResourceName")
        if grp_id in id_to_label_name:
            categories.append(id_to_label_name[grp_id])

    g_file_ases = person.get("fileAses", [])
    raw_g_file_as = (g_file_ases[0].get("value", "") if g_file_ases else "") or names.get("unstructuredName", "")
    file_as = raw_g_file_as.strip() if use_file_as else ""

    return {
        "file_as": file_as,
        "first_name": names.get("givenName", "").strip(),
        "middle_name": names.get("middleName", "").strip(),
        "last_name": names.get("familyName", "").strip(),
        "prefix": names.get("honorificPrefix", "").strip(),
        "suffix": names.get("honorificSuffix", "").strip(),
        "company": orgs.get("name", "").strip(),
        "job_title": orgs.get("title", "").strip(),
        "department": orgs.get("department", "").strip(),
        "emails": emails,
        "phones": phones,
        "addresses": addresses,
        "birthday": bday_str,
        "website": website,
        "notes": notes,
        "custom": custom,
        "categories": sorted(categories)
    }


def canonical_to_google_body(canon, label_name_to_id, use_file_as=True):
    body = {}
    name_obj = {}
    if canon["first_name"]: name_obj["givenName"] = canon["first_name"]
    if canon["middle_name"]: name_obj["middleName"] = canon["middle_name"]
    if canon["last_name"]: name_obj["familyName"] = canon["last_name"]
    if canon["prefix"]: name_obj["honorificPrefix"] = canon["prefix"]
    if canon["suffix"]: name_obj["honorificSuffix"] = canon["suffix"]
    if use_file_as and canon["file_as"]:
        name_obj["unstructuredName"] = canon["file_as"]
    if name_obj:
        body["names"] = [name_obj]
    if use_file_as:
        body["fileAses"] = [{"value": canon["file_as"]}] if canon.get("file_as") else []

    if canon["emails"]:
        body["emailAddresses"] = canon["emails"]
    if canon["phones"]:
        body["phoneNumbers"] = canon["phones"]

    org_obj = {}
    if canon["company"]: org_obj["name"] = canon["company"]
    if canon["job_title"]: org_obj["title"] = canon["job_title"]
    if canon["department"]: org_obj["department"] = canon["department"]
    if org_obj:
        body["organizations"] = [org_obj]

    if canon["addresses"]:
        body["addresses"] = [
            {
                "type": a["type"], "streetAddress": a["street"], "city": a["city"],
                "region": a["state"], "postalCode": a["postal"], "country": a["country"]
            } for a in canon["addresses"]
        ]

    if canon["birthday"]:
        y, m, d = map(int, canon["birthday"].split("-"))
        body["birthdays"] = [{"date": {"year": y, "month": m, "day": d}}]

    if canon["website"]:
        body["urls"] = [{"value": canon["website"], "type": "work"}]
    if canon["notes"]:
        body["biographies"] = [{"value": canon["notes"], "contentType": "TEXT_PLAIN"}]
    if canon["custom"]:
        body["userDefined"] = [{"key": k, "value": v} for k, v in canon["custom"].items()]

    memberships = [{"contactGroupMembership": {"contactGroupResourceName": "contactGroups/myContacts"}}]
    for cat in canon["categories"]:
        if cat in label_name_to_id:
            memberships.append({"contactGroupMembership": {"contactGroupResourceName": label_name_to_id[cat]}})
    body["memberships"] = memberships

    return body


def apply_canonical_to_outlook(item, canon, update_notes=True, dry_run=False):
    item.FirstName = canon["first_name"]
    item.MiddleName = canon["middle_name"]
    item.LastName = canon["last_name"]
    item.Title = canon["prefix"]
    item.Suffix = canon["suffix"]
    if canon["file_as"]:
        item.FileAs = canon["file_as"]
    item.CompanyName = canon["company"]
    item.JobTitle = canon["job_title"]
    item.Department = canon["department"]

    emails = canon["emails"]
    item.Email1Address = emails[0]["value"] if len(emails) > 0 else ""
    item.Email2Address = emails[1]["value"] if len(emails) > 1 else ""
    item.Email3Address = emails[2]["value"] if len(emails) > 2 else ""

    item.MobileTelephoneNumber = ""
    item.BusinessTelephoneNumber = ""
    item.Business2TelephoneNumber = ""
    item.HomeTelephoneNumber = ""
    item.Home2TelephoneNumber = ""
    item.OtherTelephoneNumber = ""

    for p in canon["phones"]:
        ptype = p["type"].lower()
        val = p["value"]
        if "mobile" in ptype and not item.MobileTelephoneNumber:
            item.MobileTelephoneNumber = val
        elif "work" in ptype:
            if not item.BusinessTelephoneNumber:
                item.BusinessTelephoneNumber = val
            elif not item.Business2TelephoneNumber:
                item.Business2TelephoneNumber = val
        elif "home" in ptype:
            if not item.HomeTelephoneNumber:
                item.HomeTelephoneNumber = val
            elif not item.Home2TelephoneNumber:
                item.Home2TelephoneNumber = val
        elif not item.OtherTelephoneNumber:
            item.OtherTelephoneNumber = val

    for prefix in ["Business", "Home", "Other"]:
        setattr(item, f"{prefix}AddressStreet", "")
        setattr(item, f"{prefix}AddressCity", "")
        setattr(item, f"{prefix}AddressState", "")
        setattr(item, f"{prefix}AddressPostalCode", "")
        setattr(item, f"{prefix}AddressCountry", "")

    for a in canon["addresses"]:
        atype = a["type"].lower()
        prefix = "Business" if "work" in atype else ("Home" if "home" in atype else "Other")
        setattr(item, f"{prefix}AddressStreet", a["street"])
        setattr(item, f"{prefix}AddressCity", a["city"])
        setattr(item, f"{prefix}AddressState", a["state"])
        setattr(item, f"{prefix}AddressPostalCode", a["postal"])
        setattr(item, f"{prefix}AddressCountry", a["country"])

    if canon["birthday"]:
        item.Birthday = datetime.strptime(canon["birthday"], "%Y-%m-%d")

    item.WebPage = canon["website"]
    if update_notes:
        item.Body = canon["notes"]
    item.Categories = ", ".join(canon["categories"])

    for idx in range(1, 5):
        key = f"Outlook_User{idx}"
        setattr(item, f"User{idx}", canon["custom"].get(key, ""))

    if not dry_run:
        item.Save()


# ==============================================================================
# GESTION DES PHOTOS
# ==============================================================================
def get_outlook_photo_bytes(contact_item):
    if not contact_item.HasPicture:
        return None
    for att in contact_item.Attachments:
        if att.FileName.lower().startswith("contactpicture"):
            tmp = os.path.join(tempfile.gettempdir(), f"ol_pic_{contact_item.EntryID[-10:]}.jpg")
            att.SaveAsFile(tmp)
            with open(tmp, "rb") as f:
                data = f.read()
            os.remove(tmp)
            return data
    return None


def set_outlook_photo_bytes(contact_item, photo_bytes, dry_run=False):
    if dry_run:
        return
    if not photo_bytes:
        if contact_item.HasPicture:
            contact_item.RemovePicture()
            contact_item.Save()
        return
    tmp = os.path.join(tempfile.gettempdir(), f"g_to_ol_{contact_item.EntryID[-10:]}.jpg")
    with open(tmp, "wb") as f:
        f.write(photo_bytes)
    contact_item.AddPicture(tmp)
    contact_item.Save()
    os.remove(tmp)


def get_google_photo_bytes(person, service):
    for p in person.get("photos", []):
        if not p.get("default", False) and p.get("url"):
            resp, content = service._http.request(p["url"])
            if resp.status == 200:
                return content
    return None


def set_google_photo_bytes(service, resource_name, photo_bytes, dry_run=False):
    if dry_run:
        return
    if photo_bytes:
        b64 = base64.b64encode(photo_bytes).decode("utf-8")
        execute_google_call_with_backoff(
            service.people().updateContactPhoto(resourceName=resource_name, body={"photoBytes": b64})
        )
        time.sleep(0.75)
    else:
        execute_google_call_with_backoff(
            service.people().deleteContactPhoto(resourceName=resource_name)
        )


# ==============================================================================
# APPLICATION PRINCIPALE (ZERO JSON EXTERNE : VAULT + REGISTRE + EXE AUTONOME)
# ==============================================================================
def get_outlook_available_accounts():
    items = []
    try:
        import pythoncom, win32com.client
        pythoncom.CoInitialize()
        ol = win32com.client.Dispatch("Outlook.Application")
        ns = ol.GetNamespace("MAPI")
        for acc in ns.Accounts:
            try:
                smtp = str(getattr(acc, "SmtpAddress", "") or getattr(acc, "DisplayName", "")).strip()
                if smtp and smtp not in items:
                    items.append(smtp)
            except Exception:
                pass
        for st in ns.Stores:
            try:
                fld = st.GetDefaultFolder(10)
                name = str(st.DisplayName).strip()
                if fld is not None and name and name not in items:
                    items.append(name)
            except Exception:
                pass
    except Exception:
        pass
    return items

CRED_CLIENT_VAULT_NAME = "SyncContactsOG_Google_Client_JSON"

def load_google_client_config_dpapi():
    """Priorite 1: Coffre-fort DPAPI natif. Priorite 2: cle par defaut embarquee dans l'exe (_MEIPASS) ou locale."""
    try:
        raw = load_secret_from_vault(CRED_CLIENT_VAULT_NAME)
        if raw:
            data = json.loads(raw)
            if isinstance(data, dict) and ("installed" in data or "web" in data):
                return data
    except Exception:
        pass

    candidates = []
    if hasattr(sys, "_MEIPASS"):
        candidates.append(os.path.join(sys._MEIPASS, "credentials.json"))
    candidates.append(os.path.join(BASE_DIR, "credentials.json"))

    for cpath in candidates:
        if os.path.exists(cpath):
            try:
                with open(cpath, "r", encoding="utf-8") as f:
                    data = json.load(f)
                if isinstance(data, dict) and ("installed" in data or "web" in data):
                    return data
            except Exception:
                pass
    return None

def save_google_client_config_dpapi(client_dict):
    save_secret_to_vault(CRED_CLIENT_VAULT_NAME, json.dumps(client_dict, ensure_ascii=False))

def clear_google_client_config_dpapi():
    try:
        delete_secret_from_vault(CRED_CLIENT_VAULT_NAME)
    except Exception:
        pass

class CleanRadioButton(ctk.CTkFrame):
    """Bouton radio épuré : anneau fin clair avec point central blanc (identique en mode sombre/clair)."""
    def __init__(self, master, text="", variable=None, value="", command=None, font=None, size=16, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self.variable = variable
        self.value = value
        self.command = command
        self._size = size
        self._hover = False

        bg_hex = self._resolve_bg()
        self.canvas = tk.Canvas(
            self, width=size + 2, height=size + 2,
            bg=bg_hex, highlightthickness=0, bd=0, cursor="hand2"
        )
        self.canvas.pack(side="left", padx=(0, 7))

        self.lbl = ctk.CTkLabel(
            self, text=text, font=font or ctk.CTkFont(size=12, weight="bold"),
            text_color=("#111827", "#F3F4F6"), cursor="hand2", anchor="w"
        )
        self.lbl.pack(side="left")

        for w in (self, self.canvas, self.lbl):
            w.bind("<Button-1>", self._on_click)
            w.bind("<Enter>", self._on_enter)
            w.bind("<Leave>", self._on_leave)

        if self.variable is not None:
            self.variable.trace_add("write", lambda *_: self._redraw())
        self.after(30, self._sync_bg_and_redraw)

    def _resolve_bg(self) -> str:
        mode = ctk.get_appearance_mode()
        return "#232730" if mode == "Dark" else "#F3F4F6"

    def _sync_bg_and_redraw(self):
        try:
            self.canvas.configure(bg=self._resolve_bg())
            self._redraw()
        except Exception:
            pass

    def _on_click(self, _event=None):
        if self.variable is not None:
            self.variable.set(self.value)
        if callable(self.command):
            self.command()

    def _on_enter(self, _event=None):
        self._hover = True
        self._redraw()

    def _on_leave(self, _event=None):
        self._hover = False
        self._redraw()

    def _redraw(self):
        try:
            self.canvas.configure(bg=self._resolve_bg())
            self.canvas.delete("all")
            s = self._size
            pad = 2
            x0, y0, x1, y1 = pad, pad, s, s
            selected = (self.variable is not None and str(self.variable.get()) == str(self.value))
            is_dark = (ctk.get_appearance_mode() == "Dark")

            if selected:
                ring_col = "#FFFFFF" if is_dark else "#111827"
                dot_col = "#FFFFFF" if is_dark else "#111827"
                self.canvas.create_oval(x0, y0, x1, y1, outline=ring_col, width=1.6)
                inset = 4.2
                self.canvas.create_oval(x0 + inset, y0 + inset, x1 - inset, y1 - inset, fill=dot_col, outline=dot_col)
            else:
                ring_col = ("#E5E7EB" if self._hover else "#9CA3AF") if is_dark else ("#374151" if self._hover else "#6B7280")
                self.canvas.create_oval(x0, y0, x1, y1, outline=ring_col, width=1.5)
        except Exception:
            pass

    def configure(self, **kwargs):
        if "text" in kwargs:
            self.lbl.configure(text=kwargs.pop("text"))
        if "font" in kwargs:
            self.lbl.configure(font=kwargs.pop("font"))
        if kwargs:
            super().configure(**kwargs)
        self._redraw()


def resource_path(relative_path):
    base_path = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)

def get_portable_data_dir():
    base_dir = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else os.path.dirname(os.path.abspath(__file__))
    data_dir = os.path.join(base_dir, "data")
    os.makedirs(data_dir, exist_ok=True)
    return data_dir

def get_portable_credentials_path():
    return os.path.join(get_portable_data_dir(), "credentials.json")

def validate_google_credentials_json(file_path):
    if not file_path or not os.path.isfile(file_path):
        return False, "Fichier introuvable."
    try:
        with open(file_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        if not isinstance(data, dict):
            return False, "Format JSON invalide."
        block = data.get("installed") or data.get("web")
        if not isinstance(block, dict) or not block.get("client_id") or not block.get("client_secret"):
            return False, "Clés OAuth 'installed'/'web' (client_id, client_secret) manquantes."
        return True, "OK"
    except Exception as e:
        return False, f"Erreur de lecture JSON : {e}"

def resolve_credentials_path(ignore_embedded=False):
    local_json = get_portable_credentials_path()
    ok, _ = validate_google_credentials_json(local_json)
    if ok:
        return local_json, "data"
    return None, "missing"

def load_portable_client_config():
    c_path, c_src = resolve_credentials_path()
    if c_src == "data" and c_path and os.path.isfile(c_path):
        with open(c_path, "r", encoding="utf-8") as f:
            return json.load(f)
    return None

_DPAPI_ENTROPY = b"SyncContactsOG::OAuth2::DPAPI::v1"
_REG_TOKEN_PATH = r"Software\SyncContactsOG"
_REG_TOKEN_NAME = "GoogleOAuthTokenDPAPI"

def _save_dpapi_registry_token(token_json_str):
    import winreg
    import win32crypt
    blob = win32crypt.CryptProtectData(
        token_json_str.encode("utf-8"),
        "SyncContactsOG Google OAuth2 Token",
        _DPAPI_ENTROPY,
        None,
        None,
        0
    )
    with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, _REG_TOKEN_PATH, 0, winreg.KEY_WRITE) as key:
        winreg.SetValueEx(key, _REG_TOKEN_NAME, 0, winreg.REG_BINARY, blob)

def _load_dpapi_registry_token():
    import winreg
    import win32crypt
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, _REG_TOKEN_PATH, 0, winreg.KEY_READ) as key:
            blob, reg_type = winreg.QueryValueEx(key, _REG_TOKEN_NAME)
            if reg_type == winreg.REG_BINARY and blob:
                _, raw_bytes = win32crypt.CryptUnprotectData(
                    blob,
                    _DPAPI_ENTROPY,
                    None,
                    None,
                    0
                )
                return json.loads(raw_bytes.decode("utf-8"))
    except FileNotFoundError:
        return None
    except Exception:
        return None
    return None

def run_google_oauth_connect(creds_json_path, log_callback=None):
    from google.oauth2.credentials import Credentials
    from google.auth.transport.requests import Request
    from google_auth_oauthlib.flow import InstalledAppFlow
    from googleapiclient.discovery import build

    scopes = ["https://www.googleapis.com/auth/contacts"]
    base_dir = os.path.dirname(creds_json_path)
    legacy_token_file = os.path.join(base_dir, "google_token.json")

    creds = None
    token_info = _load_dpapi_registry_token()
    if token_info:
        try:
            creds = Credentials.from_authorized_user_info(token_info, scopes)
        except Exception:
            creds = None

    # Migration transparente d'un éventuel ancien google_token.json vers HKCU DPAPI puis suppression sur disque
    if creds is None and os.path.isfile(legacy_token_file):
        try:
            creds = Credentials.from_authorized_user_file(legacy_token_file, scopes)
            _save_dpapi_registry_token(creds.to_json())
        except Exception:
            creds = None
    if os.path.isfile(legacy_token_file):
        try:
            os.remove(legacy_token_file)
        except Exception:
            pass

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            if log_callback:
                log_callback("Rafraîchissement du jeton OAuth Google (HKCU DPAPI)...", "Information")
            creds.refresh(Request())
        else:
            if log_callback:
                log_callback("Ouverture du navigateur pour l'authentification OAuth Google...", "Information")
            flow = InstalledAppFlow.from_client_secrets_file(creds_json_path, scopes)
            creds = flow.run_local_server(port=0, open_browser=True)
        _save_dpapi_registry_token(creds.to_json())

    service = build("people", "v1", credentials=creds, cache_discovery=False)
    res = service.people().connections().list(resourceName="people/me", pageSize=1, personFields="names").execute()
    total = res.get("totalPeople") or res.get("totalItems") or len(res.get("connections", []))
    return service, creds, f"{total} contact(s)"

class SyncContactsOGApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.config = self.load_config()
        self.lang = self.config.get("language", "FR")
        self.theme = self.config.get("theme", "Dark")

        ctk.set_appearance_mode(self.theme)
        ctk.set_default_color_theme("blue")

        self.title(f"SyncContactsOG — v{__version__}")
        try:
            _ico = resource_path("SyncContactsOG.ico")
            if os.path.exists(_ico):
                self.iconbitmap(_ico)
                self.after(200, lambda: self.iconbitmap(_ico))
        except Exception:
            pass
        if os.path.exists(ICON_FILE):
            self.iconbitmap(os.path.join(getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__))), "SyncContactsOG.ico"))

        # Taille fixe verrouillée pour garantir un alignement strictement identique en FR et EN
        self.geometry("1340x771")
        self.minsize(1340, 735)
        self.maxsize(1340, 735)
        self.resizable(False, False)

        self.is_syncing = False
        self.auto_sync_job = None
        self.tray_icon = None

        self.build_ui()
        self.apply_translations()
        self._load_persisted_exclusives()
        self.setup_system_tray()
        self.refresh_exclusives_card()
        self.after(600, lambda: self.update_tray_icon_badge(
            getattr(self, "_last_excl_ol_count", 0) + getattr(self, "_last_excl_g_count", 0)
        ))
        self.bind("<Unmap>", self.on_window_unmap)
        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.schedule_auto_sync()

        # Comptage silencieux des connecteurs Outlook en arrière-plan
        threading.Thread(target=self.refresh_outlook_counts_async, daemon=True).start()

        # Ne replier dans le System Tray au lancement que si démarré avec --minimized
        if "--minimized" in sys.argv:
            self.after(120, self.hide_to_tray)

    def t(self, key):
        return I18N.get(self.lang, I18N["FR"]).get(key, key)

    def load_config(self):
        cfg = DEFAULT_CONFIG.copy()
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, REG_SETTINGS_PATH, 0, winreg.KEY_READ) as key:
                raw_cfg, _ = winreg.QueryValueEx(key, "SettingsJson")
                cfg.update(json.loads(raw_cfg))
        except Exception:
            pass

        if not cfg.get("outlook_email"):
            detected_ol = detect_default_outlook_email()
            if detected_ol:
                cfg["outlook_email"] = detected_ol
        return cfg

    def _on_toggle_fa_fmt(self):
        if hasattr(self, "opt_fa_fmt_menu"):
            self.opt_fa_fmt_menu.configure(state="normal" if self.var_force_fa_fmt.get() else "disabled")
        self.save_config()

    def test_google_connection(self, skip_wizard=False):
        c_path, c_src = resolve_credentials_path()
        if c_src != "data" or not c_path:
            self._update_oauth_main_button(is_ok=False)
            self._write_ui_log(r"Test Google échoué : fichier .\data\credentials.json absent ou invalide.", "Error")
            return False

        self._write_ui_log("Test de connexion au compte Google en cours...", "Information")
        def _bg_test():
            try:
                svc, creds, info_str = run_google_oauth_connect(c_path, log_callback=self._write_ui_log)
                self.google_service = svc
                self.google_creds = creds
                self._write_ui_log(f"Connexion Google People API réussie ({info_str}).", "Information")
                self.after(0, lambda: self._update_oauth_main_button(is_ok=True))
            except Exception as e:
                self._write_ui_log(f"Échec connexion Google : {e}", "Error")
                self.after(0, lambda: self._update_oauth_main_button(is_ok=False))

        threading.Thread(target=_bg_test, daemon=True).start()
        return True


    def _write_ui_log(self, msg, level="Information"):
        import datetime
        ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        line = f"{ts} | {level} | {msg}\n"
        def _do_insert():
            for attr in ("txt_log", "log_box", "log_text", "textbox_log", "log_textbox"):
                box = getattr(self, attr, None)
                if box is not None:
                    try:
                        box.configure(state="normal")
                        box.insert("end", line)
                        box.see("end")
                        box.configure(state="disabled")
                        return
                    except Exception:
                        pass
            for w in self.winfo_children():
                for child in getattr(w, "winfo_children", lambda: [])():
                    if isinstance(child, ctk.CTkTextbox):
                        try:
                            child.configure(state="normal")
                            child.insert("end", line)
                            child.see("end")
                            child.configure(state="disabled")
                            return
                        except Exception:
                            pass
        try:
            self.after(0, _do_insert)
        except Exception:
            pass

    def _update_oauth_main_button(self, is_ok=None):
        _raw_lang = str(self.config.get("language") or self.config.get("lang") or self.config.get("ui_lang") or "en").lower()
        is_fr = _raw_lang.startswith("fr")
        btn_txt = "Configurer la Connexion Google" if is_fr else "Configure Google Connection"
        if is_ok is None:
            c_path, c_src = resolve_credentials_path()
            token_p = os.path.join(get_portable_data_dir(), "google_token.json")
            is_ok = (c_src == "data" and os.path.isfile(token_p))

        if hasattr(self, "btn_config_oauth"):
            if is_ok:
                self.btn_config_oauth.configure(text=btn_txt, fg_color=("#d0d5dd", "#343b48"), hover_color=("#bfc6d0", "#434c5c"), text_color=("#1a202c", "#f1f5f9"))
            else:
                self.btn_config_oauth.configure(text=btn_txt, fg_color=("#1a73e8", "#1a73e8"), hover_color=("#1557b0", "#2563eb"), text_color=("#ffffff", "#ffffff"))
        if hasattr(self, "lbl_status_icon"):
            if is_ok:
                self.lbl_status_icon.configure(text="✓", text_color=("#1e7e34", "#28a745"))
            else:
                self.lbl_status_icon.configure(text="✗", text_color=("#c93b2b", "#ff4d4d"))

    def show_credentials_manager_dialog(self, error_reason="", auto_retest=False):
        lang = str(self.config.get("language") or self.config.get("lang") or self.config.get("ui_lang") or "en").lower()
        is_fr = lang.startswith("fr")

        existing = getattr(self, "_oauth_dlg", None)
        if existing is not None:
            try:
                if existing.winfo_exists():
                    if hasattr(existing, "_update_state"):
                        existing._update_state(error_reason)
                    existing.lift()
                    existing.focus_force()
                    return
            except Exception:
                pass

        dlg = ctk.CTkToplevel(self)
        self._oauth_dlg = dlg
        dlg.title("Configuration Google OAuth (credentials.json)" if is_fr else "Google OAuth Setup (credentials.json)")
        dlg.geometry("740x535")
        dlg.resizable(False, False)
        dlg.transient(self)
        dlg.grab_set()
        try:
            _ico = resource_path("SyncContactsOG.ico")
            if os.path.exists(_ico):
                dlg.iconbitmap(_ico)
                dlg.after(200, lambda: dlg.iconbitmap(_ico) if dlg.winfo_exists() else None)
        except Exception:
            pass

        def on_close():
            self._oauth_dlg = None
            if dlg.winfo_exists():
                dlg.destroy()

        dlg.protocol("WM_DELETE_WINDOW", on_close)

        self.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() - 740) // 2
        y = self.winfo_y() + (self.winfo_height() - 535) // 2
        dlg.geometry(f"+{max(0, x)}+{max(0, y)}")

        container = ctk.CTkFrame(dlg, fg_color="transparent")
        container.pack(fill="both", expand=True, padx=18, pady=14)

        top_card = ctk.CTkFrame(container, corner_radius=10)
        top_card.pack(fill="x", pady=(0, 10))

        ctk.CTkLabel(
            top_card,
            text=r"1. Gestion du fichier credentials.json & Authentification Google" if is_fr else r"1. Manage credentials.json & Google Authentication",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color=("#1a73e8", "#4da3ff"),
            anchor="w"
        ).pack(fill="x", padx=16, pady=(12, 6))

        lbl_diag_json = ctk.CTkLabel(top_card, text="", font=ctk.CTkFont(size=12, weight="bold"), wraplength=670, anchor="w", justify="left")
        lbl_diag_json.pack(fill="x", padx=16, pady=(0, 4))

        lbl_diag_auth = ctk.CTkLabel(top_card, text="", font=ctk.CTkFont(size=12, weight="bold"), wraplength=670, anchor="w", justify="left")
        lbl_diag_auth.pack(fill="x", padx=16, pady=(0, 8))

        ctk.CTkLabel(
            top_card,
            text=(
                r"Sélectionner un fichier credentials.json sur votre disque : il sera automatiquement validé et importé dans le sous-dossier .\data à côté de l'application."
                if is_fr else
                r"Select a credentials.json file on your disk: it will be validated and imported into the .\data subfolder next to the application."
            ),
            font=ctk.CTkFont(size=11),
            text_color=("#4a5568", "#a0aec0"),
            wraplength=670,
            anchor="w",
            justify="left"
        ).pack(fill="x", padx=16, pady=(0, 10))

        btn_row = ctk.CTkFrame(top_card, fg_color="transparent")
        btn_row.pack(fill="x", padx=16, pady=(0, 14))

        state_holder = {"err": error_reason, "auth_msg": "", "auth_ok": None}

        def refresh_status():
            c_path, c_src = resolve_credentials_path()
            local_p = get_portable_credentials_path()
            token_p = os.path.join(get_portable_data_dir(), "google_token.json")
            json_ok = (c_src == "data")

            if json_ok:
                lbl_diag_json.configure(
                    text=(r"Fichier JSON : Valide et actif dans .\data\credentials.json" if is_fr else r"JSON File: Valid and active in .\data\credentials.json"),
                    text_color=("#1e7e34", "#5cd67a")
                )
                btn_import.configure(fg_color=("#d0d5dd", "#343b48"), hover_color=("#bfc6d0", "#434c5c"), text_color=("#1a202c", "#f1f5f9"))
            else:
                lbl_diag_json.configure(
                    text=(r"Fichier JSON : Aucun fichier credentials.json valide trouvé dans .\data." if is_fr else r"JSON File: No valid credentials.json found in .\data."),
                    text_color=("#c93b2b", "#ff7b6b")
                )
                btn_import.configure(fg_color=("#1a73e8", "#1a73e8"), hover_color=("#1557b0", "#2563eb"), text_color=("#ffffff", "#ffffff"))

            auth_ready = False
            if state_holder["auth_ok"] is True:
                auth_ready = True
                lbl_diag_auth.configure(text=state_holder["auth_msg"], text_color=("#1e7e34", "#5cd67a"))
            elif state_holder["auth_ok"] is False:
                auth_ready = False
                lbl_diag_auth.configure(text=state_holder["auth_msg"], text_color=("#c93b2b", "#ff7b6b"))
            elif os.path.isfile(token_p) and json_ok:
                auth_ready = True
                lbl_diag_auth.configure(
                    text=("Authentification Google : Jeton OAuth enregistré (prêt)" if is_fr else "Google Authentication: OAuth token saved (ready)"),
                    text_color=("#1e7e34", "#5cd67a")
                )
            else:
                auth_ready = False
                lbl_diag_auth.configure(
                    text=("Authentification Google : Non connecté (cliquez sur 'Connecter au compte Google')" if is_fr else "Google Authentication: Not connected (click 'Connect Google Account')"),
                    text_color=("#d97706", "#fbbf24") if json_ok else ("#c93b2b", "#ff7b6b")
                )

            if not auth_ready:
                btn_connect.configure(fg_color=("#1a73e8", "#1a73e8"), hover_color=("#1557b0", "#2563eb"), text_color=("#ffffff", "#ffffff"))
            else:
                btn_connect.configure(fg_color=("#d0d5dd", "#343b48"), hover_color=("#bfc6d0", "#434c5c"), text_color=("#1a202c", "#f1f5f9"))

            has_any = json_ok or os.path.exists(local_p)
            btn_delete.configure(state="normal" if has_any else "disabled")
            btn_connect.configure(state="normal" if json_ok else "disabled")
            self._update_oauth_main_button(is_ok=(json_ok and auth_ready))

        def update_state(new_err):
            state_holder["err"] = new_err
            if new_err:
                state_holder["auth_ok"] = False
                state_holder["auth_msg"] = f"Authentification Google : {new_err}" if is_fr else f"Google Authentication: {new_err}"
            refresh_status()

        dlg._update_state = update_state

        def run_auth_in_modal():
            c_path, c_src = resolve_credentials_path()
            if c_src != "data" or not c_path:
                refresh_status()
                return
            btn_connect.configure(state="disabled")
            btn_import.configure(state="disabled")
            state_holder["auth_ok"] = None
            lbl_diag_auth.configure(
                text=("Authentification Google : Connexion et vérification People API en cours..." if is_fr else "Google Authentication: Connecting and verifying People API..."),
                text_color=("#1a73e8", "#4da3ff")
            )
            dlg.update_idletasks()

            def _worker():
                try:
                    svc, creds, info_str = run_google_oauth_connect(c_path, log_callback=self._write_ui_log)
                    self.google_service = svc
                    self.google_creds = creds
                    self._write_ui_log(f"Connexion Google People API réussie ({info_str}).", "Information")
                    def _on_ok():
                        self._update_oauth_main_button(is_ok=True)
                        if dlg.winfo_exists():
                            state_holder["auth_ok"] = True
                            state_holder["auth_msg"] = (
                                f"Authentification Google : Connecté ({info_str}) — People API OK"
                                if is_fr else
                                f"Google Authentication: Connected ({info_str}) — People API OK"
                            )
                            btn_connect.configure(state="normal")
                            btn_import.configure(state="normal")
                            refresh_status()
                    self.after(0, _on_ok)
                except Exception as exc:
                    err_s = str(exc)
                    self._write_ui_log(f"Échec authentification Google : {err_s}", "Error")
                    def _on_err():
                        self._update_oauth_main_button(is_ok=False)
                        if dlg.winfo_exists():
                            state_holder["auth_ok"] = False
                            state_holder["auth_msg"] = (
                                f"Authentification Google : Échec ({err_s})"
                                if is_fr else
                                f"Google Authentication: Failed ({err_s})"
                            )
                            btn_connect.configure(state="normal")
                            btn_import.configure(state="normal")
                            refresh_status()
                    self.after(0, _on_err)

            threading.Thread(target=_worker, daemon=True).start()

        def on_import_json():
            from tkinter import filedialog as fd, messagebox as mb
            dlg.grab_release()
            try:
                file_selected = fd.askopenfilename(
                    parent=dlg,
                    title="Sélectionner un fichier JSON" if is_fr else "Select a JSON file",
                    filetypes=[("Fichiers JSON", "*.json"), ("Tous les fichiers", "*.*")] if is_fr else [("JSON files", "*.json"), ("All files", "*.*")]
                )
            finally:
                if dlg.winfo_exists():
                    dlg.grab_set()
                    dlg.focus_force()

            if not file_selected:
                return

            ok, msg = validate_google_credentials_json(file_selected)
            if not ok:
                mb.showerror(
                    "Fichier JSON invalide" if is_fr else "Invalid JSON file",
                    (f"Le fichier sélectionné n'est pas un credentials.json OAuth Google valide :\n{msg}" if is_fr else f"The selected file is not a valid Google OAuth credentials.json:\n{msg}"),
                    parent=dlg
                )
                return

            dest_path = get_portable_credentials_path()
            shutil.copy2(file_selected, dest_path)
            token_p = os.path.join(get_portable_data_dir(), "google_token.json")
            if os.path.isfile(token_p):
                try:
                    os.remove(token_p)
                except Exception:
                    pass
            state_holder["err"] = ""
            state_holder["auth_ok"] = None
            refresh_status()
            dlg.update_idletasks()
            self._write_ui_log(f"credentials.json importé dans {dest_path}", "Information")
            run_auth_in_modal()

        def on_delete_json():
            from tkinter import messagebox as mb
            local_p = get_portable_credentials_path()
            token_p = os.path.join(get_portable_data_dir(), "google_token.json")
            for fp in (local_p, token_p):
                if os.path.exists(fp):
                    try:
                        os.remove(fp)
                    except Exception as e:
                        mb.showerror("Erreur" if is_fr else "Error", str(e), parent=dlg)
                        return
            self.google_service = None
            self.google_creds = None
            state_holder["err"] = ""
            state_holder["auth_ok"] = False
            state_holder["auth_msg"] = (
                "Authentification Google : Déconnecté (JSON et jeton supprimés)"
                if is_fr else
                "Google Authentication: Disconnected (JSON and token deleted)"
            )
            self._update_oauth_main_button(is_ok=False)
            self._write_ui_log(r"credentials.json et jeton OAuth supprimés (.\data).", "Information")
            refresh_status()

        btn_import = ctk.CTkButton(
            btn_row,
            text="Sélectionner un fichier JSON sur le disque..." if is_fr else "Select a JSON file from disk...",
            width=265,
            height=32,
            fg_color=("#d0d5dd", "#343b48"),
            hover_color=("#bfc6d0", "#434c5c"),
            text_color=("#1a202c", "#f1f5f9"),
            font=ctk.CTkFont(size=12, weight="bold"),
            command=on_import_json
        )
        btn_import.pack(side="left", padx=(0, 8))

        btn_connect = ctk.CTkButton(
            btn_row,
            text="Connecter au compte Google" if is_fr else "Connect Google Account",
            width=205,
            height=32,
            fg_color=("#d0d5dd", "#343b48"),
            hover_color=("#bfc6d0", "#434c5c"),
            text_color=("#1a202c", "#f1f5f9"),
            font=ctk.CTkFont(size=12, weight="bold"),
            command=run_auth_in_modal
        )
        btn_connect.pack(side="left", padx=(0, 8))

        btn_delete = ctk.CTkButton(
            btn_row,
            text="Supprimer le JSON existant" if is_fr else "Delete existing JSON",
            width=195,
            height=32,
            fg_color=("#d0d5dd", "#343b48"),
            hover_color=("#bfc6d0", "#434c5c"),
            text_color=("#1a202c", "#f1f5f9"),
            font=ctk.CTkFont(size=12, weight="bold"),
            command=on_delete_json
        )
        btn_delete.pack(side="left")

        guide_card = ctk.CTkFrame(container, corner_radius=10)
        guide_card.pack(fill="both", expand=True)

        ctk.CTkLabel(
            guide_card,
            text="2. Procédure pour créer et télécharger un nouveau credentials.json" if is_fr else "2. How to create and download a new credentials.json",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color=("#1a73e8", "#4da3ff"),
            anchor="w"
        ).pack(fill="x", padx=16, pady=(12, 6))

        guide_text = (
            "1. Ouvrez Google Cloud Console : https://console.cloud.google.com/\n"
            "2. Créez un projet (ou sélectionnez un projet existant).\n"
            "3. Menu 'API et services' > 'Bibliothèque' : recherchez 'Google People API' et cliquez sur 'Activer'.\n"
            "4. Menu 'API et services' > 'Écran de consentement OAuth' :\n"
            "   - Choisissez 'Externe', renseignez le nom de l'application et votre e-mail.\n"
            "   - Dans 'Utilisateurs tests', ajoutez votre adresse Gmail (indispensable en mode Test).\n"
            "5. Menu 'API et services' > 'Identifiants' > '+ Créer des identifiants' > 'ID client OAuth' :\n"
            "   - Type d'application : 'Application de bureau' (Desktop app).\n"
            "   - Cliquez sur 'Créer', puis sur 'Télécharger le fichier JSON'.\n"
            "6. Cliquez sur 'Sélectionner un fichier JSON sur le disque...' puis sur 'Connecter au compte Google' (la fenêtre reste ouverte jusqu'à validation)."
            if is_fr else
            "1. Open Google Cloud Console: https://console.cloud.google.com/\n"
            "2. Create a new project (or select an existing one).\n"
            "3. Go to 'APIs & Services' > 'Library': search for 'Google People API' and click 'Enable'.\n"
            "4. Go to 'APIs & Services' > 'OAuth consent screen':\n"
            "   - Select 'External', fill in the app name and your email.\n"
            "   - Under 'Test users', add your Gmail address (required while in Testing mode).\n"
            "5. Go to 'APIs & Services' > 'Credentials' > '+ Create Credentials' > 'OAuth client ID':\n"
            "   - Application type: 'Desktop app'.\n"
            "   - Click 'Create', then click 'Download JSON'.\n"
            "6. Click 'Select a JSON file from disk...' then 'Connect Google Account' (the window stays open until verified)."
        )

        txt_box = ctk.CTkTextbox(guide_card, font=ctk.CTkFont(size=12), wrap="word", height=205)
        txt_box.pack(fill="both", expand=True, padx=16, pady=(0, 12))
        txt_box.insert("1.0", guide_text)

        url_str = "https://console.cloud.google.com/"
        start_idx = txt_box.search(url_str, "1.0", stopindex="end")
        if start_idx:
            end_idx = f"{start_idx}+{len(url_str)}c"
            txt_box._textbox.tag_add("cloud_link", start_idx, end_idx)
            txt_box._textbox.tag_config("cloud_link", foreground="#4da3ff", underline=True)
            txt_box._textbox.tag_bind("cloud_link", "<Button-1>", lambda e: webbrowser.open(url_str))
            txt_box._textbox.tag_bind("cloud_link", "<Enter>", lambda e: txt_box._textbox.configure(cursor="hand2"))
            txt_box._textbox.tag_bind("cloud_link", "<Leave>", lambda e: txt_box._textbox.configure(cursor=""))

        txt_box.configure(state="disabled")
        refresh_status()

    def open_google_setup_wizard(self, on_done_callback=None):
        self._update_oauth_main_button(is_ok=False)
        return
        is_fr = (self.config.get("language", "EN") == "FR")
        win = ctk.CTkToplevel(self)
        win.title("Configuration initiale Google Contacts" if is_fr else "Google Contacts First-Time Setup")
        win.geometry("620x520")
        win.transient(self)
        win.grab_set()

        ctk.CTkLabel(
            win,
            text="Connexion sécurisée à Google Contacts (4 étapes simples)" if is_fr else "Secure Google Contacts Setup (4 Simple Steps)",
            font=ctk.CTkFont(size=15, weight="bold"),
            text_color=("#1f6aa5", "#4da6ff")
        ).pack(anchor="w", padx=18, pady=(16, 6))

        ctk.CTkLabel(
            win,
            text=(
                "Aucune clé Google valide n'a été trouvée ou la clé précédente a été révoquée.\n"
                "Importez un fichier client_secret_*.json : il sera chiffré dans votre Coffre-fort Windows."
                if is_fr else
                "No valid Google app key was found or the previous key was revoked.\n"
                "Import a client_secret_*.json file: it will be encrypted inside your Windows Credential Vault."
            ),
            font=ctk.CTkFont(size=11), justify="left", wraplength=580
        ).pack(anchor="w", padx=18, pady=(0, 12))

        box = ctk.CTkFrame(win, corner_radius=8)
        box.pack(fill="both", expand=True, padx=18, pady=(0, 12))

        steps_txt = (
            "Étape 1 : Cliquez sur le bouton bleu ci-dessous pour ouvrir Google Cloud Console.\n"
            "   • Créez un projet (ex: MesContacts) et activez « Google People API » dans Bibliothèque.\n\n"
            "Étape 2 : Dans « Google Auth Platform » → « Branding » puis « Audience » :\n"
            "   • Renseignez le nom et votre e-mail, choisissez « Externe » et ajoutez votre adresse\n"
            "     Gmail dans « Utilisateurs tests » (inutile de publier l'application).\n\n"
            "Étape 3 : Dans « Clients » (ou « Identifiants ») → « + Créer un client OAuth » :\n"
            "   • Choisissez « Application de bureau », cliquez sur Créer, puis « Télécharger le fichier JSON ».\n\n"
            "Étape 4 : Cliquez sur le bouton vert ci-dessous et sélectionnez le fichier téléchargé."
            if is_fr else
            "Step 1: Click the blue button below to open Google Cloud Console.\n"
            "   • Create a project (e.g. MyContacts), then enable 'Google People API' in Library.\n\n"
            "Step 2: Go to 'Google Auth Platform' -> 'Branding' then 'Audience':\n"
            "   • Enter app name & email, select 'External', and add your Gmail address\n"
            "     under 'Test users' (no need to publish the app to production).\n\n"
            "Step 3: Go to 'Clients' (or 'Credentials') -> '+ Create OAuth client':\n"
            "   • Select 'Desktop app', click Create, then click 'Download JSON'.\n\n"
            "Step 4: Click the green button below and select the downloaded JSON file."
        )
        ctk.CTkLabel(box, text=steps_txt, font=ctk.CTkFont(size=11), justify="left", wraplength=565).pack(anchor="w", padx=14, pady=12)

        btn_row = ctk.CTkFrame(win, fg_color="transparent")
        btn_row.pack(fill="x", padx=18, pady=(0, 16))

        def _open_console():
            import webbrowser, urllib.parse
            target_url = "https://console.cloud.google.com/apis/credentials"
            g_hint = self.var_g_email.get().strip() if hasattr(self, "var_g_email") else ""
            chooser_url = "https://accounts.google.com/AccountChooser?continue=" + urllib.parse.quote(target_url, safe="")
            if g_hint and "@" in g_hint:
                chooser_url += "&Email=" + urllib.parse.quote(g_hint, safe="")
            webbrowser.open(chooser_url)

        def _import_json():
            from tkinter import filedialog
            dl_dir = os.path.join(os.path.expanduser("~"), "Downloads")
            fpath = filedialog.askopenfilename(
                parent=win,
                initialdir=dl_dir if os.path.exists(dl_dir) else BASE_DIR,
                title="Sélectionner le fichier client_secret_*.json" if is_fr else "Select client_secret_*.json",
                filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
            )
            if not fpath:
                return
            try:
                with open(fpath, "r", encoding="utf-8") as f:
                    parsed = json.load(f)
                if "installed" not in parsed and "web" not in parsed:
                    messagebox.showerror(
                        "Fichier invalide" if is_fr else "Invalid file",
                        "Ce fichier JSON n'est pas un identifiant OAuth Google valide." if is_fr else "This JSON file is not a valid Google OAuth credential.",
                        parent=win
                    )
                    return
                save_google_client_config_dpapi(parsed)
                if messagebox.askyesno(
                    "Sécurité" if is_fr else "Security",
                    "Clé enregistrée et chiffrée dans le Coffre-fort Windows !\nVoulez-vous supprimer le fichier en clair des Téléchargements ?"
                    if is_fr else
                    "Key encrypted and saved in Windows Credential Vault!\nDelete the plain JSON file from Downloads?",
                    parent=win
                ):
                    try:
                        os.remove(fpath)
                    except Exception:
                        pass
                win.destroy()
                if on_done_callback:
                    on_done_callback()
            except Exception as ex:
                messagebox.showerror("Erreur" if is_fr else "Error", str(ex), parent=win)

        ctk.CTkButton(
            btn_row,
            text="1. Ouvrir Google Cloud Console" if is_fr else "1. Open Google Cloud Console",
            width=220, height=32, command=_open_console
        ).pack(side="left", padx=(0, 10))

        ctk.CTkButton(
            btn_row,
            text="2. Importer le JSON téléchargé" if is_fr else "2. Import Downloaded JSON",
            width=240, height=32, fg_color="#16A34A", hover_color="#15803D",
            command=_import_json
        ).pack(side="left")

    def save_config(self, *_args):
        self.config["language"] = self.lang
        if hasattr(self, "_update_oauth_main_button"):
            self._update_oauth_main_button()
        self.config["theme"] = self.theme
        self.config["outlook_email"] = self.var_ol_email.get().strip()
        self.config["outlook_connector"] = self.var_ol_connector.get()
        self.config["google_email"] = self.var_g_email.get().strip()
        self.config["sync_mode"] = self.var_sync_mode.get()
        self.config["use_file_as"] = self.var_file_as.get()
        if hasattr(self, "var_force_fa_fmt"):
            self.config["force_file_as_format"] = self.var_force_fa_fmt.get()
        if hasattr(self, "var_fa_fmt_style"):
            _v_style = self.var_fa_fmt_style.get()
            self.config["file_as_format_style"] = "SURNAME, Name" if _v_style in ("SURNAME, Name", "NOM, Prénom") else "Name SURNAME"
        self.config["force_rtf"] = self._saved_force_rtf if getattr(self, "_rtf_is_disabled", False) and hasattr(self, "_saved_force_rtf") else self.var_force_rtf.get()
        self.config["sync_photos"] = self.var_sync_photos.get()
        self.config["sync_deletions"] = self.var_sync_deletions.get()
        self.config["prompt_deletion"] = self.var_prompt_deletion.get()
        self.config["force_full_compare"] = self.var_force_full.get()
        self.config["extract_local_db"] = self.var_extract_db.get()
        self.config["dry_run"] = self.var_dry_run.get()
        self.config["minimize_to_tray"] = self.var_min_tray.get()
        self.config["start_with_windows"] = self.var_win_startup.get()

        try:
            self.config["auto_sync_hours"] = int(self.var_auto_hours.get())
        except ValueError:
            self.config["auto_sync_hours"] = 0

        try:
            self.config["max_deletions_threshold"] = int(self.var_max_del.get())
        except ValueError:
            self.config["max_deletions_threshold"] = 10

        try:
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, REG_SETTINGS_PATH) as key:
                winreg.SetValueEx(key, "SettingsJson", 0, winreg.REG_SZ, json.dumps(self.config))
        except Exception:
            pass

        set_windows_startup(self.config["start_with_windows"])
        self.schedule_auto_sync()

    # --------------------------------------------------------------------------
    # SYSTEM TRAY & FENETRE WINDOWS
    # --------------------------------------------------------------------------
    def create_tray_icon_image(self, has_exclusives=False):
        meipass_ico = os.path.join(getattr(sys, "_MEIPASS", BASE_DIR), "SyncContactsOG.ico")
        for cand in (meipass_ico, os.path.join(BASE_DIR, "SyncContactsOG.ico")):
            try:
                if os.path.exists(cand):
                    raw_ico = Image.open(cand)
                    if hasattr(raw_ico, "ico") and raw_ico.ico:
                        best_sz = max(raw_ico.ico.sizes())
                        raw_ico = raw_ico.ico.getimage(best_sz)
                    img = raw_ico.convert("RGBA")
                    if has_exclusives:
                        px = img.load()
                        w, h = img.size
                        for y in range(h):
                            for x in range(w):
                                r, g, b, a = px[x, y]
                                if a > 20 and b > r + 15:
                                    # Recolorer le bleu d'origine en orange vif (#EA580C)
                                    lum = max(b, g) / 255.0
                                    px[x, y] = (min(255, int(234 * lum + 20)), min(255, int(88 * lum + 10)), min(255, int(12 * lum)), a)
                    return img
            except Exception:
                pass
        img = Image.new("RGBA", (64, 64), (234, 88, 12, 255) if has_exclusives else (31, 106, 165, 255))
        return img

    def update_tray_icon_badge(self, excl_count=0):
        try:
            if getattr(self, "tray_icon", None) is not None:
                self.tray_icon.icon = self.create_tray_icon_image(has_exclusives=(excl_count > 0))
        except Exception:
            pass

    def send_windows_toast(self, title_txt, body_txt):
        def _worker():
            try:
                ps_script = (
                    "[Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] > $null;"
                    "$xml = [Windows.UI.Notifications.ToastNotificationManager]::GetTemplateContent([Windows.UI.Notifications.ToastTemplateType]::ToastText02);"
                    "$nodes = $xml.GetElementsByTagName('text');"
                    f"$nodes.Item(0).AppendChild($xml.CreateTextNode('{title_txt.replace(chr(39), chr(39)*2)}')) > $null;"
                    f"$nodes.Item(1).AppendChild($xml.CreateTextNode('{body_txt.replace(chr(39), chr(39)*2)}')) > $null;"
                    "$toast = [Windows.UI.Notifications.ToastNotification]::new($xml);"
                    "[Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier('SyncContactsOG').Show($toast);"
                )
                subprocess.run(
                    ["powershell", "-NoProfile", "-WindowStyle", "Hidden", "-Command", ps_script],
                    creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
                    timeout=8
                )
            except Exception:
                pass
        threading.Thread(target=_worker, daemon=True).start()

    def setup_system_tray(self):
        menu = pystray.Menu(pystray.MenuItem(lambda item: f"Nettoyer exclusifs ({len(getattr(self, '_pending_exclusives', ([],[]))[0]) + len(getattr(self, '_pending_exclusives', ([],[]))[1])})", lambda icon, item: self.after(0, self.open_pending_exclusives), visible=lambda item: bool(getattr(self, "_pending_exclusives", None))), 
            pystray.MenuItem("Ouvrir SyncContactsOG", self.restore_from_tray, default=True),
            pystray.MenuItem("▶ Synchroniser maintenant", lambda: self.after(0, self.start_sync_thread)),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Quitter", self.quit_application)
        )
        _init_excl = (getattr(self, "_last_excl_ol_count", 0) + getattr(self, "_last_excl_g_count", 0)) > 0
        self.tray_icon = pystray.Icon(
            APP_REG_NAME,
            self.create_tray_icon_image(has_exclusives=_init_excl),
            "SyncContactsOG",
            menu
        )
        self.update_tray_tooltip("Prêt")
        threading.Thread(target=self.tray_icon.run, daemon=True).start()

    def hide_to_tray(self):
        self.withdraw()

    def restore_from_tray(self, _icon=None, _item=None):
        def _show():
            self.deiconify()
            self.state("normal")
            self.lift()
            self.focus_force()
        self.after(0, _show)

    def on_window_unmap(self, event):
        if event.widget == self and self.state() in ("withdrawn", "iconic"):
            if self.config.get("minimize_to_tray", True):
                self.after(10, self.hide_to_tray)

    def on_closing(self):
        self.quit_application()

    def quit_application(self, _icon=None, _item=None):
        self.save_config()
        if self.tray_icon:
            self.tray_icon.stop()
        self.after(0, self.destroy)

    # --------------------------------------------------------------------------
    # CONSTRUCTION DE L'INTERFACE GRAPHIQUE (LARGEURS VERROUILLEES FR / EN)
    # --------------------------------------------------------------------------
    def clear_google_account(self):
        em = self.var_g_email.get().strip()
        if not messagebox.askyesno("Déconnecter Google", f"Supprimer '{em}' et révoquer l'autorisation d'accès Google ?\n\nUne nouvelle authentification et autorisation Google seront demandées lors de la prochaine connexion."):
            return
        if os.path.exists("token.json"):
            try:
                import urllib.request
                c = Credentials.from_authorized_user_file("token.json")
                tok = getattr(c, "refresh_token", None) or getattr(c, "token", None)
                if tok:
                    urllib.request.urlopen(f"https://oauth2.googleapis.com/revoke?token={tok}", timeout=5)
            except Exception:
                pass
            os.remove("token.json")
        self.var_g_email.set("")
        if hasattr(self, "ent_google_account"):
            self.ent_google_account.configure(state="normal")
            self.ent_google_account.focus_set()
        if hasattr(self, "btn_clear_google"):
            self.btn_clear_google.place_forget()
        if hasattr(self, "lbl_g_conn_status"):
            self.lbl_g_conn_status.configure(text="")
        try:
            delete_secret_from_vault(CRED_TARGET_NAME)
        except Exception:
            pass
        self.log("Information", f"Compte Google déconnecté ({em}) : autorisation révoquée et token.json supprimé.")

    def update_rtf_state(self, *_):
        try:
            dis = (self.var_sync_mode.get() == "OUTLOOK_TO_GOOGLE_ONLY")
            if not hasattr(self, "sw_force_rtf"):
                return
            if dis:
                if not getattr(self, "_rtf_is_disabled", False):
                    self._saved_force_rtf = bool(self.var_force_rtf.get())
                    self._rtf_is_disabled = True
                self.sw_force_rtf.configure(
                    state="disabled",
                    progress_color=("#9CA3AF", "#4B5563"),
                    fg_color=("#9CA3AF", "#4B5563"),
                    button_color=("#D1D5DB", "#9CA3AF"),
                    text_color=("#9CA3AF", "#6B7280")
                )
                self.sw_force_rtf._status = 0.5
                self.sw_force_rtf._draw()
            else:
                if getattr(self, "_rtf_is_disabled", False):
                    self._rtf_is_disabled = False
                    if hasattr(self, "_saved_force_rtf"):
                        self.var_force_rtf.set(self._saved_force_rtf)
                self.sw_force_rtf.configure(
                    state="normal",
                    progress_color=("#3B8ED0", "#1F6AA5"),
                    fg_color=("#939BA2", "#4A4D50"),
                    button_color=("#D5D9DE", "#D5D9DE"),
                    text_color=("#111827", "#F3F4F6")
                )
                self.sw_force_rtf._status = 1.0 if self.var_force_rtf.get() else 0.0
                self.sw_force_rtf._draw()
        except Exception:
            pass

    def show_exclusives_window(self, ol_only, g_only, ol_bk, g_srv, _skip_confirm=False):
        if not ol_only and not g_only:
            return
        if not _skip_confirm:
            if not messagebox.askyesno("Contacts exclusifs", f"Présents sur un seul compte :\n• Google uniquement : {len(g_only)}\n• Outlook uniquement : {len(ol_only)}\n\nAfficher la liste pour nettoyer ?"):
                return
        win = ctk.CTkToplevel(self)
        win.title("Contacts présents sur un seul compte")
        win.geometry("900x600")
        win.transient(self)
        win.grab_set()
        tabs = ctk.CTkTabview(win)
        tabs.pack(fill="both", expand=True, padx=12, pady=12)
        g_tab_name = f"Uniquement Google ({len(g_only)})"
        ol_tab_name = f"Uniquement Outlook ({len(ol_only)})"
        t_g = tabs.add(g_tab_name)
        t_o = tabs.add(ol_tab_name)

        def _delete_and_verify_outlook(entry_id):
            import win32com.client
            ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            deleted_folder_id = ns.GetDefaultFolder(3).EntryID
            if ol_bk is not None and hasattr(ol_bk, "delete_contact"):
                ol_bk.delete_contact(entry_id)
            else:
                item = ns.GetItemFromID(entry_id)
                item.Delete()
            try:
                chk = ns.GetItemFromID(entry_id)
                if chk is not None and getattr(chk.Parent, "EntryID", "") != deleted_folder_id:
                    raise RuntimeError("Le contact est toujours présent dans Outlook.")
            except Exception as ver_err:
                if "toujours présent" in str(ver_err):
                    raise

        def _delete_and_verify_google(service, resource_name):
            try:
                service.people().deleteContact(resourceName=resource_name).execute()
            except Exception as api_err:
                if "404" not in str(api_err) and "not found" not in str(api_err).lower():
                    raise
            try:
                still_there = service.people().get(resourceName=resource_name, personFields="names").execute()
                if still_there:
                    raise RuntimeError("Google renvoie encore ce contact après suppression.")
            except Exception as ver_err:
                if "404" in str(ver_err) or "not found" in str(ver_err).lower() or "requested entity was not found" in str(ver_err).lower():
                    return
                raise

        vmaps = {}
        del_btns = {}

        for tab, items, side in ((t_g, g_only, "g"), (t_o, ol_only, "ol")):
            sf = ctk.CTkScrollableFrame(tab)
            sf.pack(fill="both", expand=True, padx=6, pady=6)
            vmap = {}
            vmaps[side] = vmap

            def _populate(scroll_f, item_list, vm):
                for w in scroll_f.winfo_children():
                    w.destroy()
                vm.clear()
                for it in item_list:
                    var = ctk.BooleanVar(value=False)
                    cb = ctk.CTkCheckBox(scroll_f, text=f"{it.get('name','')} | {it.get('email','') or '—'} | {it.get('phone','') or '—'}", variable=var)
                    cb.pack(anchor="w", pady=3, padx=6)
                    vm[it["id"]] = (var, it, cb)
                scroll_f.update_idletasks()
                try:
                    scroll_f._parent_canvas.yview_moveto(0)
                    scroll_f._parent_canvas.configure(scrollregion=scroll_f._parent_canvas.bbox("all"))
                except Exception:
                    pass

            _populate(sf, items, vmap)

            def make_del(vm, sd, tab_items, scroll_frame):
                def _del(confirm=True):
                    nonlocal g_tab_name, ol_tab_name, g_srv
                    sel = [(cid, it, cb, v) for cid, (v, it, cb) in list(vm.items()) if v.get()]
                    if not sel:
                        return
                    if confirm:
                        if not messagebox.askyesno("Confirmer", f"Supprimer définitivement {len(sel)} contact(s) ?", parent=win):
                            return

                    if sd == "g" and g_srv is None:
                        try:
                            g_srv = self.get_google_service()
                        except Exception as e:
                            messagebox.showerror("Erreur Google", f"Connexion Google impossible : {e}", parent=win)
                            return
                        if g_srv is None:
                            messagebox.showerror("Erreur Google", "Service Google Contacts indisponible.", parent=win)
                            return

                    deleted_ids = set()
                    errors = []
                    for cid, it, cb, v in sel:
                        try:
                            if sd == "g":
                                _delete_and_verify_google(g_srv, cid)
                            else:
                                _delete_and_verify_outlook(cid)
                        except Exception as del_err:
                            errors.append(f"{it.get('name', cid)} : {del_err}")
                            continue

                        deleted_ids.add(cid)
                        # Griser/desactiver immediatement avant reconstruction de la liste
                        try:
                            v.set(False)
                            cb.configure(state="disabled", text_color="#6B7280", text=f"[SUPPRIMÉ] {it.get('name','')}")
                        except Exception:
                            pass

                    if deleted_ids:
                        tab_items[:] = [x for x in tab_items if x.get("id") not in deleted_ids]
                        # Reconstruire proprement le CTkScrollableFrame sans les contacts supprimes
                        _populate(scroll_frame, tab_items, vm)

                    if sd == "g":
                        new_g_name = f"Uniquement Google ({len(g_only)})"
                        if new_g_name != g_tab_name:
                            tabs.rename(g_tab_name, new_g_name)
                            g_tab_name = new_g_name
                    else:
                        new_ol_name = f"Uniquement Outlook ({len(ol_only)})"
                        if new_ol_name != ol_tab_name:
                            tabs.rename(ol_tab_name, new_ol_name)
                            ol_tab_name = new_ol_name

                    win.update_idletasks()
                    win.update()

                    self._last_excl_ol_count = len(ol_only)
                    self._last_excl_g_count = len(g_only)
                    self.config["last_excl_ol_count"] = len(ol_only)
                    self.config["last_excl_g_count"] = len(g_only)
                    self.config["last_excl_ol_items"] = ol_only
                    self.config["last_excl_g_items"] = g_only
                    self.save_config()
                    self.refresh_exclusives_card()
                    if deleted_ids:
                        self.log("Information", f"Supprimé et vérifié {len(deleted_ids)} contact(s) exclusif(s) côté {sd.upper()}.")
                    if confirm:
                        if errors:
                            messagebox.showerror("Suppression partielle", f"{len(deleted_ids)} supprimé(s), {len(errors)} erreur(s) :\n" + "\n".join(errors[:5]), parent=win)
                        elif deleted_ids:
                            messagebox.showinfo("Fait", f"{len(deleted_ids)} contact(s) supprimé(s) et vérifié(s) absent(s) sur {sd.upper()}.", parent=win)
                return _del

            fn = make_del(vmap, side, items, sf)
            del_btns[side] = fn
            def _toggle_all_exclusive(state_val):
                target_container = tab.winfo_toplevel()
                try:
                    if 'None' != 'None' and 'None' in locals():
                        cur_tab = None.get()
                        target_container = None.tab(cur_tab)
                except Exception:
                    pass
                def _walk_cb(w):
                    for ch in w.winfo_children():
                        if isinstance(ch, ctk.CTkCheckBox):
                            ch.select() if state_val else ch.deselect()
                        _walk_cb(ch)
                _walk_cb(target_container)

            _is_fr = (getattr(self, "lang", "fr") == "fr") if hasattr(self, "lang") else True
            _btn_bar = ctk.CTkFrame(tab, fg_color="transparent")
            _btn_bar.pack(pady=(6, 12))
            ctk.CTkButton(
                _btn_bar,
                text="Tout sélectionner" if _is_fr else "Select All",
                fg_color=("#d5d9df", "#3a3f47"),
                hover_color=("#c3c9d1", "#4b525c"),
                text_color=("#1f2328", "#f0f3f6"),
                width=145,
                command=lambda: _toggle_all_exclusive(True)
            ).pack(side="left", padx=6)
            ctk.CTkButton(
                _btn_bar,
                text="Tout désélectionner" if _is_fr else "Deselect All",
                fg_color=("#d5d9df", "#3a3f47"),
                hover_color=("#c3c9d1", "#4b525c"),
                text_color=("#1f2328", "#f0f3f6"),
                width=155,
                command=lambda: _toggle_all_exclusive(False)
            ).pack(side="left", padx=6)
            ctk.CTkButton(_btn_bar, text="Supprimer la sélection", fg_color="#B91C1C", command=fn).pack(side="left", padx=6)

        win._test_vmaps = vmaps
        win._test_del_btns = del_btns
        win._test_tabs = tabs
        return win

    def update_mirror_state(self, *_):
        try:
            is_one_way = (self.var_sync_mode.get() in ("OUTLOOK_TO_GOOGLE_ONLY", "GOOGLE_TO_OUTLOOK_ONLY"))
            if hasattr(self, "chk_force_mirror"):
                self.chk_force_mirror.configure(
                    state="normal" if is_one_way else "disabled",
                    text_color=("#111827", "#F3F4F6") if is_one_way else ("#9CA3AF", "#6B7280")
                )
        except Exception:
            pass

    def notify_exclusives_tray(self, ol_only, g_only, ol_bk, g_srv):
        self._pending_exclusives = (ol_only, g_only, ol_bk, g_srv)
        tot = len(ol_only) + len(g_only)
        self.update_tray_icon_badge(tot)
        msg = f"{tot} contact(s) exclusif(s) détecté(s) (Outlook: {len(ol_only)}, Google: {len(g_only)})."
        try:
            if hasattr(self, "tray_icon") and self.tray_icon is not None:
                self.tray_icon.notify(msg, "SyncContactsOG — Contacts exclusifs")
        except Exception:
            pass
        self.send_windows_toast("SyncContactsOG — Contacts exclusifs", msg)

    def open_pending_exclusives(self):
        try:
            self.deiconify()
            self.lift()
            self.focus_force()
        except Exception:
            pass
        pend = getattr(self, "_pending_exclusives", None)
        if pend:
            self._pending_exclusives = None
            self.show_exclusives_window(*pend)

    def refresh_ol_account_dropdown(self, *_):
        try:
            if not hasattr(self, "cmb_ol_account"):
                return
            vals = []
            conn_var = getattr(self, "var_ol_connector", None)
            mode_conn = conn_var.get() if conn_var is not None else self.config.get("outlook_connector", "COM")
            if mode_conn == "GRAPH":
                for k in ("m365_email", "graph_email", "outlook_email"):
                    v = str(self.config.get(k, "") or "").strip()
                    if v and v not in vals:
                        vals.append(v)
                u_attr = str(getattr(self, "m365_user_email", "") or "").strip()
                if u_attr and u_attr not in vals:
                    vals.append(u_attr)
            else:
                vals = get_outlook_available_accounts()
            cur = self.var_ol_email.get().strip()
            if cur and cur not in vals:
                vals.insert(0, cur)
            if not vals:
                vals = [""]
            self.cmb_ol_account.configure(values=vals)
            if not cur and vals[0]:
                self.var_ol_email.set(vals[0])
        except Exception:
            pass
        # FORCE_START_IN_TRAY_WHEN_AUTOSTART
        if self.config.get("start_with_windows", False) or "--minimized" in sys.argv or "--tray" in sys.argv:
            self.config["minimize_to_tray"] = True
            if hasattr(self, "var_min_tray"):
                self.var_min_tray.set(True)
            self.after(50, self.withdraw)

    def _resolve_last_sync_display(self):
        val = getattr(self, "_last_completed_sync_str", None) or self.config.get("last_completed_sync", "")
        if val:
            return val
        for db_cand, sql in (
            (DB_FILE, "SELECT MAX(last_synced) FROM contact_map"),
            (os.path.join(self.get_archive_dir(), "contacts.db"), "SELECT MAX(extracted_at) FROM metadata"),
        ):
            try:
                if os.path.exists(db_cand):
                    c = sqlite3.connect(db_cand)
                    row = c.execute(sql).fetchone()
                    c.close()
                    if row and row[0]:
                        raw_ts = str(row[0]).strip()
                        try:
                            dt_parsed = datetime.fromisoformat(raw_ts.replace("Z", "+00:00"))
                            val = dt_parsed.strftime("%Y/%m/%d %H:%M")
                        except Exception:
                            val = raw_ts[:16].replace("T", " ")
                        self._last_completed_sync_str = val
                        self.config["last_completed_sync"] = val
                        return val
            except Exception:
                pass
        return "Jamais"

    def _resolve_max_photo_kb(self):
        kb = getattr(self, "_max_photo_kb", None)
        if kb is not None:
            return kb
        db_arch = os.path.join(self.get_archive_dir(), "contacts.db")
        try:
            if os.path.exists(db_arch):
                c = sqlite3.connect(db_arch)
                m1 = c.execute("SELECT COALESCE(MAX(LENGTH(photo_blob)), 0) FROM outlook_contacts").fetchone()[0]
                m2 = c.execute("SELECT COALESCE(MAX(LENGTH(photo_blob)), 0) FROM google_contacts").fetchone()[0]
                c.close()
                mx = max(m1 or 0, m2 or 0)
                if mx > 0:
                    self._max_photo_kb = round(mx / 1024.0, 1)
                    return self._max_photo_kb
        except Exception:
            pass
        return None

    def refresh_photo_label(self):
        try:
            if "opt_photos" in getattr(self, "switch_widgets", {}):
                base_lbl = self.t("opt_photos")
                mx_kb = self._resolve_max_photo_kb()
                full_lbl = f"{base_lbl} (max detected: {mx_kb:.1f} Ko)" if mx_kb else base_lbl
                self.switch_widgets["opt_photos"].configure(text=full_lbl)
        except Exception:
            pass

    def _load_persisted_exclusives(self):
        if getattr(self, "_excl_loaded_once", False):
            return
        self._excl_loaded_once = True
        ol_saved = self.config.get("last_excl_ol_items", [])
        g_saved = self.config.get("last_excl_g_items", [])
        cfg_ol_c = int(self.config.get("last_excl_ol_count", 0))
        cfg_g_c = int(self.config.get("last_excl_g_count", 0))

        if (cfg_ol_c > 0 or cfg_g_c > 0) and (
            all(isinstance(x, dict) and "id" in x and "email" in x for x in ol_saved)
            and all(isinstance(x, dict) and "id" in x and "email" in x for x in g_saved)
        ):
            self._last_excl_ol_count = cfg_ol_c
            self._last_excl_g_count = cfg_g_c
            self._pending_exclusives = (ol_saved, g_saved, None, None)
            return

        # Repli SQLite exact sur contacts.db vs sync_state.db
        db_arch = os.path.join(self.get_archive_dir(), "contacts.db")
        try:
            if not (os.path.exists(db_arch) and os.path.exists(DB_FILE)):
                return
            c_map = sqlite3.connect(DB_FILE)
            mapped_ol = {r[0] for r in c_map.execute("SELECT outlook_entry_id FROM contact_map WHERE outlook_entry_id IS NOT NULL").fetchall()}
            mapped_g = {r[0] for r in c_map.execute("SELECT google_resource_name FROM contact_map WHERE google_resource_name IS NOT NULL").fetchall()}
            c_map.close()

            def _first_json_val(raw_j):
                if not raw_j:
                    return ""
                try:
                    arr = json.loads(raw_j)
                    if isinstance(arr, list) and arr:
                        v = arr[0]
                        return v.get("value", "") if isinstance(v, dict) else str(v)
                except Exception:
                    pass
                return ""

            c_arch = sqlite3.connect(db_arch)
            ol_rows = c_arch.execute(
                "SELECT entry_id, COALESCE(NULLIF(file_as,''), TRIM(COALESCE(first_name,'')||' '||COALESCE(last_name,'')), entry_id), emails_json, phones_json FROM outlook_contacts"
            ).fetchall()
            g_rows = c_arch.execute(
                "SELECT resource_name, COALESCE(NULLIF(file_as,''), TRIM(COALESCE(first_name,'')||' '||COALESCE(last_name,'')), resource_name), emails_json, phones_json FROM google_contacts"
            ).fetchall()
            c_arch.close()

            ol_excl = [
                {"id": r[0], "name": r[1] or r[0], "email": _first_json_val(r[2]), "phone": _first_json_val(r[3]), "raw": {"entry_id": r[0], "name": r[1] or r[0]}}
                for r in ol_rows if r[0] and r[0] not in mapped_ol
            ]
            g_excl = [
                {"id": r[0], "name": r[1] or r[0], "email": _first_json_val(r[2]), "phone": _first_json_val(r[3]), "raw": {"resourceName": r[0], "name": r[1] or r[0]}}
                for r in g_rows if r[0] and r[0] not in mapped_g
            ]
            self._last_excl_ol_count = len(ol_excl)
            self._last_excl_g_count = len(g_excl)
            if ol_excl or g_excl:
                self._pending_exclusives = (ol_excl, g_excl, None, None)
        except Exception:
            pass

    def refresh_exclusives_card(self):
        try:
            self._load_persisted_exclusives()
            ol_c = getattr(self, "_last_excl_ol_count", 0)
            g_c = getattr(self, "_last_excl_g_count", 0)
            tot = ol_c + g_c
            btn_lbl = "Nettoyer" if self.lang == "FR" else "Clean"

            if hasattr(self, "lbl_ol_excl"):
                self.lbl_ol_excl.configure(
                    text=self.t("lbl_excl_ol").format(n=ol_c),
                    text_color=("#D97706", "#FB923C") if ol_c > 0 else ("#111827", "#FFFFFF")
                )
            if hasattr(self, "btn_clean_ol_excl"):
                self.btn_clean_ol_excl.configure(text=btn_lbl)
                if ol_c > 0:
                    self.btn_clean_ol_excl.pack(anchor="w", pady=(4, 2))
                else:
                    self.btn_clean_ol_excl.pack_forget()

            if hasattr(self, "lbl_g_excl"):
                self.lbl_g_excl.configure(
                    text=self.t("lbl_excl_g").format(n=g_c),
                    text_color=("#D97706", "#FB923C") if g_c > 0 else ("#111827", "#FFFFFF")
                )
            if hasattr(self, "btn_clean_g_excl"):
                self.btn_clean_g_excl.configure(text=btn_lbl)
                if g_c > 0:
                    self.btn_clean_g_excl.pack(anchor="w", pady=(4, 2))
                else:
                    self.btn_clean_g_excl.pack_forget()

            self.update_tray_icon_badge(tot)
            self.update_tray_tooltip()
        except Exception:
            pass
    def open_exclusives_from_card(self):
        pend = getattr(self, "_pending_exclusives", None)
        if not pend or len(pend) != 4:
            return
        ol_ex, g_ex, ol_bk, g_srv = pend
        if g_srv is None:
            try:
                g_srv = self.get_google_service()
            except Exception:
                pass
        self.show_exclusives_window(ol_ex, g_ex, ol_bk, g_srv)

    def _get_last_ok_str(self):
        if getattr(self, "_last_ok_str", ""):
            return self._last_ok_str
        for dbp, sql in (
            (DB_FILE, "SELECT MAX(last_synced) FROM contact_map"),
            (os.path.join(self.get_archive_dir(), "contacts.db"), "SELECT MAX(extracted_at) FROM metadata"),
        ):
            try:
                if os.path.exists(dbp):
                    c = sqlite3.connect(dbp)
                    val = c.execute(sql).fetchone()[0]
                    c.close()
                    if val:
                        s = str(val).strip()
                        # Essayer plusieurs formats courants (ISO, SQL, FR)
                        from datetime import datetime
                        parsed = None
                        for fmt in ("%Y-%m-%dT%H:%M:%S.%f", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y/%m/%d %H:%M:%S", "%Y/%m/%d %H:%M"):
                            try:
                                parsed = datetime.strptime(s[:26], fmt)
                                break
                            except Exception:
                                pass
                        self._last_ok_str = parsed.strftime("%Y/%m/%d %H:%M") if parsed else s[:16]
                        return self._last_ok_str
            except Exception:
                pass
        return ""

    def update_tray_tooltip(self, status_text=None):
        try:
            if getattr(self, "tray_icon", None) is None:
                return
            if status_text is not None:
                self._last_tray_status = str(status_text).strip()
            st = getattr(self, "_last_tray_status", None) or ("Sync ✓ — Prêt" if self.lang == "FR" else "Sync ✓ — Ready")
            if st in ("Prêt", "Ready"):
                st = f"Sync ✓ — {st}"
            tot_excl = getattr(self, "_last_excl_ol_count", 0) + getattr(self, "_last_excl_g_count", 0)
            if tot_excl > 0 and "Excl" not in st:
                st = f"{st} | Excl.: {tot_excl}"
            last_ok = self._get_last_ok_str()
            tip = f"{st}\nLast OK: {last_ok}" if last_ok else st
            self.tray_icon.title = tip[:127]
        except Exception:
            pass

    def build_ui(self):
        # Colonne gauche strictement figée à 660px (aucun redimensionnement selon la longueur des textes FR/EN)
        self.grid_columnconfigure(0, weight=0, minsize=660)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        left_panel = ctk.CTkFrame(self, width=646, corner_radius=12, fg_color=("gray92", "#1e222a"))
        left_panel.grid(row=0, column=0, sticky="nsew", padx=(14, 7), pady=14)
        left_panel.grid_propagate(False)
        left_panel.grid_columnconfigure(0, weight=1)

        header_row = ctk.CTkFrame(left_panel, fg_color="transparent")
        header_row.pack(fill="x", padx=14, pady=(12, 10))

        self.lbl_app_title = ctk.CTkLabel(header_row, text="", font=ctk.CTkFont(size=18, weight="bold"))
        self.lbl_app_title.pack(side="left")

        right_controls = ctk.CTkFrame(header_row, fg_color="transparent")
        right_controls.pack(side="right")

        self.lang_seg = ctk.CTkSegmentedButton(
            right_controls, values=["FR", "EN"], width=68, height=26,
            command=self.change_language
        )
        self.lang_seg.set(self.lang)
        self.lang_seg.pack(side="left", padx=(0, 8))

        self.theme_seg = ctk.CTkSegmentedButton(
            right_controls, values=[self.t("theme_light"), self.t("theme_dark")],
            width=120, height=26, command=self.change_theme
        )
        self.theme_seg.set(self.t("theme_dark") if self.theme == "Dark" else self.t("theme_light"))
        self.theme_seg.pack(side="left")

        # Grille supérieure : 2 colonnes fixes de 305px chacune
        top_grid = ctk.CTkFrame(left_panel, fg_color="transparent")
        top_grid.pack(fill="x", pady=4)
        top_grid.grid_columnconfigure(0, weight=0, minsize=311)
        top_grid.grid_columnconfigure(1, weight=1, minsize=311)

        # --- Carte 1 : Comptes & Profils (largeur fixe 305px) ---
        card_acc = ctk.CTkFrame(top_grid, width=305, height=324, corner_radius=10)
        card_acc.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        card_acc.pack_propagate(False)
        card_acc.grid_propagate(False)
        card_acc.pack_propagate(False)

        self.lbl_sec_accounts = ctk.CTkLabel(card_acc, text="", font=ctk.CTkFont(size=13, weight="bold"), text_color=("#1f6aa5", "#4da6ff"))
        self.lbl_sec_accounts.pack(anchor="w", padx=12, pady=(10, 4))

        self.lbl_ol_email = ctk.CTkLabel(card_acc, text="", font=ctk.CTkFont(size=11))
        self.lbl_ol_email.pack(anchor="w", padx=12)
        self.var_ol_email = tk.StringVar(value=self.config.get("outlook_email", ""))
        _ol_list = get_outlook_available_accounts()
        if self.config.get("outlook_email") and self.config["outlook_email"] not in _ol_list:
            _ol_list.insert(0, self.config["outlook_email"])
        if not _ol_list:
            _ol_list = [""]
        self.cmb_ol_account = ctk.CTkComboBox(
            card_acc,
            variable=self.var_ol_email,
            values=_ol_list,
            height=28,
            state="normal",
            command=lambda _v: self._on_email_changed()
        )
        self.cmb_ol_account.pack(fill="x", padx=12, pady=(2, 6))
        self.cmb_ol_account.bind("<FocusOut>", self._on_email_changed)

        self.var_ol_connector = tk.StringVar(value=self.config.get("outlook_connector", "COM"))
        self.var_ol_connector.trace_add("write", self.refresh_ol_account_dropdown)
        self.after(200, self.refresh_ol_account_dropdown)
        self.rb_ol_com = CleanRadioButton(
            card_acc, text=("Outlook Classic COM" if self.config.get("language", "EN") == "EN" else "Outlook Classique COM") + " (...)",
            variable=self.var_ol_connector, value="COM",
            command=self.save_config, font=ctk.CTkFont(size=11), size=15
        )
        self.rb_ol_com.pack(anchor="w", padx=12, pady=2)

        self.rb_ol_graph = CleanRadioButton(
            card_acc, text="Microsoft 365 Graph API (...)",
            variable=self.var_ol_connector, value="GRAPH",
            command=self.save_config, font=ctk.CTkFont(size=11), size=15
        )
        self.rb_ol_graph.pack(anchor="w", padx=12, pady=(2, 6))

        self.lbl_g_email = ctk.CTkLabel(card_acc, text="", font=ctk.CTkFont(size=11))
        self.lbl_g_email.pack(anchor="w", padx=12)
        self.var_g_email = tk.StringVar(value=self.config["google_email"])
        g_wrap = ctk.CTkFrame(card_acc, fg_color="transparent")
        g_wrap.pack(fill="x", padx=12, pady=(2, 6))
        self.ent_google_account = ctk.CTkEntry(g_wrap, textvariable=self.var_g_email, height=28, placeholder_text="ex: utilisateur@gmail.com")
        self.ent_google_account.pack(fill="x")
        self.btn_clear_google = ctk.CTkButton(g_wrap, text="✕", width=22, height=20, corner_radius=4, fg_color="transparent", hover_color=("#D1D5DB", "#374151"), font=ctk.CTkFont(size=12, weight="bold"), command=self.clear_google_account)
        if self.var_g_email.get().strip():
            self.ent_google_account.configure(state="readonly")
            self.btn_clear_google.place(relx=1.0, rely=0.5, x=-6, anchor="e")
        self.ent_google_account.bind("<FocusOut>", self.save_config)

        row_g_test = ctk.CTkFrame(card_acc, fg_color="transparent")
        row_g_test.pack(fill="x", padx=12, pady=(0, 8))
        _lang0 = str(self.config.get("language") or self.config.get("lang") or "en").lower()
        self.btn_config_oauth = ctk.CTkButton(
            row_g_test,
            text="Configurer la Connexion Google" if _lang0.startswith("fr") else "Configure Google Connection",
            width=205,
            height=28,
            fg_color=("#d0d5dd", "#343b48"),
            hover_color=("#bfc6d0", "#434c5c"),
            text_color=("#1a202c", "#f1f5f9"),
            font=ctk.CTkFont(size=11, weight="bold"),
            command=lambda: self.show_credentials_manager_dialog(auto_retest=False)
        )
        self.btn_config_oauth.pack(side="left", padx=(0, 8))

        self.lbl_status_icon = ctk.CTkLabel(
            row_g_test,
            text="",
            font=ctk.CTkFont(size=15, weight="bold"),
            text_color=("#1e7e34", "#28a745")
        )
        self.lbl_status_icon.pack(side="left")
        self._update_oauth_main_button()
        self.lbl_g_conn_status = ctk.CTkLabel(
            row_g_test, text="", font=ctk.CTkFont(size=11, weight="bold"),
            text_color=("#16A34A", "#4ADE80")
        )
        self.lbl_g_conn_status.pack(side="left", padx=(8, 0))

        auto_row = ctk.CTkFrame(card_acc, fg_color="transparent")
        auto_row.pack(anchor="w", padx=12, pady=(2, 14))
        self.lbl_auto_sync = ctk.CTkLabel(auto_row, text="", font=ctk.CTkFont(size=11))
        self.lbl_auto_sync.pack(side="left", padx=(0, 6))
        HoverToolTip(self.lbl_auto_sync, lambda: self.t("tip_auto_sync"))

        self.var_auto_hours = tk.StringVar(value=str(self.config["auto_sync_hours"]))
        CompactVerticalSpinbox(auto_row, variable=self.var_auto_hours, min_val=0, max_val=168, command=self.save_config).pack(side="left")

        # --- Carte 2 : Mode de Synchronisation (largeur fixe 305px) ---
        card_mode = ctk.CTkFrame(top_grid, corner_radius=10)
        card_mode.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        card_mode.pack_propagate(True)
        card_mode.grid_propagate(True)
        card_mode.pack_propagate(True)

        self.lbl_sec_mode = ctk.CTkLabel(card_mode, text="", font=ctk.CTkFont(size=13, weight="bold"), text_color=("#1f6aa5", "#4da6ff"))
        self.lbl_sec_mode.pack(anchor="w", padx=12, pady=(10, 6))

        self.var_sync_mode = tk.StringVar(value=self.config["sync_mode"])
        self.radio_widgets = {}
        self.radio_sublabels = {}
        self.mode_status_badges = {}
        mode_defs = [
            ("mode_prompt", "mode_prompt_sub", "tip_mode_prompt", "MERGE_PROMPT"),
            ("mode_ol_wins", "mode_ol_wins_sub", "tip_mode_ol_wins", "MERGE_OUTLOOK_WINS"),
            ("mode_g_wins", "mode_g_wins_sub", "tip_mode_g_wins", "MERGE_GOOGLE_WINS"),
            ("mode_ol_to_g", "mode_ol_to_g_sub", "tip_mode_ol_to_g", "OUTLOOK_TO_GOOGLE_ONLY"),
            ("mode_g_to_ol", "mode_g_to_ol_sub", "tip_mode_g_to_ol", "GOOGLE_TO_OUTLOOK_ONLY"),
        ]
        for label_key, sub_key, tip_key, val in mode_defs:
            item_box = ctk.CTkFrame(card_mode, fg_color="transparent")
            item_box.pack(anchor="w", fill="x", padx=12, pady=(2, 2))

            top_line = ctk.CTkFrame(item_box, fg_color="transparent")
            top_line.pack(fill="x", anchor="w")

            rb = CleanRadioButton(
                top_line, text="", variable=self.var_sync_mode, value=val,
                command=self.save_config, font=ctk.CTkFont(size=11, weight="bold"), size=15
            )
            rb.pack(side="left", anchor="w", pady=0)

            is_tested = (val == "OUTLOOK_TO_GOOGLE_ONLY")
            badge = ctk.CTkLabel(
                top_line,
                text="✔ Testé" if is_tested else "⚠ Non testé",
                width=84,
                height=19,
                corner_radius=10,
                font=ctk.CTkFont(size=10, weight="normal"),
                fg_color=("#E5E7EB", "#2B303A"),
                text_color=("#4B5563", "#9CA3AF"),
                anchor="center"
            )
            badge.pack(side="right", padx=(4, 0))

            if is_tested:
                HoverToolTip(
                    badge,
                    lambda: "Mode de synchronisation validé et testé." if self.lang == "FR" else "Synchronization mode validated and tested."
                )
            else:
                HoverToolTip(
                    badge,
                    lambda: "Attention, mode de synchronisation non encore testé" if self.lang == "FR" else "Warning: synchronization mode not yet tested"
                )

            sub = ctk.CTkLabel(
                item_box, text="", font=ctk.CTkFont(size=10),
                text_color=("gray40", "#94a3b8"), height=12
            )
            sub.pack(anchor="w", padx=(24, 0), pady=0)
            HoverToolTip(rb, lambda tk_key=tip_key: self.t(tk_key))
            self.radio_widgets[label_key] = rb
            self.radio_sublabels[sub_key] = sub
            self.mode_status_badges[val] = badge

        self.btn_sync = ctk.CTkButton(
            card_mode, text="", width=220, height=34,
            font=ctk.CTkFont(size=13, weight="bold"), fg_color="#1f6aa5", hover_color="#144870",
            command=self.start_sync_thread
        )
        self.btn_sync.pack(anchor="center", expand=True, pady=(8, 10))
        self.var_force_mirror = tk.BooleanVar(value=False)

        # Grille inférieure : 2 colonnes fixes (395px + 227px)
        bottom_grid = ctk.CTkFrame(left_panel, fg_color="transparent")
        bottom_grid.pack(fill="x", pady=(8, 4))
        bottom_grid.grid_columnconfigure(0, weight=0, minsize=395)
        bottom_grid.grid_columnconfigure(1, weight=1, minsize=227)

        # --- Carte 3 : Options & Intégrité des données (largeur fixe 389px) ---
        card_opts = ctk.CTkFrame(bottom_grid, width=389, corner_radius=10)
        card_opts.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        bottom_grid.grid_rowconfigure(0, weight=1)
        card_opts.grid_propagate(False)

        self.lbl_sec_opts = ctk.CTkLabel(card_opts, text="", font=ctk.CTkFont(size=13, weight="bold"), text_color=("#1f6aa5", "#4da6ff"))
        self.lbl_sec_opts.pack(anchor="w", padx=12, pady=(10, 6))

        self.var_sync_photos = tk.BooleanVar(value=self.config["sync_photos"])
        self.var_file_as = tk.BooleanVar(value=self.config["use_file_as"])
        self.var_force_fa_fmt = tk.BooleanVar(value=self.config.get("force_file_as_format", False))
        self.var_fa_fmt_style = tk.StringVar(value=self.config.get("file_as_format_style", "Name SURNAME"))
        self.var_force_rtf = tk.BooleanVar(value=self.config["force_rtf"])
        self.var_sync_mode.trace_add("write", self.update_rtf_state)
        self.after(80, self.update_rtf_state)
        self.var_sync_deletions = tk.BooleanVar(value=self.config["sync_deletions"])
        self.var_prompt_deletion = tk.BooleanVar(value=self.config["prompt_deletion"])
        self.var_force_full = tk.BooleanVar(value=self.config["force_full_compare"])
        self.var_extract_db = tk.BooleanVar(value=self.config["extract_local_db"])
        self.var_dry_run = tk.BooleanVar(value=self.config["dry_run"])

        self.switch_widgets = {}
        switch_defs = [
            ("opt_photos", "tip_opt_photos", self.var_sync_photos),
            ("opt_file_as", "tip_opt_file_as", self.var_file_as),
            ("opt_force_rtf", "tip_opt_force_rtf", self.var_force_rtf),
            ("opt_del", "tip_opt_del", self.var_sync_deletions),
            ("opt_prompt_del", "tip_opt_prompt_del", self.var_prompt_deletion),
            ("opt_force_full", "tip_opt_force_full", self.var_force_full),
            ("opt_extract_db", "tip_opt_extract_db", self.var_extract_db),
            ("opt_dry_run", "tip_opt_dry_run", self.var_dry_run),
        ]
        for label_key, tip_key, var in switch_defs:
            sw = ctk.CTkSwitch(
                card_opts, text="", variable=var,
                command=self.save_config, font=ctk.CTkFont(size=11)
            )
            if var is self.var_force_rtf: self.sw_force_rtf = sw
            sw.pack(anchor="w", padx=12, pady=2)
            HoverToolTip(sw, lambda tk_key=tip_key: self.t(tk_key))
            self.switch_widgets[label_key] = sw
            if label_key == "opt_file_as":
                row_fa_fmt = ctk.CTkFrame(card_opts, fg_color="transparent")
                row_fa_fmt.pack(fill="x", padx=12, pady=2)
                self.sw_force_fa_fmt = ctk.CTkSwitch(
                    row_fa_fmt, text="", variable=self.var_force_fa_fmt,
                    command=self._on_toggle_fa_fmt, font=ctk.CTkFont(size=11)
                )
                self.sw_force_fa_fmt.pack(side="left")
                HoverToolTip(self.sw_force_fa_fmt, lambda: self.t("tip_opt_force_fa_fmt"))
                self.switch_widgets["opt_force_fa_fmt"] = self.sw_force_fa_fmt
                self.opt_fa_fmt_menu = ctk.CTkOptionMenu(row_fa_fmt, values=["Name SURNAME", "SURNAME, Name"],
                    variable=self.var_fa_fmt_style, width=130, height=22,
                    font=ctk.CTkFont(size=10),
                    command=lambda _v: self.save_config(), fg_color=("#F8FAFC", "#374151"), button_color=("#E2E8F0", "#4B5563"), button_hover_color=("#CBD5E1", "#6B7280"), text_color=("#1F2937", "#F3F4F6"))
                self.opt_fa_fmt_menu.pack(side="left", padx=(6, 0))
                self.opt_fa_fmt_menu.configure(state="normal" if self.var_force_fa_fmt.get() else "disabled", fg_color=("#F8FAFC", "#374151"), button_color=("#E2E8F0", "#4B5563"), button_hover_color=("#CBD5E1", "#6B7280"), text_color=("#1F2937", "#F3F4F6"))

        del_row = ctk.CTkFrame(card_opts, fg_color="transparent")
        del_row.pack(anchor="w", padx=12, pady=(8, 14))
        self.lbl_max_del = ctk.CTkLabel(del_row, text="", font=ctk.CTkFont(size=11))
        self.lbl_max_del.pack(side="left", padx=(0, 6))
        HoverToolTip(self.lbl_max_del, lambda: self.t("tip_max_del"))

        self.var_max_del = tk.StringVar(value=str(self.config["max_deletions_threshold"]))

        # --- Carte 4a : Intégration Windows (raccourcie en hauteur par le dessous : 112px) ---
        win_stack = ctk.CTkFrame(bottom_grid, height=305, fg_color="transparent")
        win_stack.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        bottom_grid.grid_rowconfigure(0, weight=1)
        win_stack.pack_propagate(False)

        card_win = ctk.CTkFrame(win_stack, height=112, corner_radius=10)
        card_win.pack(fill="x", pady=(0, 8))
        card_win.pack_propagate(False)

        self.lbl_sec_win = ctk.CTkLabel(card_win, text="", font=ctk.CTkFont(size=13, weight="bold"), text_color=("#1f6aa5", "#4da6ff"))
        self.lbl_sec_win.pack(anchor="w", padx=12, pady=(8, 4))

        self.var_min_tray = tk.BooleanVar(value=self.config["minimize_to_tray"])
        self.var_win_startup = tk.BooleanVar(value=self.config["start_with_windows"])

        self.sw_min_tray = ctk.CTkSwitch(
            card_win, text="", variable=self.var_min_tray,
            command=self.save_config, font=ctk.CTkFont(size=11)
        )
        self.sw_min_tray.pack(anchor="w", padx=12, pady=2)
        HoverToolTip(self.sw_min_tray, lambda: self.t("tip_win_min_tray"))

        self.sw_win_startup = ctk.CTkSwitch(
            card_win, text="", variable=self.var_win_startup,
            command=self.save_config, font=ctk.CTkFont(size=11)
        )
        self.sw_win_startup.pack(anchor="w", padx=12, pady=(2, 8))
        HoverToolTip(self.sw_win_startup, lambda: self.t("tip_win_startup"))

        # --- Carte 4b : Exclusive contacts (dans l'espace libéré dessous : 185px) ---
        card_excl = ctk.CTkFrame(win_stack, corner_radius=10)
        card_excl.pack(fill="both", expand=True, pady=(6, 0))

        self.lbl_sec_excl = ctk.CTkLabel(
            card_excl, text=self.t("sec_excl"),
            font=ctk.CTkFont(size=13, weight="bold"), text_color=("#1f6aa5", "#4da6ff")
        )
        self.lbl_sec_excl.pack(anchor="w", padx=12, pady=(10, 8))

        row_ol_excl = ctk.CTkFrame(card_excl, fg_color="transparent")
        row_ol_excl.pack(fill="x", padx=12, pady=(2, 6))
        self.lbl_ol_excl = ctk.CTkLabel(
            row_ol_excl, text=self.t("lbl_excl_ol").format(n=0),
            font=ctk.CTkFont(size=11), text_color=("#111827", "#FFFFFF"),
            anchor="w", wraplength=146, justify="left"
        )
        self.lbl_ol_excl.pack(anchor="w")
        self.btn_clean_ol_excl = ctk.CTkButton(row_ol_excl, text="Nettoyer", height=24,
            font=ctk.CTkFont(size=10, weight="bold"),
            fg_color="#D97706", hover_color="#B45309",
            command=self.open_exclusives_from_card, width=120)

        row_g_excl = ctk.CTkFrame(card_excl, fg_color="transparent")
        row_g_excl.pack(fill="x", padx=12, pady=(2, 8))
        self.lbl_g_excl = ctk.CTkLabel(
            row_g_excl, text=self.t("lbl_excl_g").format(n=0),
            font=ctk.CTkFont(size=11), text_color=("#111827", "#FFFFFF"),
            anchor="w", wraplength=146, justify="left"
        )
        self.lbl_g_excl.pack(anchor="w")
        self.btn_clean_g_excl = ctk.CTkButton(row_g_excl, text="Nettoyer", height=24,
            font=ctk.CTkFont(size=10, weight="bold"),
            fg_color="#D97706", hover_color="#B45309",
            command=self.open_exclusives_from_card, width=120)



        # ======================================================================
        # PANNEAU DROIT : JOURNAL VERTICAL PLEINE HAUTEUR
        # ======================================================================
        right_panel = ctk.CTkFrame(self, corner_radius=12, fg_color=("gray92", "#161920"))
        right_panel.grid(row=0, column=1, sticky="nsew", padx=(7, 14), pady=14)
        right_panel.grid_rowconfigure(2, weight=1)
        right_panel.grid_columnconfigure(0, weight=1)

        log_header = ctk.CTkFrame(right_panel, fg_color="transparent")
        log_header.grid(row=0, column=0, sticky="ew", padx=16, pady=(14, 6))

        self.lbl_sec_log = ctk.CTkLabel(log_header, text="", font=ctk.CTkFont(size=16, weight="bold"))
        self.lbl_sec_log.pack(side="left")

        self.btn_clear_log = ctk.CTkButton(
            log_header, text="", width=115, height=28,
            fg_color=("gray70", "#333842"), hover_color=("gray60", "#454c59"),
            text_color=("gray10", "white"), command=self.clear_log_view
        )
        self.btn_clear_log.pack(side="right")

        status_box = ctk.CTkFrame(right_panel, corner_radius=8, fg_color=("gray85", "#1e222a"))
        status_box.grid(row=1, column=0, sticky="ew", padx=16, pady=(0, 10))
        right_panel.grid_rowconfigure(1, weight=0)
        right_panel.grid_rowconfigure(2, weight=1)

        self.lbl_progress = ctk.CTkLabel(
            status_box, text=self.t("status_ready"),
            font=ctk.CTkFont(family="Consolas", size=12, weight="bold"),
            text_color=("#1f6aa5", "#63b3ed")
        )
        self.lbl_progress.pack(anchor="w", padx=12, pady=(8, 4))

        self.progress_bar = ctk.CTkProgressBar(status_box, height=10, corner_radius=5)
        self.progress_bar.set(0)
        self.progress_bar.pack(fill="x", padx=12, pady=(0, 6))

        self.lbl_sub_progress = ctk.CTkLabel(
            status_box,
            text=self.t("sub_ready"),
            font=ctk.CTkFont(size=11),
            text_color=("#6B7280", "#9CA3AF"),
            anchor="w"
        )
        self.lbl_sub_progress.pack(fill="x", padx=12, pady=(2, 2))

        self.sub_progress_bar = ctk.CTkProgressBar(status_box, height=6, corner_radius=3, progress_color=("#60A5FA", "#38BDF8"))
        self.sub_progress_bar.set(0)
        self.sub_progress_bar.pack(fill="x", padx=12, pady=(0, 10))

        self.txt_log = ctk.CTkTextbox(
            right_panel, wrap="word", font=ctk.CTkFont(family="Consolas", size=12),
            fg_color=("white", "#0f1116"), text_color=("#1e293b", "#e2e8f0"), corner_radius=8
        )
        self.txt_log.grid(row=2, column=0, sticky="nsew", padx=16, pady=(0, 16))
        self.txt_log.configure(state="disabled")
        self.apply_log_colors()

    def _on_email_changed(self, _event=None):
        self.save_config()
        threading.Thread(target=self.refresh_outlook_counts_async, daemon=True).start()

    def refresh_outlook_counts_async(self):
        ol_email = self.var_ol_email.get().strip()
        com_count = count_outlook_com_contacts(ol_email)
        com_str = f"{com_count:,}".replace(",", " ") if com_count is not None else "indisponible"
        self.after(0, lambda: self.rb_ol_com.configure(text=f("Outlook Classic COM" if self.config.get("language", "EN") == "EN" else "Outlook Classique COM") + " ({com_str})"))

        graph_count = count_outlook_graph_contacts()
        graph_str = f"{graph_count:,}".replace(",", " ") if graph_count is not None else "non connecté"
        self.after(0, lambda: self.rb_ol_graph.configure(text=f"Microsoft 365 Graph API ({graph_str})"))

    def apply_log_colors(self):
        if self.theme == "Dark":
            self.txt_log._textbox.tag_config("INFO", foreground="#cbd5e1")
            self.txt_log._textbox.tag_config("WARNING", foreground="#f6ad55")
            self.txt_log._textbox.tag_config("ERROR", foreground="#fc8181")
            self.txt_log._textbox.tag_config("SUCCESS", foreground="#68d391")
        else:
            self.txt_log._textbox.tag_config("INFO", foreground="#1e293b")
            self.txt_log._textbox.tag_config("WARNING", foreground="#c05621")
            self.txt_log._textbox.tag_config("ERROR", foreground="#c53030")
            self.txt_log._textbox.tag_config("SUCCESS", foreground="#2f855a")

    def apply_translations(self):
        self.title(f"SyncContactsOG — v{__version__}")
        self.lbl_app_title.configure(text=self.t("app_title"))

        light_lbl, dark_lbl = self.t("theme_light"), self.t("theme_dark")
        self.theme_seg.configure(values=[light_lbl, dark_lbl])
        self.theme_seg.set(dark_lbl if self.theme == "Dark" else light_lbl)

        self.lbl_sec_accounts.configure(text=self.t("sec_accounts"))
        self.lbl_ol_email.configure(text=self.t("lbl_ol_email"))
        self.lbl_g_email.configure(text=self.t("lbl_g_email"))
        self.lbl_auto_sync.configure(text=self.t("lbl_auto_sync"))
        self.lbl_sec_mode.configure(text=self.t("sec_mode"))
        for key, rb in self.radio_widgets.items():
            rb.configure(text=self.t(key))
        for key, sub in self.radio_sublabels.items():
            sub.configure(text=self.t(key))
        if hasattr(self, "mode_status_badges"):
            for m_val, b_lbl in self.mode_status_badges.items():
                if m_val == "OUTLOOK_TO_GOOGLE_ONLY":
                    b_lbl.configure(text="✔ Testé" if self.lang == "FR" else "✔ Tested")
                else:
                    b_lbl.configure(text="⚠ Non testé" if self.lang == "FR" else "⚠ Untested")
        self.lbl_sec_opts.configure(text=self.t("sec_opts"))
        for key, sw in self.switch_widgets.items():
            sw.configure(text=self.t(key))
        self.lbl_max_del.configure(text=self.t("lbl_max_del"))
        self.refresh_photo_label()
        if hasattr(self, "lbl_sec_excl"):
            self.lbl_sec_excl.configure(text=self.t("sec_excl"))
        self.refresh_exclusives_card()
        self.lbl_sec_win.configure(text=self.t("sec_win"))
        self.sw_min_tray.configure(text=self.t("win_min_tray"))
        self.sw_win_startup.configure(text=self.t("win_startup"))
        self.btn_sync.configure(text=self.t("btn_start"))
        self.lbl_sec_log.configure(text=self.t("sec_log"))
        self.btn_clear_log.configure(text=self.t("btn_clear_log"))
        if hasattr(self, "btn_test_google"):
            pass
        if hasattr(self, "opt_fa_fmt_menu") and hasattr(self, "var_fa_fmt_style"):
            _cur_style = self.config.get("file_as_format_style", "Name SURNAME")
            if self.lang == "FR":
                _vals = ["Prénom NOM", "NOM, Prénom"]
                _disp = "NOM, Prénom" if _cur_style in ("SURNAME, Name", "NOM, Prénom") else "Prénom NOM"
            else:
                _vals = ["Name SURNAME", "SURNAME, Name"]
                _disp = "SURNAME, Name" if _cur_style in ("SURNAME, Name", "NOM, Prénom") else "Name SURNAME"
            self.opt_fa_fmt_menu.configure(values=_vals)
            self.var_fa_fmt_style.set(_disp)
        if not self.is_syncing:
            self.lbl_progress.configure(text=self.t("status_ready"))

    def change_language(self, new_lang):
        self.lang = new_lang
        self.apply_translations()
        self.save_config()

    def change_theme(self, selected_label):
        self.theme = "Dark" if selected_label == self.t("theme_dark") else "Light"
        ctk.set_appearance_mode(self.theme)
        self.apply_log_colors()
        self.save_config()

    def clear_log_view(self):
        self.txt_log.configure(state="normal")
        self.txt_log.delete("1.0", "end")
        self.txt_log.configure(state="disabled")

    # --------------------------------------------------------------------------
    # JOURNALISATION & INTERACTIONS THREAD-SAFE
    # --------------------------------------------------------------------------
    def log(self, level, message):
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        line = f"{ts} | {level} | {message}\n"

        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line)

        tag = "INFO"
        if level == "Warning":
            tag = "WARNING"
        elif level == "Error":
            tag = "ERROR"
        elif "complete" in message.lower():
            tag = "SUCCESS"

        def _append():
            self.txt_log.configure(state="normal")
            self.txt_log._textbox.insert("end", line, tag)
            self.txt_log.see("end")
            self.txt_log.configure(state="disabled")

        self.after(0, _append)


    def set_step_progress(self, pct, step_num, total_steps, title_fr, title_en, sub_fr="", sub_en="", sub_ratio=None):
        is_en = (self.config.get("language", "EN") == "EN")
        step_word = "Step" if is_en else "Étape"
        title = title_en if is_en else title_fr
        main_txt = f"{int(pct)}% — {step_word} {step_num}/{total_steps} : {title}"
        sub_txt = ""
        if sub_fr or sub_en:
            sub_prefix = "Sub-step: " if is_en else "Sous-étape : "
            sub_txt = sub_prefix + (sub_en if is_en else sub_fr)
        def _apply():
            self.lbl_progress.configure(text=main_txt)
            self.update_tray_tooltip(main_txt)
            self.progress_bar.set(max(0.0, min(1.0, pct / 100.0)))
            if sub_txt:
                self.lbl_sub_progress.configure(text=sub_txt)
            if sub_ratio is not None:
                self.sub_progress_bar.set(max(0.0, min(1.0, float(sub_ratio))))
        self.after(0, _apply)

    def update_status_line(self, text, current=None, total=None):
        def _upd():
            prefix = "Sub-step: " if self.config.get("language", "EN") == "EN" else "Sous-étape : "
            clean_txt = text if text.startswith(("Sub-step:", "Sous-étape")) else f"{prefix}{text}"
            self.lbl_sub_progress.configure(text=clean_txt)
            if current is not None and total and total > 0:
                self.sub_progress_bar.set(max(0.0, min(1.0, current / total)))
        self.after(0, _upd)

    def ask_user_yes_no_threadsafe(self, title, prompt_text):
        event = threading.Event()
        result = {"answer": False}

        def _ask():
            self.restore_from_tray()
            result["answer"] = messagebox.askyesno(title, prompt_text, parent=self)
            event.set()

        self.after(0, _ask)
        event.wait()
        return result["answer"]

    def ask_conflict_choice_threadsafe(self, display_name, ol_canon, g_canon):
        event = threading.Event()
        result = {"choice": "SKIP"}

        def _dialog():
            self.restore_from_tray()
            dlg = ctk.CTkToplevel(self)
            dlg.title(self.t("conflict_title"))
            dlg.geometry("560x320")
            dlg.transient(self)
            dlg.grab_set()

            ctk.CTkLabel(dlg, text=f"{self.t('conflict_header')} {display_name}", font=ctk.CTkFont(size=14, weight="bold")).pack(pady=(14, 8), padx=16, anchor="w")

            info = (
                f"OUTLOOK :\n"
                f"  Société : {ol_canon['company']} | Poste : {ol_canon['job_title']}\n"
                f"  Tél : {', '.join(p['value'] for p in ol_canon['phones'])}\n"
                f"  Email : {', '.join(e['value'] for e in ol_canon['emails'])}\n\n"
                f"GOOGLE :\n"
                f"  Société : {g_canon['company']} | Poste : {g_canon['job_title']}\n"
                f"  Tél : {', '.join(p['value'] for p in g_canon['phones'])}\n"
                f"  Email : {', '.join(e['value'] for e in g_canon['emails'])}"
            )
            box = ctk.CTkTextbox(dlg, height=160, font=ctk.CTkFont(family="Consolas", size=11))
            box.pack(fill="x", padx=16, pady=6)
            box.insert("1.0", info)
            box.configure(state="disabled")

            btn_row = ctk.CTkFrame(dlg, fg_color="transparent")
            btn_row.pack(pady=12)

            def _set(val):
                result["choice"] = val
                dlg.destroy()
                event.set()

            ctk.CTkButton(btn_row, text="Outlook Wins", width=130, command=lambda: _set("OUTLOOK")).pack(side="left", padx=8)
            ctk.CTkButton(btn_row, text="Google Wins", width=130, command=lambda: _set("GOOGLE")).pack(side="left", padx=8)
            ctk.CTkButton(btn_row, text=self.t("btn_skip"), width=130, fg_color="#4a4e58", command=lambda: _set("SKIP")).pack(side="left", padx=8)
            dlg.protocol("WM_DELETE_WINDOW", lambda: _set("SKIP"))

        self.after(0, _dialog)
        event.wait()
        return result["choice"]

    def schedule_auto_sync(self):
        if self.auto_sync_job:
            self.after_cancel(self.auto_sync_job)
            self.auto_sync_job = None
        hours = self.config.get("auto_sync_hours", 0)
        if hours and hours > 0:
            self.auto_sync_job = self.after(hours * 3600 * 1000, self.start_sync_thread)

    
    def get_google_service(self, google_email=None):
        creds = None
        raw_token = load_secret_from_vault(CRED_TARGET_NAME)
        if raw_token:
            try:
                creds = Credentials.from_authorized_user_info(json.loads(raw_token), SCOPES)
            except Exception:
                creds = None

        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                try:
                    creds.refresh(Request())
                except RefreshError:
                    self.log("Warning", "Google OAuth token expired/revoked -> restarting login flow.")
                    delete_secret_from_vault(CRED_TARGET_NAME)
                    flow = InstalledAppFlow.from_client_config(load_google_client_config_dpapi(), SCOPES)
                    creds = flow.run_local_server(port=0, prompt='consent select_account')
            else:
                flow = InstalledAppFlow.from_client_config(load_google_client_config_dpapi(), SCOPES)
                creds = flow.run_local_server(port=0, prompt='consent select_account')
            save_secret_to_vault(CRED_TARGET_NAME, creds.to_json())

        return build("people", "v1", credentials=creds)

    # --------------------------------------------------------------------------
    # ARCHIVAGE INCREMENTAL (.db avec photos + .xlsx) : DELTA / JOUR, FULL > 7 JOURS
    # --------------------------------------------------------------------------
    def get_archive_dir(self) -> str:
        custom_dir = (self.config.get("custom_archive_dir") or "").strip()
        if custom_dir:
            return custom_dir
        default_dir = os.path.join(BASE_DIR, "contacts")
        for legacy_fn in ("contacts.db", "contacts.xlsx", "contacts_NOT_UP_TO_DATE.db", "contacts_NOT_UP_TO_DATE.xlsx"):
            old_f = os.path.join(BASE_DIR, legacy_fn)
            if os.path.exists(old_f) and is_dir_writable(default_dir):
                new_f = os.path.join(default_dir, legacy_fn)
                if not os.path.exists(new_f):
                    try:
                        shutil.move(old_f, new_f)
                    except Exception:
                        pass
        return default_dir

    def ensure_archive_dir_ready(self, interactive: bool = True):
        """Vérifie les droits d'écriture du dossier d'archive; demande à l'utilisateur si lecture seule."""
        custom_dir = (self.config.get("custom_archive_dir") or "").strip()
        if custom_dir:
            if is_dir_writable(custom_dir):
                return custom_dir
            if not interactive:
                return None

        default_dir = self.get_archive_dir()
        if is_dir_writable(default_dir):
            return default_dir

        if not interactive:
            return None

        from tkinter import filedialog
        is_fr = (self.config.get("language", "EN") == "FR")
        ask_title = "Dossier d'archive en lecture seule" if is_fr else "Read-Only Application Folder"
        ask_msg = (
            f"Le dossier de l'application ({BASE_DIR}) ne dispose pas des droits d'écriture pour créer l'archive locale.\n\n"
            "Souhaitez-vous stocker automatiquement une base locale dans un autre dossier ?"
            if is_fr else
            f"The application folder ({BASE_DIR}) is not writable to store the local contacts archive.\n\n"
            "Would you like to automatically store the local archive in another folder?"
        )
        if messagebox.askyesno(ask_title, ask_msg):
            chosen = filedialog.askdirectory(
                title="Choisir le dossier de l'archive locale" if is_fr else "Select Local Archive Folder"
            )
            if chosen and is_dir_writable(chosen):
                target_sub = os.path.join(chosen, "contacts")
                os.makedirs(target_sub, exist_ok=True)
                self.config["custom_archive_dir"] = target_sub
                self.save_config()
                self.log("Information", f"Dossier d'archive locale personnalisé configuré : {target_sub}")
                return target_sub
            else:
                messagebox.showwarning(
                    "Dossier invalide" if is_fr else "Invalid Folder",
                    "Aucun dossier accessible en écriture n'a été sélectionné. L'archive locale est désactivée."
                    if is_fr else
                    "No writable folder was selected. Local archive has been disabled."
                )

        self.config["extract_local_db"] = False
        if hasattr(self, "var_extract_db"):
            self.var_extract_db.set(False)
        self.save_config()
        self.log("Warning", "Archive locale désactivée (dossier en lecture seule refusé par l'utilisateur).")
        return None

    def perform_local_db_extraction(self, ol_items, g_raw_list, id_to_label, service, cfg, force_full_refresh=False):
        archive_dir = self.ensure_archive_dir_ready(interactive=False)
        if not archive_dir:
            self.log("Warning", "Archive locale ignorée : aucun dossier accessible en écriture.")
            return

        restore_stale_archive_if_needed(archive_dir)

        final_db_path = os.path.join(archive_dir, "contacts.db")
        tmp_db_path = os.path.join(tempfile.gettempdir(), "contacts_work.db")
        export_xlsx_path = os.path.join(archive_dir, "contacts.xlsx")

        cur_db_sz = os.path.getsize(final_db_path) if os.path.exists(final_db_path) else 0
        cur_xl_sz = os.path.getsize(export_xlsx_path) if os.path.exists(export_xlsx_path) else 0
        required_bytes = max(50 * 1024 * 1024, cur_db_sz + cur_xl_sz + 25 * 1024 * 1024)

        if not has_enough_disk_space(archive_dir, required_bytes):
            renamed = mark_archive_not_up_to_date(archive_dir)
            self.log(
                "Warning",
                f"Espace disque insuffisant dans '{archive_dir}' (< {required_bytes // (1024*1024)} Mo libres). "
                f"Archive marquée NOT_UP_TO_DATE : {', '.join(renamed) if renamed else 'aucun fichier existant'}."
            )
            return

        if not os.path.exists(final_db_path):
            force_full_refresh = True

        mode_label = "FULL WEEKLY" if force_full_refresh else "INCREMENTAL DELTA"
        self.log("Information", f"Updating local archive (.db with photos + .xlsx) [{mode_label}]...")

        if os.path.exists(final_db_path) and not force_full_refresh:
            try:
                _chk_conn = sqlite3.connect(final_db_path)
                _chk_res = _chk_conn.execute("PRAGMA integrity_check").fetchone()
                _chk_conn.close()
                if not _chk_res or _chk_res[0] != "ok":
                    raise sqlite3.DatabaseError("integrity_check failed")
                shutil.copy2(final_db_path, tmp_db_path)
            except Exception as _db_err:
                self.log("Warning", f"Local archive DB corrupted ({_db_err}); rebuilding full archive cleanly...")
                try:
                    os.remove(final_db_path)
                except Exception:
                    pass
                force_full_refresh = True
                mode_label = "FULL WEEKLY"
        elif os.path.exists(tmp_db_path):
            os.remove(tmp_db_path)

        exp_conn = sqlite3.connect(tmp_db_path)
        cur = exp_conn.cursor()
        cur.execute("PRAGMA synchronous = OFF")
        cur.executescript("""
            CREATE TABLE IF NOT EXISTS metadata (
                extracted_at TEXT, mode TEXT, outlook_account TEXT, google_account TEXT,
                outlook_count INTEGER, google_count INTEGER
            );
            CREATE TABLE IF NOT EXISTS outlook_contacts (
                entry_id TEXT PRIMARY KEY, data_hash TEXT, file_as TEXT, first_name TEXT, middle_name TEXT, last_name TEXT,
                prefix TEXT, suffix TEXT, company TEXT, job_title TEXT, department TEXT,
                emails_json TEXT, phones_json TEXT, addresses_json TEXT, birthday TEXT,
                website TEXT, notes TEXT, categories TEXT, custom_fields_json TEXT, photo_blob BLOB
            );
            CREATE TABLE IF NOT EXISTS google_contacts (
                resource_name TEXT PRIMARY KEY, data_hash TEXT, file_as TEXT, first_name TEXT, middle_name TEXT, last_name TEXT,
                prefix TEXT, suffix TEXT, company TEXT, job_title TEXT, department TEXT,
                emails_json TEXT, phones_json TEXT, addresses_json TEXT, birthday TEXT,
                website TEXT, notes TEXT, labels TEXT, custom_fields_json TEXT, photo_blob BLOB
            );
            CREATE TABLE IF NOT EXISTS sync_links (outlook_entry_id TEXT, google_resource_name TEXT, last_synced TEXT);
        """)

        existing_ol = dict(cur.execute("SELECT entry_id, data_hash FROM outlook_contacts").fetchall())
        current_ol_ids = set(ol_items.keys())
        deleted_ol_ids = set(existing_ol.keys()) - current_ol_ids
        if deleted_ol_ids:
            cur.executemany("DELETE FROM outlook_contacts WHERE entry_id=?", [(eid,) for eid in deleted_ol_ids])

        arch_fast_cache = {}
        if os.path.exists(DB_FILE):
            try:
                _fc = sqlite3.connect(DB_FILE)
                arch_fast_cache = {r[0]: (r[1], r[2]) for r in _fc.execute("SELECT outlook_entry_id, last_mod_ts, canon_json FROM ol_fast_cache").fetchall()}
                _fc.close()
            except Exception:
                arch_fast_cache = {}

        self.set_step_progress(
            22, 2, 5,
            f"Archive locale ({mode_label})", f"Local archive ({mode_label})",
            f"Vérification Outlook : 0/{len(ol_items)}...", f"Checking Outlook: 0/{len(ol_items)}...",
            0.0
        )
        ol_updated_count = 0
        total_ol = len(ol_items)
        for idx, (eid, item) in enumerate(ol_items.items(), start=1):
            if idx == 1 or idx % 200 == 0 or idx == total_ol:
                pct_now = 22 + int(4 * (idx / max(1, total_ol)))
                self.set_step_progress(
                    pct_now, 2, 5,
                    f"Archive locale ({mode_label})", f"Local archive ({mode_label})",
                    f"Archive Outlook : {idx}/{total_ol} ({int(idx*100/max(1,total_ol))}%)",
                    f"Outlook archive: {idx}/{total_ol} ({int(idx*100/max(1,total_ol))}%)",
                    idx / max(1, total_ol)
                )
            try:
                mod_ts = str(getattr(item, "LastModificationTime", "") or "")
            except Exception:
                mod_ts = ""
            cached_arch = arch_fast_cache.get(eid)
            if not force_full_refresh and mod_ts and cached_arch and cached_arch[0] == mod_ts and cached_arch[1]:
                try:
                    c = json.loads(cached_arch[1])
                except Exception:
                    c = outlook_to_canonical(item, use_file_as=True)
            else:
                c = outlook_to_canonical(item, use_file_as=True)

            has_pic = bool(getattr(item, "HasPicture", False))
            row_hash = compute_hash({**c, "_has_pic": has_pic})

            if not force_full_refresh and existing_ol.get(eid) == row_hash:
                continue

            if idx % 100 == 0 or idx == total_ol:
                pct_g = int(95 + 2 * (idx / max(1, total_ol)))
                self.after(0, lambda p=pct_g, i=idx, t=total_ol, m=mode_label: (
                    self.lbl_progress.configure(text=f"{p}% — Étape 5/5 : Archive locale Outlook ({m})"),
                    self.progress_bar.set(p / 100.0),
                    self.lbl_sub_progress.configure(text=f"Archive Outlook ({m}) : {i}/{t} ({int(i*100/max(1,t))}%)"),
                    self.sub_progress_bar.set(i / max(1, t))
                ))

            photo_blob = get_outlook_photo_bytes(item) if (cfg["sync_photos"] and has_pic) else None
            cat_str = ", ".join(c["categories"])
            cur.execute(
                "INSERT OR REPLACE INTO outlook_contacts VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                (
                    eid, row_hash, c["file_as"], c["first_name"], c["middle_name"], c["last_name"],
                    c["prefix"], c["suffix"], c["company"], c["job_title"], c["department"],
                    json.dumps(c["emails"], ensure_ascii=False), json.dumps(c["phones"], ensure_ascii=False),
                    json.dumps(c["addresses"], ensure_ascii=False), c["birthday"], c["website"], c["notes"],
                    cat_str, json.dumps(c["custom"], ensure_ascii=False), photo_blob
                )
            )
            ol_updated_count += 1

        existing_g = dict(cur.execute("SELECT resource_name, data_hash FROM google_contacts").fetchall())
        current_g_ids = {p["resourceName"] for p in g_raw_list}
        deleted_g_ids = set(existing_g.keys()) - current_g_ids
        if deleted_g_ids:
            cur.executemany("DELETE FROM google_contacts WHERE resource_name=?", [(res,) for res in deleted_g_ids])

        g_updated_count = 0
        total_g = len(g_raw_list)
        for idx, person in enumerate(g_raw_list, start=1):
            if idx == 1 or idx % 200 == 0 or idx == total_g:
                pct_now = 26 + int(4 * (idx / max(1, total_g)))
                self.set_step_progress(
                    pct_now, 2, 5,
                    f"Archive locale ({mode_label})", f"Local archive ({mode_label})",
                    f"Archive Google : {idx}/{total_g} ({int(idx*100/max(1,total_g))}%)",
                    f"Google archive: {idx}/{total_g} ({int(idx*100/max(1,total_g))}%)",
                    idx / max(1, total_g)
                )
            res_name = person["resourceName"]
            c = google_to_canonical(person, id_to_label, use_file_as=True)
            photo_urls = [p.get("url", "") for p in person.get("photos", []) if not p.get("default", False)]
            row_hash = compute_hash({**c, "_photo_urls": photo_urls})

            if not force_full_refresh and existing_g.get(res_name) == row_hash:
                continue

            if idx % 100 == 0 or idx == total_g:
                pct_g = int(97 + 2 * (idx / max(1, total_g)))
                self.after(0, lambda p=pct_g, i=idx, t=total_g, m=mode_label: (
                    self.lbl_progress.configure(text=f"{p}% — Étape 5/5 : Archive locale Google ({m})"),
                    self.progress_bar.set(p / 100.0),
                    self.lbl_sub_progress.configure(text=f"Archive Google ({m}) : {i}/{t} ({int(i*100/max(1,t))}%)"),
                    self.sub_progress_bar.set(i / max(1, t))
                ))

            photo_blob = get_google_photo_bytes(person, service) if (cfg["sync_photos"] and photo_urls) else None
            cat_str = ", ".join(c["categories"])
            cur.execute(
                "INSERT OR REPLACE INTO google_contacts VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                (
                    res_name, row_hash, c["file_as"], c["first_name"], c["middle_name"], c["last_name"],
                    c["prefix"], c["suffix"], c["company"], c["job_title"], c["department"],
                    json.dumps(c["emails"], ensure_ascii=False), json.dumps(c["phones"], ensure_ascii=False),
                    json.dumps(c["addresses"], ensure_ascii=False), c["birthday"], c["website"], c["notes"],
                    cat_str, json.dumps(c["custom"], ensure_ascii=False), photo_blob
                )
            )
            g_updated_count += 1

        cur.execute("DELETE FROM sync_links")
        links = []
        if os.path.exists(DB_FILE):
            state_conn = sqlite3.connect(DB_FILE)
            links = state_conn.execute("SELECT outlook_entry_id, google_resource_name, last_synced FROM contact_map").fetchall()
            cur.executemany("INSERT INTO sync_links VALUES (?, ?, ?)", links)
            state_conn.close()

        now_iso = datetime.now().isoformat()
        cur.execute("DELETE FROM metadata")
        cur.execute(
            "INSERT INTO metadata VALUES (?, ?, ?, ?, ?, ?)",
            (now_iso, mode_label, cfg["outlook_email"], cfg["google_email"], total_ol, total_g)
        )
        exp_conn.commit()

        wb = openpyxl.Workbook(write_only=True)
        ws_meta = wb.create_sheet("Résumé")
        ws_ol = wb.create_sheet("Contacts_Outlook")
        ws_g = wb.create_sheet("Contacts_Google")
        ws_links = wb.create_sheet("Liens_Synchro")

        ws_meta.append(["Paramètre", "Valeur"])
        ws_meta.append(["Dernière mise à jour", now_iso])
        ws_meta.append(["Mode d'extraction", mode_label])
        ws_meta.append(["Compte Outlook", cfg["outlook_email"]])
        ws_meta.append(["Compte Google", cfg["google_email"]])
        ws_meta.append(["Total Contacts Outlook", total_ol])
        ws_meta.append(["Total Contacts Google", total_g])
        ws_meta.append(["Contacts liés", len(links)])

        ws_ol.append([
            "FileAs", "Prénom", "Deuxième prénom", "Nom", "Civilité", "Suffixe",
            "Société", "Poste", "Service", "E-mails", "Téléphones", "Adresses",
            "Anniversaire", "Site Web", "Catégories", "Champs Perso", "Notes", "Photo", "EntryID"
        ])
        ws_g.append([
            "Nom affiché", "Prénom", "Deuxième prénom", "Nom", "Civilité", "Suffixe",
            "Société", "Poste", "Service", "E-mails", "Téléphones", "Adresses",
            "Anniversaire", "Site Web", "Libellés", "Champs Perso", "Notes", "Photo", "ResourceName"
        ])
        ws_links.append(["Outlook EntryID", "Google ResourceName", "Dernière Synchro"])

        def fmt_items(raw_json):
            return " | ".join(f"{x.get('value', '')} ({x.get('type', '')})" for x in json.loads(raw_json or "[]"))

        def fmt_addrs(raw_json):
            out = []
            for a in json.loads(raw_json or "[]"):
                full = ", ".join(filter(None, [a.get("street"), a.get("postal"), a.get("city"), a.get("state"), a.get("country")]))
                if full:
                    out.append(f"[{a.get('type', '')}] {full}")
            return " | ".join(out)

        for row in cur.execute("""
            SELECT file_as, first_name, middle_name, last_name, prefix, suffix, company, job_title, department,
                   emails_json, phones_json, addresses_json, birthday, website, categories, custom_fields_json,
                   notes, (photo_blob IS NOT NULL), entry_id
            FROM outlook_contacts ORDER BY file_as, last_name
        """):
            custom_str = ", ".join(f"{k}={v}" for k, v in json.loads(row[15] or "{}").items())
            ws_ol.append([
                row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7], row[8],
                fmt_items(row[9]), fmt_items(row[10]), fmt_addrs(row[11]),
                row[12], row[13], row[14], custom_str, row[16], "Oui" if row[17] else "Non", row[18]
            ])

        for row in cur.execute("""
            SELECT file_as, first_name, middle_name, last_name, prefix, suffix, company, job_title, department,
                   emails_json, phones_json, addresses_json, birthday, website, labels, custom_fields_json,
                   notes, (photo_blob IS NOT NULL), resource_name
            FROM google_contacts ORDER BY file_as, last_name
        """):
            custom_str = ", ".join(f"{k}={v}" for k, v in json.loads(row[15] or "{}").items())
            ws_g.append([
                row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7], row[8],
                fmt_items(row[9]), fmt_items(row[10]), fmt_addrs(row[11]),
                row[12], row[13], row[14], custom_str, row[16], "Oui" if row[17] else "Non", row[18]
            ])

        for link_row in links:
            ws_links.append(list(link_row))

        exp_conn.close()
        try:
            shutil.move(tmp_db_path, final_db_path)
            actual_xlsx_name = os.path.basename(export_xlsx_path)
            try:
                wb.save(export_xlsx_path)
            except PermissionError:
                ts_suffix = datetime.now().strftime("%Y%m%d_%H%M%S")
                fallback_xlsx_path = os.path.join(archive_dir, f"contacts_{ts_suffix}.xlsx")
                wb.save(fallback_xlsx_path)
                actual_xlsx_name = os.path.basename(fallback_xlsx_path)
                self.log("Warning", f"Permission denied sur contacts.xlsx (fichier ouvert dans Excel). Sauvegarde effectuée sous : {actual_xlsx_name}")
        except OSError as disk_err:
            if os.path.exists(tmp_db_path):
                try:
                    os.remove(tmp_db_path)
                except Exception:
                    pass
            renamed = mark_archive_not_up_to_date(archive_dir)
            self.log(
                "Warning",
                f"Échec d'écriture de l'archive locale ({disk_err}). Fichiers marqués NOT_UP_TO_DATE : {', '.join(renamed)}."
            )
            return

        today_str = datetime.now().strftime("%Y-%m-%d")
        self.config["last_archive_date"] = today_str
        if force_full_refresh:
            self.config["last_full_archive_date"] = today_str
            self._last_ok_str = __import__("datetime").datetime.now().strftime("%Y/%m/%d %H:%M")
        self._last_ok_str = __import__("datetime").datetime.now().strftime("%Y/%m/%d %H:%M")
        self.after(0, self.save_config)

        self.log(
            "Information",
            f"Local archive ({mode_label}) complete: Outlook updated={ol_updated_count}, Google updated={g_updated_count} -> contacts.db & {actual_xlsx_name}"
        )

    # --------------------------------------------------------------------------
    # MOTEUR PRINCIPAL DE SYNCHRONISATION
    # --------------------------------------------------------------------------
    def start_sync_thread(self):
        if self.is_syncing:
            self.log("Error", "The sync thread is already running and alive, waiting for sync thread to finish")
            return
        self.save_config()
        if getattr(self, "var_extract_db", None) and self.var_extract_db.get():
            self.ensure_archive_dir_ready(interactive=True)
        self.is_syncing = True
        self.after(0, lambda: self.btn_sync.configure(state="disabled"))
        threading.Thread(target=self.run_sync_worker, daemon=True).start()

    def run_sync_worker(self):
        pythoncom.CoInitialize()
        stats = {"synced": 0, "deleted": 0, "skipped": 0, "errors": 0, "total": 0}
        cfg = self.config
        mode = cfg["sync_mode"]
        dry_run = cfg["dry_run"]
        today_dt = datetime.now()
        today_str = today_dt.strftime("%Y-%m-%d")
        last_full_sync_str = cfg.get("last_full_sync_date") or cfg.get("last_full_archive_date") or ""
        auto_full_7d = True
        if last_full_sync_str:
            try:
                days_since_full = (today_dt - datetime.strptime(last_full_sync_str, "%Y-%m-%d")).days
                auto_full_7d = (days_since_full > 7)
            except ValueError:
                auto_full_7d = True
        force_full = bool(cfg["force_full_compare"] or auto_full_7d)
        if auto_full_7d and not cfg["force_full_compare"]:
            self.log("Information", f"Dernière Full Sync > 7 jours ({last_full_sync_str or 'jamais'}) -> bascule automatique en mode FULL SYNCHRO (Outlook + Google + Archive .db/.xlsx).")
        elif cfg["force_full_compare"]:
            self.log("Information", "Option 'Force Full Compare' activée -> mode FULL SYNCHRO (Outlook + Google + Archive .db/.xlsx).")
        else:
            self.log("Information", f"Mode incrémental actif (dernière Full Sync : {last_full_sync_str}).")
        connector = cfg.get("outlook_connector", "COM")

        try:
            if not cfg["outlook_email"]:
                raise RuntimeError("Veuillez renseigner votre adresse e-mail Outlook dans 'Comptes & Profils'.")

            self.log("Information", f"Sync started ({cfg['outlook_email']} [{connector}] <-> {cfg['google_email'] or 'Google'} | Mode: {mode} | DryRun: {dry_run})")
            if dry_run:
                self.log("Warning", "DRY-RUN mode enabled: no changes will be written to Outlook or Google.")

            if connector != "COM":
                raise RuntimeError("Connecteur Microsoft 365 Graph API sélectionné mais non configuré. Aucune bascule automatique effectuée.")

            self.log("Information", "Connecting to Google...")
            _c_path, _c_src = resolve_credentials_path()
            if _c_src != "data" or not _c_path:
                if hasattr(self, "_update_oauth_main_button"):
                    self.after(0, lambda: self._update_oauth_main_button(is_ok=False))
                raise RuntimeError(r"Fichier .\data\credentials.json absent ou invalide.")
            service = None
            for _g_try in range(4):
                try:
                    _g_svc, _g_creds, _ = run_google_oauth_connect(_c_path, log_callback=getattr(self, "_write_ui_log", None))
                    service = _g_svc
                    self.google_service = _g_svc
                    self.google_creds = _g_creds
                    if hasattr(self, "_update_oauth_main_button"):
                        self.after(0, lambda: self._update_oauth_main_button(is_ok=True))
                    break
                except Exception as _g_err:
                    if _g_try < 3 and any(k in str(_g_err) for k in ("10054", "ConnectionReset", "ConnectionAborted", "timed out", "EOF")):
                        self.log("Warning", f"Connexion Google interrompue ({_g_err}), nouvelle tentative ({_g_try + 2}/4)...")
                        time.sleep(1.5 * (_g_try + 1))
                    else:
                        raise

            self.log("Information", "Connecting to Outlook Classic (COM)...")
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            ol_folder = None
            for store in namespace.Stores:
                if cfg["outlook_email"].lower() in store.DisplayName.lower():
                    ol_folder = store.GetDefaultFolder(10)
                    break
            if not ol_folder:
                raise RuntimeError(f"Compte Outlook '{cfg['outlook_email']}' introuvable dans COM.")

            self.set_step_progress(10, 1, 5, "Chargement des contacts Outlook", "Loading Outlook contacts", "Lecture du dossier Contacts Outlook...", "Reading Outlook Contacts folder...", 0.4)
            self.log("Information", "Loading Outlook contacts...")
            ol_items = {item.EntryID: item for item in ol_folder.Items if item.Class == 40}
            stats["total"] = len(ol_items)
            self.log("Information", f"Loaded {len(ol_items)} Outlook contacts.")

            self.log("Information", "Loading Google Groups...")
            groups = execute_google_call_with_backoff(service.contactGroups().list(pageSize=200)).get("contactGroups", [])
            label_name_to_id = {g["name"]: g["resourceName"] for g in groups if g.get("groupType") == "USER_CONTACT_GROUP"}
            id_to_label_name = {v: k for k, v in label_name_to_id.items()}

            self.set_step_progress(20, 2, 5, "Chargement des contacts Google", "Loading Google contacts", "Téléchargement des fiches Google...", "Downloading Google contacts...", 0.5)
            self.log("Information", "Loading Google Contacts...")
            g_raw_list = []
            page_token = None
            fields = "names,fileAses,emailAddresses,phoneNumbers,organizations,addresses,biographies,birthdays,urls,userDefined,memberships,photos,metadata"
            while True:
                resp = execute_google_call_with_backoff(
                    service.people().connections().list(
                        resourceName="people/me", pageSize=1000, pageToken=page_token, personFields=fields
                    )
                )
                g_raw_list.extend(resp.get("connections", []))
                page_token = resp.get("nextPageToken")
                if not page_token:
                    break
            g_items = {p["resourceName"]: p for p in g_raw_list}
            self.log("Information", f"Loaded {len(g_items)} Google contacts.")

            conn = init_db()
            cursor = conn.cursor()
            cursor.execute("CREATE TABLE IF NOT EXISTS ol_fast_cache (outlook_entry_id TEXT PRIMARY KEY, last_mod_ts TEXT, canon_json TEXT)")
            conn.commit()
            cursor.execute("SELECT * FROM contact_map")
            db_rows = {
                row[0]: {
                    "res_name": row[1], "ol_hash": row[2], "g_hash": row[3],
                    "ol_photo": row[4], "g_photo": row[5]
                } for row in cursor.fetchall()
            }
            mapped_res_names = {v["res_name"]: k for k, v in db_rows.items()}

            if len(db_rows) > 10 and (len(ol_items) == 0 or len(g_items) == 0):
                empty_side = "Outlook" if len(ol_items) == 0 else "Google"
                raise RuntimeError(self.t("err_zero_contacts").format(side=empty_side, linked=len(db_rows)))

            # Archivage local (.db avec photos + .xlsx) : aligné sur le mode global (Incrémental par défaut, Full si coché ou > 7 jours)


            for item in ol_items.values():
                for cat in (item.Categories or "").split(","):
                    c = cat.strip()
                    if c and c not in label_name_to_id and not dry_run:
                        created_grp = execute_google_call_with_backoff(
                            service.contactGroups().create(body={"contactGroup": {"name": c}})
                        )
                        label_name_to_id[c] = created_grp["resourceName"]
                        id_to_label_name[created_grp["resourceName"]] = c

            ol_fast_rows = {}
            try:
                cursor.execute("SELECT outlook_entry_id, last_mod_ts, canon_json FROM ol_fast_cache")
                ol_fast_rows = {r[0]: (r[1], r[2]) for r in cursor.fetchall()}
            except Exception:
                ol_fast_rows = {}

            ol_data = {}
            ol_cache_hits = 0
            ol_cache_updates = []
            tot_ol = max(1, len(ol_items))
            self.set_step_progress(35, 3, 5, "Préparation des contacts Outlook", "Preparing Outlook contacts", "Analyse des fiches Outlook...", "Analyzing Outlook contacts...", 0.0)
            for idx_o, (eid, it) in enumerate(ol_items.items(), 1):
                try:
                    mod_ts = str(getattr(it, "LastModificationTime", "") or "")
                except Exception:
                    mod_ts = ""
                cached_entry = ol_fast_rows.get(eid)
                if not force_full and mod_ts and cached_entry and cached_entry[0] == mod_ts and cached_entry[1]:
                    try:
                        canon_obj = json.loads(cached_entry[1])
                        ol_data[eid] = {"item": it, "canon": canon_obj}
                        ol_cache_hits += 1
                    except Exception:
                        canon_obj = outlook_to_canonical(it, cfg["use_file_as"])
                        ol_data[eid] = {"item": it, "canon": canon_obj}
                        if mod_ts:
                            ol_cache_updates.append((eid, mod_ts, json.dumps(canon_obj, ensure_ascii=False)))
                else:
                    canon_obj = outlook_to_canonical(it, cfg["use_file_as"])
                    ol_data[eid] = {"item": it, "canon": canon_obj}
                    if mod_ts:
                        ol_cache_updates.append((eid, mod_ts, json.dumps(canon_obj, ensure_ascii=False)))

                if idx_o == 1 or idx_o % 100 == 0 or idx_o == tot_ol:
                    self.after(0, lambda i=idx_o, t=tot_ol: (
                        self.lbl_progress.configure(text=f"{int(20 + 35*i/t)}% — Étape 3/5 : Préparation Outlook"), self.update_tray_tooltip(f"{int(20 + 35*i/t)}% — Étape 3/5 : Préparation Outlook"),
                        self.progress_bar.set(0.20 + 0.35 * (i / t)),
                        self.lbl_sub_progress.configure(text=("Sub-step: " if self.config.get("language", "EN") == "EN" else "Sous-étape : ") + f"Lecture COM & empreintes Outlook : {i}/{t} ({int(i*100/t)}%)"),
                        self.sub_progress_bar.set(i / t)
                    ))
                if idx_o % 1000 == 0:
                    self.log("Information", f"  [Outlook COM] Normalisation fiches/photos : {idx_o}/{tot_ol} (cache delta: {ol_cache_hits})...")

            if ol_cache_updates and not dry_run:
                try:
                    cursor.executemany("INSERT OR REPLACE INTO ol_fast_cache VALUES (?, ?, ?)", ol_cache_updates)
                    conn.commit()
                except Exception:
                    pass
            if not force_full:
                self.log("Information", f"  [Outlook Delta] {ol_cache_hits}/{len(ol_items)} fiches inchangées lues depuis le cache rapide ({len(ol_cache_updates)} modifiée(s)).")

            g_data = {}
            tot_g = max(1, len(g_items))
            self.set_step_progress(55, 4, 5, "Préparation des contacts Google", "Preparing Google contacts", "Normalisation des fiches Google...", "Normalizing Google contacts...", 0.0)
            for idx_g, (res, p) in enumerate(g_items.items(), 1):
                g_data[res] = {"person": p, "canon": google_to_canonical(p, id_to_label_name, cfg["use_file_as"])}
                if idx_g == 1 or idx_g % 50 == 0 or idx_g == tot_g:
                    self.after(0, lambda i=idx_g, t=tot_g: (
                        self.lbl_progress.configure(text=f"Étape 4/5 — Préparation Google ({int(55 + 25*i/t)}%)"), self.update_tray_tooltip(f"Étape 4/5 — Préparation Google ({int(55 + 25*i/t)}%)"),
                        self.progress_bar.set(0.55 + 0.25 * (i / t)),
                        self.lbl_sub_progress.configure(text=("Sub-step: " if self.config.get("language", "EN") == "EN" else "Sous-étape : ") + f"Normalisation fiches Google : {i}/{t} ({int(i*100/t)}%)"),
                        self.sub_progress_bar.set(i / t)
                    ))

            if cfg.get("force_file_as_format"):
                fa_style = cfg.get("file_as_format_style", "Name SURNAME")
                ol_is_source = mode in ("OUTLOOK_TO_GOOGLE_ONLY", "MERGE_OUTLOOK_WINS", "MERGE_GOOGLE_WINS", "MERGE_PROMPT")
                g_is_source = mode in ("GOOGLE_TO_OUTLOOK_ONLY", "MERGE_OUTLOOK_WINS", "MERGE_GOOGLE_WINS", "MERGE_PROMPT")
                fa_fixed_ol = 0
                fa_fixed_g = 0
                if ol_is_source:
                    for _eid, _od in ol_data.items():
                        _oc = _od["canon"]
                        _tgt = format_target_file_as(_oc.get("first_name"), _oc.get("last_name"), fa_style)
                        if _tgt and (_oc.get("file_as") or "") != _tgt:
                            _oc["file_as"] = _tgt
                            if not dry_run:
                                try:
                                    _od["item"].FileAs = _tgt
                                    _od["item"].Save()
                                    _mts = str(getattr(_od["item"], "LastModificationTime", ""))
                                    if _mts:
                                        cursor.execute(
                                            "INSERT OR REPLACE INTO ol_fast_cache VALUES (?, ?, ?)",
                                            (_eid, _mts, json.dumps(_oc, ensure_ascii=False))
                                        )
                                    fa_fixed_ol += 1
                                except Exception:
                                    pass
                if g_is_source:
                    for _rn, _gd in g_data.items():
                        _gc = _gd["canon"]
                        _tgt = format_target_file_as(_gc.get("first_name"), _gc.get("last_name"), fa_style)
                        if _tgt and (_gc.get("file_as") or "") != _tgt:
                            _gc["file_as"] = _tgt
                            if not dry_run:
                                try:
                                    _b = canonical_to_google_body(_gc, label_name_to_id, True)
                                    _b["etag"] = _gd["person"]["etag"]
                                    _upd = execute_google_call_with_backoff(
                                        service.people().updateContact(resourceName=_rn, updatePersonFields="names,fileAses", body=_b)
                                    )
                                    if _upd:
                                        _gd["person"] = _upd
                                    fa_fixed_g += 1
                                except Exception:
                                    pass
                if fa_fixed_ol or fa_fixed_g:
                    if not dry_run:
                        conn.commit()
                    self.log("Information", f"Force FileAs ({fa_style}): {fa_fixed_ol} Outlook + {fa_fixed_g} Google contact(s) normalized.")

            for eid, state in list(db_rows.items()):
                res_name = state["res_name"]
                if mode == "OUTLOOK_TO_GOOGLE_ONLY" and eid in ol_data and res_name not in g_data:
                    self.log("Warning", f"Google contact missing for Outlook ID {eid[-8:]} -> Auto-healing link to recreate in Google.")
                    if not dry_run:
                        cursor.execute("DELETE FROM contact_map WHERE outlook_entry_id=?", (eid,))
                        conn.commit()
                    del db_rows[eid]
                    mapped_res_names.pop(res_name, None)

            self.set_step_progress(82, 5, 5, "Rapprochement Outlook <-> Google", "Matching Outlook <-> Google", "Comparaison des fiches...", "Comparing records...", 0.3)
            self.set_step_progress(85, 5, 5, "Rapprochement & synchronisation", "Matching & synchronization", "Application des mises à jour...", "Applying updates...", 0.5)
            self.log("Information", "Matching Outlook and Google contacts...")
            unmapped_ol = [eid for eid in ol_data if eid not in db_rows]
            unmapped_g = {res: g_data[res] for res in g_data if res not in mapped_res_names}

            total_ol = len(ol_data)
            for idx, eid in enumerate(ol_data, start=1):
                c = ol_data[eid]["canon"]
                disp = c["file_as"] or f"{c['first_name']} {c['last_name']}".strip() or eid[-8:]
                if idx % 25 == 0 or idx == total_ol:
                    self.update_status_line(f"Matching contact {idx} of {total_ol} by id: {disp}", idx, total_ol)

                if eid in unmapped_ol:
                    import unicodedata
                    def _norm_id(fn, ln):
                        raw = f"{(fn or '').strip()} {(ln or '').strip()}".strip().lower()
                        raw = "".join(ch for ch in unicodedata.normalize("NFD", raw) if unicodedata.category(ch) != "Mn")
                        return " ".join(raw.split())

                    def _g_completeness(gc_obj):
                        return (
                            len(gc_obj.get("emails") or []) * 3
                            + len(gc_obj.get("phones") or []) * 2
                            + (2 if gc_obj.get("company") else 0)
                            + (2 if gc_obj.get("job_title") else 0)
                            + (1 if gc_obj.get("notes") else 0)
                        )

                    ol_emails = {e["value"].strip().lower() for e in c["emails"] if e.get("value")}
                    ol_phones = {
                        normalize_phone(ph["value"])
                        for ph in c["phones"]
                        if normalize_phone(ph.get("value")) and "assistant" not in str(ph.get("type", "")).lower()
                    }
                    ol_id = _norm_id(c.get("first_name"), c.get("last_name"))
                    ol_id_rev = _norm_id(c.get("last_name"), c.get("first_name"))

                    tier1_email_and_name = []
                    tier2_unique_email = []
                    tier3_exact_name = []
                    tier4_phone_and_name = []

                    for res_name, gd in unmapped_g.items():
                        gc = gd["canon"]
                        g_emails = {e["value"].strip().lower() for e in gc["emails"] if e.get("value")}
                        g_phones = {
                            normalize_phone(ph["value"])
                            for ph in gc["phones"]
                            if normalize_phone(ph.get("value")) and "assistant" not in str(ph.get("type", "")).lower()
                        }
                        g_id = _norm_id(gc.get("first_name"), gc.get("last_name"))

                        same_name = bool(ol_id and len(ol_id) > 2 and (ol_id == g_id or ol_id_rev == g_id))
                        shared_emails = ol_emails & g_emails
                        shared_phones = ol_phones & g_phones

                        if shared_emails:
                            if same_name or not ol_id or not g_id:
                                tier1_email_and_name.append((_g_completeness(gc), res_name))
                            else:
                                tier2_unique_email.append((_g_completeness(gc), res_name))
                        elif same_name:
                            if not (ol_emails and g_emails and not (ol_emails & g_emails)):
                                tier3_exact_name.append((_g_completeness(gc), res_name))
                        elif shared_phones and same_name:
                            if not ((not ol_emails and g_emails) or (ol_emails and not g_emails)):
                                tier4_phone_and_name.append((_g_completeness(gc), res_name))

                    matched_res = None
                    if tier1_email_and_name:
                        tier1_email_and_name.sort(reverse=True)
                        matched_res = tier1_email_and_name[0][1]
                    elif len(tier2_unique_email) == 1:
                        matched_res = tier2_unique_email[0][1]
                    elif tier3_exact_name:
                        tier3_exact_name.sort(reverse=True)
                        matched_res = tier3_exact_name[0][1]
                    elif tier4_phone_and_name:
                        tier4_phone_and_name.sort(reverse=True)
                        matched_res = tier4_phone_and_name[0][1]

                    if matched_res:
                        self.log("Information", f"Matched existing contact: \"{disp}\"")
                        ol_h = compute_hash(c)
                        g_h = compute_hash(g_data[matched_res]["canon"])
                        base_h = ol_h if ol_h == g_h else "INITIAL_DIFF"
                        db_rows[eid] = {
                            "res_name": matched_res, "ol_hash": base_h, "g_hash": base_h,
                            "ol_photo": None, "g_photo": None
                        }
                        mapped_res_names[matched_res] = eid
                        del unmapped_g[matched_res]
                        if not dry_run:
                            cursor.execute(
                                "INSERT INTO contact_map VALUES (?, ?, ?, ?, ?, ?, ?)",
                                (eid, matched_res, base_h, base_h, None, None, datetime.now().isoformat())
                            )
                            conn.commit()

            # ==================================================================
            # PHASE DE PRE-ANALYSE : CALCUL EXACT DU PLAN D'ACTION AVANT ECRITURE
            # ==================================================================
            def get_diff_fields(c1, c2):
                return [k for k in c1 if c1[k] != c2[k]]

            linked_items = [(eid, st) for eid, st in db_rows.items() if eid in ol_data and st["res_name"] in g_data]
            plan_identical = 0
            plan_ol_updates = []
            plan_g_updates = []
            sample_diffs_logged = 0

            for eid, state in linked_items:
                res_name = state["res_name"]
                ol_canon = ol_data[eid]["canon"]
                g_canon = g_data[res_name]["canon"]
                disp = ol_canon["file_as"] or f"{ol_canon['first_name']} {ol_canon['last_name']}".strip() or eid[-8:]

                cur_ol_hash = compute_hash(ol_canon)
                cur_g_hash = compute_hash(g_canon)

                if force_full:
                    is_different = (cur_ol_hash != cur_g_hash)
                    ol_changed = is_different
                    g_changed = is_different
                else:
                    ol_changed = (cur_ol_hash != state["ol_hash"])
                    g_changed = (cur_g_hash != state["g_hash"])
                    if cfg.get("use_file_as") and (ol_canon.get("file_as") or "") != (g_canon.get("file_as") or ""):
                        if mode != "GOOGLE_TO_OUTLOOK_ONLY":
                            ol_changed = True
                        else:
                            g_changed = True

                photo_needs_sync = False
                if cfg["sync_photos"]:
                    has_ol_pic = bool(getattr(ol_data[eid]["item"], "HasPicture", False))
                    has_g_pic = any(not p.get("default", False) for p in g_data[res_name]["person"].get("photos", []))
                    if mode != "GOOGLE_TO_OUTLOOK_ONLY" and (has_ol_pic != bool(state["ol_photo"])):
                        photo_needs_sync = True
                    elif mode != "OUTLOOK_TO_GOOGLE_ONLY" and (has_g_pic != bool(state["g_photo"])):
                        photo_needs_sync = True

                planned_winner = None
                if ol_changed and g_changed:
                    if mode in ("MERGE_OUTLOOK_WINS", "OUTLOOK_TO_GOOGLE_ONLY"):
                        planned_winner = "OUTLOOK"
                    elif mode in ("MERGE_GOOGLE_WINS", "GOOGLE_TO_OUTLOOK_ONLY"):
                        planned_winner = "GOOGLE"
                    elif mode == "MERGE_PROMPT":
                        planned_winner = "PROMPT"
                elif ol_changed and mode != "GOOGLE_TO_OUTLOOK_ONLY":
                    planned_winner = "OUTLOOK"
                elif g_changed and mode != "OUTLOOK_TO_GOOGLE_ONLY":
                    planned_winner = "GOOGLE"

                if planned_winner is None and not photo_needs_sync:
                    plan_identical += 1
                else:
                    if planned_winner in ("OUTLOOK", "PROMPT") or (planned_winner is None and photo_needs_sync and mode != "GOOGLE_TO_OUTLOOK_ONLY"):
                        plan_g_updates.append((eid, state, planned_winner))
                        if sample_diffs_logged < 5 and ol_changed:
                            diff_keys = get_diff_fields(ol_canon, g_canon)
                            self.log("Information", f"  [Exemple écart] \"{disp}\" -> champs différents : {', '.join(diff_keys) or 'photo'}")
                            sample_diffs_logged += 1
                    elif planned_winner == "GOOGLE" or (planned_winner is None and photo_needs_sync and mode == "GOOGLE_TO_OUTLOOK_ONLY"):
                        plan_ol_updates.append((eid, state, planned_winner))

            plan_ol_creates = [res for res in g_data if res not in mapped_res_names] if mode != "OUTLOOK_TO_GOOGLE_ONLY" else []
            plan_g_creates = [eid for eid in ol_data if eid not in db_rows] if mode != "GOOGLE_TO_OUTLOOK_ONLY" else []

            self.log("Information", "------------------------------------------------------------")
            self.log("Information", "BILAN PREVISIONNEL AVANT MISE A JOUR :")
            self.log("Information", f"  • Contacts en commun avec les mêmes données : {plan_identical}")
            self.log("Information", f"  • Contacts à créer côté Outlook             : {len(plan_ol_creates)}")
            self.log("Information", f"  • Contacts à mettre à jour côté Outlook     : {len(plan_ol_updates)}")
            self.log("Information", f"  • Contacts à créer côté Google              : {len(plan_g_creates)}")
            self.log("Information", f"  • Contacts à mettre à jour côté Google      : {len(plan_g_updates)}")
            self.log("Information", "------------------------------------------------------------")

            # Synchronisation des contacts liés (uniquement ceux nécessitant une mise à jour)
            items_to_process = plan_g_updates + plan_ol_updates
            for idx, (eid, state, _pre_winner) in enumerate(items_to_process, start=1):
                res_name = state["res_name"]
                ol_canon = ol_data[eid]["canon"]
                g_canon = g_data[res_name]["canon"]
                disp = ol_canon["file_as"] or f"{ol_canon['first_name']} {ol_canon['last_name']}".strip() or eid[-8:]

                if idx % 25 == 0 or idx == len(linked_items):
                    self.update_status_line(f"Syncing contact {idx} of {len(items_to_process)}: {disp}", idx, len(items_to_process))

                cur_ol_hash = compute_hash(ol_canon)
                cur_g_hash = compute_hash(g_canon)

                if force_full:
                    is_different = (cur_ol_hash != cur_g_hash)
                    ol_changed = is_different
                    g_changed = is_different
                else:
                    ol_changed = (cur_ol_hash != state["ol_hash"])
                    g_changed = (cur_g_hash != state["g_hash"])

                winner = None
                if ol_changed and g_changed:
                    if mode in ("MERGE_OUTLOOK_WINS", "OUTLOOK_TO_GOOGLE_ONLY"):
                        winner = "OUTLOOK"
                    elif mode in ("MERGE_GOOGLE_WINS", "GOOGLE_TO_OUTLOOK_ONLY"):
                        winner = "GOOGLE"
                    elif mode == "MERGE_PROMPT":
                        winner = self.ask_conflict_choice_threadsafe(disp, ol_canon, g_canon)
                        if winner == "SKIP":
                            stats["skipped"] += 1
                elif ol_changed and mode != "GOOGLE_TO_OUTLOOK_ONLY":
                    winner = "OUTLOOK"
                elif g_changed and mode != "OUTLOOK_TO_GOOGLE_ONLY":
                    winner = "GOOGLE"

                try:
                    if winner == "OUTLOOK":
                        self.log("Information", f"Updated Google contact from Outlook: \"{disp}\"")
                        if not dry_run:
                            body = canonical_to_google_body(ol_canon, label_name_to_id, cfg["use_file_as"])
                            body["etag"] = g_data[res_name]["person"]["etag"]
                            u_fields = "names,fileAses,emailAddresses,phoneNumbers,organizations,addresses,birthdays,urls,biographies,userDefined,memberships"
                            updated_person = execute_google_call_with_backoff(
                                service.people().updateContact(resourceName=res_name, updatePersonFields=u_fields, body=body)
                            )
                            if updated_person:
                                g_data[res_name]["person"] = updated_person
                                g_items[res_name] = updated_person
                            cur_g_hash = cur_ol_hash
                            if force_full:
                                time.sleep(0.5)
                        stats["synced"] += 1

                    elif winner == "GOOGLE":
                        allow_notes = True
                        if (cfg.get("mode") == "OUTLOOK_TO_GOOGLE_ONLY" or not cfg["force_rtf"]) and is_outlook_body_rich_rtf(ol_data[eid]["item"]):
                            allow_notes = False
                            self.log("Warning", f"Outlook contact notes body not updated, because it is RTF: {disp}")

                        self.log("Information", f"Updated Outlook contact from Google: \"{disp}\"")
                        if not dry_run:
                            apply_canonical_to_outlook(ol_data[eid]["item"], g_canon, update_notes=allow_notes, dry_run=dry_run)
                            cur_ol_hash = compute_hash(outlook_to_canonical(ol_data[eid]["item"], cfg["use_file_as"]))
                        stats["synced"] += 1

                    ol_p_hash, g_p_hash = state["ol_photo"], state["g_photo"]
                    if cfg["sync_photos"]:
                        ol_p_bytes = get_outlook_photo_bytes(ol_data[eid]["item"])
                        cur_ol_p_hash = hashlib.sha256(ol_p_bytes).hexdigest() if ol_p_bytes else None

                        if cur_ol_p_hash != ol_p_hash and mode != "GOOGLE_TO_OUTLOOK_ONLY":
                            self.log("Information", f"Updated Google photo from Outlook: \"{disp}\"")
                            set_google_photo_bytes(service, res_name, ol_p_bytes, dry_run=dry_run)
                            ol_p_hash = g_p_hash = cur_ol_p_hash
                        elif mode != "OUTLOOK_TO_GOOGLE_ONLY":
                            g_p_bytes = get_google_photo_bytes(g_data[res_name]["person"], service)
                            cur_g_p_hash = hashlib.sha256(g_p_bytes).hexdigest() if g_p_bytes else None
                            if cur_g_p_hash != g_p_hash:
                                self.log("Information", f"Updated Outlook photo from Google: \"{disp}\"")
                                set_outlook_photo_bytes(ol_data[eid]["item"], g_p_bytes, dry_run=dry_run)
                                ol_p_hash = g_p_hash = cur_g_p_hash

                    if (winner or ol_p_hash != state["ol_photo"] or g_p_hash != state["g_photo"]) and not dry_run:
                        cursor.execute("""
                            UPDATE contact_map
                            SET outlook_hash=?, google_hash=?, outlook_photo_hash=?, google_photo_hash=?, last_synced=?
                            WHERE outlook_entry_id=?
                        """, (cur_ol_hash, cur_g_hash, ol_p_hash, g_p_hash, datetime.now().isoformat(), eid))
                        conn.commit()
                except Exception as e:
                    stats["errors"] += 1
                    self.log("Error", f"Failed syncing \"{disp}\": {e}")

            if mode != "GOOGLE_TO_OUTLOOK_ONLY":
                for eid, data in ol_data.items():
                    if eid in db_rows:
                        continue
                    canon = data["canon"]
                    disp = canon["file_as"] or f"{canon['first_name']} {canon['last_name']}".strip() or eid[-8:]
                    try:
                        self.log("Information", f"Created Google contact from Outlook: \"{disp}\"")
                        if not dry_run:
                            body = canonical_to_google_body(canon, label_name_to_id, cfg["use_file_as"])
                            created = execute_google_call_with_backoff(service.people().createContact(body=body))
                            res_name = created["resourceName"]
                            p_hash = None
                            if cfg["sync_photos"]:
                                p_bytes = get_outlook_photo_bytes(data["item"])
                                if p_bytes:
                                    set_google_photo_bytes(service, res_name, p_bytes, dry_run=dry_run)
                                    p_hash = hashlib.sha256(p_bytes).hexdigest()
                            h = compute_hash(canon)
                            cursor.execute(
                                "INSERT INTO contact_map VALUES (?, ?, ?, ?, ?, ?, ?)",
                                (eid, res_name, h, h, p_hash, p_hash, datetime.now().isoformat())
                            )
                            conn.commit()
                        stats["synced"] += 1
                    except Exception as e:
                        stats["errors"] += 1
                        self.log("Error", f"Failed creating Google contact \"{disp}\": {e}")

            if mode != "OUTLOOK_TO_GOOGLE_ONLY":
                for res_name, data in g_data.items():
                    if res_name in mapped_res_names:
                        continue
                    canon = data["canon"]
                    disp = canon["file_as"] or f"{canon['first_name']} {canon['last_name']}".strip() or res_name
                    try:
                        self.log("Information", f"Created Outlook contact from Google: \"{disp}\"")
                        if not dry_run:
                            new_ol_item = ol_folder.Items.Add(2)
                            apply_canonical_to_outlook(new_ol_item, canon, update_notes=True, dry_run=dry_run)
                            eid = new_ol_item.EntryID
                            p_hash = None
                            if cfg["sync_photos"]:
                                p_bytes = get_google_photo_bytes(data["person"], service)
                                if p_bytes:
                                    set_outlook_photo_bytes(new_ol_item, p_bytes, dry_run=dry_run)
                                    p_hash = hashlib.sha256(p_bytes).hexdigest()
                            h = compute_hash(canon)
                            cursor.execute(
                                "INSERT INTO contact_map VALUES (?, ?, ?, ?, ?, ?, ?)",
                                (eid, res_name, h, h, p_hash, p_hash, datetime.now().isoformat())
                            )
                            conn.commit()
                        stats["synced"] += 1
                    except Exception as e:
                        stats["errors"] += 1
                        self.log("Error", f"Failed creating Outlook contact \"{disp}\": {e}")

            if cfg["sync_deletions"] and not force_full:
                deletions = []
                for eid, state in db_rows.items():
                    res_name = state["res_name"]
                    if eid not in ol_data and res_name in g_data and mode != "GOOGLE_TO_OUTLOOK_ONLY":
                        gc = g_data[res_name]["canon"]
                        disp = gc["file_as"] or f"{gc['first_name']} {gc['last_name']}".strip() or res_name
                        deletions.append(("DELETE_GOOGLE", eid, res_name, disp))
                    elif res_name not in g_data and eid in ol_data and mode != "OUTLOOK_TO_GOOGLE_ONLY":
                        oc = ol_data[eid]["canon"]
                        disp = oc["file_as"] or f"{oc['first_name']} {oc['last_name']}".strip() or eid[-8:]
                        deletions.append(("DELETE_OUTLOOK", eid, res_name, disp))
                    elif eid not in ol_data and res_name not in g_data:
                        deletions.append(("CLEAN_DB", eid, res_name, eid[-8:]))

                real_deletions = [d for d in deletions if d[0] != "CLEAN_DB"]
                for action, eid, res_name, disp in deletions:
                    if action == "CLEAN_DB" and not dry_run:
                        cursor.execute("DELETE FROM contact_map WHERE outlook_entry_id=?", (eid,))
                        conn.commit()

                if real_deletions:
                    limit = cfg.get("max_deletions_threshold", 10)
                    self.log("Information", f"Planned deletions detected: {len(real_deletions)} (threshold: {limit}).")

                    approved_batch = True
                    if len(real_deletions) > limit:
                        self.log("Warning", f"Deletion threshold exceeded: {len(real_deletions)} planned deletions > {limit}.")
                        if not dry_run:
                            preview_names = "\n".join(f"  • {d[3]} ({'Google' if d[0] == 'DELETE_GOOGLE' else 'Outlook'})" for d in real_deletions[:8])
                            if len(real_deletions) > 8:
                                preview_names += f"\n  ... (+{len(real_deletions) - 8} autres)"

                            approved_batch = self.ask_user_yes_no_threadsafe(
                                self.t("mass_del_title"),
                                self.t("mass_del_msg").format(count=len(real_deletions), limit=limit, preview=preview_names)
                            )

                    if not approved_batch:
                        self.log("Warning", f"All {len(real_deletions)} planned deletions were rejected by user.")
                        stats["skipped"] += len(real_deletions)
                    else:
                        os.makedirs(BACKUP_DIR, exist_ok=True)
                        bkp_path = os.path.join(BACKUP_DIR, f"google_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json")
                        with open(bkp_path, "w", encoding="utf-8") as bf:
                            json.dump(g_raw_list, bf, ensure_ascii=False, indent=2)

                        for action, eid, res_name, disp in real_deletions:
                            target_side = "Google" if action == "DELETE_GOOGLE" else "Outlook"
                            if cfg["prompt_deletion"] and len(real_deletions) <= limit and not dry_run:
                                confirmed = self.ask_user_yes_no_threadsafe(
                                    self.t("confirm_del_title"),
                                    self.t("confirm_del_msg").format(disp=disp, side=target_side)
                                )
                                if not confirmed:
                                    self.log("Information", f"Deletion skipped by user: \"{disp}\"")
                                    stats["skipped"] += 1
                                    continue

                            try:
                                if action == "DELETE_GOOGLE":
                                    self.log("Information", f"Deleted Google contact (archived in local DB): \"{disp}\"")
                                    if not dry_run:
                                        canon_snapshot = g_data[res_name]["canon"]
                                        photo_snapshot = get_google_photo_bytes(g_data[res_name]["person"], service) if cfg["sync_photos"] else None
                                        cursor.execute("""
                                            INSERT INTO deleted_contacts_history
                                            (deleted_at, target_system, display_name, outlook_entry_id, google_resource_name, contact_data_json, photo_blob)
                                            VALUES (?, ?, ?, ?, ?, ?, ?)
                                        """, (
                                            datetime.now().isoformat(), "GOOGLE", disp, eid, res_name,
                                            json.dumps(canon_snapshot, ensure_ascii=False), photo_snapshot
                                        ))
                                        execute_google_call_with_backoff(
                                            service.people().deleteContact(resourceName=res_name)
                                        )

                                elif action == "DELETE_OUTLOOK":
                                    self.log("Information", f"Deleted Outlook contact (archived in local DB): \"{disp}\"")
                                    if not dry_run:
                                        canon_snapshot = ol_data[eid]["canon"]
                                        photo_snapshot = get_outlook_photo_bytes(ol_data[eid]["item"]) if cfg["sync_photos"] else None
                                        cursor.execute("""
                                            INSERT INTO deleted_contacts_history
                                            (deleted_at, target_system, display_name, outlook_entry_id, google_resource_name, contact_data_json, photo_blob)
                                            VALUES (?, ?, ?, ?, ?, ?, ?)
                                        """, (
                                            datetime.now().isoformat(), "OUTLOOK", disp, eid, res_name,
                                            json.dumps(canon_snapshot, ensure_ascii=False), photo_snapshot
                                        ))
                                        ol_data[eid]["item"].Delete()

                                if not dry_run:
                                    cursor.execute("DELETE FROM contact_map WHERE outlook_entry_id=?", (eid,))
                                    conn.commit()
                                stats["deleted"] += 1
                            except Exception as e:
                                stats["errors"] += 1
                                self.log("Error", f"Failed deleting \"{disp}\": {e}")

            if cfg["extract_local_db"]:
                self.set_step_progress(95, 5, 5, "Mise à jour de l'archive locale", "Updating local archive", "Génération .db et .xlsx synchronisés...", "Generating synced .db and .xlsx...", 0.5)
                archive_db_exists = os.path.exists(os.path.join(self.get_archive_dir(), "contacts.db"))
                archive_xlsx_exists = os.path.exists(os.path.join(self.get_archive_dir(), "contacts.xlsx"))
                need_full_archive = bool(force_full or not archive_db_exists or not archive_xlsx_exists)
                updated_g_list = [v["person"] for v in g_data.values()]
                self.perform_local_db_extraction(
                    ol_items, updated_g_list, id_to_label_name, service, cfg, force_full_refresh=need_full_archive
                )

            # Audit automatique d integrite post-synchronisation (100% des liens)
            try:
                import unicodedata as _ud
                def _chk_n(a, b):
                    s = f"{(a or '').strip()} {(b or '').strip()}".strip().lower()
                    return "".join(ch for ch in _ud.normalize("NFD", s) if _ud.category(ch) != "Mn")
                total_checked = 0
                name_mismatches = 0
                for _eid, _st in db_rows.items():
                    _rn = _st.get("res_name")
                    if _eid in ol_data and _rn in g_data:
                        total_checked += 1
                        _oc = ol_data[_eid]["canon"]
                        _gc = g_data[_rn]["canon"]
                        _on = _chk_n(_oc.get("first_name"), _oc.get("last_name"))
                        _gn = _chk_n(_gc.get("first_name"), _gc.get("last_name"))
                        _ofa = _chk_n(_oc.get("file_as"), "")
                        _gfa = _chk_n(_gc.get("file_as"), "")
                        _fa_diff = bool(cfg.get("use_file_as") and _ofa and _gfa and _ofa != _gfa)
                        if (_on and _gn and _on != _gn) or _fa_diff:
                            name_mismatches += 1
                            if name_mismatches <= 10:
                                self.log("Warning", f"  [Audit Écart Nom] OL='{_oc.get('first_name')} {_oc.get('last_name')}' <-> G='{_gc.get('first_name')} {_gc.get('last_name')}'")
                self.log("Information", f"Integrity Audit: {total_checked - name_mismatches}/{total_checked} linked contacts verified OK ({name_mismatches} mismatch(es)).")
            except Exception as _aud_err:
                self.log("Warning", f"Integrity Audit skipped: {_aud_err}")

            conn.close()
            try:
                with sqlite3.connect(DB_FILE) as _c_end: _lk = dict(_c_end.execute('SELECT outlook_entry_id, google_resource_name FROM contact_map').fetchall())
                _ol_only = [{'id': k, 'name': (v['canon']['file_as'] or k), 'email': ','.join(x['value'] for x in v['canon']['emails']), 'phone': ','.join(x['value'] for x in v['canon']['phones']), 'raw': v['canon']} for k, v in ol_data.items() if k not in _lk]
                _g_only  = [{'id': k, 'name': (v['canon']['file_as'] or k), 'email': ','.join(x['value'] for x in v['canon']['emails']), 'phone': ','.join(x['value'] for x in v['canon']['phones']), 'raw': v['canon']} for k, v in g_data.items() if k not in set(_lk.values())]
                self.log('Information', f'RAPPORT FINAL PAR COMPTE : Outlook={len(ol_data)} (absents de Google={len(_ol_only)}) | Google={len(g_data)} (absents d Outlook={len(_g_only)})')
                if _g_only: self.log('Information', '  -> Uniquement Google : ' + ', '.join(x['name'] for x in _g_only))
                if _ol_only: self.log('Information', '  -> Uniquement Outlook : ' + ', '.join(x['name'] for x in _ol_only))
                if not cfg.get('dry_run', False) and (_ol_only or _g_only):
                    if cfg.get('force_mirror', False) and cfg.get('mode') == 'OUTLOOK_TO_GOOGLE_ONLY' and _g_only:
                        for _it in _g_only:
                            if hasattr(self, 'backup_deleted_contact'):
                                self.backup_deleted_contact('GOOGLE', _it['id'], _it['name'], _it['raw'])
                            execute_google_call_with_backoff(service.people().deleteContact(resourceName=_it['id']))
                        self.log('Information', f'Force Mirror : {len(_g_only)} contact(s) exclusif(s) supprimé(s) côté Google.')
                        _g_only = []
                    elif cfg.get('force_mirror', False) and cfg.get('mode') == 'GOOGLE_TO_OUTLOOK_ONLY' and _ol_only:
                        for _it in _ol_only:
                            if hasattr(self, 'backup_deleted_contact'):
                                self.backup_deleted_contact('OUTLOOK', _it['id'], _it['name'], _it['raw'])
                            if _it['id'] in ol_items:
                                ol_items[_it['id']].Delete()
                        self.log('Information', f'Force Mirror : {len(_ol_only)} contact(s) exclusif(s) supprimé(s) côté Outlook.')
                        _ol_only = []
                    if _ol_only or _g_only:
                        if self.state() == 'withdrawn':
                            self.after(200, lambda o=_ol_only, g=_g_only: self.notify_exclusives_tray(o, g, None, service))
                        else:
                            self.after(300, lambda o=_ol_only, g=_g_only: self.show_exclusives_window(o, g, None, service))
            except Exception as _ex: self.log('Error', f'Erreur rapport exclusifs: {_ex}')
            self.after(0, lambda: (self.sub_progress_bar.set(1.0), self.lbl_sub_progress.configure(text= ("Sub-step: " if self.config.get("language", "EN") == "EN" else "Sous-étape : ") + "Terminé (100%)")))
            _ol_ex = locals().get("_ol_only", []) or []
            _g_ex = locals().get("_g_only", []) or []
            self._last_excl_ol_count = len(_ol_ex)
            self._last_excl_g_count = len(_g_ex)
            _excl_tot = self._last_excl_ol_count + self._last_excl_g_count
            self.config["last_excl_ol_count"] = self._last_excl_ol_count
            self.config["last_excl_g_count"] = self._last_excl_g_count
            self.config["last_excl_ol_items"] = [{"id": x.get("id") or x.get("entry_id", ""), "name": x.get("name", ""), "email": x.get("email", ""), "phone": x.get("phone", "")} for x in _ol_ex if isinstance(x, dict)]
            self.config["last_excl_g_items"] = [{"id": x.get("id") or x.get("resourceName", ""), "name": x.get("name", ""), "email": x.get("email", ""), "phone": x.get("phone", "")} for x in _g_ex if isinstance(x, dict)]
            self.after(0, self.save_config)
            self.after(0, self.save_config)
            if _excl_tot > 0:
                self._pending_exclusives = (_ol_ex, _g_ex, None, service)
            self._max_photo_kb = None
            compact_sum = f"Sync ✓. Synced: {stats['synced']}/{stats['total']}. Del: {stats['deleted']}. Skip: {stats['skipped']}. Err: {stats['errors']}. Excl rest: {_excl_tot}"
            self._last_compact_summary = compact_sum
            self.config["last_compact_summary"] = compact_sum
            summary = f"Sync complete. Synced: {stats['synced']} out of {stats['total']}. Deleted: {stats['deleted']}. Skipped: {stats['skipped']}. Errors: {stats['errors']}."
            self.log("Information", summary)
            self.after(0, lambda: (self.refresh_photo_label(), self.refresh_exclusives_card()))
            self.update_status_line(summary, 100, 100)
            threading.Thread(target=self.refresh_outlook_counts_async, daemon=True).start()

        except Exception as fatal_err:
            self.log("Error", f"Sync failed: {fatal_err}")
            self.update_status_line(f"Sync failed: {fatal_err}")
        finally:
            pythoncom.CoUninitialize()
            self.is_syncing = False
            self.after(0, lambda: self.btn_sync.configure(state="normal"))



def enforce_single_instance():
    """Empêche le lancement simultané de plusieurs instances/versions sous Windows."""
    try:
        kernel32 = ctypes.windll.kernel32
        mutex = kernel32.CreateMutexW(None, False, "Local\\SyncContactsOG_SingleInstance_Mutex")
        last_err = kernel32.GetLastError()
        if last_err == 183:  # ERROR_ALREADY_EXISTS
            root_tmp = tk.Tk()
            root_tmp.withdraw()
            messagebox.showinfo(
                "SyncContactsOG",
                "Une instance de SyncContactsOG est déjà en cours d'exécution.\n\n"
                "Vérifiez l'icône dans la zone de notification (en bas à droite de la barre des tâches)."
            )
            root_tmp.destroy()
            sys.exit(0)
        return mutex
    except Exception:
        return None


if __name__ == "__main__":
    _SINGLE_INSTANCE_MUTEX = enforce_single_instance()
    app = SyncContactsOGApp()
    app.mainloop()





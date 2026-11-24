import json, sys, urllib.request, urllib.parse, urllib.error

def validate_credentials_json(path: str) -> None:
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)

    cfg = data.get("installed") or data.get("web")
    if not isinstance(cfg, dict):
        raise ValueError("Structure invalide : cle 'installed' ou 'web' manquante dans credentials.json.")

    client_id = cfg.get("client_id", "").strip()
    client_secret = cfg.get("client_secret", "").strip()
    token_uri = cfg.get("token_uri", "https://oauth2.googleapis.com/token").strip()

    if not client_id.endswith(".apps.googleusercontent.com") or not client_secret:
        raise ValueError("client_id ou client_secret malforme dans credentials.json.")

    payload = urllib.parse.urlencode({
        "client_id": client_id,
        "client_secret": client_secret,
        "code": "build_preflight_probe_code",
        "grant_type": "authorization_code",
        "redirect_uri": "http://localhost"
    }).encode("utf-8")

    req = urllib.request.Request(token_uri, data=payload, method="POST")
    req.add_header("Content-Type", "application/x-www-form-urlencoded")

    try:
        with urllib.request.urlopen(req, timeout=8) as resp:
            resp.read()
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", errors="replace")
        try:
            err_json = json.loads(body)
            err_code = err_json.get("error", "")
            err_desc = err_json.get("error_description", "")
        except Exception:
            err_code, err_desc = f"HTTP_{e.code}", body[:200]

        if err_code == "invalid_grant":
            print(f"CREDENTIALS_JSON_VALID_ONLINE ({client_id[:24]}...)")
            return
        raise RuntimeError(f"Client OAuth Google rejete/repudie ({err_code} : {err_desc}). Mettez a jour credentials.json !")
    except Exception as e:
        print(f"CREDENTIALS_JSON_STRUCTURE_OK_OFFLINE ({e})")

if __name__ == "__main__":
    validate_credentials_json(sys.argv[1])

import re, sys

def sync_specs(py_path: str, spec_fr_path: str, spec_en_path: str, version: str, today_iso: str):
    with open(py_path, "r", encoding="utf-8") as f:
        py_src = f.read()

    # Détecter quels modes sont marqués comme "testés" dans SyncContactsOG.py
    tested_modes = set()
    m_tested = re.search(r'is_tested\s*=\s*\(?([^\r\n]+)\)?', py_src)
    if m_tested:
        expr = m_tested.group(1)
        for code in ("OUTLOOK_TO_GOOGLE_ONLY", "GOOGLE_TO_OUTLOOK_ONLY", "MERGE_OUTLOOK_WINS", "MERGE_GOOGLE_WINS", "MERGE_PROMPT"):
            if code in expr:
                tested_modes.add(code)
    if not tested_modes:
        tested_modes.add("OUTLOOK_TO_GOOGLE_ONLY")

    # Mettre à jour SPECIFICATIONS.md (FR)
    with open(spec_fr_path, "r", encoding="utf-8") as f:
        fr = f.read()
    fr = re.sub(r'(?m)^(- \*\*Version courante\*\* : )`[^`]+`', rf'\1`{version}`', fr)
    fr = re.sub(r'(?m)^(- \*\*Date de mise à jour\*\* : )`[^`]+`', rf'\1`{today_iso}`', fr)

    for code in ("OUTLOOK_TO_GOOGLE_ONLY", "GOOGLE_TO_OUTLOOK_ONLY", "MERGE_OUTLOOK_WINS", "MERGE_GOOGLE_WINS", "MERGE_PROMPT"):
        if code in tested_modes:
            fr = re.sub(
                rf'(\|\s*`{code}`\s*\|)\s*[^|]+\|\s*`[^`]+`\s*\|',
                r'\1 Validé et testé en conditions réelles | `✔ Testé` |',
                fr
            )
        else:
            fr = re.sub(
                rf'(\|\s*`{code}`\s*\|)\s*[^|]+\|\s*`[^`]+`\s*\|',
                r'\1 En attente de validation terrain | `⚠ Non testé` |',
                fr
            )
    with open(spec_fr_path, "w", encoding="utf-8") as f:
        f.write(fr)

    # Mettre à jour specifications-en.md (EN)
    with open(spec_en_path, "r", encoding="utf-8") as f:
        en = f.read()
    en = re.sub(r'(?m)^(- \*\*Current Version\*\* : )`[^`]+`', rf'\1`{version}`', en)
    en = re.sub(r'(?m)^(- \*\*Last Updated\*\* : )`[^`]+`', rf'\1`{today_iso}`', en)

    for code in ("OUTLOOK_TO_GOOGLE_ONLY", "GOOGLE_TO_OUTLOOK_ONLY", "MERGE_OUTLOOK_WINS", "MERGE_GOOGLE_WINS", "MERGE_PROMPT"):
        if code in tested_modes:
            en = re.sub(
                rf'(\|\s*`{code}`\s*\|)\s*[^|]+\|\s*`[^`]+`\s*\|',
                r'\1 Validated and tested in production | `✔ Tested` |',
                en
            )
        else:
            en = re.sub(
                rf'(\|\s*`{code}`\s*\|)\s*[^|]+\|\s*`[^`]+`\s*\|',
                r'\1 Pending field validation | `⚠ Untested` |',
                en
            )
    with open(spec_en_path, "w", encoding="utf-8") as f:
        f.write(en)

    print(f"SPECS_SYNCED_FROM_CODE_OK (v{version}, tested={sorted(tested_modes)})")

if __name__ == "__main__":
    sync_specs(sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4], sys.argv[5])

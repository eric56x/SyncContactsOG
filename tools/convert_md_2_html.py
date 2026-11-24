import html, re, sys

def render(md_text):
    out, in_ul, in_tbl = [], False, False
    def fmt(s):
        s = html.escape(s)
        s = re.sub(r'`([^`]+)`', r'<code>\1</code>', s)
        s = re.sub(r'\*\*([^*]+)\*\*', r'<strong>\1</strong>', s)
        s = re.sub(r'\*([^*]+)\*', r'<em>\1</em>', s)
        s = s.replace('✔ Testé', '<span class="badge ok">✔ Testé</span>').replace('✔ Tested', '<span class="badge ok">✔ Tested</span>')
        s = s.replace('⚠ Non testé', '<span class="badge warn">⚠ Non testé</span>').replace('⚠ Untested', '<span class="badge warn">⚠ Untested</span>')
        return s.replace('$\\rightarrow$', '→').replace('$N$', 'N')

    for raw in md_text.splitlines():
        st = raw.strip()
        if not st:
            if in_ul: out.append("</ul>"); in_ul = False
            if in_tbl: out.append("</tbody></table>"); in_tbl = False
            continue
        if st == "---":
            if in_ul: out.append("</ul>"); in_ul = False
            if in_tbl: out.append("</tbody></table>"); in_tbl = False
            out.append("<hr>"); continue
        mh = re.match(r'^(#{1,4})\s+(.*)$', st)
        if mh:
            if in_ul: out.append("</ul>"); in_ul = False
            if in_tbl: out.append("</tbody></table>"); in_tbl = False
            out.append(f"<h{len(mh.group(1))}>{fmt(mh.group(2))}</h{len(mh.group(1))}>"); continue
        if st.startswith("|") and st.endswith("|"):
            cells = [c.strip() for c in st.strip("|").split("|")]
            if all(re.match(r'^:?-+:?$', c) for c in cells): continue
            if not in_tbl:
                out.append("<table><thead><tr>" + "".join(f"<th>{fmt(c)}</th>" for c in cells) + "</tr></thead><tbody>")
                in_tbl = True
            else:
                out.append("<tr>" + "".join(f"<td>{fmt(c)}</td>" for c in cells) + "</tr>")
            continue
        ml = re.match(r'^([-*]|\d+\.)\s+(.*)$', st)
        if ml:
            if not in_ul: out.append("<ul>"); in_ul = True
            out.append(f"<li>{fmt(ml.group(2))}</li>"); continue
        if in_ul: out.append("</ul>"); in_ul = False
        out.append(f"<p>{fmt(st)}</p>")
    if in_ul: out.append("</ul>")
    if in_tbl: out.append("</tbody></table>")
    return "\n".join(out)

if __name__ == "__main__":
    src_md, dst_html, ver = sys.argv[1], sys.argv[2], sys.argv[3]
    lang = sys.argv[4] if len(sys.argv) > 4 else "fr"
    title = f"SyncContactsOG v{ver} - Specifications" if lang.lower() == "en" else f"SyncContactsOG v{ver} - Spécifications"
    text = open(src_md, encoding="utf-8").read()
    css = (
        'body{font-family:"Segoe UI",system-ui,sans-serif;background:#f4f6f9;color:#1f2937;margin:0;padding:32px 16px;line-height:1.6}'
        '.container{max-width:960px;margin:0 auto;background:#fff;padding:40px 48px;border-radius:12px;box-shadow:0 4px 20px rgba(0,0,0,.06)}'
        'h1{color:#0f172a;font-size:26px;border-bottom:2px solid #2563eb;padding-bottom:10px;margin-top:0}'
        'h2{color:#1e3a8a;font-size:20px;margin-top:32px;border-left:4px solid #2563eb;padding-left:12px}'
        'h3{color:#334155;font-size:16px;margin-top:22px}'
        'h4{color:#1f2937;font-size:15px;background:#f8fafc;padding:8px 12px;border-radius:6px;margin-top:18px}'
        'p,li{font-size:14.5px;color:#374151}'
        'code{background:#f1f5f9;padding:2px 6px;border-radius:4px;font-family:Consolas,monospace;font-size:13px}'
        'hr{border:none;border-top:1px solid #e2e8f0;margin:28px 0}'
        'table{width:100%;border-collapse:collapse;margin:16px 0;font-size:14px}'
        'th{background:#1e3a8a;color:#fff;text-align:left;padding:10px 12px}'
        'td{border-bottom:1px solid #e2e8f0;padding:10px 12px;vertical-align:top}'
        'tr:nth-child(even) td{background:#f8fafc}'
        '.badge{display:inline-block;padding:2px 10px;border-radius:999px;font-size:12px;font-weight:600}'
        '.badge.ok{background:#dcfce7;color:#166534}'
        '.badge.warn{background:#f1f5f9;color:#475569;border:1px solid #cbd5e1}'
        '@media print{body{background:#fff;padding:0}.container{box-shadow:none;padding:0}}'
    )
    html_doc = f'<!DOCTYPE html><html lang="{lang.lower()}"><head><meta charset="utf-8"><title>{title}</title><style>{css}</style></head><body><div class="container">{render(text)}</div></body></html>'
    open(dst_html, "w", encoding="utf-8").write(html_doc)

"""
OSAM SEO Rank Refresher — Refresh Token Auth
"""

import streamlit as st
import requests
import math
import re
from datetime import date, datetime
from io import BytesIO
import openpyxl

st.set_page_config(page_title="OSAM SEO Rank Refresher", page_icon="📊", layout="centered")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700&family=DM+Mono:wght@400;500&display=swap');
html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }
.stApp { background: #0f1117; color: #e8eaf0; }
.app-header { background: linear-gradient(135deg,#1a1f2e,#0f1117); border:1px solid #2a2f3e; border-radius:16px; padding:32px 36px; margin-bottom:28px; display:flex; align-items:center; gap:20px; }
.header-title { font-size:26px; font-weight:700; color:#fff; margin:0; }
.header-sub { font-size:13px; color:#6b7280; margin:4px 0 0; }
.card { background:#1a1f2e; border:1px solid #2a2f3e; border-radius:12px; padding:24px; margin-bottom:20px; }
.card-title { font-size:11px; font-weight:600; letter-spacing:.1em; text-transform:uppercase; color:#4b9fff; margin-bottom:16px; }
.status-box { border-radius:10px; padding:14px 18px; margin:12px 0; font-size:13.5px; font-weight:500; line-height:1.5; }
.status-info  { background:#1e2a3e; border:1px solid #2a4a7f; color:#7eb8ff; }
.status-ok    { background:#1a2e1e; border:1px solid #2a5a30; color:#6ddb7a; }
.status-warn  { background:#2e2a1a; border:1px solid #5a4a20; color:#dbb86d; }
.status-error { background:#2e1a1a; border:1px solid #5a2a2a; color:#db6d6d; }
.metric-card { background:#0f1117; border:1px solid #2a2f3e; border-radius:10px; padding:16px; text-align:center; }
.metric-value { font-size:32px; font-weight:700; font-family:'DM Mono',monospace; line-height:1; margin-bottom:6px; }
.metric-label { font-size:11px; color:#6b7280; text-transform:uppercase; letter-spacing:.08em; }
.stButton > button { width:100%; background:linear-gradient(135deg,#2563eb,#1d4ed8)!important; color:white!important; border:none!important; border-radius:10px!important; padding:14px 24px!important; font-size:15px!important; font-weight:600!important; }
</style>
""", unsafe_allow_html=True)


def get_config():
    try:
        return {
            "client_id":     st.secrets["onedrive"]["client_id"],
            "tenant_id":     st.secrets["onedrive"]["tenant_id"],
            "file_res_id":   st.secrets["onedrive"]["file_res_id"],
            "refresh_token": st.secrets["onedrive"]["refresh_token"],
        }
    except Exception as e:
        raise Exception(f"Missing secrets: {e}")


def get_access_token(client_id, tenant_id, refresh_token):
    """Exchange refresh token for access token — no user interaction needed."""
    url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    data = {
        "client_id":     client_id,
        "grant_type":    "refresh_token",
        "refresh_token": refresh_token,
        "scope":         "https://graph.microsoft.com/Files.ReadWrite https://graph.microsoft.com/User.Read offline_access",
    }
    resp = requests.post(url, data=data, timeout=30)
    result = resp.json()
    if "access_token" in result:
        return result["access_token"]
    raise Exception(result.get("error_description") or result.get("error") or str(result))


def find_file_id(token, file_res_id):
    """Try multiple methods to find the correct Graph API file ID."""
    headers = {"Authorization": f"Bearer {token}"}
    
    # Method 1: Try direct item ID
    url = f"https://graph.microsoft.com/v1.0/me/drive/items/{file_res_id}"
    resp = requests.get(url, headers=headers, timeout=30)
    if resp.status_code == 200:
        return file_res_id

    # Method 2: Try via shares (using the share URL)
    import base64
    share_url = "https://1drv.ms/x/c/f754eedfd662ac76/IQBf4DbfaWF-SJSqxppjO-rxASQkZP7u4PvPoVutrNOQgbI?e=E50wiO"
    encoded = base64.b64encode(share_url.encode()).decode().rstrip("=").replace("/","_").replace("+","-")
    url = f"https://graph.microsoft.com/v1.0/shares/u!{encoded}/driveItem"
    resp = requests.get(url, headers=headers, timeout=30)
    if resp.status_code == 200:
        return resp.json()["id"]

    # Method 3: Search for Excel files
    url = "https://graph.microsoft.com/v1.0/me/drive/root/search(q='xlsx')"
    resp = requests.get(url, headers=headers, timeout=30)
    if resp.status_code == 200:
        items = resp.json().get("value", [])
        if items:
            # Return first xlsx file found
            return items[0]["id"]

    raise Exception(f"Could not find file. Status: {resp.status_code} - {resp.text}")

def download_workbook(token, file_res_id):
    real_id = find_file_id(token, file_res_id)
    url = f"https://graph.microsoft.com/v1.0/me/drive/items/{real_id}/content"
    headers = {"Authorization": f"Bearer {token}"}
    resp = requests.get(url, headers=headers, timeout=60, allow_redirects=True)
    resp.raise_for_status()
    # Store the real ID for upload
    import streamlit as st
    st.session_state["real_file_id"] = real_id
    return BytesIO(resp.content)


def upload_workbook(token, file_res_id, workbook_bytes):
    url = f"https://graph.microsoft.com/v1.0/me/drive/items/{file_res_id}/content"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    }
    resp = requests.put(url, headers=headers, data=workbook_bytes, timeout=120)
    resp.raise_for_status()


# ── Rank logic ───────────────────────────────────────────────────────────────

def parse_rank(val):
    if val is None: return "NA"
    s = str(val).strip().upper()
    if s in ("", "NA", "NULL"): return "NA"
    if not re.match(r'^\d+$', s): return "NA"
    n = int(s)
    return n if 1 <= n <= 100 else "NA"

def get_group(rank):
    return None if rank == "NA" else math.ceil(rank / 10)

def get_rank_status(p1, p2):
    if p1 == "NA" and p2 == "NA": return "na"
    if p1 == "NA": return "up"
    if p2 == "NA": return "down"
    g1, g2 = get_group(p1), get_group(p2)
    if g1 == g2: return "same"
    return "up" if p2 < p1 else "down"

def calculate_metrics(keywords):
    total = first_page = rank_up = rank_down = top5 = top3 = 0
    for kw in keywords:
        p1, p2 = parse_rank(kw["prev"]), parse_rank(kw["curr"])
        if p1 == "NA" and p2 == "NA": continue
        total += 1
        if p2 != "NA":
            if p2 <= 10: first_page += 1
            if p2 <= 5:  top5 += 1
            if p2 <= 3:  top3 += 1
        s = get_rank_status(p1, p2)
        if s == "up":   rank_up += 1
        if s == "down": rank_down += 1
    return {"total": total, "first_page": first_page, "rank_up": rank_up,
            "rank_down": rank_down, "top5": top5, "top3": top3}


# ── Date logic ───────────────────────────────────────────────────────────────

MONTH_MAP = {
    "january":1,"february":2,"march":3,"april":4,"may":5,"june":6,
    "july":7,"august":8,"september":9,"october":10,"november":11,"december":12,
    "jan":1,"feb":2,"mar":3,"apr":4,"jun":6,"jul":7,"aug":8,
    "sep":9,"sept":9,"oct":10,"nov":11,"dec":12
}

def parse_date_from_sheet_name(name):
    lower = name.lower()
    month_num = None
    for m, n in sorted(MONTH_MAP.items(), key=lambda x: -len(x[0])):
        if re.search(r'\b' + m + r'\b', lower):
            month_num = n; break
    if not month_num: return None
    day_match = re.search(r'\b(\d{1,2})(st|nd|rd|th)?\b', lower)
    if not day_match: return None
    day = int(day_match.group(1))
    if day < 1 or day > 31: return None
    try: return date(date.today().year, month_num, day)
    except: return None

def parse_cell_date(val):
    if val is None or val == "": return None
    if isinstance(val, (date, datetime)):
        return val.date() if isinstance(val, datetime) else val
    s = str(val).strip()
    if "\n" in s: s = s.split("\n")[-1].strip()
    m = re.match(r'^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{2,4})$', s)
    if m:
        d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if y < 100: y += 2000
        try: return date(y, mo, d)
        except: pass
    m = re.match(r'^(\d{4})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})$', s)
    if m:
        try: return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except: pass
    m = re.match(r'(\d{1,2})\s+([a-zA-Z]+)\s+(\d{4})', s)
    if m:
        mon = MONTH_MAP.get(m.group(2).lower())
        if mon:
            try: return date(int(m.group(3)), mon, int(m.group(1)))
            except: pass
    m = re.match(r'([a-zA-Z]+)\s+(\d{1,2})[,\s]+(\d{4})', s)
    if m:
        mon = MONTH_MAP.get(m.group(1).lower())
        if mon:
            try: return date(int(m.group(3)), mon, int(m.group(2)))
            except: pass
    return None

def find_closest_date_col(date_cols, target_date, tolerance_days=10):
    best, best_diff = None, float("inf")
    for dc in date_cols:
        diff = abs((dc["date"] - target_date).days)
        if diff < best_diff and diff <= tolerance_days:
            best_diff, best = diff, dc
    return best

def find_previous_month_end_col(date_cols, current_date):
    best, best_date = None, None
    for dc in date_cols:
        d = dc["date"]
        if d < current_date and d.day >= 25:
            if best_date is None or d > best_date:
                best_date, best = d, dc
    return best


# ── Sheet processing ──────────────────────────────────────────────────────────

def extract_sheet_name_from_hyperlink(cell):
    if hasattr(cell, 'hyperlink') and cell.hyperlink:
        target = cell.hyperlink.target or ""
        if target.startswith("#"):
            ref = target[1:]
            return ref.split("!")[0].strip("'") if "!" in ref else ref.strip("'")
    val = str(cell.value or "")
    m = re.search(r'HYPERLINK\s*\(\s*["\']#([^"\'!]+)', val, re.IGNORECASE)
    if m: return m.group(1).strip("'").strip()
    return str(cell.value).strip() if cell.value else None

def get_rank_summary_sheets(wb):
    return [n for n in wb.sheetnames if re.search(r'rank\s*summary', n, re.IGNORECASE)]

def process_project_sheet(proj_sheet, current_date):
    all_rows = list(proj_sheet.iter_rows(values_only=True))
    if len(all_rows) < 2: return {"error": "Sheet too short"}
    best_idx, best_count = -1, 0
    for r_idx, row in enumerate(all_rows[:3]):
        count = sum(1 for c in row if parse_cell_date(c))
        if count > best_count: best_count, best_idx = count, r_idx
    if best_idx == -1: return {"error": "No date columns found"}
    header_row = all_rows[best_idx]
    date_cols = [{"col_idx": i, "date": parse_cell_date(v)} for i, v in enumerate(header_row) if parse_cell_date(v)]
    if not date_cols: return {"error": "No parseable dates"}
    curr_col = find_closest_date_col(date_cols, current_date)
    if not curr_col: return {"error": f"No column near {current_date}"}
    prev_col = find_previous_month_end_col(date_cols, curr_col["date"])
    if not prev_col: return {"error": "No previous month-end column"}
    keywords = []
    for row in all_rows[best_idx + 1:]:
        sl = row[0] if row else None
        if not sl or str(sl).strip() == "": continue
        try: int(float(str(sl).strip()))
        except: continue
        kw = row[1] if len(row) > 1 else None
        if not kw or str(kw).strip() == "": continue
        keywords.append({
            "prev": row[prev_col["col_idx"]] if prev_col["col_idx"] < len(row) else None,
            "curr": row[curr_col["col_idx"]] if curr_col["col_idx"] < len(row) else None,
        })
    if not keywords: return {"error": "No keyword rows"}
    m = calculate_metrics(keywords)
    m["curr_date"] = curr_col["date"].strftime("%d/%m/%Y")
    m["prev_date"] = prev_col["date"].strftime("%d/%m/%Y")
    return m

def refresh_summary_sheet(wb, summary_sheet_name, progress_callback=None):
    summary_sheet = wb[summary_sheet_name]
    current_date = parse_date_from_sheet_name(summary_sheet_name)
    if not current_date:
        return 0, 0, [f"Could not parse date from: '{summary_sheet_name}'"], {}
    sheet_by_name = {s.lower(): wb[s] for s in wb.sheetnames}
    all_rows = list(summary_sheet.iter_rows(min_row=2))
    project_rows = [r for r in all_rows if len(r) > 1 and r[1] is not None and r[1].value is not None]
    success_count = error_count = 0
    error_details = []
    totals = {"total": 0, "first_page": 0, "rank_up": 0, "rank_down": 0, "top5": 0, "top3": 0}
    for idx, row in enumerate(project_rows):
        cell_b = row[1]
        proj_name = extract_sheet_name_from_hyperlink(cell_b)
        display   = str(cell_b.value).strip() if cell_b.value else f"Row {row[0].row}"
        if progress_callback: progress_callback(idx, len(project_rows), display)
        proj_sheet = sheet_by_name.get((proj_name or "").lower()) or sheet_by_name.get(display.lower())
        if not proj_sheet:
            for c in range(3, 9):
                if c < len(row): row[c].value = "Error"
            error_count += 1
            error_details.append(f"Row {row[0].row} — {display}: Tab not found")
            continue
        result = process_project_sheet(proj_sheet, current_date)
        if "error" in result:
            for c in range(3, 9):
                if c < len(row): row[c].value = "Error"
            error_count += 1
            error_details.append(f"Row {row[0].row} — {display}: {result['error']}")
        else:
            row[3].value = result["total"]
            row[4].value = result["first_page"]
            row[5].value = result["rank_up"]
            row[6].value = result["rank_down"]
            row[7].value = result["top5"]
            row[8].value = result["top3"]
            success_count += 1
            for k in totals: totals[k] += result.get(k, 0)
    return success_count, error_count, error_details, totals


# ── UI ────────────────────────────────────────────────────────────────────────

def main():
    st.markdown("""
    <div class="app-header">
        <div style="font-size:40px">📊</div>
        <div>
            <p class="header-title">OSAM SEO Rank Refresher</p>
            <p class="header-sub">Automated rank summary calculator · OneDrive sync</p>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # Load config & get token silently
    try:
        config = get_config()
        token  = get_access_token(
            config["client_id"],
            config["tenant_id"],
            config["refresh_token"]
        )
    except Exception as e:
        st.markdown(f'<div class="status-box status-error">❌ Auth error: {str(e)}</div>', unsafe_allow_html=True)
        return

    for k in ["workbook_bytes","workbook","sheets_list","connected"]:
        if k not in st.session_state:
            st.session_state[k] = None
    if "connected" not in st.session_state:
        st.session_state.connected = False

    # ── Step 1: Load File ─────────────────────────────────────
    st.markdown('<div class="card"><div class="card-title">Step 1 — Load File from OneDrive</div>', unsafe_allow_html=True)

    if not st.session_state.connected:
        st.markdown('<div class="status-box status-info">📂 Click below to load your Excel file from OneDrive.</div>', unsafe_allow_html=True)

        if st.button("📂 Load File from OneDrive"):
            with st.spinner("Loading file…"):
                try:
                    wb_bytes = download_workbook(token, config["file_res_id"])
                    st.session_state.workbook_bytes = wb_bytes.getvalue()
                    wb = openpyxl.load_workbook(BytesIO(st.session_state.workbook_bytes))
                    st.session_state.workbook    = wb
                    st.session_state.sheets_list = get_rank_summary_sheets(wb)
                    st.session_state.connected   = True
                    st.rerun()
                except Exception as e:
                    st.markdown(f'<div class="status-box status-error">❌ Failed: {str(e)}</div>', unsafe_allow_html=True)

        if st.button("🔍 Debug: List My OneDrive Files"):
            with st.spinner("Listing files…"):
                headers = {"Authorization": f"Bearer {token}"}
                # Check who we are
                me = requests.get("https://graph.microsoft.com/v1.0/me", headers=headers).json()
                st.write(f"Logged in as: {me.get('userPrincipalName', me.get('mail', 'unknown'))}")
                # List root files
                resp = requests.get("https://graph.microsoft.com/v1.0/me/drive/root/children", headers=headers)
                st.write(f"Drive status: {resp.status_code}")
                if resp.status_code == 200:
                    items = resp.json().get("value", [])
                    for item in items:
                        st.write(f"📄 {item['name']} — ID: {item['id']}")
                else:
                    st.write(resp.json())
    else:
        wb = st.session_state.workbook
        st.markdown(f"""
        <div class="status-box status-ok">
            ✅ File loaded · <strong>{len(wb.sheetnames)}</strong> sheets ·
            {len(st.session_state.sheets_list)} Rank Summary sheet(s) found
        </div>""", unsafe_allow_html=True)
        if st.button("🔄 Reload File"):
            for k in ["workbook_bytes","workbook","sheets_list"]:
                st.session_state[k] = None
            st.session_state.connected = False
            st.rerun()

    st.markdown('</div>', unsafe_allow_html=True)

    # ── Step 2: Select Sheet ──────────────────────────────────
    if st.session_state.connected and st.session_state.sheets_list:
        st.markdown('<div class="card"><div class="card-title">Step 2 — Select Rank Summary Sheet</div>', unsafe_allow_html=True)
        selected_sheet = st.selectbox("Sheet:", options=st.session_state.sheets_list, label_visibility="collapsed")
        detected_date  = parse_date_from_sheet_name(selected_sheet)
        if detected_date:
            wb = st.session_state.workbook
            project_count = sum(
                1 for row in wb[selected_sheet].iter_rows(min_row=2, values_only=True)
                if row[1] is not None and str(row[1]).strip() != ""
            )
            st.markdown(f"""
            <div class="status-box status-info">
                📅 Date: <strong>{detected_date.strftime('%d %B %Y')}</strong> &nbsp;·&nbsp;
                🗂 Projects: <strong>{project_count}</strong>
            </div>""", unsafe_allow_html=True)
        else:
            st.markdown('<div class="status-box status-warn">⚠️ Could not detect date. Use format like "March 31st Rank Summary".</div>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

        # ── Step 3: Refresh ───────────────────────────────────
        if detected_date:
            st.markdown('<div class="card"><div class="card-title">Step 3 — Refresh</div>', unsafe_allow_html=True)
            if st.button(f"🚀 Refresh '{selected_sheet}'"):
                wb = openpyxl.load_workbook(BytesIO(st.session_state.workbook_bytes))
                progress_bar  = st.progress(0)
                status_text   = st.empty()
                log_container = st.empty()
                log_lines     = []

                def progress_callback(idx, total, name):
                    progress_bar.progress(int((idx / max(total,1)) * 100))
                    status_text.markdown(f'<div class="status-box status-info">⏳ {idx+1}/{total}: <strong>{name}</strong></div>', unsafe_allow_html=True)
                    log_lines.append(f"✓ {name}")
                    log_container.text("\n".join(log_lines[-8:]))

                with st.spinner("Calculating…"):
                    try:
                        success, errors, error_details, totals = refresh_summary_sheet(wb, selected_sheet, progress_callback)
                        progress_bar.progress(100)
                        status_text.empty(); log_container.empty()

                        out = BytesIO()
                        wb.save(out)
                        updated_bytes = out.getvalue()

                        status_text.markdown('<div class="status-box status-info">☁️ Uploading to OneDrive…</div>', unsafe_allow_html=True)
                        real_id = st.session_state.get("real_file_id", config["file_res_id"])
                        upload_workbook(token, real_id, updated_bytes)
                        status_text.empty()

                        st.session_state.workbook_bytes = updated_bytes
                        st.session_state.workbook = openpyxl.load_workbook(BytesIO(updated_bytes))

                        total_projects = success + errors
                        st.markdown(f"""
                        <div class="status-box status-ok">
                            ✅ Done! <strong>{success}/{total_projects}</strong> projects updated successfully.
                            {f'<br>⚠️ {errors} had errors.' if errors > 0 else ''}
                        </div>""", unsafe_allow_html=True)

                        if success > 0:
                            metrics_display = [
                                ("Total Keywords", totals["total"],      "#4b9fff"),
                                ("1st Page",        totals["first_page"], "#818cf8"),
                                ("Rank Up ↑",       totals["rank_up"],    "#34d399"),
                                ("Rank Down ↓",     totals["rank_down"],  "#f87171"),
                                ("Top 5",           totals["top5"],       "#60a5fa"),
                                ("Top 3",           totals["top3"],       "#fbbf24"),
                            ]
                            cols = st.columns(3)
                            for i, (label, val, color) in enumerate(metrics_display):
                                with cols[i % 3]:
                                    st.markdown(f"""
                                    <div class="metric-card">
                                        <div class="metric-value" style="color:{color}">{val}</div>
                                        <div class="metric-label">{label}</div>
                                    </div>""", unsafe_allow_html=True)

                        if error_details:
                            with st.expander(f"⚠️ {errors} Error(s)"):
                                for ed in error_details:
                                    st.markdown(f"- {ed}")

                    except Exception as e:
                        status_text.empty()
                        st.markdown(f'<div class="status-box status-error">❌ Refresh failed: {str(e)}</div>', unsafe_allow_html=True)

            st.markdown('</div>', unsafe_allow_html=True)

    elif st.session_state.connected and not st.session_state.sheets_list:
        st.markdown('<div class="status-box status-warn">⚠️ No Rank Summary sheets found.</div>', unsafe_allow_html=True)

    st.markdown("""
    <div style="text-align:center;color:#374151;font-size:12px;margin-top:40px;padding-top:20px;border-top:1px solid #1e2332;">
        OSAM SEO Rank Refresher · Built for internal use · Powered by Microsoft Graph API
    </div>""", unsafe_allow_html=True)


if __name__ == "__main__":
    main()

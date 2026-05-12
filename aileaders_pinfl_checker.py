"""
AI Leaders PINFL Checker — Checkpoint + Resume
===============================================
requirements.txt:
    streamlit
    openpyxl
    pandas
    requests
"""
import re
import time
import datetime
import requests
import pandas as pd
import streamlit as st
from io import BytesIO

# ──────────────────────────────────────────
# CONFIG
# ──────────────────────────────────────────
API_URL    = "https://aileaders.uz/api/v1/check/certificates"
DELAY_SEC  = 1.5
MAX_KURS   = 10
SAVE_EVERY = 50

# ──────────────────────────────────────────
# SESSION STATE INIT
# ──────────────────────────────────────────
def init_state():
    defaults = {
        "all_kurslar":  [],
        "emails":       [],
        "holats":       [],
        "start_idx":    0,
        "result_df":    None,
        "df_orig":      None,
        "pinfl_col":    "",
        "date_col":     "",
        "finished":     False,
        "cookie_dead":  False,
        "save_counter": 0,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

init_state()

# ──────────────────────────────────────────
# cURL PARSE
# ──────────────────────────────────────────
def parse_curl(curl_text: str) -> dict:
    result = {
        "cookie": "", "user_agent": "",
        "accept": "*/*", "accept_language": "",
        "referer": "", "extra_headers": {}
    }
    for pattern in [
        r"-b\s+'([^']+)'", r'-b\s+"([^"]+)"',
        r"-H\s+'cookie:\s*([^']+)'", r'-H\s+"cookie:\s*([^"]+)"',
    ]:
        m = re.search(pattern, curl_text, re.IGNORECASE)
        if m:
            result["cookie"] = m.group(1).strip()
            break
    for m in re.finditer(r"-H\s+'([^']+)'|-H\s+\"([^\"]+)\"", curl_text):
        raw = (m.group(1) or m.group(2)).strip()
        if ":" not in raw:
            continue
        key, _, val = raw.partition(":")
        key = key.strip().lower()
        val = val.strip()
        if key == "user-agent":
            result["user_agent"] = val
        elif key == "accept" and "language" not in key:
            result["accept"] = val
        elif key == "accept-language":
            result["accept_language"] = val
        elif key == "referer":
            result["referer"] = val
        elif key not in ("cookie", "content-length"):
            result["extra_headers"][key] = val
    return result

# ──────────────────────────────────────────
# EXCEL O'QISH
# ──────────────────────────────────────────
def read_excel(file):
    xls = pd.ExcelFile(file, engine="openpyxl")
    sheet_name = xls.sheet_names[0]
    header_row = 0
    for i in range(8):
        try:
            df_tmp = pd.read_excel(xls, sheet_name=sheet_name, header=i, nrows=1, engine="openpyxl")
            cols = df_tmp.columns.astype(str).str.upper()
            if cols.str.contains(r"ПИНФЛ|PINFL", regex=True).any():
                header_row = i
                break
        except Exception:
            continue
    df = pd.read_excel(xls, sheet_name=sheet_name, header=header_row, engine="openpyxl")
    df.columns = df.columns.map(str).str.strip()
    pinfl_col = next(
        (c for c in df.columns if "ПИНФЛ" in c.upper() or "PINFL" in c.upper()), None
    )
    if pinfl_col is None:
        return pd.DataFrame(), "", ""
    date_col = next(
        (c for c in df.columns
         if "рождени" in c.lower() or "birth" in c.lower() or "sana" in c.lower()), None
    )
    df = df[df[pinfl_col].notna()].copy()
    df[pinfl_col] = (
        df[pinfl_col].astype(str).str.strip()
        .str.replace(r"\.0$", "", regex=True)
        .str.replace(r"\s+", "", regex=True)
    )
    df = df[df[pinfl_col].str.len() >= 10].reset_index(drop=True)
    if date_col:
        df[date_col] = pd.to_datetime(df[date_col], errors="coerce")
    return df, pinfl_col, date_col or ""

# ──────────────────────────────────────────
# BITTA PINFL TEKSHIRISH
# ──────────────────────────────────────────
def check_pinfl(pinfl: str, session: requests.Session) -> dict:
    empty = {"holat": "🔴 Xato", "ism": "", "email": "", "kurslar": [], "xato": ""}
    try:
        r = session.get(API_URL, params={"pinfl": pinfl}, timeout=20)
        if r.status_code in (200, 304):
            try:
                data = r.json()
            except Exception:
                return {**empty, "holat": "⚠️ JSON xato", "xato": r.text[:80]}
            all_courses = data.get("courses", [])
            completed   = [c for c in all_courses if c.get("isCompleted")]
            kurslar = []
            for c in completed:
                url    = c.get("contentCertificateUrl") or ""
                raw_ts = c.get("completedAt") or ""
                if str(raw_ts).isdigit() and int(raw_ts) > 1_000_000_000:
                    ts = int(raw_ts)
                    if ts > 1e10:
                        ts = ts / 1000
                    vaqt = datetime.datetime.fromtimestamp(ts).strftime("%d.%m.%Y")
                else:
                    vaqt = str(raw_ts)
                kurslar.append({
                    "nomi": c.get("contentName", "").strip(),
                    "url":  url,
                    "vaqt": vaqt,
                })
            has_courses = data.get("hasCourses", False)
            return {
                "holat":   "✅ Sertifikat bor" if kurslar else (
                           "⚠️ Kurs bor, sertifikat yoq" if has_courses else
                           "⚠️ Royxatda bor, kurs yoq"),
                "ism":     data.get("fullName", ""),
                "email":   data.get("email", ""),
                "kurslar": kurslar,
                "xato":    "",
            }
        elif r.status_code == 404:
            return {**empty, "holat": "❌ Topilmadi",        "kurslar": []}
        elif r.status_code == 401:
            return {**empty, "holat": "🔐 Cookie eskirgan",  "kurslar": [], "xato": "Cookie yangilang"}
        elif r.status_code == 429:
            time.sleep(60)
            return {**empty, "holat": "Rate limit",          "kurslar": [], "xato": "1 daqiqa kutildi"}
        else:
            return {**empty, "holat": f"🔴 {r.status_code}", "kurslar": [], "xato": r.text[:80]}
    except Exception as e:
        return {**empty, "holat": "🔴 Xato", "kurslar": [], "xato": str(e)[:80]}

# ──────────────────────────────────────────
# RESULT DF YASASH
# ──────────────────────────────────────────
def build_result_df():
    df_orig     = st.session_state.df_orig
    all_kurslar = st.session_state.all_kurslar
    emails      = st.session_state.emails
    holats      = st.session_state.holats

    result_df = df_orig.copy()
    result_df["Email"] = ""
    result_df["Holat"] = ""
    for n in range(MAX_KURS):
        result_df[f"Kurs {n+1} nomi"] = ""
        result_df[f"Kurs {n+1} URL"]  = ""
        result_df[f"Kurs {n+1} vaqti"]= ""

    for idx in range(len(all_kurslar)):
        result_df.at[idx, "Email"] = emails[idx] if idx < len(emails) else ""
        result_df.at[idx, "Holat"] = holats[idx] if idx < len(holats) else ""
        kurslar = all_kurslar[idx]
        for n in range(MAX_KURS):
            if n < len(kurslar):
                result_df.at[idx, f"Kurs {n+1} nomi"] = kurslar[n]["nomi"]
                result_df.at[idx, f"Kurs {n+1} URL"]  = kurslar[n]["url"]
                result_df.at[idx, f"Kurs {n+1} vaqti"]= kurslar[n]["vaqt"]
    return result_df

# ──────────────────────────────────────────
# EXCEL EKSPORT
# ──────────────────────────────────────────
def export_excel(result_df: pd.DataFrame, pinfl_col: str, date_col: str) -> bytes:
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        result_df.to_excel(writer, index=False, sheet_name="Natijalar")
        ws = writer.sheets["Natijalar"]
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            for cell in row:
                col_name = str(ws.cell(1, cell.column).value or "")
                if "ПИНФЛ" in col_name.upper() or "PINFL" in col_name.upper():
                    cell.number_format = "@"
                    if cell.value is not None:
                        cell.value = str(cell.value).replace(".0", "").strip()
                elif date_col and col_name == date_col:
                    cell.number_format = "DD.MM.YYYY"
                elif any(k in col_name for k in ["Серия", "Номер", "ID", "№"]):
                    cell.number_format = "@"
                    if cell.value is not None:
                        cell.value = str(cell.value).replace(".0", "").strip()
                elif "URL" in col_name.upper() or "vaqti" in col_name.lower():
                    cell.number_format = "@"
    return out.getvalue()

# ──────────────────────────────────────────
# TEKSHIRISH JARAYONI
# ──────────────────────────────────────────
def run_check(parsed: dict):
    session = requests.Session()
    session.headers.update({
        "User-Agent":      parsed.get("user_agent") or (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/148.0.0.0 Safari/537.36"
        ),
        "Accept":          parsed.get("accept", "*/*"),
        "Accept-Language": parsed.get("accept_language", "ru,en-US;q=0.9,en;q=0.8"),
        "Referer":         parsed.get("referer", "https://aileaders.uz/auth/login/check"),
        "Cookie":          parsed["cookie"],
        **parsed.get("extra_headers", {}),
    })

    df_orig    = st.session_state.df_orig
    pinfl_col  = st.session_state.pinfl_col
    start_from = st.session_state.start_idx
    total      = len(df_orig)

    progress_bar = st.progress(start_from / total if total else 0)
    status_text  = st.empty()
    save_info    = st.empty()

    for i in range(start_from, total):
        row   = df_orig.iloc[i]
        pinfl = str(row[pinfl_col])
        fio   = str(row.get("Полное наименование", row.get("F.I.Sh.", "")))
        status_text.text(f"Tekshirilmoqda {i+1}/{total} — {pinfl} | {fio[:40]}")
        progress_bar.progress((i + 1) / total)

        res = check_pinfl(pinfl, session)

        if res["holat"] == "🔐 Cookie eskirgan":
            st.session_state.start_idx   = i
            st.session_state.cookie_dead = True
            st.session_state.result_df   = build_result_df()
            st.error(
                f"Cookie eskirdi! {i+1}-qatorda toxtadi. "
                f"Yangi cURL kiritib 'Davom etish' tugmasini bosing."
            )
            return

        st.session_state.all_kurslar.append(res["kurslar"])
        st.session_state.emails.append(res["email"])
        st.session_state.holats.append(res["holat"])
        st.session_state.save_counter += 1

        # Har SAVE_EVERY da avtosaqlash
        if st.session_state.save_counter % SAVE_EVERY == 0:
            st.session_state.result_df = build_result_df()
            save_info.info(f"Avtosaqlash: {i+1} ta tayyor")

        time.sleep(DELAY_SEC)

    st.session_state.start_idx = total
    st.session_state.finished  = True
    st.session_state.result_df = build_result_df()
    status_text.success(f"Yakunlandi! {total} ta tekshirildi.")
    progress_bar.progress(1.0)

# ──────────────────────────────────────────
# UI
# ──────────────────────────────────────────
st.set_page_config(page_title="AI Leaders PINFL Checker", page_icon="🎓", layout="wide")
st.title("🎓 AI Leaders — PINFL Sertifikat Tekshiruvi")
st.caption("1000+ o'quvchi | Cookie eskirsa to'xtagan joydan davom etadi | Har 50 ta avtosaqlash")

# ── cURL ──
st.subheader("🔐 cURL joylashtiring")
with st.expander("📋 Qanday olish kerak?", expanded=st.session_state.cookie_dead):
    st.markdown("""
    1. **`https://aileaders.uz/auth/login/check`** ga kiring
    2. Istalgan PINFL → **Tekshirish**
    3. **F12 → Network** → `certificates?pinfl=...` → o'ng klik
    4. **Copy as cURL (bash)** → quyiga **Ctrl+V**
    """)

curl_input = st.text_area(
    "cURL:",
    placeholder="curl 'https://aileaders.uz/api/v1/check/certificates?pinfl=...' \\\n  -b 'HWWAFSESID=...' ...",
    height=100,
)

parsed  = {}
curl_ok = False
if curl_input.strip():
    parsed = parse_curl(curl_input)
    if parsed["cookie"] and "HWWAFSESID" in parsed["cookie"]:
        st.success("✅ Cookie topildi!")
        curl_ok = True
    else:
        st.error("❌ Cookie topilmadi.")

st.divider()

# ── EXCEL ──
st.subheader("📂 Excel yuklang")

if st.session_state.df_orig is None:
    uploaded = st.file_uploader("O'quvchilar ro'yxati (xlsx)", type=["xlsx"], label_visibility="collapsed")
    if uploaded:
        with st.spinner("O'qilmoqda..."):
            df_orig, pinfl_col, date_col = read_excel(uploaded)
        if df_orig.empty:
            st.error("❌ PINFL ustuni topilmadi!")
            st.stop()
        st.session_state.df_orig   = df_orig
        st.session_state.pinfl_col = pinfl_col
        st.session_state.date_col  = date_col
        st.rerun()
else:
    df_orig   = st.session_state.df_orig
    pinfl_col = st.session_state.pinfl_col
    total     = len(df_orig)
    done      = len(st.session_state.all_kurslar)

    st.success(f"✅ **{total}** ta o'quvchi | PINFL ustuni: **{pinfl_col}**")

    if done > 0:
        pct = round(done / total * 100, 1)
        st.info(f"📊 **{done}/{total}** ta tekshirilgan ({pct}%) | Qoldi: **{total - done}** ta")

    daqiqa = round((total - done) * DELAY_SEC / 60, 1)
    st.caption(f"⏱ Qolgan vaqt: ~{daqiqa} daqiqa")

    st.divider()
    st.subheader("🎮 Boshqaruv")
    col1, col2 = st.columns([2, 1])

    # COOKIE ESKIRGAN — DAVOM ETISH
    if st.session_state.cookie_dead:
        st.warning(f"⚠️ Cookie eskirdi! **{done}-qatorda** to'xtadi. Yangi cURL kiritib davom eting.")
        if curl_ok:
            if col1.button("▶️ Davom etish", type="primary", use_container_width=True):
                st.session_state.cookie_dead = False
                run_check(parsed)
                st.rerun()
        else:
            st.info("Yangi cURL kiriting ☝️")

    # TUGALLANGAN
    elif st.session_state.finished:
        st.success("✅ Barcha o'quvchilar tekshirildi!")

    # BOSHLASH / DAVOM
    else:
        if not curl_ok:
            st.warning("⚠️ cURL kiriting!")
        else:
            btn_label = "🚀 Boshlash" if done == 0 else f"▶️ Davom etish ({done}/{total} dan)"
            if col1.button(btn_label, type="primary", use_container_width=True):
                run_check(parsed)
                st.rerun()

    # YANGIDAN BOSHLASH
    if col2.button("🔄 Yangidan", use_container_width=True):
        keys = ["all_kurslar","emails","holats","start_idx","result_df",
                "df_orig","pinfl_col","date_col","finished","cookie_dead","save_counter"]
        for k in keys:
            if k in st.session_state:
                del st.session_state[k]
        st.rerun()

    # NATIJA
    if st.session_state.result_df is not None:
        st.divider()
        result_df = st.session_state.result_df

        sertifikat_bor = sum(1 for h in st.session_state.holats if "✅" in h)
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Tekshirildi",       done)
        c2.metric("✅ Sertifikat bor", sertifikat_bor)
        c3.metric("❌ Sertifikat yoq", done - sertifikat_bor)
        c4.metric("Qoldi",             total - done)

        st.dataframe(result_df, use_container_width=True)

        label = "📥 To'liq natija" if st.session_state.finished else "📥 Qisman natija (hozircha)"
        excel_bytes = export_excel(result_df, pinfl_col, st.session_state.date_col)
        st.download_button(
            label,
            data=excel_bytes,
            file_name="aileaders_natija.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

st.markdown(
    '<div style="position:fixed;left:0;bottom:0;width:100%;background:#0e1117;'
    'color:white;text-align:center;padding:8px;font-weight:bold;z-index:1000;">'
    'Tuzuvchi: Azamat Madrimov | 2026</div>',
    unsafe_allow_html=True
)

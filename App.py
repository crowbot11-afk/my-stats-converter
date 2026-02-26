import streamlit as st
import pandas as pd
import pytesseract
from PIL import Image
import io
import re
import requests
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Alliance Battle Tracker", page_icon="⚔️", layout="centered")
st.title("⚔️ Alliance Battle Report Tracker")

KNOWN_STATS = [
    "Infantry Attack", "Infantry Defense", "Infantry HP",
    "Cavalry Attack", "Cavalry Defense", "Cavalry HP",
    "Archer Attack", "Archer Defense", "Archer HP",
    "Mage Attack", "Mage Defense", "Mage HP",
    "Angel Attack", "Angel Defense", "Angel HP",
    "Golem Attack", "Golem Defense", "Golem HP",
    "Enemy Troops Attack Reduction", "Enemy Troops HP Reduction",
    "Archer Damage", "Mage Damage", "Troops Damage Taken Reduction",
    "Damage Boost when attacking", "Damage taken reduced when attacking",
    "Damage Boost when defending", "Damage taken reduced when defending",
    "Infantry Resilience", "Cavalry Resilience",
    "Archer Penetration", "Mage Penetration",
    "Archer Mastery", "Mage Mastery",
    "Damage Against Infantry Boost", "Damage Against Cavalry Boost",
    "Damage Against Angels Boost",
    "Infantry Damage Taken Reduction", "Cavalry Damage Taken Reduction",
    "Reduces Damage taken from Infantry", "Reduces Damage taken from Cavalry",
    "Reduces Damage taken from Archers", "Reduces Damage taken from Mages",
    "Mage damage increased on Infantry", "Mage damage increased on Cavalry",
    "Mage damage increased on Archers",
    "Critical Strike Rate Boost", "Critical Strike Rate Taken Reduction",
]

def ocr_image(img):
    return pytesseract.image_to_string(img, config='--psm 6')

def extract_all(all_text):
    data = {}
    m = re.search(r'Total Army\s+([\d,]+)', all_text)
    if m:
        data['March Size'] = m.group(1).replace(',', '')
    for stat in KNOWN_STATS:
        pattern = re.escape(stat).replace(r'\ ', r'\s+')
        m = re.search(pattern + r'\s+([\d,]+\.?\d*%?)', all_text, re.I)
        if m:
            data[stat] = m.group(1).strip()
    m = re.search(r'Evolution:\s*Titan Tier\s*(III|II|I|lll|ll|l|\d)', all_text, re.I)
    if m:
        tier = m.group(1).replace('lll','III').replace('ll','II').replace('l','I')
        data['Elder Titan Tier'] = 'Titan Tier ' + tier
    talent_levels = re.findall(r'Total Talent Level:\s*(\d+)', all_text)
    if talent_levels:
        data['Titan Talent Level'] = talent_levels[0]
    m = re.search(r'[Ee]volution:\s*Tier\s*(\d+)', all_text)
    if m:
        data['Beast Tier'] = m.group(1)
    if len(talent_levels) >= 2:
        data['Beast Talent Level'] = talent_levels[1]
    m = re.search(r'Total [Ss]kill Level:\s*(\d+)', all_text)
    if m:
        data['Beast Skill Level'] = m.group(1)
    m = re.search(r'(?:Level|Lv)[:\s.]*\s*(Lv\.?\s*\d+)', all_text, re.I)
    if m:
        data['Totem Level'] = re.sub(r'\s+', '', m.group(1))
    m = re.search(r'Total Special Stats Level:\s*(\d+)', all_text)
    if m:
        data['Special Stats Level'] = m.group(1)
    m = re.search(r'Total Jewels Level:\s*(\d+)', all_text)
    if m:
        data['Jewels Level'] = m.group(1)
    m = re.search(r'Zodiac Palace.*?Green Amount:\s*(\d+).*?White Amount:\s*(\d+)', all_text, re.S)
    if m:
        data['Zodiac Green'] = m.group(1)
        data['Zodiac White'] = m.group(2)
    m = re.search(r'Northern Palace.*?Green Amount:\s*(\d+)', all_text, re.S)
    if m:
        data['Northern Green'] = m.group(1)
    m = re.search(r'Colossus.*?Total Level:\s*(\d+)', all_text, re.S)
    if m:
        data['Colossus Level'] = m.group(1)
    m = re.search(r'Emblem.*?Total Level:\s*(\d+)', all_text, re.S)
    if m:
        data['Emblem Level'] = m.group(1)
    return data

def build_excel(df):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Alliance Stats')
        ws = writer.sheets['Alliance Stats']
        header_fill = PatternFill('solid', start_color='8B0000')
        header_font = Font(name='Arial', bold=True, color='FFFFFF', size=10)
        alt_fill = PatternFill('solid', start_color='FFF3E0')
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        for i, row in enumerate(ws.iter_rows(min_row=2), start=2):
            for cell in row:
                cell.font = Font(name='Arial', size=10)
                cell.alignment = Alignment(horizontal='center', vertical='center')
                if i % 2 == 0:
                    cell.fill = alt_fill
        for col in ws.columns:
            max_len = max((len(str(c.value)) if c.value else 0) for c in col)
            ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 3, 30)
        ws.row_dimensions[1].height = 30
        ws.freeze_panes = 'B2'
    buf.seek(0)
    return buf

def gdrive_url_to_direct(url):
    """Convert any Google Drive share URL to a direct download URL."""
    # Format: /file/d/FILE_ID/view or /open?id=FILE_ID
    m = re.search(r'/file/d/([a-zA-Z0-9_-]+)', url)
    if not m:
        m = re.search(r'[?&]id=([a-zA-Z0-9_-]+)', url)
    if m:
        file_id = m.group(1)
        return f"https://drive.google.com/uc?export=download&id={file_id}"
    return None

def load_from_gdrive(url):
    direct_url = gdrive_url_to_direct(url)
    if not direct_url:
        return None, "❌ Could not parse Google Drive URL. Make sure it's a valid share link."
    try:
        resp = requests.get(direct_url, timeout=15)
        if resp.status_code != 200:
            return None, f"❌ Could not download file (HTTP {resp.status_code}). Make sure sharing is set to 'Anyone with the link'."
        df = pd.read_excel(io.BytesIO(resp.content))
        return df, None
    except Exception as e:
        return None, f"❌ Error loading file: {str(e)}"

# ── Session state ──
if 'alliance_df' not in st.session_state:
    st.session_state.alliance_df = pd.DataFrame()
if 'last_data' not in st.session_state:
    st.session_state.last_data = None
if 'upload_key' not in st.session_state:
    st.session_state.upload_key = 0
if 'ocr_done' not in st.session_state:
    st.session_state.ocr_done = False

# ══════════════════════════════════════════
# STEP 1: Load existing tracker
# ══════════════════════════════════════════
st.subheader("1️⃣ Load Existing Alliance Tracker")

load_method = st.radio("How do you want to load your tracker?",
    ["Upload Excel file", "Google Drive URL"], horizontal=True)

if load_method == "Upload Excel file":
    existing_file = st.file_uploader("Upload your alliance_tracker.xlsx", type=["xlsx"], key="existing")
    if existing_file:
        st.session_state.alliance_df = pd.read_excel(existing_file)
        st.success(f"✅ Loaded {len(st.session_state.alliance_df)} players from file.")

elif load_method == "Google Drive URL":
    st.info("📋 In Google Drive: right-click your Excel file → Share → 'Anyone with the link' → Copy link")
    gdrive_url = st.text_input("Paste your Google Drive share link here", placeholder="https://drive.google.com/file/d/...")
    if gdrive_url and st.button("📥 Load from Google Drive"):
        with st.spinner("Downloading from Google Drive..."):
            df, err = load_from_gdrive(gdrive_url)
        if err:
            st.error(err)
        else:
            st.session_state.alliance_df = df
            st.success(f"✅ Loaded {len(df)} players from Google Drive!")

if not st.session_state.alliance_df.empty:
    st.caption(f"Tracker currently has **{len(st.session_state.alliance_df)} players**. Scroll down to add more.")

st.divider()

# ══════════════════════════════════════════
# STEP 2: Add a new player
# ══════════════════════════════════════════
st.subheader("2️⃣ Add New Player")

player_name_input = st.text_input("Player in-game name", placeholder="e.g. Fighterrulez", key="pname")

uploaded_files = st.file_uploader(
    "Upload all screenshots for this player",
    type=["png", "jpg", "jpeg", "webp"],
    accept_multiple_files=True,
    key=f"shots_{st.session_state.upload_key}"
)

col_extract, col_clear = st.columns([2, 1])

with col_extract:
    extract_btn = st.button("🔍 Extract Stats", disabled=not (uploaded_files and player_name_input))

with col_clear:
    if st.button("🗑️ Clear & Start Next Player"):
        st.session_state.upload_key += 1
        st.session_state.last_data = None
        st.session_state.ocr_done = False
        st.rerun()

if extract_btn and uploaded_files and player_name_input:
    all_text = ""
    progress = st.progress(0)
    status = st.empty()
    for i, f in enumerate(uploaded_files):
        status.write(f"Reading image {i+1}/{len(uploaded_files)}...")
        img = Image.open(f)
        all_text += "\n" + ocr_image(img)
        progress.progress((i + 1) / len(uploaded_files))
    status.empty()
    progress.empty()

    extracted = extract_all(all_text)
    extracted = {'Player Name': player_name_input, **extracted}
    st.session_state.last_data = extracted
    st.session_state.ocr_done = True

    st.success(f"✅ Extracted {len(extracted)-1} fields for **{player_name_input}**")

    build_keys = ['March Size', 'Elder Titan Tier', 'Titan Talent Level', 'Beast Tier',
                  'Beast Talent Level', 'Beast Skill Level', 'Totem Level',
                  'Special Stats Level', 'Jewels Level', 'Zodiac Green', 'Zodiac White',
                  'Northern Green', 'Colossus Level', 'Emblem Level']

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**🏰 Build Info**")
        build_data = {k: extracted.get(k, '—') for k in build_keys if k in extracted}
        st.dataframe(pd.DataFrame(list(build_data.items()), columns=['Field','Value']), use_container_width=True)

    st.markdown("**📊 Stats Bonus**")
    stats_data = {k: extracted[k] for k in KNOWN_STATS if k in extracted}
    if stats_data:
        items = list(stats_data.items())
        half = len(items)//2 + len(items)%2
        c1, c2 = st.columns(2)
        with c1:
            st.dataframe(pd.DataFrame(items[:half], columns=['Stat','Value']), use_container_width=True)
        with c2:
            st.dataframe(pd.DataFrame(items[half:], columns=['Stat','Value']), use_container_width=True)
        st.info(f"📊 {len(stats_data)}/47 stats extracted")
    else:
        st.warning("No stats found — check raw OCR text below.")

    with st.expander("🔤 Raw OCR Text (debugging)"):
        st.text_area("", all_text, height=200)

# ── Add to tracker button ──
if st.session_state.last_data:
    pname = st.session_state.last_data.get('Player Name', '')
    st.divider()
    if st.button(f"➕ Add {pname} to Alliance Tracker", type="primary"):
        new_row = pd.DataFrame([st.session_state.last_data])
        st.session_state.alliance_df = pd.concat(
            [st.session_state.alliance_df, new_row], ignore_index=True
        )
        st.session_state.last_data = None
        st.session_state.ocr_done = False
        # Auto-increment upload key so screenshots clear automatically
        st.session_state.upload_key += 1
        st.success(f"✅ {pname} added! {len(st.session_state.alliance_df)} players in tracker. Upload next player's screenshots above ⬆️")
        st.rerun()

st.divider()

# ══════════════════════════════════════════
# STEP 3: View & Download
# ══════════════════════════════════════════
if not st.session_state.alliance_df.empty:
    st.subheader(f"📊 Alliance Tracker — {len(st.session_state.alliance_df)} Players")
    st.dataframe(st.session_state.alliance_df, use_container_width=True)
    excel_buf = build_excel(st.session_state.alliance_df)
    st.download_button(
        label="⬇️ Download Alliance Tracker Excel",
        data=excel_buf,
        file_name="alliance_tracker.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    st.info("💡 After downloading, re-upload this file next time to keep adding players.")

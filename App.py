import streamlit as st
import pandas as pd
import pytesseract
from PIL import Image
import io
import re
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

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

ALL_COLUMNS = ['Player Name', 'March Size'] + KNOWN_STATS + [
    'Elder Titan Tier', 'Titan Talent Level', 'Beast Tier', 'Beast Talent Level',
    'Beast Skill Level', 'Totem Level', 'Special Stats Level', 'Jewels Level',
    'Zodiac Green', 'Zodiac White', 'Northern Green', 'Colossus Level', 'Emblem Level'
]

TOTAL_ROWS = 100

# ── Session state init ──
if 'df' not in st.session_state:
    st.session_state.df = pd.DataFrame(columns=ALL_COLUMNS)
if 'roster' not in st.session_state:
    st.session_state.roster = []
if 'extracted' not in st.session_state:
    st.session_state.extracted = None
if 'upload_key' not in st.session_state:
    st.session_state.upload_key = 0

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

def save_player(name, data):
    """Save player to df — update existing row or add to next free row."""
    df = st.session_state.df
    data['Player Name'] = name

    # Check if player already exists
    mask = df['Player Name'].astype(str).str.strip().str.lower() == name.strip().lower()
    if mask.any():
        idx = df.index[mask][0]
        for k, v in data.items():
            df.at[idx, k] = v
        st.session_state.df = df
        return 'updated'
    else:
        new_row = pd.DataFrame([{col: data.get(col, None) for col in ALL_COLUMNS}])
        st.session_state.df = pd.concat([df, new_row], ignore_index=True)
        return 'added'

def build_excel():
    df = st.session_state.df.copy()
    roster = st.session_state.roster

    # Ensure correct columns
    for col in ALL_COLUMNS:
        if col not in df.columns:
            df[col] = None
    df = df.reindex(columns=ALL_COLUMNS)

    wb = Workbook()

    # Hidden Roster sheet for dropdown source
    ws_roster = wb.active
    ws_roster.title = 'Roster'
    ws_roster['A1'] = 'Player Names'
    for i, name in enumerate(sorted(roster), start=2):
        ws_roster.cell(i, 1).value = name
    ws_roster.sheet_state = 'hidden'

    # Main sheet
    ws = wb.create_sheet('Alliance Stats')

    header_fill = PatternFill('solid', start_color='8B0000')
    header_font = Font(name='Arial', bold=True, color='FFFFFF', size=10)
    alt_fill = PatternFill('solid', start_color='FFF3E0')
    empty_fill = PatternFill('solid', start_color='F9F9F9')

    # Header row
    for ci, col_name in enumerate(ALL_COLUMNS, start=1):
        cell = ws.cell(1, ci, col_name)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    # Data rows — always write TOTAL_ROWS rows
    records = df.to_dict('records')
    for row_i in range(TOTAL_ROWS):
        excel_row = row_i + 2
        record = records[row_i] if row_i < len(records) else {}
        for ci, col_name in enumerate(ALL_COLUMNS, start=1):
            cell = ws.cell(excel_row, ci)
            cell.value = record.get(col_name, None)
            cell.font = Font(name='Arial', size=10)
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.fill = empty_fill if not record else (alt_fill if row_i % 2 == 0 else PatternFill())

    # Dropdown on Player Name column (A2:A101)
    if roster:
        dv = DataValidation(
            type="list",
            formula1=f"Roster!$A$2:$A${len(roster)+1}",
            allow_blank=True,
            showDropDown=False
        )
        dv.sqref = f"A2:A{TOTAL_ROWS+1}"
        ws.add_data_validation(dv)

    # Column widths
    ws.column_dimensions['A'].width = 22
    for ci, col_name in enumerate(ALL_COLUMNS, start=1):
        if ci > 1:
            ws.column_dimensions[get_column_letter(ci)].width = max(len(col_name) + 2, 12)

    ws.row_dimensions[1].height = 35
    ws.freeze_panes = 'B2'

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

def load_excel(file):
    xl = pd.ExcelFile(file)
    df = pd.read_excel(xl, sheet_name='Alliance Stats')
    df = df[df['Player Name'].notna() & (df['Player Name'].astype(str).str.strip() != '')]
    df = df.reset_index(drop=True)
    try:
        roster_df = pd.read_excel(xl, sheet_name='Roster', header=0)
        roster = roster_df.iloc[:, 0].dropna().astype(str).str.strip().tolist()
    except Exception:
        roster = df['Player Name'].dropna().astype(str).str.strip().tolist()
    return df, roster

# ════════════════════════════════════════
# SECTION 1 — Load
# ════════════════════════════════════════
st.subheader("1️⃣ Load Existing Tracker")
uploaded = st.file_uploader("Upload alliance_tracker.xlsx (skip if starting fresh)", type=["xlsx"], key="load")
if uploaded:
    df, roster = load_excel(uploaded)
    st.session_state.df = df
    st.session_state.roster = roster
    st.success(f"✅ Loaded {len(df)} players, {len(roster)} names in roster.")

players_with_data = st.session_state.df['Player Name'].dropna().tolist() if not st.session_state.df.empty and 'Player Name' in st.session_state.df.columns else []
st.caption(f"Tracker: **{len(players_with_data)} players** with data · **{len(st.session_state.roster)} names** in roster")

st.divider()

# ════════════════════════════════════════
# SECTION 2 — Roster management
# ════════════════════════════════════════
st.subheader("2️⃣ Manage Roster")

c1, c2 = st.columns(2)
with c1:
    new_name = st.text_input("➕ Add new player", placeholder="Type name then click Add")
    if st.button("Add to Roster"):
        name = new_name.strip()
        if not name:
            st.warning("Enter a name.")
        elif name in st.session_state.roster:
            st.warning(f"'{name}' already in roster.")
        else:
            st.session_state.roster.append(name)
            st.success(f"✅ '{name}' added to roster.")
            st.rerun()

with c2:
    if st.session_state.roster:
        to_kick = st.selectbox("🗑️ Kick player", options=["-- select --"] + sorted(st.session_state.roster))
        if st.button("Remove & Delete Data", type="primary"):
            if to_kick == "-- select --":
                st.warning("Select a player to remove.")
            else:
                st.session_state.roster.remove(to_kick)
                if not st.session_state.df.empty and 'Player Name' in st.session_state.df.columns:
                    st.session_state.df = st.session_state.df[
                        st.session_state.df['Player Name'].astype(str).str.strip() != to_kick
                    ].reset_index(drop=True)
                st.success(f"✅ '{to_kick}' removed.")
                st.rerun()

if st.session_state.roster:
    with st.expander(f"View full roster ({len(st.session_state.roster)} players)"):
        cols = st.columns(3)
        for i, n in enumerate(sorted(st.session_state.roster)):
            cols[i % 3].write(f"• {n}")

st.divider()

# ════════════════════════════════════════
# SECTION 3 — Upload & Extract
# ════════════════════════════════════════
st.subheader("3️⃣ Upload Screenshots & Extract Stats")

# Player selection
all_roster_names = sorted(st.session_state.roster)
if all_roster_names:
    selected = st.selectbox(
        "Select player",
        options=["-- select player --"] + all_roster_names,
        key="player_select"
    )
    player_name = selected if selected != "-- select player --" else ""
else:
    st.info("Add players to the roster first (Step 2).")
    player_name = ""

screenshots = st.file_uploader(
    "Upload all screenshots for this player",
    type=["png","jpg","jpeg","webp"],
    accept_multiple_files=True,
    key=f"shots_{st.session_state.upload_key}"
)

col_a, col_b = st.columns([2,1])
with col_a:
    do_extract = st.button("🔍 Extract Stats", disabled=not (screenshots and player_name))
with col_b:
    if st.button("🗑️ Clear"):
        st.session_state.upload_key += 1
        st.session_state.extracted = None
        st.rerun()

if do_extract and screenshots and player_name:
    all_text = ""
    bar = st.progress(0)
    msg = st.empty()
    for i, f in enumerate(screenshots):
        msg.write(f"Reading {i+1}/{len(screenshots)}...")
        all_text += "\n" + ocr_image(Image.open(f))
        bar.progress((i+1) / len(screenshots))
    msg.empty()
    bar.empty()
    st.session_state.extracted = extract_all(all_text)
    st.session_state.extracted['_player'] = player_name
    st.success(f"✅ Done! {len(st.session_state.extracted)-1} fields extracted for **{player_name}**")

    stats_found = {k: st.session_state.extracted[k] for k in KNOWN_STATS if k in st.session_state.extracted}
    build_keys = ['March Size','Elder Titan Tier','Titan Talent Level','Beast Tier',
                  'Beast Talent Level','Beast Skill Level','Totem Level',
                  'Special Stats Level','Jewels Level','Zodiac Green','Zodiac White',
                  'Northern Green','Colossus Level','Emblem Level']
    build_found = {k: st.session_state.extracted[k] for k in build_keys if k in st.session_state.extracted}

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**🏰 Build**")
        st.dataframe(pd.DataFrame(list(build_found.items()), columns=['Field','Value']), use_container_width=True)
    st.markdown("**📊 Stats**")
    if stats_found:
        items = list(stats_found.items())
        half = len(items)//2 + len(items)%2
        ca, cb = st.columns(2)
        with ca:
            st.dataframe(pd.DataFrame(items[:half], columns=['Stat','Value']), use_container_width=True)
        with cb:
            st.dataframe(pd.DataFrame(items[half:], columns=['Stat','Value']), use_container_width=True)
        st.info(f"📊 {len(stats_found)}/47 stats found")

    with st.expander("Raw OCR"):
        st.text_area("", all_text, height=150)

# ── Save button ──
if st.session_state.extracted and '_player' in st.session_state.extracted:
    pname = st.session_state.extracted['_player']
    is_update = pname in players_with_data
    st.divider()
    label = f"🔄 Update {pname}" if is_update else f"➕ Save {pname} to Tracker"
    if is_update:
        st.warning(f"This will **replace** all existing data for **{pname}**.")

    if st.button(label, type="primary", key="save_btn"):
        data_to_save = {k: v for k, v in st.session_state.extracted.items() if k != '_player'}
        action = save_player(pname, data_to_save)
        st.session_state.extracted = None
        st.session_state.upload_key += 1
        count = len(st.session_state.df)
        verb = "Updated" if action == "updated" else "Added"
        st.success(f"✅ {verb} **{pname}**! Tracker now has **{count}** players. Select next player above ⬆️")
        st.rerun()

st.divider()

# ════════════════════════════════════════
# SECTION 4 — Download
# ════════════════════════════════════════
st.subheader("4️⃣ Download Tracker")

count = len(st.session_state.df)
st.write(f"**{count}** players with data · **{len(st.session_state.roster)}** in roster · **{TOTAL_ROWS}** rows in Excel")

if count > 0:
    with st.expander("👁️ Preview"):
        st.dataframe(st.session_state.df, use_container_width=True)

excel_buf = build_excel()
st.download_button(
    label="⬇️ Download Alliance Tracker Excel",
    data=excel_buf,
    file_name="alliance_tracker.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
st.caption("💡 Save to Google Drive. Next session upload it here to continue where you left off.")

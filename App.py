import streamlit as st
import pandas as pd
import pytesseract
from PIL import Image
import io
import re
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Alliance Battle Tracker", page_icon="⚔️", layout="centered")

st.title("⚔️ Alliance Battle Report Tracker")
st.write("Upload all screenshots for one player, extract their stats, and add them to your Alliance Excel tracker.")

# ── Known stat names from the game (used to validate OCR lines) ──
STATS_BONUS_KEYS = [
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
    "Golem Defense", "Golem HP",
]

def ocr_image(img):
    return pytesseract.image_to_string(img, config='--psm 6')

def extract_player_info(all_text):
    info = {}
    # Player name
    m = re.search(r'\^\^(.+?)\^\^', all_text)
    if m:
        info['Player Name'] = m.group(1).strip()
    # Result
    m = re.search(r'\b(Victory|Defeat)\b', all_text, re.I)
    if m:
        info['Result'] = m.group(1)
    # Power Loss
    m = re.search(r'Power Loss\s+([-\d,]+)', all_text)
    if m:
        info['Power Loss'] = m.group(1).replace(',', '')
    # Battle stats
    for stat in ['Total Army', 'Kills', 'Losses', 'Wounded', 'Survivors', 'Battlers']:
        m = re.search(rf'{stat}\s+([\d,]+)', all_text)
        if m:
            info[stat] = m.group(1).replace(',', '')
    return info

def extract_stats_bonus(all_text):
    stats = {}
    lines = all_text.splitlines()
    for line in lines:
        line = line.strip()
        # Match: "Stat Name   123.4%" or "Stat Name   86"
        m = re.match(r'^([A-Za-z][A-Za-z\s]+?)\s{2,}([\d,]+\.?\d*%?)$', line)
        if m:
            key = m.group(1).strip()
            val = m.group(2).strip()
            # Only keep if it looks like a real stat (not section headers)
            if len(key) > 3 and not re.search(r'battle report|attacker|defender|stats bonus|artifact|overall|troops info|training|totem|beast|equip|palace|garden|colossus|emblem|deployed|skill|evolution|amount|level|star|green|white|northern|zodiac', key, re.I):
                stats[key] = val
    return stats

def extract_other_info(all_text):
    other = {}
    # Elder Titan
    m = re.search(r'Evolution:\s*(.+)', all_text)
    if m:
        other['Elder Titan Evolution'] = m.group(1).strip()
    m = re.search(r'Total Talent Level:\s*(\d+)', all_text)
    if m:
        other['Total Talent Level'] = m.group(1)
    # Beast
    m = re.search(r'Beast.*?Evolution:\s*Tier\s*(\d+)', all_text, re.S)
    if m:
        other['Beast Tier'] = m.group(1)
    # Totem
    m = re.search(r'Guard Totem.*?Level:\s*(Lv\.\d+)', all_text, re.S)
    if m:
        other['Totem Level'] = m.group(1)
    # Equip
    m = re.search(r'Total Special Stats Level:\s*(\d+)', all_text)
    if m:
        other['Special Stats Level'] = m.group(1)
    m = re.search(r'Total Jewels Level:\s*(\d+)', all_text)
    if m:
        other['Jewels Level'] = m.group(1)
    # Colossus
    m = re.search(r'Colossus.*?Total Level:\s*(\d+)', all_text, re.S)
    if m:
        other['Colossus Level'] = m.group(1)
    # Emblem
    m = re.search(r'Emblem.*?Total Level:\s*(\d+)', all_text, re.S)
    if m:
        other['Emblem Level'] = m.group(1)
    return other

def build_player_row(player_info, stats_bonus, other_info):
    row = {}
    row.update(player_info)
    row.update(stats_bonus)
    row.update(other_info)
    return row

def save_to_excel(existing_df, new_row):
    updated_df = pd.concat([existing_df, pd.DataFrame([new_row])], ignore_index=True)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        updated_df.to_excel(writer, index=False, sheet_name='Alliance Stats')
        ws = writer.sheets['Alliance Stats']

        # Style header row
        header_fill = PatternFill('solid', start_color='8B0000')
        header_font = Font(name='Arial', bold=True, color='FFFFFF', size=10)
        center = Alignment(horizontal='center', vertical='center')
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center

        # Alternate row colors
        light = PatternFill('solid', start_color='FFF8F0')
        lighter = PatternFill('solid', start_color='FFFFFF')
        for i, row in enumerate(ws.iter_rows(min_row=2, max_row=ws.max_row), start=2):
            for cell in row:
                cell.fill = light if i % 2 == 0 else lighter
                cell.font = Font(name='Arial', size=10)
                cell.alignment = Alignment(horizontal='center', vertical='center')

        # Auto column widths
        for col in ws.columns:
            max_len = max((len(str(cell.value)) if cell.value else 0) for cell in col)
            ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 4, 30)

        ws.freeze_panes = 'B2'

    buf.seek(0)
    return updated_df, buf

# ── Session state for accumulating players ──
if 'alliance_df' not in st.session_state:
    st.session_state.alliance_df = pd.DataFrame()
if 'last_player' not in st.session_state:
    st.session_state.last_player = None

# ── Upload existing tracker (optional) ──
st.subheader("1️⃣ Load Existing Alliance Tracker (optional)")
existing_file = st.file_uploader("Upload your existing Alliance Excel to add more players", type=["xlsx"], key="existing")
if existing_file:
    st.session_state.alliance_df = pd.read_excel(existing_file)
    st.success(f"Loaded {len(st.session_state.alliance_df)} existing players.")

# ── Upload player screenshots ──
st.subheader("2️⃣ Upload All Screenshots for One Player")
uploaded_files = st.file_uploader(
    "Upload all screenshots for this player (select multiple at once)",
    type=["png", "jpg", "jpeg", "webp"],
    accept_multiple_files=True,
    key="screenshots"
)

if uploaded_files:
    st.write(f"📸 {len(uploaded_files)} screenshots uploaded")

    if st.button("🔍 Extract Stats from Screenshots"):
        all_text = ""
        progress = st.progress(0)
        for i, f in enumerate(uploaded_files):
            img = Image.open(f)
            all_text += "\n" + ocr_image(img)
            progress.progress((i + 1) / len(uploaded_files))

        player_info = extract_player_info(all_text)
        stats_bonus = extract_stats_bonus(all_text)
        other_info = extract_other_info(all_text)
        player_row = build_player_row(player_info, stats_bonus, other_info)

        st.session_state.last_player = player_row

        st.subheader("✅ Extracted Data")

        col1, col2 = st.columns(2)
        with col1:
            st.markdown("**Player Info**")
            st.dataframe(pd.DataFrame(list(player_info.items()), columns=["Field", "Value"]), use_container_width=True)
        with col2:
            st.markdown("**Other Info**")
            st.dataframe(pd.DataFrame(list(other_info.items()), columns=["Field", "Value"]), use_container_width=True)

        st.markdown("**Stats Bonus**")
        st.dataframe(pd.DataFrame(list(stats_bonus.items()), columns=["Stat", "Value"]), use_container_width=True)

        with st.expander("🔤 Raw OCR Text (for debugging)"):
            st.text_area("", all_text, height=200)

# ── Add to alliance tracker ──
if st.session_state.last_player:
    st.subheader("3️⃣ Add to Alliance Tracker")
    player_name = st.session_state.last_player.get('Player Name', 'Unknown')
    st.write(f"Ready to add: **{player_name}**")

    if st.button(f"➕ Add {player_name} to Alliance Tracker"):
        st.session_state.alliance_df, excel_buf = save_to_excel(
            st.session_state.alliance_df,
            st.session_state.last_player
        )
        st.session_state.last_player = None
        st.success(f"✅ {player_name} added! Total players: {len(st.session_state.alliance_df)}")
        st.download_button(
            label="⬇️ Download Alliance Tracker Excel",
            data=excel_buf,
            file_name="alliance_tracker.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# ── View & download current tracker ──
if not st.session_state.alliance_df.empty:
    st.subheader(f"📊 Alliance Tracker ({len(st.session_state.alliance_df)} players)")
    st.dataframe(st.session_state.alliance_df, use_container_width=True)

    _, dl_buf = save_to_excel(st.session_state.alliance_df, {})
    # Remove the empty row we just added
    temp_df = st.session_state.alliance_df.copy()
    buf2 = io.BytesIO()
    with pd.ExcelWriter(buf2, engine='openpyxl') as writer:
        temp_df.to_excel(writer, index=False, sheet_name='Alliance Stats')
    buf2.seek(0)

    st.download_button(
        label="⬇️ Download Full Alliance Tracker",
        data=buf2,
        file_name="alliance_tracker.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

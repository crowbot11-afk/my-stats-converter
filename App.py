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

def upsert_player(df, new_data):
    player_name = new_data.get('Player Name', '')
    if df.empty or 'Player Name' not in df.columns:
        return pd.concat([df, pd.DataFrame([new_data])], ignore_index=True), 'added'
    match_idx = df.index[
        df['Player Name'].astype(str).str.strip().str.lower() == player_name.strip().lower()
    ].tolist()
    if match_idx:
        for key, val in new_data.items():
            df.at[match_idx[0], key] = val
        return df, 'updated'
    else:
        return pd.concat([df, pd.DataFrame([new_data])], ignore_index=True), 'added'

def build_excel(df, roster):
    """
    Build Excel with:
    - 'Alliance Stats' sheet: 100 fixed rows, Player Name column has dropdown
    - 'Roster' sheet: hidden name list used by the dropdown
    """
    wb = Workbook()

    # ── Roster sheet (name list for dropdown) ──
    ws_roster = wb.active
    ws_roster.title = 'Roster'
    ws_roster['A1'] = 'Player Names'
    ws_roster['A1'].font = Font(bold=True)
    for i, name in enumerate(sorted(roster), start=2):
        ws_roster.cell(i, 1).value = name
    ws_roster.sheet_state = 'hidden'

    # ── Alliance Stats sheet ──
    ws = wb.create_sheet('Alliance Stats')

    # Ensure all columns exist in df
    for col in ALL_COLUMNS:
        if col not in df.columns:
            df[col] = None

    # Reorder columns
    df = df.reindex(columns=ALL_COLUMNS)

    # Write header
    header_fill = PatternFill('solid', start_color='8B0000')
    header_font = Font(name='Arial', bold=True, color='FFFFFF', size=10)
    center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    alt_fill = PatternFill('solid', start_color='FFF3E0')
    empty_fill = PatternFill('solid', start_color='F5F5F5')

    for col_idx, col_name in enumerate(ALL_COLUMNS, start=1):
        cell = ws.cell(1, col_idx, col_name)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center

    # Write data rows (players)
    player_rows = df.to_dict('records')
    for row_idx in range(TOTAL_ROWS):
        excel_row = row_idx + 2
        if row_idx < len(player_rows):
            row_data = player_rows[row_idx]
            is_empty = False
        else:
            row_data = {}
            is_empty = True

        for col_idx, col_name in enumerate(ALL_COLUMNS, start=1):
            cell = ws.cell(excel_row, col_idx)
            cell.value = row_data.get(col_name, None) if not is_empty else None
            cell.font = Font(name='Arial', size=10)
            cell.alignment = Alignment(horizontal='center', vertical='center')
            if is_empty:
                cell.fill = empty_fill
            elif row_idx % 2 == 0:
                cell.fill = alt_fill

    # Add dropdown validation to Player Name column (col A, rows 2-101)
    roster_sheet_ref = f"Roster!$A$2:$A${len(roster)+1}"
    dv = DataValidation(
        type="list",
        formula1=roster_sheet_ref,
        allow_blank=True,
        showDropDown=False  # False = show the dropdown arrow
    )
    dv.sqref = f"A2:A{TOTAL_ROWS+1}"
    ws.add_data_validation(dv)

    # Column widths
    ws.column_dimensions['A'].width = 22  # Player Name
    for col_idx, col_name in enumerate(ALL_COLUMNS, start=1):
        if col_idx == 1:
            continue
        ws.column_dimensions[get_column_letter(col_idx)].width = max(len(col_name) + 2, 12)

    ws.row_dimensions[1].height = 35
    ws.freeze_panes = 'B2'

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

def load_tracker(file):
    """Load both the data and roster from an uploaded Excel."""
    xl = pd.ExcelFile(file)
    df = pd.read_excel(xl, sheet_name='Alliance Stats')
    df = df[df['Player Name'].notna() & (df['Player Name'].astype(str).str.strip() != '')]
    try:
        roster_df = pd.read_excel(xl, sheet_name='Roster', header=0)
        roster = roster_df.iloc[:, 0].dropna().astype(str).tolist()
    except Exception:
        roster = df['Player Name'].dropna().astype(str).tolist()
    return df, roster

# ── Session state ──
for key, default in [('alliance_df', pd.DataFrame()), ('roster', []),
                      ('last_data', None), ('upload_key', 0)]:
    if key not in st.session_state:
        st.session_state[key] = default

# ══════════════════════════════════════════════════
# STEP 1 — Load tracker
# ══════════════════════════════════════════════════
st.subheader("1️⃣ Load Your Alliance Tracker")
existing_file = st.file_uploader("Upload alliance_tracker.xlsx", type=["xlsx"], key="existing")
if existing_file:
    df, roster = load_tracker(existing_file)
    st.session_state.alliance_df = df
    st.session_state.roster = roster
    st.success(f"✅ Loaded {len(df)} players, {len(roster)} names in roster.")

st.divider()

# ══════════════════════════════════════════════════
# STEP 2 — Manage roster
# ══════════════════════════════════════════════════
st.subheader("2️⃣ Manage Alliance Roster")
st.caption("This is the master name list. Add new recruits, remove kicked players.")

col_add, col_remove = st.columns(2)

with col_add:
    st.markdown("**➕ Add player to roster**")
    new_name = st.text_input("New player name", placeholder="e.g. Fighterrulez", key="new_name")
    if st.button("Add to Roster"):
        if new_name.strip():
            if new_name.strip() in st.session_state.roster:
                st.warning(f"'{new_name}' is already in the roster.")
            else:
                st.session_state.roster.append(new_name.strip())
                st.success(f"✅ '{new_name}' added to roster.")
                st.rerun()
        else:
            st.warning("Enter a name first.")

with col_remove:
    st.markdown("**🗑️ Remove (kick) player**")
    if st.session_state.roster:
        to_remove = st.selectbox("Select player to remove", options=sorted(st.session_state.roster), key="remove_sel")
        if st.button("Remove from Roster & Delete Row", type="primary"):
            st.session_state.roster.remove(to_remove)
            if not st.session_state.alliance_df.empty and 'Player Name' in st.session_state.alliance_df.columns:
                st.session_state.alliance_df = st.session_state.alliance_df[
                    st.session_state.alliance_df['Player Name'].astype(str).str.strip() != to_remove
                ].reset_index(drop=True)
            st.success(f"✅ '{to_remove}' removed from roster and data deleted.")
            st.rerun()
    else:
        st.info("Roster is empty. Add players above.")

if st.session_state.roster:
    with st.expander(f"📋 Full Roster ({len(st.session_state.roster)} players)"):
        cols = st.columns(3)
        for i, name in enumerate(sorted(st.session_state.roster)):
            cols[i % 3].write(f"• {name}")

st.divider()

# ══════════════════════════════════════════════════
# STEP 3 — Update player stats
# ══════════════════════════════════════════════════
st.subheader("3️⃣ Update Player Stats")

if not st.session_state.roster:
    st.info("Add players to the roster above before updating stats.")
else:
    mode = st.radio("Player type:", ["Existing player", "New player (already added to roster)"], horizontal=True)

    if mode == "Existing player":
        existing_names = sorted(st.session_state.alliance_df['Player Name'].dropna().tolist()) \
            if not st.session_state.alliance_df.empty and 'Player Name' in st.session_state.alliance_df.columns else []
        if existing_names:
            selected_player = st.selectbox("Select player to update", options=existing_names)
            player_name_input = selected_player
            current = st.session_state.alliance_df[
                st.session_state.alliance_df['Player Name'] == selected_player
            ].iloc[0].dropna()
            with st.expander(f"📋 Current data for {selected_player}"):
                st.dataframe(pd.DataFrame(current).reset_index().rename(
                    columns={'index':'Field', 0:'Value'}), use_container_width=True)
        else:
            st.info("No players with data yet. Upload screenshots for a roster member.")
            player_name_input = ""
    else:
        roster_without_data = sorted([
            n for n in st.session_state.roster
            if st.session_state.alliance_df.empty
            or 'Player Name' not in st.session_state.alliance_df.columns
            or n not in st.session_state.alliance_df['Player Name'].tolist()
        ])
        if roster_without_data:
            player_name_input = st.selectbox("Select new player", options=roster_without_data)
        else:
            st.info("All roster players already have data. Use 'Existing player' to update.")
            player_name_input = ""

    if player_name_input:
        uploaded_files = st.file_uploader(
            f"Upload all screenshots for {player_name_input}",
            type=["png", "jpg", "jpeg", "webp"],
            accept_multiple_files=True,
            key=f"shots_{st.session_state.upload_key}"
        )

        col_ex, col_cl = st.columns([2, 1])
        with col_ex:
            extract_btn = st.button("🔍 Extract Stats", disabled=not uploaded_files)
        with col_cl:
            if st.button("🗑️ Clear Screenshots"):
                st.session_state.upload_key += 1
                st.session_state.last_data = None
                st.rerun()

        if extract_btn and uploaded_files:
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
            st.success(f"✅ Extracted {len(extracted)-1} fields for **{player_name_input}**")

            stats_data = {k: extracted[k] for k in KNOWN_STATS if k in extracted}
            build_keys = ['March Size','Elder Titan Tier','Titan Talent Level','Beast Tier',
                          'Beast Talent Level','Beast Skill Level','Totem Level',
                          'Special Stats Level','Jewels Level','Zodiac Green','Zodiac White',
                          'Northern Green','Colossus Level','Emblem Level']
            col1, col2 = st.columns(2)
            with col1:
                st.markdown("**🏰 Build Info**")
                build_data = {k: extracted.get(k,'—') for k in build_keys if k in extracted}
                st.dataframe(pd.DataFrame(list(build_data.items()), columns=['Field','Value']), use_container_width=True)
            st.markdown("**📊 Stats Bonus**")
            if stats_data:
                items = list(stats_data.items())
                half = len(items)//2 + len(items)%2
                c1, c2 = st.columns(2)
                with c1:
                    st.dataframe(pd.DataFrame(items[:half], columns=['Stat','Value']), use_container_width=True)
                with c2:
                    st.dataframe(pd.DataFrame(items[half:], columns=['Stat','Value']), use_container_width=True)
                st.info(f"📊 {len(stats_data)}/47 stats extracted")
            with st.expander("🔤 Raw OCR Text"):
                st.text_area("", all_text, height=200)

    if st.session_state.last_data:
        pname = st.session_state.last_data.get('Player Name', '')
        is_update = (not st.session_state.alliance_df.empty and
                     'Player Name' in st.session_state.alliance_df.columns and
                     pname in st.session_state.alliance_df['Player Name'].tolist())
        st.divider()
        label = f"🔄 Update {pname}" if is_update else f"➕ Save {pname} to Tracker"
        if is_update:
            st.warning(f"This will replace all existing data for **{pname}**.")
        if st.button(label, type="primary"):
            st.session_state.alliance_df, status = upsert_player(
                st.session_state.alliance_df, st.session_state.last_data)
            st.session_state.last_data = None
            st.session_state.upload_key += 1
            action = "Updated" if status == "updated" else "Saved"
            st.success(f"✅ {action} {pname}! Select next player above ⬆️")
            st.rerun()

st.divider()

# ══════════════════════════════════════════════════
# STEP 4 — Download
# ══════════════════════════════════════════════════
st.subheader("4️⃣ Download Alliance Tracker")

if not st.session_state.roster and st.session_state.alliance_df.empty:
    st.info("Add players to the roster and upload their stats to generate the Excel.")
else:
    total_players = len(st.session_state.alliance_df) if not st.session_state.alliance_df.empty else 0
    st.write(f"**{len(st.session_state.roster)}** names in roster · **{total_players}** players with data · **{TOTAL_ROWS}** total rows in Excel")

    if not st.session_state.alliance_df.empty:
        with st.expander("👁️ Preview tracker"):
            st.dataframe(st.session_state.alliance_df, use_container_width=True)

    excel_buf = build_excel(
        st.session_state.alliance_df if not st.session_state.alliance_df.empty else pd.DataFrame(),
        st.session_state.roster
    )
    st.download_button(
        label="⬇️ Download Alliance Tracker Excel",
        data=excel_buf,
        file_name="alliance_tracker.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    st.caption("💡 The Player Name column in Excel has a dropdown with all roster names. Empty rows are ready for future members.")

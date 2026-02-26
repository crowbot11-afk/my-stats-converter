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

st.set_page_config(page_title="Alliance Tracker", page_icon="⚔️", layout="centered")
st.title("⚔️ Alliance Tracker")

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

# ── Init ALL session state keys upfront ──
defaults = {
    'df': pd.DataFrame(columns=ALL_COLUMNS),
    'roster': [],
    'extracted': None,
    'upload_key': 0,
    'add_msg': None,
    'kick_msg': None,
    'save_msg': None,
    'picked_player': '',   # single source of truth for selected player in tab2
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

# Normalize roster every run
st.session_state.roster = sorted(set(
    s.strip() for s in st.session_state.roster if str(s).strip()
))

# ── OCR / extract ──
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
    data['Player Name'] = name
    df = st.session_state.df
    mask = df['Player Name'].astype(str).str.strip().str.lower() == name.strip().lower()
    if mask.any():
        idx = df.index[mask][0]
        for k, v in data.items():
            df.at[idx, k] = v
        st.session_state.df = df
        return 'updated'
    else:
        new_row = {col: data.get(col, None) for col in ALL_COLUMNS}
        st.session_state.df = pd.concat(
            [st.session_state.df, pd.DataFrame([new_row])], ignore_index=True
        )
        return 'added'

def build_excel():
    df = st.session_state.df.copy()
    roster = sorted(st.session_state.roster)
    for col in ALL_COLUMNS:
        if col not in df.columns:
            df[col] = None
    df = df.reindex(columns=ALL_COLUMNS)
    wb = Workbook()
    ws_r = wb.active
    ws_r.title = 'Roster'
    ws_r['A1'] = 'Player Names'
    for i, name in enumerate(roster, start=2):
        ws_r.cell(i, 1).value = name
    ws_r.sheet_state = 'hidden'
    ws = wb.create_sheet('Alliance Stats')
    h_fill = PatternFill('solid', start_color='8B0000')
    h_font = Font(name='Arial', bold=True, color='FFFFFF', size=10)
    for ci, col in enumerate(ALL_COLUMNS, 1):
        c = ws.cell(1, ci, col)
        c.fill = h_fill
        c.font = h_font
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    alt = PatternFill('solid', start_color='FFF3E0')
    empty_fill = PatternFill('solid', start_color='F9F9F9')
    records = df.to_dict('records')
    for ri in range(TOTAL_ROWS):
        er = ri + 2
        rec = records[ri] if ri < len(records) else {}
        has_data = bool(rec)
        for ci, col in enumerate(ALL_COLUMNS, 1):
            c = ws.cell(er, ci)
            c.value = rec.get(col) if has_data else None
            c.font = Font(name='Arial', size=10)
            c.alignment = Alignment(horizontal='center', vertical='center')
            c.fill = (alt if ri % 2 == 0 else PatternFill()) if has_data else empty_fill
    if roster:
        dv = DataValidation(type="list",
            formula1=f"Roster!$A$2:$A${len(roster)+1}",
            allow_blank=True, showDropDown=False)
        dv.sqref = f"A2:A{TOTAL_ROWS+1}"
        ws.add_data_validation(dv)
    ws.column_dimensions['A'].width = 22
    for ci, col in enumerate(ALL_COLUMNS, 1):
        if ci > 1:
            ws.column_dimensions[get_column_letter(ci)].width = max(len(col)+2, 12)
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
        rdf = pd.read_excel(xl, sheet_name='Roster', header=0)
        roster = rdf.iloc[:,0].dropna().astype(str).str.strip().tolist()
    except Exception:
        roster = df['Player Name'].dropna().astype(str).str.strip().tolist()
    return df, roster

# ══════════════════════════════════════════
# TABS
# ══════════════════════════════════════════
tab1, tab2, tab3 = st.tabs(["📁 Load & Roster", "📸 Add / Update Player", "⬇️ Download"])

# ─────────────────────────────────────────
# TAB 1
# ─────────────────────────────────────────
with tab1:
    st.subheader("Load Existing Tracker")
    uploaded = st.file_uploader("Upload alliance_tracker.xlsx", type=["xlsx"], key="load")
    if uploaded:
        df, roster = load_excel(uploaded)
        st.session_state.df = df
        st.session_state.roster = roster
        st.session_state.picked_player = ''
        st.success(f"✅ Loaded {len(df)} players, {len(roster)} in roster.")

    st.info(f"**{len(st.session_state.df)}** players with data · **{len(st.session_state.roster)}** in roster")

    st.divider()
    st.subheader("Add New Player to Roster")

    if st.session_state.add_msg:
        lvl, txt = st.session_state.add_msg
        st.session_state.add_msg = None
        if lvl == 'success':
            st.success(txt)
        else:
            st.warning(txt)

    with st.form("form_add", clear_on_submit=True):
        new_name = st.text_input("Player name", placeholder="Type name here")
        if st.form_submit_button("➕ Add to Roster"):
            name = new_name.strip()
            if not name:
                st.session_state.add_msg = ('warning', 'Type a name first.')
            elif name.lower() in [n.lower() for n in st.session_state.roster]:
                st.session_state.add_msg = ('warning', f"'{name}' is already in the roster.")
            else:
                st.session_state.roster = sorted(set(
                    st.session_state.roster + [name]
                ))
                st.session_state.add_msg = ('success', f"✅ '{name}' added to roster!")

    st.divider()
    st.subheader("Remove (Kick) Player")

    if st.session_state.kick_msg:
        lvl, txt = st.session_state.kick_msg
        st.session_state.kick_msg = None
        if lvl == 'success':
            st.success(txt)
        else:
            st.warning(txt)

    if st.session_state.roster:
        with st.form("form_kick", clear_on_submit=True):
            to_kick = st.selectbox("Select player to remove", ["-- select --"] + st.session_state.roster)
            if st.form_submit_button("🗑️ Remove & Delete Their Data"):
                if to_kick == "-- select --":
                    st.session_state.kick_msg = ('warning', 'Select a player first.')
                else:
                    st.session_state.roster = [n for n in st.session_state.roster
                                               if n.lower() != to_kick.lower()]
                    if 'Player Name' in st.session_state.df.columns:
                        st.session_state.df = st.session_state.df[
                            st.session_state.df['Player Name'].astype(str).str.strip().str.lower() != to_kick.lower()
                        ].reset_index(drop=True)
                    if st.session_state.picked_player.lower() == to_kick.lower():
                        st.session_state.picked_player = ''
                    st.session_state.kick_msg = ('success', f"✅ '{to_kick}' removed.")
    else:
        st.info("No players in roster yet.")

    if st.session_state.roster:
        st.divider()
        st.subheader(f"Full Roster ({len(st.session_state.roster)} players)")
        cols = st.columns(2)
        for i, name in enumerate(st.session_state.roster):
            cols[i % 2].write(f"• {name}")

# ─────────────────────────────────────────
# TAB 2
# ─────────────────────────────────────────
with tab2:
    roster = st.session_state.roster

    if not roster:
        st.info("Go to **Load & Roster** tab and add players first.")
    else:
        if st.session_state.save_msg:
            lvl, txt = st.session_state.save_msg
            st.session_state.save_msg = None
            if lvl == 'success':
                st.success(txt)
            else:
                st.warning(txt)

        st.subheader("Select Player")

        # Ensure picked_player is still valid (e.g. after kick)
        if st.session_state.picked_player not in roster:
            st.session_state.picked_player = ''

        # Use key= so radio remembers selection across reruns without index=
        # Pre-set the key only if it's missing or invalid
        if 'radio_player' not in st.session_state or st.session_state.radio_player not in (['-- select player --'] + roster):
            st.session_state.radio_player = st.session_state.picked_player if st.session_state.picked_player in roster else '-- select player --'

        st.radio(
            "Player",
            options=['-- select player --'] + roster,
            key='radio_player',
            label_visibility="collapsed"
        )

        # Persist the selection — radio key holds the value reliably
        chosen = st.session_state.radio_player
        player_name = chosen if chosen != '-- select player --' else ''
        st.session_state.picked_player = player_name

        if player_name:
            existing = st.session_state.df['Player Name'].astype(str).str.strip().tolist() if not st.session_state.df.empty else []
            if player_name in existing:
                st.warning(f"⚠️ {player_name} already has data — saving will **replace** it.")
            else:
                st.info(f"Selected: **{player_name}**")

        st.divider()
        st.subheader("Upload Screenshots")
        screenshots = st.file_uploader(
            "Select all screenshots at once",
            type=["png", "jpg", "jpeg", "webp"],
            accept_multiple_files=True,
            key=f"shots_{st.session_state.upload_key}"
        )

        col_a, col_b = st.columns([2, 1])
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
            slot = st.empty()
            for i, f in enumerate(screenshots):
                slot.write(f"Reading {i+1}/{len(screenshots)}...")
                all_text += "\n" + ocr_image(Image.open(f))
                bar.progress((i+1)/len(screenshots))
            slot.empty(); bar.empty()
            st.session_state.extracted = extract_all(all_text)
            st.session_state.extracted['_player'] = player_name
            n = len(st.session_state.extracted) - 1
            st.success(f"✅ {n} fields extracted for **{player_name}**")

            stats_found = {k: st.session_state.extracted[k] for k in KNOWN_STATS if k in st.session_state.extracted}
            build_keys = ['March Size','Elder Titan Tier','Titan Talent Level','Beast Tier',
                          'Beast Talent Level','Beast Skill Level','Totem Level',
                          'Special Stats Level','Jewels Level','Zodiac Green','Zodiac White',
                          'Northern Green','Colossus Level','Emblem Level']
            build_found = {k: st.session_state.extracted[k] for k in build_keys if k in st.session_state.extracted}
            st.markdown("**🏰 Build Info**")
            st.dataframe(pd.DataFrame(list(build_found.items()), columns=['Field','Value']), use_container_width=True)
            st.markdown("**📊 Stats Bonus**")
            if stats_found:
                items = list(stats_found.items())
                half = len(items)//2 + len(items)%2
                ca, cb = st.columns(2)
                with ca:
                    st.dataframe(pd.DataFrame(items[:half], columns=['Stat','Value']), use_container_width=True)
                with cb:
                    st.dataframe(pd.DataFrame(items[half:], columns=['Stat','Value']), use_container_width=True)
                st.info(f"📊 {len(stats_found)}/47 stats")
            with st.expander("Raw OCR"):
                st.text_area("", all_text, height=150)

        if st.session_state.extracted and st.session_state.extracted.get('_player') == player_name and player_name:
            st.divider()
            existing = st.session_state.df['Player Name'].astype(str).str.strip().tolist() if not st.session_state.df.empty else []
            label = f"🔄 Update {player_name}" if player_name in existing else f"➕ Save {player_name}"
            if st.button(label, type="primary"):
                data_to_save = {k: v for k, v in st.session_state.extracted.items() if k != '_player'}
                action = save_player(player_name, data_to_save)
                st.session_state.extracted = None
                st.session_state.upload_key += 1
                verb = "Updated" if action == "updated" else "Saved"
                st.session_state.save_msg = ('success', f"✅ {verb} **{player_name}**! {len(st.session_state.df)} players total.")
                st.rerun()

# ─────────────────────────────────────────
# TAB 3
# ─────────────────────────────────────────
with tab3:
    st.subheader("Download Alliance Tracker")
    st.write(f"**{len(st.session_state.df)}** players with data · **{len(st.session_state.roster)}** in roster · **{TOTAL_ROWS}** rows in Excel")
    if len(st.session_state.df) > 0:
        with st.expander("👁️ Preview data"):
            st.dataframe(st.session_state.df, use_container_width=True)
    excel_buf = build_excel()
    st.download_button(
        label="⬇️ Download alliance_tracker.xlsx",
        data=excel_buf,
        file_name="alliance_tracker.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    st.caption("Save to Google Drive. Next session upload it here to continue.")

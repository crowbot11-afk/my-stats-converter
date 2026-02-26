import streamlit as st
import pandas as pd
import pytesseract
from PIL import Image
import io, os, re, json
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from datetime import date

st.set_page_config(
    page_title="Event Tracker",
    page_icon="swords",
    layout="centered"
)
st.title("Event Participation Tracker")

DATA_FILE = "/tmp/event_tracker.json"


def persist_events():
    with open(DATA_FILE, "w") as f:
        json.dump(st.session_state.event_results, f)


def load_events():
    if os.path.exists(DATA_FILE):
        try:
            with open(DATA_FILE, "r") as f:
                return json.load(f)
        except Exception:
            pass
    return {}


if "ev_loaded" not in st.session_state:
    st.session_state.event_results = load_events()
    st.session_state.roster = []
    st.session_state.ev_upload_key = 0
    st.session_state.ev_last_xlsx = None
    st.session_state.ev_loaded = True


def ocr_image(img):
    return pytesseract.image_to_string(img, config="--psm 6")


def extract_participants(all_text):
    # Parse OCR text from event leaderboard screenshots
    results = {}
    lines = all_text.splitlines()
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        if not line or len(line) < 2:
            i += 1
            continue
        # Skip header/UI lines
        skip_words = [
            "battle", "alliance", "ranking", "battlers",
            "deputy", "member", "officer", "elite", "leader",
            "elite wars", "points", "times", "entered",
        ]
        line_lower = line.lower()
        is_skip = False
        for w in skip_words:
            if line_lower == w:
                is_skip = True
                break
        if is_skip:
            i += 1
            continue
        # A pure number line is a score — look back for the player name
        if re.fullmatch(r"[\d,]+", line.replace(",", "")):
            i += 1
            continue
        # Check if this line looks like a player name
        # Player names appear before role (Deputy/Member/Officer/Elite)
        # and before score or "Not Entered"
        # We gather a 5-line window and look for patterns
        window = lines[i:i+5]
        window_text = " ".join(w.strip() for w in window)
        # Detect "Not Entered" in window
        not_entered = bool(
            re.search(r"not\s+entered", window_text, re.I)
        )
        # Detect a numeric score (could be 0) in window
        has_score = bool(
            re.search(r"\d{1,7}", window_text)
        )
        # Only treat as player name if followed by role or score
        has_role = bool(
            re.search(
                r"(deputy|member|officer|elite|leader)",
                window_text,
                re.I
            )
        )
        if (has_role or has_score or not_entered) and len(line) >= 2:
            # Clean up the name
            name = line.strip()
            # Skip lines that are just roles or numbers
            if re.fullmatch(r"(deputy|member|officer|elite|leader)", name, re.I):
                i += 1
                continue
            if re.fullmatch(r"[\d,]+", name):
                i += 1
                continue
            name_key = name.lower()
            if not_entered:
                results[name_key] = (name, False)
            else:
                results[name_key] = (name, True)
        i += 1
    return results


def match_roster_to_results(roster, ocr_results):
    # For each name in roster, fuzzy-match against OCR results
    matched = {}
    for player in roster:
        player_lower = player.strip().lower()
        # Exact match first
        if player_lower in ocr_results:
            matched[player] = ocr_results[player_lower][1]
            continue
        # Partial match — OCR name contains player name or vice versa
        found = False
        for ocr_key, (ocr_name, entered) in ocr_results.items():
            if player_lower in ocr_key or ocr_key in player_lower:
                matched[player] = entered
                found = True
                break
        if not found:
            # Not found in screenshots at all = not entered
            matched[player] = False
    return matched


def write_results_to_xlsx(xlsx_file, col_name, results):
    # Add a new column to alliance_tracker.xlsx
    wb = load_workbook(xlsx_file)
    if "Alliance Stats" not in wb.sheetnames:
        return None, "Sheet Alliance Stats not found in file."
    ws = wb["Alliance Stats"]
    # Find Player Name column (col A = 1)
    # Find next empty column
    max_col = ws.max_column
    new_col = max_col + 1
    # Write header
    header_cell = ws.cell(1, new_col, col_name)
    header_cell.font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
    header_cell.fill = PatternFill("solid", start_color="8B0000")
    header_cell.alignment = Alignment(
        horizontal="center", vertical="center", wrap_text=True
    )
    # Build lookup from sheet
    green_fill = PatternFill("solid", start_color="C6EFCE")
    red_fill = PatternFill("solid", start_color="FFC7CE")
    for row in range(2, ws.max_row + 1):
        cell_name = ws.cell(row, 1).value
        if not cell_name:
            continue
        cell_name_str = str(cell_name).strip()
        cell_name_lower = cell_name_str.lower()
        entered = None
        for rname, val in results.items():
            if rname.strip().lower() == cell_name_lower:
                entered = val
                break
        result_cell = ws.cell(row, new_col)
        result_cell.font = Font(name="Arial", bold=True, size=11)
        result_cell.alignment = Alignment(
            horizontal="center", vertical="center"
        )
        if entered is True:
            result_cell.value = "+"
            result_cell.fill = green_fill
            result_cell.font = Font(
                name="Arial", bold=True, size=11, color="276221"
            )
        elif entered is False:
            result_cell.value = "-"
            result_cell.fill = red_fill
            result_cell.font = Font(
                name="Arial", bold=True, size=11, color="9C0006"
            )
    from openpyxl.utils import get_column_letter as _gcl
    ws.column_dimensions[_gcl(new_col)].width = max(len(col_name) + 2, 12)
    ws.row_dimensions[1].height = 35
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf, None


# =====================
# UI
# =====================

# Step 1: Upload tracker
st.subheader("Step 1 - Upload Alliance Tracker")
xlsx_file = st.file_uploader(
    "Upload alliance_tracker.xlsx",
    type=["xlsx"],
    key="ev_xlsx"
)

if xlsx_file and xlsx_file.name != st.session_state.ev_last_xlsx:
    try:
        xl = pd.ExcelFile(xlsx_file)
        df = pd.read_excel(xl, sheet_name="Alliance Stats")
        roster = (
            df["Player Name"]
            .dropna()
            .astype(str)
            .str.strip()
            .tolist()
        )
        roster = [r for r in roster if r]
        st.session_state.roster = roster
        st.session_state.ev_last_xlsx = xlsx_file.name
        st.success(
            "Loaded " + str(len(roster)) + " players from roster."
        )
    except Exception as e:
        st.error("Error reading file: " + str(e))

if st.session_state.roster:
    with st.expander(
        "View Roster (" + str(len(st.session_state.roster)) + " players)"
    ):
        cols = st.columns(3)
        for i, name in enumerate(st.session_state.roster):
            cols[i % 3].write("- " + name)

st.divider()

# Step 2: Event name + screenshots
st.subheader("Step 2 - Event Details")

today_str = str(date.today())
event_name = st.text_input(
    "Event name (used as column header)",
    value="Elite Wars " + today_str,
    key="ev_name"
)

st.divider()

st.subheader("Step 3 - Upload Event Screenshots")
screenshots = st.file_uploader(
    "Upload all event leaderboard screenshots",
    type=["png", "jpg", "jpeg", "webp"],
    accept_multiple_files=True,
    key="ev_shots_" + str(st.session_state.ev_upload_key)
)

has_roster = bool(st.session_state.roster)
has_shots = bool(screenshots)
has_name = bool(event_name.strip())
can_process = has_roster and has_shots and has_name

col_a, col_b = st.columns([2, 1])
with col_a:
    do_process = st.button(
        "Process Event",
        disabled=not can_process,
        type="primary",
        use_container_width=True
    )
with col_b:
    if st.button("Clear", use_container_width=True):
        st.session_state.ev_upload_key += 1
        st.session_state.ev_results_ready = None
        st.rerun()

if not has_roster:
    st.caption("Upload the alliance tracker file first.")
elif not has_name:
    st.caption("Enter an event name above.")
elif not has_shots:
    st.caption("Upload at least one screenshot.")

if do_process and can_process:
    all_text = ""
    bar = st.progress(0)
    slot = st.empty()
    for i, f in enumerate(screenshots):
        slot.write(
            "Reading " + str(i+1) + "/" + str(len(screenshots)) + "..."
        )
        all_text += chr(10) + ocr_image(Image.open(f))
        bar.progress((i+1) / len(screenshots))
    slot.empty()
    bar.empty()
    ocr_results = extract_participants(all_text)
    matched = match_roster_to_results(
        st.session_state.roster, ocr_results
    )
    st.session_state.ev_matched = matched
    st.session_state.ev_event_name = event_name.strip()
    entered_count = sum(1 for v in matched.values() if v)
    not_entered_count = sum(1 for v in matched.values() if not v)
    st.success(
        "Done! " + str(entered_count) + " entered, "
        + str(not_entered_count) + " not entered."
    )
    with st.expander("Raw OCR text"):
        st.text_area("", all_text, height=150)

if "ev_matched" in st.session_state and st.session_state.ev_matched:
    st.divider()
    st.subheader("Results Preview")
    matched = st.session_state.ev_matched
    rows = []
    for name, entered in matched.items():
        rows.append({
            "Player": name,
            "Status": "+" if entered else "-"
        })
    results_df = pd.DataFrame(rows)
    entered_df = results_df[results_df["Status"] == "+"]
    not_entered_df = results_df[results_df["Status"] == "-"]
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**Entered (" + str(len(entered_df)) + ")**")
        st.dataframe(entered_df, use_container_width=True, hide_index=True)
    with col2:
        st.markdown("**Not Entered (" + str(len(not_entered_df)) + ")**")
        st.dataframe(not_entered_df, use_container_width=True, hide_index=True)
    st.divider()
    st.subheader("Step 4 - Download Updated Tracker")
    if xlsx_file is not None:
        xlsx_file.seek(0)
        buf, err = write_results_to_xlsx(
            xlsx_file,
            st.session_state.ev_event_name,
            matched
        )
        if err:
            st.error(err)
        else:
            st.download_button(
                label="Download Updated alliance_tracker.xlsx",
                data=buf,
                file_name="alliance_tracker.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
    else:
        st.warning(
            "Re-upload the alliance_tracker.xlsx above to save results."
        )

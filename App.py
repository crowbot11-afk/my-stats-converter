import streamlit as st
import pandas as pd
import pytesseract
from PIL import Image
import io
import re
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Battle Report", page_icon="⚔️", layout="centered")

st.title("⚔️ Battle Report Extractor")
st.write("Upload your Battle Report screenshot to extract stats into Excel.")

uploaded_file = st.file_uploader("Upload screenshot", type=["png", "jpg", "jpeg", "webp"])

def parse_battle_report(text):
    lines = [l.strip() for l in text.splitlines() if l.strip()]

    battle_stats = []
    stats_bonus = []
    artifacts = []

    section = None
    attacker_vals = {}
    defender_vals = {}

    i = 0
    while i < len(lines):
        line = lines[i]

        # Detect sections
        if re.search(r'overall info|battle report', line, re.I):
            i += 1
            continue
        if re.search(r'stats bonus', line, re.I):
            section = 'stats'
            i += 1
            continue
        if re.search(r'artifact', line, re.I):
            section = 'artifact'
            i += 1
            continue
        if re.search(r'attacker|defender|troops info', line, re.I):
            section = 'battle'
            i += 1
            continue

        # Parse Stats Bonus: "Stat Name   33.6%"
        if section == 'stats':
            match = re.match(r'^(.+?)\s+([\d,]+\.?\d*%?)$', line)
            if match:
                stats_bonus.append({
                    "Stat": match.group(1).strip(),
                    "Value": match.group(2).strip()
                })
            i += 1
            continue

        # Parse Artifact lines
        if section == 'artifact':
            artifacts.append({"Info": line})
            i += 1
            continue

        # Parse battle lines: numbers with labels
        if section == 'battle':
            # Lines like: "51,199 Losses 0" or "Survivors 134 Kills 78"
            match = re.match(r'^([\w\s]+?)\s+([\d,]+)\s+([\w\s]+?)\s+([\d,]+)$', line)
            if match:
                battle_stats.append({
                    "Stat": match.group(1).strip() + " / " + match.group(3).strip(),
                    "Attacker": match.group(2).replace(',', ''),
                    "Defender": match.group(4).replace(',', '')
                })
            else:
                # Try single number line
                match2 = re.match(r'^([\w\s]+?)\s+([\d,]+)$', line)
                if match2:
                    battle_stats.append({
                        "Stat": match2.group(1).strip(),
                        "Attacker": match2.group(2).replace(',', ''),
                        "Defender": ""
                    })
            i += 1
            continue

        # Default: try to detect stat/value pairs anywhere
        match = re.match(r'^(.+?)\s{2,}([\d,]+\.?\d*%?)$', line)
        if match:
            stats_bonus.append({
                "Stat": match.group(1).strip(),
                "Value": match.group(2).strip()
            })
        i += 1

    return battle_stats, stats_bonus, artifacts

def build_excel(battle_stats, stats_bonus, artifacts, raw_text):
    wb = Workbook()

    header_font = Font(name='Arial', bold=True, color='FFFFFF', size=11)
    header_fill = PatternFill('solid', start_color='8B0000')  # Dark red like game UI
    green_fill = PatternFill('solid', start_color='C6EFCE')
    blue_fill = PatternFill('solid', start_color='BDD7EE')
    title_font = Font(name='Arial', bold=True, size=13)
    center = Alignment(horizontal='center', vertical='center')
    thin = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )

    def style_header(cell, fill=None):
        cell.font = header_font
        cell.fill = fill or header_fill
        cell.alignment = center
        cell.border = thin

    def style_cell(cell, bold=False):
        cell.font = Font(name='Arial', bold=bold, size=10)
        cell.alignment = Alignment(vertical='center')
        cell.border = thin

    # ── Sheet 1: Battle Stats ──
    ws1 = wb.active
    ws1.title = "Battle Stats"
    ws1['A1'] = "⚔️ Battle Report - Combat Stats"
    ws1['A1'].font = title_font
    ws1.merge_cells('A1:C1')
    ws1['A1'].alignment = center

    ws1.append([])
    ws1.append(["Stat", "Attacker", "Defender"])
    for cell in ws1[3]:
        style_header(cell)

    if battle_stats:
        for row in battle_stats:
            ws1.append([row.get("Stat",""), row.get("Attacker",""), row.get("Defender","")])
            r = ws1.max_row
            style_cell(ws1.cell(r, 1), bold=True)
            style_cell(ws1.cell(r, 2))
            style_cell(ws1.cell(r, 3))
    else:
        ws1.append(["No battle stats detected - check Raw Text sheet"])

    ws1.column_dimensions['A'].width = 35
    ws1.column_dimensions['B'].width = 15
    ws1.column_dimensions['C'].width = 15

    # ── Sheet 2: Stats Bonus ──
    ws2 = wb.create_sheet("Stats Bonus")
    ws2['A1'] = "📊 Stats Bonus"
    ws2['A1'].font = title_font
    ws2.merge_cells('A1:B1')
    ws2['A1'].alignment = center

    ws2.append([])
    ws2.append(["Stat Name", "Value"])
    for cell in ws2[3]:
        style_header(cell)

    if stats_bonus:
        for idx, row in enumerate(stats_bonus):
            ws2.append([row.get("Stat",""), row.get("Value","")])
            r = ws2.max_row
            fill = green_fill if idx % 2 == 0 else PatternFill('solid', start_color='EBF5EB')
            ws2.cell(r, 1).fill = fill
            ws2.cell(r, 2).fill = fill
            style_cell(ws2.cell(r, 1))
            style_cell(ws2.cell(r, 2))
    else:
        ws2.append(["No stats bonus detected - check Raw Text sheet"])

    ws2.column_dimensions['A'].width = 40
    ws2.column_dimensions['B'].width = 15

    # ── Sheet 3: Artifacts ──
    if artifacts:
        ws3 = wb.create_sheet("Artifacts")
        ws3['A1'] = "🏺 Artifact Info"
        ws3['A1'].font = title_font
        ws3.append([])
        ws3.append(["Info"])
        style_header(ws3.cell(3, 1))
        for row in artifacts:
            ws3.append([row.get("Info","")])
            style_cell(ws3.cell(ws3.max_row, 1))
        ws3.column_dimensions['A'].width = 50

    # ── Sheet 4: Raw Text ──
    ws4 = wb.create_sheet("Raw OCR Text")
    ws4['A1'] = "Raw Text (for reference / manual correction)"
    ws4['A1'].font = Font(bold=True)
    for i, line in enumerate(raw_text.splitlines(), start=2):
        ws4.cell(i, 1).value = line
    ws4.column_dimensions['A'].width = 80

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

if uploaded_file:
    image = Image.open(uploaded_file)
    st.image(image, caption="Uploaded Screenshot", use_column_width=True)

    with st.spinner("Reading text from image..."):
        raw_text = pytesseract.image_to_string(image)

    with st.expander("📄 Raw OCR Text (tap to expand)"):
        st.text_area("", raw_text, height=200)

    battle_stats, stats_bonus, artifacts = parse_battle_report(raw_text)

    st.subheader("⚔️ Battle Stats")
    if battle_stats:
        st.dataframe(pd.DataFrame(battle_stats), use_container_width=True)
    else:
        st.info("No battle stats detected. Check Raw OCR Text above.")

    st.subheader("📊 Stats Bonus")
    if stats_bonus:
        df_bonus = pd.DataFrame(stats_bonus)
        edited_bonus = st.data_editor(df_bonus, num_rows="dynamic", use_container_width=True)
        stats_bonus = edited_bonus.to_dict('records')
    else:
        st.info("No stats bonus detected. Check Raw OCR Text above.")

    if artifacts:
        st.subheader("🏺 Artifacts")
        st.dataframe(pd.DataFrame(artifacts), use_container_width=True)

    excel_buf = build_excel(battle_stats, stats_bonus, artifacts, raw_text)

    st.download_button(
        label="⬇️ Download Battle Report as Excel",
        data=excel_buf,
        file_name="battle_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

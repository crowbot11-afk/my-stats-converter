import streamlit as st
import pandas as pd
import pytesseract
from PIL import Image
import io
import re

st.set_page_config(page_title="Game Stats OCR", page_icon="🎮", layout="centered")

st.title("🎮 Game Stats Reader")
st.write("Upload a screenshot and extract the stats as Excel.")

uploaded_file = st.file_uploader("Upload your screenshot", type=["png", "jpg", "jpeg", "webp"])

if uploaded_file:
    image = Image.open(uploaded_file)
    st.image(image, caption="Uploaded Screenshot", use_column_width=True)

    with st.spinner("Reading text from image..."):
        raw_text = pytesseract.image_to_string(image)

    st.subheader("Extracted Text")
    st.text_area("Raw OCR Output", raw_text, height=200)

    lines = [line.strip() for line in raw_text.splitlines() if line.strip()]

    rows = []
    for line in lines:
        parts = re.split(r'\s{2,}|\t|:', line, maxsplit=1)
        if len(parts) == 2:
            rows.append({"Stat": parts[0].strip(), "Value": parts[1].strip()})
        else:
            rows.append({"Stat": line, "Value": ""})

    df = pd.DataFrame(rows)

    st.subheader("Parsed Table")
    edited_df = st.data_editor(df, num_rows="dynamic", use_container_width=True)

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        edited_df.to_excel(writer, index=False, sheet_name="Game Stats")
    buffer.seek(0)

    st.download_button(
        label="Download as Excel",
        data=buffer,
        file_name="game_stats.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )        parts = re.split(r'\s{2,}|\t|:|\\|', line, maxsplit=1)
        if len(parts) == 2:
            rows.append({"Stat": parts[0].strip(), "Value": parts[1].strip()})
        else:
            rows.append({"Stat": line, "Value": ""})

    df = pd.DataFrame(rows)

    st.subheader("📊 Parsed Table")
    edited_df = st.data_editor(df, num_rows="dynamic", use_container_width=True)

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        edited_df.to_excel(writer, index=False, sheet_name="Game Stats")
    buffer.seek(0)

    st.download_button(
        label="⬇️ Download as Excel",
        data=buffer,
        file_name="game_stats.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )        parts = re.split(r'\s{2,}|\t|:|\\|', line, maxsplit=1)
        if len(parts) == 2:
            rows.append({"Stat": parts[0].strip(), "Value": parts[1].strip()})
        else:
            rows.append({"Stat": line, "Value": ""})

    df = pd.DataFrame(rows)

    st.subheader("📊 Parsed Table")
    edited_df = st.data_editor(df, num_rows="dynamic", use_container_width=True)

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        edited_df.to_excel(writer, index=False, sheet_name="Game Stats")
    buffer.seek(0)

    st.download_button(
        label="⬇️ Download as Excel",
        data=buffer,
        file_name="game_stats.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )    # Split into lines, filter blanks
    lines = [line.strip() for line in raw_text.splitlines() if line.strip()]

    # Try to split each line into key/value on whitespace or common delimiters
    rows = []
    for line in lines:
        # Split on multiple spaces, tab, colon, or pipe
        parts = re.split(r'\s{2,}|\t|:|\\|', line, maxsplit=1)
        if len(parts) == 2:
            rows.append({"Stat": parts[0].strip(), "Value": parts[1].strip()})
        else:
            rows.append({"Stat": line, "Value": ""})

    df = pd.DataFrame(rows)

    st.subheader("📊 Parsed Table")
    # Make it editable so user can fix OCR errors
    edited_df = st.data_editor(df, num_rows="dynamic", use_container_width=True)

    # --- Download as Excel ---
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        edited_df.to_excel(writer, index=False, sheet_name="Game Stats")
    buffer.seek(0)

    st.download_button(
        label="⬇️ Download as Excel",
        data=buffer,
        file_name="game_stats.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
```

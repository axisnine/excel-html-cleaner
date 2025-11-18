import streamlit as st
import pandas as pd
import re
from io import BytesIO

# Allowed HTML tags
ALLOWED_TAGS = {"p", "em", "strong", "br"}

# Regex to detect HTML tags (opening, closing, self-closing)
TAG_REGEX = re.compile(r'<\s*/?\s*([a-zA-Z0-9]+)[^>]*>', re.IGNORECASE)


def find_problematic_html(value):
    """Return True if the cell contains any disallowed HTML tags."""
    if not isinstance(value, str):
        return False

    tags = TAG_REGEX.findall(value)
    if not tags:
        return False

    # If ANY tag is not in the allowed whitelist, it's problematic
    for tag in tags:
        if tag.lower() not in ALLOWED_TAGS:
            return True
    return False


def clean_html(value):
    """Remove disallowed HTML tags but KEEP inner text.
       Allowed tags remain exactly as written."""
    if not isinstance(value, str):
        return value

    def replace_tag(match):
        full_tag = match.group(0)
        tag_name = match.group(1).lower()

        if tag_name in ALLOWED_TAGS:
            return full_tag   # keep allowed tags as-is
        else:
            return ""         # strip the tag ONLY

    cleaned = TAG_REGEX.sub(replace_tag, value)

    return cleaned


# ============================
# STREAMLIT APP
# ============================

st.title("Excel HTML Cleaner")
st.write("Upload an Excel file and I’ll remove unwanted HTML while keeping allowed formatting intact.")

uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file, dtype=str)

    st.success("File uploaded successfully.")

    # Detect problematic cells
    problematic_cells = []
    for row_idx, row in df.iterrows():
        for col in df.columns:
            if find_problematic_html(row[col]):
                problematic_cells.append((row_idx, col))

    num_problematic = len(problematic_cells)

    if num_problematic == 0:
        st.info("No problematic HTML found. Your file is clean.")
    else:
        st.warning(f"Found **{num_problematic}** cells containing unwanted HTML tags.")

        with st.expander("See problematic cells (optional)"):
            for row_idx, col in problematic_cells:
                st.write(f"Row {row_idx+1}, Column '{col}': {df.at[row_idx, col]}")

        if st.button("Clean the HTML"):
            cleaned_df = df.copy()
            for row_idx, col in problematic_cells:
                cleaned_df.at[row_idx, col] = clean_html(cleaned_df.at[row_idx, col])

            # Convert cleaned df back to Excel file in memory
            output = BytesIO()
            with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                cleaned_df.to_excel(writer, index=False)
            output.seek(0)

            st.success("Cleaning complete. Download your cleaned Excel file below:")

            st.download_button(
                label="Download cleaned file",
                data=output,
                file_name="cleaned_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

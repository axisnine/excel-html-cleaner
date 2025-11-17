import streamlit as st
import pandas as pd
import re
from io import BytesIO

# Regex pattern to detect HTML tags
html_pattern = re.compile(r"<[^>]+>")

def find_html_cells(df):
    """Return a list of cells that contain HTML-like content."""
    mask = df.applymap(lambda x: bool(html_pattern.search(str(x))) if pd.notnull(x) else False)
    problems = []
    for row_idx, row in df.iterrows():
        for col_name, value in row.items():
            if mask.loc[row_idx, col_name]:
                problems.append({
                    "Row": row_idx + 1,
                    "Column": col_name,
                    "Value": value
                })
    return problems

def clean_html(x):
    """Strip HTML tags from a cell."""
    return html_pattern.sub("", x) if isinstance(x, str) else x

# ----------------------- Streamlit UI -----------------------

st.title("Excel HTML Cleaner Utility")
st.write("Upload an Excel file to detect and remove unwanted HTML tags.")

uploaded = st.file_uploader("Upload Excel file", type=["xlsx"])

if uploaded:
    df = pd.read_excel(uploaded)
    
    st.subheader("Preview of data (first 10 rows)")
    st.dataframe(df.head(10))

    problems = find_html_cells(df)

    if problems:
        st.subheader("Detected HTML-like content")
        st.write(f"Found **{len(problems)}** problematic cells.")
        st.dataframe(pd.DataFrame(problems).head(50))

        if st.checkbox("Confirm and clean HTML from all cells"):
            df_cleaned = df.applymap(clean_html)

            output = BytesIO()
            df_cleaned.to_excel(output, index=False)
            output.seek(0)

            st.success("Cleaning complete! Download your cleaned file below:")
            st.download_button(
                "Download cleaned Excel",
                data=output,
                file_name="cleaned_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.success("No HTML-like content found.")

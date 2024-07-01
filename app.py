import os
import pandas as pd
import pytesseract
from PIL import Image
import streamlit as st

# Function to extract text from image using pytesseract
def extract_text_from_image(image):
    return pytesseract.image_to_string(image)

# Function to convert extracted text to a structured table (DataFrame)
def text_to_table(text):
    rows = text.strip().split('\n')
    data = [row.split() for row in rows if row.strip()]
    return pd.DataFrame(data)

# Function to generate VBA code for Excel from the DataFrame
def generate_vba_code(df):
    vba_code = "Sub LoadDataIntoExcel()\n"
    vba_code += "    Dim ws As Worksheet\n"
    vba_code += "    Set ws = ThisWorkbook.Sheets(1)\n"
    for i, row in df.iterrows():
        for j, cell in enumerate(row):
            vba_code += f"    ws.Cells({i+1}, {j+1}).Value = \"{cell}\"\n"
    vba_code += "End Sub"
    return vba_code

# Function to generate JavaScript code for Google Sheets from the DataFrame
def generate_js_code(df):
    js_code = """
function loadDataToGoogleSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = [
"""
    for index, row in df.iterrows():
        js_code += "    ["
        js_code += ", ".join([f'"{cell}"' for cell in row])
        js_code += "],\n"
    js_code = js_code.rstrip(",\n") + "\n  ];\n"
    js_code += """
  for (var i = 0; i < data.length; i++) {
    for (var j = 0; j < data[i].length; j++) {
      sheet.getRange(i + 1, j + 1).setValue(data[i][j]);
    }
  }
}
"""
    return js_code

# Streamlit App
def main():
    st.title("Batch Image to Code Converter for Excel and Google Sheets")
    st.write("Upload multiple images to extract text and generate code.")

    uploaded_files = st.file_uploader("Choose image files", type=["jpg", "jpeg", "png"], accept_multiple_files=True)

    if uploaded_files:
        code_type = st.selectbox("Select Code Type", ["VBA for Excel", "JavaScript for Google Sheets"])
        st.write(f"Processing {len(uploaded_files)} image(s)...")
        
        for i, uploaded_file in enumerate(uploaded_files):
            st.write(f"Image {i + 1}:")
            image = Image.open(uploaded_file)
            st.image(image, caption=f'Uploaded Image {i + 1}', use_column_width=True)
            
            st.write("Extracting text...")
            extracted_text = extract_text_from_image(image)
            st.text_area(f"Extracted Text from Image {i + 1}", value=extracted_text, height=100)
            
            st.write("Converting text to table...")
            df = text_to_table(extracted_text)
            st.write(df)
            
            if code_type == "VBA for Excel":
                st.write("Generating VBA code for Excel...")
                vba_code = generate_vba_code(df)
                st.code(vba_code, language='vbscript')
            else:
                st.write("Generating JavaScript code for Google Sheets...")
                js_code = generate_js_code(df)
                st.code(js_code, language='javascript')
            
            st.markdown("---")  # Separator between images

if __name__ == "__main__":
    main()

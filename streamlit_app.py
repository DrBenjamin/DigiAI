import streamlit as st
import os
import io
import pandas as pd
import openai
import xlsxwriter
import openpyxl
import xlrd
import PyPDF2




#### Functions
### Function: export_excel = Pandas dataframe to MS Excel Makro File (xlsm)
def export_excel(sheet, data, excel_file_name = 'Digital_Landscape_GIZ.xlsm'):
    # Create an Excel file filled with a pandas dataframe using XlsxWriter as engine
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine = 'xlsxwriter') as writer:
        # Add dataframe data to worksheet
        data.to_excel(writer, sheet_name = sheet, index = False)

        # Define worksheet
        worksheet = writer.sheets[sheet]

        # Add a table to the worksheet
        #span = "A1:A2"
        #worksheet.add_table(span, {'columns': columns})
        range_table = "A:A"
        worksheet.set_column(range_table, 30)

        # Add Excel VBA code
        workbook = writer.book
        workbook.add_vba_project('vbaProject.bin')

        # Saving changes
        workbook.close()
        writer.close()

        # Download Button
        st.download_button(label = 'Download Excel document', data = buffer, file_name = excel_file_name,
                           mime = "application/vnd.ms-excel.sheet.macroEnabled.12")




#### Main App
st.header("Digitalization Advisor")
st.subheader('What is your project / approach about?')
st.write("Welcome to the Digitalization Advisor. This tool will help you to identify the best digitalization initiatives in GIZ to support your individual project. Please answer the following questions to get started.")

st.checkbox("The project is with a partner organisation")
input_text = '"""'
input_text += st.text_area("What is your project about (also add keywords to describe your project)?")
st.write("... you can also upload a describing PDF file (also in addition to the text above)")

# Upload PDF file
uploaded_file = st.file_uploader(label = 'Choose a PDF file to upload', type = 'pdf')
if uploaded_file is not None:
    file_name = os.path.join('PDFs', uploaded_file.name)
    file = open(file_name, 'wb')
    file.write(uploaded_file.getvalue())
    file.close()
    
# Extact text from PDF document
try:
    reader = PyPDF2.PdfReader(file_name)
    reader_text = ""
    for i in range(len(reader.pages)):
        reader_text += reader.pages[i].extract_text()
    reader_text = reader_text.replace('\n', ' ')
except:
    print('No PDF file uploaded')

# Download Excel file
submitted = st.button("Submit")
if submitted:
    st.write("Thank you for your submission. Your Excel document is in preparation...")
    
    # Using ChatGPT from OpenAI to shorten PDF extracted text
    # Set API key
    openai.api_key = st.secrets['openai']['key']
                
    # Doing the requests to OpenAI for summarizing / keyword extracting the question
    try:
        # Creating summary of user question
        model = 'gpt-3.5-turbo'
        response_summary = openai.ChatCompletion.create(model = model, messages = [{"role": "system", "content": "You do summarization."}, {"role": "user", "content": reader_text[:3000]},])
        summary_text = response_summary['choices'][0]['message']['content'].lstrip()
        summary_text = summary_text.replace('\n', ' ')
        input_text += " " + summary_text
    except:
        print('ChatGPT failed')
    input_text += '"""'
    st.write('Download your personalized Excel document.')
    export_excel(sheet = 'Project Description', data = pd.DataFrame([input_text]))
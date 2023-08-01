import streamlit as st
import os
import PyPDF2


### Function: export_excel = Pandas dataframe to MS Excel Makro File (xlsm)
def export_excel(sheet, column, columns, length, data,
                 sheet2 = 'N0thing', column2 = 'A', columns2 = '', length2 = '', data2 = '',
                 sheet3 = 'N0thing', column3 = 'A', columns3 = '', length3 = '', data3 = '',
                 sheet4 = 'N0thing', column4 = 'A', columns4 = '', length4 = '', data4 = '',
                 sheet5 = 'N0thing', column5 = 'A', columns5 = '', length5 = '', data5 = '',
                 sheet6 = 'N0thing', column6 = 'A', columns6 = '', length6 = '', data6 = '',
                 sheet7 = 'N0thing', column7 = 'A', columns7 = '', length7 = '', data7 = '',
                 image = 'NoImage', image_pos = 'D1', excel_file_name = 'Export.xlsm'):


    ## Store fuction arguments in array
    # Create empty array
    func_arr = []

    # Add function arguments to array
    func_arr.append([sheet, column, columns, length, data])
    func_arr.append([sheet2, column2, columns2, length2, data2])
    func_arr.append([sheet3, column3, columns3, length3, data3])
    func_arr.append([sheet4, column4, columns4, length4, data4])
    func_arr.append([sheet5, column5, columns5, length5, data5])
    func_arr.append([sheet6, column6, columns6, length6, data6])
    func_arr.append([sheet7, column7, columns7, length7, data7])


    ## Create an Excel file filled with a pandas dataframe using XlsxWriter as engine
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine = 'xlsxwriter') as writer:
        for i in range(7):
            if (func_arr[i][0] != 'N0thing'):
                # Add dataframe data to worksheet
                func_arr[i][4].to_excel(writer, sheet_name = func_arr[i][0], index = False)

                # Define worksheet
                worksheet = writer.sheets[func_arr[i][0]]

                # Add a table to the worksheet
                if func_arr[i][1] != 'A':
                    span = "A1:%s%s" % (func_arr[i][1], func_arr[i][3])
                    worksheet.add_table(span, {'columns': func_arr[i][2]})
                    range_table = "A:" + func_arr[i][1]
                    worksheet.set_column(range_table, 30)

                # Add image to worksheet
                if (image != 'NoImage'):
                    # Saving image as png to a buffer
                    # byteIO = io.BytesIO()
                    # image.save(byteIO, format = 'PNG')
                    # pic = byteIO.getvalue()

                    # Saving image as png temp file
                    f = open('files/temp.png', 'wb')
                    f.write(image)
                    f.close()

                    # Insert in worksheet
                    worksheet.insert_image(image_pos, 'files/temp.png')


        ## Add Excel VBA code
        workbook = writer.book
        workbook.add_vba_project('files/vbaProject.bin')


        ## Saving changes
        workbook.close()
        writer.save()


        ## Download Button
        st.download_button(label = 'Download Excel document', data = buffer, file_name = excel_file_name,
                           mime = "application/vnd.ms-excel.sheet.macroEnabled.12")




#### Header
st.header("Digitalization Advisor")
st.subheader('What is your project / approach about?')
st.write("Welcome to the Digitalization Advisor. This tool will help you to identify the best digitalization initiatives in GIZ to support your individual project. Please answer the following questions to get started.")

st.checkbox("The project is with a partner organisation")
input_text = ""
input_text = st.text_area("What is your project about?")
st.write("... you can also upload a describing PDF file (also in addition to the text above)")

# Upload PDF file
uploaded_file = st.file_uploader(label = 'Choose a PDF file to upload', type = 'pdf')
if uploaded_file is not None:
    file_name = os.path.join('PDFs', uploaded_file.name)
    file = open(file_name, 'wb')
    file.write(uploaded_file.getvalue())
    file.close()
    
# Extact text from PDF document
reader = PyPDF2.PdfReader(file_name)
for i in range(len(reader.pages)):
    input_text += reader.pages[i].extract_text()

# Download Excel file
submitted = st.button("Submit")
if submitted:
    st.write("Thank you for your submission. Download your personalized Excel document.")
    st.text_area("Your project description", input_text)

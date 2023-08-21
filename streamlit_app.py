##### `streamlit_app.py`
##### DigiAI
##### Open-Source, hosted on https://github.com/DrBenjamin/DigiAI
##### Please reach out to ben@benbox.org for any questions
#### Loading needed Python libraries
import streamlit as st
import os
import io
import sys
from zipfile import ZipFile
from zipfile import BadZipfile
import pandas as pd
import openai
import xlsxwriter
import openpyxl
import xlrd
import PyPDF2
from PIL import Image




#### Session states
## Session states
if 'version' not in st.session_state:
    st.session_state['version'] = 'V1.0'
if 'font_color' not in st.session_state:
    st.session_state['font_color'] = 'vbWhite'




#### Functions
### Function: check_password = Password / user checking
def check_password():
    # Session states
    if ("username" not in st.session_state):
        st.session_state["username"] = ''
    if ("password" not in st.session_state):
        st.session_state["password"] = ''
    if ("admin" not in st.session_state):
        st.session_state["admin"] = False
    if ("password_correct" not in st.session_state):
        st.session_state["password_correct"] = False
    if ('logout' not in st.session_state):
        st.session_state['logout'] = False
    
    # Checks whether a password entered by the user is correct
    def password_entered():
        try:
            if st.session_state["username"] in st.secrets["passwords"] and st.session_state["password"] == st.secrets["passwords"][
                st.session_state["username"]]:
                st.session_state["password_correct"] = True
                st.session_state["admin"] = False
                
                # Delete username + password
                del st.session_state["password"]
                del st.session_state["username"]
            
            # Checks whether a password entered by the user is correct for admins
            elif st.session_state["username"] in st.secrets["admins"] and st.session_state["password"] == st.secrets["admins"][
                st.session_state["username"]]:
                st.session_state["password_correct"] = True
                st.session_state["admin"] = True
                
                # Delete username + password
                del st.session_state["password"]
                del st.session_state["username"]
            
            # No combination fits
            else:
                st.session_state["password_correct"] = False
        except Exception as e:
            print('Exception in `password_entered` function. Error: ', e)
            st.session_state["password_correct"] = False
    
    ## Sidebar
    st.sidebar.header('Digitalization Advisor')
    # Show Sidebar Header Image
    #st.sidebar.image('Images/Africa_Digitalization.jpg', use_column_width = True)
    
    # First run, show inputs for username + password
    if "password_correct" not in st.session_state:
        st.sidebar.subheader('Please enter username and password')
        st.sidebar.text_input(label = "Username", on_change = password_entered, key = "username")
        st.sidebar.text_input(label = "OTP", type = "password", on_change = password_entered, key = "password")
        return False
    
    # Password not correct, show input + error
    elif not st.session_state["password_correct"]:
        st.sidebar.text_input(label = "Username", on_change = password_entered, key = "username")
        st.sidebar.text_input(label = "OTP", type = "password", on_change = password_entered, key = "password")
        if (st.session_state['logout']):
            st.sidebar.success('Logout successful!', icon = "✅")
        else:
            st.sidebar.error(body = "User not known or OTP incorrect!", icon = "🚨")
        return False
    
    else:
        # Password correct
        st.sidebar.success(body = ' You are logged in.', icon = "✅")
        st.sidebar.info(body = ' You can close this menu now.', icon = '☝🏾️')
        st.sidebar.button(label = 'Logout', on_click = logout)
        return True



### Funtion: logout = Logout button
def logout():
    # Set `logout` to get logout-message
    st.session_state['logout'] = True
    
    # Set password to `false`
    st.session_state["password_correct"] = False



### Function: extract_macro = Extract VBA code from MS Excel document (xlsm)
def extract_macro(xlsm_file = 'Excel/Digital_Landscape_GIZ.xlsm', vba_filename = 'vbaProject.bin'):
    try:
        # Open the Excel xlsm file as a zip file.
        xlsm_zip = ZipFile(xlsm_file, 'r')

        # Read the xl/vbaProject.bin file.
        vba_data = xlsm_zip.read('xl/' + vba_filename)

        # Write the vba data to a local file.
        vba_file = open('Excel/' + vba_filename, "wb")
        vba_file.write(vba_data)
        vba_file.close()
        print("Extracted: %s" % vba_filename)

    except IOError as e:
        print("File error: %s" % str(e))
        print("Not Extraced, using chached version.")

    except KeyError as e:
        # Usually when there isn't a xl/vbaProject.bin member in the file.
        print("File error: %s" % str(e))
        print("File may not be an Excel xlsm macro file: '%s'" % xlsm_file)
        print("Not Extraced, using chached version.")

    except BadZipfile as e:
        # Usually if the file is a xls file and not a xlsm file.
        print("File error: %s: '%s'" % (str(e), xlsm_file))
        print("File may not be an Excel xlsm macro file.")
        print("Not Extraced, using chached version.")

    except Exception as e:
        # Catch any other exceptions.
        print("File error: %s" % str(e))
        print("Not Extraced, using chached version.")



### Function: import_excel = Read pandas dataframe from MS Excel document (xlsx)
def import_excel(excel_file_name):
    try:
        df = pd.read_excel(excel_file_name)
        return df
    except FileNotFoundError:
        print(f"File '{excel_file_name}' not found.")



### Function: export_excel = Pandas dataframe to MS Excel Makro File (xlsm)
def export_excel(sheet, data, sheet2, keywords, sheet3, landscape, excel_file_name = 'Digital_Landscape_GIZ.xlsm', image = 'NoImage'):
    # Create an Excel file filled with a pandas dataframe using XlsxWriter as engine
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine = 'xlsxwriter') as writer:
        # Add data to worksheet
        data.to_excel(writer, sheet_name = sheet, index = False)
        keywords.to_excel(writer, sheet_name = sheet2, index = False)
        landscape.to_excel(writer, sheet_name = sheet3, index = False)
        pd.DataFrame([st.session_state['font_color']]).to_excel(writer, sheet_name = "Wallpaper", index = False)

        # Add image to worksheet
        worksheet = writer.sheets['Wallpaper']
        if (image != 'NoImage'):
            # Saving image as png to a buffer
            byteIO = io.BytesIO()
            image.save(byteIO, format = 'JPEG')
            pic = byteIO.getvalue()

            # Saving image as png temp file
            f = open('Images/temp.jpg', 'wb')
            f.write(pic)
            f.close()
            
            # Insert in worksheet
            worksheet.insert_image("A3", 'Images/temp.jpg')

        # Add Excel VBA code
        workbook = writer.book
        workbook.add_vba_project('Excel/vbaProject.bin')

        # Saving changes
        workbook.close()
        writer.close()

        # Download Button
        st.download_button(label = 'Download Excel document', data = buffer, file_name = excel_file_name,
                           mime = "application/vnd.ms-excel.sheet.macroEnabled.12")




#### Main App
if check_password():
    st.header("Digitalization Advisor")
    st.subheader('Get help to find support in GIZ for your digitalization project')
    st.write("Welcome to the Digitalization Advisor. This tool will help you to identify the best digitalization initiatives in GIZ to support your individual project. Please answer the following questions to get started.")

    # Upload Excel file
    uploaded_file = st.file_uploader(label = 'Do you want to upload a new Digital landscape GIZ file version?', type = 'xlsx')
    if uploaded_file is not None:
        file_name = os.path.join('Excel', uploaded_file.name)
        file = open(file_name, 'wb')
        file.write(uploaded_file.getvalue())
        file.close()

    # Get a list of files in a folder
    filez = os.listdir('Excel/')
    versions = []
    for file in filez:
        if file[:27] == 'Digital_Landscape_GIZ_List_' and file[-5:] == '.xlsx':
            versions.append(file[27:31])
    st.session_state['version'] = st.selectbox(label = "Which Digital Landscape GIZ file version should be used?", options = versions, index = len(versions) - 1, disabled = False)

    # Upload Wallpaper image
    uploaded_file = st.file_uploader(label = 'Do you want to upload a customized Wallpaper?', type = 'jpg')
    if uploaded_file is not None:
        file_name = os.path.join('Images', uploaded_file.name)
        file = open(file_name, 'wb')
        file.write(uploaded_file.getvalue())
        file.close()
        excel_image = file_name
        st.session_state['font_color'] = st.selectbox(label = "Which font color should be used?", options = ['Black', 'White'])
    else:
        excel_image = 'Images/Wallpaper.jpg'
        st.session_state['font_color'] = 'White'

    # User Input
    input_text = '"""'
    input_text += st.text_area("What is your digitalization project about?", placeholder = "NO project numbers, names, mail adresses, phone numbers and confidential information!")
    st.warning("NO confidential information, as data will be processed by OpenAI!", icon = "🔥")
    input_keywords = ''
    input_keywords = st.text_area("What are the keywords of your digitalization project?", placeholder = "NO project numbers, names, mail adresses, phone numbers and confidential information!")

    # Upload PDF file
    uploaded_file = st.file_uploader(label = 'Do you want to upload a PDF file with unconfidential / public information?', type = 'pdf')
    if uploaded_file is not None:
        file_name = os.path.join('PDFs', uploaded_file.name)
        file = open(file_name, 'wb')
        file.write(uploaded_file.getvalue())
        file.close()
    st.warning("NO confidential document, as data will be processed by OpenAI!", icon = "🔥")
        
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
        st.write("Thank you for your submission.")
        with st.spinner('Wait for your Excel document...'):
            ## Using ChatGPT from OpenAI to shorten PDF extracted text
            # Set API key
            openai.api_key = st.secrets['openai']['key']
                        
            # Doing the requests to OpenAI for summarizing
            model = 'gpt-3.5-turbo'
            try:
                # Creating summary of user question
                response_summary = openai.ChatCompletion.create(model = model, messages = [{"role": "system", "content": "You do summarization."}, {"role": "user", "content": reader_text[:3000]},])
                summary_text = response_summary['choices'][0]['message']['content'].lstrip()
                summary_text = summary_text.replace('\n', ' ')
                input_text += " " + summary_text
                print('ChatGPT summarization successful')
            except:
                print('ChatGPT summarization failed')
            input_text += '"""'

            # Doing the requests to OpenAI for keyword extracting
            try:
                # Extracting keywords
                response_keywords = openai.ChatCompletion.create(model = model, messages = [{"role": "system", "content": "You do keyword extraction."}, {"role": "user", "content": input_text},])
                keywords = response_keywords['choices'][0]['message']['content'].lstrip()
                input_keywords += ", " + keywords
                input_keywords = set(input_keywords.split(', '))
                input_keywords = ', '.join(input_keywords)
                print('ChatGPT keyword extraction successful')
            except:
                print('ChatGPT keyword extraction failed')


            ## Export Excel file
            extract_macro()
            export_excel(sheet = 'Project description', data = pd.DataFrame([input_text]), sheet2 = 'Project keywords', keywords = pd.DataFrame([input_keywords]), sheet3 = 'Digital landscape GIZ', landscape = import_excel(excel_file_name = 'Excel/Digital_Landscape_GIZ_List_' + st.session_state['version'] + '.xlsx'), image = Image.open(excel_image))
        st.toast("Your Excel document is ready for download.", icon = "👍")
        st.balloons()
else:
    st.info("Please enter your username and password on the left side.", icon = "🔒")
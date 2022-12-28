import streamlit as st
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import plotly.graph_objects as go
import plotly.io as pio

import xlsxwriter
import base64
import io
import requests
import datetime

from PIL import Image
from streamlit_extras.badges import badge


def main():
    col1, col2, col3 = st.columns([0.05, 0.265, 0.035])
    
    with col1:
        url = 'https://github.com/tsu2000/gpa_calculator/raw/main/images/grade.png'
        response = requests.get(url)
        img = Image.open(io.BytesIO(response.content))
        st.image(img, output_format = 'png')

    with col2:
        st.title('&nbsp; Module GPA Calculator')

    with col3:
        badge(type = 'github', name = 'tsu2000/gpa_calculator', url = 'https://github.com/tsu2000/gpa_calculator')

    # Create sidebar with options
    with st.sidebar:  
        st.markdown('# 🔀 &nbsp; Navigation Bar') 
        st.markdown('###')
        selected_uni = st.selectbox('Select the university you are part of:', ['NTU (Nanyang Technological University)',
                                                                               'SMU (Singapore Management University)',
                                                                               'SUTD (Singapore University of Technology & Design)'])
        st.markdown('---')
        feature = st.radio('Select a feature:', ['Current GPA Analysis',
                                                 'GPA Calculation Explanation'])   

    # Select option
    if feature == 'Current GPA Analysis':
        calc(selected_uni = selected_uni)
    elif feature == 'GPA Calculation Explanation':
        explain()
    
    
def calc(selected_uni):
    st.markdown('#### :bar_chart: &nbsp; Current GPA Analysis & Module Tracker')

    st.markdown('This feature allows NTU/SUTD/SMU students to keep track of their modules to calculate their CAP and provides a brief analysis on the modules taken. It also allows users to download their analysis as a PDF file and selected module data to an Excel file. For NUS students, please see the NUS-exclusive app [**here**](https://nus-cap.streamlit.app). ')

    st.markdown('---')

    if 'all_module_data' not in st.session_state:
        st.session_state['all_module_data'] = []

    if 'upload_status' not in st.session_state:
        st.session_state['upload_status'] = False

    # Unique grade-to-cap conversion for each uni
    ntu_grades_to_cap = {'A+': 5.0,
                         'A': 5.0,
                         'A-': 4.5, 
                         'B+': 4.0, 
                         'B': 3.5, 
                         'B-': 3.0, 
                         'C+': 2.5, 
                         'C': 2.0, 
                         'D+': 1.5, 
                         'D': 1.0, 
                         'F': 0.0, 
                         'S': None, 
                         'U': None,
                         'EX': None,
                         'TC': None,
                         'LOA': None,
                         'IP': None,
                         '*': None,
                         '#': None}
    
    smu_grades_to_cap = {'A+': 4.3,
                         'A': 4.0,
                         'A-': 3.7, 
                         'B+': 3.3, 
                         'B': 3.0, 
                         'B-': 2.7, 
                         'C+': 2.3, 
                         'C': 2.0,
                         'C-': 1.7, 
                         'D+': 1.3, 
                         'D': 1.0, 
                         'F': 0.0,
                         'W': None}

    sutd_grades_to_cap = {'A+': 5.3,
                          'A': 5.0,
                          'A-': 4.5, 
                          'B+': 4.0, 
                          'B': 3.5, 
                          'B-': 3.0, 
                          'C+': 2.5, 
                          'C': 2.0, 
                          'D+': 1.5, 
                          'D': 1.0, 
                          'F': 0.0}

    # Select unique grade_to_cap dictionary based on selected uni
    if selected_uni == 'NTU (Nanyang Technological University)':
        grades_to_cap = ntu_grades_to_cap
    elif selected_uni == 'SMU (Singapore Management University)':
        grades_to_cap = smu_grades_to_cap
    elif selected_uni == 'SUTD (Singapore University of Technology & Design)':
        grades_to_cap = sutd_grades_to_cap

    # Generic template for adding modules
    col_left, col_middle, col_right = st.columns(3)
    with col_left:
        mod_code = st.text_input('Input your module code here:')
    with col_middle:
        mod_title = st.text_input('Input your module title here:')
    with col_right:
        mod_mcs = st.number_input('Input module MCs/AUs:', value = 4.0)

    mod_grade = st.selectbox('Select grade you have obtained for the respective module:', grades_to_cap)
    mod_score = grades_to_cap[mod_grade]

    now = datetime.datetime.now()

    amb_col, rmb_col, clear_col = st.columns([1, 4.2, 0.8]) 

    with amb_col:
        amb = st.button('Add Module')
        if amb:
            st.session_state.all_module_data.append([mod_code, mod_title, mod_mcs, mod_grade, mod_score])

    with rmb_col:
        rmb = st.button(u'\u21ba')
        if rmb and st.session_state['all_module_data'] != []:
            st.session_state.all_module_data.remove(st.session_state.all_module_data[-1])

    with clear_col:
        clear = st.button('Clear All')
        if clear:
            st.session_state['all_module_data'] = []

    # Functionality to add mdoules to existing spreadsheet
    upload_xlsx = st.file_uploader('Or, upload an existing .xlsx file with recorded modules in the same format:')

    if upload_xlsx is not None and st.session_state['upload_status'] == False:
        df_upload = pd.read_excel(upload_xlsx)
        for row in range(len(df_upload)):
            st.session_state.all_module_data.append([i for i in df_upload.iloc[row]])
        st.session_state['upload_status'] = True

    elif upload_xlsx is None:
        st.session_state['upload_status'] = False

    df = pd.DataFrame(columns = ['Module Code', 'Module Title', 'No. of MC/AUs', 'Grade', 'Grade Points'],
                      data = st.session_state['all_module_data'])

    st.markdown('###### Add a module and grade to view and download the data table:')

    # Display module data in DataFrame
    if st.session_state['all_module_data'] != []:
        st.dataframe(df.style.format(precision = 1), use_container_width = True)
        
    analysis_col, export_col = st.columns([1, 0.265]) 

    with export_col:
        def to_excel(df):
            output = io.BytesIO()
            writer = pd.ExcelWriter(output, engine = 'xlsxwriter')

            df.to_excel(writer, sheet_name = 'nus_mods', index = False)
            workbook = writer.book
            worksheet = writer.sheets['nus_mods']

            # Add formats and templates here        
            font_color = '#000000'
            header_color = '#ffff00'

            string_template = workbook.add_format(
                {
                    'font_color': font_color, 
                }
            )

            grade_template = workbook.add_format(
                {
                    'font_color': font_color, 
                    'align': 'center',
                    'bold': True
                }
            )

            float_template = workbook.add_format(
                {
                    'num_format': '0.0',
                    'font_color': font_color, 
                }
            )

            header_template = workbook.add_format(
                {
                    'bg_color': header_color, 
                    'border': 1
                }
            )

            column_formats = {
                'A': [string_template, 15],
                'B': [string_template, 50],
                'C': [float_template, 15],
                'D': [grade_template, 15],
                'E': [float_template, 15]
            }

            for column in column_formats.keys():
                worksheet.set_column(f'{column}:{column}', column_formats[column][1], column_formats[column][0])
                worksheet.conditional_format(f'{column}1:{column}1', {'type': 'no_errors', 'format': header_template})

            # Automatically apply Filter function on shape of dataframe
            worksheet.autofilter(0, 0, df.shape[0], df.shape[1]-1)

            # Saving and returning data
            writer.save()
            processed_data = output.getvalue()

            return processed_data

        def get_table_download_link(df):
            """Generates a link allowing the data in a given Pandas DataFrame to be downloaded
            in:  dataframe
            out: href string
            """
            val = to_excel(df)
            b64 = base64.b64encode(val)

            return f'<a href="data:application/octet-stream;base64,{b64.decode()}" download="mod_cap_details.xlsx">:inbox_tray: Download (.xlsx)</a>' 

        if st.session_state['all_module_data'] != []:
            st.markdown(get_table_download_link(df), unsafe_allow_html = True)

    with analysis_col:
        if st.session_state['all_module_data'] != []:
            analysis = st.button('View Analysis')
        else:
            analysis = None

    if analysis and st.session_state['all_module_data'] != []:

        df2 = df.dropna()
        cap = sum(df2['No. of MC/AUs'] * df2['Grade Points']) / sum(df2['No. of MC/AUs'])
        total_mcs_gpa = sum(df2['No. of MC/AUs'])

        final_cap = round(cap, 2)
        dp4_cap = round(cap, 4)

        # Degree classification
        if selected_uni == 'NTU (Nanyang Technological University)' or selected_uni == 'SUTD (Singapore University of Technology & Design)':
            if final_cap >= 4.50:
                degree_class = 'Honours (Highest Distinction)'
            elif final_cap >= 4.00:
                degree_class = 'Honours (Distinction)'
            elif final_cap >= 3.50:
                degree_class = 'Honours (Merit)'            
            elif final_cap >= 3.00:
                degree_class = 'Honours'        
            elif final_cap >= 2.00:
                degree_class = 'Pass'
            else:
                degree_class = 'Below requirements for graduation'

        elif selected_uni == 'SMU (Singapore Management University)':
            if final_cap >= 3.8:
                degree_class = 'Summa Cum Laude'
            elif final_cap >= 3.6:
                degree_class = 'Magna Cum Laude'
            elif final_cap >= 3.4:
                degree_class = 'Cum Laude'            
            elif final_cap >= 3.2:
                degree_class = 'High Merit'        
            elif final_cap >= 3.0:
                degree_class = 'Merit'
            elif final_cap >= 2.5:
                degree_class = 'Pass'
            else:
                degree_class = 'Below requirements for graduation'

        # Cumulative modules
        table_dict = {'Final GPA': final_cap,
                      'Degree Classification': degree_class,
                      'Your GPA (To 4 d.p.)': dp4_cap,
                      'No. of MC/AUs used to calculate GPA': total_mcs_gpa,
                      'Date of Overview': now.strftime('%d %b %Y')}

        col_fill_colors = ['azure']*2 + ['lavender']*2 + ['honeydew']
        font_colors = ['mediumblue']*2 + ['indigo']*2 + ['darkgreen']

        fig = go.Figure(data = [go.Table(columnwidth = [2.5, 1.5],
                                    header = dict(values = ['<b>Module Overview & Detailed Analysis<b>', 
                                                            '<b>Result<b>'],
                                                fill_color = 'lightskyblue',
                                                line_color = 'black',
                                                align = 'center',
                                                font = dict(color = 'black', 
                                                            size = 14,
                                                            family = 'Georgia')),
                                    cells = dict(values = [list(table_dict.keys()),
                                                        list(table_dict.values())], 
                                                fill_color = [col_fill_colors, col_fill_colors],
                                                line_color = 'black',
                                                align = ['right', 'left'],
                                                font = dict(color = [font_colors, font_colors], 
                                                            size = [14, 14],
                                                            family = ['Georgia', 'Georgia Bold']),
                                                height = 25))])

        fig.update_layout(height = 167.5, width = 700, margin = dict(l = 5, r = 5, t = 5, b = 5))
        st.plotly_chart(fig, use_container_width = True)

        # Create an in-memory buffer
        buffer = io.BytesIO()

        # Save the figure as a pdf to the buffer
        fig.write_image(file = buffer, scale = 6, format = 'pdf')

        # Download the pdf from the buffer
        st.download_button(
            label = 'Download Analysis as PDF',
            data = buffer,
            file_name = 'cap_overview.pdf',
            mime = 'application/octet-stream',
            help = 'Downloads the module analysis as a PDF File'
        )

    st.markdown('---')
                       
def explain():
    st.markdown('#### :bulb: &nbsp; CAP/GPA Calculation Explanation')

    st.markdown('This feature describes how Cumulative Average Point (CAP) or Grade Point Average (GPA) is calculated briefly.')

    st.markdown('---')
    
    st.markdown('To calculate the CAP/GPA for $n$ number of modules:')
    st.write('&nbsp;')
    
    st.markdown(r'''$G = \text{Module Grade Points}$''')
    st.markdown(r'''$G_n = \text{Specific Module Grade Points for the } n^\text{th} \text{ module used in CAP/GPA calculation}$''')
    st.markdown(r'''$MC = \text{Module Credits/Academic Units}$''')
    st.markdown(r'''$MC_n = \text{Specific Module Credits/Academic Units for the } n^\text{th} \text{ module used in CAP/GPA calculation}$''')
    
    st.write('&nbsp;')
    
    st.latex(r'''\text{CAP} = \frac{G_1\times{MC_1} + G_2\times{MC_2} + ... + G_n\times{MC_n}}{MC_1 + MC_2 + ... + MC_n}''')
   
    st.latex(r'''= \sum\limits_{i=1}^{n} \frac{{G_i}\times{MC_i}}{MC_i}''')
    
    st.markdown('---')
        
    
if __name__ == "__main__":
    st.set_page_config(page_title = 'GPA Calculator', page_icon = '📝')
    main()

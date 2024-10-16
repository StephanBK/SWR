import pandas as pd
import streamlit as st
from datetime import datetime
import xlsxwriter
from io import BytesIO

# Conversion factors
inches_to_mm = 25.4
sq_inches_to_sq_feet = 1 / 144

st.title("Glass Processing and Export App")

# File upload
uploaded_file = st.file_uploader("Upload a CSV file", type="csv")
if uploaded_file is not None:
    df = pd.read_csv(uploaded_file)
    st.write("Here's a preview of your data:")
    st.dataframe(df.head())
    
    # System Type selection
    system_type = st.selectbox("Select System Type", ["SWR-IG", "SWR-VIG", "SWR", "Other"])
    if system_type == "SWR-IG" or system_type == "SWR-VIG":
        glass_offset = 11.1125
    elif system_type == "SWR":
        glass_offset = 7.571
    else:
        glass_offset = st.number_input("Enter Glass Offset (in inches)", value=0.0)

    # Project details inputs
    project_name = st.text_input("Enter Project Name")
    project_number = st.text_input("Enter Project Number")
    part_number = st.text_input("Enter Part #")
    
    # Dimension inputs
    glass_cutting_tolerance = st.number_input("Enter Glass Cutting Tolerance (in inches)", value=0.0)
    joint_top = st.number_input("Enter Joint Top (in inches)", value=0.0)
    joint_bottom = st.number_input("Enter Joint Bottom (in inches)", value=0.0)
    joint_left = st.number_input("Enter Joint Left (in inches)", value=0.0)
    joint_right = st.number_input("Enter Joint Right (in inches)", value=0.0)
    
    # Get the current date for file headers
    creation_date = datetime.now().strftime('%Y-%m-%d')
    
    # Display entered values for confirmation
    st.write(f"**System Type:** {system_type}")
    st.write(f"**Project Name:** {project_name}")
    st.write(f"**Project Number:** {project_number}")
    st.write(f"**Part #:** {part_number}")
    st.write(f"**Glass Offset:** {glass_offset}")
    st.write(f"**Glass Cutting Tolerance:** {glass_cutting_tolerance}")
    st.write(f"**Joint Dimensions (in inches):** Top={joint_top}, Bottom={joint_bottom}, Left={joint_left}, Right={joint_right}")
    
    # Calculations for Glass dimensions
    df['Overall Width mm'] = df['Overall Width in'] * inches_to_mm
    df['Overall Height mm'] = df['Overall Height in'] * inches_to_mm
    df['Unit Area ft²'] = (df['Overall Width in'] * df['Overall Height in']) * sq_inches_to_sq_feet
    df['Total Area ft²'] = df['Unit Area ft²'] * df['Qty']

    joint_left_mm = joint_left * inches_to_mm
    joint_right_mm = joint_right * inches_to_mm
    joint_top_mm = joint_top * inches_to_mm
    joint_bottom_mm = joint_bottom * inches_to_mm
    df['SWR Width mm'] = df['Overall Width mm'] - joint_left_mm - joint_right_mm
    df['SWR Height mm'] = df['Overall Height mm'] - joint_top_mm - joint_bottom_mm
    mm_to_inches = 1 / inches_to_mm
    df['SWR Width in'] = df['SWR Width mm'] * mm_to_inches
    df['SWR Height in'] = df['SWR Height mm'] * mm_to_inches
    
    # Glass Offset calculation
    glass_offset_mm = glass_offset * inches_to_mm
    df['Glass Width mm'] = df['SWR Width mm'] - (2 * glass_offset_mm)
    df['Glass Height mm'] = df['SWR Height mm'] - (2 * glass_offset_mm)
    df['Glass Width in'] = df['Glass Width mm'] * mm_to_inches
    df['Glass Height in'] = df['Glass Height mm'] * mm_to_inches

    # Button to generate and download Excel files
    if st.button("Generate Excel Files"):
        # Buffer for writing the Excel files
        excel_buffer = BytesIO()

        # Writing to Excel
        with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='Sheet1', index=False)
            
            # Set up the workbook and worksheet
            workbook = writer.book
            worksheet = writer.sheets['Sheet1']
            worksheet.write('A1', "Project Name:")
            worksheet.write('B1', project_name)
            worksheet.write('A2', "Project Number:")
            worksheet.write('B2', project_number)
            worksheet.write('A3', "Date Created:")
            worksheet.write('B3', creation_date)
            
        # Download the file
        excel_buffer.seek(0)
        st.download_button(
            label="Download Glass Excel File",
            data=excel_buffer,
            file_name="Glass.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

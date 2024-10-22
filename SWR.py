
import streamlit as st
import pandas as pd
from datetime import datetime
import xlsxwriter
from io import BytesIO

# Display the logo at the top of the app
st.image("ilogo.png", width=200)  # Adjust the width as needed

# Title for the app
st.title("SWR Cutlist")

# Project details input fields
project_name = st.text_input("Enter Project Name")
project_number = st.text_input("Enter Project Number")

# System Type selection with automatic Glass Offset logic
system_type = st.selectbox("Select System Type", ["SWR-IG", "SWR-VIG", "SWR", "Custom"])

# Finish selection
finish = st.selectbox("Select Finish", ["Mil Finish", "Clear Anodized", "Black Anodized", "Painted"])

# Set default Glass Offset and assign profile number based on the selected system type
if system_type == "SWR-IG":
    glass_offset = 11.1125
    profile_number = '03003'
elif system_type == "SWR-VIG":
    glass_offset = 11.1125
    profile_number = '03004'
elif system_type == "SWR":
    glass_offset = 7.571
    profile_number = '03002'
else:
    glass_offset = st.number_input("Enter Glass Offset (in inches)", value=0.0)
    profile_number = None

# Display the automatically set Glass Offset for confirmation or modification if necessary
if system_type != "Custom":
    st.write(f"Using a Glass Offset of {glass_offset} inches for system type {system_type} with profile number {profile_number}")

# Additional project details input fields with default values and 3 decimal places
glass_cutting_tolerance = st.number_input("Enter Glass Cutting Tolerance (in inches)", value=0.625, format="%.3f")
joint_top = st.number_input("Enter the Joint Top (in inches)", value=0.5, format="%.3f")
joint_bottom = st.number_input("Enter the Joint Bottom (in inches)", value=0.125, format="%.3f")
joint_left = st.number_input("Enter the Joint Left (in inches)", value=0.25, format="%.3f")
joint_right = st.number_input("Enter the Joint Right (in inches)", value=0.25, format="%.3f")

# Combine System Type, Project Number, and Finish
part_number = f"{system_type}-{profile_number}"

# File upload
uploaded_file = st.file_uploader("Upload a CSV file", type="csv")

# Provide a download button for the template file
with open("SWR template.csv", "rb") as template_file:
    template_data = template_file.read()
    st.download_button("Download Template", data=template_data, file_name="SWR_template.csv", mime="text/csv")

# Load and process the uploaded CSV file
if uploaded_file:
    df = pd.read_csv(uploaded_file)

    # Perform some operations on the DataFrame
    st.dataframe(df)

    # Function to add the INOVUES logo to the Excel files
    def add_logo(worksheet, writer):
        worksheet.insert_image('A1', 'ilogo.png', {'x_scale': 0.25, 'y_scale': 0.25})  # Quarter size logo

    # Function to write the project details at row 8
    def write_project_details(worksheet):
        worksheet.write('A8', "Project Name:")
        worksheet.write('A9', "Project Number:")
        worksheet.write('A10', "Date Created:")
        worksheet.write('B8', project_name)
        worksheet.write('B9', project_number)
        worksheet.write('B10', datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

    # Function to write the data starting 11 rows down to accommodate the logo and project details
    def write_data_with_offset(writer, sheet_name):
        df.to_excel(writer, sheet_name=sheet_name, startrow=11, index=False)

    # Example: Generate Glass File Export with logo and project details
    glass_file = BytesIO()
    with pd.ExcelWriter(glass_file, engine='xlsxwriter') as writer:
        write_data_with_offset(writer, 'Glass')
        worksheet = writer.sheets['Glass']
        add_logo(worksheet, writer)
        write_project_details(worksheet)

    # Example: Generate AggCutOnly Export with logo and project details
    agg_file = BytesIO()
    with pd.ExcelWriter(agg_file, engine='xlsxwriter') as writer:
        write_data_with_offset(writer, 'AggCutOnly')
        worksheet = writer.sheets['AggCutOnly']
        add_logo(worksheet, writer)
        write_project_details(worksheet)

    # Example: Generate TagDetails Export with "Color/Finish" column, logo, and project details
    tag_file = BytesIO()
    with pd.ExcelWriter(tag_file, engine='xlsxwriter') as writer:
        # Add a new column "Color/Finish" with the value of the finish variable
        df['Color/Finish'] = finish
        write_data_with_offset(writer, 'TagDetails')
        worksheet = writer.sheets['TagDetails']
        add_logo(worksheet, writer)
        write_project_details(worksheet)

    # Example: Generate SWR Table Export with logo and project details
    swr_table_file = BytesIO()
    with pd.ExcelWriter(swr_table_file, engine='xlsxwriter') as writer:
        worksheet = writer.book.add_worksheet()
        write_project_details(worksheet)
        write_data_with_offset(writer, 'Sheet1')
        worksheet = writer.sheets['Sheet1']
        add_logo(worksheet, writer)

    # Provide download buttons for each file
    st.download_button("Download Glass File", data=glass_file.getvalue(), file_name="Glass.xlsx")
    st.download_button("Download AggCutOnly File", data=agg_file.getvalue(), file_name="AggCutOnly.xlsx")
    st.download_button("Download TagDetails File", data=tag_file.getvalue(), file_name="TagDetails.xlsx")
    st.download_button("Download SWR Table File", data=swr_table_file.getvalue(), file_name="SWR_table.xlsx")

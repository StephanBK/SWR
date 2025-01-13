import streamlit as st
import pandas as pd
from datetime import datetime
import xlsxwriter
from io import BytesIO
import os

# Conversion constants
inches_to_mm = 25.4
mm_to_inches = 1 / inches_to_mm
sq_inches_to_sq_feet = 1 / 144

# Display the logo at the top of the app
st.image("ilogo.png", width=200)  # Adjust the width as needed

# Title for the app
st.title("SWR Cutlist")

# Project details input fields
project_name = st.text_input("Enter Project Name", value="INO-")
project_number = st.text_input("Enter Project Number")

# System Type selection with automatic Glass Offset logic
system_type = st.selectbox("Select System Type", ["SWR-IG", "SWR-VIG", "SWR", "Custom"])

# Finish selection
finish = st.selectbox("Select Finish", ["Mil Finish", "Clear Anodized", "Black Anodized", "Painted"])

# Set default Glass Offset and assign profile number based on the selected system type
if system_type == "SWR-IG":
    glass_offset = 11.1125  # in millimeters
    profile_number = '03003'
elif system_type == "SWR-VIG":
    glass_offset = 11.1125  # in millimeters
    profile_number = '03004'
elif system_type == "SWR":
    glass_offset = 7.571  # in millimeters
    profile_number = '03002'
else:
    # Glass Offset toggle and input
    unit_offset = st.radio("Select Unit for Glass Offset", ["Inches", "Millimeters"], index=0)
    if unit_offset == "Inches":
        glass_offset = st.number_input("Enter Glass Offset (in inches)", value=0.0) * inches_to_mm
    else:
        glass_offset = st.number_input("Enter Glass Offset (in mm)", value=0.0)

# Display the automatically set Glass Offset for confirmation or modification if necessary
if system_type != "Custom":
    st.write(f"Using a Glass Offset of {glass_offset} mm for system type {system_type} with profile number {profile_number}")

# Input fields for Glass Cutting Tolerance and Joint Dimensions with toggles
st.subheader("Input Parameters")

# Glass Cutting Tolerance
unit_tolerance = st.radio("Select Unit for Glass Cutting Tolerance", ["Inches", "Millimeters"], index=0)
if unit_tolerance == "Inches":
    glass_cutting_tolerance = st.number_input("Enter Glass Cutting Tolerance", value=0.0625, format="%.4f")
else:
    glass_cutting_tolerance = st.number_input("Enter Glass Cutting Tolerance (in mm)", value=0.0625 * inches_to_mm, format="%.3f")
    glass_cutting_tolerance = glass_cutting_tolerance * mm_to_inches

# Joint Top
unit_joint_top = st.radio("Select Unit for Joint Top", ["Inches", "Millimeters"], index=0)
if unit_joint_top == "Inches":
    joint_top = st.number_input("Enter the Joint Top", value=0.5, format="%.3f")
else:
    joint_top = st.number_input("Enter the Joint Top (in mm)", value=0.5 * inches_to_mm, format="%.3f")
    joint_top = joint_top * mm_to_inches

# Joint Bottom
unit_joint_bottom = st.radio("Select Unit for Joint Bottom", ["Inches", "Millimeters"], index=0)
if unit_joint_bottom == "Inches":
    joint_bottom = st.number_input("Enter the Joint Bottom", value=0.125, format="%.3f")
else:
    joint_bottom = st.number_input("Enter the Joint Bottom (in mm)", value=0.125 * inches_to_mm, format="%.3f")
    joint_bottom = joint_bottom * mm_to_inches

# Joint Left
unit_joint_left = st.radio("Select Unit for Joint Left", ["Inches", "Millimeters"], index=0)
if unit_joint_left == "Inches":
    joint_left = st.number_input("Enter the Joint Left", value=0.25, format="%.3f")
else:
    joint_left = st.number_input("Enter the Joint Left (in mm)", value=0.25 * inches_to_mm, format="%.3f")
    joint_left = joint_left * mm_to_inches

# Joint Right
unit_joint_right = st.radio("Select Unit for Joint Right", ["Inches", "Millimeters"], index=0)
if unit_joint_right == "Inches":
    joint_right = st.number_input("Enter the Joint Right", value=0.25, format="%.3f")
else:
    joint_right = st.number_input("Enter the Joint Right (in mm)", value=0.25 * inches_to_mm, format="%.3f")
    joint_right = joint_right * mm_to_inches

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

    # Convert all input measurements to mm for internal calculations
    df['Overall Width mm'] = df['Overall Width in'] * inches_to_mm
    df['Overall Height mm'] = df['Overall Height in'] * inches_to_mm

    # Convert joint dimensions to mm for subtraction
    joint_left_mm = joint_left * inches_to_mm
    joint_right_mm = joint_right * inches_to_mm
    joint_top_mm = joint_top * inches_to_mm
    joint_bottom_mm = joint_bottom * inches_to_mm

    # SWR Width/Height calculations in mm (after subtracting joint dimensions)
    df['SWR Width mm'] = df['Overall Width mm'] - joint_left_mm - joint_right_mm
    df['SWR Height mm'] = df['Overall Height mm'] - joint_top_mm - joint_bottom_mm

    # Convert SWR dimensions from mm back to inches immediately after calculation
    df['SWR Width in'] = df['SWR Width mm'] * mm_to_inches
    df['SWR Height in'] = df['SWR Height mm'] * mm_to_inches

    # Glass Offset calculation in mm (subtracted from SWR dimensions without further conversion)
    df['Glass Width mm'] = df['SWR Width mm'] - (2 * glass_offset)
    df['Glass Height mm'] = df['SWR Height mm'] - (2 * glass_offset)

    # Convert final glass dimensions back to inches for the Glass file output
    df['Glass Width in'] = df['Glass Width mm'] * mm_to_inches
    df['Glass Height in'] = df['Glass Height mm'] * mm_to_inches

    # ==================== Glass File Export ====================
    output_df = pd.DataFrame({'Item': range(1, len(df) + 1)})
    output_df['Glass Width in'] = df['Glass Width in']
    output_df['Glass Height in'] = df['Glass Height in']
    output_df['Area Each (ft²)'] = (output_df['Glass Width in'] * output_df['Glass Height in']) * sq_inches_to_sq_feet
    output_df['Qty'] = df['Qty']
    output_df['Area Total (ft²)'] = output_df['Qty'] * output_df['Area Each (ft²)']
    totals_row = pd.DataFrame([['Totals', None, None, None, output_df['Qty'].sum(), output_df['Area Total (ft²)'].sum()]],
                              columns=output_df.columns)
    output_df = pd.concat([output_df, totals_row], ignore_index=True)

    # Save to Excel and prepare for download
    glass_file = BytesIO()
    with pd.ExcelWriter(glass_file, engine='xlsxwriter') as writer:
        worksheet = writer.book.add_worksheet("Sheet1")
        
        # Insert logo
        worksheet.insert_image('A1', 'ilogo.png', {'x_scale': 0.2, 'y_scale': 0.2})
        
        worksheet.write('A7', "Project Name:")
        worksheet.write('A8', "Project Number:")
        worksheet.write('A9', "Date Created:")
        worksheet.write('B7', project_name)
        worksheet.write('B8', project_number)
        worksheet.write('B9', datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        output_df.to_excel(writer, sheet_name='Sheet1', startrow=12, index=False)

    # ==================== AggCutOnly File Export ====================
    df['Qty x 2'] = df['Qty'] * 2
    width_counts = df.groupby('SWR Width in')['Qty'].sum().sort_values(ascending=False)
    height_counts = df.groupby('SWR Height in')['Qty'].sum().sort_values(ascending=False)
    unique_dimensions = pd.Index(width_counts.index.tolist() + height_counts.index.tolist()).unique()
    
    agg_df = pd.DataFrame(0, index=unique_dimensions, columns=['Part #', 'Miter'] + df['Tag'].unique().tolist() + ['Total QTY'])
    agg_df['Part #'] = part_number
    agg_df['Miter'] = "**"
    
    for i, row in df.iterrows():
        width, height, tag, qty_x_2 = row['SWR Width in'], row['SWR Height in'], row['Tag'], row['Qty x 2']
        if width in agg_df.index and tag in agg_df.columns:
            agg_df.at[width, tag] += qty_x_2
        if height in agg_df.index and tag in agg_df.columns:
            agg_df.at[height, tag] += qty_x_2
    
    agg_df['Total QTY'] = agg_df[df['Tag'].unique()].sum(axis=1)
    agg_df.index.name = "Finished Length in"
    agg_df = agg_df.reset_index()
    
    agg_file = BytesIO()
    with pd.ExcelWriter(agg_file, engine='xlsxwriter') as writer:
        worksheet = writer.book.add_worksheet('Sheet1')
        
        # Insert logo
        worksheet.insert_image('A1', 'ilogo.png', {'x_scale': 0.2, 'y_scale': 0.2})
        
        worksheet.write('A7', "Project Name:")
        worksheet.write('A8', "Project Number:")
        worksheet.write('A9', "Date Created:")
        worksheet.write('B7', project_name)
        worksheet.write('B8', project_number)
        worksheet.write('B9', datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        agg_df.to_excel(writer, sheet_name='Sheet1', startrow=12, index=False)

    # ==================== TagDetails File Export ====================
    tag_file = BytesIO()
    with pd.ExcelWriter(tag_file, engine='xlsxwriter') as writer:
        for tag in df['Tag'].unique():
            tag_df = df[df['Tag'] == tag]
            table_data = {'Item': [], 'Position': [], 'Quantity': [], 'Length (mm)': [], 'Length (inch)': []}
            for idx, row in tag_df.iterrows():
                swr_width_mm, swr_height_mm, swr_width_in, swr_height_in = row['SWR Width mm'], row['SWR Height mm'], row['SWR Width in'], row['SWR Height in']
                qty_x2 = row['Qty'] * 2
                table_data['Item'].extend([idx + 1, idx + 1, idx + 1, idx + 1])
                table_data['Position'].extend(['left', 'right', 'top', 'bottom'])
                table_data['Quantity'].extend([qty_x2, qty_x2, qty_x2, qty_x2])
                table_data['Length (mm)'].extend([swr_width_mm, swr_width_mm, swr_height_mm, swr_height_mm])
                table_data['Length (inch)'].extend([swr_width_in, swr_width_in, swr_height_in, swr_height_in])
            tag_output_df = pd.DataFrame(table_data)
            worksheet = writer.book.add_worksheet(str(tag))
            
            # Insert logo
            worksheet.insert_image('A1', 'ilogo.png', {'x_scale': 0.2, 'y_scale': 0.2})
            
            worksheet.write('A7', "Project Name:")
            worksheet.write('A8', "Project Number:")
            worksheet.write('A9', "Date Created:")
            worksheet.write('B7', project_name)
            worksheet.write('B8', project_number)
            worksheet.write('B9', datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
            tag_output_df.to_excel(writer, sheet_name=str(tag), startrow=12, index=False)

    # ==================== SWR Table Export ====================
    swr_table_file = BytesIO()
    with pd.ExcelWriter(swr_table_file, engine='xlsxwriter') as writer:
        worksheet = writer.book.add_worksheet()
        
        # Insert logo
        worksheet.insert_image('A1', 'ilogo.png', {'x_scale': 0.2, 'y_scale': 0.2})
        
        worksheet.write('A7', "Project Name:")
        worksheet.write('A8', "Project Number:")
        worksheet.write('A9', "Date Created:")
        worksheet.write('B7', project_name)
        worksheet.write('B8', project_number)
        worksheet.write('B9', datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        df.to_excel(writer, sheet_name='Sheet1', startrow=12, index=False)

    # Provide download buttons for each file with the updated filenames
    prefix = f"INO_{project_number}_SWR_"
    st.download_button("Download Glass File", data=glass_file.getvalue(), file_name=f"{prefix}Glass.xlsx")
    st.download_button("Download AggCutOnly File", data=agg_file.getvalue(), file_name=f"{prefix}AggCutOnly.xlsx")
    st.download_button("Download TagDetails File", data=tag_file.getvalue(), file_name=f"{prefix}TagDetails.xlsx")
    st.download_button("Download SWR Table File", data=swr_table_file.getvalue(), file_name=f"{prefix}Table.xlsx")
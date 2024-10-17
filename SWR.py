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
profile_number = st.number_input("Enter Profile Number", step=1)
# System Type selection with automatic Glass Offset logic
system_type = st.selectbox("Select System Type", ["SWR-IG", "SWR-VIG", "SWR", "Custom"])

# Set default Glass Offset based on the selected system type
if system_type in ["SWR-IG", "SWR-VIG"]:
    glass_offset = 11.1125
elif system_type == "SWR":
    glass_offset = 7.571
else:
    glass_offset = st.number_input("Enter Glass Offset (in inches)", value=0.0)
    
# Display the automatically set Glass Offset for confirmation or modification if necessary
if system_type != "Custom":
    st.write(f"Using a Glass Offset of {glass_offset} inches for system type {system_type}")

# Additional project details input fields with default values and 3 decimal places
glass_cutting_tolerance = st.number_input("Enter Glass Cutting Tolerance (in inches)", value=0.625, format="%.3f")
joint_top = st.number_input("Enter the Joint Top (in inches)", value=0.5, format="%.3f")
joint_bottom = st.number_input("Enter the Joint Bottom (in inches)", value=0.125, format="%.3f")
joint_left = st.number_input("Enter the Joint Left (in inches)", value=0.25, format="%.3f")
joint_right = st.number_input("Enter the Joint Right (in inches)", value=0.25, format="%.3f")

# Combine System Type and Project Number

part_number = f"{system_type}-{profile_number}"

# File upload
uploaded_file = st.file_uploader("Upload a CSV file", type="csv")

# Provide a download button for the template file
with open("SWR template.csv", "rb") as template_file:
    template_data = template_file.read()
st.download_button(
    label="Download Template File",
    data=template_data,
    file_name="SWR_template.csv",
    mime="text/csv"
)

# Date creation for headers
creation_date = datetime.now().strftime('%Y-%m-%d')

if uploaded_file is not None:
    # Load the CSV data
    df = pd.read_csv(uploaded_file)
    st.write("Here’s a preview of your data:")
    st.dataframe(df.head())

    # Conversion factors
    inches_to_mm = 25.4
    sq_inches_to_sq_feet = 1 / 144

    # Example calculation for Glass dimensions
    df['Overall Width mm'] = df['Overall Width in'] * inches_to_mm
    df['Overall Height mm'] = df['Overall Height in'] * inches_to_mm
    df['Unit Area ft²'] = (df['Overall Width in'] * df['Overall Height in']) * sq_inches_to_sq_feet
    df['Total Area ft²'] = df['Unit Area ft²'] * df['Qty']

    # Calculate joint dimensions in mm
    joint_left_mm = joint_left * inches_to_mm
    joint_right_mm = joint_right * inches_to_mm
    joint_top_mm = joint_top * inches_to_mm
    joint_bottom_mm = joint_bottom * inches_to_mm

    # SWR Width/Height calculations
    df['SWR Width mm'] = df['Overall Width mm'] - joint_left_mm - joint_right_mm
    df['SWR Height mm'] = df['Overall Height mm'] - joint_top_mm - joint_bottom_mm
    mm_to_inches = 1 / inches_to_mm
    df['SWR Width in'] = df['SWR Width mm'] * mm_to_inches
    df['SWR Height in'] = df['SWR Height mm'] * mm_to_inches

    # Glass Offset calculation
    glass_offset_mm = glass_offset * inches_to_mm
    df['Glass Width mm'] = df['SWR Width mm'] - (2 * glass_offset_mm)
    df['Glass Height mm'] = df['SWR Height mm'] - (2 * glass_offset_mm)

    # Define the rounding function
    def round_to_nearest(value, base=0.0625):
        return round(value / base) * base

    # Calculate Glass Width and Height in inches and round to nearest 0.0625
    df['Glass Width in'] = df['Glass Width mm'] * mm_to_inches
    df['Glass Height in'] = df['Glass Height mm'] * mm_to_inches
    df['Glass Width in'] = df['Glass Width in'].apply(round_to_nearest)
    df['Glass Height in'] = df['Glass Height in'].apply(round_to_nearest)

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
        worksheet.write('A1', "Project Name:")
        worksheet.write('A2', "Project Number:")
        worksheet.write('A3', "Date Created:")
        worksheet.write('B1', project_name)
        worksheet.write('B2', part_number)
        worksheet.write('B3', creation_date)
        output_df.to_excel(writer, sheet_name='Sheet1', startrow=6, index=False)

    # ==================== AggCutOnly File Export ====================
    df['Qty x 2'] = df['Qty'] * 2
    width_counts = df.groupby('SWR Width in')['Qty'].sum().sort_values(ascending=False)
    height_counts = df.groupby('SWR Height in')['Qty'].sum().sort_values(ascending=False)
    unique_dimensions = pd.Index(width_counts.index.tolist() + height_counts.index.tolist()).unique()

    # Prepare the AggCutOnly DataFrame
    agg_df = pd.DataFrame(0, index=unique_dimensions, columns=['Part #', 'Miter', 'Finished Length in'] + df['Tag'].unique().tolist() + ['Total QTY'])

    # Set all values in the 'Part #' column to the new part_number
    agg_df['Part #'] = part_number  # This ensures all rows have the new part number

    agg_df['Miter'] = "**"
    agg_df['Finished Length in'] = agg_df.index

    # Process each row to populate the tags
    for i, row in df.iterrows():
        width, height, tag, qty_x_2 = row['SWR Width in'], row['SWR Height in'], row['Tag'], row['Qty x 2']
        if width in agg_df.index and tag in agg_df.columns:
            agg_df.at[width, tag] += qty_x_2
        if height in agg_df.index and tag in agg_df.columns:
            agg_df.at[height, tag] += qty_x_2

    # Sum quantities across all tags
    agg_df['Total QTY'] = agg_df[df['Tag'].unique()].sum(axis=1)

    # Add a totals row for the AggCutOnly file
    totals_row = pd.DataFrame([{
        'Part #': None,
        'Miter': None,
        'Finished Length in': 'Total',
        **{col: agg_df[col].sum() for col in agg_df.columns if col not in ['Part #', 'Miter', 'Finished Length in']}
    }])
    agg_df = pd.concat([agg_df, totals_row], ignore_index=True)

    agg_file = BytesIO()
    with pd.ExcelWriter(agg_file, engine='xlsxwriter') as writer:
        worksheet = writer.book.add_worksheet('Sheet1')
        worksheet.write('A1', "Project Name:")
        worksheet.write('A2', "Project Number:")
        worksheet.write('A3', "Date Created:")
        worksheet.write('B1', project_name)
        worksheet.write('B2', part_number)  # Use part_number here as well
        worksheet.write('B3', creation_date)
        agg_df.to_excel(writer, sheet_name='Sheet1', startrow=6, index=False)
    # ==================== TagDetails File Export ====================
    tag_file = BytesIO()
    with pd.ExcelWriter(tag_file, engine='xlsxwriter') as writer:
        for tag in df['Tag'].unique():
            tag_df = df[df['Tag'] == tag]
        
            # Initialize the table data with new columns for Type and Profile #
            table_data = {
                'Item': [],
                'Type': [],
                'Profile #': [],
                'Position': [],
                'Quantity': [],
                'Length (mm)': [],
                'Length (inch)': []
            }
        
            for idx, row in tag_df.iterrows():
                swr_width_mm, swr_height_mm, swr_width_in, swr_height_in = row['SWR Width mm'], row['SWR Height mm'], row['SWR Width in'], row['SWR Height in']
                qty_x2 = row['Qty'] * 2

                # Populate Item, Type, and Profile # for each position (left, right, top, bottom)
                table_data['Item'].extend([idx + 1, idx + 1, idx + 1, idx + 1])
                table_data['Type'].extend(['Alum Profile'] * 4)
                table_data['Profile #'].extend([part_number] * 4)  # Set Profile # to part_number
                table_data['Position'].extend(['left', 'right', 'top', 'bottom'])
                table_data['Quantity'].extend([qty_x2, qty_x2, qty_x2, qty_x2])
                table_data['Length (mm)'].extend([swr_width_mm, swr_width_mm, swr_height_mm, swr_height_mm])
                table_data['Length (inch)'].extend([swr_width_in, swr_width_in, swr_height_in, swr_height_in])

            # Create a DataFrame from table_data and export each tag to a separate sheet
            tag_output_df = pd.DataFrame(table_data)
            worksheet = writer.book.add_worksheet(str(tag))
            worksheet.write('A1', "Project Name:")
            worksheet.write('A2', "Project Number:")
            worksheet.write('A3', "Date Created:")
            worksheet.write('B1', project_name)
            worksheet.write('B2', part_number)  # Use part_number here for project reference
            worksheet.write('B3', creation_date)
            tag_output_df.to_excel(writer, sheet_name=str(tag), startrow=6, index=False)

    # ==================== SWR Table Export ====================
    swr_table_file = BytesIO()
    with pd.ExcelWriter(swr_table_file, engine='xlsxwriter') as writer:
        worksheet = writer.book.add_worksheet()
        worksheet.write('A1', "Project Name:")
        worksheet.write('A2', "Project Number:")
        worksheet.write('A3', "Date Created:")
        worksheet.write('B1', project_name)
        worksheet.write('B2', project_number)
        worksheet.write('B3', creation_date)
        df.to_excel(writer, sheet_name='Sheet1', startrow=6, index=False)

    # Provide download buttons for each file
    st.download_button("Download Glass File", data=glass_file.getvalue(), file_name="Glass.xlsx")
    st.download_button("Download AggCutOnly File", data=agg_file.getvalue(), file_name="AggCutOnly.xlsx")
    st.download_button("Download TagDetails File", data=tag_file.getvalue(), file_name="TagDetails.xlsx")
    st.download_button("Download SWR Table File", data=swr_table_file.getvalue(), file_name="SWR_table.xlsx")
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
st.image("ilogo.png", width=200)

# Title for the app
st.title("SWR Cutlist")

# Project details input fields
project_name = st.text_input("Enter Project Name")
project_number = st.text_input("Enter Project Number", value="INO-")
prepared_by = st.text_input("Prepared By")  # ← new field

# System Type selection with automatic Glass Offset logic
system_type = st.selectbox("Select System Type", ["SWR-IG", "SWR-VIG", "SWR", "Custom"])
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
    unit_offset = st.radio("Select Unit for Glass Offset", ["Inches", "Millimeters"], index=0)
    if unit_offset == "Inches":
        glass_offset = st.number_input("Enter Glass Offset (in inches)", value=0.0) * inches_to_mm
    else:
        glass_offset = st.number_input("Enter Glass Offset (in mm)", value=0.0)

if system_type != "Custom":
    st.write(f"Using a Glass Offset of {glass_offset:.3f} mm for system {system_type} (Profile {profile_number})")

# Input Parameters section
st.subheader("Input Parameters")

unit_tolerance = st.radio("Select Unit for Glass Cutting Tolerance", ["Inches", "Millimeters"], index=0)
if unit_tolerance == "Inches":
    glass_cutting_tolerance = st.number_input("Enter Glass Cutting Tolerance (in inches)", value=0.0625, format="%.4f")
else:
    val_mm = st.number_input("Enter Glass Cutting Tolerance (in mm)", value=0.0625 * inches_to_mm, format="%.3f")
    glass_cutting_tolerance = val_mm * mm_to_inches

unit_joint_top = st.radio("Select Unit for Joint Top", ["Inches", "Millimeters"], index=0)
if unit_joint_top == "Inches":
    joint_top = st.number_input("Enter Joint Top (in inches)", value=0.5, format="%.3f")
else:
    val_mm = st.number_input("Enter Joint Top (in mm)", value=0.5 * inches_to_mm, format="%.3f")
    joint_top = val_mm * mm_to_inches

unit_joint_bottom = st.radio("Select Unit for Joint Bottom", ["Inches", "Millimeters"], index=0)
if unit_joint_bottom == "Inches":
    joint_bottom = st.number_input("Enter Joint Bottom (in inches)", value=0.125, format="%.3f")
else:
    val_mm = st.number_input("Enter Joint Bottom (in mm)", value=0.125 * inches_to_mm, format="%.3f")
    joint_bottom = val_mm * mm_to_inches

unit_joint_left = st.radio("Select Unit for Joint Left", ["Inches", "Millimeters"], index=0)
if unit_joint_left == "Inches":
    joint_left = st.number_input("Enter Joint Left (in inches)", value=0.25, format="%.3f")
else:
    val_mm = st.number_input("Enter Joint Left (in mm)", value=0.25 * inches_to_mm, format="%.3f")
    joint_left = val_mm * mm_to_inches

unit_joint_right = st.radio("Select Unit for Joint Right", ["Inches", "Millimeters"], index=0)
if unit_joint_right == "Inches":
    joint_right = st.number_input("Enter Joint Right (in inches)", value=0.25, format="%.3f")
else:
    val_mm = st.number_input("Enter Joint Right (in mm)", value=0.25 * inches_to_mm, format="%.3f")
    joint_right = val_mm * mm_to_inches

# Part number
part_number = f"{system_type}-{profile_number}"

# File upload & template download
uploaded_file = st.file_uploader("Upload a CSV file", type="csv")
with open("SWR template.csv", "rb") as template_file:
    st.download_button("Download Template", template_file.read(), "SWR_template.csv", "text/csv")

if uploaded_file:
    df = pd.read_csv(uploaded_file)
    st.dataframe(df)

    # Convert dims
    df['Overall Width mm'] = df['Overall Width in'] * inches_to_mm
    df['Overall Height mm'] = df['Overall Height in'] * inches_to_mm
    j_l = joint_left * inches_to_mm
    j_r = joint_right * inches_to_mm
    j_t = joint_top * inches_to_mm
    j_b = joint_bottom * inches_to_mm

    df['SWR Width mm'] = df['Overall Width mm'] - j_l - j_r
    df['SWR Height mm'] = df['Overall Height mm'] - j_t - j_b
    df['SWR Width in'] = df['SWR Width mm'] * mm_to_inches
    df['SWR Height in'] = df['SWR Height mm'] * mm_to_inches

    df['Glass Width mm'] = df['SWR Width mm'] - (2 * glass_offset)
    df['Glass Height mm'] = df['SWR Height mm'] - (2 * glass_offset)
    df['Glass Width in'] = df['Glass Width mm'] * mm_to_inches
    df['Glass Height in'] = df['Glass Height mm'] * mm_to_inches

    # --- Glass File Export ---
    output_df = pd.DataFrame({'Item': range(1, len(df) + 1)})
    output_df['Glass Width in'] = df['Glass Width in']
    output_df['Glass Width (nearest 1/16)'] = output_df['Glass Width in'].apply(
        lambda x: f"{int(round(x * 16))//16} {int(round(x * 16))%16}/16" if round(x * 16)%16 else f"{int(round(x))}"
    )
    output_df['Glass Height in'] = df['Glass Height in']
    output_df['Glass Height (nearest 1/16)'] = output_df['Glass Height in'].apply(
        lambda x: f"{int(round(x * 16))//16} {int(round(x * 16))%16}/16" if round(x * 16)%16 else f"{int(round(x))}"
    )
    output_df['Area Each (ft²)'] = (output_df['Glass Width in'] * output_df['Glass Height in']) * sq_inches_to_sq_feet
    output_df['Qty'] = df['Qty']
    output_df['Area Total (ft²)'] = output_df['Qty'] * output_df['Area Each (ft²)']

    totals = pd.DataFrame(
        [['Totals', None, None, None, None, None,
          output_df['Qty'].sum(), output_df['Area Total (ft²)'].sum()]],
        columns=output_df.columns
    )
    output_df = pd.concat([output_df, totals], ignore_index=True)

    glass_file = BytesIO()
    with pd.ExcelWriter(glass_file, engine='xlsxwriter') as writer:
        ws = writer.book.add_worksheet("Sheet1")
        ws.insert_image('A1', 'ilogo.png', {'x_scale':0.2,'y_scale':0.2})
        ws.write('A7', "Project Name:");      ws.write('B7', project_name)
        ws.write('A8', "Project Number:");    ws.write('B8', project_number)
        ws.write('A9', "Date Created:");      ws.write('B9', datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        ws.write('A10',"Prepared By:");       ws.write('B10', prepared_by)
        output_df.to_excel(writer, sheet_name='Sheet1', startrow=12, index=False)

    st.download_button("Download Glass File", data=glass_file.getvalue(),
                       file_name=f"INO_{project_number}_SWR_Glass.xlsx")

    # --- AggCutOnly File Export ---
    df['Qty x 2'] = df['Qty'] * 2
    width_counts = df.groupby('SWR Width in')['Qty'].sum().sort_values(ascending=False)
    height_counts = df.groupby('SWR Height in')['Qty'].sum().sort_values(ascending=False)
    unique_dims = pd.Index(width_counts.index.tolist() + height_counts.index.tolist()).unique()

    agg_df = pd.DataFrame(0, index=unique_dims,
                          columns=['Part #','Miter'] + df['Tag'].unique().tolist() + ['Total QTY'])
    agg_df['Part #'] = part_number
    agg_df['Miter'] = "**"

    for _, row in df.iterrows():
        w, h, tag, q2 = row['SWR Width in'], row['SWR Height in'], row['Tag'], row['Qty x 2']
        if tag in agg_df.columns:
            agg_df.at[w, tag] += q2
            agg_df.at[h, tag] += q2
    agg_df['Total QTY'] = agg_df[df['Tag'].unique()].sum(axis=1)
    agg_df.index.name = "Finished Length in"
    agg_df = agg_df.reset_index()

    agg_file = BytesIO()
    with pd.ExcelWriter(agg_file, engine='xlsxwriter') as writer:
        ws = writer.book.add_worksheet("Sheet1")
        ws.insert_image('A1','ilogo.png',{'x_scale':0.2,'y_scale':0.2})
        ws.write('A7',"Project Name:");    ws.write('B7', project_name)
        ws.write('A8',"Project Number:");  ws.write('B8', project_number)
        ws.write('A9',"Date Created:");    ws.write('B9', datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        ws.write('A10',"Prepared By:");   ws.write('B10', prepared_by)
        ws.write('A11',"Finish:");        ws.write('B11', finish)
        agg_df.to_excel(writer, sheet_name='Sheet1', startrow=12, index=False)

    st.download_button("Download AggCutOnly File", data=agg_file.getvalue(),
                       file_name=f"INO_{project_number}_SWR_AggCutOnly.xlsx")

    # --- TagDetails File Export ---
    tag_file = BytesIO()
    with pd.ExcelWriter(tag_file, engine='xlsxwriter') as writer:
        for tag in df['Tag'].unique():
            tag_df = df[df['Tag'] == tag]
            rows = {'Item':[], 'Position':[], 'Quantity':[], 'Length (mm)':[], 'Length (in)':[]}
            for idx, r in tag_df.iterrows():
                for pos, length in [('left', r['SWR Width mm']),
                                    ('right',r['SWR Width mm']),
                                    ('top',  r['SWR Height mm']),
                                    ('bottom',r['SWR Height mm'])]:
                    rows['Item'].append(idx+1)
                    rows['Position'].append(pos)
                    rows['Quantity'].append(r['Qty']*2)
                    rows['Length (mm)'].append(length)
                    rows['Length (in)'].append(length * mm_to_inches)

            tag_output_df = pd.DataFrame(rows)
            ws = writer.book.add_worksheet(str(tag))
            ws.insert_image('A1','ilogo.png',{'x_scale':0.2,'y_scale':0.2})
            ws.write('A7',"Project Name:");    ws.write('B7', project_name)
            ws.write('A8',"Project Number:");  ws.write('B8', project_number)
            ws.write('A9',"Date Created:");    ws.write('B9', datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
            ws.write('A10',"Prepared By:");   ws.write('B10', prepared_by)
            tag_output_df.to_excel(writer, sheet_name=str(tag), startrow=12, index=False)

    st.download_button("Download TagDetails File", data=tag_file.getvalue(),
                       file_name=f"INO_{project_number}_SWR_TagDetails.xlsx")

    # --- SWR Table Export (with all inputs) ---
    swr_table_file = BytesIO()
    with pd.ExcelWriter(swr_table_file, engine='xlsxwriter') as writer:
        ws = writer.book.add_worksheet("Sheet1")
        ws.insert_image('A1','ilogo.png',{'x_scale':0.2,'y_scale':0.2})
        # metadata
        ws.write('A7',"Project Name:");                   ws.write('B7', project_name)
        ws.write('A8',"Project Number:");                 ws.write('B8', project_number)
        ws.write('A9',"Date Created:");                   ws.write('B9', datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        ws.write('A10',"Prepared By:");                   ws.write('B10', prepared_by)
        ws.write('A11',"Finish:");                        ws.write('B11', finish)
        ws.write('A12',"Glass Cutting Tolerance (in):");  ws.write('B12', glass_cutting_tolerance)
        ws.write('A13',"Joint Top (in):");                ws.write('B13', joint_top)
        ws.write('A14',"Joint Bottom (in):");             ws.write('B14', joint_bottom)
        ws.write('A15',"Joint Left (in):");               ws.write('B15', joint_left)
        ws.write('A16',"Joint Right (in):");              ws.write('B16', joint_right)
        # table itself
        df.drop(columns=["Qty x 2"], errors="ignore")\
          .to_excel(writer, sheet_name='Sheet1', startrow=17, index=False)

    st.download_button("Download SWR Table File", data=swr_table_file.getvalue(),
                       file_name=f"INO_{project_number}_SWR_Table.xlsx")
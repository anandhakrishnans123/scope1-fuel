import streamlit as st
import pandas as pd
import random
from io import BytesIO

# Function to merge sheets from an Excel file
def merge_sheets(file_path, sheets_to_merge):
    merged_data = pd.DataFrame()
    for sheet_name in sheets_to_merge:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        merged_data = pd.concat([merged_data, df], ignore_index=True)
    return merged_data

# Function to process fuel data for the FZE entity
def process_fuel_data(client_data, template_workbook_path, column_mapping, output_path, sheet_name):
    template_df = pd.read_excel(template_workbook_path, sheet_name=None)
    template_data = template_df[sheet_name]

    preserved_header = template_data.iloc[:0, :]

    matched_data = pd.DataFrame(columns=template_data.columns)

    for client_col, template_col in column_mapping.items():
        if client_col in client_data.columns and template_col in template_data.columns:
            matched_data[template_col] = client_data[client_col]

    if 'Start Date' in client_data.columns:
        matched_data['Res_Date'] = pd.to_datetime(matched_data['Res_Date']).dt.date

    # Fill missing values with random choices
    matched_data['Activity Unit'] = matched_data['Activity Unit'].apply(lambda x: random.choice(['Litres']) if pd.isna(x) else x)
    matched_data['Fuel Unit'] = matched_data['Fuel Unit'].apply(lambda x: random.choice(['Litres']) if pd.isna(x) else x)
    matched_data['CF Factor'] = matched_data['CF Factor'].apply(lambda x: random.choice(['IMO']) if pd.isna(x) else x)
    matched_data['GAS Type'] = matched_data['GAS Type'].apply(lambda x: random.choice(['CO2']) if pd.isna(x) else x)
    matched_data['Activity'] = matched_data['Activity'].apply(lambda x: random.choice([0.001]) if pd.isna(x) else x)
    matched_data['Fuel Type'] = matched_data['Fuel Type'].apply(lambda x: random.choice(['Diesel']) if pd.isna(x) else x)
    matched_data['Source'] = matched_data['Facility']

    final_data = pd.concat([preserved_header, matched_data], ignore_index=True)
    final_data.dropna(subset=['Res_Date'], inplace=True)

    final_data.to_excel(output_path, index=False)

# Function to process SSL data
def process_ssl_data(client_data, template_workbook_path, column_mapping, output_path, sheet_name):
    template_df = pd.read_excel(template_workbook_path, sheet_name=None)
    template_data = template_df[sheet_name]

    preserved_header = template_data.iloc[:0, :]

    matched_data = pd.DataFrame(columns=template_data.columns)

    for client_col, template_col in column_mapping.items():
        if client_col in client_data.columns and template_col in template_data.columns:
            matched_data[template_col] = client_data[client_col]

    if 'Start Date' in client_data.columns:
        matched_data['Res_Date'] = pd.to_datetime(matched_data['Res_Date']).dt.date

    matched_data['Activity Unit'] = matched_data['Activity Unit'].apply(lambda x: random.choice(['MT']) if pd.isna(x) else x)
    matched_data['Fuel Unit'] = matched_data['Fuel Unit'].apply(lambda x: random.choice(['MT']) if pd.isna(x) else x)
    matched_data['CF Factor'] = matched_data['CF Factor'].apply(lambda x: random.choice(['IMO']) if pd.isna(x) else x)
    matched_data['GAS Type'] = matched_data['GAS Type'].apply(lambda x: random.choice(['CO2']) if pd.isna(x) else x)

    final_data = pd.concat([preserved_header, matched_data], ignore_index=True)
    final_data.dropna(subset=['Res_Date'], inplace=True)
    final_data.dropna(subset=['Activity'], inplace=True)
    final_data = final_data[final_data['Activity'] != 0]

    # Ensure 'Fuel Type' is a string and strip leading/trailing spaces
    if 'Fuel Type' in final_data.columns:
        final_data['Fuel Type'] = final_data['Fuel Type'].astype(str).str.strip()

        # Replace specific values
        final_data['Fuel Type'].replace({
            'LFO Consumed (in MT)': 'LFO',
            'HFO Consumed (in MT)': 'HFO',
            'DGO Consumed (in MT)': 'DGO'
        }, inplace=True)
    final_data.insert(1, 'Department', "")
    
    final_data.insert(3, 'End Date', final_data['Res_Date'])
    final_data.insert(3, 'Start Date', final_data['Res_Date'])
    
    
    final_data = final_data.dropna(subset=["Fuel Consumption"])

    return final_data

# Streamlit app setup
st.title('Fuel Data Processing')

# Dropdown menu to select the entity
entity = st.selectbox('Select Entity', ['Select', 'FZE', 'SSL'])

# File uploader
uploaded_file = st.file_uploader("Upload the source file", type=["xlsx"])

if uploaded_file is not None and entity != 'Select':
    # Define parameters based on selected entity
    if entity == 'FZE':
        sheets = ['FORKLIFT-16934', 'FORKLIFT-16935']
        column_mapping = {
            'End Date': 'Res_Date',
            'Remark': 'Facility',
            'Fuel Consumed (Litres)': 'Fuel Consumption'
        }
        template_path = 'Fuel-Type-Sample_scope1.xlsx'
        output_path = "output_client.xlsx"
    elif entity == 'SSL':
        sheets = ['TBC BADRINATH', 'TBC KAILASH', 'SSL KRISHNA', 'SSL VISHAKAPATNAM',
                  'SSL MUMBAI', 'SSL BRAMHAPUTRA', 'SSL GANGA', 'SSL BHARAT', 'SSL SABRIMALAI',
                  'SSL GUJARAT', 'SSL DELHI', 'SSL GODAVARI', 'SSL THAMIRABARANI']
        column_mapping = {
            'Start Date': 'Res_Date',
            'Vessel Name': 'Facility',
            'Vessel Type': 'Source',
            'Distance travelled (In NM)': 'Activity',
            'Fuel Type': "Fuel Type",
            'Consumed (in MT)': 'Fuel Consumption'
        }
        template_path = 'Fuel-Type-Sample_scope1.xlsx'
        output_path = "output_client.xlsx"

    # Process the uploaded file
    client_data = merge_sheets(uploaded_file, sheets)
    
    if entity == 'FZE':
        process_fuel_data(client_data, template_path, column_mapping, output_path, 'Fuel Type')
    elif entity == 'SSL':
        ssl_data_melted = client_data.melt(id_vars=['Location/Unit/Factory ID', 'Start Date', 'End Date', 'Vessel Name',
                                                      'Vessel Category', 'Vessel Type', 'Distance travelled (In NM)'],
                                             value_vars=['DGO Consumed (in MT)', 'HFO Consumed (in MT)', 'LFO Consumed (in MT)'],
                                             var_name='Fuel Type',
                                             value_name='Consumed (in MT)')
        
        # Process SSL data
        final_data = process_ssl_data(ssl_data_melted, template_path, column_mapping, output_path, 'Fuel Type')
        
        # Save processed data to Excel
        final_data.to_excel(output_path, index=False)

    st.write(f'{entity} Data processed and saved to {output_path}.')

    # Provide download link
    def get_file_download_link(file_path):
        with open(file_path, "rb") as f:
            data = f.read()
        return data

    st.download_button('Download Processed Data', get_file_download_link(output_path), file_name='Processed_Data.xlsx')

import streamlit as st
import pandas as pd
import io
import zipfile
from datetime import datetime
from pytz import timezone
from PIL import Image


# set page configurations
st.set_page_config(
    page_title="Excel Cleaner",
    page_icon="🤖",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# the custom CSS lives here:
hide_default_format = """
        <style>
            .reportview-container .main footer {visibility: hidden;}    
            #MainMenu, footer {visibility: hidden;}
            section.main > div:has(~ footer ) {
                padding-bottom: 1px;
                padding-left: 20px;
                padding-right: 40px;
                padding-top: 25px;
            }
            [data-testid="collapsedControl"] {
                display: none;
                } 
            [data-testid="stDecoration"] {
                background-image: linear-gradient(90deg, rgb(0,92,185), rgb(0,92,185));
                height: 35%;
                }
            div.stActionButton{visibility: hidden;}
        </style>
       """

# inject the CSS
st.markdown(hide_default_format, unsafe_allow_html=True)

# main title
st.markdown(
    "<p style='color:#000000; font-weight: 900; font-size: 46px'>Excel Cleaning Robot (version 2)</p>", unsafe_allow_html=True)

# sub title
st.markdown(
    "<p style='color:#000000; font-weight: 400; font-size: 20px'><em>This friendly user interface provides an automated workflow for cleaning ABI invoices. Put this robot to work!</em></p>", unsafe_allow_html=True)

st.write("")
st.write("")


# Function to clean each dataframe that's uploaded
def clean_dataframe(df):
    """Cleans a DataFrame by capturing values from specific cells and creating a new DataFrame.

    Args:
        df (pd.DataFrame): The input DataFrame.

    Returns:
        pd.DataFrame: The cleaned DataFrame.
    """

    # Extract values from specified cells
    invoice_date = df.iloc[0, 1].strftime("%m/%d/%Y")
    vendor_invoice_number = df.iloc[1, 1]
    abi_contract_name = df.iloc[2, 1]
    abi_contract_number = df.iloc[3, 1]
    project_WO_name = df.iloc[4, 1]
    abi_project_number = df.iloc[5, 1]

    # special case for this value, which may be "N/A" in the spreadsheet
    abi_WO_number = df.iloc[6, 1]
    abi_WO_number = abi_WO_number if not pd.isna(abi_WO_number) else "N/A"

    # rest of the values
    total_contract_WO_amount = '${:,.2f}'.format(df.iloc[7, 1])
    current_invoice_amount = '${:,.2f}'.format(df.iloc[8, 1])
    abi_cost_code = df.iloc[9, 1]
    prime_contract_vendor = df.iloc[10, 1]

    # Extract non-blank vendor/subcontractor values from Column 0 (first column)
    vendor_subcontractor_names = df.iloc[13:32, 0].dropna().tolist()

    # Get the indices of rows that have non-blank values in the 'vendors' list
    vendor_subcontractor_indices = df.iloc[13:32, 0].dropna().index

    # Extract corresponding values from the other columns using the same indices
    certification = df.iloc[vendor_subcontractor_indices, 1].tolist()
    race_ethnicity = df.iloc[vendor_subcontractor_indices, 2].tolist()
    additional_DBE_types = df.iloc[vendor_subcontractor_indices, 3].tolist()
    net_invoice_amount = ['${:,.2f}'.format(
        item) for item in df.iloc[vendor_subcontractor_indices, 4].tolist()]
    net_contracted_amount = ['${:,.2f}'.format(
        item) for item in df.iloc[vendor_subcontractor_indices, 5].tolist()]
    total_invoiced_toDate = ['${:,.2f}'.format(
        item) for item in df.iloc[vendor_subcontractor_indices, 6].tolist()]
    newly_added = df.iloc[vendor_subcontractor_indices, 7].tolist()

    # the total will be at the bottom of the table
    invoice_total = '${:,.2f}'.format(df.iloc[33, 4])

    # Determine Prime/Sub based on matching values in "Prime Contract/Vendor" and "Vendor/Subcontractor"
    prime_sub_column = ["Prime" if vendor ==
                        prime_contract_vendor else "Sub" for vendor in vendor_subcontractor_names]

    # Create a new DataFrame with the extracted values
    cleaned_df = pd.DataFrame({
        "Invoice Date": [invoice_date] * len(vendor_subcontractor_names),
        "Vendor Invoice #": [vendor_invoice_number] * len(vendor_subcontractor_names),
        "ABI Contract Name": [abi_contract_name] * len(vendor_subcontractor_names),
        "ABI Contract #": [abi_contract_number] * len(vendor_subcontractor_names),
        "Project or Work Order Name": [project_WO_name] * len(vendor_subcontractor_names),
        "ABI Project #": [abi_project_number] * len(vendor_subcontractor_names),
        "ABI Work Order #": [abi_WO_number] * len(vendor_subcontractor_names),
        "Total Contract/Work Order Amt": [total_contract_WO_amount] * len(vendor_subcontractor_names),
        "Current Invoice Amount": [current_invoice_amount] * len(vendor_subcontractor_names),
        "ABI Cost Code #": [abi_cost_code] * len(vendor_subcontractor_names),
        "Prime Contractor/Vendor": [prime_contract_vendor] * len(vendor_subcontractor_names),
        "Vendor/Subcontractor": vendor_subcontractor_names,
        "Prime/Sub": prime_sub_column,
        "Certification": certification,
        "Race/Ethnicity": race_ethnicity,
        "Additional DBE Types": additional_DBE_types,
        "Net Invoice Amount ($)": net_invoice_amount,
        "Net Contracted Amount ($)": net_contracted_amount,
        "Total Invoiced to Date ($)": total_invoiced_toDate,
        "Newly added?": newly_added,
        "Invoice Total": [invoice_total] * len(vendor_subcontractor_names)
    })

    # fill in missing values
    cleaned_df["Certification"] = cleaned_df["Certification"].fillna(
        "N/A").replace("Select", "N/A")
    cleaned_df["Race/Ethnicity"] = cleaned_df["Race/Ethnicity"].fillna(
        "N/A").replace("Select", "N/A")
    cleaned_df["Additional DBE Types"] = cleaned_df["Additional DBE Types"].fillna(
        "N/A").replace("Select", "N/A")
    cleaned_df["Newly added?"] = cleaned_df["Newly added?"].fillna(
        "N/A").replace("Select", "N/A")

    return cleaned_df


# Function to handle file uploading and cleaning
def handle_upload():

    uploaded_files = st.file_uploader(
        label="Upload Excel file(s) to be processed:",
        accept_multiple_files=True
    )

    cleaned_dataframes = {}

    if uploaded_files:
        for file in uploaded_files:
            try:
                # Read Excel file into a Pandas DataFrame
                df = pd.read_excel(file)

                # Clean the DataFrame (modify this based on your cleaning requirements)
                cleaned_df = clean_dataframe(df)

                # Get the filename without extension
                filename = file.name.split('.')[0]

                # add filename as another column
                cleaned_df['Original file name'] = filename

                # Save cleaned dataframe with the original filename
                cleaned_dataframes[filename] = cleaned_df

            except Exception as e:
                st.markdown(
                    f"<p style='color:#8B0000; font-weight: 200; font-size: 14px'><b>Beep-bop, Layla!</b> I found an error with the following file: <u>{file.name}</u>. If you're curious, the error is: {e}.</p>",
                    unsafe_allow_html=True
                )
                continue

        # Create a ZipFile object to store individual Excel files
        tz = timezone("America/Atlanta")
        timestamp = datetime.now(tz).strftime("%m-%d-%Y_%I.%M%p")
        zip_file_name = f"cleaned_files_{timestamp}.zip"

        buffer_zip = io.BytesIO()

        # instantiate a variable to keep track of the number of files to be included in the Zip file
        number_of_files = 0

        # Create a ZipFile object to store individual Excel files
        with zipfile.ZipFile(buffer_zip, 'w') as zip_file:
            # Save each dataframe as a separate Excel file in the zip archive
            for filename, cleaned_df in cleaned_dataframes.items():

                excel_file = io.BytesIO()
                with pd.ExcelWriter(excel_file, engine='xlsxwriter') as writer:
                    cleaned_df.to_excel(
                        writer,
                        sheet_name='Table 1',
                        index=False
                    )

                    # Get the xlsxwriter workbook and worksheet objects
                    workbook = writer.book
                    worksheet = writer.sheets["Table 1"]

                    # autofit columns
                    worksheet.autofit()

                    # Close the Pandas Excel writer and save the Excel file
                    writer.close()

                excel_file.seek(0)
                zip_file.writestr(
                    f'{filename}_clean.xlsx', excel_file.read())

                number_of_files += 1  # increment the number of files

        # Close the ZipFile and output the zip file to the buffer
        buffer_zip.seek(0)

        # show user how many files were uploaded
        st.markdown(
            f"<p style='color:#000000; font-weight: 600; font-size: 18px'><em>Total files processed: {number_of_files}</em></p>", unsafe_allow_html=True)

        st.download_button(
            label="Download as zip folder",
            data=buffer_zip,
            file_name=zip_file_name,
            mime="application/zip"
        )


handle_upload()

# draw logo at lower-right corner of dashboard
st.write("")
st.write("")
st.write("")
st.write("")
col1, col2 = st.columns([3, 1])
im = Image.open('Content/abi_2.png')
col2.image(im, width=150)

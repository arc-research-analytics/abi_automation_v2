import streamlit as st
import pandas as pd
import math
import io
import zipfile
from datetime import datetime
import pytz
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
    company_name = df.iloc[0, 1]
    project_name = df.iloc[1, 1]
    project_number = df.iloc[2, 1]
    total_bid_amount = df.iloc[3, 1]
    company_contact = df.iloc[21, 1]
    contact_phone = df.iloc[22, 1]
    contact_email = df.iloc[23, 1]
    date = df.iloc[24, 1]

    # Extract non-blank vendor/subcontractor values from Column 0 (first column)
    contractor_names = df.iloc[6:19, 0].dropna().tolist()

    # Get the indices of rows that have non-blank values in the 'vendors' list
    contractor_indices = df.iloc[6:19, 0].dropna().index

    # Extract corresponding values from the other columns using the same indices
    certifying_agency = df.iloc[contractor_indices, 1].tolist()
    certification = df.iloc[contractor_indices, 2].tolist()
    race_ethnicity = df.iloc[contractor_indices, 3].tolist()
    additional_DBE_groups = df.iloc[contractor_indices, 4].tolist()
    description_of_work = df.iloc[contractor_indices, 5].tolist()
    dollar_value = df.iloc[contractor_indices, 6].tolist()

    # st.write(f'dollar values: {dollar_value}')

    # Determine Prime/Sub based on matching values in "Prime Contract/Vendor" and "Vendor/Subcontractor"
    prime_sub_column = ["Prime" if vendor ==
                        company_name else "Sub" for vendor in contractor_names]

    # Create a new DataFrame with the extracted values
    cleaned_df = pd.DataFrame({
        "Company Name": [company_name] * len(contractor_names),
        "Project Name": [project_name] * len(contractor_names),
        "Project # (if applicable)": [project_number] * len(contractor_names),
        "Total Bid Amount ($)": [total_bid_amount] * len(contractor_names),
        "Primary & Subcontractor Names": contractor_names,
        "Prime/Sub": prime_sub_column,
        "Certifying Agency": certifying_agency,
        "Certification": certification,
        "Race/Ethnicity": race_ethnicity,
        "Additional DBE Groups": additional_DBE_groups,
        "Description of Work": description_of_work,
        "$ Value": dollar_value,
        "Company Contact": [company_contact] * len(contractor_indices),
        "Contact Phone": [contact_phone] * len(contractor_indices),
        "Contact Email": [contact_email] * len(contractor_indices),
        "Date": [date] * len(contractor_indices),
    })

    # fill in missing values
    cleaned_df["Race/Ethnicity"] = cleaned_df["Race/Ethnicity"].fillna(
        "N/A").replace("Select", "N/A")
    cleaned_df["Additional DBE Groups"] = cleaned_df["Additional DBE Groups"].fillna(
        "N/A").replace("Select", "N/A")
    cleaned_df["Description of Work"] = cleaned_df["Description of Work"].fillna(
        "N/A").replace("Select", "N/A")

    # Display extracted values
    # st.write("Extracted values:")
    # st.dataframe(cleaned_df)

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

                # Save cleaned dataframe with the original filename
                cleaned_dataframes[filename] = cleaned_df

            except Exception as e:
                st.markdown(
                    f"<p style='color:#8B0000; font-weight: 200; font-size: 14px'><b>Beep-bop, Layla!</b> I found an error with the following file: <u>{file.name}</u>. If you're curious, the error is: {e}.</p>",
                    unsafe_allow_html=True
                )
                continue

        # Create a ZipFile object to store individual Excel files
        est_timezone = pytz.timezone("US/Eastern")
        timestamp = datetime.now(est_timezone).strftime("%m-%d-%Y_%I.%M%p")
        zip_file_name = f"cleaned_files_{timestamp}.zip"

        buffer_zip = io.BytesIO()

        # instantiate a variable to keep track of the number of files to be included in the Zip file
        number_of_files = 0

        # Create a ZipFile object to store individual Excel files
        with zipfile.ZipFile(buffer_zip, 'w') as zip_file:
            # Save each dataframe as a separate Excel file in the zip archive
            for filename, cleaned_df in cleaned_dataframes.items():
                # Ensure "Invoice Date" is in date format without time
                cleaned_df['Date'] = pd.to_datetime(
                    cleaned_df['Date']).dt.date

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

                    # Define a date format in the "MM/DD/YYYY" format
                    date_format = workbook.add_format(
                        {'num_format': 'mm/dd/yyyy'})

                    # Define a currency format with dollar signs and thousands separators
                    dollar_format = workbook.add_format(
                        {'num_format': '$#,##0.00'})

                    # Set the number format for the specified columns
                    worksheet.set_column("P:P", None, date_format)
                    worksheet.set_column("D:D", None, dollar_format)
                    worksheet.set_column("L:L", None, dollar_format)

                    # autofit columns
                    worksheet.autofit()

                    # Close the Pandas Excel writer and save the Excel file
                    writer.close()

                excel_file.seek(0)
                zip_file.writestr(
                    f'{filename}.xlsx', excel_file.read())

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

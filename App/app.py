# -----------------------------------------------------------
# Imports
# -----------------------------------------------------------
import streamlit as st
import pandas as pd
import datetime as dt
import os
from dotenv import load_dotenv
from openpyxl import load_workbook
import io
import msoffcrypto

# -----------------------------------------------------------
# Variables
# -----------------------------------------------------------
# Key to join the dataframes
join_key = 'Bedrijfsnaam'

# Password to secure the app
load_dotenv()
password = os.getenv('password')

# Insert logo
tdf_logo = 'https://static.wixstatic.com/media/35c2c9_f5b19d5824844a429b9168156112a314~mv2.png/v1/fill/w_81,h_75,al_c,q_85,usm_0.66_1.00_0.01,enc_auto/TDF%20logo%20wit%20.png'

# -----------------------------------------------------------
# Data
# -----------------------------------------------------------
# Load the file
file = "20240208_tdf_dmu.xlsx"

# Open the file
with open(file, "rb") as f:
    decryptor = msoffcrypto.OfficeFile(f)
    decryptor.load_key(password=password)  # Use password to decrypt file

    decrypted = io.BytesIO()
    decryptor.decrypt(decrypted)

    # Load workbook with openpyxl
    decrypted.seek(0)
    workbook = load_workbook(decrypted, read_only=True)


# Load your data into a pandas DataFrame
df_employees = pd.DataFrame(workbook['Sheet1'].values)
df_companies = pd.DataFrame(workbook['Sheet2'].values)

# Promote first row as headers
df_employees.columns = df_employees.iloc[0]
df_employees = df_employees[1:]
df_employees.reset_index(drop=True, inplace=True)

df_companies.columns = df_companies.iloc[0]
df_companies = df_companies[1:]
df_companies.reset_index(drop=True, inplace=True)


# -----------------------------------------------------------
# The streamlit app
# -----------------------------------------------------------
# Set the page to use a wide layout
st.set_page_config(
    page_title="TDF | DMU",
    layout="wide")

# Add a password input field to the sidebar
password_input = st.sidebar.text_input("Voer wachtwoord in:", type="password")

if password_input == password:
    # -----------------------------------------------------------
    # Header
    # -----------------------------------------------------------

    st.image(tdf_logo, width=100)
    st.title('The Digital Federation')

    st.markdown("")
    st.header("De prospects:")

    # -----------------------------------------------------------
    # Nested table with companies and employees
    # -----------------------------------------------------------
    for _, company in df_companies.iterrows():
        with st.expander(f"{company[join_key]} "):
            filtered_employees = df_employees[df_employees[join_key]
                                              == company[join_key]]

            # Select the columns to display
            columns_to_display = ['URL afbeelding', 'Volledige naam', 'Functie categorie', 'E-mail 1',
                                  'E-mail 2', 'Telefoonnummer 1', 'Telefoonnummer 2', 'LinkedIn URL']
            filtered_employees_display = filtered_employees[columns_to_display]

            st.data_editor(
                filtered_employees_display,
                column_config={
                    "LinkedIn URL": st.column_config.LinkColumn(
                        "LinkedIn URL",
                        display_text="Open profiel"
                    ),
                    "URL afbeelding": st.column_config.ImageColumn(
                        "Image", help="Afbeelding van desbetreffende persoon"
                    )
                },
                hide_index=True,
            )
    # -----------------------------------------------------------
    # Download button to get the whole dataset
    # -----------------------------------------------------------
    st.markdown("")
    st.markdown("")
    st.download_button(
        label="Download de gehele dataset",
        data=df_employees.to_csv(index=False),
        file_name=f'{dt.datetime.now().strftime("%Y%m%d")}_tdf_dmu.csv',
        mime='text/csv',
    )


else:
    # If the password is incorrect, display an error message to
    st.error("Incorrect password. Please try again in the sidebar.")

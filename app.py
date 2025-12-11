import streamlit as st
import pandas as pd
import requests
import time
import io


# --- Helper Function: Query CMS API ---
def get_taxonomy_data(npi_number):
    """
    Queries the NPI Registry API version 2.1 for a given NPI number.
    Returns a list of dictionaries containing taxonomy details or None if failed.
    """
    url = "https://npiregistry.cms.hhs.gov/api/"
    params = {
        'number': npi_number,
        'version': '2.1'
    }

    try:
        response = requests.get(url, params=params, timeout=5)
        if response.status_code == 200:
            data = response.json()
            if data.get('result_count', 0) > 0 and 'results' in data:
                return data['results'][0].get('taxonomies', [])
    except Exception as e:
        pass

    return []


# --- Helper Function: Convert DF to Excel bytes ---
def convert_df_to_excel(df):
    """
    Converts a pandas DataFrame to an Excel file in memory.
    """
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Taxonomy Results')
    processed_data = output.getvalue()
    return processed_data


# --- Main Streamlit App ---
def main():
    st.set_page_config(page_title="NPI Taxonomy Lookup", layout="centered")

    # --- CUSTOM CSS FOR FILE UPLOADER ---
    # This block targets the Streamlit file uploader to make it bigger and change color on hover
    st.markdown("""
        <style>
        /* Target the drop zone container */
        [data-testid='stFileUploader'] section {
            padding: 3rem; /* Makes the drop zone much taller/bigger */
            background-color: #f0f2f6; /* Light gray background by default */
            border: 2px dashed #ccc;
            border-radius: 10px;
            text-align: center;
        }

        /* Change color when hovering (simulates drag-over effect) */
        [data-testid='stFileUploader'] section:hover {
            background-color: #e3f2fd; /* Light blue background */
            border-color: #2196f3;     /* Blue border */
        }

        /* Optional: Increase the font size of the text inside */
        [data-testid='stFileUploader'] section > div {
             font-size: 1.2rem;
        }
        </style>
    """, unsafe_allow_html=True)
    # -------------------------------------

    st.title("🏥 NPI Taxonomy Lookup Tool")
    st.markdown("""
    Upload your Excel file to find taxonomy codes for NPI numbers.

    **Instructions:**
    1. Ensure your NPIs are in **Column A** of the last tab (or a tab named "Missing NPIs (kelvin)").
    2. Upload the file below.
    3. Download the processed file with Taxonomy codes.
    """)

    uploaded_file = st.file_uploader("Drag and drop Excel file here", type=['xlsx', 'xls'])

    if uploaded_file is not None:
        try:
            # Try to load the specific sheet first, fallback to the last sheet
            try:
                df = pd.read_excel(uploaded_file, sheet_name='Missing NPIs (kelvin)', dtype=str)
                st.info("Loaded sheet: 'Missing NPIs (kelvin)'")
            except ValueError:
                df = pd.read_excel(uploaded_file, sheet_name=-1, dtype=str)
                st.info("Target sheet not found. Loaded the last sheet in the file.")

            # Get NPIs from Column A (index 0)
            # Remove NaNs and duplicates
            npi_list = df.iloc[:, 0].dropna().unique()

            st.write(f"**Found {len(npi_list)} unique NPIs to process.**")

            if st.button("Start Processing"):
                progress_bar = st.progress(0)
                status_text = st.empty()

                output_rows = []
                total_npis = len(npi_list)

                for i, npi in enumerate(npi_list):
                    # Clean NPI string
                    npi = str(npi).strip().replace('.0', '')  # Handle if excel read as float

                    # Update status
                    status_text.text(f"Processing {i + 1}/{total_npis}: {npi}")
                    progress_bar.progress((i + 1) / total_npis)

                    if len(npi) != 10 or not npi.isdigit():
                        output_rows.append({'NPI': npi, 'Taxonomy Code': 'Invalid Format'})
                        continue

                    # Fetch data
                    taxonomies = get_taxonomy_data(npi)

                    if taxonomies:
                        for tax in taxonomies:
                            output_rows.append({
                                'NPI': npi,
                                'Taxonomy Code': tax.get('code', ''),
                                'Taxonomy Description': tax.get('desc', ''),
                                'Primary Taxonomy': tax.get('primary', ''),
                                'State': tax.get('state', ''),
                                'License': tax.get('license', '')
                            })
                    else:
                        output_rows.append({
                            'NPI': npi,
                            'Taxonomy Code': 'Not Found'
                        })

                    # Rate limiting (10 requests per second is the limit usually, stay safe with small sleep)
                    time.sleep(0.05)

                # Finalize
                result_df = pd.DataFrame(output_rows)

                st.success("Processing Complete!")
                st.dataframe(result_df.head())

                # Create download button
                excel_data = convert_df_to_excel(result_df)

                st.download_button(
                    label="📥 Download Results as Excel",
                    data=excel_data,
                    file_name="npi_taxonomy_results.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        except Exception as e:
            st.error(f"Error reading file: {e}")


if __name__ == "__main__":
    main()
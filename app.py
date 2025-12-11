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
    st.markdown("""
        <style>
        [data-testid='stFileUploader'] section {
            padding: 80px 20px;
            background-color: #f8f9fa;
            border: 3px dashed #4CAF50;
            border-radius: 20px;
            text-align: center;
            transition: all 0.2s ease-in-out;
        }

        [data-testid='stFileUploader'] section:hover,
        [data-testid='stFileUploader'] section:focus-within {
            background-color: #e8f5e9;
            border-color: #2E7D32;
            transform: scale(1.01);
        }

        [data-testid='stFileUploader'] section > div {
             font-size: 1.5rem;
             font-weight: bold;
             color: #333;
        }
        </style>
    """, unsafe_allow_html=True)

    st.title("🏥 NPI Taxonomy Lookup Tool")
    st.markdown("""
    **Instructions:**
    1. Drag your file into the green box below.
    2. Select the **Sheet Name** containing your data.
    3. Ensure there is a column header named **"NPI"**.
    """)

    uploaded_file = st.file_uploader("Drop Excel File Here", type=['xlsx', 'xls'])

    if uploaded_file is not None:
        try:
            # 1. Load the Excel File wrapper (does not read data yet, just metadata)
            xls = pd.ExcelFile(uploaded_file)

            # 2. Let user pick the sheet
            sheet_names = xls.sheet_names
            selected_sheet = st.selectbox("Select the sheet containing NPIs:", sheet_names)

            if selected_sheet:
                # Load the specific sheet
                df = pd.read_excel(uploaded_file, sheet_name=selected_sheet, dtype=str)

                # Clean column names (remove spaces around headers)
                df.columns = df.columns.str.strip()

                # 3. Look for "NPI" column header
                if 'NPI' in df.columns:
                    # Extract NPIs from that specific column
                    npi_list = df['NPI'].dropna().unique()
                    st.success(f"Found column 'NPI' with {len(npi_list)} unique records.")

                    if st.button("Start Processing"):
                        progress_bar = st.progress(0)
                        status_text = st.empty()

                        output_rows = []
                        total_npis = len(npi_list)

                        for i, npi in enumerate(npi_list):
                            # Clean NPI string
                            npi = str(npi).strip().replace('.0', '')

                            status_text.text(f"Processing {i + 1}/{total_npis}: {npi}")
                            progress_bar.progress((i + 1) / total_npis)

                            if len(npi) != 10 or not npi.isdigit():
                                output_rows.append({'NPI': npi, 'Taxonomy Code': 'Invalid Format'})
                                continue

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

                            time.sleep(0.05)

                        result_df = pd.DataFrame(output_rows)

                        st.success("Processing Complete!")
                        st.dataframe(result_df.head())

                        excel_data = convert_df_to_excel(result_df)

                        st.download_button(
                            label="📥 Download Results",
                            data=excel_data,
                            file_name="npi_taxonomy_results.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                else:
                    st.error("❌ Could not find a column named 'NPI' in the selected sheet. Please check your headers.")

        except Exception as e:
            st.error(f"Error reading file: {e}")


if __name__ == "__main__":
    main()
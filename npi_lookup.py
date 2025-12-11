import pandas as pd
import requests
import time


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
        response = requests.get(url, params=params, timeout=10)
        if response.status_code == 200:
            data = response.json()
            if data.get('result_count', 0) > 0 and 'results' in data:
                return data['results'][0].get('taxonomies', [])
    except Exception as e:
        print(f"Error fetching NPI {npi_number}: {e}")

    return []


def process_npi_file(input_file_path, output_file_path):
    print("Reading file...")
    try:
        # Load the specific sheet "Missing NPIs (kelvin)"
        # header=None assumes the first row might be data or a header.
        # If row 1 is a header, change to header=0 and adjust row access.
        df = pd.read_excel(input_file_path, sheet_name='Missing NPIs (kelvin)', dtype=str)
    except ValueError:
        print("Sheet 'Missing NPIs (kelvin)' not found. Trying the last sheet instead.")
        df = pd.read_excel(input_file_path, sheet_name=-1, dtype=str)

    # Assumes NPIs are in Column A (index 0)
    # We drop NaN values to avoid querying empty rows
    npi_column_data = df.iloc[:, 0].dropna().unique()

    output_rows = []

    print(f"Found {len(npi_column_data)} unique NPIs. Starting lookup...")

    for npi in npi_column_data:
        npi = npi.strip()
        if not npi.isdigit() or len(npi) != 10:
            print(f"Skipping invalid NPI format: {npi}")
            continue

        print(f"Looking up NPI: {npi}")
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
            # Keep the NPI in the list even if no data found
            output_rows.append({
                'NPI': npi,
                'Taxonomy Code': 'Not Found',
                'Taxonomy Description': '',
                'Primary Taxonomy': '',
                'State': '',
                'License': ''
            })

        # Be nice to the API
        time.sleep(0.1)

    # Create new DataFrame and save
    print("Saving results...")
    result_df = pd.DataFrame(output_rows)
    result_df.to_excel(output_file_path, index=False)
    print(f"Done! Saved to {output_file_path}")


# --- EXECUTION ---
if __name__ == "__main__":
    # Replace with your actual file name
    input_file = "input_data.xlsx"
    output_file = "npi_taxonomy_results.xlsx"

    # Check if file exists before running
    import os

    if os.path.exists(input_file):
        process_npi_file(input_file, output_file)
    else:
        print(f"Please create a file named '{input_file}' or update the script with your filename.")
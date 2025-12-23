import pandas as pd
import os
from pathlib import Path

def extract_id_number(assigned_id):
    """Extract the last section (number) from Assigned ID"""
    if pd.isna(assigned_id):
        return ""
    # Split by '-' and get the last part
    parts = str(assigned_id).split('-')
    return parts[-1] if parts else ""

def clean_whatsapp_number(phone):
    """Remove +91 prefix if present"""
    if pd.isna(phone):
        return ""
    phone_str = str(phone).strip()
    # Remove +91 prefix if present
    if phone_str.startswith('+91'):
        return phone_str[3:]
    return phone_str

def process_excel_to_csv(excel_file_path, output_folder='output_csvs'):
    """
    Process Excel file and generate separate CSV files for each Assigned Naqeeb
    
    Parameters:
    - excel_file_path: Path to the input Excel file
    - output_folder: Folder where CSV files will be saved (default: 'output_csvs')
    """
    
    # Create output folder if it doesn't exist
    Path(output_folder).mkdir(parents=True, exist_ok=True)
    
    # Read the Excel file
    print(f"Reading Excel file: {excel_file_path}")
    df = pd.read_excel(excel_file_path)
    
    # Map column names (assuming 0-indexed: A=0, B=1, C=2, etc.)
    # B=1 (Assigned ID#), C=2 (Full Name), E=4 (WhatsApp#), AO=40 (Assigned Naqeeb)
    
    # Get column names or use index if headers exist
    # Adjust these if your columns have different positions
    assigned_id_col = df.columns[1]  # Column B
    full_name_col = df.columns[2]     # Column C
    whatsapp_col = df.columns[4]      # Column E
    naqeeb_col = df.columns[40]       # Column AO
    
    print(f"Processing columns:")
    print(f"  - Assigned ID: {assigned_id_col}")
    print(f"  - Full Name: {full_name_col}")
    print(f"  - WhatsApp: {whatsapp_col}")
    print(f"  - Assigned Naqeeb: {naqeeb_col}")
    
    # Filter out rows where Assigned Naqeeb is blank
    df_filtered = df[df[naqeeb_col].notna() & (df[naqeeb_col] != '')].copy()
    
    print(f"\nTotal rows: {len(df)}")
    print(f"Rows with Assigned Naqeeb: {len(df_filtered)}")
    
    # Process data
    df_filtered['Name'] = df_filtered[assigned_id_col].astype(str) + '_' + df_filtered[full_name_col].astype(str)
    df_filtered['Mobile'] = df_filtered[whatsapp_col].apply(clean_whatsapp_number)
    
    # Group by Assigned Naqeeb and create separate CSV files
    grouped = df_filtered.groupby(naqeeb_col)
    
    print(f"\nGenerating CSV files for {len(grouped)} Naqeeb(s):")
    
    for naqeeb_name, group in grouped:
        # Clean the naqeeb name for use as filename
        clean_name = str(naqeeb_name).strip().replace('/', '_').replace('\\', '_').replace(' ', '_')
        csv_filename = f"{clean_name}.csv"
        csv_path = os.path.join(output_folder, csv_filename)
        
        # Select only the required columns
        output_df = group[['Name', 'Mobile']].copy()
        
        # Save to CSV
        output_df.to_csv(csv_path, index=False)
        print(f"  ✓ {csv_filename} - {len(output_df)} records")
    
    print(f"\nAll CSV files saved to: {output_folder}/")

# Main execution
if __name__ == "__main__":
    # Replace with your Excel file path
    excel_file = r"D:\Work\py-ws\whatsapp-registration-extractor\contacts-generator\Hyd Contacts.xlsx"  # Change this to your actual file path
    
    # Optional: specify custom output folder
    output_folder = r"D:\Work\py-ws\whatsapp-registration-extractor\contacts-generator\output_csvs"
    
    try:
        process_excel_to_csv(excel_file, output_folder)
        print("\n✓ Processing completed successfully!")
    except FileNotFoundError:
        print(f"\n✗ Error: File '{excel_file}' not found.")
        print("Please update the 'excel_file' variable with the correct path.")
    except Exception as e:
        print(f"\n✗ Error occurred: {str(e)}")
        import traceback
        traceback.print_exc()
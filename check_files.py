import pandas as pd

def check_files():
    # Read source file
    print("Reading source file (HealthOptimisation.xlsx)...")
    source_df = pd.read_excel('HealthOptimisation.xlsx')
    print("\nSource file columns:", source_df.columns.tolist())
    print("\nFirst few rows of source file:")
    print(source_df.head())
    
    # Read contact file
    print("\n\nReading contact file (output/HealthOptimization_contact_info_organized.xlsx)...")
    contact_df = pd.read_excel('output/HealthOptimization_contact_info_organized.xlsx')
    print("\nContact file columns:", contact_df.columns.tolist())
    print("\nFirst few rows of contact file:")
    print(contact_df.head())

if __name__ == "__main__":
    check_files() 
import os
import glob
import shutil
import pandas as pd

def filter_and_export_csv(directory, output_directory):
    # Ensure the output directory exists
    if not os.path.exists(output_directory):
        os.makedirs(output_directory)
        print(f"Created output directory: {output_directory}")
    
    # Find all CSV files in the directory
    csv_files = glob.glob(os.path.join(directory, '*.csv'))
    
    # Process each CSV file
    file_paths = []
    for file_path in csv_files:
        file_name = os.path.basename(file_path)
        base_name, _ = os.path.splitext(file_name)
        temp_directory = os.path.join(output_directory,"time_temp")
        if not os.path.exists(temp_directory):
            os.makedirs(temp_directory)
            print(f"Created temp directory: {temp_directory}")
        output_file = os.path.join(temp_directory, f"{base_name}.xlsx")
        
        print(f"Processing file: {file_path}")
        
        # Read the CSV file into a DataFrame
        df = pd.read_csv(file_path)
        #find out location
        if 'Location' in df.columns and df['Location'].str.contains('Carlsbad', case=False, na=False).any():
            loc = 'CB'
            #Add salary:
            salary_rows = pd.DataFrame([{'Employee': 'Jane Doe', 'Regular Pay': 500.00, 'Total Pay': 500.00},
                                    {'Employee': 'John Doe', 'Regular Pay': 100.00, 'Total Pay': 100.00}], columns=df.columns)
            loc = 'DM'
            salary_rows = pd.DataFrame([{'Employee': 'Jane Doe', 'Regular Pay': 800.00, 'Total Pay': 800.00},
                                    {'Employee': 'John Doe', 'Regular Pay': 400.00, 'Total Pay': 400.00}],columns=df.columns)
        elif 'Location' in df.columns and df['Location'].str.contains('Encinitas', case=False, na=False).any():
            loc = 'EN'
            salary_rows = pd.DataFrame([{'Employee': 'Jane Doe', 'Regular Pay': 300.00, 'Total Pay': 300.00},
                                    {'Employee': 'John Doe', 'Regular Pay': 100.00, 'Total Pay': 100.00}],columns=df.columns)
        else:
            loc = 'Unknown'
        output_file = os.path.join(temp_directory, f"{loc}_{base_name}.xlsx")
        
        # Filter out rows where 'Job Title' contains 'Generic' or 'Driver'
        filtered_df = df[~df['Job Title'].str.contains('Generic|Driver', case=False, na=False)]
        
        # Process 'Half Day Server' and 'Server' job titles
        half_day_server_jobs = filtered_df[filtered_df['Job Title'].isin(['Half Day Server', 'Half Day Busser'])] 
        for _, half_day_row in half_day_server_jobs.iterrows():
            server_rows = filtered_df[(filtered_df['Job Title'].isin(['Server', 'Busser','Cashier'])) & (filtered_df['Employee'] == half_day_row['Employee'])]
            if not server_rows.empty:
                idx = server_rows.index[0]
                for column in filtered_df.columns:
                    if column not in ['Employee', 'Job Title','Hourly Rate','Location']:
                        filtered_df.at[idx, column] += half_day_row[column]
                filtered_df = filtered_df[filtered_df.index != half_day_row.name]
        
        
        filtered_df = pd.concat([filtered_df, salary_rows], ignore_index=True)

        # Export the filtered DataFrame to an Excel file
        filtered_df.to_excel(output_file, index=False)
        print(f"Filtered data exported to {output_file}")
        
        file_paths.append(output_file)

    # Combine all filtered Excel files into one workbook
    output_file = os.path.basename(output_file)
    output_file_name,_=os.path.splitext(output_file)
    combined_file = os.path.join(output_directory, f"{output_file_name}.xlsx")
    # Check if the combined file already exists and delete it
    if os.path.exists(combined_file):
        os.remove(combined_file)
        print(f"Deleted existing file: {combined_file}")
    with pd.ExcelWriter(combined_file, engine='openpyxl') as writer:
        for file_path in file_paths:
            df = pd.read_excel(file_path)
            # Extract date part from the file name for the sheet name
            sheet_name = '_'.join(os.path.splitext(os.path.basename(file_path))[0].split('_')[2:])
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            print(f"Added {file_path} to {combined_file} as sheet {sheet_name}")

    print(f"All files combined into {combined_file}")
    shutil.rmtree(temp_directory)
    print("Deleting temporary files")

# Set the directory paths
current_directory = os.getcwd()
src_directory = os.path.join(current_directory, 'time_src')
output_directory = os.path.join(current_directory, 'output')

filter_and_export_csv(src_directory, output_directory)
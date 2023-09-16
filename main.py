import pandas as pd
import os

def process_excel_files_in_folder():
    try:
        script_directory = os.path.dirname(os.path.abspath(__file__))
        excel_files = [f for f in os.listdir(script_directory) if f.endswith(".xlsx")]

        if not excel_files:
            print("No Excel files found in the folder.")
            return

        for excel_file in excel_files:
            input_file_path = os.path.join(script_directory, excel_file)
            
            # Specify the data types for columns GS1CompanyPrefix and GTIN as 'str' (string)
            dtype = {'GS1CompanyPrefix': str, 'GTIN': str}
            df = pd.read_excel(input_file_path, dtype=dtype)
            
            df = pd.read_excel(input_file_path, dtype=dtype)

            gtin_targets = {}

            for index, row in df.iterrows():
                gtin = row['GTIN']
                target_markets = row['TargetMarkets']
                
                if gtin in gtin_targets:
                    gtin_targets[gtin].append(target_markets)
                else:
                    gtin_targets[gtin] = [target_markets]

            for gtin, targets in gtin_targets.items():
                concatenated_targets = '~'.join(map(str, targets))
                df.loc[df['GTIN'] == gtin, 'TargetMarkets'] = concatenated_targets

            df.drop_duplicates(subset='GTIN', keep='first', inplace=True)
            output_file_name = f"{os.path.splitext(excel_file)[0]}-output.xlsx"
            output_file_path = os.path.join(script_directory, output_file_name)
            
            # Save the DataFrame to Excel with the specified data types
            writer = pd.ExcelWriter(output_file_path, engine='openpyxl')
            df.to_excel(writer, index=False, sheet_name='Sheet1', float_format='0')  # Use custom format '0'
            writer.save()
            
            print(f"Processed '{excel_file}' and saved as '{output_file_path}'")

    except Exception as e:
        print(f"An error occurred: {str(e)}")

if __name__ == "__main__":
    process_excel_files_in_folder()

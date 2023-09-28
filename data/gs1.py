import pandas as pd

def process_excel_files_in_folder(file):
    dtype = {'GS1COMPANYPREFIX': str, 'GTIN': str}
    try:
        df = pd.read_excel(file.stream, engine='openpyxl', dtype=dtype)
            
        # Specify the data types for columns GS1CompanyPrefix and GTIN as 'str' (string)
        
        gtin_targets = {}

        for index, row in df.iterrows():
            gtin = row['GTIN']
            target_markets = row['TARGETMARKETS']
                
            if gtin in gtin_targets:
                gtin_targets[gtin].append(target_markets)
            else:
                gtin_targets[gtin] = [target_markets]

        for gtin, targets in gtin_targets.items():
            concatenated_targets = '~'.join(map(str, targets))
            df.loc[df['GTIN'] == gtin, 'TARGETMARKETS'] = concatenated_targets

        df.drop_duplicates(subset='GTIN', keep='first', inplace=True)
            
        # Save the DataFrame to Excel with the specified data types
        writer = pd.ExcelWriter(file, engine='openpyxl')
        df.to_excel(writer, index=False, sheet_name='Sheet1', float_format='0') 
        df.to_excel('/home/garciagi/GS1/GS1-report.xlsx', index=False, sheet_name='Sheet1', float_format='0') # Use custom format '0'
        writer.save()

    except Exception as e:
        print(f"An error occurred: {str(e)}")
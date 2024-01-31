import pandas as pd

def process_excel_files_in_folder(file):
    dtype = {'GS1COMPANYPREFIX': str, 'GTIN': str}
    try:
        df = pd.read_excel(file.stream, engine='openpyxl', dtype=dtype)
        #df = pd.read_excel(file, engine='openpyxl', dtype=dtype)

        # Data manipulation in 'DESC1LANGUAGE' column
        df['DESC1LANGUAGE'].replace({'chs': 'zh', 'cht': 'zh', 'zhh': 'zh', 'mx': 'es', 'rs': 'sr'}, inplace=True)

        # Validate and add leading zeros to 'GTIN' column
        df['GTIN'] = df['GTIN'].apply(lambda x: x.zfill(14) if pd.notna(x) else x)

        # Check for spaces and dashes in 'GTIN', and delete the rows
        df = df[~df['GTIN'].str.contains(' ') & ~df['GTIN'].str.contains('-')]

        # Create new column 'GS1COMPANYPREFIX' with the specified calculation
        df['GS1COMPANYPREFIX'] = df['GTIN'].apply(lambda x: x[1:8] if len(x) >= 8 else '')

        # Group data by 'GTIN' and concatenate 'TARGETMARKETS'
        df['TARGETMARKETS'] = df.groupby('GTIN')['TARGETMARKETS'].transform(lambda x: '~'.join(x))

        # Drop duplicates based on 'GTIN'
        df.drop_duplicates(subset='GTIN', keep='first', inplace=True)

        # Save the DataFrame to Excel with the specified data types
        writer = pd.ExcelWriter('/home/garciagi/GS1/GS1-report.xlsx', engine='openpyxl')
        #writer = pd.ExcelWriter('GS1-report.xlsx', engine='openpyxl')
        df.to_excel(writer, index=False, sheet_name='Sheet1', float_format='0')
        writer.save()

    except Exception as e:
        print(f"An error occurred: {str(e)}")

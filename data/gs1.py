import pandas as pd

def process_excel_files_in_folder(file):
    dtype = {'GS1COMPANYPREFIX': str, 'GTIN': str}
    try:
        df = pd.read_excel(file.stream, engine='openpyxl', dtype=dtype)
        #df = pd.read_excel(file, engine='openpyxl', dtype=dtype)

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


# This is a "fix-it" python file that just handles xls to xlsx files on the assumption they're not working.

import pandas as pd



def convert_xls_to_xlsx(self, input_path:str, output_path:str):
    # Convert xls to xlsx if needed

    # Read the .xls file using pandas
    df: DataFrame = pd.read_excel(input_path, engine='xlrd')

    # Write the DataFrame to a new .xlsx file
    df.to_excel(output_path, index=False)


if __name__ == "__main__":
    # Example usage
    input_path = 'Attendance_List.xls'
    output_path = 'Attendance_List.xlsx'
    convert_xls_to_xlsx(input_path, output_path)
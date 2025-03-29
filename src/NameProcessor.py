# Name Processor
# Work through the excel file and process the names.
# Update the file format if it is not in the correct format.
# Check for duplicates and remove them.
# Check for invalid names and flag them.

import openpyxl
import pandas as pd

class NameProcessor:
    """ Name Processor Class
        This class is responsible for processing names in an excel file.
        It will check for duplicates, invalid names, and save the processed names back to the file.
        It will also convert the file format if it is not in the correct format.
    """ 
    
    def __init__(self, file_path):
        self.file_path = file_path
        self.excel_vals: dict[str, str] = {}
        self.target_worksheet: str = 'Attendees' # Default Worksheet
        self.target_column: str = 'C' # Default column for names

    def load_names(self):
        # Load names from the excel file
        workbook = openpyxl.load_workbook(self.file_path)
        
        worksheet = workbook['Attendees']
        max_row = worksheet.max_row
        for row_num in range(2, max_row):
            cell = worksheet[f'{self.target_column}{row_num}']
            email = worksheet[f'F{row_num}']
            if cell.value is None:
                pass
            else:
                self.excel_vals[cell.value] = email.value


    def process_names(self):
        # Process the names and check for duplicates and invalid names
        pass

    def save_names(self):
        # Save the processed names back to the excel file
        pass

    def check_duplicates(self):
        # Check for duplicate names
        pass

    def check_invalid_names(self):
        # Check for invalid names
        pass
    
    def convert_xls_to_xlsx(self, input_path:str, output_path:str):
        # Convert xls to xlsx if needed

        # Read the .xls file using pandas
        df: DataFrame = pd.read_excel(input_path, engine='xlrd')

        # Write the DataFrame to a new .xlsx file
        df.to_excel(output_path, index=False)
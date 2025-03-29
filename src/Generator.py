from docx import Document as DC
import os
import win32com
import re

from util.template_path import TemplatePath
from typing import Optional
import datetime

class CertificateGenerator:
    """ Certificate Generator Class
        Create the Certificate PDF and in the correct folder.
    """

    def __init__(self, name, email, template_path: TemplatePath =None):
        self.name = name
        self.email = email
        self.templates: TemplatePath = template_path

    def generate_pdf(self, file_name):

        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = 0

        doc=word.Documents.open(file_name)
        new_file = "Tester"
        doc.SaveAs(new_file, FileFormat=17)
        doc.Close(SaveChanges=False)

        word.Quit()


    def CreateCerts(self):

        doc_path = self.templates.word_template # Set this to by dynamically assigned
        doc = DC(doc_path)

        try:
            os.chdir('Certificates')
        except:
            # Tooling has been run before, so we need to do something with the old certificates
            print("Certificates folder Doesn't exist. Need to create it.")
            os.mkdir('Certificates')
            os.chdir('Certificates')


        unique_path = self.name.replace(" ", "_")
        os.mkdir(unique_path)
        target_text = "Name Surname"
        for paragraph in doc.paragraphs:
            if target_text in paragraph.text:
                for run in paragraph.runs:
                    run.text = run.text.replace(target_text, self.name)

        modified_doc_path = unique_path + '.docx'
        save_location = unique_path + '/' + modified_doc_path
        print(save_location)
        doc.save(save_location)

        os.chdir(unique_path)

        email_contents = EmailTemplateGenerator(template_path=self.templates)
        email_contents.main(replacement_text=self.name)
        email_contents.store_email(input_string=self.email)


        print(modified_doc_path)
        self.generate_pdf(modified_doc_path)
        print(f"Document saved with the updated text")
        os.chdir('../../')
        
        
        
"""EmailGenerator
    Handling general email interactions for Certificate Generators

"""

import csv

class EmailTemplateGenerator():
    """
    Create the .txt file to be used by power automate
    to create the email
    """

    def __init__(self, template_path: str = None):
        self.template_path: Optional[TemplatePath] = template_path

    def main(self, replacement_text):
        """_summary_

        Args:
            replacement_text (_type_): _description_
        """
        input_file = self.template_path.email_template
        output_file = 'Email_input.txt'

        find_text = 'XXXXX'

        with open(input_file, mode='r', encoding='utf-8') as file_in, open(output_file, 'w', encoding='utf-8') as file_out:
            for line in file_in:
                updated_line = line.replace(find_text, replacement_text)
                try:
                    file_out.write(updated_line)
                except UnicodeEncodeError:
                    print("CANNOT CREATE THIS FILE DUE TO INVALID NAME CHARACTERS. SKIPPING....")
                    print(f"Issue name is: {updated_line}")

        print(f"Text replaced and saved to '{output_file}'.")

    def store_email(self, input_string: str):
        """_summary_

        Args:
            input_string (_type_): _description_
        """
        file_name = "output.csv"
        with open(file_name, 'w', newline='', encoding='utf-8') as file:
            writer = csv.writer(file)
            writer.writerow([input_string])

        print(f"String saved to {file_name}")

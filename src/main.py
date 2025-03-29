# Certificate Generation Tooling

# Import
import pyfiglet

from util.util import print_title
from NameProcessor import NameProcessor
from Generator import CertificateGenerator
from util.template_path import TemplatePath

import os

def main() -> None:
    """
    Main entry point for the certificate generation tooling.
    """
    print_title("Certificate Generation Tooling")

    # File path should be loaded dynamically, ideally from an input directory.
    
    SOURCE_TEMPLATES =  TemplatePath(os.path.dirname(__file__))
    
    source_data = NameProcessor(file_path=SOURCE_TEMPLATES.excel_template)
    source_data.load_names()

    # With the names sourced, we can dynamically generate the certificates & email text file
    
    for name, email in source_data.excel_vals.items():
        # Create a certificate for each name
        print(f"Generating certificate for {name} with email {email}.")
        generator = CertificateGenerator(name, email, SOURCE_TEMPLATES)
        generator.CreateCerts()



if __name__ == "__main__":
    # Run the main function
    main()
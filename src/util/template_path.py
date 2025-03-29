import os

class TemplatePath:
    def __init__(self, base_dir):
        """
        Initialize the TemplatePath with the base directory.
        :param base_dir: The base directory where the templates folder is located.
        """
        self.base_dir = os.path.abspath(base_dir)
        self.templates_dir = os.path.join(self.base_dir, 'templates')

        # Initialize specific template paths
        self.word_template = os.path.join(self.templates_dir, 'Certificate.docx')
        self.email_template = os.path.join(self.templates_dir, 'Email_template.txt')
        self.excel_template = os.path.join(self.templates_dir, 'Attendance_List.xlsx')

        def get_word_template(self):
            """
            Get the path of the word template.
            :return: The absolute path to the word template.
            """
            return self.word_template

        def set_word_template(self, filename):
            """
            Set the path of the word template.
            :param filename: The name of the word template file.
            """
            self.word_template = os.path.join(self.templates_dir, filename)

        def get_email_template(self):
            """
            Get the path of the email template.
            :return: The absolute path to the email template.
            """
            return self.email_template

        def set_email_template(self, filename):
            """
            Set the path of the email template.
            :param filename: The name of the email template file.
            """
            self.email_template = os.path.join(self.templates_dir, filename)

        def get_excel_template(self):
            """
            Get the path of the excel template.
            :return: The absolute path to the excel template.
            """
            return self.excel_template

        def set_excel_template(self, filename):
            """
            Set the path of the excel template.
            :param filename: The name of the excel template file.
            """
            self.excel_template = os.path.join(self.templates_dir, filename)

        def get_template_path(self, template_name):
            """
            Get the absolute path of a template within the templates folder.
            :param template_name: The name of the template file.
            :return: The absolute path to the template file.
            """
            return os.path.abspath(os.path.join(self.templates_dir, template_name))
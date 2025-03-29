This folder MUST contain the following:

The source Attendance list spreadsheet titled: Attendance_List.xlsx
The Certificate template file titled: Certificate.docx
The email template file titled: Email_Template.txt

The source file path component of this code repository assumes they are in this folder with those file names. 

If you don't like this as the options, you can update the code with src/util/template_path accordingly.

**About the Attendance List spreadsheet**

This still assumes that the first sheet is titled "Attendees"
It also still assumes that the target column containing the names of the attendees is column C, and their emails are column F.

If you want to use a different format, these can be updated accordingly in src/NameProcessor.py.

FAQ's:

1. I saved the Attendance_List in the templates folder but it won't recognise it
    Are you sure? Did you save it as a xls instead of xlsx? 
    Running "Fix_Attendance.sh" should convert the file for you. Please note though that this is a speculative bash script assuming you've done something *predictably idiotic*. I'd recommend simply opening the problem file and saving it correctly. That's all the bash script does anyway.
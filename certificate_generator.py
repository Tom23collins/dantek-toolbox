import fitz
import openpyxl
import win32com.client
import re
import os

def generate_certificates(path_to_data, preview_email):
    delegates_processed = 0
    certificates_processed = 0
    emails_processed = 0
    course_failures = 0
    invalid_emails = 0
    data = openpyxl.load_workbook(path_to_data).active

    if not os.path.exists(f'{os.environ['USERPROFILE']}\\Desktop\\Certificates'):
        os.makedirs(f'{os.environ['USERPROFILE']}\\Desktop\\Certificates')

    for row in range(1, data.max_row):
        delegate_name = str(data['L'][row].value)
        delegate_email = str(data['O'][row].value)
        delegate_percentage = int((data['F'][row].value / 13) * 100)
        completion_date = str(data['C'][row].value).split(' ')[0].split('-')
        issue_date = f'{completion_date[2]}/{completion_date[1]}/{completion_date[0]}'
        expiry_date = f'{completion_date[2]}/{completion_date[1]}/{int(completion_date[0])+2}'
        delegates_processed += 1

        # Verify Score
        if not delegate_percentage > 70:
            course_failures += 1
        else:
            pdf_width = 595
            custom_font = f'{os.getcwd()}\\static\\OpenSans-SemiBold.ttf'
            doc = fitz.open(f'{os.getcwd()}\\static\\certificate_template.pdf')
            page = doc[0]
            page.insert_font(fontfile=custom_font, fontname="OpenSans")
            page.insert_text(fitz.Point((pdf_width / 2) - ((len(delegate_name) * 16 ) /2 ), 430), str(delegate_name), fontname="OpenSans", color=None, fontsize=30)
            page.insert_text(fitz.Point(46, 580), str(issue_date), fontname="OpenSans", color=None, fontsize=12)
            page.insert_text(fitz.Point(487, 580), str(expiry_date), fontname="OpenSans", color=None, fontsize=12)
            doc.save(f'{f'{os.environ['USERPROFILE']}\\Desktop\\Certificates'}\\Legionella Awareness Certificate - {delegate_name}.pdf')
            certificates_processed += 1

        # Verify Email
        email_pattern = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,7}\b'
        if not re.match(email_pattern,delegate_email):
            invalid_emails += 1
        else:
            ol=win32com.client.Dispatch("outlook.application")
            olmailitem=0x0 #size of the new email
            newmail=ol.CreateItem(olmailitem)
            newmail.Subject= 'Legionella Awareness Certificate'
            newmail.To= delegate_email
            newmail.Body= f'Hi {delegate_name.split(' ')[0]},\n\nCongratulations on completing your course on Legionella Awareness!\nPlease find attached your certificate.'

            attach= f'{f'{os.environ['USERPROFILE']}\\Desktop\\Certificates'}\\Legionella Awareness Certificate - {delegate_name}.pdf'
            newmail.Attachments.Add(attach)
            emails_processed += 1

            if preview_email:
                newmail.Display() 
            else:
                newmail.Send()

    return [delegates_processed, certificates_processed, emails_processed, course_failures, invalid_emails]
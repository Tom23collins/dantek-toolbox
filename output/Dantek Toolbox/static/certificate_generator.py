import fitz
import win32com.client
import re
import os
import pandas as pd

def generate_certificates(training_course, path_to_data):
    dataframe = pd.read_excel(path_to_data)

    delegates_processed = 0
    certificates_processed = 0
    emails_processed = 0
    course_failures = 0
    invalid_emails = 0

    if not os.path.exists(f'{os.environ['USERPROFILE']}\\Desktop\\Certificates'):
        os.makedirs(f'{os.environ['USERPROFILE']}\\Desktop\\Certificates')

    for row in range(1, len(dataframe.index)):
        delegate_name = dataframe['Candidate name'].values[row-1]
        delegate_email = dataframe['Email address'].values[row-1]
        delegate_percentage = int((dataframe['Total points'].values[row-1] / 13) * 100)
        completion_date = str(dataframe['Completion time'].values[row-1])[0:10].split('-')
        completion_date = completion_date
        issue_date = f'{completion_date[2]}/{completion_date[1]}/{completion_date[0]}'
        expiry_date = f'{completion_date[2]}/{completion_date[1]}/{int(completion_date[0])+3}'
        delegates_processed += 1

        # Verify Score
        if not delegate_percentage > 70:
            course_failures += 1
        else:
            #Create PDF
            pdf_width = 595
            custom_font = f'{os.getcwd()}\\static\\OpenSans-SemiBold.ttf'
            doc = fitz.open(f'{os.getcwd()}\\static\\certificate_template.pdf')
            page = doc[0]
            page.insert_font(fontfile=custom_font, fontname="OpenSans")
            page.insert_text(fitz.Point((pdf_width / 2) - ((len(training_course.split(' - ')[0]) * 16 ) /2 ), 285), str(training_course.split(' - ')[0]), fontname="OpenSans", color=None, fontsize=30)
            page.insert_text(fitz.Point((pdf_width / 2) - ((len(training_course.split(' - ')[1]) * 16 ) /2 ), 320), str(training_course.split(' - ')[1]), fontname="OpenSans", color=None, fontsize=30)
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
                # Prepare Email
                ol=win32com.client.Dispatch("outlook.application")
                olmailitem=0x0
                newmail=ol.CreateItem(olmailitem)
                newmail.Subject= 'Legionella Awareness Certificate'
                newmail.To= delegate_email
                newmail.Body= f'Hi {delegate_name.split(' ')[0]},\n\nI hope you are well.\n\nPlease find attached your {training_course.split(' - ')[0]} {training_course.split(' - ')[1]} Training Certificate attended on {issue_date}.\n\nWe hope you enjoyed the course.\n\nWith kind regards\n\nEmma'

                attach= f'{f'{os.environ['USERPROFILE']}\\Desktop\\Certificates'}\\Legionella Awareness Certificate - {delegate_name}.pdf'
                newmail.Attachments.Add(attach)
                emails_processed += 1

                newmail.Display() 
            

    return [delegates_processed, certificates_processed, emails_processed, course_failures, invalid_emails]
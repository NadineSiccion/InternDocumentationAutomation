from classes import Report
from docx import Document
from docx.shared import Inches
from pathlib import Path
from datetime import date


def edit_document(template_path, output_path, image_path, data:dict):
    doc = Document(template_path)

    img_file1 = image_path + f'd{data["[DJNUM]"] + ' (1).PNG'}'
    img_file2 = image_path + f'd{data["[DJNUM]"] + ' (2).PNG'}'

    for paragraph in doc.paragraphs:
        for key, value in data.items():
            if key in paragraph.text:
                    for run in paragraph.runs:
                        run.text = run.text.replace(key, value)
                    print('Replaced "' + key + '" with "' + value + '".')
    
    print(f'looking for file "{img_file1}"')
    doc.add_picture(img_file1, width=Inches(6))
    print('Picture added.')

    print(f'looking for file "{img_file2}"')
    doc.add_picture(img_file2, width=Inches(6))
    print('Picture added.')
    
    doc.save(output_path)
    print('Saved output at ' + str(output_path) + '.')


# run if main
if __name__ == '__main__':
     cwd = Path.cwd()

     template_path = cwd /'templates'/'DAILY_JOURNAL_TEMPLATE.docx'
     output_path = cwd /'outputs'/'DailyJounalNo###.docx'
     image_path = cwd / 'imgs'
     report_date = date(2024, 9, 26)
     # image file names must follow the format "d# (1).PNG"
     uip_internship_report = Report('uip', 561, 6, template_path, output_path, image_path, report_date)

     uip_internship_report.update_hours()
     uip_internship_report.save_report()
     uip_internship_report.edit_document()


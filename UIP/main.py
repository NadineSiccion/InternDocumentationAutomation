from classes import Internship, Report
from docx import Document
from docx.shared import Inches


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
     pass
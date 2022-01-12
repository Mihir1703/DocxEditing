from docx import Document
from docx2pdf import convert
import pandas as pd
import os

# email_info = pd.read_csv('Book1.csv', header=None)
# print(email_info)

# info = [(Name of startup) for Name of startup in zip(email_info[:][1], email_info[:][2]) if '@gmail' in mail]


def main1(tc):

    template_file_path = 'E:/test/Express.docx'
    output_file_path = 'E:/test/'+'output/'+tc + '.docx'

    variables = {
        "Express Stores": tc,
        # "${EMPLOEE_TITLE}": "Software Engineer",
        # "${EMPLOEE_ID}": "302929393",
        # "${EMPLOEE_ADDRESS}": "דרך השלום מנחם בגין דוגמא",
        # "${EMPLOEE_PHONE}": "+972-5056000000",
        # "${EMPLOEE_EMAIL}": "example@example.com",
        # "${START_DATE}": "03 Jan, 2021",
        # "${SALARY}": "10,000",
        # "${SALARY_30}": "3,000",
        # "${SALARY_70}": "7,000",
    }

    template_document = Document(template_file_path)

    for variable_key, variable_value in variables.items():
        for paragraph in template_document.paragraphs:
            replace_text_in_paragraph(paragraph, variable_key, variable_value)

        for table in template_document.tables:
            for col in table.columns:
                for cell in col.cells:
                    for paragraph in cell.paragraphs:
                        replace_text_in_paragraph(
                            paragraph, variable_key, variable_value)

    template_document.save(output_file_path)


def replace_text_in_paragraph(paragraph, key, value):
    if key in paragraph.text:
        inline = paragraph.runs
        for item in inline:
            if key in item.text:
                item.text = item.text.replace(key, value)


if __name__ == '__main__':
    email_info = pd.read_csv('Book1.csv', header=None)
    info = [i for i in email_info[:][0]]
    # print(email_info)
    for num in info:
        # print(num)
        main1(num)
    convert("output/")
filelist = [f for f in os.listdir("E:/test/output") if f.endswith(".docx")]
for f in filelist:
    os.remove(os.path.join("E:/test/output", f))

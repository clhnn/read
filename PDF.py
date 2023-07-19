import json
import pandas as pd
import pdfplumber
from PyPDF2 import PdfReader

class PDFProcessor:
    def __init__(self, pdf_file):
        self.pdf_file = pdf_file
    
    def extract_images(self):
        reader = PdfReader(self.pdf_file)
        for i in range(len(reader.pages)):
            page = reader.pages[i]
            for j, image in enumerate(page.images):    
                with open(f"image{j+1}.png", "wb") as f:
                    f.write(image.data)

    def extract_tables(self):
        pdf = pdfplumber.open(self.pdf_file)
        result_df = pd.DataFrame()
        for i in range(len(pdf.pages)):
            page = pdf.pages[i]
            tables = page.extract_tables()
            namenum = 0
            if not tables:
                continue
            for table in tables:
                if namenum+1 > len(tables):
                    continue
                xlsx_name = [f'text{namenum+1}.xlsx']
                df_detail = pd.DataFrame(table[1:], columns=table[0])
                df_detail.to_excel(xlsx_name[0])
                namenum += 1

    def create_simple_outline(self):
        with open(self.pdf_file, 'rb') as file:
            pdf_reader = PdfReader(file)
            num_pages = len(pdf_reader.pages)
            outline = {
                'children': []
            }
            for page_num in range(num_pages):
                page = pdf_reader.pages[page_num]
                content = page.extract_text()
                page_node = {
                    'title': f'Page {page_num + 1}',
                    'page': page_num + 1,
                    'content': content
                }
                outline['children'].append(page_node)
        return outline

    def process_pdf(self):
        self.extract_images()
        self.extract_tables()
        pdf_outline = self.create_simple_outline()

        result_dict = {
            'outline': pdf_outline
        }

        return result_dict

    def process_and_export_json(self):
        result = self.process_pdf()
        json_result = json.dumps(result, indent=4, ensure_ascii=False)
        with open("PDFtxt.json", "w", encoding='utf-8') as file:
            file.write(json_result)
        print("JSON result exported to PDFtxt.json")

if __name__ == '__main__':
    pdf_processor = PDFProcessor("your file name.pdf")
    pdf_processor.extract_images()
    pdf_processor.extract_tables()
    pdf_processor.create_simple_outline()
    pdf_processor.process_pdf()
    pdf_processor.process_and_export_json()


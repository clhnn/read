import json
import pandas as pd
import pdfplumber
import os
import fitz
from PyPDF2 import PdfReader
import PyPDF2

class PDFProcessor:
    
    # 初始化方法，設定要處理的 PDF 檔案名稱    
    def __init__(self, pdf_file):
        self.pdf_file = pdf_file
        self.font_size_threshold = 12     #字體大小閾值，可根據需要進行調整
        
    # 讀取 PDF 的圖片並儲存成檔案   
    def extract_images(self):
        reader = PdfReader(self.pdf_file)
        for i in range(len(reader.pages)):
            page = reader.pages[i]
            for j, image in enumerate(page.images):    
                with open(f"image{j+1}.png", "wb") as f:
                    f.write(image.data)
                    
    # 讀取 PDF 的表格並儲存成獨立的 csv 檔案
    def extract_tables(self, odname=None):
        odname = None
        pdf_path = self.pdf_file
        pdf = pdfplumber.open(pdf_path)
        already_taken = 'False'
        count = 1

        for pagenum, page in enumerate(pdf.pages):
            print('>>checking table at page %d'%(pagenum))
            tables = page.extract_tables()

            if not tables:
                print('>>skipped table at page %d'%(pagenum))
                continue
            
            table_bottom = pdf.pages[pagenum].bbox[3]-pdf.pages[pagenum].find_tables()[-1].bbox[3]-30 <= pdf.pages[pagenum].chars[-1].get('y0')
            
            if table_bottom:
                if count > 1:
                    count -= 1
                    continue
                for t, table in enumerate(tables):
                    if odname != None:
                        csv_name = os.path.join(odname, f'table{t+1}_{pagenum+1}.csv',index=False)
                    else:
                        csv_name = f'table{t+1}_{pagenum+1}.csv'
                    if already_taken == 'True':
                        if table == tables[0]:
                            continue
                    if table == tables[-1] and (pagenum < len(pdf.pages)-1):
                        for c in range(1, len(pdf.pages)+1):
                            if pdf.pages[pagenum+c].extract_tables()[0] != pdf.pages[pagenum+c].extract_tables()[-1]:
                                break
                            if pdf.pages[pagenum+c].bbox[3]-pdf.pages[pagenum+c].find_tables()[-1].bbox[3]-30 > pdf.pages[pagenum+c].chars[-1].get('y0'):
                                break
                            count += 1
                        for page_table in range(1, count+1):
                            table += pdf.pages[pagenum+page_table].extract_tables()[0]
                        combined_table = pd.DataFrame(table[1:], columns = table[0])
                        combined_table.to_csv(csv_name,index=False)
                        already_taken = 'True'
                        continue
                    df_detail = pd.DataFrame(table[1:], columns = table[0])
                    df_detail.to_csv(csv_name,index=False)
        
            else:
                if count > 1:
                    count -= 1
                    continue
                for t2, table in enumerate(tables):
                    if odname != None:
                        csv_name = os.path.join(odname, f'table{t2+1}_{pagenum+1}.csv',index=False)
                    else:
                        csv_name = f'table{t2+1}_{pagenum+1}.csv'
                    if already_taken == 'True':
                        if table == tables[0]:
                            already_taken = 'False'
                            continue
                    df_detail = pd.DataFrame(table[1:], columns = table[0])
                    df_detail.to_csv(csv_name,index=False)
                    already_taken = 'False'
        return
    
    # 讀取 PDF 頁首頁碼內文
    def classify_text_by_font_size(self):
        page_texts = {'header': [], 'content': []}
        pdf_document = fitz.open(self.pdf_file)
        for page_num in range(pdf_document.page_count):
            page = pdf_document[page_num]
            text = page.get_text()
            header_text = ''
            content_text = ''
            for line in text.split('\n'):
                if line.strip():
                    font_size = None
                    for block in page.get_text("dict")['blocks']:
                        for line_info in block['lines']:
                            for span in line_info['spans']:
                                if line in span['text']:
                                    font_size = span['size']
                                    break
                            if font_size is not None:
                                break
                        if font_size is not None:
                            break
    
                    if font_size is not None and font_size < self.font_size_threshold:
                        header_text += line + '\n'
                    else:
                        content_text += line + '\n'
    
            # Store classified text in dictionary
            page_texts['header'].append(header_text)
            content_all_text = self.extract_paragraphs(content_text, page_num)  # Fix: remove the extra argument
            page_texts['content'].append(content_all_text)
        return page_texts

    # 讀取 PDF 的每頁的段落
    def extract_paragraphs(self,text,page_num):
        paragraphs = text.split(' \n') 
        all_para=[]
        while '' in paragraphs:
            paragraphs.remove('')
        for i, paragraph in enumerate(paragraphs):
            paragraph_number = f"Page {page_num + 1}, Paragraph {i + 1}:{paragraph}\n"
            #print(paragraph_number)
            all_para.append(paragraph_number)
        return all_para
                                
    # 建立 PDF 的簡單大綱，包含頁碼和內容
    def create_simple_outline(self):
        with open(self.pdf_file, 'rb') as file:
            pdf_reader = PdfReader(file)
            num_pages = len(pdf_reader.pages)
            outline = {
                'children': []
            }
            content_all=self.classify_text_by_font_size()
            for page_num in range(num_pages):
                for category, texts in content_all.items():
                    content = texts[page_num]
                    if category == 'header':
                        page_node = {
                            'title': f'Page {page_num + 1}',
                            'page': page_num + 1,
                            'category': category,
                            'header': content,
                            
                        }
                        outline['children'].append(page_node)
                    else:
                        page_node = {
                            'title': f'Page {page_num + 1}',
                            'page': page_num + 1,
                            'category': category,
                            'content': content,
                            
                        }
                        outline['children'].append(page_node)
        return outline
        
    # 處理 PDF，提取圖片、表格並建立大綱
    def process_pdf(self):
        self.extract_images()
        self.extract_tables()
        pdf_outline = self.create_simple_outline()

        result_dict = {
            'outline': pdf_outline
        }
        return result_dict
        
    # 處理 PDF 並將結果匯出成 JSON 檔案
    def process_and_export_json(self):
        result = self.process_pdf()
        json_result = json.dumps(result, indent=4, ensure_ascii=False)
        with open("PDFtxt.json", "w", encoding='utf-8') as file:
            file.write(json_result)
        print("JSON result exported to PDFtxt.json")
        
# 主程式部分
if __name__ == '__main__':
    
    # 請將 "your file name.pdf" 替換為要處理的實際 PDF 檔案名稱
    pdf_processor = PDFProcessor("demo2-5.pdf")

    #從PDF中提取圖片
    pdf_processor.extract_images()

    # 從 PDF 中提取表格並儲存為單獨的 csv 檔案
    pdf_processor.extract_tables()

    # 建立包含頁碼和內容的 PDF 簡單大綱
    pdf_processor.create_simple_outline()

    # 處理 PDF，提取圖片、表格並建立大綱
    pdf_processor.process_pdf()

    # 處理 PDF 並將結果匯出成 JSON 檔案
    pdf_processor.process_and_export_json()

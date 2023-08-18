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
        
    # 讀取 PDF 的圖片並儲存成檔案   
    def extract_images(self):
        reader = PdfReader(self.pdf_file)
        for i in range(len(reader.pages)):
            page = reader.pages[i]
            for j, image in enumerate(page.images):    
                with open(f"image{j+1}.png", "wb") as f:
                    f.write(image.data)
                    
    # 讀取 PDF 的表格並儲存成獨立的 csv 檔案
    def extract_tables(self):
        pdf_path = self.pdf_file
        pdf = pdfplumber.open(pdf_path)
        already_taken = 'False'
        count = 1
        csv_name = []
        csv_name_up = []
        csv_name_down = []
        
        for pagenum, page in enumerate(pdf.pages):
            if pagenum > 9:
                break
            print('>>checking table at page %d'%(pagenum))
            tables = page.extract_tables()
            table_num = 0
        
            if not tables:
                print('>>skipped table at page %d'%(pagenum))
                continue
            chars = pdf.pages[pagenum].chars
            table_bottom = pdf.pages[pagenum].bbox[3]-pdf.pages[pagenum].find_tables()[-1].bbox[3]-30 <= pdf.pages[pagenum].chars[-1].get('y0')
    
            if table_bottom:
                if count > 2:
                    count -= 1
                    continue
                if count == 2:
                    count -=1
                for t, table in enumerate(tables):
                    if already_taken == 'True':
                        if table == tables[0]:
                            table_num += 1
                            already_taken = 'False'
                            continue
                    length_up = []
                    length_down = []
                    csv_name = []
                    csv_name_up = []
                    csv_name_down = []
                    if table == tables[-1] and (pagenum < len(pdf.pages)-1):
                        count += 1
                        for c in range(1, len(pdf.pages)+1):
                            if pdf.pages[pagenum+c].extract_tables()[0] != pdf.pages[pagenum+c].extract_tables()[-1]:
                                break
                            if pdf.pages[pagenum+c].bbox[3]-pdf.pages[pagenum+c].find_tables()[-1].bbox[3]-30 > pdf.pages[pagenum+c].chars[-1].get('y0'):
                                break
                            count += 1
                        for page_table in range(1, count):
                            table += pdf.pages[pagenum+page_table].extract_tables()[0]
                    chars2 = pdf.pages[pagenum+count-1].chars
                    table_y0 = pdf.pages[pagenum].find_tables()[table_num].bbox[1]
                    table_y1 = pdf.pages[pagenum].find_tables()[table_num].bbox[3]
                    if count > 1:
                        table_y1 = pdf.pages[pagenum+count-1].find_tables()[0].bbox[3]
                    for char in range(0, len(chars)):
                        char_info = pdf.pages[pagenum].chars[char]
                        if char_info.get('bottom') < table_y0:
                            bottom = int(table_y0-(char_info.get('bottom')))
                            length_up.append(bottom)
                    for char in range(0, len(chars2)):
                        char_info = pdf.pages[pagenum+count-1].chars[char]
                        if char_info.get('top') > table_y1:
                            top = int(char_info.get('top')-table_y1)
                            length_down.append(top)
                    for char in range(0, len(chars)):
                        char_info = pdf.pages[pagenum].chars[char]
                        if char_info.get('bottom') < table_y0:
                            lone_up = int(table_y0-(char_info.get('bottom')))
                            if lone_up == min(length_up):
                                csv_name_up += char_info.get('text')
                                csv_name_up = "".join(csv_name_up)
                    for char in range(0, len(chars2)):
                        char_info = pdf.pages[pagenum+count-1].chars[char]
                        if char_info.get('top') > table_y1:
                            lone_down = int(char_info.get('top')-table_y1)
                            if lone_down == min(length_down):
                                csv_name_down += char_info.get('text')
                                csv_name_down = "".join(csv_name_down)
                    if csv_name_up == ' ':
                        csv_name = f'{csv_name_down}.csv'
                    else:
                        csv_name = f'{csv_name_up}.csv'
                    if count > 1:
                        combined_table = pd.DataFrame(table[1:], columns = table[0])
                        combined_table.rename(columns={None: '#'}, inplace=True)
                        for index, row in combined_table.iterrows():
                            for col in range(0, len(combined_table.columns)):
                                if combined_table.iat[index, col] == None:
                                    combined_table.iat[index, col] = '#'
                        combined_table.to_csv(csv_name, index=False, encoding='utf-8')
                        table_num += 1
                        already_taken = 'True'
                        continue
                    df_detail = pd.DataFrame(table[1:], columns = table[0])
                    df_detail.rename(columns={None: '#'}, inplace=True)
                    for index, row in df_detail.iterrows():
                        for col in range(0, len(df_detail.columns)):
                            if df_detail.iat[index, col] == None:
                                df_detail.iat[index, col] = '#'
                    df_detail.to_csv(csv_name, index=False, encoding='utf-8')
                    table_num += 1
        
            else:
                if count > 2:
                    count -= 1
                    continue
                for t2, table in enumerate(tables):
                    if already_taken == 'True':
                        if table == tables[0]:
                            table_num += 1
                            already_taken = 'False'
                            continue
                    length_up = []
                    length_down = []
                    csv_name = []
                    csv_name_up = []
                    csv_name_down = []
                    table_y0 = pdf.pages[pagenum].find_tables()[table_num].bbox[1]
                    table_y1 = pdf.pages[pagenum].find_tables()[table_num].bbox[3]
                    for char in range(0, len(chars)):
                        char_info = pdf.pages[pagenum].chars[char]
                        if char_info.get('bottom') < table_y0:
                            bottom = int(table_y0-(char_info.get('bottom')))
                            length_up.append(bottom)
                        elif char_info.get('top') > table_y1:
                            top = int(char_info.get('top')-table_y1)
                            length_down.append(top)
                    for char in range(0, len(chars)):
                        char_info = pdf.pages[pagenum].chars[char]
                        if char_info.get('bottom') < table_y0:
                            lone_up = int(table_y0-(char_info.get('bottom')))
                            if lone_up == min(length_up):
                                csv_name_up += char_info.get('text')
                                csv_name_up = "".join(csv_name_up)
                        elif char_info.get('top') > table_y1:
                            lone_down = int(char_info.get('top')-table_y1)
                            if lone_down == min(length_down):
                                csv_name_down += char_info.get('text')
                                csv_name_down = "".join(csv_name_down)
                    if csv_name_up == ' ':
                        csv_name = f'{csv_name_down}.csv'
                    else:
                        csv_name = f'{csv_name_up}.csv'
                    df_detail = pd.DataFrame(table[1:], columns = table[0])
                    df_detail.rename(columns={None: '#'}, inplace=True)
                    for index, row in df_detail.iterrows():
                        for col in range(0, len(df_detail.columns)):
                            if df_detail.iat[index, col] == None:
                                df_detail.iat[index, col] = '#'
                    df_detail.to_csv(csv_name, index=False, encoding='utf-8')
                    table_num += 1
                    already_taken = 'False'
        return
    
    #讀取 PDF 頁內文
    def classify_text_by_row_data(self):
        content_texts = {'content' : []}
        content = []
        pdf_document = fitz.open(self.pdf_file)
        for page_num in range(pdf_document.page_count):
            page = pdf_document[page_num]
            text = self.extract_paragraphs(page.get_text(),page_num)
            content.append(text)
            #content.append(page)
        content_texts['content'] = content
        return content_texts
    
    # 讀取 PDF 頁首頁碼內文
    def classify_text_by_font_size(self):
        header_texts = {'header': []}
        content_texts = {'content': []}
        texts = []
        item = ''
        pdf_document = fitz.open(self.pdf_file)
        paragraph_texts = ''
        for page_num in range(pdf_document.page_count):
            page = pdf_document[page_num]
            text = page.get_text()
            paragraph_text = []
            for line in text.split('\n'):
                if line.strip():
                    line = line.strip()
                    font_size = None
                    for block in page.get_text("dict")['blocks']:
                        if 'lines' in block:
                            for line_info in block['lines']:
                                for span in line_info['spans']:
                                    if line in span['text']:
                                        font_size = span['size']
                                        break
                                if font_size is not None:
                                    break
                        if font_size is not None:
                            break
                    if self.is_line_potential_size(line) :
                        item += line
                    else:
                        if item != '':
                            item += line
                            paragraph_text.append(item)
                            item = ''
                        paragraph_text.append(line)
                else:
                    paragraph_text.append('##')    ##為空白行                    
            content_all_text=self.extract_paragraphs(paragraph_text,page_num)
            content_texts['content'].append(content_all_text)
            texts = []
        return content_texts
    
    #判斷line大小
    def is_line_potential_size(self,line):
    ##如果item大小不小於4可設正列
     #  item =['1.','2.,.....]
     #  if line == item:
     #     return
        if len(line) <= 4 :
            return True
        return False

    # 讀取 PDF 的每頁的段落
    def extract_paragraphs(self,text,page_num):
        #page_text = text.strip('\n')
        #paragraphs = page_text.split(' \n')
        paragraphs = text 
        return paragraphs  
                         
    # 建立 PDF 的簡單大綱，包含頁碼和內容
    def create_simple_outline(self):
        with open(self.pdf_file, 'rb') as file:
            pdf_reader = PdfReader(file)
            num_pages = len(pdf_reader.pages)
            outline = {
                'file_info':[],
                'pages':[]
            }
            file_info={
                'filename':self.pdf_file,
                'pages':num_pages
                }
            outline['file_info'].append(file_info)
            content_texts=self.classify_text_by_font_size()  #PDF text 依行及font size隔開
            #content_texts=self.classify_text_by_row_data()  #PDF text 依整頁進行讀取
            for page_num in range(num_pages):
                for category2, texts2 in content_texts.items():
                    content = texts2[page_num]
                    page_node = {
                        'page': page_num + 1,
                        'paragraphs' : content
                        }
                    outline['pages'].append(page_node)
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
        return
        
# 主程式部分
if __name__ == '__main__':
    
    # 請將 "your file name.pdf" 替換為要處理的實際 PDF 檔案名稱
    pdf_processor = PDFProcessor("your name file.pdf")

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

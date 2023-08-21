# **Introduction**
此示範程式皆為Python

## PDF
'PDFProcessor'類用於處理PDF文件，包括提取圖片、提取表格並保存為獨立的Excel文件，創建PDF的簡單大綱，以及將處理結果導出為JSON文件。該類包含以下方法

`注意!在執行程式前請確保已安裝所需的Python庫:'json'、'pandas'、'pdfplumber' 和 'PyPDF2'`

###### 初始化方法
初始化方法用於設置要處理的PDF文件名
```js
# 初始化方法，設定要處理的 PDF 檔案名稱
def __init__(self, pdf_file):
    self.pdf_file = pdf_file
```

###### 提取圖片
'extract_images'方法從PDF文件中提取圖片，並將每個圖片保存為單獨的PNG文件
```js
# 讀取 PDF 的圖片並儲存成檔案
def extract_images(self):
    reader = PdfReader(self.pdf_file)
    for i in range(len(reader.pages)):
        page = reader.pages[i]
        for j, image in enumerate(page.images):    
            with open(f"image{j+1}.png", "wb") as f:
                f.write(image.data)
```

###### 提取表格
'extract_tables'方法從PDF文件中提取表格，並將每個表格保存為單獨的csv文件，而不規則表格中如出現合併表格，除了輸出表示該格的值，其餘格子將輸出'#'
```js
def extract_tables(self):
    pdf_path = self.pdf_file
    pdf = pdfplumber.open(pdf_path)
    already_taken = 'False'
    count = 1
    csv_name = []
    csv_name_up = []
    csv_name_down = []

    for pagenum, page in enumerate(pdf.pages):
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
```

###### 讀取內文
`有兩種分類方法`

1.'classify_text_by_row_data()' 函數：這個方法將文本按行進行分類，不考慮字體大小。適用於文本結構較簡單的 PDF
```js
def classify_text_by_row_data(self):
    content_texts = {'content' : []}
    content = []
    pdf_document = fitz.open(self.pdf_file)
    for page_num in range(pdf_document.page_count):
        page = pdf_document[page_num]
        text = self.extract_paragraphs(page.get_text(),page_num)
        content.append(text)
    content_texts['content'] = content
    return content_texts
```

2.'classify_text_by_font_size()' 函數：將 PDF 文件中的文字按照字體大小分類。它會遍歷每一頁，然後根據字體大小閾值將文字分為不同的段落，並返回每個段落的內容
```js
def classify_text_by_font_size(self):
    header_texts = {'header': []}
    content_texts = {'content': []}
    texts = []
    item = ''
    pdf_document = fitz.open(self.pdf_file)
    paragraph_texts = ''
    for page_num in range(pdf_document.page_count):
        all_para={}
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
    return content_texts
```

###### 判斷line大小
'is_line_potential_size()'此函數用於判斷'classify_text_by_font_size()'給定的文本行是否具有潛在的大小特徵，並根據一個特定的標準來確定是否可以將該行視為有效行
`##為若遇到長度判斷問題時能進行修改成正列假設`
```js
def is_line_potential_size(self,line):
##如果item大小不小於4可設正列
     #  item =['1.','2.,.....]
     #  if line == item:
     #     return
        if len(line) <= 4 :
            return True
        return False
```

###### 讀取 PDF 的每頁的段落
'extract_paragraphs()'此函數獲取文本行列表並將它們分組為段落。它消除空行並為每頁生成段落列表。函數內部調用該函數classify_text_by_font_size()來處理每一頁文本的行
```js
def extract_paragraphs(self,text,page_num):
    #page_text = text.strip('\n')
    #paragraphs = page_text.split(' \n')
    paragraphs = text 
    return paragraphs
```

###### 創建簡單大綱
'create_simple_outline'方法創建簡單大綱，包括每一頁的頁碼和內容。大綱以字典形式返回
```js
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
            content_texts=self.classify_text_by_font_size()
            for page_num in range(num_pages):
                for category2, texts2 in content_texts.items():
                    content = texts2[page_num]
                    page_node = {
                        'page': page_num + 1,
                        'paragraphs' : content
                        }
                    outline['pages'].append(page_node)
        return outline
```

###### 處理PDF
'process_pdf'方法是綜合處理PDF的方法，他調用'extract_images'、'extract_tables'和'create_simple_outline'方法，提取圖片、表格並創建大綱。最終將結果以字典形式返回
```js
# 處理 PDF，提取圖片、表格並建立大綱
def process_pdf(self):
    self.extract_images()
    self.extract_tables()
    pdf_outline = self.create_simple_outline()

    result_dict = {
        'outline': pdf_outline
    }
    return result_dict
```

###### 導出JSON
'process_and_export_json'方法用於處理PDF並將結果導出為JSON文件。他調用'_pdf'方法獲取處理結果，並將結果以JSON格是寫入名為'PDFtxt.json'的文件中
```js
# 處理 PDF 並將結果匯出成 JSON 檔案
def process_and_export_json(self):
    result = self.process_pdf()
    json_result = json.dumps(result, indent=4, ensure_ascii=False)
    with open("PDFtxt.json", "w", encoding='utf-8') as file:
        file.write(json_result)
    print("JSON result exported to PDFtxt.json")
```

###### 主程式
在主程式部分，你可以創建一個'PDFProcessor'對象並調用其方法來處理PDF文件。以下是可以執行的操作:
* 從PDF中提取圖片並保存為PNG文件
* 從PDF中提取表格並保存為獨立的csv文件
* 創建包含頁碼和內容的PDF簡單大綱
* 處理PDF，提取圖片、表格和大綱
* 將處理結果導出為JSON文件
```js
#主程式部分
if __name__ == '__main__':
    pdf_processor = PDFProcessor("your file name.pdf")
    pdf_processor.extract_images()
    pdf_processor.extract_tables()
    pdf_processor.create_simple_outline()
    pdf_processor.process_pdf()
    pdf_processor.process_and_export_json()
```

`注意!在執行程式前請確保'your file name.pdf'為實際的PDF文件名`

## Word
'Word'類圖供了處理Word文檔的功能，包括讀取文本段落、提取圖片和讀取表格內容。該類包含以下方法

`注意!請確保在執行程式前，請確保已安裝所需的Python庫:'json'和'docx'`

###### 初始化方法
初始化方法用於設置要處理的Word文檔名
```js
#初始化方法，設定要處理的 Word 檔案
def __init__(self, doc):
    self.doc = docx.Document(doc)
```

###### 讀取文本段落
'read_text'方法用於從Word文檔中讀取所有文本段落。並以列表行形式返回
```js
# 讀取 Word 文件中的所有文本段落
def read_text(self):
    paragraphs = self.doc.paragraphs
    text_list = [paragraph.text for paragraph in paragraphs]
    return text_list
```

###### 提取圖片
'extract_images'方法用於提取Word文檔中的所有圖片，並將每個圖片保存為單獨的PNG文件。方法返回保存圖片的文件名列表
```js
# 提取 Word 文件中的所有圖片並儲存成檔案
def extract_images(self):
    rels = self.doc.part.rels
    images = []     
    for rel in rels:
        rel_type = rels[rel].reltype
        if "image" in rel_type:
            image_part = rels[rel]._target
            image_data = image_part.blob
            image_filename = f"image{len(images)+1}.png"
            with open(image_filename, "wb") as f:
                f.write(image_data)
            images.append(image_filename)
    return images
```

###### 讀取表格內容
'read_table'方法用於讀取Word文檔中表格的內容<並將內容已遷到列表的形式返回。每個列表都表示為一個列表，其中每個字典代表一個表格行，字典的鍵是表頭，值是表格單元格的內容
```js
# 讀取 Word 文件中的表格內容並返回列表形式
def read_table(self):
    tables_data = []
    for table in self.doc.tables:
        data = []
        keys = None
        for i, row in enumerate(table.rows):
            text = [cell.text.strip() for cell in row.cells]
            if i == 0:
                keys = tuple(text)
                continue
            row_data = dict(zip(keys, text))
            data.append(row_data)
        tables_data.append(data)
    return tables_data
```

###### 讀取標題樣式段落
'readheading'方法用於讀取Word文檔中的標題樣式段落，並以列表形式返回
```js
# 讀取 Word 文件中的標題樣式段落並返回列表
def readheading(self):
    paragraphs = self.doc.paragraphs
    heading=[]
    for paragraph in paragraphs:
        if paragraph.style.name.startswith('Heading'):
            heading.append(paragraph.text)
    return heading
```

###### 主程式
在主程式部分，你可以創建一個'Word'最像並調用其他方法處理Word文檔。以下是可以執行的操作:

```js
if __name__ == '__main__':
    word = Word("your file name.docx")
    text_list = word.read_text()
    images_list = word.extract_images()
    tables_data = word.read_table()
    heading = word.readheading()

    result = {
        'Heading': heading,
        'text': text_list,
        'images': images_list,
        'tables': tables_data        
    }

    with open('wordtxt.json', 'w') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print("JSON result exported to wordtxt.json")
```
通過調用相應的方法或去處理結果，並將結果整合成一個字典。最後，將結果以JSON格式寫入名為'wordtxt.json'的文件中

`注意!請確保在執行程式前，將'your file name.docx'更改為您將使用的Word文檔名`

# Install
Python適用

## PDF

```js
pip install PyPDF2
```
PyPDF2是一個用於處理PDF文件的庫，它可以讀取和操作PDF文件。此程式中，其用於讀取PDF文件的頁面並提取頁面中的圖像數據

```js
pip install pdfplumber
```
pdfplumber是一個用於處理PDF文件的庫。它提供提取文本、圖像、表格等內容的功能。在此程式中，pdfplumber用於提取PDF文件中的表格數據

```js
pip install pandas
```
Pandas是一個強大的樹去分析處理庫，它提供了綾羅高校的數據結構和數據分析工具。在此程式中，pdfplumber用於提取PDF文件中的表格數據。

## Word

```js
pip install docx
```
用於處理Microsoft Word文檔(.docx)的Python庫。他提供一組功能，用於讀取、編輯和創建Word文檔，在此程式碼中，'docx'庫被用於讀取Word文檔的內容。包括文本段落、表格和圖片

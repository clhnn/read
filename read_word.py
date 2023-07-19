from word import Word
import json
if __name__ == '__main__':
    
    # 請將 "your file name.docx" 替換為要處理的實際 Word 檔案名稱
    word = Word("your file name.docx")

    # 讀取 Word 文件中的所有文本段落
    text_list = word.read_text()

    # 提取 Word 文件中的所有圖片並儲存成檔案
    images_list = word.extract_images()

    # 讀取 Word 文件中的表格內容並返回列表形式
    tables_data = word.read_table()

    # 讀取 Word 文件中的標題樣式段落並返回列表
    heading = word.readheading()

    # 將結果整合成字典
    result = {
        'Heading': heading,
        'text': text_list,
        'images': images_list,
        'tables': tables_data        
    }

    # 將結果寫入 JSON 檔案
    with open('wordtxt.json', 'w') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print("JSON result exported to wordtxt.json")   

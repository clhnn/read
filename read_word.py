from word import Word
import json
if __name__ == '__main__':
    word = Word("111025042.docx")
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
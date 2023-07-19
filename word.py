import docx
import json

class Word:
    def __init__(self, doc):
        self.doc = docx.Document(doc)

    def read_text(self):
        paragraphs = self.doc.paragraphs
        text_list = [paragraph.text for paragraph in paragraphs]
        return text_list

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
    
    def readheading(self):
        paragraphs = self.doc.paragraphs
        heading=[]
        for paragraph in paragraphs:
            if paragraph.style.name.startswith('Heading'):
                heading.append(paragraph.text)
        return heading
 
       
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

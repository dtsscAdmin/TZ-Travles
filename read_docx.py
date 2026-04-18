import zipfile
import xml.etree.ElementTree as ET

def read_docx(path):
    with zipfile.ZipFile(path) as docx:
        tree = ET.XML(docx.read('word/document.xml'))
        text = []
        for node in tree.iter():
            if node.tag == '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t' and node.text:
                text.append(node.text)
            elif node.tag == '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p':
                text.append('\n')
        return ''.join(text)

print(read_docx('c:/Users/USER/Downloads/TripZoneTravels_Proposal.docx'))

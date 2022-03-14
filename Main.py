import xml.etree.ElementTree as ET
import zipfile
import re

def get_docx_text(path):
    """
    Take the path of a docx file as argument, return the text in unicode.
    """
    document = zipfile.ZipFile(path)
    xml_content = document.read('word/document.xml')
    document.close()

    xml_structured = ET.fromstring(xml_content)
    
    return xml_structured

def addPropertiesMarks(xml):

   prefix = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
   text = ""
   for body in xml:
      for paragraph in body:
         text += "<p>"
         for sentence in paragraph:
            text += ''.join(sentence.itertext())
            for rProrT in sentence:
               if rProrT.findall(prefix+"b"):
                  targetSentence = ''.join(sentence.itertext())
                  textTag = "<b>" + targetSentence + "</b>"
                  text = re.sub(targetSentence, textTag, text)
               
               if rProrT.findall(prefix+"u"):
                  targetSentence = ''.join(sentence.itertext())
                  textTag = "<u>" + targetSentence + "</u>"
                  text = re.sub(targetSentence, textTag, text)

               if rProrT.findall(prefix+"i"):
                  targetSentence = ''.join(sentence.itertext())
                  textTag = "<i>" + targetSentence + "</i>"
                  text = re.sub(targetSentence, textTag, text)

         text += "</p>"

   return text

xml_text = get_docx_text("Lorem ipsum dolor sit amet.zip")
x = addPropertiesMarks(xml_text)
print(x)

ET.dump(xml_text)
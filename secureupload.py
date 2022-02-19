import os
from zipfile import ZipFile
import shutil
import time
from PIL import Image
from PyPDF2 import PdfFileWriter, PdfFileReader
from PyPDF2.generic import NameObject, createStringObject


path_1 = input("What is the path of your file ?")
directory = os.path.dirname(path_1)
name_ = os.path.basename(path_1)
name, ext = name_.split(".")

localdir = os.getcwd()


datadocx = (os.getcwd())
if ext == ("docx"):
    zipfile = path_1 + ".zip"
    
    os.rename(path_1, zipfile)
    
    with ZipFile(zipfile, 'r') as zip:
        zip.extractall(datadocx)
    path = "docProps\\core.xml".format(os.getcwd())
    file = open(path, "r")
    header, content = file.readlines()

    position1 = content.find("<dc:creator>")
    position2 = content.find("</dc:creator>")
    position3 = content.find("<cp:lastModifiedBy>")
    position4 = content.find("</cp:lastModifiedBy>")
    

    remove_stepone = content[position1+12: position2]
    remove_steptwo = content[position3 + 19: position4]


    content1 = content.replace(remove_stepone, "")
    content2 = content1.replace(remove_steptwo, "")
    file.close()
    file = open(path, "r")
    


    data = file.read()
    file.seek(0)
    file.close()
    file = open(path, "wb")
    bheader = header.encode("utf-8")
    bcontent = content2.encode("utf-8")
    file.write(bheader + bcontent)
    file.truncate() 
    file.close()

    with ZipFile("doc.zip", 'w') as zip:
        zip.write("_rels\\")
        zip.write("docProps\\")
        zip.write("word\\")
        zip.write("word\\_rels")
        zip.write("word\\theme")
        zip.write("word\\document.xml")    
        zip.write("word\\fontTable.xml")
        zip.write("word\\settings.xml")
        zip.write("word\\styles.xml")
        zip.write("word\\webSettings.xml")
        zip.write("word\\_rels\\document.xml.rels")
        zip.write("xl\\theme\\theme1.xml")
        zip.write("docProps\\app.xml")
        zip.write("docProps\\core.xml")
        zip.write('[Content_Types].xml')
        zip.write("_rels\\.rels")
    
    os.rename("doc.zip", "{}e.docx".format(name_))
    shutil.move("{}_(cleaned).docx".format(name), directory)
    shutil.rmtree("xl\\")
    shutil.rmtree("_rels\\")
    os.remove("[Content_Types].xml")
    os.remove(zipfile)
    shutil.rmtree("docProps\\")
if ext == ("xlsx"):
    zipfile = path_1 + ".zip"
    
    os.rename(path_1, zipfile)
    
    with ZipFile(zipfile, 'r') as zip:
        zip.extractall(datadocx)
    path = "docProps\\core.xml".format(os.getcwd())
    file = open(path, "r")
    header, content = file.readlines()

    position1 = content.find("<dc:creator>")
    position2 = content.find("</dc:creator>")
    position3 = content.find("<cp:lastModifiedBy>")
    position4 = content.find("</cp:lastModifiedBy>")
    

    remove_stepone = content[position1+12: position2]
    remove_steptwo = content[position3 + 19: position4]


    content1 = content.replace(remove_stepone, "")
    content2 = content1.replace(remove_steptwo, "")
    file.close()
    file = open(path, "r")
    


    data = file.read()
    file.seek(0)
    file.close()
    file = open(path, "wb")
    bheader = header.encode("utf-8")
    bcontent = content2.encode("utf-8")
    file.write(bheader + bcontent)
    file.truncate() 
    file.close()

    with ZipFile("doc.zip", 'w') as zip:
        zip.write("_rels\\")
        zip.write("_rels\\.rels")
        zip.write("docProps\\")
        zip.write("xl\\")
        zip.write("xl\\_rels")
        zip.write("xl\\theme")
        zip.write("xl\\worksheets")
        worksheet = True
        sheet = 1
            
        while worksheet:
            try:
                zip.write("xl\\worksheets\\sheet{}.xml".format(sheet))
                sheet = sheet+1
            except FileNotFoundError():
                worksheet = False
        theme = 1
        worksheet = True
        while worksheet:
            try:
                zip.write("xl\\worksheets\\sheet{}.xml".format(theme))
                theme = theme +1
            except FileNotFoundError():
                worksheet = False
            
        
        zip.write("docProps\\app.xml")
        zip.write("docProps\\core.xml")
        zip.write('[Content_Types].xml')
        zip.write("\\xl\\_rels\\workbook.xml.rels")
    
    os.rename("doc.zip", "{}e.xlsx".format(name_))
    shutil.move("{}_(cleaned).xlsx".format(name), directory)
    shutil.rmtree("xl\\")
    shutil.rmtree("_rels\\")
    os.remove("[Content_Types].xml")
    os.remove(zipfile)
    shutil.rmtree("docProps\\")
elif ext == ("jpg" or "png"):
    image = Image.open(path_1)
    image_data = list(image.getdata())
    image_without_exif = Image.new(image.mode, image.size)
    image_without_exif.putdata(image_data)
    image_without_exif.save(u""+ directory +"\\cleaned_{}".format(name))
elif ext == ("pdf"):
    input_pdf = PdfFileReader(path_1, strict=False)
    info = input_pdf.getDocumentInfo()
    output = PdfFileWriter()
    infoDict = output._info.getObject()
    infoDict.update({NameObject('/Title'): createStringObject(u'e'),NameObject('/Author'): createStringObject(u'e'),NameObject('/Subject'): createStringObject(u'e'),NameObject('/Creator'): createStringObject(u'e'),NameObject('/Producer'): createStringObject(u'e'),NameObject('/Keywords'): createStringObject(u'z')})
    for page in range(input_pdf.getNumPages()):
        output.addPage(input_pdf.getPage(page))
    outputStream = open(path_1 + "{}_clear.pdf", 'wb'.format(name))
    output.write(outputStream)

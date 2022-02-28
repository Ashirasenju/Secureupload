import os
from zipfile import ZipFile
import shutil
import time
from PIL import Image
from PyPDF2 import PdfFileWriter, PdfFileReader
from PyPDF2.generic import NameObject, createStringObjec
import logging
# lets configure the program
path_1 = input("What is the path of your file ?")
directory = os.path.dirname(path_1)
name_ = os.path.basename(path_1)
name, ext = name_.split(".")

localdir = os.getcwd()
logging.basicConfig(filename="logs.log", filemode="w", format="%(levelname)s - %(message)s")

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
elif ext == ("xlsx"):
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
elif ext == ("pptx"):
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
        zip.write("docProps\\app.xml")
        zip.write("docProps\\core.xml")
        zip.write("ppt\\")
        zip.write("ppt\\_rels")
        zip.write("ppt\\_rels\\presentation.xml.rels")
        slidelayout = True
        boucle = 1
            
        while slidelayout:
            try:
                zip.write("ppt\\slideLayouts\\slideLayout{}.xml".format(boucle))
                boucle = boucle+1
            except FileNotFoundError():
                slidelayout = False
        zip.write("[Content_Types].xml")
        zip.write("\\ppt\\slideLayouts\\_rels")
        
        
        slideMaster = True
        boucle = 1
        
        while slideMaster:
            try:
                zip.write("ppt\\slideMasters\\slideMaster{}.xml".format(boucle))
                boucle = boucle +1
            except FileNotFoundError():
                slideMaster = False

        slides = True
        boucle = 1
        
        while slides:
            try:
                zip.write("ppt\\slides\\slide{}.xml".format(boucle))
                boucle = boucle + 1
            except FileNotFoundError():
                slides = False
        
        theme = True
        boucle = 1
        while theme:
            try:
                zip.write("ppt\\theme\\theme{}.xml".format(boucle))
                boucle = boucle +1
            except FileNotFoundError():
                theme = False
        zip.write("ppt\\presentation.xml")
        zip.write("ppt\\presProps.xml")
        zip.write("ppt\\tableStyles.xml")
        zip.write("ppt\\viewProps.xml")
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
        
elif ext == ("pptx"):
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
        zip.write("docProps\\app.xml")
        zip.write("docProps\\core.xml")
        zip.write("ppt\\")
        zip.write("ppt\\_rels")
        zip.write("ppt\\_rels\\presentation.xml.rels")
        slidelayout = True
        boucle = 1
            
        while slidelayout:
            try:
                zip.write("ppt\\slideLayouts\\slideLayout{}.xml".format(boucle))
                boucle = boucle+1
            except FileNotFoundError():
                slidelayout = False
        zip.write("[Content_Types].xml")
        zip.write("\\ppt\\slideLayouts\\_rels")
        
        
        slideMaster = True
        boucle = 1
        
        while slideMaster:
            try:
                zip.write("ppt\\slideMasters\\slideMaster{}.xml".format(boucle))
                boucle = boucle +1
            except FileNotFoundError():
                slideMaster = False

        slides = True
        boucle = 1
        
        while slides:
            try:
                zip.write("ppt\\slides\\slide{}.xml".format(boucle))
                boucle = boucle + 1
            except FileNotFoundError():
                slides = False
        
        theme = True
        boucle = 1
        while theme:
            try:
                zip.write("ppt\\theme\\theme{}.xml".format(boucle))
                boucle = boucle +1
            except FileNotFoundError():
                theme = False
        zip.write("ppt\\presentation.xml")
        zip.write("ppt\\presProps.xml")
        zip.write("ppt\\tableStyles.xml")
        zip.write("ppt\\viewProps.xml")
    
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

def checkcode(path):
    logs = open("logs.txt", "w")
    with file(path, "r") as f:
        for lines in f:
            if "subprocess.run" in lines:
                logging.warn("Your file is using a unsecure methode. SUID can be used. Please use an another one.")
            elif "input" in lines:
                end_line = lines.find(")\n")
                start = linex.find("=")
                remove = lines[start: end_line]
                var = lines.replace(remove, "")

                

                
                

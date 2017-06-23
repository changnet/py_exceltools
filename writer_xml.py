#! /bin/python

import os
import xlrd

from optparse import OptionParser

from xml.dom.minidom import getDOMImplementation
from xml.dom.minidom import Element

impl = getDOMImplementation()
OutputUnitTypeDefines = False

def isInVector(i, v):
    try: v.index(i)
    except: return -1
    else: return v.index(i)


class XLSDoc:

    def __init__(self, input_path, output_path):
        self.sheets = []
        self.sheetPageName = []
        self.inputPath = input_path
        self.outputPath = output_path

    def readDoc(self, fileName):
        doc = xlrd.open_workbook(self.inputPath + fileName)
        for i in range(doc.nsheets):
            #sheet = Sheet(fileName.split('.')[0])
            sheet = Sheet("emprise_villes")
            sheetToAdd = doc.sheet_by_index(i)
            nameToAdd = sheetToAdd.name.lower()
            nameToAdd = nameToAdd.split('_')[-1] 
            if(isInVector(nameToAdd, self.sheetPageName) < 0):
                self.sheetPageName.append(nameToAdd)
            sheet.readSheet(sheetToAdd)
            self.sheets.append(sheet)

    def outPutToXML(self):
        for sheet in self.sheets:
            doc = sheet.toXml()
            xmlFile = open(self.outputPath + sheet.fileName, 'w')
            #xmlFile.write(doc.toxml('utf-8'))
            doc.writexml(xmlFile, "    ", "    ", "\n", "utf-8")

        if OutputUnitTypeDefines:                        
            i = 0                       
            enumFile = open(self.outputPath + "../../../src/game/UnitTypeDefines.h", 'w')
            enumFile.write('//-------Do NOT Change!  Auto generated---------\n')
            enumFile.write('\n#ifndef __UNITTYPEDEFINES_H__ \n')
            enumFile.write('#define __UNITTYPEDEFINES_H__ \n')
            enumFile.write('\n')                 
            enumFile.write('enum UNIT_TYPE_ID'+'\n')
            enumFile.write('{'+'\n')                
            i = 0         
            for nameItem in self.sheetPageName:
                enumFile.write('\t' + 'UNIT_TYPE_'+ nameItem.upper()+' = ' + str(i) +',\n')  
                i = i + 1
            enumFile.write('\t' + 'UNIT_TYPE_COUNT'+ ' = ' + str(i) +'\n')
            enumFile.write('};'+'\n\n')

            enumFile.write('static const char* UNIT_TYPE_NAME[UNIT_TYPE_COUNT] = '+'\n')   
            enumFile.write('{'+'\n')    
            for nameItem in self.sheetPageName:
                enumFile.write('\t'+ '\"' + nameItem +'\",' +'\n')  
                i = i + 1
            enumFile.write('};'+'\n')
            enumFile.write('\n#endif'+'\n')
            enumFile.close()
                   



class Sheet:

    def __init__(self, rootName):
        self.rootName = rootName
        self.sheetName = ""
        self.nodeNames = []
        self.contents = []
        self.fileName = ""

    def readSheet(self, sheetObj):
        self.sheetName = sheetObj.name
        self.fileName = self.sheetName + ".xml"
        self.nodeNames = sheetObj.row(0)
        for i in range(1, sheetObj.nrows):
            self.contents.append(sheetObj.row(i))

    def toXml(self):
        doc = impl.createDocument(None, None, None)
        root = doc.createElement(self.rootName)
        for content in self.contents:
            row = doc.createElement(self.sheetName)
            for i in range(len(self.nodeNames)):
                cell = doc.createElement(unicode(self.nodeNames[i].value))
                value = content[i].value
                cell.appendChild(
                        doc.createTextNode(self._convertValue(content[i].value)))
                row.appendChild(cell)
            root.appendChild(row)
        doc.appendChild(root)
        return doc

    def _convertValue(self, value):
        if(type(value) == float and int(value) == value):
            return str(int(value))
        elif(type(value) != str and type(value) != unicode):
            return str(value)
        else:
            return value



if __name__ == '__main__':

    parser = OptionParser()

    parser.add_option("-i", "--input", dest="input_path",
                     default="../doc/",
                     help="read all .xls file from this path")
    parser.add_option("-o", "--output", dest="output_path",
                     default="../data/",
                     help="write all xml file to this path")

    options, args = parser.parse_args()

    fileNames = os.listdir(options.input_path)
	
    XLS = XLSDoc(options.input_path, options.output_path)

    for fileName in fileNames:
        if(fileName == 'kingdorms_data.xls'):
            XLS.readDoc(fileName)

    XLS.outPutToXML()

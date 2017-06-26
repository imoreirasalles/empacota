from xml.etree import ElementTree as ET
from openpyxl import Workbook

#comentário adulterado novamente outra vez

with open("Teste_Literatura_1050.xml", "rt") as f:
    tree = ET.parse(f)
root = tree.getroot()

ns = { 'cumulus' : 'http://www.canto.com/ns/Export/1.0'}

wb = Workbook()
ws = wb.active
ws.title = 'Cumulus'

def mk_header():
   """preenche primeira linha da planilha
      com os campos exportados, retornando
      dicionário com uid : nome do campo"""
   row = 1
   col = 'A'
   uidcampo = {}
   for node in root[0][0].findall('cumulus:Field', ns):
      print(node.attrib, node.text)
      campo = node.find('cumulus:Name', ns)
      ws[ col + str(row)] = campo.text
      col = chr(ord(col) + 1)
      uidcampo[node.attrib['uid']] = campo.text
   return uidcampo

def fill_record(dic):
   row = 2
   for ITEM in root[1]:
      col = 'A'
      for FieldValue in ITEM.findall('cumulus:FieldValue', ns):
         if FieldValue.attrib['uid'] in dic.keys():
            chave = FieldValue.attrib['uid']
            for campo in FieldValue.iter():
                col = chr( 97 + list(dic.keys()).index(chave) )
                ws[ col + str(row)] = campo.text
      row += 1

dic = mk_header()
fill_record(dic)

wb.save('teste_1050.xlsx')

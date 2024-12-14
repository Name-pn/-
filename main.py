import docx
import pygments
import os
import io
import time

from docx.shared import Pt, Cm
from pygments.lexers import CppLexer
from pygments.formatters import RtfFormatter
from spire.doc import *
from spire.doc.common import *

USE_ADD_FILE = 1
NAME_RES = "listings.docx"

def readFolder():
    print("Введите путь к папке с кодом для генерации Doc")
    folderName = input()
    return folderName

def readFile(file):
    f = open(file, "r", encoding='utf-8')
    str = f.read()
    lexer = CppLexer()
    formatter = RtfFormatter()
    highlighted_code = pygments.highlight(str, lexer, formatter)
    f.close()
    return highlighted_code

def insertOne(file, i, listCode, listHead):
    code = readFile(file)
    file = file[file.rfind("\\")+1:]
    listCode.append(code)
    listHead.append(f"Листинг №{i}: \"Файл проекта {file}\"")

def formDoc(codes, headers):
    doc = Document()
    s = doc.AddSection()
    for i in range(0, len(codes)):
        doc.LastSection.BreakCode = SectionBreakType.NoBreak
        p = doc.LastSection.AddParagraph()
        p.AppendText(headers[i])

        stream = Stream(codes[i].encode())
        doc.InsertTextFromStream(stream, FileFormat.Rtf)

    doc.SaveToFile(NAME_RES)

def changeFormatDoc():
    doc = docx.Document(NAME_RES)
    ps = doc.paragraphs

    # Предположим, что мы хотим удалить первый параграф
    p = doc.paragraphs[0]  # Получаем первый параграф

    # Удаляем параграф
    paragraph_element = p._element
    paragraph_element.getparent().remove(paragraph_element)
    for p in ps[1:]:
        for one_run in p.runs:
            if not p.text.find("Листинг"):
                p.paragraph_format.space_after = Pt(8)
                p.paragraph_format.space_before = Pt(8)
                font = one_run.font
                font.size = Pt(14)
            else:
                p.paragraph_format.left_indent = Cm(1.5)
                font = one_run.font
                font.name = 'Consolas'
                font.size = Pt(12)
    doc.save(NAME_RES)

def createDoc(name):
    files = [each for each in os.listdir(name) if each.endswith('.cpp') or each.endswith('.h')]
    i = 1
    listCode, listHead = [], []
    if (name[-1] != '\\'):
        name += '\\'
    for file in files:
        insertOne(name + file, i, listCode, listHead)
        i += 1
    formDoc(listCode, listHead)


if __name__ == '__main__':
    name = readFolder()

    start_time = time.time()
    createDoc(name)
    changeFormatDoc()
    end_time = time.time()

    print(f"Перенос листингов в Docx завершен, он занял {end_time - start_time} секунд")

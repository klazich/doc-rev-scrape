import json
import datetime
from os import walk
from os.path import normpath, splitext, join

import docx2txt
import PyPDF2
import win32com.client as win32


TARGET_DIR = normpath('H:/Document Registration')
EXT_TARGETS = ['.docx', '.doc', '.xlsx', '.xls']


class Com:

    def __init__(self):
        self.Excel = win32.gencache.EnsureDispatch('Excel.Application')
        self.Word = win32.gencache.EnsureDispatch('Word.Application')

    def check(self):
        try:
            self.Excel.Name
        except:
            self.Excel = win32.gencache.EnsureDispatch('Excel.Application')

        try:
            self.Word.Name
        except:
            self.Word = win32.gencache.EnsureDispatch('Word.Application')

    def done(self):
        self.Excel.Quit()
        self.Word.Quit()


def do_pdf(pathname):

    pdfFileObj = open(pathname, 'rb')
    pdfReader = PyPDF2.PdfFileReader(pdfFileObj)

    first_page = pdfReader.getPage(0)
    text = first_page.extractText()

    return text


def get_first_2_rows(workbook):

    cells = workbook.Sheets(1).Range("A1:Z1").Value[0]  \
          + workbook.Sheets(1).Range("A2:Z2").Value[0]
    cells = [str(x) for x in cells if x is not None]
    text = ' '.join(cells)

    return text


def do_excel(filepath, com):

    com.check()

    try:
        workbook = com.Excel.Workbooks.Open(filepath)
    except:
        return 1

    workbook.Visible = False

    if com.Excel.Workbooks.Count > 1:
        for wb in com.Excel.Workbooks:
            if wb.Name != workbook.Name:
                wb.Close(SaveChanges=False)

    text = ''
    try:
        text += workbook.Sheets(1).PageSetup.LeftHeader + ' '
    except:
        return 1
    else:
        text += workbook.Sheets(1).PageSetup.CenterHeader + ' '
        text += workbook.Sheets(1).PageSetup.RightHeader + ' '
        text += get_first_2_rows(workbook)

    return text


def do_doc(filepath, com):

    com.check()

    try:
        document = com.Word.Documents.Open(filepath)
    except:
        return 1

    document.Visible = False

    if com.Word.Documents.Count > 1:
        for doc in com.Word.Documents:
            if doc.Name != document.Name:
                doc.Close(SaveChanges=False)

    text = ''
    try:
        text += document.Sections(1).Headers(1).Range.Text + ' '
    except:
        return 1
    else:
        text += document.Sections(1).Headers(2).Range.Text + ' '
        text += document.Sections(1).Headers(3).Range.Text

    return text


def do_docx(filepath):
    try:
        text = docx2txt.process(filepath)
    except:
        return 1
    else:
        return text


def gather_filenames(dir=TARGET_DIR, skip=[]):

    files = {}
    file_count = 0

    for (dirpath, dirnames, filenames) in walk(dir):

        for filename in filenames:

            ext = splitext(filename)[1]
            filepath = join(dirpath, filename)

            if ext in EXT_TARGETS:
                files[filename] = {'path': filepath, 'rev': None}
                file_count += 1

    to_json = {
        'data': {
            'count': file_count,
            'timestamp': datetime.datetime.now().isoformat()
        },
        'files': files
    }

    return json.dumps(to_json, indent=4)

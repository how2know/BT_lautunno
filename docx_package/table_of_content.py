from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import win32com.client
import inspect, os


class TableOfContent:
    def __init__(self, report_document):
        self.report = report_document
        # self.report_name = report_file_name

    def write(self):
        paragraph = self.report.add_paragraph()
        run = paragraph.add_run()
        fldChar = OxmlElement('w:fldChar')  # creates a new element
        fldChar.set(qn('w:fldCharType'), 'begin')  # sets attribute on element
        instrText = OxmlElement('w:instrText')
        instrText.set(qn('xml:space'), 'preserve')  # sets attribute on element
        instrText.text = 'TOC \\o "1-3" \\h \\z \\u'  # change 1-3 depending on heading levels you need

        fldChar2 = OxmlElement('w:fldChar')
        fldChar2.set(qn('w:fldCharType'), 'separate')
        fldChar3 = OxmlElement('w:t')
        fldChar3.text = 'Press "Ctrl + A" to select everything and then "F9" to update fields.'
        fldChar2.append(fldChar3)

        fldChar4 = OxmlElement('w:fldChar')
        fldChar4.set(qn('w:fldCharType'), 'end')

        r_element = run._r
        r_element.append(fldChar)
        r_element.append(instrText)
        r_element.append(fldChar2)
        r_element.append(fldChar4)
        p_element = paragraph._p

    def update(self, report_file_path):
        word = win32com.client.DispatchEx("Word.Application")
        doc = word.Documents.Open(report_file_path)
        doc.TablesOfContents(1).Update()
        doc.Close(SaveChanges=True)
        word.Quit()

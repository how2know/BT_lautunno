from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.enum.section import WD_SECTION, WD_ORIENT
from docx.shared import Pt, Cm, RGBColor
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import nsdecls, qn
from docx.oxml import parse_xml
from docx.oxml.shared import OxmlElement
from datetime import date
import docx.oxml.ns as ns
from docx.oxml.ns import qn
from docx.oxml import parse_xml
from docx.oxml.shared import OxmlElement


# TODO: write this in a module
def capitalize_first_letter(string):
    """
    Capitalizes the first letter of a string to make it look like a title.
    """
    return string[:1].upper() + string[1:]

def create_element(name):
    return OxmlElement(name)

def create_attribute(element, name, value):
    element.set(ns.qn(name), value)


class Header:
    def __init__(self, section,  list_of_parameters):
        self.section = section
        self.header = self.section.header

        # header of both sections are not linked
        self.header.is_linked_to_previous = False

        self.first_line = self.header.paragraphs[0]
        self.second_line = self.header.add_paragraph()

        self.first_line.paragraph_format.tab_stops.clear_all()
        self.add_tab_stops(self.first_line)
        self.add_tab_stops(self.second_line)

        self.parameters = list_of_parameters

    # add three tab stops (left, center, right)
    def add_tab_stops(self, paragraph):
        paragraph.paragraph_format.tab_stops.add_tab_stop(Cm(0), WD_TAB_ALIGNMENT.LEFT, WD_TAB_LEADER.SPACES)
        paragraph.paragraph_format.tab_stops.add_tab_stop(Cm(8), WD_TAB_ALIGNMENT.CENTER, WD_TAB_LEADER.SPACES)
        paragraph.paragraph_format.tab_stops.add_tab_stop(Cm(16), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.SPACES)

    def write(self):
        firm = capitalize_first_letter(self.parameters['Firm name'])
        title = capitalize_first_letter(self.parameters['Header title'])
        version = capitalize_first_letter(self.parameters['Version / ID'])

        # today's date
        today_date = date.today()
        date_string = today_date.strftime('%d.%m.%Y')

        self.first_line.text = '{} \t {} \t {}'.format(firm, title, version)
        self.second_line.text = ' \t \t {}'.format(date_string)


class Footer:
    def __init__(self, section):
        self.section = section
        self.footer = self.section.footer

        # footer of both sections are not linked
        self.footer.is_linked_to_previous = False

        self.first_line = self.footer.paragraphs[0]
        self.second_line = self.footer.add_paragraph()

        self.first_line.paragraph_format.tab_stops.clear_all()
        self.add_tab_stops(self.first_line)
        self.add_tab_stops(self.second_line)

    # add three tab stops (left, center, right)
    def add_tab_stops(self, paragraph):
        paragraph.paragraph_format.tab_stops.add_tab_stop(Cm(0), WD_TAB_ALIGNMENT.LEFT, WD_TAB_LEADER.SPACES)
        paragraph.paragraph_format.tab_stops.add_tab_stop(Cm(8), WD_TAB_ALIGNMENT.CENTER, WD_TAB_LEADER.SPACES)
        paragraph.paragraph_format.tab_stops.add_tab_stop(Cm(16), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.SPACES)

    def add_page_number(self, run):
        fldChar1 = create_element('w:fldChar')
        create_attribute(fldChar1, 'w:fldCharType', 'begin')

        instrText = create_element('w:instrText')
        create_attribute(instrText, 'xml:space', 'preserve')
        instrText.text = "PAGE"

        fldChar2 = create_element('w:fldChar')
        create_attribute(fldChar2, 'w:fldCharType', 'end')

        run._r.append(fldChar1)
        run._r.append(instrText)
        run._r.append(fldChar2)

    def write(self):

        self.second_line.text = ' \t \t'

        self.add_page_number(self.second_line.add_run())

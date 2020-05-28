from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
import os
from zipfile import ZipFile
from bs4 import BeautifulSoup

# from Writing_text import layout
# from Reading_text import text_reading

from docx_package import text_reading, layout

class Chapter:
    '''A class that defines a chapter ...'''

    def __init__(self, report_document, text_input_document, text_input_soup, title, list_of_tables, parameters_dictionary):
        self.report = report_document
        self.text_input = text_input_document
        self.text_input_soup = text_input_soup
        self.title = title
        self.list_of_tables = list_of_tables
        self.parameters_dictionary = parameters_dictionary

    def heading_index(self):
        """
        Find the heading of the chapter and return the corresponding paragraph index.
        """

        for paragraph_index, paragraph in enumerate(self.text_input.paragraphs):
            if paragraph.text == self.title and 'Heading' in paragraph.style.name:
                return paragraph_index

    def next_heading_index(self):
        """
        Return the index of the following heading.
        """

        previous_index = self.heading_index()

        for paragraph_index, paragraph in enumerate(self.text_input.paragraphs[previous_index + 1:]):
            if 'Heading' in paragraph.style.name:
                return paragraph_index + previous_index + 1

    @ property
    def paragraphs(self):
        """
        Return a list of all paragraphs (as string) of the chapter.
        """

        list_of_paragraphs = []
        heading_index = self.heading_index()
        next_heading_index = self.next_heading_index()

        for paragraph in self.text_input.paragraphs[heading_index + 1: next_heading_index]:
            list_of_paragraphs.append(paragraph.text)

        return list_of_paragraphs

    @ property
    def parameters(self):
        """
        Read the dropdown lists of the parameter table and return their value in a list.
        """

        return text_reading.get_dropdown_list_of_table(self.text_input_soup,
                                                       self.list_of_tables.index('{} parameter table'.format(self.title))
                                                       )

    def write_chapter(self):
        """
        Write the heading and the paragraphs of a chapter, including the parameters.
        """

        # read heading style and write heading
        heading_style = self.text_input.paragraphs[self.heading_index()].style.name
        self.report.add_paragraph(self.title, heading_style)

        parameters_values = ['', '', '']

        '''Create variables in order to call property only once, and not in a loop.'''
        parameters = self.parameters
        paragraphs = self.paragraphs

        # stores values of corresponding parameter keys in a list
        for parameter_index, parameter in enumerate(parameters):
            if parameter != '-':
                parameters_values[parameter_index] = self.parameters_dictionary[parameter]

        # write paragraphs including values of parameters
        for paragraph in paragraphs:
            new_paragraph = self.report.add_paragraph(
                paragraph.format(parameters_values[0], parameters_values[1], parameters_values[2],)
            )
            new_paragraph.style.name = 'Normal'
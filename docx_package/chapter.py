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

from docx_package import text_writing, text, text_reading, layout

class Chapter:
    '''A class that defines a chapter ...'''

    TEXT_INPUT_FILE = 'Text_input.docx'
    REPORT_FILE = 'Report.docx'
    DEFINITIONS_FILE = 'Terms_definitions.docx'
    INPUTS_DIRECTORY = 'Inputs'

    def __init__(self, report_document, text_input_path, title, list_of_tables, parameters_dictionary):
        self.report = report_document
        self.text_input_path = text_input_path
        self.text_input = Document(text_input_path)
        self.title = title
        self.list_of_tables = list_of_tables
        self.parameters_dictionary = parameters_dictionary
        # self.heading_level = heading_level
        # self.parameter_table_index = parameter_table_index
        # self.picture_table = picture_table

    '''
    @ property
    def text_input(self):
        text_input = File(self.TEXT_INPUT_FILE, self.INPUTS_DIRECTORY)
        return Document(text_input.path)
    '''

    #  find a heading given his title and return the corresponding paragraph index
    def heading_name_index(self):
        for i in range(len(self.text_input.paragraphs)):
            if self.text_input.paragraphs[i].style.name == 'Heading 1':
                if self.text_input.paragraphs[i].text == self.title:
                    return i
            elif self.text_input.paragraphs[i].style.name == 'Heading 2':
                if self.text_input.paragraphs[i].text == self.title:
                    return i
            elif self.text_input.paragraphs[i].style.name == 'Heading 3':
                if self.text_input.paragraphs[i].text == self.title:
                    return i
            elif self.text_input.paragraphs[i].style.name == 'Heading 4':
                if self.text_input.paragraphs[i].text == self.title:
                    return i

            '''
            if self.text_input.paragraphs[i].style.name == 'Heading {}'.format(1 or 2 or 3 or 4):  # look for paragraphs with corresponding style
                if self.text_input.paragraphs[i].text == self.title:  # look for paragraphs with corresponding title
                    return i  # return the index of the paragraphs
            '''

    # return the index of the next heading corresponding to a style given the index of the previous heading
    def next_heading_index(self, previous_index):
        for i in range(previous_index + 1, len(self.text_input.paragraphs)):
            if self.text_input.paragraphs[i].style.name == 'Heading 1':
                return i
            elif self.text_input.paragraphs[i].style.name == 'Heading 2':
                return i
            elif self.text_input.paragraphs[i].style.name == 'Heading 3':
                return i
            elif self.text_input.paragraphs[i].style.name == 'Heading 4':
                return i

            '''
            if self.text_input.paragraphs[i].style.name == 'Heading {}'.format(1 or 2 or 3 or 4):  # look for paragraphs with corresponding style
                return i  # return the index of the paragraph
            '''

    # store all paragraphs and their corresponding style between a given heading and the next one
    def paragraph_after_heading_with_styles(self, list_of_paragraphs, list_of_styles):
        heading_index = self.heading_name_index()  # index of the given heading
        next_index = self.next_heading_index(heading_index)  # index of the next heading
        for i in range(heading_index + 1, next_index):  # loop over all paragraphs between the given heading and the next one
            list_of_paragraphs.append(self.text_input.paragraphs[i])  # store all paragraphs in a given list
            list_of_styles.append(self.text_input.paragraphs[i].style.name)  # store all styles in a given list

    # return the paragraphs following the given heading
    @ property
    def paragraphs(self):
        list_of_paragraphs = []
        heading_index = self.heading_name_index()  # index of the given heading
        next_index = self.next_heading_index(heading_index)  # index of the next heading
        for i in range(heading_index + 1, next_index):  # loop over all paragraphs between the given heading and the next one
            list_of_paragraphs.append(self.text_input.paragraphs[i])  # store all paragraphs in a given list

        return list_of_paragraphs

    @ property
    def parameters(self):
        """
        Read the dropdown lists of the parameter table and return their value in a list.
        """

        return text_reading.get_dropdown_list_of_table(self.text_input_path,
                                                       self.list_of_tables.index('{} parameter table'.format(self.title))
                                                       )

    def write_chapter(self):
        """
        Write the heading and the paragraphs of a chapter, including the parameters.
        """

        # read heading style and write heading
        heading_style = self.text_input.paragraphs[self.heading_name_index()].style.name
        self.report.add_paragraph(self.title, heading_style)

        parameters_values = ['', '', '']

        '''Create variables in order to call property only once, and not in a loop.'''
        parameters = self.parameters
        paragraphs = self.paragraphs

        # stores values of corresponding parameter keys in a list as lower case string
        for i in range(len(parameters)):
            if parameters[i] != '-':
                parameters_values[i] = self.parameters_dictionary[parameters[i]].lower()

        # write paragraphs including values of parameters
        for i in range(len(paragraphs)):
            paragraph = self.report.add_paragraph(
                paragraphs[i].text.format(parameters_values[0], parameters_values[1], parameters_values[2])
            )
            paragraph.style.name = 'Normal'
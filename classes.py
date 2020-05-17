from docx import Document
import os
from zipfile import ZipFile
from bs4 import BeautifulSoup

class File:
    ''' Class that defines a file '''

    def __init__(self, file_name, directory_name):
        self.name = file_name
        self.directory = directory_name

    # return the path of a file given its name and the name of its directory
    @ property
    def path(self):
        text_input_directory = os.path.dirname(self.name)
        path = os.path.join(text_input_directory, self.directory, self.name)

        return path


class Chapter:
    '''A class that defines a chapter ...'''

    TEXT_INPUT_FILE = 'Text_input.docx'
    REPORT_FILE = 'Report.docx'
    DEFINITIONS_FILE = 'Terms_definitions.docx'
    INPUTS_DIRECTORY = 'Inputs'

    def __init__(self, report_document, text_input_path, title, heading_level, parameter_table_index):
        self.report = report_document
        self.text_input_path = text_input_path
        self.text_input = Document(text_input_path)
        self.title = title
        self.heading_level = heading_level
        self.parameter_table_index = parameter_table_index
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

        list_of_value = []

        # open docx file as a zip file and store its relevant xml data
        zip_file = ZipFile(self.text_input_path)
        xml_data = zip_file.read('word/document.xml')
        zip_file.close()

        # parse the xml data with BeautifulSoup
        soup = BeautifulSoup(xml_data, 'xml')

        # look for all values of dropdown lists in the data and store them
        tables = soup.find_all('tbl')
        dd_lists_content = tables[self.parameter_table_index].find_all('sdtContent')
        for i in dd_lists_content:
            list_of_value.append(i.find('t').string)

        return list_of_value

        '''
        parameters = []
        for i in range(1, 4):
            parameters.append(self.parameter_table.cell(i, 2).text)
            
        return parameters
        '''

    def write_chapter(self):
        self.report.add_heading(self.title, self.heading_level)
        for i in range(len(self.paragraphs)):
            paragraph = self.report.add_paragraph(
                self.paragraphs[i].text.format(self.parameters[0], self.parameters[1], self.parameters[2])
            )
            paragraph.style.name = 'Normal'

    # read dropdown lists and store their value in a list
    def dropdown_lists_value(self, file_name):

        list_of_value = []

        # open docx file as a zip file and store its relevant xml data
        zip_file = ZipFile(file_name)
        xml_data = zip_file.read('word/document.xml')
        zip_file.close()

        # parse the xml data with BeautifulSoup
        soup = BeautifulSoup(xml_data, 'xml')

        # look for all values of dropdown lists in the data and store them
        tables = soup.find_all('tbl')
        dd_lists_content = tables[self.parameter_table_index].find_all('sdtContent')
        for i in dd_lists_content:
            list_of_value.append(i.find('t').string)




class Parameters:
    def __init__(self, text_input_document):
        self.text_input = text_input_document

    def read_standard_tables(self, parameters_dictionary, table_index):
        for row in self.text_input.tables[table_index].rows:
            key = row.cells[0].text
            parameters_dictionary[key] = row.cells[1].text

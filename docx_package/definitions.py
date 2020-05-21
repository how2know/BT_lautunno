from docx import Document
import time

from docx_package import text_writing, text, text_reading, layout


class Definitions:
    STANDARDS_NAMES = ['EU Regulation 2017/745', 'IEC 62366-1', 'FDA Guidance']

    def __init__(self, report_document, text_input_document, text_input_soup, definitions_document, title, list_of_tables):
        self.report = report_document
        self.text_input = text_input_document
        self.text_input_soup = text_input_soup
        self.definitions = definitions_document
        self.title = title
        self.list_of_tables = list_of_tables

    def standard_heading_index(self, standard_name):
        for paragraph_index, paragraph in enumerate(self.definitions.paragraphs):
            if paragraph.text == standard_name and 'Heading' in paragraph.style.name:
                return paragraph_index

    def next_standard_heading_index(self, previous_index):
        for paragraph_index, paragraph in enumerate(self.definitions.paragraphs[previous_index + 1:]):
            if paragraph.text in self.STANDARDS_NAMES and 'Heading' in paragraph.style.name:
                return paragraph_index + previous_index + 1

    def standard_wanted_terms(self, standard_name):
        for table_index, table in enumerate(self.list_of_tables):
            if standard_name in table:

                list_of_yes_no = text_reading.get_dropdown_list_of_table(self.text_input_soup, table_index)

                list_of_terms = []
                for row in self.text_input.tables[table_index].rows:
                    for cell in row.cells:
                        if cell.text:
                            list_of_terms.append(cell.text)

                list_of_defined_terms = []
                for i in range(len(list_of_terms)):
                    if list_of_yes_no[i] == 'Yes':
                        list_of_defined_terms.append(list_of_terms[i])

                return list_of_defined_terms

    def write_terms_definitions(self, standard_name, list_of_terms):
        standard_heading_index = self.standard_heading_index(standard_name)
        next_index = self.next_standard_heading_index(standard_heading_index)

        terms_heading_indexes = []

        for paragraph_index, paragraph in enumerate(self.definitions.paragraphs[standard_heading_index: next_index]):
            if 'Heading' in paragraph.style.name:
                terms_heading_indexes.append(paragraph_index + standard_heading_index)

        list_of_paragraphs = []
        list_of_styles = []

        for index, terms_heading_index in enumerate(terms_heading_indexes):
            if self.definitions.paragraphs[terms_heading_index].text in list_of_terms:
                for i in range(terms_heading_index, terms_heading_indexes[index + 1]):
                    list_of_paragraphs.append(self.definitions.paragraphs[i].text)
                    list_of_styles.append(self.definitions.paragraphs[i].style.name)

        for index, paragraph in enumerate(list_of_paragraphs):
            self.report.add_paragraph(paragraph, list_of_styles[index])

    def write_all_definitions(self):
        self.report.add_heading(self.title)
        for standard_name in self.STANDARDS_NAMES:
            first_index = self.standard_heading_index(standard_name)

            wanted_terms = self.standard_wanted_terms(standard_name)
            self.write_terms_definitions(standard_name, wanted_terms)
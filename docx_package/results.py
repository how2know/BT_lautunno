from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
import os
from zipfile import ZipFile
from bs4 import BeautifulSoup

from docx_package import layout, text_reading, text_writing, text

# from Writing_text import layout
# from Reading_text import text_reading

class Results:
    def __init__(self, report_document, text_input_path, title, list_of_tables, parameters_dictionary):
        self.report = report_document
        self.text_input_path = text_input_path
        self.text_input = Document(text_input_path)
        self.title = title
        self.list_of_tables = list_of_tables
        self.parameters_dictionary = parameters_dictionary

    def visualization(self):
        # index of the input tables
        task_table_index = self.list_of_tables.index('Effectiveness analysis tasks and problems table')
        problem_table_index = self.list_of_tables.index('Effectiveness analysis problem type table')
        video_table_index = self.list_of_tables.index('Effectiveness analysis video table')

        # input tables
        task_table = self.text_input.tables[task_table_index]
        problem_table = self.text_input.tables[problem_table_index]
        video_table = self.text_input.tables[video_table_index]

        # color
        light_grey_10 = 'D0CECE'
        green = '00B050'
        red = 'FF0000'
        orange = 'FFC000'
        yellow = 'FFFF00'

        # create a table for the results visualization
        result_table_rows = self.parameters_dictionary['Number of critical tasks'] + 1
        result_table_cols = self.parameters_dictionary['Number of participants'] + 1
        result_table = self.report.add_table(result_table_rows, result_table_cols)
        layout.define_table_style(result_table)

        # write the information of the input table in the result table
        for i in range(result_table_rows):
            for j in range(result_table_cols):
                cell = result_table.cell(i, j)

                # skip the first row and first column
                if i != 0 and j != 0:
                    cell.text = task_table.cell(i, j).text
                    cell.paragraphs[0].runs[0].font.bold = True

                # first row
                elif i == 0 and j != 0:
                    cell.text = task_table.cell(i, j).text
                    layout.set_cell_shading(cell, light_grey_10)     # color the cell in light_grey_10
                    cell.paragraphs[0].runs[0].font.bold = True

                # first column
                elif i != 0 and j == 0:
                    cell.text = self.parameters_dictionary['Critical task {} name'.format(i)]
                    layout.set_cell_shading(cell, light_grey_10)     # color the cell in light_grey_10
                    cell.paragraphs[0].runs[0].font.bold = True

                # bolds all text
                if cell.text:
                    cell.paragraphs[0].runs[0].font.bold = True

        # color the cell according to the type of problem
        for i in range(1, result_table_rows):
            for j in range(1, result_table_cols):
                cell = result_table.cell(i, j)

                if cell.text:     # check if the text string is not empty
                    index = int(cell.text)

                    if self.parameters_dictionary['Problem {} type'.format(index)] == 'Important problem':
                        layout.set_cell_shading(cell, orange)

                    if self.parameters_dictionary['Problem {} type'.format(index)] == 'Marginal problem':
                        layout.set_cell_shading(cell, yellow)

                    if self.parameters_dictionary['Problem {} type'.format(index)] == 'Critical problem':
                        layout.set_cell_shading(cell, red)
                else:
                    layout.set_cell_shading(cell, green)

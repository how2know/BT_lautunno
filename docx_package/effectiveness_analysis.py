from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT, WD_ROW_HEIGHT_RULE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, Cm, RGBColor
import os
from zipfile import ZipFile
from bs4 import BeautifulSoup

from docx_package import layout, text_reading, text_writing, text
from docx_package.results import ResultsChapter


class EffectivenessAnalysis:

    TITLE = 'Effectiveness analysis'
    TITLE_STYLE = 'Heading 2'

    # color for cell shading
    LIGHT_GREY_10 = 'D0CECE'
    GREEN = '00B050'
    RED = 'FF0000'
    ORANGE = 'FFC000'
    YELLOW = 'FFFF00'

    def __init__(self, report_document, text_input_document, text_input_soup, title, list_of_tables, parameters_dictionary):
        self.report = report_document
        self.text_input = text_input_document
        self.text_input_soup = text_input_soup
        self.title = title
        self.list_of_tables = list_of_tables
        self.parameters_dictionary = parameters_dictionary

    def add_table(self):
        # index of the input tables
        task_table_index = self.list_of_tables.index('Effectiveness analysis tasks and problems table')
        problem_table_index = self.list_of_tables.index('Effectiveness analysis problem type table')
        video_table_index = self.list_of_tables.index('Effectiveness analysis video table')

        # input tables
        task_table = self.text_input.tables[task_table_index]
        problem_table = self.text_input.tables[problem_table_index]
        video_table = self.text_input.tables[video_table_index]

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
                    layout.set_cell_shading(cell, self.LIGHT_GREY_10)     # color the cell in light_grey_10
                    cell.paragraphs[0].runs[0].font.bold = True

                # first column
                elif i != 0 and j == 0:
                    cell.text = self.parameters_dictionary['Critical task {} name'.format(i)]
                    layout.set_cell_shading(cell, self.LIGHT_GREY_10)     # color the cell in light_grey_10
                    cell.paragraphs[0].runs[0].font.bold = True

                # bolds all text
                if cell.text:
                    cell.paragraphs[0].runs[0].font.bold = True

        # color the cell according to the type of problem
        for i in range(1, result_table_rows):
            for j in range(1, result_table_cols):
                cell = result_table.cell(i, j)

                if cell.text:     # check if the text string is not empty
                    problem_index = int(cell.text)

                    if self.parameters_dictionary['Problem {} type'.format(problem_index)] == 'Important problem':
                        layout.set_cell_shading(cell, self.ORANGE)

                    if self.parameters_dictionary['Problem {} type'.format(problem_index)] == 'Marginal problem':
                        layout.set_cell_shading(cell, self.YELLOW)

                    if self.parameters_dictionary['Problem {} type'.format(problem_index)] == 'Critical problem':
                        layout.set_cell_shading(cell, self.RED)
                else:
                    layout.set_cell_shading(cell, self.GREEN)

    # TODO: finish this function
    def add_description_table(self):
        self.report.add_paragraph()

        description_table = self.report.add_table(3, 8)
        description_table.allow_autofit = False

        for row in description_table.rows:
            row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY

        description_table.rows[0].height = Cm(0.5)
        description_table.rows[1].height = Cm(0.1)
        description_table.rows[2].height = Cm(0.5)

        for cell in description_table.columns[0].cells:
            cell.width = Cm(2)

        # description_table.columns[0].cells[0].width = Cm(2)
        description_table.columns[1].cells[0].width = Cm(0.5)
        description_table.columns[2].cells[0].width = Cm(3.8)
        description_table.columns[3].cells[0].width = Cm(1.6)
        description_table.columns[4].cells[0].width = Cm(1.7)
        description_table.columns[5].cells[0].width = Cm(0.5)
        description_table.columns[6].cells[0].width = Cm(3.8)
        description_table.columns[7].cells[0].width = Cm(2)


        layout.set_cell_shading(description_table.cell(0, 1), self.GREEN)
        layout.set_cell_shading(description_table.cell(0, 5), self.ORANGE)
        layout.set_cell_shading(description_table.cell(2, 1), self.YELLOW)
        layout.set_cell_shading(description_table.cell(2, 5), self.RED)

        description_table.cell(0, 2).text = 'No problem found'
        description_table.cell(0, 6).text = 'Important problem found'
        description_table.cell(2, 2).text = 'Marginal problem found'
        description_table.cell(2, 6).text = 'Critical problem found'

        description_table.cell(0, 2).paragraphs[0].runs[0].font.size = Pt(9)
        description_table.cell(0, 6).paragraphs[0].runs[0].font.size = Pt(9)
        description_table.cell(2, 2).paragraphs[0].runs[0].font.size = Pt(9)
        description_table.cell(2, 6).paragraphs[0].runs[0].font.size = Pt(9)

        '''
        for cell in description_table.rows[0].cells:
            cell.paragraphs[0].runs[0].font.size = Pt(9)
        '''

    def add_problem_description(self):
        for i in range(1, self.parameters_dictionary['Number of problems'] + 1):
            self.report.add_paragraph(self.parameters_dictionary['Problem {} description'.format(i)],
                                      'List Number')

    def write_chapter(self):

        effectiveness_analysis = ResultsChapter(self.report, self.text_input, self.text_input_soup, self.title, self.list_of_tables, self.parameters_dictionary)

        self.report.add_paragraph(self.TITLE, self.TITLE_STYLE)
        self.add_table()
        self.add_description_table()

        self.report.add_paragraph('Problems description', 'Heading 3')
        self.add_problem_description()

        self.report.add_paragraph('Analysis', 'Heading 3')
        effectiveness_analysis.write_chapter()


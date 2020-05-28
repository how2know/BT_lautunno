from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT, WD_ROW_HEIGHT_RULE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, Cm, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import os
from zipfile import ZipFile
from bs4 import BeautifulSoup

from docx_package import layout, text_reading, text_writing, text
from docx_package.results import ResultsChapter


# TODO: write this function in a module
def set_column_width(column, size):
    """
    Set the width of the given column to the given size (in cm).
    To make it works, the autofit of the corresponding table must be disabled beforehand (table.autofit = False).
    """
    for cell in column.cells:
        cell.width = Cm(size)


# TODO: write this function in a module
def set_cell_border(cell, **kwargs):
    """
    Set cell`s border
    Usage:

    set_cell_border(
        cell,
        top={"sz": 12, "val": "single", "color": "#FF0000", "space": "0"},
        bottom={"sz": 12, "color": "#00FF00", "val": "single"},
        start={"sz": 24, "val": "dashed", "shadow": "true"},
        end={"sz": 12, "val": "dashed"},
    )
    """
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()

    # check for tag existence, if none found, then create one
    tcBorders = tcPr.first_child_found_in("w:tcBorders")
    if tcBorders is None:
        tcBorders = OxmlElement('w:tcBorders')
        tcPr.append(tcBorders)

    # list over all available tags
    for edge in ('start', 'top', 'end', 'bottom', 'insideH', 'insideV'):
        edge_data = kwargs.get(edge)
        if edge_data:
            tag = 'w:{}'.format(edge)

            # check for tag existence, if none found, then create one
            element = tcBorders.find(qn(tag))
            if element is None:
                element = OxmlElement(tag)
                tcBorders.append(element)

            # looks like order of attributes is important
            for key in ["sz", "val", "color", "space", "shadow"]:
                if key in edge_data:
                    element.set(qn('w:{}'.format(key)), str(edge_data[key]))


# TODO: comment everything
class EffectivenessAnalysis:

    TITLE = 'Effectiveness analysis'
    TITLE_STYLE = 'Heading 2'
    DESCRIPTION_TITLE = 'Problems description'
    DESCRIPTION_STYLE = 'Heading 3'
    ANALYSIS_TITLE = 'Analysis'
    ANALYSIS_STYLE = 'Heading 3'

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

    # TODO: make rows the same height and text appears in the middle
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

        set_cell_border(result_table.cell(0, 0),
                        top={"color": "#FFFFFF"},
                        start={"color": "#FFFFFF"},
                        )

        for row in result_table.rows:
            for cell in row.cells:
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    def add_description_table(self):
        self.report.add_paragraph()

        description_table = self.report.add_table(3, 8)
        # description_table.allow_autofit = False
        description_table.autofit = False

        for row in description_table.rows:
            row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY

        description_table.rows[0].height = Cm(0.5)
        description_table.rows[1].height = Cm(0.2)
        description_table.rows[2].height = Cm(0.5)

        set_column_width(description_table.columns[0], 2)
        set_column_width(description_table.columns[1], 0.5)
        set_column_width(description_table.columns[2], 3.8)
        set_column_width(description_table.columns[3], 1.6)
        set_column_width(description_table.columns[4], 1.7)
        set_column_width(description_table.columns[5], 0.5)
        set_column_width(description_table.columns[6], 3.8)
        set_column_width(description_table.columns[7], 2)

        layout.set_cell_shading(description_table.cell(0, 1), self.GREEN)
        layout.set_cell_shading(description_table.cell(0, 5), self.ORANGE)
        layout.set_cell_shading(description_table.cell(2, 1), self.YELLOW)
        layout.set_cell_shading(description_table.cell(2, 5), self.RED)

        # TODO: maybe make a function for this
        set_cell_border(description_table.cell(0, 1),
                        top={"sz": 4, "val": "single", "color": "#000000"},
                        bottom={"sz": 4, "val": "single", "color": "#000000"},
                        start={"sz": 4, "val": "single", "color": "#000000"},
                        end={"sz": 4, "val": "single", "color": "#000000"},
                        )

        set_cell_border(description_table.cell(0, 5),
                        top={"sz": 4, "val": "single", "color": "#000000",},
                        bottom={"sz": 4, "val": "single", "color": "#000000"},
                        start={"sz": 4, "val": "single", "color": "#000000"},
                        end={"sz": 4, "val": "single", "color": "#000000"},
                        )

        set_cell_border(description_table.cell(2, 1),
                        top={"sz": 4, "val": "single", "color": "#000000",},
                        bottom={"sz": 4, "val": "single", "color": "#000000"},
                        start={"sz": 4, "val": "single", "color": "#000000"},
                        end={"sz": 4, "val": "single", "color": "#000000"},
                        )

        set_cell_border(description_table.cell(2, 5),
                        top={"sz": 4, "val": "single", "color": "#000000", },
                        bottom={"sz": 4, "val": "single", "color": "#000000"},
                        start={"sz": 4, "val": "single", "color": "#000000"},
                        end={"sz": 4, "val": "single", "color": "#000000"},
                        )

        description_table.cell(0, 2).text = 'No problem found'
        description_table.cell(0, 6).text = 'Important problem found'
        description_table.cell(2, 2).text = 'Marginal problem found'
        description_table.cell(2, 6).text = 'Critical problem found'

        description_table.cell(0, 2).paragraphs[0].runs[0].font.size = Pt(9)
        description_table.cell(0, 6).paragraphs[0].runs[0].font.size = Pt(9)
        description_table.cell(2, 2).paragraphs[0].runs[0].font.size = Pt(9)
        description_table.cell(2, 6).paragraphs[0].runs[0].font.size = Pt(9)

        for row in description_table.rows:
            for cell in row.cells:
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    def add_problem_description(self):
        for i in range(1, self.parameters_dictionary['Number of problems'] + 1):
            self.report.add_paragraph(self.parameters_dictionary['Problem {} description'.format(i)],
                                      'List Number')

    def write_chapter(self):

        effectiveness_analysis = ResultsChapter(self.report, self.text_input, self.text_input_soup, self.title,
                                                self.list_of_tables, self.parameters_dictionary)

        self.report.add_paragraph(self.TITLE, self.TITLE_STYLE)
        self.add_table()
        self.add_description_table()

        self.report.add_paragraph(self.DESCRIPTION_TITLE, self.DESCRIPTION_STYLE)
        self.add_problem_description()

        self.report.add_paragraph(self.ANALYSIS_TITLE, self.ANALYSIS_STYLE)
        effectiveness_analysis.write_chapter()


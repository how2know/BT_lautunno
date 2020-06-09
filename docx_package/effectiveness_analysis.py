from docx import Document
from docx.document import Document
from docx.table import Table
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
from typing import List, Dict, Union

from docx_package import layout, text_reading
from docx_package.results import ResultsChapter


# TODO: improve comments
class EffectivenessAnalysis:
    """
    Class that represents the 'Effectiveness analysis' chapter and the visualization of its results.
    """

    # name of tables as they appear in the tables list
    TASK_TABLE_NAME = 'Effectiveness analysis tasks and problems table'
    PROBLEM_TABLE_NAME = 'Effectiveness analysis problem type table'
    VIDEO_TABLE_NAME = 'Effectiveness analysis video table'

    # information about the headings of this chapter
    TITLE = 'Effectiveness analysis'
    TITLE_STYLE = 'Heading 2'
    DESCRIPTION_TITLE = 'Problems description'
    DESCRIPTION_STYLE = 'Heading 3'
    DISCUSSION_TITLE = 'Discussion'
    DISCUSSION_STYLE = 'Heading 3'

    # hexadecimals of colors for cell shading
    LIGHT_GREY_10 = 'D0CECE'
    GREEN = '00B050'
    RED = 'FF0000'
    ORANGE = 'FFC000'
    YELLOW = 'FFFF00'

    def __init__(self, report_document: Document,
                 text_input_document: Document,
                 text_input_soup: BeautifulSoup,
                 list_of_tables: List[str],
                 parameters_dictionary: Dict[str, Union[str, int]]
                 ):
        """
        Args:
            report_document: .docx file where the report is written.
            text_input_document: .docx file where all inputs are written.
            text_input_soup: BeautifulSoup of the xml of the input .docx file.
            list_of_tables: List of all table names.
            parameters_dictionary: Dictionary of all input parameters (key = parameter name, value = parameter value)
        """

        self.report = report_document
        self.text_input = text_input_document
        self.text_input_soup = text_input_soup
        self.tables = list_of_tables
        self.parameters = parameters_dictionary

    @ property
    def task_table(self) -> Table:
        """
        Returns:
            Table of the input .docx file where the information about the effectiveness analysis are written.
        """

        task_table_index = self.tables.index(self.TASK_TABLE_NAME)
        return self.text_input.tables[task_table_index]

    @ property
    def problem_table(self) -> Table:
        """
        Returns:
            Table of the input .docx file where the description of the problems are written.
        """

        problem_table_index = self.tables.index(self.PROBLEM_TABLE_NAME)
        return self.text_input.tables[problem_table_index]

    @ property
    def video_table(self) -> Table:
        """
        Returns:
            Table of the input .docx file where the information about the task videos are written.
        """

        video_table_index = self.tables.index(self.VIDEO_TABLE_NAME)
        return self.text_input.tables[video_table_index]

    # TODO: why is the table too large
    # TODO: make rows the same height and text appears in the middle
    def make_result_table(self):
        """
        Create a table for the visualization of the effectiveness analysis.
        """

        # create a table for the results visualization
        result_table_rows = self.parameters['Number of critical tasks'] + 1
        result_table_cols = self.parameters['Number of participants'] + 1
        result_table = self.report.add_table(result_table_rows, result_table_cols)
        layout.define_table_style(result_table)      # TODO: do we need this line

        # write the information of the input table in the result table
        for i in range(result_table_rows):
            for j in range(result_table_cols):
                cell = result_table.cell(i, j)

                # skip the first row and first column
                if i != 0 and j != 0:
                    cell.text = self.task_table.cell(i, j).text
                    cell.paragraphs[0].runs[0].font.bold = True

                # first row
                elif i == 0 and j != 0:
                    cell.text = self.task_table.cell(i, j).text
                    layout.set_cell_shading(cell, self.LIGHT_GREY_10)     # color the cell in light_grey_10
                    cell.paragraphs[0].runs[0].font.bold = True

                # first column
                elif i != 0 and j == 0:
                    cell.text = self.parameters['Critical task {} name'.format(i)]
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

                    if self.parameters['Problem {} type'.format(problem_index)] == 'Important problem':
                        layout.set_cell_shading(cell, self.ORANGE)

                    if self.parameters['Problem {} type'.format(problem_index)] == 'Marginal problem':
                        layout.set_cell_shading(cell, self.YELLOW)

                    if self.parameters['Problem {} type'.format(problem_index)] == 'Critical problem':
                        layout.set_cell_shading(cell, self.RED)
                else:
                    layout.set_cell_shading(cell, self.GREEN)

        # color the top left cell borders in white
        layout.set_cell_border(result_table.cell(0, 0),
                               top={"color": "#FFFFFF"},
                               start={"color": "#FFFFFF"}
                               )

        # set text vertical alignment of all cells to center
        for row in result_table.rows:
            for cell in row.cells:
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    def make_description_table(self):
        """
        Create a table that describes the color of the effectiveness analysis table.
        """

        # add a line for space between the two tables
        separation_line = self.report.add_paragraph(' ')
        separation_line.runs[0].font.size = Pt(5)

        # create table
        description_table = self.report.add_table(3, 8)
        description_table.autofit = False

        # set height of rows
        layout.set_row_height(description_table.rows[0], 0.5)
        layout.set_row_height(description_table.rows[1], 0.2)
        layout.set_row_height(description_table.rows[2], 0.5)

        # set width of columns
        layout.set_column_width(description_table.columns[0], 2)
        layout.set_column_width(description_table.columns[1], 0.5)
        layout.set_column_width(description_table.columns[2], 3.8)
        layout.set_column_width(description_table.columns[3], 1.6)
        layout.set_column_width(description_table.columns[4], 1.7)
        layout.set_column_width(description_table.columns[5], 0.5)
        layout.set_column_width(description_table.columns[6], 3.8)
        layout.set_column_width(description_table.columns[7], 2)

        # TODO: maybe make a function for this
        # color four cells and set their borders
        layout.set_cell_shading(description_table.cell(0, 1), self.GREEN)
        layout.set_cell_shading(description_table.cell(0, 5), self.ORANGE)
        layout.set_cell_shading(description_table.cell(2, 1), self.YELLOW)
        layout.set_cell_shading(description_table.cell(2, 5), self.RED)
        layout.set_cell_border(description_table.cell(0, 1),
                               top={"sz": 4, "val": "single", "color": "#000000"},
                               bottom={"sz": 4, "val": "single", "color": "#000000"},
                               start={"sz": 4, "val": "single", "color": "#000000"},
                               end={"sz": 4, "val": "single", "color": "#000000"},
                               )
        layout.set_cell_border(description_table.cell(0, 5),
                               top={"sz": 4, "val": "single", "color": "#000000"},
                               bottom={"sz": 4, "val": "single", "color": "#000000"},
                               start={"sz": 4, "val": "single", "color": "#000000"},
                               end={"sz": 4, "val": "single", "color": "#000000"},
                               )
        layout.set_cell_border(description_table.cell(2, 1),
                               top={"sz": 4, "val": "single", "color": "#000000"},
                               bottom={"sz": 4, "val": "single", "color": "#000000"},
                               start={"sz": 4, "val": "single", "color": "#000000"},
                               end={"sz": 4, "val": "single", "color": "#000000"},
                               )
        layout.set_cell_border(description_table.cell(2, 5),
                               top={"sz": 4, "val": "single", "color": "#000000"},
                               bottom={"sz": 4, "val": "single", "color": "#000000"},
                               start={"sz": 4, "val": "single", "color": "#000000"},
                               end={"sz": 4, "val": "single", "color": "#000000"},
                               )

        # write description next to the four cells and set font size
        description_table.cell(0, 2).text = 'No problem found'
        description_table.cell(0, 6).text = 'Important problem found'
        description_table.cell(2, 2).text = 'Marginal problem found'
        description_table.cell(2, 6).text = 'Critical problem found'
        description_table.cell(0, 2).paragraphs[0].runs[0].font.size = Pt(9)
        description_table.cell(0, 6).paragraphs[0].runs[0].font.size = Pt(9)
        description_table.cell(2, 2).paragraphs[0].runs[0].font.size = Pt(9)
        description_table.cell(2, 6).paragraphs[0].runs[0].font.size = Pt(9)

        # set text vertical alignment of all cells to center
        for row in description_table.rows:
            for cell in row.cells:
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    def write_problem_description(self):
        """
        Write the description of the different problems.
        """

        for i in range(1, self.parameters['Number of problems'] + 1):
            self.report.add_paragraph(self.parameters['Problem {} description'.format(i)],
                                      'List Number')

    def write_chapter(self):
        """
        Write the whole chapter 'Effectiveness analysis', including the tables.
        """

        effectiveness_analysis = ResultsChapter(self.report, self.text_input, self.text_input_soup, self.TITLE,
                                                self.tables, self.parameters)

        self.report.add_paragraph(self.TITLE, self.TITLE_STYLE)
        self.make_result_table()
        self.make_description_table()

        self.report.add_paragraph(self.DESCRIPTION_TITLE, self.DESCRIPTION_STYLE)
        self.write_problem_description()

        self.report.add_paragraph(self.DISCUSSION_TITLE, self.DISCUSSION_STYLE)
        effectiveness_analysis.write_chapter()


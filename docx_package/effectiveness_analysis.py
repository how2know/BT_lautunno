from docx.document import Document
from docx.table import Table
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT, WD_ROW_HEIGHT_RULE
from docx.shared import Pt
from bs4 import BeautifulSoup
from typing import List, Dict, Union
import numpy as np
import seaborn as sns
import matplotlib.pyplot as plt
import pandas as pd

from docx_package.layout import Layout
from docx_package.results import ResultsChapter


class EffectivenessAnalysis:
    """
    Class that represents the 'Effectiveness analysis' chapter and the visualization of its results.
    """

    # name of tables as they appear in the tables list
    TASK_TABLE_NAME = 'Effectiveness analysis tasks and problems table'
    PROBLEM_TABLE_NAME = 'Effectiveness analysis problem type table'
    VIDEO_TABLE_NAME = 'Effectiveness analysis video table'

    # parameter keys
    TASKS_NUMBER_KEY = 'Number of critical tasks'
    PARTICIPANTS_NUMBER_KEY = 'Number of participants'
    PROBLEMS_NUMBER_KEY = 'Number of problems'

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

    # list of the colors table columns width and table rows height
    COLORS_TABLE_WIDTHS = [2, 0.5, 3.8, 1.6, 1.7, 0.5, 3.8, 2]
    COLORS_TABLE_HEIGHTS = [0.5, 0.2, 0.5]

    def __init__(self,
                 report_document: Document,
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

    '''
    def problems_dataframe(self):
        # dataframe = pd.DataFrame()

        rows_number = self.parameters[self.TASKS_NUMBER_KEY] + 1
        cols_number = self.parameters[self.PARTICIPANTS_NUMBER_KEY] + 1

        problems_number = self.parameters[self.PROBLEMS_NUMBER_KEY]

        dataframe = pd.DataFrame(columns=['Participant{}'.format(i) for i in range(1, cols_number)],
                                 index=['Problem{}'.format(i) for i in range(1, problems_number + 1)])

        for i in range(rows_number):
            for j in range(cols_number):

                # skip the first row and first column
                if i != 0 and j != 0:
                    problem = self.task_table.cell(i, j).text

                    if problem:
                        dataframe.loc['Problem{}'.format(problem)].at['Participant{}'.format(j)] = 'Yes'

        print(dataframe)
    '''

    def make_result_table(self):
        """
        Create a table for the visualization of the effectiveness analysis.
        """

        # create a table for the results visualization
        rows_number = self.parameters[self.TASKS_NUMBER_KEY] + 1
        cols_number = self.parameters[self.PARTICIPANTS_NUMBER_KEY] + 1
        result_table = self.report.add_table(rows_number, cols_number)
        result_table.style = 'Table Grid'
        result_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        result_table.autofit = False

        # write the information of the input table in the result table
        for i in range(rows_number):
            for j in range(cols_number):
                cell = result_table.cell(i, j)

                # skip the first row and first column
                if i != 0 and j != 0:
                    cell.text = self.task_table.cell(i, j).text
                    cell.paragraphs[0].runs[0].font.bold = True

                # first row
                elif i == 0 and j != 0:
                    cell.text = self.task_table.cell(i, j).text
                    Layout.set_cell_shading(cell, self.LIGHT_GREY_10)     # color the cell in light_grey_10
                    cell.paragraphs[0].runs[0].font.size = Pt(9)
                    cell.paragraphs[0].runs[0].font.bold = True

                # first column
                elif i != 0 and j == 0:
                    cell.text = self.parameters['Critical task {} name'.format(i)]
                    Layout.set_cell_shading(cell, self.LIGHT_GREY_10)     # color the cell in light_grey_10
                    cell.paragraphs[0].runs[0].font.bold = True

                # bolds all text
                if cell.text:
                    cell.paragraphs[0].runs[0].font.bold = True

        # color the cell according to the type of problem
        for i in range(1, rows_number):
            for j in range(1, cols_number):
                cell = result_table.cell(i, j)

                if cell.text:     # check if the text string is not empty
                    problem_index = int(cell.text)

                    if self.parameters['Problem {} type'.format(problem_index)] == 'Important problem':
                        Layout.set_cell_shading(cell, self.ORANGE)

                    if self.parameters['Problem {} type'.format(problem_index)] == 'Marginal problem':
                        Layout.set_cell_shading(cell, self.YELLOW)

                    if self.parameters['Problem {} type'.format(problem_index)] == 'Critical problem':
                        Layout.set_cell_shading(cell, self.RED)
                else:
                    Layout.set_cell_shading(cell, self.GREEN)

        # color the top left cell borders in white
        Layout.set_cell_border(result_table.cell(0, 0),
                               top={"color": "#FFFFFF"},
                               start={"color": "#FFFFFF"}
                               )

        # set the vertical and horizontal alignment of all cells
        for row in result_table.rows:
            for cell in row.cells:
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        for i in range(len(result_table.rows)):
            for j in range(1, len(result_table.columns)):
                result_table.cell(i, j).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

        # set the width of the columns
        Layout.set_column_width(result_table.columns[0], 2.4)
        for j in range(1, cols_number):
            width = (15.9 - 2.4) / (cols_number - 1)
            Layout.set_column_width(result_table.columns[j], width)

        # set the height of the rows
        Layout.set_row_height(result_table.rows[0], 0.5)
        for i in range(1, rows_number):
            Layout.set_row_height(result_table.rows[i], 1.1, rule=WD_ROW_HEIGHT_RULE.AT_LEAST)

    @ staticmethod
    def add_color_description(table: Table, cell_row: int, cell_column: int, color: str, description: str):
        """
        Color the cell of a table and add a description in the following cell.

        Args:
            table: Table that will contain the colors and their description.
            cell_row: Row in which the color cell is located.
            cell_column: Column in which the color cell is located.
            color: Color of the cell.
            description: Description / meaning of the color.
        """

        color_cell = table.cell(cell_row, cell_column)
        description_cell = table.cell(cell_row, cell_column + 1)

        # color the cell and set its borders
        Layout.set_cell_shading(color_cell, color)
        Layout.set_cell_border(color_cell,
                               top={"sz": 4, "val": "single", "color": "#000000"},
                               bottom={"sz": 4, "val": "single", "color": "#000000"},
                               start={"sz": 4, "val": "single", "color": "#000000"},
                               end={"sz": 4, "val": "single", "color": "#000000"},
                               )

        # write description next to the color cell and set the font size
        description_cell.text = description
        description_cell.paragraphs[0].runs[0].font.size = Pt(9)
        description_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    def make_colors_table(self):
        """
        Create a table that describes the color of the effectiveness analysis table.
        """

        # add a line for space between the two tables
        separation_line = self.report.add_paragraph(' ')
        separation_line.runs[0].font.size = Pt(5)

        # create table
        colors_table = self.report.add_table(3, 8)
        colors_table.autofit = False

        # set the height of all rows
        for idx, row in enumerate(colors_table.rows):
            Layout.set_row_height(row, self.COLORS_TABLE_HEIGHTS[idx])

        # set the width of all columns
        for idx, column in enumerate(colors_table.columns):
            Layout.set_column_width(column, self.COLORS_TABLE_WIDTHS[idx])

        # add the color description to the table in the corresponding cells
        self.add_color_description(colors_table, 0, 1, self.GREEN, 'No problem found')
        self.add_color_description(colors_table, 0, 5, self.ORANGE, 'Important problem found')
        self.add_color_description(colors_table, 2, 1, self.YELLOW, 'Marginal problem found')
        self.add_color_description(colors_table, 2, 5, self.RED, 'Critical problem found')

    def write_problem_description(self):
        """
        Write the description of the different problems.
        """

        for i in range(1, self.parameters['Number of problems'] + 1):
            self.report.add_paragraph(self.parameters['Problem {} description'.format(i)],
                                      'List Number 2')

    def write_chapter(self):
        """
        Write the whole chapter 'Effectiveness analysis', including the tables.
        """

        effectiveness_analysis = ResultsChapter(self.report, self.text_input, self.text_input_soup, self.TITLE,
                                                self.tables, self.parameters)

        self.report.add_paragraph(self.TITLE, self.TITLE_STYLE)
        self.make_result_table()
        self.make_colors_table()

        self.report.add_paragraph(self.DESCRIPTION_TITLE, self.DESCRIPTION_STYLE)
        self.write_problem_description()

        self.report.add_paragraph(self.DISCUSSION_TITLE, self.DISCUSSION_STYLE)
        effectiveness_analysis.write_chapter()

        '''self.problems_dataframe()'''

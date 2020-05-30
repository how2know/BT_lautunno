from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
import numpy as np
import seaborn as sns
import matplotlib.pyplot as plt
import pandas as pd

from docx_package import layout
from docx_package.results import ResultsChapter


class DwellTimesAndRevisits:

    TITLE = 'Dwell times and revisits'
    TITLE_STYLE = 'Heading 2'
    DISCUSSION_TITLE = 'Discussion'
    DISCUSSION_STYLE = 'Heading 3'

    START_TIME = 'Start time'
    END_TIME = 'End time'
    FIXATION_TIME = 'Fixation time'
    LABEL = 'Label'

    TABLE_FIRST_ROW = ['AOI', 'Dwell times [ms]', 'Revisits']

    # color for cell shading
    LIGHT_GREY_10 = 'D0CECE'

    def __init__(self, report_document, text_input_document, text_input_soup, list_of_tables, parameters_dictionary, txt_data):
        self.report = report_document
        self.text_input = text_input_document
        self.text_input_soup = text_input_soup
        self.tables = list_of_tables
        self.parameters = parameters_dictionary
        self.txt_data = txt_data

    @ property
    def areas_of_interest(self):
        labels_list = self.txt_data.index.values.tolist()
        areas_of_interest = []

        for label in labels_list:
            if label not in areas_of_interest:
                areas_of_interest.append(label)

        return areas_of_interest

    def dwell_times(self):
        aois = self.areas_of_interest
        dwell_times_vector = np.zeros(len(aois))

        '''
        for index, aoi in enumerate(aois):
            for line in self.txt_data.iterrows():
                if line[0] == aoi:
                    print(line['Fixation time'])
                    # dwell_times_vector[index] += line[3]
        '''
        '''
        for index, aoi in enumerate(aois):
            for idx, line in self.txt_data.iterrows():

                print(line.index)
        '''
        '''
                if line[0] == aoi:
                    print(line['Fixation time'])
                    # dwell_times_vector[index] += line[3]
                '''


        print(dwell_times_vector)

        print(dwell_times_vector)

    def add_table(self):
        aois = self.areas_of_interest
        table = self.report.add_table(len(aois) + 1, 3)

        table.style = 'Table Grid'  # set the table style
        table.alignment = WD_TABLE_ALIGNMENT.CENTER  # set the table alignment
        table.autofit = True

        for index, label in enumerate(self.TABLE_FIRST_ROW):
            cell = table.rows[0].cells[index]
            cell.text = label
            layout.set_cell_shading(cell, self.LIGHT_GREY_10)  # color the cell in light_grey_10
            cell.paragraphs[0].runs[0].font.bold = True

        for index, aoi in enumerate(aois):
            cell = table.columns[0].cells[index+1]
            cell.text = aoi

        # set the vertical and horizontal alignment of all cells
        for row in table.rows:
            for cell in row.cells:
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                # cell.paragraphs[0].style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
                # cell.paragraphs[0].style.name = 'Table'

    def write_chapter(self):
        time_on_tasks = ResultsChapter(self.report, self.text_input, self.text_input_soup, self.TITLE,
                                       self.tables, self.parameters)

        self.report.add_paragraph(self.TITLE, self.TITLE_STYLE)

        self.add_table()

        self.dwell_times()

        self.report.add_paragraph(self.DISCUSSION_TITLE, self.DISCUSSION_STYLE)
        time_on_tasks.write_chapter()

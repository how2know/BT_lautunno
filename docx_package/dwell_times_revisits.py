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

    def __init__(self, report_document, text_input_document, text_input_soup, list_of_tables, parameters_dictionary, list_of_dataframes):
        self.report = report_document
        self.text_input = text_input_document
        self.text_input_soup = text_input_soup
        self.tables = list_of_tables
        self.parameters = parameters_dictionary
        self.cGOM_dataframes = list_of_dataframes

    def areas_of_interest(self, dataframe):
        labels_list = dataframe.index.values.tolist()
        areas_of_interest = []

        for label in labels_list:
            if label not in areas_of_interest:
                areas_of_interest.append(label)

        return areas_of_interest

    def dwell_times(self, aois, dataframe):
        dwell_times_vector = np.zeros(len(aois))

        for idx, aoi in enumerate(aois):
            data_of_aoi = dataframe[dataframe.index == aoi]
            dwell_times = data_of_aoi['Fixation time'].sum()
            dwell_times_vector[idx] = dwell_times

        return dwell_times_vector

    def revisits(self, aois, dataframe):
        revisits_list = []

        for idx, aoi in enumerate(aois):
            data_of_aoi = dataframe[dataframe.index == aoi]
            revisits = len(data_of_aoi['Fixation time']) - 1
            revisits_list.append(revisits)

        return revisits_list

    def dwell_times_stat(self):
        dwell_times_df = pd.DataFrame()

        for idx, dataframe in enumerate(self.cGOM_dataframes):
            aois = self.areas_of_interest(dataframe)
            dwell_times = self.dwell_times(aois, dataframe)

            participant_dwell_times = pd.DataFrame(index=['Participant {}'.format(idx+1)],
                                                   columns=aois,
                                                   data=dwell_times.reshape(1, -1)
                                                   )

            dwell_times_df = dwell_times_df.append(participant_dwell_times)

        dwell_times_mean = dwell_times_df.mean().to_numpy()

        dwell_times_mean_df = pd.DataFrame(index=['Mean'],
                                           columns=dwell_times_df.columns,
                                           data=dwell_times_mean.reshape(1, -1)
                                           )

        dwell_times_df = dwell_times_df.append(dwell_times_mean_df)

        print(dwell_times_df)

        return dwell_times_df

    def revisits_stat(self):
        revisits_df = pd.DataFrame()

        for idx, dataframe in enumerate(self.cGOM_dataframes):
            aois = self.areas_of_interest(dataframe)
            revisits = self.revisits(aois, dataframe)

            participant_revisits = pd.DataFrame(index=['Participant {}'.format(idx + 1)],
                                                columns=aois,
                                                data=[revisits]
                                                )

            revisits_df = revisits_df.append(participant_revisits)

        revisits_mean = revisits_df.mean().to_numpy()

        revisits_mean_df = pd.DataFrame(index=['Mean'],
                                        columns=revisits_df.columns,
                                        data=[revisits_mean]
                                        )

        revisits_df = revisits_df.append(revisits_mean_df)

        print(revisits_df)

        return revisits_df

    def add_table(self):
        dwell_times_df = self.dwell_times_stat()
        revisits_df = self.revisits_stat()

        aois = dwell_times_df.columns

        table = self.report.add_table(len(aois) + 1, 3)

        table.style = 'Table Grid'  # set the table style
        table.alignment = WD_TABLE_ALIGNMENT.CENTER  # set the table alignment
        table.autofit = True

        for index, label in enumerate(self.TABLE_FIRST_ROW):
            cell = table.rows[0].cells[index]
            cell.text = label
            layout.set_cell_shading(cell, self.LIGHT_GREY_10)  # color the cell in light_grey_10
            cell.paragraphs[0].runs[0].font.bold = True

        for idx, aoi in enumerate(aois):
            cell = table.columns[0].cells[idx+1]
            cell.text = aoi

        for idx, dwell_times in enumerate(dwell_times_df.loc['Mean'].to_numpy()):
            cell = table.columns[1].cells[idx+1]
            cell.text = str(round(dwell_times, 4))

        for idx, revisits in enumerate(revisits_df.loc['Mean'].to_numpy()):
            cell = table.columns[2].cells[idx+1]
            cell.text = str(round(revisits, 4))

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

        self.report.add_paragraph(self.DISCUSSION_TITLE, self.DISCUSSION_STYLE)
        time_on_tasks.write_chapter()

from docx.document import Document
from docx.table import Table
from bs4 import BeautifulSoup
from typing import List, Dict, Union
import numpy as np
import seaborn as sns
import matplotlib.pyplot as plt
import pandas as pd

from docx_package.results import ResultsChapter
from txt_package import plot


class TimeOnTasks:
    """
    Class that represents the 'Time on tasks' chapter and the visualization of its results.
    """

    # name of table as it appears in the tables list
    TIME_ON_TASK_TABLE_NAME = 'Time on tasks table'

    # parameter keys as they appear in the parameters dictionary
    PARTICIPANTS_NUMBER_KEY = 'Number of participants'
    TASKS_NUMBER_KEY = 'Number of critical tasks'

    # information about the headings of this chapter
    TITLE = 'Time on tasks'
    TITLE_STYLE = 'Heading 2'
    DISCUSSION_TITLE = 'Discussion'
    DISCUSSION_STYLE = 'Heading 3'

    # path to plot image files
    PARTICIPANT_FIGURE_PATH = 'Outputs/Time_on_task_participant{}.png'
    MAIN_FIGURE_PATH = 'Outputs/Time_on_task.png'

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
        self.parameters = parameters_dictionary
        self.tables = list_of_tables
        input_table_index = self.tables.index(self.TIME_ON_TASK_TABLE_NAME)
        self.input_table = text_input_document.tables[input_table_index]

    @ property
    def tasks(self) -> List[str]:
        """
        Returns:
            List of task names.
        """

        tasks = []
        for i in range(1, self.parameters[self.TASKS_NUMBER_KEY] + 1):
            tasks.append(self.parameters['Critical task {} name'.format(i)])
        return tasks

    @ property
    def participants(self) -> List[str]:
        """
        Returns:
            List of participants, i.e. [Participant 1, Participant 2, ...].
        """

        participants = ['participant {}'.format(i) for i in range(1, self.parameters[self.PARTICIPANTS_NUMBER_KEY] + 1)]
        return participants

    @ property
    def times(self) -> np.ndarray:
        """
        Returns:
            Matrix of task completion times.
        """

        rows = self.parameters[self.TASKS_NUMBER_KEY]
        columns = self.parameters[self.PARTICIPANTS_NUMBER_KEY]
        times = np.zeros((rows, columns))
        for i in range(rows):
            for j in range(columns):
                time = float(self.input_table.cell(i+1, j+1).text)
                times[i, j] = time

        # return the transposed matrix to have participants as rows and tasks as columns
        return times.transpose()

    #TODO: check if pd.DataFrame = pandas.core.frame.DataFrame which is the type of the return value here
    @ property
    def times_df(self) -> pd.DataFrame:
        """
        Returns:
            Data frame of tasks completion times with participants as index and task names as columns
        """

        data_frame = pd.DataFrame(self.times, index=self.participants, columns=self.tasks)
        return data_frame

    def make_plots(self):
        for participant in self.participants:
            participant_times = self.times_df.loc[participant].to_numpy()
            participant_times_df = pd.DataFrame(data=[participant_times],
                                                columns=self.times_df.columns)

            plot.make_barplot(data_frame=participant_times_df,
                              figure_save_path='Time_on_task_{}.png'.format(participant),
                              title='Time on task {}'.format(participant),
                              ylabel='Completion time [s]')



    def write_chapter(self):
        # plot.make_barplot(self.times_df, self.FIGURE_NAME, ylabel='Completion time [s]')
        # plot.make_boxplot(self.times_df, 'Outputs/Hey39.png', title='Hey', ylabel='Completion time [s]')

        self.times_df
        self.make_plots()

        time_on_tasks = ResultsChapter(self.report, self.text_input, self.text_input_soup, self.TITLE,
                                       self.tables, self.parameters)

        self.report.add_paragraph(self.TITLE, self.TITLE_STYLE)
        self.report.add_picture(self.MAIN_FIGURE_PATH)
        self.report.add_picture('Outputs/Hey39.png')

        self.report.add_paragraph(self.DISCUSSION_TITLE, self.DISCUSSION_STYLE)
        time_on_tasks.write_chapter()


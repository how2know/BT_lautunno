from docx.document import Document
from docx.table import Table
from bs4 import BeautifulSoup
from typing import List, Dict, Union
import numpy as np
import seaborn as sns
import matplotlib.pyplot as plt
import pandas as pd

from docx_package.results import ResultsChapter
from txt_package import plot, eye_tracking


class Transitions:
    """
    Class that represents the 'Transitions' chapter and the visualization of its results.
    """

    # information about the headings of this chapter
    TITLE = 'Transitions'
    TITLE_STYLE = 'Heading 2'
    DISCUSSION_TITLE = 'Discussion'
    DISCUSSION_STYLE = 'Heading 3'

    # path to plot image files
    PARTICIPANT_FIGURE_PATH = 'Outputs/Transitions_participant{}.png'
    HEAT_MAP_FIGURE_PATH = 'Outputs/Transitions_heat_map.png'

    def __init__(self,
                 report_document: Document,
                 text_input_document: Document,
                 text_input_soup: BeautifulSoup,
                 list_of_tables: List[str],
                 parameters_dictionary: Dict[str, Union[str, int]],
                 list_of_dataframes: List[pd.DataFrame]
                 ):
        """
        Args:
            report_document: .docx file where the report is written.
            text_input_document: .docx file where all inputs are written.
            text_input_soup: BeautifulSoup of the xml of the input .docx file.
            list_of_tables: List of all table names.
            parameters_dictionary: Dictionary of all input parameters (key = parameter name, value = parameter value)
            list_of_dataframes: List of data frames containing the cGOM data of each participant
        """

        self.report = report_document
        self.text_input = text_input_document
        self.text_input_soup = text_input_soup
        self.tables = list_of_tables
        self.parameters = parameters_dictionary
        self.cGOM_dataframes = list_of_dataframes

    def makes_plot(self):
        """
        Create heat maps to visualize the transitions percentage between the AOIs.

        One heat map for each participant is created.
        One heat map with the data of all participants are created.
        """

        # TODO: choose which kind of data we take for the main heat map

        # main data frame that will contain the data from the data frames from all participants
        all_transitions = pd.DataFrame()

        # main data frame that will contain the transitions percentage of all participants
        all_data = pd.DataFrame()

        # create a data frame with the transitions percentage for each participant, create a heat map with it,
        # and append it to the main data frame
        for idx, dataframe in enumerate(self.cGOM_dataframes):
            aois = eye_tracking.areas_of_interest(dataframe)
            participant_transitions = eye_tracking.transitions(aois, dataframe)
            plot.make_heatmap(data_frame=participant_transitions,
                              figure_save_path=self.PARTICIPANT_FIGURE_PATH.format(str(idx + 1)),
                              title='Transitions: participant {}'.format(str(idx + 1)),
                              xlabel='AOI destination (to)',
                              ylabel='AOI source (from)'
                              )
            all_transitions = all_transitions.append(participant_transitions)
            '''all_data = all_data.append(dataframe)'''

        # create a data frame with the mean of transitions for each AOI and append it
        # to a data frame that will contain all means, and then create a heat map with all means
        # Problem: The sum of all percentage is not equal to 1.
        all_aois = all_transitions.columns.tolist()
        transitions_stat = pd.DataFrame()
        for idx, aoi in enumerate(all_aois):
            data_of_aoi = all_transitions[all_transitions.index == aoi]
            transitions_mean = data_of_aoi.mean().to_numpy()
            transitions_mean_df = pd.DataFrame(index=[aoi],
                                               columns=all_transitions.columns,
                                               data=[transitions_mean]
                                               )
            transitions_stat = transitions_stat.append(transitions_mean_df)

        plot.make_heatmap(transitions_stat,
                          figure_save_path=self.HEAT_MAP_FIGURE_PATH,
                          title='Transitions',
                          xlabel='AOI destination (to)',
                          ylabel='AOI source (from)'
                          )

        '''print(transitions_stat.to_numpy().sum())     # sum of all percentage != 1'''

        '''
        # create a data frame with the transitions of all participants and a heat map out of it
        # Problem: The data of all participants is consecutively contained in a list. 
        #          A transition will therefore be added from the last AOI of one participant 
        #          to the first AOI of the next participant, and this is not correct.
        all_aois = self.areas_of_interest(all_data)
        transitions_stat = eye_tracking.transitions(all_aois, all_data)
        plot.make_heatmap(transitions_stat,
                          figure_save_path=self.HEAT_MAP_FIGURE_PATH,
                          title='Transitions',
                          xlabel='AOI destination (to)',
                          ylabel='AOI source (from)'
                          )
        print(transitions_stat.to_numpy().sum())     # sum of all percentage = 1
        '''

    def write_chapter(self):
        """
        Write the whole chapter 'Transitions', including plot.
        """

        self.makes_plot()

        transitions = ResultsChapter(self.report, self.text_input, self.text_input_soup, self.TITLE,
                                     self.tables, self.parameters)

        self.report.add_paragraph(self.TITLE, self.TITLE_STYLE)
        self.report.add_picture(self.HEAT_MAP_FIGURE_PATH)

        self.report.add_paragraph(self.DISCUSSION_TITLE, self.DISCUSSION_STYLE)
        transitions.write_chapter()

"""Eye tracking module

Some methods that represents basic eye tracking data features are implemented here.

Those methods are needed to create the visualization part of the report.
Thus, they are called in the class AverageFixation, DwellTimesAndRevisits and Transitions.
"""

from docx.document import Document
from docx.table import Table
from bs4 import BeautifulSoup
from typing import List, Dict, Union
from docx import Document
import os
from zipfile import ZipFile
from bs4 import BeautifulSoup
import time
from itertools import islice
import numpy as np
import seaborn as sns
import matplotlib.pyplot as plt
import pandas as pd

# TODO: names of columns as constant variable, e.g. START_TIME = 'Start time'

def areas_of_interest(dataframe: pd.DataFrame) -> List[str]:
    """
    Args:
        dataframe: Data frame that have AOIs as index.

    Returns:
        List of AOIs given in a data frame.
    """

    # read all indexes of a data frame and append it to a list if it is already in
    aois = []
    labels_list = dataframe.index.values.tolist()
    for label in labels_list:
        if label not in aois:
            aois.append(label)

    return aois


def fixations(aois: List[str], dataframe: pd.DataFrame) -> pd.DataFrame:
    """
    Args:
        aois: List of AOIs.
        dataframe: Data frame that have a column 'Fixation time'.

    Returns:
        Data frame that contains all fixations for each AOI with the AOIs as columns.
    """

    # main fixations data frame
    fixations_df = pd.DataFrame()

    # create a data frame for each AOI with all its fixations and append it to the main data frame
    for idx, aoi in enumerate(aois):
        data_of_aoi = dataframe[dataframe.index == aoi]
        fixations_vector = data_of_aoi['Fixation time'].to_numpy()
        aoi_fixations = pd.DataFrame(columns=[aoi], data=fixations_vector)
        fixations_df = fixations_df.append(aoi_fixations, ignore_index=True)

    return fixations_df

'''
def dwell_times1(aois: List[str], dataframe: pd.DataFrame) -> np.ndarray:
    """
    Args:
        aois: List of AOIs.
        dataframe: Data frame that have a column 'Fixation time'.

    Returns:
        Vector of the dwell times for each AOI.
    """

    # main dwell times vector
    dwell_times_vector = np.zeros(len(aois))

    # append the sum of all fixations, i.e. the dwell time, of each AOI to the main vector
    for idx, aoi in enumerate(aois):
        data_of_aoi = dataframe[dataframe.index == aoi]
        dwell_times = data_of_aoi['Fixation time'].sum()
        dwell_times_vector[idx] = dwell_times

    return dwell_times_vector
'''


def dwell_times(aois: List[str], dataframe: pd.DataFrame) -> pd.DataFrame:
    """
    Calculate the dwell times and make some statistics out of them.

    A dwell time is defined as the time one look at AOI, i.e. the time between
    the start time of fixation to a AOI and the end time of the last fixation to the AOI
    before one change to another AOI.

    Args:
        aois: List of AOIs.
        dataframe: Data frame that have columns 'Start time' and 'End time'.

    Returns:
        Data frame with all dwell times and the statistics for each AOI.

        The indexes are the AOIs and the columns are the statistics,
        i.e. 'Dwell times', 'Sum', 'Mean', 'Max', 'Min'.
    """

    # main data frame that contains all the dwell times
    dwell_times_df = pd.DataFrame()

    # values for the first fixation
    previous_aoi = dataframe.index[0]
    previous_start = dataframe.iloc[0]['Start time']
    previous_end = dataframe.iloc[0]['End time']

    for row in dataframe.iterrows():

        # values of the current fixation
        aoi = row[0]
        start_time = row[1]['Start time']
        end_time = row[1]['End time']

        # if the AOI has not changed, the dwell time is not finish
        # thus, the start time remains the same but the end time changes to the one of the current fixation
        if aoi == previous_aoi:
            previous_end = end_time

        # if the AOI has changed, the dwell time is finish
        if aoi != previous_aoi:

            # create a data frame for the dwell time and append it to the main data frame
            dwell_time_df = pd.DataFrame(data=previous_end - previous_start,
                                         index=[previous_aoi],
                                         columns=['Dwell times']
                                         )
            dwell_times_df = dwell_times_df.append(dwell_time_df)

            # both start and end time change to the one of the current fixation
            previous_start = start_time
            previous_end = end_time

        # change to the next fixation
        previous_aoi = aoi

    # append the sum of all fixations, i.e. the dwell time, of each AOI to the main vector
    for idx, aoi in enumerate(aois):
        data_of_aoi = dwell_times_df[dwell_times_df.index == aoi]

        # make statistics for the dwell times of this AOI
        time_sum = data_of_aoi['Dwell times'].sum()
        time_mean = data_of_aoi['Dwell times'].mean()
        time_max = data_of_aoi['Dwell times'].max()
        time_min = data_of_aoi['Dwell times'].min()

        # create data frames with the statistics and append them to the main data frame
        sum_df = pd.DataFrame(data=time_sum, index=[aoi], columns=['Sum'])
        mean_df = pd.DataFrame(data=time_mean, index=[aoi], columns=['Mean'])
        max_df = pd.DataFrame(data=time_max, index=[aoi], columns=['Max'])
        min_df = pd.DataFrame(data=time_min, index=[aoi], columns=['Min'])
        dwell_times_df = dwell_times_df.append(sum_df)
        dwell_times_df = dwell_times_df.append(mean_df)
        dwell_times_df = dwell_times_df.append(max_df)
        dwell_times_df = dwell_times_df.append(min_df)

    return dwell_times_df


def transitions(aois: List[str], dataframe: pd.DataFrame) -> pd.DataFrame:
    """
    Args:
        aois: List of AOIs.
        dataframe: Data frame that have AOIs as index.

    Returns:
        Data frame that have AOIs as indexes and columns and the fixation ratio for all AOIs as entries.

    """

    # create a data frame with the AOIs as columns and indexes and zeros as entries
    transitions_table = pd.DataFrame(index=aois,
                                     columns=aois,
                                     data=np.zeros((len(aois), len(aois)))
                                     )

    # list of all AOIs that were looked (fixations) in the order they appeared
    all_fixations_aoi = dataframe.index.values.tolist()

    # add one transition from the last fixation AOI (index) to the actual fixation AOI (column)
    last_fixation_aoi = all_fixations_aoi[0]
    for fixation_aoi in all_fixations_aoi[1:]:
        transitions_table.loc[last_fixation_aoi].at[fixation_aoi] += 1
        last_fixation_aoi = fixation_aoi

    # divide all entries by the total number of fixations to get a ratio (or percentage)
    transitions_number = transitions_table.to_numpy().sum()
    transitions_table = transitions_table.div(transitions_number)

    return transitions_table


def revisits(aois: List[str], dataframe: pd.DataFrame) -> List[int]:
    """
    Args:
        aois: List of AOIs.
        dataframe: Data frame that have a column 'Fixation time'.

    Returns:
        List of the number of revisits for each AOI.
    """

    # main revisits list
    revisits_list = []

    # append the number of fixations - 1, i.e. the number of revisits, of each AOI to the main list
    for idx, aoi in enumerate(aois):
        data_of_aoi = dataframe[dataframe.index == aoi]
        revisits = len(data_of_aoi['Fixation time']) - 1
        revisits_list.append(revisits)

    return revisits_list



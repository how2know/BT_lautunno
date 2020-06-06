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


def dwell_times(aois: List[str], dataframe: pd.DataFrame) -> np.ndarray:
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



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


# TODO: delete the label column
def make_dataframe(txt_file_path):
    """
    Reads the .txt file that contains the data from cGOM and returns a data frame out of it.
    The indexes of the data frame are the label, i.e. the AOI.
    The columns of the data frame are the start time of a fixation, end time of a fixation, duration of a fixation.
    """

    # names of the data frame columns
    start_time = 'Start time'
    end_time = 'End time'
    fixation_time = 'Fixation time'
    # label = 'Label'

    start_times_list = []
    end_times_list = []
    labels_list = []

    # reads the file and stores the value in the corresponding lists
    with open(txt_file_path, 'r') as file:
        for line in islice(file, 1, None):
            start_times_list.append(float(line.split()[0]))
            end_times_list.append(float(line.split()[1]))
            labels_list.append(line.split()[2])
    file.close()

    # creates numpy arrays from lists
    start_times_vector = np.array(start_times_list)
    end_times_vector = np.array(end_times_list)
    fixation_times_vector = end_times_vector - start_times_vector

    # creates pandas data frame
    data = pd.DataFrame(index=labels_list, columns=[start_time, end_time, fixation_time])
    data[start_time] = start_times_list
    data[end_time] = end_times_list
    data[fixation_time] = fixation_times_vector
    # data[label] = labels_list

    return data

# TODO: maybe give the directory_path as argument
def make_dataframes_list(parameters):
    """
    Creates a data frame from the cGOM data of each participant and returns a list of the data frames.
    The files containing the data must be named Participant<Number>.txt, e.g. 'Participant3.txt', and stored in
    the Inputs/Data directory.
    """
    # path to the .txt files
    directory_path = 'Inputs/Data/Participant{}.txt'

    participants_number = parameters['Number of participants']

    dataframes_list = []

    for i in range(1, participants_number + 1):
        txt_file_path = directory_path.format(str(i))

        # stores a data frame in the list or passes if the file is not provided
        try:
            dataframes_list.append(make_dataframe(txt_file_path))
        except FileNotFoundError:
            pass

    return dataframes_list


def areas_of_interest(dataframe):
    labels_list = dataframe.index.values.tolist()
    areas_of_interest = []

    for label in labels_list:
        if label not in areas_of_interest:
            areas_of_interest.append(label)

    return areas_of_interest


def transitions(self, aois, dataframe):
    table = pd.DataFrame(index=aois,
                         columns=aois,
                         data=np.zeros((len(aois), len(aois)))
                         )

    all_fixations = dataframe.index.values.tolist()

    last_fixation = all_fixations[0]

    for fixation in all_fixations[1:]:
        # table.at[last_fixation, fixation] += 1
        table.loc[last_fixation].at[fixation] += 1

        last_fixation = fixation

    transitions_number = table.to_numpy().sum()

    table = table.div(transitions_number)
    # table = table.mul(100)

    return table


def dwell_times(aois, dataframe):
    dwell_times_vector = np.zeros(len(aois))

    for idx, aoi in enumerate(aois):
        data_of_aoi = dataframe[dataframe.index == aoi]
        dwell_times = data_of_aoi['Fixation time'].sum()
        dwell_times_vector[idx] = dwell_times

    return dwell_times_vector


def revisits(aois, dataframe):
    revisits_list = []

    for idx, aoi in enumerate(aois):
        data_of_aoi = dataframe[dataframe.index == aoi]
        revisits = len(data_of_aoi['Fixation time']) - 1
        revisits_list.append(revisits)

    return revisits_list


def fixations(aois, dataframe):
    participant_fixations_df = pd.DataFrame()

    for idx, aoi in enumerate(aois):
        data_of_aoi = dataframe[dataframe.index == aoi]
        fixations = data_of_aoi['Fixation time'].to_numpy()
        aoi_fixations = pd.DataFrame(columns=[aoi], data=fixations)

        participant_fixations_df = participant_fixations_df.append(aoi_fixations, ignore_index=True)

    return participant_fixations_df



from docx import Document
import os
from zipfile import ZipFile
from bs4 import BeautifulSoup
import time
from os import listdir
from itertools import islice
import numpy as np
import seaborn as sns
import matplotlib.pyplot as plt
import pandas as pd


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
    dataframe = pd.DataFrame(index=labels_list, columns=[start_time, end_time, fixation_time])
    dataframe[start_time] = start_times_list
    dataframe[end_time] = end_times_list
    dataframe[fixation_time] = fixation_times_vector

    return dataframe


# TODO: maybe give the directory_path as argument
def make_dataframes_list(parameters):
    """
    Creates a data frame from the cGOM data of each participant and returns a list of the data frames.
    The files containing the data must be named Participant<Number>.txt, e.g. 'Participant3.txt', and stored in
    the Inputs/Data directory.
    """

    # list of all files stored in the directory 'Inputs/Data'
    files = listdir('Inputs/Data')

    directory_path2 = 'Inputs/Data/{}'

    dataframes_list2 = []

    print(files)

    biggest_number = 1

    for file in files:
        if file.startswith('Participant') and file.endswith('.txt'):
            number = file.replace('Participant', '').replace('.txt', '')
            if number.isdigit():
                if int(number) > biggest_number:
                    biggest_number = int(number)

                '''
                txt_file_path = directory_path2.format(file)

                dataframe = make_dataframe(txt_file_path)
                if not dataframe.empty:
                    dataframes_list2.append(dataframe)
                '''

                '''

                # stores a data frame in the list or passes if the file is not provided
                try:
                    dataframe = make_dataframe(txt_file_path)
                    if not dataframe.empty:
                        dataframes_list2.append(dataframe)
                except FileNotFoundError:
                    pass
                '''


    # path to the .txt files
    directory_path = 'Inputs/Data/Participant{}.txt'

    # participants_number = parameters['Number of participants']


    dataframes_list = []

    for i in range(biggest_number+1):
        txt_file_path = directory_path.format(str(i))

        # stores a data frame in the list or passes if the file is not provided
        try:
            dataframes_list.append(make_dataframe(txt_file_path))
            print(txt_file_path)
        except FileNotFoundError:
            pass

    return dataframes_list





from os import listdir
from itertools import islice
import numpy as np
import pandas as pd
from typing import List


class cGOM:
    """
    Class that represents the cGOM data and creates a readable form of it, i.e. data frames.

    cGOM data must be loaded in .txt files in the 'Inputs/cGOM_data' directory.
    """

    # path to the cGOM directory and cGOM .txt files
    cGOM_DIRECTORY_PATH = 'Inputs/cGOM_data'
    cGOM_FILES_PATH = 'Inputs/cGOM_data/Participant{}.txt'

    # names of the columns of the cGOM .txt files
    START_TIME = 'Start time'
    END_TIME = 'End time'
    FIXATION_TIME = 'Fixation time'

    def __init__(self):
        pass

    def make_dataframe(self, txt_file_path: str) -> pd.DataFrame:
        """
        Args:
            txt_file_path: Path of the .txt file that contains the data.

        Returns:
            Data frame with the data of the cGOM .txt file.

            The indexes of the data frame are the label, i.e. the AOI.
            The columns of the data frame are the start time of a fixation,
            end time of a fixation, and duration of a fixation.
        """

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
        dataframe = pd.DataFrame(index=labels_list, columns=[self.START_TIME, self.END_TIME, self.FIXATION_TIME])
        dataframe[self.START_TIME] = start_times_list
        dataframe[self.END_TIME] = end_times_list
        dataframe[self.FIXATION_TIME] = fixation_times_vector

        # rename BG in Background
        dataframe = dataframe.rename(index={'BG': 'Background'})

        return dataframe

    @ classmethod
    def make_dataframes_list(cls) -> List[pd.DataFrame]:
        """
        Creates a data frame from the cGOM data of each participant and returns a list of the data frames.

        Notes:
            The files containing the data must be named Participant<Number>.txt, e.g. 'Participant3.txt',
            and stored in the Inputs/cGOM_data directory.

        Returns:
            List of data frames that contain the cGOM data of each participant.
        """

        cGOM = cls()

        # list of all files stored in the directory 'Inputs/Data'
        files = listdir(cGOM.cGOM_DIRECTORY_PATH)

        biggest_number = 1

        # look for all files that are named in the form 'Participant<Number>.txt' and get the biggest value of <Number>
        for file in files:
            if file.startswith('Participant') and file.endswith('.txt'):
                try:
                    number = int(file.replace('Participant', '').replace('.txt', ''))
                    if number > biggest_number:
                        biggest_number = int(number)
                except ValueError:
                    pass

        # path to the .txt files
        files_path = cGOM.cGOM_FILES_PATH

        dataframes_list = []

        for i in range(biggest_number + 1):
            txt_file_path = files_path.format(str(i))

            # store the data frames in the list and skip the empty ones
            try:
                dataframe = cGOM.make_dataframe(txt_file_path)
                if not dataframe.empty:
                    dataframes_list.append(dataframe)

            # pass if the file is not provided
            except FileNotFoundError:
                pass

        return dataframes_list

from typing import Dict, Union
import pandas as pd
from os import listdir


class TobiiData:
    """
    Class that represents the Tobii data and creates a readable form of it, i.e. data frames.

    Tobii data must be loaded in .tsv files in the 'Inputs/Tobii_data' directory.
    """

    # labels of the columns of the data frame
    PARTICIPANTS_LABEL = 'Participant name'
    TIMESTAMPS_LABEL = 'Recording timestamp'
    EVENT_LABEL = 'Event'
    SECONDS_LABEL = 'Seconds'

    def __init__(self, parameters_dictionary: Dict[str, Union[str, int]]):
        """
        Args:
            parameters_dictionary: Dictionary of all input parameters (key = parameter name, value = parameter value).
        """

        self.parameters = parameters_dictionary

    def make_dataframe(self, tsv_file_path) -> pd.DataFrame:
        """
        Args:
            tsv_file_path: Path to the .tsv file that contains the Tobii data.

        Returns:
            Data frame which has the participants as index, the tasks in the 'Event' column
            and the times in the 'Seconds' column.
        """

        # create a data frame from the .tsv file
        tobii_df = pd.read_csv(tsv_file_path, sep='\t')

        # create a data frame with the relevant columns, i.e. 'Recording timestamp', 'Event',
        # and 'Participant name' as index
        timestamps = tobii_df[self.TIMESTAMPS_LABEL]
        events = tobii_df[self.EVENT_LABEL]
        participants = tobii_df[self.PARTICIPANTS_LABEL]
        tasks_times_df = pd.concat([timestamps, events, participants], axis=1)
        tasks_times_df.set_index(self.PARTICIPANTS_LABEL, inplace=True)

        # delete all irrelevant rows and keep only the ones that describes a task
        tasks_times_df = tasks_times_df.dropna()
        tasks_times_df = tasks_times_df[tasks_times_df[self.EVENT_LABEL].str.contains('Task')]

        # convert the timestamp, that are in microseconds, in seconds and create a new column
        microseconds = tasks_times_df[self.TIMESTAMPS_LABEL].to_numpy()
        seconds = microseconds / 1000000
        tasks_times_df[self.SECONDS_LABEL] = seconds

        return tasks_times_df

    @ classmethod
    def make_main_dataframe(cls, parameters_dictionary: Dict[str, Union[str, int]]) -> pd.DataFrame:
        """
        Create a data frame that contains the relevant data from Tobii, i.e. 'Participant name', 'Event' of the tasks,
        and times in 'Seconds'.

        The data frame is created either from a .tsv file that already contains the data of all participants or
        by combining the provided data frames of every participant.

        Args:
            parameters_dictionary: Dictionary of all input parameters (key = parameter name, value = parameter value).

        Returns:
            Data frame which has the participants as index, the tasks in the 'Event' column
            and the times in the 'Seconds' column.
        """

        tobii = cls(parameters_dictionary)

        # list of all files stored in the directory 'Inputs/Data'
        files = listdir('Inputs/Tobii_data')

        # create directly a data frame with .tsv files of all participants if there is one
        if 'All_participants.tsv' in files:
            tobii_df = tobii.make_dataframe('Inputs/Tobii_data/All_participants.tsv')

        # create a data frame with the .tsv files provided for the different participants
        else:
            # main data frame that will contain the data of all participants
            tobii_df = pd.DataFrame()

            # create a data frame with the Tobii data of each participants
            # and append it to the main data frame
            for i in range(16):
                try:
                    participant_df = tobii.make_dataframe('Inputs/Tobii_data/Participant{}.tsv'.format(i))
                    indexes = ['Participant{}'.format(i)] * len(participant_df)
                    participant_df.index = indexes
                    tobii_df = tobii_df.append(participant_df)

                # do nothing if no file is provided for a participant
                except FileNotFoundError:
                    pass

        # create an empty with the relevant columns if no .tsv file was provided
        if tobii_df.empty:
            tobii_df = pd.DataFrame(columns=[tobii.EVENT_LABEL, tobii.SECONDS_LABEL])

        return tobii_df

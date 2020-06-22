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


class TobiiData:

    def __init__(self):
        pass

    @ staticmethod
    def make_dataframe(txt_file_path):
        """
        Reads the .txt file that contains the data from cGOM and returns a data frame out of it.
        The indexes of the data frame are the label, i.e. the AOI.
        The columns of the data frame are the start time of a fixation, end time of a fixation, duration of a fixation.
        """

        dataframe = pd.DataFrame()

        start = time.time()

        tobii_dataframe = pd.read_csv(txt_file_path, sep='\t')
        # print(tobii_dataframe)

        end = time.time()
        print('read csv: ', end-start)

        start = time.time()

        tobii_dataframe2 = pd.read_table(txt_file_path, sep='\t')
        # print(tobii_dataframe2)

        end = time.time()
        print('read table: ', end-start)

        timestamps = tobii_dataframe['Recording timestamp']
        timestamps2 = tobii_dataframe['Recording timestamp'].to_numpy()
        events = tobii_dataframe['Event']

        events2 = tobii_dataframe['Event'].to_numpy()

        print(pd.DataFrame(data=events2, columns=['Event']))
        print(pd.DataFrame(data=timestamps2, columns=['Recording timestamp']))

        # print(events)
        # print(timestamps)

        # dataframe = dataframe.append(timestamps)
        # dataframe = dataframe.append(events)

        # dataframe = dataframe.append(pd.DataFrame(data=timestamps2, columns=['Recording timestamp']))
        # dataframe = dataframe.append(pd.DataFrame(data=events2, columns=['Event']))

        '''dataframe = pd.concat([pd.DataFrame(data=timestamps2, columns=['Recording timestamp']), pd.DataFrame(data=events2, columns=['Event'])], axis=1)'''

        dataframe = pd.concat([timestamps, events], axis=1)

        print(dataframe)

        dataframe = dataframe.dropna()

        print(dataframe)

        return dataframe


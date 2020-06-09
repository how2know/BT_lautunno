#####     TEXT READING MODULE     #####

# Functions that read the text from the text input file are implemented in this module.

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


# return the path of a file given its name and the name of its directory
def get_path(file_name, directory_name):
    text_input_directory = os.path.dirname(file_name)
    path = os.path.join(text_input_directory, directory_name, file_name)

    return path

'''
# read dropdown lists and store their value in a list
def read_dropdown_lists(file_name, list_of_value):

    # open docx file as a zip file and store its relevant xml data
    zip_file = ZipFile(file_name)
    xml_data = zip_file.read('word/document.xml')
    zip_file.close()

    # parse the xml data with BeautifulSoup
    soup = BeautifulSoup(xml_data, 'xml')

    # look for all values of dropdown lists in the data and store them
    dd_lists_content = soup.find_all('sdtContent')
    for i in dd_lists_content:
        list_of_value.append(i.find('t').string)
'''


def parse_xml_with_bs4(text_input_path):
    """
    Opens a docx file as a zip file, stores the xml data containing the infos about the docx document
    and return the parsed xml data as BeautifulSoup object.

    This takes some time, that's why it is better to run it once at the beginning and not in a loop.
    """

    # open docx file as a zip file and store its relevant xml data
    zip_file = ZipFile(text_input_path)
    xml_data = zip_file.read('word/document.xml')
    zip_file.close()

    # parse the xml data with BeautifulSoup
    return BeautifulSoup(xml_data, 'xml')


# return a list of the value of all dropdown lists in a table
def get_dropdown_list_of_table(text_input_soup, table_index):
    '''

    '''

    list_of_value = []

    # look for all values of dropdown lists in the data and store them
    tables = text_input_soup.find_all('tbl')
    dd_lists_content = tables[table_index].find_all('sdtContent')
    for i in dd_lists_content:
        list_of_value.append(i.find('t').string)

    return list_of_value

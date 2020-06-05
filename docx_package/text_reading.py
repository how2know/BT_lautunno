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

'''
#
def get_parameters_from_tables(text_input_document, text_input_soup, list_of_table, parameters_dictionary):
    """
    Function that reads all parameters stored in the tables of the text input document and stores them in a dictionary.

    Standard tables that define the parameters have two columns, which contain the key and the value of parameters.
    The parameters of these tables are read in the same way for all tables, the only difference is made when the value
    should be an integer or a string.

    Special tables differs from standard tables, e.g. they have more than two columns, contain dropdown lists, etc...
    The parameters of these tables are read are read in a different way for all tables.

    :param text_input_path:
    :param list_of_table:
    :param parameters_dictionary:
    :return:
    """

    # text_input_document = Document(text_input_path)

    # tables with two columns having parameters in each row
    # those tables are handled the same way
    standard_parameters_table = [
        'Report table',
        'Study table',
        'Header table',
        'Approval table',
        'Participants number table',
        'Critical tasks number table',
        'Effectiveness analysis problem number table',
    ]

    # tables that have special features (e.g. more than two columns, dropdown lists, ...)
    # those tables are handled separately
    special_parameters_table = [
        'Critical tasks description table',
        'Effectiveness analysis problem type table'
    ]

    # get the parameters from the standard table
    for table_name in standard_parameters_table:
        table_index = list_of_table.index(table_name)
        table = text_input_document.tables[table_index]

        for row in table.rows:
            key = row.cells[0].text

            if key.startswith('Number of'):
                parameters_dictionary[key] = int(row.cells[1].text)
            else:
                parameters_dictionary[key] = row.cells[1].text

    # get the parameters from the special table
    for table_name in special_parameters_table:

        # get the parameters from the critical tasks description table
        if table_name.startswith('Critical tasks description'):
            table_index = list_of_table.index(table_name)
            table = text_input_document.tables[table_index]

            for i in range(1, parameters_dictionary['Number of critical tasks'] + 1):
                type_key = table.cell(i, 0).text + ' name'
                description_key = table.cell(i, 0).text + ' description'

                parameters_dictionary[type_key] = table.cell(i, 1).text
                parameters_dictionary[description_key] = table.cell(i, 2).text

        # get the parameters from the effectiveness analysis problem type table
        if table_name.startswith('Effectiveness analysis problem'):
            table_index = list_of_table.index(table_name)
            table = text_input_document.tables[table_index]

            """Cells are not recognized as cells by word if they contain a dropdown list. That is why I had to create 
            a work around to get the values here."""

            list_of_text = []     # list of the text of all cells, except those ones containing dropdown list
            stop = False

            number_of_problems = parameters_dictionary['Number of problems']

            # stores the text of every relevant cell in a list
            while not stop:
                for i in range(1, number_of_problems + 1):
                    for cell in table.rows[i].cells:
                        text = cell.text

                        if not text:
                            list_of_text.pop()
                            stop = True
                        else:
                            list_of_text.append(text)

            # delete the last item to ensure that there is no key without a corresponding value
            if len(list_of_text) % 2 != 0:
                list_of_text.pop()

            dropdown_lists_values = get_dropdown_list_of_table(text_input_soup, table_index)

            value = 0

            # stores the keys and corresponding values in the dictionary
            for i in range(len(list_of_text)):
                if i % 2 == 0:
                    problem_number = list_of_text[i]
                    description = list_of_text[i + 1]

                    type_key = problem_number + ' type'
                    description_key = problem_number + ' description'

                    parameters_dictionary[type_key] = dropdown_lists_values[value]
                    parameters_dictionary[description_key] = description

                    value += 1
'''

'''
def read_txt(txt_file_path):
    start_time = 'Start time'
    end_time = 'End time'
    fixation_time = 'Fixation time'
    label = 'Label'

    start_times_list = []
    end_times_list = []
    labels_list = []

    with open(txt_file_path, 'r') as file:
        for line in islice(file, 1, None):
            start_times_list.append(float(line.split()[0]))
            end_times_list.append(float(line.split()[1]))
            labels_list.append(line.split()[2])
    file.close()

    start_times_vector = np.array(start_times_list)
    end_times_vector = np.array(end_times_list)
    fixation_times_vector = end_times_vector - start_times_vector

    data = pd.DataFrame(index=labels_list, columns=[start_time, end_time, fixation_time, label])

    data[start_time] = start_times_list
    data[end_time] = end_times_list
    data[fixation_time] = fixation_times_vector
    data[label] = labels_list

    print(data)
    return data
'''


#####     TEXT READING MODULE     #####

# Functions that read the text from the text input file are implemented in this module.

import os
from zipfile import ZipFile
from bs4 import BeautifulSoup

class chapter_input():

    def __init__(self, heading, parameter_table, picture_table):
        heading = heading
        parameter_table = parameter_table
        picture_table = picture_table


# return the path of a file given its name and the name of its directory
def get_path(file_name, directory_name):
    text_input_directory = os.path.dirname(file_name)
    path = os.path.join(text_input_directory, directory_name, file_name)

    return path


#  find a heading with his title and style and return the corresponding paragraph index
def heading_name_index(paragraphs, title, style):
    for i in range(len(paragraphs)):              # loop over all paragraphs
        if paragraphs[i].style.name == style:     # look for paragraphs with corresponding style
            if paragraphs[i].text == title:       # look for paragraphs with corresponding title
                return i                          # return the index of the paragraphs


# return the index of the next heading corresponding to a style given the index of the previous heading
def next_heading_index(paragraphs, style, previous_index):
    for i in range(previous_index + 1, len(paragraphs)):     # loop over paragraphs coming after the given paragraph index
        if paragraphs[i].style.name == style:                # look for paragraphs with corresponding style
            return i                                         # return the index of the paragraph


# return the index of the next heading corresponding to a style given the index of the previous heading
def next_different_heading_index(paragraphs, style, previous_index):
    for i in range(previous_index + 1, len(paragraphs)):     # loop over paragraphs coming after the given paragraph index
        if paragraphs[i].style.name == style:                # look for paragraphs with corresponding style
            return i                                         # return the index of the paragraph


# store all paragraphs and their corresponding style between a given heading and the next one
def paragraph_after_heading_with_styles(paragraphs, list_of_paragraphs, list_of_styles, heading_title, heading_style):
    heading_index = heading_name_index(paragraphs, heading_title, heading_style)     # index of the given heading
    next_index = next_heading_index(paragraphs, heading_style, heading_index)        # index of the next heading
    for i in range(heading_index + 1, next_index):                                   # loop over all paragraphs between the given heading and the next one
        list_of_paragraphs.append(paragraphs[i])                                     # store all paragraphs in a given list
        list_of_styles.append(paragraphs[i].style.name)                              # store all styles in a given list


# store all paragraphs between a given heading and the next one
def paragraph_after_heading(paragraphs, list_of_paragraphs, heading_title, heading_style):
    heading_index = heading_name_index(paragraphs, heading_title, heading_style)     # index of the given heading
    next_index = next_heading_index(paragraphs, heading_style, heading_index)        # index of the next heading
    for i in range(heading_index + 1, next_index):                                   # loop over all paragraphs between the given heading and the next one
        list_of_paragraphs.append(paragraphs[i])                                     # store all paragraphs in a given list


# store all paragraphs between a given heading and the next one
def paragraph_after_heading_different(paragraphs, list_of_paragraphs, heading_title, heading_style1, heading_style2):
    heading_index = heading_name_index(paragraphs, heading_title, heading_style1)     # index of the given heading
    next_index = next_heading_index(paragraphs, heading_style2, heading_index)        # index of the next heading
    for i in range(heading_index + 1, next_index):                                   # loop over all paragraphs between the given heading and the next one
        list_of_paragraphs.append(paragraphs[i])                                     # store all paragraphs in a given list


# find all terms that you want to define and store them
def find_definitions(table, columns_indexes_list, definitions_list):
    for j in columns_indexes_list:                                     # loop over the columns that contains "Yes" or "No"
        for i in range(len(table.columns[j].cells)):                   # loop over all cells of the columns
            if table.cell(i, j).text == 'Yes':                         # find all cells that contains "Yes"
                definitions_list.append(table.cell(i, j - 1).text)     # store the terms that correspond to a "Yes"


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

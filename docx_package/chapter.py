from docx import Document
from docx.document import Document
from docx.table import Table
from bs4 import BeautifulSoup
from typing import List, Dict, Union
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.shared import Pt, Cm, RGBColor
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
import os
from zipfile import ZipFile
from bs4 import BeautifulSoup
from PIL import Image, UnidentifiedImageError

# from Writing_text import layout
# from Reading_text import text_reading

from docx_package import text_reading, layout


class Chapter:
    """
    Class that represents a chapter.

    Every classic chapter are represented by this class, i.e. all chapters except
    the 'Terms definitions' one and those of the 'Results' section.
    """

    def __init__(self,
                 report_document: Document,
                 text_input_document: Document,
                 text_input_soup: BeautifulSoup,
                 title: str,
                 list_of_tables: List[str],
                 picture_paths_list: List[str],
                 parameters_dictionary: Dict[str, Union[str, int]]
                 ):
        """
        Args:
            report_document: .docx file where the report is written.
            text_input_document: .docx file where all inputs are written.
            text_input_soup: BeautifulSoup of the xml of the input .docx file.
            title: Title of the chapter.
            list_of_tables: List of all table names.
            picture_paths_list: List of the path of all input pictures.
            parameters_dictionary: Dictionary of all input parameters (key = parameter name, value = parameter value)
        """

        self.report = report_document
        self.text_input = text_input_document
        self.text_input_soup = text_input_soup
        self.title = title
        self.list_of_tables = list_of_tables
        self.picture_paths = picture_paths_list
        self.parameters_dictionary = parameters_dictionary

    def heading_index(self) -> int:
        """
        Returns:
            Paragraph index of the chapter heading in the text input document.
        """
        for paragraph_index, paragraph in enumerate(self.text_input.paragraphs):
            if paragraph.text == self.title and 'Heading' in paragraph.style.name:
                return paragraph_index

    def next_heading_index(self) -> int:
        """
        Returns:
            Paragraph index of the following heading in the text input document.
        """

        previous_index = self.heading_index()
        for paragraph_index, paragraph in enumerate(self.text_input.paragraphs[previous_index + 1:]):
            if 'Heading' in paragraph.style.name:
                return paragraph_index + previous_index + 1

    @ property
    def paragraphs(self) -> List[str]:
        """
        Returns:
            List of all paragraphs (as text string) of the chapter.
        """

        list_of_paragraphs = []
        heading_index = self.heading_index()
        next_heading_index = self.next_heading_index()

        # store all paragraphs that are between the two heading indexes in a list
        for paragraph in self.text_input.paragraphs[heading_index + 1: next_heading_index]:
            list_of_paragraphs.append(paragraph.text)

        return list_of_paragraphs

    @ property
    def parameters(self) -> List[str]:
        """
        Returns:
            List of parameters needed to be writen in the chapter.
        """

        return text_reading.get_dropdown_list_of_table(self.text_input_soup,
                                                       self.list_of_tables.index('{} parameter table'.format(self.title))
                                                       )

    @ property
    def picture_name(self) -> str:
        """
        Returns:
            Title with underscores instead of spaces, e.g. 'Use environment' becomes 'Use_environment'.
        """

        return self.title.replace(' ', '_')

    def add_picture(self):
        """
        Load a picture from the input files and add it to the report.



        Returns:
            True if a picture was added, and False if not.
        """

        picture_added = False

        for i in range(1, 4):
            numbered_picture_name = self.picture_name + str(i)

            # find the files that are relevant for the cover page
            for picture_path in self.picture_paths:
                if numbered_picture_name in picture_path:
                    print(numbered_picture_name)
                    print(picture_path)

                    picture = Image.open(picture_path)
                    print(picture.width)
                    print(picture.height)




                    picture_paragraph = self.report.add_paragraph(style='Picture')

                    if picture.width < 378:
                        picture_paragraph.add_run().add_picture(picture_path)
                    else:
                        picture_paragraph.add_run().add_picture(picture_path, width=Cm(10))



    def write_chapter(self):
        """
        Write the heading and the paragraphs of a chapter, including the parameters.
        """

        # read heading style and write heading
        heading_style = self.text_input.paragraphs[self.heading_index()].style.name
        self.report.add_paragraph(self.title, heading_style)

        parameters_values = ['', '', '']

        '''Create variables in order to call property only once, and not in a loop.'''
        parameters = self.parameters
        paragraphs = self.paragraphs

        # stores values of corresponding parameter keys in a list
        for parameter_index, parameter in enumerate(parameters):
            if parameter != '-':
                parameters_values[parameter_index] = self.parameters_dictionary[parameter]

        # write paragraphs including values of parameters
        for paragraph in paragraphs:
            new_paragraph = self.report.add_paragraph(
                paragraph.format(parameters_values[0], parameters_values[1], parameters_values[2],)
            )
            new_paragraph.style.name = 'Normal'

        self.add_picture()

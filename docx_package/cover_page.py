from docx.document import Document
from docx.text.paragraph import Paragraph
from docx.table import Table
from bs4 import BeautifulSoup
from typing import List, Dict, Union
from docx.oxml.ns import nsdecls, qn
from docx.oxml import parse_xml
from docx.oxml.shared import OxmlElement
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER
import numpy as np
from PIL import Image, UnidentifiedImageError

from docx_package import layout


class CoverPage:
    """
    Class that represents and creates the cover page of the report.
    """

    # title and subtitle parameter key
    TITLE_KEY = 'Title'
    SUBTITLE_KEY = 'Subtitle'

    # keys for the approval table
    AUTHOR_NAME_KEY = 'Author’s name'
    AUTHOR_FUNCTION_KEY = 'Author’s function'
    REVIEWER_NAME_KEY = 'Reviewer’s name'
    REVIEWER_FUNCTION_KEY = 'Reviewer’s function'
    APPROVER_NAME_KEY = 'Approver’s name'
    APPROVER_FUNCTION_KEY = 'Approver’s function'

    # hexadecimal of color for cell shading
    LIGHT_GREY_10 = 'D0CECE'

    # picture file name without the extension
    PICTURE_NAME = 'Cover_page'

    def __init__(self, report_document: Document,
                 picture_paths_list: List[str],
                 parameters_dictionary: Dict[str, Union[str, int]]):
        """
        Args:
            report_document: .docx file where the report is written.
            parameters_dictionary: Dictionary of all input parameters (key = parameter name, value = parameter value).
        """

        self.report = report_document
        self.picture_paths = picture_paths_list
        self.parameters = parameters_dictionary

    def write_title(self):
        """
        Create and add a title with a first capital letter and border below it.
        """

        title_text = layout.capitalize_first_letter(self.parameters[self.TITLE_KEY])
        title = self.report.add_paragraph(title_text, 'Title')
        layout.insert_horizontal_border(title)

    def write_subtitle(self) -> Paragraph:
        """
        Create and add a subtitle with a first capital letter.

        Returns:
            Paragraph of the subtitle.
        """

        subtitle_text = layout.capitalize_first_letter(self.parameters[self.SUBTITLE_KEY])
        subtitle = self.report.add_paragraph(subtitle_text, 'Subtitle')
        return subtitle

    def add_approval_table(self):
        """
        Add a table for the approval of the document.
        """

        # name of the different persons
        author_name = self.parameters[self.AUTHOR_NAME_KEY]
        reviewer_name = self.parameters[self.REVIEWER_NAME_KEY]
        approver_name = self.parameters[self.APPROVER_NAME_KEY]

        # store content of cells in a matrix
        approval_cells_text = np.array((['Role', 'Name / Function', 'Date', 'Signature'],
                                        ['Author', author_name, '', ''],
                                        ['Reviewer', reviewer_name, '', ''],
                                        ['Approver', approver_name, '', '']))

        # create table, define its style and fill it
        approval_table = self.report.add_table(rows=4, cols=4)
        layout.define_table_style(approval_table)
        for i in range(0, 4):
            for j in range(0, 4):
                approval_table.cell(i, j).text = approval_cells_text[i, j]

        # set the shading of the first row to light_grey_10 and make it bold
        for cell in approval_table.rows[0].cells:
            layout.set_cell_shading(cell, self.LIGHT_GREY_10)
            cell.paragraphs[0].runs[0].font.bold = True

        # function of the different persons
        author_function = layout.capitalize_first_letter(self.parameters[self.AUTHOR_FUNCTION_KEY])
        reviewer_function = layout.capitalize_first_letter(self.parameters[self.REVIEWER_FUNCTION_KEY].capitalize())
        approver_function = layout.capitalize_first_letter(self.parameters[self.APPROVER_FUNCTION_KEY].capitalize())

        # add the function to the person
        approval_table.cell(1, 1).add_paragraph(author_function)
        approval_table.cell(2, 1).add_paragraph(reviewer_function)
        approval_table.cell(3, 1).add_paragraph(approver_function)

        # make the function italic or do nothing when no function were given
        try:
            for i in range(1, 4):
                approval_table.rows[i].cells[1].paragraphs[1].runs[0].font.italic = True
        except IndexError:
            pass

    # TODO: add function to detect image file
    def add_picture(self) -> bool:
        """
        Load a picture from the input files and add it to the report.

        The longest side (height or width) is set to 14 cm and the ratio is kept.
        The picture is added in the center regarding the side margin and spacing at the top and the bottom
        of the picture is set according to the height.

        Returns:
            True if a picture was added, and False if not.
        """

        picture_added = False

        # find the files that are relevant for the cover page
        for picture_path in self.picture_paths:
            if self.PICTURE_NAME in picture_path:

                # if no picture was added yet, add one
                if not picture_added:
                    try:
                        picture = Image.open(picture_path)

                        # find the longest side and set it to 14 cm when adding the picture
                        # case where the width is the longest side
                        if picture.width >= picture.height:
                            picture_paragraph = self.report.add_paragraph(style='Picture')

                            # set the spacing before and after the picture according the height/width ratio
                            if picture.height / picture.width * 14 < 5:
                                space = 5
                            elif picture.height / picture.width * 14 < 10:
                                space = 3
                            elif picture.height / picture.width * 14 < 14:
                                space = 1
                            picture_paragraph.paragraph_format.space_before = Cm(space)
                            picture_paragraph.paragraph_format.space_after = Cm(space)

                            # add the picture
                            picture_paragraph.add_run().add_picture(picture_path, width=Cm(14))

                        # case where the height is the longest side
                        else:
                            picture_paragraph = self.report.add_paragraph(style='Picture')

                            # spacing before and after the picture is always 1 cm because height is always 14 cm
                            space = 1
                            picture_paragraph.paragraph_format.space_before = Cm(space)
                            picture_paragraph.paragraph_format.space_after = Cm(space)

                            # add the picture
                            picture_paragraph.add_run().add_picture(picture_path, height=Cm(14))

                        picture_added = True

                    # print an error message if a input file is not an image
                    except UnidentifiedImageError:
                        print(picture_path, 'is not an picture file.')

                # if a picture was already added, print an error message if another file is given for the cover page
                else:
                    print('Too many input images for the cover page!')

        return picture_added

    def create(self):
        """
        Create the cover page with a title, a subtitle, a picture and a table for the approval of the report.
        """

        # add title, subtitle, picture (if an image file is provided) and approval table
        self.write_title()
        subtitle = self.write_subtitle()
        picture_added = self.add_picture()
        self.add_approval_table()

        # set the spacing between subtitle and approval table to 16 cm if no picture was added
        if not picture_added:
            subtitle.paragraph_format.space_after = Cm(16)







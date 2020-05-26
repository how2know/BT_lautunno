from docx.oxml.ns import nsdecls, qn
from docx.oxml import parse_xml
from docx.oxml.shared import OxmlElement
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER
import numpy as np
from PIL import Image

from docx_package import layout


# TODO: write this in a module
def capitalize_first_letter(string):
    """
    Capitalizes the first letter of a string to make it look like a title.
    """
    return string[:1].upper() + string[1:]


# insert an horizontal border under a given paragraph
def insert_horizontal_border(paragraph):
    p = paragraph._p                    # p is the <w:p> XML element
    pPr = p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    pPr.insert_element_before(pBdr,
        'w:shd', 'w:tabs', 'w:suppressAutoHyphens', 'w:kinsoku', 'w:wordWrap',
        'w:overflowPunct', 'w:topLinePunct', 'w:autoSpaceDE', 'w:autoSpaceDN',
        'w:bidi', 'w:adjustRightInd', 'w:snapToGrid', 'w:spacing', 'w:ind',
        'w:contextualSpacing', 'w:mirrorIndents', 'w:suppressOverlap', 'w:jc',
        'w:textDirection', 'w:textAlignment', 'w:textboxTightWrap',
        'w:outlineLvl', 'w:divId', 'w:cnfStyle', 'w:rPr', 'w:sectPr',
        'w:pPrChange'
    )
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '6')
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), 'auto')
    pBdr.append(bottom)


class CoverPage:
    """
    This class represents and creates the cover page of the report.
    """

    TITLE_KEY = 'Title'
    SUBTITLE_KEY = 'Subtitle'
    AUTHOR_NAME_KEY = 'Author’s name'
    AUTHOR_FUNCTION_KEY = 'Author’s function'
    REVIEWER_NAME_KEY = 'Reviewer’s name'
    REVIEWER_FUNCTION_KEY = 'Reviewer’s function'
    APPROVER_NAME_KEY = 'Approver’s name'
    APPROVER_FUNCTION_KEY = 'Approver’s function'

    def __init__(self, report_document, parameters_dictionary):
        self.report = report_document
        self.parameters = parameters_dictionary

    def write_title(self):
        title_text = capitalize_first_letter(self.parameters[self.TITLE_KEY])
        title = self.report.add_paragraph(title_text, 'Title')
        insert_horizontal_border(title)

    def write_subtitle(self):
        subtitle_text = capitalize_first_letter(self.parameters[self.SUBTITLE_KEY])
        subtitle = self.report.add_paragraph(subtitle_text, 'Subtitle')

    def add_approval_table(self):

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

        # set the shading of the first row to light_grey_10 (RGB Hex: D0CECE)
        for cell in approval_table.rows[0].cells:
            layout.set_cell_shading(cell, 'D0CECE')

        # make the first row bold
        for col in approval_table.columns:
            col.cells[0].paragraphs[0].runs[0].font.bold = True

        # function of the different persons
        author_function = capitalize_first_letter(self.parameters[self.AUTHOR_FUNCTION_KEY])
        reviewer_function = capitalize_first_letter(self.parameters[self.REVIEWER_FUNCTION_KEY].capitalize())
        approver_function = capitalize_first_letter(self.parameters[self.APPROVER_FUNCTION_KEY].capitalize())

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
    def add_picture(self):
        """
        Load a picture from the input files and add it to the report.
        The longest side (height or width) is set to 14 cm and the ratio is kept.
        The picture is added in the center regarding the side margin and spacing at the top and the bottom
        of the picture is set according to the height.
        Return True if a picture was added, and False if not.
        """

        try:
            picture_path = 'Inputs/Pictures/Cover_page.png'
            # picture_path = 'Inputs/Pictures/Cover_page2.jpg'

            picture = Image.open(picture_path)

            # find the longest side and set it to 14 cm when adding the picture
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

                picture_paragraph.add_run().add_picture(picture_path, width=Cm(14))
            else:
                picture_paragraph = self.report.add_paragraph(style='Picture')

                # spacing before and after the picture is always 1 cm because height is always 14 cm
                space = 1

                picture_paragraph.paragraph_format.space_before = Cm(space)
                picture_paragraph.paragraph_format.space_after = Cm(space)

                picture_paragraph.add_run().add_picture(picture_path, height=Cm(14))
            return True

        except FileNotFoundError:
            return False

    def create(self):
        """
        Create the cover page with a title, a subtitle, a picture and a table for the approval of the report.
        """
        
        self.write_title()
        self.write_subtitle()
        self.add_picture()
        self.add_approval_table()





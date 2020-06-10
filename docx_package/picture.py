from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, Cm, RGBColor


# TODO: add update function
class Picture:
    """
    Class that represents everything that have something to do with the pictures in the report,
    i.e. adding pictures, captions and list of figures.
    """

    # information for the caption
    CAPTION_LABEL = 'Figure '
    CAPTION_STYLE = 'Caption'

    def __init__(self,
                 report_document,
                 picture_paths,
                 picture_name,
                 caption_text,
                 picture_width,
                 picture_height):
        """
        Args:
            report_document: .docx file where the report is written.
            picture_paths: List of paths of all input pictures.
            picture_name: Name of the picture file without the extension.
            caption_text: Text of the picture caption.
            picture_width: Width of the picture in cm as it appears in the report.
            picture_height: Height of the picture in cm as it appears in the report.
        """

        self.report = report_document
        self.picture_paths = picture_paths
        self.picture_name = picture_name
        self.caption = caption_text
        self.width = picture_width
        self.height = picture_height

    def add_picture(self) -> bool:
        """
        Add a picture to the report.

        Returns:
            True if a picture was added, and False if not.
        """

        picture_added = False

        # find the files that correspond to the picture file name
        for picture_path in self.picture_paths:
            if self.picture_name in picture_path:

                # add a picture with the given size in the center of the side margin
                picture_paragraph = self.report.add_paragraph(style='Picture')
                picture_paragraph.add_run().add_picture(picture_path, width=self.width, height=self.height)

                picture_added = True

        return picture_added

    def add_caption(self):
        """
        Add a caption of the form: 'Figure <figure number>: <caption text>, e.g. 'Figure 3: A medical device.'

        The caption will not appear in this form at the first time.
        It has to be updated by pressing Ctrl + A, and then F9.
        """

        # add the label of the caption
        paragraph = self.report.add_paragraph(self.CAPTION_LABEL, style=self.CAPTION_STYLE)

        # add XML elements and set their attributes so that the caption is considered as a caption and can be updated
        run = paragraph.add_run()
        r = run._r
        fldChar = OxmlElement('w:fldChar')
        fldChar.set(qn('w:fldCharType'), 'begin')
        r.append(fldChar)
        instrText = OxmlElement('w:instrText')
        instrText.text = 'SEQ Figure \\* ARABIC'
        r.append(instrText)
        fldChar = OxmlElement('w:fldChar')
        fldChar.set(qn('w:fldCharType'), 'end')
        r.append(fldChar)

        # add the text of the caption
        paragraph.add_run(': {}'.format(self.caption))

    @ classmethod
    def add_picture_and_caption(cls,
                                report_document,
                                picture_paths,
                                picture_name,
                                caption,
                                width=None,
                                height=None,
                                ):
        """
        Add a picture to the report if there is one that corresponds to the picture file name
        and a caption after the picture if one was added.

        Args:
            report_document:
            picture_paths:
            picture_name:
            caption:
            width:
            height:
        """
        picture = cls(report_document, picture_paths, picture_name, caption, width, height)
        picture_added = picture.add_picture()

        if picture_added:
            picture.add_caption()

    @ staticmethod
    def add_figures_list(report_document):
        heading = report_document.add_paragraph('List of figures', 'Heading 2')

        paragraph = report_document.add_paragraph()
        run = paragraph.add_run()
        fldChar = OxmlElement('w:fldChar')  # creates a new element
        fldChar.set(qn('w:fldCharType'), 'begin')  # sets attribute on element
        instrText = OxmlElement('w:instrText')
        instrText.set(qn('xml:space'), 'preserve')  # sets attribute on element
        instrText.text = 'TOC \\h \\z \\c \"Figure\"'  # change 1-3 depending on heading levels you need

        fldChar2 = OxmlElement('w:fldChar')
        fldChar2.set(qn('w:fldCharType'), 'separate')
        fldChar3 = OxmlElement('w:t')
        fldChar3.text = 'Press "Ctrl + A" to select everything and then "F9" to update fields.'
        fldChar2.append(fldChar3)

        fldChar4 = OxmlElement('w:fldChar')
        fldChar4.set(qn('w:fldCharType'), 'end')

        r_element = run._r
        r_element.append(fldChar)
        r_element.append(instrText)
        r_element.append(fldChar2)
        r_element.append(fldChar4)
        p_element = paragraph._p

# TODO: comment
#####     LAYOUT DEFINITION     #####

# Some functions that define the layout and formatting of the report are implemented in this module.
# It includes: document layout, header, footer, styles, tables, ...
from docx.document import Document
from docx.section import Section
from docx.text.paragraph import Paragraph
from docx.table import _Row, _Column, _Cell
from docx.enum.base import EnumValue
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.enum.section import WD_SECTION, WD_ORIENT
from docx.shared import Pt, Cm, RGBColor
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT, WD_ROW_HEIGHT_RULE
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import nsdecls, qn
from docx.oxml import parse_xml
from docx.oxml.shared import OxmlElement


class Layout:
    """
    Class that represents and defines the layout and formatting of the report.
    """

    # TODO: use this
    # define some colors
    BLACK = RGBColor(0, 0, 0)  # Hex: 000000
    BLACK_35 = RGBColor(90, 90, 90)  # Hex: 5A5A5A
    LIGHT_GREY_10 = RGBColor(208, 206, 206)  # Hex: D0CECE

    def __init__(self, report_document):
        """
        Args:
            report_document: .docx file where the report is written.
        """
        self.report = report_document

    @ staticmethod
    def define_page_format(section: Section):
        """
        Define the page setup of a section as default A4 setup (21 cm x 29.7 cm) with 2.5 cm margin.

        Args:
            section: Section whose setup will be defined.
        """

        section.orientation = WD_ORIENT.PORTRAIT
        section.page_width = Cm(21)
        section.page_height = Cm(29.7)
        section.top_margin = Cm(2.5)
        section.bottom_margin = Cm(2.5)
        section.right_margin = Cm(2.5)
        section.left_margin = Cm(2.5)

    def define_style(self,
                     name: str,
                     font: str,
                     size: int,
                     color: RGBColor,
                     alignment: EnumValue,
                     italic=False,
                     bold=False,
                     space_before=None,
                     space_after=None
                     ):
        """
        Define the characteristics of a style in the report.

        Args:
            name: Name of the style that is defined.
            font: Font name
            size: Font size in Pt.
            color: Font color
            alignment: Alignment of the paragraph (left, right, center, justify).
            italic (optional): Boolean to know if it should be italic.
            bold (optional): Boolean to know if it should be bold.
            space_before (optional): Space before the paragraph. None if inherited from the style hierarchy.
            space_after (optional): Space after the paragraph. None if inherited from the style hierarchy.
        """

        # add a new style if it does not already exist
        if name not in self.report.styles:
            self.report.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)

        style = self.report.styles[name]

        # this let us modify the font of some styles that are kind of "blocked"
        try:
            style.element.xpath('w:rPr/w:rFonts')[0].attrib.clear()
        except IndexError:
            pass

        style.font.name = font
        style.font.size = Pt(size)
        style.font.color.rgb = color
        style.paragraph_format.alignment = alignment
        style.font.italic = italic
        style.font.bold = bold
        style.paragraph_format.space_before = space_before
        style.paragraph_format.space_after = space_after

    @ classmethod
    def define_all_styles(cls, report_document: Document):
        """
        Define all relevant styles of the report.

        Args:
            report_document: .docx file where the report is written.
        """

        layout = cls(report_document)

        layout.define_style('Title', 'Calibri Light', 32, layout.BLACK, WD_ALIGN_PARAGRAPH.CENTER)
        layout.define_style('Subtitle', 'Calibri Light', 24, layout.BLACK_35, WD_ALIGN_PARAGRAPH.CENTER)
        layout.define_style('Heading 1', 'Calibri', 16, layout.BLACK, WD_ALIGN_PARAGRAPH.LEFT, bold=True,
                            space_before=Pt(12))
        layout.define_style('Heading 2', 'Calibri Light', 14, layout.BLACK, WD_ALIGN_PARAGRAPH.LEFT, bold=True,
                            space_before=Pt(2))
        layout.define_style('Heading 3', 'Calibri Light', 12, layout.BLACK, WD_ALIGN_PARAGRAPH.LEFT, bold=True,
                            space_before=Pt(2))
        layout.define_style('Normal', 'Calibri', 11, layout.BLACK, WD_ALIGN_PARAGRAPH.JUSTIFY,
                            space_before=Pt(0), space_after=Pt(0))
        # layout.define_style('Table', 'Calibri', 11, layout.BLACK, WD_ALIGN_PARAGRAPH.LEFT)
        layout.define_style('Picture', 'Calibri', 11, layout.BLACK, WD_ALIGN_PARAGRAPH.CENTER,
                            space_before=Pt(8), space_after=Pt(5))
        layout.define_style('Caption', 'Calibri', 9, layout.BLACK, WD_ALIGN_PARAGRAPH.CENTER, bold=True)

    @ staticmethod
    def capitalize_first_letter(string: str) -> str:
        """
        Args:
            string: String whose first letter must be capitalized.

        Returns:
            String with a capital first letter.
        """

        return string[:1].upper() + string[1:]

    @ staticmethod
    def set_cell_shading(cell: _Cell, color_hex: str):
        """
        Color a cell.

        Args:
            cell: Cell that must be colored.
            color_hex: Hexadecimal representation of the color.
        """

        shading_elm = parse_xml(r'<w:shd {0} w:fill="{1}"/>'.format(nsdecls('w'), color_hex))
        cell._tc.get_or_add_tcPr().append(shading_elm)

    @ staticmethod
    def insert_horizontal_border(paragraph: Paragraph):
        """
        Add an horizontal border under a paragraph.

        Args:
            paragraph: Paragraph under which you want to add an horizontal border.
        """

        # access to XML paragraph element <w:p> and its properties
        p = paragraph._p
        pPr = p.get_or_add_pPr()

        # create new XML element and insert it to the paragraph element
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

        # create new XML element, set its properties and add it to the pBdr element
        bottom = OxmlElement('w:bottom')
        bottom.set(qn('w:val'), 'single')
        bottom.set(qn('w:sz'), '6')
        bottom.set(qn('w:space'), '1')
        bottom.set(qn('w:color'), 'auto')
        pBdr.append(bottom)

    @ staticmethod
    def set_row_height(row: _Row, height: float, rule=WD_ROW_HEIGHT_RULE.EXACTLY):
        """
        Set the height of a table row.

        Args:
            row: Row whose height is to be changed.
            height: Height of the column in cm.
            rule (optional): Rule for determining the height of a table row, e.g. rule=WD_ROW_HEIGHT_RULE.AT_LEAST.
        """

        row.height_rule = rule
        row.height = Cm(height)

    @ staticmethod
    def set_column_width(column: _Column, width: float):
        """
        Set the width of a table column.

        Note:
            To make it work, the autofit of the corresponding table must be disabled beforehand (table.autofit = False).

        Args:
            column: Column whose width is to be changed.
            width: Width of the column in cm.
        """

        for cell in column.cells:
            cell.width = Cm(width)

    @ staticmethod
    def set_cell_border(cell: _Cell, **kwargs):
        """
        Set the border of a cell.

        Usage example:
        set_cell_border(cell,
                        top={"sz": 12, "val": "single", "color": "#FF0000", "space": "0"},     # top border
                        bottom={"sz": 12, "color": "#00FF00", "val": "single"},     # bottom border
                        start={"sz": 24, "val": "dashed", "shadow": "true"},     # left border
                        end={"sz": 12, "val": "dashed"}     # right border
                        )

        Available attributes can be found here: http://officeopenxml.com/WPtableBorders.php

        Args:
            cell: Cell with borders to be changed.
        """

        # access to XML element <w:tc> and its properties
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()

        # check for tag existence, if none found, then create one
        tcBorders = tcPr.first_child_found_in("w:tcBorders")
        if tcBorders is None:
            tcBorders = OxmlElement('w:tcBorders')
            tcPr.append(tcBorders)

        # list over all available tags
        for edge in ('start', 'top', 'end', 'bottom', 'insideH', 'insideV'):
            edge_data = kwargs.get(edge)
            if edge_data:
                tag = 'w:{}'.format(edge)

                # check for tag existence, if none found, then create one
                element = tcBorders.find(qn(tag))
                if element is None:
                    element = OxmlElement(tag)
                    tcBorders.append(element)

                # looks like order of attributes is important
                for key in ["sz", "val", "color", "space", "shadow"]:
                    if key in edge_data:
                        element.set(qn('w:{}'.format(key)), str(edge_data[key]))

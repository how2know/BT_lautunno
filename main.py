from docx import Document
from docx.enum.section import WD_SECTION
import os
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import win32com.client
import inspect

import time

from docx_package import text_reading, layout, picture

from docx_package.layout import Layout
from docx_package.chapter import Chapter
from docx_package.effectiveness_analysis import EffectivenessAnalysis
from docx_package.definitions import Definitions
from docx_package.header_footer import Header, Footer
from docx_package.table_of_content import TableOfContent
from docx_package.time_on_tasks import TimeOnTasks
from docx_package.cover_page import CoverPage
from docx_package.dwell_times_revisits import DwellTimesAndRevisits
from docx_package.average_fixation import AverageFixation
from docx_package.transitions import Transitions
from docx_package.parameters import Parameters
from docx_package.picture import Picture

from eye_tracking_package import cGOM_data
from eye_tracking_package.tobii_data import TobiiData


def list_number(doc, par, prev=None, level=None, num=True):
    """
    Makes a paragraph into a list item with a specific level and
    optional restart.

    An attempt will be made to retreive an abstract numbering style that
    corresponds to the style of the paragraph. If that is not possible,
    the default numbering or bullet style will be used based on the
    ``num`` parameter.

    Parameters
    ----------
    doc : docx.document.Document
        The document to add the list into.
    par : docx.paragraph.Paragraph
        The paragraph to turn into a list item.
    prev : docx.paragraph.Paragraph or None
        The previous paragraph in the list. If specified, the numbering
        and styles will be taken as a continuation of this paragraph.
        If omitted, a new numbering scheme will be started.
    level : int or None
        The level of the paragraph within the outline. If ``prev`` is
        set, defaults to the same level as in ``prev``. Otherwise,
        defaults to zero.
    num : bool
        If ``prev`` is :py:obj:`None` and the style of the paragraph
        does not correspond to an existing numbering style, this will
        determine wether or not the list will be numbered or bulleted.
        The result is not guaranteed, but is fairly safe for most Word
        templates.
    """
    xpath_options = {
        True: {'single': 'count(w:lvl)=1 and ', 'level': 0},
        False: {'single': '', 'level': level},
    }

    def style_xpath(prefer_single=True):
        """
        The style comes from the outer-scope variable ``par.style.name``.
        """
        style = par.style.style_id
        return (
            'w:abstractNum['
                '{single}w:lvl[@w:ilvl="{level}"]/w:pStyle[@w:val="{style}"]'
            ']/@w:abstractNumId'
        ).format(style=style, **xpath_options[prefer_single])

    def type_xpath(prefer_single=True):
        """
        The type is from the outer-scope variable ``num``.
        """
        type = 'decimal' if num else 'bullet'
        return (
            'w:abstractNum['
                '{single}w:lvl[@w:ilvl="{level}"]/w:numFmt[@w:val="{type}"]'
            ']/@w:abstractNumId'
        ).format(type=type, **xpath_options[prefer_single])

    def get_abstract_id():
        """
        Select as follows:

            1. Match single-level by style (get min ID)
            2. Match exact style and level (get min ID)
            3. Match single-level decimal/bullet types (get min ID)
            4. Match decimal/bullet in requested level (get min ID)
            3. 0
        """
        for fn in (style_xpath, type_xpath):
            for prefer_single in (True, False):
                xpath = fn(prefer_single)
                ids = numbering.xpath(xpath)
                if ids:
                    return min(int(x) for x in ids)
        return 0

    if (prev is None or
            prev._p.pPr is None or
            prev._p.pPr.numPr is None or
            prev._p.pPr.numPr.numId is None):
        if level is None:
            level = 0
        numbering = doc.part.numbering_part.numbering_definitions._numbering
        # Compute the abstract ID first by style, then by num
        anum = get_abstract_id()
        # Set the concrete numbering based on the abstract numbering ID
        num = numbering.add_num(anum)
        # Make sure to override the abstract continuation property
        num.add_lvlOverride(ilvl=level).add_startOverride(1)
        # Extract the newly-allocated concrete numbering ID
        num = num.numId
    else:
        if level is None:
            level = prev._p.pPr.numPr.ilvl.val
        # Get the previous concrete numbering ID
        num = prev._p.pPr.numPr.numId.val
    par._p.get_or_add_pPr().get_or_add_numPr().get_or_add_numId().val = num
    par._p.get_or_add_pPr().get_or_add_numPr().get_or_add_ilvl().val = level


def main():
    start = time.time()

    # file name of the report
    report_file = 'Report.docx'

    # create document
    report = Document()

    # name of the directory and the input files
    input_directory = 'Inputs'
    text_input_file = 'Text_input.docx'
    definitions_file = 'Terms_definitions.docx'

    # path of the input files
    text_input_path = text_reading.get_path(text_input_file, input_directory)
    definitions_path = text_reading.get_path(definitions_file, input_directory)

    # load text input files with python-docx
    text_input = Document(text_input_path)
    definitions = Document(definitions_path)

    picture_paths = Picture.get_picture_paths()

    #
    text_input_soup = text_reading.parse_xml_with_bs4(text_input_path)

    # list of all tables in the order they appear in the text input document
    # useful to get the index of a table in the document
    tables = [
        'Report table',
        'Study table',
        'Header table',
        'Approval table',
        'Cover page table',
        'Purpose parameter table',
        'Purpose caption table',
        'Background parameter table',
        'Background caption table',
        'Scope parameter table',
        'Scope caption table',
        'EU Regulation 2017/745 definitions table',
        'IEC 62366-1 definitions table',
        'FDA Guidance definitions table',
        'Ethics statement parameter table',
        'Ethics statement caption table',
        'Device specifications parameter table',
        'Device specifications caption table',
        'Goal parameter table',
        'Goal caption table',
        'Participants number table',
        'Participants description table',
        'Participants parameter table',
        'Participants caption table',
        'Use environment parameter table',
        'Use environment caption table',
        'Critical tasks number table',
        'Critical tasks description table',
        'Use scenarios parameter table',
        'Use scenarios caption table',
        'Setup parameter table',
        'Setup caption table',
        'Effectiveness analysis tasks and problems table',
        'Effectiveness analysis problem number table',
        'Effectiveness analysis problem type table',
        'Effectiveness analysis video table',
        'Effectiveness analysis parameter table',
        'Effectiveness analysis caption table',
        'Time on tasks table',
        'Time on tasks plot type table',
        'Time on tasks parameter table',
        'Time on tasks caption table',
        'Dwell times and revisits parameter table',
        'Dwell times and revisits caption table',
        'Average fixation plot type table',
        'Average fixation parameter table',
        'Average fixation caption table',
        'Transitions parameter table',
        'Transitions caption table',
        'Conclusion parameter table',
        'Conclusion caption table',
    ]

    parameters = Parameters.get_all(text_input, text_input_soup, tables)

    print(parameters)   # TODO: delete this line

    cGOM_dataframes = cGOM_data.make_dataframes_list(parameters)

    '''tobii_data = TobiiData.make_dataframe('Inputs/Data/Participant1.tsv')'''

    # define all styles used in the document
    # layout.define_all_styles(report)
    Layout.define_all_styles(report)

    # section1 includes cover page and table of content
    section1 = report.sections[0]
    layout.define_page_format(section1)

    cover_page = CoverPage(report, text_input, tables, picture_paths, parameters)
    cover_page.create()

    # add a page break
    report.add_page_break()

    # TODO: write this better (do not set it as heading 1)
    # add table of content
    report.add_paragraph('Table of content', 'Heading 1')

    ''' commented out to save time
    # TODO: write this better (function in class TableOfContent)
    # path to the file needed to update the table of content
    script_dir = os.path.dirname(os.path.abspath(inspect.getfile(inspect.currentframe())))
    report_file_path = os.path.join(script_dir, report_file)
    '''

    table_of_content = TableOfContent(report)
    table_of_content.write()

    # add a new section
    section2 = report.add_section(WD_SECTION.NEW_PAGE)
    layout.define_page_format(section2)

    # add a footer with page number
    footer = Footer(section2)
    footer.write()

    # add a header
    header = Header(section2, parameters)
    header.write()



    p0 = report.add_paragraph('Item 1', style='List Bullet')
    list_number(report, p0, level=0, num=False)
    p1 = report.add_paragraph('Item A', style='List Bullet 2')
    list_number(report, p1, p0, level=1)
    p2 = report.add_paragraph('Item 2', style='List Bullet')
    list_number(report, p2, p1, level=0)
    p3 = report.add_paragraph('Item B', style='List Bullet 2')
    list_number(report, p3, p2, level=1)




    '''  create and write all the chapters '''

    purpose = Chapter(report, text_input, text_input_soup, 'Purpose', tables, picture_paths, parameters)
    purpose.write_chapter()

    background = Chapter(report, text_input, text_input_soup, 'Background', tables, picture_paths, parameters)
    background.write_chapter()

    scope = Chapter(report, text_input, text_input_soup, 'Scope', tables, picture_paths, parameters)
    scope.write_chapter()

    def_chapter = Definitions(report, text_input, text_input_soup, definitions, 'Terms definitions', tables)
    def_chapter.write_all_definitions()

    ethics = Chapter(report, text_input, text_input_soup, 'Ethics statement', tables, picture_paths, parameters)
    ethics.write_chapter()

    device = Chapter(report, text_input, text_input_soup, 'Device specifications', tables, picture_paths, parameters)
    device.write_chapter()

    # TODO: write this better
    report.add_paragraph('Test procedure', 'Heading 2')

    goal = Chapter(report, text_input, text_input_soup, 'Goal', tables, picture_paths, parameters)
    goal.write_chapter()

    participants = Chapter(report, text_input, text_input_soup, 'Participants', tables, picture_paths, parameters)
    participants.write_chapter()

    environment = Chapter(report, text_input, text_input_soup, 'Use environment', tables, picture_paths, parameters)
    environment.write_chapter()

    scenarios = Chapter(report, text_input, text_input_soup, 'Use scenarios', tables, picture_paths, parameters)
    scenarios.write_chapter()

    setup = Chapter(report, text_input, text_input_soup, 'Setup', tables, picture_paths, parameters)
    setup.write_chapter()


    # TODO: write this better
    report.add_paragraph('Results', 'Heading 1')

    start1 = time.time()
    effectiveness_analysis = EffectivenessAnalysis(report, text_input, text_input_soup, tables, parameters)
    effectiveness_analysis.write_chapter()
    end1 = time.time()
    print('Effectiveness analysis: ', end1-start1)

    start2 = time.time()
    time_on_tasks = TimeOnTasks(report, text_input, text_input_soup, tables, parameters)
    time_on_tasks.write_chapter()
    end2 = time.time()
    print('Time on tasks: ', end2-start2)

    start3 = time.time()
    dwell_times_and_revisits = DwellTimesAndRevisits(report, text_input, text_input_soup, tables, parameters, cGOM_dataframes)
    dwell_times_and_revisits.write_chapter()
    end3 = time.time()
    print('Dwell times: ', end3-start3)
    
    start4 = time.time()
    average_fixation = AverageFixation(report, text_input, text_input_soup, tables, parameters, cGOM_dataframes)
    average_fixation.write_chapter()
    end4 = time.time()
    print('Average fixation: ', end4-start4)

    start5 = time.time()
    transitions = Transitions(report, text_input, text_input_soup, tables, parameters, cGOM_dataframes)
    transitions.write_chapter()
    end5 = time.time()
    print('Transitions: ', end5-start5)

    conclusion = Chapter(report, text_input, text_input_soup, 'Conclusion', tables, picture_paths, parameters)
    conclusion.write_chapter()

    '''Picture.error_message(picture_paths)'''

    # add a page break
    report.add_page_break()

    Picture.add_figures_list(report)

    # save the report
    report.save(report_file)

    '''This works but it is very long...'''
    # update the table of content
    # table_of_content.update(report_file_path)

    # open the report with the default handler for .docx (Word)
    os.startfile(report_file)

    end = time.time()

    print(end - start)


if __name__ == '__main__':
    main()

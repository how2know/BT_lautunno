from docx import Document
from docx.enum.section import WD_SECTION
import os
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import win32com.client
import inspect

import time

from docx_package import text_reading, layout, picture

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

from txt_package import cGOM_data


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

    # define all styles used in the document
    layout.define_all_styles(report)

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

    '''
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
    '''

    start3 = time.time()
    dwell_times_and_revisits = DwellTimesAndRevisits(report, text_input, text_input_soup, tables, parameters, cGOM_dataframes)
    dwell_times_and_revisits.write_chapter()
    end3 = time.time()
    print('Dwell times: ', end3-start3)

    '''
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
    '''

    conclusion = Chapter(report, text_input, text_input_soup, 'Conclusion', tables, picture_paths, parameters)
    conclusion.write_chapter()

    Picture.error_message(picture_paths)

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

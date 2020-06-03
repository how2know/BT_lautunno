from docx import Document
from docx.enum.section import WD_SECTION
import os
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import win32com.client
import inspect

import time

from docx_package import text_reading, layout

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

    txt_file_path = 'Inputs/Data/Participant1.txt'

    txt_file_data = text_reading.read_txt(txt_file_path)


    #
    text_input_soup = text_reading.parse_xml_with_bs4(text_input_path)

    # list of all tables in the order they appear in the text input document
    # useful to get the index of a table in the document
    tables = [
        'Report table',
        'Study table',
        'Header table',
        'Approval table',
        'Purpose parameter table',
        'Background parameter table',
        'Scope parameter table',
        'EU Regulation 2017/745 definitions table',
        'IEC 62366-1 definitions table',
        'FDA Guidance definitions table',
        'Ethics statement parameter table',
        'Device specifications parameter table',
        'Goal parameter table',
        'Participants number table',
        'Participants description table',
        'Participants parameter table',
        'Use environment parameter table',
        'Critical tasks number table',
        'Critical tasks description table',
        'Use scenarios parameter table',
        'Setup parameter table',
        'Effectiveness analysis tasks and problems table',
        'Effectiveness analysis problem number table',
        'Effectiveness analysis problem type table',
        'Effectiveness analysis video table',
        'Effectiveness analysis parameter table',
        'Time on tasks table',
        'Time on tasks parameter table',
        'Dwell times and revisits parameter table',
        'Average fixation parameter table',
        'Transitions parameter table',
        'Conclusion parameter table',
    ]

    parameters = {}

    text_reading.get_parameters_from_tables(text_input, text_input_soup, tables, parameters)

    print(parameters)

    cGOM_dataframes = cGOM_data.make_dataframes_list(parameters)

    # define all styles used in the document
    layout.define_all_styles(report)

    # section1 includes cover page and table of content
    section1 = report.sections[0]
    layout.define_page_format(section1)

    cover_page = CoverPage(report, parameters)
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

    purpose = Chapter(report, text_input, text_input_soup, 'Purpose', tables, parameters)
    purpose.write_chapter()

    background = Chapter(report, text_input, text_input_soup, 'Background', tables, parameters)
    background.write_chapter()

    scope = Chapter(report, text_input, text_input_soup, 'Scope', tables, parameters)
    scope.write_chapter()

    def_chapter = Definitions(report, text_input, text_input_soup, definitions, 'Terms definitions', tables)
    def_chapter.write_all_definitions()

    ethics = Chapter(report, text_input, text_input_soup, 'Ethics statement', tables, parameters)
    ethics.write_chapter()

    device = Chapter(report, text_input, text_input_soup, 'Device specifications', tables, parameters)
    device.write_chapter()

    # TODO: write this better
    report.add_paragraph('Test procedure', 'Heading 2')

    goal = Chapter(report, text_input, text_input_soup, 'Goal', tables, parameters)
    goal.write_chapter()

    participants = Chapter(report, text_input, text_input_soup, 'Participants', tables, parameters)
    participants.write_chapter()

    environment = Chapter(report, text_input, text_input_soup, 'Use environment', tables, parameters)
    environment.write_chapter()

    scenarios = Chapter(report, text_input, text_input_soup, 'Use scenarios', tables, parameters)
    scenarios.write_chapter()

    setup = Chapter(report, text_input, text_input_soup, 'Setup', tables, parameters)
    setup.write_chapter()

    # TODO: write this better
    report.add_paragraph('Results', 'Heading 1')

    effectiveness_analysis = EffectivenessAnalysis(report, text_input, text_input_soup, tables, parameters)
    effectiveness_analysis.write_chapter()

    # average_fixation = AverageFixation(report, text_input, text_input_soup, tables, parameters, txt_file_data, cGOM_dataframes)
    # average_fixation.write_chapter()

    time_on_tasks = TimeOnTasks(report, text_input, text_input_soup, tables, parameters)
    time_on_tasks.write_chapter()

    dwell_times_and_revisits = DwellTimesAndRevisits(report, text_input, text_input_soup, tables, parameters, txt_file_data, cGOM_dataframes)
    dwell_times_and_revisits.write_chapter()

    average_fixation = AverageFixation(report, text_input, text_input_soup, tables, parameters, txt_file_data, cGOM_dataframes)
    average_fixation.write_chapter()

    transitions = Transitions(report, text_input, text_input_soup, tables, parameters)
    transitions.write_chapter()

    conclusion = Chapter(report, text_input, text_input_soup, 'Conclusion', tables, parameters)
    conclusion.write_chapter()

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

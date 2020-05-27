from docx import Document
from docx.enum.section import WD_SECTION
import os

import time

# from classes import *
from docx_package.chapter import *
from docx_package.effectiveness_analysis import *
from docx_package.definitions import *
from docx_package.header_footer import *
from docx_package.table_of_content import *
from docx_package.time_on_tasks import TimeOnTasks
from docx_package.cover_page import CoverPage


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

    #
    text_input_soup = text_reading.parse_xml_with_bs4(text_input_path)

    # list of all tables in text input document
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
        'Dwell time and revisits parameter table',
        'Average fixation parameter table',
        'Transitions parameter table',
        'Conclusion parameter table',
    ]

    parameters = {}

    text_reading.get_parameters_from_tables(text_input, text_input_soup, tables, parameters)

    print(parameters)

    # define all styles used in the document
    layout.define_all_styles(report)

    # section1 includes cover page and table of content
    section1 = report.sections[0]
    layout.define_page_format(section1)

    cover_page = CoverPage(report, parameters)
    cover_page.create()

    # add a page break
    report.add_page_break()

    # add table of content
    report.add_heading(text.toc_title, 1)

    # TODO: write this better (function in class TableOfContent)
    script_dir = os.path.dirname(os.path.abspath(inspect.getfile(inspect.currentframe())))
    report_file_path = os.path.join(script_dir, report_file)

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

    # text_writing.write_definitions_chapter(report, text.definitions_title, text.defined_terms, text.definitions_list, text.definitions_styles_list)

    def_chapter = Definitions(report, text_input, text_input_soup, definitions, 'Terms definitions', tables)
    def_chapter.write_all_definitions()

    ethics = Chapter(report, text_input, text_input_soup, 'Ethics statement', tables, parameters)
    ethics.write_chapter()

    device = Chapter(report, text_input, text_input_soup, 'Device specifications', tables, parameters)
    device.write_chapter()

    report.add_heading(text.procedure_title, 1)

    goal = Chapter(report, text_input, text_input_soup, 'Goal', tables, parameters)
    goal.write_chapter()

    # text_writing.write_chapter(report, text.goal_title, 2, text.goal_paragraphs)

    participants = Chapter(report, text_input, text_input_soup, 'Participants', tables, parameters)
    participants.write_chapter()

    environment = Chapter(report, text_input, text_input_soup, 'Use environment', tables, parameters)
    environment.write_chapter()

    scenarios = Chapter(report, text_input, text_input_soup, 'Use scenarios', tables, parameters)
    scenarios.write_chapter()

    setup = Chapter(report, text_input, text_input_soup, 'Setup', tables, parameters)
    setup.write_chapter()

    report.add_heading(text.results_title, 1)

    effectiveness_analysis = EffectivenessAnalysis(report, text_input, text_input_soup, 'Effectiveness analysis', tables, parameters)
    effectiveness_analysis.write_chapter()

    time_on_tasks = TimeOnTasks(report, text_input, tables, parameters)
    time_on_tasks.insert_plot()


    # text_writing.write_chapter(report, text.results_title, 1, text.results_paragraphs)

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

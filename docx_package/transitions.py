from docx_package.results import ResultsChapter


class Transitions:

    TITLE = 'Transitions'
    TITLE_STYLE = 'Heading 2'
    ANALYSIS_TITLE = 'Analysis'
    ANALYSIS_STYLE = 'Heading 3'

    def __init__(self, report_document, text_input_document, text_input_soup, list_of_tables, parameters_dictionary):
        self.report = report_document
        self.text_input = text_input_document
        self.text_input_soup = text_input_soup
        self.tables = list_of_tables
        self.parameters = parameters_dictionary

    def write_chapter(self):
        time_on_tasks = ResultsChapter(self.report, self.text_input, self.text_input_soup, self.TITLE,
                                       self.tables, self.parameters)

        self.report.add_paragraph(self.TITLE, self.TITLE_STYLE)

        self.report.add_paragraph(self.ANALYSIS_TITLE, self.ANALYSIS_STYLE)
        time_on_tasks.write_chapter()
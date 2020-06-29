from docx.document import Document
from bs4 import BeautifulSoup
from typing import List

from docx_package import text_reading


class Parameters:
    """
    Class that represents and defines the parameter used to write the chapters of the report.
    """

    # tables with two columns having parameters in each row
    # all those tables are handled the same way
    STANDARD_PARAMETERS_TABLES = [
        'Report table',
        'Study table',
        'Header table',
        'Approval table',
        'Participants number table',
        'Critical tasks number table',
        'Effectiveness analysis problem number table',
    ]

    # tables that have special features (e.g. more than two columns, dropdown lists, ...)
    # those tables are handled separately
    TASKS_TABLE = 'Critical tasks description table'
    EFFECTIVENESS_TABLE = 'Effectiveness analysis problem type table'

    def __init__(self,
                 text_input_document: Document,
                 text_input_soup: BeautifulSoup,
                 list_of_tables: List[str]
                 ):
        """
        Args:
            text_input_document: .docx file where all inputs are written.
            text_input_soup: BeautifulSoup of the xml of the input .docx file.
            list_of_tables: List of all table names.
        """

        self.text_input = text_input_document
        self.text_input_soup = text_input_soup
        self.tables = list_of_tables

        # dictionary where keys and values of all parameters will be stored
        self.dictionary = {}

    def get_from_standard_tables(self):
        """
        Read the parameters from the standard tables and stored them in the dictionary.

        Standard tables have two columns, the first one contains the keys and the second one the values of parameters.
        """

        for table_name in self.STANDARD_PARAMETERS_TABLES:
            table_index = self.tables.index(table_name)
            table = self.text_input.tables[table_index]

            for row in table.rows:
                key = row.cells[0].text

                # a difference is made when the value should be an integer or a string
                if key.startswith('Number of'):
                    self.dictionary[key] = int(row.cells[1].text)
                else:
                    self.dictionary[key] = row.cells[1].text

    def get_from_tasks_table(self):
        """
        Read the parameters from the critical tasks table and stored them in the dictionary.

        The critical tasks table differs from the standard table because it has 3 columns.
        """

        table_index = self.tables.index(self.TASKS_TABLE)
        table = self.text_input.tables[table_index]

        for i in range(1, self.dictionary['Number of critical tasks'] + 1):
            type_key = table.cell(i, 0).text + ' name'
            description_key = table.cell(i, 0).text + ' description'

            self.dictionary[type_key] = table.cell(i, 1).text
            self.dictionary[description_key] = table.cell(i, 2).text

    def get_from_problems_table(self):
        """
        Read the parameters from the problems table and stored them in the dictionary.

        The problems table differs from the standard table because it has 3 columns and contains dropdown lists.
        """

        table_index = self.tables.index(self.EFFECTIVENESS_TABLE)
        table = self.text_input.tables[table_index]

        number_of_problems = self.dictionary['Number of problems']

        # cells are not recognized as cells by word if they contain a dropdown list,
        # so here is a work around to get the values

        # list of the text of all cells, except those ones containing dropdown list
        list_of_text = []

        # stores the text of every relevant cell in the list
        stop = False
        while not stop:
            for i in range(1, number_of_problems + 1):
                for cell in table.rows[i].cells:
                    text = cell.text
                    if not text:
                        list_of_text.pop()
                        stop = True
                    else:
                        list_of_text.append(text)

        # delete the last item to ensure that there is no key without a corresponding value
        if len(list_of_text) % 2 != 0:
            list_of_text.pop()

        dropdown_lists_values = text_reading.get_dropdown_list_of_table(self.text_input_soup, table_index)

        value = 0

        # stores the keys and corresponding values in the dictionary
        for i in range(len(list_of_text)):

            # the problem number appears every two items
            if i % 2 == 0:

                # problem types
                problem_number = list_of_text[i]
                type_key = problem_number + ' type'
                self.dictionary[type_key] = dropdown_lists_values[value]

                # problem descriptions
                description = list_of_text[i + 1]
                description_key = problem_number + ' description'
                self.dictionary[description_key] = description

                value += 1

    @ classmethod
    def get_all(cls, text_input_document, text_input_soup, list_of_tables):
        """
        Args:
            text_input_document: .docx file where all inputs are written.
            text_input_soup: BeautifulSoup of the xml of the input .docx file.
            list_of_tables: List of all table names.

        Returns:
            Dictionary containing values and keys of all parameters.
        """

        parameters = cls(text_input_document, text_input_soup, list_of_tables)

        parameters.get_from_standard_tables()
        parameters.get_from_tasks_table()
        parameters.get_from_problems_table()

        return parameters.dictionary

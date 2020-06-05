from docx_package import text_reading


class Parameters:
    """
            Function that reads all parameters stored in the tables of the text input document and stores them in a dictionary.

        Standard tables that define the parameters have two columns, which contain the key and the value of parameters.
        The parameters of these tables are read in the same way for all tables, the only difference is made when the value
        should be an integer or a string.

        Special tables differs from standard tables, e.g. they have more than two columns, contain dropdown lists, etc...
        The parameters of these tables are read are read in a different way for all tables.
    """

    # tables with two columns having parameters in each row
    # those tables are handled the same way
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

    def __init__(self, text_input_document, text_input_soup, list_of_table):
        self.text_input = text_input_document
        self.text_input_soup = text_input_soup
        self.tables = list_of_table
        self.dictionary = {}

    def get_from_standard_tables(self):
        """
        Get the parameters from the standard tables
        """

        for table_name in self.STANDARD_PARAMETERS_TABLES:
            table_index = self.tables.index(table_name)
            table = self.text_input.tables[table_index]

            for row in table.rows:
                key = row.cells[0].text

                if key.startswith('Number of'):
                    self.dictionary[key] = int(row.cells[1].text)
                else:
                    self.dictionary[key] = row.cells[1].text

    def get_from_tasks_table(self):
        # get the parameters from the critical tasks description table
        table_index = self.tables.index(self.TASKS_TABLE)
        table = self.text_input.tables[table_index]

        for i in range(1, self.dictionary['Number of critical tasks'] + 1):
            type_key = table.cell(i, 0).text + ' name'
            description_key = table.cell(i, 0).text + ' description'

            self.dictionary[type_key] = table.cell(i, 1).text
            self.dictionary[description_key] = table.cell(i, 2).text

    def get_from_effectiveness_table(self):
        # get the parameters from the effectiveness analysis problem type table
        table_index = self.tables.index(self.EFFECTIVENESS_TABLE)
        table = self.text_input.tables[table_index]

        '''Cells are not recognized as cells by word if they contain a dropdown list. That is why I had to create 
        a work around to get the values here.'''

        list_of_text = []  # list of the text of all cells, except those ones containing dropdown list
        stop = False

        number_of_problems = self.dictionary['Number of problems']

        # stores the text of every relevant cell in a list
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
            if i % 2 == 0:
                problem_number = list_of_text[i]
                description = list_of_text[i + 1]

                type_key = problem_number + ' type'
                description_key = problem_number + ' description'

                self.dictionary[type_key] = dropdown_lists_values[value]
                self.dictionary[description_key] = description

                value += 1

    @classmethod
    def get_all(cls, text_input_document, text_input_soup, list_of_table):

        parameters = cls(text_input_document, text_input_soup, list_of_table)

        parameters.get_from_standard_tables()
        parameters.get_from_tasks_table()
        parameters.get_from_effectiveness_table()

        return parameters.dictionary


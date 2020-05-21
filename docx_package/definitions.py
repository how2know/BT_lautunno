from docx import Document

# store all terms that we want to define in a list
defined_terms =[]
text_reading.find_definitions(tab_definition, [1, 4], defined_terms)

# store the definitions paragraphs and their styles in lists
definitions_list = []
definitions_styles_list = []
for i in range(len(defined_terms)):
    new_definition = []
    new_styles = []
    text_reading.paragraph_after_heading_with_styles(definitions.paragraphs, new_definition, new_styles, defined_terms[i], 'Heading 2')
    definitions_list.append(new_definition)
    definitions_styles_list.append(new_styles)

# find all terms that you want to define and store them
def find_definitions(table, columns_indexes_list, definitions_list):
    for j in columns_indexes_list:                                     # loop over the columns that contains "Yes" or "No"
        for i in range(len(table.columns[j].cells)):                   # loop over all cells of the columns
            if table.cell(i, j).text == 'Yes':                         # find all cells that contains "Yes"
                definitions_list.append(table.cell(i, j - 1).text)     # store the terms that correspond to a "Yes" 

class Definitions:
    def __init__(self):
        pass

    def write_definitions_chapter(self, document, heading_title, defined_terms_list, definitions_list,
                                  definitions_styles_list):
        document.add_heading(heading_title, 1)
        for i in range(len(defined_terms_list)):
            document.add_heading(defined_terms_list[i], 2)
            for j in range(len(definitions_list[i])):
                if definitions_styles_list[i][j] != 'Heading 1':
                    document.add_paragraph(definitions_list[i][j].text, definitions_styles_list[i][j])
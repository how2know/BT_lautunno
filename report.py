from docx import Document
from docx.enum.section import WD_SECTION
import os

from Writing_text import text, text_writing, layout


if __name__ == '__main__':
    # file name of the report
    filename = 'report.docx'

    # create document
    document = Document()

    # define all styles used in the document
    layout.define_all_styles(document)

    # section1 includes cover page and table of content
    section1 = document.sections[0]
    layout.define_page_format(section1)

    # add title
    title = text_writing.write_title(document, text.title)

    # add subtitle
    subtitle = document.add_heading(text.subtitle, 0)
    subtitle.style = document.styles['Subtitle']

    # create table of the document approval
    approval_table = document.add_table(rows=4, cols=4)

    # approval_table.style = 'Table Grid'
    layout.define_table_style(approval_table)
    for i in range(0, 4):
        for j in range(0, 4):
            approval_table.cell(i, j).text = text.approval_cells[i, j]

    # make the first row bold
    for col in approval_table.columns:
        col.cells[0].paragraphs[0].runs[0].font.bold = True

    # add the function to the person
    approval_table.cell(1, 1).add_paragraph(text.function_author)
    approval_table.cell(2, 1).add_paragraph(text.function_reviewer)
    approval_table.cell(3, 1).add_paragraph(text.function_approver)

    # make the function italic
    for i in range(1, 4):
        approval_table.rows[i].cells[1].paragraphs[1].runs[0].font.italic = True

    # add a page break
    document.add_page_break()

    # add table of content
    table_of_content = document.add_heading(text.toc_title, 1)

    # add a new section
    section2 = document.add_section(WD_SECTION.NEW_PAGE)
    layout.define_page_format(section2)

    # add a header
    layout.create_header(section2, text.first_header, text.second_header)

    text_writing.write_chapter(document, text.purpose_title, 1, text.purpose_paragraphs)

    text_writing.write_chapter(document, text.background_title, 1, text.background_paragraphs)

    text_writing.write_chapter(document, text.scope_title, 1, text.scope_paragraphs)

    text_writing.write_definitions_chapter(document, text.definitions_title, text.defined_terms, text.definitions_list,
                                           text.definitions_styles_list)

    text_writing.write_chapter(document, text.ethics_title, 1, text.ethics_paragraphs)

    text_writing.write_chapter(document, text.device_title, 1, text.device_paragraphs)

    document.add_heading(text.procedure_title, 1)

    text_writing.write_chapter(document, text.goal_title, 2, text.goal_paragraphs)

    text_writing.write_chapter(document, text.participants_title, 2, text.participants_paragraphs)

    text_writing.write_chapter(document, text.environment_title, 2, text.environment_paragraphs)

    text_writing.write_chapter(document, text.scenarios_title, 2, text.scenarios_paragraphs)

    text_writing.write_chapter(document, text.setup_title, 2, text.setup_paragraphs)

    text_writing.write_chapter(document, text.results_title, 1, text.results_paragraphs)

    text_writing.write_chapter(document, text.conclusion_title, 1, text.conclusion_paragraphs)

    '''
    for i in range(len(text.list_of_paragraphs)):
        document.add_heading(text.list_of_title[i].text, 1)
        for j in range(len(text.list_of_paragraphs[i])):
            document.add_paragraph(text.list_of_paragraphs[i][j].text)
    '''

    # save the report
    document.save(filename)

    # open the report with the default handler for .docx (Word)
    os.startfile(filename)


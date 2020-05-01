from Writing_text import layout


def write_title(document, text):
    title = document.add_paragraph(text)
    title.style = document.styles['Title']
    layout.insert_horizontal_border(title)

    return title


def write_chapter(document, heading_title, heading_level, paragraphs):
    document.add_heading(heading_title, heading_level)
    for i in range(len(paragraphs)):
        document.add_paragraph(paragraphs[i].text)


def write_definitions_chapter(document, heading_title, defined_terms_list, definitions_list, definitions_styles_list):
    document.add_heading(heading_title, 1)
    for i in range(len(defined_terms_list)):
        document.add_heading(defined_terms_list[i], 2)
        for j in range(len(definitions_list[i])):
            if definitions_styles_list[i][j] != 'Heading 1':
                document.add_paragraph(definitions_list[i][j].text, definitions_styles_list[i][j])
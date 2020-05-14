class Chapter:

    def __init__(self, title, style, parameter_table, picture_table):
        self.title = title
        self.style = style
        self.parameter_table = parameter_table
        self.picture_table = picture_table


    #  find a heading with his title and style and return the corresponding paragraph index
    def heading_name_index(self, paragraphs):
        for i in range(len(paragraphs)):  # loop over all paragraphs
            if paragraphs[i].style.name == self.style:  # look for paragraphs with corresponding style
                if paragraphs[i].text == self.title:  # look for paragraphs with corresponding title
                    return i  # return the index of the paragraphs

    # return the index of the next heading corresponding to a style given the index of the previous heading
    def next_heading_index(self, paragraphs, style, previous_index):
        for i in range(previous_index + 1, len(paragraphs)):  # loop over paragraphs coming after the given paragraph index
            if paragraphs[i].style.name == style:  # look for paragraphs with corresponding style
                return i  # return the index of the paragraph

    # return the index of the next heading corresponding to a style given the index of the previous heading
    def next_different_heading_index(self, paragraphs, style, previous_index):
        for i in range(previous_index + 1,
                       len(paragraphs)):  # loop over paragraphs coming after the given paragraph index
            if paragraphs[i].style.name == style:  # look for paragraphs with corresponding style
                return i  # return the index of the paragraph

    # store all paragraphs and their corresponding style between a given heading and the next one
    def paragraph_after_heading_with_styles(self, paragraphs, list_of_paragraphs, list_of_styles, heading_title,
                                            heading_style):
        heading_index = heading_name_index(paragraphs, heading_title, heading_style)  # index of the given heading
        next_index = next_heading_index(paragraphs, heading_style, heading_index)  # index of the next heading
        for i in range(heading_index + 1,
                       next_index):  # loop over all paragraphs between the given heading and the next one
            list_of_paragraphs.append(paragraphs[i])  # store all paragraphs in a given list
            list_of_styles.append(paragraphs[i].style.name)  # store all styles in a given list

    # store all paragraphs between a given heading and the next one
    def paragraph_after_heading(self, paragraphs, list_of_paragraphs, heading_title, heading_style):
        heading_index = heading_name_index(paragraphs, heading_title, heading_style)  # index of the given heading
        next_index = next_heading_index(paragraphs, heading_style, heading_index)  # index of the next heading
        for i in range(heading_index + 1,
                       next_index):  # loop over all paragraphs between the given heading and the next one
            list_of_paragraphs.append(paragraphs[i])  # store all paragraphs in a given list

    # store all paragraphs between a given heading and the next one
    def paragraph_after_heading_different(self, paragraphs, list_of_paragraphs, heading_title, heading_style1,
                                          heading_style2):
        heading_index = heading_name_index(paragraphs, heading_title, heading_style1)  # index of the given heading
        next_index = next_heading_index(paragraphs, heading_style2, heading_index)  # index of the next heading
        for i in range(heading_index + 1,
                       next_index):  # loop over all paragraphs between the given heading and the next one
            list_of_paragraphs.append(paragraphs[i])  # store all paragraphs in a given list
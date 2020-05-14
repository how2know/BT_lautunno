class Chapter:

    def __init__(self, document, title, style, parameter_table, picture_table):
        self.document = document
        self.title = title
        self.style = style
        self.parameter_table = parameter_table
        self.picture_table = picture_table


    #  find a heading with his title and style and return the corresponding paragraph index
    def heading_name_index(self):
        for i in range(len(self.document.paragraphs)):  # loop over all paragraphs
            if self.document.paragraphs[i].style.name == 'Heading {}'.format(1 or 2 or 3 or 4):  # look for paragraphs with corresponding style
                if self.document.paragraphs[i].text == self.title:  # look for paragraphs with corresponding title
                    return i  # return the index of the paragraphs

    # return the index of the next heading corresponding to a style given the index of the previous heading
    def next_heading_index(self, previous_index):
        for i in range(previous_index + 1, len(self.document.paragraphs)):  # loop over paragraphs coming after the given paragraph index
            if self.document.paragraphs[i].style.name == 'Heading {}'.format(1 or 2 or 3 or 4):  # look for paragraphs with corresponding style
                return i  # return the index of the paragraph

    # store all paragraphs and their corresponding style between a given heading and the next one
    def paragraph_after_heading_with_styles(self, list_of_paragraphs, list_of_styles):
        heading_index = self.heading_name_index()  # index of the given heading
        next_index = self.next_heading_index(heading_index)  # index of the next heading
        for i in range(heading_index + 1, next_index):  # loop over all paragraphs between the given heading and the next one
            list_of_paragraphs.append(self.document.paragraphs[i])  # store all paragraphs in a given list
            list_of_styles.append(self.document.paragraphs[i].style.name)  # store all styles in a given list

    # store all paragraphs between a given heading and the next one
    def paragraph_after_heading(self, paragraphs, list_of_paragraphs):
        heading_index = self.heading_name_index()  # index of the given heading
        next_index = self.next_heading_index(heading_index)  # index of the next heading
        for i in range(heading_index + 1, next_index):  # loop over all paragraphs between the given heading and the next one
            list_of_paragraphs.append(self.document.paragraphs[i])  # store all paragraphs in a given list
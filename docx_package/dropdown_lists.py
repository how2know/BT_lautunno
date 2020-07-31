from zipfile import ZipFile
from bs4 import BeautifulSoup
from typing import List


class DropDownLists:

    @ staticmethod
    def get_soup(text_input_path) -> BeautifulSoup:
        """
        Opens a .docx file as a .zip file and stores the XML data containing the infos about the .docx document
        in a BeautifulSoup object.

        Args:
            text_input_path: Path of the text input form.

        Returns:
            BeautifulSoup object that contains the XML data of the .docx document.
        """

        # open docx file as a zip file and store its relevant xml data
        zip_file = ZipFile(text_input_path)
        xml_data = zip_file.read('word/document.xml')
        zip_file.close()

        # parse the xml data with BeautifulSoup
        return BeautifulSoup(xml_data, 'xml')

    # return a list of the value of all dropdown lists in a table
    @ staticmethod
    def get_from_table(text_input_soup: BeautifulSoup, table_index: int) -> List[str]:
        """
        Args:
            text_input_soup: BeautifulSoup object that contains the XML data of the .docx document.
            table_index: Index of the table in the text input form.

        Returns:
            List of values of all dropdown lists in a table.
        """

        list_of_value = []

        # look for all values of dropdown lists in the data and store them
        tables = text_input_soup.find_all('tbl')
        dd_lists_content = tables[table_index].find_all('sdtContent')
        for i in dd_lists_content:
            list_of_value.append(i.find('t').string)

        return list_of_value

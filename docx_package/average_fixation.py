import seaborn as sns
import matplotlib.pyplot as plt
import pandas as pd


from docx_package.results import ResultsChapter


class AverageFixation:

    TITLE = 'Average fixation'
    TITLE_STYLE = 'Heading 2'
    DISCUSSION_TITLE = 'Discussion'
    DISCUSSION_STYLE = 'Heading 3'

    FIGURE_NAME = 'Average_fixation.png'

    def __init__(self, report_document, text_input_document, text_input_soup, list_of_tables, parameters_dictionary, txt_data, list_of_dataframes):
        self.report = report_document
        self.text_input = text_input_document
        self.text_input_soup = text_input_soup
        self.tables = list_of_tables
        self.parameters = parameters_dictionary
        self.txt_data = txt_data
        self.cGOM_dataframes = list_of_dataframes

    def make_plot(self, data_frame, title, save_path):
        sns.set(style='whitegrid')
        plot = sns.boxplot(data=data_frame)
        plt.ylabel('Fixation duration [s]')
        plot.set_title(title)
        figure = plot.get_figure()
        figure.savefig(save_path)
        plt.show()

    def areas_of_interest(self, dataframe):
        labels_list = dataframe.index.values.tolist()
        areas_of_interest = []

        for label in labels_list:
            if label not in areas_of_interest:
                areas_of_interest.append(label)

        return areas_of_interest

    def fixations(self, aois, dataframe):
        participant_fixations_df = pd.DataFrame()

        for idx, aoi in enumerate(aois):
            data_of_aoi = dataframe[dataframe.index == aoi]
            fixations = data_of_aoi['Fixation time'].to_numpy()
            aoi_fixations = pd.DataFrame(columns=[aoi], data=fixations)

            participant_fixations_df = participant_fixations_df.append(aoi_fixations, ignore_index=True)

        return participant_fixations_df

    def average_fixation_stat(self):
        average_fixation_df = pd.DataFrame()

        for idx, dataframe in enumerate(self.cGOM_dataframes):
            aois = self.areas_of_interest(dataframe)

            participant_fixations = self.fixations(aois, dataframe)

            average_fixation_df = average_fixation_df.append(participant_fixations, ignore_index=True)

            self.make_plot(participant_fixations,
                           'Participant {}'.format(str(idx+1)),
                           'Average_fixation_Participant{}.png'.format(str(idx+1))
                           )

        return average_fixation_df

    def write_chapter(self):
        time_on_tasks = ResultsChapter(self.report, self.text_input, self.text_input_soup, self.TITLE,
                                       self.tables, self.parameters)

        self.make_plot(self.average_fixation_stat(),
                       'Average fixation',
                       self.FIGURE_NAME
                       )

        self.report.add_paragraph(self.TITLE, self.TITLE_STYLE)
        self.report.add_picture(self.FIGURE_NAME)

        self.report.add_paragraph(self.DISCUSSION_TITLE, self.DISCUSSION_STYLE)
        time_on_tasks.write_chapter()

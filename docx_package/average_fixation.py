import seaborn as sns
import matplotlib.pyplot as plt
import pandas as pd


from docx_package.results import ResultsChapter


class AverageFixation:

    TITLE = 'Average fixation'
    TITLE_STYLE = 'Heading 2'
    DISCUSSION_TITLE = 'Discussion'
    DISCUSSION_STYLE = 'Heading 3'

    def __init__(self, report_document, text_input_document, text_input_soup, list_of_tables, parameters_dictionary, txt_data, list_of_dataframes):
        self.report = report_document
        self.text_input = text_input_document
        self.text_input_soup = text_input_soup
        self.tables = list_of_tables
        self.parameters = parameters_dictionary
        self.txt_data = txt_data
        self.cGOM_dataframes = list_of_dataframes

    def areas_of_interest(self, dataframe):
        labels_list = dataframe.index.values.tolist()
        areas_of_interest = []

        for label in labels_list:
            if label not in areas_of_interest:
                areas_of_interest.append(label)

        return areas_of_interest

    def fixations(self, aois, dataframe, title):
        participant_fixations_df = pd.DataFrame()

        for idx, aoi in enumerate(aois):
            data_of_aoi = dataframe[dataframe.index == aoi]
            fixations = data_of_aoi['Fixation time'].to_numpy()
            aoi_fixations = pd.DataFrame(columns=[aoi], data=fixations)

            # print(aoi_fixations)

            participant_fixations_df = participant_fixations_df.append(aoi_fixations, ignore_index=True)

            # print(participant_fixations_df)

            '''
            sns.set(style='whitegrid')
            plot = sns.boxplot(data=participant_fixations_df)
            # plt.xlabel('Tasks')
            plt.ylabel('Fixation duration [ms]')
            # figure = plot.get_figure()
            # figure.savefig(self.FIGURE_NAME)
            plt.show()
            '''


        print(participant_fixations_df)

        sns.set(style='whitegrid')
        plot = sns.boxplot(data=participant_fixations_df)
        plt.ylabel('Fixation duration [ms]')
        plot.set_title(title)
        plt.show()


        return participant_fixations_df

    def average_fixation_stat(self):
        average_fixation_df = pd.DataFrame()

        for idx, dataframe in enumerate(self.cGOM_dataframes):
            aois = self.areas_of_interest(dataframe)

            # print(self.fixations(aois, dataframe))

            average_fixation_df = average_fixation_df.append(self.fixations(aois, dataframe, str(idx)), ignore_index=True)

        # print(average_fixation_df)

        test = average_fixation_df.dropna(subset=['Sink'])

        # print(test['Sink'])

        print(average_fixation_df)

        return average_fixation_df

        # return test

    def make_plot(self):
        sns.set(style='whitegrid')
        '''
        plot = sns.boxplot(data=self.txt_data,
                           y='Fixation time',
                           x='Label'
                           )
        '''
        plot = sns.boxplot(data=self.average_fixation_stat())
        # plt.xlabel('Tasks')
        plt.ylabel('Fixation duration [ms]')
        plot.set_title('Main graph bordel')
        figure = plot.get_figure()
        figure.savefig('Average_fixation.png')
        plt.show()

    def write_chapter(self):
        time_on_tasks = ResultsChapter(self.report, self.text_input, self.text_input_soup, self.TITLE,
                                       self.tables, self.parameters)

        self.report.add_paragraph(self.TITLE, self.TITLE_STYLE)

        self.report.add_paragraph(self.DISCUSSION_TITLE, self.DISCUSSION_STYLE)
        time_on_tasks.write_chapter()
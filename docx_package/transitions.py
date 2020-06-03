import numpy as np
import seaborn as sns
import matplotlib.pyplot as plt
import pandas as pd

from docx_package.results import ResultsChapter


class Transitions:

    TITLE = 'Transitions'
    TITLE_STYLE = 'Heading 2'
    DISCUSSION_TITLE = 'Discussion'
    DISCUSSION_STYLE = 'Heading 3'

    # TODO: delete txt data
    def __init__(self, report_document, text_input_document, text_input_soup, list_of_tables, parameters_dictionary, txt_data, list_of_dataframes):
        self.report = report_document
        self.text_input = text_input_document
        self.text_input_soup = text_input_soup
        self.tables = list_of_tables
        self.parameters = parameters_dictionary
        self.txt_data = txt_data
        self.cGOM_dataframes = list_of_dataframes

    # TODO: write % somewhere
    def make_plot(self, data_frame, title, save_path):
        plot = sns.heatmap(data=data_frame,
                           vmin=0, vmax=100,
                           annot=True,
                           linewidths=.5,
                           cmap='YlOrRd',
                           cbar=False,
                           fmt='.2%'
                           )
        plt.ylabel('AOI source (from)')
        plt.xlabel('AOI destination (to)')
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

    def transitions(self, aois, dataframe):
        table = pd.DataFrame(index=aois,
                             columns=aois,
                             data=np.zeros((len(aois), len(aois)))
                             )

        all_fixations = dataframe.index.values.tolist()

        last_fixation = all_fixations[0]

        for fixation in all_fixations[1:]:
            # table.at[last_fixation, fixation] += 1
            table.loc[last_fixation].at[fixation] += 1

            last_fixation = fixation

        transitions_number = table.to_numpy().sum()

        table = table.div(transitions_number)
        # table = table.mul(100)

        return table

    def transitions_stat(self):
        all_transitions = pd.DataFrame()
        all_data = pd.DataFrame()

        for idx, dataframe in enumerate(self.cGOM_dataframes):
            aois = self.areas_of_interest(dataframe)

            participant_transitions = self.transitions(aois, dataframe)

            self.make_plot(participant_transitions,
                           'Participant {}'.format(str(idx + 1)),
                           'Transitions_Participant{}.png'.format(str(idx + 1))
                           )

            all_transitions = all_transitions.append(participant_transitions)
            all_data = all_data.append(dataframe)

        all_aois = self.areas_of_interest(all_data)

        transitions_stat1 = self.transitions(all_aois, all_data)
        self.make_plot(transitions_stat1,
                       'Transitions main graph 1',
                       'Transitions1.png'
                       )

        transitions_stat2 = pd.DataFrame()
        for idx, aoi in enumerate(all_aois):
            data_of_aoi = all_transitions[all_transitions.index == aoi]

            mean = data_of_aoi.mean()

            print(mean)

            transitions_stat2 = transitions_stat2.append(mean, ignore_index=True)

            transitions_stat2.rename(index={idx: str(aoi)})

        print(transitions_stat2)
        self.make_plot(transitions_stat2,
                       'Transitions main graph 2',
                       'Transitions2.png'
                       )

    def write_chapter(self):
        time_on_tasks = ResultsChapter(self.report, self.text_input, self.text_input_soup, self.TITLE,
                                       self.tables, self.parameters)

        self.report.add_paragraph(self.TITLE, self.TITLE_STYLE)

        self.report.add_paragraph(self.DISCUSSION_TITLE, self.DISCUSSION_STYLE)
        time_on_tasks.write_chapter()
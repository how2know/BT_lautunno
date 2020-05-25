import numpy as np
import seaborn as sns
import matplotlib.pyplot as plt
import pandas as pd


class TimeOnTasks:

    FIGURE_NAME = 'Time_on_task.png'

    def __init__(self, report_document, text_input_document, list_of_tables, parameters_dictionary):
        self.report = report_document
        self.parameters = parameters_dictionary
        input_table_index = list_of_tables.index('Time on tasks table')
        self.input_table = text_input_document.tables[input_table_index]

    @ property
    def tasks(self):
        tasks = []
        for i in range(1, self.parameters['Number of critical tasks'] + 1):
            tasks.append(self.parameters['Critical task {} name'.format(i)])
        return tasks

    @ property
    def participants(self):
        participants = ['Participants {}'.format(i) for i in range(1, self.parameters['Number of participants'] + 1)]
        return participants

    @ property
    def times(self):
        rows = self.parameters['Number of critical tasks']
        columns = self.parameters['Number of participants']
        times = np.zeros((rows, columns))
        for i in range(rows):
            for j in range(columns):
                time = float(self.input_table.cell(i+1, j+1).text)
                times[i, j] = time

        transpose = times.transpose()
        return transpose

    @ property
    def data(self):
        data = pd.DataFrame(self.times, index=self.participants, columns=self.tasks)
        return data

    def make_plot(self):
        sns.set(style="whitegrid")
        plot = sns.barplot(data=self.data)
        figure = plot.get_figure()
        figure.savefig(self.FIGURE_NAME)
        # plt.show()

    def insert_plot(self):
        self.make_plot()
        self.report.add_picture(self.FIGURE_NAME)


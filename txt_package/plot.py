import numpy as np
import seaborn as sns
import matplotlib.pyplot as plt
import pandas as pd

import time


def make_barplot(data_frame, figure_save_path, title=None, xlabel=None, ylabel=None):
    """
    Create a barplot out of a data frame and save its figure.

    Args:
        data_frame: Data frame containing the data from which the plot will be made.
        figure_save_path: Path where the figure of the plot will be saved.
        title (optional): Plot title written on the figure.
        xlabel (optional): Label of the x-axis.
        ylabel (optional): Label of the y-axis.
    """

    sns.set(style='whitegrid')

    plot = sns.barplot(data=data_frame,
                       capsize=0.1,     # length of the caps at the endpoint of the confidence interval bar
                       errwidth=1.5,     # width of the confidence interval bar
                       ci=95     # size of the confidence interval
                       )

    if title:
        plot.set_title(title)
    if xlabel:
        plt.xlabel(xlabel)
    if ylabel:
        plt.ylabel(ylabel)

    figure = plot.get_figure()
    figure.savefig(figure_save_path)

    plt.show()


def make_boxplot(data_frame, figure_save_path, title=None, xlabel=None, ylabel=None):
    """
    Create a box plot out of a data frame and save its figure.

    Args:
        data_frame: Data frame containing the data from which the plot will be made.
        figure_save_path: Path where the figure of the plot will be saved.
        title (optional): Plot title written on the figure.
        xlabel (optional): Label of the x-axis.
        ylabel (optional): Label of the y-axis.
    """

    sns.set(style='whitegrid')
    plot = sns.boxplot(data=data_frame)

    if title:
        plot.set_title(title)
    if xlabel:
        plt.xlabel(xlabel)
    if ylabel:
        plt.ylabel(ylabel)

    figure = plot.get_figure()
    figure.savefig(figure_save_path)

    plt.show()


def make_heatmap(data_frame, figure_save_path, title=None, xlabel=None, ylabel=None):
    """
    Create a heat map out of a data frame and save its figure.

    Args:
        data_frame: Data frame containing the data from which the plot will be made.
        figure_save_path: Path where the figure of the plot will be saved.
        title (optional): Plot title written on the figure.
        xlabel (optional): Label of the x-axis.
        ylabel (optional): Label of the y-axis.
    """

    plot = sns.heatmap(data=data_frame,
                       vmin=0, vmax=1,     # max and min value
                       annot=True,     # annotate each cell
                       linewidths=.5,     # width of the line between each cell
                       cmap='YlOrRd',     # color of the cells
                       cbar=False,     # bar showing the colors
                       fmt='.2%'     # formatting of the annotation
                       )

    if title:
        plot.set_title(title)
    if xlabel:
        plt.xlabel(xlabel)
    if ylabel:
        plt.ylabel(ylabel)

    figure = plot.get_figure()
    figure.savefig(figure_save_path)

    plt.show()

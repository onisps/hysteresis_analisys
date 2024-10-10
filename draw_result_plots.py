import os, time
import pandas as pd
from glob2 import glob
from matplotlib import pyplot as plt
import matplotlib
import pathlib
import xlsxwriter
import xlrd
import numpy as np
import openpyxl, re
from matplotlib import rc
import matplotlib as mpl
import matplotlib.ticker as ticker

def comma_format(x, p):
    return format(x, "3,.2f").replace(".", ",")

pathGlobal = str(pathlib.Path(__file__).parent.resolve())

if __name__ == '__main__':
    if not os.path.exists(os.path.join(pathGlobal, 'pics', 'all_loops')):
        os.makedirs(os.path.join(pathGlobal, 'pics', 'all_loops'))
    # if not os.path.exists(os.path.join(pathGlobal, 'pics', '3_last_loops')):
    #     os.makedirs(os.path.join(pathGlobal, 'pics', '3_last_loops'))
    subfolders = [x[0] for x in os.walk(os.path.join(pathGlobal, 'csv'))][1:]
    steps = [10, 20, 30, 40, 50, 75, 100, 150, 200]
    for folder in subfolders:
        load_list = glob(os.path.join(pathGlobal, folder + '/load*.csv'), recursive=True)
        unload_list = glob(os.path.join(pathGlobal, folder + '/unload*.csv'), recursive=True)
        load_list.sort()
        unload_list.sort()
        try:
            f = plt.figure(figsize=(10,10))
            mpl.rcParams['axes.spines.right'] = False
            mpl.rcParams['axes.spines.top'] = False
            plt.gca().get_yaxis().set_major_formatter(ticker.FuncFormatter(comma_format))
            # max_load = 0
            max_load_for_draw = 0
            max_elong = 0
            for load_file, unload_file in zip(load_list, unload_list):
                data_load = pd.read_csv(load_file)
                data_load = data_load.to_numpy()[:, 1:].T
                data_unload = pd.read_csv(unload_file)
                data_unload = data_unload.to_numpy()[:, 1:].T
                if load_file.split('.')[0].split('-')[-1] in ['6', '7', '8']:
                    # max_load = np.max((data_load[1,-1], max_load))
                    max_elong = np.max((data_load[0,-1], max_elong))
                    color_of_line = 'firebrick'
                else:
                    max_load_for_draw = np.max((data_load[1, -1], max_load_for_draw))
                    color_of_line = 'slategrey'
                data_plot = np.hstack((data_load, data_unload))
                plt.plot(data_plot[0], data_plot[1], color=color_of_line, linewidth=2)
                # plt.plot(data_load[0], data_load[1], color=color_of_line, linewidth=2)
                # plt.plot(data_unload[0], data_unload[1], color=color_of_line, linewidth=2)

            # plt.vlines(x=min(steps, key=lambda x: abs(x - max_elong)), ymin=0, ymax=max_load_for_draw*1.2,
            #            colors='grey', linewidth=1.25, alpha=0.5)
            plt.text(x=max_elong, y=max_load_for_draw,
                     s=f'{max_load_for_draw:2.4} МПа\n{max_elong:2.4} %'.replace('.', ','),
                     horizontalalignment='left', fontsize=12)
            plt.title(f'Гистерезис {str(" ").join(folder.split("/")[-1].split(".")[0].split(" ")[2:-1])}', fontsize=16)
            plt.xlim((0, min(steps, key=lambda x: abs(x - max_elong))))
            plt.xlabel('Относительное удлинение, %', fontsize=16)
            plt.ylabel('Напряжение, МПа', fontsize=16)
            plt.yticks(np.arange(0, np.ceil(max_load_for_draw)*2, (np.ceil(max_load_for_draw)*2)/20))
            plt.ylim(0, np.ceil(max_load_for_draw))
            f.savefig(os.path.join(pathGlobal,
                                   'pics/all_loops',
                                   str(" ").join(folder.split("/")[-1].split(".")[0].split(" ")[2:-1])+'.png'),
                      dpi=350)
            plt.close()
        except Exception as e:
            print(f'Exception: {e}')

        # try:
        #     f = plt.figure(figsize=(10,10))
        #     plt.gca().get_yaxis().set_major_formatter(ticker.FuncFormatter(comma_format))
        #     mpl.rcParams['axes.spines.right'] = False
        #     mpl.rcParams['axes.spines.top'] = False
        #     max_load = 0
        #     max_elong = 0
        #     for load_file, unload_file in zip(load_list, unload_list):
        #
        #         data_load = pd.read_csv(load_file)
        #         data_load = data_load.to_numpy()[:, 1:].T
        #         data_unload = pd.read_csv(unload_file)
        #         data_unload = data_unload.to_numpy()[:, 1:].T
        #         if load_file.split('.')[0].split('-')[-1] in ['6', '7', '8']:
        #             max_load = np.max((data_load[1,-1], max_load))
        #             max_elong = np.max((data_load[0,-1], max_elong))
        #             color_of_line = 'slategrey'
        #         else:
        #             continue
        #         data_plot = np.hstack((data_load, data_unload))
        #         plt.plot(data_plot[0], data_plot[1], color=color_of_line, linewidth=2)
        #
        #     plt.vlines(x=min(steps, key=lambda x: abs(x - max_elong)), ymin=0, ymax=max_load*1.2,
        #                colors='grey', linewidth=1.25, alpha=0.5)
        #     plt.text(x=max_elong, y=max_load,
        #              s=f'{max_load:2.4} МПа\n{max_elong:2.4} %'.replace('.', ','),
        #              horizontalalignment='right')
        #     plt.title(f'3 last loops of {str(" ").join(folder.split("/")[-1].split(".")[0].split(" ")[2:-1])}')
        #     plt.xlim((0, 250))
        #     plt.ylim((0, max_load * 1.2))
        #
        #     f.savefig(os.path.join(pathGlobal,
        #                            'pics/3_last_loops',
        #                            str(" ").join(folder.split("/")[-1].split(".")[0].split(" ")[2:-1]) + '.png'),
        #               dpi=350)
        #     plt.close()
        # except Exception as e:
        #     print(f'Exception: {e}')
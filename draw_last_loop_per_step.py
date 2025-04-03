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

def sort_human(l):
    convert = lambda text: float(text) if text.isdigit() else text
    alphanum = lambda key: [convert(c) for c in re.split('([-+]?[0-9]*\.?[0-9]*)', key)]
    l.sort(key=alphanum)
    return l

pathGlobal = str(pathlib.Path(__file__).parent.resolve())

if __name__ == '__main__':
    if not os.path.exists(os.path.join(pathGlobal, 'pics', 'last_loops_by_step')):
        os.makedirs(os.path.join(pathGlobal, 'pics', 'last_loops_by_step'))
    # via https://matplotlib.org/stable/_images/sphx_glr_named_colors_003.png
    color_dict = [
        '#048F98',
        '#6CA0DF',
        '#7459C1',
        'tab:green',
        'tab:blue',
        'gold',
        'darkgrey',
        'lightgrey',
        'lightcoral',
        'firebrick',
        'sienna',
        'darkorange',
        'tan',
        'gold',
        'darkkhaki',
        'olive',
        'yellowgreen',
        'darkseagreen',
        'forestgreen',
        'aquamarine',
        'mediumturquoise',
        'aqua',
        'deepskyblue',
        'steelblue',
        'navy',
        'slateblue',
        'mediumpurple',
        'indigo',
        'darkviolet',
        'violet',
        'deeppink',
        'crimson'
    ]
    # subfolders = [x[0] for x in os.walk(os.path.join(pathGlobal, 'csv'))][1:]
    subfolders = ['E:\\work\\python\\hysteresis_analisys\\csv\\17_02_2025 гистерезис  12 PVA 146-186 DMSO-H2O = 60-40 1 loop 1',
                  'E:\\work\\python\\hysteresis_analisys\\csv\\17_02_2025 гистерезис  12 PVA 146-186 1 loop 2',
                  'E:\\work\\python\\hysteresis_analisys\\csv\\31_07_2024 гистерезис PVA-CNT-L 2']
    steps = [10, 20, 30, 40, 50, 75, 100, 150, 200]
    for step in steps:
        max_load = 0
        max_elong = 0
        max_load_for_draw = 0
        try:
            legend_array = []
            f = plt.figure(figsize=(10, 10))
            mpl.rcParams['axes.spines.right'] = False
            mpl.rcParams['axes.spines.top'] = False
            plt.gca().get_yaxis().set_major_formatter(ticker.FuncFormatter(comma_format))
            for folder, color_of_line in zip(subfolders, color_dict):
                if os.path.exists(os.path.join(folder,f'load_curve_{step}-8.csv')):
                    data_load = pd.read_csv(os.path.join(folder,f'load_curve_{step}-8.csv'))
                    data_load = data_load.to_numpy()[:, 1:].T
                    data_unload = pd.read_csv(os.path.join(folder,f'unload_curve_{step}-8.csv'))
                    data_unload = data_unload.to_numpy()[:, 1:].T
                    max_load = np.max((data_load[1,-1], max_load))
                    max_elong = np.max((data_load[0,-1], max_elong))
                    max_load_for_draw = np.max((data_load[1, -1], max_load_for_draw))
                    # color_of_line = 'firebrick'
                    data_plot = np.hstack((data_load, data_unload))
                    plt.plot(data_plot[0], data_plot[1], color=color_of_line, linewidth=2)
                    plt.text(x=data_load[0,-1], y=data_load[1,-1],
                             s=f'{str(" ").join(folder.split("/")[-1].split(".")[0].split(" ")[2:-1])}',
                             horizontalalignment='right', alpha=0.5)
                    legend_array.append(f'{str(" ").join(folder.split("/")[-1].split(".")[0].split(" ")[2:-1])}')
                else:
                    continue
            # plt.vlines(x=min(steps, key=lambda x: abs(x - max_elong)), ymin=0, ymax=max_load*1.2,
            #            colors='grey', linewidth=1.25, alpha=0.5)
            plt.title(f'Петли с {step} % удлинения', fontsize=16)
            plt.xlim((0, min(steps, key=lambda x: abs(x - max_elong))))
            plt.yticks(np.arange(0, np.ceil(max_load_for_draw) * 2, (np.ceil(max_load_for_draw) * 2) / 20))
            plt.ylim(0, np.ceil(max_load_for_draw))
            plt.xlabel('Относительное удлинение, %', fontsize=16)
            plt.ylabel('Напряжение, МПа', fontsize=16)
            plt.legend(legend_array)
            plt.gcf().bbox.containsx(15)
            f.savefig(os.path.join(pathGlobal,
                                   'pics/last_loops_by_step',
                                   f'step-{step}'+'.png'),
                      dpi=350)
            plt.close()
        except Exception as e:
            print(f'Exception: {e}')
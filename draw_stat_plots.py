import numpy as np
import pandas as pd
from matplotlib import pyplot as plt
import matplotlib as mpl

if __name__ == '__main__':
    file_names = ['res_stat_by_math.xlsx', 'res_stat_one_by_one.xlsx']
    file_name = file_names[1]
    data = pd.read_excel(file_name)
    materials = np.unique(data['Материал'].to_numpy())
    data_np = data.to_numpy()
    titles = data.columns
    ind_titles = 3

    for curr_ind in np.arange(ind_titles, len(titles)):
        if 'удлинение' in titles[curr_ind].lower() or 'деформация' in titles[curr_ind].lower() or 'расстояние' in titles[curr_ind].lower():
            ylabel = 'Относительное удлинение, %'
        elif 'напряжение' in titles[curr_ind].lower() or 'напряжения' in titles[curr_ind].lower():
            ylabel = 'Напряжение, МПа'
        elif 'площадь' in titles[curr_ind].lower():
            ylabel = 'Площадь, МПа*мм/мм'
        f = plt.figure(figsize=(15, 10))
        mpl.rcParams['axes.spines.right'] = False
        mpl.rcParams['axes.spines.top'] = False
        for mat in materials:
            curr_data = data_np[data_np[:, 1] == mat, :]
            steps = np.unique(curr_data[:, 2])
            plot_data = list()
            for step in steps:
                plot_data.append(np.average(curr_data[curr_data[:,2] == step,curr_ind]))
            plt.plot(steps, plot_data)
        plt.legend(materials)
        plt.title(f'{titles[curr_ind]}')
        plt.xlabel('Удлинение, %')
        plt.ylabel(ylabel)
        plt.savefig(f'./pics/{titles[curr_ind]}.png')

    print(data)
import numpy as np
import pandas as pd
from matplotlib import pyplot as plt
import matplotlib as mpl

if __name__ == '__main__':
    file_name = 'res_stat.xlsx'
    data = pd.read_excel(file_name)
    materials = np.unique(data['Материал'].to_numpy())
    data_np = data.to_numpy()
    titles = data.columns
    ind_titles = 3

    for curr_ind in np.arange(ind_titles, len(titles)):
        f = plt.figure(figsize=(15, 10))
        mpl.rcParams['axes.spines.right'] = False
        mpl.rcParams['axes.spines.top'] = False
        for mat in materials:
            curr_data = data_np[data_np[:, 1] == mat, :]
            steps = curr_data[:, 2]
            plt.plot(steps, curr_data[:, curr_ind])
        plt.legend(materials)
        plt.title(f'{titles[curr_ind]}')
        plt.savefig(f'./pics/{titles[curr_ind]}.png')

    print(data)
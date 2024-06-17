import os, time
import pandas as pd
from glob2 import glob
from matplotlib import pyplot as plt
import pathlib
import xlsxwriter
import xlrd
import numpy as np
import openpyxl
import matplotlib as mpl

pathGlobal = str(pathlib.Path(__file__).parent.resolve())

class colors:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKCYAN = '\033[96m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'


def split_arrays(elong, stress):
    s_arr = stress[:np.argwhere(stress == np.max(stress[:int(2 / 3 * len(stress))]))[0][0]]
    top_point = len(s_arr)

    elong_top = elong[:top_point]
    stress_top = stress[:top_point]

    temp = stress[top_point + 1:]
    s_arr = temp[:np.argwhere(temp == np.min(temp[:int(2 / 3 * len(temp))]))[0][0]]
    bot_point = len(s_arr)
    stress_bot = stress[top_point + 1:top_point + 1 + bot_point]
    elong_bot = elong[top_point + 1:top_point + 1 + bot_point]
    return [elong_top, stress_top], [elong_bot, stress_bot]


def compute_square(arr_s, arr_e):
    square = 0
    for i in np.arange(1, len(arr_s)):
        square += arr_s[i] * abs(arr_e[i] - arr_e[i - 1])
    return square


if __name__ == '__main__':
    if not os.path.exists('./pics'):
        os.makedirs('./pics')
    row = ['Дата изменения','Файл', 'Шаг', 'Площадь', 'Накопленное удлинение vs 1', 'Накопленное удлинение vs prev',
           'Удлинение петли', 'Удлинение из 0 начала петли']
    res = pd.DataFrame(columns=row)
    res.to_excel('./res.xlsx', startrow=0, startcol=-1)
    workbook = openpyxl.load_workbook('./res.xlsx')
    worksheet = workbook.active
    paths = [
             'data/test'
             #'data/гистерезис PVA',
             # 'data/Гистерезис SIBS',
             # 'data/Гистерезис перикадр',
             # 'data/Полимеры'
             ]
    saved_elongation_v1 = 0
    for path in paths:
        for file in glob(os.path.join(pathGlobal, path + '/*.xls'), recursive=True):
            data = pd.ExcelFile(file)
            drops = ['Параметры', 'Результаты', 'Статистика']
            sheets = data.sheet_names
            for drop in drops:
                sheets = [value for value in sheets if value != drop]

            for sheet in sheets:

                data_loaded = pd.read_excel(file, sheet_name=sheet)
                elongation = np.array(data_loaded.to_numpy()[2:, 0], dtype='float64')
                stress = np.array(data_loaded.to_numpy()[2:, 1], dtype='float64')
                stress[stress < 5e-4] = 0
                top_side, bot_side = split_arrays(elongation, stress)
                try:
                    last = np.argwhere(top_side[1] == 0)[-1]
                    top_side[1][:last[0]] = 0
                except:
                    pass
                if sheet[-1] == '1':
                    saved_elongation_v1 = top_side[0][-1]

                if len(bot_side[0]) < len(top_side[0]) * 0.25:
                    print(colors.FAIL + f'skipped {(file.split("/")[-1]).split(".")[0]} - {sheet}' + colors.ENDC)
                    continue
                # else:
                #     print(colors.OKGREEN + f'processing {(file.split("/")[-1]).split(".")[0]} - {sheet}' + colors.ENDC)
                square_top = compute_square(arr_e=top_side[0], arr_s=top_side[1])
                square_bot = compute_square(arr_e=bot_side[0], arr_s=bot_side[1])
                f = plt.figure(figsize=(15, 10))
                mpl.rcParams['axes.spines.right'] = False
                mpl.rcParams['axes.spines.top'] = False
                plt.plot(top_side[0], top_side[1], 'r-')
                plt.text(x=top_side[0][-1], y=top_side[1][-1],
                         s=(f'el: {top_side[0][-1]:2.2f}\ns: {top_side[1][-1]:2.2f}'))
                plt.plot(top_side[0][-1], top_side[1][-1], 'or')

                plt.plot(bot_side[0], bot_side[1], 'c-')
                plt.text(x=bot_side[0][-1], y=bot_side[1][-1],
                         s=(f'el: {bot_side[0][-1]:2.2f}\ns: {bot_side[1][-1]:2.2f}'))
                plt.plot(bot_side[0][-1], bot_side[1][-1], 'or')
                top_loop_start = [0, 0]
                try:
                    arg = np.argwhere(top_side[1] == 0)[-1]
                    top_loop_start[0] = top_side[0][arg][0]
                    top_loop_start[1] = top_side[1][arg][0]
                except:
                    pass
                fs = 14
                plt.plot(top_loop_start[0], top_loop_start[1], 'xk')
                plt.text(x=top_loop_start[0], y=top_loop_start[1],
                         s=(f'el: {top_loop_start[0]:2.2f}\ns: {top_loop_start[1]:2.2f}'))
                plt.title(
                    f'top: {float(square_top):2.2} |'
                    f' bot: {float(square_bot):2.2f} |'
                    f' sq: {float(square_top - square_bot):2.2f}'
                )


                plt.savefig('./pics/' + file.split('/')[-1].split('.')[0] + '_' + sheet + '.jpg')
                plt.close()

                if sheet[-1] == '1':
                    row = [
                            time.ctime(os.path.getctime(file)),
                            (file.split('/')[-1]).split('.')[0],
                            sheet, square_top - square_bot, 0, 0,
                            bot_side[0][-1] - top_loop_start[0],
                            top_loop_start[0]
                           ]
                else:
                    row = [
                            time.ctime(os.path.getctime(file)),
                            (file.split('/')[-1]).split('.')[0],
                            sheet, square_top - square_bot,
                            top_side[0][-1] - saved_elongation_v1,
                            top_side[0][-1] - saved_elongation,
                            bot_side[0][-1] - top_loop_start[0],
                            top_loop_start[0]
                           ]
                saved_elongation = top_side[0][-1]
                worksheet.append(row)
                workbook.save('./res.xlsx')

import os, time
import pandas as pd
from glob2 import glob
from matplotlib import pyplot as plt
import pathlib
import xlsxwriter
import xlrd
import numpy as np
import openpyxl, re
import matplotlib as mpl
from scipy import interpolate

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
    # s_arr = stress[:np.argwhere(stress == np.max(stress[:int(2 / 3 * len(stress))]))[0][0]]
    s_arr = stress[:np.argwhere(stress == np.max(stress))[0][0]]
    top_point = len(s_arr)

    elong_top = elong[:top_point]
    stress_top = stress[:top_point]

    temp = stress[top_point + 1:]
    # s_arr = temp[:np.argwhere(temp == np.min(temp[:int(2 / 3 * len(temp))]))[0][0]]
    s_arr = temp[:np.argwhere(temp == np.min(temp))[0][0]]
    bot_point = len(s_arr)
    stress_bot = stress[top_point + 1:top_point + 1 + bot_point]
    elong_bot = elong[top_point + 1:top_point + 1 + bot_point]
    return [elong_top, stress_top], [elong_bot, stress_bot]


def interpolate_data(arr_s, arr_e, size=500):
    interpolator = interpolate.interp1d(arr_e, arr_s)
    new_e = np.linspace(arr_e[0], arr_e[-1], size)
    new_s = interpolator(np.linspace(arr_s[0], arr_s[-1], size))
    return new_e, new_s


def compute_square(arr_s, arr_e):
    square = 0
    for i in np.arange(1, len(arr_s)):
        square += arr_s[i] * abs(arr_e[i] - arr_e[i - 1])
    return square


if __name__ == '__main__':
    pd.set_option('max_colwidth', 50)
    trim_coeff = 0.0005
    if not os.path.exists('./pics'):
        os.makedirs('./pics')
    # row = ['Дата изменения', 'Файл', 'Шаг', 'Цикл', 'Площадь',
    #        'Накопленное удлинение vs 1', 'Накопленное удлинение vs prev',
    #        'Удлинение петли', 'Удлинение из 0 начала петли',
    #        'Деградация напряжение vs 1', 'Деградация напряжения vs prev',
    #        'Макс напряжение']
    # res = pd.DataFrame(columns=row)
    # res.to_excel('./res.xlsx', startrow=0, startcol=-1, sheet_name='descriptive')
    workbook = openpyxl.load_workbook('./res.xlsx')
    worksheet = workbook.active
    paths = [
        'data/test'
        # 'data/гистерезис PVA',
        # 'data/Гистерезис SIBS',
        # 'data/Полимеры',
        # 'data/Гистерезис перикадр'
    ]
    saved_elongation_v1 = 0
    is_break = False
    for path in paths:
        for file in glob(os.path.join(pathGlobal, path + '/*.xls'), recursive=True):
            print(colors.BOLD+colors.HEADER+f'current file is {file}!'+colors.ENDC)
            file_name = ""
            for v in (file.split('/')[-1]).split('.')[0:-1]:
                file_name += f'{v} '
            if not os.path.exists(f'./pics/line/{file_name}'):
                os.makedirs(f'./pics/line/{file_name}')
            if not os.path.exists(f'./pics/scatter/{file_name}'):
                os.makedirs(f'./pics/scatter/{file_name}')
            if not os.path.exists(f'./pics/one-by-one/{file_name}'):
                os.makedirs(f'./pics/one-by-one/{file_name}')
            data = pd.ExcelFile(file)
            drops = ['Параметры', 'Результаты', 'Статистика']
            sheets = data.sheet_names
            for drop in drops:
                sheets = [value for value in sheets if value != drop]

            # draw loops in one plot
            flag = ''
            list_of_steps = list()
            perecardium_set = list()
            for sheet in sheets:
                temp = sheet.split('-')

                temp_float = list()
                for v in temp:
                    try:
                        temp_float.append(float(v.translate(str().maketrans(',', '.'))))
                    except:
                        pass
                if len(temp_float) == 2:
                    if flag == '':
                        flag = 'poly'
                    list_of_steps.append(sheet.split('-')[0])
                    temp_float = list()
                elif len(temp_float) == 4:
                    if flag == '':
                        flag = 'perecardium'
                    perecardium_set.append(sheet)
                    list_of_steps.append(sheet.split('-')[2])
                    temp_float = list()

            list_of_steps = np.array(list_of_steps)
            perecardium_set = np.array(perecardium_set)
            unique_steps = np.unique(list_of_steps)
            unique_steps = [int(x) for x in unique_steps]
            unique_steps.sort()
            unique_steps = [str(x) for x in unique_steps]
            for step in unique_steps:
                if flag == 'poly':
                    temp_sheet = list()
                    indexes = np.argwhere(list_of_steps == step)
                    range_of_cycles = np.arange(1, len(indexes) + 1)
                    for v in range_of_cycles:
                        temp_sheet.append(str(step) + '-' + str(v))
                else:
                    res = list()
                    for v in range(len(perecardium_set)):
                        if re.search('-' + step + '-', perecardium_set[v]):
                            res.append(True)
                        else:
                            res.append(False)
                    temp_sheet = perecardium_set[res]
                curve = np.array([[0], [0]], dtype='float64')
                is_saved = False
                for t_sheet in temp_sheet:
                    try:
                        print(colors.OKGREEN + f'proc: {t_sheet}' + colors.ENDC, end=' > ')
                        data_loaded = pd.read_excel(file, sheet_name=t_sheet)
                        elongation = np.array(data_loaded.to_numpy()[2:, 0], dtype='float64')
                        stress = np.array(data_loaded.to_numpy()[2:, 1], dtype='float64')
                        # stress[stress < np.max(stress) * trim_coeff] = 0
                        if stress[0] > 0:
                            stress -= np.average(stress[0:100])
                        stress[stress < trim_coeff] = 0
                        top_side, bot_side = split_arrays(elongation, stress)
                        try:
                            last = np.argwhere(top_side[1] == 0)[-1]
                            top_side[1][:last[0]] = 0
                        except:
                            pass
                        if not is_saved:
                            curve = top_side
                            is_saved = True
                        else:
                            curve = np.hstack((curve, top_side))

                        curve = np.hstack((curve, bot_side))
                        print(colors.OKGREEN + f'done' + colors.ENDC, end='\n')
                    except:
                        print(colors.FAIL + f'failed' + colors.ENDC, end='\n')
                        pass
                ind_to_delete = np.argwhere(curve[1, :] == 0)
                curve = np.delete(curve, ind_to_delete, axis=1)
                try:
                    if not os.path.exists('./pics/scatter'):
                        os.makedirs('./pics/scatter')
                    if not os.path.exists('./pics/line'):
                        os.makedirs('./pics/line')
                    f = plt.figure(figsize=(15, 10))
                    mpl.rcParams['axes.spines.right'] = False
                    mpl.rcParams['axes.spines.top'] = False
                    plt.plot(curve[0], curve[1], 'c-', linewidth=2)
                    plt.axhline(y=np.max(curve[1]), c='grey', linestyle='--')
                    plt.axvline(x=np.max(curve[0]), c='grey', linestyle='--')
                    plt.xlim(xmin=min(0, min(curve[0])), xmax=np.max(curve[0] * 1.1))
                    plt.ylim(0, np.max(curve[1] * 1.1))
                    plt.text(x=np.max(curve[0]), y=np.max(curve[1]), s=f'max stress: {np.max(curve[1]):0.4}',
                             horizontalalignment='right', verticalalignment='bottom')
                    plt.title(
                        f'Hysteresys of {file.split("/")[-1].split(".")[0].split(" ")[-2]} {file.split("/")[-1].split(".")[0].split(" ")[-1]} '
                        f'with elongation of {float(step.translate(str().maketrans(",", ".")))}%')
                    plt.savefig(
                        f'./pics/line/{file_name}/{file_name}_{float(step.translate(str().maketrans(",", ".")))}.png')
                    plt.close()

                    f = plt.figure(figsize=(15, 10))
                    mpl.rcParams['axes.spines.right'] = False
                    mpl.rcParams['axes.spines.top'] = False
                    plt.scatter(curve[0], curve[1], c='c', marker='.', s=2)
                    plt.axhline(y=np.max(curve[1]), c='grey', linestyle='--')
                    plt.axvline(x=np.max(curve[0]), c='grey', linestyle='--')
                    plt.xlim(xmin=min(0, min(curve[0])), xmax=np.max(curve[0] * 1.1))
                    plt.ylim(0, np.max(curve[1] * 1.1))
                    plt.text(x=np.max(curve[0]), y=np.max(curve[1]), s=f'max stress: {np.max(curve[1]):0.4}',
                             horizontalalignment='right', verticalalignment='bottom')
                    plt.title(
                        f'Hysteresys of {file.split("/")[-1].split(".")[0].split(" ")[-2]} {file.split("/")[-1].split(".")[0].split(" ")[-1]} '
                        f'with elongation of {float(step.translate(str().maketrans(",", ".")))}%')
                    plt.savefig(
                        f'./pics/scatter/{file_name}/{file_name}_{float(step.translate(str().maketrans(",", ".")))}.png')
                    plt.close()
                except:
                    pass

            # del curve, f, top_side, bot_side, temp

            # calc square and draw one-by-one
            for sheet in sheets:
                if not os.path.exists('./pics/one-by-one'):
                    os.makedirs('./pics/one-by-one')
                data_loaded = pd.read_excel(file, sheet_name=sheet)
                if data_loaded.to_numpy()[np.argmax(data_loaded.to_numpy()[2:, 1]), 0] < np.max(data_loaded.to_numpy()[2:, 0]) - 5:
                    print(colors.FAIL + f'Failed with {file_name} -> {sheet}! '
                                        f'{data_loaded.to_numpy()[np.argmax(data_loaded.to_numpy()[2:, 1]), 0]} but limit '
                                        f'{np.max(data_loaded.to_numpy()[2:, 0]) - 5}' + colors.ENDC)
                    is_break = True
                    break
                elongation = np.array(data_loaded.to_numpy()[2:, 0], dtype='float64')
                stress = np.array(data_loaded.to_numpy()[2:, 1], dtype='float64')
                # stress[stress < np.max(stress) * trim_coeff] = 0
                if stress[0] > 0:
                    stress -= np.average(stress[0:100])
                stress[stress < trim_coeff] = 0
                top_side, bot_side = split_arrays(elongation, stress)
                try:
                    last = np.argwhere(top_side[1] == 0)[-1]
                    top_side[1][:last[0]] = 0
                except:
                    pass
                if sheet[-1] == '1':
                    saved_elongation_v1 = top_side[0][-1]
                    saved_stress_v1 = top_side[1][-1]

                square_top = compute_square(arr_e=top_side[0], arr_s=top_side[1])
                square_bot = compute_square(arr_e=bot_side[0], arr_s=bot_side[1])
                f = plt.figure(figsize=(15, 10))
                mpl.rcParams['axes.spines.right'] = False
                mpl.rcParams['axes.spines.top'] = False
                plt.plot(top_side[0], top_side[1], 'r-')
                plt.text(x=top_side[0][-1], y=top_side[1][-1],
                         s=(f'el: {top_side[0][-1]:2.2f}\ns: {top_side[1][-1]:2.2f}'),
                         horizontalalignment='right', verticalalignment='bottom')
                plt.plot(top_side[0][-1], top_side[1][-1], 'or')

                plt.plot(bot_side[0], bot_side[1], 'c-')
                plt.text(x=bot_side[0][-1], y=bot_side[1][-1],
                         s=(f'el: {bot_side[0][-1]:2.2f}\ns: {bot_side[1][-1]:2.2f}'),
                         horizontalalignment='right', verticalalignment='bottom')
                plt.plot(bot_side[0][-1], bot_side[1][-1], 'or')
                top_loop_start = [0, 0]
                try:
                    arg = np.argwhere(top_side[1] == 0)[-1]
                    top_loop_start[0] = top_side[0][arg][0]
                    top_loop_start[1] = top_side[1][arg][0]
                except:
                    pass
                plt.plot(top_loop_start[0], top_loop_start[1], 'xk')
                plt.text(x=top_loop_start[0], y=top_loop_start[1],
                         s=(f'el: {top_loop_start[0]:2.2f}\ns: {top_loop_start[1]:2.2f}'),
                         horizontalalignment='right', verticalalignment='bottom')
                plt.title(
                    f'top: {float(square_top):2.2} |'
                    f' bot: {float(square_bot):2.2f} |'
                    f' sq: {float(square_top - square_bot):2.2f}'
                )

                plt.savefig(f'./pics/one-by-one/{file_name}/' + file_name + '_' + sheet + '.jpg')
                plt.close()

                if sheet[-1] == '1':
                    row = [
                        time.ctime(os.path.getctime(file)),
                        file_name,
                        sheet.split('-')[0],
                        sheet.split('-')[1],
                        square_top - square_bot,
                        0,
                        0,
                        bot_side[0][-1] - top_loop_start[0],
                        top_loop_start[0],
                        0,
                        0,
                        np.max(top_side[1])
                    ]
                else:
                    row = [
                        time.ctime(os.path.getctime(file)),
                        file_name,
                        sheet.split('-')[0],
                        sheet.split('-')[1],
                        square_top - square_bot,
                        top_side[0][-1] - saved_elongation_v1,
                        top_side[0][-1] - saved_elongation,
                        bot_side[0][-1] - top_loop_start[0],
                        top_loop_start[0],
                        top_side[1][-1] - saved_stress_v1,
                        top_side[1][-1] - saved_stress,
                        np.max(top_side[1])
                    ]
                saved_elongation = top_side[0][-1]
                saved_stress = top_side[1][-1]
                worksheet.append(row)
                workbook.save('./res.xlsx')
            # if is_break:
            #     break

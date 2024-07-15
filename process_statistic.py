import pandas as pd
import numpy as np
import openpyxl

if __name__ == '__main__':
    xlsx_files = [
        'res.xlsx'
    ]
    row = ['Файл', 'Материал', 'Шаг', 'Площадь 1 петли', 'Avg площадь 3 последних',
           'Необратимая деформация 1 петли, %', 'Необратимая деформация 3 последних, %',
           'Avg расстояние в петле, %', 'Avg расстояние в 3 последних, %',
           'Avg потеря напряжения, МПа', 'Avg потеря напряжения 3 последних, МПа',
           'Avg напряжение, МПа', 'Avg напряжение 3 последних, МПа', 'Макс напряжение, МПа']
    res = pd.DataFrame(columns=row)
    res.to_excel('./res_stat.xlsx', startrow=0, startcol=-1, sheet_name='statistic')
    workbook = openpyxl.load_workbook('./res_stat.xlsx')
    worksheet = workbook.active
    # xlsx_file = xlsx_files[0]
    for xlsx_file in xlsx_files:
        data = pd.read_excel(xlsx_file)
        steps = data['Шаг'].values
        unique_steps = np.unique(steps)
        np_data = data.to_numpy()
        materials = np.unique(np_data[:, 1])
        for mat in materials:
            curr = np_data[np_data[:, 1] == mat, :]
            try:
                for step in unique_steps:
                    curr_data = curr[curr[:, 2] == step, 1:]
                    mat_name = ""
                    for v in mat.split(" ")[2:]:
                        mat_name += f'{v} '
                    row = [xlsx_file.split(".")[0], mat_name, step, curr_data[0, 3], np.average(curr_data[-3:, 3]),
                           curr_data[0, 7], np.average(curr_data[-3:, 7]),
                           np.average(curr_data[:, 7]), np.average(curr_data[-3:, 7]),
                           np.average(curr_data[:, -2]), np.average(curr_data[-3:, -2]),
                           np.average(curr_data[:, -1]), np.average(curr_data[-3:, -1]), np.max(curr_data[:, -1])
                           ]
                    worksheet.append(row)
                    workbook.save('./res_stat.xlsx')
            except:
                pass
    print('hi')
    workbook.close()

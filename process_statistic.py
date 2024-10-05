import pandas as pd
import numpy as np
import openpyxl

if __name__ == '__main__':
    xlsx_files = [
        'res.xlsx'
    ]
    file_name_end = 'res_stat_one_by_one.xlsx'
    row = ['Материал', 'Группа материалов', 'Шаг', 'Площадь 1 петли',
           'Необратимая деформация 1 петли, %', 'Необратимая деформация 3 последних, %',
           'Avg расстояние в петле, %', 'Avg расстояние в 3 последних, %',
           'Avg потеря напряжения, МПа', 'Avg потеря напряжения 3 последних, МПа',
           'Avg напряжение, МПа', 'Avg напряжение 3 последних, МПа', 'Макс напряжение, МПа',
           'Emod нагрузки, 3 последних', 'Emod разгрузки, 3 последних']

    res = pd.DataFrame(columns=row)
    res.to_excel(f'./{file_name_end}', startrow=0, startcol=-1, sheet_name='statistic')
    workbook = openpyxl.load_workbook(f'./{file_name_end}')
    worksheet = workbook.active
    for xlsx_file in xlsx_files:
        data = pd.read_excel(xlsx_file)
        steps = data['Шаг'].values
        unique_steps = np.unique(steps)
        np_data = data.to_numpy()
        materials = [
            '15% PVA-L', '15% PVA 0,5% CNT', '15% PVA 0,5% CNT-L',
            'SIBS 188k 1% SCNT',
            'МедЛаб-КТ', 'ePTFE',
            'PSU-FGO',
            'Перикард вдоль', 'Перикард поперек', 'Модифицированный Перикард вдоль', 'Модифицированный Перикард поперек'
        ]
        materials_types = [
            'PVA', 'PVA', 'PVA',
            'SIBS',
            'ePTFE', 'ePTFE',
            'PSU-FGO',
            'Перикард', 'Перикард', 'Перикард', 'Перикард'
        ]
        for mat, mat_type in zip(materials, materials_types):
            rows_with_mat = []
            for curr_mat_ind in np.arange(0, len(np_data[:, 1])):
                if mat.lower() in str(np_data[curr_mat_ind, 1]).lower():
                    rows_with_mat.append(curr_mat_ind)
            try:
                curr = np_data[rows_with_mat, :]
                for step in unique_steps:
                    curr_data = curr[curr[:, 2] == step, 1:]
                    # mat_name = ""
                    # for v in mat.split(" ")[2:]:
                    #     mat_name += f'{v} '

                    data_1st_loop = curr_data[curr_data[:, 2] == 1, :]
                    data_2nd_loop = curr_data[curr_data[:, 2] == 2, :]
                    data_3rd_loop = curr_data[curr_data[:, 2] == 3, :]
                    data_4th_loop = curr_data[curr_data[:, 2] == 4, :]
                    data_5th_loop = curr_data[curr_data[:, 2] == 5, :]
                    data_6th_loop = curr_data[curr_data[:, 2] == 6, :]
                    data_7th_loop = curr_data[curr_data[:, 2] == 7, :]
                    data_8th_loop = curr_data[curr_data[:, 2] == 8, :]
                    square_1st_loop = np.average(data_1st_loop[:, 3])
                    avg_square_3_last_loops = np.average([data_8th_loop[:, 3], data_7th_loop[:, 3], data_6th_loop[:, 3]])
                    saved_deformation = np.average(data_1st_loop[:, 7])
                    avg_saved_deformation_3_last = np.average([data_8th_loop[:, 7], data_7th_loop[:, 7], data_6th_loop[:, 7]])
                    avg_loop_distance = np.average(
                        [data_8th_loop[:, 8],
                         data_7th_loop[:, 8],
                         data_6th_loop[:, 8],
                         data_5th_loop[:, 8],
                         data_4th_loop[:, 8],
                         data_3rd_loop[:, 8],
                         data_2nd_loop[:, 8],
                         data_1st_loop[:, 8]
                         ]
                    )
                    avg_loop_distance_3_last = np.average(
                        [data_8th_loop[:, 8],
                         data_7th_loop[:, 8],
                         data_6th_loop[:, 8]
                         ]
                    )
                    avg_stress_loss = np.average(
                        [data_8th_loop[:, -4],
                         data_7th_loop[:, -4],
                         data_6th_loop[:, -4],
                         data_5th_loop[:, -4],
                         data_4th_loop[:, -4],
                         data_3rd_loop[:, -4],
                         data_2nd_loop[:, -4],
                         data_1st_loop[:, -4]
                         ]
                    )
                    avg_stress_loss_3_last = np.average(
                        [data_8th_loop[:, -4],
                         data_7th_loop[:, -4],
                         data_6th_loop[:, -4]
                         ]
                    )
                    avg_stress = np.average(
                        [data_8th_loop[:, -3],
                         data_7th_loop[:, -3],
                         data_6th_loop[:, -3],
                         data_5th_loop[:, -3],
                         data_4th_loop[:, -3],
                         data_3rd_loop[:, -3],
                         data_2nd_loop[:, -3],
                         data_1st_loop[:, -3]
                         ]
                    )
                    avg_stress_3_last = np.average(
                        [data_8th_loop[:, -3],
                         data_7th_loop[:, -3],
                         data_6th_loop[:, -3]
                         ]
                    )
                    max_stress = np.max(
                        [data_8th_loop[:, -3],
                         data_7th_loop[:, -3],
                         data_6th_loop[:, -3],
                         data_5th_loop[:, -3],
                         data_4th_loop[:, -3],
                         data_3rd_loop[:, -3],
                         data_2nd_loop[:, -3],
                         data_1st_loop[:, -3]
                         ]
                    )
                    emod_load_3_last = np.average(
                        [data_8th_loop[:, -2],
                         data_7th_loop[:, -2],
                         data_6th_loop[:, -2]
                        ]
                    )
                    emod_unload_3_last = np.average(
                        [data_8th_loop[:, -1],
                         data_7th_loop[:, -1],
                         data_6th_loop[:, -1]
                         ]
                    )
                    row = [mat, mat_type, step, square_1st_loop,
                           saved_deformation, avg_saved_deformation_3_last,
                           avg_loop_distance, avg_loop_distance_3_last,
                           avg_stress_loss, avg_stress_loss_3_last,
                           avg_stress, avg_stress_3_last, max_stress, emod_load_3_last,  emod_unload_3_last
                           ]
                    worksheet.append(row)
                    workbook.save(f'./{file_name_end}')
            except:
                pass
    print('hi')
    workbook.close()

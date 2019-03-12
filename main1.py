#Программа чтения и обработки данных с логгера
#Синчинов Никита
from typing import List

import xlrd
import matplotlib.pyplot as plt


exel_data_file = xlrd.open_workbook('test_data_2.xlsx')
sheet = exel_data_file.sheet_by_index(0)

row_nubmer = sheet.nrows
cols_number = sheet.ncols

#Номера столбцов напряжение тока и Времени
voltage_col = 0
current_col = 1
date_col = 3

logg_date = []
logg_voltage = []
logg_current = []

# Берем дату и время записи
for row in range(1, row_nubmer):
    logg_date.append(sheet.cell_value(row, date_col))
    logg_voltage.append(sheet.cell_value(row, voltage_col))
    logg_current.append(sheet.cell_value(row, current_col))
    #"Перевернем" ток
    logg_current[row-1] = 1024 - logg_current[row-1]

    # Если по какой-либо причине данные в ячейке имеют некорректный формат, записать в ячейку время из предыдущей + 1 минута
    if type(logg_date[row - 1]) != float:
        logg_date[row - 1] = list(logg_date[row - 2])
        logg_date[row - 1][4] = logg_date[row - 1][4] + 1
    else:
        logg_date[row - 1] = list(xlrd.xldate_as_tuple(logg_date[row - 1], exel_data_file.datemode))


#y_curr - набор значений тока от времени
#y_voltage - набор значений напряжения от времени
#length - длина (колво точек) к которому необхожимо нормализовать (привести)
#На выходе процедуры нормализованный набор точек без потери и изменения характера графика
def normalization(y_curr, y_voltage, length):
    #Так как длина вектора значний тока и напряжения одинакова, отталкиваемся от тока (y_curr), например
    if len(y_curr) > length:
        while len(y_curr) > length:
            norm_y_curr = []
            norm_y_voltage = []
            m = len(y_curr) // length
            for point in range(0, len(y_curr)-m, m):
                aver = 0
                aver2 = 0
                for subpoint in range(0, m):
                    aver = aver + y_curr[point + subpoint]
                    aver2 = aver2 + y_voltage[point + subpoint]
                norm_y_curr.append(aver / m)
                norm_y_voltage.append(aver2 / m)
            y_curr = norm_y_curr
            y_voltage = norm_y_voltage
    elif len(y_curr) < length:
        iter = 0
        while len(y_curr) < length:
            mayak = len(y_curr)
            #Добавляем по одному значению в вектор, пока он не будет желаемой длины
            #За первую итерацию между 1 и 2 значением, за вторую между 2 и 3 и т.д
            #Добавляем среднее между двумя соседними
            norm_y_curr = [0] * (len(y_curr) + 1)
            y1_curr = y_curr[iter]
            y2_curr = y_curr[iter + 1]
            norm_y_curr[iter] = y1_curr
            norm_y_curr[iter+1] = (y1_curr + y2_curr)/2
            norm_y_curr[:iter] = y_curr[:iter]
            norm_y_curr[iter+2:] = y_curr[iter+1:]
            y_curr = norm_y_curr

            norm_y_voltage = [0] * (len(y_voltage) + 1)
            y1_voltage = y_voltage[iter]
            y2_voltage = y_voltage[iter + 1]
            norm_y_voltage[iter] = y1_voltage
            norm_y_voltage[iter + 1] = (y1_voltage + y2_voltage) / 2
            norm_y_voltage[:iter] = y_voltage[:iter]
            norm_y_voltage[iter + 2:] = y_voltage[iter + 1:]
            y_voltage = norm_y_voltage

            iter += 2
            #Когда дошли до последнего элемента, обнуляем счетчик и начинаем снова
            if iter >= mayak:
                iter = 0
                mayak = len(y_curr)
    return [y_curr, y_voltage]


#Процедура поиска графиков на заданном отрезке и их нормализация
#sp - start point
#ep - end point
#qu - quantity of points - колиичетсво точек до которого мы нормализуем каждый график
def graphic_normalization(sp, ep, qu):
    #На выходе в этот массив запишутся все найденные и нормализованные данные зарядок
    mass = []
    for point in range(sp, ep):
        X = []
        Curr = []
        Volt = []
        I1 = logg_current[point]
        I2 = logg_current[point + 1]
        # Находим точку начала зарядки по значению и производной от тока
        if I2 > 600 and I2 - I1 > 20:
            print("___________________________")
            print("Начало: " + str(logg_date[point][3]) + ":" + str(logg_date[point][4]))
            for end_point in range(point, ep - 1):
                #Вектор X содержит массив данных времени для каждой зарядки (пока не используется)
                X.append(logg_date[end_point])

                Curr.append(logg_current[end_point])
                Volt.append(logg_voltage[end_point])

                I1 = logg_current[end_point]
                I2 = logg_current[end_point + 1]
                U1 = logg_current[end_point]
                U2 = logg_current[end_point + 1]

                if I2 < 600 and U2 - U1 < -30:
                    print("Конец: " + str(logg_date[end_point][3]) + ":" + str(logg_date[end_point][4]))
                    print("___________________________")
                    break
            #Добавляем в конец массива нормализованные данные тока и напряжения от новой найденной зарядки
            mass.append(normalization(Curr, Volt, qu))
    return mass

Q_points = 100
a = graphic_normalization(0, 1380, Q_points)

x = []
for i in range(0, Q_points):
    x.append(i)

#Блок вывода графиков
fig, axes = plt.subplots(nrows=2, ncols=1, figsize=(8, 8))
fig.subplots_adjust(left=0.09, bottom=0.11, right=0.95, top=0.95, hspace=0.3)
axes[0].grid(True)
axes[1].grid(True)

for num_graph in range(0, len(a[:])):
    axes[0].plot(x, a[num_graph][0])
    axes[1].plot(x, a[num_graph][1])

axes[0].set_title('Токи')
axes[1].set_title('Напряжения')
#Открыть окно с графиками

plt.show()
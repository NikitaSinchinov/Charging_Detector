#Программа чтения и обработки данных с логгера
#Синчинов Никита


import xlrd
import matplotlib.pyplot as plt


exel_data_file = xlrd.open_workbook('test_data.xlsx')
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

X = []
Y = []
k = []
Y2 = []
# Поиск пар точек
for point in range(0, 1300): #len(logg_date)-1
    I1 = logg_current[point]
    I2 = logg_current[point+1]
    # Находим точку начала зарядки по значению и производной от тока
    if I2 > 600 and I2-I1>20:
        print("")
        print("Начало: " + str(logg_date[point][3]) + ":"+ str(logg_date[point][4]))
        # Начинаем искать точку окончания зарядки
        for end_point in range(point, len(logg_date)-1):
            X.append(logg_date[end_point])
            Y.append(logg_current[end_point])
            Y2.append(logg_voltage[end_point])
            k.append(end_point)
            I1 = logg_current[end_point]
            I2 = logg_current[end_point + 1]
            U1 = logg_current[end_point]
            U2 = logg_current[end_point + 1]
            if I2 < 600 and U2 - U1 < -30:
                print("Конец: " + str(logg_date[end_point][3]) + ":" + str(logg_date[end_point][4]))
                break
print("График состоит из ", len(X), " наборов.")

mY = []
mY2 = []
mk = []
for point in range(0, len(Y)//2-1):
    mY.append(Y[point*2])
    mY2.append(Y2[point*2])
    mk.append(point)


average_Y = []
average_Y2 = []
average_k = []
for point in range(0, len(Y)//2-2):
    average_Y.append((Y[point*2] + Y[point*2+1])/2)
    average_Y2.append((Y2[point*2] + Y2[point*2+1])/2)
    average_k.append(point)





#Настройки окна с графиками
fig, axes = plt.subplots(nrows=3, ncols=1, figsize=(8, 8))
fig.subplots_adjust(left=0.09, bottom=0.11, right=0.95, top=0.95, hspace=0.3)
axes[0].grid(True)
axes[1].grid(True)
axes[2].grid(True)
axes[0].plot(k, Y2, k, Y)
axes[1].plot(mk, mY2, mk, mY)
axes[2].plot(average_k, average_Y2, average_k, average_Y)
axes[0].legend(('Напряжение', 'Ток'))


axes[0].set_title('График как он есть')
axes[1].set_title('Каждая вторая точка')
axes[2].set_title('Среднее значение между двумя соседними')
#Открыть окно с графиками
plt.show()
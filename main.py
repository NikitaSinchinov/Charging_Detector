import xlrd

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


for point in range(0, len(logg_date)-1):
    I1 = logg_current[point]
    I2 = logg_current[point+1]
    if I2 > 600 and I2-I1>20:
        print("")
        print(str(logg_date[point][3]) + ":"+ str(logg_date[point][4]))

        for end_point in range(point, len(logg_date)-1):
            I1 = logg_current[end_point]
            I2 = logg_current[end_point + 1]
            U1 = logg_current[end_point]
            U2 = logg_current[end_point + 1]
            if I2 < 600 and U2 - U1 < -30:
                print(str(logg_date[end_point][3]) + ":" + str(logg_date[end_point][4]))
                break


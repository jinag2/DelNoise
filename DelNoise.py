#! python3

import csv
import sys
import os
import copy
from openpyxl import Workbook
from openpyxl import styles
from openpyxl import chart
from openpyxl import drawing
from openpyxl.drawing import line
from openpyxl.drawing import colors


def main():
    csv_list = read_csv(sys.argv[1])
    vol_list = abstract_float_col(csv_list, 0)
    adc1_list = abstract_int_col(csv_list, 1)
    adc2_list = abstract_int_col(csv_list, 2)
    adc3_list = abstract_int_col(csv_list, 3)
    adc4_list = abstract_int_col(csv_list, 4)
    adc5_list = abstract_int_col(csv_list, 5)
    
    # csv_out = cut_flag(adc_list)
    # output_file = write_csv(csv_out)
    # print('Creat ' + output_file + ' OK!')
    
    data_list = [vol_list, adc1_list, adc2_list]
    excel_file = 'D:\\CVSROOT\\Python\\ADC1_out.xlsx'
    if create_excel(excel_file, data_list):
        print('Create ' + excel_file + ' succeeded!')
    else:
        print('Create ' + excel_file + ' failed!')


def read_csv(filename):
    csv_file = open(filename)
    csv_reader = csv.reader(csv_file)
    return list(csv_reader)


def abstract_float_col(data_list, col):
    value_list = []
    for row in data_list:
        value_list.append(float(row[col]))
    return value_list


def abstract_int_col(data_list, col):
    value_list = []
    for row in data_list:
        value_list.append(int(row[col]))
    return value_list


def cut_flag(data_list):
    flag_size = len(data_list[0])
    new_list = copy.deepcopy(data_list)
    for row in new_list:
        del row[flag_size - 1]
        del row[flag_size - 2]
        del row[flag_size - 3]
    return new_list


def write_csv(data_list):
    input_file_tuple = os.path.split(sys.argv[1])
    file_tuple = os.path.splitext(input_file_tuple[1])
    filename = input_file_tuple[0] + '\\' + file_tuple[0] + '_noflag' + file_tuple[1]
    
    csv_file = open(filename, 'w', newline = '')
    csv_writer = csv.writer(csv_file)
    for row in data_list:
        csv_writer.writerow(row)
    csv_file.close()
    
    return filename


def create_excel(filename, data_list):
    if len(data_list) < 2:
        return False
    
    wb = Workbook()
    sheet = wb.active
    font12 = styles.Font(size = 12)

    serial_title = ['Voltage']
    adc_num = len(data_list) - 1
    for i in range(adc_num):
        serial_title.append('ADC ' + str(i))
    sheet.append(serial_title)
    
    for col in range(len(data_list)):
        for row in range(len(data_list[col])):
            cell = sheet.cell(row = row + 2, column = col + 1, value = data_list[col][row])
            cell.font = font12
    
    adc_chart = chart.ScatterChart()
    adc_chart.title = "Scatter Chart"
    adc_chart.style = 13
    adc_chart.y_axis.title = 'ADC'
    adc_chart.x_axis.title = 'Voltage'

    x_values = chart.Reference(sheet, min_col = 1, min_row = 2, max_row = 100)
    for i in range(adc_num):
        values = chart.Reference(sheet, min_col = i+2, min_row = 2, max_row = 100)
        series = chart.Series(values, x_values, title_from_data = True)
        adc_chart.series.append(series)

    adc_series = adc_chart.series[0]
    # fill = PatternFillProperties(prst = "pct5")
    # fill.foreground = ColorChoice(prstClr = "red")
    # fill.background = ColorChoice(prstClr = "blue")
    
    line_prop = drawing.line.LineProperties(solidFill = drawing.colors.ColorChoice(prstClr = 'red'))
    adc_series.graphicalProperties.line = line_prop
    
    sheet.add_chart(adc_chart, "F5")
    
    # serial_obj = openpyxl.chart.series_factory.SeriesFactory(data, title='First series of values')
    # chart_obj = openpyxl.chart.BarChart()
    # chart_obj.title = 'My Chart'
    # chart_obj.append(serial_obj)
    # sheet.add_chart(chart_obj, 'C5')
    
    wb.save(filename)
    return True


main()

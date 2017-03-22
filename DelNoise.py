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


MAX_ADC_NUM = 10


def main():
    csv_list = read_csv(sys.argv[1])
    data_list = [abstract_float_col(csv_list, 0)]
    
    if len(csv_list[0]) - 4 > MAX_ADC_NUM:
        adc_num = MAX_ADC_NUM
    else:
        adc_num = len(csv_list[0]) - 4
    
    total_adc_list1 = []
    if adc_num >= 1:
        for i in range(adc_num):
            adc_list = abstract_int_col(csv_list, i+1)
            total_adc_list1.append(adc_list)
            data_list.append(adc_list)
    else:
        print('No ADC data.')
        sys.exit()
        
    # csv_out = cut_flag(adc_list)
    # output_file = write_csv(csv_out)
    # print('Creat ' + output_file + ' OK!')

    data_list.append(cal_median(total_adc_list1))
    excel_file = generate_excel_file(sys.argv[1]) #'D:\\CVSROOT\\Python\\ADC1_out.xlsx'
    if create_excel(excel_file, data_list):
        print('Create ' + excel_file + ' succeeded!')
    else:
        print('Create ' + excel_file + ' failed!')


def read_csv(filename):
    try:
        csv_file = open(filename)
    except Exception as err:
        print('Can not open ' + ' (' + str(err) + ')')
        sys.exit()
        
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


def cal_median(adc_list):
    adc_num = len(adc_list)
    odd_num = True
    odd_index = int(adc_num // 2)
    if (adc_num % 2) == 0:
        odd_num = False
    
    median_list = []
    elem_num = len(adc_list[0])
    for i in range(elem_num):
        elem_list = []
        for j in range(adc_num):
            elem_list.append(adc_list[j][i])
        
        elem_list.sort()
        if odd_num:
            median_list.append(elem_list[odd_index])
        else:
            elem = int((elem_list[odd_index - 1] + elem_list[odd_index]) / 2)
            median_list.append(elem)
    
    return median_list


def generate_excel_file(csv_filename):
    input_file_tuple = os.path.split(csv_filename)
    file_tuple = os.path.splitext(input_file_tuple[1])
    filename = input_file_tuple[0] + '\\' + file_tuple[0] + '_out.xlsx'
    return filename

    
def create_excel(filename, data_list):
    if len(data_list) < 2:
        return False
    
    wb = Workbook()
    sheet = wb.active
    font12 = styles.Font(size = 12)
    sheet.column_dimensions['A'].width = 11

    serial_title = ['Voltage']
    serials_num = len(data_list) - 1
    adc_num = serials_num - 1
    for i in range(adc_num):
        serial_title.append('ADC ' + str(i))
    serial_title.append('Median')
    sheet.append(serial_title)
    
    for col_idx in range(len(data_list)):
        for row_idx in range(len(data_list[col_idx])):
            cell = sheet.cell(row = row_idx + 2, column = col_idx + 1, value = data_list[col_idx][row_idx])
            cell.font = font12
    
    adc_chart = chart.ScatterChart()
    adc_chart.title = "Scatter Chart"
    adc_chart.style = 13
    adc_chart.y_axis.title = 'ADC'
    adc_chart.x_axis.title = 'Voltage'
    adc_chart.height = 12
    adc_chart.width = 16

    max_item_num = 200
    line_color = ('416FA6', 'A8423F', '86A44A', '6E548D', '3D96AE',
                  'B8860B', 'E9967A', 'DA8137', '8EA5CB', '808000',
                  'A0522D', '2E8B57', 'B0E0E6', '000080', 'FFDEAD',
                  'FF0000')

    x_values = chart.Reference(sheet, min_col = 1, min_row = 2, max_row = max_item_num+1)
    for i in range(serials_num):
        values = chart.Reference(sheet, min_col = i + 2, min_row = 1, max_row = max_item_num+1)
        series = chart.Series(values, x_values, title_from_data = True)
        if i == serials_num - 1:
            series.graphicalProperties.line = drawing.line.LineProperties(solidFill = line_color[15])
        else:
            series.graphicalProperties.line = drawing.line.LineProperties(solidFill = line_color[i])
        series.graphicalProperties.line.width = 27432  # width in EMUs, EMU = pixel * 914400 / 96, assume pixel = 75
        adc_chart.series.append(series)
    
    sheet.add_chart(adc_chart, "I5")
    
    try:
        wb.save(filename)
    except Exception as err:
        print('Can not save ' + ' (' + str(err) + ')')
        sys.exit()
        
    return True


main()

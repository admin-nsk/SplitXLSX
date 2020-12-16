#!/usr/bin/python3
# -*- coding: utf-8 -*-

import pandas as pd
from logger_config import get_logger

log = get_logger('Slice_XLS_logger')
path = 'data/2020.11.24 Выгрузка УК ноябрь.xlsx'


class SliceXLS:

    def __init__(self, input_xls):
        '''python 3.9
        Class for split one file to many
        :param input_xls: File for splitting
        '''
        self.input_xls = input_xls
        self.all_data_frame = pd.DataFrame()
        self.column_names = ''
        self.sheet_names = ''

    def read_sheets(self):
        '''
        Open Excel file and read sheet names
        :return: list sheet names
        '''
        try:
            with pd.ExcelFile(self.input_xls) as source_xls:
                self.sheet_names = source_xls.sheet_names
                return self.sheet_names
        except:
            log.debug('Don\'t read file')
            return 'Don\'t read file'

    def read_data(self, selected_sheet):
        '''
        Read data frame on file
        :param selected_sheet:
        :return: list column_names
        '''
        with pd.ExcelFile(self.input_xls) as source_xls:
            self.all_data_frame = pd.read_excel(source_xls, sheet_name=selected_sheet)
            self.column_names = self.all_data_frame.columns
            return self.column_names

    def split_file(self, selected_sheet, selected_column):
        '''
        Split data to many file by selected column
        :param selected_sheet:
        :param selected_column:
        :return: None
        '''
        log.debug('Start split_file')
        unique_data_on_column = set(self.all_data_frame[selected_column])
        for item in unique_data_on_column:
            filtered_data = self.all_data_frame[(self.all_data_frame[selected_column] == item)]
            out_file = item.replace('/', '_') + '.xlsx'
            try:
                log.debug('Write frame to file')
                self.write_xls(out_file, filtered_data, selected_sheet)
            except:
                log.debug('Don\'t write frame to file')

    def write_xls(self, filename, frame, sheet_name):
        '''
        Write data frame in file
        :param filename:
        :param frame:
        :param sheet_name:
        :return:
        '''
        out_data = pd.DataFrame()
        out_data = out_data.append(frame, ignore_index=True)
        writer = pd.ExcelWriter(filename)
        out_data.to_excel(writer, sheet_name=sheet_name)
        writer.save()


if __name__ == '__main__':
    slxl = SliceXLS(path)
    slxl.read_data(selected_sheet='Телеметрия')
    slxl.split_file(selected_sheet='Телеметрия', selected_column='Адрес')

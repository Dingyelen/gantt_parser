# -*- coding: utf-8 -*-
# @Author: dingyelen
# @Date:   2024-08-27 11:08:06
# @Last Modified by:   dingyelen
# @Last Modified time: 2024-08-27 14:37:00


import pandas as pd
# from datetime import date, datetime
from data_handling import ReadExcelData


class GanttData(object):
    def __init__(self, file_path: str, sheetname: str, date_row: int, data_only: bool = True):
        """[summary]
        
        [description]
        
        Args:
            file_path (str): [文件地址]
            data_only (bool): [空值不要读取] (default: `True`)
        """

        self.file_path = file_path
        self.data_only = data_only
        self.sheetname = sheetname
        self.date_row = date_row

    def task_to_date(self):
        data = ReadExcelData(file_path=self.file_path, data_only=self.data_only)
        sheet = data.open_sheet(sheetname=self.sheetname)
        mergecell_list = data.get_mergecell(sheetname=self.sheetname)

        gantt_task_data = []

        for _mergecell in mergecell_list:
            start_row = _mergecell[0]
            start_col = _mergecell[2]
            end_col = _mergecell[3]
            task_name = sheet.cell(row=start_row, column=start_col).value
            # task_name = data.get_cell_value(sheetname=sheetname, row=start_row, column=start_col)
            start_date = sheet.cell(row=self.date_row, column=start_col).value
            end_date = sheet.cell(row=self.date_row, column=end_col).value
            if start_date and end_date:
                duration = (end_date - start_date).days + 1
                # task_data【活动名，开始日期，结束日期，天数】
                task_data = tuple([task_name, start_date, end_date, duration])  
            gantt_task_data.append(task_data)    

        return gantt_task_data

    def date_to_task(self, start_date, end_date):
        start_date = pd.to_datetime(start_date)
        end_date = pd.to_datetime(end_date)

        gantt_task_data = self.task_to_date()
        gantt_task_data_df = pd.DataFrame(gantt_task_data, columns=['task_name', 'start_date', 'end_date', 'duration'])
        gantt_task_data_df['start_date_fit'] = gantt_task_data_df['start_date'].apply(lambda x: x if x>=start_date else start_date)
        gantt_task_data_df['end_date_fit'] = gantt_task_data_df['end_date'].apply(lambda x: x if x<=end_date else end_date)
        gantt_task_data_df['duration_fit'] = gantt_task_data_df.apply(lambda x: (x.end_date_fit-x.start_date_fit).days+1, axis=1)

        gantt_task_data_select = gantt_task_data_df[(gantt_task_data_df.end_date>=start_date)&(gantt_task_data_df.start_date<=end_date)]
        gantt_task_data_select = gantt_task_data_select.loc[:, ['task_name', 'start_date_fit', 'end_date_fit', 'duration_fit']]
        # gantt_date_data【活动名，开始日期，结束日期，天数】
        gantt_date_data = [tuple(row) for row in gantt_task_data_select.to_numpy()]
        
        return gantt_date_data


if __name__ == '__main__':
    file_path = r'D:\我的坚果云\work\Sincetimes_DOW\2023-09-21 XCY DYL DOW damo 活动历史对比\damo活动\【日本】霸王天下-活动表2020年.xlsx'
    test = GanttData(file_path, '活动排期', 4)

    # print(test.task_to_date())
    test.date_to_task('2020-07-01', '2020-07-07')
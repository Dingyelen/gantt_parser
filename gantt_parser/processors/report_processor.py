# -*- coding: utf-8 -*-
# @Author: dingyelen
# @Date:   2024-08-28 10:18:34
# @Last Modified by:   dingyelen
# @Last Modified time: 2024-08-28 15:03:33

import os
import pandas as pd
import plotly.io as pio
from plotly.figure_factory import create_gantt

# from processors.gantt_processor import GanttData
from gantt_parser.processors.gantt_processor import GanttData
from gantt_parser.untils.date_ranges import DateFuncNew


class GanttPlot(object):
    def __init__(self, gantt_data: list, task_col: str='task_name', start_col: str='start_date', end_col: str='end_date', 
                 output_path: str = r'D:\我的坚果云\work\GIT\gantt_parser\data\output\graph'):
        self.gantt_data = gantt_data
        self.task_col = task_col
        self.start_col = start_col
        self.end_col = end_col
        self.output_path = output_path

    def gantt_data_process(self):
        select_col = [self.task_col, self.start_col, self.end_col]
        gantt_data = [{key: value for key, value in task.items() if key in select_col} for task in self.gantt_data]
        gantt_data = [{'Task' if key == self.task_col else 'Start' if key == self.start_col else 'Finish': value for key, value in task.items()} for task in gantt_data]

        return gantt_data

    def gantt_plot(self, graph_name: str='gantt_graph.png'):
        fig_path = os.path.join(self.output_path, graph_name)
        gantt_data = self.gantt_data_process()
        fig = create_gantt(gantt_data, colors='Blues', group_tasks=True, 
                           show_colorbar=True, bar_width=0.5, 
                           showgrid_x=True, showgrid_y=True, )
        fig.update_xaxes(tickmode='linear', tickformat='%m-%d')
        fig.update_layout(showlegend=False)
        pio.write_image(fig, fig_path, format='png')

        return fig_path

class GanttReport(object):
    def __init__(self, gantt_data: list, output_path: str = r'D:\我的坚果云\work\GIT\gantt_parser\data\output'):
        self.gantt_data = gantt_data
        self.output_path = output_path

    def gantt_data_analysis(self):
        gantt_data_df = pd.DataFrame(self.gantt_data)
        date_range = DateFuncNew('day')
        gantt_data_df['日期'] = gantt_data_df.apply(lambda x: date_range.get_start_and_end_range(x.start_date.strftime('%Y-%m-%d'), x.end_date.strftime('%Y-%m-%d')), axis=1)
        gantt_date_df_explode = gantt_data_df.explode('日期')
        gantt_date_group = gantt_date_df_explode.groupby(['日期'])['task_name'].count().rename('活动数').reset_index()
        
        task_num = len(gantt_data_df)
        start_date = gantt_data_df['start_date'].max().strftime('%Y%m%d')
        end_date = gantt_data_df['end_date'].max().strftime('%Y%m%d')
        report_path = os.path.join(self.output_path, f'{start_date}-{end_date} 活动情况.txt')

        report = f'任务清单\n'
        report += f'{start_date}-{end_date} 期间一共有 {task_num} 个任务\n'
        report += "=" * 30 + '\n\n'
        report += '每日任务数量清单\n'
        report += f'{gantt_date_group.to_string(index=False)}\n'
        report += "=" * 30 + '\n\n'
        report += '活动详情如下\n'

        for task in gantt_data:
            task_name = task.get('task_name')
            start_date = task['start_date'].strftime('%Y-%m-%d')
            end_date = task['end_date'].strftime('%Y-%m-%d')
            duration = task['duration']
        
            report += f"任务名称: {task_name}\n"
            report += f"开始日期: {start_date}\n"
            report += f"结束日期: {end_date}\n"
            report += f"持续时间: {duration} 天\n"
            report += "-" * 20 + "\n"

        with open(report_path, 'w', encoding='utf-8') as file:g
            file.write(report)
            
        return report
        


if __name__ == '__main__':
    file_path = r'D:\我的坚果云\work\Sincetimes_DOW\2023-09-21 XCY DYL DOW damo 活动历史对比\damo活动\【日本】霸王天下-活动表2020年.xlsx'
    test = GanttData(file_path, '活动排期', 4)
    gantt_data = test.date_to_task('2020-07-01', '2020-07-07')

    gantt_test = GanttPlot(gantt_data)
    # print(gantt_data)
    # print(gantt_test.gantt_data_process())
    # gantt_test.gantt_plot()
    report_test = GanttReport(gantt_data)
    print(report_test.gantt_data_analysis())
    # print(gantt_data)
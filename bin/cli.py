# Created by Michael Scales
# Version 1.0
# This program will take user modified excel spreadsheets 
# parses the data, collates the information and generates a mostly completed Agent Report
# Mar 03 2016

# -*- coding: utf-8 -*-
import traceback
from pyexcel import Sheet, get_book, get_sheet
from os.path import dirname, abspath, join
from types import MethodType
from utilities.utilities import safe_div, create_schedule, read_time_card, open_wb, read_feature_card


VALID_DAYS_IN_MONTH = 21


def in_schedule(schedule):
    for k, v in schedule.items():
        yield(
            (
                r'{0} {1}({2})'.format(v.f_name, v.l_name, v.ext), v
            )
        )


def main():
    directory = join(dirname(dirname(abspath(__file__))), 'raw_files')

    reports = {
        'agent_time_card': open_wb(get_book(file_name=join(directory, 'Agent_Time_Card.xlsx'))),
        'feature_trace': open_wb(get_book(file_name=join(directory, 'Agent_Realtime_Feature_Trace.xlsx')))
    }

    schedule = create_schedule(
        get_sheet(file_name=join(r'M:/Help Desk/', 'Schedules for OPS.xlsx'),
                  name_columns_by_row=0,
                  name_rows_by_column=0)
    )

    # Bind read function to respective reports
    check_time_card = MethodType(read_time_card, reports['agent_time_card'])
    check_feature_card = MethodType(read_feature_card, reports['feature_trace'])

    output = Sheet(colnames=['', '% Avail', 'Absences', 'Late'])
    try:
        for sheet_name, emp_data in in_schedule(schedule):
            tc_data = check_time_card(sheet_name, emp_data)
            dnd_time = check_feature_card(sheet_name)
            output.row += [
                '{row_name}'.format(
                    row_name=emp_data.ext
                ),
                '{val:.1%}'.format(
                    val=safe_div(dnd_time, tc_data['Duration'])
                ),
                VALID_DAYS_IN_MONTH - tc_data['Log Events'],
                tc_data['Late']
            ]
        # for rpt_name, rpt in reports.items():
        #     if isinstance(rpt, Book):
        #         r_filter = (
        #             {
        #                 'filter_col': reports[rpt][0],
        #                 'filter_val': reports[rpt][1]
        #             } if isinstance(reports.get(rpt, None), tuple)
        #             else None
        #         )
        #         for sheet in rpt:
        #             for row_name in sheet.rownames:
        #                 try:
        #                     if r_filter:
        #                         if sheet[row_name, r_filter['filter_col']] == r_filter['filter_val']:
        #                             cell_value = sheet[row_name, 'Duration']
        #                         else:
        #                             raise ValueError()
        #                     else:
        #                         cell_value = sheet[row_name, 'Duration']
        #                 except ValueError:
        #                     pass
        #                 else:
        #                     try:
        #                         summary[sheet.name][rpt_name] += get_sec(cell_value)
        #                     except KeyError:
        #                         summary[sheet.name] = {
        #                             **summary.get(sheet.name, {}),
        #                             rpt_name: get_sec(cell_value)
        #                         }

        # for agent, col_data in summary.items():
        #     if 'feature_trace' in col_data.keys():
        #         output.row += [
        #             '{row_name}'.format(
        #                 row_name=int(search(r"([0-9]+)", agent).group(0))
        #             ),
        #             '{val:.1%}'.format(
        #                 val=safe_div(col_data['feature_trace'], col_data['agent_time_card'])
        #             )
        #         ]
        else:
            output.name_rows_by_column(0)
            print(output)
            output.save_as(filename=join(dirname(dirname(abspath(__file__))), 'output', 'outfile.xlsx'))
    except Exception:
        print(traceback.format_exc())
    print('completed life cycle')


if __name__ == "__main__":
    main()

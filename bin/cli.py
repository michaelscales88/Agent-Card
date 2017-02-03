# Created by Michael Scales
# Version 1.0
# This program will take user modified excel spreadsheets 
# parses the data, collates the information and generates a mostly completed Agent Report
# Mar 03 2016

# -*- coding: utf-8 -*-
import traceback
from re import search
import pyexcel as pe
from pyexcel import Sheet, get_book
from pyexcel.book import Book
from os.path import dirname, abspath, join
from collections import OrderedDict
from utilities.utilities import get_sec, header_filter, safe_div


# def get_my_book(**keywords):
#     book_stream = _get_book(**keywords)
#     my_book = CustomBook(book_stream.to_dict(),
#                          filename=book_stream.filename,
#                          path=book_stream.path)
#     return my_book
#
#
# class CustomBook(Book):
#     def __hash__(self):
#         return hash((self.name, self.location))
#
#     def __eq__(self, other):
#         return (self.name, self.location) == (other.name, other.location)
#
#     def __ne__(self, other):
#         return not (self == other)


def main():
    directory = join(dirname(dirname(abspath(__file__))), 'raw_files')
    reports = {
        'agent_time_card': get_book(file_name=join(directory, 'Agent_Time_Card.xlsx')),
        'feature_trace': get_book(file_name=join(directory, 'Agent_Realtime_Feature_Trace.xlsx'))
    }
    summary = OrderedDict()
    reports[reports['feature_trace']] = ('Feature Type', 'Do Not Disturb')
    try:
        wb_filter = pe.filters.RowFilter(header_filter)
        for rpt_name, rpt in reports.items():
            if isinstance(rpt, Book):
                rpt.remove_sheet('Summary')
                r_filter = (
                    {
                        'filter_col': reports[rpt][0],
                        'filter_val': reports[rpt][1]
                    } if isinstance(reports.get(rpt, None), tuple)
                    else None
                )
                for sheet in rpt:
                    sheet.filter(wb_filter)
                    sheet.name_columns_by_row(0)
                    sheet.name_rows_by_column(0)
                    for row_name in sheet.rownames:
                        try:
                            if r_filter:
                                if sheet[row_name, r_filter['filter_col']] == r_filter['filter_val']:
                                    cell_value = sheet[row_name, 'Duration']
                                else:
                                    raise ValueError()
                            else:
                                cell_value = sheet[row_name, 'Duration']
                        except ValueError:
                            pass
                        else:
                            try:
                                summary[sheet.name][rpt_name] += get_sec(cell_value)
                            except KeyError:
                                summary[sheet.name] = {
                                    **summary.get(sheet.name, {}),
                                    rpt_name: get_sec(cell_value)
                                }
        # print(
        #     '\n'.join(['{agent}\n{values}'.format(agent=k, values=v) for k, v in summary.items()])
        # )
        output = Sheet(colnames=['', '% Avail'])
        for agent, col_data in summary.items():
            if 'feature_trace' in col_data.keys():
                output.row += [
                    '{row_name}'.format(
                        row_name=int(search(r"([0-9]+)", agent).group(0))
                    ),
                    '{val:.1%}'.format(
                        val=safe_div(col_data['feature_trace'], col_data['agent_time_card'])
                    )
                ]
        else:
            output.name_rows_by_column(0)
            print(output)
            output.save_as(filename=join(dirname(dirname(abspath(__file__))), 'output', 'outfile.xlsx'))
    except Exception:
        print(traceback.format_exc())
    print('completed life cycle')


if __name__ == "__main__":
    main()

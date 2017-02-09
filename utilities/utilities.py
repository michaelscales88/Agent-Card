import re
from datetime import datetime, timedelta
from collections import namedtuple
from dateutil.parser import parse


def get_sec(time_string):  # returns time provided as as string in hours, minutes, sec
    try:
        h, m, s = [int(float(i)) for i in time_string.split(':')]
    except TypeError:
        return 0
    return convert_sec(h, m, s)  # Converts hours, minutes, sec to seconds


def convert_sec(h, m, s):  # Converts hours, minutes, sec to seconds
    return (3600 * int(h)) + (60 * int(m)) + int(s)


def header_filter(row_index, row):
    corner_case = re.split('\(| - ', row[0])
    bad_word = corner_case[0].split(' ')[0] not in ('Feature', 'Call', 'Event')
    return True if len(corner_case) > 1 else bad_word


def safe_div(n, d):
    rtn = 0
    try:
        rtn = n / d
    except ZeroDivisionError:
        pass
    return 1 if rtn > 1 else rtn


def date_tuple_factory(date_string):
        dt = namedtuple('date', 'start end')
        try:
            (raw_start, raw_end) = split_str(date_string)
        except ValueError:
            return_dt = None
        else:
            dt_start = datetime.strptime(raw_start, '%H:%M').time()
            dt_end = datetime.strptime(raw_end, '%H:%M').time()
            return_dt = dt(start=dt_start, end=dt_end)
        return return_dt


def split_str(t_string):
    return t_string.split('-')


def safe_parse(dt=None):
    try:
        return parse(dt)
    except ValueError:
        # print('Could not parse date_time: {dt}'.format(dt=dt))
        return None


def create_schedule(data):
    return_dict = {}
    new_schedule = namedtuple('this_emp', 'Monday Tuesday Wednesday Thursday Friday '
                                          'Saturday Sunday f_name l_name data ext')
    new_schedule.__new__.__defaults__ = (None,) * len(new_schedule._fields)
    for emp in data.rownames:
        emp_schedule = new_schedule(f_name=data[emp, 'First'],
                                    l_name=data[emp, 'Last'],
                                    ext=int(emp),
                                    Monday=date_tuple_factory(data[emp, 'Monday']),
                                    Tuesday=date_tuple_factory(data[emp, 'Tuesday']),
                                    Wednesday=date_tuple_factory(data[emp, 'Wednesday']),
                                    Thursday=date_tuple_factory(data[emp, 'Thursday']),
                                    Friday=date_tuple_factory(data[emp, 'Friday']),
                                    Saturday=date_tuple_factory(data[emp, 'Saturday']),
                                    Sunday=date_tuple_factory(data[emp, 'Sunday']))
        return_dict[emp] = emp_schedule
    return return_dict


def read_time_card(time_card, sheet_name, emp_data):
    rtn_data = {
        'Late': 0,
        'Duration': 0,
        'Log Events': []
    }
    sheet = time_card[sheet_name]
    for row_name in sheet.rownames:
        # print(row_name)
        start_time = safe_parse(dt=sheet[row_name, "Logged In"])
        end_time = safe_parse(dt=sheet[row_name, "Logged Out"])
        if start_time.date() == end_time.date() if start_time and end_time else start_time:  # Excludes roll over days
            try:
                shift_start = emp_data[start_time.weekday()].start  # checks if the agent is scheduled this date
            except AttributeError:
                pass
                # print(start_time)
                # print('**Extra Day**')
            else:
                rtn_data['Duration'] += get_sec(sheet[row_name, "Duration"])  # Any duration on a valid day counts.
                if start_time.date() not in rtn_data['Log Events']:  # "hashes" date - counts ea. login one time
                    rtn_data['Log Events'].append(start_time.date())
                    grace_time = datetime.combine(start_time.date(), shift_start) + timedelta(minutes=5)
                    if start_time > grace_time:  # Can only be late once per non-absent day
                        rtn_data['Late'] += 1
    rtn_data['Log Events'] = len(rtn_data['Log Events'])
    return rtn_data


def read_feature_card(feature_card, sheet_name):
    rtn_data = 0
    sheet = feature_card[sheet_name]
    for row_name in sheet.rownames:
        if sheet[row_name, 'Feature Type'] == 'Do Not Disturb':
            rtn_data += get_sec(sheet[row_name, 'Duration'])
    return rtn_data


def open_wb(wb):
    wb.remove_sheet('Summary')
    for sheet in wb:
        del sheet.row[header_filter]
        sheet.name_columns_by_row(0)
        sheet.name_rows_by_column(0)
    return wb

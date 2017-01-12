import re


def get_sec(time_string):  # returns time provided as as string in hours, minutes, sec
    try:
        h, m, s = [int(float(i)) for i in time_string.split(':')]
    except TypeError:
        return 0
    return convert_sec(h, m, s)  # Converts hours, minutes, sec to seconds


def convert_sec(h, m, s):  # Converts hours, minutes, sec to seconds
    return (3600 * int(h)) + (60 * int(m)) + int(s)


def header_filter(row):
    corner_case = re.split('\(| - ', row[0])
    bad_word = corner_case[0].split(' ')[0] not in ('Feature', 'Call', 'Event')
    return True if len(corner_case) > 1 else bad_word


def safe_div(n, d):
    try:
        return n / d
    except ZeroDivisionError:
        return 0

# Created by Michael Scales
# Version 1.0
# This program will take user modified excel spreadsheets 
# parses the data, collates the information and generates a mostly completed Agent Report
# Mar 03 2016

# -*- coding: utf-8 -*-
import traceback
from utilities.utilities import (open_reports, read_reports, writeTimeCard)


def main():
    try:
        agent_report, dnd_report = open_reports()  # opens raw files
        time_card_accumulator = []  # container for all agents
        read_reports(agent_report, dnd_report, time_card_accumulator)  # takes data from raw files and adds to Agent container
        writeTimeCard(time_card_accumulator)  # writes output to outfile
    except (SystemExit, KeyboardInterrupt, UserWarning):  # Passes any valid system exit command
        pass
    except Exception as e: # T/S
        exc_type, exc_value, exc_traceback = sys.exc_info()
        traceback.print_exception(exc_type, exc_value, exc_traceback, limit=2, file=sys.stdout)
        print(e)
    finally:
        print("Packing up the tools to quit.")


if __name__ == "__main__":
    from os import path, sys
    sys.path.append(path.dirname(path.dirname(path.abspath(__file__))))
    main()


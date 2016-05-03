# Created by Michael Scales
# Version 1.0
# This program will take user modified excel spreadsheets 
# parses the data, collates the information and generates a mostly completed Agent Report
# Mar 03 2016

# -*- coding: utf-8 -*-
import openpyxl
import traceback
import string

import sys
import os
from os import path

sys.path.append(path.dirname(path.dirname(path.abspath(__file__))))

from utilities import Agent
from utilities.utilities import ( openReports, readReports, writeTimeCard )


def main():
    try:
        agentReport, dndReport = openReports() # opens raw files
        timeCard = [] # container for all agents
        readReports(agentReport, dndReport, timeCard) # takes data from raw files and adds to Agent container
        writeTimeCard(timeCard) # writes output to outfile
    except (SystemExit, KeyboardInterrupt, UserWarning): # Passes any valid system exit command
        pass
    except Exception as e: # T/S
        exc_type, exc_value, exc_traceback = sys.exc_info()
        traceback.print_exception(exc_type, exc_value, exc_traceback, limit=2, file=sys.stdout)
        print(e)
    finally:
        print("Packing up the tools to quit.") 
        
if (__name__ == "__main__"):
    main()
    

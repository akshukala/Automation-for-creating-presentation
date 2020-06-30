# -*- coding: utf-8 -*-
"""
Created on Wed Nov 20 10:45:03 2019

@author: akshay.kale
"""
import tkinter as tk
from tkcalendar import DateEntry
from datetime import date

'''Defining all the constant variable '''
class Variables:
    def __init__(self):
        # QRC Data file path required in QRC_report.py and QRC_circlewise_report.py
        self.QRC_PATH = "C:\\Users\\akshay.kale\\Documents\\Scripts\\AT_Increase_Automation\\Data_files\\QRC_05_08Dec.xlsx"
        
        self.PPT_PATH = "C:\\Users\\akshay.kale\\Desktop\\Analysis.pptx" # PPT store path
        
        self.IMG_PATH = "C:\\Users\\akshay.kale\\Desktop\\" # Image store path
        
        self.DATE_1 = '08-09 Dec' # Date required in QRC_report.py and QRC_circlewise_report.py
        
        self.DATE_2 = '10-11 Dec' # Date required in QRC_report.py and QRC_circlewise_report.py
        
        self.MONTH_1 = 'Sep'
        
        self.MONTH_2 = 'Oct'
        
        self.MONTH_3 = 'Nov'
        
        #Service Trend file path required in ServiceTrend.py file
        self.ST_PATH = "C:\\Users\\akshay.kale\\Documents\\Scripts\\AT_Increase_Automation\\Data_files\\ServiceTrendReport.xlsx" 
        
        self.ST_DATE_1 = '2020-01-06'  # Date required in ServiceTrend.py file
        
        self.ST_DATE_2 = '2020-01-07' # Date required in ServiceTrend.py file
        
        # IVR Data file path required in function_report_automation.py
        self.IVR_PATH = "C:\\Users\\akshay.kale\\Documents\\Scripts\\AT_Increase_Automation\\Data_files\\"
        
        self.NETWORK_PARAMETER_PREPAID = {'Q--Coverage Availability', 'C-Network Down', 'C-Speed Related', 'C-Accessibility',
                                      'C -Coverage'}
        
        self.NETWORK_PARAMETER_POSTPAID = {'Q-Coverage Availability', 'C-Network Down', 'C-Speed Related', 'C-Accessibility',
                                       'C- Coverage'}

variable = Variables()

'''Creating a datepicker'''
class MyDateEntry(DateEntry):
    def __init__(self, master=None, **kw):
        DateEntry.__init__(self, master=None, date_pattern='y-mm-dd',  **kw)
        # add black border around drop-down calendar
        self._top_cal.configure(bg='black', bd=4)
        # add label displaying today's date below
        tk.Label(self._top_cal, bg='gray90', anchor='w',
                 text='Today: %s' % date.today().strftime('%Y-%m-%d')).pack(fill='x')




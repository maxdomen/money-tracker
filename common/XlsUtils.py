from datetime import datetime

__author__ = 'Max'
import xlrd
def date_from_xls(xlsdate,book2_datemode):
     tdate=xlrd.xldate_as_tuple(xlsdate,book2_datemode)
     date=datetime(tdate[0],tdate[1],tdate[2],tdate[3],tdate[4],tdate[5])
     return date


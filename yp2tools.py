import sys
from openpyxl import load_workbook
from mmap import mmap,ACCESS_READ
from xlrd import open_workbook

basefile = (sys.argv[0])

#print(open_workbook(basefile))

wb = load_workbook(filename = basefile)
#!/usr/bin/env python


import os
from openpyxl import load_workbook
filename = 'C:\\Macomb ROW\\Returned Docs Delivery Script\\test\\Armada_Not_Drawn.xlsx'
lpdir = 'C:/Macomb ROW/Returned Docs Delivery Script/test/Armada Liber Page Docs' 
wb = load_workbook(filename, use_iterators=True)
ws = wb.get_active_sheet()

#Define function to create a list of all file names (without extension) in a directory (recursive).
#The list output should contain all liber pages in the directory.
#TODO: Would generator be more efficient?
def get_lps(lpdir):
    fnlist = []
    for path,dirs,files in os.walk(lpdir):
        for fn in files:
            fnlist.append(fn[:fn.rfind('.')])
    return fnlist    
           
#Define function that reads liber page values from a spreadsheet into a list (liber pages must be in column 'B'
# and assumes row 1 is a header row and all liber pages are "Not Drawn")
# TODO: Modify code to determine if liber page is 'Drawn' or 'Not Drawn', and then
# pull liber page values for only 'Not Drawn' into list
def get_not_drawn(filename):
    ndlist = []
    hr = ws.get_highest_row()
    for i in range(2,hr):
        lp = ws.cell(row = i, column = 2).value
        if lp != None:
            ndlist.append(str(lp))
        else:
            return ndlist
    return ndlist
        
 # Define function that compares the two lists
 # TODO: Modify to actually delete the 'drawn' liber pages from the directory.
 # Currently, it prints which liber pages should be deleted.
def comp_lists(lpdir,filename):
    for i in get_lps(lpdir):
        if i not in get_not_drawn(filename):
            print i, "should be deleted."


# Define generator object to recursively yield empty directory paths (deleting the "Drawn" liber pages will leave empty directories
# that should be cleaned up)
def get_empty_dir(lpdir):
    for path,dirs,files in os.walk(lpdir):
        if not dirs and not files:
            yield path 

# Define function to delete empty directories using generated list from GetEmptyDir
#TODO: This function can be combined with generator object above. Not necessary to have seperate functions. Is generator neccessary??
def del_empty_dir(lpdir):
    empty_dir = list(get_empty_dir(lpdir))
    for i in empty_dir:
        os.rmdir(i)
        print i, 'is empty and has been deleted.'
    print 'All empty directories have been deleted.'

#!/usr/bin/env python


import os
from openpyxl import load_workbook
filename = raw_input("Enter full path to spreadsheet file (include file extension): ")
#'C:\\Macomb ROW\\Returned Docs Delivery Script\\test\\Armada_Not_Drawn.xlsx'
lpdir = raw_input("Enter full path of directory containing the documents: ")
#'C:/Macomb ROW/Returned Docs Delivery Script/test/Armada Liber Page Docs' 
wb = load_workbook(filename, use_iterators=True)
ws = wb.get_active_sheet()

#Define generator to get all files (full paths) in a directory (recursive).
def get_docs(lpdir):
    for path,dirs,files in os.walk(lpdir):
        for fn in files:
            yield os.path.join(path,fn)

#Define function that reads liber page values from a spreadsheet into a list (liber pages must be in column 'B'
# and assumes row 1 is a header row and all liber pages are "Not Drawn")
# TODO: Modify code to determine if liber page is 'Drawn' or 'Not Drawn', and then
# pull liber page values for only 'Not Drawn' into list
def get_not_drawn(filename):
    ndlist = []
    hr = ws.get_highest_row()
    for i in range(2,hr):
        lp = ws.cell(row = i, column = 2).value
        if lp:
            ndlist.append(str(lp))
        else:
            return ndlist
    return ndlist
        
 # Define function that compares the two lists
 # TODO: Modify to actually delete the 'drawn' liber pages from the directory.
 # Currently, it prints which liber pages should be deleted.
def comp_lists(lpdir,filename):
    for i in list(get_docs(lpdir)):
        if os.path.splitext(os.path.split(i)[1])[0] not in get_not_drawn(filename):
            os.remove(i)
            print i, "has been deleted."


# Define generator object to recursively yield empty directory paths (deleting the "Drawn" liber pages will leave empty directories
# that should be cleaned up)
def get_empty_dir(lpdir):
    for path,dirs,files in os.walk(lpdir):
        if not dirs and not files:
            yield path 

# Define function to delete empty directories using generated list from GetEmptyDir
#TODO: Currently deletes empty directories at time of generation, but does not delete the empty directories
# created after deleting the first "round" of empty directories.
def del_empty_dir(lpdir):
    empty_dir = list(get_empty_dir(lpdir))
    for i in empty_dir:
        os.rmdir(i)
        print i, 'is empty and has been deleted.'
    print 'All empty directories have been deleted.'


#print comp_lists(lpdir,filename)
#print del_empty_dir(lpdir)

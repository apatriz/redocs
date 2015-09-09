#!/usr/bin/env python
#NAME: redocs.py
##TODO: transform user input for column choice ("A","B",etc.) into corresponding integer val
## --current functionality requires all parameter inputs are strings (this can be modified if not using raw_input()) to obtain parameter values

""" This python script reads a spreadsheet column containing file names into a list,
    and compares it to a list of all files in a specific directory. The purpose is to identify files
    that should be preserved, delete the rest, execute a cleanup of the directory structure to eliminate
    empty directories and produce a deliverable consisting of only the needed files.
    It also creates a new workbook.xlsx , containing only the entries for the preserved files."""


import os
from openpyxl import load_workbook
from openpyxl import Workbook

##raw_input("Enter full path to spreadsheet file (include file extension): ")
filename = 'C:\\Macomb ROW\\script_testing\\redocs\\test\\Armada_Not_Drawn.xlsx'

##raw_input("Enter full path of directory containing the documents: ")
lpdir = 'C:\\Macomb ROW\\script_testing\\redocs\\test\\Armada Liber Page Docs'

##raw_input("Enter the spreadsheet column which lists the file names to be preserved (i.e. Enter '0' for column 'A', '1' for column 'B',etc.): ")
fn_column = '1'

##raw_input("Enter the spreadsheet column which lists the values for identifying which files should be preserved: ")
ident_column = '4'

##raw_input("Enter the value that identifies which files should be preserved: ")
keep_val = 'n'

##raw_input("Enter the spreadsheet column which lists the file locations: ")
fold_col = '0'

##raw_input("Enter the spreadsheet column which lists the comments: ")
comment_col = '2'

##raw_input("Enter the column labels for the output spreadsheet (i.e. Folder,Liber Page,Comments for column 1, 2 and 3 labels respectively): ")
col_header = 'Folder','Liber Page','Comments'

wb = load_workbook(filename, use_iterators= True)
ws = wb.get_active_sheet()


#define generator to get all files (full paths) in a directory (recursive).
def get_docs(lpdir):
    for path,dirs,files in os.walk(lpdir):
        for fn in files:
            yield os.path.join(path,fn)

#define function that reads values from a spreadsheet into a list, based on user input

def tokeep(filename):
    keeplist = []
    for row in ws.iter_rows():
        if row[int(ident_column)].value == keep_val:
            keeplist.append(row[int(fn_column)].value)
    return keeplist


# define function that compares the two lists and deletes files that are not in 'tokeep' list
def del_unwanted(lpdir,filename):
    for i in get_docs(lpdir):
        if os.path.splitext(os.path.split(i)[1])[0] not in tokeep(filename):
            # print i, "will be deleted"
            os.remove(i)
            print i, "has been deleted."
    print "All unwanted files deleted."


# define generator object to recursively yield empty directory paths (deleting files may leave empty directories
# that should be cleaned up)
def get_empty_dir(lpdir):
    for path,dirs,files in os.walk(lpdir):
        if not dirs and not files:
            yield path

# define function to delete empty directories using generated list from GetEmptyDir.
# will run until no empty directories are found.
def cleanup(lpdir):
    while True:
        empty_dir = list(get_empty_dir(lpdir))
        for i in empty_dir:
            os.rmdir(i)
            print i, 'is empty and has been deleted.'
        if empty_dir == []:
            print 'All empty directories deleted.'
            break

# define function that returns a list of data entries for the preserved files. Each entry in the list is a tuple containing the needed data for each file (i.e.(folder, filename, comments)).
# the first entry will always be the column header labels
def create_entry_list(filename):
    entrylist = [(col_header)]
    prevfol = ''
    for row in ws.iter_rows():
        folder = row[int(fold_col)].value
        fn = row[int(fn_column)].value
        comment = row[int(comment_col)].value
        if not folder:
            folder = prevfol
        else:
            prevfol = folder
        if row[int(ident_column)].value == keep_val:
            entry = folder,fn,comment
            entrylist.append(entry)
    return entrylist

# define function that creates new workbook.xlsx named "Returned" inside the directory containing the preserved files,
# which contains the data entries for the preserved files (from create_entry_list)
def create_new_xl(filename):
    nb = Workbook(write_only=True)
    ns = nb.create_sheet()
    entrylist = create_entry_list(filename)
    for entry in entrylist:
        ns.append([value for value in entry])
    nb.save(lpdir + '/Returned.xlsx')
    print "Saved new spreadsheet containing preserved file entries to {0}/Returned.xlsx".format(lpdir)


# define the main function
def main():
    del_unwanted(lpdir,filename)
    cleanup(lpdir)
    create_new_xl(filename)

# call the main function
if __name__ == "__main__":
    main()
    

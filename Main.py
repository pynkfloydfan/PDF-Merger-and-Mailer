import sys
import os
import PySimpleGUI as sg
import win32com.client as w32
import win32api
from collections import namedtuple
from itertools import groupby
from PyPDF2 import PdfFileMerger, PdfFileReader, PdfFileWriter
import reportlab
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from operator import attrgetter
import pandas as pd



"""
    A PySimpleGUI based program that will display the Outlook folder heirarchy, allow you to choose a folder and then
    will extract all PDF files (according to any required filters), combine then into a single file and email them
"""

# Base64 versions of images of a folder and a file. PNG files (may not work with PySimpleGUI27, swap with GIFs)

folder_icon = b'iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAACXBIWXMAAAsSAAALEgHS3X78AAABnUlEQVQ4y8WSv2rUQRSFv7vZgJFFsQg2EkWb4AvEJ8hqKVilSmFn3iNvIAp21oIW9haihBRKiqwElMVsIJjNrprsOr/5dyzml3UhEQIWHhjmcpn7zblw4B9lJ8Xag9mlmQb3AJzX3tOX8Tngzg349q7t5xcfzpKGhOFHnjx+9qLTzW8wsmFTL2Gzk7Y2O/k9kCbtwUZbV+Zvo8Md3PALrjoiqsKSR9ljpAJpwOsNtlfXfRvoNU8Arr/NsVo0ry5z4dZN5hoGqEzYDChBOoKwS/vSq0XW3y5NAI/uN1cvLqzQur4MCpBGEEd1PQDfQ74HYR+LfeQOAOYAmgAmbly+dgfid5CHPIKqC74L8RDyGPIYy7+QQjFWa7ICsQ8SpB/IfcJSDVMAJUwJkYDMNOEPIBxA/gnuMyYPijXAI3lMse7FGnIKsIuqrxgRSeXOoYZUCI8pIKW/OHA7kD2YYcpAKgM5ABXk4qSsdJaDOMCsgTIYAlL5TQFTyUIZDmev0N/bnwqnylEBQS45UKnHx/lUlFvA3fo+jwR8ALb47/oNma38cuqiJ9AAAAAASUVORK5CYII='
file_icon = b'iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAACXBIWXMAAAsSAAALEgHS3X78AAABU0lEQVQ4y52TzStEURiHn/ecc6XG54JSdlMkNhYWsiILS0lsJaUsLW2Mv8CfIDtr2VtbY4GUEvmIZnKbZsY977Uwt2HcyW1+dTZvt6fn9557BGB+aaNQKBR2ifkbgWR+cX13ubO1svz++niVTA1ArDHDg91UahHFsMxbKWycYsjze4muTsP64vT43v7hSf/A0FgdjQPQWAmco68nB+T+SFSqNUQgcIbN1bn8Z3RwvL22MAvcu8TACFgrpMVZ4aUYcn77BMDkxGgemAGOHIBXxRjBWZMKoCPA2h6qEUSRR2MF6GxUUMUaIUgBCNTnAcm3H2G5YQfgvccYIXAtDH7FoKq/AaqKlbrBj2trFVXfBPAea4SOIIsBeN9kkCwxsNkAqRWy7+B7Z00G3xVc2wZeMSI4S7sVYkSk5Z/4PyBWROqvox3A28PN2cjUwinQC9QyckKALxj4kv2auK0xAAAAAElFTkSuQmCC'

# ------ Create a COM link to Outlook and get the default inbox and the mailbox it comes from ------
outlook = w32.Dispatch('Outlook.Application')
mapi = outlook.GetNamespace('MAPI')
inbox = mapi.GetDefaultFolder(6)
mailbox = inbox.Parent
pdflist =list()

# ------ Some constants ------ #
savefolder = os.path.join(os.path.expandvars("%userprofile%"), 'Documents')
firstname = win32api.GetUserNameEx(3)
firstname = firstname.split()[1]
firstname = re.findall('^[A-Za-z]+', firstname)[0]

sg.theme('LightGreen')

# ------ Set up the Tree with the root node as the default mailbox of Outlook ------ #
treedata = sg.TreeData()
treedata.Insert(parent='', key=mailbox.Name, text=mailbox.Name, 
                values=[mailbox.StoreID, mailbox.EntryID], icon=folder_icon)

# ------ Custom Class to handle the format of the pdf filenames so the right ones can be grouped and merged in the right order ------ #
class MyFileName:
    def __init__(self, filename):
        self.filename = filename
        temp = filename.split()
        # is file a Section or an Appendix?
        self.section = 'Sec' in temp
        self.appendix = 'Appendix' in temp

        if self.section:
            self.freq = filename.split('Sec ')[0].strip()
            self.part = filename.split('Sec ')[1].split()[0].strip()
            self.region = filename.split('Sec ' = self.part + ' ')[1].strip()
        elif self.appendix:
            self.freq = filename.split('Appendix ')[0].strip()
            self.part = filename.split('Appendix ')[1].split()[0].strip()
            self.region = filename.split('Appendix ' = self.part + ' ')[1].strip()

    def __repr__(self):
        return self.filename

# ------ Define and run function to fill the TreeData with all Outlook folders ------ #
def add_folders(parent, tree=treedata):
    """
    Iterates through the parent mailbox adding all the folders and subfolders
    to the tree

    Params:
    parent: a reference to an Outlook Folder item
    tree: a PySimpleGUI treedata object

    Returns:
    Nothing
    """
    for f in parent.Folders:
        tree.Insert(parent=parent.Name, key=f.Name, text=f.Name, values=[f.StoreID, f.EntryID], icon=folder_icon)
        # recursive so all folders and subfolders are added
        add_folders(f)

# Add all outlook folders to the tree
add_folders(mailbox)


# ------ Define the layout of the window with the Tree ------ #
layout = [[sg.Text('Outlook folder browser\nSelect folder with pdfs to merge',
                   size=(50,2)),
           sg.Text('Select pdfs to merge',
                   size=(40,2))],
          [sg.Tree(data=treedata,
                   headings=['StoreID', 'EntryID'],
                   auto_size_columns=True,
                   visible_column_map={False, False},
                   num_rows=20,
                   col0_width=30,
                   key='-TREE-',
                   show_expanded=False,
                   enable_events=True),
           sg.Listbox(values='',
                      enable_events=True,
                      size=(50,24),
                      key='-FILES-',
                      select_mode=sg.LISTBOX_SELECT_MODE_EXTENDED)],
          [sg.Button('Get pdfs'), sg.Button('Cancel'), sg.Button('Merge pdfs', disabled=True, key='-MERGE-')]]

# ------ Display the Tree GUI ------
window = sg.Window('Tree Element Test', layout,
                   grab_anywhere=False,
                   keep_on_top=False,
                   resizable=True)

# ------ Create a reference to the treedata dictionary ------ #
TreeDict = treedata.tree_dict


# ------ Get a list of all pdf files in the folder ------ #
def get_pdfs(folder):
    """
    Creates a list of Named Tuples of (filename, entryID, index) 
    for all pdf files in the folder.
    filename is a MyFileName object

    Params:
    folder: a reference to an Outlook Folder object

    Return:
    A list of named tuples
    """
    messages = folder.Items

    # ------ what about using namedtuple instead to hold filename & entryID? ------ #
    pdf = namedtuple('pdf', 'filename entryID index')

    pdflist = [[pdf(filename=MyFileName(att.Filename), entryID=m.EntryID, index=att.Index) for att in m.Attachments
                if att.Filename.endswith('pdf')]
               for m in messages if m.Attachments.Count > 0]

    # ------ flatten the list of lists and sort by filename ------ #
    pdflist = list(itertools.chain.from_iterable(pdflist))
    pdflist = sorted(pdflist, key=attrgetter('filename.filename'))

    return pdflist

# To use GetItemFromID you need the entryID of the Item (e.g. MailItem)
# The StoreID of the folder it came from is optional


# ------ Event loop to catch selected Outlook folder and update the listbox ------ #
while True:
    event, values = window.read()
    if event in (None, 'Cancel'):
        window.close()
        sys.exit()
    if event is 'Get pdfs':
        storeID, entryID = TreeDict[values['-TREE-'][0]].values
        pdf_list = get_pdfs(outlook.GetFolderFromID(entryID, storeID))
        window['-FILES-'].update(pdf_list)
        window['-MERGE-'].update(disabled=False)
    if event is '-MERGE-':
        print(event, values)
        break


def merge_pdf(pdf_files):
    """
    Iterates through a list of namedtuples of type=pdf(filename, entryID) and merges them into a single file.
    All the files are attachments to Outlook messages stored in a single Outlook folder
    :param pdf_files: a list of namedtuples of type=pdf
    :return: a single pdf file saved at the chosen location
    """
    save_folder = sg.PopupGetFolder('Choose merged pdf file save location')
    merger = PdfFileMerger()
    for f in pdf_files:
        # ------ first create a reference to the message the pdf file is stored in ------ #
        msg = outlook.GetItemFromID(f.entryID) # storeID is optional
        # ------ then get the pdf file ------ #
        pdf_file = msg.Attachments.Item(f.index)
        # ------ save the pdf file and then add it to the merger job ------ #
        pdf_file.SaveAsFile(os.path.join(save_folder, f.filename))
        merger.append(fileobj=os.path.join(save_folder, f.filename))
    # ------ save the merged file ------ #
    merger.write(os.path.join(save_folder, 'MergedPdf.pdf'))
    merger.close()
    sg.Popup('MergedPdf saved at {}'.format(save_folder))
    # ------ ask whether to keep the source pdf files or delete them ------ #
    del_choice = sg.PopupYesNo('Delete original pdfs from saved location?')
    if del_choice == 'Yes':
        for f in pdf_files:
            os.remove(os.path.join(save_folder, f.filename))
        sg.Popup('Source pdfs have been deleted')



merge_pdf(values['-FILES-'])


# print(pdf_dict.values())
print(pdf_list)

work_folder = outlook.GetFolderFromID(entryID, storeID)
print('There are {} messages in this folder'.format(work_folder.Items.Count))

work_folder_Restricted = work_folder.Items.Restrict("[ReceivedTime] >= '23/12/2019'") #December 23, 2019
print('There are {} messages in this folder since 23rd Dec 2019'.format(work_folder_Restricted.Count))
for msg in work_folder_Restricted:
    print(msg.Subject)

window.close()
sys.exit()


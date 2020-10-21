import sys
import os
import PySimpleGUI as sg
import win32com.client as w32
import win32api
from collections import namedtuple
from itertools import groupby, chain
from PyPDF2 import PdfFileMerger, PdfFileReader, PdfFileWriter
import reportlab
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from operator import attrgetter
import pandas as pd
import re


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
pdflist = list()

# ------ Some constants ------ #
save_folder = os.path.join(os.path.expandvars("%userprofile%"), 'Documents')
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
    """
    The names of pdf files are in the format:
    {freq} + {Sec | Appendix} + {part} + {region}
    e.g. "Monthly Sec 1 EMEA"

    This class takes in a filename and determines those named parts of it
    """

    def __init__(self, filename):
        self.filename = filename
        temp = filename.split()
        # is file a Section or an Appendix?
        self.section = 'Sec' in temp
        self.appendix = 'Appendix' in temp

        if self.section:
            self.freq = filename.split('Sec ')[0].strip()
            self.part = filename.split('Sec ')[1].split()[0].strip()
            self.region = filename.split('Sec ' + self.part + ' ')[1].strip()
        elif self.appendix:
            self.freq = filename.split('Appendix ')[0].strip()
            self.part = filename.split('Appendix ')[1].split()[0].strip()
            self.region = filename.split(
                'Appendix ' + self.part + ' ')[1].strip()

    def __repr__(self):
        return self.filename

# ------ Define and run function to fill the TreeData with all Outlook folders ------ #
def add_folders(parent, tree=treedata):
    """
    Iterates through the parent mailbox adding all the folders and subfolders
    to the tree

    Parameters:
    -----------
    parent: a reference to an Outlook Folder item
    tree: a PySimpleGUI treedata object

    Returns:
    --------
    Nothing
    """
    for f in parent.Folders:
        tree.Insert(parent=parent.Name, key=f.Name, text=f.Name,
                    values=[f.StoreID, f.EntryID], icon=folder_icon)
        # recursive so all folders and subfolders are added
        add_folders(f)


# Add all outlook folders to the tree
add_folders(mailbox)


# ------ Define the layout of the window with the Tree ------ #
layout = [[sg.Text('Outlook folder browser\nSelect folder with pdfs to merge',
                   size=(50, 2)),
           sg.Text('Select pdfs to merge',
                   size=(40, 2))],
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
                      size=(50, 24),
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
    pdflist = list(chain.from_iterable(pdflist))
    pdflist = sorted(pdflist, key=attrgetter('filename.filename'))

    return pdflist

# To use GetItemFromID you need the entryID of the Item (e.g. MailItem)
# The StoreID of the folder it came from is optional


def group_pdfs(pdfs):
    """
    Given a list of pdf NamedTuples (as defined above), the function will
    return a list of lists of pdfs grouped by region/frequency/{sec|appendix}
    Thus, each sublist will be an ordered list of pdfs to be merged into single docs

    Parameters:
    -----------
    pdfs: a list of pdf NamedTuples
    
    Returns:
    --------
    A list of lists of pdfs in page order, each sublist a separate document

    """
    
    # Sort the list of filenames by region/freq/appendex
    pdfs.sort(key=lambda x: [x.filename.region, x.filename.freq, x.filename.appendix])
    
    # Get unique pairs of (freq, region)
    merge_these = {(x.filename.freq, x.filename.region) for x in pdfs}
    
    # Split all files into separate lists by freq/region
    result = [[x for x in pdfs if (x.filename.freq, x.filename.region) == m] for m in merge_these]
    
    return result

def send_mail(send_to, send_cc=None, subject='Test', attachment_path=None, send_bcc=None, body_text='Test mail'):
    """
    Create an outlook mail and prepare to send it
    """
    new_mail = outlook.CreateItem(0x0) # 0x0='olMailItem'
    new_mail.To = send_to
    
    if send_cc:
        new_mail.CC = send_cc
    else:
        new_mail.CC = ''
        
    if subject:
        new_mail.Subject = subject
    else:
        new_mail.Subject = ''
    
    if send_bcc:
        new_mail.BCC = send_bcc
    else:
        new_mail.BCC = ''
        
    if attachment_path:
        new_mail.Attachments.Add(attachment_path)
    
    if body_text:
        new_mail.Body = body_text
    else:
        new_mail.Body = ''
    
    new_mail.Display(False)


def merge_pdf(pdf_files):
    """
    Iterates through a list of namedtuples of type=pdf(filename, entryID) and merges them into a single file.
    All the files are attachments to Outlook messages stored in a single Outlook folder
    
    Parameters:
    -----------
    pdf_files: a list of namedtuples of type=pdf
    
    Returns:
    --------
    The path to a single pdf file saved at the chosen location
    """
    
    # save_folder = sg.PopupGetFolder('Choose merged pdf file save location')
    
    merged_filename = ' '.join([pdf_files[0].filename.freq, pdf_files[0].filename.region])
    
    if not merged_filename.endswith('.pdf'):
        merged_filename += '.pdf'
        
    print (f'Working on {merged_filename}')
    
    merger = PdfFileMerger()
    for f in pdf_files:
        # ------ first create a reference to the message the pdf file is stored in ------ #
        msg = mapi.GetItemFromID(f.entryID)  # storeID is optional
        # ------ then get the pdf file ------ #
        pdf_file = msg.Attachments.Item(f.index)
        # ------ save the pdf file and then add it to the merger job ------ #
        pdf_file.SaveAsFile(os.path.join(save_folder, f.filename.filename))
        merger.append(fileobj=os.path.join(save_folder, f.filename.filename))
        
    # ------ save the merged file ------ #
    merger.write(os.path.join(save_folder, merged_filename))
    merger.close()
    #sg.Popup('MergedPdf saved at {}'.format(save_folder))
    
    # ------ ask whether to keep the source pdf files or delete them ------ #
    del_choice = sg.PopupYesNo('Delete original pdfs from saved location?')
    if del_choice == 'Yes':
        for f in pdf_files:
            os.remove(os.path.join(save_folder, f.filename))
        sg.Popup('Source pdfs have been deleted')
        
    return os.path.join(save_folder, merged_filename)


def import_mailing_list(mailing_list=None):
    """
    Reads a csv file into a pandas dataframe and returns that dataframe
    
    Parameters:
    -----------
    mailing_list: the filename and path location of the mailing list.
        The file headers must be:
            To:
            Cc:
            Bcc:
            Addressee
            Report Name
            Team Name
    
    Returns:
    --------
    pandas dataframe
    """
    
    # if no path is supplied then ask for a specific file
    if not mailing_list:
        email_data = sg.PopupGetFile('Select csv file wth email data', 
                                     default_path=os.path.join(os.path.expandvars('%userprofile%'),
                                                               'Documents',
                                                               'emaildata.csv'))
    else:
        email_data = mailing_list
        
    result = pd.read_csv(email_data)
    return result


def create_pdf_canvas(source, output):
    """
    Creates a blank pdf file with the same number of pages as the source file.
    Each page will be the same size as the source file and numbered at the 
    bottom centre of each page
    
    The purpose is to add page numbers to files that don't have them

    Args:
        source (pdf file): a file opened with PdfFileReader
        output (string): the name of the output file
    """
    
    # Create the base canvas the pdf doc
    c = canvas.Canvas(output)
    
    # Add a page for each page in the source to the base
    for i in range(source.getNumPages()) :
        c.setFont('Helvetica', 7)
        w = source.pages[i].mediaBox[2]
        h = source.pages[i].mediaBox[3]
        c.setPageSize((w, h))
        
        # add the page number
        if i > 0:
            c.drawString(w//2, (4)*mm, str(i)) 
        c.showPage()
    c.save()
    return


def combine_merged_canvas(basepdf, pagespdf):
    """
    Combines the single pdf file created with merge_pdf() with the blank canvas
    pdf with page numbers only from create_pdf_canvas

    Args:
        basepdf (pdf file): merged_pdf filepath and filename
        pagespdf (pdf file): pdf_canvas filepath and filename
    
    Returns:
        overwrites basepdf
    """
    
    # read the pdf files
    corePdf = PdfFileReader(basepdf)
    numberPdf = PdfFileReader(pagespdf)
    
    # combine the two pdf files
    output = PdfFileWriter()
    for p in range(corePdf.getNumPages()):
        page = corePdf.getPage(p)
        numberLayer = numberPdf.getPage(p)
        page.mergePage(numberLayer)
        output.addPage(page)
        
    with open(basepdf, 'wb') as f:
        output.write(f)
    
    
def create_merged_pdf_and_mail(pdfs):
    """
    Performs all the following actions:
    1. Groups all pdfs into those that should be joined together
    2. Merges them into separate files
    3. Attaches merged file to a mail for specific recipients
    
    Parameters:
    -----------
        pdfs: a list of NamedTuples of pdfs chosen to go through this process
    """        
    
    # 1. Import the mailing list
    email_list = import_mailing_list()
    
    # 2. Get list of lists of pdfs grouped by freq & region
    grp_pdfs = group_pdfs(pdfs)
    
    # 3. Join pdfs in each group and attach to a new mail
    for gp in grp_pdfs:
        print(f'There are {len(gp)} pdf files to merge...')
        single_pdf = merge_pdf(gp)
        outputfilename = os.path.join(save_folder, 'pdfpages.pdf')
        create_pdf_canvas(PdfFileReader(single_pdf), outputfilename)
        combine_merged_canvas(single_pdf, outputfilename)
        
        freq = gp[0].filename.freq
        reg = gp[0].filename.region
        
        if reg.endswith('.pdf'):
            reg = reg[:-4]
        
        mask = ((email_list['Report Name'] == freq) & (email_list['Team Name'] == reg))
        print(f'Number of filess with corresponding Report Name & Team Name is {mask.sum()}')
        
        if mask.sum() == 0:
            print(f'''There are no combinations of
                  Report Name = {freq}
                  Team Name = {reg}\n''')
        else:
            email_to = email_list.loc[mask, 'To:'].values[0]
            email_cc = email_list.loc[mask, 'Cc:'].values[0]
            email_bcc = email_list.loc[mask, 'Bcc:'].values[0]
            email_addressee = email_list.loc[mask, 'Addressee'].values[0]
            
            print(f'''
                  Frequency: {freq}
                  Region: {reg}
                  To: {email_to})
                  Cc: {email_cc}
                  Bcc: {email_bcc}
                  Addressee: {email_addressee}
                  ''')
            
            mail_text = 'Dear {0}\n\nPlease find attached your {1} MI.\n\nKind Regards,\n\n{2}'.format(
                email_addressee,
                freq.split()[0],
                firstname
            )
            
            send_mail(send_to=email_to,
                      send_cc=email_cc,
                      send_bcc=email_bcc,
                      subject=' '.join([freq, reg]),
                      attachment_path=single_pdf,
                      body_text=mail_text)
            

# ------ Event loop to catch selected Outlook folder and update the listbox ------ #
while True:
    event, values = window.read()
    if event in (None, 'Cancel'):
        window.close()
        sys.exit()
    if event == 'Get pdfs':
        storeID, entryID = TreeDict[values['-TREE-'][0]].values
        pdf_list = get_pdfs(mapi.GetFolderFromID(entryID, storeID))
        window['-FILES-'].update([x.filename.filename for x in pdf_list])
        window['-MERGE-'].update(disabled=False)
    if event == '-MERGE-':
        merging_pdfs = ([x for x in pdf_list if x.filename.filename in values['-FILES-']])
        create_merged_pdf_and_mail(merging_pdfs)
        window.close()
        sys.exit()


 
# work_folder = outlook.GetFolderFromID(entryID, storeID)
# print('There are {} messages in this folder'.format(work_folder.Items.Count))

# work_folder_Restricted = work_folder.Items.Restrict(
#     "[ReceivedTime] >= '23/12/2019'")  # December 23, 2019
# print('There are {} messages in this folder since 23rd Dec 2019'.format(
#     work_folder_Restricted.Count))
# for msg in work_folder_Restricted:
#     print(msg.Subject)

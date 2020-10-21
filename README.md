# PDF Merger and Mailer
 
 This script first creates a tree of all your Microsoft Outlook folders. Once
 one of them is selected it will go through all the mails in that folder and
 provide a list of all the pdf attachments.

 The purpose of the program was originally to capture individual pdf pages that 
 were emailed and then collate the pages into a single doc. The pdf page 
 filenames were formatted in this form:

 {frequency} + {Sec | Appendix} + {part} + {region}
    e.g. "Monthly Sec 1 EMEA"

 So, continuing from there then all the attachments are merged into single
 pdf files, grouped by common frequency & region.

 Lastly, the pages were sent on email to specific persons the list of which is 
 kept in a csv file. The column headers of that file are:
    To:
    Cc:
    Bcc:
    Addressee
    Report Name
    Team Name

This python program required a number of interesting packages:
    pysimplegui: for creating the outlook folder tree and pop-up messages
    win32api & win32com.client: for connecting with the Outlook COM interface
    PyPDF2 & reportlab: for reading, writing and creating pdf files
# Excel VBA for Sending Outlook Appointment Calendar Invites

## Purpose

During exam season, immense amounts of time are dedicated to sending out 
Outlook Calendar appointment invitations. Each participant of the 
meeting has to be looked up and their E-mail address added individually. 
Furthermore, while templates are available for the various invites, much 
of the content has to be customized, including the meeting dates / time, 
the meeting room (location), the list of instructors & support staff, 
list of support rooms (for larger exams) etc. 

While innate variations 
between exams exist, and some customizations are probably always present, 
much of the content could still be automated to minimize the amount of time 
spent searching, scrolling and pasting information. This set of VBA code 
is meant to be imported into a Macro-Enabled Excel file containing exam 
information, and it is hoped that most of the drafting work for 
Calendar Invites can be automated by the code included.

## Prepare Your Excel for VBA

1. Open a blank Excel File
1. Save as a Macro-Enabled Spreadsheet
1. Go to "File"
    1. => "Save As"
    1. => choose "Excel Macro-Enabled Workbook (*.xlsm)"
1. Enable Developer Ribbon
    1. Go to "File"
    1. => "Options"
    1. => "Customize Ribbon"
    1. => in menu, find and tick checkmark for "Developer"
    1. => "Ok"
1. Enable References for Excel VBAProject
    1. Go to "Developer"
    1. => Visual Basic
    1. => "Tools"
    1. => "References"
    1. => ensure the following are checked:
        1. "Visual Basic For Applications"
        1. "Microsoft Excel 16.0 Object Library"
        1. "Microsoft Office 16.0 Object Library"
        1. "Microsoft Outlook 16.0 Object Library"
        1. "OLE Automation"
1. Import .bas File into Project
    1. In VBA window, go to "File"
    1. => "Import File"
    1. => choose the file from this repo

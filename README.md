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
spent searching, scrolling and pasting information. The Macro-Containing 
Spreadsheet in this repo includes VBA code that automates much of these 
procedures.

## How to Use
1. Ensure Outlook is open
1. Download the Excel File in this repo "send_exam_invites.xlsm" 
(or fetch/pull repo through GitHub)
1. Open "send_exam_invites.xlsm"
1. If you see a warning ribbon on top, click "Enable Macro" or "Enable Content"
1. Open your browser, go to the google spreadsheet for Sauder Exams
    - Ask LS Ops for link/access if you don't have it
1. Go to the worksheet for the relevant term in the **google sheet**
1. Copy content in the **google sheet worksheet** 
1. Paste content from **google sheet** into "Exam Sheet" in "send_exam_invites.xlsm"
    - Ensure header is in first row of spreadsheet
1. Go to the "Mail List" worksheet in the **google sheet**
1. Copy content in the worksheet
1. Paste content from "Mail List" from **google sheet** 
into "Mail List" in "send_exam_invites.xlsm"
1. Check that formatting in "send_exam_invites.xlsm" is up to standards
1. Once ready, select the courses in the first column that needs invite drafted
1. Click on "RECOMMENDED: Draft Selected" in top left
    - Alternatively: click "Draft / Send with Custom Settings" to customize settings
1. Calendar invites will be made
1. Once done, check the "CALENDAR INVITE" column to see if any courses "FAILED"
1. Check the "PEOPLE INVITED" column to see who was found for the invite
1. Update the "CALENDAR INVITE" column on **google sheet** (can copy and paste)
1. BEFORE EXAM: make necessary tweaks and double check info are correct

## Importing VBA into Blank Spreadsheet
These steps are only for importing the code into a blank sheet. 
If you are unsure what this means then only refer to "How to Use" to run the spreadsheet.
1. Pull this repo
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
        1. "Microsoft Scripting Runtime"
        1. "OLE Automation"
        1. "Miscrosoft VBScript Regular Expression 5.5"
1. Import .bas and .cls files into Project
    1. In VBA window, go to "File"
    1. => "Import File"
    1. => Choose the files from the "src" directory pulled out of this repo
    1. => Files have to be imported one by one
1. Done

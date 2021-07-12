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
1. Download the Excel File in this repo "send_exam_invites.xlsm" 
(or fetch/pull repo through GitHub)
1. Open "send_exam_invites.xlsm"
1. If you see a warning ribbon on top, click "Enable Macro" or "Enable Content"
1. Open your browser, go to the google spreadsheet for Sauder Exams
    - Ask LS Ops for link/access if you don't have it
1. Go to the worksheet for the relevant term in the **google sheet**
1. Copy content in the worksheet 
1. Paste content from **google sheet** into "Exam Sheet" in "send_exam_invites.xlsm"
1. Go to the "Mail List" worksheet in the **google sheet**
1. Copy content in the worksheet
1. Paste content from "Mail List" from **google sheet** 
into "Mail List" in "send_exam_invites.xlsm"
1. Check that formatting in "send_exam_invites.xlsm" is up to standards
1. On the right end of "Exam Sheet" in "send_invites.xlsm", 
there's a button to "Send / Save / Update Invite(s)"
1. Once you're ready, click the button and follow instructions to send invites
1. After program has finished running:
    1. Copy "Exam Sheet" content back to **google sheet**
    1. At the very least, copy the "CALENDAR INVITE" column back to **google sheet**
    1. Check outlook to see if invites are good
    1. Report to team any errors, either in the invite or when program was running
1. Note that after running the program, spreadsheet 
"send_exam_invites.xlsm" is saved into an "output" folder


## Importing VBA into Blank Spreadsheet
These steps are only for importing the code into a blank sheet. 
If you are unsure what this means then only refer to "How to Use" to run the spreadsheet.
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
        1. "Miscrosoft VBScript Regular Expression 5.5"
1. Import .bas File into Project
    1. In VBA window, go to "File"
    1. => "Import File"
    1. => choose the file from this repo

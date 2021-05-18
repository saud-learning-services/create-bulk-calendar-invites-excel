#Excel VBA for Sending Outlook Appointment Calendar Invites
##Prepare Your Excel for VBA
<ol>
    <li>Open a blank Excel File</li>
    <li>Save as a Macro-Enabled Spreadsheet
        <ol>
            <li>Go to "File"</li>
            <li>... "Save As"</li>
            <li>... choose "Excel Macro-Enabled Workbook (*.xlsm)"</li>
        </ol>
    </li>
    <li>Enable Developer Ribbon
        <ol>
            <li>Go to "File"</li>
            <li>... "Options"</li>
            <li>... "Customize Ribbon"</li>
            <li>... in menu, find and tick checkmark for "Developer"</li>
            <li>... "Ok"</li>
        </ol>
    </li>
    <li>Enable References for Excel VBAProject
        <ol>
            <li>Go to "Developer"</li>
            <li>... Visual Basic</li>
            <li>... "Tools"</li>
            <li>... "References"</li>
            <li>... ensure the following are checked:
                <ul>
                    <li>"Visual Basic For Applications"</li>
                    <li>"Microsoft Excel 16.0 Object Library"</li>
                    <li>"Microsoft Office 16.0 Object Library"</li>
                    <li>"Microsoft Outlook 16.0 Object Library"</li>
                    <li>"OLE Automation"</li>
                </ul>
            </li>
        </ol>
    </li>
    <li>Import .bas File into Project
        <ol>
            <li>In VBA window, go to "File"</li>
            <li>... "Import File"</li>
            <li>... choose the file from this repo</li>
        </ol>
    </li>
</ol>
##Requirements
In order to make use of the code, the following MUST be enabled 
in your Excel VBA.

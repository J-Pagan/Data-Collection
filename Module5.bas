Attribute VB_Name = "Module5"
Sub phytochemical_identifier_extract()

Dim I As Long

'Keeps track of the Excel rows
Dim rowcounter As Long
rowcounter = 3

For I = 1 To 9
    'Creates new URLS to navigate to each phytochemical's webpage
    Url = "https://cb.imsc.res.in/imppat/phytochemical-detailedpage/IMPHY00000" & I
    
    Dim ig As Object
    
    'Loads the webpage in Internet Explorer
    Set ig = CreateObject("InternetExplorer.Application")
    ig.Visible = False
    ig.navigate Url
    
    'Loops until the webpage is fully loaded
    Do While ig.Busy: DoEvents: Loop
    Do Until ig.readyState = 4: DoEvents: Loop
    
    'Saves the Excel sheet to a variable
    Dim mys As Worksheet
    Set mys = ThisWorkbook.Sheets("Sheet3")
    
    'Saves the phytochemical summary data to a variable
    Dim topInfo As MSHTML.IHTMLElementCollection
    Dim Top As Object
    Set topInfo = ig.document.getElementsByClassName("col-8 pt-0 mt-0 ml-2 pl-2")
    
    'Notes the phytochemical identifiers before every extraction
    mys.Cells(rowcounter - 2, 1).Value = "IMPPAT Phytochemical identifier: IMPHY00000" & I
    
    'Saves the summary data to the Excel sheet
    On Error Resume Next
    For Each Top In topInfo
        mys.Cells(rowcounter - 1, 1).Value = Top.innerText
    Next Top
    
    rowcounter = rowcounter + 3
    
    'Quits Internet Explorer
    ig.Quit
Next I

For I = 10 To 99
    'Creates new URLS to navigate to each phytochemical's webpage
    Url = "https://cb.imsc.res.in/imppat/phytochemical-detailedpage/IMPHY0000" & I
    
    'Loads the webpage in Internet Explorer
    Set ig = CreateObject("InternetExplorer.Application")
    ig.Visible = False
    ig.navigate Url
    
    'Loops until the webpage is fully loaded
    Do While ig.Busy: DoEvents: Loop
    Do Until ig.readyState = 4: DoEvents: Loop
    
    'Saves the Excel sheet to a variable
    Set mys = ThisWorkbook.Sheets("Sheet3")
    
    'Saves the phytochemical summary data to a variable
    Set topInfo = ig.document.getElementsByClassName("col-8 pt-0 mt-0 ml-2 pl-2")
    
    'Notes the phytochemical identifiers before every extraction
    mys.Cells(rowcounter - 2, 1).Value = "IMPPAT Phytochemical identifier: IMPHY0000" & I
    
    'Saves the summary data to the Excel sheet
    On Error Resume Next
    For Each Top In topInfo
        mys.Cells(rowcounter - 1, 1).Value = Top.innerText
    Next Top
    
    rowcounter = rowcounter + 3
    
    'Quits Internet Explorer
    ig.Quit
Next I

For I = 100 To 999
    'Creates new URLS to navigate to each phytochemical's webpage
    Url = "https://cb.imsc.res.in/imppat/phytochemical-detailedpage/IMPHY000" & I
    
    'Loads the webpage in Internet Explorer
    Set ig = CreateObject("InternetExplorer.Application")
    ig.Visible = False
    ig.navigate Url
    
    'Loops until the webpage is fully loaded
    Do While ig.Busy: DoEvents: Loop
    Do Until ig.readyState = 4: DoEvents: Loop
    
    'Saves the Excel sheet to a variable
    Set mys = ThisWorkbook.Sheets("Sheet3")
    
    'Saves the phytochemical summary data to a variable
    Set topInfo = ig.document.getElementsByClassName("col-8 pt-0 mt-0 ml-2 pl-2")
    
    'Notes the phytochemical identifiers before every extraction
    mys.Cells(rowcounter - 2, 1).Value = "IMPPAT Phytochemical identifier: IMPHY000" & I
    
    'Saves the summary data to the Excel sheet
    On Error Resume Next
    For Each Top In topInfo
        mys.Cells(rowcounter - 1, 1).Value = Top.innerText
    Next Top
    
    rowcounter = rowcounter + 3
    
    'Quits Internet Explorer
    ig.Quit
Next I

For I = 1000 To 9999
    'Creates new URLS to navigate to each phytochemical's webpage
    Url = "https://cb.imsc.res.in/imppat/phytochemical-detailedpage/IMPHY00" & I
    
    'Loads the webpage in Internet Explorer
    Set ig = CreateObject("InternetExplorer.Application")
    ig.Visible = False
    ig.navigate Url
    
    'Loops until the webpage is fully loaded
    Do While ig.Busy: DoEvents: Loop
    Do Until ig.readyState = 4: DoEvents: Loop
    
    'Saves the Excel sheet to a variable
    Set mys = ThisWorkbook.Sheets("Sheet3")
    
    'Saves the phytochemical summary data to a variable
    Set topInfo = ig.document.getElementsByClassName("col-8 pt-0 mt-0 ml-2 pl-2")
    
    'Notes the phytochemical identifiers before every extraction
    mys.Cells(rowcounter - 2, 1).Value = "IMPPAT Phytochemical identifier: IMPHY00" & I
    
    'Saves the summary data to the Excel sheet
    On Error Resume Next
    For Each Top In topInfo
        mys.Cells(rowcounter - 1, 1).Value = Top.innerText
    Next Top
    
    rowcounter = rowcounter + 3
    
    'Quits Internet Explorer
    ig.Quit
Next I

'There are 17,967 phytochemicals in the database
For I = 10000 To 17967
    'Creates new URLS to navigate to each phytochemical's webpage
    Url = "https://cb.imsc.res.in/imppat/phytochemical-detailedpage/IMPHY0" & I
    
    'Loads the webpage in Internet Explorer
    Set ig = CreateObject("InternetExplorer.Application")
    ig.Visible = False
    ig.navigate Url
    
    'Loops until the webpage is fully loaded
    Do While ig.Busy: DoEvents: Loop
    Do Until ig.readyState = 4: DoEvents: Loop
    
    'Saves the Excel sheet to a variable
    Set mys = ThisWorkbook.Sheets("Sheet3")
    
    'Saves the phytochemical summary data to a variable
    Set topInfo = ig.document.getElementsByClassName("col-8 pt-0 mt-0 ml-2 pl-2")
    
    'Notes the phytochemical identifiers before every extraction
    mys.Cells(rowcounter - 2, 1).Value = "IMPPAT Phytochemical identifier: IMPHY0" & I
    
    'Saves the summary data to the Excel sheet
    On Error Resume Next
    For Each Top In topInfo
        mys.Cells(rowcounter - 1, 1).Value = Top.innerText
    Next Top
    
    rowcounter = rowcounter + 3
    
    'Quits Internet Explorer
    ig.Quit
Next I

End Sub

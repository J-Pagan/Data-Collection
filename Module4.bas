Attribute VB_Name = "Module4"
Sub dropdown_menu_extract2()

Dim IE As New SHDocVw.InternetExplorer
Dim HTMLDoc As New MSHTML.HTMLDocument
Dim list As MSHTML.IHTMLElement
Dim I As Long

'Hides the web browser
IE.Visible = False
'Opens the medicinal plant's page on Internet Explorer
IE.navigate "https://cb.imsc.res.in/imppat/"

'Loops until the webpage is fully loaded
Do While IE.Busy Or IE.readyState <> 4
    Application.Wait DateAdd("s", 1, Now)
Loop

'Saves the HTML code to a variable
Set HTMLDoc = IE.document

'Saves the elements in the dropdown menu to a variable
Set list = HTMLDoc.getElementsByClassName("homeselect form-control")(1)

'Keeps track of the Excel rows and columns
Dim rowcounter As Long
Dim columncounter As Long
rowcounter = 4
columncounter = 2

'Loops through dropdown menu elements
For I = 1 To 3000
    'Creates new URLS to navigate to each dropdown menu option's webpage
    newURL = "https://cb.imsc.res.in/imppat/phytochemicalplants/" & list.Options(I).Text
    
    Dim ig As Object
    
    'Loads the webpage in Internet Explorer
    Set ig = CreateObject("InternetExplorer.Application")
    ig.Visible = False
    ig.navigate newURL
    
    'Loops until the webpage is fully loaded
    Do While ig.Busy: DoEvents: Loop
    Do Until ig.readyState = 4: DoEvents: Loop
    
    'Saves the table data to a variable
    Dim tb As HTMLTable
    Set tb = ig.document.getElementById("table_id")
    
    Dim tro As HTMLTableRow
    Dim tdc As HTMLTableCell
    Dim thu
    Dim mys As Worksheet
    
    'Saves the Excel sheet to a variable
    Set mys = ThisWorkbook.Sheets("Sheet2")
    
    'Notes the phytochemicals by ID before every extraction
    mys.Cells(rowcounter - 1, 1).Value = "Phytochemical ID#: " & I
    
    'Extracts the table data from the page using HTML tags
    On Error Resume Next
    For Each tro In tb.getElementsByTagName("tr")
        For Each thu In tro.getElementsByTagName("th")
            mys.Cells(rowcounter, columncounter).Value = thu.innerText
            columncounter = columncounter + 1
        Next thu
        For Each tdc In tro.getElementsByTagName("td")
            mys.Cells(rowcounter, columncounter).Value = tdc.innerText
            columncounter = columncounter + 1
        Next tdc
        columncounter = 1
        rowcounter = rowcounter + 1
    Next tro
    rowcounter = rowcounter + 2
    
    'Quits Internet Explorer
    ig.Quit
Next I

For I = 3001 To 6000
    'Creates new URLS to navigate to each dropdown menu option's webpage
    newURL = "https://cb.imsc.res.in/imppat/phytochemicalplants/" & list.Options(I).Text
    
    'Loads the webpage in Internet Explorer
    Set ig = CreateObject("InternetExplorer.Application")
    ig.Visible = False
    ig.navigate newURL
    
    'Loops until the webpage is fully loaded
    Do While ig.Busy: DoEvents: Loop
    Do Until ig.readyState = 4: DoEvents: Loop
    
    'Saves the table data to a variable
    Set tb = ig.document.getElementById("table_id")
    
    'Saves the Excel sheet to a variable
    Set mys = ThisWorkbook.Sheets("Sheet2")
    
    'Notes the phytochemicals by ID before every extraction
    mys.Cells(rowcounter - 1, 1).Value = "Phytochemical ID#: " & I
    
    'Extracts the table data from the page using HTML tags
    On Error Resume Next
    For Each tro In tb.getElementsByTagName("tr")
        For Each thu In tro.getElementsByTagName("th")
            mys.Cells(rowcounter, columncounter).Value = thu.innerText
            columncounter = columncounter + 1
        Next thu
        For Each tdc In tro.getElementsByTagName("td")
            mys.Cells(rowcounter, columncounter).Value = tdc.innerText
            columncounter = columncounter + 1
        Next tdc
        columncounter = 1
        rowcounter = rowcounter + 1
    Next tro
    rowcounter = rowcounter + 2
    
    'Quits Internet Explorer
    ig.Quit
Next I

For I = 6001 To 9000
    'Creates new URLS to navigate to each dropdown menu option's webpage
    newURL = "https://cb.imsc.res.in/imppat/phytochemicalplants/" & list.Options(I).Text
    
    'Loads the webpage in Internet Explorer
    Set ig = CreateObject("InternetExplorer.Application")
    ig.Visible = False
    ig.navigate newURL
    
    'Loops until the webpage is fully loaded
    Do While ig.Busy: DoEvents: Loop
    Do Until ig.readyState = 4: DoEvents: Loop
    
    'Saves the table data to a variable
    Set tb = ig.document.getElementById("table_id")
    
    'Saves the Excel sheet to a variable
    Set mys = ThisWorkbook.Sheets("Sheet2")
    
    'Notes the phytochemicals by ID before every extraction
    mys.Cells(rowcounter - 1, 1).Value = "Phytochemical ID#: " & I
    
    'Extracts the table data from the page using HTML tags
    On Error Resume Next
    For Each tro In tb.getElementsByTagName("tr")
        For Each thu In tro.getElementsByTagName("th")
            mys.Cells(rowcounter, columncounter).Value = thu.innerText
            columncounter = columncounter + 1
        Next thu
        For Each tdc In tro.getElementsByTagName("td")
            mys.Cells(rowcounter, columncounter).Value = tdc.innerText
            columncounter = columncounter + 1
        Next tdc
        columncounter = 1
        rowcounter = rowcounter + 1
    Next tro
    rowcounter = rowcounter + 2
    
    'Quits Internet Explorer
    ig.Quit
Next I

For I = 9001 To 12000
    'Creates new URLS to navigate to each dropdown menu option's webpage
    newURL = "https://cb.imsc.res.in/imppat/phytochemicalplants/" & list.Options(I).Text
    
    'Loads the webpage in Internet Explorer
    Set ig = CreateObject("InternetExplorer.Application")
    ig.Visible = False
    ig.navigate newURL
    
    'Loops until the webpage is fully loaded
    Do While ig.Busy: DoEvents: Loop
    Do Until ig.readyState = 4: DoEvents: Loop
    
    'Saves the table data to a variable
    Set tb = ig.document.getElementById("table_id")
    
    'Saves the Excel sheet to a variable
    Set mys = ThisWorkbook.Sheets("Sheet2")
    
    'Notes the phytochemicals by ID before every extraction
    mys.Cells(rowcounter - 1, 1).Value = "Phytochemical ID#: " & I
    
    'Extracts the table data from the page using HTML tags
    On Error Resume Next
    For Each tro In tb.getElementsByTagName("tr")
        For Each thu In tro.getElementsByTagName("th")
            mys.Cells(rowcounter, columncounter).Value = thu.innerText
            columncounter = columncounter + 1
        Next thu
        For Each tdc In tro.getElementsByTagName("td")
            mys.Cells(rowcounter, columncounter).Value = tdc.innerText
            columncounter = columncounter + 1
        Next tdc
        columncounter = 1
        rowcounter = rowcounter + 1
    Next tro
    rowcounter = rowcounter + 2
    
    'Quits Internet Explorer
    ig.Quit
Next I

For I = 12001 To 15000
    'Creates new URLS to navigate to each dropdown menu option's webpage
    newURL = "https://cb.imsc.res.in/imppat/phytochemicalplants/" & list.Options(I).Text
    
    'Loads the webpage in Internet Explorer
    Set ig = CreateObject("InternetExplorer.Application")
    ig.Visible = False
    ig.navigate newURL
    
    'Loops until the webpage is fully loaded
    Do While ig.Busy: DoEvents: Loop
    Do Until ig.readyState = 4: DoEvents: Loop
    
    'Saves the table data to a variable
    Set tb = ig.document.getElementById("table_id")
    
    'Saves the Excel sheet to a variable
    Set mys = ThisWorkbook.Sheets("Sheet2")
    
    'Notes the phytochemicals by ID before every extraction
    mys.Cells(rowcounter - 1, 1).Value = "Phytochemical ID#: " & I
    
    'Extracts the table data from the page using HTML tags
    On Error Resume Next
    For Each tro In tb.getElementsByTagName("tr")
        For Each thu In tro.getElementsByTagName("th")
            mys.Cells(rowcounter, columncounter).Value = thu.innerText
            columncounter = columncounter + 1
        Next thu
        For Each tdc In tro.getElementsByTagName("td")
            mys.Cells(rowcounter, columncounter).Value = tdc.innerText
            columncounter = columncounter + 1
        Next tdc
        columncounter = 1
        rowcounter = rowcounter + 1
    Next tro
    rowcounter = rowcounter + 2
    
    'Quits Internet Explorer
    ig.Quit
Next I

'17968 options in the dropdown menu
For I = 15001 To list.Options.Length
    'Creates new URLS to navigate to each dropdown menu option's webpage
    newURL = "https://cb.imsc.res.in/imppat/phytochemicalplants/" & list.Options(I).Text
    
    'Loads the webpage in Internet Explorer
    Set ig = CreateObject("InternetExplorer.Application")
    ig.Visible = False
    ig.navigate newURL
    
    'Loops until the webpage is fully loaded
    Do While ig.Busy: DoEvents: Loop
    Do Until ig.readyState = 4: DoEvents: Loop
    
    'Saves the table data to a variable
    Set tb = ig.document.getElementById("table_id")
    
    'Saves the Excel sheet to a variable
    Set mys = ThisWorkbook.Sheets("Sheet2")
    
    'Notes the phytochemicals by ID before every extraction
    mys.Cells(rowcounter - 1, 1).Value = "Phytochemical ID#: " & I
    
    'Extracts the table data from the page using HTML tags
    On Error Resume Next
    For Each tro In tb.getElementsByTagName("tr")
        For Each thu In tro.getElementsByTagName("th")
            mys.Cells(rowcounter, columncounter).Value = thu.innerText
            columncounter = columncounter + 1
        Next thu
        For Each tdc In tro.getElementsByTagName("td")
            mys.Cells(rowcounter, columncounter).Value = tdc.innerText
            columncounter = columncounter + 1
        Next tdc
        columncounter = 1
        rowcounter = rowcounter + 1
    Next tro
    rowcounter = rowcounter + 2
    
    'Quits Internet Explorer
    ig.Quit
Next I

'Quits Internet Explorer
IE.Quit

End Sub


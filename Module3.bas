Attribute VB_Name = "Module3"
Sub dropdown_menu_extract()

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
Set list = HTMLDoc.getElementsByClassName("homeselect form-control")(0)

'Keeps track of the Excel rows and columns
Dim rowcounter As Long
Dim columncounter As Long
rowcounter = 4
columncounter = 2

'Loops through dropdown menu elements
For I = 1 To 1000
    'Splits each option separated by a space into two elements in the array
    testArray = Split(list.Options(I).Text, " ")
    'Creates new URLS to navigate to each dropdown menu option's webpage
    newURL = "https://cb.imsc.res.in/imppat/phytochemical/" & testArray(0) & "%20" & testArray(1)
    
    Dim ig As Object
    
    'Loads the webpage in Internet Explorer
    Set ig = CreateObject("InternetExplorer.Application")
    ig.Visible = False
    ig.navigate newURL
    
    'Loops until the webpage is fully loaded
    Do While ig.Busy: DoEvents: Loop
    Do Until ig.readyState = 4: DoEvents: Loop
    
    Dim tb As HTMLTable
    Dim topInfo As MSHTML.IHTMLElementCollection
    Dim Top As Object
    
    'Saves the table data to a variable
    Set tb = ig.document.getElementById("table_id")
    'Saves the data above the table into a variable
    Set topInfo = ig.document.getElementsByClassName("col-lg-8")
    
    Dim tro As HTMLTableRow
    Dim tdc As HTMLTableCell
    Dim thu
    Dim mys As Worksheet
    
    'Saves the Excel sheet to a variable
    Set mys = ThisWorkbook.Sheets("Sheet1")
    
    'Separates the medicinal plants by ID before every extraction
    mys.Cells(rowcounter - 2, 1).Value = "Plant ID#: " & I
    
    'Extracts the data above the table on each page
    On Error Resume Next
    For Each Top In topInfo
        mys.Cells(rowcounter - 1, 1).Value = Top.innerText
    Next Top
    
    'Extracts the table data from the page using HTML tags
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
    rowcounter = rowcounter + 5
    
    'Quits Internet Explorer
    ig.Quit
Next I

For I = 1001 To 2000
    'Splits each option separated by a space into two elements in the array
    testArray = Split(list.Options(I).Text, " ")
    'Creates new URLS to navigate to each dropdown menu option's webpage
    newURL = "https://cb.imsc.res.in/imppat/phytochemical/" & testArray(0) & "%20" & testArray(1)
    
    'Loads the webpage in Internet Explorer
    Set ig = CreateObject("InternetExplorer.Application")
    ig.Visible = False
    ig.navigate newURL
    
    'Loops until the webpage is fully loaded
    Do While ig.Busy: DoEvents: Loop
    Do Until ig.readyState = 4: DoEvents: Loop
    
    'Saves the table data to a variable
    Set tb = ig.document.getElementById("table_id")
    'Saves the data above the table into a variable
    Set topInfo = ig.document.getElementsByClassName("col-lg-8")
    
    'Saves the Excel sheet to a variable
    Set mys = ThisWorkbook.Sheets("Sheet1")
    
    'Separates the medicinal plants by ID before every extraction
    mys.Cells(rowcounter - 2, 1).Value = "Plant ID#: " & I
    
    'Extracts the data above the table on each page
    On Error Resume Next
    For Each Top In topInfo
        mys.Cells(rowcounter - 1, 1).Value = Top.innerText
    Next Top
    
    'Extracts the table data from the page using HTML tags
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
    rowcounter = rowcounter + 5
    
    'Quits Internet Explorer
    ig.Quit
Next I

For I = 2001 To 3000
    'Splits each option separated by a space into two elements in the array
    testArray = Split(list.Options(I).Text, " ")
    'Creates new URLS to navigate to each dropdown menu option's webpage
    newURL = "https://cb.imsc.res.in/imppat/phytochemical/" & testArray(0) & "%20" & testArray(1)
    
    'Loads the webpage in Internet Explorer
    Set ig = CreateObject("InternetExplorer.Application")
    ig.Visible = False
    ig.navigate newURL
    
    'Loops until the webpage is fully loaded
    Do While ig.Busy: DoEvents: Loop
    Do Until ig.readyState = 4: DoEvents: Loop
    
    'Saves the table data to a variable
    Set tb = ig.document.getElementById("table_id")
    'Saves the data above the table into a variable
    Set topInfo = ig.document.getElementsByClassName("col-lg-8")
    
    'Saves the Excel sheet to a variable
    Set mys = ThisWorkbook.Sheets("Sheet1")
    
    'Separates the medicinal plants by ID before every extraction
    mys.Cells(rowcounter - 2, 1).Value = "Plant ID#: " & I
    
    'Extracts the data above the table on each page
    On Error Resume Next
    For Each Top In topInfo
        mys.Cells(rowcounter - 1, 1).Value = Top.innerText
    Next Top
    
    'Extracts the table data from the page using HTML tags
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
    rowcounter = rowcounter + 5
    
    'Quits Internet Explorer
    ig.Quit
Next I

'4011 options in the dropdown menu
For I = 3001 To list.Options.Length
    'Splits each option separated by a space into two elements in the array
    testArray = Split(list.Options(I).Text, " ")
    'Creates new URLS to navigate to each dropdown menu option's webpage
    newURL = "https://cb.imsc.res.in/imppat/phytochemical/" & testArray(0) & "%20" & testArray(1)
    
    'Loads the webpage in Internet Explorer
    Set ig = CreateObject("InternetExplorer.Application")
    ig.Visible = False
    ig.navigate newURL
    
    'Loops until the webpage is fully loaded
    Do While ig.Busy: DoEvents: Loop
    Do Until ig.readyState = 4: DoEvents: Loop
    
    'Saves the table data to a variable
    Set tb = ig.document.getElementById("table_id")
    'Saves the data above the table into a variable
    Set topInfo = ig.document.getElementsByClassName("col-lg-8")
    
    'Saves the Excel sheet to a variable
    Set mys = ThisWorkbook.Sheets("Sheet1")
    
    'Separates the medicinal plants by ID before every extraction
    mys.Cells(rowcounter - 2, 1).Value = "Plant ID#: " & I
    
    'Extracts the data above the table on each page
    On Error Resume Next
    For Each Top In topInfo
        mys.Cells(rowcounter - 1, 1).Value = Top.innerText
    Next Top
    
    'Extracts the table data from the page using HTML tags
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
    rowcounter = rowcounter + 5
    
    'Quits Internet Explorer
    ig.Quit
Next I

'Quits Internet Explorer
IE.Quit

End Sub

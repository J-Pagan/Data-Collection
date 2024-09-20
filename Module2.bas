Attribute VB_Name = "Module2"
Sub pullDataFromWeb()

    Dim IE As Object
    Dim doc As HTMLDocument
    
    Set IE = CreateObject("InternetExplorer.Application")
    'Shows the browser launching
    IE.Visible = True
    
    'Opens the webpage
    IE.navigate "https://www.signupanytime.com/plugins/links/front/linksviews.aspx?v=results&fmt=nohead&ax=2655&t=15859#"
    
    'Loops until the webpage is fully loaded
    Do While IE.Busy Or IE.readyState <> 4
        Application.Wait DateAdd("s", 1, Now)
    Loop
    
    Set doc = IE.document
    
    'Clicks on each option
    doc.getElementsByClassName("player")(0).getElementsByTagName("a")(0).Click
    doc.getElementsByClassName("player")(1).getElementsByTagName("a")(0).Click
    doc.getElementsByClassName("player")(2).getElementsByTagName("a")(0).Click
    
    'Code from: https://stackoverflow.com/questions/47939045/retrieve-data-from-a-table-of-aspx-page-using-excel-vba
    'Column headers
    Set eleColth = doc.getElementsByTagName("th")
    j = 0 'start with the first value in the th collection
            For Each eleCol In eleColth 'for each element in the td collection
                ThisWorkbook.Sheets("Sheet5").Range("A1").Offset(I, j).Value = eleCol.innerText 'paste the inner text of the td element, and offset at the same time
                j = j + 1 'move to next element in td collection
            Next eleCol 'rinse and repeat


    'Content
    Set eleColtr = doc.getElementsByTagName("tr")

    'This section populates Excel
        I = 0 'start with first value in tr collection
        For Each eleRow In eleColtr 'for each element in the tr collection
            Set eleColtd = doc.getElementsByTagName("tr")(I).getElementsByTagName("td") 'get all the td elements in that specific tr
            j = 0 'start with the first value in the td collection
            For Each eleCol In eleColtd 'for each element in the td collection
                ThisWorkbook.Sheets("Sheet5").Range("D3").Offset(I, j).Value = eleCol.innerText 'paste the inner text of the td element, and offset at the same time
                j = j + 1 'move to next element in td collection
            Next eleCol 'rinse and repeat
            I = I + 1 'move to next element in td collection
        Next eleRow 'rinse and repeat

    IE.Quit
    Set IE = Nothing
    
End Sub

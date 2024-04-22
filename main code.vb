Sub ImportAndOrganizeData()
    Dim filePath As String
    Dim textLine As String
    Dim textData() As String
    Dim i As Long, j As Long
    Dim newWorkbook As Workbook
    Dim ws As Worksheet
    Dim itemNumber As String
    Dim termDate As String
    Dim orderQuantity As String
    Dim reference As String
    Dim yearWeek As String
    Dim lastRow As Long

    ' Prompt user to select CSV file
    filePath = Application.GetOpenFilename("Text Files (*.csv),*.csv", , "Please select the file:")
    
    If filePath = "False" Then Exit Sub ' If user cancels, exit sub
    
    ' Create a new workbook
    Set newWorkbook = Application.Workbooks.Add
    Set ws = newWorkbook.Sheets(1)
    
    ' Set column headers
    ws.Cells(1, 1).Value = "Item Number"
    ws.Cells(1, 2).Value = "Term"
    ws.Cells(1, 3).Value = "Order Quantity"
    ws.Cells(1, 4).Value = "Reference"
    ws.Cells(1, 5).Value = "Year-Week"
    
    ' Open file in input mode
    Open filePath For Input As #1
    
    i = 2 ' Start at the second row because the first row is headers
    Do Until EOF(1)
        Line Input #1, textLine ' Read line from file
        textLine = Replace(textLine, ";;", ";") ' Replace double semicolons with single
        textData = Split(textLine, ";") ' Split the line into an array
        
        ' Reset variables for each line
        itemNumber = ""
        termDate = ""
        orderQuantity = ""
        reference = ""
        
        ' Parse each segment in the line based on rules provided
        For j = LBound(textData) To UBound(textData)
            If Len(textData(j)) = 10 And Left(textData(j), 1) = "4" Then
                itemNumber = textData(j)
            ElseIf IsDate(textData(j)) And termDate = "" Then
                termDate = textData(j)
                yearWeek = Format(CDate(termDate), "yyyy") & "-" & Format(CDate(termDate), "ww")
            ElseIf IsNumeric(textData(j)) And orderQuantity = "" And termDate <> "" Then
                orderQuantity = textData(j)
            ElseIf (Len(textData(j)) = 8 And Left(textData(j), 1) = "3") Or textData(j) = "forecast" Then
                reference = textData(j)
            End If
        Next j
        
        ' Write the parsed data to the worksheet
        ws.Cells(i, 1).Value = itemNumber
        ws.Cells(i, 2).Value = termDate
        ws.Cells(i, 3).Value = orderQuantity
        ws.Cells(i, 4).Value = reference
        ws.Cells(i, 5).Value = yearWeek
        
        i = i + 1 ' Move to next row
    Loop
    
    Close #1 ' Close file
    
    ' Remove any completely blank rows as a failsafe
    For i = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row To 2 Step -1
        If Application.CountA(ws.Rows(i)) = 0 Then
            ws.Rows(i).Delete
        End If
    Next i

    ' Find the last row with data after clean-up
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    If lastRow > 1 Then
        ' There is data present, so we can convert the range to a Table
        Dim table As ListObject
        Set table = ws.ListObjects.Add(SourceType:=xlSrcRange, Source:=ws.Range("A1:E" & lastRow), XlListObjectHasHeaders:=xlYes)
        table.Name = "DataTable"
        table.TableStyle = "TableStyleMedium9"
        
        ' Center everything in the table
        table.Range.HorizontalAlignment = xlCenter
        table.Range.VerticalAlignment = xlCenter
        
        ' Apply filters to the top row
        ws.Range("A1:E1").AutoFilter
    Else
        MsgBox "No data was found to format as a table and apply filters.", vbExclamation
    End If
    
    ' Optimize the view
    ws.Columns("A:E").AutoFit

    ' Activate the new workbook
    newWorkbook.Activate
End Sub


Function GetYearWeek(ByVal targetDate As Date) As String
    Dim weekNum As Integer
    weekNum = DatePart("ww", targetDate, vbMonday, vbFirstFourDays)
    GetYearWeek = Year(targetDate) & "-" & Format(weekNum, "00")
End Function

Function CustomXLookup(lookupValue As String, lookupArray As Range, returnArray As Range) As Variant
    Dim matchFound As Boolean
    matchFound = False
    Dim i As Long
    For i = 1 To lookupArray.Cells.Count
        If lookupArray.Cells(i).Value = lookupValue Then
            CustomXLookup = returnArray.Cells(i).Value
            matchFound = True
            Exit For
        End If
    Next i
    If Not matchFound Then CustomXLookup = "N/A"
End Function

Sub UpdateWorksheets()
    On Error GoTo ErrorHandler
    Dim wsHF As Worksheet, wsPINGRM As Worksheet, wsSummary As Worksheet
    Dim lastRow As Long, i As Long
    Dim yearWeek As String, itemYearWeek As String
    Dim matchPING As String, matchHF As String, CO As String

    Debug.Print "UpdateWorksheets started at " & Now

    Set wsHF = ThisWorkbook.Sheets("HF Sheet")
    Debug.Print "wsHF set to 'HF Sheet'"

    Set wsPINGRM = ThisWorkbook.Sheets("PINGRM's Sheet")
    Debug.Print "wsPINGRM set to 'PINGRM's Sheet'"

    Set wsSummary = ThisWorkbook.Sheets("Summary Sheet")
    Debug.Print "wsSummary set to 'Summary Sheet'"

    ' Update HF Sheet
    lastRow = wsHF.Cells(wsHF.Rows.Count, "A").End(xlUp).Row
    Debug.Print "Updating HF Sheet, last row: " & lastRow

    For i = 2 To lastRow ' Assuming row 1 has headers
        yearWeek = GetYearWeek(wsHF.Cells(i, "PromDlvry").Value)
        wsHF.Cells(i, "Year-Week").Value = yearWeek
        itemYearWeek = wsHF.Cells(i, "Item Number").Value & yearWeek
        matchPING = CustomXLookup(itemYearWeek, wsPINGRM.Columns("Ref"), wsPINGRM.Columns("Match?"))
        wsHF.Cells(i, "Match PING?").Value = matchPING
        Debug.Print "Row " & i & ": Year-Week updated to " & yearWeek & "; Match PING? updated to " & matchPING
    Next i

    Debug.Print "Update on HF Sheet completed."

    ' TODO: Insert logic to update PINGRM's Sheet here...
    ' TODO: Insert logic to update Summary Sheet here...

    Debug.Print "UpdateWorksheets completed at " & Now
    Exit Sub

ErrorHandler:
    Debug.Print "Error encountered: " & Err.Description & " (Error " & Err.Number & ")"
    MsgBox "An error occurred: " & Err.Description, vbCritical, "Error " & Err.Number
    Resume Next
End Sub

Sub TrimAndFilterSheet()
    Dim ws As Worksheet
    Dim wb As Workbook
    Dim fd As FileDialog
    Dim selectedFile As String
    Dim lastRow As Long
    Dim col As Range, delCols As Range
    Dim requiredColumns As Variant
    Dim found As Boolean
    Dim i As Long, j As Long

    ' Create and configure the FileDialog object
    Set fd = Application.FileDialog(msoFileDialogOpen)
    With fd
        .AllowMultiSelect = False
        .Title = "Select the Excel file"
        .Filters.Clear
        .Filters.Add "Excel files", "*.xlsx; *.xls; *.xlsm"
        
        ' Show the dialog box to the user
        If .Show = True Then
            selectedFile = .SelectedItems(1)
        Else
            ' Exit the subroutine if no file is selected
            MsgBox "No file selected.", vbExclamation
            Exit Sub
        End If
    End With

    ' Open the selected workbook
    Set wb = Workbooks.Open(selectedFile)
    ' Prompt the user to enter the sheet name or select from list
    Set ws = wb.Sheets(1) ' ERROR: OBJECT REQUIRED

    requiredColumns = Array("CustID", "Customer PONumber", "CONumber", "Ln", "Item Number", "Item Description", "Cust Item Number", "OrderQty", "Open Qty", "PromDlvry")
    
    ' Delete non-required columns
    Application.ScreenUpdating = False
    For i = ws.UsedRange.Columns.Count To 1 Step -1
        found = False
        For j = LBound(requiredColumns) To UBound(requiredColumns)
            If ws.Cells(1, i).Value = requiredColumns(j) Then
                found = True
                Exit For
            End If
        Next j
        If Not found Then
            If delCols Is Nothing Then
                Set delCols = ws.Columns(i)
            Else
                Set delCols = Union(delCols, ws.Columns(i))
            End If
        End If
    Next i
    If Not delCols Is Nothing Then delCols.Delete
    
     ' Add Year-Week and Match PING? columns
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    With ws
        .Cells(1, .UsedRange.Columns.Count + 1).Value = "Year-Week"
        .Cells(1, .UsedRange.Columns.Count + 2).Value = "Match PING?"
    
        ' Clear any existing filters before applying a new one
        .AutoFilterMode = False
        .Range("A1").AutoFilter Field:=1, Criteria1:="PINGRM"
    End With
    
    Application.ScreenUpdating = True
    wb.Save ' Optionally save the workbook
    wb.Close ' Close the workbook

End Sub

Sub ImportAndOrganizePINGRMData()
    Dim filePath As String
    Dim textLine As String
    Dim textData() As String
    Dim i As Long, j As Long
    Dim newWorkbook As Workbook
    Dim ws As Worksheet
    Dim itemNumber As String
    Dim termDate As String
    Dim orderQuantity As String
    Dim reference As String
    Dim yearWeek As String
    Dim lastRow As Long

    ' Prompt user to select CSV file
    filePath = Application.GetOpenFilename("Text Files (*.csv),*.csv", , "Please select the file:")
    If filePath = "False" Then Exit Sub ' If user cancels, exit sub
    
    ' Create a new workbook and setup worksheet
    Set newWorkbook = Application.Workbooks.Add
    Set ws = newWorkbook.Sheets(1)
    ws.Cells(1, 1).Value = "Item Number"
    ws.Cells(1, 2).Value = "Term"
    ws.Cells(1, 3).Value = "Order Quantity"
    ws.Cells(1, 4).Value = "Reference"
    ws.Cells(1, 5).Value = "Year-Week"
    
    ' Open file and read data
    Open filePath For Input As #1
    i = 2 ' Start from the second row for data entries
    Do Until EOF(1)
        Line Input #1, textLine
        textLine = Replace(textLine, ";;", ";") ' Handle double semicolons
        textData = Split(textLine, ";") ' Split data into array
        
        ' Reset variables for new line
        itemNumber = ""
        termDate = ""
        orderQuantity = ""
        reference = ""
        
        ' Parse and assign data based on conditions
        For j = LBound(textData) To UBound(textData)
            Select Case True
                Case Len(textData(j)) = 10 And Left(textData(j), 1) = "4"
                    itemNumber = textData(j)
                Case IsDate(textData(j)) And termDate = ""
                    termDate = textData(j)
                    yearWeek = Format(CDate(termDate), "yyyy") & "-" & Format(CDate(termDate), "ww")
                Case IsNumeric(textData(j)) And orderQuantity = "" And termDate <> ""
                    orderQuantity = textData(j)
                Case (Len(textData(j)) = 8 And Left(textData(j), 1) = "3") Or textData(j) = "forecast"
                    reference = textData(j)
            End Select
        Next j
        
        ' Write parsed data to the worksheet
        ws.Cells(i, 1).Value = itemNumber
        ws.Cells(i, 2).Value = termDate
        ws.Cells(i, 3).Value = orderQuantity
        ws.Cells(i, 4).Value = reference
        ws.Cells(i, 5).Value = yearWeek
        i = i + 1
    Loop
    Close #1 ' Close file after reading
    
    ' Find the last row with data after clean-up
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    If lastRow > 1 Then
    ' There is data present, so we can convert the range to a Table
        Dim table As ListObject
        Set table = ws.ListObjects.Add(SourceType:=xlSrcRange, Source:=ws.Range("A1:E" & lastRow), XlListObjectHasHeaders:=xlYes)
        table.Name = "DataTable"
        table.TableStyle = "TableStyleMedium9"
        
        ' Center everything in the table
        table.Range.HorizontalAlignment = xlCenter
        table.Range.VerticalAlignment = xlCenter
        
        ' Apply filters to the top row
        ws.Range("A1:E1").AutoFilter
    Else
        MsgBox "No data was found to format as a table and apply filters.", vbExclamation
    End If
        
    ' Optimize column widths and activate the workbook
    ws.Columns("A:E").AutoFit
    newWorkbook.Activate
End Sub


Sub ShowMyUserForm()
    UserForm1.Show
End Sub

'Flow of the program - User presses button, which opens a user form - then they press the button to upload there CIMT sheet, then they press another button to upload there RAW PINGRM sheet.
' Then the program cuts the columns its doenst need from the CIMT sheet then runs the Sub ImportAndOrganizeData().
' The CIMT summarised sheet should go into the first sheet and the Summarised PINGRM sheet should go into the second sheet all in a new workbook.

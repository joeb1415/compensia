
Public invoiceMonth As Date
Public invoiceMonthName As String
Public monthSheet As Worksheet
Public monthHeaders As Range
Public letterDate As Date
Public template As Worksheet
    
Public clientName As String
Public clientRow As Integer
Public clientHourlyData As Range

Public compensiaSheet As Worksheet
Public compensiaHeaders As Range
Public compensiaKey As Range
Public clientDataSheet As Worksheet
Public clientDataHeaders As Range
Public clientDataKey As Range



Sub ComputeAndPrintReport()
'
' ComputeAndPrintReport Macro
'
' Author: Joe Bryan
' 2014-10-20
'
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
        
    invoiceMonth = Range("InvoiceMonth").value
    invoiceMonthName = Format(invoiceMonth, "mmmm yyyy")
    Set monthSheet = Worksheets(invoiceMonthName & " backup")
    Set template = Worksheets("Template")
    
    Dim totalColumn As Integer
    clientName = Range("ClientName").Text
    clientRow = monthSheet.Cells(1, 1).EntireColumn.Find(clientName, Lookat:=xlWhole).row
    Set monthHeaders = monthSheet.Cells(2, 1).EntireRow
    totalColumn = monthHeaders.Find("Total", Lookat:=xlPart).Column
    Set clientHourlyData = monthSheet.Range( _
        monthSheet.Cells(clientRow, 2), _
        monthSheet.Cells(clientRow, totalColumn - 1) _
    )
    
    Set compensiaStaffSheet = Worksheets("Compensia Staff")
    Set compensiaStaffHeaders = compensiaStaffSheet.Cells(1, 1).EntireRow
    Set compensiaStaffKey = compensiaStaffSheet.Cells(1, 1).EntireColumn
    Set clientDataSheet = Worksheets("Client Data")
    Set clientDataHeaders = clientDataSheet.Cells(1, 1).EntireRow
    Set clientDataKey = clientDataSheet.Cells(1, 1).EntireColumn
    
    '''''''''''''''''''''''''''''''''''''''
    ' Page 1
    '''''''''''''''''''''''''''''''''''''''
    
    template.Cells.EntireRow.Hidden = False
    template.Cells.EntireColumn.Hidden = False
    template.Columns("A:Z").EntireColumn.Delete
    template.Rows("1:100").EntireRow.Delete
    template.Cells.Font.Name = "Arial"
    template.Cells.Font.Size = 10
    template.Columns("A").ColumnWidth = 68.89
    template.Columns("B").ColumnWidth = 14.14
    
    Dim editCell As Range
    Set editCell = Worksheets("Template").Cells(1, 1)
    
    editCell.value = "Compensia"
    editCell.Font.Name = "Goudy Old Style"
    editCell.Font.Size = 22
    editCell.Font.ColorIndex = 9 ' maroon
    editCell.Characters(Start:=5, Length:=5).Font.ColorIndex = 11 ' blue
    editCell.HorizontalAlignment = xlRight
    Set editCell = Offset(editCell, True) ' Compensia header
    BottomBoarder editCell.Offset(-1, 0)
    
    editCell.value = "125 S Market Street  Suite 1000  San Jose  California  95113  408 876 4025  408 876 4027 fax"
    editCell.Font.Name = "Tw Cen MT"
    editCell.Font.Size = 8
    Set editCell = Offset(editCell, True) ' Compensia address
    
    Set editCell = Offset(editCell, True) ' blank 1
    Set editCell = Offset(editCell, True) ' blank 2
    Set editCell = Offset(editCell, True) ' blank 3
    
    editCell.value = "Via Email"
    SetFontHeader editCell
    editCell.Font.Italic = True
    Set editCell = Offset(editCell, True)
    
    Set editCell = Offset(editCell, True) ' blank 1
    Set editCell = Offset(editCell, True) ' blank 2
    
    Dim letterDateType As String
    letterDateType = Range("LetterDateType").value
    If letterDateType = "Today" Then
        letterDate = Date
    End If
    If letterDateType = "Last of month" Then
        letterDate = DateSerial(Year(invoiceMonth), Month(invoiceMonth) + 1, 0)
    End If
    If letterDateType = "First of next month" Then
        letterDate = DateSerial(Year(invoiceMonth), Month(invoiceMonth) + 1, 1)
    End If
    
    editCell.numberFormat = "@"
    editCell.value = Format(letterDate, "mmmm d, yyyy")
    SetFontHeader editCell
    Set editCell = Offset(editCell, True) ' letter date
    
    Set editCell = Offset(editCell, True) ' blank 1
    Set editCell = Offset(editCell, True) ' blank 2
    
    Dim addresseeFullName As String
    addresseeFullName = ClientVLookup("Addressee Full Name")
    If Len(addresseeFullName) > 0 Then
        editCell.value = addresseeFullName
        SetFontHeader editCell
        Set editCell = Offset(editCell, True) ' Addressee. Only offset if value exists
    End If
    
    Dim addresseeTitle As String
    addresseeTitle = ClientVLookup("Addressee Title")
    If Len(addresseeTitle) > 0 Then
        editCell.value = addresseeTitle
        SetFontHeader editCell
        Set editCell = Offset(editCell, True)
    End If
    
    Dim longClientName As String
    longClientName = ClientVLookup("Long Client Name")
    If Len(longClientName) > 0 Then
        editCell.value = longClientName
        SetFontHeader editCell
        Set editCell = Offset(editCell, True)
    End If
    
    Set editCell = Print4LineAddress(editCell)
    Set editCell = Offset(editCell, True) ' blank 1
    
    Dim addresseeShortName As String
    addresseeShortName = ClientVLookup("Addressee Short Name")
    editCell.value = "Dear " & Trim(addresseeShortName) & ","
    SetFontHeader editCell
    Set editCell = Offset(editCell, True)
    
    Set editCell = Offset(editCell, True) ' blank 1
    
    editCell.value = monthSheet.Cells(clientRow, monthHeaders.Find("Message Body", Lookat:=xlWhole).Column)
    SetFontHeader editCell
    SetMultiLineCellHeight editCell
    Set editCell = Offset(editCell, True)
    
    Set editCell = Offset(editCell, True) ' blank 1
    
    editCell.value = "Please let me know if you have any questions."
    SetFontHeader editCell
    Set editCell = Offset(editCell, True)
    
    Set editCell = Offset(editCell, True) ' blank 1
    
    editCell.value = "Sincerely,"
    SetFontHeader editCell
    Set editCell = Offset(editCell, True)
    
    Set editCell = Offset(editCell, True) ' blank 1
    Set editCell = Offset(editCell, True) ' blank 2
    
    Dim managerName As String
    managerName = ClientVLookup("Manager Name")
    editCell.value = compensiaStaffSheet.Cells(compensiaStaffKey.Find(managerName, Lookat:=xlWhole).row, compensiaStaffHeaders.Find("Signature Name", Lookat:=xlWhole).Column)
    SetFontHeader editCell
    Set editCell = Offset(editCell, True)
    editCell.value = compensiaStaffSheet.Cells(compensiaStaffKey.Find(managerName, Lookat:=xlWhole).row, compensiaStaffHeaders.Find("Title", Lookat:=xlWhole).Column)
    SetFontHeader editCell
    Set editCell = Offset(editCell, True)
    
    Set editCell = Offset(editCell, True) ' blank 1
    Set editCell = Offset(editCell, True) ' blank 2
    
    Dim cc As String
    cc = ClientVLookup("Client cc's")
    If Len(cc) > 0 Then
        cc = Replace(cc, "" & Chr(10) & "", "" & Chr(10) & "      ")
        editCell.value = "cc: " & cc
        SetFontHeader editCell
        SetMultiLineCellHeight editCell
        Set editCell = Offset(editCell, True)
    End If
    
    
    
    '''''''''''''''''''''''''''''''''''''''
    ' Page 2
    '''''''''''''''''''''''''''''''''''''''
    
    Dim page2StartRow As Integer
    page2StartRow = editCell.row
    Range(Worksheets("Template").Cells(1, 1), Worksheets("Template").Cells(2, 2)).Copy
    Worksheets("Template").Paste Destination:=editCell
    Worksheets("Template").ResetAllPageBreaks
    editCell.PageBreak = xlPageBreakManual
    Set editCell = Offset(editCell, True) ' Compensia header
    Set editCell = Offset(editCell, True) ' Compensia address
    
    Set editCell = Offset(editCell, True) ' blank 1
    Set editCell = Offset(editCell, True) ' blank 2
    
    editCell.numberFormat = "@"
    Dim lastOfMonth As Date ' invoice always has last of month, regardless of letterDate
    lastOfMonth = DateSerial(Year(invoiceMonth), Month(invoiceMonth) + 1, 0)
    editCell.value = Format(lastOfMonth, "mmmm d, yyyy")
    SetFontHeader editCell
    Set editCell = Offset(editCell, True)
    
    Set editCell = Offset(editCell, True) ' blank 1
    Set editCell = Offset(editCell, True) ' blank 2
    
    Dim invoiceNumber As String
    invoiceNumber = monthSheet.Cells(clientRow, monthHeaders.Find("Invoice #", Lookat:=xlWhole).Column).value
    If Len(invoiceNumber) > 0 Then
        editCell.value = "INVOICE #" & invoiceNumber ' Invoice line
        SetFontHeader editCell
        editCell.Font.Bold = True
        Set editCell = Offset(editCell, True)
        Set editCell = Offset(editCell, True) ' blank 1
    End If
    
    editCell.value = "Client: " & ClientVLookup("Long Client Name")
    SetFontHeader editCell
    Set editCell = Offset(editCell, True)
    
    Dim includeAddressOnInvoice As String
    includeAddressOnInvoice = ClientVLookup("Include Address On Invoice")
    If includeAddressOnInvoice = "Yes" Then
        Set editCell = Print4LineAddress(editCell)
    End If
    
    Set editCell = Offset(editCell, True) ' blank 1

    editCell.value = "Period: " & invoiceMonthName
    SetFontHeader editCell
    Set editCell = Offset(editCell, True)
    
    Set editCell = Offset(editCell, True) ' blank 1

    Dim poNumber As String
    poNumber = ClientVLookup("Purchase Order")
    If Len(poNumber) > 0 Then
        editCell.value = "PO Number: " & poNumber
        SetFontHeader editCell
        Set editCell = Offset(editCell, True)
        Set editCell = Offset(editCell, True) ' blank 1
    End If
    
    editCell.value = "Payment terms: " & ClientVLookup("Payment Terms")
    SetFontHeader editCell
    Set editCell = Offset(editCell, True)
    
    Set editCell = Offset(editCell, True) ' blank 1
    Set editCell = Offset(editCell, True) ' blank 2
    Set editCell = Offset(editCell, True) ' blank 3
    
    
    Dim employeeName As String
    Dim employeeFLast As String
    Dim employeeShortTitle As String
    Dim hours As Double
    Dim hoursString As String
    Dim invoiceType As String
    
    Dim fees As Double
    Dim totalFees As Double
    Dim travel As Double
    Dim cell As Range
    
    editCell.value = "Professional fees:"
    SetFontHeader editCell
    Set editCell = Offset(editCell, True)
    
    invoiceType = ClientVLookup("Invoice Type")
    totalFees = 0
    Dim firstLineItem As Boolean: firstLineItem = True
    For Each cell In clientHourlyData
        If IsNumeric(cell.Value2) Then
            hours = cell.Value2
        Else
            hours = 0
        End If
        If hours > 0 Then
            employeeName = monthSheet.Cells(2, cell.Column).Text
            
            employeeFLast = compensiaStaffSheet.Cells(compensiaStaffKey.Find(employeeName, Lookat:=xlWhole).row, compensiaStaffHeaders.Find("F. Last", Lookat:=xlWhole).Column)
            employeeShortTitle = compensiaStaffSheet.Cells(compensiaStaffKey.Find(employeeName, Lookat:=xlWhole).row, compensiaStaffHeaders.Find("Short Title", Lookat:=xlWhole).Column)
            billingRate = compensiaStaffSheet.Cells(compensiaStaffKey.Find(employeeName, Lookat:=xlWhole).row, compensiaStaffHeaders.Find("Billing Rate", Lookat:=xlWhole).Column)
            If hours = 1 Then
                hoursString = "1.0 hour"
            Else
                hoursString = Format(hours, "0.00") & " hours"
            End If
            
            fees = hours * billingRate
            totalFees = totalFees + fees
            If invoiceType = "Standard" Then
                If editCell.MergeCells Then
                    editCell.UnMerge
                End If
                
                Dim lineItemName As String
                lineItemName = "'- " & employeeFLast & ", " & employeeShortTitle & ", " & hoursString & " @ $" & billingRate
                
                Dim numberFormat As String: numberFormat = "#,##0.00"
                If firstLineItem Then
                    numberFormat = "$#,##0.00"
                End If
                SetLineItemValue editCell, lineItemName, fees, numberFormat
                
                firstLineItem = False
                Set editCell = Offset(editCell, False)
            End If
        End If
    Next cell
    If invoiceType = "Standard" Then
        BottomBoarder editCell.Offset(-1, 0).Offset(0, 1) ' straight offset(-1, 1) adds column off of a merged cell
    End If
    
    totalFees = monthSheet.Cells(clientRow, totalColumn + 1).Value2 ' using total column rather than computing here
    SetLineItemValue editCell, "          Total professional fees:", totalFees, "$#,##0.00"
    Set editCell = Offset(editCell, False)
    
    editCell.value = "Expenses:"
    SetFontHeader editCell
    Set editCell = Offset(editCell, True)
    
    travel = monthSheet.Cells(clientRow, totalColumn + 5).Value2
    SetLineItemValue editCell, "'- Travel/out-of-pocket expenses", travel, "#,##0.00"
    Set editCell = Offset(editCell, False)
    
    Dim overhead As Double
    overhead = monthSheet.Cells(clientRow, totalColumn + 2).Value2
    SetLineItemValue editCell, "'- Standard administrative/research overhead", overhead, "#,##0.00"
    BottomBoarder editCell.Offset(0, 1)
    Set editCell = Offset(editCell, False)
    
    SetLineItemValue editCell, "          Total expenses:", travel + overhead, "$#,##0.00"
    Set editCell = Offset(editCell, False)
    
    Set editCell = Offset(editCell, True) ' blank 1
    
    SetLineItemValue editCell, "          Grand Total:", totalFees + travel + overhead, "$#,##0.00"
    editCell.Font.Size = 12
    editCell.Font.Bold = True
    editCell.Font.Italic = True
    editCell.Offset(0, 1).Font.Size = 12
    editCell.Offset(0, 1).Font.Italic = True
    Set editCell = Offset(editCell, False)
    
    Do Until editCell.row = page2StartRow + 44
        Set editCell = Offset(editCell, True) ' blank 1
    Loop
    
    editCell.value = "Please remit payment to:"
    SetFontHeader editCell
    editCell.Font.Bold = True
    editCell.Font.Underline = True
    editCell.Font.Italic = True
    Set editCell = Offset(editCell, True)
    
    Set editCell = Offset(editCell, True) ' blank 1
    
    editCell.value = "Compensia, Inc."
    SetFontHeader editCell
    Set editCell = Offset(editCell, True)
    
    editCell.value = "125 S. Market Street"
    SetFontHeader editCell
    Set editCell = Offset(editCell, True)
    
    editCell.value = "Suite 1000"
    SetFontHeader editCell
    Set editCell = Offset(editCell, True)
    
    editCell.value = "San Jose, CA  95113"
    SetFontHeader editCell
    Set editCell = editCell.Offset(0, 1) ' don't merge. Offset right. This is lastCell.
    
    '''''''''''''''''''''''''''''''''''''''''''''
    ' Define printing
    '''''''''''''''''''''''''''''''''''''''''''''
    
    template.PageSetup.printArea = Range(template.Cells(1, 1), editCell).Address
    
    Range(editCell.Offset(1, 0), template.Cells(Rows.Count, 1)).EntireRow.Hidden = True
    Range(editCell.Offset(0, 1), template.Cells(1, Columns.Count)).EntireColumn.Hidden = True
    
    Dim printTo As String
    printTo = Range("PrintTo").Text
    
    PrintPage printTo

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub
Function Print4LineAddress(cell As Range) As Range

    Dim addressLine1 As String
    addressLine1 = ClientVLookup("Address Line 1")
    If Len(addressLine1) > 0 Then
        cell.value = addressLine1
        SetFontHeader cell
        Set cell = Offset(cell, True)
    End If
    
    Dim addressLine2 As String
    addressLine2 = ClientVLookup("Address Line 2")
    If Len(addressLine2) > 0 Then
        cell.value = addressLine2
        SetFontHeader cell
        Set cell = Offset(cell, True)
    End If
    
    Dim addressLine3 As String
    addressLine3 = ClientVLookup("Address Line 3")
    If Len(addressLine3) > 0 Then
        cell.value = addressLine3
        SetFontHeader cell
        Set cell = Offset(cell, True)
    End If
    
    Dim addressLine4 As String
    addressLine4 = ClientVLookup("Address Line 4")
    If Len(addressLine4) > 0 Then
        cell.value = addressLine4
        SetFontHeader cell
        Set cell = Offset(cell, True)
    End If
    
    Set Print4LineAddress = cell
End Function
Function ClientVLookup(columnName As String) As String
    ClientVLookup = clientDataSheet.Cells(clientDataKey.Find(clientName, Lookat:=xlWhole).row, clientDataHeaders.Find(columnName, Lookat:=xlWhole).Column).Text
End Function
Sub BottomBoarder(cell As Range)
    With cell.MergeArea.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 9 ' maroon
    End With
End Sub
Sub SetFontHeader(cell As Range)
    cell.HorizontalAlignment = xlLeft
    cell.Font.Name = "Franklin Gothic Book"
    cell.Font.Size = 11
End Sub
Sub SetLineItemValue(cell As Range, lineItemName As String, value As Double, numberFormat As String)
    With cell
        .value = lineItemName
        .ClearFormats
        .Font.Name = "Arial"
        .Font.Size = 10
        
        With .Offset(0, 1)
            .value = value
            .numberFormat = numberFormat
            .HorizontalAlignment = xlRight
            .Font.Name = "Arial"
            .Font.Size = 10
        End With
    End With
End Sub
Function Offset(cell As Range, mergeFirst As Boolean) As Range
    If mergeFirst And Not cell.MergeCells Then
        Range(cell, cell.Offset(0, 1)).merge
    End If
    If Not IsEmpty(cell) Then
        cell.EntireRow.AutoFit
    End If
    Set Offset = cell.Offset(1, 0)
End Function
Sub SetMultiLineCellHeight(multiLineCell As Range)
    If Not multiLineCell.MergeCells Then
        Range(multiLineCell, multiLineCell.Offset(0, 1)).merge
    End If
        
    Dim cell As Range
    Dim totalWidth As Double
    totalWidth = 0#
    For Each cell In multiLineCell.MergeArea
        totalWidth = totalWidth + cell.ColumnWidth
    Next cell
    
    multiLineCell.WrapText = True
    multiLineCell.Copy
    Dim pasteCell As Range
    Set pasteCell = multiLineCell.Parent.Cells(multiLineCell.row, multiLineCell.Cells(1, 1).Column + 2)
    multiLineCell.Parent.Paste Destination:=pasteCell
    With pasteCell
        .Formula = "=" & multiLineCell.Address
        .Font.Size = multiLineCell.Font.Size
        .ColumnWidth = totalWidth
        .WrapText = True
        .EntireRow.AutoFit
    End With
End Sub
Sub PrintPage(printTo As String)
'
' PrintPage Macro
'
' Author: Joe Bryan
' 2014-10-20
'

    Dim filepath As String
    filepath = ThisWorkbook.Path & "\" & invoiceMonthName
    If Dir(filepath, vbDirectory) = "" Then
        MkDir (filepath)
    End If
    
    Dim managerName As String
    managerName = ClientVLookup("Manager Name")
    filepath = filepath & "\" & managerName
    If Dir(filepath, vbDirectory) = "" Then
        MkDir (filepath)
    End If
    
    Dim filename As String
    filename = clientName & " " & invoiceMonthName & " invoice"
    
    Set monthSheet = Worksheets(invoiceMonthName & " backup")
    clientName = Range("ClientName").Text
    clientRow = monthSheet.Cells(1, 1).EntireColumn.Find(clientName, Lookat:=xlWhole).row
    Set monthHeaders = monthSheet.Cells(2, 1).EntireRow
    invoiceNumber = monthSheet.Cells(clientRow, monthHeaders.Find("Invoice #", Lookat:=xlWhole).Column).value
    If Len(invoiceNumber) > 0 Then
        filename = filename & " " & invoiceNumber
    End If

    Dim printArea As Range
    Set printArea = Range(Worksheets("Template").PageSetup.printArea)
    
    '''''''
    
    If printTo = "PDF" Then
        Worksheets("Template").ExportAsFixedFormat Type:=xlTypePDF, filename:=filepath & "\" & filename & ".pdf"
    End If
    
    '''''''
    
    If printTo = "Word" Then
    
        Dim wordApp
        Set wordApp = CreateObject("Word.Application")
        wordApp.Visible = False
        
        Dim wordDoc
        Set wordDoc = wordApp.Documents.Add
        
        printArea.Copy
        
        wordDoc.Paragraphs(1).Range.PasteExcelTable _
            LinkedToExcel:=False, _
            WordFormatting:=False, _
            RTF:=False
            
        CutCopyMode = False
            
        With wordDoc.PageSetup
            .TopMargin = wordApp.InchesToPoints(0.4) ' space for header
            .BottomMargin = wordApp.InchesToPoints(1)
            .LeftMargin = wordApp.InchesToPoints(1.25)
            .RightMargin = wordApp.InchesToPoints(0.75)
        End With

        Dim table As table
        For Each table In wordDoc.Tables
            table.Select
            With wordApp.Selection.ParagraphFormat
                .SpaceBefore = 0
                .SpaceBeforeAuto = False
                .SpaceAfter = 0
                .SpaceAfterAuto = False
                .LineSpacingRule = wdLineSpaceSingle
                .LineUnitBefore = 0
                .LineUnitAfter = 0
            End With
        Next table
        
        If Dir(filepath & "\" & filename & ".docx") <> "" Then
            Kill filepath & "\" & filename & ".docx"
        End If
        wordDoc.SaveAs (filepath & "\" & filename & ".docx")
        
        wordDoc.Close
        wordApp.Quit
        Set wordDoc = Nothing
        Set wordApp = Nothing
    
    End If
    
    '''''''
    
    If printTo = "Excel" Then
        
        Application.ScreenUpdating = False
        Dim masterWorkbook As Workbook
        Dim ws As Worksheet
        Set masterWorkbook = ThisWorkbook
        Set ws = masterWorkbook.Worksheets("Template")
        
        Dim newWorkbook As Workbook
        Set newWorkbook = Workbooks.Add
        
        Dim newWorksheet As Worksheet
        Set newWorksheet = newWorkbook.Worksheets(1)
        newWorksheet.Cells.Font.Name = "Arial"
        newWorksheet.Cells.Font.Size = 10
        
        ws.Range(ws.PageSetup.printArea).Cells.Copy
        With newWorksheet.Cells(1, 1)
            .PasteSpecial xlPasteValues
            .PasteSpecial xlPasteFormats
            .PasteSpecial xlPasteColumnWidths
        End With
        
        Dim lastCell As Range
        Dim page2Start As Range
        Set lastCell = newWorksheet.Cells.SpecialCells(xlCellTypeLastCell)
        Set page2Start = newWorksheet.Range(Cells(2, 1), lastCell).Find("Compensia", Lookat:=xlWhole)
        newWorksheet.Cells(1, 1).Characters(Start:=5, Length:=5).Font.ColorIndex = 11
        page2Start.Characters(Start:=5, Length:=5).Font.ColorIndex = 11
        newWorksheet.PageSetup.printArea = Range(newWorksheet.Cells(1, 1), lastCell).Address
        page2Start.PageBreak = xlPageBreakManual
        
        Dim messageBodyCell As Range
        Set messageBodyCell = Range(newWorksheet.Cells(1, 1), lastCell).Find("Dear ", Lookat:=xlPart).Offset(2, 0)
        SetMultiLineCellHeight messageBodyCell
        
        Range(lastCell.Offset(1, 0), newWorksheet.Cells(Rows.Count, 1)).EntireRow.Hidden = True
        Range(lastCell.Offset(0, 1), newWorksheet.Cells(1, Columns.Count)).EntireColumn.Hidden = True
        
        With newWorksheet.PageSetup
            .TopMargin = Application.InchesToPoints(0.4)
            .BottomMargin = Application.InchesToPoints(1)
            .LeftMargin = Application.InchesToPoints(1)
            .RightMargin = Application.InchesToPoints(0.75)
        End With
        
        Dim sheetName As String
        sheetName = invoiceMonthName & " invoice"
        newWorksheet.Name = sheetName
        For Each sheet In newWorkbook.Worksheets
            Application.DisplayAlerts = False
            If sheet.Name <> sheetName Then
                sheet.Delete
            End If
        Next
        Application.DisplayAlerts = True
        
        If Dir(filepath & "\" & filename & ".xlsx") <> "" Then
            Kill filepath & "\" & filename & ".xlsx"
        End If
        newWorkbook.SaveAs (filepath & "\" & filename & ".xlsx")
        
        newWorkbook.Close
        Set newWorkbook = Nothing
        
    End If
    
    '''''''
    
    If printTo = "Email Draft" Then
        PrintPage "PDF"
        PrintPage "Email Draft - Existing PDF"
    End If
    
    '''''''
    
    If printTo = "Email Draft - Existing PDF" Then
        Dim emailTo As String
        emailTo = ClientVLookup("Email 1")
        
        Dim emailCc As String
        Dim emailCcList As String
        For i = 2 To 5
            emailCc = ClientVLookup("Email " & i)
            emailCcList = emailCcList & emailCc & ";"
        Next i
        
        Dim bodyText As String
        bodyText = "Dear "
        bodyText = bodyText & ClientVLookup("Addressee Email Name")
        bodyText = bodyText & ", " & vbNewLine & vbNewLine
        bodyText = bodyText & Range("EmailMessageBody").Text
        
        Set olApp = CreateObject("Outlook.Application")
        Set Outmail = olApp.CreateItem(olMailItem)
        With Outmail
            .To = emailTo
            .cc = emailCcList
            .Subject = filename & " from Compensia"
            .body = bodyText
            .Attachments.Add filepath & "\" & filename & ".pdf"
            .Save
        End With
    End If
    
    ' else, do nothing
    
End Sub
Sub SaveAllXlsxAsPdfMaster()
    invoiceMonth = Range("InvoiceMonth").value
    invoiceMonthName = Format(invoiceMonth, "mmmm yyyy")
    
    Dim principalName As String
    principalName = Range("PrincipalName").Text
    
    Dim filepath As String
    filepath = ThisWorkbook.Path & "\" & invoiceMonthName
    If Dir(filepath, vbDirectory) = "" Then
        MkDir (filepath)
    End If
    
    filepath = filepath & "\" & principalName
    If Dir(filepath, vbDirectory) = "" Then
        MkDir (filepath)
    End If
    
    If principalName = "Name" Then
        SaveAllXlsxAsPdf ThisWorkbook.Path & "\" & invoiceMonthName
    Else
        SaveAllXlsxAsPdf ThisWorkbook.Path & "\" & invoiceMonthName & "\" & principalName
    End If
End Sub
Sub SaveAllXlsxAsPdf(filepath As String)
'
' Save all .xlsx files in this folder and subfolders as PDF
'
' Author: Joe Bryan
' 2014-10-26
'
    Dim objFSO As Object
    Dim objFolder As Object
    Dim objSubFolder As Object
    Dim objFile As Object

    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFSO.GetFolder(filepath)
    
    For Each objSubFolder In objFolder.subfolders
        SaveAllXlsxAsPdf objSubFolder.Path ' recursive
    Next objSubFolder
    
    For Each objFile In objFolder.Files
        If Right(objFile.Name, 5) = ".xlsx" And Left(objFile.Name, 2) <> "~$" Then
            SaveXlsxAsPdf objFile.Path
        End If
    Next objFile

End Sub

Sub SaveXlsxAsPdf(xlsxPath As String)
'
' Save an .xlsx file as PDF
'
' Author: Joe Bryan
' 2014-10-27
'
    Dim newApplication As New Application
    newApplication.Visible = False
    
    Dim book As Workbook
    Set book = newApplication.Workbooks.Open(xlsxPath)
    
    Dim pdfPath As String
    pdfPath = Left(xlsxPath, Len(xlsxPath) - 5) & ".pdf"
    If Dir(pdfPath) <> "" Then
        Kill pdfPath
    End If
    book.Worksheets(1).ExportAsFixedFormat Type:=xlTypePDF, filename:=pdfPath
    
    book.Close SaveChanges:=False
    newApplication.Quit
    Set newApplication = Nothing
        
End Sub
Sub ReadAllRevisedBodiesMaster()
    invoiceMonth = Range("InvoiceMonth").value
    invoiceMonthName = Format(invoiceMonth, "mmmm yyyy")
    
    Dim principalName As String
    principalName = Range("PrincipalName").Text
    
    If principalName = "Name" Then
        ReadAllRevisedBodies ThisWorkbook.Path & "\" & invoiceMonthName
    Else
        ReadAllRevisedBodies ThisWorkbook.Path & "\" & invoiceMonthName & "\" & principalName
    End If
End Sub
Sub ReadAllRevisedBodies(filepath As String)
'
' Read message body from revised .xlsx file into master document
'
' Author: Joe Bryan
' 2014-10-28
'
    Dim fso As Object
    Dim rootFolder As Object
    Dim subFolder As Object
    Dim file As Object

    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set Folder = fso.GetFolder(filepath)
    
    For Each subFolder In Folder.subfolders
        ReadAllRevisedBodies subFolder.Path ' recursive
    Next subFolder
    
    For Each file In Folder.Files
        If Right(file.Name, 5) = ".xlsx" Then
            ReadRevisedBody file.Path
        End If
    Next file
    
End Sub
Sub ReadRevisedBody(filepath As String)
'
' Read message body from revised .xlsx file into master document
'
' Author: Joe Bryan
' 2014-10-28
'
    Dim newApplication As New Application
    newApplication.Visible = False
    
    Dim book As Workbook
    Set book = newApplication.Workbooks.Open(filepath)
    
    Dim lastCell As Range
    Dim messageBodyCell As Range
    Dim newBody As String
    Set lastCell = book.Worksheets(1).Cells.SpecialCells(xlCellTypeLastCell)
    Set messageBodyCell = book.Worksheets(1).Range(book.Worksheets(1).Cells(1, 1), lastCell).Find("Dear ", Lookat:=xlPart).Offset(2, 0)
    newBody = messageBodyCell.Text
    
    book.Close SaveChanges:=False
    newApplication.Quit
    Set newApplication = Nothing
    
    invoiceMonth = Range("InvoiceMonth").value
    invoiceMonthName = Format(invoiceMonth, "mmmm yyyy")
    Set monthSheet = Worksheets(invoiceMonthName & " backup")
    
    Dim filename As String
    filename = Mid(filepath, InStrRev(filepath, "\") + 1, 100)
    clientName = Left(filename, InStr(filename, invoiceMonthName) - 2)
    clientRow = monthSheet.Cells(1, 1).EntireColumn.Find(clientName, Lookat:=xlWhole).row
    Set monthHeaders = monthSheet.Cells(2, 1).EntireRow
    
    Set clientDataSheet = Worksheets("Client Data")
    Set clientDataHeaders = clientDataSheet.Cells(1, 1).EntireRow
    Set clientDataKey = clientDataSheet.Cells(1, 1).EntireColumn
    
    With monthSheet.Cells(clientRow, monthHeaders.Find("Message Body", Lookat:=xlWhole).Column)
        .value = newBody
        .WrapText = False
    End With
    With ClientVLookup("Message Body")
        .value = newBody
        .WrapText = False
    End With
    clientDataSheet.Cells(clientDataKey.Find(clientName, Lookat:=xlWhole).row, clientDataHeaders.Find("Message Body Last Updated", Lookat:=xlWhole).Column).value = invoiceMonth
    
End Sub
Sub PrintForPrincipal()
'
' Use specified "Print To" method to print all invoices for specified principal
'
' Author: Joe Bryan
' 2014-10-28
'

    invoiceMonth = Range("InvoiceMonth").value
    invoiceMonthName = Format(invoiceMonth, "mmmm yyyy")
    Set monthSheet = Worksheets(invoiceMonthName & " backup")
    
    Set compensiaStaffSheet = Worksheets("Compensia Staff")
    Set compensiaStaffHeaders = compensiaStaffSheet.Cells(1, 1).EntireRow
    Set compensiaStaffKey = compensiaStaffSheet.Cells(1, 1).EntireColumn
    Set clientDataSheet = Worksheets("Client Data")
    Set clientDataHeaders = clientDataSheet.Cells(1, 1).EntireRow
    Set clientDataKey = clientDataSheet.Cells(1, 1).EntireColumn
    
    Dim principalName As String
    Dim principalinitials As String
    Dim managerNameColumn As Integer
    Dim clientNameCell As Range
    Dim clientDataCell As Range
    principalName = Range("PrincipalName").Text
    principalinitials = compensiaStaffSheet.Cells(compensiaStaffKey.Find(principalName, Lookat:=xlWhole).row, compensiaStaffHeaders.Find("Initials", Lookat:=xlWhole).Column).Text
    managerNameColumn = clientDataHeaders.Find("Manager Name", Lookat:=xlWhole).Column
    
    For row = 3 To monthSheet.Cells.SpecialCells(xlCellTypeLastCell).row
        Set clientNameCell = monthSheet.Cells(row, 1)
        If clientNameCell.Rows.Hidden Then Exit For
        
        clientName = clientNameCell.Text
        If clientName = "TOTAL" Then Exit For
        
        Dim managerInitials As String
        managerInitials = clientNameCell.Offset(0, 1).Text
        If managerInitials <> principalinitials Then
            GoTo nextIteration
        End If
        
        Set clientDataCell = clientDataKey.Find(clientName, Lookat:=xlWhole)
        If clientDataCell Is Nothing Then
            MsgBox clientName & " not found in ""Client Data"" tab. Skipping client"
            GoTo nextIteration
        End If
        
        If clientDataSheet.Cells(clientDataCell.row, managerNameColumn).Text = principalName Or principalName = "Name" Then
            Range("ClientName").value = clientName
            ComputeAndPrintReport
        End If
nextIteration:
    Next row
    
End Sub



Attribute VB_Name = "Module1"
Sub formatingCell()
    'Set s1 = Sheets("Summary Seller")
    's1.Range("B23:B59").NumberFormat = "0"
End Sub

Sub changeColumns()
    Sheets("credit_note_less_21").Visible = True
    Sheets("credit_note_less_21").Select
    Columns("E").ColumnWidth = 21
    Columns("G").ColumnWidth = 30
    Columns("K").ColumnWidth = 13
    Columns("L").ColumnWidth = 6.43
    Columns("L").ColumnWidth = 4.43
    
End Sub

Sub createFolder(path As String)
    Dim MainFolder As String
    MainFolder = ""
    Dim a As Variant ' = Split(path)
    a = Split(path, "\")
    For i = 0 To UBound(a)
        If i = 0 Then
            MainFolder = a(i)
        End If
        If a(i) <> "" And i > 0 Then
            MainFolder = MainFolder & "\" & a(i)
            If Len(Dir(MainFolder, vbDirectory)) = 0 Then
               MkDir (MainFolder)
            End If
        End If
    Next i
End Sub

Sub createFormularForCredit_note(credit_temp As String)
   
    Set s1 = Sheets("credit_note_less_21")
    Set s2 = Sheets(credit_temp)
    Dim col As String
    col = "A5:N18"
    Dim range1 As Range
    Dim range2 As Range
    Set range1 = s1.Range(col)
    Set range2 = s2.Range(col)
    range1.Copy range2
    
End Sub

Sub createSheet()
    Dim heads(15) As String
    heads(0) = "No."
    heads(1) = "Order No."
    heads(2) = "Merchant SKU"
    heads(3) = ""
    heads(4) = ""
    heads(5) = ""
    heads(6) = "Item"
    heads(7) = "Tax Invoice No."
    heads(8) = "Invoice Date"
    heads(9) = "Qty"
    heads(10) = "Commission (Net)"
    heads(11) = "GST*"
    heads(12) = ""
    heads(13) = "Refunded Charge"
    heads(14) = "Date"
    Call deleteSheet
   
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets.Add(After:= _
             ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.name = "temp"
    Sheets("temp").Range("A1:N1").Value = heads
    'IFERROR(INDEX(Seller_CN_index!E:E,MATCH('Summary Seller'!$B$10,Seller_CN_index!H:H,0)),"-")
    Dim strFormula As String
    strFormula = "=IFERROR(INDEX(historic_for_credit_note!F:F,MATCH('Summary Seller'!$B$10,historic_for_credit_note!I:I)),INDEX(historic_for_credit_note!F:F,MATCH('Summary Seller'!$B$76,historic_for_credit_note!R:R)))"
    Range("P2").Formula = strFormula
    Range("O2").Value = "=DATE(2016,MONTH(TODAY()-30),10)"
    
    Set ws = ThisWorkbook.Sheets.Add(After:= _
             ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.name = "create_credit_note"
    Sheets("Finance overview by Item").Select
    Range("D2:BZ2").Select
    Application.CutCopyMode = False
    Selection.Copy
    
    Sheets("create_credit_note").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Sheets("Detailed sales report").Select
    Range("A6").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Call deleteRowsOfCredit(400)
End Sub

Sub deleteRowsOfCredit(t As Integer)
    Application.DisplayAlerts = False
    For i = 1 To t Step 1
        Sheets("credit_note").Rows(21).EntireRow.Delete
    Next
End Sub

Sub CopyCreteToTemp(temp As String, create_note As String)
    Sheets(create_note).Select
    Sheets(temp).Range("A2:N10000").ClearContents
    Dim row As Integer
    row = 2
    Range("C1").Select
    
    Selection.End(xlDown).Select
    'row = ActiveCell.row
    'On Error GoTo notrow
    row = ActiveCell.row
'notrow:
    'Dim tempA As String
    'Call ClearBorder
    Call CopyDataToSheet(temp, create_note, "A", "B2", row)
    Call CopyDataToSheet(temp, create_note, "E", "C2", row)
    Call CopyDataToSheet(temp, create_note, "F", "G2", row)
    Call CopyDataToSheet(temp, create_note, "Q", "N2", row)
    Sheets(temp).calculate
    Call ProcessData(temp)
End Sub

Sub CopyDataToSheet(temp As String, create_note As String, col As String, colDest As String, row As Integer)
    Sheets(create_note).Select
    Range(col & "2:" & col & row).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets(temp).Select
    Range(colDest).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub


Sub ProcessData(name As String)
    Sheets(name).Select
    Dim row As Integer
    row = 1
    Range("C1").Select
    Selection.End(xlDown).Select
    'On Error GoTo notrow
    row = ActiveCell.row
    'notrow:
    Dim DataRange  As Variant
    Dim MyVar      As Double
    Sheets(name).Select
    Range("A2:N" & row).Select
    DataRange = Selection.Value
    Dim i As Integer
    Dim j As Integer
    Dim rate As Double
    Dim sdate As Date
    sdate = Range("O2").Value
    Dim invoice As String
    invoice = "-"
    On Error GoTo notrow
    invoice = Range("P2").Value
notrow:
    rate = Sheets("Seller_CN_index").Range("P2").Value
    j = 1
    i = 1
    For t = 2 To row
        DataRange(j, 1) = i
        DataRange(j, 10) = 1
        MyVar = DataRange(j, 14)
        MyVar = -1 * MyVar
        DataRange(j, 14) = MyVar
        DataRange(j, 11) = MyVar / (1 + rate)
        DataRange(j, 12) = (MyVar / (1 + rate)) * rate
        DataRange(j, 9) = sdate
        DataRange(j, 8) = invoice
        i = i + 1
        j = j + 1
    Next t
    Range("A2:N" & row).Value = DataRange
End Sub

Sub deleteSheet()
    Application.DisplayAlerts = False
    For Each Sheet In Application.Worksheets
        If Sheet.name = "temp" Or Sheet.name = "create_credit_note" Or Sheet.name = "temp-final" Then
            Sheet.Delete
        End If
    Next Sheet
    Application.DisplayAlerts = True
End Sub
Sub hide_sheet()
    Application.DisplayAlerts = False
    For Each Sheet In Application.Worksheets
        If Sheet.name <> "Automatic PDF Generation" Then
            Sheet.Visible = False
        End If
    Next Sheet
    Application.DisplayAlerts = True
End Sub
Sub show_all_sheet()
    For Each Sheet In Application.Worksheets
        Sheet.Visible = True
    Next Sheet
    Sheets("Automatic PDF Generation").Select
End Sub

'Automatic PDF Generation

Sub ClearBorder()
     Sheets("credit_note_less_21").calculate
     Call harcode_credit_note("credit_note_less_21")
     Call deleteRowsOfCredit(100)
End Sub


Sub CopyDataFromTempToCredit()
    Dim rng As Range
    Dim address As String
    Dim lastRow As Integer
    Dim sheetTemp As String
    sheetTemp = "temp"
    Set s1 = Sheets("create_credit_note")
    Set s2 = Sheets(sheetTemp)
    
    lastRow = Sheets("create_credit_note").Cells(Sheets("create_credit_note").Rows.Count, "B").End(xlUp).row
    Set shtCreditNote = Sheets("credit_note")
    ' Clear Text And Border
    Call ClearBorder
     ' draw border for table
    With shtCreditNote.Range("A21:N21").Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = 1
    End With
    
    
    Dim ik As Integer
    Dim iIndex As Integer
    If lastRow > 1 Then
         Dim t As Integer
         t = lastRow + 19
         Call MergeCells(t)
         Dim strRaneSource As String
         Dim strRaneDest As String
         
         strRaneSource = "A2:" & "N" & lastRow
         strRaneDest = "A21:" & "N" & t
         shtCreditNote.Range("G21:G" & (t + 1)).WrapText = True
         shtCreditNote.Range("G21:G" & t).Rows.AutoFit
         shtCreditNote.Range("G21:G" & t).HorizontalAlignment = xlLeft
         Dim rangeSource As Range
         Dim rangeDest As Range
         Set rangeSource = s2.Range(strRaneSource)
         Set rangeDest = shtCreditNote.Range(strRaneDest)
         rangeDest.Value = rangeSource.Value
         lastRow = shtCreditNote.Cells(shtCreditNote.Rows.Count, "B").End(xlUp).row
         
         
         Call drawBorder(lastRow)
    End If
    If lastRow = 1 Then
        lastRow = 21
        Call MergeCells(lastRow)
        Call drawBorder(lastRow)
    End If
    iIndex = lastRow + 1
    With shtCreditNote.Range("A" & iIndex + 2 & ":N" & iIndex + 2).Borders(xlEdgeBottom)
         .LineStyle = xlContinuous
         .Weight = xlThick
         .ColorIndex = 1
    End With
    
    Dim strString As String
    strString = "*GST rate: Singapore (7%); Malaysia (6%); Hong Kong (zero-rated); Taiwan (zero-rated)"
         
    ik = iIndex + 4
    shtCreditNote.Cells(ik, 3).Value = strString
         
    strString = "This is a computer generated credit note. No signature is required."
    ik = iIndex + 6
   
    shtCreditNote.Cells(ik, 3).Value = strString
         
    With shtCreditNote.Range("A" & iIndex + 9 & ":N" & iIndex + 9).Borders(xlEdgeBottom)
       .LineStyle = xlContinuous
       .Weight = xlThick
       .ColorIndex = 1
    End With
    row_note_sheet = ik
  
    
End Sub

Sub drawBorder(lastRow As Integer)
        Set shtCreditNote = Sheets("credit_note")
        Set rng = shtCreditNote.Range("A21:N" & (lastRow + 2))
        rng.Font.name = "Arial"
        rng.Font.FontStyle = "Regular"
        rng.Font.Size = 8
        rng.Font.Bold = False
        
        With shtCreditNote.Range("A21:N21").Borders(xlEdgeTop)
             .LineStyle = xlContinuous
             .Weight = xlThin
             .ColorIndex = 1
         End With
         
         
         
         With shtCreditNote.Range("B21:B" & lastRow).Borders(xlEdgeLeft)
             .LineStyle = xlContinuous
             .Weight = xlThin
             .ColorIndex = 1
         End With
         
         With shtCreditNote.Range("C21:C" & lastRow).Borders(xlEdgeLeft)
             .LineStyle = xlContinuous
             .Weight = xlThin
             .ColorIndex = 1
         End With
        
         
         With shtCreditNote.Range("G21:G" & lastRow).Borders(xlEdgeLeft)
             .LineStyle = xlContinuous
             .Weight = xlThin
             .ColorIndex = 1
         End With
         
         With shtCreditNote.Range("H21:H" & lastRow).Borders(xlEdgeLeft)
             .LineStyle = xlContinuous
             .Weight = xlThin
             .ColorIndex = 1
         End With
         
         With shtCreditNote.Range("I21:I" & lastRow).Borders(xlEdgeLeft)
             .LineStyle = xlContinuous
             .Weight = xlThin
             .ColorIndex = 1
         End With
         
         With shtCreditNote.Range("J21:J" & lastRow).Borders(xlEdgeLeft)
             .LineStyle = xlContinuous
             .Weight = xlThin
             .ColorIndex = 1
         End With
         
         With shtCreditNote.Range("K21:K" & lastRow).Borders(xlEdgeLeft)
             .LineStyle = xlContinuous
             .Weight = xlThin
             .ColorIndex = 1
         End With
         
         With shtCreditNote.Range("L21:L" & lastRow).Borders(xlEdgeLeft)
             .LineStyle = xlContinuous
             .Weight = xlThin
             .ColorIndex = 1
         End With
                          
         With shtCreditNote.Range("A" & lastRow & ":N" & lastRow).Borders(xlEdgeBottom)
             .LineStyle = xlContinuous
             .Weight = xlThin
             .ColorIndex = 1
         End With
         
         
         
         
          
         
        
         
         'Add Formula for sum
         Dim line_a As Integer
         line_a = lastRow
         iIndex = lastRow + 1
         If flag_tw = True Then
            With shtCreditNote.Range("A" & lastRow + 1 & ":N" & lastRow + 1).Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .ColorIndex = 1
            End With
            'shtCreditNote.Range(row_hide).EntireRow.Hidden = True
            Sheets("credit_note").Range("A" & iIndex & ":F" & iIndex).MergeCells = True
            shtCreditNote.Cells(iIndex, 1).Value = "Tax witheld and paid by customer(WHT)"
            lAddress = shtCreditNote.Cells(iIndex, 14).address(RowAbsolute:=False, ColumnAbsolute:=False)
            shtCreditNote.Range(lAddress).Formula = "=-20%*Sum(N21:N" & lastRow & ")"
            shtCreditNote.Range(lAddress).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
            shtCreditNote.Cells(iIndex, 14).Font.Bold = True
            iIndex = lastRow + 2
            line_a = line_a + 1
         End If
         
        
         With shtCreditNote.Range("A21:A" & line_a).Borders(xlEdgeLeft)
             .LineStyle = xlContinuous
             .Weight = xlThin
             .ColorIndex = 1
         End With
         
          With shtCreditNote.Range("N19:N" & iIndex).Borders(xlEdgeLeft)
             .LineStyle = xlContinuous
             .Weight = xlThin
             .ColorIndex = 1
         End With
         
         With shtCreditNote.Range("N19:N" & iIndex).Borders(xlEdgeRight)
             .LineStyle = xlContinuous
             .Weight = xlThin
             .ColorIndex = 1
         End With
         
         With shtCreditNote.Range("I" & iIndex & ":I" & iIndex).Borders(xlEdgeLeft)
             .LineStyle = xlContinuous
             .Weight = xlThin
             .ColorIndex = 1
         End With
         
         With shtCreditNote.Range("J" & iIndex & ":J" & iIndex).Borders(xlEdgeLeft)
             .LineStyle = xlContinuous
             .Weight = xlThin
             .ColorIndex = 1
         End With
         With shtCreditNote.Range("K" & iIndex & ":K" & iIndex).Borders(xlEdgeLeft)
             .LineStyle = xlContinuous
             .Weight = xlThin
             .ColorIndex = 1
         End With
         With shtCreditNote.Range("L" & iIndex & ":L" & iIndex).Borders(xlEdgeLeft)
             .LineStyle = xlContinuous
             .Weight = xlThin
             .ColorIndex = 1
         End With
         
         With shtCreditNote.Range("I" & iIndex & ":N" & iIndex).Borders(xlEdgeBottom)
             .LineStyle = xlContinuous
             .Weight = xlThin
             .ColorIndex = 1
         End With
         
         Sheets("credit_note").Range("L" & iIndex & ":M" & iIndex).MergeCells = True
         
         lAddress = shtCreditNote.Cells(iIndex, 10).address(RowAbsolute:=False, ColumnAbsolute:=False)
         shtCreditNote.Range(lAddress).Formula = "=Sum(J21:J" & lastRow & ")"
         
         lAddress = shtCreditNote.Cells(iIndex, 11).address(RowAbsolute:=False, ColumnAbsolute:=False)
         shtCreditNote.Range(lAddress).Formula = "=Sum(K21:K" & lastRow & ")"
         
         lAddress = shtCreditNote.Cells(iIndex, 12).address(RowAbsolute:=False, ColumnAbsolute:=False)
         shtCreditNote.Range(lAddress).Formula = "=Sum(L21:L" & lastRow & ")"
         
         lAddress = shtCreditNote.Cells(iIndex, 14).address(RowAbsolute:=False, ColumnAbsolute:=False)
         shtCreditNote.Range(lAddress).Formula = "=Sum(N21:N" & (iIndex - 1) & ")"
         
         
         shtCreditNote.Cells(iIndex, 9).Value = "Total"
         shtCreditNote.Cells(iIndex, 9).Font.Bold = True
         shtCreditNote.Cells(iIndex, 10).Font.Bold = True
         shtCreditNote.Cells(iIndex, 11).Font.Bold = True
         shtCreditNote.Cells(iIndex, 12).Font.Bold = True
         shtCreditNote.Cells(iIndex, 14).Font.Bold = True
         shtCreditNote.Range("K21:N" & iIndex).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
         
End Sub

Sub MergeCells(r As Integer)
    Set shtCreditNote = Sheets("credit_note")
    Dim i As Integer
    Dim t As Integer
    t = r + 1
    i = 21
    For k = 21 To t
        Sheets("credit_note").Range("A" & i & ":C" & i).MergeCells = False
        Sheets("credit_note").Range("C" & i & ":F" & i).MergeCells = True
        Sheets("credit_note").Range("L" & i & ":M" & i).MergeCells = True
        i = i + 1
    Next k
    'Sheets("credit_note").Range("L" & i & ":M" & i).MergeCells = True
    'i = i + 1
    For k = 1 To 100
        Sheets("credit_note").Range("C" & i & ":F" & i).MergeCells = False
        Sheets("credit_note").Range("L" & i & ":M" & i).MergeCells = False
        i = i + 1
    Next k
End Sub



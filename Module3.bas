Attribute VB_Name = "Module3"
Sub RunSelectedSellers()
    Application.ScreenUpdating = False
    Sheets("Automatic PDF Generation").Select
    Sheets("Automatic PDF Generation").Range("I1:I10000").ClearContents
    
    Dim R7 As Integer
    On Error GoTo error
    Range("F42").Select
    Selection.End(xlDown).Select
    R7 = ActiveCell.row
    Range("F42:F" & R7).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("I1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    
    Dim seller_name As String
    seller_name = Sheets("seller_CN_index").Cells(9, 7).Value
    Call show_all_sheet
    Call createSheet
    Dim i As Integer
    i = 2
    Call isCheckTw
    Call formatingCell
    Call showHiderowsfortw
    Do While Sheets("Automatic PDF Generation").Cells(i, 9).Value <> ""
        seller_name = Sheets("Automatic PDF Generation").Cells(i, 9).Value
        Call ExportBySeller(seller_name)
        i = i + 1
    Loop
error:
    Sheets("Automatic PDF Generation").Range("I1:I10000").ClearContents
    Call deleteSheet
    Call hide_sheet
    Sheets("Automatic PDF Generation").Select
End Sub

Sub ExportBySeller(seller_name As String)
    Dim excelFile As String
    excelFile = Sheets("Automatic PDF Generation").Range("C2").Value & Sheets("Seller_CN_index").Range("K4").Value & Sheets("Automatic PDF Generation").Range("C3").Value & " closing\Tools & Reports\Output\Excel Files\"
    Call createFolder(excelFile)
    Dim summaryFile As String
    summaryFile = Sheets("Automatic PDF Generation").Range("C2").Value & Sheets("Seller_CN_index").Range("K4").Value & Sheets("Automatic PDF Generation").Range("C3").Value & " closing\Tools & Reports\Output\Seller Reports\"
    Call createFolder(summaryFile)
    Dim creditFile As String
    creditFile = Sheets("Automatic PDF Generation").Range("C2").Value & Sheets("Seller_CN_index").Range("K4").Value & Sheets("Automatic PDF Generation").Range("C3").Value & " closing\Tools & Reports\Output\Credit Notes\"
    Call createFolder(creditFile)
    Dim invoiceFile As String
    invoiceFile = Sheets("Automatic PDF Generation").Range("C2").Value & Sheets("Seller_CN_index").Range("K4").Value & Sheets("Automatic PDF Generation").Range("C3").Value & " closing\Tools & Reports\Output\Tax Invoices\"
    Call createFolder(invoiceFile)
    Dim R7 As Integer
    Call Export_Seller(seller_name)
    Sheets("Detailed sales report").Select
    Range("A6").Select
    Selection.End(xlDown).Select
    R7 = 6
    On Error GoTo error
    R7 = ActiveCell.row
    If R7 > 6 Then
        Call savexls_sheet(seller_name, excelFile)
        Call show_hiderows
        Sheets("Summary Seller").calculate
        Call savepdfsumm_sheet(seller_name, summaryFile)
        Sheets("Detailed sales report").Columns("A:AZ").EntireColumn.Hidden = False
        Call Export_Invoice(seller_name, invoiceFile)
        Call Export_Credit_Note(seller_name, creditFile)
    End If
error:
End Sub
Sub Export_Invoice(seller_name As String, path As String)
      Call filterdetail_invoice_sheet(seller_name)
        With Sheets("Summary Seller")
        .Range("B10").Value = " "
        .Range("B10").Value = seller_name
        .calculate
        End With
        Sheets("Detailed sales report").calculate
        Sheets("Summary Seller").calculate
        Sheets("Tax Invoice").calculate
        Dim R7 As Integer
        Sheets("Detailed sales report").Select
        Range("A6").Select
        Selection.End(xlDown).Select
        R7 = 6
        On Error GoTo error
        R7 = ActiveCell.row
        If R7 > 6 Then
            Call savepdfinv_sheet(seller_name, path)
        End If
error:
End Sub
Sub Export_Credit_Note(seller_name As String, path As String)
    With Sheets("Summary Seller")
        .Range("B10").Value = " "
        .Range("B10").Value = seller_name
        .calculate
        End With
        Sheets("Detailed sales report").calculate
        Sheets("Summary Seller").calculate
        Sheets("credit_note_less_21").calculate
        Call filterdetail_credit_note_sheet(seller_name)
        Dim R7 As Integer
        R7 = 1
        Sheets("create_credit_note").Select
        Range("A1").Select
        Selection.End(xlDown).Select
        On Error GoTo error
        R7 = ActiveCell.row
        If R7 > 1 Then
            Call save_pdf_credit_note_sheet(seller_name, path)
        End If
error:
End Sub


'create data Validation by vba
Sub createDataValidation()
    Dim ws As Worksheet
    Dim ws1 As Worksheet
    Dim range1 As Range, rng As Range
    Set ws = Sheets("Seller_CN_index")
    Set ws1 = Sheets("Automatic PDF Generation")
    Dim lastRow As Integer
    lastRow = ws.Cells(ws.Rows.Count, "G").End(xlUp).row
    Set range1 = ws.Range("G2:G" & (lastRow + 1))
    Dim t As Integer
    Dim i As Integer
    i = 43
    t = 60
    ws1.Range("F43:F60").ClearContents
    For i = 43 To t
        Set rng = ws1.Range("F" & i)
        With rng.Validation
            .Delete 'delete previous validation
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
            Formula1:="='" & ws.name & "'!" & range1.address
        End With
    Next
    
End Sub

Sub ProFinaceOverview()
   ' On Error GoTo error
    
    Set ws = Sheets("Finance overview by seller_")
    Sheets("Finance overview by seller_").Select
    Dim R7 As Integer
    Range("A2").Select
    Selection.End(xlDown).Select
    R7 = ActiveCell.row
    Dim last As Integer
    last = R7 + 1
    
    Set rng = ws.Range("A2:A" & R7)
    rng.Font.Bold = True
    Set rng = ws.Range("B" & R7 & ":AB" & last)
    rng.Font.Bold = True
    Set rng = ws.Range("B2:AB2")
    rng.Font.Bold = True
    
    Set rng = ws.Range("A2:Z" & R7)
    rng.WrapText = True
    rng.VerticalAlignment = xlTop
    
    With ws.Range("B" & R7 & ":AB" & R7).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With ws.Range("B" & last & ":AB" & last).Borders(xlEdgeBottom)
        .LineStyle = xlDouble
        .Weight = xlThick
    End With
    Dim shtSource As Worksheet
    Set shtSource = Sheets("Finance overview by seller_")
    shtSource.Range("A" & last).Value = "Grand Total"
    shtSource.Range("A" & last).Font.Bold = True
    Dim col As String
    Dim strFormula As String
    Dim i As Integer
    For i = 2 To 28
        col = fnColumnToLetter_Split(1, i, shtSource)
        strFormula = "=sum(" & col & "3:" & col & R7 & ")"
        shtSource.Range(col & last).Formula = strFormula
    Next
    Worksheets("Finance overview by seller_").Columns("AA:AB").AutoFit
    Sheets("Finance overview by seller_").Select
    Columns("AC").Select
    Selection.ColumnWidth = 28
    Columns("A").Select
    Selection.ColumnWidth = 45
    Columns("B:Z").Select
    Selection.ColumnWidth = 15
    On Error GoTo error
    ws.Rows("1:65536").Rows.Ungroup
    Selection.Columns.Ungroup
error:
    Columns("Z:AC").Select
    ' un group
    On Error GoTo error1
        Selection.Columns.Ungroup
error1:
    Selection.Columns.Group
    Sheets("Finance overview by seller_").calculate
End Sub

Function fnColumnToLetter_Split(ByVal postion As Integer, ByVal intColumnNumber As Integer, shtSource As Worksheet)
    fnColumnToLetter_Split = Split(shtSource.Cells(postion, intColumnNumber).address, "$")(1)
End Function


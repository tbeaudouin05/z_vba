Attribute VB_Name = "i_excel_report_"
Sub ExcelReport()

    Application.ScreenUpdating = False
    Dim FName As String, ws As Worksheet
    Dim i As Integer, k As Integer, wb As Workbook
    Dim MainFolder As String
    Dim CN As String
    
    Call show_all

    MainFolder = Sheets("Automatic PDF Generation").Range("C2").Value & Sheets("Seller_CN_index").Range("K4").Value & Sheets("Automatic PDF Generation").Range("C3").Value & " closing\Tools & Reports\Output\Excel Files\"
    
    i = 2

    If Len(Dir(MainFolder, vbDirectory)) = 0 Then
            MkDir (MainFolder)
    End If
    
    
    Sheets("Finance overview by Item").Cells.AutoFilter

    Do While Sheets("seller_CN_index").Cells(i, 1).Value <> ""
        With Sheets("Summary Seller")
        .Range("B10").Value = ""
        .Range("B10").Value = Sheets("seller_CN_index").Cells(i, 7).Value
        .calculate
        End With
        
        Call filterdetail_credit_note_excel(i)
        
        If Sheets("Summary Seller").Cells(19, 5).Value <> "-" Then
        With Sheets("Finance overview by Item")
        k = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
        End With
        Else: k = 1
        End If
        
        If k <= 21 Then
        CN = "credit_note_less_21"
        
        ElseIf k > 21 And k <= 68 Then
        CN = "credit_note_less_68"
        
        ElseIf k > 68 And k <= 115 Then
        CN = "credit_note_less_115"
        
        ElseIf k > 115 And k <= 162 Then
        CN = "credit_note_less_162"
        
        ElseIf k > 162 And k <= 200 Then
        CN = "credit_note_less_200"
        
        ElseIf k > 200 And k <= 250 Then
        CN = "credit_note_less_250"
        
        ElseIf k > 250 Then
        CN = "credit_note_less_300"
        
        End If
        
        Sheets("Detailed sales report").calculate
        Sheets("Summary Seller").calculate
        Sheets(CN).calculate
        
        Call harcode_credit_note(CN)
        
        Call filterdetail_invoice_excel(i)
        Sheets("Detailed sales report").calculate
        Sheets("Summary Seller").calculate
        Sheets("Tax Invoice").calculate
        
        Call hardcode_invoice
        
        Call filterdetail(i)
        Sheets("Detailed sales report").calculate
        Sheets("Summary Seller").calculate
        
        Call savexls(i, k, MainFolder)
        
        i = i + 1
        
    Loop
    
    Call hide_all
    
    Sheets("Automatic PDF Generation").Select

End Sub

Sub savexls(ix As Integer, k As Integer, path As String)

    Dim brand
    
    brand = legalized(Sheets("seller_CN_index").Cells(ix, 7).Value)
    
    FName = path & brand & " - Seller Report" & _
        " " & Sheets("seller_CN_index").Cells(2, 10).Value & ".xlsx"
        
        ' erase the columns used as flags (don t want them to appear for sellers)
        Sheets("Detailed sales report").Columns("AO:AZ").ClearContents
        
        Sheets("Finance overview by Item").Cells.AutoFilter
    
    Set wb = Workbooks.Add(xlWBATWorksheet)
    With wb
        ThisWorkbook.Sheets(Array("Summary Seller", "Detailed sales report", "Tax Invoice_", "credit_note")).Copy After:=.Worksheets(.Worksheets.Count)
        With ActiveSheet.UsedRange
            .Value = .Value
        End With
        Application.DisplayAlerts = False
        .Worksheets(1).Delete
        Application.DisplayAlerts = True
        .SaveAs filename:=FName
        .Close False
    End With
    
    
End Sub


Sub harcode_credit_note(CN As String)

    Sheets(CN).Select
    Cells.Select
    Selection.Copy
    Sheets("credit_note").Select
    Range("A1").Select
    ActiveSheet.Paste
    Cells.Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub


Sub hardcode_invoice()

    Sheets("Tax Invoice").Select
    Cells.Select
    Selection.Copy
    Sheets("Tax Invoice_").Select
    Range("A1").Select
    ActiveSheet.Paste
    Cells.Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

End Sub

Sub ExcelReport_seller()

    Application.ScreenUpdating = False
    Dim FName As String, ws As Worksheet
    Dim i As Integer, j As Integer, k As Integer, wb As Workbook
    Dim MainFolder As String
    Dim CN As String
    
    Call show_all

    MainFolder = Sheets("Automatic PDF Generation").Range("C2").Value & Sheets("Seller_CN_index").Range("K4").Value & Sheets("Automatic PDF Generation").Range("C3").Value & " closing\Tools & Reports\Output\Excel Files\"
    
    i = 2
    j = 33

    If Len(Dir(MainFolder, vbDirectory)) = 0 Then
            MkDir (MainFolder)
    End If
    
    
    Sheets("Finance overview by Item").Cells.AutoFilter

    Do While Sheets("seller_CN_index").Cells(i, 1).Value <> ""
    
    If Sheets("seller_CN_index").Cells(i, 7).Value = Sheets("Old macro Thomas").Cells(j, 6).Value Then
    
        With Sheets("Summary Seller")
        .Range("B10").Value = ""
        .Range("B10").Value = Sheets("seller_CN_index").Cells(i, 7).Value
        .calculate
        End With
        
        Call filterdetail_credit_note_excel(i)
        
        If Sheets("Summary Seller").Cells(19, 5).Value <> "-" Then
        With Sheets("Finance overview by Item")
        k = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
        End With
        Else: k = 1
        End If
        
        If k <= 21 Then
        CN = "credit_note_less_21"
        
        ElseIf k > 21 And k <= 68 Then
        CN = "credit_note_less_68"
        
        ElseIf k > 68 And k <= 115 Then
        CN = "credit_note_less_115"
        
        ElseIf k > 115 And k <= 162 Then
        CN = "credit_note_less_162"
        
        ElseIf k > 162 And k <= 200 Then
        CN = "credit_note_less_200"
        
        ElseIf k > 200 And k <= 250 Then
        CN = "credit_note_less_250"
        
        ElseIf k > 250 Then
        CN = "credit_note_less_300"
        
        End If
        
        Sheets("Detailed sales report").calculate
        Sheets("Summary Seller").calculate
        Sheets(CN).calculate
        
        Call harcode_credit_note(CN)
        
        Call filterdetail_invoice_excel(i)
        Sheets("Detailed sales report").calculate
        Sheets("Summary Seller").calculate
        Sheets("Tax Invoice").calculate
        
        Call hardcode_invoice
        
        Call filterdetail(i)
        Sheets("Detailed sales report").calculate
        Sheets("Summary Seller").calculate
        
        Call savexls(i, k, MainFolder)
        
        j = j + 1
        
        End If
        
        i = i + 1
        
    Loop
    
    Call hide_all
    
    Sheets("Automatic PDF Generation").Select

End Sub

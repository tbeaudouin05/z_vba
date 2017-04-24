Attribute VB_Name = "i_finance_overview"
Sub hardcode_finance_overview()
    Call deleteSheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets.Add(After:= _
             ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    Dim temp As String
    temp = "temp-final"
    ws.name = temp
    
    Sheets("Finance overview by seller").Select
    Cells.Select
    Selection.Copy
    Sheets(temp).Select
    Range("A1").Select
    ActiveSheet.Paste
    Cells.Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    Call copyFinance(temp)
    
    Call deleteSheet
    
    End Sub
        
    Sub copyFinance(temp As String)
        Dim R4 As Integer
        R4 = 2
        Set ws = Sheets("Finance overview by Item")
        country = ws.Range("B3").Value
        Sheets(temp).Select
        Sheets(temp).Cells.AutoFilter
        Sheets(temp).Select
        Range("A2").Select
        Selection.End(xlDown).Select
        R4 = ActiveCell.row
        Sheets(temp).Select
        Range("$A$2:$AD$2").AutoFilter
        ActiveSheet.Range("A2:AD" & R4).AutoFilter Field:=1, Criteria1:=country
        Range("B2").Select
        Selection.End(xlDown).Select
        R4 = ActiveCell.row
                
        Range("B1:AD" & R4).Select
        Selection.Copy
       
        Sheets("Finance overview by seller_").Select
        Range("A1").Select
        ActiveSheet.Paste
        Cells.Select
        Application.CutCopyMode = False
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
               
    End Sub
    
    
    Sub hardcode_index()
    
    Sheets("Seller_CN_index").Select
    Cells.Select
    Selection.Copy
    Sheets("Seller_CN_index_").Select
    Range("A1").Select
    ActiveSheet.Paste
    Cells.Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    End Sub
    
    Sub overview()
    
    Application.ScreenUpdating = False
    Dim FName As String, ws As Worksheet
    Dim wb As Workbook
    Dim MainFolder As String
    
    Call show_all
    
    MainFolder = Sheets("Automatic PDF Generation").Range("C2").Value & Sheets("Seller_CN_index").Range("K4").Value & Sheets("Automatic PDF Generation").Range("C3").Value & " closing\Tools & Reports\Output\"
    
    Call createFolder(MainFolder)
    If Len(Dir(MainFolder, vbDirectory)) = 0 Then
            MkDir (MainFolder)
    End If
    
    
    Call hardcode_index
    
    Call hardcode_finance_overview
    
    Call save_overview(MainFolder)
    
    Sheets("Finance overview by seller_").Select
    Cells.Select
    Selection.Delete
    
    Sheets("Seller_CN_index_").Select
    Cells.Select
    Selection.Delete
    
    Call hide_all
    
    Application.ScreenUpdating = True
    
    End Sub
    
    
    Sub save_overview(path As String)
    
    Call ProFinaceOverview
    FName = path & "Finance Overview" & _
        " - " & Sheets("Seller_CN_index").Range("K3").Value & " - " & Sheets("seller_CN_index").Cells(2, 10).Value & ".xlsx"
    
    Set wb = Workbooks.Add(xlWBATWorksheet)
    With wb
        ThisWorkbook.Sheets(Array("Finance overview by seller_", "Finance overview by Item", "Seller_CN_index_")).Copy After:=.Worksheets(.Worksheets.Count)
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


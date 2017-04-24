Attribute VB_Name = "b_seller_CN_index_overviews"
Sub seller_CN_index_and_overviews()

Dim R2, i As Integer

 ' adjust number of lines for finance overviews
    
    
    ' erase lines finance overview by items
    
    Sheets("Finance overview by Item").Select
    Sheets("Finance overview by Item").Cells.AutoFilter
    
    Sheets("Finance overview by Item").Select
    Range("$A$2:$BZ$2").AutoFilter
    
        Range("C2").Select
        Selection.End(xlDown).Select
        R2 = ActiveCell.row
    
    Range("A2:BZ" & R2).Select
    Selection.ClearContents
    
    ' update finance overview by items
    
    Sheets("Orders data for macro & pivot").Select

        Range("D1").Select
        Selection.End(xlDown).Select
        R2 = ActiveCell.row

    Range("B1:BZ" & R2).Select
    Selection.Copy
    
    Sheets("Finance overview by Item").Select
    Range("A2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
        Range("C2").Select
        Selection.End(xlDown).Select
        R2 = ActiveCell.row
    
    ' erase (blank)s
    
    Cells.Select
    Selection.Replace What:="(blank)", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    ' sort finance overview by items on seller names (for seller_CN_index)
    
    Columns("C:C").Select
    ActiveWorkbook.Worksheets("Finance overview by Item").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Finance overview by Item").Sort.SortFields.Add Key:=Range( _
        "C2"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Finance overview by Item").Sort
        .SetRange Range("A2:BZ" & R2)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

' Update seller_CN_index

    'erase sellers last months
    
    Sheets("seller_CN_index").Select
    Sheets("seller_CN_index").Cells.AutoFilter
    
        Range("C2").Select
        Selection.End(xlDown).Select
        R2 = ActiveCell.row
    
    Range("G1:I" & R2).Select
    Selection.ClearContents
    
    'seller_name_summary
    
    Sheets("Finance overview by Item").Select
    
        Range("C2").Select
        Selection.End(xlDown).Select
        R2 = ActiveCell.row
    
    Range("C2:C" & R2).Select
    Selection.Copy
    
    Sheets("seller_CN_index").Select
    Range("G1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    ActiveSheet.Range("$G:$G").RemoveDuplicates Columns:=1, Header:= _
        xlYes
        
    Range("G1").Value = "seller_name_summary"
    
    ' seller_name_CN
    
    Sheets("Finance overview by Item").Select
    
    Range("$A$2:$BZ$2").AutoFilter
    ActiveSheet.Range("A2:BZ" & R2).AutoFilter Field:=51, Criteria1:="1"
        
    Range("C2:C" & R2).Select
    Selection.Copy
    
    Sheets("seller_CN_index").Select
    Range("H1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    ActiveSheet.Range("$H:$H").RemoveDuplicates Columns:=1, Header:= _
        xlYes
        
    Range("H1").Value = "seller_name_CN"
    
    Range("I1").Value = "seller_name_invoice"
    
    Sheets("Finance overview by Item").Select
    Range("A2:A" & R2).Select
    Selection.Copy
    
    Sheets("seller_CN_index").Select
    Range("Q1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    ActiveSheet.Range("$Q:$Q").RemoveDuplicates Columns:=1, Header:= _
        xlYes
    
    Range("Q1").Value = "short_code_seller"
        
    'seller_name_invoice
    
    Sheets("Finance overview by Item").Cells.AutoFilter
    Sheets("Finance overview by Item").Select
    
    Range("$A$2:$BZ$2").AutoFilter
    ActiveSheet.Range("A2:BZ" & R2).AutoFilter Field:=52, Criteria1:="1"
    
    Range("C2:C" & R2).Select
    Selection.Copy
    
    Sheets("seller_CN_index").Select
    Range("I1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    ActiveSheet.Range("$I:$I").RemoveDuplicates Columns:=1, Header:= _
        xlYes
    
    Range("I1").Value = "seller_name_invoice"
    
    Sheets("Finance overview by Item").Select
    Range("A2:A" & R2).Select
    Selection.Copy
    
    Sheets("seller_CN_index").Select
    Range("R1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    ActiveSheet.Range("$R:$R").RemoveDuplicates Columns:=1, Header:= _
        xlYes
    
    Range("R1").Value = "short_code_invoice"
        
    ' adjust number of lines seller_CN_index
    
        Range("C2").Select
        Selection.End(xlDown).Select
        R2 = ActiveCell.row
     
    Range("A4:F4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("A3:F3").Select
    Selection.AutoFill Destination:=Range("A3:F" & R2), Type:=xlFillDefault
    
    Sheets("seller_CN_index").calculate
    
    ' erase the (blank)s in Finance overview by Item
    
    Sheets("Finance overview by Item").Select
    Cells.Select
    Selection.Replace What:="(blank)", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
    Sheets("Finance overview by Item").Cells.AutoFilter
    
    Call adjust_lines_check
    
     
End Sub

Sub adjust_lines_check()

Dim R10 As Integer

Sheets("Finance overview by seller").Select

    Range("A3").Select
    Selection.End(xlDown).Select
    R10 = ActiveCell.row

    Range("AD4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("AD3").Select
    Selection.AutoFill Destination:=Range("AD3:AD" & R10)
    ActiveSheet.calculate
End Sub

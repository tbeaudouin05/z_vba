Attribute VB_Name = "h_filter_details_"
Sub filterdetail(ix As Integer)

Dim R3 As Integer


    Sheets("Detailed sales report").Range("A7:BB10000").ClearContents
    
    Sheets("Finance overview by Item").Cells.AutoFilter
    
        Sheets("Finance overview by Item").Select
        Range("C2").Select
        Selection.End(xlDown).Select
        R3 = ActiveCell.row
    
    Sheets("Finance overview by Item").Select
    Range("$A$2:$BZ$2").AutoFilter
    ActiveSheet.Range("A2:BZ" & R3).AutoFilter Field:=3, Criteria1:=Sheets("seller_CN_index").Cells(ix, 7).Value
        
        Range("C2").Select
        Selection.End(xlDown).Select
        R3 = ActiveCell.row
        
            
    Range("D2:BZ" & R3).Select
    Application.CutCopyMode = False
    Selection.Copy
    
    Sheets("Detailed sales report").Select
    Range("A6").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    
End Sub

Sub filterdetail_credit_note(ix As Integer)

Dim R4 As Integer

R4 = 2

    Sheets("Detailed sales report").Range("A7:BB10000").ClearContents
        
    Sheets("Finance overview by Item").Cells.AutoFilter
        
        Sheets("Finance overview by Item").Select
        Range("C2").Select
        Selection.End(xlDown).Select
        R4 = ActiveCell.row
         
    Sheets("Finance overview by Item").Select
    Range("$A$2:$BZ$2").AutoFilter
    ActiveSheet.Range("A2:BZ" & R4).AutoFilter Field:=3, Criteria1:=Sheets("seller_CN_index").Cells(ix, 8).Value
    ActiveSheet.Range("A2:BZ" & R4).AutoFilter Field:=51, Criteria1:="1"
            
        Range("C2").Select
        Selection.End(xlDown).Select
        R4 = ActiveCell.row
                
    Range("D2:BZ" & R4).Select
    Application.CutCopyMode = False
    Selection.Copy
    
    Sheets("Detailed sales report").Select
    Range("A6").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
End Sub

Sub filterdetail_credit_note_excel(ix As Integer)

Dim R5 As Integer

    Sheets("Detailed sales report").Range("A7:BB10000").ClearContents
    
    If Sheets("Summary Seller").Cells(19, 5).Value <> "-" Then
        
    Sheets("Finance overview by Item").Cells.AutoFilter
        
        Sheets("Finance overview by Item").Select
        Range("C2").Select
        Selection.End(xlDown).Select
        R5 = ActiveCell.row
        
    Sheets("Finance overview by Item").Select
    Range("$A$2:$BZ$2").AutoFilter
    ActiveSheet.Range("A2:BZ" & R5).AutoFilter Field:=3, Criteria1:=Sheets("Summary Seller").Range("B10").Value
    ActiveSheet.Range("A2:BZ" & R5).AutoFilter Field:=51, Criteria1:="1"
        
        Range("C2").Select
        Selection.End(xlDown).Select
        R5 = ActiveCell.row
        
    Range("D2:BZ" & R5).Select
    Application.CutCopyMode = False
    Selection.Copy
    
    Sheets("Detailed sales report").Select
    Range("A6").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    End If
    
End Sub


Sub filterdetail_invoice(ix As Integer)

Dim R6 As Integer

R6 = 2

    Sheets("Detailed sales report").Range("A7:BB10000").ClearContents
        
    Sheets("Finance overview by Item").Cells.AutoFilter
        
        Sheets("Finance overview by Item").Select
        Range("C2").Select
        Selection.End(xlDown).Select
        R6 = ActiveCell.row
        
    Sheets("Finance overview by Item").Select
    Range("$A$2:$BZ$2").AutoFilter
    ActiveSheet.Range("A2:BZ" & R6).AutoFilter Field:=3, Criteria1:=Sheets("seller_CN_index").Cells(ix, 9).Value
    ActiveSheet.Range("A2:BZ" & R6).AutoFilter Field:=52, Criteria1:="1"
        
        Range("C2").Select
        Selection.End(xlDown).Select
        R6 = ActiveCell.row
    
    Range("D2:BZ" & R6).Select
    Application.CutCopyMode = False
    Selection.Copy
    
    Sheets("Detailed sales report").Select
    Range("A6").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
End Sub


Sub filterdetail_invoice_excel(ix As Integer)

Dim R7 As Integer

    Sheets("Detailed sales report").Range("A7:BB10000").ClearContents
    
    If Sheets("Summary Seller").Cells(18, 5).Value <> "-" Then
        
        Sheets("Finance overview by Item").Cells.AutoFilter
    
        Sheets("Finance overview by Item").Select
        Range("C2").Select
        Selection.End(xlDown).Select
        R7 = ActiveCell.row
        
        Sheets("Finance overview by Item").Select
        Range("$A$2:$BZ$2").AutoFilter
        ActiveSheet.Range("A2:BZ" & R7).AutoFilter Field:=3, Criteria1:=Sheets("Summary Seller").Range("B10").Value
        ActiveSheet.Range("A2:BZ" & R7).AutoFilter Field:=52, Criteria1:="1"
            
            Range("C2").Select
            Selection.End(xlDown).Select
            R7 = ActiveCell.row
            
        Range("D2:BZ" & R7).Select
        Application.CutCopyMode = False
        Selection.Copy
        
        Sheets("Detailed sales report").Select
        Range("A6").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        
    End If
    
End Sub

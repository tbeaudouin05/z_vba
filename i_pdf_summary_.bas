Attribute VB_Name = "i_pdf_summary_"
Sub PDFSummary()
    
    Application.ScreenUpdating = False
    Dim FName As String, ws As Worksheet
    Dim i As Integer, wb As Workbook
    Dim MainFolder As String

    MainFolder = Sheets("Automatic PDF Generation").Range("C2").Value & Sheets("Seller_CN_index").Range("K4").Value & Sheets("Automatic PDF Generation").Range("C3").Value & " closing\Tools & Reports\Output\Seller Reports\"
    
    i = 2
    
    If Len(Dir(MainFolder, vbDirectory)) = 0 Then
            MkDir (MainFolder)
    End If
    
    Call show_all
    
    Sheets("Finance overview by Item").Cells.AutoFilter
    
    Do While Sheets("seller_CN_index").Cells(i, 1).Value <> ""
    
        Call filterdetail(i)
        
        With Sheets("Summary Seller")
        .Range("B10").Value = ""
        .Range("B10").Value = Sheets("seller_CN_index").Cells(i, 7).Value
        .calculate
        End With
        
        Sheets("Summary Seller").calculate
        Sheets("Detailed sales report").calculate
        
        If Sheets("Seller_CN_index").Range("J2").Value = "MPT" Then
        Sheets("Detailed sales report").Columns("N:N").EntireColumn.Hidden = True
        Sheets("Detailed sales report").Columns("AK:AK").EntireColumn.Hidden = True
        End If
        
        'customer shipping fee adapt for TW
        
        If Application.SumIf(Sheets("Detailed sales report").Range("R7:R1300"), "<>") = 0 Then
        Sheets("Detailed sales report").Columns("R:R").EntireColumn.Hidden = True
        End If
        
        'return shipping fee
        
        If Application.SumIf(Sheets("Detailed sales report").Range("Y7:Y1300"), "<>") = 0 Then
        Sheets("Detailed sales report").Columns("X:Y").EntireColumn.Hidden = True
        End If
        
        'vouchers
        
        If Application.SumIf(Sheets("Detailed sales report").Range("AA7:AA1300"), "<>") = 0 Then
        Sheets("Detailed sales report").Columns("Z:AA").EntireColumn.Hidden = True
        End If
        
        'cart rule
        
        If Application.SumIf(Sheets("Detailed sales report").Range("AB7:AB1300"), "<>") = 0 Then
        Sheets("Detailed sales report").Columns("AB:AB").EntireColumn.Hidden = True
        End If
        
        'delivery fee waiver
        
        If Application.SumIf(Sheets("Detailed sales report").Range("AC7:AC1300"), "<>") = 0 Then
        Sheets("Detailed sales report").Columns("AC:AC").EntireColumn.Hidden = True
        End If
        
        'return penalty waiver
              
        If Application.SumIf(Sheets("Detailed sales report").Range("AD7:AD1300"), "<>") = 0 Then
        Sheets("Detailed sales report").Columns("AD:AD").EntireColumn.Hidden = True
        End If
              
        'cancellation penalty waiver
        
        If Application.SumIf(Sheets("Detailed sales report").Range("AE7:AE1300"), "<>") = 0 Then
        Sheets("Detailed sales report").Columns("AE:AF").EntireColumn.Hidden = True
        End If
        
        'exceptional refund to seller
        
        If Application.SumIf(Sheets("Detailed sales report").Range("AG7:AG1300"), "<>") = 0 Then
        Sheets("Detailed sales report").Columns("AG:AG").EntireColumn.Hidden = True
        End If
        
        'production services fee
        
        
        If Application.SumIf(Sheets("Detailed sales report").Range("AH7:AH1300"), "<>") = 0 Then
        Sheets("Detailed sales report").Columns("AH:AH").EntireColumn.Hidden = True
        End If
        
        'correction of commission
        
        If Application.SumIf(Sheets("Detailed sales report").Range("AI7:AI1300"), "<>") = 0 Then
        Sheets("Detailed sales report").Columns("AI:AI").EntireColumn.Hidden = True
        End If
        
        'other seller revenues
        
        If Application.SumIf(Sheets("Detailed sales report").Range("AJ7:AJ1300"), "<>") = 0 Then
        Sheets("Detailed sales report").Columns("AJ:AJ").EntireColumn.Hidden = True
        End If
        
        'other fees
        
        Sheets("Summary Seller").calculate
        
        Call savepdfsumm(i, MainFolder)
        
        Sheets("Detailed sales report").Columns("A:AZ").EntireColumn.Hidden = False
        
        i = i + 1
        
    Loop
    
    Call hide_all
    
    Sheets("Automatic PDF Generation").Select
    
End Sub

Sub savepdfsumm(ix As Integer, path As String)
    Dim brand As String
    brand = legalized(Sheets("seller_CN_index").Cells(ix, 7).Value)
    
    FName = path & brand & " - Seller Report" & _
        " " & Sheets("seller_CN_index").Cells(2, 10).Value & ".pdf"
        
    Sheets("Detailed sales report").Select
    Range("A6").Select
    Selection.End(xlDown).Select
    lastRow = ActiveCell.row
    Sheets("Detailed sales report").PageSetup.PrintArea = "A1:AL" & lastRow
           
    Sheets(Array("Summary Seller", "Detailed sales report")).Select
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, filename:=FName _
        , Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
        :=False, OpenAfterPublish:=False
End Sub


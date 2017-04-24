Attribute VB_Name = "i_pdf_credit_note"
Sub pdf_credit_note()

    Application.ScreenUpdating = False
    Dim FName As String, ws As Worksheet
    Dim i As Integer, wb As Workbook
    Dim MainFolder As String
    Dim k As Integer
    Dim j As Integer
    
    Call show_all
    
    MainFolder = Sheets("Automatic PDF Generation").Range("C2").Value & Sheets("Seller_CN_index").Range("K4").Value & Sheets("Automatic PDF Generation").Range("C3").Value & " closing\Tools & Reports\Output\Credit Notes\"
    
    
    i = 2
            
    If Len(Dir(MainFolder, vbDirectory)) = 0 Then
            MkDir (MainFolder)
    End If
    
    Sheets("Finance overview by Item").Cells.AutoFilter
    
    Do While Sheets("seller_CN_index").Cells(i, 2).Value <> ""
    
        Call filterdetail_credit_note(i)
        
        With Sheets("Finance overview by Item")
        k = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
        End With
        
        With Sheets("Summary Seller")
        .Range("B10").Value = " "
        .Range("B10").Value = Sheets("seller_CN_index").Cells(i, 8).Value
        .calculate
        End With
        
        If k <= 21 Then
        Sheets("credit_note_less_21").calculate
        
        ElseIf k > 21 And k <= 68 Then
        Sheets("credit_note_less_68").calculate
        
        ElseIf k > 68 And k <= 115 Then
        Sheets("credit_note_less_115").calculate
        
        ElseIf k > 115 And k <= 162 Then
        Sheets("credit_note_less_162").calculate
        
        ElseIf k > 162 And k <= 200 Then
        Sheets("credit_note_less_200").calculate
        
        ElseIf k > 200 And k <= 250 Then
        Sheets("credit_note_less_250").calculate
        
        ElseIf k > 250 And k <= 300 Then
        Sheets("credit_note_less_300").calculate
        
        End If
        
        Sheets("Detailed sales report").calculate
        
        Call save_pdf_credit_note(i, MainFolder, k)
        
        i = i + 1
        
    Loop
    
    Sheets("Finance overview by Item").Cells.AutoFilter

    Call hide_all
    
    Sheets("Automatic PDF Generation").Select
    
End Sub

Sub save_pdf_credit_note(ix As Integer, path As String, kx As Integer)

    Dim brand As String
    
    brand = legalized(Sheets("seller_CN_index").Cells(ix, 8).Value)
    
    FName = path & brand & " - Credit_Note" & _
        " " & Sheets("seller_CN_index").Cells(2, 10).Value & ".pdf"
    

        If kx <= 21 Then
        Sheets("credit_note_less_21").ExportAsFixedFormat Type:=xlTypePDF, filename:=FName _
        , Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
        :=False, OpenAfterPublish:=False
        
        ElseIf kx > 21 And kx <= 68 Then
        Sheets("credit_note_less_68").ExportAsFixedFormat Type:=xlTypePDF, filename:=FName _
        , Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
        :=False, OpenAfterPublish:=False
    
        ElseIf kx > 68 And kx <= 115 Then
        Sheets("credit_note_less_115").ExportAsFixedFormat Type:=xlTypePDF, filename:=FName _
        , Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
        :=False, OpenAfterPublish:=False
        
        ElseIf kx > 115 And kx <= 162 Then
        Sheets("credit_note_less_162").ExportAsFixedFormat Type:=xlTypePDF, filename:=FName _
        , Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
        :=False, OpenAfterPublish:=False
        
        ElseIf kx > 162 And kx <= 200 Then
        Sheets("credit_note_less_200").ExportAsFixedFormat Type:=xlTypePDF, filename:=FName _
        , Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
        :=False, OpenAfterPublish:=False
        
        ElseIf kx > 200 And kx <= 250 Then
        Sheets("credit_note_less_250").ExportAsFixedFormat Type:=xlTypePDF, filename:=FName _
        , Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
        :=False, OpenAfterPublish:=False
        
        ElseIf kx > 250 And kx <= 300 Then
        Sheets("credit_note_less_300").ExportAsFixedFormat Type:=xlTypePDF, filename:=FName _
        , Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
        :=False, OpenAfterPublish:=False
        
        Else: MsgBox "You need to create a new credit note template with more lines AND change the macro sub save_pdf_credit_note"
       
        End If
    
End Sub


Attribute VB_Name = "CreateFile"
Sub save_pdf_credit_note_sheet(seller_name As String, path As String)
    Dim brand As String
    brand = legalized(seller_name)
    FName = path & brand & " - Credit_Note" & _
        " " & Sheets("seller_CN_index").Cells(2, 10).Value & ".pdf"
    Sheets("credit_note").ExportAsFixedFormat Type:=xlTypePDF, filename:=FName _
        , Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
        :=True, OpenAfterPublish:=False
End Sub


Sub savexls_sheet(seller_name As String, path As String)

    Dim brand
    
    brand = legalized(seller_name) 'legalized(Sheets("seller_CN_index").Cells(ix, 7).Value)
    
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


Sub savepdfsumm_sheet(seller_name As String, path As String)
    Dim brand As String
    brand = legalized(seller_name) 'legalized(Sheets("seller_CN_index").Cells(ix, 7).Value)
    
    FName = path & brand & " - Seller Report" & _
        " " & Sheets("seller_CN_index").Cells(2, 10).Value & ".pdf"
        
    Sheets("Detailed sales report").Select
    Range("A1").Select
    Selection.End(xlDown).Select
    lastRow = ActiveCell.row
    If lastRow = 6 Then
        lastRow = 7
    End If
    Sheets("Detailed sales report").PageSetup.PrintArea = "A1:AL" & lastRow
    Sheets(Array("Summary Seller", "Detailed sales report")).Select
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, filename:=FName _
        , Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
        :=False, OpenAfterPublish:=False
End Sub
Sub savepdfinv_sheet(seller_name As String, path As String)
    Dim brand As String
    brand = legalized(seller_name) ' legalized(Sheets("seller_CN_index").Cells(ix, 9).Value)
    
    FName = path & brand & " - Tax Invoice" & _
        " " & Sheets("seller_CN_index").Cells(2, 10).Value & ".pdf"

    Sheets("Tax Invoice").ExportAsFixedFormat Type:=xlTypePDF, filename:=FName _
        , Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
        :=False, OpenAfterPublish:=False
End Sub




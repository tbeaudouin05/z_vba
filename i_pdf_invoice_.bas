Attribute VB_Name = "i_pdf_invoice_"
Sub PDFInvoice()

    Application.ScreenUpdating = False
    Dim FName As String, ws As Worksheet
    Dim i As Integer, wb As Workbook
    Dim MainFolder As String
    
    MainFolder = Sheets("Automatic PDF Generation").Range("C2").Value & Sheets("Seller_CN_index").Range("K4").Value & Sheets("Automatic PDF Generation").Range("C3").Value & " closing\Tools & Reports\Output\Tax Invoices\"
    
    i = 2
            
    If Len(Dir(MainFolder, vbDirectory)) = 0 Then
            MkDir (MainFolder)
    End If
    
    Call show_all
    
    Sheets("Finance overview by Item").Cells.AutoFilter
    
    Do While Sheets("seller_CN_index").Cells(i, 3).Value <> ""
        Call filterdetail_invoice(i)
        
        With Sheets("Summary Seller")
        .Range("B10").Value = " "
        .Range("B10").Value = Sheets("seller_CN_index").Cells(i, 9).Value
        .calculate
        End With
        
        Sheets("Summary Seller").calculate
        
        Sheets("Tax Invoice").calculate
        
        Sheets("Detailed sales report").calculate
        
        Call savepdfinv(i, MainFolder)
        
        i = i + 1
        
    Loop
    
    Sheets("Finance overview by Item").Cells.AutoFilter

    Call hide_all
    
    Sheets("Automatic PDF Generation").Select
    
End Sub


Sub savepdfinv(ix As Integer, path As String)
    Dim brand As String
    brand = legalized(Sheets("seller_CN_index").Cells(ix, 9).Value)
    
    FName = path & brand & " - Tax Invoice" & _
        " " & Sheets("seller_CN_index").Cells(2, 10).Value & ".pdf"

    Sheets("Tax Invoice").ExportAsFixedFormat Type:=xlTypePDF, filename:=FName _
        , Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
        :=False, OpenAfterPublish:=False
End Sub



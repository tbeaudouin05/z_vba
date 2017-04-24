Attribute VB_Name = "runmacro"
Public row_note_sheet As Integer
Public index_sum As Integer
Public country As String
Public flag_tw As Boolean

Sub isCheckTw()
    flag_tw = False
    If Sheets("Seller_CN_index").Range("K2").Value = "MPT" Then
        flag_tw = True
    End If
End Sub

Sub Run_credit_note()
    Application.ScreenUpdating = False
    Call pdf_credit_note_sheet
    Application.ScreenUpdating = True
    Dim MainFolder As String
    MainFolder = Sheets("Automatic PDF Generation").Range("C2").Value & Sheets("Seller_CN_index").Range("K4").Value & Sheets("Automatic PDF Generation").Range("C3").Value & " closing\Tools & Reports\Output\Credit Notes\"
    Call endmacro(MainFolder)
    
End Sub

Sub Run_ExcelReport()
    Application.ScreenUpdating = False
    Call changeColumns
    Call ExcelReport_sheet
    Dim MainFolder As String
    Application.ScreenUpdating = True
    MainFolder = Sheets("Automatic PDF Generation").Range("C2").Value & Sheets("Seller_CN_index").Range("K4").Value & Sheets("Automatic PDF Generation").Range("C3").Value & " closing\Tools & Reports\Output\Excel Files\"
    Call endmacro(MainFolder)
    
End Sub

Sub Run_PdfSummary()
    Application.ScreenUpdating = False
    Call PDFSummary_sheet
    Application.ScreenUpdating = True
    Dim MainFolder As String
    MainFolder = Sheets("Automatic PDF Generation").Range("C2").Value & Sheets("Seller_CN_index").Range("K4").Value & Sheets("Automatic PDF Generation").Range("C3").Value & " closing\Tools & Reports\Output\Seller Reports\"
    Call endmacro(MainFolder)
    
End Sub

Sub Run_PdfInvoice()
    Application.ScreenUpdating = False
    Call PDFInvoice_sheet
    Application.ScreenUpdating = True
    Dim MainFolder As String
    MainFolder = Sheets("Automatic PDF Generation").Range("C2").Value & Sheets("Seller_CN_index").Range("K4").Value & Sheets("Automatic PDF Generation").Range("C3").Value & " closing\Tools & Reports\Output\Tax Invoices\"
    Call endmacro(MainFolder)
    
End Sub


Sub pdf_credit_note_sheet() 'pdf_credit_note
    Application.ScreenUpdating = False
    Dim FName As String, ws As Worksheet
    Dim i As Integer, wb As Workbook
    Dim MainFolder As String
    Dim k As Integer
    Dim j As Integer
    MainFolder = Sheets("Automatic PDF Generation").Range("C2").Value & Sheets("Seller_CN_index").Range("K4").Value & Sheets("Automatic PDF Generation").Range("C3").Value & " closing\Tools & Reports\Output\Credit Notes\"
    i = 2
    On Error GoTo create_temp_file
    Call show_all_sheet
    Call createSheet
    Call createFolder(MainFolder)
    Dim seller_name As String
    Sheets("Finance overview by Item").Cells.AutoFilter
    Call isCheckTw
    Call formatingCell
    Call showHiderowsfortw
    Do While Sheets("seller_CN_index").Cells(i, 2).Value <> ""
        seller_name = Sheets("seller_CN_index").Cells(i, 8).Value
        With Sheets("Summary Seller")
        .Range("B10").Value = " "
        .Range("B10").Value = seller_name
        .calculate
        End With
        Sheets("Detailed sales report").calculate
        Sheets("Summary Seller").calculate
        Sheets("credit_note_less_21").calculate
        Call filterdetail_credit_note_sheet(seller_name)
        Call save_pdf_credit_note_sheet(seller_name, MainFolder)
        i = i + 1
      
    Loop
create_temp_file:
    Call deleteSheet
    Call hide_sheet
    Sheets("Automatic PDF Generation").Select
End Sub

Sub Export_Seller(seller_name As String)
    With Sheets("Summary Seller")
        .Range("B10").Value = ""
        .Range("B10").Value = seller_name 'Sheets("seller_CN_index").Cells(i, 7).Value
        .calculate
        End With
        'seller_name = Sheets("seller_CN_index").Cells(i, 7).Value
        Call filterdetail_invoice_excel_sheet(seller_name)
        Sheets("Detailed sales report").calculate
        Sheets("Summary Seller").calculate
        Sheets("Tax Invoice").calculate
        Call hardcode_invoice
        Call filterdetail_sheet(seller_name)
        Sheets("Detailed sales report").calculate
        Sheets("Summary Seller").calculate
        Sheets("credit_note_less_21").calculate
        Call filterdetail_credit_note_inexcel_sheet
End Sub

Sub ExcelReport_sheet()
    Application.ScreenUpdating = False
    row_note_sheet = 60
    Dim FName As String, ws As Worksheet
    Dim i As Integer, k As Integer, wb As Workbook
    Dim MainFolder As String
    Dim CN As String
    index_sum = 1
    'MainFolder = Sheets("Automatic PDF Generation").Range("output_path[Value]").Value & "\Excel Files\"
    MainFolder = Sheets("Automatic PDF Generation").Range("C2").Value & Sheets("Seller_CN_index").Range("K4").Value & Sheets("Automatic PDF Generation").Range("C3").Value & " closing\Tools & Reports\Output\Excel Files\"
    
   ' On Error GoTo create_temp_file2
    Call show_all_sheet
    Call createSheet
    Call createFolder(MainFolder)
    i = 2
    Sheets("Finance overview by Item").Cells.AutoFilter
    Dim seller_name As String
    seller_name = ""
    Call isCheckTw
    Call formatingCell
    Call showHiderowsfortw
    Do While Sheets("seller_CN_index").Cells(i, 1).Value <> ""
        seller_name = Sheets("seller_CN_index").Cells(i, 7).Value
        Call Export_Seller(seller_name)
        Call savexls_sheet(seller_name, MainFolder)
        i = i + 1
       
    Loop
'create_temp_file2:
    Call deleteSheet
    Call hide_sheet
    Sheets("Automatic PDF Generation").Select
    'Application.ScreenUpdating = True
End Sub

Sub PDFSummary_sheet()
    Application.ScreenUpdating = False
    Dim FName As String, ws As Worksheet
    Dim i As Integer, k As Integer, wb As Workbook
    Dim MainFolder As String
    MainFolder = Sheets("Automatic PDF Generation").Range("C2").Value & Sheets("Seller_CN_index").Range("K4").Value & Sheets("Automatic PDF Generation").Range("C3").Value & " closing\Tools & Reports\Output\Seller Reports\"
    Call show_all_sheet
    Call createFolder(MainFolder)
    i = 2
    Sheets("Finance overview by Item").Cells.AutoFilter
    Dim seller_name As String
    Call isCheckTw
    Call formatingCell
    Call showHiderowsfortw
    Do While Sheets("seller_CN_index").Cells(i, 1).Value <> ""
        Call filterdetail(i)
        seller_name = Sheets("seller_CN_index").Cells(i, 7).Value
        With Sheets("Summary Seller")
        .Range("B10").Value = ""
        .Range("B10").Value = seller_name 'Sheets("seller_CN_index").Cells(i, 7).Value
        .calculate
        End With
        Sheets("Summary Seller").calculate
        Sheets("Detailed sales report").calculate
        Call show_hiderows
        Sheets("Summary Seller").calculate
        Call savepdfsumm_sheet(seller_name, MainFolder)
        Sheets("Detailed sales report").Columns("A:AZ").EntireColumn.Hidden = False
        i = i + 1
       
    Loop
    Call hide_sheet
    Sheets("Automatic PDF Generation").Select
End Sub

Sub PDFInvoice_sheet()

    Application.ScreenUpdating = False
    Dim FName As String, ws As Worksheet
    Dim i As Integer, wb As Workbook
    Dim MainFolder As String
    MainFolder = Sheets("Automatic PDF Generation").Range("C2").Value & Sheets("Seller_CN_index").Range("K4").Value & Sheets("Automatic PDF Generation").Range("C3").Value & " closing\Tools & Reports\Output\Tax Invoices\"
    i = 2
    Call show_all_sheet
    Call createFolder(MainFolder)
    Sheets("Finance overview by Item").Cells.AutoFilter
    Call isCheckTw
    Call formatingCell
    Call showHiderowsfortw
    Do While Sheets("seller_CN_index").Cells(i, 3).Value <> ""
        Call filterdetail_invoice(i)
        With Sheets("Summary Seller")
        .Range("B10").Value = " "
        .Range("B10").Value = Sheets("seller_CN_index").Cells(i, 9).Value
        .calculate
        End With
        Sheets("Detailed sales report").calculate
        Sheets("Summary Seller").calculate
        Sheets("Tax Invoice").calculate
        Call savepdfinv(i, MainFolder)
        i = i + 1
    Loop
    Sheets("Finance overview by Item").Cells.AutoFilter
    Call hide_sheet
    Sheets("Automatic PDF Generation").Select
End Sub

Sub generate_hk_sheet()

'Application.ScreenUpdating = False

Call show_all_sheet

'generate hk

Call filter_country("hk")
        
Call seller_CN_index_and_overviews
    
Call adapt_template_for_TW

Sheets("Seller_CN_index").calculate

Call generate_all_sheet

'Application.ScreenUpdating = True

End Sub


Sub Run_generate_hk()
    Application.ScreenUpdating = False
    Call generate_hk_sheet
    Application.ScreenUpdating = True
    Dim MainFolder As String
    MainFolder = Sheets("Automatic PDF Generation").Range("C2").Value & Sheets("Seller_CN_index").Range("K4").Value & Sheets("Automatic PDF Generation").Range("C3").Value & " closing\Tools & Reports\Output\"
    Call endmacro(MainFolder)
    
End Sub
Sub Run_generate_sg()
    Application.ScreenUpdating = False
    Call generate_sg_sheet
    Application.ScreenUpdating = True
    Dim MainFolder As String
    MainFolder = Sheets("Automatic PDF Generation").Range("C2").Value & Sheets("Seller_CN_index").Range("K4").Value & Sheets("Automatic PDF Generation").Range("C3").Value & " closing\Tools & Reports\Output\"
    Call endmacro(MainFolder)
    
End Sub
Sub Run_generate_tw()
    Application.ScreenUpdating = False
    Call generate_tw_sheet
    Application.ScreenUpdating = True
    Dim MainFolder As String
    MainFolder = Sheets("Automatic PDF Generation").Range("C2").Value & Sheets("Seller_CN_index").Range("K4").Value & Sheets("Automatic PDF Generation").Range("C3").Value & " closing\Tools & Reports\Output\"
    Call endmacro(MainFolder)
    
End Sub
Sub Run_generate_my()
    Application.ScreenUpdating = False
    Call generate_my_sheet
    Application.ScreenUpdating = True
    Dim MainFolder As String
    MainFolder = Sheets("Automatic PDF Generation").Range("C2").Value & Sheets("Seller_CN_index").Range("K4").Value & Sheets("Automatic PDF Generation").Range("C3").Value & " closing\Tools & Reports\Output\"
    Call endmacro(MainFolder)
    
End Sub

Sub generate_sg_sheet()

'Application.ScreenUpdating = False

Call show_all_sheet

'generate hk

Call filter_country("sg")
        
Call seller_CN_index_and_overviews
    
Call adapt_template_for_TW

Sheets("Seller_CN_index").calculate

Call generate_all_sheet

'Application.ScreenUpdating = True

End Sub

Sub generate_tw_sheet()

'Application.ScreenUpdating = False

Call show_all_sheet

'generate hk

Call filter_country("tw")
        
Call seller_CN_index_and_overviews
    
Call adapt_template_for_TW

Sheets("Seller_CN_index").calculate

Call generate_all_sheet

'Application.ScreenUpdating = True

End Sub

Sub generate_my_sheet()

'Application.ScreenUpdating = False

Call show_all_sheet

'generate hk

Call filter_country("my")
        
Call seller_CN_index_and_overviews
    
Call adapt_template_for_TW

Sheets("Seller_CN_index").calculate

Call generate_all_sheet

'Application.ScreenUpdating = True

End Sub

Sub generate_all_countries_sheet()

Application.ScreenUpdating = False

Call show_all_sheet

'generate hk

Call filter_country("hk")
        
Call seller_CN_index_and_overviews
    
Call adapt_template_for_TW

Sheets("Seller_CN_index").calculate

Call generate_all_sheet


' generate sg

Call show_all_sheet

Call filter_country("sg")
        
Call seller_CN_index_and_overviews
    
Call adapt_template_for_TW

Sheets("Seller_CN_index").calculate

Call generate_all_sheet

        
' generate tw

Call show_all_sheet

Call filter_country("tw")
        
Call seller_CN_index_and_overviews
    
Call adapt_template_for_TW

Sheets("Seller_CN_index").calculate

Call generate_all_sheet


' generate my

Call show_all_sheet

Call filter_country("my")
        
Call seller_CN_index_and_overviews
    
Call adapt_template_for_TW

Sheets("Seller_CN_index").calculate

Call generate_all_sheet
Application.ScreenUpdating = True
Dim MainFolder As String
MainFolder = Sheets("Automatic PDF Generation").Range("C2").Value
Call endmacro(MainFolder)
End Sub

Sub Run_Generate_all()
    Application.ScreenUpdating = False
    Call generate_all_sheet
    Application.ScreenUpdating = True
    Dim MainFolder As String
    MainFolder = Sheets("Automatic PDF Generation").Range("C2").Value & Sheets("Seller_CN_index").Range("K4").Value & Sheets("Automatic PDF Generation").Range("C3").Value & " closing\Tools & Reports\Output"
    Call endmacro(MainFolder)
    
End Sub
Sub generate_all_sheet()
    Call overview
    Call ExcelReport_sheet
    Call PDFSummary_sheet
    Call PDFInvoice_sheet
    Call pdf_credit_note_sheet
End Sub


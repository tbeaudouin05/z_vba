Attribute VB_Name = "a_prepare_data"
    Sub calculate()
    
    Call show_all
        
    Sheets("Input").calculate
    
    Sheets("Seller_CN_index").calculate
        
    ActiveWorkbook.RefreshAll
    
    Sheets("Seller_CN_index").calculate
    
    Application.ScreenUpdating = False
    
    Call check_sellers_data
    Call check_soi_data
    Call check_historic
    Call check_disputes
    Call check_ap_aging
    Call check_promotion_data
    
    Application.ScreenUpdating = True
           
    MsgBox ("Model has been calculated for all countries in " & Sheets("Seller_CN_index").Range("J2").Value)
    
    Call hide_all
    
    End Sub
    
    Sub update_data_path()
    
    Application.ScreenUpdating = False
    
    Call populate_empty_template
    
    Call erase_blank_rows
    
    Application.ScreenUpdating = True
    
    ActiveWorkbook.Connections("Query - Parameter").Refresh
    
    MsgBox ("Path to data folder and year&month were updated")

    End Sub
    
    Sub format_data_country()
    
    Application.ScreenUpdating = False
    
    Call show_all
    
    Call seller_CN_index_and_overviews
    
    Call adapt_template_for_TW

    Sheets("Seller_CN_index").calculate
    
    Call hide_all
    
    Application.ScreenUpdating = True
    
    MsgBox ("Model has been formatted for " & Sheets("Seller_CN_index").Range("K3").Value & " in " & Sheets("Seller_CN_index").Range("J2").Value)

    End Sub
    
    Sub filter_country(c As String)
 country = c
 
 Sheets("Sellers data for macro").Select
 Cells.Select
 Selection.ClearContents
   
 Sheets("Orders data for macro & pivot").Select

    ActiveSheet.PivotTables("soi_data").PivotFields( _
        "[soi_data].[Venture code].[Venture code]").VisibleItemsList = Array( _
        "[soi_data].[Venture code].&[" & c & "]")
        
    
 Sheets("Sellers data for macro_").Select
 
    ActiveSheet.ListObjects("sellers_data").Range.AutoFilter Field:=24, _
        Criteria1:=c
        
Sheets("historic_for_credit_note").Select

    ActiveSheet.ListObjects("historic").Range.AutoFilter Field:=17, Criteria1:= _
        c
        
    Sheets("Sellers data for macro_").Select
    Cells.Select
    Selection.Copy
    Sheets("Sellers data for macro").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
End Sub
    
    Sub filter_sg()
    
    Application.ScreenUpdating = False
    
    Call show_all
    
    Call filter_country("sg")
    
    Application.ScreenUpdating = True
    
    Call format_data_country
    
    Sheets("Seller_CN_index").calculate
    Call createDataValidation
    Sheets("Automatic PDF Generation").calculate
    
    Call hide_all
    
    End Sub
    
    Sub filter_hk()
    
    Application.ScreenUpdating = False
    
    Call show_all
    
    Call filter_country("hk")
    
    Application.ScreenUpdating = True
    
    Call format_data_country
    
    Sheets("Seller_CN_index").calculate
    Call createDataValidation
    Sheets("Automatic PDF Generation").calculate
    
    Call hide_all
    
    End Sub
    
    Sub filter_tw()
    
    Application.ScreenUpdating = False
    
    Call show_all
    
    Call filter_country("tw")
    
    Application.ScreenUpdating = True
    
    Call format_data_country
    
    Sheets("Seller_CN_index").calculate
    Call createDataValidation
    Sheets("Automatic PDF Generation").calculate
    
    Call hide_all
    
    End Sub
    
    Sub filter_my()
    
    Application.ScreenUpdating = False
    
    Call show_all
    
    Call filter_country("my")
    
    Application.ScreenUpdating = True
    
    Call format_data_country
    
    Sheets("Seller_CN_index").calculate
    Call createDataValidation
    Sheets("Automatic PDF Generation").calculate
    
    Call hide_all
    
    End Sub
    
    Sub check_sellers_data()
    
    Dim R11 As Integer
    
    Sheets("Sellers data for macro_").Select
        
    ActiveSheet.ListObjects("sellers_data").Range.AutoFilter Field:=24
        
        Range("X2").Select
        Selection.End(xlDown).Select
        R11 = ActiveCell.row
    
    Range("X2:X" & R11).Select
    Selection.Copy
    
    Sheets("Automatic PDF Generation").Select
    Range("G1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    ActiveSheet.Range("$G:$G").RemoveDuplicates Columns:=1, Header:= _
        xlYes
    
    Range("G1").Value = "sellers_data"
    
    End Sub
    
    Sub check_soi_data()
    
    Dim R12 As Integer
    
    Sheets("Orders data for macro & pivot").Select
        
    ActiveSheet.PivotTables("soi_data").PivotFields( _
        "[soi_data].[Venture code].[Venture code]").VisibleItemsList = Array("")
        
    Columns("C:C").Select
    Selection.Copy
    
    Sheets("Automatic PDF Generation").Select
    Range("H1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    ActiveSheet.Range("$H:$H").RemoveDuplicates Columns:=1, Header:= _
        xlYes
    
    Range("H1").Value = "soi_data"
    
    End Sub
    
    Sub check_historic()
    
    Sheets("historic_for_credit_note").Select
        
    ActiveSheet.ListObjects("historic").Range.AutoFilter Field:=17
    
    Range("Q:Q").Select
    Selection.Copy
    
    Sheets("Automatic PDF Generation").Select
    Range("I1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    ActiveSheet.Range("$I:$I").RemoveDuplicates Columns:=1, Header:= _
        xlYes
    
    Range("I1").Value = "historic"
    
    End Sub

  Sub check_disputes()
    
    Sheets("disputes").Select
        
    ActiveSheet.ListObjects("disputes").Range.AutoFilter Field:=27
    
    Range("AA:AA").Select
    Selection.Copy
    
    Sheets("Automatic PDF Generation").Select
    Range("J1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    ActiveSheet.Range("$J:$J").RemoveDuplicates Columns:=1, Header:= _
        xlYes
    
    Range("J1").Value = "disputes?"
    
    End Sub


  Sub check_ap_aging()
    
    Sheets("ap_aging").Select
        
    ActiveSheet.ListObjects("ap_aging").Range.AutoFilter Field:=27
    
    Range("AA:AA").Select
    Selection.Copy
    
    Sheets("Automatic PDF Generation").Select
    Range("K1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    ActiveSheet.Range("$K:$K").RemoveDuplicates Columns:=1, Header:= _
        xlYes
    
    Range("K1").Value = "ap_aging?"
    
    End Sub
    
    
  Sub check_promotion_data()
    
    Sheets("promotion_data").Select
        
    ActiveSheet.ListObjects("promotion_data").Range.AutoFilter Field:=7
    
    Range("G:G").Select
    Selection.Copy
    
    Sheets("Automatic PDF Generation").Select
    Range("L1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    ActiveSheet.Range("$L:$L").RemoveDuplicates Columns:=1, Header:= _
        xlYes
    
    Range("L1").Value = "promotion_data?"
    
    End Sub


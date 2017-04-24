Attribute VB_Name = "c_adapt_template_for_TW_"
Sub adapt_template_for_TW()
Attribute adapt_template_for_TW.VB_ProcData.VB_Invoke_Func = " \n14"

Sheets("Seller_CN_index").calculate

If Sheets("Seller_CN_index").Range("K2").Value = "MPT" Then

Sheets("Summary Seller").Select
    Rows("30:30").Select
    Selection.EntireRow.Hidden = False
    Rows("31:32").Select
    Selection.EntireRow.Hidden = False
    Rows("54:55").Select
    Selection.EntireRow.Hidden = False
    Rows("68:73").Select
    Selection.EntireRow.Hidden = False
    
    Range("C24:E58").Select
    Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
    
Sheets("Tax Invoice").Select
    Rows("32:33").Select
    Selection.EntireRow.Hidden = False
    Rows("53:54").Select
    Selection.EntireRow.Hidden = False
    Rows("74:75").Select
    Selection.EntireRow.Hidden = False
    
    Rows("22:60").Select
    Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
    
Sheets("Detailed sales report").Select
    Range("H7:H5000,K7:N5000,Q7:S5000,R7:S5000,V7:V5000,X7:AZ5000").Select
    Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
    Rows("4:4").Select
    Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
    
Sheets("Finance overview by seller").Select
    Cells.Select
    Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
    
Sheets("Finance overview by Item").Select
    Range("K:K,N:Q,S:V,Y:AP,AT:AV").Select
    Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
      
Sheets(Array("credit_note_less_21", "credit_note_less_68", "credit_note_less_115", "credit_note_less_162")).Select
    Sheets("credit_note_less_21").Activate
    Rows("21:400").Select
    Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
    Range("A:F,J:J").Select
    Selection.NumberFormat = "0"
    Columns("I:I").Select
    Selection.NumberFormat = "dd/mm/yyyy"
    
Sheets("credit_note_less_21").Select
    Rows("42:42").Select
    Selection.EntireRow.Hidden = False
    
Sheets("credit_note_less_68").Select
    Rows("89:89").Select
    Selection.EntireRow.Hidden = False
    
Sheets("credit_note_less_115").Select
    Rows("136:136").Select
    Selection.EntireRow.Hidden = False

Sheets("credit_note_less_162").Select
    Rows("183:183").Select
    Selection.EntireRow.Hidden = False
    
    
Else

Sheets("Summary Seller").Select
    Rows("30:30").Select
    Selection.EntireRow.Hidden = True
    Rows("31:32").Select
    Selection.EntireRow.Hidden = True
    Rows("54:55").Select
    Selection.EntireRow.Hidden = True
    Rows("68:73").Select
    Selection.EntireRow.Hidden = True
    
    Range("C24:E58").Select
    Selection.NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
    
Sheets("Tax Invoice").Select
    Rows("32:33").Select
    Selection.EntireRow.Hidden = True
    Rows("53:54").Select
    Selection.EntireRow.Hidden = True
    Rows("74:75").Select
    Selection.EntireRow.Hidden = True
    
    Rows("22:60").Select
    Selection.NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
    
Sheets("Detailed sales report").Select
    Range("H7:H5000,K7:N5000,Q7:S5000,R7:S5000,V7:V5000,X7:AZ5000").Select
    Selection.NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
    Rows("4:4").Select
    Selection.NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
    
Sheets("Finance overview by seller").Select
    Cells.Select
    Selection.NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
    
Sheets("Finance overview by Item").Select
    Range("K:K,N:Q,S:V,Y:AP,AT:AV").Select
    Selection.NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
      
Sheets(Array("credit_note_less_21", "credit_note_less_68", "credit_note_less_115", "credit_note_less_162")).Select
    Sheets("credit_note_less_21").Activate
    Rows("21:400").Select
    Selection.NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
    Range("A:F,J:J").Select
    Selection.NumberFormat = "0"
    Columns("I:I").Select
    Selection.NumberFormat = "dd/mm/yyyy"
    
Sheets("credit_note_less_21").Select
    Rows("42:42").Select
    Selection.EntireRow.Hidden = True
    
Sheets("credit_note_less_68").Select
    Rows("89:89").Select
    Selection.EntireRow.Hidden = True
    
Sheets("credit_note_less_115").Select
    Rows("136:136").Select
    Selection.EntireRow.Hidden = True

Sheets("credit_note_less_162").Select
    Rows("183:183").Select
    Selection.EntireRow.Hidden = True
    
End If

End Sub

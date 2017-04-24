Attribute VB_Name = "p_show_hide"
Sub show_all()

    Sheets("Input").Visible = True
    Sheets("INPUT>>").Visible = True
    Sheets("Orders data for macro & pivot").Visible = True
    Sheets("Sellers data for macro").Visible = True
    Sheets("seller_CN_index").Visible = True
    Sheets("seller_CN_index_").Visible = True
    Sheets("historic_for_credit_note").Visible = True
    Sheets("REPORTING>>").Visible = True
    Sheets("Automatic PDF Generation").Visible = True
    Sheets("FINANCE OVERVIEW>>").Visible = True
    Sheets("Finance overview by seller").Visible = True
    Sheets("Finance overview by seller_").Visible = True
    Sheets("Finance overview by Item").Visible = True
    Sheets("REPORT TEMPLATE->").Visible = True
    Sheets("Summary Seller").Visible = True
    Sheets("Detailed sales report").Visible = True
    Sheets("Tax Invoice").Visible = True
    Sheets("Tax Invoice_").Visible = True
    Sheets("credit_note").Visible = True
    Sheets("credit_note_less_21").Visible = True
    Sheets("credit_note_less_68").Visible = True
    Sheets("credit_note_less_115").Visible = True
    Sheets("credit_note_less_162").Visible = True
    Sheets("credit_note_less_200").Visible = True
    Sheets("credit_note_less_250").Visible = True
    Sheets("credit_note_less_300").Visible = True
    Sheets("disputes").Visible = True
    Sheets("ap_aging").Visible = True
    Sheets("promotion_data").Visible = True
    Sheets("Sellers data for macro_").Visible = True
    
    Sheets("Automatic PDF Generation").Select
    

End Sub

Sub hide_all()

    Sheets("Input").Visible = False
    Sheets("INPUT>>").Visible = False
    Sheets("Orders data for macro & pivot").Visible = False
    Sheets("Sellers data for macro").Visible = False
    Sheets("seller_CN_index").Visible = False
    Sheets("seller_CN_index_").Visible = False
    Sheets("historic_for_credit_note").Visible = False
    Sheets("REPORTING>>").Visible = False
    Sheets("Automatic PDF Generation").Visible = True
    Sheets("FINANCE OVERVIEW>>").Visible = False
    Sheets("Finance overview by seller").Visible = False
    Sheets("Finance overview by seller_").Visible = False
    Sheets("Finance overview by Item").Visible = False
    Sheets("REPORT TEMPLATE->").Visible = False
    Sheets("Summary Seller").Visible = False
    Sheets("Detailed sales report").Visible = False
    Sheets("Tax Invoice").Visible = False
    Sheets("Tax Invoice_").Visible = False
    Sheets("credit_note").Visible = False
    Sheets("credit_note_less_21").Visible = False
    Sheets("credit_note_less_68").Visible = False
    Sheets("credit_note_less_115").Visible = False
    Sheets("credit_note_less_162").Visible = False
    Sheets("credit_note_less_200").Visible = False
    Sheets("credit_note_less_250").Visible = False
    Sheets("credit_note_less_300").Visible = False
    Sheets("disputes").Visible = False
    Sheets("ap_aging").Visible = False
    Sheets("promotion_data").Visible = False
    Sheets("Sellers data for macro_").Visible = False
    

End Sub


Sub show_hiderows()
        If Sheets("Seller_CN_index").Range("J2").Value = "MPT" Then
            Sheets("Detailed sales report").Columns("N:N").EntireColumn.Hidden = True
            Sheets("Detailed sales report").Columns("AK:AK").EntireColumn.Hidden = True
            Else
                Sheets("Detailed sales report").Columns("N:N").EntireColumn.Hidden = False
            Sheets("Detailed sales report").Columns("AK:AK").EntireColumn.Hidden = False
        End If
    'customer shipping fee adapt for TW
        
        If Application.SumIf(Sheets("Detailed sales report").Range("R7:R1300"), "<>") = 0 Then
            Sheets("Detailed sales report").Columns("R:R").EntireColumn.Hidden = True
            Else
                Sheets("Detailed sales report").Columns("R:R").EntireColumn.Hidden = False
        End If
        
        'return shipping fee
        
        If Application.SumIf(Sheets("Detailed sales report").Range("Y7:Y1300"), "<>") = 0 Then
                Sheets("Detailed sales report").Columns("X:Y").EntireColumn.Hidden = True
            Else
                Sheets("Detailed sales report").Columns("X:Y").EntireColumn.Hidden = False
        End If
        
        'vouchers
        
        If Application.SumIf(Sheets("Detailed sales report").Range("AA7:AA1300"), "<>") = 0 Then
                Sheets("Detailed sales report").Columns("Z:AA").EntireColumn.Hidden = True
            Else
                Sheets("Detailed sales report").Columns("Z:AA").EntireColumn.Hidden = False
        End If
        
        'cart rule
        
        If Application.SumIf(Sheets("Detailed sales report").Range("AB7:AB1300"), "<>") = 0 Then
                Sheets("Detailed sales report").Columns("AB:AB").EntireColumn.Hidden = True
            Else
                Sheets("Detailed sales report").Columns("AB:AB").EntireColumn.Hidden = False
        End If
        
        'delivery fee waiver
        
        If Application.SumIf(Sheets("Detailed sales report").Range("AC7:AC1300"), "<>") = 0 Then
                Sheets("Detailed sales report").Columns("AC:AC").EntireColumn.Hidden = True
            Else
                Sheets("Detailed sales report").Columns("AC:AC").EntireColumn.Hidden = False
        End If
        
        'return penalty waiver
              
        If Application.SumIf(Sheets("Detailed sales report").Range("AD7:AD1300"), "<>") = 0 Then
                Sheets("Detailed sales report").Columns("AD:AD").EntireColumn.Hidden = True
            Else
                Sheets("Detailed sales report").Columns("AD:AD").EntireColumn.Hidden = False
        End If
              
        'cancellation penalty waiver
        
        If Application.SumIf(Sheets("Detailed sales report").Range("AE7:AE1300"), "<>") = 0 Then
                Sheets("Detailed sales report").Columns("AE:AF").EntireColumn.Hidden = True
            Else
                Sheets("Detailed sales report").Columns("AE:AF").EntireColumn.Hidden = False
        End If
        
        'exceptional refund to seller
        
        If Application.SumIf(Sheets("Detailed sales report").Range("AG7:AG1300"), "<>") = 0 Then
                Sheets("Detailed sales report").Columns("AG:AG").EntireColumn.Hidden = True
            Else
                Sheets("Detailed sales report").Columns("AG:AG").EntireColumn.Hidden = False
        End If
        
        'production services fee
        
        
        If Application.SumIf(Sheets("Detailed sales report").Range("AH7:AH1300"), "<>") = 0 Then
                Sheets("Detailed sales report").Columns("AH:AH").EntireColumn.Hidden = True
            Else
                Sheets("Detailed sales report").Columns("AH:AH").EntireColumn.Hidden = False
        End If
        
        'correction of commission
        
        If Application.SumIf(Sheets("Detailed sales report").Range("AI7:AI1300"), "<>") = 0 Then
                Sheets("Detailed sales report").Columns("AI:AI").EntireColumn.Hidden = True
            Else
                Sheets("Detailed sales report").Columns("AI:AI").EntireColumn.Hidden = False
        End If
        
        'other seller revenues
        
        If Application.SumIf(Sheets("Detailed sales report").Range("AJ7:AJ1300"), "<>") = 0 Then
                Sheets("Detailed sales report").Columns("AJ:AJ").EntireColumn.Hidden = True
            Else
                Sheets("Detailed sales report").Columns("AJ:AJ").EntireColumn.Hidden = False
        End If
        
        'other fees
End Sub

Sub showHiderowsfortw()
 If flag_tw = True Then
        Sheets("Tax Invoice").Rows("57").EntireRow.Hidden = False
        Sheets("Tax Invoice_").Rows("57").EntireRow.Hidden = False
    Else
        Sheets("Tax Invoice").Rows("57").EntireRow.Hidden = True
        Sheets("Tax Invoice_").Rows("57").EntireRow.Hidden = True
 End If
End Sub


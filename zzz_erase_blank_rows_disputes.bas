Attribute VB_Name = "zzz_erase_blank_rows_disputes"
Sub erase_blank_rows()

Application.DisplayAlerts = False

Dim wb1, wb2, wb3, wb4, wb5, wb6, wb7, wb8, wb9, wb10, wb11, wb12 As Excel.Workbook
Dim Rz As Integer

    path_to_folder = Sheets("Automatic PDF Generation").Range("C2").Value
    year_month = Sheets("Automatic PDF Generation").Range("C3").Value
    
    
    path_to_disputes_tw = path_to_folder & "\M005) Marketplace TW\MPT " & year_month & " closing\Tools & Reports\Input\disputes.xlsx"
    path_to_disputes_sg = path_to_folder & "\M006) Marketplace SG\MPS " & year_month & " closing\Tools & Reports\Input\disputes.xlsx"
    path_to_disputes_hk = path_to_folder & "\M007) Marketplace HK\MPH " & year_month & " closing\Tools & Reports\Input\disputes.xlsx"
    path_to_disputes_my = path_to_folder & "\M009) Marketplace MY\MPM " & year_month & " closing\Tools & Reports\Input\disputes.xlsx"
    
    path_to_ap_aging_tw = path_to_folder & "\M005) Marketplace TW\MPT " & year_month & " closing\Tools & Reports\Input\ap_aging.xlsx"
    path_to_ap_aging_sg = path_to_folder & "\M006) Marketplace SG\MPS " & year_month & " closing\Tools & Reports\Input\ap_aging.xlsx"
    path_to_ap_aging_hk = path_to_folder & "\M007) Marketplace HK\MPH " & year_month & " closing\Tools & Reports\Input\ap_aging.xlsx"
    path_to_ap_aging_my = path_to_folder & "\M009) Marketplace MY\MPM " & year_month & " closing\Tools & Reports\Input\ap_aging.xlsx"
    
    path_to_promotion_tw = path_to_folder & "\M005) Marketplace TW\MPT " & year_month & " closing\Tools & Reports\Input\promotion_data.xlsx"
    path_to_promotion_sg = path_to_folder & "\M006) Marketplace SG\MPS " & year_month & " closing\Tools & Reports\Input\promotion_data.xlsx"
    path_to_promotion_hk = path_to_folder & "\M007) Marketplace HK\MPH " & year_month & " closing\Tools & Reports\Input\promotion_data.xlsx"
    path_to_promotion_my = path_to_folder & "\M009) Marketplace MY\MPM " & year_month & " closing\Tools & Reports\Input\promotion_data.xlsx"
    
    
' erase blank rows disputes
    
    ' TW
    Workbooks.Open (path_to_disputes_tw)
    Workbooks(2).Activate
    Sheets("disputes").Select
    
    If Range("M5").Value <> 0 Then
    Range("M4").Select
    Selection.End(xlDown).Select
    Rz = ActiveCell.row
    
    Rows(Rz + 1 & ":" & Rz + 1).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    End If
    
    Workbooks(2).Close True, path_to_disputes_tw
    
    ' SG
    Workbooks.Open (path_to_disputes_sg)
    Workbooks(2).Activate
    Sheets("disputes").Select
    
    If Range("M5").Value <> 0 Then
    Range("M4").Select
    Selection.End(xlDown).Select
    Rz = ActiveCell.row
    
    Rows(Rz + 1 & ":" & Rz + 1).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    End If
    
    Workbooks(2).Close True, path_to_disputes_sg
    
    ' HK
    Workbooks.Open (path_to_disputes_hk)
    Workbooks(2).Activate
    Sheets("disputes").Select
    
    If Range("M5").Value <> 0 Then
    Range("M4").Select
    Selection.End(xlDown).Select
    Rz = ActiveCell.row
    
    Rows(Rz + 1 & ":" & Rz + 1).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    End If
    
    Workbooks(2).Close True, path_to_disputes_hk
    
    ' MY
    Workbooks.Open (path_to_disputes_my)
    Workbooks(2).Activate
    Sheets("disputes").Select
    
    If Range("M5").Value <> 0 Then
    Range("M4").Select
    Selection.End(xlDown).Select
    Rz = ActiveCell.row
    
    Rows(Rz + 1 & ":" & Rz + 1).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    End If
    
    Workbooks(2).Close True, path_to_disputes_my
    
' erase blank rows ap_aging
    
    ' TW
    Workbooks.Open (path_to_ap_aging_tw)
    Workbooks(2).Activate
    Sheets("ap_aging").Select
    
    If Range("M5").Value <> 0 Then
    Range("M4").Select
    Selection.End(xlDown).Select
    Rz = ActiveCell.row
    
    Rows(Rz + 1 & ":" & Rz + 1).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    End If
    
    Workbooks(2).Close True, path_to_ap_aging_tw
    
    ' SG
    Workbooks.Open (path_to_ap_aging_sg)
    Workbooks(2).Activate
    Sheets("ap_aging").Select
    
    If Range("M5").Value <> 0 Then
    Range("M4").Select
    Selection.End(xlDown).Select
    Rz = ActiveCell.row
    
    Rows(Rz + 1 & ":" & Rz + 1).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    End If
    
    Workbooks(2).Close True, path_to_ap_aging_sg

    ' HK
    Workbooks.Open (path_to_ap_aging_hk)
    Workbooks(2).Activate
    Sheets("ap_aging").Select
    
    If Range("M5").Value <> 0 Then
    Range("M4").Select
    Selection.End(xlDown).Select
    Rz = ActiveCell.row
    
    Rows(Rz + 1 & ":" & Rz + 1).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    End If
    
    Workbooks(2).Close True, path_to_ap_aging_hk
    
    
    ' MY
    Workbooks.Open (path_to_ap_aging_my)
    Workbooks(2).Activate
    Sheets("ap_aging").Select
    
    If Range("M5").Value <> 0 Then
    Range("M4").Select
    Selection.End(xlDown).Select
    Rz = ActiveCell.row
    
    Rows(Rz + 1 & ":" & Rz + 1).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    End If
    
    Workbooks(2).Close True, path_to_ap_aging_my

' erase blank rows promotion_data
    
    ' TW
    Workbooks.Open (path_to_promotion_tw)
    Workbooks(2).Activate
    Sheets("promotion_data").Select
    
    If Range("A2").Value <> 0 Then
    Range("A1").Select
    Selection.End(xlDown).Select
    Rz = ActiveCell.row
    
    Rows(Rz + 1 & ":" & Rz + 1).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    End If
    
    Workbooks(2).Close True, path_to_promotion_tw
    
    ' SG
    Workbooks.Open (path_to_promotion_sg)
    Workbooks(2).Activate
    Sheets("promotion_data").Select
    
    If Range("A2").Value <> 0 Then
    Range("A1").Select
    Selection.End(xlDown).Select
    Rz = ActiveCell.row
    
    Rows(Rz + 1 & ":" & Rz + 1).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    End If
    
    Workbooks(2).Close True, path_to_promotion_sg
    
    ' HK
    Workbooks.Open (path_to_promotion_hk)
    Workbooks(2).Activate
    Sheets("promotion_data").Select
    
    If Range("A2").Value <> 0 Then
    Range("A1").Select
    Selection.End(xlDown).Select
    Rz = ActiveCell.row
    
    Rows(Rz + 1 & ":" & Rz + 1).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    End If
    
    Workbooks(2).Close True, path_to_promotion_hk
    
    ' MY
    Workbooks.Open (path_to_promotion_my)
    Workbooks(2).Activate
    Sheets("promotion_data").Select
    
    If Range("A2").Value <> 0 Then
    Range("A1").Select
    Selection.End(xlDown).Select
    Rz = ActiveCell.row
    
    Rows(Rz + 1 & ":" & Rz + 1).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    End If
    
    Workbooks(2).Close True, path_to_promotion_my

End Sub

Attribute VB_Name = "zz_populate_empty_template"
Sub populate_empty_template()

    path_to_folder = Sheets("Automatic PDF Generation").Range("C2").Value
    year_month = Sheets("Automatic PDF Generation").Range("C3").Value
    
    path_from_disputes = path_to_folder & "\MP Generation Tool\empty_templates\disputes.xlsx"
    path_to_disputes_tw = path_to_folder & "\M005) Marketplace TW\MPT " & year_month & " closing\Tools & Reports\Input\disputes.xlsx"
    path_to_disputes_sg = path_to_folder & "\M006) Marketplace SG\MPS " & year_month & " closing\Tools & Reports\Input\disputes.xlsx"
    path_to_disputes_hk = path_to_folder & "\M007) Marketplace HK\MPH " & year_month & " closing\Tools & Reports\Input\disputes.xlsx"
    path_to_disputes_my = path_to_folder & "\M009) Marketplace MY\MPM " & year_month & " closing\Tools & Reports\Input\disputes.xlsx"
    
    path_from_ap_aging = path_to_folder & "\MP Generation Tool\empty_templates\ap_aging.xlsx"
    path_to_ap_aging_tw = path_to_folder & "\M005) Marketplace TW\MPT " & year_month & " closing\Tools & Reports\Input\ap_aging.xlsx"
    path_to_ap_aging_sg = path_to_folder & "\M006) Marketplace SG\MPS " & year_month & " closing\Tools & Reports\Input\ap_aging.xlsx"
    path_to_ap_aging_hk = path_to_folder & "\M007) Marketplace HK\MPH " & year_month & " closing\Tools & Reports\Input\ap_aging.xlsx"
    path_to_ap_aging_my = path_to_folder & "\M009) Marketplace MY\MPM " & year_month & " closing\Tools & Reports\Input\ap_aging.xlsx"
    
    path_from_promotion = path_to_folder & "\MP Generation Tool\empty_templates\promotion_data.xlsx"
    path_to_promotion_tw = path_to_folder & "\M005) Marketplace TW\MPT " & year_month & " closing\Tools & Reports\Input\promotion_data.xlsx"
    path_to_promotion_sg = path_to_folder & "\M006) Marketplace SG\MPS " & year_month & " closing\Tools & Reports\Input\promotion_data.xlsx"
    path_to_promotion_hk = path_to_folder & "\M007) Marketplace HK\MPH " & year_month & " closing\Tools & Reports\Input\promotion_data.xlsx"
    path_to_promotion_my = path_to_folder & "\M009) Marketplace MY\MPM " & year_month & " closing\Tools & Reports\Input\promotion_data.xlsx"
    
    
    ' copy disputes empty templates
    
    If Dir(path_to_disputes_tw) = "" Then
    FileCopy path_from_disputes, path_to_disputes_tw
    End If
    
    If Dir(path_to_disputes_sg) = "" Then
    FileCopy path_from_disputes, path_to_disputes_sg
    End If
    
    If Dir(path_to_disputes_hk) = "" Then
    FileCopy path_from_disputes, path_to_disputes_hk
    End If
    
    If Dir(path_to_disputes_my) = "" Then
    FileCopy path_from_disputes, path_to_disputes_my
    End If
    
    ' copy ap_aging empty templates
    
    If Dir(path_to_ap_aging_tw) = "" Then
    FileCopy path_from_ap_aging, path_to_ap_aging_tw
    End If
    
    If Dir(path_to_ap_aging_sg) = "" Then
    FileCopy path_from_ap_aging, path_to_ap_aging_sg
    End If
    
    If Dir(path_to_ap_aging_hk) = "" Then
    FileCopy path_from_ap_aging, path_to_ap_aging_hk
    End If
    
    If Dir(path_to_ap_aging_my) = "" Then
    FileCopy path_from_ap_aging, path_to_ap_aging_my
    End If
    
    ' copy promotion_data empty templates
    
    If Dir(path_to_promotion_tw) = "" Then
    FileCopy path_from_promotion, path_to_promotion_tw
    End If
    
    If Dir(path_to_promotion_sg) = "" Then
    FileCopy path_from_promotion, path_to_promotion_sg
    End If

    If Dir(path_to_promotion_hk) = "" Then
    FileCopy path_from_promotion, path_to_promotion_hk
    End If
    
    If Dir(path_to_promotion_my) = "" Then
    FileCopy path_from_promotion, path_to_promotion_my
    End If
    
End Sub
 

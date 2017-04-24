Attribute VB_Name = "l_endmacro_legalized"
Sub endmacro(location As String)
    Application.ScreenUpdating = True
    Sheets("Finance overview by seller").Activate       'Reporting sheet
    MsgBox "Generation Complete. Files are located in " & location
End Sub

Function legalized(filename As String) As String

    Dim i As Integer
    Const illegals = "\/:*?""<>|"
    
    legalized = filename
    
    For i = 1 To Len(illegals)
        legalized = Replace(legalized, Mid$(illegals, i, 1), "0")
    Next i

End Function

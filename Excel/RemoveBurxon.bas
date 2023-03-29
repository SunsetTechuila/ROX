Attribute VB_Name = "RemoveBurxon"
Sub RmBurxon()
Attribute RmBurxon.VB_Description = "Убирает Бурхона из накладных"
Attribute RmBurxon.VB_ProcData.VB_Invoke_Func = "B\n14"
'
' RemoveBurxon Макрос
' Убирает Бурхона из накладных
'
' Сочетание клавиш: Ctrl+Shift+B
'
Dim cell As Range
Dim firstAddress As String

savedCalcMode = Application.Calculation
Application.Calculation = xlCalculationManual

With Worksheets("Кол-во единица").Range("F:H")
    Set cell = .Find("ТП: Бурхон")
    If Not cell Is Nothing Then
        firstAddress = cell.Address
        
        Do
            cell.Value2 = vbNullString
            Set cell = Range(cell.Address).Offset(1)
            cell.Value2 = vbNullString
            Set cell = .FindNext(cell)
            If cell Is Nothing Then Exit Do
        Loop While cell.Address <> fistAddress
    End If
End With

Application.Calculation = savedCalcMode
End Sub

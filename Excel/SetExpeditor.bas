Attribute VB_Name = "SetExpeditor"
Sub SetExp()
Attribute SetExp.VB_Description = "Проставляет на всех накладных введенного экспедитора"
Attribute SetExp.VB_ProcData.VB_Invoke_Func = "E\n14"
'
' SetExpeditor Макрос
' Проставляет на всех накладных имя введенного экспедитора
'
' Сочетание клавиш: Ctrl+Shift+E
'
Dim cell As Range
Dim firstAddress As String, expeditorName As String

savedCalcMode = Application.Calculation
Application.Calculation = xlCalculationManual

With Worksheets("Кол-во единица").Range("C:F")
    Set cell = .Find("Накладная №")
    If Not cell Is Nothing Then
        firstAddress = cell.Address
        expeditorName = InputBox("Введите имя экспедитора")
        If expeditorName <> vbNullString Then
            expeditorName = UCase(Left(expeditorName, 1)) & Right(LCase(expeditorName), Len(expeditorName) - 1)
            Do
                Range(cell.Address).Offset(0, 3).Value2 = "Экспедитор: " & expeditorName
                Set cell = .FindNext(cell)
                If cell Is Nothing Then Exit Do
            Loop While cell.Address <> firstAddress
        End If
    End If
End With

Application.Calculation = savedCalcMode
End Sub

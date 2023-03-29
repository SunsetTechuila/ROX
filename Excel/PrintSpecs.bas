Attribute VB_Name = "PrintSpecs"
Type Order
startAddress As String
endAddress As String
End Type

Sub PrintSpec()
Attribute PrintSpec.VB_Description = "Печатает накладные в указанном количестве"
Attribute PrintSpec.VB_ProcData.VB_Invoke_Func = "W\n14"
'
' PrintSpecs Макрос
' Печатает накладные в указанном количестве
'
' Сочетание клавиш: Ctrl+Shift+W
'
Dim firstAddress As String, inpStr As String
Dim orders() As Order
Dim cop As Integer
Dim cellOne As Range, cellTwo As Range
Dim savedCalcMode As XlCalculation
Dim i As Long

savedCalcMode = Application.Calculation
Application.Calculation = xlCalculationManual

With Worksheets("Кол-во единица").Range("A:H")
    Set cellOne = .Find("Накладная")
    If Not cellOne Is Nothing Then
        inpStr = InputBox("Введите количество копий")
        If inpStr = vbNullString Then
            Application.Calculation = savedCalcMode
            Exit Sub
        Else
            cop = inpStr
        End If
        
        firstAddress = cellOne.Address
        ReDim orders(0)
        Set cellTwo = Range("A1")
        
        Do
            i = UBound(orders)
            orders(i).startAddress = Replace(cellOne.Address, "C", "A")
            Set cellTwo = .Find("Принял: ____________________________", After:=cellTwo)
            orders(i).endAddress = Replace(cellTwo.Address, "E", "H")
            Set cellOne = .Find("Накладная №", After:=cellOne)
            If cellOne.Address = firstAddress Then
                Exit Do
            Else
                ReDim Preserve orders(UBound(orders) + 1)
            End If
        Loop
        
        For i = 0 To UBound(orders)
            Range(orders(i).startAddress & ":" & orders(i).endAddress).PrintOut Copies:=cop
        Next
    End If
End With

Application.Calculation = savedCalcMode

End Sub

Attribute VB_Name = "CopyExpeditorsAndNumbers"
Type Order
    NumCell As Range
    Agent As String
    Client As String
    ClientVar2 As String
    ClientVar3 As String
    Amount As Long
    Quantity As Long
    ExpCell As Range
End Type

Type ExportedOrder
    Number As String
    Agent As String
    Client As String
    Amount As Double
    Quantity As Long
    ExpName As String
    Balance As Long
End Type

Sub CopyExpAndNum()
'
' CopyExpeditorAndNumber Макрос
' Копирует экспедиторов и номера заказов из файла экспорта, сопоставляя данные заказов
'
Dim ordersWb As Workbook, exportWb As Workbook
Dim ordersWs As Worksheet, exportWs As Worksheet
Dim i As Long, a As Long
Dim firstAddress As String, clientName1 As String, clientName2 As String
Dim cellOrder As Range, cellAmount As Range, cell As Range
Dim posINN As Integer
Dim exportOrders() As ExportedOrder
Dim orders() As Order
i = 0
a = 0
ReDim exportOrders(0)
ReDim orders(0)

For Each wb In Workbooks
    If wb.Name Like "Zagruz*" Then
        Set ordersWb = wb
        Set ordersWs = ordersWb.Worksheets("Кол-во единица")
        i = i + 1
        If i > 1 Then
            MsgBox "Файл с накладными должен быть открыт в единственном экземпляре!"
            Exit Sub
        End If
    ElseIf wb.Name Like "Data export*" Then
        Set exportWb = wb
        Set exportWs = exportWb.Worksheets("Sheet1")
        a = a + 1
        If a > 1 Then
            MsgBox "Файл экспорта должен быть открыт в единственном экземпляре!"
            Exit Sub
        End If
    End If
Next
If i < 1 Then
    MsgBox "Откройте файл с накладными!"
    Exit Sub
ElseIf a < 1 Then
    MsgBox "Откройте файл экспорта!"
    Exit Sub
End If

exportWs.Activate
Set cell = Range("B3")
Do
    exportOrders(UBound(exportOrders)).Number = cell.Value
    If exportOrders(UBound(exportOrders)).Number = "" Then
        If UBound(exportOrders) = 0 Then
            Exit Sub
        Else
            ReDim Preserve exportOrders(UBound(exportOrders) - 1)
            Exit Do
        End If
    End If
    Set cell = Range(cell.Address).Offset(0, 4)
    exportOrders(UBound(exportOrders)).Client = Trim$(cell.Value)
    posINN = InStr(exportOrders(UBound(exportOrders)).Client, "ИНН:")
    If posINN <> 0 Then exportOrders(UBound(exportOrders)).Client = Trim$(Mid(exportOrders(UBound(exportOrders)).Client, 1, posINN - 1))
    Set cell = Range(cell.Address).Offset(0, 1)
    exportOrders(UBound(exportOrders)).Quantity = cell.Value
    Set cell = Range(cell.Address).Offset(0, 1)
    exportOrders(UBound(exportOrders)).Amount = GetAmount(cell.Value)
    Set cell = Range(cell.Address).Offset(0, 4)
    exportOrders(UBound(exportOrders)).Agent = cell.Value
    Set cell = Range(cell.Address).Offset(0, 1)
    exportOrders(UBound(exportOrders)).ExpName = cell.Value
    Set cell = Range(cell.Address).Offset(0, 3)
    exportOrders(UBound(exportOrders)).Balance = cell.Value
    
    For i = (UBound(exportOrders) - 1) To 0 Step -1
        If exportOrders(UBound(exportOrders)).Client = exportOrders(i).Client And exportOrders(UBound(exportOrders)).Balance = exportOrders(i).Balance Then
            exportOrders(i).Amount = exportOrders(i).Amount + exportOrders(UBound(exportOrders)).Amount
            exportOrders(i).Quantity = exportOrders(i).Quantity + exportOrders(UBound(exportOrders)).Quantity
            exportOrders(i).Number = exportOrders(i).Number & " + " & exportOrders(UBound(exportOrders)).Number
            If exportOrders(i).ExpName <> exportOrders(UBound(exportOrders)).ExpName Then
                If exportOrders(i).ExpName = "" Then
                    exportOrders(i).ExpName = exportOrders(i).ExpName & "Без экспедитора/" & exportOrders(UBound(exportOrders)).ExpName
                ElseIf exportOrders(UBound(exportOrders)).ExpName = "" Then
                    exportOrders(i).ExpName = exportOrders(i).ExpName & "/Без экспедитора"
                Else
                    exportOrders(i).ExpName = exportOrders(i).ExpName & "/" & exportOrders(UBound(exportOrders)).ExpName
                End If
            End If
            ReDim Preserve exportOrders(UBound(exportOrders) - 1)
            Exit For
        End If
    Next
    
    Set cell = Range(cell.Address).Offset(1, -14)
    ReDim Preserve exportOrders(UBound(exportOrders) + 1)
Loop

ordersWs.Activate
With ordersWs.Range("A:H")
    Set cellOrder = .Find("Накладная")
    If Not cellOrder Is Nothing Then
        firstAddress = cellOrder.Address
        Set cellAmount = Range("A1")
        Do
            Set orders(UBound(orders)).NumCell = cellOrder
            Set cell = Range(cellOrder.Address).Offset(1, -1)
            clientName1 = Trim$(Split(Replace(cell.Value, "Кому: ", ""), " - ")(0))
            clientName2 = Split(Replace(cell.Value, "Кому: ", ""), " - ")(1)
            clientName2 = Trim$(Replace(Replace(clientName2, "(", ""), ")", ""))
            orders(UBound(orders)).Client = clientName1
            If clientName2 <> "" Then
                orders(UBound(orders)).ClientVar2 = clientName1 & " " & clientName2
                orders(UBound(orders)).ClientVar3 = clientName1 & clientName2
            End If
            Set cell = Range(cellOrder.Address).Offset(1, 3)
            orders(UBound(orders)).Agent = Replace(cell.Value, "ТП: ", "")
            Set orders(UBound(orders)).ExpCell = Range(cellOrder.Address).Offset(0, 3)
            Set cellAmount = .Find("Принял: ____________________________", After:=cellAmount)
            Set cell = Range(cellAmount.Address).Offset(-1)
            Set cell = Range(Replace(cell.Address, "E", "H"))
            orders(UBound(orders)).Amount = Replace(Replace(cell.Value, " сум", ""), ",", "")
            Set cell = Range(cell.Address).Offset(0, -3)
            If cell.Value = "" Then Set cell = Range(cell.Address).Offset(-3)
            orders(UBound(orders)).Quantity = cell.Value
            Set cellOrder = .Find("Накладная", After:=cellOrder)
            If cellOrder.Address = firstAddress Then
                Exit Do
            Else: ReDim Preserve orders(UBound(orders) + 1)
            End If
        Loop
    Else: Exit Sub
    End If
End With

For i = UBound(orders) To 0 Step -1
    For a = 0 To UBound(exportOrders)
        Select Case False
            Case orders(i).Agent = exportOrders(a).Agent
            Case orders(i).Client = exportOrders(a).Client Or orders(i).ClientVar2 = exportOrders(a).Client Or orders(i).ClientVar3 = exportOrders(a).Client
            Case orders(i).Amount = CLng(exportOrders(a).Amount)
            Case orders(i).Quantity = exportOrders(a).Quantity
        Case Else
            If exportOrders(a).ExpName <> "" Then orders(i).ExpCell.Value = "Экспедитор: " & exportOrders(a).ExpName
            orders(i).NumCell.Value = Replace(orders(i).NumCell.Value, "№", "№" & exportOrders(a).Number)
            Exit For
        End Select
    Next
Next

End Sub

Function GetAmount(amountString As String) As Double
Dim commaPos As Integer, dotPos As Integer

commaPos = InStr(amountString, ",")
dotPos = InStr(amountString, ".")

If commaPos = 0 And dotPos = 0 Then
    GetAmount = amountString
ElseIf commaPos <> 0 And dotPos <> 0 Then
    GetAmount = Replace(Replace(amountString, ",", ""), ".", ",")
Else
    If Len(amountString) - commaPos > 2 Then
        GetAmount = Replace(amountString, ",", "")
    Else
        GetAmount = amountString
    End If
End If

End Function

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
    ExpArr() As String
    ExpStr As String
    Balance As Long
End Type

Sub CopyExpAndNum()
'
' CopyExpeditorAndNumber Макрос
' Копирует экспедиторов и номера заказов из файла экспорта, сопоставляя данные заказов
'
Dim ordersWs As Worksheet, exportWs As Worksheet
Dim i As Long, a As Long, c As Long, b As Long
Dim firstAddress As String, clientName1 As String, clientName2 As String
Dim cellOrder As Range, cellAmount As Range, cell As Range
Dim posINN As Integer
Dim exportOrders() As ExportedOrder
Dim orders() As Order
a = 0
i = 0
ReDim exportOrders(0)
ReDim exportOrders(0).ExpArr(0)
ReDim orders(0)

For Each wb In Workbooks
    If wb.Name Like "Zagruz*" Then
        i = i + 1
        If i > 1 Then
            MsgBox "Файл с накладными должен быть открыт в единственном экземпляре!"
            Exit Sub
        End If
        Set ordersWs = wb.Worksheets("Кол-во единица")
    ElseIf wb.Name Like "Data export*" Then
        a = a + 1
        If a > 1 Then
            MsgBox "Файл экспорта должен быть открыт в единственном экземпляре!"
            Exit Sub
        End If
        Set exportWs = wb.Worksheets("Sheet1")
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
    i = UBound(exportOrders)

    exportOrders(i).Number = cell.Value2
    If exportOrders(i).Number = vbNullString Then
        If i = 0 Then
            Exit Sub
        Else
            ReDim Preserve exportOrders(i - 1)
            Exit Do
        End If
    End If

    Set cell = Range(cell.Address).Offset(0, 4)
    exportOrders(i).Client = Trim$(cell.Value2)
    posINN = InStr(exportOrders(i).Client, "ИНН:")
    If posINN <> 0 Then exportOrders(i).Client = Trim$(Mid(exportOrders(i).Client, 1, posINN - 1))
    Set cell = Range(cell.Address).Offset(0, 1)
    exportOrders(i).Quantity = cell.Value2
    Set cell = Range(cell.Address).Offset(0, 1)
    exportOrders(i).Amount = GetAmount(cell.Value2)
    Set cell = Range(cell.Address).Offset(0, 4)
    exportOrders(i).Agent = cell.Value2
    Set cell = Range(cell.Address).Offset(0, 1)
    exportOrders(i).ExpArr() = cell.Value2
    Set cell = Range(cell.Address).Offset(0, 3)
    exportOrders(i).Balance = cell.Value2
    
    For a = (i - 1) To 0 Step -1
        If exportOrders(i).Client = exportOrders(a).Client And exportOrders(i).Balance = exportOrders(a).Balance Then
            exportOrders(a).Amount = exportOrders(a).Amount + exportOrders(i).Amount
            exportOrders(a).Quantity = exportOrders(a).Quantity + exportOrders(i).Quantity
            exportOrders(a).Number = exportOrders(a).Number & "+" & exportOrders(i).Number
            exportOrders(a).ExpArr = 
            ReDim Preserve exportOrders(i - 1)
            Exit For
        End If
    Next
    
    Set cell = Range(cell.Address).Offset(1, -14)
    ReDim Preserve exportOrders(i + 1)
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
            clientName1 = Trim$(Split(Replace(cell.Value2, "Кому: ", ""), " - ")(0))
            clientName2 = Split(Replace(cell.Value2, "Кому: ", ""), " - ")(1)
            clientName2 = Trim$(Replace(Replace(clientName2, "(", ""), ")", ""))
            orders(UBound(orders)).Client = clientName1
            If clientName2 <> "" Then
                orders(UBound(orders)).ClientVar2 = clientName1 & " " & clientName2
                orders(UBound(orders)).ClientVar3 = clientName1 & clientName2
            End If
            Set cell = Range(cellOrder.Address).Offset(1, 3)
            orders(UBound(orders)).Agent = Replace(cell.Value2, "ТП: ", "")
            Set orders(UBound(orders)).ExpCell = Range(cellOrder.Address).Offset(0, 3)
            Set cellAmount = .Find("Принял: ____________________________", After:=cellAmount)
            Set cell = Range(cellAmount.Address).Offset(-1)
            Set cell = Range(Replace(cell.Address, "E", "H"))
            orders(UBound(orders)).Amount = Replace(Replace(cell.Value2, " сум", ""), ",", "")
            Set cell = Range(cell.Address).Offset(0, -3)
            If cell.Value2 = "" Then Set cell = Range(cell.Address).Offset(-3)
            orders(UBound(orders)).Quantity = cell.Value2
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
            If exportOrders(a).ExpName <> "" Then orders(i).ExpCell.Value2 = "Экспедитор: " & exportOrders(a).ExpName
            orders(i).NumCell.Value2 = Replace(orders(i).NumCell.Value2, "№", "№" & exportOrders(a).Number)
            Exit For
        End Select
    Next
Next

End Sub

Function GetAmount(amountStr As String) As Double
Dim commaPos As Integer, dotPos As Integer

commaPos = InStr(amountStr, ",")
dotPos = InStr(amountStr, ".")

If commaPos = 0 And dotPos = 0 Then
    GetAmount = amountStr
ElseIf commaPos <> 0 And dotPos <> 0 Then
    GetAmount = Replace(Replace(amountStr, ",", ""), ".", ",")
Else
    If Len(amountStr) - commaPos > 2 Then
        GetAmount = Replace(amountStr, ",", "")
    Else
        GetAmount = amountStr
    End If
End If

End Function

Function GetExpStr(arrExpName()) As String
Dim lastExpName As String

End Function

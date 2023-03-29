Attribute VB_Name = "SetExpeditor2"
Type Order
  cellExp As Range
  Agent As String
End Type

Public choiceButtonClicked As Boolean

Sub SetExp2()
Attribute SetExp2.VB_Description = "Проставляет экспедитора на основании выбранных агентов"
Attribute SetExp2.VB_ProcData.VB_Invoke_Func = "R\n14"
'
' SetExpeditor2 Макрос
' Проставляет экспедитора на основании выбранных агентов
'
Dim cellAgent As Range
Dim orders() As Order
Dim i As Long, a As Long
Dim firstAddress As String, expeditorName As String, arrAgentName() As String
Dim found As Boolean

With Worksheets("Кол-во единица").Range("F:H")
  Set cellAgent = .Find("ТП: ")
  If Not cellAgent Is Nothing Then
    firstAddress = cellAgent.Address
    ReDim arrAgentName(0)
    ReDim orders(0)
    Set chooseForm = AgentChooseForm
    Set agentList = chooseForm.AgentListBox
    choiceButtonClicked = False
    agentList.Clear
    
    expeditorName = InputBox("Введите имя экспедитора")
    If expeditorName = vbNullString Then Exit Sub
    expeditorName = UCase(Left(expeditorName, 1)) & Right(LCase(expeditorName), Len(expeditorName) - 1)
    
    Do
      Set orders(UBound(orders)).cellExp = Range(cellAgent.Address).Offset(-1)
      orders(UBound(orders)).Agent = Replace(cellAgent.Value, "ТП: ", vbNullString)
      
      found = False
      For i = 0 To agentList.ListCount - 1
        If StrComp(orders(UBound(orders)).Agent, agentList.List(i), vbTextCompare) = 0 Then
          found = True
          Exit For
        End If
      Next
      If found = False Then agentList.AddItem orders(UBound(orders)).Agent
      
      Set cellAgent = .FindNext(cellAgent)
      If cellAgent Is Nothing Then Exit Do
      ReDim Preserve orders(UBound(orders) + 1)
    Loop While cellAgent.Address <> firstAddress
    If orders(UBound(orders)).Agent = vbNullString Then ReDim Preserve orders(UBound(orders) - 1)
    
    chooseForm.Show
    If choiceButtonClicked = False Then Exit Sub
    
    For i = 0 To agentList.ListCount - 1
      If agentList.Selected(i) Then
        arrAgentName(UBound(arrAgentName)) = agentList.List(i)
        ReDim Preserve arrAgentName(UBound(arrAgentName) + 1)
      End If
    Next
    If arrAgentName(0) = vbNullString Then Exit Sub
    If arrAgentName(UBound(arrAgentName)) = vbNullString Then ReDim Preserve arrAgentName(UBound(arrAgentName) - 1)
    
    For i = 0 To UBound(orders)
      For a = 0 To UBound(arrAgentName)
        If orders(i).Agent = arrAgentName(a) Then orders(i).cellExp.Value = "Экспедитор: " & expeditorName
      Next
    Next
    
  End If
End With
End Sub

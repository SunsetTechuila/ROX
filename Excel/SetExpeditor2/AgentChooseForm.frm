VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AgentChooseForm 
   Caption         =   "Выбор агентов"
   ClientHeight    =   5450
   ClientLeft      =   50
   ClientTop       =   390
   ClientWidth     =   4000
   OleObjectBlob   =   "AgentChooseForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AgentChooseForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ChoiceButton_Click()
SetExpeditor2.choiceButtonClicked = True
AgentChooseForm.Hide
End Sub

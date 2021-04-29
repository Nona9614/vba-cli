VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
End
Attribute VB_Name = "Class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private mValue As String

Public Function Init(Value as String)
  mValue = Value
End Function

Public Property Get ClassValue() As IRibbonUI
  ClassValue = mValue
End Property

Public Property Let ClassValue(Value As IRibbonUI)
  mValue = Value
End Property

Public Sub ClassFunction()
  MsgBox "Welcome!", vbButtonType, "VBA CLI"
End Sub
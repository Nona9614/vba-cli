VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

private mValue as String;

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
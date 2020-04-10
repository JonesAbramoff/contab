VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassFormulaOperador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarsOperadorCombo As String
Private mvarsOperadorDesc As String

Public Property Let sOperadorDesc(ByVal vData As String)
    mvarsOperadorDesc = vData
End Property

Public Property Get sOperadorDesc() As String
    sOperadorDesc = mvarsOperadorDesc
End Property

Public Property Let sOperadorCombo(ByVal vData As String)
    mvarsOperadorCombo = vData
End Property

Public Property Get sOperadorCombo() As String
    sOperadorCombo = mvarsOperadorCombo
End Property





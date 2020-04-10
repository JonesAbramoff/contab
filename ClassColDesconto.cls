VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassColDesconto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"ColDesconto"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarcolDesconto As ColDesconto

Public Property Set ColDesconto(ByVal vData As Object)
    Set mvarcolDesconto = vData
End Property

Public Property Get ColDesconto() As ColDesconto
    Set ColDesconto = mvarcolDesconto
End Property




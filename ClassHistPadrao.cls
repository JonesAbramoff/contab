VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassHistPadrao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariHistPadrao As Integer
Private mvarsDescHistPadrao As String

Public Property Let sDescHistPadrao(ByVal vData As String)
    mvarsDescHistPadrao = vData
End Property

Public Property Get sDescHistPadrao() As String
    sDescHistPadrao = mvarsDescHistPadrao
End Property

Public Property Let iHistPadrao(ByVal vData As Integer)
    mvariHistPadrao = vData
End Property

Public Property Get iHistPadrao() As Integer
    iHistPadrao = mvariHistPadrao
End Property



Private Sub Class_Initialize()
    mvarsDescHistPadrao = String(STRING_HISTORICO, 0)
End Sub

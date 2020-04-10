VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassTipoDeProdutoCategoria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariTipoDeProduto As Integer
Private mvarsCategoria As String
Private mvarsItem As String

Public Property Let sItem(ByVal vData As String)
    mvarsItem = vData
End Property

Public Property Get sItem() As String
    sItem = mvarsItem
End Property

Public Property Let sCategoria(ByVal vData As String)
    mvarsCategoria = vData
End Property

Public Property Get sCategoria() As String
    sCategoria = mvarsCategoria
End Property

Public Property Let iTipoDeProduto(ByVal vData As Integer)
    mvariTipoDeProduto = vData
End Property

Public Property Get iTipoDeProduto() As Integer
    iTipoDeProduto = mvariTipoDeProduto
End Property




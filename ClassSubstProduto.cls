VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassSubstProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"colItemPedido"
Option Explicit

Private mvarsCodProdutoSubstituto As String
Private mvarsCodProduto As String
Private mvarcolItemPedido As New colItemPedido

Public Property Set colItemPedido(ByVal vData As Object)
    Set mvarcolItemPedido = vData
End Property

Public Property Get colItemPedido() As colItemPedido
    Set colItemPedido = mvarcolItemPedido
End Property

Public Property Let sCodProduto(ByVal vData As String)
    mvarsCodProduto = vData
End Property

Public Property Get sCodProduto() As String
    sCodProduto = mvarsCodProduto
End Property

Public Property Let sCodProdutoSubstituto(ByVal vData As String)
    mvarsCodProdutoSubstituto = vData
End Property

Public Property Get sCodProdutoSubstituto() As String
    sCodProdutoSubstituto = mvarsCodProdutoSubstituto
End Property




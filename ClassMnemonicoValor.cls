VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassMnemonicoValor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Contém o(s) valor(es) do mnemonico"
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarsMnemonico As String
Private mvariParam As Integer
Private mvarvParam(1 To 3) As Variant
Private mvarcolValor As Collection

Public Property Set colValor(ByVal vData As Object)
Attribute colValor.VB_Description = "coleção contendo os valores do mnenmonico."
    Set mvarcolValor = vData
End Property

Public Property Get colValor() As Collection
    Set colValor = mvarcolValor
End Property

Public Property Let vParam(ByVal iIndice As Integer, ByVal vData As Variant)
Attribute vParam.VB_Description = "Valor do Parametro 3"
    mvarvParam(iIndice) = vData
End Property

Public Property Set vParam(ByVal iIndice As Integer, ByVal vData As Object)
    Set mvarvParam(iIndice) = vData
End Property

Public Property Get vParam(ByVal iIndice As Integer) As Variant
    If IsObject(mvarvParam(iIndice)) Then
        Set vParam = mvarvParam(iIndice)
    Else
        vParam = mvarvParam(iIndice)
    End If
End Property

Public Property Let iParam(ByVal vData As Integer)
Attribute iParam.VB_Description = "Número de Parametros"
    mvariParam = vData
End Property

Public Property Get iParam() As Integer
    iParam = mvariParam
End Property

Public Property Let sMnemonico(ByVal vData As String)
Attribute sMnemonico.VB_Description = "Nome do Mnemonico"
    mvarsMnemonico = vData
End Property

Public Property Get sMnemonico() As String
    sMnemonico = mvarsMnemonico
End Property




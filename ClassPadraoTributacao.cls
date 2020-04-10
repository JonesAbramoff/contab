VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassPadraoTribSaida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarsNaturezaOperacao As String
Private mvarsSiglaMovto As String
Private mvarsCategoriaFilialCliente As String
Private mvarsItemCategoria As String
Private mvariTipoTributacaoPadrao As Integer

Public Property Let iTipoTributacaoPadrao(ByVal vData As Integer)
    mvariTipoTributacaoPadrao = vData
End Property

Public Property Get iTipoTributacaoPadrao() As Integer
    iTipoTributacaoPadrao = mvariTipoTributacaoPadrao
End Property

Public Property Let sItemCategoria(ByVal vData As String)
    mvarsItemCategoria = vData
End Property

Public Property Get sItemCategoria() As String
    sItemCategoria = mvarsItemCategoria
End Property

Public Property Let sCategoriaFilialCliente(ByVal vData As String)
    mvarsCategoriaFilialCliente = vData
End Property

Public Property Get sCategoriaFilialCliente() As String
    sCategoriaFilialCliente = mvarsCategoriaFilialCliente
End Property

Public Property Let sSiglaMovto(ByVal vData As String)
    mvarsSiglaMovto = vData
End Property

Public Property Get sSiglaMovto() As String
    sSiglaMovto = mvarsSiglaMovto
End Property

Public Property Let sNaturezaOperacao(ByVal vData As String)
    mvarsNaturezaOperacao = vData
End Property

Public Property Get sNaturezaOperacao() As String
    sNaturezaOperacao = mvarsNaturezaOperacao
End Property




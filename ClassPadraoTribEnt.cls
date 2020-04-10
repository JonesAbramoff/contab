VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassPadraoTribEnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarsNaturezaOperacao As String
Private mvarsSiglaMovto As String
Private mvarsCategoriaProduto As String
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

Public Property Let sCategoriaProduto(ByVal vData As String)
    mvarsCategoriaProduto = vData
End Property

Public Property Get sCategoriaProduto() As String
    sCategoriaProduto = mvarsCategoriaProduto
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





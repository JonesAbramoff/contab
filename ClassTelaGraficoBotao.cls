VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTelaGraficoBotao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarsTextoExibicao As String
Private mvarsNomeFuncao As String
Private mvarcolParametros As New Collection
Private mvarsNome As String
Private mvariAtualizaRetornoClick As Integer

Public Property Let iAtualizaRetornoClick(ByVal vData As Integer)
    mvariAtualizaRetornoClick = vData
End Property

Public Property Get iAtualizaRetornoClick() As Integer
    iAtualizaRetornoClick = mvariAtualizaRetornoClick
End Property

Public Property Set colParametros(ByVal vData As Collection)
    Set mvarcolParametros = vData
End Property

Public Property Get colParametros() As Collection
    Set colParametros = mvarcolParametros
End Property

Public Property Let sTextoExibicao(ByVal vData As String)
    mvarsTextoExibicao = vData
End Property

Public Property Get sTextoExibicao() As String
    sTextoExibicao = mvarsTextoExibicao
End Property

Public Property Let sNomeFuncao(ByVal vData As String)
    mvarsNomeFuncao = vData
End Property

Public Property Get sNomeFuncao() As String
    sNomeFuncao = mvarsNomeFuncao
End Property

Public Property Let sNome(ByVal vData As String)
    mvarsNome = vData
End Property

Public Property Get sNome() As String
    sNome = mvarsNome
End Property

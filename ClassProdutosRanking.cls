VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassProdutosRanking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvariNumIntRel As Integer
Private mvarlRanking As Long
Private mvarsProduto As String
Private mvarsNomeReduzido As String
Private mvardQuantidade As Double
Private mvardValor As Double
Private mvardVlrPercParticip As Double
Private mvardQuantPercParticip As Double
Private mvarsItemCategoria As String

Public Property Let iNumIntRel(ByVal vData As Integer)
    mvariNumIntRel = vData
End Property

Public Property Get iNumIntRel() As Integer
    iNumIntRel = mvariNumIntRel
End Property

Public Property Let lRanking(ByVal vData As Double)
    mvarlRanking = vData
End Property

Public Property Get lRanking() As Double
    lRanking = mvarlRanking
End Property

Public Property Let sItemCategoria(ByVal vData As String)
    mvarsItemCategoria = vData
End Property

Public Property Get sItemCategoria() As String
    sItemCategoria = mvarsItemCategoria
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let sNomeReduzido(ByVal vData As String)
    mvarsNomeReduzido = vData
End Property

Public Property Get sNomeReduzido() As String
    sNomeReduzido = mvarsNomeReduzido
End Property

Public Property Let dQuantidade(ByVal vData As Double)
    mvardQuantidade = vData
End Property

Public Property Get dQuantidade() As Double
    dQuantidade = mvardQuantidade
End Property

Public Property Let dQuantPercParticip(ByVal vData As Double)
    mvardQuantPercParticip = vData
End Property

Public Property Get dQuantPercParticip() As Double
    dQuantPercParticip = mvardQuantPercParticip
End Property

Public Property Let dVlrPercParticip(ByVal vData As Double)
    mvardVlrPercParticip = vData
End Property

Public Property Get dVlrPercParticip() As Double
    dVlrPercParticip = mvardVlrPercParticip
End Property

Public Property Let dValor(ByVal vData As Double)
    mvardValor = vData
End Property

Public Property Get dValor() As Double
    dValor = mvardValor
End Property


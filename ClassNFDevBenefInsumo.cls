VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassNFDevBenefInsumo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarsProduto As String
Private mvarsUM As String
Private mvardQuantidade As Double
Private mvarcolItemRem As New Collection 'Coleção com os itens de NF remessa que estão sendo devolvidos
Private mvarcolItemOP As New Collection 'Coleção com os itens de OP que estão sendo devolvidos
Private mvardQuantDevolvida As Double
Private mvardQuantADevolver As Double
Private mvarlNumIntItemOP As Long

Public Property Let lNumIntItemOP(ByVal vData As Long)
    mvarlNumIntItemOP = vData
End Property

Public Property Get lNumIntItemOP() As Long
    lNumIntItemOP = mvarlNumIntItemOP
End Property

Public Property Let dQuantADevolver(ByVal vData As Double)
    mvardQuantADevolver = vData
End Property

Public Property Get dQuantADevolver() As Double
    dQuantADevolver = mvardQuantADevolver
End Property

Public Property Set colItemOP(ByVal vData As Object)
    Set mvarcolItemOP = vData
End Property

Public Property Get colItemOP() As Collection
    Set colItemOP = mvarcolItemOP
End Property

Public Property Let dQuantDevolvida(ByVal vData As Double)
    mvardQuantDevolvida = vData
End Property

Public Property Get dQuantDevolvida() As Double
    dQuantDevolvida = mvardQuantDevolvida
End Property

Public Property Set colItemRem(ByVal vData As Object)
    Set mvarcolItemRem = vData
End Property

Public Property Get colItemRem() As Collection
    Set colItemRem = mvarcolItemRem
End Property

Public Property Let dQuantidade(ByVal vData As Double)
    mvardQuantidade = vData
End Property

Public Property Get dQuantidade() As Double
    dQuantidade = mvardQuantidade
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let sUM(ByVal vData As String)
    mvarsUM = vData
End Property

Public Property Get sUM() As String
    sUM = mvarsUM
End Property


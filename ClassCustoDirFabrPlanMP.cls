VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCustoDirFabrPlanMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvariFilialEmpresa As Integer
Private mvariAno As Integer
Private mvarsProduto As String
Private mvariSeq As Integer
Private mvarsProdutoMP As String
Private mvardQtde As Double
Private mvarsProdutoPai As String
Private mvarsProdutoPai1 As String

'os campos abaixo foram incluidos para facilitar a implementacao da tela MatPrim
Private mvardCustoUnitario As Double
Private mvardParticipacao As Double

Public Property Let dParticipacao(ByVal vData As Double)
    mvardParticipacao = vData
End Property

Public Property Get dParticipacao() As Double
    dParticipacao = mvardParticipacao
End Property

Public Property Let dCustoUnitario(ByVal vData As Double)
    mvardCustoUnitario = vData
End Property

Public Property Get dCustoUnitario() As Double
    dCustoUnitario = mvardCustoUnitario
End Property

Public Property Let dQtde(ByVal vData As Double)
    mvardQtde = vData
End Property

Public Property Get dQtde() As Double
    dQtde = mvardQtde
End Property

Public Property Let sProdutoMP(ByVal vData As String)
    mvarsProdutoMP = vData
End Property

Public Property Get sProdutoMP() As String
    sProdutoMP = mvarsProdutoMP
End Property

Public Property Let iSeq(ByVal vData As Integer)
    mvariSeq = vData
End Property

Public Property Get iSeq() As Integer
    iSeq = mvariSeq
End Property

Public Property Let iAno(ByVal vData As Integer)
    mvariAno = vData
End Property

Public Property Get iAno() As Integer
    iAno = mvariAno
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let sProdutoPai(ByVal vData As String)
    mvarsProdutoPai = vData
End Property

Public Property Get sProdutoPai() As String
    sProdutoPai = mvarsProdutoPai
End Property

Public Property Let sProdutoPai1(ByVal vData As String)
    mvarsProdutoPai1 = vData
End Property

Public Property Get sProdutoPai1() As String
    sProdutoPai1 = mvarsProdutoPai1
End Property




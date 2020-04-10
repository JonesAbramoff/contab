VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCustoDirFabrPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvariFilialEmpresa As Integer
Private mvariAno As Integer
Private mvarsProduto As String
Private mvariSeq As Integer
Private mvarsProdutoPai1 As String
Private mvarsProdutoPai As String
Private mvarsProdutoFilho As String
Private mvardQuantPai As Double
Private mvardParticipacao As Double
Private mvardPesoFator1 As Double
Private mvardPesoFator2 As Double
Private mvardPesoFator3 As Double
Private mvardPesoFator4 As Double
Private mvardPesoFator5 As Double
Private mvardPesoFator6 As Double

Private mvarcolMP As New Collection
Private mvarcolEmb As New Collection

Public Property Set colEmb(ByVal vData As Object)
    Set mvarcolEmb = vData
End Property

Public Property Get colEmb() As Collection
    Set colEmb = mvarcolEmb
End Property

Public Property Set colMP(ByVal vData As Object)
    Set mvarcolMP = vData
End Property

Public Property Get colMP() As Collection
    Set colMP = mvarcolMP
End Property

Public Property Let dPesoFator4(ByVal vData As Double)
    mvardPesoFator4 = vData
End Property

Public Property Get dPesoFator4() As Double
    dPesoFator4 = mvardPesoFator4
End Property

Public Property Let dPesoFator3(ByVal vData As Double)
    mvardPesoFator3 = vData
End Property

Public Property Get dPesoFator3() As Double
    dPesoFator3 = mvardPesoFator3
End Property

Public Property Let dPesoFator2(ByVal vData As Double)
    mvardPesoFator2 = vData
End Property

Public Property Get dPesoFator2() As Double
    dPesoFator2 = mvardPesoFator2
End Property

Public Property Let dPesoFator1(ByVal vData As Double)
    mvardPesoFator1 = vData
End Property

Public Property Get dPesoFator1() As Double
    dPesoFator1 = mvardPesoFator1
End Property

Public Property Let dParticipacao(ByVal vData As Double)
    mvardParticipacao = vData
End Property

Public Property Get dParticipacao() As Double
    dParticipacao = mvardParticipacao
End Property

Public Property Let dQuantPai(ByVal vData As Double)
    mvardQuantPai = vData
End Property

Public Property Get dQuantPai() As Double
    dQuantPai = mvardQuantPai
End Property

Public Property Let sProdutoFilho(ByVal vData As String)
    mvarsProdutoFilho = vData
End Property

Public Property Get sProdutoFilho() As String
    sProdutoFilho = mvarsProdutoFilho
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

Public Property Let dPesoFator5(ByVal vData As Double)
    mvardPesoFator5 = vData
End Property

Public Property Get dPesoFator5() As Double
    dPesoFator5 = mvardPesoFator5
End Property

Public Property Let dPesoFator6(ByVal vData As Double)
    mvardPesoFator6 = vData
End Property

Public Property Get dPesoFator6() As Double
    dPesoFator6 = mvardPesoFator6
End Property




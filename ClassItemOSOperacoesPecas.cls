VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassItemOSOperacoesPecas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarlNumIntDocOper As Long
Private mvarsProduto As String
Private mvardQuantidade As Double
Private mvarsUMProduto As String
Private mvariComposicao As Integer
Private mvardPercentualPerda As Double
Private mvardCustoStandard As Double
Private mvarsVersaoKitComp As String
Public Property Let sVersaoKitComp(ByVal vData As String)
    mvarsVersaoKitComp = vData
End Property

Public Property Get sVersaoKitComp() As String
    sVersaoKitComp = mvarsVersaoKitComp
End Property

Public Property Let dCustoStandard(ByVal vData As Double)
    mvardCustoStandard = vData
End Property

Public Property Get dCustoStandard() As Double
    dCustoStandard = mvardCustoStandard
End Property

Public Property Let dPercentualPerda(ByVal vData As Double)
    mvardPercentualPerda = vData
End Property

Public Property Get dPercentualPerda() As Double
    dPercentualPerda = mvardPercentualPerda
End Property

Public Property Let iComposicao(ByVal vData As Integer)
    mvariComposicao = vData
End Property

Public Property Get iComposicao() As Integer
    iComposicao = mvariComposicao
End Property

Public Property Let sUMProduto(ByVal vData As String)
    mvarsUMProduto = vData
End Property

Public Property Get sUMProduto() As String
    sUMProduto = mvarsUMProduto
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

Public Property Let lNumIntDocOper(ByVal vData As Long)
    mvarlNumIntDocOper = vData
End Property

Public Property Get lNumIntDocOper() As Long
    lNumIntDocOper = mvarlNumIntDocOper
End Property






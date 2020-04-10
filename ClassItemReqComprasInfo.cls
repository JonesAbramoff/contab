VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassItemReqComprasInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlNumIntDoc As Long
Private mvarsProduto As String
Private mvarsDescProduto As String
Private mvardQuantRecebida As Double
Private mvarsUM As String
Private mvarlReqCompra As Long
Private mvardQuantRecebidoPedido As Double
Private mvardQuantReceber As Double
Private mvariFilialEmpresa As Integer
Private mvarsFilialEmpresaNomeRed As String
Private mvariUrgente As Integer
Private mvariItemPC As Integer
Private mvariItemRC As Integer
Private mvarlPedCompra As Long
Private mvariTipoTributacao As Integer
Private mvarlNumIntDocItemPC As Long
Private mvardFator As Double
Private mvariClasseUM As Integer
Private mvarsSiglaUMCompra As String

Public Property Let sSiglaUMCompra(ByVal vData As String)
    mvarsSiglaUMCompra = vData
End Property

Public Property Get sSiglaUMCompra() As String
    sSiglaUMCompra = mvarsSiglaUMCompra
End Property

Public Property Let iClasseUM(ByVal vData As Integer)
    mvariClasseUM = vData
End Property

Public Property Get iClasseUM() As Integer
    iClasseUM = mvariClasseUM
End Property

Public Property Let dFator(ByVal vData As Double)
    mvardFator = vData
End Property

Public Property Get dFator() As Double
    dFator = mvardFator
End Property

Public Property Let lNumIntDocItemPC(ByVal vData As Long)
    mvarlNumIntDocItemPC = vData
End Property

Public Property Get lNumIntDocItemPC() As Long
    lNumIntDocItemPC = mvarlNumIntDocItemPC
End Property

Public Property Let iTipoTributacao(ByVal vData As Integer)
    mvariTipoTributacao = vData
End Property

Public Property Get iTipoTributacao() As Integer
    iTipoTributacao = mvariTipoTributacao
End Property

Public Property Let lPedCompra(ByVal vData As Long)
    mvarlPedCompra = vData
End Property

Public Property Get lPedCompra() As Long
    lPedCompra = mvarlPedCompra
End Property

Public Property Let iItemRC(ByVal vData As Integer)
    mvariItemRC = vData
End Property

Public Property Get iItemRC() As Integer
    iItemRC = mvariItemRC
End Property

Public Property Let iItemPC(ByVal vData As Integer)
    mvariItemPC = vData
End Property

Public Property Get iItemPC() As Integer
    iItemPC = mvariItemPC
End Property

Public Property Let iUrgente(ByVal vData As Integer)
    mvariUrgente = vData
End Property

Public Property Get iUrgente() As Integer
    iUrgente = mvariUrgente
End Property

Public Property Let sFilialEmpresaNomeRed(ByVal vData As String)
    mvarsFilialEmpresaNomeRed = vData
End Property

Public Property Get sFilialEmpresaNomeRed() As String
    sFilialEmpresaNomeRed = mvarsFilialEmpresaNomeRed
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let dQuantReceber(ByVal vData As Double)
    mvardQuantReceber = vData
End Property

Public Property Get dQuantReceber() As Double
    dQuantReceber = mvardQuantReceber
End Property

Public Property Let dQuantRecebidoPedido(ByVal vData As Double)
    mvardQuantRecebidoPedido = vData
End Property

Public Property Get dQuantRecebidoPedido() As Double
    dQuantRecebidoPedido = mvardQuantRecebidoPedido
End Property

Public Property Let lReqCompra(ByVal vData As Long)
    mvarlReqCompra = vData
End Property

Public Property Get lReqCompra() As Long
    lReqCompra = mvarlReqCompra
End Property

Public Property Let sUM(ByVal vData As String)
    mvarsUM = vData
End Property

Public Property Get sUM() As String
    sUM = mvarsUM
End Property

Public Property Let dQuantRecebida(ByVal vData As Double)
    mvardQuantRecebida = vData
End Property

Public Property Get dQuantRecebida() As Double
    dQuantRecebida = mvardQuantRecebida
End Property

Public Property Let sDescProduto(ByVal vData As String)
    mvarsDescProduto = vData
End Property

Public Property Get sDescProduto() As String
    sDescProduto = mvarsDescProduto
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
Attribute lNumIntDoc.VB_Description = "Num Int do ítem de Requisição de Compras"
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property




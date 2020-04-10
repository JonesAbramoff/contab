VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassDevCheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvariFilialEmpresa As Integer
Private mvarlCodigo As Long
Private mvardtData As Date
Private mvarlFornecedor As Long
Private mvariFilial As Integer
Private mvardtDataVencimento As Date
Private mvardValorCredito As Double
Private mvarlSeqChq As Long
Private mvarlNumIntChqBord As Long
Private mvarlNumIntBaixasParcRecCanc As Long
Private mvarlNumIntDoc As Long
Private mvariContaCorrente As Integer
Private mvarlNumIntTituloPag As Long
Private mvarlNumIntCheque As Long

Public Property Let lNumIntCheque(ByVal vData As Long)
    mvarlNumIntCheque = vData
End Property

Public Property Get lNumIntCheque() As Long
    lNumIntCheque = mvarlNumIntCheque
End Property

Public Property Let lNumIntTituloPag(ByVal vData As Long)
    mvarlNumIntTituloPag = vData
End Property

Public Property Get lNumIntTituloPag() As Long
    lNumIntTituloPag = mvarlNumIntTituloPag
End Property

Public Property Let iContaCorrente(ByVal vData As Integer)
    mvariContaCorrente = vData
End Property

Public Property Get iContaCorrente() As Integer
    iContaCorrente = mvariContaCorrente
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let lNumIntBaixasParcRecCanc(ByVal vData As Long)
    mvarlNumIntBaixasParcRecCanc = vData
End Property

Public Property Get lNumIntBaixasParcRecCanc() As Long
    lNumIntBaixasParcRecCanc = mvarlNumIntBaixasParcRecCanc
End Property


Public Property Let lNumIntChqBord(ByVal vData As Long)
    mvarlNumIntChqBord = vData
End Property

Public Property Get lNumIntChqBord() As Long
    lNumIntChqBord = mvarlNumIntChqBord
End Property

Public Property Let lSeqChq(ByVal vData As Long)
    mvarlSeqChq = vData
End Property

Public Property Get lSeqChq() As Long
    lSeqChq = mvarlSeqChq
End Property

Public Property Let dValorCredito(ByVal vData As Double)
    mvardValorCredito = vData
End Property

Public Property Get dValorCredito() As Double
    dValorCredito = mvardValorCredito
End Property

Public Property Let dtDataVencimento(ByVal vData As Date)
    mvardtDataVencimento = vData
End Property

Public Property Get dtDataVencimento() As Date
    dtDataVencimento = mvardtDataVencimento
End Property

Public Property Let iFilial(ByVal vData As Integer)
    mvariFilial = vData
End Property

Public Property Get iFilial() As Integer
    iFilial = mvariFilial
End Property

Public Property Let lFornecedor(ByVal vData As Long)
    mvarlFornecedor = vData
End Property

Public Property Get lFornecedor() As Long
    lFornecedor = mvarlFornecedor
End Property

Public Property Let dtData(ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let lCodigo(ByVal vData As Long)
    mvarlCodigo = vData
End Property

Public Property Get lCodigo() As Long
    lCodigo = mvarlCodigo
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
        iFilialEmpresa = mvariFilialEmpresa
End Property

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassGeracaoNFe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvariOpcaoNFiscais As Integer
Private mvarsSerieDe As String
Private mvarsSerieAte As String
Private mvarlNFiscalDe As Long
Private mvarlNFiscalAte As Long
Private mvarlClienteDe As Long
Private mvarlClienteAte As Long
Private mvarlFornecedorDe As Long
Private mvarlFornecedorAte As Long
Private mvardtEmissaoDe As Date
Private mvardtEmissaoAte As Date
Private mvarsOrdenacao As String

Public Property Let sOrdenacao(ByVal vData As String)
    mvarsOrdenacao = vData
End Property

Public Property Get sOrdenacao() As String
    sOrdenacao = mvarsOrdenacao
End Property

Public Property Let dtEmissaoAte(ByVal vData As Date)
    mvardtEmissaoAte = vData
End Property

Public Property Get dtEmissaoAte() As Date
    dtEmissaoAte = mvardtEmissaoAte
End Property

Public Property Let dtEmissaoDe(ByVal vData As Date)
    mvardtEmissaoDe = vData
End Property

Public Property Get dtEmissaoDe() As Date
    dtEmissaoDe = mvardtEmissaoDe
End Property

Public Property Let lFornecedorAte(ByVal vData As Long)
    mvarlFornecedorAte = vData
End Property

Public Property Get lFornecedorAte() As Long
    lFornecedorAte = mvarlFornecedorAte
End Property

Public Property Let lFornecedorDe(ByVal vData As Long)
    mvarlFornecedorDe = vData
End Property

Public Property Get lFornecedorDe() As Long
    lFornecedorDe = mvarlFornecedorDe
End Property

Public Property Let lClienteAte(ByVal vData As Long)
    mvarlClienteAte = vData
End Property

Public Property Get lClienteAte() As Long
    lClienteAte = mvarlClienteAte
End Property

Public Property Let lClienteDe(ByVal vData As Long)
    mvarlClienteDe = vData
End Property

Public Property Get lClienteDe() As Long
    lClienteDe = mvarlClienteDe
End Property

Public Property Let lNFiscalAte(ByVal vData As Long)
    mvarlNFiscalAte = vData
End Property

Public Property Get lNFiscalAte() As Long
    lNFiscalAte = mvarlNFiscalAte
End Property

Public Property Let lNFiscalDe(ByVal vData As Long)
    mvarlNFiscalDe = vData
End Property

Public Property Get lNFiscalDe() As Long
    lNFiscalDe = mvarlNFiscalDe
End Property

Public Property Let sSerieAte(ByVal vData As String)
    mvarsSerieAte = vData
End Property

Public Property Get sSerieAte() As String
    sSerieAte = mvarsSerieAte
End Property

Public Property Let sSerieDe(ByVal vData As String)
    mvarsSerieDe = vData
End Property

Public Property Get sSerieDe() As String
    sSerieDe = mvarsSerieDe
End Property

Public Property Let iOpcaoNFiscais(ByVal vData As Integer)
    mvariOpcaoNFiscais = vData
End Property

Public Property Get iOpcaoNFiscais() As Integer
    iOpcaoNFiscais = mvariOpcaoNFiscais
End Property




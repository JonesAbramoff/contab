VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFluxoContratoItemNFPag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarlFluxoId As Long
Private mvarsContrato As String
Private mvardtDataPag As Date
Private mvariItemNF As Integer
Private mvardValor As Double
Private mvarsSerie As String
Private mvarlNumNotaFiscal As Long
Private mvardtDataEmissaoNF As Date
Private mvarlFornecedor As Long
Private mvariFilialForn As Integer
Private mvariSeqContrato As Integer
Private mvarlNumIntNF As Long
Private mvarlNumIntContrato As Long

Public Property Let lNumIntContrato(ByVal vData As Long)
    mvarlNumIntContrato = vData
End Property

Public Property Get lNumIntContrato() As Long
    lNumIntContrato = mvarlNumIntContrato
End Property

Public Property Let lNumIntNF(ByVal vData As Long)
    mvarlNumIntNF = vData
End Property

Public Property Get lNumIntNF() As Long
    lNumIntNF = mvarlNumIntNF
End Property

Public Property Let iSeqContrato(ByVal vData As Integer)
    mvariSeqContrato = vData
End Property

Public Property Get iSeqContrato() As Integer
    iSeqContrato = mvariSeqContrato
End Property

Public Property Let iFilialForn(ByVal vData As Integer)
    mvariFilialForn = vData
End Property

Public Property Get iFilialForn() As Integer
    iFilialForn = mvariFilialForn
End Property

Public Property Let lFornecedor(ByVal vData As Long)
    mvarlFornecedor = vData
End Property

Public Property Get lFornecedor() As Long
    lFornecedor = mvarlFornecedor
End Property

Public Property Let dtDataEmissaoNF(ByVal vData As Date)
    mvardtDataEmissaoNF = vData
End Property

Public Property Get dtDataEmissaoNF() As Date
    dtDataEmissaoNF = mvardtDataEmissaoNF
End Property

Public Property Let lNumNotaFiscal(ByVal vData As Long)
    mvarlNumNotaFiscal = vData
End Property

Public Property Get lNumNotaFiscal() As Long
    lNumNotaFiscal = mvarlNumNotaFiscal
End Property

Public Property Let sSerie(ByVal vData As String)
    mvarsSerie = vData
End Property

Public Property Get sSerie() As String
    sSerie = mvarsSerie
End Property

Public Property Let dValor(ByVal vData As Double)
    mvardValor = vData
End Property

Public Property Get dValor() As Double
    dValor = mvardValor
End Property

Public Property Let iItemNF(ByVal vData As Integer)
    mvariItemNF = vData
End Property

Public Property Get iItemNF() As Integer
    iItemNF = mvariItemNF
End Property

Public Property Let dtDataPag(ByVal vData As Date)
    mvardtDataPag = vData
End Property

Public Property Get dtDataPag() As Date
    dtDataPag = mvardtDataPag
End Property

Public Property Let sContrato(ByVal vData As String)
    mvarsContrato = vData
End Property

Public Property Get sContrato() As String
    sContrato = mvarsContrato
End Property


Public Property Let lFluxoId(ByVal vData As Long)
    mvarlFluxoId = vData
End Property

Public Property Get lFluxoId() As Long
    lFluxoId = mvarlFluxoId
End Property




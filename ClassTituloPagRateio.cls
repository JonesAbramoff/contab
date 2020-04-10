VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTituloPagRateio"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlNumIntDocPag As Long
Private mvariSeq As Integer
Private mvarlCliente As Long
Private mvardValor As Double
Private mvarlHistorico As Long
Private mvarsReferencia As String
Private mvarsDossie As String
Private mvariCobrar As Integer
Private mvarlNumIntDocRec As Long
Private mvarlND As Long
Private mvardtDataGerND As Date

Public Property Let dtDataGerND(ByVal vData As Date)
    mvardtDataGerND = vData
End Property

Public Property Get dtDataGerND() As Date
    dtDataGerND = mvardtDataGerND
End Property

Public Property Let lND(ByVal vData As Long)
    mvarlND = vData
End Property

Public Property Get lND() As Long
    lND = mvarlND
End Property

Public Property Let lNumIntDocRec(ByVal vData As Long)
    mvarlNumIntDocRec = vData
End Property

Public Property Get lNumIntDocRec() As Long
    lNumIntDocRec = mvarlNumIntDocRec
End Property

Public Property Let iCobrar(ByVal vData As Integer)
    mvariCobrar = vData
End Property

Public Property Get iCobrar() As Integer
    iCobrar = mvariCobrar
End Property

Public Property Let sDossie(ByVal vData As String)
    mvarsDossie = vData
End Property

Public Property Get sDossie() As String
    sDossie = mvarsDossie
End Property

Public Property Let sReferencia(ByVal vData As String)
    mvarsReferencia = vData
End Property


Public Property Get sReferencia() As String
    sReferencia = mvarsReferencia
End Property

Public Property Let lHistorico(ByVal vData As Long)
    mvarlHistorico = vData
End Property

Public Property Get lHistorico() As Long
    lHistorico = mvarlHistorico
End Property

Public Property Let dValor(ByVal vData As Double)
    mvardValor = vData
End Property

Public Property Get dValor() As Double
    dValor = mvardValor
End Property

Public Property Let lCliente(ByVal vData As Long)
    mvarlCliente = vData
End Property

Public Property Get lCliente() As Long
    lCliente = mvarlCliente
End Property

Public Property Let iSeq(ByVal vData As Integer)
    mvariSeq = vData
End Property

Public Property Get iSeq() As Integer
    iSeq = mvariSeq
End Property

Public Property Let lNumIntDocPag(ByVal vData As Long)
    mvarlNumIntDocPag = vData
End Property

Public Property Get lNumIntDocPag() As Long
    lNumIntDocPag = mvarlNumIntDocPag
End Property

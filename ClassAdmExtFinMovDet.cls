VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassAdmExtFinMovDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarlNumIntDoc As Long
Private mvarlNumIntMov As Long
Private mvarlSeq As Long
Private mvariTipo As Integer
Private mvardValor As Double
Private mvarsNumCartao As String
Private mvardtDataCompra As Date
Private mvariNumParcela As Integer
Private mvarsAutorizacao As String
Private mvarsRO As String
Private mvarsNSU As String
Private mvarlNumIntBaixaParcRec As Long
Private mvarlNumIntParcRec As Long
Private mvardSaldoParcRec As Double
Private mvarlNumTitulo As Long
Private mvariCodErro As Integer
Private mvarlNumIntParc As Long
Private mvarlNumIntTitulo As Long
'########################################
'AUXILIARES
Private mvariQtdParcDif As Integer
Private mvariQtdParc As Integer
Private mvariQtdParcEnc As Integer
Private mvardVlrParcDif As Double
Private mvardVlrParc As Double
Private mvardVlrParcEnc As Double

Public Property Let dVlrParcDif(ByVal vData As Double)
    mvardVlrParcDif = vData
End Property

Public Property Get dVlrParcDif() As Double
    dVlrParcDif = mvardVlrParcDif
End Property

Public Property Let dVlrParc(ByVal vData As Double)
    mvardVlrParc = vData
End Property

Public Property Get dVlrParc() As Double
    dVlrParc = mvardVlrParc
End Property

Public Property Let dVlrParcEnc(ByVal vData As Double)
    mvardVlrParcEnc = vData
End Property

Public Property Get dVlrParcEnc() As Double
    dVlrParcEnc = mvardVlrParcEnc
End Property

Public Property Let iQtdParcDif(ByVal vData As Integer)
    mvariQtdParcDif = vData
End Property

Public Property Get iQtdParcDif() As Integer
    iQtdParcDif = mvariQtdParcDif
End Property

Public Property Let iQtdParc(ByVal vData As Integer)
    mvariQtdParc = vData
End Property

Public Property Get iQtdParc() As Integer
    iQtdParc = mvariQtdParc
End Property

Public Property Let iQtdParcEnc(ByVal vData As Integer)
    mvariQtdParcEnc = vData
End Property

Public Property Get iQtdParcEnc() As Integer
    iQtdParcEnc = mvariQtdParcEnc
End Property
'AUXILIARES
'########################################
Public Property Let dSaldoParcRec(ByVal vData As Double)
    mvardSaldoParcRec = vData
End Property

Public Property Get dSaldoParcRec() As Double
    dSaldoParcRec = mvardSaldoParcRec
End Property

Public Property Let lNumIntParcRec(ByVal vData As Long)
    mvarlNumIntParcRec = vData
End Property

Public Property Get lNumIntParcRec() As Long
    lNumIntParcRec = mvarlNumIntParcRec
End Property

Public Property Let lNumIntBaixaParcRec(ByVal vData As Long)
    mvarlNumIntBaixaParcRec = vData
End Property

Public Property Get lNumIntBaixaParcRec() As Long
    lNumIntBaixaParcRec = mvarlNumIntBaixaParcRec
End Property

Public Property Let sNSU(ByVal vData As String)
    mvarsNSU = vData
End Property

Public Property Get sNSU() As String
    sNSU = mvarsNSU
End Property

Public Property Let sRO(ByVal vData As String)
    mvarsRO = vData
End Property

Public Property Get sRO() As String
    sRO = mvarsRO
End Property

Public Property Let sAutorizacao(ByVal vData As String)
    mvarsAutorizacao = vData
End Property

Public Property Get sAutorizacao() As String
    sAutorizacao = mvarsAutorizacao
End Property

Public Property Let iNumParcela(ByVal vData As Integer)
    mvariNumParcela = vData
End Property

Public Property Get iNumParcela() As Integer
    iNumParcela = mvariNumParcela
End Property

Public Property Let dtDataCompra(ByVal vData As Date)
    mvardtDataCompra = vData
End Property

Public Property Get dtDataCompra() As Date
    dtDataCompra = mvardtDataCompra
End Property

Public Property Let sNumCartao(ByVal vData As String)
    mvarsNumCartao = vData
End Property

Public Property Get sNumCartao() As String
    sNumCartao = mvarsNumCartao
End Property

Public Property Let dValor(ByVal vData As Double)
    mvardValor = vData
End Property

Public Property Get dValor() As Double
    dValor = mvardValor
End Property

Public Property Let iTipo(ByVal vData As Integer)
    mvariTipo = vData
End Property

Public Property Get iTipo() As Integer
    iTipo = mvariTipo
End Property

Public Property Let iCodErro(ByVal vData As Integer)
    mvariCodErro = vData
End Property

Public Property Get iCodErro() As Integer
    iCodErro = mvariCodErro
End Property

Public Property Let lSeq(ByVal vData As Long)
    mvarlSeq = vData
End Property

Public Property Get lSeq() As Long
    lSeq = mvarlSeq
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let lNumIntMov(ByVal vData As Long)
    mvarlNumIntMov = vData
End Property

Public Property Get lNumIntMov() As Long
    lNumIntMov = mvarlNumIntMov
End Property

Public Property Let lNumTitulo(ByVal vData As Long)
    mvarlNumTitulo = vData
End Property

Public Property Get lNumTitulo() As Long
    lNumTitulo = mvarlNumTitulo
End Property

Public Property Let lNumIntParc(ByVal vData As Long)
    mvarlNumIntParc = vData
End Property

Public Property Get lNumIntParc() As Long
    lNumIntParc = mvarlNumIntParc
End Property

Public Property Let lNumIntTitulo(ByVal vData As Long)
    mvarlNumIntTitulo = vData
End Property

Public Property Get lNumIntTitulo() As Long
    lNumIntTitulo = mvarlNumIntTitulo
End Property

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTributacaoPRJProp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariFilialEmpresa As Integer
Private mvarsCodProposta As String
Private mvariTaxacaoAutomatica As Integer
Private mvariTipoTributacao As Integer
Private mvariTipoTributacaoManual As Integer
Private mvardICMSBase As Double
Private mvariICMSBaseManual As Integer
Private mvardICMSValor As Double
Private mvariICMSValorManual As Integer
Private mvardICMSSubstBase As Double
Private mvariICMSSubstBaseManual As Integer
Private mvardICMSSubstValor As Double
Private mvariICMSSubstValorManual As Integer
Private mvardIPIBase As Double
Private mvariIPIBaseManual As Integer
Private mvardIPIValor As Double
Private mvariIPIValorManual As Integer
Private mvardIRRFBase As Double
Private mvardIRRFAliquota As Double
Private mvariIRRFAliquotaManual As Integer
Private mvardIRRFValor As Double
Private mvariIRRFValorManual As Integer
Private mvariISSIncluso As Integer
Private mvardISSBase As Double
Private mvardISSAliquota As Double
Private mvariISSAliquotaManual As Integer
Private mvardISSValor As Double
Private mvariISSValorManual As Integer
Private mvariISSInclusoManual As Integer
Private mvardISSRetido As Double
Private mvariISSRetidoManual As Integer

Private mvardValorINSS As Double
Private mvariINSSRetido As Integer
Private mvardINSSBase As Double
Private mvardINSSDeducoes As Double
Private mvariINSSBaseManual As Integer
Private mvariINSSDeducoesManual As Integer
Private mvariINSSValorManual As Integer
Private mvariINSSRetidoManual As Integer
Private mvardPISCredito As Double
Private mvariPISCreditoManual As Integer
Private mvardCOFINSCredito As Double
Private mvariCOFINSCreditoManual As Integer
Private mvardIPICredito As Double
Private mvariIPICreditoManual As Integer
Private mvardICMSCredito As Double
Private mvariICMSCreditoManual As Integer

Private mvarobjTributacaoFrete As New ClassTributacaoItemPV
Private mvarobjTributacaoSeguro As New ClassTributacaoItemPV
Private mvarobjTributacaoDesconto As New ClassTributacaoItemPV
Private mvarobjTributacaoOutras As New ClassTributacaoItemPV

Private mvariPISRetidoManual As Integer
Private mvariCOFINSRetidoManual As Integer
Private mvariCSLLRetidoManual As Integer

Private mvariPISManual As Integer
Private mvariCOFINSManual As Integer

Private mvardPISRetido As Double
Private mvardCOFINSRetido As Double
Private mvardCSLLRetido As Double

Public Property Let iPISManual(ByVal vData As Integer)
    mvariPISManual = vData
End Property

Public Property Get iPISManual() As Integer
    iPISManual = mvariPISManual
End Property

Public Property Let iCOFINSManual(ByVal vData As Integer)
    mvariCOFINSManual = vData
End Property

Public Property Get iCOFINSManual() As Integer
    iCOFINSManual = mvariCOFINSManual
End Property

Public Property Let dISSRetido(ByVal vData As Double)
    mvardISSRetido = vData
End Property

Public Property Get dISSRetido() As Double
    dISSRetido = mvardISSRetido
End Property

Public Property Let iISSRetidoManual(ByVal vData As Integer)
    mvariISSRetidoManual = vData
End Property

Public Property Get iISSRetidoManual() As Integer
    iISSRetidoManual = mvariISSRetidoManual
End Property

Public Property Let iPISRetidoManual(ByVal vData As Integer)
    mvariPISRetidoManual = vData
End Property

Public Property Get iPISRetidoManual() As Integer
    iPISRetidoManual = mvariPISRetidoManual
End Property

Public Property Let iCOFINSRetidoManual(ByVal vData As Integer)
    mvariCOFINSRetidoManual = vData
End Property

Public Property Get iCOFINSRetidoManual() As Integer
    iCOFINSRetidoManual = mvariCOFINSRetidoManual
End Property

Public Property Let iCSLLRetidoManual(ByVal vData As Integer)
    mvariCSLLRetidoManual = vData
End Property

Public Property Get iCSLLRetidoManual() As Integer
    iCSLLRetidoManual = mvariCSLLRetidoManual
End Property

Public Property Let dCSLLRetido(ByVal vData As Double)
    mvardCSLLRetido = vData
End Property

Public Property Get dCSLLRetido() As Double
    dCSLLRetido = mvardCSLLRetido
End Property

Public Property Let dCOFINSRetido(ByVal vData As Double)
    mvardCOFINSRetido = vData
End Property

Public Property Get dCOFINSRetido() As Double
    dCOFINSRetido = mvardCOFINSRetido
End Property

Public Property Let dPISRetido(ByVal vData As Double)
    mvardPISRetido = vData
End Property

Public Property Get dPISRetido() As Double
    dPISRetido = mvardPISRetido
End Property

Public Property Get objTributacaoFrete() As ClassTributacaoItemPV
    Set objTributacaoFrete = mvarobjTributacaoFrete
End Property

Public Property Set objTributacaoFrete(ByVal vData As Object)
    Set mvarobjTributacaoFrete = vData
End Property

Public Property Get objTributacaoSeguro() As ClassTributacaoItemPV
    Set objTributacaoSeguro = mvarobjTributacaoSeguro
End Property

Public Property Set objTributacaoSeguro(ByVal vData As Object)
    Set mvarobjTributacaoSeguro = vData
End Property

Public Property Get objTributacaoDesconto() As ClassTributacaoItemPV
    Set objTributacaoDesconto = mvarobjTributacaoDesconto
End Property

Public Property Set objTributacaoDesconto(ByVal vData As Object)
    Set mvarobjTributacaoDesconto = vData
End Property

Public Property Get objTributacaoOutras() As ClassTributacaoItemPV
    Set objTributacaoOutras = mvarobjTributacaoOutras
End Property

Public Property Set objTributacaoOutras(ByVal vData As Object)
    Set mvarobjTributacaoOutras = vData
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let sCodProposta(ByVal vData As String)
    mvarsCodProposta = vData
End Property

Public Property Get sCodProposta() As String
    sCodProposta = mvarsCodProposta
End Property

Public Property Let iIRRFValorManual(ByVal vData As Integer)
    mvariIRRFValorManual = vData
End Property

Public Property Get iIRRFValorManual() As Integer
    iIRRFValorManual = mvariIRRFValorManual
End Property

Public Property Let dIRRFValor(ByVal vData As Double)
    mvardIRRFValor = vData
End Property

Public Property Get dIRRFValor() As Double
    dIRRFValor = mvardIRRFValor
End Property

Public Property Let iISSIncluso(ByVal vData As Integer)
    mvariISSIncluso = vData
End Property

Public Property Get iISSIncluso() As Integer
    iISSIncluso = mvariISSIncluso
End Property

Public Property Let iISSValorManual(ByVal vData As Integer)
    mvariISSValorManual = vData
End Property

Public Property Get iISSValorManual() As Integer
    iISSValorManual = mvariISSValorManual
End Property

Public Property Let iISSInclusoManual(ByVal vData As Integer)
    mvariISSInclusoManual = vData
End Property

Public Property Get iISSInclusoManual() As Integer
    iISSInclusoManual = mvariISSInclusoManual
End Property

Public Property Let dISSValor(ByVal vData As Double)
    mvardISSValor = vData
End Property

Public Property Get dISSValor() As Double
    dISSValor = mvardISSValor
End Property

Public Property Let iIPIValorManual(ByVal vData As Integer)
    mvariIPIValorManual = vData
End Property

Public Property Get iIPIValorManual() As Integer
    iIPIValorManual = mvariIPIValorManual
End Property

Public Property Let dIPIValor(ByVal vData As Double)
    mvardIPIValor = vData
End Property

Public Property Get dIPIValor() As Double
    dIPIValor = mvardIPIValor
End Property

Public Property Let iICMSSubstValorManual(ByVal vData As Integer)
    mvariICMSSubstValorManual = vData
End Property

Public Property Get iICMSSubstValorManual() As Integer
    iICMSSubstValorManual = mvariICMSSubstValorManual
End Property

Public Property Let dICMSSubstValor(ByVal vData As Double)
    mvardICMSSubstValor = vData
End Property

Public Property Get dICMSSubstValor() As Double
    dICMSSubstValor = mvardICMSSubstValor
End Property

Public Property Let iICMSSubstBaseManual(ByVal vData As Integer)
    mvariICMSSubstBaseManual = vData
End Property

Public Property Get iICMSSubstBaseManual() As Integer
    iICMSSubstBaseManual = mvariICMSSubstBaseManual
End Property

Public Property Let dICMSSubstBase(ByVal vData As Double)
    mvardICMSSubstBase = vData
End Property

Public Property Get dICMSSubstBase() As Double
    dICMSSubstBase = mvardICMSSubstBase
End Property

Public Property Let iICMSValorManual(ByVal vData As Integer)
    mvariICMSValorManual = vData
End Property

Public Property Get iICMSValorManual() As Integer
    iICMSValorManual = mvariICMSValorManual
End Property

Public Property Let dICMSValor(ByVal vData As Double)
    mvardICMSValor = vData
End Property

Public Property Get dICMSValor() As Double
    dICMSValor = mvardICMSValor
End Property

Public Property Let iICMSBaseManual(ByVal vData As Integer)
    mvariICMSBaseManual = vData
End Property

Public Property Get iICMSBaseManual() As Integer
    iICMSBaseManual = mvariICMSBaseManual
End Property

Public Property Let dICMSBase(ByVal vData As Double)
    mvardICMSBase = vData
End Property

Public Property Get dICMSBase() As Double
    dICMSBase = mvardICMSBase
End Property

Public Property Let iTipoTributacaoManual(ByVal vData As Integer)
    mvariTipoTributacaoManual = vData
End Property

Public Property Get iTipoTributacaoManual() As Integer
    iTipoTributacaoManual = mvariTipoTributacaoManual
End Property

Public Property Let iTipoTributacao(ByVal vData As Integer)
    mvariTipoTributacao = vData
End Property

Public Property Get iTipoTributacao() As Integer
    iTipoTributacao = mvariTipoTributacao
End Property

Public Property Let iTaxacaoAutomatica(ByVal vData As Integer)
    mvariTaxacaoAutomatica = vData
End Property

Public Property Get iTaxacaoAutomatica() As Integer
    iTaxacaoAutomatica = mvariTaxacaoAutomatica
End Property

Public Property Let iIPIBaseManual(ByVal vData As Integer)
    mvariIPIBaseManual = vData
End Property

Public Property Get iIPIBaseManual() As Integer
    iIPIBaseManual = mvariIPIBaseManual
End Property

Public Property Let dIPIBase(ByVal vData As Double)
    mvardIPIBase = vData
End Property

Public Property Get dIPIBase() As Double
    dIPIBase = mvardIPIBase
End Property

Public Property Let iISSAliquotaManual(ByVal vData As Integer)
    mvariISSAliquotaManual = vData
End Property

Public Property Get iISSAliquotaManual() As Integer
    iISSAliquotaManual = mvariISSAliquotaManual
End Property

Public Property Let dISSAliquota(ByVal vData As Double)
    mvardISSAliquota = vData
End Property

Public Property Get dISSAliquota() As Double
    dISSAliquota = mvardISSAliquota
End Property

Public Property Let iIRRFAliquotaManual(ByVal vData As Integer)
    mvariIRRFAliquotaManual = vData
End Property

Public Property Get iIRRFAliquotaManual() As Integer
    iIRRFAliquotaManual = mvariIRRFAliquotaManual
End Property

Public Property Let dIRRFAliquota(ByVal vData As Double)
    mvardIRRFAliquota = vData
End Property

Public Property Get dIRRFAliquota() As Double
    dIRRFAliquota = mvardIRRFAliquota
End Property

Public Property Let dISSBase(ByVal vData As Double)
    mvardISSBase = vData
End Property

Public Property Get dISSBase() As Double
    dISSBase = mvardISSBase
End Property

Public Property Let dIRRFBase(ByVal vData As Double)
    mvardIRRFBase = vData
End Property

Public Property Get dIRRFBase() As Double
    dIRRFBase = mvardIRRFBase
End Property

Private Sub Class_Initialize()

    mvarobjTributacaoFrete.lNumIntDoc = ITEM_TIPO_FRETE
    mvarobjTributacaoSeguro.lNumIntDoc = ITEM_TIPO_SEGURO
    mvarobjTributacaoDesconto.lNumIntDoc = ITEM_TIPO_DESCONTO
    mvarobjTributacaoOutras.lNumIntDoc = ITEM_TIPO_OUTRAS_DESP

End Sub

Public Property Let iINSSBaseManual(ByVal vData As Integer)
    mvariINSSBaseManual = vData
End Property

Public Property Get iINSSBaseManual() As Integer
    iINSSBaseManual = mvariINSSBaseManual
End Property

Public Property Let iINSSDeducoesManual(ByVal vData As Integer)
    mvariINSSDeducoesManual = vData
End Property

Public Property Get iINSSDeducoesManual() As Integer
    iINSSDeducoesManual = mvariINSSDeducoesManual
End Property

Public Property Let iINSSValorManual(ByVal vData As Integer)
    mvariINSSValorManual = vData
End Property

Public Property Get iINSSValorManual() As Integer
    iINSSValorManual = mvariINSSValorManual
End Property

Public Property Let iINSSRetidoManual(ByVal vData As Integer)
    mvariINSSRetidoManual = vData
End Property

Public Property Get iINSSRetidoManual() As Integer
    iINSSRetidoManual = mvariINSSRetidoManual
End Property

Public Property Let dINSSDeducoes(ByVal vData As Double)
    mvardINSSDeducoes = vData
End Property

Public Property Get dINSSDeducoes() As Double
    dINSSDeducoes = mvardINSSDeducoes
End Property

Public Property Let dINSSBase(ByVal vData As Double)
    mvardINSSBase = vData
End Property

Public Property Get dINSSBase() As Double
    dINSSBase = mvardINSSBase
End Property

Public Property Let iINSSRetido(ByVal vData As Integer)
    mvariINSSRetido = vData
End Property

Public Property Get iINSSRetido() As Integer
    iINSSRetido = mvariINSSRetido
End Property

Public Property Let dValorINSS(ByVal vData As Double)
    mvardValorINSS = vData
End Property

Public Property Get dValorINSS() As Double
    dValorINSS = mvardValorINSS
End Property

Public Property Let dINSSValor(ByVal vData As Double)
    mvardValorINSS = vData
End Property

Public Property Get dINSSValor() As Double
    dINSSValor = mvardValorINSS
End Property

Public Property Let iPISCreditoManual(ByVal vData As Integer)
    mvariPISCreditoManual = vData
End Property

Public Property Get iPISCreditoManual() As Integer
    iPISCreditoManual = mvariPISCreditoManual
End Property

Public Property Let iCOFINSCreditoManual(ByVal vData As Integer)
    mvariCOFINSCreditoManual = vData
End Property

Public Property Get iCOFINSCreditoManual() As Integer
    iCOFINSCreditoManual = mvariCOFINSCreditoManual
End Property

Public Property Let dCOFINSCredito(ByVal vData As Double)
    mvardCOFINSCredito = vData
End Property

Public Property Get dCOFINSCredito() As Double
    dCOFINSCredito = mvardCOFINSCredito
End Property

Public Property Let dPISCredito(ByVal vData As Double)
    mvardPISCredito = vData
End Property

Public Property Get dPISCredito() As Double
    dPISCredito = mvardPISCredito
End Property

Public Property Let dICMSCredito(ByVal vData As Double)
    mvardICMSCredito = vData
End Property

Public Property Get dICMSCredito() As Double
    dICMSCredito = mvardICMSCredito
End Property

Public Property Let dIPICredito(ByVal vData As Double)
    mvardIPICredito = vData
End Property

Public Property Get dIPICredito() As Double
    dIPICredito = mvardIPICredito
End Property

Public Property Let iIPICreditoManual(ByVal vData As Integer)
    mvariIPICreditoManual = vData
End Property

Public Property Get iIPICreditoManual() As Integer
    iIPICreditoManual = mvariIPICreditoManual
End Property

Public Property Let iICMSCreditoManual(ByVal vData As Integer)
    mvariICMSCreditoManual = vData
End Property

Public Property Get iICMSCreditoManual() As Integer
    iICMSCreditoManual = mvariICMSCreditoManual
End Property

Public Property Get dTotalRetencoes() As Double
'READ-ONLY, calculada
    dTotalRetencoes = mvardIRRFValor + mvardPISRetido + mvardCOFINSRetido + mvardCSLLRetido + IIf(mvariINSSRetido, mvardValorINSS, 0) + mvardISSRetido
End Property


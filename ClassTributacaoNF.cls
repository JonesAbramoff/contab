VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTributacaoNF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlNumIntDoc As Long 'local copy
Private mvarsNaturezaOpInterna As String 'local copy
Private mvariTipoTributacao As Integer 'local copy
Private mvardIPIBase As Double 'local copy
Private mvardIPIValor As Double 'local copy
Private mvardIPICredito As Double 'local copy
Private mvardICMSBase As Double 'local copy
Private mvardICMSValor As Double 'local copy
Private mvardICMSSubstBase As Double 'local copy
Private mvardICMSSubstValor As Double 'local copy
Private mvardICMSCredito As Double 'local copy
Private mvariISSIncluso As Integer 'local copy
Private mvardISSBase As Double 'local copy
Private mvardISSAliquota As Double 'local copy
Private mvardISSValor As Double 'local copy
Private mvardISSRetido As Double
Private mvardIRRFBase As Double 'local copy
Private mvardIRRFAliquota As Double 'local copy
Private mvardIRRFValor As Double 'local copy
Private mvardValorINSS As Double
Private mvariINSSRetido As Integer
Private mvardINSSBase As Double
Private mvardINSSDeducoes As Double

Private mvariNaturezaOpManual As Integer
Private mvariTipoTributacaoManual As Integer
Private mvariICMSBaseManual As Integer
Private mvariICMSValorManual As Integer
Private mvariICMSSubstBaseManual As Integer
Private mvariICMSSubstValorManual As Integer
Private mvariIPIBaseManual As Integer
Private mvariIPIValorManual As Integer
Private mvariIRRFAliquotaManual As Integer
Private mvariIRRFValorManual As Integer
Private mvariISSAliquotaManual As Integer
Private mvariISSValorManual As Integer
Private mvariISSRetidoManual As Integer

Private mvariINSSBaseManual As Integer
Private mvariINSSDeducoesManual As Integer
Private mvariINSSValorManual As Integer
Private mvariINSSRetidoManual As Integer

Private mvarobjTribComplNFFrete As ClassTribComplNF
Private mvarobjTribComplNFSeguro As ClassTribComplNF
Private mvarobjTribComplNFOutras As ClassTribComplNF
Private mvarobjTribComplNFDesconto As ClassTribComplNF

Private mvardPISCredito As Double
Private mvariPISCreditoManual As Integer
Private mvardCOFINSCredito As Double
Private mvariCOFINSCreditoManual As Integer

Private mvariPISManual As Integer
Private mvariCOFINSManual As Integer

Private mvariPISRetidoManual As Integer
Private mvariCOFINSRetidoManual As Integer
Private mvariCSLLRetidoManual As Integer

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

Public Property Let iISSValorManual(ByVal vData As Integer)
    mvariISSValorManual = vData
End Property

Public Property Get iISSValorManual() As Integer
    iISSValorManual = mvariISSValorManual
End Property

Public Property Let iIPIValorManual(ByVal vData As Integer)
    mvariIPIValorManual = vData
End Property

Public Property Get iIPIValorManual() As Integer
    iIPIValorManual = mvariIPIValorManual
End Property

Public Property Let iICMSSubstValorManual(ByVal vData As Integer)
    mvariICMSSubstValorManual = vData
End Property

Public Property Get iICMSSubstValorManual() As Integer
    iICMSSubstValorManual = mvariICMSSubstValorManual
End Property

Public Property Let iICMSSubstBaseManual(ByVal vData As Integer)
    mvariICMSSubstBaseManual = vData
End Property

Public Property Get iICMSSubstBaseManual() As Integer
    iICMSSubstBaseManual = mvariICMSSubstBaseManual
End Property

Public Property Let iICMSValorManual(ByVal vData As Integer)
    mvariICMSValorManual = vData
End Property

Public Property Get iICMSValorManual() As Integer
    iICMSValorManual = mvariICMSValorManual
End Property

Public Property Let iICMSBaseManual(ByVal vData As Integer)
    mvariICMSBaseManual = vData
End Property

Public Property Get iICMSBaseManual() As Integer
    iICMSBaseManual = mvariICMSBaseManual
End Property

Public Property Let iTipoTributacaoManual(ByVal vData As Integer)
    mvariTipoTributacaoManual = vData
End Property

Public Property Get iTipoTributacaoManual() As Integer
    iTipoTributacaoManual = mvariTipoTributacaoManual
End Property

Public Property Let iIPIBaseManual(ByVal vData As Integer)
    mvariIPIBaseManual = vData
End Property

Public Property Get iIPIBaseManual() As Integer
    iIPIBaseManual = mvariIPIBaseManual
End Property

Public Property Let iISSAliquotaManual(ByVal vData As Integer)
    mvariISSAliquotaManual = vData
End Property

Public Property Get iISSAliquotaManual() As Integer
    iISSAliquotaManual = mvariISSAliquotaManual
End Property

Public Property Let iIRRFAliquotaManual(ByVal vData As Integer)
    mvariIRRFAliquotaManual = vData
End Property

Public Property Get iIRRFAliquotaManual() As Integer
    iIRRFAliquotaManual = mvariIRRFAliquotaManual
End Property

Public Property Let iIRRFValorManual(ByVal vData As Integer)
    mvariIRRFValorManual = vData
End Property

Public Property Get iIRRFValorManual() As Integer
    iIRRFValorManual = mvariIRRFValorManual
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

Public Property Let dIPICredito(ByVal vData As Double)
    mvardIPICredito = vData
End Property

Public Property Get dIPICredito() As Double
    dIPICredito = mvardIPICredito
End Property

Public Property Let dICMSCredito(ByVal vData As Double)
    mvardICMSCredito = vData
End Property

Public Property Get dICMSCredito() As Double
    dICMSCredito = mvardICMSCredito
End Property



Public Property Get objTributacaoFrete() As ClassTribComplNF
    Set objTributacaoFrete = mvarobjTribComplNFFrete
End Property

Public Property Set objTributacaoFrete(ByVal vData As Object)
    Set mvarobjTribComplNFFrete = vData
End Property

Public Property Get objTributacaoSeguro() As ClassTribComplNF
    Set objTributacaoSeguro = mvarobjTribComplNFSeguro
End Property

Public Property Set objTributacaoSeguro(ByVal vData As Object)
    Set mvarobjTribComplNFSeguro = vData
End Property

Public Property Get objTributacaoDesconto() As ClassTribComplNF
    Set objTributacaoDesconto = mvarobjTribComplNFDesconto
End Property

Public Property Set objTributacaoDesconto(ByVal vData As Object)
    Set mvarobjTribComplNFDesconto = vData
End Property

Public Property Get objTributacaoOutras() As ClassTribComplNF
    Set objTributacaoOutras = mvarobjTribComplNFOutras
End Property

Public Property Set objTributacaoOutras(ByVal vData As Object)
    Set mvarobjTribComplNFOutras = vData
End Property


Public Property Let dIRRFValor(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dIRRFValor = 5
    mvardIRRFValor = vData
End Property


Public Property Get dIRRFValor() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dIRRFValor
    dIRRFValor = mvardIRRFValor
End Property



Public Property Let dIRRFAliquota(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dIRRFAliquota = 5
    mvardIRRFAliquota = vData
End Property


Public Property Get dIRRFAliquota() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dIRRFAliquota
    dIRRFAliquota = mvardIRRFAliquota
End Property



Public Property Let dIRRFBase(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dIRRFBase = 5
    mvardIRRFBase = vData
End Property


Public Property Get dIRRFBase() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dIRRFBase
    dIRRFBase = mvardIRRFBase
End Property



Public Property Let dISSValor(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dISSValor = 5
    mvardISSValor = vData
End Property


Public Property Get dISSValor() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dISSValor
    dISSValor = mvardISSValor
End Property



Public Property Let dISSAliquota(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dISSAliquota = 5
    mvardISSAliquota = vData
End Property


Public Property Get dISSAliquota() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dISSAliquota
    dISSAliquota = mvardISSAliquota
End Property



Public Property Let dISSBase(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dISSBase = 5
    mvardISSBase = vData
End Property


Public Property Get dISSBase() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dISSBase
    dISSBase = mvardISSBase
End Property



Public Property Let iISSIncluso(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iISSIncluso = 5
    mvariISSIncluso = vData
End Property


Public Property Get iISSIncluso() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iISSIncluso
    iISSIncluso = mvariISSIncluso
End Property



Public Property Let dICMSSubstValor(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dICMSSubstValor = 5
    mvardICMSSubstValor = vData
End Property


Public Property Get dICMSSubstValor() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dICMSSubstValor
    dICMSSubstValor = mvardICMSSubstValor
End Property



Public Property Let dICMSSubstBase(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dICMSSubstBase = 5
    mvardICMSSubstBase = vData
End Property


Public Property Get dICMSSubstBase() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dICMSSubstBase
    dICMSSubstBase = mvardICMSSubstBase
End Property



Public Property Let dICMSValor(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dICMSValor = 5
    mvardICMSValor = vData
End Property


Public Property Get dICMSValor() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dICMSValor
    dICMSValor = mvardICMSValor
End Property



Public Property Let dICMSBase(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dICMSBase = 5
    mvardICMSBase = vData
End Property


Public Property Get dICMSBase() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dICMSBase
    dICMSBase = mvardICMSBase
End Property



Public Property Let dIPIValor(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dIPIValor = 5
    mvardIPIValor = vData
End Property


Public Property Get dIPIValor() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dIPIValor
    dIPIValor = mvardIPIValor
End Property



Public Property Let dIPIBase(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dIPIBase = 5
    mvardIPIBase = vData
End Property


Public Property Get dIPIBase() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dIPIBase
    dIPIBase = mvardIPIBase
End Property



Public Property Let iTipoTributacao(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iTipoTributacao = 5
    mvariTipoTributacao = vData
End Property


Public Property Get iTipoTributacao() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iTipoTributacao
    iTipoTributacao = mvariTipoTributacao
End Property



Public Property Let sNaturezaOpInterna(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sNaturezaOpInterna = 5
    mvarsNaturezaOpInterna = vData
End Property


Public Property Get sNaturezaOpInterna() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sNaturezaOpInterna
    sNaturezaOpInterna = mvarsNaturezaOpInterna
End Property

Public Property Let iNaturezaOpManual(ByVal vData As Integer)
    mvariNaturezaOpManual = vData
End Property

Public Property Get iNaturezaOpManual() As Integer
    iNaturezaOpManual = mvariNaturezaOpManual
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lNumIntDoc = 5
    mvarlNumIntDoc = vData
End Property


Public Property Get lNumIntDoc() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lNumIntDoc
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Get dTotalRetencoes() As Double
'READ-ONLY, calculada
    dTotalRetencoes = mvardIRRFValor + mvardPISRetido + mvardCOFINSRetido + mvardCSLLRetido + IIf(mvariINSSRetido, mvardValorINSS, 0) + mvardISSRetido
End Property



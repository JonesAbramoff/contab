VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTributacaoDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

Private mvarcolItens As New Collection
Private mvarcolTiposTribICMS As New Collection
Private mvarcolTiposTribIPI As New Collection
Private mvarcolTiposTribPISCOFINS As New Collection
Private mvarcolTiposTribICMSSimples As New Collection
Private mvarcolTiposTribISS As New Collection

Private mvarobjDoc As Object
Private mvarobjTribDocFrete As Object
Private mvarobjTribDocSeguro As Object
Private mvarobjTribDocDespesa As Object
Private mvarobjTribDocDesconto As Object
Private mvarsNaturezaOpInterna As String 'local copy
Private mvariNaturezaOpManual As Integer
Private mvardValorFrete As Double
Private mvardValorDesconto As Double
Private mvardValorSeguro As Double
Private mvardValorOutrasDespesas As Double

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvariTipoDoc As Integer
Private mvarlNumIntDoc As Long
Private mvariTaxacaoAutomatica As Integer
Private mvariTipoTributacao As Integer
Private mvariTipoTributacaoManual As Integer
Private mvardIPIBase As Double
Private mvariIPIBaseManual As Integer
Private mvardIPIValor As Double
Private mvariIPIValorManual As Integer
Private mvardICMSBase As Double
Private mvariICMSBaseManual As Integer
Private mvardICMSValor As Double
Private mvariICMSValorManual As Integer
Private mvardICMSSubstBase As Double
Private mvariICMSSubstBaseManual As Integer
Private mvardICMSSubstValor As Double
Private mvariICMSSubstValorManual As Integer
Private mvariISSIncluso As Integer
Private mvardISSBase As Double
Private mvariISSBaseManual As Integer
Private mvardISSAliquota As Double
Private mvariISSAliquotaManual As Integer
Private mvardISSValor As Double
Private mvariISSValorManual As Integer
Private mvardIRRFBase As Double
Private mvariIRRFBaseManual As Integer
Private mvardIRRFAliquota As Double
Private mvariIRRFAliquotaManual As Integer
Private mvardIRRFValor As Double
Private mvariIRRFValorManual As Integer
Private mvardPISRetido As Double
Private mvardCOFINSRetido As Double
Private mvardCSLLRetido As Double
Private mvariPISRetidoManual As Integer
Private mvariCOFINSRetidoManual As Integer
Private mvariCSLLRetidoManual As Integer
Private mvardISSRetido As Double
Private mvariISSRetidoManual As Integer
Private mvardPISCredito As Double
Private mvariPISCreditoManual As Integer
Private mvardCOFINSCredito As Double
Private mvariCOFINSCreditoManual As Integer
Private mvardIPICredito As Double
Private mvariIPICreditoManual As Integer
Private mvardICMSCredito As Double
Private mvariICMSCreditoManual As Integer
Private mvardValorINSS As Double
Private mvariValorINSSManual As Integer
Private mvariINSSRetido As Integer
Private mvariINSSRetidoManual As Integer
Private mvardINSSValorBase As Double
Private mvariINSSValorBaseManual As Integer
Private mvardINSSValorDeducoes As Double
Private mvariINSSValorDeducoesManual As Integer
Private mvardValorAux As Double
Private mvardPISValor As Double
Private mvardCOFINSValor As Double

Private mvardTotTrib As Double
Private mvariTotTribManual As Integer
Private mvariTotTribTipo As Integer
Private mvariTotTribTipoManual As Integer

'nfe 3.10
Public dICMSValorIsento As Double
Public dISSValorDeducao As Double
Public dISSValorOutrasRet As Double
Public dISSValorDescIncond As Double
Public dISSValorDescCond As Double
Public iDestino As Integer
Public iDestinoManual As Integer
Public iFinalidadeNFe As Integer
Public iFinalidadeNFeManual As Integer
Public iIndConsumidorFinal As Integer
Public iIndConsumidorFinalManual As Integer
Public iIndPresenca As Integer
Public iIndPresencaManual As Integer
Public iRegTribEspecial As Integer
Public iRegTribEspecialManual As Integer
Public dtDataPrestServico As Date
Public iDataPrestServicoManual As Integer
Public dIIValor As Double
Public sISSMunicFaturamento As String

'IBPTax 0.0.9
Public dTotTribFed As Double
Public dTotTribEst As Double
Public dTotTribMunic As Double
Public iTotTribFedManual As Integer
Public iTotTribEstManual As Integer
Public iTotTribMunicManual As Integer

'NT 2015.003 (ICMS Interestadual)
Public dICMSInterestVlrFCPUFDest As Double
Public dICMSInterestVlrUFDest As Double
Public dICMSInterestVlrUFRemet As Double

'nfe 4.00
Public dICMSVlrFCP As Double
Public dICMSVlrFCPST As Double
Public dICMSVlrFCPSTRet As Double
Public dIPIVlrDevolvido As Double

Public Property Get dPISSTValor() As Double
Dim objTribDocItem As ClassTributacaoDocItem
Dim dAux As Double
    
    dAux = 0

    For Each objTribDocItem In colItens
        dAux = Arredonda_Moeda(dAux + objTribDocItem.dPISSTValor)
    Next
    
    dPISSTValor = dAux

End Property

Public Property Get dCOFINSSTValor() As Double
Dim objTribDocItem As ClassTributacaoDocItem
Dim dAux As Double
    
    dAux = 0

    For Each objTribDocItem In colItens
        dAux = Arredonda_Moeda(dAux + objTribDocItem.dCOFINSSTValor)
    Next
    
    dCOFINSSTValor = dAux

End Property

Public Property Let dPISValor(ByVal vData As Double)
    mvardPISValor = vData
End Property

Public Property Get dPISValor() As Double
    dPISValor = mvardPISValor
End Property

Public Property Let dCOFINSValor(ByVal vData As Double)
    mvardCOFINSValor = vData
End Property

Public Property Get dCOFINSValor() As Double
    dCOFINSValor = mvardCOFINSValor
End Property

Public Property Let dValorAux(ByVal vData As Double)
    mvardValorAux = vData
End Property

Public Property Get dValorAux() As Double
    dValorAux = mvardValorAux
End Property

Public Property Let iTipoDoc(ByVal vData As Integer)
    mvariTipoDoc = vData
End Property

Public Property Get iTipoDoc() As Integer
    iTipoDoc = mvariTipoDoc
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let iTaxacaoAutomatica(ByVal vData As Integer)
    mvariTaxacaoAutomatica = vData
End Property

Public Property Get iTaxacaoAutomatica() As Integer
    iTaxacaoAutomatica = mvariTaxacaoAutomatica
End Property

Public Property Let iTipoTributacao(ByVal vData As Integer)
    mvariTipoTributacao = vData
End Property

Public Property Get iTipoTributacao() As Integer
    iTipoTributacao = mvariTipoTributacao
End Property

Public Property Let iTipoTributacaoManual(ByVal vData As Integer)
    mvariTipoTributacaoManual = vData
End Property

Public Property Get iTipoTributacaoManual() As Integer
    iTipoTributacaoManual = mvariTipoTributacaoManual
End Property

Public Property Let dIPIBase(ByVal vData As Double)
    mvardIPIBase = vData
End Property

Public Property Get dIPIBase() As Double
    dIPIBase = mvardIPIBase
End Property

Public Property Let iIPIBaseManual(ByVal vData As Integer)
    mvariIPIBaseManual = vData
End Property

Public Property Get iIPIBaseManual() As Integer
    iIPIBaseManual = mvariIPIBaseManual
End Property

Public Property Let dIPIValor(ByVal vData As Double)
    mvardIPIValor = vData
End Property

Public Property Get dIPIValor() As Double
    dIPIValor = mvardIPIValor
End Property

Public Property Let iIPIValorManual(ByVal vData As Integer)
    mvariIPIValorManual = vData
End Property

Public Property Get iIPIValorManual() As Integer
    iIPIValorManual = mvariIPIValorManual
End Property

Public Property Let dICMSBase(ByVal vData As Double)
    mvardICMSBase = vData
End Property

Public Property Get dICMSBase() As Double
    dICMSBase = mvardICMSBase
End Property

Public Property Let iICMSBaseManual(ByVal vData As Integer)
    mvariICMSBaseManual = vData
End Property

Public Property Get iICMSBaseManual() As Integer
    iICMSBaseManual = mvariICMSBaseManual
End Property

Public Property Let dICMSValor(ByVal vData As Double)
    mvardICMSValor = vData
End Property

Public Property Get dICMSValor() As Double
    dICMSValor = mvardICMSValor
End Property

Public Property Let iICMSValorManual(ByVal vData As Integer)
    mvariICMSValorManual = vData
End Property

Public Property Get iICMSValorManual() As Integer
    iICMSValorManual = mvariICMSValorManual
End Property

Public Property Let dICMSSubstBase(ByVal vData As Double)
    mvardICMSSubstBase = vData
End Property

Public Property Get dICMSSubstBase() As Double
    dICMSSubstBase = mvardICMSSubstBase
End Property

Public Property Let iICMSSubstBaseManual(ByVal vData As Integer)
    mvariICMSSubstBaseManual = vData
End Property

Public Property Get iICMSSubstBaseManual() As Integer
    iICMSSubstBaseManual = mvariICMSSubstBaseManual
End Property

Public Property Let dICMSSubstValor(ByVal vData As Double)
    mvardICMSSubstValor = vData
End Property

Public Property Get dICMSSubstValor() As Double
    dICMSSubstValor = mvardICMSSubstValor
End Property

Public Property Let iICMSSubstValorManual(ByVal vData As Integer)
    mvariICMSSubstValorManual = vData
End Property

Public Property Get iICMSSubstValorManual() As Integer
    iICMSSubstValorManual = mvariICMSSubstValorManual
End Property

Public Property Let iISSIncluso(ByVal vData As Integer)
    mvariISSIncluso = vData
End Property

Public Property Get iISSIncluso() As Integer
    iISSIncluso = mvariISSIncluso
End Property

Public Property Let dISSBase(ByVal vData As Double)
    mvardISSBase = vData
End Property

Public Property Get dISSBase() As Double
    dISSBase = mvardISSBase
End Property

Public Property Let iISSBaseManual(ByVal vData As Integer)
    mvariISSBaseManual = vData
End Property

Public Property Get iISSBaseManual() As Integer
    iISSBaseManual = mvariISSBaseManual
End Property

Public Property Let dISSAliquota(ByVal vData As Double)
    mvardISSAliquota = vData
End Property

Public Property Get dISSAliquota() As Double
    dISSAliquota = mvardISSAliquota
End Property

Public Property Let iISSAliquotaManual(ByVal vData As Integer)
    mvariISSAliquotaManual = vData
End Property

Public Property Get iISSAliquotaManual() As Integer
    iISSAliquotaManual = mvariISSAliquotaManual
End Property

Public Property Let dISSValor(ByVal vData As Double)
    mvardISSValor = vData
End Property

Public Property Get dISSValor() As Double
    dISSValor = mvardISSValor
End Property

Public Property Let iISSValorManual(ByVal vData As Integer)
    mvariISSValorManual = vData
End Property

Public Property Get iISSValorManual() As Integer
    iISSValorManual = mvariISSValorManual
End Property

Public Property Let dIRRFBase(ByVal vData As Double)
    mvardIRRFBase = vData
End Property

Public Property Get dIRRFBase() As Double
    dIRRFBase = mvardIRRFBase
End Property

Public Property Let iIRRFBaseManual(ByVal vData As Integer)
    mvariIRRFBaseManual = vData
End Property

Public Property Get iIRRFBaseManual() As Integer
    iIRRFBaseManual = mvariIRRFBaseManual
End Property

Public Property Let dIRRFAliquota(ByVal vData As Double)
    mvardIRRFAliquota = vData
End Property

Public Property Get dIRRFAliquota() As Double
    dIRRFAliquota = mvardIRRFAliquota
End Property

Public Property Let iIRRFAliquotaManual(ByVal vData As Integer)
    mvariIRRFAliquotaManual = vData
End Property

Public Property Get iIRRFAliquotaManual() As Integer
    iIRRFAliquotaManual = mvariIRRFAliquotaManual
End Property

Public Property Let dIRRFValor(ByVal vData As Double)
    mvardIRRFValor = vData
End Property

Public Property Get dIRRFValor() As Double
    dIRRFValor = mvardIRRFValor
End Property

Public Property Let iIRRFValorManual(ByVal vData As Integer)
    mvariIRRFValorManual = vData
End Property

Public Property Get iIRRFValorManual() As Integer
    iIRRFValorManual = mvariIRRFValorManual
End Property

Public Property Let dPISRetido(ByVal vData As Double)
    mvardPISRetido = vData
End Property

Public Property Get dPISRetido() As Double
    dPISRetido = mvardPISRetido
End Property

Public Property Let dCOFINSRetido(ByVal vData As Double)
    mvardCOFINSRetido = vData
End Property

Public Property Get dCOFINSRetido() As Double
    dCOFINSRetido = mvardCOFINSRetido
End Property

Public Property Let dCSLLRetido(ByVal vData As Double)
    mvardCSLLRetido = vData
End Property

Public Property Get dCSLLRetido() As Double
    dCSLLRetido = mvardCSLLRetido
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

Public Property Let dCOFINSCredito(ByVal vData As Double)
    mvardCOFINSCredito = vData
End Property

Public Property Get dCOFINSCredito() As Double
    dCOFINSCredito = mvardCOFINSCredito
End Property

Public Property Let iCOFINSCreditoManual(ByVal vData As Integer)
    mvariCOFINSCreditoManual = vData
End Property

Public Property Get iCOFINSCreditoManual() As Integer
    iCOFINSCreditoManual = mvariCOFINSCreditoManual
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

Public Property Let dICMSCredito(ByVal vData As Double)
    mvardICMSCredito = vData
End Property

Public Property Get dICMSCredito() As Double
    dICMSCredito = mvardICMSCredito
End Property

Public Property Let iICMSCreditoManual(ByVal vData As Integer)
    mvariICMSCreditoManual = vData
End Property

Public Property Get iICMSCreditoManual() As Integer
    iICMSCreditoManual = mvariICMSCreditoManual
End Property

Public Property Let dValorINSS(ByVal vData As Double)
    mvardValorINSS = vData
End Property

Public Property Get dValorINSS() As Double
    dValorINSS = mvardValorINSS
End Property

Public Property Let iValorINSSManual(ByVal vData As Integer)
    mvariValorINSSManual = vData
End Property

Public Property Get iValorINSSManual() As Integer
    iValorINSSManual = mvariValorINSSManual
End Property

Public Property Let iINSSRetido(ByVal vData As Integer)
    mvariINSSRetido = vData
End Property

Public Property Get iINSSRetido() As Integer
    iINSSRetido = mvariINSSRetido
End Property

Public Property Let iINSSRetidoManual(ByVal vData As Integer)
    mvariINSSRetidoManual = vData
End Property

Public Property Get iINSSRetidoManual() As Integer
    iINSSRetidoManual = mvariINSSRetidoManual
End Property

Public Property Let dINSSValorBase(ByVal vData As Double)
    mvardINSSValorBase = vData
End Property

Public Property Get dINSSValorBase() As Double
    dINSSValorBase = mvardINSSValorBase
End Property

Public Property Let iINSSValorBaseManual(ByVal vData As Integer)
    mvariINSSValorBaseManual = vData
End Property

Public Property Get iINSSValorBaseManual() As Integer
    iINSSValorBaseManual = mvariINSSValorBaseManual
End Property

Public Property Let dINSSValorDeducoes(ByVal vData As Double)
    mvardINSSValorDeducoes = vData
End Property

Public Property Get dINSSValorDeducoes() As Double
    dINSSValorDeducoes = mvardINSSValorDeducoes
End Property

Public Property Let iINSSValorDeducoesManual(ByVal vData As Integer)
    mvariINSSValorDeducoesManual = vData
End Property

Public Property Get iINSSValorDeducoesManual() As Integer
    iINSSValorDeducoesManual = mvariINSSValorDeducoesManual
End Property

Public Property Set colItens(ByVal vData As Object)
    Set mvarcolItens = vData
End Property

Public Property Get colItens() As Collection
    Set colItens = mvarcolItens
End Property

Public Property Set objDoc(ByVal vData As Object)
    Set mvarobjDoc = vData
End Property

Public Property Get objDoc() As Object
    Set objDoc = mvarobjDoc
End Property

Public Property Set objTribDocFrete(ByVal vData As Object)
    Set mvarobjTribDocFrete = vData
End Property

Public Property Get objTribDocFrete() As Object
    Set objTribDocFrete = mvarobjTribDocFrete
End Property

Public Property Set objTribDocSeguro(ByVal vData As Object)
    Set mvarobjTribDocSeguro = vData
End Property

Public Property Get objTribDocSeguro() As Object
    Set objTribDocSeguro = mvarobjTribDocSeguro
End Property

Public Property Set objTribDocDespesa(ByVal vData As Object)
    Set mvarobjTribDocDespesa = vData
End Property

Public Property Get objTribDocDespesa() As Object
    Set objTribDocDespesa = mvarobjTribDocDespesa
End Property

Public Property Set objTribDocDesconto(ByVal vData As Object)
    Set mvarobjTribDocDesconto = vData
End Property

Public Property Get objTribDocDesconto() As Object
    Set objTribDocDesconto = mvarobjTribDocDesconto
End Property

Public Property Let dTotTrib(ByVal vData As Double)
    mvardTotTrib = vData
End Property

Public Property Get dTotTrib() As Double
    dTotTrib = mvardTotTrib
End Property

Public Property Let iTotTribManual(ByVal vData As Integer)
    mvariTotTribManual = vData
End Property

Public Property Get iTotTribManual() As Integer
    iTotTribManual = mvariTotTribManual
End Property

Public Property Let iTotTribTipo(ByVal vData As Integer)
    mvariTotTribTipo = vData
End Property

Public Property Get iTotTribTipo() As Integer
    iTotTribTipo = mvariTotTribTipo
End Property

Public Property Let iTotTribTipoManual(ByVal vData As Integer)
    mvariTotTribTipoManual = vData
End Property

Public Property Get iTotTribTipoManual() As Integer
    iTotTribTipoManual = mvariTotTribTipoManual
End Property

Public Sub Copia(ByVal objTribDocOriginal As ClassTributacaoDoc)

Dim objTribDocItemOrig As ClassTributacaoDocItem
Dim objTribDocItem As ClassTributacaoDocItem

On Error GoTo Erro_Copia

    Call Copia2(objTribDocOriginal)
    
    Set colItens = New Collection
    For Each objTribDocItemOrig In objTribDocOriginal.colItens
        Set objTribDocItem = New ClassTributacaoDocItem
        Call objTribDocItem.Copia(objTribDocItemOrig)
        colItens.Add objTribDocItem
    Next
    
    Exit Sub

Erro_Copia:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202900)

    End Select

    Exit Sub
    
End Sub

Public Sub Copia2(objTribDocOriginal As Variant)

On Error GoTo Erro_Copia2

    iTipoDoc = objTribDocOriginal.iTipoDoc
    lNumIntDoc = objTribDocOriginal.lNumIntDoc
    iTaxacaoAutomatica = objTribDocOriginal.iTaxacaoAutomatica
    iTipoTributacao = objTribDocOriginal.iTipoTributacao
    iTipoTributacaoManual = objTribDocOriginal.iTipoTributacaoManual
    dIPIBase = objTribDocOriginal.dIPIBase
    iIPIBaseManual = objTribDocOriginal.iIPIBaseManual
    dIPIValor = objTribDocOriginal.dIPIValor
    iIPIValorManual = objTribDocOriginal.iIPIValorManual
    dICMSBase = objTribDocOriginal.dICMSBase
    iICMSBaseManual = objTribDocOriginal.iICMSBaseManual
    dICMSValor = objTribDocOriginal.dICMSValor
    iICMSValorManual = objTribDocOriginal.iICMSValorManual
    dICMSSubstBase = objTribDocOriginal.dICMSSubstBase
    iICMSSubstBaseManual = objTribDocOriginal.iICMSSubstBaseManual
    dICMSSubstValor = objTribDocOriginal.dICMSSubstValor
    iICMSSubstValorManual = objTribDocOriginal.iICMSSubstValorManual
    iISSIncluso = objTribDocOriginal.iISSIncluso
    dISSBase = objTribDocOriginal.dISSBase
    iISSBaseManual = objTribDocOriginal.iISSBaseManual
    dISSAliquota = objTribDocOriginal.dISSAliquota
    iISSAliquotaManual = objTribDocOriginal.iISSAliquotaManual
    dISSValor = objTribDocOriginal.dISSValor
    iISSValorManual = objTribDocOriginal.iISSValorManual
    dIRRFBase = objTribDocOriginal.dIRRFBase
    iIRRFBaseManual = objTribDocOriginal.iIRRFBaseManual
    dIRRFAliquota = objTribDocOriginal.dIRRFAliquota
    iIRRFAliquotaManual = objTribDocOriginal.iIRRFAliquotaManual
    dIRRFValor = objTribDocOriginal.dIRRFValor
    iIRRFValorManual = objTribDocOriginal.iIRRFValorManual
    dPISRetido = objTribDocOriginal.dPISRetido
    dCOFINSRetido = objTribDocOriginal.dCOFINSRetido
    dCSLLRetido = objTribDocOriginal.dCSLLRetido
    iPISRetidoManual = objTribDocOriginal.iPISRetidoManual
    iCOFINSRetidoManual = objTribDocOriginal.iCOFINSRetidoManual
    iCSLLRetidoManual = objTribDocOriginal.iCSLLRetidoManual
    dISSRetido = objTribDocOriginal.dISSRetido
    iISSRetidoManual = objTribDocOriginal.iISSRetidoManual
    dPISCredito = objTribDocOriginal.dPISCredito
    iPISCreditoManual = objTribDocOriginal.iPISCreditoManual
    dCOFINSCredito = objTribDocOriginal.dCOFINSCredito
    iCOFINSCreditoManual = objTribDocOriginal.iCOFINSCreditoManual
    dIPICredito = objTribDocOriginal.dIPICredito
    iIPICreditoManual = objTribDocOriginal.iIPICreditoManual
    dICMSCredito = objTribDocOriginal.dICMSCredito
    iICMSCreditoManual = objTribDocOriginal.iICMSCreditoManual
    dValorINSS = objTribDocOriginal.dValorINSS
    iValorINSSManual = objTribDocOriginal.iValorINSSManual
    iINSSRetido = objTribDocOriginal.iINSSRetido
    iINSSRetidoManual = objTribDocOriginal.iINSSRetidoManual
    dINSSValorBase = objTribDocOriginal.dINSSValorBase
    iINSSValorBaseManual = objTribDocOriginal.iINSSValorBaseManual
    dINSSValorDeducoes = objTribDocOriginal.dINSSValorDeducoes
    iINSSValorDeducoesManual = objTribDocOriginal.iINSSValorDeducoesManual
    sNaturezaOpInterna = objTribDocOriginal.sNaturezaOpInterna
    dTotTrib = objTribDocOriginal.dTotTrib
    iTotTribManual = objTribDocOriginal.iTotTribManual
    iTotTribTipo = objTribDocOriginal.iTotTribTipo
    iTotTribTipoManual = objTribDocOriginal.iTotTribTipoManual
    iNaturezaOpManual = objTribDocOriginal.iNaturezaOpManual
    
    dTotTribFed = objTribDocOriginal.dTotTribFed
    iTotTribFedManual = objTribDocOriginal.iTotTribFedManual
    dTotTribEst = objTribDocOriginal.dTotTribEst
    iTotTribEstManual = objTribDocOriginal.iTotTribEstManual
    dTotTribMunic = objTribDocOriginal.dTotTribMunic
    iTotTribMunicManual = objTribDocOriginal.iTotTribMunicManual
    
    
    'nfe 3.10
    dPISValor = objTribDocOriginal.dPISValor
    dCOFINSValor = objTribDocOriginal.dCOFINSValor
    dICMSValorIsento = objTribDocOriginal.dICMSValorIsento
    dISSValorDeducao = objTribDocOriginal.dISSValorDeducao
    dISSValorOutrasRet = objTribDocOriginal.dISSValorOutrasRet
    dISSValorDescIncond = objTribDocOriginal.dISSValorDescIncond
    dISSValorDescCond = objTribDocOriginal.dISSValorDescCond
    iDestino = objTribDocOriginal.iDestino
    iDestinoManual = objTribDocOriginal.iDestinoManual
    iFinalidadeNFe = objTribDocOriginal.iFinalidadeNFe
    iFinalidadeNFeManual = objTribDocOriginal.iFinalidadeNFeManual
    iIndConsumidorFinal = objTribDocOriginal.iIndConsumidorFinal
    iIndConsumidorFinalManual = objTribDocOriginal.iIndConsumidorFinalManual
    iIndPresenca = objTribDocOriginal.iIndPresenca
    iIndPresencaManual = objTribDocOriginal.iIndPresencaManual
    iRegTribEspecial = objTribDocOriginal.iRegTribEspecial
    iRegTribEspecialManual = objTribDocOriginal.iRegTribEspecialManual
    dtDataPrestServico = objTribDocOriginal.dtDataPrestServico
    iDataPrestServicoManual = objTribDocOriginal.iDataPrestServicoManual
    dIIValor = objTribDocOriginal.dIIValor
    
    'NT 2015.003 (ICMS Interestadual)
    dICMSInterestVlrFCPUFDest = objTribDocOriginal.dICMSInterestVlrFCPUFDest
    dICMSInterestVlrUFDest = objTribDocOriginal.dICMSInterestVlrUFDest
    dICMSInterestVlrUFRemet = objTribDocOriginal.dICMSInterestVlrUFRemet

    'nfe 4.00
    dICMSVlrFCP = objTribDocOriginal.dICMSVlrFCP
    dICMSVlrFCPST = objTribDocOriginal.dICMSVlrFCPST
    dICMSVlrFCPSTRet = objTribDocOriginal.dICMSVlrFCPSTRet
    dIPIVlrDevolvido = objTribDocOriginal.dIPIVlrDevolvido

    Exit Sub

Erro_Copia2:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202899)

    End Select

    Exit Sub
    
End Sub

Public Function ObterVar_Padrao(ByVal sIdentificador As String, ByVal objVariavelCalculo As ClassVariavelCalculo, Optional ByVal objClassP As Object = Nothing) As Long

Dim lErro As Long, objProduto As New ClassProduto
Dim objClass As Object
Dim vValorAux As Variant

On Error GoTo Erro_ObterVar_Padrao

    If objClassP Is Nothing Then
        Set objClass = Me
    Else
        Set objClass = objClassP
    End If

    Select Case sIdentificador

        Case "ISS_MUNIC_FATURAMENTO"
            objVariavelCalculo.vValor = objClass.sISSMunicFaturamento
            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO
        
        Case "NAT_OPERACAO"
            objVariavelCalculo.vValor = objClass.sNaturezaOp
            objVariavelCalculo.iStatus = objClass.iNaturezaOpManual
                
        Case "TIPO_TRIB"
            objVariavelCalculo.vValor = objClass.iTipoTributacao
            objVariavelCalculo.iStatus = objClass.iTipoTributacaoManual
            
        'para ICMS
        Case "ICMS_BASE"
            objVariavelCalculo.vValor = objClass.dICMSBase
            objVariavelCalculo.iStatus = objClass.iICMSBaseManual
        
        Case "ICMS_VALOR"
            objVariavelCalculo.vValor = objClass.dICMSValor
            objVariavelCalculo.iStatus = objClass.iICMSValorManual
        
        Case "ICMS_SUBST_BASE"
            objVariavelCalculo.vValor = objClass.dICMSSubstBase
            objVariavelCalculo.iStatus = objClass.iICMSSubstBaseManual
        
        Case "ICMS_SUBST_VALOR"
            objVariavelCalculo.vValor = objClass.dICMSSubstValor
            objVariavelCalculo.iStatus = objClass.iICMSSubstValorManual

        Case "ICMS_VALOR_CREDITO"
            objVariavelCalculo.vValor = objClass.dICMSCredito
            objVariavelCalculo.iStatus = objClass.iICMSCreditoManual
        
        'para IPI
        Case "IPI_BASE"
            objVariavelCalculo.vValor = objClass.dIPIBase
            objVariavelCalculo.iStatus = objClass.iIPIBaseManual

        Case "IPI_VALOR"
            objVariavelCalculo.vValor = objClass.dIPIValor
            objVariavelCalculo.iStatus = objClass.iIPIValorManual

        Case "IPI_VALOR_CREDITO"
            objVariavelCalculo.vValor = objClass.dIPICredito
            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO
        
        'para ISS
        Case "ISS_INCLUSO"
            objVariavelCalculo.vValor = objClass.iISSIncluso
            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO
        
        Case "ISS_BASE"
            objVariavelCalculo.vValor = objClass.dISSBase
            objVariavelCalculo.iStatus = objClass.iISSBaseManual

        Case "ISS_ALIQUOTA"
            objVariavelCalculo.vValor = objClass.dISSAliquota
            objVariavelCalculo.iStatus = objClass.iISSAliquotaManual

        Case "ISS_VALOR"
            objVariavelCalculo.vValor = objClass.dISSValor
            objVariavelCalculo.iStatus = objClass.iISSValorManual

        'para IR
        Case "IR_BASE"
            objVariavelCalculo.vValor = objClass.dIRRFBase
            objVariavelCalculo.iStatus = objClass.iIRRFBaseManual
            
        Case "IR_ALIQUOTA"
            objVariavelCalculo.vValor = objClass.dIRRFAliquota
            objVariavelCalculo.iStatus = objClass.iIRRFAliquotaManual
            
        Case "IR_VALOR"
            objVariavelCalculo.vValor = objClass.dIRRFValor
            objVariavelCalculo.iStatus = objClass.iIRRFValorManual
            
        'para INSS
        Case "INSS_BASE"
            objVariavelCalculo.vValor = objClass.dINSSBase
            objVariavelCalculo.iStatus = objClass.iINSSBaseManual
        
        Case "INSS_VALOR_DEDUCOES"
            objVariavelCalculo.vValor = objClass.dINSSDeducoes
            objVariavelCalculo.iStatus = objClass.iINSSDeducoesManual
        
        Case "INSS_VALOR"
            objVariavelCalculo.vValor = objClass.dINSSValor
            objVariavelCalculo.iStatus = objClass.iINSSValorManual

        Case "INSS_RETIDO"
            objVariavelCalculo.vValor = objClass.iINSSRetido
            objVariavelCalculo.iStatus = objClass.iINSSRetidoManual
            
        Case "PIS_VALOR_CREDITO"
            objVariavelCalculo.vValor = objClass.dPISCredito
            objVariavelCalculo.iStatus = objClass.iPISCreditoManual
        
        Case "COFINS_VALOR_CREDITO"
            objVariavelCalculo.vValor = objClass.dCOFINSCredito
            objVariavelCalculo.iStatus = objClass.iCOFINSCreditoManual
        
        Case "PIS_RETIDO"
            objVariavelCalculo.vValor = objClass.dPISRetido
            objVariavelCalculo.iStatus = objClass.iPISRetidoManual
            
        Case "ISS_RETIDO"
            objVariavelCalculo.vValor = objClass.dISSRetido
            objVariavelCalculo.iStatus = objClass.iISSRetidoManual
            
        Case "COFINS_RETIDO"
            objVariavelCalculo.vValor = objClass.dCOFINSRetido
            objVariavelCalculo.iStatus = objClass.iCOFINSRetidoManual
            
        Case "CSLL_RETIDO"
            objVariavelCalculo.vValor = objClass.dCSLLRetido
            objVariavelCalculo.iStatus = objClass.iCSLLRetidoManual
            
        Case "BASE_ICMS_IMPORTACAO_SIMPLES"
            objVariavelCalculo.vValor = objClass.dIPIBase + objClass.dIPIValor + objClass.dValorOutrasDespesas
            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO
            
        Case "PIS_TIPO_OBJ"
            Set vValorAux = objVariavelCalculo.vValor
            Call objClass.PISCOFINSTipo_Obter(vValorAux)
            Set objVariavelCalculo.vValor = vValorAux
            
        Case "COFINS_TIPO_OBJ"
            Set vValorAux = objVariavelCalculo.vValor
            Call objClass.PISCOFINSTipo_Obter(vValorAux)
            Set objVariavelCalculo.vValor = vValorAux
        
        Case "ICMS_SIMPLES_TIPO_OBJ"
            Set vValorAux = objVariavelCalculo.vValor
            Call objClass.ICMSSimplesTipo_Obter(vValorAux)
            Set objVariavelCalculo.vValor = vValorAux
            
        Case "ISS_TIPO_OBJ"
            Set vValorAux = objVariavelCalculo.vValor
            Call objClass.ISSTipo_Obter(vValorAux)
            Set objVariavelCalculo.vValor = vValorAux
        
        Case "TOT_TRIB"
            objVariavelCalculo.vValor = objClass.dTotTrib
            objVariavelCalculo.iStatus = objClass.iTotTribManual

        Case "TOT_TRIB_TIPO"
            objVariavelCalculo.vValor = objClass.iTotTribTipo
            objVariavelCalculo.iStatus = objClass.iTotTribTipoManual
        
        Case "TOT_TRIB_FED"
            objVariavelCalculo.vValor = objClass.dTotTribFed
            objVariavelCalculo.iStatus = objClass.iTotTribFedManual
        
        Case "TOT_TRIB_EST"
            objVariavelCalculo.vValor = objClass.dTotTribEst
            objVariavelCalculo.iStatus = objClass.iTotTribEstManual
        
        Case "TOT_TRIB_MUNIC"
            objVariavelCalculo.vValor = objClass.dTotTribMunic
            objVariavelCalculo.iStatus = objClass.iTotTribMunicManual
        
        'nfe 3.10
        Case "PIS_VALOR"
            objVariavelCalculo.vValor = objClass.dPISValor
            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO

        Case "COFINS_VALOR"
            objVariavelCalculo.vValor = objClass.dCOFINSValor
            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO

        Case "ICMS_VALOR_ISENTO"
            objVariavelCalculo.vValor = objClass.dICMSValorIsento
            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO

        Case "ISS_VALOR_DEDUCAO"
            objVariavelCalculo.vValor = objClass.dISSValorDeducao
            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO

        Case "ISS_VALOR_OUTRAS_RET"
            objVariavelCalculo.vValor = objClass.dISSValorOutrasRet
            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO

        Case "ISS_VALOR_DESC_INCOND"
            objVariavelCalculo.vValor = objClass.dISSValorDescIncond
            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO

        Case "ISS_VALOR_DESC_COND"
            objVariavelCalculo.vValor = objClass.dISSValorDescCond
            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO

        Case "DESTINO"
            objVariavelCalculo.vValor = objClass.iDestino
            objVariavelCalculo.iStatus = objClass.iDestinoManual
        
        Case "FINALIDADE_NFE"
            objVariavelCalculo.vValor = objClass.iFinalidadeNFe
            objVariavelCalculo.iStatus = objClass.iFinalidadeNFeManual
        
        Case "IND_CONSUMIDOR_FINAL"
            objVariavelCalculo.vValor = objClass.iIndConsumidorFinal
            objVariavelCalculo.iStatus = objClass.iIndConsumidorFinalManual
        
        Case "IND_PRESENCA"
            objVariavelCalculo.vValor = objClass.iIndPresenca
            objVariavelCalculo.iStatus = objClass.iIndPresencaManual
        
        Case "REG_TRIB_ESPECIAL"
            objVariavelCalculo.vValor = objClass.iRegTribEspecial
            objVariavelCalculo.iStatus = objClass.iRegTribEspecialManual
        
        Case "DATA_PREST_SERVICO"
            objVariavelCalculo.vValor = objClass.dtDataPrestServico
            objVariavelCalculo.iStatus = objClass.iDataPrestServicoManual
        
        Case "ICMS_INTEREST_VLR_FCP_UF_DEST"
            objVariavelCalculo.vValor = objClass.dICMSInterestVlrFCPUFDest
            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO

        Case "ICMS_INTEREST_VLR_UF_DEST"
            objVariavelCalculo.vValor = objClass.dICMSInterestVlrUFDest
            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO

        Case "ICMS_INTEREST_VLR_UF_REMET"
            objVariavelCalculo.vValor = objClass.dICMSInterestVlrUFRemet
            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO
            
        'nfe 4.00
        Case "ICMS_VLR_FCP"
            objVariavelCalculo.vValor = objClass.dICMSVlrFCP
            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO
            
        Case "ICMS_VLR_FCP_ST"
            objVariavelCalculo.vValor = objClass.dICMSVlrFCPST
            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO
            
        Case "ICMS_VLR_FCP_RET"
            objVariavelCalculo.vValor = objClass.dICMSVlrFCPSTRet
            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO
            
        Case "IPI_VLR_DEVOLVIDO"
            objVariavelCalculo.vValor = objClass.dIPIVlrDevolvido
            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO
            
        Case Else
            gError 202850
        
    End Select

    ObterVar_Padrao = SUCESSO

    Exit Function

Erro_ObterVar_Padrao:

    ObterVar_Padrao = gErr

    Select Case gErr
    
        Case 202850
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202851)

    End Select

    Exit Function

End Function

Public Function GuardarVar_Padrao(ByVal objVariavelCalculo As ClassVariavelCalculo, Optional ByVal objClassP As Object = Nothing) As Long

Dim lErro As Long, objProduto As New ClassProduto
Dim objClass As Object

On Error GoTo Erro_GuardarVar_Padrao

    If objClassP Is Nothing Then
        Set objClass = Me
    Else
        Set objClass = objClassP
    End If
    
    Select Case objVariavelCalculo.sIdentificador
    
        Case "ISS_MUNIC_FATURAMENTO"
            objClass.sISSMunicFaturamento = objVariavelCalculo.vValor
        
        Case "NAT_OPERACAO"
            objClass.sNaturezaOp = objVariavelCalculo.vValor
            objClass.iNaturezaOpManual = objVariavelCalculo.iStatus
            
        Case "TIPO_TRIB"
            objClass.iTipoTributacao = objVariavelCalculo.vValor
            objClass.iTipoTributacaoManual = objVariavelCalculo.iStatus

        'para ICMS
        Case "ICMS_BASE"
            objClass.dICMSBase = objVariavelCalculo.vValor
            objClass.iICMSBaseManual = objVariavelCalculo.iStatus
        
        Case "ICMS_VALOR"
            objClass.dICMSValor = objVariavelCalculo.vValor
            objClass.iICMSValorManual = objVariavelCalculo.iStatus
        
        Case "ICMS_SUBST_BASE"
            objClass.dICMSSubstBase = objVariavelCalculo.vValor
            objClass.iICMSSubstBaseManual = objVariavelCalculo.iStatus
        
        Case "ICMS_SUBST_VALOR"
            objClass.dICMSSubstValor = objVariavelCalculo.vValor
            objClass.iICMSSubstValorManual = objVariavelCalculo.iStatus
        
        Case "ICMS_VALOR_CREDITO"
            objClass.dICMSCredito = objVariavelCalculo.vValor
            objClass.iICMSCreditoManual = objVariavelCalculo.iStatus
       
        'para IPI
        Case "IPI_BASE"
            objClass.dIPIBase = objVariavelCalculo.vValor
            objClass.iIPIBaseManual = objVariavelCalculo.iStatus
        
        Case "IPI_VALOR"
            objClass.dIPIValor = objVariavelCalculo.vValor
            objClass.iIPIValorManual = objVariavelCalculo.iStatus
        
        Case "IPI_VALOR_CREDITO"
            objClass.dIPICredito = objVariavelCalculo.vValor
            objClass.iIPICreditoManual = objVariavelCalculo.iStatus
        
        'para ISS
        Case "ISS_INCLUSO"
            objClass.iISSIncluso = objVariavelCalculo.vValor
            'objClass.iISSInclusoManual = objVariavelCalculo.iStatus
        
        Case "ISS_BASE"
            objClass.dISSBase = objVariavelCalculo.vValor
            objClass.iISSBaseManual = objVariavelCalculo.iStatus
        
        Case "ISS_ALIQUOTA"
            objClass.dISSAliquota = objVariavelCalculo.vValor
            objClass.iISSAliquotaManual = objVariavelCalculo.iStatus
        
        Case "ISS_VALOR"
            objClass.dISSValor = objVariavelCalculo.vValor
            objClass.iISSValorManual = objVariavelCalculo.iStatus
        
        'para IR
        Case "IR_BASE"
            objClass.dIRRFBase = objVariavelCalculo.vValor
            objClass.iIRRFBaseManual = objVariavelCalculo.iStatus
        
        Case "IR_ALIQUOTA"
            objClass.dIRRFAliquota = objVariavelCalculo.vValor
            objClass.iIRRFAliquotaManual = objVariavelCalculo.iStatus
        
        Case "IR_VALOR"
            objClass.dIRRFValor = objVariavelCalculo.vValor
            objClass.iIRRFValorManual = objVariavelCalculo.iStatus
            
        'para INSS
        Case "INSS_BASE"
            objClass.dINSSBase = objVariavelCalculo.vValor
            objClass.iINSSBaseManual = objVariavelCalculo.iStatus

        Case "INSS_VALOR_DEDUCOES"
            objClass.dINSSDeducoes = objVariavelCalculo.vValor
            objClass.iINSSDeducoesManual = objVariavelCalculo.iStatus

        Case "INSS_VALOR"
            objClass.dINSSValor = objVariavelCalculo.vValor
            objClass.iINSSValorManual = objVariavelCalculo.iStatus

        Case "INSS_RETIDO"
            objClass.iINSSRetido = objVariavelCalculo.vValor
            objClass.iINSSRetidoManual = objVariavelCalculo.iStatus
        
        Case "PIS_VALOR_CREDITO"
            objClass.dPISCredito = objVariavelCalculo.vValor
            objClass.iPISCreditoManual = objVariavelCalculo.iStatus
        
        Case "COFINS_VALOR_CREDITO"
            objClass.dCOFINSCredito = objVariavelCalculo.vValor
            objClass.iCOFINSCreditoManual = objVariavelCalculo.iStatus
        
        Case "PIS_RETIDO"
            objClass.dPISRetido = objVariavelCalculo.vValor
            objClass.iPISRetidoManual = objVariavelCalculo.iStatus
        
        Case "ISS_RETIDO"
            objClass.dISSRetido = objVariavelCalculo.vValor
            objClass.iISSRetidoManual = objVariavelCalculo.iStatus
        
        Case "COFINS_RETIDO"
            objClass.dCOFINSRetido = objVariavelCalculo.vValor
            objClass.iCOFINSRetidoManual = objVariavelCalculo.iStatus
        
        Case "CSLL_RETIDO"
            objClass.dCSLLRetido = objVariavelCalculo.vValor
            objClass.iCSLLRetidoManual = objVariavelCalculo.iStatus

        Case "DESCONTO_VALOR"
            objClass.dValorDesconto = objVariavelCalculo.vValor
        
        Case "FRETE_VALOR"
            objClass.dValorFrete = objVariavelCalculo.vValor
        
        Case "SEGURO_VALOR"
            objClass.dValorSeguro = objVariavelCalculo.vValor
        
        Case "DESP_ACESS_VALOR"
            objClass.dValorOutrasDespesas = objVariavelCalculo.vValor
            
        Case "TOT_TRIB"
            objClass.dTotTrib = objVariavelCalculo.vValor
            objClass.iTotTribManual = objVariavelCalculo.iStatus
        
        Case "TOT_TRIB_FED"
            objClass.dTotTribFed = objVariavelCalculo.vValor
            objClass.iTotTribFedManual = objVariavelCalculo.iStatus
        
        Case "TOT_TRIB_EST"
            objClass.dTotTribEst = objVariavelCalculo.vValor
            objClass.iTotTribEstManual = objVariavelCalculo.iStatus
        
        Case "TOT_TRIB_MUNIC"
            objClass.dTotTribMunic = objVariavelCalculo.vValor
            objClass.iTotTribMunicManual = objVariavelCalculo.iStatus
        
        Case "TOT_TRIB_TIPO"
            objClass.iTotTribTipo = objVariavelCalculo.vValor
            objClass.iTotTribTipoManual = objVariavelCalculo.iStatus
        
        'nfe 3.10
        Case "PIS_VALOR"
            objClass.dPISValor = objVariavelCalculo.vValor

        Case "COFINS_VALOR"
            objClass.dCOFINSValor = objVariavelCalculo.vValor

        Case "ICMS_VALOR_ISENTO"
            objClass.dICMSValorIsento = objVariavelCalculo.vValor

        Case "ISS_VALOR_DEDUCAO"
            objClass.dISSValorDeducao = objVariavelCalculo.vValor

        Case "ISS_VALOR_OUTRAS_RET"
            objClass.dISSValorOutrasRet = objVariavelCalculo.vValor

        Case "ISS_VALOR_DESC_INCOND"
            objClass.dISSValorDescIncond = objVariavelCalculo.vValor

        Case "ISS_VALOR_DESC_COND"
            objClass.dISSValorDescCond = objVariavelCalculo.vValor

        Case "DESTINO"
            objClass.iDestino = objVariavelCalculo.vValor
            objClass.iDestinoManual = objVariavelCalculo.iStatus
        
        Case "FINALIDADE_NFE"
            objClass.iFinalidadeNFe = objVariavelCalculo.vValor
            objClass.iFinalidadeNFeManual = objVariavelCalculo.iStatus
        
        Case "IND_CONSUMIDOR_FINAL"
            objClass.iIndConsumidorFinal = objVariavelCalculo.vValor
            objClass.iIndConsumidorFinalManual = objVariavelCalculo.iStatus
        
        Case "IND_PRESENCA"
            objClass.iIndPresenca = objVariavelCalculo.vValor
            objClass.iIndPresencaManual = objVariavelCalculo.iStatus
        
        Case "REG_TRIB_ESPECIAL"
            objClass.iRegTribEspecial = objVariavelCalculo.vValor
            objClass.iRegTribEspecialManual = objVariavelCalculo.iStatus
        
        Case "DATA_PREST_SERVICO"
            objClass.dtDataPrestServico = objVariavelCalculo.vValor
            objClass.iDataPrestServicoManual = objVariavelCalculo.iStatus
        
        Case "ICMS_INTEREST_VLR_FCP_UF_DEST"
            objClass.dICMSInterestVlrFCPUFDest = objVariavelCalculo.vValor

        Case "ICMS_INTEREST_VLR_UF_DEST"
            objClass.dICMSInterestVlrUFDest = objVariavelCalculo.vValor

        Case "ICMS_INTEREST_VLR_UF_REMET"
            objClass.dICMSInterestVlrUFRemet = objVariavelCalculo.vValor
        
        'nfe 4.00
        Case "ICMS_VLR_FCP"
            objClass.dICMSVlrFCP = objVariavelCalculo.vValor
            
        Case "ICMS_VLR_FCP_ST"
            objClass.dICMSVlrFCPST = objVariavelCalculo.vValor
            
        Case "ICMS_VLR_FCP_RET"
            objClass.dICMSVlrFCPSTRet = objVariavelCalculo.vValor
            
        Case "IPI_VLR_DEVOLVIDO"
            objClass.dIPIVlrDevolvido = objVariavelCalculo.vValor
            
        Case Else
            gError 202852
        
    End Select

    GuardarVar_Padrao = SUCESSO

    Exit Function

Erro_GuardarVar_Padrao:

    GuardarVar_Padrao = gErr

    Select Case gErr
    
        Case 202852
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202853)

    End Select

    Exit Function

End Function

Public Property Let sNaturezaOpInterna(ByVal vData As String)
    mvarsNaturezaOpInterna = vData
End Property

Public Property Get sNaturezaOpInterna() As String
    sNaturezaOpInterna = mvarsNaturezaOpInterna
End Property

Public Property Let sNaturezaOp(ByVal vData As String)
    mvarsNaturezaOpInterna = vData
End Property

Public Property Get sNaturezaOp() As String
    sNaturezaOp = mvarsNaturezaOpInterna
End Property

Public Property Let iNaturezaOpManual(ByVal vData As Integer)
    mvariNaturezaOpManual = vData
End Property

Public Property Get iNaturezaOpManual() As Integer
    iNaturezaOpManual = mvariNaturezaOpManual
End Property

Public Property Let dINSSBase(ByVal vData As Double)
    mvardINSSValorBase = vData
End Property

Public Property Get dINSSBase() As Double
    dINSSBase = mvardINSSValorBase
End Property

Public Property Let iINSSBaseManual(ByVal vData As Integer)
    mvariINSSValorBaseManual = vData
End Property

Public Property Get iINSSBaseManual() As Integer
    iINSSBaseManual = mvariINSSValorBaseManual
End Property

Public Property Let dINSSDeducoes(ByVal vData As Double)
    mvardINSSValorDeducoes = vData
End Property

Public Property Get dINSSDeducoes() As Double
    dINSSDeducoes = mvardINSSValorDeducoes
End Property

Public Property Let iINSSDeducoesManual(ByVal vData As Integer)
    mvariINSSValorDeducoesManual = vData
End Property

Public Property Get iINSSDeducoesManual() As Integer
    iINSSDeducoesManual = mvariINSSValorDeducoesManual
End Property

Public Property Let dINSSValor(ByVal vData As Double)
    mvardValorINSS = vData
End Property

Public Property Get dINSSValor() As Double
    dINSSValor = mvardValorINSS
End Property

Public Property Let iINSSValorManual(ByVal vData As Integer)
    mvariValorINSSManual = vData
End Property

Public Property Get iINSSValorManual() As Integer
    iINSSValorManual = mvariValorINSSManual
End Property

Public Property Get dTotalRetencoes() As Double
    dTotalRetencoes = mvardIRRFValor + mvardPISRetido + mvardCOFINSRetido + mvardCSLLRetido + IIf(mvariINSSRetido, mvardValorINSS, 0) + mvardISSRetido
End Property

Public Property Set colTiposTribICMS(ByVal vData As Object)
    Set mvarcolTiposTribICMS = vData
End Property

Public Property Get colTiposTribICMS() As Collection
    Set colTiposTribICMS = mvarcolTiposTribICMS
End Property

Public Property Set colTiposTribICMSSimples(ByVal vData As Object)
    Set mvarcolTiposTribICMSSimples = vData
End Property

Public Property Get colTiposTribICMSSimples() As Collection
    Set colTiposTribICMSSimples = mvarcolTiposTribICMSSimples
End Property

Public Property Set colTiposTribISS(ByVal vData As Object)
    Set mvarcolTiposTribISS = vData
End Property

Public Property Get colTiposTribISS() As Collection
    Set colTiposTribISS = mvarcolTiposTribISS
End Property

Public Property Set colTiposTribIPI(ByVal vData As Object)
    Set mvarcolTiposTribIPI = vData
End Property

Public Property Get colTiposTribIPI() As Collection
    Set colTiposTribIPI = mvarcolTiposTribIPI
End Property

Public Property Set colTiposTribPISCOFINS(ByVal vData As Object)
    Set mvarcolTiposTribPISCOFINS = vData
End Property

Public Property Get colTiposTribPISCOFINS() As Collection
    Set colTiposTribPISCOFINS = mvarcolTiposTribPISCOFINS
End Property

Public Sub PISCOFINSTipo_Obter(vValor As Variant)
    Dim iCodigo As Integer
    iCodigo = vValor.iTipo
    For Each vValor In colTiposTribPISCOFINS
        If iCodigo = vValor.iTipo Then Exit For
    Next
End Sub

Public Sub ISSTipo_Obter(vValor As Variant)
    Dim iCodigo As Integer
    iCodigo = vValor.iTipo
    For Each vValor In colTiposTribISS
        If iCodigo = vValor.iTipo Then Exit For
    Next
End Sub

Public Sub ICMSSimplesTipo_Obter(vValor As Variant)
    Dim iCodigo As Integer
    iCodigo = vValor.iTipo
    For Each vValor In colTiposTribICMSSimples
        If iCodigo = vValor.iTipo Then Exit For
    Next
End Sub

Private Function Altera_VAR_PREENCH(ByVal iStatus As Integer)
    iTipoTributacaoManual = iStatus
    iIPIBaseManual = iStatus
    iIPIValorManual = iStatus
    iICMSBaseManual = iStatus
    iICMSValorManual = iStatus
    iICMSSubstBaseManual = iStatus
    iICMSSubstValorManual = iStatus
    iISSBaseManual = iStatus
    iISSAliquotaManual = iStatus
    iISSValorManual = iStatus
    iIRRFBaseManual = iStatus
    iIRRFAliquotaManual = iStatus
    iIRRFValorManual = iStatus
    iPISRetidoManual = iStatus
    iCOFINSRetidoManual = iStatus
    iCSLLRetidoManual = iStatus
    iISSRetidoManual = iStatus
    iPISCreditoManual = iStatus
    iCOFINSCreditoManual = iStatus
    iIPICreditoManual = iStatus
    iICMSCreditoManual = iStatus
    iValorINSSManual = iStatus
    iINSSRetidoManual = iStatus
    iINSSValorBaseManual = iStatus
    iINSSValorDeducoesManual = iStatus
    iNaturezaOpManual = iStatus
    iTotTribManual = iStatus
    iTotTribTipoManual = iStatus
    
    'nfe 3.10
    iDestinoManual = iStatus
    iFinalidadeNFeManual = iStatus
    iIndConsumidorFinalManual = iStatus
    iIndPresencaManual = iStatus
    iRegTribEspecialManual = iStatus
    iDataPrestServicoManual = iStatus
    
    'IBPTax 0.0.9
    iTotTribFedManual = iStatus
    iTotTribEstManual = iStatus
    iTotTribMunicManual = iStatus
    
End Function

Public Sub Coloca_Manual()
    Call Altera_VAR_PREENCH(VAR_PREENCH_MANUAL)
End Sub

Public Sub Coloca_Auto()
    Call Altera_VAR_PREENCH(VAR_PREENCH_AUTOMATICO)
End Sub

Public Property Let dValorSeguro(ByVal vData As Double)
    mvardValorSeguro = vData
End Property

Public Property Get dValorSeguro() As Double
    dValorSeguro = mvardValorSeguro
End Property

Public Property Let dValorDesconto(ByVal vData As Double)
    mvardValorDesconto = vData
End Property

Public Property Get dValorDesconto() As Double
    dValorDesconto = mvardValorDesconto
End Property

Public Property Let dValorFrete(ByVal vData As Double)
    mvardValorFrete = vData
End Property

Public Property Get dValorFrete() As Double
    dValorFrete = mvardValorFrete
End Property

Public Property Let dValorOutrasDespesas(ByVal vData As Double)
    mvardValorOutrasDespesas = vData
End Property

Public Property Get dValorOutrasDespesas() As Double
    dValorOutrasDespesas = mvardValorOutrasDespesas
End Property

Private Sub Class_Initialize()
    dtDataPrestServico = DATA_NULA
End Sub

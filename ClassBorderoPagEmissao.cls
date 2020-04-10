VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassBorderoPagEmissao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"colInfoParcPag"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariCta As Integer
Private mvarlNumero As Long
Private mvarlNumeroInt As Long
Private mvardtEmissao As Date
Private mvardtContabil As Date
Private mvardtVenctoDe As Date
Private mvardtVencto As Date
Private mvardValorMaximo As Double
Private mvariTipoCobranca As Integer
Private mvariLiqTitOutroBco As Integer
Private mvariPodeDepCtaOutroBco As Integer
Private mvariQtdeParcelasSelecionadas As Integer
Private mvardValorParcelasSelecionadas As Double
Private mvarcolInfoParcPag As New Collection
Private mvariPodeDOCOutroEstado As Integer
Private mvariBcoCta As Integer
Private mvarobjEvolucao As Object
Private mvarobjTelaAtualizacao As Object 'tela de acompanhamento da atualizacao de arquivos. Necessaria para a contabilizacao.

Private mvardValorBaixado As Double
Private mvardValorMulta As Double
Private mvardValorJuros As Double
Private mvardValorDesconto As Double
Private mvaradValorDiferenca() As Double

Private mvarobjColBaixaParcPagar As ColBaixaParcPagar

Public Property Get objColBaixaParcPagar() As ColBaixaParcPagar
    Set objColBaixaParcPagar = mvarobjColBaixaParcPagar
End Property

Public Property Set objColBaixaParcPagar(ByVal vData As ColBaixaParcPagar)
    Set mvarobjColBaixaParcPagar = vData
End Property

Public Property Get adValorDiferenca_UBound() As Integer
    adValorDiferenca_UBound = UBound(mvaradValorDiferenca)
End Property

Public Property Get adValorDiferenca_LBound() As Integer
    adValorDiferenca_LBound = LBound(mvaradValorDiferenca)
End Property

Public Property Let adValorDiferenca(ByVal vIndex As Variant, ByVal vData As Double)
    If vIndex > UBound(mvaradValorDiferenca) Then ReDim Preserve mvaradValorDiferenca(vIndex + 1)
    mvaradValorDiferenca(vIndex) = vData
End Property

Public Property Get adValorDiferenca(ByVal vIndex As Variant) As Double
    adValorDiferenca = mvaradValorDiferenca(vIndex)
End Property

Public Property Let dValorBaixado(ByVal vData As Double)
    mvardValorBaixado = vData
End Property

Public Property Get dValorBaixado() As Double
    dValorBaixado = mvardValorBaixado
End Property

Public Property Let dValorMulta(ByVal vData As Double)
    mvardValorMulta = vData
End Property

Public Property Get dValorMulta() As Double
    dValorMulta = mvardValorMulta
End Property

Public Property Let dValorJuros(ByVal vData As Double)
    mvardValorJuros = vData
End Property

Public Property Get dValorJuros() As Double
    dValorJuros = mvardValorJuros
End Property

Public Property Let dValorDesconto(ByVal vData As Double)
    mvardValorDesconto = vData
End Property

Public Property Get dValorDesconto() As Double
    dValorDesconto = mvardValorDesconto
End Property


Public Property Let objTelaAtualizacao(ByVal vData As Object)
    Set mvarobjTelaAtualizacao = vData
End Property

Public Property Get objTelaAtualizacao() As Object
    Set objTelaAtualizacao = mvarobjTelaAtualizacao
End Property

Public Property Let objEvolucao(ByVal vData As Object)
    Set mvarobjEvolucao = vData
End Property

Public Property Get objEvolucao() As Object
    Set objEvolucao = mvarobjEvolucao
End Property

Public Property Set objEvolucao(ByVal vData As Object)
    Set mvarobjEvolucao = vData
End Property

Public Property Let iBcoCta(ByVal vData As Integer)
    mvariBcoCta = vData
End Property

Public Property Get iBcoCta() As Integer
    iBcoCta = mvariBcoCta
End Property

Public Property Let iPodeDOCOutroEstado(ByVal vData As Integer)
    mvariPodeDOCOutroEstado = vData
End Property

Public Property Get iPodeDOCOutroEstado() As Integer
    iPodeDOCOutroEstado = mvariPodeDOCOutroEstado
End Property

Public Property Get colInfoParcPag() As Collection
    Set colInfoParcPag = mvarcolInfoParcPag
End Property

Public Property Let dValorParcelasSelecionadas(ByVal vData As Double)
    mvardValorParcelasSelecionadas = vData
End Property

Public Property Get dValorParcelasSelecionadas() As Double
    dValorParcelasSelecionadas = mvardValorParcelasSelecionadas
End Property

Public Property Let iQtdeParcelasSelecionadas(ByVal vData As Integer)
    mvariQtdeParcelasSelecionadas = vData
End Property

Public Property Get iQtdeParcelasSelecionadas() As Integer
    iQtdeParcelasSelecionadas = mvariQtdeParcelasSelecionadas
End Property

Public Property Let iPodeDepCtaOutroBco(ByVal vData As Integer)
    mvariPodeDepCtaOutroBco = vData
End Property

Public Property Get iPodeDepCtaOutroBco() As Integer
    iPodeDepCtaOutroBco = mvariPodeDepCtaOutroBco
End Property

Public Property Let iLiqTitOutroBco(ByVal vData As Integer)
    mvariLiqTitOutroBco = vData
End Property

Public Property Get iLiqTitOutroBco() As Integer
    iLiqTitOutroBco = mvariLiqTitOutroBco
End Property

Public Property Let iTipoCobranca(ByVal vData As Integer)
    mvariTipoCobranca = vData
End Property

Public Property Get iTipoCobranca() As Integer
    iTipoCobranca = mvariTipoCobranca
End Property

Public Property Let dValorMaximo(ByVal vData As Double)
    mvardValorMaximo = vData
End Property

Public Property Get dValorMaximo() As Double
    dValorMaximo = mvardValorMaximo
End Property

Public Property Let dtVencto(ByVal vData As Date)
    mvardtVencto = vData
End Property

Public Property Get dtVencto() As Date
    dtVencto = mvardtVencto
End Property

Public Property Let dtVenctoDe(ByVal vData As Date)
    mvardtVenctoDe = vData
End Property

Public Property Get dtVenctoDe() As Date
    dtVenctoDe = mvardtVenctoDe
End Property

Public Property Let dtContabil(ByVal vData As Date)
    mvardtContabil = vData
End Property

Public Property Get dtContabil() As Date
    dtContabil = mvardtContabil
End Property

Public Property Let dtEmissao(ByVal vData As Date)
    mvardtEmissao = vData
End Property

Public Property Get dtEmissao() As Date
    dtEmissao = mvardtEmissao
End Property

Public Property Let lNumero(ByVal vData As Long)
    mvarlNumero = vData
End Property

Public Property Get lNumero() As Long
    lNumero = mvarlNumero
End Property

Public Property Let lNumeroInt(ByVal vData As Long)
    mvarlNumeroInt = vData
End Property

Public Property Get lNumeroInt() As Long
    lNumeroInt = mvarlNumeroInt
End Property

Public Property Let iCta(ByVal vData As Integer)
    mvariCta = vData
End Property

Public Property Get iCta() As Integer
    iCta = mvariCta
End Property

Function GeraContabilizacao(objContabAutomatica As ClassContabAutomatica, vParams As Variant) As Long

    GeraContabilizacao = mvarobjTelaAtualizacao.GeraContabilizacao(objContabAutomatica, vParams)
    
End Function

Private Sub Class_Initialize()
    ReDim Preserve mvaradValorDiferenca(1)
End Sub

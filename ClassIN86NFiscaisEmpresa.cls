VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassIN86NFiscaisEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarsIndicadorMov As String
Private mvariModeloDoc As Integer
Private mvarsSerie As String
Private mvarlNumeroDoc As Long
Private mvardtDataEmissao As Date
Private mvardValorMercadoria As Double
Private mvardValorTotDesconto As Double
Private mvardValorFrete As Double
Private mvardValorSeguro As Double
Private mvardValorOutrasDesp As Double
Private mvardValorTotalIPI As Double
Private mvardValorTotalICMSSubsTRB As Double
Private mvardValorTotalNF As Double
Private mvarsInscEstSubsTRB As String
Private mvariViaTransporte As String
Private mvariTransportador As String
Private mvariQtdVolumes As Integer
Private mvarsEspecieVolume As String
Private mvardPesoBruto As Double
Private mvardPesoLiquido As Double
Private mvariModalidadeFrete As String
Private mvarsIDVeiculo As String
Private mvarsTipoFatura As String
Private mvarsObservacao As String
Private mvarsIndicadorMovCMP As String
Private mvariEmitente As Integer
Private mvariDestinatario As Integer
Private mvariOrigem As Integer
Private mvarlFornecedor As Long
Private mvariFilialForn As Integer
Private mvarlCliente As Long
Private mvariFilialCliente As Integer
Private mvardtDataEntrada As Date
Private mvardtDataSaida As Date
Private mvariStatus As Integer
Private mvariFilialEmpresa As Integer
Private mvariClasseDocCPR As Integer
Private mvarlNumIntDocCPR As Long
Private mvariRemetente As Integer
Private mvariTipoTipoDocInfo As Integer 'Incluído por Luiz Nogueira em 28/01/04

Public Property Let lNumIntDocCPR(ByVal vData As Long)
    mvarlNumIntDocCPR = vData
End Property

Public Property Get lNumIntDocCPR() As Long
    lNumIntDocCPR = mvarlNumIntDocCPR
End Property

Public Property Let iClasseDocCPR(ByVal vData As Integer)
    mvariClasseDocCPR = vData
End Property

Public Property Get iClasseDocCPR() As Integer
    iClasseDocCPR = mvariClasseDocCPR
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let iStatus(ByVal vData As Integer)
    mvariStatus = vData
End Property

Public Property Get iStatus() As Integer
    iStatus = mvariStatus
End Property

Public Property Let dtDataSaida(ByVal vData As Date)
    mvardtDataSaida = vData
End Property

Public Property Get dtDataSaida() As Date
    dtDataSaida = mvardtDataSaida
End Property

Public Property Let dtDataEntrada(ByVal vData As Date)
    mvardtDataEntrada = vData
End Property

Public Property Get dtDataEntrada() As Date
    dtDataEntrada = mvardtDataEntrada
End Property

Public Property Let iFilialCliente(ByVal vData As Integer)
    mvariFilialCliente = vData
End Property

Public Property Get iFilialCliente() As Integer
    iFilialCliente = mvariFilialCliente
End Property

Public Property Let lCliente(ByVal vData As Long)
    mvarlCliente = vData
End Property

Public Property Get lCliente() As Long
    lCliente = mvarlCliente
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

Public Property Let iOrigem(ByVal vData As Integer)
    mvariOrigem = vData
End Property

Public Property Get iOrigem() As Integer
    iOrigem = mvariOrigem
End Property

Public Property Let iDestinatario(ByVal vData As Integer)
    mvariDestinatario = vData
End Property

Public Property Get iDestinatario() As Integer
    iDestinatario = mvariDestinatario
End Property

Public Property Let iEmitente(ByVal vData As Integer)
    mvariEmitente = vData
End Property

Public Property Get iEmitente() As Integer
    iEmitente = mvariEmitente
End Property

Public Property Let sIndicadorMovCMP(ByVal vData As String)
    mvarsIndicadorMovCMP = vData
End Property

Public Property Get sIndicadorMovCMP() As String
    sIndicadorMovCMP = mvarsIndicadorMovCMP
End Property

Public Property Let sObservacao(ByVal vData As String)
    mvarsObservacao = vData
End Property

Public Property Get sObservacao() As String
    sObservacao = mvarsObservacao
End Property

Public Property Let sTipoFatura(ByVal vData As String)
    mvarsTipoFatura = vData
End Property

Public Property Get sTipoFatura() As String
    sTipoFatura = mvarsTipoFatura
End Property

Public Property Let sIDVeiculo(ByVal vData As String)
    mvarsIDVeiculo = vData
End Property

Public Property Get sIDVeiculo() As String
    sIDVeiculo = mvarsIDVeiculo
End Property

Public Property Let iModalidadeFrete(ByVal vData As Integer)
    mvariModalidadeFrete = vData
End Property

Public Property Get iModalidadeFrete() As Integer
    iModalidadeFrete = mvariModalidadeFrete
End Property

Public Property Let dPesoLiquido(ByVal vData As Double)
    mvardPesoLiquido = vData
End Property

Public Property Get dPesoLiquido() As Double
    dPesoLiquido = mvardPesoLiquido
End Property

Public Property Let dPesoBruto(ByVal vData As Double)
    mvardPesoBruto = vData
End Property

Public Property Get dPesoBruto() As Double
    dPesoBruto = mvardPesoBruto
End Property

Public Property Let sEspecieVolume(ByVal vData As String)
    mvarsEspecieVolume = vData
End Property

Public Property Get sEspecieVolume() As String
    sEspecieVolume = mvarsEspecieVolume
End Property

Public Property Let iQtdVolumes(ByVal vData As Integer)
    mvariQtdVolumes = vData
End Property

Public Property Get iQtdVolumes() As Integer
    iQtdVolumes = mvariQtdVolumes
End Property

Public Property Let iTransportador(ByVal vData As Integer)
    mvariTransportador = vData
End Property

Public Property Get iTransportador() As Integer
    iTransportador = mvariTransportador
End Property

Public Property Let iViaTransporte(ByVal vData As Integer)
    mvariViaTransporte = vData
End Property

Public Property Get iViaTransporte() As Integer
    iViaTransporte = mvariViaTransporte
End Property

Public Property Let sInscEstSubsTRB(ByVal vData As String)
    mvarsInscEstSubsTRB = vData
End Property

Public Property Get sInscEstSubsTRB() As String
    sInscEstSubsTRB = mvarsInscEstSubsTRB
End Property

Public Property Let dValorTotalNF(ByVal vData As Double)
    mvardValorTotalNF = vData
End Property

Public Property Get dValorTotalNF() As Double
    dValorTotalNF = mvardValorTotalNF
End Property

Public Property Let dValorTotalICMSSubsTRB(ByVal vData As Double)
    mvardValorTotalICMSSubsTRB = vData
End Property

Public Property Get dValorTotalICMSSubsTRB() As Double
    dValorTotalICMSSubsTRB = mvardValorTotalICMSSubsTRB
End Property

Public Property Let dValorTotalIPI(ByVal vData As Double)
    mvardValorTotalIPI = vData
End Property

Public Property Get dValorTotalIPI() As Double
    dValorTotalIPI = mvardValorTotalIPI
End Property

Public Property Let dValorOutrasDesp(ByVal vData As Double)
    mvardValorOutrasDesp = vData
End Property

Public Property Get dValorOutrasDesp() As Double
    dValorOutrasDesp = mvardValorOutrasDesp
End Property

Public Property Let dValorSeguro(ByVal vData As Double)
    mvardValorSeguro = vData
End Property

Public Property Get dValorSeguro() As Double
    dValorSeguro = mvardValorSeguro
End Property

Public Property Let dValorFrete(ByVal vData As Double)
    mvardValorFrete = vData
End Property

Public Property Get dValorFrete() As Double
    dValorFrete = mvardValorFrete
End Property

Public Property Let dValorTotDesconto(ByVal vData As Double)
    mvardValorTotDesconto = vData
End Property

Public Property Get dValorTotDesconto() As Double
    dValorTotDesconto = mvardValorTotDesconto
End Property

Public Property Let dValorMercadoria(ByVal vData As Double)
    mvardValorMercadoria = vData
End Property

Public Property Get dValorMercadoria() As Double
    dValorMercadoria = mvardValorMercadoria
End Property

Public Property Let dtDataEmissao(ByVal vData As Date)
    mvardtDataEmissao = vData
End Property

Public Property Get dtDataEmissao() As Date
    dtDataEmissao = mvardtDataEmissao
End Property

Public Property Let lNumeroDoc(ByVal vData As Long)
    mvarlNumeroDoc = vData
End Property

Public Property Get lNumeroDoc() As Long
    lNumeroDoc = mvarlNumeroDoc
End Property

Public Property Let sSerie(ByVal vData As String)
    mvarsSerie = vData
End Property

Public Property Get sSerie() As String
    sSerie = mvarsSerie
End Property

Public Property Let iModeloDoc(ByVal vData As Integer)
    mvariModeloDoc = vData
End Property

Public Property Get iModeloDoc() As Integer
    iModeloDoc = mvariModeloDoc
End Property

Public Property Let sIndicadorMov(ByVal vData As String)
    mvarsIndicadorMov = vData
End Property

Public Property Get sIndicadorMov() As String
    sIndicadorMov = mvarsIndicadorMov
End Property

Public Property Let iRemetente(ByVal vData As Integer)
    mvariRemetente = vData
End Property

Public Property Get iRemetente() As Integer
    iRemetente = mvariRemetente
End Property

'Incluído por Luiz Nogueira em 28/01/04
Public Property Let iTipoTipoDocInfo(ByVal vData As Integer)
    mvariTipoTipoDocInfo = vData
End Property

'Incluído por Luiz Nogueira em 28/01/04
Public Property Get iTipoTipoDocInfo() As Integer
    iTipoTipoDocInfo = mvariTipoTipoDocInfo
End Property


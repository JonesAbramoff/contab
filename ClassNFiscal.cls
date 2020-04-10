VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassNFiscal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"colComissoesNF"
Attribute VB_Ext_KEY = "Member1" ,"colItensNF"
Attribute VB_Ext_KEY = "Member2" ,"colParcelaPagar"
Attribute VB_Ext_KEY = "Member3" ,"colParcelaReceber"
Attribute VB_Ext_KEY = "Member4" ,"colParcelaReceber"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

Private Const ITEM_TIPO_FRETE = 1
Private Const ITEM_TIPO_SEGURO = 2
Private Const ITEM_TIPO_DESCONTO = 3
Private Const ITEM_TIPO_OUTRAS_DESP = 4

Public objInfoUsu As Object

'Daniel 23/05/2002
Private mvarlNumOrcamentoVenda As Long
Private mvariFilialOrcamento As Integer

Private mvarlNumIntDoc As Long
Private mvariFilialEmpresa As Integer
Private mvarsSerie As String
Private mvarlNumNotaFiscal As Long
Private mvarlCliente As Long
Private mvariFilialCli As Integer
Private mvariFilialEntrega As Integer
Private mvarlFornecedor As Long
Private mvariFilialForn As Integer
Private mvardtDataEmissao As Date
Private mvardtDataSaida As Date
Private mvarlNumPedidoVenda As Long
Private mvarsNumPedidoTerc As String
Private mvardValorProdutos As Double
Private mvardValorFrete As Double
Private mvardValorSeguro As Double
Private mvardValorOutrasDespesas As Double
Private mvardValorDesconto As Double
Private mvariCodTransportadora As Integer
Private mvarsMensagemNota As String
Private mvariTabelaPreco As Integer
Private mvariTipoNFiscal As Integer
Private mvarsNaturezaOp As String
Private mvardPesoLiq As Double
Private mvardPesoBruto As Double
Private mvardtDataVencimento As Date
Private mvarlNumIntTrib As Long
Private mvarsPlaca As String
Private mvarsPlacaUF As String
Private mvarlVolumeQuant As Long
Private mvarlVolumeEspecie As Long
Private mvarlVolumeMarca As Long
Private mvariCanal As Integer
Private mvarlNumIntNotaOriginal As Long
Private mvarcolItensNF As New colItensNF
Private mvarcolComissoesNF As New Collection
Private mvarcolParcelaPagar As New colParcelaPagar
Private mvarcolParcelaReceber As New colParcelaReceber
Private mvardtDataEntrada As Date
Private mvardValorTotal As Double
Private mvariClasseDocCPR As Integer
Private mvarlNumIntDocCPR As Long
Private mvariStatus As Integer
Private mvarobjTributacaoNF As New ClassTributacaoDoc
Private mvarsCodUsuario As String
Private mvariFilialPedido As Integer
Private mvarlClienteBenef As Long
Private mvariFilialCliBenef As Integer
Private mvarlFornecedorBenef As Long
Private mvariFilialFornBenef As Integer
Private mvarobjMovEstoque As ClassMovEstoque
Private mvariCondicaoPagto As Integer
Private mvarsVolumeNumero As String
Private mvariFreteRespons As Integer
Private mvardtDataRegistro As Date
Private mvardtDataReferencia As Date
Private mvarlNumRecebimento As Long
Private mvariTipoDocInfo As Long
Private mvarsObservacao As String
Private mvarsCodUsuarioCancel As String
Private mvarsMotivoCancel As String
Private mvarobjConhecimentoFrete As New ClassConhecimentoFrete
Private mvarobjRastreamento As Object
Private mvarsNomeTelaNFiscal As String
Private mvardtHoraEntrada As Date
Private mvardtHoraSaida As Date
Private mvarsDestino As String
Private mvarsOrigem As String
Private mvardValorContainer As Double
Private mvardValorMercadoria As Double
Private mvarcolComprovServ As New Collection
Private mvariCodTranspRedesp As Integer
Private mvariDetPagFrete As Integer
Private mvariSemDataSaida As Integer
Private mvardVolumeTotal As Double
Private mvariMoeda As Integer
Private mvardTaxaMoeda As Double
Private mvariNaoVerificaCredito As Integer
Private mvarsMensagemCorpoNota As String
'Private mvarsNatureza As String
Private mvarobjProjetoInfo As Object
Private mvariNaoImpCobranca As Integer
Private mvarlNumNFe As Long
Private mvarsCodVerificacaoNFe As String
Private mvariRecibo As Integer

Private mvarcolPacksNF As New Collection

'informaçoes para notas de importacao
Public objNFImportacao As New ClassNFImportacao

Private mvarobjTela As Object
Private mvarsStatusNFeFed As String
Private mvariNFe As Integer
Private mvarcolItensSRV As New Collection
Private mvardValorDescontoTit As Double
Private mvarobjInfoAdic As New ClassInfoAdic
Private mvardValorDescontoItens As Double
Private mvardValorItens As Double
Private mvarobjContabil As New ClassContabil
Private mvarobjNFDevBenefNFRem As Object 'Guarda informações sobre o cálculo da NF de devolução beneficamento com base em dados da remessa e produções
Private mvarlFornEntTerc As Long
Private mvariFilialFornEntTerc As Integer
Private mvarsChvNFe As String

Public dtDataCancel As Date
Public dtDataRegCancel As Date
Public dtHoraRegCancel As Date

'nfe 3.10
Public dtHoraEmissao As Date

Public dtDataCadastro As Date
Public dtHoraCadastro As Date
Public sUsuCadastro As String

Private mvarsCGCEmitente As String
Private mvariRollBack As Integer 'indica que a rotina de cancelamento deve ser desfeita pois da primeira vez so deve testar se o cancelamento é possivel

Public iNumNFManual As Integer '0:normal, 1: o numero já é passado

Public lNumIntNFAnterior As Long 'usado na alteracao de nf
Private mvarcolReqCompraBaixadas As New Collection 'Coleção auxiliar para tratamento de eventos no workflow
Private mvarName As String

Private mvarlCliIntermediario As Long
Private mvariFilialCliIntermediario As Integer
Private mvarsSerieNFPOrig As String
Private mvarlNumNFPOrig As Long

Public sQRCode As String

Public Property Let lNumNFPOrig(ByVal vData As Long)
    mvarlNumNFPOrig = vData
End Property

Public Property Get lNumNFPOrig() As Long
    lNumNFPOrig = mvarlNumNFPOrig
End Property

Public Property Let sSerieNFPOrig(ByVal vData As String)
    mvarsSerieNFPOrig = vData
End Property

Public Property Get sSerieNFPOrig() As String
    sSerieNFPOrig = mvarsSerieNFPOrig
End Property

Public Property Let iFilialCliIntermediario(ByVal vData As Integer)
    mvariFilialCliIntermediario = vData
End Property

Public Property Get iFilialCliIntermediario() As Integer
    iFilialCliIntermediario = mvariFilialCliIntermediario
End Property

Public Property Let lCliIntermediario(ByVal vData As Long)
    mvarlCliIntermediario = vData
End Property

Public Property Get lCliIntermediario() As Long
    lCliIntermediario = mvarlCliIntermediario
End Property


Public Property Let Name(ByVal vData As String)
    mvarName = vData
End Property

Public Property Get Name() As String
    Name = mvarName
End Property

Public Property Set colReqCompraBaixadas(ByVal vData As Collection)
    Set mvarcolReqCompraBaixadas = vData
End Property

Public Property Get colReqCompraBaixadas() As Collection
    Set colReqCompraBaixadas = mvarcolReqCompraBaixadas
End Property

Public Property Let iRollBack(ByVal vData As Integer)
    mvariRollBack = vData
End Property

Public Property Get iRollBack() As Integer
    iRollBack = mvariRollBack
End Property

Public Property Let sCGCEmitente(ByVal vData As String)
    mvarsCGCEmitente = vData
End Property

Public Property Get sCGCEmitente() As String
    sCGCEmitente = mvarsCGCEmitente
End Property

Public Property Let sChvNFe(ByVal vData As String)
    mvarsChvNFe = vData
End Property

Public Property Get sChvNFe() As String
    sChvNFe = mvarsChvNFe
End Property

Public Property Let iFilialFornEntTerc(ByVal vData As Integer)
    mvariFilialFornEntTerc = vData
End Property

Public Property Get iFilialFornEntTerc() As Integer
    iFilialFornEntTerc = mvariFilialFornEntTerc
End Property

Public Property Let lFornEntTerc(ByVal vData As Long)
    mvarlFornEntTerc = vData
End Property

Public Property Get lFornEntTerc() As Long
    lFornEntTerc = mvarlFornEntTerc
End Property

Public Property Set objNFDevBenefNFRem(ByVal vData As Object)
    Set mvarobjNFDevBenefNFRem = vData
End Property

Public Property Get objNFDevBenefNFRem() As Object
    Set objNFDevBenefNFRem = mvarobjNFDevBenefNFRem
End Property

Public Property Set objContabil(ByVal vData As ClassContabil)
    Set mvarobjContabil = vData
End Property

Public Property Get objContabil() As ClassContabil
    Set objContabil = mvarobjContabil
End Property

Public Property Let dValorItens(ByVal vData As Double)
    mvardValorItens = vData
End Property

Public Property Get dValorItens() As Double
    dValorItens = mvardValorItens
End Property

Public Property Let dValorDescontoItens(ByVal vData As Double)
    mvardValorDescontoItens = vData
End Property

Public Property Get dValorDescontoItens() As Double
    dValorDescontoItens = mvardValorDescontoItens
End Property

Public Property Set objInfoAdic(ByVal vData As ClassInfoAdic)
    Set mvarobjInfoAdic = vData
End Property

Public Property Get objInfoAdic() As ClassInfoAdic
    Set objInfoAdic = mvarobjInfoAdic
End Property

Public Property Let dValorDescontoTit(ByVal vData As Double)
    mvardValorDescontoTit = vData
End Property

Public Property Get dValorDescontoTit() As Double
    dValorDescontoTit = mvardValorDescontoTit
End Property

Public Property Set colItensSRV(ByVal vData As Collection)
    Set mvarcolItensSRV = vData
End Property

Public Property Get colItensSRV() As Collection
    Set colItensSRV = mvarcolItensSRV
End Property

Public Property Let iNFe(ByVal vData As Integer)
    mvariNFe = vData
End Property

Public Property Get iNFe() As Integer
    iNFe = mvariNFe
End Property

Public Property Let sStatusNFeFed(ByVal vData As String)
    mvarsStatusNFeFed = vData
End Property

Public Property Get sStatusNFeFed() As String
    sStatusNFeFed = mvarsStatusNFeFed
End Property

Public Property Set objTela(ByVal vData As Object)
    Set mvarobjTela = vData
End Property

Public Property Get objTela() As Object
    Set objTela = mvarobjTela
End Property

Public Property Let iRecibo(ByVal vData As Integer)
    mvariRecibo = vData
End Property

Public Property Get iRecibo() As Integer
    iRecibo = mvariRecibo
End Property

Public Property Let sCodVerificacaoNFe(ByVal vData As String)
    mvarsCodVerificacaoNFe = vData
End Property

Public Property Get sCodVerificacaoNFe() As String
    sCodVerificacaoNFe = mvarsCodVerificacaoNFe
End Property

Public Property Let lNumNFe(ByVal vData As Long)
    mvarlNumNFe = vData
End Property

Public Property Get lNumNFe() As Long
    lNumNFe = mvarlNumNFe
End Property

Public Property Let iNaoImpCobranca(ByVal vData As Integer)
    mvariNaoImpCobranca = vData
End Property

Public Property Get iNaoImpCobranca() As Integer
    iNaoImpCobranca = mvariNaoImpCobranca
End Property

Public Property Let sNatureza(ByVal vData As String)
    mvarobjInfoAdic.sNatureza = vData
End Property

Public Property Get sNatureza() As String
    sNatureza = mvarobjInfoAdic.sNatureza
End Property

Public Property Set objProjetoInfo(ByVal vData As Object)
    Set mvarobjProjetoInfo = vData
End Property

Public Property Get objProjetoInfo() As Object
    Set objProjetoInfo = mvarobjProjetoInfo
End Property

Public Property Let sMensagemCorpoNota(ByVal vData As String)
    mvarsMensagemCorpoNota = vData
End Property

Public Property Get sMensagemCorpoNota() As String
    sMensagemCorpoNota = mvarsMensagemCorpoNota
End Property

Public Property Let iNaoVerificaCredito(ByVal vData As Integer)
    mvariNaoVerificaCredito = vData
End Property

Public Property Get iNaoVerificaCredito() As Integer
    iNaoVerificaCredito = mvariNaoVerificaCredito
End Property

Public Property Let dVolumeTotal(ByVal vData As Double)
    mvardVolumeTotal = vData
End Property

Public Property Get dVolumeTotal() As Double
    dVolumeTotal = mvardVolumeTotal
End Property

Public Property Let iMoeda(ByVal vData As Integer)
    mvariMoeda = vData
End Property

Public Property Get iMoeda() As Integer
    iMoeda = mvariMoeda
End Property

Public Property Let dTaxaMoeda(ByVal vData As Double)
    mvardTaxaMoeda = vData
End Property

Public Property Get dTaxaMoeda() As Double
    dTaxaMoeda = mvardTaxaMoeda
End Property

Public Property Let iSemDataSaida(ByVal vData As Integer)
    mvariSemDataSaida = vData
End Property

Public Property Get iSemDataSaida() As Integer
    iSemDataSaida = mvariSemDataSaida
End Property

Public Property Let iDetPagFrete(ByVal vData As Integer)
    mvariDetPagFrete = vData
End Property

Public Property Get iDetPagFrete() As Integer
    iDetPagFrete = mvariDetPagFrete
End Property

Public Property Let iCodTranspRedesp(ByVal vData As Integer)
    mvariCodTranspRedesp = vData
End Property

Public Property Get iCodTranspRedesp() As Integer
    iCodTranspRedesp = mvariCodTranspRedesp
End Property

Public Property Let dtHoraSaida(ByVal vData As Date)
    mvardtHoraSaida = vData
End Property

Public Property Get dtHoraSaida() As Date
    dtHoraSaida = mvardtHoraSaida
End Property

Public Property Let dtHoraEntrada(ByVal vData As Date)
    mvardtHoraEntrada = vData
End Property

Public Property Get dtHoraEntrada() As Date
    dtHoraEntrada = mvardtHoraEntrada
End Property

Public Property Let sNomeTelaNFiscal(ByVal vData As String)
    mvarsNomeTelaNFiscal = vData
End Property

Public Property Get sNomeTelaNFiscal() As String
    sNomeTelaNFiscal = mvarsNomeTelaNFiscal
End Property

Public Property Set objRastreamento(ByVal vData As Object)
    Set mvarobjRastreamento = vData
End Property

Public Property Get objRastreamento() As Object
    Set objRastreamento = mvarobjRastreamento
End Property

Public Property Set objConhecimentoFrete(ByVal vData As ClassConhecimentoFrete)
    Set mvarobjConhecimentoFrete = vData
End Property

Public Property Get objConhecimentoFrete() As ClassConhecimentoFrete
    Set objConhecimentoFrete = mvarobjConhecimentoFrete
End Property

Public Property Let sMotivoCancel(ByVal vData As String)
    mvarsMotivoCancel = vData
End Property

Public Property Get sMotivoCancel() As String
    sMotivoCancel = mvarsMotivoCancel
End Property

Public Property Let sCodUsuarioCancel(ByVal vData As String)
    mvarsCodUsuarioCancel = vData
End Property

Public Property Get sCodUsuarioCancel() As String
    sCodUsuarioCancel = mvarsCodUsuarioCancel
End Property

Public Property Let sObservacao(ByVal vData As String)
    mvarsObservacao = vData
End Property

Public Property Get sObservacao() As String
    sObservacao = mvarsObservacao
End Property

Public Property Let iTipoDocInfo(ByVal vData As Integer)
    mvariTipoDocInfo = vData
End Property

Public Property Get iTipoDocInfo() As Integer
    iTipoDocInfo = mvariTipoDocInfo
End Property

Public Property Let lNumRecebimento(ByVal vData As Long)
Attribute lNumRecebimento.VB_Description = "Número do Recebimento"
    mvarlNumRecebimento = vData
End Property

Public Property Get lNumRecebimento() As Long
    lNumRecebimento = mvarlNumRecebimento
End Property

Public Property Let dtDataReferencia(ByVal vData As Date)
    mvardtDataReferencia = vData
End Property

Public Property Get dtDataReferencia() As Date
    dtDataReferencia = mvardtDataReferencia
End Property

Public Property Let dtDataRegistro(ByVal vData As Date)
    mvardtDataRegistro = vData
End Property

Public Property Get dtDataRegistro() As Date
    dtDataRegistro = mvardtDataRegistro
End Property

Public Property Let iFreteRespons(ByVal vData As Integer)
    mvariFreteRespons = vData
End Property

Public Property Get iFreteRespons() As Integer
    iFreteRespons = mvariFreteRespons
End Property

Public Property Let sVolumeNumero(ByVal vData As String)
    mvarsVolumeNumero = vData
End Property


Public Property Get sVolumeNumero() As String
    sVolumeNumero = mvarsVolumeNumero
End Property

Public Property Let iCondicaoPagto(ByVal vData As Integer)
    mvariCondicaoPagto = vData
End Property

Public Property Get iCondicaoPagto() As Integer
    iCondicaoPagto = mvariCondicaoPagto
End Property

Public Property Set objMovEstoque(ByVal vData As ClassMovEstoque)
    Set mvarobjMovEstoque = vData
End Property

Public Property Get objMovEstoque() As ClassMovEstoque
        Set objMovEstoque = mvarobjMovEstoque
End Property

Public Property Let iFilialFornBenef(ByVal vData As Integer)
Attribute iFilialFornBenef.VB_Description = "Filial do Fornecedor que vai beneficiar os nossos produtos."
    mvariFilialFornBenef = vData
End Property

Public Property Get iFilialFornBenef() As Integer
    iFilialFornBenef = mvariFilialFornBenef
End Property

Public Property Let lFornecedorBenef(ByVal vData As Long)
Attribute lFornecedorBenef.VB_Description = "Código do Fornecedor que vai beneficiar os nossos produtos."
    mvarlFornecedorBenef = vData
End Property

Public Property Get lFornecedorBenef() As Long
    lFornecedorBenef = mvarlFornecedorBenef
End Property

Public Property Let iFilialCliBenef(ByVal vData As Integer)
Attribute iFilialCliBenef.VB_Description = "Código da Filial do Cliente cujos produtos serão beneficiados por nós."
    mvariFilialCliBenef = vData
End Property

Public Property Get iFilialCliBenef() As Integer
    iFilialCliBenef = mvariFilialCliBenef
End Property

Public Property Let lClienteBenef(ByVal vData As Long)
Attribute lClienteBenef.VB_Description = "Código do Cliente cujos produtos serão beneficiados por nós."
    mvarlClienteBenef = vData
End Property

Public Property Get lClienteBenef() As Long
    lClienteBenef = mvarlClienteBenef
End Property

Public Property Let iFilialPedido(ByVal vData As Integer)
    mvariFilialPedido = vData
End Property

Public Property Get iFilialPedido() As Integer
    iFilialPedido = mvariFilialPedido
End Property

Public Property Let sCodUsuario(ByVal vData As String)
Attribute sCodUsuario.VB_Description = "Codigo do Usuario que autorizou o crédito (se preenchido)"
    mvarsCodUsuario = vData
End Property

Public Property Get sCodUsuario() As String
    sCodUsuario = mvarsCodUsuario
End Property

Public Property Set objTributacaoNF(ByVal vData As Object)
    Set mvarobjTributacaoNF = vData
End Property

Public Property Get objTributacaoNF() As ClassTributacaoDoc
    Set objTributacaoNF = mvarobjTributacaoNF
End Property

Public Property Let iStatus(ByVal vData As Integer)
    mvariStatus = vData
End Property

Public Property Get iStatus() As Integer
    iStatus = mvariStatus
End Property

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

Public Property Let dValorTotal(ByVal vData As Double)
    mvardValorTotal = vData
End Property

Public Property Get dValorTotal() As Double
    dValorTotal = mvardValorTotal
End Property

Public Property Let dtDataEntrada(ByVal vData As Date)
    mvardtDataEntrada = vData
End Property

Public Property Get dtDataEntrada() As Date
    dtDataEntrada = mvardtDataEntrada
End Property

Public Property Set colParcelaReceber(ByVal vData As Object)
    Set mvarcolParcelaReceber = vData
End Property

Public Property Get colParcelaReceber() As colParcelaReceber
    Set colParcelaReceber = mvarcolParcelaReceber
End Property

Public Property Set colParcelaPagar(ByVal vData As Object)
    Set mvarcolParcelaPagar = vData
End Property

Public Property Get colParcelaPagar() As colParcelaPagar
    Set colParcelaPagar = mvarcolParcelaPagar
End Property

Public Property Set colComissoesNF(ByVal vData As Object)
    Set mvarcolComissoesNF = vData
End Property

Public Property Get colComissoesNF() As Collection
    Set colComissoesNF = mvarcolComissoesNF
End Property

Public Property Set colItensNF(ByVal vData As Object)
    Set mvarcolItensNF = vData
End Property

Public Property Get colItensNF() As colItensNF
    Set colItensNF = mvarcolItensNF
End Property

Public Property Set colPacksNF(ByVal vData As Object)
    Set mvarcolPacksNF = vData
End Property

Public Property Get colPacksNF() As Collection
    Set colPacksNF = mvarcolPacksNF
End Property

Public Property Let lNumIntNotaOriginal(ByVal vData As Long)
    mvarlNumIntNotaOriginal = vData
End Property

Public Property Get lNumIntNotaOriginal() As Long
    lNumIntNotaOriginal = mvarlNumIntNotaOriginal
End Property

Public Property Let iCanal(ByVal vData As Integer)
    mvariCanal = vData
End Property

Public Property Get iCanal() As Integer
    iCanal = mvariCanal
End Property

Public Property Let lVolumeMarca(ByVal vData As Long)
    mvarlVolumeMarca = vData
End Property

Public Property Get lVolumeMarca() As Long
    lVolumeMarca = mvarlVolumeMarca
End Property

Public Property Let lVolumeEspecie(ByVal vData As Long)
    mvarlVolumeEspecie = vData
End Property

Public Property Get lVolumeEspecie() As Long
    lVolumeEspecie = mvarlVolumeEspecie
End Property

Public Property Let lVolumeQuant(ByVal vData As Long)
    mvarlVolumeQuant = vData
End Property

Public Property Get lVolumeQuant() As Long
    lVolumeQuant = mvarlVolumeQuant
End Property

Public Property Let sPlacaUF(ByVal vData As String)
    mvarsPlacaUF = vData
End Property

Public Property Get sPlacaUF() As String
    sPlacaUF = mvarsPlacaUF
End Property

Public Property Let sPlaca(ByVal vData As String)
    mvarsPlaca = vData
End Property

Public Property Get sPlaca() As String
    sPlaca = mvarsPlaca
End Property

Public Property Let lNumIntTrib(ByVal vData As Long)
    mvarlNumIntTrib = vData
End Property

Public Property Get lNumIntTrib() As Long
    lNumIntTrib = mvarlNumIntTrib
End Property

Public Property Let dtDataVencimento(ByVal vData As Date)
    mvardtDataVencimento = vData
End Property

Public Property Get dtDataVencimento() As Date
    dtDataVencimento = mvardtDataVencimento
End Property

Public Property Let dPesoBruto(ByVal vData As Double)
    mvardPesoBruto = vData
End Property

Public Property Get dPesoBruto() As Double
    dPesoBruto = mvardPesoBruto
End Property

Public Property Let dPesoLiq(ByVal vData As Double)
    mvardPesoLiq = vData
End Property

Public Property Get dPesoLiq() As Double
    dPesoLiq = mvardPesoLiq
End Property

Public Property Let sNaturezaOp(ByVal vData As String)
    mvarsNaturezaOp = vData
End Property

Public Property Get sNaturezaOp() As String
    sNaturezaOp = mvarsNaturezaOp
End Property

Public Property Let iTipoNFiscal(ByVal vData As Integer)
    mvariTipoNFiscal = vData
End Property

Public Property Get iTipoNFiscal() As Integer
    iTipoNFiscal = mvariTipoNFiscal
End Property

Public Property Let iTabelaPreco(ByVal vData As Integer)
    mvariTabelaPreco = vData
End Property

Public Property Get iTabelaPreco() As Integer
    iTabelaPreco = mvariTabelaPreco
End Property

Public Property Let sMensagemNota(ByVal vData As String)
    mvarsMensagemNota = vData
End Property

Public Property Get sMensagemNota() As String
    sMensagemNota = mvarsMensagemNota
End Property

Public Property Let iCodTransportadora(ByVal vData As Integer)
    mvariCodTransportadora = vData
End Property

Public Property Get iCodTransportadora() As Integer
    iCodTransportadora = mvariCodTransportadora
End Property

Public Property Let dValorDesconto(ByVal vData As Double)
    mvardValorDesconto = vData
End Property

Public Property Get dValorDesconto() As Double
    dValorDesconto = mvardValorDesconto
End Property

Public Property Let dValorOutrasDespesas(ByVal vData As Double)
    mvardValorOutrasDespesas = vData
End Property

Public Property Get dValorOutrasDespesas() As Double
    dValorOutrasDespesas = mvardValorOutrasDespesas
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

Public Property Let dValorProdutos(ByVal vData As Double)
    mvardValorProdutos = vData
End Property

Public Property Get dValorProdutos() As Double
    dValorProdutos = mvardValorProdutos
End Property

'Public Property Let lNumIntTitPag(ByVal vData As Long)
'    mvarlNumIntTitPag = vData
'End Property
'
'Public Property Get lNumIntTitPag() As Long
'    lNumIntTitPag = mvarlNumIntTitPag
'End Property
'
'Public Property Let lNumIntTitRec(ByVal vData As Long)
'    mvarlNumIntTitRec = vData
'End Property
'
'Public Property Get lNumIntTitRec() As Long
'    lNumIntTitRec = mvarlNumIntTitRec
'End Property
'
Public Property Let lNumPedidoVenda(ByVal vData As Long)
    mvarlNumPedidoVenda = vData
End Property

Public Property Get lNumPedidoVenda() As Long
    lNumPedidoVenda = mvarlNumPedidoVenda
End Property

Public Property Let dtDataSaida(ByVal vData As Date)
    mvardtDataSaida = vData
End Property

Public Property Get dtDataSaida() As Date
    dtDataSaida = mvardtDataSaida
End Property

Public Property Let dtDataEmissao(ByVal vData As Date)
    mvardtDataEmissao = vData
End Property

Public Property Get dtDataEmissao() As Date
    dtDataEmissao = mvardtDataEmissao
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

Public Property Let iFilialEntrega(ByVal vData As Integer)
    mvariFilialEntrega = vData
End Property

Public Property Get iFilialEntrega() As Integer
    iFilialEntrega = mvariFilialEntrega
End Property

Public Property Let iFilialCli(ByVal vData As Integer)
    mvariFilialCli = vData
End Property

Public Property Get iFilialCli() As Integer
    iFilialCli = mvariFilialCli
End Property

Public Property Let lCliente(ByVal vData As Long)
    mvarlCliente = vData
End Property

Public Property Get lCliente() As Long
    lCliente = mvarlCliente
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

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let sNumPedidoTerc(ByVal vData As String)
    mvarsNumPedidoTerc = vData
End Property

Public Property Get sNumPedidoTerc() As String
    sNumPedidoTerc = mvarsNumPedidoTerc
End Property

Public Property Get dValorLivroAux() As Double
    dValorLivroAux = Round(mvardValorTotal - objTributacaoNF.dIRRFValor, 2)
End Property

Public Property Let sDestino(ByVal vData As String)
    mvarsDestino = vData
End Property

Public Property Get sDestino() As String
    sDestino = mvarsDestino
End Property

Public Property Let sOrigem(ByVal vData As String)
    mvarsOrigem = vData
End Property

Public Property Get sOrigem() As String
    sOrigem = mvarsOrigem
End Property

Public Property Let dValorContainer(ByVal vData As Double)
    mvardValorContainer = vData
End Property

Public Property Get dValorContainer() As Double
    dValorContainer = mvardValorContainer
End Property

Public Property Let dValorMercadoria(ByVal vData As Double)
    mvardValorMercadoria = vData
End Property

Public Property Get dValorMercadoria() As Double
    dValorMercadoria = mvardValorMercadoria
End Property

Public Property Set colComprovServ(ByVal vData As Object)
    Set mvarcolComprovServ = vData
End Property

Public Property Get colComprovServ() As Collection
    Set colComprovServ = mvarcolComprovServ
End Property

'Daniel 23/05/2002
Public Property Let lNumOrcamentoVenda(ByVal vData As Long)
    mvarlNumOrcamentoVenda = vData
End Property

'Daniel 23/05/2002
Public Property Get lNumOrcamentoVenda() As Long
    lNumOrcamentoVenda = mvarlNumOrcamentoVenda
End Property

'Daniel 23/05/2002
Public Property Let iFilialOrcamento(ByVal vData As Long)
    mvariFilialOrcamento = vData
End Property

'Daniel 23/05/2002
Public Property Get iFilialOrcamento() As Long
    iFilialOrcamento = mvariFilialOrcamento
End Property

Sub Inicializa_Tributacao()
'
Dim objItemNF As Object

    Set objTributacaoNF = New ClassTributacaoDoc
    
    For Each objItemNF In colItensNF
        Call objItemNF.Inicializa_Tributacao
    Next

End Sub

Public Property Set colItens(ByVal vData As Object)
    Dim vValor As ClassItemNF
    Dim colItensAux As New colItensNF
    For Each vValor In vData
        colItensAux.Add1 vValor
    Next
    Set mvarcolItensNF = colItensAux
End Property

Public Property Get colItens() As Collection
    Dim vValor As Variant
    Dim colItensAux As New Collection
    For Each vValor In colItensNF
        colItensAux.Add vValor
    Next
    Set colItens = colItensAux
End Property

Public Property Set objTributacao(ByVal vData As Object)
    Set mvarobjTributacaoNF = vData
End Property

Public Property Get objTributacao() As ClassTributacaoDoc
    Set objTributacao = mvarobjTributacaoNF
End Property

Private Sub Class_Initialize()

    dtDataCancel = DATA_NULA
    dtDataRegCancel = DATA_NULA
    dtDataCadastro = DATA_NULA

End Sub

Function Calcula_Mnemonico(objMnemonicoValor As ClassMnemonicoValor, Optional objContexto As Object) As Long
'Calcula_Mnemonico do WorkFlow genérico para todas notas fiscais

Dim lErro As Long
Dim objCli As New ClassCliente, objFilCli As New ClassFilialCliente
Dim objVend As New ClassVendedor, objEnd As New ClassEndereco

On Error GoTo Erro_Calcula_Mnemonico

    Select Case objMnemonicoValor.sMnemonico
                        
        Case "NumNotaFiscal"
            objMnemonicoValor.colValor.Add lNumNotaFiscal
            
        Case "FilialEmpresa"
            objMnemonicoValor.colValor.Add iFilialEmpresa
            
        Case "DataEmissao"
            objMnemonicoValor.colValor.Add dtDataEmissao
            
        Case "DataEmissaoTxt"
            objMnemonicoValor.colValor.Add Format(dtDataEmissao, "dd/mm/yyyy")
            
        Case "Serie"
            objMnemonicoValor.colValor.Add sSerie
            
        Case "Cliente"
            objMnemonicoValor.colValor.Add lCliente
        
        Case "FilialCliente"
            objMnemonicoValor.colValor.Add iFilialCli
        
        Case "NomeCliente"
        
            objCli.lCodigo = lCliente
                
            lErro = CF("Cliente_Le", objCli)
            If lErro <> SUCESSO And lErro <> 12293 Then gError ERRO_SEM_MENSAGEM
        
            objMnemonicoValor.colValor.Add objCli.sRazaoSocial
                       
        Case "CGC"
        
            objFilCli.lCodCliente = lCliente
            objFilCli.iCodFilial = iFilialCli
            
            'Verifica se já existe a Filial do Cliente lido na tabela FiliaisClientes
            lErro = CF("FilialCliente_Le", objFilCli)
            If lErro <> SUCESSO And lErro <> 12567 Then gError ERRO_SEM_MENSAGEM
            
            objMnemonicoValor.colValor.Add objFilCli.sCgc
                       
        Case "Vendedor"
            
            objFilCli.lCodCliente = lCliente
            objFilCli.iCodFilial = iFilialCli
            
            'Verifica se já existe a Filial do Cliente lido na tabela FiliaisClientes
            lErro = CF("FilialCliente_Le", objFilCli)
            If lErro <> SUCESSO And lErro <> 12567 Then gError ERRO_SEM_MENSAGEM
            
            objMnemonicoValor.colValor.Add objFilCli.iVendedor
            
        Case "NomeVendedor"
            
            objFilCli.lCodCliente = lCliente
            objFilCli.iCodFilial = iFilialCli
            
            'Verifica se já existe a Filial do Cliente lido na tabela FiliaisClientes
            lErro = CF("FilialCliente_Le", objFilCli)
            If lErro <> SUCESSO And lErro <> 12567 Then gError ERRO_SEM_MENSAGEM
            
            objVend.iCodigo = objFilCli.iVendedor
            
            lErro = CF("Vendedor_Le", objVend)
            If lErro <> SUCESSO And lErro <> 12582 Then gError ERRO_SEM_MENSAGEM
            
            objMnemonicoValor.colValor.Add objVend.sNome
            
        Case "EmailVendedor"
            
            objFilCli.lCodCliente = lCliente
            objFilCli.iCodFilial = iFilialCli
            
            'Verifica se já existe a Filial do Cliente lido na tabela FiliaisClientes
            lErro = CF("FilialCliente_Le", objFilCli)
            If lErro <> SUCESSO And lErro <> 12567 Then gError ERRO_SEM_MENSAGEM
            
            objVend.iCodigo = objFilCli.iVendedor
            
            lErro = CF("Vendedor_Le", objVend)
            If lErro <> SUCESSO And lErro <> 12582 Then gError ERRO_SEM_MENSAGEM
            
            objEnd.lCodigo = objVend.lEndereco
        
            'Endereco de Retirada
            lErro = CF("Endereco_Le", objEnd)
            If lErro <> SUCESSO And lErro <> 12309 Then gError ERRO_SEM_MENSAGEM
            
            objMnemonicoValor.colValor.Add objEnd.sEmail
            
        Case "ValorTotal"
            objMnemonicoValor.colValor.Add Format(dValorTotal, "STANDARD")
                    
        Case Else
            gError 178683

    End Select

    Calcula_Mnemonico = SUCESSO

    Exit Function

Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case 178683
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178684)

    End Select

    Exit Function

End Function

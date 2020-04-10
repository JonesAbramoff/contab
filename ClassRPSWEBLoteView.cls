VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
End
Attribute VB_Name = "ClassRPSWEBLoteView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level", "; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvariStatus AS Integer
Private mvarsSerie AS String
Private mvarlNumNotaFiscal AS Long
Private mvarlCliente AS Long
Private mvariFilialCli AS Integer
Private mvariFilialEntrega AS Integer
Private mvarlFornecedor AS Long
Private mvariFilialForn AS Integer
Private mvardtDataEmissao AS Date
Private mvardtDataEntrada AS Date
Private mvardtDataSaida AS Date
Private mvardtDataVencimento AS Date
Private mvardtDataReferencia AS Date
Private mvarlNumPedidoVenda AS Long
Private mvarsNumPedidoTerc AS String
Private mvariClasseDocCPR AS Integer
Private mvariFilialPedido AS Integer
Private mvarlNumIntDocCPR AS Long
Private mvardValorTotal AS Double
Private mvardValorProdutos AS Double
Private mvardValorFrete AS Double
Private mvardValorSeguro AS Double
Private mvardValorOutrasDespesas AS Double
Private mvardValorDesconto AS Double
Private mvariCodTransportadora AS Integer
Private mvarsMensagemNota AS String
Private mvariTabelaPreco AS Integer
Private mvariTipoNFiscal AS Integer
Private mvarsNaturezaOp AS String
Private mvardPesoLiq AS Double
Private mvardPesoBruto AS Double
Private mvarlNumIntTrib AS Long
Private mvarsPlaca AS String
Private mvarsPlacaUF AS String
Private mvarlVolumeQuant AS Integer
Private mvarlVolumeEspecie AS Long
Private mvarlVolumeMarca AS Long
Private mvarsVolumeNumero AS String
Private mvariCanal AS Integer
Private mvarlNumIntNotaOriginal AS Long
Private mvarlClienteBenef AS Long
Private mvariFilialCliBenef AS Integer
Private mvarlFornecedorBenef AS Long
Private mvariFilialFornBenef AS Integer
Private mvariFreteRespons AS Integer
Private mvarlNumRecebimento AS Long
Private mvarsObservacao AS String
Private mvarsCodUsuarioCancel AS String
Private mvarsMotivoCancel AS String
Private mvardHoraEntrada AS Double
Private mvardHoraSaida AS Double
Private mvarsOrigemNF AS String
Private mvarsDestinoNF AS String
Private mvardValorContainer AS Double
Private mvardValorMercadoria AS Double
Private mvardtDataAtualizacao AS Date
Private mvariCodTranspRedesp AS Integer
Private mvariDetPagFrete AS Integer
Private mvariSemDataSaida AS Integer
Private mvardVolumeTotal AS Double
Private mvariMoeda AS Integer
Private mvardTaxaMoeda AS Double
Private mvarsMensagemCorpoNota AS String
Private mvarsNatureza AS String
Private mvariNaoImpCobranca AS Integer
Private mvariRPS AS Integer
Private mvarlNumNFe AS Long
Private mvarsCodVerificacaoNFe AS String
Private mvariMexiDataEmissao AS Integer
Private mvarsRazaoSocialCli AS String
Private mvarsNomeReduzidoCli AS String
Private mvarlNumIntNF AS Long
Private mvarlLote AS Long
Private mvariFilialEmpresa AS Integer
Private mvardtData AS Date
Private mvardHora AS Double
Private mvarsUsuario AS String
Private mvarsRazaoSocialForn AS String
Private mvarsNomeReduzidoForn AS String

Public Property Let iStatus (ByVal vData As Integer)
    mvariStatus = vData
End Property

Public Property Get iStatus () AS Integer
    iStatus= mvariStatus
End Property

Public Property Let sSerie (ByVal vData As String)
    mvarsSerie = vData
End Property

Public Property Get sSerie () AS String
    sSerie= mvarsSerie
End Property

Public Property Let lNumNotaFiscal (ByVal vData As Long)
    mvarlNumNotaFiscal = vData
End Property

Public Property Get lNumNotaFiscal () AS Long
    lNumNotaFiscal= mvarlNumNotaFiscal
End Property

Public Property Let lCliente (ByVal vData As Long)
    mvarlCliente = vData
End Property

Public Property Get lCliente () AS Long
    lCliente= mvarlCliente
End Property

Public Property Let iFilialCli (ByVal vData As Integer)
    mvariFilialCli = vData
End Property

Public Property Get iFilialCli () AS Integer
    iFilialCli= mvariFilialCli
End Property

Public Property Let iFilialEntrega (ByVal vData As Integer)
    mvariFilialEntrega = vData
End Property

Public Property Get iFilialEntrega () AS Integer
    iFilialEntrega= mvariFilialEntrega
End Property

Public Property Let lFornecedor (ByVal vData As Long)
    mvarlFornecedor = vData
End Property

Public Property Get lFornecedor () AS Long
    lFornecedor= mvarlFornecedor
End Property

Public Property Let iFilialForn (ByVal vData As Integer)
    mvariFilialForn = vData
End Property

Public Property Get iFilialForn () AS Integer
    iFilialForn= mvariFilialForn
End Property

Public Property Let dtDataEmissao (ByVal vData As Date)
    mvardtDataEmissao = vData
End Property

Public Property Get dtDataEmissao () AS Date
    dtDataEmissao= mvardtDataEmissao
End Property

Public Property Let dtDataEntrada (ByVal vData As Date)
    mvardtDataEntrada = vData
End Property

Public Property Get dtDataEntrada () AS Date
    dtDataEntrada= mvardtDataEntrada
End Property

Public Property Let dtDataSaida (ByVal vData As Date)
    mvardtDataSaida = vData
End Property

Public Property Get dtDataSaida () AS Date
    dtDataSaida= mvardtDataSaida
End Property

Public Property Let dtDataVencimento (ByVal vData As Date)
    mvardtDataVencimento = vData
End Property

Public Property Get dtDataVencimento () AS Date
    dtDataVencimento= mvardtDataVencimento
End Property

Public Property Let dtDataReferencia (ByVal vData As Date)
    mvardtDataReferencia = vData
End Property

Public Property Get dtDataReferencia () AS Date
    dtDataReferencia= mvardtDataReferencia
End Property

Public Property Let lNumPedidoVenda (ByVal vData As Long)
    mvarlNumPedidoVenda = vData
End Property

Public Property Get lNumPedidoVenda () AS Long
    lNumPedidoVenda= mvarlNumPedidoVenda
End Property

Public Property Let sNumPedidoTerc (ByVal vData As String)
    mvarsNumPedidoTerc = vData
End Property

Public Property Get sNumPedidoTerc () AS String
    sNumPedidoTerc= mvarsNumPedidoTerc
End Property

Public Property Let iClasseDocCPR (ByVal vData As Integer)
    mvariClasseDocCPR = vData
End Property

Public Property Get iClasseDocCPR () AS Integer
    iClasseDocCPR= mvariClasseDocCPR
End Property

Public Property Let iFilialPedido (ByVal vData As Integer)
    mvariFilialPedido = vData
End Property

Public Property Get iFilialPedido () AS Integer
    iFilialPedido= mvariFilialPedido
End Property

Public Property Let lNumIntDocCPR (ByVal vData As Long)
    mvarlNumIntDocCPR = vData
End Property

Public Property Get lNumIntDocCPR () AS Long
    lNumIntDocCPR= mvarlNumIntDocCPR
End Property

Public Property Let dValorTotal (ByVal vData As Double)
    mvardValorTotal = vData
End Property

Public Property Get dValorTotal () AS Double
    dValorTotal= mvardValorTotal
End Property

Public Property Let dValorProdutos (ByVal vData As Double)
    mvardValorProdutos = vData
End Property

Public Property Get dValorProdutos () AS Double
    dValorProdutos= mvardValorProdutos
End Property

Public Property Let dValorFrete (ByVal vData As Double)
    mvardValorFrete = vData
End Property

Public Property Get dValorFrete () AS Double
    dValorFrete= mvardValorFrete
End Property

Public Property Let dValorSeguro (ByVal vData As Double)
    mvardValorSeguro = vData
End Property

Public Property Get dValorSeguro () AS Double
    dValorSeguro= mvardValorSeguro
End Property

Public Property Let dValorOutrasDespesas (ByVal vData As Double)
    mvardValorOutrasDespesas = vData
End Property

Public Property Get dValorOutrasDespesas () AS Double
    dValorOutrasDespesas= mvardValorOutrasDespesas
End Property

Public Property Let dValorDesconto (ByVal vData As Double)
    mvardValorDesconto = vData
End Property

Public Property Get dValorDesconto () AS Double
    dValorDesconto= mvardValorDesconto
End Property

Public Property Let iCodTransportadora (ByVal vData As Integer)
    mvariCodTransportadora = vData
End Property

Public Property Get iCodTransportadora () AS Integer
    iCodTransportadora= mvariCodTransportadora
End Property

Public Property Let sMensagemNota (ByVal vData As String)
    mvarsMensagemNota = vData
End Property

Public Property Get sMensagemNota () AS String
    sMensagemNota= mvarsMensagemNota
End Property

Public Property Let iTabelaPreco (ByVal vData As Integer)
    mvariTabelaPreco = vData
End Property

Public Property Get iTabelaPreco () AS Integer
    iTabelaPreco= mvariTabelaPreco
End Property

Public Property Let iTipoNFiscal (ByVal vData As Integer)
    mvariTipoNFiscal = vData
End Property

Public Property Get iTipoNFiscal () AS Integer
    iTipoNFiscal= mvariTipoNFiscal
End Property

Public Property Let sNaturezaOp (ByVal vData As String)
    mvarsNaturezaOp = vData
End Property

Public Property Get sNaturezaOp () AS String
    sNaturezaOp= mvarsNaturezaOp
End Property

Public Property Let dPesoLiq (ByVal vData As Double)
    mvardPesoLiq = vData
End Property

Public Property Get dPesoLiq () AS Double
    dPesoLiq= mvardPesoLiq
End Property

Public Property Let dPesoBruto (ByVal vData As Double)
    mvardPesoBruto = vData
End Property

Public Property Get dPesoBruto () AS Double
    dPesoBruto= mvardPesoBruto
End Property

Public Property Let lNumIntTrib (ByVal vData As Long)
    mvarlNumIntTrib = vData
End Property

Public Property Get lNumIntTrib () AS Long
    lNumIntTrib= mvarlNumIntTrib
End Property

Public Property Let sPlaca (ByVal vData As String)
    mvarsPlaca = vData
End Property

Public Property Get sPlaca () AS String
    sPlaca= mvarsPlaca
End Property

Public Property Let sPlacaUF (ByVal vData As String)
    mvarsPlacaUF = vData
End Property

Public Property Get sPlacaUF () AS String
    sPlacaUF= mvarsPlacaUF
End Property

Public Property Let lVolumeQuant (ByVal vData As Integer)
    mvarlVolumeQuant = vData
End Property

Public Property Get lVolumeQuant () AS Integer
    lVolumeQuant= mvarlVolumeQuant
End Property

Public Property Let lVolumeEspecie (ByVal vData As Long)
    mvarlVolumeEspecie = vData
End Property

Public Property Get lVolumeEspecie () AS Long
    lVolumeEspecie= mvarlVolumeEspecie
End Property

Public Property Let lVolumeMarca (ByVal vData As Long)
    mvarlVolumeMarca = vData
End Property

Public Property Get lVolumeMarca () AS Long
    lVolumeMarca= mvarlVolumeMarca
End Property

Public Property Let sVolumeNumero (ByVal vData As String)
    mvarsVolumeNumero = vData
End Property

Public Property Get sVolumeNumero () AS String
    sVolumeNumero= mvarsVolumeNumero
End Property

Public Property Let iCanal (ByVal vData As Integer)
    mvariCanal = vData
End Property

Public Property Get iCanal () AS Integer
    iCanal= mvariCanal
End Property

Public Property Let lNumIntNotaOriginal (ByVal vData As Long)
    mvarlNumIntNotaOriginal = vData
End Property

Public Property Get lNumIntNotaOriginal () AS Long
    lNumIntNotaOriginal= mvarlNumIntNotaOriginal
End Property

Public Property Let lClienteBenef (ByVal vData As Long)
    mvarlClienteBenef = vData
End Property

Public Property Get lClienteBenef () AS Long
    lClienteBenef= mvarlClienteBenef
End Property

Public Property Let iFilialCliBenef (ByVal vData As Integer)
    mvariFilialCliBenef = vData
End Property

Public Property Get iFilialCliBenef () AS Integer
    iFilialCliBenef= mvariFilialCliBenef
End Property

Public Property Let lFornecedorBenef (ByVal vData As Long)
    mvarlFornecedorBenef = vData
End Property

Public Property Get lFornecedorBenef () AS Long
    lFornecedorBenef= mvarlFornecedorBenef
End Property

Public Property Let iFilialFornBenef (ByVal vData As Integer)
    mvariFilialFornBenef = vData
End Property

Public Property Get iFilialFornBenef () AS Integer
    iFilialFornBenef= mvariFilialFornBenef
End Property

Public Property Let iFreteRespons (ByVal vData As Integer)
    mvariFreteRespons = vData
End Property

Public Property Get iFreteRespons () AS Integer
    iFreteRespons= mvariFreteRespons
End Property

Public Property Let lNumRecebimento (ByVal vData As Long)
    mvarlNumRecebimento = vData
End Property

Public Property Get lNumRecebimento () AS Long
    lNumRecebimento= mvarlNumRecebimento
End Property

Public Property Let sObservacao (ByVal vData As String)
    mvarsObservacao = vData
End Property

Public Property Get sObservacao () AS String
    sObservacao= mvarsObservacao
End Property

Public Property Let sCodUsuarioCancel (ByVal vData As String)
    mvarsCodUsuarioCancel = vData
End Property

Public Property Get sCodUsuarioCancel () AS String
    sCodUsuarioCancel= mvarsCodUsuarioCancel
End Property

Public Property Let sMotivoCancel (ByVal vData As String)
    mvarsMotivoCancel = vData
End Property

Public Property Get sMotivoCancel () AS String
    sMotivoCancel= mvarsMotivoCancel
End Property

Public Property Let dHoraEntrada (ByVal vData As Double)
    mvardHoraEntrada = vData
End Property

Public Property Get dHoraEntrada () AS Double
    dHoraEntrada= mvardHoraEntrada
End Property

Public Property Let dHoraSaida (ByVal vData As Double)
    mvardHoraSaida = vData
End Property

Public Property Get dHoraSaida () AS Double
    dHoraSaida= mvardHoraSaida
End Property

Public Property Let sOrigemNF (ByVal vData As String)
    mvarsOrigemNF = vData
End Property

Public Property Get sOrigemNF () AS String
    sOrigemNF= mvarsOrigemNF
End Property

Public Property Let sDestinoNF (ByVal vData As String)
    mvarsDestinoNF = vData
End Property

Public Property Get sDestinoNF () AS String
    sDestinoNF= mvarsDestinoNF
End Property

Public Property Let dValorContainer (ByVal vData As Double)
    mvardValorContainer = vData
End Property

Public Property Get dValorContainer () AS Double
    dValorContainer= mvardValorContainer
End Property

Public Property Let dValorMercadoria (ByVal vData As Double)
    mvardValorMercadoria = vData
End Property

Public Property Get dValorMercadoria () AS Double
    dValorMercadoria= mvardValorMercadoria
End Property

Public Property Let dtDataAtualizacao (ByVal vData As Date)
    mvardtDataAtualizacao = vData
End Property

Public Property Get dtDataAtualizacao () AS Date
    dtDataAtualizacao= mvardtDataAtualizacao
End Property

Public Property Let iCodTranspRedesp (ByVal vData As Integer)
    mvariCodTranspRedesp = vData
End Property

Public Property Get iCodTranspRedesp () AS Integer
    iCodTranspRedesp= mvariCodTranspRedesp
End Property

Public Property Let iDetPagFrete (ByVal vData As Integer)
    mvariDetPagFrete = vData
End Property

Public Property Get iDetPagFrete () AS Integer
    iDetPagFrete= mvariDetPagFrete
End Property

Public Property Let iSemDataSaida (ByVal vData As Integer)
    mvariSemDataSaida = vData
End Property

Public Property Get iSemDataSaida () AS Integer
    iSemDataSaida= mvariSemDataSaida
End Property

Public Property Let dVolumeTotal (ByVal vData As Double)
    mvardVolumeTotal = vData
End Property

Public Property Get dVolumeTotal () AS Double
    dVolumeTotal= mvardVolumeTotal
End Property

Public Property Let iMoeda (ByVal vData As Integer)
    mvariMoeda = vData
End Property

Public Property Get iMoeda () AS Integer
    iMoeda= mvariMoeda
End Property

Public Property Let dTaxaMoeda (ByVal vData As Double)
    mvardTaxaMoeda = vData
End Property

Public Property Get dTaxaMoeda () AS Double
    dTaxaMoeda= mvardTaxaMoeda
End Property

Public Property Let sMensagemCorpoNota (ByVal vData As String)
    mvarsMensagemCorpoNota = vData
End Property

Public Property Get sMensagemCorpoNota () AS String
    sMensagemCorpoNota= mvarsMensagemCorpoNota
End Property

Public Property Let sNatureza (ByVal vData As String)
    mvarsNatureza = vData
End Property

Public Property Get sNatureza () AS String
    sNatureza= mvarsNatureza
End Property

Public Property Let iNaoImpCobranca (ByVal vData As Integer)
    mvariNaoImpCobranca = vData
End Property

Public Property Get iNaoImpCobranca () AS Integer
    iNaoImpCobranca= mvariNaoImpCobranca
End Property

Public Property Let iRPS (ByVal vData As Integer)
    mvariRPS = vData
End Property

Public Property Get iRPS () AS Integer
    iRPS= mvariRPS
End Property

Public Property Let lNumNFe (ByVal vData As Long)
    mvarlNumNFe = vData
End Property

Public Property Get lNumNFe () AS Long
    lNumNFe= mvarlNumNFe
End Property

Public Property Let sCodVerificacaoNFe (ByVal vData As String)
    mvarsCodVerificacaoNFe = vData
End Property

Public Property Get sCodVerificacaoNFe () AS String
    sCodVerificacaoNFe= mvarsCodVerificacaoNFe
End Property

Public Property Let iMexiDataEmissao (ByVal vData As Integer)
    mvariMexiDataEmissao = vData
End Property

Public Property Get iMexiDataEmissao () AS Integer
    iMexiDataEmissao= mvariMexiDataEmissao
End Property

Public Property Let sRazaoSocialCli (ByVal vData As String)
    mvarsRazaoSocialCli = vData
End Property

Public Property Get sRazaoSocialCli () AS String
    sRazaoSocialCli= mvarsRazaoSocialCli
End Property

Public Property Let sNomeReduzidoCli (ByVal vData As String)
    mvarsNomeReduzidoCli = vData
End Property

Public Property Get sNomeReduzidoCli () AS String
    sNomeReduzidoCli= mvarsNomeReduzidoCli
End Property

Public Property Let lNumIntNF (ByVal vData As Long)
    mvarlNumIntNF = vData
End Property

Public Property Get lNumIntNF () AS Long
    lNumIntNF= mvarlNumIntNF
End Property

Public Property Let lLote (ByVal vData As Long)
    mvarlLote = vData
End Property

Public Property Get lLote () AS Long
    lLote= mvarlLote
End Property

Public Property Let iFilialEmpresa (ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa () AS Integer
    iFilialEmpresa= mvariFilialEmpresa
End Property

Public Property Let dtData (ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData () AS Date
    dtData= mvardtData
End Property

Public Property Let dHora (ByVal vData As Double)
    mvardHora = vData
End Property

Public Property Get dHora () AS Double
    dHora= mvardHora
End Property

Public Property Let sUsuario (ByVal vData As String)
    mvarsUsuario = vData
End Property

Public Property Get sUsuario () AS String
    sUsuario= mvarsUsuario
End Property

Public Property Let sRazaoSocialForn (ByVal vData As String)
    mvarsRazaoSocialForn = vData
End Property

Public Property Get sRazaoSocialForn () AS String
    sRazaoSocialForn= mvarsRazaoSocialForn
End Property

Public Property Let sNomeReduzidoForn (ByVal vData As String)
    mvarsNomeReduzidoForn = vData
End Property

Public Property Get sNomeReduzidoForn () AS String
    sNomeReduzidoForn= mvarsNomeReduzidoForn
End Property

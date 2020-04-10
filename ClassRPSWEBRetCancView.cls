VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
End
Attribute VB_Name = "ClassRPSWEBRetCancView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level", "; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarsCorrecao AS String
Private mvarsMsg AS String
Private mvarsCodMsg AS String
Private mvardtData AS Date
Private mvardHora AS Double
Private mvariFilialEmpresa AS Integer
Private mvarlNumIntNF AS Long
Private mvarsVersao AS String
Private mvaritpAmb AS Integer
Private mvarsCodigoCancelamento AS String
Private mvarsId AS String
Private mvarsCNPJ AS String
Private mvarlCodigoMunicipio AS Long
Private mvarsInscricaoMunicipal AS String
Private mvarsNumero AS String
Private mvarsIdConfirmacao AS String
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
Private mvariFilialPedido AS Integer
Private mvarlNumPedidoVenda AS Long
Private mvarsNumPedidoTerc AS String
Private mvardValorTotal AS Double
Private mvardValorProdutos AS Double
Private mvardValorFrete AS Double
Private mvardValorSeguro AS Double
Private mvardValorOutrasDespesas AS Double
Private mvardValorDesconto AS Double
Private mvariCodTransportadora AS Integer
Private mvarsCodUsuarioCancel AS String
Private mvarsMotivoCancel AS String
Private mvardHoraEntrada AS Double
Private mvardHoraSaida AS Double
Private mvarsOrigemNF AS String
Private mvarsDestinoNF AS String
Private mvarsNatureza AS String
Private mvariNaoImpCobranca AS Integer
Private mvariRPS AS Integer
Private mvarlNumNFe AS Long
Private mvarsCodVerificacaoNFe AS String
Private mvarsRazaoSocialCli AS String
Private mvarsNomeReduzidoCli AS String
Private mvarsRazaoSocialForn AS String
Private mvarsNomeReduzidoForn AS String

Public Property Let sCorrecao (ByVal vData As String)
    mvarsCorrecao = vData
End Property

Public Property Get sCorrecao () AS String
    sCorrecao= mvarsCorrecao
End Property

Public Property Let sMsg (ByVal vData As String)
    mvarsMsg = vData
End Property

Public Property Get sMsg () AS String
    sMsg= mvarsMsg
End Property

Public Property Let sCodMsg (ByVal vData As String)
    mvarsCodMsg = vData
End Property

Public Property Get sCodMsg () AS String
    sCodMsg= mvarsCodMsg
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

Public Property Let iFilialEmpresa (ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa () AS Integer
    iFilialEmpresa= mvariFilialEmpresa
End Property

Public Property Let lNumIntNF (ByVal vData As Long)
    mvarlNumIntNF = vData
End Property

Public Property Get lNumIntNF () AS Long
    lNumIntNF= mvarlNumIntNF
End Property

Public Property Let sVersao (ByVal vData As String)
    mvarsVersao = vData
End Property

Public Property Get sVersao () AS String
    sVersao= mvarsVersao
End Property

Public Property Let itpAmb (ByVal vData As Integer)
    mvaritpAmb = vData
End Property

Public Property Get itpAmb () AS Integer
    itpAmb= mvaritpAmb
End Property

Public Property Let sCodigoCancelamento (ByVal vData As String)
    mvarsCodigoCancelamento = vData
End Property

Public Property Get sCodigoCancelamento () AS String
    sCodigoCancelamento= mvarsCodigoCancelamento
End Property

Public Property Let sId (ByVal vData As String)
    mvarsId = vData
End Property

Public Property Get sId () AS String
    sId= mvarsId
End Property

Public Property Let sCNPJ (ByVal vData As String)
    mvarsCNPJ = vData
End Property

Public Property Get sCNPJ () AS String
    sCNPJ= mvarsCNPJ
End Property

Public Property Let lCodigoMunicipio (ByVal vData As Long)
    mvarlCodigoMunicipio = vData
End Property

Public Property Get lCodigoMunicipio () AS Long
    lCodigoMunicipio= mvarlCodigoMunicipio
End Property

Public Property Let sInscricaoMunicipal (ByVal vData As String)
    mvarsInscricaoMunicipal = vData
End Property

Public Property Get sInscricaoMunicipal () AS String
    sInscricaoMunicipal= mvarsInscricaoMunicipal
End Property

Public Property Let sNumero (ByVal vData As String)
    mvarsNumero = vData
End Property

Public Property Get sNumero () AS String
    sNumero= mvarsNumero
End Property

Public Property Let sIdConfirmacao (ByVal vData As String)
    mvarsIdConfirmacao = vData
End Property

Public Property Get sIdConfirmacao () AS String
    sIdConfirmacao= mvarsIdConfirmacao
End Property

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

Public Property Let iFilialPedido (ByVal vData As Integer)
    mvariFilialPedido = vData
End Property

Public Property Get iFilialPedido () AS Integer
    iFilialPedido= mvariFilialPedido
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

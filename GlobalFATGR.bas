Attribute VB_Name = "GlobalFATGR"
Option Explicit
'Constante para nFiscal
Public Const NUMINTNOTA_COMPROVANTE = 0

'Constante para a tabela TipoEmbalagem
Public Const STRING_TIPOEMBALAGEM_DESCRICAO = 50

'Constante para a tabela OrigemDestino
Public Const STRING_ORIGEMDESTINO_ORIGEMDESTINO = 50
Public Const STRING_ORIGEMDESTINO_UF = 2
Public Const STRING_ESTADOS_SIGLA = 2

'Constantes para tabela ItemServico
Public Const STRING_ITEMSERVICO_DESCRICAO = 100
Public Const STRING_ITEMSERVICO_DOCUMENTO = 100

'Constantes para tabela TipoContainer
'Por Tulio em 10/12/01
Public Const STRING_TIPOCONTAINER_DESCRICAO = 100
'Por Cyntia em 24/06/02
Public Const STRING_SOLICITACAO_PORTO = 40

'Constantes para a tabela ProgNavio
Public Const STRING_PROGNAVIO_VIAGEM = 10
Public Const STRING_PROGNAVIO_NAVIO = 50
Public Const STRING_PROGNAVIO_TERMINAL = 50
Public Const STRING_PROGNAVIO_ARMADOR = 50
Public Const STRING_PROGNAVIO_AGMARITIMA = 50
Public Const STRING_PROGNAVIO_OBSERVACAO = 255
    
'Constantes para a tabela Despachante
'Public Const NUM_MAX_CONTATOS = 200
Public Const STRING_DESPACHANTE_CGC = 14
Public Const STRING_DESPACHANTE_NOME = 50
Public Const STRING_DESPACHANTE_NOMEREDUZIDO = 20
Public Const STRING_CONTATO_CONTATO = 50
Public Const STRING_CONTATO_FAX = 18
Public Const STRING_CONTATO_EMAIL = 50
Public Const STRING_CONTATO_TELEFONE = 18
Public Const STRING_CONTATO_SETOR = 50
Public Const DESPACHANTE = 1

'Constantes para CompServico
Public Const NUM_MAX_ITEMSERVICO = 100
Public Const STRING_MATERIAL = 50
Public Const STRING_TARA = 20
Public Const STRING_CODCONTAINER = 20
Public Const STRING_LACRE = 20
Public Const STRING_DOCUMENTO = 100
Public Const STRING_DOC_NUMERO = 20
Public Const STRING_MOTORISTA = 20
Public Const STRING_PLACA = 10
Public Const SOLICITACAO = "solicitacao"
Public Const COMPROVANTE = "comprovante"
Public Const DELAY_DEMURRAGE = 9
Public Const QUANTIDADE_DEFAULT = 1
Public Const STATUS_FATURAVEL = 2
Public Const STATUS_CONCLUIDO = 1

'constantes para a TabPreco
Public Const STRING_TABPRECO_OBSERVACAO = 255
Public Const STRING_TABPRECOITENS_PRODUTO = 20
Public Const STRING_PRODUTOS_DESCRICAO = 50
Public Const STRING_CLIENTES_NOMEREDUZIDO = 20
Public Const STRING_PRODUTOS_CODIGO = 20

'Constantes para Documento
Public Const STRING_DOCUMENTO_DESCRICAO = 100
Public Const STRING_DOCUMENTO_NOMEREDUZIDO = 20
Public Const DOCUMENTO_EXTERNO = 1
Public Const DOCUMENTO_INTERNO = 0
Public Const STRING_DOCUMENTO_DOCUMENTO = 100

'Constantes para PropostaCotacao
'Por Tulio em 30/01/2002
Public Const NUM_MAX_SERVICOS = 100
Public Const NUM_MAX_DESTORIGEM = 50
Public Const NUM_MAX_CONTAINER = 50
Public Const ABERTA = "Aberta"
Public Const APROVADA = "Aprovada"
Public Const PERDIDA = "Perdida"
Public Const STRING_COTACAO_CLIENTE = 50
Public Const STRING_COTACAO_DESCCARGASOLTA = 255
Public Const STRING_COTACAO_ENVIOCOMPLEMENTO = 100
Public Const STRING_COTACAO_INDICACAO = 50
Public Const STRING_COTACAO_OBSDESTORIGEM = 250

Public Const STRING_COTACAO_OBSERVACAO = 255
Public Const STRING_COTACAO_OBSRESULTADO = 255
Public Const TIPO_DOC_ORIGEM_COTACAO = 2
Public Const STRING_COTACAOORIGEMDESTINO_ORIGEM = 50
Public Const STRING_COTACAOORIGEMDESTINO_DESTINO = 50
Public Const STRING_COTACAOORIGEMDESTINO_SERVICO = 100
Public Const STRING_COTACAOSERVICO_PRODUTO = 20
Public Const STRING_PRODUTO_CODIGO = 21
Public Const INCLUI_CARGA = 1
Public Const INCLUI_DESCARGA = 1
Public Const INCLUI_OVA = 1
Public Const INCLUI_DESOVA = 1
Public Const CARGA_SOLTA = 1

'Constantes para Tela de SolicitacaoServico
Public Const STRING_SOLICITACAOSERVICO_NUMREFERENCIA = 50
Public Const STRING_SOLICITACAOSERVICO_MATERIAL = 50
Public Const STRING_SOLICITACAOSERVICO_UM = 20
Public Const STRING_SOLICITACAOSERVICO_BOOKING = 20
Public Const STRING_SOLICITACAOSERVICO_OBSERVACAO = 255
Public Const STRING_ENDERECO_BAIRRO = 12
Public Const STRING_ENDERECO_CEP = 8
Public Const STRING_ENDERECO_CIDADE = 15
Public Const STRING_ENDERECO_CONTATO = 50
Public Const STRING_ENDERECO_EMAIL = 50
Public Const STRING_ENDERECO_ENDERECO = 40
Public Const STRING_ENDERECO_FAX = 18
Public Const STRING_ENDERECO_SIGLAESTADO = 2
Public Const STRING_ENDERECO_TELEFONE1 = 18
Public Const STRING_ENDERECO_TELEFONE2 = 18
Public Const STRING_SOLSERVSERVICO_PRODUTO = 20

'Type para endereços
'Feito po Cyntia em 01/02/02
Type typeEnderecos

    lCodigo As Long
    sEndereco As String
    sBairro As String
    sCidade As String
    sSiglaEstado As String
    iCodigoPais As Integer
    sCEP As String
    sTelefone1 As String
    sTelefone2 As String
    sEmail As String
    sFax As String
    sContato As String
    
End Type

'Type da tabela ProgNavio
Type TypeProgNavio
    lCodigo As Long
    sNavio As String
    sTerminal As String
    sArmador As String
    sAgMaritima As String
    sViagem As String
    sObservacao As String
    dtDataChegada As Date
    dtDataDeadLine As Date
    dHoraChegada As Double
    dHoradeadLine As Double
End Type

Type typeSolicitacaoServico
    lNumIntDoc As Long
    iFilialEmpresa As Integer
    lNumero As Long
    lCodTabPreco As Long
    lCliente As Long
    dtDataEmissao As Date
    sNumReferencia As String
    dtDataPedido As Date
    iTipoOperacao As Integer 'O=Importação 1=Exportação 2=Mercado Interno
    iDespachante As Integer 'O=Importação 1=Exportação 2=Mercado Interno
    sMaterial As String
    dQuantMaterial As Double
    sUM As String
    dValorMercadoria As Double
    iTipoEmbalagem As Integer
    iTipoContainer As Integer
    lCodProgNavio As Long
    sBooking As String
    dtDataPrevInicio As Date
    dHoraPrevInicio As Double
    dtDataPrevFim As Date
    dHoraPrevFim As Double
    sObservacao As String
    sPorto As String
    lEnderecoOrigem As Long
    lEnderecoDestino As Long
End Type

Type typeCotacaoGR
    lNumIntDoc As Long
    iFilialEmpresa As Integer
    lCodigo As Long
    dtData As Date
    dValorMerc As Double
    iTipoOperacao As Integer 'O=Importação 1=Exportação 2=Mercado Interno
    sCliente As String
    iEnvio As Integer 'O=Carta 1=E-mail 2=Fax 3=Telefone
    sEnvioComplemento As String
    iCodVendedor As Integer
    sIndicacao As String
    sObservacao As String
    sObsDestOrigem As String
    dtDataPrevInicio As Date
    iTipoEmbalagem As Integer
    iAjudantes As Integer
    iCarga As Integer
    iCargaPorConta As Integer
    iDesCarga As Integer
    iDesCargaPorConta As Integer
    iOva As Integer
    iOvaPorConta As Integer
    iDesova As Integer
    iDesovaPorConta As Integer
    iCargaSolta As Integer
    sDescCargaSolta As String
    iCondicaoPagto As Integer
    iSituacao As Integer
    iJustificativa As Integer
    sObsResultado As String
    colCotacaoOrigemDestino As New Collection
    colCotacaoContainer As New Collection
    colCotacaoServico As New Collection
    colContato As New Collection
End Type

Type typeCotacaoContainer
    lNumIntDoc As Long
    lNumIntDocOrigem As Long
    iTipoContainer As Integer
    iQuantidade As Integer
End Type

Type typeCotacaoOrigemDestino
    lNumIntDoc As Long
    lNumIntDocOrigem As Long
    sServico As String
    sOrigem As String
    sDestino As String
End Type

Type typeCotacaoServico
    lNumIntDoc As Long
    lNumIntDocOrigem As Long
    sProduto As String
    dQuantidade As Double
    dPrecoUnitario As Double
    dAdValoren As Double
    dPedagio As Double
    iOrigem As Integer
    iDestino As Integer
End Type

Type typeDespachante
    iCodigo As Integer
    sCGC As String
    sNome As String
    sNomeReduzido As String
    lEndereco As Long
    objEndereco As ClassEndereco
    colContato As New Collection
End Type

Type typeDocumento
    iCodigo As Integer
    sDescricao As String
    sNomeReduzido As String
    iTipoDoc As Integer '0=Interno 1=Externo
    sDocumento As String
End Type

Type typeServItemServ
    sProduto As String
    iCodItemServico As Integer
    iOrdem As Integer
End Type

Type typeTabPreco
    lCodigo As Long
    lCliente As Long
    iOrigem As Integer
    iDestino As Integer
    dPedagio As Double
    dAdValoren As Double
    sObservacao As String
    colTabPrecoItens As New Collection
    dtDataVigencia As Date
End Type

Type typeTabPrecoItens
    lCodTabela As Long
    sProduto As String
    dtDataVigencia As Date
    dPreco As Double
    sDescricao As String
End Type

Type typeCompServ
    lNumIntDoc As Long
    lNumIntDocOrigem As Long
    iFilialEmpresa As Integer
    lCodigo As Long
    dtDataEmissao As Date
    sProduto As String
    dQuantMaterial As Double
    sUM As String
    dValorMercadoria As Double
    dFretePeso As Double
    dPedagio As Double
    dAdValoren As Double
    dtDataDemurrage As Date
    sCodigoContainer As String
    dTara As Double
    sLacre As String
    sObservacao As String
    lNumIntNota As Long
    iSituacao As Integer
    dQuantidade As Double
    dValorContainer As Double
    sCidadeDestino As String
    sCidadeOrigem As String
    sUFDestino As String
    sUFOrigem As String
    sEmbalagem As String
    sMaterial As String
End Type

Type typeCompServItem
    lNumIntDoc As Long
    lNumIntDocOrigem As Long
    iCodItemServico As Integer
    dtDataPrev As Date
    dHoraPrev As Double
    dtDataInicio As Date
    dHoraInicio As Double
    dtDataFim As Date
    dHoraFim As Double
    iDocIntTipo As Integer
    sDocIntNumero As String
    dtDocIntDataEmissao As Date
    iDocExtTipo As Integer
    sDocExtNumero As String
    dtDocExtDataEmissao As Date
    dtDocExtDataRec As Date
    dDocExtHoraRec As Double
    sPlacaCaminhao As String
    sPlacaCarreta As String
    sMotorista As String
    sObservacao As String
End Type

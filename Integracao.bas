Attribute VB_Name = "Integracao"
Option Explicit

Public Const STRING_TABELA = 50
Public Const STRING_SIGLA_ARQ = 10
Public Const STRING_NOME_EXTERNO_INT = 50

Public Const ROTINA_IMPORTACAO_DADOS = 1
Public Const ROTINA_EXPORTACAO_DADOS = 2

Public Const TIPO_INTEGRACAO_IMPORTACAO = 1
Public Const TIPO_INTEGRACAO_EXPORTACAO = 2

Public Const EXPORTAR_DADOS_TODOS = 1
Public Const EXPORTAR_DADOS_TODOS_NAO_EXPORTADOS = 2
Public Const EXPORTAR_DADOS_POR_PERIODO = 3

Public Const STRING_INT_TIPOCLIENTE_DESC = 50

Public Const INT_ETAPA_EXP_TABINT = 1
Public Const INT_ETAPA_EXP_GERARQ = 2
Public Const INT_ETAPA_IMP_TABINT = 3
Public Const INT_ETAPA_IMP_ATUALI = 4

Public Const TIPO_ARQ_CLIENTE = 1
Public Const TIPO_ARQ_PV = 2
Public Const TIPO_ARQ_SLDPROD = 3

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeArqExportacao
    lNumIntArq As Long
    sNomeArquivo As String
    iTipoArq As Integer
    dtDataExportacao As Date
    dHoraExportacao As Double
    sUsuario As String
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeTipoArqIntegracao
    iCodigo As Integer
    sDescricao As String
    sSiglaArq As String
    sTabela As String
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeArqExportacaoAux
    lNumIntGeracao As Long
    dtDataGeracao As Date
    dHoraGeracao As Double
    sUsuario As String
    iExportar As Integer
    dtExpDataDe As Date
    dtExpDataAte As Date
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeArqImportacao
    lNumIntArq As Long
    iTipoArq As Integer
    sNomeArquivo As String
    dtDataImportacao As Date
    dHoraImportacao As Double
    dtDataAtualizacao As Date
    dHoraAtualizacao As Double
    sUsuario As String
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeIntegracaoCliente
    lNumIntDoc As Long
    lNumIntGer As Long
    lNumIntArq As Long
    lSeqRegistro As Long
    iTipoInt As Integer
    lCodCliente As Long
    iCodFilial As Integer
    sRazaoSocial As String
    sNomeReduzido As String
    iTipo As Integer
    sObservacao As String
    dLimiteCredito As Double
    iCondicaoPagto As Integer
    iAtivo As Integer
    sFilialNome As String
    sCgc As String
    sRG As String
    sInscricaoEstadual As String
    sInscricaoMunicipal As String
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
    sEnderecoEnt As String
    sBairroEnt As String
    sCidadeEnt As String
    sSiglaEstadoEnt As String
    iCodigoPaisEnt As Integer
    sCEPEnt As String
    sTelefone1Ent As String
    sTelefone2Ent As String
    sEmailEnt As String
    sFaxEnt As String
    sContatoEnt As String
    sEnderecoCobr As String
    sBairroCobr As String
    sCidadeCobr As String
    sSiglaEstadoCobr As String
    iCodigoPaisCobr As Integer
    sCEPCobr As String
    sTelefone1Cobr As String
    sTelefone2Cobr As String
    sEmailCobr As String
    sFaxCobr As String
    sContatoCobr As String
    iComErro As Integer
    dtDataAtualizacao As Date
    iVendedor As Integer
    sObservacaoFilial As String
    lCodExterno As Long
    sTipoCliente As String
    sLogradouro As String
    sComplemento As String
    sTipoLogradouro As String
    sEmail2 As String
    lNumero As Long
    iTelDDD1 As Integer
    iTelDDD2 As Integer
    iFaxDDD As Integer
    sTelNumero1 As String
    sTelNumero2 As String
    sFaxNumero As String
    sLogradouroEnt As String
    sComplementoEnt As String
    sTipoLogradouroEnt As String
    sEmail2Ent As String
    lNumeroEnt As Long
    iTelDDD1Ent As Integer
    iTelDDD2Ent As Integer
    iFaxDDDEnt As Integer
    sTelNumero1Ent As String
    sTelNumero2Ent As String
    sFaxNumeroEnt As String
    iTabelaPreco As Integer
    sReferencia As String
    sReferenciaEnt As String
End Type


'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeIntegracaoPV
    lNumIntDoc As Long
    lNumIntGer As Long
    lNumIntArq As Long
    lSeqRegistro As Long
    iTipoInt As Integer
    lCodPedidoExterno As Long
    dtDataEmissao As Date
    lCodClienteExterno As Long
    lCodClienteCorporator As Long
    iCodFilialCorporator As Integer
    sCGCCliente As String
    sNaturezaOp As String
    iCodTabelaPreco As Integer
    iCodCondPagto As Integer
    sNomeCondPagtoExterno As String
    iCodCondPagtoExterno As Integer
    iFilialEmpresaFat As Integer
    dValorDescontoPedido As Double
    iFrete As Integer
    dValorFretePedido As Double
    dValorSeguroPedido As Double
    dValorOutrasDespesasPedido As Double
    iFilialEmpresaEnt As Integer
    iTrazerTranspAuto As Integer
    iCodTransportadora As Integer
    iCodTransportadoraExterno As Integer
    sNomeTransportadoraExterno As String
    iTrazerMensagemAuto As Integer
    sMensagemPedido As String
    iTrazerPesoAuto As Integer
    dPesoBruto As Double
    dPesoLiquido As Double
    iTrazerCanalVendaAuto As Integer
    iCanalVenda As Integer
    iCodVendedor As Integer
    iCodVendedorExterno As Integer
    sNomeVendedorExterno As String
    iTrazerComissaoAuto As Integer
    dValorComissao As Double
    iTrazerReservaAuto As Integer
    iCodAlmoxarifado As Integer
    iCodAlmoxarifadoExterno As Integer
    sNomeAlmoxarifadoExterno As String
    iItem As Integer
    sCodProduto As String
    sCodProdutoExterno As String
    iTrazerDescricaoAuto As Integer
    sDescricaoItem As String
    dQuantidadePedida As Double
    dQuantidadeCancelada As Double
    sUM As String
    dPrecoUnitario As Double
    dValorDescontoItem As Double
    dtDataEntrega As Date
    iComErro As Integer
    dtDataAtualizacao As Date
    iCodTabelaPrecoItem As Integer
    sPedRepr As String
End Type

Type typeIntegracaoSldProd
    lNumIntDoc As Long
    lNumIntGer As Long
    lNumIntArq As Long
    lSeqRegistro As Long
    iTipoInt As Integer
    sCodProduto As String
    iAlmoxarifado As Integer
    dSaldoDisp As Double
    iComErro As Integer
    dtDataAtualizacao As Date
End Type



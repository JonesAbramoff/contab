Attribute VB_Name = "GlobalMAT2"
''Option Explicit

''Public Const STRING_PEDIDOCOTACAO_CONTATO = 50
''Public Const STRING_COTACAO_DESCRICAO = 50
''Public Const STRING_COTACAOPRODUTO_PRODUTO = 20
''Public Const STRING_COTACAOPRODUTO_UM = 5
''
''Public Const STRING_NOME_FILIAL_FORN = 50
''Public Const STRING_NOME_FILIAL_CLI = 50
''
'''codigos de TiposDocInfo ligados a Compras
''''''Public Const DOCINFO_NRFPCO = 80
''''''Public Const DOCINFO_NRFFCO = 81
''Public Const DOCINFO_NFEENCO = 82
''Public Const DOCINFO_NFIENCO = 83
''Public Const DOCINFO_NFEEFNCO = 84
''Public Const DOCINFO_NFIEFNCO = 85
''Public Const DOCINFO_NFIEIPICO = 86
''Public Const DOCINFO_NFIECPCO = 87
''Public Const DOCINFO_NFIEICMCO = 88
''Public Const DOCINFO_NFEECPCO = 89
''Public Const DOCINFO_NFEEICMCO = 90
''Public Const DOCINFO_NFEEIPICO = 91
''Public Const DOCINFO_NFIEFIPICO = 92
''Public Const DOCINFO_NFIEFICMCO = 93
''Public Const DOCINFO_NFIEFCPCO = 94
''Public Const DOCINFO_NFEEFICMCO = 95
''Public Const DOCINFO_NFEEFCPCO = 96
''Public Const DOCINFO_NFEEFIPICO = 97
''Public Const DOCINFO_NFEEBFCOM = 100
''Public Const DOCINFO_NFIEBFCOM = 101
''Public Const DOCINFO_NFIEFBFCOM = 102
''Public Const DOCINFO_NFEEFBFCOM = 103
''
'''Para saber se vai ser ou não calculado os Parametros  de Ponto Pedido
''Public Const PRODUTOFILIAL_CALCULA_VALORES = 1
''Public Const PRODUTOFILIAL_NAO_CALCULA_VALORES = 0
''
''
''Type typeTipoDeProduto
''    sSiglaUMCompra As String
''    sSiglaUMEstoque As String
''    sSiglaUMVenda As String
''    sDescricao As String
''    sSigla As String
''    sIPICodigo As String
''    sIPICodDIPI As String
''    sISSCodigo As String
''    sContaContabil As String
''    sContaProducao As String
''    iTipo As Integer
''    iClasseUM As Integer
''    iCompras As Integer
''    iControleEstoque As Integer
''    iFaturamento As Integer
''    iKitBasico As Integer
''    iKitInt As Integer
''    iPCP As Integer
''    iPrazoValidade As Integer
''    iIRIncide As Integer
''    iICMSAgregaCusto As Integer
''    iIPIAgregaCusto As Integer
''    iFreteAgregaCusto As Integer
''    iApropriacaoCusto As Integer
''    iIntRessup As Integer
''    iMesesConsumoMedio As Integer
''    iConsideraQuantCotacaoAnterior As Integer
''    iTemFaixaReceb As Integer
''    iRecebFaixaFora As Integer
''    iNatureza As Integer
''    colCategoriaItem As Collection
''    dIPIAliquota As Double
''    dISSAliquota As Double
''    dTempoRessupMax As Double
''    dConsumoMedioMax As Double
''    dResiduo As Double
''    dPercentMaisQuantCotacaoAnterior As Double
''    dPercentMenosQuantCotacaoAnterior As Double
''    dPercentMaisReceb As Double
''    dPercentMenosReceb As Double
''
''End Type
''
''
''Type typeCotacao
''
''    lNumIntDoc As Long
''    iFilialEmpresa As Integer
''    lCodigo As Long
''    sDescricao As String
''    dtData As Date
''    iTipoDestino As Integer
''    lFornCliDestino As Long
''    iFilialDestino As Integer
''    iComprador As Integer
''
''End Type
''
''Type typePedidoCotacao
''    lNumIntDoc As Long
''    iFilialEmpresa As Integer
''    lCodigo As Long
''    lFornecedor As Long
''    iFilial As Integer
''    sContato As String
''    dtDataEmissao As Date
''    dtData As Date
''    dtDataValidade As Date
''    iTipoFrete As Integer
''    iStatus As Integer
''    iCondPagtoPrazo As Integer
''End Type
''
''Type typeCotacaoProduto
''    lNumIntDoc As Long
''    lCotacao As Long
''    sProduto As String
''    dQuantidade As Double
''    sUM As String
''    lFornecedor As Long
''    iFilial As Integer
''End Type
''
''Type typeItemPedCotacao
''    lNumIntDoc As Long
''    sProduto As String
''    dQuantidade As Double
''    sUM As String
''    lCotacaoProduto As Long
''End Type
''
''Type typeItemCotacao
''    lNumIntDoc As Long
''    iCondPagto As Integer
''    dtDataReferencia As Date
''    dPrecoUnitario As Double
''    dOutrasDespesas As Double
''    dValorSeguro As Double
''    dValorDesconto As Double
''    dValorTotal As Double
''    dValorIPI As Double
''    dAliquotaIPI As Double
''    dAliquotaICMS As Double
''    iPrazoEntrega As Integer
''    dQuantEntrega As Double
''    lObservacao As Long
''    dValorFrete As Double
''End Type
''
''Type typeOrdemProducao
''    iFilialEmpresa As Integer
''    sCodigo As String
''    dtDataEmissao As Date
''    iNumItens As Integer
''    iNumItensBaixados As Integer
''    iGeraReqCompra As Integer
''    iGeraOP As Integer
''End Type
''
''Type typeFornecedorProdutoFF
''    iFilialEmpresa As Integer
''    sProduto As String
''    lFornecedor As Long
''    iFilialForn As Integer
''    sProdutoFornecedor As String
''    dLoteMinimo As Double
''    iNota As Integer
''    dQuantPedAbertos As Double
''    dtDataUltimaCompra As Date
''    iTempoRessup As Integer
''    dQuantPedida As Double
''    dQuantRecebida As Double
''    dtDataPedido As Date
''    dtDataReceb As Date
''    dPrecoTotal As Double
''    dUltimaCotacao As Double
''    dtDataUltimaCotacao As Date
''    iTipoFreteUltimaCotacao As Integer
''    dQuantUltimaCotacao As Double
''    iPadrao As Integer
''    iCondPagto As Integer
''    sCondPagto As String
''End Type
''

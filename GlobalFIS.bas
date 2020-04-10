Attribute VB_Name = "GlobalFIS"
Option Explicit

'SPED FISCAL
Public Const TITULO_SPED_FISCAL = "Sped Fiscal"
Public Const ROTINA_SPED_FISCAL = 2
Public Const ROTINA_SPED_FISCAL_PIS = 3
Public Const ROTINA_ECF = 4

Public Const REG_APUR_PIS_COFINS_TIPO_PIS = 1
Public Const REG_APUR_PIS_COFINS_TIPO_COFINS = 2

Public Const EFD_TABELAS_OBRIG_PISCOFINS = 1
Public Const EFD_TABELAS_OBRIG_ICMS = 2
Public Const EFD_TABELAS_OBRIG_IPI = 3
Public Const EFD_TABELAS_OBRIG_CONTRIBPREV = 4

Public Const STRING_GUIASICMS_NUMERO = 12
Public Const STRING_GUIASICMS_ORGAOARRECADADOR = 20
Public Const STRING_GUIASICMS_LOCALENTREGA = 50
Public Const STRING_GUIASICMS_CODRECEITA = 50
Public Const STRING_GUIASICMS_CODOBRIGRECOLHER = 3

'Seções para GERAIN86
Public Const STRING_IN86TIPOSARQUIVOS_DESCRICAO = 50
Public Const STRING_IN86TIPOSARQUIVOS_PREFIXONOME = 20
Public Const STRING_IN86TIPOSARQUIVOS_ROTINAGERADORA = 50
Public Const STRING_IN86TIPOSARQUIVOS_LAYOUTARQUIVO = 100
Public Const STRING_IN86MEIOENTREGA_DESCRICAO = 50
Public Const STRING_IN86MODELOS_MODELO = 50
Public Const STRING_IN86ARQUIVOS_NOME = 100
Public Const PATH_IN86 = "\IN86\"
Public Const TIPO_ARQUIVO_AUXILIAR = 1
Public Const ARQUIVO_NOME_LEIAME = "LEIAME"
Public Const TITULO_TELABATCH_GERACAO_ARQ_IN86 = "Geração de Arquivos IN86"
Public Const ROTINA_ARQIN86_BATCH = 1

Public Const LIVREGES_SAIDA_TODAS = 0
Public Const LIVREGES_SAIDA_NACIONAL = 1
Public Const LIVREGES_SAIDA_EXTERNA = 2

Public Const LIVREGES_ENTRADA_TODAS = 0
Public Const LIVREGES_ENTRADA_NACIONAL = 1
Public Const LIVREGES_ENTRADA_EXTERNA = 2

'Constantes utilizadas na geração de relatórios IN86
Public Const QUANT_REGISTROS_DUMP = 30
Public Const ARQUIVO_LOCAL_DATA = "Local e Data: "
Public Const ARQUIVO_RESPONSAVEL = "Responsável: "
Public Const ARQUIVO_LINHA_GRANDE = "___________________________________________"
Public Const ARQUIVO_ERRO_GERACAO_LAYOUT = "Erro na geração do Relatório de Layout referente ao "
Public Const ARQUIVO_ERRO_GERACAO_ACOMPANHAMENTO = "Erro na geração do Relatório de Acompanhamento referente ao Arquivo "
Public Const ARQUIVO_CNPJ = "CNPJ: "
Public Const ARQUIVO_NOME_EMPRESA = "Nome Empresarial: "
Public Const ARQUIVO_NOME_ARQ = "Nome do Arquivo: "
Public Const ARQUIVO_DATA_GERACAO = "Data da Geração: "
Public Const ARQUIVO_CONTEUDO_ARQ = "Conteúdo do Arquivo: "
Public Const ARQUIVO_MEIO_FISICO = "Meio Físico de Entrega: "
Public Const ARQUIVO_DUMP_PRIMEIROS = "Dump dos Trinta Primeiros Registros: "
Public Const ARQUIVO_DUMP_ULTIMOS = "Dump dos Trinta Últimos Registros: "
Public Const ARQUIVO_QTD_PAGINAS = " (quantidade de páginas)"
Public Const ARQUIVO_DESCRICAO_DETALHADA = "Descrição Detalhada do Arquivo: "
Public Const ARQUIVO_QTD_VOLUMES = "Quantidade de volumes: "
Public Const ARQUIVO_QTD_REGISTROS = "Quantidade de Registros: "
Public Const ARQUIVO_TAMANHO_ARQ_BYTES = "Tamanho do Arquivo (em Bytes): "
Public Const ARQUIVO_OUTROS = "Outros Parâmetros: "
Public Const ARQUIVO_CONTRIBUINTE_RESP_PREPOSTO = "Contribuinte/Responsável ou Preposto"
Public Const ARQUIVO_RESP_TECNICO = "Responsável Técnico pela Geração do Arquivo"
Public Const ARQUIVO_NOME = "Nome: "
Public Const ARQUIVO_ASSINATURA = "Assinatura: "
Public Const ARQUIVO_CPF = "CPF: "
Public Const ARQUIVO_TELEFONE = "Telefone: "
Public Const ARQUIVO_FAX = "FAX: "
Public Const ARQUIVO_EMAIL = "E-Mail: "
Public Const ARQUIVO_LINHA_PEQUENA = "________"
Public Const ARQUIVO_LINHA_MEDIA = "________________________ "

'Incluído por Luiz Nogueira em 28/01/04
Public Const ARQUIVO_MARGEM_ESQUERDA As String = "          "
Public Const ARQUIVO_TITULO_RELDUMP1 As String = "Relatório de Dump dos XX primeiros registros do arquivo "
Public Const ARQUIVO_TITULO_RELDUMP2 As String = "Relatório de Dump dos XX últimos registros do arquivo "
Public Const ARQUIVO_TITULO_RELACOMPANHAMENTO As String = "Relatório de Acompanhamento"

'Seções para Tipo Apuração
Public Const SECAO_DEBITO = 0
Public Const SECAO_CREDITO = 1
Public Const SECAO_APURACAO = 2

'Indentifica se o tipo Apuracao é Pré-Cadastrada
Public Const TIPOREGAPURACAO_PRE_CADASTRADO = 1

'Se permite lancamento ou não permite lancamento
Public Const TIPOREGAPURACAO_PERMITE_LANCAMENTO = 1
Public Const TIPOREGAPURACAO_NAO_PERMITE_LANCAMENTO = 0

'Se a apuração está aberta
Public Const APURACAO_ICMS_ABERTA = 0
Public Const APURACAO_IPI_ABERTA = 0

'Indica se o Livro está Fechado
Public Const LIVRO_ABERTO = 0

Public Const ARQUIVO_ICMS_ABERTO = 0

'String's para Buffer
Public Const STRING_DESCRICAO_APURACAO = 255
Public Const STRING_LOCALENTREGA = 32
Public Const STRING_OBSERVACAO = 255
Public Const STRING_FILIALEMPRESA_NOME = 35
Public Const STRING_LOGRADOURO = 34
Public Const STRING_COMPLEMENTO = 22
Public Const STRING_CONTATO_REGAPURACAO = 28
Public Const STRING_TELCONTATO = 12
Public Const STRING_MUNICIPIO = 30
Public Const STRING_BAIRRO_REGAPURACAO = 15
Public Const STRING_NOME_REGES = 255
Public Const STRING_CONVENIO = 30
Public Const STRING_NUMERO = 12
Public Const STRING_CODIGONCM = 8
Public Const STRING_ARQICMS_UM = 6
Public Const STRING_ARQICMS_DESCRICAO_PRODUTO = 53

Public Const LIVRO_REG_ENTRADA_ICMS_IPI_CODIGO = 1
Public Const LIVRO_REG_SAIDA_ICMS_IPI_CODIGO = 2
Public Const LIVRO_REG_INVENTARIO_CODIGO = 3
Public Const LIVRO_APURACAO_ICMS_CODIGO = 4
Public Const LIVRO_LISTA_EMITENTES_CODIGO = 5
Public Const LIVRO_CODIGO_MERCADORIAS_CODIGO = 6
Public Const LIVRO_OPERACOES_INTERESTADUAIS_CODIGO = 7
Public Const LIVRO_PRESTACOES_INTERESTADUAIS_CODIGO = 8
Public Const LIVRO_DADOS_RECOLHIMENTO_CODIGO = 9
Public Const LIVRO_APURACAO_IPI_CODIGO = 10
Public Const LIVRO_APURACAO_ISS_CODIGO = 11
Public Const LIVRO_REG_ENTRADA_ISS_CODIGO = 12
Public Const LIVRO_REG_SAIDA_ISS_CODIGO = 13

Public Const TRIBUTO_COM_LIVROFISCAL = 1


'StatusLivro na Tabela de LivRegES
Public Const STATUS_FIS_ORIGINAL = 0
Public Const STATUS_FIS_ALTERADA = 1
Public Const STATUS_FIS_ORIGINAL_EXCLUIDA = 2
Public Const STATUS_FIS_ALTERADA_EXCLUIDA = 3

'Coluna na Tipo Tributacao
Public Const TIPO_TRIBUTADO = 1
Public Const TIPO_ISENTO_NAO_TRIBUTADO = 2
Public Const TIPO_OUTRAS = 3

Public Const SITUACAO_NORMAL = 0

'Tipo na Tabela de LivRegES
Public Const TIPO_REGES_SAIDA = 1
Public Const TIPO_REGES_ENTRADA = 0

'TipoNumIntDocOrigem na Tabela de LivRegES
Public Const TIPO_NUMINTDOC_ORIGEM_NOTAFISCAL = 0

'Número máximo de lançamentos (Grid)
Public Const NUM_MAX_LANCAMENTOS = 100

'Mnemonicos Globais usados no Fis
Public Const MNEUMONICO_FRETE = "CtaFrete"
Public Const MNEUMONICO_SEGURO = "CtaSeguro"
Public Const MNEUMONICO_OUTRAS_DESP = "CtaOutrasDesp"

Public Const IMPRIME_LIVROFISCAL = 1

'Constates do Tipo de Guia
Public Const GNRICMS_TIPO_ARQUIVOICMS = 1
Public Const GNRICMS_TIPO_APURACAOICMS = 2

Public Const STRING_GNRICMS_CODRECEITA = 50
Public Const STRING_GNRICMS_CODOBRIGRECOLHER = 3

Type typeIN86TiposArquivos
    iCodigo As Integer
    sDescricao As String
    sPrefixoNome As String
    sRotinaGeradora As String
    iAuxiliar As Integer
    sLayoutArquivo As String
End Type

Type typeIN86MeioEntrega
    iCodigo As Integer
    iUsaEtiqueta As Integer
    sDescricao As String
End Type

Type typeIN86Modelos
    iCodigo As Integer
    sModelo As String
    dtDataInicio As Date
    dtDataFim As Date
    iMeioEntrega As Integer
    iEtiquetas As Integer
End Type

Type typeIN86Arquivos
    iModelo As Integer
    iTipo As Integer
    iSelecionado As Integer
    sNome As String
    iDUMP As Integer
    iRelatAcompanhamento As Integer
    iLayout As Integer
    iFilialEmpresa As Integer
    iNumEtiqueta As Integer
End Type

Type typeTiposRegApuracao

    iCodigo As Integer
    sDescricao As String
    iSecao As Integer
    iPreCadastrado As Integer
    iLancamento As Integer
    
End Type

Type typeRegApuracaoItens
    
    lNumIntDoc As Long
    lNumIntDocApuracao As Long
    iTipoReg As Integer
    sDescricao As String
    dtData As Date
    dValor As Double
    iFilialEmpresa As Integer
    
End Type

Type typeLivrosFilial

    iCodLivro As Integer
    iFilialEmpresa As Integer
    iImprime As Integer
    iNumeroProxLivro As Integer
    iNumeroProxFolha As Integer
    iPeriodicidade As Integer
    dtDataInicial As Date
    dtDataFinal As Date
    dtImpressoEm As Date
    sDescricao As String
    
End Type

Type typeRegApuracao

    lNumIntDoc As Long
    iFilialEmpresa As Integer
    dtDataInicial As Date
    dtDataFinal As Date
    lNumIntDocLivFechado As Long
    dSaldoCredorInicial As Double
    dSaldoCredorFinal As Double
    dtDataEntregaGIA As Date
    sLocalEntregaGIA As String
    sObservacoes As String
    sCgc As String
    sInscricaoEstadual As String
    sNome As String
    sMunicipio As String
    sUF As String
    sLogradouro As String
    lNumero As Long
    sComplemento As String
    sBairro As String
    sCEP As String
    sContato As String
    sTelContato As String
    iNumeroLivro As Integer
    dtDataImpressao As Date
    iFolhaInicial As Integer
    
End Type

Type typeGNRICMS

    lNumIntDoc As Long
    lCodigo As Long
    iTipo As Integer
    dtDataPagto As Date
    sCGCSubstTrib As String
    sInscricaoEstadual As String
    sUFSubstTrib As String
    sUFDestino As String
    iBanco As Integer
    iAgencia As Integer
    sNumero As String
    dValor As Double
    dtDataVencimento As Date
    dtDataRef As Date
    sConvenio As String
    lNumIntRegApuracaoICMS As Long
    lNumIntArqICMS As Long
    sCodReceita As String
    sCodObrigRecolher As String

End Type

Type typeLivRegES

    lNumIntDocOrigem As Long
    lNumIntDoc As Long
    lNumIntLivroFechado As Long
    lNumIntNF As Long
    iStatusLivro As Integer
    iTipo As Integer
    sCgc As String
    sInscricaoEstadual As String
    sNome As String
    dtData As Date
    sUF As String
    sSerie As String
    lNumNotaFiscal As Long
    iSituacao As Integer
    iEmitente As Integer
    iDestinatario As Integer
    iOrigem As Integer
    lNumIntEmitente As Long
    lNumIntArqICMS As Long
    lNumIntRegApuracaoICMS As Long
    lNumIntRegApuracaoIPI As Long
    iModelo As Integer
    iTipoNumIntDocOrigem As Integer
    iFilialEmpresa As Integer
    dtDataEmissao As Date
    iCIF_FOB As Integer
    dPISValor As Double
    dCOFINSValor As Double
    iIEIsento As Integer
End Type


Type typeLivRegESLinha

    lNumIntDoc As Long
    lNumIntDocRegES As Long
    sNaturezaOp As String
    dValorTotal As Double
    dValorICMSBase As Double
    dValorICMS As Double
    dValorICMSIsentoNaoTrib As Double
    dValorICMSOutras As Double
    dValorICMSSubstBase As Double
    dValorICMSSubstRet As Double
    dAliquotaICMS As Double
    dValorDespAcess As Double
    dValorIPI As Double
    dAliquotaIPI As Double
    dValorIPIBase As Double
    dValorIPIIsentoNaoTrib As Double
    dValorIPIOutras As Double
    sClassifContabil As String
    dValorContabil As Double
    sObservacaoLivFisc As String
    iCodigoICMS As Integer
    iCodigoIPI As Integer

End Type

Type typeLivRegESItemNF

    lNumIntDoc As Long
    lNumIntDocRegES As Long
    iNumItem As Integer
    sCFOP As String
    lNumIntCadProd As Long
    dQuantidade As Double
    dValorProduto As Double
    dValorDescontoDespAcess As Double
    dValorICMSBase As Double
    dValorICMSSubstBase As Double
    dValorIPI As Double
    dAliquotaICMS As Double
    iTipoTribICMS As Integer
    iTipoTribIPI As Integer
    dValorICMS As Double
    dValorIPIBase As Double
    dAliquotaIPI As Double
    dRedBaseICMS As Double
    dRedBaseIPI As Double
    iTipoTributacao As Integer
    dAliquotaSubst As Double
    dValorSubst As Double
    
End Type

Type typeTributo
    
    iCodigo As Integer
    sDescricao As String
    iApuracaoPeriodicidade As Integer
    iLivro As Integer
    
End Type

Type typeLivroFiscal
    
    iCodigo As Integer
    sDescricao As String
    iCodTributo As Integer
    iPeriodicidade As Integer
    iApuracao As Integer

End Type

Type typeLivroFechado
    
    lNumIntDoc As Long
    iCodLivro As Integer
    iFilialEmpresa As Integer
    iNumeroLivro As Integer
    dtDataInicial As Date
    dtDataFinal As Date
    dtDataImpressao As Date
    iFolhaInicial As Integer
    iFolhaFinal As Integer
    
End Type

Type typeRegInventario
    
    iFilialEmpresa As Integer
    sProduto As String
    dtData As Date
    lNumIntDocLivFechado As Long
    sDescricao As String
    sModelo As String
    sIPICodigo As String
    sSiglaUMEstoque As String
    dQuantidadeUMEstoque As Double
    dValorUnitario As Double
    iNatureza As Integer
    dQtdeNossaEmTerc As Double
    dQtdeDeTercConosco As Double
    sObservacoes As String
    sContaContabil As String
    dQuantConserto As Double
    dQuantDemo As Double
    dQuantConsig As Double
    dQuantBenef As Double
    dQuantOutras As Double
    dCustoConserto As Double
    dCustoDemo As Double
    dCustoConsig As Double
    dCustoBenef As Double
    dCustoOutras As Double
    dQuantConserto3 As Double
    dQuantDemo3 As Double
    dQuantConsig3 As Double
    dQuantBenef3 As Double
    dQuantOutras3 As Double
    dCustoConserto3 As Double
    dCustoDemo3 As Double
    dCustoConsig3 As Double
    dCustoBenef3 As Double
    dCustoOutras3 As Double
    dValorEstoque As Double
    dValorBenef As Double
    dValorBenef3 As Double
    dValorConserto As Double
    dValorConserto3 As Double
    dValorConsig As Double
    dValorConsig3 As Double
    dValorDemo As Double
    dValorDemo3 As Double
    dValorOutras As Double
    dValorOutras3 As Double
    
End Type

Type typeLivRegESCadProd

    lNumIntDoc As Long
    lNumIntArqICMS As Long
    sProduto As String
    dtDataInicial As Date
    dtDataFinal As Date
    sCodigoNCM As String
    sDescricao As String
    sSiglaUM As String
    sSituacaoTrib As String
    dAliquotaIPI As Double
    dAliquotaICMS As Double
    dReducaoBaseCalculoICMS As Double
    dBaseCalculoICMSSubst As Double

End Type

Type typeInfoArqICMS

    lNumIntDoc As Long
    dtDataInicial As Date
    dtDataFinal As Date
    sCgc As String
    sInscricaoEstadual As String
    sNome As String
    sMunicipio As String
    sUF As String
    sLogradouro As String
    lNumero As Long
    sComplemento As String
    sBairro As String
    sCEP As String
    sContato As String
    sTelContato As String
    sNomeArquivo As String

End Type

Type typeRegInventarioAlmox
    
    iFilialEmpresa As Integer
    sProduto As String
    dtData As Date
    iAlmoxarifado As Integer
    dQuantidadeUMEstoque As Double
    dQtdeDeTercConosco As Double
    dQuantConsig3 As Double
    dQuantDemo3 As Double
    dQuantConserto3 As Double
    dQuantOutras3 As Double
    dQuantBenef3 As Double
    
End Type

'Constantes para geração dos arquivos do IN86
Public Const STRING_IN86_ALIQUOTA = 5
Public Const STRING_IN86_ITEMNF = 3
Public Const STRING_IN86_COD_GENERICO = 14
Public Const STRING_IN86_COD_PRODUTO = 20
Public Const STRING_IN86_CGC_CPF = 14
Public Const STRING_IN86_INSCR_ESTADUAL = 14
Public Const STRING_IN86_INSCR_MUNICIPAL = 14
Public Const STRING_IN86_RAZAO_SOCIAL = 70
Public Const STRING_IN86_ENDERECO = 60
Public Const STRING_IN86_BAIRRO = 20
Public Const STRING_IN86_MUNICIPIO = 20
Public Const STRING_IN86_UF = 2
Public Const STRING_IN86_PAIS = 20
Public Const STRING_IN86_CEP = 8
Public Const STRING_IN86_CONTA_CTB = 28
Public Const STRING_IN86_TIPOCONTA_CTB = 1
Public Const STRING_IN86_DESCRICAO = 45
Public Const STRING_IN86_CCL = 28
Public Const STRING_IN86_DATA_ATUALIZACAO = 8
Public Const STRING_IN86_NATUREZAOP = 6
Public Const STRING_IN86_CFOP = 4
Public Const STRING_IN86_FILIAL = 4
Public Const STRING_IN86_VALOR = 17
Public Const STRING_IN86_DEB_CRED = 1
Public Const STRING_IN86_ARQUIVAMENTO = 12
Public Const STRING_IN86_LANCAMENTO = 12
Public Const STRING_IN86_HISTORICO = 150
Public Const STRING_IN86_FILIALEMPRESA = 5
Public Const STRING_IN86_ORIGEM = 50
Public Const STRING_IN86_CANCELAMENTO = 1
Public Const STRING_IN86_ESPECIEVOLUME = 10
Public Const STRING_IN86_IDVEICULO = 15
Public Const STRING_IN86_INDICADORMOV = 1
Public Const STRING_IN86_MODALIDADEFRETE = 3
Public Const STRING_IN86_MODELODOC = 2
Public Const STRING_IN86_OBSERVACAO = 45
Public Const STRING_IN86_SERIE = 5
Public Const STRING_IN86_TIPOFATURA = 1
Public Const STRING_IN86_TRANSPORTADOR = 14
Public Const STRING_IN86_VIA_TRANSPORTE = 15
Public Const STRING_IN86_NUMERO_NF = 9
Public Const STRING_IN86_PESO = 17
Public Const STRING_IN86_UM = 3
Public Const STRING_IN86_IPICODIGO = 8
Public Const STRING_IN86_SITUACAO_TRB = 3
Public Const STRING_IN86_HISTORICOMOVEST = 50
Public Const STRING_IN86_NUMERO_DOC = 12
Public Const STRING_IN86_TIPO_DOC = 3
Public Const STRING_IN86_SIGLA_DOC = 5
Public Const STRING_IN86_DESC_TIPOSMOVEST = 100
Public Const STRING_IN86_COD_TIPOSMOVEST = 5
Public Const STRING_IN86_PERCENTUALPERDA = 5
Public Const STRING_IN86_TIPO_OPERACAO = 1
Public Const STRING_IN86_PARCELA = 3

Public Const SITUACAOEST_NOSSO_EM_PODER_CONTRIBUITE = 1
Public Const SITUACAOEST_NOSSO_EM_PODER_TERCEIROS = 2
Public Const SITUACAOEST_TERCEIROS_EM_NOSSO_PODER = 3

Public Const IN86_TIPO_IPI_SAIDA_TRIBUTADO = 1
Public Const IN86_TIPO_IPI_SAIDA_NAOTRIBUTADO = 2
Public Const IN86_TIPO_IPI_SAIDA_OUTROS = 3

Public Const IN86_TIPO_IPI_ENTRADA_RECUPERA = 1
Public Const IN86_TIPO_IPI_ENTRADA_NAOTRIBUTADO = 2
Public Const IN86_TIPO_IPI_ENTRADA_OUTROS = 3

Public Const IN86_TIPO_ICMS_TRIBUTADO = 1
Public Const IN86_TIPO_ICMS_NAOTRIBUTADO = 2
Public Const IN86_TIPO_ICMS_OUTROS = 3

Public Const IN86_INDICADORMOV_ENTRADA As String = "E" 'Incluído por Luiz Nogueira em 28/01/04
Public Const IN86_INDICADORMOV_SAIDA As String = "S" 'Incluído por Luiz Nogueira em 28/01/04

'Constantes que guardam o prefixo dos arquivos que serão gerados pelo IN86
'Ver tabela IN86TipoArquivos
Public Const IN86_ARQ_FORNECEDORHISTORICO As String = "IN86Fornecedores"
Public Const IN86_ARQ_CLIENTEHISTORICO As String = "IN86Clientes"
Public Const IN86_ARQ_TRANSPORTADORAHISTORICO As String = "IN86Transportadoras"
Public Const IN86_ARQ_NATOPHISTORICO As String = "IN86NaturezaOP"
Public Const IN86_ARQ_CCLHISTORICO As String = "IN86Ccl"
Public Const IN86_ARQ_PRODUTOHISTORICO As String = "IN86Produtos"
Public Const IN86_ARQ_PLANOCONTAHISTORICO As String = "IN86PlanoContas"
Public Const IN86_ARQ_LANCAMENTO As String = "IN86LancamentosCTB"
Public Const IN86_ARQ_SALDOSMENSAIS As String = "IN86SldMensais"
Public Const IN86_ARQ_NF_EMPRESAS As String = "IN86NFEmpresa"
Public Const IN86_ARQ_ITENSNF_EMPRESAS As String = "IN86ItensNFEmpresa"
Public Const IN86_ARQ_NF_TERCEIROS As String = "IN86NFTerceiros"
Public Const IN86_ARQ_ITENSNF_TERCEIROS As String = "IN86ItensNFTerceiros"
Public Const IN86_ARQ_MOVESTOQUE As String = "IN86MovEstoque"
Public Const IN86_ARQ_TIPOSMOVESTOQUE As String = "IN86TiposMovEstoque"
Public Const IN86_ARQ_REGINVENTARIO As String = "IN86RegInventario"
Public Const IN86_ARQ_INSUMOS As String = "IN86Insumos"
Public Const IN86_ARQ_CTBCLIENTES As String = "IN86CTBClientes"
Public Const IN86_ARQ_CTBFORN As String = "IN86CTBForn"


Type typeSaldosMensaisIN86

    dtDataSldInicial As Date
    sConta As String
    dValorSaldoInicial As Double
    iSldIniDebCred As Integer
    dValorTotalDeb As Double
    dValorTotalCred As Double
    dValorSaldoFinal As Double
    iSldFimDebCred As Integer
    iFilialEmpresa As Integer
    
End Type

Type typeNFiscaisEmpresaIN86
    sIndicadorMov As String
    iModeloDoc As Integer
    sSerie As String
    lNumeroDoc As Long
    dtDataEmissao As Date
    dValorMercadoria As Double
    dValorTotDesconto As Double
    dValorFrete As Double
    dValorSeguro As Double
    dValorOutrasDesp As Double
    dValorTotalIPI As Double
    dValorTotalICMSSubsTRB As Double
    dValorTotalNF As Double
    sInscEstSubsTRB As String
    iViaTransporte As Integer
    iTransportador As Integer
    iQtdVolumes As Integer
    sEspecieVolume As String
    dPesoBruto As Double
    dPesoLiquido As Double
    iModalidadeFrete As Integer
    sIDVeiculo As String
    sTipoFatura As String
    sObservacao As String
    sIndicadorMovCMP As String
    iEmitente As Integer
    iDestinatario As Integer
    iRemetente As Integer
    iOrigem As Integer
    lFornecedor As Long
    iFilialForn As Integer
    lCliente As Long
    iFilialCliente As Integer
    dtDataSaida As Date
    dtDataEntrada As Date
    iStatus As Integer
    iFilialEmpresa As Integer
    iClasseDocCPR As Integer
    lNumIntDocCPR As Long
    iTipoTipoDocInfo As Integer 'Incluído por Luiz Nogueira em 28/01/04
End Type

Type typeItensNFiscaisEmpresaIN86
    iColunaNoLivro As Integer
    iColunaNoLivroEntrada As Integer
    iColunaNoLivroSaida As Integer
    iOrigem As Integer
    sIndicadorMov As String
    sIndicadorMovCMP As String
    iModeloDoc As Integer
    sSerie As String
    lNumeroNF As Long
    dtDataEmissao As Date
    iItemNF As Integer
    sProduto As String
    sDescProd As String
    sNaturezaOp As String
    sIPICodigo As String
    dQuantidade As Double
    sUM As String
    dPrecoUnitario As Double
    dValorDesconto As Double
    iIPITipo As Integer
    dIPIAliquota As Double
    dIPIBaseCalculo As Double
    dIPIValor As Double
    iOrigemMerc As Integer
    iTipoTribCST As Integer
    iICMSTipo As Integer
    dICMSAliquota As Double
    dICMSBase As Double
    dICMSValor As Double
    dICMSSubstBase As Double
    dICMSSubstValor As Double
    iTipoMovEstoque As Integer
    iFilialEmpresa As Integer
    lFornecedor As Long
    lCliente As Long
    iFilialForn As Integer
    iFilialCli As Integer
    iTipoTipoDocInfo As Integer 'Incluído por Luiz Nogueira em 28/01/04
    sCSTIPI As String
End Type

Type typeIN86MovEstoque
    
    sProduto As String
    iTipoNumIntDocOrigem As Integer
    dtDataMov As Date
    sSiglaUM As String
    sSiglaDoc As String
    dQuantidade As Double
    sEntradaouSaida As String
    sEntradaSaidaCMP As String
    dCustoUnitario As Double
    dValorTotal As Double
    lNumIntDocOrigem As Long
    iFilialEmpresa As Integer
    sSerie As String
    sNumeroDoc As String
    sHistorico As String
    lCodigo As Long
End Type

Type typeIN86TiposMovEstoque
    
    iCodigo As Integer
    sDescricao As String
    sSigla As String
    
End Type

Type typeIN86Insumos

    dtData As Date
    sProdutoPai As String
    sVersao As String
    iNivel As Integer
    iSeq As Integer
    sProduto As String
    iSeqPai As Integer
    iComposicao As Integer
    dQuantidade As Double
    sUnidadeMedInsumo As String
    dPercentualPerda As Double
    sUnidadeMedPai As String

End Type

Type typeIN86CTBForn

    sContaContabil As String
    sContaContabilAux As String
    lFornecedor As Long
    iFilial As Integer
    dtDataOperacao As Date
    sHistorico As String
    dValorOperacao As Double
    sTipoOperacao As String
    sTipoDocumento As String
    lDocumento As Long
    dValorTitulo As Double
    dtDataEmissao As Date
    dtDataVencimento As Date
    sArquivamento As String
    iNumParcela As Integer
    iFilialEmpresa As Integer

End Type

Type typeIN86CTBClientes

    sContaContabil As String
    sContaContabilAux As String
    lCliente As Long
    iFilial As Integer
    dtDataOperacao As Date
    sHistorico As String
    dValorOperacao As Double
    sTipoOperacao As String
    sTipoDocumento As String
    lDocumento As Long
    dValorTitulo As Double
    dtDataEmissao As Date
    dtDataVencimento As Date
    sArquivamento As String
    iNumParcela As Integer
    iFilialEmpresa As Integer

End Type

'type para geração de arquivos 60 e 61
Type typeLojaArqFisICMS
'Modificado por Wagner

    iCodECF As Integer
    iFilialEmpresa As Integer
    dtData As Date
    sNumSerieECF As String
    lCOOIni As Long
    lCOOFim As Long
    lCRZ As Long
    dGrandeTotal As Double
    dVendaBruta As Double
    lCRO As Long
    dTotalizadorCupom As Double
    sSituacaoTrib As String
    dAliquota As Double
    dTotalizador As Double
    
End Type
    
Type typeGuiasICMS
    dtData As Date
    dtDataEntrega As Date
    dValor As Double
    iFilialEmpresa As Integer
    sNumero As String
    sOrgaoArrecadador As String
    sLocalEntrega As String
    dtApuracaoDe As Date
    dtApuracaoAte As Date
    sCodReceita As String
    sCodObrigRecolher As String
    dtDataVencimento As Date
    sUF As String
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeSpedDocFiscais
    lNumIntSped As Long
    iTipoDoc As Integer
    lNumIntDoc As Long
    sBloco As String
    sRegistro As String
    iFilialEmpresa As Integer
    lCliente As Long
    iFilialCli As Integer
    lFornecedor As Long
    iFilialForn As Integer
    dtDataEmissao As Date
    dtDataES As Date
    dValorTotal As Double
    sOperacao As String
    sEmitente As String
    sModelo As String
    iSituacao As Integer
    sSerie As String
    lNumDocumento As Long
    sChaveEletronica As String
    sFrete As String
    sPagamento As String
    dValorDesconto As Double
    dValorFrete As Double
    dValorSeguro As Double
    dValorDespesas As Double
    dValorServico As Double
    dValorNaoTributado As Double
    dValorBaseICMS As Double
    dValorICMS As Double
    dValorBaseICMSST As Double
    dValorICMSST As Double
    dValorBaseIPI As Double
    dValorIPI As Double
    dValorBasePIS As Double
    dValorPIS As Double
    dValorBaseCofins As Double
    dValorCofins As Double
    dValorPisRetido As Double
    dValorPisST As Double
    dValorCofinsRetido As Double
    dValorCofinsST As Double
    dValorISS As Double
    sCgc As String
    lCodMensagem As Long
    dValorMercadoria As Double
    sIndNatFrtPis As String
    sIndNatFrtCofins As String
    dValorRecebido As Double
    dValorNoCR As Double
    dValorRecContrPrev As Double
    sCodMunicServ As String
    dValorBaseISS As Double
    dValorBaseISSRet As Double
    dValorISSRet As Double
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeSpedDocFiscaisItens
    lNumIntSped As Long
    iTipoDoc As Integer
    lNumIntDoc As Long
    iItem As Integer
    sRegistro As String
    sProduto As String
    sDescricao As String
    dQuantidade As Double
    sUM As String
    sMovFisica As String
    sICMSCST As String
    sCFOP As String
    dICMSAliquota As Double
    dICMSSTAliquota As Double
    sIPICST As String
    sIPICodEnq As String
    dIPIAliquota As Double
    sPISCST As String
    dPISAliquota As Double
    dPISBCQtd As Double
    dPISAliquotaQtd As Double
    sCofinsCST As String
    dCOFINSAliquota As Double
    dCofinsBCQtd As Double
    dCofinsAliquotaQtd As Double
    dValorDesconto As Double
    dValorFrete As Double
    dValorSeguro As Double
    dValorDespesas As Double
    dValorServico As Double
    dValorNaoTributado As Double
    dValorBaseICMS As Double
    dValorICMS As Double
    dValorBaseICMSST As Double
    dValorICMSST As Double
    dValorBaseIPI As Double
    dValorIPI As Double
    dValorBasePIS As Double
    dValorPIS As Double
    dValorBaseCofins As Double
    dValorCofins As Double
    dValorPisRetido As Double
    dValorPisST As Double
    dValorCofinsRetido As Double
    dValorCofinsST As Double
    dValorISS As Double
    dPrecoUnitario As Double
    dPrecoTotal As Double
    sNatBCCred As String
    dValorRecebido As Double
    dValorNoCR As Double
    dValorRecContrPrev As Double
    dAliquotaContrPrev As Double
    sCodAtividadeTab511 As String
    sNCM As String
    sCodISSServ As String
    dValorBaseISS As Double
    dAliquotaISS As Double
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeRegApuracaoPISCofins
    lNumIntDoc As Long
    iFilialEmpresa As Integer
    iTipo As Integer
    iAno As Integer
    iMes As Integer
    iOrigCred As Integer
    sCNPJCedCred As String
    sCodCred As String
    dVlCredApu As Double
    dVlCredExtApu As Double
    dVlTotCredApu As Double
    dVlCredDescPAAnt As Double
    dVlCredPerPAAnt As Double
    dVlCredDCompPAAnt As Double
    dSdCredDispEFD As Double
    dVlCredDescEFD As Double
    dVlCredPerEFD As Double
    dVlCredDCompEFD As Double
    dVlCredTrans As Double
    dVlCredOut As Double
    dSdCredFim As Double
End Type

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AdmSelect2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''Option Explicit
''
''Function Produto_Le(objProduto As ClassProduto) As Long
'''Lê os dados do produto, cujo codigo foi passado como parâmetro
''
''Dim lErro As Long
''Dim lComando As Long
''Dim tProduto As typeProduto
''
''On Error GoTo Erro_Produto_Le
''
''    'Abre o comando
''    lComando = Comando_Abrir()
''    If lComando = 0 Then Error 28027
''
''    'Inicializa as variáveis do tipo String
''    With tProduto
''
''        .sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
''        .sSiglaUMEstoque = String(STRING_PRODUTO_SIGLAUMESTOQUE, 0)
''        .sSiglaUMCompra = String(STRING_PRODUTO_SIGLAUMCOMPRA, 0)
''        .sSiglaUMVenda = String(STRING_PRODUTO_SIGLAUMVENDA, 0)
''        .sSubstituto1 = String(STRING_PRODUTO_SUBSTITUTO1, 0)
''        .sSubstituto2 = String(STRING_PRODUTO_SUBSTITUTO2, 0)
''        .sNomeReduzido = String(STRING_PRODUTO_NOME_REDUZIDO, 0)
''        .sModelo = String(STRING_PRODUTO_MODELO, 0)
''        .sCodigoBarras = String(STRING_PRODUTO_CODIGO_BARRAS, 0)
''        .sCor = String(STRING_PRODUTO_COR, 0)
''        .sIPICodDIPI = String(STRING_PRODUTO_IPI_COD_DIPI, 0)
''        .sIPICodigo = String(STRING_PRODUTO_IPI_CODIGO, 0)
''        .sObsFisica = String(STRING_PRODUTO_OBS_FISICA, 0)
''        .sContaContabil = String(STRING_CONTA, 0)
''        .sContaContabilProducao = String(STRING_CONTA, 0)
''
''        'Lê o Produto
''        lErro = Comando_Executar(lComando, "SELECT  Descricao, Gerencial, Ativo, Faturamento, ClasseUM, SiglaUMEstoque, SiglaUMCompra, SiglaUMVenda, Substituto1, Substituto2, Nivel, Tipo, NomeReduzido, Modelo, PrazoValidade, CodigoBarras, EtiquetasCodBarras, PesoLiq, PesoBruto, Comprimento, Espessura, Largura, Cor, ObsFisica, Compras, PCP, KitBasico, KitInt, IPIAliquota, IPICodigo, IPICodDIPI, ControleEstoque, Apropriacao, ContaContabil, ContaContabilProducao, Residuo, Natureza ,CustoReposicao, OrigemMercadoria, TabelaPreco, PercentMaisQuantCotacaoAnterior, PercentMenosQuantCotacaoAnterior, ConsideraQuantCotacaoAnterior, TemFaixaReceb, PercentMaisReceb, PercentMenosReceb, RecebForaFaixa FROM Produtos WHERE Codigo = ?" _
''        , .sDescricao, .iGerencial, .iAtivo, .iFaturamento, .iClasseUM, .sSiglaUMEstoque, .sSiglaUMCompra, .sSiglaUMVenda, .sSubstituto1, .sSubstituto2, .iNivel, .iTipo, .sNomeReduzido, .sModelo, .iPrazoValidade, .sCodigoBarras, .iEtiquetasCodBarras, .dPesoLiq, .dPesoBruto, .dComprimento, .dEspessura, .dLargura, .sCor, .sObsFisica, .iCompras, .iPCP, .iKitBasico, .iKitInt, .dIPIAliquota, .sIPICodigo, .sIPICodDIPI, .iControleEstoque, .iApropriacaoCusto, .sContaContabil, .sContaContabilProducao, .dResiduo, .iNatureza, .dCustoReposicao, .iOrigemMercadoria, .iTabelaPreco, .dPercentMaisQuantCotacaoAnterior, .dPercentMenosQuantCotacaoAnterior, .iConsideraQuantCotacaoAnterior, .iTemFaixaReceb, .dPercentMaisReceb, .dPercentMenosReceb, .iRecebForaFaixa, objProduto.sCodigo)
''        If lErro <> AD_SQL_SUCESSO Then Error 28028
''
''        lErro = Comando_BuscarPrimeiro(lComando)
''        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 28029
''
''        'Se não encontrou o Produto --> Erro
''        If lErro = AD_SQL_SEM_DADOS Then Error 28030
''
''        'Coloca os dados lidos no objProduto
''        objProduto.sDescricao = .sDescricao
''        objProduto.iGerencial = .iGerencial
''        objProduto.iAtivo = .iAtivo
''        objProduto.iFaturamento = .iFaturamento
''        objProduto.iClasseUM = .iClasseUM
''        objProduto.sSiglaUMEstoque = .sSiglaUMEstoque
''        objProduto.sSiglaUMCompra = .sSiglaUMCompra
''        objProduto.sSiglaUMVenda = .sSiglaUMVenda
''        objProduto.sSubstituto1 = .sSubstituto1
''        objProduto.sSubstituto2 = .sSubstituto2
''        objProduto.iTipo = .iTipo
''        objProduto.sNomeReduzido = .sNomeReduzido
''        objProduto.sModelo = .sModelo
''        objProduto.iNivel = .iNivel
''        objProduto.iPrazoValidade = .iPrazoValidade
''        objProduto.sCodigoBarras = .sCodigoBarras
''        objProduto.iEtiquetasCodBarras = .iEtiquetasCodBarras
''        objProduto.dPesoLiq = .dPesoLiq
''        objProduto.dPesoBruto = .dPesoBruto
''        objProduto.dComprimento = .dComprimento
''        objProduto.dEspessura = .dEspessura
''        objProduto.dLargura = .dLargura
''        objProduto.sCor = .sCor
''        objProduto.sObsFisica = .sObsFisica
''        objProduto.iCompras = .iCompras
''        objProduto.iPCP = .iPCP
''        objProduto.iKitBasico = .iKitBasico
''        objProduto.iKitInt = .iKitInt
''        objProduto.dIPIAliquota = .dIPIAliquota
''        objProduto.sIPICodigo = .sIPICodigo
''        objProduto.sIPICodDIPI = .sIPICodDIPI
''        objProduto.iControleEstoque = .iControleEstoque
''        objProduto.iApropriacaoCusto = .iApropriacaoCusto
''        objProduto.sContaContabil = .sContaContabil
''        objProduto.sContaContabilProducao = .sContaContabilProducao
''        objProduto.dResiduo = .dResiduo
''        objProduto.iNatureza = .iNatureza
''        objProduto.dCustoReposicao = .dCustoReposicao
''        objProduto.iOrigemMercadoria = .iOrigemMercadoria
''        objProduto.iTabelaPreco = .iTabelaPreco
''        objProduto.dPercentMaisQuantCotacaoAnterior = .dPercentMaisQuantCotacaoAnterior
''        objProduto.dPercentMenosQuantCotacaoAnterior = .dPercentMenosQuantCotacaoAnterior
''        objProduto.iConsideraQuantCotacaoAnterior = .iConsideraQuantCotacaoAnterior
''        objProduto.iTemFaixaReceb = .iTemFaixaReceb
''        objProduto.dPercentMaisReceb = .dPercentMaisReceb
''        objProduto.dPercentMenosReceb = .dPercentMenosReceb
''        objProduto.iRecebForaFaixa = .iRecebForaFaixa
''
''    End With
''
''    'Fecha o comando
''    Call Comando_Fechar(lComando)
''
''    Produto_Le = SUCESSO
''
''    Exit Function
''
''Erro_Produto_Le:
''
''    Produto_Le = Err
''
''    Select Case Err
''
''        Case 28027
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
''
''        Case 28028, 28029
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS", Err, objProduto.sCodigo)
''
''        Case 28030
''            'Produto não encontrado. Será tratado na rotina chamadora.
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142609)
''
''    End Select
''
''    Call Comando_Fechar(lComando)
''
''    Exit Function
''
''End Function
''
''
''
''

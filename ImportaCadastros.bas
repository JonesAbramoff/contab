Attribute VB_Name = "ImportaCadastros"
Option Explicit

Private Declare Function Conexao_AbrirExt Lib "ADSQLMN.DLL" Alias "AD_Conexao_Abrir" (ByVal driver_sql As Integer, ByVal lpParamIn As String, ByVal ParamLenIn As Integer, ByVal lpParamOut As String, lpParamLenOut As Integer) As Long
Private Declare Function Conexao_FecharExt Lib "ADSQLMN.DLL" Alias "AD_Conexao_Fechar" (ByVal lConexao As Long) As Long

'Tipos de Frete
Private Const TIPO_CIF = 0
Private Const TIPO_FOB = 1

Type typeImportCli
    lCodCliente As Long
    sRazaoSocial As String
    sNomeReduzido As String
    iTipo As Integer
    sObservacao As String
    dLimiteCredito As Double
    iCondicaoPagto As Integer
    dDesconto As Double
    iCodPadraoCobranca As Integer
    iCodMensagem As Integer
    iTabelaPreco As Integer
    lNumPagamentos As Long
    iProxCodFilial As Integer
    iCodFilial As Integer
    sFilialNome As String
    sFilialCGC As String
    sFilialInscEstadual As String
    sFilialInscMunicipal As String
    iFilialCodTransportadora As Integer
    sFilialObservacao1 As String
    sFilialContaContabil As String
    iFilialVendedor As Integer
    dFilialComissaoVendas As Double
    iFilialRegiao As Integer
    iFilialFreqVisitas As Integer
    dtFilialDataUltVisita As Date
    iFilialCodCobrador As Integer
    iFilialICMSBaseCalculoIPI As Integer
    lFilialRevendedor As Long
    sFilialTipoFrete As String
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
    sEndereco1 As String
    sBairro1 As String
    sCidade1 As String
    sSiglaEstado1 As String
    iCodigoPais1 As Integer
    sCEP1 As String
    sTelefone11 As String
    sTelefone21 As String
    sEmail1 As String
    sFax1 As String
    sContato1 As String
    sEndereco2 As String
    sBairro2 As String
    sCidade2 As String
    sSiglaEstado2 As String
    iCodigoPais2 As Integer
    sCEP2 As String
    sTelefone12 As String
    sTelefone22 As String
    sEmail2 As String
    sFax2 As String
    sContato2 As String
End Type

Type typeImportForn
    lCodigo As Long
    sRazaoSocial As String
    sNomeReduzido As String
    iTipo As Integer
    sObservacao As String
    iCondicaoPagto As Integer
    dDesconto As Double
    iProxCodFilial As Integer
    iFilialCod As Integer
    sFilialNome As String
    sFilialCGC As String
    sFilialInscEstadual As String
    sFilialInscMunicipal As String
    sFilialContaContabil As String
    iFilialBanco As Integer
    sFilialAgencia As String
    sFilialContaCorrente As String
    sFilialObservacao1 As String
    iFilialTipoFrete As Integer
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

Type typeImportProd
    sCodigo As String
    iTipo As Integer
    sDescricao As String
    sNomeReduzido As String
    sModelo As String
    iGerencial As Integer
    iNivel As Integer
    sSubstituto1 As String
    sSubstituto2 As String
    iPrazoValidade As Integer
    sCodigoBarras As String
    iEtiquetasCodBarras As Integer
    dPesoLiq As Double
    dPesoBruto As Double
    dComprimento As Double
    dEspessura As Double
    dLargura As Double
    sCor As String
    sObsFisica As String
    iClasseUM As Integer
    sSiglaUMEstoque As String
    sSiglaUMCompra As String
    sSiglaUMVenda As String
    iAtivo As Integer
    iFaturamento As Integer
    iCompras As Integer
    iPCP As Integer
    iKitBasico As Integer
    iKitInt As Integer
    dIPIAliquota As Double
    sIPICodigo As String
    sIPICodDIPI As String
    iControleEstoque As Integer
    iICMSAgregaCusto As Integer
    iIPIAgregaCusto As Integer
    iFreteAgregaCusto As Integer
    iApropriacaoCusto As Integer
    sContaContabil As String
    sContaContabilProducao As String
    dResiduo As Double
    iNatureza As Integer
    dCustoReposicao As Double
    iOrigemMercadoria As Integer
    iTabelaPreco As Integer
    iTemFaixaReceb As Integer
    dPercentMaisReceb As Double
    dPercentMenosReceb As Double
    iRecebForaFaixa As Integer
    iTempoProducao As Integer
    iRastro As Integer
    lHorasMaquina As Long
    dPesoEspecifico As Double
    iCreditoIPI As Integer
    iCreditoICMS As Integer
    iLinha As Integer
    sDimEmbalagem As String
End Type

Function Importa_Produtos() As Long
'Importa os dados referentes aos Produtos (Tabelas: ImportProd,ImportProdAux,ImportProdDesc)

Dim lErro As Long, iPosPonto As Integer, iCont As Integer
Dim tImportProd As typeImportProd
Dim objProduto As ClassProduto, objProdutoPai As ClassProduto
Dim lComando As Long, lComando2 As Long, lComando3 As Long
Dim lTransacao As Long
Dim lTamanho As Long
Dim iCreditoICMS As Integer
Dim iCreditoIPI As Integer
Dim sCodProduto As String, sCodProdutoAux As String
Dim colTabelaPrecoItem As New Collection
Dim sProdutoPai As String, objProdutoCategoria As ClassProdutoCategoria
Dim sSegmento1 As String, sSegmento2 As String

On Error GoTo Erro_Importa_Produtos
    
    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 76348
    
    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then gError 76348
    
    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then gError 76348
    
    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 76349
    
    tImportProd.sCodigo = String(STRING_PRODUTO, 0)
    tImportProd.sCodigoBarras = String(STRING_PRODUTO_CODIGO_BARRAS, 0)
    tImportProd.sContaContabil = String(STRING_CONTA, 0)
    tImportProd.sContaContabilProducao = String(STRING_CONTA, 0)
    tImportProd.sCor = String(STRING_PRODUTO_COR, 0)
    tImportProd.sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
    tImportProd.sIPICodDIPI = String(STRING_PRODUTO_IPI_COD_DIPI, 0)
    tImportProd.sIPICodigo = String(STRING_PRODUTO_IPI_CODIGO, 0)
    tImportProd.sModelo = String(STRING_PRODUTO_MODELO, 0)
    tImportProd.sNomeReduzido = String(STRING_PRODUTO_NOME_REDUZIDO, 0)
    tImportProd.sObsFisica = String(STRING_PRODUTO_OBS_FISICA, 0)
    tImportProd.sSiglaUMCompra = String(STRING_PRODUTO_SIGLAUMCOMPRA, 0)
    tImportProd.sSiglaUMEstoque = String(STRING_PRODUTO_SIGLAUMESTOQUE, 0)
    tImportProd.sSiglaUMVenda = String(STRING_PRODUTO_SIGLAUMVENDA, 0)
    tImportProd.sSubstituto1 = String(STRING_PRODUTO_SUBSTITUTO1, 0)
    tImportProd.sSubstituto2 = String(STRING_PRODUTO_SUBSTITUTO2, 0)
    tImportProd.sDimEmbalagem = String(255, 0)
    
    'Lê os registros da tabela ImportProd
    With tImportProd
    lErro = Comando_Executar(lComando, "SELECT Codigo,Tipo,Descricao,NomeReduzido,Modelo,Gerencial,Nivel,Substituto1,Substituto2,PrazoValidade,CodigoBarras,EtiquetasCodBarras,PesoLiq,PesoBruto," _
        & "Comprimento,Espessura,Largura,Cor,ObsFisica,ClasseUM,SiglaUMEstoque,SiglaUMCompra,SiglaUMVenda,Ativo,Faturamento,Compras,PCP," _
        & "KitBasico,KitInt,IPIAliquota,IPICodigo,IPICodDIPI,ControleEstoque,ICMSAgregaCusto,IPIAgregaCusto,FreteAgregaCusto,Apropriacao,ContaContabil,ContaContabilProducao,TemFaixaReceb,PercentMaisReceb,PercentMenosReceb,RecebForaFaixa,CreditoICMS,CreditoIPI,Residuo,Natureza," _
        & "CustoReposicao,OrigemMercadoria,TabelaPreco,TempoProducao,Rastro,HorasMaquina,PesoEspecifico FROM ImportProd ORDER BY Codigo", .sCodigo, .iTipo, .sDescricao, .sNomeReduzido, .sModelo, .iGerencial, .iNivel, .sSubstituto1, .sSubstituto2, .iPrazoValidade, .sCodigoBarras, .iEtiquetasCodBarras, .dPesoLiq, .dPesoBruto, .dComprimento, .dEspessura, .dLargura, .sCor, _
        .sObsFisica, .iClasseUM, .sSiglaUMEstoque, .sSiglaUMCompra, .sSiglaUMVenda, .iAtivo, .iFaturamento, .iCompras, .iPCP, .iKitBasico, .iKitInt, .dIPIAliquota, .sIPICodigo, .sIPICodDIPI, .iControleEstoque, .iICMSAgregaCusto, .iIPIAgregaCusto, .iFreteAgregaCusto, .iApropriacaoCusto, .sContaContabil, .sContaContabilProducao, _
        .iTemFaixaReceb, .dPercentMaisReceb, .dPercentMenosReceb, .iRecebForaFaixa, .iCreditoICMS, .iCreditoIPI, .dResiduo, .iNatureza, .dCustoReposicao, .iOrigemMercadoria, .iTabelaPreco, .iTempoProducao, .iRastro, .lHorasMaquina, .dPesoEspecifico)
    End With
    If lErro <> AD_SQL_SUCESSO Then gError 76350
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 76351
    
    Do While lErro = AD_SQL_SUCESSO
    
        sCodProduto = Replace(tImportProd.sCodigo, "'", "")
        
        Set objProduto = New ClassProduto

        objProduto.sNomeReduzido = Trim(sCodProduto)
        
        iPosPonto = InStr(sCodProduto, ".")
        If iPosPonto = 0 Then
            sSegmento1 = "SRV"
            sSegmento2 = sCodProduto
        Else
            sSegmento1 = left(sCodProduto, iPosPonto - 1)
            sSegmento2 = Mid(sCodProduto, iPosPonto + 1)
        End If
        
        If Len(sSegmento1) < 3 Then
            sSegmento1 = sSegmento1 & String(3 - Len(sSegmento1), " ")
        End If
        If Len(sSegmento2) < 15 Then
            sSegmento2 = sSegmento2 & String(15 - Len(sSegmento2), " ")
        End If
        
        sCodProduto = UCase(sSegmento1 & sSegmento2)
        sProdutoPai = UCase(sSegmento1 & String(18 - Len(sSegmento1), " "))
        
        objProduto.sCodigo = sCodProduto
        'Verifica se o Produto já está cadastrado
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 76352

        'Se não existe o Produto na tabela Produtos e o codigo eh valido
        If lErro = 28030 Then
        
            If sProdutoPai <> "" Then

                Set objProdutoPai = New ClassProduto

                objProdutoPai.sCodigo = sProdutoPai

                objProdutoPai.sDescricao = tImportProd.sDescricao
                objProdutoPai.sNomeReduzido = Trim(sProdutoPai) & "."

                'Se não estiver, grava o produto "pai"
                lErro = Produto_Define_ProdutoPai(objProdutoPai, tImportProd)
                If lErro <> SUCESSO Then gError 76411

            End If
            
            'Preenche objProduto a partir de tImportProd
            lErro = Produto_PreencheObjetos(tImportProd, objProduto)
            If lErro <> SUCESSO Then gError 76353
                  
            'Grava o Produto
            lErro = CF("Produto_Grava_Trans", objProduto, colTabelaPrecoItem)
            If lErro <> SUCESSO Then
                gError 76354
            End If

            lErro = Comando_Executar(lComando3, "UPDATE importprod SET codigo_corporator = ? WHERE Codigo = ?", objProduto.sCodigo, tImportProd.sCodigo)
            If lErro <> AD_SQL_SUCESSO Then gError 76354
            
            iCont = iCont + 1
            
'            If iCont = 500 Then Exit Do
        
        End If
        
        'Busca o proximo registro de ImportProd
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 76355
        
    Loop
'    MsgBox (CStr(iCont))
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 76356
        
    'Fecha o comando
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    
    Importa_Produtos = SUCESSO
    
    Exit Function
    
Erro_Importa_Produtos:

    Importa_Produtos = gErr
    
    Select Case gErr
    
        Case 76348
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 76349
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 76350, 76351, 76355
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_IMPORTPROD", gErr)
        
        Case 76352, 76353, 76354, 76373, 76400, 76411
            'Erros tratados nas rotinas chamadas
            
        Case 76356
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT_TRANSACAO", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161813)
            
    End Select
    
    Call Transacao_Rollback
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    
    Exit Function
    
End Function

Function Produto_PreencheObjetos(tImportProd As typeImportProd, ByVal objProduto As ClassProduto) As Long
'Preenche objProduto a partir dos dados existentes em tImportProd

Dim lErro As Long
Dim objProdAux As New ClassProduto

On Error GoTo Erro_Produto_PreencheObjetos
    
    objProduto.sDescricao = Trim(tImportProd.sDescricao)
    If objProduto.sDescricao = "" Then objProduto.sDescricao = "SEM DESCRIÇÃO"
    
    'objProduto.sNomeReduzido = Trim(tImportProd.sNomeReduzido)
    'If objProduto.sNomeReduzido = "" Then objProduto.sNomeReduzido = left(objProduto.sDescricao, STRING_PRODUTO_NOME_REDUZIDO)
'
'    objProdAux.sNomeReduzido = objProduto.sNomeReduzido
'
'    lErro = CF("Produto_Le_NomeReduzido", objProdAux)
'    If lErro <> SUCESSO And lErro <> 26927 Then gError 89869
'
'    'se o produto estiver cadastrado
'    If lErro = SUCESSO Then
'        If objProdAux.sCodigo <> objProduto.sCodigo Then
'            objProduto.sNomeReduzido = objProduto.sCodigo
'        End If
'    End If
    
    objProduto.dComprimento = tImportProd.dComprimento
    objProduto.dCustoReposicao = tImportProd.dCustoReposicao
    objProduto.dEspessura = tImportProd.dEspessura
    objProduto.dIPIAliquota = tImportProd.dIPIAliquota
    
    If objProduto.dIPIAliquota > 1 Then objProduto.dIPIAliquota = objProduto.dIPIAliquota / 100
    
    objProduto.dLargura = tImportProd.dLargura
    objProduto.dPercentMenosReceb = tImportProd.dPercentMenosReceb
    objProduto.dPesoBruto = tImportProd.dPesoBruto
    objProduto.dPesoEspecifico = tImportProd.dPesoEspecifico
    objProduto.dPesoLiq = tImportProd.dPesoLiq
    objProduto.dResiduo = tImportProd.dResiduo
    objProduto.iAtivo = PRODUTO_ATIVO 'tImportProd.iAtivo
    
    objProduto.iControleEstoque = PRODUTO_CONTROLE_ESTOQUE 'tImportProd.iControleEstoque
    objProduto.iCreditoICMS = tImportProd.iCreditoICMS
    objProduto.iCreditoIPI = tImportProd.iCreditoIPI
    objProduto.iEtiquetasCodBarras = tImportProd.iEtiquetasCodBarras
    objProduto.iFreteAgregaCusto = tImportProd.iFreteAgregaCusto
    objProduto.iGerencial = 0 'tImportProd.iGerencial
    objProduto.iICMSAgregaCusto = tImportProd.iICMSAgregaCusto
    objProduto.iIPIAgregaCusto = tImportProd.iIPIAgregaCusto
    objProduto.iKitBasico = 1 'tImportProd.iKitBasico
    objProduto.iKitInt = 1 'tImportProd.iKitInt
    objProduto.iNatureza = tImportProd.iNatureza
    
    If left(objProduto.sCodigo, 3) = "SRV" Then
        objProduto.iNatureza = NATUREZA_PROD_SERVICO
    End If
    
    If objProduto.iNatureza = 0 Then objProduto.iNatureza = NATUREZA_PROD_PRODUTO_REVENDA
    If objProduto.iNatureza = NATUREZA_PROD_SERVICO Then
        objProduto.iControleEstoque = PRODUTO_CONTROLE_SEM_ESTOQUE
    End If
    
    objProduto.iFaturamento = 1 'tImportProd.iFaturamento
    
    'Verifica se o produto pode ser comprado
    If objProduto.iNatureza = NATUREZA_PROD_SERVICO Or objProduto.iNatureza = NATUREZA_PROD_SUBPRODUTO Or objProduto.iNatureza = NATUREZA_PROD_PRODUTO_INTERMEDIARIO Or objProduto.iNatureza = NATUREZA_PROD_PRODUTO_ACABADO Then
        objProduto.iCompras = PRODUTO_PRODUZIVEL
    Else
        objProduto.iCompras = PRODUTO_COMPRAVEL
    End If
    
    'Verifica a Apropriacao de Custo do Produto
    If objProduto.iCompras = 0 Then
        objProduto.iApropriacaoCusto = APROPR_CUSTO_REAL
    Else
        objProduto.iApropriacaoCusto = APROPR_CUSTO_MEDIO
    End If
    
'    objProduto.iNivel = IIf(InStr(objProduto.sCodigo, ".") <> 0, 2, 1)
    objProduto.iNivel = 2
    
    objProduto.iPCP = tImportProd.iPCP
    
    objProduto.iPrazoValidade = tImportProd.iPrazoValidade
    objProduto.iRastro = tImportProd.iRastro
    objProduto.iTabelaPreco = tImportProd.iTabelaPreco
    objProduto.iTempoProducao = tImportProd.iTempoProducao
    objProduto.iTipo = tImportProd.iTipo
    objProduto.lHorasMaquina = tImportProd.lHorasMaquina
    
    If objProduto.iTipo = 0 Then objProduto.iTipo = 1
    
    objProduto.iKitVendaComp = 0
    
    objProduto.sCodigoBarras = Trim(tImportProd.sCodigoBarras)
    If objProduto.sCodigoBarras <> "" Then
        Call objProduto.colCodBarras.Add(objProduto.sCodigoBarras)
    End If
    
    objProduto.sContaContabil = tImportProd.sContaContabil
    objProduto.sContaContabilProducao = tImportProd.sContaContabilProducao
    objProduto.sCor = Trim(tImportProd.sCor)
    objProduto.sIPICodDIPI = Trim(tImportProd.sIPICodDIPI)
    objProduto.sIPICodigo = IIf(Len(Trim(tImportProd.sIPICodigo)) <> 0, tImportProd.sIPICodigo, "")
    objProduto.sModelo = Trim(tImportProd.sModelo)
    
    'Define a Classe de UM do Produto
    lErro = Produto_Define_ClasseUM(tImportProd)
    If lErro <> SUCESSO Then gError 76435
    
    objProduto.iClasseUM = tImportProd.iClasseUM
    objProduto.sSiglaUMCompra = UCase(Trim(tImportProd.sSiglaUMEstoque))
    objProduto.sSiglaUMEstoque = UCase(Trim(tImportProd.sSiglaUMEstoque))
    'UMVenda ficará igual a UMEstoque
    objProduto.sSiglaUMVenda = UCase(Trim(tImportProd.sSiglaUMEstoque))
    objProduto.sSiglaUMTrib = UCase(Trim(tImportProd.sSiglaUMEstoque))
    
    objProduto.sSubstituto1 = tImportProd.sSubstituto1
    objProduto.sSubstituto2 = tImportProd.sSubstituto2
    
    objProduto.iOrigemMercadoria = tImportProd.iOrigemMercadoria
    
    'Informações referentes a Compras
    objProduto.iConsideraQuantCotAnt = 1
    objProduto.dPercentMenosQuantCotAnt = 0
    objProduto.dPercentMaisQuantCotAnt = 0
    objProduto.iTemFaixaReceb = 0
    objProduto.dPercentMaisReceb = 0
    objProduto.dPercentMenosReceb = 0
    objProduto.iRecebForaFaixa = 1
    
    objProduto.sObsFisica = Trim(tImportProd.sObsFisica)
        
    Produto_PreencheObjetos = SUCESSO
    
    Exit Function
    
Erro_Produto_PreencheObjetos:

    Produto_PreencheObjetos = gErr
    
    Select Case gErr
    
        Case 76404, 76435
            'Erros tratados nas rotinas chamadas
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161814)
            
    End Select
    
    Exit Function
    
End Function

Function Produto_Define_ClasseUM(tImportProd As typeImportProd) As Long
'Define ClasseUM e SiglaUM do Produto, a partir dos dados lidos em ImportProd

Dim lErro As Long, iIndice As Integer
Dim alComando(1 To 4) As Long
Dim objClasseUM As New ClassClasseUM
Dim objUM As New ClassUnidadeDeMedida
Dim dQuantidade As Double
Dim iClasseUM As Integer
Dim iProx As Integer, sDescr As String
Dim colSiglas As New Collection

On Error GoTo Erro_Produto_Define_ClasseUM

    'Abrir comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir
        If alComando(iIndice) = 0 Then gError 125827
    Next

    'Realiza a leitura
    lErro = Comando_Executar(alComando(1), "SELECT Classe FROM ClasseUM WHERE Sigla = ? AND Classe BETWEEN 100 AND 800 ORDER BY Classe DESC ", _
                                        iClasseUM, tImportProd.sSiglaUMEstoque)
    If lErro <> AD_SQL_SUCESSO Then gError 125828
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 125829
    
    If lErro <> AD_SQL_SUCESSO Then 'gError 76471
    
        lErro = Comando_Executar(alComando(2), "SELECT MAX(Classe) FROM ClasseUM WHERE Classe BETWEEN 100 AND 800 ", iProx)
        If lErro <> AD_SQL_SUCESSO Then gError 125828
        
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 125829
        
        If iProx < 100 Then iProx = 100
        
        sDescr = String(STRING_MAXIMO, 0)
        
        lErro = Comando_Executar(alComando(3), "SELECT Descricao FROM ImportUM WHERE Sigla = ? ", sDescr, tImportProd.sSiglaUMEstoque)
        If lErro <> AD_SQL_SUCESSO Then gError 125828
        
        lErro = Comando_BuscarPrimeiro(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 125829
        
        If lErro <> AD_SQL_SUCESSO Then
            sDescr = Trim(tImportProd.sSiglaUMEstoque)
        End If
        
        iProx = iProx + 1
        
        objClasseUM.iClasse = iProx
        objClasseUM.sDescricao = sDescr
        objClasseUM.sSiglaUMBase = Trim(UCase(tImportProd.sSiglaUMEstoque))
        
        objUM.iClasse = objClasseUM.iClasse
        objUM.sSigla = objClasseUM.sSiglaUMBase
        objUM.sNome = sDescr
        objUM.dQuantidade = 1
        objUM.sSiglaUMBase = objClasseUM.sSiglaUMBase

        iClasseUM = objClasseUM.iClasse
        
        colSiglas.Add objUM
        
        lErro = CF("ClasseUM_Grava_EmTrans", objClasseUM, colSiglas)
        If lErro <> SUCESSO Then gError 125828
    
    End If

    tImportProd.iClasseUM = iClasseUM
    
    If tImportProd.iClasseUM = 0 Then gError 76472
    
    'Libera comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Produto_Define_ClasseUM = SUCESSO
    
    Exit Function
    
Erro_Produto_Define_ClasseUM:

    Produto_Define_ClasseUM = gErr
    
    Select Case gErr
    
        Case 125828, 125829
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLASSEUM", gErr)
        
        Case 76471
            Call Rotina_Erro(vbOKOnly, "ERRO_UM_NAO_CADASTRADA1", gErr, tImportProd.sSiglaUMEstoque)
        
        Case 76472
            Call Rotina_Erro(vbOKOnly, "ERRO_CLASSEUM_NAO_INFORMADA", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161815)
            
    End Select

    'Libera comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function
        
End Function

Private Function Produto_Define_ProdutoPai(ByVal objProdutoPai As ClassProduto, tImportProd As typeImportProd) As Long
'verifica se já está cadastrado em Produtos. Se não estiver, grava o
'produto "pai" na tabela de Produtos
            
Dim lErro As Long
Dim lComando As Long
Dim sCodProduto2 As String
Dim sCodigo As String
Dim sDescricao As String
Dim sProdutoPai As String
Dim colTabelaPrecoItem As New Collection

On Error GoTo Erro_Produto_Define_ProdutoPai

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 76399

    sProdutoPai = objProdutoPai.sCodigo
    
    'Lê o Produto "Pai"
    lErro = CF("Produto_Le", objProdutoPai)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 76398
    
    'Se o produto "pai" não estiver cadastrado
    If lErro = 28030 Then
    
        objProdutoPai.sCodigo = sProdutoPai
        
        'Preenche Produto "Pai" com os mesmos dados do Produto "Filho" que estão em tImportProd
        lErro = Produto_PreencheObjetos(tImportProd, objProdutoPai)
        If lErro <> SUCESSO Then gError 76410
        
        'Altera os dados específicos do produto "pai", que não são iguais ao "filho"
        '??? HICARE objProdutoPai.sDescricao = Trim(sDescricao)
        objProdutoPai.iGerencial = 1
        objProdutoPai.iNivel = 1 '??? HICARE
        
        'Grava o "produto pai" no BD
        lErro = CF("Produto_Grava_Trans", objProdutoPai, colTabelaPrecoItem)
        If lErro <> SUCESSO Then
            gError 76403
        End If
    
    End If

    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Produto_Define_ProdutoPai = SUCESSO
    
    Exit Function
    
Erro_Produto_Define_ProdutoPai:

    Produto_Define_ProdutoPai = gErr
    
    Select Case gErr
    
        Case 76399
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 76398, 76403, 76410
            'Erros tratados nas rotinas chamadas
            
        Case 76401, 76402
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_IMPORTPRODAUX", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161816)
            
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Public Function Importa_Transportadoras() As Long
'Importa os dados da tabela ImportTRANSP para a Tabela Transportadoras

Dim alComando(1 To 3) As Long
Dim lTransacao As Long
Dim lErro As Long
Dim sNome As String
Dim sNomeReduzido As String
Dim sCgc As String
Dim sInscricaoEstadual As String, sInscricaoMunicipal As String
Dim sGuia As String
Dim sEndereco As String
Dim iCodigo As Integer
Dim iCodTransp As Integer
Dim sBairro As String
Dim sCidade As String, sCidade2 As String
Dim sUF As String
Dim sCEP As String
Dim sFone As String, sTelefone2 As String
Dim sFax As String
Dim objTransportadora As New ClassTransportadora
Dim objEndereco As New ClassEndereco
Dim colTransportadora As New Collection
Dim colEndereco As New Collection, sEmail As String
Dim iIndice As Integer, iCodigoPais As Integer
Dim lCidade As Long, iViaTransporte As Integer
Dim lCodCid As Long, dPesoMinimo As Double, sContato As String

On Error GoTo Erro_Importa_Transportadoras

    'Abre a transação
    lTransacao = Transacao_Abrir
    If lTransacao = 0 Then gError 125826

    'Abrir comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir
        If alComando(iIndice) = 0 Then gError 125827
    Next

    'Inicializa as strings
    sNome = String(255, 0)
    sNomeReduzido = String(255, 0)
    sCgc = String(255, 0)
    sInscricaoEstadual = String(255, 0)
    sInscricaoMunicipal = String(255, 0)
    sGuia = String(255, 0)
    sEndereco = String(255, 0)
    sBairro = String(255, 0)
    sCidade = String(255, 0)
    sUF = String(255, 0)
    sCEP = String(255, 0)
    sFone = String(255, 0)
    sTelefone2 = String(255, 0)
    sFax = String(255, 0)
    sEmail = String(255, 0)
    sContato = String(255, 0)
    
    'Realiza a leitura na tabela TRANSPS
    lErro = Comando_Executar(alComando(1), "SELECT Codigo,Nome,NomeReduzido,CGC,InscricaoEstadual,InscricaoMunicipal,ViaTransporte,Guia,PesoMinimo,Endereco,Bairro,Cidade,SiglaEstado,CodigoPais,CEP,Telefone1,Telefone2,Email,Fax,Contato FROM ImportTransp ORDER BY Codigo", _
                                        iCodigo, sNome, sNomeReduzido, sCgc, sInscricaoEstadual, sInscricaoMunicipal, iViaTransporte, sGuia, dPesoMinimo, sEndereco, sBairro, sCidade, sUF, iCodigoPais, sCEP, sFone, sTelefone2, sEmail, sFax, sContato)
    If lErro <> AD_SQL_SUCESSO Then gError 125828
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 125829
    
    Do While lErro = AD_SQL_SUCESSO
    
        Set objTransportadora = New ClassTransportadora
        Set objEndereco = New ClassEndereco
    
        'Preenche o objTransportadora e o objEndereco
        With objTransportadora
        
            .iCodigo = iCodigo
            .iViaTransporte = 3
            .dPesoMinimo = 0
            .sCgc = sCgc
            .sGuia = sGuia
            .sInscricaoEstadual = sInscricaoEstadual
            .sNome = sNome
            .sNomeReduzido = sNomeReduzido
            
        End With
        
        With objEndereco
        
            .sBairro = sBairro
            .sCEP = sCEP
            .sCidade = sCidade
            .sEndereco = sEndereco
            .sFax = sFax
            .sSiglaEstado = sUF
            .sTelefone1 = sFone
            .sTelefone2 = sTelefone2
            .iCodigoPais = 1
            .sEmail = sEmail
            .sContato = sContato
            
        End With
        
        If lCidade <> 0 Then
        
            'Verifica se a Cidade está cadastrada no BD
            sCidade2 = String(255, 0)
            lErro = Comando_Executar(alComando(2), "SELECT Descricao FROM Cidades WHERE Codigo = ?", sCidade2, lCidade)
            If lErro <> AD_SQL_SUCESSO Then gError 125830
            
            lErro = Comando_BuscarPrimeiro(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 125831
            
            'se não estiver cadastrada
            If lErro = AD_SQL_SEM_DADOS Then gError 125832
            
'                lErro = Comando_Executar(alComando(3), "INSERT INTO Cidades (Codigo, Descricao) VALUES (?,?)", lCidade, sCidade)
'                If lErro <> AD_SQL_SUCESSO Then gError 125832
'
'            End If
        
            objEndereco.sCidade = sCidade2
            
        End If
        
        'Preenche as coleções: transportadora e Endereço
        colTransportadora.Add objTransportadora
        colEndereco.Add objEndereco
        
        'Busca o Próximo elemento
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 125833
        
    Loop
    
    'Libera comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Fecha a Transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 125834

    iIndice = 0
    
    'Realiza a gravação dos dados das coleções para o BD
    For Each objTransportadora In colTransportadora
    
        iIndice = iIndice + 1
    
        Set objEndereco = colEndereco.Item(iIndice)
    
        lErro = CF("Transportadora_Grava", objTransportadora, objEndereco)
        If lErro <> SUCESSO Then MsgBox (objTransportadora.iCodigo)
        
    Next
    
    Importa_Transportadoras = SUCESSO
    
    Exit Function
    
Erro_Importa_Transportadoras:

    Importa_Transportadoras = gErr
    
    Select Case gErr
    
        Case 125826
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 125827
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 125828, 125829, 125833
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRANSPS", gErr)
            
        Case 125830, 125831
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CIDADES", gErr, lCidade)
        
        Case 125832
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_CIDADES", gErr, lCidade)
            
        Case 125834
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case 125835, 125836
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161817)
        
    End Select
    
    'Libera comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Call Transacao_Rollback
    
    Exit Function
    
End Function

Function Importa_Kit()

Dim lErro As Long, iIndice As Integer, lTransacao As Long
Dim alComando(1 To 2) As Long
Dim tKit As typeKit
Dim objKit As ClassKit
Dim tProdutoKit As typeProdutoKit
Dim objProdutoKit As ClassProdutoKit
Dim dQuantidade As Double
Dim sUM As String
Dim iPos As Integer

On Error GoTo Erro_Importa_Kit

    'Abre a transação
    lTransacao = Transacao_Abrir
    If lTransacao = 0 Then gError 125826

    'Abrir comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir
        If alComando(iIndice) = 0 Then gError 125827
    Next

    tKit.sObservacao = String(255, 0)
    tKit.sProdutoRaiz = String(255, 0)
    tKit.sVersao = String(255, 0)
    sUM = String(255, 0)
    
    lErro = Comando_Executar(alComando(1), "SELECT ProdutoRaiz, Versao, Data, Observacao, Situacao, Quantidade, UnidadeMedida FROM ImportKit ORDER BY ProdutoRaiz, Versao", _
    tKit.sProdutoRaiz, tKit.sVersao, tKit.dtData, tKit.sObservacao, tKit.iSituacao, dQuantidade, sUM)
    If lErro <> AD_SQL_SUCESSO Then gError 999
    
    lErro = Comando_BuscarProximo(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 999
    
    iIndice = 0
    
    Do While lErro = AD_SQL_SUCESSO
    
        Set objProdutoKit = New ClassProdutoKit
        Set objKit = New ClassKit
        
        objKit.colComponentes.Add objProdutoKit
                
        objKit.dtData = tKit.dtData
        objKit.iSituacao = tKit.iSituacao
        objKit.iVersaoFormPreco = tKit.iSituacao
        objKit.sObservacao = tKit.sObservacao
        objKit.sProdutoRaiz = tKit.sProdutoRaiz
        objKit.sVersao = tKit.sVersao

'        'Parolibor
'        If Len(objKit.sProdutoRaiz) = 7 Then
'            objKit.sProdutoRaiz = objKit.sProdutoRaiz & Chr(32)
'        End If
'
'        If IsNumeric(Trim(Right(objKit.sProdutoRaiz, 4))) Then
'            If StrParaLong(Trim(Right(objKit.sProdutoRaiz, 4))) = 0 Then
'                objKit.sProdutoRaiz = Left(objKit.sProdutoRaiz, 4) & Chr(32) & Chr(32) & Chr(32) & Chr(32)
'            End If
'        End If
        
        iPos = 0
        
         Set objProdutoKit = New ClassProdutoKit
        Set objKit = New ClassKit
        
        objKit.colComponentes.Add objProdutoKit
       
        objProdutoKit.dQuantidade = dQuantidade
        objProdutoKit.iNivel = KIT_NIVEL_RAIZ
        objProdutoKit.sProduto = objKit.sProdutoRaiz
        objProdutoKit.sProdutoRaiz = objKit.sProdutoRaiz
        objProdutoKit.sVersao = objKit.sVersao
        objProdutoKit.sUnidadeMed = sUM
        objProdutoKit.iSeq = 1
        objProdutoKit.iSeqPai = 0
        objProdutoKit.dPercentualPerda = 0
        objProdutoKit.iPosicaoArvore = iPos
    
        iIndice = iIndice + 1
        
        tProdutoKit.sProduto = String(255, 0)
        tProdutoKit.sUnidadeMed = String(255, 0)

        lErro = Comando_Executar(alComando(2), "SELECT Ordem,ProdutoInsumo,QuantidadeInsumo,UnidadeMedInsumo,PercentualPerda FROM ImportKitInsumos WHERE ProdutoRaiz = ? AND Versao = ? ORDER BY Ordem ", _
        tProdutoKit.iSeq, tProdutoKit.sProduto, tProdutoKit.dQuantidade, tProdutoKit.sUnidadeMed, tProdutoKit.dPercentualPerda, tKit.sProdutoRaiz, tKit.sVersao)
        If lErro <> AD_SQL_SUCESSO Then gError 999
        
        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 999
        
        iIndice = 0
        
        Do While lErro = AD_SQL_SUCESSO
        
            iPos = iPos + 1
        
            Set objProdutoKit = New ClassProdutoKit
            
            objProdutoKit.dQuantidade = tProdutoKit.dQuantidade
            objProdutoKit.iNivel = 1
            objProdutoKit.sProduto = tProdutoKit.sProduto
            objProdutoKit.sProdutoRaiz = objKit.sProdutoRaiz
            objProdutoKit.sVersao = objKit.sVersao
            objProdutoKit.sUnidadeMed = tProdutoKit.sUnidadeMed
            objProdutoKit.iSeq = tProdutoKit.iSeq
            objProdutoKit.iSeqPai = 1
            objProdutoKit.dPercentualPerda = 0
            objProdutoKit.iPosicaoArvore = iPos
            
'            'Parolibor
'            If Len(objProdutoKit.sProduto) = 7 Then
'                objProdutoKit.sProduto = objProdutoKit.sProduto & Chr(32)
'            End If
'
'            If IsNumeric(Trim(Right(objProdutoKit.sProduto, 4))) Then
'                If StrParaLong(Trim(Right(objProdutoKit.sProduto, 4))) = 0 Then
'                    objProdutoKit.sProduto = Left(objProdutoKit.sProduto, 4) & Chr(32) & Chr(32) & Chr(32) & Chr(32)
'                End If
'            End If
            
            objKit.colComponentes.Add objProdutoKit
        
            lErro = Comando_BuscarProximo(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 999
        
        Loop
        
        lErro = CF("Kit_Grava_EmTrans", objKit)
        If lErro <> SUCESSO Then
            gError 99999
        End If
        
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 999
    
    Loop
    
    'Libera comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Fecha a Transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 125834

    Importa_Kit = SUCESSO
     
    Exit Function
    
Erro_Importa_Kit:

    Importa_Kit = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161818)
     
    End Select
     
    'Libera comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Call Transacao_Rollback
    
    Exit Function

End Function

Function Importa_Cidades()

Dim lErro As Long, iIndice As Long, lTransacao As Long
Dim alComando(1 To 3) As Long, sDescricao As String, lQtde As Long

On Error GoTo Erro_Importa_Cidades

    'Abre a transação
    lTransacao = Transacao_Abrir
    If lTransacao = 0 Then gError 125826

    'Abrir comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir
        If alComando(iIndice) = 0 Then gError 125827
    Next

    sDescricao = String(255, 0)
    
    lErro = Comando_Executar(alComando(1), "SELECT Descricao, SUM(Quantidade) FROM ImportCidades WHERE Descricao NOT IN (SELECT Descricao FROM Cidades) GROUP BY Descricao ORDER BY SUM(Quantidade) DESC", sDescricao, lQtde)
    If lErro <> AD_SQL_SUCESSO Then gError 999
    
    lErro = Comando_BuscarProximo(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 999
    
    lErro = CF("Config_ObterAutomatico_EmTrans", "FATConfig", "NUM_PROX_CIDADECADASTRO", "Cidades", "Codigo", iIndice)
    If lErro <> SUCESSO Then gError 125053
    
    Do While lErro = AD_SQL_SUCESSO
    
        If Len(Trim(sDescricao)) > 1 Then
        
            lErro = Comando_Executar(alComando(2), "INSERT INTO Cidades (Codigo,Descricao, CodIBGE) VALUES (?,?,?)", iIndice, Trim(sDescricao), "")
            If lErro <> AD_SQL_SUCESSO Then gError 999
            
        End If
        
        iIndice = iIndice + 1
        
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 999
    
    Loop
    
    'Libera comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Fecha a Transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 125834

    Importa_Cidades = SUCESSO
     
    Exit Function
    
Erro_Importa_Cidades:

    Importa_Cidades = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161818)
     
    End Select
     
    'Libera comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Call Transacao_Rollback
    
    Exit Function

End Function

Public Function Importa_Vendedores() As Long
'Importa os dados da tabela ImportTRANSP para a Tabela Transportadoras

Dim alComando(1 To 3) As Long
Dim lTransacao As Long
Dim lErro As Long
Dim sNome As String
Dim sNomeReduzido As String
Dim sCgc As String
Dim sInscricaoEstadual As String, sInscricaoMunicipal As String
Dim sGuia As String
Dim sEndereco As String
Dim iCodigo As Integer
Dim iCodTransp As Integer
Dim sBairro As String
Dim sCidade As String, sCidade2 As String
Dim sUF As String
Dim sCEP As String
Dim sFone As String, sTelefone2 As String
Dim sFax As String
Dim objVendedor As New ClassVendedor
Dim objEndereco As New ClassEndereco
Dim colVend As New Collection, dPercComissao As Double, dPercComissaoEmissao As Double, dPercComissaoBaixa As Double
Dim colEndereco As New Collection, sEmail As String
Dim iIndice As Integer, iCodigoPais As Integer
Dim lCidade As Long, iViaTransporte As Integer, sRazaoSocial As String
Dim lCodCid As Long, dPesoMinimo As Double, sContato As String
Dim objVendAux As ClassVendedor

On Error GoTo Erro_Importa_Vendedores

    'Abre a transação
    lTransacao = Transacao_Abrir
    If lTransacao = 0 Then gError 125826

    'Abrir comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir
        If alComando(iIndice) = 0 Then gError 125827
    Next

    'Inicializa as strings
    sNome = String(255, 0)
    sNomeReduzido = String(255, 0)
    sCgc = String(255, 0)
    sInscricaoEstadual = String(255, 0)
    sInscricaoMunicipal = String(255, 0)
    sGuia = String(255, 0)
    sEndereco = String(255, 0)
    sBairro = String(255, 0)
    sCidade = String(255, 0)
    sUF = String(255, 0)
    sCEP = String(255, 0)
    sFone = String(255, 0)
    sTelefone2 = String(255, 0)
    sFax = String(255, 0)
    sEmail = String(255, 0)
    sContato = String(255, 0)
    sRazaoSocial = String(255, 0)
    
    'Realiza a leitura na tabela
    lErro = Comando_Executar(alComando(1), "SELECT Codigo,Nome,NomeReduzido,PercComissao,PercComissaoBaixa,PercComissaoEmissao," & _
        "CGC,InscricaoEstadual,RazaoSocial,Endereco,Bairro,Cidade,SiglaEstado,CodigoPais,CEP,Telefone1,Telefone2,Email,Fax,Contato FROM ImportVend WHERE Codigo > 0 AND Codigo NOT IN (SELECT Codigo FROM Vendedores) ORDER BY Codigo", _
        iCodigo, sNome, sNomeReduzido, dPercComissao, dPercComissaoBaixa, _
        dPercComissaoEmissao, _
        sCgc, sInscricaoEstadual, sRazaoSocial, sEndereco, sBairro, sCidade, sUF, iCodigoPais, sCEP, _
        sFone, sTelefone2, sEmail, sFax, sContato)
    If lErro <> AD_SQL_SUCESSO Then gError 125828
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 125829
    
    Do While lErro = AD_SQL_SUCESSO
    
        Set objVendedor = New ClassVendedor
        Set objEndereco = New ClassEndereco
    
        'Preenche o objTransportadora e o objEndereco
        With objVendedor
        
            .iCodigo = iCodigo
            .sCgc = Trim(sCgc)
            
            .sCgc = Replace(.sCgc, " ", "")
            .sCgc = Replace(.sCgc, ".", "")
            .sCgc = Replace(.sCgc, "-", "")
            .sCgc = Replace(.sCgc, "/", "")
            .sCgc = Replace(.sCgc, "\", "")
            .sCgc = Replace(.sCgc, "_", "")
            .sCgc = Replace(.sCgc, " ", "")
            
            If Len(.sCgc) > 11 And Len(.sCgc) < 14 Then
                .sCgc = String(14 - Len(.sCgc), "0") & .sCgc
            ElseIf Len(.sCgc) > 1 And Len(.sCgc) < 11 Then
                .sCgc = String(11 - Len(.sCgc), "0") & .sCgc
            End If
            
            .sInscricaoEstadual = Trim(sInscricaoEstadual)
            .sInscricaoEstadual = Replace(.sInscricaoEstadual, " ", "")
            .sInscricaoEstadual = Replace(.sInscricaoEstadual, ".", "")
            .sInscricaoEstadual = Replace(.sInscricaoEstadual, "-", "")
            .sInscricaoEstadual = Replace(.sInscricaoEstadual, "/", "")
            .sInscricaoEstadual = Replace(.sInscricaoEstadual, "\", "")
            .sInscricaoEstadual = Replace(.sInscricaoEstadual, "_", "")
            .sInscricaoEstadual = Replace(.sInscricaoEstadual, "ISENTO", "")
            .sInscricaoEstadual = Replace(.sInscricaoEstadual, "ISENTA", "")
            
            
            .sNome = Trim(sNome)
            If .sNome = "" Then .sNome = "SEM NOME"
            
            .sNomeReduzido = Trim(sNomeReduzido)
            If .sNomeReduzido = "" Then .sNomeReduzido = left(.sNome, 20)
            
            .dPercComissao = dPercComissao
            .dPercComissaoBaixa = dPercComissaoBaixa
            .dPercComissaoEmissao = dPercComissaoEmissao
            
            If .dPercComissaoBaixa = 0 And .dPercComissaoEmissao = 0 Then
                .dPercComissaoEmissao = 1
            End If
            
            .iAtivo = 1
            .iComissaoFrete = 1
            .iComissaoICM = 1
            .iComissaoIPI = 1
            .iComissaoSeguro = 1
            .iComissaoSobreTotal = 1
            
        End With
        
        With objEndereco
        
            .sBairro = Trim(sBairro)
            .sCEP = sCEP
            
            .sCEP = Replace(.sCEP, " ", "")
            .sCEP = Replace(.sCEP, ".", "")
            .sCEP = Replace(.sCEP, "-", "")
            .sCEP = Replace(.sCEP, "/", "")
            .sCEP = Replace(.sCEP, "\", "")
            .sCEP = Replace(.sCEP, "_", "")
            .sCEP = Replace(.sCEP, " ", "")
            
            .sCEP = String(8 - Len(.sCEP), "0") + .sCEP
            If .sCEP = String(8, "0") Then .sCEP = ""
    
            .sCidade = Trim(sCidade)
            .sEndereco = Trim(sEndereco)
            .sFax = Trim(sFax)
            .sSiglaEstado = Trim(sUF)
            .sTelefone1 = Trim(sFone)
            .sTelefone2 = Trim(sTelefone2)
            .iCodigoPais = 1
            .sEmail = Trim(sEmail)
            .sContato = Trim(sContato)
            
        End With
'
'        If lCidade <> 0 Then
'
'            'Verifica se a Cidade está cadastrada no BD
'            sCidade2 = String(255, 0)
'            lErro = Comando_Executar(alComando(2), "SELECT Descricao FROM Cidades WHERE Codigo = ?", sCidade2, lCidade)
'            If lErro <> AD_SQL_SUCESSO Then gError 125830
'
'            lErro = Comando_BuscarPrimeiro(alComando(2))
'            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 125831
'
'            'se não estiver cadastrada
'            If lErro = AD_SQL_SEM_DADOS Then gError 125832
'
''                lErro = Comando_Executar(alComando(3), "INSERT INTO Cidades (Codigo, Descricao) VALUES (?,?)", lCidade, sCidade)
''                If lErro <> AD_SQL_SUCESSO Then gError 125832
''
''            End If
'
'            objEndereco.sCidade = sCidade2
'
'        End If
'
        'Preenche as coleções: transportadora e Endereço
        colVend.Add objVendedor
        colEndereco.Add objEndereco
        
        'Busca o Próximo elemento
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 125833
        
    Loop
    
    'Libera comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Fecha a Transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 125834

    iIndice = 0
    
    'Realiza a gravação dos dados das coleções para o BD
    For Each objVendedor In colVend
    
        iIndice = iIndice + 1
    
        Set objEndereco = colEndereco.Item(iIndice)
        
        Set objVendAux = New ClassVendedor
        objVendAux.sNomeReduzido = objVendedor.sNomeReduzido
        
        lErro = CF("Vendedor_Le_NomeReduzido", objVendAux)
        If lErro <> SUCESSO And lErro <> 25008 Then gError 125833
        
        If lErro = SUCESSO Then
            If objVendAux.iCodigo <> objVendedor.iCodigo Then
                objVendedor.sNomeReduzido = left(objVendedor.sNomeReduzido, 14) + "-" + CStr(objVendedor.iCodigo)
            End If
        End If
    
        lErro = CF("Vendedor_Grava", objVendedor, objEndereco)
        If lErro <> SUCESSO Then MsgBox (objVendedor.iCodigo)
        
    Next
    
    Importa_Vendedores = SUCESSO
    
    Exit Function
    
Erro_Importa_Vendedores:

    Importa_Vendedores = gErr
    
    Select Case gErr
    
        Case 125826
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 125827
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 125828, 125829, 125833
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRANSPS", gErr)
            
        Case 125830, 125831
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CIDADES", gErr, lCidade)
        
        Case 125832
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_CIDADES", gErr, lCidade)
            
        Case 125834
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case 125835, 125836
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161819)
        
    End Select
    
    'Libera comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Call Transacao_Rollback
    
    Exit Function
    
End Function

Function Importa_Fornecedores() As Long
'Importa os dados da tabela ImportForn
'???se o fornecedor nao tem filial=1 troca a primeira filial encontrada para filial1
Dim lErro As Long
Dim tImportForn As typeImportForn
Dim lComando As Long
Dim lTransacao As Long
Dim objFilialForn As New ClassFilialFornecedor
Dim objFornecedor As New ClassFornecedor
Dim objEndereco As New ClassEndereco
Dim bPrimeiraFilial As Boolean

On Error GoTo Erro_Importa_Fornecedores

    bPrimeiraFilial = False
    
    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 76357
    
    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 76358
    
    tImportForn.sBairro = String(255, 0)
    tImportForn.sCEP = String(255, 0)
    tImportForn.sCidade = String(255, 0)
    tImportForn.sContato = String(255, 0)
    tImportForn.sEmail = String(255, 0)
    tImportForn.sEndereco = String(255, 0)
    tImportForn.sFax = String(255, 0)
    tImportForn.sFilialAgencia = String(255, 0)
    tImportForn.sFilialCGC = String(255, 0)
    tImportForn.sFilialContaContabil = String(255, 0)
    tImportForn.sFilialContaCorrente = String(255, 0)
    tImportForn.sFilialInscEstadual = String(255, 0)
    tImportForn.sFilialInscMunicipal = String(255, 0)
    tImportForn.sFilialNome = String(255, 0)
    tImportForn.sFilialObservacao1 = String(255, 0)
    tImportForn.sNomeReduzido = String(255, 0)
    tImportForn.sObservacao = String(255, 0)
    tImportForn.sRazaoSocial = String(255, 0)
    tImportForn.sSiglaEstado = String(255, 0)
    tImportForn.sTelefone1 = String(255, 0)
    tImportForn.sTelefone2 = String(255, 0)
    
    'Lê os registros da tabela ImportForn
    lErro = Comando_Executar(lComando, "SELECT Codigo,RazaoSocial,NomeReduzido,Tipo,Observacao,CondicaoPagto,Desconto,ProxCodFilial,FilialCod," _
    & "FilialNome,FilialCGC,FilialInscricaoEstadual,FilialInscricaoMunicipal,FilialContaContabil,FilialBanco,FilialAgencia,FilialContaCorrente,FilialObservacao1,FilialTipoFrete,Endereco,Bairro,Cidade," _
    & "SiglaEstado,CodigoPais,CEP,Telefone1,Telefone2,Email,Fax,Contato FROM ImportForn ORDER BY Codigo,FilialCod", tImportForn.lCodigo, tImportForn.sRazaoSocial, tImportForn.sNomeReduzido, tImportForn.iTipo, tImportForn.sObservacao, _
    tImportForn.iCondicaoPagto, tImportForn.dDesconto, tImportForn.iProxCodFilial, tImportForn.iFilialCod, tImportForn.sFilialNome, tImportForn.sFilialCGC, tImportForn.sFilialInscEstadual, tImportForn.sFilialInscMunicipal, _
    tImportForn.sFilialContaContabil, tImportForn.iFilialBanco, tImportForn.sFilialAgencia, tImportForn.sFilialContaCorrente, tImportForn.sFilialObservacao1, tImportForn.iFilialTipoFrete, tImportForn.sEndereco, tImportForn.sBairro, _
    tImportForn.sCidade, tImportForn.sSiglaEstado, tImportForn.iCodigoPais, tImportForn.sCEP, tImportForn.sTelefone1, tImportForn.sTelefone2, tImportForn.sEmail, tImportForn.sFax, tImportForn.sContato)
    If lErro <> AD_SQL_SUCESSO Then gError 76359
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 76360
    
    Do While lErro = AD_SQL_SUCESSO
        
'        bPrimeiraFilial = False
'
'        'Verifica se mudou o fornecedor
'        If tImportForn.lCodigo <> objFilialForn.lCodFornecedor Then
'            bPrimeiraFilial = True
'        End If
        
        objFilialForn.lCodFornecedor = tImportForn.lCodigo
        objFilialForn.iCodFilial = tImportForn.iFilialCod
        
        If objFilialForn.iCodFilial = 0 Then objFilialForn.iCodFilial = FILIAL_MATRIZ
        
        'Verifica se já existe a Filial do Fornecedor lido na tabela FiliaisFornecedores
        lErro = CF("FilialFornecedor_Le", objFilialForn)
        If lErro <> SUCESSO And lErro <> 12929 Then gError 76361

        'Se não existe a filial na tabela FiliaisFornecedores
        If lErro = 12929 Then
            
            Call ImportForn_Preenche_Endereco(tImportForn, objEndereco)
            
'            'Se o a primeira filial do Fornecedor lido não é a matriz
'            If bPrimeiraFilial = True And objFilialForn.iCodFilial <> FILIAL_MATRIZ Then
'                'muda o codigo da filial
'                objFilialForn.iCodFilial = FILIAL_MATRIZ
'            End If
'
'            If objFilialForn.iCodFilial = FILIAL_MATRIZ Then
            
                'Preenche objFornecedor a partir de tImportForn
                lErro = Fornecedor_PreencheObjetos(tImportForn, objFornecedor)
                If lErro <> SUCESSO Then gError 76362
                
                'Grava o Fornecedor
                lErro = CF("Fornecedor_Grava_EmTrans", objFornecedor, objEndereco)
                If lErro <> SUCESSO Then gError 76363

'            Else
'
'                'Preenche objFilialForn a partir de tImportForn
'                lErro = FilialFornecedor_PreencheObjetos(tImportForn, objFilialForn)
'                If lErro <> SUCESSO Then gError 76367
'
'                'Grava a Filial Fornecedor
'                lErro = CF("FiliaisFornecedores_Grava_EmTrans", objFilialForn, objEndereco)
'                If lErro <> SUCESSO Then gError 76368
'
'            End If
        
        End If
       
        'Busca o proximo registro de ImportForn
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 76369
        
    Loop
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 76370
    
    Call Comando_Fechar(lComando)

    Importa_Fornecedores = SUCESSO
    
    Exit Function
    
Erro_Importa_Fornecedores:

    Importa_Fornecedores = gErr
    
    Select Case gErr
    
        Case 76357
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 76358
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 76359, 76360, 76369
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_IMPORTFORN", gErr)
        
        Case 76361, 76362, 76363, 76367, 76368
            'Erros tratados nas rotinas chamadas
            
        Case 76370
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT_TRANSACAO", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161820)
            
    End Select
    
    Call Transacao_Rollback
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function Fornecedor_PreencheObjetos(tImportForn As typeImportForn, ByVal objFornecedor As ClassFornecedor) As Long
'Preenche objFornecedor e colEndereco a partir dos dados existentes em tImportForn

Dim lErro As Long
Dim objFornAUx As New ClassFornecedor

On Error GoTo Erro_Fornecedor_PreencheObjetos

    objFornecedor.dDesconto = tImportForn.dDesconto
    objFornecedor.iBanco = tImportForn.iFilialBanco
    objFornecedor.iCondicaoPagto = tImportForn.iCondicaoPagto
    objFornecedor.iProxCodFilial = 2
    objFornecedor.iTipo = tImportForn.iTipo
    objFornecedor.iAtivo = MARCADO
    
    If objFornecedor.iTipo = 0 Then objFornecedor.iTipo = 9999
    objFornecedor.lCodigo = tImportForn.lCodigo
    objFornecedor.sAgencia = Trim(tImportForn.sFilialAgencia)
    objFornecedor.sContaContabil = tImportForn.sFilialContaContabil
    objFornecedor.sContaCorrente = Trim(tImportForn.sFilialContaCorrente)
    
    objFornecedor.sInscricaoEstadual = tImportForn.sFilialInscEstadual
    objFornecedor.sInscricaoEstadual = Replace(objFornecedor.sInscricaoEstadual, " ", "")
    objFornecedor.sInscricaoEstadual = Replace(objFornecedor.sInscricaoEstadual, ".", "")
    objFornecedor.sInscricaoEstadual = Replace(objFornecedor.sInscricaoEstadual, "-", "")
    objFornecedor.sInscricaoEstadual = Replace(objFornecedor.sInscricaoEstadual, "/", "")
    objFornecedor.sInscricaoEstadual = Replace(objFornecedor.sInscricaoEstadual, "\", "")
    objFornecedor.sInscricaoEstadual = Replace(objFornecedor.sInscricaoEstadual, "_", "")
    objFornecedor.sInscricaoEstadual = Replace(objFornecedor.sInscricaoEstadual, "ISENTO", "")
    objFornecedor.sInscricaoEstadual = Replace(objFornecedor.sInscricaoEstadual, "ISENTA", "")
    If Len(Trim(objFornecedor.sInscricaoEstadual)) > 0 Then
        objFornecedor.iIEIsento = DESMARCADO
        objFornecedor.iIENaoContrib = DESMARCADO
    Else
        objFornecedor.iIEIsento = MARCADO
        objFornecedor.iIENaoContrib = MARCADO
    End If
    objFornecedor.iRegimeTributario = REGIME_TRIBUTARIO_NORMAL
    
    objFornecedor.sInscricaoMunicipal = tImportForn.sFilialInscMunicipal
    objFornecedor.sInscricaoMunicipal = Replace(objFornecedor.sInscricaoMunicipal, " ", "")
    objFornecedor.sInscricaoMunicipal = Replace(objFornecedor.sInscricaoMunicipal, ".", "")
    objFornecedor.sInscricaoMunicipal = Replace(objFornecedor.sInscricaoMunicipal, "-", "")
    objFornecedor.sInscricaoMunicipal = Replace(objFornecedor.sInscricaoMunicipal, "/", "")
    objFornecedor.sInscricaoMunicipal = Replace(objFornecedor.sInscricaoMunicipal, "\", "")
    objFornecedor.sInscricaoMunicipal = Replace(objFornecedor.sInscricaoMunicipal, "_", "")
    objFornecedor.sInscricaoMunicipal = Replace(objFornecedor.sInscricaoMunicipal, "ISENTO", "")
    objFornecedor.sInscricaoMunicipal = Replace(objFornecedor.sInscricaoMunicipal, "ISENTA", "")
        
    objFornecedor.sObservacao = Trim(tImportForn.sObservacao)
    objFornecedor.sObservacao2 = Trim(tImportForn.sFilialObservacao1)
    objFornecedor.sRazaoSocial = Trim(tImportForn.sRazaoSocial)
    If objFornecedor.sRazaoSocial = "" Then objFornecedor.sRazaoSocial = "SEM NOME"
    
    objFornecedor.sNomeReduzido = Trim(tImportForn.sNomeReduzido)
    If objFornecedor.sNomeReduzido = "" Then objFornecedor.sNomeReduzido = left(objFornecedor.sRazaoSocial, STRING_FORNECEDOR_NOME_REDUZIDO)
    
    objFornAUx.sNomeReduzido = objFornecedor.sNomeReduzido

    lErro = CF("Fornecedor_Le_NomeReduzido", objFornAUx)
    If lErro <> SUCESSO And lErro <> 6681 Then gError ERRO_SEM_MENSAGEM

    If lErro = SUCESSO Then
        If objFornAUx.lCodigo <> objFornecedor.lCodigo Then
            objFornecedor.sNomeReduzido = left(objFornecedor.sNomeReduzido, STRING_FORNECEDOR_NOME_REDUZIDO - 7) + "-" + CStr(objFornecedor.lCodigo)
        End If
    End If
    
    objFornecedor.sCgc = tImportForn.sFilialCGC
    objFornecedor.sCgc = Replace(objFornecedor.sCgc, " ", "")
    objFornecedor.sCgc = Replace(objFornecedor.sCgc, ".", "")
    objFornecedor.sCgc = Replace(objFornecedor.sCgc, "-", "")
    objFornecedor.sCgc = Replace(objFornecedor.sCgc, "/", "")
    objFornecedor.sCgc = Replace(objFornecedor.sCgc, "\", "")
    objFornecedor.sCgc = Replace(objFornecedor.sCgc, "_", "")
    objFornecedor.sCgc = Replace(objFornecedor.sCgc, " ", "")
    
    If Len(objFornecedor.sCgc) > 11 And Len(objFornecedor.sCgc) < 14 Then
        objFornecedor.sCgc = String(14 - Len(objFornecedor.sCgc), "0") & objFornecedor.sCgc
    ElseIf Len(objFornecedor.sCgc) > 1 And Len(objFornecedor.sCgc) < 11 Then
        objFornecedor.sCgc = String(11 - Len(objFornecedor.sCgc), "0") & objFornecedor.sCgc
    End If
    
'    If Len(Trim(tImportForn.sFilialCGC)) > 11 And Len(Trim(tImportForn.sFilialCGC)) <> 14 Then
'        objFornecedor.sCgc = Format(Trim(tImportForn.sFilialCGC), "00000000000000")
'    Else
'        If Len(Trim(tImportForn.sFilialCGC)) > 8 And Len(Trim(tImportForn.sFilialCGC)) <> 11 Then
'            objFornecedor.sCgc = Format(Trim(tImportForn.sFilialCGC), "00000000000")
'        Else
'            objFornecedor.sCgc = tImportForn.sFilialCGC
'        End If
'    End If
        
''    Select Case Len(Trim(tImportForn.sFilialCGC))
''
''    Case STRING_CPF 'CPF
''
''        'Critica CPF
''        lErro = Cpf_Critica(tImportForn.sFilialCGC)
''        If lErro <> SUCESSO Then gError 76364
''
''        objFornecedor.sCgc = tImportForn.sFilialCGC
''
''    Case STRING_CGC 'CGC
''
''        'Critica CGC
''        lErro = Cgc_Critica(tImportForn.sFilialCGC)
''        If lErro <> SUCESSO Then gError 76365
''
''        objFornecedor.sCgc = tImportForn.sFilialCGC
''
''    Case Else
''
''        objFornecedor.sCgc = ""
''
''    End Select

    Fornecedor_PreencheObjetos = SUCESSO
    
    Exit Function
    
Erro_Fornecedor_PreencheObjetos:

    Fornecedor_PreencheObjetos = gErr
    
    Select Case gErr
    
        Case 76364, 76365
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161821)
            
    End Select
    
    Exit Function
    
End Function

Sub ImportForn_Preenche_Endereco(tImportForn As typeImportForn, objEndereco As ClassEndereco)
'Preenche objEndereco
    
    Set objEndereco = New ClassEndereco
    
    objEndereco.iCodigoPais = tImportForn.iCodigoPais
    If objEndereco.iCodigoPais = 0 Then objEndereco.iCodigoPais = 1
    
    objEndereco.sBairro = Trim(tImportForn.sBairro)
    objEndereco.sCEP = Trim(tImportForn.sCEP)
    
    objEndereco.sCEP = Replace(objEndereco.sCEP, " ", "")
    objEndereco.sCEP = Replace(objEndereco.sCEP, ".", "")
    objEndereco.sCEP = Replace(objEndereco.sCEP, "-", "")
    objEndereco.sCEP = Replace(objEndereco.sCEP, "/", "")
    objEndereco.sCEP = Replace(objEndereco.sCEP, "\", "")
    objEndereco.sCEP = Replace(objEndereco.sCEP, "_", "")
    objEndereco.sCEP = Replace(objEndereco.sCEP, " ", "")
    
    objEndereco.sCEP = String(8 - Len(objEndereco.sCEP), "0") + objEndereco.sCEP
    If objEndereco.sCEP = String(8, "0") Then objEndereco.sCEP = ""
    
    objEndereco.sCidade = Trim(tImportForn.sCidade)
    If Len(Trim(objEndereco.sCidade)) < 2 Then objEndereco.sCidade = ""

    objEndereco.sContato = Trim(tImportForn.sContato)
    objEndereco.sEmail = Trim(tImportForn.sEmail)
    objEndereco.sEndereco = Trim(tImportForn.sEndereco)
    objEndereco.sSiglaEstado = Trim(tImportForn.sSiglaEstado)
    
    If objEndereco.sSiglaEstado = "" Then objEndereco.sSiglaEstado = "RJ"
    
    objEndereco.sTelefone1 = Trim(tImportForn.sTelefone1)
    
    objEndereco.sTelefone2 = Trim(tImportForn.sTelefone2)
    
    objEndereco.sFax = Trim(tImportForn.sFax)
    
End Sub

Function FilialFornecedor_PreencheObjetos(tImportForn As typeImportForn, objFilialForn As ClassFilialFornecedor) As Long
'Preenche objFilialForn a partir dos dados existentes em tImportForn

Dim lErro As Long

On Error GoTo Erro_FilialFornecedor_PreencheObjetos

    objFilialForn.iBanco = tImportForn.iFilialBanco
    objFilialForn.iCodFilial = tImportForn.iFilialCod
    objFilialForn.iTipoFrete = tImportForn.iFilialTipoFrete
    objFilialForn.lCodFornecedor = tImportForn.lCodigo
    objFilialForn.sAgencia = tImportForn.sFilialAgencia
    objFilialForn.sContaContabil = tImportForn.sFilialContaContabil
    objFilialForn.sContaCorrente = tImportForn.sFilialContaCorrente
    objFilialForn.sContato = tImportForn.sContato
    objFilialForn.sInscricaoEstadual = tImportForn.sFilialInscEstadual
    objFilialForn.sInscricaoMunicipal = tImportForn.sFilialInscMunicipal
    objFilialForn.sNome = tImportForn.sFilialNome
    objFilialForn.sObservacao = tImportForn.sFilialObservacao1
    
'    If Len(Trim(tImportForn.sFilialCGC)) > 11 And Len(Trim(tImportForn.sFilialCGC)) <> 14 Then
'        objFilialForn.sCgc = Format(Trim(tImportForn.sFilialCGC), "00000000000000")
'    Else
'        If Len(Trim(tImportForn.sFilialCGC)) > 8 And Len(Trim(tImportForn.sFilialCGC)) <> 11 Then
'            objFilialForn.sCgc = Format(Trim(tImportForn.sFilialCGC), "00000000000")
'        Else
'            objFilialForn.sCgc = tImportForn.sFilialCGC
'        End If
'    End If

    objFilialForn.sCgc = tImportForn.sFilialCGC
    objFilialForn.sCgc = Replace(objFilialForn.sCgc, " ", "")
    objFilialForn.sCgc = Replace(objFilialForn.sCgc, ".", "")
    objFilialForn.sCgc = Replace(objFilialForn.sCgc, "-", "")
    objFilialForn.sCgc = Replace(objFilialForn.sCgc, "/", "")
    objFilialForn.sCgc = Replace(objFilialForn.sCgc, "\", "")
    objFilialForn.sCgc = Replace(objFilialForn.sCgc, "_", "")
    objFilialForn.sCgc = Replace(objFilialForn.sCgc, " ", "")

    If Len(objFilialForn.sCgc) > 11 And Len(objFilialForn.sCgc) < 14 Then
        objFilialForn.sCgc = String(14 - Len(objFilialForn.sCgc), "0") & objFilialForn.sCgc
    ElseIf Len(objFilialForn.sCgc) > 1 And Len(objFilialForn.sCgc) < 11 Then
        objFilialForn.sCgc = String(11 - Len(objFilialForn.sCgc), "0") & objFilialForn.sCgc
    End If
    
    If Replace(objFilialForn.sCgc, "0", "") = "" Then objFilialForn.sCgc = ""

        
''    Select Case Len(Trim(tImportForn.sFilialCGC))
''
''    Case STRING_CPF 'CPF
''
''        'Critica CPF
''        lErro = Cpf_Critica(tImportForn.sFilialCGC)
''        If lErro <> SUCESSO Then gError 76371
''
''        objFilialForn.sCgc = tImportForn.sFilialCGC
''
''    Case STRING_CGC 'CGC
''
''        'Critica CGC
''        lErro = Cgc_Critica(tImportForn.sFilialCGC)
''        If lErro <> SUCESSO Then gError 76372
''
''        objFilialForn.sCgc = tImportForn.sFilialCGC
''
''    Case Else
''
''        objFilialForn.sCgc = ""
''
''    End Select

    FilialFornecedor_PreencheObjetos = SUCESSO
    
    Exit Function
    
Erro_FilialFornecedor_PreencheObjetos:

    FilialFornecedor_PreencheObjetos = gErr
    
    Select Case gErr
    
        Case 76371, 76372
            'Erros tratados nas rotinas chamadas
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161822)
            
    End Select
    
    Exit Function
    
End Function

Function Importa_Clientes() As Long
'Importa os dados da tabela ImportCli para criar clientes e suas filiais
'obs.: se o cliente nao tem filial=1 troca a primeira filial encontrada para filial1

Dim lErro As Long, iFilial As Integer
Dim tImportCli As typeImportCli
Dim lComando As Long, lComando2 As Long
Dim lTransacao As Long
Dim objFilialCliente As New ClassFilialCliente
Dim objCliente As New ClassCliente
Dim colEndereco As New Collection
Dim lCodigo As Long, lClienteAnterior As Long
Dim sRamoAtividade As String 'Kit Gourmet'
Dim sCGCAnterior As String, iQtdeCadastrada As Integer
Dim bTerminou As Boolean

On Error GoTo Erro_Importa_Clientes

    bTerminou = False
    
    Do While Not bTerminou
    
        iQtdeCadastrada = 0

        'Abre o comando
        lComando = Comando_Abrir()
        If lComando = 0 Then gError 76329
        
        lComando2 = Comando_Abrir()
        If lComando2 = 0 Then gError 76329
       
        'Inicia a Transacao
        lTransacao = Transacao_Abrir()
        If lTransacao = 0 Then gError 76342
        
        tImportCli.sBairro = String(255, 0)
        tImportCli.sBairro1 = String(255, 0)
        tImportCli.sBairro2 = String(255, 0)
        tImportCli.sCEP = String(255, 0)
        tImportCli.sCEP1 = String(255, 0)
        tImportCli.sCEP2 = String(255, 0)
        tImportCli.sCidade = String(255, 0)
        tImportCli.sCidade1 = String(255, 0)
        tImportCli.sCidade2 = String(255, 0)
        tImportCli.sContato = String(255, 0)
        tImportCli.sContato1 = String(255, 0)
        tImportCli.sContato2 = String(255, 0)
        tImportCli.sEmail = String(255, 0)
        tImportCli.sEmail1 = String(255, 0)
        tImportCli.sEmail2 = String(255, 0)
        tImportCli.sEndereco = String(255, 0)
        tImportCli.sEndereco1 = String(255, 0)
        tImportCli.sEndereco2 = String(255, 0)
        tImportCli.sFax = String(255, 0)
        tImportCli.sFax1 = String(255, 0)
        tImportCli.sFax2 = String(255, 0)
        tImportCli.sFilialCGC = String(255, 0)
        tImportCli.sFilialContaContabil = String(255, 0)
        tImportCli.sFilialInscEstadual = String(255, 0)
        tImportCli.sFilialInscMunicipal = String(255, 0)
        tImportCli.sFilialNome = String(255, 0)
        tImportCli.sFilialObservacao1 = String(255, 0)
        tImportCli.sFilialTipoFrete = String(255, 0)
        tImportCli.sNomeReduzido = String(255, 0)
        tImportCli.sObservacao = String(255, 0)
        tImportCli.sRazaoSocial = String(255, 0)
        tImportCli.sSiglaEstado = String(255, 0)
        tImportCli.sSiglaEstado1 = String(255, 0)
        tImportCli.sSiglaEstado2 = String(255, 0)
        tImportCli.sTelefone1 = String(255, 0)
        tImportCli.sTelefone11 = String(255, 0)
        tImportCli.sTelefone12 = String(255, 0)
        tImportCli.sTelefone2 = String(255, 0)
        tImportCli.sTelefone21 = String(255, 0)
        tImportCli.sTelefone22 = String(255, 0)
        sRamoAtividade = String(255, 0)
        
        lClienteAnterior = -1
        sCGCAnterior = ""
        iQtdeCadastrada = 0
        
        'Lê os registros da tabela ImportCli
        With tImportCli
        lErro = Comando_ExecutarPos(lComando, "SELECT Codigo,RazaoSocial,NomeReduzido,Tipo,Observacao,LimiteCredito,CondicaoPagto,Desconto,CodPadraoCobranca,CodMensagem,TabelaPreco,NumPagamentos,CodFilial," _
            & "FilialNome,FilialCGC,FilialInscricaoEstadual,FilialInscricaoMunicipal,FilialCodTransportadora,FilialObservacao1,FilialContaContabil,FilialVendedor,FilialComissaoVendas,FilialRegiao,FilialFreqVisitas,FilialDataUltVisita,FilialCodCobrador," _
            & "FilialICMSBaseCalculoComIPI,FilialRevendedor,FilialTipoFrete,Endereco,Bairro,Cidade,SiglaEstado,CodigoPais,CEP,Telefone1,Telefone2,Email,Fax,Contato,Endereco1,Bairro1,Cidade1,SiglaEstado1,CodigoPais1,CEP1,Telefone11,Telefone21,Email1,Fax1,Contato1," _
            & "Endereco2,Bairro2,Cidade2,SiglaEstado2,CodigoPais2,CEP2,Telefone12,Telefone22,Email2,Fax2,Contato2 FROM ImportCli WHERE codigo_corporator = 0 ORDER BY Codigo", 0, .lCodCliente, .sRazaoSocial, .sNomeReduzido, .iTipo, .sObservacao, .dLimiteCredito, _
            .iCondicaoPagto, .dDesconto, .iCodPadraoCobranca, .iCodMensagem, .iTabelaPreco, .lNumPagamentos, .iCodFilial, .sFilialNome, .sFilialCGC, .sFilialInscEstadual, .sFilialInscMunicipal, _
            .iFilialCodTransportadora, .sFilialObservacao1, .sFilialContaContabil, .iFilialVendedor, .dFilialComissaoVendas, .iFilialRegiao, .iFilialFreqVisitas, .dtFilialDataUltVisita, .iFilialCodCobrador, .iFilialICMSBaseCalculoIPI, _
            .lFilialRevendedor, .sFilialTipoFrete, .sEndereco, .sBairro, .sCidade, .sSiglaEstado, .iCodigoPais, tImportCli.sCEP, .sTelefone1, .sTelefone2, .sEmail, .sFax, .sContato, .sEndereco1, .sBairro1, _
            .sCidade1, .sSiglaEstado1, .iCodigoPais1, .sCEP1, .sTelefone11, .sTelefone21, .sEmail1, .sFax1, tImportCli.sContato1, .sEndereco2, .sBairro2, .sCidade2, .sSiglaEstado2, .iCodigoPais2, .sCEP2, .sTelefone12, _
            .sTelefone22, .sEmail2, .sFax2, .sContato2)
        End With
        If lErro <> AD_SQL_SUCESSO Then gError 76330
        
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 76331
        
        Do While lErro = AD_SQL_SUCESSO
        
    '        If sCGCAnterior <> tImportCli.sFilialCGC Or tImportCli.sFilialCGC = "" Then
    '
    '            sCGCAnterior = tImportCli.sFilialCGC
                
                objFilialCliente.lCodCliente = tImportCli.lCodCliente
                objFilialCliente.iCodFilial = tImportCli.iCodFilial
                
                If objFilialCliente.iCodFilial = 0 Then objFilialCliente.iCodFilial = FILIAL_MATRIZ

    
                'Verifica se já existe a Filial do Cliente lido na tabela FiliaisClientes
                lErro = CF("FilialCliente_Le", objFilialCliente)
                If lErro <> SUCESSO And lErro <> 12567 Then gError 76332
    
                'Se não existe a filial na tabela FiliaisClientes
                If lErro = 12567 Then
    '
    '                If lClienteAnterior <> objFilialCliente.lCodCliente Then
    '
    '                    iFilial = FILIAL_MATRIZ
    '
    '                Else
    '
    '                    iFilial = iFilial + 1
    '
    '                End If
    '
    '                objFilialCliente.iCodFilial = iFilial
                    
                    Set colEndereco = New Collection


                        
    '                If objFilialCliente.iCodFilial = FILIAL_MATRIZ Then
                    
                        'Preenche objCliente e colEndereco a partir de tImportCli
                        lErro = Cliente_PreencheObjetosImportacao(tImportCli, objCliente, colEndereco)
                        If lErro <> SUCESSO Then gError 76333
                        
                        'Grava o Cliente
                        lErro = CF("Cliente_Grava_EmTrans", objCliente, colEndereco)
                        If lErro <> SUCESSO Then gError 76334
        
                        lErro = Comando_ExecutarPos(lComando2, "UPDATE ImportCli SET codigo_corporator = ?, filial_corporator = ?", lComando, objCliente.lCodigo, FILIAL_MATRIZ)
                        If lErro <> AD_SQL_SUCESSO Then gError 76334
                        
    '                Else
    '
    '                    'Preenche objFilialCliente e colEndereco a partir de tImportCli
    '                    lErro = FilialCliente_PreencheObjetosImportacao(tImportCli, objFilialCliente, colEndereco)
    '                    If lErro <> SUCESSO Then gError 76335
    '
    '                    'Grava a Filial Cliente
    '                    lErro = CF("FiliaisClientes_Grava_EmTrans", objFilialCliente, colEndereco)
    '                    If lErro <> SUCESSO Then gError 76336
    '
    '                    lErro = Comando_ExecutarPos(lComando2, "UPDATE ImportCli SET codigo_corporator = ?, filial_corporator = ?", lComando, objFilialCliente.lCodCliente, objFilialCliente.iCodFilial)
    '                    If lErro <> AD_SQL_SUCESSO Then gError 76334
    '
    '                End If
    '
                    iQtdeCadastrada = iQtdeCadastrada + 1
                    
    '
    '            End If
    '
    '            lClienteAnterior = objFilialCliente.lCodCliente
    '
            End If
            
            'Busca o proximo registro de ImportCli
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 76341
            
            If iQtdeCadastrada > 3000 Then
                Exit Do
            End If
                    
        Loop
        
        If lErro = AD_SQL_SEM_DADOS Then bTerminou = True

        
        'Confirma a transação
        lErro = Transacao_Commit()
        If lErro <> AD_SQL_SUCESSO Then gError 76343
        
        Call Comando_Fechar(lComando)
        Call Comando_Fechar(lComando2)
        
    Loop

    Importa_Clientes = SUCESSO
    
    Exit Function
    
Erro_Importa_Clientes:

    Importa_Clientes = gErr
    
    Select Case gErr
    
        Case 76329
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 76330, 76331, 76341
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_IMPORTCLI", gErr)
        
        Case 76332, 76333, 76334, 76335, 76336
            'Erros tratados nas rotinas chamadas
            
        Case 76342
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 76343
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT_TRANSACAO", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161823)
            
    End Select
    
    Call Transacao_Rollback
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    
    Exit Function
    
End Function

Private Function Cliente_PreencheObjetosImportacao(tImportCli As typeImportCli, ByVal objCliente As ClassCliente, ByVal colEndereco As Collection) As Long
'Preenche objCliente e colEndereco a partir dos dados existentes em tImportCli

Dim lErro As Long
Dim objEndereco As New ClassEndereco
Dim objCliAUx As New ClassCliente

On Error GoTo Erro_Cliente_PreencheObjetosImportacao
    
    objCliente.lCodigo = tImportCli.lCodCliente

    objCliente.sRazaoSocial = Trim(tImportCli.sRazaoSocial)
    If objCliente.sRazaoSocial = "" Then objCliente.sRazaoSocial = "SEM NOME"
    
    objCliente.sNomeReduzido = Trim(tImportCli.sNomeReduzido)
    If objCliente.sNomeReduzido = "" Then objCliente.sNomeReduzido = left(objCliente.sRazaoSocial, STRING_CLIENTE_NOME_REDUZIDO)
    
    objCliAUx.sNomeReduzido = objCliente.sNomeReduzido

    lErro = CF("Cliente_Le_NomeReduzido", objCliAUx)
    If lErro <> SUCESSO And lErro <> 12348 Then gError ERRO_SEM_MENSAGEM

    If lErro = SUCESSO Then
        If objCliAUx.lCodigo <> objCliente.lCodigo Then
            objCliente.sNomeReduzido = left(objCliente.sNomeReduzido, STRING_CLIENTE_NOME_REDUZIDO - 7) + "-" + CStr(objCliente.lCodigo)
        End If
    End If
    
    'Todo cliente está recebendo Tipo=1
    objCliente.iTipo = tImportCli.iTipo
    
    objCliente.sObservacao = Trim(tImportCli.sObservacao)
    objCliente.dLimiteCredito = tImportCli.dLimiteCredito
    objCliente.dDesconto = tImportCli.dDesconto
    objCliente.iTabelaPreco = tImportCli.iTabelaPreco
    
    'nao incluir cnodicao de pagto
    objCliente.iCondicaoPagto = tImportCli.iCondicaoPagto
    objCliente.iAtivo = 1 '???
    objCliente.iCodMensagem = tImportCli.iCodMensagem
    objCliente.lNumPagamentos = tImportCli.lNumPagamentos
    objCliente.iCodPadraoCobranca = tImportCli.iCodPadraoCobranca
    objCliente.iProxCodFilial = 2 '???
    
    lErro = FilialCliente_PreencheInfoImportacao(objCliente, tImportCli)
    If lErro <> SUCESSO Then gError 76337
    
    lErro = FilialCliente_PreencheEnderecosImportacao(tImportCli, colEndereco)
    If lErro <> SUCESSO Then gError 76338
        
    Cliente_PreencheObjetosImportacao = SUCESSO
    
    Exit Function
    
Erro_Cliente_PreencheObjetosImportacao:

    Cliente_PreencheObjetosImportacao = gErr
    
    Select Case gErr
    
        Case 76337, 76338
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161824)
            
    End Select
    
    Exit Function
    
End Function

Private Function FilialCliente_PreencheObjetosImportacao(tImportCli As typeImportCli, ByVal objFilialCliente As ClassFilialCliente, ByVal colEndereco As Collection) As Long
'Preenche objFilialCliente e colEndereco a partir dos dados existentes em tImportCli

Dim lErro As Long
Dim objEndereco As New ClassEndereco

On Error GoTo Erro_FilialCliente_PreencheObjetosImportacao

'    objFilialCliente.iCodFilial = tImportCli.iCodFilial
    objFilialCliente.sNome = "FILIAL " & CStr(objFilialCliente.iCodFilial)
    objFilialCliente.lCodCliente = tImportCli.lCodCliente
    
    lErro = FilialCliente_PreencheInfoImportacao(objFilialCliente, tImportCli)
    If lErro <> SUCESSO Then gError 76339
    
    lErro = FilialCliente_PreencheEnderecosImportacao(tImportCli, colEndereco)
    If lErro <> SUCESSO Then gError 76340
    
    FilialCliente_PreencheObjetosImportacao = SUCESSO
    
    Exit Function
    
Erro_FilialCliente_PreencheObjetosImportacao:

    FilialCliente_PreencheObjetosImportacao = gErr
    
    Select Case gErr
    
        Case 76339, 76340
            'Erros tratados nas rotinas chamadas
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161825)
            
    End Select
    
    Exit Function
    
End Function

Private Function FilialCliente_PreencheEnderecosImportacao(tImportCli As typeImportCli, ByVal colEndereco As Collection) As Long
'Preenche colEndereco a partir dos dados existentes em tImportCli
'(0 = principal , 1 = entrega , 2 = cobranca)
Dim lErro As Long
Dim objEndereco As New ClassEndereco
Dim iPais As Integer, sUF As String

On Error GoTo Erro_FilialCliente_PreencheEnderecosImportacao

    'Preenche colEndereco
    Set objEndereco = New ClassEndereco
    
    objEndereco.iCodigoPais = tImportCli.iCodigoPais
    If objEndereco.iCodigoPais = 0 Then objEndereco.iCodigoPais = 1
    
    iPais = objEndereco.iCodigoPais
    
    objEndereco.sBairro = Trim(tImportCli.sBairro)
    objEndereco.sCEP = tImportCli.sCEP
    
    objEndereco.sCEP = Replace(objEndereco.sCEP, " ", "")
    objEndereco.sCEP = Replace(objEndereco.sCEP, ".", "")
    objEndereco.sCEP = Replace(objEndereco.sCEP, "-", "")
    objEndereco.sCEP = Replace(objEndereco.sCEP, "/", "")
    objEndereco.sCEP = Replace(objEndereco.sCEP, "\", "")
    objEndereco.sCEP = Replace(objEndereco.sCEP, "_", "")
    objEndereco.sCEP = Replace(objEndereco.sCEP, " ", "")
    
    objEndereco.sCEP = String(8 - Len(objEndereco.sCEP), "0") + objEndereco.sCEP
    If objEndereco.sCEP = String(8, "0") Then objEndereco.sCEP = ""
    
    'Verifica se o tamanho da string Cidade lida é maior que o permitido
    objEndereco.sCidade = Trim(tImportCli.sCidade)
    If Len(Trim(objEndereco.sCidade)) < 2 Then objEndereco.sCidade = ""
    
    objEndereco.sContato = Trim(tImportCli.sContato)
    objEndereco.sEmail = Trim(tImportCli.sEmail)
    objEndereco.sEndereco = Trim(tImportCli.sEndereco)
    objEndereco.sFax = Trim(tImportCli.sFax)
    objEndereco.sSiglaEstado = Trim(tImportCli.sSiglaEstado)
    
    If objEndereco.sSiglaEstado = "" Then objEndereco.sSiglaEstado = "RJ"
    
    sUF = objEndereco.sSiglaEstado
    
    'Verifica se o tamanho da string Telefone lida é maior que o permitido
    objEndereco.sTelefone1 = left(Trim(tImportCli.sTelefone1), STRING_TELEFONE)
    
    objEndereco.sTelefone2 = left(Trim(tImportCli.sTelefone2), STRING_TELEFONE)
    
    colEndereco.Add objEndereco
    
    Set objEndereco = New ClassEndereco
    
    objEndereco.iCodigoPais = tImportCli.iCodigoPais1
    If objEndereco.iCodigoPais = 0 Then objEndereco.iCodigoPais = iPais
    
    objEndereco.sBairro = Trim(tImportCli.sBairro1)
    objEndereco.sCEP = Trim(tImportCli.sCEP1)
    
    objEndereco.sCEP = Replace(objEndereco.sCEP, " ", "")
    objEndereco.sCEP = Replace(objEndereco.sCEP, ".", "")
    objEndereco.sCEP = Replace(objEndereco.sCEP, "-", "")
    objEndereco.sCEP = Replace(objEndereco.sCEP, "/", "")
    objEndereco.sCEP = Replace(objEndereco.sCEP, "\", "")
    objEndereco.sCEP = Replace(objEndereco.sCEP, "_", "")
    objEndereco.sCEP = Replace(objEndereco.sCEP, " ", "")
    
    objEndereco.sCEP = String(8 - Len(objEndereco.sCEP), "0") + objEndereco.sCEP
    If objEndereco.sCEP = String(8, "0") Then objEndereco.sCEP = ""
    
    objEndereco.sCidade = Trim(tImportCli.sCidade1)
    
    objEndereco.sContato = Trim(tImportCli.sContato1)
    objEndereco.sEmail = Trim(tImportCli.sEmail1)
    objEndereco.sEndereco = Trim(tImportCli.sEndereco1)
    objEndereco.sFax = Trim(tImportCli.sFax1)
    objEndereco.sSiglaEstado = Trim(tImportCli.sSiglaEstado1)
    
    If objEndereco.sSiglaEstado = "" Then objEndereco.sSiglaEstado = sUF
    
    objEndereco.sTelefone1 = Trim(tImportCli.sTelefone11)
    
    objEndereco.sTelefone2 = Trim(tImportCli.sTelefone21)
    
    colEndereco.Add objEndereco
    
    Set objEndereco = New ClassEndereco
    
    objEndereco.iCodigoPais = tImportCli.iCodigoPais2
    If objEndereco.iCodigoPais = 0 Then objEndereco.iCodigoPais = iPais
    
    objEndereco.sBairro = Trim(tImportCli.sBairro2)
    objEndereco.sCEP = tImportCli.sCEP2
    
    objEndereco.sCEP = Replace(objEndereco.sCEP, " ", "")
    objEndereco.sCEP = Replace(objEndereco.sCEP, ".", "")
    objEndereco.sCEP = Replace(objEndereco.sCEP, "-", "")
    objEndereco.sCEP = Replace(objEndereco.sCEP, "/", "")
    objEndereco.sCEP = Replace(objEndereco.sCEP, "\", "")
    objEndereco.sCEP = Replace(objEndereco.sCEP, "_", "")
    objEndereco.sCEP = Replace(objEndereco.sCEP, " ", "")
    
    objEndereco.sCEP = String(8 - Len(objEndereco.sCEP), "0") + objEndereco.sCEP
    If objEndereco.sCEP = String(8, "0") Then objEndereco.sCEP = ""
    
    objEndereco.sCidade = Trim(tImportCli.sCidade2)
    
    objEndereco.sContato = Trim(tImportCli.sContato2)
    objEndereco.sEmail = Trim(tImportCli.sEmail2)
    objEndereco.sEndereco = Trim(tImportCli.sEndereco2)
    objEndereco.sFax = Trim(tImportCli.sFax2)
    objEndereco.sSiglaEstado = Trim(tImportCli.sSiglaEstado2)
    
    If objEndereco.sSiglaEstado = "" Then objEndereco.sSiglaEstado = sUF
    
    objEndereco.sTelefone1 = Trim(tImportCli.sTelefone12)
    
    objEndereco.sTelefone2 = Trim(tImportCli.sTelefone22)
    
    colEndereco.Add objEndereco
    
    FilialCliente_PreencheEnderecosImportacao = SUCESSO
    
    Exit Function
    
Erro_FilialCliente_PreencheEnderecosImportacao:

    FilialCliente_PreencheEnderecosImportacao = gErr
    
    Select Case gErr
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161826)
            
    End Select
    
    Exit Function
    
End Function

Private Function FilialCliente_PreencheInfoImportacao(ByVal objFilialCliente As Object, tImportCli As typeImportCli) As Long

Dim lErro As Long
Dim objRegiaoVenda As New ClassRegiaoVenda

On Error GoTo Erro_FilialCliente_PreencheInfoImportacao

    objFilialCliente.sCgc = tImportCli.sFilialCGC
    objFilialCliente.sCgc = Replace(objFilialCliente.sCgc, " ", "")
    objFilialCliente.sCgc = Replace(objFilialCliente.sCgc, ".", "")
    objFilialCliente.sCgc = Replace(objFilialCliente.sCgc, "-", "")
    objFilialCliente.sCgc = Replace(objFilialCliente.sCgc, "/", "")
    objFilialCliente.sCgc = Replace(objFilialCliente.sCgc, "\", "")
    objFilialCliente.sCgc = Replace(objFilialCliente.sCgc, "_", "")
    objFilialCliente.sCgc = Replace(objFilialCliente.sCgc, " ", "")
'    If Len(Trim(tImportCli.sFilialCGC)) > 11 And Len(Trim(tImportCli.sFilialCGC)) <> 14 Then
'        objFilialCliente.sCgc = Format(Trim(tImportCli.sFilialCGC), "00000000000000")
'    Else
'        If Len(Trim(tImportCli.sFilialCGC)) > 8 And Len(Trim(tImportCli.sFilialCGC)) < 11 And Len(Trim(tImportCli.sFilialCGC)) <> 11 Then
'            objFilialCliente.sCgc = Format(Trim(tImportCli.sFilialCGC), "00000000000")
'        Else
'            objFilialCliente.sCgc = tImportCli.sFilialCGC
'        End If
'    End If

    If Len(objFilialCliente.sCgc) > 11 And Len(objFilialCliente.sCgc) < 14 Then
        objFilialCliente.sCgc = String(14 - Len(objFilialCliente.sCgc), "0") & objFilialCliente.sCgc
    ElseIf Len(objFilialCliente.sCgc) > 1 And Len(objFilialCliente.sCgc) < 11 Then
        objFilialCliente.sCgc = String(11 - Len(objFilialCliente.sCgc), "0") & objFilialCliente.sCgc
    End If
    
    If Replace(objFilialCliente.sCgc, "0", "") = "" Then objFilialCliente.sCgc = ""

    objFilialCliente.sInscricaoEstadual = tImportCli.sFilialInscEstadual
    objFilialCliente.sInscricaoEstadual = Replace(objFilialCliente.sInscricaoEstadual, " ", "")
    objFilialCliente.sInscricaoEstadual = Replace(objFilialCliente.sInscricaoEstadual, ".", "")
    objFilialCliente.sInscricaoEstadual = Replace(objFilialCliente.sInscricaoEstadual, "-", "")
    objFilialCliente.sInscricaoEstadual = Replace(objFilialCliente.sInscricaoEstadual, "/", "")
    objFilialCliente.sInscricaoEstadual = Replace(objFilialCliente.sInscricaoEstadual, "\", "")
    objFilialCliente.sInscricaoEstadual = Replace(objFilialCliente.sInscricaoEstadual, "_", "")
    objFilialCliente.sInscricaoEstadual = Replace(objFilialCliente.sInscricaoEstadual, "ISENTO", "")
    objFilialCliente.sInscricaoEstadual = Replace(objFilialCliente.sInscricaoEstadual, "ISENTA", "")
    If Replace(objFilialCliente.sInscricaoEstadual, "0", "") = "" Then objFilialCliente.sInscricaoEstadual = ""
    
    If Len(Trim(objFilialCliente.sInscricaoEstadual)) > 0 Then
        objFilialCliente.iIEIsento = DESMARCADO
        objFilialCliente.iIENaoContrib = DESMARCADO
    Else
        objFilialCliente.iIEIsento = MARCADO
        objFilialCliente.iIENaoContrib = MARCADO
    End If
        
    objFilialCliente.iRegimeTributario = REGIME_TRIBUTARIO_NORMAL
    objFilialCliente.sInscricaoMunicipal = tImportCli.sFilialInscMunicipal
    
    objFilialCliente.sInscricaoMunicipal = Replace(objFilialCliente.sInscricaoMunicipal, " ", "")
    objFilialCliente.sInscricaoMunicipal = Replace(objFilialCliente.sInscricaoMunicipal, ".", "")
    objFilialCliente.sInscricaoMunicipal = Replace(objFilialCliente.sInscricaoMunicipal, "-", "")
    objFilialCliente.sInscricaoMunicipal = Replace(objFilialCliente.sInscricaoMunicipal, "/", "")
    objFilialCliente.sInscricaoMunicipal = Replace(objFilialCliente.sInscricaoMunicipal, "\", "")
    objFilialCliente.sInscricaoMunicipal = Replace(objFilialCliente.sInscricaoMunicipal, "_", "")
    objFilialCliente.sInscricaoMunicipal = Replace(objFilialCliente.sInscricaoMunicipal, "ISENTO", "")
    objFilialCliente.sInscricaoMunicipal = Replace(objFilialCliente.sInscricaoMunicipal, "ISENTA", "")

    'objFilialCliente.sObservacao = tImportCli.sObservacao
    objFilialCliente.iVendedor = tImportCli.iFilialVendedor
    objFilialCliente.iRegiao = tImportCli.iFilialRegiao
    objFilialCliente.sContaContabil = tImportCli.sFilialContaContabil
    objFilialCliente.dComissaoVendas = tImportCli.dFilialComissaoVendas
    objFilialCliente.iCodCobrador = tImportCli.iFilialCodCobrador
    objFilialCliente.iFreqVisitas = tImportCli.iFilialFreqVisitas
    objFilialCliente.dtDataUltVisita = tImportCli.dtFilialDataUltVisita
    objFilialCliente.iCodTransportadora = tImportCli.iFilialCodTransportadora
    
    If tImportCli.sFilialTipoFrete = "F" Then
        objFilialCliente.iTipoFrete = TIPO_FOB
    ElseIf tImportCli.sFilialTipoFrete = "C" Then
        objFilialCliente.iTipoFrete = TIPO_CIF
    End If
    
    FilialCliente_PreencheInfoImportacao = SUCESSO
     
    Exit Function
    
Erro_FilialCliente_PreencheInfoImportacao:

    FilialCliente_PreencheInfoImportacao = gErr
     
    Select Case Err
          
        Case 76339, 76340, 76385
            'Erros tratados nas rotinas chamadas
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161827)
     
    End Select
     
    Exit Function

End Function

Function Importa_TitRec()

Dim lErro As Long, iIndice As Integer, lTransacao As Long
Dim alComando(1 To 5) As Long, lAux As Long, iAux As Integer
Dim sSiglaDocumento As String, lNumTitulo As Long, sParcela As String, iParcela As Integer
Dim sNatureza As String, dtDataEmissao As Date, dtDataVencimento As Date, dtDataVenctoReal As Date
Dim dtDataBaixa As Date, lCliente As Long, iFilial As Integer, dValor As Double, dSaldo As Double
Dim sObservacao As String, sNossoNumero As String, sID As String
Dim sSiglaAnt As String, lNumAnt As Long
Dim objTitRec As ClassTituloReceber
Dim objParcRec As ClassParcelaReceber
Dim colParcelaReceber As colParcelaReceber
Dim colComissaoEmissao As colComissao
Dim colcolComissao As colcolComissao
Dim colComissao As colComissao
Dim colcolDesconto As colcolDesconto
Dim colDesconto As colDesconto
Dim objContabil As ClassContabil
Dim colBaixaParcReceber As ColBaixaParcRec
Dim objBaixaReceber As ClassBaixaReceber
Dim colID As Collection, vValor As Variant

On Error GoTo Erro_Importa_TitRec

    'Abre a transação
    lTransacao = Transacao_Abrir
    If lTransacao = 0 Then gError 211850

    'Abrir comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir
        If alComando(iIndice) = 0 Then gError 211851
    Next

    sSiglaDocumento = String(255, 0)
    sParcela = String(255, 0)
    sNatureza = String(255, 0)
    sObservacao = String(255, 0)
    sNossoNumero = String(255, 0)
    sID = String(255, 0)
    
    lErro = Comando_Executar(alComando(1), "SELECT SiglaDocumento,NumTitulo,Parcela,Natureza, DataEmissao,DataVencimento,DataVenctoReal,DataBaixa,Cliente,Filial,Valor,Saldo,Observacao,NossoNumero,ID FROM ImportTitRec WHERE NumIntTitulo_Corporator = 0 ORDER BY SiglaDocumento, NumTitulo, Parcela", _
    sSiglaDocumento, lNumTitulo, sParcela, sNatureza, dtDataEmissao, dtDataVencimento, dtDataVenctoReal, dtDataBaixa, lCliente, iFilial, dValor, dSaldo, sObservacao, sNossoNumero, sID)
    If lErro <> AD_SQL_SUCESSO Then gError 211852
    
    lErro = Comando_BuscarProximo(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211853
    
    iIndice = 1
    
    Do While lErro = AD_SQL_SUCESSO
    
        If sSiglaAnt <> sSiglaDocumento Or lNumAnt <> lNumTitulo Then
                               
            If lNumAnt <> 0 Then
                lErro = CF("TituloReceber_Grava_EmTrans", objTitRec, colComissaoEmissao, colParcelaReceber, colcolComissao, colcolDesconto, objContabil)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                For Each objParcRec In colParcelaReceber
                    
                    If Abs(objParcRec.dValor - objParcRec.dValorAux) > DELTA_VALORMONETARIO Then
                    
                        Set colBaixaParcReceber = New ColBaixaParcRec
                        Set objBaixaReceber = New ClassBaixaReceber
                        
                        objBaixaReceber.iStatus = 1
                        objBaixaReceber.iMotivo = 4
                        objBaixaReceber.dtData = dtDataBaixa
                        objBaixaReceber.dtDataContabil = dtDataBaixa
                        objBaixaReceber.dtDataRegistro = gdtDataHoje
                        objBaixaReceber.sHistorico = "Importação de Título Baixado da Microsiga"

                        colBaixaParcReceber.Add 0, 0, objParcRec.lNumIntDoc, 0, 1, 0, 0, 0, objParcRec.dValor - objParcRec.dValorAux, objParcRec.dValor - objParcRec.dValorAux, 0
    
                        'Grava BaixaReceber e BaixasReceberParcela associadas
                        lErro = CF("BaixaReceber_Grava", objBaixaReceber, colBaixaParcReceber, objContabil, objTitRec.lCliente, objTitRec.iFilial)
                        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    
                    End If
                
                Next
                
                For Each vValor In colID
                    lErro = Comando_Executar(alComando(5), "UPDATE ImportTitRec SET NumIntTitulo_Corporator = ? WHERE ID = ? ", objTitRec.lNumIntDoc, vValor)
                    If lErro <> AD_SQL_SUCESSO Then gError 211852
                Next
                
            End If

            'preencher o objeto e mandar gravar
            Set objTitRec = New ClassTituloReceber
            Set colID = New Collection
        
            objTitRec.lCliente = lCliente
            objTitRec.iFilial = iFilial
            objTitRec.lNumTitulo = lNumTitulo
            objTitRec.dtDataEmissao = dtDataEmissao
            objTitRec.iFilialEmpresa = giFilialEmpresa
            objTitRec.sSiglaDocumento = sSiglaDocumento
            objTitRec.sNatureza = sNatureza
            objTitRec.dtDataRegistro = DATA_NULA
            objTitRec.dtDataEstorno = DATA_NULA
            objTitRec.sObservacao = left(Trim(sObservacao), 50)
            
            lErro = Comando_Executar(alComando(2), "SELECT Codigo FROM Clientes WHERE Codigo = ?", lAux, lCliente)
            If lErro <> AD_SQL_SUCESSO Then gError 211857
            
            lErro = Comando_BuscarProximo(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211858
            
            If lErro <> SUCESSO Then gError 211859

            lErro = Comando_Executar(alComando(3), "SELECT ClasseDocCPR FROM TiposDeDocumento WHERE Sigla = ?", iAux, sSiglaDocumento)
            If lErro <> AD_SQL_SUCESSO Then gError 211860
            
            lErro = Comando_BuscarProximo(alComando(3))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211861
            
            If lErro <> SUCESSO Then gError 211862
            
            lErro = Comando_Executar(alComando(4), "SELECT Tipo FROM NatMovCta WHERE Codigo = ?", iAux, sNatureza)
            If lErro <> AD_SQL_SUCESSO Then gError 211863
            
            lErro = Comando_BuscarProximo(alComando(4))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211864
            
            If lErro <> SUCESSO Then gError 211865
            
            Set colParcelaReceber = New colParcelaReceber
            Set colComissaoEmissao = New colComissao
            Set colcolComissao = New colcolComissao
            Set colcolDesconto = New colcolDesconto
            
            sSiglaAnt = sSiglaDocumento
            lNumAnt = lNumTitulo
        
        End If
        Set objParcRec = New ClassParcelaReceber
        Set colComissao = New colComissao
        Set colDesconto = New colDesconto
    
        colParcelaReceber.AddObj objParcRec
        colcolDesconto.Add colDesconto
        colcolComissao.Add colComissao
        
        colID.Add sID
    
        objTitRec.iNumParcelas = objTitRec.iNumParcelas + 1
        objParcRec.iNumParcela = colParcelaReceber.Count
        objParcRec.dtDataVencimento = dtDataVencimento
        objParcRec.dtDataVencimentoReal = dtDataVenctoReal
        objParcRec.dValor = dValor
        objParcRec.dSaldo = dValor
        objParcRec.dValorOriginal = dValor
        objParcRec.dValorAux = dSaldo
        
        objTitRec.dValor = objTitRec.dValor + dValor
        objTitRec.dSaldo = objTitRec.dValor
        objParcRec.sNumTitCobrador = sNossoNumero
        objParcRec.sObservacao = sObservacao
        
        objParcRec.dtDataCredito = DATA_NULA
        objParcRec.dtDataDepositoCheque = DATA_NULA
        objParcRec.dtDataEmissaoCheque = DATA_NULA
        objParcRec.dtDataPrevReceb = DATA_NULA
        objParcRec.dtDataProxCobr = DATA_NULA
        objParcRec.dtDataTransacaoCartao = DATA_NULA
        objParcRec.dtDesconto1Ate = DATA_NULA
        objParcRec.dtDesconto2Ate = DATA_NULA
        objParcRec.dtDesconto3Ate = DATA_NULA
        objParcRec.dtValidadeCartao = DATA_NULA
        
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211854
    
    Loop
    

     If lNumAnt <> 0 Then
         lErro = CF("TituloReceber_Grava_EmTrans", objTitRec, colComissaoEmissao, colParcelaReceber, colcolComissao, colcolDesconto, objContabil)
         If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
         
         For Each objParcRec In colParcelaReceber
             
             If Abs(objParcRec.dValor - objParcRec.dValorAux) > DELTA_VALORMONETARIO Then
             
                 Set colBaixaParcReceber = New ColBaixaParcRec
                 Set objBaixaReceber = New ClassBaixaReceber
                 
                 objBaixaReceber.iStatus = 1
                 objBaixaReceber.iMotivo = 4
                 objBaixaReceber.dtData = dtDataBaixa
                 objBaixaReceber.dtDataContabil = dtDataBaixa
                 objBaixaReceber.dtDataRegistro = gdtDataHoje
                 objBaixaReceber.sHistorico = "Importação de Título Baixado da Microsiga"

                 colBaixaParcReceber.Add 0, 0, objParcRec.lNumIntDoc, 0, 1, 0, 0, 0, objParcRec.dValor - objParcRec.dValorAux, objParcRec.dValor - objParcRec.dValorAux, 0

                 'Grava BaixaReceber e BaixasReceberParcela associadas
                 lErro = CF("BaixaReceber_Grava", objBaixaReceber, colBaixaParcReceber, objContabil, objTitRec.lCliente, objTitRec.iFilial)
                 If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
             
             End If
         
         Next
         
     End If
    
    'Libera comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Fecha a Transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 211855

    Importa_TitRec = SUCESSO
     
    Exit Function
    
Erro_Importa_TitRec:

    Importa_TitRec = gErr
     
    Select Case gErr
          
        Case ERRO_SEM_MENSAGEM
        
        Case 211850
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 211851
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 211852 To 211854
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_IMPORTTITREC", gErr)
        
        Case 211855
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT_TRANSACAO", gErr)
            
        Case 211857, 211858
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLIENTES", gErr)
          
        Case 211859
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, lCliente)
          
        Case 211860, 211861
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOSDEDOCUMENTO", gErr)
          
        Case 211862
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_DOCUMENTO_NAO_CADASTRADO", gErr, sSiglaDocumento)
          
        Case 211863, 211864
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NATMOVCTA", gErr)
          
        Case 211865
            Call Rotina_Erro(vbOKOnly, "ERRO_NATMOVCTA_NAO_CADASTRADA", gErr, sNatureza)
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211856)
     
    End Select
     
    'Libera comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Call Transacao_Rollback
    
    Exit Function

End Function

Function Importa_Kits()

Dim lErro As Long, iIndice As Integer, lTransacao As Long
Dim alComando(1 To 5) As Long, lAux As Long, iAux As Integer
Dim objKit As ClassKit
Dim objProdutoKit As ClassProdutoKit
Dim sProdutoRaiz As String, dQtdeRaiz As Double, sUMRaiz As String
Dim sProduto As String, dtData As Date, sVersao As String, dQtde As Double
Dim sUM As String, iComposicao As Integer, dPerda As Double, sObservacao As String
Dim sProdRaizAnt As String, sVersaoAnt As String
Dim objProduto As ClassProduto, iSeq As Integer

On Error GoTo Erro_Importa_Kits

    'Abre a transação
    lTransacao = Transacao_Abrir
    If lTransacao = 0 Then gError 211850

    'Abrir comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir
        If alComando(iIndice) = 0 Then gError 211851
    Next
    
    sProdutoRaiz = String(255, 0)
    sUMRaiz = String(255, 0)
    sProduto = String(255, 0)
    sVersao = String(255, 0)
    sUM = String(255, 0)
    sObservacao = String(255, 0)
   
    lErro = Comando_Executar(alComando(1), "SELECT ProdutoRaiz,QtdeRaiz,UMRaiz,Produto,Data,Versao,Qtde,UM,Composicao,Perda,Observacao FROM ImportKits ORDER BY ProdutoRaiz, Versao, Seq", _
    sProdutoRaiz, dQtdeRaiz, sUMRaiz, sProduto, dtData, sVersao, dQtde, sUM, iComposicao, dPerda, sObservacao)
    If lErro <> AD_SQL_SUCESSO Then gError 211852
    
    lErro = Comando_BuscarProximo(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211853
    
    iIndice = 1
    
    Do While lErro = AD_SQL_SUCESSO
    
        iSeq = iSeq + 1
    
        If sProdRaizAnt <> sProdutoRaiz Or sVersaoAnt <> sVersao Then
    
            If sProdRaizAnt <> "" Then
    
'                lErro = CF("Kit_Valida_Quantidade", objKit)
'                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                'Verifica recursividade circular, ou seja, o produto final não pode estar contido como insumo de um dos componentes, direta ou indiretamente desde que e igual ou maior quantidade
'                'ntre outros
'                lErro = CF("Kit_Valida", objKit)
'                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                lErro = CF("Kit_Grava_EmTrans", objKit)
                If lErro <> SUCESSO And lErro <> 21779 And lErro <> 21780 Then gError ERRO_SEM_MENSAGEM
                
            End If
            sProdRaizAnt = sProdutoRaiz
            sVersaoAnt = sVersao
            iSeq = 2
            
            Set objKit = New ClassKit
    
            If dtData <> DATA_NULA Then
                objKit.dtData = dtData
            Else
            
            End If
            objKit.iSituacao = KIT_SITUACAO_PADRAO
            
            objKit.iVersaoFormPreco = objKit.iSituacao
            objKit.sObservacao = sObservacao
            objKit.sProdutoRaiz = sProdutoRaiz
            
            If sVersao <> "" Then
                objKit.sVersao = sVersao
            Else
                objKit.sVersao = "1.00"
            End If
            
            Set objProduto = New ClassProduto
            objProduto.sCodigo = objKit.sProdutoRaiz
            
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
            
            If lErro <> SUCESSO Then gError 211857
            
            Set objProdutoKit = New ClassProdutoKit
            objKit.colComponentes.Add objProdutoKit
            
            If dQtdeRaiz <> 0 Then
                objProdutoKit.dQuantidade = dQtdeRaiz
            Else
                objProdutoKit.dQuantidade = 1
            End If
            objProdutoKit.iNivel = KIT_NIVEL_RAIZ
            objProdutoKit.sProduto = objKit.sProdutoRaiz
            objProdutoKit.sProdutoRaiz = objKit.sProdutoRaiz
            objProdutoKit.sVersao = objKit.sVersao
            If sUMRaiz <> "" Then
                objProdutoKit.sUnidadeMed = sUMRaiz
            Else
                objProdutoKit.sUnidadeMed = objProduto.sSiglaUMEstoque
            End If
            objProdutoKit.iSeq = 1
            objProdutoKit.iSeqPai = 0
            objProdutoKit.dPercentualPerda = 0
            objProdutoKit.iPosicaoArvore = 0
            objProdutoKit.iComposicao = 1
            
        End If
        
        Set objProduto = New ClassProduto
        objProduto.sCodigo = sProduto
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
        
        If lErro <> SUCESSO Then gError 211858
        
        Set objProdutoKit = New ClassProdutoKit
        objKit.colComponentes.Add objProdutoKit
        
        If dQtde <> 0 Then
            objProdutoKit.dQuantidade = dQtde
        Else
            objProdutoKit.dQuantidade = 1
        End If
        objProdutoKit.iNivel = 1
        objProdutoKit.sProduto = sProduto
        objProdutoKit.sProdutoRaiz = objKit.sProdutoRaiz
        objProdutoKit.sVersao = objKit.sVersao
        If sUM <> "" Then
            objProdutoKit.sUnidadeMed = sUM
        Else
            objProdutoKit.sUnidadeMed = objProduto.sSiglaUMEstoque
        End If
        objProdutoKit.iSeq = iSeq
        objProdutoKit.iSeqPai = 1
        objProdutoKit.dPercentualPerda = dPerda
        objProdutoKit.iPosicaoArvore = iSeq - 1
        objProdutoKit.iComposicao = 1
        
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211854
    
    Loop
    
    If sProdRaizAnt <> "" Then

'        lErro = CF("Kit_Valida_Quantidade", objKit)
'        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
'
'        'Verifica recursividade circular, ou seja, o produto final não pode estar contido como insumo de um dos componentes, direta ou indiretamente desde que e igual ou maior quantidade
'        'ntre outros
'        lErro = CF("Kit_Valida", objKit)
'        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        lErro = CF("Kit_Grava_EmTrans", objKit)
        If lErro <> SUCESSO And lErro <> 21779 And lErro <> 21780 Then gError ERRO_SEM_MENSAGEM
        
    End If
    
    'Libera comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Fecha a Transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 211855

    Importa_Kits = SUCESSO
     
    Exit Function
    
Erro_Importa_Kits:

    Importa_Kits = gErr
     
    Select Case gErr
          
        Case ERRO_SEM_MENSAGEM
        
        Case 211850
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 211851
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 211852 To 211854
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_IMPORTKITS", gErr)
       
        Case 211855
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT_TRANSACAO", gErr)

        Case 211857, 211858
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211856)
     
    End Select
     
    'Libera comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Call Transacao_Rollback
    
    Exit Function

End Function

Function Importa_TitPag()

Dim lErro As Long, iIndice As Integer, lTransacao As Long
Dim alComando(1 To 5) As Long, lAux As Long, iAux As Integer
Dim sSiglaDocumento As String, sNumTitulo As String, sParcela As String, iParcela As Integer
Dim sNatureza As String, dtDataEmissao As Date, dtDataVencimento As Date, dtDataVenctoReal As Date
Dim dtDataBaixa As Date, lFornecedor As Long, iFilial As Integer, dValor As Double, dSaldo As Double
Dim sObservacao As String, sID As String
Dim sSiglaAnt As String, sNumAnt As String, lFornAnt As Long
Dim objTitPag As ClassTituloPagar
Dim objParcPag As ClassParcelaPagar
Dim colParcelaPagar As colParcelaPagar
Dim objContabil As ClassContabil
Dim colBaixaParcPagar As ColBaixaParcRec
Dim objBaixaPagar As ClassBaixaPagar
Dim colID As Collection, vValor As Variant, lFat As Long

On Error GoTo Erro_Importa_TitPag

    'Abre a transação
    lTransacao = Transacao_Abrir
    If lTransacao = 0 Then gError 211850

    'Abrir comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir
        If alComando(iIndice) = 0 Then gError 211851
    Next

    sSiglaDocumento = String(255, 0)
    sParcela = String(255, 0)
    sNatureza = String(255, 0)
    sObservacao = String(255, 0)
    sID = String(255, 0)
    sNumTitulo = String(255, 0)
    
    lFat = 900000
    
    lErro = Comando_Executar(alComando(1), "SELECT SiglaDocumento,NumTitulo,Parcela,Natureza, DataEmissao,DataVencimento,DataVenctoReal,DataBaixa,Fornecedor,Filial,Valor,Saldo,Observacao,ID FROM ImportTitPag WHERE NumIntTitulo_Corporator = 0 ORDER BY SiglaDocumento, NumTitulo, Fornecedor, ID", _
    sSiglaDocumento, sNumTitulo, sParcela, sNatureza, dtDataEmissao, dtDataVencimento, dtDataVenctoReal, dtDataBaixa, lFornecedor, iFilial, dValor, dSaldo, sObservacao, sID)
    If lErro <> AD_SQL_SUCESSO Then gError 211852
    
    lErro = Comando_BuscarProximo(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211853
    
    iIndice = 1
    
    Do While lErro = AD_SQL_SUCESSO
    
        iIndice = 1 + iIndice
    
        If sSiglaAnt <> sSiglaDocumento Or sNumAnt <> sNumTitulo Or lFornAnt <> lFornecedor Then
                               
            If sNumAnt <> "" Then
            
                lErro = CF("NFFatPag_Grava_EmTrans", objTitPag, colParcelaPagar, objContabil)
                If lErro <> SUCESSO Then
                    gError ERRO_SEM_MENSAGEM
                End If
                
                For Each vValor In colID
                    lErro = Comando_Executar(alComando(5), "UPDATE ImportTitPag SET NumIntTitulo_Corporator = ? WHERE ID = ? ", objTitPag.lNumIntDoc, vValor)
                    If lErro <> AD_SQL_SUCESSO Then gError 211852
                Next
                
            End If

            'preencher o objeto e mandar gravar
            Set objTitPag = New ClassTituloPagar
            Set colID = New Collection
        
            objTitPag.sObservacao = sObservacao
        
            objTitPag.lFornecedor = lFornecedor
            objTitPag.iFilial = iFilial
            
            If IsNumeric(sNumTitulo) Then
                objTitPag.lNumTitulo = StrParaLong(sNumTitulo)
            ElseIf IsNumeric(sID) Then
                objTitPag.lNumTitulo = StrParaLong(sID)
            Else
                lFat = lFat + 1
                objTitPag.lNumTitulo = lFat
            End If
            objTitPag.dtDataEmissao = dtDataEmissao
            objTitPag.iFilialEmpresa = giFilialEmpresa
            objTitPag.sSiglaDocumento = sSiglaDocumento
            objTitPag.sNatureza = sNatureza
            objTitPag.dtDataRegistro = DATA_NULA
            objTitPag.dtDataEstorno = DATA_NULA
            
            lErro = Comando_Executar(alComando(2), "SELECT Codigo FROM Fornecedores WHERE Codigo = ?", lAux, lFornecedor)
            If lErro <> AD_SQL_SUCESSO Then gError 211857
            
            lErro = Comando_BuscarProximo(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211858
            
            If lErro <> SUCESSO Then gError 211859

            lErro = Comando_Executar(alComando(3), "SELECT ClasseDocCPR FROM TiposDeDocumento WHERE Sigla = ?", iAux, sSiglaDocumento)
            If lErro <> AD_SQL_SUCESSO Then gError 211860
            
            lErro = Comando_BuscarProximo(alComando(3))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211861
            
            If lErro <> SUCESSO Then gError 211862
            
            lErro = Comando_Executar(alComando(4), "SELECT Tipo FROM NatMovCta WHERE Codigo = ?", iAux, sNatureza)
            If lErro <> AD_SQL_SUCESSO Then gError 211863
            
            lErro = Comando_BuscarProximo(alComando(4))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211864
            
            If lErro <> SUCESSO Then gError 211865
            
            Set colParcelaPagar = New colParcelaPagar
            
            sSiglaAnt = sSiglaDocumento
            sNumAnt = sNumTitulo
            lFornAnt = lFornecedor
        
        End If
        
        Set objParcPag = colParcelaPagar.Add(0, 0, 0, 0, dtDataVencimento, dtDataVenctoReal, dSaldo, dValor, 0, 0, 0, 0, "", "")
    
        colID.Add sID
    
        objTitPag.iNumParcelas = objTitPag.iNumParcelas + 1
        objParcPag.iNumParcela = colParcelaPagar.Count
        objParcPag.dtDataVencimento = dtDataVencimento
        objParcPag.dtDataVencimentoReal = dtDataVenctoReal
        objParcPag.dValor = dValor
        objParcPag.dSaldo = dValor
        objParcPag.dValorOriginal = dValor
        objParcPag.iTipoCobranca = 1
        
        objTitPag.dValorTotal = objTitPag.dValorTotal + dValor
        objTitPag.dValorProdutos = objTitPag.dValorTotal
        objTitPag.dSaldo = objTitPag.dValorTotal
                
        objParcPag.dtDataUltimaBaixa = DATA_NULA
        
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211854
    
    Loop
    
    If sNumAnt <> 0 Then
        lErro = CF("NFFatPag_Grava_EmTrans", objTitPag, colParcelaPagar, objContabil)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        For Each vValor In colID
            lErro = Comando_Executar(alComando(5), "UPDATE ImportTitPag SET NumIntTitulo_Corporator = ? WHERE ID = ? ", objTitPag.lNumIntDoc, vValor)
            If lErro <> AD_SQL_SUCESSO Then gError 211852
        Next
        
    End If
    
    'Libera comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Fecha a Transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 211855

    Importa_TitPag = SUCESSO
     
    Exit Function
    
Erro_Importa_TitPag:

    Importa_TitPag = gErr
     
    Select Case gErr
          
        Case ERRO_SEM_MENSAGEM
        
        Case 211850
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 211851
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 211852 To 211854
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_IMPORTTITREC", gErr)
        
        Case 211855
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT_TRANSACAO", gErr)
            
        Case 211857, 211858
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FornecedorS", gErr)
          
        Case 211859
            Call Rotina_Erro(vbOKOnly, "ERRO_Fornecedor_NAO_CADASTRADO", gErr, lFornecedor)
          
        Case 211860, 211861
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOSDEDOCUMENTO", gErr)
          
        Case 211862
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_DOCUMENTO_NAO_CADASTRADO", gErr, sSiglaDocumento)
          
        Case 211863, 211864
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NATMOVCTA", gErr)
          
        Case 211865
            Call Rotina_Erro(vbOKOnly, "ERRO_NATMOVCTA_NAO_CADASTRADA", gErr, sNatureza)
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211856)
     
    End Select
     
    'Libera comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Call Transacao_Rollback
    
    Exit Function

End Function

Function Importa_CondPagto() As Long

Dim lErro As Long
Dim lTransacao As Long
Dim alComando(1 To 5) As Long
Dim iIndice As Integer
Dim lCodAnt As Long, sDescricao As String, sParcelas As String, iTipo As Integer
Dim objCondicaoPagto As ClassCondicaoPagto
Dim objParc As ClassCondicaoPagtoParc, objParcAux As ClassCondicaoPagtoParc
Dim colParcs As Collection, colParcsOrd As Collection, colCampos As Collection, vValor As Variant
Dim iPos As Integer, sAux As String, dPercFalta As Double, iProxCodigo As Integer
Dim iCodAux As Integer

On Error GoTo Erro_Importa_CondPagto
    
    'abrir transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 205034

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 205035
    Next
    
    iProxCodigo = 1
    
    sDescricao = String(STRING_MAXIMO, 0)
    sParcelas = String(STRING_MAXIMO, 0)

    lErro = Comando_ExecutarPos(alComando(1), "SELECT CodAnt, Descricao, Parcelas, Tipo FROM ImportCondPagto WHERE CodCorporator = 0 ", 0, lCodAnt, sDescricao, sParcelas, iTipo)
    If lErro <> AD_SQL_SUCESSO Then gError 211852
    
    lErro = Comando_BuscarProximo(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211853

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objCondicaoPagto = New ClassCondicaoPagto
        Set colParcs = New Collection
        Set colParcsOrd = New Collection
        Set colCampos = New Collection
        
        iProxCodigo = iProxCodigo + 1
        
        objCondicaoPagto.iCodigo = iProxCodigo
        objCondicaoPagto.sDescricao = Trim(sDescricao)
        objCondicaoPagto.sDescReduzida = Trim(sDescricao)
        objCondicaoPagto.iEmRecebimento = MARCADO
        objCondicaoPagto.iEmPagamento = MARCADO
    
        lErro = Comando_Executar(alComando(5), "SELECT Codigo FROM CondicoesPagto WHERE DescReduzida = ? ", iCodAux, sDescricao)
        If lErro <> AD_SQL_SUCESSO Then gError 211852
        
        lErro = Comando_BuscarProximo(alComando(5))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211853
        
        If lErro = AD_SQL_SUCESSO Then
            objCondicaoPagto.sDescReduzida = objCondicaoPagto.sDescReduzida & "(" & CStr(lCodAnt) & ")"
        End If
            
        If iTipo = 2 Then
            sAux = Trim(sDescricao)
            iTipo = 1
        ElseIf iTipo = 9 Then
            sAux = "1"
            iTipo = 1
        Else
            sAux = Trim(sParcelas)
        End If
        
        Do While sAux <> ""
            Set objParc = New ClassCondicaoPagtoParc
            iPos = InStr(1, sAux, ",")
            If iPos = 0 Then
                objParc.iDias = StrParaInt(sAux)
                colParcs.Add objParc
                sAux = ""
            Else
                objParc.iDias = StrParaInt(left(sAux, iPos - 1))
                colParcs.Add objParc
                sAux = Trim(Mid(sAux, iPos + 1))
            End If
        Loop
        
        colCampos.Add "iDias"
        
        Call Ordena_Colecao(colParcs, colParcsOrd, colCampos)
        
        Select Case iTipo
        
            Case 1 'Dias entre parcelas
                iIndice = 0
                dPercFalta = 1
                For Each objParcAux In colParcsOrd
                    iIndice = iIndice + 1
                    If iIndice = 1 Then
                        objCondicaoPagto.iDiasParaPrimeiraParcela = objParcAux.iDias
                        objCondicaoPagto.iNumeroParcelas = colParcs.Count
                    End If
                    Set objParc = New ClassCondicaoPagtoParc
                    
                    objParc.iCodigo = objCondicaoPagto.iCodigo
                    objParc.iSeq = iIndice
                    objParc.iTipoDataBase = CONDPAGTO_TIPODATABASE_EMISSAO
                    objParc.iTipoIntervalo = CONDPAGTO_TIPOINTERVALO_DIAS
                    objParc.iDias = objParcAux.iDias
                    objParc.iModificador = CONDPAGTO_MODIFICADOR_VAZIO
                    
                    If iIndice = colParcs.Count Then
                        objParc.dPercReceb = dPercFalta
                    Else
                        objParc.dPercReceb = Round(1 / colParcs.Count, 8)
                        dPercFalta = dPercFalta - objParc.dPercReceb
                    End If
                    
                    objCondicaoPagto.colParcelas.Add objParc
                Next

            Case 3 'xVezes dia fixo
            
                objCondicaoPagto.iDataFixa = MARCADO
                objCondicaoPagto.iMensal = MARCADO
                dPercFalta = 1
                Select Case colParcs.Count
                
                    Case 3
                    
                        objCondicaoPagto.iNumeroParcelas = colParcs.Item(1).iDias
                        objCondicaoPagto.iDiaDoMes = colParcs.Item(3).iDias
                        
                        For iIndice = 1 To objCondicaoPagto.iNumeroParcelas
                        
                            Set objParc = New ClassCondicaoPagtoParc
                            
                            objParc.iCodigo = objCondicaoPagto.iCodigo
                            objParc.iSeq = iIndice
                            objParc.iTipoDataBase = CONDPAGTO_TIPODATABASE_DATAFIXA
                            objParc.iTipoIntervalo = CONDPAGTO_TIPOINTERVALO_DIAMES
                            objParc.iDias = objCondicaoPagto.iDiaDoMes
                            objParc.iModificador = CONDPAGTO_MODIFICADOR_VAZIO
                            
                            If iIndice = objCondicaoPagto.iNumeroParcelas Then
                                objParc.dPercReceb = dPercFalta
                            Else
                                objParc.dPercReceb = Round(1 / colParcs.Count, 8)
                                dPercFalta = dPercFalta - objParc.dPercReceb
                            End If
                            
                            objCondicaoPagto.colParcelas.Add objParc
                            
                        Next
                
                    
                    Case 4
                    
                        objCondicaoPagto.iNumeroParcelas = colParcs.Item(1).iDias
                        objCondicaoPagto.iDiaDoMes = colParcs.Item(3).iDias
                        
                        For iIndice = 1 To objCondicaoPagto.iNumeroParcelas
                        
                            Set objParc = New ClassCondicaoPagtoParc
                            
                            objParc.iCodigo = objCondicaoPagto.iCodigo
                            objParc.iSeq = iIndice
                            objParc.iTipoDataBase = CONDPAGTO_TIPODATABASE_DATAFIXA
                            objParc.iTipoIntervalo = CONDPAGTO_TIPOINTERVALO_DIAMES
                            
                            If iIndice Mod 2 = 0 Then
                                objParc.iDias = colParcs.Item(4).iDias
                            Else
                                objParc.iDias = objCondicaoPagto.iDiaDoMes
                            End If
                            
                            objParc.iModificador = CONDPAGTO_MODIFICADOR_VAZIO
                            
                            If iIndice = objCondicaoPagto.iNumeroParcelas Then
                                objParc.dPercReceb = dPercFalta
                            Else
                                objParc.dPercReceb = Round(1 / colParcs.Count, 8)
                                dPercFalta = dPercFalta - objParc.dPercReceb
                            End If
                            
                            objCondicaoPagto.colParcelas.Add objParc
                            
                        Next
                    
                    Case Else
                        gError 999999
                
                End Select
                
            Case 7 'xVezes dia fixo

                objCondicaoPagto.iDataFixa = MARCADO
                objCondicaoPagto.iMensal = MARCADO
                dPercFalta = 1
                    
                objCondicaoPagto.iNumeroParcelas = colParcs.Item(1).iDias
                objCondicaoPagto.iDiaDoMes = colParcs.Item(2).iDias
                
                For iIndice = 1 To objCondicaoPagto.iNumeroParcelas
                
                    Set objParc = New ClassCondicaoPagtoParc
                    
                    objParc.iCodigo = objCondicaoPagto.iCodigo
                    objParc.iSeq = iIndice
                    objParc.iTipoDataBase = CONDPAGTO_TIPODATABASE_DATAFIXA
                    objParc.iTipoIntervalo = CONDPAGTO_TIPOINTERVALO_DIAMES
                    objParc.iDias = objCondicaoPagto.iDiaDoMes
                    objParc.iModificador = CONDPAGTO_MODIFICADOR_VAZIO
                    
                    If iIndice = objCondicaoPagto.iNumeroParcelas Then
                        objParc.dPercReceb = dPercFalta
                    Else
                        objParc.dPercReceb = Round(1 / colParcs.Count, 8)
                        dPercFalta = dPercFalta - objParc.dPercReceb
                    End If
                    
                    objCondicaoPagto.colParcelas.Add objParc

                Next
                
            Case Else
                gError 99999
                
        End Select
    
        lErro = Comando_Executar(alComando(3), "INSERT INTO CondicoesPagto (Codigo, DescReduzida, Descricao, EmPagamento, EmRecebimento, NumeroParcelas, DiasParaPrimeiraParcela, IntervaloParcelas, Mensal, DiaDoMes, AcrescimoFinanceiro, Modificador, DataFixa, CargoMinimo, FormaPagamento, CodExterno) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", objCondicaoPagto.iCodigo, objCondicaoPagto.sDescReduzida, objCondicaoPagto.sDescricao, objCondicaoPagto.iEmPagamento, objCondicaoPagto.iEmRecebimento, objCondicaoPagto.iNumeroParcelas, objCondicaoPagto.iDiasParaPrimeiraParcela, objCondicaoPagto.iIntervaloParcelas, objCondicaoPagto.iMensal, objCondicaoPagto.iDiaDoMes, objCondicaoPagto.dAcrescimoFinanceiro, objCondicaoPagto.iModificador, objCondicaoPagto.iDataFixa, objCondicaoPagto.iCargoMinimo, objCondicaoPagto.iFormaPagamento, lCodAnt)
        If lErro <> AD_SQL_SUCESSO Then gError 205042
        
        For Each objParc In objCondicaoPagto.colParcelas
        
            With objParc
                lErro = Comando_Executar(alComando(4), "INSERT INTO CondicoesPagtoParc (Codigo, Seq, TipoDataBase, TipoIntervalo, Dias, Modificador, PercReceb) VALUES (?,?,?,?,?,?,?)", .iCodigo, .iSeq, .iTipoDataBase, .iTipoIntervalo, .iDias, .iModificador, .dPercReceb)
            End With
            If lErro <> AD_SQL_SUCESSO Then gError 205042
        
        Next
        
        lErro = Comando_ExecutarPos(alComando(2), "UPDATE ImportCondPagto SET CodCorporator = ?", alComando(1), objCondicaoPagto.iCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 205042
        
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211853
    
    Loop
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
     'fechar transacao
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 205044
    
    Importa_CondPagto = SUCESSO
    
    Exit Function
    
Erro_Importa_CondPagto:
    
    Importa_CondPagto = gErr
    
    Select Case gErr
        
        Case 205034
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 205035
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 205036, 205037
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ADMCONFIG", gErr)
            
        Case 205038 To 205041
        
        Case 205042
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_CONDICAOPAGTO", gErr, objCondicaoPagto.iCodigo)
        
        Case 205043
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_ADMCONFIG", gErr)
        
        Case 205044
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205045)
            
    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Call Transacao_Rollback
    
    Exit Function
    
End Function



Function Replica_Clientes() As Long

Dim lErro As Long
Dim lComando As Long
Dim lCliente As Long
Dim objCliente As ClassCliente, colCli As New Collection, colEnderecos As colEndereco
Dim lConexaoAnt As Long, lConexaoEmp2 As Long
Dim sDSN As String
Dim iLenDSN As Integer
Dim sParamOut As String
Dim iLenParamOut As Integer
Dim objEmpresa As New ClassDicEmpresa
Dim colEnd As Collection, objEnd As ClassEndereco

On Error GoTo Erro_Replica_Clientes

    lConexaoAnt = GL_lConexao

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 76329
    
    lErro = Comando_Executar(lComando, "SELECT Codigo FROM Clientes ORDER BY Codigo", lCliente)
    If lErro <> AD_SQL_SUCESSO Then gError 76330
        
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 76331
        
    Do While lErro = AD_SQL_SUCESSO
    
        Set objCliente = New ClassCliente
        Set colEnderecos = New colEndereco
        
        objCliente.lCodigo = lCliente
        
        lErro = CF("Cliente_Le", objCliente)
        If lErro <> SUCESSO And lErro <> 12293 Then gError 12431
        
        lErro = CF("Enderecos_Le_Cliente", colEnderecos, objCliente)
        If lErro <> SUCESSO Then gError 12304
        
        Set objCliente.objInfoUsu = colEnderecos
        
        objCliente.iVendedor = 0
        objCliente.iTabelaPreco = 0
        
        colCli.Add objCliente
            
        'Busca o proximo registro de ImportCli
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 76341
                           
    Loop
    
    Call Comando_Fechar(lComando)
    
    objEmpresa.lCodigo = 2
    
    lErro = Empresa_Le(objEmpresa)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    sDSN = objEmpresa.sStringConexao
    iLenDSN = Len(sDSN)
    sParamOut = String(1024, 0)
    iLenParamOut = 1024

    lConexaoEmp2 = Conexao_AbrirExt(AD_SQL_DRIVER_ODBC, sDSN, iLenDSN, sParamOut, iLenParamOut)
    If lConexaoEmp2 = 0 Then gError 182724
    
    GL_lConexao = lConexaoEmp2
    
    For Each objCliente In colCli
    
        Set colEnd = New Collection
        For Each objEnd In objCliente.objInfoUsu
            colEnd.Add objEnd
        Next
    
        lErro = CF("Cliente_Grava", objCliente, colEnd)
        If lErro <> SUCESSO Then gError 43294
    
    Next
   
    GL_lConexao = lConexaoAnt

    Call Conexao_FecharExt(lConexaoEmp2)

    Replica_Clientes = SUCESSO
    
    Exit Function
    
Erro_Replica_Clientes:

    Replica_Clientes = gErr
    
    Select Case gErr
    
        Case 76329
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 76330, 76331, 76341
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_IMPORTCLI", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161823)
            
    End Select
    
    GL_lConexao = lConexaoAnt
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function Replica_Produtos() As Long

Dim lErro As Long
Dim lComando As Long
Dim sProduto As String
Dim lConexaoAnt As Long, lConexaoEmp2 As Long
Dim sDSN As String
Dim iLenDSN As Integer
Dim sParamOut As String
Dim iLenParamOut As Integer
Dim objEmpresa As New ClassDicEmpresa
Dim objProduto As ClassProduto, colProd As New Collection
Dim colTabelaPrecoItem As New Collection

On Error GoTo Erro_Replica_Produtos

    lConexaoAnt = GL_lConexao

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 76329
    
    sProduto = String(STRING_MAXIMO, 0)
    
    lErro = Comando_Executar(lComando, "SELECT Codigo FROM Produtos ORDER BY Codigo", sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 76330
        
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 76331
        
    Do While lErro = AD_SQL_SUCESSO
    
        Set objProduto = New ClassProduto
        
        objProduto.sCodigo = sProduto
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 133752
        
        lErro = CF("CodigosBarra_Le_Produto", objProduto)
        If lErro <> SUCESSO Then gError 101730
               
        colProd.Add objProduto
            
        'Busca o proximo registro de ImportCli
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 76341
                           
    Loop
    
    Call Comando_Fechar(lComando)
    
    objEmpresa.lCodigo = 2
    
    lErro = Empresa_Le(objEmpresa)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    sDSN = objEmpresa.sStringConexao
    iLenDSN = Len(sDSN)
    sParamOut = String(1024, 0)
    iLenParamOut = 1024

    lConexaoEmp2 = Conexao_AbrirExt(AD_SQL_DRIVER_ODBC, sDSN, iLenDSN, sParamOut, iLenParamOut)
    If lConexaoEmp2 = 0 Then gError 182724
    
    GL_lConexao = lConexaoEmp2
    
    For Each objProduto In colProd
    
        lErro = CF("Produto_Grava", objProduto, colTabelaPrecoItem)
        If lErro <> SUCESSO Then gError 43294
    
    Next
   
    GL_lConexao = lConexaoAnt

    Call Conexao_FecharExt(lConexaoEmp2)

    Replica_Produtos = SUCESSO
    
    Exit Function
    
Erro_Replica_Produtos:

    Replica_Produtos = gErr
    
    Select Case gErr
    
        Case 76329
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 76330, 76331, 76341
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_IMPORTCLI", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161823)
            
    End Select
    
    GL_lConexao = lConexaoAnt
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

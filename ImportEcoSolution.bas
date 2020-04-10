Attribute VB_Name = "ImportEcoSolution"
Option Explicit

'******* Favor nao retirar este type deste lugar. Mario.
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
    dPercentMaisQuantCotAnt  As Double
    dPercentMenosQuantCotAnt As Double
    iConsideraQuantCotAnt As Integer
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
    iGrupo As Integer
    iSubGrupo As Integer
    sffl_CodClasse As String
    sffl_DescrClasse As String
    sffl_ONU As String
    iffl_R As Integer
End Type

Function Importa_Produtos() As Long
'Importa os dados referentes aos Produtos (Tabelas: ImportProd,ImportProdAux,ImportProdDesc)

Dim lErro As Long
Dim tImportProd As typeImportProd
Dim objProduto As New ClassProduto
Dim objProdutoPai As New ClassProduto
Dim lComando As Long
Dim lTransacao As Long
Dim lTamanho As Long
Dim iCreditoICMS As Integer
Dim iCreditoIPI As Integer
Dim sCodProduto As String
Dim colTabelaPrecoItem As New Collection
Dim sProdutoPai As String

On Error GoTo Erro_Importa_Produtos
    
    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 76348
    
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
    tImportProd.sffl_CodClasse = String(STRING_CATEGORIAPRODUTO_DESCRICAO, 0)
    tImportProd.sffl_DescrClasse = String(STRING_CATEGORIAPRODUTO_DESCRICAO, 0)
    tImportProd.sffl_ONU = String(STRING_CATEGORIAPRODUTO_DESCRICAO, 0)
    
    'Lê os registros da tabela ImportProd
    With tImportProd
    lErro = Comando_Executar(lComando, "SELECT Codigo,Tipo,Descricao,NomeReduzido,Modelo,Gerencial,Nivel,Substituto1,Substituto2,PrazoValidade,CodigoBarras,EtiquetasCodBarras,PesoLiq,PesoBruto," _
        & "Comprimento,Espessura,Largura,Cor,ObsFisica,ClasseUM,SiglaUMEstoque,SiglaUMCompra,SiglaUMVenda,Ativo,Faturamento,Compras,PCP," _
        & "KitBasico,KitInt,IPIAliquota,IPICodigo,IPICodDIPI,ControleEstoque,ICMSAgregaCusto,IPIAgregaCusto,FreteAgregaCusto,Apropriacao,ContaContabil,ContaContabilProducao,TemFaixaReceb,PercentMaisReceb,PercentMenosReceb,RecebForaFaixa,CreditoICMS,CreditoIPI,Residuo,Natureza," _
        & "CustoReposicao,OrigemMercadoria,TabelaPreco,TempoProducao,Rastro,HorasMaquina,PesoEspecifico,Linha,Grupo,SubGrupo,ffl_CodClasse,ffl_DescrClasse,ffl_ONU,ffl_R FROM ImportProd ORDER BY Codigo", .sCodigo, .iTipo, .sDescricao, .sNomeReduzido, .sModelo, .iGerencial, .iNivel, .sSubstituto1, .sSubstituto2, .iPrazoValidade, .sCodigoBarras, .iEtiquetasCodBarras, .dPesoLiq, .dPesoBruto, .dComprimento, .dEspessura, .dLargura, .sCor, _
        .sObsFisica, .iClasseUM, .sSiglaUMEstoque, .sSiglaUMCompra, .sSiglaUMVenda, .iAtivo, .iFaturamento, .iCompras, .iPCP, .iKitBasico, .iKitInt, .dIPIAliquota, .sIPICodigo, .sIPICodDIPI, .iControleEstoque, .iICMSAgregaCusto, .iIPIAgregaCusto, .iFreteAgregaCusto, .iApropriacaoCusto, .sContaContabil, .sContaContabilProducao _
        , .iTemFaixaReceb, .dPercentMaisReceb, .dPercentMenosReceb, .iRecebForaFaixa, .iCreditoICMS, .iCreditoIPI, .dResiduo, .iNatureza, .dCustoReposicao, .iOrigemMercadoria, .iTabelaPreco, .iTempoProducao, .iRastro, .lHorasMaquina, .dPesoEspecifico, .iLinha, .iGrupo, .iSubGrupo, .sffl_CodClasse, .sffl_DescrClasse, .sffl_ONU, .iffl_R)
    End With
    If lErro <> AD_SQL_SUCESSO Then gError 76350
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 76351
    
    Do While lErro = AD_SQL_SUCESSO
    
        Set objProduto = New ClassProduto

        objProduto.sCodigo = tImportProd.sCodigo
        
        If tImportProd.sCodigo <> "11191111" And tImportProd.sCodigo <> "11191201" And tImportProd.sCodigo <> "12144412" And tImportProd.sCodigo <> "13177202" And tImportProd.sCodigo <> "13209202" Then
        
            objProduto.iNatureza = tImportProd.iNatureza
            objProduto.iTipo = tImportProd.iTipo
    
            'Verifica se o Produto já está cadastrado
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 76352
        
            'Se não existe o Produto na tabela Produtos
            If lErro = 28030 Then
            
                'Preenche objProduto a partir de tImportProd
                lErro = Produto_PreencheObjetos(tImportProd, objProduto)
                If lErro <> SUCESSO Then gError 76353
                      
                'Grava o Produto
                lErro = CF("Produto_Grava_Trans", objProduto, colTabelaPrecoItem)
                If lErro <> SUCESSO Then gError 76354
    
            End If
        
        End If
        
        'Busca o proximo registro de ImportProd
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 76355
        
    Loop
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 76356
        
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161830)
            
    End Select
    
    Call Transacao_Rollback
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function Produto_PreencheObjetos(tImportProd As typeImportProd, ByVal objProduto As ClassProduto) As Long
'Preenche objProduto a partir dos dados existentes em tImportProd

Dim lErro As Long

On Error GoTo Erro_Produto_PreencheObjetos
        
    objProduto.sCodigo = tImportProd.sCodigo
    objProduto.sNomeReduzido = tImportProd.sNomeReduzido
    objProduto.dComprimento = tImportProd.dComprimento
    objProduto.dCustoReposicao = tImportProd.dCustoReposicao
    objProduto.dEspessura = tImportProd.dEspessura
    objProduto.dIPIAliquota = tImportProd.dIPIAliquota
    objProduto.dLargura = tImportProd.dLargura
    objProduto.dPercentMaisQuantCotAnt = tImportProd.dPercentMaisQuantCotAnt
    objProduto.dPercentMaisReceb = tImportProd.dPercentMaisReceb
    objProduto.dPercentMenosQuantCotAnt = tImportProd.dPercentMenosQuantCotAnt
    objProduto.dPercentMenosReceb = tImportProd.dPercentMenosReceb
    objProduto.dPesoBruto = tImportProd.dPesoBruto
    objProduto.dPesoEspecifico = tImportProd.dPesoEspecifico
    objProduto.dPesoLiq = tImportProd.dPesoLiq
    objProduto.dResiduo = tImportProd.dResiduo
    objProduto.iAtivo = tImportProd.iAtivo
    objProduto.iControleEstoque = tImportProd.iControleEstoque
    objProduto.iCreditoICMS = tImportProd.iCreditoICMS
    objProduto.iCreditoIPI = tImportProd.iCreditoIPI
    objProduto.iEtiquetasCodBarras = tImportProd.iEtiquetasCodBarras
    objProduto.iFreteAgregaCusto = tImportProd.iFreteAgregaCusto
    objProduto.iGerencial = tImportProd.iGerencial
    objProduto.iICMSAgregaCusto = tImportProd.iICMSAgregaCusto
    objProduto.iIPIAgregaCusto = tImportProd.iIPIAgregaCusto
    objProduto.iKitBasico = tImportProd.iKitBasico
    objProduto.iKitInt = tImportProd.iKitInt
    
    objProduto.iFaturamento = PRODUTO_VENDAVEL
    objProduto.iCompras = tImportProd.iCompras
    
    'Verifica a Apropriacao de Custo do Produto
    If objProduto.iCompras <> PRODUTO_COMPRAVEL Then
        objProduto.iApropriacaoCusto = APROPR_CUSTO_REAL
    Else
        objProduto.iApropriacaoCusto = APROPR_CUSTO_MEDIO
    End If
    
    objProduto.iNivel = tImportProd.iNivel
    
    objProduto.iPCP = tImportProd.iPCP
    objProduto.iPrazoValidade = tImportProd.iPrazoValidade
    objProduto.iRastro = tImportProd.iRastro
    objProduto.iTabelaPreco = tImportProd.iTabelaPreco
    objProduto.iTempoProducao = tImportProd.iTempoProducao
    objProduto.iTipo = tImportProd.iTipo
    objProduto.lHorasMaquina = tImportProd.lHorasMaquina
    objProduto.sCodigoBarras = tImportProd.sCodigoBarras
    objProduto.sContaContabil = tImportProd.sContaContabil
    objProduto.sContaContabilProducao = tImportProd.sContaContabilProducao
    objProduto.sCor = tImportProd.sCor
    objProduto.sDescricao = Trim(tImportProd.sDescricao)
    objProduto.sIPICodDIPI = tImportProd.sIPICodDIPI
    objProduto.sIPICodigo = tImportProd.sIPICodigo
    objProduto.sModelo = tImportProd.sModelo
    
    objProduto.iClasseUM = tImportProd.iClasseUM
    objProduto.sSiglaUMCompra = tImportProd.sSiglaUMCompra
    objProduto.sSiglaUMEstoque = tImportProd.sSiglaUMEstoque
    objProduto.sSiglaUMVenda = tImportProd.sSiglaUMVenda
    
    objProduto.sSubstituto1 = tImportProd.sSubstituto1
    objProduto.sSubstituto2 = tImportProd.sSubstituto2
    
    'Todo produto tem Origem= Nacional
    objProduto.iOrigemMercadoria = 0
    
    'Informações referentes a Compras
    objProduto.iConsideraQuantCotAnt = 1
    objProduto.dPercentMenosQuantCotAnt = 0
    objProduto.dPercentMaisQuantCotAnt = 0
    objProduto.iTemFaixaReceb = 0
    objProduto.dPercentMaisReceb = 0
    objProduto.dPercentMenosReceb = 0
    objProduto.iRecebForaFaixa = 1
    
    objProduto.sObsFisica = tImportProd.sObsFisica
    
    Produto_PreencheObjetos = SUCESSO
    
    Exit Function
    
Erro_Produto_PreencheObjetos:

    Produto_PreencheObjetos = gErr
    
    Select Case gErr
    
        Case 76404, 76435
            'Erros tratados nas rotinas chamadas
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161831)
            
    End Select
    
    Exit Function
    
End Function


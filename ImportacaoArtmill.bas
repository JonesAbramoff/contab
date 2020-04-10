Attribute VB_Name = "ImportacaoArtmill"
Option Explicit

'??? select * from pd_prodts where right(PD_PR_ID,2) <> cor_id
'??? produtos 9789 e 957704

Const STRING_TRANS_NOME = 100
Const STRING_TRANS_NOMEREDUZIDO = 50
Const STRING_TRANS_INSCRICAOESTADUAL = 15
Const STRING_TRANS_CGC = 14
Const STRING_TRANS_ENDERECO = 100
Const STRING_TRANS_GUIA = 10
Const STRING_TRANS_BAIRRO = 20
Const STRING_TRANS_CIDADE = 50
Const STRING_TRANS_UF = 2
Const STRING_TRANS_CEP = 8
Const STRING_TRANS_FONE = 12
Const STRING_TRANS_FAX = 12

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

Public Function GravaDadosTransps_Transportadoras() As Long
'Realiza o transporte dos dados da tabela TRANSPS para a Tabela Transportadoras

Dim alComando(1 To 3) As Long
Dim lTransacao As Long
Dim lErro As Long
Dim sNome As String
Dim sNomeReduzido As String
Dim sCGC As String
Dim sInscricaoEstadual As String
Dim sGuia As String
Dim sEndereco As String
Dim iCodigo As Integer
Dim iCodTransp As Integer
Dim sBairro As String
Dim sCidade As String, sCidade2 As String
Dim sUF As String
Dim sCEP As String
Dim sFone As String
Dim sFax As String
Dim objTransportadora As New ClassTransportadora
Dim objEndereco As New ClassEndereco
Dim colTransportadora As New Collection
Dim colEndereco As New Collection
Dim iIndice As Integer
Dim lCidade As Long
Dim lCodCid As Long

On Error GoTo Erro_GravaDadosTransps_Transportadoras

    STRING_ENDERECO = 255
    STRING_BAIRRO = 255
    STRING_CIDADE = 255
    STRING_TRANSPORTADORA_NOME = 255
    STRING_TRANSPORTADORA_NOME_REDUZIDO = 255
    
    'Abre a transação
    lTransacao = Transacao_Abrir
    If lTransacao = 0 Then gError 125826

    'Abrir comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir
        If alComando(iIndice) = 0 Then gError 125827
    Next

    'Inicializa as strings
    sNome = String(STRING_TRANS_NOME, 0)
    sNomeReduzido = String(STRING_TRANS_NOMEREDUZIDO, 0)
    sCGC = String(STRING_TRANS_CGC, 0)
    sInscricaoEstadual = String(STRING_TRANS_INSCRICAOESTADUAL, 0)
    sGuia = String(STRING_TRANS_GUIA, 0)
    sEndereco = String(STRING_TRANS_ENDERECO, 0)
    sBairro = String(STRING_TRANS_BAIRRO, 0)
    sCidade = String(STRING_TRANS_CIDADE, 0)
    sUF = String(STRING_TRANS_UF, 0)
    sCEP = String(STRING_TRANS_CEP, 0)
    sFone = String(STRING_TRANS_FONE, 0)
    sFax = String(STRING_TRANS_FAX, 0)

    'Realiza a leitura na tabela TRANSPS
    lErro = Comando_Executar(alComando(1), "SELECT PD_ID_TRANS, PD_TRANS, PD_TRANS_APELIDO, PD_TRANS_CGC, PD_TRANS_IE, PD_TRANS_END, PD_TRANS_GUIA, PD_TRANS_BAI, PD_TRANS_CID, ID_UF, PD_TRANS_CEP, PD_TRANS_FONE, PD_TRANS_FAX, CidadeID FROM PD_TRANSPS", _
                                        iCodigo, sNome, sNomeReduzido, sCGC, sInscricaoEstadual, sEndereco, sGuia, sBairro, sCidade, sUF, sCEP, sFone, sFax, lCidade)
    If lErro <> AD_SQL_SUCESSO Then gError 125828
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 125829
    
    Do While lErro = AD_SQL_SUCESSO
    
        Set objTransportadora = New ClassTransportadora
        Set objEndereco = New ClassEndereco
    
        'Preenche o objTransportadora e o objEndereco
        With objTransportadora
        
            .iCodigo = iCodigo
            .iViaTransporte = 7
            .dPesoMinimo = 0
            .sCGC = sCGC
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
            .iCodigoPais = 1
        
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
    
    GravaDadosTransps_Transportadoras = SUCESSO
    
    Exit Function
    
Erro_GravaDadosTransps_Transportadoras:

    GravaDadosTransps_Transportadoras = gErr
    
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 177296)
        
    End Select
    
    'Libera comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Call Transacao_Rollback
    
    Exit Function
    
End Function

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
Dim sProdutoPai As String, objProdutoCategoria As ClassProdutoCategoria

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
    tImportProd.sDimEmbalagem = String(255, 0)
    
    'Lê os registros da tabela ImportProd
    With tImportProd
    lErro = Comando_Executar(lComando, "SELECT Codigo,Tipo,Descricao,NomeReduzido,Modelo,Gerencial,Nivel,Substituto1,Substituto2,PrazoValidade,CodigoBarras,EtiquetasCodBarras,PesoLiq,PesoBruto," _
        & "Comprimento,Espessura,Largura,Cor,ObsFisica,ClasseUM,SiglaUMEstoque,SiglaUMCompra,SiglaUMVenda,Ativo,Faturamento,Compras,PCP," _
        & "KitBasico,KitInt,IPIAliquota,IPICodigo,IPICodDIPI,ControleEstoque,ICMSAgregaCusto,IPIAgregaCusto,FreteAgregaCusto,Apropriacao,ContaContabil,ContaContabilProducao,TemFaixaReceb,PercentMaisReceb,PercentMenosReceb,RecebForaFaixa,CreditoICMS,CreditoIPI,Residuo,Natureza," _
        & "CustoReposicao,OrigemMercadoria,TabelaPreco,TempoProducao,Rastro,HorasMaquina,PesoEspecifico,Linha, Dimensoes FROM ImportProd ORDER BY Codigo", .sCodigo, .iTipo, .sDescricao, .sNomeReduzido, .sModelo, .iGerencial, .iNivel, .sSubstituto1, .sSubstituto2, .iPrazoValidade, .sCodigoBarras, .iEtiquetasCodBarras, .dPesoLiq, .dPesoBruto, .dComprimento, .dEspessura, .dLargura, .sCor, _
        .sObsFisica, .iClasseUM, .sSiglaUMEstoque, .sSiglaUMCompra, .sSiglaUMVenda, .iAtivo, .iFaturamento, .iCompras, .iPCP, .iKitBasico, .iKitInt, .dIPIAliquota, .sIPICodigo, .sIPICodDIPI, .iControleEstoque, .iICMSAgregaCusto, .iIPIAgregaCusto, .iFreteAgregaCusto, .iApropriacaoCusto, .sContaContabil, .sContaContabilProducao, _
        .iTemFaixaReceb, .dPercentMaisReceb, .dPercentMenosReceb, .iRecebForaFaixa, .iCreditoICMS, .iCreditoIPI, .dResiduo, .iNatureza, .dCustoReposicao, .iOrigemMercadoria, .iTabelaPreco, .iTempoProducao, .iRastro, .lHorasMaquina, .dPesoEspecifico, .iLinha, .sDimEmbalagem)
    End With
    If lErro <> AD_SQL_SUCESSO Then gError 76350
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 76351
    
    Do While lErro = AD_SQL_SUCESSO
    
        Set objProduto = New ClassProduto

        'guarda o código do produto
        
        '??? CODIGO ESPECIFICO P/ARTMILL
        

        sCodProduto = "01" & Format(tImportProd.sCodigo, "000000") & "  "
        tImportProd.sCodigo = sCodProduto
        
        objProduto.sCodigo = tImportProd.sCodigo
        
        '??? CODIGO ESPECIFICO P/ARTMILL
        'define a natureza e o tipo do produto
        objProduto.iNatureza = NATUREZA_PROD_PRODUTO_ACABADO
        objProduto.iTipo = 1
        
        'Verifica se o Produto já está cadastrado
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 76352
    
        'Se não existe o Produto na tabela Produtos
        If lErro = 28030 Then
        
            'Busca o código do produto "pai"
            '??? especifico ARTMILL
            sProdutoPai = Left(objProduto.sCodigo, Len(objProduto.sCodigo) - 4) & "    "
            
            Set objProdutoPai = New ClassProduto
            
            objProdutoPai.sCodigo = sProdutoPai
            
            '??? especifico ARTMILL
            objProdutoPai.sDescricao = tImportProd.sDescricao
            
            '??? CODIGO ESPECIFICO P/ARTMILL
            'Preenche colCategoriaItem de objProduto
            If tImportProd.iLinha <> 0 Then
                
                Set objProdutoCategoria = New ClassProdutoCategoria
                objProdutoCategoria.sCategoria = "Setor"
                objProdutoCategoria.sProduto = objProdutoPai.sCodigo
                objProdutoCategoria.sItem = CStr(tImportProd.iLinha)
                objProdutoPai.colCategoriaItem.Add objProdutoCategoria
                
                Set objProdutoCategoria = New ClassProdutoCategoria
                objProdutoCategoria.sCategoria = "Setor"
                objProdutoCategoria.sProduto = objProduto.sCodigo
                objProdutoCategoria.sItem = CStr(tImportProd.iLinha)
                objProduto.colCategoriaItem.Add objProdutoCategoria
            End If
                  
            'já está cadastrado em Produtos. Se não estiver, grava o produto "pai"
            lErro = Produto_Define_ProdutoPai(objProdutoPai, tImportProd)
            If lErro <> SUCESSO Then gError 76411
                        
            'Preenche objProduto a partir de tImportProd
            lErro = Produto_PreencheObjetos(tImportProd, objProduto)
            If lErro <> SUCESSO Then gError 76353
                  
            'Grava o Produto
            lErro = CF("Produto_Grava_Trans", objProduto, colTabelaPrecoItem)
            If lErro <> SUCESSO Then gError 76354

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177297)
            
    End Select
    
    Call Transacao_Rollback
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Private Function Produto_Define_ProdutoPai(ByVal objProdutoPai As ClassProduto, tImportProd As typeImportProd) As Long
'Busca em ImportProdAux o produto "pai" e verifica se já está cadastrado em Produtos. Se não estiver, grava o
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
    
        '??? CODIGO ESPECIFICO P/ARTMILL
        'define a natureza e o tipo do produto
        objProdutoPai.iNatureza = NATUREZA_PROD_PRODUTO_ACABADO
        objProdutoPai.iTipo = 1
        
        'Preenche Produto "Pai" com os mesmos dados do Produto "Filho" que estão em tImportProd
        lErro = Produto_PreencheObjetos(tImportProd, objProdutoPai)
        If lErro <> SUCESSO Then gError 76410
        
        'Altera os dados específicos do produto "pai", que não são iguais ao "filho"
        '??? ARTMILL objProdutoPai.sDescricao = Trim(sDescricao)
        objProdutoPai.iGerencial = 1
        objProdutoPai.sCodigo = sProdutoPai
        objProdutoPai.sNomeReduzido = "P" & objProdutoPai.sCodigo
        objProdutoPai.iNivel = 2 '??? ARTMILL
        
        'Grava o "produto pai" no BD
        lErro = CF("Produto_Grava_Trans", objProdutoPai, colTabelaPrecoItem)
        If lErro <> SUCESSO Then gError 76403
    
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177298)
            
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function Produto_PreencheObjetos(tImportProd As typeImportProd, ByVal objProduto As ClassProduto) As Long
'Preenche objProduto a partir dos dados existentes em tImportProd

Dim lErro As Long

On Error GoTo Erro_Produto_PreencheObjetos
        
'    objProduto.sCodigo = tImportProd.sCodigo
    objProduto.sNomeReduzido = Trim("P" & objProduto.sCodigo)
    objProduto.dComprimento = tImportProd.dComprimento
    objProduto.dCustoReposicao = tImportProd.dCustoReposicao
    objProduto.dEspessura = tImportProd.dEspessura
    objProduto.dIPIAliquota = tImportProd.dIPIAliquota
    objProduto.dLargura = tImportProd.dLargura
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
    
    'Verifica se o produto pode ser comprado
    If objProduto.iNatureza = NATUREZA_PROD_PRODUTO_ACABADO Or objProduto.iNatureza = NATUREZA_PROD_PRODUTO_INTERMEDIARIO Then
        objProduto.iCompras = PRODUTO_NAO_COMPRAVEL
    Else
        objProduto.iCompras = PRODUTO_COMPRAVEL
    End If
    
    'Verifica a Apropriacao de Custo do Produto
    If objProduto.iNatureza = NATUREZA_PROD_PRODUTO_ACABADO Or objProduto.iNatureza = NATUREZA_PROD_PRODUTO_INTERMEDIARIO Then
        objProduto.iApropriacaoCusto = APROPR_CUSTO_REAL
    Else
        objProduto.iApropriacaoCusto = tImportProd.iApropriacaoCusto
    End If
    
    objProduto.iNivel = tImportProd.iNivel
    
    'Verifica se Produto pode participar da producao
    If objProduto.iNatureza = NATUREZA_PROD_MATERIA_PRIMA Or objProduto.iNatureza = NATUREZA_PROD_PRODUTO_INTERMEDIARIO Then
        objProduto.iPCP = PRODUTO_PCP_PODE
    Else
        objProduto.iPCP = PRODUTO_PCP_NAOPODE
    End If
    
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
    objProduto.sIPICodigo = IIf(Len(Trim(tImportProd.sIPICodigo)) <> 0, tImportProd.sIPICodigo & "00", "")
    objProduto.sModelo = tImportProd.sModelo
    
    'Define a Classe de UM do Produto
    lErro = Produto_Define_ClasseUM(tImportProd)
    If lErro <> SUCESSO Then gError 76435
    
    objProduto.iClasseUM = tImportProd.iClasseUM
    objProduto.sSiglaUMCompra = UCase(Trim(tImportProd.sSiglaUMCompra))
    objProduto.sSiglaUMEstoque = UCase(Trim(tImportProd.sSiglaUMEstoque))
    'UMVenda ficará igual a UMEstoque
    objProduto.sSiglaUMVenda = UCase(Trim(tImportProd.sSiglaUMEstoque))
    
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
        
    Set objProduto.objInfoUsu = New ClassProdutoInfoUsu
    
    objProduto.objInfoUsu.sDimEmbalagem = tImportProd.sDimEmbalagem
    
    Produto_PreencheObjetos = SUCESSO
    
    Exit Function
    
Erro_Produto_PreencheObjetos:

    Produto_PreencheObjetos = gErr
    
    Select Case gErr
    
        Case 76404, 76435
            'Erros tratados nas rotinas chamadas
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177299)
            
    End Select
    
    Exit Function
    
End Function

Function Produto_Define_ClasseUM(tImportProd As typeImportProd) As Long
'Define ClasseUM e SiglaUM do Produto, a partir dos dados lidos em ImportProd

Dim lErro As Long
Dim objClasseUM As New ClassClasseUM
Dim objUM As New ClassUnidadeDeMedida
Dim dQuantidade As Double
Dim bClasseDefinida As Boolean

On Error GoTo Erro_Produto_Define_ClasseUM
    
    Select Case UCase(Trim(tImportProd.sSiglaUMEstoque))
    
        Case "PÇ"
        
            tImportProd.iClasseUM = 100
        
        Case "KG"
            tImportProd.iClasseUM = 101
               
        Case "CJ"
            tImportProd.iClasseUM = 102
    
    End Select
    
    If tImportProd.iClasseUM = 0 Then gError 99999
    
    Produto_Define_ClasseUM = SUCESSO
    
    Exit Function
    
Erro_Produto_Define_ClasseUM:

    Produto_Define_ClasseUM = gErr
    
    Select Case gErr
        
        Case 76471
            'Erro tratado na rotina chamada
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177300)
            
    End Select
    
    Exit Function

            
        
End Function

Function Produto_Preenche_ColCategoria(tImportProd As typeImportProd, ByVal objProduto As ClassProduto) As Long
'Preenche colCategoriaItem de objProduto

Dim lErro As Long
Dim iItem As Integer
Dim objCategoriaProduto As New ClassCategoriaProduto
Dim objProdutoCategoria As New ClassProdutoCategoria
Dim colItensCategoria As New Collection
Dim objCategoriaItem As New ClassCategoriaProdutoItem

On Error GoTo Erro_Produto_Preenche_ColCategoria

    Set objProduto.colCategoriaItem = New Collection
    Set objProdutoCategoria = New ClassProdutoCategoria
    
    'Verifica se Linha está preenchida
    If tImportProd.iLinha > 0 Then
    
        'Preenche o objCategoriaProduto com a Categoria
        objCategoriaProduto.sCategoria = "Setor"

        'Verifica se a Categoria Produto existe. Se nao existir, insere no BD
        lErro = Valida_CategoriaProduto(objCategoriaProduto, colItensCategoria, tImportProd.iLinha, tImportProd.sCodigo, objProduto)
        If lErro <> SUCESSO Then gError 76378
        
    End If
    
    
    Produto_Preenche_ColCategoria = SUCESSO
    
    Exit Function
    
Erro_Produto_Preenche_ColCategoria:

    Produto_Preenche_ColCategoria = gErr
    
    Select Case gErr
    
        Case 76378, 76380, 76381, 76391, 76393, 76395
            'Erros tratados nas rotinas chamadas
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177301)
            
    End Select
    
    Exit Function
    
End Function

Function Valida_CategoriaProdutoItem(ByVal objCategoriaProduto As ClassCategoriaProduto, ByVal objCategoriaItem As ClassCategoriaProdutoItem, ByVal colItensCategoria As Collection) As Long
'Verifica se o Item da Categoria existe. Se nao existir, insere no BD

Dim lErro As Long

On Error GoTo Erro_Valida_CategoriaProdutoItem

    'Le o Item da Categoria
    lErro = CF("CategoriaProduto_Le_Item", objCategoriaItem)
    If lErro <> SUCESSO And lErro <> 22603 Then gError 76382
    
    'Se nao encontrou o Item ==> inclusao
    If lErro <> SUCESSO Then
    
        lErro = CF("CategoriaProduto_Grava_NovosItens", objCategoriaProduto, colItensCategoria)
        If lErro <> SUCESSO Then gError 76383
    
    End If
    
    Valida_CategoriaProdutoItem = SUCESSO
    
    Exit Function
    
Erro_Valida_CategoriaProdutoItem:

    Valida_CategoriaProdutoItem = gErr
    
    Select Case gErr
        
        Case 76382, 76383
            'Erros tratados nas rotinas chamadas
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177302)
            
    End Select
    
    Exit Function
            
End Function

Function Valida_CategoriaProduto(ByVal objCategoriaProduto As ClassCategoriaProduto, colItensCategoria As Collection, Item As Variant, ByVal sCodigo As String, ByVal objProduto As ClassProduto) As Long
'Verifica se a CategoriaProduto já está cadastrada. Se não estiver, grava a categoria no BD

Dim lErro As Long
Dim objCategoriaItem As New ClassCategoriaProdutoItem
Dim objProdutoCategoria As New ClassProdutoCategoria

On Error GoTo Erro_Valida_CategoriaProduto

    Set colItensCategoria = New Collection
    Set objCategoriaItem = New ClassCategoriaProdutoItem
    Set objProdutoCategoria = New ClassProdutoCategoria
        
    'Lê Categoria de Produto no BD
    lErro = CF("CategoriaProduto_Le", objCategoriaProduto)
    If lErro <> SUCESSO And lErro <> 22540 Then gError 76374
    
    'Se a Categoria não está cadastrada
    If lErro <> SUCESSO Then
        
        'Grava a Categoria no BD
        lErro = CategoriaProduto_Grava_EmTrans(objCategoriaProduto, colItensCategoria)
        If lErro <> SUCESSO Then gError 76375
        
    End If
    
    Set colItensCategoria = New Collection
    Set objCategoriaItem = New ClassCategoriaProdutoItem
    
    objCategoriaItem.sCategoria = objCategoriaProduto.sCategoria
    objCategoriaItem.sItem = Item
    colItensCategoria.Add objCategoriaItem
    
    'Verifica se o Item da Categoria existe. Se nao existir, insere no BD
    lErro = Valida_CategoriaProdutoItem(objCategoriaProduto, objCategoriaItem, colItensCategoria)
    If lErro <> SUCESSO Then gError 76379
    
    Set objProdutoCategoria = New ClassProdutoCategoria
    
    objProdutoCategoria.sCategoria = objCategoriaProduto.sCategoria
    objProdutoCategoria.sProduto = sCodigo
    objProdutoCategoria.sItem = objCategoriaItem.sItem
    objProduto.colCategoriaItem.Add objProdutoCategoria
    
    Valida_CategoriaProduto = SUCESSO
    
    Exit Function
    
Erro_Valida_CategoriaProduto:

    Valida_CategoriaProduto = gErr
    
    Select Case gErr
    
        Case 76374, 76375, 76379
            'Erros tratados nas rotinas chamadas
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177303)
            
    End Select
    
    Exit Function
    
End Function

Function Produto_DefineNatureza(ByVal objProduto As ClassProduto, tImportProd As typeImportProd) As Long
'Define a natureza e o tipo do Produto a partir do código do produto informado

Dim lErro As Long
Dim lTamanho As Long

On Error GoTo Erro_Produto_DefineNatureza

    'Se o código do produto comeca por letra
    If Not (IsNumeric(Left(objProduto.sCodigo, 1))) Then
    
        'se o código do produto começa com "Z"
        If Left(objProduto.sCodigo, 1) = "Z" Then
        
            objProduto.iNatureza = NATUREZA_PROD_PRODUTO_INTERMEDIARIO
            objProduto.iTipo = 11
            
        'se começa com qualquer outra letra
        Else
            objProduto.iNatureza = NATUREZA_PROD_MATERIA_PRIMA
            objProduto.iTipo = 12
            
        End If
        
    'Se o código do produto comeca por número
    Else
    
        'se o código tem 7 dígitos
        If Len(Trim(objProduto.sCodigo)) = 7 Then
        
            'se o produto comeca com 1 ==> manutenção
            If Left(objProduto.sCodigo, 1) = 1 Then
                objProduto.iNatureza = NATUREZA_PROD_PRODUTO_MANUTENCAO
                objProduto.iTipo = 1
            'se o produto começa com 2 ==> material de laboratório
            ElseIf Left(objProduto.sCodigo, 1) = 2 Then
                objProduto.iNatureza = NATUREZA_PROD_OUTROS
                objProduto.iTipo = 2
            'se o produto começa com 3 ==> segurança
            ElseIf Left(objProduto.sCodigo, 1) = 3 Then
                objProduto.iNatureza = NATUREZA_PROD_OUTROS
                objProduto.iTipo = 3
            'se o produto começa com 4 ==> limpeza
            ElseIf Left(objProduto.sCodigo, 1) = 4 Then
                objProduto.iNatureza = NATUREZA_PROD_OUTROS
                objProduto.iTipo = 4
            'se o produto começa com 5 ==> expediente p/ papelaria
            ElseIf Left(objProduto.sCodigo, 1) = 5 Then
                objProduto.iNatureza = NATUREZA_PROD_OUTROS
                objProduto.iTipo = 5
            'se o produto começa com 6 ==> embalagem
            ElseIf Left(objProduto.sCodigo, 1) = 6 Then
                objProduto.iNatureza = NATUREZA_PROD_EMBALAGENS
                objProduto.iTipo = 6
            'se o produto começa com 7 ==> material para industrialização
            ElseIf Left(objProduto.sCodigo, 1) = 7 Then
                objProduto.iNatureza = NATUREZA_PROD_OUTROS
                objProduto.iTipo = 7
            'se o produto começa com 8 ==> outros
            ElseIf Left(objProduto.sCodigo, 1) = 8 Then
                objProduto.iNatureza = NATUREZA_PROD_OUTROS
                objProduto.iTipo = 8
            'se o produto começa com 9 ==> outros
            ElseIf Left(objProduto.sCodigo, 1) = 9 Then
                objProduto.iNatureza = NATUREZA_PROD_OUTROS
                objProduto.iTipo = 9
            End If
            
        'se o código tem até 4 dígitos
        ElseIf Len(Trim(objProduto.sCodigo)) <= 4 Then
            
            lTamanho = Len(Trim(objProduto.sCodigo))
            '???Confirmar formatacao
            'coloca o código do produto no formato de 7 dígitos
            objProduto.sCodigo = Format(objProduto.sCodigo, "0000000")
            tImportProd.sCodigo = objProduto.sCodigo
            objProduto.iNatureza = NATUREZA_PROD_PRODUTO_ACABADO
            objProduto.iTipo = 10
            
        End If
            
    End If
    
    Produto_DefineNatureza = SUCESSO
    
    Exit Function
    
Erro_Produto_DefineNatureza:

    Produto_DefineNatureza = gErr
    
    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177304)
            
    End Select
    
    Exit Function
    
End Function

Function ClasseUM_Grava_EmTrans(ByVal objClasseUM As ClassClasseUM, ByVal colSiglas As Collection) As Long
'Inclui ou altera uma Classe de Unidade de Medida
'Tabelas:ClasseUM e UnidadeDeMedida

Dim lErro As Long
Dim iIndice As Integer
Dim iEditavel As Integer
Dim colUMCopia As New Collection
Dim sDescricao As String, sSigla As String
Dim iClasse As Integer, sSiglaUM As String, sNome As String, dQuantidade As Double, sSiglaUMBase As String
Dim objUM As New ClassUnidadeDeMedida
Dim alComando(1 To 8) As Long
Dim iTotalClasseUM As Integer

On Error GoTo Erro_ClasseUM_Grava_EmTrans

    'Cria uma cópia "de trabalho" da coleção passada como parametro
    For Each objUM In colSiglas
        colUMCopia.Add objUM
    Next

    For iIndice = LBound(alComando) To UBound(alComando)

        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 22445

    Next

    sSigla = String(STRING_UM_SIGLA, 0)
    sDescricao = String(STRING_CLASSEUM_DESCRICAO, 0)

    'Pesquisa no BD a Classe em questão
    lErro = Comando_ExecutarPos(alComando(1), "SELECT Descricao, Sigla FROM ClasseUM WHERE Classe = ?", 0, sDescricao, sSigla, objClasseUM.iClasse)
    If lErro <> AD_SQL_SUCESSO Then Error 22447

    'Lê a Classe, se estiver no BD
    lErro = Comando_BuscarProximo(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 22448

    'Se a Classe já existe
    If lErro = AD_SQL_SUCESSO Then

        'Se trocou a Descrição da classe ou a sigla base
        If sDescricao <> objClasseUM.sDescricao Or sSigla <> objClasseUM.sSiglaUMBase Then

            'Alterar a ClasseUM
            lErro = Comando_ExecutarPos(alComando(2), "UPDATE ClasseUM SET Descricao = ?, Sigla = ?", alComando(1), objClasseUM.sDescricao, objClasseUM.sSiglaUMBase)
            If lErro <> AD_SQL_SUCESSO Then Error 22449

        End If

        sSiglaUM = String(STRING_UM_SIGLA, 0)
        sNome = String(STRING_CLASSEUM_NOME, 0)
        sSiglaUMBase = String(STRING_UM_SIGLA, 0)

        'Percorre as siglas da Classe no BD
        lErro = Comando_ExecutarPos(alComando(3), "SELECT Sigla, Nome, Quantidade, SiglaUMBase, Editavel FROM UnidadesDeMedida WHERE Classe = ?", 0, sSiglaUM, sNome, dQuantidade, sSiglaUMBase, iEditavel, objClasseUM.iClasse)
        If lErro <> AD_SQL_SUCESSO Then Error 22450

        'Lê a Sigla da Classe, se estiver no BD
        lErro = Comando_BuscarProximo(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 22451

        Do While lErro <> AD_SQL_SEM_DADOS

            'Procura a Sigla na coleção correspondente ao grid
            iIndice = CF("ClasseUM_Grava_SiglaNaColecao", colUMCopia, sSiglaUM)

            If iIndice <> 0 Then

                Set objUM = colUMCopia.Item(iIndice)

                'Se houver sido alterada a Descrição(Nome), o Fator de Conversão(Quantidade) ou a Sigla Base(SiglaUMBase)
                If objUM.sNome <> sNome Or objUM.dQuantidade <> dQuantidade Or objUM.sSiglaUMBase <> sSiglaUMBase Then

                    'se a conversao já foi utilizada a sigla nao é editavel
                    If iEditavel <> UM_EDITAVEL And (objUM.sSiglaUMBase <> sSiglaUMBase Or objUM.dQuantidade <> dQuantidade) Then
                        Error 22920
                    End If

                    'Altera a Tabela UnidadesDeMedida
                    lErro = Comando_ExecutarPos(alComando(4), "UPDATE UnidadesDeMedida SET Nome = ?, Quantidade = ?, SiglaUMBase = ?", alComando(3), objUM.sNome, objUM.dQuantidade, objUM.sSiglaUMBase)
                    If lErro <> AD_SQL_SUCESSO Then Error 22452

                End If

                'Retira da coleção
                colUMCopia.Remove (iIndice)

            Else

                'se a conversao já foi utilizada a sigla nao é editavel
                If iEditavel <> UM_EDITAVEL Then Error 22921

                'Se o par (classe,sigla) estiver sendo usado em Produtos, TiposDeProduto, Itens de Pedido de venda,.... nao poderá ser excluido
                lErro = ClasseUM_Exclui2(objClasseUM.iClasse, sSiglaUM)
                If lErro <> SUCESSO Then Error 22455

                'Excluir registro em UnidadesDeMedida
                lErro = Comando_ExecutarPos(alComando(5), "DELETE FROM UnidadesDeMedida", alComando(3))
                If lErro <> AD_SQL_SUCESSO Then Error 22454

            End If

            'Lê a Sigla da Classe, se estiver no BD
            lErro = Comando_BuscarProximo(alComando(3))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 22453

        Loop

    Else 'A Classe não existe

        'Se for Versão Light
        If giTipoVersao = VERSAO_LIGHT Then
            
            'Lê o Número Total de Classes
            lErro = CF("ClasseUM_Le_Total", iTotalClasseUM)
            If lErro <> SUCESSO Then Error 61191
            
            'Se ultrapassou o número máximo de Classes UM ---> ERRO
            If iTotalClasseUM >= LIMITE_CLASSE_UM_VGLIGHT Then Error 61192
        
        End If

        'Insere em ClasseUM, criando uma nova Classe
        lErro = Comando_Executar(alComando(6), "INSERT INTO ClasseUM (Classe,Descricao, Sigla) VALUES(?,?,?)", objClasseUM.iClasse, objClasseUM.sDescricao, objClasseUM.sSiglaUMBase)
        If lErro <> AD_SQL_SUCESSO Then Error 22456

    End If

    'Grava as Siglas que ainda não faziam parte da Classe
    lErro = CF("ClasseUM_Grava_NovasSiglas", objClasseUM, colUMCopia)
    If lErro <> SUCESSO Then Error 22457

    'libera comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    ClasseUM_Grava_EmTrans = SUCESSO

    Exit Function

Erro_ClasseUM_Grava_EmTrans:

    ClasseUM_Grava_EmTrans = Err

    Select Case Err

        Case 22457, 22455, 61191

        Case 22445
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 22447, 22448
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLASSEUM", Err)

        Case 22449
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MODIFICACAO_CLASSEUM", Err)

        Case 22452
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MODIFICACAO_UNIDADESDEMEDIDA", Err)

        Case 22454
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_UNIDADESDEMEDIDA", Err, objClasseUM.iClasse)

        Case 22450, 22451, 22453
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_UNIDADESDEMEDIDA", Err)

        Case 22456
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_CLASSEUM", Err, , objClasseUM.iClasse)

        Case 22920, 22921
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SIGLA_UM_NAO_EDITAVEL", Err, sSiglaUM, objClasseUM.iClasse)
        
        Case 61192
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LIMITE_CLASSE_UM", Err, LIMITE_CLASSE_UM_VGLIGHT)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 177305)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'Alterado por Ivan em 04/04/03
Function CategoriaProduto_Grava_EmTrans(ByVal objCategoriaProduto As ClassCategoriaProduto, ByVal colItensCategoria As Collection) As Long
'inclui ou altera uma categoria de produtos e seus valores
'tabelas:CategoriaProduto e CategoriaProdutoItem

Dim lErro As Long, iIndice As Integer, colItensCategoriaCopia As New Collection
Dim sCategoriaDescricao As String, sCategoriaSigla As String
Dim objCategoriaProdutoItem As New ClassCategoriaProdutoItem
Dim alComando(1 To 6) As Long
Dim tCategoriaItem As typeCategoriaProdutoItem

On Error GoTo Erro_CategoriaProduto_Grava_EmTrans

    'cria uma copia "de trabalho" da colecao passada como parametro
    For Each objCategoriaProdutoItem In colItensCategoria
        colItensCategoriaCopia.Add objCategoriaProdutoItem
    Next
    
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 27121
    Next

    sCategoriaDescricao = String(STRING_CATEGORIAPRODUTO_DESCRICAO, 0)
    sCategoriaSigla = String(STRING_CATEGORIAPRODUTO_SIGLA, 0)
    
    'Pesquisa no BD a categoria em questão
    lErro = Comando_ExecutarPos(alComando(1), "SELECT Descricao, Sigla FROM CategoriaProduto WHERE Categoria = ?", 0, sCategoriaDescricao, sCategoriaSigla, objCategoriaProduto.sCategoria)
    If lErro <> AD_SQL_SUCESSO Then Error 27124

    'Lê a categoria, se estiver no BD
    lErro = Comando_BuscarProximo(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 27125

    'Se a categoria existir...
    If lErro = AD_SQL_SUCESSO Then
    
        'Se a descrição ou sigla da categoria for diferente
        If sCategoriaDescricao <> objCategoriaProduto.sDescricao Or sCategoriaSigla <> objCategoriaProduto.sSigla Then
             
            'Alterar a descrição ou a sigla da categoria
            lErro = Comando_ExecutarPos(alComando(2), "UPDATE CategoriaProduto SET Descricao = ?, Sigla = ?", alComando(1), objCategoriaProduto.sDescricao, objCategoriaProduto.sSigla)
            If lErro <> AD_SQL_SUCESSO Then Error 27126
            
        End If
                
        tCategoriaItem.sDescricao = String(STRING_CATEGORIAPRODUTOITEM_DESCRICAO, 0)
        tCategoriaItem.sItem = String(STRING_CATEGORIAPRODUTOITEM_ITEM, 0)
        
        'Percorrer todos os itens atuais da categoria no bd
        With tCategoriaItem
            lErro = Comando_ExecutarPos(alComando(3), "SELECT Item, Ordem, Descricao, Valor1, Valor2, Valor3, Valor4, Valor5, Valor6, Valor7, Valor8 FROM CategoriaProdutoItem WHERE Categoria = ?", 0, _
                .sItem, .iOrdem, .sDescricao, .dValor1, .dValor2, .dValor3, .dValor4, .dValor5, .dValor6, .dValor7, .dValor8, objCategoriaProduto.sCategoria)
        End With
        If lErro <> AD_SQL_SUCESSO Then Error 27127
        
        'Lê o item da categoria, se estiver no BD
        lErro = Comando_BuscarProximo(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 27128

        Do While lErro <> AD_SQL_SEM_DADOS
            
            'Procura o item na coleção correspondente ao grid
            iIndice = CF("CategoriaProduto_Grava_ItemNaColecao", colItensCategoriaCopia, tCategoriaItem.sItem)
            
            'Se (ainda) existir no grid uma linha com a mesma chave
            If iIndice <> 0 Then
            
                Set objCategoriaProdutoItem = colItensCategoriaCopia.Item(iIndice)
                           
                'Alterar o registro no Bd
                 With objCategoriaProdutoItem
                    lErro = Comando_ExecutarPos(alComando(4), "UPDATE CategoriaProdutoItem SET Item =?, Ordem = ?, Descricao = ?, Valor1 = ?, Valor2 = ?, Valor3 = ?, Valor4 = ?, Valor5 = ?, Valor6 = ?, Valor7 = ?, Valor8 = ?", alComando(3), _
                        .sItem, .iOrdem, .sDescricao, .dValor1, .dValor2, .dValor3, .dValor4, .dValor5, .dValor6, .dValor7, .dValor8)
                 End With
                 If lErro <> AD_SQL_SUCESSO Then Error 27129

                'excluir o item da colecao
                colItensCategoriaCopia.Remove (iIndice)
                
            Else
            
                lErro = CF("CategoriaProdutoItem_NaoUtilizado", objCategoriaProduto.sCategoria, tCategoriaItem.sCategoria)
                If lErro Then Error 27156
                
                'Excluir o item do bd
                lErro = Comando_ExecutarPos(alComando(5), "DELETE FROM CategoriaProdutoItem", alComando(3))
                If lErro <> AD_SQL_SUCESSO Then Error 27130
                
            End If
            
            'Lê o item da categoria, se estiver no BD
            lErro = Comando_BuscarProximo(alComando(3))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 27131

        Loop
        
    Else 'a categoria não existe...
    
        'Criá-la ( incluí-la em CategoriaProduto )
        lErro = Comando_Executar(alComando(6), "INSERT INTO CategoriaProduto (Categoria, Descricao, Sigla) VALUES(?,?,?)", objCategoriaProduto.sCategoria, objCategoriaProduto.sDescricao, objCategoriaProduto.sSigla)
        If lErro <> AD_SQL_SUCESSO Then Error 27132
        
    End If
    
    'Grava os itens que ainda nao faziam parte da categoria
    lErro = CF("CategoriaProduto_Grava_NovosItens", objCategoriaProduto, colItensCategoriaCopia)
    If lErro <> SUCESSO Then Error 27133
    
    'Libera comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    CategoriaProduto_Grava_EmTrans = SUCESSO

    Exit Function

Erro_CategoriaProduto_Grava_EmTrans:

    CategoriaProduto_Grava_EmTrans = Err

    Select Case Err

        Case 27133, 27156
        
        Case 27121
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 27124, 27125
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CATEGORIAPRODUTO", Err)
        
        Case 27126
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MODIFICACAO_CATEGORIAPRODUTO", Err)
        
        Case 27129
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MODIFICACAO_CATEGORIAPRODUTOITEM", Err)
        
        Case 27130
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CATEGORIAPRODUTOITEM", Err)
        
        Case 27127, 27128, 27131
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CATEGORIAPRODUTOITENS_CATEGORIA", Err)
        
        Case 27132
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_CATEGORIAPRODUTO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 177306)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function
        
End Function

Private Function ClasseUM_Exclui2(ByVal iClasseUM As Integer, ByVal sSiglaUM As String) As Long
'Retorna SUCESSO se a Classe e a Sigla não estiverem sendo usadas nas tabelas Produtos , TiposDeProduto e ItensPedidoDeVenda

Dim lErro As Long, lComando As Long, iClasse As Integer, sSigla As String

On Error GoTo Erro_ClasseUM_Exclui2

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 22431

    sSigla = String(STRING_UM_SIGLA, 0)

    'Pesquisa se a Classe e a Sigla estão sendo usadas na tabela Produtos
    lErro = Comando_Executar(lComando, "SELECT ClasseUM, SiglaUMEstoque FROM Produtos WHERE ClasseUM = ? AND SiglaUMEstoque = ?", iClasse, sSigla, iClasseUM, sSiglaUM)
    If lErro <> AD_SQL_SUCESSO Then Error 22432

    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 22433

    'se a Classe e a Sigla estão sendo usadas na tabela Produtos => erro
    If lErro <> AD_SQL_SEM_DADOS Then Error 22434

    'Pesquisa se a Classe e a sigla estão sendo usadas na tabela TiposDeProduto
    lErro = Comando_Executar(lComando, "SELECT ClasseUM, SiglaUMEstoque FROM TiposDeProduto WHERE ClasseUM = ? AND SiglaUMEstoque = ?", iClasse, sSigla, iClasseUM, sSiglaUM)
    If lErro <> AD_SQL_SUCESSO Then Error 57819

    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 57820

    'Se a Classe está sendo usada na tabela TiposDeProduto => erro
    If lErro <> AD_SQL_SEM_DADOS Then Error 57821
    
    sSigla = String(STRING_UM_SIGLA, 0)

    'Pesquisa se a Classe e a Sigla estão sendo usadas na tabela ItensPedidoDeVenda
    lErro = Comando_Executar(lComando, "SELECT ClasseUM, UnidadeMed FROM ItensPedidoDeVenda WHERE ClasseUM = ? AND UnidadeMed = ?", iClasse, sSigla, iClasseUM, sSiglaUM)
    If lErro <> AD_SQL_SUCESSO Then Error 22435

    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 22436

    'se a categoria está associada a algum produto => erro
    If lErro <> AD_SQL_SEM_DADOS Then Error 22437

    Call Comando_Fechar(lComando)
    
    ClasseUM_Exclui2 = SUCESSO

    Exit Function

Erro_ClasseUM_Exclui2:

    ClasseUM_Exclui2 = Err

    Select Case Err

        Case 22431
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 22432, 22433
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS1", Err)

        Case 22435, 22436
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDIDODEVENDA", Err)

        Case 22434
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLASSEUM_E_SIGLAUM_UTILIZADAS_PRODUTOS", Err, iClasseUM, sSiglaUM)

        Case 22437
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLASSEUM_E_SIGLAUM_UTILIZADAS_ITENSPEDIDODEVENDA", Err, iClasseUM, sSiglaUM)

        Case 57819, 57820
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOSDEPRODUTO1", Err)
            
        Case 57821
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLASSEUM_E_SIGLAUM_UTILIZADAS_TIPOSDEPRODUTO", Err, iClasseUM, sSiglaUM)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 177307)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function Importacao_Vendedores() As Long

Dim lErro As Long
Dim alComando(1) As Long
Dim lTransacao As Long
Dim tVendedor As typeVendedor
Dim tEndereco As typeEndereco
Dim objVendedor As ClassVendedor
Dim objEndereco As New ClassEndereco
Dim sArquivo As String
Dim tTipoVendedor As typeTipoVendedor
Dim objTipoVendedor As ClassTipoVendedor
Dim lCodigo As Long
Dim iIndice As Integer
Dim colTiposVendedores As New Collection
Dim sCGCAux As String

Const STRING_VENDEDOR_NOME_REDUZIDO_USU = STRING_VENDEDOR_NOME_REDUZIDO + 30
Const STRING_ENDERECO_USU = 50
Const STRING_BAIRRO_USU = 50
Const STRING_CEP_USU = STRING_CEP
Const STRING_CIDADE_USU = 50
Const STRING_TELEFONE_USU = 12
Const STRING_AGENCIA_USU = STRING_AGENCIA + 3
Const STRING_CONTA_CORRENTE_USU = STRING_CONTA_CORRENTE + 1

On Error GoTo Erro_Importacao_Vendedores

    STRING_ENDERECO = 255
    STRING_BAIRRO = 255
    STRING_CIDADE = 255
    
    sArquivo = App.Path & "\Vendedores_Log_Importacao.txt"

    If Len(Dir(sArquivo)) > 0 Then Kill sArquivo

    'Arquivo de log
    Open sArquivo For Append As #1

    'Executa abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then
        Print #1, "Erro: não foi possível abrir transação para executar importação da tabela VendedoresOrigem."
        gError 1000
    End If

    'Executa a abertura do Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then
            Print #1, "Erro: não foi possível abrir comando para executar leitura da tabela VendedoresOrigem."
            gError 1000
        End If
    Next

    'Inicializa as strings que serão utilizadas na leitura
    tVendedor.sNome = String(STRING_VENDEDOR_NOME, 0)
    tVendedor.sNomeReduzido = String(STRING_VENDEDOR_NOME_REDUZIDO_USU, 0)
    tVendedor.sAgencia = String(STRING_AGENCIA_USU, 0)
    tVendedor.sContaCorrente = String(STRING_CONTA_CORRENTE_USU, 0)
    tVendedor.sCGC = String(STRING_CGC, 0)
    tVendedor.sInscricaoEstadual = String(STRING_INSCR_EST, 0)

    tTipoVendedor.sDescricao = String(STRING_TIPO_DE_VENDEDOR_DESCRICAO, 0)
    tEndereco.sEndereco = String(STRING_ENDERECO_USU, 0)
    tEndereco.sBairro = String(STRING_BAIRRO_USU, 0)
    tEndereco.sCidade = String(STRING_CIDADE_USU, 0)
    tEndereco.sCEP = String(STRING_CEP_USU, 0)
    tEndereco.sSiglaEstado = String(STRING_ESTADO, 0)
    tEndereco.sTelefone1 = String(STRING_TELEFONE_USU, 0)
    tEndereco.sTelefone2 = String(STRING_TELEFONE_USU, 0)
    tEndereco.sFax = String(STRING_TELEFONE_USU, 0)
    tEndereco.sContato = String(STRING_CONTATO, 0)
    tEndereco.sEmail = String(STRING_EMAIL, 0)

    '*** GRAVAÇÃO DE TIPO DE VENDEDOR ********
    'Lê os diferentes de tipos de vendedor existentes na tabela VendedoresOrigem
    lErro = Comando_Executar(alComando(0), "SELECT DISTINCT Tipo FROM VendedoresOrigem WHERE Tipo <>'' ORDER BY Tipo", tTipoVendedor.sDescricao)
    If lErro <> AD_SQL_SUCESSO Then
        Print #1, "Erro: não foi possível ler a tabela VendedoresOrigem."
        gError 1000
    End If

    'Busca o primeiro tipo de vendedor encontrado
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then
        Print #1, "Erro: não foi possível ler a tabela VendedoresOrigem."
        gError 1000
    End If

    'Se não encontrou nenhum tipo
    If lErro = AD_SQL_SEM_DADOS Then Print #1, "Erro: nenhum tipo de vendedor foi encontrado."

    'Enquanto houverem tipos de vendedores
    Do While lErro = AD_SQL_SUCESSO

        'Instancia um novo obj
        Set objTipoVendedor = New ClassTipoVendedor

        'Obtém um código para o novo tipo de vendedor
        lErro = CF("Config_ObterAutomatico_EmTrans", "CPRConfig", "NUM_PROX_TIPO_VENDEDOR", "TiposDeVendedor", "Codigo", lCodigo)
        If lErro <> SUCESSO Then
            Print #1, "Não foi possível gravar o tipo de vendedor " & objTipoVendedor.sDescricao
        Else

            With tTipoVendedor

                'Transfere o código lido para o obj
                objTipoVendedor.iCodigo = lCodigo
                objTipoVendedor.sDescricao = .sDescricao
                objTipoVendedor.dPercComissao = 0
                objTipoVendedor.dPercComissaoEmissao = 1
                objTipoVendedor.dPercComissaoBaixa = 0
            End With

            lErro = TipoVendedor_Grava_Importacao_EmTrans(objTipoVendedor, 1)
            If lErro <> SUCESSO Then
                Print #1, "Não foi possível gravar o tipo de vendedor " & objTipoVendedor.sDescricao
            Else
                'Guarda o tipo na coleção de tipos gravados
                colTiposVendedores.Add objTipoVendedor
            End If

        End If

        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then
            Print #1, "Erro: não foi possível ler a tabela TransportadoraOrigem."
            gError 1000
        End If

    Loop
    '*** FIM DA GRAVAÇÃO DE TIPO DE VENDEDOR ********

    'tem que reinicializar a variável, pois a mesma já foi utilizada em uma leitura anterior
    tTipoVendedor.sDescricao = String(STRING_TIPO_DE_VENDEDOR_DESCRICAO, 0)
    
    '*** GRAVAÇÃO DE DE VENDEDOR ********
    'Lê os diferentes de tipos de vendedor existentes na tabela VendedoresOrigem
    lErro = Comando_Executar(alComando(1), "SELECT DISTINCT Codigo, Nome, NomeReduzido, PercComissao, Tipo, Endereco, Bairro, Cidade, CEP, SiglaEstado, Telefone1, Telefone2, Fax, Email, Contato, Banco, Agencia, ContaCorrente, CGC, InscricaoEstadual FROM VendedoresOrigem ORDER BY Codigo", tVendedor.iCodigo, tVendedor.sNome, tVendedor.sNomeReduzido, tVendedor.dPercComissao, tTipoVendedor.sDescricao, tEndereco.sEndereco, tEndereco.sBairro, tEndereco.sCidade, tEndereco.sCEP, tEndereco.sSiglaEstado, tEndereco.sTelefone1, tEndereco.sTelefone2, tEndereco.sFax, tEndereco.sEmail, tEndereco.sContato, tVendedor.iBanco, tVendedor.sAgencia, tVendedor.sContaCorrente, tVendedor.sCGC, tVendedor.sInscricaoEstadual)
    If lErro <> AD_SQL_SUCESSO Then
        Print #1, "Erro: não foi possível ler a tabela VendedoresOrigem."
        gError 1000
    End If

    'Busca o primeiro tipo de vendedor encontrado
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then
        Print #1, "Erro: não foi possível ler a tabela VendedoresOrigem."
        gError 1000
    End If

    'Se não encontrou nenhum tipo
    If lErro = AD_SQL_SEM_DADOS Then Print #1, "Erro: nenhum vendedor foi encontrado."

    'Enquanto houverem tipos de vendedores
    Do While lErro = AD_SQL_SUCESSO

        Set objVendedor = New ClassVendedor

        With tVendedor

            objVendedor.iCodigo = .iCodigo

            'Se o nome não foi informado na tabela origem
            If Len(Trim(.sNome)) = 0 Then

                'Monta um nome fictício para o vendedor
                objVendedor.sNome = .sNomeReduzido

                'Grava no arquivo de log o erro referente a esse vendedor
                Print #1, "Vendedor: " & .iCodigo & "|Erro: o vendedor não estava com o nome preenchido. O registro foi gravado com nome " & objVendedor.sNome & "."

            'Se o nome está preenchido
            Else
                'Transfere o nome lido para o obj
                objVendedor.sNome = .sNome

            End If

            'Se o nome reduzido não foi preenchido
            If Len(Trim(.sNomeReduzido)) = 0 Then

                'Monta um nome reduzido fictício para o vendedor
                objVendedor.sNomeReduzido = "Vendedor" & CStr(.iCodigo)

                'Grava no arquivo de log o erro referente a essa transportadora
                Print #1, "Vendedor: " & .iCodigo & "|Erro: o vendedor não estava com o nome reduzido preenchido. O registro foi gravado com nome " & objVendedor.sNomeReduzido & "."

            'Se o nome reduzido está preenchido
            Else

                'Se o nome for maior que 20 =>
                If Len(Trim(.sNomeReduzido)) > STRING_VENDEDOR_NOME_REDUZIDO Then
                
                    'Indica no arquivo de log que o nome do cliente foi truncado
                    Print #1, "Vendedor: " & .iCodigo & "|Erro: o nome reduzido do vendedor foi alterado de " & .sNomeReduzido & " para " & Mid(.sNomeReduzido, 1, STRING_VENDEDOR_NOME_REDUZIDO) & "."
                
                End If
                
                'Transfere o nome reduzido lido para o obj
                objVendedor.sNomeReduzido = Mid(.sNomeReduzido, 1, STRING_VENDEDOR_NOME_REDUZIDO)

                'Se o nome reduzido começar com um caracter numérico => alerta que é inválido
                If IsNumeric(Mid(.sNomeReduzido, 1, 1)) Then Print #1, "O vendedor " & objVendedor.iCodigo & SEPARADOR & objVendedor.sNomeReduzido & " foi importado com nome reduzido inválido, pois nome reduzido não pode começar com número."

            End If

            'Se o tipo de vendedor está preenchido
            If Len(Trim(tTipoVendedor.sDescricao)) > 0 Then

                'Procura o código para o tipo de vendedor
                For Each objTipoVendedor In colTiposVendedores
                    If UCase(Trim(tTipoVendedor.sDescricao)) = UCase(Trim(objTipoVendedor.sDescricao)) Then objVendedor.iTipo = objTipoVendedor.iCodigo
                Next

            End If

            'Se não encontrou código para o tipo de vendedor => grava erro na tabela de log
            If objVendedor.iTipo = 0 Then Print #1, "O vendedor " & objVendedor.iCodigo & " não foi vinculado a nenhum tipo de vendedor."

            'Guarda os percentuais de comissão. Está forçando que a comissão seja toda na baixa
            objVendedor.dPercComissao = .dPercComissao / 100
            objVendedor.dPercComissaoEmissao = 0
            objVendedor.dPercComissaoBaixa = 1
            
            'Guarda o código do banco
            objVendedor.iBanco = tVendedor.iBanco
            
            'Se a agência foi preenchida
            If Len(Trim(.sAgencia)) > 0 Then
            
                'Se for maior que o padrão CORPORATOR =>
                If Len(Trim(.sAgencia)) > STRING_AGENCIA Then
                
                    'Indica no arquivo de log que a agência foi truncada
                    Print #1, "Vendedor: " & .iCodigo & "|Erro: a agência do vendedor foi alterada de " & .sAgencia & " para " & Mid(.sAgencia, 1, STRING_AGENCIA) & "."
                
                End If
                
                'Transfere a agência lida para o obj
                objVendedor.sAgencia = Mid(.sAgencia, 1, STRING_AGENCIA)
            
            End If

            'Se a conta-corrente foi preenchida
            If Len(Trim(.sContaCorrente)) > 0 Then
            
                'Se for maior que o padrão CORPORATOR =>
                If Len(Trim(.sContaCorrente)) > STRING_CONTA_CORRENTE Then
                
                    'Indica no arquivo de log que a C/C. foi truncada
                    Print #1, "Vendedor: " & .iCodigo & "|Erro: a conta-corrente do vendedor foi alterada de " & .sContaCorrente & " para " & Mid(.sContaCorrente, 1, STRING_CONTA_CORRENTE) & "."
                
                End If
                
                'Transfere a agência lida para o obj
                objVendedor.sContaCorrente = Mid(.sContaCorrente, 1, STRING_CONTA_CORRENTE)
            
            End If
            
            'Se o CGC foi preenchido
            If Len(Trim(.sCGC)) > 0 Then

                sCGCAux = ""

                Call Formata_String_Numero(.sCGC, sCGCAux)

'                'Verifica se é um CGC válido
'                lErro = Cgc_Critica(sCGCAux)
'
'                'Se o cgc não for válido
'                If lErro <> SUCESSO Then
'
'                    'Grava no arquivo de log o erro referente a essa transportadora
'                    Print #1, "Vendedor: " & .iCodigo & "|Erro: CGC inválido."
'
'                End If

                'Transfere o cgc lido para o obj
                'mesmo sendo inválido o CGC é gravado para facilitar a correção do mesmo
                .sCGC = sCGCAux

            'Se o CGC não foi preenchido
            Else

                'Grava no arquivo de log o erro referente a esse vendedor
                Print #1, "Vendedor: " & .iCodigo & "|Erro: CGC não preenchido."

            End If

            'se a inscrição estadual foi preenchida
            If Len(Trim(.sInscricaoEstadual)) > 0 Then
                objVendedor.sInscricaoEstadual = .sInscricaoEstadual
            
            'Se não foi preenchida
            Else
            
                'Grava no arquivo de log o erro referente a esse vendedor
                Print #1, "Vendedor: " & .iCodigo & "|Erro: Inscrição estadual não preenchida."
            
            End If
            
        End With

        'Move o endereço para a memória
        lErro = Move_Endereco_Memoria(tEndereco, objEndereco, objVendedor.iCodigo, "Vendedor")
        If lErro <> SUCESSO Then
            Print #1, "Ocorreu erro ao guardar na memória o endereço do cliente " & objVendedor.iCodigo & "."
        End If

        'Grava a transportadora e o endereço
        lErro = CF("Vendedor_Grava_EmTrans", objVendedor, objEndereco)
        If lErro <> SUCESSO Then
            Print #1, "Erro: não foi possível gravar o Vendedor com código " & objVendedor.iCodigo
            gError 1000
        End If

        'Busca a próxima transportadora na tabela TransportadoraOrigem
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then
            Print #1, "Erro: não foi possível ler a tabela VendedoresOrigem."
            gError 1000
        End If

    Loop

    'Executa o fechamento do Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Print #1, "Erro ao efetuar commit da importação de transportadoras."

    'Fecha o arquivo de log
    Close #1

    Importacao_Vendedores = SUCESSO

    Exit Function

Erro_Importacao_Vendedores:

    Importacao_Vendedores = gErr

    Select Case gErr

        Case 1000
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177308)

    End Select

    'Executa o fechamento do Comando
    'Executa o fechamento do Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Call Transacao_Rollback

    'Fecha o arquivo de log
    Close #1

End Function

Function TipoVendedor_Grava_Importacao_EmTrans(ByVal objTipoVendedor As ClassTipoVendedor, ByVal iArquivoLog As Integer) As Long
'Atualiza ou insere um novo registro na tabela TiposDeVendedor

Dim lErro As Long
Dim iCodigo As Integer
Dim iVendedor As Integer
Dim lComando As Long
Dim lComando1 As Long
Dim lComando2 As Long

On Error GoTo Erro_TipoVendedor_Grava_Importacao_EmTrans

    'Inicializa comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then
        Print #iArquivoLog, " Erro ao abrir comando para gravação do tipo de vendedor " & objTipoVendedor.iCodigo
        gError 1000
    End If

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then
        Print #iArquivoLog, " Erro ao abrir comando para gravação do tipo de vendedor " & objTipoVendedor.iCodigo
        gError 1000
    End If

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then
        Print #iArquivoLog, " Erro ao abrir comando para gravação do tipo de vendedor " & objTipoVendedor.iCodigo
        gError 1000
    End If

    'Pesquisa descrição dos outros Tipos de Vendedores no BD
    lErro = Comando_Executar(lComando, "SELECT Codigo FROM TiposDeVendedor WHERE Codigo <> ? AND Descricao = ?", iVendedor, objTipoVendedor.iCodigo, objTipoVendedor.sDescricao)
    If lErro <> AD_SQL_SUCESSO Then
        Print #iArquivoLog, " Erro na leitura da tabela TiposDeVendedor."
        gError 1000
    End If

    'Verifica resultado da pesquisa
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then
        Print #iArquivoLog, " Erro na leitura da tabela TiposDeVendedor."
        gError 1000
    End If

    'Encontrou TipoVendedor com essa descrição
    If lErro = AD_SQL_SUCESSO Then
        Print #iArquivoLog, "O tipo de vendedor " & objTipoVendedor.iCodigo & SEPARADOR & objTipoVendedor.sDescricao & " não foi gravado, pois já existe outro tipo com essa descrição."
    Else

        'Pesquisa Tipo de Vendedor no BD
        lErro = Comando_ExecutarPos(lComando1, "SELECT Codigo FROM TiposDeVendedor WHERE Codigo = ? ", 0, iCodigo, objTipoVendedor.iCodigo)
        If lErro <> AD_SQL_SUCESSO Then
            Print #iArquivoLog, " Erro na leitura da tabela TiposDeVendedor."
            gError 1000
        End If

        'Verifica resultado da pesquisa
        lErro = Comando_BuscarPrimeiro(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then
            Print #iArquivoLog, " Erro na leitura da tabela TiposDeVendedor."
            gError 1000
        End If

        If lErro = AD_SQL_SEM_DADOS Then
        'Tipo de Vendedor não pertence ao BD

            'Insere novo registro na tabela TiposDeVendedor
            lErro = Comando_Executar(lComando2, "INSERT INTO TiposDeVendedor (Codigo, Descricao, PercComissao, PercComissaoBaixa, PercComissaoEmissao, ComissaoSobreTotal, ComissaoFrete, ComissaoSeguro, ComissaoICM, ComissaoIPI) VALUES (?,?,?,?,?,?,?,?,?,?)", objTipoVendedor.iCodigo, objTipoVendedor.sDescricao, objTipoVendedor.dPercComissao, objTipoVendedor.dPercComissaoBaixa, objTipoVendedor.dPercComissaoEmissao, objTipoVendedor.iComissaoSobreTotal, objTipoVendedor.iComissaoFrete, objTipoVendedor.iComissaoSeguro, objTipoVendedor.iComissaoICM, objTipoVendedor.iComissaoIPI)
            If lErro <> AD_SQL_SUCESSO Then
                Print #iArquivoLog, " Erro ao gravar na tabela TiposDeVendedor. O tipo " & objTipoVendedor.iCodigo & SEPARADOR & objTipoVendedor.sDescricao & " não foi gravado."
                gError 1000
            End If

        Else
            Print #iArquivoLog, "O tipo " & objTipoVendedor.iCodigo & SEPARADOR & objTipoVendedor.sDescricao & " não foi gravado, pois já existe outro tipo com esse código."
        End If

    End If


    'Libera comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

    TipoVendedor_Grava_Importacao_EmTrans = SUCESSO

    Exit Function

Erro_TipoVendedor_Grava_Importacao_EmTrans:

    TipoVendedor_Grava_Importacao_EmTrans = gErr

    Select Case Err

        Case 1000
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177309)

    End Select

    Call Transacao_Rollback

    'Libera comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

    Exit Function

End Function

Private Function Move_Endereco_Memoria(tEndereco As typeEndereco, objEndereco As ClassEndereco, lCodigo As Long, sTextoAux As String) As Long

Dim sCEPAux As String

On Error GoTo Erro_Move_Endereco_Memoria

        'Instancia um novo obj para armazenar o endereco do cliente
        Set objEndereco = New ClassEndereco

        With tEndereco

            'Transfere os dados lidos para o objendereco

            objEndereco.sEndereco = .sEndereco

            objEndereco.sBairro = .sBairro
            
            Select Case .sCidade
            
                    Case "0urinhos", "Brasília", "Cariacica", "Cláudio", "Goiânia", "Itu", "JI-Paraná", "Nerópolis", "Riachão do Jacuipe"
                        objEndereco.sCidade = ""
                    
                    Case Else
                        objEndereco.sCidade = .sCidade
            End Select
            
            objEndereco.sSiglaEstado = .sSiglaEstado
            
            Call Formata_String_Numero(.sCEP, sCEPAux)
            
            objEndereco.sCEP = sCEPAux
            
            'Se o telefone lido for maior que o endereço máximo permitido no Corporator
            If Len(Trim(StringZ(.sTelefone1))) > STRING_TELEFONE Then

                'Indica no arquivo de log que o endereço foi truncado
                Print #1, sTextoAux & ": " & lCodigo & "|Erro: o telefone do " & sTextoAux & " foi alterado de " & .sTelefone1 & " para " & Mid(Trim(.sTelefone1), 1, STRING_TELEFONE) & "."

            End If

            objEndereco.sTelefone1 = Mid(Trim(.sTelefone1), 1, STRING_TELEFONE)
            
            'Se o telefone lido for maior que o endereço máximo permitido no Corporator
            If Len(Trim(StringZ(.sTelefone2))) > STRING_TELEFONE Then

                'Indica no arquivo de log que o endereço foi truncado
                Print #1, sTextoAux & ": " & lCodigo & "|Erro: o telefone do " & sTextoAux & " foi alterado de " & .sTelefone2 & " para " & Mid(Trim(.sTelefone2), 1, STRING_TELEFONE) & "."

            End If

            objEndereco.sTelefone2 = Mid(Trim(.sTelefone2), 1, STRING_TELEFONE)
            
            objEndereco.sFax = .sFax
            objEndereco.sContato = .sContato

            'Se o código do país não foi preenchido
            If .iCodigoPais = 0 Then

                'Coloca o endereço como sendo no Brasil
                objEndereco.iCodigoPais = PAIS_BRASIL

            'Se o país foi preenchido
            Else

                '??? tem que implementar depois a importação automática de país

            End If

        End With

    Move_Endereco_Memoria = SUCESSO

    Exit Function

Erro_Move_Endereco_Memoria:

    Move_Endereco_Memoria = gErr

End Function

Public Function Importacao_Clientes() As Long

Dim lErro As Long
Dim lTransacao As Long
Dim alComando(3) As Long
Dim iIndice As Integer
Dim tCliente As typeCliente
Dim tEndereco As typeEndereco
Dim tEndereco2 As typeEndereco
Dim tEndereco3 As typeEndereco
Dim sCPF As String, sCidadeAux As String
Dim sCPFAux As String
Dim sCGCAux As String
Dim objCliente As ClassCliente
Dim objFilialCliente As ClassFilialCliente
Dim objEndereco As ClassEndereco
Dim objEndereco2 As ClassEndereco
Dim objEndereco3 As ClassEndereco
Dim colEnderecos As Collection
Dim sArquivo As String
Dim ivTransportadora As Integer
Dim iTransportadora As Integer
Dim sItemCategoria As String
Dim objVendedor As ClassVendedor
Dim objFilialCliCategoria As ClassFilialCliCategoria
Dim sEndereco3 As String
Dim ivVendedor As Integer
Dim ivTipoCliente As Integer
Dim iFilialCliente As Integer
Dim lCodigo As Long
Dim lCidade As Long

Dim sGuia As String 'incluído exclusivamente para importação da Artmill... posteriormente poderá ser removido

Const STRING_CLIENTE_RAZAO_SOCIAL_USU = 100
Const STRING_CLIENTE_NOME_REDUZIDO_USU = 50
Const STRING_ENDERECO_USU = 100
Const STRING_BAIRRO_USU = 50
Const STRING_CEP_USU = STRING_CEP
Const STRING_CIDADE_USU = 50
Const STRING_TELEFONE_USU = 50
Const STRING_CGC_USU = 14
Const STRING_INSCR_EST_USU = 17
Const STRING_INSCR_MUN_USU = 18
Const STRING_CATEGORIACLIENTEITEM_ITEM_USU = 20

Const VENDEDOR_NUMERICO = 1
Const TIPO_NUMERICO = 1
Const TRANSPORTADORA_NUMERICO = 1
Const CIDADE_NUMERICO = 1

On Error GoTo Erro_Importacao_Clientes

    STRING_ENDERECO = 100
    STRING_BAIRRO = 50
    STRING_CIDADE = 50
    STRING_CLIENTE_RAZAO_SOCIAL = 100
    STRING_CLIENTE_NOME_REDUZIDO = 50
    
    sArquivo = App.Path & "\Cliente_Log_Importacao.txt"

    If Len(Dir(sArquivo)) > 0 Then Kill sArquivo

    'Arquivo de log
    Open sArquivo For Append As #1

     'Executa abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then
        Print #1, "Erro: não foi possível abrir transação para executar importação dos clientes."
        gError 1000
    End If

    'Executa a abertura do Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then
            Print #1, "Erro: não foi possível abrir comando para executar leitura da tabela ClienteOrigem."
            gError 1000
        End If
    Next

    'Inicializa as strings que serão utilizadas na leitura
    tCliente.sNome = String(255, 0)
    tCliente.sNomeReduzido = String(255, 0)
    tEndereco.sEndereco = String(255, 0)
    tEndereco.sBairro = String(255, 0)
    tEndereco.sCEP = String(255, 0)
    tEndereco.sCidade = String(255, 0)
    tEndereco.sSiglaEstado = String(255, 0)
    tEndereco.sTelefone1 = String(255, 0)
    tEndereco.sTelefone2 = String(255, 0)
    tEndereco.sFax = String(255, 0)
    tEndereco.sEmail = String(255, 0)
    tEndereco.sContato = String(255, 0)
    tEndereco2.sEndereco = String(255, 0)
    tEndereco2.sBairro = String(255, 0)
    tEndereco2.sCEP = String(255, 0)
    tEndereco2.sCidade = String(255, 0)
    tEndereco2.sSiglaEstado = String(255, 0)
    tEndereco2.sTelefone1 = String(255, 0)
    tEndereco2.sFax = String(255, 0)
    tEndereco2.sContato = String(255, 0)
    sEndereco3 = String(255, 0)
    tCliente.sCGC = String(255, 0)
    sCPF = String(255, 0)
    tCliente.sInscricaoSuframa = String(255, 0)
    tCliente.sInscricaoEstadual = String(255, 0)
    tCliente.sInscricaoMunicipal = String(255, 0)
    tCliente.sObservacao = String(255, 0)
    sItemCategoria = String(255, 0)
    sGuia = String(255, 0)

    'Lê os clientes da tabela Origem
    lErro = Comando_Executar(alComando(0), "SELECT Codigo, Nome, NomeReduzido, LimiteCredito, Endereco, Bairro, CEP, Cidade, SiglaEstado, Telefone1, Telefone2, Fax, Email, Contato, Endereco2, Bairro2, CEP2, Cidade2, SiglaEstado2, Telefone3, Endereco3, CGC, CPF, InscricaoSuframa, InscricaoEstadual, InscricaoMunicipal, Observacao, Transportadora, Vendedor, ItemCategoria, CondPagto, Tipo, CidadeID, Guia FROM ClienteOrigem WHERE NomeReduzido NOT IN (select nomereduzido from clienteorigem group by nomereduzido having count(*) > 1) ORDER BY CGC, Codigo", _
        tCliente.lCodigo, tCliente.sNome, tCliente.sNomeReduzido, tCliente.dLimiteCredito, tEndereco.sEndereco, tEndereco.sBairro, tEndereco.sCEP, tEndereco.sCidade, tEndereco.sSiglaEstado, tEndereco.sTelefone1, tEndereco.sTelefone2, tEndereco.sFax, tEndereco.sEmail, tEndereco.sContato, tEndereco2.sEndereco, tEndereco2.sBairro, tEndereco2.sCEP, tEndereco2.sCidade, tEndereco2.sSiglaEstado, tEndereco2.sTelefone1, sEndereco3, tCliente.sCGC, sCPF, tCliente.sInscricaoSuframa, tCliente.sInscricaoEstadual, tCliente.sInscricaoMunicipal, tCliente.sObservacao, ivTransportadora, ivVendedor, sItemCategoria, tCliente.iCondicaoPagto, ivTipoCliente, lCidade, sGuia)
    If lErro <> AD_SQL_SUCESSO Then
        Print #1, "Erro: não foi possível ler a tabela ClienteOrigem."
        gError 1000
    End If

    'Busca o primeiro cliente
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then
        Print #1, "Erro: não foi possível ler a tabela ClienteOrigem."
        gError 1000
    End If

    'Se não encontrou nenhum cliente
    If lErro = AD_SQL_SEM_DADOS Then Print #1, "Erro: nenhum cliente foi encontrado na tabela origem."

    iFilialCliente = 1

    'Enquanto houverem clientes na tabela origem
    Do While lErro = AD_SQL_SUCESSO

        'Instancia um novo obj
        Set objCliente = New ClassCliente

        With tCliente

            'Transfere o código lido para o obj
            objCliente.lCodigo = .lCodigo

            objCliente.dLimiteCredito = .dLimiteCredito

            'Se a condição pagto for diferente de zero, significa que é a vista
            If .iCondicaoPagto <> 0 Then objCliente.iCondicaoPagto = COD_A_VISTA

            'Guarda o tipo do cliente
            objCliente.iTipo = ivTipoCliente

            'Se a categoria está preenchida
            If Len(Trim(sItemCategoria)) > 0 Then

                'Instancia um novo obj para categoria
                Set objFilialCliCategoria = New ClassFilialCliCategoria

                'Guarda os dados da categoria
                objFilialCliCategoria.sCategoria = "Situação"
                objFilialCliCategoria.sItem = sItemCategoria

                objCliente.colCategoriaItem.Add objFilialCliCategoria

            End If

            'Se o nome não foi informado na tabela origem
            If Len(Trim(.sNome)) = 0 Then

                'Monta um nome fictício para a transportadora
                objCliente.sRazaoSocial = "Cliente " & CStr(.lCodigo)

                'Grava no arquivo de log o erro referente a essa transportadora
                Print #1, "Cliente: " & .lCodigo & "|Erro: nome não preenchido. O registro foi gravado com nome " & objCliente.sRazaoSocial

            'Se o nome está preenchido
            Else

                'Se o nome for maior que 40 =>
                If Len(Trim(.sNome)) > STRING_CLIENTE_RAZAO_SOCIAL Then

                    'Indica no arquivo de log que o nome do cliente foi truncado
                    Print #1, "Cliente: " & .lCodigo & "|Erro: a razão social do cliente foi alterada de " & .sNome & " para " & Mid(.sNome, 1, STRING_CLIENTE_RAZAO_SOCIAL) & "."

                End If

                'Transfere o nome lido para o obj
                objCliente.sRazaoSocial = Mid(Trim(.sNome), 1, STRING_CLIENTE_RAZAO_SOCIAL)

            End If

            'Se o nome reduzido não foi preenchido
            If Len(Trim(.sNomeReduzido)) = 0 Then

                'Monta um nome reduzido fictício para a transportadora
                objCliente.sNomeReduzido = "Cliente" & CStr(.lCodigo)

                'Grava no arquivo de log o erro referente a essa transportadora
                Print #1, "Cliente: " & .lCodigo & "|Erro: nome reduzido não preenchido. O registro foi gravado com nome reduzido " & objCliente.sNomeReduzido

            'Se o nome reduzido está preenchido
            Else

                'Se o nome reduzido começar com um caracter numérico => alerta que é inválido
                If IsNumeric(Mid(.sNomeReduzido, 1, 1)) Then Print #1, "Cliente: " & .lCodigo & "|Erro: nome reduzido inválido, pois nome reduzido não pode começar com número."

                'Se o nome for maior que 20 =>
                If Len(Trim(.sNomeReduzido)) > STRING_CLIENTE_NOME_REDUZIDO Then

                    'Indica no arquivo de log que o nome do cliente foi truncado
                    Print #1, "Cliente: " & .lCodigo & "|Erro: o nome reduzido do cliente foi alterado de " & .sNomeReduzido & " para " & Mid(.sNomeReduzido, 1, STRING_CLIENTE_NOME_REDUZIDO) & "."

                End If

                'Transfere o nome reduzido lido para o obj
                objCliente.sNomeReduzido = Mid(Trim(.sNomeReduzido), 1, STRING_CLIENTE_NOME_REDUZIDO)

            End If

            'Se o CGC não foi preenchido
            If Len(Trim(.sCGC)) = 0 Then

                'Se o CPF também não foi preenchido
                If Len(Trim(sCPF)) = 0 Then

                    'Grava no arquivo de log o erro referente a essa transportadora
                    Print #1, "Cliente: " & .lCodigo & "|Erro: CGC e CPF não preenchidos."

                'Se o CPF foi preenchido
                Else
                    
                    Call Formata_String_Numero(sCPF, sCPFAux)
                    
'                    'Verifica se é um cpf válido
'                    lErro = Cgc_Critica(sCPFAux)
'
'                    'Se o cpf não for válido
'                    If lErro <> SUCESSO Then
'
'                        'Grava no arquivo de log o erro referente a essa transportadora
'                        Print #1, "Cliente: " & .lCodigo & "|Erro: CPF " & sCPF & " é inválido."
'
'                    End If

                    'Transfere o cpf lido para o obj
                    'mesmo sendo inválido o CPF é gravado para facilitar a correção do mesmo
                    objCliente.sCGC = sCPFAux

                End If

            'Se o CGC foi preenchido
            Else

                Call Formata_String_Numero(.sCGC, sCGCAux)
                
'                'Verifica se é um CGC válido
'                lErro = Cgc_Critica(sCGCAux)
'
'                'Se o cgc não for válido
'                If lErro <> SUCESSO Then
'
'                    'Grava no arquivo de log o erro referente a essa transportadora
'                    Print #1, "Cliente: " & .lCodigo & "|Erro: CGC inválido."
'
'                End If
                
                'Transfere o cgc lido para o obj
                'mesmo sendo inválido o CGC é gravado para facilitar a correção do mesmo
                objCliente.sCGC = sCGCAux
                
            End If
            
            'Se a inscrição estadual foi preenchida
            If Len(Trim(.sInscricaoEstadual)) > 0 Then
                
                'se a inscrição estadual é maior do que o padrão do CORPORATOR
                If Len(Trim(.sInscricaoEstadual)) > STRING_INSCR_EST Then
                
                    'Indica no arquivo de log que a inscrição estadual do cliente foi truncada
                    Print #1, "Cliente: " & .lCodigo & "|Erro: a inscrição estadual do cliente foi alterada de " & .sInscricaoEstadual & " para " & Mid(.sInscricaoEstadual, 1, STRING_INSCR_EST) & "."
                
                End If
                
'                'Guarda a inscrição estadual na memória
'                .sInscricaoEstadual = Mid(.sInscricaoEstadual, 1, STRING_INSCR_EST)
            
            'Senão
            Else
            
                'Grava no arquivo de log o erro referente a esse cliente
                Print #1, "Cliente: " & .lCodigo & "|Erro: Inscrição Estadual não preenchida."
            
            End If
                    
            'Se a inscrição estadual foi preenchida
            If Len(Trim(.sInscricaoSuframa)) > 0 Then
                
                'se a inscrição estadual é maior do que o padrão do CORPORATOR
                If Len(Trim(.sInscricaoSuframa)) > STRING_INSCR_SUF Then
                
                    'Indica no arquivo de log que a inscrição estadual do cliente foi truncada
                    Print #1, "Cliente: " & .lCodigo & "|Erro: a Inscrição Suframa do cliente foi alterada de " & .sInscricaoSuframa & " para " & Mid(.sInscricaoSuframa, 1, STRING_INSCR_SUF) & "."
                
                End If
                
'                'Guarda a inscrição estadual na memória
'                .sInscricaoSuframa = Mid(.sInscricaoSuframa, 1, STRING_INSCR_SUF)
            
            'Senão
            Else
            
                'Grava no arquivo de log o erro referente a esse cliente
                Print #1, "Cliente: " & .lCodigo & "|Erro: Inscrição Suframa não preenchida."
            
            End If
                    
                'Guarda o código do vendedor em objCliente
            objCliente.iVendedor = ivVendedor

            '***********************************************************************

            '************ Tratamento para obter a transportadora associada ************
            .iCodTransportadora = ivTransportadora
                
            '************ Tratamento para obter a cidade associada ************
            If lCidade > 0 And CIDADE_NUMERICO = 1 Then
            
                sCidadeAux = String(255, 0)
                
                'Tenta ler a cidade, para obter a descrição
                lErro = Comando_Executar(alComando(2), "SELECT Descricao FROM Cidades WHERE Codigo=?", sCidadeAux, CInt(lCidade))
                If lErro <> AD_SQL_SUCESSO Then Print #1, "Cliente: " & objCliente.lCodigo & "|Erro: não conseguiu ler a cidade do cliente."
            
                lErro = Comando_BuscarPrimeiro(alComando(2))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Print #1, "Cliente: " & objCliente.lCodigo & "|Erro: não conseguiu ler a transportadora do cliente."

                'Se não encontrou a cidade
                If lErro <> SUCESSO Then

                    sCidadeAux = ""
                    
                    'Grava erro no arquivo de log
                    Print #1, "Cliente: " & objCliente.lCodigo & "|Erro: a cidade do cliente não foi encontrada."

                End If
            
            Else
            
                sCidadeAux = ""
                
            End If
            '***********************************************************************
            
            'Guarda a guia no obj
            objCliente.sGuia = sGuia
            
        End With

        Set colEnderecos = New Collection
        
        'Move o endereço para a memória
        lErro = Move_Endereco_Memoria(tEndereco, objEndereco, objCliente.lCodigo, "Cliente")
        If lErro <> SUCESSO Then
            Print #1, "Ocorreu erro ao guardar na memória o endereço do cliente " & objCliente.lCodigo & "."
        End If

        objEndereco.sCidade = sCidadeAux

        'Guarda o endereço na coleção
        colEnderecos.Add objEndereco

        'Move o endereço de cobrança para a memória
        lErro = Move_Endereco_Memoria(tEndereco2, objEndereco2, objCliente.lCodigo, "Cliente")
        If lErro <> SUCESSO Then
            Print #1, "Ocorreu erro ao guardar na memória o endereço de cobrança do cliente " & objCliente.lCodigo & "."
        End If

        objEndereco2.sCidade = sCidadeAux
        
        'Guarda o endereço de cobrança na coleção
        colEnderecos.Add objEndereco2

        'Move o endereço de entrega para a memória
        lErro = Move_Endereco_Memoria(tEndereco3, objEndereco3, objCliente.lCodigo, "Cliente")
        If lErro <> SUCESSO Then
            Print #1, "Ocorreu erro ao guardar na memória o endereço do cliente " & objCliente.lCodigo & "."
        End If

        objEndereco3.sCidade = sCidadeAux
        
        'Guarda o endereço na coleção
        colEnderecos.Add objEndereco3
        
        If sCGCAux <> "" Then
        
            'Lê em Filial Cliente algum Cliente com o Mesmo CGC ou CPF
            lErro = Comando_Executar(alComando(3), "SELECT CodCliente FROM FiliaisClientes WHERE CGC LIKE ? AND CodCliente <> ? AND CGC<>''", lCodigo, Mid(sCGCAux, 1, 8) & "%", tCliente.lCodigo)
            If lErro <> AD_SQL_SUCESSO Then
                Print #1, "Ocorreu erro durante a leitura da tabela FiliaisClientes."
            End If
    
            lErro = Comando_BuscarPrimeiro(alComando(3))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then
                Print #1, "Ocorreu erro durante a leitura da tabela FiliaisClientes."
            End If
        
        End If
        
        'Se encontrou algum cliente com a mesma raiz de CGC
        If sCGCAux <> "" And lErro = AD_SQL_SUCESSO Then

            objCliente.lCodigo = lCodigo
            
            iFilialCliente = iFilialCliente + 1

            'instancia um novo objfilial
            Set objFilialCliente = New ClassFilialCliente

            'Transfere os dados para um objfilial
            lErro = Move_objCliente_objFilial(objCliente, objFilialCliente)
            If lErro <> SUCESSO Then
                Print #1, "Ocorreu erro ao tentar criar os dados para uma nova filial do cliente " & objCliente.lCodigo & "."
            End If

            objFilialCliente.iCodFilial = iFilialCliente
            objFilialCliente.sNome = objFilialCliente.sNome & iFilialCliente

'            MsgBox ("filcli:" & CStr(objFilialCliente.lCodCliente) & "/" & CStr(objFilialCliente.iCodFilial))
            
            'Insere a nova filial de cliente
            lErro = CF("FiliaisClientes_Grava_EmTrans", objFilialCliente, colEnderecos)
            If lErro <> SUCESSO Then
                Print #1, "|Erro: não foi possível gravar o Cliente com código " & objCliente.lCodigo
                gError 1000
            End If

        Else

'            MsgBox ("cliente: " & CStr(objCliente.lCodigo))
                    
            
            iFilialCliente = 1

            'Grava o cliente e o endereço
            lErro = CF("Cliente_Grava_EmTrans", objCliente, colEnderecos)
            If lErro <> SUCESSO Then
                Print #1, "Cliente: " & objCliente.lCodigo & "|Erro: não foi possível gravar o Cliente."
                gError 1000
            End If

        End If

        'Busca o próximo cliente na tabela ClienteOrigem
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then
            Print #1, "Erro: não foi possível ler a tabela ClienteOrigem."
            gError 1000
        End If
        
        sCGCAux = ""

    Loop

    'Executa o fechamento do Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Print #1, " Erro ao efetuar commit da importação de clientes."

    'Fecha o arquivo de log
    Close #1

    Importacao_Clientes = SUCESSO

    Exit Function

Erro_Importacao_Clientes:

    Importacao_Clientes = gErr

    Select Case gErr

        Case 1000
            MsgBox ("erro 1000")
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177310)

    End Select

    'Executa o fechamento do Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Call Transacao_Rollback

    'Fecha o arquivo de log
    Close #1

End Function

Private Function Move_objCliente_objFilial(objCliente As ClassCliente, objFilialCliente As ClassFilialCliente) As Long

    With objFilialCliente

        .lCodCliente = objCliente.lCodigo
        .iCodTransportadora = objCliente.iCodTransportadora
        .iVendedor = objCliente.iVendedor
        .sCGC = objCliente.sCGC
        .sInscricaoEstadual = objCliente.sInscricaoEstadual
        .sInscricaoMunicipal = objCliente.sInscricaoMunicipal
        .sNome = "Filial "
        .sObservacao = objCliente.sObservacao
        .sGuia = objCliente.sGuia

    End With

End Function



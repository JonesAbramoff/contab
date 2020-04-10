VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassImpExpNFPhar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const TIPOREG_NF = "N"
Const TIPOREG_ITEMNF = "I"
Const TIPOREG_RASTROITEM = "L"
Const ALMOXARIFADO_NF_DEVOLUCAO_PHAR = 1
Const FORNECEDOR_PHAR = 44 'KEEPERS
Const FILIALFORNECEDOR_PHAR = 1
Const LOTE_NAO_INFORMADO = "NAO INFORMADO"
Const SERIE_PHAR = "0"
Const TIPO_TRIBUTACAO_PHAR = 33 'SEM TRIBUTAÇÃO

Dim NATUREZAOP_INTERNA_PHAR As String

Function NFiscais_Importar(ByVal sNomeArquivo As String) As Long

Dim lErro As Long
Dim lTransacao As Long
Dim sTipoRegistro As String
Dim sRegistro As String
Dim sNomeArquivoCritica As String
Dim dValorProdutos As Double
Dim dTotalRastreado As Double
Dim vVariavel As Variant
Dim iLinha As Integer
Dim iLinhaNF As Integer
Dim iLinhaItemNF As Integer
Dim objNFiscal As New ClassNFiscal
Dim objItemNF As New ClassItemNF
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objTipoDocInfo As ClassTipoDocInfo
Dim objRastroItemNF As ClassRastroItemNF
Dim colNFs As New Collection 'Coleção de Notas Geradas
Dim colCriticas As New Collection 'Coleção de Críticas
Dim colMsgNotas As New Collection 'Coleção de Notas não Gravadas
Dim colMsgNotasG As New Collection 'Coleção de Notas Gravadas
Dim objTribTab As ClassTribTab

On Error GoTo Erro_NFiscais_Importar
    
    'Abre a transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 177314
   
    lErro = Conexao_DesativarLocks(DESATIVAR_LOCKS)
    If lErro <> SUCESSO Then gError 177410
   
    iLinha = 0
   
    'Cria nome de arquivo de Log com base no nome do arquivo de Importação
    sNomeArquivoCritica = left(sNomeArquivo, InStr(1, sNomeArquivo, ".") - 1) & "_LOG.txt"
   
    'Abre aquivo de importação para leitura
    Open sNomeArquivo For Input As #1
    
    'Cria/abre arquivo de log para gravação
    Open sNomeArquivoCritica For Output As #2
    
    'Cabeçalho do arquivo de Log
    Print #2, "Log de Importação do arquivo " & sNomeArquivo
    Print #2, "Empresa: " & gsNomeEmpresa & ", Filial: " & giFilialEmpresa & SEPARADOR & gsNomeFilialEmpresa
    Print #2, "DATA: " & CStr(Now)
    Print #2, ""
    
    'Tenta ler o almoxarifado utilizado na distribuição (SEMPRE O MESMO)
    objAlmoxarifado.iCodigo = ALMOXARIFADO_NF_DEVOLUCAO_PHAR
    objAlmoxarifado.iFilialEmpresa = giFilialEmpresa

    lErro = CF("Almoxarifado_Le", objAlmoxarifado)
    If lErro <> SUCESSO And lErro <> 25056 Then gError 177315

    'Enquanto não chegar ao final do arquivo de importação
    Do While Not EOF(1)
    
        'Obtém a Primeita Linha
        Line Input #1, sRegistro
    
        iLinha = iLinha + 1

        'O primeiro caracter identifica o tipo de registro
        sTipoRegistro = Mid(sRegistro, 1, 1)
        
        Select Case sTipoRegistro
        
            'Se for Nota Fiscal
            Case TIPOREG_NF
            
                'Obtém o Valor dos Produtos da NF anterior
                dValorProdutos = 0
                For Each objItemNF In objNFiscal.ColItensNF
                    dValorProdutos = dValorProdutos + (objItemNF.dPrecoUnitario * objItemNF.dQuantidade)
                Next
                
                objNFiscal.dValorProdutos = dValorProdutos
                
                'Verifica se o total dos produtos é igual ao total da Nota (NÃO TEM DESCONTO, FRETE, ETC ...)
                If Abs(dValorProdutos - objNFiscal.dValorTotal) > DELTA_VALORMONETARIO Then
                    colCriticas.Add "LINHA " & CStr(iLinhaNF) & ". O Somatório dos Valores dos Itens " & Format(dValorProdutos, "STANDARD") & " é diferente do Valor Total da Nota Fiscal " & Format(objNFiscal.dValorTotal, "STANDARD") & "."
                End If
            
                'NF Nova
                iLinhaNF = iLinha
            
                Set objNFiscal = New ClassNFiscal
                
                'Preenche o objNFiscal com base nos registros da Linha
                Call Preenche_NFiscal(colNFs, objNFiscal, sRegistro, iLinha, colCriticas)
        
            'Se for um item de Nota Fiscal
            Case TIPOREG_ITEMNF
            
                'Obtém o total rastreado
                dTotalRastreado = 0
                For Each objRastroItemNF In objItemNF.colRastreamento
                    dTotalRastreado = dTotalRastreado + objRastroItemNF.dLoteQdtAlocada
                Next
                
                'Verifica se a quantidade rastreada é superior a quantidade do almoxarifado
                If dTotalRastreado - objItemNF.dQuantidade > QTDE_ESTOQUE_DELTA Then
                    colCriticas.Add "LINHA " & CStr(iLinhaItemNF) & ". A Quantidade Rastreada " & Formata_Estoque(dTotalRastreado) & " é superior a quantidade do Almoxarifado " & Formata_Estoque(objItemNF.dQuantidade) & "."
                End If
                
                'Item Novo
                iLinhaItemNF = iLinha
            
                Set objItemNF = New ClassItemNF
        
                'Preenche o objItemNF com base nos registros da Linha
                Call Preenche_ItemNFiscal(objNFiscal, objItemNF, objAlmoxarifado, sRegistro, iLinha, colCriticas)
        
            'Se for um rastreamento do Item
            Case TIPOREG_RASTROITEM
        
                'Preenche o rastreamento do item com base nos registros da Linha
                'DESABILITAR ATÉ ELES COMEÇAREM A UTILIZAR CORRETAMENTE.
                Call Preenche_Rastreamento(objNFiscal, objItemNF, objAlmoxarifado, sRegistro, iLinha, colCriticas)
                
            'Se esse tipo não está sendo tratado
            Case Else
                colCriticas.Add "LINHA " & CStr(iLinha) & ".Não existe tratamento para o tipo de registro " & sTipoRegistro & "."
        
        End Select
        
    Loop
    
    'Obtém o Valor dos Produtos da NF anterior
    dValorProdutos = 0
    For Each objItemNF In objNFiscal.ColItensNF
        dValorProdutos = dValorProdutos + Arredonda_Moeda(objItemNF.dPrecoUnitario * objItemNF.dQuantidade)
    Next
    
    objNFiscal.dValorProdutos = dValorProdutos
    
    'Verifica se o total dos produtos é igual ao total da Nota (NÃO TEM DESCONTO, FRETE, ETC ...)
    If Abs(dValorProdutos - objNFiscal.dValorTotal) > DELTA_VALORMONETARIO Then
        colCriticas.Add "LINHA " & CStr(iLinhaNF) & ".O Somatório dos Itens " & Format(dValorProdutos, "STANDARD") & " é diferente do Valor Total da Nota Fiscal " & Format(objNFiscal.dValorTotal, "STANDARD") & "."
    End If
                
    'Obtém o total rastreado
    dTotalRastreado = 0
    For Each objRastroItemNF In objItemNF.colRastreamento
        dTotalRastreado = dTotalRastreado + objRastroItemNF.dLoteQdtAlocada
    Next
    
    'Verifica se a quantidade rastreada é superior a quantidade do almoxarifado
    If dTotalRastreado - objItemNF.dQuantidade > QTDE_ESTOQUE_DELTA Then
        colCriticas.Add "LINHA " & CStr(iLinhaItemNF) & ". A Quantidade Rastreada " & Formata_Estoque(dTotalRastreado) & " é superior a quantidade do Almoxarifado " & Formata_Estoque(objItemNF.dQuantidade) & "."
    End If
                
    'Se teve erro no arquivo
    If colCriticas.Count <> 0 Then
    
        'Imprime os erros no arquivo de LOG
        For Each vVariavel In colCriticas
            Print #2, vVariavel
        Next
        
        Print #2, ""
        Print #2, "Não foi possível importar o arquivo."
       
        'Sai com erro na geração
        gError 177316
    
    End If
    
    'Para cada Nota Gerada com sucesso
    For Each objNFiscal In colNFs

        Set objTipoDocInfo = New ClassTipoDocInfo
        
        objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
        
        'Lê o Tipo de Documento
        lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
        If lErro <> SUCESSO And lErro <> 31415 Then gError 177324
        
        'pesquisa a existencia da nota fiscal com os dados passados como parametro e não exibe mensagens
        lErro = CF("NFiscalEntrada_Verifica_Existencia2", objNFiscal, objTipoDocInfo, True)
        If lErro <> SUCESSO And lErro <> 61414 And lErro <> 89723 Then gError 177325
                      
        'Se for uma nota nova
        If lErro = SUCESSO Then
        
            'Calcula o Peso Líquido e Bruto
            Call NFiscal_Peso_Calcula(objNFiscal)
            
            Set objTribTab = New ClassTribTab
            
            lErro = objTribTab.TributacaoNF_Reset(objNFiscal)
            If lErro <> SUCESSO Then gError 177325
            
            'obtem dados do contrato
            Call objTribTab.TipoNFiscal_Definir(objTipoDocInfo.iCodigo, objTipoDocInfo.sSigla)
            Call objTribTab.DataEmissao_Alterada(objNFiscal.dtDataEmissao)
            Call objTribTab.Serie_Alterada(objNFiscal.sSerie)
            
            lErro = gobjTributacao.AtualizaImpostos(objTribTab.mvarobjNFTributacao, 0)
            If lErro <> SUCESSO Then gError 177325
            
            lErro = objTribTab.PreencherNF(objNFiscal)
            If lErro <> SUCESSO Then gError 177325
            
            Set objTribTab = Nothing
            
            objNFiscal.sNaturezaOp = objNFiscal.objTributacaoNF.sNaturezaOpInterna
        
            'Grava a Nota em Transação
            lErro = CF("NFiscalEntradaDevolucao_Grava_EmTrans", objNFiscal, Nothing)
            If lErro <> SUCESSO Then gError 177317 'Não pode continuar sem dar RollBack
            
            'Adiciona a coleção de Notas Gravadas
            colMsgNotasG.Add "Série: " & objNFiscal.sSerie & ", Número: " & CStr(objNFiscal.lNumNotaFiscal) & ", Data Emissão: " & Format(objNFiscal.dtDataEmissao, "dd/mm/yyyy")
        
        Else 'Se a nota já existe no sistema
            'Adiciona a coleção de Notas não Gravadas
            colMsgNotas.Add "Série: " & objNFiscal.sSerie & ", Número: " & CStr(objNFiscal.lNumNotaFiscal) & ", Data Emissão: " & Format(objNFiscal.dtDataEmissao, "dd/mm/yyyy") & ", Motivo da não Gravação: Já Existe no Sistema."
        End If
        
    Next
    
    'Rodapé do Log
    Print #2, "Importação feita com sucesso."
    Print #2, ""
    Print #2, "Total de Notas Gravadas = " & CStr(colMsgNotasG.Count)
    
    'Grava no Log a lista de Notas Gravadas com sucesso
    For Each vVariavel In colMsgNotasG
        Print #2, vVariavel
    Next
    
    Print #2, ""
    Print #2, "Total de Notas não Gravadas = " & CStr(colMsgNotas.Count)
    
    'Grava no Log a lista de Notas que não foram Gravadas
    For Each vVariavel In colMsgNotas
        Print #2, vVariavel
    Next
    
    lErro = Conexao_DesativarLocks(REATIVAR_LOCKS)
    If lErro <> SUCESSO Then gError 177411
    
    'COMMIT
    lErro = Transacao_Commit()
    If lErro <> SUCESSO Then gError 177318
    
    'Fecha o arquivo de Importação
    Close #1
    
    'Fecha o arquivo de Log
    Close #2
    
    'Avisa que a rotina foi finalizada com sucesso
    Call Rotina_Aviso(vbOKOnly, "AVISO_IMPORTACAO_SUCESSO", sNomeArquivoCritica)
    
    NFiscais_Importar = SUCESSO
    
    Exit Function
    
Erro_NFiscais_Importar:

    NFiscais_Importar = gErr

    Select Case gErr
    
        Case 177314
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 177315, 177317, 177324, 177325
        
        Case 177316
            Call Rotina_Erro(vbOKOnly, "ERRO_IMPORTACAO_NF", gErr, sNomeArquivoCritica)
        
        Case 177318
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT_TRANSACAO", gErr)
        
        Case 177410
            Call Rotina_Erro(vbOKOnly, "ERRO_DESATIVACAO_LOCKS", gErr)
             
        Case 177411
            Call Rotina_Erro(vbOKOnly, "ERRO_REATIVACAO_LOCKS", gErr)
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 177319)
    
    End Select
    
    Call Conexao_DesativarLocks(REATIVAR_LOCKS)
    
    'Desfaz a transação
    Call Transacao_Rollback
    
    'Fecha o arquivo de Importação
    Close #1
    
    'Fecha o arquivo de Log
    Close #2
    
    Exit Function

End Function

Function Preenche_NFiscal(ByVal colNFs As Collection, ByVal objNF As ClassNFiscal, ByVal sRegistro As String, ByVal iLinha As Integer, ByVal colCriticas As Collection) As Long
'Preenche o objNF

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_Preenche_NFiscal

    'Obtém dados fixos
    objNF.iFilialEmpresa = giFilialEmpresa
    objNF.iTipoNFiscal = DOCINFO_NFEEDSO 'DOCINFO_NFEEODFSE
    objNF.iTipoDocInfo = TIPODOCINFO_TIPO_NFEXT
    objNF.iNaoVerificaCredito = NAO_VERIFICA_CREDITO_CLIENTE
    objNF.iStatus = STATUS_BAIXADO
    objNF.dtDataReferencia = DATA_NULA
    objNF.iFreteRespons = FRETE_EMITENTE
    objNF.dtDataRegistro = gdtDataAtual
    objNF.dtDataSaida = DATA_NULA
    objNF.dtDataVencimento = DATA_NULA
    objNF.sSerie = SERIE_PHAR
    
    'Obtém dados do Registro
    objNF.lNumNotaFiscal = StrParaLong(Mid(sRegistro, 5, 6))
    
    objFornecedor.sNomeReduzido = Trim(Mid(sRegistro, 11, 7))
    
    lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
    If lErro <> SUCESSO And lErro <> 6681 Then
        'Avisa que teve erro na leitura do fornecedor
        colCriticas.Add "LINHA " & CStr(iLinha) & ".Erro na leitura da tabela de Fornecedores."
    End If
    If lErro = 6681 Then
        'Avisa que o fornecedor não está cadastrado
        colCriticas.Add "LINHA " & CStr(iLinha) & ".O Fornecedor com nome reduzido " & objFornecedor.sNomeReduzido & " não está cadastrado."
    End If
    
    objNF.lFornecedor = objFornecedor.lCodigo
    objNF.iFilialForn = StrParaInt(Mid(sRegistro, 18, 3))
    objNF.dtDataEmissao = StrParaDate(Mid(sRegistro, 21, 2) & "/" & Mid(sRegistro, 23, 2) & "/" & Mid(sRegistro, 25, 4))
    objNF.dtDataEntrada = StrParaDate(Mid(sRegistro, 29, 2) & "/" & Mid(sRegistro, 31, 2) & "/" & Mid(sRegistro, 33, 4))
    objNF.sNaturezaOp = Mid(sRegistro, 37, 4)
    objNF.dValorTotal = StrParaDbl(Mid(sRegistro, 41, 11))
    
    Select Case objNF.sNaturezaOp
    
        Case "5902"
            NATUREZAOP_INTERNA_PHAR = "1902"
        
        Case "5906"
            NATUREZAOP_INTERNA_PHAR = "1906"
        
        Case "5907"
            NATUREZAOP_INTERNA_PHAR = "1907"
    
        Case Else
            NATUREZAOP_INTERNA_PHAR = "1949"
    End Select
    
    'Inicializa a parte tributária
    Call objNF.Inicializa_Tributacao
    Call objNF.objTributacaoNF.Coloca_Auto
    
    objNF.objTributacaoNF.iTipoTributacao = TIPO_TRIBUTACAO_PHAR
    objNF.objTributacaoNF.iNaturezaOpManual = VAR_PREENCH_MANUAL
    objNF.objTributacaoNF.sNaturezaOpInterna = NATUREZAOP_INTERNA_PHAR
    objNF.objTributacaoNF.iTipoTributacaoManual = VAR_PREENCH_MANUAL
    
'    objNF.objTributacaoNF.objTributacaoDesconto.iTipoTributacao = TIPO_TRIBUTACAO_PHAR
'    objNF.objTributacaoNF.objTributacaoDesconto.sNaturezaOp = NATUREZAOP_INTERNA_PHAR
'    objNF.objTributacaoNF.objTributacaoFrete.iTipoTributacao = TIPO_TRIBUTACAO_PHAR
'    objNF.objTributacaoNF.objTributacaoFrete.sNaturezaOp = NATUREZAOP_INTERNA_PHAR
'    objNF.objTributacaoNF.objTributacaoOutras.iTipoTributacao = TIPO_TRIBUTACAO_PHAR
'    objNF.objTributacaoNF.objTributacaoOutras.sNaturezaOp = NATUREZAOP_INTERNA_PHAR
'    objNF.objTributacaoNF.objTributacaoSeguro.iTipoTributacao = TIPO_TRIBUTACAO_PHAR
'    objNF.objTributacaoNF.objTributacaoSeguro.sNaturezaOp = NATUREZAOP_INTERNA_PHAR
    'XXXX
    
    'Inicializa o rastreamento
    Set objNF.objRastreamento = New ClassRastreamento
    
    'Adiciona a Nota a coleção de Notas geradas
    colNFs.Add objNF
 
    Preenche_NFiscal = SUCESSO
    
    Exit Function
    
Erro_Preenche_NFiscal:

    Preenche_NFiscal = gErr

    Select Case gErr
        
        Case Else
            colCriticas.Add "LINHA " & CStr(iLinha) & "." & Error$ & "."
    
    End Select
    
    Exit Function

End Function

Function Preenche_ItemNFiscal(ByVal objNF As ClassNFiscal, ByVal objItemNF As ClassItemNF, ByVal objAlmoxarifado As ClassAlmoxarifado, ByVal sRegistro As String, ByVal iLinha As Integer, ByVal colCriticas As Collection) As Long
'Preenche o objItemNF

Dim lErro As Long
Dim sProduto As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim objItemNFAlocacao As New ClassItemNFAlocacao
Dim dFator As Double

On Error GoTo Erro_Preenche_ItemNFiscal

    'Obtém dados fixos
    objItemNF.dtDataEntrega = DATA_NULA

    'Obtém dados do Registro
    objItemNF.iItem = StrParaInt(Mid(sRegistro, 11, 4))
    
    sProduto = Mid(sRegistro, 15, 27)
    
    lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 177323
    
    objProduto.sCodigo = sProdutoFormatado
    
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then
        'Avisa que houve erro na leitura do produto
        colCriticas.Add "LINHA " & CStr(iLinha) & ". Erro na formatação do produto" & sProduto & "."
    End If
    If lErro = 28030 Then
        'Avisa que o produto não está cadastrado
        colCriticas.Add "LINHA " & CStr(iLinha) & ". O Produto com código " & sProdutoFormatado & " não está cadastrado no sistema."
    End If
        
    objItemNF.sProduto = sProdutoFormatado
    objItemNF.sDescricaoItem = objProduto.sDescricao
'    objItemNF.sUnidadeMed = Trim(Mid(sRegistro, 42, 3))
    objItemNF.dQuantidade = Arredonda_Estoque(StrParaDbl(Mid(sRegistro, 45, 9)))
    objItemNF.dPrecoUnitario = StrParaDbl(Mid(sRegistro, 54, 12))
    
    objItemNF.sUnidadeMed = objProduto.sSiglaUMEstoque
    objItemNF.sUMEstoque = objProduto.sSiglaUMEstoque
    objItemNF.sUMVenda = objProduto.sSiglaUMVenda
    objItemNF.iClasseUM = objProduto.iClasseUM
    
    Select Case Trim(objItemNF.sProduto)

        Case "5033001", "5033105"

            objItemNF.sUnidadeMed = "UN"
            objItemNF.dQuantidade = objItemNF.dQuantidade * 12
            objItemNF.dPrecoUnitario = objItemNF.dPrecoUnitario / 12

        Case "5033012", "5033015", "5033018", "5033031"

            objItemNF.sUnidadeMed = "UN"
            objItemNF.dQuantidade = objItemNF.dQuantidade * 6
            objItemNF.dPrecoUnitario = objItemNF.dPrecoUnitario / 6

    End Select
    
     'Converte a unidade de medida de UM de Venda para a UM da linha do Grid
    lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objItemNF.sUnidadeMed, objProduto.sSiglaUMEstoque, dFator)
    If lErro <> SUCESSO Then gError 177321
    
    'Cria distribuição do Item
    Set objItemNFAlocacao = objItemNF.ColAlocacoes.Add(objAlmoxarifado.iCodigo, objAlmoxarifado.sNomeReduzido, Arredonda_Estoque(objItemNF.dQuantidade * dFator))
    
    objItemNFAlocacao.sUnidadeMed = objProduto.sSiglaUMEstoque
    
    objNF.ColItensNF.Add1 objItemNF
 
    'Inicializa parte tributária
    Call objItemNF.Inicializa_Tributacao
    Call objItemNF.objTributacaoItemNF.Coloca_Auto
    
    objItemNF.objTributacaoItemNF.iTipoTributacao = TIPO_TRIBUTACAO_PHAR
    objItemNF.objTributacaoItemNF.iNaturezaOpManual = VAR_PREENCH_MANUAL
    objItemNF.objTributacaoItemNF.sNaturezaOp = objNF.objTributacaoNF.sNaturezaOpInterna
    objItemNF.objTributacaoItemNF.iNaturezaOpManual = VAR_PREENCH_MANUAL
    
    objProduto.sCodigo = objItemNF.sProduto
    objItemNF.objTributacaoItemNF.sProduto = objProduto.sCodigo
    objItemNF.objTributacaoItemNF.iExTIPI = objProduto.iExTIPI
    objItemNF.objTributacaoItemNF.sGenero = objProduto.sGenero
    objItemNF.objTributacaoItemNF.iProdutoEspecifico = objProduto.iProdutoEspecifico
    objItemNF.objTributacaoItemNF.sUMTrib = objProduto.sSiglaUMTrib
    objItemNF.objTributacaoItemNF.sIPICodProduto = objProduto.sIPICodigo
    objItemNF.objTributacaoItemNF.dQtdTrib = objItemNF.dQuantidade
    objItemNF.objTributacaoItemNF.dValorUnitTrib = objItemNF.dPrecoUnitario
    objItemNF.objTributacaoItemNF.dDescontoGrid = objItemNF.dValorDesconto
    objItemNF.objTributacaoItemNF.dPrecoTotal = objItemNF.dValorTotal
    objItemNF.objTributacaoItemNF.sProdutoDescricao = objProduto.sDescricao
    objItemNF.objTributacaoItemNF.sISSQN = objProduto.sISSQN
        
    Preenche_ItemNFiscal = SUCESSO
    
    Exit Function
    
Erro_Preenche_ItemNFiscal:

    Preenche_ItemNFiscal = gErr

    Select Case gErr
    
        Case 177321
            colCriticas.Add "LINHA " & CStr(iLinha) & ". Não foi possível a conversão da unidade de medida " & objItemNF.sUnidadeMed & " para a unidade de medida de estoque " & objProduto.sSiglaUMEstoque & " do produto " & objProduto.sCodigo & ". Classe UM = ."
  
        Case Else
            colCriticas.Add "LINHA " & CStr(iLinha) & "." & Error$ & "."
    
    End Select
    
    Exit Function

End Function

Function Preenche_Rastreamento(ByVal objNF As ClassNFiscal, ByVal objItemNF As ClassItemNF, ByVal objAlmoxarifado As ClassAlmoxarifado, ByVal sRegistro As String, ByVal iLinha As Integer, ByVal colCriticas As Collection) As Long
'Preenche o objRastroItemNF em objItemNF

Dim lErro As Long
Dim objRastroItemNF As ClassRastroItemNF
Dim objRastroLote As New ClassRastreamentoLote

On Error GoTo Erro_Preenche_Rastreamento

    'Obtém os dados fixos
    Set objRastroItemNF = New ClassRastroItemNF

    objRastroItemNF.iAlmoxCodigo = objAlmoxarifado.iCodigo
    objRastroItemNF.dAlmoxQtdAlocada = objItemNF.dQuantidade
    objRastroItemNF.iItem = objItemNF.iItem
    objRastroItemNF.sProduto = objItemNF.sProduto
    objRastroItemNF.sSiglaUM = objItemNF.sUnidadeMed
 
    'Obtém os dados do registro
    objRastroItemNF.sLote = Trim(Mid(sRegistro, 15, 29))
    objRastroItemNF.dLoteQdtAlocada = StrParaDbl(Mid(sRegistro, 44, 9))
    objRastroItemNF.iLoteFilialOP = StrParaInt(Mid(sRegistro, 53, 1))
    
    'Se o lote não estiver marcado como não informado
    If UCase(Trim(objRastroItemNF.sLote)) <> LOTE_NAO_INFORMADO Then
    
        'Le o Lote
        Set objRastroLote = New ClassRastreamentoLote
        
        objRastroLote.sCodigo = objRastroItemNF.sLote
        objRastroLote.sProduto = objRastroItemNF.sProduto
        objRastroLote.iFilialOP = objRastroItemNF.iLoteFilialOP
    
        lErro = CF("RastreamentoLote_Le", objRastroLote)
        If lErro <> SUCESSO And lErro <> 75710 Then
            'Avisa que houve erro na leitura do Lote
            colCriticas.Add "LINHA " & CStr(iLinha) & ". Erro na Leitura do Lote."
        End If
        If lErro = 75710 Then
            'Avisa que o Lote não existe
            colCriticas.Add "LINHA " & CStr(iLinha) & ". O Lote " & objRastroItemNF.sLote & " não está cadastrado para o produto " & objRastroItemNF.sProduto & "."
        End If
                
        objRastroItemNF.dtLoteData = objRastroLote.dtDataEntrada
            
        'Verifica se o total do lote é menor que o total do almoxarifado
        If objRastroItemNF.dAlmoxQtdAlocada + QTDE_ESTOQUE_DELTA < objRastroItemNF.dLoteQdtAlocada Then
            colCriticas.Add "LINHA " & CStr(iLinha) & ". A quantidade alocada no Lote " & Formata_Estoque(objRastroItemNF.dLoteQdtAlocada) & " é maior que a alocada no almoxarifado " & Formata_Estoque(objRastroItemNF.dAlmoxQtdAlocada) & "."
        End If
    
        'Adiciona o rastreamento
        objItemNF.colRastreamento.Add objRastroItemNF
    End If
 
    Preenche_Rastreamento = SUCESSO
    
    Exit Function
    
Erro_Preenche_Rastreamento:

    Preenche_Rastreamento = gErr

    Select Case gErr
        
        Case Else
            colCriticas.Add "LINHA " & CStr(iLinha) & "." & Error$ & "."
    
    End Select
    
    Exit Function

End Function

Private Function NFiscal_Peso_Calcula(objNFiscal As ClassNFiscal)
'Calcula o Peso Líquido e Bruto de uma Nota Fiscal

Dim lErro As Long
Dim objProduto As ClassProduto
Dim objItemNF As ClassItemNF
Dim dPesoL As Double
Dim dPesoB As Double
Dim dFator As Double

On Error GoTo Erro_NFiscal_Peso_Calcula

    'Para cada item da Nota Fiscal
    For Each objItemNF In objNFiscal.ColItensNF
    
        'Le o produto
        Set objProduto = New ClassProduto
    
        objProduto.sCodigo = objItemNF.sProduto
    
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 177320
    
        'Converte a unidade de medida de UM de Venda para a UM da linha do Grid
        lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objItemNF.sUnidadeMed, objProduto.sSiglaUMVenda, dFator)
        If lErro <> SUCESSO Then gError 177321
    
        'Calcula o peso Líquido e o Peso Bruto
        dPesoL = dPesoL + (objProduto.dPesoLiq * objItemNF.dQuantidade * dFator)
        dPesoB = dPesoB + (objProduto.dPesoBruto * objItemNF.dQuantidade * dFator)
    
    Next
    
    objNFiscal.dPesoBruto = dPesoB
    objNFiscal.dPesoLiq = dPesoL
    
    NFiscal_Peso_Calcula = SUCESSO
    
    Exit Function

Erro_NFiscal_Peso_Calcula:

    NFiscal_Peso_Calcula = gErr

    Select Case gErr
    
        Case 177320, 177321

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177322)

    End Select
       
    Exit Function

End Function

Public Function Importa_Xml_NFe_Cust() As Long

Dim lErro As Long, iIndice As Integer
Dim lTransacao As Long
Dim alComando(0 To 12) As Long
Dim objNF As New ClassNFiscal
Dim objItemNF As New ClassItemNF
Dim objTipoDocInfo As ClassTipoDocInfo
Dim lForn As Long, iFilForn As Integer
Dim lNumNF As Long, sSerie As String, dtDataEmi As Date, sMsg(1 To 4) As String
Dim sMensagem As String, dTotal As Double, sChvNFe As String, sCFOP As String
Dim iItem As Integer, sProdXML As String, sProdDesc As String, sCFOPInterno As String
Dim sCFOPXML As String, sUMXML As String, dQtd As Double, dPreco As Double
Dim objProduto As ClassProduto, dTotalItens As Integer, sProduto As String
Dim objTribTab As New ClassTribTab, sNFs As String
Dim objItemNFAlocacao As New ClassItemNFAlocacao
Dim objAlmoxarifado As New ClassAlmoxarifado, dFator As Double, sEANXML As String, lNumIntNF As Long
Dim colArqsSep As Collection, sPedidos As String, vsNomeArq As Variant, sLoteXML As String

Const CGC_EMITENTE = "11239205000140"

On Error GoTo Erro_Importa_Xml_NFe_Cust

    'Imorta os arquivos de separação necessários para identificação dos lotes usados
    lErro = Importa_Arq_Sep_LMLog()
    If lErro <> AD_SQL_SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Abrir transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 209866

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 209867
    Next
    
    'Tenta ler o almoxarifado utilizado na distribuição (SEMPRE O MESMO)
    objAlmoxarifado.iCodigo = ALMOXARIFADO_NF_DEVOLUCAO_PHAR
    objAlmoxarifado.iFilialEmpresa = giFilialEmpresa

    lErro = CF("Almoxarifado_Le", objAlmoxarifado)
    If lErro <> SUCESSO And lErro <> 25056 Then gError ERRO_SEM_MENSAGEM
       
    'Busca o código do fornecedor
    lErro = Comando_Executar(alComando(0), "SELECT CodFornecedor, CodFilial FROM FiliaisFornecedores WHERE CGC = ?", lForn, iFilForn, CGC_EMITENTE)
    If lErro <> AD_SQL_SUCESSO Then gError 209868
    
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209869
    
    If lErro <> AD_SQL_SUCESSO Then gError 209870 ' Fornecedor não cadastrado
       
    sSerie = String(STRING_MAXIMO, 0)
    sChvNFe = String(STRING_MAXIMO, 0)
    For iIndice = 1 To 4
        sMsg(iIndice) = String(STRING_MAXIMO, 0)
    Next
    
    'Só traz as NFe de um determinado fornecedor para filial empresa corrente que
    'não tenham sido importadas anteriormente e que não possuam
    'nem tributação e nem nenhum outro valor que afete o total da NF que não seja os itens
    lErro = Comando_Executar(alComando(1), "SELECT ChvNFe, NumNF, Serie, DataEmissao, ValorTotalNF, SUBSTRING(Mensagem,1,250), SUBSTRING(Mensagem,251,250), SUBSTRING(Mensagem,501,250), SUBSTRING(Mensagem,751,250) FROM ImportNFeXml WHERE NumIntNF = 0 AND CGCEmitente = ? AND FilialEmpresa = ? AND ABS(ValorProdutos - ValorTotalNF) < 0.001 AND ICMSBase = 0 AND Mensagem LIKE '%REFERENTE AOS PEDIDOS:%'", _
                                                sChvNFe, lNumNF, sSerie, dtDataEmi, dTotal, sMsg(1), sMsg(2), sMsg(3), sMsg(4), CGC_EMITENTE, giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 209873
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209874
    
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objNF = New ClassNFiscal
    
        'Inicializa a parte tributária
        Call objNF.Inicializa_Tributacao
        Call objNF.objTributacaoNF.Coloca_Auto
    
        objNF.lNumNotaFiscal = lNumNF
        objNF.sSerie = sSerie & "-e"
        objNF.dtDataEmissao = dtDataEmi
        objNF.iNFe = MARCADO
        For iIndice = 1 To 4
            objNF.sMensagemNota = objNF.sMensagemNota & sMsg(iIndice)
        Next
        objNF.iFilialEmpresa = giFilialEmpresa
        objNF.iTipoNFiscal = DOCINFO_NFEEDSO 'DOCINFO_NFERS 'NF externa de entrada remessa simbólica
        objNF.iTipoDocInfo = TIPODOCINFO_TIPO_NFEXT
        objNF.iNaoVerificaCredito = NAO_VERIFICA_CREDITO_CLIENTE
        objNF.iStatus = STATUS_BAIXADO
        objNF.dtDataReferencia = DATA_NULA
        objNF.iFreteRespons = FRETE_EMITENTE
        objNF.dtDataRegistro = Date
        objNF.dtDataSaida = DATA_NULA
        objNF.dtDataVencimento = DATA_NULA
        objNF.lFornecedor = lForn
        objNF.iFilialForn = iFilForn
        objNF.dtDataEntrada = objNF.dtDataEmissao 'Date
        objNF.sChvNFe = sChvNFe
        
        objNF.dValorTotal = dTotal
        objNF.dValorProdutos = dTotal
        
        '=========================================
        'Identifica as separações que estão sendo devolvidas
        Set colArqsSep = New Collection
        
        sPedidos = Replace(Mid(objNF.sMensagemNota, InStr(1, objNF.sMensagemNota, "REFERENTE AOS PEDIDOS:") + Len("REFERENTE AOS PEDIDOS:")), "|", "")
        
        vsNomeArq = String(STRING_MAXIMO, 0)
        
        lErro = Comando_Executar(alComando(11), "SELECT DISTINCT NomeArq FROM PharLMLogSep WHERE NumPedido IN (" & sPedidos & ") AND NumIntNF = 0 ORDER BY NomeArq ", vsNomeArq)
        If lErro <> AD_SQL_SUCESSO Then gError 211056
        
        lErro = Comando_BuscarPrimeiro(alComando(11))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211057
        
        Do While lErro <> AD_SQL_SEM_DADOS
        
            colArqsSep.Add vsNomeArq
        
            lErro = Comando_BuscarProximo(alComando(11))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211057
        
        Loop
        '=========================================
        
        'Inicializa o rastreamento
        Set objNF.objRastreamento = New ClassRastreamento
        
        sProdXML = String(STRING_MAXIMO, 0)
        sProdDesc = String(STRING_MAXIMO, 0)
        sCFOPXML = String(STRING_MAXIMO, 0)
        sUMXML = String(STRING_MAXIMO, 0)
        sEANXML = String(STRING_MAXIMO, 0)
        sLoteXML = String(STRING_MAXIMO, 0)
    
        'Lê os itens da NF
        lErro = Comando_Executar(alComando(2), "SELECT Item, Produto, SUBSTRING(DescProd,1,250), CFOP, UM, Quantidade, PrecoUnitario, EAN, InfAdProd FROM ImportNFeItensXml WHERE ChvNFE = ? ORDER BY Item", _
            iItem, sProdXML, sProdDesc, sCFOPXML, sUMXML, dQtd, dPreco, sEANXML, sLoteXML, sChvNFe)
        If lErro <> AD_SQL_SUCESSO Then gError 209875
        
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209876
        
        sCFOP = ""
        sCFOPInterno = ""
        'Para cada item
        Do While lErro <> AD_SQL_SEM_DADOS
        
            Set objProduto = New ClassProduto
            Set objItemNF = New ClassItemNF
        
            'Inicializa parte tributária
            Call objItemNF.Inicializa_Tributacao
            Call objItemNF.objTributacaoItemNF.Coloca_Auto
            
            'Obtém dados fixos
            objItemNF.dtDataEntrega = DATA_NULA
            objItemNF.iAlmoxarifado = objAlmoxarifado.iCodigo
        
            'Obtém dados do Registro
            objItemNF.iItem = iItem
            
            sCFOP = Replace(sCFOPXML, "Item", "")
            Select Case sCFOP
                Case "5902"
                    sCFOPInterno = "1902"
                Case "5906", "5907"
                    sCFOPInterno = "1907"
                Case Else
                    sCFOPInterno = "1949"
            End Select
            
            lErro = CF("Importa_Xml_NFe_Loc_Prod", giFilialEmpresa, lForn, iFilForn, sProdXML, sEANXML, objProduto)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
            objItemNF.sProduto = objProduto.sCodigo
            objItemNF.sDescricaoItem = sProdDesc
            objItemNF.dQuantidade = dQtd
            objItemNF.dPrecoUnitario = dPreco
            
            objItemNF.sUnidadeMed = objProduto.sSiglaUMCompra
            objItemNF.sUMEstoque = objProduto.sSiglaUMEstoque
            objItemNF.sUMVenda = objProduto.sSiglaUMVenda
            objItemNF.iClasseUM = objProduto.iClasseUM
            
            objNF.ColItensNF.Add1 objItemNF
            
            objItemNF.objTributacaoItemNF.iTipoTributacao = 33
            objItemNF.objTributacaoItemNF.iNaturezaOpManual = VAR_PREENCH_MANUAL
            objItemNF.objTributacaoItemNF.sNaturezaOp = sCFOPInterno
            objItemNF.objTributacaoItemNF.iNaturezaOpManual = VAR_PREENCH_MANUAL
            
            objItemNF.objTributacaoItemNF.sProduto = objProduto.sCodigo
            objItemNF.objTributacaoItemNF.iExTIPI = objProduto.iExTIPI
            objItemNF.objTributacaoItemNF.sGenero = objProduto.sGenero
            objItemNF.objTributacaoItemNF.iProdutoEspecifico = objProduto.iProdutoEspecifico
            objItemNF.objTributacaoItemNF.sUMTrib = objProduto.sSiglaUMTrib
            objItemNF.objTributacaoItemNF.sIPICodProduto = objProduto.sIPICodigo
            objItemNF.objTributacaoItemNF.dQtdTrib = objItemNF.dQuantidade
            objItemNF.objTributacaoItemNF.dValorUnitTrib = objItemNF.dPrecoUnitario
            objItemNF.objTributacaoItemNF.dDescontoGrid = objItemNF.dValorDesconto
            objItemNF.objTributacaoItemNF.dPrecoTotal = objItemNF.dValorTotal
            objItemNF.objTributacaoItemNF.sProdutoDescricao = objProduto.sDescricao
            objItemNF.objTributacaoItemNF.sISSQN = objProduto.sISSQN
            
            'Converte a unidade de medida de UM de Venda para a UM da linha do Grid
            lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objItemNF.sUnidadeMed, objProduto.sSiglaUMEstoque, dFator)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            'Cria distribuição do Item
            Set objItemNFAlocacao = objItemNF.ColAlocacoes.Add(objAlmoxarifado.iCodigo, objAlmoxarifado.sNomeReduzido, Arredonda_Estoque(objItemNF.dQuantidade * dFator))
            objItemNFAlocacao.sUnidadeMed = objProduto.sSiglaUMEstoque
            
            lErro = Import_Xml_NFe_Preenche_Lote(objNF, objItemNF, sLoteXML)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
            lErro = Comando_BuscarProximo(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209877
        
        Loop
        
        objNF.sNaturezaOp = sCFOP
               
        objNF.objTributacaoNF.iTipoTributacao = 33
        objNF.objTributacaoNF.iNaturezaOpManual = VAR_PREENCH_MANUAL
        objNF.objTributacaoNF.sNaturezaOpInterna = sCFOPInterno
        objNF.objTributacaoNF.iTipoTributacaoManual = VAR_PREENCH_MANUAL
        
        objNF.lNumPedidoVenda = 0
        
        'GRAVA A NF
        '****************************************************************************
        'Verifica se a existe nota fiscal está cadastrada
        lErro = CF("NFiscal_Le_1", objNF)
        If lErro <> SUCESSO And lErro <> 83971 Then gError ERRO_SEM_MENSAGEM
        
        'Se ela não está cadastrada
        If lErro <> SUCESSO Then

            Set objTipoDocInfo = New ClassTipoDocInfo
            
            objTipoDocInfo.iCodigo = objNF.iTipoNFiscal
            
            'Lê o Tipo de Documento
            lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
            If lErro <> SUCESSO And lErro <> 31415 Then gError ERRO_SEM_MENSAGEM
            
            'pesquisa a existencia da nota fiscal com os dados passados como parametro e não exibe mensagens
            lErro = CF("NFiscalEntrada_Verifica_Existencia2", objNF, objTipoDocInfo, True)
            If lErro <> SUCESSO And lErro <> 61414 And lErro <> 89723 Then gError ERRO_SEM_MENSAGEM
                          
            'Se for uma nota nova
            If lErro = SUCESSO Then
                
                Set objTribTab = New ClassTribTab
                
                lErro = objTribTab.TributacaoNF_Reset(objNF)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                'obtem dados do contrato
                Call objTribTab.TipoNFiscal_Definir(objTipoDocInfo.iCodigo, objTipoDocInfo.sSigla)
                Call objTribTab.DataEmissao_Alterada(objNF.dtDataEmissao)
                Call objTribTab.Serie_Alterada(objNF.sSerie)
                
                lErro = gobjTributacao.AtualizaImpostos(objTribTab.mvarobjNFTributacao, 0)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                lErro = objTribTab.PreencherNF(objNF)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                Set objTribTab = Nothing
                
                objNF.sNaturezaOp = objNF.objTributacaoNF.sNaturezaOpInterna
            
                'Grava a Nota em Transação
                lErro = CF("NFiscalEntradaDevolucao_Grava_EmTrans", objNF, Nothing)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                If Len(Trim(sNFs)) <> 0 Then sNFs = sNFs & SEPARADOR
                sNFs = sNFs & CStr(objNF.lNumNotaFiscal)
                
            End If
        End If

        lErro = Comando_Executar(alComando(3), "UPDATE ImportNFeXml SET NumIntNF = ? WHERE ChvNFe = ?", objNF.lNumIntDoc, sChvNFe)
        If lErro <> AD_SQL_SUCESSO Then gError 209878
        
        For Each vsNomeArq In colArqsSep
        
            lErro = Comando_Executar(alComando(12), "UPDATE PharLMLogSep SET NumIntNF = ? WHERE NomeArq = ?", objNF.lNumIntDoc, vsNomeArq)
            If lErro <> AD_SQL_SUCESSO Then gError 211059
            
        Next
    
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209879
    
    Loop
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Fechar transacao
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 209880
    
    Call Rotina_Aviso(vbOKOnly, "AVISO_NF_IMPORT_COM_SUCESSO", sNFs)

    Importa_Xml_NFe_Cust = SUCESSO

    Exit Function

Erro_Importa_Xml_NFe_Cust:

    Importa_Xml_NFe_Cust = gErr

    Select Case gErr

        Case 209866
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 209867
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 209868, 209869
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISFORNECEDORES", gErr)

        Case 209870
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORN_CGC_NAO_CADASTRADO", gErr, CGC_EMITENTE)
            
        Case 209871, 209872
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISEMPRESA", gErr)

        Case 209873, 209874, 209879
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_IMPORTNFEXML", gErr)

        Case 209875, 209876, 209877
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_IMPORTNFEXML", gErr)

        Case 209878
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_IMPORTNFEXML", gErr)

        Case 209880
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case 211056 To 211059
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PHARLMLOGSEP", gErr)
            
        Case 211060
            Call Rotina_Erro(vbOKOnly, "ERRO_PHARLMLOGSEP_NAO_LOCALIZADO", gErr, objNF.lNumNotaFiscal)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209881)

    End Select
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Call Transacao_Rollback

    Exit Function
    
End Function

'Private Function Import_Xml_NFe_Preenche_Lote(ByVal objNF As ClassNFiscal, ByVal objItemNF As ClassItemNF) As Long
Private Function Import_Xml_NFe_Preenche_Lote(ByVal objNF As ClassNFiscal, ByVal objItemNF As ClassItemNF, ByVal sLoteXML As String) As Long

Dim lErro As Long, iIndice As Integer
Dim alComando(0 To 1) As Long
Dim lNumPedido As Long, sProdMask As String
Dim sLote As String, dQtde As Double
Dim objRastroItemNF As ClassRastroItemNF
Dim objRastroLote As New ClassRastreamentoLote

On Error GoTo Erro_Import_Xml_NFe_Preenche_Lote

'    'Abertura de Comando
'    For iIndice = LBound(alComando) To UBound(alComando)
'        alComando(iIndice) = Comando_Abrir()
'        If alComando(iIndice) = 0 Then gError 211000
'    Next
'
    Call Mascara_RetornaProdutoTela(objItemNF.sProduto, sProdMask)
'
'    sLote = String(STRING_MAXIMO, 0)
'
'    lErro = Comando_Executar(alComando(1), "SELECT Lote, SUM(Quantidade) FROM PharLMLogSep AS P, PharLMLogSepItens AS I WHERE P.NumPedido = I.NumPedido AND P.Situacao = 'SE' AND P.NumIntNF = 0 AND P.DataPedido <= ? AND I.Produto = ? AND I.Lote <> '' AND P.NomeArq IN (" & sNomeArqs & ") GROUP BY Lote ", _
'        sLote, dQtde, objNF.dtDataEmissao, sProdMask)
'    If lErro <> AD_SQL_SUCESSO Then gError 211004
'
'    lErro = Comando_BuscarPrimeiro(alComando(1))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211005
'
'    Do While lErro = AD_SQL_SUCESSO

        sLote = sLoteXML
        dQtde = objItemNF.dQuantidade
        
    If Len(Trim(sLote)) > 0 Then
     
        'Obtém os dados fixos
        Set objRastroItemNF = New ClassRastroItemNF
     
        objRastroItemNF.iAlmoxCodigo = objItemNF.iAlmoxarifado
        objRastroItemNF.dAlmoxQtdAlocada = objItemNF.dQuantidade
        objRastroItemNF.iItem = objItemNF.iItem
        objRastroItemNF.sProduto = objItemNF.sProduto
        objRastroItemNF.sSiglaUM = objItemNF.sUnidadeMed
      
        'Obtém os dados do registro
        objRastroItemNF.sLote = Trim(sLote)
        objRastroItemNF.dLoteQdtAlocada = dQtde
        objRastroItemNF.iLoteFilialOP = 0
     
        'Le o Lote
        Set objRastroLote = New ClassRastreamentoLote
         
        objRastroLote.sCodigo = objRastroItemNF.sLote
        objRastroLote.sProduto = objRastroItemNF.sProduto
        objRastroLote.iFilialOP = objRastroItemNF.iLoteFilialOP
     
        lErro = CF("RastreamentoLote_Le", objRastroLote)
        If lErro <> SUCESSO And lErro <> 75710 Then gError ERRO_SEM_MENSAGEM
         
        If lErro <> SUCESSO Then
            gError 211006
        End If
         
        objItemNF.colRastreamento.Add objRastroItemNF
        
    End If
     
'        lErro = Comando_BuscarProximo(alComando(1))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211007
'
'    Loop
'
'    'Fecha Comando
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
    
    Import_Xml_NFe_Preenche_Lote = SUCESSO

    Exit Function

Erro_Import_Xml_NFe_Preenche_Lote:

    Import_Xml_NFe_Preenche_Lote = gErr

    Select Case gErr
    
        Case 211000
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 211004, 211005, 211007
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PHARLMLOGSEPITENS", gErr)
            
        Case 211006
            Call Rotina_Erro(vbOKOnly, "ERRO_LOTE_NAO_CADASTRADO_RASTREAMENTO", gErr, sLote, sProdMask, 0)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211008)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function
    
End Function

Private Function Importa_Arq_Sep_LMLog() As Long

Dim lErro As Long, iIndice As Integer, sDir As String
Dim objFSO As New FileSystemObject, objFolder As Folder
Dim objFile As File, objTS As TextStream, sRegistro As String
Dim lTransacao As Long, bImportou As Boolean, bArqAberto As Boolean
Dim alComando(0 To 4) As Long, dtDataImport As Date, dtData As Date, sData As String
Dim iSeq As Integer, colNFs As New Collection, objNF As ClassNFiscal

On Error GoTo Erro_Importa_Arq_Sep_LMLog

    bArqAberto = False
    
    lErro = CF("Config_Le", "CRFATConfig", CRFATCFG_DIR_IMPORT_XML_NFE, EMPRESA_TODA, sDir)
    If lErro <> SUCESSO And lErro <> 208279 Then gError ERRO_SEM_MENSAGEM
        
    'Se a Pasta Proc não existir tem que criar
    If Len(Dir(sDir & "proc\", vbDirectory)) = 0 Then MkDir sDir & "Proc\"
    
    'Pega todos os aquivos da pasta
    Set objFolder = objFSO.GetFolder(sDir)
    
    'Para cada arquivo
    For Each objFile In objFolder.Files

        bArqAberto = False
        bImportou = False

        'Abrir transacao
        lTransacao = Transacao_Abrir()
        If lTransacao = 0 Then gError 211009
    
        'Abertura de Comando
        For iIndice = LBound(alComando) To UBound(alComando)
            alComando(iIndice) = Comando_Abrir()
            If alComando(iIndice) = 0 Then gError 211010
        Next
            
        If InStr(1, objFile.Name, "SEPAITENS_0003_") <> 0 Then
        
            lErro = Comando_Executar(alComando(2), "SELECT DataImport FROM PharLMLogSepItens WHERE NomeArq = ? ", _
                dtDataImport, objFile.Name)
            If lErro <> AD_SQL_SUCESSO Then gError 211011
            
            lErro = Comando_BuscarPrimeiro(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211012
            
            If lErro = AD_SQL_SUCESSO Then gError 211013 'O arquivo já foi importado anteriormente
        
            Set objTS = objFile.OpenAsTextStream()
            bArqAberto = True
            iSeq = 0
            Do While Not objTS.AtEndOfLine
            
                sRegistro = objTS.ReadLine
                iSeq = iSeq + 1
                
                sData = LMLog_Campo(sRegistro, 6)
                If Len(Trim(sData)) > 0 Then sData = left(sData, 2) & "/" & Mid(sData, 3, 2) & "/" & right(sData, 4)
                dtData = StrParaDate(sData)
                
                lErro = Comando_Executar(alComando(0), "INSERT INTO PharLMLogSepItens(NomeArq,DataImport,HoraImport,NumPedido,Seq, Produto,CodigoBarras,Quantidade,Lote,ValidadeLote) VALUES (?,?,?,?,?,?,?,?,?,?)", _
                objFile.Name, Date, CDbl(Time), LMLog_Campo(sRegistro, 1), iSeq, LMLog_Campo(sRegistro, 2), LMLog_Campo(sRegistro, 3), StrParaDbl(LMLog_Campo(sRegistro, 4)), LMLog_Campo(sRegistro, 5), dtData)
                If lErro <> AD_SQL_SUCESSO Then gError 211014

            Loop
        
            Call objTS.Close
            bArqAberto = False
            bImportou = True
        
        ElseIf InStr(1, objFile.Name, "SEPAPED_0003_") <> 0 Then
        
            lErro = Comando_Executar(alComando(3), "SELECT DataImport FROM PharLMLogSep WHERE NomeArq = ? ", _
                dtDataImport, objFile.Name)
            If lErro <> AD_SQL_SUCESSO Then gError 211015
            
            lErro = Comando_BuscarPrimeiro(alComando(3))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211016
            
            If lErro = AD_SQL_SUCESSO Then gError 211017 'O arquivo já foi importado anteriormente

            Set objTS = objFile.OpenAsTextStream()
            bArqAberto = True
            Do While Not objTS.AtEndOfLine
            
                sRegistro = objTS.ReadLine
                
                sData = LMLog_Campo(sRegistro, 2)
                If Len(Trim(sData)) > 0 Then sData = left(sData, 2) & "/" & Mid(sData, 3, 2) & "/" & right(sData, 4)
                dtData = StrParaDate(sData)
                
                lErro = Comando_Executar(alComando(1), "INSERT INTO PharLMLogSep(NomeArq,DataImport,HoraImport,NumPedido,DataPedido,NumPedidoLMLog,Situacao,MotivoDivergencia,PesoTotal,QuantidadeVolumes,CNPJTransportadora) VALUES (?,?,?,?,?,?,?,?,?,?,?)", _
                objFile.Name, Date, CDbl(Time), LMLog_Campo(sRegistro, 1), dtData, LMLog_Campo(sRegistro, 3), LMLog_Campo(sRegistro, 4), LMLog_Campo(sRegistro, 5), StrParaDbl(LMLog_Campo(sRegistro, 6)), StrParaDbl(LMLog_Campo(sRegistro, 7)), LMLog_Campo(sRegistro, 8))
                If lErro <> AD_SQL_SUCESSO Then gError 211018
                
                Set objNF = New ClassNFiscal
                
                objNF.lNumNotaFiscal = LMLog_Campo(sRegistro, 1)
                objNF.dtDataEmissao = dtData
                objNF.lVolumeQuant = StrParaLong(LMLog_Campo(sRegistro, 7))
                                
                colNFs.Add objNF

                lErro = Comando_Executar(alComando(4), "UPDATE NFiscal SET VolumeQuant = ? WHERE NumNotaFiscal = ? AND DataEmissao = ? AND FornEntTerc = ?", _
                objNF.lVolumeQuant, objNF.lNumNotaFiscal, objNF.dtDataEmissao, 403)
                If lErro <> AD_SQL_SUCESSO Then gError 211018

            Loop
        
            Call objTS.Close
            bArqAberto = False
            bImportou = True
        
        End If

        'Fecha Comando
        For iIndice = LBound(alComando) To UBound(alComando)
            Call Comando_Fechar(alComando(iIndice))
        Next
    
        'Fechar transacao
        lErro = Transacao_Commit()
        If lErro <> AD_SQL_SUCESSO Then gError 211019
                
        If bImportou Then
            FileCopy sDir & objFile.Name, sDir & "Proc\" & objFile.Name
            Kill sDir & objFile.Name
        End If

    Next
    
    lErro = PharLMLogSep_Valida(colNFs)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Importa_Arq_Sep_LMLog = SUCESSO

    Exit Function

Erro_Importa_Arq_Sep_LMLog:

    Importa_Arq_Sep_LMLog = gErr

    Select Case gErr
    
        Case 211009
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
    
        Case 211010
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 211011, 211012
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PHARLMLOGSEPITENS", gErr)

        Case 211013
            Call Rotina_Erro(vbOKOnly, "ERRO_PHARLMLOGSEP_JA_IMPORTADO", gErr, objFile.Name, dtDataImport)

        Case 211014
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERT_PHARLMLOGSEPITENS", gErr)

        Case 211015, 211016
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PHARLMLOGSEP", gErr)

        Case 211017
            Call Rotina_Erro(vbOKOnly, "ERRO_PHARLMLOGSEP_JA_IMPORTADO", gErr, objFile.Name, dtDataImport)

        Case 211018
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERT_PHARLMLOGSEP", gErr)

        Case 211019
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211020)

    End Select
    
    If bArqAberto Then Call objTS.Close
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Call Transacao_Rollback

    Exit Function
    
End Function

Private Function LMLog_Campo(ByVal sRegistro As String, ByVal iCampo As Integer) As String

Dim iPos1 As Integer
Dim iPos2 As Integer
Dim iIndice As Integer
Dim sAux As String

    iPos2 = 0
    iPos1 = InStr(iPos2 + 1, sRegistro, "|")
    For iIndice = 1 To iCampo - 1
        iPos2 = iPos1
        iPos1 = InStr(iPos2 + 1, sRegistro, "|")
        If iPos1 = 0 Then iPos1 = Len(sRegistro) + 1
    Next
    sAux = Mid(sRegistro, iPos2 + 1, iPos1 - iPos2 - 1)
    'If iCampo = 1 Then sAux = Mid(sAux, 2)
    LMLog_Campo = sAux

End Function

Private Function PharLMLogSep_Valida(ByVal colNFs As Collection) As Long

Dim lErro As Long
Dim lNumIntNF As Long, lComando As Long
Dim objNFiscal As ClassNFiscal
Dim lTransacao As Long

Const OL_CODFORN_LMLOG = 403

On Error GoTo Erro_PharLMLogSep_Valida

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 211194

    For Each objNFiscal In colNFs
  
        lErro = Comando_Executar(lComando, "SELECT NumIntDoc FROM NFiscal WHERE Status <> 7 AND DataEmissao BETWEEN ? AND ? AND FilialEmpresa = ? AND FornEntTerc = ? AND NumNotaFiscal = ?", _
                                        lNumIntNF, objNFiscal.dtDataEmissao - PERIODO_EMISSAO, objNFiscal.dtDataEmissao + PERIODO_EMISSAO, giFilialEmpresa, OL_CODFORN_LMLOG, objNFiscal.lNumNotaFiscal)
        If lErro <> AD_SQL_SUCESSO Then gError 211195
            
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211196
    
        If lErro = AD_SQL_SUCESSO Then
        
            lErro = CF("NF_Valida_PharSep", lNumIntNF, True)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            objNFiscal.lNumIntDoc = lNumIntNF
                      
        End If
                
    Next
    
    'Fecha os comandos
    Call Comando_Fechar(lComando)

    'Abrir transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 211009

    For Each objNFiscal In colNFs
    
        If objNFiscal.lNumIntDoc <> 0 Then
                      
            lErro = CF("NFiscal_Trata_MsgItem", objNFiscal)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
        End If
        
    Next

    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 211019
            
    PharLMLogSep_Valida = SUCESSO
    
    Exit Function
    
Erro_PharLMLogSep_Valida:

    PharLMLogSep_Valida = gErr
    
    Select Case gErr

        Case 211194
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 211195 To 211196
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", gErr)

        Case 211009
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
    
        Case 211019
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211197)

    End Select
    
    'Fecha os comandos
    Call Comando_Fechar(lComando)
    
    Call Transacao_Rollback
       
    Exit Function
    
End Function

Public Function NF_Valida_PharSep(ByVal lNumIntNF As Long, Optional ByVal bSoAvisa As Boolean = False) As Long

Dim lErro As Long, lComando As Long
Dim objNF As New ClassNFiscal, objItemNF As ClassItemNF
Dim objNFSep As New ClassNFiscal, objItemNFSep As ClassItemNF
Dim sProdMask As String, dQtde As Double, sProdFormat As String
Dim iProdutoPreenchido As Integer, bAchou As Boolean

Const OL_CODFORN_LMLOG = 403

On Error GoTo Erro_NF_Valida_PharSep

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 211187

    objNF.lNumIntDoc = lNumIntNF
    
    lErro = CF("NFiscal_Le", objNF)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Se for uma NF que deveria ter sido entregue pela LM
    If objNF.lFornEntTerc = OL_CODFORN_LMLOG Then
    
        'Lê os itens da Nota Fiscal
        lErro = CF("NFiscalItens_Le", objNF)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
   
        sProdMask = String(STRING_MAXIMO, 0)
    
        lErro = Comando_Executar(lComando, "SELECT Produto, SUM(Quantidade) AS Qtde FROM PharLMLogSepItens WHERE NumPedido = ? AND DataImport BETWEEN ? AND ? GROUP BY Produto ORDER BY Produto ", sProdMask, dQtde, objNF.lNumNotaFiscal, objNF.dtDataEmissao - PERIODO_EMISSAO, objNF.dtDataEmissao + PERIODO_EMISSAO)
        If lErro <> AD_SQL_SUCESSO Then gError 211188
        
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211189
    
        Do While lErro <> AD_SQL_SEM_DADOS
        
            Set objItemNFSep = New ClassItemNF
            
'            lErro = CF("Produto_Formata", sProdMask, sProdFormat, iProdutoPreenchido)
'            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            sProdFormat = Replace(sProdMask, ".", "")
            sProdFormat = sProdFormat & String(10 - Len(sProdFormat), " ")
            
            objItemNFSep.sProduto = sProdFormat
            objItemNFSep.dQuantidade = dQtde
            
            objNFSep.ColItensNF.Add1 objItemNFSep
        
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211190
        
        Loop
        
        For Each objItemNF In objNF.ColItensNF
            bAchou = False
            For Each objItemNFSep In objNFSep.ColItensNF
                If objItemNF.sProduto = objItemNFSep.sProduto Then
                    bAchou = True
                    Exit For
                End If
            Next
            If Not bAchou Then
                If bSoAvisa Then
                    Call Rotina_Aviso(vbOKOnly, "ERRO_ITEM_NF_SEM_SEPARACAO", objItemNF.sProduto, objNF.lNumNotaFiscal, objNF.dtDataEmissao)
                Else
                    gError 211198
                End If
            Else
                If Abs(objItemNF.dQuantidade - objItemNFSep.dQuantidade) > QTDE_ESTOQUE_DELTA Then
                    If bSoAvisa Then
                        Call Rotina_Aviso(vbOKOnly, "ERRO_QTDE_SEP_DIF_QTDE_NF", objItemNF.sProduto, objNF.lNumNotaFiscal, objNF.dtDataEmissao, Formata_Estoque(objItemNF.dQuantidade), Formata_Estoque(objItemNFSep.dQuantidade))
                    Else
                        gError 211191
                    End If
                End If
            End If
        
        Next
    
    End If
    
    'Fecha os comandos
    Call Comando_Fechar(lComando)

    NF_Valida_PharSep = SUCESSO
    
    Exit Function
    
Erro_NF_Valida_PharSep:

    NF_Valida_PharSep = gErr
    
    Select Case gErr

        Case 211187
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 211188 To 211190
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PHARLMLOGSEPITENS", gErr)
            
        Case 211191
            Call Rotina_Erro(vbOKOnly, "ERRO_QTDE_SEP_DIF_QTDE_NF", gErr, objItemNF.sProduto, objNF.lNumNotaFiscal, objNF.dtDataEmissao, Formata_Estoque(objItemNF.dQuantidade), Formata_Estoque(objItemNFSep.dQuantidade))

        Case 211198
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEM_NF_SEM_SEPARACAO", gErr, objItemNF.sProduto, objNF.lNumNotaFiscal, objNF.dtDataEmissao)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211192)

    End Select
    
    'Fecha os comandos
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function



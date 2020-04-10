VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFATGrava"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'1 - 08/2000 Márcio - Função de gravação alterada para incluir preço base no ItemPV
'2 - 09/2000 Márcio - Função de gravação alterada para incluir Cliente, Vendedor e Preço Unitário

Option Explicit

Private Declare Function Conexao_AbrirExt Lib "ADSQLMN.DLL" Alias "AD_Conexao_Abrir" (ByVal driver_sql As Integer, ByVal lpParamIn As String, ByVal ParamLenIn As Integer, ByVal lpParamOut As String, lpParamLenOut As Integer) As Long
Private Declare Function Conexao_FecharExt Lib "ADSQLMN.DLL" Alias "AD_Conexao_Fechar" (ByVal lConexao As Long) As Long

Private Declare Function Comando_BindVarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_BindVar" (ByVal lComando As Long, lpVar As Variant) As Long
Private Declare Function Comando_PrepararInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Preparar" (ByVal lComando As Long, ByVal lpSQLStmt As String) As Long
Private Declare Function Comando_ExecutarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Executar" (ByVal lComando As Long) As Long

Private Sub INPAL_ConverteNatop43(ByVal sNatOp As String, sNatop3 As String)
'retorna em sNatop3 o codigo correspondente à natureza de operacao sNatop com 3 digitos
'sNatop pode ter 3 ou 4 digitos

    If Len(sNatOp) = 3 Then
        sNatop3 = sNatOp
    Else
    
        Select Case sNatOp
        
            Case "5101", "5109", "5111", "5116", "5118", "5122"
                sNatop3 = "511"
                
            Case "5102", "5110", "5112", "5117", "5119", "5123"
                sNatop3 = "512"
            
            Case "5124", "5125"
                sNatop3 = "513"
        
            Case "6101", "6107", "6109", "6111", "6116", "6118", "6122"
                sNatop3 = "611"
            
            Case "6102", "6108", "6110", "6112", "6117", "6119", "6123"
                sNatop3 = "612"
        
            Case "6124", "6125"
                sNatop3 = "613"
            
            Case Else
                sNatop3 = left(sNatOp, 3)
                
        End Select
        
    End If

End Sub

Public Function Gera_ArqComissoes(ByVal dtDataInicio As Date, ByVal dtDataFim As Date) As Long
'Lê os itens de notas fiscais de saída que geram faturamento dentro do período entre dtDataInicio e dtDataFim
'Monta o registro em uma única linha
'Insere o registro no arquivo

Dim lErro As Long
Dim lComando As Long
Dim sNomeArq As String
Dim bGerou As Boolean
Dim iSequencial As Integer
Dim tArqComissoes As typeArqComissoes
Dim sRegistro As String
Dim dFator As Double
Dim sStatusNota As String
Dim objProduto As ClassProduto, sNatOp As String, sNatOpInt As String

On Error GoTo Erro_Gera_ArqComissoes

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 94839
    
    'Enquanto não tiver gerado o nome do arquivo
    Do While Not bGerou
    
        'Limpa a variável que receberá o nome do arquivo
        sNomeArq = ""
        
        'Gera o sequencial que irá integrar o nome do arquivo
        iSequencial = iSequencial + 1
        
        'Monta o nome do arquivo
        sNomeArq = App.Path & "\" & sNomeArq & "ArqComissoes_" & Format(gdtDataHoje, "DDMMYY") & Formata_Campo(ALINHAMENTO_ESQUERDA, 2, "0", iSequencial) & ".txt"
        
        'Se o len retornar 0, significa que não existe arquivo com sNomeArq,
        'portanto, esse nome pode ser utilizado para geração do arquivo
        If Len(Trim(Dir(sNomeArq))) = 0 Then bGerou = True
    
    Loop
    
    'Abre o arquivo
    Open sNomeArq For Output As #1
    
    With tArqComissoes
        
        .sSerie = String(STRING_SERIE, 0)
        .sProduto = String(STRING_PRODUTO, 0)
        .sNomeProduto = String(STRING_PRODUTO_DESCRICAO, 0)
        .sLinha = String(STRING_CATEGORIAPRODUTOITEM_ITEM, 0)
        .sGrupo = String(STRING_CATEGORIAPRODUTOITEM_ITEM, 0)
        .sSubGrupo = String(STRING_CATEGORIAPRODUTOITEM_ITEM, 0)
        .sNomeGrupo = String(STRING_CATEGORIAPRODUTOITEM_DESCRICAO, 0)
        .sClassFiscal = String(STRING_PRODUTO_IPI_CODIGO, 0)
        .sRazaoSocCliente = String(STRING_CLIENTE_RAZAO_SOCIAL, 0)
        .sNomeVendedor = String(STRING_VENDEDOR_NOME, 0)
        .sNatOpInt = String(STRING_NATUREZAOP_CODIGO, 0)
        .sNatOp = String(STRING_NATUREZAOP_CODIGO, 0)
        .sUM = String(STRING_UM_SIGLA, 0)

        'Lê no banco de dados os registro que serão inseridos no arquivo
        lErro = Comando_Executar(lComando, "SELECT FilialEmp, NumNotaFiscal, Serie, ItemNF, DataEmissao, CondPagto, Produto, NomeProduto, Linha, Grupo, SubGrupo, NomeGrupo, ClassFiscal, CodCliente, FilialCliente, RazaoSocCliente, Regiao, NomeVendedor, NatOpInt, NatOp, Quant, UM, PesoSemConv, PrecoUn, DespFinanceira, AliquotaIPI, StatusNota, Desconto, PrecoAVista FROM ArqComissoes_Inpal WHERE DataEmissao>=? AND DataEmissao<= ? ORDER BY NumNotaFiscal, ItemNF", .iFilialEmp, .lNumNotaFiscal, .sSerie, .iItemNF, .dtDataEmissao, .iCondPagto, .sProduto, .sNomeProduto, .sLinha, .sGrupo, .sSubGrupo, .sNomeGrupo, .sClassFiscal, .lCodCliente, .iFilialCliente, .sRazaoSocCliente, .iRegiao, .sNomeVendedor, .sNatOpInt, .sNatOp, .dQuant, .sUM, .dPesoSemConv, .dPrecoUn, .dDespFinanceira, .dAliquotaIPI, .iStatusNota, .dDesconto, .dPrecoAVista, dtDataInicio, dtDataFim)
        If lErro <> AD_SQL_SUCESSO Then gError 94840
        
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 94841
        
        'Se não encontrou => erro
        If lErro = AD_SQL_SEM_DADOS Then gError 94842
    
        'Enquanto retornar sucesso
        Do While lErro = SUCESSO
            
            'Instancia um novo objProduto
            Set objProduto = New ClassProduto
            
            'Guarda o código no obj
            objProduto.sCodigo = .sProduto
            
            'Lê os dados do produto (a leitura serve para obter a classeUM do produto)
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO Then gError 94854
            
            If Not UnidadeDeMedida_IgnorarNaVenda(.sUM) Then
            
                'O peso será a quantidade do produto convertida para a UM Kg
                'Obtém o fator de conversão que será aplicado para converter o peso em KG
                lErro = CF("UM_Conversao", objProduto.iClasseUM, .sUM, "Kg", dFator)
                If lErro <> SUCESSO Then dFator = 1 'Jones 01/06/03
                
                'Faz a conversão
                .dPesoSemConv = .dPesoSemConv * dFator
            
            End If
            
            'Guarda o status da nota como Normal (S-Normal; C-Cancelada)
            sStatusNota = "S"
            
            'Se a nota foi cancelada, altera o status
            If .iStatusNota = STATUS_CANCELADO Then sStatusNota = "C"
        
            'Indica que a nota é referente a uma venda (todas as notas representam vendas)
            .sVenda = "S"
            
            Call INPAL_ConverteNatop43(.sNatOpInt, sNatOpInt)
            Call INPAL_ConverteNatop43(.sNatOp, sNatOp)
            
            'Monta o registro que será inserido no arquivo
            sRegistro = Formata_Campo(ALINHAMENTO_ESQUERDA, 2, "0", .iFilialEmp) & Formata_Campo(ALINHAMENTO_ESQUERDA, 6, "0", .lNumNotaFiscal) & Formata_Campo(ALINHAMENTO_DIREITA, 3, " ", .sSerie) & Formata_Campo(ALINHAMENTO_ESQUERDA, 2, "0", .iItemNF) & Format(.dtDataEmissao, "DDMMYY") & Formata_Campo(ALINHAMENTO_ESQUERDA, 3, "0", .iCondPagto) & Formata_Campo(ALINHAMENTO_DIREITA, 4, "0", .sProduto) & Formata_Campo(ALINHAMENTO_DIREITA, 40, " ", .sNomeProduto) & Formata_Campo(ALINHAMENTO_ESQUERDA, 2, "0", .sLinha) & Formata_Campo(ALINHAMENTO_ESQUERDA, 2, "0", .sGrupo) & Formata_Campo(ALINHAMENTO_DIREITA, 2, "0", .sSubGrupo) & Formata_Campo(ALINHAMENTO_DIREITA, 30, " ", .sNomeGrupo) & Formata_Campo(ALINHAMENTO_ESQUERDA, 10, "0", .sClassFiscal) & Formata_Campo(ALINHAMENTO_ESQUERDA, 5, "0", .lCodCliente) & Formata_Campo(ALINHAMENTO_ESQUERDA, 2, "0", .iFilialCliente) & Formata_Campo(ALINHAMENTO_DIREITA, 40, " ", .sRazaoSocCliente) _
                & Formata_Campo(ALINHAMENTO_ESQUERDA, 4, "0", .iRegiao) & Formata_Campo(ALINHAMENTO_DIREITA, 30, " ", .sNomeVendedor) & Formata_Campo(ALINHAMENTO_ESQUERDA, 2, "0", sNatOpInt) & Formata_Campo(ALINHAMENTO_ESQUERDA, 3, "0", sNatOp) & .sVenda & Replace(Formata_Campo(ALINHAMENTO_ESQUERDA, 7, " ", Format(.dQuant, "####0.0")), ",", ".") & Formata_Campo(ALINHAMENTO_ESQUERDA, 2, " ", .sUM) & Replace(Formata_Campo(ALINHAMENTO_ESQUERDA, 7, " ", Format(.dPesoSemConv, "####0.0")), ",", ".") & Replace(Formata_Campo(ALINHAMENTO_ESQUERDA, 12, " ", Format(.dPrecoUn, "########0.00")), ",", ".") & Replace(Formata_Campo(ALINHAMENTO_ESQUERDA, 12, " ", Format(.dDespFinanceira, "########0.00")), ",", ".") & Formata_Campo(ALINHAMENTO_ESQUERDA, 2, "0", .dAliquotaIPI * 100) & sStatusNota & Replace(Formata_Campo(ALINHAMENTO_ESQUERDA, 12, " ", Format(.dDesconto, "########0.00")), ",", ".") & Replace(Formata_Campo(ALINHAMENTO_ESQUERDA, 12, " ", Format(.dPrecoAVista, "########0.00")), ",", ".")
            
            
''''******Shirley 27/06/2002*********
''''            & Formata_Campo(ALINHAMENTO_ESQUERDA, 4, "0", .iRegiao) & Formata_Campo(ALINHAMENTO_DIREITA, 30, " ", .sNomeVendedor) & Formata_Campo(ALINHAMENTO_ESQUERDA, 2, "0", .sNatOpInt) & Formata_Campo(ALINHAMENTO_ESQUERDA, 3, "0", .sNatOp) & .sVenda & Formata_Campo(ALINHAMENTO_ESQUERDA, 7, " ", Format(.dQuant, "####0.0")) & Formata_Campo(ALINHAMENTO_ESQUERDA, 2, " ", .sUM) & Formata_Campo(ALINHAMENTO_ESQUERDA, 7, " ", Format(.dPesoSemConv, "####0.0")) & Formata_Campo(ALINHAMENTO_ESQUERDA, 12, " ", Format(.dPrecoUn, "########0.00")) & Formata_Campo(ALINHAMENTO_ESQUERDA, 12, " ", Format(.dDespFinanceira, "########0.00")) & Formata_Campo(ALINHAMENTO_ESQUERDA, 2, "0", .dAliquotaIPI * 100) & sStatusNota & Formata_Campo(ALINHAMENTO_ESQUERDA, 12, " ", Format(.dDesconto, "########0.00")) & Formata_Campo(ALINHAMENTO_ESQUERDA, 12, " ", Format(.dPrecoAVista, "########0.00"))
''''*********************************
            
            'Insere o registro
            Print #1, sRegistro
            
            'Busca os dados do próximo registro
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 94843
        
        Loop
    
    End With
        
    'Fecha o arquivo
    Close #1
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    'Exibe um aviso de sucesso com o nome do arquivo informado
    Call Rotina_Aviso(vbOKOnly, "AVISO_ARQCOMISSOES_INPAL_SUCESSO", sNomeArq)
    
    Gera_ArqComissoes = SUCESSO
    
    Exit Function
    
Erro_Gera_ArqComissoes:

    Gera_ArqComissoes = gErr
    
    Select Case gErr
    
        Case 94853
        
        Case 94839
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 94840, 94841, 94843
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ARQCOMISSOES_INPAL", gErr)
        
        Case 94842
            Call Rotina_Erro(vbOKOnly, "ERRO_ARQCOMISSOES_INPAL_INEXISTENTE", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
            
    End Select
    
    'Fecha o arquivo
    Close #1
    
    'Exclui o arquivo
    Kill sNomeArq
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

'2 - 09/2000 Márcio - Função de gravação alterada para incluir Cliente, Vendedor e Preço Unitário
'Incluido Cliente, Vendedor e Produto na chave primária
Function PrevVenda_Exclui(ByVal objPrevVenda As ClassPrevVenda) As Long
'Exclui a previsão de venda do BD

Dim lErro As Long
Dim lComando1 As Long
Dim lComando2 As Long
Dim lTransacao As Long
Dim sCodigo As String
Dim iFilialEmpresa As Integer

On Error GoTo Erro_PrevVenda_Exclui

    'Abre os comandos
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 34535

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 34536

    'Inicia a transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 34537

    'Inicializa string com o tamanho do BD
    sCodigo = String(STRING_PREVVENDA_CODIGO, 0)

    'Pequisa a previsão em questão no BD
    lErro = Comando_ExecutarPos(lComando1, "SELECT FilialEmpresa, Codigo FROM PrevVenda WHERE FilialEmpresa = ? AND Codigo = ? AND Cliente = ? AND Vendedor = ? AND Produto = ?", 0, iFilialEmpresa, sCodigo, objPrevVenda.iFilialEmpresa, objPrevVenda.sCodigo, objPrevVenda.lCliente, objPrevVenda.iVendedor, objPrevVenda.sProduto)
    If lErro <> AD_SQL_SUCESSO Then Error 34538

    'Lê a previsão se estiver no BD
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 34539

    'Não encontrou a previsão no BD
    If lErro = AD_SQL_SEM_DADOS Then Error 34540

    'Lock da previsão
    lErro = Comando_LockExclusive(lComando1)
    If lErro <> AD_SQL_SUCESSO Then Error 34541

    'Exclui a previsão do BD - Falta alterar filialempresa
    lErro = Comando_ExecutarPos(lComando2, "DELETE FROM PrevVenda", lComando1)
    If lErro <> AD_SQL_SUCESSO Then Error 34542

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 34543

    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando1)

    PrevVenda_Exclui = SUCESSO

    Exit Function

Erro_PrevVenda_Exclui:

    PrevVenda_Exclui = Err

    Select Case Err

        Case 34535, 34536
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 34537
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 34538, 34539
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PREVVENDA", Err, objPrevVenda.sCodigo)

        Case 34540
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PREVVENDA_NAO_CADASTRADA", Err, objPrevVenda.sCodigo)

        Case 34541
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PREVVENDA", Err, objPrevVenda.sCodigo)

        Case 34542
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PREVVENDA", Err, objPrevVenda.sCodigo)

        Case 34543
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$)
    End Select

    Call Transacao_Rollback
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

'2 - 09/2000 Márcio - Função de gravação alterada para incluir Cliente, Vendedor e Preço Unitário
Function PrevVenda_Grava(ByVal objPrevVenda As ClassPrevVenda) As Long
'Grava a Previsão de Venda no BD

Dim alComando(1 To 6) As Long
Dim lTransacao As Long
Dim lErro As Long
Dim sCodigo As String
Dim tProduto As typeProduto
Dim iAlmoxarifado As Integer
Dim iIndice As Integer
Dim objProdutoFilial As New ClassProdutoFilial
Dim iFilialEmpresa As Integer

On Error GoTo Erro_PrevVenda_Grava

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 58256
    Next
    
    'Inicia a transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 58257

    If objPrevVenda.iCodRegiao > 0 Then
        
        'Lock na Regiao de Venda
        lErro = CF("RegiaoVenda_Lock", alComando(4), objPrevVenda.iCodRegiao)
        If lErro <> SUCESSO Then gError 58259
    
    End If
    
    objProdutoFilial.sProduto = objPrevVenda.sProduto
    objProdutoFilial.iFilialEmpresa = objPrevVenda.iFilialEmpresa
    
    'Faz lock do registro associado na tabela ProdutoFilial
    lErro = CF("ProdutoFilial_Le_Lock", alComando(5), objProdutoFilial)
    If lErro <> SUCESSO And lErro <> 41770 Then gError 25656
    If lErro = 41770 Then gError 25657
    
    'Lê no BD as caracteristicas do Produto e Locka (Gerencial, Faturamento)
    lErro = Comando_ExecutarLockado(alComando(1), "SELECT Gerencial, Faturamento FROM Produtos WHERE Codigo = ? ", tProduto.iGerencial, tProduto.iFaturamento, objPrevVenda.sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 58258
    
    'Faz Lock Shared do Produto
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO Then gError 58267
        
    'Faz Lock Shared do Produto
    lErro = Comando_LockShared(alComando(1))
    If lErro <> AD_SQL_SUCESSO Then gError 58259
    
    'Se for um Produto Gerencial ---> ERRO
    If tProduto.iGerencial = GERENCIAL Then gError 58260
    
    'Se for um Produto não Vendavel ----> ERRO
    If tProduto.iFaturamento = PRODUTO_NAO_VENDAVEL Then gError 58261
    
    'Verifica se já existe uma previsão de Vendas com o mesmo Código e Filial Empresa e data inicio e fim diferentes
    lErro = Comando_Executar(alComando(6), "SELECT FilialEmpresa FROM PrevVenda WHERE FilialEmpresa = ? AND Codigo = ? AND (DataInicio <> ? OR DataFim <> ?)", iFilialEmpresa, objPrevVenda.iFilialEmpresa, objPrevVenda.sCodigo, objPrevVenda.dtDataInicio, objPrevVenda.dtDataFim)
    If lErro <> AD_SQL_SUCESSO Then gError 500199
    
    lErro = Comando_BuscarPrimeiro(alComando(6))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 500200
    
    'Se encontrou a Previsão de vendas, erro
    If lErro = AD_SQL_SUCESSO Then gError 500201
    
    'Inicializa string com o tamanho do BD
    sCodigo = String(STRING_PREVVENDA_CODIGO, 0)

    'Pesquisa no BD a previsão em questão
    lErro = Comando_ExecutarPos(alComando(2), "SELECT Codigo FROM PrevVenda WHERE FilialEmpresa = ? AND Codigo = ? AND Cliente = ? AND Vendedor = ? AND Produto = ?", 0, sCodigo, giFilialEmpresa, objPrevVenda.sCodigo, objPrevVenda.lCliente, objPrevVenda.iVendedor, objPrevVenda.sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 58262

    'Lê a previsão se estiver no BD
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 58263
    
    'Se encontrou a previsão -> Alteração
    If lErro = AD_SQL_SUCESSO Then

        'Atualiza a Previsão
        lErro = Comando_ExecutarPos(alComando(3), "UPDATE PrevVenda SET DataPrevisao = ? , Produto = ?, CodRegiao = ?, Quantidade = ?, Valor = ?, PrecoUnitario = ?", alComando(2), objPrevVenda.dtDataPrevisao, objPrevVenda.sProduto, objPrevVenda.iCodRegiao, objPrevVenda.dQuantidade, objPrevVenda.dValor, objPrevVenda.dPrecoUnitario)
        If lErro <> AD_SQL_SUCESSO Then gError 58264

    Else 'Não encontrou a previsão -> Inserção

        'Insere a Previsão
        lErro = Comando_Executar(alComando(3), "INSERT INTO PrevVenda (FilialEmpresa, Codigo, Cliente, Vendedor, DataPrevisao, Produto, DataInicio, DataFim, CodRegiao, Quantidade, Valor, PrecoUnitario) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)", giFilialEmpresa, objPrevVenda.sCodigo, objPrevVenda.lCliente, objPrevVenda.iVendedor, objPrevVenda.dtDataPrevisao, objPrevVenda.sProduto, objPrevVenda.dtDataInicio, objPrevVenda.dtDataFim, objPrevVenda.iCodRegiao, objPrevVenda.dQuantidade, objPrevVenda.dValor, objPrevVenda.dPrecoUnitario)
        If lErro <> AD_SQL_SUCESSO Then gError 58265

    End If

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 58266

    'Fechamento comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PrevVenda_Grava = SUCESSO

    Exit Function

Erro_PrevVenda_Grava:

    PrevVenda_Grava = gErr

    Select Case gErr

        Case 25656, 58259 'Tratado na Rotina chamada
        
        Case 25657
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_FILIAL_NAO_CADASTRADO", gErr, objProdutoFilial.sProduto, objProdutoFilial.iFilialEmpresa)
     
        Case 58256
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 58257
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 58258, 58259, 58267, 500199, 500200
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS", gErr, objPrevVenda.sProduto)
        
        Case 58260
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", gErr, objPrevVenda.sProduto)
        
        Case 58261
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PODE_SER_VENDIDO2", gErr, objPrevVenda.sProduto)
        
        Case 58262, 58263
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PREVVENDA", gErr, objPrevVenda.sCodigo)

        Case 58264
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_PREVVENDA", gErr, objPrevVenda.sCodigo)

        Case 58265
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PREVVENDA", gErr, objPrevVenda.sCodigo)

        Case 58266
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case 500201
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PREVVENDAS_JA_CADASTRADA", gErr, objPrevVenda.sCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Call Transacao_Rollback
    
    'Fechamento comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'Funcoes Referentes a tela PlanComissoesInpal
'Colocadas aqui 20/03 por Tulio

Public Function ComissoesInpal_Grava(ByVal objComissoesInpalPlan As ClassComisInpalPlan, ByVal lPlanilhaIrma As Long) As Long
'Grava as informacoes da planilha no BD...
'objComissoesInpalPlan RECEBE(INPUT) os dados a serem gravados
'lPlanilhaIrma RECEBE(INPUT) a outra planilha, que sera passada para a funcao que faz a gravacao efetiva

Dim lErro As Long
Dim lTransacao As Long
Dim objComissoesInpalRegras As ClassComisInpalRegras
Dim objComissoesRegras As ClassComissoesRegras
Dim colComissoesRegras As New Collection

On Error GoTo Erro_ComissoesInpal_Grava

    'abre a trans
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 98863
    
    'grava os dados da planilha
    lErro = ComissoesInpalPlan_Grava_EmTrans(objComissoesInpalPlan, lPlanilhaIrma)
    If lErro <> SUCESSO Then gError 98867
    
    'Exclui do bd as regras em questao...
    lErro = ComissoesInpalRegras_Exclui_EmTrans(objComissoesInpalPlan.lCodigo)
    If lErro <> SUCESSO Then gError 98881
    
    'para cada regra da planilha
    For Each objComissoesInpalRegras In objComissoesInpalPlan.colComissoesInpalRegras
        
        'guarda o numintdoc da planilha em objComissoesInpalRegras
        objComissoesInpalRegras.lCodPlanilha = objComissoesInpalPlan.lCodigo
    
        'grava as regras da planilha
        lErro = ComissoesInpalRegras_Grava_EmTrans(objComissoesInpalRegras)
        If lErro <> SUCESSO Then gError 98868
    
    Next
    
    'Monta as regras no formato da tabela ComissoesRegras
    lErro = ComissoesRegras_Gera_Inpal(objComissoesInpalPlan, colComissoesRegras)
    If lErro <> SUCESSO Then gError 98903
    
    'Para cada regra em colComissoesRegras
    For Each objComissoesRegras In colComissoesRegras
    
        'Grava as regras com novo formato na tabela ComissoesRegras
        lErro = CF("ComissoesRegras_Grava_EmTrans", objComissoesRegras)
        If lErro <> SUCESSO Then gError 98869
    
    Next

    'Confirma transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 98864

    ComissoesInpal_Grava = SUCESSO

    Exit Function

Erro_ComissoesInpal_Grava:

    ComissoesInpal_Grava = gErr
    
    Select Case gErr
    
        Case 98863
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
    
        Case 98864
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
    
        Case 98867, 98868, 98869, 98881, 98903
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select
    
    Call Transacao_Rollback
    
    Exit Function
    
End Function

Private Function ComissoesInpalPlan_Grava_EmTrans(ByVal objComissoesInpalPlan As ClassComisInpalPlan, ByVal lPlanilhaIrma As Long) As Long
'faz o lock do vendedor e chama as funcoes que irao inserir ou atualizar efetivamente os dados da planilha
'DEVE SER CHAMADA EM TRANSAÇÃO!
'objComissoesInpalPlan RECEBE (INPUT) os dados a serem gravados
'e RETORNA(OUTPUT) o NumIntDoc da planilha
'lPlanilhaIrma RECEBE(INPUT) a outra planilha, que vai sofrer alteracao

Dim alComando(0 To 5) As Long
Dim iIndex As Integer
Dim lErro As Long, lErro2 As Long, lErro3 As Long
Dim lCod As Long
Dim iVendedor As Integer, iTecnico As Integer
Dim bEhGravacao As Boolean
Dim dAjudaCusto As Double
Dim iTipoAjudaCusto As Integer

On Error GoTo Erro_ComissoesInpalPlan_Grava_EmTrans

    'abertura comando..
    For iIndex = LBound(alComando) To UBound(alComando)
    
        alComando(iIndex) = Comando_Abrir()
        If alComando(iIndex) = 0 Then gError 98870
    
    Next

    'locka o vendedor
    lErro = CF("Vendedor_Lock", alComando(0), objComissoesInpalPlan.iVendedor)
    If lErro <> SUCESSO Then gError 98871
    
    'Obtem um conjunto unitario de registros de ComissoesInpalPlan
    lErro = Comando_ExecutarPos(alComando(1), "SELECT Codigo, Vendedor, Tecnico FROM ComissoesInpalPlan WHERE Codigo = ?", 0, lCod, iVendedor, iTecnico, objComissoesInpalPlan.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 98876
    
    'Tenta se posicionar no 1o registro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98877
    
    'Se o vendedor e o a flag tecnico nao forem alterados, nao precisa fazer as checagens
    If Not (objComissoesInpalPlan.iVendedor = iVendedor And objComissoesInpalPlan.iTecnico = iTecnico) Then
    
        'Verifica se ja existe outro registro com a flag vendedor/tecnico iguais ao que sera gravado
        lErro2 = Comando_ExecutarPos(alComando(3), "SELECT Codigo FROM ComissoesInpalPlan WHERE Vendedor = ? AND Tecnico = ? AND Codigo <> ?", 0, lCod, objComissoesInpalPlan.iVendedor, objComissoesInpalPlan.iTecnico, objComissoesInpalPlan.lCodigo)
        If lErro2 <> AD_SQL_SUCESSO Then gError 98954
        
        'se posiciona no 1o registro...
        lErro2 = Comando_BuscarPrimeiro(alComando(3))
        If lErro2 <> AD_SQL_SUCESSO And lErro2 <> AD_SQL_SEM_DADOS Then gError 98955
    
        'se encontrou => erro
        If lErro2 <> AD_SQL_SEM_DADOS Then gError 98956
            
    End If
    
    'se nao achou a planilha, entao eh insercao
    If lErro <> AD_SQL_SUCESSO Then
    
        'Faz a insercao em ComissaoInpalPlan
        With objComissoesInpalPlan
            lErro = Comando_Executar(alComando(2), "INSERT INTO ComissoesInpalPlan (Codigo, Vendedor, Tecnico, PercComissaoEmissao, PercComissaoBaixa, ComissaoSobreTotal, ComissaoFrete, ComissaoDesp, ComissaoIPI, ComissaoSeguro, AjudaCusto, TipoAjudaCusto) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)", .lCodigo, .iVendedor, .iTecnico, .dPercComissaoEmissao, .dPercComissaoBaixa, .iComissaoSobreTotal, .iComissaoFrete, .iComissaoDesp, .iComissaoIPI, .iComissaoSeguro, .dAjudaCusto, .iTipoAjudaCusto)
        End With
            
        If lErro <> AD_SQL_SUCESSO Then gError 98873
            
    'se achou, entao eh atualizacao
    Else
    
        'faz o lock exclusive da planilha
        lErro = Comando_LockExclusive(alComando(1))
        If lErro <> SUCESSO Then gError 98940
    
        'Atualiza o registro em questao...
        With objComissoesInpalPlan
            lErro = Comando_ExecutarPos(alComando(2), "UPDATE ComissoesInpalPlan SET Vendedor = ?, Tecnico = ?, PercComissaoEmissao = ?, PercComissaoBaixa = ?, ComissaoSobreTotal = ?, ComissaoFrete = ?, ComissaoDesp = ?, ComissaoIPI = ?, ComissaoSeguro =?, AjudaCusto = ?, TipoAjudaCusto = ?", alComando(1), .iVendedor, .iTecnico, .dPercComissaoEmissao, .dPercComissaoBaixa, .iComissaoSobreTotal, .iComissaoFrete, .iComissaoDesp, .iComissaoIPI, .iComissaoSeguro, .dAjudaCusto, .iTipoAjudaCusto)
        End With

        If lErro <> AD_SQL_SUCESSO Then gError 98884

    End If
    
    'verifica se eh necessario atualizar a planilha irma
    'a planilha irma eh a planilha que tem o mesmo codigo do vendedor, mas diferente valor para a flag itecnico
    'tivemos que fazer esse upgrade porque senao poderiamos ter valores e tipos diferentes de ajuda de custo para
    'um mesmo vendedor, o que causaria problema na parte de comissoes...
    'aparentemente, esta no local errado, deveria estar no cadastro de vendedores...
    'mas como eh uma coisa especifica da inpal, nao quisemos mexer no cadastro de vendedor
    If lPlanilhaIrma > 0 Then
        
        'obtem a planilha necessaria para a atualizacao
        'a ajudacusto e o tipo de ajuda custo obtidas sao obsoletas, so usei como variavel para selecionar a linha..
        lErro = Comando_ExecutarPos(alComando(4), "SELECT AjudaCusto, TipoAjudaCusto FROM ComissoesInpalPlan WHERE Codigo = ?", 0, dAjudaCusto, iTipoAjudaCusto, lPlanilhaIrma)
        If lErro <> AD_SQL_SUCESSO Then gError 101694
        
        'posiciona alComando(4) sobre a planilha que sera manipulada
        lErro = Comando_BuscarPrimeiro(alComando(4))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 101695
        
        'atualiza a planilha caso tenha encontrado
        If lErro <> AD_SQL_SEM_DADOS Then
            
            'atualiza a planilha irma (Ajuda de Custo e Tipo da Ajuda de Custo)
            lErro = Comando_ExecutarPos(alComando(5), "UPDATE ComissoesInpalPlan SET AjudaCusto = ?, TipoAjudaCusto = ?", alComando(4), objComissoesInpalPlan.dAjudaCusto, objComissoesInpalPlan.iTipoAjudaCusto)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 101666
            
        End If
    
    End If
    
    'fechando comando...
    For iIndex = LBound(alComando) To UBound(alComando)
    
        Call Comando_Fechar(alComando(iIndex))
    
    Next
    
    ComissoesInpalPlan_Grava_EmTrans = SUCESSO
    
    Exit Function
    
Erro_ComissoesInpalPlan_Grava_EmTrans:

    ComissoesInpalPlan_Grava_EmTrans = gErr

    Select Case gErr
    
        Case 98870
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 98871, 98872, 98875
        
        Case 98873
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_COMISSOESINPALPLAN", gErr, objComissoesInpalPlan.lCodigo)
        
        Case 98876, 98877, 98954, 98955, 101660, 101661, 101694, 101695
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMISSOESINPALPLAN", gErr, objComissoesInpalPlan.lCodigo)
    
        Case 98884
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_COMISSOESINPALPLAN", gErr, objComissoesInpalPlan.lCodigo)
    
        Case 98940
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_COMISSOESINPALPLAN", gErr, objComissoesInpalPlan.lCodigo)
    
        Case 98956
            Call Rotina_Erro(vbOKOnly, "ERRO_PLANILHA_JA_EXISTENTE", gErr, lCod)
    
        Case 101666
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_COMISSOESINPALPLAN", gErr, lPlanilhaIrma)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr)
            
    End Select
    
    'fechando comando...
    For iIndex = LBound(alComando) To UBound(alComando)
    
        Call Comando_Fechar(alComando(iIndex))
    
    Next
    
    Exit Function

End Function

Private Function ComissoesInpalRegras_Grava_EmTrans(ByVal objComissoesInpalRegras As ClassComisInpalRegras) As Long
'Grava efetivamente os dados das regras no bd..
'DEVE SER CHAMADA EM TRANSACAO!
'objComissoesInpalRegras RECEBE (INPUT) os dados q serao gravados
'e RETORNA (OUTPUT) o numintdoc da regra que sera excluida

Dim lComando As Long
Dim lErro As Long
Dim lNumIntDoc As Long

On Error GoTo Erro_ComissoesInpalRegras_Grava_EmTrans

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 98874

    'Faz Lock nas tabelas relacionadas
    lErro = CF("ComissoesInpalRegras_Le_Tabelas_Lock", objComissoesInpalRegras)
    If lErro <> SUCESSO Then gError 98880
    
    'Obtem o NumIntDoc da regra que sera gravada...
    lErro = CF("Config_ObterAutomatico_EmTrans", "FATConfig", "NUM_PROX_COMISSOESINPALREGRAS", "ComissoesInpalRegras", "NumIntDoc", lNumIntDoc)
    If lErro <> SUCESSO Then gError 98882
    
    'coloca o numintdoc no obj
    objComissoesInpalRegras.lNumIntDoc = lNumIntDoc
    
    'faz a insercao
    With objComissoesInpalRegras
        lErro = Comando_Executar(lComando, "INSERT INTO ComissoesInpalRegras (NumIntDoc, CodPlanilha, RegiaoVenda, Cliente, FilialCliente, CategoriaProduto, ItemCatProduto, PercTabelaA, PercTabelaB) VALUES (?,?,?,?,?,?,?,?,?)", .lNumIntDoc, .lCodPlanilha, .iRegiaoVenda, .lCliente, .iFilialCliente, .sCategoriaProduto, .sItemCatProduto, .dPercTabelaA, .dPercTabelaB)
    End With
    
    If lErro <> AD_SQL_SUCESSO Then gError 98885
    
    'fecha o comando
    Call Comando_Fechar(lComando)
    
    ComissoesInpalRegras_Grava_EmTrans = SUCESSO
    
    Exit Function
    
Erro_ComissoesInpalRegras_Grava_EmTrans:

    ComissoesInpalRegras_Grava_EmTrans = gErr
    
    Select Case gErr
    
        Case 98874
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 98880, 98882
        
        Case 98885
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_COMISSOESINPALREGRAS", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr)

    End Select
    
    'fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Private Function ComissoesRegras_Gera_Inpal(ByVal objComissoesInpalPlan As ClassComisInpalPlan, ByVal colComissoesRegras As Collection) As Long
'Monta as regras para serem gravadas em ComissoesRegras
'objComissoesInpalPlan RECEBE (INPUT) as regras da planilha
'colComissoesRegras RETORNA (OUTPUT) as regras no formato da tabela ComissoesRegras e na ordem adequada

Dim lErro As Long
Dim objComissoesInpalRegras As ClassComisInpalRegras
Dim objComissoesRegras As ClassComissoesRegras
Dim sValorBase As String
Dim iIndex As Integer

On Error GoTo Erro_ComissoesRegras_Gera_Inpal
        
    'monta o conteudo que sera inserido no campo: ValorBase
    lErro = ComissoesRegras_Inpal_Gera_ValorBase(objComissoesInpalPlan, sValorBase)
    If lErro <> SUCESSO Then gError 98890
    
    'Para cada regra na colecao
    For Each objComissoesInpalRegras In objComissoesInpalPlan.colComissoesInpalRegras
        
        'aloca uma area de memoria e faz com que objComissoesRegras aponte para a mesma
        Set objComissoesRegras = New ClassComissoesRegras
        
        'soma o peso da linha à regra
        'utiliza o indice e o numero de itens da colecao para calcular o peso
        'abstracao da formula: peso = totalRegras - indice
        objComissoesRegras.lOrdenacao = objComissoesInpalPlan.colComissoesInpalRegras.Count - iIndex
        
        'incrementa o indice
        iIndex = iIndex + 1
        
        'Monta o conteudo que sera inserido no campo: Regra
        lErro = ComissoesRegras_Inpal_Gera_Regra(objComissoesInpalRegras, objComissoesRegras)
        If lErro <> SUCESSO Then gError 98901
        
        'Se a regra será utilizada para um vendedor não técnico
        If objComissoesInpalPlan.iTecnico <> VENDEDOR_INDIRETO Then
        
            'Inclui a condição do vendedor na regra
            objComissoesRegras.sRegra = MNEMONICO_COMISSOES_VENDEDOR & OPERADOR_IGUAL & objComissoesInpalPlan.iVendedor & E_ENTRE_ESPACOS & objComissoesRegras.sRegra
        
        End If

        'guarda o valor base no objComissoesRegras
        objComissoesRegras.sValorBase = sValorBase
    
        'Monta o conteudo que sera inserido no campo: PercComissao
        lErro = ComissoesRegras_Inpal_Gera_PercComissao(objComissoesInpalRegras, objComissoesRegras)
        If lErro <> SUCESSO Then gError 98902
    
        'se a flag tecnico estiver marcada
        If objComissoesInpalPlan.iTecnico = VENDEDOR_INDIRETO Then
            
            'seta flag VendedorIndireto = a flag tecnico
            objComissoesRegras.iVendedorIndireto = VENDEDOR_INDIRETO
            
            'guarda o codigo do vendedor em objComissoesRegras
            objComissoesRegras.iVendedor = objComissoesInpalPlan.iVendedor
            
        End If
        
        'guarda o numintdoc da regra na tabela da inpal no objcomissoesregras
        objComissoesRegras.lNumIntDocOrigem = objComissoesInpalRegras.lNumIntDoc
        
        'Guarda o percentual da emissao em objComissoesRegras
        objComissoesRegras.dPercComissaoEmiss = objComissoesInpalPlan.dPercComissaoEmissao
        
        'guarda o vendedor
        objComissoesRegras.iVendedor = objComissoesInpalPlan.iVendedor
        
        'adiciona a comissao na colecao
        colComissoesRegras.Add objComissoesRegras
    
    Next
    
    'ordena a colecao a partir do campo ordenacao
    Call Ordena_Regras(colComissoesRegras, colComissoesRegras.Count)

    ComissoesRegras_Gera_Inpal = SUCESSO
    
    Exit Function
    
Erro_ComissoesRegras_Gera_Inpal:

    ComissoesRegras_Gera_Inpal = gErr
    
    Select Case gErr
        
        Case 98890, 98901, 98902
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
            
    End Select

    Exit Function

End Function

Private Sub Ordena_Regras(ByVal colRegras As Collection, iFim As Integer)
'sub recursiva que ordena uma colecao com objetos da Classe ClassComissoesRegras
'ordena de forma decrescente levando em conta a property "lOrdenacao"
'colRegras RECEBE (INPUT) as regras de forma desordenada
'colRegras RETORNA (OUTPUT) as regras de forma ordenada
'iFim RECEBE (INPUT) o indice do ultimo elemento da colecao

Dim iIndex As Integer
Dim iMenor As Integer
Dim objMenor As ClassComissoesRegras
Dim objAux As ClassComissoesRegras

    'caso trivial..
    'se so tem 1 elemento ou nenhum, nao ha o q ordenar, logo.. retorna
    If iFim = 1 Or iFim = 0 Then Exit Sub
    
    'inicializa indice de menor com o indice do ultimo elemento
    iMenor = iFim
        
    'inicializa objmenor com o elemento do fim
    Set objMenor = colRegras(iFim)
        
    'loop para achar o menor valor da colecao
    For iIndex = iFim - 1 To 1 Step -1

        'faz com que a variavel auxiliar (objaux) aponte para o valor corrente da colecao
        Set objAux = colRegras(iIndex)
        
        'verifica se o peso de objAux eh menor do q o peso de menor
        If objAux.lOrdenacao < objMenor.lOrdenacao Then
        
            'se for, entao objmenor passa a ser objaux
            Set objMenor = objAux
            
            'e o indice do menor passa a ser o indice de aux
            iMenor = iIndex
                    
        End If
       
    Next
    
    'ja de posse do menor elemento, faco com que ele seja o ultimo da colecao...
    'primeiramente, removo ele da colecao
    colRegras.Remove (iMenor)
    
    'depois disso, adiciono ele no fim
    colRegras.Add objMenor, , , iFim - 1

    'Agora ja possuo uma colecao parcialmente ordenada... chamo essa mesma funcao para
    'resolver o problema no restante da colecao..
    Call Ordena_Regras(colRegras, iFim - 1)

End Sub

Private Function ComissoesRegras_Inpal_Gera_Regra(ByVal objComissoesInpalRegras As ClassComisInpalRegras, ByVal objComissoesRegras As ClassComissoesRegras) As Long
'Monta a regra de acordo com o conteudo do objcomissoesinpalregras
'objComissoesInpalRegras RECEBE (INPUT) as regras que serao utilizadas na formacao da Regra e no calculo do peso
'objComissoesRegras RETORNA (OUTPUT) a clausula da regra e acumula o peso

Dim sRegra As String
Dim lPeso As Long

On Error GoTo Erro_ComissoesRegras_Inpal_Gera_Regra

    'se a regiao estiver preenchida
    If objComissoesInpalRegras.iRegiaoVenda > 0 Then
        
        'inclui mnemonico de regiao na regra
        sRegra = E_ENTRE_ESPACOS & MNEMONICO_COMISSOES_REGIAO & OPERADOR_IGUAL & objComissoesInpalRegras.iRegiaoVenda
        
        'acerta o formato da regra
        Call Regra_Trim(sRegra)
        
        'soma 100 ao peso da regra
        lPeso = PESO_COMISSOES_REGIAO
    
    End If
        
    'se o cliente estiver preenchido
    If objComissoesInpalRegras.lCliente > 0 Then
        
        'inclui mnemonico Cliente na regra
        sRegra = sRegra & E_ENTRE_ESPACOS & MNEMONICO_COMISSOES_CLIENTE & OPERADOR_IGUAL & objComissoesInpalRegras.lCliente
        
        'acerta o formato da regra
        Call Regra_Trim(sRegra)
        
        'soma 1000 ao peso da regra
        lPeso = lPeso + PESO_COMISSOES_CLIENTE
        
    End If
    
    'se a filial do cliente estiver preenchida
    If objComissoesInpalRegras.iFilialCliente > 0 Then
    
        'inclui o mnemonico de filial na regra
        sRegra = sRegra & E_ENTRE_ESPACOS & MNEMONICO_COMISSOES_FILIALCLI & OPERADOR_IGUAL & objComissoesInpalRegras.iFilialCliente
        
        'acerta o formato da regra
        Call Regra_Trim(sRegra)
        
        'soma 10000 ao peso da regra
        lPeso = lPeso + PESO_COMISSOES_FILIALCLI
        
    End If
    
    'se a categoria estiver preenchida
    If Len(Trim(objComissoesInpalRegras.sCategoriaProduto)) > 0 Then
    
        'inclui o mnemonico itemcatproduto na regra
        sRegra = sRegra & E_ENTRE_ESPACOS & MNEMONICO_COMISSOES_ITEMCATPRODUTO & CARACTER_TODAS_LINHAS_GRID & "(" & """" & objComissoesInpalRegras.sCategoriaProduto & """" & ")" & OPERADOR_IGUAL & """" & objComissoesInpalRegras.sItemCatProduto & """"
                
        'acerta o formato da regra
        Call Regra_Trim(sRegra)
        
        'soma 10 ao peso da regra
        lPeso = lPeso + PESO_COMISSOES_ITEMCATPRODUTO
    
    End If
    
    'atribui a regra ao obj removendo o "E com espacos" no inicio da regra
    objComissoesRegras.sRegra = Mid(sRegra, 4)
    
    'acumula o peso da regra no obj
    objComissoesRegras.lOrdenacao = objComissoesRegras.lOrdenacao + lPeso

    ComissoesRegras_Inpal_Gera_Regra = SUCESSO
    
    Exit Function
    
Erro_ComissoesRegras_Inpal_Gera_Regra:

    ComissoesRegras_Inpal_Gera_Regra = gErr
    
    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr)
    
    End Select
    
    Exit Function

End Function

Private Sub Regra_Trim(sRegra As String)
'Se o ultimo caracter da regra for um espaco em branco, remove o dito cujo...

    'se o ultimo caracter eh espaco
    If Mid(sRegra, Len(sRegra)) = STRING_ESPACO Then
        
        'regra = regra sem o ultimo caracter
        sRegra = Mid(sRegra, 1, Len(sRegra) - 1)
    
    End If


End Sub

Private Function ComissoesRegras_Inpal_Gera_ValorBase(ByVal objComissoesInpalPlan As ClassComisInpalPlan, sValorBase As String) As Long
'gera o conteudo do valor base de comissoesregras
'objComissoesInpalPlan RECEBE (INPUT) os valores que estavam nas checks da tela...
'sValorBase RETORNA (OUTPUT) o valor construido para o valorbase

On Error GoTo Erro_ComissoesRegras_Inpal_Gera_ValorBase

    'se ComissaoSobreTotal estiver marcado
    If objComissoesInpalPlan.iComissaoSobreTotal > 0 Then
        
        'inclui o mnemonico produto_vlrtotal no valorbase
        sValorBase = MNEMONICO_COMISSOES_PRODUTO_VLRTOTAL & CARACTER_TODAS_LINHAS_GRID & OPERADOR_SOMA

    Else
    
        'inclui o mnemonico produto_vlrvenda no valorbase
        sValorBase = MNEMONICO_COMISSOES_PRODUTO_VLRVENDA & CARACTER_TODAS_LINHAS_GRID & OPERADOR_SUBTRACAO & MNEMONICO_COMISSOES_PRODUTO_VLRDESCONTO & OPERADOR_SOMA
        
        'se ComissaoSobreFrete estiver marcado
        'coloquei o teste ">0".. ia utilizar "=vbchecked"
        If objComissoesInpalPlan.iComissaoFrete > 0 Then
        
            'inclui o mnemonico produto_vlrfrete no valorbase
            sValorBase = sValorBase & MNEMONICO_COMISSOES_PRODUTO_VLRFRETE & CARACTER_TODAS_LINHAS_GRID & OPERADOR_SOMA
            
        End If
        
        'se ComissaoSobreSeguro estiver marcado
        'coloquei o teste ">0".. ia utilizar "=vbchecked"
        If objComissoesInpalPlan.iComissaoSeguro > 0 Then
        
            'inclui o mnemonico produto_vlrseguro no valorbase
            sValorBase = sValorBase & MNEMONICO_COMISSOES_PRODUTO_VLRSEGURO & CARACTER_TODAS_LINHAS_GRID & OPERADOR_SOMA
            
        End If
        
        'se comissaosobreoutrasdesp estiver marcado
        'coloquei o teste ">0".. ia utilizar "=vbchecked"
        If objComissoesInpalPlan.iComissaoDesp > 0 Then
        
            'inclui o mnemonico produto_vlroutrasdesp no valorbase
            sValorBase = sValorBase & MNEMONICO_COMISSOES_PRODUTO_VLROUTRASDESP & CARACTER_TODAS_LINHAS_GRID & OPERADOR_SOMA
            
        End If
        
        'se comissaosobreipi estiver marcado
        'coloquei o teste ">0".. ia utilizar "=vbchecked"
        If objComissoesInpalPlan.iComissaoIPI > 0 Then
        
            'inclui o mnemonico produto_vlripi no valorbase
            sValorBase = sValorBase & MNEMONICO_COMISSOES_PRODUTO_VLRIPI & CARACTER_TODAS_LINHAS_GRID & OPERADOR_SOMA
            
        End If
    
    End If
    
    sValorBase = Mid(sValorBase, 1, Len(sValorBase) - Len(OPERADOR_SOMA))
    
    ComissoesRegras_Inpal_Gera_ValorBase = SUCESSO
    
    Exit Function

Erro_ComissoesRegras_Inpal_Gera_ValorBase:

    ComissoesRegras_Inpal_Gera_ValorBase = gErr
    
    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
            
    End Select
            
    Exit Function

End Function

Private Function ComissoesRegras_Inpal_Gera_PercComissao(ByVal objComissoesInpalRegras As ClassComisInpalRegras, ByVal objComissoesRegras As ClassComissoesRegras) As Long
'gera o conteudo do campo sPercComissao do ojbComissoesRegras
'objComissoesInpalRegras RECEBE (INPUT) os valores de tabela A e tabela B que serao utilizados
'objComissoesRegras RETORNA (OUTPUT) o valor do campo: sPercComissao

On Error GoTo Erro_ComissoesRegras_Inpal_Gera_PercComissao

    'se tabelaB estiver preenchida
    If objComissoesInpalRegras.dPercTabelaB > 0 Then
        
        'guarda a formula para encontrar o % comissao no PercComissao
        objComissoesRegras.sPercComissao = FORMULA_PERCCOMISSAO & SEPARADOR_FORMULA_FUNCAO & Replace(CStr(objComissoesInpalRegras.dPercTabelaA), SEPARADOR_FORMULA_FUNCAO, ".") & SEPARADOR_FORMULA_FUNCAO & Replace(CStr(objComissoesInpalRegras.dPercTabelaB), SEPARADOR_FORMULA_FUNCAO, ".") & ")"
        
    Else
        
        'guarda o valor de Tabela A em PercComissao
        objComissoesRegras.sPercComissao = Replace(CStr(objComissoesInpalRegras.dPercTabelaA), SEPARADOR_FORMULA_FUNCAO, ".")
    
    End If
    
    ComissoesRegras_Inpal_Gera_PercComissao = SUCESSO
    
    Exit Function

Erro_ComissoesRegras_Inpal_Gera_PercComissao:

    ComissoesRegras_Inpal_Gera_PercComissao = gErr
    
    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr)
            
    End Select
            
    Exit Function

End Function

Public Function ComissoesInpalPlan_Exclui(ByVal objComissoesInpalPlan As ClassComisInpalPlan) As Long
'Abre as transacoes e chama a funcao que ira excluir a planilha efetivamente do bd
'objComissoesInpalPlan RECEBE (INPUT) a chave para que se exclua a planilha do BD

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_ComissoesInpalPlan_Exclui

    'Abre transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 98913
    
    'Exclui a planilha do bd
    lErro = ComissoesInpalPlan_Exclui_EmTrans(objComissoesInpalPlan)
    If lErro <> SUCESSO Then gError 98914

    'Confirma transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 98915

    ComissoesInpalPlan_Exclui = SUCESSO
    
    Exit Function
    
Erro_ComissoesInpalPlan_Exclui:

    ComissoesInpalPlan_Exclui = gErr
    
    Select Case gErr
        
        Case 98913
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 98914
        
        Case 98915
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr)
            
    End Select
    
    Call Transacao_Rollback
    
    Exit Function

End Function

Private Function ComissoesInpalPlan_Exclui_EmTrans(ByVal objComissoesInpalPlan As ClassComisInpalPlan) As Long
'Exclui efetivamente uma planilha do bd...
'DEVE SER CHAMADA DENTRO DE TRANSAÇÃO
'objComissoesInpalPlan RECEBE (INPUT) a chave primaria que sera usada para identificar
'o registro a ser excluido

Dim lErro As Long
Dim lComando As Long
Dim lCodigo As Long
Dim lComando2 As Long

On Error GoTo Erro_ComissoesInpalPlan_Exclui_EmTrans

    'Abre comando
    lComando = Comando_Abrir()
    lComando2 = Comando_Abrir()
    If lComando = 0 Or lComando2 = 0 Then gError 98915
        
    'obtem o conjunto unitario de registros que serao excluidos do bd..
    lErro = Comando_ExecutarPos(lComando, "SELECT Codigo FROM ComissoesInpalPlan WHERE Codigo = ?", 0, lCodigo, objComissoesInpalPlan.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 98916

    'se posiciona no 1o registro do conjunto obtido anteriormente
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98917

    'se nao estiver cadastrada
    If lErro <> AD_SQL_SUCESSO Then gError 98918
        
    'faz o lock da planilha
    'o lock utilizado eh o "Exclusive" pois essa tabela eh a principal da tela...
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then gError 98919

    'exclui as regras dessa planilha
    lErro = ComissoesInpalRegras_Exclui_EmTrans(lCodigo)
    If lErro <> SUCESSO Then gError 98920

    'exclui o registro referenciado pelo obj da tabela comissoesinpalplan
    lErro = Comando_ExecutarPos(lComando2, "DELETE FROM ComissoesInpalPlan", lComando)
    If lErro <> AD_SQL_SUCESSO Then gError 98921

    'fecha comando
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)

    ComissoesInpalPlan_Exclui_EmTrans = SUCESSO
    
    Exit Function

Erro_ComissoesInpalPlan_Exclui_EmTrans:

    ComissoesInpalPlan_Exclui_EmTrans = gErr
    
    Select Case gErr
        
        Case 98915
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 98916, 98917
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMISSOESINPALPLAN", gErr, objComissoesInpalPlan.lCodigo)
        
        Case 98918
            Call Rotina_Erro(vbOKOnly, "ERRO_COMISSOESINPALPLAN_NAO_CADASTRADA", gErr, objComissoesInpalPlan.lCodigo)
               
        Case 98919
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_COMISSOESINPALPLAN", gErr, objComissoesInpalPlan.lCodigo)
               
        Case 98920
        
        Case 98921
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_COMISSOESINPALPLAN", gErr, objComissoesInpalPlan.lCodigo)
               
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr)
            
    End Select
    
    'fecha comando
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    
    Exit Function

End Function

Private Function ComissoesInpalRegras_Exclui_EmTrans(ByVal lCodPlanilha As Long) As Long
'Exclui efetivamente as regras do bd associadas a uma planilha...
'DEVE SER CHAMADA EM TRANSAÇÃO!
'lNumIntPlanilha RECEBE (INPUT) a chave estrangeira que representa o relacionamento
'entre a planilha e suas regras, ou seja, é a partir de NumIntPlanilha que se faz a exclusão dos registros

Dim lErro As Long
Dim lComando As Long
Dim lComando2 As Long
Dim lNumIntDoc As Long

On Error GoTo Erro_ComissoesInpalRegras_Exclui_EmTrans

    'Abre comando
    lComando = Comando_Abrir()
    lComando2 = Comando_Abrir()
    If lComando = 0 Or lComando2 = 0 Then gError 98924

    'obtem um conjunto de regras do bd a partir do Numero Interno da Planilha
    lErro = Comando_ExecutarPos(lComando, "SELECT NumIntDoc FROM ComissoesInpalRegras WHERE CodPlanilha = ?", 0, lNumIntDoc, lCodPlanilha)
    If lErro <> AD_SQL_SUCESSO Then gError 98925

    'se posiciona na primeira regra do conjunto obtido anteriormente
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98926

    'Varre o conjunto de regras visando deleta-las do bd
    Do While lErro <> AD_SQL_SEM_DADOS
    
        'deleta a regra corrente do bd
        lErro = Comando_ExecutarPos(lComando2, "DELETE FROM ComissoesInpalRegras", lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98928
    
        'deleta a regra correspondente na tabela ComissoesRegras
        lErro = CF("ComissoesRegras_Exclui_NumIntDocOrigem_EmTrans", lNumIntDoc)
        If lErro <> SUCESSO Then gError 98929
        
        'busca o proximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98947
        
    
    Loop

    'fechando comandos...
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    
    ComissoesInpalRegras_Exclui_EmTrans = SUCESSO
    
    Exit Function

Erro_ComissoesInpalRegras_Exclui_EmTrans:

    ComissoesInpalRegras_Exclui_EmTrans = gErr
    
    Select Case gErr
    
        Case 98924
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 98925, 98926, 98947
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMISSOESINPALREGRAS1", gErr, lCodPlanilha)
    
        Case 98929
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_COMISSOESINPALREGRAS", gErr, lCodPlanilha)
        
        Case 98929
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr)
            
    End Select
    
    'fechando comandos...
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    
    Exit Function

End Function

'Public Function ComissoesInpalPlan_AjudaCusto_Critica(ByVal objComissoesInpalPlan As ClassComisInpalPlan, lPlanilhaIrma As Long) As Long
''Verifica se existe uma planilha irma a planilha em questão, onde será preciso atualizar os campos TipoAjudaCusto e AjudaCusto
''Planilha irmã é uma planilha pertencente ao mesmo vendedor da planilha em questão
''Retorna em lPlanilhaIrma o codigo da planilha que vai sofrer o update
''se a planilhairma for 0, significa que nao precisa fazer o update
''objComissoesInpalPlan eh parametro de INPUT
''lPlanilhaIrma eh parametro de OUTPUT
'
'Dim lComando As Long
'Dim lErro As Long
'Dim lPlanilhaIrmaAux As Long
'Dim dAjudaCusto As Double
'Dim vbRet As VbMsgBoxResult
'
'On Error GoTo Erro_ComissoesInpalPlan_AjudaCusto_Critica
'
'    'abre comando
'    lComando = Comando_Abrir()
'    If lComando = 0 Then gError 101692
'
'    'faz o select q ira retornar a ajuda custo do vendedor
'    lErro = Comando_Executar(lComando, "SELECT Codigo, AjudaCusto FROM ComissoesInpalPlan WHERE Vendedor = ? AND Codigo <> ?", lPlanilhaIrmaAux, dAjudaCusto, objComissoesInpalPlan.iVendedor, objComissoesInpalPlan.lCodigo)
'    If lErro <> AD_SQL_SUCESSO Then gError 101660
'
'    'busca o 1o registro do conjunto obtido anteriormente..
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 101661
'
'    'fecha comando
'    Call Comando_Fechar(lComando)
'
'    'se achou e ajuda de custo do existente eh diferente do q vai ser gravado
'    If lErro = AD_SQL_SUCESSO And Abs(dAjudaCusto - objComissoesInpalPlan.dAjudaCusto) >= DELTA_VALORMONETARIO Then
'
'        'se ajuda de custo do que vai ser gravado for 0, avisar que o outro sera mudado pra 0
'        If objComissoesInpalPlan.dAjudaCusto <= DELTA_VALORMONETARIO Then
'
'            'mostra o aviso
'            vbRet = Rotina_Aviso(vbYesNo, "AJUDACUSTO_ATUALIZACAO_ZERO", lPlanilhaIrmaAux, objComissoesInpalPlan.iVendedor, dAjudaCusto)
'
'            'se a resposta ao aviso anterior foi negativa, dispara erro
'            If vbRet = vbNo Then gError 101691
'
'        End If
'
'    End If
'
'    'retorna o codigo da planilha irma
'    lPlanilhaIrma = lPlanilhaIrmaAux
'
'    ComissoesInpalPlan_AjudaCusto_Critica = SUCESSO
'
'    Exit Function
'
'Erro_ComissoesInpalPlan_AjudaCusto_Critica:
'
'    ComissoesInpalPlan_AjudaCusto_Critica = gErr
'
'    Select Case gErr
'
'        Case 101660, 101661
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA", gErr, "ComissoesInpalPlan")
'
'            'fecha comando
'            Call Comando_Fechar(lComando)
'
'        Case 101691
'            'Erro nao tratado propositalmente... nao eh necessario tratar na rotina chamadora...
'
'        Case 101692
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'            'fecha comando
'            Call Comando_Fechar(lComando)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
'
'            'fecha comando
'            Call Comando_Fechar(lComando)
'
'    End Select
'
'    Exit Function
'
'End Function

'tulio261202
'sub para redimensionar o tamanho de objcontrole para itam
'cuidado AO USAR!!!
'o objetivo da criacao dessa rotina foi a customizacao do tamanho do campo mensagemnotafiscal
'para a inpal
Public Sub Redimensiona_Tam_MsgNF(ByVal objControle As Object)

    'Call MsgBox("redimensionei INPAL")

    'a inpal pediu o dobro do tamanho do campo da nota fiscal
    objControle.MaxLength = STRING_NFISCAL_MENSAGEM * 2

End Sub

Private Function RelPrecosLista_Prepara1(ByVal objRelPrecosLista As ClassRelPrecosLista, ByVal iAno As Integer, ByVal iFilialEmpresa As Integer, ByVal lComando As Long) As Long

Dim lErro As Long, dQtde As Double

On Error GoTo Erro_RelPrecosLista_Prepara1

    '??? ORC 2003 deveria ser obtido de alguma outra maneira
    'obter kgs orcados
    lErro = CF("Produto_ObtemKGsOrcados", iFilialEmpresa, objRelPrecosLista.sProduto, iAno, "ORC 2003", dQtde)
    If lErro <> SUCESSO Then gError 106874
                    
    With objRelPrecosLista
        lErro = Comando_Executar(lComando, "INSERT INTO RelPrecosLista (NumIntRel, Produto, Descricao, Valor1, Valor2, Valor3, Valor4, EmbPadrao, KGsOrcados) VALUES (?,?,?,?,?,?,?,?,?)", _
            .lNumIntRel, .sProduto, .sDescricao, .dValor1, .dValor2, .dValor3, .dValor4, .dEmbPadrao, dQtde)
    End With
    If lErro <> AD_SQL_SUCESSO Then gError 106875
    
    RelPrecosLista_Prepara1 = SUCESSO
     
    Exit Function
    
Erro_RelPrecosLista_Prepara1:

    RelPrecosLista_Prepara1 = gErr
     
    Select Case gErr
          
        Case 106874
        
        Case 106875
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RELPRECOSLISTA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    Exit Function

End Function

Public Function RelPrecosLista_Prepara(ByVal iFilialEmpresa As Integer, ByVal dtDataVigencia As Date, ByVal iTabelaPreco As Integer, lNumIntRel As Long) As Long
'Preenche a tabela RelPrecosLista para a execucao do relatorio "Precos de Lista" e devolve lNumIntRel

Dim lErro As Long, iIndice As Integer
Dim lTransacao As Long, alComando(1 To 2) As Long
Dim sProduto As String, sDescricao As String, dCapacidade As Double, iCodTabela As Integer, dPreco As Double
Dim sProdAnterior As String
Dim objRelPrecosLista As ClassRelPrecosLista

On Error GoTo Erro_RelPrecosLista_Prepara

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 106876
    Next
    
    'Inicia a transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 106877

    'obtem o nº automatico
    lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_PROX_NUMINTREL", lNumIntRel)
    If lErro <> SUCESSO Then gError 119613
    
    sProduto = String(STRING_PRODUTO, 0)
    sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
    
    If iTabelaPreco = 0 Then
        If giSQLTipoOrdParam = SQL_ORD_PARAM_ESQ_DIR Then
            lErro = Comando_Executar(alComando(1), "SELECT T1.CodProduto, Produtos.Descricao, ProdEmbPadrao.Capacidade, T1.CodTabela, T1.Preco FROM TabelasDePrecoItens T1, Produtos, ProdEmbPadrao WHERE T1.FilialEmpresa = ? AND T1.CodProduto = Produtos.Codigo AND Produtos.Codigo = ProdEmbPadrao.CodProduto AND DataVigencia = (SELECT MAX(DataVigencia) FROM TabelasDePrecoItens T2 WHERE T2.CodTabela = T1.CodTabela AND T2.FilialEmpresa = T1.FilialEmpresa AND T2.CodProduto = T1.CodProduto AND T2.CodTabela = T1.CodTabela AND T2.DataVigencia <= ? ) ORDER BY T1.CodProduto", _
                sProduto, sDescricao, dCapacidade, iCodTabela, dPreco, iFilialEmpresa, dtDataVigencia)
        Else
            lErro = Comando_Executar(alComando(1), "SELECT T1.CodProduto, Produtos.Descricao, ProdEmbPadrao.Capacidade, T1.CodTabela, T1.Preco FROM TabelasDePrecoItens T1, Produtos, ProdEmbPadrao WHERE T1.FilialEmpresa = ? AND T1.CodProduto = Produtos.Codigo AND Produtos.Codigo = ProdEmbPadrao.CodProduto AND DataVigencia = (SELECT MAX(DataVigencia) FROM TabelasDePrecoItens T2 WHERE T2.CodTabela = T1.CodTabela AND T2.FilialEmpresa = T1.FilialEmpresa AND T2.CodProduto = T1.CodProduto AND T2.CodTabela = T1.CodTabela AND T2.DataVigencia <= ? ) ORDER BY T1.CodProduto", _
                sProduto, sDescricao, dCapacidade, iCodTabela, dPreco, dtDataVigencia, iFilialEmpresa)
        End If
    Else
        If giSQLTipoOrdParam = SQL_ORD_PARAM_ESQ_DIR Then
            lErro = Comando_Executar(alComando(1), "SELECT T1.CodProduto, Produtos.Descricao, ProdEmbPadrao.Capacidade, T1.CodTabela, T1.Preco FROM TabelasDePrecoItens T1, Produtos, ProdEmbPadrao WHERE T1.CodTabela = ? AND T1.FilialEmpresa = ? AND T1.CodProduto = Produtos.Codigo AND Produtos.Codigo = ProdEmbPadrao.CodProduto AND DataVigencia = (SELECT MAX(DataVigencia) FROM TabelasDePrecoItens T2 WHERE T2.CodTabela = T1.CodTabela AND T2.FilialEmpresa = T1.FilialEmpresa AND T2.CodProduto = T1.CodProduto AND T2.DataVigencia <= ? ) ORDER BY T1.CodProduto", _
                sProduto, sDescricao, dCapacidade, iCodTabela, dPreco, iTabelaPreco, iFilialEmpresa, dtDataVigencia)
        Else
            lErro = Comando_Executar(alComando(1), "SELECT T1.CodProduto, Produtos.Descricao, ProdEmbPadrao.Capacidade, T1.CodTabela, T1.Preco FROM TabelasDePrecoItens T1, Produtos, ProdEmbPadrao WHERE T1.CodTabela = ? AND T1.FilialEmpresa = ? AND T1.CodProduto = Produtos.Codigo AND Produtos.Codigo = ProdEmbPadrao.CodProduto AND DataVigencia = (SELECT MAX(DataVigencia) FROM TabelasDePrecoItens T2 WHERE T2.CodTabela = T1.CodTabela AND T2.FilialEmpresa = T1.FilialEmpresa AND T2.CodProduto = T1.CodProduto AND T2.DataVigencia <= ? ) ORDER BY T1.CodProduto", _
                sProduto, sDescricao, dCapacidade, iCodTabela, dPreco, dtDataVigencia, iTabelaPreco, iFilialEmpresa)
        End If
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 106878
    
    lErro = Comando_BuscarProximo(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106879
        
    If lErro = AD_SQL_SUCESSO Then
    
        Do While lErro = AD_SQL_SUCESSO
        
            'se trocou de produto
            If sProdAnterior <> sProduto Then
    
                'se nao for o 1o registro lido
                If sProdAnterior <> "" Then
                
                    lErro = RelPrecosLista_Prepara1(objRelPrecosLista, Year(dtDataVigencia), iFilialEmpresa, alComando(2))
                    If lErro <> SUCESSO Then gError 106880
    
                End If
                
                Set objRelPrecosLista = New ClassRelPrecosLista
                
                With objRelPrecosLista
                    .lNumIntRel = lNumIntRel
                    .sProduto = sProduto
                    .sDescricao = sDescricao
                    .dEmbPadrao = dCapacidade
                End With
                
                sProdAnterior = sProduto
                
            End If
            
            Select Case iCodTabela
            
                Case 1
                    objRelPrecosLista.dValor1 = dPreco
                
                Case 2
                    objRelPrecosLista.dValor2 = dPreco
                
                Case 3
                    objRelPrecosLista.dValor3 = dPreco
                
                Case 4
                    objRelPrecosLista.dValor4 = dPreco
                
            End Select
            
            lErro = Comando_BuscarProximo(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106881
        
        Loop
        
        lErro = RelPrecosLista_Prepara1(objRelPrecosLista, Year(dtDataVigencia), iFilialEmpresa, alComando(2))
        If lErro <> SUCESSO Then gError 106882
    
    End If
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 106883
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    RelPrecosLista_Prepara = SUCESSO
     
    Exit Function
    
Erro_RelPrecosLista_Prepara:

    RelPrecosLista_Prepara = gErr
     
    Select Case gErr
          
        Case 106876
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 106877
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
    
        Case 106878, 106879, 106881
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PREP_RELPRECOSLISTA", gErr)
        
        Case 106883
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case 106880, 106882
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    Call Transacao_Rollback
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
       
    Exit Function

End Function

Function NFiscaisRaiz_Importar(ByVal colNFs As Collection) As Long
'gera nfs a partir de arquivo gerado pela balança

Dim lErro As Long
Dim lTransacao As Long
Dim lComando As Long
Dim iLinha As Integer
Dim lNFPri As Long
Dim lNFUlt As Long
Dim objNFiscal As ClassNFiscal
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim objContabil As ClassContabil
Dim objCtb As ClassImpRaizCtb, objContabAutomatica As ClassContabAutomatica

On Error GoTo Erro_NFiscaisRaiz_Importar

    'Abre a transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 181920
   
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 181932
   
'    lErro = Conexao_DesativarLocks(DESATIVAR_LOCKS)
'    If lErro <> SUCESSO Then gError 181921
    
    objTipoDocInfo.iCodigo = DOCINFO_NFIEFN
    
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 181922
   
    iLinha = 0
    
    'Para cada Nota
    For Each objNFiscal In colNFs
        
        iLinha = iLinha + 1
                
        'preencher a nf
        lErro = NFiscaisRaiz_Importar1(objNFiscal, objTipoDocInfo, objContabil)
        If lErro <> SUCESSO Then gError 181923
        
        'gravar a nf
        lErro = CF("NFiscalFatEntrada_Grava_EmTrans", objNFiscal, objContabil)
        If lErro <> SUCESSO Then gError 181924
        
        Set objContabAutomatica = New ClassContabAutomatica
        Set objCtb = New ClassImpRaizCtb
        
        'gravar a contabilizacao
        lErro = objContabAutomatica.Inicializa_Contab(objCtb, MODULO_ESTOQUE, "EST", objNFiscal.dtDataEntrada, objNFiscal.dtDataEntrada)
        If lErro <> SUCESSO Then gError 130627
        
        lErro = objContabAutomatica.GeraContabilizacao(objNFiscal)
        If lErro <> SUCESSO Then gError 130628
        
        lErro = objContabAutomatica.Finaliza_Contab()
        If lErro <> SUCESSO Then gError 130629
        
        'grava o link entre a nf e a tabela das pesagens
        lErro = NFiscaisRaiz_Importar4(objNFiscal, lComando)
        If lErro <> SUCESSO Then gError 181923
        
        'para dar msg no final
        If iLinha = 1 Then lNFPri = objNFiscal.lNumNotaFiscal
        If iLinha = colNFs.Count Then lNFUlt = objNFiscal.lNumNotaFiscal
        
    Next
    
    Call Comando_Fechar(lComando)
    
'    lErro = Conexao_DesativarLocks(REATIVAR_LOCKS)
'    If lErro <> SUCESSO Then gError 181925
    
    'COMMIT
    lErro = Transacao_Commit()
    If lErro <> SUCESSO Then gError 181926
    
    'Avisa que a rotina foi finalizada com sucesso
    Call Rotina_Aviso(vbOKOnly, "AVISO_IMPORTACAO_NFRAIZ_SUCESSO", lNFPri, lNFUlt)
    
    NFiscaisRaiz_Importar = SUCESSO
     
    Exit Function
    
Erro_NFiscaisRaiz_Importar:

    NFiscaisRaiz_Importar = gErr
     
    Select Case gErr
          
        Case 181920
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
          
        Case 181921
            Call Rotina_Erro(vbOKOnly, "ERRO_DESATIVACAO_LOCKS", gErr)
            
        Case 181922 To 181924, 130627, 130628, 130629
             
        Case 181925
            Call Rotina_Erro(vbOKOnly, "ERRO_REATIVACAO_LOCKS", gErr)
            
        Case 181926
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case 181932
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
    
    Call Comando_Fechar(lComando)
    
    Call Conexao_DesativarLocks(REATIVAR_LOCKS)
    
    'Desfaz a transação
    Call Transacao_Rollback

    Exit Function

End Function

Private Function NFiscaisRaiz_Importar1(ByVal objNFiscal As ClassNFiscal, ByVal objTipoDocInfo As ClassTipoDocInfo, ByVal objContabil As ClassContabil) As Long
'Preenche os dados default da nota e acerta a parte tributária e da distribuição

Dim lErro As Long
Dim objItemNF As ClassItemNF
Dim objTribTab As ClassTribTab
Dim objItemNFAlocacao As ClassItemNFAlocacao
Dim objProduto As ClassProduto
Dim iAlmoxarifado As Integer
Dim objAlmoxarifado As ClassAlmoxarifado
Dim dFator As Double
Dim iIndice As Integer
Dim iIndice2 As Integer
Dim dFrete As Double

Const NATUREZAOP_INPAL = "1101"
Const TIPOTRIB_INPAL = 521
Const SERIE_INPAL = "1"

On Error GoTo Erro_NFiscaisRaiz_Importar1

    With objNFiscal
        .iTipoNFiscal = objTipoDocInfo.iCodigo
        .iFilialEmpresa = giFilialEmpresa
        .iTipoDocInfo = objTipoDocInfo.iTipo
        .iStatus = STATUS_LANCADO
        .dtDataReferencia = DATA_NULA
        .iFreteRespons = FRETE_EMITENTE
        .dtDataRegistro = gdtDataAtual
        .dtDataSaida = DATA_NULA
        .dtDataVencimento = DATA_NULA
        .sSerie = SERIE_INPAL
        
        Call .Inicializa_Tributacao
        Call .objTributacaoNF.Coloca_Auto
    
        .sNaturezaOp = NATUREZAOP_INPAL
        .objTributacaoNF.sNaturezaOpInterna = NATUREZAOP_INPAL
        .objTributacaoNF.iNaturezaOpManual = VAR_PREENCH_MANUAL
        
        .objTributacaoNF.iTipoTributacao = TIPOTRIB_INPAL
        .objTributacaoNF.iTipoTributacaoManual = VAR_PREENCH_MANUAL

    
    End With
    
    iIndice = 0
    dFrete = 0
    For Each objItemNF In objNFiscal.ColItensNF
    
        iIndice = iIndice + 1
        
        dFrete = dFrete + objItemNF.dQuantidade * StrParaDbl(objNFiscal.objInfoUsu.sTick_CampoUsu2)
    
        Call objItemNF.Inicializa_Tributacao
        Call objItemNF.objTributacaoItemNF.Coloca_Auto
        
        objItemNF.iItem = iIndice
    
        objItemNF.objTributacaoItemNF.sNaturezaOp = NATUREZAOP_INPAL
        objItemNF.objTributacaoItemNF.iNaturezaOpManual = VAR_PREENCH_MANUAL
    
        objItemNF.objTributacaoItemNF.iTipoTributacao = TIPOTRIB_INPAL
        objItemNF.objTributacaoItemNF.iTipoTributacaoManual = VAR_PREENCH_MANUAL
                    
        'Distribuição
        Set objProduto = New ClassProduto
        
        objProduto.sCodigo = objItemNF.sProduto
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 181935
        
        If lErro = 28030 Then gError 181936
        
        objItemNF.sDescricaoItem = objProduto.sDescricao
        objItemNF.sUMEstoque = objProduto.sSiglaUMEstoque
        objItemNF.sUMVenda = objProduto.sSiglaUMVenda
        objItemNF.iClasseUM = objProduto.iClasseUM
        
        If objProduto.iControleEstoque <> PRODUTO_SEM_ESTOQUE Then
            
            lErro = CF("AlmoxarifadoPadrao_Le", objNFiscal.iFilialEmpresa, objProduto.sCodigo, iAlmoxarifado)
            If lErro <> SUCESSO And lErro <> 23796 Then gError 181937
        
            If iAlmoxarifado = 0 Then gError 181938
            
            Set objAlmoxarifado = New ClassAlmoxarifado
            
            objAlmoxarifado.iCodigo = iAlmoxarifado
            
            lErro = CF("Almoxarifado_Le", objAlmoxarifado)
            If lErro <> SUCESSO And lErro <> 25056 Then gError 181939
    
            If lErro = 25056 Then gError 181940
    
             'Converte a unidade de medida de UM de Venda para a UM da linha do Grid
            lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objItemNF.sUnidadeMed, objProduto.sSiglaUMEstoque, dFator)
            If lErro <> SUCESSO Then gError 181941
            
            For iIndice2 = objItemNF.ColAlocacoes.Count To 1 Step -1
                objItemNF.ColAlocacoes.Remove iIndice2
            Next
        
            'Cria distribuição do Item
            Set objItemNFAlocacao = objItemNF.ColAlocacoes.Add(objAlmoxarifado.iCodigo, objAlmoxarifado.sNomeReduzido, objItemNF.dQuantidade * dFator)
            
            objItemNFAlocacao.sUnidadeMed = objProduto.sSiglaUMEstoque
            
        End If
    
    Next

    Set objTribTab = New ClassTribTab
    
    lErro = objTribTab.TributacaoNF_Reset(objNFiscal)
    If lErro <> SUCESSO Then gError 181923
    
    'obtem dados do contrato
    Call objTribTab.TipoNFiscal_Definir(objTipoDocInfo.iCodigo, objTipoDocInfo.sSigla)
    Call objTribTab.DataEmissao_Alterada(objNFiscal.dtDataEmissao)
    Call objTribTab.Serie_Alterada(objNFiscal.sSerie)

    objTribTab.objNFTributacao.iISSInclusoManual = VAR_PREENCH_AUTOMATICO
    
    lErro = gobjTributacao.AtualizaImpostos(objTribTab.mvarobjNFTributacao, 0)
    If lErro <> SUCESSO Then gError 181924
    
    lErro = objTribTab.PreencherNF(objNFiscal)
    If lErro <> SUCESSO Then gError 181925
    
    Set objTribTab = Nothing
    
    objNFiscal.sNaturezaOp = objNFiscal.objTributacaoNF.sNaturezaOpInterna
    
    'acertar valor total e da parcela, considerando os tributos, retencóes,...
    objNFiscal.dValorTotal = objNFiscal.dValorTotal + objNFiscal.objTributacaoNF.dICMSSubstValor + objNFiscal.objTributacaoNF.dIPIValor + objNFiscal.dValorSeguro + objNFiscal.dValorOutrasDespesas + IIf(objNFiscal.objTributacaoNF.iISSIncluso = 0, objNFiscal.objTributacaoNF.dISSValor, 0)
    
    Set objNFiscal.objRastreamento = New ClassRastreamento
    
    'Preenche a parte de pagamento
    lErro = NFiscaisRaiz_Importar2(objNFiscal, dFrete)
    If lErro <> SUCESSO Then gError 181926
    
    'Preenche a parte contábil
    lErro = NFiscaisRaiz_Importar3(objNFiscal, objContabil)
    If lErro <> SUCESSO Then gError 181927
    
    NFiscaisRaiz_Importar1 = SUCESSO
     
    Exit Function
    
Erro_NFiscaisRaiz_Importar1:

    NFiscaisRaiz_Importar1 = gErr
     
    Select Case gErr
    
        Case 181923 To 181927, 181935, 181937, 181939, 181941
        
        Case 181936
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
       
        Case 181938
            Call Rotina_Erro(vbOKOnly, "ERRO_NAO_EXISTE_ALMOXARIFADO", gErr, objProduto.sCodigo)
  
        Case 181940
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", gErr, iAlmoxarifado)
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    Exit Function

End Function

Private Function NFiscaisRaiz_Importar2(ByVal objNFiscal As ClassNFiscal, ByVal dFrete As Double)
'Preenche a parte das parcelas

Dim lErro As Long
Dim dValor1 As Double
Dim dValor2 As Double
Dim iIndice As Integer
Dim dINSSRetido As Double
Dim dtVencimento As Date

Const DOMINGO = 1
Const SEGUNDA = 2
Const TERCA = 3
Const QUARTA = 4
Const QUINTA = 5
Const SEXTA = 6
Const SABADO = 7

Const DOMINGO_CONDPAGTO = 116
Const SEGUNDA_CONDPAGTO = 154
Const TERCA_CONDPAGTO = 110
Const QUARTA_CONDPAGTO = 112
Const QUINTA_CONDPAGTO = 113
Const SEXTA_CONDPAGTO = 114
Const SABADO_CONDPAGTO = 115

Const PORTADOR_CARTEIRA_PR_CH1 = 3
Const PORTADOR_CARTEIRA_PR_CH2 = 4

On Error GoTo Erro_NFiscaisRaiz_Importar2

    If objNFiscal.objTributacaoNF.iINSSRetido = MARCADO Then
        dINSSRetido = objNFiscal.objTributacaoNF.dValorINSS
    Else
        dINSSRetido = 0
    End If

    dValor1 = objNFiscal.dValorTotal - dFrete - dINSSRetido - (objNFiscal.objTributacaoNF.dCOFINSRetido + objNFiscal.objTributacaoNF.dCSLLRetido + objNFiscal.objTributacaoNF.dIRRFValor + objNFiscal.objTributacaoNF.dPISRetido)
    dValor2 = dFrete
    
    Select Case Weekday(objNFiscal.dtDataEmissao)
    
        Case DOMINGO
            objNFiscal.iCondicaoPagto = DOMINGO_CONDPAGTO
            dtVencimento = DateAdd("d", 5, objNFiscal.dtDataEmissao)
    
        Case SEGUNDA
            objNFiscal.iCondicaoPagto = SEGUNDA_CONDPAGTO
            dtVencimento = DateAdd("d", 4, objNFiscal.dtDataEmissao)
    
        Case TERCA
            objNFiscal.iCondicaoPagto = TERCA_CONDPAGTO
            dtVencimento = DateAdd("d", 3, objNFiscal.dtDataEmissao)
    
        Case QUARTA
            objNFiscal.iCondicaoPagto = QUARTA_CONDPAGTO
            dtVencimento = DateAdd("d", 2, objNFiscal.dtDataEmissao)
    
        Case QUINTA
            objNFiscal.iCondicaoPagto = QUINTA_CONDPAGTO
            dtVencimento = DateAdd("d", 8, objNFiscal.dtDataEmissao)
    
        Case SEXTA
            objNFiscal.iCondicaoPagto = SEXTA_CONDPAGTO
            dtVencimento = DateAdd("d", 7, objNFiscal.dtDataEmissao)
    
        Case SABADO
            objNFiscal.iCondicaoPagto = SABADO_CONDPAGTO
            dtVencimento = DateAdd("d", 6, objNFiscal.dtDataEmissao)
    
    End Select

    For iIndice = objNFiscal.ColParcelaPagar.Count To 1 Step -1
        objNFiscal.ColParcelaPagar.Remove iIndice
    Next
    
    'Coloca os valores das Parcelas no Grid Parcelas
    objNFiscal.ColParcelaPagar.Add 0, 0, 1, STATUS_ABERTO, dtVencimento, dtVencimento, dValor1, dValor1, PORTADOR_CARTEIRA_PR_CH1, 1, TIPO_COBRANCA_CARTEIRA, 0, "", ""
    
    If dValor2 <> 0 Then
        objNFiscal.ColParcelaPagar.Add 0, 0, 2, STATUS_ABERTO, dtVencimento, dtVencimento, dValor2, dValor2, PORTADOR_CARTEIRA_PR_CH2, 1, TIPO_COBRANCA_CARTEIRA, 0, "", ""
    End If
    
    NFiscaisRaiz_Importar2 = SUCESSO
    
    Exit Function

Erro_NFiscaisRaiz_Importar2:

    NFiscaisRaiz_Importar2 = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select
       
    Exit Function

End Function

Private Function NFiscaisRaiz_Importar3(objNFiscal As ClassNFiscal, ByVal objContabil As ClassContabil)
'Preenche a parte contabil

Dim lErro As Long

On Error GoTo Erro_NFiscaisRaiz_Importar3

    NFiscaisRaiz_Importar3 = SUCESSO
    
    Exit Function

Erro_NFiscaisRaiz_Importar3:

    NFiscaisRaiz_Importar3 = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select
       
    Exit Function

End Function

Private Function NFiscaisRaiz_Importar4(objNFiscal As ClassNFiscal, ByVal lComando As Long)
'Grava a associação da Nota com o RegTick

Dim lErro As Long

On Error GoTo Erro_NFiscaisRaiz_Importar4

    lErro = Comando_Executar(lComando, "INSERT INTO NF_RAIZ (Indice_RegTick, NumIntDocNF) VALUES (?,?)", objNFiscal.objInfoUsu.lTick_Indice, objNFiscal.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 181934

    NFiscaisRaiz_Importar4 = SUCESSO
    
    Exit Function

Erro_NFiscaisRaiz_Importar4:

    NFiscaisRaiz_Importar4 = gErr

    Select Case gErr
    
        Case 181934
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NF_RAIZ", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select
       
    Exit Function

End Function


Function RelAnaliseVendas_Prepara(ByVal objRelAnaliseVendas As ClassRelAnaliseVendas) As Long

Dim lErro As Long, sSelect As String, iIndice As Integer
Dim lTransacao As Long, alComando(0 To 1) As Long, lNumIntRel As Long
Dim iVendedor As Integer, lCliente As Long, iFilialCli As Integer, sProduto As String, iClasseUM As Integer, sUnidadeMed As String, iFilialEmpresa As Integer, iFaturamento As Integer, iTipoNFiscal As Integer, iTabelaPreco As Integer, dQuantidade As Double, dValorTotal As Double

On Error GoTo Erro_RelAnaliseVendas_Prepara

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 130348
    Next

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 130349

    'obtem o nº automatico
    lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_PROX_REL_ANALISE_VENDAS", lNumIntRel)
    If lErro <> SUCESSO Then gError 130350
    
    objRelAnaliseVendas.lNumIntRel = lNumIntRel
    
    'monta o SELECT
    lErro = RelAnaliseVendas_Prepara1(objRelAnaliseVendas, sSelect)
    If lErro <> SUCESSO Then gError 130351
    
    'executa a preparacao da parte fixa do SELECT
    sProduto = String(STRING_PRODUTO, 0)
    sUnidadeMed = String(STRING_UM_SIGLA, 0)
    lErro = RelAnaliseVendas_Prepara2(alComando(0), sSelect, iVendedor, lCliente, iFilialCli, sProduto, iClasseUM, sUnidadeMed, iFilialEmpresa, iFaturamento, iTipoNFiscal, iTabelaPreco, dQuantidade, dValorTotal)
    If lErro <> SUCESSO Then gError 130352
    
    'complementa a passagem dos parametros que variam de acordo com a selecao do usuario
    'e executa o SELECT p/obtencao das parcelas
    lErro = RelAnaliseVendas_Prepara3(objRelAnaliseVendas, alComando(0))
    If lErro <> SUCESSO Then gError 130353
    
    'processa os registros retornados pelo SELECT executado acima
    lErro = RelAnaliseVendas_Prepara4(objRelAnaliseVendas, alComando, iVendedor, lCliente, iFilialCli, sProduto, iClasseUM, sUnidadeMed, iFilialEmpresa, iFaturamento, iTipoNFiscal, iTabelaPreco, dQuantidade, dValorTotal)
    If lErro <> SUCESSO Then gError 130354
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 130355
    
    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next
   
    RelAnaliseVendas_Prepara = SUCESSO
     
    Exit Function
    
Erro_RelAnaliseVendas_Prepara:

    RelAnaliseVendas_Prepara = gErr
     
    Select Case gErr
          
        Case 130348
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
 
        Case 130349
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    Call Transacao_Rollback
    
    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next
   
    Exit Function

End Function

Private Function RelAnaliseVendas_Prepara1(ByVal objRelAnaliseVendas As ClassRelAnaliseVendas, sSelect As String) As Long
'monta o SELECT em funcao dos parametros passados

Dim lErro As Long
Dim sFrom As String, sWhere As String, sFields As String, sOrderBy As String
Dim sGroupBy As String

On Error GoTo Erro_RelAnaliseVendas_Prepara1
    
    sFields = "NFiscal.FilialEmpresa, FiliaisClientes.Vendedor, NFiscal.Cliente, NFiscal.FilialCli, ItensNFiscal.Produto, Produtos.ClasseUM, ItensNFiscal.UnidadeMed, TiposDocInfo.Faturamento, TiposDocInfo.Codigo, NFiscal.TabelaPreco, " & _
            "SUM(ItensNFiscal.Quantidade) As Quantidade, SUM(ItensNFiscal.Quantidade * ItensNFiscal.PrecoUnitario - ItensNFiscal.ValorDesconto) AS ValorTotal"
    sFrom = " FROM NFiscal, TiposDocInfo, ItensNFiscal, Produtos, FiliaisClientes"
    sWhere = " WHERE NFiscal.DataEmissao >= ? AND NFiscal.DataEmissao <= ? AND NFiscal.NumIntDoc = ItensNFiscal.NumIntNF AND TiposDocInfo.Codigo = NFiscal.TipoNFiscal AND (TiposDocInfo.Faturamento IN (1, 2)) AND ItensNFiscal.Produto = Produtos.Codigo AND NFiscal.Status <> 7  AND (TiposDocInfo.Complementar = 0 OR TiposDocInfo.Codigo IN (?,?)) AND FiliaisClientes.CodCliente = NFiscal.Cliente AND FiliaisClientes.CodFilial = NFiscal.FilialCli"
    sOrderBy = " ORDER BY NFiscal.FilialEmpresa, FiliaisClientes.Vendedor, NFiscal.Cliente, NFiscal.FilialCli, ItensNFiscal.Produto, Produtos.ClasseUM, ItensNFiscal.UnidadeMed"
    sGroupBy = " GROUP BY NFiscal.FilialEmpresa, FiliaisClientes.Vendedor, NFiscal.Cliente, NFiscal.FilialCli, ItensNFiscal.Produto, Produtos.ClasseUM, ItensNFiscal.UnidadeMed, TiposDocInfo.Faturamento, TiposDocInfo.Codigo, NFiscal.TabelaPreco"
    
    If objRelAnaliseVendas.iFilialEmpresa <> EMPRESA_TODA Then
        sWhere = sWhere & " AND NFiscal.FilialEmpresa = ?"
    End If
    
    If objRelAnaliseVendas.iVendedorDe <> 0 Then
        sWhere = sWhere & " AND FiliaisClientes.Vendedor >= ?"
    End If
    
    If objRelAnaliseVendas.iVendedorAte <> 0 Then
        sWhere = sWhere & " AND FiliaisClientes.Vendedor <= ?"
    End If
        
    If objRelAnaliseVendas.lClienteDe <> 0 Then
        sWhere = sWhere & " AND NFiscal.Cliente >= ?"
    End If
    
    If objRelAnaliseVendas.lClienteAte <> 0 Then
        sWhere = sWhere & " AND NFiscal.Cliente <= ?"
    End If
        
    If Len(Trim(objRelAnaliseVendas.sProdutoDe)) <> 0 Then
        sWhere = sWhere & " AND ItensNFiscal.Produto >= ?"
    End If
    
    If Len(Trim(objRelAnaliseVendas.sProdutoAte)) <> 0 Then
        sWhere = sWhere & " AND ItensNFiscal.Produto <= ?"
    End If
        
    sSelect = "SELECT " & sFields & sFrom & sWhere & sGroupBy & sOrderBy
    
    RelAnaliseVendas_Prepara1 = SUCESSO
     
    Exit Function
    
Erro_RelAnaliseVendas_Prepara1:

    RelAnaliseVendas_Prepara1 = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    Exit Function

End Function

Private Function RelAnaliseVendas_Prepara2(ByVal lComando As Long, ByVal sSelect As String, viVendedor As Variant, vlCliente As Variant, viFilialCli As Variant, vsProduto As Variant, viClasseUM As Variant, vsUnidadeMed As Variant, viFilialEmpresa As Variant, viFaturamento As Variant, viTipoNFiscal As Variant, viTabelaPreco As Variant, vdQuantidade As Variant, vdValorTotal As Variant) As Long
'executa a preparacao da parte fixa do SELECT

Dim lErro As Long

On Error GoTo Erro_RelAnaliseVendas_Prepara2
    
    lErro = Comando_PrepararInt(lComando, sSelect)
    If lErro <> AD_SQL_SUCESSO Then gError 130356

    lErro = Comando_BindVarInt(lComando, viFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 130357

    lErro = Comando_BindVarInt(lComando, viVendedor)
    If lErro <> AD_SQL_SUCESSO Then gError 130358

    lErro = Comando_BindVarInt(lComando, vlCliente)
    If lErro <> AD_SQL_SUCESSO Then gError 130359

    lErro = Comando_BindVarInt(lComando, viFilialCli)
    If lErro <> AD_SQL_SUCESSO Then gError 130360

    lErro = Comando_BindVarInt(lComando, vsProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 130361

    lErro = Comando_BindVarInt(lComando, viClasseUM)
    If lErro <> AD_SQL_SUCESSO Then gError 130362

    lErro = Comando_BindVarInt(lComando, vsUnidadeMed)
    If lErro <> AD_SQL_SUCESSO Then gError 130363

    lErro = Comando_BindVarInt(lComando, viFaturamento)
    If lErro <> AD_SQL_SUCESSO Then gError 130364

    lErro = Comando_BindVarInt(lComando, viTipoNFiscal)
    If lErro <> AD_SQL_SUCESSO Then gError 130365

    lErro = Comando_BindVarInt(lComando, viTabelaPreco)
    If lErro <> AD_SQL_SUCESSO Then gError 130366

    lErro = Comando_BindVarInt(lComando, vdQuantidade)
    If lErro <> AD_SQL_SUCESSO Then gError 130367

    lErro = Comando_BindVarInt(lComando, vdValorTotal)
    If lErro <> AD_SQL_SUCESSO Then gError 130368

    RelAnaliseVendas_Prepara2 = SUCESSO
     
    Exit Function
    
Erro_RelAnaliseVendas_Prepara2:

    RelAnaliseVendas_Prepara2 = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    Exit Function

End Function

Private Function RelAnaliseVendas_Prepara3(ByVal objRelAnaliseVendas As ClassRelAnaliseVendas, ByVal lComando As Long) As Long
'complementa a passagem dos parametros que variam de acordo com a selecao do usuario
'e executa o SELECT p/obtencao das parcelas

Dim lErro As Long
Dim viFilialEmpresa As Variant, viVendedorDe As Variant, viVendedorAte As Variant, vlClienteDe As Variant, vlClienteAte As Variant, vsProdutoDe As Variant, vsProdutoAte As Variant
Dim vdtDataEmissaoDe As Variant, vdtDataEmissaoAte As Variant, viTipoComplPreco1 As Variant, viTipoComplPreco2 As Variant

On Error GoTo Erro_RelAnaliseVendas_Prepara3
    
    vdtDataEmissaoDe = objRelAnaliseVendas.dtDataEmissaoDe
    lErro = Comando_BindVarInt(lComando, vdtDataEmissaoDe)
    If lErro <> AD_SQL_SUCESSO Then gError 130369
    
    vdtDataEmissaoAte = objRelAnaliseVendas.dtDataEmissaoAte
    lErro = Comando_BindVarInt(lComando, vdtDataEmissaoAte)
    If lErro <> AD_SQL_SUCESSO Then gError 130369
    
    viTipoComplPreco1 = CInt(DOCINFO_NFISCP)
    lErro = Comando_BindVarInt(lComando, viTipoComplPreco1)
    If lErro <> AD_SQL_SUCESSO Then gError 130369
    
    viTipoComplPreco2 = CInt(DOCINFO_NFISFCP)
    lErro = Comando_BindVarInt(lComando, viTipoComplPreco2)
    If lErro <> AD_SQL_SUCESSO Then gError 130369
    
    If objRelAnaliseVendas.iFilialEmpresa <> EMPRESA_TODA Then
        viFilialEmpresa = objRelAnaliseVendas.iFilialEmpresa
        lErro = Comando_BindVarInt(lComando, viFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 130369
    End If
    
    If objRelAnaliseVendas.iVendedorDe <> 0 Then
        viVendedorDe = objRelAnaliseVendas.iVendedorDe
        lErro = Comando_BindVarInt(lComando, viVendedorDe)
        If lErro <> AD_SQL_SUCESSO Then gError 130370
    End If
    
    If objRelAnaliseVendas.iVendedorAte <> 0 Then
        viVendedorAte = objRelAnaliseVendas.iVendedorAte
        lErro = Comando_BindVarInt(lComando, viVendedorAte)
        If lErro <> AD_SQL_SUCESSO Then gError 130371
    End If
        
    If objRelAnaliseVendas.lClienteDe <> 0 Then
        vlClienteDe = objRelAnaliseVendas.lClienteDe
        lErro = Comando_BindVarInt(lComando, vlClienteDe)
        If lErro <> AD_SQL_SUCESSO Then gError 130372
    End If
    
    If objRelAnaliseVendas.lClienteAte <> 0 Then
        vlClienteAte = objRelAnaliseVendas.lClienteAte
        lErro = Comando_BindVarInt(lComando, vlClienteAte)
        If lErro <> AD_SQL_SUCESSO Then gError 130373
    End If
        
    If Len(Trim(objRelAnaliseVendas.sProdutoDe)) <> 0 Then
        vsProdutoDe = objRelAnaliseVendas.sProdutoDe
        lErro = Comando_BindVarInt(lComando, vsProdutoDe)
        If lErro <> AD_SQL_SUCESSO Then gError 130374
    End If
    
    If Len(Trim(objRelAnaliseVendas.sProdutoAte)) <> 0 Then
        vsProdutoAte = objRelAnaliseVendas.sProdutoAte
        lErro = Comando_BindVarInt(lComando, vsProdutoAte)
        If lErro <> AD_SQL_SUCESSO Then gError 130375
    End If
    
    lErro = Comando_ExecutarInt(lComando)
    If lErro <> AD_SQL_SUCESSO Then gError 130376

    RelAnaliseVendas_Prepara3 = SUCESSO
     
    Exit Function
    
Erro_RelAnaliseVendas_Prepara3:

    RelAnaliseVendas_Prepara3 = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    Exit Function

End Function

Private Function RelAnaliseVendas_Prepara4(ByVal objRelAnaliseVendas As ClassRelAnaliseVendas, alComando() As Long, iVendedor As Integer, lCliente As Long, iFilialCli As Integer, sProduto As String, iClasseUM As Integer, sUnidadeMed As String, iFilialEmpresa As Integer, iFaturamento As Integer, iTipoNFiscal As Integer, iTabelaPreco As Integer, dQuantidade As Double, dValorTotal As Double) As Long

Dim lErro As Long, dFator As Double
Dim iVendedorAnt As Integer, lClienteAnt As Long, iFilialCliAnt As Integer, sProdutoAnt As String, iFilialEmpresaAnt As Integer, iTabelaPrecoAnt As Integer
Dim dQuantidadeAcum As Double, dValorTotalAcum As Double, iFilialEmpresaCalc As Integer
Dim dQuantidadeDevAcum As Double, dValorTotalDevAcum As Double, bFiltraLucroDe As Boolean, bFiltraLucroAte As Boolean, dLucroDe As Double, dLucroAte As Double
Dim objMargContr As New ClassMargContr, dPrecoPraticadoInformado As Double, bInsereRegistro As Boolean

On Error GoTo Erro_RelAnaliseVendas_Prepara4

    If Len(Trim(objRelAnaliseVendas.sLucroDe)) <> 0 Or Len(Trim(objRelAnaliseVendas.sLucroDe)) <> 0 Then
    
        If Len(Trim(objRelAnaliseVendas.sLucroDe)) <> 0 Then
            bFiltraLucroDe = True
            dLucroDe = StrParaDbl(objRelAnaliseVendas.sLucroDe)
        End If
        
        If Len(Trim(objRelAnaliseVendas.sLucroAte)) <> 0 Then
            bFiltraLucroAte = True
            dLucroAte = StrParaDbl(objRelAnaliseVendas.sLucroAte)
        End If
        
    End If
    
    lErro = Comando_BuscarProximo(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130377

    If lErro = AD_SQL_SEM_DADOS Then gError 130378

    With objMargContr
        .iColunaCli = 3
        .iColunaPraticado = 4
        .iLinhaPercMarg = 17
        .iLinhaPercLucro = 20
        .iLinhaValorMarg = 21
    End With

    iVendedorAnt = iVendedor
    lClienteAnt = lCliente
    iFilialCliAnt = iFilialCli
    sProdutoAnt = sProduto
    iFilialEmpresaAnt = iFilialEmpresa
    iTabelaPrecoAnt = iTabelaPreco
            
    Do While lErro = AD_SQL_SUCESSO

        'se trocou algo que compõe a chave do registro em RelAnaliseVendas
        If iVendedorAnt <> iVendedor Or lClienteAnt <> lCliente Or iFilialCliAnt <> iFilialCli Or sProdutoAnt <> sProduto Or iFilialEmpresaAnt <> iFilialEmpresa Then
        
            If dQuantidadeAcum <> 0 Then
                
                dPrecoPraticadoInformado = Arredonda_Moeda(dValorTotalAcum / dQuantidadeAcum, 2)
            
                With objMargContr
                
                    .dQuantidade = dQuantidadeAcum
                    .iCodFilial = iFilialCliAnt
                    .iFilialFaturamento = iFilialEmpresaAnt
                    .lCodCliente = lClienteAnt
                    .sProduto = sProdutoAnt
                    .iTabelaPreco = iTabelaPrecoAnt
                    .sUnidadeMed = "KG"
                    .dPrecoPraticadoInformado = dPrecoPraticadoInformado
                    
                End With
    
                lErro = CF("MargContr_Calcula", objMargContr)
                If lErro <> SUCESSO Then gError 130382
                
                bInsereRegistro = True
                
                If bFiltraLucroDe Or bFiltraLucroDe Then
                
                    If bFiltraLucroDe Then
                    
                        Select Case objRelAnaliseVendas.iTipoFiltroProduto
                        
                            Case 1 '% lucro
                                If (objMargContr.dUltPercLucroPrat - dLucroDe) < DELTA_VALORMONETARIO2 Then bInsereRegistro = False
                                
                            Case 2 '% marg
                                If (objMargContr.dUltPercMargPrat - dLucroDe) < DELTA_VALORMONETARIO2 Then bInsereRegistro = False
                            
                            Case 3 'R$ marg
                                If (objMargContr.dUltValorMargPrat - dLucroDe) < DELTA_VALORMONETARIO2 Then bInsereRegistro = False
                            
                        End Select
                        
                    End If
                    
                    If bInsereRegistro And bFiltraLucroAte Then
                    
                        Select Case objRelAnaliseVendas.iTipoFiltroProduto
                        
                            Case 1 '% lucro
                                If (objMargContr.dUltPercLucroPrat - dLucroAte) > DELTA_VALORMONETARIO2 Then bInsereRegistro = False
                                
                            Case 2 '% marg
                                If (objMargContr.dUltPercMargPrat - dLucroAte) > DELTA_VALORMONETARIO2 Then bInsereRegistro = False
                            
                            Case 3 'R$ marg
                                If (objMargContr.dUltValorMargPrat - dLucroAte) > DELTA_VALORMONETARIO2 Then bInsereRegistro = False
                            
                        End Select
                        
                    End If
                    
                End If
                
                If bInsereRegistro Then
                    lErro = Comando_Executar(alComando(1), "INSERT INTO RelAnaliseVendas (NumIntRel, Vendedor, Cliente, Filial, Produto, FilialEmpresa, UnidadeMed, Quantidade, PrecoVenda, PrecoCliente, PercMargem, PercLucro, ValorMargem, QtdeDevolvida, ValorDevolvido, TabelaDePreco) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                        objRelAnaliseVendas.lNumIntRel, iVendedorAnt, lClienteAnt, iFilialCliAnt, sProdutoAnt, iFilialEmpresaAnt, "KG", dQuantidadeAcum, dPrecoPraticadoInformado, objMargContr.dUltPrecoCli, objMargContr.dUltPercMargPrat, objMargContr.dUltPercLucroPrat, objMargContr.dUltValorMargPrat, dQuantidadeDevAcum, dValorTotalDevAcum, iTabelaPrecoAnt)
                    If lErro <> AD_SQL_SUCESSO Then gError 130390
                End If
            
            End If
            
            dQuantidadeAcum = 0
            dValorTotalAcum = 0
            dQuantidadeDevAcum = 0
            dValorTotalDevAcum = 0
            
            iVendedorAnt = iVendedor
            lClienteAnt = lCliente
            iFilialCliAnt = iFilialCli
            sProdutoAnt = sProduto
            iFilialEmpresaAnt = iFilialEmpresa
            
        End If
            
        'Faturada
        If iFaturamento = 1 Then
            dValorTotalAcum = dValorTotalAcum + dValorTotal
            
            If Not UnidadeDeMedida_IgnorarNaVenda(sUnidadeMed) Then
                lErro = CF("UM_Conversao", iClasseUM, sUnidadeMed, "KG", dFator)
                If lErro <> SUCESSO Then
                    MsgBox (CStr(iClasseUM) & " - " & "sUnidadeMed")
                    dFator = 0
                    'gError 130383
                End If

                If iTipoNFiscal <> DOCINFO_NFISCP And iTipoNFiscal <> DOCINFO_NFISFCP Then
                    dQuantidadeAcum = dQuantidadeAcum + (dQuantidade * dFator)
                End If
            End If

        Else
            'Devolvida
            dValorTotalDevAcum = dValorTotalDevAcum + dValorTotal

            If Not UnidadeDeMedida_IgnorarNaVenda(sUnidadeMed) Then
                lErro = CF("UM_Conversao", iClasseUM, sUnidadeMed, "KG", dFator)
                If lErro <> SUCESSO Then gError 130384
                
                dQuantidadeDevAcum = dQuantidadeDevAcum + (dQuantidade * dFator)
            End If
        
        End If
        
        iTabelaPrecoAnt = iTabelaPreco
        
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130385

    Loop
    
    If dQuantidadeAcum <> 0 Then
        dPrecoPraticadoInformado = Arredonda_Moeda(dValorTotalAcum / dQuantidadeAcum, 2)
            
        With objMargContr
        
            .dQuantidade = dQuantidadeAcum
            .iCodFilial = iFilialCliAnt
            .iFilialFaturamento = iFilialEmpresaAnt
            .lCodCliente = lClienteAnt
            .sProduto = sProdutoAnt
            .iTabelaPreco = iTabelaPrecoAnt
            .sUnidadeMed = "KG"
            .dPrecoPraticadoInformado = dPrecoPraticadoInformado
            
        End With
    
        lErro = CF("MargContr_Calcula", objMargContr)
        If lErro <> SUCESSO Then gError 130391
        
        bInsereRegistro = True
        
        If bFiltraLucroDe Or bFiltraLucroDe Then
        
            If bFiltraLucroDe Then
            
                Select Case objRelAnaliseVendas.iTipoFiltroProduto
                
                    Case 1 '% lucro
                        If (objMargContr.dUltPercLucroPrat - dLucroDe) < DELTA_VALORMONETARIO2 Then bInsereRegistro = False
                        
                    Case 2 '% marg
                        If (objMargContr.dUltPercMargPrat - dLucroDe) < DELTA_VALORMONETARIO2 Then bInsereRegistro = False
                    
                    Case 3 'R$ marg
                        If (objMargContr.dUltValorMargPrat - dLucroDe) < DELTA_VALORMONETARIO2 Then bInsereRegistro = False
                    
                End Select
                
            End If
            
            If bInsereRegistro And bFiltraLucroAte Then
            
                Select Case objRelAnaliseVendas.iTipoFiltroProduto
                
                    Case 1 '% lucro
                        If (objMargContr.dUltPercLucroPrat - dLucroAte) > DELTA_VALORMONETARIO2 Then bInsereRegistro = False
                        
                    Case 2 '% marg
                        If (objMargContr.dUltPercMargPrat - dLucroAte) > DELTA_VALORMONETARIO2 Then bInsereRegistro = False
                    
                    Case 3 'R$ marg
                        If (objMargContr.dUltValorMargPrat - dLucroAte) > DELTA_VALORMONETARIO2 Then bInsereRegistro = False
                    
                End Select
                
            End If
            
        End If
        
        If bInsereRegistro Then
        
            lErro = Comando_Executar(alComando(1), "INSERT INTO RelAnaliseVendas (NumIntRel, Vendedor, Cliente, Filial, Produto, FilialEmpresa, UnidadeMed, Quantidade, PrecoVenda, PrecoCliente, PercMargem, PercLucro, ValorMargem, QtdeDevolvida, ValorDevolvido, TabelaDePreco) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                objRelAnaliseVendas.lNumIntRel, iVendedorAnt, lClienteAnt, iFilialCliAnt, sProdutoAnt, iFilialEmpresaAnt, "KG", dQuantidadeAcum, dPrecoPraticadoInformado, objMargContr.dUltPrecoCli, objMargContr.dUltPercMargPrat, objMargContr.dUltPercLucroPrat, objMargContr.dUltValorMargPrat, dQuantidadeDevAcum, dValorTotalDevAcum, iTabelaPrecoAnt)
            If lErro <> AD_SQL_SUCESSO Then gError 130392
           
        End If
        
    End If
    
    RelAnaliseVendas_Prepara4 = SUCESSO
     
    Exit Function
    
Erro_RelAnaliseVendas_Prepara4:

    RelAnaliseVendas_Prepara4 = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    Exit Function

End Function


Private Function MargContr_CalculaDVV(ByVal objMargContr As ClassMargContr, ByVal iColuna As Integer) As Long

Dim lErro As Long, iLinha As Integer, objDVVLinCol As ClassDVVLinCol
Dim colPlanilhas As New Collection, objPlanilhas As ClassPlanilhas, sFormula As String
Dim objContexto As New ClassContextoPlan, dValor As Double
Dim objDVVLinColEsq As ClassDVVLinCol

On Error GoTo Erro_MargContr_CalculaDVV

    If objMargContr.colDVVLin.Count <> 0 Then
    
        For iLinha = 1 To objMargContr.colDVVLin.Count
                
            Set objDVVLinCol = objMargContr.colDVVLinCol(objMargContr.IndDVV(iLinha, iColuna - 1))
            
            If Len(Trim(objDVVLinCol.sFormula)) = 0 Then
            
                If iColuna > 1 Then
                
                    Set objDVVLinColEsq = objMargContr.colDVVLinCol(objMargContr.IndDVV(iLinha, iColuna - 2))
                    sFormula = objDVVLinColEsq.sFormula
                    If Len(Trim(sFormula)) = 0 Then
                    
                        If iColuna > 2 Then
                        
                            Set objDVVLinColEsq = objMargContr.colDVVLinCol(objMargContr.IndDVV(iLinha, iColuna - 3))
                            sFormula = objDVVLinColEsq.sFormula
                            If Len(Trim(sFormula)) = 0 Then sFormula = "0"
                            
                        Else
                        
                            sFormula = "0"
                            
                        End If
                        
                    End If
                Else
                
                    sFormula = "0"
                    
                End If
                
            Else
            
                sFormula = objDVVLinCol.sFormula
                
            End If
            
            Set objPlanilhas = New ClassPlanilhas
            
            With objPlanilhas
                .iTipoPlanilha = PLANILHA_TIPO_DVV
                .iFilialEmpresa = objMargContr.iFilialFaturamento
                .iEscopo = MNEMONICOFPRECO_ESCOPO_PRODUTO
                .iLinha = iLinha
                .sExpressao = sFormula
            End With
            
            colPlanilhas.Add objPlanilhas
        
        Next
        
        With objContexto
            .iFilialFaturamento = objMargContr.iFilialFaturamento
            .sProduto = objMargContr.sProduto
            .dQuantidade = objMargContr.dQuantidade
            .iFilialCli = objMargContr.iCodFilial
            .lCliente = objMargContr.lCodCliente
            .iVendedor = objMargContr.iCodVendedor
            .iTabelaPreco = objMargContr.iTabelaPreco
            .iAno = Year(gdtDataAtual)
            .dtDataCustoFixo = objMargContr.dtDataCF
            .iRotinaOrigem = FORMACAO_PRECO_ANALISE_MARGCONTR_REL
            Set .colComissoes = objMargContr.colComissoes
            .dPrecoPraticado = objMargContr.dPrecoComissoes
            .sNomeRedCliente = objMargContr.sUltNomeRedCli
            .sUM = objMargContr.sUnidadeMed
        End With
        
        'Executa as formulas da planilha de preço. Retorna o valor da planilha em dValor (que é o valor da última linha da planilha) e o valor de cada linha em colPlanilhas.Item(?).dValor
        lErro = CF("Avalia_Expressao_FPreco3", colPlanilhas, dValor, objContexto)
        If lErro <> SUCESSO Then gError 130046
    
        Set objPlanilhas = colPlanilhas.Item(colPlanilhas.Count)
        
        Select Case (iColuna - 1)
        
            Case 1
                objMargContr.sDVVTotal1 = Format(objPlanilhas.dValor * 100, "###,##0.00###")
            Case 2
                objMargContr.sDVVTotal2 = Format(objPlanilhas.dValor * 100, "###,##0.00###")
            Case 3
                objMargContr.sDVVTotal3 = Format(objPlanilhas.dValor * 100, "###,##0.00###")
                
        End Select
    
    End If
    
    MargContr_CalculaDVV = SUCESSO
     
    Exit Function
    
Erro_MargContr_CalculaDVV:

    MargContr_CalculaDVV = gErr
     
    Select Case gErr
          
        Case 130046
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    Exit Function

End Function

Private Function Analise_RecalcularColuna(ByVal objMargContr As ClassMargContr, ByVal iColuna As Integer, ByVal iColunaCli As Integer, ByVal iColunaPraticado As Integer) As Long

Dim lErro As Long, iLinha As Integer, objLinCol As ClassPlanMargContrLinCol, objLin As ClassPlanMargContrLin
Dim colPlanilhas As New Collection, objPlanilhas As ClassPlanilhas, sFormula As String
Dim objContexto As New ClassContextoPlan, dValor As Double, colPlanilhas2 As New Collection
Dim objMnemonicoValor As ClassMnemonicoValor, objContexto2 As New ClassContextoPlan, iCol As Integer
Dim dValorL1 As Double

On Error GoTo Erro_Analise_RecalcularColuna

    'calcular o preco para o cliente se a coluna for iColunaCli
    If iColuna = iColunaCli Then
    
        With objContexto2
            .iFilialFaturamento = objMargContr.iFilialFaturamento
            .sProduto = objMargContr.sProduto
            .dQuantidade = objMargContr.dQuantidade
            .iFilialCli = objMargContr.iCodFilial
            .lCliente = objMargContr.lCodCliente
            .iVendedor = objMargContr.iCodVendedor
            .iTabelaPreco = objMargContr.iTabelaPreco
            .iAno = Year(gdtDataAtual)
            .dtDataCustoFixo = objMargContr.dtDataCF
            .iRotinaOrigem = FORMACAO_PRECO_ANALISE_MARGCONTR_REL
            Set .colComissoes = objMargContr.colComissoes
            .dPrecoPraticado = objMargContr.dPrecoComissoes
            .dTaxaDescPadrao = objMargContr.dTaxaDescPadrao
            .dTaxaValFut = objMargContr.dTaxaValFut
            .dDiasValFut = objMargContr.dDiasValFut
        End With
        
        'guarda os resultados do grid dvv para todas as colunas
        Call DVVTotal_GuardaValores(objMargContr, objContexto2)
        
        For iLinha = 1 To objMargContr.colPlanMargContrLin.Count
        
            Set objLin = objMargContr.colPlanMargContrLin(iLinha)
            Set objLinCol = objMargContr.colPlanMargContrLinCol(objMargContr.IndAnalise(iLinha, iColuna - 1))
            
            sFormula = objLinCol.sFormula
            If iLinha = 1 And Len(Trim(sFormula)) <> 0 Then Exit For
            If Len(Trim(objLin.sFormulaL1)) = 0 And Len(Trim(sFormula)) = 0 Then sFormula = objLin.sFormulaGeral
            
            If Len(Trim(sFormula)) = 0 Then sFormula = "0"
            
            Set objPlanilhas = New ClassPlanilhas
            
            With objPlanilhas
                .iTipoPlanilha = PLANILHA_TIPO_TODOS
                .iFilialEmpresa = objMargContr.iFilialFaturamento
                .iEscopo = MNEMONICOFPRECO_ESCOPO_GERAL
                .iLinha = iLinha
                .sExpressao = sFormula
            End With
            
            colPlanilhas2.Add objPlanilhas
            
            'se possui formula p/calculo da linha 1 e o valor está preenchido...
            If Len(Trim(objLin.sFormulaL1)) <> 0 And ((Len(Trim(objLinCol.sFormula)) <> 0)) Then
            
                Set objPlanilhas = New ClassPlanilhas
                
                With objPlanilhas
                    .iTipoPlanilha = PLANILHA_TIPO_TODOS
                    .iFilialEmpresa = objMargContr.iFilialFaturamento
                    .iEscopo = MNEMONICOFPRECO_ESCOPO_GERAL
                    .iLinha = iLinha + 1
                    .sExpressao = objLin.sFormulaL1
                End With
                
                colPlanilhas2.Add objPlanilhas
        
                'Executa as formulas da planilha de preço. Retorna o valor da planilha em dValor (que é o valor da última linha da planilha) e o valor de cada linha em colPlanilhas.Item(?).dValor
                lErro = CF("Avalia_Expressao_FPreco3", colPlanilhas2, dValor, objContexto2)
                If lErro <> SUCESSO Then gError 106721

                dValorL1 = dValor
                                
                Exit For
                
            End If
            
        Next
        
    End If
    
    If iColuna = iColunaCli Then
     
        objMargContr.dUltPrecoCli = dValorL1
        
    Else
    
        'guarda os resultados do grid dvv para todas as colunas
        Call DVVTotal_GuardaValores(objMargContr, objContexto)
        
        For iLinha = 1 To objMargContr.colPlanMargContrLin.Count
        
            Set objLin = objMargContr.colPlanMargContrLin(iLinha)
            Set objLinCol = objMargContr.colPlanMargContrLinCol(objMargContr.IndAnalise(iLinha, iColuna - 1))
                
            If iLinha = 1 And Len(Trim(objLinCol.sFormula)) = 0 And Len(Trim(objLin.sFormulaGeral)) = 0 And objLin.iEditavel <> 0 Then
            
                'pega conteudo da tela, da propria celula
                sFormula = CStr(dValorL1)
                If Len(Trim(sFormula)) <> 0 And objLin.iFormato = GRID_FORMATO_PERCENTAGEM Then sFormula = CStr(StrParaDbl(sFormula) / 100)
                Call TrocaPontoVirgula(sFormula)
                
            Else
            
                sFormula = objLinCol.sFormula
                If Len(Trim(sFormula)) = 0 Then sFormula = objLin.sFormulaGeral
            
            End If
            
            If Len(Trim(sFormula)) = 0 Then sFormula = "0"
            
            Set objPlanilhas = New ClassPlanilhas
            
            With objPlanilhas
                .iTipoPlanilha = PLANILHA_TIPO_TODOS
                .iFilialEmpresa = objMargContr.iFilialFaturamento
                .iEscopo = MNEMONICOFPRECO_ESCOPO_PRODUTO
                .iLinha = iLinha
                .sExpressao = sFormula
            End With
            
            colPlanilhas.Add objPlanilhas
        
        Next
        
        With objContexto
            .iFilialFaturamento = objMargContr.iFilialFaturamento
            .sProduto = objMargContr.sProduto
            .dQuantidade = objMargContr.dQuantidade
            .iFilialCli = objMargContr.iCodFilial
            .lCliente = objMargContr.lCodCliente
            .iVendedor = objMargContr.iCodVendedor
            .iTabelaPreco = objMargContr.iTabelaPreco
            .iAno = Year(gdtDataAtual)
            .dtDataCustoFixo = objMargContr.dtDataCF
            .iRotinaOrigem = FORMACAO_PRECO_ANALISE_MARGCONTR_REL
            Set .colComissoes = objMargContr.colComissoes
            .dPrecoPraticado = objMargContr.dPrecoComissoes
            .dTaxaDescPadrao = objMargContr.dTaxaDescPadrao
            .dTaxaValFut = objMargContr.dTaxaValFut
            .dDiasValFut = objMargContr.dDiasValFut
        End With
        
        'Executa as formulas da planilha de preço. Retorna o valor da planilha em dValor (que é o valor da última linha da planilha) e o valor de cada linha em colPlanilhas.Item(?).dValor
        lErro = CF("Avalia_Expressao_FPreco3", colPlanilhas, dValor, objContexto)
        If lErro <> SUCESSO Then gError 106721
    
        For Each objPlanilhas In colPlanilhas
        
            If objPlanilhas.iLinha = 1 Then
                objMargContr.dUltPrecoPraticado = StrParaDbl(Format(objPlanilhas.dValor, "###,##0.00###"))
            End If
            
            If objPlanilhas.iLinha = objMargContr.iLinhaPercMarg Then
                objMargContr.dUltPercMargPrat = StrParaDbl(Format(objPlanilhas.dValor * 100, "###,##0.00###"))
            End If
    
            If objPlanilhas.iLinha = objMargContr.iLinhaPercLucro Then
                objMargContr.dUltPercLucroPrat = StrParaDbl(Format(objPlanilhas.dValor * 100, "###,##0.00###"))
            End If
    
            If objPlanilhas.iLinha = objMargContr.iLinhaValorMarg Then
                objMargContr.dUltValorMargPrat = StrParaDbl(Format(objPlanilhas.dValor, "###,##0.00###"))
            End If
    
        Next
    
    End If
    
    Analise_RecalcularColuna = SUCESSO
     
    Exit Function
    
Erro_Analise_RecalcularColuna:

    Analise_RecalcularColuna = gErr
     
    Select Case gErr
          
        Case 106721
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    Exit Function

End Function

Private Sub DVVTotal_GuardaValores(ByVal objMargContr As ClassMargContr, objContexto As ClassContextoPlan)

Dim iCol As Integer, objMnemonicoValor As ClassMnemonicoValor

    'guarda os resultados do grid dvv para todas as colunas
    For iCol = 1 To 3
    
        Set objMnemonicoValor = New ClassMnemonicoValor
        Set objMnemonicoValor.colValor = New Collection
        
        objMnemonicoValor.sMnemonico = "DVVTotal"
        objMnemonicoValor.vParam(1) = CDbl(iCol)
    
        Select Case iCol
        
            Case 1
                objMnemonicoValor.colValor.Add StrParaDbl(StrParaDbl(objMargContr.sDVVTotal1) / 100)
            Case 2
                objMnemonicoValor.colValor.Add StrParaDbl(StrParaDbl(objMargContr.sDVVTotal2) / 100)
            Case 3
                objMnemonicoValor.colValor.Add StrParaDbl(StrParaDbl(objMargContr.sDVVTotal3) / 100)
        
        End Select
        
        objContexto.colMnemonicoValor.Add objMnemonicoValor
        
    Next

End Sub

Private Sub TrocaPontoVirgula(sNumero As String)
Dim iTam As Integer, sResult As String, i As Integer, sCaracter As String

    For i = 1 To Len(sNumero)
    
        sCaracter = Mid(sNumero, i, 1)
        Select Case sCaracter
        
            Case ","
                sCaracter = "."
                
            Case "."
                sCaracter = ""
            
        End Select
        
        sResult = sResult & sCaracter
        
    Next
        
    sNumero = sResult

End Sub

Private Function MargContr_TestaDVV(ByVal objMargContr As ClassMargContr) As Long

Dim lErro As Long, dPercDVV As Double, lComando As Long

On Error GoTo Erro_MargContr_TestaDVV

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 106666
    
    lErro = Comando_Executar(lComando, "SELECT PercDVV FROM DVVClienteProd WHERE FilialEmpresa = ? AND CodCliente = ? AND CodFilial = ? AND Produto = ?", dPercDVV, objMargContr.iFilialFaturamento, objMargContr.lCodCliente, objMargContr.iCodFilial, objMargContr.sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 106667

    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106668
    If lErro <> AD_SQL_SUCESSO Then gError 130055
    
    Call Comando_Fechar(lComando)
    
    MargContr_TestaDVV = SUCESSO
     
    Exit Function
    
Erro_MargContr_TestaDVV:

    MargContr_TestaDVV = gErr
     
    Select Case gErr
          
        Case 106666
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 106667, 106668
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_DVVCLIENTEPROD", gErr)

        Case 130055
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function MargContr_Calcula(ByVal objMargContr As ClassMargContr) As Long

Dim lErro As Long, objTelaComissoes As Object, objCodigoNome As New AdmlCodigoNome
Dim dtDataCFTemp As Date, dTaxaDescPadraoTemp As Double, dTaxaValFutTemp As Double, dDiasValFutTemp As Double
Dim colComissoes As Collection, dPrecoComissoes As Double

On Error GoTo Erro_MargContr_Calcula

    If objMargContr.objTelaComissoes Is Nothing Then
    
        lErro = CF("MargContr_Le_DVV", objMargContr)
        If lErro <> SUCESSO Then gError 130379
        
        lErro = CF("MargContr_Le_Analise", objMargContr)
        If lErro <> SUCESSO Then gError 130380
        
        lErro = Chama_Tela_Nova_Instancia1("ComissoesCalcula", objTelaComissoes)
        If lErro <> SUCESSO Then gError 130381
        
        Set objMargContr.objTelaComissoes = objTelaComissoes
    
    Else
    
        Set objTelaComissoes = objMargContr.objTelaComissoes
        
    End If
    
    'se trocou de filialempresa
    If objMargContr.iUltFilialEmpresa <> objMargContr.iFilialFaturamento Then
        
        lErro = CF("FilialEmpresa_ObtemDataCustoFixo", objMargContr.iFilialFaturamento, dtDataCFTemp)
        If lErro <> SUCESSO Then gError 130386
    
        lErro = CF("CalcMP_ObterValores", objMargContr.iFilialFaturamento, dTaxaDescPadraoTemp, dTaxaValFutTemp, dDiasValFutTemp)
        If lErro <> SUCESSO Then gError 130387
        
        With objMargContr
            .dtDataCF = dtDataCFTemp
            .dTaxaDescPadrao = dTaxaDescPadraoTemp
            .dTaxaValFut = dTaxaValFutTemp
            .dDiasValFut = dDiasValFutTemp
        End With
        objMargContr.iUltFilialEmpresa = objMargContr.iFilialFaturamento
        
    End If

    'se trocou de cliente
    If objMargContr.lUltCliente <> objMargContr.lCodCliente Then
    
        objCodigoNome.lCodigo = objMargContr.lCodCliente
        lErro = CF("Cliente_Le_NomeRed", objCodigoNome)
        If lErro <> SUCESSO And lErro <> 12553 Then gError 130388
        If lErro <> SUCESSO Then gError 130389
        
        objMargContr.sUltNomeRedCli = objCodigoNome.sNome
        
        objMargContr.lUltCliente = objMargContr.lCodCliente
        
    End If

    lErro = MargContr_TestaDVV(objMargContr)
    If lErro <> SUCESSO Then
    
        objMargContr.dUltPrecoCli = 0
        objMargContr.dUltPrecoPraticado = 0
        objMargContr.dUltPercMargPrat = 0
        objMargContr.dUltPercLucroPrat = 0
        objMargContr.dUltValorMargPrat = 0
    
    Else
    
        lErro = CF("MargContr_CalculaComissoes", colComissoes, objMargContr, dPrecoComissoes, objTelaComissoes, objMargContr.sUltNomeRedCli, objMargContr.sUnidadeMed)
        If lErro <> SUCESSO Then gError 130045
        
        Set objMargContr.colComissoes = colComissoes
        objMargContr.dPrecoComissoes = dPrecoComissoes
    
        'calcula a DVV para o cliente
        lErro = MargContr_CalculaDVV(objMargContr, objMargContr.iColunaCli)
        If lErro <> SUCESSO Then gError 130047
    
        'calcula a DVV para o preco praticado
        lErro = MargContr_CalculaDVV(objMargContr, objMargContr.iColunaPraticado)
        If lErro <> SUCESSO Then gError 130052
    
        lErro = Analise_RecalcularColuna(objMargContr, objMargContr.iColunaCli, objMargContr.iColunaCli, objMargContr.iColunaPraticado)
        If lErro <> SUCESSO Then gError 130053
        
        lErro = Analise_RecalcularColuna(objMargContr, objMargContr.iColunaPraticado, objMargContr.iColunaCli, objMargContr.iColunaPraticado)
        If lErro <> SUCESSO Then gError 130054
    
    End If

    MargContr_Calcula = SUCESSO
     
    Exit Function
    
Erro_MargContr_Calcula:

    MargContr_Calcula = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    Exit Function

End Function

Function PVItens_Grava_Customizado(ByVal objPV As ClassPedidoDeVenda) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(0 To 2) As Long
Dim objItemPV As ClassItemPedido
Dim lNumIntDoc As Long
Dim sTabela As String

On Error GoTo Erro_PVItens_Grava_Customizado

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 141994
    Next

    For Each objItemPV In objPV.colItensPedido
        
        sTabela = "ItensPedidoDeVenda"
        
        'Pesquisa os ítens vinculados ao item
        lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM " & sTabela & " WHERE NumIntDoc = ? ", 0, lNumIntDoc, objItemPV.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 141995

        lErro = Comando_BuscarPrimeiro(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 141996

        If lErro <> AD_SQL_SUCESSO Then

            sTabela = "ItensPedidoDeVendaBaixados"

            'Pesquisa os ítens vinculados ao item
            lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM " & sTabela & " WHERE NumIntDoc = ? ", 0, lNumIntDoc, objItemPV.lNumIntDoc)
            If lErro <> AD_SQL_SUCESSO Then gError 141997
    
            lErro = Comando_BuscarPrimeiro(alComando(0))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 141998

        End If

        If lErro = AD_SQL_SUCESSO Then
        
            lErro = Comando_ExecutarPos(alComando(1), "UPDATE " & sTabela & " SET PrecoBase = ? ", alComando(0), objItemPV.dPrecoBase)
            If lErro <> AD_SQL_SUCESSO Then gError 141999

        End If

    Next

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PVItens_Grava_Customizado = SUCESSO

    Exit Function

Erro_PVItens_Grava_Customizado:

    PVItens_Grava_Customizado = gErr

    Select Case gErr

        Case 141994
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 141995 To 141998
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDIDODEVENDA", gErr)

        Case 141999
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_ITENSPEDIDODEVENDA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 177000)

    End Select

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function RelListaPrecoCalc_Prepara(lNumIntRel As Long, ByVal iFilialEmpresa As Integer, ByVal dtDataAtual As Date, ByVal dtDataAnt As Date, ByVal sProdIni As String, ByVal sProdFim As String, ByVal sCodPrev As String) As Long
'gera registros na tabela RelListaPrecoCalc para a execucao do relatorio de lista de precos

Dim lErro As Long, dQtdeOrcada As Double, dQtdeOrcadaAux As Double, lTransacao As Long
Dim iIndice As Integer, alComando(1 To 5) As Long, sProduto As String
Dim dCustoMPAnt As Double, dCustoEmbAnt As Double, dCustoFreteAnt As Double, dCustoDiretoAnt As Double, dCustoFixoAnt As Double, dDVVSemICMSAnt As Double, dDevDuvAnt As Double
Dim dCustoMPAtual As Double, dCustoEmbAtual As Double, dCustoFreteAtual As Double, dCustoDiretoAtual As Double, dCustoFixoAtual As Double, dDVVSemICMSAtual As Double, dDevDuvAtual As Double
Dim dPrecoCalc1Atual As Double, dPrecoCalc2Atual As Double, dPrecoCalc3Atual As Double, dPrecoCalc2Ant As Double, dPrecoVenda2Ant As Double, dPrecoVenda2Atual As Double

On Error GoTo Erro_RelListaPrecoCalc_Prepara

    'abrir comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 106940
    Next
    
    'Inicia a transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 106941

    'obtem o nº automatico
    lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_PROX_NUMINTREL", lNumIntRel)
    If lErro <> SUCESSO Then gError 106942
    
    sProduto = String(STRING_PRODUTO, 0)
    
    If sProdIni = "" And sProdFim = "" Then
        lErro = Comando_Executar(alComando(1), "SELECT DISTINCT Produto FROM FormacaoPrecoCalc WHERE FilialEmpresa = ? AND Data = ?", sProduto, iFilialEmpresa, dtDataAtual)
    ElseIf sProdIni <> "" And sProdFim = "" Then
        lErro = Comando_Executar(alComando(1), "SELECT DISTINCT Produto FROM FormacaoPrecoCalc WHERE FilialEmpresa = ? AND Data = ? AND Produto >= ?", sProduto, iFilialEmpresa, dtDataAtual, sProdIni)
    ElseIf sProdIni = "" And sProdFim <> "" Then
        lErro = Comando_Executar(alComando(1), "SELECT DISTINCT Produto FROM FormacaoPrecoCalc WHERE FilialEmpresa = ? AND Data = ? AND Produto <= ?", sProduto, iFilialEmpresa, dtDataAtual, sProdFim)
    Else
        lErro = Comando_Executar(alComando(1), "SELECT DISTINCT Produto FROM FormacaoPrecoCalc WHERE FilialEmpresa = ? AND Data = ? AND Produto >= ? AND Produto <= ?", sProduto, iFilialEmpresa, dtDataAtual, sProdIni, sProdFim)
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 106943
    
    lErro = Comando_BuscarProximo(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106944
    
    Do While lErro = AD_SQL_SUCESSO
    
        lErro = CF("Produto_ObtemKGsOrcados", iFilialEmpresa, sProduto, Year(dtDataAtual), sCodPrev, dQtdeOrcada, alComando(5))
        If lErro <> SUCESSO Then gError 106945
        
        'obter valores do calculo anterior
        lErro = RelListaPrecoCalc_Prepara1(iFilialEmpresa, dtDataAnt, 1, sProduto, dCustoMPAnt, dCustoEmbAnt, dCustoFreteAnt, dCustoDiretoAnt, dCustoFixoAnt, dDVVSemICMSAnt, dDevDuvAnt, alComando())
        If lErro <> SUCESSO Then gError 106945
        
        'obter preco de venda anterior da tabela 2
        lErro = CF("Produto_ObtemPrecoData", iFilialEmpresa, 2, sProduto, dtDataAnt, dPrecoVenda2Ant)
        If lErro <> SUCESSO Then gError 106946
        
        'obter valores do calculo atual
        lErro = RelListaPrecoCalc_Prepara1(iFilialEmpresa, dtDataAtual, 1, sProduto, dCustoMPAtual, dCustoEmbAtual, dCustoFreteAtual, dCustoDiretoAtual, dCustoFixoAtual, dDVVSemICMSAtual, dDevDuvAtual, alComando())
        If lErro <> SUCESSO Then gError 106947
        
        'obter preco de venda atual da tabela 2
        lErro = CF("Produto_ObtemPrecoData", iFilialEmpresa, 2, sProduto, dtDataAtual, dPrecoVenda2Atual)
        If lErro <> SUCESSO Then gError 106948
        
        'obter preco calculado atual da tabela 1
        lErro = RelListaPrecoCalc_Prepara2(iFilialEmpresa, dtDataAtual, 1, sProduto, dPrecoCalc1Atual, alComando())
        If lErro <> SUCESSO Then gError 106949
        
        'obter preco calculado atual da tabela 2
        lErro = RelListaPrecoCalc_Prepara2(iFilialEmpresa, dtDataAtual, 2, sProduto, dPrecoCalc2Atual, alComando())
        If lErro <> SUCESSO Then gError 106950
        
        'obter preco calculado atual da tabela 3
        lErro = RelListaPrecoCalc_Prepara2(iFilialEmpresa, dtDataAtual, 3, sProduto, dPrecoCalc3Atual, alComando())
        If lErro <> SUCESSO Then gError 106951
        
        'obter preco calculado anterior da tabela 2
        lErro = RelListaPrecoCalc_Prepara2(iFilialEmpresa, dtDataAnt, 2, sProduto, dPrecoCalc2Ant, alComando())
        If lErro <> SUCESSO Then gError 106950
        
        'grava o registro na tabela
        lErro = Comando_Executar(alComando(2), "INSERT INTO RelListaPrecoCalc (NumIntRel, Produto, QtdeOrcada, CustoMPAnt, CustoEmbAnt, CustoFreteAnt, CustoMPAtual, CustoEmbAtual, CustoFreteAtual, CustoDiretoAtual, CustoFixoAtual, DVVSemICMSAtual, DevDuvAtual, PrecoCalc1Atual, PrecoCalc2Atual, PrecoCalc3Atual, PrecoCalc2Ant, PrecoVenda2Ant, PrecoVenda2Atual) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
            lNumIntRel, sProduto, dQtdeOrcada, dCustoMPAnt, dCustoEmbAnt, dCustoFreteAnt, dCustoMPAtual, dCustoEmbAtual, dCustoFreteAtual, dCustoDiretoAtual, dCustoFixoAtual, dDVVSemICMSAtual, dDevDuvAtual, dPrecoCalc1Atual, dPrecoCalc2Atual, dPrecoCalc3Atual, dPrecoCalc2Ant, dPrecoVenda2Ant, dPrecoVenda2Atual)
        If lErro <> AD_SQL_SUCESSO Then gError 106952
    
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106944
    
    Loop
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 106953
    
   'fechar comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
          
    RelListaPrecoCalc_Prepara = SUCESSO
     
    Exit Function
    
Erro_RelListaPrecoCalc_Prepara:

    RelListaPrecoCalc_Prepara = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    Call Transacao_Rollback
    
   'fechar comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
          
    Exit Function

End Function

Private Function RelListaPrecoCalc_Prepara1(ByVal iFilialEmpresa As Integer, ByVal dtData As Date, ByVal iTabelaPreco As Integer, ByVal sProduto As String, dCustoMP As Double, dCustoEmb As Double, dCustoFrete As Double, dCustoDireto As Double, dCustoFixo As Double, dDVVSemICMS As Double, dDevDuv As Double, alComando() As Long) As Long
'obtem valores correspondentes à estas colunas:
'CustoMP , CustoEmb, CustoFrete, CustoDireto, CustoFixo, DVVSemICMS, DevDuv

Dim lErro As Long, sExpressao As String, dValor As Double

On Error GoTo Erro_RelListaPrecoCalc_Prepara1

    'inicializa valores
    dCustoMP = 0
    dCustoEmb = 0
    dCustoFrete = 0
    dCustoDireto = 0
    dCustoFixo = 0
    dDVVSemICMS = 0
    dDevDuv = 0

    sExpressao = String(STRING_FORMACAOPRECO_EXPRESSAO, 0)
    
    lErro = Comando_Executar(alComando(3), "SELECT Expressao, Valor FROM FormacaoPrecoCalcLin WHERE FilialEmpresa = ? AND Data = ? AND TabelaPreco = ? AND Produto = ?", _
        sExpressao, dValor, iFilialEmpresa, dtData, iTabelaPreco, sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 106937
        
    lErro = Comando_BuscarProximo(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106938
            
    Do While lErro = AD_SQL_SUCESSO
    
        If InStr(1, sExpressao, MNEMONICOFPRECO_CUSTOMP, vbTextCompare) <> 0 Then dCustoMP = dValor
        If InStr(1, sExpressao, MNEMONICOFPRECO_CUSTOEMB, vbTextCompare) <> 0 Then dCustoEmb = dValor
        If InStr(1, sExpressao, "ValorCategoria(""DVV"",3)", vbTextCompare) <> 0 Then dCustoFrete = dValor
        If InStr(1, sExpressao, MNEMONICOFPRECO_CUSTODIRETO, vbTextCompare) <> 0 Then dCustoDireto = dValor
        If InStr(1, sExpressao, MNEMONICOFPRECO_CUSTOFIXO, vbTextCompare) <> 0 Then dCustoFixo = dValor
        If InStr(1, sExpressao, "ValorCategoria(""DVV"",1)", vbTextCompare) <> 0 Then dDVVSemICMS = dValor
        If InStr(1, sExpressao, "DevDuvPadrao", vbTextCompare) <> 0 Then dDevDuv = dValor
        
        lErro = Comando_BuscarProximo(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106939
    
    Loop
        
    RelListaPrecoCalc_Prepara1 = SUCESSO
     
    Exit Function
    
Erro_RelListaPrecoCalc_Prepara1:

    RelListaPrecoCalc_Prepara1 = gErr
     
    Select Case gErr
          
        Case 106937, 106938, 106939
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FORMACAOPRECOCALCLIN", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    Exit Function

End Function

Private Function RelListaPrecoCalc_Prepara2(ByVal iFilialEmpresa As Integer, ByVal dtData As Date, ByVal iTabelaPreco As Integer, ByVal sProduto As String, dPrecoCalc As Double, alComando() As Long) As Long
'obtem o preco calculado pela rotina de calculo de precos

Dim lErro As Long, dValor As Double

On Error GoTo Erro_RelListaPrecoCalc_Prepara2

    lErro = Comando_Executar(alComando(4), "SELECT Valor FROM FormacaoPrecoCalc WHERE FilialEmpresa = ? AND Data = ? AND TabelaPreco = ? AND Produto = ?", _
        dValor, iFilialEmpresa, dtData, iTabelaPreco, sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 106954
        
    lErro = Comando_BuscarProximo(alComando(4))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106955

    If lErro = AD_SQL_SUCESSO Then
        dPrecoCalc = dValor
    Else
        dPrecoCalc = 0
    End If
    
    RelListaPrecoCalc_Prepara2 = SUCESSO
     
    Exit Function
    
Erro_RelListaPrecoCalc_Prepara2:

    RelListaPrecoCalc_Prepara2 = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    Exit Function

End Function

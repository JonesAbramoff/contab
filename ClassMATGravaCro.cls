VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassMATGrava"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function ProdutoKit_Insere(ByVal objKit As ClassKit, ByVal objProdutoKit As ClassProdutoKit, lComando As Long) As Long

Dim lErro As Long
Dim objProdutoKitCro As New ClassProdutoKitCro

On Error GoTo Erro_ProdutoKit_Insere

    If Not (objProdutoKit.objInfoUsu Is Nothing) Then
        Set objProdutoKitCro = objProdutoKit.objInfoUsu
    End If
    
    '################################################
    'Alterado por Wagner 23/11/04
    lErro = Comando_Executar(lComando, "INSERT INTO ProdutoKit (ProdutoRaiz, Nivel, Seq, Versao, Produto, SeqPai, Quantidade, UnidadeMed, Composicao,PosicaoArvore, PercentualPerda, CustoStandard, VersaoKitComp, TipoCarga, GrupoPesagem) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", objKit.sProdutoRaiz, objProdutoKit.iNivel, objProdutoKit.iSeq, objKit.sVersao, objProdutoKit.sProduto, objProdutoKit.iSeqPai, objProdutoKit.dQuantidade, objProdutoKit.sUnidadeMed, objProdutoKit.iComposicao, objProdutoKit.iPosicaoArvore, objProdutoKit.dPercentualPerda, objProdutoKit.dCustoStandard, objProdutoKit.sVersaoKitComp, objProdutoKitCro.iTipoCarga, objProdutoKitCro.sGrupoPesagem)
    If lErro <> AD_SQL_SUCESSO Then gError 21771
    '################################################
        
    ProdutoKit_Insere = SUCESSO
    
    Exit Function
    
Erro_ProdutoKit_Insere:

    ProdutoKit_Insere = gErr

    Select Case gErr

        Case 21771
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PRODUTOKIT", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184295)

    End Select
    
    Exit Function

End Function

'###############################################
'Inserido por Wagner 16/09/05
Function MovimentoEstoque_Alterar_Customizado(ByVal objMovEstoque As ClassMovEstoque) As Long
'altera o responsável
'ATENCAO: Esta Funcao tem que ser chamada dentro de transacao

Dim alComando(1 To 2) As Long
Dim lErro As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim objItemMovEst As ClassItemMovEstoque
Dim objMovEstoqueCro As New ClassMovEstoqueCro

On Error GoTo Erro_MovimentoEstoque_Alterar_Customizado

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 140231
    Next
    
    If Not (objMovEstoque.objInfoUsu Is Nothing) Then
    
        Set objMovEstoqueCro = objMovEstoque.objInfoUsu

        For Each objItemMovEst In objMovEstoque.colItens
    
            'le o codigo do movimento de estoque passado como parametro
            lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntDoc FROM MovimentoEstoque WHERE NumIntDoc = ?", 0, lNumIntDoc, objItemMovEst.lNumIntDoc)
            If lErro <> AD_SQL_SUCESSO Then gError 140232
        
            lErro = Comando_BuscarPrimeiro(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 140233
        
            'se o movimento não está cadastrado
            If lErro = AD_SQL_SEM_DADOS Then gError 140234
        
            lErro = Comando_ExecutarPos(alComando(2), "UPDATE MovimentoEstoque SET Responsavel = ?", alComando(1), objMovEstoqueCro.sResponsavel)
            If lErro <> AD_SQL_SUCESSO Then gError 140235
        
       Next
       
    End If
   
   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    MovimentoEstoque_Alterar_Customizado = SUCESSO
    
    Exit Function

Erro_MovimentoEstoque_Alterar_Customizado:

    MovimentoEstoque_Alterar_Customizado = gErr

    Select Case gErr

        Case 140231
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 140232, 140233
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE", gErr)

        Case 140234
            Call Rotina_Erro(vbOKOnly, "ERRO_MOVESTOQUE_NAO_CADASTRADO1", gErr, objItemMovEst.lNumIntDoc)

        Case 140235
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_MOVIMENTOESTOQUE", gErr, objItemMovEst.iFilialEmpresa, objItemMovEst.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function
'###############################################

'###############################################
'Inserido por Wagner 23/02/2007
Function RastreamentoLote_Grava_Customizado(ByVal objRastroLote As ClassRastreamentoLote) As Long
'Grava o RastreamentoLote

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(0 To 1) As Long
Dim iStatus As Integer
Dim objRastroLoteCro As ClassRastreamentoLoteCro

On Error GoTo Erro_RastreamentoLote_Grava_Customizado

    'Abrir comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 185923
    Next
        
    If Not (objRastroLote.objInfoUsu Is Nothing) Then
    
        Set objRastroLoteCro = objRastroLote.objInfoUsu
   
        'Verifica se o RastreamentoLote já está cadastrado
        lErro = Comando_ExecutarPos(alComando(0), "SELECT Status FROM RastreamentoLote WHERE Produto = ? AND Lote = ? AND FilialOP = ?", 0, iStatus, objRastroLote.sProduto, objRastroLote.sCodigo, objRastroLote.iFilialOP)
        If lErro <> AD_SQL_SUCESSO Then gError 185924
    
        lErro = Comando_BuscarPrimeiro(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185925
    
        'Se o Lote está baixado, erro
        If iStatus = RATREAMENTOLOTE_STATUS_BAIXADO Then gError 185926
    
        'Se encontrou o RastreamentoLote, ALTERAÇÃO
        If lErro = AD_SQL_SUCESSO Then

            'Altera dados do RastreamentoLote
            lErro = Comando_ExecutarPos(alComando(1), "UPDATE RastreamentoLote Set Fornecedor = ?, FilialForn = ?, Localizacao = ?", alComando(0), objRastroLoteCro.lFornecedor, objRastroLoteCro.iFilialForn, objRastroLoteCro.sLocalizacao)
            If lErro <> AD_SQL_SUCESSO Then gError 185927
    
        End If
        
    End If

    'Fechar comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    RastreamentoLote_Grava_Customizado = SUCESSO

    Exit Function

Erro_RastreamentoLote_Grava_Customizado:

    RastreamentoLote_Grava_Customizado = gErr

    Select Case gErr

        Case 185921

        Case 185923
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 185924, 185925
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RASTREAMENTOLOTE", gErr)

        Case 185926
            Call Rotina_Erro(vbOKOnly, "ERRO_LOTE_BAIXADO", gErr, objRastroLote.sCodigo)

        Case 185927
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_RASTREAMENTOLOTE", gErr, objRastroLote.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function
'######################################################

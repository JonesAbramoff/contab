VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTRBGrava"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function ICMSExcecao_Grava(ByVal objICMSExcecoes As ClassICMSExcecao) As Long
Dim lErro As Long
Dim iTipo As Integer
Dim lTransacao As Long
Dim sDescricao1 As String
Dim sDescricao2 As String
Dim sDescricao3 As String
Dim alComando(6) As Long
Dim iIndice As Integer

On Error GoTo Erro_ICMSExcecao_Grava

    'Abre transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 21543

    'Abre Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 21544
    Next

    '###############################################
    'Alterado por Wagner 29/09/05
    If objICMSExcecoes.iTipoCliForn = ICMSEXCECOES_TIPOCLIFORN_CLIENTE Then
    
        If Len(objICMSExcecoes.sCategoriaCliente) > 0 Then
    
            sDescricao1 = String(STRING_CATEGORIACLIENTEITEM_DESCRICAO, 0)
    
            'Leitura da tabela de itens das categorias de clientes
            lErro = Comando_ExecutarPos(alComando(3), "SELECT Descricao FROM CategoriaClienteItem WHERE Categoria = ? AND Item =? ", 0, _
            sDescricao1, objICMSExcecoes.sCategoriaCliente, objICMSExcecoes.sCategoriaClienteItem)
            If lErro <> AD_SQL_SUCESSO Then Error 21548
    
            lErro = Comando_BuscarPrimeiro(alComando(3))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 21549
    
            If lErro = AD_SQL_SEM_DADOS Then Error 21550
    
            'Lock das tabelas
            lErro = Comando_LockShared(alComando(3))
            If lErro <> AD_SQL_SUCESSO Then Error 21558
    
        End If
        
    Else
    
        If Len(objICMSExcecoes.sCategoriaFornecedor) > 0 Then
    
            sDescricao2 = String(STRING_CATEGORIAFORNECEDORITEM_DESCRICAO, 0)
    
            'Leitura da tabela de itens das categorias de clientes
            lErro = Comando_ExecutarPos(alComando(3), "SELECT Descricao FROM CategoriaFornecedorItem WHERE Categoria = ? AND Item =? ", 0, _
            sDescricao2, objICMSExcecoes.sCategoriaFornecedor, objICMSExcecoes.sCategoriaFornecedorItem)
            If lErro <> AD_SQL_SUCESSO Then Error 21548
    
            lErro = Comando_BuscarPrimeiro(alComando(3))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 21549
    
            If lErro = AD_SQL_SEM_DADOS Then Error 21550
    
            'Lock das tabelas
            lErro = Comando_LockShared(alComando(3))
            If lErro <> AD_SQL_SUCESSO Then Error 21558
    
        End If
    
    End If
    '###############################################

    If Len(objICMSExcecoes.sCategoriaProduto) > 0 Then

        sDescricao3 = String(STRING_CATEGORIAPRODUTOITEM_DESCRICAO, 0)

        'Leitura da tabela de itens de categorias de produtos
        lErro = Comando_ExecutarPos(alComando(5), "SELECT Descricao FROM CategoriaProdutoItem WHERE Categoria = ? AND Item =? ", 0, _
        sDescricao3, objICMSExcecoes.sCategoriaProduto, objICMSExcecoes.sCategoriaProdutoItem)
        If lErro <> AD_SQL_SUCESSO Then Error 21554

        lErro = Comando_BuscarPrimeiro(alComando(5))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 21555

        If lErro = AD_SQL_SEM_DADOS Then Error 21556

        lErro = Comando_LockShared(alComando(5))
        If lErro <> AD_SQL_SUCESSO Then Error 21560

    End If

    'Inserção ou atualização dos dados
    lErro = ICMSExcecao_Grava1(objICMSExcecoes, alComando())
    If lErro <> SUCESSO Then Error 21570

    'Confirma transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 21565

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    ICMSExcecao_Grava = SUCESSO

    Exit Function

Erro_ICMSExcecao_Grava:

    ICMSExcecao_Grava = Err

    Select Case Err

        Case 21543
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 21572
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 21548, 21549
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CATEGORIACLIENTEITEM2", Err, objICMSExcecoes.sCategoriaCliente, objICMSExcecoes.sCategoriaClienteItem)

        Case 21550
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIACLIENTEITEM_INEXISTENTE", Err, objICMSExcecoes.sCategoriaCliente, objICMSExcecoes.sCategoriaClienteItem)

        Case 21554, 21555
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CATEGORIAPRODUTOITEM2", Err, objICMSExcecoes.sCategoriaProduto, objICMSExcecoes.sCategoriaProdutoItem)

        Case 21556
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTOITEM_INEXISTENTE", Err, objICMSExcecoes.sCategoriaProduto, objICMSExcecoes.sCategoriaProdutoItem)

        Case 21558
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_CATEGORIACLIENTEITEM", Err)

        Case 21560
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_CATEGORIAPRODUTOITEM", Err)

        Case 21565
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case 21570

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153635)

    End Select

    Call Transacao_Rollback

    For iIndice = LBound(alComando) To UBound(alComando)

        Call Comando_Fechar(alComando(iIndice))

    Next

    Exit Function

End Function

Private Function ICMSExcecao_Grava1(ByVal objICMSExcecoes As ClassICMSExcecao, alComando() As Long) As Long
'grava as alterações/inclusões realizadas pelo usuário nos dados de ICMSExcecoes para um registro.

Dim lErro As Long
Dim iTipo As Integer
Dim sDescricao As String
Dim iIndice As Integer

On Error GoTo Erro_ICMSExcecao_Grava1

    '##################################################################
    'Alterado por Wagner 29/09/05
    'Leitura da tabela de exceções de ICMS
'    lErro = Comando_ExecutarPos(alComando(1), "SELECT Tipo FROM ICMSExcecoes WHERE EstadoDestino = ? AND CategoriaProduto = ? AND CategoriaProdutoItem =? AND CategoriaCliente = ? AND CategoriaClienteItem =? AND EstadoOrigem = ? AND CategoriaFornecedor = ? AND CategoriaFornecedorItem =?", 0, _
'    iTipo, objICMSExcecoes.sEstadoDestino, objICMSExcecoes.sCategoriaProduto, objICMSExcecoes.sCategoriaProdutoItem, objICMSExcecoes.sCategoriaCliente, objICMSExcecoes.sCategoriaClienteItem, objICMSExcecoes.sEstadoOrigem, objICMSExcecoes.sCategoriaFornecedor, objICMSExcecoes.sCategoriaFornecedorItem)
'    If lErro <> AD_SQL_SUCESSO Then Error 21561
    '##################################################################
    
    lErro = Comando_Executar(alComando(1), "SELECT Tipo FROM ICMSExcecoes WHERE GrupoOrigemMercadoria = ? AND EstadoDestino = ? AND CategoriaProduto = ? AND CategoriaProdutoItem =? AND CategoriaCliente = ? AND CategoriaClienteItem =? AND EstadoOrigem = ? AND CategoriaFornecedor = ? AND CategoriaFornecedorItem =? AND TipoAplicacao = ?", _
        iTipo, objICMSExcecoes.iGrupoOrigemMercadoria, objICMSExcecoes.sEstadoDestino, objICMSExcecoes.sCategoriaProduto, objICMSExcecoes.sCategoriaProdutoItem, objICMSExcecoes.sCategoriaCliente, objICMSExcecoes.sCategoriaClienteItem, objICMSExcecoes.sEstadoOrigem, objICMSExcecoes.sCategoriaFornecedor, objICMSExcecoes.sCategoriaFornecedorItem, objICMSExcecoes.iTipoAplicacao)
    If lErro <> AD_SQL_SUCESSO Then Error 21561

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 21562

    If lErro = AD_SQL_SEM_DADOS Then

        '##################################################################
        'Alterado por Wagner 29/09/05
        'Se não existe, realiza inserção
        lErro = Comando_Executar(alComando(2), "INSERT INTO ICMSExcecoes (GrupoOrigemMercadoria, EstadoDestino ,CategoriaProduto ,CategoriaProdutoItem ,CategoriaCliente ,CategoriaClienteItem, EstadoOrigem, Tipo, PercRedBaseCalculo, Aliquota, PercMargemLucro, Fundamentacao, Prioridade,CategoriaFornecedor ,CategoriaFornecedorItem, TipoCliForn, UsaPauta, ValorPauta, PercRedBaseCalculoSubst, TipoSimples, ICMSPercFCP, TipoAplicacao,ICMSSTBaseDupla,ICMSSTBaseDuplaIni,cBenef,ICMSMotivo) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
            objICMSExcecoes.iGrupoOrigemMercadoria, objICMSExcecoes.sEstadoDestino, objICMSExcecoes.sCategoriaProduto, objICMSExcecoes.sCategoriaProdutoItem, objICMSExcecoes.sCategoriaCliente, objICMSExcecoes.sCategoriaClienteItem, objICMSExcecoes.sEstadoOrigem, objICMSExcecoes.iTipo, objICMSExcecoes.dPercRedBaseCalculo, objICMSExcecoes.dAliquota, objICMSExcecoes.dPercMargemLucro, objICMSExcecoes.sFundamentacao, objICMSExcecoes.iPrioridade, objICMSExcecoes.sCategoriaFornecedor, objICMSExcecoes.sCategoriaFornecedorItem, objICMSExcecoes.iTipoCliForn, objICMSExcecoes.iUsaPauta, objICMSExcecoes.dValorPauta, objICMSExcecoes.dPercRedBaseCalculoSubst, objICMSExcecoes.iTipoSimples, objICMSExcecoes.dICMSPercFCP, objICMSExcecoes.iTipoAplicacao, objICMSExcecoes.iICMSSTBaseDupla, objICMSExcecoes.dtICMSSTBaseDuplaIni, objICMSExcecoes.scBenef, objICMSExcecoes.iICMSMotivo)
        If lErro <> AD_SQL_SUCESSO Then Error 21563
        '##################################################################

    Else
        
        'Se existe, atualiza
'        lErro = Comando_ExecutarPos(alComando(2), "UPDATE ICMSExcecoes SET Tipo = ? , PercRedBaseCalculo = ? , Aliquota = ? , PercMargemLucro = ?, Fundamentacao = ? , Prioridade = ? , TipoCliForn = ?, UsaPauta = ?, ValorPauta = ?, PercRedBaseCalculoSubst = ?", alComando(1), _
'            objICMSExcecoes.iTipo, objICMSExcecoes.dPercRedBaseCalculo, objICMSExcecoes.dAliquota, objICMSExcecoes.dPercMargemLucro, objICMSExcecoes.sFundamentacao, objICMSExcecoes.iPrioridade, objICMSExcecoes.iTipoCliForn, objICMSExcecoes.iUsaPauta, objICMSExcecoes.dValorPauta, objICMSExcecoes.dPercRedBaseCalculoSubst)
'        If lErro <> AD_SQL_SUCESSO Then Error 21564

        lErro = Comando_Executar(alComando(2), "UPDATE ICMSExcecoes SET Tipo = ? , PercRedBaseCalculo = ? , Aliquota = ? , PercMargemLucro = ?, Fundamentacao = ? , Prioridade = ? , TipoCliForn = ?, UsaPauta = ?, ValorPauta = ?, PercRedBaseCalculoSubst = ?, TipoSimples = ?, ICMSPercFCP = ?,ICMSSTBaseDupla = ?,ICMSSTBaseDuplaIni = ?, cBenef = ?, ICMSMotivo = ? WHERE GrupoOrigemMercadoria = ? AND EstadoDestino = ? AND CategoriaProduto = ? AND CategoriaProdutoItem =? AND CategoriaCliente = ? AND CategoriaClienteItem =? AND EstadoOrigem = ? AND CategoriaFornecedor = ? AND CategoriaFornecedorItem =?", _
            objICMSExcecoes.iTipo, objICMSExcecoes.dPercRedBaseCalculo, objICMSExcecoes.dAliquota, objICMSExcecoes.dPercMargemLucro, objICMSExcecoes.sFundamentacao, objICMSExcecoes.iPrioridade, objICMSExcecoes.iTipoCliForn, objICMSExcecoes.iUsaPauta, objICMSExcecoes.dValorPauta, objICMSExcecoes.dPercRedBaseCalculoSubst, objICMSExcecoes.iTipoSimples, objICMSExcecoes.dICMSPercFCP, objICMSExcecoes.iICMSSTBaseDupla, objICMSExcecoes.dtICMSSTBaseDuplaIni, objICMSExcecoes.scBenef, objICMSExcecoes.iICMSMotivo, objICMSExcecoes.iGrupoOrigemMercadoria, objICMSExcecoes.sEstadoDestino, objICMSExcecoes.sCategoriaProduto, objICMSExcecoes.sCategoriaProdutoItem, objICMSExcecoes.sCategoriaCliente, objICMSExcecoes.sCategoriaClienteItem, objICMSExcecoes.sEstadoOrigem, objICMSExcecoes.sCategoriaFornecedor, objICMSExcecoes.sCategoriaFornecedorItem)
        If lErro <> AD_SQL_SUCESSO Then Error 21564

    End If

    ICMSExcecao_Grava1 = SUCESSO

    Exit Function

Erro_ICMSExcecao_Grava1:

    ICMSExcecao_Grava1 = Err

    Select Case Err

        Case 21561, 21562
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ICMS_EXCECOES", Err)

        Case 21563
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ICMSEXCECOES", Err, objICMSExcecoes.sEstadoDestino, objICMSExcecoes.sCategoriaCliente, objICMSExcecoes.sCategoriaClienteItem, objICMSExcecoes.sCategoriaProduto, objICMSExcecoes.sCategoriaProdutoItem)

        Case 21564
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_ICMSEXCECOES", Err, objICMSExcecoes.sEstadoDestino, objICMSExcecoes.sCategoriaCliente, objICMSExcecoes.sCategoriaClienteItem, objICMSExcecoes.sCategoriaProduto, objICMSExcecoes.sCategoriaProdutoItem)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153636)

    End Select

    Exit Function

End Function

Function ICMSExcecao_Exclui(ByVal objICMSExcecoes As ClassICMSExcecao) As Long
'grava as exclusões realizadas pelo usuário nos dados de ICMSExcecoes para um registro.

Dim lErro As Long
Dim iTipo As Integer
Dim lTransacao As Long
Dim lComando As Long
Dim lComando1 As Long

On Error GoTo Erro_ICMSExcecao_Exclui

    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 21490

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 21491

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 21530

'    lErro = Comando_ExecutarPos(lComando, "SELECT Tipo FROM ICMSExcecoes WHERE EstadoDestino = ?  AND EstadoOrigem = ? AND CategoriaProduto = ? AND CategoriaProdutoItem =? AND CategoriaCliente = ? AND CategoriaClienteItem =? AND CategoriaFornecedor = ? AND CategoriaFornecedorItem =? AND TipoCliForn = ?", 0, _
'        iTipo, objICMSExcecoes.sEstadoDestino, objICMSExcecoes.sEstadoOrigem, objICMSExcecoes.sCategoriaProduto, objICMSExcecoes.sCategoriaProdutoItem, objICMSExcecoes.sCategoriaCliente, objICMSExcecoes.sCategoriaClienteItem, objICMSExcecoes.sCategoriaFornecedor, objICMSExcecoes.sCategoriaFornecedorItem, objICMSExcecoes.iTipoCliForn)
'    If lErro <> AD_SQL_SUCESSO Then Error 21492

    lErro = Comando_Executar(lComando, "SELECT Tipo FROM ICMSExcecoes WHERE GrupoOrigemMercadoria = ? AND EstadoDestino = ?  AND EstadoOrigem = ? AND CategoriaProduto = ? AND CategoriaProdutoItem =? AND CategoriaCliente = ? AND CategoriaClienteItem =? AND CategoriaFornecedor = ? AND CategoriaFornecedorItem =? AND TipoCliForn = ? AND TipoAplicacao = ?", _
        iTipo, objICMSExcecoes.iGrupoOrigemMercadoria, objICMSExcecoes.sEstadoDestino, objICMSExcecoes.sEstadoOrigem, objICMSExcecoes.sCategoriaProduto, objICMSExcecoes.sCategoriaProdutoItem, objICMSExcecoes.sCategoriaCliente, objICMSExcecoes.sCategoriaClienteItem, objICMSExcecoes.sCategoriaFornecedor, objICMSExcecoes.sCategoriaFornecedorItem, objICMSExcecoes.iTipoCliForn, objICMSExcecoes.iTipoAplicacao)
    If lErro <> AD_SQL_SUCESSO Then Error 21492
    
    'Lê o primeiro registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 21493

    If lErro = AD_SQL_SEM_DADOS Then Error 21494

'    lErro = Comando_ExecutarPos(lComando1, "DELETE FROM ICMSExcecoes", lComando)
'    If lErro <> AD_SQL_SUCESSO Then Error 21495
    
    lErro = Comando_Executar(lComando1, "DELETE FROM ICMSExcecoes WHERE GrupoOrigemMercadoria = ? AND EstadoDestino = ?  AND EstadoOrigem = ? AND CategoriaProduto = ? AND CategoriaProdutoItem =? AND CategoriaCliente = ? AND CategoriaClienteItem =? AND CategoriaFornecedor = ? AND CategoriaFornecedorItem =? AND TipoCliForn = ? AND TipoAplicacao = ?", _
        objICMSExcecoes.iGrupoOrigemMercadoria, objICMSExcecoes.sEstadoDestino, objICMSExcecoes.sEstadoOrigem, objICMSExcecoes.sCategoriaProduto, objICMSExcecoes.sCategoriaProdutoItem, objICMSExcecoes.sCategoriaCliente, objICMSExcecoes.sCategoriaClienteItem, objICMSExcecoes.sCategoriaFornecedor, objICMSExcecoes.sCategoriaFornecedorItem, objICMSExcecoes.iTipoCliForn, objICMSExcecoes.iTipoAplicacao)
    If lErro <> AD_SQL_SUCESSO Then Error 21495

    'Confirma transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 21496

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    ICMSExcecao_Exclui = SUCESSO

    Exit Function

Erro_ICMSExcecao_Exclui:

    ICMSExcecao_Exclui = Err

    Select Case Err

        Case 21490, 21530
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 21491
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 21492, 21493
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ICMS_EXCECOES", Err)

        Case 21494
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ICMSEXCECOES_INEXISTENTE", Err, objICMSExcecoes.sEstadoDestino, objICMSExcecoes.sCategoriaCliente, objICMSExcecoes.sCategoriaClienteItem, objICMSExcecoes.sCategoriaProduto, objICMSExcecoes.sCategoriaProdutoItem)

        Case 21495
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ICMSEXCECOES", Err, objICMSExcecoes.sEstadoDestino, objICMSExcecoes.sCategoriaCliente, objICMSExcecoes.sCategoriaClienteItem, objICMSExcecoes.sCategoriaProduto, objICMSExcecoes.sCategoriaProdutoItem)

        Case 21496
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153637)

    End Select

    Call Transacao_Rollback

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Function TipoTributacao_Grava(ByVal objTipoTributacao As ClassTipoDeTributacaoMovto) As Long
'Inclui ou altera um tipo de tributação

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(2) As Long
Dim lTransacao As Long
Dim iTipo As Integer, iEntrada As Integer

On Error GoTo Erro_TipoTributacao_Grava

    'Abre os comandos
    For iIndice = 0 To 1
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 33302
    Next

    'Abre a transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 33303

    'Pesquisar o Tipo de Tributação pelo código do tipo
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Tipo, Entrada FROM TiposDeTributacaoMovto WHERE Tipo = ?", 0, iTipo, iEntrada, objTipoTributacao.iTipo)
    If lErro <> AD_SQL_SUCESSO Then Error 33304

    'Tenta ler
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 33305

    'Encontrou o Tipo de Tributação
    If lErro = AD_SQL_SUCESSO Then 'MODIFICAÇÃO

        '"Lock" do Tipo de Tributação
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then Error 33306

        If iEntrada <> objTipoTributacao.iEntrada Then Error 59377
        
        'Atualiza a tabela de TiposDeTributacaoMovto
        lErro = Comando_ExecutarPos(alComando(1), "UPDATE TiposDeTributacaoMovto SET Descricao = ?, ICMSIncide = ?, ICMSTipo = ?, ICMSBaseComIPI = ?, ICMSCredita = ?, IPIIncide = ?, IPITipo = ?, IPIFrete = ?, IPIDestaca = ?, IPICredita = ?, ISSIncide = ?, IRIncide = ?, IRAliquota = ?, INSSAliquota = ?, INSSRetencaoMinima = ?, INSSIncide = ?, PISCredita = ?, PISRetencao = ?, ISSRetencao = ?, COFINSCredita = ?, COFINSRetencao = ?, CSLLRetencao = ? , ISSTipo =? , PISTipo = ?, COFINSTipo = ?, ICMSSimplesTipo = ?, RegimeTributario = ?, NatBCCred = ?, ISSIndExigibilidade = ?, IPICodEnq = ?", _
            alComando(0), objTipoTributacao.sDescricao, objTipoTributacao.iICMSIncide, objTipoTributacao.iICMSTipo, objTipoTributacao.iICMSBaseComIPI, objTipoTributacao.iICMSCredita, objTipoTributacao.iIPIIncide, objTipoTributacao.iIPITipo, objTipoTributacao.iIPIFrete, objTipoTributacao.iIPIDestaca, objTipoTributacao.iIPICredita, objTipoTributacao.iISSIncide, objTipoTributacao.iIRIncide, objTipoTributacao.dIRAliquota, objTipoTributacao.dINSSAliquota, objTipoTributacao.dINSSRetencaoMinima, objTipoTributacao.iINSSIncide, _
            objTipoTributacao.iPISCredita, objTipoTributacao.iPISRetencao, objTipoTributacao.iISSRetencao, objTipoTributacao.iCOFINSCredita, objTipoTributacao.iCOFINSRetencao, objTipoTributacao.iCSLLRetencao, objTipoTributacao.iISSTipo, objTipoTributacao.iPISTipo, objTipoTributacao.iCOFINSTipo, objTipoTributacao.iICMSSimplesTipo, objTipoTributacao.iRegimeTributario, objTipoTributacao.sNatBCCred, objTipoTributacao.iISSIndExigibilidade, objTipoTributacao.sIPICodEnq)
        If lErro <> AD_SQL_SUCESSO Then Error 33307

    Else 'INCLUSAO

        'Insere um novo registro na tabela de TiposDeTributacaoMovto
        lErro = Comando_Executar(alComando(0), "INSERT INTO TiposDeTributacaoMovto(Tipo, Descricao, Entrada, ICMSIncide, ICMSTipo, ICMSBaseComIPI, ICMSCredita, IPIIncide, IPITipo, IPIFrete, IPIDestaca, IPICredita, ISSIncide, IRIncide, IRAliquota, INSSAliquota, INSSRetencaoMinima, INSSIncide, PISCredita, PISRetencao, ISSRetencao, COFINSCredita, COFINSRetencao, CSLLRetencao, ISSTipo, PISTipo, COFINSTipo, ICMSSimplesTipo, RegimeTributario, NatBCCred, ISSIndExigibilidade,IPICodEnq) " & _
            "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", objTipoTributacao.iTipo, objTipoTributacao.sDescricao, objTipoTributacao.iEntrada, objTipoTributacao.iICMSIncide, objTipoTributacao.iICMSTipo, objTipoTributacao.iICMSBaseComIPI, objTipoTributacao.iICMSCredita, objTipoTributacao.iIPIIncide, objTipoTributacao.iIPITipo, objTipoTributacao.iIPIFrete, objTipoTributacao.iIPIDestaca, objTipoTributacao.iIPICredita, objTipoTributacao.iISSIncide, objTipoTributacao.iIRIncide, objTipoTributacao.dIRAliquota, objTipoTributacao.dINSSAliquota, objTipoTributacao.dINSSRetencaoMinima, objTipoTributacao.iINSSIncide, _
            objTipoTributacao.iPISCredita, objTipoTributacao.iPISRetencao, objTipoTributacao.iISSRetencao, objTipoTributacao.iCOFINSCredita, objTipoTributacao.iCOFINSRetencao, objTipoTributacao.iCSLLRetencao, objTipoTributacao.iISSTipo, objTipoTributacao.iPISTipo, objTipoTributacao.iCOFINSTipo, objTipoTributacao.iICMSSimplesTipo, objTipoTributacao.iRegimeTributario, objTipoTributacao.sNatBCCred, objTipoTributacao.iISSIndExigibilidade, objTipoTributacao.sIPICodEnq)
        If lErro <> AD_SQL_SUCESSO Then Error 33308

    End If

    'Faz o commit da transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 33376

   'Fecha os comandos
    For iIndice = 0 To 1
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    TipoTributacao_Grava = SUCESSO

    Exit Function

Erro_TipoTributacao_Grava:

    TipoTributacao_Grava = Err

    Select Case Err

        Case 33302
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 33303
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 33304, 33305
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPO_TRIBUTACAO", Err, objTipoTributacao.iTipo)

        Case 33306
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_TIPOSDETRIBUTACAOMOVTO", Err, objTipoTributacao.iTipo)

        Case 33307
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_TIPOSDETRIBUTACAOMOVTO", Err, objTipoTributacao.iTipo)

        Case 33308
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TIPOSDETRIBUTACAOMOVTO", Err, objTipoTributacao.iTipo)

        Case 33376
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case 59377
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOTRIBMOV_ALT_ENTSAI", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153638)

    End Select

    Call Transacao_Rollback

    'Fecha os comandos
    For iIndice = 0 To 1
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function TipoTributacao_Exclui(ByVal objTipoTributacao As ClassTipoDeTributacaoMovto) As Long
'Exclui um tipo de tributação paassado como parâmetro
'objTipoDeTributacao.iTipo tem que estar preenchido

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(3) As Long
Dim lTransacao As Long
Dim iTipo As Integer

On Error GoTo Erro_TipoTributacao_Exclui

    'Abertura comandos
    For iIndice = 0 To 2

        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 33313

    Next

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 33314

    'Pesquisa o Tipo de Tributação
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Tipo FROM TiposDeTributacaoMovto WHERE Tipo = ?", 0, iTipo, objTipoTributacao.iTipo)
    If lErro <> AD_SQL_SUCESSO Then Error 33315

    'Tenta ler
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 33316

    'Se não encontrou o Tipo de Tributação ==> erro
    If lErro = AD_SQL_SEM_DADOS Then Error 33317

    'Lock da tabela de TiposDeTributacaoMovto
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> AD_SQL_SUCESSO Then Error 33318

    'verifica se pode excluir
    lErro = TipoTributacao_PodeExcluir(objTipoTributacao, alComando(1))
    If lErro <> SUCESSO Then Error 41510

    'Exclui o Tipo de Tributação
    lErro = Comando_ExecutarPos(alComando(2), "DELETE FROM TiposDeTributacaoMovto", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then Error 33323

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 33324

    'Fechamento comandos
    For iIndice = 0 To 2
        Call Comando_Fechar(alComando(iIndice))
    Next

    TipoTributacao_Exclui = SUCESSO

    Exit Function

Erro_TipoTributacao_Exclui:

    TipoTributacao_Exclui = Err

    Select Case Err

        Case 33313
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 33314
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 33315, 33316
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPO_TRIBUTACAO", Err, objTipoTributacao.iTipo)

        Case 33317
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_TRIBUTACAO_NAO_CADASTRADO", Err, objTipoTributacao.iTipo)

        Case 33318
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_TIPOSDETRIBUTACAOMOVTO", Err, objTipoTributacao.iTipo)

        Case 33323
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TIPOSDETRIBUTACAOMOVTO", Err, objTipoTributacao.iTipo)

        Case 33324
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case 41510
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153639)

    End Select

    'Fechamento transação
    Call Transacao_Rollback

    'Fechamento comandos
    For iIndice = 0 To 2
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function PadraoTribSaida_Grava(ByVal objPadraoTribSaida As ClassPadraoTribSaida) As Long
'Inclui ou Altera um Padrão de Tributação Saída

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(6) As Long
Dim lTransacao As Long
Dim iTipoTributacaoPadrao As Integer

On Error GoTo Erro_PadraoTribSaida_Grava

    'Abre os comandos
    For iIndice = 0 To 5
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 33348
    Next

    'Abre a transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 33349

    'Pesquisar o Padrão de Tributação Saída
    lErro = Comando_ExecutarPos(alComando(0), "SELECT TipoTributacaoPadrao FROM PadroesTribSaida WHERE NaturezaOperacao = ? AND SiglaMovto = ? AND CategoriaFilialCliente = ? AND ItemCategoria = ?", 0, iTipoTributacaoPadrao, objPadraoTribSaida.sNaturezaOperacao, objPadraoTribSaida.sSiglaMovto, objPadraoTribSaida.sCategoriaFilialCliente, objPadraoTribSaida.sItemCategoria)
    If lErro <> AD_SQL_SUCESSO Then Error 33350

    'Tenta ler
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 33351

    'Encontrou o Padrão de Tributação
    If lErro = AD_SQL_SUCESSO Then 'MODIFICAÇÃO

        '"Lock" do Padrão de Tributação Saída
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then Error 33352

        If iTipoTributacaoPadrao <> objPadraoTribSaida.iTipoTributacaoPadrao Then

            'Atualiza a tabela de Padrão de Tributação Saída
            lErro = Comando_ExecutarPos(alComando(1), "UPDATE PadroesTribSaida SET TipoTributacaoPadrao = ?", alComando(0), objPadraoTribSaida.iTipoTributacaoPadrao)
            If lErro <> AD_SQL_SUCESSO Then Error 33353

        End If

    Else 'CRIAÇÃO

        'Verifica se a CategoriaCliente foi preenchida
        If objPadraoTribSaida.sCategoriaFilialCliente <> "" Then

            'Verifica se a Item da CategoriaCliente foi preenchido
            If objPadraoTribSaida.sItemCategoria <> "" Then
            
                'Pequisa a o Item da Categoria
                lErro = CF("CategoriaClienteItem_Lock", objPadraoTribSaida.sCategoriaFilialCliente, objPadraoTribSaida.sItemCategoria, alComando(2))
                If lErro <> SUCESSO And lErro <> 40555 Then Error 40559
                
                'Se não encontrou ==> erro
                If lErro = 40555 Then Error 40560

            End If

        End If

        'Verifica se a Natureza da Operação foi preenchida
        If objPadraoTribSaida.sNaturezaOperacao <> "" Then
        
            'Pequisa a Natureza da Operação
            lErro = CF("NaturezaOperacao_Lock", objPadraoTribSaida.sNaturezaOperacao, alComando(3))
            If lErro <> SUCESSO And lErro <> 40551 Then Error 40557
            
            'Se não encontrou ==> erro
            If lErro = 40551 Then Error 40558
            
        End If
        'Verifica se o Tipo Documento foi preenchido
        If objPadraoTribSaida.sSiglaMovto <> "" Then
        
            'Pesquisa o Tipo de Documento e faz "Lock" do Tipo do Documento
            lErro = CF("TipoDocInfo_Lock_Sigla", objPadraoTribSaida.sSiglaMovto, alComando(4))
            If lErro <> SUCESSO And lErro <> 40569 Then Error 40571
        
            'Se não encontrou ----> Erro
            If lErro = 40569 Then Error 40572
                
        End If
        
        'Pesquisa o tipo de Tributação
        lErro = CF("TipoTributacao_Lock", objPadraoTribSaida.iTipoTributacaoPadrao, alComando(5))
        If lErro <> SUCESSO And lErro <> 40563 Then Error 40565
        
        'Se não encontrou -----> Erro
        If lErro = 40563 Then Error 40566
        
        'Insere um novo registro na tabela de PadroesTribSaida
        lErro = Comando_Executar(alComando(0), "INSERT INTO PadroesTribSaida(NaturezaOperacao, SiglaMovto, CategoriaFilialCliente, ItemCategoria, TipoTributacaoPadrao) VALUES (?,?,?,?,?)", objPadraoTribSaida.sNaturezaOperacao, objPadraoTribSaida.sSiglaMovto, objPadraoTribSaida.sCategoriaFilialCliente, objPadraoTribSaida.sItemCategoria, objPadraoTribSaida.iTipoTributacaoPadrao)
        If lErro <> AD_SQL_SUCESSO Then Error 33374

    End If

    'Faz o commit da transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 33375

    'Fecha os comandos
    For iIndice = 0 To 5
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    PadraoTribSaida_Grava = SUCESSO

    Exit Function

Erro_PadraoTribSaida_Grava:

    PadraoTribSaida_Grava = Err

    Select Case Err

        Case 33348
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 33349
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 33350, 33351
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PADROESTRIBUTACAO1", Err, objPadraoTribSaida.sNaturezaOperacao, objPadraoTribSaida.sSiglaMovto, objPadraoTribSaida.sCategoriaFilialCliente, objPadraoTribSaida.sItemCategoria)

        Case 33352
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PADROESTRIBUTACAO", Err, objPadraoTribSaida.sNaturezaOperacao, objPadraoTribSaida.sSiglaMovto, objPadraoTribSaida.sCategoriaFilialCliente, objPadraoTribSaida.sItemCategoria)

        Case 33353
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_PADROESTRIBUTACAO", Err, objPadraoTribSaida.sNaturezaOperacao, objPadraoTribSaida.sSiglaMovto, objPadraoTribSaida.sCategoriaFilialCliente, objPadraoTribSaida.sItemCategoria)

        Case 33374
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PADROESTRIBUTACAO", Err, objPadraoTribSaida.sNaturezaOperacao, objPadraoTribSaida.sSiglaMovto, objPadraoTribSaida.sCategoriaFilialCliente, objPadraoTribSaida.sItemCategoria)

        Case 33375
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
        
        Case 40557, 40559, 40565, 40571
            
        Case 40558
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_INEXISTENTE", Err, objPadraoTribSaida.sNaturezaOperacao)
        
        Case 40560
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIACLIENTEITEM_INEXISTENTE", Err, objPadraoTribSaida.sItemCategoria, objPadraoTribSaida.sCategoriaFilialCliente)
       
        Case 40566
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_TRIBUTACAO_NAO_CADASTRADO", Err, objPadraoTribSaida.iTipoTributacaoPadrao)
                           
        Case 40572
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPODOCINFO_NAO_CADASTRADO", Err, objPadraoTribSaida.sSiglaMovto)
                           
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153640)

    End Select

    Call Transacao_Rollback

    'Fecha os comandos
    For iIndice = 0 To 5
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function PadraoTribSaida_Exclui(ByVal objPadraoTribSaida As ClassPadraoTribSaida) As Long
'Exclui um Padrão de Tributação Saída existente

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(2) As Long
Dim lTransacao As Long
Dim iTipoTributacaoPadrao As Integer

On Error GoTo Erro_PadraoTribSaida_Exclui

    'Abertura comandos
    For iIndice = 0 To 1

        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 33389

    Next

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 33390

    'Pesquisa o Padrão de Tributação Saída
    lErro = Comando_ExecutarPos(alComando(0), "SELECT TipoTributacaoPadrao FROM PadroesTribSaida WHERE NaturezaOperacao = ? AND SiglaMovto = ? AND CategoriaFilialCliente = ? AND ItemCategoria = ?", 0, iTipoTributacaoPadrao, objPadraoTribSaida.sNaturezaOperacao, objPadraoTribSaida.sSiglaMovto, objPadraoTribSaida.sCategoriaFilialCliente, objPadraoTribSaida.sItemCategoria)
    If lErro <> AD_SQL_SUCESSO Then Error 33391

    'Tenta ler
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 33392

    'Se não encontrou o Padrão de Tributação Saída ==> erro
    If lErro = AD_SQL_SEM_DADOS Then Error 33393

    'Lock da tabela de PadroesTribSaida
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> AD_SQL_SUCESSO Then Error 33394

    'Exclui o Padrão de Tributação Saída
    lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM PadroesTribSaida", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then Error 33395

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 33396

    'Fechamento comandos
    For iIndice = 0 To 1
        Call Comando_Fechar(alComando(iIndice))
    Next

    PadraoTribSaida_Exclui = SUCESSO

    Exit Function

Erro_PadraoTribSaida_Exclui:

    PadraoTribSaida_Exclui = Err

    Select Case Err

        Case 33389
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 33390
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 33391, 33392
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PADROESTRIBUTACAO1", Err, objPadraoTribSaida.sNaturezaOperacao, objPadraoTribSaida.sSiglaMovto, objPadraoTribSaida.sCategoriaFilialCliente, objPadraoTribSaida.sItemCategoria)

        Case 33393
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PADRAO_TRIBUTACAO_NAO_CADASTRADO", Err, objPadraoTribSaida.sNaturezaOperacao, objPadraoTribSaida.sSiglaMovto, objPadraoTribSaida.sCategoriaFilialCliente, objPadraoTribSaida.sItemCategoria)

        Case 33394
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PADROESTRIBUTACAO", Err, objPadraoTribSaida.sNaturezaOperacao, objPadraoTribSaida.sSiglaMovto, objPadraoTribSaida.sCategoriaFilialCliente, objPadraoTribSaida.sItemCategoria)

        Case 33395
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PADROESTRIBUTACAO", Err, objPadraoTribSaida.sNaturezaOperacao, objPadraoTribSaida.sSiglaMovto, objPadraoTribSaida.sCategoriaFilialCliente, objPadraoTribSaida.sItemCategoria)

        Case 33396
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153641)

    End Select

    'Fechamento transação
    Call Transacao_Rollback

    'Fechamento comandos
    For iIndice = 0 To 1
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function PadraoTribEntrada_Grava(ByVal objPadraoTribEnt As ClassPadraoTribEnt) As Long
'Inclui ou Altera um Padrão de Tributação Entrada

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(6) As Long
Dim lTransacao As Long
Dim iTipoTributacaoPadrao As Integer

On Error GoTo Erro_PadraoTribEntrada_Grava

    'Abre os comandos
    For iIndice = 0 To 5
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 33490
    Next

    'Abre a transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 33491

    'Pesquisar o Padrão de Tributação Entrada
    lErro = Comando_ExecutarPos(alComando(0), "SELECT TipoTributacaoPadrao FROM PadroesTribEntrada WHERE NaturezaOperacao = ? AND SiglaMovto = ? AND CategoriaProduto = ? AND ItemCategoria = ?", 0, iTipoTributacaoPadrao, objPadraoTribEnt.sNaturezaOperacao, objPadraoTribEnt.sSiglaMovto, objPadraoTribEnt.sCategoriaProduto, objPadraoTribEnt.sItemCategoria)
    If lErro <> AD_SQL_SUCESSO Then Error 33492

    'Tenta ler
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 33493

    'Encontrou o Padrão de Tributação Entrada
    If lErro = AD_SQL_SUCESSO Then 'MODIFICAÇÃO

        '"Lock" do Padrão de Tributação
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then Error 33494

        If iTipoTributacaoPadrao <> objPadraoTribEnt.iTipoTributacaoPadrao Then

            'Atualiza a tabela de Padrão de Tributação Entrada
            lErro = Comando_ExecutarPos(alComando(1), "UPDATE PadroesTribEntrada SET TipoTributacaoPadrao = ?", alComando(0), objPadraoTribEnt.iTipoTributacaoPadrao)
            If lErro <> AD_SQL_SUCESSO Then Error 33495

        End If

    Else 'GRAVAÇÃO

        'Verifica se a CategoriaProduto foi preenchida
        If objPadraoTribEnt.sCategoriaProduto <> "" Then

            'Verifica se a Item da CategoriaProduto foi preenchido
            If objPadraoTribEnt.sItemCategoria <> "" Then

                'Pequisa  o Item da Categoria
                lErro = CF("CategoriaProdutoItem_Lock", objPadraoTribEnt.sCategoriaProduto, objPadraoTribEnt.sItemCategoria, alComando(2))
                If lErro <> SUCESSO And lErro <> 40577 Then Error 40573
                
                'Se não encontrou ==> erro
                If lErro = 40577 Then Error 40574

            End If

        End If

        'Verifica se a Natureza da Operação foi preenchida
        If objPadraoTribEnt.sNaturezaOperacao <> "" Then

            'Pequisa a Natureza da Operação
            lErro = CF("NaturezaOperacao_Lock", objPadraoTribEnt.sNaturezaOperacao, alComando(3))
            If lErro <> SUCESSO And lErro <> 40551 Then Error 40579
            
            'Se não encontrou ==> erro
            If lErro = 40551 Then Error 40580
       
        End If

        'Verifica se o Tipo Documento foi preenchido
        If objPadraoTribEnt.sSiglaMovto <> "" Then
            
            'Pesquisa o Tipo de Documento e faz "Lock" do Tipo do Documento
            lErro = CF("TipoDocInfo_Lock_Sigla", objPadraoTribEnt.sSiglaMovto, alComando(4))
            If lErro <> SUCESSO And lErro <> 40569 Then Error 40581
        
            'Se não encontrou ----> Erro
            If lErro = 40569 Then Error 40582
       
        End If
        
        'Pesquisa o tipo de Tributação
        lErro = CF("TipoTributacao_Lock", objPadraoTribEnt.iTipoTributacaoPadrao, alComando(5))
        If lErro <> SUCESSO And lErro <> 40563 Then Error 40583
        
        'Se não encontrou -----> Erro
        If lErro = 40563 Then Error 40584
               
        'Insere um novo registro na tabela de PadroesTribEntrada
        lErro = Comando_Executar(alComando(0), "INSERT INTO PadroesTribEntrada(NaturezaOperacao, SiglaMovto, CategoriaProduto, ItemCategoria, TipoTributacaoPadrao) VALUES (?,?,?,?,?)", objPadraoTribEnt.sNaturezaOperacao, objPadraoTribEnt.sSiglaMovto, objPadraoTribEnt.sCategoriaProduto, objPadraoTribEnt.sItemCategoria, objPadraoTribEnt.iTipoTributacaoPadrao)
        If lErro <> AD_SQL_SUCESSO Then Error 33512

    End If

    'Faz o commit da transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 33513

    'Fecha os comandos
    For iIndice = 0 To 5
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    PadraoTribEntrada_Grava = SUCESSO

    Exit Function

Erro_PadraoTribEntrada_Grava:

    PadraoTribEntrada_Grava = Err

    Select Case Err

        Case 33490
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 33491
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 33492, 33493
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PADROESTRIBENTRADA", Err, objPadraoTribEnt.sNaturezaOperacao, objPadraoTribEnt.sSiglaMovto, objPadraoTribEnt.sCategoriaProduto, objPadraoTribEnt.sItemCategoria)

        Case 33494
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PADROESTRIBENTRADA", Err, objPadraoTribEnt.sNaturezaOperacao, objPadraoTribEnt.sSiglaMovto, objPadraoTribEnt.sCategoriaProduto, objPadraoTribEnt.sItemCategoria)

        Case 33495
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_PADROESTRIBENTRADA", Err, objPadraoTribEnt.sNaturezaOperacao, objPadraoTribEnt.sSiglaMovto, objPadraoTribEnt.sCategoriaProduto, objPadraoTribEnt.sItemCategoria)
        
        Case 33512
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PADROESTRIBENTRADA", Err, objPadraoTribEnt.sNaturezaOperacao, objPadraoTribEnt.sSiglaMovto, objPadraoTribEnt.sCategoriaProduto, objPadraoTribEnt.sItemCategoria)

        Case 33513
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
        
        Case 40573, 40579, 40581, 40583
        
        Case 40574
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTOITEM_INEXISTENTE", Err, objPadraoTribEnt.sItemCategoria, objPadraoTribEnt.sCategoriaProduto)
        
        Case 40580
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_INEXISTENTE", Err, objPadraoTribEnt.sNaturezaOperacao)
             
        Case 40582
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPODOCINFO_NAO_CADASTRADO", Err, objPadraoTribEnt.sSiglaMovto)

        Case 40584
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_TRIBUTACAO_NAO_CADASTRADO", Err, objPadraoTribEnt.iTipoTributacaoPadrao)
               
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153642)

    End Select

    Call Transacao_Rollback

    'Fecha os comandos
    For iIndice = 0 To 5
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function PadraoTribEntrada_Exclui(ByVal objPadraoTribEnt As ClassPadraoTribEnt) As Long
'Exclui um Padrão de Tributação Entrada existente

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(2) As Long
Dim lTransacao As Long
Dim iTipoTributacaoPadrao As Integer

On Error GoTo Erro_PadraoTribEntrada_Exclui

    'Abertura comandos
    For iIndice = 0 To 1

        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 33520

    Next

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 33521

    'Pesquisa o Padrão de Tributação Entrada
    lErro = Comando_ExecutarPos(alComando(0), "SELECT TipoTributacaoPadrao FROM PadroesTribEntrada WHERE NaturezaOperacao = ? AND SiglaMovto = ? AND CategoriaProduto = ? AND ItemCategoria = ?", 0, iTipoTributacaoPadrao, objPadraoTribEnt.sNaturezaOperacao, objPadraoTribEnt.sSiglaMovto, objPadraoTribEnt.sCategoriaProduto, objPadraoTribEnt.sItemCategoria)
    If lErro <> AD_SQL_SUCESSO Then Error 33522

    'Tenta ler
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 33523

    'Se não encontrou o Padrão de Tributação Entrada ==> erro
    If lErro = AD_SQL_SEM_DADOS Then Error 33524

    'Lock da tabela de PadroesTribEntrada
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> AD_SQL_SUCESSO Then Error 33525

    'Exclui o Padrão de Tributação Entrada
    lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM PadroesTribEntrada", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then Error 33526

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 33527

    'Fechamento comandos
    For iIndice = 0 To 1
        Call Comando_Fechar(alComando(iIndice))
    Next

    PadraoTribEntrada_Exclui = SUCESSO

    Exit Function

Erro_PadraoTribEntrada_Exclui:

    PadraoTribEntrada_Exclui = Err

    Select Case Err

        Case 33520
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 33521
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 33522, 33523
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PADROESTRIBENTRADA", Err, objPadraoTribEnt.sNaturezaOperacao, objPadraoTribEnt.sSiglaMovto, objPadraoTribEnt.sCategoriaProduto, objPadraoTribEnt.sItemCategoria)

        Case 33524
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PADRAOTRIBENTRADA_NAO_CADASTRADO", Err, objPadraoTribEnt.sNaturezaOperacao, objPadraoTribEnt.sSiglaMovto, objPadraoTribEnt.sCategoriaProduto, objPadraoTribEnt.sItemCategoria)

        Case 33525
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PADROESTRIBENTRADA", Err, objPadraoTribEnt.sNaturezaOperacao, objPadraoTribEnt.sSiglaMovto, objPadraoTribEnt.sCategoriaProduto, objPadraoTribEnt.sItemCategoria)

        Case 33526
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PADROESTRIBENTRADA", Err, objPadraoTribEnt.sNaturezaOperacao, objPadraoTribEnt.sSiglaMovto, objPadraoTribEnt.sCategoriaProduto, objPadraoTribEnt.sItemCategoria)

        Case 33527
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153643)

    End Select

    'Fechamento transação
    Call Transacao_Rollback

    'Fechamento comandos
    For iIndice = 0 To 1
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function
    
End Function

Private Function TipoTributacao_PodeExcluir(ByVal objTipoTributacao As ClassTipoDeTributacaoMovto, ByVal lComando As Long) As Long
'verifica se o tipo de tributacao nao está sendo utilizado, o que impediria a sua exclusao
    
Dim lErro As Long
Dim iTipoPadrao As Integer
On Error GoTo Erro_TipoTributacao_PodeExcluir

    'Verifica se o Tipo é usado em padrão de saida
    lErro = Comando_Executar(lComando, "SELECT TipoTributacaoPadrao FROM PadroesTribSaida WHERE TipoTributacaoPadrao = ?", iTipoPadrao, objTipoTributacao.iTipo)
    If lErro <> AD_SQL_SUCESSO Then Error 33319

    'Tenta ler
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 33320

    'Achou o Tipo de Tributação Padrão ==> erro
    If lErro = AD_SQL_SUCESSO Then Error 33321

    'Verifica se o Tipo é usado em padrão de entrada
    lErro = Comando_Executar(lComando, "SELECT TipoTributacaoPadrao FROM PadroesTribEntrada WHERE TipoTributacaoPadrao = ?", iTipoPadrao, objTipoTributacao.iTipo)
    If lErro <> AD_SQL_SUCESSO Then Error 41507

    'Tenta ler
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 41508

    'Achou o Tipo de Tributação Padrão ==> erro
    If lErro = AD_SQL_SUCESSO Then Error 41509

    'Verifica se o Tipo é usado em NFiscal
    lErro = Comando_Executar(lComando, "SELECT TipoTributacao FROM TributacaoNF WHERE TipoTributacao = ?", iTipoPadrao, objTipoTributacao.iTipo)
    If lErro <> AD_SQL_SUCESSO Then Error 58841

    'Tenta ler
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 58842

    'Achou o Tipo de Tributação Padrão ==> erro
    If lErro = AD_SQL_SUCESSO Then Error 58843
    
    'Verifica se o Tipo é usado em TipoDocInfo
    lErro = Comando_Executar(lComando, "SELECT TipoOperacaoTrib FROM TiposDocInfo WHERE TipoOperacaoTrib = ?", iTipoPadrao, objTipoTributacao.iTipo)
    If lErro <> AD_SQL_SUCESSO Then Error 58844

    'Tenta ler
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 58845

    'Achou o Tipo de Tributação Padrão ==> erro
    If lErro = AD_SQL_SUCESSO Then Error 58846
    
    'Verifica se o Tipo é usado em Ped. de Venda
    lErro = Comando_Executar(lComando, "SELECT TipoTributacao FROM TributacaoPV WHERE (TipoTributacao =? OR TipoTributacaoManual =?)", iTipoPadrao, objTipoTributacao.iTipo, objTipoTributacao.iTipo)
    If lErro <> AD_SQL_SUCESSO Then Error 58847

    'Tenta ler
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 58848

    'Achou o Tipo de Tributação Padrão ==> erro
    If lErro = AD_SQL_SUCESSO Then Error 58849
    
    TipoTributacao_PodeExcluir = SUCESSO
    
    Exit Function
    
Erro_TipoTributacao_PodeExcluir:

    TipoTributacao_PodeExcluir = Err
    
    Select Case Err

        Case 33319, 33320, 41507, 41508
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PADROESTRIBUTACAO", Err, objTipoTributacao.iTipo)

        Case 33321, 41509
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TIPO_TRIBUTACAO_PADRAO", Err, objTipoTributacao.iTipo)
        
        Case 58841, 58842
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRIBUTACAONF", Err)
        
        Case 58843
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TIPO_TRIBUTACAO_NF", Err, objTipoTributacao.iTipo)
        
        Case 58844, 58845
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPODOCINFO1", Err)
        
        Case 58846
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TIPO_TRIBUTACAO_TIPODOCINFO", Err, objTipoTributacao.iTipo)
        
        Case 58847, 58848
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRIBUTACAOPV", Err)
        
        Case 58849
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TIPO_TRIBUTACAO_PV", Err, objTipoTributacao.iTipo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153644)

    End Select
    
    Exit Function

End Function

Function TipoTributacao_Lock(ByVal iTipoTributacaoPadrao As Integer, ByVal lComando As Long) As Long
'faz lock shared de um tipo de tributacao para garantir a sua existencia

Dim lErro As Long
Dim iTipo As Integer

On Error GoTo Erro_TipoTributacao_Lock

    'Pesquisa o Tipo de Tributação
    lErro = Comando_ExecutarLockado(lComando, "SELECT Tipo FROM TiposDeTributacaoMovto WHERE Tipo = ?", iTipo, iTipoTributacaoPadrao)
    If lErro <> AD_SQL_SUCESSO Then Error 40561

    'Tenta ler
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 40562

    'Se não encontrou ==> erro
    If lErro = AD_SQL_SEM_DADOS Then Error 40563

    '"Lock" da TiposDeTributacaoMovto
    lErro = Comando_LockShared(lComando)
    If lErro <> SUCESSO Then Error 40564

    TipoTributacao_Lock = SUCESSO
    
    Exit Function
    
Erro_TipoTributacao_Lock:

    TipoTributacao_Lock = Err
    
    Select Case Err

        Case 40561, 40562
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPO_TRIBUTACAO", Err, iTipoTributacaoPadrao)
        
        Case 40563
        
        Case 40564
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_TIPOSDETRIBUTACAOMOVTO", Err, iTipoTributacaoPadrao)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153645)

    End Select

    Exit Function
    
End Function

'*** INICIO

Function IPIExcecao_Grava(ByVal objIPIExcecoes As ClassIPIExcecao) As Long
'grava as alterações/inclusões realizadas pelo usuário nos dados de IPIExcecoes para um registro.

Dim lErro As Long
Dim iTipo As Integer
Dim lTransacao As Long
Dim sDescricao As String
Dim alComando(5) As Long
Dim iIndice As Integer

On Error GoTo Erro_IPIExcecao_Grava

    'Abre transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 21625

    'Abre Comandos
    For iIndice = LBound(alComando) To UBound(alComando)

        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 21626

    Next

    If Len(objIPIExcecoes.sCategoriaCliente) > 0 Then

        sDescricao = String(STRING_CATEGORIACLIENTEITEM_DESCRICAO, 0)

        'Leitura da tabela de itens das categorias de clientes
        lErro = Comando_ExecutarPos(alComando(3), "SELECT Descricao FROM CategoriaClienteItem WHERE Categoria = ? AND Item =? ", 0, _
        sDescricao, objIPIExcecoes.sCategoriaCliente, objIPIExcecoes.sCategoriaClienteItem)
        If lErro <> AD_SQL_SUCESSO Then Error 21630

        lErro = Comando_BuscarPrimeiro(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 21631

        If lErro = AD_SQL_SEM_DADOS Then Error 21632

        lErro = Comando_LockShared(alComando(3))
        If lErro <> AD_SQL_SUCESSO Then Error 21640

    End If

    If Len(objIPIExcecoes.sCategoriaProduto) > 0 Then

        sDescricao = String(STRING_CATEGORIAPRODUTOITEM_DESCRICAO, 0)

        'Leitura da tabela de itens de categorias de produtos
        lErro = Comando_ExecutarPos(alComando(5), "SELECT Descricao FROM CategoriaProdutoItem WHERE Categoria = ? AND Item =? ", 0, _
        sDescricao, objIPIExcecoes.sCategoriaProduto, objIPIExcecoes.sCategoriaProdutoItem)
        If lErro <> AD_SQL_SUCESSO Then Error 21636

        lErro = Comando_BuscarPrimeiro(alComando(5))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 21637

        If lErro = AD_SQL_SEM_DADOS Then Error 21638

        lErro = Comando_LockShared(alComando(5))
        If lErro <> AD_SQL_SUCESSO Then Error 21642

    End If

    'Inserção ou atualização dos dados
    lErro = IPIExcecao_Grava1(objIPIExcecoes, alComando())
    If lErro <> SUCESSO Then Error 21643

    'Confirma transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 21644

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    IPIExcecao_Grava = SUCESSO

    Exit Function

Erro_IPIExcecao_Grava:

    IPIExcecao_Grava = Err

    Select Case Err

        Case 21625
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 21626
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 21627, 21628
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CATEGORIACLIENTE", Err)

        Case 21629
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIACLIENTE_INEXISTENTE", Err, objIPIExcecoes.sCategoriaCliente)

        Case 21630, 21631
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CATEGORIACLIENTEITEM2", Err, objIPIExcecoes.sCategoriaCliente, objIPIExcecoes.sCategoriaClienteItem)

        Case 21632
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIACLIENTEITEM_INEXISTENTE", Err, objIPIExcecoes.sCategoriaCliente, objIPIExcecoes.sCategoriaClienteItem)

        Case 21633, 21634
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CATEGORIAPRODUTO", Err)

        Case 21635
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTO1_INEXISTENTE", Err, objIPIExcecoes.sCategoriaProduto)

        Case 21636, 21637
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CATEGORIAPRODUTOITEM2", Err, objIPIExcecoes.sCategoriaProduto, objIPIExcecoes.sCategoriaProdutoItem)

        Case 21638
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTOITEM_INEXISTENTE", Err, objIPIExcecoes.sCategoriaProduto, objIPIExcecoes.sCategoriaProdutoItem)

        Case 21639
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_CATEGORIACLIENTE", Err)

        Case 21640
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_CATEGORIACLIENTEITEM", Err)

        Case 21641
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_CATEGORIAPRODUTO", Err)

        Case 21642
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_CATEGORIAPRODUTOITEM", Err)

        Case 21644
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case 21643

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153646)

    End Select

    Call Transacao_Rollback

    For iIndice = LBound(alComando) To UBound(alComando)

        Call Comando_Fechar(alComando(iIndice))

    Next

    Exit Function

End Function

Private Function IPIExcecao_Grava1(ByVal objIPIExcecoes As ClassIPIExcecao, alComando() As Long) As Long
'grava as alterações/inclusões realizadas pelo usuário nos dados de IPIExcecoes para um registro.

Dim lErro As Long
Dim iTipo As Integer
Dim sDescricao As String
Dim iIndice As Integer

On Error GoTo Erro_IPIExcecao_Grava1

    'Leitura da tabela de exceções de IPI
    lErro = Comando_ExecutarPos(alComando(1), "SELECT Tipo FROM IPIExcecoes WHERE CategoriaProduto = ? AND CategoriaProdutoItem =? AND CategoriaCliente = ? AND CategoriaClienteItem =? ", 0, _
    iTipo, objIPIExcecoes.sCategoriaProduto, objIPIExcecoes.sCategoriaProdutoItem, objIPIExcecoes.sCategoriaCliente, objIPIExcecoes.sCategoriaClienteItem)
    If lErro <> AD_SQL_SUCESSO Then Error 21645

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 21646

    If lErro = AD_SQL_SEM_DADOS Then

        'Se não existe, realiza inserção
        lErro = Comando_Executar(alComando(1), "INSERT INTO IPIExcecoes (CategoriaProduto ,CategoriaProdutoItem ,CategoriaCliente ,CategoriaClienteItem, Tipo, PercRedBaseCalculo, Aliquota, PercMargemLucro, Fundamentacao, Prioridade, TipoCalculo, AliquotaRS) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)", _
            objIPIExcecoes.sCategoriaProduto, objIPIExcecoes.sCategoriaProdutoItem, objIPIExcecoes.sCategoriaCliente, objIPIExcecoes.sCategoriaClienteItem, objIPIExcecoes.iTipo, objIPIExcecoes.dPercRedBaseCalculo, objIPIExcecoes.dAliquota, objIPIExcecoes.dPercMargemLucro, objIPIExcecoes.sFundamentacao, objIPIExcecoes.iPrioridade, objIPIExcecoes.iTipoCalculo, objIPIExcecoes.dAliquotaRS)
        If lErro <> AD_SQL_SUCESSO Then Error 21647

    Else

        'Se existe, atualiza
        lErro = Comando_ExecutarPos(alComando(2), "UPDATE IPIExcecoes SET Tipo = ? , PercRedBaseCalculo = ? , Aliquota = ? , PercMargemLucro = ?, Fundamentacao = ? , Prioridade = ?, TipoCalculo = ?, AliquotaRS = ? ", alComando(1), _
            objIPIExcecoes.iTipo, objIPIExcecoes.dPercRedBaseCalculo, objIPIExcecoes.dAliquota, objIPIExcecoes.dPercMargemLucro, objIPIExcecoes.sFundamentacao, objIPIExcecoes.iPrioridade, objIPIExcecoes.iTipoCalculo, objIPIExcecoes.dAliquotaRS)
        If lErro <> AD_SQL_SUCESSO Then Error 21648

    End If

    IPIExcecao_Grava1 = SUCESSO

    Exit Function

Erro_IPIExcecao_Grava1:

    IPIExcecao_Grava1 = Err

    Select Case Err

        Case 21645, 21646
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_IPIEXCECOES", Err)

        Case 21647
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_IPIEXCECOES", Err, objIPIExcecoes.sCategoriaCliente, objIPIExcecoes.sCategoriaClienteItem, objIPIExcecoes.sCategoriaProduto, objIPIExcecoes.sCategoriaProdutoItem)

        Case 21648
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_IPIEXCECOES", Err, objIPIExcecoes.sCategoriaCliente, objIPIExcecoes.sCategoriaClienteItem, objIPIExcecoes.sCategoriaProduto, objIPIExcecoes.sCategoriaProdutoItem)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153647)

    End Select

    Exit Function

End Function

Function IPIExcecao_Exclui(ByVal objIPIExcecoes As ClassIPIExcecao) As Long
'grava as exclusões realizadas pelo usuário nos dados de IPIExcecoes para um registro.

Dim lErro As Long
Dim iTipo As Integer
Dim lTransacao As Long
Dim lComando As Long
Dim lComando1 As Long

On Error GoTo Erro_IPIExcecao_Exclui

    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 21649

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 21650

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 21651

    lErro = Comando_ExecutarPos(lComando, "SELECT Tipo FROM IPIExcecoes WHERE CategoriaProduto = ? AND CategoriaProdutoItem =? AND CategoriaCliente = ? AND CategoriaClienteItem =? ", 0, _
    iTipo, objIPIExcecoes.sCategoriaProduto, objIPIExcecoes.sCategoriaProdutoItem, objIPIExcecoes.sCategoriaCliente, objIPIExcecoes.sCategoriaClienteItem)
    If lErro <> AD_SQL_SUCESSO Then Error 21652

    'Lê o primeiro registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 21655

    If lErro = AD_SQL_SEM_DADOS Then Error 21656

    lErro = Comando_ExecutarPos(lComando1, "DELETE FROM IPIExcecoes ", lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 21657

    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 21658

    'Confirma transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 21659

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    IPIExcecao_Exclui = SUCESSO

    Exit Function

Erro_IPIExcecao_Exclui:

    IPIExcecao_Exclui = Err

    Select Case Err

        Case 21649, 21651
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 21650
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 21652, 21653, 21654, 21655, 21658
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_IPIEXCECOES", Err)

        Case 21656
            lErro = Rotina_Erro(vbOKOnly, "ERRO_IPIEXCECOES_INEXISTENTE", Err, objIPIExcecoes.sCategoriaCliente, objIPIExcecoes.sCategoriaClienteItem, objIPIExcecoes.sCategoriaProduto, objIPIExcecoes.sCategoriaProdutoItem)

        Case 21657
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_IPIEXCECOES", Err, objIPIExcecoes.sCategoriaCliente, objIPIExcecoes.sCategoriaClienteItem, objIPIExcecoes.sCategoriaProduto, objIPIExcecoes.sCategoriaProdutoItem)

        Case 21659
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153648)

    End Select

    Call Transacao_Rollback

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Function NFiscal_GravaTributacao(ByVal objNFiscal As ClassNFiscal, ByVal lComando1 As Long, ByVal lComando2 As Long) As Long

Dim lErro As Long

On Error GoTo Erro_NFiscal_GravaTributacao

'    With objNFiscal.objTributacaoNF
'        lErro = Comando_Executar(lComando1, "INSERT INTO TributacaoNF (NumIntDoc, NaturezaOpInterna, TipoTributacao, IPIBase, IPIValor, IPICredito, ICMSBase, ICMSValor, ICMSSubstBase, ICMSSubstValor, ICMSCredito, ISSIncluso, ISSBase, ISSAliquota, ISSValor, IRRFBase, IRRFAliquota, IRRFValor, ValorINSS, INSSRetido, INSSValorBase, INSSValorDeducoes, PISCredito, COFINSCredito, PISRetido, ISSRetido, COFINSRetido, CSLLRetido ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
'            objNFiscal.lNumIntDoc, .sNaturezaOpInterna, .iTipoTributacao, .dIPIBase, .dIPIValor, .dIPICredito, .dICMSBase, .dICMSValor, .dICMSSubstBase, .dICMSSubstValor, .dICMSCredito, .iISSIncluso, .dISSBase, .dISSAliquota, .dISSValor, .dIRRFBase, .dIRRFAliquota, .dIRRFValor, .dValorINSS, .iINSSRetido, .dINSSBase, .dINSSDeducoes, .dPISCredito, .dCOFINSCredito, .dPISRetido, .dISSRetido, .dCOFINSRetido, .dCSLLRetido)
'    End With
'
'    If lErro <> AD_SQL_SUCESSO Then Error 22858
    
'    'inserir os complementos
'    lErro = CF("NFiscal_GravaTributacaoCompl", objNFiscal, objNFiscal.objTributacaoNF.objTributacaoDesconto, lComando2)
'    If lErro <> SUCESSO Then Error 22859
'
'    lErro = CF("NFiscal_GravaTributacaoCompl", objNFiscal, objNFiscal.objTributacaoNF.objTributacaoFrete, lComando2)
'    If lErro <> SUCESSO Then Error 22860
'
'    lErro = CF("NFiscal_GravaTributacaoCompl", objNFiscal, objNFiscal.objTributacaoNF.objTributacaoOutras, lComando2)
'    If lErro <> SUCESSO Then Error 22861
'
'    lErro = CF("NFiscal_GravaTributacaoCompl", objNFiscal, objNFiscal.objTributacaoNF.objTributacaoSeguro, lComando2)
'    If lErro <> SUCESSO Then Error 22862

    objNFiscal.objTributacao.lNumIntDoc = objNFiscal.lNumIntDoc
    objNFiscal.objTributacao.iTipoDoc = TIPODOC_TRIB_NF

    lErro = CF("TributacaoDoc_Grava", objNFiscal.objTributacao)
    If lErro <> SUCESSO Then Error 22862
    
    NFiscal_GravaTributacao = SUCESSO

    Exit Function

Erro_NFiscal_GravaTributacao:

    NFiscal_GravaTributacao = Err

    Select Case Err
'
'        Case 22858
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRIBNFISCAL", Err)
        
        Case 22859 To 22862
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153649)

    End Select

    Exit Function

End Function


'Function NFiscal_GravaTributacaoCompl(ByVal objNFiscal As ClassNFiscal, ByVal objTribComplNF As ClassTribComplNF, ByVal lComando As Long) As Long
'
'Dim lErro As Long
'
'On Error GoTo Erro_NFiscal_GravaTributacaoCompl
'
''    With objTribComplNF
''        lErro = Comando_Executar(lComando, "INSERT INTO TributacaoComplNF (NumIntDoc, Tipo, NaturezaOp, TipoTributacao, IPITipo, IPIBaseCalculo, IPIPercRedBase, IPIAliquota, IPIValor, IPICredito, ICMSTipo, ICMSBase, ICMSPercRedBase, ICMSAliquota, ICMSValor, ICMSCredito, ICMSSubstBase, ICMSSubstAliquota, ICMSSubstValor, PISCredito, COFINSCredito ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
''            objNFiscal.lNumIntDoc, .iTipo, .sNaturezaOp, .iTipoTributacao, .iIPITipo, .dIPIBaseCalculo, .dIPIPercRedBase, .dIPIAliquota, .dIPIValor, .dIPICredito, .iICMSTipo, .dICMSBase, .dICMSPercRedBase, .dICMSAliquota, .dICMSValor, .dICMSCredito, .dICMSSubstBase, .dICMSSubstAliquota, .dICMSSubstValor, .dPISCredito, .dCOFINSCredito)
''    End With
''
''    If lErro <> AD_SQL_SUCESSO Then Error 22863
'
'    NFiscal_GravaTributacaoCompl = SUCESSO
'
'    Exit Function
'
'Erro_NFiscal_GravaTributacaoCompl:
'
'    NFiscal_GravaTributacaoCompl = Err
'
'    Select Case Err
'
'        Case 22863
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRIBCOMPLNFISCAL", Err)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153650)
'
'    End Select
'
'    Exit Function
'
'End Function

Function NFiscal_GravaTributacaoItem(ByVal objItemNF As ClassItemNF, ByVal lComando As Long, ByVal lComandoProd As Long) As Long

Dim lErro As Long, tProduto As typeProduto

On Error GoTo Erro_NFiscal_GravaTributacaoItem
 
'    If objItemNF.objTributacaoItemNF.sIPICodProduto = "" Then
'
'        'obter o cod do ipi e a origem do produto p/obter a cst
'        tProduto.sIPICodigo = String(STRING_PRODUTO_IPI_CODIGO, 0)
'        lErro = Comando_Executar(lComandoProd, "SELECT OrigemMercadoria, IPICodigo FROM Produtos WHERE Codigo = ?", tProduto.iOrigemMercadoria, tProduto.sIPICodigo, objItemNF.sProduto)
'        If lErro <> AD_SQL_SUCESSO Then Error 56959
'
'        lErro = Comando_BuscarProximo(lComandoProd)
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 56960
'
'        If lErro <> AD_SQL_SUCESSO Then Error 56961
'
'        objItemNF.objTributacaoItemNF.sIPICodProduto = tProduto.sIPICodigo
'
'    End If
    
'    With objItemNF.objTributacaoItemNF
'        lErro = Comando_Executar(lComando, "INSERT INTO TributacaoItemNF (NumIntNF, Item, NumIntDoc, NaturezaOp, TipoTributacao, IPITipo, IPICodProduto, IPIBaseCalculo, IPIPercRedBase, IPIAliquota, IPIValor, IPICredito, ICMSTipo, ICMSBase, ICMSPercRedBase, ICMSAliquota, ICMSValor, ICMSCredito, ICMSSubstBase, ICMSSubstAliquota, ICMSSubstValor, PISCredito, COFINSCredito, ICMSAliquotaAdicaoDI, DespImpICMSBase, DespImpICMSValor, DespImpICMSCredito, ICMSSubstPercRedBase, ICMSSubstPercMVA, " & _
'        "PISTipo, PISTipoCalculo, PISBase, PISAliquota, PISAliquotaValor, PISQtde, PISValor, PISSTTipoCalculo, PISSTBase, PISSTAliquota, PISSTAliquotaValor, PISSTQtde, PISSTValor, COFINSTipo, COFINSTipoCalculo, COFINSBase, COFINSAliquota, COFINSAliquotaValor, COFINSQtde, COFINSValor, COFINSSTTipoCalculo, COFINSSTBase, COFINSSTAliquota, COFINSSTAliquotaValor, COFINSSTQtde,COFINSSTValor, CST, ISSQN, ISSBase, ISSAliquota, ISSCidadeIBGE, ISSValor, IIBase, IIDespAduaneira, IIIOF, IIValor, ICMSBaseModalidade, ICMSSubstBaseModalidade, IPIEnquadramentoClasse, IPIEnquadramentoCodigo, IPICNPJProdutor, IPISeloCodigo, IPISeloQtde, IPITipoCalculo, IPIUnidadePadraoQtde, IPIUnidadePadraoValor, ValorFreteItem, ValorSeguroItem, ValorOutrasDespesasItem, ValorDescontoItem, OrigemMercadoria, ExTIPI, Genero, ProdutoEspecifico, EAN, EANTrib, QtdTrib, UMTrib, ValorUnitTrib" & _
'        ") VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
'            objItemNF.lNumIntNF, objItemNF.iItem, objItemNF.lNumIntDoc, .sNaturezaOp, .iTipoTributacao, .iIPITipo, .sIPICodProduto, .dIPIBaseCalculo, .dIPIPercRedBase, .dIPIAliquota, .dIPIValor, .dIPICredito, .iICMSTipo, .dICMSBase, .dICMSPercRedBase, .dICMSAliquota, .dICMSValor, .dICMSCredito, .dICMSSubstBase, .dICMSSubstAliquota, .dICMSSubstValor, .dPISCredito, .dCOFINSCredito, .dICMSAliquotaAdicaoDI, .dDespImpICMSBase, .dDespImpICMSValor, .dDespImpICMSCredito, .dICMSSubstPercRedBase, .dICMSSubstPercMVA, _
'            .iPISTipo, .iPISTipoCalculo, .dPISBase, .dPISAliquota, .dPISAliquotaValor, .dPISQtde, .dPISValor, .iPISSTTipoCalculo, .dPISSTBase, .dPISSTAliquota, .dPISSTAliquotaValor, .dPISSTQtde, .dPISSTValor, .iCOFINSTipo, .iCOFINSTipoCalculo, .dCOFINSBase, .dCOFINSAliquota, .dCOFINSAliquotaValor, .dCOFINSQtde, .dCOFINSValor, .iCOFINSSTTipoCalculo, .dCOFINSSTBase, .dCOFINSSTAliquota, .dCOFINSSTAliquotaValor, .dCOFINSSTQtde, .dCOFINSSTValor, .sCST, .sISSQN, .dISSBase, .dISSAliquota, .sISSCidadeIBGE, .dISSValor, .dIIBase, .dIIDespAduaneira, .dIIIOF, .dIIValor, .iICMSBaseModalidade, .iICMSSubstBaseModalidade, .sIPIEnquadramentoClasse, .sIPIEnquadramentoCodigo, .sIPICNPJProdutor, .sIPISeloCodigo, .lIPISeloQtde, .iIPITipoCalculo, .dIPIUnidadePadraoQtde, .dIPIUnidadePadraoValor, .dValorFreteItem, .dValorSeguroItem, .dValorOutrasDespesasItem, .dValorDescontoItem, .iOrigemMercadoria, .iExTIPI, .sGenero, .iProdutoEspecifico, .sEAN, .sEANTrib, .dQtdTrib, .sUMTrib, .dValorUnitTrib)
'    End With
'
'    If lErro <> AD_SQL_SUCESSO Then
'        Error 22857
'    End If

    objItemNF.objTributacao.lNumIntDoc = objItemNF.lNumIntNF
    objItemNF.objTributacao.lNumIntDocItem = objItemNF.lNumIntDoc
    objItemNF.objTributacao.iItem = objItemNF.iItem
    objItemNF.objTributacao.dDescontoGrid = objItemNF.dValorDesconto
    objItemNF.objTributacao.dPrecoUnitario = objItemNF.dPrecoUnitario
    objItemNF.objTributacao.dPrecoUnitarioMoeda = objItemNF.dPrecoUnitarioMoeda
    objItemNF.objTributacao.dQuantidade = objItemNF.dQuantidade
    objItemNF.objTributacao.sProduto = objItemNF.sProduto
    objItemNF.objTributacao.sProdutoDescricao = objItemNF.sDescricaoItem
    objItemNF.objTributacao.sUnidadeMed = objItemNF.sUnidadeMed
    objItemNF.objTributacao.iTipoDoc = TIPODOC_TRIB_NF

    lErro = CF("TributacaoDocItem_Grava", objItemNF.objTributacao)
    If lErro <> SUCESSO Then Error 56961
    
    NFiscal_GravaTributacaoItem = SUCESSO

    Exit Function

Erro_NFiscal_GravaTributacaoItem:

    NFiscal_GravaTributacaoItem = Err

    Select Case Err

        Case 56959, 56960
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS", Err, objItemNF.sProduto)
        
        Case 56961
            'lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", Err, objItemNF.sProduto)
        
        Case 22857
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRIBITEMNFISCAL", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153651)

    End Select

    Exit Function

End Function

Function NFiscal_ExcluiTributacao(ByVal objNFiscal As ClassNFiscal) As Long
'Exclui do bd os registros referentes à tributacao associados à nf passada como parametro

Dim lErro As Long, iIndice As Integer, alComando(1 To 6) As Long, lNumIntDocAux As Long, iItem As Integer, iTipo As Integer

On Error GoTo Erro_NFiscal_ExcluiTributacao

'    'Abre os comandos
'     For iIndice = LBound(alComando) To UBound(alComando)
'        alComando(iIndice) = Comando_Abrir()
'        If alComando(iIndice) = 0 Then gError 81542
'    Next
'
'    'seleciona e exlui registro de tributacaonf
'    lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntDoc FROM TributacaoNF WHERE NumIntDoc = ?", 0, lNumIntDocAux, objNFiscal.lNumIntDoc)
'    If lErro <> AD_SQL_SUCESSO Then gError 81543
'
'    lErro = Comando_BuscarProximo(alComando(1))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81544
'
'    If lErro <> AD_SQL_SEM_DADOS Then
'
'        lErro = Comando_ExecutarPos(alComando(2), "DELETE FROM TributacaoNF", alComando(1))
'        If lErro <> AD_SQL_SUCESSO Then gError 81545
'
'    End If
'
'    'seleciona e exlui registros de tributacaoitemnf
'    lErro = Comando_ExecutarPos(alComando(3), "SELECT NumIntNF, Item FROM TributacaoItemNF WHERE NumIntNF = ?", 0, lNumIntDocAux, iItem, objNFiscal.lNumIntDoc)
'    If lErro <> AD_SQL_SUCESSO Then gError 81546
'
'    lErro = Comando_BuscarProximo(alComando(3))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81547
'
'    Do While lErro <> AD_SQL_SEM_DADOS
'
'        lErro = Comando_ExecutarPos(alComando(4), "DELETE FROM TributacaoItemNF", alComando(3))
'        If lErro <> AD_SQL_SUCESSO Then gError 81548
'
'        lErro = Comando_BuscarProximo(alComando(3))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81549
'
'    Loop
'
'    'seleciona e exlui registros de tributacaocomplnf
'    lErro = Comando_ExecutarPos(alComando(5), "SELECT NumIntDoc, Tipo FROM TributacaoComplNF WHERE NumIntDoc = ?", 0, lNumIntDocAux, iTipo, objNFiscal.lNumIntDoc)
'    If lErro <> AD_SQL_SUCESSO Then gError 81550
'
'    lErro = Comando_BuscarProximo(alComando(5))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81551
'
'    Do While lErro <> AD_SQL_SEM_DADOS
'
'        lErro = Comando_ExecutarPos(alComando(6), "DELETE FROM TributacaoComplNF", alComando(5))
'        If lErro <> AD_SQL_SUCESSO Then gError 81552
'
'        lErro = Comando_BuscarProximo(alComando(5))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81553
'
'    Loop
'
'    'Fechamento dos comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next

    Select Case objNFiscal.iTipoNFiscal
        
        Case TIPODOCINFO_CONHECIMENTOFRETE_FATURA, TIPODOCINFO_CONHECIMENTOFRETE, TIPODOCINFO_CONHECIMENTOFRETE_COMPLICMS, TIPODOCINFO_CONHECIMENTOFRETE_FATURA_COMPLICMS
        
        Case Else
            lErro = CF("TributacaoDoc_Exclui", TIPODOC_TRIB_NF, objNFiscal.lNumIntDoc)
            If lErro <> SUCESSO Then gError 81552
    
    End Select
    
    NFiscal_ExcluiTributacao = SUCESSO
     
    Exit Function
    
Erro_NFiscal_ExcluiTributacao:

    NFiscal_ExcluiTributacao = gErr
     
    Select Case gErr
'
'        Case 81542
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 81543, 81544
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRIBUTACAONF", gErr)
'
'        Case 81545
'            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRIBUTACAONF", gErr)
'
'        Case 81546, 81547, 81549
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRIBUTACAOITEMNF", gErr)
'
'        Case 81548
'            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRIBUTACAOITEMNF", gErr)
'
'        Case 81550, 81551, 81553
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRIBUTACAOCOMPLNF", gErr)
                
        Case 81552
            'Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRIBUTACAOCOMPLNF", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153652)
     
    End Select
     
    'Fechamento dos comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Function Tributacao_Iniciar(ByVal colDriversTributos As Collection) As Long

Dim lErro As Long
Dim objTributoDriver As ClassTributoDriver

On Error GoTo Erro_Tributacao_Iniciar

    Set objTributoDriver = New ClassTributoISS
    lErro = objTributoDriver.Ativar()
    If lErro <> SUCESSO Then Error 27247
    colDriversTributos.Add objTributoDriver
    
    Set objTributoDriver = New ClassTributoIR
    lErro = objTributoDriver.Ativar()
    If lErro <> SUCESSO Then Error 27248
    colDriversTributos.Add objTributoDriver
    
    Set objTributoDriver = New ClassTributoIPI
    lErro = objTributoDriver.Ativar()
    If lErro <> SUCESSO Then Error 27249
    colDriversTributos.Add objTributoDriver
    
    Set objTributoDriver = New ClassTributoICMS
    lErro = objTributoDriver.Ativar()
    If lErro <> SUCESSO Then Error 27250
    colDriversTributos.Add objTributoDriver
        
    Set objTributoDriver = New ClassTributoINSS
    lErro = objTributoDriver.Ativar()
    If lErro <> SUCESSO Then Error 32251
    colDriversTributos.Add objTributoDriver
        
    'valendo a partir de 01/12/2002 pela mp 66
    Set objTributoDriver = New ClassTributoPIS
    lErro = objTributoDriver.Ativar()
    If lErro <> SUCESSO Then Error 32251
    colDriversTributos.Add objTributoDriver
    
    'valendo a partir de 01/02/2004 pela mp 135
    Set objTributoDriver = New ClassTributoCOFINS
    lErro = objTributoDriver.Ativar()
    If lErro <> SUCESSO Then Error 32251
    colDriversTributos.Add objTributoDriver
    
    'valendo a partir de 01/02/2004 pela mp 135
    Set objTributoDriver = New ClassTribArt30
    lErro = objTributoDriver.Ativar()
    If lErro <> SUCESSO Then Error 32251
    colDriversTributos.Add objTributoDriver
    
    'valendo a partir de 08/06/2013 pela lei 12.741/12
    Set objTributoDriver = New ClassTributo12741
    lErro = objTributoDriver.Ativar()
    If lErro <> SUCESSO Then Error 32251
    colDriversTributos.Add objTributoDriver
    
    Tributacao_Iniciar = SUCESSO
     
    Exit Function
    
Erro_Tributacao_Iniciar:

    Tributacao_Iniciar = gErr
     
    Select Case gErr
          
        Case 27247 To 27250, 32251
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153653)
     
    End Select
     
    Exit Function

End Function

Public Function TributacaoDoc_Grava(ByVal objTributacaoDoc As ClassTributacaoDoc) As Long

Dim lErro As Long
Dim alComando(1 To 1) As Long
Dim iIndice As Integer

On Error GoTo Erro_TributacaoDoc_Grava

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 202859
    Next

    With objTributacaoDoc
        lErro = Comando_Executar(alComando(1), "INSERT INTO TributacaoDoc( NaturezaOpInterna,TipoDoc, NumIntDoc, TaxacaoAutomatica, TipoTributacao,TipoTributacaoManual, IPIBase, IPIBaseManual, IPIValor, IPIValorManual, ICMSBase, ICMSBaseManual, ICMSValor, ICMSValorManual, ICMSSubstBase, " & _
                "ICMSSubstBaseManual, ICMSSubstValor, ICMSSubstValorManual, ISSIncluso, ISSBase, ISSBaseManual, ISSAliquota, ISSAliquotaManual, ISSValor, ISSValorManual, IRRFBase, IRRFBaseManual, IRRFAliquota, IRRFAliquotaManual, IRRFValor, " & _
                "IRRFValorManual, PISRetido, COFINSRetido, CSLLRetido, PISRetidoManual, COFINSRetidoManual, CSLLRetidoManual, ISSRetido, ISSRetidoManual, PISCredito, PISCreditoManual, COFINSCredito, COFINSCreditoManual, IPICredito, IPICreditoManual, " & _
                "ICMSCredito, ICMSCreditoManual, ValorINSS, ValorINSSManual, INSSRetido, INSSRetidoManual, INSSValorBase, INSSValorBaseManual, INSSValorDeducoes, INSSValorDeducoesManual, TotTrib, TotTribManual, TotTribTipo, TotTribTipoManual,NaturezaOpManual, " & _
                "PISValor, COFINSValor, ICMSValorIsento, ISSValorDeducao, ISSValorOutrasRet, ISSValorDescIncond, ISSValorDescCond, IIValor, Destino, DestinoManual, FinalidadeNFe, " & _
                "FinalidadeNFeManual , IndConsumidorFinal, IndConsumidorFinalManual, IndPresenca, IndPresencaManual, RegTribEspecial, RegTribEspecialManual, DataPrestServico, DataPrestServicoManual,TotTribFederal,TotTribEstadual,TotTribMunicipal,TotTribFederalManual,TotTribEstadualManual,TotTribMunicipalManual,NFImpAbateIIPrecoUnit,ICMSInterestVlrFCPUFDest,ICMSInterestVlrUFDest,ICMSInterestVlrUFRemet, " & _
                "ICMSVlrFCP, ICMSVlrFCPST, ICMSVlrFCPSTRet, IPIVlrDevolvido)" & _
                "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                .sNaturezaOpInterna, .iTipoDoc, .lNumIntDoc, .iTaxacaoAutomatica, .iTipoTributacao, .iTipoTributacaoManual, .dIPIBase, .iIPIBaseManual, .dIPIValor, .iIPIValorManual, .dICMSBase, .iICMSBaseManual, .dICMSValor, .iICMSValorManual, .dICMSSubstBase, _
                .iICMSSubstBaseManual, .dICMSSubstValor, .iICMSSubstValorManual, .iISSIncluso, .dISSBase, .iISSBaseManual, .dISSAliquota, .iISSAliquotaManual, .dISSValor, .iISSValorManual, .dIRRFBase, .iIRRFBaseManual, .dIRRFAliquota, .iIRRFAliquotaManual, .dIRRFValor, _
                .iIRRFValorManual, .dPISRetido, .dCOFINSRetido, .dCSLLRetido, .iPISRetidoManual, .iCOFINSRetidoManual, .iCSLLRetidoManual, .dISSRetido, .iISSRetidoManual, .dPISCredito, .iPISCreditoManual, .dCOFINSCredito, .iCOFINSCreditoManual, .dIPICredito, .iIPICreditoManual, _
                .dICMSCredito, .iICMSCreditoManual, .dValorINSS, .iValorINSSManual, .iINSSRetido, .iINSSRetidoManual, .dINSSValorBase, .iINSSValorBaseManual, .dINSSValorDeducoes, .iINSSValorDeducoesManual, .dTotTrib, .iTotTribManual, .iTotTribTipo, .iTotTribTipoManual, .iNaturezaOpManual, _
                .dPISValor, .dCOFINSValor, .dICMSValorIsento, .dISSValorDeducao, .dISSValorOutrasRet, .dISSValorDescIncond, .dISSValorDescCond, .dIIValor, .iDestino, .iDestinoManual, _
                .iFinalidadeNFe, .iFinalidadeNFeManual, .iIndConsumidorFinal, .iIndConsumidorFinalManual, .iIndPresenca, .iIndPresencaManual, .iRegTribEspecial, .iRegTribEspecialManual, .dtDataPrestServico, .iDataPrestServicoManual, .dTotTribFed, .dTotTribEst, .dTotTribMunic, .iTotTribFedManual, .iTotTribEstManual, .iTotTribMunicManual, gobjCRFAT.iNFImportacaoTribFlag12, .dICMSInterestVlrFCPUFDest, .dICMSInterestVlrUFDest, .dICMSInterestVlrUFRemet, _
                .dICMSVlrFCP, .dICMSVlrFCPST, .dICMSVlrFCPSTRet, .dIPIVlrDevolvido)
    End With
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202860

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TributacaoDoc_Grava = SUCESSO

    Exit Function

Erro_TributacaoDoc_Grava:

    TributacaoDoc_Grava = gErr

    Select Case gErr

        Case 202859
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 202860
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRIBUTACAODOC", gErr, objTributacaoDoc.iTipoDoc, objTributacaoDoc.lNumIntDoc)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202861)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function TributacaoDocItem_Grava(ByVal objTributacaoDocItem As ClassTributacaoDocItem) As Long

Dim lErro As Long
Dim alComando(1 To 1) As Long
Dim iIndice As Integer

On Error GoTo Erro_TributacaoDocItem_Grava

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 202862
    Next

    With objTributacaoDocItem
        lErro = Comando_Executar(alComando(1), "INSERT INTO TributacaoDocItem( TipoDoc, NumIntDocItem, NumIntDoc, Item,Produto, ProdutoDescricao, UnidadeMed, Quantidade, PrecoTotal, PrecoUnitario, PrecoUnitarioMoeda, DescontoGrid, NaturezaOp, TipoTributacao, IPITipo, IPICodProduto, IPIBaseCalculo, IPIPercRedBase, IPIAliquota, IPIValor, IPICredito, ICMSTipo, ICMSBase, ICMSPercRedBase, ICMSAliquota, ICMSValor, ICMSCredito, ICMSSubstBase, ICMSSubstAliquota, ICMSSubstValor, PISCredito, COFINSCredito, ICMSAliquotaAdicaoDI, ICMSPercRedBaseAdicaoDI, DespImpICMSBase, DespImpICMSValor, DespImpICMSCredito, ICMSSubstPercRedBase, ICMSSubstPercMVA, PISTipo, PISTipoCalculo, PISBase, PISAliquota, PISAliquotaValor, PISQtde, PISValor, PISSTTipoCalculo, PISSTBase, PISSTAliquota, PISSTAliquotaValor, " & _
                "PISSTQtde, PISSTValor, COFINSTipo, COFINSTipoCalculo, COFINSBase, COFINSAliquota, COFINSAliquotaValor, COFINSQtde, COFINSValor, COFINSSTTipoCalculo, COFINSSTBase, COFINSSTAliquota, COFINSSTAliquotaValor, COFINSSTQtde, COFINSSTValor, CST, ISSQN, ISSBase, ISSAliquota, ISSCidadeIBGE, ISSValor, IIBase, IIDespAduaneira, IIIOF, IIValor, ICMSBaseModalidade, ICMSSubstBaseModalidade, IPIEnquadramentoClasse, IPIEnquadramentoCodigo, IPICNPJProdutor, IPISeloCodigo, IPISeloQtde, IPITipoCalculo, IPIUnidadePadraoQtde, IPIUnidadePadraoValor, ValorFreteItem, ValorSeguroItem, ValorOutrasDespesasItem, ValorDescontoItem, OrigemMercadoria, ExTIPI, Genero, ProdutoEspecifico, EAN, EANTrib, QtdTrib, UMTrib, ValorUnitTrib, NaturezaOpManual, TipoTributacaoManual, " & _
                "IPITipoManual, IPICodProdutoManual, IPIBaseCalculoManual, IPIPercRedBaseManual, IPIAliquotaManual, IPIValorManual, IPICreditoManual, ICMSTipoManual, ICMSBaseManual, ICMSPercRedBaseManual, ICMSAliquotaManual, ICMSValorManual, ICMSCreditoManual, ICMSSubstBaseManual, ICMSSubstAliquotaManual,ICMSSubstValorManual, PISCreditoManual, COFINSCreditoManual, ICMSAliquotaAdicaoDIManual, ICMSPercRedBaseAdicaoDIManual, DespImpICMSBaseManual, DespImpICMSValorManual, DespImpICMSCreditoManual, ICMSSubstPercRedBaseManual, ICMSSubstPercMVAManual, PISTipoManual, " & _
                "PISTipoCalculoManual, PISBaseManual, PISAliquotaManual, PISAliquotaValorManual, PISQtdeManual, PISValorManual, PISSTTipoCalculoManual, PISSTBaseManual, PISSTAliquotaManual, PISSTAliquotaValorManual, PISSTQtdeManual, PISSTValorManual, COFINSTipoManual, COFINSTipoCalculoManual, COFINSBaseManual, COFINSAliquotaManual, COFINSAliquotaValorManual, COFINSQtdeManual, COFINSValorManual, COFINSSTTipoCalculoManual, COFINSSTBaseManual, COFINSSTAliquotaManual, COFINSSTAliquotaValorManual, COFINSSTQtdeManual, COFINSSTValorManual, " & _
                "CSTManual, ISSQNManual, ISSBaseManual, ISSAliquotaManual, ISSCidadeIBGEManual, ISSValorManual, IIBaseManual, IIDespAduaneiraManual, IIIOFManual, IIValorManual, ICMSBaseModalidadeManual, ICMSSubstBaseModalidadeManual, IPIEnquadramentoClasseManual, IPIEnquadramentoCodigoManual, IPICNPJProdutorManual, IPISeloCodigoManual, IPISeloQtdeManual, IPITipoCalculoManual, IPIUnidadePadraoQtdeManual, IPIUnidadePadraoValorManual, ValorFreteItemManual, ValorSeguroItemManual, ValorOutrasDespesasItemManual, ValorDescontoItemManual, OrigemMercadoriaManual, FCI, FCIManual, ICMSSTCobrAntBase,ICMSSTCobrAntValor,ICMSSTCobrAntBaseManual,ICMSSTCobrAntValorManual, " & _
                "ICMSpercBaseOperacaoPropria,ICMSUFDevidoST,ICMSvBCSTRet,ICMSvICMSSTRet,ICMSvBCSTDest,ICMSvICMSSTDest,ICMSpCredSN,ICMSvCredSN,ISSCST,ICMSValorIsento,ICMSMotivo,ICMSSimplesTipo,ICMSpercBaseOperacaoPropriaManual,ICMSUFDevidoSTManual,ICMSvBCSTRetManual,ICMSvICMSSTRetManual,ICMSvBCSTDestManual,ICMSvICMSSTDestManual,ICMSpCredSNManual,ICMSvCredSNManual, ICMSValorIsentoManual, ICMSMotivoManual, ICMSSimplesTipoManual, ISSTipo, ISSTipoManual, RegimeTributario, CSOSN, NatBCCred, NatBCCredManual, TotTrib, TotTribManual, TotTribTipo, TotTribTipoManual, " & _
                "ICMS51ValorOp, ICMSPercDifer, ICMSValorDif, ICMS51ValorOpManual, ICMSPercDiferManual, ICMSValorDifManual, ISSValorDeducao, ISSValorOutrasRet, ISSValorDescIncond, ISSValorDescCond, ISSValorRet, ISSIndExigibilidade, ISSCodServ, ISSMunicIncidImp, ISSCodPais, ISSNumProcesso, ISSIndIncentivo, ISSValorDeducaoManual, ISSValorOutrasRetManual, ISSValorDescIncondManual , ISSValorDescCondManual, ISSValorRetManual, ISSIndExigibilidadeManual, ISSCodServManual, ISSMunicIncidImpManual, ISSCodPaisManual, ISSNumProcessoManual, ISSIndIncentivoManual,TotTribAliqFederal,TotTribAliqEstadual,TotTribAliqMunicipal,TotTribAliqFederalManual,TotTribAliqEstadualManual,TotTribAliqMunicipalManual, " & _
                "CEST,ICMSInterestBCUFDest,ICMSInterestBCUFDestManual,ICMSInterestPercFCPUFDest,ICMSInterestPercFCPUFDestManual,ICMSInterestAliqUFDest,ICMSInterestAliqUFDestManual,ICMSInterestPercPartilha,ICMSInterestPercPartilhaManual,ICMSInterestVlrFCPUFDest,ICMSInterestVlrFCPUFDestManual,ICMSInterestVlrUFDest,ICMSInterestVlrUFDestManual,ICMSInterestVlrUFRemet,ICMSInterestVlrUFRemetManual,ICMSInterestAliq,ICMSInterestAliqManual, " & _
                "ICMSInterestBCFCPUFDest,ICMSInterestBCFCPUFDestManual,ICMSvBCFCPST,ICMSvBCFCPSTManual,ICMSpFCPST,ICMSpFCPSTManual,ICMSvFCPST,ICMSvFCPSTManual,ICMSvBCFCPSTRet,ICMSvBCFCPSTRetManual,ICMSpFCPSTRet,ICMSpFCPSTRetManual,ICMSvFCPSTRet,ICMSvFCPSTRetManual,ICMSSTCobrAntAliquota,ICMSSTCobrAntAliquotaManual,ICMSvBCFCP,ICMSvBCFCPManual,ICMSpFCP,ICMSpFCPManual,ICMSvFCP,ICMSvFCPManual,cBenef,indEscala,CNPJFab,IPIVlrDevolvido,IPIVlrDevolvidoManual,pDevol,pDevolManual,ICMSSTBaseDupla,ICMSSTBaseDuplaManual,ICMSInterestBaseDupla,ICMSInterestBaseDuplaManual,cBenefManual) " & _
                "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                .iTipoDoc, .lNumIntDocItem, .lNumIntDoc, .iItem, .sProduto, .sProdutoDescricao, .sUnidadeMed, .dQuantidade, .dPrecoTotal, .dPrecoUnitario, .dPrecoUnitarioMoeda, .dDescontoGrid, .sNaturezaOp, .iTipoTributacao, .iIPITipo, .sIPICodProduto, .dIPIBaseCalculo, .dIPIPercRedBase, .dIPIAliquota, .dIPIValor, .dIPICredito, .iICMSTipo, .dICMSBase, .dICMSPercRedBase, .dICMSAliquota, .dICMSValor, .dICMSCredito, .dICMSSubstBase, .dICMSSubstAliquota, .dICMSSubstValor, .dPISCredito, .dCOFINSCredito, .dICMSAliquotaAdicaoDI, .dICMSPercRedBaseAdicaoDI, .dDespImpICMSBase, .dDespImpICMSValor, .dDespImpICMSCredito, .dICMSSubstPercRedBase, .dICMSSubstPercMVA, .iPISTipo, .iPISTipoCalculo, .dPISBase, .dPISAliquota, .dPISAliquotaValor, .dPISQtde, .dPISValor, .iPISSTTipoCalculo, .dPISSTBase, .dPISSTAliquota, .dPISSTAliquotaValor, _
                .dPISSTQtde, .dPISSTValor, .iCOFINSTipo, .iCOFINSTipoCalculo, .dCOFINSBase, .dCOFINSAliquota, .dCOFINSAliquotaValor, .dCOFINSQtde, .dCOFINSValor, .iCOFINSSTTipoCalculo, .dCOFINSSTBase, .dCOFINSSTAliquota, .dCOFINSSTAliquotaValor, .dCOFINSSTQtde, .dCOFINSSTValor, .sCST, .sISSQN, .dISSBase, .dISSAliquota, .sISSCidadeIBGE, .dISSValor, .dIIBase, .dIIDespAduaneira, .dIIIOF, .dIIValor, .iICMSBaseModalidade, .iICMSSubstBaseModalidade, .sIPIEnquadramentoClasse, .sIPIEnquadramentoCodigo, .sIPICNPJProdutor, .sIPISeloCodigo, .lIPISeloQtde, .iIPITipoCalculo, .dIPIUnidadePadraoQtde, .dIPIUnidadePadraoValor, .dValorFreteItem, .dValorSeguroItem, .dValorOutrasDespesasItem, .dValorDescontoItem, .iOrigemMercadoria, .iExTIPI, .sGenero, .iProdutoEspecifico, .sEAN, .sEANTrib, .dQtdTrib, .sUMTrib, .dValorUnitTrib, .iNaturezaOpManual, .iTipoTributacaoManual, _
                .iIPITipoManual, .iIPICodProdutoManual, .iIPIBaseCalculoManual, .iIPIPercRedBaseManual, .iIPIAliquotaManual, .iIPIValorManual, .iIPICreditoManual, .iICMSTipoManual, .iICMSBaseManual, .iICMSPercRedBaseManual, .iICMSAliquotaManual, .iICMSValorManual, .iICMSCreditoManual, .iICMSSubstBaseManual, .iICMSSubstAliquotaManual, .iICMSSubstValorManual, .iPISCreditoManual, .iCOFINSCreditoManual, .iICMSAliquotaAdicaoDIManual, .iICMSPercRedBaseAdicaoDIManual, _
                .iDespImpICMSBaseManual, .iDespImpICMSValorManual, .iDespImpICMSCreditoManual, .iICMSSubstPercRedBaseManual, .iICMSSubstPercMVAManual, .iPISTipoManual, .iPISTipoCalculoManual, .iPISBaseManual, .iPISAliquotaManual, .iPISAliquotaValorManual, .iPISQtdeManual, .iPISValorManual, .iPISSTTipoCalculoManual, .iPISSTBaseManual, .iPISSTAliquotaManual, .iPISSTAliquotaValorManual, .iPISSTQtdeManual, .iPISSTValorManual, .iCOFINSTipoManual, .iCOFINSTipoCalculoManual, .iCOFINSBaseManual, .iCOFINSAliquotaManual, .iCOFINSAliquotaValorManual, .iCOFINSQtdeManual, .iCOFINSValorManual, .iCOFINSSTTipoCalculoManual, .iCOFINSSTBaseManual, .iCOFINSSTAliquotaManual, .iCOFINSSTAliquotaValorManual, _
                .iCOFINSSTQtdeManual, .iCOFINSSTValorManual, .iCSTManual, .iISSQNManual, .iISSBaseManual, .iISSAliquotaManual, .iISSCidadeIBGEManual, .iISSValorManual, .iIIBaseManual, .iIIDespAduaneiraManual, .iIIIOFManual, .iIIValorManual, .iICMSBaseModalidadeManual, .iICMSSubstBaseModalidadeManual, .iIPIEnquadramentoClasseManual, .iIPIEnquadramentoCodigoManual, .iIPICNPJProdutorManual, .iIPISeloCodigoManual, .iIPISeloQtdeManual, .iIPITipoCalculoManual, .iIPIUnidadePadraoQtdeManual, .iIPIUnidadePadraoValorManual, .iValorFreteItemManual, .iValorSeguroItemManual, .iValorOutrasDespesasItemManual, .iValorDescontoItemManual, .iOrigemMercadoriaManual, .sFCI, .iFCIManual, _
                .dICMSSTCobrAntBase, .dICMSSTCobrAntValor, .iICMSSTCobrAntBaseManual, .iICMSSTCobrAntValorManual, .dICMSpercBaseOperacaoPropria, .sICMSUFDevidoST, .dICMSvBCSTRet, .dICMSvICMSSTRet, .dICMSvBCSTDest, .dICMSvICMSSTDest, .dICMSpCredSN, .dICMSvCredSN, .sISSCST, .dICMSValorIsento, .iICMSMotivo, .iICMSSimplesTipo, .iICMSpercBaseOperacaoPropriaManual, .iICMSUFDevidoSTManual, .iICMSvBCSTRetManual, .iICMSvICMSSTRetManual, .iICMSvBCSTDestManual, .iICMSvICMSSTDestManual, .iICMSpCredSNManual, .iICMSvCredSNManual, .iICMSValorIsentoManual, .iICMSMotivoManual, .iICMSSimplesTipoManual, .iISSTipo, .iISSTipoManual, .iRegimeTributario, .sCSOSN, .sNatBCCred, .iNatBCCredManual, .dTotTrib, .iTotTribManual, .iTotTribTipo, .iTotTribTipoManual, _
                .dICMS51ValorOp, .dICMSPercDifer, .dICMSValorDif, .iICMS51ValorOpManual, .iICMSPercDiferManual, .iICMSValorDifManual, .dISSValorDeducao, .dISSValorOutrasRet, .dISSValorDescIncond, .dISSValorDescCond, .dISSValorRet, .iISSIndExigibilidade, .sISSCodServ, .sISSMunicIncidImp, .iISSCodPais, .sISSNumProcesso, .iISSIndIncentivo, .iISSValorDeducaoManual, .iISSValorOutrasRetManual, .iISSValorDescIncondManual, .iISSValorDescCondManual, .iISSValorRetManual, .iISSIndExigibilidadeManual, .iISSCodServManual, .iISSMunicIncidImpManual, .iISSCodPaisManual, .iISSNumProcessoManual, .iISSIndIncentivoManual, .dTotTribAliqFed, .dTotTribAliqEst, .dTotTribAliqMunic, .iTotTribAliqFedManual, .iTotTribAliqEstManual, .iTotTribAliqMunicManual, _
                .sCEST, .dICMSInterestBCUFDest, .iICMSInterestBCUFDestManual, .dICMSInterestPercFCPUFDest, .iICMSInterestPercFCPUFDestManual, .dICMSInterestAliqUFDest, .iICMSInterestAliqUFDestManual, .dICMSInterestPercPartilha, .iICMSInterestPercPartilhaManual, .dICMSInterestVlrFCPUFDest, .iICMSInterestVlrFCPUFDestManual, .dICMSInterestVlrUFDest, .iICMSInterestVlrUFDestManual, .dICMSInterestVlrUFRemet, .iICMSInterestVlrUFRemetManual, .dICMSInterestAliq, .iICMSInterestAliqManual, _
                .dICMSInterestBCFCPUFDest, .iICMSInterestBCFCPUFDestManual, .dICMSvBCFCPST, .iICMSvBCFCPSTManual, .dICMSpFCPST, .iICMSpFCPSTManual, .dICMSvFCPST, .iICMSvFCPSTManual, .dICMSvBCFCPSTRet, .iICMSvBCFCPSTRetManual, .dICMSpFCPSTRet, .iICMSpFCPSTRetManual, .dICMSvFCPSTRet, .iICMSvFCPSTRetManual, .dICMSSTCobrAntAliquota, .iICMSSTCobrAntAliquotaManual, .dICMSvBCFCP, .iICMSvBCFCPManual, .dICMSpFCP, .iICMSpFCPManual, .dICMSvFCP, .iICMSvFCPManual, .scBenef, .sindEscala, .sCNPJFab, .dIPIVlrDevolvido, .iIPIVlrDevolvidoManual, .dpDevol, .ipDevolManual, .iICMSSTBaseDupla, .iICMSSTBaseDuplaManual, .iICMSInterestBaseDupla, .iICMSInterestBaseDuplaManual, .icBenefManual)
    End With
    If lErro <> AD_SQL_SUCESSO Then gError 202863

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TributacaoDocItem_Grava = SUCESSO

    Exit Function

Erro_TributacaoDocItem_Grava:

    TributacaoDocItem_Grava = gErr

    Select Case gErr

        Case 202862
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 202863
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRIBUTACAODOCITEM", gErr, objTributacaoDocItem.iTipoDoc, objTributacaoDocItem.lNumIntDocItem)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202864)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function TributacaoDoc_Exclui(ByVal iTipoDoc As Integer, ByVal lNumIntDoc As Long, Optional ByVal bEvitaErroNaoCad As Boolean = False) As Long

Dim lErro As Long
Dim alComando(0 To 3) As Long
Dim iIndice As Integer
Dim lNumIntDocAux As Long

On Error GoTo Erro_TributacaoDoc_Exclui

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 202865
    Next

    'Le a tabelaTributacaoDoc
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM TributacaoDoc WHERE TipoDoc= ?  AND NumIntDoc= ? ", _
                0, lNumIntDocAux, iTipoDoc, lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 202866

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202867

    'Se não existir => ERRO
    If lErro = AD_SQL_SEM_DADOS And Not bEvitaErroNaoCad Then gError 202868

    If lErro <> AD_SQL_SEM_DADOS Then

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 202869
    
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM TributacaoDoc", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 202870
    
        'Le a tabela TributacaoDocItem
        lErro = Comando_ExecutarPos(alComando(2), "SELECT NumIntDocItem FROM TributacaoDocItem WHERE TipoDoc= ?  AND NumIntDoc= ? ", _
                    0, lNumIntDocAux, iTipoDoc, lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 202871
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202872
        
        Do While lErro <> AD_SQL_SEM_DADOS
        
             'Trava registro contra alterações/Leituras
            lErro = Comando_LockExclusive(alComando(2))
            If lErro <> SUCESSO Then gError 202873
        
            lErro = Comando_ExecutarPos(alComando(3), "DELETE FROM TributacaoDocItem", alComando(2))
            If lErro <> AD_SQL_SUCESSO Then gError 202874
        
            'Busca Próximo
            lErro = Comando_BuscarPrimeiro(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202875
        
        Loop
        
    End If

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TributacaoDoc_Exclui = SUCESSO

    Exit Function

Erro_TributacaoDoc_Exclui:

    TributacaoDoc_Exclui = gErr

    Select Case gErr

        Case 202865
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 202866, 202867
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRIBUTACAODOC", gErr, iTipoDoc, lNumIntDoc)

        Case 202868
            Call Rotina_Erro(vbOKOnly, "ERRO_TRIBUTACAODOC_NAO_CADASTRADO", gErr, iTipoDoc, lNumIntDoc)

        Case 202869
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRIBUTACAODOC", gErr, iTipoDoc, lNumIntDoc)

        Case 202870
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRIBUTACAODOC", gErr, iTipoDoc, lNumIntDoc)
            
        Case 202871, 202872, 202875
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRIBUTACAODOCITEM", gErr, iTipoDoc, lNumIntDocAux)

        Case 202873
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRIBUTACAODOCITEM", gErr, iTipoDoc, lNumIntDocAux)

        Case 202874
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRIBUTACAODOCITEM", gErr, iTipoDoc, lNumIntDocAux)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202876)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function TributacaoDocItem_Exclui(ByVal iTipoDoc As Integer, ByVal lNumIntDocItem As Long) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lNumIntDocAux As Long

On Error GoTo Erro_TributacaoDocItem_Exclui

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 202877
    Next

    'Le a tabelaTributacaoDoc
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM TributacaoDocItem WHERE TipoDoc= ?  AND NumIntDocItem= ? ", _
                0, lNumIntDocAux, iTipoDoc, lNumIntDocItem)
    If lErro <> AD_SQL_SUCESSO Then gError 202878

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202879

    If lErro <> AD_SQL_SEM_DADOS Then

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 202881
    
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM TributacaoDocItem", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 202882

    End If

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TributacaoDocItem_Exclui = SUCESSO

    Exit Function

Erro_TributacaoDocItem_Exclui:

    TributacaoDocItem_Exclui = gErr

    Select Case gErr

        Case 202877
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 202878, 202879
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRIBUTACAODOCITEM", gErr, iTipoDoc, lNumIntDocItem)

        Case 202880
            Call Rotina_Erro(vbOKOnly, "ERRO_TRIBUTACAODOCITEM_NAO_CADASTRADO", gErr, iTipoDoc, lNumIntDocItem)

        Case 202881
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRIBUTACAODOCITEM", gErr, iTipoDoc, lNumIntDocItem)

        Case 202882
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRIBUTACAODOCITEM", gErr, iTipoDoc, lNumIntDocItem)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202883)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function Tributacao_Atualiza_Versao() As Long
'Verifica se a tributação precisa ser atualizada e se sim chama a função para tal

Dim lErro As Long
Dim lComando1 As Long
Dim lComando2 As Long
Dim lTransacao As Long
Dim sVersao As String
Dim iIndice As Integer

Const VERSAO_TRIB_ATUAL = 2

On Error GoTo Erro_Tributacao_Atualiza_Versao

    'Abre o comando
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 202890

    sVersao = String(STRING_CONFIG_CONTEUDO, 0)

    'Lê a versão usada atualmente na tributação
    lErro = Comando_Executar(lComando1, "SELECT Conteudo FROM AdmConfig WHERE Codigo = ? ", sVersao, "VERSAO_TRIBUTACAO")
    If lErro <> AD_SQL_SUCESSO Then gError 202893

    'Busca o Primeiro
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202894
    
    'Se tem versão
    If lErro <> AD_SQL_SEM_DADOS Then
    
        'Percorre as versões até chegar a versão atual
        For iIndice = StrParaInt(sVersao) + 1 To VERSAO_TRIB_ATUAL
                
            'MsgBox ("Vai atualizar a versao da tributacao: de " & CStr(StrParaInt(sVersao) + 1) & " (" & CStr(iIndice) & ") para " & sVersao)
                
            lTransacao = Transacao_Abrir()
            If lTransacao = 0 Then gError 202892
                
            Call Conexao_DesativarLocks(DESATIVAR_LOCKS)
            
            lComando2 = Comando_Abrir()
            If lComando2 = 0 Then gError 202891
                        
            'manda atualizar a versão
            lErro = CF("Tributacao_Atualiza_Versao" & CStr(iIndice))
            If lErro <> SUCESSO Then gError 202895
                       
'            'Marca a nova versão no BD
'            lErro = Comando_Executar(lComando2, "UPDATE Controle SET Conteudo = ? WHERE Codigo = ? ", CStr(iIndice), CONTROLE_VERSAO_TRIB)
            lErro = Comando_Executar(lComando2, "UPDATE AdmConfig SET Conteudo = ? WHERE Codigo = ? ", CStr(iIndice), "VERSAO_TRIBUTACAO")
            If lErro <> AD_SQL_SUCESSO Then gError 202896
                       
            Call Comando_Fechar(lComando2)
            
            Call Conexao_DesativarLocks(REATIVAR_LOCKS)
            
            lErro = Transacao_Commit()
            If lErro <> AD_SQL_SUCESSO Then gError 202897
                       
        Next
        
    End If
       
    Call Comando_Fechar(lComando1)
    
    Tributacao_Atualiza_Versao = SUCESSO
    
    Exit Function

Erro_Tributacao_Atualiza_Versao:

    Tributacao_Atualiza_Versao = gErr

    Select Case gErr
   
        Case 202890, 202891
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 202892
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
       
        Case 202893, 202894
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTROLE", gErr)
            
        Case 202895
            
        Case 202896
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_CONTROLE", gErr)
        
        Case 202897
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202898)
        
    End Select
    
    Call Transacao_Rollback
   
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

    Exit Function

End Function

Public Function Tributacao_Atualiza_Versao2() As Long
'Essa função atualiza a tributação da versão 1 para versão 2 (TributacaoNF->TributacaoDoc)
'Para Reatualizar a versão é necessário:
'1 - Voltar a versão usada na tabela Controle do Dic
'(DIC) UPDATE Controle SET Conteudo = 1 WHERE Codigo = 2 (Volta para versão 1)
'2 - Marcar o Campo Atualizado na tabela desejaca como = 0
'(DADOS) UPDATE TRIB_OLD_TributacaoNF SET Atualizado = 0 (Marca para reatualizar as NFs)

Dim lErro As Long, iIndice As Integer, iTipoDoc As Integer, iTipoChave As Integer
Dim sSQLTribDoc As String, sSQLTribCompl As String, sSQLTribDocItem As String, sSQLDoc As String, sSQLUpdateDoc As String
Const TIPO_CHAVE_NUMINTDOC = 0
Const TIPO_CHAVE_COD_FILIAL = 1
Const TIPO_CHAVE_COD_FILIAL_STRING = 2

On Error GoTo Erro_Tributacao_Atualiza_Versao2

    'Para cada conjunto de tabelas a serem convertidas
    For iIndice = 1 To 4
    
        'Monta as leituras adequadas para passar para função de leitura e transformação
        Select Case iIndice
        
            Case 1 'NF
                iTipoDoc = TIPODOC_TRIB_NF
                iTipoChave = TIPO_CHAVE_NUMINTDOC
                sSQLTribDoc = "SELECT 0 AS FilialEmpresa, 0 AS TaxacaoAutomatica, NumIntDoc,NaturezaOpInterna,TipoTributacao,IPIBase,IPIValor,IPICredito,ICMSBase,ICMSValor,ICMSSubstBase,ICMSSubstValor,ICMSCredito,ISSIncluso,ISSBase,ISSAliquota,ISSValor,IRRFBase,IRRFAliquota,IRRFValor,ValorINSS,INSSRetido,INSSValorBase,INSSValorDeducoes,PISCredito,COFINSCredito,PISRetido,COFINSRetido,CSLLRetido,ISSRetido FROM TRIB_OLD_TributacaoNF WHERE Atualizado = 0 "
                sSQLTribCompl = "SELECT Tipo,IPIBaseCalculo,IPIValor,IPICredito,ICMSBase,ICMSValor,ICMSCredito,ICMSSubstBase,ICMSSubstValor,PISCredito,COFINSCredito FROM TRIB_OLD_TributacaoComplNF WHERE NumIntDoc = ? AND 0 = ? "
                sSQLTribDocItem = "SELECT T.NumIntNF,T.Item,T.NumIntDoc,T.NaturezaOp,T.TipoTributacao,T.IPITipo,T.IPICodProduto,T.IPIBaseCalculo,T.IPIPercRedBase,T.IPIAliquota,T.IPIValor,T.IPICredito,T.ICMSTipo,T.ICMSBase,T.ICMSPercRedBase,T.ICMSAliquota,T.ICMSValor,T.ICMSCredito,T.ICMSSubstBase,T.ICMSSubstAliquota,T.ICMSSubstValor,T.PISCredito,T.COFINSCredito,T.ICMSAliquotaAdicaoDI,T.DespImpICMSBase,T.DespImpICMSValor,T.DespImpICMSCredito,T.ICMSSubstPercRedBase,T.ICMSSubstPercMVA, " & _
                                                    "I.Produto , I.DescricaoItem, I.UnidadeMed, I.Quantidade, I.Quantidade * I.PrecoUnitario - I.ValorDesconto, I.PrecoUnitario, I.PrecoUnitarioMoeda, I.ValorDesconto, P.OrigemMercadoria " & _
                                                    "FROM TRIB_OLD_TributacaoItemNF AS T, ItensNFiscal AS I, Produtos AS P WHERE T.NumIntDoc = I.NumIntDoc AND P.Codigo = I.Produto AND I.NumIntNF = ? ORDER By I.Item "
                sSQLDoc = "SELECT NumIntDoc, NaturezaOP, ValorTotal, ValorFrete, ValorSeguro, ValorDesconto, ValorOutrasDespesas FROM NFiscal WHERE NumIntDoc = ? AND 0 = ?"
                sSQLUpdateDoc = "UPDATE TRIB_OLD_TributacaoNF SET Atualizado = ? WHERE NumIntDoc = ? AND 0 = ? "
                
            Case 2 'PV - Aberto
                iTipoDoc = TIPODOC_TRIB_PV
                iTipoChave = TIPO_CHAVE_COD_FILIAL
                sSQLTribDoc = "SELECT FilialEmpresa, TaxacaoAutomatica, CodPedido As NumIntDoc,'' As NaturezaOpInterna, TipoTributacao,IPIBase,IPIValor,IPICredito,ICMSBase,ICMSValor,ICMSSubstBase,ICMSSubstValor,ICMSCredito,ISSIncluso,ISSBase,ISSAliquota,ISSValor,IRRFBase,IRRFAliquota,IRRFValor,ValorINSS,INSSRetido,INSSValorBase,INSSValorDeducoes,PISCredito,COFINSCredito,PISRetido,COFINSRetido,CSLLRetido,ISSRetido FROM TRIB_OLD_TributacaoPV WHERE Atualizado = 0 "
                sSQLTribCompl = "SELECT Item AS Tipo,IPIBaseCalculo,IPIValor,IPICredito,ICMSBase,ICMSValor,ICMSCredito,ICMSSubstBase,ICMSSubstValor,PISCredito,COFINSCredito FROM TRIB_OLD_TributacaoComplPV WHERE CodPedido = ? AND FilialEmpresa = ? "
                sSQLTribDocItem = "SELECT V.NumIntDoc,0 AS Item,I.NumIntDoc,T.NaturezaOp,T.TipoTributacao,T.IPITipo,P.IPICodigo,T.IPIBaseCalculo,T.IPIPercRedBase,T.IPIAliquota,T.IPIValor,T.IPICredito,T.ICMSTipo,T.ICMSBase,T.ICMSPercRedBase,T.ICMSAliquota,T.ICMSValor,T.ICMSCredito,T.ICMSSubstBase,T.ICMSSubstAliquota,T.ICMSSubstValor,T.PISCredito,T.COFINSCredito,0 AS ICMSAliquotaAdicaoDI,0 AS DespImpICMSBase,0 AS DespImpICMSValor,0 AS DespImpICMSCredito,T.ICMSSubstPercRedBase,T.ICMSSubstPercMVA,  " & _
                                                    "I.Produto , I.Descricao, I.UnidadeMed, I.Quantidade, I.Quantidade * I.PrecoUnitario - I.ValorDesconto, I.PrecoUnitario, 0 AS PrecoUnitarioMoeda, I.ValorDesconto, P.OrigemMercadoria " & _
                                                    "FROM TRIB_OLD_TributacaoItemPV AS T, ItensPedidoDeVendaTodos AS I, Produtos AS P, PedVenTodos AS V WHERE T.NumIntDoc = I.NumIntDoc AND V.Codigo = I.CodPedido AND V.FilialEmpresa = I.FilialEmpresa AND P.Codigo = I.Produto AND V.NumIntDoc = ? ORDER By I.NumIntDoc "
                sSQLDoc = "SELECT NumIntDoc, NaturezaOP, ValorTotal, ValorFrete, ValorSeguro, ValorDesconto, ValorOutrasDespesas FROM PedVenTodos WHERE Codigo = ? AND FilialEmpresa = ? "
                sSQLUpdateDoc = "UPDATE TRIB_OLD_TributacaoPV SET Atualizado = ? WHERE CodPedido = ? AND FilialEmpresa = ? "
            
            Case 3 'PV - Baixados
                iTipoDoc = TIPODOC_TRIB_PV
                iTipoChave = TIPO_CHAVE_COD_FILIAL
                sSQLTribDoc = "SELECT FilialEmpresa, TaxacaoAutomatica, CodPedido As NumIntDoc,'' As NaturezaOpInterna, TipoTributacao,IPIBase,IPIValor,IPICredito,ICMSBase,ICMSValor,ICMSSubstBase,ICMSSubstValor,ICMSCredito,ISSIncluso,ISSBase,ISSAliquota,ISSValor,IRRFBase,IRRFAliquota,IRRFValor,ValorINSS,INSSRetido,INSSValorBase,INSSValorDeducoes,PISCredito,COFINSCredito,PISRetido,COFINSRetido,CSLLRetido,ISSRetido FROM TRIB_OLD_TributacaoPVBaixado WHERE Atualizado = 0 "
                sSQLTribCompl = "SELECT Item AS Tipo,IPIBaseCalculo,IPIValor,IPICredito,ICMSBase,ICMSValor,ICMSCredito,ICMSSubstBase,ICMSSubstValor,PISCredito,COFINSCredito FROM TRIB_OLD_TributacaoComplPVBaixado WHERE CodPedido = ? AND FilialEmpresa = ? "
                sSQLTribDocItem = "SELECT V.NumIntDoc,0 AS Item,I.NumIntDoc,T.NaturezaOp,T.TipoTributacao,T.IPITipo,P.IPICodigo,T.IPIBaseCalculo,T.IPIPercRedBase,T.IPIAliquota,T.IPIValor,T.IPICredito,T.ICMSTipo,T.ICMSBase,T.ICMSPercRedBase,T.ICMSAliquota,T.ICMSValor,T.ICMSCredito,T.ICMSSubstBase,T.ICMSSubstAliquota,T.ICMSSubstValor,T.PISCredito,T.COFINSCredito,0 AS ICMSAliquotaAdicaoDI,0 AS DespImpICMSBase,0 AS DespImpICMSValor,0 AS DespImpICMSCredito,T.ICMSSubstPercRedBase,T.ICMSSubstPercMVA,  " & _
                                                    "I.Produto , I.Descricao, I.UnidadeMed, I.Quantidade, I.Quantidade * I.PrecoUnitario - I.ValorDesconto, I.PrecoUnitario, 0 AS PrecoUnitarioMoeda, I.ValorDesconto, P.OrigemMercadoria " & _
                                                    "FROM TRIB_OLD_TributacaoItemPVBaixado AS T, ItensPedidoDeVendaTodos AS I, Produtos AS P, PedVenTodos AS V WHERE T.NumIntDoc = I.NumIntDoc AND V.Codigo = I.CodPedido AND V.FilialEmpresa = I.FilialEmpresa AND P.Codigo = I.Produto AND V.NumIntDoc = ? ORDER By I.NumIntDoc "
                sSQLDoc = "SELECT NumIntDoc, NaturezaOP, ValorTotal, ValorFrete, ValorSeguro, ValorDesconto, ValorOutrasDespesas FROM PedVenTodos WHERE Codigo = ? AND FilialEmpresa = ? "
                sSQLUpdateDoc = "UPDATE TRIB_OLD_TributacaoPVBaixado SET Atualizado = ? WHERE CodPedido = ? AND FilialEmpresa = ?"
                
            Case 4 'OV
                iTipoDoc = TIPODOC_TRIB_OV
                iTipoChave = TIPO_CHAVE_COD_FILIAL
                sSQLTribDoc = "SELECT FilialEmpresa, TaxacaoAutomatica, CodOrcamento As NumIntDoc,'' As NaturezaOpInterna, TipoTributacao,IPIBase,IPIValor,IPICredito,ICMSBase,ICMSValor,ICMSSubstBase,ICMSSubstValor,ICMSCredito,ISSIncluso,ISSBase,ISSAliquota,ISSValor,IRRFBase,IRRFAliquota,IRRFValor,ValorINSS,INSSRetido,INSSValorBase,INSSValorDeducoes,PISCredito,COFINSCredito,PISRetido,COFINSRetido,CSLLRetido,ISSRetido FROM TRIB_OLD_TributacaoOV WHERE Atualizado = 0 "
                sSQLTribCompl = "SELECT Item AS Tipo,IPIBaseCalculo,IPIValor,IPICredito,ICMSBase,ICMSValor,ICMSCredito,ICMSSubstBase,ICMSSubstValor,PISCredito,COFINSCredito FROM TRIB_OLD_TributacaoComplOV WHERE CodOrcamento = ? AND FilialEmpresa = ? "
                sSQLTribDocItem = "SELECT V.NumIntDoc,0 AS Item,I.NumIntDoc,T.NaturezaOp,T.TipoTributacao,T.IPITipo,P.IPICodigo,T.IPIBaseCalculo,T.IPIPercRedBase,T.IPIAliquota,T.IPIValor,T.IPICredito,T.ICMSTipo,T.ICMSBase,T.ICMSPercRedBase,T.ICMSAliquota,T.ICMSValor,T.ICMSCredito,T.ICMSSubstBase,T.ICMSSubstAliquota,T.ICMSSubstValor,T.PISCredito,T.COFINSCredito,0 AS ICMSAliquotaAdicaoDI,0 AS DespImpICMSBase,0 AS DespImpICMSValor,0 AS DespImpICMSCredito,T.ICMSSubstPercRedBase,T.ICMSSubstPercMVA, " & _
                                                    "I.Produto , I.Descricao, I.UnidadeMed, I.Quantidade, I.Quantidade * I.PrecoUnitario - I.ValorDesconto, I.PrecoUnitario, 0 AS PrecoUnitarioMoeda, I.ValorDesconto, P.OrigemMercadoria " & _
                                                    "FROM TRIB_OLD_TributacaoItemOV AS T, ItensOrcamentoVenda AS I, Produtos AS P, OrcamentoVenda AS V WHERE T.NumIntDoc = I.NumIntDoc AND V.Codigo = I.CodOrcamento AND V.FilialEmpresa = I.FilialEmpresa AND P.Codigo = I.Produto AND V.NumIntDoc = ? ORDER By I.NumIntDoc "
                sSQLDoc = "SELECT NumIntDoc, NaturezaOP, ValorTotal, ValorFrete, ValorSeguro, ValorDesconto, ValorOutrasDespesas FROM OrcamentoVenda WHERE Codigo = ? AND FilialEmpresa = ? "
                sSQLUpdateDoc = "UPDATE TRIB_OLD_TributacaoOV SET Atualizado = ? WHERE CodOrcamento = ? AND FilialEmpresa = ?"
            
'''            Case 5 'OSRV
'''                iTipoDoc = TIPODOC_TRIB_OSRV
'''                iTipoChave = TIPO_CHAVE_COD_FILIAL
'''                sSQLTribDoc = "SELECT  FilialEmpresa, TaxacaoAutomatica, CodOrcamento As NumIntDoc,'' As NaturezaOpInterna, TipoTributacao,IPIBase,IPIValor,IPICredito,ICMSBase,ICMSValor,ICMSSubstBase,ICMSSubstValor,ICMSCredito,ISSIncluso,ISSBase,ISSAliquota,ISSValor,IRRFBase,IRRFAliquota,IRRFValor,ValorINSS,INSSRetido,INSSValorBase,INSSValorDeducoes,PISCredito,COFINSCredito,PISRetido,COFINSRetido,CSLLRetido,0 AS ISSRetido FROM TRIB_OLD_TributacaoOrcSRV WHERE Atualizado = 0 "
'''                sSQLTribCompl = "SELECT Item AS Tipo,IPIBaseCalculo,IPIValor,IPICredito,ICMSBase,ICMSValor,ICMSCredito,ICMSSubstBase,ICMSSubstValor,PISCredito,COFINSCredito FROM TRIB_OLD_TributacaoComplOrcSRV WHERE CodOrcamento = ? AND FilialEmpresa = ? "
'''                sSQLTribDocItem = "SELECT V.NumIntDoc,0 AS Item,I.NumIntDoc,T.NaturezaOp,T.TipoTributacao,T.IPITipo,P.IPICodigo,T.IPIBaseCalculo,T.IPIPercRedBase,T.IPIAliquota,T.IPIValor,T.IPICredito,T.ICMSTipo,T.ICMSBase,T.ICMSPercRedBase,T.ICMSAliquota,T.ICMSValor,T.ICMSCredito,T.ICMSSubstBase,T.ICMSSubstAliquota,T.ICMSSubstValor,T.PISCredito,T.COFINSCredito,0 AS ICMSAliquotaAdicaoDI,0 AS DespImpICMSBase,0 AS DespImpICMSValor,0 AS DespImpICMSCredito,0 AS ICMSSubstPercRedBase,0 AS ICMSSubstPercMVA, " & _
'''                                                    "I.Produto , I.Descricao, I.UnidadeMed, I.Quantidade, I.Quantidade * I.PrecoUnitario - I.ValorDesconto, I.PrecoUnitario, 0 AS PrecoUnitarioMoeda, I.ValorDesconto, P.OrigemMercadoria  " & _
'''                                                    "FROM TRIB_OLD_TributacaoItemOrcSRV AS T, ItensOrcamentoSRV AS I, Produtos AS P, OrcamentoSRV AS V WHERE T.NumIntDoc = I.NumIntDoc AND V.Codigo = I.CodOrcamento AND V.FilialEmpresa = I.FilialEmpresa AND P.Codigo = I.Produto AND V.NumIntDoc = ? ORDER By I.NumIntDoc "
'''                sSQLDoc = "SELECT NumIntDoc, NaturezaOP, ValorTotal, ValorFrete, ValorSeguro, ValorDesconto, ValorOutrasDespesas FROM OrcamentoSRV WHERE Codigo = ? AND FilialEmpresa = ? "
'''                sSQLUpdateDoc = "UPDATE TRIB_OLD_TributacaoOrcSRV SET Atualizado = ? WHERE CodOrcamento = ? AND FilialEmpresa = ?"
'''
'''            Case 6 'PSRV
'''                iTipoDoc = TIPODOC_TRIB_PSRV
'''                iTipoChave = TIPO_CHAVE_COD_FILIAL
'''                sSQLTribDoc = "SELECT FilialEmpresa, TaxacaoAutomatica, CodPedido As NumIntDoc,'' As NaturezaOpInterna, TipoTributacao,IPIBase,IPIValor,IPICredito,ICMSBase,ICMSValor,ICMSSubstBase,ICMSSubstValor,ICMSCredito,ISSIncluso,ISSBase,ISSAliquota,ISSValor,IRRFBase,IRRFAliquota,IRRFValor,ValorINSS,INSSRetido,INSSValorBase,INSSValorDeducoes,PISCredito,COFINSCredito,PISRetido,COFINSRetido,CSLLRetido, ISSRetido FROM TRIB_OLD_TributacaoPedSRV WHERE Atualizado = 0 "
'''                sSQLTribCompl = "SELECT Item AS Tipo,IPIBaseCalculo,IPIValor,IPICredito,ICMSBase,ICMSValor,ICMSCredito,ICMSSubstBase,ICMSSubstValor,PISCredito,COFINSCredito FROM TRIB_OLD_TributacaoComplPSRV WHERE CodPedido = ? AND FilialEmpresa = ? "
'''                sSQLTribDocItem = "SELECT V.NumIntDoc,0 AS Item,I.NumIntDoc,T.NaturezaOp,T.TipoTributacao,T.IPITipo,P.IPICodigo,T.IPIBaseCalculo,T.IPIPercRedBase,T.IPIAliquota,T.IPIValor,T.IPICredito,T.ICMSTipo,T.ICMSBase,T.ICMSPercRedBase,T.ICMSAliquota,T.ICMSValor,T.ICMSCredito,T.ICMSSubstBase,T.ICMSSubstAliquota,T.ICMSSubstValor,T.PISCredito,T.COFINSCredito,0 AS ICMSAliquotaAdicaoDI,0 AS DespImpICMSBase,0 AS DespImpICMSValor,0 AS DespImpICMSCredito,0 AS ICMSSubstPercRedBase,0 AS ICMSSubstPercMVA, " & _
'''                                                    "I.Produto , I.Descricao, I.UnidadeMed, I.Quantidade, I.Quantidade * I.PrecoUnitario - I.ValorDesconto, I.PrecoUnitario, 0 AS PrecoUnitarioMoeda, I.ValorDesconto, P.OrigemMercadoria " & _
'''                                                    "FROM TRIB_OLD_TributacaoItemPedidoSRV AS T, ItensPedidoSRV AS I, Produtos AS P, PedidoServico AS V WHERE T.NumIntDoc = I.NumIntDoc AND V.Codigo = I.CodPedido AND V.FilialEmpresa = I.FilialEmpresa AND P.Codigo = I.Produto AND V.NumIntDoc = ? ORDER By I.NumIntDoc "
'''                sSQLDoc = "SELECT NumIntDoc, NaturezaOP, ValorTotal, ValorFrete, ValorSeguro, ValorDesconto, ValorOutrasDespesas FROM PedidoServico WHERE Codigo = ? AND FilialEmpresa = ? "
'''                sSQLUpdateDoc = "UPDATE TRIB_OLD_TributacaoPedSRV SET Atualizado = ? WHERE CodPedido = ? AND FilialEmpresa = ?"
'''
'''            Case 7 'PRJ PROP
'''                iTipoDoc = TIPODOC_TRIB_PRJ_PROP
'''                iTipoChave = TIPO_CHAVE_COD_FILIAL_STRING
'''                sSQLTribDoc = "SELECT FilialEmpresa, TaxacaoAutomatica, CodProposta As NumIntDoc,'' As NaturezaOpInterna, TipoTributacao,IPIBase,IPIValor,IPICredito,ICMSBase,ICMSValor,ICMSSubstBase,ICMSSubstValor,ICMSCredito,ISSIncluso,ISSBase,ISSAliquota,ISSValor,IRRFBase,IRRFAliquota,IRRFValor,ValorINSS,INSSRetido,INSSValorBase,INSSValorDeducoes,PISCredito,COFINSCredito,PISRetido,COFINSRetido,CSLLRetido, ISSRetido FROM TRIB_OLD_TributacaoPRJProp WHERE Atualizado = 0 "
'''                sSQLTribCompl = "SELECT Item AS Tipo,IPIBaseCalculo,IPIValor,IPICredito,ICMSBase,ICMSValor,ICMSCredito,ICMSSubstBase,ICMSSubstValor,PISCredito,COFINSCredito FROM TRIB_OLD_TributacaoComplPRJProp WHERE CodProposta = ? AND FilialEmpresa = ? "
'''                sSQLTribDocItem = "SELECT V.NumIntDoc,0 AS Item,I.NumIntDoc,T.NaturezaOp,T.TipoTributacao,T.IPITipo,P.IPICodigo,T.IPIBaseCalculo,T.IPIPercRedBase,T.IPIAliquota,T.IPIValor,T.IPICredito,T.ICMSTipo,T.ICMSBase,T.ICMSPercRedBase,T.ICMSAliquota,T.ICMSValor,T.ICMSCredito,T.ICMSSubstBase,T.ICMSSubstAliquota,T.ICMSSubstValor,T.PISCredito,T.COFINSCredito,0 AS ICMSAliquotaAdicaoDI,0 AS DespImpICMSBase,0 AS DespImpICMSValor,0 AS DespImpICMSCredito,0 AS ICMSSubstPercRedBase,0 AS ICMSSubstPercMVA, " & _
'''                                                    "I.Produto , I.DescProd, I.UnidadeMed, I.Quantidade, I.Quantidade * I.PrecoUnitario - I.ValorDesconto, I.PrecoUnitario, 0 AS PrecoUnitarioMoeda, I.ValorDesconto, P.OrigemMercadoria " & _
'''                                                    "FROM TRIB_OLD_TributacaoPRJPropItem AS T, PRJPropostaItens AS I, Produtos AS P, PRJPropostas AS V WHERE T.NumIntDoc = I.NumIntDoc AND V.NumIntDoc = I.NumIntDocProposta AND P.Codigo = I.Produto AND V.NumIntDoc = ? ORDER By I.NumIntDoc "
'''                sSQLDoc = "SELECT NumIntDoc, NaturezaOP, ValorTotal, ValorFrete, ValorSeguro, ValorDesconto, ValorOutrasDespesas FROM PRJPropostas WHERE Codigo = ? AND FilialEmpresa = ? "
'''                sSQLUpdateDoc = "UPDATE TRIB_OLD_TributacaoPRJProp SET Atualizado = ? WHERE CodProposta = ? AND FilialEmpresa = ?"
'''
'''            Case 8 'PRJ CTR
'''                iTipoDoc = TIPODOC_TRIB_PRJ_CTR
'''                iTipoChave = TIPO_CHAVE_COD_FILIAL_STRING
'''                sSQLTribDoc = "SELECT FilialEmpresa, TaxacaoAutomatica, CodContrato As NumIntDoc,'' As NaturezaOpInterna, TipoTributacao,IPIBase,IPIValor,IPICredito,ICMSBase,ICMSValor,ICMSSubstBase,ICMSSubstValor,ICMSCredito,ISSIncluso,ISSBase,ISSAliquota,ISSValor,IRRFBase,IRRFAliquota,IRRFValor,ValorINSS,INSSRetido,INSSValorBase,INSSValorDeducoes,PISCredito,COFINSCredito,PISRetido,COFINSRetido,CSLLRetido, ISSRetido FROM TRIB_OLD_TributacaoPRJCTR WHERE Atualizado = 0 "
'''                sSQLTribCompl = "SELECT Item AS Tipo,IPIBaseCalculo,IPIValor,IPICredito,ICMSBase,ICMSValor,ICMSCredito,ICMSSubstBase,ICMSSubstValor,PISCredito,COFINSCredito FROM TRIB_OLD_TributacaoComplPRJCTR WHERE CodContrato = ? AND FilialEmpresa = ? "
'''                sSQLTribDocItem = "SELECT V.NumIntDoc,0 AS Item,I.NumIntDoc,T.NaturezaOp,T.TipoTributacao,T.IPITipo,P.IPICodigo,T.IPIBaseCalculo,T.IPIPercRedBase,T.IPIAliquota,T.IPIValor,T.IPICredito,T.ICMSTipo,T.ICMSBase,T.ICMSPercRedBase,T.ICMSAliquota,T.ICMSValor,T.ICMSCredito,T.ICMSSubstBase,T.ICMSSubstAliquota,T.ICMSSubstValor,T.PISCredito,T.COFINSCredito,0 AS ICMSAliquotaAdicaoDI,0 AS DespImpICMSBase,0 AS DespImpICMSValor,0 AS DespImpICMSCredito,0 AS ICMSSubstPercRedBase,0 AS ICMSSubstPercMVA,  " & _
'''                                                    "I.Produto , I.DescProd, I.UnidadeMed, I.Quantidade, I.Quantidade * I.PrecoUnitario - I.ValorDesconto, I.PrecoUnitario, 0 AS PrecoUnitarioMoeda, I.ValorDesconto, P.OrigemMercadoria " & _
'''                                                    "FROM TRIB_OLD_TributacaoPRJCTRItem AS T, PRJContratoItens AS I, Produtos AS P, PRJContratos AS V WHERE T.NumIntDoc = I.NumIntDoc AND V.NumIntDoc = I.NumIntDocContrato AND P.Codigo = I.Produto AND V.NumIntDoc = ? ORDER By I.NumIntDoc "
'''                sSQLDoc = "SELECT NumIntDoc, NaturezaOP, ValorTotal, ValorFrete, ValorSeguro, ValorDesconto, ValorOutrasDespesas FROM PRJPropostas WHERE Codigo = ? AND FilialEmpresa = ? "
'''                sSQLUpdateDoc = "UPDATE TRIB_OLD_TributacaoPRJCTR SET Atualizado = ? WHERE CodContrato = ? AND FilialEmpresa = ? "
                
        End Select
    
        'Le os select passados e repassa para função que vai converter um a um
        lErro = Tributacao_Atualiza_Versao2_Aux(iTipoChave, iTipoDoc, sSQLTribDoc, sSQLTribCompl, sSQLTribDocItem, sSQLDoc, sSQLUpdateDoc)
        If lErro <> SUCESSO Then gError 202903
    
    Next

    Tributacao_Atualiza_Versao2 = SUCESSO

    Exit Function

Erro_Tributacao_Atualiza_Versao2:

    Tributacao_Atualiza_Versao2 = gErr

    Select Case gErr
    
        Case 202903

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202904)

    End Select

    Exit Function
    
End Function

Public Function Tributacao_Atualiza_Versao2_Aux(ByVal iTipoChave As Integer, ByVal iTipoDoc As Integer, ByVal sSQLTribDoc As String, ByVal sSQLTribCompl As String, ByVal sSQLTribDocItem As String, ByVal sSQLDoc As String, ByVal sSQLUpdateDoc As String) As Long
'Função que monta a estrutura conforme os selects passados e repassa para função que faz o acerto de fato

Dim lErro As Long
'Dim lTransacao As Long
Dim alComando(1 To 5) As Long
Dim iIndice As Integer
Dim objTribDoc As ClassTributacaoDoc
Dim objTribDocItem As ClassTributacaoDocItem
Dim tTribDoc As typeTributacaoDoc
Dim tTribDocCompl As typeTributacaoDoc
Dim objTribDocFrete As ClassTributacaoDoc
Dim objTribDocSeguro As ClassTributacaoDoc
Dim objTribDocDespesa As ClassTributacaoDoc
Dim objTribDocDesconto As ClassTributacaoDoc
Dim tTribDocItem As typeTributacaoDocItem
Dim dValorFrete As Double, dValorSeguro As Double, dValorDesconto As Double
Dim dValorDespesa As Double, dValorTotal As Double
Dim iTipo As Integer, iFilialEmpresa As Integer
Dim lNumIntDoc As Long, sNaturezaOp As String, lCodigo As Long, vChave As Variant
Const TIPO_CHAVE_NUMINTDOC = 0
Const TIPO_CHAVE_COD_FILIAL = 1
Const TIPO_CHAVE_COD_FILIAL_STRING = 2

On Error GoTo Erro_Tributacao_Atualiza_Versao2_Aux

'    'Abertura de transação
'    lTransacao = Transacao_Abrir()
'    If lTransacao = 0 Then gError 202905

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 202906
    Next
    
    If iTipoChave = TIPO_CHAVE_COD_FILIAL_STRING Then
        vChave = String(255, 0)
    Else
        vChave = CLng(0)
    End If
    
    'Le a Tributação
    With tTribDoc
        .sNaturezaOpInterna = String(STRING_NATUREZAOP_CODIGO, 0)
        lErro = Comando_Executar(alComando(1), sSQLTribDoc, iFilialEmpresa, .iTaxacaoAutomatica, vChave, .sNaturezaOpInterna, .iTipoTributacao, .dIPIBase, .dIPIValor, .dIPICredito, .dICMSBase, .dICMSValor, .dICMSSubstBase, .dICMSSubstValor, .dICMSCredito, .iISSIncluso, .dISSBase, .dISSAliquota, .dISSValor, .dIRRFBase, .dIRRFAliquota, .dIRRFValor, .dValorINSS, .iINSSRetido, .dINSSValorBase, .dINSSValorDeducoes, .dPISCredito, .dCOFINSCredito, .dPISRetido, .dCOFINSRetido, .dCSLLRetido, .dISSRetido)
    End With
    If lErro <> AD_SQL_SUCESSO Then gError 202907
    
    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202908
    
    Do While lErro <> AD_SQL_SEM_DADOS

        Set objTribDoc = New ClassTributacaoDoc
        Set objTribDocFrete = New ClassTributacaoDoc
        Set objTribDocSeguro = New ClassTributacaoDoc
        Set objTribDocDespesa = New ClassTributacaoDoc
        Set objTribDocDesconto = New ClassTributacaoDoc
        
        tTribDoc.iTipoDoc = iTipoDoc
        Call objTribDoc.Copia2(tTribDoc)
        
        'Le o documento
        sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
        lErro = Comando_Executar(alComando(2), sSQLDoc, lNumIntDoc, sNaturezaOp, dValorTotal, dValorFrete, dValorSeguro, dValorDesconto, dValorDespesa, vChave, iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 202909
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202910
        
        objTribDoc.dValorAux = dValorTotal
        
        objTribDoc.lNumIntDoc = lNumIntDoc
        If iTipoChave <> TIPO_CHAVE_NUMINTDOC Then
            objTribDoc.sNaturezaOpInterna = sNaturezaOp
        End If
        
        'Le o Complemento
        With tTribDocCompl
            lErro = Comando_Executar(alComando(3), sSQLTribCompl, iTipo, .dIPIBase, .dIPIValor, .dIPICredito, .dICMSBase, .dICMSValor, .dICMSCredito, .dICMSSubstBase, .dICMSSubstValor, .dPISCredito, .dCOFINSCredito, vChave, iFilialEmpresa)
        End With
        If lErro <> AD_SQL_SUCESSO Then gError 202911
        
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202912
        
        Do While lErro <> AD_SQL_SEM_DADOS
        
            'Preenche os dados do complemento no obj
            Select Case iTipo
                Case ITEM_TIPO_FRETE
                    Call objTribDocFrete.Copia2(tTribDocCompl)
                    objTribDocFrete.dValorAux = dValorFrete
                Case ITEM_TIPO_SEGURO
                    Call objTribDocSeguro.Copia2(tTribDocCompl)
                    objTribDocSeguro.dValorAux = dValorSeguro
                Case ITEM_TIPO_OUTRAS_DESP
                    Call objTribDocDespesa.Copia2(tTribDocCompl)
                    objTribDocDespesa.dValorAux = dValorDespesa
                Case ITEM_TIPO_DESCONTO
                    Call objTribDocDesconto.Copia2(tTribDocCompl)
                    objTribDocDesconto.dValorAux = dValorDesconto
            End Select

            'Busca Próximo
            lErro = Comando_BuscarProximo(alComando(3))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202913

        Loop

        'Le os Itens
        With tTribDocItem
            .sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
            .sIPICodProduto = String(STRING_PRODUTO_IPI_CODIGO, 0)
            .sProduto = String(STRING_PRODUTO, 0)
            .sProdutoDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
            .sUnidadeMed = String(STRING_UM_SIGLA, 0)
            lErro = Comando_Executar(alComando(4), sSQLTribDocItem, .lNumIntDoc, .iItem, .lNumIntDocItem, .sNaturezaOp, .iTipoTributacao, .iIPITipo, .sIPICodProduto, .dIPIBaseCalculo, .dIPIPercRedBase, .dIPIAliquota, .dIPIValor, .dIPICredito, .iICMSTipo, .dICMSBase, .dICMSPercRedBase, .dICMSAliquota, .dICMSValor, .dICMSCredito, .dICMSSubstBase, .dICMSSubstAliquota, .dICMSSubstValor, .dPISCredito, .dCOFINSCredito, .dICMSAliquotaAdicaoDI, .dDespImpICMSBase, .dDespImpICMSValor, .dDespImpICMSCredito, .dICMSSubstPercRedBase, .dICMSSubstPercMVA, _
                        .sProduto, .sProdutoDescricao, .sUnidadeMed, .dQuantidade, .dPrecoTotal, .dPrecoUnitario, .dPrecoUnitarioMoeda, .dDescontoGrid, .iOrigemMercadoria, objTribDoc.lNumIntDoc)
        End With
        If lErro <> AD_SQL_SUCESSO Then gError 202914
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(4))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202915
        
        Do While lErro <> AD_SQL_SEM_DADOS
        
            Set objTribDocItem = New ClassTributacaoDocItem

            tTribDocItem.iTipoDoc = tTribDoc.iTipoDoc
            Call objTribDocItem.Copia2(tTribDocItem)
            
            objTribDoc.colItens.Add objTribDocItem

            'Busca Próximo
            lErro = Comando_BuscarProximo(alComando(4))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202916

        Loop
        
        Set objTribDoc.objTribDocFrete = objTribDocFrete
        Set objTribDoc.objTribDocSeguro = objTribDocSeguro
        Set objTribDoc.objTribDocDespesa = objTribDocDespesa
        Set objTribDoc.objTribDocDesconto = objTribDocDesconto
        
        'Função que redistribui os valores e preenche os campos novos
        lErro = Tributacao_Atualiza_Versao2_AcertaValores(objTribDoc)
        If lErro <> SUCESSO Then gError 202917
    
        'Marca a Tributação do documento como atualizada
        lErro = Comando_Executar(alComando(5), sSQLUpdateDoc, MARCADO, vChave, iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 202918

        'Busca Próximo
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202919
    
    Loop
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

'    'Confirma a transação
'    lErro = Transacao_Commit()
'    If lErro <> AD_SQL_SUCESSO Then gError 202920

    Tributacao_Atualiza_Versao2_Aux = SUCESSO

    Exit Function

Erro_Tributacao_Atualiza_Versao2_Aux:

    Tributacao_Atualiza_Versao2_Aux = gErr

    Select Case gErr

        Case 202905
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 202906
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 202907, 202908, 202919
            Call Rotina_Erro(vbOKOnly, "ERRO_EXECUCAO_COMANDO_SQL", gErr, sSQLTribDoc)
        
        Case 202909, 202910
            Call Rotina_Erro(vbOKOnly, "ERRO_EXECUCAO_COMANDO_SQL", gErr, sSQLDoc)
        
        Case 202911 To 202913
            Call Rotina_Erro(vbOKOnly, "ERRO_EXECUCAO_COMANDO_SQL", gErr, sSQLTribCompl)
        
        Case 202914 To 202916
            Call Rotina_Erro(vbOKOnly, "ERRO_EXECUCAO_COMANDO_SQL", gErr, sSQLTribDocItem)
        
        Case 202917
        
        Case 202918
            Call Rotina_Erro(vbOKOnly, "ERRO_EXECUCAO_COMANDO_SQL", gErr, sSQLUpdateDoc)
        
        Case 202920
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202921)

    End Select
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Desfaz Transação
'    Call Transacao_Rollback

    Exit Function

End Function

Private Function Tributacao_Atualiza_Versao2_AcertaValores(ByVal objTribDoc As ClassTributacaoDoc) As Long
'Função que converte um obj ClassTributacaoDoc com os dados de valores, itens e complementos preenchidos
'no formato novo do BD

Dim lErro As Long
Dim objTribDocItem As ClassTributacaoDocItem
Dim objTribDocFrete As ClassTributacaoDoc
Dim objTribDocSeguro As ClassTributacaoDoc
Dim objTribDocDespesa As ClassTributacaoDoc
Dim objTribDocDesconto As ClassTributacaoDoc
Dim dValorBaseRateio As Double
Dim vCampo As Variant
Dim vValor As Variant
Dim colCampos As New Collection
Dim colValores As New Collection
Dim iIndice As Integer

On Error GoTo Erro_Tributacao_Atualiza_Versao2_AcertaValores

    'Não dá mais erro por conta de NF de Complemento
    'If objTribDoc.colItens.Count = 0 Then gError 202927

    Set objTribDocFrete = objTribDoc.objTribDocFrete
    Set objTribDocSeguro = objTribDoc.objTribDocSeguro
    Set objTribDocDespesa = objTribDoc.objTribDocDespesa
    Set objTribDocDesconto = objTribDoc.objTribDocDesconto

    'PASSO 0 -Acrescenta Frete\Seguro e Outras Despesas na Base e Valor dos Impostos do Documento
    'Não é necessário porque a nível de NF esses valores já foram colocados
'    objTribDoc.dIPIBase = objTribDoc.dIPIBase + objTribDocFrete.dIPIBase + objTribDocSeguro.dIPIBase + objTribDocDespesa.dIPIBase
'    objTribDoc.dIPIValor = objTribDoc.dIPIValor + objTribDocFrete.dIPIValor + objTribDocSeguro.dIPIValor + objTribDocDespesa.dIPIValor
'    objTribDoc.dIPICredito = objTribDoc.dIPICredito + objTribDocFrete.dIPICredito + objTribDocSeguro.dIPICredito + objTribDocDespesa.dIPICredito
'    objTribDoc.dICMSBase = objTribDoc.dICMSBase + objTribDocFrete.dICMSBase + objTribDocSeguro.dICMSBase + objTribDocDespesa.dICMSBase
'    objTribDoc.dICMSValor = objTribDoc.dICMSValor + objTribDocFrete.dICMSValor + objTribDocSeguro.dICMSValor + objTribDocDespesa.dICMSValor
'    objTribDoc.dICMSCredito = objTribDoc.dICMSCredito + objTribDocFrete.dICMSCredito + objTribDocSeguro.dICMSCredito + objTribDocDespesa.dICMSCredito
'    objTribDoc.dICMSSubstBase = objTribDoc.dICMSSubstBase + objTribDocFrete.dICMSSubstBase + objTribDocSeguro.dICMSSubstBase + objTribDocDespesa.dICMSSubstBase
'    objTribDoc.dICMSSubstValor = objTribDoc.dICMSSubstValor + objTribDocFrete.dICMSSubstValor + objTribDocSeguro.dICMSSubstValor + objTribDocDespesa.dICMSSubstValor
'    objTribDoc.dPISCredito = objTribDoc.dPISCredito + objTribDocFrete.dPISCredito + objTribDocSeguro.dPISCredito + objTribDocDespesa.dPISCredito
'    objTribDoc.dCOFINSCredito = objTribDoc.dCOFINSCredito + objTribDocFrete.dCOFINSCredito + objTribDocSeguro.dCOFINSCredito + objTribDocDespesa.dCOFINSCredito

    'PASSO 1 -Distribui Frete\Seguro\Outras Despesas e Desconto pelos itens e os impostos que não tinham
    dValorBaseRateio = 0
    For Each objTribDocItem In objTribDoc.colItens
        dValorBaseRateio = dValorBaseRateio + objTribDocItem.dPrecoTotal
    Next
    colCampos.Add "dValorDescontoItem"
    colCampos.Add "dValorOutrasDespesasItem"
    colCampos.Add "dValorSeguroItem"
    colCampos.Add "dValorFreteItem"
    colCampos.Add "dIPIBaseCalculo"
    colCampos.Add "dIPIValor"
    colCampos.Add "dIPICredito"
    colCampos.Add "dICMSBase"
    colCampos.Add "dICMSValor"
    colCampos.Add "dICMSCredito"
    colCampos.Add "dICMSSubstBase"
    colCampos.Add "dICMSSubstValor"
    colCampos.Add "dPISCredito"
    colCampos.Add "dCOFINSCredito"
    colCampos.Add "dCOFINSBase"
    colCampos.Add "dCOFINSValor"
    colCampos.Add "dPISBase"
    colCampos.Add "dPISValor"
    colCampos.Add "dISSBase"
    colCampos.Add "dISSValor"
    
    colValores.Add objTribDocDesconto.dValorAux
    colValores.Add objTribDocDespesa.dValorAux
    colValores.Add objTribDocSeguro.dValorAux
    colValores.Add objTribDocFrete.dValorAux
    colValores.Add objTribDocFrete.dIPIBase + objTribDocSeguro.dIPIBase + objTribDocDespesa.dIPIBase
    colValores.Add objTribDocFrete.dIPIValor + objTribDocSeguro.dIPIValor + objTribDocDespesa.dIPIValor
    colValores.Add objTribDocFrete.dIPICredito + objTribDocSeguro.dIPICredito + objTribDocDespesa.dIPICredito
    colValores.Add objTribDocFrete.dICMSBase + objTribDocSeguro.dICMSBase + objTribDocDespesa.dICMSBase
    colValores.Add objTribDocFrete.dICMSValor + objTribDocSeguro.dICMSValor + objTribDocDespesa.dICMSValor
    colValores.Add objTribDocFrete.dICMSCredito + objTribDocSeguro.dICMSCredito + objTribDocDespesa.dICMSCredito
    colValores.Add objTribDocFrete.dICMSSubstBase + objTribDocSeguro.dICMSSubstBase + objTribDocDespesa.dICMSSubstBase
    colValores.Add objTribDocFrete.dICMSSubstValor + objTribDocSeguro.dICMSSubstValor + objTribDocDespesa.dICMSSubstValor
    colValores.Add objTribDoc.dPISCredito
    colValores.Add objTribDoc.dCOFINSCredito
    colValores.Add IIf(objTribDoc.dCOFINSRetido > DELTA_VALORMONETARIO, objTribDoc.dValorAux, 0) 'CALCULAR BASE DO COFINS
    colValores.Add objTribDoc.dCOFINSRetido
    colValores.Add IIf(objTribDoc.dPISRetido > DELTA_VALORMONETARIO, objTribDoc.dValorAux, 0) 'CALCULAR BASE DO PIS
    colValores.Add objTribDoc.dPISRetido
    colValores.Add objTribDoc.dISSBase
    colValores.Add objTribDoc.dISSValor
    
    iIndice = 0
    For Each vCampo In colCampos
        iIndice = iIndice + 1
        vValor = colValores.Item(iIndice)
        'Ratea o valor passado entre os itens para o campo passado
        lErro = Tributacao_Atualiza_Versao2_AcertaValores2(vValor, dValorBaseRateio, vCampo, objTribDoc.colItens)
        If lErro <> SUCESSO Then gError 202922
    Next
    
    'PASSO 2 - Acerta os demais campos
    For Each objTribDocItem In objTribDoc.colItens
        'Verifica a incidência de PIS e COFINS para atualizar o campo tipo (novo)
        If objTribDocItem.dPISValor > DELTA_VALORMONETARIO Then
            objTribDocItem.iPISTipo = 1 'Incide
            objTribDocItem.iPISTipoCalculo = TRIB_TIPO_CALCULO_PERMITE_PERC
        Else
            objTribDocItem.iPISTipo = 8 'Não Incide
            objTribDocItem.iPISTipoCalculo = TRIB_TIPO_CALCULO_PERMITE_NADA
        End If
        If objTribDocItem.dCOFINSValor > DELTA_VALORMONETARIO Then
            objTribDocItem.iCOFINSTipo = 1 'Incide
            objTribDocItem.iCOFINSTipoCalculo = TRIB_TIPO_CALCULO_PERMITE_PERC
        Else
            objTribDocItem.iCOFINSTipo = 8 'Não Incide
            objTribDocItem.iCOFINSTipoCalculo = TRIB_TIPO_CALCULO_PERMITE_NADA
        End If
        objTribDocItem.iICMSBaseModalidade = ICMS_MODALIDADE_VALOR
        objTribDocItem.iICMSSubstBaseModalidade = ICMS_ST_MODALIDADE_MARGEM
        objTribDocItem.iIPITipoCalculo = TRIB_TIPO_CALCULO_PERMITE_PERC
        objTribDocItem.dISSAliquota = objTribDoc.dISSAliquota
        objTribDocItem.dQtdTrib = objTribDocItem.dQuantidade
        objTribDocItem.sUMTrib = objTribDocItem.sUnidadeMed
        objTribDocItem.dValorUnitTrib = objTribDocItem.dPrecoUnitario
        'Calcula a alíquota de PIS e COFINS pos a mesma não existe na tributação do documento
        If objTribDocItem.dPISValor > DELTA_VALORMONETARIO Then
            objTribDocItem.dPISAliquota = objTribDocItem.dPISValor / objTribDocItem.dPISBase
        End If
        If objTribDocItem.dCOFINSValor > DELTA_VALORMONETARIO Then
            objTribDocItem.dCOFINSAliquota = objTribDocItem.dCOFINSValor / objTribDocItem.dCOFINSBase
        End If
        
        'Tenta ver se a alíquota é próxima das alíquotas mais comuns e converte para evitar erro de arredondamento
        If Abs(objTribDocItem.dPISAliquota - 1.65) < 0.02 Then objTribDocItem.dPISAliquota = 1.65
        If Abs(objTribDocItem.dPISAliquota - 0.65) < 0.02 Then objTribDocItem.dPISAliquota = 0.65
    
        If Abs(objTribDocItem.dCOFINSAliquota - 7.6) < 0.02 Then objTribDocItem.dCOFINSAliquota = 7.6
        If Abs(objTribDocItem.dCOFINSAliquota - 3) < 0.02 Then objTribDocItem.dCOFINSAliquota = 3
    
    Next
    
    'PASSO 3 - Grava Dados
    'Exclui os dados caso já tenham sido atualizados antes
    lErro = TributacaoDoc_Exclui(objTribDoc.iTipoDoc, objTribDoc.lNumIntDoc, True)
    If lErro <> SUCESSO Then gError 202923
    
    'Marca os dados como se tivessem sofrido alterações na mão para não recalcular mais e grava
    Call objTribDoc.Coloca_Manual
    lErro = TributacaoDoc_Grava(objTribDoc)
    If lErro <> SUCESSO Then gError 202923
    
    For Each objTribDocItem In objTribDoc.colItens
        Call objTribDocItem.Coloca_Manual
        lErro = TributacaoDocItem_Grava(objTribDocItem)
        If lErro <> SUCESSO Then gError 202924
    Next

    Tributacao_Atualiza_Versao2_AcertaValores = SUCESSO

    Exit Function

Erro_Tributacao_Atualiza_Versao2_AcertaValores:

    Tributacao_Atualiza_Versao2_AcertaValores = gErr

    Select Case gErr
    
        Case 202922 To 202924
        
        Case 202927
            Call Rotina_Erro(vbOKOnly, "ERRO_NF_SEM_ITENSTRIB", gErr, objTribDoc.lNumIntDoc)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202925)

    End Select
     
    Exit Function

End Function

Private Function Tributacao_Atualiza_Versao2_AcertaValores2(ByVal dValorRateio As Double, ByVal dValorBaseRateio As Double, ByVal sNomeCampo As String, ByVal colItens As Collection) As Long
'Essa função recebe:
'dValorRateio = valor a ser rateado,
'dValorBaseRateio = soma do preço total de todos itens (o preço total vai servir como parâmetro para redistribuição)
'sNomeCampo = Campo que vai receber o valor rateado
'colItens = coleção de iten que vão receber o rateo
Dim dValorAcum As Double, dValorItem As Double, dDif As Double
Dim objDocItem As ClassTributacaoDocItem
Dim dValorAnterior As Double

On Error GoTo Erro_Tributacao_Atualiza_Versao2_AcertaValores2
    
    'Se tem algum valor a ser rateado
    If dValorRateio > DELTA_VALORMONETARIO And dValorBaseRateio > DELTA_VALORMONETARIO Then
                 
        'percorrer os itens guardando valores "aproximados"
        dValorAcum = 0
        For Each objDocItem In colItens
        
            dValorItem = Int(100 * objDocItem.dPrecoTotal * dValorRateio / dValorBaseRateio) / 100
            dValorAcum = dValorAcum + dValorItem
            objDocItem.dValorAux = dValorItem
    
        Next
         
        dDif = dValorRateio - dValorAcum
        
        'Se deu diferença para mais
        If dDif > DELTA_VALORMONETARIO Then
            Do While dDif > DELTA_VALORMONETARIO
                'percorrer os itens aumentando em R$ 0,01 o valor
                For Each objDocItem In colItens
                    objDocItem.dValorAux = objDocItem.dValorAux + 0.01
                    dDif = dDif - 0.01
                    If Abs(dDif) < DELTA_VALORMONETARIO Then Exit For
                Next
            Loop
        'Se deu diferença para menos
        ElseIf dDif < -DELTA_VALORMONETARIO Then
            Do While dDif < -DELTA_VALORMONETARIO
                'percorrer os itens diminuindo em R$ 0,01 o valor
                For Each objDocItem In colItens
                    objDocItem.dValorAux = objDocItem.dValorAux - 0.01
                    dDif = dDif + 0.01
                    If Abs(dDif) < DELTA_VALORMONETARIO Then Exit For
                Next
            Loop
        End If
        
        'Coloca o valor rateado no campo certo
        For Each objDocItem In colItens
            dValorAnterior = CallByName(objDocItem, sNomeCampo, VbGet)
            Call CallByName(objDocItem, sNomeCampo, VbLet, objDocItem.dValorAux + dValorAnterior)
        Next
        
    End If
    
    Tributacao_Atualiza_Versao2_AcertaValores2 = SUCESSO

    Exit Function

Erro_Tributacao_Atualiza_Versao2_AcertaValores2:

    Tributacao_Atualiza_Versao2_AcertaValores2 = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB_TRAT1", gErr, sNomeCampo, Error, 202926)

    End Select
     
    Exit Function
    
End Function

Public Function Tributacao_Imposto_ObtemBaseItem(ByVal iFilialEmpresa As Integer, ByVal sImposto As String, ByVal objDoc As ClassTributoDoc, ByVal objDocItem As ClassTributoDocItem, dBase As Double) As Long

Dim lErro As Long
Dim objVarItemValor As New ClassVariavelCalculo
Dim objVarItemOutValor As New ClassVariavelCalculo
Dim objVarItemIPI As New ClassVariavelCalculo
Dim objVarItemDesc As New ClassVariavelCalculo
Dim objVarItemFrete As New ClassVariavelCalculo
Dim objVarTipoTrib As New ClassVariavelCalculo
Dim objVarItemValorBruto As New ClassVariavelCalculo
Dim objVarISSIncluso As New ClassVariavelCalculo
Dim objVarISSValor As New ClassVariavelCalculo
Dim objVarDataEmissao As New ClassVariavelCalculo
Dim objVarIndConsumidorFinal As New ClassVariavelCalculo
Dim objTRBConfig As New ClassTRBConfig
Dim objTipoTrib As New ClassTipoDeTributacaoMovto
Dim objVarTipoDocInfo As New ClassVariavelCalculo
Dim objVarItemNFOrig As New ClassVariavelCalculo, objItemNFOrig As ClassItemNF
Dim dtDataEmissao As Date, bConsumidorFinal As Boolean

On Error GoTo Erro_Tributacao_Imposto_ObtemBaseItem

    dBase = 0
    
    'data de emissão
    lErro = objDoc.ObterVar("DATA_EMISSAO", objVarDataEmissao)
    If lErro <> SUCESSO Then gError 202927
    
    dtDataEmissao = objVarDataEmissao.vValor

    'Indicador de consumidor final
    lErro = objDoc.ObterVar("IND_CONSUMIDOR_FINAL", objVarIndConsumidorFinal)
    If lErro <> SUCESSO Then gError 202927
    
    bConsumidorFinal = objVarIndConsumidorFinal.vValor
    
    'Preço Item
    lErro = objDocItem.ObterVar("PRODUTO_VALOR", objVarItemValor)
    If lErro <> SUCESSO Then gError 202927
    
    lErro = objDoc.ObterVar("TIPO_DOC_CODIGO", objVarTipoDocInfo)
    If lErro <> SUCESSO Then gError 202927
            
    'Frete\Seguro, etc Rateado
    lErro = objDocItem.ObterVar("PRODUTO_OUTROS_VALORES", objVarItemOutValor)
    If lErro <> SUCESSO Then gError 202928
    
    lErro = objDocItem.ObterVar("TIPO_TRIB", objVarTipoTrib)
    If lErro <> SUCESSO Then gError 202929
    
    'obter mais info sobre o tipo trib
    objTipoTrib.iTipo = objVarTipoTrib.vValor
    lErro = objTRBConfig.gobjTributacao.TributacaoTipo_Obter(objTipoTrib)
    If lErro <> SUCESSO Then gError 202930
    
    'obter valor do IPI do item de objdocitem
    lErro = objDocItem.ObterVar("IPI_VALOR", objVarItemIPI)
    If lErro <> SUCESSO Then gError 202931
    
    'Obtem o desconto
    lErro = objDocItem.ObterVar("VALOR_DESCONTO_ITEM", objVarItemDesc)
    If lErro <> SUCESSO Then gError 202932
    
    'Obtem o frete
    lErro = objDocItem.ObterVar("VALOR_FRETE_ITEM", objVarItemFrete)
    If lErro <> SUCESSO Then gError 202932
    
    'Valor Bruto
    lErro = objDocItem.ObterVar("PRODUTO_VALOR_BRUTO", objVarItemValorBruto)
    If lErro <> SUCESSO Then gError 202933
    
    'O ISS está incluso
    lErro = objDoc.ObterVar("ISS_INCLUSO", objVarISSIncluso)
    If lErro <> SUCESSO Then gError 202934
         
    'Valor do ISS
    lErro = objDoc.ObterVar("ISS_VALOR", objVarISSValor)
    If lErro <> SUCESSO Then gError 202935

    'obter o valor de ICMS ST que possa estar embutido no campo de outras despesas para abate-lo
    If objVarTipoDocInfo.vValor = DOCINFO_NFIEDV Then
    
        lErro = objDocItem.ObterVar("OBJ_ITEM_NF_ORIG", objVarItemNFOrig)
        If lErro <> SUCESSO Then Error 27721
        
        Set objItemNFOrig = objVarItemNFOrig.vValor
        
        If Not (objItemNFOrig Is Nothing) Then
    
            objVarItemOutValor.vValor = Arredonda_Moeda(objVarItemOutValor.vValor - objItemNFOrig.objTributacaoItemNF.dICMSSubstValor)
    
        End If
        
    End If
    
    Select Case sImposto
    
        Case "PIS"
            dBase = objVarItemValor.vValor + objVarItemOutValor.vValor
    
        Case "COFINS"
            dBase = objVarItemValor.vValor + objVarItemOutValor.vValor
    
        Case "ICMS"
            dBase = objVarItemValor.vValor + objVarItemOutValor.vValor + IIf(objTipoTrib.iICMSBaseComIPI Or (bConsumidorFinal And gobjCRFAT.iICMSIncluiIPIBaseConsFinal = MARCADO), objVarItemIPI.vValor, 0)
    
        Case "IPI"
            If objVarTipoDocInfo.vValor <> DOCINFO_NFIEIMP And objVarTipoDocInfo.vValor <> DOCINFO_NFIEIMPSE And objVarTipoDocInfo.vValor <> DOCINFO_NFIEIMPPC Then
                If gobjCRFAT.iIPIAbateDesconto = 0 And Not (dtDataEmissao >= gobjCRFAT.dtIPIIniDescExcBase And gobjCRFAT.dtIPIIniDescExcBase <> DATA_NULA) Then
                    dBase = objVarItemValorBruto.vValor + objVarItemOutValor.vValor + objVarItemDesc.vValor - IIf(objTipoTrib.iIPIFrete, 0, objVarItemFrete.vValor)
                Else
                    dBase = objVarItemValor.vValor + objVarItemOutValor.vValor - IIf(objTipoTrib.iIPIFrete, 0, objVarItemFrete.vValor)
                End If
            Else
                If gobjCRFAT.iIPIAbateDesconto = 0 And Not (dtDataEmissao >= gobjCRFAT.dtIPIIniDescExcBase And gobjCRFAT.dtIPIIniDescExcBase <> DATA_NULA) Then
                    dBase = objVarItemValorBruto.vValor + objVarItemDesc.vValor
                Else
                    dBase = objVarItemValor.vValor
                End If
            End If
        Case "CSLL"
            dBase = objVarItemValor.vValor + IIf(objVarISSIncluso.vValor = 0, objVarISSValor.vValor, 0)
            
        Case "ISS"
            dBase = objVarItemValor.vValor
        
        Case "INSS"
            dBase = objVarItemValor.vValor + objVarItemOutValor.vValor
            
        Case "IR"
            dBase = objVarItemValor.vValor + IIf(objVarISSIncluso.vValor = 0, objVarISSValor.vValor, 0)
    
        Case "ICMS_II"
            dBase = objVarItemValor.vValor + IIf(objTipoTrib.iICMSBaseComIPI, objVarItemIPI.vValor, 0) 'As despesas de importação são adicionadas depois junto com a própria base quando for o caso
    
    End Select

    dBase = Arredonda_Moeda(dBase)

    Tributacao_Imposto_ObtemBaseItem = SUCESSO

    Exit Function

Erro_Tributacao_Imposto_ObtemBaseItem:

    Tributacao_Imposto_ObtemBaseItem = gErr

    Select Case gErr
    
        Case 202927 To 202935

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202926)

    End Select
     
    Exit Function
    
End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'EXCLUSÃO
Public Function PISCOFINSExcecoes_Exclui(ByVal objPISCOFINSExcecao As ClassPISCOFINSExcecao) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_PISCOFINSExcecoes_Exclui

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 205477

    lErro = CF("PISCOFINSExcecoes_Exclui_EmTrans", objPISCOFINSExcecao)
    If lErro <> SUCESSO Then gError 205478

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 205479

    PISCOFINSExcecoes_Exclui = SUCESSO

    Exit Function

Erro_PISCOFINSExcecoes_Exclui:

    PISCOFINSExcecoes_Exclui = gErr

    Select Case gErr

        Case 205477
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 205478

        Case 205479
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205480)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function PISCOFINSExcecoes_Exclui_EmTrans(ByVal objPISCOFINSExcecao As ClassPISCOFINSExcecao) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim iAux As Integer

On Error GoTo Erro_PISCOFINSExcecoes_Exclui_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 205481
    Next

    'Le a tabelaPISCOFINSExcecoes
    lErro = Comando_ExecutarPos(alComando(0), "SELECT 1 FROM PISCOFINSExcecoes WHERE CategoriaCliente= ?  AND CategoriaClienteItem= ?  AND CategoriaProduto= ?  AND CategoriaProdutoItem= ? AND CategoriaFornecedor = ?  AND CategoriaFornecedorItem= ? AND TipoCliForn = ?", _
                0, iAux, objPISCOFINSExcecao.sCategoriaCliente, objPISCOFINSExcecao.sCategoriaClienteItem, objPISCOFINSExcecao.sCategoriaProduto, objPISCOFINSExcecao.sCategoriaProdutoItem, objPISCOFINSExcecao.sCategoriaFornecedor, objPISCOFINSExcecao.sCategoriaFornecedorItem, objPISCOFINSExcecao.iTipoCliForn)
    If lErro <> AD_SQL_SUCESSO Then gError 205482

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 205483

    'Se não existir => ERRO
    If lErro = AD_SQL_SEM_DADOS Then gError 205484

    'Trava registro contra alterações/Leituras
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> SUCESSO Then gError 205485

    lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM PISCOFINSExcecoes", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 205486

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PISCOFINSExcecoes_Exclui_EmTrans = SUCESSO

    Exit Function

Erro_PISCOFINSExcecoes_Exclui_EmTrans:

    PISCOFINSExcecoes_Exclui_EmTrans = gErr

    Select Case gErr

        Case 205481
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 205482, 205483
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PISCOFINSEXCECOES", gErr)

        Case 205484
            Call Rotina_Erro(vbOKOnly, "ERRO_PISCOFINSEXCECOES_NAO_CADASTRADO", gErr, objPISCOFINSExcecao.sCategoriaCliente, objPISCOFINSExcecao.sCategoriaClienteItem, objPISCOFINSExcecao.sCategoriaProduto, objPISCOFINSExcecao.sCategoriaProdutoItem)

        Case 205485
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_PISCOFINSEXCECOES", gErr)

        Case 205486
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PISCOFINSEXCECOES", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205487)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'GRAVAÇÃO
Public Function PISCOFINSExcecoes_Grava(ByVal objPISCOFINSExcecao As ClassPISCOFINSExcecao) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_PISCOFINSExcecoes_Grava

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 205488

    lErro = CF("PISCOFINSExcecoes_Grava_EmTrans", objPISCOFINSExcecao)
    If lErro <> SUCESSO Then gError 205489

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 205490

    PISCOFINSExcecoes_Grava = SUCESSO

    Exit Function

Erro_PISCOFINSExcecoes_Grava:

    PISCOFINSExcecoes_Grava = gErr

    Select Case gErr

        Case 205488
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 205489

        Case 205490
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205491)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function PISCOFINSExcecoes_Grava_EmTrans(ByVal objPISCOFINSExcecao As ClassPISCOFINSExcecao) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim iAux As Integer

On Error GoTo Erro_PISCOFINSExcecoes_Grava_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 205492
    Next

    'Le a tabelaPISCOFINSExcecoes
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Tipo FROM PISCOFINSExcecoes WHERE CategoriaCliente= ?  AND CategoriaClienteItem= ?  AND CategoriaProduto= ?  AND CategoriaProdutoItem= ? AND CategoriaFornecedor = ? AND CategoriaFornecedorItem = ? AND TipoCliForn = ?", _
        0, iAux, objPISCOFINSExcecao.sCategoriaCliente, objPISCOFINSExcecao.sCategoriaClienteItem, objPISCOFINSExcecao.sCategoriaProduto, objPISCOFINSExcecao.sCategoriaProdutoItem, objPISCOFINSExcecao.sCategoriaFornecedor, objPISCOFINSExcecao.sCategoriaFornecedorItem, objPISCOFINSExcecao.iTipoCliForn)
    If lErro <> AD_SQL_SUCESSO Then gError 205493

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 205494

    'Se existir => UPDATE, senão => INSERT
    If lErro = AD_SQL_SUCESSO Then

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 205495

        lErro = Comando_ExecutarPos(alComando(1), "UPDATE PISCOFINSExcecoes SET Tipo= ? , TipoPIS= ? , TipoCOFINS= ? , PISTipoCalculo= ? , " & _
                "COFINSTipoCalculo= ? , AliquotaPisRS= ? , AliquotaPisPerc= ? , AliquotaCofinsRS= ? , AliquotaCofinsPerc= ? , " & _
                "Fundamentacao= ? , Prioridade= ? , TipoPISE= ? , TipoCOFINSE= ? ", alComando(0), _
                objPISCOFINSExcecao.iTipo, objPISCOFINSExcecao.iTipoPIS, objPISCOFINSExcecao.iTipoCOFINS, objPISCOFINSExcecao.iPISTipoCalculo, _
                objPISCOFINSExcecao.iCOFINSTipoCalculo, objPISCOFINSExcecao.dAliquotaPisRS, objPISCOFINSExcecao.dAliquotaPisPerc, objPISCOFINSExcecao.dAliquotaCofinsRS, objPISCOFINSExcecao.dAliquotaCofinsPerc, _
                objPISCOFINSExcecao.sFundamentacao, objPISCOFINSExcecao.iPrioridade, objPISCOFINSExcecao.iTipoPISE, objPISCOFINSExcecao.iTipoCOFINSE)
        If lErro <> AD_SQL_SUCESSO Then gError 205496

    Else


        lErro = Comando_Executar(alComando(1), "INSERT INTO PISCOFINSExcecoes( CategoriaCliente, CategoriaClienteItem, CategoriaProduto, CategoriaProdutoItem, " & _
                "Tipo, TipoPIS, TipoCOFINS, PISTipoCalculo, COFINSTipoCalculo, " & _
                "AliquotaPisRS, AliquotaPisPerc, AliquotaCofinsRS, AliquotaCofinsPerc, Fundamentacao, " & _
                "Prioridade, CategoriaFornecedor, CategoriaFornecedorItem, TipoCliForn, TipoPISE, TipoCOFINSE) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                objPISCOFINSExcecao.sCategoriaCliente, objPISCOFINSExcecao.sCategoriaClienteItem, objPISCOFINSExcecao.sCategoriaProduto, objPISCOFINSExcecao.sCategoriaProdutoItem, _
                objPISCOFINSExcecao.iTipo, objPISCOFINSExcecao.iTipoPIS, objPISCOFINSExcecao.iTipoCOFINS, objPISCOFINSExcecao.iPISTipoCalculo, objPISCOFINSExcecao.iCOFINSTipoCalculo, _
                objPISCOFINSExcecao.dAliquotaPisRS, objPISCOFINSExcecao.dAliquotaPisPerc, objPISCOFINSExcecao.dAliquotaCofinsRS, objPISCOFINSExcecao.dAliquotaCofinsPerc, objPISCOFINSExcecao.sFundamentacao, _
                objPISCOFINSExcecao.iPrioridade, objPISCOFINSExcecao.sCategoriaFornecedor, objPISCOFINSExcecao.sCategoriaFornecedorItem, objPISCOFINSExcecao.iTipoCliForn, objPISCOFINSExcecao.iTipoPISE, objPISCOFINSExcecao.iTipoCOFINSE)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 205497

    End If

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PISCOFINSExcecoes_Grava_EmTrans = SUCESSO

    Exit Function

Erro_PISCOFINSExcecoes_Grava_EmTrans:

    PISCOFINSExcecoes_Grava_EmTrans = gErr

    Select Case gErr

        Case 205492
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 205493, 205494
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PISCOFINSEXCECOES", gErr)

        Case 205495
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_PISCOFINSEXCECOES", gErr)

        Case 205496
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_PISCOFINSEXCECOES", gErr)

        Case 205497
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PISCOFINSEXCECOES", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205498)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function



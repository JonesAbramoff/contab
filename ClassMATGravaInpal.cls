VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassMATGravaInpal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function ItemOP_Baixa_EmTrans(objItemOP As ClassItemOP) As Long
'Baixa o item da ordem de producao

Dim sComando_SQL As String
Dim lErro As Long
Dim tItemOP As typeItemOP
Dim alComando(0 To 10) As Long
Dim iIndice As Integer
Dim objEstoqueProduto As New ClassEstoqueProduto

On Error GoTo Erro_ItemOP_Baixa_EmTrans
        
    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 93660
    Next
   
    sComando_SQL = "SELECT Situacao, Almoxarifado, Quantidade, QuantidadeProd FROM ItensOrdemProducao WHERE FilialEmpresa = ? And Produto = ? And Codigo = ?"

    'Lê ItemOP da Tabela ItensOrdemProducao a partir de giFilialEmpresa , objItemOP.OPCodigo e objItemOP.Produto
    lErro = Comando_ExecutarPos(alComando(0), sComando_SQL, 0, tItemOP.iSituacao, tItemOP.iAlmoxarifado, tItemOP.dQuantidade, tItemOP.dQuantidadeProd, objItemOP.iFilialEmpresa, objItemOP.sProduto, objItemOP.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 93651

    'Faz busca ao registro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 93652

    'Se não encontrou nenhum registro
    If lErro = AD_SQL_SEM_DADOS Then gError 93659

    If tItemOP.iSituacao = ITEMOP_SITUACAO_DESAB Then gError 93653

    If tItemOP.iSituacao = ITEMOP_SITUACAO_BAIXADA Then gError 93654

    'Lock no registro encontrado
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 93655

    lErro = Comando_ExecutarPos(alComando(1), "UPDATE ItensOrdemProducao Set Situacao = ?", alComando(0), ITEMOP_SITUACAO_BAIXADA)
    If lErro <> AD_SQL_SUCESSO Then gError 93656
    
    'altera a qtde em OP para o produto-almoxarifado na tabela EstoqueProduto
    If tItemOP.dQuantidade > tItemOP.dQuantidadeProd Then

        objEstoqueProduto.sProduto = objItemOP.sProduto
        objEstoqueProduto.iAlmoxarifado = tItemOP.iAlmoxarifado
    
        lErro = CF("EstoqueProduto_AlterarQtdOP", objEstoqueProduto, -(tItemOP.dQuantidade - tItemOP.dQuantidadeProd))
        If lErro <> SUCESSO Then gError 95376

    End If

    'Atualiza o campo NumItensBaixados=NumItensBaixados+1 na Tabela OrdensdeProducao
    'Exclui empenhos associados ao ItemOP
    lErro = CF("OrdemDeProducao_TrataBaixaItem", objItemOP, alComando)
    If lErro <> SUCESSO Then gError 93657

    'Verifica se todos os itens da OP estão baixados, se tiverem ==> baixa a OP
    'Se todos os itens estiverem baixados, insere na tabela OrdensdeProducaoBaixadas
    'Insere na tabela ItensOrdemProducaoBaixados
    'Exclui Empenho associados ao itemOp
    'Exclui o registro da tabela ItensOrdemProducao
    'Exclui o registro da tabela OrdemdeProducao
    lErro = CF("OrdemDeProducao_Baixa1", alComando(), objItemOP.sCodigo)
    
    If lErro <> SUCESSO Then gError 93658

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next
    
    ItemOP_Baixa_EmTrans = SUCESSO

    Exit Function

Erro_ItemOP_Baixa_EmTrans:

    ItemOP_Baixa_EmTrans = gErr

    Select Case gErr

        Case 95376
        
        Case 93660
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
           
        Case 93656
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_ITENSORDENSDEPRODUCAO", gErr, tItemOP.iItem, objItemOP.sCodigo, tItemOP.iFilialEmpresa)

        Case 93651, 93652
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSORDENSPRODUCAO", gErr)

        Case 93659
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEMOP_SITUACAO_BAIXADA", gErr, objItemOP.sProduto, objItemOP.sCodigo)
        
        Case 93655
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_ITENSORDENSDEPRODUCAO", gErr)

        Case 93653
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEMOP_SITUACAO_DESABILITADA", gErr, objItemOP.sProduto, objItemOP.sCodigo)

        Case 93654
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEMOP_SITUACAO_BAIXADA", gErr, objItemOP.sProduto, objItemOP.sCodigo)

        Case 93658
            Call Rotina_Erro(vbOKOnly, "ERRO_ORDEMPRODUCAO_BAIXADA", gErr, objItemOP.sCodigo)
     
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function ItemOP_Baixa(objItemOP As ClassItemOP) As Long
'Baixa o item da ordem de producao

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_ItemOP_Baixa

    'Inicia a transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 93661

    lErro = ItemOP_Baixa_EmTrans(objItemOP)
    If lErro <> SUCESSO Then gError 81737

    'confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 93662

    ItemOP_Baixa = SUCESSO

    Exit Function

Erro_ItemOP_Baixa:

    ItemOP_Baixa = gErr

    Select Case gErr

        Case 81737
        
        Case 93661
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
           
        Case 93662
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    'confirma a transação
    Call Transacao_Rollback

    Exit Function

End Function

Function ItensOP_Baixa(ColItemOP As Collection) As Long
'Baixa os itens de ordem de producao passados na colecao

Dim lErro As Long
Dim lTransacao As Long
Dim objItemOP As ClassItemOP

On Error GoTo Erro_ItensOP_Baixa

    'Inicia a transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 81738

    For Each objItemOP In ColItemOP
    
        lErro = ItemOP_Baixa_EmTrans(objItemOP)
        If lErro <> SUCESSO Then gError 81740
    
    Next
    
    'confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 81739

    ItensOP_Baixa = SUCESSO

    Exit Function

Erro_ItensOP_Baixa:

    ItensOP_Baixa = gErr

    Select Case gErr

        Case 81740
        
        Case 81738
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
           
        Case 81739
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    'confirma a transação
    Call Transacao_Rollback

    Exit Function

End Function



Attribute VB_Name = "Module2"
'Responsavel pela Rotina Atualizacao Inventário: Mario
'Data: 30/10/98
'Pendencias:

'Atualização e Desatualização de InvLote

Option Explicit

Public Function Rotina_Atualiza_InvLote_Int(iID_Atualizacao As Integer, objAtuInvLoteAux As Object) As Long

Dim lErro As Long
Dim lComando As Long
Dim iIndice As Integer
Dim objInvLote As ClassInvLote
Dim tInvLote As typeInvLote
Dim colInvLote As New Collection
Dim sComando_SQL As String
Dim vbMsgBox As VbMsgBoxResult

On Error GoTo Erro_Rotina_Atualiza_InvLote_Int

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 41369

    sComando_SQL = "SELECT FilialEmpresa, Lote FROM InvLotePendente WHERE IdAtualizacao = ? ORDER BY FilialEmpresa, Lote"

    'Pesquisa os lotes de inventário pendentes no banco de dados
    lErro = Comando_Executar(lComando, sComando_SQL, tInvLote.iFilialEmpresa, tInvLote.iLote, iID_Atualizacao)
    If lErro <> AD_SQL_SUCESSO Then Error 41370

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 41371

    Do While lErro <> AD_SQL_SEM_DADOS

        Set objInvLote = New ClassInvLote

        objInvLote.iFilialEmpresa = tInvLote.iFilialEmpresa
        objInvLote.iLote = tInvLote.iLote

        colInvLote.Add objInvLote

        'le o proximo lote de inventário pendente
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 41372

    Loop

    TelaAcompanhaBatchEST.dValorTotal = colInvLote.Count

    For Each objInvLote In colInvLote

        tInvLote.iFilialEmpresa = objInvLote.iFilialEmpresa
        tInvLote.iLote = objInvLote.iLote

        'para cada lote pendente
        lErro = Atualiza_InvLote(tInvLote, iID_Atualizacao, objAtuInvLoteAux)
'        If lErro <> SUCESSO Then Error 41373
'retirado o tratamento de Erro para poder processar varios lotes mesmo que um ou outro apresente erro

        lErro = DoEvents()
        
        TelaAcompanhaBatchEST.dValorAtual = TelaAcompanhaBatchEST.dValorAtual + 1

        TelaAcompanhaBatchEST.ProgressBar1.Value = CInt((TelaAcompanhaBatchEST.dValorAtual / TelaAcompanhaBatchEST.dValorTotal) * 100)
        
        If TelaAcompanhaBatchEST.iCancelaBatch = CANCELA_BATCH Then
        
            vbMsgBox = Rotina_Aviso(vbYesNo, "AVISO_CANCELAR_ATUALIZACAO_INVLOTE")
            
            If vbMsgBox = vbYes Then Error 41538
                
            TelaAcompanhaBatchEST.iCancelaBatch = 0
                
        End If


    Next

    Call Comando_Fechar(lComando)

    Rotina_Atualiza_InvLote_Int = SUCESSO

    Exit Function

Erro_Rotina_Atualiza_InvLote_Int:

    Rotina_Atualiza_InvLote_Int = Err

    Select Case Err

        Case 41369
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 41370, 41371, 41372
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_INVLOTEPENDENTE", Err)

        Case 41373, 41538

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", lErro, Error$, 159533)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function InvLotePendente_Le(lComando As Long, tInvLote As typeInvLote, iFilialEmpresa As Integer, iLote As Integer) As Long

Dim sComando_SQL As String
Dim lErro As Long

On Error GoTo Erro_InvLotePendente_Le

    tInvLote.sDescricao = String(STRING_INVLOTE_DESCRICAO, 0)

    sComando_SQL = "SELECT Descricao, NumItensInf, NumItensAtual, IdAtualizacao FROM InvLotePendente WHERE FilialEmpresa = ? AND Lote = ?"

    'Pesquisa o lote de inventário pendente no banco de dados
    lErro = Comando_ExecutarPos(lComando, sComando_SQL, 0, tInvLote.sDescricao, tInvLote.iNumItensInf, tInvLote.iNumItensAtual, tInvLote.iIDAtualizacao, iFilialEmpresa, iLote)
    If lErro <> AD_SQL_SUCESSO Then Error 41376

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 41377

    'Lock do Lote
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 41379
    
    tInvLote.iFilialEmpresa = iFilialEmpresa
    tInvLote.iLote = iLote
    
    InvLotePendente_Le = SUCESSO

    Exit Function

Erro_InvLotePendente_Le:

    InvLotePendente_Le = Err

    Select Case Err
        Case 41376, 41377
        Case 41379
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_INVLOTEPENDENTE", Err, tInvLote.iLote, tInvLote.iFilialEmpresa)
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", lErro, Error$, 159534)
    End Select

    Exit Function

End Function

Public Sub Inicializa_Inventario(tInventario As typeInventario, tItemInventario As typeItemInventario)

    tInventario.sCodigo = String(STRING_INVENTARIO_CODIGO, 0)
    tItemInventario.sProduto = String(STRING_PRODUTO, 0)
    tItemInventario.sEtiqueta = String(STRING_INVENTARIO_ETIQUETA, 0)
    tItemInventario.sSiglaUM = String(STRING_UM_SIGLA, 0)
    tItemInventario.sContaContabilEst = String(STRING_CONTA, 0)
    tItemInventario.sContaContabilInv = String(STRING_CONTA, 0)
    tItemInventario.sLoteProduto = String(STRING_LOTE_RASTREAMENTO, 0)

End Sub

Public Sub Inicializa_MovEstoque(objMovEstoque As ClassMovEstoque, iFilialEmpresa As Integer, lCodAuto As Long, tInventario As typeInventario)

    Set objMovEstoque = New ClassMovEstoque

    objMovEstoque.iFilialEmpresa = iFilialEmpresa
    objMovEstoque.dtData = tInventario.dtData
    objMovEstoque.dtHora = tInventario.dHora
    objMovEstoque.lCodigo = lCodAuto
    objMovEstoque.iTipoMov = 0

End Sub

Public Function Atualiza_InvLote(tInvLote As typeInvLote, ByVal iID_Atualizacao As Integer, objAtuInvLoteAux As ClassAtualizacaoInvLoteAux) As Long

Dim lErro As Long
Dim lTransacao As Long
Dim objInvLote As New ClassInvLote

On Error GoTo Erro_Atualiza_InvLote

    'abre transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 41375
    
    objInvLote.iFilialEmpresa = tInvLote.iFilialEmpresa
    objInvLote.iIDAtualizacao = tInvLote.iIDAtualizacao
    objInvLote.iLote = tInvLote.iLote
    objInvLote.iNumItensAtual = tInvLote.iNumItensAtual
    objInvLote.iNumItensInf = tInvLote.iNumItensInf
    objInvLote.sDescricao = tInvLote.sDescricao

    lErro = Atualiza_InvLote0(objInvLote, iID_Atualizacao, objAtuInvLoteAux)
    If lErro <> SUCESSO Then gError 105204

    tInvLote.iFilialEmpresa = objInvLote.iFilialEmpresa
    tInvLote.iIDAtualizacao = objInvLote.iIDAtualizacao
    tInvLote.iLote = objInvLote.iLote
    tInvLote.iNumItensAtual = objInvLote.iNumItensAtual
    tInvLote.iNumItensInf = objInvLote.iNumItensInf
    tInvLote.sDescricao = objInvLote.sDescricao

    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 41392

    Atualiza_InvLote = SUCESSO

    Exit Function

Erro_Atualiza_InvLote:

    Atualiza_InvLote = gErr

    Select Case gErr

        Case 41375
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 41392
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case 105204

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159535)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function

Public Sub Armazena_Inventario(objInventario As ClassInventario, tInventario As typeInventario)
'armazena os dados de tInventario em objInventario

    Set objInventario = New ClassInventario
    Set objInventario.colItens = New ColItemInventario
    
    objInventario.dtData = tInventario.dtData
    objInventario.iFilialEmpresa = tInventario.iFilialEmpresa
    objInventario.iLote = tInventario.iLote
    objInventario.sCodigo = tInventario.sCodigo
    objInventario.dtHora = tInventario.dHora
    

End Sub

Public Function Armazena_ItemInventario(tItemInventario As typeItemInventario, lComando As Long, lComando2 As Long, objInventario As ClassInventario) As Long
'armazena os dados de tItemInventario em objInventario.colItens e exclui os dados da tabela de Inventario Pendente

Dim lErro As Long
Dim objItemInventario As New ClassItemInventario

On Error GoTo Erro_Armazena_ItemInventario


    Set objItemInventario = objInventario.colItens.Add(tItemInventario.lNumIntDoc, tItemInventario.sProduto, tItemInventario.sProdutoDesc, tItemInventario.sSiglaUM, tItemInventario.dQuantidade, tItemInventario.dCusto, tItemInventario.iAlmoxarifado, tItemInventario.sAlmoxarifadoNomeRed, tItemInventario.iTipo, tItemInventario.sEtiqueta, tItemInventario.sContaContabilEst, tItemInventario.sContaContabilInv, 0, 0)
    objItemInventario.dQuantEst = tItemInventario.dQuantEst
    
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 41393

    lErro = Comando_ExecutarPos(lComando2, "DELETE FROM InventarioPendente", lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 41395

    Armazena_ItemInventario = SUCESSO
    
    Exit Function
    
Erro_Armazena_ItemInventario:

    Armazena_ItemInventario = Err

    Select Case Err
    
        Case 41393
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_INVENTARIOPENDENTE", Err, objInventario.sCodigo, objInventario.iFilialEmpresa)
            
        Case 41395
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_INVENTARIOPENDENTE", Err, objInventario.sCodigo, objInventario.iFilialEmpresa)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", lErro, Error$, 159536)
            
    End Select

    Exit Function

End Function

Public Function MovEstoque_Prepara(tItemInventario As typeItemInventario, tEstoqueProduto As typeEstoqueProduto, objMovEstoque As ClassMovEstoque, dFator As Double, objProduto As ClassProduto) As Long

Dim lErro As Long
Dim iTipoMov As Integer
Dim dCalculado As Double, dCusto As Double
Dim objItemMovEstoque As ClassItemMovEstoque
Dim colRastreamentoMovto As New Collection
Dim objItemInventario As New ClassItemInventario

On Error GoTo Erro_MovEstoque_Prepara

    objProduto.sCodigo = tItemInventario.sProduto
    
    lErro = MovEstoque_ObterDados(tItemInventario, tEstoqueProduto, iTipoMov, dCalculado, dFator)
    If lErro <> SUCESSO Then Error 41385

    If dCalculado < 0 Then dCalculado = -dCalculado
        
    'o custo em item inventario é o custo unitário na unidade informada pelo usuario no inventario
    dCusto = tItemInventario.dCusto * dCalculado
    
    objItemInventario.dQuantidade = tItemInventario.dQuantidade
    objItemInventario.dQuantEst = tItemInventario.dQuantEst
    objItemInventario.iFilialOP = tItemInventario.iFilialOP
    objItemInventario.sLote = tItemInventario.sLoteProduto
    objItemInventario.sProduto = tItemInventario.sProduto
    objItemInventario.sSiglaUM = tItemInventario.sSiglaUM
     
    'Move o Rastro do objInventário para a Colecao de Rastreamento
    lErro = Move_Rastro_Inventario_Estoque(objItemInventario, objProduto, colRastreamentoMovto)
    If lErro <> SUCESSO Then Error 78400
    
    dCalculado = dCalculado * dFator
    
    Set objItemMovEstoque = objMovEstoque.colItens.Add(0, iTipoMov, dCusto, 0, tItemInventario.sProduto, tItemInventario.sProdutoDesc, objProduto.sSiglaUMEstoque, dCalculado, tItemInventario.iAlmoxarifado, tItemInventario.sAlmoxarifadoNomeRed, tItemInventario.lNumIntDoc, "", 0, "", "", "", "", 0, colRastreamentoMovto, Nothing, DATA_NULA)
    
    objItemMovEstoque.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_INVENTARIO
        
    MovEstoque_Prepara = SUCESSO

    Exit Function

Erro_MovEstoque_Prepara:

    MovEstoque_Prepara = Err

    Select Case Err
    
        Case 41385, 78400
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", lErro, Error$, 159537)
            
    End Select

    Exit Function

End Function

Public Function Atualiza_Inventario(lComando1 As Long, objInventario As ClassInventario, objItemInventario As ClassItemInventario, objItemMovEstoque As ClassItemMovEstoque) As Long
'atualiza os itens de inventario

Dim lErro As Long
Dim sComando_SQL As String

On Error GoTo Erro_Atualiza_Inventario

    sComando_SQL = "INSERT INTO Inventario (NumIntDoc, FilialEmpresa, Lote, Codigo, Data, Produto, SiglaUM, Quantidade, QuantEst, Custo, Almoxarifado, Etiqueta, Tipo, ContaContabilEst, ContaContabilInv, Hora) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"

    lErro = Comando_Executar(lComando1, sComando_SQL, objItemInventario.lNumIntDoc, objInventario.iFilialEmpresa, objInventario.iLote, objInventario.sCodigo, objInventario.dtData, objItemInventario.sProduto, objItemInventario.sSiglaUM, objItemInventario.dQuantidade, objItemInventario.dQuantEst, objItemMovEstoque.dCusto, objItemInventario.iAlmoxarifado, objItemInventario.sEtiqueta, objItemInventario.iTipo, objItemInventario.sContaContabilEst, objItemInventario.sContaContabilInv, CDbl(objInventario.dtHora))
    If lErro <> AD_SQL_SUCESSO Then Error 41394

    Atualiza_Inventario = SUCESSO

    Exit Function

Erro_Atualiza_Inventario:

    Atualiza_Inventario = Err

    Select Case Err

        Case 41394
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_INVENTARIO", Err, objInventario.sCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", lErro, Error$, 159538)

    End Select

    Exit Function

End Function

Public Function EstoqueProduto_ObterDados(lComando As Long, tItemInventario As typeItemInventario, tEstoqueProduto As typeEstoqueProduto) As Long

Dim lErro As Long
Dim sComando_SQL As String

On Error GoTo Erro_EstoqueProduto_ObterDados

    tEstoqueProduto.sProduto = String(STRING_PRODUTO, 0)

    sComando_SQL = "SELECT Produto, Almoxarifado, DataInventario, SaldoInicial, DataInicial, QuantidadeInicial, QuantDispNossa, QuantRecIndl, QuantIndOutras, " & _
    "QuantDefeituosa, QuantConsig3, QuantDemo3, QuantConserto3, QuantOutras3, QuantBenef3, QuantReservadaConsig FROM EstoqueProduto WHERE Produto = ? AND Almoxarifado = ?"

    lErro = Comando_ExecutarPos(lComando, sComando_SQL, 0, tEstoqueProduto.sProduto, tEstoqueProduto.iAlmoxarifado, tEstoqueProduto.dtDataInventario, tEstoqueProduto.dSaldoInicial, tEstoqueProduto.dtDataInicial, tEstoqueProduto.dQuantidadeInicial, tEstoqueProduto.dQuantDispNossa, tEstoqueProduto.dQuantRecIndl, tEstoqueProduto.dQuantInd, tEstoqueProduto.dQuantDefeituosa, tEstoqueProduto.dQuantConsig3, tEstoqueProduto.dQuantDemo3, tEstoqueProduto.dQuantConserto3, tEstoqueProduto.dQuantOutras3, tEstoqueProduto.dQuantBenef3, tEstoqueProduto.dQuantReservadaConsig, tItemInventario.sProduto, tItemInventario.iAlmoxarifado)
    If lErro <> AD_SQL_SUCESSO Then Error 41405

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 41406

    If lErro = AD_SQL_SEM_DADOS Then Error 41407

    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 41408

    EstoqueProduto_ObterDados = SUCESSO

    Exit Function

Erro_EstoqueProduto_ObterDados:

    EstoqueProduto_ObterDados = Err

    Select Case Err

        Case 41405, 41406
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ESTOQUEPRODUTO", Err, tEstoqueProduto.sProduto, tEstoqueProduto.iAlmoxarifado)
            
        Case 41407
            Call Rotina_Erro(vbOKOnly, "ERRO_ESTOQUEPRODUTO_INEXISTENTE", Err, tEstoqueProduto.sProduto, tEstoqueProduto.iAlmoxarifado)
            
        Case 41408
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_ESTOQUEPRODUTO", Err, tEstoqueProduto.sProduto, tEstoqueProduto.iAlmoxarifado)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", lErro, Error$, 159539)

    End Select

    Exit Function

End Function

Public Function InvLote_Atualiza(lComando As Long, lComando1 As Long, lComando2 As Long, tInvLote As typeInvLote) As Long
'marca o lote como atualizado
Dim sComando_SQL As String
Dim lErro As Long

On Error GoTo Erro_InvLote_Atualiza

    sComando_SQL = "DELETE FROM InvLotePendente"

    lErro = Comando_ExecutarPos(lComando1, sComando_SQL, lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 41409

    sComando_SQL = "INSERT INTO InvLote (FilialEmpresa, Lote, Descricao, NumItensInf, NumItensAtual, IdAtualizacao) VALUES (?,?,?,?,?,?)"

    lErro = Comando_Executar(lComando2, sComando_SQL, tInvLote.iFilialEmpresa, tInvLote.iLote, tInvLote.sDescricao, tInvLote.iNumItensInf, tInvLote.iNumItensAtual, tInvLote.iIDAtualizacao)
    If lErro <> AD_SQL_SUCESSO Then Error 41410

    InvLote_Atualiza = SUCESSO

    Exit Function

Erro_InvLote_Atualiza:

    InvLote_Atualiza = Err

    Select Case Err

        Case 41409
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_INVLOTEPENDENTE", Err, tInvLote.iLote, tInvLote.iFilialEmpresa)
        Case 41410
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_INVLOTE", Err, tInvLote.iLote, tInvLote.iFilialEmpresa)
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", lErro, Error$, 159540)

    End Select

    Exit Function

End Function

Private Function MovEstoque_ObterDados(tItemInventario As typeItemInventario, tEstoqueProduto As typeEstoqueProduto, iTipoMov As Integer, dCalculado As Double, ByVal dFator As Double) As Long
' prepara os itens de movimentação de estoque

Dim lErro As Long

On Error GoTo Erro_MovEstoque_ObterDados

    dCalculado = tItemInventario.dQuantidade - tItemInventario.dQuantEst

    If tItemInventario.iTipo = TIPO_QUANT_DISPONIVEL_NOSSA Then

        If tItemInventario.iAtualizaSoLote = INVENTARIO_ATUALIZA_SO_LOTE Then

            If dCalculado >= 0 Then
                iTipoMov = MOV_EST_ACRES_INVENT_DISP_NOSSA_SOLOTE
            ElseIf dCalculado < 0 Then
                iTipoMov = MOV_EST_DECR_INVENT_DISP_NOSSA_SOLOTE
            End If

        Else

            If dCalculado >= 0 Then
                iTipoMov = MOV_EST_ACRESCIMO_INVENT_DISPONIVEL_NOSSA
            ElseIf dCalculado < 0 Then
                iTipoMov = MOV_EST_DECRESCIMO_INVENT_DISPONIVEL_NOSSA
            End If

        End If

    ElseIf tItemInventario.iTipo = TIPO_QUANT_RECEB_INDISP Then

        If tItemInventario.iAtualizaSoLote = INVENTARIO_ATUALIZA_SO_LOTE Then

            If dCalculado >= 0 Then
                iTipoMov = MOV_EST_ACRES_INVENT_RECEB_IND_SOLOTE
            ElseIf dCalculado < 0 Then
                iTipoMov = MOV_EST_DECR_INVENT_RECEB_IND_SOLOTE
            End If

        Else

            If dCalculado >= 0 Then
                iTipoMov = MOV_EST_ACRESCIMO_INVENT_RECEB_INDISP
            ElseIf dCalculado < 0 Then
                iTipoMov = MOV_EST_DECRESCIMO_INVENT_RECEB_INDISP
            End If

        End If

    ElseIf tItemInventario.iTipo = TIPO_QUANT_OUTRAS_INDISP Then

        If tItemInventario.iAtualizaSoLote = INVENTARIO_ATUALIZA_SO_LOTE Then
        
            If dCalculado >= 0 Then
                iTipoMov = MOV_EST_ACRES_INVENT_IND_OUTRAS_SOLOTE
            ElseIf dCalculado < 0 Then
                iTipoMov = MOV_EST_DECR_INVENT_IND_OUTRAS_SOLOTE
            End If

        Else

            If dCalculado >= 0 Then
                iTipoMov = MOV_EST_ACRESCIMO_INVENT_INDISP_OUTRAS
            ElseIf dCalculado < 0 Then
                iTipoMov = MOV_EST_DECRESCIMO_INVENT_INDISP_OUTRAS
            End If

        End If

    ElseIf tItemInventario.iTipo = TIPO_QUANT_DEFEIT Then

        If tItemInventario.iAtualizaSoLote = INVENTARIO_ATUALIZA_SO_LOTE Then
        
            If dCalculado >= 0 Then
                iTipoMov = MOV_EST_ACRES_INVENT_DEFEITUOSO_SOLOTE
            ElseIf dCalculado < 0 Then
                iTipoMov = MOV_EST_DECR_INVENT_DEFEITUOSO_SOLOTE
            End If

        Else
        
            If dCalculado >= 0 Then
                iTipoMov = MOV_EST_ACRESCIMO_INVENT_DEFEITUOSO
            ElseIf dCalculado < 0 Then
                iTipoMov = MOV_EST_DECRESCIMO_INVENT_DEFEITUOSO
            End If

        End If
        
    ElseIf tItemInventario.iTipo = TIPO_QUANT_3_CONSIG Then

        If tItemInventario.iAtualizaSoLote = INVENTARIO_ATUALIZA_SO_LOTE Then

            If dCalculado >= 0 Then
                iTipoMov = MOV_EST_ACRES_INVENT_CONSIG_TERC_SOLOTE
            ElseIf dCalculado < 0 Then
                iTipoMov = MOV_EST_DECR_INVENT_CONSIG_TERC_SOLOTE
            End If

        Else
        
            If dCalculado >= 0 Then
                iTipoMov = MOV_EST_ACRESCIMO_INVENT_CONSIG_TERC
            ElseIf dCalculado < 0 Then
                iTipoMov = MOV_EST_DECRESCIMO_INVENT_CONSIG_TERC
            End If

        End If

    ElseIf tItemInventario.iTipo = TIPO_QUANT_3_DEMO Then

        If tItemInventario.iAtualizaSoLote = INVENTARIO_ATUALIZA_SO_LOTE Then

            If dCalculado >= 0 Then
                iTipoMov = MOV_EST_ACRES_INVENT_DEMO_TERC_SOLOTE
            ElseIf dCalculado < 0 Then
                iTipoMov = MOV_EST_DECR_INVENT_DEMO_TERC_SOLOTE
            End If

        Else
        
            If dCalculado >= 0 Then
                iTipoMov = MOV_EST_ACRESCIMO_INVENT_DEMO_TERC
            ElseIf dCalculado < 0 Then
                iTipoMov = MOV_EST_DECRESCIMO_INVENT_DEMO_TERC
            End If
        
        End If
        
    ElseIf tItemInventario.iTipo = TIPO_QUANT_3_CONSERTO Then

        If tItemInventario.iAtualizaSoLote = INVENTARIO_ATUALIZA_SO_LOTE Then

            If dCalculado >= 0 Then
                iTipoMov = MOV_EST_ACRES_INVENT_CONS_TERC_SOLOTE
            ElseIf dCalculado < 0 Then
                iTipoMov = MOV_EST_DECR_INVENT_CONS_TERC_SOLOTE
            End If

        Else
        
            If dCalculado >= 0 Then
                iTipoMov = MOV_EST_ACRESCIMO_INVENT_CONSERTO_TERC
            ElseIf dCalculado < 0 Then
                iTipoMov = MOV_EST_DECRESCIMO_INVENT_CONSERTO_TERC
            End If
        
        End If
        
    ElseIf tItemInventario.iTipo = TIPO_QUANT_3_OUTRAS Then

        If tItemInventario.iAtualizaSoLote = INVENTARIO_ATUALIZA_SO_LOTE Then

            If dCalculado >= 0 Then
                iTipoMov = MOV_EST_ACRES_INVENT_OUTROS_TERC_SOLOTE
            ElseIf dCalculado < 0 Then
                iTipoMov = MOV_EST_DECR_INVENT_OUTROS_TERC_SOLOTE
            End If

        Else
        
            If dCalculado >= 0 Then
                iTipoMov = MOV_EST_ACRESCIMO_INVENT_OUTROS_TERC
            ElseIf dCalculado < 0 Then
                iTipoMov = MOV_EST_DECRESCIMO_INVENT_OUTROS_TERC
            End If
        
        End If

    ElseIf tItemInventario.iTipo = TIPO_QUANT_3_BENEF Then

        If tItemInventario.iAtualizaSoLote = INVENTARIO_ATUALIZA_SO_LOTE Then

            If dCalculado >= 0 Then
                iTipoMov = MOV_EST_ACRES_INVENT_BENEF_TERC_SOLOTE
            ElseIf dCalculado < 0 Then
                iTipoMov = MOV_EST_DECR_INVENT_BENEF_TERC_SOLOTE
            End If

        Else
        
            If dCalculado >= 0 Then
                iTipoMov = MOV_EST_ACRESCIMO_INVENT_BENEF_TERC
            ElseIf dCalculado < 0 Then
                iTipoMov = MOV_EST_DECRESCIMO_INVENT_BENEF_TERC
            End If
        
        End If
        
    ElseIf tItemInventario.iTipo = TIPO_QUANT_DISPONIVEL_NOSSA_CI Then

        If tItemInventario.iAtualizaSoLote = INVENTARIO_ATUALIZA_SO_LOTE Then

            If dCalculado >= 0 Then
                iTipoMov = MOV_EST_ACRES_INVENT_DISP_NOSSA_SOLOTE_CI
            ElseIf dCalculado < 0 Then
                iTipoMov = MOV_EST_DECR_INVENT_DISP_NOSSA_SOLOTE_CI
            End If

        Else

            If dCalculado >= 0 Then
                iTipoMov = MOV_EST_ACRESCIMO_INVENT_DISPONIVEL_NOSSA_CI
            ElseIf dCalculado < 0 Then
                iTipoMov = MOV_EST_DECRESCIMO_INVENT_DISPONIVEL_NOSSA_CI
            End If

        End If

    ElseIf tItemInventario.iTipo = TIPO_QUANT_RECEB_INDISP_CI Then

        If tItemInventario.iAtualizaSoLote = INVENTARIO_ATUALIZA_SO_LOTE Then

            If dCalculado >= 0 Then
                iTipoMov = MOV_EST_ACRES_INVENT_RECEB_IND_SOLOTE_CI
            ElseIf dCalculado < 0 Then
                iTipoMov = MOV_EST_DECR_INVENT_RECEB_IND_SOLOTE_CI
            End If

        Else

            If dCalculado >= 0 Then
                iTipoMov = MOV_EST_ACRESCIMO_INVENT_RECEB_INDISP_CI
            ElseIf dCalculado < 0 Then
                iTipoMov = MOV_EST_DECRESCIMO_INVENT_RECEB_INDISP_CI
            End If

        End If

    ElseIf tItemInventario.iTipo = TIPO_QUANT_OUTRAS_INDISP_CI Then

        If tItemInventario.iAtualizaSoLote = INVENTARIO_ATUALIZA_SO_LOTE Then
        
            If dCalculado >= 0 Then
                iTipoMov = MOV_EST_ACRES_INVENT_IND_OUTRAS_SOLOTE_CI
            ElseIf dCalculado < 0 Then
                iTipoMov = MOV_EST_DECR_INVENT_IND_OUTRAS_SOLOTE_CI
            End If

        Else

            If dCalculado >= 0 Then
                iTipoMov = MOV_EST_ACRESCIMO_INVENT_INDISP_OUTRAS_CI
            ElseIf dCalculado < 0 Then
                iTipoMov = MOV_EST_DECRESCIMO_INVENT_INDISP_OUTRAS_CI
            End If

        End If

    ElseIf tItemInventario.iTipo = TIPO_QUANT_DEFEIT_CI Then

        If tItemInventario.iAtualizaSoLote = INVENTARIO_ATUALIZA_SO_LOTE Then
        
            If dCalculado >= 0 Then
                iTipoMov = MOV_EST_ACRES_INVENT_DEFEITUOSO_SOLOTE_CI
            ElseIf dCalculado < 0 Then
                iTipoMov = MOV_EST_DECR_INVENT_DEFEITUOSO_SOLOTE_CI
            End If

        Else
        
            If dCalculado >= 0 Then
                iTipoMov = MOV_EST_ACRESCIMO_INVENT_DEFEITUOSO_CI
            ElseIf dCalculado < 0 Then
                iTipoMov = MOV_EST_DECRESCIMO_INVENT_DEFEITUOSO_CI
            End If

        End If
        
    ElseIf tItemInventario.iTipo = TIPO_QUANT_3_CONSIG_CI Then

        If tItemInventario.iAtualizaSoLote = INVENTARIO_ATUALIZA_SO_LOTE Then

            If dCalculado >= 0 Then
                iTipoMov = MOV_EST_ACRES_INVENT_CONSIG_TERC_SOLOTE_CI
            ElseIf dCalculado < 0 Then
                iTipoMov = MOV_EST_DECR_INVENT_CONSIG_TERC_SOLOTE_CI
            End If

        Else
        
            If dCalculado >= 0 Then
                iTipoMov = MOV_EST_ACRESCIMO_INVENT_CONSIG_TERC_CI
            ElseIf dCalculado < 0 Then
                iTipoMov = MOV_EST_DECRESCIMO_INVENT_CONSIG_TERC_CI
            End If

        End If
        
    End If

    MovEstoque_ObterDados = SUCESSO

    Exit Function

Erro_MovEstoque_ObterDados:

    MovEstoque_ObterDados = Err

    Select Case Err
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", lErro, Error$, 159541)
            
    End Select

    Exit Function

End Function

Public Function EstoqueProduto_AtualizaDataInventario(lComando As Long, lComando1 As Long, dtData As Date) As Long
'grava a data do ultimo inventario do produto anterior em estoqueproduto
Dim lErro As Long
Dim sComando_SQL As String

On Error GoTo Erro_EstoqueProduto_AtualizaDataInventario

    sComando_SQL = "UPDATE EstoqueProduto SET DataInventario = ?"

    lErro = Comando_ExecutarPos(lComando1, sComando_SQL, lComando, dtData)
    If lErro <> AD_SQL_SUCESSO Then Error 41412

    lErro = Comando_Unlock(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 41413

    EstoqueProduto_AtualizaDataInventario = SUCESSO

    Exit Function

Erro_EstoqueProduto_AtualizaDataInventario:

    EstoqueProduto_AtualizaDataInventario = Err

    Select Case Err

        Case 41412
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_ESTOQUEPRODUTO1", Err)
        Case 41413
            Call Rotina_Erro(vbOKOnly, "ERRO_UNLOCK_ESTOQUEPRODUTO", Err)
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", lErro, Error$, 159542)

    End Select

    Exit Function

End Function

'???? Fernando está Função já está em Class Mat Grava como private
Private Function Move_Rastro_Inventario_Estoque(objItemInventario As ClassItemInventario, objProduto As ClassProduto, colRastreamentoMovto As Collection) As Long
'Move o Rastro do Inventario para a coleção que será passada no Reg. Inventário

Dim lErro As Long
Dim objRastreamentoMovto As New ClassRastreamentoMovto

On Error GoTo Erro_Move_Rastro_Inventario_Estoque
    
    If objProduto.iRastro <> PRODUTO_RASTRO_NENHUM And Len(Trim(objItemInventario.sLote)) > 0 Then
    
        If objProduto.iRastro = PRODUTO_RASTRO_LOTE Then
            
            objRastreamentoMovto.sLote = objItemInventario.sLote
            
        ElseIf objProduto.iRastro = PRODUTO_RASTRO_OP Then
            
            objRastreamentoMovto.sLote = objItemInventario.sLote
            objRastreamentoMovto.iFilialOP = objItemInventario.iFilialOP
            
        End If
        
        objRastreamentoMovto.iTipoDocOrigem = TIPO_RASTREAMENTO_MOVTO_MOVTO_ESTOQUE
        objRastreamentoMovto.sProduto = objItemInventario.sProduto
        objRastreamentoMovto.dQuantidade = Abs(objItemInventario.dQuantidade - objItemInventario.dQuantEst)
        objRastreamentoMovto.sSiglaUM = objItemInventario.sSiglaUM
            
        colRastreamentoMovto.Add objRastreamentoMovto
        
    End If
    
    Move_Rastro_Inventario_Estoque = SUCESSO
    
    Exit Function
    
Erro_Move_Rastro_Inventario_Estoque:

    Move_Rastro_Inventario_Estoque = gErr
    
    Select Case gErr
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 159543)
    
    End Select
    
    Exit Function
    
End Function

Public Function Atualiza_InvLote0(objInvLote As ClassInvLote, ByVal iID_Atualizacao As Integer, objAtuInvLoteAux As ClassAtualizacaoInvLoteAux) As Long

Dim lErro As Long

On Error GoTo Erro_Atualiza_InvLote0

    lErro = CF("Atualiza_InvLote_Trans", objInvLote, iID_Atualizacao, objAtuInvLoteAux)
    If lErro <> SUCESSO Then gError 105205

    Atualiza_InvLote0 = SUCESSO

    Exit Function

Erro_Atualiza_InvLote0:

    Atualiza_InvLote0 = gErr

    Select Case gErr

        Case 105205

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159544)

    End Select

    Exit Function

End Function



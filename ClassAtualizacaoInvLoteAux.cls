VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassAtualizacaoInvLoteAux"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarobjTelaAtualizacao As Object 'tela de acompanhamento da atualizacao de arquivos. Necessaria para a contabilizacao.
Private mvardtContabil As Date
Private mvariLoteContabil As Integer
Private giExercicio As Integer, giPeriodo As Integer
Private gcolAlmoxFilial As New Collection

Private gobjContabAutomatica As ClassContabAutomatica
Private gobjInventario As ClassInventario
Private gobjMovEstoque As ClassMovEstoque

'Mnemonicos
Private Const CODIGO1 As String = "Codigo"
Private Const DATA1 As String = "Data"
Private Const PRODUTO1 As String = "Produto_Codigo"
Private Const UNIDADE_MED As String = "Unidade_Med"
Private Const QUANTIDADE1 As String = "Quantidade"
Private Const QUANT_DIFERENCA As String = "QuantDiferenca"
Private Const DESCRICAO_ITEM As String = "Descricao_Item"
Private Const ALMOXARIFADO1 As String = "Almoxarifado"
Private Const TIPO1 As String = "Tipo"
Private Const CUSTO_UNITARIO As String = "Custo_Unitario"
Private Const ETIQUETA1 As String = "Etiqueta"
Private Const CONTACONTABIL1 As String = "ContaContabilEst"
Private Const CTAAJUSTEINV As String = "CtaAjusteInv"
Private Const QUANT_ESTOQUE As String = "Quant_Estoque"

Public Property Let iLoteContabil(ByVal vData As Integer)
    mvariLoteContabil = vData
End Property

Public Property Get iLoteContabil() As Integer
    iLoteContabil = mvariLoteContabil
End Property

Public Property Let dtContabil(ByVal vData As Date)
    mvardtContabil = vData
End Property

Public Property Get dtContabil() As Date
    dtContabil = mvardtContabil
End Property

Public Property Let objTelaAtualizacao(ByVal vData As Object)
    Set mvarobjTelaAtualizacao = vData
End Property

Public Property Get objTelaAtualizacao() As Object
    Set objTelaAtualizacao = mvarobjTelaAtualizacao
End Property

Public Function Calcula_Mnemonico(objMnemonicoValor As ClassMnemonicoValor) As Long

Dim lErro As Long
Dim sTipo As String
Dim sProdutoTela As String, sAlmoxNomeRed As String
Dim objItemInventario As ClassItemInventario
Dim objAlmoxarifado As ClassAlmoxarifado
Dim sContaMascarada As String

On Error GoTo Erro_Calcula_Mnemonico

    Select Case objMnemonicoValor.sMnemonico

        Case QUANT_ESTOQUE
            For Each objItemInventario In gobjInventario.colItens

                objMnemonicoValor.colValor.Add objItemInventario.dQuantEst
            
            Next
            
        Case CODIGO1
            objMnemonicoValor.colValor.Add gobjInventario.sCodigo

        Case DATA1
            objMnemonicoValor.colValor.Add gobjInventario.dtData

        Case PRODUTO1
            For Each objItemInventario In gobjInventario.colItens

                lErro = Mascara_RetornaProdutoTela(objItemInventario.sProduto, sProdutoTela)
                If lErro <> SUCESSO Then Error 59420
                    
                objMnemonicoValor.colValor.Add sProdutoTela
            
            Next

         Case ESCANINHO_CUSTO
            For Each objItemInventario In gobjInventario.colItens
            
                Select Case objItemInventario.iTipo
                
                    Case TIPO_QUANT_3_CONSIG, TIPO_QUANT_3_CONSIG_CI
                        objMnemonicoValor.colValor.Add ESCANINHO_3_EM_CONSIGNACAO
                        
                    Case TIPO_QUANT_3_DEMO
                        objMnemonicoValor.colValor.Add ESCANINHO_3_EM_DEMO
                    
                    Case TIPO_QUANT_3_OUTRAS
                        objMnemonicoValor.colValor.Add ESCANINHO_3_EM_OUTROS
                    
                    Case TIPO_QUANT_3_BENEF
                        objMnemonicoValor.colValor.Add ESCANINHO_3_EM_BENEF
                    
                    Case Else
                        objMnemonicoValor.colValor.Add ESCANINHO_NOSSO
                        
                End Select
                
            Next
            
        Case UNIDADE_MED
            For Each objItemInventario In gobjInventario.colItens

                objMnemonicoValor.colValor.Add objItemInventario.sSiglaUM
            
            Next

        Case QUANTIDADE1
            For Each objItemInventario In gobjInventario.colItens

                objMnemonicoValor.colValor.Add objItemInventario.dQuantidade
            
            Next

        Case QUANT_DIFERENCA
            For Each objItemInventario In gobjInventario.colItens

                objMnemonicoValor.colValor.Add (objItemInventario.dQuantidade - objItemInventario.dQuantEst)
            
            Next

        Case DESCRICAO_ITEM
            For Each objItemInventario In gobjInventario.colItens

                objMnemonicoValor.colValor.Add objItemInventario.sProdutoDesc
            
            Next

        Case ALMOXARIFADO1
            
            If gcolAlmoxFilial.Count = 0 Then
            
                lErro = CF("Almoxarifados_Le_FilialEmpresa", gobjInventario.iFilialEmpresa, gcolAlmoxFilial)
                If lErro <> SUCESSO Then Error 59421
                
            End If
            
            For Each objItemInventario In gobjInventario.colItens

                sAlmoxNomeRed = ""
                
                For Each objAlmoxarifado In gcolAlmoxFilial
                
                    If objItemInventario.iAlmoxarifado = objAlmoxarifado.iCodigo Then
                    
                        sAlmoxNomeRed = objAlmoxarifado.sNomeReduzido
                        Exit For
                        
                    End If
                    
                Next
                
                objMnemonicoValor.colValor.Add sAlmoxNomeRed
            
            Next

        Case TIPO1
            For Each objItemInventario In gobjInventario.colItens

                Select Case objItemInventario.iTipo
                
                    Case TIPO_QUANT_DISPONIVEL_NOSSA
                        sTipo = STRING_QUANT_DISPONIVEL_NOSSA
                    
                    Case TIPO_QUANT_RECEB_INDISP
                        sTipo = STRING_QUANT_RECEB_INDISP
                    
                    Case TIPO_QUANT_OUTRAS_INDISP
                        sTipo = STRING_QUANT_OUTRAS_INDISP
                    
                    Case TIPO_QUANT_DEFEIT
                        sTipo = STRING_QUANT_DEFEIT
                    
                    Case TIPO_QUANT_3_CONSIG
                        sTipo = STRING_QUANT_3_CONSIG
                    
                    Case TIPO_QUANT_3_DEMO
                        sTipo = STRING_QUANT_3_DEMO
                        
                    Case TIPO_QUANT_3_CONSERTO
                        sTipo = STRING_QUANT_3_CONSERTO
                    
                    Case TIPO_QUANT_3_OUTRAS
                        sTipo = STRING_QUANT_3_OUTRAS
                    
                    Case TIPO_QUANT_3_BENEF
                        sTipo = STRING_QUANT_3_BENEF
                
                    Case TIPO_QUANT_DISPONIVEL_NOSSA_CI
                        sTipo = STRING_QUANT_DISPONIVEL_NOSSA_CI
                    
                    Case TIPO_QUANT_RECEB_INDISP_CI
                        sTipo = STRING_QUANT_RECEB_INDISP_CI
                    
                    Case TIPO_QUANT_OUTRAS_INDISP_CI
                        sTipo = STRING_QUANT_OUTRAS_INDISP_CI
                    
                    Case TIPO_QUANT_DEFEIT_CI
                        sTipo = STRING_QUANT_DEFEIT_CI
                    
                    Case TIPO_QUANT_3_CONSIG_CI
                        sTipo = STRING_QUANT_3_CONSIG_CI
                
                End Select
                
                objMnemonicoValor.colValor.Add sTipo
            
            Next

        Case CUSTO_UNITARIO
            For Each objItemInventario In gobjInventario.colItens

                objMnemonicoValor.colValor.Add objItemInventario.dCusto
            
            Next

        Case ETIQUETA1
            For Each objItemInventario In gobjInventario.colItens

                objMnemonicoValor.colValor.Add objItemInventario.sEtiqueta
            
            Next

        Case CONTACONTABIL1
            For Each objItemInventario In gobjInventario.colItens

                If Len(Trim(objItemInventario.sContaContabilEst)) <> 0 Then
                                
                    sContaMascarada = String(STRING_CONTA, 0)

                    lErro = Mascara_RetornaContaTela(objItemInventario.sContaContabilEst, sContaMascarada)
                    If lErro <> SUCESSO Then Error 59426

                    objMnemonicoValor.colValor.Add sContaMascarada
                    
                Else
                
                    objMnemonicoValor.colValor.Add ""
                    
                End If
            
            Next

        Case CTAAJUSTEINV
            For Each objItemInventario In gobjInventario.colItens

                If Len(Trim(objItemInventario.sContaContabilInv)) <> 0 Then
                                
                    sContaMascarada = String(STRING_CONTA, 0)

                    lErro = Mascara_RetornaContaTela(objItemInventario.sContaContabilInv, sContaMascarada)
                    If lErro <> SUCESSO Then Error 59427

                    objMnemonicoValor.colValor.Add sContaMascarada
                    
                Else
                
                    objMnemonicoValor.colValor.Add ""
                    
                End If
            
            Next

        Case Else
            Error 59422

    End Select

    Calcula_Mnemonico = SUCESSO

    Exit Function

Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = Err

    Select Case Err

        Case 59422
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO

        Case 59420, 59421, 59426, 59427

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 144643)

    End Select

    Exit Function

End Function

Public Function GeraContabilizacao(objContabAutomatica As ClassContabAutomatica, vParams As Variant) As Long
'esta funcao é chamada a cada atualizacao de inventario e é responsavel por gerar a contabilizacao correspondente

Dim lErro As Long, lDoc As Long, objItem As ClassItemNF, iNumAlocacoes As Integer
Dim objPeriodo As New ClassPeriodo

On Error GoTo Erro_GeraContabilizacao

    Set gobjContabAutomatica = objContabAutomatica
    Set gobjInventario = vParams(0)
    Set gobjMovEstoque = vParams(1)

    mvarobjTelaAtualizacao.Controls("GridItens").Tag = gobjInventario.colItens.Count

    If giPeriodo = 0 Then
    
        'Coloca o periodo relativo a data na tela
        lErro = CF("Periodo_Le", dtContabil, objPeriodo)
        If lErro <> SUCESSO Then Error 59419
        
        giPeriodo = objPeriodo.iPeriodo
        giExercicio = objPeriodo.iExercicio
    
    End If

    'obter numero de Doc
    lErro = CF("Voucher_Automatico_Trans", gobjInventario.iFilialEmpresa, giExercicio, giPeriodo, MODULO_ESTOQUE, lDoc)
    If lErro <> SUCESSO Then Error 59417
    
    If gobjMovEstoque.colItens.Count > 0 Then
    
        'grava a contabilizacao
        lErro = objContabAutomatica.Gravar_Registro(mvarobjTelaAtualizacao, "InventarioLote", gobjMovEstoque.colItens(1).lNumIntDoc, 0, 0, LANPENDENTE_NAO_APROPR_CRPROD, lDoc, gobjInventario.iFilialEmpresa, mvariLoteContabil)
        If lErro <> SUCESSO Then Error 59418

    End If
    
    GeraContabilizacao = SUCESSO
     
    Exit Function
    
Erro_GeraContabilizacao:

    GeraContabilizacao = Err
     
    Select Case Err
          
        Case 59417, 59418, 59419
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 144644)
     
    End Select
     
    Exit Function

End Function

Private Sub Class_Terminate()

    Set gobjContabAutomatica = Nothing
    Set gobjInventario = Nothing
    Set gobjMovEstoque = Nothing
    Set gcolAlmoxFilial = Nothing
    
End Sub

Public Function Atualiza_InvLote_Trans(objInvLote As ClassInvLote, ByVal iID_Atualizacao As Integer, objAtuInvLoteAux As ClassAtualizacaoInvLoteAux) As Long

Dim lErro As Long, lErro1 As Long
Dim alComando(0 To 10) As Long, lCodAuto As Long
Dim tInvLote1 As typeInvLote, tEstoqueProduto As typeEstoqueProduto
Dim iIndice As Integer, objMovEstoque As ClassMovEstoque
Dim sComando_SQL As String, sCodigoAtual As String
Dim tItemInventario As typeItemInventario, tInventario As typeInventario
Dim objInventario As ClassInventario
Dim objItemMovEstoque As ClassItemMovEstoque
Dim objItemInventario As ClassItemInventario
Dim dFator As Double, objContabAutomatica As ClassContabAutomatica
Dim objProduto As New ClassProduto

On Error GoTo Erro_Atualiza_InvLote_Trans

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 41374
    Next

    GL_lUltimoErro = 0

    lErro = InvLotePendente_Le(alComando(0), tInvLote1, objInvLote.iFilialEmpresa, objInvLote.iLote)
    If lErro <> SUCESSO And lErro <> 41379 Then gError 41419

    'Verifica se o lote deve ser atualizado por este processo
    If iID_Atualizacao <> tInvLote1.iIDAtualizacao Then gError 41378

    Call Inicializa_Inventario(tInventario, tItemInventario)

    'pesquisa os inventarios pendentes do lote ordenados por Codigo
    sComando_SQL = "SELECT NumIntDoc, FilialEmpresa, Lote, Codigo, Data, Produto, SiglaUM, Quantidade, QuantEst, Custo, Almoxarifado, Etiqueta, Tipo, ContaContabilEst, ContaContabilInv, LoteProduto, FilialOP, Hora, AtualizaSoLote FROM InventarioPendente WHERE FilialEmpresa = ? AND Lote = ? ORDER BY Codigo"
    
    lErro1 = Comando_ExecutarPos(alComando(3), sComando_SQL, 0, tItemInventario.lNumIntDoc, tInventario.iFilialEmpresa, tInventario.iLote, tInventario.sCodigo, tInventario.dtData, tItemInventario.sProduto, tItemInventario.sSiglaUM, tItemInventario.dQuantidade, tItemInventario.dQuantEst, tItemInventario.dCusto, tItemInventario.iAlmoxarifado, tItemInventario.sEtiqueta, tItemInventario.iTipo, tItemInventario.sContaContabilEst, tItemInventario.sContaContabilInv, tItemInventario.sLoteProduto, tItemInventario.iFilialOP, tInventario.dHora, tItemInventario.iAtualizaSoLote, objInvLote.iFilialEmpresa, objInvLote.iLote)
    If lErro1 <> AD_SQL_SUCESSO Then gError 41381

    'Le o primeiro inventario pendente
    lErro1 = Comando_BuscarPrimeiro(alComando(3))
    If lErro1 <> AD_SQL_SUCESSO And lErro1 <> AD_SQL_SEM_DADOS Then gError 41382

    Do While lErro1 <> AD_SQL_SEM_DADOS

        sCodigoAtual = tInventario.sCodigo

        'retorna o próximo codigo de Movimento de Estoque disponivel
        lErro = CF("MovEstoque_Automatico_EmTransacao1", alComando(1), alComando(2), alComando(8), objInvLote.iFilialEmpresa, lCodAuto)
        If lErro <> SUCESSO Then gError 41380

        Call Inicializa_MovEstoque(objMovEstoque, objInvLote.iFilialEmpresa, lCodAuto, tInventario)
        
        'armazena os dados de tInventario em objInventario
        Call Armazena_Inventario(objInventario, tInventario)

        Do While sCodigoAtual = tInventario.sCodigo And lErro1 = AD_SQL_SUCESSO

            'pesquisa os dados referentes ao estoque do produto
            lErro = EstoqueProduto_ObterDados(alComando(4), tItemInventario, tEstoqueProduto)
            If lErro <> SUCESSO Then gError 41383
            
'            If tInventario.dtData < tEstoqueProduto.dtDataInventario Then gError 55464

            'grava a data do ultimo inventario do produto anterior em estoqueproduto
            lErro = EstoqueProduto_AtualizaDataInventario(alComando(4), alComando(6), tInventario.dtData)
            If lErro <> SUCESSO Then gError 41387

            objProduto.sCodigo = tItemInventario.sProduto
    
            'Lê o produto no BD para obter UM de estoque
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 20843
    
            If lErro = 28030 Then gError 20844
            
            lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, tItemInventario.sSiglaUM, objProduto.sSiglaUMEstoque, dFator)
            If lErro <> SUCESSO Then gError 20845
    
            'prepara os itens de movimentação de estoque
            lErro = MovEstoque_Prepara(tItemInventario, tEstoqueProduto, objMovEstoque, dFator, objProduto)
            If lErro <> SUCESSO Then gError 41391
            
            'se não mencionar o lote ==> verificar se a quantidade encontrada é menor do que a quantidade em reserva, se for ==> erro (ou seja, tem que abrir mão das reservas antes).
            'se mencionar lote, não pode verificar a quantidade vs a quantidade reservada pois o produto pode estar distribuido por varios lotes. Para isto, existe um relatório que lista os produtos que tem no estoque quantidade abaixo da reservada.
            'o usuário poderá acertar esta situação pelo inventário que permite alterar só o estoque, o estoque junto com o lote ou só o lote.
            If gobjMAT.iAceitaEstoqueNegativo = DESMARCADO Then
            
                If Len(Trim(tItemInventario.sLoteProduto)) = 0 Then
                    If tItemInventario.iTipo = TIPO_QUANT_DISPONIVEL_NOSSA Or tItemInventario.iTipo = TIPO_QUANT_DISPONIVEL_NOSSA_CI Then
                        If (tItemInventario.dQuantidade * dFator) - tEstoqueProduto.dQuantReservada < 0 Then gError 61390
                    ElseIf tItemInventario.iTipo = TIPO_QUANT_3_CONSIG Or tItemInventario.iTipo = TIPO_QUANT_3_CONSIG_CI Then
                        If (tItemInventario.dQuantidade * dFator) - tEstoqueProduto.dQuantReservadaConsig < 0 Then gError 61391
                    End If
                End If
                
            End If
   
            'armazena os dados de tItemInventario em objInventario.colItens e exclui os dados da tabela de Inventario Pendente
            lErro = Armazena_ItemInventario(tItemInventario, alComando(3), alComando(9), objInventario)
            If lErro <> SUCESSO Then gError 55473

            'le o proximo inventario
            lErro1 = Comando_BuscarProximo(alComando(3))
            If lErro1 <> AD_SQL_SUCESSO And lErro1 <> AD_SQL_SEM_DADOS Then gError 41386

        Loop

        'grava os movimentos de estoque
        lErro = CF("MovEstoque_Grava", objMovEstoque)
        If lErro <> SUCESSO Then gError 41389

        For iIndice = 1 To objInventario.colItens.Count
    
            Set objItemInventario = objInventario.colItens(iIndice)
            Set objItemMovEstoque = objMovEstoque.colItens(iIndice)

            'atualiza os itens de inventario
            lErro = Atualiza_Inventario(alComando(5), objInventario, objItemInventario, objItemMovEstoque)
            If lErro <> SUCESSO Then gError 41384

        Next

        If (gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO) Then
        
            Set objContabAutomatica = New ClassContabAutomatica
        
            lErro = objContabAutomatica.Inicializa_Contab(objAtuInvLoteAux, MODULO_ESTOQUE, "EST", objAtuInvLoteAux.dtContabil, objMovEstoque.dtData)
            If lErro <> SUCESSO Then gError 59415
            
            lErro = objContabAutomatica.GeraContabilizacao(objInventario, objMovEstoque)
            If lErro <> SUCESSO Then gError 44258
    
            lErro = objContabAutomatica.Finaliza_Contab()
            If lErro <> SUCESSO Then gError 59416
        
        End If

    Loop

    'atualiza o lote
    lErro = InvLote_Atualiza(alComando(0), alComando(7), alComando(10), tInvLote1)
    If lErro <> SUCESSO Then gError 41388
    
    If GL_lUltimoErro <> 0 Then gError 105122
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Atualiza_InvLote_Trans = SUCESSO

    Exit Function

Erro_Atualiza_InvLote_Trans:

    Atualiza_InvLote_Trans = gErr

    Select Case gErr

        Case 20843, 20845, 41380, 41383, 41384, 41387, 41388, 41389, 41391, 41423, 44258, 55465, 55473, 59415, 59416
        
        Case 20844
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case 41374
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 41375
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 41378
            Call Rotina_Erro(vbOKOnly, "ERRO_IMPOSSIVEL_ATUALIZACAO", gErr, objInvLote.iLote, objInvLote.iFilialEmpresa)
            
        Case 41381, 41382, 41386
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_INVENTARIOPENDENTE", gErr, objInvLote.iLote, objInvLote.iFilialEmpresa)
            
        Case 41392
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case 41419
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_INVLOTEPENDENTE", gErr, objInvLote.iLote, objInvLote.iFilialEmpresa)
            
        Case 55464
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INVENTARIO_MENOR", gErr, tItemInventario.sProduto, CStr(tItemInventario.iAlmoxarifado), CStr(tEstoqueProduto.dtDataInventario), CStr(tInventario.dtData))
            
        Case 61390
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INV_LIBERAR_QUANTIDADE_RESERVADA", gErr, tItemInventario.sProduto, tInventario.sCodigo, tEstoqueProduto.dQuantReservada, tItemInventario.dQuantidade * dFator)
            
        Case 61391
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INV_LIBERAR_QUANTIDADE_RES_CONSIG", gErr, iIndice, tEstoqueProduto.dQuantReservadaConsig, tItemInventario.dQuantidade * dFator)

        Case 105122

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144645)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFISSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function EstoqueNegativo_Verifica(ByVal iOperacao As Integer, ByVal objItemMovEst As ClassItemMovEstoque, ByVal objTipoMovEstoque As ClassTipoMovEst) As Long
    
Dim bTestar As Boolean
Dim lErro As Long, sComandoSQL As String, tSldDiaEstAlm As typeSldDiaEstAlm
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim testprod As typeEstoqueProduto, objRegInventario As New ClassRegInventario
Dim iIndice As Integer, alComando(1 To 3) As Long
Dim dQuantRecIndl As Double

On Error GoTo Erro_EstoqueNegativo_Verifica

    bTestar = False
    
    If objTipoMovEstoque.iAtualizaSoLote = 0 Then
    
        'se for um estorno ou uma exclusao ==>
        If objTipoMovEstoque.iAtualizaMovEstoque = TIPOMOV_EST_ESTORNOMOV Or iOperacao = EXCLUSAO_MOVIMENTO Or objTipoMovEstoque.iAtualizaMovEstoque = TIPOMOV_EST_EXCLUIMOV Then
        
            If ((objTipoMovEstoque.iAtualizaConserto = TIPOMOV_EST_ADICIONACONSERTO Or _
                objTipoMovEstoque.iAtualizaConsig = TIPOMOV_EST_ADICIONACONSIGNACAO Or _
                objTipoMovEstoque.iAtualizaDemo = TIPOMOV_EST_ADICIONADEMO Or _
                objTipoMovEstoque.iAtualizaOutras = TIPOMOV_EST_ADICIONAOUTRAS Or _
                objTipoMovEstoque.iAtualizaBenef = TIPOMOV_EST_ADICIONABENEF Or _
                objTipoMovEstoque.iAtualizaNossaDisp = TIPOMOV_EST_ADICIONANOSSADISP) And objItemMovEst.dQuantidadeEst > 0) Or _
                ((objTipoMovEstoque.iAtualizaConserto = TIPOMOV_EST_SUBTRAICONSERTO Or _
                objTipoMovEstoque.iAtualizaConsig = TIPOMOV_EST_SUBTRAICONSIGNACAO Or _
                objTipoMovEstoque.iAtualizaDemo = TIPOMOV_EST_SUBTRAIDEMO Or _
                objTipoMovEstoque.iAtualizaOutras = TIPOMOV_EST_SUBTRAIOUTRAS Or _
                objTipoMovEstoque.iAtualizaBenef = TIPOMOV_EST_SUBTRAIBENEF Or _
                objTipoMovEstoque.iAtualizaNossaDisp = TIPOMOV_EST_SUBTRAINOSSADISP) And objItemMovEst.dQuantidadeEst < 0) Then bTestar = True
        Else
        
            If ((objTipoMovEstoque.iAtualizaConserto = TIPOMOV_EST_SUBTRAICONSERTO Or _
                objTipoMovEstoque.iAtualizaConsig = TIPOMOV_EST_SUBTRAICONSIGNACAO Or _
                objTipoMovEstoque.iAtualizaDemo = TIPOMOV_EST_SUBTRAIDEMO Or _
                objTipoMovEstoque.iAtualizaOutras = TIPOMOV_EST_SUBTRAIOUTRAS Or _
                objTipoMovEstoque.iAtualizaBenef = TIPOMOV_EST_SUBTRAIBENEF Or _
                objTipoMovEstoque.iAtualizaNossaDisp = TIPOMOV_EST_SUBTRAINOSSADISP) And objItemMovEst.dQuantidadeEst > 0) Or _
                ((objTipoMovEstoque.iAtualizaConserto = TIPOMOV_EST_ADICIONACONSERTO Or _
                objTipoMovEstoque.iAtualizaConsig = TIPOMOV_EST_ADICIONACONSIGNACAO Or _
                objTipoMovEstoque.iAtualizaDemo = TIPOMOV_EST_ADICIONADEMO Or _
                objTipoMovEstoque.iAtualizaOutras = TIPOMOV_EST_ADICIONAOUTRAS Or _
                objTipoMovEstoque.iAtualizaBenef = TIPOMOV_EST_ADICIONABENEF Or _
                objTipoMovEstoque.iAtualizaNossaDisp = TIPOMOV_EST_ADICIONANOSSADISP) And objItemMovEst.dQuantidadeEst < 0) Then bTestar = True
                
        End If
    End If
    
    If bTestar Then
    
        'Abre os comandos
        For iIndice = LBound(alComando) To UBound(alComando)
            alComando(iIndice) = Comando_Abrir()
            If alComando(iIndice) = 0 Then gError 130437
        Next

        'pegar estoque inicial de estoque produto e variacao por slddiaestalm
        With testprod
        
            lErro = Comando_Executar(alComando(1), "SELECT " & _
                "QuantidadeInicial , QuantDispNossa, QuantReservada, QuantReservadaConsig, QuantRecIndl, QuantIndOutras, " & _
                "QuantDefeituosa , QuantConsig3, QuantConsig, QuantDemo3, " & _
                "QuantDemo , QuantConserto3, QuantConserto, QuantOutras3, " & _
                "QuantOutras, QuantBenef, QuantBenef3, " & _
                "QuantInicialConsig3 , QuantInicialConsig, QuantInicialDemo3, QuantInicialDemo, " & _
                "QuantInicialConserto3 , QuantInicialConserto, QuantInicialOutras3, QuantInicialOutras, QuantInicialBenef, " & _
                "QuantInicialBenef3 FROM EstoqueProduto WHERE Produto = ? AND Almoxarifado = ?", _
                .dQuantidadeInicial, .dQuantDispNossa, .dQuantReservada, .dQuantReservadaConsig, .dQuantRecIndl, .dQuantInd, _
                .dQuantDefeituosa, .dQuantConsig3, .dQuantConsig, .dQuantDemo3, _
                .dQuantDemo, .dQuantConserto3, .dQuantConserto, .dQuantOutras3, _
                .dQuantOutras, .dQuantBenef, .dQuantBenef3, _
                .dQuantInicialConsig3, .dQuantInicialConsig, .dQuantInicialDemo3, .dQuantInicialDemo, _
                .dQuantInicialConserto3, .dQuantInicialConserto, .dQuantInicialOutras3, .dQuantInicialOutras, .dQuantInicialBenef, _
                .dQuantInicialBenef3, objItemMovEst.sProduto, objItemMovEst.iAlmoxarifado)
        End With
            
        If lErro <> AD_SQL_SUCESSO Then gError 130438
            
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130439
        If lErro <> AD_SQL_SUCESSO Then gError 130440
    
        'atribui estoques iniciais
        With objRegInventario
            .dQuantBenef = testprod.dQuantInicialBenef
            .dQuantBenef3 = testprod.dQuantInicialBenef3
            .dQuantConserto = testprod.dQuantInicialConserto
            .dQuantConserto3 = testprod.dQuantInicialConserto3
            .dQuantConsig = testprod.dQuantInicialConsig
            .dQuantConsig3 = testprod.dQuantInicialConsig3
            .dQuantDemo = testprod.dQuantInicialDemo
            .dQuantDemo3 = testprod.dQuantInicialDemo3
            .dQuantOutras = testprod.dQuantInicialOutras
            .dQuantOutras3 = testprod.dQuantInicialOutras3
            .dQuantidadeUMEstoque = testprod.dQuantidadeInicial + testprod.dQuantInicialBenef3 + testprod.dQuantInicialConserto3 + testprod.dQuantInicialConsig3 + testprod.dQuantInicialDemo3 + testprod.dQuantInicialOutras3
        End With
        
        '----------------------------- Leitura de SldDiaEstAlm ------------------
        sComandoSQL = "SELECT Data, QuantEntrada, QuantSaida, QuantEntRecIndl, QuantSaiRecIndl, ValorEntrada, ValorSaida, ValorEntRecIndl, ValorSaiRecIndl, " _
        & "QuantEntConsig3 , QuantSaiConsig3, QuantEntDemo3 , QuantSaiDemo3, " _
        & "QuantEntConserto3 , QuantSaiConserto3, QuantEntOutros3 , QuantSaiOutros3, " _
        & "QuantEntBenef3 , QuantSaiBenef3, QuantEntConsig , QuantSaiConsig, " _
        & "QuantEntDemo , QuantSaiDemo, QuantEntConserto , QuantSaiConserto, " _
        & "QuantEntOutros , QuantSaiOutros, QuantEntBenef , QuantSaiBenef "
    
        'Tabela, Filtro
        sComandoSQL = sComandoSQL & " FROM SldDiaEstAlm WHERE Produto = ? AND Almoxarifado = ? ORDER BY Data"
    
        With tSldDiaEstAlm
            lErro = Comando_Executar(alComando(2), sComandoSQL, .dtData, .dQuantEntrada, .dQuantSaida, .dQuantEntRecIndl, .dQuantSaiRecIndl, .dValorEntrada, .dValorSaida, .dValorEntRecIndl, .dValorSaiRecIndl, _
            .dQuantEntConsig3, .dQuantSaiConsig3, .dQuantEntDemo3, .dQuantSaiDemo3, _
            .dQuantEntConserto3, .dQuantSaiConserto3, .dQuantEntOutros3, .dQuantSaiOutros3, _
            .dQuantEntBenef3, .dQuantSaiBenef3, .dQuantEntConsig, .dQuantSaiConsig, _
            .dQuantEntDemo, .dQuantSaiDemo, .dQuantEntConserto, .dQuantSaiConserto, _
            .dQuantEntOutros, .dQuantSaiOutros, .dQuantEntBenef, .dQuantSaiBenef, _
            objItemMovEst.sProduto, objItemMovEst.iAlmoxarifado)
        End With
        If lErro <> AD_SQL_SUCESSO Then gError 130441
    
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130442
    
        Do While lErro = AD_SQL_SUCESSO
        
            With tSldDiaEstAlm
                            
                '------------------------Acumula os valores de SdlDiaEstAlm ------------------
                objRegInventario.dQuantidadeUMEstoque = objRegInventario.dQuantidadeUMEstoque + .dQuantEntrada - .dQuantSaida - (.dQuantEntRecIndl - .dQuantSaiRecIndl)
                
                dQuantRecIndl = dQuantRecIndl + (.dQuantEntRecIndl - .dQuantSaiRecIndl)
                
                objRegInventario.dQuantBenef = objRegInventario.dQuantBenef + .dQuantEntBenef - .dQuantSaiBenef
                objRegInventario.dQuantConserto = objRegInventario.dQuantConserto + .dQuantEntConserto - .dQuantSaiConserto
                objRegInventario.dQuantConsig = objRegInventario.dQuantConsig + .dQuantEntConsig - .dQuantSaiConsig
                objRegInventario.dQuantDemo = objRegInventario.dQuantDemo + .dQuantEntDemo - .dQuantSaiDemo
                objRegInventario.dQuantOutras = objRegInventario.dQuantOutras + .dQuantEntOutros - .dQuantSaiOutros
        
                objRegInventario.dQuantBenef3 = objRegInventario.dQuantBenef3 + .dQuantEntBenef3 - .dQuantSaiBenef3
                objRegInventario.dQuantConserto3 = objRegInventario.dQuantConserto3 + .dQuantEntConserto3 - .dQuantSaiConserto3
                objRegInventario.dQuantConsig3 = objRegInventario.dQuantConsig3 + .dQuantEntConsig3 - .dQuantSaiConsig3
                objRegInventario.dQuantDemo3 = objRegInventario.dQuantDemo3 + .dQuantEntDemo3 - .dQuantSaiDemo3
                objRegInventario.dQuantOutras3 = objRegInventario.dQuantOutras3 + .dQuantEntOutros3 - .dQuantSaiOutros3
        
            End With
            
            If objRegInventario.dQuantidadeUMEstoque < -QTDE_ESTOQUE_DELTA Then gError 130425
            
            If dQuantRecIndl < -QTDE_ESTOQUE_DELTA Then gError 130426
            
            If objRegInventario.dQuantBenef < -QTDE_ESTOQUE_DELTA Then gError 130427
            If objRegInventario.dQuantConserto < -QTDE_ESTOQUE_DELTA Then gError 130428
            If objRegInventario.dQuantConsig < -QTDE_ESTOQUE_DELTA Then gError 130429
            If objRegInventario.dQuantDemo < -QTDE_ESTOQUE_DELTA Then gError 130430
            If objRegInventario.dQuantOutras < -QTDE_ESTOQUE_DELTA Then gError 130431
            
            If objRegInventario.dQuantBenef3 < -QTDE_ESTOQUE_DELTA Then gError 130432
            If objRegInventario.dQuantConserto3 < -QTDE_ESTOQUE_DELTA Then gError 130433
            If objRegInventario.dQuantConsig3 < -QTDE_ESTOQUE_DELTA Then gError 130434
            If objRegInventario.dQuantDemo3 < -QTDE_ESTOQUE_DELTA Then gError 130435
            If objRegInventario.dQuantOutras3 < -QTDE_ESTOQUE_DELTA Then gError 130436
            
            lErro = Comando_BuscarProximo(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130443
    
        Loop
    
        'verificar se os saldos finais acumulados batem com os de estoqueproduto
        '??? If Abs((testprod.dQuantDispNossa + testprod.dQuantReservada) - objRegInventario.dQuantidadeUMEstoque) > QTDE_ESTOQUE_DELTA Then gError 130445
        
        If Abs(testprod.dQuantRecIndl - dQuantRecIndl) > QTDE_ESTOQUE_DELTA Then gError 130446
        
        If Abs(testprod.dQuantBenef - objRegInventario.dQuantBenef) > QTDE_ESTOQUE_DELTA Then gError 130447
        If Abs(testprod.dQuantConserto - objRegInventario.dQuantConserto) > QTDE_ESTOQUE_DELTA Then gError 130448
        If Abs((testprod.dQuantConsig + testprod.dQuantReservadaConsig) - objRegInventario.dQuantConsig) > QTDE_ESTOQUE_DELTA Then gError 130449
        If Abs(testprod.dQuantDemo - objRegInventario.dQuantDemo) > QTDE_ESTOQUE_DELTA Then gError 130450
        If Abs(testprod.dQuantOutras - objRegInventario.dQuantOutras) > QTDE_ESTOQUE_DELTA Then gError 130451
        
        If Abs(testprod.dQuantBenef3 - objRegInventario.dQuantBenef3) > QTDE_ESTOQUE_DELTA Then gError 130452
        If Abs(testprod.dQuantConserto3 - objRegInventario.dQuantConserto3) > QTDE_ESTOQUE_DELTA Then gError 130453
        If Abs(testprod.dQuantConsig3 - objRegInventario.dQuantConsig3) > QTDE_ESTOQUE_DELTA Then gError 130454
        If Abs(testprod.dQuantDemo3 - objRegInventario.dQuantDemo3) > QTDE_ESTOQUE_DELTA Then gError 130455
        If Abs(testprod.dQuantOutras3 - objRegInventario.dQuantOutras3) > QTDE_ESTOQUE_DELTA Then gError 130456
    
        'Fecha os comandos
        For iIndice = LBound(alComando) To UBound(alComando)
            lErro = Comando_Fechar(alComando(iIndice))
        Next

    End If
    
    EstoqueNegativo_Verifica = SUCESSO
     
    Exit Function
    
Erro_EstoqueNegativo_Verifica:

    EstoqueNegativo_Verifica = gErr
     
    Select Case gErr
          
        Case 130445 To 130456
            Call Rotina_Erro(vbOKOnly, "ERRO_TESTE_ESTOQUEPRODUTO", gErr, objItemMovEst.sProduto, objItemMovEst.iAlmoxarifado)
            
        Case 130437
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 130438, 130439, 130440
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ESTOQUEPRODUTO1", gErr, objItemMovEst.sProduto, objItemMovEst.iAlmoxarifado)
        
        Case 130441, 130442
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDDIAESTALM1", gErr, objItemMovEst.iAlmoxarifado, objItemMovEst.sProduto)
        
        Case 130427
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SALDO_MAT_BENEF", gErr, objItemMovEst.sProduto, objItemMovEst.iAlmoxarifado, objRegInventario.dQuantBenef)
            
        Case 130428
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SALDO_MAT_CONSERTO", gErr, objItemMovEst.sProduto, objItemMovEst.iAlmoxarifado, objRegInventario.dQuantConserto)
            
        Case 130429
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SALDO_MAT_CONSIG", gErr, objItemMovEst.sProduto, objItemMovEst.iAlmoxarifado, objRegInventario.dQuantConsig)
            
        Case 130430
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SALDO_MAT_DEMO", gErr, objItemMovEst.sProduto, objItemMovEst.iAlmoxarifado, objRegInventario.dQuantDemo)
            
        Case 130431
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SALDO_MAT_OUTRAS", gErr, objItemMovEst.sProduto, objItemMovEst.iAlmoxarifado, objRegInventario.dQuantOutras)
            
        Case 130432
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SALDO_MAT_BENEF3", gErr, objEstoqueProduto.sProduto, objEstoqueProduto.iAlmoxarifado, objRegInventario.dQuantBenef3)
        
        Case 130433
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SALDO_MAT_CONSERTO3", gErr, objItemMovEst.sProduto, objItemMovEst.iAlmoxarifado, objRegInventario.dQuantConserto3)
            
        Case 130434
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SALDO_MAT_CONSIG3", gErr, objItemMovEst.sProduto, objItemMovEst.iAlmoxarifado, objRegInventario.dQuantConsig3)
            
        Case 130435
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SALDO_MAT_DEMO3", gErr, objItemMovEst.sProduto, objItemMovEst.iAlmoxarifado, objRegInventario.dQuantDemo3)
            
        Case 130436
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SALDO_MAT_OUTRAS3", gErr, objItemMovEst.sProduto, objItemMovEst.iAlmoxarifado, objRegInventario.dQuantOutras3)
            
        Case 130425
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SALDO_MAT_DISPONIVEL", gErr, objItemMovEst.sProduto, objItemMovEst.iAlmoxarifado, objRegInventario.dQuantidadeUMEstoque)
            
        Case 130426
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SALDO_MAT_INDISPONIVEL", gErr, objItemMovEst.sProduto, objItemMovEst.iAlmoxarifado, dQuantRecIndl)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    Exit Function
    
End Function

Public Function ApuracaoRes_ObtemGuias(ByVal dtDataDe As Date, ByVal dtDataAte As Date, ByVal iFilialEmpresa As Integer, ByVal colGuias As Collection) As Long
'preenche colGuias com as guias de ICMS ordenadas por data e numero
'ATENCAO: apenas para apoio a adrelvb

Dim lErro As Long, lComando As Long, objGuia As ClassGuiasICMS, tGuia As typeGuiasICMS
Dim dtDataDePer As Date, dtDataAtePer As Date

On Error GoTo Erro_ApuracaoRes_ObtemGuias

    dtDataDePer = DateAdd("m", 1, dtDataDe)
    dtDataAtePer = DateAdd("m", 1, dtDataAte)
    
    lComando = Comando_AbrirExt(GL_lConexaoRel)
    If lComando = 0 Then gError 124106
    
    With tGuia
    
        .sLocalEntrega = String(STRING_GUIASICMS_LOCALENTREGA, 0)
        .sNumero = String(STRING_GUIASICMS_NUMERO, 0)
        .sOrgaoArrecadador = String(STRING_GUIASICMS_ORGAOARRECADADOR, 0)
        
    End With
    
    lErro = Comando_Executar(lComando, "SELECT Numero, Data, Valor, OrgaoArrecadador, DataEntrega, LocalEntrega FROM GuiasICMS WHERE FilialEmpresa = ? AND Data >= ? AND Data <= ? ORDER BY Data, Numero", _
        tGuia.sNumero, tGuia.dtData, tGuia.dValor, tGuia.sOrgaoArrecadador, tGuia.dtDataEntrega, tGuia.sLocalEntrega, iFilialEmpresa, dtDataDePer, dtDataAtePer)
    If lErro <> AD_SQL_SUCESSO Then gError 124107
        
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 124108
    
    Do While lErro = AD_SQL_SUCESSO
    
        Set objGuia = New ClassGuiasICMS
        
        With objGuia
        
            .dtData = tGuia.dtData
            .dtDataEntrega = tGuia.dtDataEntrega
            .dValor = tGuia.dValor
            .iFilialEmpresa = iFilialEmpresa
            .sLocalEntrega = tGuia.sLocalEntrega
            .sNumero = tGuia.sNumero
            .sOrgaoArrecadador = tGuia.sOrgaoArrecadador
        
        End With
        
        colGuias.Add objGuia
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 124109
    
    Loop
    
    Call Comando_Fechar(lComando)
    
    ApuracaoRes_ObtemGuias = SUCESSO
     
    Exit Function
    
Erro_ApuracaoRes_ObtemGuias:

    ApuracaoRes_ObtemGuias = gErr
     
    Select Case gErr
          
        Case 124106
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 124107, 124108, 124109
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_GUIASICMSREL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150470)
     
    End Select
     
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function


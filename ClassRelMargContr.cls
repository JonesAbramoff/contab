VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelMargContr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'variaveis globais
Private gobjMargContr As ClassMargContr
Private gobjTelaComissoes As Object
Private gcolComissoes As Collection
Private gdPrecoComissoes As Double
Private dtDataCF As Date
Private sDVVTotal(1 To 3) As String
Private dTaxaDescPadrao As Double, dTaxaValFut As Double, dDiasValFut As Double

'Identificação
Private iUltFilialEmpresa As Integer
Private sUltProduto As String
Private lUltCliente As Long
Private iUltFilialCli As Integer
Private dUltQtde As Double

Private sUltUM As String
Private sUltNomeRedCli As String

'Valores armazenados para a última pesquisa
Private dUltPrecoCli As Double
Private dUltPrecoPraticado As Double
Private dUltPercMargPrat As Double
Private dUltPercLucroPrat As Double

Private alComando(1 To 1) As Long

Function MargContr_ObtemPrecoCli(dValor As Double, ByVal iFilialFaturamento As Integer, ByVal lCliente As Long, ByVal iFilialCli As Integer, ByVal sProduto As String, ByVal dQtde As Double, ByVal sUM As String, ByVal iTabelaPreco As Integer, ByVal iColunaCli As Integer, ByVal iColunaPraticado As Integer, ByVal iLinhaPercMarg As Integer, ByVal iLinhaPercLucro As Integer) As Long

Dim lErro As Long

On Error GoTo Erro_MargContr_ObtemPrecoCli

    If UCase(gsNomePrinc) = "SGEPRINC" Then

        lErro = MargContr_Verifica(iFilialFaturamento, lCliente, iFilialCli, sProduto, dQtde, sUM, iTabelaPreco, iColunaCli, iColunaPraticado, iLinhaPercMarg, iLinhaPercLucro)
        If lErro <> SUCESSO Then gError 130048
        
        dValor = dUltPrecoCli
    
    Else
    
        gError 216152
    
        dValor = 0
    
    End If
    
    MargContr_ObtemPrecoCli = SUCESSO
     
    Exit Function
    
Erro_MargContr_ObtemPrecoCli:

    MargContr_ObtemPrecoCli = gErr
     
    Select Case gErr
          
        Case 130048
        
        Case 216152
            Call Rotina_Erro(vbOKOnly, "ERRO_ADRELVB_NAO_PODE_SER_CALC_FORA_SGEPRINC", gErr, "MargContr", gsNomePrinc)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153229)
     
    End Select
     
    Exit Function

End Function

Function MargContr_ObtemPrecoPrat(dValor As Double, ByVal iFilialFaturamento As Integer, ByVal lCliente As Long, ByVal iFilialCli As Integer, ByVal sProduto As String, ByVal dQtde As Double, ByVal sUM As String, ByVal iTabelaPreco As Integer, ByVal iColunaCli As Integer, ByVal iColunaPraticado As Integer, ByVal iLinhaPercMarg As Integer, ByVal iLinhaPercLucro As Integer) As Long

Dim lErro As Long

On Error GoTo Erro_MargContr_ObtemPrecoPrat

    lErro = MargContr_Verifica(iFilialFaturamento, lCliente, iFilialCli, sProduto, dQtde, sUM, iTabelaPreco, iColunaCli, iColunaPraticado, iLinhaPercMarg, iLinhaPercLucro)
    If lErro <> SUCESSO Then gError 130048
    
    dValor = dUltPrecoPraticado
    
    MargContr_ObtemPrecoPrat = SUCESSO
     
    Exit Function
    
Erro_MargContr_ObtemPrecoPrat:

    MargContr_ObtemPrecoPrat = gErr
     
    Select Case gErr
          
        Case 130048
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153230)
     
    End Select
     
    Exit Function

End Function

Function MargContr_ObtemMargPrat(dValor As Double, ByVal iFilialFaturamento As Integer, ByVal lCliente As Long, ByVal iFilialCli As Integer, ByVal sProduto As String, ByVal dQtde As Double, ByVal sUM As String, ByVal iTabelaPreco As Integer, ByVal iColunaCli As Integer, ByVal iColunaPraticado As Integer, ByVal iLinhaPercMarg As Integer, ByVal iLinhaPercLucro As Integer) As Long

Dim lErro As Long

On Error GoTo Erro_MargContr_ObtemMargPrat

    lErro = MargContr_Verifica(iFilialFaturamento, lCliente, iFilialCli, sProduto, dQtde, sUM, iTabelaPreco, iColunaCli, iColunaPraticado, iLinhaPercMarg, iLinhaPercLucro)
    If lErro <> SUCESSO Then gError 130048
    
    dValor = dUltPercMargPrat / 100
    
    MargContr_ObtemMargPrat = SUCESSO
     
    Exit Function
    
Erro_MargContr_ObtemMargPrat:

    MargContr_ObtemMargPrat = gErr
     
    Select Case gErr
          
        Case 130048
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153231)
     
    End Select
     
    Exit Function

End Function

Function MargContr_ObtemLucroPrat(dValor As Double, ByVal iFilialFaturamento As Integer, ByVal lCliente As Long, ByVal iFilialCli As Integer, ByVal sProduto As String, ByVal dQtde As Double, ByVal sUM As String, ByVal iTabelaPreco As Integer, ByVal iColunaCli As Integer, ByVal iColunaPraticado As Integer, ByVal iLinhaPercMarg As Integer, ByVal iLinhaPercLucro As Integer) As Long

Dim lErro As Long

On Error GoTo Erro_MargContr_ObtemLucroPrat

    lErro = MargContr_Verifica(iFilialFaturamento, lCliente, iFilialCli, sProduto, dQtde, sUM, iTabelaPreco, iColunaCli, iColunaPraticado, iLinhaPercMarg, iLinhaPercLucro)
    If lErro <> SUCESSO Then gError 130048
    
    dValor = dUltPercLucroPrat / 100
    
    MargContr_ObtemLucroPrat = SUCESSO
     
    Exit Function
    
Erro_MargContr_ObtemLucroPrat:

    MargContr_ObtemLucroPrat = gErr
     
    Select Case gErr
          
        Case 130048
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153232)
     
    End Select
     
    Exit Function

End Function

Private Function MargContr_Verifica(ByVal iFilialFaturamento As Integer, ByVal lCliente As Long, ByVal iFilialCli As Integer, ByVal sProduto As String, ByVal dQtde As Double, ByVal sUM As String, ByVal iTabelaPreco As Integer, ByVal iColunaCli As Integer, ByVal iColunaPraticado As Integer, ByVal iLinhaPercMarg As Integer, ByVal iLinhaPercLucro As Integer) As Long

Dim lErro As Long, objCodigoNome As New AdmlCodigoNome

On Error GoTo Erro_MargContr_Verifica

    If gobjMargContr Is Nothing Then
    
        Set gobjMargContr = New ClassMargContr
    
        lErro = CF("MargContr_Le_DVV", gobjMargContr)
        If lErro <> SUCESSO Then gError 130041
        
        lErro = CF("MargContr_Le_Analise", gobjMargContr)
        If lErro <> SUCESSO Then gError 130042
        
        lErro = Chama_Tela_Nova_Instancia1("ComissoesCalcula", gobjTelaComissoes)
        If lErro <> SUCESSO Then gError 130049
    
        lErro = CF("FilialEmpresa_ObtemDataCustoFixo", iFilialFaturamento, dtDataCF)
        If lErro <> SUCESSO Then gError 130050
    
        lErro = CF("CalcMP_ObterValores", giFilialEmpresa, dTaxaDescPadrao, dTaxaValFut, dDiasValFut)
        If lErro <> SUCESSO Then gError 130051

    End If
    
    'se trocou algum parametro entao fazer o calculo novamente
    If iUltFilialEmpresa <> iFilialFaturamento Or sUltProduto <> sProduto Or lUltCliente <> lCliente Or iUltFilialCli <> iFilialCli Or dUltQtde <> dQtde Or sUltUM <> sUM Then

        With gobjMargContr
        
            .dQuantidade = dQtde
            .iCodFilial = iFilialCli
            .iFilialFaturamento = iFilialFaturamento
            .lCodCliente = lCliente
            .sProduto = sProduto
            .iTabelaPreco = iTabelaPreco
            
        End With
        
        If lCliente <> lUltCliente Then
        
            objCodigoNome.lCodigo = lCliente
            lErro = CF("Cliente_Le_NomeRed", objCodigoNome)
            If lErro <> SUCESSO And lErro <> 12553 Then gError 130043
            If lErro <> SUCESSO Then gError 130044
            
            sUltNomeRedCli = objCodigoNome.sNome
            
        End If
        
        sUltUM = sUM
        
        lErro = MargContr_TestaDVV
        If lErro <> SUCESSO Then
        
            dUltPrecoCli = 0
            dUltPrecoPraticado = 0
            dUltPercMargPrat = 0
            dUltPercLucroPrat = 0
        
        Else
        
            lErro = CF("MargContr_CalculaComissoes", gcolComissoes, gobjMargContr, gdPrecoComissoes, gobjTelaComissoes, sUltNomeRedCli, sUM)
            If lErro <> SUCESSO Then gError 130045
        
            'calcula a DVV para o cliente
            lErro = MargContr_CalculaDVV(iColunaCli)
            If lErro <> SUCESSO Then gError 130047
        
            'calcula a DVV para o preco praticado
            lErro = MargContr_CalculaDVV(iColunaPraticado)
            If lErro <> SUCESSO Then gError 130052
        
            lErro = Analise_RecalcularColuna(iColunaCli, iColunaCli, iColunaPraticado, iLinhaPercMarg, iLinhaPercLucro)
            If lErro <> SUCESSO Then gError 130053
            
            lErro = Analise_RecalcularColuna(iColunaPraticado, iColunaCli, iColunaPraticado, iLinhaPercMarg, iLinhaPercLucro)
            If lErro <> SUCESSO Then gError 130054
        
        End If
        
        iUltFilialEmpresa = iFilialFaturamento
        sUltProduto = sProduto
        lUltCliente = lCliente
        iUltFilialCli = iFilialCli
        dUltQtde = dQtde
            
    End If
    
    MargContr_Verifica = SUCESSO
     
    Exit Function
    
Erro_MargContr_Verifica:

    MargContr_Verifica = gErr
     
    Select Case gErr
          
        Case 130041, 130042, 130043, 130045, 130047, 130049 To 130054
        
        Case 130044
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_ENCONTRADO", gErr, lCliente)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153233)
     
    End Select
     
    Exit Function

End Function

Private Sub Class_Terminate()
    
Dim iIndice As Integer
    
    Set gobjMargContr = Nothing
   
    If Not gobjTelaComissoes Is Nothing Then
    
        gobjTelaComissoes.Unload gobjTelaComissoes
        
    End If
    
    Set gcolComissoes = Nothing
    
    For iIndice = LBound(alComando) To UBound(alComando)
        If alComando(iIndice) <> 0 Then Call Comando_Fechar(alComando(iIndice))
    Next

End Sub

Private Function MargContr_CalculaDVV(ByVal iColuna As Integer) As Long

Dim lErro As Long, iLinha As Integer, objDVVLinCol As ClassDVVLinCol
Dim colPlanilhas As New Collection, objPlanilhas As ClassPlanilhas, sFormula As String
Dim objContexto As New ClassContextoPlan, dValor As Double

On Error GoTo Erro_MargContr_CalculaDVV

    If gobjMargContr.colDVVLin.Count <> 0 Then
    
        For iLinha = 1 To gobjMargContr.colDVVLin.Count
                
            Set objDVVLinCol = gobjMargContr.colDVVLinCol(gobjMargContr.IndDVV(iLinha, iColuna - 1))
            
            If Len(Trim(objDVVLinCol.sFormula)) = 0 Then
            
                sFormula = "0" '??? modificado em relacao a tela
                
            Else
            
                sFormula = objDVVLinCol.sFormula
                
            End If
            
            Set objPlanilhas = New ClassPlanilhas
            
            With objPlanilhas
                .iTipoPlanilha = PLANILHA_TIPO_DVV
                .iFilialEmpresa = gobjMargContr.iFilialFaturamento
                .iEscopo = MNEMONICOFPRECO_ESCOPO_PRODUTO
                .iLinha = iLinha
                .sExpressao = sFormula
            End With
            
            colPlanilhas.Add objPlanilhas
        
        Next
        
        With objContexto
            .iFilialFaturamento = gobjMargContr.iFilialFaturamento
            .sProduto = gobjMargContr.sProduto
            .dQuantidade = gobjMargContr.dQuantidade
            .iFilialCli = gobjMargContr.iCodFilial
            .lCliente = gobjMargContr.lCodCliente
            .iVendedor = gobjMargContr.iCodVendedor
            .iTabelaPreco = gobjMargContr.iTabelaPreco
            .iAno = Year(gdtDataAtual)
            .dtDataCustoFixo = dtDataCF
            .iRotinaOrigem = FORMACAO_PRECO_ANALISE_MARGCONTR_REL
            Set .colComissoes = gcolComissoes
            .dPrecoPraticado = gdPrecoComissoes
            .sNomeRedCliente = sUltNomeRedCli
            .sUM = sUltUM
        End With
        
        'Executa as formulas da planilha de preço. Retorna o valor da planilha em dValor (que é o valor da última linha da planilha) e o valor de cada linha em colPlanilhas.Item(?).dValor
        lErro = CF("Avalia_Expressao_FPreco3", colPlanilhas, dValor, objContexto)
        If lErro <> SUCESSO Then gError 130046
    
        Set objPlanilhas = colPlanilhas.Item(colPlanilhas.Count)
        
        sDVVTotal(iColuna - 1) = Format(objPlanilhas.dValor * 100, "###,##0.00###")
    
    End If
    
    MargContr_CalculaDVV = SUCESSO
     
    Exit Function
    
Erro_MargContr_CalculaDVV:

    MargContr_CalculaDVV = gErr
     
    Select Case gErr
          
        Case 130046
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153234)
     
    End Select
     
    Exit Function

End Function

Private Function Analise_RecalcularColuna(ByVal iColuna As Integer, ByVal iColunaCli As Integer, ByVal iColunaPraticado As Integer, ByVal iLinhaPercMarg As Integer, ByVal iLinhaPercLucro As Integer) As Long

Dim lErro As Long, iLinha As Integer, objLinCol As ClassPlanMargContrLinCol, objLin As ClassPlanMargContrLin
Dim colPlanilhas As New Collection, objPlanilhas As ClassPlanilhas, sFormula As String
Dim objContexto As New ClassContextoPlan, dValor As Double, colPlanilhas2 As New Collection
Dim objMnemonicoValor As ClassMnemonicoValor, objContexto2 As New ClassContextoPlan, iCol As Integer
Dim dValorL1 As Double

On Error GoTo Erro_Analise_RecalcularColuna

    'calcular o preco para o cliente se a coluna for iColunaCli
    If iColuna = iColunaCli Then
    
        With objContexto2
            .iFilialFaturamento = gobjMargContr.iFilialFaturamento
            .sProduto = gobjMargContr.sProduto
            .dQuantidade = gobjMargContr.dQuantidade
            .iFilialCli = gobjMargContr.iCodFilial
            .lCliente = gobjMargContr.lCodCliente
            .iVendedor = gobjMargContr.iCodVendedor
            .iTabelaPreco = gobjMargContr.iTabelaPreco
            .iAno = Year(gdtDataAtual)
            .dtDataCustoFixo = dtDataCF
            .iRotinaOrigem = FORMACAO_PRECO_ANALISE_MARGCONTR_REL
            Set .colComissoes = gcolComissoes
            .dPrecoPraticado = gdPrecoComissoes
            .dTaxaDescPadrao = dTaxaDescPadrao
            .dTaxaValFut = dTaxaValFut
            .dDiasValFut = dDiasValFut
        End With
        
        'guarda os resultados do grid dvv para todas as colunas
        Call DVVTotal_GuardaValores(objContexto2)
        
        For iLinha = 1 To gobjMargContr.colPlanMargContrLin.Count
        
            Set objLin = gobjMargContr.colPlanMargContrLin(iLinha)
            Set objLinCol = gobjMargContr.colPlanMargContrLinCol(gobjMargContr.IndAnalise(iLinha, iColuna - 1))
            
            sFormula = objLinCol.sFormula
            If iLinha = 1 And Len(Trim(sFormula)) <> 0 Then Exit For
            If Len(Trim(objLin.sFormulaL1)) = 0 And Len(Trim(sFormula)) = 0 Then sFormula = objLin.sFormulaGeral
            
            If Len(Trim(sFormula)) = 0 Then sFormula = "0"
            
            Set objPlanilhas = New ClassPlanilhas
            
            With objPlanilhas
                .iTipoPlanilha = PLANILHA_TIPO_TODOS
                .iFilialEmpresa = gobjMargContr.iFilialFaturamento
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
                    .iFilialEmpresa = gobjMargContr.iFilialFaturamento
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
     
        dUltPrecoCli = dValorL1
        
    Else
    
        'guarda os resultados do grid dvv para todas as colunas
        Call DVVTotal_GuardaValores(objContexto)
        
        For iLinha = 1 To gobjMargContr.colPlanMargContrLin.Count
        
            Set objLin = gobjMargContr.colPlanMargContrLin(iLinha)
            Set objLinCol = gobjMargContr.colPlanMargContrLinCol(gobjMargContr.IndAnalise(iLinha, iColuna - 1))
                
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
                .iFilialEmpresa = gobjMargContr.iFilialFaturamento
                .iEscopo = MNEMONICOFPRECO_ESCOPO_PRODUTO
                .iLinha = iLinha
                .sExpressao = sFormula
            End With
            
            colPlanilhas.Add objPlanilhas
        
        Next
        
        With objContexto
            .iFilialFaturamento = gobjMargContr.iFilialFaturamento
            .sProduto = gobjMargContr.sProduto
            .dQuantidade = gobjMargContr.dQuantidade
            .iFilialCli = gobjMargContr.iCodFilial
            .lCliente = gobjMargContr.lCodCliente
            .iVendedor = gobjMargContr.iCodVendedor
            .iTabelaPreco = gobjMargContr.iTabelaPreco
            .iAno = Year(gdtDataAtual)
            .dtDataCustoFixo = dtDataCF
            .iRotinaOrigem = FORMACAO_PRECO_ANALISE_MARGCONTR_REL
            Set .colComissoes = gcolComissoes
            .dPrecoPraticado = gdPrecoComissoes
            .dTaxaDescPadrao = dTaxaDescPadrao
            .dTaxaValFut = dTaxaValFut
            .dDiasValFut = dDiasValFut
        End With
        
        'Executa as formulas da planilha de preço. Retorna o valor da planilha em dValor (que é o valor da última linha da planilha) e o valor de cada linha em colPlanilhas.Item(?).dValor
        lErro = CF("Avalia_Expressao_FPreco3", colPlanilhas, dValor, objContexto)
        If lErro <> SUCESSO Then gError 106721
    
        For Each objPlanilhas In colPlanilhas
        
            If objPlanilhas.iLinha = 1 Then
                dUltPrecoPraticado = StrParaDbl(Format(objPlanilhas.dValor, "###,##0.00###"))
            End If
            
            If objPlanilhas.iLinha = iLinhaPercMarg Then
                dUltPercMargPrat = StrParaDbl(Format(objPlanilhas.dValor * 100, "###,##0.00###"))
            End If
    
            If objPlanilhas.iLinha = iLinhaPercLucro Then
                dUltPercLucroPrat = StrParaDbl(Format(objPlanilhas.dValor * 100, "###,##0.00###"))
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153235)
     
    End Select
     
    Exit Function

End Function

Private Sub DVVTotal_GuardaValores(objContexto As ClassContextoPlan)

Dim iCol As Integer, objMnemonicoValor As ClassMnemonicoValor

    'guarda os resultados do grid dvv para todas as colunas
    For iCol = 1 To 3
    
        Set objMnemonicoValor = New ClassMnemonicoValor
        Set objMnemonicoValor.colValor = New Collection
        
        objMnemonicoValor.sMnemonico = "DVVTotal"
        objMnemonicoValor.vParam(1) = CDbl(iCol)
    
        objMnemonicoValor.colValor.Add StrParaDbl(StrParaDbl(sDVVTotal(iCol)) / 100)
        
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

Private Function MargContr_TestaDVV() As Long

Dim lErro As Long, dPercDVV As Double, lComando As Long

On Error GoTo Erro_MargContr_TestaDVV

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 106666
    
    lErro = Comando_Executar(lComando, "SELECT PercDVV FROM DVVClienteProd WHERE FilialEmpresa = ? AND CodCliente = ? AND CodFilial = ? AND Produto = ?", dPercDVV, gobjMargContr.iFilialFaturamento, gobjMargContr.lCodCliente, gobjMargContr.iCodFilial, gobjMargContr.sProduto)
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153236)
     
    End Select
     
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function ProdutoCli_Obtem_PrecoDesde(ByVal iFilialEmpresa As Integer, ByVal lCliente As Long, ByVal iFilial As Integer, ByVal sProduto As String, ByVal dtData As Date, ByVal dValor As Double, sDesde As String) As Long

Dim lErro As Long, iIndice As Integer, dtDataEmissaoNF As Date
Dim dPrecoUnitario As Double, dPrecoDesde As Date, dPrecoUnitInicial As Double

On Error GoTo Erro_ProdutoCli_Obtem_PrecoDesde

    If alComando(UBound(alComando)) = 0 Then
        For iIndice = LBound(alComando) To UBound(alComando)
            alComando(iIndice) = Comando_AbrirExt(GL_lConexaoRel)
            If alComando(iIndice) = 0 Then gError 130784
        Next
    End If
    
    dPrecoDesde = dtData
    dPrecoUnitInicial = dValor
    
    lErro = Comando_Executar(alComando(1), "SELECT DataEmissao, PrecoUnitario FROM NFiscal, ItensNFiscal, TiposDocInfo WHERE NFiscal.TipoNFiscal = Codigo AND NFiscal.NumIntDoc = NumIntNF AND NFiscal.Status <> 7 AND Faturamento=1 AND FilialEmpresa = ? AND Cliente = ? AND FilialCli = ? AND Produto = ? AND DataEmissao < ? ORDER BY DataEmissao DESC", _
        dtDataEmissaoNF, dPrecoUnitario, iFilialEmpresa, lCliente, iFilial, sProduto, dtData)
    If lErro <> AD_SQL_SUCESSO Then gError 130787
    
    lErro = Comando_BuscarProximo(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130788

    Do While lErro = AD_SQL_SUCESSO
            
        If Abs(dPrecoUnitario - dPrecoUnitInicial) > DELTA_VALORMONETARIO Then
            Exit Do
        Else
            dPrecoDesde = dtDataEmissaoNF
        End If
        
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130789
    
    Loop
    
    sDesde = Format(dPrecoDesde, "dd/mm/yy")
    
    ProdutoCli_Obtem_PrecoDesde = SUCESSO
     
    Exit Function
    
Erro_ProdutoCli_Obtem_PrecoDesde:

    ProdutoCli_Obtem_PrecoDesde = gErr
     
    Select Case gErr
          
        Case 130784
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 130785 To 130790
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ULTVENDASPV", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    Exit Function

End Function

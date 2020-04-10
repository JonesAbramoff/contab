VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassMargContrAuxCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'variaveis globais
Private gobjMargContr As ClassMargContr

Private bComPrecoPraticado As Boolean

Private lUltCliente As Long
Private sUltNomeRedCli As String

Private sDVVTotal(1 To 3) As String

Public Function MargContr_InicioCalc(ByVal iFilialFaturamento As Integer, objMargContr As ClassMargContr, ByVal bComPrecoPrat As Boolean) As Long

Dim objTelaComissoes As Object
Dim dtDataCF As Date, lErro As Long
Dim dTaxaDescPadrao As Double, dTaxaValFut As Double, dDiasValFut As Double
        
On Error GoTo Erro_MargContr_InicioCalc

    Set objMargContr = New ClassMargContr
    Set gobjMargContr = objMargContr

    gobjMargContr.iFilialFaturamento = iFilialFaturamento
    
    bComPrecoPraticado = bComPrecoPrat
    
    lErro = CF("MargContr_Le_DVV", gobjMargContr)
    If lErro <> SUCESSO Then gError 130041
    
    lErro = CF("MargContr_Le_Analise", gobjMargContr)
    If lErro <> SUCESSO Then gError 130042
    
    lErro = Chama_Tela_Nova_Instancia1("ComissoesCalcula", objTelaComissoes)
    If lErro <> SUCESSO Then gError 130049

    Set gobjMargContr.objTelaComissoes = objTelaComissoes
    
    lErro = CF("FilialEmpresa_ObtemDataCustoFixo", iFilialFaturamento, dtDataCF)
    If lErro <> SUCESSO Then gError 130050

    gobjMargContr.dtDataCF = dtDataCF
    
    lErro = CF("CalcMP_ObterValores", iFilialFaturamento, dTaxaDescPadrao, dTaxaValFut, dDiasValFut)
    If lErro <> SUCESSO Then gError 130051
    
    With gobjMargContr
    
        .dTaxaDescPadrao = dTaxaDescPadrao
        .dTaxaValFut = dTaxaValFut
        .dDiasValFut = dDiasValFut
        .iColunaCli = 3
        .iColunaPraticado = 4
        .iColunaSimulacao = 9
        .iLinhaPercMarg = 17
        .iLinhaPrazo = 2
        .iLinhaPercLucro = 20
        .iLinhaValorMarg = 21
        
    End With
        
    MargContr_InicioCalc = SUCESSO
    
    Exit Function
    
Erro_MargContr_InicioCalc:

    MargContr_InicioCalc = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Public Function MargContr_FimCalc() As Long

On Error GoTo Erro_MargContr_FimCalc

    If Not gobjMargContr.objTelaComissoes Is Nothing Then
    
        gobjMargContr.objTelaComissoes.Unload gobjMargContr.objTelaComissoes
        Set gobjMargContr.objTelaComissoes = Nothing
        
    End If
    
    Set gobjMargContr.colComissoes = Nothing
    Set gobjMargContr = Nothing
   
    MargContr_FimCalc = SUCESSO
    
    Exit Function
    
Erro_MargContr_FimCalc:

    MargContr_FimCalc = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Public Function MargContr_DetCalc(ByVal lCliente As Long, ByVal iFilialCli As Integer, ByVal sProduto As String, ByVal dQtde As Double, ByVal sUM As String, ByVal iTabelaPreco As Integer) As Long

Dim lErro As Long, objCodigoNome As New AdmlCodigoNome, colComissoesAux As Collection, dPrecoComissoesAux As Double
Dim objTelaComissoesAux As Object

On Error GoTo Erro_MargContr_DetCalc

    With gobjMargContr
    
        .dQuantidade = dQtde
        .iCodFilial = iFilialCli
'            .iFilialFaturamento = iFilialFaturamento
        .lCodCliente = lCliente
        .sProduto = sProduto
        .iTabelaPreco = iTabelaPreco
        .sUnidadeMed = sUM
        
    End With
    
    If lCliente <> lUltCliente Then
    
        objCodigoNome.lCodigo = lCliente
        lErro = CF("Cliente_Le_NomeRed", objCodigoNome)
        If lErro <> SUCESSO And lErro <> 12553 Then gError 130043
        If lErro <> SUCESSO Then gError 130044
        
        sUltNomeRedCli = objCodigoNome.sNome
        lUltCliente = lCliente
        
    End If
    
    lErro = MargContr_TestaDVV
    If lErro <> SUCESSO Then
            
        With gobjMargContr
            .dUltPrecoCli = 0
            .dUltPrecoPraticado = 0
            .dUltPercMargPrat = 0
            .dUltPercLucroPrat = 0
        End With
        
    Else
    
        Set objTelaComissoesAux = gobjMargContr.objTelaComissoes
        
        lErro = CF("MargContr_CalculaComissoes", colComissoesAux, gobjMargContr, dPrecoComissoesAux, objTelaComissoesAux, sUltNomeRedCli, gobjMargContr.sUnidadeMed)
        If lErro <> SUCESSO Then gError 130045
    
        Set gobjMargContr.colComissoes = colComissoesAux
        gobjMargContr.dPrecoComissoes = dPrecoComissoesAux
    
        'calcula a DVV para o cliente
        lErro = MargContr_CalculaDVV(gobjMargContr.iColunaCli)
        If lErro <> SUCESSO Then gError 130047
    
        lErro = Analise_RecalcularColuna(gobjMargContr.iColunaCli, gobjMargContr.iColunaCli, gobjMargContr.iColunaPraticado, gobjMargContr.iLinhaPercMarg, gobjMargContr.iLinhaPercLucro)
        If lErro <> SUCESSO Then gError 130053
        
        If bComPrecoPraticado Then
        
            'calcula a DVV para o preco praticado
            lErro = MargContr_CalculaDVV(gobjMargContr.iColunaPraticado)
            If lErro <> SUCESSO Then gError 130052
    
            lErro = Analise_RecalcularColuna(gobjMargContr.iColunaPraticado, gobjMargContr.iColunaCli, gobjMargContr.iColunaPraticado, gobjMargContr.iLinhaPercMarg, gobjMargContr.iLinhaPercLucro)
            If lErro <> SUCESSO Then gError 130054
    
        End If
        
    End If
        
    MargContr_DetCalc = SUCESSO
     
    Exit Function
    
Erro_MargContr_DetCalc:

    MargContr_DetCalc = gErr
     
    Select Case gErr
          
        Case 130041, 130042, 130043, 130045, 130047, 130049 To 130054
        
        Case 130044
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_ENCONTRADO", gErr, lCliente)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    Exit Function

End Function

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
            .dtDataCustoFixo = gobjMargContr.dtDataCF
            .iRotinaOrigem = FORMACAO_PRECO_ANALISE_MARGCONTR_REL
            Set .colComissoes = gobjMargContr.colComissoes
            .dPrecoPraticado = gobjMargContr.dPrecoComissoes
            .sNomeRedCliente = sUltNomeRedCli
            .sUM = gobjMargContr.sUnidadeMed
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
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
            .dtDataCustoFixo = gobjMargContr.dtDataCF
            .iRotinaOrigem = FORMACAO_PRECO_ANALISE_MARGCONTR_REL
            Set .colComissoes = gobjMargContr.colComissoes
            .dPrecoPraticado = gobjMargContr.dPrecoComissoes
            .dTaxaDescPadrao = gobjMargContr.dTaxaDescPadrao
            .dTaxaValFut = gobjMargContr.dTaxaValFut
            .dDiasValFut = gobjMargContr.dDiasValFut
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
     
        gobjMargContr.dUltPrecoCli = dValorL1
        
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
            .dtDataCustoFixo = gobjMargContr.dtDataCF
            .iRotinaOrigem = FORMACAO_PRECO_ANALISE_MARGCONTR_REL
            Set .colComissoes = gobjMargContr.colComissoes
            .dPrecoPraticado = gobjMargContr.dPrecoComissoes
            .dTaxaDescPadrao = gobjMargContr.dTaxaDescPadrao
            .dTaxaValFut = gobjMargContr.dTaxaValFut
            .dDiasValFut = gobjMargContr.dDiasValFut
        End With
        
        'Executa as formulas da planilha de preço. Retorna o valor da planilha em dValor (que é o valor da última linha da planilha) e o valor de cada linha em colPlanilhas.Item(?).dValor
        lErro = CF("Avalia_Expressao_FPreco3", colPlanilhas, dValor, objContexto)
        If lErro <> SUCESSO Then gError 106721
    
        For Each objPlanilhas In colPlanilhas
        
            If objPlanilhas.iLinha = 1 Then
                gobjMargContr.dUltPrecoPraticado = StrParaDbl(Format(objPlanilhas.dValor, "###,##0.00###"))
            End If
            
            If objPlanilhas.iLinha = iLinhaPercMarg Then
                gobjMargContr.dUltPercMargPrat = StrParaDbl(Format(objPlanilhas.dValor * 100, "###,##0.00###"))
            End If
    
            If objPlanilhas.iLinha = iLinhaPercLucro Then
                gobjMargContr.dUltPercLucroPrat = StrParaDbl(Format(objPlanilhas.dValor * 100, "###,##0.00###"))
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function MargContr_DetCalcMarg(dPercMarg As Double, ByVal dQuantidade As Double, ByVal lCliente As Long, ByVal iFilialCli As Integer, ByVal iFilialEmpresaFaturamento As Integer, ByVal sProduto As String, ByVal iTabelaPreco As Integer, ByVal sUnidadeMed As String, ByVal iCondicaoPagto As Integer, ByVal sNaturezaOp As String, ByVal dPrecoVenda As Double, ByVal dFrete As Double) As Long

Dim lErro As Long, objCodigoNome As New AdmlCodigoNome, colComissoesAux As Collection, dPrecoComissoesAux As Double
Dim objTelaComissoesAux As Object, dPrazo As Double, objCondicaoPagto As New ClassCondicaoPagto, dFator As Double, objProduto As New ClassProduto

On Error GoTo Erro_MargContr_DetCalcMarg

    With gobjMargContr
    
        .dQuantidade = dQuantidade
        .iCodFilial = iFilialCli
        .iFilialFaturamento = iFilialEmpresaFaturamento
        .lCodCliente = lCliente
        .sProduto = sProduto
        .iTabelaPreco = iTabelaPreco
        .sUnidadeMed = sUnidadeMed
        
    End With
    
    If lUltCliente = 0 Then
    
        objCodigoNome.lCodigo = lCliente
        lErro = CF("Cliente_Le_NomeRed", objCodigoNome)
        If lErro <> SUCESSO And lErro <> 12553 Then gError 130043
        If lErro <> SUCESSO Then gError 130044
        
        sUltNomeRedCli = objCodigoNome.sNome
        lUltCliente = lCliente
        
    End If
    
    lErro = MargContr_TestaDVV
    If lErro <> SUCESSO And lErro <> 130055 Then gError ERRO_SEM_MENSAGEM
    
    If lErro = 130055 Then 'faltou informacao para calcular
    
        dPercMarg = 0
        
    Else
    
        Set objTelaComissoesAux = gobjMargContr.objTelaComissoes
        
        lErro = CF("MargContr_CalculaComissoes", colComissoesAux, gobjMargContr, dPrecoComissoesAux, objTelaComissoesAux, sUltNomeRedCli, gobjMargContr.sUnidadeMed)
        If lErro <> SUCESSO Then gError 130045
    
        Set gobjMargContr.colComissoes = colComissoesAux
        gobjMargContr.dPrecoComissoes = dPrecoComissoesAux
    
        'calcula a DVV para o cliente
        lErro = MargContr_CalculaDVV(gobjMargContr.iColunaCli)
        If lErro <> SUCESSO Then gError 130047
    
        lErro = Analise_RecalcularColuna(gobjMargContr.iColunaCli, gobjMargContr.iColunaCli, gobjMargContr.iColunaPraticado, gobjMargContr.iLinhaPercMarg, gobjMargContr.iLinhaPercLucro)
        If lErro <> SUCESSO Then gError 130053
        
        'força a DVV para simulacao ser igual à do cliente
        sDVVTotal(3) = sDVVTotal(2)
    
        If iCondicaoPagto <> 0 And iCondicaoPagto <> COD_A_VISTA Then
            
            objCondicaoPagto.iCodigo = iCondicaoPagto
            lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
            If lErro <> SUCESSO And lErro <> 19205 Then gError ERRO_SEM_MENSAGEM
    
            'Se a condição de pagamento não está cadastrada, erro
            If lErro = 19205 Then gError 106675
    
            'Calcula a média ponderada de dias para a taxa financeira
            Call CF("CalculaDias_CondPagamento", objCondicaoPagto, dPrazo)
        
        End If
        
        'se for remessa em consignacao colocar + 30 dias (D9DQ1NMM7D)
        If sNaturezaOp = "6917" Then dPrazo = dPrazo + 30
        
        objProduto.sCodigo = sProduto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        If objProduto.sSiglaUMEstoque <> gobjMargContr.sUnidadeMed Then
        
            lErro = CF("UM_Conversao", objProduto.iClasseUM, gobjMargContr.sUnidadeMed, objProduto.sSiglaUMEstoque, dFator)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            gobjMargContr.dQuantidade = gobjMargContr.dQuantidade * dFator
            dPrecoVenda = dPrecoVenda / dFator
            dFrete = dFrete / dFator
            gobjMargContr.sUnidadeMed = objProduto.sSiglaUMEstoque
        
        End If
        
        lErro = Analise_RecalcularColuna2(dPercMarg, dPrecoVenda, dPrazo, dFrete)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    End If
    
    MargContr_DetCalcMarg = SUCESSO
     
    Exit Function
    
Erro_MargContr_DetCalcMarg:

    MargContr_DetCalcMarg = gErr
     
    Select Case gErr
          
        Case 130041, 130042, 130043, 130045, 130047, 130049 To 130054, ERRO_SEM_MENSAGEM
        
        Case 130044
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_ENCONTRADO", gErr, lCliente)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    Exit Function

End Function

Private Function Analise_RecalcularColuna2(dPercMarg As Double, ByVal dPrecoVenda As Double, ByVal dPrazo As Double, ByVal dFrete As Double) As Long

Dim lErro As Long, iLinha As Integer, objLinCol As ClassPlanMargContrLinCol, objLin As ClassPlanMargContrLin
Dim colPlanilhas As New Collection, objPlanilhas As ClassPlanilhas, sFormula As String
Dim objContexto As New ClassContextoPlan, dValor As Double, colPlanilhas2 As New Collection
Dim objMnemonicoValor As ClassMnemonicoValor, objContexto2 As New ClassContextoPlan, iCol As Integer
Dim dValorL1 As Double

On Error GoTo Erro_Analise_RecalcularColuna2

    dPercMarg = 0
    
    'guarda os resultados do grid dvv para todas as colunas
    Call DVVTotal_GuardaValores(objContexto)
    
    For iLinha = 1 To gobjMargContr.colPlanMargContrLin.Count
    
        Set objLin = gobjMargContr.colPlanMargContrLin(iLinha)
        Set objLinCol = gobjMargContr.colPlanMargContrLinCol(gobjMargContr.IndAnalise(iLinha, gobjMargContr.iColunaSimulacao - 1))
            
        If iLinha = 1 Then
        
            'pega conteudo da tela, da propria celula
            sFormula = CStr(dPrecoVenda)
            Call TrocaPontoVirgula(sFormula)
            
        Else
        
            If iLinha = 8 Then 'frete
            
                'pega conteudo da tela, da propria celula
                sFormula = CStr(dFrete)
                Call TrocaPontoVirgula(sFormula)
                
            Else
        
                If iLinha = gobjMargContr.iLinhaPrazo Then
                
                    sFormula = CStr(dPrazo)
                    Call TrocaPontoVirgula(sFormula)
                
                Else
                
                    sFormula = objLinCol.sFormula
                    If Len(Trim(sFormula)) = 0 Then sFormula = objLin.sFormulaGeral
            
                End If
        
            End If
        
        End If
        
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
        .dtDataCustoFixo = gobjMargContr.dtDataCF
        .iRotinaOrigem = FORMACAO_PRECO_ANALISE_MARGCONTR_REL
        Set .colComissoes = gobjMargContr.colComissoes
        .dPrecoPraticado = dPrecoVenda
        .dTaxaDescPadrao = gobjMargContr.dTaxaDescPadrao
        .dTaxaValFut = gobjMargContr.dTaxaValFut
        .dDiasValFut = gobjMargContr.dDiasValFut
    End With
    
    'Executa as formulas da planilha de preço. Retorna o valor da planilha em dValor (que é o valor da última linha da planilha) e o valor de cada linha em colPlanilhas.Item(?).dValor
    lErro = CF("Avalia_Expressao_FPreco3", colPlanilhas, dValor, objContexto)
    If lErro <> SUCESSO Then gError 106721

    For Each objPlanilhas In colPlanilhas
    
        If objPlanilhas.iLinha = gobjMargContr.iLinhaPercMarg Then
            dPercMarg = StrParaDbl(Format(objPlanilhas.dValor * 100, "###,##0.00###"))
        End If

    Next
    
    Analise_RecalcularColuna2 = SUCESSO
     
    Exit Function
    
Erro_Analise_RecalcularColuna2:

    Analise_RecalcularColuna2 = gErr
     
    Select Case gErr
          
        Case 106721
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153235)
     
    End Select
     
    Exit Function

End Function


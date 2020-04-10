VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelMargContr3"
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
Private sUltNomeRedCli As String
Private dUltQtde As Double
Private sUltUM As String
Private dUltPrazo As Double
Private dUltPrecoVenda As Double

'Valores armazenados para a última pesquisa
Private dUltPercMarg As Double

Function MargContr_ObtemMargItemNF(dValor As Double, ByVal lNumIntItemNF As Long, ByVal iColunaSimulacao As Integer, ByVal iLinhaPercMarg As Integer, ByVal iColunaDVVCliente As Integer) As Long

Dim lErro As Long, dPrecoVenda As Double, lComando As Long, iCondicaoPagto As Integer
Dim iFilialFaturamento As Integer, lCliente As Long, iFilialCli As Integer, sProduto As String, dQtde As Double, sUM As String, iTabelaPreco As Integer, dPrazo As Double
Dim objCondicaoPagto As New ClassCondicaoPagto, iLinhaPrazo As Integer
Dim objProduto As New ClassProduto, dFator As Double, sNaturezaOp As String

On Error GoTo Erro_MargContr_ObtemMargItemNF

    If UCase(gsNomePrinc) = "SGEPRINC" Then
    
        lComando = Comando_AbrirExt(GL_lConexaoRel)
        If lComando = 0 Then gError 81810
        
        sUM = String(STRING_UM_SIGLA, 0)
        sProduto = String(STRING_PRODUTO, 0)
        sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
        
        lErro = Comando_Executar(lComando, "SELECT TDI.NaturezaOp, NF.FilialEmpresa, NF.Cliente, NF.FilialCli, INF.Produto, INF.Quantidade, INF.UnidadeMed, INF.PrecoUnitario, NF.TabelaPreco, TIT.CondicaoPagto FROM TributacaoDocItem TDI, ItensNFiscal INF, NFiscal NF LEFT OUTER JOIN TitulosRecTodos TIT ON NF.ClasseDocCPR=2 AND TIT.NumIntDoc = NF.NumIntDocCPR WHERE INF.NumIntDoc = TDI.NumIntDocItem AND TDI.TipoDoc = ? AND INF.NumIntNF = NF.NumIntDoc AND INF.NumIntDoc = ?", _
            sNaturezaOp, iFilialFaturamento, lCliente, iFilialCli, sProduto, dQtde, sUM, dPrecoVenda, iTabelaPreco, iCondicaoPagto, TIPODOC_TRIB_NF, lNumIntItemNF)
        If lErro <> AD_SQL_SUCESSO Then gError 201429
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201430
            
        Call Comando_Fechar(lComando)
            
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
        
        iLinhaPrazo = 2 '???? deveria vir como parametro
        
        objProduto.sCodigo = sProduto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        If objProduto.sSiglaUMEstoque <> sUM Then
        
            lErro = CF("UM_Conversao", objProduto.iClasseUM, sUM, objProduto.sSiglaUMEstoque, dFator)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            dQtde = dQtde * dFator
            dPrecoVenda = dPrecoVenda / dFator
            sUM = objProduto.sSiglaUMEstoque
        
        End If
        
        lErro = MargContr_Verifica(iFilialFaturamento, lCliente, iFilialCli, sProduto, dQtde, sUM, iTabelaPreco, iColunaSimulacao, iLinhaPercMarg, dPrecoVenda, dPrazo, iColunaDVVCliente, iLinhaPrazo)
        If lErro <> SUCESSO Then gError 130048
        
        dValor = dUltPercMarg
    
    Else
    
        gError 216152
    
        dValor = 0
    
    End If
    
    MargContr_ObtemMargItemNF = SUCESSO
     
    Exit Function
    
Erro_MargContr_ObtemMargItemNF:

    MargContr_ObtemMargItemNF = gErr
     
    Select Case gErr
          
        Case 130048, ERRO_SEM_MENSAGEM
        
        Case 216152
            Call Rotina_Erro(vbOKOnly, "ERRO_ADRELVB_NAO_PODE_SER_CALC_FORA_SGEPRINC", gErr, "MargContr", gsNomePrinc)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153231)
     
    End Select
     
    Exit Function

End Function

Private Function MargContr_Verifica(ByVal iFilialFaturamento As Integer, ByVal lCliente As Long, ByVal iFilialCli As Integer, ByVal sProduto As String, ByVal dQtde As Double, ByVal sUM As String, ByVal iTabelaPreco As Integer, ByVal iColunaSimulacao As Integer, ByVal iLinhaPercMarg As Integer, ByVal dPrecoVenda As Double, dPrazo As Double, ByVal iColunaDVVCliente As Integer, ByVal iLinhaPrazo As Integer) As Long

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
    If iUltFilialEmpresa <> iFilialFaturamento Or sUltProduto <> sProduto Or lUltCliente <> lCliente Or iUltFilialCli <> iFilialCli Or dUltQtde <> dQtde Or sUltUM <> sUM Or dPrecoVenda <> dUltPrecoVenda Or dPrazo <> dUltPrazo Then

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
        
            dUltPercMarg = 0
        
        Else
        
            lErro = CF("MargContr_CalculaComissoes", gcolComissoes, gobjMargContr, gdPrecoComissoes, gobjTelaComissoes, sUltNomeRedCli, sUM)
            If lErro <> SUCESSO Then gError 130045
        
            'calcula a DVV para o cliente
            lErro = MargContr_CalculaDVV(iColunaDVVCliente)
            If lErro <> SUCESSO Then gError 130047
            
            lErro = Analise_RecalcularColuna(iColunaSimulacao, iLinhaPercMarg, dPrecoVenda, dPrazo, dDiasValFut, iLinhaPrazo)
            If lErro <> SUCESSO Then gError 130053
            
        End If
        
        iUltFilialEmpresa = iFilialFaturamento
        sUltProduto = sProduto
        lUltCliente = lCliente
        iUltFilialCli = iFilialCli
        dUltQtde = dQtde
        dUltPrecoVenda = dPrecoVenda
        dUltPrazo = dPrazo
        
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
    
    Set gobjMargContr = Nothing
   
    If Not gobjTelaComissoes Is Nothing Then
    
        gobjTelaComissoes.Unload gobjTelaComissoes
        
    End If
    
    Set gcolComissoes = Nothing

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
        
        sDVVTotal(iColuna) = Format(objPlanilhas.dValor * 100, "###,##0.00###") '??? tirei o -1 para colocar para a coluna de simulacao
    
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

Private Function Analise_RecalcularColuna(ByVal iColuna As Integer, ByVal iLinhaPercMarg As Integer, ByVal dPrecoVenda As Double, ByVal dPrazo As Double, ByVal dDiasValFut As Double, ByVal iLinhaPrazo As Integer) As Long

Dim lErro As Long, iLinha As Integer, objLinCol As ClassPlanMargContrLinCol, objLin As ClassPlanMargContrLin
Dim colPlanilhas As New Collection, objPlanilhas As ClassPlanilhas, sFormula As String
Dim objContexto As New ClassContextoPlan, dValor As Double, colPlanilhas2 As New Collection
Dim objMnemonicoValor As ClassMnemonicoValor, objContexto2 As New ClassContextoPlan, iCol As Integer
Dim dValorL1 As Double

On Error GoTo Erro_Analise_RecalcularColuna

    'guarda os resultados do grid dvv para todas as colunas
    Call DVVTotal_GuardaValores(objContexto)
    
    For iLinha = 1 To gobjMargContr.colPlanMargContrLin.Count
    
        Set objLin = gobjMargContr.colPlanMargContrLin(iLinha)
        Set objLinCol = gobjMargContr.colPlanMargContrLinCol(gobjMargContr.IndAnalise(iLinha, iColuna - 1))
            
        If iLinha = 1 Then
        
            'pega conteudo da tela, da propria celula
            sFormula = CStr(dPrecoVenda)
            Call TrocaPontoVirgula(sFormula)
            
        Else
        
            If iLinha = iLinhaPrazo Then
            
                sFormula = CStr(dPrazo)
                Call TrocaPontoVirgula(sFormula)
            
            Else
            
                sFormula = objLinCol.sFormula
                If Len(Trim(sFormula)) = 0 Then sFormula = objLin.sFormulaGeral
        
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
        .dtDataCustoFixo = dtDataCF
        .iRotinaOrigem = FORMACAO_PRECO_ANALISE_MARGCONTR_REL
        Set .colComissoes = gcolComissoes
        .dPrecoPraticado = dPrecoVenda
        .dTaxaDescPadrao = dTaxaDescPadrao
        .dTaxaValFut = dTaxaValFut
        .dDiasValFut = dDiasValFut
    End With
    
    'Executa as formulas da planilha de preço. Retorna o valor da planilha em dValor (que é o valor da última linha da planilha) e o valor de cada linha em colPlanilhas.Item(?).dValor
    lErro = CF("Avalia_Expressao_FPreco3", colPlanilhas, dValor, objContexto)
    If lErro <> SUCESSO Then gError 106721

    For Each objPlanilhas In colPlanilhas
    
        If objPlanilhas.iLinha = iLinhaPercMarg Then
            dUltPercMarg = StrParaDbl(Format(objPlanilhas.dValor * 100, "###,##0.00###"))
        End If

    Next
    
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

Private Function CalculaDias_PagtoPV(ByVal lNumIntPV As Long, dDias As Double) As Long
'calcula o numero medio de dias para um pedido de vendas

Dim iDiaParcelaAtual As Integer
Dim iTotalDias As Integer
Dim iNumeroParcelas As Integer
Dim iIntervalo As Integer


    CalculaDias_PagtoPV = SUCESSO

End Function




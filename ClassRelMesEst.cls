VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelMesEst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


'criada para atender às chamadas de EXT_FUNC nos relatorios atraves de pesquisa, principalmente, nas tabelas SldMesEst, SldDiaEst
'existem grupos de funcoes:
'1)que precisa apenas das totalizacoes de movimentos num periodo
'2)que precisa do saldo inicial no periodo e as movimentacoes diarias uma a uma para obtencao de estoque medio
'3)que obtem totalizacoes em sldmesest sobre os 12 ultimos meses

''IMPORTANTE:
'   grandes otimizacoes podem ser feitas para relatorios que listem uma faixa de produtos numa ordem fixa evitando a execucao de SELECTs produto a produto
'   outra otimizacao pode ser feita p/relatorios que precisem de movtos e estoque medio quando o est medio poderia pegar as duas coisas

''IMPORTANTE: NAO ESQUECA DE TRATAR DATA_NULA COMO POSSIVEL DATA INICIAL OU FINAL
'

'Identificação
Private iUltFilialEmpresa As Integer
Private dtUltDataInicial As Date
Private dtUltDataFinal As Date
Private sUltProduto As String 'último produto pesquisado. Obs.: só preencher se conseguir obter os dados solicitados, ou seja, se os "valores obtidos" abaixo forem válidos.

'Valores armazenados para o último Produto pesquisado
Private dUltQuantidadeVendida As Double
Private dUltCustoTotal As Double
Private dUltQuantidadeEntrada As Double
Private dUltQuantidadeSaida As Double
Private dUltValorEntrada As Double
Private dUltValorSaida As Double

Private lComando As Long
Private alComandoSE(0 To 4) As Long


'Identificação p/Calculo de Estoque Medio (EM)
Private iUltFilialEmpresaEM As Integer
Private dtUltDataInicialEM As Date
Private dtUltDataFinalEM As Date
Private sUltProdutoEM As String 'último produto pesquisado. Obs.: só preencher se conseguir obter os dados solicitados, ou seja, se os "valores obtidos" abaixo forem válidos.

Private dUltEstMedioQtde As Double
Private dUltEstMedioValor As Double

Private alComandoEM(0 To 1) As Long

'Identificação p/Calculo de Consumo Medio nos ultimos 12 meses
Private iUltFilialEmpresaCM As Integer
Private sUltProdutoCM As String 'último produto pesquisado. Obs.: só preencher se conseguir obter os dados solicitados, ou seja, se os "valores obtidos" abaixo forem válidos.

Private dUltConsumoMedio As Double
Private dUltVendaMediaQtde As Double

Private alComandoCM(0 To 1) As Long

'Identificação p/Calculo de Custo

Private sUltProdutoCusto As String
Private iUltFilialEmpresaCusto As Integer
Private iUltApropriacaoCusto As Integer
Private iUltAnoEstAbertoCusto As Integer
Private iUltMesEstAbertoCusto As Integer
Private iUltAnoEstApuradoCusto As Integer
Private iUltMesEstApuradoCusto As Integer
Private dUltCusto As Double

'Identificação p/Calculo de Custo Filial

Private sUltProdutoCustoFilial As String
Private iUltApropriacaoCustoFilial As Integer
Private dUltCustoFilial As Double
Private iUltFilialEmpresaCustoFilial As Integer

Public Function ObterQtdeVendProd(sProduto As String, iFilialEmpresa As Integer, dVendaMediaQtde As Double) As Long
'Para obter a venda media em qtde de uma filialempresa de um produto nos ultimos 12 meses
'só contarao p/o calculo os meses apos a implantacao do produto na filial

Dim lErro As Long

On Error GoTo Erro_ObterQtdeVendProd
    
    If sProduto <> sUltProdutoCM Or iFilialEmpresa <> iUltFilialEmpresaCM Then
        
        lErro = ObterRelMesMedia12Meses(sProduto, iFilialEmpresa)
        If lErro <> SUCESSO Then Error 41963
    
    End If
    
    dVendaMediaQtde = dUltVendaMediaQtde
    
    ObterQtdeVendProd = SUCESSO
    
    Exit Function
    
Erro_ObterQtdeVendProd:

    ObterQtdeVendProd = Err
    
    Select Case Err

        Case 41963
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153237)

    End Select

    Exit Function
    
End Function

Public Function ObterQtdeConsProd(sProduto As String, iFilialEmpresa As Integer, dConsumoMedio As Double) As Long
'Para obter o consumo medio de uma filialempresa de um produto nos ultimos 12 meses
'só contarao p/o calculo os meses apos a implantacao do produto na filial
'obs: se o campo da tabela ProdutoFilial passar a estar atualizado pode-se pegar o valor direto de lá

Dim lErro As Long

On Error GoTo Erro_ObterQtdeConsProd
    
    If sProduto <> sUltProdutoCM Or iFilialEmpresa <> iUltFilialEmpresaCM Then
        
        lErro = ObterRelMesMedia12Meses(sProduto, iFilialEmpresa)
        If lErro <> SUCESSO Then Error 41960
    
    End If
    
    dConsumoMedio = dUltConsumoMedio
    
    ObterQtdeConsProd = SUCESSO
    
    Exit Function
    
Erro_ObterQtdeConsProd:

    ObterQtdeConsProd = Err
    
    Select Case Err

        Case 41960
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153238)

    End Select

    Exit Function
    
End Function

Public Function ObterEstMedProdPerQtde(sProduto As String, dtDataInicial As Date, dtDataFinal As Date, iFilialEmpresa As Integer, dQuantidadeMedia As Double) As Long
'Para obter a qtde media ( diaria ) de estoque mantida por uma filialempresa de um produto num periodo

Dim lErro As Long

On Error GoTo Erro_ObterEstMedProdPerQtde
    
    If sProduto <> sUltProdutoEM Or dtDataInicial <> dtUltDataInicialEM Or dtDataFinal <> dtUltDataFinalEM Or iFilialEmpresa <> iUltFilialEmpresaEM Then
        
        lErro = ObterRelMesEstMedio(sProduto, dtDataInicial, dtDataFinal, iFilialEmpresa)
        If lErro <> SUCESSO Then Error 41941
    
    End If
    
    dQuantidadeMedia = dUltEstMedioQtde
    
    ObterEstMedProdPerQtde = SUCESSO
    
    Exit Function
    
Erro_ObterEstMedProdPerQtde:

    ObterEstMedProdPerQtde = Err
    
    Select Case Err

        Case 41941
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153239)

    End Select

    Exit Function
    
End Function

Public Function ObterEstMedProdPerVal(sProduto As String, dtDataInicial As Date, dtDataFinal As Date, iFilialEmpresa As Integer, dValorMedio As Double) As Long
'Para obter o valor medio ( diario ) de estoque mantido por uma filialempresa de um produto num periodo

Dim lErro As Long

On Error GoTo Erro_ObterEstMedProdPerVal
    
    If sProduto <> sUltProdutoEM Or dtDataInicial <> dtUltDataInicialEM Or dtDataFinal <> dtUltDataFinalEM Or iFilialEmpresa <> iUltFilialEmpresaEM Then
        
        lErro = ObterRelMesEstMedio(sProduto, dtDataInicial, dtDataFinal, iFilialEmpresa)
        If lErro <> SUCESSO Then Error 41942
    
    End If
    
    dValorMedio = dUltEstMedioValor
    
    ObterEstMedProdPerVal = SUCESSO
    
    Exit Function
    
Erro_ObterEstMedProdPerVal:

    ObterEstMedProdPerVal = Err
    
    Select Case Err

        Case 41942
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153240)

    End Select

    Exit Function
    
End Function

Public Function ObterQtdeVendidaProdPer(sProduto As String, dtDataInicial As Date, dtDataFinal As Date, iFilialEmpresa As Integer, dQuantidadeVendida As Double) As Long
'Para obter a qtde vendida por uma filialempresa de um produto num periodo

Dim lErro As Long

On Error GoTo Erro_ObterQtdeVendidaProdPer
    
    If sProduto <> sUltProduto Or dtDataInicial <> dtUltDataInicial Or dtDataFinal <> dtUltDataFinal Or iFilialEmpresa <> iUltFilialEmpresa Then
        
        lErro = ObterRelMesEst(sProduto, dtDataInicial, dtDataFinal, iFilialEmpresa)
        If lErro <> SUCESSO Then Error 43324
    
    End If
    
    dQuantidadeVendida = dUltQuantidadeVendida
    
    ObterQtdeVendidaProdPer = SUCESSO
    
    Exit Function
    
Erro_ObterQtdeVendidaProdPer:

    ObterQtdeVendidaProdPer = Err
    
    Select Case Err

        Case 43324
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153241)

    End Select

    Exit Function
    
End Function

Public Function ObterCustoTotalProdPer(sProduto As String, dtDataInicial As Date, dtDataFinal As Date, iFilialEmpresa As Integer, dCustoTotal As Double) As Long
'Para obter o custo do material vendido por uma filialempresa de um produto num periodo

Dim lErro As Long

On Error GoTo Erro_ObterCustoTotalProdPer
    
    If sProduto <> sUltProduto Or dtDataInicial <> dtUltDataInicial Or dtDataFinal <> dtUltDataFinal Or iFilialEmpresa <> iUltFilialEmpresa Then
        
        lErro = ObterRelMesEst(sProduto, dtDataInicial, dtDataFinal, iFilialEmpresa)
        If lErro <> SUCESSO Then Error 43325
        
    End If
    
    dCustoTotal = dUltCustoTotal
    
    ObterCustoTotalProdPer = SUCESSO
    
    Exit Function
    
Erro_ObterCustoTotalProdPer:

    ObterCustoTotalProdPer = Err
    
    Select Case Err

        Case 43325
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153242)

    End Select

    Exit Function

End Function

Public Function ObterQtdeEntradaProdPer(sProduto As String, dtDataInicial As Date, dtDataFinal As Date, iFilialEmpresa As Integer, dQuantidadeEntrada As Double) As Long
'Para obter a qtde entrada em uma filialempresa de um produto num periodo

Dim lErro As Long

On Error GoTo Erro_ObterQtdeEntradaProdPer
    
    If sProduto <> sUltProduto Or dtDataInicial <> dtUltDataInicial Or dtDataFinal <> dtUltDataFinal Or iFilialEmpresa <> iUltFilialEmpresa Then
        
        lErro = ObterRelMesEst(sProduto, dtDataInicial, dtDataFinal, iFilialEmpresa)
        If lErro <> SUCESSO Then Error 43351
    
    End If
    
    dQuantidadeEntrada = dUltQuantidadeEntrada
    
    ObterQtdeEntradaProdPer = SUCESSO
    
    Exit Function
    
Erro_ObterQtdeEntradaProdPer:

    ObterQtdeEntradaProdPer = Err
    
    Select Case Err

        Case 43351
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153243)

    End Select

    Exit Function
    
End Function

Public Function ObterQtdeSaidaProdPer(sProduto As String, dtDataInicial As Date, dtDataFinal As Date, iFilialEmpresa As Integer, dQuantidadeSaida As Double) As Long
'Para obter a qtde saida de uma filialempresa de um produto num periodo

Dim lErro As Long

On Error GoTo Erro_ObterQtdeSaidaProdPer
    
    If sProduto <> sUltProduto Or dtDataInicial <> dtUltDataInicial Or dtDataFinal <> dtUltDataFinal Or iFilialEmpresa <> iUltFilialEmpresa Then
        
        lErro = ObterRelMesEst(sProduto, dtDataInicial, dtDataFinal, iFilialEmpresa)
        If lErro <> SUCESSO Then Error 43352
    
    End If
    
    dQuantidadeSaida = dUltQuantidadeSaida
    
    ObterQtdeSaidaProdPer = SUCESSO
    
    Exit Function
    
Erro_ObterQtdeSaidaProdPer:

    ObterQtdeSaidaProdPer = Err
    
    Select Case Err

        Case 43352
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153244)

    End Select

    Exit Function
    
End Function

Public Function ObterValorEntradaProdPer(sProduto As String, dtDataInicial As Date, dtDataFinal As Date, iFilialEmpresa As Integer, dValorEntrada As Double) As Long
'Para obter o custo das entradas em uma filialempresa de um produto num periodo

Dim lErro As Long

On Error GoTo Erro_ObterValorEntradaProdPer
    
    If sProduto <> sUltProduto Or dtDataInicial <> dtUltDataInicial Or dtDataFinal <> dtUltDataFinal Or iFilialEmpresa <> iUltFilialEmpresa Then
        
        lErro = ObterRelMesEst(sProduto, dtDataInicial, dtDataFinal, iFilialEmpresa)
        If lErro <> SUCESSO Then Error 43353
    
    End If
    
    dValorEntrada = dUltValorEntrada
    
    ObterValorEntradaProdPer = SUCESSO
    
    Exit Function
    
Erro_ObterValorEntradaProdPer:

    ObterValorEntradaProdPer = Err
    
    Select Case Err

        Case 43353
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153245)

    End Select

    Exit Function
    
End Function

Public Function ObterValorSaidaProdPer(sProduto As String, dtDataInicial As Date, dtDataFinal As Date, iFilialEmpresa As Integer, dValorSaida As Double) As Long
'Para obter o custo das saidas uma filialempresa de um produto num periodo

Dim lErro As Long

On Error GoTo Erro_ObterValorSaidaProdPer
    
    If sProduto <> sUltProduto Or dtDataInicial <> dtUltDataInicial Or dtDataFinal <> dtUltDataFinal Or iFilialEmpresa <> iUltFilialEmpresa Then
        
        lErro = ObterRelMesEst(sProduto, dtDataInicial, dtDataFinal, iFilialEmpresa)
        If lErro <> SUCESSO Then Error 43354
    
    End If
    
    dValorSaida = dUltValorSaida
    
    ObterValorSaidaProdPer = SUCESSO
    
    Exit Function
    
Erro_ObterValorSaidaProdPer:

    ObterValorSaidaProdPer = Err
    
    Select Case Err

        Case 43354
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153246)

    End Select

    Exit Function
    
End Function

Private Function ObterRelMesEst(sProduto As String, dtDataInicial As Date, dtDataFinal As Date, iFilialEmpresa As Integer) As Long
'Para obter os valores a partir de sldmesest/slddiaest e torna-los disponiveis para as outras funcoes

Dim lErro As Long, dQuantidadeVendida As Double, dCustoTotal As Double, dValorFaturado As Double
Dim dQuantidadeEntrada As Double, dQuantidadeSaida As Double, dValorEntrada As Double, dValorSaida As Double
Dim dQuantidadeVendidaSemEst As Double, dCustoTotalSemEst As Double

On Error GoTo Erro_ObterRelMesEst

    lErro = CF("Produto_ObterRelMesEst", sProduto, dtDataInicial, dtDataFinal, iFilialEmpresa, dQuantidadeVendida, dCustoTotal, dQuantidadeEntrada, dQuantidadeSaida, dValorEntrada, dValorSaida, lComando)
    If lErro <> SUCESSO Then Error 43327
    
    lErro = CF("Produto_ObterRelMesSemEst", sProduto, dtDataInicial, dtDataFinal, iFilialEmpresa, dQuantidadeVendidaSemEst, dCustoTotalSemEst, alComandoSE(0), alComandoSE(1), alComandoSE(2), alComandoSE(3), alComandoSE(4))
    If lErro <> SUCESSO Then Error 43327
    
    sUltProduto = sProduto
    dtUltDataInicial = dtDataInicial
    dtUltDataFinal = dtDataFinal
    iUltFilialEmpresa = iFilialEmpresa
    
    dUltQuantidadeVendida = dQuantidadeVendida + dQuantidadeVendidaSemEst
    dUltCustoTotal = dCustoTotal + dCustoTotalSemEst
    dUltQuantidadeEntrada = dQuantidadeEntrada
    dUltQuantidadeSaida = dQuantidadeSaida
    dUltValorEntrada = dValorEntrada
    dUltValorSaida = dValorSaida
    
    ObterRelMesEst = SUCESSO
    
    Exit Function
    
Erro_ObterRelMesEst:

    ObterRelMesEst = Err
    
    Select Case Err

        Case 43327
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153247)

    End Select

    Exit Function
    
End Function

Private Function ObterRelMesEstMedio(sProduto As String, dtDataInicial As Date, dtDataFinal As Date, iFilialEmpresa As Integer) As Long
'Para obter os valores a partir de sldmesest/slddiaest e torna-los disponiveis para as outras funcoes

Dim lErro As Long
Dim dQuantidade As Double, dValor As Double

On Error GoTo Erro_ObterRelMesEstMedio

    lErro = CF("Produto_ObterRelMesEstMedio", sProduto, dtDataInicial, dtDataFinal, iFilialEmpresa, dQuantidade, dValor, alComandoEM)
    If lErro <> SUCESSO Then Error 41949
    
    sUltProdutoEM = sProduto
    dtUltDataInicialEM = dtDataInicial
    dtUltDataFinalEM = dtDataFinal
    iUltFilialEmpresaEM = iFilialEmpresa
    
    dUltEstMedioQtde = dQuantidade
    dUltEstMedioValor = dValor
    
    ObterRelMesEstMedio = SUCESSO
    
    Exit Function
    
Erro_ObterRelMesEstMedio:

    ObterRelMesEstMedio = Err
    
    Select Case Err

        Case 41949
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153248)

    End Select

    Exit Function
    
End Function

Private Function ObterRelMesMedia12Meses(sProduto As String, iFilialEmpresa As Integer) As Long
'Para obter o consumo medio

Dim lErro As Long
Dim dQtdeConsumo As Double, dQtdeVendida As Double

On Error GoTo Erro_ObterRelMesMedia12Meses

    lErro = CF("Produto_ObterRelMesMedia12Meses", sProduto, iFilialEmpresa, dQtdeConsumo, dQtdeVendida, alComandoCM)
    If lErro <> SUCESSO Then Error 41961
    
    sUltProdutoCM = sProduto
    iUltFilialEmpresaCM = iFilialEmpresa
    
    dUltConsumoMedio = dQtdeConsumo
    dUltVendaMediaQtde = dQtdeVendida
    
    ObterRelMesMedia12Meses = SUCESSO
    
    Exit Function
    
Erro_ObterRelMesMedia12Meses:

    ObterRelMesMedia12Meses = Err
    
    Select Case Err

        Case 41961
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153249)

    End Select

    Exit Function
    
End Function

Public Function ObterRelCusto(dCusto As Double, ByVal iFilialEmpresa As Integer, ByVal iApropriacaoCusto As Integer, ByVal iAnoEstAberto As Integer, ByVal iMesEstAberto As Integer, ByVal iAnoEstApurado As Integer, ByVal iMesEstApurado As Integer, ByVal sProduto As String) As Long
'Para obter o custo do produtos passado como parametro e torna-lo disponivel para as outras funcoes

Dim lErro As Long
Dim objSldMesEst As New ClassSldMesEst

On Error GoTo Erro_ObterRelCusto

    If iFilialEmpresa <> iUltFilialEmpresaCusto Or iUltApropriacaoCusto <> iApropriacaoCusto Or iUltAnoEstAbertoCusto <> iAnoEstAberto Or iUltMesEstAbertoCusto <> iMesEstAberto Or _
    iUltAnoEstApuradoCusto <> iAnoEstApurado Or iUltMesEstApuradoCusto <> iMesEstApurado Or sUltProdutoCusto <> sProduto Then

        Select Case iApropriacaoCusto
            
            Case APROPR_CUSTO_REAL
                objSldMesEst.iAno = iAnoEstApurado
                
            Case APROPR_CUSTO_MEDIO, APROPR_CUSTO_STANDARD
                objSldMesEst.iAno = iAnoEstAberto
            
        End Select
        
        objSldMesEst.sProduto = sProduto
        objSldMesEst.iFilialEmpresa = iFilialEmpresa

        lErro = CF("SldMesEst_Le", objSldMesEst)
        If lErro <> SUCESSO And lErro <> 25429 Then gError 92903
        
        If lErro = SUCESSO Then

            lErro = CF("Calcula_Custo", dCusto, iFilialEmpresa, iApropriacaoCusto, iAnoEstAberto, iMesEstAberto, iAnoEstApurado, iMesEstApurado, sProduto)
            If lErro <> SUCESSO Then gError 92904
            
        Else

            dCusto = 0

        End If

        sUltProdutoCusto = sProduto
        iUltFilialEmpresaCusto = iFilialEmpresa
        iUltApropriacaoCusto = iApropriacaoCusto
        iUltAnoEstAbertoCusto = iAnoEstAberto
        iUltMesEstAbertoCusto = iMesEstAberto
        iUltAnoEstApuradoCusto = iAnoEstApurado
        iUltMesEstApuradoCusto = iMesEstApurado
        dUltCusto = dCusto
    
    Else
    
        dCusto = dUltCusto
        
    End If
    
    ObterRelCusto = SUCESSO
    
    Exit Function
    
Erro_ObterRelCusto:

    ObterRelCusto = gErr
    
    Select Case gErr

        Case 92903, 92904
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153250)

    End Select

    Exit Function
    
End Function

Public Function ObterRelCustoFilial(dCusto As Double, ByVal iFilialEmpresa As Integer, ByVal iApropriacaoCusto As Integer, ByVal sProduto As String) As Long
'Para obter o custo do produtos passado como parametro

Dim lErro As Long
Dim objEstoqueMes As New ClassEstoqueMes
Dim iAnoEstAberto As Integer, iMesEstAberto As Integer, iAnoEstApurado As Integer, iMesEstApurado As Integer
Dim objSldMesEst As New ClassSldMesEst

On Error GoTo Erro_ObterRelCustoFilial

    If iFilialEmpresa <> iUltFilialEmpresaCustoFilial Or iUltApropriacaoCustoFilial <> iApropriacaoCusto Or sProduto <> sUltProdutoCustoFilial Then

'        If iFilialEmpresa <> iUltFilialEmpresaCustoFilial Then

            objEstoqueMes.iFilialEmpresa = iFilialEmpresa
            
            '???? Pode ser melhorado
            If iApropriacaoCusto = APROPR_CUSTO_MEDIO Or iApropriacaoCusto = APROPR_CUSTO_STANDARD Then
            
                'Ler o mês e o ano que esta aberto passando como parametro filialEmpresa  e Fechamento
                lErro = CF("EstoqueMes_Le_Aberto", objEstoqueMes)
                If lErro <> SUCESSO And lErro <> 40673 Then gError 55154
            
                If lErro = 40673 Then gError 55155

                iAnoEstAberto = objEstoqueMes.iAno
                iMesEstAberto = objEstoqueMes.iMes

            End If


            If iApropriacaoCusto = APROPR_CUSTO_REAL Then
            

                objEstoqueMes.iFilialEmpresa = giFilialEmpresa
    
                'le o ultimo ano/mes apurado
                lErro = CF("EstoqueMes_Le_Apurado", objEstoqueMes)
                If lErro <> SUCESSO And lErro <> 46225 Then gError 55156
            
                If lErro = 46225 Then
                    iAnoEstApurado = 0
                    iMesEstApurado = 0
                Else
                    iAnoEstApurado = objEstoqueMes.iAno
                    iMesEstApurado = objEstoqueMes.iMes
                End If
            
            End If
            
            iUltFilialEmpresaCustoFilial = iFilialEmpresa

'        End If

        Select Case iApropriacaoCusto
            
            Case APROPR_CUSTO_REAL
                objSldMesEst.iAno = iAnoEstApurado
                
            Case APROPR_CUSTO_MEDIO, APROPR_CUSTO_STANDARD
                objSldMesEst.iAno = iAnoEstAberto
            
        End Select
        
        objSldMesEst.sProduto = sProduto
        objSldMesEst.iFilialEmpresa = iFilialEmpresa

        lErro = CF("SldMesEst_Le", objSldMesEst)
        If lErro <> SUCESSO And lErro <> 25429 Then gError 92905

        If lErro = SUCESSO Then

            lErro = CF("Calcula_Custo", dCusto, iFilialEmpresa, iApropriacaoCusto, iAnoEstAberto, iMesEstAberto, iAnoEstApurado, iMesEstApurado, sProduto)
            If lErro <> SUCESSO Then gError 55157

        Else
        
            dCusto = 0
            
        End If

        sUltProdutoCustoFilial = sProduto
        iUltApropriacaoCustoFilial = iApropriacaoCusto
        dUltCustoFilial = dCusto
    
    Else
    
        dCusto = dUltCustoFilial
        
    End If
    
    ObterRelCustoFilial = SUCESSO
    
    Exit Function
    
Erro_ObterRelCustoFilial:

    ObterRelCustoFilial = gErr
    
    Select Case gErr

        Case 55154, 55156, 55157, 92905
        
        Case 55155
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NAOEXISTE_MES_ABERTO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153251)

    End Select

    Exit Function
    
End Function


Private Sub Class_Terminate()

    If lComando <> 0 Then Call Comando_Fechar(lComando)
    If alComandoEM(0) <> 0 Then Call Comando_Fechar(alComandoEM(0))
    If alComandoEM(1) <> 0 Then Call Comando_Fechar(alComandoEM(1))
    If alComandoCM(0) <> 0 Then Call Comando_Fechar(alComandoCM(0))
    If alComandoCM(1) <> 0 Then Call Comando_Fechar(alComandoCM(1))
    If alComandoSE(0) <> 0 Then Call Comando_Fechar(alComandoSE(0))
    If alComandoSE(1) <> 0 Then Call Comando_Fechar(alComandoSE(1))
    If alComandoSE(2) <> 0 Then Call Comando_Fechar(alComandoSE(2))
    If alComandoSE(3) <> 0 Then Call Comando_Fechar(alComandoSE(3))
    If alComandoSE(4) <> 0 Then Call Comando_Fechar(alComandoSE(4))
    
End Sub

'*****************************************************************
'*****************************************************************
'************** observacoes para otimizacoes futuras *************
'*****************************************************************
'*****************************************************************
''''valores armazenados para otimizar a performance
'''Private UltAnoFechado As Integer 'obtido da tabela EstoqueMes
'''Private UltMesFechado As Integer 'obtido da tabela EstoqueMes
'''Private UltAnoApurado As Integer 'obtido da tabela EstoqueMes
'''Private UltMesApurado As Integer 'obtido da tabela EstoqueMes
'''Private iAnoSaldoIniValido As Integer 'o ano menor ou igual ao da data inicial que contenha saldo inicial valido
'''Private iLeAnoAnterior As Integer 'se precisa ler anos anteriores para obter "saldo inicial valido"
'''Private lComandoAnoAnt As Long
'''Private iUsaDiaEst As Integer 'se precisa ler a tabela SldDiaEst
'''Private lComandoDiaEst As Long
'''
'''Private alComando(10) As Long '??? 10?
'''
'''Private Function Estoque_ObterAnoSaldoIniValido(iFilialEmpresa As Integer)
''''obter o ano menor ou igual ao da data inicial que contenha saldo inicial valido
'''
'''    'se data inicial = DATA_NULA ...
'''
'''        'o ano é o 1o da tabela EstoqueMes para a filial
'''
'''    'senao
'''
'''        'obter o ano da data inicial do periodo
'''
'''        '"SELECT Ano, Mes, Fechamento FROM EstoqueMes WHERE FilialEmpresa = ? AND Ano <= ? ORDER BY Ano DESC, Mes DESC"
'''
'''        'percorrer o resultado testando "Fechamento" e obter a resposta
'''        'se nenhum mes estiver fechado, o ano valido é o do ultimo registro lido
'''
'''End Function
'
'Relatório Análise da Movimentação de Estoque: periodo entre dias
'1 - Ext_Func para:
'
'      .Saldo das Movimentacoes do Periodo (Quantidade e Valor): quanto entrou menos qto saiu
'
'    Algoritmo Sugerido:
'
'        Usar apenas 1 select envolvendo somatorio em slddiaest da data inicial a final
'
'    Algoritmo otimizado: (versao futura)
'
'        Na 1a vez que for calcular considerar:
'            DATA_NULA pode ser data inicial ou final
'        'obter:
'            se a data final é do mesmo ano-mes que a data inicial
'            se a data inicial é 1o dia de um mes
'            se a data final é ultimo dia do mes (cuidado com ano bissexto)
'            se a data final é do mes seguinte ao da data inicial
'
'        se a data inicial nao for a do 1o dia de um mes:
'            se a data final for no mesmo mes que a data inicial:
'                => usar apenas slddiaest obtendo somatorio do movto entre as datas
'            senao
'                usar slddiaest obtendo somatorio da movto da data inicial ao fim do mes
'                se a data final for ult dia de mes:
'                    => usar sldmesest obtendo somatorio do movto entre o mes seguinte ao da data inicial e o mes da data final
'                Else
'                    se a data final nao é do mes seguinte ao da data inicial
'                        => usar sldmesest obtendo somatorio do movto entre o mes seguinte ao da data inicial e o mes anterior ao da data final
'                    usar slddiaest obtendo somatorio do movto entre o 1o dia do mes da data final e a data final
'                End If
'        Else
'            se a data final for no mesmo mes que a data inicial:
'                se a data final for ult dia de mes:
'                    => usar apenas sldmesest obtendo movto no mes
'                senao
'                    => usar apenas slddiaest obtendo somatorio do movto entre as datas
'            senao
'                se a data final for ult dia de mes:
'                    => usar sldmesest obtendo somatorio do movto entre o mes da data inicial e o mes da data final
'                Else
'                    se a data final nao é do mes seguinte ao da data inicial
'                        => usar sldmesest obtendo somatorio do movto entre o mes da data inicial e o mes anterior ao da data final
'                    usar slddiaest obtendo somatorio do movto entre o 1o dia do mes da data final e a data final
'                End If
'        End If
'
'        Em resumo, qualquer dos caminhos acima pode envolver, em sequencia, apenas os comandos abaixo:
'            acesso a slddiaest
'            acesso a sldmesest
'            acesso a slddiaest
'        variando os limites (datas). Portanto, a partir da analise acima pode-se preencher um array com
'            3 elementos indicando se o comando correspondente deve ser executado e com que parametros.
'
'        Com este array preenchido, cada vez que for necessario calcular basta percorrer o array executando os comandos necessarios com os parametros devidos.
'



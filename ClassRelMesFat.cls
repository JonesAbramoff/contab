VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelMesFat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
 Option Explicit

'ALTERAÇÕES INPAL
'Márcio 09/2000 - incluido cliente e Vendedor

'criada para atender às chamadas de EXT_FUNC nos relatorios atraves de pesquisa, principalmente, nas tabelas SldMesEst, SldDiaEst
'
'IMPORTANTE: NAO ESQUECA DE TRATAR DATA_NULA COMO POSSIVEL DATA INICIAL OU FINAL

'Identificação
Private iUltFilialEmpresa As Integer
Private dtUltDataInicial As Date
Private dtUltDataFinal As Date
Private sUltProduto As String 'último produto pesquisado. Obs.: só preencher se conseguir obter os dados solicitados, ou seja, se os "valores obtidos" abaixo forem válidos.

'Márcio 09/2000 - incluido cliente e Vendedor
Private iUltVendedor As Integer
Private lUltCliente As Long

'Valores armazenados para o último Produto pesquisado
Private dUltValorFaturado As Double
Private dUltDesconto As Double
Private dUltQuantidadePedida As Double
Private dUltQuantidadeFaturada As Double
Private dUltQuantidadeDevolvida As Double
Private dUltValorDevolvido As Double
Private dUltValorCusto As Double

Private dUltQuantFaturada As Double
Private dUltQuantDevolvida As Double


Private lComando As Long
Private lComando2 As Long

Private iUltFilialCli As Integer
Private dtUltData As Date
Private iUltCodRegiao As Integer
Private iUltMes As Integer
Dim bUltSemDev As Boolean

Private iUltFilialEmpresaCusto As Integer
Private dtUltDataInicialCusto As Date
Private dtUltDataFinalCusto As Date
Private sUltProdutoCusto As String 'último produto pesquisado. Obs.: só preencher se conseguir obter os dados solicitados, ou seja, se os "valores obtidos" abaixo forem válidos.


Public Function ObterValorFatProdPer(sProduto As String, dtDataInicial As Date, dtDataFinal As Date, iFilialEmpresa As Integer, dValorFaturado As Double, Optional ByVal bSemDev As Boolean = False) As Long
'Para obter o valor faturado por uma filialempresa de um produto num periodo

Dim lErro As Long

On Error GoTo Erro_ObterValorFatProdPer
    
    If bUltSemDev <> bSemDev Or sProduto <> sUltProduto Or dtDataInicial <> dtUltDataInicial Or dtDataFinal <> dtUltDataFinal Or iFilialEmpresa <> iUltFilialEmpresa Then
        
        lErro = ObterRelMesFat(sProduto, dtDataInicial, dtDataFinal, iFilialEmpresa, bSemDev)
        If lErro <> SUCESSO Then Error 43319
        
        bUltSemDev = bSemDev
    
    End If
    
    dValorFaturado = dUltValorFaturado
    
    ObterValorFatProdPer = SUCESSO
    
    Exit Function
    
Erro_ObterValorFatProdPer:

    ObterValorFatProdPer = Err
    
    Select Case Err

        Case 43319
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153265)

    End Select

    Exit Function
    
End Function

Public Function ObterDescontoFatProdPer(sProduto As String, dtDataInicial As Date, dtDataFinal As Date, iFilialEmpresa As Integer, dDesconto As Double) As Long
'Para obter o valor de descontos no faturamento de uma filialempresa de um produto num periodo

Dim lErro As Long

On Error GoTo Erro_ObterDescontoFatProdPer
    
    If sProduto <> sUltProduto Or dtDataInicial <> dtUltDataInicial Or dtDataFinal <> dtUltDataFinal Or iFilialEmpresa <> iUltFilialEmpresa Then
        
        lErro = ObterRelMesFat(sProduto, dtDataInicial, dtDataFinal, iFilialEmpresa)
        If lErro <> SUCESSO Then Error 43320
        
    End If
    
    dDesconto = dUltDesconto
    
    ObterDescontoFatProdPer = SUCESSO
    
    Exit Function
    
Erro_ObterDescontoFatProdPer:

    ObterDescontoFatProdPer = Err
    
    Select Case Err

        Case 43320
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153266)

    End Select

    Exit Function

End Function

Public Function ObterQtdePedidaFatProdPer(sProduto As String, dtDataInicial As Date, dtDataFinal As Date, iFilialEmpresa As Integer, dQuantidadePedida As Double) As Long
'Para a qtde em pedidos de venda de uma filialempresa para um produto num periodo

Dim lErro As Long

On Error GoTo Erro_ObterQtdePedidaFatProdPer
    
    If sProduto <> sUltProduto Or dtDataInicial <> dtUltDataInicial Or dtDataFinal <> dtUltDataFinal Or iFilialEmpresa <> iUltFilialEmpresa Then
        
        lErro = ObterRelMesFat(sProduto, dtDataInicial, dtDataFinal, iFilialEmpresa)
        If lErro <> SUCESSO Then Error 43321
        
    End If
    
    dQuantidadePedida = dUltQuantidadePedida
    
    ObterQtdePedidaFatProdPer = SUCESSO
    
    Exit Function
    
Erro_ObterQtdePedidaFatProdPer:

    ObterQtdePedidaFatProdPer = Err
    
    Select Case Err

        Case 43321
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153267)

    End Select

    Exit Function

End Function

Public Function ObterQtdeFaturadaFatProdPer(sProduto As String, dtDataInicial As Date, dtDataFinal As Date, iFilialEmpresa As Integer, dQuantidadeFaturada As Double) As Long
'Para obter a qtde faturada por uma filialempresa de um produto num periodo

Dim lErro As Long

On Error GoTo Erro_ObterQtdeFaturadaFatProdPer
    
    If sProduto <> sUltProduto Or dtDataInicial <> dtUltDataInicial Or dtDataFinal <> dtUltDataFinal Or iFilialEmpresa <> iUltFilialEmpresa Then
        
        lErro = ObterRelMesFat(sProduto, dtDataInicial, dtDataFinal, iFilialEmpresa)
        If lErro <> SUCESSO Then Error 43322
    
    End If
    
    dQuantidadeFaturada = dUltQuantidadeFaturada
    
    ObterQtdeFaturadaFatProdPer = SUCESSO
    
    Exit Function
    
Erro_ObterQtdeFaturadaFatProdPer:

    ObterQtdeFaturadaFatProdPer = Err
    
    Select Case Err

        Case 43322
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153268)

    End Select

    Exit Function

End Function

Public Function ObterQtdeDevolvidaFatProdPer(sProduto As String, dtDataInicial As Date, dtDataFinal As Date, iFilialEmpresa As Integer, dQuantidadeDevolvida As Double) As Long
'Para obter a qtde devolvida por uma filialempresa de um produto num periodo

Dim lErro As Long

On Error GoTo Erro_ObterQtdeDevolvidaFatProdPer
    
    If sProduto <> sUltProduto Or dtDataInicial <> dtUltDataInicial Or dtDataFinal <> dtUltDataFinal Or iFilialEmpresa <> iUltFilialEmpresa Then
        
        lErro = ObterRelMesFat(sProduto, dtDataInicial, dtDataFinal, iFilialEmpresa)
        If lErro <> SUCESSO Then Error 43349
    
    End If
    
    dQuantidadeDevolvida = dUltQuantidadeDevolvida
    
    ObterQtdeDevolvidaFatProdPer = SUCESSO
    
    Exit Function
    
Erro_ObterQtdeDevolvidaFatProdPer:

    ObterQtdeDevolvidaFatProdPer = Err
    
    Select Case Err

        Case 43349
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153269)

    End Select

    Exit Function

End Function

Private Function ObterRelMesFat(sProduto As String, dtDataInicial As Date, dtDataFinal As Date, iFilialEmpresa As Integer, Optional ByVal bSemDev As Boolean = False) As Long
'Para obter os valores a partir de sldmesfat/slddiafat e torna-los disponiveis para as outras funcoes

Dim lErro As Long, dValorFaturado As Double, dDesconto As Double, dQuantidadePedida As Double, dQuantidadeFaturada As Double, dQuantidadeDevolvida As Double

On Error GoTo Erro_ObterRelMesFat

    If Not bSemDev Then
        lErro = CF("Produto_ObterRelMesFat", sProduto, dtDataInicial, dtDataFinal, iFilialEmpresa, dValorFaturado, dDesconto, dQuantidadePedida, dQuantidadeFaturada, dQuantidadeDevolvida, lComando)
    Else
        lErro = CF("Produto_ObterRelMesFat_SemDev", sProduto, dtDataInicial, dtDataFinal, iFilialEmpresa, dValorFaturado, dDesconto, dQuantidadePedida, dQuantidadeFaturada, dQuantidadeDevolvida, lComando)
    End If
    If lErro <> SUCESSO Then Error 43217
    
    sUltProduto = sProduto
    dtUltDataInicial = dtDataInicial
    dtUltDataFinal = dtDataFinal
    iUltFilialEmpresa = iFilialEmpresa
    
    dUltValorFaturado = dValorFaturado
    dUltDesconto = dDesconto
    dUltQuantidadePedida = dQuantidadePedida
    dUltQuantidadeFaturada = dQuantidadeFaturada
    dUltQuantidadeDevolvida = dQuantidadeDevolvida
        
    ObterRelMesFat = SUCESSO
    
    Exit Function
    
Erro_ObterRelMesFat:

    ObterRelMesFat = Err
    
    Select Case Err

        Case 43217
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153270)

    End Select

    Exit Function
    
End Function

Private Sub Class_Terminate()

    If lComando <> 0 Then Call Comando_Fechar(lComando)
    If lComando2 <> 0 Then Call Comando_Fechar(lComando2)

End Sub



'********* Inicio - Márcio 09/2000 - incluido cliente e Vendedor ********

Public Function ObterQtdeDevolvidaFatProdVendCliPer(sProduto As String, dtDataInicial As Date, dtDataFinal As Date, iFilialEmpresa As Integer, iVendedor As Integer, lCliente As Long, dQuantidadeDevolvida As Double) As Long
'Para obter a qtde devolvida por uma filialempresa de um produto num periodo

Dim lErro As Long

On Error GoTo Erro_ObterQtdeDevolvidaFatProdVendCliPer
    
    If sProduto <> sUltProduto Or dtDataInicial <> dtUltDataInicial Or dtDataFinal <> dtUltDataFinal Or iFilialEmpresa <> iUltFilialEmpresa And iVendedor <> iUltVendedor And lCliente <> lUltCliente Then
        
        lErro = ObterRelMesFatVendCli(sProduto, dtDataInicial, dtDataFinal, iFilialEmpresa, iVendedor, lCliente)
        If lErro <> SUCESSO Then gError 500216
    
    End If
    
    dQuantidadeDevolvida = dUltQuantidadeDevolvida
    
    ObterQtdeDevolvidaFatProdVendCliPer = SUCESSO
    
    Exit Function
    
Erro_ObterQtdeDevolvidaFatProdVendCliPer:

    ObterQtdeDevolvidaFatProdVendCliPer = gErr
    
    Select Case gErr

        Case 500216
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153271)

    End Select

    Exit Function

End Function

Public Function ObterQtdeFaturadaFatProdVendCliPer(sProduto As String, dtDataInicial As Date, dtDataFinal As Date, iFilialEmpresa As Integer, iVendedor As Integer, lCliente As Long, dQuantidadeFaturada As Double) As Long
'Para obter a qtde faturada por uma filialempresa de um produto num periodo

Dim lErro As Long

On Error GoTo Erro_ObterQtdeFaturadaFatProdVendCliPer
    
    If sProduto <> sUltProduto Or dtDataInicial <> dtUltDataInicial Or dtDataFinal <> dtUltDataFinal Or iFilialEmpresa <> iUltFilialEmpresa And iUltVendedor <> iVendedor And lUltCliente <> lCliente Then
        
        lErro = ObterRelMesFatVendCli(sProduto, dtDataInicial, dtDataFinal, iFilialEmpresa, iVendedor, lCliente)
        If lErro <> SUCESSO Then gError 500224
    
    End If
    
    dQuantidadeFaturada = dUltQuantidadeFaturada
    
    ObterQtdeFaturadaFatProdVendCliPer = SUCESSO
    
    Exit Function
    
Erro_ObterQtdeFaturadaFatProdVendCliPer:

    ObterQtdeFaturadaFatProdVendCliPer = gErr
    
    Select Case gErr

        Case 500224
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153272)

    End Select

    Exit Function

End Function

Public Function ObterValorFatProdVendCliPer(sProduto As String, dtDataInicial As Date, dtDataFinal As Date, iFilialEmpresa As Integer, iVendedor As Integer, lCliente As Long, dValorFaturado As Double) As Long
'Para obter o valor faturado por uma filialempresa de um produto num periodo

Dim lErro As Long

On Error GoTo Erro_ObterValorFatProdVendCliPer
    
    If sProduto <> sUltProduto Or dtDataInicial <> dtUltDataInicial Or dtDataFinal <> dtUltDataFinal Or iFilialEmpresa <> iUltFilialEmpresa And iVendedor <> iUltVendedor And lCliente <> lUltCliente Then
        
        lErro = ObterRelMesFatVendCli(sProduto, dtDataInicial, dtDataFinal, iFilialEmpresa, iVendedor, lCliente)
        If lErro <> SUCESSO Then gError 500225
    
    End If
    
    dValorFaturado = dUltValorFaturado
    
    ObterValorFatProdVendCliPer = SUCESSO
    
    Exit Function
    
Erro_ObterValorFatProdVendCliPer:

    ObterValorFatProdVendCliPer = gErr
    
    Select Case gErr

        Case 500225
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153273)

    End Select

    Exit Function
    
End Function

Private Function ObterRelMesFatVendCli(sProduto As String, dtDataInicial As Date, dtDataFinal As Date, iFilialEmpresa As Integer, iVendedor As Integer, lCliente As Long) As Long
'Para obter os valores a partir de sldmesfat/slddiafat e torna-los disponiveis para as outras funcoes

Dim lErro As Long, dValorFaturado As Double, dQuantidadeFaturada As Double, dQuantidadeDevolvida As Double

On Error GoTo Erro_ObterRelMesFatVendCli

    lErro = CF("Produto_ObterRelMesFatVendCli", sProduto, dtDataInicial, dtDataFinal, iFilialEmpresa, iVendedor, lCliente, dValorFaturado, dQuantidadeFaturada, dQuantidadeDevolvida, lComando, lComando2)
    If lErro <> SUCESSO Then gError 500217
    
    sUltProduto = sProduto
    dtUltDataInicial = dtDataInicial
    dtUltDataFinal = dtDataFinal
    iUltFilialEmpresa = iFilialEmpresa
    
    dUltValorFaturado = dValorFaturado
    dUltQuantidadeFaturada = dQuantidadeFaturada
    dUltQuantidadeDevolvida = dQuantidadeDevolvida
        
    iUltVendedor = iVendedor
    lUltCliente = lCliente
    
    ObterRelMesFatVendCli = SUCESSO
    
    Exit Function
    
Erro_ObterRelMesFatVendCli:

    ObterRelMesFatVendCli = gErr
    
    Select Case gErr

        Case 500217
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153274)

    End Select

    Exit Function
    
End Function

'Function Produto_ObterCutoRelMesFat(ByVal sProduto As String, ByVal dtDataInicial As Date, ByVal dtDataFinal As Date, ByVal iFilialEmpresa As Integer, dValorCusto As Double, lComando As Long) As Long

Private Function ObterCustoRelMesFat(sProduto As String, dtDataInicial As Date, dtDataFinal As Date, iFilialEmpresa As Integer) As Long
'Para obter os valores a partir de movimentoestoque e torna-los disponiveis para as outras funcoes

Dim lErro As Long, dValorCusto As Double

On Error GoTo Erro_ObterCustoRelMesFat

'    MsgBox ("entrou em ObterCustoRelMesFat")

    lErro = CF("Produto_ObterCustoRelMesFat", sProduto, dtDataInicial, dtDataFinal, iFilialEmpresa, dValorCusto)
    If lErro <> SUCESSO Then gError 204873
    
    sUltProdutoCusto = sProduto
    dtUltDataInicialCusto = dtDataInicial
    dtUltDataFinalCusto = dtDataFinal
    iUltFilialEmpresaCusto = iFilialEmpresa
    
    dUltValorCusto = dValorCusto
        
 '   MsgBox ("vai sair de ObterCustoRelMesFat")
    
'    MsgBox ("dVUltValorCusto = " & dUltValorCusto)
        
    ObterCustoRelMesFat = SUCESSO
    
    Exit Function
    
Erro_ObterCustoRelMesFat:

    ObterCustoRelMesFat = gErr
    
    Select Case gErr

        Case 204873
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 204874)

    End Select

    Exit Function
    
End Function

Public Function ObterCustoFatProdPer(sProduto As String, dtDataInicial As Date, dtDataFinal As Date, iFilialEmpresa As Integer, dValorCusto As Double) As Long
'Para obter o custo por uma filialempresa de um produto num periodo

Dim lErro As Long

On Error GoTo Erro_ObterCustoFatProdPer
    
 '   MsgBox ("entrou em ObterCustoFatProdPer de ClassRelMesFat")
    
 '   MsgBox "sProduto = " & sProduto & " sUltProduto = " & sUltProduto & " dtDataFinal = " & CStr(dtDataFinal) & " dtUltDataFinal = " & CStr(dtUltDataFinal)
    
    If sProduto <> sUltProdutoCusto Or dtDataInicial <> dtUltDataInicialCusto Or dtDataFinal <> dtUltDataFinalCusto Or iFilialEmpresa <> iUltFilialEmpresaCusto Then
        
'        MsgBox ("vai chamar ObterCustoRelMesFat")
        
        lErro = ObterCustoRelMesFat(sProduto, dtDataInicial, dtDataFinal, iFilialEmpresa)
        If lErro <> SUCESSO Then gError 204875
        
  '       MsgBox ("chamou ObterCustoRelMesFat")
        
    End If
    
  '  MsgBox ("vai sair de ObterCustoFatProdPer de ClassRelMesFat")
    
    dValorCusto = dUltValorCusto
    
  '  MsgBox ("dValorCusto = " & dValorCusto)
    
    ObterCustoFatProdPer = SUCESSO
    
    Exit Function
    
Erro_ObterCustoFatProdPer:

    ObterCustoFatProdPer = gErr
    
    Select Case gErr

        Case 204875
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 204876)

    End Select

    Exit Function
    
End Function


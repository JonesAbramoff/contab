VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelTitPag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Criada para atender emissão de Notas Fiscais

Private lNumIntUltimoTitulo As Long
Private colParcPag As New ColParcelaPagar
Private lComando As Long

'Variaveis Globais
Private lFornSaldoAntUlt As Double
Private dSaldoAdiantamentosUlt As Double
Private lComandoFornAnt As Long
Private lFornSaldoCredUlt As Double
Private dSaldoCreditosUlt As Double
Private lComandoFornCred As Long
Private lComandoFornCredTotal As Long
Private lComandoFornAntTotal As Long

Public Function ObterDataParcPag(dtDataParcPag As Date, lNumIntTitulo As Long, iParcela As Integer) As Long
'para obter a data de vencimento da Parcela a Pagar

Dim lErro As Long

On Error GoTo Erro_ObterDataParcPag

    If lNumIntTitulo <> lNumIntUltimoTitulo Then
        
        Set colParcPag = New ColParcelaPagar
        
        If lNumIntTitulo <> 0 Then
        
            lErro = CF("ParcelasPag_Le_DataValor", lNumIntTitulo, colParcPag, lComando)
            If lErro <> SUCESSO Then Error 58229
        
        End If
        
        lNumIntUltimoTitulo = lNumIntTitulo
        
    End If
    
    If iParcela <= colParcPag.Count Then
        dtDataParcPag = colParcPag(iParcela).dtDataVencimento
    Else
        dtDataParcPag = DATA_NULA
    End If
    
    ObterDataParcPag = SUCESSO
    
    Exit Function
    
Erro_ObterDataParcPag:

    ObterDataParcPag = Err
    
    Select Case Err

        Case 58229 'Tratado na Rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153371)

    End Select

    Exit Function

End Function

Public Function ObterValorParcPag(dValor As Double, lNumIntTitulo As Long, iParcela As Integer) As Long
'para obter o valor da parcela a Pagar

Dim lErro As Long

On Error GoTo Erro_ObterValorParcPag

    If lNumIntTitulo <> lNumIntUltimoTitulo Then
        
        Set colParcPag = New ColParcelaPagar
        
        If lNumIntTitulo <> 0 Then
        
            lErro = CF("ParcelasPag_Le_DataValor", lNumIntTitulo, colParcPag, lComando)
            If lErro <> SUCESSO Then Error 58230
        
        End If
        
        lNumIntUltimoTitulo = lNumIntTitulo
        
    End If
    
    If iParcela <= colParcPag.Count Then
        dValor = colParcPag(iParcela).dValor
    Else
        dValor = 0
    End If
    
    ObterValorParcPag = SUCESSO
    
    Exit Function
    
Erro_ObterValorParcPag:

    ObterValorParcPag = Err
    
    Select Case Err

        Case 58230 'Tratado na Rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153372)

    End Select

    Exit Function

End Function

Private Sub Class_Terminate()

    If lComando <> 0 Then Call Comando_Fechar(lComando)
    If lComandoFornAnt <> 0 Then Call Comando_Fechar(lComandoFornAnt)
    If lComandoFornCred <> 0 Then Call Comando_Fechar(lComandoFornCred)
    If lComandoFornAntTotal <> 0 Then Call Comando_Fechar(lComandoFornAntTotal)
    If lComandoFornCredTotal <> 0 Then Call Comando_Fechar(lComandoFornCredTotal)

End Sub

Public Function Fornecedor_Obter_SaldoAdiantamentos(dSaldoAdiantamentos As Double, lFornecedor As Long) As Long
'Retorna o Saldo de Adiantamentos do Fornecedor

Dim lErro As Long
Dim dPagamentosAntecipados As Double

On Error GoTo Erro_Fornecedor_Obter_SaldoAdiantamentos

    If lFornecedor <> lFornSaldoAntUlt Then

        'Le no BD os Pagamentos antecipados do Fornecedor
        lErro = CF("Fornecedor_Le_SaldoPagAntec", lFornecedor, dPagamentosAntecipados, lComandoFornAnt)
        If lErro <> SUCESSO And lErro <> 58604 Then Error 58605
        
        'Preenche na Variavel Global
        lFornSaldoAntUlt = lFornecedor
        dSaldoAdiantamentosUlt = dPagamentosAntecipados

    End If

    dSaldoAdiantamentos = dSaldoAdiantamentosUlt
    
    Fornecedor_Obter_SaldoAdiantamentos = SUCESSO
    
    Exit Function

Erro_Fornecedor_Obter_SaldoAdiantamentos:

    Fornecedor_Obter_SaldoAdiantamentos = Err
    
    Select Case Err
        
        Case 58605 'Tratado na Rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153373)

    End Select
            
    Exit Function
    
End Function

Public Function Fornecedor_Obter_SaldoCreditos(dSaldoCreditos As Double, lFornecedor As Long) As Long
'Retorna o Saldo de Creditos do Fornecedor

Dim lErro As Long
Dim dSaldoCredito As Double

On Error GoTo Erro_Fornecedor_Obter_SaldoCreditos

    If lFornecedor <> lFornSaldoCredUlt Then

        'Le o Saldo de Crédito para o Fornecedor passado
        lErro = CF("Fornecedor_Le_SaldoCreditos", lFornecedor, dSaldoCredito, lComandoFornCred)
        If lErro <> SUCESSO And lErro <> 58611 Then Error 58612
            
        'Coloca nas Variaveis Globais
        lFornSaldoCredUlt = lFornecedor
        dSaldoCreditosUlt = dSaldoCredito

    End If

    dSaldoCreditos = dSaldoCreditosUlt

    Fornecedor_Obter_SaldoCreditos = SUCESSO
    
    Exit Function
    
Erro_Fornecedor_Obter_SaldoCreditos:
    
    Fornecedor_Obter_SaldoCreditos = Err
    
    Select Case Err
        
        Case 58612 'Tratado na Rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153374)

    End Select
            
    Exit Function
    
End Function

Public Function Fornecedor_Obter_SaldoAdiantamentosTotal(lFornInicial As Long, lFornFinal As Long, dTotalSaldoAdiantamentos As Double) As Long
'Retorna o Total dos Adiantamentos dos Fornecedores (Todos os Fornecedores dentro da Seleção)

Dim lErro As Long

On Error GoTo Erro_Fornecedor_Obter_SaldoAdiantamentosTotal

    'Le no BD o Total dos Adiantamentos dos Fornecedores (Todos os Fornecedores dentro da Seleção)
    lErro = CF("Fornecedor_Le_SaldoPagAntecTotal", lFornInicial, lFornFinal, dTotalSaldoAdiantamentos, lComandoFornAntTotal)
    If lErro <> SUCESSO And lErro <> 58604 Then Error 58620
    
    Fornecedor_Obter_SaldoAdiantamentosTotal = SUCESSO
    
    Exit Function

Erro_Fornecedor_Obter_SaldoAdiantamentosTotal:

    Fornecedor_Obter_SaldoAdiantamentosTotal = Err
    
    Select Case Err
        
        Case 58620 'Tratado na Rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153375)

    End Select
            
    Exit Function
    
End Function

Public Function Fornecedor_Obter_SaldoCreditosTotal(lFornInicial As Long, lFornFinal As Long, dTotalSaldoCredito As Double) As Long
'Retorna o Total dos Saldo de Creditos todos os Fornecedor

Dim lErro As Long

On Error GoTo Erro_Fornecedor_Obter_SaldoCreditosTotal

    'Le a Soma dos Saldos de Crédito para todos os Fornecedores passado
    lErro = CF("Fornecedor_Le_SaldoCreditosTotal", lFornInicial, lFornFinal, dTotalSaldoCredito, lComandoFornCredTotal)
    If lErro <> SUCESSO And lErro <> 58611 Then Error 58621

    Fornecedor_Obter_SaldoCreditosTotal = SUCESSO
    
    Exit Function
    
Erro_Fornecedor_Obter_SaldoCreditosTotal:
    
    Fornecedor_Obter_SaldoCreditosTotal = Err
    
    Select Case Err
        
        Case 58621 'Tratado na Rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153376)

    End Select
            
    Exit Function
    
End Function

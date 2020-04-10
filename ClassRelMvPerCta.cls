VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelMvPerCta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Function MvPerConta_Obtem_Total(dTotal As Double, iPeriodoPedido As Integer, iExercicioPedido As Integer, iFilialEmpresa As Integer, sConta As String) As Long
'Obtem o Total dos 12 periodos anteriores ao periodo pedido.

Dim lErro As Long
Dim lComando As Long
Dim dSoma As Double
Dim iExercAtual As Integer

On Error GoTo Erro_MvPerConta_Obtem_Total

    'ABre o coamndo
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 86077

    'Calcula o periodo anterior
    iExercAtual = iExercicioPedido - 1
    
    'Inicializa a média com o valor zero
    dTotal = 0
    
    'Lê o MvPercta do exercicio pedido
    lErro = Comando_Executar(lComando, "SELECT (Cre01-Deb01)+(Cre02-Deb02)+(Cre03+Deb03)+(Cre04-Deb04)+(Cre05-Deb05)+(Cre06-Deb06)+(Cre07-Deb07)+(Cre08-Deb08)+(Cre09-Deb09)+(Cre10-Deb10)+(Cre11-Deb11)+(Cre12-Deb12) FROM MvPerCta WHERE FilialEmpresa = ? AND Exercicio = ? AND Conta = ?", dSoma, iFilialEmpresa, iExercAtual, sConta)
    If lErro <> AD_SQL_SUCESSO Then Error 86078

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 86079
    If lErro = AD_SQL_SUCESSO Then dTotal = dSoma
                    
    'Fecha o comando
    Call Comando_Fechar(lComando)

    MvPerConta_Obtem_Total = SUCESSO
    
    Exit Function

Erro_MvPerConta_Obtem_Total:

    MvPerConta_Obtem_Total = Err
    
    Select Case Err
    
        Case 86077
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 86078, 86079
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVPERCTA", Err)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153279)
    
    End Select
            
    Call Comando_Fechar(lComando)
            
    Exit Function
    
End Function


Public Function MvPerConta_Obtem_Media(dMedia As Double, iPeriodoPedido As Integer, iExercicioPedido As Integer, iFilialEmpresa As Integer, sConta As String) As Long
'Obtem a média dos 12 periodos anteriores ao periodo pedido.

Dim lErro As Long
Dim lComando As Long
Dim iDifExercicio As Integer, iDifPer As Integer
Dim iNumPeriodos As Integer
Dim iContPer As Integer
Dim adSomaTotal(1 To 12) As Double
Dim dSoma As Double
Dim iIndice As Integer
Dim objLancamento As New ClassLancamento_Detalhe
Dim iExercAtual As Integer

On Error GoTo Erro_MvPerConta_Obtem_Media

    'ABre o coamndo
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 62866

    objLancamento.sConta = sConta
    objLancamento.iFilialEmpresa = iFilialEmpresa
    
    'Lê o exercício e o periodo inicial da conta para a filial passada
    lErro = LancamentoInicial_Le(objLancamento)
    If lErro <> SUCESSO And lErro <> 62865 Then Error 62867

    'Inicializa a média com o valor zero
    dMedia = 0
    
    'Verifica o intervalo entre o exerc. do primeiro lancto e de exercicio pedido
    iDifExercicio = iExercicioPedido - objLancamento.iExercicio
    
    If iDifExercicio >= 0 And lErro = SUCESSO Then
        'Se o exercicio pedido é o mesmo que o inicial
        If iDifExercicio = 0 Then
            'calcula o número de periodos para calculo da média
            iNumPeriodos = iPeriodoPedido - objLancamento.iPeriodoLan
        'Se o exercício inicial é anterior ao o pedido
        ElseIf iDifExercicio = 1 Then
            'calcula o númeo de periodos a tratar contanto os 2 exercicios
            iNumPeriodos = (iPeriodoPedido + 13) - objLancamento.iPeriodoLan
        'Se o exerc. inic. é muito anterior ao pedido
        Else
            'O número de periodos para calcular a média é 12
            iNumPeriodos = 12
        End If
        'Se algum cálculo ultrapassou 12, atribui o 12
        If iNumPeriodos > 12 Then iNumPeriodos = 12
              
        'Se há periodos a calcular
        If iNumPeriodos > 0 Then
            'Lê o MvPercta do exercicio pedido
            lErro = Comando_Executar(lComando, "SELECT (Cre01-Deb01), (Cre02-Deb02), (Cre03-Deb03), (Cre04-Deb04), (Cre05-Deb05), (Cre06-Deb06), (Cre07-Deb07), (Cre08-Deb08),  (Cre09-Deb09), (Cre10-Deb10), (Cre11-Deb11), (Cre12-Deb12) FROM MvPerCta WHERE FilialEmpresa = ? AND Exercicio = ? AND Conta = ?", adSomaTotal(1), adSomaTotal(2), adSomaTotal(3), adSomaTotal(4), adSomaTotal(5), adSomaTotal(6), adSomaTotal(7), adSomaTotal(8), adSomaTotal(9), adSomaTotal(10), adSomaTotal(11), adSomaTotal(12), iFilialEmpresa, iExercicioPedido, sConta)
            If lErro <> AD_SQL_SUCESSO Then Error 62868
        
            lErro = Comando_BuscarPrimeiro(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 62869
            If lErro <> AD_SQL_SUCESSO Then Error 62870
            
            'Acumula os valores dos periodos a contar no exercicio pedido
            For iIndice = (iPeriodoPedido - 1) To LBound(adSomaTotal) Step -1
                dSoma = dSoma + adSomaTotal(iIndice)
                iContPer = iContPer + 1
                If iContPer = iNumPeriodos Then Exit For
            Next
            
            'Se o calculo envolver periodos do exercício anterior
            If iContPer < iNumPeriodos Then
                'Zera o vertor de soma
                For iIndice = LBound(adSomaTotal) To UBound(adSomaTotal)
                    adSomaTotal(iIndice) = 0
                Next
                'Calcula o periodo anterior
                iExercAtual = iExercicioPedido - 1
                
                'Lê o MvPerCta do Exercicio anterior
                lErro = Comando_Executar(lComando, "SELECT (Cre01-Deb01), (Cre02-Deb02), (Cre03+Deb03), (Cre04-Deb04), (Cre05-Deb05), (Cre06-Deb06), (Cre07-Deb07), (Cre08-Deb08),  (Cre09-Deb09), (Cre10-Deb10), (Cre11-Deb11), (Cre12-Deb12) FROM MvPerCta WHERE FilialEmpresa = ? AND Exercicio = ? AND Conta = ?", adSomaTotal(1), adSomaTotal(2), adSomaTotal(3), adSomaTotal(4), adSomaTotal(5), adSomaTotal(6), adSomaTotal(7), adSomaTotal(8), adSomaTotal(9), adSomaTotal(10), adSomaTotal(11), adSomaTotal(12), iFilialEmpresa, iExercAtual, sConta)
                If lErro <> AD_SQL_SUCESSO Then Error 62871
                
                lErro = Comando_BuscarPrimeiro(lComando)
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 62872
                If lErro <> AD_SQL_SUCESSO Then Error 62873
                
                'Inclui no acumulador os periodos do Exercicio anterior ]
                'que fazem parte do cálculo da média
                For iIndice = UBound(adSomaTotal) To LBound(adSomaTotal) Step -1
                    dSoma = dSoma + adSomaTotal(iIndice)
                    iContPer = iContPer + 1
                    If iContPer = iNumPeriodos Then Exit For
                Next
        
            End If
            'Calcula a média
            dMedia = dSoma / iNumPeriodos
        
        End If
    
    End If
    
    'Fecha o comando
    Call Comando_Fechar(lComando)

    MvPerConta_Obtem_Media = SUCESSO
    
    Exit Function

Erro_MvPerConta_Obtem_Media:

    MvPerConta_Obtem_Media = Err
    
    Select Case Err
    
        Case 62866
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 62867
            
        Case 62868, 62869, 62870, 62871, 62872, 62873
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVPERCTA", Err)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153280)
    
    End Select
            
    Call Comando_Fechar(lComando)
            
    Exit Function
    
End Function

Function LancamentoInicial_Le(objLancamento As ClassLancamento_Detalhe) As Long
'Lê o Exercício e o Periodo  do primeiro lançamento feito para a conta passada na filial passada

Dim lErro As Long
Dim lComando As Long
Dim tLanc As typeLancamento

On Error GoTo Erro_LancamentoInicial_Le
    
    'Abre o comando
    lComando = Comando_Abrir
    If lComando = 0 Then gError 62862
    
    With tLanc
        'Lê o primeiro lancamento
        lErro = Comando_Executar(lComando, "SELECT Exercicio,PeriodoLan FROM Lancamentos WHERE Conta = ? AND FilialEmpresa = ? ORDER BY Data , PeriodoLan", .iExercicio, .iPeriodoLan, objLancamento.sConta, objLancamento.iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 62863
        
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 62864
        If lErro <> AD_SQL_SUCESSO Then gError 62865
        
        'Guarda o exercício e o período
        objLancamento.iExercicio = .iExercicio
        objLancamento.iPeriodoLan = .iPeriodoLan
    
    End With

    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    LancamentoInicial_Le = SUCESSO
    
    Exit Function
    
Erro_LancamentoInicial_Le:

    LancamentoInicial_Le = gErr
    
    Select Case gErr
    
        Case 62862
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 62863, 62864
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LANCAMENTOS2", gErr, objLancamento.sConta)
        
        Case 62865
           
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153281)
            
    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function


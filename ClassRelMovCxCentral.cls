VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelMovCxCentral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim iFilialEmpresaAnt As Integer
Dim iCaixaAnt As Integer
Dim dtDataAnt As Date
Dim dSaldoAnt As Double
Dim dDebitoAnt As Double
Dim dCreditoAnt As Double
Dim iTipoMeioPagtoAnt As Integer

'Function Obtem_MvCaixaCentral_SldInicial(dtDataIn As Date, dSaldoInicial As Double, iCodCaixa As Integer, iFilialEmpresa As Integer) As Long
''Função que Obtem po Saldo a partir da data inicial do Relatorio
'
'Dim lErro As Long
'Dim lTransacao As Long
'Dim sSQLDeb As String
'Dim sSQLCred As String
'Dim alComando(1) As Long
'Dim dtdataAux As Date
'Dim iMesAux As Integer
'Dim iAnoAux As Integer
'Dim iIndice As Integer
'Dim Deb(1 To 12) As Double
'Dim Cred(1 To 12) As Double
'Dim dDebMem As Double
'Dim dCredMem As Double
'Dim iCodCaixaAux As Integer
'Dim iFilialAux As Integer
'
'On Error GoTo Erro_Obtem_MvCaixaCentral_SldInicial
'
'    'Abre o comando
'    For iIndice = LBound(alComando) To UBound(alComando)
'
'        alComando(iIndice) = Comando_Abrir
'        If alComando(iIndice) = 0 Then gError 113156
'
'    Next
'
'    'Verifica o Ano para Adquirir o saldo
'    iAnoAux = Year(dtDataIn)
'
'    'Selecionar o saldo aglutinado de Movimentos de Caixa para o ano q passou
'    lErro = Comando_Executar(alComando(0), "SELECT  SaldoInicial FROM CCMov WHERE FilialEmpresa = ? AND Ano = ? AND CodCaixa = ? ", dSaldoInicial, iFilialEmpresa, iAnoAux, iCodCaixa)
'    If lErro <> AD_SQL_SUCESSO Then gError 113163
'
'    lErro = Comando_BuscarPrimeiro(alComando(0))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 113158
'
'    iIndice = 1
'
'    'Cria Uma String Com todos os Meses do ano
'    Do While iIndice <= 12
'
'        sSQLDeb = sSQLDeb & " Deb0" & iIndice & " +"
'        sSQLCred = sSQLCred & " Cred0" & iIndice & " +"
'        iIndice = iIndice + 1
'        If iIndice > 9 Then
'            sSQLDeb = sSQLDeb & " Deb10 + "
'            sSQLCred = sSQLCred & " Cred10 + "
'            sSQLDeb = sSQLDeb & " Deb11 + "
'            sSQLCred = sSQLCred & " Cred11  + "
'            sSQLDeb = sSQLDeb & " Deb12 +"
'            sSQLCred = sSQLCred & " Cred12 +"
'            Exit Do
'        End If
'    Loop
'
'    sSQLDeb = Left(sSQLDeb, Len(sSQLDeb) - 1)
'    sSQLCred = Left(sSQLCred, Len(sSQLCred) - 1)
'
'    'Selecionar o saldo aglutinado de Movimentos de Caixa para o ano q passou
'    lErro = Comando_Executar(alComando(1), "SELECT CodCaixa , FilialEmpresa , SUM(" & sSQLDeb & "), SUM (" & sSQLCred & ") FROM CCMov WHERE FilialEmpresa = ? AND Ano = ? AND CodCaixa = ? GROUP BY (CodCaixa) , (FilialEmpresa), ( SaldoInicial)", iCodCaixaAux, iFilialAux, dDebMem, dCredMem, iFilialEmpresa, iAnoAux, iCodCaixa)
'    If lErro <> AD_SQL_SUCESSO Then gError 113164
'
'    lErro = Comando_BuscarPrimeiro(alComando(1))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 113165
'
'    'Atualiza o Saldo de Movimentos até o mês anterior
'    dSaldoInicial = dSaldoInicial + dDebMem + dCredMem
'
'    'fecha os comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'
'        Call Comando_Fechar(alComando(iIndice))
'
'    Next
'
'    Obtem_MvCaixaCentral_SldInicial = SUCESSO
'
'    Exit Function
'
'Erro_Obtem_MvCaixaCentral_SldInicial:
'
'    Obtem_MvCaixaCentral_SldInicial = gErr
'
'    Select Case gErr
'
'        Case 113156
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 113158, 113163, 113164, 113165
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_SELECAO_CCMOV", gErr, iCodCaixa, iAnoAux, iFilialEmpresa)
'
'       Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153275)
'
'    End Select
'
'    'fecha os comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'
'        Call Comando_Fechar(alComando(iIndice))
'
'    Next
'
'    Exit Function
'
'End Function

Function Obtem_MvCaixaCentral_SldInicial(dtData As Date, dSaldo As Double, iCaixa As Integer, iFilialEmpresa As Integer) As Long
'Função que Obtem po Saldo a partir da data inicial do Relatorio

Dim lErro As Long
Dim iMes As Integer
Dim iAno As Integer
Dim dSaldoDiario As Double

On Error GoTo Erro_Obtem_MvCaixaCentral_SldInicial

    If iFilialEmpresa <> iFilialEmpresaAnt Or dtData <> dtDataAnt Or iCaixa <> iCaixaAnt Or 0 <> iTipoMeioPagtoAnt Then
            
        'Busca dSaldo na tabela MvPerCli
        lErro = CF("CCMov_Le_SldAnterior", dSaldo, iCaixa, dtData, iFilialEmpresa)
        If lErro <> SUCESSO Then gError 105645

        iMes = Month(dtData)
        iAno = Year(dtData)
        
        'Se nao for o ultimo dia do mes, consulta a tabela MvDiaCli
        'If Dias_Mes(iMes, iAno) <> Day(dtData) Then
        'Essa lógica estava errada e presente em vários .cls, mas só irei alterar esse
        'Se eu quero o saldo inicial de 31/12 não posso pegar o valor completo de Dezembro pois ele inclui a movimentação
        'do dia 31 que não deveria ser computada pois é o saldo inicial dessa data
        'Antes também nem funcionava para o último dia porque não setava iMes e iAno e aí duplicava o valor
        If 1 <> Day(dtData) Then
        
            lErro = CF("CCMovDia_Le_SldAnterior", dSaldoDiario, iCaixa, dtData, iFilialEmpresa)
            If lErro <> SUCESSO Then gError 105646
            
            dSaldo = dSaldo + dSaldoDiario
        
        End If
        
        iFilialEmpresaAnt = iFilialEmpresa
        iCaixaAnt = iCaixa
        dtDataAnt = dtData
        dSaldoAnt = dSaldo
        iTipoMeioPagtoAnt = 0
        
    End If
    
    dSaldo = dSaldoAnt
        
    Exit Function
    
    Obtem_MvCaixaCentral_SldInicial = SUCESSO
    
Erro_Obtem_MvCaixaCentral_SldInicial:

    Obtem_MvCaixaCentral_SldInicial = gErr
    
    Select Case gErr
        
        Case 105645, 105646
            'Erros tratados nas rotinas chamadas
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153276)

    End Select

    Exit Function

End Function

Function Obtem_MvCaixaCentral_SldInicial_TMP(dtData As Date, dSaldo As Double, iCaixa As Integer, iFilialEmpresa As Integer, iTipoMeioPagto As Integer) As Long
'Função que Obtem po Saldo a partir da data inicial do Relatorio

Dim lErro As Long
Dim iMes As Integer
Dim iAno As Integer
Dim dSaldoDiario As Double
Dim dDebito As Double, dCredito As Double

On Error GoTo Erro_Obtem_MvCaixaCentral_SldInicial_TMP

    If iFilialEmpresa <> iFilialEmpresaAnt Or dtData <> dtDataAnt Or iCaixa <> iCaixaAnt Or iTipoMeioPagto <> iTipoMeioPagtoAnt Then
                
        'Busca dSaldo na tabela MvPerCli
        lErro = CF("CCMov_Le_SldAnterior", dSaldo, iCaixa, dtData, iFilialEmpresa, iTipoMeioPagto)
        If lErro <> SUCESSO Then gError 105645

        iMes = Month(dtData)
        iAno = Year(dtData)
        
        'Se nao for o ultimo dia do mes, consulta a tabela MvDiaCli
        'If Dias_Mes(iMes, iAno) <> Day(dtData) Then
        'Essa lógica estava errada e presente em vários .cls, mas só irei alterar esse
        'Se eu quero o saldo inicial de 31/12 não posso pegar o valor completo de Dezembro pois ele inclui a movimentação
        'do dia 31 que não deveria ser computada pois é o saldo inicial dessa data
        'Antes também nem funcionava para o último dia porque não setava iMes e iAno e aí duplicava o valor
        If 1 <> Day(dtData) Then
        
            lErro = CF("CCMovDia_Le_SldAnterior", dSaldoDiario, iCaixa, dtData, iFilialEmpresa, iTipoMeioPagto)
            If lErro <> SUCESSO Then gError 105646
            
            dSaldo = dSaldo + dSaldoDiario
        
        End If
        
        lErro = CF("CCMovDia_Le_DebCred", dDebito, dCredito, iCaixa, dtData, iFilialEmpresa, iTipoMeioPagto)
        If lErro <> SUCESSO Then gError 105646
        
        iFilialEmpresaAnt = iFilialEmpresa
        iCaixaAnt = iCaixa
        dtDataAnt = dtData
        dSaldoAnt = dSaldo
        dDebitoAnt = dDebito
        dCreditoAnt = dCredito
        iTipoMeioPagtoAnt = iTipoMeioPagto
        
    End If
    
    dSaldo = dSaldoAnt
        
    Exit Function
    
    Obtem_MvCaixaCentral_SldInicial_TMP = SUCESSO
    
Erro_Obtem_MvCaixaCentral_SldInicial_TMP:

    Obtem_MvCaixaCentral_SldInicial_TMP = gErr
    
    Select Case gErr
        
        Case 105645, 105646
            'Erros tratados nas rotinas chamadas
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153276)

    End Select

    Exit Function

End Function

Function Obtem_MvCaixaCentral_SldInicial_TMP_Cust(dtData As Date, dSaldo As Double, iCaixa As Integer, iFilialEmpresa As Integer, iTipoMeioPagto As Integer) As Long
'Função que Obtem po Saldo a partir da data inicial do Relatorio

Dim lErro As Long
Dim iMes As Integer
Dim iAno As Integer
Dim dSaldoDiario As Double
Dim dDebito As Double, dCredito As Double

On Error GoTo Erro_Obtem_MvCaixaCentral_SldInicial_TMP_Cust

    If iFilialEmpresa <> iFilialEmpresaAnt Or dtData <> dtDataAnt Or iCaixa <> iCaixaAnt Or iTipoMeioPagto <> iTipoMeioPagtoAnt Then
                
        'Busca dSaldo na tabela MvPerCli
        lErro = CF("CCMov_Le_SldAnterior", dSaldo, iCaixa, dtData, iFilialEmpresa, iTipoMeioPagto, "CCMovCust")
        If lErro <> SUCESSO Then gError 105645
        
        iMes = Month(dtData)
        iAno = Year(dtData)

        'Se nao for o ultimo dia do mes, consulta a tabela MvDiaCli
        'If Dias_Mes(iMes, iAno) <> Day(dtData) Then
        'Essa lógica estava errada e presente em vários .cls, mas só irei alterar esse
        'Se eu quero o saldo inicial de 31/12 não posso pegar o valor completo de Dezembro pois ele inclui a movimentação
        'do dia 31 que não deveria ser computada pois é o saldo inicial dessa data
        'Antes também nem funcionava para o último dia porque não setava iMes e iAno e aí duplicava o valor
        If 1 <> Day(dtData) Then
        
            lErro = CF("CCMovDia_Le_SldAnterior", dSaldoDiario, iCaixa, dtData, iFilialEmpresa, iTipoMeioPagto, "CCMovDiaCust")
            If lErro <> SUCESSO Then gError 105646
            
            dSaldo = dSaldo + dSaldoDiario
        
        End If
        
        lErro = CF("CCMovDia_Le_DebCred", dDebito, dCredito, iCaixa, dtData, iFilialEmpresa, iTipoMeioPagto, "CCMovDiaCust")
        If lErro <> SUCESSO Then gError 105646
        
        iFilialEmpresaAnt = iFilialEmpresa
        iCaixaAnt = iCaixa
        dtDataAnt = dtData
        dSaldoAnt = dSaldo
        dDebitoAnt = dDebito
        dCreditoAnt = dCredito
        iTipoMeioPagtoAnt = iTipoMeioPagto
        
    End If
    
    dSaldo = dSaldoAnt
        
    Exit Function
    
    Obtem_MvCaixaCentral_SldInicial_TMP_Cust = SUCESSO
    
Erro_Obtem_MvCaixaCentral_SldInicial_TMP_Cust:

    Obtem_MvCaixaCentral_SldInicial_TMP_Cust = gErr
    
    Select Case gErr
        
        Case 105645, 105646
            'Erros tratados nas rotinas chamadas
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153276)

    End Select

    Exit Function

End Function

Function Obtem_MvCaixaCentral_Cred_TMP_Cust(dtData As Date, dCredito As Double, iCaixa As Integer, iFilialEmpresa As Integer, iTipoMeioPagto As Integer) As Long

Dim lErro As Long
Dim dSaldo As Double

On Error GoTo Erro_Obtem_MvCaixaCentral_Cred_TMP_Cust

    lErro = Obtem_MvCaixaCentral_SldInicial_TMP_Cust(dtData, dSaldo, iCaixa, iFilialEmpresa, iTipoMeioPagto)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    dCredito = dCreditoAnt
        
    Exit Function
    
    Obtem_MvCaixaCentral_Cred_TMP_Cust = SUCESSO
    
Erro_Obtem_MvCaixaCentral_Cred_TMP_Cust:

    Obtem_MvCaixaCentral_Cred_TMP_Cust = gErr
    
    Select Case gErr
        
        Case ERRO_SEM_MENSAGEM
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153276)

    End Select

    Exit Function

End Function

Function Obtem_MvCaixaCentral_Deb_TMP_Cust(dtData As Date, dDebito As Double, iCaixa As Integer, iFilialEmpresa As Integer, iTipoMeioPagto As Integer) As Long

Dim lErro As Long
Dim dSaldo As Double

On Error GoTo Erro_Obtem_MvCaixaCentral_Deb_TMP_Cust

    lErro = Obtem_MvCaixaCentral_SldInicial_TMP_Cust(dtData, dSaldo, iCaixa, iFilialEmpresa, iTipoMeioPagto)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    dDebito = dDebitoAnt
        
    Exit Function
    
    Obtem_MvCaixaCentral_Deb_TMP_Cust = SUCESSO
    
Erro_Obtem_MvCaixaCentral_Deb_TMP_Cust:

    Obtem_MvCaixaCentral_Deb_TMP_Cust = gErr
    
    Select Case gErr
        
        Case ERRO_SEM_MENSAGEM
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153276)

    End Select

    Exit Function

End Function

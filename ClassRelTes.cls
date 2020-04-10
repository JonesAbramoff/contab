VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelTes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'criada para atender às chamadas de EXT_FUNC nos relatorios atraves de pesquisa, principalmente, nas tabelas CCIMov, CCIMovDia
'
'IMPORTANTE: NAO ESQUECA DE TRATAR DATA_NULA COMO POSSIVEL DATA INICIAL

'Identificação
Private dtUltDataInicial As Date
Private dtUltDataFinal As Date
Private iUltCodCCISldIni As Integer 'última conta corrente pesquisada para saldo inicial. Obs.: só preencher se conseguir obter os dados solicitados, ou seja, se os "valores obtidos" abaixo forem válidos.
Private iUltCodCCISldAplic As Integer 'última conta corrente pesquisada para saldo de aplicações. Obs.: só preencher se conseguir obter os dados solicitados, ou seja, se os "valores obtidos" abaixo forem válidos.

'Valores armazenados para a última CCI pesquisada
Private dUltSldIni As Double 'saldo inicial na data em questão
Private dUltSldAplic As Double 'saldo aplicado na data em questão

Private dUltSldIni_Total As Double
Private dtUltDataInicial_Total As Date
Private dUltSldAplic_Total As Double
Private dtUltDataFinal_Total As Date

'Para Valores Entrada e Saída
Private dtUltDataInicialEntSai As Date
Private dtUltDataFinalEntSai As Date
Private dUltValorEntrada As Double
Private dUltValorSaida As Double
Private iUltCodigoConta As Integer

'################################
'Inserido por Wagner
Private iUltCodCCISldIniConc As Integer
Private dtUltDataInicialConc As Date
Private dUltSldIniConc As Double
'################################

Private lUltNumMovto As Long
Private sUltFornCliTit As String
Private sUltHistTit As String

Private alComando(1 To 4) As Long
Private alComando1(1 To 4) As Long
Private alComando2(1 To 4) As Long
Private alComando3(1 To 4) As Long
Private alComando4(1 To 4) As Long
Private alComando5(1 To 4) As Long
Private alComando6(1 To 4) As Long
Private alComando7(1 To 7) As Long

Private alComando8(1 To 4) As Long

Public Function ObterSldIni(iCodCCI As Integer, dtData As Date, dSldIni As Double) As Long
'Para obter o saldo inicial de uma conta corrente numa data

Dim lErro As Long

On Error GoTo Erro_ObterSldIni
    
    If iCodCCI <> iUltCodCCISldIni Or dtData <> dtUltDataInicial Then
        
        lErro = ObterRelTes(iCodCCI, dtData)
        If lErro <> SUCESSO Then Error 60735
    
    End If
    
    dSldIni = dUltSldIni
    
    ObterSldIni = SUCESSO
    
    Exit Function
    
Erro_ObterSldIni:

    ObterSldIni = Err
    
    Select Case Err

        Case 60735
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153358)

    End Select

    Exit Function
    
End Function

Public Function ObterSldIni_Total(dtData As Date, dSldIni As Double) As Long
'Para obter o saldo inicial de uma conta corrente numa data

Dim lErro As Long

On Error GoTo Erro_ObterSldIni_Total
    
    If dtData <> dtUltDataInicial_Total Then
        
        lErro = CF("CCI_ObterRelTes_Total", dtData, dSldIni, alComando)
        If lErro <> SUCESSO Then gError 69073
    
        dtUltDataInicial_Total = dtData
    
        dUltSldIni_Total = dSldIni

    End If
    
    dSldIni = dUltSldIni_Total
    
    ObterSldIni_Total = SUCESSO
    
    Exit Function
    
Erro_ObterSldIni_Total:

    ObterSldIni_Total = gErr
    
    Select Case gErr

        Case 69073
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153359)

    End Select

    Exit Function
    
End Function

Private Function ObterRelTes(iCodCCI As Integer, dtData As Date) As Long
'Para obter os valores a partir de ccimov/ccimovdia e torna-los disponiveis para as outras funcoes

Dim lErro As Long
Dim dSldIni As Double

On Error GoTo Erro_ObterRelTes

    lErro = CF("CCI_ObterRelTes", iCodCCI, dtData, dSldIni, alComando1)
    If lErro <> SUCESSO Then Error 60736
    
    iUltCodCCISldIni = iCodCCI
    dtUltDataInicial = dtData
    
    dUltSldIni = dSldIni
        
    ObterRelTes = SUCESSO
    
    Exit Function
    
Erro_ObterRelTes:

    ObterRelTes = Err
    
    Select Case Err

        Case 60736
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153360)

    End Select

    Exit Function
    
End Function

Public Function ObterSldAplic(iCodCCI As Integer, dtData As Date, dSldAplic As Double) As Long
'Para obter o saldo das aplicações associados a uma conta corrente numa data

Dim lErro As Long

On Error GoTo Erro_ObterSldAplic
    
    If iCodCCI <> iUltCodCCISldAplic Or dtData <> dtUltDataFinal Then
        
        lErro = ObterRelTesAplic(iCodCCI, dtData)
        If lErro <> SUCESSO Then Error 60747
    
    End If
    
    dSldAplic = dUltSldAplic
    
    ObterSldAplic = SUCESSO
    
    Exit Function
    
Erro_ObterSldAplic:

    ObterSldAplic = Err
    
    Select Case Err

        Case 60747
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153361)

    End Select

    Exit Function
    
End Function

Public Function ObterSldAplic_Total(dtData As Date, dSldAplic As Double) As Long
'Para obter o saldo das aplicações até uma data

Dim lErro As Long

On Error GoTo Erro_ObterSldAplic_Total
    
    If dtData <> dtUltDataFinal_Total Then
        
        lErro = CF("CCI_ObterRelTesAplic_Total", dtData, dSldAplic, alComando2)
        If lErro <> SUCESSO Then gError 69072
        
        dtUltDataFinal_Total = dtData
        dUltSldAplic_Total = dSldAplic
    
    End If
    
    dSldAplic = dUltSldAplic_Total
    
    ObterSldAplic_Total = SUCESSO
    
    Exit Function
    
Erro_ObterSldAplic_Total:

    ObterSldAplic_Total = gErr
    
    Select Case gErr

        Case 69072
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153362)

    End Select

    Exit Function
    
End Function

Private Function ObterRelTesAplic(iCodCCI As Integer, dtData As Date) As Long
'Para obter os valores a partir de ccimov/ccimovdia e torna-los disponiveis para as outras funcoes

Dim lErro As Long
Dim dSldAplic As Double

On Error GoTo Erro_ObterRelTesAplic

    lErro = CF("CCI_ObterRelTesAplic", iCodCCI, dtData, dSldAplic, alComando3)
    If lErro <> SUCESSO Then Error 60748
    
    iUltCodCCISldAplic = iCodCCI
    dtUltDataFinal = dtData
    
    dUltSldAplic = dSldAplic
        
    ObterRelTesAplic = SUCESSO
    
    Exit Function
    
Erro_ObterRelTesAplic:

    ObterRelTesAplic = Err
    
    Select Case Err

        Case 60748
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153363)

    End Select

    Exit Function
    
End Function


Private Sub Class_Terminate()

Dim iIndice As Integer

    For iIndice = LBound(alComando) To UBound(alComando)
         Comando_Fechar (alComando(iIndice))
    Next

    For iIndice = LBound(alComando1) To UBound(alComando1)
         Comando_Fechar (alComando1(iIndice))
    Next

    For iIndice = LBound(alComando2) To UBound(alComando2)
         Comando_Fechar (alComando2(iIndice))
    Next

    For iIndice = LBound(alComando3) To UBound(alComando3)
         Comando_Fechar (alComando3(iIndice))
    Next

    For iIndice = LBound(alComando4) To UBound(alComando4)
         Comando_Fechar (alComando4(iIndice))
    Next

    For iIndice = LBound(alComando5) To UBound(alComando5)
         Comando_Fechar (alComando5(iIndice))
    Next

    For iIndice = LBound(alComando6) To UBound(alComando6)
         Comando_Fechar (alComando6(iIndice))
    Next

    For iIndice = LBound(alComando7) To UBound(alComando7)
         Comando_Fechar (alComando7(iIndice))
    Next

    For iIndice = LBound(alComando8) To UBound(alComando8)
         Comando_Fechar (alComando8(iIndice))
    Next
    
End Sub


Public Function ObterSldEntrada(dtDataInicial As Date, dtDataFinal As Date, iCodConta As Integer, dValorEntrada As Double) As Long
'Para obter o saldo de Entrada

Dim lErro As Long
Dim dValorSaida As Double

On Error GoTo Erro_ObterSldEntrada
    
    If dtDataInicial <> dtUltDataInicialEntSai Or dtDataFinal <> dtUltDataFinalEntSai Or iCodConta <> iUltCodigoConta Then
        
        lErro = CF("CCI_ObterRelTesEntradaSaida", iCodConta, dtDataInicial, dtDataFinal, dValorEntrada, dValorSaida, alComando4)
        If lErro <> SUCESSO Then gError 69083
        
        dtUltDataInicialEntSai = dtDataInicial
        dtUltDataFinalEntSai = dtDataFinal
        dUltValorEntrada = dValorEntrada
        dUltValorSaida = dValorSaida
        iUltCodigoConta = iCodConta
        
    End If
    
    dValorEntrada = dUltValorEntrada
    
    ObterSldEntrada = SUCESSO
    
    Exit Function
    
Erro_ObterSldEntrada:

    ObterSldEntrada = gErr
    
    Select Case gErr

        Case 69083
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153364)

    End Select

    Exit Function
    
End Function

Public Function ObterSldSaida(dtDataInicial As Date, dtDataFinal As Date, iCodConta As Integer, dValorSaida As Double) As Long
'Para obter o Saldo de Saida

Dim lErro As Long
Dim dValorEntrada As Double

On Error GoTo Erro_ObterSldSaida
    
    If dtDataInicial <> dtUltDataInicialEntSai Or dtDataFinal <> dtUltDataFinalEntSai Or iCodConta <> iUltCodigoConta Then
        
        lErro = CF("CCI_ObterRelTesEntradaSaida", iCodConta, dtDataInicial, dtDataFinal, dValorEntrada, dValorSaida, alComando5)
        If lErro <> SUCESSO Then gError 69084
        
        dtUltDataInicialEntSai = dtDataInicial
        dtUltDataFinalEntSai = dtDataFinal
        dUltValorEntrada = dValorEntrada
        dUltValorSaida = dValorSaida
        iUltCodigoConta = iCodConta
    
    End If
    
    dValorSaida = dUltValorSaida
    
    ObterSldSaida = SUCESSO
    
    Exit Function
    
Erro_ObterSldSaida:

    ObterSldSaida = gErr
    
    Select Case gErr

        Case 69084
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153365)

    End Select

    Exit Function
    
End Function

Public Function ObterDataVenctoReal(dtDataVenctoReal As Date, ByVal dtData As Date) As Long
'Retorna data de Vencimento real

Dim lErro As Long

On Error GoTo Erro_ObterDataVenctoReal
    
    If dtData > 0 Then
    
        lErro = CF("DataVencto_Real", dtData, dtDataVenctoReal)
        If lErro <> SUCESSO Then gError 93347
    
    End If
        
    ObterDataVenctoReal = SUCESSO
    
    Exit Function

Erro_ObterDataVenctoReal:

    ObterDataVenctoReal = gErr
    
    Select Case gErr
                
        Case 93347
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153366)
        
    End Select
    
    Exit Function
    
End Function

'Copiada de RotinasCRFAT
Function DataVencto_Real(dtDataVencimento As Date, dtDataVenctoReal As Date) As Long
'Gera a Data de Vencimento Real em cima da Data de Vencimento passada por parâmetro
'os feriados considerados sao da filialempesa corrente
'AVISO: ESSA ROTINA TEM QUE SER CHAMADA DENTRO DE UMA TRANSACAO

Dim vbDiaDaSemana As VbDayOfWeek
Dim dtFeriado As Date
Dim lErro As Long
Dim lComando As Long

On Error GoTo Erro_DataVencto_Real

    'Abre comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 18444
    
    dtDataVenctoReal = dtDataVencimento
    
    vbDiaDaSemana = Weekday(dtDataVenctoReal)
    
    'Se for a data for sábado ou domingo transferimos para a segunda
    If vbDiaDaSemana = vbSunday Then dtDataVenctoReal = dtDataVenctoReal + 1
    If vbDiaDaSemana = vbSaturday Then dtDataVenctoReal = dtDataVenctoReal + 2

    'Pesquisa-se as na Tabela de Feriados as datas maiores ou iguais a Data de Vencimento
    lErro = Comando_Executar(lComando, "SELECT Data FROM Feriados WHERE FilialEmpresa = ? AND Data >= ? ORDER BY Data", dtFeriado, giFilialEmpresa, dtDataVenctoReal)
    If lErro <> AD_SQL_SUCESSO Then Error 18445
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 18446
    
    Do While lErro = AD_SQL_SUCESSO
    
        'Se for maior é porque a data de vencimento não é feriado
        If (dtFeriado > dtDataVenctoReal) Then Exit Do

        'Incrementa-se a data
        dtDataVenctoReal = dtDataVenctoReal + 1
        
        vbDiaDaSemana = Weekday(dtDataVenctoReal)
        
        'Verifica se é sábado ou domingo
        If vbDiaDaSemana = vbSunday Then dtDataVenctoReal = dtDataVenctoReal + 1
        If vbDiaDaSemana = vbSaturday Then dtDataVenctoReal = dtDataVenctoReal + 2
        
        'Loop para adiantar DataFeriado se for menor que DataVenctoReal
        Do While dtFeriado < dtDataVenctoReal And lErro <> AD_SQL_SEM_DADOS
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 18447
        Loop
    Loop
    
    Call Comando_Fechar(lComando)
    
    DataVencto_Real = SUCESSO
    
    Exit Function
    
Erro_DataVencto_Real:

    DataVencto_Real = Err
    
    Select Case Err
    
        Case 18444
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 18445, 18446, 18447
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FERIADOS", Err)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153367)
            
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
        
End Function

'#####################################################
'Inserido por Wagner
Private Function ObterRelTesConciliado(iCodCCI As Integer, dtData As Date) As Long
'Para obter os valores a partir de ccimov/ccimovdia e torna-los disponiveis para as outras funcoes

Dim lErro As Long
Dim dSldIni As Double

On Error GoTo Erro_ObterRelTesConciliado

    lErro = CF("CCI_ObterRelTesConciliado", iCodCCI, dtData, dSldIni, alComando6)
    If lErro <> SUCESSO Then gError 132070
    
    iUltCodCCISldIniConc = iCodCCI
    dtUltDataInicialConc = dtData
    
    dUltSldIniConc = dSldIni
        
    ObterRelTesConciliado = SUCESSO
    
    Exit Function
    
Erro_ObterRelTesConciliado:

    ObterRelTesConciliado = gErr
    
    Select Case gErr

        Case 132070
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153368)

    End Select

    Exit Function
    
End Function

Public Function ObterSldIniConc(iCodCCI As Integer, dtData As Date, dSldIni As Double) As Long
'Para obter o saldo inicial de uma conta corrente numa data

Dim lErro As Long

On Error GoTo Erro_ObterSldIniConc
    
    If iCodCCI <> iUltCodCCISldIniConc Or dtData <> dtUltDataInicialConc Then
        
        lErro = ObterRelTesConciliado(iCodCCI, dtData)
        If lErro <> SUCESSO Then gError 132072
    
    End If
    
    dSldIni = dUltSldIniConc
    
    ObterSldIniConc = SUCESSO
    
    Exit Function
    
Erro_ObterSldIniConc:

    ObterSldIniConc = gErr
    
    Select Case gErr

        Case 132072
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153369)

    End Select

    Exit Function
    
End Function

Public Function ObterHistContabil(sHistorico As String, ByVal lNumMovto As Long) As Long
'Para obter o saldo inicial de uma conta corrente numa data

Dim lErro As Long

On Error GoTo Erro_ObterHistContabil
            
    lErro = CF("CCI_ObterHistContabil", sHistorico, lNumMovto, alComando7)
    If lErro <> SUCESSO Then gError 132182
    
    ObterHistContabil = SUCESSO
    
    Exit Function
    
Erro_ObterHistContabil:

    ObterHistContabil = gErr
    
    Select Case gErr

        Case 132182
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153370)

    End Select

    Exit Function
    
End Function

Public Function Obter_CliForn(ByVal iTipo As Integer, ByVal lNumMovto As Long, sCliForn As String) As Long

Dim lErro As Long

On Error GoTo Erro_Obter_CliForn
            
    If lUltNumMovto <> lNumMovto Then
        lErro = Obter_Dados_Titulo(iTipo, lNumMovto)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    End If
    
    sCliForn = sUltFornCliTit
    
    Obter_CliForn = SUCESSO
    
    Exit Function
    
Erro_Obter_CliForn:

    Obter_CliForn = gErr
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 213853)

    End Select

    Exit Function
    
End Function

Public Function Obter_HistoricoTit(ByVal iTipo As Integer, ByVal lNumMovto As Long, sHistorico As String) As Long

Dim lErro As Long

On Error GoTo Erro_Obter_HistoricoTit
            
    If lUltNumMovto <> lNumMovto Then
        lErro = Obter_Dados_Titulo(iTipo, lNumMovto)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    End If
    
    sHistorico = sUltHistTit

    Obter_HistoricoTit = SUCESSO
    
    Exit Function
    
Erro_Obter_HistoricoTit:

    Obter_HistoricoTit = gErr
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 213854)

    End Select

    Exit Function
    
End Function

Public Function Obter_Dados_Titulo(ByVal iTipo As Integer, ByVal lNumMovto As Long) As Long

Dim lErro As Long, iIndice As Integer
Dim sTipoCliForn As String, lCliForn As Long, sNomeCliForn As String, sHistorico As String
Dim lCliFornAnt As Long, sCliFornAux As String, sHistAux As String
Dim sSQL As String

On Error GoTo Erro_Obter_Dados_Titulo

    For iIndice = LBound(alComando8) To UBound(alComando8)
        If alComando8(iIndice) = 0 Then
            alComando8(iIndice) = Comando_AbrirExt(GL_lConexaoRel)
            If alComando8(iIndice) = 0 Then gError 213855
        End If
    Next
            
    Select Case iTipo
    
        Case 6 'Adiantamento para Fornecedor
        
            sTipoCliForn = "F"
               
            sNomeCliForn = String(STRING_MAXIMO, 0)
               
            lErro = Comando_Executar(alComando8(1), "SELECT P.Fornecedor, F.NomeReduzido FROM PagtosAntecipados AS P, Fornecedores AS F WHERE P.Fornecedor = F.Codigo AND NumMovto = ?", _
            lCliForn, sNomeCliForn, lNumMovto)
            If lErro <> AD_SQL_SUCESSO Then gError 213856
            
            lErro = Comando_BuscarPrimeiro(alComando8(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213857
            
            sHistorico = ""
            
            sCliFornAux = sTipoCliForn & ":" & CStr(lCliForn) & "-" & sNomeCliForn
            sHistAux = sHistorico
                    
        Case 7 'Adiantamento de Cliente
        
            sTipoCliForn = "C"
        
            sNomeCliForn = String(STRING_MAXIMO, 0)
        
            lErro = Comando_Executar(alComando8(2), "SELECT R.Cliente, C.NomeReduzido FROM RecebAntecipados AS R, Clientes AS C WHERE R.Cliente = C.Codigo AND NumMovto = ?", _
            lCliForn, sNomeCliForn, lNumMovto)
            If lErro <> AD_SQL_SUCESSO Then gError 213858
            
            lErro = Comando_BuscarPrimeiro(alComando8(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213859
            
            sHistorico = ""
            
            sCliFornAux = sTipoCliForn & ":" & CStr(lCliForn) & "-" & sNomeCliForn
            sHistAux = sHistorico
            
        Case 8, 24 'Recebimento de cliente - > BaixasRec
        
            sTipoCliForn = "C"
            
            sNomeCliForn = String(STRING_MAXIMO, 0)
            sHistorico = String(STRING_MAXIMO, 0)
        
            sSQL = "SELECT DISTINCT T.Cliente, C.NomeReduzido FROM TitulosRecTodos AS T, ParcelasRecTodas AS P, BaixasParcRec As BP, Clientes AS C, BaixasRec As B"
            sSQL = sSQL & " " & "WHERE T.NumIntDoc = P.NumIntTitulo And P.NumIntDoc = BP.NumIntParcela AND T.Cliente = C.Codigo AND B.NumIntBaixa = BP.NumIntBaixa AND BP.Status <> 5 AND B.Status <> 5"
            sSQL = sSQL & " " & "AND B.NumMovCta = ? ORDER BY T.Cliente"
        
            lErro = Comando_Executar(alComando8(3), sSQL, lCliForn, sNomeCliForn, lNumMovto)
            If lErro <> AD_SQL_SUCESSO Then gError 213860
            
            lErro = Comando_BuscarPrimeiro(alComando8(3))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213861
            
            Do While lErro = AD_SQL_SUCESSO
        
                If lCliFornAnt <> lCliForn Then
                    If Len(Trim(sCliFornAux)) > 0 Then
                        sCliFornAux = sCliFornAux & "; "
                    Else
                        sCliFornAux = sTipoCliForn & ":"
                    End If
                    sCliFornAux = sCliFornAux & CStr(lCliForn) & "-" & sNomeCliForn
                    lCliFornAnt = lCliForn
                End If
        
                lErro = Comando_BuscarProximo(alComando8(3))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213862
             
            Loop
            
            sHistAux = ""
        
        Case 9, 10, 12 'Pagto de Título por Cheque/Borderô/Dinheiro - > BaixasPag
        
            sTipoCliForn = "F"
        
            sNomeCliForn = String(STRING_MAXIMO, 0)
            sHistorico = String(STRING_MAXIMO, 0)
        
            sSQL = "SELECT T.Fornecedor, C.NomeReduzido, T.Historico FROM TitulosPagTodos AS T, ParcelasPagTodas AS P, BaixasParcPag As BP, Fornecedores AS C, BaixasPag As B"
            sSQL = sSQL & " " & "WHERE T.NumIntDoc = P.NumIntTitulo And P.NumIntDoc = BP.NumIntParcela AND T.Fornecedor = C.Codigo AND B.NumIntBaixa = BP.NumIntBaixa AND BP.Status <> 5 AND B.Status <> 5"
            sSQL = sSQL & " " & "AND B.NumMovCta = ? ORDER BY T.Fornecedor"
        
            lErro = Comando_Executar(alComando8(4), sSQL, lCliForn, sNomeCliForn, sHistorico, lNumMovto)
            If lErro <> AD_SQL_SUCESSO Then gError 213863
            
            lErro = Comando_BuscarPrimeiro(alComando8(4))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213864
            
            Do While lErro = AD_SQL_SUCESSO
        
                If lCliFornAnt <> lCliForn Then
                    If Len(Trim(sCliFornAux)) > 0 Then
                        sCliFornAux = sCliFornAux & "; "
                    Else
                        sCliFornAux = sTipoCliForn & ":"
                    End If
                    sCliFornAux = sCliFornAux & CStr(lCliForn) & "-" & sNomeCliForn
                    lCliFornAnt = lCliForn
                End If
                If Len(Trim(sHistAux)) > 0 And Len(Trim(sHistorico)) > 0 Then sHistAux = sHistAux & "; "
                sHistAux = sHistAux & sHistorico
                
                lErro = Comando_BuscarProximo(alComando8(4))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213865
             
            Loop
    
    End Select
    
    lUltNumMovto = lNumMovto
    sUltFornCliTit = left(sCliFornAux, 250)
    sUltHistTit = left(sHistAux, 250)
    
    Obter_Dados_Titulo = SUCESSO
    
    Exit Function
    
Erro_Obter_Dados_Titulo:

    Obter_Dados_Titulo = gErr
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case 213855
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 213856, 213857
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PAGTO_ANTECIPADO2", gErr)

        Case 213858, 213859
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RECEB_ANTEC", gErr)

        Case 213860 To 213862
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PARCELAS_REC", gErr)

        Case 213863 To 213865
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PARCELAS_PAG2", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 213866)

    End Select

    Exit Function
    
End Function
'###############################################################

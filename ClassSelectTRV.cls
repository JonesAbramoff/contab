VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassSelectTRV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function TRVRateio_Le(ByVal colTRVRateio As Collection) As Long
'le os lançamentos relativos ao TRVRateio

Dim lComando As Long
Dim lErro As Long
Dim objTRVRateio As ClassTRVRateio
Dim sConta As String
Dim sCcl As String
Dim sDescricao As String
Dim iPeriodo As Integer
Dim iExercicio As Integer


On Error GoTo Erro_TRVRateio_Le
    
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 197622

    sConta = String(STRING_CONTA, 0)
    sCcl = String(STRING_CCL, 0)
    sDescricao = String(STRING_RATEIO_DESCRICAO, 0)
    
    lErro = Comando_Executar(lComando, "SELECT  Exercicio, Periodo, Conta, Ccl, Descricao  FROM TRVRateio ORDER BY Seq", iExercicio, iPeriodo, sConta, sCcl, sDescricao)
    If lErro <> AD_SQL_SUCESSO Then gError 197623

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197624

    Do While lErro <> AD_SQL_SEM_DADOS

        Set objTRVRateio = New ClassTRVRateio

        objTRVRateio.iExercicio = iExercicio
        objTRVRateio.iPeriodo = iPeriodo
        objTRVRateio.sConta = sConta
        objTRVRateio.sCcl = sCcl
        objTRVRateio.sDescricao = sDescricao

        colTRVRateio.Add objTRVRateio

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197625
        
    Loop
    
    Call Comando_Fechar(lComando)
    
    TRVRateio_Le = SUCESSO
    
    Exit Function
 
Erro_TRVRateio_Le:

    TRVRateio_Le = gErr

    Select Case gErr
    
        Case 197622
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 197623 To 197625
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVRATEIO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197626)
    
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function PeriodosFilial_Le(ByVal objPeriodosFilial As ClassPeriodosFilial) As Long
'le o registro referente ao periodo em questão e coloca os dados em objPeriodosFilial

Dim lComando As Long
Dim lErro As Long
Dim tPeriodo As typePeriodosFilial


On Error GoTo Erro_PeriodosFilial_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 10157
    
    lErro = Comando_Executar(lComando, "SELECT Lancamentos, Apurado, Lote, Fechado, DataApuracao, DocApuracao, FechadoCTB FROM PeriodosFilial WHERE FilialEmpresa = ? AND Exercicio = ? AND Periodo = ?", tPeriodo.iLancamentos, tPeriodo.iApurado, tPeriodo.iLote, tPeriodo.iFechado, tPeriodo.dtDataApuracao, tPeriodo.lDocApuracao, tPeriodo.iFechadoCTB, objPeriodosFilial.iFilialEmpresa, objPeriodosFilial.iExercicio, objPeriodosFilial.iPeriodo)
    If lErro <> AD_SQL_SUCESSO Then Error 10158
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 10159
    
    Call Comando_Fechar(lComando)
    
    objPeriodosFilial.iLancamentos = tPeriodo.iLancamentos
    objPeriodosFilial.iApurado = tPeriodo.iApurado
    objPeriodosFilial.iLote = tPeriodo.iLote
    
    If objPeriodosFilial.sOrigem = MODULO_CONTABILIDADE Then
        objPeriodosFilial.iFechado = tPeriodo.iFechadoCTB
    Else
        objPeriodosFilial.iFechado = tPeriodo.iFechado
    End If
        
    If objPeriodosFilial.sOrigem = "TMP" Then
        objPeriodosFilial.iFechado = DESMARCADO
    End If
        
    objPeriodosFilial.dtDataApuracao = tPeriodo.dtDataApuracao
    objPeriodosFilial.lDocApuracao = tPeriodo.lDocApuracao
    
    PeriodosFilial_Le = SUCESSO

    Exit Function

Erro_PeriodosFilial_Le:

    PeriodosFilial_Le = Err
    
    Select Case Err
    
    Case 10157
        lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
    Case 10158, 10159
        lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PERIODOSFILIAL", Err, objPeriodosFilial.iFilialEmpresa, objPeriodosFilial.iExercicio, objPeriodosFilial.iPeriodo)
    
    Case Else
        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153387)
        
    End Select
    
    Call Comando_Fechar(lComando)
       
    Exit Function
    
End Function

Function Periodo_Le_Todos_Exercicio(ByVal iFilialEmpresa As Integer, ByVal iExercicio As Integer, ByVal colPeriodos As Collection) As Long
'le os periodos referentes ao exercicio iExercicio e coloca-os em colPeriodos

Dim lComando As Long
Dim lErro As Long
Dim iFim_de_Arquivo As Integer
Dim objPeriodo As ClassPeriodo
Dim tPeriodo As typePeriodo

On Error GoTo Erro_Periodo_Le_Todos_Exercicio

    tPeriodo.sNomeExterno = String(STRING_PERIODO_NOME_EXTERNO, 0)

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 207135

    'Pesquisa os periodos do exercicio em questão
    lErro = Comando_Executar(lComando, "SELECT PeriodosFilial.Exercicio, PeriodosFilial.Periodo, NomeExterno, Lancamentos, Apurado, DataInicio, DataFim, Lote, Fechado, FechadoCTB FROM PeriodosFilial, Periodo WHERE PeriodosFilial.FilialEmpresa = ? AND PeriodosFilial.Exercicio = ? AND PeriodosFilial.Exercicio = Periodo.Exercicio AND PeriodosFilial.Periodo = Periodo.Periodo ORDER BY PeriodosFilial.Periodo", tPeriodo.iExercicio, tPeriodo.iPeriodo, tPeriodo.sNomeExterno, tPeriodo.iLancamentos, tPeriodo.iApurado, tPeriodo.dtDataInicio, tPeriodo.dtDataFim, tPeriodo.iLote, tPeriodo.iFechado, tPeriodo.iFechadoCTB, iFilialEmpresa, iExercicio)
    If lErro <> AD_SQL_SUCESSO Then gError 207136
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 207137
    
    Do While lErro = AD_SQL_SUCESSO

        Set objPeriodo = New ClassPeriodo
        
        objPeriodo.iExercicio = tPeriodo.iExercicio
        objPeriodo.iPeriodo = tPeriodo.iPeriodo
        objPeriodo.sNomeExterno = tPeriodo.sNomeExterno
        objPeriodo.iLancamentos = tPeriodo.iLancamentos
        objPeriodo.iApurado = tPeriodo.iApurado
        objPeriodo.dtDataInicio = tPeriodo.dtDataInicio
        objPeriodo.dtDataFim = tPeriodo.dtDataFim
        objPeriodo.iLote = tPeriodo.iLote
        objPeriodo.iFechado = tPeriodo.iFechado
        objPeriodo.iFechadoCTB = tPeriodo.iFechadoCTB
        
        colPeriodos.Add objPeriodo
        
        'le o proximo periodo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 207138
                
    Loop
    
    Call Comando_Fechar(lComando)
    
    Periodo_Le_Todos_Exercicio = SUCESSO
    
    Exit Function
    
    
Erro_Periodo_Le_Todos_Exercicio:

    Periodo_Le_Todos_Exercicio = gErr
    
    Select Case gErr
    
        Case 207135
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 207136 To 207138
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PERIODOS", gErr)
           
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207139)
        
    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function
    
End Function


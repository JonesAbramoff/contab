VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFLHGrava"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function Lote_Contabil_Folha(objIntfCtb As Object, objLote As ClassLote) As Long
'Cria o lote contabil para a rotina de interação com a folha de pagamento

Dim lTransacao As Long
Dim lErro As Long

On Error GoTo Erro_Lote_Contabil_Folha

    'Inicia a Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 92919
    
    objLote.iFilialEmpresa = objIntfCtb.iFilialEmpresa
    objLote.iExercicio = objIntfCtb.iExercicio
    objLote.iPeriodo = objIntfCtb.iPeriodo
    objLote.sOrigem = MODULO_FOLHA
    
    'obter numero do lote
    lErro = CF("Lote_Automatico1", objLote)
    If lErro <> SUCESSO Then gError 92918
    
    objLote.iStatus = LOTE_DESATUALIZADO
    
    'insere o lote pendente no banco de dados
    lErro = CF("LotePendente_Grava_Trans", objLote)
    If lErro <> SUCESSO Then gError 92920

    'Confirma a Transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 92921

    Lote_Contabil_Folha = SUCESSO
    
    Exit Function

Erro_Lote_Contabil_Folha:

    Lote_Contabil_Folha = gErr

    Select Case gErr

        Case 92918, 92920

        Case 92919
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 92921
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150471)
    
    End Select
    
    'Fechamento transação
    Call Transacao_Rollback

    Exit Function

End Function

Function Periodo_Le5(ByVal dtData As Date, ByVal objPeriodo As ClassPeriodo, ByVal iFilialEmpresa As Integer) As Long
'le o registro referente ao periodo em questão e coloca os dados em objPeriodo

Dim lComando As Long
Dim lErro As Long
Dim tPeriodo As typePeriodo

On Error GoTo Erro_Periodo_Le5

    tPeriodo.sNomeExterno = String(STRING_PERIODO_NOME_EXTERNO, 0)
    
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 6046
    
    lErro = Comando_Executar(lComando, "SELECT Periodo.Exercicio, Periodo.Periodo, NomeExterno, Lancamentos, Apurado, DataInicio, DataFim, Lote, Fechado FROM Periodo, PeriodosFilial WHERE FilialEmpresa = ? AND Periodo.DataInicio<=? AND Periodo.DataFim>=? AND Periodo.Exercicio = PeriodosFilial.Exercicio AND Periodo.Periodo = PeriodosFilial.Periodo", tPeriodo.iExercicio, tPeriodo.iPeriodo, tPeriodo.sNomeExterno, tPeriodo.iLancamentos, tPeriodo.iApurado, tPeriodo.dtDataInicio, tPeriodo.dtDataFim, tPeriodo.iLote, tPeriodo.iFechado, iFilialEmpresa, dtData, dtData)
    If lErro <> AD_SQL_SUCESSO Then Error 6047
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 6048
    
    Call Comando_Fechar(lComando)
    
    objPeriodo.iExercicio = tPeriodo.iExercicio
    objPeriodo.iPeriodo = tPeriodo.iPeriodo
    objPeriodo.sNomeExterno = tPeriodo.sNomeExterno
    objPeriodo.iLancamentos = tPeriodo.iLancamentos
    objPeriodo.iApurado = tPeriodo.iApurado
    objPeriodo.dtDataInicio = tPeriodo.dtDataInicio
    objPeriodo.dtDataFim = tPeriodo.dtDataFim
    objPeriodo.iLote = tPeriodo.iLote
    objPeriodo.iFechado = tPeriodo.iFechado
    
    Periodo_Le5 = SUCESSO

    Exit Function

Erro_Periodo_Le5:

    Periodo_Le5 = Err
    
    Select Case Err
    
    Case 6046
        lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
    Case 6047, 6048
        lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PERIODO", Err)
    
    Case Else
        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150472)
        
    End Select
    
    Call Comando_Fechar(lComando)
       
    Exit Function
    
End Function


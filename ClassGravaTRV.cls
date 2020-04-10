VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassGravaTRV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Function TRVRateio_Grava(ByVal colTRVRateio As Collection) As Long
'grava na tabela de Rateios Automaticos TRVRateio

Dim lComando As Long
Dim lComando1 As Long
Dim lComando2 As Long
Dim lErro As Long
Dim sConta As String
Dim iStatus As Integer
Dim objTRVRateio As ClassTRVRateio
Dim iTipoConta As Integer
Dim lTransacao As Long

On Error GoTo Erro_TRVRateio_Grava

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 197535

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 197536
    
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 197537
    
    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then gError 197538
    
    'Faz inserção do Lançamento na Tabela de RateioOff
    lErro = Comando_Executar(lComando1, "DELETE FROM TRVRateio")
    If lErro <> AD_SQL_SUCESSO Then gError 197539
    
    'Inicia Loop de GRAVACAO com Lançamentos do GRID
    For Each objTRVRateio In colTRVRateio
                
        'Pesquisa Conta associada a um Lançamento
        lErro = Comando_ExecutarLockado(lComando, "SELECT TipoConta, Ativo FROM PlanoConta WHERE Conta=?", iTipoConta, iStatus, objTRVRateio.sConta)
        If lErro <> AD_SQL_SUCESSO Then gError 197540
        
       'Lê a Conta
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197541
        
        If lErro = AD_SQL_SEM_DADOS Then gError 197542
        
        lErro = Comando_LockShared(lComando)
        If lErro <> AD_SQL_SUCESSO Then gError 197543
            
        'verifica se a conta está ativa
        If iStatus <> CONTA_ATIVA Then gError 197544
               
        'verifica se a conta é analítica
        If iTipoConta <> CONTA_ANALITICA Then gError 197545
        
        sConta = String(STRING_CONTA, 0)
        
        'Pesquisa a associação Conta x Ccl
        lErro = Comando_ExecutarLockado(lComando2, "SELECT Conta FROM ContaCcl WHERE Conta=? AND Ccl=?", sConta, objTRVRateio.sConta, objTRVRateio.sCcl)
        If lErro <> AD_SQL_SUCESSO Then gError 197546
        
        lErro = Comando_BuscarPrimeiro(lComando2)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197547
            
        If lErro = AD_SQL_SEM_DADOS Then gError 197548
            
        'lock da associacao conta x ccl
        lErro = Comando_LockExclusive(lComando2)
        If lErro <> AD_SQL_SUCESSO Then gError 197549
            
        lErro = CF("Ccl_Critica_Analitica", objTRVRateio.sCcl)
        If lErro <> SUCESSO Then gError 197550
        
        'Faz inserção do Lançamento na Tabela de RateioOff
        lErro = Comando_Executar(lComando1, "INSERT INTO TRVRateio (Conta, Ccl, Descricao, Exercicio, Periodo, Seq ) VALUES (?,?,?,?,?,?)", objTRVRateio.sConta, objTRVRateio.sCcl, objTRVRateio.sDescricao, objTRVRateio.iExercicio, objTRVRateio.iPeriodo, objTRVRateio.iSeq)
        If lErro <> AD_SQL_SUCESSO Then gError 197551
        
    Next

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    
    'Confirma transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 197552

    TRVRateio_Grava = SUCESSO
    
    Exit Function
    
Erro_TRVRateio_Grava:

    TRVRateio_Grava = gErr
    
    Select Case gErr
    
        Case 197535
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 197536 To 197538
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 197539
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRVRATEIO", gErr)
    
        Case 197540, 197541
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTA3", gErr, objTRVRateio.sConta)
    
        Case 197542
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTA1", gErr, objTRVRateio.sConta)

        Case 197543
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_PLANOCONTA", gErr, objTRVRateio.sConta)
    
        Case 197544
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTA_INATIVA", gErr)
    
        Case 197545
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_ANALITICA", gErr, objTRVRateio.sConta)
        
        Case 197546, 197547
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTACCL3", gErr, objTRVRateio.sConta, objTRVRateio.sCcl)
    
        Case 197548
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTACCL_NAO_CADASTRADO", gErr, objTRVRateio.sConta, objTRVRateio.sCcl)
        
        Case 197549
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_CONTACCL", gErr, objTRVRateio.sConta, objTRVRateio.sCcl)
    
        Case 197550
        
        Case 197551
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRVRATEIO", gErr)
    
        Case 197552
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 197553)
                
    End Select
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    
    Call Transacao_Rollback
    
    Exit Function

End Function

Public Function RateioOff_Gera_TRVRateio(ByVal colRateioOff As Collection, ByVal iExcluiRateio) As Long

Dim alComando(1 To 5) As Long
Dim lTransacao As Long
Dim iIndice As Integer
Dim lErro As Long
Dim objTRVRateio As ClassTRVRateio
Dim objRateioOff As ClassRateioOff
Dim sCcl As String
Dim dValor As Double
Dim colcolRateioOff As New Collection
Dim dValorTotal As Double
Dim dPercTotal As Double
Dim lCodigo As Long
Dim iSeq As Integer
Dim iAchou As Integer
Dim colTRVRateio As New Collection

On Error GoTo Erro_RateioOff_Gera_TRVRateio
    
    'Abrir transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 197628
    
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 197629
    Next

    If iExcluiRateio = 1 Then

        lErro = Comando_Executar(alComando(4), "DELETE FROM RateioOffContas WHERE EXISTS (SELECT * FROM RateioOff RA WHERE LEFT(RA.Descricao,3) = '(A)' AND RateioOffContas.Codigo = RA.Codigo) ")
        If lErro <> AD_SQL_SUCESSO Then gError 197630
    
        lErro = Comando_Executar(alComando(5), "DELETE FROM RateioOff WHERE LEFT(Descricao,3) = '(A)'")
        If lErro <> AD_SQL_SUCESSO Then gError 197631

    End If

    lErro = CF("TRVRateio_Le", colTRVRateio)
    If lErro <> SUCESSO Then gError 197632
    
    For Each objTRVRateio In colTRVRateio

        'Mostra número do proximo rateio disponível
        lErro = CF("RateioOff_Automatico_Trans", lCodigo)
        If lErro <> SUCESSO Then gError 197633
        
'        Set objRateioOff = New ClassRateioOff
'
'        iAchou = 0
'
'        For Each colRateioOff In colcolRateioOff
'
'            If colRateioOff.Count > 1 Then
'
'                Set objRateioOff = colRateioOff(1)
'
'                If Left(objRateioOff.sCcl, 1) = Left(objTRVRateio.sCcl, 1) Then
'                    iAchou = 1
'                    Exit For
'                End If
'
'            End If
'
'        Next
'
'
'        If iAchou = 0 Then
'
'            Set colRateioOff = New Collection
'
'            sCcl = String(STRING_CCL, 0)
'
'            dValorTotal = 0
'
'            lErro = Comando_Executar(alComando(1), "SELECT Ccl, SUM(Cre" & Format(objTRVRateio.iPeriodo, "00") & " - Deb" & Format(objTRVRateio.iPeriodo, "00") & ")" & _
'            " FROM MvPerCcl WHERE FilialEmpresa = ? AND Exercicio = ? AND LEFT(Conta,1) = '5' AND LEFT(Ccl,1) = ? AND SUBSTRING(Ccl,3,1) >= '3' GROUP BY Ccl", sCcl, dValor, iFilialEmpresa, objTRVRateio.iExercicio, Left(objTRVRateio.sCcl, 1))
'            If lErro <> AD_SQL_SUCESSO Then gError 197634
'
'            lErro = Comando_BuscarPrimeiro(alComando(1))
'            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197635
'
'            If lErro = AD_SQL_SUCESSO Then
'
'            Do While lErro = AD_SQL_SUCESSO
'
'                If dValor <> 0 Then
'
'                    Set objRateioOff = New ClassRateioOff
'
'                    objRateioOff.sCcl = sCcl
'                    objRateioOff.dValor = dValor
'
'                    colRateioOff.Add objRateioOff
'
'                    dValorTotal = dValorTotal + Abs(dValor)
'
'                End If
'
'                lErro = Comando_BuscarProximo(alComando(1))
'                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197636
'
'            Loop
'
'            dPercTotal = 0
'
'            For Each objRateioOff In colRateioOff
'
'                objRateioOff.dPercentual = Round(Abs(objRateioOff.dValor) / dValorTotal, 4)
'                dPercTotal = dPercTotal + objRateioOff.dPercentual
'
'            Next
'
'            If dPercTotal <> 1 Then
'
'                Set objRateioOff = colRateioOff(1)
'
'                objRateioOff.dPercentual = objRateioOff.dPercentual + Round((1 - dPercTotal), 4)
'
'            End If
'
'            colcolRateioOff.Add colRateioOff
'
'        End If
'
'        End If

        iSeq = 0

        For Each objRateioOff In colRateioOff
        
            iSeq = iSeq + 1
        
            lErro = Comando_Executar(alComando(2), "INSERT INTO RateioOff (Codigo, Seq, CclOrigem, ContaCre, Tipo, Conta, Percentual, Descricao, Ccl) VALUES (?,?,?,?,?,?,?,?,?)", _
                    lCodigo, iSeq, objTRVRateio.sCcl, objTRVRateio.sConta, TIPO_RATEIOOFF_MENSAL, objTRVRateio.sConta, objRateioOff.dPercentual, "(A)" & objTRVRateio.sDescricao, objRateioOff.sCcl)
            If lErro <> AD_SQL_SUCESSO Then gError 197637
        
        Next
        
        lErro = Comando_Executar(alComando(3), "INSERT INTO RateioOffContas (Codigo, Item, ContaInicial, ContaFinal) VALUES (?,?,?,?)", _
                lCodigo, 1, objTRVRateio.sConta, objTRVRateio.sConta)
        If lErro <> AD_SQL_SUCESSO Then gError 197638
        
    Next
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    'Fechar transacao
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 197639
    
    RateioOff_Gera_TRVRateio = SUCESSO
    
    Exit Function
 
Erro_RateioOff_Gera_TRVRateio:

    RateioOff_Gera_TRVRateio = gErr

    Select Case gErr
    
        Case 197628
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
    
        Case 197629
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 197630
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_RATEIOOFFCONTAS", gErr)
    
        Case 197631
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_RATEIOOFF", gErr)
    
        Case 197632, 197633
    
        Case 197634, 197635, 197636
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVPERCCL", gErr)
    
        Case 197637
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RATEIOOFF", gErr)
    
        Case 197638
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RATEIOOFFCONTAS", gErr)
    
        Case 197639
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197640)
    
    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Call Transacao_Rollback
    
    Exit Function

End Function

Public Function Excel_Le_RateioOff(ByVal sArquivo As String, ByVal objMsg As Object, ByVal objProgresso As Object, ByVal objTela As Object, colRateioOff As Collection) As Long

Dim lErro As Long
'Declarações para manipulação do Excel
Dim objPastaTrabalho As Object 'Excel.Workbook
Dim objPlanilhaExcel As Object 'Excel.Worksheet
Dim objFaixaCelulasPreencher As Object 'Excel.Range
Dim iIndice As Integer
Dim sArq As String
Dim iLinha As Integer
Dim iLinhas As Integer
Dim dTotalRateio As Double
Dim objRateioOff As ClassRateioOff
Dim sCcl As String

On Error GoTo Erro_Excel_Le_RateioOff

    sArq = sArquivo
    
    'Abre o excel
    lErro = CF("Excel_Abrir")
    If lErro <> 0 Then gError 197751
    
    'Desabilita as mensagens do Excel
    GL_objExcelSist.DisplayAlerts = False
    
    objMsg.AddItem "Abrindo a planilha " & sArq
    
    Set objPastaTrabalho = GL_objExcelSist.Workbooks.Open(sArq)
    
    DoEvents
    
    If objTela.giStop = 1 Then gError 197752
    
    objMsg.AddItem "Planilha aberta, iniciando importação"
    
    'Seleciona a planilha ativa na pasta de trabalho criada
    Set objPlanilhaExcel = objPastaTrabalho.ActiveSheet
    
    sCcl = objPlanilhaExcel.Cells(2, 1).Value
    iLinhas = 0
    
    Do While Len(Trim(sCcl)) <> 0
        iLinhas = iLinhas + 1
        sCcl = objPlanilhaExcel.Cells(iLinhas + 1, 1).Value
    Loop
    
    iLinhas = iLinhas - 1
    objProgresso.Max = iLinhas
    objProgresso.Min = 0
    objProgresso.Value = 0
    
    dTotalRateio = 0
    
    For iLinha = 1 To iLinhas
    
        Set objRateioOff = New ClassRateioOff
        
        objRateioOff.sCcl = objPlanilhaExcel.Cells(iLinha + 1, 1).Value
        
        objRateioOff.dPercentual = StrParaDbl(objPlanilhaExcel.Cells(iLinha + 1, 2).Value)
    
        colRateioOff.Add objRateioOff
    
        dTotalRateio = dTotalRateio + objRateioOff.dPercentual
                    
        objProgresso.Value = objProgresso.Value + 1
                    
        DoEvents
        
        If objTela.giStop = 1 Then gError 197753
                    
    Next
    
    If Abs(dTotalRateio - 1) > DELTA_VALORMONETARIO Then gError 197754
    
    Call CF("Excel_Fechar")
    
    objMsg.AddItem "Arquivo " & sArq & " importado com sucesso."
    
    Excel_Le_RateioOff = SUCESSO
    
    Exit Function
    
Erro_Excel_Le_RateioOff:

    Excel_Le_RateioOff = gErr
    
    Select Case gErr

        Case 197751
            Call Rotina_Erro(vbOKOnly, "ERRO_IMPORTACAO_LINHA", gErr)

        Case 197752, 197753
            Call Rotina_Erro(vbOKOnly, "ERRO_IMPORTACAO_INTERROMPIDA_USUARIO", gErr)
        
        Case 197754
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 197755)

    End Select

    'Fecha o Excel
    Call CF("Excel_Fechar")

    Exit Function

End Function

Public Function Executa_ComissaoRegra1_Cust(bForcaInclusao As Boolean) As Long
    bForcaInclusao = True
    Executa_ComissaoRegra1_Cust = SUCESSO
End Function

Function PeriodosFilial_Insere_Exercicio(ByVal iExercicio As Integer, ByVal iNumPeriodos As Integer, ByVal colFiliais As Collection) As Long
'Insere os PeriodosFilial relativos a um exercicio
'ATENÇÃO: Esta rotina tem que ser chamada por um programa que esteja em transação

Dim lErro As Long
Dim lComando As Long
Dim iPeriodo As Integer
Dim objFiliais As AdmFiliais

On Error GoTo Erro_PeriodosFilial_Insere_Exercicio

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 207108

    For iPeriodo = 1 To iNumPeriodos

        For Each objFiliais In colFiliais
    
            If objFiliais.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO Or objFiliais.iCodFilial = EMPRESA_TODA Then
    
                'Insere o Periodo em questão
                lErro = Comando_Executar(lComando, "INSERT INTO PeriodosFilial (FilialEmpresa, Exercicio, Periodo, Lancamentos, Apurado, Lote, Fechado, DocApuracao, FechadoCTB) VALUES (?,?,?,?,?,?,?,?,?)", objFiliais.iCodFilial, iExercicio, iPeriodo, LANCAMENTOS_NO_BD, PERIODO_NAO_APURADO, 0, PERIODO_ABERTO, 0, PERIODO_ABERTO)
                If lErro <> AD_SQL_SUCESSO Then gError 207109
            
            End If
            
        Next

    Next

    Call Comando_Fechar(lComando)

    PeriodosFilial_Insere_Exercicio = SUCESSO

    Exit Function

Erro_PeriodosFilial_Insere_Exercicio:

    PeriodosFilial_Insere_Exercicio = gErr

    Select Case gErr

        Case 207108
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 207109
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PERIODOSFILIAL", gErr, objFiliais.iCodFilial, iExercicio, iPeriodo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207110)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function PeriodosFilial_Altera_Status(ByVal iFilialEmpresa As Integer, ByVal iExercicio As Integer, ByVal colPeriodos As Collection) As Long
'altera o status dos periodos do exercicio passado como parametro. O Exercicio não possui movimento e não teve alterado o numero de periodos.
'TEM QUE SER EXECUTADO EM TRANSACAO

Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim iPeriodo As Integer
Dim objPeriodo As ClassPeriodo
Dim iIndice As Integer

On Error GoTo Erro_PeriodosFilial_Altera_Status

    For iIndice = LBound(alComando) To UBound(alComando)
    
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 207111
        
    Next

    'Alterar status dos períodos da filial em questão
    lErro = Comando_ExecutarPos(alComando(1), "SELECT Periodo FROM PeriodosFilial WHERE FilialEmpresa=? AND Exercicio = ?", 0, iPeriodo, iFilialEmpresa, iExercicio)
    If lErro <> AD_SQL_SUCESSO Then gError 207112

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 207113

    Do While lErro = AD_SQL_SUCESSO

        Set objPeriodo = colPeriodos.Item(iPeriodo)

        lErro = Comando_ExecutarPos(alComando(2), "UPDATE PeriodosFilial SET Fechado=?, FechadoCTB=?", alComando(1), objPeriodo.iFechado, objPeriodo.iFechadoCTB)
        If lErro <> AD_SQL_SUCESSO Then gError 207114

        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 207115
        
    Loop

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PeriodosFilial_Altera_Status = SUCESSO

    Exit Function

Erro_PeriodosFilial_Altera_Status:

    PeriodosFilial_Altera_Status = gErr

    Select Case gErr
    
        Case 207111
             lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 107112, 207113, 207115
             lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PERIODOSFILIAL2", gErr, iFilialEmpresa, iExercicio)

        Case 207114
             lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_PERIODOSFILIAL", gErr, iPeriodo, iExercicio, iFilialEmpresa)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 207116)
             
    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function


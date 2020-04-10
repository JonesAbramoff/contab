VERSION 1.0 CLASS
BEGIN
  MultiUse = 0   'False
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1batch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Quando selecionar os lotes a serem atualizados no interativo
'ao acionar o botao OK os lotes serão atualizados com um número
'único (gerado pelo sistema) e este número será passado como parametro
'para o programa que faz a atualizacao. Baseado neste numero o programa
'de Atualizacao loca cada lote e atualiza-o. Nao precisa locar todos
'os lotes logo no inicio. O fato de um lote estar marcado com um numero'nao impede que o mesmo apareça no interativo.

Function Rotina_Atualizacao(ByVal sNomeArqParam As String, ByVal iID_Atualizacao As Integer) As Long

Dim lErro As Long

On Error GoTo Erro_Rotina_Atualizacao

    lErro = Sistema_Abrir_Batch(sNomeArqParam)
    If lErro <> SUCESSO Then Error 45055

    Set gcolModulo = New AdmColModulo
    
    lErro = CF("Modulos_Le_Empresa_Filial", glEmpresa, giFilialEmpresa, gcolModulo)
    If lErro <> SUCESSO Then Error 55968

    TelaAcompanhaBatch.Caption = TITULO_TELABATCH_ATUALIZA_LOTE
    TelaAcompanhaBatch.sNomeArqParam = sNomeArqParam
    TelaAcompanhaBatch.iIdAtualizacao_Param = iID_Atualizacao
    TelaAcompanhaBatch.iRotinaBatch = ROTINA_ATUALIZACAO_BATCH
    
    TelaAcompanhaBatch.Show

    Rotina_Atualizacao = SUCESSO
    
    Exit Function

Erro_Rotina_Atualizacao:

    Rotina_Atualizacao = Err

    Select Case Err

        Case 45055, 55968
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144617)

    End Select
    
    Exit Function

End Function

Function Rotina_Desatualizacao(ByVal sOrigem As String, ByVal iExercicio As Integer, ByVal iPeriodo As Integer, ByVal iLote As Integer) As Long
'desatualiza o lote passado como parametro

Dim lComando1 As Long
Dim lComando As Long
Dim lComando2 As Long
Dim lComando3 As Long
Dim lComando4 As Long
Dim lErro As Long
Dim iStatus As Integer
Dim tLote As typeLote_batch
Dim lTransacao As Long
Dim iUsoCcl As Integer

On Error GoTo Erro_Rotina_Desatualizacao
    
    lComando = 0
    lComando1 = 0
    lComando2 = 0
    lComando3 = 0
    lComando4 = 0
    lTransacao = 0

    tLote.sOrigem = sOrigem
    tLote.iExercicio = iExercicio
    tLote.iLote = iLote
    tLote.iPeriodo = iPeriodo
    
    lComando4 = Comando_Abrir()
    If lComando4 = 0 Then Error 9403

    lErro = Comando_Executar(lComando4, "SELECT UsoCcl FROM Configuracao", iUsoCcl)
    If lErro <> AD_SQL_SUCESSO Then Error 9404
    
    lErro = Comando_BuscarPrimeiro(lComando4)
    If lErro <> AD_SQL_SUCESSO Then Error 9404

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 5129

    'Pesquisa o Exercicio em questão
    lErro = Comando_ExecutarLockado(lComando2, "SELECT Status FROM Exercicios WHERE Exercicio = ?", iStatus, tLote.iExercicio)
    If lErro <> AD_SQL_SUCESSO Then Error 5130

    'le o Exercicio
    lErro = Comando_BuscarPrimeiro(lComando2)
    If lErro <> AD_SQL_SUCESSO Then Error 5131

    'Se o Exercicio não estiver aberto ==> erro
    If iStatus <> EXERCICIO_ABERTO Then Error 5132

    'não permite mudança no status do exercicio
    lErro = Comando_LockShared(lComando2)
    If lErro <> AD_SQL_SUCESSO Then Error 5133

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 5055

    'Pesquisa o periodo em questão
    lErro = Comando_ExecutarLockado(lComando1, "SELECT Exercicio,Periodo FROM Periodo WHERE Exercicio = ? AND Periodo = ? )", iExercicio, iPeriodo, tLote.iExercicio, tLote.iPeriodo)
    If lErro <> AD_SQL_SUCESSO Then Error 5056

    'Le o periodo
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO Then Error 5057

    'Lock do Periodo
    lErro = Comando_LockExclusive(lComando1)
    If lErro <> AD_SQL_SUCESSO Then Error 5058

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 5059

    'Pesquisa o lote no banco de dados
    lErro = Comando_ExecutarPos(lComando, "SELECT Status FROM Lote WHERE Origem = ? AND Exercicio = ? AND Periodo = ? AND Lote = ?", 0, iStatus, tLote.sOrigem, tLote.iExercicio, tLote.iPeriodo, tLote.iLote)
    If lErro <> AD_SQL_SUCESSO Then Error 5060

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 5061

    'Verifica se o lote não está atualizado ==> Erro
    If iStatus <> LOTE_ATUALIZADO Then Error 5062

    'Lock do Lote
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 5063

    'Inicia a transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 5243

    'Processa os lançamentos do lote
    lErro = Processa_Lancamentos(tLote, DESATUALIZACAO, iUsoCcl, 0)
    If lErro <> SUCESSO Then Error 5242

    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then Error 5240

    'marca o lote como desatualizado
    lErro = Comando_ExecutarPos(lComando3, "UPDATE Lote SET Status = ?", lComando, LOTE_DESATUALIZADO)
    If lErro <> AD_SQL_SUCESSO Then Error 5241

    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 5244

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    Call Comando_Fechar(lComando4)

    Rotina_Desatualizacao = SUCESSO

    Exit Function

Erro_Rotina_Desatualizacao:

    Rotina_Desatualizacao = Err

    Select Case Err

        Case 5055
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 5056, 5057
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PERIODOS", Err, tLote.iExercicio, tLote.iPeriodo)

        Case 5058
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_PERIODO", Err, iExercicio, iLote)

        Case 5059
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 5060, 5061
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LOTE", Err, tLote.sOrigem, tLote.iExercicio, tLote.iPeriodo, tLote.iLote)
    
        Case 5062
            Call Rotina_Erro(vbOKOnly, "ERRO_LOTE_NAO_ATUALIZADO", Err, tLote.sOrigem, tLote.iExercicio, tLote.iPeriodo, tLote.iLote)

        Case 5063
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_LOTE", Err, tLote.sOrigem, tLote.iExercicio, tLote.iPeriodo, tLote.iLote)

        Case 5129
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 5130, 5131
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EXERCICIO", Err, tLote.iExercicio)

        Case 5132
            Call Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_NAO_ABERTO", Err, tLote.iExercicio)

        Case 5133
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCACAO_EXERCICIO", Err, tLote.iExercicio)

        Case 5240
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 5241
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_LOTE", Err, tLote.sOrigem, tLote.iExercicio, tLote.iPeriodo, tLote.iLote)

        Case 5242

        Case 5243
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 5244
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case 9403
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 9404, 9405
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONFIGURACAO", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144618)

    End Select

    Call Transacao_Rollback
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    Call Comando_Fechar(lComando4)
        
    Exit Function

End Function

Function Rotina_Apura_Exercicio(ByVal sNomeArqParam As String, ByVal iFilialEmpresa As Integer, ByVal iExercicio As Integer, ByVal iLote As Integer, ByVal sHistorico As String, ByVal sContaResultado As String, ByVal colContasApuracao As Collection) As Long
'realiza a apuração do exercicio iExercicio para as receitas e despesas passadas como parametro e gera um lote contendo a conta resultado da apuracao.

Dim lErro As Long

On Error GoTo Erro_Rotina_Apura_Exercicio

    lErro = Sistema_Abrir_Batch(sNomeArqParam)
    If lErro <> SUCESSO Then Error 45054

    lErro = Carrega_ColFiliais()
    If lErro <> SUCESSO Then Error 55839

    TelaAcompanhaBatch.Caption = TITULO_TELABATCH_APURA_EXERCICIO
    TelaAcompanhaBatch.sNomeArqParam = sNomeArqParam
    TelaAcompanhaBatch.iFilialEmpresa = iFilialEmpresa
    TelaAcompanhaBatch.iExercicio = iExercicio
    TelaAcompanhaBatch.iLote = iLote
    TelaAcompanhaBatch.sHistorico = sHistorico
    TelaAcompanhaBatch.sContaResultado = sContaResultado
    TelaAcompanhaBatch.iRotinaBatch = ROTINA_APURA_EXERCICIO_BATCH
    Set TelaAcompanhaBatch.colContasApuracao = colContasApuracao
    
    TelaAcompanhaBatch.Show

    Rotina_Apura_Exercicio = SUCESSO
    
    Exit Function

Erro_Rotina_Apura_Exercicio:

    Rotina_Apura_Exercicio = Err

    Select Case Err

        Case 45054, 55839
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144619)

    End Select
    
    Exit Function

End Function

Function Rotina_Fechamento_Exercicio(ByVal sNomeArqParam As String, ByVal iExercicio As Integer, ByVal sConta_Ativo_Inicial As String, ByVal sConta_Ativo_Final As String, ByVal sConta_Passivo_Inicial As String, ByVal sConta_Passivo_Final As String) As Long

Dim lErro As Long

On Error GoTo Erro_Rotina_Fechamento_Exercicio

    lErro = Sistema_Abrir_Batch(sNomeArqParam)
    If lErro <> SUCESSO Then Error 45057

    lErro = Carrega_ColFiliais()
    If lErro <> SUCESSO Then Error 55840

    TelaAcompanhaBatch.Caption = TITULO_TELABATCH_FECHA_EXERCICIO
    TelaAcompanhaBatch.sNomeArqParam = sNomeArqParam
    TelaAcompanhaBatch.iExercicio = iExercicio
    TelaAcompanhaBatch.sConta_Ativo_Inicial = sConta_Ativo_Inicial
    TelaAcompanhaBatch.sConta_Ativo_Final = sConta_Ativo_Final
    TelaAcompanhaBatch.sConta_Passivo_Inicial = sConta_Passivo_Inicial
    TelaAcompanhaBatch.sConta_Passivo_Final = sConta_Passivo_Final
    TelaAcompanhaBatch.iRotinaBatch = ROTINA_FECHAMENTO_EXERCICIO_BATCH
    
    TelaAcompanhaBatch.Show

    Rotina_Fechamento_Exercicio = SUCESSO
    
    Exit Function

Erro_Rotina_Fechamento_Exercicio:

    Rotina_Fechamento_Exercicio = Err

    Select Case Err

        Case 45057, 55840
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144620)

    End Select
    
    Exit Function

End Function

Function Rotina_Reabertura_Exercicio(ByVal sNomeArqParam As String, ByVal iExercicio As Integer, ByVal sConta_Ativo_Inicial As String, ByVal sConta_Ativo_Final As String, ByVal sConta_Passivo_Inicial As String, ByVal sConta_Passivo_Final As String) As Long

Dim lErro As Long

On Error GoTo Erro_Rotina_Reabertura_Exercicio

    lErro = Sistema_Abrir_Batch(sNomeArqParam)
    If lErro <> SUCESSO Then Error 45059

    lErro = Carrega_ColFiliais()
    If lErro <> SUCESSO Then Error 55841

    TelaAcompanhaBatch.Caption = TITULO_TELABATCH_REABRE_EXERCICIO
    TelaAcompanhaBatch.sNomeArqParam = sNomeArqParam
    TelaAcompanhaBatch.iExercicio = iExercicio
    TelaAcompanhaBatch.sConta_Ativo_Inicial = sConta_Ativo_Inicial
    TelaAcompanhaBatch.sConta_Ativo_Final = sConta_Ativo_Final
    TelaAcompanhaBatch.sConta_Passivo_Inicial = sConta_Passivo_Inicial
    TelaAcompanhaBatch.sConta_Passivo_Final = sConta_Passivo_Final
    TelaAcompanhaBatch.iRotinaBatch = ROTINA_REABERTURA_EXERCICIO_BATCH
    
    TelaAcompanhaBatch.Show

    Rotina_Reabertura_Exercicio = SUCESSO
    
    Exit Function

Erro_Rotina_Reabertura_Exercicio:

    Rotina_Reabertura_Exercicio = Err

    Select Case Err

        Case 45059, 55841
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144621)

    End Select
    
    Exit Function

End Function

Function Rotina_Apura_Periodos(ByVal sNomeArqParam As String, ByVal iFilialEmpresa As Integer, ByVal iExercicio As Integer, ByVal iPeriodo_Inicial As Integer, ByVal iPeriodo_Final As Integer, ByVal sContaResultado As String, ByVal sContaPonte As String, ByVal colContasApuracao As Collection, ByVal sHistorico As String, Optional ByVal iZeraRD As Integer = 0) As Long

Dim lErro As Long

On Error GoTo Erro_Rotina_Apura_Periodos

    lErro = Sistema_Abrir_Batch(sNomeArqParam)
    If lErro <> SUCESSO Then Error 27440
    
    lErro = Carrega_ColFiliais()
    If lErro <> SUCESSO Then Error 55842
    
    TelaAcompanhaBatch.Caption = TITULO_TELABATCH_APURA_PERIODO
    TelaAcompanhaBatch.sNomeArqParam = sNomeArqParam
    TelaAcompanhaBatch.iFilialEmpresa = iFilialEmpresa
    TelaAcompanhaBatch.iExercicio = iExercicio
    TelaAcompanhaBatch.iPeriodo_Inicial = iPeriodo_Inicial
    TelaAcompanhaBatch.iPeriodo_Final = iPeriodo_Final
    TelaAcompanhaBatch.sHistorico = sHistorico
    TelaAcompanhaBatch.sContaResultado = sContaResultado
    TelaAcompanhaBatch.sContaPonte = sContaPonte
    TelaAcompanhaBatch.iRotinaBatch = ROTINA_APURA_PERIODOS_BATCH
    TelaAcompanhaBatch.iZeraRD = iZeraRD
    Set TelaAcompanhaBatch.colContasApuracao = colContasApuracao
    
    TelaAcompanhaBatch.Show

    Rotina_Apura_Periodos = SUCESSO
    
    Exit Function

Erro_Rotina_Apura_Periodos:

    Rotina_Apura_Periodos = Err

    Select Case Err

        Case 27440, 55842

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144622)

    End Select
    
    Exit Function

End Function

Function Rotina_Reprocessamento(ByVal sNomeArqParam As String, ByVal iFilialEmpresa As Integer, ByVal iExercicio As Integer, ByVal iPeriodo As Integer) As Long

Dim lErro As Long

On Error GoTo Erro_Rotina_Reprocessamento

    lErro = Sistema_Abrir_Batch(sNomeArqParam)
    If lErro <> SUCESSO Then Error 45060

    lErro = Carrega_ColFiliais()
    If lErro <> SUCESSO Then Error 55843

    TelaAcompanhaBatch.Caption = TITULO_TELABATCH_REPROCESSAMENTO
    TelaAcompanhaBatch.sNomeArqParam = sNomeArqParam
    TelaAcompanhaBatch.iFilialEmpresa = iFilialEmpresa
    TelaAcompanhaBatch.iExercicio = iExercicio
    TelaAcompanhaBatch.iPeriodo = iPeriodo
    TelaAcompanhaBatch.iRotinaBatch = ROTINA_REPROCESSAMENTO_BATCH
    
    TelaAcompanhaBatch.Show

    Rotina_Reprocessamento = SUCESSO
    
    Exit Function

Erro_Rotina_Reprocessamento:

    Rotina_Reprocessamento = Err

    Select Case Err

        Case 45060, 55843
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144623)

    End Select
    
    Exit Function

End Function

Function Rotina_RateioOff(ByVal sNomeArqParam As String, ByVal iLote As Integer, ByVal dtData As Date, ByVal iPeriodoInicial As Integer, ByVal iPeriodoFinal As Integer, ByVal colRateios As Collection, ByVal iFilialEmpresa As Integer, ByVal sHistorico As String) As Long

Dim lErro As Long
Dim objRateioOffBatch As New ClassRateioOffBatch

On Error GoTo Erro_Rotina_RateioOff

    lErro = Sistema_Abrir_Batch(sNomeArqParam)
    If lErro <> SUCESSO Then Error 45058

    objRateioOffBatch.iLote = iLote
    objRateioOffBatch.dtData = dtData
    objRateioOffBatch.iPeriodoInicial = iPeriodoInicial
    objRateioOffBatch.iPeriodoFinal = iPeriodoFinal
    Set objRateioOffBatch.colRateios = colRateios
    objRateioOffBatch.iFilialEmpresa = iFilialEmpresa
    objRateioOffBatch.sHistorico = sHistorico

    TelaAcompanhaBatch.Caption = TITULO_TELABATCH_RATEIOOFF
    TelaAcompanhaBatch.sNomeArqParam = sNomeArqParam
    Set TelaAcompanhaBatch.objRateioOffBatch = objRateioOffBatch
    TelaAcompanhaBatch.iRotinaBatch = ROTINA_RATEIOOFF_BATCH
    
    TelaAcompanhaBatch.Show

    Rotina_RateioOff = SUCESSO
    
    Exit Function

Erro_Rotina_RateioOff:

    Rotina_RateioOff = Err

    Select Case Err

        Case 45058
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144624)

    End Select
    
    Exit Function

End Function

Public Function Rotina_Gerar_DRE_DRP(ByVal sNomeArqParam As String, ByVal colModelos As Collection, ByVal iExercicio As Integer, ByVal iPeriodo As Integer, ByVal iFilialEmpresa As Integer, ByVal sDiretorio As String, Optional ByVal iGrupoEmpresarial As Integer = 0) As Long

Dim lErro As Long

On Error GoTo Erro_Rotina_Gerar_DRE_DRP

    Set TelaAcompanhaBatch2.colModelos = colModelos
    
    TelaAcompanhaBatch2.iExercicio = iExercicio
    TelaAcompanhaBatch2.iPeriodo = iPeriodo
    TelaAcompanhaBatch2.iFilialEmpresa = iFilialEmpresa
    TelaAcompanhaBatch2.sDiretorio = sDiretorio
    TelaAcompanhaBatch2.iGrupoEmpresarial = iGrupoEmpresarial

    TelaAcompanhaBatch2.Caption = "Geração de DRE\DRP em excel"
    TelaAcompanhaBatch2.sNomeArqParam = sNomeArqParam

    'Exibe a tela de acompanhamento do batch
    TelaAcompanhaBatch2.Show

    Rotina_Gerar_DRE_DRP = SUCESSO

    Exit Function

Erro_Rotina_Gerar_DRE_DRP:

    Rotina_Gerar_DRE_DRP = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187119)

    End Select

    Exit Function

End Function

Function Rotina_Desapura_Exercicio(ByVal sNomeArqParam As String, ByVal iFilialEmpresa As Integer, ByVal iExercicio As Integer, ByVal iLote As Integer, ByVal sHistorico As String, ByVal sContaResultado As String, ByVal colContasApuracao As Collection) As Long
'realiza a apuração do exercicio iExercicio para as receitas e despesas passadas como parametro e gera um lote contendo a conta resultado da apuracao.

Dim lErro As Long

On Error GoTo Erro_Rotina_Desapura_Exercicio

    lErro = Sistema_Abrir_Batch(sNomeArqParam)
    If lErro <> SUCESSO Then gError 188420

    lErro = Carrega_ColFiliais()
    If lErro <> SUCESSO Then gError 55839

    TelaAcompanhaBatch.Caption = TITULO_TELABATCH_DESAPURA_EXERCICIO
    TelaAcompanhaBatch.sNomeArqParam = sNomeArqParam
    TelaAcompanhaBatch.iFilialEmpresa = iFilialEmpresa
    TelaAcompanhaBatch.iExercicio = iExercicio
    TelaAcompanhaBatch.iLote = iLote
    TelaAcompanhaBatch.sHistorico = sHistorico
    TelaAcompanhaBatch.sContaResultado = sContaResultado
    TelaAcompanhaBatch.iRotinaBatch = ROTINA_DESAPURA_EXERCICIO_BATCH
    Set TelaAcompanhaBatch.colContasApuracao = colContasApuracao
    
    TelaAcompanhaBatch.Show

    Rotina_Desapura_Exercicio = SUCESSO
    
    Exit Function

Erro_Rotina_Desapura_Exercicio:

    Rotina_Desapura_Exercicio = gErr

    Select Case gErr

        Case 188420, 188421
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 188422)

    End Select
    
    Exit Function

End Function

Function Rotina_Sped_Contabil_Diario(ByVal sNomeArqParam As String, ByVal iFilialEmpresa As Integer, ByVal sDiretorio As String, ByVal dtDataIni As Date, ByVal dtDataFim As Date, ByVal lNumOrd As Long, ByVal sContaOutros As String, ByVal iIndSituacaoPer As Integer, ByVal iIndSitEspecial As Integer, Optional ByVal iCodVersao As Integer = 1, Optional ByVal iIndNIRE As Integer = 0, Optional ByVal iFinalidade As Integer = 0, Optional ByVal sHashEscrSubst As String = "", Optional ByVal sNIRESubst As String = "", Optional ByVal iEmpGrandePorte As Integer = 0, Optional ByVal iTipoECD As Integer = 0, Optional ByVal sCodSCP As String = "", Optional ByVal colSCPs As Collection = Nothing) As Long
'executa o diario sped

Dim lErro As Long

On Error GoTo Erro_Rotina_Sped_Contabil_Diario

    lErro = Sistema_Abrir_Batch(sNomeArqParam)
    If lErro <> SUCESSO Then gError 203251

    TelaAcompanhaBatch.Caption = TITULO_SPED_DIARIO
    TelaAcompanhaBatch.sNomeArqParam = sNomeArqParam
    TelaAcompanhaBatch.iFilialEmpresa = iFilialEmpresa
    TelaAcompanhaBatch.sDiretorio = sDiretorio
    TelaAcompanhaBatch.dtDataIni = dtDataIni
    TelaAcompanhaBatch.dtDataFim = dtDataFim
    TelaAcompanhaBatch.lNumOrd = lNumOrd
    TelaAcompanhaBatch.iRotinaBatch = ROTINA_SPED_DIARIO
    TelaAcompanhaBatch.sContaOutros = sContaOutros
    
    TelaAcompanhaBatch.iIndSituacaoPer = iIndSituacaoPer
    TelaAcompanhaBatch.iIndSitEspecial = iIndSitEspecial
    
    TelaAcompanhaBatch.iCodVersao = iCodVersao
    TelaAcompanhaBatch.iIndNIRE = iIndNIRE
    TelaAcompanhaBatch.iFinalidade = iFinalidade
    TelaAcompanhaBatch.sHashEscrSubst = sHashEscrSubst
    TelaAcompanhaBatch.sNIRESubst = sNIRESubst
    TelaAcompanhaBatch.iEmpGrandePorte = iEmpGrandePorte
    
    TelaAcompanhaBatch.iTipoECD = iTipoECD
    TelaAcompanhaBatch.sCodSCP = sCodSCP
    Set TelaAcompanhaBatch.colSCPs = colSCPs
    
    TelaAcompanhaBatch.Show

    Rotina_Sped_Contabil_Diario = SUCESSO
    
    Exit Function

Erro_Rotina_Sped_Contabil_Diario:

    Rotina_Sped_Contabil_Diario = gErr

    Select Case gErr

        Case 203251
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 203252)

    End Select
    
    Exit Function

End Function

Function Rotina_Sped_FCont(ByVal sNomeArqParam As String, ByVal iFilialEmpresa As Integer, ByVal sDiretorio As String, ByVal dtDataIni As Date, ByVal dtDataFim As Date, ByVal lNumOrd As Long, ByVal sContaOutros As String, ByVal iIndSituacaoPer As Integer, ByVal iIndSitEspecial As Integer) As Long
'executa o diario sped

Dim lErro As Long

On Error GoTo Erro_Rotina_Sped_FCont

    lErro = Sistema_Abrir_Batch(sNomeArqParam)
    If lErro <> SUCESSO Then gError 203251

    TelaAcompanhaBatch.Caption = TITULO_FCONT
    TelaAcompanhaBatch.sNomeArqParam = sNomeArqParam
    TelaAcompanhaBatch.iFilialEmpresa = iFilialEmpresa
    TelaAcompanhaBatch.sDiretorio = sDiretorio
    TelaAcompanhaBatch.dtDataIni = dtDataIni
    TelaAcompanhaBatch.dtDataFim = dtDataFim
    TelaAcompanhaBatch.lNumOrd = lNumOrd
    TelaAcompanhaBatch.iIndSituacaoPer = iIndSituacaoPer
    TelaAcompanhaBatch.iIndSitEspecial = iIndSitEspecial
    TelaAcompanhaBatch.iRotinaBatch = ROTINA_FCONT
    TelaAcompanhaBatch.sContaOutros = sContaOutros
    
    TelaAcompanhaBatch.Show

    Rotina_Sped_FCont = SUCESSO
    
    Exit Function

Erro_Rotina_Sped_FCont:

    Rotina_Sped_FCont = gErr

    Select Case gErr

        Case 203251
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 203252)

    End Select
    
    Exit Function

End Function

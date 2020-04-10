VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1online"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function Rotina_Atualiza_Lancamentos(ByVal objLancamento As ClassLancamento_Cabecalho, Optional ByVal bExclusaoFisica As Boolean = False) As Long
'atualiza os lançamentos de um voucher
'TEM QUE SER CHAMADO DENTRO DE TRANSACAO

Dim lErro As Long
Dim lPosicao As Long
Dim tLancamento_Sort As typeLancamento_Sort
Dim lComando1 As Long
Dim lComando2 As Long
Dim tProcessa_Lancamento As typeProcessa_Lancamento
Dim alComando(1 To 6) As Long
Dim alComando1(1 To 3) As Long
Dim iIndice As Integer

On Error GoTo Erro_Rotina_Atualiza_Lancamentos

    lComando1 = 0
    lComando2 = 0

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 9535

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then gError 9538

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 55835
    Next

    For iIndice = 1 To 3
        alComando1(iIndice) = alComando(iIndice)
    Next
    
    lErro = Atualiza_ExercicioFilial_Nao_Apurado(alComando1(), objLancamento.iExercicio, objLancamento.iFilialEmpresa)
    If lErro <> SUCESSO Then gError 55836

    For iIndice = 1 To 3
        alComando1(iIndice) = alComando(iIndice + 3)
    Next
    
    lErro = Atualiza_Periodo_Não_Apurado(alComando1(), objLancamento.iExercicio, objLancamento.iPeriodoLan, objLancamento.iFilialEmpresa)
    If lErro <> SUCESSO Then gError 55837

    'verifica se deve aglutinar por dia o lançamento
    lErro = Verifica_Aglutina_Lancam_Por_Dia(objLancamento.sOrigem, tProcessa_Lancamento.iAglutinaLancamPorDia)
    If lErro <> SUCESSO Then gError 55859

    lErro = Comando_Executar(lComando1, "SELECT UsoCcl FROM Configuracao", tProcessa_Lancamento.iUsoCcl)
    If lErro <> AD_SQL_SUCESSO Then gError 9536
    
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO Then gError 9537

    tProcessa_Lancamento.iFilialEmpresa = objLancamento.iFilialEmpresa
    If bExclusaoFisica Then
        tProcessa_Lancamento.iOperacao = DESATUALIZACAO
    Else
        tProcessa_Lancamento.iOperacao = ATUALIZACAO
    End If
    tProcessa_Lancamento.iPeriodo = objLancamento.iPeriodoLan
    tProcessa_Lancamento.sPeriodo = Format(tProcessa_Lancamento.iPeriodo, "00")
    tProcessa_Lancamento.iExercicio = objLancamento.iExercicio
    tProcessa_Lancamento.iOperacao1 = ROTINA_ATUALIZACAO_ONLINE
    tProcessa_Lancamento.tLancamento.iFilialEmpresa = objLancamento.iFilialEmpresa
    tProcessa_Lancamento.tLancamento.sOrigem = objLancamento.sOrigem
    tProcessa_Lancamento.tLancamento.iExercicio = objLancamento.iExercicio
    tProcessa_Lancamento.tLancamento.iPeriodoLan = objLancamento.iPeriodoLan
    tProcessa_Lancamento.tLancamento.lDoc = objLancamento.lDoc
    tProcessa_Lancamento.tLancamento.iPeriodoLote = objLancamento.iPeriodoLote
    
    tProcessa_Lancamento.tLancamento.sConta = String(STRING_CONTA, 0)
    tProcessa_Lancamento.tLancamento.sCcl = String(STRING_CCL, 0)
    tProcessa_Lancamento.tLancamento.sHistorico = String(STRING_HISTORICO, 0)
    tProcessa_Lancamento.tLancamento.sProduto = String(STRING_PRODUTO, 0)
    tProcessa_Lancamento.tLancamento.sDocOrigem = String(STRING_DOCORIGEM, 0)
    tProcessa_Lancamento.tLancamento.sModelo = String(STRING_PADRAOCONTAB_MODELO, 0)
    tProcessa_Lancamento.sCclAglutinado = "***"

    'Pesquisa os lançamentos do voucher
    lErro = Comando_ExecutarPos(lComando2, "SELECT  Seq, Data, Conta, Ccl, Historico, Valor, NumIntDoc, FilialCliForn, CliForn, Transacao, Aglutina, Produto, ApropriaCRProd, ContaSimples, SeqContraPartida, EscaninhoCusto, ValorLivroAuxiliar, ClienteFornecedor, DocOrigem, Quantidade, DataEstoque, Status, Modelo, Gerencial, Subtipo  FROM LanPendente WHERE FilialEmpresa = ? AND Origem = ? AND Exercicio = ? AND PeriodoLan = ? AND Doc = ? ORDER BY Conta, Data", 0, _
    tProcessa_Lancamento.tLancamento.iSeq, tProcessa_Lancamento.tLancamento.dtData, tProcessa_Lancamento.tLancamento.sConta, tProcessa_Lancamento.tLancamento.sCcl, tProcessa_Lancamento.tLancamento.sHistorico, tProcessa_Lancamento.tLancamento.dValor, tProcessa_Lancamento.tLancamento.lNumIntDoc, tProcessa_Lancamento.tLancamento.iFilialCliForn, tProcessa_Lancamento.tLancamento.lCliForn, tProcessa_Lancamento.tLancamento.iTransacao, tProcessa_Lancamento.tLancamento.iAglutina, tProcessa_Lancamento.tLancamento.sProduto, tProcessa_Lancamento.tLancamento.iApropriaCRProd, tProcessa_Lancamento.tLancamento.lContaSimples, tProcessa_Lancamento.tLancamento.iSeqContraPartida, tProcessa_Lancamento.tLancamento.iEscaninho_Custo, tProcessa_Lancamento.tLancamento.dValorLivroAuxiliar, tProcessa_Lancamento.tLancamento.iClienteFornecedor, tProcessa_Lancamento.tLancamento.sDocOrigem, tProcessa_Lancamento.tLancamento.dQuantidade, tProcessa_Lancamento.tLancamento.dtDataEstoque, tProcessa_Lancamento.tLancamento.iStatus, _
    tProcessa_Lancamento.tLancamento.sModelo, tProcessa_Lancamento.tLancamento.iGerencial, tProcessa_Lancamento.tLancamento.iSubTipo, _
    tProcessa_Lancamento.tLancamento.iFilialEmpresa, tProcessa_Lancamento.tLancamento.sOrigem, tProcessa_Lancamento.tLancamento.iExercicio, tProcessa_Lancamento.tLancamento.iPeriodoLan, tProcessa_Lancamento.tLancamento.lDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 9539

    tProcessa_Lancamento.lComando2 = lComando2

    'Le o primeiro lançamento pertencente ao lote
    lErro = Comando_BuscarPrimeiro(lComando2)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 9540
    
    If lErro = AD_SQL_SUCESSO Then

        lErro = Processa_Lancamentos_1(tProcessa_Lancamento)
        If lErro <> SUCESSO Then gError 83582

    End If

    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    tProcessa_Lancamento.lComando2 = 0
    
    Rotina_Atualiza_Lancamentos = SUCESSO

    Exit Function

Erro_Rotina_Atualiza_Lancamentos:

    Rotina_Atualiza_Lancamentos = gErr

    Select Case gErr

        Case 9535, 9538, 55835
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 9539, 9540
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LANCAMENTOS_PENDENTES", gErr)

        Case 9536, 9537
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONFIGURACAO", gErr)

        Case 55836, 55837, 55859, 83582

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144625)

    End Select

    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function Rotina_Reprocessamento_DocOrigem(ByVal iOrigemLcto As Integer, ByVal lNumIntDocOrigem As Long, ByVal colExercicio As Collection, ByVal iFilialEmpresa As Integer) As Long
'utilizado pelo reprocessamento para alterar os lançamentos associados a docorigem que tiverem o produto preenchido e a apropriação seja pelo custo médio ou standard

Dim lErro As Long
Dim lComando1 As Long
Dim lComando2 As Long
Dim lComando3 As Long
Dim tProcessa_Lancamento As typeProcessa_Lancamento
Dim iUsoCcl As Integer

On Error GoTo Erro_Rotina_Reprocessamento_DocOrigem

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 83529

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then gError 83530

    tProcessa_Lancamento.lComando2 = lComando2

    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then gError 83823

    lErro = Comando_Executar(lComando1, "SELECT UsoCcl FROM Configuracao", iUsoCcl)
    If lErro <> AD_SQL_SUCESSO Then gError 83527
    
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO Then gError 83528

    tProcessa_Lancamento.iFilialEmpresa = iFilialEmpresa
    tProcessa_Lancamento.iOperacao = ATUALIZACAO
    tProcessa_Lancamento.iUsoCcl = iUsoCcl
    tProcessa_Lancamento.iOperacao1 = ROTINA_REPROC_CUSTO_MEDIO
    tProcessa_Lancamento.sCclAglutinado = "***"
    tProcessa_Lancamento.iAglutinaLancamPorDia = AGLUTINA_LANCAM_POR_DIA

    tProcessa_Lancamento.tLancamento.sOrigem = String(STRING_ORIGEM, 0)
    Set tProcessa_Lancamento.colExercicio = colExercicio

    'Pesquisa os lançamentos pertencentes ao NumIntDoc/OrigemLcto
    lErro = Comando_Executar(lComando3, "SELECT DISTINCT FilialEmpresa, Origem, Exercicio, PeriodoLan, Doc FROM Lancamentos, TransacaoCTBCodigo WHERE FilialEmpresa = ? AND NumIntDoc = ? AND Lancamentos.Transacao = TransacaoCTBCodigo.Codigo AND TransacaoCTBCodigo.OrigemLcto = ? ORDER BY FilialEmpresa, Origem, Exercicio, PeriodoLan, Doc", _
    tProcessa_Lancamento.tLancamento.iFilialEmpresa, tProcessa_Lancamento.tLancamento.sOrigem, tProcessa_Lancamento.tLancamento.iExercicio, tProcessa_Lancamento.tLancamento.iPeriodoLan, tProcessa_Lancamento.tLancamento.lDoc, iFilialEmpresa, lNumIntDocOrigem, iOrigemLcto)
    If lErro <> AD_SQL_SUCESSO Then gError 83824

    'Le o primeiro lançamento pertencente ao NumIntDoc/OrigemLcto
    lErro = Comando_BuscarPrimeiro(lComando3)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 83825
    
    Do While lErro = AD_SQL_SUCESSO
    
        tProcessa_Lancamento.tLancamento.sConta = String(STRING_CONTA, 0)
        tProcessa_Lancamento.tLancamento.sCcl = String(STRING_CCL, 0)
        tProcessa_Lancamento.tLancamento.sHistorico = String(STRING_HISTORICO, 0)
        tProcessa_Lancamento.tLancamento.sProduto = String(STRING_PRODUTO, 0)
        tProcessa_Lancamento.tLancamento.sDocOrigem = String(STRING_DOCORIGEM, 0)
        tProcessa_Lancamento.tLancamento.sModelo = String(STRING_PADRAOCONTAB_MODELO, 0)
    
        'Pesquisa os lançamentos pertencentes a FilialEmpresa/Origem/Exercicio/PeriodoLan/Doc e que tenham o campo produto preenchido
        lErro = Comando_ExecutarPos(lComando2, "SELECT Seq, Lote, PeriodoLote, Data, Conta, Ccl, Historico, Valor, NumIntDoc, FilialCliForn, CliForn, Transacao, DocAglutinado, SeqAglutinado, Aglutinado, Produto, ApropriaCRProd, ContaSimples, SeqContraPartida, EscaninhoCusto, ValorLivroAuxiliar, ClienteFornecedor, DocOrigem, Quantidade, DataEstoque, Status, Modelo, Subtipo FROM Lancamentos WHERE FilialEmpresa = ? AND Origem = ? AND Exercicio = ? AND PeriodoLan = ? AND Doc = ? AND Produto <> '' ORDER BY Conta, Data, Ccl", 0, _
        tProcessa_Lancamento.tLancamento.iSeq, tProcessa_Lancamento.tLancamento.iLote, tProcessa_Lancamento.tLancamento.iPeriodoLote, tProcessa_Lancamento.tLancamento.dtData, tProcessa_Lancamento.tLancamento.sConta, tProcessa_Lancamento.tLancamento.sCcl, tProcessa_Lancamento.tLancamento.sHistorico, tProcessa_Lancamento.tLancamento.dValor, tProcessa_Lancamento.tLancamento.lNumIntDoc, tProcessa_Lancamento.tLancamento.iFilialCliForn, tProcessa_Lancamento.tLancamento.lCliForn, tProcessa_Lancamento.tLancamento.iTransacao, tProcessa_Lancamento.tLancamento.lDocAglutinado, tProcessa_Lancamento.tLancamento.iSeqAglutinado, tProcessa_Lancamento.tLancamento.iAglutina, tProcessa_Lancamento.tLancamento.sProduto, tProcessa_Lancamento.tLancamento.iApropriaCRProd, _
                                        tProcessa_Lancamento.tLancamento.lContaSimples, tProcessa_Lancamento.tLancamento.iSeqContraPartida, tProcessa_Lancamento.tLancamento.iEscaninho_Custo, tProcessa_Lancamento.tLancamento.dValorLivroAuxiliar, tProcessa_Lancamento.tLancamento.iClienteFornecedor, tProcessa_Lancamento.tLancamento.sDocOrigem, tProcessa_Lancamento.tLancamento.dQuantidade, tProcessa_Lancamento.tLancamento.dtDataEstoque, tProcessa_Lancamento.tLancamento.iStatus, tProcessa_Lancamento.tLancamento.sModelo, tProcessa_Lancamento.tLancamento.iSubTipo, tProcessa_Lancamento.tLancamento.iFilialEmpresa, tProcessa_Lancamento.tLancamento.sOrigem, tProcessa_Lancamento.tLancamento.iExercicio, tProcessa_Lancamento.tLancamento.iPeriodoLan, tProcessa_Lancamento.tLancamento.lDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 83531
        
        'Le o primeiro lançamento pertencente ao lote
        lErro = Comando_BuscarPrimeiro(lComando2)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 83532
        
        tProcessa_Lancamento.iPeriodo = tProcessa_Lancamento.tLancamento.iPeriodoLan
        tProcessa_Lancamento.sPeriodo = Format(tProcessa_Lancamento.tLancamento.iPeriodoLan, "00")
        tProcessa_Lancamento.iExercicio = tProcessa_Lancamento.tLancamento.iExercicio
        
        If lErro = AD_SQL_SUCESSO Then
    
            lErro = Processa_Lancamentos_1(tProcessa_Lancamento)
            If lErro <> SUCESSO Then gError 83542
    
        End If

        'Le o proximo lançamento pertencente ao NumIntDoc/OrigemLcto
        lErro = Comando_BuscarProximo(lComando3)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 83826

    Loop

    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    
    'Atualiza o custo dos lançamentos pendentes que tem o documento origem passado como parametro, se existirem.
    lErro = CF("LanPendente_Reprocessa", iOrigemLcto, lNumIntDocOrigem, tProcessa_Lancamento.iFilialEmpresa)
    If lErro <> SUCESSO Then gError 83820
    
    tProcessa_Lancamento.lComando2 = 0
    
    Rotina_Reprocessamento_DocOrigem = SUCESSO

    Exit Function

Erro_Rotina_Reprocessamento_DocOrigem:

    Rotina_Reprocessamento_DocOrigem = gErr

    Select Case gErr

        Case 83527, 83528
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONFIGURACAO", gErr)

        Case 83529, 83530, 83823
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 83824, 83825, 83826
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LANCAMENTOS9", gErr, lNumIntDocOrigem, iOrigemLcto)

        Case 83542, 83820
        
        Case 83531, 83532
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LANCAMENTOS8", gErr, tProcessa_Lancamento.tLancamento.iFilialEmpresa, tProcessa_Lancamento.tLancamento.sOrigem, tProcessa_Lancamento.tLancamento.iExercicio, tProcessa_Lancamento.tLancamento.iPeriodoLan, tProcessa_Lancamento.tLancamento.lDoc)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144626)

    End Select

    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)

    Exit Function

End Function

Function Rotina_Reproc_Exclui_Lanc(ByVal iFilialEmpresa As Integer, ByVal iOrigemLcto As Integer, ByVal lNumIntDocOrigem As Long, ByVal colExercicio As Collection) As Long
'utilizado pelo reprocessamento para excluir as contabilizações associados ao ajuste do custo standard

Dim lErro As Long

On Error GoTo Erro_Rotina_Reproc_Exclui_Lanc

    'exclui os lançamentos pendentes associados a Filial/OrigemLcto/lNumIntDocOrigem se houverem
    lErro = Processa_Exclusao_LanPendente(iFilialEmpresa, iOrigemLcto, lNumIntDocOrigem, colExercicio)
    If lErro <> SUCESSO Then gError 83625
        
    'exclui os lançamentos já contabilizados  associados a Filial/OrigemLcto/lNumIntDocOrigem se houverem
    lErro = Processa_Exclusao_Lancamentos(iFilialEmpresa, iOrigemLcto, lNumIntDocOrigem, colExercicio)
    If lErro <> AD_SQL_SUCESSO Then gError 83624
    
    Rotina_Reproc_Exclui_Lanc = SUCESSO

    Exit Function

Erro_Rotina_Reproc_Exclui_Lanc:

    Rotina_Reproc_Exclui_Lanc = gErr

    Select Case gErr

        Case 83624, 83625
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144627)

    End Select

    Exit Function

End Function

Function Rotina_Reprocessamento_CProd(ByVal iFilialEmpresa As Integer, ByVal iMes As Integer, ByVal iAno As Integer, ByVal colExercicio As Collection) As Long
'altera os lançamentos do mes especificado com relação ao custo de produção

Dim lErro As Long
Dim lComando2 As Long
Dim lComando1 As Long
Dim tProcessa_Lancamento As typeProcessa_Lancamento
Dim iUsoCcl As Integer
Dim dtDataInicial As Date
Dim dtDataFinal As Date

On Error GoTo Erro_Rotina_Reprocessamento_CProd

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 83544

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then gError 83545

    lErro = Comando_Executar(lComando1, "SELECT UsoCcl FROM Configuracao", iUsoCcl)
    If lErro <> AD_SQL_SUCESSO Then gError 83546
    
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO Then gError 83547

    tProcessa_Lancamento.iFilialEmpresa = iFilialEmpresa
    tProcessa_Lancamento.iOperacao = ATUALIZACAO
    tProcessa_Lancamento.iUsoCcl = iUsoCcl
    tProcessa_Lancamento.iOperacao1 = ROTINA_REPROC_CUSTO_PRODUCAO
    tProcessa_Lancamento.sCclAglutinado = "***"
    tProcessa_Lancamento.iAglutinaLancamPorDia = AGLUTINA_LANCAM_POR_DIA

    tProcessa_Lancamento.tLancamento.sOrigem = String(STRING_ORIGEM, 0)
    tProcessa_Lancamento.tLancamento.sConta = String(STRING_CONTA, 0)
    tProcessa_Lancamento.tLancamento.sCcl = String(STRING_CCL, 0)
    tProcessa_Lancamento.tLancamento.sHistorico = String(STRING_HISTORICO, 0)
    tProcessa_Lancamento.tLancamento.sProduto = String(STRING_PRODUTO, 0)
    tProcessa_Lancamento.tLancamento.sDocOrigem = String(STRING_DOCORIGEM, 0)
    tProcessa_Lancamento.tLancamento.sModelo = String(STRING_PADRAOCONTAB_MODELO, 0)
    
    dtDataInicial = CDate("1/" & CStr(iMes) & "/" & CStr(iAno))
    dtDataFinal = DateAdd("m", 1, dtDataInicial) - 1

    Set tProcessa_Lancamento.colExercicio = colExercicio
    
    'Pesquisa os lançamentos pertencentes ao lote
    lErro = Comando_ExecutarPos(lComando2, "SELECT FilialEmpresa, Origem, Exercicio, PeriodoLan, Doc, Seq, Lote, PeriodoLote, Data, Conta, Ccl, Historico, Valor, NumIntDoc, FilialCliForn, CliForn, Transacao, DocAglutinado, SeqAglutinado, Aglutinado, Produto, ApropriaCRProd, ContaSimples, SeqContraPartida, EscaninhoCusto, ValorLivroAuxiliar, ClienteFornecedor, DocOrigem, Quantidade, DataEstoque, Status, Modelo, Subtipo FROM Lancamentos WHERE FilialEmpresa = ? AND DataEstoque >= ? And DataEstoque <= ? And Produto <> '' ORDER BY Conta, Data, Ccl", _
    0, tProcessa_Lancamento.tLancamento.iFilialEmpresa, tProcessa_Lancamento.tLancamento.sOrigem, tProcessa_Lancamento.tLancamento.iExercicio, tProcessa_Lancamento.tLancamento.iPeriodoLan, tProcessa_Lancamento.tLancamento.lDoc, tProcessa_Lancamento.tLancamento.iSeq, tProcessa_Lancamento.tLancamento.iLote, tProcessa_Lancamento.tLancamento.iPeriodoLote, tProcessa_Lancamento.tLancamento.dtData, tProcessa_Lancamento.tLancamento.sConta, tProcessa_Lancamento.tLancamento.sCcl, tProcessa_Lancamento.tLancamento.sHistorico, tProcessa_Lancamento.tLancamento.dValor, tProcessa_Lancamento.tLancamento.lNumIntDoc, tProcessa_Lancamento.tLancamento.iFilialCliForn, tProcessa_Lancamento.tLancamento.lCliForn, tProcessa_Lancamento.tLancamento.iTransacao, tProcessa_Lancamento.tLancamento.lDocAglutinado, tProcessa_Lancamento.tLancamento.iSeqAglutinado, tProcessa_Lancamento.tLancamento.iAglutina, tProcessa_Lancamento.tLancamento.sProduto, tProcessa_Lancamento.tLancamento.iApropriaCRProd, _
                                    tProcessa_Lancamento.tLancamento.lContaSimples, tProcessa_Lancamento.tLancamento.iSeqContraPartida, tProcessa_Lancamento.tLancamento.iEscaninho_Custo, tProcessa_Lancamento.tLancamento.dValorLivroAuxiliar, tProcessa_Lancamento.tLancamento.iClienteFornecedor, tProcessa_Lancamento.tLancamento.sDocOrigem, tProcessa_Lancamento.tLancamento.dQuantidade, tProcessa_Lancamento.tLancamento.dtDataEstoque, tProcessa_Lancamento.tLancamento.iStatus, tProcessa_Lancamento.tLancamento.sModelo, tProcessa_Lancamento.tLancamento.iSubTipo, iFilialEmpresa, dtDataInicial, dtDataFinal)
    If lErro <> AD_SQL_SUCESSO Then gError 83548

    'Le o primeiro lançamento pertencente ao periodo
    lErro = Comando_BuscarPrimeiro(lComando2)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 83549

    tProcessa_Lancamento.iPeriodo = tProcessa_Lancamento.tLancamento.iPeriodoLote
    tProcessa_Lancamento.sPeriodo = Format(tProcessa_Lancamento.tLancamento.iPeriodoLote, "00")
    tProcessa_Lancamento.iExercicio = tProcessa_Lancamento.tLancamento.iExercicio
    
    tProcessa_Lancamento.lComando2 = lComando2

    If lErro = AD_SQL_SUCESSO Then

        lErro = Processa_Lancamentos_1(tProcessa_Lancamento)
        If lErro <> SUCESSO Then gError 83550

    End If

    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    
    tProcessa_Lancamento.lComando2 = 0
    
    Rotina_Reprocessamento_CProd = SUCESSO

    Exit Function

Erro_Rotina_Reprocessamento_CProd:

    Rotina_Reprocessamento_CProd = gErr

    Select Case gErr

        Case 83544, 83545
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 83546, 83547
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONFIGURACAO", gErr)

        Case 83548, 83549
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LANCAMENTOS3", gErr)

        Case 83550
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144628)

    End Select

    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

    Exit Function

End Function




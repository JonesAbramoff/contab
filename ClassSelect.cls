VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'''????Fernando. Favor Transferir para GlobaisContab. Mario
Const STRING_LANPENDENTE_DOCORIGEM = 50


Function Periodo_Le1(ByVal dtData As Date, ByVal objPeriodo As ClassPeriodo, ByVal iFilialEmpresa As Integer) As Long
'le o registro referente ao periodo em questão e coloca os dados em objPeriodo

Dim lComando As Long
Dim lErro As Long
Dim tPeriodo As typePeriodo

On Error GoTo Erro_Periodo_Le1

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
    
    Periodo_Le1 = SUCESSO

    Exit Function

Erro_Periodo_Le1:

    Periodo_Le1 = Err
    
    Select Case Err
    
    Case 6046
        lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
    Case 6047, 6048
        lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PERIODO", Err)
        'MsgBox ("filialempresa: " & CStr(iFilialEmpresa) & " " & CStr(dtData))
    
    Case Else
        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153386)
        
    End Select
    
    Call Comando_Fechar(lComando)
       
    Exit Function
    
End Function

Function Periodo_Le(ByVal dtData As Date, ByVal objPeriodo As ClassPeriodo) As Long
'Le o registro referente ao periodo em questão e coloca os dados em objPeriodo
'Específica p/giFilialEmpresa

    Periodo_Le = CF("Periodo_Le1", dtData, objPeriodo, giFilialEmpresa)
    
End Function

Function PeriodosFilial_Le(ByVal objPeriodosFilial As ClassPeriodosFilial) As Long
'le o registro referente ao periodo em questão e coloca os dados em objPeriodosFilial

Dim lComando As Long
Dim lErro As Long
Dim tPeriodo As typePeriodosFilial

On Error GoTo Erro_PeriodosFilial_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 10157
    
    lErro = Comando_Executar(lComando, "SELECT Lancamentos, Apurado, Lote, Fechado, DataApuracao, DocApuracao FROM PeriodosFilial WHERE FilialEmpresa = ? AND Exercicio = ? AND Periodo = ?", tPeriodo.iLancamentos, tPeriodo.iApurado, tPeriodo.iLote, tPeriodo.iFechado, tPeriodo.dtDataApuracao, tPeriodo.lDocApuracao, objPeriodosFilial.iFilialEmpresa, objPeriodosFilial.iExercicio, objPeriodosFilial.iPeriodo)
    If lErro <> AD_SQL_SUCESSO Then Error 10158
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 10159
    
    Call Comando_Fechar(lComando)
    
    objPeriodosFilial.iLancamentos = tPeriodo.iLancamentos
    objPeriodosFilial.iApurado = tPeriodo.iApurado
    objPeriodosFilial.iLote = tPeriodo.iLote
    objPeriodosFilial.iFechado = tPeriodo.iFechado
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

Function Ccl_Le_Todos(ByVal colCcl As Collection) As Long
'le todos os centros de custo/lucro contidos na tabela Ccl e coloca-os em colCcl

Dim lComando As Long
Dim lErro As Long
Dim objCcl As ClassCcl
Dim tCcl As typeCcl
    
On Error GoTo Erro_Ccl_Le_Todos

    tCcl.sCcl = String(STRING_CCL, 0)
    tCcl.sDescCcl = String(STRING_CCL_DESCRICAO, 0)

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 5915

    lErro = Comando_Executar(lComando, "SELECT Ccl, DescCcl, TipoCcl FROM Ccl ORDER BY Ccl", tCcl.sCcl, tCcl.sDescCcl, tCcl.iTipoCcl)
    If lErro <> AD_SQL_SUCESSO Then Error 5916
    
    'le o primeiro centro de custo
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 5917
    
    Do While lErro <> AD_SQL_SEM_DADOS

        Set objCcl = New ClassCcl
        
        objCcl.sCcl = tCcl.sCcl
        objCcl.sDescCcl = tCcl.sDescCcl
        objCcl.iTipoCcl = tCcl.iTipoCcl
        
        'coloca o centro de custo lido na coleção
        colCcl.Add objCcl
        
        'le o proximo centro de custo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 5918

    Loop

    Call Comando_Fechar(lComando)
    
    Ccl_Le_Todos = SUCESSO

    Exit Function

Erro_Ccl_Le_Todos:

    Ccl_Le_Todos = Err
    
    Select Case Err

        Case 5915
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 5916, 5917, 5918
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CCL1", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153388)

    End Select
    
    Call Comando_Fechar(lComando)
     
    Exit Function
    
End Function

Function Ccl_Le_Todos_Analiticos(ByVal colCcl As Collection) As Long
'le todos os centros de custo/lucro analiticos contidos na tabela Ccl e coloca-os em colCcl

Dim lComando As Long
Dim lErro As Long
Dim objCcl As ClassCcl
Dim tCcl As typeCcl
    
On Error GoTo Erro_Ccl_Le_Todos_Analiticos

    tCcl.sCcl = String(STRING_CCL, 0)
    tCcl.sDescCcl = String(STRING_CCL_DESCRICAO, 0)

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 10478

    lErro = Comando_Executar(lComando, "SELECT Ccl, DescCcl, TipoCcl FROM Ccl WHERE TipoCcl=? ORDER BY Ccl", tCcl.sCcl, tCcl.sDescCcl, tCcl.iTipoCcl, CCL_ANALITICA)
    If lErro <> AD_SQL_SUCESSO Then Error 10479
    
    'le o primeiro centro de custo
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 10480
    
    Do While lErro <> AD_SQL_SEM_DADOS

        Set objCcl = New ClassCcl
        
        objCcl.sCcl = tCcl.sCcl
        objCcl.sDescCcl = tCcl.sDescCcl
        objCcl.iTipoCcl = tCcl.iTipoCcl
        
        'coloca o centro de custo lido na coleção
        colCcl.Add objCcl
        
        'le o proximo centro de custo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 10481

    Loop

    Call Comando_Fechar(lComando)
    
    Ccl_Le_Todos_Analiticos = SUCESSO

    Exit Function

Erro_Ccl_Le_Todos_Analiticos:

    Ccl_Le_Todos_Analiticos = Err
    
    Select Case Err

        Case 10478
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 10479, 10480, 10481
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CCL1", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153389)

    End Select
    
    Call Comando_Fechar(lComando)
     
    Exit Function
    
End Function

Function PlanoConta_Le_Todos(ByVal colPlanoConta As Collection) As Long
'le todos as contas contidas na tabela de Plano de Contas e coloca-as em colPlanoConta

Dim lComando As Long
Dim lErro As Long
Dim objPlanoConta As ClassPlanoConta
Dim tPlanoConta As typePlanoConta

On Error GoTo Erro_PlanoConta_Le_Todos

    tPlanoConta.sConta = String(STRING_CONTA, 0)
    tPlanoConta.sDescConta = String(STRING_CONTA_DESCRICAO, 0)

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 5907
    
    'le a tabela de plano de Contas ordenado por conta
    lErro = Comando_Executar(lComando, "SELECT Conta, NivelConta, DescConta, TipoConta, ContaSimples, Natureza, DigitoVerif, DigitoVerifSimples, HistPadrao, Ativo, UsaContaSimples, Categoria,FluxoCaixa, NaturezaSped FROM PlanoConta ORDER BY Conta", tPlanoConta.sConta, tPlanoConta.iNivelConta, tPlanoConta.sDescConta, tPlanoConta.iTipoConta, tPlanoConta.lContaSimples, tPlanoConta.iNatureza, tPlanoConta.iDigitoVerif, tPlanoConta.iDigitoVerifSimples, tPlanoConta.iHistPadrao, tPlanoConta.iAtivo, tPlanoConta.iUsaContaSimples, tPlanoConta.iCategoria, tPlanoConta.iFluxoCaixa, tPlanoConta.iNaturezaSped)
    If lErro <> AD_SQL_SUCESSO Then Error 5908
    
    'le a primeira conta
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 5909
    
    Do While lErro <> AD_SQL_SEM_DADOS

        Set objPlanoConta = New ClassPlanoConta
        
        objPlanoConta.sConta = tPlanoConta.sConta
        objPlanoConta.iNivelConta = tPlanoConta.iNivelConta
        objPlanoConta.sDescConta = tPlanoConta.sDescConta
        objPlanoConta.iTipoConta = tPlanoConta.iTipoConta
        objPlanoConta.lContaSimples = tPlanoConta.lContaSimples
        objPlanoConta.iNatureza = tPlanoConta.iNatureza
        objPlanoConta.iDigitoVerif = tPlanoConta.iDigitoVerif
        objPlanoConta.iDigitoVerifSimples = tPlanoConta.iDigitoVerifSimples
        objPlanoConta.iHistPadrao = tPlanoConta.iHistPadrao
        objPlanoConta.iAtivo = tPlanoConta.iAtivo
        objPlanoConta.iUsaContaSimples = tPlanoConta.iUsaContaSimples
        objPlanoConta.iCategoria = tPlanoConta.iCategoria
        objPlanoConta.iFluxoCaixa = tPlanoConta.iFluxoCaixa
        objPlanoConta.iNaturezaSped = tPlanoConta.iNaturezaSped
        
        'coloca a conta lida na coleção
        colPlanoConta.Add objPlanoConta
        
        'le a proxima conta
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 5910

    Loop

    Call Comando_Fechar(lComando)
    
    PlanoConta_Le_Todos = SUCESSO

    Exit Function

Erro_PlanoConta_Le_Todos:

    PlanoConta_Le_Todos = Err
    
    Select Case Err

        Case 5907
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 5908, 5909, 5910
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTA", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153390)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function HistPadrao_Le_Todos(ByVal colHistPadrao As Collection) As Long
'le todos os historicos contidos na tabela de historico padrão e coloca-os em colHistPadrao

Dim lComando As Long
Dim lErro As Long
Dim objHistPadrao As ClassHistPadrao
Dim iHistPadrao As Integer
Dim sDescHistPadrao As String
    
On Error GoTo Erro_HistPadrao_Le_Todos

    sDescHistPadrao = String(STRING_HISTORICO, 0)

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 5886

    'le a tabela de historico padrão
    lErro = Comando_Executar(lComando, "SELECT HistPadrao, DescHistPadrao FROM HistPadrao ORDER BY HistPadrao", iHistPadrao, sDescHistPadrao)
    If lErro <> AD_SQL_SUCESSO Then Error 5887

    'le o primeiro historico
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 5888
    
    Do While lErro <> AD_SQL_SEM_DADOS

        Set objHistPadrao = New ClassHistPadrao
        
        objHistPadrao.iHistPadrao = iHistPadrao
        objHistPadrao.sDescHistPadrao = sDescHistPadrao
        
        'coloca o histórico lido na coleção
        colHistPadrao.Add objHistPadrao
        
        'le o proximo historico
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 5890

    Loop

    Call Comando_Fechar(lComando)
    
    HistPadrao_Le_Todos = SUCESSO

    Exit Function

Erro_HistPadrao_Le_Todos:

    HistPadrao_Le_Todos = Err
    
    Select Case Err

        Case 5886
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 5887, 5888, 5890
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_HISTPADRAO1", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153391)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function Conta_SelecionaUma(ByVal sConta As String, ByVal objPlanoConta As ClassPlanoConta, ByVal sSiglaModulo As String) As Long
'O argumento sConta deve ser o número da conta FORMATADO BD
'le a conta passada como parametro e retorna os dados em objPlanoConta.

Dim lComando As Long
Dim lErro As Long
Dim tPlanoConta As typePlanoConta
Dim sSelecao As String

On Error GoTo Erro_Conta_SelecionaUma

    tPlanoConta.sConta = String(STRING_CONTA, 0)
    tPlanoConta.sDescConta = String(STRING_CONTA_DESCRICAO, 0)
    
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 6028
    
    If sSiglaModulo = MODULO_CONTABILIDADE Then

        lErro = Comando_Executar(lComando, "SELECT Conta, NivelConta, DescConta, TipoConta, ContaSimples, Natureza, DigitoVerif, DigitoVerifSimples, HistPadrao, Ativo, UsaContaSimples, Categoria, FluxoCaixa, NaturezaSped FROM PlanoConta WHERE Conta=?", tPlanoConta.sConta, tPlanoConta.iNivelConta, tPlanoConta.sDescConta, tPlanoConta.iTipoConta, tPlanoConta.lContaSimples, tPlanoConta.iNatureza, tPlanoConta.iDigitoVerif, tPlanoConta.iDigitoVerifSimples, tPlanoConta.iHistPadrao, tPlanoConta.iAtivo, tPlanoConta.iUsaContaSimples, tPlanoConta.iCategoria, tPlanoConta.iFluxoCaixa, tPlanoConta.iNaturezaSped, sConta)
        If lErro <> AD_SQL_SUCESSO Then Error 6029
        
    Else

        lErro = Comando_Executar(lComando, "SELECT Conta, NivelConta, DescConta, TipoConta, ContaSimples, Natureza, DigitoVerif, DigitoVerifSimples, HistPadrao, Ativo, UsaContaSimples, Categoria, FluxoCaixa, NaturezaSped FROM PlanoConta WHERE Conta=? AND " & sSiglaModulo & "=1", tPlanoConta.sConta, tPlanoConta.iNivelConta, tPlanoConta.sDescConta, tPlanoConta.iTipoConta, tPlanoConta.lContaSimples, tPlanoConta.iNatureza, tPlanoConta.iDigitoVerif, tPlanoConta.iDigitoVerifSimples, tPlanoConta.iHistPadrao, tPlanoConta.iAtivo, tPlanoConta.iUsaContaSimples, tPlanoConta.iCategoria, tPlanoConta.iFluxoCaixa, tPlanoConta.iNaturezaSped, sConta)
        If lErro <> AD_SQL_SUCESSO Then Error 10055
    
    End If
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 6031
        
    If lErro = AD_SQL_SEM_DADOS Then Error 6030
        
    objPlanoConta.sConta = tPlanoConta.sConta
    objPlanoConta.iNivelConta = tPlanoConta.iNivelConta
    objPlanoConta.sDescConta = tPlanoConta.sDescConta
    objPlanoConta.iTipoConta = tPlanoConta.iTipoConta
    objPlanoConta.lContaSimples = tPlanoConta.lContaSimples
    objPlanoConta.iNatureza = tPlanoConta.iNatureza
    objPlanoConta.iDigitoVerif = tPlanoConta.iDigitoVerif
    objPlanoConta.iDigitoVerifSimples = tPlanoConta.iDigitoVerifSimples
    objPlanoConta.iHistPadrao = tPlanoConta.iHistPadrao
    objPlanoConta.iAtivo = tPlanoConta.iAtivo
    objPlanoConta.iUsaContaSimples = tPlanoConta.iUsaContaSimples
    objPlanoConta.dSldIni = tPlanoConta.dSldIni
    objPlanoConta.iCategoria = tPlanoConta.iCategoria
    objPlanoConta.iFluxoCaixa = tPlanoConta.iFluxoCaixa
    objPlanoConta.iNaturezaSped = tPlanoConta.iNaturezaSped
    
    Call Comando_Fechar(lComando)
    Conta_SelecionaUma = SUCESSO
    
Exit Function
 
Erro_Conta_SelecionaUma:

    Conta_SelecionaUma = Err
    
    Select Case Err
    
    Case 6028
        lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
    Case 6029, 6031, 10055
        lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTA3", Err, sConta)
    
    Case 6030
                
    Case Else
        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153392)
    
    End Select
    
    Call Comando_Fechar(lComando)
      
    Exit Function

End Function

Function LotePendente_Le(ByVal objLote As ClassLote) As Long
'le um lote a partir da chave fornecida através de objLote. Devolve os dados em objLote

Dim lComando As Long
Dim lErro As Long
Dim tLote As typeLote

    
On Error GoTo Erro_LotePendente_Le

    tLote.sIdOriginal = String(STRING_IDORIGINAL, 0)

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 5433
    
    If giSetupLotePorPeriodo = LOTE_INICIALIZADO_POR_PERIODO Then
    
        lErro = Comando_Executar(lComando, "SELECT Periodo, TotCre, TotDeb, TotInf, Status, IdOriginal, NumLancInf, NumLancAtual, NumDocInf, NumDocAtual FROM LotePendente WHERE FilialEmpresa = ? AND Origem = ? AND Exercicio=? AND Periodo=? AND Lote=?", tLote.iPeriodo, tLote.dTotCre, tLote.dTotDeb, tLote.dTotInf, tLote.iStatus, tLote.sIdOriginal, tLote.iNumLancInf, tLote.iNumLancAtual, tLote.iNumDocInf, tLote.iNumDocAtual, objLote.iFilialEmpresa, objLote.sOrigem, objLote.iExercicio, objLote.iPeriodo, objLote.iLote)
        If lErro <> AD_SQL_SUCESSO Then Error 5434
        
    Else
        
        lErro = Comando_Executar(lComando, "SELECT Periodo, TotCre, TotDeb, TotInf, Status, IdOriginal, NumLancInf, NumLancAtual, NumDocInf, NumDocAtual FROM LotePendente WHERE FilialEmpresa = ? AND Origem = ? AND Exercicio=? AND Lote=?", tLote.iPeriodo, tLote.dTotCre, tLote.dTotDeb, tLote.dTotInf, tLote.iStatus, tLote.sIdOriginal, tLote.iNumLancInf, tLote.iNumLancAtual, tLote.iNumDocInf, tLote.iNumDocAtual, objLote.iFilialEmpresa, objLote.sOrigem, objLote.iExercicio, objLote.iLote)
        If lErro <> AD_SQL_SUCESSO Then Error 5827
        
    End If
    
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 5442
        
    If lErro = AD_SQL_SEM_DADOS Then Error 5435
    
    objLote.iPeriodo = tLote.iPeriodo
    objLote.dTotCre = tLote.dTotCre
    objLote.dTotDeb = tLote.dTotDeb
    objLote.dTotInf = tLote.dTotInf
    objLote.iStatus = tLote.iStatus
    objLote.sIdOriginal = tLote.sIdOriginal
    objLote.iNumLancInf = tLote.iNumLancInf
    objLote.iNumLancAtual = tLote.iNumLancAtual
    objLote.iNumDocInf = tLote.iNumDocInf
    objLote.iNumDocAtual = tLote.iNumDocAtual
    
    Call Comando_Fechar(lComando)
    
    LotePendente_Le = SUCESSO

Exit Function

Erro_LotePendente_Le:

    LotePendente_Le = Err
    
    Select Case Err
    
        Case 5433
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 5434, 5442, 5827
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LOTEPENDENTE", Err, objLote.iFilialEmpresa, objLote.sOrigem, objLote.iExercicio, objLote.iPeriodo, objLote.iLote)
           
        Case 5435
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153393)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function HistPadrao_Le(ByVal objHistPadrao As ClassHistPadrao) As Long
'Le o historico padrao a partir da chave fornecida em objHistPadrao. Retorna os dados em objHistPadrao

Dim lComando As Long
Dim lErro As Long
Dim sDescHistPadrao As String

    
On Error GoTo Erro_HistPadrao_Le

    sDescHistPadrao = String(STRING_HISTORICO, 0)

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 5444
    
    'seleciona o historico cujo codigo foi passado como parametro
    lErro = Comando_Executar(lComando, "SELECT DescHistPadrao FROM HistPadrao WHERE HistPadrao = ?", sDescHistPadrao, objHistPadrao.iHistPadrao)
    If lErro <> AD_SQL_SUCESSO Then Error 5445
    
    'le o historico
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 5447
        
    'se não encontrou o historico ==> erro
    If lErro = AD_SQL_SEM_DADOS Then Error 5446
    
    objHistPadrao.sDescHistPadrao = sDescHistPadrao
    
    Call Comando_Fechar(lComando)
    
    HistPadrao_Le = SUCESSO

    Exit Function

Erro_HistPadrao_Le:

    HistPadrao_Le = Err
    
    Select Case Err
    
        Case 5444
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 5445, 5447
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_HISTPADRAO", Err, objHistPadrao.iHistPadrao)
        
        Case 5446
           
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153394)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function PlanoConta_Le_ContaSimples(ByVal lContaSimples As Long, ByVal objPlanoConta As ClassPlanoConta) As Long
'le uma conta a partir da conta simplificada fornecida em lContaSimples. Retorna os dados em objPlanoConta

Dim lComando As Long
Dim lErro As Long
Dim tPlanoConta As typePlanoConta

On Error GoTo Erro_PlanoConta_Le_ContaSimples

    tPlanoConta.sConta = String(STRING_CONTA, 0)
    tPlanoConta.sDescConta = String(STRING_CONTA_DESCRICAO, 0)

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 5449

    lErro = Comando_Executar(lComando, "SELECT Conta, NivelConta, DescConta, TipoConta, ContaSimples, Natureza, DigitoVerif, DigitoVerifSimples, HistPadrao, Ativo, UsaContaSimples, Categoria, CP, CR, TES, FAT, EST, FluxoCaixa FROM PlanoConta WHERE ContaSimples=?", tPlanoConta.sConta, tPlanoConta.iNivelConta, tPlanoConta.sDescConta, tPlanoConta.iTipoConta, tPlanoConta.lContaSimples, tPlanoConta.iNatureza, tPlanoConta.iDigitoVerif, tPlanoConta.iDigitoVerifSimples, tPlanoConta.iHistPadrao, tPlanoConta.iAtivo, tPlanoConta.iUsaContaSimples, tPlanoConta.iCategoria, tPlanoConta.iCP, tPlanoConta.iCR, tPlanoConta.iTES, tPlanoConta.iFAT, tPlanoConta.iEST, tPlanoConta.iFluxoCaixa, lContaSimples)
    If lErro <> AD_SQL_SUCESSO Then Error 5450
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 5452
        
    If lErro = AD_SQL_SEM_DADOS Then Error 5451
        
    objPlanoConta.sConta = tPlanoConta.sConta
    objPlanoConta.iNivelConta = tPlanoConta.iNivelConta
    objPlanoConta.sDescConta = tPlanoConta.sDescConta
    objPlanoConta.iTipoConta = tPlanoConta.iTipoConta
    objPlanoConta.lContaSimples = tPlanoConta.lContaSimples
    objPlanoConta.iNatureza = tPlanoConta.iNatureza
    objPlanoConta.iDigitoVerif = tPlanoConta.iDigitoVerif
    objPlanoConta.iDigitoVerifSimples = tPlanoConta.iDigitoVerifSimples
    objPlanoConta.iHistPadrao = tPlanoConta.iHistPadrao
    objPlanoConta.iAtivo = tPlanoConta.iAtivo
    objPlanoConta.iUsaContaSimples = tPlanoConta.iUsaContaSimples
    objPlanoConta.dSldIni = tPlanoConta.dSldIni
    objPlanoConta.iCategoria = tPlanoConta.iCategoria
    objPlanoConta.iFluxoCaixa = tPlanoConta.iFluxoCaixa
    
    If tPlanoConta.iCP = USA_MODULO Then objPlanoConta.colModulo.Add MODULO_CONTASAPAGAR
    If tPlanoConta.iCR = USA_MODULO Then objPlanoConta.colModulo.Add MODULO_CONTASARECEBER
    If tPlanoConta.iTES = USA_MODULO Then objPlanoConta.colModulo.Add MODULO_TESOURARIA
    If tPlanoConta.iFAT = USA_MODULO Then objPlanoConta.colModulo.Add MODULO_FATURAMENTO
    If tPlanoConta.iEST = USA_MODULO Then objPlanoConta.colModulo.Add MODULO_ESTOQUE
        
    Call Comando_Fechar(lComando)
    
    PlanoConta_Le_ContaSimples = SUCESSO
    
    Exit Function
 
Erro_PlanoConta_Le_ContaSimples:

    PlanoConta_Le_ContaSimples = Err
    
    Select Case Err
    
        Case 5449
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 5450, 5452
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTA2", Err, lContaSimples)
    
        Case 5451
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153395)
    
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function PlanoContas_Le_HistPadrao(ByVal iHistPadrao As Integer) As Long
'Le o historico padrao no plano de contas.

Dim lComando As Long
Dim lErro As Long
Dim iHistPadrao1 As Integer
    
On Error GoTo Erro_PlanoContas_Le_HistPadrao

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 6163
    
    lErro = Comando_Executar(lComando, "SELECT HistPadrao FROM PlanoConta WHERE HistPadrao = ?", iHistPadrao1, iHistPadrao)
    If lErro <> AD_SQL_SUCESSO Then Error 6164
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 6166
        
    If lErro = AD_SQL_SEM_DADOS Then Error 6165
    
    Call Comando_Fechar(lComando)
    
    PlanoContas_Le_HistPadrao = SUCESSO

    Exit Function

Erro_PlanoContas_Le_HistPadrao:

    PlanoContas_Le_HistPadrao = Err
    
    Select Case Err
    
        Case 6163
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 6164, 6166
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTA", Err)

        Case 6165
   
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153396)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function Lancamento_Le_Conta(ByVal sConta As String) As Long
'verifica se existe no BD pelo menos um lancamento para a conta fornecida

Dim lComando As Long
Dim lErro As Long
Dim lDoc As Long

On Error GoTo Erro_Lancamento_Le_Conta

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 5455
    
    lErro = Comando_Executar(lComando, "SELECT Doc FROM Lancamentos WHERE Conta=?", lDoc, sConta)
    If lErro <> AD_SQL_SUCESSO Then Error 5456
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 5458
    
    If lErro = AD_SQL_SEM_DADOS Then Error 5457
        
    Call Comando_Fechar(lComando)
    
    Lancamento_Le_Conta = SUCESSO
    
    Exit Function
 
Erro_Lancamento_Le_Conta:

    Lancamento_Le_Conta = Err
    
    Select Case Err
    
        Case 5455
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 5456, 5458
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LANCAMENTOS2", Err, sConta)
            
        Case 5457
               
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153397)
    
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function LanPendente_Le_Conta(ByVal sConta As String) As Long
'verifica se existe no BD pelo menos um lancamento pendente para a conta fornecida

Dim lComando As Long
Dim lErro As Long
Dim lDoc As Long

On Error GoTo Erro_LanPendente_Le_Conta

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 5459
    
    lErro = Comando_Executar(lComando, "SELECT Doc FROM LanPendente WHERE Conta=?", lDoc, sConta)
    If lErro <> AD_SQL_SUCESSO Then Error 5460
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 5462
        
    If lErro = AD_SQL_SEM_DADOS Then Error 5461
        
    Call Comando_Fechar(lComando)
    
    LanPendente_Le_Conta = SUCESSO
    
    Exit Function
 
Erro_LanPendente_Le_Conta:

    LanPendente_Le_Conta = Err
    
    Select Case Err
    
        Case 5459
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 5460, 5462
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LANPENDENTE", Err, sConta)
            
        Case 5461
          
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153398)
    
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function ContaCcl_Le_Conta(ByVal sConta As String) As Long
'verifica se existe no BD pelo menos uma associacao da conta fornecida com centros de custo

Dim lComando As Long
Dim lErro As Long
Dim sCcl As String

On Error GoTo Erro_ContaCcl_Le_Conta

    sCcl = String(STRING_CCL, 0)
    
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 5467
    
    lErro = Comando_Executar(lComando, "SELECT Ccl FROM ContaCcl WHERE Conta=?", sCcl, sConta)
    If lErro <> AD_SQL_SUCESSO Then Error 5468
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 5470
        
    If lErro = AD_SQL_SEM_DADOS Then Error 5469
        
    Call Comando_Fechar(lComando)
    
    ContaCcl_Le_Conta = SUCESSO
    
    Exit Function
 
Erro_ContaCcl_Le_Conta:

    ContaCcl_Le_Conta = Err
    
    Select Case Err
    
        Case 5467
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 5468, 5470
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTACCL1", Err, sConta)
           
        Case 5469
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153399)
    
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function PlanoConta_Le_ContaMaior(ByVal sConta As String, ByVal objPlanoConta As ClassPlanoConta) As Long
'le uma conta imediatamente maior que a conta passada como parametro. Devolve os dados em objPlanoConta

Dim lComando As Long
Dim lErro As Long
Dim tPlanoConta As typePlanoConta

On Error GoTo Erro_PlanoConta_Le_ContaMaior

    tPlanoConta.sConta = String(STRING_CONTA, 0)
    tPlanoConta.sDescConta = String(STRING_CONTA_DESCRICAO, 0)
    
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 5473

    lErro = Comando_Executar(lComando, "SELECT Conta, NivelConta, DescConta, TipoConta, ContaSimples, Natureza, DigitoVerif, DigitoVerifSimples, HistPadrao, Ativo, UsaContaSimples, Categoria, FluxoCaixa FROM PlanoConta WHERE Conta>? ORDER BY Conta", tPlanoConta.sConta, tPlanoConta.iNivelConta, tPlanoConta.sDescConta, tPlanoConta.iTipoConta, tPlanoConta.lContaSimples, tPlanoConta.iNatureza, tPlanoConta.iDigitoVerif, tPlanoConta.iDigitoVerifSimples, tPlanoConta.iHistPadrao, tPlanoConta.iAtivo, tPlanoConta.iUsaContaSimples, tPlanoConta.iCategoria, tPlanoConta.iFluxoCaixa, sConta)
    If lErro <> AD_SQL_SUCESSO Then Error 5474
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 5476
        
    If lErro = AD_SQL_SEM_DADOS Then Error 5475
        
    objPlanoConta.sConta = tPlanoConta.sConta
    objPlanoConta.iNivelConta = tPlanoConta.iNivelConta
    objPlanoConta.sDescConta = tPlanoConta.sDescConta
    objPlanoConta.iTipoConta = tPlanoConta.iTipoConta
    objPlanoConta.lContaSimples = tPlanoConta.lContaSimples
    objPlanoConta.iNatureza = tPlanoConta.iNatureza
    objPlanoConta.iDigitoVerif = tPlanoConta.iDigitoVerif
    objPlanoConta.iDigitoVerifSimples = tPlanoConta.iDigitoVerifSimples
    objPlanoConta.iHistPadrao = tPlanoConta.iHistPadrao
    objPlanoConta.iAtivo = tPlanoConta.iAtivo
    objPlanoConta.iUsaContaSimples = tPlanoConta.iUsaContaSimples
    objPlanoConta.iCategoria = tPlanoConta.iCategoria
    objPlanoConta.iFluxoCaixa = tPlanoConta.iFluxoCaixa
    
    Call Comando_Fechar(lComando)
    
    PlanoConta_Le_ContaMaior = SUCESSO
    
    Exit Function
 
Erro_PlanoConta_Le_ContaMaior:

    PlanoConta_Le_ContaMaior = Err
    
    Select Case Err
    
        Case 5473
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 5474, 5476
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTA", Err)
              
        Case 5475
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153400)
    
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function Ccl_Le_CclMaior(ByVal sCcl As String, ByVal objCcl As ClassCcl) As Long
'le um centro de custo/lucro imediatamente maior que o centro de custo/lucro passado como parametro. Devolve os dados em objCcl

Dim lComando As Long
Dim lErro As Long
Dim tCcl As typeCcl

On Error GoTo Erro_Ccl_Le_CclMaior

    tCcl.sCcl = String(STRING_CCL, 0)
    tCcl.sDescCcl = String(STRING_CCL_DESCRICAO, 0)
    
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 10369
    
    lErro = Comando_Executar(lComando, "SELECT Ccl, DescCcl, TipoCcl FROM Ccl WHERE Ccl>? ORDER BY Ccl", tCcl.sCcl, tCcl.sDescCcl, tCcl.iTipoCcl, sCcl)
    If lErro <> AD_SQL_SUCESSO Then Error 10370
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 10371
        
    If lErro = AD_SQL_SEM_DADOS Then Error 10372
        
    objCcl.sCcl = tCcl.sCcl
    objCcl.sDescCcl = tCcl.sDescCcl
    objCcl.iTipoCcl = tCcl.iTipoCcl
        
    Call Comando_Fechar(lComando)
    
    Ccl_Le_CclMaior = SUCESSO
    
    Exit Function
 
Erro_Ccl_Le_CclMaior:

    Ccl_Le_CclMaior = Err
    
    Select Case Err
    
        Case 10369
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 10370, 10371
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CCL1", Err)
              
        Case 10372
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153401)
    
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function Lancamento_Le_Lote(ByVal objLote As ClassLote) As Long
'verifica se existe no BD pelo menos um lancamento para o lote fornecido

Dim lComando As Long
Dim lErro As Long
Dim lDoc As Long

On Error GoTo Erro_Lancamento_Le_Lote

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 5518
    
    lErro = Comando_Executar(lComando, "SELECT Doc FROM Lancamentos WHERE Origem=? AND Exercicio=? AND PeriodoLote=? AND Lote=?", lDoc, objLote.sOrigem, objLote.iExercicio, objLote.iPeriodo, objLote.iLote)
    If lErro <> AD_SQL_SUCESSO Then Error 5519
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 5521
        
    If lErro = AD_SQL_SEM_DADOS Then Error 5520
        
    Call Comando_Fechar(lComando)
    
    Lancamento_Le_Lote = SUCESSO
    
    Exit Function
 
Erro_Lancamento_Le_Lote:

    Lancamento_Le_Lote = Err
    
    Select Case Err
    
        Case 5518
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 5519, 5521
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LANCAMENTOS", Err, objLote.sOrigem, objLote.iExercicio, objLote.iPeriodo, objLote.iLote)
            
        Case 5520
           
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153402)
    
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function Lancamento_Le_Doc(ByVal objLancamento_Cabecalho As ClassLancamento_Cabecalho) As Long
'Tenta ler o documento em questão.

Dim lComando1 As Long
Dim lErro As Long
Dim lDoc1 As Long

On Error GoTo Erro_Lancamento_Le_Doc

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 5727

    'verifica se este número ainda não foi utilizado
    If giSetupDocPorPeriodo = DOC_INICIALIZADO_POR_PERIODO Then
    
        lErro = Comando_Executar(lComando1, "SELECT Doc FROM Lancamentos WHERE FilialEmpresa = ? AND Origem = ? AND Exercicio = ? AND PeriodoLan = ? AND Doc = ? ORDER BY Doc", lDoc1, objLancamento_Cabecalho.iFilialEmpresa, objLancamento_Cabecalho.sOrigem, objLancamento_Cabecalho.iExercicio, objLancamento_Cabecalho.iPeriodoLan, objLancamento_Cabecalho.lDoc)
        If lErro <> AD_SQL_SUCESSO Then Error 5728
        
    Else
    
        lErro = Comando_Executar(lComando1, "SELECT Doc FROM Lancamentos WHERE FilialEmpresa = ? AND Origem = ? AND Exercicio = ? AND Doc = ? ORDER BY Doc", lDoc1, objLancamento_Cabecalho.iFilialEmpresa, objLancamento_Cabecalho.sOrigem, objLancamento_Cabecalho.iExercicio, objLancamento_Cabecalho.lDoc)
        If lErro <> AD_SQL_SUCESSO Then Error 5729
        
    End If
        
    'le o primeiro lancamento
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 5725

    If lErro = AD_SQL_SEM_DADOS Then Error 5726

    Call Comando_Fechar(lComando1)
    
    Lancamento_Le_Doc = SUCESSO
        
    Exit Function
    
    
Erro_Lancamento_Le_Doc:

    Lancamento_Le_Doc = Err
    
    Select Case Err
    
        Case 5725, 5728, 5729
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LANCAMENTOS3", Err)
           
        Case 5726
               
        Case 5727
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153403)
        
    End Select
    
    Call Comando_Fechar(lComando1)

    Exit Function
    
End Function

Function Lancamento_Le_Doc1(ByVal objLancamento_Cabecalho As ClassLancamento_Cabecalho) As Long
'Tenta ler o documento em questão e verifica se é um documento estornado ou estornador

Dim lComando1 As Long
Dim lErro As Long
Dim lDoc1 As Long
Dim iStatus As Integer

On Error GoTo Erro_Lancamento_Le_Doc1

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 83859

    'verifica se este número ainda não foi utilizado
    If giSetupDocPorPeriodo = DOC_INICIALIZADO_POR_PERIODO Then
    
        lErro = Comando_Executar(lComando1, "SELECT Status FROM Lancamentos WHERE FilialEmpresa = ? AND Origem = ? AND Exercicio = ? AND PeriodoLan = ? AND Doc = ? AND Data = ? ORDER BY Doc", lDoc1, objLancamento_Cabecalho.iFilialEmpresa, objLancamento_Cabecalho.sOrigem, objLancamento_Cabecalho.iExercicio, objLancamento_Cabecalho.iPeriodoLan, objLancamento_Cabecalho.lDoc, objLancamento_Cabecalho.dtData)
        If lErro <> AD_SQL_SUCESSO Then gError 83860
        
    Else
    
        lErro = Comando_Executar(lComando1, "SELECT Status FROM Lancamentos WHERE FilialEmpresa = ? AND Origem = ? AND Exercicio = ? AND Doc = ? AND Data = ? ORDER BY Doc", lDoc1, objLancamento_Cabecalho.iFilialEmpresa, objLancamento_Cabecalho.sOrigem, objLancamento_Cabecalho.iExercicio, objLancamento_Cabecalho.lDoc, objLancamento_Cabecalho.dtData)
        If lErro <> AD_SQL_SUCESSO Then gError 83861
        
    End If
        
    'le o primeiro lancamento
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 83862

    If lErro = AD_SQL_SEM_DADOS Then gError 83863

    If iStatus = VOUCHER_ESTORNADO Then gError 83864

    If iStatus = VOUCHER_ESTORNADOR Then gError 83865

    Call Comando_Fechar(lComando1)
    
    Lancamento_Le_Doc1 = SUCESSO
        
    Exit Function
    
Erro_Lancamento_Le_Doc1:

    Lancamento_Le_Doc1 = gErr
    
    Select Case gErr
    
        Case 83859
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 83860, 83861, 83862
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LANCAMENTOS3", gErr)
           
        Case 83863, 83864, 83865
               
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153404)
        
    End Select
    
    Call Comando_Fechar(lComando1)

    Exit Function
    
End Function


Function LanPendente_Le_Lote(ByVal objLote As ClassLote) As Long
'verifica se existe no BD pelo menos um lancamento pendente para o lote fornecido

Dim lComando As Long
Dim lErro As Long
Dim lDoc As Long

On Error GoTo Erro_LanPendente_Le_Lote

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 5522
    
    lErro = Comando_Executar(lComando, "SELECT Doc FROM LanPendente WHERE FilialEmpresa=? AND Origem=? AND Exercicio=? AND PeriodoLote=? AND Lote=?", lDoc, objLote.iFilialEmpresa, objLote.sOrigem, objLote.iExercicio, objLote.iPeriodo, objLote.iLote)
    If lErro <> AD_SQL_SUCESSO Then Error 5523
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 5525
        
    If lErro = AD_SQL_SEM_DADOS Then Error 5524
        
    Call Comando_Fechar(lComando)
    
    LanPendente_Le_Lote = SUCESSO
    
    Exit Function
 
Erro_LanPendente_Le_Lote:

    LanPendente_Le_Lote = Err
    
    Select Case Err
    
        Case 5522
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 5523, 5525
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LANPENDENTE1", Err, objLote.iFilialEmpresa, objLote.sOrigem, objLote.iExercicio, objLote.iPeriodo, objLote.iLote)
            
        Case 5524
          
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153405)
    
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function RateioOn_Le_Conta(ByVal sConta As String) As Long
'verifica se existe na tabela RateioOn pelo menos um rateio com a conta em questão.

Dim lComando As Long
Dim lErro As Long
Dim iCodigo As Integer
Dim iSeq As Integer

On Error GoTo Erro_RateioOn_Le_Conta

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 5564
    
    lErro = Comando_Executar(lComando, "SELECT Codigo, Seq FROM RateioOn WHERE Conta=?", iCodigo, iSeq, sConta)
    If lErro <> AD_SQL_SUCESSO Then Error 5565
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 5567
        
    If lErro = AD_SQL_SEM_DADOS Then Error 5566
        
    Call Comando_Fechar(lComando)
    
    RateioOn_Le_Conta = SUCESSO
    
    Exit Function
 
Erro_RateioOn_Le_Conta:

    RateioOn_Le_Conta = Err
    
    Select Case Err
    
        Case 5564
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 5565, 5567
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RATEIOON", Err, sConta)
           
        Case 5566
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153406)
    
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function RateioOn_Le_Doc(ByVal objRateioOn As ClassRateioOn, ByVal colRateioOns As Collection) As Long
'le os lançamentos relativos ao Rateio passado como parametro e coloca-os em colRateioOns

Dim lComando As Long
Dim lErro As Long
Dim tRateioOn As typeRateioOn
Dim objRateioOn1 As ClassRateioOn

On Error GoTo Erro_RateioOn_Le_Doc

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 11137

    tRateioOn.sConta = String(STRING_CONTA, 0)
    tRateioOn.sCcl = String(STRING_CCL, 0)
    tRateioOn.sHistorico = String(STRING_HISTORICO, 0)
    tRateioOn.sDescricao = String(STRING_RATEIO_DESCRICAO, 0)

    lErro = Comando_Executar(lComando, "SELECT  Seq, Descricao, Conta,  Ccl,  Historico, Percentual FROM RateioOn WHERE Codigo=?", tRateioOn.iSeq, tRateioOn.sDescricao, tRateioOn.sConta, tRateioOn.sCcl, tRateioOn.sHistorico, tRateioOn.dPercentual, objRateioOn.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 11138

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 11139

    If lErro = AD_SQL_SEM_DADOS Then Error 11136

    Do While lErro <> AD_SQL_SEM_DADOS

        Set objRateioOn1 = New ClassRateioOn

        objRateioOn1.iCodigo = objRateioOn.iCodigo
        objRateioOn1.iSeq = tRateioOn.iSeq
        objRateioOn1.sConta = tRateioOn.sConta
        objRateioOn1.sCcl = tRateioOn.sCcl
        objRateioOn1.sHistorico = tRateioOn.sHistorico
        objRateioOn1.dPercentual = tRateioOn.dPercentual
        objRateioOn1.sDescricao = tRateioOn.sDescricao

        colRateioOns.Add objRateioOn1

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 11140
        
    Loop
    
    Call Comando_Fechar(lComando)
    
    RateioOn_Le_Doc = SUCESSO
    
    Exit Function
 
Erro_RateioOn_Le_Doc:

    RateioOn_Le_Doc = Err
    
    Select Case Err
    
        Case 11136
           
        Case 11137
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
              
        Case 11138, 11139, 11140
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RATEIOON", Err)
                 
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153407)
    
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function RateioOff_Le_Conta(ByVal sConta As String) As Long
'verifica se existe na tabela RateioOff pelo menos um rateio com a conta em questão.

Dim lComando As Long
Dim lErro As Long
Dim lSeq As Long


On Error GoTo Erro_RateioOff_Le_Conta

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 5572
    
    lErro = Comando_Executar(lComando, "SELECT Seq FROM RateioOff WHERE Conta=? OR ContaCre = ?", lSeq, sConta, sConta)
    If lErro <> AD_SQL_SUCESSO Then Error 5573
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 5575
        
    If lErro = AD_SQL_SEM_DADOS Then Error 5574
        
    Call Comando_Fechar(lComando)
    
    RateioOff_Le_Conta = SUCESSO
    
    Exit Function
 
Erro_RateioOff_Le_Conta:

    RateioOff_Le_Conta = Err
    
    Select Case Err
    
        Case 5572
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 5573, 5575
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RATEIOOFF", Err, sConta)
               
        Case 5574
                  
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153408)
    
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function RateioOff_Le_Doc(ByVal objRateioOff As ClassRateioOff, ByVal colRateioOff As Collection) As Long
'le os lançamentos relativos ao Rateio passado como parametro e coloca-os em colRateioOff

Dim lComando As Long
Dim lErro As Long
Dim tRateioOff As typeRateioOff
Dim objRateioOff1 As ClassRateioOff


On Error GoTo Erro_RateioOff_Le_Doc
    
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 11342

    tRateioOff.sContaCre = String(STRING_CONTA, 0)
    tRateioOff.sConta = String(STRING_CONTA, 0)
    tRateioOff.sCclOrigem = String(STRING_CCL, 0)
    tRateioOff.sCcl = String(STRING_CCL, 0)
    tRateioOff.sDescricao = String(STRING_RATEIO_DESCRICAO, 0)
    
    lErro = Comando_Executar(lComando, "SELECT  CclOrigem, Descricao, Percentual, Seq, Conta, Ccl, ContaCre, Tipo  FROM RateioOff WHERE Codigo = ?", tRateioOff.sCclOrigem, tRateioOff.sDescricao, tRateioOff.dPercentual, tRateioOff.lSeq, tRateioOff.sConta, tRateioOff.sCcl, tRateioOff.sContaCre, tRateioOff.iTipo, objRateioOff.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 11343

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 11344

    If lErro = AD_SQL_SEM_DADOS Then Error 11345

    objRateioOff.sCclOrigem = tRateioOff.sCclOrigem
    objRateioOff.iTipo = tRateioOff.iTipo
    objRateioOff.sContaCre = tRateioOff.sContaCre
    objRateioOff.sDescricao = tRateioOff.sDescricao

    Do While lErro <> AD_SQL_SEM_DADOS

        Set objRateioOff1 = New ClassRateioOff

        objRateioOff1.lCodigo = objRateioOff.lCodigo
        objRateioOff1.sCclOrigem = tRateioOff.sCclOrigem
        objRateioOff1.sDescricao = tRateioOff.sDescricao
        objRateioOff1.sConta = tRateioOff.sConta
        objRateioOff1.sCcl = tRateioOff.sCcl
        objRateioOff1.sContaCre = tRateioOff.sContaCre
        objRateioOff1.lSeq = tRateioOff.lSeq
        objRateioOff1.dPercentual = tRateioOff.dPercentual
        objRateioOff1.iTipo = tRateioOff.iTipo

        colRateioOff.Add objRateioOff1

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 11346
        
    Loop
    
    Call Comando_Fechar(lComando)
    
    RateioOff_Le_Doc = SUCESSO
    
    Exit Function
 
Erro_RateioOff_Le_Doc:

    RateioOff_Le_Doc = Err

    Select Case Err
    
        Case 11342
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 11343, 11344, 11346
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RATEIOOFF", Err)
        
        Case 11345
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153409)
    
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function DocAuto_Le_Conta(ByVal sConta As String) As Long
'verifica se existe na tabela DocAuto pelo menos um documento automático com a conta em questão.

Dim lComando As Long
Dim lErro As Long
Dim lDoc As Long
Dim iSeq As Integer

On Error GoTo Erro_DocAuto_Le_Conta

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 5576
    
    lErro = Comando_Executar(lComando, "SELECT Doc, Seq FROM DocAuto WHERE Conta=?", lDoc, iSeq, sConta)
    If lErro <> AD_SQL_SUCESSO Then Error 5577
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 5579
        
    If lErro = AD_SQL_SEM_DADOS Then Error 5578
        
    Call Comando_Fechar(lComando)
    
    DocAuto_Le_Conta = SUCESSO
    
    Exit Function
 
Erro_DocAuto_Le_Conta:

    DocAuto_Le_Conta = Err
    
    Select Case Err
    
        Case 5576
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 5577, 5579
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_DOCAUTO", Err, sConta)
              
        Case 5578
                  
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153410)
    
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function DocAuto_Le_Doc(ByVal objDocAuto As ClassDocAuto, ByVal colDocAutos As Collection) As Long
'le os lançamentos relativos ao documento passado como parametro e coloca-os em colDocAutos

Dim lComando As Long
Dim lErro As Long
Dim tDocAuto As typeDocAuto
Dim objDocAuto1 As ClassDocAuto

On Error GoTo Erro_DocAuto_Le_Doc

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 11028

    tDocAuto.sOrigem = String(STRING_ORIGEM, 0)
    tDocAuto.sConta = String(STRING_CONTA, 0)
    tDocAuto.sCcl = String(STRING_CCL, 0)
    tDocAuto.sHistorico = String(STRING_HISTORICO, 0)
    tDocAuto.sDescricao = String(STRING_DOCAUTO_DESCRICAO, 0)

    lErro = Comando_Executar(lComando, "SELECT  Seq, Origem, Descricao, Conta,  Ccl,  Historico, Valor, SeqContraPartida FROM DocAuto WHERE Doc=? ORDER BY SEQ", tDocAuto.iSeq, tDocAuto.sOrigem, tDocAuto.sDescricao, tDocAuto.sConta, tDocAuto.sCcl, tDocAuto.sHistorico, tDocAuto.dValor, tDocAuto.iSeqContraPartida, objDocAuto.lDoc)
    If lErro <> AD_SQL_SUCESSO Then Error 11029

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 11030

    If lErro = AD_SQL_SEM_DADOS Then Error 11017

    Do While lErro <> AD_SQL_SEM_DADOS

        Set objDocAuto1 = New ClassDocAuto

        objDocAuto1.lDoc = objDocAuto.lDoc
        objDocAuto1.sOrigem = tDocAuto.sOrigem
        objDocAuto1.iSeq = tDocAuto.iSeq
        objDocAuto1.sConta = tDocAuto.sConta
        objDocAuto1.sCcl = tDocAuto.sCcl
        objDocAuto1.sHistorico = tDocAuto.sHistorico
        objDocAuto1.dValor = tDocAuto.dValor
        objDocAuto1.sDescricao = tDocAuto.sDescricao
        objDocAuto1.iSeqContraPartida = tDocAuto.iSeqContraPartida

        colDocAutos.Add objDocAuto1

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 11031
        
    Loop
    
    Call Comando_Fechar(lComando)
    
    DocAuto_Le_Doc = SUCESSO
    
    Exit Function
 
Erro_DocAuto_Le_Doc:

    DocAuto_Le_Doc = Err
    
    Select Case Err
    
        Case 11017
         
        Case 11028
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
              
        Case 11029, 11030, 11031
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_DOCAUTO", Err)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153411)
    
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

'Le Todos os dados dos Documentos Para fazer o carregamento da Arvore de Documento Automatico
Function DocAuto_Le_Todos(ByVal colDocAuto As Collection) As Long

Dim lComando As Long
Dim lErro As Long
Dim objDocAuto As ClassDocAuto
Dim tDocAuto As typeDocAuto

On Error GoTo Erro_DocAuto_Le_Todos
        
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 11354
    
    tDocAuto.sDescricao = String(STRING_DOCAUTO_DESCRICAO, 0)
    
    'le a tabela de Documento Automatico ordenado por Documento
    lErro = Comando_Executar(lComando, "SELECT DISTINCT Doc, Descricao FROM DocAuto ORDER BY Doc", tDocAuto.lDoc, tDocAuto.sDescricao)
    If lErro <> AD_SQL_SUCESSO Then Error 11355
    
    'le o primeiro Documento Automatico
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 11356
    
    Do While lErro <> AD_SQL_SEM_DADOS

        Set objDocAuto = New ClassDocAuto
        
        objDocAuto.lDoc = tDocAuto.lDoc
        objDocAuto.sDescricao = tDocAuto.sDescricao
        
        'coloca a conta lida na coleção
        colDocAuto.Add objDocAuto
        
        'le a proxima conta
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 11357

    Loop

    Call Comando_Fechar(lComando)
    
    DocAuto_Le_Todos = SUCESSO

    Exit Function

Erro_DocAuto_Le_Todos:

    DocAuto_Le_Todos = Err

    Select Case Err

        Case 11354
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 11355, 11356, 11357
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_DOCAUTO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153412)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function PlanoConta_Tem_Filho(ByVal sConta As String, iTemFilho As Integer) As Long
'testa se a conta em questão tem contas abaixo dela, ou seja tem contas filha.

Dim objPlanoConta As New ClassPlanoConta
Dim lErro As Long
Dim iNivel As Integer
Dim iNivel1 As Integer


On Error GoTo Erro_PlanoConta_Tem_Filho

    
    'le a conta com número imediatamente maior que a conta em questão
    lErro = CF("PlanoConta_Le_ContaMaior", sConta, objPlanoConta)
    If lErro <> SUCESSO And lErro <> 5475 Then Error 5591
    
    If lErro = SUCESSO Then
    
        'obtem o nivel da conta em questão
        lErro = Mascara_Conta_ObterNivel(sConta, iNivel)
        If lErro <> SUCESSO Then Error 5592
        
        'obtem o nivel da conta lida
        lErro = Mascara_Conta_ObterNivel(objPlanoConta.sConta, iNivel1)
        If lErro <> SUCESSO Then Error 5593
    
        'se o nivel da conta em questão for menor que a conta lida ==> tem filho
        If iNivel < iNivel1 Then
            iTemFilho = 1
        Else
            iTemFilho = 0
        End If
        
    Else
         iTemFilho = 0
    End If
    
    PlanoConta_Tem_Filho = SUCESSO
    
    Exit Function
 
Erro_PlanoConta_Tem_Filho:

    PlanoConta_Tem_Filho = Err

    Select Case Err
    
        Case 5591
    
        Case 5592
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_CONTA_OBTERNIVEL", Err, sConta)
        
        Case 5593
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_CONTA_OBTERNIVEL", Err, objPlanoConta.sConta)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153413)
    
    End Select
    
    Exit Function
    
    
End Function

Function Ccl_Tem_Filho(ByVal sCcl As String, iTemFilho As Integer) As Long
'testa se o centro de custo em questão tem centros de custo/lucro filhos.

Dim objCcl As New ClassCcl
Dim lErro As Long
Dim iNivel As Integer
Dim iNivel1 As Integer


On Error GoTo Erro_Ccl_Tem_Filho

    'le o centro de custo/lucro com número imediatamente maior que o centro de custo/lucro em questão
    lErro = Ccl_Le_CclMaior(sCcl, objCcl)
    If lErro <> SUCESSO And lErro <> 10372 Then Error 10373
    
    If lErro = SUCESSO Then
    
        'obtem o nivel do centro de custo/lucro em questão
        lErro = Mascara_Ccl_ObterNivel(sCcl, iNivel)
        If lErro <> SUCESSO Then Error 10374
        
        'obtem o nivel do centro de custo/lucro lido
        lErro = Mascara_Ccl_ObterNivel(objCcl.sCcl, iNivel1)
        If lErro <> SUCESSO Then Error 10375
    
        'se o nivel do centro de custo/lucro em questão for menor que o centro de custo/lucro lido ==> tem filho
        If iNivel < iNivel1 Then
            iTemFilho = CCL_TEM_FILHOS
        Else
            iTemFilho = CCL_NAO_TEM_FILHOS
        End If
        
    Else
         iTemFilho = CCL_NAO_TEM_FILHOS
    End If
    
    Ccl_Tem_Filho = SUCESSO
    
    Exit Function
 
Erro_Ccl_Tem_Filho:

    Ccl_Tem_Filho = Err

    Select Case Err
    
        Case 10373
    
        Case 10374
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_CCL_OBTERNIVEL", Err, sCcl)
        
        Case 10375
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_CCL_OBTERNIVEL", Err, objCcl.sCcl)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153414)
    
    End Select
    
    Exit Function
    
    
End Function

Function Ccl_Le(ByVal objCcl As ClassCcl) As Long
'le o ccl a partir da chave fornecida em objCcl. Devolve os dados em objCcl

Dim lComando As Long
Dim lErro As Long
Dim tCcl As typeCcl

On Error GoTo Erro_Ccl_Le

    tCcl.sDescCcl = String(STRING_CCL_DESCRICAO, 0)

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 5597

    'Verifica se o Ccl existe
    lErro = Comando_Executar(lComando, "SELECT DescCcl, TipoCcl, AtivoCcl FROM Ccl WHERE Ccl=?", tCcl.sDescCcl, tCcl.iTipoCcl, tCcl.iAtivo, objCcl.sCcl)
    If lErro <> AD_SQL_SUCESSO Then Error 5598

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 5600
        
    'Ccl não está cadastrado
    If lErro = AD_SQL_SEM_DADOS Then Error 5599

    objCcl.sDescCcl = tCcl.sDescCcl
    objCcl.iTipoCcl = tCcl.iTipoCcl
    objCcl.iAtivo = tCcl.iAtivo
    
    Call Comando_Fechar(lComando)

    Ccl_Le = SUCESSO

Exit Function

Erro_Ccl_Le:

    Ccl_Le = Err
    
    Select Case Err

        Case 5597
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 5598, 5600
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CCL", Err, objCcl.sCcl)
         
        Case 5599
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153415)

    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function
    
Function ContaCcl_Le_Ccl(ByVal sCcl As String) As Long
'verifica se existe no BD pelo menos uma associacao do centro de custo fornecido com conta

Dim lComando As Long
Dim lErro As Long
Dim sConta As String

On Error GoTo Erro_ContaCcl_Le_Ccl

    sConta = String(STRING_CONTA, 0)
    
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 5601
    
    lErro = Comando_Executar(lComando, "SELECT Conta FROM ContaCcl WHERE Ccl=?", sConta, sCcl)
    If lErro <> AD_SQL_SUCESSO Then Error 5602
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 5604
        
    If lErro = AD_SQL_SEM_DADOS Then Error 5603
        
    Call Comando_Fechar(lComando)
    
    ContaCcl_Le_Ccl = SUCESSO
    
    Exit Function
 
Erro_ContaCcl_Le_Ccl:

    ContaCcl_Le_Ccl = Err
    
    Select Case Err
    
        Case 5601
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 5602, 5604
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTACCL2", Err, sConta)
                
        Case 5603
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153416)
    
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function Lancamento_Le_Ccl(ByVal sCcl As String) As Long
'verifica se existe no BD pelo menos um lancamento para o centro de custo/lucro fornecido

Dim lComando As Long
Dim lErro As Long
Dim lDoc As Long

On Error GoTo Erro_Lancamento_Le_Ccl

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 5605
    
    lErro = Comando_Executar(lComando, "SELECT Doc FROM Lancamentos WHERE Ccl=?", lDoc, sCcl)
    If lErro <> AD_SQL_SUCESSO Then Error 5606
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 5608
        
    If lErro = AD_SQL_SEM_DADOS Then Error 5607
        
    Call Comando_Fechar(lComando)
    
    Lancamento_Le_Ccl = SUCESSO
    
    Exit Function
 
Erro_Lancamento_Le_Ccl:

    Lancamento_Le_Ccl = Err
    
    Select Case Err
    
        Case 5605
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 5606, 5608
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LANCAMENTOS4", Err, sCcl)
                
        Case 5607
           
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153417)
    
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function LanPendente_Le_Ccl(ByVal sCcl As String) As Long
'verifica se existe no BD pelo menos um lancamento pendente para o centro deu custo/lucro fornecido

Dim lComando As Long
Dim lErro As Long
Dim lDoc As Long

On Error GoTo Erro_LanPendente_Le_Ccl

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 5609
    
    lErro = Comando_Executar(lComando, "SELECT Doc FROM LanPendente WHERE Ccl=?", lDoc, sCcl)
    If lErro <> AD_SQL_SUCESSO Then Error 5610
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 5612
        
    If lErro = AD_SQL_SEM_DADOS Then Error 5611
        
    Call Comando_Fechar(lComando)
    
    LanPendente_Le_Ccl = SUCESSO
    
    Exit Function
 
Erro_LanPendente_Le_Ccl:

    LanPendente_Le_Ccl = Err
    
    Select Case Err
    
        Case 5609
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 5610, 5612
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LANPENDENTE2", Err, sCcl)
         
        Case 5611
                   
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153418)
    
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function
    
Function DocAuto_Le_Ccl(ByVal sCcl As String) As Long
'verifica se existe na tabela DocAuto pelo menos um documento automático com o Ccl em questão.

Dim lComando As Long
Dim lErro As Long
Dim lDoc As Long
Dim iSeq As Integer

On Error GoTo Erro_DocAuto_Le_Ccl

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 8044
    
    lErro = Comando_Executar(lComando, "SELECT Doc, Seq FROM DocAuto WHERE Ccl=?", lDoc, iSeq, sCcl)
    If lErro <> AD_SQL_SUCESSO Then Error 8045
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 8047
        
    If lErro = AD_SQL_SEM_DADOS Then Error 8097
        
    Call Comando_Fechar(lComando)
    
    DocAuto_Le_Ccl = SUCESSO
    
    Exit Function
 
Erro_DocAuto_Le_Ccl:

    DocAuto_Le_Ccl = Err
    
    Select Case Err
    
        Case 8044
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 8045, 8047
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_DOCAUTO", Err, sCcl)
           
        Case 8097
                   
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153419)
    
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function RateioOn_Le_Ccl(ByVal sCcl As String) As Long
'verifica se existe na tabela RateioOn pelo menos um rateio com o Ccl em questão.

Dim lComando As Long
Dim lErro As Long
Dim iCodigo As Integer
Dim iSeq As Integer

On Error GoTo Erro_RateioOn_Le_Ccl

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 8087
    
    lErro = Comando_Executar(lComando, "SELECT Codigo, Seq FROM RateioOn WHERE Ccl=?", iCodigo, iSeq, sCcl)
    If lErro <> AD_SQL_SUCESSO Then Error 8088
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 8090
    
    If lErro = AD_SQL_SEM_DADOS Then Error 8098
        
    Call Comando_Fechar(lComando)
    
    RateioOn_Le_Ccl = SUCESSO
    
    Exit Function
 
Erro_RateioOn_Le_Ccl:

    RateioOn_Le_Ccl = Err
    
    Select Case Err
    
        Case 8087
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 8088, 8090
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RATEIOON", Err, sCcl)
        
        Case 8098
          
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153420)
    
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function RateioOff_Le_Ccl(ByVal sCcl As String) As Long
'verifica se existe na tabela RateioOff pelo menos um rateio com o Ccl em questão.

Dim lComando As Long
Dim lErro As Long
Dim lSeq As Long

On Error GoTo Erro_RateioOff_Le_Ccl

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 5977
    
    lErro = Comando_Executar(lComando, "SELECT Seq FROM RateioOff WHERE CclOrigem=? Or Ccl=?", lSeq, sCcl, sCcl)
    If lErro <> AD_SQL_SUCESSO Then Error 5978
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 5979
        
    If lErro = AD_SQL_SEM_DADOS Then Error 8100
        
    Call Comando_Fechar(lComando)
    
    RateioOff_Le_Ccl = SUCESSO
    
    Exit Function
 
Erro_RateioOff_Le_Ccl:

    RateioOff_Le_Ccl = Err
    
    Select Case Err
    
        Case 5977
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 5978, 5979
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RATEIOOFF", Err, sCcl)
              
        Case 8100
                   
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153421)
    
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function HistPadrao_Le_Todos_Historicos(ByVal colHistPadrao As Collection) As Long
'Preenche coleção para preencher listbox de Históricos

Dim lErro As Long
Dim iFim_de_Arquivo As Integer
Dim lComando As Long
Dim iCodigo As Integer
Dim sCodigo As String
Dim sDescricao As String
Dim sListBoxItem As String
Dim sEspacos As String
    
On Error GoTo Erro_HistPadrao_Le_Todos_Historicos

    sDescricao = Space(STRING_HISTORICO)
    
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 6171
    
    lErro = Comando_Executar(lComando, "SELECT HistPadrao, DescHistPadrao FROM HistPadrao", iCodigo, sDescricao)
    If lErro <> AD_SQL_SUCESSO Then Error 6172
    
    iFim_de_Arquivo = Comando_BuscarPrimeiro(lComando)
        
    Do While iFim_de_Arquivo = AD_SQL_SUCESSO
    
        'Espaços que faltam para completar tamanho STRING_CODIGO_HISTORICO
        sEspacos = Space(STRING_CODIGO_HISTORICO - Len(CStr(iCodigo)))
        
        'Concatena Codigo e Nome do HistPadrao
        sCodigo = sEspacos & CStr(iCodigo)
        sListBoxItem = sCodigo & SEPARADOR & Trim(sDescricao)
        
        'Adiciona à Coleção
        colHistPadrao.Add (sListBoxItem)

        iFim_de_Arquivo = Comando_BuscarProximo(lComando)

    Loop
    
    'Se o erro não foi fim de arquivo desvia para rotina de erro
    If iFim_de_Arquivo <> AD_SQL_SEM_DADOS Then Error 6173
    
    Call Comando_Fechar(lComando)
    
    HistPadrao_Le_Todos_Historicos = SUCESSO
    
    Exit Function

Erro_HistPadrao_Le_Todos_Historicos:

    HistPadrao_Le_Todos_Historicos = Err
    
    Select Case Err
                
        Case 6171
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 6172, 6173
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_HISTPADRAO1", Err)
                        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 153422)
        
        End Select
        
        Call Comando_Fechar(lComando)
        
    Exit Function

End Function

Public Function ContaCcl_Le_Todos_Conta(ByVal colContas As Collection, ByVal iTipoConta As Integer) As Long
'Preenche a colecao colContas com o conteudo de PlanoContas, de acordo com iTipoConta

Dim lErro As Long
Dim lComando As Long
Dim sConta As String
Dim sDescricao As String
Dim objPlanoConta As ClassPlanoConta
    
On Error GoTo Erro_ContaCcl_Le_Todos_Conta

    
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 8105
    
    sConta = String(STRING_CONTA, 0)
    sDescricao = String(STRING_CONTA_DESCRICAO, 0)
    
    lErro = Comando_Executar(lComando, "SELECT Conta, DescConta FROM PlanoConta WHERE TipoConta=? ORDER BY Conta", sConta, sDescricao, iTipoConta)
    If lErro <> AD_SQL_SUCESSO Then Error 8106
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9915
    
    If lErro = AD_SQL_SEM_DADOS Then Error 8127
            
    Do While lErro = AD_SQL_SUCESSO
    
        Set objPlanoConta = New ClassPlanoConta
    
        objPlanoConta.sConta = sConta
        objPlanoConta.sDescConta = sDescricao
        
        colContas.Add objPlanoConta

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 8107

    Loop
    
    Call Comando_Fechar(lComando)
    
    ContaCcl_Le_Todos_Conta = SUCESSO
    
    Exit Function

Erro_ContaCcl_Le_Todos_Conta:

    ContaCcl_Le_Todos_Conta = Err
    
    Select Case Err
                
        Case 8105
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 8106, 8107, 9915
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTA", Err)
            
        Case 8127
           
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 153423)
        
        End Select
        
        Call Comando_Fechar(lComando)
        
    Exit Function
            
End Function

Public Function ContaCcl_Le_Todos_Ccl(ByVal colCcl As Collection) As Long
'Preenche colecao para preencher listbox de Ccl

Dim lErro As Long
Dim lComando As Long
Dim sCcl As String
    
On Error GoTo Erro_ContaCcl_Le_Todos_Ccl

    sCcl = String(STRING_CCL, 0)
    
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 8108
    
    lErro = Comando_Executar(lComando, "SELECT Ccl FROM Ccl", sCcl)
    If lErro <> AD_SQL_SUCESSO Then Error 8109
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9917
    
    If lErro = AD_SQL_SEM_DADOS Then Error 8128
        
    Do While lErro = AD_SQL_SUCESSO
    
        colCcl.Add sCcl

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9918

    Loop
    
    Call Comando_Fechar(lComando)
    
    ContaCcl_Le_Todos_Ccl = SUCESSO
    
    Exit Function

Erro_ContaCcl_Le_Todos_Ccl:

    ContaCcl_Le_Todos_Ccl = Err
    
    Select Case Err
                
        Case 8108
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 8109, 9917, 9918
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CCL", Err, sCcl)
                  
        Case 8110
                      
        Case 8128
            'Relação vazia
                        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 153424)
        
        End Select
        
        Call Comando_Fechar(lComando)
        
    Exit Function

End Function

Public Function ContaCcl_Associa_Ccl(ByVal colCcl As Collection, ByVal sConta As String) As Long
'Retorna em colCcl os centros de custo/lucro associados à conta sConta. /m

Dim lErro As Long
Dim lComando1 As Long
Dim sCcl As String

On Error GoTo Erro_ContaCcl_Associa_Ccl
    
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 8111
    
    sCcl = String(STRING_CCL, 0)
    
    'pesquisa as associações da conta em questão com centros de custo
    lErro = Comando_Executar(lComando1, "SELECT Ccl FROM ContaCcl WHERE Conta=?", sCcl, sConta)
    If lErro <> AD_SQL_SUCESSO Then Error 8112
        
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9875
    
    'Marca os Ccl's específicos na ListBox Ccl
    Do While lErro = AD_SQL_SUCESSO
        
        colCcl.Add sCcl
               
        lErro = Comando_BuscarProximo(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 8113
                   
    Loop

    Call Comando_Fechar(lComando1)
    
    ContaCcl_Associa_Ccl = SUCESSO
    
    Exit Function

Erro_ContaCcl_Associa_Ccl:

    ContaCcl_Associa_Ccl = Err
    
    Select Case Err
                
        Case 8111
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 8112, 8113, 9875
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTACCL2", Err, sConta)
                        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 153425)
        
    End Select
        
    Call Comando_Fechar(lComando1)
        
    Exit Function

End Function

Public Function PlanoConta_Le_Conta(ByVal sConta As String) As Long
'Verifica no BD a existencia de sConta

Dim lErro As Long
Dim lComando1 As Long
Dim sConta1 As String
Dim iFim_de_Arquivo  As Integer
Dim iIndice As Integer

On Error GoTo Erro_PlanoConta_Le_Conta
  
    sConta1 = String(STRING_CONTA, 0)
    
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 8121
        
    'Seleciona conta em PlanoConta
    lErro = Comando_Executar(lComando1, "SELECT Conta FROM PlanoConta WHERE Conta = ?", sConta1, sConta)
    If lErro <> AD_SQL_SUCESSO Then Error 8122
           
    'Verifica se a conta existe
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 8123
    
    If lErro = AD_SQL_SEM_DADOS Then Error 10051
    
    Call Comando_Fechar(lComando1)
       
    PlanoConta_Le_Conta = SUCESSO
       
    Exit Function
       
Erro_PlanoConta_Le_Conta:

    PlanoConta_Le_Conta = Err
    
    Select Case Err
                
        Case 8121
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 8122, 8123
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTA", Err, sConta)
            
        Case 10051
                        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 153426)
        
        End Select
        
        Call Comando_Fechar(lComando1)
    
End Function

Public Function ContaCcl_Associa_Conta(ByVal colContas As Collection, ByVal sCcl As String) As Long
'Retorna em colContas as contas associadas ao centro de custo sCcl. /m

Dim lErro As Long
Dim lComando1 As Long
Dim sConta As String

On Error GoTo Erro_ContaCcl_Associa_Conta
    
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 8124
    
    sConta = String(STRING_CONTA, 0)
    
    'pesquisa as associações da conta com centros de custo
    lErro = Comando_Executar(lComando1, "SELECT Conta FROM ContaCcl WHERE Ccl=?", sConta, sCcl)
    If lErro <> AD_SQL_SUCESSO Then Error 8125
        
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9879
    
    'Marca as contas específicas na ListBox Contas
    Do While lErro = AD_SQL_SUCESSO
        
        colContas.Add sConta
        
        lErro = Comando_BuscarProximo(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 8126
                   
    Loop

    Call Comando_Fechar(lComando1)
    
    ContaCcl_Associa_Conta = SUCESSO
    
    Exit Function

Erro_ContaCcl_Associa_Conta:

    ContaCcl_Associa_Conta = Err
    
    Select Case Err
                
        Case 8124
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 8125, 8126, 9879
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTACCL2", Err, sCcl)
          
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 153427)
        
    End Select
        
    Call Comando_Fechar(lComando1)
    
    Exit Function

End Function

Function Exercicio_Le(ByVal iExercicio As Integer, ByVal objExercicio As ClassExercicio) As Long
'retorna os dados do exercicio em questão

Dim lComando As Long
Dim lErro As Long
Dim tExercicio As typeExercicio

On Error GoTo Erro_Exercicio_Le

    tExercicio.sNomeExterno = String(STRING_EXERCICIO_NOME_EXTERNO, 0)
    
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 5822
    
    'seleciona os dados do exercicio em questão
    lErro = Comando_Executar(lComando, "SELECT NumPeriodos, Status, NomeExterno, DataInicio, DataFim FROM Exercicios WHERE Exercicio=?", tExercicio.iNumPeriodos, tExercicio.iStatus, tExercicio.sNomeExterno, tExercicio.dtDataInicio, tExercicio.dtDataFim, iExercicio)
    If lErro <> AD_SQL_SUCESSO Then Error 5823
    
    'le o exercicio em questão
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 5824
    
    'se o exercicio não estiver cadastrado
    If lErro = AD_SQL_SEM_DADOS Then Error 10083
    
    objExercicio.iExercicio = iExercicio
    objExercicio.iNumPeriodos = tExercicio.iNumPeriodos
    objExercicio.iStatus = tExercicio.iStatus
    objExercicio.sNomeExterno = tExercicio.sNomeExterno
    objExercicio.dtDataInicio = tExercicio.dtDataInicio
    objExercicio.dtDataFim = tExercicio.dtDataFim
    
    Call Comando_Fechar(lComando)
    
    Exercicio_Le = SUCESSO

    Exit Function

Erro_Exercicio_Le:

    Exercicio_Le = Err
    
    Select Case Err
    
        Case 5822
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 5823, 5824
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EXERCICIO2", Err, iExercicio)
            
        Case 10083
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153428)
        
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function Periodo_Le_ExercicioPeriodo(ByVal iExercicio As Integer, ByVal iPeriodo As Integer, ByVal objPeriodo As ClassPeriodo) As Long
'le o registro referente ao periodo em questão e coloca os dados em objPeriodo

Dim lComando As Long
Dim lErro As Long
Dim tPeriodo As typePeriodo

On Error GoTo Erro_Periodo_Le_ExercicioPeriodo

    tPeriodo.sNomeExterno = String(STRING_PERIODO_NOME_EXTERNO, 0)
    
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 5829
    
    lErro = Comando_Executar(lComando, "SELECT NomeExterno, DataInicio, DataFim FROM Periodo WHERE Exercicio=? AND Periodo=?", tPeriodo.sNomeExterno, tPeriodo.dtDataInicio, tPeriodo.dtDataFim, iExercicio, iPeriodo)
    If lErro <> AD_SQL_SUCESSO Then Error 5830
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 5831
    
    objPeriodo.iExercicio = iExercicio
    objPeriodo.iPeriodo = iPeriodo
    objPeriodo.sNomeExterno = tPeriodo.sNomeExterno
    objPeriodo.dtDataInicio = tPeriodo.dtDataInicio
    objPeriodo.dtDataFim = tPeriodo.dtDataFim
    
    Call Comando_Fechar(lComando)
    
    Periodo_Le_ExercicioPeriodo = SUCESSO

    Exit Function

Erro_Periodo_Le_ExercicioPeriodo:

    Periodo_Le_ExercicioPeriodo = Err
    
    Select Case Err
    
        Case 5829
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 5830, 5831
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PERIODO", Err)
     
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153429)
        
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function LanPendente_Le_Doc(ByVal objDoc As ClassLancamento_Cabecalho, ByVal colLancamentos As Collection) As Long
'le os lançamentos relativos ao documento passado como parametro e coloca-os em colLancamentos

Dim lComando As Long
Dim lErro As Long
Dim objLanc As ClassLancamento_Detalhe

Dim tLanc As typeLancamento

On Error GoTo Erro_LanPendente_Le_Doc

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 5835
        
    tLanc.sOrigem = String(STRING_ORIGEM, 0)
    tLanc.sConta = String(STRING_CONTA, 0)
    tLanc.sCcl = String(STRING_CCL, 0)
    tLanc.sHistorico = String(STRING_HISTORICO, 0)
    tLanc.sProduto = String(STRING_PRODUTO, 0)
    tLanc.sModelo = String(STRING_PADRAOCONTAB_MODELO, 0)
    tLanc.sDocOrigem = String(STRING_LANPENDENTE_DOCORIGEM, 0)
                            
    If giSetupDocPorPeriodo = DOC_INICIALIZADO_POR_PERIODO Then

        lErro = Comando_Executar(lComando, "SELECT Origem, Exercicio, PeriodoLan, Doc, Seq, Lote, PeriodoLote, Data, Conta, Ccl, Historico, Valor, NumIntDoc, FilialCliForn, CliForn, Transacao, ContaSimples, SeqContraPartida, Produto, DataEstoque, Status, Modelo, ValorLivroAuxiliar, ClienteFornecedor, FilialEmpresa,  Aglutina, ApropriaCrProd, EscaninhoCusto, DocOrigem, Quantidade, Gerencial " & _
        "FROM LanPendente WHERE FilialEmpresa=? AND Origem=? AND Exercicio=? AND PeriodoLan=? AND Doc=? ORDER BY Seq", tLanc.sOrigem, tLanc.iExercicio, tLanc.iPeriodoLan, tLanc.lDoc, tLanc.iSeq, tLanc.iLote, tLanc.iPeriodoLote, tLanc.dtData, tLanc.sConta, tLanc.sCcl, tLanc.sHistorico, tLanc.dValor, tLanc.lNumIntDoc, tLanc.iFilialCliForn, tLanc.lCliForn, tLanc.iTransacao, tLanc.lContaSimples, tLanc.iSeqContraPartida, tLanc.sProduto, tLanc.dtDataEstoque, tLanc.iStatus, tLanc.sModelo, tLanc.dValorLivroAuxiliar, tLanc.iClienteFornecedor, tLanc.iFilialEmpresa, tLanc.iAglutina, tLanc.iApropriaCRProd, tLanc.iEscaninho_Custo, tLanc.sDocOrigem, tLanc.dQuantidade, tLanc.iGerencial, objDoc.iFilialEmpresa, objDoc.sOrigem, objDoc.iExercicio, objDoc.iPeriodoLan, objDoc.lDoc)
        If lErro <> AD_SQL_SUCESSO Then Error 5836
        
    Else

        lErro = Comando_Executar(lComando, "SELECT Origem, Exercicio, PeriodoLan, Doc, Seq, Lote, PeriodoLote, Data, Conta, Ccl, Historico, Valor, NumIntDoc, FilialCliForn, CliForn, Transacao, ContaSimples, SeqContraPartida, Produto, DataEstoque, Status, Modelo, ValorLivroAuxiliar, ClienteFornecedor, FilialEmpresa, Aglutina, ApropriaCrProd, EscaninhoCusto, DocOrigem, Quantidade, Gerencial " & _
        "FROM LanPendente WHERE FilialEmpresa=? AND Origem=? AND Exercicio=? AND Doc=? ORDER BY Seq", tLanc.sOrigem, tLanc.iExercicio, tLanc.iPeriodoLan, tLanc.lDoc, tLanc.iSeq, tLanc.iLote, tLanc.iPeriodoLote, tLanc.dtData, tLanc.sConta, tLanc.sCcl, tLanc.sHistorico, tLanc.dValor, tLanc.lNumIntDoc, tLanc.iFilialCliForn, tLanc.lCliForn, tLanc.iTransacao, tLanc.lContaSimples, tLanc.iSeqContraPartida, tLanc.sProduto, tLanc.dtDataEstoque, tLanc.iStatus, tLanc.sModelo, tLanc.dValorLivroAuxiliar, tLanc.iClienteFornecedor, tLanc.iFilialEmpresa, tLanc.iAglutina, tLanc.iApropriaCRProd, tLanc.iEscaninho_Custo, tLanc.sDocOrigem, tLanc.dQuantidade, tLanc.iGerencial, objDoc.iFilialEmpresa, objDoc.sOrigem, objDoc.iExercicio, objDoc.lDoc)
        If lErro <> AD_SQL_SUCESSO Then Error 5837
        
    End If
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 5838
        
    If lErro = AD_SQL_SEM_DADOS Then Error 5842
        
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objLanc = New ClassLancamento_Detalhe
        
        objLanc.iFilialEmpresa = objDoc.iFilialEmpresa
        objLanc.sOrigem = tLanc.sOrigem
        objLanc.iExercicio = tLanc.iExercicio
        objLanc.iPeriodoLan = tLanc.iPeriodoLan
        objLanc.lDoc = tLanc.lDoc
        objLanc.iSeq = tLanc.iSeq
        objLanc.iLote = tLanc.iLote
        objLanc.iPeriodoLote = tLanc.iPeriodoLote
        objLanc.dtData = tLanc.dtData
        objLanc.sConta = tLanc.sConta
        objLanc.sCcl = tLanc.sCcl
        objLanc.sHistorico = tLanc.sHistorico
        objLanc.dValor = tLanc.dValor
        objLanc.lNumIntDoc = tLanc.lNumIntDoc
        objLanc.iFilialCliForn = tLanc.iFilialCliForn
        objLanc.lCliForn = tLanc.lCliForn
        objLanc.iTransacao = tLanc.iTransacao
        objLanc.lContaSimples = tLanc.lContaSimples
        objLanc.iSeqContraPartida = tLanc.iSeqContraPartida
        objLanc.sProduto = tLanc.sProduto
        objLanc.dtDataEstoque = tLanc.dtDataEstoque
        objLanc.iStatus = tLanc.iStatus
        objLanc.sModelo = tLanc.sModelo
        objLanc.dValorLivroAuxiliar = tLanc.dValorLivroAuxiliar
        objLanc.iClienteFornecedor = tLanc.iClienteFornecedor
        objLanc.iApropriaCRProd = tLanc.iApropriaCRProd
        objLanc.iEscaninho_Custo = tLanc.iEscaninho_Custo
        objLanc.sDocOrigem = tLanc.sDocOrigem
        objLanc.dQuantidade = tLanc.dQuantidade
        objLanc.iGerencial = tLanc.iGerencial
        
        colLancamentos.Add objLanc
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 5839
        
    Loop
    
    Call Comando_Fechar(lComando)
    
    LanPendente_Le_Doc = SUCESSO
    
    Exit Function
 
Erro_LanPendente_Le_Doc:

    LanPendente_Le_Doc = Err
    
    Select Case Err
    
        Case 5835
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 5836, 5837, 5838, 5839
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LANPENDENTE3", Err)
           
        Case 5842
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153430)
    
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function ContaCcl_Le(ByVal objContaCcl As ClassContaCcl) As Long
'le a contaccl a partir da chave fornecida em objContaCcl. Devolve os dados em objContaCcl

Dim lComando As Long
Dim lErro As Long
Dim sConta As String

On Error GoTo Erro_ContaCcl_Le

    sConta = String(STRING_CONTA, 0)
    
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 5868

    'Verifica se a ContaCcl existe
    lErro = Comando_Executar(lComando, "SELECT Conta FROM ContaCcl WHERE Conta=? AND Ccl=?", sConta, objContaCcl.sConta, objContaCcl.sCcl)
    If lErro <> AD_SQL_SUCESSO Then Error 5869

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 5870
    
    'ContaCcl não está cadastrado
    If lErro = AD_SQL_SEM_DADOS Then Error 5871

    Call Comando_Fechar(lComando)

    ContaCcl_Le = SUCESSO

Exit Function

Erro_ContaCcl_Le:

    ContaCcl_Le = Err
    
    Select Case Err

        Case 5868
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 5869, 5870
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTACCL3", Err, objContaCcl.sConta, objContaCcl.sCcl)
          
        Case 5871
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153431)

    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Configuracao_Le(ByVal objConfiguracao As ClassConfiguracao) As Long
'le o registro da tabela de Configuracao

Dim lComando As Long
Dim lErro As Long
Dim tConfiguracao As typeConfiguracao

On Error GoTo Erro_Configuracao_Le

    tConfiguracao.sGrupoAtivo = String(STRING_CONTA, 0)
    tConfiguracao.sGrupoPassivo = String(STRING_CONTA, 0)
    tConfiguracao.sGrupoReceita = String(STRING_CONTA, 0)
    tConfiguracao.sGrupoDespesa = String(STRING_CONTA, 0)
    
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 5956

    lErro = Comando_Executar(lComando, "SELECT GrupoReceita, GrupoDespesa, GrupoAtivo, GrupoPassivo, IdAtualizacao, UsoCcl, DocPorPeriodo, LotePorPeriodo, TipoContaDefault, NaturezaDefault, HistPadrao, DataImplantacao, ContabCentralizada, ContabGerencial, FilialAuxiliar FROM Configuracao", tConfiguracao.sGrupoReceita, tConfiguracao.sGrupoDespesa, tConfiguracao.sGrupoAtivo, tConfiguracao.sGrupoPassivo, tConfiguracao.iIdAtualizacao, tConfiguracao.iUsoCcl, tConfiguracao.iDocPorPeriodo, tConfiguracao.iLotePorPeriodo, tConfiguracao.iTipoContaDefault, tConfiguracao.iNaturezaDefault, tConfiguracao.iHistPadrao, tConfiguracao.dtDataImplantacao, tConfiguracao.iContabCentralizada, tConfiguracao.iContabGerencial, tConfiguracao.iFilialAuxiliar)
    If lErro <> AD_SQL_SUCESSO Then Error 5957

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 5958

    If lErro = AD_SQL_SEM_DADOS Then Error 5959

    objConfiguracao.sGrupoReceita = tConfiguracao.sGrupoReceita
    objConfiguracao.sGrupoDespesa = tConfiguracao.sGrupoDespesa
    objConfiguracao.sGrupoAtivo = tConfiguracao.sGrupoAtivo
    objConfiguracao.sGrupoPassivo = tConfiguracao.sGrupoPassivo
    objConfiguracao.iIdAtualizacao = tConfiguracao.iIdAtualizacao
    objConfiguracao.iUsoCcl = tConfiguracao.iUsoCcl
    objConfiguracao.iDocPorPeriodo = tConfiguracao.iDocPorPeriodo
    objConfiguracao.iLotePorPeriodo = tConfiguracao.iLotePorPeriodo
    objConfiguracao.iTipoContaDefault = tConfiguracao.iTipoContaDefault
    objConfiguracao.iNaturezaDefault = tConfiguracao.iNaturezaDefault
    objConfiguracao.iHistPadrao = tConfiguracao.iHistPadrao
    objConfiguracao.dtDataImplantacao = tConfiguracao.dtDataImplantacao
    objConfiguracao.iContabCentralizada = tConfiguracao.iContabCentralizada
    objConfiguracao.iContabGerencial = tConfiguracao.iContabGerencial
    objConfiguracao.iFilialAuxiliar = tConfiguracao.iFilialAuxiliar
    
    Call Comando_Fechar(lComando)

    Configuracao_Le = SUCESSO

    Exit Function

Erro_Configuracao_Le:

    Configuracao_Le = Err
    
    Select Case Err

        Case 5956
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 5957, 5958
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONFIGURACAO", Err)
        
        Case 5959
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONFIGURACAO_NAO_CADASTRADA", Err)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153432)

    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Exercicio_Le_Data(ByVal dtData As Date, ByVal objExercicio As ClassExercicio) As Long
'retorna os dados do exercicio para a data em questão.

Dim lComando As Long
Dim lErro As Long
Dim tExercicio As typeExercicio

On Error GoTo Erro_Exercicio_Le_Data

    tExercicio.sNomeExterno = String(STRING_EXERCICIO_NOME_EXTERNO, 0)
    
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 5962
    
    lErro = Comando_Executar(lComando, "SELECT Exercicio, NumPeriodos, Status, NomeExterno, DataInicio, DataFim FROM Exercicios WHERE DataInicio <= ? And DataFim >= ?", tExercicio.iExercicio, tExercicio.iNumPeriodos, tExercicio.iStatus, tExercicio.sNomeExterno, tExercicio.dtDataInicio, tExercicio.dtDataFim, dtData, dtData)
    If lErro <> AD_SQL_SUCESSO Then Error 5963
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 5964
    
    If lErro = AD_SQL_SEM_DADOS Then Error 5965
    
    Call Comando_Fechar(lComando)
    
    objExercicio.iExercicio = tExercicio.iExercicio
    objExercicio.iNumPeriodos = tExercicio.iNumPeriodos
    objExercicio.iStatus = tExercicio.iStatus
    objExercicio.sNomeExterno = tExercicio.sNomeExterno
    objExercicio.dtDataInicio = tExercicio.dtDataInicio
    objExercicio.dtDataFim = tExercicio.dtDataFim
    
    Exercicio_Le_Data = SUCESSO

    Exit Function

Erro_Exercicio_Le_Data:

    Exercicio_Le_Data = Err
    
    Select Case Err
    
    Case 5962
        lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
    Case 5963, 5964
        lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EXERCICIO1", Err)
        
    Case 5965
    
    Case Else
        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153433)
        
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function Exercicios_Abertos_Le_Todos(ByVal colExerciciosAbertos As Collection) As Long
'le todos os exercicios abertos e coloca-os em colExerciciosAbertos

Dim lComando As Long
Dim lErro As Long
Dim tExercicio As typeExercicio
Dim objExercicio As ClassExercicio
    
On Error GoTo Erro_Exercicios_Abertos_Le_Todos

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 5980
    
    tExercicio.sNomeExterno = String(STRING_EXERCICIO_NOME_EXTERNO, 0)

    lErro = Comando_Executar(lComando, "SELECT Exercicio, NumPeriodos, Status, NomeExterno, DataInicio, DataFim FROM Exercicios  WHERE Status = ? ORDER BY Exercicio", tExercicio.iExercicio, tExercicio.iNumPeriodos, tExercicio.iStatus, tExercicio.sNomeExterno, tExercicio.dtDataInicio, tExercicio.dtDataFim, EXERCICIO_ABERTO)
    If lErro <> AD_SQL_SUCESSO Then Error 5981
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 5982
    
    Do While lErro = AD_SQL_SUCESSO
    
        Set objExercicio = New ClassExercicio
    
        objExercicio.iExercicio = tExercicio.iExercicio
        objExercicio.iNumPeriodos = tExercicio.iNumPeriodos
        objExercicio.iStatus = tExercicio.iStatus
        objExercicio.sNomeExterno = tExercicio.sNomeExterno
        objExercicio.dtDataInicio = tExercicio.dtDataInicio
        objExercicio.dtDataFim = tExercicio.dtDataFim

        colExerciciosAbertos.Add objExercicio
        
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 5983
        
    Loop
        
    Call Comando_Fechar(lComando)
    
    Exercicios_Abertos_Le_Todos = SUCESSO
    
    Exit Function
    
    
Erro_Exercicios_Abertos_Le_Todos:

    Exercicios_Abertos_Le_Todos = Err
    
    Select Case Err
    
        Case 5980
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 5981, 5982, 5983
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EXERCICIO1", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153434)
        
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
    If lComando = 0 Then Error 5388

    'Pesquisa os periodos do exercicio em questão
    lErro = Comando_Executar(lComando, "SELECT PeriodosFilial.Exercicio, PeriodosFilial.Periodo, NomeExterno, Lancamentos, Apurado, DataInicio, DataFim, Lote, Fechado FROM PeriodosFilial, Periodo WHERE PeriodosFilial.FilialEmpresa = ? AND PeriodosFilial.Exercicio = ? AND PeriodosFilial.Exercicio = Periodo.Exercicio AND PeriodosFilial.Periodo = Periodo.Periodo ORDER BY PeriodosFilial.Periodo", tPeriodo.iExercicio, tPeriodo.iPeriodo, tPeriodo.sNomeExterno, tPeriodo.iLancamentos, tPeriodo.iApurado, tPeriodo.dtDataInicio, tPeriodo.dtDataFim, tPeriodo.iLote, tPeriodo.iFechado, iFilialEmpresa, iExercicio)
    If lErro <> AD_SQL_SUCESSO Then Error 5389
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 5985
    
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
        
        colPeriodos.Add objPeriodo
        
        'le o proximo periodo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 5986
                
    Loop
    
    Call Comando_Fechar(lComando)
    
    Periodo_Le_Todos_Exercicio = SUCESSO
    
    Exit Function
    
    
Erro_Periodo_Le_Todos_Exercicio:

    Periodo_Le_Todos_Exercicio = Err
    
    Select Case Err
    
        Case 5388
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 5389, 5985, 5986
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PERIODOS", Err)
           
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153435)
        
    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function
    
End Function

Function Exercicio_Le_NumRegistros(lNumReg As Long) As Long
'retorna o número de registros contidos na tabela de Exercicio

Dim lErro As Long
Dim lComando As Long

On Error GoTo Erro_Exercicio_Le_NumRegistros

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 9839

    'Le o número de registro na tabela de Exercicio
    lErro = Comando_Executar(lComando, "SELECT COUNT(*) FROM Exercicios", lNumReg)
    If lErro <> AD_SQL_SUCESSO Then Error 9840

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 9841
    
    Call Comando_Fechar(lComando)
    
    Exercicio_Le_NumRegistros = SUCESSO

    Exit Function

Erro_Exercicio_Le_NumRegistros:

    Exercicio_Le_NumRegistros = Err
    
    Select Case Err

        Case 9839
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 9840, 9841
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EXERCICIO1", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153436)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function LotePendente_Le_Posicionado(ByVal colLote As Collection, ByVal lPosicao As Long, ByVal lNumReg As Long, ByVal lComando As Long, ByVal sOrigem As String, ByVal iExercicio As Integer, ByVal iPeriodo As Integer, ByVal iLote As Integer) As Long
'le os lotes lNumReg a partir da posicao lPosicao e coloca-os e colLote.

Dim lErro As Long
Dim objLote As ClassLote
Dim lNumRegLido As Long
    
On Error GoTo Erro_LotePendente_Le_Posicionado

    sOrigem = String(STRING_ORIGEM, 0)

    If lComando = 0 Then Error 9013
    
    lErro = Comando_Executar(lComando, "SELECT Origem, Exercicio, Periodo, Lote FROM LotePendente", sOrigem, iExercicio, iPeriodo, iLote)
    If lErro <> AD_SQL_SUCESSO Then Error 9014

    lErro = Comando_BuscarAbsoluto(lComando, lPosicao)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9015

    lNumRegLido = 1

    Do While lErro <> AD_SQL_SEM_DADOS And lNumRegLido <= lNumReg

        Set objLote = New ClassLote

        objLote.sOrigem = sOrigem
        objLote.iExercicio = iExercicio
        objLote.iPeriodo = iPeriodo
        objLote.iLote = iLote

        colLote.Add objLote

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9016

        lNumRegLido = lNumRegLido + 1
    Loop
    
    LotePendente_Le_Posicionado = SUCESSO

    Exit Function

Erro_LotePendente_Le_Posicionado:

    LotePendente_Le_Posicionado = Err
    
    Select Case Err
    
        Case 9013
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 9014, 9015, 9016
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LOTEPENDENTE1", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153437)

    End Select
    
    
    Exit Function

End Function

Function Exercicio_Le_NomeExterno(ByVal lComando As Long, ByVal iExercicio As Integer, sNomeExterno As String) As Long
'le o nome externo do exercicio passado como parametro

Dim lErro As Long

On Error GoTo Erro_Exercicio_Le_NomeExterno

    sNomeExterno = String(STRING_EXERCICIO_NOME_EXTERNO, 0)
    
    lErro = Comando_Executar(lComando, "SELECT NomeExterno FROM Exercicios WHERE Exercicio=?", sNomeExterno, iExercicio)
    If lErro <> AD_SQL_SUCESSO Then Error 9265
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 9266
    
    Exercicio_Le_NomeExterno = SUCESSO

    Exit Function

Erro_Exercicio_Le_NomeExterno:

    Exercicio_Le_NomeExterno = Err
    
    Select Case Err
    
        Case 9265, 9266
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EXERCICIO2", Err, iExercicio)
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153438)
        
    End Select
    
    Exit Function
    
End Function

Function Periodo_Le_NomeExterno(ByVal lComando As Long, ByVal iExercicio As Integer, ByVal iPeriodo As Integer, sNomeExterno As String) As Long
'le o nome externo referente ao periodo em questão

Dim lErro As Long

On Error GoTo Erro_Periodo_Le_NomeExterno

    sNomeExterno = String(STRING_PERIODO_NOME_EXTERNO, 0)
    
    lErro = Comando_Executar(lComando, "SELECT NomeExterno FROM Periodo WHERE Exercicio=? AND Periodo=?", sNomeExterno, iExercicio, iPeriodo)
    If lErro <> AD_SQL_SUCESSO Then Error 9263
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 9264
    
    Periodo_Le_NomeExterno = SUCESSO

    Exit Function

Erro_Periodo_Le_NomeExterno:

    Periodo_Le_NomeExterno = Err
    
    Select Case Err
    
        Case 9263, 9264
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PERIODO", Err)
            lErro = Comando_Fechar(lComando)
           
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153439)
        
    End Select
    
    Exit Function
    
End Function

Function Origem_Le_Todos(ByVal colOrigem As Collection) As Long
'le todos as origens contidas na tabela Origem e coloca-as em colOrigem

Dim lComando As Long
Dim lErro As Long
Dim objOrigem As ClassOrigemContab
Dim tOrigem As typeOrigem

On Error GoTo Erro_Origem_Le_Todos

    tOrigem.sOrigem = String(STRING_ORIGEM, 0)
    tOrigem.sDescricao = String(STRING_ORIGEM_DESCRICAO, 0)

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 9272
    
    'le todos os registros da tabela Origem
    lErro = Comando_Executar(lComando, "SELECT Origem, Descricao, Contabilidade FROM Origem", tOrigem.sOrigem, tOrigem.sDescricao, tOrigem.iContabilidade)
    If lErro <> AD_SQL_SUCESSO Then Error 9273
    
    'le a primeira origem
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9274
    
    Do While lErro <> AD_SQL_SEM_DADOS

        Set objOrigem = New ClassOrigemContab
        
        objOrigem.sOrigem = tOrigem.sOrigem
        objOrigem.sDescricao = tOrigem.sDescricao
        objOrigem.iContabilidade = tOrigem.iContabilidade
        
        'coloca a origem lida na coleção
        colOrigem.Add objOrigem
        
        'le a proxima origem
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9275

    Loop

    Call Comando_Fechar(lComando)
    
    Origem_Le_Todos = SUCESSO

    Exit Function

Erro_Origem_Le_Todos:

    Origem_Le_Todos = Err
    
    Select Case Err

        Case 9272
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 9273, 9274, 9275
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORIGEM", Err)
         
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153440)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function


Function PlanoConta_Le_Maior_Menor_Conta(sMaior As String, sMenor As String) As Long

Dim lErro As Long
Dim lComando As Long
Dim sConta As String

On Error GoTo Erro_PlanoConta_Le_maior_Menor_Conta

    sConta = String(STRING_CONTA, 0)
    
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 13093

    lErro = Comando_Executar(lComando, "SELECT Conta FROM PlanoConta ORDER BY Conta", sConta)
    If lErro <> AD_SQL_SUCESSO Then Error 13094

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then Error 13095

    sMenor = sConta

    lErro = Comando_BuscarUltimo(lComando)
    If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then Error 13096

    sMaior = sConta

    Call Comando_Fechar(lComando)

    PlanoConta_Le_Maior_Menor_Conta = SUCESSO
    
    Exit Function
    
Erro_PlanoConta_Le_maior_Menor_Conta:

    PlanoConta_Le_Maior_Menor_Conta = Err
    
    Select Case Err

        Case 13093
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 13094, 13095, 13096
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELATORIOOPCOES", Err)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153441)

    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Lote_Le(ByVal objLote As ClassLote) As Long
'le um lote a partir da chave fornecida através de objLote. Devolve os dados em objLote

Dim lComando As Long
Dim lErro As Long
Dim tLote As typeLote
    
On Error GoTo Erro_Lote_Le

    tLote.sIdOriginal = String(STRING_IDORIGINAL, 0)
    
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 9289
    
    If giSetupLotePorPeriodo = LOTE_INICIALIZADO_POR_PERIODO Then
    
        lErro = Comando_Executar(lComando, "SELECT Periodo, TotCre, TotDeb, TotInf, Status, IdOriginal, NumLancInf, NumLancAtual, NumDocInf, NumDocAtual FROM Lote WHERE FilialEmpresa = ? AND Origem = ? AND Exercicio=? AND Periodo=? AND Lote=?", tLote.iPeriodo, tLote.dTotCre, tLote.dTotDeb, tLote.dTotInf, tLote.iStatus, tLote.sIdOriginal, tLote.iNumLancInf, tLote.iNumLancAtual, tLote.iNumDocInf, tLote.iNumDocAtual, objLote.iFilialEmpresa, objLote.sOrigem, objLote.iExercicio, objLote.iPeriodo, objLote.iLote)
        If lErro <> AD_SQL_SUCESSO Then Error 9290
        
    Else
        
        lErro = Comando_Executar(lComando, "SELECT Periodo, TotCre, TotDeb, TotInf, Status, IdOriginal, NumLancInf, NumLancAtual, NumDocInf, NumDocAtual FROM Lote WHERE FilialEmpresa = ? AND Origem = ? AND Exercicio=? AND Lote=?", tLote.iPeriodo, tLote.dTotCre, tLote.dTotDeb, tLote.dTotInf, tLote.iStatus, tLote.sIdOriginal, tLote.iNumLancInf, tLote.iNumLancAtual, tLote.iNumDocInf, tLote.iNumDocAtual, objLote.iFilialEmpresa, objLote.sOrigem, objLote.iExercicio, objLote.iLote)
        If lErro <> AD_SQL_SUCESSO Then Error 9291
        
    End If
    
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9292
        
    If lErro = AD_SQL_SEM_DADOS Then Error 9293
    
    objLote.iPeriodo = tLote.iPeriodo
    objLote.dTotCre = tLote.dTotCre
    objLote.dTotDeb = tLote.dTotDeb
    objLote.dTotInf = tLote.dTotInf
    objLote.iStatus = tLote.iStatus
    objLote.sIdOriginal = tLote.sIdOriginal
    objLote.iNumLancInf = tLote.iNumLancInf
    objLote.iNumLancAtual = tLote.iNumLancAtual
    objLote.iNumDocInf = tLote.iNumDocInf
    objLote.iNumDocAtual = tLote.iNumDocAtual
    
    Call Comando_Fechar(lComando)
    
    Lote_Le = SUCESSO

    Exit Function

Erro_Lote_Le:

    Lote_Le = Err
    
    Select Case Err
    
        Case 9289
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 9290, 9291, 9292
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LOTEPENDENTE", Err, objLote.iFilialEmpresa, objLote.sOrigem, objLote.iExercicio, objLote.iPeriodo, objLote.iLote)
                    
        Case 9293
           
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153442)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function Segmento_Le_Codigo(ByVal objSegmento As ClassSegmento, ByVal colSegmentos As Collection) As Long
'le os campos (nivel , tipo , tamanho , delimitador , preenchimento) da tabela de Segmentos relativos ao codigo passado por parametro e coloca-os em colSegmentos

Dim lComando As Long
Dim lErro As Long
Dim objSeg As ClassSegmento
Dim tSegmento As typeSegmento

On Error GoTo Erro_Segmento_Le_Codigo

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 12102
    
    'inicializa as variaveis
    tSegmento.sDelimitador = Space(STRING_SEGMENTO_DELIMITADOR)
    
    'seleciona cada segmento da conta separado por nivel , tamanho , delimitador e preenchimento
    lErro = Comando_Executar(lComando, "SELECT Nivel, Tipo, Tamanho, Delimitador, Preenchimento FROM Segmento WHERE Codigo =? ", tSegmento.iNivel, tSegmento.iTipo, tSegmento.iTamanho, tSegmento.sDelimitador, tSegmento.iPreenchimento, objSegmento.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 12103
           
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 12104
        
    If lErro = AD_SQL_SEM_DADOS Then Error 12105
        
    Do While lErro = AD_SQL_SUCESSO
    
        Set objSeg = New ClassSegmento
        
        objSeg.sCodigo = objSegmento.sCodigo
        objSeg.iNivel = tSegmento.iNivel
        objSeg.iTipo = tSegmento.iTipo
        objSeg.iTamanho = tSegmento.iTamanho
        objSeg.sDelimitador = tSegmento.sDelimitador
        objSeg.iPreenchimento = tSegmento.iPreenchimento
        
        colSegmentos.Add objSeg
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 12106
        
    Loop
    
    Call Comando_Fechar(lComando)
    
    Segmento_Le_Codigo = SUCESSO
    
    Exit Function
 
Erro_Segmento_Le_Codigo:

    Segmento_Le_Codigo = Err
    
    Select Case Err
    
        Case 12102
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 12103, 12104, 12106
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SEGMENTO", Err)
               
        Case 12105
            'nao achou nenhum dado na tabela , a funcao vai retornar AD_SQL_SEM_DADOS
            Segmento_Le_Codigo = AD_SQL_SEM_DADOS
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153443)
    
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function PlanoConta_ExisteConta(iExiste As Integer) As Long
'Verifica se existe pelo menos uma conta cadastrada em PlanoContano BD
'se existir iExiste = POSSUI_CONTA , caso contrario iExiste = NAO_POSSUI_CONTA

Dim lComando As Long
Dim lErro As Long
Dim sConta As String

On Error GoTo Erro_PlanoConta_ExisteConta
   
    sConta = String(STRING_CONTA, 0)
    
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 12107
    
    lErro = Comando_Executar(lComando, "SELECT conta FROM PlanoConta", sConta)
    If lErro <> AD_SQL_SUCESSO Then Error 12108
           
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 12109
        
    If lErro = AD_SQL_SUCESSO Then
        iExiste = POSSUI_CONTA
    Else
        iExiste = NAO_POSSUI_CONTA
    End If
        
    Call Comando_Fechar(lComando)
    
    PlanoConta_ExisteConta = SUCESSO
    
    Exit Function
 
Erro_PlanoConta_ExisteConta:

    PlanoConta_ExisteConta = Err
    
    Select Case Err
    
        Case 12107
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 12108, 12109
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTA", Err)
               
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153444)
    
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function Exercicios_Le_Todos(ByVal colExercicios As Collection) As Long
'le  todos os exercicios  e coloca-os em colExercicios

Dim lComando As Long
Dim lErro As Long
Dim tExercicio As typeExercicio
Dim objExercicio As ClassExercicio
    
On Error GoTo Erro_Exercicios_Le_Todos

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 9381
    
    tExercicio.sNomeExterno = String(STRING_EXERCICIO_NOME_EXTERNO, 0)

    lErro = Comando_Executar(lComando, "SELECT Exercicio, NumPeriodos, Status, NomeExterno, DataInicio, DataFim FROM Exercicios ORDER BY Exercicio", tExercicio.iExercicio, tExercicio.iNumPeriodos, tExercicio.iStatus, tExercicio.sNomeExterno, tExercicio.dtDataInicio, tExercicio.dtDataFim)
    If lErro <> AD_SQL_SUCESSO Then Error 9382
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9383
    
    Do While lErro = AD_SQL_SUCESSO
    
        Set objExercicio = New ClassExercicio
    
        objExercicio.iExercicio = tExercicio.iExercicio
        objExercicio.iNumPeriodos = tExercicio.iNumPeriodos
        objExercicio.iStatus = tExercicio.iStatus
        objExercicio.sNomeExterno = tExercicio.sNomeExterno
        objExercicio.dtDataInicio = tExercicio.dtDataInicio
        objExercicio.dtDataFim = tExercicio.dtDataFim

        colExercicios.Add objExercicio
        
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9384
        
    Loop
        
    Call Comando_Fechar(lComando)
    
    Exercicios_Le_Todos = SUCESSO
    
    Exit Function
    
    
Erro_Exercicios_Le_Todos:

    Exercicios_Le_Todos = Err
    
    Select Case Err
    
        Case 9381
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 9382, 9383, 9384
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EXERCICIO1", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153445)
        
    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Periodo_Le_Todos(ByVal colPeriodos As Collection) As Long
'le todos os periodos de todos os exercicios e coloca-os em colPeriodos

Dim lComando As Long
Dim lErro As Long
Dim objPeriodo As ClassPeriodo
Dim tPeriodo As typePeriodo

On Error GoTo Erro_Periodo_Le_Todos

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 9385

    tPeriodo.sNomeExterno = String(STRING_PERIODO_NOME_EXTERNO, 0)

    'Pesquisa os periodos do exercicio em questão
    lErro = Comando_Executar(lComando, "SELECT Exercicio, Periodo, NomeExterno, DataInicio, DataFim FROM Periodo", tPeriodo.iExercicio, tPeriodo.iPeriodo, tPeriodo.sNomeExterno, tPeriodo.dtDataInicio, tPeriodo.dtDataFim)
    If lErro <> AD_SQL_SUCESSO Then Error 9386
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9387
    
    Do While lErro = AD_SQL_SUCESSO

        Set objPeriodo = New ClassPeriodo
        
        objPeriodo.iExercicio = tPeriodo.iExercicio
        objPeriodo.iPeriodo = tPeriodo.iPeriodo
        objPeriodo.sNomeExterno = tPeriodo.sNomeExterno
        objPeriodo.dtDataInicio = tPeriodo.dtDataInicio
        objPeriodo.dtDataFim = tPeriodo.dtDataFim
        
        colPeriodos.Add objPeriodo
        
        'le o proximo periodo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9388
                
    Loop
    
    Call Comando_Fechar(lComando)
    
    Periodo_Le_Todos = SUCESSO
    
    Exit Function
    
Erro_Periodo_Le_Todos:

    Periodo_Le_Todos = Err
    
    Select Case Err
    
        Case 9385
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 9386, 9387, 9388
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PERIODOS", Err)
           
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153446)
        
    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function
    
End Function

Function Orcamento_Le(ByVal objOrcamento As ClassOrcamento, ByVal colOrcamento As Collection) As Long
'le os registros da tabela Orcamento que satisfazem as condicoes passadas como parametro e coloca-os em colOrcamento

Dim lComando As Long
Dim lErro As Long
Dim tOrcamento As typeOrcamento
Dim objOrcamento1 As ClassOrcamento

On Error GoTo Erro_Orcamento_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 16015

    tOrcamento.sCcl = String(STRING_CCL, 0)

    lErro = Comando_Executar(lComando, "SELECT  Periodo, Valor, Ccl FROM Orcamento WHERE FilialEmpresa = ? AND Exercicio=? AND Conta=? AND Ccl=?", tOrcamento.iPeriodo, tOrcamento.dValor, tOrcamento.sCcl, objOrcamento.iFilialEmpresa, objOrcamento.iExercicio, objOrcamento.sConta, objOrcamento.sCcl)
    If lErro <> AD_SQL_SUCESSO Then Error 16016

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 16017

    If lErro = AD_SQL_SEM_DADOS Then Error 16019

    Do While lErro <> AD_SQL_SEM_DADOS

        Set objOrcamento1 = New ClassOrcamento

        objOrcamento1.iFilialEmpresa = objOrcamento.iFilialEmpresa
        objOrcamento1.sConta = objOrcamento.sConta
        objOrcamento1.iExercicio = objOrcamento.iExercicio
        objOrcamento1.sCcl = objOrcamento.sCcl
        objOrcamento1.iPeriodo = tOrcamento.iPeriodo
        objOrcamento1.dValor = tOrcamento.dValor
        
        colOrcamento.Add objOrcamento1

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 16018
        
    Loop
    
    Call Comando_Fechar(lComando)
    
    Orcamento_Le = SUCESSO
    
    Exit Function
 
Erro_Orcamento_Le:

    Orcamento_Le = Err
    
    Select Case Err
    
        Case 16015
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
              
        Case 16016, 16017, 16018
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORCAMENTO", Err)
          
        Case 16019
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153447)
    
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function LanPendente_Le_Doc1(ByVal objDoc As ClassLancamento_Cabecalho) As Long
'Tenta ler o documento em questão.

Dim lComando1 As Long
Dim lErro As Long
Dim iLote As Integer
Dim iPeriodo As Integer

On Error GoTo Erro_LanPendente_Le_Doc1

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 9591

    'verifica se este número ainda não foi utilizado
    If giSetupDocPorPeriodo = DOC_INICIALIZADO_POR_PERIODO Then
    
        lErro = Comando_Executar(lComando1, "SELECT PeriodoLan, Lote FROM LanPendente WHERE FilialEmpresa = ? AND Origem = ? AND Exercicio = ? AND PeriodoLan = ? AND Doc = ? ORDER BY Doc", iPeriodo, iLote, objDoc.iFilialEmpresa, objDoc.sOrigem, objDoc.iExercicio, objDoc.iPeriodoLan, objDoc.lDoc)
        If lErro <> AD_SQL_SUCESSO Then Error 9592
        
    Else
    
        lErro = Comando_Executar(lComando1, "SELECT PeriodoLan, Lote FROM LanPendente WHERE FilialEmpresa = ? AND Origem = ? AND Exercicio = ? AND Doc = ? ORDER BY Doc", iPeriodo, iLote, objDoc.iFilialEmpresa, objDoc.sOrigem, objDoc.iExercicio, objDoc.lDoc)
        If lErro <> AD_SQL_SUCESSO Then Error 9593
        
    End If
        
    'le o primeiro lancamento
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9589

    'se encontrou um documento com o mesmo numero e outro periodo ==> erro
    If lErro = SUCESSO And iPeriodo <> objDoc.iPeriodoLan Then Error 20770
    
    If lErro = AD_SQL_SEM_DADOS Then Error 9590

    objDoc.iLote = iLote

    Call Comando_Fechar(lComando1)
    
    LanPendente_Le_Doc1 = SUCESSO
        
    Exit Function
    
    
Erro_LanPendente_Le_Doc1:

    LanPendente_Le_Doc1 = Err
    
    Select Case Err
    
        Case 9589, 9592, 9593
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LANPENDENTE3", Err)
           
        Case 9590
               
        Case 9591
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 20770
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DOC_PENDENTE_OUTRO_PERIODO", Err, objDoc.iFilialEmpresa, objDoc.sOrigem, objDoc.iExercicio, iPeriodo, objDoc.lDoc)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153448)
        
    End Select
    
    Call Comando_Fechar(lComando1)

    Exit Function
    
End Function

Function Exercicio_Le_Prim_Aberto(ByVal objExercicio As ClassExercicio) As Long

Dim lErro As Long
Dim lComando As Long
Dim sNomeExterno As String
Dim iExercicio As Integer
Dim iStatus As Integer

On Error GoTo Erro_Exercicio_Le_Prim_Aberto

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 11709
    
    sNomeExterno = String(STRING_NOME_EXTERNO, 0)
    
    lErro = Comando_Executar(lComando, "SELECT Exercicio, NomeExterno, Status FROM Exercicios WHERE Status <> ? ORDER BY Exercicio", iExercicio, sNomeExterno, iStatus, EXERCICIO_FECHADO)
    If lErro <> AD_SQL_SUCESSO Then Error 11711
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 11712
    
    If lErro = AD_SQL_SEM_DADOS Then Error 11713
                
    objExercicio.sNomeExterno = sNomeExterno
    objExercicio.iExercicio = iExercicio
    objExercicio.iStatus = iStatus
               
    Call Comando_Fechar(lComando)
    
    Exercicio_Le_Prim_Aberto = SUCESSO
    
    Exit Function
    
Erro_Exercicio_Le_Prim_Aberto:

    Exercicio_Le_Prim_Aberto = Err

    Select Case Err
    
        Case 11709
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 11711, 11712
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EXERCICIOS", Err)
        
        Case 11713
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TODOS_EXERCICIOS_FECHADOS", Err)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153449)
            
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function Exercicio_Le_Ultimo_Fechado(ByVal objExercicio As ClassExercicio) As Long

Dim lErro As Long
Dim lComando As Long
Dim sNomeExterno As String
Dim iExercicio As Integer

On Error GoTo Erro_Exercicio_Le_Ultimo_Fechado

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 11744
    
    sNomeExterno = String(STRING_NOME_EXTERNO, 0)
    
    lErro = Comando_Executar(lComando, "SELECT Exercicio, NomeExterno FROM Exercicios WHERE Status = ?  ORDER BY Exercicio DESC", iExercicio, sNomeExterno, EXERCICIO_FECHADO)
    If lErro <> AD_SQL_SUCESSO Then Error 11778
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 11712
    
    If lErro = AD_SQL_SEM_DADOS Then Error 11779
        
    objExercicio.sNomeExterno = sNomeExterno
    objExercicio.iExercicio = iExercicio
                   
    Call Comando_Fechar(lComando)
    
    Exercicio_Le_Ultimo_Fechado = SUCESSO
    
    Exit Function
    
Erro_Exercicio_Le_Ultimo_Fechado:

    Exercicio_Le_Ultimo_Fechado = Err

    Select Case Err
    
        Case 11744
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 11777
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
                    
        Case 11778, 11712
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EXERCICIOS", Err)
        
        Case 11779
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TODOS_EXERCICIOS_ABERTOS", Err)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153450)
            
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function LotePendente_Le_Desatualizados(ByVal iFilialEmpresa As Integer, ByVal colLote As Collection) As Long
'Le todos os lotes com status = LOTE_DESATUALIZADO e IdAtualizacao = 0

Dim lComando As Long
Dim lErro As Long
Dim objLote As ClassLote
Dim tLote As typeLote

On Error GoTo Erro_LotePendente_Le_Desatualizados
   
    tLote.sOrigem = String(STRING_ORIGEM, 0)
    tLote.sIdOriginal = String(STRING_IDORIGINAL, 0)
     
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 12184
    
    lErro = Comando_Executar(lComando, "SELECT Origem , Exercicio, Periodo, Lote, IdAtualizacao, NumDocInf, TotInf FROM LotePendente WHERE FilialEmpresa = ? AND Status = ? AND IdAtualizacao = ? ORDER BY Origem, Exercicio, Periodo, Lote ", tLote.sOrigem, tLote.iExercicio, tLote.iPeriodo, tLote.iLote, tLote.iIdAtualizacao, tLote.iNumDocInf, tLote.dTotInf, iFilialEmpresa, LOTE_DESATUALIZADO, 0)
    If lErro <> AD_SQL_SUCESSO Then Error 12185
        
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 12186
        
    Do While lErro <> AD_SQL_SEM_DADOS

        Set objLote = New ClassLote
        
            objLote.iFilialEmpresa = iFilialEmpresa
            objLote.sOrigem = tLote.sOrigem
            objLote.iExercicio = tLote.iExercicio
            objLote.iPeriodo = tLote.iPeriodo
            objLote.iLote = tLote.iLote
            objLote.iIdAtualizacao = tLote.iIdAtualizacao
            objLote.iNumDocInf = tLote.iNumDocInf
            objLote.dTotInf = tLote.dTotInf
            
            colLote.Add objLote

            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 12188
        
    Loop
    
    Call Comando_Fechar(lComando)
    
    LotePendente_Le_Desatualizados = SUCESSO

Exit Function

Erro_LotePendente_Le_Desatualizados:

    LotePendente_Le_Desatualizados = Err

    Select Case Err
    
        Case 12184
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 12185, 12186, 12188
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LOTEPENDENTE1", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153451)

    End Select
    
    Call Comando_Fechar(lComando)
        
    Exit Function

End Function

Function LotePendente_Le_DesatualizadoII(ByVal iFilialEmpresa As Integer, ByVal colLote As Collection) As Long
'Le todos os lotes com status = LOTE_DESATUALIZADO

Dim lComando As Long
Dim lErro As Long
Dim objLote As ClassLote
Dim tLote As typeLote

On Error GoTo Erro_LotePendente_Le_DesatualizadoII

    tLote.sOrigem = String(STRING_ORIGEM, 0)
    tLote.sIdOriginal = String(STRING_IDORIGINAL, 0)
    
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 12206
    
    lErro = Comando_Executar(lComando, "SELECT Origem , Exercicio, Periodo, Lote, IdAtualizacao, TotInf, NumDocInf FROM LotePendente WHERE FilialEmpresa = ? AND Status = ? ORDER BY Origem, Exercicio, Periodo, Lote ", tLote.sOrigem, tLote.iExercicio, tLote.iPeriodo, tLote.iLote, tLote.iIdAtualizacao, tLote.dTotInf, tLote.iNumDocInf, iFilialEmpresa, LOTE_DESATUALIZADO)
    If lErro <> AD_SQL_SUCESSO Then Error 12207
        
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 12208
     
    Do While lErro <> AD_SQL_SEM_DADOS

        Set objLote = New ClassLote
        
            objLote.iFilialEmpresa = iFilialEmpresa
            objLote.sOrigem = tLote.sOrigem
            objLote.iExercicio = tLote.iExercicio
            objLote.iPeriodo = tLote.iPeriodo
            objLote.iLote = tLote.iLote
            objLote.iIdAtualizacao = tLote.iIdAtualizacao
            objLote.iNumDocInf = tLote.iNumDocInf
            objLote.dTotInf = tLote.dTotInf

            colLote.Add objLote

            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 12212
        
    Loop
    
    Call Comando_Fechar(lComando)
    
    LotePendente_Le_DesatualizadoII = SUCESSO

Exit Function

Erro_LotePendente_Le_DesatualizadoII:

    LotePendente_Le_DesatualizadoII = Err

    Select Case Err
    
        Case 12206
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 12207, 12208, 12212
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LOTEPENDENTE1", Err)
                                                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153452)

    End Select
    
    Call Comando_Fechar(lComando)
   
    Exit Function

End Function

Function ContaCategoria_Le(ByVal objContaCategoria As ClassContaCategoria) As Long
'Le a categoria da conta a partir da chave fornecida em objContaCategoria. Retorna os dados em objContaCategoria

Dim lComando As Long
Dim lErro As Long
Dim iApuracao As Integer
Dim sNome As String
    
On Error GoTo Erro_ContaCategoria_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 9648
    
    sNome = String(STRING_CONTA_CATEGORIA_NOME, 0)
    
    lErro = Comando_Executar(lComando, "SELECT Nome, Apuracao FROM ContaCategoria WHERE Codigo = ?", sNome, iApuracao, objContaCategoria.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 9649
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9650
        
    If lErro = AD_SQL_SEM_DADOS Then Error 9651
    
    objContaCategoria.iApuracao = iApuracao
    objContaCategoria.sNome = sNome
    
    Call Comando_Fechar(lComando)
    
    ContaCategoria_Le = SUCESSO

    Exit Function

Erro_ContaCategoria_Le:

    ContaCategoria_Le = Err
    
    Select Case Err
    
        Case 9648
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 9649, 9650
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTACATEGORIA", Err, objContaCategoria.iCodigo)
        
        Case 9651
           
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153453)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function PlanoContas_Le_Categoria(ByVal iCodigo As Integer) As Long
'Le a categoria no plano de contas.

Dim lComando As Long
Dim lErro As Long
Dim iCodigo1 As Integer
    
On Error GoTo Erro_PlanoContas_Le_Categoria

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 9663
    
    lErro = Comando_Executar(lComando, "SELECT Categoria FROM PlanoConta WHERE Categoria = ?", iCodigo1, iCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 9664
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9665
        
    If lErro = AD_SQL_SEM_DADOS Then Error 9666
    
    Call Comando_Fechar(lComando)
    
    PlanoContas_Le_Categoria = SUCESSO

    Exit Function

Erro_PlanoContas_Le_Categoria:

    PlanoContas_Le_Categoria = Err
    
    Select Case Err
    
        Case 9663
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 9664, 9665
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTA", Err)

        Case 9666
   
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153454)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function ContaCategoria_Le_Todos(ByVal colContaCategoria As Collection) As Long
'le todos as categorias contidas na tabela de categorias e coloca-as em colContaCategoria

Dim lComando As Long
Dim lErro As Long
Dim objContaCategoria As ClassContaCategoria
Dim iApuracao As Integer
Dim sNome As String
Dim iCodigo As Integer
    
On Error GoTo Erro_ContaCategoria_Le_Todos

    sNome = String(STRING_CONTA_CATEGORIA_NOME, 0)

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 9675

    'le a tabela ContaCategoria
    lErro = Comando_Executar(lComando, "SELECT Codigo, Nome, Apuracao FROM ContaCategoria ORDER BY Codigo", iCodigo, sNome, iApuracao)
    If lErro <> AD_SQL_SUCESSO Then Error 9676

    'le a primeira categoria
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9677
    
    Do While lErro <> AD_SQL_SEM_DADOS

        Set objContaCategoria = New ClassContaCategoria
        
        objContaCategoria.iCodigo = iCodigo
        objContaCategoria.sNome = sNome
        objContaCategoria.iApuracao = iApuracao
        
        'coloca a categoria lida na coleção
        colContaCategoria.Add objContaCategoria
        
        'le a proxima categoria
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9678

    Loop

    Call Comando_Fechar(lComando)
    
    ContaCategoria_Le_Todos = SUCESSO

    Exit Function

Erro_ContaCategoria_Le_Todos:

    ContaCategoria_Le_Todos = Err
    
    Select Case Err

        Case 9675
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 9676, 9677, 9678
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTACATEGORIA1", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153455)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function PlanoConta_Le_Categoria_Nivel(ByVal iCategoria As Integer, ByVal iNivelConta As Integer, ByVal objPlanoConta As ClassPlanoConta) As Long
'Le a conta na categoria e no nivel especificados

Dim lComando As Long
Dim lErro As Long
Dim tPlanoConta As typePlanoConta

On Error GoTo Erro_PlanoConta_Le_Categoria_Nivel

    tPlanoConta.sConta = String(STRING_CONTA, 0)
    tPlanoConta.sDescConta = String(STRING_CONTA_DESCRICAO, 0)

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 9724

    lErro = Comando_Executar(lComando, "SELECT Conta, NivelConta, DescConta, TipoConta, ContaSimples, Natureza, DigitoVerif, DigitoVerifSimples, HistPadrao, Ativo, UsaContaSimples, Categoria, FluxoCaixa, NaturezaSped FROM PlanoConta WHERE Categoria=? AND NivelConta=?", tPlanoConta.sConta, tPlanoConta.iNivelConta, tPlanoConta.sDescConta, tPlanoConta.iTipoConta, tPlanoConta.lContaSimples, tPlanoConta.iNatureza, tPlanoConta.iDigitoVerif, tPlanoConta.iDigitoVerifSimples, tPlanoConta.iHistPadrao, tPlanoConta.iAtivo, tPlanoConta.iUsaContaSimples, tPlanoConta.iCategoria, tPlanoConta.iFluxoCaixa, tPlanoConta.iNaturezaSped, iCategoria, iNivelConta)
    If lErro <> AD_SQL_SUCESSO Then Error 9725
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9727
        
    If lErro = AD_SQL_SEM_DADOS Then Error 9726
        
    objPlanoConta.sConta = tPlanoConta.sConta
    objPlanoConta.iNivelConta = tPlanoConta.iNivelConta
    objPlanoConta.sDescConta = tPlanoConta.sDescConta
    objPlanoConta.iTipoConta = tPlanoConta.iTipoConta
    objPlanoConta.lContaSimples = tPlanoConta.lContaSimples
    objPlanoConta.iNatureza = tPlanoConta.iNatureza
    objPlanoConta.iDigitoVerif = tPlanoConta.iDigitoVerif
    objPlanoConta.iDigitoVerifSimples = tPlanoConta.iDigitoVerifSimples
    objPlanoConta.iHistPadrao = tPlanoConta.iHistPadrao
    objPlanoConta.iAtivo = tPlanoConta.iAtivo
    objPlanoConta.iUsaContaSimples = tPlanoConta.iUsaContaSimples
    objPlanoConta.dSldIni = tPlanoConta.dSldIni
    objPlanoConta.iCategoria = tPlanoConta.iCategoria
    objPlanoConta.iFluxoCaixa = tPlanoConta.iFluxoCaixa
    objPlanoConta.iNaturezaSped = tPlanoConta.iNaturezaSped
    
    Call Comando_Fechar(lComando)
    
    PlanoConta_Le_Categoria_Nivel = SUCESSO
    
    Exit Function
 
Erro_PlanoConta_Le_Categoria_Nivel:

    PlanoConta_Le_Categoria_Nivel = Err
    
    Select Case Err
    
        Case 9724
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 9725, 9727
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTA4", Err, iCategoria, iNivelConta)
    
        Case 9726
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153456)
    
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function ContaCategoria_Le_Nome(ByVal sNome As String, ByVal objContaCategoria As ClassContaCategoria) As Long
'Le a categoria da conta que possui o sNome especificado. Retorna os dados em objContaCategoria

Dim lComando As Long
Dim lErro As Long
Dim iApuracao As Integer
Dim iCodigo As Integer
    
On Error GoTo Erro_ContaCategoria_Le_Nome

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 9729
    
    lErro = Comando_Executar(lComando, "SELECT Codigo, Apuracao FROM ContaCategoria WHERE Nome = ?", iCodigo, iApuracao, sNome)
    If lErro <> AD_SQL_SUCESSO Then Error 9730
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9731
        
    If lErro = AD_SQL_SEM_DADOS Then Error 9732
    
    objContaCategoria.iApuracao = iApuracao
    objContaCategoria.sNome = sNome
    objContaCategoria.iCodigo = iCodigo
    
    Call Comando_Fechar(lComando)
    
    ContaCategoria_Le_Nome = SUCESSO

    Exit Function

Erro_ContaCategoria_Le_Nome:

    ContaCategoria_Le_Nome = Err
    
    Select Case Err
    
        Case 9729
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 9730, 9731
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTACATEGORIA2", Err, sNome)
        
        Case 9732
           
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153457)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function PlanoConta_Le_Todas_Categorias(ByVal colPlanoConta As Collection) As Long
'le todas as contas que possuam categoria no plano de contas

Dim lComando As Long
Dim lErro As Long
Dim objPlanoConta As ClassPlanoConta
Dim tPlanoConta As typePlanoConta

On Error GoTo Erro_PlanoConta_Le_Todas_Categorias

    tPlanoConta.sConta = String(STRING_CONTA, 0)
    tPlanoConta.sDescConta = String(STRING_CONTA_DESCRICAO, 0)

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 9745

    'le a tabela de plano de Contas selecionando as contas que possuam categoria
    lErro = Comando_Executar(lComando, "SELECT Conta, NivelConta, DescConta, TipoConta, ContaSimples, Natureza, DigitoVerif, DigitoVerifSimples, HistPadrao, Ativo, UsaContaSimples, Categoria, FluxoCaixa FROM PlanoConta WHERE Categoria > 0", tPlanoConta.sConta, tPlanoConta.iNivelConta, tPlanoConta.sDescConta, tPlanoConta.iTipoConta, tPlanoConta.lContaSimples, tPlanoConta.iNatureza, tPlanoConta.iDigitoVerif, tPlanoConta.iDigitoVerifSimples, tPlanoConta.iHistPadrao, tPlanoConta.iAtivo, tPlanoConta.iUsaContaSimples, tPlanoConta.iCategoria, tPlanoConta.iFluxoCaixa)
    If lErro <> AD_SQL_SUCESSO Then Error 9746
    
    'le a primeira conta
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9747
    
    Do While lErro <> AD_SQL_SEM_DADOS

        Set objPlanoConta = New ClassPlanoConta
        
        objPlanoConta.sConta = tPlanoConta.sConta
        objPlanoConta.iNivelConta = tPlanoConta.iNivelConta
        objPlanoConta.sDescConta = tPlanoConta.sDescConta
        objPlanoConta.iTipoConta = tPlanoConta.iTipoConta
        objPlanoConta.lContaSimples = tPlanoConta.lContaSimples
        objPlanoConta.iNatureza = tPlanoConta.iNatureza
        objPlanoConta.iDigitoVerif = tPlanoConta.iDigitoVerif
        objPlanoConta.iDigitoVerifSimples = tPlanoConta.iDigitoVerifSimples
        objPlanoConta.iHistPadrao = tPlanoConta.iHistPadrao
        objPlanoConta.iAtivo = tPlanoConta.iAtivo
        objPlanoConta.iUsaContaSimples = tPlanoConta.iUsaContaSimples
        objPlanoConta.dSldIni = tPlanoConta.dSldIni
        objPlanoConta.iCategoria = tPlanoConta.iCategoria
        objPlanoConta.iFluxoCaixa = tPlanoConta.iFluxoCaixa
        
        'coloca a conta lida na coleção
        colPlanoConta.Add objPlanoConta
        
        'le a proxima conta
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9748

    Loop

    Call Comando_Fechar(lComando)
    
    PlanoConta_Le_Todas_Categorias = SUCESSO

    Exit Function

Erro_PlanoConta_Le_Todas_Categorias:

    PlanoConta_Le_Todas_Categorias = Err
    
    Select Case Err

        Case 9745
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 9746, 9747, 9748
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTA", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153458)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function CTBConfig_Le(ByVal objCTBConfig As ClassCTBConfig) As Long
'Le o registro de configuração especificado no codigo de objCTBConfig e retorna os demais dados em objCTBConfig

Dim lComando As Long
Dim lErro As Long
Dim tCTBConfig As typeCTBConfig
    
On Error GoTo Erro_CTBConfig_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 9752
    
    tCTBConfig.sDescricao = String(STRING_CTBCONFIG_DESCRICAO, 0)
    tCTBConfig.sConteudo = String(STRING_CTBCONFIG_CONTEUDO, 0)
    
    lErro = Comando_Executar(lComando, "SELECT Descricao, Tipo, Conteudo FROM CTBConfig WHERE Codigo = ? AND FilialEmpresa = ?", tCTBConfig.sDescricao, tCTBConfig.iTipo, tCTBConfig.sConteudo, objCTBConfig.sCodigo, objCTBConfig.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then Error 9753
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9754
        
    If lErro = AD_SQL_SEM_DADOS Then Error 9755
    
    objCTBConfig.sDescricao = tCTBConfig.sDescricao
    objCTBConfig.iTipo = tCTBConfig.iTipo
    objCTBConfig.sConteudo = tCTBConfig.sConteudo
    
    Call Comando_Fechar(lComando)
    
    CTBConfig_Le = SUCESSO

    Exit Function

Erro_CTBConfig_Le:

    CTBConfig_Le = Err
    
    Select Case Err
    
        Case 9752
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 9753, 9754
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CTBCONFIG", Err, objCTBConfig.sCodigo)
        
        Case 9755
           
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153459)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function LotePendente_Le_NumRegistros(lNumReg As Long, ByVal lComando As Long) As Long
'retorna o número de registros contidos na tabela de lotes pendentes.

Dim lErro As Long

On Error GoTo Erro_LotePendente_Le_NumRegistros

    If lComando = 0 Then Error 9008

    'Le o número de registro na tabela de lotes pendentes
    lErro = Comando_Executar(lComando, "SELECT COUNT(*) FROM LotePendente", lNumReg)
    If lErro <> AD_SQL_SUCESSO Then Error 9009

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 9010
    
    LotePendente_Le_NumRegistros = SUCESSO

    Exit Function

Erro_LotePendente_Le_NumRegistros:

    LotePendente_Le_NumRegistros = Err
    
    Select Case Err

        Case 9008
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 9009, 9010
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LOTEPENDENTE1", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153460)

    End Select
    
    Exit Function

End Function

Function Exercicio_Le_Primeiro(ByVal objExercicio As ClassExercicio) As Long
'retorna o primeiro exercicio do sistema

Dim lErro As Long
Dim lComando As Long
Dim sNomeExterno As String
Dim iStatus As Integer
Dim tExercicio As typeExercicio

On Error GoTo Erro_Exercicio_Le_Primeiro

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 9844
    
    tExercicio.sNomeExterno = String(STRING_NOME_EXTERNO, 0)
    
    lErro = Comando_Executar(lComando, "SELECT Exercicio, NumPeriodos, Status, NomeExterno, DataInicio, DataFim FROM Exercicios ORDER BY Exercicio", tExercicio.iExercicio, tExercicio.iNumPeriodos, tExercicio.iStatus, tExercicio.sNomeExterno, tExercicio.dtDataInicio, tExercicio.dtDataFim)
    If lErro <> AD_SQL_SUCESSO Then Error 9845
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9846
    
    If lErro = AD_SQL_SEM_DADOS Then Error 9847
                
    objExercicio.iExercicio = tExercicio.iExercicio
    objExercicio.iNumPeriodos = tExercicio.iNumPeriodos
    objExercicio.iStatus = tExercicio.iStatus
    objExercicio.iLoteApuracao = tExercicio.iLoteApuracao
    objExercicio.sNomeExterno = tExercicio.sNomeExterno
    objExercicio.dtDataInicio = tExercicio.dtDataInicio
    objExercicio.dtDataFim = tExercicio.dtDataFim
               
    Call Comando_Fechar(lComando)
    
    Exercicio_Le_Primeiro = SUCESSO
    
    Exit Function
    
Erro_Exercicio_Le_Primeiro:

    Exercicio_Le_Primeiro = Err

    Select Case Err
    
        Case 9844
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 9845, 9846
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EXERCICIO1", Err)
        
        Case 9847
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153461)
            
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function LanPendente_Le_ContaCcl(ByVal sConta As String, ByVal sCcl As String) As Long
'verifica se existe no BD pelo menos um lancamento pendente para a conta e centro de custo/lucro fornecido

Dim lComando As Long
Dim lErro As Long
Dim lDoc As Long

On Error GoTo Erro_LanPendente_Le_ContaCcl

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 9891
    
    lErro = Comando_Executar(lComando, "SELECT Doc FROM LanPendente WHERE Ccl=? And Conta=?", lDoc, sCcl, sConta)
    If lErro <> AD_SQL_SUCESSO Then Error 9892
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9893
        
    If lErro = AD_SQL_SEM_DADOS Then Error 9894
        
    Call Comando_Fechar(lComando)
    
    LanPendente_Le_ContaCcl = SUCESSO
    
    Exit Function
 
Erro_LanPendente_Le_ContaCcl:

    LanPendente_Le_ContaCcl = Err
    
    Select Case Err
    
        Case 9891
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 9892, 9893
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LANPENDENTE4", Err, sConta, sCcl)
            
        Case 9894
          
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153462)
    
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function RateioOn_Le_ContaCcl(ByVal sConta As String, ByVal sCcl As String) As Long
'verifica se existe na tabela RateioOn pelo menos um rateio com a conta e centro de custo em questão.

Dim lComando As Long
Dim lErro As Long
Dim iCodigo As Integer
Dim iSeq As Integer

On Error GoTo Erro_RateioOn_Le_ContaCcl

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 9900
    
    lErro = Comando_Executar(lComando, "SELECT Codigo, Seq FROM RateioOn WHERE Conta=? And Ccl=?", iCodigo, iSeq, sConta, sCcl)
    If lErro <> AD_SQL_SUCESSO Then Error 9901
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9902
        
    If lErro = AD_SQL_SEM_DADOS Then Error 9903
        
    Call Comando_Fechar(lComando)
    
    RateioOn_Le_ContaCcl = SUCESSO
    
    Exit Function
 
Erro_RateioOn_Le_ContaCcl:

    RateioOn_Le_ContaCcl = Err
    
    Select Case Err
    
        Case 9900
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 9901, 9902
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RATEIOON1", Err, sConta, sCcl)
           
        Case 9903
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153463)
    
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function Exercicios_Nao_Fechados_Le_Todos(ByVal colExercicios As Collection) As Long
'le todos os exercicios não fechados e coloca-os em colExercicios

Dim lComando As Long
Dim lErro As Long
Dim tExercicio As typeExercicio
Dim objExercicio As ClassExercicio
    
On Error GoTo Erro_Exercicios_Nao_Fechados_Le_Todos

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 5980
    
    tExercicio.sNomeExterno = String(STRING_EXERCICIO_NOME_EXTERNO, 0)

    lErro = Comando_Executar(lComando, "SELECT Exercicio, NumPeriodos, Status, NomeExterno, DataInicio, DataFim FROM Exercicios  WHERE Status <> ? ORDER BY Exercicio", tExercicio.iExercicio, tExercicio.iNumPeriodos, tExercicio.iStatus, tExercicio.sNomeExterno, tExercicio.dtDataInicio, tExercicio.dtDataFim, EXERCICIO_FECHADO)
    If lErro <> AD_SQL_SUCESSO Then Error 5981
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 5982
    
    Do While lErro = AD_SQL_SUCESSO
    
        Set objExercicio = New ClassExercicio
    
        objExercicio.iExercicio = tExercicio.iExercicio
        objExercicio.iNumPeriodos = tExercicio.iNumPeriodos
        objExercicio.iStatus = tExercicio.iStatus
        objExercicio.sNomeExterno = tExercicio.sNomeExterno
        objExercicio.dtDataInicio = tExercicio.dtDataInicio
        objExercicio.dtDataFim = tExercicio.dtDataFim

        colExercicios.Add objExercicio
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 5983
        
    Loop
        
    Call Comando_Fechar(lComando)
    
    Exercicios_Nao_Fechados_Le_Todos = SUCESSO
    
    Exit Function
    
    
Erro_Exercicios_Nao_Fechados_Le_Todos:

    Exercicios_Nao_Fechados_Le_Todos = Err
    
    Select Case Err
    
        Case 5980
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 5981, 5982, 5983
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EXERCICIO1", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153464)
        
    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function ContaCcl_Le_Todos_Conta1(ByVal sConta As String, ByVal colContaCcl As Collection) As Long
'Le todas as associacoes da conta em questão com centros de custo/lucro e carrega-os em colContaCcl

Dim lErro As Long
Dim lComando As Long
Dim sCcl As String
Dim dSldIni As Double
Dim objContaCcl As ClassContaCcl
    
On Error GoTo Erro_ContaCcl_Le_Todos_Conta1

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 9942
    
    sCcl = String(STRING_CCL, 0)
    
    lErro = Comando_Executar(lComando, "SELECT Ccl FROM ContaCcl WHERE Conta=?", sCcl, sConta)
    If lErro <> AD_SQL_SUCESSO Then Error 9943
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9944
    
    Do While lErro = AD_SQL_SUCESSO
    
        Set objContaCcl = New ClassContaCcl
        
        objContaCcl.sConta = sConta
        objContaCcl.sCcl = sCcl
        objContaCcl.dSldIni = dSldIni
    
        colContaCcl.Add objContaCcl

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9945

    Loop
    
    Call Comando_Fechar(lComando)
    
    ContaCcl_Le_Todos_Conta1 = SUCESSO
    
    Exit Function

Erro_ContaCcl_Le_Todos_Conta1:

    ContaCcl_Le_Todos_Conta1 = Err
    
    Select Case Err
                
        Case 9942
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 9943, 9944, 9945
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTACCL1", Err, sConta)
                  
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 153465)
        
        End Select
        
        Call Comando_Fechar(lComando)
        
    Exit Function

End Function

Public Function ContaCcl_Le_Todas_Contas_Distintas(ByVal colConta As Collection) As Long
'Le todas as contas que possuem associação com centro de custo e coloca-as em colConta

Dim lErro As Long
Dim lComando As Long
Dim sConta As String
Dim sDescConta As String
Dim objPlanoConta As ClassPlanoConta
    
On Error GoTo Erro_ContaCcl_Le_Todas_Contas_Distintas

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 9949
    
    sConta = String(STRING_CONTA, 0)
    sDescConta = String(STRING_CONTA_DESCRICAO, 0)
    
    lErro = Comando_Executar(lComando, "SELECT DISTINCT ContaCcl.Conta, DescConta FROM ContaCcl, PlanoConta WHERE ContaCcl.Conta = PlanoConta.Conta ORDER BY ContaCcl.Conta", sConta, sDescConta)
    If lErro <> AD_SQL_SUCESSO Then Error 9950
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9951
    
    Do While lErro = AD_SQL_SUCESSO
    
        Set objPlanoConta = New ClassPlanoConta
    
        objPlanoConta.sConta = sConta
        objPlanoConta.sDescConta = sDescConta
    
        colConta.Add objPlanoConta

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9952

    Loop
    
    Call Comando_Fechar(lComando)
    
    ContaCcl_Le_Todas_Contas_Distintas = SUCESSO
    
    Exit Function

Erro_ContaCcl_Le_Todas_Contas_Distintas:

    ContaCcl_Le_Todas_Contas_Distintas = Err
    
    Select Case Err
                
        Case 9949
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 9950, 9951, 9952
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTACCL", Err)
                  
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 153466)
        
        End Select
        
        Call Comando_Fechar(lComando)
        
    Exit Function

End Function

Function PlanoConta_Le_Conta1(ByVal sConta As String, ByVal objPlanoConta As ClassPlanoConta) As Long
'O argumento sConta deve ser o número da conta FORMATADO BD
'le a conta passada como parametro e retorna os dados em objPlanoConta.

Dim lComando As Long
Dim lErro As Long
Dim tPlanoConta As typePlanoConta

On Error GoTo Erro_PlanoConta_Le_Conta1

    tPlanoConta.sConta = String(STRING_CONTA, 0)
    tPlanoConta.sDescConta = String(STRING_CONTA_DESCRICAO, 0)

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 6028

    lErro = Comando_Executar(lComando, "SELECT Conta, NivelConta, DescConta, TipoConta, ContaSimples, Natureza, DigitoVerif, DigitoVerifSimples, HistPadrao, Ativo, UsaContaSimples, Categoria, CP, CR, TES, FAT, EST, COM, FIS, LJ, FluxoCaixa, NaturezaSped FROM PlanoConta WHERE Conta=?", tPlanoConta.sConta, tPlanoConta.iNivelConta, tPlanoConta.sDescConta, tPlanoConta.iTipoConta, tPlanoConta.lContaSimples, tPlanoConta.iNatureza, tPlanoConta.iDigitoVerif, tPlanoConta.iDigitoVerifSimples, tPlanoConta.iHistPadrao, tPlanoConta.iAtivo, tPlanoConta.iUsaContaSimples, tPlanoConta.iCategoria, tPlanoConta.iCP, tPlanoConta.iCR, tPlanoConta.iTES, tPlanoConta.iFAT, tPlanoConta.iEST, tPlanoConta.iCOM, tPlanoConta.iFIS, tPlanoConta.iLJ, tPlanoConta.iFluxoCaixa, tPlanoConta.iNaturezaSped, sConta)
    If lErro <> AD_SQL_SUCESSO Then Error 6029
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 6031
        
    If lErro = AD_SQL_SEM_DADOS Then Error 6030
        
    objPlanoConta.sConta = tPlanoConta.sConta
    objPlanoConta.iNivelConta = tPlanoConta.iNivelConta
    objPlanoConta.sDescConta = tPlanoConta.sDescConta
    objPlanoConta.iTipoConta = tPlanoConta.iTipoConta
    objPlanoConta.lContaSimples = tPlanoConta.lContaSimples
    objPlanoConta.iNatureza = tPlanoConta.iNatureza
    objPlanoConta.iDigitoVerif = tPlanoConta.iDigitoVerif
    objPlanoConta.iDigitoVerifSimples = tPlanoConta.iDigitoVerifSimples
    objPlanoConta.iHistPadrao = tPlanoConta.iHistPadrao
    objPlanoConta.iAtivo = tPlanoConta.iAtivo
    objPlanoConta.iUsaContaSimples = tPlanoConta.iUsaContaSimples
    objPlanoConta.iCategoria = tPlanoConta.iCategoria
    objPlanoConta.iFluxoCaixa = tPlanoConta.iFluxoCaixa
    objPlanoConta.iNaturezaSped = tPlanoConta.iNaturezaSped
    
    If tPlanoConta.iCP = USA_MODULO Then objPlanoConta.colModulo.Add MODULO_CONTASAPAGAR
    If tPlanoConta.iCR = USA_MODULO Then objPlanoConta.colModulo.Add MODULO_CONTASARECEBER
    If tPlanoConta.iTES = USA_MODULO Then objPlanoConta.colModulo.Add MODULO_TESOURARIA
    If tPlanoConta.iFAT = USA_MODULO Then objPlanoConta.colModulo.Add MODULO_FATURAMENTO
    If tPlanoConta.iEST = USA_MODULO Then objPlanoConta.colModulo.Add MODULO_ESTOQUE
    If tPlanoConta.iCOM = USA_MODULO Then objPlanoConta.colModulo.Add MODULO_COMPRAS
    If tPlanoConta.iFIS = USA_MODULO Then objPlanoConta.colModulo.Add MODULO_LIVROSFISCAIS
    If tPlanoConta.iLJ = USA_MODULO Then objPlanoConta.colModulo.Add MODULO_LOJA
    
    Call Comando_Fechar(lComando)
    
    PlanoConta_Le_Conta1 = SUCESSO
    
Exit Function
 
Erro_PlanoConta_Le_Conta1:

    PlanoConta_Le_Conta1 = Err
    
    Select Case Err
    
    Case 6028
        lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
    Case 6029, 6031
        lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTA3", Err, sConta)
    
    Case 6030
                
    Case Else
        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153467)
    
    End Select
    
    Call Comando_Fechar(lComando)
      
    Exit Function

End Function

Function PlanoConta_Le_Todos_Modulo(ByVal colAnaliticas As Collection, ByVal colSinteticas As Collection, ByVal sSiglaModulo As String) As Long
'le todas as contas analíticas, visiveis para o modulo em questão, e as conta sinteticas contidas na tabela de Plano de Contas e coloca-as em colAnaliticas (contas analiticas) e colSinteticas (contas sinteticas)

Dim lComando As Long
Dim lErro As Long
Dim objPlanoConta As ClassPlanoConta
Dim tPlanoConta As typePlanoConta

On Error GoTo Erro_PlanoConta_Le_Todos_Modulo

    tPlanoConta.sConta = String(STRING_CONTA, 0)
    tPlanoConta.sDescConta = String(STRING_CONTA_DESCRICAO, 0)

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 10056
    
    'le a tabela de plano de Contas ordenado por conta
    lErro = Comando_Executar(lComando, "SELECT Conta, NivelConta, DescConta, TipoConta FROM PlanoConta WHERE (TipoConta=? AND " & sSiglaModulo & "=1) OR TipoConta=? ORDER BY Conta", tPlanoConta.sConta, tPlanoConta.iNivelConta, tPlanoConta.sDescConta, tPlanoConta.iTipoConta, CONTA_ANALITICA, CONTA_SINTETICA)
    If lErro <> AD_SQL_SUCESSO Then Error 10057
    
    'le a primeira conta analitica do modulo em questão
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 10058
    
    Do While lErro <> AD_SQL_SEM_DADOS

        Set objPlanoConta = New ClassPlanoConta
    
        objPlanoConta.sConta = tPlanoConta.sConta
        objPlanoConta.iNivelConta = tPlanoConta.iNivelConta
        objPlanoConta.sDescConta = tPlanoConta.sDescConta
        objPlanoConta.iTipoConta = tPlanoConta.iTipoConta

        If tPlanoConta.iTipoConta = CONTA_ANALITICA Then
        
            'coloca a conta lida na coleção
            colAnaliticas.Add objPlanoConta
        
        Else
        
            colSinteticas.Add objPlanoConta, tPlanoConta.sConta
            
        End If
        
        'le a proxima conta analitica do modulo em questão
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 10059

    Loop

    Call Comando_Fechar(lComando)
    
    PlanoConta_Le_Todos_Modulo = SUCESSO

    Exit Function

Erro_PlanoConta_Le_Todos_Modulo:

    PlanoConta_Le_Todos_Modulo = Err
    
    Select Case Err

        Case 10056
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 10057, 10058, 10059
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTA", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153468)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function ContasPatrimoniaisNivelUm_Le(ByVal colXContas As Collection) As Long
'preenche uma colecao com os codigos das contas patrimoniais de nivel 1
'p/evitar problemas o tipo de dados Variant vs cta que comeca por numererico vou
'colocar todas as ctas precedidas pela letra 'X'. A colecao deverá conter portanto
'elementos como "X1000000" e "X2000000"
Dim lErro As Long
Dim lComando As Long
Dim sConta As String, SXConta As String

On Error GoTo Erro_ContasPatrimoniaisNivelUm_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 7161
    
    sConta = String(STRING_CONTA, 0)
    lErro = Comando_Executar(lComando, "SELECT PlanoConta.Conta FROM PlanoConta, ContaCategoria WHERE PlanoConta.NivelConta = 1 AND ContaCategoria.Apuracao = 0 AND PlanoConta.Categoria = ContaCategoria.Codigo", sConta)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 7158
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 7159
    
    Do While lErro <> AD_SQL_SEM_DADOS

        SXConta = "X" & sConta
        
        'coloca o centro de custo lido na coleção
        colXContas.Add SXConta
        
        'le o proximo centro de custo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 7160

    Loop
    
    Call Comando_Fechar(lComando)
    
    ContasPatrimoniaisNivelUm_Le = SUCESSO
    
    Exit Function

Erro_ContasPatrimoniaisNivelUm_Le:

    ContasPatrimoniaisNivelUm_Le = Err

    Select Case Err

        Case 7161
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 7158, 7159, 7160
            '???? colocar msg de erro na leitura de plano de ctas + contas categoria
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153469)

    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Monta_Filtro_Ctas_Patrimoniais(ByVal objRelOpcoes As AdmRelOpcoes, sExprCtasPatr As String) As Long
'monta expressao do tipo "GP(@TC1,@TC2) E GP(@TC3,@TC4)"
'onde GP é uima macroa com a expressao: Conta >= $T0 E Conta <= $T1
'esta funcao vai incluir tambem os valores dos @T? que criar
'Os valores de T1, T2, T3, T4 serao algo como: "1000000", "19999999", "20000000", "29999999"
Dim colXContas As New Collection
Dim lErro As Long, iContaArrobas
Dim vXConta As Variant, sConta As String, sContaMax As String
Dim sArrobaCtaIni As String, sArrobaCtaFim As String
On Error GoTo Erro_Monta_Filtro_Ctas_Patrimoniais

    iContaArrobas = 0
    sExprCtasPatr = "("
    
    lErro = CF("ContasPatrimoniaisNivelUm_Le", colXContas)
    If lErro <> SUCESSO Then Error 7162

    For Each vXConta In colXContas  ' Create list of names.
        sConta = Mid(CStr(vXConta), 2)
        'obter maior conta possivel no grupo
        lErro = Mascara_RetornaUltimaConta(sConta, sContaMax)
            
        sArrobaCtaIni = "TC" & CStr(iContaArrobas)
        sArrobaCtaFim = "TC" & CStr(iContaArrobas + 1)
        iContaArrobas = iContaArrobas + 2
        
        If sExprCtasPatr <> "(" Then sExprCtasPatr = sExprCtasPatr & " OU "
        sExprCtasPatr = sExprCtasPatr & "GP(@" & sArrobaCtaIni & ",@" & sArrobaCtaFim & ")"
        
        lErro = objRelOpcoes.IncluirParametro(sArrobaCtaIni, sConta)
        If lErro <> AD_BOOL_TRUE Then Error 7163
        
        lErro = objRelOpcoes.IncluirParametro(sArrobaCtaFim, sContaMax)
        If lErro <> AD_BOOL_TRUE Then Error 7164
        
    Next vXConta
            
    If sExprCtasPatr <> "(" Then
        sExprCtasPatr = sExprCtasPatr & ")"
    Else
        sExprCtasPatr = ""
    End If
    
    Monta_Filtro_Ctas_Patrimoniais = SUCESSO
    
    Exit Function
    
Erro_Monta_Filtro_Ctas_Patrimoniais:

    Monta_Filtro_Ctas_Patrimoniais = Err

    Select Case Err
        
        Case 7162
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153470)

    End Select
    
    Exit Function
    
End Function

Function Orcamento_Le_Exercicio(ByVal iExercicio As Integer) As Long
'verifica se existe no BD pelo menos um orçamento para o exercício fornecido

Dim lComando As Long
Dim lErro As Long
Dim iExercicio1 As Integer

On Error GoTo Erro_Orcamento_Le_Exercicio

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 11404

    lErro = Comando_Executar(lComando, "SELECT Exercicio FROM Orcamento WHERE Exercicio = ?", iExercicio1, iExercicio)
    If lErro <> AD_SQL_SUCESSO Then Error 11405

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 11406

    If lErro = AD_SQL_SEM_DADOS Then Error 11407

    Call Comando_Fechar(lComando)

    Orcamento_Le_Exercicio = SUCESSO

    Exit Function

Erro_Orcamento_Le_Exercicio:

    Orcamento_Le_Exercicio = Err

    Select Case Err

        Case 11404
             lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 11405, 11406
             lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORCAMENTO1", Err, iExercicio)

        Case 11407

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153471)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Movimento_Le_Exercicio(ByVal iExercicio As Integer) As Long
'verifica se existe no BD movimento contábil para o exercício fornecido /m

Dim lComando As Long
Dim lComando1 As Long
Dim lErro   As Integer
Dim sConta As String
Dim dCredito(NUM_MAX_PERIODOS) As Double
Dim dDebito(NUM_MAX_PERIODOS) As Double
Dim dMovimento As Double
Dim iIndice As Integer
Dim iTemMovimento As Integer

On Error GoTo Erro_Movimento_Le_Exercicio

    sConta = String(STRING_CONTA, 0)
    
    iTemMovimento = 0

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 13600

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 13601

    'busca todas as contas da tabela PlanoConta com nível de conta = 1
    lErro = Comando_Executar(lComando, "SELECT Conta FROM PlanoConta WHERE NivelConta=1", sConta)
    If lErro <> AD_SQL_SUCESSO Then Error 13602

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 13603

    If lErro = AD_SQL_SEM_DADOS Then Error 13604

    'para conta de nível 1 da tabela PlanoConta, procurar em MvPerCta se existe algum débito ou crédito
    Do While lErro <> AD_SQL_SEM_DADOS

        lErro = Comando_Executar(lComando1, "SELECT Deb01, Deb02, Deb03, Deb04, Deb05, Deb06, Deb07, Deb08, Deb09, Deb10, Deb11, Deb12, Cre01, Cre02, Cre03, Cre04, Cre05, Cre06, Cre07, Cre08, Cre09, Cre10, Cre11, Cre12 FROM MvPerCta WHERE Exercicio=? AND Conta=?", dDebito(1), dDebito(2), dDebito(3), dDebito(4), dDebito(5), dDebito(6), dDebito(7), dDebito(8), dDebito(9), dDebito(10), dDebito(11), dDebito(12), dCredito(1), dCredito(2), dCredito(3), dCredito(4), dCredito(5), dCredito(6), dCredito(7), dCredito(8), dCredito(9), dCredito(10), dCredito(11), dCredito(12), iExercicio, sConta)
        If lErro <> SUCESSO Then Error 13605

        lErro = Comando_BuscarPrimeiro(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 13606

        'encontrou a conta
        If lErro = SUCESSO Then

            'faz o somatório dos débitos e créditos
            dMovimento = 0
            
            For iIndice = 1 To 12
                dMovimento = dMovimento + dCredito(iIndice) + dDebito(iIndice)
            Next

            'verifica se houve movimentação para a conta no exercício
            If dMovimento <> 0 Then
                iTemMovimento = 1
                Exit Do
            End If

        End If

        'próxima conta de nível = 1
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 13607

    Loop

    If iTemMovimento = 0 Then Error 13604

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Movimento_Le_Exercicio = SUCESSO
    
    Exit Function

Erro_Movimento_Le_Exercicio:

    Movimento_Le_Exercicio = Err

    Select Case Err

        Case 13600, 13601
             lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 13602, 13603, 13607
             lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTA", Err)

        Case 13604

        Case 13605, 13606
             lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVPERCTA1", Err, iExercicio, sConta)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153472)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Function LanPendente_Le_Exercicio(ByVal iExercicio As Integer) As Long
'verifica se existe no BD pelo menos um lancamento pendente para o exercício fornecido

Dim lComando As Long
Dim lErro As Long
Dim lDoc As Long

On Error GoTo Erro_LanPendente_Le_Exercicio

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 13608

    lErro = Comando_Executar(lComando, "SELECT Doc FROM LanPendente WHERE Exercicio=?", lDoc, iExercicio)
    If lErro <> AD_SQL_SUCESSO Then Error 13609

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 13610

    If lErro = AD_SQL_SEM_DADOS Then Error 13611

    Call Comando_Fechar(lComando)

    LanPendente_Le_Exercicio = SUCESSO

    Exit Function

Erro_LanPendente_Le_Exercicio:

    LanPendente_Le_Exercicio = Err

    Select Case Err

        Case 13608
             lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 13609, 13610
             lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LANPENDENTE5", Err, iExercicio)

        Case 13611

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153473)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function LotePendente_Le_Exercicio(ByVal iExercicio As Integer) As Long
'verifica se existe no BD pelo menos um lote pendente para o exercício fornecido

Dim lComando As Long
Dim lErro   As Integer
Dim iExercicio1 As Integer

On Error GoTo Erro_LotePendente_Le_Exercicio

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 13612

    lErro = Comando_Executar(lComando, "SELECT Exercicio FROM LotePendente WHERE Exercicio=?", iExercicio1, iExercicio)
    If lErro <> AD_SQL_SUCESSO Then Error 13613

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 13614

    If lErro = AD_SQL_SEM_DADOS Then Error 13615

    Call Comando_Fechar(lComando)

    LotePendente_Le_Exercicio = SUCESSO

Exit Function

Erro_LotePendente_Le_Exercicio:

    LotePendente_Le_Exercicio = Err

    Select Case Err

        Case 13612
             lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 13613, 13614
             lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LOTEPENDENTE1", Err, iExercicio)

        Case 13615

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153474)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function SaldoInicialConta_Le(ByVal objSaldoInicialConta As ClassSaldoInicialConta) As Long
'Le o saldo inicial da conta passada como parametro e devolve o resultado em objPlanoConta.

Dim lErro As Long
Dim lComando As Long
Dim dSldIni1 As Double

On Error GoTo Erro_SaldoInicialConta_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 10343

    'acessa o saldo inicial da conta em questão
    lErro = Comando_Executar(lComando, "SELECT SldIni FROM SaldoInicialConta WHERE FilialEmpresa=? AND Conta=?", dSldIni1, objSaldoInicialConta.iFilialEmpresa, objSaldoInicialConta.sConta)
    If lErro <> AD_SQL_SUCESSO Then Error 10344

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 10345

    objSaldoInicialConta.dSldIni = dSldIni1

    Call Comando_Fechar(lComando)

    SaldoInicialConta_Le = SUCESSO

    Exit Function

Erro_SaldoInicialConta_Le:

    SaldoInicialConta_Le = Err

    Select Case Err

        Case 10343
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 10344, 10345
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SALDOINICIALCONTA", Err, objSaldoInicialConta.iFilialEmpresa, objSaldoInicialConta.sConta)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153475)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function SaldoInicialContaCcl_Le_Todos_Conta(ByVal objSaldoInicialContaCcl As ClassSaldoInicialContaCcl, ByVal colSaldoInicialContaCcl As Collection) As Long
'Le todas as associacoes da conta em questão com centros de custo/lucro e carrega-os em colContaCcl

Dim lErro As Long
Dim lComando As Long
Dim sCcl As String
Dim dSldIni As Double
Dim sDescCcl As String
Dim objSaldoInicialContaCcl1 As ClassSaldoInicialContaCcl
    
On Error GoTo Erro_SaldoInicialContaCcl_Le_Todos_Conta

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 10347
    
    sCcl = String(STRING_CCL, 0)
    sDescCcl = String(STRING_CCL_DESCRICAO, 0)
    
    lErro = Comando_Executar(lComando, "SELECT SaldoInicialContaCcl.Ccl, SldIni, DescCcl FROM SaldoInicialContaCcl, Ccl WHERE FilialEmpresa=? AND Conta=? AND SaldoInicialContaCcl.Ccl = Ccl.Ccl", sCcl, dSldIni, sDescCcl, objSaldoInicialContaCcl.iFilialEmpresa, objSaldoInicialContaCcl.sConta)
    If lErro <> AD_SQL_SUCESSO Then Error 10348
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 10349
    
    Do While lErro = AD_SQL_SUCESSO
    
        Set objSaldoInicialContaCcl1 = New ClassSaldoInicialContaCcl
        
        objSaldoInicialContaCcl1.iFilialEmpresa = objSaldoInicialContaCcl.iFilialEmpresa
        objSaldoInicialContaCcl1.sConta = objSaldoInicialContaCcl.sConta
        objSaldoInicialContaCcl1.sCcl = sCcl
        objSaldoInicialContaCcl1.dSldIni = dSldIni
        objSaldoInicialContaCcl1.sDescCcl = sDescCcl
    
        colSaldoInicialContaCcl.Add objSaldoInicialContaCcl1

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 10350

    Loop
    
    Call Comando_Fechar(lComando)
    
    SaldoInicialContaCcl_Le_Todos_Conta = SUCESSO
    
    Exit Function

Erro_SaldoInicialContaCcl_Le_Todos_Conta:

    SaldoInicialContaCcl_Le_Todos_Conta = Err
    
    Select Case Err
                
        Case 10347
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 10348, 10349, 10350
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SALDOINICIALCONTACCL3", Err, objSaldoInicialContaCcl.iFilialEmpresa, objSaldoInicialContaCcl.sConta)
                  
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 153476)
        
        End Select
        
        Call Comando_Fechar(lComando)
        
    Exit Function

End Function

Function Exercicio_Le_Codigo(ByVal objExercicio As ClassExercicio) As Long
'Lê o código do Exercício à partir do nome externoo passado como parâmetro. Retorna os dados em objExercício

Dim lErro As Long
Dim lComando As Long
Dim tExercicio As typeExercicio
    
On Error GoTo Erro_Exercicio_Le_Codigo

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 28729
    
    lErro = Comando_Executar(lComando, "SELECT Exercicio, NumPeriodos, Status, DataInicio, DataFim FROM Exercicios WHERE NomeExterno = ?", tExercicio.iExercicio, tExercicio.iNumPeriodos, tExercicio.iStatus, tExercicio.dtDataInicio, tExercicio.dtDataFim, objExercicio.sNomeExterno)
    If lErro <> AD_SQL_SUCESSO Then Error 28730
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 28731
        
    If lErro = AD_SQL_SEM_DADOS Then Error 28732
    
    objExercicio.iExercicio = tExercicio.iExercicio
    objExercicio.iNumPeriodos = tExercicio.iNumPeriodos
    objExercicio.iStatus = tExercicio.iStatus
    objExercicio.dtDataInicio = tExercicio.dtDataInicio
    objExercicio.dtDataFim = tExercicio.dtDataFim
    
    Call Comando_Fechar(lComando)
    
    Exercicio_Le_Codigo = SUCESSO

    Exit Function

Erro_Exercicio_Le_Codigo:

    Exercicio_Le_Codigo = Err
    
    Select Case Err
    
        Case 28729
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 28730, 28731
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EXERCICIOS1", Err, objExercicio.sNomeExterno)
        
        Case 28732
            'Não encontrou o Exercício. Erro tratado na rotina chamadora.
           
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153477)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function Periodo_Le_Codigo(ByVal objPeriodo As ClassPeriodo) As Long
'Lê o código do Período à partir do nome externo e o exercício passados como parâmetro. Retorna os dados em objPeriodo

Dim lErro As Long
Dim lComando As Long
Dim tPeriodo As typePeriodo
    
On Error GoTo Erro_Periodo_Le_Codigo

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 28733
    
    lErro = Comando_Executar(lComando, "SELECT Periodo, DataInicio, DataFim FROM Periodo WHERE Exercicio = ? AND NomeExterno = ?", tPeriodo.iPeriodo, tPeriodo.dtDataInicio, tPeriodo.dtDataFim, objPeriodo.iExercicio, objPeriodo.sNomeExterno)
    If lErro <> AD_SQL_SUCESSO Then Error 28734
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 28735
        
    If lErro = AD_SQL_SEM_DADOS Then Error 28736
    
    objPeriodo.iPeriodo = tPeriodo.iPeriodo
    objPeriodo.dtDataInicio = tPeriodo.dtDataInicio
    objPeriodo.dtDataFim = tPeriodo.dtDataFim
    
    Call Comando_Fechar(lComando)
    
    Periodo_Le_Codigo = SUCESSO

    Exit Function

Erro_Periodo_Le_Codigo:

    Periodo_Le_Codigo = Err
    
    Select Case Err
    
        Case 28733
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 28734, 28735
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PERIODO3", Err, objPeriodo.iExercicio, objPeriodo.sNomeExterno)
        
        Case 28736
            'Não encontrou o Período. Erro tratado na rotina chamadora.
           
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153478)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function ExerciciosFilial_Le(ByVal objExerciciosFilial As ClassExerciciosFilial) As Long
'retorna os dados do exerciciofilial em questão

Dim lComando As Long
Dim lErro As Long
Dim tExerciciosFilial As typeExerciciosFilial

On Error GoTo Erro_ExerciciosFilial_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 20386
    
    'seleciona os dados do exerciciofilial em questão
    lErro = Comando_Executar(lComando, "SELECT Status, LoteApuracao, DataApuracao, DocApuracao FROM ExerciciosFilial WHERE FilialEmpresa=? AND Exercicio=?", tExerciciosFilial.iStatus, tExerciciosFilial.iLoteApuracao, tExerciciosFilial.dtDataApuracao, tExerciciosFilial.lDocApuracao, objExerciciosFilial.iFilialEmpresa, objExerciciosFilial.iExercicio)
    If lErro <> AD_SQL_SUCESSO Then Error 20387
    
    'le o exerciciofilial em questão
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 20388
    
    'se o exerciciofilial não estiver cadastrado
    If lErro = AD_SQL_SEM_DADOS Then Error 20389
    
    objExerciciosFilial.iStatus = tExerciciosFilial.iStatus
    objExerciciosFilial.iLoteApuracao = tExerciciosFilial.iLoteApuracao
    objExerciciosFilial.dtDataApuracao = tExerciciosFilial.dtDataApuracao
    objExerciciosFilial.lDocApuracao = tExerciciosFilial.lDocApuracao
    
    Call Comando_Fechar(lComando)
    
    ExerciciosFilial_Le = SUCESSO

    Exit Function

Erro_ExerciciosFilial_Le:

    ExerciciosFilial_Le = Err
    
    Select Case Err
    
        Case 20386
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 20387, 20388
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EXERCICIOSFILIAL", Err, objExerciciosFilial.iFilialEmpresa, objExerciciosFilial.iExercicio)
            
        Case 20389
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153479)
        
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function PadraoContabItem_Le(ByVal objPadraoContab As ClassPadraoContab, ByVal colPadraoContabItem As ClassColPadraoContabItem) As Long
'retorna os lançamentos normais de objPadraoContab . Os lançamentos são retornados em colPadraoContabItem
'NÃO INCLUI OS LANÇAMENTOS DE CUSTO.

Dim lComando As Long
Dim lErro As Long
Dim tPadraoContabItem As typePadraoContabItem
Dim objPadraoContabItem As ClassPadraoContabItem

On Error GoTo Erro_PadraoContabItem_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 20463
    
    tPadraoContabItem.sConta = String(STRING_PADRAOCONTABITEM_CONTA, 0)
    tPadraoContabItem.sCcl = String(STRING_PADRAOCONTABITEM_CCL, 0)
    tPadraoContabItem.sCredito = String(STRING_PADRAOCONTABITEM_CREDITO, 0)
    tPadraoContabItem.sDebito = String(STRING_PADRAOCONTABITEM_DEBITO, 0)
    tPadraoContabItem.sHistorico = String(STRING_PADRAOCONTABITEM_HISTORICO, 0)
    tPadraoContabItem.sProduto = String(STRING_PADRAOCONTABITEM_PRODUTO, 0)
    
    'seleciona os lancamentos padrao do Modulo/Transacao/Modelo passado como parametro
    lErro = Comando_Executar(lComando, "SELECT Conta, Ccl, Credito, Debito, Historico, Item, Custo, Produto, Aglutina, Gerencial FROM PadraoContabItem WHERE Modulo=? AND Transacao=? AND SubTipo=? AND Modelo=? AND Custo = ? ORDER BY Item", tPadraoContabItem.sConta, tPadraoContabItem.sCcl, tPadraoContabItem.sCredito, tPadraoContabItem.sDebito, tPadraoContabItem.sHistorico, tPadraoContabItem.iItem, tPadraoContabItem.iCusto, tPadraoContabItem.sProduto, tPadraoContabItem.iAglutina, tPadraoContabItem.iGerencial, objPadraoContab.sModulo, objPadraoContab.iTransacao, objPadraoContab.iSubTipo, objPadraoContab.sModelo, PADRAOCONTABITEM_NORMAL)
    If lErro <> AD_SQL_SUCESSO Then Error 20464
    
    'le o primeiro lancamento
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 20465
    
    Do While lErro = AD_SQL_SUCESSO
    
        Set objPadraoContabItem = New ClassPadraoContabItem
        
        objPadraoContabItem.sModulo = objPadraoContab.sModulo
        objPadraoContabItem.iTransacao = objPadraoContab.iTransacao
        objPadraoContabItem.iSubTipo = objPadraoContab.iSubTipo
        objPadraoContabItem.sModelo = objPadraoContab.sModelo
        objPadraoContabItem.iItem = tPadraoContabItem.iItem
        objPadraoContabItem.sConta = tPadraoContabItem.sConta
        objPadraoContabItem.sCcl = tPadraoContabItem.sCcl
        objPadraoContabItem.sCredito = tPadraoContabItem.sCredito
        objPadraoContabItem.sDebito = tPadraoContabItem.sDebito
        objPadraoContabItem.sHistorico = tPadraoContabItem.sHistorico
        objPadraoContabItem.iCusto = tPadraoContabItem.iCusto
        objPadraoContabItem.sProduto = tPadraoContabItem.sProduto
        objPadraoContabItem.iAglutina = tPadraoContabItem.iAglutina
        objPadraoContabItem.iGerencial = tPadraoContabItem.iGerencial
        
        colPadraoContabItem.Add objPadraoContabItem
        
        'le o proximo lancamento
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 20466
        
    Loop
    
    Call Comando_Fechar(lComando)
    
    PadraoContabItem_Le = SUCESSO

    Exit Function

Erro_PadraoContabItem_Le:

    PadraoContabItem_Le = Err
    
    Select Case Err
    
        Case 20463
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 20464, 20465, 20466
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PADRAOCONTABITEM", Err, objPadraoContab.sModulo, objPadraoContab.iTransacao, objPadraoContab.sModelo)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153480)
        
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function PadraoContabItem_Le_Custo(ByVal objPadraoContab As ClassPadraoContab, ByVal colPadraoContabItem As ClassColPadraoContabItem) As Long
'retorna os lançamentos de custo de objPadraoContab . Os lançamentos são retornados em colPadraoContabItem
'NÃO INCLUI OS LANÇAMENTOS NORMAIS

Dim lComando As Long
Dim lErro As Long
Dim tPadraoContabItem As typePadraoContabItem
Dim objPadraoContabItem As ClassPadraoContabItem

On Error GoTo Erro_PadraoContabItem_Le_Custo

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 36404
    
    tPadraoContabItem.sConta = String(STRING_PADRAOCONTABITEM_CONTA, 0)
    tPadraoContabItem.sCcl = String(STRING_PADRAOCONTABITEM_CCL, 0)
    tPadraoContabItem.sCredito = String(STRING_PADRAOCONTABITEM_CREDITO, 0)
    tPadraoContabItem.sDebito = String(STRING_PADRAOCONTABITEM_DEBITO, 0)
    tPadraoContabItem.sHistorico = String(STRING_PADRAOCONTABITEM_HISTORICO, 0)
    tPadraoContabItem.sProduto = String(STRING_PADRAOCONTABITEM_PRODUTO, 0)
    
    'seleciona os lancamentos padrao de custo do Modulo/Transacao/Modelo passado como parametro
    lErro = Comando_Executar(lComando, "SELECT Conta, Ccl, Credito, Debito, Historico, Item, Custo, Produto, Aglutina, Gerencial FROM PadraoContabItem WHERE Modulo=? AND Transacao=? AND SubTipo = ? AND Modelo=? AND Custo = ? ORDER BY Item", tPadraoContabItem.sConta, tPadraoContabItem.sCcl, tPadraoContabItem.sCredito, tPadraoContabItem.sDebito, tPadraoContabItem.sHistorico, tPadraoContabItem.iItem, tPadraoContabItem.iCusto, tPadraoContabItem.sProduto, tPadraoContabItem.iAglutina, tPadraoContabItem.iGerencial, objPadraoContab.sModulo, objPadraoContab.iTransacao, objPadraoContab.iSubTipo, objPadraoContab.sModelo, PADRAOCONTABITEM_CUSTO)
    If lErro <> AD_SQL_SUCESSO Then Error 36405
    
    'le o primeiro lancamento
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 36406
    
    Do While lErro = AD_SQL_SUCESSO
    
        Set objPadraoContabItem = New ClassPadraoContabItem
        
        objPadraoContabItem.sModulo = objPadraoContab.sModulo
        objPadraoContabItem.iTransacao = objPadraoContab.iTransacao
        objPadraoContabItem.iSubTipo = objPadraoContab.iSubTipo
        objPadraoContabItem.sModelo = objPadraoContab.sModelo
        objPadraoContabItem.iItem = tPadraoContabItem.iItem
        objPadraoContabItem.sConta = tPadraoContabItem.sConta
        objPadraoContabItem.sCcl = tPadraoContabItem.sCcl
        objPadraoContabItem.sCredito = tPadraoContabItem.sCredito
        objPadraoContabItem.sDebito = tPadraoContabItem.sDebito
        objPadraoContabItem.sHistorico = tPadraoContabItem.sHistorico
        objPadraoContabItem.iCusto = tPadraoContabItem.iCusto
        objPadraoContabItem.sProduto = tPadraoContabItem.sProduto
        objPadraoContabItem.iAglutina = tPadraoContabItem.iAglutina
        objPadraoContabItem.iGerencial = tPadraoContabItem.iGerencial
        
        colPadraoContabItem.Add objPadraoContabItem
        
        'le o proximo lancamento
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 36407
        
    Loop
    
    Call Comando_Fechar(lComando)
    
    PadraoContabItem_Le_Custo = SUCESSO

    Exit Function

Erro_PadraoContabItem_Le_Custo:

    PadraoContabItem_Le_Custo = Err
    
    Select Case Err
    
        Case 36404
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 36405, 36406, 36407
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PADRAOCONTABITEM", Err, objPadraoContab.sModulo, objPadraoContab.iTransacao, objPadraoContab.sModelo)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153481)
        
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function MnemonicoCTB_Le(ByVal sModulo As String, ByVal iTransacao As Integer, ByVal colMnemonico As Collection) As Long
'retorna os mnemonicos associados ao modulo/transacao passados como parametro. Os mnemonicos são retornados em colMnemonico

Dim lComando As Long
Dim sMnemonico As String
Dim iTipo As Integer
Dim tMnemonicoCTB As typeMnemonicoCTB
Dim sNomeGrid As String
Dim sMnemonicoCombo As String
Dim sMnemonicoDesc As String
Dim lErro As Long
Dim objMnemonico As ClassMnemonico

On Error GoTo Erro_MnemonicoCTB_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 36021
    
    sMnemonico = String(STRING_MNEMONICOCTB_MNEMONICO, 0)
    sNomeGrid = String(STRING_NOME_CONTROLE, 0)
    tMnemonicoCTB.sMnemonicoCombo = String(STRING_MNEMONICOCTB_MNEMONICOCOMBO, 0)
    tMnemonicoCTB.sMnemonicoDesc = String(STRING_MNEMONICOCTB_MNEMONICODESC, 0)
    
    'seleciona os mnemonicos do Modulo/Transacao passado como parametro
    lErro = Comando_Executar(lComando, "SELECT Mnemonico, Tipo, NumParam, Param1, Param2, Param3, NomeGrid, MnemonicoCombo, MnemonicoDesc FROM MnemonicoCTB WHERE Modulo=? AND Transacao=?", sMnemonico, iTipo, tMnemonicoCTB.iNumParam, tMnemonicoCTB.iParam1, tMnemonicoCTB.iParam2, tMnemonicoCTB.iParam3, sNomeGrid, tMnemonicoCTB.sMnemonicoCombo, tMnemonicoCTB.sMnemonicoDesc, sModulo, iTransacao)
    If lErro <> AD_SQL_SUCESSO Then Error 36022
    
    'le o primeiro mnemonico
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 36023
    
    Do While lErro = AD_SQL_SUCESSO
    
        Set objMnemonico = New ClassMnemonico
    
        objMnemonico.sModulo = sModulo
        objMnemonico.iTransacao = iTransacao
        objMnemonico.sMnemonico = sMnemonico
        objMnemonico.iTipo = iTipo
        objMnemonico.iNumParam = tMnemonicoCTB.iNumParam
        objMnemonico.iParam1 = tMnemonicoCTB.iParam1
        objMnemonico.iParam2 = tMnemonicoCTB.iParam2
        objMnemonico.iParam3 = tMnemonicoCTB.iParam3
        objMnemonico.sGrid = sNomeGrid
        objMnemonico.sMnemonicoCombo = tMnemonicoCTB.sMnemonicoCombo
        objMnemonico.sMnemonicoDesc = tMnemonicoCTB.sMnemonicoDesc
        
        colMnemonico.Add objMnemonico
        
        'le o proximo mnemonico
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 36024
        
    Loop
    
    Call Comando_Fechar(lComando)
    
    MnemonicoCTB_Le = SUCESSO

    Exit Function

Erro_MnemonicoCTB_Le:

    MnemonicoCTB_Le = Err
    
    Select Case Err
    
        Case 36021
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 36022, 36023, 36024
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MNEMONICOCTB", Err, sModulo, iTransacao)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153482)
        
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function PadraoContab_Le(ByVal objPadraoContab As ClassPadraoContab) As Long
'retorna o modelo de geracao padrao deste Modulo/Transacao

Dim lComando As Long
Dim lErro As Long
Dim sModelo As String

On Error GoTo Erro_PadraoContab_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 20458
    
    sModelo = String(STRING_PADRAOCONTAB_MODELO, 0)
    
    'seleciona o modelo de geracao padrao deste Modulo/Transacao
    lErro = Comando_Executar(lComando, "SELECT Modelo FROM PadraoContab WHERE Modulo=? AND Transacao=? AND SubTipo=? AND Padrao=?", sModelo, objPadraoContab.sModulo, objPadraoContab.iTransacao, objPadraoContab.iSubTipo, PADRAOCONTAB_MODELO_PADRAO)
    If lErro <> AD_SQL_SUCESSO Then Error 20459
    
    'le o modelo em questão
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 20460
    
    'se o modelo não estiver cadastrado
    If lErro = AD_SQL_SEM_DADOS Then Error 20461
    
    objPadraoContab.sModelo = sModelo
    
    Call Comando_Fechar(lComando)
    
    PadraoContab_Le = SUCESSO

    Exit Function

Erro_PadraoContab_Le:

    PadraoContab_Le = Err
    
    Select Case Err
    
        Case 20458
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 20459, 20460
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PADRAOCONTAB", Err, objPadraoContab.sModulo, objPadraoContab.iTransacao)
            
        Case 20461
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PADRAOCONTAB_SEM_MODELO_PADRAO", Err, objPadraoContab.sModulo, objPadraoContab.iTransacao)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153483)
        
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function MnemonicoCTB_Le_Mnemonico(ByVal objMnemonico As ClassMnemonico) As Long
'retorna os dados do mnemonico passado como parametro

Dim lComando As Long
Dim sMnemonico As String
Dim iTipo As Integer
Dim tMnemonicoCTB As typeMnemonicoCTB
Dim sNomeGrid As String
Dim sMnemonicoCombo As String
Dim sMnemonicoDesc As String
Dim lErro As Long

On Error GoTo Erro_MnemonicoCTB_Le_Mnemonico

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 36079
    
    sNomeGrid = String(STRING_NOME_CONTROLE, 0)
    tMnemonicoCTB.sMnemonico = String(STRING_MNEMONICOCTB_MNEMONICO, 0)
    tMnemonicoCTB.sMnemonicoDesc = String(STRING_MNEMONICOCTB_MNEMONICODESC, 0)
    
    'seleciona o mnemonico passado como parametro
    lErro = Comando_Executar(lComando, "SELECT Mnemonico, Tipo, NumParam, Param1, Param2, Param3, NomeGrid, MnemonicoDesc FROM MnemonicoCTB WHERE Modulo=? AND Transacao=? AND MnemonicoCombo=?", tMnemonicoCTB.sMnemonico, tMnemonicoCTB.iTipo, tMnemonicoCTB.iNumParam, tMnemonicoCTB.iParam1, tMnemonicoCTB.iParam2, tMnemonicoCTB.iParam3, sNomeGrid, tMnemonicoCTB.sMnemonicoDesc, objMnemonico.sModulo, objMnemonico.iTransacao, objMnemonico.sMnemonicoCombo)
    If lErro <> AD_SQL_SUCESSO Then Error 36080
    
    'le o mnemonico
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 36081
    
    If lErro = AD_SQL_SEM_DADOS Then Error 36082
    
    objMnemonico.iTipo = tMnemonicoCTB.iTipo
    objMnemonico.iNumParam = tMnemonicoCTB.iNumParam
    objMnemonico.iParam1 = tMnemonicoCTB.iParam1
    objMnemonico.iParam2 = tMnemonicoCTB.iParam2
    objMnemonico.iParam3 = tMnemonicoCTB.iParam3
    objMnemonico.sMnemonico = tMnemonicoCTB.sMnemonico
    objMnemonico.sMnemonicoDesc = tMnemonicoCTB.sMnemonicoDesc
    objMnemonico.sGrid = tMnemonicoCTB.sNomeGrid
    
    Call Comando_Fechar(lComando)
    
    MnemonicoCTB_Le_Mnemonico = SUCESSO

    Exit Function

Erro_MnemonicoCTB_Le_Mnemonico:

    MnemonicoCTB_Le_Mnemonico = Err
    
    Select Case Err
    
        Case 36079
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 36080, 36081
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MNEMONICOCTB1", Err, objMnemonico.sModulo, objMnemonico.iTransacao, objMnemonico.sMnemonicoCombo)
            
        Case 36082

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153484)
        
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function FormulaFuncao_Le_Todos(ByVal colFormulaFuncao As Collection) As Long
'Le todas as funções e carrega-as em colFormulaFuncao

Dim lErro As Long
Dim lComando As Long
Dim objFormulaFuncao As ClassFormulaFuncao
Dim tFormulaFuncao As typeFormulaFuncao
    
On Error GoTo Erro_FormulaFuncao_Le_Todos

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 36081
    
    tFormulaFuncao.sFuncaoCombo = String(STRING_FORMULAFUNCAO_FUNCAOCOMBO, 0)
    tFormulaFuncao.sFuncaoDesc = String(STRING_FORMULAFUNCAO_FUNCAODESC, 0)
    
    lErro = Comando_Executar(lComando, "SELECT FuncaoCombo, FuncaoDesc FROM FormulaFuncao", tFormulaFuncao.sFuncaoCombo, tFormulaFuncao.sFuncaoDesc)
    If lErro <> AD_SQL_SUCESSO Then Error 36082
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 36083
    
    Do While lErro = AD_SQL_SUCESSO
    
        Set objFormulaFuncao = New ClassFormulaFuncao
        
        objFormulaFuncao.sFuncaoCombo = tFormulaFuncao.sFuncaoCombo
        objFormulaFuncao.sFuncaoDesc = tFormulaFuncao.sFuncaoDesc
    
        colFormulaFuncao.Add objFormulaFuncao

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 36084

    Loop
    
    Call Comando_Fechar(lComando)
    
    FormulaFuncao_Le_Todos = SUCESSO
    
    Exit Function

Erro_FormulaFuncao_Le_Todos:

    FormulaFuncao_Le_Todos = Err
    
    Select Case Err
                
        Case 36081
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 36082, 36083, 36084
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FORMULAFUNCAO", Err)
                  
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 153485)
        
        End Select
        
        Call Comando_Fechar(lComando)
        
    Exit Function

End Function

Function FormulaFuncao_Le(ByVal objFormulaFuncao As ClassFormulaFuncao) As Long
'Le os dados da função passada como parametro

Dim lErro As Long
Dim lComando As Long
Dim tFormulaFuncao As typeFormulaFuncao
    
On Error GoTo Erro_FormulaFuncao_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 36085
    
    tFormulaFuncao.sFuncaoDesc = String(STRING_FORMULAFUNCAO_FUNCAODESC, 0)
    
    lErro = Comando_Executar(lComando, "SELECT FuncaoDesc FROM FormulaFuncao WHERE FuncaoCombo = ?", tFormulaFuncao.sFuncaoDesc, objFormulaFuncao.sFuncaoCombo)
    If lErro <> AD_SQL_SUCESSO Then Error 36086
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 36087
    
    If lErro = AD_SQL_SEM_DADOS Then Error 36088
    
    objFormulaFuncao.sFuncaoDesc = tFormulaFuncao.sFuncaoDesc
    
    Call Comando_Fechar(lComando)
    
    FormulaFuncao_Le = SUCESSO
    
    Exit Function

Erro_FormulaFuncao_Le:

    FormulaFuncao_Le = Err
    
    Select Case Err
                
        Case 36085
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 36086, 36087
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FORMULAFUNCAO1", Err, objFormulaFuncao.sFuncaoCombo)
            
        Case 36088
                  
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 153486)
        
        End Select
        
        Call Comando_Fechar(lComando)
        
    Exit Function

End Function

Function FormulaOperador_Le_Todos(ByVal colFormulaOperador As Collection) As Long
'Le todos os operadores e carrega-as em colFormulaOperador

Dim lErro As Long
Dim lComando As Long
Dim objFormulaOperador As ClassFormulaOperador
Dim tFormulaOperador As typeFormulaOperador
    
On Error GoTo Erro_FormulaOperador_Le_Todos

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 36091
    
    tFormulaOperador.sOperadorCombo = String(STRING_FORMULAOPERADOR_OPERADORCOMBO, 0)
    tFormulaOperador.sOperadorDesc = String(STRING_FORMULAOPERADOR_OPERADORDESC, 0)
    
    lErro = Comando_Executar(lComando, "SELECT OperadorCombo, OperadorDesc FROM FormulaOperador", tFormulaOperador.sOperadorCombo, tFormulaOperador.sOperadorDesc)
    If lErro <> AD_SQL_SUCESSO Then Error 36092
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 36093
    
    Do While lErro = AD_SQL_SUCESSO
    
        Set objFormulaOperador = New ClassFormulaOperador
        
        objFormulaOperador.sOperadorCombo = tFormulaOperador.sOperadorCombo
        objFormulaOperador.sOperadorDesc = tFormulaOperador.sOperadorDesc
    
        colFormulaOperador.Add objFormulaOperador

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 36094

    Loop
    
    Call Comando_Fechar(lComando)
    
    FormulaOperador_Le_Todos = SUCESSO
    
    Exit Function

Erro_FormulaOperador_Le_Todos:

    FormulaOperador_Le_Todos = Err
    
    Select Case Err
                
        Case 36091
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 36092, 36093, 36094
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FORMULAOPERADOR", Err)
                  
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 153487)
        
        End Select
        
        Call Comando_Fechar(lComando)
        
    Exit Function

End Function

Function FormulaOperador_Le(ByVal objFormulaOperador As ClassFormulaOperador) As Long
'Le os dados do operador passado como parametro

Dim lErro As Long
Dim lComando As Long
Dim tFormulaOperador As typeFormulaOperador
    
On Error GoTo Erro_FormulaOperador_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 36095
    
    tFormulaOperador.sOperadorDesc = String(STRING_FORMULAOPERADOR_OPERADORDESC, 0)
    
    lErro = Comando_Executar(lComando, "SELECT OperadorDesc FROM FormulaOperador WHERE OperadorCombo = ?", tFormulaOperador.sOperadorDesc, objFormulaOperador.sOperadorCombo)
    If lErro <> AD_SQL_SUCESSO Then Error 36096
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 36097
    
    If lErro = AD_SQL_SEM_DADOS Then Error 36098
    
    objFormulaOperador.sOperadorDesc = tFormulaOperador.sOperadorDesc
    
    Call Comando_Fechar(lComando)
    
    FormulaOperador_Le = SUCESSO
    
    Exit Function

Erro_FormulaOperador_Le:

    FormulaOperador_Le = Err
    
    Select Case Err
                
        Case 36095
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 36096, 36097
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FORMULAOPERADOR1", Err, objFormulaOperador.sOperadorCombo)
            
        Case 36098
                  
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 153488)
        
        End Select
        
        Call Comando_Fechar(lComando)
        
    Exit Function

End Function

Function TransacaoCTB_Le_Todos(ByVal sSiglaModulo As String, ByVal colTransacao As Collection) As Long
'Le todas as transacoes relativas ao modulo em questão e carrega-as em colTransacao

Dim lErro As Long
Dim lComando As Long
Dim objTransacao As ClassTransacaoCTB
Dim sTransacao As String
Dim sTransacaoTela As String
Dim iCodigo As Integer
    
On Error GoTo Erro_TransacaoCTB_Le_Todos

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 36103
    
    sTransacao = String(STRING_TRANSACAOCTB_TRANSACAO, 0)
    sTransacaoTela = String(STRING_TRANSACAOCTB_TRANSACAOTELA, 0)
    
    lErro = Comando_Executar(lComando, "SELECT Transacao, Codigo, TransacaoTela FROM TransacaoCTB WHERE Sigla = ?", sTransacao, iCodigo, sTransacaoTela, sSiglaModulo)
    If lErro <> AD_SQL_SUCESSO Then Error 36104
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 36105
    
    Do While lErro = AD_SQL_SUCESSO
    
        Set objTransacao = New ClassTransacaoCTB
        
        objTransacao.sSigla = sSiglaModulo
        objTransacao.sTransacao = sTransacao
        objTransacao.iCodigo = iCodigo
        objTransacao.sTransacaoTela = sTransacaoTela
    
        colTransacao.Add objTransacao

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 36106

    Loop
    
    Call Comando_Fechar(lComando)
    
    TransacaoCTB_Le_Todos = SUCESSO
    
    Exit Function

Erro_TransacaoCTB_Le_Todos:

    TransacaoCTB_Le_Todos = Err
    
    Select Case Err
                
        Case 36103
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 36104, 36105, 36106
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRANSACAOCTB", Err, sSiglaModulo)
                  
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 153489)
        
        End Select
        
        Call Comando_Fechar(lComando)
        
    Exit Function

End Function

Function PadraoContab_Le_Todos(ByVal sSiglaModulo As String, ByVal iTransacao As Integer, ByVal colPadraoContab As Collection, ByVal iSubTipo As Integer) As Long
'leitura dos modelos de contabilizacao referentes ao modulo/transacao em questao

Dim lComando As Long
Dim lErro As Long
Dim sModelo As String
Dim objPadraoContab As ClassPadraoContab
Dim iPadrao As Integer

On Error GoTo Erro_PadraoContab_Le_Todos

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 36118
    
    sModelo = String(STRING_PADRAOCONTAB_MODELO, 0)
    
    'seleciona os modelos deste Modulo/Transacao
    lErro = Comando_Executar(lComando, "SELECT Modelo, Padrao FROM PadraoContab WHERE Modulo=? AND Transacao=? AND SubTipo=?", sModelo, iPadrao, sSiglaModulo, iTransacao, iSubTipo)
    If lErro <> AD_SQL_SUCESSO Then Error 36119
    
    'le o primeiro modelo
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 36120
    
    Do While lErro = AD_SQL_SUCESSO
        
        Set objPadraoContab = New ClassPadraoContab
        
        objPadraoContab.sModulo = sSiglaModulo
        objPadraoContab.iTransacao = iTransacao
        objPadraoContab.sModelo = sModelo
        objPadraoContab.iPadrao = iPadrao
        objPadraoContab.iSubTipo = iSubTipo
        
        colPadraoContab.Add objPadraoContab
        
        'le o proximo modelo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 36121
        
    Loop
    
    Call Comando_Fechar(lComando)
    
    PadraoContab_Le_Todos = SUCESSO

    Exit Function

Erro_PadraoContab_Le_Todos:

    PadraoContab_Le_Todos = Err
    
    Select Case Err
    
        Case 36118
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 36119, 36120, 36121
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PADRAOCONTAB", Err, sSiglaModulo, iTransacao)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153490)
        
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function


End Function

Function PadraoContab_Le1(ByVal objPadraoContab As ClassPadraoContab) As Long
'retorna os dados do  modelo passado como parametro. Retorna seus dados em objPadraoContab.

Dim lComando As Long
Dim lErro As Long
Dim iPadrao As Integer

On Error GoTo Erro_PadraoContab_Le1

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 36157
    
    'seleciona o modelo de geracao padrao deste Modulo/Transacao
    lErro = Comando_Executar(lComando, "SELECT Padrao FROM PadraoContab WHERE Modulo=? AND Transacao=? AND SubTipo=? AND Modelo=?", iPadrao, objPadraoContab.sModulo, objPadraoContab.iTransacao, objPadraoContab.iSubTipo, objPadraoContab.sModelo)
    If lErro <> AD_SQL_SUCESSO Then Error 36158
    
    'le o modelo em questão
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 36159
    
    'se o modelo não estiver cadastrado
    If lErro = AD_SQL_SEM_DADOS Then Error 36160
    
    objPadraoContab.iPadrao = iPadrao
    
    Call Comando_Fechar(lComando)
    
    PadraoContab_Le1 = SUCESSO

    Exit Function

Erro_PadraoContab_Le1:

    PadraoContab_Le1 = Err
    
    Select Case Err
    
        Case 36157
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 36158, 36159
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PADRAOCONTAB1", Err, objPadraoContab.sModulo, objPadraoContab.iTransacao, objPadraoContab.sModelo)
            
        Case 36160
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153491)
        
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function LotePendente_Le1(ByVal iFilialEmpresa As Integer) As Long
'Lê todos os lotes de uma FilialEmpresa

Dim lErro As Long
Dim lComando As Long
Dim objLote As ClassLote
Dim tLote As typeLote

On Error GoTo Erro_LotePendente_Le1
   
    tLote.sOrigem = String(STRING_ORIGEM, 0)
    tLote.sIdOriginal = String(STRING_IDORIGINAL, 0)
     
    'Inicializa comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 33606
    
    lErro = Comando_Executar(lComando, "SELECT Origem , Exercicio, Periodo, Lote, IdAtualizacao FROM LotePendente WHERE FilialEmpresa = ? ORDER BY Origem, Exercicio, Periodo, Lote", tLote.sOrigem, tLote.iExercicio, tLote.iPeriodo, tLote.iLote, tLote.iIdAtualizacao, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then Error 33607
        
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 33608
    
    'Não encontrou nenhum lote
    If lErro = AD_SQL_SEM_DADOS Then Error 33611
        
    'Libera comando
    Call Comando_Fechar(lComando)
    
    LotePendente_Le1 = SUCESSO

Exit Function

Erro_LotePendente_Le1:

    LotePendente_Le1 = Err

    Select Case Err
    
        Case 33606
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 33607, 33608
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LOTEPENDENTE1", Err)
            
        Case 33611
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153492)

    End Select
    
    Call Comando_Fechar(lComando)
        
    Exit Function

End Function

Function LanPendente_Le_Doc2(ByVal iFilialEmpresa As Integer, ByVal iTransacao As Integer, ByVal lNumIntDoc As Long, colLancamentos As Collection, ByVal sSiglaModulo As String) As Long
'le os lançamentos pendentes relativos ao documento passado como parametro e coloca-os em colLancamentos

Dim lComando As Long
Dim lErro As Long
Dim objLanc As ClassLancamento_Detalhe
Dim tLanc As typeLancamento

On Error GoTo Erro_LanPendente_Le_Doc2

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 36287
    
    tLanc.sOrigem = String(STRING_ORIGEM, 0)
    tLanc.sConta = String(STRING_CONTA, 0)
    tLanc.sCcl = String(STRING_CCL, 0)
    tLanc.sHistorico = String(STRING_HISTORICO, 0)
    tLanc.sProduto = String(STRING_PRODUTO, 0)
    tLanc.sModelo = String(STRING_PADRAOCONTAB_MODELO, 0)
    tLanc.sDocOrigem = String(STRING_LANPENDENTE_DOCORIGEM, 0)
    tLanc.sProduto = String(STRING_PRODUTO, 0)
    
    lErro = Comando_Executar(lComando, "SELECT  Origem, Exercicio, PeriodoLan, Doc, Seq, Lote, PeriodoLote, Data, Conta, Ccl, Historico, Valor, NumIntDoc, FilialCliForn, CliForn, Transacao, Aglutina, ContaSimples, SeqContraPartida, Status, Modelo, ValorLivroAuxiliar, ClienteFornecedor, FilialEmpresa, Produto, ApropriaCrProd, EscaninhoCusto, DocOrigem, Quantidade, DataEstoque " & _
    "FROM LanPendente WHERE FilialEmpresa = ? AND Transacao = ? AND NumIntDoc = ? AND Origem = ? ", tLanc.sOrigem, tLanc.iExercicio, tLanc.iPeriodoLan, tLanc.lDoc, tLanc.iSeq, tLanc.iLote, tLanc.iPeriodoLote, tLanc.dtData, tLanc.sConta, tLanc.sCcl, tLanc.sHistorico, tLanc.dValor, tLanc.lNumIntDoc, tLanc.iFilialCliForn, tLanc.lCliForn, tLanc.iTransacao, tLanc.iAglutina, tLanc.lContaSimples, tLanc.iSeqContraPartida, tLanc.iStatus, tLanc.sModelo, tLanc.dValorLivroAuxiliar, tLanc.iClienteFornecedor, tLanc.iFilialEmpresa, tLanc.sProduto, tLanc.iApropriaCRProd, tLanc.iEscaninho_Custo, tLanc.sDocOrigem, tLanc.dQuantidade, tLanc.dtDataEstoque, iFilialEmpresa, iTransacao, lNumIntDoc, sSiglaModulo)
    If lErro <> AD_SQL_SUCESSO Then Error 36288
        
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 36289
        
    If lErro = AD_SQL_SEM_DADOS Then Error 36290
        
    Set colLancamentos = New Collection

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objLanc = New ClassLancamento_Detalhe
        
        objLanc.iFilialEmpresa = iFilialEmpresa
        objLanc.sOrigem = tLanc.sOrigem
        objLanc.iExercicio = tLanc.iExercicio
        objLanc.iPeriodoLan = tLanc.iPeriodoLan
        objLanc.lDoc = tLanc.lDoc
        objLanc.iSeq = tLanc.iSeq
        objLanc.iLote = tLanc.iLote
        objLanc.iPeriodoLote = tLanc.iPeriodoLote
        objLanc.dtData = tLanc.dtData
        objLanc.sConta = tLanc.sConta
        objLanc.sCcl = tLanc.sCcl
        objLanc.sHistorico = tLanc.sHistorico
        objLanc.dValor = tLanc.dValor
        objLanc.lNumIntDoc = tLanc.lNumIntDoc
        objLanc.iFilialCliForn = tLanc.iFilialCliForn
        objLanc.lCliForn = tLanc.lCliForn
        objLanc.iTransacao = tLanc.iTransacao
        objLanc.iAglutina = tLanc.iAglutina
        objLanc.lContaSimples = tLanc.lContaSimples
        objLanc.iSeqContraPartida = tLanc.iSeqContraPartida
        objLanc.iStatus = tLanc.iStatus
        objLanc.sProduto = tLanc.sProduto
        objLanc.sModelo = tLanc.sModelo
        objLanc.dValorLivroAuxiliar = tLanc.dValorLivroAuxiliar
        objLanc.iClienteFornecedor = tLanc.iClienteFornecedor
        objLanc.sProduto = tLanc.sProduto
        objLanc.iApropriaCRProd = tLanc.iApropriaCRProd
        objLanc.iEscaninho_Custo = tLanc.iEscaninho_Custo
        objLanc.sDocOrigem = tLanc.sDocOrigem
        objLanc.dQuantidade = tLanc.dQuantidade
        objLanc.dtDataEstoque = tLanc.dtDataEstoque
        
        colLancamentos.Add objLanc
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 36291
        
    Loop
    
    Call Comando_Fechar(lComando)
    
    LanPendente_Le_Doc2 = SUCESSO
    
    Exit Function
 
Erro_LanPendente_Le_Doc2:

    LanPendente_Le_Doc2 = Err
    
    Select Case Err
    
        Case 36287
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 36288, 36289, 36291
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LANPENDENTE3", Err)
           
        Case 36290
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153493)
    
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function LanPendente_Le_Doc3(ByVal iFilialEmpresa As Integer, ByVal iTransacao As Integer, ByVal lNumIntDoc As Long, colLancamentos As Collection, ByVal sSiglaModulo As String, sModelo As String, dtData As Date) As Long
'le os lançamentos pendentes relativos ao documento passado como parametro e coloca-os em colLancamentos

Dim lComando As Long
Dim lErro As Long
Dim objLanc As ClassLancamento_Detalhe
Dim tLanc As typeLancamento

On Error GoTo Erro_LanPendente_Le_Doc3

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 92016
    
    tLanc.sOrigem = String(STRING_ORIGEM, 0)
    tLanc.sConta = String(STRING_CONTA, 0)
    tLanc.sCcl = String(STRING_CCL, 0)
    tLanc.sHistorico = String(STRING_HISTORICO, 0)
    tLanc.sProduto = String(STRING_PRODUTO, 0)
    tLanc.sModelo = String(STRING_PADRAOCONTAB_MODELO, 0)
    tLanc.sDocOrigem = String(STRING_LANPENDENTE_DOCORIGEM, 0)
    
    lErro = Comando_Executar(lComando, "SELECT  Origem, Exercicio, PeriodoLan, Doc, Seq, Lote, PeriodoLote, Data, Conta, Ccl, Historico, Valor, NumIntDoc, FilialCliForn, CliForn, Transacao, Aglutina, ContaSimples, SeqContraPartida, Status, Modelo, ValorLivroAuxiliar, ClienteFornecedor, Produto, ApropriaCrProd, EscaninhoCusto, DocOrigem, Quantidade, DataEstoque, Gerencial FROM LanPendente WHERE FilialEmpresa = ? AND Transacao = ? AND NumIntDoc = ?", tLanc.sOrigem, tLanc.iExercicio, tLanc.iPeriodoLan, tLanc.lDoc, tLanc.iSeq, tLanc.iLote, tLanc.iPeriodoLote, tLanc.dtData, tLanc.sConta, tLanc.sCcl, tLanc.sHistorico, tLanc.dValor, tLanc.lNumIntDoc, tLanc.iFilialCliForn, tLanc.lCliForn, tLanc.iTransacao, tLanc.iAglutina, tLanc.lContaSimples, tLanc.iSeqContraPartida, tLanc.iStatus, tLanc.sModelo, tLanc.dValorLivroAuxiliar, tLanc.iClienteFornecedor, tLanc.sProduto, tLanc.iApropriaCRProd, tLanc.iEscaninho_Custo, tLanc.sDocOrigem, tLanc.dQuantidade, tLanc.dtDataEstoque, _
    tLanc.iGerencial, iFilialEmpresa, iTransacao, lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 92017
        
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92018
        
    If lErro = AD_SQL_SEM_DADOS Then gError 92019
        
    'devolve o modelo que foi usado para criar os lançamentos, se houver
    sModelo = tLanc.sModelo
    dtData = tLanc.dtData
        
    Set colLancamentos = New Collection

    Do While lErro <> AD_SQL_SEM_DADOS
    
        If tLanc.sOrigem = sSiglaModulo Then
        
            Set objLanc = New ClassLancamento_Detalhe
            
            objLanc.iFilialEmpresa = iFilialEmpresa
            objLanc.sOrigem = tLanc.sOrigem
            objLanc.iExercicio = tLanc.iExercicio
            objLanc.iPeriodoLan = tLanc.iPeriodoLan
            objLanc.lDoc = tLanc.lDoc
            objLanc.iSeq = tLanc.iSeq
            objLanc.iLote = tLanc.iLote
            objLanc.iPeriodoLote = tLanc.iPeriodoLote
            objLanc.dtData = tLanc.dtData
            objLanc.sConta = tLanc.sConta
            objLanc.sCcl = tLanc.sCcl
            objLanc.sHistorico = tLanc.sHistorico
            objLanc.dValor = tLanc.dValor
            objLanc.lNumIntDoc = tLanc.lNumIntDoc
            objLanc.iFilialCliForn = tLanc.iFilialCliForn
            objLanc.lCliForn = tLanc.lCliForn
            objLanc.iTransacao = tLanc.iTransacao
            objLanc.iAglutina = tLanc.iAglutina
            objLanc.lContaSimples = tLanc.lContaSimples
            objLanc.iSeqContraPartida = tLanc.iSeqContraPartida
            objLanc.iStatus = tLanc.iStatus
            objLanc.sProduto = tLanc.sProduto
            objLanc.sModelo = tLanc.sModelo
            objLanc.dValorLivroAuxiliar = tLanc.dValorLivroAuxiliar
            objLanc.iClienteFornecedor = tLanc.iClienteFornecedor
            objLanc.iApropriaCRProd = tLanc.iApropriaCRProd
            objLanc.iEscaninho_Custo = tLanc.iEscaninho_Custo
            objLanc.sDocOrigem = tLanc.sDocOrigem
            objLanc.dQuantidade = tLanc.dQuantidade
            objLanc.dtDataEstoque = tLanc.dtDataEstoque
            objLanc.iGerencial = tLanc.iGerencial
            
            colLancamentos.Add objLanc
            
        End If
            
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92020
        
    Loop
    
    If colLancamentos.Count = 0 Then gError 92019
    
    Call Comando_Fechar(lComando)
    
    LanPendente_Le_Doc3 = SUCESSO
    
    Exit Function
 
Erro_LanPendente_Le_Doc3:

    LanPendente_Le_Doc3 = gErr
    
    Select Case gErr
    
        Case 92016
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 92017, 92018, 92020
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LANPENDENTE3", gErr)
           
        Case 92019
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153494)
    
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function


Function LanPendente_Le_GrupoTransExc(ByVal iGrupoTransExc As Integer, ByVal lNumIntDocOrigem As Long, ByVal colLancamentos As Collection) As Long
'IMPORTANTE: está funcao é especifica da operacao de exclusao/estorno de contabilizacao
'Consulta a tabela de Lançamentos Pendentes com a chave NumIntDoc p/transacao do grupo identificado na tabela TransacaoCTB p/iGrupoTransExc.
'retorna lancamentos ordenados por filialempresa

Dim lErro As Long, iIndice As Integer
Dim lComando As Long
Dim objLanc As ClassLancamento_Detalhe
Dim tLanc As typeLancamento, iFilialEmpresa As Integer

On Error GoTo Erro_LanPendente_Le_GrupoTransExc

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 41990
    
    tLanc.sOrigem = String(STRING_ORIGEM, 0)
    tLanc.sConta = String(STRING_CONTA, 0)
    tLanc.sCcl = String(STRING_CCL, 0)
    tLanc.sHistorico = String(STRING_HISTORICO, 0)
    tLanc.sProduto = String(STRING_PRODUTO, 0)
    tLanc.sModelo = String(STRING_PADRAOCONTAB_MODELO, 0)
    tLanc.sDocOrigem = String(STRING_LANPENDENTE_DOCORIGEM, 0)
    
    lErro = Comando_Executar(lComando, "SELECT DISTINCT FilialEmpresa, Origem, Exercicio, PeriodoLan, Doc, Seq, Lote, PeriodoLote, Data, Conta, Ccl, Historico, Valor, NumIntDoc, FilialCliForn , CliForn, LanPendente.Transacao, ContaSimples, SeqContraPartida, Produto, Quantidade, DataEstoque, Status, Modelo, ValorLivroAuxiliar, ClienteFornecedor, ApropriaCRProd, LanPendente.Aglutina, LanPendente.EscaninhoCusto, LanPendente.DocOrigem FROM LanPendente, TransacaoCTB WHERE LanPendente.Transacao = TransacaoCTB.Codigo AND NumIntDoc = ? AND TransacaoCTB.GrupoTransExc = ? ORDER BY FilialEmpresa" _
        , iFilialEmpresa, tLanc.sOrigem, tLanc.iExercicio, tLanc.iPeriodoLan, tLanc.lDoc, tLanc.iSeq, tLanc.iLote, tLanc.iPeriodoLote, tLanc.dtData, tLanc.sConta, tLanc.sCcl, tLanc.sHistorico, tLanc.dValor, tLanc.lNumIntDoc, tLanc.iFilialCliForn, tLanc.lCliForn, tLanc.iTransacao, tLanc.lContaSimples, tLanc.iSeqContraPartida, tLanc.sProduto, tLanc.dQuantidade, tLanc.dtDataEstoque, tLanc.iStatus, tLanc.sModelo, tLanc.dValorLivroAuxiliar, tLanc.iClienteFornecedor, tLanc.iApropriaCRProd, tLanc.iAglutina, tLanc.iEscaninho_Custo, tLanc.sDocOrigem, lNumIntDocOrigem, iGrupoTransExc)
    If lErro <> AD_SQL_SUCESSO Then Error 41991
        
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 41992
        
    If lErro = AD_SQL_SEM_DADOS Then Error 41993
        
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objLanc = New ClassLancamento_Detalhe
        
        objLanc.iFilialEmpresa = iFilialEmpresa
        objLanc.sOrigem = tLanc.sOrigem
        objLanc.iExercicio = tLanc.iExercicio
        objLanc.iPeriodoLan = tLanc.iPeriodoLan
        objLanc.lDoc = tLanc.lDoc
        objLanc.iSeq = tLanc.iSeq
        objLanc.iLote = tLanc.iLote
        objLanc.iPeriodoLote = tLanc.iPeriodoLote
        objLanc.dtData = tLanc.dtData
        objLanc.sConta = tLanc.sConta
        objLanc.sCcl = tLanc.sCcl
        objLanc.sHistorico = tLanc.sHistorico
        objLanc.dValor = tLanc.dValor
        objLanc.lNumIntDoc = tLanc.lNumIntDoc
        objLanc.iFilialCliForn = tLanc.iFilialCliForn
        objLanc.lCliForn = tLanc.lCliForn
        objLanc.iTransacao = tLanc.iTransacao
        objLanc.iAglutina = LANCAMENTO_NAO_AGLUTINADO
        objLanc.lContaSimples = tLanc.lContaSimples
        objLanc.iSeqContraPartida = tLanc.iSeqContraPartida
        objLanc.iAglutina = tLanc.iAglutina
        objLanc.iEscaninho_Custo = tLanc.iEscaninho_Custo
        objLanc.sDocOrigem = tLanc.sDocOrigem
        
        'reprocessamento
        objLanc.sProduto = tLanc.sProduto
        objLanc.dQuantidade = tLanc.dQuantidade
        objLanc.dtDataEstoque = tLanc.dtDataEstoque
        
        objLanc.sModelo = tLanc.sModelo
        objLanc.dValorLivroAuxiliar = tLanc.dValorLivroAuxiliar
        objLanc.iClienteFornecedor = tLanc.iClienteFornecedor
        objLanc.iStatus = tLanc.iStatus
        objLanc.iApropriaCRProd = tLanc.iApropriaCRProd
        
        colLancamentos.Add objLanc
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 41994
        
    Loop
    
    Call Comando_Fechar(lComando)
    
    LanPendente_Le_GrupoTransExc = SUCESSO
     
    Exit Function
    
Erro_LanPendente_Le_GrupoTransExc:

    LanPendente_Le_GrupoTransExc = Err
     
    Select Case Err
          
        Case 41990
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 41991, 41992, 41994
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LANPENDENTE3", Err)
        
        Case 41993 'nao achou lcto
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153495)
     
    End Select
     
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function Lancamento_Le_GrupoTransExc(ByVal iGrupoTransExc As Integer, ByVal lNumIntDocOrigem As Long, ByVal colLancamentos As Collection) As Long
'IMPORTANTE: está funcao é especifica da operacao de exclusao/estorno de contabilizacao
'Consulta a tabela de Lançamentos com a chave NumIntDoc p/transacao que na tabela TransacaoCTB tenha GrupoTransExc igual a iGrupoTransExc.
'retorna lancamentos ordenados por filialempresa

Dim lErro As Long, iIndice As Integer
Dim lComando As Long
Dim objLanc As ClassLancamento_Detalhe
Dim tLanc As typeLancamento, iFilialEmpresa As Integer

On Error GoTo Erro_Lancamento_Le_GrupoTransExc

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 41985
    
    tLanc.sOrigem = String(STRING_ORIGEM, 0)
    tLanc.sConta = String(STRING_CONTA, 0)
    tLanc.sCcl = String(STRING_CCL, 0)
    tLanc.sHistorico = String(STRING_HISTORICO, 0)
    tLanc.sProduto = String(STRING_PRODUTO, 0)
    tLanc.sModelo = String(STRING_PADRAOCONTAB_MODELO, 0)
    tLanc.sDocOrigem = String(STRING_LANPENDENTE_DOCORIGEM, 0)
    
    'reprocessamento
    lErro = Comando_Executar(lComando, "SELECT DISTINCT FilialEmpresa, Origem, Exercicio, PeriodoLan, Doc, Seq, Lote, PeriodoLote, Data, Conta, Ccl, Historico, Valor, NumIntDoc, FilialCliForn , CliForn, Lancamentos.Transacao, DocAglutinado, ContaSimples, SeqContraPartida, Produto, Quantidade, DataEstoque, Status, Modelo, ValorLivroAuxiliar, ClienteFornecedor,  ApropriaCRProd, Lancamentos.EscaninhoCusto, Lancamentos.DocOrigem FROM Lancamentos, TransacaoCTB WHERE Lancamentos.Transacao = TransacaoCTB.Codigo AND NumIntDoc = ? AND TransacaoCTB.GrupoTransExc = ? ORDER BY FilialEmpresa", _
        iFilialEmpresa, tLanc.sOrigem, tLanc.iExercicio, tLanc.iPeriodoLan, tLanc.lDoc, tLanc.iSeq, tLanc.iLote, tLanc.iPeriodoLote, tLanc.dtData, tLanc.sConta, tLanc.sCcl, tLanc.sHistorico, tLanc.dValor, tLanc.lNumIntDoc, tLanc.iFilialCliForn, tLanc.lCliForn, tLanc.iTransacao, tLanc.lDocAglutinado, tLanc.lContaSimples, tLanc.iSeqContraPartida, tLanc.sProduto, tLanc.dQuantidade, tLanc.dtDataEstoque, tLanc.iStatus, tLanc.sModelo, tLanc.dValorLivroAuxiliar, tLanc.iClienteFornecedor, tLanc.iApropriaCRProd, tLanc.iEscaninho_Custo, tLanc.sDocOrigem, lNumIntDocOrigem, iGrupoTransExc)
    If lErro <> AD_SQL_SUCESSO Then Error 41986
        
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 41987
        
    If lErro = AD_SQL_SEM_DADOS Then Error 41988
        
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objLanc = New ClassLancamento_Detalhe
        
        objLanc.iFilialEmpresa = iFilialEmpresa
        objLanc.sOrigem = tLanc.sOrigem
        objLanc.iExercicio = tLanc.iExercicio
        objLanc.iPeriodoLan = tLanc.iPeriodoLan
        objLanc.lDoc = tLanc.lDoc
        objLanc.iSeq = tLanc.iSeq
        objLanc.iLote = tLanc.iLote
        objLanc.iPeriodoLote = tLanc.iPeriodoLote
        objLanc.dtData = tLanc.dtData
        objLanc.sConta = tLanc.sConta
        objLanc.sCcl = tLanc.sCcl
        objLanc.sHistorico = tLanc.sHistorico
        objLanc.dValor = tLanc.dValor
        objLanc.lNumIntDoc = tLanc.lNumIntDoc
        objLanc.iFilialCliForn = tLanc.iFilialCliForn
        objLanc.lCliForn = tLanc.lCliForn
        objLanc.iTransacao = tLanc.iTransacao
        If tLanc.lDocAglutinado <> 0 Then objLanc.iAglutina = LANCAMENTO_AGLUTINADO
        objLanc.lContaSimples = tLanc.lContaSimples
        objLanc.iSeqContraPartida = tLanc.iSeqContraPartida
        objLanc.iEscaninho_Custo = tLanc.iEscaninho_Custo
        objLanc.sDocOrigem = tLanc.sDocOrigem
                
        'reprocessamento
        objLanc.sProduto = tLanc.sProduto
        objLanc.dQuantidade = tLanc.dQuantidade
        objLanc.dtDataEstoque = tLanc.dtDataEstoque
        objLanc.iStatus = tLanc.iStatus
        
        objLanc.sModelo = tLanc.sModelo
        objLanc.dValorLivroAuxiliar = tLanc.dValorLivroAuxiliar
        objLanc.iClienteFornecedor = tLanc.iClienteFornecedor
        objLanc.iApropriaCRProd = tLanc.iApropriaCRProd
        
        colLancamentos.Add objLanc
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 41989
        
    Loop
    
    Call Comando_Fechar(lComando)
    
    Lancamento_Le_GrupoTransExc = SUCESSO
     
    Exit Function
    
Erro_Lancamento_Le_GrupoTransExc:

    Lancamento_Le_GrupoTransExc = Err
     
    Select Case Err
          
        Case 41985
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 41986, 41987, 41989
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LANCAMENTOS3", Err)
        
        Case 41988 'nao achou lcto
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153496)
     
    End Select
     
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function Lancamento_Le_Doc2(ByVal iFilialEmpresa As Integer, ByVal iTransacao As Integer, ByVal lNumIntDoc As Long, ByVal colLancamentos As Collection) As Long
'le os lançamentos relativos ao documento passado como parametro e coloca-os em colLancamentos

Dim lComando As Long
Dim lErro As Long
Dim objLanc As ClassLancamento_Detalhe
Dim tLanc As typeLancamento

On Error GoTo Erro_Lancamento_Le_Doc2

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 36292
    
    tLanc.sOrigem = String(STRING_ORIGEM, 0)
    tLanc.sConta = String(STRING_CONTA, 0)
    tLanc.sCcl = String(STRING_CCL, 0)
    tLanc.sHistorico = String(STRING_HISTORICO, 0)
    tLanc.sProduto = String(STRING_PRODUTO, 0)
    tLanc.sModelo = String(STRING_PADRAOCONTAB_MODELO, 0)
    tLanc.sDocOrigem = String(STRING_LANPENDENTE_DOCORIGEM, 0)
     
    lErro = Comando_Executar(lComando, "SELECT  Origem, Exercicio, PeriodoLan, Doc, Seq, Lote, PeriodoLote, Data, Conta, Ccl, Historico, Valor, NumIntDoc, FilialCliForn , CliForn, Transacao, Aglutinado, DocAglutinado, ContaSimples, SeqContraPartida, Produto, Quantidade, DataEstoque, Modelo, ApropriaCrProd, EscaninhoCusto, ValorLivroAuxiliar, ClienteFornecedor, DocOrigem, Status  FROM Lancamentos WHERE FilialEmpresa = ? AND Transacao = ? AND NumIntDoc = ? AND Status = 0", _
    tLanc.sOrigem, tLanc.iExercicio, tLanc.iPeriodoLan, tLanc.lDoc, tLanc.iSeq, tLanc.iLote, tLanc.iPeriodoLote, tLanc.dtData, tLanc.sConta, tLanc.sCcl, tLanc.sHistorico, tLanc.dValor, tLanc.lNumIntDoc, tLanc.iFilialCliForn, tLanc.lCliForn, tLanc.iTransacao, tLanc.iAglutina, tLanc.lDocAglutinado, tLanc.lContaSimples, tLanc.iSeqContraPartida, tLanc.sProduto, tLanc.dQuantidade, tLanc.dtDataEstoque, tLanc.sModelo, tLanc.iApropriaCRProd, tLanc.iEscaninho_Custo, tLanc.dValorLivroAuxiliar, tLanc.iClienteFornecedor, tLanc.sDocOrigem, tLanc.iStatus, iFilialEmpresa, iTransacao, lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then Error 36293
        
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 36294
        
    If lErro = AD_SQL_SEM_DADOS Then Error 36295
        
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objLanc = New ClassLancamento_Detalhe
        
        objLanc.iFilialEmpresa = iFilialEmpresa
        objLanc.sOrigem = tLanc.sOrigem
        objLanc.iExercicio = tLanc.iExercicio
        objLanc.iPeriodoLan = tLanc.iPeriodoLan
        objLanc.lDoc = tLanc.lDoc
        objLanc.iSeq = tLanc.iSeq
        objLanc.iLote = tLanc.iLote
        objLanc.iPeriodoLote = tLanc.iPeriodoLote
        objLanc.dtData = tLanc.dtData
        objLanc.sConta = tLanc.sConta
        objLanc.sCcl = tLanc.sCcl
        objLanc.sHistorico = tLanc.sHistorico
        objLanc.dValor = tLanc.dValor
        objLanc.lNumIntDoc = tLanc.lNumIntDoc
        objLanc.iFilialCliForn = tLanc.iFilialCliForn
        objLanc.lCliForn = tLanc.lCliForn
        objLanc.iTransacao = tLanc.iTransacao
        objLanc.iAglutina = tLanc.iAglutina
        objLanc.lContaSimples = tLanc.lContaSimples
        objLanc.iSeqContraPartida = tLanc.iSeqContraPartida
        objLanc.sProduto = tLanc.sProduto
        objLanc.dQuantidade = tLanc.dQuantidade
        objLanc.dtDataEstoque = tLanc.dtDataEstoque
        objLanc.sModelo = tLanc.sModelo
        objLanc.iApropriaCRProd = tLanc.iApropriaCRProd
        objLanc.iEscaninho_Custo = tLanc.iEscaninho_Custo
        objLanc.dValorLivroAuxiliar = tLanc.dValorLivroAuxiliar
        objLanc.iClienteFornecedor = tLanc.iClienteFornecedor
        objLanc.sDocOrigem = tLanc.sDocOrigem
        objLanc.iStatus = tLanc.iStatus
        
        colLancamentos.Add objLanc
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 36296
        
    Loop
    
    Call Comando_Fechar(lComando)
    
    Lancamento_Le_Doc2 = SUCESSO
    
    Exit Function
 
Erro_Lancamento_Le_Doc2:

    Lancamento_Le_Doc2 = Err
    
    Select Case Err
    
        Case 36292
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 36293, 36294, 36296
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LANCAMENTOS3", Err)
           
        Case 36295
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153497)
    
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function Lancamento_Le_Doc4(ByVal iFilialEmpresa As Integer, ByVal iTransacao As Integer, ByVal lNumIntDoc As Long, ByVal colLancamentos As Collection, ByVal sSiglaModulo As String, sModelo As String, dtData As Date) As Long
'le os lançamentos relativos ao documento passado como parametro e coloca-os em colLancamentos

Dim lComando As Long
Dim lErro As Long
Dim objLanc As ClassLancamento_Detalhe
Dim tLanc As typeLancamento

On Error GoTo Erro_Lancamento_Le_Doc4

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 92021
    
    tLanc.sOrigem = String(STRING_ORIGEM, 0)
    tLanc.sConta = String(STRING_CONTA, 0)
    tLanc.sCcl = String(STRING_CCL, 0)
    tLanc.sHistorico = String(STRING_HISTORICO, 0)
    tLanc.sProduto = String(STRING_PRODUTO, 0)
    tLanc.sModelo = String(STRING_PADRAOCONTAB_MODELO, 0)
    tLanc.sDocOrigem = String(STRING_LANPENDENTE_DOCORIGEM, 0)
    
    lErro = Comando_Executar(lComando, "SELECT  Origem, Exercicio, PeriodoLan, Doc, Seq, Lote, PeriodoLote, Data, Conta, Ccl, Historico, Valor, NumIntDoc, FilialCliForn , CliForn, Transacao, Aglutinado, DocAglutinado, ContaSimples, SeqContraPartida, Produto, Quantidade, DataEstoque, Modelo, ApropriaCrProd, EscaninhoCusto, ValorLivroAuxiliar, ClienteFornecedor, DocOrigem, Gerencial FROM Lancamentos WHERE FilialEmpresa = ? AND Transacao = ? AND NumIntDoc = ? AND Status = 0", _
    tLanc.sOrigem, tLanc.iExercicio, tLanc.iPeriodoLan, tLanc.lDoc, tLanc.iSeq, tLanc.iLote, tLanc.iPeriodoLote, tLanc.dtData, tLanc.sConta, tLanc.sCcl, tLanc.sHistorico, tLanc.dValor, tLanc.lNumIntDoc, tLanc.iFilialCliForn, tLanc.lCliForn, tLanc.iTransacao, tLanc.iAglutina, tLanc.lDocAglutinado, tLanc.lContaSimples, tLanc.iSeqContraPartida, tLanc.sProduto, tLanc.dQuantidade, tLanc.dtDataEstoque, tLanc.sModelo, tLanc.iApropriaCRProd, tLanc.iEscaninho_Custo, tLanc.dValorLivroAuxiliar, tLanc.iClienteFornecedor, tLanc.sDocOrigem, tLanc.iGerencial, iFilialEmpresa, iTransacao, lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 92022
        
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92023
        
    If lErro = AD_SQL_SEM_DADOS Then gError 92024
        
    sModelo = tLanc.sModelo
    dtData = tLanc.dtData
        
    Do While lErro <> AD_SQL_SEM_DADOS
    
        If tLanc.sOrigem = sSiglaModulo Then
    
            Set objLanc = New ClassLancamento_Detalhe
            
            objLanc.iFilialEmpresa = iFilialEmpresa
            objLanc.sOrigem = tLanc.sOrigem
            objLanc.iExercicio = tLanc.iExercicio
            objLanc.iPeriodoLan = tLanc.iPeriodoLan
            objLanc.lDoc = tLanc.lDoc
            objLanc.iSeq = tLanc.iSeq
            objLanc.iLote = tLanc.iLote
            objLanc.iPeriodoLote = tLanc.iPeriodoLote
            objLanc.dtData = tLanc.dtData
            objLanc.sConta = tLanc.sConta
            objLanc.sCcl = tLanc.sCcl
            objLanc.sHistorico = tLanc.sHistorico
            objLanc.dValor = tLanc.dValor
            objLanc.lNumIntDoc = tLanc.lNumIntDoc
            objLanc.iFilialCliForn = tLanc.iFilialCliForn
            objLanc.lCliForn = tLanc.lCliForn
            objLanc.iTransacao = tLanc.iTransacao
            objLanc.iAglutina = tLanc.iAglutina
            objLanc.lContaSimples = tLanc.lContaSimples
            objLanc.iSeqContraPartida = tLanc.iSeqContraPartida
            objLanc.sProduto = tLanc.sProduto
            objLanc.dQuantidade = tLanc.dQuantidade
            objLanc.dtDataEstoque = tLanc.dtDataEstoque
            objLanc.sModelo = tLanc.sModelo
            objLanc.iApropriaCRProd = tLanc.iApropriaCRProd
            objLanc.dValorLivroAuxiliar = tLanc.dValorLivroAuxiliar
            objLanc.iEscaninho_Custo = tLanc.iEscaninho_Custo
            objLanc.iClienteFornecedor = tLanc.iClienteFornecedor
            objLanc.sDocOrigem = tLanc.sDocOrigem
            objLanc.iGerencial = tLanc.iGerencial
            
            colLancamentos.Add objLanc
        
        End If
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92025
        
    Loop
    
    If colLancamentos.Count = 0 Then gError 92024
    
    Call Comando_Fechar(lComando)
    
    Lancamento_Le_Doc4 = SUCESSO
    
    Exit Function
 
Erro_Lancamento_Le_Doc4:

    Lancamento_Le_Doc4 = gErr
    
    Select Case gErr
    
        Case 92021
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 92022, 92023, 92025
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LANCAMENTOS3", gErr)
           
        Case 92024
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153498)
    
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function TransacaoCTB_Le(ByVal objTransacaoCTB As ClassTransacaoCTB) As Long
'Le os dados da transacao passada como parametro e devolve os dados em objTransacaoCTB

Dim lErro As Long
Dim lComando As Long
Dim iCodigo As Integer
Dim sTransacaoTela As String
Dim iCliFornecedor As Integer
Dim iObrigatoriedade As Integer, iOrigemLcto As Integer
Dim iGerencial As Integer
    
On Error GoTo Erro_TransacaoCTB_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 36365
    
    sTransacaoTela = String(STRING_TRANSACAOCTB_TRANSACAOTELA, 0)
    
    lErro = Comando_Executar(lComando, "SELECT Codigo, TransacaoTela, CliFornecedor, Obrigatoriedade, OrigemLcto, Gerencial FROM TransacaoCTB WHERE Sigla = ? AND Transacao = ? AND SubTipo = ?", iCodigo, sTransacaoTela, iCliFornecedor, iObrigatoriedade, iOrigemLcto, iGerencial, objTransacaoCTB.sSigla, objTransacaoCTB.sTransacao, objTransacaoCTB.iSubTipo)
    If lErro <> AD_SQL_SUCESSO Then Error 36366
        
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 36367
    
    objTransacaoCTB.iCodigo = iCodigo
    objTransacaoCTB.sTransacaoTela = sTransacaoTela
    objTransacaoCTB.iCliFornecedor = iCliFornecedor
    objTransacaoCTB.iObrigatoriedade = iObrigatoriedade
    objTransacaoCTB.iOrigemLcto = iOrigemLcto
    objTransacaoCTB.iGerencial = iGerencial
    
    Call Comando_Fechar(lComando)
    
    TransacaoCTB_Le = SUCESSO
    
    Exit Function

Erro_TransacaoCTB_Le:

    TransacaoCTB_Le = Err
    
    Select Case Err
                
        Case 36365
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 36366, 36367
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRANSACAOCTB1", Err, objTransacaoCTB.sSigla, objTransacaoCTB.sTransacao)
                  
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 153499)
        
        End Select
        
        Call Comando_Fechar(lComando)
        
    Exit Function

End Function

Function TransacaoCTB_Le1(ByVal objTransacaoCTB As ClassTransacaoCTB) As Long
'Le os dados da transacao passada como parametro e devolve os dados em objTransacaoCTB

Dim lErro As Long
Dim lComando As Long
Dim iCodigo As Integer
Dim sTransacaoTela As String
Dim iCliFornecedor As Integer
    
On Error GoTo Erro_TransacaoCTB_Le1

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 83229
    
    sTransacaoTela = String(STRING_TRANSACAOCTB_TRANSACAOTELA, 0)
    
    lErro = Comando_Executar(lComando, "SELECT Codigo, TransacaoTela, CliFornecedor FROM TransacaoCTB WHERE Sigla = ? AND Transacao = ?", iCodigo, sTransacaoTela, iCliFornecedor, objTransacaoCTB.sSigla, objTransacaoCTB.sTransacao)
    If lErro <> AD_SQL_SUCESSO Then gError 83230
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then gError 83231
    
    objTransacaoCTB.iCodigo = iCodigo
    objTransacaoCTB.sTransacaoTela = sTransacaoTela
    objTransacaoCTB.iCliFornecedor = iCliFornecedor
    
    Call Comando_Fechar(lComando)
    
    TransacaoCTB_Le1 = SUCESSO
    
    Exit Function

Erro_TransacaoCTB_Le1:

    TransacaoCTB_Le1 = gErr
    
    Select Case gErr
                
        Case 83229
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 83230, 83231
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRANSACAOCTB1", gErr, objTransacaoCTB.sSigla, objTransacaoCTB.sTransacao)
                  
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153500)
        
        End Select
        
        Call Comando_Fechar(lComando)
        
    Exit Function

End Function

Function LotePendente_Le2(ByVal objLote As ClassLote) As Long
'Verifica se tem um lote pendente que não esteja em processo de atualização com a Filial/Origem/Exercicio/Periodo passados como parametro e se tiver devolve seus dados.

Dim lErro As Long
Dim lComando As Long
Dim tLote As typeLote

On Error GoTo Erro_LotePendente_Le2
   
    tLote.sIdOriginal = String(STRING_IDORIGINAL, 0)
     
    'Inicializa comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 36414
    
    lErro = Comando_Executar(lComando, "SELECT Lote, TotCre, TotDeb, TotInf, Status, IdOriginal, NumLancInf, NumLancAtual, NumDocInf, NumDocAtual FROM LotePendente WHERE FilialEmpresa = ? AND Origem = ? AND Exercicio=? AND Periodo=? AND IdAtualizacao = 0", tLote.iLote, tLote.dTotCre, tLote.dTotDeb, tLote.dTotInf, tLote.iStatus, tLote.sIdOriginal, tLote.iNumLancInf, tLote.iNumLancAtual, tLote.iNumDocInf, tLote.iNumDocAtual, objLote.iFilialEmpresa, objLote.sOrigem, objLote.iExercicio, objLote.iPeriodo)
    If lErro <> AD_SQL_SUCESSO Then Error 36415
        
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 36416
    
    'Não encontrou nenhum lote
    If lErro = AD_SQL_SEM_DADOS Then Error 36417
        
    objLote.iLote = tLote.iLote
    objLote.dTotCre = tLote.dTotCre
    objLote.dTotDeb = tLote.dTotDeb
    objLote.dTotInf = tLote.dTotInf
    objLote.iStatus = tLote.iStatus
    objLote.sIdOriginal = tLote.sIdOriginal
    objLote.iNumLancInf = tLote.iNumLancInf
    objLote.iNumLancAtual = tLote.iNumLancAtual
    objLote.iNumDocInf = tLote.iNumDocInf
    objLote.iNumDocAtual = tLote.iNumDocAtual
        
    'Libera comando
    Call Comando_Fechar(lComando)
    
    LotePendente_Le2 = SUCESSO

Exit Function

Erro_LotePendente_Le2:

    LotePendente_Le2 = Err

    Select Case Err
    
        Case 36414
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 36415, 36416
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LOTEPENDENTE4", Err, objLote.iFilialEmpresa, objLote.sOrigem, objLote.iExercicio, objLote.iPeriodo)
            
        Case 36417
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153501)

    End Select
    
    Call Comando_Fechar(lComando)
        
    Exit Function

End Function

Function Produto_Le_Apropriacao(ByVal lComando As Long, ByVal sProduto As String, iApropriacao As Integer) As Long
'Lê o atributo Apropriação do produto, cujo codigo foi passado como parâmetro

Dim lErro As Long

On Error GoTo Erro_Produto_Le_Apropriacao

    lErro = Comando_Executar(lComando, "SELECT Apropriacao FROM Produtos WHERE Codigo = ? ", iApropriacao, sProduto)
    If lErro <> AD_SQL_SUCESSO Then Error 36528

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 36529

    Produto_Le_Apropriacao = SUCESSO

    Exit Function

Erro_Produto_Le_Apropriacao:

    Produto_Le_Apropriacao = Err

    Select Case Err

        Case 36528, 36529
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS", Err, sProduto)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153502)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function SldMesEst_Le_CustoMedio(ByVal lComando As Long, ByVal iFilialEmpresa As Integer, ByVal iAno As Integer, ByVal sProduto As String, ByVal iMes As Integer, dCustoMedio As Double) As Long
'le os dados de custo medio de SldMesEst relativos ao Ano, FilialEmpresa, Produto e mes passados como parametro
    
Dim lErro As Long
Dim tSldMesEst As typeSldMesEst
Dim dQuantAtual As Double
Dim dValorAtual As Double
Dim iIndice As Integer

On Error GoTo Erro_SldMesEst_Le_CustoMedio
    
    'seleciona o registro de SldMesEst relativo ao Ano, FilialEmpresa e Produto passados como parametro
    lErro = Comando_ExecutarPos(lComando, "SELECT QuantInicialCusto, ValorInicialCusto, SaldoQuantCusto1, SaldoValorCusto1, SaldoQuantCusto2, SaldoValorCusto2, SaldoQuantCusto3, SaldoValorCusto3, SaldoQuantCusto4, SaldoValorCusto4, SaldoQuantCusto5, SaldoValorCusto5, SaldoQuantCusto6, SaldoValorCusto6, SaldoQuantCusto7, SaldoValorCusto7, SaldoQuantCusto8, SaldoValorCusto8, SaldoQuantCusto9, SaldoValorCusto9, SaldoQuantCusto10, SaldoValorCusto10, SaldoQuantCusto11, SaldoValorCusto11, SaldoQuantCusto12, SaldoValorCusto12 FROM SldMesEst WHERE Ano=? AND FilialEmpresa=? AND Produto=?", 0, _
                              tSldMesEst.dQuantInicialCusto, tSldMesEst.dValorInicialCusto, tSldMesEst.adSaldoQuantCusto(1), tSldMesEst.adSaldoValorCusto(1), tSldMesEst.adSaldoQuantCusto(2), tSldMesEst.adSaldoValorCusto(2), tSldMesEst.adSaldoQuantCusto(3), tSldMesEst.adSaldoValorCusto(3), tSldMesEst.adSaldoQuantCusto(4), tSldMesEst.adSaldoValorCusto(4), tSldMesEst.adSaldoQuantCusto(5), tSldMesEst.adSaldoValorCusto(5), tSldMesEst.adSaldoQuantCusto(6), tSldMesEst.adSaldoValorCusto(6), tSldMesEst.adSaldoQuantCusto(7), tSldMesEst.adSaldoValorCusto(7), tSldMesEst.adSaldoQuantCusto(8), tSldMesEst.adSaldoValorCusto(8), tSldMesEst.adSaldoQuantCusto(9), tSldMesEst.adSaldoValorCusto(9), tSldMesEst.adSaldoQuantCusto(10), tSldMesEst.adSaldoValorCusto(10), _
                              tSldMesEst.adSaldoQuantCusto(11), tSldMesEst.adSaldoValorCusto(11), tSldMesEst.adSaldoQuantCusto(12), tSldMesEst.adSaldoValorCusto(12), iAno, iFilialEmpresa, sProduto)
    If lErro <> AD_SQL_SUCESSO Then Error 36519
    
    'le o SldMesEst
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 36520
    
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 36573
    
    dQuantAtual = tSldMesEst.dQuantInicialCusto
    dValorAtual = tSldMesEst.dValorInicialCusto

    For iIndice = 1 To iMes
    
        dQuantAtual = dQuantAtual + tSldMesEst.adSaldoQuantCusto(iIndice)
        dValorAtual = dValorAtual + tSldMesEst.adSaldoValorCusto(iIndice)
        
    Next
    
    If Abs(dQuantAtual) > QTDE_ESTOQUE_DELTA Then
    
        dCustoMedio = dValorAtual / dQuantAtual
        
    Else
    
        dCustoMedio = 0
        
    End If

    SldMesEst_Le_CustoMedio = SUCESSO

    Exit Function

Erro_SldMesEst_Le_CustoMedio:
    
    SldMesEst_Le_CustoMedio = Err
    
    Select Case Err
    
        Case 36519, 36520
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESEST", Err, iAno, iFilialEmpresa, sProduto)
            
        Case 36573
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_SLDMESEST", Err, iAno, iFilialEmpresa, sProduto)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153503)
        
    End Select
    
    Exit Function

End Function

Function SldMesEst_Le_CustoStandard(ByVal lComando As Long, ByVal iFilialEmpresa As Integer, ByVal iAno As Integer, ByVal sProduto As String, ByVal iMes As Integer, dCustoStandard As Double) As Long
'le os dados de custo standard de SldMesEst relativos ao Ano, FilialEmpresa, Produto e mes passados como parametro
    
Dim lErro As Long
    
On Error GoTo Erro_SldMesEst_Le_CustoStandard
    
    'seleciona o custo standard de SldMesEst relativo ao Ano, FilialEmpresa, Produto, Mes passados como parametro
    lErro = Comando_ExecutarPos(lComando, "SELECT CustoStandard" + CStr(iMes) + " FROM SldMesEst WHERE Ano=? AND FilialEmpresa=? AND Produto=?", 0, dCustoStandard, iAno, iFilialEmpresa, sProduto)
    If lErro <> AD_SQL_SUCESSO Then Error 36521
    
    'le o SldMesEst
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 36522
    
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 36574
    
    SldMesEst_Le_CustoStandard = SUCESSO

    Exit Function

Erro_SldMesEst_Le_CustoStandard:
    
    SldMesEst_Le_CustoStandard = Err
    
    Select Case Err
    
        Case 36521, 36522
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESEST", Err, iAno, iFilialEmpresa, sProduto)
            
        Case 36574
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_SLDMESEST", Err, iAno, iFilialEmpresa, sProduto)
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153504)
        
    End Select
    
    Exit Function

End Function

Function PlanoConta_Le_Netos(ByVal sConta As String, ByVal colPlanoConta As Collection) As Long
'le todos os netos de sConta e coloca-as em colPlanoConta

Dim lComando As Long
Dim lErro As Long
Dim objPlanoConta As ClassPlanoConta
Dim tPlanoConta As typePlanoConta
Dim sConta1 As String
Dim sContaEnxuta As String
Dim iNivel As Integer

On Error GoTo Erro_PlanoConta_Le_Netos

    tPlanoConta.sConta = String(STRING_CONTA, 0)
    tPlanoConta.sDescConta = String(STRING_CONTA_DESCRICAO, 0)

    lErro = Mascara_RetornaContaEnxuta(sConta, sContaEnxuta)
    If lErro <> SUCESSO Then Error 36757
    
    lErro = Mascara_Conta_ObterNivel(sConta, iNivel)
    If lErro <> SUCESSO Then Error 36758

    sConta1 = sContaEnxuta + String(STRING_CONTA - Len(sContaEnxuta), "Z")

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 36759
    
    'le todos os netos de uma conta ordenado por conta
    lErro = Comando_Executar(lComando, "SELECT Conta, NivelConta, DescConta, TipoConta, ContaSimples, Natureza, DigitoVerif, DigitoVerifSimples, HistPadrao, Ativo, UsaContaSimples, Categoria, FluxoCaixa FROM PlanoConta WHERE NivelConta = ? AND Conta > ? AND Conta < ? ORDER BY Conta", tPlanoConta.sConta, tPlanoConta.iNivelConta, tPlanoConta.sDescConta, tPlanoConta.iTipoConta, tPlanoConta.lContaSimples, tPlanoConta.iNatureza, tPlanoConta.iDigitoVerif, tPlanoConta.iDigitoVerifSimples, tPlanoConta.iHistPadrao, tPlanoConta.iAtivo, tPlanoConta.iUsaContaSimples, tPlanoConta.iCategoria, tPlanoConta.iFluxoCaixa, iNivel + 2, sConta, sConta1)
    If lErro <> AD_SQL_SUCESSO Then Error 36760
    
    'le a primeira conta
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 36761
    
    Do While lErro <> AD_SQL_SEM_DADOS

        Set objPlanoConta = New ClassPlanoConta
        
        objPlanoConta.sConta = tPlanoConta.sConta
        objPlanoConta.iNivelConta = tPlanoConta.iNivelConta
        objPlanoConta.sDescConta = tPlanoConta.sDescConta
        objPlanoConta.iTipoConta = tPlanoConta.iTipoConta
        objPlanoConta.lContaSimples = tPlanoConta.lContaSimples
        objPlanoConta.iNatureza = tPlanoConta.iNatureza
        objPlanoConta.iDigitoVerif = tPlanoConta.iDigitoVerif
        objPlanoConta.iDigitoVerifSimples = tPlanoConta.iDigitoVerifSimples
        objPlanoConta.iHistPadrao = tPlanoConta.iHistPadrao
        objPlanoConta.iAtivo = tPlanoConta.iAtivo
        objPlanoConta.iUsaContaSimples = tPlanoConta.iUsaContaSimples
        objPlanoConta.iCategoria = tPlanoConta.iCategoria
        objPlanoConta.iFluxoCaixa = tPlanoConta.iFluxoCaixa
        
        'coloca a conta lida na coleção
        colPlanoConta.Add objPlanoConta
        
        'le a proxima conta
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 36762

    Loop

    Call Comando_Fechar(lComando)
    
    PlanoConta_Le_Netos = SUCESSO

    Exit Function

Erro_PlanoConta_Le_Netos:

    PlanoConta_Le_Netos = Err
    
    Select Case Err

        Case 36757
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, sConta)

        Case 36758
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_CONTA_OBTERNIVEL", Err, sConta)

        Case 36759
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 36760, 36761, 36762
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTA", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153505)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function PlanoConta_Le_Netos_Modulo(ByVal sConta As String, ByVal colPlanoConta As Collection, ByVal sSiglaModulo As String) As Long
'le todos os netos de sConta visiveis ao modulo sSiglaModulo e coloca-as em colPlanoConta

Dim lComando As Long
Dim lErro As Long
Dim objPlanoConta As ClassPlanoConta
Dim tPlanoConta As typePlanoConta
Dim sConta1 As String
Dim sContaEnxuta As String
Dim iNivel As Integer
Dim sSQL As String

On Error GoTo Erro_PlanoConta_Le_Netos_Modulo

    tPlanoConta.sConta = String(STRING_CONTA, 0)
    tPlanoConta.sDescConta = String(STRING_CONTA_DESCRICAO, 0)

    lErro = Mascara_RetornaContaEnxuta(sConta, sContaEnxuta)
    If lErro <> SUCESSO Then Error 44010
    
    lErro = Mascara_Conta_ObterNivel(sConta, iNivel)
    If lErro <> SUCESSO Then Error 44011

    sConta1 = sContaEnxuta + String(STRING_CONTA - Len(sContaEnxuta), "Z")

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 44012

    sSQL = "SELECT Conta, NivelConta, DescConta, TipoConta, ContaSimples, Natureza, DigitoVerif, DigitoVerifSimples, HistPadrao, Ativo, UsaContaSimples, Categoria, FluxoCaixa FROM PlanoConta WHERE NivelConta = ? AND Conta > ? AND Conta < ? AND ("
    
    Select Case sSiglaModulo
    
        Case MODULO_CONTASAPAGAR
            sSQL = sSQL & "CP = 1"
            
        Case MODULO_CONTASARECEBER
            sSQL = sSQL & "CR = 1"
            
        Case MODULO_ESTOQUE
            sSQL = sSQL & "EST = 1"
    
        Case MODULO_FATURAMENTO
            sSQL = sSQL & "FAT = 1"
    
        Case MODULO_TESOURARIA
            sSQL = sSQL & "TES = 1"
    
        Case Else
            Error 44013
            
    End Select
    
    sSQL = sSQL & " OR TipoConta = ?) ORDER BY Conta"
    
    'le todos os netos de uma conta ordenado por conta
    lErro = Comando_Executar(lComando, sSQL, tPlanoConta.sConta, tPlanoConta.iNivelConta, tPlanoConta.sDescConta, tPlanoConta.iTipoConta, tPlanoConta.lContaSimples, tPlanoConta.iNatureza, tPlanoConta.iDigitoVerif, tPlanoConta.iDigitoVerifSimples, tPlanoConta.iHistPadrao, tPlanoConta.iAtivo, tPlanoConta.iUsaContaSimples, tPlanoConta.iCategoria, tPlanoConta.iFluxoCaixa, iNivel + 2, sConta, sConta1, CONTA_SINTETICA)
    If lErro <> AD_SQL_SUCESSO Then Error 44014
    
    'le a primeira conta
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 44015
    
    Do While lErro <> AD_SQL_SEM_DADOS

        Set objPlanoConta = New ClassPlanoConta
        
        objPlanoConta.sConta = tPlanoConta.sConta
        objPlanoConta.iNivelConta = tPlanoConta.iNivelConta
        objPlanoConta.sDescConta = tPlanoConta.sDescConta
        objPlanoConta.iTipoConta = tPlanoConta.iTipoConta
        objPlanoConta.lContaSimples = tPlanoConta.lContaSimples
        objPlanoConta.iNatureza = tPlanoConta.iNatureza
        objPlanoConta.iDigitoVerif = tPlanoConta.iDigitoVerif
        objPlanoConta.iDigitoVerifSimples = tPlanoConta.iDigitoVerifSimples
        objPlanoConta.iHistPadrao = tPlanoConta.iHistPadrao
        objPlanoConta.iAtivo = tPlanoConta.iAtivo
        objPlanoConta.iUsaContaSimples = tPlanoConta.iUsaContaSimples
        objPlanoConta.iCategoria = tPlanoConta.iCategoria
        objPlanoConta.iFluxoCaixa = tPlanoConta.iFluxoCaixa
        
        'coloca a conta lida na coleção
        colPlanoConta.Add objPlanoConta
        
        'le a proxima conta
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 44016

    Loop

    Call Comando_Fechar(lComando)
    
    PlanoConta_Le_Netos_Modulo = SUCESSO

    Exit Function

Erro_PlanoConta_Le_Netos_Modulo:

    PlanoConta_Le_Netos_Modulo = Err
    
    Select Case Err

        Case 44010
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, sConta)

        Case 44011
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_CONTA_OBTERNIVEL", Err, sConta)

        Case 44012
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 44013
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MODULO_INEXISTENTE", Err, sSiglaModulo)

        Case 44014, 44015, 44016
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTA", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153506)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function PlanoConta_Le_Niveis0e1(ByVal colPlanoConta As Collection) As Long
'le todas as contas de nível 0 e 1 e coloca-as em colPlanoConta

Dim lComando As Long
Dim lErro As Long
Dim objPlanoConta As ClassPlanoConta
Dim tPlanoConta As typePlanoConta
Dim sConta1 As String
Dim sContaEnxuta As String
Dim iNivel As Integer

On Error GoTo Erro_PlanoConta_Le_Niveis0e1

    tPlanoConta.sConta = String(STRING_CONTA, 0)
    tPlanoConta.sDescConta = String(STRING_CONTA_DESCRICAO, 0)

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 36764
    
    'le todos os filhos de uma conta ordenado por conta
    lErro = Comando_Executar(lComando, "SELECT Conta, NivelConta, DescConta, TipoConta, ContaSimples, Natureza, DigitoVerif, DigitoVerifSimples, HistPadrao, Ativo, UsaContaSimples, Categoria, FluxoCaixa FROM PlanoConta WHERE NivelConta = 1 OR NivelConta = 2 ORDER BY Conta", tPlanoConta.sConta, tPlanoConta.iNivelConta, tPlanoConta.sDescConta, tPlanoConta.iTipoConta, tPlanoConta.lContaSimples, tPlanoConta.iNatureza, tPlanoConta.iDigitoVerif, tPlanoConta.iDigitoVerifSimples, tPlanoConta.iHistPadrao, tPlanoConta.iAtivo, tPlanoConta.iUsaContaSimples, tPlanoConta.iCategoria, tPlanoConta.iFluxoCaixa)
    If lErro <> AD_SQL_SUCESSO Then Error 36765
    
    'le a primeira conta
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 36766
    
    Do While lErro <> AD_SQL_SEM_DADOS

        Set objPlanoConta = New ClassPlanoConta
        
        objPlanoConta.sConta = tPlanoConta.sConta
        objPlanoConta.iNivelConta = tPlanoConta.iNivelConta
        objPlanoConta.sDescConta = tPlanoConta.sDescConta
        objPlanoConta.iTipoConta = tPlanoConta.iTipoConta
        objPlanoConta.lContaSimples = tPlanoConta.lContaSimples
        objPlanoConta.iNatureza = tPlanoConta.iNatureza
        objPlanoConta.iDigitoVerif = tPlanoConta.iDigitoVerif
        objPlanoConta.iDigitoVerifSimples = tPlanoConta.iDigitoVerifSimples
        objPlanoConta.iHistPadrao = tPlanoConta.iHistPadrao
        objPlanoConta.iAtivo = tPlanoConta.iAtivo
        objPlanoConta.iUsaContaSimples = tPlanoConta.iUsaContaSimples
        objPlanoConta.iCategoria = tPlanoConta.iCategoria
        objPlanoConta.iFluxoCaixa = tPlanoConta.iFluxoCaixa
        
        'coloca a conta lida na coleção
        colPlanoConta.Add objPlanoConta
        
        'le a proxima conta
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 36767

    Loop

    Call Comando_Fechar(lComando)
    
    PlanoConta_Le_Niveis0e1 = SUCESSO

    Exit Function

Erro_PlanoConta_Le_Niveis0e1:

    PlanoConta_Le_Niveis0e1 = Err
    
    Select Case Err

        Case 36764
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 36765, 36766, 36767
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTA", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153507)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function PlanoConta_Le_Niveis0e1_Modulo(ByVal colPlanoConta As Collection, ByVal sSiglaModulo As String) As Long
'le todas as contas de nível 0 e 1 visiveis ao Modulo (sSiglaModulo) e coloca-as em colPlanoConta

Dim lComando As Long
Dim lErro As Long
Dim objPlanoConta As ClassPlanoConta
Dim tPlanoConta As typePlanoConta
Dim sConta1 As String
Dim sContaEnxuta As String
Dim iNivel As Integer
Dim sSQL As String

On Error GoTo Erro_PlanoConta_Le_Niveis0e1_Modulo

    tPlanoConta.sConta = String(STRING_CONTA, 0)
    tPlanoConta.sDescConta = String(STRING_CONTA_DESCRICAO, 0)

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 44004

    sSQL = "SELECT Conta, NivelConta, DescConta, TipoConta, ContaSimples, Natureza, DigitoVerif, DigitoVerifSimples, HistPadrao, Ativo, UsaContaSimples, Categoria, FluxoCaixa FROM PlanoConta WHERE (NivelConta = 1 OR NivelConta = 2) AND ("
    
    Select Case sSiglaModulo
    
        Case MODULO_CONTASAPAGAR
            sSQL = sSQL & "CP = 1"
            
        Case MODULO_CONTASARECEBER
            sSQL = sSQL & "CR = 1"
            
        Case MODULO_ESTOQUE
            sSQL = sSQL & "EST = 1"
    
        Case MODULO_FATURAMENTO
            sSQL = sSQL & "FAT = 1"
    
        Case MODULO_TESOURARIA
            sSQL = sSQL & "TES = 1"
    
        Case MODULO_SERVICOS
            sSQL = sSQL & "SRV = 1"
    
        Case Else
            Error 44005
            
    End Select
    
    sSQL = sSQL & " OR TipoConta = ?) ORDER BY Conta"
    
    'le todos as contas de nivel 0 e 1 visiveis ao modulo em questão
    lErro = Comando_Executar(lComando, sSQL, tPlanoConta.sConta, tPlanoConta.iNivelConta, tPlanoConta.sDescConta, tPlanoConta.iTipoConta, tPlanoConta.lContaSimples, tPlanoConta.iNatureza, tPlanoConta.iDigitoVerif, tPlanoConta.iDigitoVerifSimples, tPlanoConta.iHistPadrao, tPlanoConta.iAtivo, tPlanoConta.iUsaContaSimples, tPlanoConta.iCategoria, tPlanoConta.iFluxoCaixa, CONTA_SINTETICA)
    If lErro <> AD_SQL_SUCESSO Then Error 44006
    
    'le a primeira conta
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 44007
    
    Do While lErro <> AD_SQL_SEM_DADOS

        Set objPlanoConta = New ClassPlanoConta
        
        objPlanoConta.sConta = tPlanoConta.sConta
        objPlanoConta.iNivelConta = tPlanoConta.iNivelConta
        objPlanoConta.sDescConta = tPlanoConta.sDescConta
        objPlanoConta.iTipoConta = tPlanoConta.iTipoConta
        objPlanoConta.lContaSimples = tPlanoConta.lContaSimples
        objPlanoConta.iNatureza = tPlanoConta.iNatureza
        objPlanoConta.iDigitoVerif = tPlanoConta.iDigitoVerif
        objPlanoConta.iDigitoVerifSimples = tPlanoConta.iDigitoVerifSimples
        objPlanoConta.iHistPadrao = tPlanoConta.iHistPadrao
        objPlanoConta.iAtivo = tPlanoConta.iAtivo
        objPlanoConta.iUsaContaSimples = tPlanoConta.iUsaContaSimples
        objPlanoConta.iCategoria = tPlanoConta.iCategoria
        objPlanoConta.iFluxoCaixa = tPlanoConta.iFluxoCaixa
        
        'coloca a conta lida na coleção
        colPlanoConta.Add objPlanoConta
        
        'le a proxima conta
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 44008

    Loop

    Call Comando_Fechar(lComando)
    
    PlanoConta_Le_Niveis0e1_Modulo = SUCESSO

    Exit Function

Erro_PlanoConta_Le_Niveis0e1_Modulo:

    PlanoConta_Le_Niveis0e1_Modulo = Err
    
    Select Case Err

        Case 44004
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 44005
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MODULO_INEXISTENTE", Err, sSiglaModulo)

        Case 44006, 44007, 44008
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTA", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153508)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function PadraoContabItem_Le_Modelo(ByVal objPadraoContab As ClassPadraoContab, ByVal colPadraoContabItem As ClassColPadraoContabItem) As Long
'retorna os lançamentos de objPadraoContab referentes ao modelo em questão. Os lançamentos são retornados em colPadraoContabItem

Dim lComando As Long
Dim lErro As Long
Dim tPadraoContabItem As typePadraoContabItem
Dim objPadraoContabItem As ClassPadraoContabItem

On Error GoTo Erro_PadraoContabItem_Le_Modelo

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 36805
    
    tPadraoContabItem.sConta = String(STRING_PADRAOCONTABITEM_CONTA, 0)
    tPadraoContabItem.sCcl = String(STRING_PADRAOCONTABITEM_CCL, 0)
    tPadraoContabItem.sCredito = String(STRING_PADRAOCONTABITEM_CREDITO, 0)
    tPadraoContabItem.sDebito = String(STRING_PADRAOCONTABITEM_DEBITO, 0)
    tPadraoContabItem.sHistorico = String(STRING_PADRAOCONTABITEM_HISTORICO, 0)
    tPadraoContabItem.sProduto = String(STRING_PADRAOCONTABITEM_PRODUTO, 0)
    
    'seleciona os lancamentos padrao do Modulo/Transacao/Modelo passado como parametro
    lErro = Comando_Executar(lComando, "SELECT Conta, Ccl, Credito, Debito, Historico, Item, Custo, Produto, Aglutina, Gerencial FROM PadraoContabItem WHERE Modulo=? AND Transacao=? AND SubTipo=? AND Modelo=? ORDER BY Item", tPadraoContabItem.sConta, tPadraoContabItem.sCcl, tPadraoContabItem.sCredito, tPadraoContabItem.sDebito, tPadraoContabItem.sHistorico, tPadraoContabItem.iItem, tPadraoContabItem.iCusto, tPadraoContabItem.sProduto, tPadraoContabItem.iAglutina, tPadraoContabItem.iGerencial, objPadraoContab.sModulo, objPadraoContab.iTransacao, objPadraoContab.iSubTipo, objPadraoContab.sModelo)
    If lErro <> AD_SQL_SUCESSO Then Error 36806
    
    'le o primeiro lancamento
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 36807
    
    Do While lErro = AD_SQL_SUCESSO
    
        Set objPadraoContabItem = New ClassPadraoContabItem
        
        objPadraoContabItem.sModulo = objPadraoContab.sModulo
        objPadraoContabItem.iTransacao = objPadraoContab.iTransacao
        objPadraoContabItem.iSubTipo = objPadraoContab.iSubTipo
        objPadraoContabItem.sModelo = objPadraoContab.sModelo
        objPadraoContabItem.iItem = tPadraoContabItem.iItem
        objPadraoContabItem.sConta = tPadraoContabItem.sConta
        objPadraoContabItem.sCcl = tPadraoContabItem.sCcl
        objPadraoContabItem.sCredito = tPadraoContabItem.sCredito
        objPadraoContabItem.sDebito = tPadraoContabItem.sDebito
        objPadraoContabItem.sHistorico = tPadraoContabItem.sHistorico
        objPadraoContabItem.iCusto = tPadraoContabItem.iCusto
        objPadraoContabItem.sProduto = tPadraoContabItem.sProduto
        objPadraoContabItem.iAglutina = tPadraoContabItem.iAglutina
        objPadraoContabItem.iGerencial = tPadraoContabItem.iGerencial
        
        colPadraoContabItem.Add objPadraoContabItem
        
        'le o proximo lancamento
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 36808
        
    Loop
    
    Call Comando_Fechar(lComando)
    
    PadraoContabItem_Le_Modelo = SUCESSO

    Exit Function

Erro_PadraoContabItem_Le_Modelo:

    PadraoContabItem_Le_Modelo = Err
    
    Select Case Err
    
        Case 36805
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 36806, 36807, 36808
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PADRAOCONTABITEM", Err, objPadraoContab.sModulo, objPadraoContab.iTransacao, objPadraoContab.sModelo)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153509)
        
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function RateioOff_Le_Doc1(ByVal lComando As Long, ByVal objRateioOff As ClassRateioOff, ByVal colRateioOff As Collection) As Long
'le os lançamentos relativos ao Rateio passado como parametro e coloca-os em colRateioOff

Dim lErro As Long
Dim tRateioOff As typeRateioOff
Dim objRateioOff1 As ClassRateioOff


On Error GoTo Erro_RateioOff_Le_Doc1
    
    tRateioOff.sContaCre = String(STRING_CONTA, 0)
    tRateioOff.sConta = String(STRING_CONTA, 0)
    tRateioOff.sCclOrigem = String(STRING_CCL, 0)
    tRateioOff.sCcl = String(STRING_CCL, 0)
    tRateioOff.sDescricao = String(STRING_RATEIO_DESCRICAO, 0)
    
    lErro = Comando_Executar(lComando, "SELECT  CclOrigem, Descricao, Percentual, Seq, Conta, Ccl, ContaCre, Tipo  FROM RateioOff WHERE Codigo = ?", tRateioOff.sCclOrigem, tRateioOff.sDescricao, tRateioOff.dPercentual, tRateioOff.lSeq, tRateioOff.sConta, tRateioOff.sCcl, tRateioOff.sContaCre, tRateioOff.iTipo, objRateioOff.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 36814

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 36815

    If lErro = AD_SQL_SEM_DADOS Then Error 36816

    objRateioOff.sCclOrigem = tRateioOff.sCclOrigem
    objRateioOff.iTipo = tRateioOff.iTipo
    objRateioOff.sContaCre = tRateioOff.sContaCre
    objRateioOff.sDescricao = tRateioOff.sDescricao

    Do While lErro <> AD_SQL_SEM_DADOS

        Set objRateioOff1 = New ClassRateioOff

        objRateioOff1.lCodigo = objRateioOff.lCodigo
        objRateioOff1.sCclOrigem = tRateioOff.sCclOrigem
        objRateioOff1.sDescricao = tRateioOff.sDescricao
        objRateioOff1.sConta = tRateioOff.sConta
        objRateioOff1.sCcl = tRateioOff.sCcl
        objRateioOff1.sContaCre = tRateioOff.sContaCre
        objRateioOff1.lSeq = tRateioOff.lSeq
        objRateioOff1.dPercentual = tRateioOff.dPercentual
        objRateioOff1.iTipo = tRateioOff.iTipo

        colRateioOff.Add objRateioOff1

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 36817
        
    Loop
    
    RateioOff_Le_Doc1 = SUCESSO
    
    Exit Function
 
Erro_RateioOff_Le_Doc1:

    RateioOff_Le_Doc1 = Err

    Select Case Err
    
        Case 36814, 36815, 36817
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RATEIOOFF", Err)
        
        Case 36816
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153510)
    
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function LotePendente_Le_Lock(ByVal lComando As Long, ByVal objLote As ClassLote) As Long
'le e loca um lote a partir da chave fornecida através de objLote. Devolve os dados em objLote

Dim lErro As Long
Dim tLote As typeLote
    
On Error GoTo Erro_LotePendente_Le_Lock

    tLote.sIdOriginal = String(STRING_IDORIGINAL, 0)

    If giSetupLotePorPeriodo = LOTE_INICIALIZADO_POR_PERIODO Then
    
        lErro = Comando_Executar(lComando, "SELECT Periodo, TotCre, TotDeb, TotInf, Status, IdOriginal, NumLancInf, NumLancAtual, NumDocInf, NumDocAtual FROM LotePendente WHERE FilialEmpresa = ? AND Origem = ? AND Exercicio=? AND Periodo=? AND Lote=?", tLote.iPeriodo, tLote.dTotCre, tLote.dTotDeb, tLote.dTotInf, tLote.iStatus, tLote.sIdOriginal, tLote.iNumLancInf, tLote.iNumLancAtual, tLote.iNumDocInf, tLote.iNumDocAtual, objLote.iFilialEmpresa, objLote.sOrigem, objLote.iExercicio, objLote.iPeriodo, objLote.iLote)
        If lErro <> AD_SQL_SUCESSO Then Error 36832
        
    Else
        
        lErro = Comando_Executar(lComando, "SELECT Periodo, TotCre, TotDeb, TotInf, Status, IdOriginal, NumLancInf, NumLancAtual, NumDocInf, NumDocAtual FROM LotePendente WHERE FilialEmpresa = ? AND Origem = ? AND Exercicio=? AND Lote=?", tLote.iPeriodo, tLote.dTotCre, tLote.dTotDeb, tLote.dTotInf, tLote.iStatus, tLote.sIdOriginal, tLote.iNumLancInf, tLote.iNumLancAtual, tLote.iNumDocInf, tLote.iNumDocAtual, objLote.iFilialEmpresa, objLote.sOrigem, objLote.iExercicio, objLote.iLote)
        If lErro <> AD_SQL_SUCESSO Then Error 36833
        
    End If
    
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 36834
        
    If lErro = AD_SQL_SEM_DADOS Then Error 36835
    
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 36836
    
    objLote.iPeriodo = tLote.iPeriodo
    objLote.dTotCre = tLote.dTotCre
    objLote.dTotDeb = tLote.dTotDeb
    objLote.dTotInf = tLote.dTotInf
    objLote.iStatus = tLote.iStatus
    objLote.sIdOriginal = tLote.sIdOriginal
    objLote.iNumLancInf = tLote.iNumLancInf
    objLote.iNumLancAtual = tLote.iNumLancAtual
    objLote.iNumDocInf = tLote.iNumDocInf
    objLote.iNumDocAtual = tLote.iNumDocAtual
    
    LotePendente_Le_Lock = SUCESSO

Exit Function

Erro_LotePendente_Le_Lock:

    LotePendente_Le_Lock = Err
    
    Select Case Err
    
        Case 36832, 36833, 36834
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LOTEPENDENTE", Err, objLote.iFilialEmpresa, objLote.sOrigem, objLote.iExercicio, objLote.iPeriodo, objLote.iLote)
           
        Case 36835
        
        Case 36836
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_LOTEPENDENTE", Err, objLote.iFilialEmpresa, objLote.sOrigem, objLote.iExercicio, objLote.iPeriodo, objLote.iLote)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153511)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function RateioOff_Le_ContaCcl(ByVal sConta As String, ByVal sCcl As String) As Long
'verifica se existe na tabela RateioOff pelo menos um rateio com a conta e centro de custo em questão.

Dim lComando As Long
Dim lErro As Long
Dim lCodigo As Long
Dim lSeq As Long

On Error GoTo Erro_RateioOff_Le_ContaCcl

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 44046
    
    lErro = Comando_Executar(lComando, "SELECT Codigo, Seq FROM RateioOff WHERE Conta=? And Ccl=?", lCodigo, lSeq, sConta, sCcl)
    If lErro <> AD_SQL_SUCESSO Then Error 44047
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 44048
        
    If lErro = AD_SQL_SEM_DADOS Then Error 44049
        
    Call Comando_Fechar(lComando)
    
    RateioOff_Le_ContaCcl = SUCESSO
    
    Exit Function
 
Erro_RateioOff_Le_ContaCcl:

    RateioOff_Le_ContaCcl = Err
    
    Select Case Err
    
        Case 44046
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 44047, 44048
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RATEIOOFF2", Err, sConta, sCcl)
           
        Case 44049
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153512)
    
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function MnemonicoCTBValor_Le(ByVal objMnemonico As ClassMnemonicoCTBValor) As Long
'retorna os dados do mnemonico passado como parametro

Dim lComando As Long
Dim sMnemonico As String
Dim tMnemonicoCTBValor As typeMnemonicoCTBValor
Dim sDescricao As String
Dim lErro As Long
Dim sValor As String

On Error GoTo Erro_MnemonicoCTBValor_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 39687
    
    tMnemonicoCTBValor.sDescricao = String(STRING_MNEMONICOCTBVALOR_DESCRICAO, 0)
    tMnemonicoCTBValor.sValor = String(STRING_MNEMONICOCTBVALOR_VALOR, 0)
    
    'seleciona o mnemonico passado como parametro
    lErro = Comando_Executar(lComando, "SELECT Descricao, Valor FROM MnemonicoCTBValor WHERE Mnemonico = ?", tMnemonicoCTBValor.sDescricao, tMnemonicoCTBValor.sValor, objMnemonico.sMnemonico)
    If lErro <> AD_SQL_SUCESSO Then Error 39688
    
    'le o mnemonico
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 39689
    
    If lErro = AD_SQL_SEM_DADOS Then Error 39690
    
    objMnemonico.sDescricao = tMnemonicoCTBValor.sDescricao
    objMnemonico.sValor = tMnemonicoCTBValor.sValor
    
    Call Comando_Fechar(lComando)
    
    MnemonicoCTBValor_Le = SUCESSO

    Exit Function

Erro_MnemonicoCTBValor_Le:

    MnemonicoCTBValor_Le = Err
    
    Select Case Err
    
        Case 39687
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 39688, 39689
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MNEMONICOCTBVALOR", Err, objMnemonico.sMnemonico)
            
        Case 39690

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153513)
        
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function RelDRE_Le_Modelos_Distintos(ByVal sRelatorio As String, ByVal colModelos As Collection) As Long
'Le todos os modelos disponíveis no BD

Dim lErro As Long
Dim sComando_SQL As String
Dim sModelo As String
Dim lComando As Long

On Error GoTo Erro_RelDRE_Le_Modelos_Distintos

    sModelo = String(STRING_DRE_MODELO, 0)

    sComando_SQL = "SELECT DISTINCT Modelo FROM RelDRE WHERE Relatorio = ?"

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 44532

    lErro = Comando_Executar(lComando, sComando_SQL, sModelo, sRelatorio)
    If lErro <> AD_SQL_SUCESSO Then Error 44533

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 44534

    Do While lErro <> AD_SQL_SEM_DADOS

        colModelos.Add sModelo

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 44535

    Loop

    Call Comando_Fechar(lComando)

    RelDRE_Le_Modelos_Distintos = SUCESSO

    Exit Function

Erro_RelDRE_Le_Modelos_Distintos:

    RelDRE_Le_Modelos_Distintos = Err

    Select Case Err

        Case 44532
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 44533, 44534, 44535
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELDRE", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153514)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function RelDRE_Le_Modelo(ByVal sRelatorio As String, ByVal sModelo As String, ByVal colRelDRE As Collection) As Long
'Le o modelo em questão e seu componentes

Dim lErro As Long, iIndice As Integer
Dim lComando As Long
Dim sComando_SQL As String
Dim tRelDRE As typeRelDRE
Dim objRelDRE As ClassRelDRE

On Error GoTo Erro_RelDRE_Le_Modelo

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 44536

    tRelDRE.sModelo = String(STRING_DRE_MODELO, 0)
    tRelDRE.sTitulo = String(STRING_DRE_TITULO, 0)

    sComando_SQL = "SELECT Modelo, Codigo, Tipo, Nivel, Titulo, Posicao, Imprime, Exercicio FROM RelDRE WHERE Relatorio = ? AND Modelo = ? ORDER BY Posicao"

    lErro = Comando_Executar(lComando, sComando_SQL, tRelDRE.sModelo, tRelDRE.iCodigo, tRelDRE.iTipo, tRelDRE.iNivel, tRelDRE.sTitulo, tRelDRE.iPosicao, tRelDRE.iImprime, tRelDRE.iExercicio, sRelatorio, sModelo)
    If lErro <> AD_SQL_SUCESSO Then Error 44537

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 44538

    Do While lErro <> AD_SQL_SEM_DADOS

        Set objRelDRE = New ClassRelDRE

        objRelDRE.sModelo = tRelDRE.sModelo
        objRelDRE.iCodigo = tRelDRE.iCodigo
        objRelDRE.iTipo = tRelDRE.iTipo
        objRelDRE.iNivel = tRelDRE.iNivel
        objRelDRE.sTitulo = tRelDRE.sTitulo
        objRelDRE.iPosicao = tRelDRE.iPosicao
        objRelDRE.iImprime = tRelDRE.iImprime
        objRelDRE.iExercicio = tRelDRE.iExercicio

        colRelDRE.Add objRelDRE, "X" & CStr(tRelDRE.iCodigo)

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 44539

    Loop

    Call Comando_Fechar(lComando)

    RelDRE_Le_Modelo = SUCESSO

    Exit Function

Erro_RelDRE_Le_Modelo:

    RelDRE_Le_Modelo = Err

    Select Case Err

        Case 44536
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 44537, 44538, 44539
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELDRE_MODELO", Err, sModelo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153515)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function RelDREConta_Le_Modelo(ByVal sRelatorio As String, ByVal sModelo As String, ByVal colRelDREConta As Collection) As Long
'Le os dados de Contas do modelo atual para a coleção

Dim lErro As Long
Dim sComando_SQL As String
Dim tRelDREConta As typeRelDREConta
Dim objRelDREConta As ClassRelDREConta
Dim lComando As Long

On Error GoTo Erro_RelDREConta_Le_Modelo

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 44540

    tRelDREConta.sModelo = String(STRING_DRE_MODELO, 0)
    tRelDREConta.sContaFinal = String(STRING_CONTA, 0)
    tRelDREConta.sContaInicial = String(STRING_CONTA, 0)
    tRelDREConta.sCclInicial = String(STRING_CCL, 0)
    tRelDREConta.sCclFinal = String(STRING_CCL, 0)

    sComando_SQL = "SELECT Modelo, Codigo, Item, ContaInicial, ContaFinal, CclInicial, CclFinal FROM RelDREConta WHERE Relatorio = ? AND Modelo = ?"

    lErro = Comando_Executar(lComando, sComando_SQL, tRelDREConta.sModelo, tRelDREConta.iCodigo, tRelDREConta.iItem, tRelDREConta.sContaInicial, tRelDREConta.sContaFinal, tRelDREConta.sCclInicial, tRelDREConta.sCclFinal, sRelatorio, sModelo)
    If lErro <> AD_SQL_SUCESSO Then Error 39818

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 39819

    Do While lErro <> AD_SQL_SEM_DADOS

        Set objRelDREConta = New ClassRelDREConta

        objRelDREConta.sModelo = tRelDREConta.sModelo
        objRelDREConta.iCodigo = tRelDREConta.iCodigo
        objRelDREConta.iItem = tRelDREConta.iItem
        objRelDREConta.sContaInicial = tRelDREConta.sContaInicial
        objRelDREConta.sContaFinal = tRelDREConta.sContaFinal
        objRelDREConta.sCclInicial = tRelDREConta.sCclInicial
        objRelDREConta.sCclFinal = tRelDREConta.sCclFinal

        colRelDREConta.Add objRelDREConta

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 44541

    Loop

    Call Comando_Fechar(lComando)

    RelDREConta_Le_Modelo = SUCESSO

    Exit Function

Erro_RelDREConta_Le_Modelo:

    RelDREConta_Le_Modelo = Err

    Select Case Err

        Case 44540
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 39818, 39819, 44541
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELDRECONTA", Err, sModelo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153516)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function RelDREFormula_Le_Modelo(ByVal sRelatorio As String, ByVal sModelo As String, ByVal colRelDREFormula As Collection) As Long
'Le os dados de Fórmulas do modelo atual para a coleção

Dim lErro As Long
Dim sComando_SQL As String
Dim tRelDREFormula As typeRelDREFormula
Dim objRelDREFormula As ClassRelDREFormula
Dim lComando As Long

On Error GoTo Erro_RelDREFormula_Le_Modelo

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 44542

    tRelDREFormula.sModelo = String(STRING_DRE_MODELO, 0)
    tRelDREFormula.sFormula = String(STRING_DRE_TITULO, 0)

    sComando_SQL = "SELECT Modelo, Codigo, Item, CodigoFormula, Operacao, Formula FROM RelDREFormula WHERE Relatorio = ? AND Modelo = ? ORDER BY Codigo, Item"

    lErro = Comando_Executar(lComando, sComando_SQL, tRelDREFormula.sModelo, tRelDREFormula.iCodigo, tRelDREFormula.iItem, tRelDREFormula.iCodigoFormula, tRelDREFormula.iOperacao, tRelDREFormula.sFormula, sRelatorio, sModelo)
    If lErro <> AD_SQL_SUCESSO Then Error 44543

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 44544

    Do While lErro <> AD_SQL_SEM_DADOS

        Set objRelDREFormula = New ClassRelDREFormula

        objRelDREFormula.sModelo = tRelDREFormula.sModelo
        objRelDREFormula.iCodigo = tRelDREFormula.iCodigo
        objRelDREFormula.iItem = tRelDREFormula.iItem
        objRelDREFormula.iCodigoFormula = tRelDREFormula.iCodigoFormula
        objRelDREFormula.iOperacao = tRelDREFormula.iOperacao
        objRelDREFormula.sFormula = tRelDREFormula.sFormula

        colRelDREFormula.Add objRelDREFormula

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 44545

    Loop

    Call Comando_Fechar(lComando)

    RelDREFormula_Le_Modelo = SUCESSO

    Exit Function

Erro_RelDREFormula_Le_Modelo:

    RelDREFormula_Le_Modelo = Err

    Select Case Err

        Case 44542
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 44543, 44544, 44545
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELDREFORMULA", Err, sModelo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153517)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Ccl_ExisteCcl(iExiste As Integer) As Long
'verifica se tem algum centro de custo cadastrado.
'se existir iExiste = POSSUI_CCL , caso contrario iExiste = NAO_POSSUI_CCL

Dim lComando As Long
Dim lErro As Long
Dim iTipoCcl As Integer

On Error GoTo Erro_Ccl_ExisteCcl

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 44737

    'Verifica se o Ccl existe
    lErro = Comando_Executar(lComando, "SELECT TipoCcl FROM Ccl", iTipoCcl)
    If lErro <> AD_SQL_SUCESSO Then Error 44738

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 44739
        
    'Ccl está cadastrado
    If lErro = SUCESSO Then
        iExiste = POSSUI_CCL
    Else
        iExiste = NAO_POSSUI_CCL
    End If

    Call Comando_Fechar(lComando)

    Ccl_ExisteCcl = SUCESSO

    Exit Function

Erro_Ccl_ExisteCcl:

    Ccl_ExisteCcl = Err
    
    Select Case Err

        Case 44737
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 44738, 44739
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CCL1", Err)
         
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153518)

    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function PadraoContab_ObtemModeloPadrao(ByVal sModulo As String, ByVal sTransacao As String, sModelo As String, ByVal iSubTipo As Integer) As Long
'retorna o modelo padrao p/o modulo-transacao

Dim lErro As Long, lComando As Long

On Error GoTo Erro_PadraoContab_ObtemModeloPadrao

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 32134
    
    'seleciona o modelo de geracao padrao deste Modulo/Transacao
    sModelo = String(STRING_PADRAOCONTAB_MODELO, 0)
    lErro = Comando_Executar(lComando, "SELECT PadraoContab.Modelo FROM PadraoContab, TransacaoCTB WHERE TransacaoCTB.Sigla = ? AND TransacaoCTB.Transacao = ? AND TransacaoCTB.SubTipo = ? AND TransacaoCTB.Codigo = PadraoContab.Transacao AND TransacaoCTB.SubTipo = PadraoContab.SubTipo AND PadraoContab.Modulo = ? AND PadraoContab.Padrao = ?", sModelo, sModulo, sTransacao, iSubTipo, sModulo, PADRAOCONTAB_MODELO_PADRAO)
    If lErro <> AD_SQL_SUCESSO Then Error 32135
    
    'le o modelo em questão
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 32136
    
    'se nao houver modelo padrao cadastrado
    If lErro = AD_SQL_SEM_DADOS Then Error 32137
    
    Call Comando_Fechar(lComando)
    
    PadraoContab_ObtemModeloPadrao = SUCESSO
     
    Exit Function
    
Erro_PadraoContab_ObtemModeloPadrao:

    PadraoContab_ObtemModeloPadrao = Err
     
    Select Case Err
          
        Case 32134
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 32135, 32136
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PADRAOCONTAB_MODPADRAO", Err, sTransacao, sModulo)
        
        Case 32137
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_PADRAOCONTAB_SEMMODPADRAO", Err, sTransacao, sModulo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153519)
     
    End Select
     
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function Lote_LotePendente_Le(ByVal objLote As ClassLote) As Long
'verifica se para o lote passado como parametro já tem um lote ou lote pendente cadastrado com o mesmo numero

Dim lComando As Long
Dim lErro As Long
Dim iPeriodo As Integer

On Error GoTo Erro_Lote_LotePendente_Le

    If giSetupLotePorPeriodo = LOTE_INICIALIZADO_POR_EXERCICIO Then
    
        lComando = Comando_Abrir()
        If lComando = 0 Then Error 55145
    
        lErro = Comando_Executar(lComando, "SELECT Periodo FROM LotePendente WHERE FilialEmpresa = ? AND Origem = ? AND Exercicio=? AND Periodo<>? AND Lote=?", iPeriodo, objLote.iFilialEmpresa, objLote.sOrigem, objLote.iExercicio, objLote.iPeriodo, objLote.iLote)
        If lErro <> AD_SQL_SUCESSO Then Error 55146
        
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 55147
        
        If lErro = AD_SQL_SUCESSO Then Error 55148
        
        lErro = Comando_Executar(lComando, "SELECT Periodo FROM Lote WHERE FilialEmpresa = ? AND Origem = ? AND Exercicio=? AND Periodo<>? AND Lote=?", iPeriodo, objLote.iFilialEmpresa, objLote.sOrigem, objLote.iExercicio, objLote.iPeriodo, objLote.iLote)
        If lErro <> AD_SQL_SUCESSO Then Error 55149
        
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 55150
        
        If lErro = AD_SQL_SUCESSO Then Error 55151
        
        Call Comando_Fechar(lComando)
    
    End If
    
    Lote_LotePendente_Le = SUCESSO

    Exit Function

Erro_Lote_LotePendente_Le:

    Lote_LotePendente_Le = Err
    
    Select Case Err
    
        Case 55145
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 55146, 55147
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LOTEPENDENTE", Err, objLote.iFilialEmpresa, objLote.sOrigem, objLote.iExercicio, objLote.iPeriodo, objLote.iLote)
           
        Case 55148, 55151
            Call Rotina_Erro(vbOKOnly, "ERRO_LOTE_JA_CADASTRADO_OUTRO_PERIODO", Err, objLote.iFilialEmpresa, objLote.sOrigem, objLote.iExercicio, iPeriodo, objLote.iLote)
           
        Case 55149, 55150
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LOTE", Err, objLote.iFilialEmpresa, objLote.sOrigem, objLote.iExercicio, objLote.iPeriodo, objLote.iLote)
           
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153520)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function MvPerCta_Calcula_Valor(ByVal iFilialEmpresa As Integer, ByVal iExercicio As Integer, ByVal sContaInicial As String, ByVal sContaFinal As String, dValorExercAtual As Double, dValorExercAnt As Double) As Long
'Calcula Valor pela soma de SldIni com os valores de crédito e subtrai os débitos

Dim lErro As Long
Dim lComando1 As Long
Dim dSldIni As Double
Dim adValCre(1 To 12) As Double
Dim adValDeb(1 To 12) As Double
Dim iIndice As Integer
Dim dValor As Double
Dim iExercicio1 As Integer
    
On Error GoTo Erro_MvPerCta_Calcula_Valor

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 43604
    
    dValorExercAnt = 0
    dValorExercAtual = 0
    
    lErro = Comando_Executar(lComando1, "SELECT Exercicio, SUM(SldIni+Cre01+Cre02+Cre03+Cre04+Cre05+Cre06+Cre07+Cre08+Cre09+Cre10+Cre11+Cre12-Deb01-Deb02-Deb03-Deb04-Deb05-Deb06-Deb07-Deb08-Deb09-Deb10-Deb11-Deb12) FROM MvPerCta, PlanoConta WHERE MvPerCta.Conta = PlanoConta.Conta AND TipoConta = 3 AND FilialEmpresa = ? AND (Exercicio = ? Or Exercicio =?) AND MvPerCta.Conta >= ? AND MvPerCta.Conta <= ? GROUP BY Exercicio", _
        iExercicio1, dValor, iFilialEmpresa, iExercicio, iExercicio - 1, sContaInicial, sContaFinal)
    If lErro <> AD_SQL_SUCESSO Then Error 43592

    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then Error 43593

    Do While lErro = AD_SQL_SUCESSO
    
        If iExercicio = iExercicio1 Then
            dValorExercAtual = dValor
        Else
            dValorExercAnt = dValor
        End If
        
        lErro = Comando_BuscarProximo(lComando1)
        If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then Error 43593
        
    Loop
        
    'Libera comando
    Call Comando_Fechar(lComando1)
    
    MvPerCta_Calcula_Valor = SUCESSO
    
    Exit Function
      
Erro_MvPerCta_Calcula_Valor:

    MvPerCta_Calcula_Valor = Err
    
    Select Case Err
    
        Case 43604
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 43592, 43593
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVPERCTA", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153521)
        
    End Select
    
    'Libera comando
    Call Comando_Fechar(lComando1)

    Exit Function
    
End Function

Function MvPerCta_Calcula_Valor_Periodo(ByVal iFilialEmpresa As Integer, ByVal iExercicio As Integer, ByVal iPeriodo, ByVal sContaInicial As String, ByVal sContaFinal As String, dValorPerAnt As Double, dValorPerAtual As Double, dValorAcumPerAnt As Double, dValorAcumPerAtual As Double) As Long
'Calcula Valor pela soma de SldIni com os valores de crédito e débito do Exercicio/Periodo em questao

Dim lErro As Long
Dim lComando1 As Long
Dim dSldIni As Double
Dim dValorPeriodo As Double
Dim dValorAcumPeriodo As Double
Dim iIndice As Integer
Dim sSQL As String
Dim iExercicio1 As Integer
Dim iPeriodo1 As Integer

    
On Error GoTo Erro_MvPerCta_Calcula_Valor_Periodo

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 55898
    
    dValorPerAtual = 0
    dValorPerAnt = 0
    dValorAcumPerAtual = 0
    dValorAcumPerAnt = 0
    
    sSQL = "SUM(SldIni + "
    
    For iPeriodo1 = 1 To iPeriodo
        sSQL = sSQL + " Cre" + Format(iPeriodo1, "00") + " - Deb" + Format(iPeriodo1, "00") + " + "
    Next
    
    sSQL = left(sSQL, Len(sSQL) - 3) + ")"
    
     lErro = Comando_Executar(lComando1, "SELECT Exercicio, SUM(Cre" + Format(iPeriodo, "00") + " - Deb" + Format(iPeriodo, "00") + "), " + sSQL + " FROM MvPerCta, PlanoConta WHERE MvPerCta.Conta = PlanoConta.Conta AND TipoConta = 3 AND FilialEmpresa = ? AND (Exercicio = ? OR Exercicio = ?) AND MvPerCta.Conta >= ? AND MvPerCta.Conta <= ? GROUP BY Exercicio", _
        iExercicio1, dValorPeriodo, dValorAcumPeriodo, iFilialEmpresa, iExercicio, iExercicio - 1, sContaInicial, sContaFinal)
    If lErro <> AD_SQL_SUCESSO Then Error 55899

    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then Error 55900

    Do While lErro = AD_SQL_SUCESSO

        If iExercicio = iExercicio1 Then
            dValorPerAtual = dValorPeriodo
            dValorAcumPerAtual = dValorAcumPeriodo
        Else
            dValorPerAnt = dValorPeriodo
            dValorAcumPerAnt = dValorAcumPeriodo
        End If
        
        lErro = Comando_BuscarProximo(lComando1)
        If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then Error 43593

    Loop

    'Libera comando
    Call Comando_Fechar(lComando1)
    
    MvPerCta_Calcula_Valor_Periodo = SUCESSO
    
    Exit Function
      
Erro_MvPerCta_Calcula_Valor_Periodo:

    MvPerCta_Calcula_Valor_Periodo = Err
    
    Select Case Err
    
        Case 55898
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 55899, 55900
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVPERCTA", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153522)
        
    End Select
    
    'Libera comando
    Call Comando_Fechar(lComando1)

    Exit Function
    
End Function

Function Lancamentos_Le_Doc(ByVal objDoc As ClassLancamento_Cabecalho, ByVal colLancamentos As Collection) As Long
'Lê os lançamentos relativos ao documento passado como parâmetro e coloca-os em colLancamentos

Dim lErro As Long
Dim lComando As Long
Dim objLanc As ClassLancamento_Detalhe
Dim tLanc As typeLancamento

On Error GoTo Erro_Lancamentos_Le_Doc

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 28696

    tLanc.sConta = String(STRING_CONTA, 0)
    tLanc.sCcl = String(STRING_CCL, 0)
    tLanc.sHistorico = String(STRING_HISTORICO, 0)
    tLanc.sProduto = String(STRING_PRODUTO, 0)
    tLanc.sModelo = String(STRING_PADRAOCONTAB_MODELO, 0)
    tLanc.sDocOrigem = String(STRING_LANPENDENTE_DOCORIGEM, 0)
    
    If giSetupDocPorPeriodo = DOC_INICIALIZADO_POR_PERIODO Then
    
        lErro = Comando_Executar(lComando, "SELECT  Seq, Lote, PeriodoLote, Data, Conta, Ccl, Historico, Valor, NumIntDoc, FilialCliForn, CliForn, Transacao, SeqContraPartida, Status, Produto, ApropriaCrProd, ContaSimples, EscaninhoCusto, ValorLivroAuxiliar, ClienteFornecedor, DocOrigem, Quantidade, DataEstoque, Modelo, Gerencial FROM Lancamentos WHERE FilialEmpresa = ? AND Origem = ? AND Exercicio = ? AND PeriodoLan = ? AND Doc = ?", _
            tLanc.iSeq, tLanc.iLote, tLanc.iPeriodoLote, tLanc.dtData, tLanc.sConta, tLanc.sCcl, tLanc.sHistorico, tLanc.dValor, tLanc.lNumIntDoc, tLanc.iFilialCliForn, tLanc.lCliForn, tLanc.iTransacao, tLanc.iSeqContraPartida, tLanc.iStatus, tLanc.sProduto, tLanc.iApropriaCRProd, tLanc.lContaSimples, tLanc.iEscaninho_Custo, tLanc.dValorLivroAuxiliar, tLanc.iClienteFornecedor, tLanc.sDocOrigem, tLanc.dQuantidade, tLanc.dtDataEstoque, tLanc.sModelo, tLanc.iGerencial, objDoc.iFilialEmpresa, objDoc.sOrigem, objDoc.iExercicio, objDoc.iPeriodoLan, objDoc.lDoc)
        If lErro <> AD_SQL_SUCESSO Then Error 28697

    Else

        lErro = Comando_Executar(lComando, "SELECT  Seq, Lote, PeriodoLote, Data, Conta, Ccl, Historico, Valor, NumIntDoc, FilialCliForn, CliForn, Transacao, SeqContraPartida, Status, Produto, ApropriaCrProd, ContaSimples, EscaninhoCusto, ValorLivroAuxiliar, ClienteFornecedor, DocOrigem, Quantidade, DataEstoque, Modelo,PeriodoLan, Gerencial FROM Lancamentos WHERE FilialEmpresa = ? AND Origem = ? AND Exercicio = ? AND Doc = ?", _
            tLanc.iSeq, tLanc.iLote, tLanc.iPeriodoLote, tLanc.dtData, tLanc.sConta, tLanc.sCcl, tLanc.sHistorico, tLanc.dValor, tLanc.lNumIntDoc, tLanc.iFilialCliForn, tLanc.lCliForn, tLanc.iTransacao, tLanc.iSeqContraPartida, tLanc.iStatus, tLanc.sProduto, tLanc.iApropriaCRProd, tLanc.lContaSimples, tLanc.iEscaninho_Custo, tLanc.dValorLivroAuxiliar, tLanc.iClienteFornecedor, tLanc.sDocOrigem, tLanc.dQuantidade, tLanc.dtDataEstoque, tLanc.sModelo, tLanc.iPeriodoLan, tLanc.iGerencial, objDoc.iFilialEmpresa, objDoc.sOrigem, objDoc.iExercicio, objDoc.lDoc)
        If lErro <> AD_SQL_SUCESSO Then Error 28698

    End If

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 28699

    If lErro = AD_SQL_SEM_DADOS Then Error 28700

    Do While lErro <> AD_SQL_SEM_DADOS

        Set objLanc = New ClassLancamento_Detalhe
        
        objLanc.iSeq = tLanc.iSeq
        objLanc.iLote = tLanc.iLote
        objLanc.iPeriodoLote = tLanc.iPeriodoLote
        objLanc.dtData = tLanc.dtData
        objLanc.sConta = tLanc.sConta
        objLanc.sCcl = tLanc.sCcl
        objLanc.sHistorico = tLanc.sHistorico
        objLanc.dValor = tLanc.dValor
        objLanc.lNumIntDoc = tLanc.lNumIntDoc
        objLanc.iFilialCliForn = tLanc.iFilialCliForn
        objLanc.lCliForn = tLanc.lCliForn
        objLanc.iTransacao = tLanc.iTransacao
        objLanc.iSeqContraPartida = tLanc.iSeqContraPartida
        objLanc.iStatus = tLanc.iStatus
        objLanc.sProduto = tLanc.sProduto
        objLanc.sModelo = tLanc.sModelo
        objLanc.dValorLivroAuxiliar = tLanc.dValorLivroAuxiliar
        objLanc.iClienteFornecedor = tLanc.iClienteFornecedor
        objLanc.iPeriodoLan = tLanc.iPeriodoLan
        objLanc.sProduto = tLanc.sProduto
        objLanc.iApropriaCRProd = tLanc.iApropriaCRProd
        objLanc.lContaSimples = tLanc.lContaSimples
        objLanc.iEscaninho_Custo = tLanc.iEscaninho_Custo
        objLanc.sDocOrigem = tLanc.sDocOrigem
        objLanc.dQuantidade = tLanc.dQuantidade
        objLanc.dtDataEstoque = tLanc.dtDataEstoque
        objLanc.iGerencial = tLanc.iGerencial
               
        colLancamentos.Add objLanc

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 28701

    Loop

    Call Comando_Fechar(lComando)

    Lancamentos_Le_Doc = SUCESSO

    Exit Function

Erro_Lancamentos_Le_Doc:

    Lancamentos_Le_Doc = Err

    Select Case Err

        Case 28696
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 28697, 28698, 28699, 28701
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LANCAMENTOS3", Err)

        Case 28700

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153523)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function MnemonicoCTBValor_Le_Globais(ByVal colMnemonicoGlobal As Collection) As Long
'le todos os Mnemônicos contidos na tabela MnemonicoCTBValor e coloca-os em colMnemonicoGlobal

Dim lComando As Long
Dim lErro As Long
Dim objMnemonicoGlobal As ClassMnemonicoCTBValor
Dim tMnemonicoGlobal As typeMnemonicoCTBValor

On Error GoTo Erro_MnemonicoCTBValor_Le_Globais

    tMnemonicoGlobal.sDescricao = String(STRING_MNEMONICOCTB_DESCRICAO, 0)
    tMnemonicoGlobal.sMnemonico = String(STRING_MNEMONICOCTB_MNEMONICO, 0)
    tMnemonicoGlobal.sValor = String(STRING_MNEMONICOCTB_VALOR, 0)
    
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 39772
    
    'lê todos os registros da tabela MnemonicoCTBValor
    lErro = Comando_Executar(lComando, "SELECT Mnemonico, Descricao, Valor FROM MnemonicoCTBValor ORDER BY Mnemonico", tMnemonicoGlobal.sMnemonico, tMnemonicoGlobal.sDescricao, tMnemonicoGlobal.sValor)
    If lErro <> AD_SQL_SUCESSO Then Error 39773
    
    'le o primeiro Mnemônico
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 39774
    
    Do While lErro <> AD_SQL_SEM_DADOS

        Set objMnemonicoGlobal = New ClassMnemonicoCTBValor
        
        objMnemonicoGlobal.sMnemonico = tMnemonicoGlobal.sMnemonico
        objMnemonicoGlobal.sDescricao = tMnemonicoGlobal.sDescricao
        objMnemonicoGlobal.sValor = tMnemonicoGlobal.sValor
        
        'coloca o mnemônico lido na coleção
        colMnemonicoGlobal.Add objMnemonicoGlobal
        
        'le o próximo mnemônico
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 39775

    Loop

    Call Comando_Fechar(lComando)
    
    MnemonicoCTBValor_Le_Globais = SUCESSO

    Exit Function

Erro_MnemonicoCTBValor_Le_Globais:

    MnemonicoCTBValor_Le_Globais = Err
    
    Select Case Err

        Case 39772
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 39773, 39774, 39775
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MNEMONICOCTBVALOR", Err)
         
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153524)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function Conta_Analitica_Lock(ByVal sConta As String, ByVal sSiglaModulo As String) As Long
'O argumento sConta deve ser o número da conta FORMATADO BD
'le a conta passada como parametro e faz um lock shared

Dim lComando As Long
Dim lErro As Long
Dim tPlanoConta As typePlanoConta
Dim sSelecao As String

On Error GoTo Erro_Conta_Analitica_Lock

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 56666
    
    If sSiglaModulo = MODULO_CONTABILIDADE Then
    
        lErro = Comando_ExecutarLockado(lComando, "SELECT TipoConta, Ativo FROM PlanoConta WHERE Conta=?", tPlanoConta.iTipoConta, tPlanoConta.iAtivo, sConta)
        If lErro <> AD_SQL_SUCESSO Then Error 56667
        
    Else
    
        lErro = Comando_ExecutarLockado(lComando, "SELECT TipoConta, Ativo FROM PlanoConta WHERE Conta=? AND " & sSiglaModulo & "=1", tPlanoConta.iTipoConta, tPlanoConta.iAtivo, sConta)
        If lErro <> AD_SQL_SUCESSO Then Error 56668
    
    End If
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 56669
        
    If lErro = AD_SQL_SEM_DADOS Then Error 56670
        
    If tPlanoConta.iAtivo <> CONTA_ATIVA Then Error 56671
    
    If tPlanoConta.iTipoConta <> CONTA_ANALITICA Then Error 56672
    
    lErro = Comando_LockShared(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 56673
    
    Call Comando_Fechar(lComando)
    
    Conta_Analitica_Lock = SUCESSO
    
    Exit Function
 
Erro_Conta_Analitica_Lock:

    Conta_Analitica_Lock = Err
    
    Select Case Err
    
        Case 56666
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 56667, 56668, 56669
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTA3", Err, sConta)
        
        Case 56670
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_CADASTRADA", Err, sConta)
                
        Case 56671
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_INATIVA", Err, sConta)
        
        Case 56672
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_ANALITICA", Err, sConta)
                    
        Case 56673
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PLANOCONTA", Err, sConta)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153525)
    
    End Select
    
    Call Comando_Fechar(lComando)
      
    Exit Function

End Function


Function UM_Le(ByVal objUnidadeMedida As ClassUnidadeDeMedida) As Long
'Lê a unidade de medida a partir da Classe e da Sigla -->
'e devolve dados no Objeto

Dim lErro As Long
Dim lComando As Long
Dim tUnidMed As typeUnidadeMedida

On Error GoTo Erro_UM_Le

    tUnidMed.sNome = String(STRING_UM_NOME, 0)
    tUnidMed.sSiglaUMBase = String(STRING_UM_SIGLA_UM_BASE, 0)

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 23772

    lErro = Comando_Executar(lComando, "SELECT Nome, SiglaUMBase, Quantidade, Editavel FROM UnidadesdeMedida WHERE Classe=? AND Sigla=?", tUnidMed.sNome, tUnidMed.sSiglaUMBase, tUnidMed.dQuantidade, tUnidMed.iEditavel, objUnidadeMedida.iClasse, objUnidadeMedida.sSigla)
    If lErro <> AD_SQL_SUCESSO Then Error 23773

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 23774

    If lErro = AD_SQL_SEM_DADOS Then Error 23775

    With tUnidMed

        objUnidadeMedida.dQuantidade = .dQuantidade
        objUnidadeMedida.sNome = .sNome
        objUnidadeMedida.sSiglaUMBase = .sSiglaUMBase
        objUnidadeMedida.iEditavel = .iEditavel
        
    End With

    Call Comando_Fechar(lComando)

    UM_Le = SUCESSO

    Exit Function

Erro_UM_Le:

    UM_Le = Err

    Select Case Err

        Case 23772
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 23773, 23774
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA_UNIDADESDEMEDIDA", Err, objUnidadeMedida.iClasse, objUnidadeMedida.sSigla)

        Case 23775

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 153526)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function RateioOff_Le_TodosDistintos(ByVal colRateioOff As Collection) As Long
'Le todos os rateios off distintos do BD

Dim lErro As Long
Dim sComando_SQL As String
Dim objRateioOff As ClassRateioOff
Dim lCodigo As Long
Dim sDescricao As String
Dim lComando As Long

On Error GoTo Erro_RateioOff_Le_TodosDistintos

    sComando_SQL = "SELECT DISTINCT Codigo, Descricao FROM RateioOff"

    sDescricao = String(STRING_RATEIO_DESCRICAO, 0)

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 41527

    lErro = Comando_Executar(lComando, sComando_SQL, lCodigo, sDescricao)
    If lErro <> AD_SQL_SUCESSO Then Error 41528

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 41529

    Do While lErro <> AD_SQL_SEM_DADOS

        Set objRateioOff = New ClassRateioOff

        objRateioOff.lCodigo = lCodigo
        objRateioOff.sDescricao = sDescricao

        colRateioOff.Add objRateioOff

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 41530

    Loop

    Call Comando_Fechar(lComando)

    RateioOff_Le_TodosDistintos = SUCESSO

    Exit Function

Erro_RateioOff_Le_TodosDistintos:

    RateioOff_Le_TodosDistintos = Err

    Select Case Err

        Case 41527
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 41528, 41529, 41530
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RATEIOOFF1", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153527)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function RateioOffContas_Le_Doc(ByVal objRateioOff As ClassRateioOff, ByVal colContas As Collection) As Long
'le as contas de origem relativos ao Rateio passado como parametro e coloca-os em colContas

Dim lComando As Long
Dim lErro As Long
Dim sContaInicio As String
Dim sContaFim As String
Dim lCodigo As Long
Dim iItem As Integer
Dim objRateioOffContas As ClassRateioOffContas


On Error GoTo Erro_RateioOffContas_Le_Doc
    
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 55811

    sContaInicio = String(STRING_CONTA, 0)
    sContaFim = String(STRING_CONTA, 0)
    
    lErro = Comando_Executar(lComando, "SELECT  Codigo, Item, ContaInicial, ContaFinal  FROM RateioOffContas WHERE Codigo = ?", lCodigo, iItem, sContaInicio, sContaFim, objRateioOff.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 55812

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 55813

    Do While lErro <> AD_SQL_SEM_DADOS

        Set objRateioOffContas = New ClassRateioOffContas

        objRateioOffContas.lCodigo = lCodigo
        objRateioOffContas.iItem = iItem
        objRateioOffContas.sContaInicio = sContaInicio
        objRateioOffContas.sContaFim = sContaFim

        colContas.Add objRateioOffContas

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 55814
        
    Loop
    
    Call Comando_Fechar(lComando)
    
    RateioOffContas_Le_Doc = SUCESSO
    
    Exit Function
 
Erro_RateioOffContas_Le_Doc:

    RateioOffContas_Le_Doc = Err

    Select Case Err
    
        Case 55811
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 55812, 55813, 55814
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RATEIOOFFCONTAS", Err, objRateioOff.lCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153528)
    
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function RateioOffContas_Le_Doc1(ByVal lComando As Long, ByVal objRateioOff As ClassRateioOff, ByVal colContas As Collection) As Long
'le as contas de origem relativos ao Rateio passado como parametro e coloca-os em colContas

Dim lErro As Long
Dim sContaInicio As String
Dim sContaFim As String
Dim lCodigo As Long
Dim iItem As Integer
Dim objRateioOffContas As ClassRateioOffContas


On Error GoTo Erro_RateioOffContas_Le_Doc1
    
    sContaInicio = String(STRING_CONTA, 0)
    sContaFim = String(STRING_CONTA, 0)
    
    lErro = Comando_Executar(lComando, "SELECT  Codigo, Item, ContaInicial, ContaFinal  FROM RateioOffContas WHERE Codigo = ?", lCodigo, iItem, sContaInicio, sContaFim, objRateioOff.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 55821

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 55822

    Do While lErro <> AD_SQL_SEM_DADOS

        Set objRateioOffContas = New ClassRateioOffContas

        objRateioOffContas.lCodigo = lCodigo
        objRateioOffContas.iItem = iItem
        objRateioOffContas.sContaInicio = sContaInicio
        objRateioOffContas.sContaFim = sContaFim

        colContas.Add objRateioOffContas

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 55823
        
    Loop
    
    RateioOffContas_Le_Doc1 = SUCESSO
    
    Exit Function
 
Erro_RateioOffContas_Le_Doc1:

    RateioOffContas_Le_Doc1 = Err

    Select Case Err
    
        Case 55821, 55822, 55823
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RATEIOOFFCONTAS", Err, objRateioOff.lCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153529)
    
    End Select
    
    Exit Function

End Function

Function RelDMPL_Le_Modelo(ByVal sRelatorio As String, ByVal sModelo As String, ByVal colRel As Collection) As Long
'Le o modelo em questão e seu componentes

Dim lErro As Long, iIndice As Integer
Dim lComando As Long
Dim sComando_SQL As String
Dim tRel As typeRelDMPL
Dim objRel As ClassRelDMPL

On Error GoTo Erro_RelDMPL_Le_Modelo

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 60578

    tRel.sModelo = String(STRING_DRE_MODELO, 0)
    tRel.sTitulo = String(STRING_DRE_TITULO, 0)

    sComando_SQL = "SELECT Modelo, Linha, Coluna, Tipo, Titulo, Exercicio FROM RelDMPL WHERE Relatorio = ? AND Modelo = ? ORDER BY Tipo, Exercicio, Linha, Coluna"

    lErro = Comando_Executar(lComando, sComando_SQL, tRel.sModelo, tRel.iLinha, tRel.iColuna, tRel.iTipo, tRel.sTitulo, tRel.iExercicio, sRelatorio, sModelo)
    If lErro <> AD_SQL_SUCESSO Then Error 60579

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 60580

    Do While lErro <> AD_SQL_SEM_DADOS

        Set objRel = New ClassRelDMPL

        objRel.sModelo = tRel.sModelo
        objRel.iLinha = tRel.iLinha
        objRel.iColuna = tRel.iColuna
        objRel.iTipo = tRel.iTipo
        objRel.sTitulo = tRel.sTitulo
        objRel.iExercicio = tRel.iExercicio

        colRel.Add objRel

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 60581

    Loop

    Call Comando_Fechar(lComando)

    RelDMPL_Le_Modelo = SUCESSO

    Exit Function

Erro_RelDMPL_Le_Modelo:

    RelDMPL_Le_Modelo = Err

    Select Case Err

        Case 60578
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 60579, 60580, 60581
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELDMPL_MODELO", Err, sModelo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153530)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function RelDMPLConta_Le_Modelo(ByVal sRelatorio As String, ByVal sModelo As String, ByVal colRelConta As Collection) As Long
'Le os dados de Contas do modelo atual para a coleção

Dim lErro As Long
Dim sComando_SQL As String
Dim tRelConta As typeRelDMPLConta
Dim objRelConta As ClassRelDMPLConta
Dim lComando As Long

On Error GoTo Erro_RelDMPLConta_Le_Modelo

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 60582

    tRelConta.sModelo = String(STRING_DRE_MODELO, 0)
    tRelConta.sContaFinal = String(STRING_CONTA, 0)
    tRelConta.sContaInicial = String(STRING_CONTA, 0)

    sComando_SQL = "SELECT Modelo, Linha, Coluna, Item, ContaInicial, ContaFinal FROM RelDMPLConta WHERE Relatorio = ? AND Modelo = ? ORDER BY Linha, Coluna, Item"

    lErro = Comando_Executar(lComando, sComando_SQL, tRelConta.sModelo, tRelConta.iLinha, tRelConta.iColuna, tRelConta.iItem, tRelConta.sContaInicial, tRelConta.sContaFinal, sRelatorio, sModelo)
    If lErro <> AD_SQL_SUCESSO Then Error 60583

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 60584

    Do While lErro <> AD_SQL_SEM_DADOS

        Set objRelConta = New ClassRelDMPLConta

        objRelConta.sModelo = tRelConta.sModelo
        objRelConta.iLinha = tRelConta.iLinha
        objRelConta.iColuna = tRelConta.iColuna
        objRelConta.iItem = tRelConta.iItem
        objRelConta.sContaInicial = tRelConta.sContaInicial
        objRelConta.sContaFinal = tRelConta.sContaFinal

        colRelConta.Add objRelConta

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 60585

    Loop

    Call Comando_Fechar(lComando)

    RelDMPLConta_Le_Modelo = SUCESSO

    Exit Function

Erro_RelDMPLConta_Le_Modelo:

    RelDMPLConta_Le_Modelo = Err

    Select Case Err

        Case 60582
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 60583, 60584, 60585
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELDMPLCONTA", Err, sModelo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153531)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function RelDMPLFormula_Le_Modelo(ByVal sRelatorio As String, ByVal sModelo As String, ByVal colRelFormula As Collection) As Long
'Le os dados de Fórmulas do modelo atual para a coleção

Dim lErro As Long
Dim sComando_SQL As String
Dim tRelFormula As typeRelDMPLFormula
Dim objRelFormula As ClassRelDMPLFormula
Dim lComando As Long

On Error GoTo Erro_RelDMPLFormula_Le_Modelo

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 60586

    tRelFormula.sModelo = String(STRING_DRE_MODELO, 0)

    sComando_SQL = "SELECT Modelo, Linha, Coluna, Item, LinhaFormula, ColunaFormula, Operacao FROM RelDMPLFormula WHERE Relatorio = ? AND Modelo = ? ORDER BY Linha, Coluna, Item"

    lErro = Comando_Executar(lComando, sComando_SQL, tRelFormula.sModelo, tRelFormula.iLinha, tRelFormula.iColuna, tRelFormula.iItem, tRelFormula.iLinhaFormula, tRelFormula.iColunaFormula, tRelFormula.iOperacao, sRelatorio, sModelo)
    If lErro <> AD_SQL_SUCESSO Then Error 60587

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 60588

    Do While lErro <> AD_SQL_SEM_DADOS

        Set objRelFormula = New ClassRelDMPLFormula

        objRelFormula.sModelo = tRelFormula.sModelo
        objRelFormula.iLinha = tRelFormula.iLinha
        objRelFormula.iColuna = tRelFormula.iColuna
        objRelFormula.iItem = tRelFormula.iItem
        objRelFormula.iLinhaFormula = tRelFormula.iLinhaFormula
        objRelFormula.iColunaFormula = tRelFormula.iColunaFormula
        objRelFormula.iOperacao = tRelFormula.iOperacao

        colRelFormula.Add objRelFormula

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 60589

    Loop

    Call Comando_Fechar(lComando)

    RelDMPLFormula_Le_Modelo = SUCESSO

    Exit Function

Erro_RelDMPLFormula_Le_Modelo:

    RelDMPLFormula_Le_Modelo = Err

    Select Case Err

        Case 60586
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 60587, 60588, 60589
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELDMPLFORMULA", Err, sModelo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153532)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function RelDMPL_Le_Modelos_Distintos(ByVal sRelatorio As String, ByVal colModelos As Collection) As Long
'Le todos os modelos disponíveis no BD

Dim lErro As Long
Dim sComando_SQL As String
Dim sModelo As String
Dim lComando As Long

On Error GoTo Erro_RelDMPL_Le_Modelos_Distintos

    sModelo = String(STRING_DRE_MODELO, 0)

    sComando_SQL = "SELECT DISTINCT Modelo FROM RelDMPL WHERE Relatorio = ?"

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 60606

    lErro = Comando_Executar(lComando, sComando_SQL, sModelo, sRelatorio)
    If lErro <> AD_SQL_SUCESSO Then Error 60607

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 60608

    Do While lErro <> AD_SQL_SEM_DADOS

        colModelos.Add sModelo

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 60609

    Loop

    Call Comando_Fechar(lComando)

    RelDMPL_Le_Modelos_Distintos = SUCESSO

    Exit Function

Erro_RelDMPL_Le_Modelos_Distintos:

    RelDMPL_Le_Modelos_Distintos = Err

    Select Case Err

        Case 60606
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 60607, 60608, 60609
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELDRE", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153533)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function MvPerCta_Calcula_Valor1(ByVal iFilialEmpresa As Integer, ByVal iExercicio As Integer, ByVal sContaInicial As String, ByVal sContaFinal As String, dValorExercAtual As Double) As Long
'Calcula Valor pela soma de SldIni com os valores de crédito e subtrai os débitos

Dim lErro As Long
Dim lComando1 As Long
Dim dSldIni As Double
Dim adValCre(1 To 12) As Double
Dim adValDeb(1 To 12) As Double
Dim iIndice As Integer
Dim dValor As Double
    
On Error GoTo Erro_MvPerCta_Calcula_Valor1

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 60676
    
    dValorExercAtual = 0
    
    lErro = Comando_Executar(lComando1, "SELECT SUM(SldIni+Cre01+Cre02+Cre03+Cre04+Cre05+Cre06+Cre07+Cre08+Cre09+Cre10+Cre11+Cre12-Deb01-Deb02-Deb03-Deb04-Deb05-Deb06-Deb07-Deb08-Deb09-Deb10-Deb11-Deb12) FROM MvPerCta, PlanoConta WHERE MvPerCta.Conta = PlanoConta.Conta AND TipoConta = 3 AND FilialEmpresa = ? AND Exercicio = ? AND MvPerCta.Conta >= ? AND MvPerCta.Conta <= ?", _
        dValor, iFilialEmpresa, iExercicio, sContaInicial, sContaFinal)
    If lErro <> AD_SQL_SUCESSO Then Error 60677

    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then Error 60678

    If lErro = AD_SQL_SUCESSO Then dValorExercAtual = dValor
        
    'Libera comando
    Call Comando_Fechar(lComando1)
    
    MvPerCta_Calcula_Valor1 = SUCESSO
    
    Exit Function
      
Erro_MvPerCta_Calcula_Valor1:

    MvPerCta_Calcula_Valor1 = Err
    
    Select Case Err
    
        Case 60676
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 60677, 60678
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVPERCTA", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153534)
        
    End Select
    
    'Libera comando
    Call Comando_Fechar(lComando1)

    Exit Function
    
End Function

Function MvPerCta_Calcula_Valor1_Periodo(ByVal iFilialEmpresa As Integer, ByVal iExercicio As Integer, ByVal iPeriodo As Integer, ByVal sContaInicial As String, ByVal sContaFinal As String, dValorExercAtual As Double) As Long
'Calcula Valor pela soma de SldIni com os valores de crédito e subtrai os débitos

Dim lErro As Long
Dim lComando1 As Long
Dim dSldIni As Double
Dim adValCre(1 To 12) As Double
Dim adValDeb(1 To 12) As Double
Dim iIndice As Integer
Dim dValor As Double
    
On Error GoTo Erro_MvPerCta_Calcula_Valor1_Periodo

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 60676
    
    dValorExercAtual = 0
    
    lErro = Comando_Executar(lComando1, "SELECT SUM(SldIni+Cre" & Format(iPeriodo, "00") & " -Deb" & Format(iPeriodo, "00") & ") FROM MvPerCta, PlanoConta WHERE MvPerCta.Conta = PlanoConta.Conta AND TipoConta = 3 AND FilialEmpresa = ? AND Exercicio = ? AND MvPerCta.Conta >= ? AND MvPerCta.Conta <= ?", _
        dValor, iFilialEmpresa, iExercicio, sContaInicial, sContaFinal)
    If lErro <> AD_SQL_SUCESSO Then Error 60677

    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then Error 60678

    If lErro = AD_SQL_SUCESSO Then dValorExercAtual = dValor
        
    'Libera comando
    Call Comando_Fechar(lComando1)
    
    MvPerCta_Calcula_Valor1_Periodo = SUCESSO
    
    Exit Function
      
Erro_MvPerCta_Calcula_Valor1_Periodo:

    MvPerCta_Calcula_Valor1_Periodo = Err
    
    Select Case Err
    
        Case 60676
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 60677, 60678
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVPERCTA", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153534)
        
    End Select
    
    'Libera comando
    Call Comando_Fechar(lComando1)

    Exit Function
    
End Function

Function Lote_Gera_Automatico(ByVal iFilialEmpresa As Integer, ByVal sCodUsuario As String, ByVal sOrigem As String, ByVal iExercicio As Integer, ByVal iPeriodo As Integer, iLote As Integer) As Long
'retorna o lote automatico para a contabilidade caso o usuário tenha optado por geração automatica do numero do lote

Dim lErro As Long
Dim iLote1 As Integer
Dim alComando(0 To 3) As Long
Dim iIndice As Integer

On Error GoTo Erro_Lote_Gera_Automatico

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 60880
    Next

    'pesquisa um lote já registrado para este modulo
    lErro = Comando_Executar(alComando(0), "SELECT Lote FROM UsuarioLote WHERE CodUsuario = ? AND Origem=?", iLote, sCodUsuario, sOrigem)
    If lErro <> AD_SQL_SUCESSO Then Error 60881
    
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 60882
    
    'se não tiver um lote já registrado para este modulo
    If lErro = AD_SQL_SEM_DADOS Then
    
        'Pesquisa as numerações de lote pendentes para esta origem e o periodo em questão
        lErro = Comando_Executar(alComando(1), "SELECT Lote FROM LotePendente WHERE FilialEmpresa = ? AND Origem = ? AND Exercicio = ? AND Periodo = ? ORDER BY Lote", iLote, iFilialEmpresa, sOrigem, iExercicio, iPeriodo)
        If lErro <> AD_SQL_SUCESSO Then Error 60883
            
        'le o primeiro lote pendente
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 60884
    
        'não há lote pendente para o modulo em questão.
        If lErro <> AD_SQL_SUCESSO Then Error 60885
    
    Else
    
        'se já tiver um lote cadastrado, verifica se está pendente e se pertence ao periodo em questão
        lErro = Comando_Executar(alComando(2), "SELECT Lote FROM LotePendente WHERE FilialEmpresa = ? AND Origem = ? AND Exercicio = ? AND Periodo = ? AND Lote = ?", iLote1, iFilialEmpresa, sOrigem, iExercicio, iPeriodo, iLote)
        If lErro <> AD_SQL_SUCESSO Then Error 60886
        
        'le o lote pendente
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 60887
    
        'se o lote que está guardado não está mais pendente ou não pertence ao periodo em questão
        If lErro <> AD_SQL_SUCESSO Then
    
            'Pesquisa as numerações de lote pendentes para esta origem e o periodo em questão
            lErro = Comando_Executar(alComando(1), "SELECT Lote FROM LotePendente WHERE FilialEmpresa = ? AND Origem = ? AND Exercicio = ? AND Periodo = ? ORDER BY Lote", iLote, iFilialEmpresa, sOrigem, iExercicio, iPeriodo)
            If lErro <> AD_SQL_SUCESSO Then Error 60888
                
            'le o primeiro lote pendente
            lErro = Comando_BuscarPrimeiro(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 60889
        
            'não há lote pendente para o modulo em questão.
            If lErro <> AD_SQL_SUCESSO Then Error 60890
    
        End If
    
    End If
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Lote_Gera_Automatico = SUCESSO
        
    Exit Function
    
    
Erro_Lote_Gera_Automatico:

    Lote_Gera_Automatico = Err

    Select Case Err
    
    Case 60880
        lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
    Case 60881, 60882
        lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_USUARIOLOTE", Err, sCodUsuario, sOrigem)
        
    Case 60883, 60884, 60888, 60889
        lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LOTEPENDENTE5", Err, iFilialEmpresa, sOrigem, iExercicio, iPeriodo)
        
    Case 60885, 60890
        
    Case 60886, 60887
        lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LOTEPENDENTE", Err, iFilialEmpresa, sOrigem, iExercicio, iPeriodo, iLote)
        
    Case Else
        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153535)
        
    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function CustoMedioProducaoApurado_Le_Mes(ByVal sProduto As String, ByVal iMes As Integer, ByVal iAno As Integer, dCMPAtual As Double) As Long
'Calcula Custo Médio de Produção Apurado do Produto passado
'Se o mês  passado ainda não foi apurado retorna ZERO

Dim lErro As Long
Dim objEstoqueMes As New ClassEstoqueMes

On Error GoTo Erro_CustoMedioProducaoApurado_Le_Mes

    objEstoqueMes.iFilialEmpresa = giFilialEmpresa
    
    lErro = CF("EstoqueMes_Le_Apurado", objEstoqueMes)
    If lErro <> SUCESSO And lErro <> 46225 Then gError 64497
        
    'Se já existe um Mes apurado
    If lErro = SUCESSO Then
        
        'Se o Ano o Mês passado for Menor ou igual ao ultimo Ano-Mes apurado
        If iAno <= objEstoqueMes.iAno And iMes <= objEstoqueMes.iMes Then
            'calcula o custo medio de produção do produto em questão
            lErro = CF("Calcula_CustoMedioProducao", giFilialEmpresa, sProduto, iAno, iMes, dCMPAtual)
            If lErro <> SUCESSO And lErro <> 25433 And lErro <> 55052 Then gError 64498
            
            'Se não encontrou o SldMesEst
            If lErro = 25433 Then gError 69860
            
            'Se não encontrou o SldMesEst
            If lErro = 55052 Then gError 69861
            
        End If
    
    End If
    
    CustoMedioProducaoApurado_Le_Mes = SUCESSO

    Exit Function

Erro_CustoMedioProducaoApurado_Le_Mes:

    CustoMedioProducaoApurado_Le_Mes = gErr

    Select Case gErr

        Case 64497, 64498 'Tratado na rotina chamada
        
        Case 69860
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SLDMESEST_NAO_CADASTRADO", gErr, giFilialEmpresa, iAno, sProduto)
        
        Case 69861
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SLDMESEST_NAO_CADASTRADO1", gErr, giFilialEmpresa, sProduto)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153536)

    End Select

    Exit Function

End Function


Function EstoqueMes_Le_Apurado(ByVal objEstoqueMes As ClassEstoqueMes) As Long
'Le o último Ano/Mes apurado

Dim lErro As Long
Dim lComando As Long
Dim tEstoqueMes As typeEstoqueMes

On Error GoTo Erro_EstoqueMes_Le_Apurado

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 46222

    lErro = Comando_Executar(lComando, "SELECT Ano, Mes FROM EstoqueMes WHERE CustoProdApurado = ? AND FilialEmpresa = ? ORDER BY Ano DESC, Mes DESC", tEstoqueMes.iAno, tEstoqueMes.iMes, CUSTO_APURADO, objEstoqueMes.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then Error 46223

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 46224

    If lErro = AD_SQL_SEM_DADOS Then Error 46225
        
    objEstoqueMes.iAno = tEstoqueMes.iAno
    objEstoqueMes.iMes = tEstoqueMes.iMes

    'Fechamento comando
    Call Comando_Fechar(lComando)

    EstoqueMes_Le_Apurado = SUCESSO

    Exit Function

Erro_EstoqueMes_Le_Apurado:

    EstoqueMes_Le_Apurado = Err

    Select Case Err

        Case 46222
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 45125, 45126, 46223, 46224
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ESTOQUEMES", Err, objEstoqueMes.iFilialEmpresa, objEstoqueMes.iAno, objEstoqueMes.iMes)

        Case 46225

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153537)

    End Select

   'Fechamento comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Calcula_CustoMedioProducao(ByVal iFilialEmpresa As Integer, ByVal sProduto As String, ByVal iAno As Integer, ByVal iMes As Integer, dCMPAtual As Double) As Long
'Calcula o Custo Médio de Produção do Produto sProduto
'iAno = indica o ano de onde deve ser lido os dados de SldMesEst
'iMes = indica o ultimo mes apurado ou zero. Zero indica que não houve apuração até o momento

Dim lErro As Long
Dim objEstoqueMes As New ClassEstoqueMes
Dim objSldMesEst As New ClassSldMesEst
Dim iIndice As Integer
Dim dQuantAcumulada As Double
Dim dValorAcumulado As Double
Dim iMesFinal As Integer

On Error GoTo Erro_Calcula_CustoMedioProducao

    'se iMes = 0 ==> não há nenhum mes apurado ==> o custo medio de produção é o inicial
'    If iMes = 0 Then
    
    objSldMesEst.sProduto = sProduto
    objSldMesEst.iFilialEmpresa = iFilialEmpresa

    'Le os dados iniciais do registro da tabela SldMesEst relativo ao Primeiro Ano do Produto (FilialEmpresa pega em giFilialEmpresa)
    lErro = CF("SldMesEst_Le_PrimeiroAno", objSldMesEst)
    If lErro <> SUCESSO And lErro <> 55050 Then Error 55051

    If lErro = 55050 Then Error 55052

    dCMPAtual = objSldMesEst.dCustoMedioProducaoInicial
    
'    Else
    
    If iMes <> 0 And objSldMesEst.iAno <= iAno Then
        
        Set objSldMesEst = New ClassSldMesEst
        dCMPAtual = 0

        objSldMesEst.iAno = iAno
        objSldMesEst.sProduto = sProduto
        objSldMesEst.iFilialEmpresa = iFilialEmpresa

        lErro = CF("SldMesEst_Le", objSldMesEst)
        If lErro <> SUCESSO And lErro <> 25429 Then Error 25432
        
        If lErro = 25429 Then Error 25433
        
        'Cálculo de saldo de quantidade
        '-------------------------------
        
        'Quantidade no início do ano
        'dQuantAcumulada = objSldMesEst.dQuantInicial
        dQuantAcumulada = objSldMesEst.dQuantInicialCusto
    
        'Acumula entrada e saída dos meses
        For iIndice = 1 To iMes
            'dQuantAcumulada = dQuantAcumulada + objSldMesEst.dQuantEnt(iIndice) - objSldMesEst.dQuantSai(iIndice)
            dQuantAcumulada = dQuantAcumulada + objSldMesEst.dSaldoQuantCusto(iIndice)
        Next
        
        If dQuantAcumulada > QTDE_ESTOQUE_DELTA2 Then
        
            'Cálculo de saldo de valor
            '-------------------------
            
            'Valor inicial
            'dValorAcumulado = objSldMesEst.dValorInicial
            dValorAcumulado = objSldMesEst.dValorInicialCusto
            
            'Meses
            For iIndice = 1 To iMes
                'dValorAcumulado = dValorAcumulado + objSldMesEst.dValorEnt(iIndice) - objSldMesEst.dValorSai(iIndice)
                dValorAcumulado = dValorAcumulado + objSldMesEst.dSaldoValorCusto(iIndice)
            Next
            
            'Calcula CustoMedioProducaoAtual
            '-------------------------------
            dCMPAtual = dValorAcumulado / dQuantAcumulada
            
        ElseIf dQuantAcumulada = 0 Then
        
            'Procura o último mês em que houve saída (apropr=CMP)
            For iMesFinal = iMes To 1 Step -1
            
                If objSldMesEst.dSaldoQuantCusto(iMesFinal) > 0 Then Exit For
            
            Next
        
            'se houve saida em algum mes
            If iMesFinal <> 0 Then
            
                'Custo Médio é o valor da última saída mensal dividido pela quantidade
                dCMPAtual = objSldMesEst.dSaldoValorCusto(iMesFinal) / objSldMesEst.dSaldoQuantCusto(iMesFinal)
                            
            Else 'Todas as quantidades de saída e de entrada do Produto estão zeradas nesse ano
            
                dCMPAtual = objSldMesEst.dCustoMedioProducaoInicial
            
            End If
        
        End If
    
    End If
    
    Calcula_CustoMedioProducao = SUCESSO

    Exit Function

Erro_Calcula_CustoMedioProducao:

    Calcula_CustoMedioProducao = Err

    Select Case Err

        Case 25432, 55051 'Tratado na rotina chamada
        
        Case 25433
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_SLDMESEST_NAO_CADASTRADO", Err, iFilialEmpresa, objSldMesEst.iAno, objSldMesEst.sProduto)
        
        Case 55052
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_SLDMESEST_NAO_CADASTRADO1", Err, iFilialEmpresa, objSldMesEst.sProduto)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153538)

    End Select

    Exit Function

End Function


Function SldMesEst_Le_PrimeiroAno(ByVal objSldMesEst As ClassSldMesEst) As Long
'Le os dados iniciais do registro da tabela SldMesEst relativo ao Primeiro Ano do Produto (FilialEmpresa pega em giFilialEmpresa)

Dim lErro As Long
Dim lComando As Long
Dim tSldMesEst As typeSldMesEst

On Error GoTo Erro_SldMesEst_Le_PrimeiroAno

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 55047
    
    With tSldMesEst
    
        lErro = Comando_Executar(lComando, "SELECT Ano, CustoMedioProducaoInicial, QuantInicial, ValorInicial  FROM SldMesEst WHERE  FilialEmpresa = ? AND Produto = ? ORDER BY Ano", .iAno, .dCustoMedioProducaoInicial, .dQuantInicial, .dValorInicial, objSldMesEst.iFilialEmpresa, objSldMesEst.sProduto)
    
    End With
    
    If lErro <> AD_SQL_SUCESSO Then Error 55048

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 55049
    
    If lErro = AD_SQL_SEM_DADOS Then Error 55050
    
    With tSldMesEst
        
        objSldMesEst.dCustoMedioProducaoInicial = .dCustoMedioProducaoInicial
        objSldMesEst.dQuantInicial = .dQuantInicial
        objSldMesEst.dValorInicial = .dValorInicial
        objSldMesEst.iAno = .iAno
    
    End With

    Call Comando_Fechar(lComando)

    SldMesEst_Le_PrimeiroAno = SUCESSO

    Exit Function

Erro_SldMesEst_Le_PrimeiroAno:

    SldMesEst_Le_PrimeiroAno = Err

    Select Case Err

        Case 55047
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 55048, 55049
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESEST3", Err, objSldMesEst.iFilialEmpresa, objSldMesEst.sProduto)
        
        Case 55050  'Tratado na rotina chamadora
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153539)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function SldMesEst_Le(ByVal objSldMesEst As ClassSldMesEst) As Long
'Le na tabela SldMesEst registro relativo a Ano, Produto (FilialEmpresa pega em giFilialEmpresa)

Dim lErro As Long
Dim lComando As Long
Dim sComandoSQL As String
Dim tSldMesEst As typeSldMesEst
Dim iMes As Integer
Dim iIndice As Integer
Dim sMes As String
Dim sCampo(1 To 15) As String

On Error GoTo Erro_SldMesEst_Le

    sCampo(1) = "QuantEnt"
    sCampo(2) = "ValorEnt"
    sCampo(3) = "QuantSai"
    sCampo(4) = "QuantCons"
    sCampo(5) = "QuantVend"
    sCampo(6) = "ValorSai"
    sCampo(7) = "CustoMedio"
    sCampo(8) = "CustoProducao"
    sCampo(9) = "CustoStandard"
    sCampo(10) = "Status"
    sCampo(11) = "QuantComp"
    sCampo(12) = "ValorComp"
    sCampo(13) = "ValorVend"
    sCampo(14) = "SaldoQuantCusto"
    sCampo(15) = "SaldoValorCusto"
    
    sComandoSQL = "SELECT CustoMedioProducaoInicial, QuantInicial, ValorInicial, QuantInicialCusto, ValorInicialCusto"
    
    For iIndice = LBound(sCampo) To UBound(sCampo)
        For iMes = 1 To 12
            
            sMes = CStr(iMes)
            sComandoSQL = sComandoSQL & ", " & sCampo(iIndice) & sMes
    
        Next
    Next
    
    sComandoSQL = sComandoSQL & " FROM SldMesEst WHERE  Ano = ? AND FilialEmpresa = ? AND Produto = ?"
    
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 25426
    
    With tSldMesEst
    
        lErro = Comando_Executar(lComando, sComandoSQL, .dCustoMedioProducaoInicial, .dQuantInicial, .dValorInicial, .dQuantInicialCusto, .dValorInicialCusto, .adQuantEnt(1), .adQuantEnt(2), .adQuantEnt(3), .adQuantEnt(4), .adQuantEnt(5), .adQuantEnt(6), .adQuantEnt(7), .adQuantEnt(8), .adQuantEnt(9), .adQuantEnt(10), .adQuantEnt(11), .adQuantEnt(12), .adValorEnt(1), .adValorEnt(2), .adValorEnt(3), .adValorEnt(4), .adValorEnt(5), .adValorEnt(6), .adValorEnt(7), .adValorEnt(8), .adValorEnt(9), .adValorEnt(10), .adValorEnt(11), .adValorEnt(12), _
            .adQuantSai(1), .adQuantSai(2), .adQuantSai(3), .adQuantSai(4), .adQuantSai(5), .adQuantSai(6), .adQuantSai(7), .adQuantSai(8), .adQuantSai(9), .adQuantSai(10), .adQuantSai(11), .adQuantSai(12), .adQuantCons(1), .adQuantCons(2), .adQuantCons(3), .adQuantCons(4), .adQuantCons(5), .adQuantCons(6), .adQuantCons(7), .adQuantCons(8), .adQuantCons(9), .adQuantCons(10), .adQuantCons(11), .adQuantCons(12), _
            .adQuantVend(1), .adQuantVend(2), .adQuantVend(3), .adQuantVend(4), .adQuantVend(5), .adQuantVend(6), .adQuantVend(7), .adQuantVend(8), .adQuantVend(9), .adQuantVend(10), .adQuantVend(11), .adQuantVend(12), .adValorSai(1), .adValorSai(2), .adValorSai(3), .adValorSai(4), .adValorSai(5), .adValorSai(6), .adValorSai(7), .adValorSai(8), .adValorSai(9), .adValorSai(10), .adValorSai(11), .adValorSai(12), _
            .adCustoMedio(1), .adCustoMedio(2), .adCustoMedio(3), .adCustoMedio(4), .adCustoMedio(5), .adCustoMedio(6), .adCustoMedio(7), .adCustoMedio(8), .adCustoMedio(9), .adCustoMedio(10), .adCustoMedio(11), .adCustoMedio(12), .adCustoProducao(1), .adCustoProducao(2), .adCustoProducao(3), .adCustoProducao(4), .adCustoProducao(5), .adCustoProducao(6), .adCustoProducao(7), .adCustoProducao(8), .adCustoProducao(9), .adCustoProducao(10), .adCustoProducao(11), .adCustoProducao(12), _
            .adCustoStandard(1), .adCustoStandard(2), .adCustoStandard(3), .adCustoStandard(4), .adCustoStandard(5), .adCustoStandard(6), .adCustoStandard(7), .adCustoStandard(8), .adCustoStandard(9), .adCustoStandard(10), .adCustoStandard(11), .adCustoStandard(12), .aiStatus(1), .aiStatus(2), .aiStatus(3), .aiStatus(4), .aiStatus(5), .aiStatus(6), .aiStatus(7), .aiStatus(8), .aiStatus(9), .aiStatus(10), .aiStatus(11), .aiStatus(12), _
            .adQuantComp(1), .adQuantComp(2), .adQuantComp(3), .adQuantComp(4), .adQuantComp(5), .adQuantComp(6), .adQuantComp(7), .adQuantComp(8), .adQuantComp(9), .adQuantComp(10), .adQuantComp(11), .adQuantComp(12), .adValorComp(1), .adValorComp(2), .adValorComp(3), .adValorComp(4), .adValorComp(5), .adValorComp(6), .adValorComp(7), .adValorComp(8), .adValorComp(9), .adValorComp(10), .adValorComp(11), .adValorComp(12), _
            .adValorVend(1), .adValorVend(2), .adValorVend(3), .adValorVend(4), .adValorVend(5), .adValorVend(6), .adValorVend(7), .adValorVend(8), .adValorVend(9), .adValorVend(10), .adValorVend(11), .adValorVend(12), _
            .adSaldoQuantCusto(1), .adSaldoQuantCusto(2), .adSaldoQuantCusto(3), .adSaldoQuantCusto(4), .adSaldoQuantCusto(5), .adSaldoQuantCusto(6), .adSaldoQuantCusto(7), .adSaldoQuantCusto(8), .adSaldoQuantCusto(9), .adSaldoQuantCusto(10), .adSaldoQuantCusto(11), .adSaldoQuantCusto(12), _
            .adSaldoValorCusto(1), .adSaldoValorCusto(2), .adSaldoValorCusto(3), .adSaldoValorCusto(4), .adSaldoValorCusto(5), .adSaldoValorCusto(6), .adSaldoValorCusto(7), .adSaldoValorCusto(8), .adSaldoValorCusto(9), .adSaldoValorCusto(10), .adSaldoValorCusto(11), .adSaldoValorCusto(12), _
            objSldMesEst.iAno, objSldMesEst.iFilialEmpresa, objSldMesEst.sProduto)
    
    End With
    
    If lErro <> AD_SQL_SUCESSO Then Error 25427

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 25428
    If lErro = AD_SQL_SEM_DADOS Then Error 25429
    
    With tSldMesEst
        
        objSldMesEst.dCustoMedioProducaoInicial = .dCustoMedioProducaoInicial
        objSldMesEst.dQuantInicial = .dQuantInicial
        objSldMesEst.dValorInicial = .dValorInicial
        objSldMesEst.dQuantInicialCusto = .dQuantInicialCusto
        objSldMesEst.dValorInicialCusto = .dValorInicialCusto
    
        For iMes = 1 To 12
        
            objSldMesEst.dQuantEnt(iMes) = .adQuantEnt(iMes)
            objSldMesEst.dValorEnt(iMes) = .adValorEnt(iMes)
            objSldMesEst.dQuantSai(iMes) = .adQuantSai(iMes)
            objSldMesEst.dQuantCons(iMes) = .adQuantCons(iMes)
            objSldMesEst.dQuantVend(iMes) = .adQuantVend(iMes)
            objSldMesEst.dValorSai(iMes) = .adValorSai(iMes)
            objSldMesEst.dCustoMedio(iMes) = .adCustoMedio(iMes)
            objSldMesEst.dCustoProducao(iMes) = .adCustoProducao(iMes)
            objSldMesEst.dCustoStandard(iMes) = .adCustoStandard(iMes)
            objSldMesEst.iStatus(iMes) = .aiStatus(iMes)
            objSldMesEst.dQuantComp(iMes) = .adQuantComp(iMes)
            objSldMesEst.dValorComp(iMes) = .adValorComp(iMes)
            objSldMesEst.dValorVend(iMes) = .adValorVend(iMes)
            objSldMesEst.dSaldoQuantCusto(iMes) = .adSaldoQuantCusto(iMes)
            objSldMesEst.dSaldoValorCusto(iMes) = .adSaldoValorCusto(iMes)
        
        Next
    
    End With

    Call Comando_Fechar(lComando)

    SldMesEst_Le = SUCESSO

    Exit Function

Erro_SldMesEst_Le:

    SldMesEst_Le = Err

    Select Case Err

        Case 25426
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 25427, 25428
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESEST", Err, objSldMesEst.iAno, objSldMesEst.iFilialEmpresa, objSldMesEst.sProduto)
        
        Case 25429  'Tratado na rotina chamadora
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153540)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function SldMesEst_Le_CustoRealProducao(ByVal lComando As Long, ByVal iFilialEmpresa As Integer, ByVal iAno As Integer, ByVal sProduto As String, ByVal iMes As Integer, dCustoRealProducao As Double) As Long
'le os dados de custo real de Produção de SldMesEst relativos ao Ano, FilialEmpresa, Produto e mes passados como parametro
    
Dim lErro As Long
    
On Error GoTo Erro_SldMesEst_Le_CustoRealProducao
    
    'seleciona o custo standard de SldMesEst relativo ao Ano, FilialEmpresa, Produto, Mes passados como parametro
    lErro = Comando_ExecutarPos(lComando, "SELECT CustoProducao" + CStr(iMes) + " FROM SldMesEst WHERE Ano=? AND FilialEmpresa=? AND Produto=?", 0, dCustoRealProducao, iAno, iFilialEmpresa, sProduto)
    If lErro <> AD_SQL_SUCESSO Then Error 60908
    
    'le o SldMesEst
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 60909
    
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 60910
    
    SldMesEst_Le_CustoRealProducao = SUCESSO

    Exit Function

Erro_SldMesEst_Le_CustoRealProducao:
    
    SldMesEst_Le_CustoRealProducao = Err
    
    Select Case Err
    
        Case 60908, 60909
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESEST", Err, iAno, iFilialEmpresa, sProduto)
            
        Case 60910
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_SLDMESEST", Err, iAno, iFilialEmpresa, sProduto)
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153541)
        
    End Select
    
    Exit Function

End Function

Function Ccl_Le_Total(ByVal iTotalCcl As Integer) As Long
'Lê a quantidade de Ccl's já gravados na Tabela de Ccl e que são analíticos

Dim lErro As Long
Dim lComando As Long

On Error GoTo Erro_Ccl_Le_Total

    'Abre comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 64080

    'Seleciona a quantidade de Ccl Gravados
    lErro = Comando_Executar(lComando, "SELECT COUNT(*) FROM Ccl WHERE TipoCcl = ?", iTotalCcl, CCL_ANALITICA)
    If lErro <> AD_SQL_SUCESSO Then Error 64081
    
    'Tenta ler
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 64082
            
    'Fecha comando
    Call Comando_Fechar(lComando)
        
    Ccl_Le_Total = SUCESSO
    
    Exit Function
    
Erro_Ccl_Le_Total:
    
    Ccl_Le_Total = Err
    
    Select Case Err
    
        Case 64080
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 64081, 64082
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CCL1", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153542)
            
    End Select
        
    'Fecha comando
    Call Comando_Fechar(lComando)
        
    Exit Function
        
End Function

Function SldMesEst1_Le_CustoMedioConserto3(ByVal lComando As Long, ByVal iFilialEmpresa As Integer, ByVal iAno As Integer, ByVal sProduto As String, ByVal iMes As Integer, dCustoMedio As Double) As Long

Dim tSldMesEst1 As typeSldMesEst1
Dim lErro As Long
Dim iIndice As Integer
Dim dQuantAtual As Double
Dim dValorAtual As Double

On Error GoTo Erro_SldMesEst1_Le_CustoMedioConserto3

    'seleciona o registro de SldMesEst1 relativo ao Ano, FilialEmpresa e Produto passados como parametro
    lErro = Comando_Executar(lComando, "SELECT QuantInicialConserto3, ValorInicialConserto3, SaldoQuantConserto31, SaldoValorConserto31, SaldoQuantConserto32, SaldoValorConserto32, SaldoQuantConserto33, SaldoValorConserto33, SaldoQuantConserto34, SaldoValorConserto34, SaldoQuantConserto35, SaldoValorConserto35, SaldoQuantConserto36, SaldoValorConserto36, SaldoQuantConserto37, SaldoValorConserto37, SaldoQuantConserto38, SaldoValorConserto38, SaldoQuantConserto39, SaldoValorConserto39, SaldoQuantConserto310, SaldoValorConserto310, SaldoQuantConserto311, SaldoValorConserto311, SaldoQuantConserto312, SaldoValorConserto312  FROM SldMesEst1 WHERE Ano=? AND FilialEmpresa=? AND Produto=?", _
                              tSldMesEst1.dQuantInicialConserto3, tSldMesEst1.dValorInicialConserto3, tSldMesEst1.adSaldoQuantConserto3(1), tSldMesEst1.adSaldoValorConserto3(1), tSldMesEst1.adSaldoQuantConserto3(2), tSldMesEst1.adSaldoValorConserto3(2), tSldMesEst1.adSaldoQuantConserto3(3), tSldMesEst1.adSaldoValorConserto3(3), tSldMesEst1.adSaldoQuantConserto3(4), tSldMesEst1.adSaldoValorConserto3(4), tSldMesEst1.adSaldoQuantConserto3(5), tSldMesEst1.adSaldoValorConserto3(5), tSldMesEst1.adSaldoQuantConserto3(6), tSldMesEst1.adSaldoValorConserto3(6), tSldMesEst1.adSaldoQuantConserto3(7), tSldMesEst1.adSaldoValorConserto3(7), tSldMesEst1.adSaldoQuantConserto3(8), tSldMesEst1.adSaldoValorConserto3(8), tSldMesEst1.adSaldoQuantConserto3(9), tSldMesEst1.adSaldoValorConserto3(9), _
                              tSldMesEst1.adSaldoQuantConserto3(10), tSldMesEst1.adSaldoValorConserto3(10), tSldMesEst1.adSaldoQuantConserto3(11), tSldMesEst1.adSaldoValorConserto3(11), tSldMesEst1.adSaldoQuantConserto3(12), tSldMesEst1.adSaldoValorConserto3(12), iAno, iFilialEmpresa, sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 71508
    
    'le o SldMesEst
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 71509
    
    dQuantAtual = tSldMesEst1.dQuantInicialConserto3
    dValorAtual = tSldMesEst1.dValorInicialConserto3

    For iIndice = 1 To iMes
    
        dQuantAtual = dQuantAtual + tSldMesEst1.adSaldoQuantConserto3(iIndice)
        dValorAtual = dValorAtual + tSldMesEst1.adSaldoValorConserto3(iIndice)
        
    Next
    
    If Abs(dQuantAtual) > QTDE_ESTOQUE_DELTA Then

        dCustoMedio = dValorAtual / dQuantAtual
        
    Else
    
        dCustoMedio = 0
        
    End If
    
    
    SldMesEst1_Le_CustoMedioConserto3 = SUCESSO

    Exit Function

Erro_SldMesEst1_Le_CustoMedioConserto3:
    
    SldMesEst1_Le_CustoMedioConserto3 = gErr
    
    Select Case gErr
    
        Case 71508, 71509
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESEST11", gErr, iAno, iFilialEmpresa, sProduto)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153543)
        
    End Select
    
    Exit Function

End Function

Function SldMesEst1_Le_CustoMedioDemo3(ByVal lComando As Long, ByVal iFilialEmpresa As Integer, ByVal iAno As Integer, ByVal sProduto As String, ByVal iMes As Integer, dCustoMedio As Double) As Long

Dim tSldMesEst1 As typeSldMesEst1
Dim lErro As Long
Dim iIndice As Integer
Dim dQuantAtual As Double
Dim dValorAtual As Double

On Error GoTo Erro_SldMesEst1_Le_CustoMedioDemo3

    'seleciona o registro de SldMesEst1 relativo ao Ano, FilialEmpresa e Produto passados como parametro
    lErro = Comando_Executar(lComando, "SELECT QuantInicialDemo3, ValorInicialDemo3, SaldoQuantDemo31, SaldoValorDemo31, SaldoQuantDemo32, SaldoValorDemo32, SaldoQuantDemo33, SaldoValorDemo33, SaldoQuantDemo34, SaldoValorDemo34, SaldoQuantDemo35, SaldoValorDemo35, SaldoQuantDemo36, SaldoValorDemo36, SaldoQuantDemo37, SaldoValorDemo37, SaldoQuantDemo38, SaldoValorDemo38, SaldoQuantDemo39, SaldoValorDemo39, SaldoQuantDemo310, SaldoValorDemo310, SaldoQuantDemo311, SaldoValorDemo311, SaldoQuantDemo312, SaldoValorDemo312  FROM SldMesEst1 WHERE Ano=? AND FilialEmpresa=? AND Produto=?", _
                              tSldMesEst1.dQuantInicialDemo3, tSldMesEst1.dValorInicialDemo3, tSldMesEst1.adSaldoQuantDemo3(1), tSldMesEst1.adSaldoValorDemo3(1), tSldMesEst1.adSaldoQuantDemo3(2), tSldMesEst1.adSaldoValorDemo3(2), tSldMesEst1.adSaldoQuantDemo3(3), tSldMesEst1.adSaldoValorDemo3(3), tSldMesEst1.adSaldoQuantDemo3(4), tSldMesEst1.adSaldoValorDemo3(4), tSldMesEst1.adSaldoQuantDemo3(5), tSldMesEst1.adSaldoValorDemo3(5), tSldMesEst1.adSaldoQuantDemo3(6), tSldMesEst1.adSaldoValorDemo3(6), tSldMesEst1.adSaldoQuantDemo3(7), tSldMesEst1.adSaldoValorDemo3(7), tSldMesEst1.adSaldoQuantDemo3(8), tSldMesEst1.adSaldoValorDemo3(8), tSldMesEst1.adSaldoQuantDemo3(9), tSldMesEst1.adSaldoValorDemo3(9), _
                              tSldMesEst1.adSaldoQuantDemo3(10), tSldMesEst1.adSaldoValorDemo3(10), tSldMesEst1.adSaldoQuantDemo3(11), tSldMesEst1.adSaldoValorDemo3(11), tSldMesEst1.adSaldoQuantDemo3(12), tSldMesEst1.adSaldoValorDemo3(12), iAno, iFilialEmpresa, sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 71510
    
    'le o SldMesEst
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 71511
    
    dQuantAtual = tSldMesEst1.dQuantInicialDemo3
    dValorAtual = tSldMesEst1.dValorInicialDemo3

    For iIndice = 1 To iMes
    
        dQuantAtual = dQuantAtual + tSldMesEst1.adSaldoQuantDemo3(iIndice)
        dValorAtual = dValorAtual + tSldMesEst1.adSaldoValorDemo3(iIndice)
        
    Next
    
    If Abs(dQuantAtual) > QTDE_ESTOQUE_DELTA Then

        dCustoMedio = dValorAtual / dQuantAtual
        
    Else
    
        dCustoMedio = 0
        
    End If
    
    SldMesEst1_Le_CustoMedioDemo3 = SUCESSO

    Exit Function

Erro_SldMesEst1_Le_CustoMedioDemo3:
    
    SldMesEst1_Le_CustoMedioDemo3 = gErr
    
    Select Case gErr
    
        Case 71510, 71511
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESEST11", gErr, iAno, iFilialEmpresa, sProduto)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153544)
        
    End Select
    
    Exit Function

End Function

Function SldMesEst1_Le_CustoMedioBenef3(ByVal lComando As Long, ByVal iFilialEmpresa As Integer, ByVal iAno As Integer, ByVal sProduto As String, ByVal iMes As Integer, dCustoMedio As Double) As Long

Dim tSldMesEst1 As typeSldMesEst1
Dim lErro As Long
Dim iIndice As Integer
Dim dQuantAtual As Double
Dim dValorAtual As Double

On Error GoTo Erro_SldMesEst1_Le_CustoMedioBenef3

    'seleciona o registro de SldMesEst1 relativo ao Ano, FilialEmpresa e Produto passados como parametro
    lErro = Comando_Executar(lComando, "SELECT QuantInicialBenef3, ValorInicialBenef3, SaldoQuantBenef31, SaldoValorBenef31, SaldoQuantBenef32, SaldoValorBenef32, SaldoQuantBenef33, SaldoValorBenef33, SaldoQuantBenef34, SaldoValorBenef34, SaldoQuantBenef35, SaldoValorBenef35, SaldoQuantBenef36, SaldoValorBenef36, SaldoQuantBenef37, SaldoValorBenef37, SaldoQuantBenef38, SaldoValorBenef38, SaldoQuantBenef39, SaldoValorBenef39, SaldoQuantBenef310, SaldoValorBenef310, SaldoQuantBenef311, SaldoValorBenef311, SaldoQuantBenef312, SaldoValorBenef312  FROM SldMesEst1 WHERE Ano=? AND FilialEmpresa=? AND Produto=?", _
                              tSldMesEst1.dQuantInicialBenef3, tSldMesEst1.dValorInicialBenef3, tSldMesEst1.adSaldoQuantBenef3(1), tSldMesEst1.adSaldoValorBenef3(1), tSldMesEst1.adSaldoQuantBenef3(2), tSldMesEst1.adSaldoValorBenef3(2), tSldMesEst1.adSaldoQuantBenef3(3), tSldMesEst1.adSaldoValorBenef3(3), tSldMesEst1.adSaldoQuantBenef3(4), tSldMesEst1.adSaldoValorBenef3(4), tSldMesEst1.adSaldoQuantBenef3(5), tSldMesEst1.adSaldoValorBenef3(5), tSldMesEst1.adSaldoQuantBenef3(6), tSldMesEst1.adSaldoValorBenef3(6), tSldMesEst1.adSaldoQuantBenef3(7), tSldMesEst1.adSaldoValorBenef3(7), tSldMesEst1.adSaldoQuantBenef3(8), tSldMesEst1.adSaldoValorBenef3(8), tSldMesEst1.adSaldoQuantBenef3(9), tSldMesEst1.adSaldoValorBenef3(9), _
                              tSldMesEst1.adSaldoQuantBenef3(10), tSldMesEst1.adSaldoValorBenef3(10), tSldMesEst1.adSaldoQuantBenef3(11), tSldMesEst1.adSaldoValorBenef3(11), tSldMesEst1.adSaldoQuantBenef3(12), tSldMesEst1.adSaldoValorBenef3(12), iAno, iFilialEmpresa, sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 71514
    
    'le o SldMesEst
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 71515
    
    dQuantAtual = tSldMesEst1.dQuantInicialBenef3
    dValorAtual = tSldMesEst1.dValorInicialBenef3

    For iIndice = 1 To iMes
    
        dQuantAtual = dQuantAtual + tSldMesEst1.adSaldoQuantBenef3(iIndice)
        dValorAtual = dValorAtual + tSldMesEst1.adSaldoValorBenef3(iIndice)
        
    Next
    
    If Abs(dQuantAtual) > QTDE_ESTOQUE_DELTA Then

        dCustoMedio = dValorAtual / dQuantAtual
        
    Else
    
        dCustoMedio = 0
        
    End If
    
    SldMesEst1_Le_CustoMedioBenef3 = SUCESSO

    Exit Function

Erro_SldMesEst1_Le_CustoMedioBenef3:
    
    SldMesEst1_Le_CustoMedioBenef3 = gErr
    
    Select Case gErr
    
        Case 71514, 71515
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESEST11", gErr, iAno, iFilialEmpresa, sProduto)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153545)
        
    End Select
    
    Exit Function

End Function

Function SldMesEst1_Le_CustoMedioOutros3(ByVal lComando As Long, ByVal iFilialEmpresa As Integer, ByVal iAno As Integer, ByVal sProduto As String, ByVal iMes As Integer, dCustoMedio As Double) As Long

Dim tSldMesEst1 As typeSldMesEst1
Dim lErro As Long
Dim iIndice As Integer
Dim dQuantAtual As Double
Dim dValorAtual As Double

On Error GoTo Erro_SldMesEst1_Le_CustoMedioOutros3

    'seleciona o registro de SldMesEst1 relativo ao Ano, FilialEmpresa e Produto passados como parametro
    lErro = Comando_Executar(lComando, "SELECT QuantInicialOutros3, ValorInicialOutros3, SaldoQuantOutros31, SaldoValorOutros31, SaldoQuantOutros32, SaldoValorOutros32, SaldoQuantOutros33, SaldoValorOutros33, SaldoQuantOutros34, SaldoValorOutros34, SaldoQuantOutros35, SaldoValorOutros35, SaldoQuantOutros36, SaldoValorOutros36, SaldoQuantOutros37, SaldoValorOutros37, SaldoQuantOutros38, SaldoValorOutros38, SaldoQuantOutros39, SaldoValorOutros39, SaldoQuantOutros310, SaldoValorOutros310, SaldoQuantOutros311, SaldoValorOutros311, SaldoQuantOutros312, SaldoValorOutros312  FROM SldMesEst1 WHERE Ano=? AND FilialEmpresa=? AND Produto=?", _
                              tSldMesEst1.dQuantInicialOutros3, tSldMesEst1.dValorInicialOutros3, tSldMesEst1.adSaldoQuantOutros3(1), tSldMesEst1.adSaldoValorOutros3(1), tSldMesEst1.adSaldoQuantOutros3(2), tSldMesEst1.adSaldoValorOutros3(2), tSldMesEst1.adSaldoQuantOutros3(3), tSldMesEst1.adSaldoValorOutros3(3), tSldMesEst1.adSaldoQuantOutros3(4), tSldMesEst1.adSaldoValorOutros3(4), tSldMesEst1.adSaldoQuantOutros3(5), tSldMesEst1.adSaldoValorOutros3(5), tSldMesEst1.adSaldoQuantOutros3(6), tSldMesEst1.adSaldoValorOutros3(6), tSldMesEst1.adSaldoQuantOutros3(7), tSldMesEst1.adSaldoValorOutros3(7), tSldMesEst1.adSaldoQuantOutros3(8), tSldMesEst1.adSaldoValorOutros3(8), tSldMesEst1.adSaldoQuantOutros3(9), tSldMesEst1.adSaldoValorOutros3(9), _
                              tSldMesEst1.adSaldoQuantOutros3(10), tSldMesEst1.adSaldoValorOutros3(10), tSldMesEst1.adSaldoQuantOutros3(11), tSldMesEst1.adSaldoValorOutros3(11), tSldMesEst1.adSaldoQuantOutros3(12), tSldMesEst1.adSaldoValorOutros3(12), iAno, iFilialEmpresa, sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 71512
    
    'le o SldMesEst
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 71513
    
    dQuantAtual = tSldMesEst1.dQuantInicialOutros3
    dValorAtual = tSldMesEst1.dValorInicialOutros3

    For iIndice = 1 To iMes
    
        dQuantAtual = dQuantAtual + tSldMesEst1.adSaldoQuantOutros3(iIndice)
        dValorAtual = dValorAtual + tSldMesEst1.adSaldoValorOutros3(iIndice)
        
    Next
    
    If Abs(dQuantAtual) > QTDE_ESTOQUE_DELTA Then

        dCustoMedio = dValorAtual / dQuantAtual
        
    Else
    
        dCustoMedio = 0
        
    End If
    
    SldMesEst1_Le_CustoMedioOutros3 = SUCESSO

    Exit Function

Erro_SldMesEst1_Le_CustoMedioOutros3:
    
    SldMesEst1_Le_CustoMedioOutros3 = gErr
    
    Select Case gErr
    
        Case 71512, 71513
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESEST11", gErr, iAno, iFilialEmpresa, sProduto)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153546)
        
    End Select
    
    Exit Function

End Function

Function SldMesEst2_Le_CustoMedioConserto(ByVal lComando As Long, ByVal iFilialEmpresa As Integer, ByVal iAno As Integer, ByVal sProduto As String, ByVal iMes As Integer, dCustoMedio As Double) As Long
'calcula o custo médio de material em conserto nosso em poder de terceiros.

Dim tSldMesEst2 As typeSldMesEst2
Dim lErro As Long
Dim iIndice As Integer
Dim dQuantAtual As Double
Dim dValorAtual As Double

On Error GoTo Erro_SldMesEst2_Le_CustoMedioConserto

    'seleciona o registro de SldMesEst2 relativo ao Ano, FilialEmpresa e Produto passados como parametro
    lErro = Comando_Executar(lComando, "SELECT QuantInicialConserto, ValorInicialConserto, SaldoQuantConserto1, SaldoValorConserto1, SaldoQuantConserto2, SaldoValorConserto2, SaldoQuantConserto3, SaldoValorConserto3, SaldoQuantConserto4, SaldoValorConserto4, SaldoQuantConserto5, SaldoValorConserto5, SaldoQuantConserto6, SaldoValorConserto6, SaldoQuantConserto7, SaldoValorConserto7, SaldoQuantConserto8, SaldoValorConserto8, SaldoQuantConserto9, SaldoValorConserto9, SaldoQuantConserto10, SaldoValorConserto10, SaldoQuantConserto11, SaldoValorConserto11, SaldoQuantConserto12, SaldoValorConserto12  FROM SldMesEst2 WHERE Ano=? AND FilialEmpresa=? AND Produto=?", _
                              tSldMesEst2.dQuantInicialConserto, tSldMesEst2.dValorInicialConserto, tSldMesEst2.adSaldoQuantConserto(1), tSldMesEst2.adSaldoValorConserto(1), tSldMesEst2.adSaldoQuantConserto(2), tSldMesEst2.adSaldoValorConserto(2), tSldMesEst2.adSaldoQuantConserto(3), tSldMesEst2.adSaldoValorConserto(3), tSldMesEst2.adSaldoQuantConserto(4), tSldMesEst2.adSaldoValorConserto(4), tSldMesEst2.adSaldoQuantConserto(5), tSldMesEst2.adSaldoValorConserto(5), tSldMesEst2.adSaldoQuantConserto(6), tSldMesEst2.adSaldoValorConserto(6), tSldMesEst2.adSaldoQuantConserto(7), tSldMesEst2.adSaldoValorConserto(7), tSldMesEst2.adSaldoQuantConserto(8), tSldMesEst2.adSaldoValorConserto(8), tSldMesEst2.adSaldoQuantConserto(9), tSldMesEst2.adSaldoValorConserto(9), _
                              tSldMesEst2.adSaldoQuantConserto(10), tSldMesEst2.adSaldoValorConserto(10), tSldMesEst2.adSaldoQuantConserto(11), tSldMesEst2.adSaldoValorConserto(11), tSldMesEst2.adSaldoQuantConserto(12), tSldMesEst2.adSaldoValorConserto(12), iAno, iFilialEmpresa, sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 60967
    
    'le o SldMesEst
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 60968
    
    dQuantAtual = tSldMesEst2.dQuantInicialConserto
    dValorAtual = tSldMesEst2.dValorInicialConserto

    For iIndice = 1 To iMes
    
        dQuantAtual = dQuantAtual + tSldMesEst2.adSaldoQuantConserto(iIndice)
        dValorAtual = dValorAtual + tSldMesEst2.adSaldoValorConserto(iIndice)
        
    Next
    
    If Abs(dQuantAtual) > QTDE_ESTOQUE_DELTA Then

        dCustoMedio = dValorAtual / dQuantAtual
        
    Else
    
        dCustoMedio = 0
        
    End If
    
    
    SldMesEst2_Le_CustoMedioConserto = SUCESSO

    Exit Function

Erro_SldMesEst2_Le_CustoMedioConserto:
    
    SldMesEst2_Le_CustoMedioConserto = gErr
    
    Select Case gErr
    
        Case 60967, 60968
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESEST21", gErr, iAno, iFilialEmpresa, sProduto)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153547)
        
    End Select
    
    Exit Function

End Function

Function SldMesEst1_Le_CustoMedioConsig3(ByVal lComando As Long, ByVal iFilialEmpresa As Integer, ByVal iAno As Integer, ByVal sProduto As String, ByVal iMes As Integer, dCustoMedio As Double) As Long

Dim tSldMesEst1 As typeSldMesEst1
Dim lErro As Long
Dim iIndice As Integer
Dim dQuantAtual As Double
Dim dValorAtual As Double

On Error GoTo Erro_SldMesEst1_Le_CustoMedioConsig3

    'seleciona o registro de SldMesEst1 relativo ao Ano, FilialEmpresa e Produto passados como parametro
    lErro = Comando_Executar(lComando, "SELECT QuantInicialConsig3, ValorInicialConsig3, SaldoQuantConsig31, SaldoValorConsig31, SaldoQuantConsig32, SaldoValorConsig32, SaldoQuantConsig33, SaldoValorConsig33, SaldoQuantConsig34, SaldoValorConsig34, SaldoQuantConsig35, SaldoValorConsig35, SaldoQuantConsig36, SaldoValorConsig36, SaldoQuantConsig37, SaldoValorConsig37, SaldoQuantConsig38, SaldoValorConsig38, SaldoQuantConsig39, SaldoValorConsig39, SaldoQuantConsig310, SaldoValorConsig310, SaldoQuantConsig311, SaldoValorConsig311, SaldoQuantConsig312, SaldoValorConsig312  FROM SldMesEst1 WHERE Ano=? AND FilialEmpresa=? AND Produto=?", _
                              tSldMesEst1.dQuantInicialConsig3, tSldMesEst1.dValorInicialConsig3, tSldMesEst1.adSaldoQuantConsig3(1), tSldMesEst1.adSaldoValorConsig3(1), tSldMesEst1.adSaldoQuantConsig3(2), tSldMesEst1.adSaldoValorConsig3(2), tSldMesEst1.adSaldoQuantConsig3(3), tSldMesEst1.adSaldoValorConsig3(3), tSldMesEst1.adSaldoQuantConsig3(4), tSldMesEst1.adSaldoValorConsig3(4), tSldMesEst1.adSaldoQuantConsig3(5), tSldMesEst1.adSaldoValorConsig3(5), tSldMesEst1.adSaldoQuantConsig3(6), tSldMesEst1.adSaldoValorConsig3(6), tSldMesEst1.adSaldoQuantConsig3(7), tSldMesEst1.adSaldoValorConsig3(7), tSldMesEst1.adSaldoQuantConsig3(8), tSldMesEst1.adSaldoValorConsig3(8), tSldMesEst1.adSaldoQuantConsig3(9), tSldMesEst1.adSaldoValorConsig3(9), _
                              tSldMesEst1.adSaldoQuantConsig3(10), tSldMesEst1.adSaldoValorConsig3(10), tSldMesEst1.adSaldoQuantConsig3(11), tSldMesEst1.adSaldoValorConsig3(11), tSldMesEst1.adSaldoQuantConsig3(12), tSldMesEst1.adSaldoValorConsig3(12), iAno, iFilialEmpresa, sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 60922

    'le o SldMesEst
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 60923
    
    dQuantAtual = tSldMesEst1.dQuantInicialConsig3
    dValorAtual = tSldMesEst1.dValorInicialConsig3

    For iIndice = 1 To iMes
    
        dQuantAtual = dQuantAtual + tSldMesEst1.adSaldoQuantConsig3(iIndice)
        dValorAtual = dValorAtual + tSldMesEst1.adSaldoValorConsig3(iIndice)
        
    Next
    
    If Abs(dQuantAtual) > QTDE_ESTOQUE_DELTA Then

        dCustoMedio = dValorAtual / dQuantAtual
        
    Else
    
        dCustoMedio = 0
        
    End If
    
    SldMesEst1_Le_CustoMedioConsig3 = SUCESSO

    Exit Function

Erro_SldMesEst1_Le_CustoMedioConsig3:
    
    SldMesEst1_Le_CustoMedioConsig3 = gErr
    
    Select Case gErr
    
        Case 60922, 60923
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESEST11", gErr, iAno, iFilialEmpresa, sProduto)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153548)
        
    End Select
    
    Exit Function

End Function

Function SldMesEst2_Le_CustoMedioDemo(ByVal lComando As Long, ByVal iFilialEmpresa As Integer, ByVal iAno As Integer, ByVal sProduto As String, ByVal iMes As Integer, dCustoMedio As Double) As Long
'calcula o custo médio de material em Demonstração nosso em poder de terceiros.

Dim tSldMesEst2 As typeSldMesEst2
Dim lErro As Long
Dim iIndice As Integer
Dim dQuantAtual As Double
Dim dValorAtual As Double

On Error GoTo Erro_SldMesEst2_Le_CustoMedioDemo

    'seleciona o registro de SldMesEst2 relativo ao Ano, FilialEmpresa e Produto passados como parametro
    lErro = Comando_Executar(lComando, "SELECT QuantInicialDemo, ValorInicialDemo, SaldoQuantDemo1, SaldoValorDemo1, SaldoQuantDemo2, SaldoValorDemo2, SaldoQuantDemo3, SaldoValorDemo3, SaldoQuantDemo4, SaldoValorDemo4, SaldoQuantDemo5, SaldoValorDemo5, SaldoQuantDemo6, SaldoValorDemo6, SaldoQuantDemo7, SaldoValorDemo7, SaldoQuantDemo8, SaldoValorDemo8, SaldoQuantDemo9, SaldoValorDemo9, SaldoQuantDemo10, SaldoValorDemo10, SaldoQuantDemo11, SaldoValorDemo11, SaldoQuantDemo12, SaldoValorDemo12  FROM SldMesEst2 WHERE Ano=? AND FilialEmpresa=? AND Produto=?", _
                              tSldMesEst2.dQuantInicialDemo, tSldMesEst2.dValorInicialDemo, tSldMesEst2.adSaldoQuantDemo(1), tSldMesEst2.adSaldoValorDemo(1), tSldMesEst2.adSaldoQuantDemo(2), tSldMesEst2.adSaldoValorDemo(2), tSldMesEst2.adSaldoQuantDemo(3), tSldMesEst2.adSaldoValorDemo(3), tSldMesEst2.adSaldoQuantDemo(4), tSldMesEst2.adSaldoValorDemo(4), tSldMesEst2.adSaldoQuantDemo(5), tSldMesEst2.adSaldoValorDemo(5), tSldMesEst2.adSaldoQuantDemo(6), tSldMesEst2.adSaldoValorDemo(6), tSldMesEst2.adSaldoQuantDemo(7), tSldMesEst2.adSaldoValorDemo(7), tSldMesEst2.adSaldoQuantDemo(8), tSldMesEst2.adSaldoValorDemo(8), tSldMesEst2.adSaldoQuantDemo(9), tSldMesEst2.adSaldoValorDemo(9), _
                              tSldMesEst2.adSaldoQuantDemo(10), tSldMesEst2.adSaldoValorDemo(10), tSldMesEst2.adSaldoQuantDemo(11), tSldMesEst2.adSaldoValorDemo(11), tSldMesEst2.adSaldoQuantDemo(12), tSldMesEst2.adSaldoValorDemo(12), iAno, iFilialEmpresa, sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 60970
    
    'le o SldMesEst
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 60971
    
    dQuantAtual = tSldMesEst2.dQuantInicialDemo
    dValorAtual = tSldMesEst2.dValorInicialDemo

    For iIndice = 1 To iMes
    
        dQuantAtual = dQuantAtual + tSldMesEst2.adSaldoQuantDemo(iIndice)
        dValorAtual = dValorAtual + tSldMesEst2.adSaldoValorDemo(iIndice)
        
    Next
    
    If Abs(dQuantAtual) > QTDE_ESTOQUE_DELTA Then

        dCustoMedio = dValorAtual / dQuantAtual
        
    Else
    
        dCustoMedio = 0
        
    End If
    
    SldMesEst2_Le_CustoMedioDemo = SUCESSO

    Exit Function

Erro_SldMesEst2_Le_CustoMedioDemo:
    
    SldMesEst2_Le_CustoMedioDemo = gErr
    
    Select Case gErr
    
        Case 60970, 60971
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESEST21", gErr, iAno, iFilialEmpresa, sProduto)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153549)
        
    End Select
    
    Exit Function

End Function

Function SldMesEst2_Le_CustoMedioOutros(ByVal lComando As Long, ByVal iFilialEmpresa As Integer, ByVal iAno As Integer, ByVal sProduto As String, ByVal iMes As Integer, dCustoMedio As Double) As Long
'calcula o custo médio de Outros materiais nossos em poder de terceiros.

Dim tSldMesEst2 As typeSldMesEst2
Dim lErro As Long
Dim iIndice As Integer
Dim dQuantAtual As Double
Dim dValorAtual As Double

On Error GoTo Erro_SldMesEst2_Le_CustoMedioOutros

    'seleciona o registro de SldMesEst2 relativo ao Ano, FilialEmpresa e Produto passados como parametro
    lErro = Comando_Executar(lComando, "SELECT QuantInicialOutros, ValorInicialOutros, SaldoQuantOutros1, SaldoValorOutros1, SaldoQuantOutros2, SaldoValorOutros2, SaldoQuantOutros3, SaldoValorOutros3, SaldoQuantOutros4, SaldoValorOutros4, SaldoQuantOutros5, SaldoValorOutros5, SaldoQuantOutros6, SaldoValorOutros6, SaldoQuantOutros7, SaldoValorOutros7, SaldoQuantOutros8, SaldoValorOutros8, SaldoQuantOutros9, SaldoValorOutros9, SaldoQuantOutros10, SaldoValorOutros10, SaldoQuantOutros11, SaldoValorOutros11, SaldoQuantOutros12, SaldoValorOutros12  FROM SldMesEst2 WHERE Ano=? AND FilialEmpresa=? AND Produto=?", _
                              tSldMesEst2.dQuantInicialOutros, tSldMesEst2.dValorInicialOutros, tSldMesEst2.adSaldoQuantOutros(1), tSldMesEst2.adSaldoValorOutros(1), tSldMesEst2.adSaldoQuantOutros(2), tSldMesEst2.adSaldoValorOutros(2), tSldMesEst2.adSaldoQuantOutros(3), tSldMesEst2.adSaldoValorOutros(3), tSldMesEst2.adSaldoQuantOutros(4), tSldMesEst2.adSaldoValorOutros(4), tSldMesEst2.adSaldoQuantOutros(5), tSldMesEst2.adSaldoValorOutros(5), tSldMesEst2.adSaldoQuantOutros(6), tSldMesEst2.adSaldoValorOutros(6), tSldMesEst2.adSaldoQuantOutros(7), tSldMesEst2.adSaldoValorOutros(7), tSldMesEst2.adSaldoQuantOutros(8), tSldMesEst2.adSaldoValorOutros(8), tSldMesEst2.adSaldoQuantOutros(9), tSldMesEst2.adSaldoValorOutros(9), _
                              tSldMesEst2.adSaldoQuantOutros(10), tSldMesEst2.adSaldoValorOutros(10), tSldMesEst2.adSaldoQuantOutros(11), tSldMesEst2.adSaldoValorOutros(11), tSldMesEst2.adSaldoQuantOutros(12), tSldMesEst2.adSaldoValorOutros(12), iAno, iFilialEmpresa, sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 60973

    'le o SldMesEst
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 60974
    
    dQuantAtual = tSldMesEst2.dQuantInicialOutros
    dValorAtual = tSldMesEst2.dValorInicialOutros

    For iIndice = 1 To iMes
    
        dQuantAtual = dQuantAtual + tSldMesEst2.adSaldoQuantOutros(iIndice)
        dValorAtual = dValorAtual + tSldMesEst2.adSaldoValorOutros(iIndice)
        
    Next
    
    If Abs(dQuantAtual) > QTDE_ESTOQUE_DELTA Then

        dCustoMedio = dValorAtual / dQuantAtual
        
    Else
    
        dCustoMedio = 0
        
    End If
    
    SldMesEst2_Le_CustoMedioOutros = SUCESSO

    Exit Function

Erro_SldMesEst2_Le_CustoMedioOutros:
    
    SldMesEst2_Le_CustoMedioOutros = gErr
    
    Select Case gErr
    
        Case 60973, 60974
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESEST21", gErr, iAno, iFilialEmpresa, sProduto)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153550)
        
    End Select
    
    Exit Function

End Function

Function SldMesEst2_Le_CustoMedioBenef(ByVal lComando As Long, ByVal iFilialEmpresa As Integer, ByVal iAno As Integer, ByVal sProduto As String, ByVal iMes As Integer, dCustoMedio As Double) As Long
'calcula o custo médio de materiais em beneficiamento nossos em poder de terceiros.

Dim tSldMesEst2 As typeSldMesEst2
Dim lErro As Long
Dim iIndice As Integer
Dim dQuantAtual As Double
Dim dValorAtual As Double

On Error GoTo Erro_SldMesEst2_Le_CustoMedioBenef

    'seleciona o registro de SldMesEst2 relativo ao Ano, FilialEmpresa e Produto passados como parametro
    lErro = Comando_Executar(lComando, "SELECT QuantInicialBenef, ValorInicialBenef, SaldoQuantBenef1, SaldoValorBenef1, SaldoQuantBenef2, SaldoValorBenef2, SaldoQuantBenef3, SaldoValorBenef3, SaldoQuantBenef4, SaldoValorBenef4, SaldoQuantBenef5, SaldoValorBenef5, SaldoQuantBenef6, SaldoValorBenef6, SaldoQuantBenef7, SaldoValorBenef7, SaldoQuantBenef8, SaldoValorBenef8, SaldoQuantBenef9, SaldoValorBenef9, SaldoQuantBenef10, SaldoValorBenef10, SaldoQuantBenef11, SaldoValorBenef11, SaldoQuantBenef12, SaldoValorBenef12  FROM SldMesEst2 WHERE Ano=? AND FilialEmpresa=? AND Produto=?", _
                              tSldMesEst2.dQuantInicialBenef, tSldMesEst2.dValorInicialBenef, tSldMesEst2.adSaldoQuantBenef(1), tSldMesEst2.adSaldoValorBenef(1), tSldMesEst2.adSaldoQuantBenef(2), tSldMesEst2.adSaldoValorBenef(2), tSldMesEst2.adSaldoQuantBenef(3), tSldMesEst2.adSaldoValorBenef(3), tSldMesEst2.adSaldoQuantBenef(4), tSldMesEst2.adSaldoValorBenef(4), tSldMesEst2.adSaldoQuantBenef(5), tSldMesEst2.adSaldoValorBenef(5), tSldMesEst2.adSaldoQuantBenef(6), tSldMesEst2.adSaldoValorBenef(6), tSldMesEst2.adSaldoQuantBenef(7), tSldMesEst2.adSaldoValorBenef(7), tSldMesEst2.adSaldoQuantBenef(8), tSldMesEst2.adSaldoValorBenef(8), tSldMesEst2.adSaldoQuantBenef(9), tSldMesEst2.adSaldoValorBenef(9), _
                              tSldMesEst2.adSaldoQuantBenef(10), tSldMesEst2.adSaldoValorBenef(10), tSldMesEst2.adSaldoQuantBenef(11), tSldMesEst2.adSaldoValorBenef(11), tSldMesEst2.adSaldoQuantBenef(12), tSldMesEst2.adSaldoValorBenef(12), iAno, iFilialEmpresa, sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 60976


    'le o SldMesEst
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 60977
    
    dQuantAtual = tSldMesEst2.dQuantInicialBenef
    dValorAtual = tSldMesEst2.dValorInicialBenef

    For iIndice = 1 To iMes
    
        dQuantAtual = dQuantAtual + tSldMesEst2.adSaldoQuantBenef(iIndice)
        dValorAtual = dValorAtual + tSldMesEst2.adSaldoValorBenef(iIndice)
        
    Next
    
    If Abs(dQuantAtual) > QTDE_ESTOQUE_DELTA Then
    
        dCustoMedio = dValorAtual / dQuantAtual
            
    Else
        
        dCustoMedio = 0
            
    End If
    
    
    
    SldMesEst2_Le_CustoMedioBenef = SUCESSO

    Exit Function

Erro_SldMesEst2_Le_CustoMedioBenef:
    
    SldMesEst2_Le_CustoMedioBenef = gErr
    
    Select Case gErr
    
        Case 60976, 60977
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESEST21", gErr, iAno, iFilialEmpresa, sProduto)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153551)
        
    End Select
    
    Exit Function

End Function

Function SldMesEst2_Le_CustoMedioConsig(ByVal lComando As Long, ByVal iFilialEmpresa As Integer, ByVal iAno As Integer, ByVal sProduto As String, ByVal iMes As Integer, dCustoMedio As Double) As Long
'calcula o custo médio de materiais em consignação nossos em poder de terceiros.

Dim tSldMesEst2 As typeSldMesEst2
Dim lErro As Long
Dim iIndice As Integer
Dim dQuantAtual As Double
Dim dValorAtual As Double
Dim sMes As String

On Error GoTo Erro_SldMesEst2_Le_CustoMedioConsig
    
    sMes = CStr(iMes)

    'seleciona o registro de SldMesEst2 relativo ao Ano, FilialEmpresa e Produto passados como parametro
    lErro = Comando_Executar(lComando, "SELECT QuantInicialConsig, ValorInicialConsig, SaldoQuantConsig1, SaldoValorConsig1, SaldoQuantConsig2, SaldoValorConsig2, SaldoQuantConsig3, SaldoValorConsig3, SaldoQuantConsig4, SaldoValorConsig4, SaldoQuantConsig5, SaldoValorConsig5, SaldoQuantConsig6, SaldoValorConsig6, SaldoQuantConsig7, SaldoValorConsig7, SaldoQuantConsig8, SaldoValorConsig8, SaldoQuantConsig9, SaldoValorConsig9, SaldoQuantConsig10, SaldoValorConsig10, SaldoQuantConsig11, SaldoValorConsig11, SaldoQuantConsig12, SaldoValorConsig12  FROM SldMesEst2 WHERE Ano=? AND FilialEmpresa=? AND Produto=?", _
                              tSldMesEst2.dQuantInicialConsig, tSldMesEst2.dValorInicialConsig, tSldMesEst2.adSaldoQuantConsig(1), tSldMesEst2.adSaldoValorConsig(1), tSldMesEst2.adSaldoQuantConsig(2), tSldMesEst2.adSaldoValorConsig(2), tSldMesEst2.adSaldoQuantConsig(3), tSldMesEst2.adSaldoValorConsig(3), tSldMesEst2.adSaldoQuantConsig(4), tSldMesEst2.adSaldoValorConsig(4), tSldMesEst2.adSaldoQuantConsig(5), tSldMesEst2.adSaldoValorConsig(5), tSldMesEst2.adSaldoQuantConsig(6), tSldMesEst2.adSaldoValorConsig(6), tSldMesEst2.adSaldoQuantConsig(7), tSldMesEst2.adSaldoValorConsig(7), tSldMesEst2.adSaldoQuantConsig(8), tSldMesEst2.adSaldoValorConsig(8), tSldMesEst2.adSaldoQuantConsig(9), tSldMesEst2.adSaldoValorConsig(9), _
                              tSldMesEst2.adSaldoQuantConsig(10), tSldMesEst2.adSaldoValorConsig(10), tSldMesEst2.adSaldoQuantConsig(11), tSldMesEst2.adSaldoValorConsig(11), tSldMesEst2.adSaldoQuantConsig(12), tSldMesEst2.adSaldoValorConsig(12), iAno, iFilialEmpresa, sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 60979
    
    'le o SldMesEst
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 60980
    
    dQuantAtual = tSldMesEst2.dQuantInicialConsig
    dValorAtual = tSldMesEst2.dValorInicialConsig

    For iIndice = 1 To iMes
    
        dQuantAtual = dQuantAtual + tSldMesEst2.adSaldoQuantConsig(iIndice)
        dValorAtual = dValorAtual + tSldMesEst2.adSaldoValorConsig(iIndice)
        
    Next
    
    If Abs(dQuantAtual) > QTDE_ESTOQUE_DELTA Then
    
        dCustoMedio = dValorAtual / dQuantAtual
            
    Else
        
        dCustoMedio = 0
            
    End If
    
    SldMesEst2_Le_CustoMedioConsig = SUCESSO

    Exit Function

Erro_SldMesEst2_Le_CustoMedioConsig:
    
    SldMesEst2_Le_CustoMedioConsig = gErr
    
    Select Case gErr
    
        Case 60979, 60980
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESEST21", gErr, iAno, iFilialEmpresa, sProduto)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153552)
        
    End Select
    
    Exit Function

End Function

Function TransacaoCTB_Le_TransacaoTela(ByVal objTransacaoCTB As ClassTransacaoCTB) As Long
'Le os dados da transacao passada como parametro e devolve os dados em objTransacaoCTB

Dim lErro As Long
Dim lComando As Long
Dim iCodigo As Integer
Dim sTransacao As String
Dim iCliFornecedor As Integer
Dim iSubTipo As Integer
Dim iObrigatoriedade As Integer
    
On Error GoTo Erro_TransacaoCTB_Le_TransacaoTela

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 71691
    
    sTransacao = String(STRING_TRANSACAOCTB_TRANSACAO, 0)
    
    lErro = Comando_Executar(lComando, "SELECT Codigo, CliFornecedor, SubTipo, Transacao, Obrigatoriedade FROM TransacaoCTB WHERE Sigla = ? AND TransacaoTela = ?", iCodigo, iCliFornecedor, iSubTipo, sTransacao, iObrigatoriedade, objTransacaoCTB.sSigla, objTransacaoCTB.sTransacaoTela)
    If lErro <> AD_SQL_SUCESSO Then gError 71692
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then gError 71693
    
    objTransacaoCTB.iCodigo = iCodigo
    objTransacaoCTB.sTransacao = sTransacao
    objTransacaoCTB.iCliFornecedor = iCliFornecedor
    objTransacaoCTB.iSubTipo = iSubTipo
    objTransacaoCTB.iObrigatoriedade = iObrigatoriedade
    
    Call Comando_Fechar(lComando)
    
    TransacaoCTB_Le_TransacaoTela = SUCESSO
    
    Exit Function

Erro_TransacaoCTB_Le_TransacaoTela:

    TransacaoCTB_Le_TransacaoTela = gErr
    
    Select Case gErr
                
        Case 71691
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 71692, 71693
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRANSACAOCTB1", gErr, objTransacaoCTB.sSigla, objTransacaoCTB.sTransacaoTela)
                  
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153553)
        
        End Select
        
        Call Comando_Fechar(lComando)
        
    Exit Function

End Function

Function Obtem_Subtipo_Contabil(ByVal sTabela As String, ByVal iCodigo As Integer, iSubTipo As Integer) As Long

Dim lErro As Long
Dim lComando As Long
Dim sSQL As String

On Error GoTo Erro_Obtem_Subtipo_Contabil

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 79668
    
    'Define o comando SQL que será utilizado
    sSQL = "SELECT SubTipoContabil FROM " & sTabela & " WHERE Codigo = ?"
    
    'Lê o subtipo na tabela passada como parâmetro
    lErro = Comando_Executar(lComando, sSQL, iSubTipo, iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 79669
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 79670
    
    'Se não encontrou => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 79671
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Obtem_Subtipo_Contabil = SUCESSO
    
    Exit Function
    
Erro_Obtem_Subtipo_Contabil:

    Obtem_Subtipo_Contabil = gErr
    
    Select Case gErr
    
        Case 79668
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 79669, 79670
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA", gErr, sTabela)
        
        Case 79671 'Subtipo não encontrado. Deve ser tratado na rotina chamadora
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153554)
        
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function MvPerCcl_Le_SldAnterior(ByVal iFilialEmpresa As Integer, ByVal iExercicio As Integer, ByVal sCcl As String, ByVal iPeriodo As Integer, dSaldo As Double, lComando As Long) As Long
'Le o Saldo anterior do iPeriodo passado

Dim lErro As Long
Dim dCredito(NUM_MAX_PERIODOS) As Double
Dim dDebito(NUM_MAX_PERIODOS) As Double
Dim dMovimento As Double
Dim dSldIni As Double
Dim iIndice As Integer

On Error GoTo Erro_MvPerCcl_Le_SldAnterior

    If lComando = 0 Then
        lComando = Comando_AbrirExt(GL_lConexaoRel)
        If lComando = 0 Then gError 78080
    End If

    'le os saldos de centro de custo
    lErro = Comando_Executar(lComando, "SELECT SldIni, Deb01, Deb02, Deb03, Deb04, Deb05, Deb06, Deb07, Deb08, Deb09, Deb10, Deb11, Deb12, Cre01, Cre02, Cre03, Cre04, Cre05, Cre06, Cre07, Cre08, Cre09, Cre10, Cre11, Cre12 FROM MvPerCcl WHERE Ccl = ? AND FilialEmpresa = ? AND Exercicio = ? AND LEN(Conta) = 0", dSldIni, dDebito(1), dDebito(2), dDebito(3), dDebito(4), dDebito(5), dDebito(6), dDebito(7), dDebito(8), dDebito(9), dDebito(10), dDebito(11), dDebito(12), dCredito(1), dCredito(2), dCredito(3), dCredito(4), dCredito(5), dCredito(6), dCredito(7), dCredito(8), dCredito(9), dCredito(10), dCredito(11), dCredito(12), sCcl, iFilialEmpresa, iExercicio)
    If lErro <> AD_SQL_SUCESSO Then Error 78081
    
    'le o primeiro saldo de centro de custo
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 78082
    
    dMovimento = dSldIni
    
    'acumula os saldos em dMovimento
    For iIndice = 1 To iPeriodo - 1
        dMovimento = dMovimento + dCredito(iIndice) - dDebito(iIndice)
    Next
        
    dSaldo = dMovimento
    
    MvPerCcl_Le_SldAnterior = SUCESSO
    
    Exit Function
 
Erro_MvPerCcl_Le_SldAnterior:

    MvPerCcl_Le_SldAnterior = Err

    Select Case Err
    
        Case 78080
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 78081, 78082
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVPERCCL3", Err, sCcl)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153555)
    
    End Select
    
    Exit Function

End Function

Function MvPerCli_Le_SldAnterior(dSaldo As Double, ByVal lCliente As Long, ByVal dtData As Date, ByVal iFilialEmpresa As Integer, lComando As Long) As Long
'Le o Saldo anterior do cliente passado, a partir da tabela MvPerCli

Dim lErro As Long
Dim dCredito(NUM_MAX_PERIODOS) As Double
Dim dDebito(NUM_MAX_PERIODOS) As Double
Dim dMovimento As Double
Dim dSldIni As Double
Dim iIndice As Integer
Dim iMes As Integer
Dim iDia As Integer
Dim iIndice2 As Integer

On Error GoTo Erro_MvPerCli_Le_SldAnterior

    If lComando = 0 Then
        lComando = Comando_AbrirExt(GL_lConexaoRel)
        If lComando = 0 Then gError 76443
    End If

    'le os saldos do cliente
    lErro = Comando_Executar(lComando, "SELECT SldIni, Deb01, Deb02, Deb03, Deb04, Deb05, Deb06, Deb07, Deb08, Deb09, Deb10, Deb11, Deb12, Cre01, Cre02, Cre03, Cre04, Cre05, Cre06, Cre07, Cre08, Cre09, Cre10, Cre11, Cre12 FROM MvPerCli, Exercicios WHERE Cliente = ? AND FilialEmpresa = ? AND MvPerCli.Exercicio = Exercicios.Exercicio AND DataInicio <= ? And DataFim >= ?", dSldIni, dDebito(1), dDebito(2), dDebito(3), dDebito(4), dDebito(5), dDebito(6), dDebito(7), dDebito(8), dDebito(9), dDebito(10), dDebito(11), dDebito(12), dCredito(1), dCredito(2), dCredito(3), dCredito(4), dCredito(5), dCredito(6), dCredito(7), dCredito(8), dCredito(9), dCredito(10), dCredito(11), dCredito(12), lCliente, iFilialEmpresa, dtData, dtData)
    If lErro <> AD_SQL_SUCESSO Then gError 76444
    
    'le o primeiro saldo do cliente
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 76445
    
    Do While lErro = AD_SQL_SUCESSO
        
        dMovimento = Round(dMovimento + dSldIni, 2)
    
        iMes = Month(dtData)
        iDia = Day(dtData)
        
        'Verifica se é o ultimo dia do mes
        If iDia = Dias_Mes(iMes, Year(dtData)) Then
            'considera o mês da data informada
            iIndice2 = iMes
        Else
            'considera até o mês anterior à data informada
            iIndice2 = iMes - 1
        End If
        
        'acumula os saldos em dMovimento
        For iIndice = 1 To iIndice2
            dMovimento = Round(dMovimento + dCredito(iIndice) - dDebito(iIndice), 2)
        Next
            
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 76445
        
    Loop
        
    dSaldo = dMovimento
        
    MvPerCli_Le_SldAnterior = SUCESSO
    
    Exit Function
 
Erro_MvPerCli_Le_SldAnterior:

    MvPerCli_Le_SldAnterior = gErr

    Select Case gErr
    
        Case 76443
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 76444, 76445
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVPERCLI", gErr)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153556)
    
    End Select
    
    Exit Function

End Function

Function MvDiaCli_Le_SldAnterior(dSaldoDiario As Double, ByVal lCliente As Long, ByVal dtData As Date, ByVal iFilialEmpresa As Integer, lComando2 As Long) As Long
'Le o Saldo anterior do cliente passado, a partir da tabela MvDiaCli

Dim lErro As Long
Dim dCredito As Double
Dim dDebito As Double
Dim dMovimento As Double
Dim dSldIni As Double
Dim iIndice As Integer
Dim iMes As Integer
Dim iAno As Integer
Dim dtDataAnterior As Date

On Error GoTo Erro_MvDiaCli_Le_SldAnterior

    If lComando2 = 0 Then
        lComando2 = Comando_AbrirExt(GL_lConexaoRel)
        If lComando2 = 0 Then gError 76447
    End If
    
    iMes = Month(dtData)
    iAno = Year(dtData)
    dtDataAnterior = CDate("01/" & Month(dtData) & "/" & Year(dtData))
    
    'le os saldos do cliente, no intervalo de dias informado
    lErro = Comando_Executar(lComando2, "SELECT Deb, Cre FROM MvDiaCli WHERE Cliente = ? AND FilialEmpresa = ? AND Data>=? AND Data<=?", dDebito, dCredito, lCliente, iFilialEmpresa, dtDataAnterior, dtData)
    If lErro <> AD_SQL_SUCESSO Then gError 76448
    
    'le o primeiro saldo do cliente
    lErro = Comando_BuscarPrimeiro(lComando2)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 76449
    
    Do While lErro = AD_SQL_SUCESSO
    
        'Soma os valores encontrados
        dMovimento = dMovimento + dCredito - dDebito
    
        lErro = Comando_BuscarProximo(lComando2)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 76450
        
    Loop
    
    dSaldoDiario = dMovimento
        
    MvDiaCli_Le_SldAnterior = SUCESSO
    
    Exit Function
 
Erro_MvDiaCli_Le_SldAnterior:

    MvDiaCli_Le_SldAnterior = gErr

    Select Case gErr
    
        Case 76447
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 76448, 76449, 76450
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVDIACLI", gErr)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153557)
    
    End Select
    
    Exit Function

End Function

Function MvPerForn_Le_SldAnterior(dSaldo As Double, ByVal lFornecedor As Long, ByVal dtData As Date, ByVal iFilialEmpresa As Integer, lComando As Long) As Long
'Le o Saldo anterior do Fornecedor passado, a partir da tabela MvPerForn

Dim lErro As Long
Dim dCredito(NUM_MAX_PERIODOS) As Double
Dim dDebito(NUM_MAX_PERIODOS) As Double
Dim dMovimento As Double
Dim dSldIni As Double
Dim iIndice As Integer
Dim iMes As Integer
Dim iDia As Integer
Dim iIndice2 As Integer

On Error GoTo Erro_MvPerForn_Le_SldAnterior

    If lComando = 0 Then
        lComando = Comando_AbrirExt(GL_lConexaoRel)
        If lComando = 0 Then gError 76454
    End If

    'le os saldos do fornecedor
    lErro = Comando_Executar(lComando, "SELECT SldIni, Deb01, Deb02, Deb03, Deb04, Deb05, Deb06, Deb07, Deb08, Deb09, Deb10, Deb11, Deb12, Cre01, Cre02, Cre03, Cre04, Cre05, Cre06, Cre07, Cre08, Cre09, Cre10, Cre11, Cre12 FROM MvPerForn, Exercicios WHERE Fornecedor = ? AND FilialEmpresa = ? AND MvPerForn.Exercicio = Exercicios.Exercicio AND DataInicio <= ? And DataFim >= ?", dSldIni, dDebito(1), dDebito(2), dDebito(3), dDebito(4), dDebito(5), dDebito(6), dDebito(7), dDebito(8), dDebito(9), dDebito(10), dDebito(11), dDebito(12), dCredito(1), dCredito(2), dCredito(3), dCredito(4), dCredito(5), dCredito(6), dCredito(7), dCredito(8), dCredito(9), dCredito(10), dCredito(11), dCredito(12), lFornecedor, iFilialEmpresa, dtData, dtData)
    If lErro <> AD_SQL_SUCESSO Then gError 76455
    
    'le o primeiro saldo do fornecedor
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 76456
    
    Do While lErro = AD_SQL_SUCESSO
    
        dMovimento = Round(dMovimento + dSldIni, 2)
        
        iMes = Month(dtData)
        iDia = Day(dtData)
        
        'Verifica se é o ultimo dia do mes
        If iDia = Dias_Mes(iMes, Year(dtData)) Then
            'considera o mês da data informada
            iIndice2 = iMes
        Else
            'considera até o mês anterior à data informada
            iIndice2 = iMes - 1
        End If
    
        'acumula os saldos em dMovimento
        For iIndice = 1 To iIndice2
            dMovimento = Round(dMovimento + dCredito(iIndice) - dDebito(iIndice), 2)
        Next
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 76456
    
    Loop
        
    dSaldo = dMovimento
        
    MvPerForn_Le_SldAnterior = SUCESSO
    
    Exit Function
 
Erro_MvPerForn_Le_SldAnterior:

    MvPerForn_Le_SldAnterior = gErr

    Select Case gErr
    
        Case 76454
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 76455, 76456
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVPERFORN", gErr, lFornecedor)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153558)
    
    End Select
    
    Exit Function

End Function

Function MvDiaForn_Le_SldAnterior(dSaldoDiario As Double, ByVal lFornecedor As Long, ByVal dtData As Date, ByVal iFilialEmpresa As Integer, lComando2 As Long) As Long
'Le o Saldo anterior do Fornecedor passado, a partir da tabela MvDiaForn

Dim lErro As Long
Dim dCredito As Double
Dim dDebito As Double
Dim dMovimento As Double
Dim dSldIni As Double
Dim iIndice As Integer
Dim iMes As Integer
Dim iAno As Integer
Dim dtDataAnterior As Date

On Error GoTo Erro_MvDiaForn_Le_SldAnterior

    If lComando2 = 0 Then
        lComando2 = Comando_AbrirExt(GL_lConexaoRel)
        If lComando2 = 0 Then gError 76457
    End If
    
    iMes = Month(dtData)
    iAno = Year(dtData)
    dtDataAnterior = CDate("01/" & Month(dtData) & "/" & Year(dtData))
    
    'le os saldos do fornecedor
    lErro = Comando_Executar(lComando2, "SELECT Deb, Cre FROM MvDiaForn WHERE Fornecedor = ? AND FilialEmpresa = ? AND Data>=? AND Data<=?", dDebito, dCredito, lFornecedor, iFilialEmpresa, dtDataAnterior, dtData)
    If lErro <> AD_SQL_SUCESSO Then gError 76458
    
    'le o primeiro saldo do fornecedor
    lErro = Comando_BuscarPrimeiro(lComando2)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 76459
    
    Do While lErro = AD_SQL_SUCESSO
    
        dMovimento = dMovimento + dCredito - dDebito
    
        lErro = Comando_BuscarProximo(lComando2)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 76460
        
    Loop
    
    dSaldoDiario = dMovimento
        
    MvDiaForn_Le_SldAnterior = SUCESSO
    
    Exit Function
 
Erro_MvDiaForn_Le_SldAnterior:

    MvDiaForn_Le_SldAnterior = gErr

    Select Case gErr
    
        Case 76457
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 76458, 76459, 76460
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVDIAFORN", gErr)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153559)
    
    End Select
    
    Exit Function

End Function

Function CustoMedio_Le(alComando() As Long, ByVal objLancamento_Detalhe As ClassLancamento_Detalhe, ByVal iApropriacao As Integer) As Long
'le o custo medio, standard ou custo dos escaninhos

Dim lErro As Long
Dim dCustoMedio As Double
Dim dCustoStandard As Double

On Error GoTo Erro_CustoMedio_Le

    If iApropriacao = APROPR_CUSTO_MEDIO And objLancamento_Detalhe.iEscaninho_Custo = ESCANINHO_NOSSO Then

        lErro = CF("SldMesEst_Le_CustoMedio", alComando(4), objLancamento_Detalhe.iFilialEmpresa, Year(objLancamento_Detalhe.dtDataEstoque), objLancamento_Detalhe.sProduto, Month(objLancamento_Detalhe.dtDataEstoque), dCustoMedio)
        If lErro <> SUCESSO Then gError 36525
        
        objLancamento_Detalhe.dValor = objLancamento_Detalhe.dValor * dCustoMedio
        
'            reprocessamento
'            objLancamento_Detalhe.sProduto = ""
        
    ElseIf iApropriacao = APROPR_CUSTO_STANDARD And objLancamento_Detalhe.iEscaninho_Custo = ESCANINHO_NOSSO Then
    
        lErro = CF("SldMesEst_Le_CustoStandard", alComando(5), objLancamento_Detalhe.iFilialEmpresa, Year(objLancamento_Detalhe.dtDataEstoque), objLancamento_Detalhe.sProduto, Month(objLancamento_Detalhe.dtDataEstoque), dCustoStandard)
        If lErro <> SUCESSO Then gError 36533
        
        objLancamento_Detalhe.dValor = objLancamento_Detalhe.dValor * dCustoStandard
    
'            reprocessamento
'            objLancamento_Detalhe.sProduto = ""
        
    ElseIf iApropriacao = APROPR_CUSTO_MEDIO Or iApropriacao = APROPR_CUSTO_STANDARD Then
    
        If objLancamento_Detalhe.iEscaninho_Custo = ESCANINHO_NOSSO_EM_CONSIGNACAO Then
    
            lErro = CF("SldMesEst2_Le_CustoMedioConsig", alComando(6), objLancamento_Detalhe.iFilialEmpresa, Year(objLancamento_Detalhe.dtDataEstoque), objLancamento_Detalhe.sProduto, Month(objLancamento_Detalhe.dtDataEstoque), dCustoMedio)
            If lErro <> SUCESSO Then gError 71568
                
        ElseIf objLancamento_Detalhe.iEscaninho_Custo = ESCANINHO_NOSSO_EM_CONSERTO Then
        
            lErro = CF("SldMesEst2_Le_CustoMedioConserto", alComando(7), objLancamento_Detalhe.iFilialEmpresa, Year(objLancamento_Detalhe.dtDataEstoque), objLancamento_Detalhe.sProduto, Month(objLancamento_Detalhe.dtDataEstoque), dCustoMedio)
            If lErro <> SUCESSO Then gError 71569

        ElseIf objLancamento_Detalhe.iEscaninho_Custo = ESCANINHO_NOSSO_EM_DEMO Then
        
            lErro = CF("SldMesEst2_Le_CustoMedioDemo", alComando(8), objLancamento_Detalhe.iFilialEmpresa, Year(objLancamento_Detalhe.dtDataEstoque), objLancamento_Detalhe.sProduto, Month(objLancamento_Detalhe.dtDataEstoque), dCustoMedio)
            If lErro <> SUCESSO Then gError 71570
            
        ElseIf objLancamento_Detalhe.iEscaninho_Custo = ESCANINHO_NOSSO_EM_OUTROS Then
        
            lErro = CF("SldMesEst2_Le_CustoMedioOutros", alComando(9), objLancamento_Detalhe.iFilialEmpresa, Year(objLancamento_Detalhe.dtDataEstoque), objLancamento_Detalhe.sProduto, Month(objLancamento_Detalhe.dtDataEstoque), dCustoMedio)
            If lErro <> SUCESSO Then gError 71571
            
        ElseIf objLancamento_Detalhe.iEscaninho_Custo = ESCANINHO_NOSSO_EM_BENEF Then
        
            lErro = CF("SldMesEst2_Le_CustoMedioBenef", alComando(10), objLancamento_Detalhe.iFilialEmpresa, Year(objLancamento_Detalhe.dtDataEstoque), objLancamento_Detalhe.sProduto, Month(objLancamento_Detalhe.dtDataEstoque), dCustoMedio)
            If lErro <> SUCESSO Then gError 71572
        
        ElseIf objLancamento_Detalhe.iEscaninho_Custo = ESCANINHO_3_EM_CONSIGNACAO Then
        
            lErro = CF("SldMesEst1_Le_CustoMedioConsig3", alComando(11), objLancamento_Detalhe.iFilialEmpresa, Year(objLancamento_Detalhe.dtDataEstoque), objLancamento_Detalhe.sProduto, Month(objLancamento_Detalhe.dtDataEstoque), dCustoMedio)
            If lErro <> SUCESSO Then gError 71573
        
        ElseIf objLancamento_Detalhe.iEscaninho_Custo = ESCANINHO_3_EM_CONSERTO Then
        
            lErro = CF("SldMesEst1_Le_CustoMedioConserto3", alComando(12), objLancamento_Detalhe.iFilialEmpresa, Year(objLancamento_Detalhe.dtDataEstoque), objLancamento_Detalhe.sProduto, Month(objLancamento_Detalhe.dtDataEstoque), dCustoMedio)
            If lErro <> SUCESSO Then gError 71574

        ElseIf objLancamento_Detalhe.iEscaninho_Custo = ESCANINHO_3_EM_DEMO Then
        
            lErro = CF("SldMesEst1_Le_CustoMedioDemo3", alComando(13), objLancamento_Detalhe.iFilialEmpresa, Year(objLancamento_Detalhe.dtDataEstoque), objLancamento_Detalhe.sProduto, Month(objLancamento_Detalhe.dtDataEstoque), dCustoMedio)
            If lErro <> SUCESSO Then gError 71575
            
        ElseIf objLancamento_Detalhe.iEscaninho_Custo = ESCANINHO_3_EM_OUTROS Then
        
            lErro = CF("SldMesEst1_Le_CustoMedioOutros3", alComando(14), objLancamento_Detalhe.iFilialEmpresa, Year(objLancamento_Detalhe.dtDataEstoque), objLancamento_Detalhe.sProduto, Month(objLancamento_Detalhe.dtDataEstoque), dCustoMedio)
            If lErro <> SUCESSO Then gError 71576
            
        ElseIf objLancamento_Detalhe.iEscaninho_Custo = ESCANINHO_3_EM_BENEF Then
        
            lErro = CF("SldMesEst1_Le_CustoMedioBenef3", alComando(15), objLancamento_Detalhe.iFilialEmpresa, Year(objLancamento_Detalhe.dtDataEstoque), objLancamento_Detalhe.sProduto, Month(objLancamento_Detalhe.dtDataEstoque), dCustoMedio)
            If lErro <> SUCESSO Then gError 71577
            
        End If
        
        objLancamento_Detalhe.dValor = objLancamento_Detalhe.dValor * dCustoMedio
        
'            reprocessamento
'            objLancamento_Detalhe.sProduto = ""
        
    End If

    CustoMedio_Le = SUCESSO
    
    Exit Function
    
Erro_CustoMedio_Le:

    CustoMedio_Le = gErr
    
    Select Case gErr
                   
        Case 36525, 36533, 71568, 71569, 71570, 71571, 71572, 71573, 71574, 71575, 71576, 71577
                   
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153560)
            
    End Select
    
    Exit Function


End Function

Function Lancamento_Le_Doc3(ByVal iFilialEmpresa As Integer, ByVal iTransacao As Integer, ByVal lNumIntDoc As Long, ByVal colLancamentos As Collection) As Long
'le os lançamentos relativos ao documento passado como parametro inclusive os com status <> 0 e coloca-os em colLancamentos

Dim lComando As Long
Dim lErro As Long
Dim objLanc As ClassLancamento_Detalhe
Dim tLanc As typeLancamento

On Error GoTo Erro_Lancamento_Le_Doc3

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 89097
    
    tLanc.sOrigem = String(STRING_ORIGEM, 0)
    tLanc.sConta = String(STRING_CONTA, 0)
    tLanc.sCcl = String(STRING_CCL, 0)
    tLanc.sHistorico = String(STRING_HISTORICO, 0)
    tLanc.sProduto = String(STRING_PRODUTO, 0)
    tLanc.sModelo = String(STRING_PADRAOCONTAB_MODELO, 0)
    tLanc.sDocOrigem = String(STRING_LANPENDENTE_DOCORIGEM, 0)
    
    lErro = Comando_Executar(lComando, "SELECT  Origem, Exercicio, PeriodoLan, Doc, Seq, Lote, PeriodoLote, Data, Conta, Ccl, Historico, Valor, NumIntDoc, FilialCliForn , CliForn, Transacao, Aglutinado, DocAglutinado, ContaSimples, SeqContraPartida, Produto, Quantidade, DataEstoque, ApropriaCrProd, EscaninhoCusto, ValorLivroAuxiliar, ClienteFornecedor, DocOrigem, Status, Modelo FROM Lancamentos WHERE FilialEmpresa = ? AND Transacao = ? AND NumIntDoc = ? ", _
    tLanc.sOrigem, tLanc.iExercicio, tLanc.iPeriodoLan, tLanc.lDoc, tLanc.iSeq, tLanc.iLote, tLanc.iPeriodoLote, tLanc.dtData, tLanc.sConta, tLanc.sCcl, tLanc.sHistorico, tLanc.dValor, tLanc.lNumIntDoc, tLanc.iFilialCliForn, tLanc.lCliForn, tLanc.iTransacao, tLanc.iAglutina, tLanc.lDocAglutinado, tLanc.lContaSimples, tLanc.iSeqContraPartida, tLanc.sProduto, tLanc.dQuantidade, tLanc.dtDataEstoque, tLanc.iApropriaCRProd, tLanc.iEscaninho_Custo, tLanc.dValorLivroAuxiliar, tLanc.iClienteFornecedor, tLanc.sDocOrigem, tLanc.iStatus, tLanc.sModelo, iFilialEmpresa, iTransacao, lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 89098
        
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89099
        
    If lErro = AD_SQL_SEM_DADOS Then gError 89100
        
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objLanc = New ClassLancamento_Detalhe
        
        objLanc.iFilialEmpresa = iFilialEmpresa
        objLanc.sOrigem = tLanc.sOrigem
        objLanc.iExercicio = tLanc.iExercicio
        objLanc.iPeriodoLan = tLanc.iPeriodoLan
        objLanc.lDoc = tLanc.lDoc
        objLanc.iSeq = tLanc.iSeq
        objLanc.iLote = tLanc.iLote
        objLanc.iPeriodoLote = tLanc.iPeriodoLote
        objLanc.dtData = tLanc.dtData
        objLanc.sConta = tLanc.sConta
        objLanc.sCcl = tLanc.sCcl
        objLanc.sHistorico = tLanc.sHistorico
        objLanc.dValor = tLanc.dValor
        objLanc.lNumIntDoc = tLanc.lNumIntDoc
        objLanc.iFilialCliForn = tLanc.iFilialCliForn
        objLanc.lCliForn = tLanc.lCliForn
        objLanc.iTransacao = tLanc.iTransacao
        objLanc.iAglutina = tLanc.iAglutina
        objLanc.lContaSimples = tLanc.lContaSimples
        objLanc.iSeqContraPartida = tLanc.iSeqContraPartida
        objLanc.sProduto = tLanc.sProduto
        objLanc.dQuantidade = tLanc.dQuantidade
        objLanc.dtDataEstoque = tLanc.dtDataEstoque
        objLanc.sModelo = tLanc.sModelo
        objLanc.dValorLivroAuxiliar = tLanc.dValorLivroAuxiliar
        objLanc.iApropriaCRProd = tLanc.iApropriaCRProd
        objLanc.iEscaninho_Custo = tLanc.iEscaninho_Custo
        objLanc.iClienteFornecedor = tLanc.iClienteFornecedor
        objLanc.sDocOrigem = tLanc.sDocOrigem
        objLanc.iStatus = tLanc.iStatus
        objLanc.sModelo = tLanc.sModelo
                
        colLancamentos.Add objLanc
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89101
        
    Loop
    
    Call Comando_Fechar(lComando)
    
    Lancamento_Le_Doc3 = SUCESSO
    
    Exit Function
 
Erro_Lancamento_Le_Doc3:

    Lancamento_Le_Doc3 = gErr
    
    Select Case gErr
    
        Case 89097
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 89098, 89099, 89101
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LANCAMENTOS3", gErr)
           
        Case 89100
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153561)
    
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function Lancamento_Le_Doc5(ByVal iFilialEmpresa As Integer, ByVal lNumIntDocBaixaParcPag As Long, ByVal lNumIntDocBaixa As Long, ByVal colLancamentos As Collection) As Long
'le os lançamentos relativos ao documento passado como parametro inclusive os com status <> 0 e coloca-os em colLancamentos
'le os lancamentos relativos as origens de lançamento (OrigemLcto) = 2(BaixasPag) ou 3(BaixasParcPag)

Dim lComando As Long
Dim lErro As Long
Dim objLanc As ClassLancamento_Detalhe
Dim tLanc As typeLancamento

On Error GoTo Erro_Lancamento_Le_Doc5

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 92671
    
    tLanc.sOrigem = String(STRING_ORIGEM, 0)
    tLanc.sConta = String(STRING_CONTA, 0)
    tLanc.sCcl = String(STRING_CCL, 0)
    tLanc.sHistorico = String(STRING_HISTORICO, 0)
    tLanc.sProduto = String(STRING_PRODUTO, 0)
    tLanc.sModelo = String(STRING_PADRAOCONTAB_MODELO, 0)
    tLanc.sDocOrigem = String(STRING_LANPENDENTE_DOCORIGEM, 0)
    
    lErro = Comando_Executar(lComando, "SELECT  Origem, Exercicio, PeriodoLan, Doc, Seq, Lote, PeriodoLote, Data, Conta, Ccl, Historico, Valor, NumIntDoc, FilialCliForn , CliForn, Lancamentos.Transacao, Aglutinado, DocAglutinado, ContaSimples, SeqContraPartida, Produto, Quantidade, DataEstoque, ApropriaCrProd, EscaninhoCusto, ValorLivroAuxiliar, ClienteFornecedor, DocOrigem, Status, Modelo FROM Lancamentos, TransacaoCTB WHERE FilialEmpresa = ? AND Lancamentos.Transacao = Codigo AND ((OrigemLcto = ? AND NumIntDoc = ?) OR (OrigemLcto = ? AND NumIntDoc = ?))  ", _
    tLanc.sOrigem, tLanc.iExercicio, tLanc.iPeriodoLan, tLanc.lDoc, tLanc.iSeq, tLanc.iLote, tLanc.iPeriodoLote, tLanc.dtData, tLanc.sConta, tLanc.sCcl, tLanc.sHistorico, tLanc.dValor, tLanc.lNumIntDoc, tLanc.iFilialCliForn, tLanc.lCliForn, tLanc.iTransacao, tLanc.iAglutina, tLanc.lDocAglutinado, tLanc.lContaSimples, tLanc.iSeqContraPartida, tLanc.sProduto, tLanc.dQuantidade, tLanc.dtDataEstoque, tLanc.iApropriaCRProd, tLanc.iEscaninho_Custo, tLanc.dValorLivroAuxiliar, tLanc.iClienteFornecedor, tLanc.sDocOrigem, tLanc.iStatus, tLanc.sModelo, iFilialEmpresa, TRANSACAOCTBORIGEM_BAIXASPAG, lNumIntDocBaixa, TRANSACAOCTBORIGEM_BAIXASPARCPAG, lNumIntDocBaixaParcPag)
    If lErro <> AD_SQL_SUCESSO Then gError 92672
        
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92673
        
    If lErro = AD_SQL_SEM_DADOS Then gError 92674
        
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objLanc = New ClassLancamento_Detalhe
        
        objLanc.iFilialEmpresa = iFilialEmpresa
        objLanc.sOrigem = tLanc.sOrigem
        objLanc.iExercicio = tLanc.iExercicio
        objLanc.iPeriodoLan = tLanc.iPeriodoLan
        objLanc.lDoc = tLanc.lDoc
        objLanc.iSeq = tLanc.iSeq
        objLanc.iLote = tLanc.iLote
        objLanc.iPeriodoLote = tLanc.iPeriodoLote
        objLanc.dtData = tLanc.dtData
        objLanc.sConta = tLanc.sConta
        objLanc.sCcl = tLanc.sCcl
        objLanc.sHistorico = tLanc.sHistorico
        objLanc.dValor = tLanc.dValor
        objLanc.lNumIntDoc = tLanc.lNumIntDoc
        objLanc.iFilialCliForn = tLanc.iFilialCliForn
        objLanc.lCliForn = tLanc.lCliForn
        objLanc.iTransacao = tLanc.iTransacao
        objLanc.iAglutina = tLanc.iAglutina
        objLanc.lContaSimples = tLanc.lContaSimples
        objLanc.iSeqContraPartida = tLanc.iSeqContraPartida
        objLanc.sProduto = tLanc.sProduto
        objLanc.dQuantidade = tLanc.dQuantidade
        objLanc.dtDataEstoque = tLanc.dtDataEstoque
        objLanc.sModelo = tLanc.sModelo
        objLanc.dValorLivroAuxiliar = tLanc.dValorLivroAuxiliar
        objLanc.iApropriaCRProd = tLanc.iApropriaCRProd
        objLanc.iEscaninho_Custo = tLanc.iEscaninho_Custo
        objLanc.iClienteFornecedor = tLanc.iClienteFornecedor
        objLanc.sDocOrigem = tLanc.sDocOrigem
        objLanc.iStatus = tLanc.iStatus
        objLanc.sModelo = tLanc.sModelo
                
        colLancamentos.Add objLanc
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92675
        
    Loop
    
    Call Comando_Fechar(lComando)
    
    Lancamento_Le_Doc5 = SUCESSO
    
    Exit Function
 
Erro_Lancamento_Le_Doc5:

    Lancamento_Le_Doc5 = gErr
    
    Select Case gErr
    
        Case 92671
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 92672, 92673, 92675
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LANCAMENTOS3", gErr)
           
        Case 92674
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153562)
    
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function LanPendente_Le_Doc4(ByVal iFilialEmpresa As Integer, ByVal lNumIntDocBaixaParcPag As Long, ByVal lNumIntDocBaixa As Long, colLancamentos As Collection) As Long
'le os lançamentos pendentes relativos ao documento passado como parametro e coloca-os em colLancamentos
'le os lancamentos pendentes relativos às origens de lançamento (OrigemLcto) = 2(BaixasPag) ou 3(BaixasParcPag)

Dim lComando As Long
Dim lErro As Long
Dim objLanc As ClassLancamento_Detalhe
Dim tLanc As typeLancamento

On Error GoTo Erro_LanPendente_Le_Doc4

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 92676
    
    tLanc.sOrigem = String(STRING_ORIGEM, 0)
    tLanc.sConta = String(STRING_CONTA, 0)
    tLanc.sCcl = String(STRING_CCL, 0)
    tLanc.sHistorico = String(STRING_HISTORICO, 0)
    tLanc.sProduto = String(STRING_PRODUTO, 0)
    tLanc.sModelo = String(STRING_PADRAOCONTAB_MODELO, 0)
    tLanc.sDocOrigem = String(STRING_LANPENDENTE_DOCORIGEM, 0)
    
    lErro = Comando_Executar(lComando, "SELECT  Origem, Exercicio, PeriodoLan, Doc, Seq, Lote, PeriodoLote, Data, Conta, Ccl, Historico, Valor, NumIntDoc, FilialCliForn, CliForn, LanPendente.Transacao, Aglutina, ContaSimples, SeqContraPartida, Status, Modelo, ValorLivroAuxiliar, ClienteFornecedor, Produto, ApropriaCrProd, EscaninhoCusto, DocOrigem, Quantidade, DataEstoque FROM LanPendente, TransacaoCTB WHERE FilialEmpresa = ? AND LanPendente.Transacao = Codigo AND ((OrigemLcto = ? AND NumIntDoc = ?) OR (OrigemLcto = ? AND NumIntDoc = ?)) ", _
    tLanc.sOrigem, tLanc.iExercicio, tLanc.iPeriodoLan, tLanc.lDoc, tLanc.iSeq, tLanc.iLote, tLanc.iPeriodoLote, tLanc.dtData, tLanc.sConta, tLanc.sCcl, tLanc.sHistorico, tLanc.dValor, tLanc.lNumIntDoc, tLanc.iFilialCliForn, tLanc.lCliForn, tLanc.iTransacao, tLanc.iAglutina, tLanc.lContaSimples, tLanc.iSeqContraPartida, tLanc.iStatus, tLanc.sModelo, tLanc.dValorLivroAuxiliar, tLanc.iClienteFornecedor, tLanc.sProduto, tLanc.iApropriaCRProd, tLanc.iEscaninho_Custo, tLanc.sDocOrigem, tLanc.dQuantidade, tLanc.dtDataEstoque, iFilialEmpresa, TRANSACAOCTBORIGEM_BAIXASPAG, lNumIntDocBaixa, TRANSACAOCTBORIGEM_BAIXASPARCPAG, lNumIntDocBaixaParcPag)
    If lErro <> AD_SQL_SUCESSO Then gError 92677
        
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92678
        
    If lErro = AD_SQL_SEM_DADOS Then gError 92679
        
    Set colLancamentos = New Collection

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objLanc = New ClassLancamento_Detalhe
        
        objLanc.iFilialEmpresa = iFilialEmpresa
        objLanc.sOrigem = tLanc.sOrigem
        objLanc.iExercicio = tLanc.iExercicio
        objLanc.iPeriodoLan = tLanc.iPeriodoLan
        objLanc.lDoc = tLanc.lDoc
        objLanc.iSeq = tLanc.iSeq
        objLanc.iLote = tLanc.iLote
        objLanc.iPeriodoLote = tLanc.iPeriodoLote
        objLanc.dtData = tLanc.dtData
        objLanc.sConta = tLanc.sConta
        objLanc.sCcl = tLanc.sCcl
        objLanc.sHistorico = tLanc.sHistorico
        objLanc.dValor = tLanc.dValor
        objLanc.lNumIntDoc = tLanc.lNumIntDoc
        objLanc.iFilialCliForn = tLanc.iFilialCliForn
        objLanc.lCliForn = tLanc.lCliForn
        objLanc.iTransacao = tLanc.iTransacao
        objLanc.iAglutina = tLanc.iAglutina
        objLanc.lContaSimples = tLanc.lContaSimples
        objLanc.iSeqContraPartida = tLanc.iSeqContraPartida
        objLanc.iStatus = tLanc.iStatus
        objLanc.sProduto = tLanc.sProduto
        objLanc.sModelo = tLanc.sModelo
        objLanc.dValorLivroAuxiliar = tLanc.dValorLivroAuxiliar
        objLanc.iClienteFornecedor = tLanc.iClienteFornecedor
        objLanc.iApropriaCRProd = tLanc.iApropriaCRProd
        objLanc.iEscaninho_Custo = tLanc.iEscaninho_Custo
        objLanc.sDocOrigem = tLanc.sDocOrigem
        objLanc.dQuantidade = tLanc.dQuantidade
        objLanc.dtDataEstoque = tLanc.dtDataEstoque
        
        colLancamentos.Add objLanc
            
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92680
        
    Loop
    
    Call Comando_Fechar(lComando)
    
    LanPendente_Le_Doc4 = SUCESSO
    
    Exit Function
 
Erro_LanPendente_Le_Doc4:

    LanPendente_Le_Doc4 = gErr
    
    Select Case gErr
    
        Case 92676
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 92677, 92678, 92680
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LANPENDENTE3", gErr)
           
        Case 92679
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153563)
    
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

'Revisado - OK Luiz 19/02/02
Public Function PlanoContaHistorico_Le_Todos(ByVal colPlanoContaHistorico As Collection) As Long
'Lê todos os registros da tabela PlanoContaHistorico e devolve-os na coleção
'colPlanoContaHistorico RETORNA(Output) os registros lidos

Dim lErro As Long
Dim lComando As Long
Dim sSelect As String
Dim tPlanoContaHistorico As typePlanoContaHistorico
Dim objPlanoContaHistorico As ClassPlanoContaHistorico

On Error GoTo Erro_PlanoContaHistorico_Le_Todos

    'Abre o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 97811
    
    'Carrega sSelect com a Expressão SQL
    sSelect = "SELECT PlanoContaHistorico.DataAtualizacao, PlanoContaHistorico.Codigo, PlanoContaHistorico.Tipo, PlanoContaHistorico.ContaPai, PlanoContaHistorico.Descricao FROM PlanoContaHistorico"
    
    With tPlanoContaHistorico
        
        'Inicializa as Strings do Type
        .sCodigo = String(STRING_CONTA, 0)
        .sDescricao = String(STRING_CONTA_DESCRICAO, 0)
        .sContaPai = String(STRING_CONTA, 0)
            
        'Executa o Comando SQL
        lErro = Comando_Executar(lComando, sSelect, .dtDataAtualizacao, .sCodigo, .iTipo, .sContaPai, .sDescricao)
        If lErro <> AD_SQL_SUCESSO Then gError 97812
    
    End With
        
    'Busca o Primeiro Registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97813

    'Se nao encontrou => Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 97810
    
    'Enquanto tem PlanoConta => adiciona na colecao
    Do While lErro <> AD_SQL_SEM_DADOS
        
        'Inicializa o obj
        Set objPlanoContaHistorico = New ClassPlanoContaHistorico
        
        With objPlanoContaHistorico
    
            .dtDataAtualizacao = tPlanoContaHistorico.dtDataAtualizacao
            .iTipo = tPlanoContaHistorico.iTipo
            .sCodigo = tPlanoContaHistorico.sCodigo
            .sDescricao = tPlanoContaHistorico.sDescricao
            .sContaPai = tPlanoContaHistorico.sContaPai
            
        End With
    
        'Adiciona na Colecao
        colPlanoContaHistorico.Add objPlanoContaHistorico
    
        'Busca o Próximo Registro
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97814
            
    Loop
    
    'Fecha o Comando
    Call Comando_Fechar(lComando)
    
    PlanoContaHistorico_Le_Todos = SUCESSO
    
    Exit Function
    
Erro_PlanoContaHistorico_Le_Todos:

    PlanoContaHistorico_Le_Todos = gErr

    Select Case gErr
        
        Case 97811
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 97812, 97813, 97814
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTAHISTORICO", gErr)
        
        Case 97810 'Sem dados
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153564)
    
    End Select

    'Fecha o Comando
    Call Comando_Fechar(lComando)
    
End Function

'Revisado - OK Luiz 19/02/02
Public Function CclHistorico_Le_Todos(ByVal colCclHistorico As Collection) As Long
'Lê todos os registros da tabela CclHistorico e devolve-os na coleção
'colCclHistorico RETORNA(Output) os registros lidos

Dim lErro As Long
Dim lComando As Long
Dim sSelect As String
Dim tCclHistorico As typeCclHistorico
Dim objCclHistorico As ClassCclHistorico

On Error GoTo Erro_CclHistorico_Le_Todos

    'Abre o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 97815
    
    'Carrega sSelect com a Expressão SQL
    sSelect = "SELECT CclHistorico.DataAtualizacao, CclHistorico.CodigoCcl, CclHistorico.TipoCcl, CclHistorico.DescCcl FROM CclHistorico"
    
    With tCclHistorico
        
        'Inicializa as Strings do Type
        .sCodigoCcl = String(STRING_CCL, 0)
        .sDescCcl = String(STRING_CCL_DESCRICAO, 0)
            
        'Executa o Comando SQL
        lErro = Comando_Executar(lComando, sSelect, .dtDataAtualizacao, .sCodigoCcl, .iTipoCcl, .sDescCcl)
        If lErro <> AD_SQL_SUCESSO Then gError 97816
    
    End With
        
    'Busca o Primeiro Registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97817

    'Se nao encontrou => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 97818
    
    'Enquanto tem Ccl => adiciona na colecao
    Do While lErro <> AD_SQL_SEM_DADOS
                                    
        'Inicializa o Obj
        Set objCclHistorico = New ClassCclHistorico
        
        With objCclHistorico
    
            .dtDataAtualizacao = tCclHistorico.dtDataAtualizacao
            .iTipoCcl = tCclHistorico.iTipoCcl
            .sCodigoCcl = tCclHistorico.sCodigoCcl
            .sDescCcl = tCclHistorico.sDescCcl
            
        End With
    
        'Adiciona na Colecao
        colCclHistorico.Add objCclHistorico
    
        'Busca o Próximo Registro
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97819
        
    Loop
    
    'Fecha o Comando
    Call Comando_Fechar(lComando)
        
    CclHistorico_Le_Todos = SUCESSO
    
    Exit Function
    
Erro_CclHistorico_Le_Todos:

    CclHistorico_Le_Todos = gErr

    Select Case gErr
        
        Case 97815
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 97816, 97817, 97819
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CCLHISTORICO", gErr)
            
        Case 97818 'Sem dados
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153565)
    
    End Select

    'Fecha o Comando
    Call Comando_Fechar(lComando)
    
End Function

Public Function Preenche_Historico_Lanc_Aglutinado(objLancamento As ClassLancamentos) As Long
'Coloca o Histórico do Lancamento Aglutinado passado como parametro = "movimento nesta data"

On Error GoTo Erro_Preenche_Historico_Lanc_Aglutinado

    'Preenche o Histórico do Lançamento
    objLancamento.sHistorico = "Movimento nesta data"
    
    Preenche_Historico_Lanc_Aglutinado = SUCESSO
    
    Exit Function
    
Erro_Preenche_Historico_Lanc_Aglutinado:

    Preenche_Historico_Lanc_Aglutinado = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153566)
    
    End Select

End Function

Function MvPerCcl_Calcula_Valor(ByVal iFilialEmpresa As Integer, ByVal iExercicio As Integer, ByVal sContaInicial As String, ByVal sContaFinal As String, dValorExercAtual As Double, dValorExercAnt As Double, ByVal sCclInicial As String, ByVal sCclFinal As String, Optional ByVal lComando As Long = 0) As Long
'Calcula Valor pela soma de SldIni com os valores de crédito e subtrai os débitos

Dim lErro As Long
Dim lComando1 As Long
Dim dSldIni As Double
Dim adValCre(1 To 12) As Double
Dim adValDeb(1 To 12) As Double
Dim iIndice As Integer
Dim dValor As Double
Dim iExercicio1 As Integer
Dim alComando(1 To 2) As Long
Dim dValorAtual As Double
Dim dValorAnt As Double
    
On Error GoTo Erro_MvPerCcl_Calcula_Valor

    If lComando = 0 Then
        lComando1 = Comando_Abrir()
        If lComando1 = 0 Then gError 178923
    Else
        lComando1 = lComando
    End If
    
    'Abrir comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 207355
    Next
    
    
    dValorExercAnt = 0
    dValorExercAtual = 0
    
    If Len(Trim(sContaInicial)) > 0 And Len(Trim(sCclInicial)) = 0 Then
    
        lErro = Comando_Executar(lComando1, "SELECT Exercicio, SUM(SldIni+Cre01+Cre02+Cre03+Cre04+Cre05+Cre06+Cre07+Cre08+Cre09+Cre10+Cre11+Cre12-Deb01-Deb02-Deb03-Deb04-Deb05-Deb06-Deb07-Deb08-Deb09-Deb10-Deb11-Deb12) FROM MvPerCta, PlanoConta WHERE MvPerCta.Conta = PlanoConta.Conta AND TipoConta = 3 AND FilialEmpresa = ? AND (Exercicio = ? Or Exercicio =?) AND MvPerCta.Conta >= ? AND MvPerCta.Conta <= ? GROUP BY Exercicio", _
            iExercicio1, dValor, iFilialEmpresa, iExercicio, iExercicio - 1, sContaInicial, sContaFinal)
        If lErro <> AD_SQL_SUCESSO Then gError 178918
    
        Select Case iFilialEmpresa
            Case EMPRESA_TODA
                If giContabGerencial <> 0 Then
                    lErro = Comando_Executar(alComando(1), "SELECT SUM(Valor) From Lancamentos WHERE FilialEmpresa < ? AND Exercicio = ? AND Origem IN ('APE','EAE') AND Conta >= ? AND Conta <= ?", _
                        dValorAnt, Abs(giFilialAuxiliar), iExercicio - 1, sContaInicial, sContaFinal)
                Else
                    lErro = Comando_Executar(alComando(1), "SELECT SUM(Valor) From Lancamentos WHERE Exercicio = ? AND Origem IN ('APE','EAE') AND Conta >= ? AND Conta <= ?", _
                        dValorAnt, iExercicio - 1, sContaInicial, sContaFinal)
                End If
                
            Case Abs(giFilialAuxiliar)
                lErro = Comando_Executar(alComando(1), "SELECT SUM(Valor) From Lancamentos WHERE FilialEmpresa > ? AND Exercicio = ? AND Origem IN ('APE','EAE') AND Conta >= ? AND Conta <= ?", _
                    dValorAnt, Abs(giFilialAuxiliar), iExercicio - 1, sContaInicial, sContaFinal)
            
            Case Else
                lErro = Comando_Executar(alComando(1), "SELECT SUM(Valor) From Lancamentos WHERE FilialEmpresa = ? AND Exercicio = ? AND Origem IN ('APE','EAE') AND Conta >= ? AND Conta <= ?", _
                    dValorAnt, iFilialEmpresa, iExercicio - 1, sContaInicial, sContaFinal)
        End Select
        If lErro <> AD_SQL_SUCESSO Then gError 207356
    
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then gError 207357
    
        Select Case iFilialEmpresa
            Case EMPRESA_TODA
                If giContabGerencial <> 0 Then
                    lErro = Comando_Executar(alComando(2), "SELECT SUM(Valor) From Lancamentos WHERE FilialEmpresa < ? AND Exercicio = ? AND Origem IN ('APE','EAE') AND Conta >= ? AND Conta <= ?", _
                        dValorAtual, Abs(giFilialAuxiliar), iExercicio, sContaInicial, sContaFinal)
                Else
                    lErro = Comando_Executar(alComando(2), "SELECT SUM(Valor) From Lancamentos WHERE Exercicio = ? AND Origem IN ('APE','EAE') AND Conta >= ? AND Conta <= ?", _
                        dValorAtual, iExercicio, sContaInicial, sContaFinal)
                End If
                
            Case Abs(giFilialAuxiliar)
                lErro = Comando_Executar(alComando(2), "SELECT SUM(Valor) From Lancamentos WHERE FilialEmpresa > ? AND Exercicio = ? AND Origem IN ('APE','EAE') AND Conta >= ? AND Conta <= ?", _
                    dValorAtual, Abs(giFilialAuxiliar), iExercicio, sContaInicial, sContaFinal)
            
            Case Else
                lErro = Comando_Executar(alComando(2), "SELECT SUM(Valor) From Lancamentos WHERE FilialEmpresa = ? AND Exercicio = ? AND Origem IN ('APE','EAE') AND Conta >= ? AND Conta <= ?", _
                    dValorAtual, iFilialEmpresa, iExercicio, sContaInicial, sContaFinal)
        End Select
        If lErro <> AD_SQL_SUCESSO Then gError 207358
    
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then gError 207359
    
        
    ElseIf Len(Trim(sContaInicial)) = 0 And Len(Trim(sCclInicial)) > 0 Then
    
        lErro = Comando_Executar(lComando1, "SELECT Exercicio, SUM(SldIni+Cre01+Cre02+Cre03+Cre04+Cre05+Cre06+Cre07+Cre08+Cre09+Cre10+Cre11+Cre12-Deb01-Deb02-Deb03-Deb04-Deb05-Deb06-Deb07-Deb08-Deb09-Deb10-Deb11-Deb12) FROM MvPerCcl, Ccl WHERE MvPerCcl.Ccl = Ccl.Ccl AND TipoCcl = 3 AND FilialEmpresa = ? AND (Exercicio = ? Or Exercicio =?) AND MvPerCcl.Ccl >= ? AND MvPerCcl.Ccl <= ? AND MvPerCcl.Conta <> '' GROUP BY Exercicio", _
            iExercicio1, dValor, iFilialEmpresa, iExercicio, iExercicio - 1, sCclInicial, sCclFinal)
        If lErro <> AD_SQL_SUCESSO Then gError 178919
    
        Select Case iFilialEmpresa
            Case EMPRESA_TODA
                If giContabGerencial <> 0 Then
                    lErro = Comando_Executar(alComando(1), "SELECT SUM(Valor) From Lancamentos WHERE FilialEmpresa < ? AND Exercicio = ? AND Origem IN ('APE','EAE') AND Conta >= ? AND Conta <= ?", _
                        dValorAnt, Abs(giFilialAuxiliar), iExercicio - 1, sContaInicial, sContaFinal)
                Else
                lErro = Comando_Executar(alComando(1), "SELECT SUM(Valor) From Lancamentos WHERE Exercicio = ? AND Origem IN ('APE','EAE') AND Ccl >= ? AND Ccl <= ? ", _
                    dValorAnt, iExercicio - 1, sCclInicial, sCclFinal)
                End If
                
            Case Abs(giFilialAuxiliar)
                lErro = Comando_Executar(alComando(1), "SELECT SUM(Valor) From Lancamentos WHERE FilialEmpresa > ? AND Exercicio = ? AND Origem IN ('APE','EAE') AND Ccl >= ? AND Ccl <= ? ", _
                    dValorAnt, Abs(giFilialAuxiliar), iExercicio - 1, sCclInicial, sCclFinal)
            
            Case Else
                lErro = Comando_Executar(alComando(1), "SELECT SUM(Valor) From Lancamentos WHERE FilialEmpresa = ? AND Exercicio = ? AND Origem IN ('APE','EAE') AND Ccl >= ? AND Ccl <= ? ", _
                    dValorAnt, iFilialEmpresa, iExercicio - 1, sCclInicial, sCclFinal)
        End Select
        If lErro <> AD_SQL_SUCESSO Then gError 207360
    
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then gError 207361
    
        Select Case iFilialEmpresa
            Case EMPRESA_TODA
                If giContabGerencial <> 0 Then
                    lErro = Comando_Executar(alComando(2), "SELECT SUM(Valor) From Lancamentos WHERE FilialEmpresa < ? AND Exercicio = ? AND Origem IN ('APE','EAE') AND Conta >= ? AND Conta <= ?", _
                        dValorAtual, Abs(giFilialAuxiliar), iExercicio, sContaInicial, sContaFinal)
                Else
                lErro = Comando_Executar(alComando(2), "SELECT SUM(Valor) From Lancamentos WHERE Exercicio = ? AND Origem IN ('APE','EAE') AND Ccl >= ? AND Ccl <= ? ", _
                    dValorAtual, iExercicio, sCclInicial, sCclFinal)
                End If
                
            Case Abs(giFilialAuxiliar)
                lErro = Comando_Executar(alComando(2), "SELECT SUM(Valor) From Lancamentos WHERE FilialEmpresa > ? AND Exercicio = ? AND Origem IN ('APE','EAE') AND Ccl >= ? AND Ccl <= ? ", _
                    dValorAtual, Abs(giFilialAuxiliar), iExercicio, sCclInicial, sCclFinal)
            
            Case Else
                lErro = Comando_Executar(alComando(2), "SELECT SUM(Valor) From Lancamentos WHERE FilialEmpresa = ? AND Exercicio = ? AND Origem IN ('APE','EAE') AND Ccl >= ? AND Ccl <= ? ", _
                    dValorAtual, iFilialEmpresa, iExercicio, sCclInicial, sCclFinal)
        End Select
        If lErro <> AD_SQL_SUCESSO Then gError 207362
    
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then gError 207363
    
    
    ElseIf Len(Trim(sContaInicial)) > 0 And Len(Trim(sCclInicial)) > 0 Then
    
        lErro = Comando_Executar(lComando1, "SELECT Exercicio, SUM(SldIni+Cre01+Cre02+Cre03+Cre04+Cre05+Cre06+Cre07+Cre08+Cre09+Cre10+Cre11+Cre12-Deb01-Deb02-Deb03-Deb04-Deb05-Deb06-Deb07-Deb08-Deb09-Deb10-Deb11-Deb12) FROM MvPerCcl, Ccl, PlanoConta WHERE MvPerCcl.Ccl = Ccl.Ccl AND MvPerCcl.Conta = PlanoConta.Conta AND TipoCcl = 3 AND TipoConta = 3 AND FilialEmpresa = ? AND (Exercicio = ? Or Exercicio =?) AND MvPerCcl.Ccl >= ? AND MvPerCcl.Ccl <= ? AND MvPerCcl.Conta >= ? AND MvPerCcl.Conta <= ? GROUP BY Exercicio", _
            iExercicio1, dValor, iFilialEmpresa, iExercicio, iExercicio - 1, sCclInicial, sCclFinal, sContaInicial, sContaFinal)
        If lErro <> AD_SQL_SUCESSO Then gError 178920
        
        Select Case iFilialEmpresa
            Case EMPRESA_TODA
                If giContabGerencial <> 0 Then
                    lErro = Comando_Executar(alComando(1), "SELECT SUM(Valor) From Lancamentos WHERE FilialEmpresa < ? AND Exercicio = ? AND Origem IN ('APE','EAE') AND Ccl >= ? AND Ccl <= ? AND Conta >= ? AND Conta <= ?", _
                        dValorAnt, Abs(giFilialAuxiliar), iExercicio - 1, sCclInicial, sCclFinal, sContaInicial, sContaFinal)
                Else
                    lErro = Comando_Executar(alComando(1), "SELECT SUM(Valor) From Lancamentos WHERE Exercicio = ? AND Origem IN ('APE','EAE') AND Ccl >= ? AND Ccl <= ? AND Conta >= ? AND Conta <= ?", _
                        dValorAnt, iExercicio - 1, sCclInicial, sCclFinal, sContaInicial, sContaFinal)
                End If
                
            Case Abs(giFilialAuxiliar)
                lErro = Comando_Executar(alComando(1), "SELECT SUM(Valor) From Lancamentos WHERE FilialEmpresa > ? AND Exercicio = ? AND Origem IN ('APE','EAE') AND Ccl >= ? AND Ccl <= ? AND Conta >= ? AND Conta <= ?", _
                    dValorAnt, Abs(giFilialAuxiliar), iExercicio - 1, sCclInicial, sCclFinal, sContaInicial, sContaFinal)
            
            Case Else
                lErro = Comando_Executar(alComando(1), "SELECT SUM(Valor) From Lancamentos WHERE FilialEmpresa = ? AND Exercicio = ? AND Origem IN ('APE','EAE') AND Ccl >= ? AND Ccl <= ? AND Conta >= ? AND Conta <= ?", _
                    dValorAnt, iFilialEmpresa, iExercicio - 1, sCclInicial, sCclFinal, sContaInicial, sContaFinal)
        End Select
        If lErro <> AD_SQL_SUCESSO Then gError 207364
    
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then gError 207365
    
        Select Case iFilialEmpresa
            Case EMPRESA_TODA
                If giContabGerencial <> 0 Then
                    lErro = Comando_Executar(alComando(2), "SELECT SUM(Valor) From Lancamentos WHERE FilialEmpresa < ? AND Exercicio = ? AND Origem IN ('APE','EAE') AND Ccl >= ? AND Ccl <= ? AND Conta >= ? AND Conta <= ?", _
                        dValorAtual, Abs(giFilialAuxiliar), iExercicio, sCclInicial, sCclFinal, sContaInicial, sContaFinal)
                Else
                    lErro = Comando_Executar(alComando(2), "SELECT SUM(Valor) From Lancamentos WHERE Exercicio = ? AND Origem IN ('APE','EAE') AND Ccl >= ? AND Ccl <= ? AND Conta >= ? AND Conta <= ?", _
                        dValorAtual, iExercicio, sCclInicial, sCclFinal, sContaInicial, sContaFinal)
                End If
                
            Case Abs(giFilialAuxiliar)
                lErro = Comando_Executar(alComando(2), "SELECT SUM(Valor) From Lancamentos WHERE FilialEmpresa > ? AND Exercicio = ? AND Origem IN ('APE','EAE') AND Ccl >= ? AND Ccl <= ? AND Conta >= ? AND Conta <= ?", _
                    dValorAtual, Abs(giFilialAuxiliar), iExercicio, sCclInicial, sCclFinal, sContaInicial, sContaFinal)
            
            Case Else
                lErro = Comando_Executar(alComando(2), "SELECT SUM(Valor) From Lancamentos WHERE FilialEmpresa = ? AND Exercicio = ? AND Origem IN ('APE','EAE') AND Ccl >= ? AND Ccl <= ? AND Conta >= ? AND Conta <= ?", _
                    dValorAtual, iFilialEmpresa, iExercicio, sCclInicial, sCclFinal, sContaInicial, sContaFinal)
        End Select
        If lErro <> AD_SQL_SUCESSO Then gError 207366
    
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then gError 207367
        
    
    End If
    
    
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then gError 178921

    Do While lErro = AD_SQL_SUCESSO
    
        If iExercicio = iExercicio1 Then
            dValorExercAtual = dValor - dValorAtual
        Else
            dValorExercAnt = dValor - dValorAnt
        End If
        
        lErro = Comando_BuscarProximo(lComando1)
        If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then gError 178922
        
    Loop
        
    'Libera comando
    If lComando = 0 Then Call Comando_Fechar(lComando1)
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    MvPerCcl_Calcula_Valor = SUCESSO
    
    Exit Function
      
Erro_MvPerCcl_Calcula_Valor:

    MvPerCcl_Calcula_Valor = gErr
    
    Select Case gErr
    
        Case 178918 To 178922
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVPERCTA", gErr)
        
        Case 178923
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 207356 To 207367
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LANCAMENTO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178924)
        
    End Select
    
    'Libera comando
    If lComando = 0 Then Call Comando_Fechar(lComando1)

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function
    
End Function


Function MvPerCcl_Calcula_Valor_Periodo(ByVal iFilialEmpresa As Integer, ByVal iExercicio As Integer, ByVal iPeriodo, ByVal sContaInicial As String, ByVal sContaFinal As String, dValorPerAnt As Double, dValorPerAtual As Double, dValorAcumPerAnt As Double, dValorAcumPerAtual As Double, ByVal sCclInicial As String, ByVal sCclFinal As String, Optional ByVal lComando As Long = 0) As Long
'Calcula Valor pela soma de SldIni com os valores de crédito e débito do Exercicio/Periodo em questao

Dim lErro As Long
Dim lComando1 As Long
Dim dSldIni As Double
Dim dValorPeriodo As Double
Dim dValorAcumPeriodo As Double
Dim iIndice As Integer
Dim sSQL As String
Dim iExercicio1 As Integer
Dim iPeriodo1 As Integer
    
On Error GoTo Erro_MvPerCcl_Calcula_Valor_Periodo

    If lComando = 0 Then
        lComando1 = Comando_Abrir()
        If lComando1 = 0 Then gError 178925
    Else
        lComando1 = lComando
    End If
    
    dValorPerAtual = 0
    dValorPerAnt = 0
    dValorAcumPerAtual = 0
    dValorAcumPerAnt = 0
    
    sSQL = "SUM(SldIni + "
    
    For iPeriodo1 = 1 To iPeriodo
        sSQL = sSQL + " Cre" + Format(iPeriodo1, "00") + " - Deb" + Format(iPeriodo1, "00") + " + "
    Next
    
    sSQL = left(sSQL, Len(sSQL) - 3) + ")"
    
    If Len(Trim(sContaInicial)) > 0 And Len(Trim(sCclInicial)) = 0 Then
    
        lErro = Comando_Executar(lComando1, "SELECT Exercicio, SUM(Cre" + Format(iPeriodo, "00") + " - Deb" + Format(iPeriodo, "00") + "), " + sSQL + " FROM MvPerCta, PlanoConta WHERE MvPerCta.Conta = PlanoConta.Conta AND TipoConta = 3 AND FilialEmpresa = ? AND (Exercicio = ? OR Exercicio = ?) AND MvPerCta.Conta >= ? AND MvPerCta.Conta <= ? GROUP BY Exercicio", _
            iExercicio1, dValorPeriodo, dValorAcumPeriodo, iFilialEmpresa, iExercicio, iExercicio - 1, sContaInicial, sContaFinal)
        If lErro <> AD_SQL_SUCESSO Then gError 178926

    ElseIf Len(Trim(sContaInicial)) = 0 And Len(Trim(sCclInicial)) > 0 Then
    
        lErro = Comando_Executar(lComando1, "SELECT Exercicio, SUM(Cre" + Format(iPeriodo, "00") + " - Deb" + Format(iPeriodo, "00") + "), " + sSQL + " FROM MvPerCcl, Ccl WHERE MvPerCcl.Ccl = Ccl.Ccl AND TipoCcl = 3 AND FilialEmpresa = ? AND (Exercicio = ? OR Exercicio = ?) AND MvPerCcl.Ccl >= ? AND MvPerCcl.Ccl <= ? ANd MvPerCcl.Conta <> '' GROUP BY Exercicio", _
            iExercicio1, dValorPeriodo, dValorAcumPeriodo, iFilialEmpresa, iExercicio, iExercicio - 1, sCclInicial, sCclFinal)
        If lErro <> AD_SQL_SUCESSO Then gError 178927

    ElseIf Len(Trim(sContaInicial)) > 0 And Len(Trim(sCclInicial)) > 0 Then

        lErro = Comando_Executar(lComando1, "SELECT Exercicio, SUM(Cre" + Format(iPeriodo, "00") + " - Deb" + Format(iPeriodo, "00") + "), " + sSQL + " FROM MvPerCcl, Ccl, PlanoConta WHERE MvPerCcl.Ccl = Ccl.Ccl AND MvPerCcl.Conta = PlanoConta.Conta AND TipoCcl = 3 AND TipoConta = 3 AND FilialEmpresa = ? AND (Exercicio = ? OR Exercicio = ?) AND MvPerCcl.Ccl >= ? AND MvPerCcl.Ccl <= ? AND MvPerCcl.Conta >= ? AND MvPerCcl.Conta <= ? GROUP BY Exercicio", _
            iExercicio1, dValorPeriodo, dValorAcumPeriodo, iFilialEmpresa, iExercicio, iExercicio - 1, sCclInicial, sCclFinal, sContaInicial, sContaFinal)
        If lErro <> AD_SQL_SUCESSO Then gError 178928

    End If

    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then gError 178929

    Do While lErro = AD_SQL_SUCESSO

        If iExercicio = iExercicio1 Then
            dValorPerAtual = dValorPeriodo
            dValorAcumPerAtual = dValorAcumPeriodo
        Else
            dValorPerAnt = dValorPeriodo
            dValorAcumPerAnt = dValorAcumPeriodo
        End If
        
        lErro = Comando_BuscarProximo(lComando1)
        If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then gError 178930

    Loop

    'Libera comando
    If lComando = 0 Then Call Comando_Fechar(lComando1)
    
    MvPerCcl_Calcula_Valor_Periodo = SUCESSO
    
    Exit Function
      
Erro_MvPerCcl_Calcula_Valor_Periodo:

    MvPerCcl_Calcula_Valor_Periodo = gErr
    
    Select Case gErr
    
        Case 178925
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 178926 To 178930
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVPERCTA", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178931)
        
    End Select
    
    'Libera comando
    If lComando = 0 Then Call Comando_Fechar(lComando1)

    Exit Function
    
End Function

Public Function DRE_Move_Dados_Formato_Excel(ByVal objPlanilha As ClassPlanilhaExcel, ByVal sModelo As String, ByVal sTipoRelatorio As String, ByVal sDiretorio As String) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(0 To 0) As Long
Dim sTitulo As String
Dim sRelatorio As String
Dim iTipo As Integer
Dim dValor As Double
Dim dValorExercAnt As Double
Dim dValorPerAcumAtual As Double
Dim dValorPerAcumAnt As Double
Dim objColunas As ClassColunasExcel
Dim objCelulas As ClassCelulasExcel
Dim dLarguraColuna As Double
Dim iNumColunas As Integer
Dim sNome As String
Dim iNivel As Integer
Dim sIdentacao As String

On Error GoTo Erro_DRE_Move_Dados_Formato_Excel

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 187133
    Next
    
    'Guarda o nome que será exibido para a planilha
    'objPlanilha.sNomePlanilha = sDiretorio & "DRE_" & sModelo & Format(gdtDataHoje, "yyyymmdd")
    objPlanilha.sNomePlanilha = sModelo
    objPlanilha.iModoGravacao = MARCADO
     
    If sTipoRelatorio = RELDRE Then
        iNumColunas = 3
        objPlanilha.sNomeDiretorioArq = sDiretorio & "DRE_" & sModelo & Format(gdtDataHoje, "yyyymmdd")
    Else
        iNumColunas = 5
        objPlanilha.sNomeDiretorioArq = sDiretorio & "DRP_" & sModelo & Format(gdtDataHoje, "yyyymmdd")
    End If
    
    'Obtém a largura da coluna no padrão do Excel
    lErro = CF("Excel_Obtem_Largura_Coluna", 2000, dLarguraColuna)
    If lErro <> SUCESSO Then gError 187134
       
    For iIndice = 1 To iNumColunas
    
        'Instancia um novo obj para armazenar os dados da coluna
        Set objColunas = New ClassColunasExcel
        
        If iIndice = 1 Then
            'Guarda no objeto a largura dessa coluna
            objColunas.dLarguraColuna = dLarguraColuna * 3
        Else
            'Guarda no objeto a largura dessa coluna
            objColunas.dLarguraColuna = dLarguraColuna
        End If
        
        'Adiciona a coluna à coleção de colunas
        objPlanilha.colColunas.Add objColunas
    
        'Instancia um novo objeto para armazenar os dados da primeira célula da coluna em questão
        Set objCelulas = New ClassCelulasExcel
        
        Select Case iIndice
        
            Case 1
                sNome = " "
            Case 2
                If sTipoRelatorio = RELDRE Then
                    sNome = "EXERCÍCIO ATUAL"
                Else
                    sNome = "PER.ATUAL"
                End If
            
            Case 3
                If sTipoRelatorio = RELDRE Then
                    sNome = "EXERCÍCIO ANTERIOR"
                Else
                    sNome = "PER.ANTERIOR ACUM"
                End If
            
            Case 4
                sNome = "PER.ATUAL ACUM"
            
            Case 5
                sNome = "PER.ANTERIOR"
        
        End Select

        'Guarda o título da coluna em questão que será exibido na primeira célula
        objCelulas.vValor = sNome
        objCelulas.bFonteNegrito = True
         
        'Adiciona a célula à coleção de células
        objColunas.colCelulas.Add objCelulas

    Next
     
    sRelatorio = String(255, 0)
    sTitulo = String(255, 0)

    lErro = Comando_Executar(alComando(0), "SELECT Nivel, Titulo, Relatorio, Tipo, Valor,ValorExercAnt,ValorPerAcumAtual,ValorPerAcumAnt FROM RelDRE WHERE Imprime = 1 AND Modelo = ? ORDER BY Posicao", iNivel, sTitulo, sRelatorio, iTipo, dValor, dValorExercAnt, dValorPerAcumAtual, dValorPerAcumAnt, sModelo)
    If lErro <> AD_SQL_SUCESSO Then gError 187135

    lErro = Comando_BuscarProximo(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 187136

    Do While lErro <> AD_SQL_SEM_DADOS
       
        If sTipoRelatorio <> RELDRE Or (sTipoRelatorio = RELDRE And sRelatorio = RELDRE) Then
        
            sIdentacao = String((iNivel - 1) * 3, Asc(" "))
    
            For iIndice = 1 To iNumColunas
    
                'aponta para a coluna em questão
                Set objColunas = objPlanilha.colColunas(iIndice)
                
                'Instancia um novo objeto para armazenar os dados da célula
                Set objCelulas = New ClassCelulasExcel
                
                Select Case iIndice
                
                    Case 1
                        objCelulas.vValor = sIdentacao & sTitulo
                        
                        If iNivel = 1 Then
                            objCelulas.bFonteNegrito = True
                        End If
                    
                    Case 2
                        If iTipo <> 2 Then
                            objCelulas.vValor = dValor
                        Else
                            objCelulas.vValor = Null
                        End If
                        objCelulas.sNumberFormat = "#,##0.00"
                    
                    Case 3
                        If iTipo <> 2 Then
                            objCelulas.vValor = dValorExercAnt
                        Else
                            objCelulas.vValor = Null
                        End If
                        objCelulas.sNumberFormat = "#,##0.00"
                    
                    Case 4
                        If iTipo <> 2 Then
                            objCelulas.vValor = dValorPerAcumAtual
                        Else
                            objCelulas.vValor = Null
                        End If
                        objCelulas.sNumberFormat = "#,##0.00"
                    
                    Case 5
                        If iTipo <> 2 Then
                            objCelulas.vValor = dValorPerAcumAnt
                        Else
                            objCelulas.vValor = Null
                        End If
                        objCelulas.sNumberFormat = "#,##0.00"
                        
                End Select
                                   
                'adiciona a célula à coleção de células da coluna em questão
                objColunas.colCelulas.Add objCelulas
                
            Next
            
        End If

        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 187137

    Loop
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
        
    DRE_Move_Dados_Formato_Excel = SUCESSO
    
    Exit Function
    
Erro_DRE_Move_Dados_Formato_Excel:
    
    DRE_Move_Dados_Formato_Excel = gErr
    
    Select Case gErr

        Case 187133
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 187134
        
        Case 187135, 187136, 187137
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELDRE", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 187138)

    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function
    
End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function PlanoContaRefModelo_Le(ByVal objPlanoContaRefModelo As ClassPlanoContaRefModelo) As Long

Dim lErro As Long
Dim lComando As Long
Dim tPlanoContaRefModelo As typePlanoContaRefModelo

On Error GoTo Erro_PlanoContaRefModelo_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 200871

    'Alocação de espaço no buffer
    tPlanoContaRefModelo.sDescricao = String(STRING_PLANOCONTAREFMODELO_DESCRICAO, 0)
    tPlanoContaRefModelo.sUsuarioCad = String(STRING_USUARIO_CODIGO, 0)
    tPlanoContaRefModelo.sUsuarioAlt = String(STRING_USUARIO_CODIGO, 0)

    'Le a tabelaPlanoContaRefModelo
    lErro = Comando_Executar(lComando, "SELECT Codigo, Descricao, Oficial, AnoVigencia, UsuarioCad, DataCad, " & _
                "HoraCad, UsuarioAlt, DataAlt, HoraAlt, Tipo FROM PlanoContaRefModelo WHERE Codigo= ? ", _
                tPlanoContaRefModelo.lCodigo, tPlanoContaRefModelo.sDescricao, tPlanoContaRefModelo.iOficial, tPlanoContaRefModelo.iAnoVigencia, _
                tPlanoContaRefModelo.sUsuarioCad, tPlanoContaRefModelo.dtDataCad, tPlanoContaRefModelo.dHoraCad, tPlanoContaRefModelo.sUsuarioAlt, tPlanoContaRefModelo.dtDataAlt, _
                tPlanoContaRefModelo.dHoraAlt, tPlanoContaRefModelo.iTipo, _
                objPlanoContaRefModelo.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 200872

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200873

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objPlanoContaRefModelo.lCodigo = tPlanoContaRefModelo.lCodigo
    objPlanoContaRefModelo.sDescricao = tPlanoContaRefModelo.sDescricao
    objPlanoContaRefModelo.iOficial = tPlanoContaRefModelo.iOficial
    objPlanoContaRefModelo.iAnoVigencia = tPlanoContaRefModelo.iAnoVigencia
    objPlanoContaRefModelo.sUsuarioCad = tPlanoContaRefModelo.sUsuarioCad
    objPlanoContaRefModelo.dtDataCad = tPlanoContaRefModelo.dtDataCad
    objPlanoContaRefModelo.dHoraCad = tPlanoContaRefModelo.dHoraCad
    objPlanoContaRefModelo.sUsuarioAlt = tPlanoContaRefModelo.sUsuarioAlt
    objPlanoContaRefModelo.dtDataAlt = tPlanoContaRefModelo.dtDataAlt
    objPlanoContaRefModelo.dHoraAlt = tPlanoContaRefModelo.dHoraAlt
    
    objPlanoContaRefModelo.iTipo = tPlanoContaRefModelo.iTipo
    
    lErro = CF("PlanoContaRef_Le_Modelo", objPlanoContaRefModelo)
    If lErro <> SUCESSO Then gError 200874

    'Fecha Comando
    Call Comando_Fechar(lComando)

    PlanoContaRefModelo_Le = SUCESSO

    Exit Function

Erro_PlanoContaRefModelo_Le:

    PlanoContaRefModelo_Le = gErr

    Select Case gErr

        Case 200871
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 200872, 200873
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTAREFMODELO", gErr)
            
        Case 200874

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200875)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function PlanoContaRef_Le(ByVal objPlanoContaRef As ClassPlanoContaRef) As Long

Dim lErro As Long
Dim lComando As Long
Dim tPlanoContaRef As typePlanoContaRef

On Error GoTo Erro_PlanoContaRef_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 200876

    'Alocação de espaço no buffer
    tPlanoContaRef.sConta = String(STRING_CONTA, 0)
    tPlanoContaRef.sContaImp = String(STRING_CONTA, 0)
    tPlanoContaRef.sContaPai = String(STRING_CONTA, 0)
    tPlanoContaRef.sDescricao = String(STRING_PLANOCONTAREF_DESCRICAO, 0)
    tPlanoContaRef.sOrientacao = String(STRING_PLANOCONTAREF_ORIENTACAO, 0)

    'Le a tabelaPlanoContaRef
    lErro = Comando_Executar(lComando, "SELECT CodigoModelo, Conta, Descricao, Orientacao, Tipo, ContaImp, TipoImp, ValidadeDe, ValidadeAte, Natureza, ContaPai FROM PlanoContaRef WHERE CodigoModelo= ? AND Conta = ? ", _
                tPlanoContaRef.lCodigoModelo, tPlanoContaRef.sConta, tPlanoContaRef.sDescricao, tPlanoContaRef.sOrientacao, _
                tPlanoContaRef.iTipo, tPlanoContaRef.sContaImp, tPlanoContaRef.iTipoImp, tPlanoContaRef.dtValidadeDe, tPlanoContaRef.dtValidadeAte, tPlanoContaRef.iNatureza, tPlanoContaRef.sContaPai, objPlanoContaRef.lCodigoModelo, objPlanoContaRef.sConta)
    If lErro <> AD_SQL_SUCESSO Then gError 200877

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200878

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objPlanoContaRef.lCodigoModelo = tPlanoContaRef.lCodigoModelo
    objPlanoContaRef.sConta = tPlanoContaRef.sConta
    objPlanoContaRef.sDescricao = tPlanoContaRef.sDescricao
    objPlanoContaRef.sOrientacao = tPlanoContaRef.sOrientacao
    objPlanoContaRef.iTipo = tPlanoContaRef.iTipo
    objPlanoContaRef.sContaImp = tPlanoContaRef.sContaImp
    objPlanoContaRef.iTipoImp = tPlanoContaRef.iTipoImp
    objPlanoContaRef.dtValidadeDe = tPlanoContaRef.dtValidadeDe
    objPlanoContaRef.dtValidadeAte = tPlanoContaRef.dtValidadeAte
    objPlanoContaRef.iNatureza = tPlanoContaRef.iNatureza
    objPlanoContaRef.sContaPai = tPlanoContaRef.sContaPai

    lErro = CF("PlanoContaRefConfig_Le", objPlanoContaRef)
    If lErro <> SUCESSO Then gError 200879

    'Fecha Comando
    Call Comando_Fechar(lComando)

    PlanoContaRef_Le = SUCESSO

    Exit Function

Erro_PlanoContaRef_Le:

    PlanoContaRef_Le = gErr

    Select Case gErr

        Case 200876
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 200877, 200878
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTAREF", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200880)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function PlanoContaRef_Le_Modelo(ByVal objPlanoContaRefModelo As ClassPlanoContaRefModelo) As Long

Dim lErro As Long
Dim lComando As Long
Dim tPlanoContaRef As typePlanoContaRef
Dim objPlanoContaRef As ClassPlanoContaRef

On Error GoTo Erro_PlanoContaRef_Le_Modelo

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 200881

    'Alocação de espaço no buffer
    tPlanoContaRef.sConta = String(STRING_CONTA, 0)
    tPlanoContaRef.sConta0 = String(STRING_CONTA, 0)
    tPlanoContaRef.sContaImp = String(STRING_CONTA, 0)
    tPlanoContaRef.sContaPai = String(STRING_CONTA, 0)
    tPlanoContaRef.sDescricao = String(STRING_PLANOCONTAREF_DESCRICAO, 0)
    tPlanoContaRef.sOrientacao = String(STRING_PLANOCONTAREF_ORIENTACAO, 0)

    'Le a tabelaPlanoContaRef
    lErro = Comando_Executar(lComando, "SELECT CodigoModelo, Conta, Descricao, Orientacao, Tipo, ContaImp, TipoImp, ValidadeDe, ValidadeAte, Conta0, Natureza, ContaPai FROM PlanoContaRef WHERE CodigoModelo= ? ORDER BY Conta ", _
                tPlanoContaRef.lCodigoModelo, tPlanoContaRef.sConta, tPlanoContaRef.sDescricao, tPlanoContaRef.sOrientacao, _
                tPlanoContaRef.iTipo, tPlanoContaRef.sContaImp, tPlanoContaRef.iTipoImp, tPlanoContaRef.dtValidadeDe, tPlanoContaRef.dtValidadeAte, tPlanoContaRef.sConta0, tPlanoContaRef.iNatureza, tPlanoContaRef.sContaPai, objPlanoContaRefModelo.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 200882

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200883

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objPlanoContaRef = New ClassPlanoContaRef

        objPlanoContaRef.lCodigoModelo = tPlanoContaRef.lCodigoModelo
        objPlanoContaRef.sConta = tPlanoContaRef.sConta
        objPlanoContaRef.sConta0 = tPlanoContaRef.sConta0
        objPlanoContaRef.sDescricao = tPlanoContaRef.sDescricao
        objPlanoContaRef.sOrientacao = tPlanoContaRef.sOrientacao
        objPlanoContaRef.iTipo = tPlanoContaRef.iTipo
        objPlanoContaRef.sContaImp = tPlanoContaRef.sContaImp
        objPlanoContaRef.iTipoImp = tPlanoContaRef.iTipoImp
        objPlanoContaRef.dtValidadeDe = tPlanoContaRef.dtValidadeDe
        objPlanoContaRef.dtValidadeAte = tPlanoContaRef.dtValidadeAte
    
        objPlanoContaRef.sContaPai = tPlanoContaRef.sContaPai
        objPlanoContaRef.iNatureza = tPlanoContaRef.iNatureza
    
        lErro = CF("PlanoContaRefConfig_Le", objPlanoContaRef)
        If lErro <> SUCESSO Then gError 200884
        
        objPlanoContaRefModelo.colContas.Add objPlanoContaRef, objPlanoContaRef.sConta

        'Busca Próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200885

    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    PlanoContaRef_Le_Modelo = SUCESSO

    Exit Function

Erro_PlanoContaRef_Le_Modelo:

    PlanoContaRef_Le_Modelo = gErr

    Select Case gErr

        Case 200881
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 200882, 200883, 200885
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTAREF", gErr)

        Case 200884

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200885)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function PlanoContaRefConfig_Le(ByVal objPlanoContaRef As ClassPlanoContaRef) As Long

Dim lErro As Long
Dim lComando As Long
Dim tPlanoContaRefConfig As typePlanoContaRefConfig
Dim objPlanoContaRefConfig As ClassPlanoContaRefConfig

On Error GoTo Erro_PlanoContaRefConfig_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 200886

    'Alocação de espaço no buffer
    tPlanoContaRefConfig.sConta = String(STRING_CONTA, 0)
    tPlanoContaRefConfig.sContaInicial = String(STRING_CONTA, 0)
    tPlanoContaRefConfig.sContaFinal = String(STRING_CONTA, 0)
    tPlanoContaRefConfig.sCclInicial = String(STRING_CCL, 0)
    tPlanoContaRefConfig.sCclFinal = String(STRING_CCL, 0)

    'Le a tabelaPlanoContaRefConfig
    lErro = Comando_Executar(lComando, "SELECT CodigoModelo, Conta, Seq, ContaInicial, ContaFinal, CclInicial, " & _
                "CclFinal, Subtrai FROM PlanoContaRefConfig WHERE CodigoModelo= ? AND Conta = ? ORDER By Seq ", _
                tPlanoContaRefConfig.lCodigoModelo, tPlanoContaRefConfig.sConta, tPlanoContaRefConfig.iSeq, tPlanoContaRefConfig.sContaInicial, _
                tPlanoContaRefConfig.sContaFinal, tPlanoContaRefConfig.sCclInicial, tPlanoContaRefConfig.sCclFinal, tPlanoContaRefConfig.iSubtrai, _
                objPlanoContaRef.lCodigoModelo, objPlanoContaRef.sConta)
    If lErro <> AD_SQL_SUCESSO Then gError 200887

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200888

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objPlanoContaRefConfig = New ClassPlanoContaRefConfig

        objPlanoContaRefConfig.lCodigoModelo = tPlanoContaRefConfig.lCodigoModelo
        objPlanoContaRefConfig.sConta = tPlanoContaRefConfig.sConta
        objPlanoContaRefConfig.iSeq = tPlanoContaRefConfig.iSeq
        objPlanoContaRefConfig.sContaInicial = tPlanoContaRefConfig.sContaInicial
        objPlanoContaRefConfig.sContaFinal = tPlanoContaRefConfig.sContaFinal
        objPlanoContaRefConfig.sCclInicial = tPlanoContaRefConfig.sCclInicial
        objPlanoContaRefConfig.sCclFinal = tPlanoContaRefConfig.sCclFinal
        objPlanoContaRefConfig.iSubtrai = tPlanoContaRefConfig.iSubtrai
        
        objPlanoContaRef.colConfig.Add objPlanoContaRefConfig

        'Busca Próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200889

    Loop
    'Fecha Comando
    Call Comando_Fechar(lComando)

    PlanoContaRefConfig_Le = SUCESSO

    Exit Function

Erro_PlanoContaRefConfig_Le:

    PlanoContaRefConfig_Le = gErr

    Select Case gErr

        Case 200886
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 200887, 200888, 200889
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTAREFConfig", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200890)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Lancamentos_Abre_Tela(ByVal iOrigem1 As Integer, ByVal lNumIntDoc1 As Long, ByVal iOrigem2 As Integer, ByVal lNumIntDoc2 As Long) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(1 To 2) As Long
Dim objLanc As ClassLancamento_Detalhe
Dim tLanc As typeLancamento
Dim sTela As String
Dim sFiltro As String

On Error GoTo Erro_Lancamentos_Abre_Tela

    GL_objMDIForm.MousePointer = vbHourglass

    If lNumIntDoc1 = 0 And lNumIntDoc2 = 0 Then gError 202700

    'Abrir comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 202701
    Next
    
    If giFilialAuxiliar <> 0 Then
        If Abs(giFilialAuxiliar) > giFilialEmpresa Then
            sFiltro = " AND FilialEmpresa < " & CStr(Abs(giFilialAuxiliar)) & " "
        Else
            sFiltro = " AND FilialEmpresa > " & CStr(Abs(giFilialAuxiliar)) & " "
        End If
    End If
    
    sTela = "LancamentoConsulta"

    tLanc.sOrigem = String(STRING_ORIGEM, 0)
    
    lErro = Comando_Executar(alComando(1), "SELECT FilialEmpresa, Origem, Exercicio, PeriodoLan, Doc FROM TransacaoCTB AS T, Lancamentos As L WHERE L.Transacao = T.Codigo AND ((OrigemLcto = ? AND NumIntDoc = ?) OR (OrigemLcto = ? AND NumIntDoc = ?))" & sFiltro, _
    tLanc.iFilialEmpresa, tLanc.sOrigem, tLanc.iExercicio, tLanc.iPeriodoLan, tLanc.lDoc, iOrigem1, lNumIntDoc1, iOrigem2, lNumIntDoc2)
    If lErro <> AD_SQL_SUCESSO Then gError 202702

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202703
    
    If lErro = AD_SQL_SEM_DADOS Then
        sTela = "Lancamentos"
        
        lErro = Comando_Executar(alComando(2), "SELECT FilialEmpresa, Origem, Exercicio, PeriodoLan, Doc FROM TransacaoCTB AS T, LanPendente As L WHERE L.Transacao = T.Codigo AND ((OrigemLcto = ? AND NumIntDoc = ?) OR (OrigemLcto = ? AND NumIntDoc = ?))" & sFiltro, _
        tLanc.iFilialEmpresa, tLanc.sOrigem, tLanc.iExercicio, tLanc.iPeriodoLan, tLanc.lDoc, iOrigem1, lNumIntDoc1, iOrigem2, lNumIntDoc2)
        If lErro <> AD_SQL_SUCESSO Then gError 202702
    
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202703
        
    End If

    If lErro = AD_SQL_SEM_DADOS Then gError 202704
    
    Set objLanc = New ClassLancamento_Detalhe
    
    objLanc.iPeriodoLan = tLanc.iPeriodoLan
    objLanc.iFilialEmpresa = tLanc.iFilialEmpresa
    objLanc.sOrigem = tLanc.sOrigem
    objLanc.iExercicio = tLanc.iExercicio
    objLanc.lDoc = tLanc.lDoc
               
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Call Chama_Tela(sTela, objLanc)
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Lancamentos_Abre_Tela = SUCESSO

    Exit Function

Erro_Lancamentos_Abre_Tela:

    GL_objMDIForm.MousePointer = vbDefault

    Lancamentos_Abre_Tela = gErr

    Select Case gErr
    
        Case 202700 'ERRO_NENHUM_DOC_SELECIONADO
            Call Rotina_Erro(vbOKOnly, "ERRO_NENHUM_DOC_SELECIONADO", gErr)

        Case 202701
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 202702, 202703
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LANCAMENTOS3", gErr)

        Case 202704
            Call Rotina_Erro(vbOKOnly, "ERRO_LANCAMENTO_INEXISTENTE2", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202705)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Public Function Lcto_Le_Ccl_Cta_Result_Doc(ByVal iOrigemLcto As Integer, ByVal lNumIntDoc As Long, sCclRet As String) As Long
'Lê o centro de custo do lançamento da conta de resultado do documento passado

Dim lErro As Long
Dim lComando As Long
Dim sCcl As String

On Error GoTo Erro_Lcto_Le_Ccl_Cta_Result_Doc

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 213471

    'Alocação de espaço no buffer
    sCcl = String(STRING_CCL, 0)

    lErro = Comando_Executar(lComando, "SELECT Ccl FROM LancamentosTodos AS L, PlanoConta AS C WHERE L.Conta = C.Conta AND NumIntDoc = ? AND Transacao IN (SELECT Codigo FROM TransacaoCTB WHERE OrigemLcto = ?) AND C.NaturezaSped = 4", sCcl, lNumIntDoc, iOrigemLcto)
    If lErro <> AD_SQL_SUCESSO Then gError 213472

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213473
    
    If lErro = AD_SQL_SUCESSO Then
        sCclRet = sCcl
    Else
        sCclRet = ""
    End If
    
    'Fecha Comando
    Call Comando_Fechar(lComando)

    Lcto_Le_Ccl_Cta_Result_Doc = SUCESSO

    Exit Function

Erro_Lcto_Le_Ccl_Cta_Result_Doc:

    Lcto_Le_Ccl_Cta_Result_Doc = gErr

    Select Case gErr

        Case 213471
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 213472, 213473
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213474)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function TiposPlanoContaRef_Le(ByVal objTiposPlanoContaRef As ClassTiposPlanoContaRef) As Long

Dim lErro As Long
Dim lComando As Long
Dim tTiposPlanoContaRef As typeTiposPlanoContaRef

On Error GoTo Erro_TiposPlanoContaRef_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 213840

    'Alocação de espaço no buffer
    tTiposPlanoContaRef.sDescricao = String(STRING_MAXIMO, 0)

    'Le a tabelaTiposPlanoContaRef
    lErro = Comando_Executar(lComando, "SELECT AnoBase, Tipo, Descricao FROM TiposPlanoContaRef WHERE AnoBase <= ? AND Tipo= ? ORDER BY AnoBase DESC", _
                tTiposPlanoContaRef.iAnoBase, tTiposPlanoContaRef.iTipo, tTiposPlanoContaRef.sDescricao, _
                objTiposPlanoContaRef.iAnoBase, objTiposPlanoContaRef.iTipo)
    If lErro <> AD_SQL_SUCESSO Then gError 213841

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213842

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objTiposPlanoContaRef.iAnoBase = tTiposPlanoContaRef.iAnoBase
    objTiposPlanoContaRef.iTipo = tTiposPlanoContaRef.iTipo
    objTiposPlanoContaRef.sDescricao = tTiposPlanoContaRef.sDescricao

    'Fecha Comando
    Call Comando_Fechar(lComando)

    TiposPlanoContaRef_Le = SUCESSO

    Exit Function

Erro_TiposPlanoContaRef_Le:

    TiposPlanoContaRef_Le = gErr

    Select Case gErr

        Case 213840
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 213841, 213842
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOSPLANOCONTAREF", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213843)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function



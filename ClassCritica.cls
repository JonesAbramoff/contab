VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCritica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const STRING_CONTACONTABIL = 20



Function Conta_Critica(ByVal sConta As String, sContaFormatada As String, ByVal objPlanoConta As ClassPlanoConta, ByVal sSiglaModulo As String) As Long

Dim lErro As Long
Dim iContaPreenchida As Integer

On Error GoTo Erro_Conta_Critica


    If Len(sConta) > 0 Then
    
        lErro = CF("Conta_Formata", sConta, sContaFormatada, iContaPreenchida)
        If lErro Then Error 5705
    
        If iContaPreenchida = CONTA_PREENCHIDA Then
    
            lErro = CF("Conta_SelecionaUma", sContaFormatada, objPlanoConta, sSiglaModulo)
            If lErro <> SUCESSO And lErro <> 6030 Then Error 5699
    
            'Ausencia de Conta no BD
            If lErro = 6030 Then Error 5700
            
            If objPlanoConta.iAtivo <> CONTA_ATIVA Then Error 5856
            
            If objPlanoConta.iTipoConta <> CONTA_ANALITICA Then Error 5857
            
        
        End If
    
    End If
    
    Conta_Critica = SUCESSO
    
    Exit Function

Erro_Conta_Critica:

    Conta_Critica = Err
    
    Select Case Err
    
        Case 5699, 5705
        
        Case 5700
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_CADASTRADA", Err, sConta)
        
        Case 5856
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_INATIVA", Err, sConta)
        
        Case 5857
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_ANALITICA", Err, sConta)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149222)
        
    End Select
    
    Exit Function

End Function


Function Exercicio_Critica(ByVal iExercicio As Integer, iExercicioFechado As Integer) As Long
'Verifica se o exercicio está fechado

Dim iExercicioStatus As Integer
Dim lErro As Long
Dim lComando As Long

On Error GoTo Erro_Exercicio_Critica

    lComando = 0
    
    'Consulta BD para checar STATUS do EXERCICIO de iMes,iAno
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 6005

    lErro = Comando_Executar(lComando, "SELECT Status FROM Exercicios WHERE Exercicio=?", iExercicioStatus, iExercicio)
    If lErro <> AD_SQL_SUCESSO Then Error 6006
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 5437

    If iExercicioStatus = EXERCICIO_FECHADO Then
        iExercicioFechado = 1
    Else
        iExercicioFechado = 0
    End If

    Call Comando_Fechar(lComando)

    Exercicio_Critica = SUCESSO
    
    Exit Function

Erro_Exercicio_Critica:

    Exercicio_Critica = Err

    Select Case Err
    
        Case 6005
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        Case 5437, 6006
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EXERCICIO2", Err, iExercicio)
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149223)
            
    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Lote_Critica_Atualizado(ByVal objLote As ClassLote, iLoteAtualizado As Integer) As Long
'testa se o lote já foi atualizado. Se já foi iLoteAtualizado = LOTE_ATUALIZADO senao retorna 0.

Dim lComando1 As Long
Dim lErro As Long
Dim iLote As Integer

On Error GoTo Erro_Lote_Critica_Atualizado
    
    lComando1 = 0
    
    iLoteAtualizado = 0

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 5429

    'pesquisa no BD o lote em questão
    lErro = Comando_Executar(lComando1, "SELECT Lote FROM Lote WHERE FilialEmpresa =? AND Origem=? AND Exercicio=? AND Periodo=? AND Lote=?", iLote, objLote.iFilialEmpresa, objLote.sOrigem, objLote.iExercicio, objLote.iPeriodo, objLote.iLote)
    If lErro <> AD_SQL_SUCESSO Then Error 5430

    'le o lote, se estiver no BD
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 5431

    If lErro = AD_SQL_SUCESSO Then iLoteAtualizado = LOTE_ATUALIZADO
    
    Call Comando_Fechar(lComando1)

    Lote_Critica_Atualizado = SUCESSO

    Exit Function

Erro_Lote_Critica_Atualizado:

    Lote_Critica_Atualizado = Err

    Select Case Err

        Case 5429
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 5430, 5431
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LOTEPENDENTE", Err, objLote.iFilialEmpresa, objLote.sOrigem, objLote.iExercicio, objLote.iPeriodo, objLote.iLote)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149224)

    End Select

    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Function ContaAnalitica_Critica_Exclusao(ByVal sConta As String) As Long
'verifica se a conta possui lancamentos pendentes, rateios associados, etc.

Dim lErro As Long
Dim iTemLanPendente As Integer

On Error GoTo Erro_ContaAnalitica_Critica_Exclusao

        
    'Verificar se existem lancamentos pendentes
    lErro = CF("Conta_Critica_Possui_LanPendente", sConta, iTemLanPendente)
    If lErro <> SUCESSO Then gError 5669
        
    If iTemLanPendente = 1 Then gError 5670
        
'     'verifica se a conta apreseta vinculo com outras tabelas e views
'    lErro = Conta_Critica_Vinculo_Compras(sConta)
'    If lErro <> SUCESSO Then gError 85206

    'verifica se a conta não tem rateios on-line associados
    lErro = CF("RateioOn_Le_Conta", sConta)
    If lErro <> SUCESSO And lErro <> 5566 Then gError 5581
    
    If lErro = SUCESSO Then gError 5583
    
    'verifica se a conta não tem rateios off-line associados
    lErro = CF("RateioOff_Le_Conta", sConta)
    If lErro <> SUCESSO And lErro <> 5574 Then gError 5582
    
    If lErro = SUCESSO Then gError 5584
    
    ContaAnalitica_Critica_Exclusao = SUCESSO
    
    Exit Function
        
Erro_ContaAnalitica_Critica_Exclusao:

    ContaAnalitica_Critica_Exclusao = gErr
    
    Select Case gErr
    
        Case 5581, 5582, 5669
            
        Case 5583
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CONTA_COM_RATEIOON", gErr)
        
        Case 5584
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CONTA_COM_RATEIOOFF", gErr)
            
        Case 5670
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CONTA_COM_LANC_PEND1", gErr)
        
        Case 85206
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149225)
            
    End Select
        
    Exit Function
    


End Function

Function Conta_Critica_Vinculo_Compras(ByVal sConta As String) As Long
Dim lErro As Long
Dim lComando As Long
Dim sContaContabil As String


On Error GoTo Erro_Conta_Critica_Vinculo_Compras

    'abertura dos comandos
    lComando = Comando_Abrir
    If lComando = 0 Then gError 85217
    
    'inicializa a variavel
    sContaContabil = String(STRING_CONTACONTABIL, 0)
    

    'verifica se a conta apresenta algum vinculo com a view ItensPedCompraTodos
    lErro = Comando_Executar(lComando, "SELECT ContaContabil FROM LocalizacaoItensPC WHERE ContaContabil= ?", sContaContabil, sConta)
    If lErro <> AD_SQL_SUCESSO Then gError 85213

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 85214

    'se encontrou dados com a sconta então erro
    If lErro = AD_SQL_SUCESSO Then gError 85215

'    'verifica se a conta apresenta algum vinculo com a tabela ItensReqModelo
'    lErro = Comando_Executar(lComando, "SELECT ContaContabil FROM ItensReqModelo WHERE ContaContabil= ?", sContaContabil, sConta)
'    If lErro <> AD_SQL_SUCESSO Then gError 85208
'
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 85209
'
'    'se encontrou dados com a sconta então erro
'    If lErro = AD_SQL_SUCESSO Then gError 85210

    lErro = Comando_Executar(lComando, "SELECT ContaContabil FROM ItensReqCompraTodos WHERE ContaContabil= ?", sContaContabil, sConta)
    If lErro <> AD_SQL_SUCESSO Then gError 85203

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 85204

    'se encontrou dados com a sconta então erro
    If lErro = AD_SQL_SUCESSO Then gError 85205

    Call Comando_Fechar(lComando)

    Conta_Critica_Vinculo_Compras = SUCESSO

    Exit Function

Erro_Conta_Critica_Vinculo_Compras:

    Conta_Critica_Vinculo_Compras = gErr

    Select Case gErr

        Case 85217
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 85213, 85214
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LOCALIZACAOITENSPC", gErr)

        Case 85215
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_VINCULADA_LOCALIZACAOITENSPC", gErr, sContaContabil)

        Case 85208, 85209
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSREQMODELO", gErr)

        Case 852010
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_VINCULADA_ITENSREQMODELO", gErr, sContaContabil)

        Case 85203, 85204
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSREQCOMPRATODOS", gErr)

        Case 85205
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_VINCULADA_ITENSREQCOMPRATODOS", gErr, sContaContabil)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149226)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Conta_Critica_Possui_Movimento(ByVal sConta As String, iTemMovimento As Integer) As Long
'verifica se a conta possui movimentação nos exercicios existentes.

Dim lComando As Long
Dim lErro As Long
Dim dCredito(NUM_MAX_PERIODOS) As Double
Dim dDebito(NUM_MAX_PERIODOS) As Double
Dim dMovimento As Double
Dim dSldIni As Double
Dim iIndice As Integer
Dim iFim_de_Arquivo As Integer

On Error GoTo Erro_Conta_Critica_Possui_Movimento

    lComando = 0

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 5461
    
    lErro = Comando_Executar(lComando, "SELECT SldIni, Deb01, Deb02, Deb03, Deb04, Deb05, Deb06, Deb07, Deb08, Deb09, Deb10, Deb11, Deb12, Cre01, Cre02, Cre03, Cre04, Cre05, Cre06, Cre07, Cre08, Cre09, Cre10, Cre11, Cre12 FROM MvPerCta WHERE Conta = ?", dSldIni, dDebito(1), dDebito(2), dDebito(3), dDebito(4), dDebito(5), dDebito(6), dDebito(7), dDebito(8), dDebito(9), dDebito(10), dDebito(11), dDebito(12), dCredito(1), dCredito(2), dCredito(3), dCredito(4), dCredito(5), dCredito(6), dCredito(7), dCredito(8), dCredito(9), dCredito(10), dCredito(11), dCredito(12), sConta)
    If lErro <> AD_SQL_SUCESSO Then Error 5642
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9883
    
    dMovimento = 0
    
    Do While lErro = AD_SQL_SUCESSO And dMovimento = 0
        
        For iIndice = 1 To 12
            dMovimento = dMovimento + dCredito(iIndice) + dDebito(iIndice)
        Next
            
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9884
            
    Loop
    
    If dMovimento <> 0 Then
        iTemMovimento = 1
    Else
        iTemMovimento = 0
    End If
    
    Call Comando_Fechar(lComando)
    
    Conta_Critica_Possui_Movimento = SUCESSO
    
    Exit Function
 
Erro_Conta_Critica_Possui_Movimento:

    Conta_Critica_Possui_Movimento = Err

    Select Case Err
    
        Case 5641
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 5642, 9883, 9884
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVPERCTA2", Err, sConta)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149227)
    
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function Ccl_Critica_Exclusao(ByVal sCcl As String) As Long
'verifica se o Ccl possui rateios associados, etc. /m

Dim lErro As Long
Dim iTemMovimento As Integer
Dim iTemLanPendente As Integer

On Error GoTo Erro_Ccl_Critica_Exclusao

    'Verifica se existem movimentações para este Ccl. Se existem, abortar.
    lErro = CF("Ccl_Critica_Possui_Movimento", sCcl, iTemMovimento)
    If lErro <> SUCESSO Then Error 8119
        
    If iTemMovimento = 1 Then Error 8053
        
    'Verificar se existem lancamentos pendentes
    lErro = CF("Ccl_Critica_Possui_LanPendente", sCcl, iTemLanPendente)
    If lErro <> SUCESSO Then Error 8120
        
    If iTemLanPendente = 1 Then Error 8121
        
    'verifica se o Ccl não tem rateios on-line associados
    lErro = CF("RateioOn_Le_Ccl", sCcl)
    If lErro <> SUCESSO And lErro <> 8098 Then Error 8078
    
    'se existe rateio com o centro de custo em questão
    If lErro = SUCESSO Then Error 8077
    
    'verifica se o Ccl não tem rateios off-line associados
    lErro = CF("RateioOff_Le_Ccl", sCcl)
    If lErro <> SUCESSO And lErro <> 8100 Then Error 8080
    
    'se existe rateio com o centro de custo em questão
    If lErro = SUCESSO Then Error 8079
    
    Ccl_Critica_Exclusao = SUCESSO
    
    Exit Function
        
Erro_Ccl_Critica_Exclusao:
    
    Ccl_Critica_Exclusao = Err
    
    Select Case Err
    
        Case 8053
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CCL_COM_MOVIMENTACAO", Err)
    
        Case 8078, 8080, 8082, 8119, 8120
            
        Case 8077
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CCL_COM_RATEIOON", Err)
            
        Case 8079, 8081
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CCL_COM_RATEIOOFF", Err)
            
        Case 8121
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CCL_COM_LANC_PEND", Err)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149228)
            
    End Select
        
    Exit Function

End Function

Function CclExtraContabil_Critica_Exclusao(ByVal sCcl As String) As Long
'verifica se o Ccl possui rateios associados, etc.

Dim lErro As Long
Dim iTemMovimento As Integer
Dim iTemLanPendente As Integer

On Error GoTo Erro_CclExtraContabil_Critica_Exclusao

    'Verificar se existem lancamentos pendentes
    lErro = CF("Ccl_Critica_Possui_LanPendente", sCcl, iTemLanPendente)
    If lErro <> SUCESSO Then Error 10441
        
    If iTemLanPendente = 1 Then Error 10442
        
    'verifica se o Ccl não tem rateios on-line associados
    lErro = CF("RateioOn_Le_Ccl", sCcl)
    If lErro <> SUCESSO And lErro <> 8098 Then Error 10443
    
    'se existe rateio com o centro de custo em questão
    If lErro = SUCESSO Then Error 10444
    
    'verifica se o Ccl não tem rateios off-line associados
    lErro = CF("RateioOff_Le_Ccl", sCcl)
    If lErro <> SUCESSO And lErro <> 8100 Then Error 10445
    
    'se existe rateio com o centro de custo em questão
    If lErro = SUCESSO Then Error 10446
    
    CclExtraContabil_Critica_Exclusao = SUCESSO
    
    Exit Function
        
Erro_CclExtraContabil_Critica_Exclusao:
    
    CclExtraContabil_Critica_Exclusao = Err
    
    Select Case Err
    
        Case 10441, 10443, 10445
            
        Case 10444
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CCL_COM_RATEIOON", Err)
            
        Case 10446
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CCL_COM_RATEIOOFF", Err)
            
        Case 10442
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CCL_COM_LANC_PEND", Err)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149229)
            
    End Select
        
    Exit Function

End Function



Function CclContabil_Critica_Exclusao(ByVal sCcl As String) As Long
'verifica se o Ccl possui rateios associados, etc.

Dim lErro As Long
Dim sConta As String
Dim lComando As Long

On Error GoTo Erro_CclContabil_Critica_Exclusao

    lComando = 0

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 10422

    sConta = String(STRING_CONTA, 0)

    'pesquisa as associações do ccl com contas
    lErro = Comando_ExecutarPos(lComando, "SELECT Conta FROM ContaCcl WHERE Ccl = ?", 0, sConta, sCcl)
    If lErro <> AD_SQL_SUCESSO Then Error 10423

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 10424
    
    Do While lErro = SUCESSO
    
        'verifica se o centro de custo/lucro possui lancamentos associados
        lErro = CF("ContaAnalitica_Critica_Exclusao", sConta)
        If lErro <> SUCESSO Then Error 10425

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 10426
        
    Loop

    CclContabil_Critica_Exclusao = SUCESSO
    
    Exit Function
        
Erro_CclContabil_Critica_Exclusao:
    
    CclContabil_Critica_Exclusao = Err
    
    Select Case Err
    
        Case 10422
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 10423, 10424, 10426
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTACCL2", Err, sCcl)
            
        Case 10425
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149230)
            
    End Select
        
    Exit Function


End Function


Function Ccl_Critica_Possui_Movimento(ByVal sCcl As String, iTemMovimento As Integer) As Long
'verifica se o Ccl possui movimentação nos exercicios existentes. /m

Dim lComando As Long
Dim lErro As Long
Dim dCredito(NUM_MAX_PERIODOS) As Double
Dim dDebito(NUM_MAX_PERIODOS) As Double
Dim dMovimento As Double
Dim dSldIni As Double
Dim iIndice As Integer

On Error GoTo Erro_Ccl_Critica_Possui_Movimento

    lComando = 0

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 8073
    
    'le os saldos de centro de custo para o centro de custo em questão
    lErro = Comando_Executar(lComando, "SELECT SldIni, Deb01, Deb02, Deb03, Deb04, Deb05, Deb06, Deb07, Deb08, Deb09, Deb10, Deb11, Deb12, Cre01, Cre02, Cre03, Cre04, Cre05, Cre06, Cre07, Cre08, Cre09, Cre10, Cre11, Cre12 FROM MvPerCcl WHERE Ccl = ?", dSldIni, dDebito(1), dDebito(2), dDebito(3), dDebito(4), dDebito(5), dDebito(6), dDebito(7), dDebito(8), dDebito(9), dDebito(10), dDebito(11), dDebito(12), dCredito(1), dCredito(2), dCredito(3), dCredito(4), dCredito(5), dCredito(6), dCredito(7), dCredito(8), dCredito(9), dCredito(10), dCredito(11), dCredito(12), sCcl)
    If lErro <> AD_SQL_SUCESSO Then Error 8074
    
    'le o primeiro saldo de centro de custo
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9862
    
    dMovimento = 0
    
    Do While lErro = AD_SQL_SUCESSO And dMovimento = 0
        
        'acumula os saldos em dMovimento
        For iIndice = 1 To 12
            dMovimento = dMovimento + dCredito(iIndice) + dDebito(iIndice)
        Next
            
        'le o proximo saldo de centro de custo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9863
            
    Loop
    
    If dMovimento <> 0 Then
        iTemMovimento = 1
    Else
        iTemMovimento = 0
    End If
    
    Call Comando_Fechar(lComando)
    
    Ccl_Critica_Possui_Movimento = SUCESSO
    
    Exit Function
 
Erro_Ccl_Critica_Possui_Movimento:

    Ccl_Critica_Possui_Movimento = Err

    Select Case Err
    
        Case 8073
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 8074, 9862, 9863
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVPERCCL3", Err, sCcl)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149231)
    
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function Ccl_Critica_Possui_LanPendente(ByVal sCcl As String, iTemLanPendente As Integer) As Long
'Verifica se existem lançamentos pendentes para o centro de custo em questão. Se existir retorna iTemLanPendente = 1, senao iTemLanPendente = 0

Dim lErro As Long

On Error GoTo Erro_Ccl_Critica_Possui_LanPendente

    iTemLanPendente = 0
    
    'Verificar se existem lancamentos pendentes
    lErro = CF("LanPendente_Le_Ccl", sCcl)
    If lErro <> SUCESSO And lErro <> 5611 Then Error 5668
    
    'Se sucesso ==> Existem lançamentos pendentes
    If lErro = SUCESSO Then
        iTemLanPendente = 1
    End If

    Ccl_Critica_Possui_LanPendente = SUCESSO
    
    Exit Function
 
Erro_Ccl_Critica_Possui_LanPendente:

    Ccl_Critica_Possui_LanPendente = Err

    Select Case Err
    
        Case 5668
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149232)
    
    End Select
    
    Exit Function

End Function

Function Conta_Critica_Possui_LanPendente(ByVal sConta As String, iTemLanPendente As Integer) As Long
'Verifica se existem lançamentos pendentes para a conta em questão. Se existir retorna iTemLanPendente = 1, senao iTemLanPendente = 0

Dim lErro As Long

On Error GoTo Erro_Conta_Critica_Possui_LanPendente

    iTemLanPendente = 0
    
    'Verificar se existem lancamentos pendentes
    lErro = CF("LanPendente_Le_Conta", sConta)
    If lErro <> SUCESSO And lErro <> 5461 Then Error 5671
    
    'Se sucesso ==> Existem lançamentos pendentes
    If lErro = SUCESSO Then
        iTemLanPendente = 1
    End If

    Conta_Critica_Possui_LanPendente = SUCESSO
    
    Exit Function
 
Erro_Conta_Critica_Possui_LanPendente:

    Conta_Critica_Possui_LanPendente = Err

    Select Case Err
    
        Case 5671
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149233)
    
    End Select
    
    Exit Function

End Function


Function Ccl_Critica(ByVal sCcl As String, sCclFormatada As String, ByVal objCcl As ClassCcl) As Long
'verifica se o centro de custo/lucro está em condições de receber lançamentos. Devolve os dados de Ccl em objCcl.

Dim lErro As Long
Dim iCclPreenchida As Integer

On Error GoTo Erro_Ccl_Critica

    If Len(sCcl) > 0 Then
    
        lErro = CF("Ccl_Formata", sCcl, sCclFormatada, iCclPreenchida)
        If lErro <> SUCESSO Then Error 5706
    
        If iCclPreenchida = CCL_PREENCHIDA Then
        
            objCcl.sCcl = sCclFormatada
    
            lErro = CF("Ccl_Le", objCcl)
            If lErro <> SUCESSO And lErro <> 5599 Then Error 5702
    
            'Ausencia de Ccl no BD
            If lErro = 5599 Then Error 5703
            
            If objCcl.iTipoCcl <> CCL_ANALITICA Then Error 10515
            
            If objCcl.iAtivo = 0 Then Error 32294
            
        End If
        
        
    End If
    
    Ccl_Critica = SUCESSO
    
    Exit Function

Erro_Ccl_Critica:

    Ccl_Critica = Err
    
    Select Case Err
    
        Case 5702, 5703, 5706
    
        Case 10515
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_ANALITICA1", Err, objCcl.sCcl)
  
        Case 32294
            Call Rotina_Erro(vbOKOnly, "ERRO_CCL_INATIVO", Err, objCcl.sCcl)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149234)
        
    End Select
    
    Exit Function

End Function

Function Ccl_Critica_Analitica(ByVal sCcl As String) As Long
'verifica se o centro de custo é analitico

Dim lErro As Long
Dim lComando As Long
Dim iTipoCcl As Integer, iAtivoCcl As Integer

On Error GoTo Erro_Ccl_Critica_Analitica

    If Len(Trim(sCcl)) <> 0 Then
    
        lComando = Comando_Abrir()
        If lComando = 0 Then Error 10491
    
        'pesquisa o centro de custo em questão no BD
        lErro = Comando_Executar(lComando, "SELECT TipoCcl, AtivoCcl FROM Ccl WHERE Ccl=?", iTipoCcl, iAtivoCcl, sCcl)
        If lErro <> AD_SQL_SUCESSO Then Error 10492
    
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 10493
        
        'verifica se o centro de custo não é analitico
        If iTipoCcl <> CCL_ANALITICA Then Error 10494
    
        If iAtivoCcl = 0 Then Error 32311
            
        Call Comando_Fechar(lComando)
    
    End If
    
    Ccl_Critica_Analitica = SUCESSO
    
    Exit Function
    
Erro_Ccl_Critica_Analitica:

    Ccl_Critica_Analitica = Err
    
    Select Case Err
    
        Case 10491
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 32311
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_INATIVO", Err, sCcl)
            
        Case 10492, 10493
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CCL", Err, sCcl)
        
        Case 10494
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_ANALITICA1", Err, sCcl)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149235)
        
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
        
End Function

Function LanPendente_Critica_TotaisLote(ByVal objLote As ClassLote, iTotaisIguais As Integer) As Long

Dim lErro As Long
Dim lComando As Long
Dim dCreditos As Double
Dim dDebitos As Double
Dim iNumLanc As Integer
Dim iNumDoc As Integer
Dim dValor As Double
Dim lDoc As Long
Dim lDoc1 As Long

On Error GoTo Erro_LanPendente_Critica_TotaisLote

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 8031

    lErro = Comando_Executar(lComando, "SELECT Valor, Doc FROM LanPendente WHERE FilialEmpresa = ? AND Origem=? AND Exercicio=? AND PeriodoLote=? AND Lote=? ORDER BY Doc", dValor, lDoc, objLote.iFilialEmpresa, objLote.sOrigem, objLote.iExercicio, objLote.iPeriodo, objLote.iLote)
    If lErro <> AD_SQL_SUCESSO Then Error 8032
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9885
    
    dCreditos = 0
    dDebitos = 0
    iNumLanc = 0
    iNumDoc = 0
    
    Do While lErro = SUCESSO
    
        lDoc1 = lDoc
        
        Do While lErro = SUCESSO And lDoc1 = lDoc
    
            'Contabiliza credito/debito
            If dValor > 0 Then
                dCreditos = dCreditos + Round(dValor, 2)
            Else
                dDebitos = dDebitos - Round(dValor, 2)
            End If
            
            'Contabiliza lancamentos pendentes
            iNumLanc = iNumLanc + 1
            
            'Le o proximo lancamento pendente
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9886
                
        Loop
                
        iNumDoc = iNumDoc + 1
        
    Loop
    
    objLote.iNumLancAtual = iNumLanc
    
    'Verificar se existe diferenca entre totais do lote e valores atuais
    If Round(objLote.dTotCre, 2) <> Round(dCreditos, 2) Or Round(objLote.dTotDeb, 2) <> Round(dDebitos, 2) Or objLote.iNumDocAtual <> iNumDoc Then
        
        objLote.dTotCre = Round(dCreditos, 2)
        objLote.dTotDeb = Round(dDebitos, 2)
        objLote.iNumDocAtual = iNumDoc
    
    End If
    
    'Verificar se existe diferenca entre informado e atual
    If Round(objLote.dTotInf, 2) <> Round(dCreditos, 2) Or Round(objLote.dTotInf, 2) <> Round(dDebitos, 2) Or objLote.iNumDocAtual <> iNumDoc Then
        
        iTotaisIguais = 1
    
    Else
        iTotaisIguais = 0
    End If
    
    Call Comando_Fechar(lComando)
   
    LanPendente_Critica_TotaisLote = SUCESSO
    
    Exit Function

Erro_LanPendente_Critica_TotaisLote:

    LanPendente_Critica_TotaisLote = Err

    Select Case Err
    
        Case 8031
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 8032, 9885, 9886
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LANPENDENTE1", Err, objLote.iFilialEmpresa, objLote.sOrigem, objLote.iExercicio, objLote.iPeriodo, objLote.iLote)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149236)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function ContaCcl_Critica_Colecao_ExtraContab(ByVal sConta As String, ByVal colCcl As Collection) As Long
'Verifica se os Ccl's associados a sConta estao presentes em  colCcl /m
'PARA EXECUTAR ESTA ROTINA O PROGRAMA TEM QUE ESTAR EM TRANSACAO

Dim lErro As Long
Dim sCcl As Variant
Dim sCcl1 As String
Dim lComando As Long
Dim iEsta_na_Colecao As Integer

On Error GoTo Erro_ContaCcl_Critica_Colecao_ExtraContab

    lComando = 0

    sCcl1 = String(STRING_CCL, 0)
    
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 8205
   
    'pesquisa os centros de custo/lucro associados à conta em questão
    lErro = Comando_Executar(lComando, "SELECT Ccl FROM ContaCcl WHERE Conta = ?", sCcl1, sConta)
    If lErro <> AD_SQL_SUCESSO Then Error 8150
   
    'le a primeira associação
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9881

    'Se o Centro de custo/lucro nao estiver na colecao, precisa ser excluido
    Do While lErro = AD_SQL_SUCESSO
   
        iEsta_na_Colecao = 0
   
        For Each sCcl In colCcl
       
            If sCcl = sCcl1 Then
        
                iEsta_na_Colecao = 1
                Exit For
            
            End If
           
        Next
       
        'Se o centro de custo/lucro nao esta na colecao
        If iEsta_na_Colecao = 0 Then
        
            'exclui a associacao
            lErro = CF("Ccl_Exclui_Sem_Assoc_ExtraContabil", sConta, sCcl1)
            If lErro <> SUCESSO And lErro <> 8156 Then Error 8206
            
       End If
       
       lErro = Comando_BuscarProximo(lComando)
       If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9882
       
    Loop
   
    Call Comando_Fechar(lComando)
   
    ContaCcl_Critica_Colecao_ExtraContab = SUCESSO
    
    Exit Function

Erro_ContaCcl_Critica_Colecao_ExtraContab:

    ContaCcl_Critica_Colecao_ExtraContab = Err
    
    Select Case Err
                
        Case 8150, 9881, 9882
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTACCL1", Err, sConta)
               
        Case 8205
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 8206
                   
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 149237)
        
    End Select
        
    Call Comando_Fechar(lComando)
        
End Function

Public Function ContaCcl_Critica_Colecao_Contabil(ByVal sConta As String, ByVal colCcl As Collection) As Long
'Verifica se os Ccl's associados a sConta estao presentes em  colCcl /m
'PARA EXECUTAR ESTA ROTINA O PROGRAMA TEM QUE ESTAR EM TRANSACAO

Dim lErro As Long
Dim sCcl As Variant
Dim sCcl1 As String
Dim lComando As Long
Dim iEsta_na_Colecao As Integer
Dim iContaPreenchida As Integer

On Error GoTo Erro_ContaCcl_Critica_Colecao_Contabil

    lComando = 0

    sCcl1 = String(STRING_CCL, 0)
    
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 8318
       
    'pesquisa os centros de custo/lucro associados à conta em questão
    lErro = Comando_Executar(lComando, "SELECT Ccl FROM ContaCcl WHERE Conta = ?", sCcl1, sConta)
    If lErro <> AD_SQL_SUCESSO Then Error 8320
   
    'le a primeira associação
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9887

    'Se o Ccl nao estiver na colecao, precisa ser excluida
    Do While lErro = AD_SQL_SUCESSO
   
        iEsta_na_Colecao = 0
   
        For Each sCcl In colCcl
       
            If sCcl = sCcl1 Then
            
                iEsta_na_Colecao = 1
                Exit For
                
            End If
           
        Next
       
        'Se o centro de custo/lucro nao esta na colecao
        If iEsta_na_Colecao = 0 Then
        
            'exclui a associacao
            lErro = CF("Ccl_Exclui_Sem_Assoc_Contabil", sConta, sCcl1)
            If lErro <> SUCESSO And lErro <> 8354 Then Error 8321
            
        End If
       
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9888
       
    Loop
   
    Call Comando_Fechar(lComando)
   
    ContaCcl_Critica_Colecao_Contabil = SUCESSO
    
    Exit Function

Erro_ContaCcl_Critica_Colecao_Contabil:

    ContaCcl_Critica_Colecao_Contabil = Err
    
    Select Case Err
                
        Case 8318
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 8320, 9887, 9888
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTACCL1", Err, sConta)
               
        Case 8321
                   
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 149238)
        
    End Select
        
    Call Comando_Fechar(lComando)
        
    Exit Function
   
End Function

Function ContaCcl_Critica_Possui_Movimento(ByVal sConta As String, ByVal sCcl As String, iTemMovimento As Integer) As Long
'verifica se a assoc possui movimentação nos exercicios existentes.
'PARA EXECUTAR ESTA ROTINA O PROGRAMA TEM QUE ESTAR EM TRANSACAO

Dim lComando As Long
Dim lErro As Long
Dim dCredito(NUM_MAX_PERIODOS) As Double
Dim dDebito(NUM_MAX_PERIODOS) As Double
Dim dMovimento As Double
Dim dSldIni As Double
Dim iIndice As Integer

On Error GoTo Erro_ContaCcl_Critica_Possui_Movimento

    lComando = 0

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 8167
    
    lErro = Comando_Executar(lComando, "SELECT SldIni, Deb01, Deb02, Deb03, Deb04, Deb05, Deb06, Deb07, Deb08, Deb09, Deb10, Deb11, Deb12, Cre01, Cre02, Cre03, Cre04, Cre05, Cre06, Cre07, Cre08, Cre09, Cre10, Cre11, Cre12 FROM MvPerCcl WHERE Conta=? AND Ccl = ?", dSldIni, dDebito(1), dDebito(2), dDebito(3), dDebito(4), dDebito(5), dDebito(6), dDebito(7), dDebito(8), dDebito(9), dDebito(10), dDebito(11), dDebito(12), dCredito(1), dCredito(2), dCredito(3), dCredito(4), dCredito(5), dCredito(6), dCredito(7), dCredito(8), dCredito(9), dCredito(10), dCredito(11), dCredito(12), sConta, sCcl)
    If lErro <> AD_SQL_SUCESSO Then Error 8168
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9889
    
    dMovimento = 0
    
    Do While lErro = AD_SQL_SUCESSO And dMovimento = 0
    
        For iIndice = 1 To 12
            dMovimento = dMovimento + dCredito(iIndice) + dDebito(iIndice)
        Next
            
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9890
            
    Loop
    
    If dMovimento <> 0 Then
        iTemMovimento = 1
    Else
        iTemMovimento = 0
    End If
    
    Call Comando_Fechar(lComando)
    
    ContaCcl_Critica_Possui_Movimento = SUCESSO
    
    Exit Function
 
Erro_ContaCcl_Critica_Possui_Movimento:

    ContaCcl_Critica_Possui_Movimento = Err

    Select Case Err
    
        Case 8167
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 8168, 9889, 9890
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVPERCCL1", Err, sCcl, sConta)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149239)
    
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function Saldo_Inicial_Critica(iDisponivel As Integer) As Long
'verifica se o saldo inicial pode ser alterado

Dim objConfiguracao As New ClassConfiguracao
Dim objExercicio As New ClassExercicio
Dim lErro As Long

On Error GoTo Erro_Saldo_Inicial_Critica

    lErro = CF("Configuracao_Le", objConfiguracao)
    If lErro <> SUCESSO Then Error 5961
    
    lErro = CF("Exercicio_Le_Data", objConfiguracao.dtDataImplantacao, objExercicio)
    If lErro <> SUCESSO And lErro <> 5965 Then Error 5966
    
    'exercicio não encontrado
    If lErro = 5965 Then
        iDisponivel = SLDINI_NAO_DISPONIVEL
    Else
    
        'se o primeiro exercicio está fechado ==> não pode habilitar o saldo inicial da conta
        If objExercicio.iStatus <> EXERCICIO_FECHADO Then
            iDisponivel = SLDINI_DISPONIVEL
        Else
            iDisponivel = SLDINI_NAO_DISPONIVEL
        End If
        
    End If
        
    Saldo_Inicial_Critica = SUCESSO
    
    Exit Function
    
Erro_Saldo_Inicial_Critica:

    Saldo_Inicial_Critica = Err

    Select Case Err
    
        Case 5961, 5966, 5967, 5968
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149240)
        
    End Select

    Exit Function
        
End Function

Function Conta_Critica_ContaPai(ByVal sConta As String, ByVal sSiglaModulo As String) As Long
'verifica se a conta tem um pai sintetico

Dim sConta1 As String
Dim iNivel As Integer
Dim lErro As Long
Dim objPlanoConta As New ClassPlanoConta

On Error GoTo Erro_Conta_Critica_ContaPai

    'retorna o nivel da conta
    lErro = Mascara_Conta_ObterNivel(sConta, iNivel)
    If lErro <> SUCESSO Then Error 5481
    
    'se a conta tem um "pai"
    If iNivel > 1 Then
    
        sConta1 = String(STRING_CONTA, 0)
    
        'descobre a conta "pai"
        lErro = Mascara_RetornaContaNoNivel(iNivel - 1, sConta, sConta1)
        If lErro <> SUCESSO Then Error 5482
    
        'verifica se a conta "pai" está cadastrada
        lErro = CF("Conta_SelecionaUma", sConta1, objPlanoConta, sSiglaModulo)
        If lErro <> SUCESSO And lErro <> 6030 Then Error 5484
        
        'nao encontrou a conta
        If lErro = 6030 Then Error 5483
        
        'verifica se a conta pai é sintetica. Se for analitica é erro.
        If objPlanoConta.iTipoConta = CONTA_ANALITICA Then Error 5654
        
    End If
        
    Conta_Critica_ContaPai = SUCESSO
    
    Exit Function
    
Erro_Conta_Critica_ContaPai:

    Conta_Critica_ContaPai = Err
    
    Select Case Err
    
        Case 5481
    
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_CONTA_OBTERNIVEL", Err, sConta)
    
        Case 5482
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_RetornaContaNoNivel", Err, sConta, iNivel - 1)
    
        Case 5484
            
        Case 5483
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTAPAI_INEXISTENTE", Err)
            
        Case 5654
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTAPAI_ANALITICA", Err)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149241)
        
    End Select

    Exit Function

End Function

Function Ccl_Critica_CclPai(ByVal sCcl As String) As Long
'verifica se o centro de custo/lucro tem um pai sintetico

Dim sCcl1 As String
Dim iNivel As Integer
Dim lErro As Long
Dim objCcl As New ClassCcl

On Error GoTo Erro_Ccl_Critica_CclPai

    'retorna o nivel do centro de custo/lucro
    lErro = Mascara_Ccl_ObterNivel(sCcl, iNivel)
    If lErro <> SUCESSO Then Error 10457
    
    'se o centro de custo/lucro tem um "pai"
    If iNivel > 1 Then
    
        sCcl1 = String(STRING_CCL, 0)
    
        'descobre o centro de custo/lucro "pai"
        lErro = Mascara_RetornaCclNoNivel(iNivel - 1, sCcl, sCcl1)
        If lErro <> SUCESSO Then Error 10458
    
        objCcl.sCcl = sCcl1
        
        'verifica se o centro de custo/lucro "pai" está cadastrado
        lErro = CF("Ccl_Le", objCcl)
        If lErro <> SUCESSO And lErro <> 5599 Then Error 10459
        
        'nao encontrou o centro de custo/lucro
        If lErro = 5599 Then Error 10460
        
        'verifica se o centro de custo/lucro pai é sintetico. Se for analitico é erro.
        If objCcl.iTipoCcl = CCL_ANALITICA Then Error 10461
        
    End If
        
    Ccl_Critica_CclPai = SUCESSO
    
    Exit Function
    
Erro_Ccl_Critica_CclPai:

    Ccl_Critica_CclPai = Err
    
    Select Case Err
    
        Case 10457
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_CCL_OBTERNIVEL", Err, sCcl)
    
        Case 10458
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_RetornaCclNoNivel", Err, sCcl, iNivel - 1)
    
        Case 10459
            
        Case 10460
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCLPAI_INEXISTENTE", Err)
            
        Case 10461
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCLPAI_ANALITICA", Err)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149242)
        
    End Select

    Exit Function

End Function

Function Conta_Critica_Categoria(ByVal sConta As String, ByVal iCategoria As Integer) As Long
'verifica se a conta é de nivel 1 e tem a categoria preenchida.

Dim iNivel As Integer
Dim lErro As Long

On Error GoTo Erro_Conta_Critica_Categoria

    'retorna o nivel da conta
    lErro = Mascara_Conta_ObterNivel(sConta, iNivel)
    If lErro <> SUCESSO Then Error 9718
    
    'se é uma conta de nivel um tem que especificar uma categoria. Senão ==> erro.
    If iNivel = 1 And iCategoria = 0 Then Error 9719
    
    Conta_Critica_Categoria = SUCESSO
    
    Exit Function
    
Erro_Conta_Critica_Categoria:

    Conta_Critica_Categoria = Err
    
    Select Case Err
    
        Case 9718
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_CONTA_OBTERNIVEL", Err, sConta)
    
        Case 9719
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NIVEL1_CATEGORIA", Err)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149243)
        
    End Select

    Exit Function

End Function

Function ContaCcl_Critica_Possui_LanPendente(ByVal sConta As String, ByVal sCcl As String, iTemLanPendente As Integer) As Long
'Verifica se existem lançamentos pendentes para a conta e centro de custo extra contabil em questão. Se existir retorna iTemLanPendente = 1, senao iTemLanPendente = 0

Dim lErro As Long

On Error GoTo Erro_ContaCcl_Critica_Possui_LanPendente

    iTemLanPendente = 0
    
    'Verificar se existem lancamentos pendentes
    lErro = CF("LanPendente_Le_ContaCcl", sConta, sCcl)
    If lErro <> SUCESSO And lErro <> 9894 Then Error 9896
    
    'Se sucesso ==> Existem lançamentos pendentes
    If lErro = SUCESSO Then
        iTemLanPendente = 1
    End If

    ContaCcl_Critica_Possui_LanPendente = SUCESSO
    
    Exit Function
 
Erro_ContaCcl_Critica_Possui_LanPendente:

    ContaCcl_Critica_Possui_LanPendente = Err

    Select Case Err
    
        Case 9896
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149244)
    
    End Select
    
    Exit Function

End Function

Function ContaCcl_Critica_Exclusao(ByVal sConta As String, ByVal sCcl As String) As Long
'verifica se a conta+ccl possui movimento, lancamentos pendentes, rateioon associado e rateiooff associado.

Dim lErro As Long
Dim iTemLanPendente As Integer
Dim iTemMovimento As Integer

On Error GoTo Erro_ContaCcl_Critica_Exclusao

    'Verifica se existem movimentações para esta Assoc. Se existem, abortar.
    lErro = CF("ContaCcl_Critica_Possui_Movimento", sConta, sCcl, iTemMovimento)
    If lErro <> SUCESSO Then Error 8158
        
    If iTemMovimento = 1 Then Error 8159

    'Verificar se existem lancamentos pendentes
    lErro = CF("ContaCcl_Critica_Possui_LanPendente", sConta, sCcl, iTemLanPendente)
    If lErro <> SUCESSO Then Error 9905
        
    If iTemLanPendente = 1 Then Error 9906
        
    'verifica se a conta não tem rateios on-line associados
    lErro = CF("RateioOn_Le_ContaCcl", sConta, sCcl)
    If lErro <> SUCESSO And lErro <> 9903 Then Error 9907
    
    If lErro = SUCESSO Then Error 9908
    
    'verifica se existe na tabela RateioOff pelo menos um rateio com a conta e centro de custo em questão.
    lErro = CF("RateioOff_Le_ContaCcl", sConta, sCcl)
    If lErro <> SUCESSO And lErro <> 44049 Then Error 20607
    
    If lErro = SUCESSO Then Error 20608
    
    ContaCcl_Critica_Exclusao = SUCESSO
    
    Exit Function
        
Erro_ContaCcl_Critica_Exclusao:

    ContaCcl_Critica_Exclusao = Err
    
    Select Case Err
    
        Case 5583
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CONTA_COM_RATEIOON", Err)
        
        Case 5584
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CONTA_COM_RATEIOOFF", Err)
            
        Case 5670
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CONTA_COM_LANC_PEND1", Err)
        
        Case 8158, 9905, 9907, 20607
        
        Case 8159
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ASSOC_CCLCONTA_COM_MOV", Err, sConta, sCcl)
        
        Case 9906
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CCLCONTA_COM_LANPENDENTE", Err, sConta, sCcl)
            
        Case 9908
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CONTACCL_COM_RATEIOON", Err, sConta, sCcl)
        
        Case 20608
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CONTACCL_COM_RATEIOOFF", Err, sConta, sCcl)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149245)
            
    End Select
        
    Exit Function
    


End Function

Function Exercicio_Critica_Movimento(ByVal iExercicio As Integer) As Long
'Verifica se existe movimento, lote ou lançamento pendente no exercício iExercicio

Dim lErro As Long
Dim objExercicio As New ClassExercicio

On Error GoTo Erro_Exercicio_Critica_Movimento

    If iExercicio <> 0 Then

        lErro = CF("Exercicio_Le", iExercicio, objExercicio)
        If lErro <> SUCESSO And lErro <> 10083 Then Error 10089

        'se o exercicio está cadastrado
        If lErro = SUCESSO Then

            'verifica se o exercicio está aberto
            If objExercicio.iStatus <> EXERCICIO_ABERTO Then Error 13663
    
            'verifica se existe lote pendente para o exercicio em questão
            lErro = CF("LotePendente_Le_Exercicio", iExercicio)
            If lErro <> SUCESSO And lErro <> 13615 Then Error 13654
    
            If lErro = SUCESSO Then Error 13663
    
            'verifica se existe lançamento pendente para o exercicio em questão
            lErro = CF("LanPendente_Le_Exercicio", iExercicio)
            If lErro <> SUCESSO And lErro <> 13611 Then Error 13655
    
            If lErro = SUCESSO Then Error 13663
    
            'verifica se existe orçamento para o exercicio em questão
            lErro = CF("Orcamento_Le_Exercicio", iExercicio)
            If lErro <> SUCESSO And lErro <> 11407 Then Error 11408
    
            If lErro = SUCESSO Then Error 13663
    
            'verifica se existe movimento para o exercicio em questão
            lErro = CF("Movimento_Le_Exercicio", iExercicio)
            If lErro <> SUCESSO And lErro <> 13604 Then Error 13656
    
            If lErro = SUCESSO Then Error 13663

        End If

    End If

    Exercicio_Critica_Movimento = SUCESSO

    Exit Function

Erro_Exercicio_Critica_Movimento:

    Exercicio_Critica_Movimento = Err

    Select Case Err

        Case 10089, 11408

        Case 13663, 13654, 13655, 13656

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149246)

    End Select

    Exit Function

End Function

Function Ccl_Critica_Tipo(ByVal objCcl As ClassCcl) As Long
'critica o tipo do centro de custo/lucro

Dim lErro As Long
Dim iTemFilho As Integer

On Error GoTo Erro_Ccl_Critica_Tipo

    'Se o centro de custo está sendo colocada como sintetico.
    'Verifica:
    '  Se não tem associação com conta

    'critica o tipo do centro de custo/lucro
    If objCcl.iTipoCcl = CCL_SINTETICA Then
    
        'pesquisa as associações do centro de custo/lucro com contas
        lErro = CF("ContaCcl_Le_Ccl", objCcl.sCcl)
        If lErro <> SUCESSO And lErro <> 5603 Then Error 10575
            
        'Se encontrou alguma associacao com conta
        If lErro = SUCESSO Then Error 10576
        
        'verifica se foi utilizado em movimento de estoque
        lErro = CF("MovimentoEstoque_Le_Ccl", objCcl.sCcl)
        If lErro <> SUCESSO And lErro <> 60868 Then Error 60870
        
        'se encontrou algum movimento de estoque utilizando o centro de custo
        If lErro = SUCESSO Then Error 60871
        
    Else
    
        'Se o centro de custo está sendo colocado como analitico verifica se
        'não tem centros de custo/lucro filhos
        lErro = CF("Ccl_Tem_Filho", objCcl.sCcl, iTemFilho)
        If lErro <> SUCESSO Then Error 10577
        
        If iTemFilho = CCL_TEM_FILHOS Then Error 10578
        
    End If

    Ccl_Critica_Tipo = SUCESSO
    
    Exit Function
    
Erro_Ccl_Critica_Tipo:

    Ccl_Critica_Tipo = Err

    Select Case Err

        Case 10575, 10577, 60870
    
        Case 10576
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_SINTETICA_ASSOCIADA_CONTA", Err, objCcl.sCcl)
        
        Case 10578
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_ANALITICA_COM_FILHOS", Err, objCcl.sCcl)
            
        Case 60871
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_SINTETICA_USADA_EM_MOVESTOQUE", Err, objCcl.sCcl)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149247)

    End Select
    
    Exit Function

End Function

Function Conta_Critica_Tipo(ByVal sConta As String, ByVal iTipoConta As Integer) As Long
'critica o tipo da conta

Dim sConta1 As String
Dim objPlanoConta As New ClassPlanoConta
Dim lErro As Long
Dim iNivel As Integer
Dim iTemFilho As Integer
Dim sCcl As String

On Error GoTo Erro_Conta_Critica_Tipo

    'Se a conta está sendo colocada como sintetica.
    'Verifica:
    '  a) Se não tem lançamentos
    '  b) Se não tem associação com centro de custo (se estiver trabalhando com centro de custo extra - contabil)
    If iTipoConta = CONTA_SINTETICA Then
    
        'verifica se tem algum lançamento associado a conta em questão
        lErro = CF("Lancamento_Le_Conta", sConta)
        If lErro <> SUCESSO And lErro <> 5457 Then Error 5464
        
        'Se encontrou algum lancamento
        If lErro = SUCESSO Then Error 5463
        
        'verifica se tem algum lançamento pendente associado a conta em questão
        lErro = CF("LanPendente_Le_Conta", sConta)
        If lErro <> SUCESSO And lErro <> 5461 Then Error 5466
        
        'Se encontrou algum lancamento pendente
        If lErro = SUCESSO Then Error 5465
        
        'se for centro de custo extra-contabil
        If giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then
    
            'pesquisa as associações da conta com centros de custo/lucro
            lErro = CF("ContaCcl_Le_Conta", sConta)
            If lErro <> SUCESSO And lErro <> 5469 Then Error 5472
            
            'Se encontrou alguma associacao com centro de custo/lucro
            If lErro = SUCESSO Then Error 5471
            
        ElseIf giSetupUsoCcl = CCL_USA_CONTABIL Then
     
            sCcl = String(STRING_CCL, 0)
     
            'pesquisa as associações da conta com centros de custo/lucro
            lErro = Mascara_RetornaCcl(sConta, sCcl)
            If lErro <> SUCESSO Then Error 10464
            
            If Len(sCcl) > 0 Then Error 10465
            
        End If
        
    Else
    
        'Se a conta está sendo colocada como analitica verifica se
        'não tem contas filhas
        lErro = CF("PlanoConta_Tem_Filho", sConta, iTemFilho)
        If lErro <> SUCESSO Then Error 5478
        
        If iTemFilho = 1 Then Error 5480
        
    End If
    
    Conta_Critica_Tipo = SUCESSO
    
    Exit Function
        
Erro_Conta_Critica_Tipo:

    Conta_Critica_Tipo = Err
    
    Select Case Err
    
        Case 5463
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_SINTETICA_COM_LANCAMENTOS", Err)
    
        Case 5464, 5466, 5472, 5478, 5479
        
        Case 5477
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_CONTA_OBTERNIVEL", Err, sConta)
            
        Case 5465
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_SINTETICA_COM_LANC_PEND", Err)
        
        Case 5471, 10465
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_SINTETICA_ASSOCIADA_CCL", Err)
        
        Case 5480
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_ANALITICA_COM_FILHAS", Err)
            
        Case 10464
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_RetornaCcl", Err, sConta)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149248)
        
    End Select
        
    Exit Function
    
End Function

Function ContaSimples_Critica(ByVal sContaSimplesText As String, ByVal sContaSimplesClipText As String, ByVal objPlanoConta As ClassPlanoConta) As Long
'verifica se a contasimples está em condições de receber lançamentos. Devolve os dados da ContaSimples em objPlanoConta

Dim lErro As Long
Dim iContaPreenchida As Integer
Dim lContaSimples As Long
Dim sConta As String

On Error GoTo Erro_ContaSimples_Critica

    If Left(sContaSimplesText, 1) = " " Then Error 44033

    If Len(sContaSimplesClipText) > 0 Then
    
        'verifica se é um numero inteiro entre 1 e 99.999
        lErro = Long_Critica1(sContaSimplesClipText)
    
        'se for verifica se é uma conta simples cadastrada
        If lErro = SUCESSO Then
    
            lContaSimples = CLng(sContaSimplesClipText)
    
            lErro = CF("PlanoConta_Le_ContaSimples", lContaSimples, objPlanoConta)
            If lErro <> SUCESSO And lErro <> 5451 Then Error 44041
            
            If lErro = 5451 Then Error 44037
            
            If objPlanoConta.iAtivo <> CONTA_ATIVA Then Error 44034
            
            If objPlanoConta.iTipoConta <> CONTA_ANALITICA Then Error 44035
            
        Else
            Error 44033
        End If
    
    Else
        Error 44033
    End If
    
    ContaSimples_Critica = SUCESSO
    
    Exit Function

Erro_ContaSimples_Critica:

    ContaSimples_Critica = Err
    
    Select Case Err
    
        Case 44033
        
        Case 44034
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_INATIVA", Err, sContaSimplesText)
        
        Case 44035
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_ANALITICA", Err, sContaSimplesText)
        
        Case 44037 'não está cadastrado
        
        Case 44041
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149249)
        
    End Select
    
    Exit Function

End Function

Function ContaSimples_Critica_Modulo(ByVal sContaSimplesText As String, ByVal sContaSimplesClipText As String, ByVal objPlanoConta As ClassPlanoConta, ByVal sSiglaModulo As String) As Long
'verifica se a contasimples está em condições de receber lançamentos. Devolve os dados da ContaSimples em objPlanoConta

Dim lErro As Long
Dim iContaPreenchida As Integer
Dim lContaSimples As Long
Dim sConta As String

On Error GoTo Erro_ContaSimples_Critica_Modulo

    If Left(sContaSimplesText, 1) = " " Then Error 44096

    If Len(sContaSimplesClipText) > 0 Then
    
        'verifica se é um numero inteiro entre 1 e 99.999
        lErro = Long_Critica1(sContaSimplesClipText)
    
        'se for verifica se é uma conta simples cadastrada
        If lErro = SUCESSO Then
    
            lContaSimples = CLng(sContaSimplesClipText)
    
            lErro = CF("PlanoConta_Le_ContaSimples", lContaSimples, objPlanoConta)
            If lErro <> SUCESSO And lErro <> 5451 Then Error 44097
            
            If lErro = 5451 Then Error 44098
            
            If objPlanoConta.iAtivo <> CONTA_ATIVA Then Error 44099
            
            If objPlanoConta.iTipoConta <> CONTA_ANALITICA Then Error 44100
            
            If objPlanoConta.UsaModulo(sSiglaModulo) = NAO_USA_MODULO Then Error 44101
            
        Else
            Error 44096
        End If
    
    Else
        Error 44096
    End If
    
    ContaSimples_Critica_Modulo = SUCESSO
    
    Exit Function

Erro_ContaSimples_Critica_Modulo:

    ContaSimples_Critica_Modulo = Err
    
    Select Case Err
    
        Case 44096, 44097, 44098
        
        Case 44099
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_INATIVA", Err, sContaSimplesText)
        
        Case 44100
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_ANALITICA", Err, sContaSimplesText)
        
        Case 44101
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_VISIVEL_MODULO", Err, sContaSimplesText)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149250)
        
    End Select
    
    Exit Function

End Function



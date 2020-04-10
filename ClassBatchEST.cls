VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassBatchEST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Responsavel pela Rotina Atualizacao Inventário: Mario
'Data: 30/10/98
'Pendencias: Não tem

Option Explicit

Function Rotina_CustoMedioProducao_Calcula(ByVal sNomeArqParam As String, ByVal iAno As Integer, ByVal iMes As Integer) As Long
'calcula o custo médio de produção para mes/ano passados e valora os movimentos de estoque

On Error GoTo Erro_Rotina_CustoMedioProducao_Calcula

    TelaAcompanhaBatchEST.Caption = TITULO_TELABATCH_CUSTO_MEDIO_PRODUCAO
    TelaAcompanhaBatchEST.sNomeArqParam = sNomeArqParam
    TelaAcompanhaBatchEST.iRotinaBatch = ROTINA_CUSTO_MEDIO_PRODUCAO_BATCH

    'Parametros da rotina
    TelaAcompanhaBatchEST.iAno = iAno
    TelaAcompanhaBatchEST.iMes = iMes

    'Exibe a tela de acompanhamento do batch
    TelaAcompanhaBatchEST.Show

    Rotina_CustoMedioProducao_Calcula = SUCESSO

    Exit Function

Erro_Rotina_CustoMedioProducao_Calcula:

    Rotina_CustoMedioProducao_Calcula = Err

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 144649)

    End Select

    Exit Function

End Function

Function Rotina_Atualizacao_InvLote(ByVal sNomeArqParam As String, ByVal iID_Atualizacao As Integer, ByVal dtDataContabil As Date, ByVal iLoteContabil As Integer) As Long

On Error GoTo Erro_Rotina_Atualizacao_InvLote

    TelaAcompanhaBatchEST.Caption = TITULO_TELABATCH_ATUALIZA_INVLOTE
    TelaAcompanhaBatchEST.sNomeArqParam = sNomeArqParam
    TelaAcompanhaBatchEST.iIdAtualizacao_Param = iID_Atualizacao
    Set TelaAcompanhaBatchEST.gobjAtuInvLoteAux = New ClassAtualizacaoInvLoteAux
    TelaAcompanhaBatchEST.gobjAtuInvLoteAux.dtContabil = dtDataContabil
    TelaAcompanhaBatchEST.gobjAtuInvLoteAux.iLoteContabil = iLoteContabil
    TelaAcompanhaBatchEST.iRotinaBatch = ROTINA_ATUALIZA_INVLOTE_BATCH
    
    'Exibe a tela de acompanhamento do batch
    TelaAcompanhaBatchEST.Show

    Rotina_Atualizacao_InvLote = SUCESSO
    
    Exit Function

Erro_Rotina_Atualizacao_InvLote:

    Rotina_Atualizacao_InvLote = Err

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144650)

    End Select
    
    Exit Function

End Function

Function Rotina_Reproc_MovEstoque(ByVal sNomeArqParam As String, ByVal objReprocessamentoEst As ClassReprocessamentoEST) As Long
'reprocessa os movimentos de estoque a partir da data passada como parametro

Dim lErro As Long

On Error GoTo Erro_Rotina_Reproc_MovEstoque

'*** Para depurar, usando o BatchEst como .dll, o trecho abaixo deve estar comentado
    TelaAcompanhaBatchEST.Caption = TITULO_TELABATCH_REPROCESSAMENTO_MOVEST
    TelaAcompanhaBatchEST.sNomeArqParam = sNomeArqParam
    TelaAcompanhaBatchEST.iRotinaBatch = ROTINA_REPROCESSAMENTO_MOVEST_BATCH

    'Parametros da rotina
    Set TelaAcompanhaBatchEST.objReprocessamentoEst = objReprocessamentoEst
    'Exibe a tela de acompanhamento do batch
    TelaAcompanhaBatchEST.Show
'***

    '*** Para depurar, usando o BatchEst como .dll, o trecho abaixo deve estar descomentado
'    lErro = Rotina_Reprocessamento_MovEstoque_Int(objReprocessamentoEst)
'    If lErro <> SUCESSO Then gError 83543
    '***
    
    Rotina_Reproc_MovEstoque = SUCESSO

    Exit Function

Erro_Rotina_Reproc_MovEstoque:

    Rotina_Reproc_MovEstoque = gErr

    Select Case gErr
        
        Case 83543
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144651)

    End Select

    Exit Function

End Function
'
'Function Rotina_Reprocessamento_MovEstoque(sNomeArqParam As String, objReprocessamentoEst As ClassReprocessamentoEST) As Long
''reprocessa os movimentos de estoque a partir da data passada como parametro
'
'Dim lErro As Long
'
'On Error GoTo Erro_Rotina_Reprocessamento_MovEstoque
'
''*** Luiz, retirar o comentário do trecho abaixo ***
''    TelaAcompanhaBatchEST.Caption = TITULO_TELABATCH_REPROCESSAMENTO_MOVEST
''    TelaAcompanhaBatchEST.sNomeArqParam = sNomeArqParam
''    TelaAcompanhaBatchEST.iRotinaBatch = ROTINA_REPROCESSAMENTO_MOVEST_BATCH
''
''    'Parametros da rotina
''    Set TelaAcompanhaBatchEST.objReprocessamentoEST = objReprocessamentoEST
''    'Exibe a tela de acompanhamento do batch
''    TelaAcompanhaBatchEST.Show
''*** Fim do trecho que deve ser "descomentado" ***
'
'    '*** Luiz, retirar o trecho de código abaixo
'    lErro = Rotina_Reprocessamento_MovEstoque_Int(objReprocessamentoEst)
'    If lErro <> SUCESSO Then gError 83543
'    '*** Fim do trecho a ser retirado ***
'
'    Rotina_Reprocessamento_MovEstoque = SUCESSO
'
'    Exit Function
'
'Erro_Rotina_Reprocessamento_MovEstoque:
'
'    Rotina_Reprocessamento_MovEstoque = gErr
'
'    Select Case gErr
'
'        Case 83543
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144652)
'
'    End Select
'
'    Exit Function
'
'End Function

Function Rotina_Reproc_TestaIntegridade(ByVal sNomeArqParam As String, ByVal objReprocessamentoEst As ClassReprocessamentoEST) As Long
'Testa a integridade do último reprocessamento aplicado sobre o BD

Dim lErro As Long

On Error GoTo Erro_Rotina_Reproc_TestaIntegridade

'*** Luiz, retirar o comentário do trecho abaixo ***
'    TelaAcompanhaBatchEST.Caption = TITULO_TELABATCH_REPROCESSAMENTO_TESTAINT
'    TelaAcompanhaBatchEST.sNomeArqParam = sNomeArqParam
'    TelaAcompanhaBatchEST.iRotinaBatch = ROTINA_REPROCESSAMENTO_TESTAINT_BATCH
'
'    'Parametros da rotina
'    Set TelaAcompanhaBatchEST.objReprocessamentoEST = objReprocessamentoEST
'    'Exibe a tela de acompanhamento do batch
'    TelaAcompanhaBatchEST.Show
'*** Fim do trecho que deve ser "descomentado" ***

    '*** Luiz, retirar o trecho de código abaixo
'    lErro = Rotina_Reproc_TestaIntegridade_Int(objReprocessamentoEst)
'    If lErro <> SUCESSO Then gError 99999
    '*** Fim do trecho a ser retirado ***
    
    Rotina_Reproc_TestaIntegridade = SUCESSO

    Exit Function

Erro_Rotina_Reproc_TestaIntegridade:

    Rotina_Reproc_TestaIntegridade = gErr

    Select Case gErr
        
        Case 83543
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144653)

    End Select

    Exit Function

End Function


'###############################################
'INSERIDO POR WAGNER
Public Function Rotina_NFiscalContrato_Gera(ByVal sNomeArqParam As String, ByVal objGeracaoFatContrato As ClassGeracaoFatContrato) As Long

Dim lErro As Long

On Error GoTo Erro_Rotina_NFiscalContrato_Gera

    TelaAcompanhaBatchEST.Caption = TITULO_TELABATCH_GERACONTRATOCOBRANCA
    TelaAcompanhaBatchEST.sNomeArqParam = sNomeArqParam
    TelaAcompanhaBatchEST.iRotinaBatch = ROTINA_GERACONTRATOCOBRANCA_BATCH

    'Parametros da rotina
    Set TelaAcompanhaBatchEST.objGeracaoFatContrato = objGeracaoFatContrato

    'Exibe a tela de acompanhamento do batch
    TelaAcompanhaBatchEST.Show

    Rotina_NFiscalContrato_Gera = SUCESSO

    Exit Function

Erro_Rotina_NFiscalContrato_Gera:

    Rotina_NFiscalContrato_Gera = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144654)

    End Select

    Exit Function

End Function
'INSERIDO POR WAGNER
'###############################################

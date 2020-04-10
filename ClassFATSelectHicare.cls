VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFATSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function ContratoPropaganda_Le(ByVal objContratoPropaganda As ClassContratoPropag) As Long
'Lê do BD o ContratoPropaganda com as informações passadas no Obj

Dim lErro As Long
Dim lComando As Long
Dim tContratoPropaganda As typeContratoPropaganda

On Error GoTo Erro_ContratoPropaganda_Le

    'Abre Comando
    lComando = Comando_Abrir
    If lComando = 0 Then gError 128080
    
    'Realiza a seleção do ContratoPropaganda com a condição passada
    lErro = Comando_Executar(lComando, "SELECT Cliente, FilialEmpresa, PeriodoDe, PeriodoAte, Percentual FROM ContratoPropaganda WHERE Cliente = ? AND FilialEmpresa = ? AND PeriodoDe = ? AND PeriodoAte = ?", tContratoPropaganda.lCliente, tContratoPropaganda.iFilialEmpresa, tContratoPropaganda.dtPeriodoDe, tContratoPropaganda.dtPeriodoAte, tContratoPropaganda.dPercentual, objContratoPropaganda.lCliente, objContratoPropaganda.iFilialEmpresa, objContratoPropaganda.dtPeriodoDe, objContratoPropaganda.dtPeriodoAte)
    If lErro <> AD_SQL_SUCESSO Then gError 128081
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And AD_SQL_SEM_DADOS Then gError 128082
    
    'ContratoPropaganda não cadastrado --> ERRO
    If lErro = AD_SQL_SEM_DADOS Then gError 128083
    
    'Preenche o Obj
    With tContratoPropaganda
    
        objContratoPropaganda.lCliente = .lCliente
        objContratoPropaganda.iFilialEmpresa = .iFilialEmpresa
        objContratoPropaganda.dtPeriodoDe = .dtPeriodoDe
        objContratoPropaganda.dtPeriodoAte = .dtPeriodoAte
        objContratoPropaganda.dPercentual = .dPercentual
        
    End With
        
    'Fecha o Comando
    Call Comando_Fechar(lComando)
        
    ContratoPropaganda_Le = SUCESSO
    
    Exit Function
    
Erro_ContratoPropaganda_Le:

    ContratoPropaganda_Le = gErr
    
    Select Case gErr
    
        Case 128080
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 128081, 128082
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTRATOPROPAGANDA", gErr, objContratoPropaganda.lCliente, objContratoPropaganda.dtPeriodoDe, objContratoPropaganda.dtPeriodoAte)
        
        Case 128083
            'registro não encontrado no banco
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179190)
            
    End Select
    
    Call Comando_Fechar(lComando)
        
    Exit Function
        
End Function

Public Function CamposGenericos_ValidaExclusao_Situacao() As Long
'Verifica se é válida a exclusão de valores para o campo Requisitante

Dim lComando As Long
Dim lErro As Long
Dim lSituacao As Long
Dim lNumero As Long
Dim lFornecedor As Long
Dim iFilial As Integer

On Error GoTo Erro_CamposGenericos_ValidaExclusao_Situacao
    
    'Abre comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 132201
    
    'Busca Situacoes em NFsPad_Todas que não existam na tabela CamposGenericos Valores
    lErro = Comando_Executar(lComando, "SELECT Situacao, Fornecedor, Filial, NumNotaFiscal FROM NFsPag_Todas WHERE Situacao NOT IN (SELECT CodValor FROM CamposGenericosValores WHERE CodCampo = ?) AND Situacao <> 0", lSituacao, lNumero, lFornecedor, iFilial, CAMPOSGENERICOS_SITUACAO)
    If lErro <> AD_SQL_SUCESSO Then gError 132202
    
    'Posiciona no primeiro registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 132203
    
    'Se encontrou algum registro => erro de inconsistência
    If lErro = AD_SQL_SUCESSO Then gError 132204
    
    'Busca Situacoes em TitulosPagTodos que não existam na tabela CamposGenericos Valores
    lErro = Comando_Executar(lComando, "SELECT Situacao, Fornecedor, Filial, NumTitulo FROM TitulosPag_Todos WHERE Situacao NOT IN (SELECT CodValor FROM CamposGenericosValores WHERE CodCampo = ?) AND Situacao <> 0", lSituacao, lNumero, lFornecedor, iFilial, CAMPOSGENERICOS_SITUACAO)
    If lErro <> AD_SQL_SUCESSO Then gError 132205
    
    'Posiciona no primeiro registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 132206
    
    'Se encontrou algum registro => erro de inconsistência
    If lErro = AD_SQL_SUCESSO Then gError 132207
    
    'Fecha comando
    Call Comando_Fechar(lComando)
    
    CamposGenericos_ValidaExclusao_Situacao = SUCESSO
    
    Exit Function
    
Erro_CamposGenericos_ValidaExclusao_Situacao:

    CamposGenericos_ValidaExclusao_Situacao = gErr
    
    Select Case gErr
    
        Case 132201
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 132202, 132203
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFPAG_CAMPOSGENERICOS", gErr)
        
        Case 132204
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_SITUACAO_NFPAG", gErr, lSituacao, lNumero, lFornecedor, iFilial)
            
        Case 132205, 132206
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOSPAG_CAMPOSGENERICOS", gErr)
        
        Case 132207
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_SITUACAO_TITULOSPAG", gErr, lSituacao, lNumero, lFornecedor, iFilial)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179193)
    
    End Select

    'Fecha comando
    Call Comando_Fechar(lComando)

End Function

Public Function CamposGenericos_ValidaExclusao_Historico() As Long
'Verifica se é válida a exclusão de valores para o campo Requisitante

Dim lComando As Long
Dim lErro As Long
Dim lHistorico As Long
Dim lNumero As Long
Dim lFornecedor As Long
Dim iFilial As Integer

On Error GoTo Erro_CamposGenericos_ValidaExclusao_Historico
    
    'Abre comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 132208
    
    'Busca Históricos em NFsPad_Todas que não existam na tabela CamposGenericos Valores
    lErro = Comando_Executar(lComando, "SELECT Historico, Fornecedor, Filial, NumNotaFiscal FROM NFsPag_Todas WHERE Historico NOT IN (SELECT CodValor FROM CamposGenericosValores WHERE CodCampo = ?) AND Historico <> 0", lHistorico, lNumero, lFornecedor, iFilial, CAMPOSGENERICOS_HISTORICO)
    If lErro <> AD_SQL_SUCESSO Then gError 132209
    
    'Posiciona no primeiro registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 132210
    
    'Se encontrou algum registro => erro de inconsistência
    If lErro = AD_SQL_SUCESSO Then gError 132211
    
    'Busca Históricos em TitulosPagTodos que não existam na tabela CamposGenericos Valores
    lErro = Comando_Executar(lComando, "SELECT Historico, Fornecedor, Filial, NumTitulo FROM TitulosPag_Todos WHERE Historico NOT IN (SELECT CodValor FROM CamposGenericosValores WHERE CodCampo = ?) AND Historico <> 0", lHistorico, lNumero, lFornecedor, iFilial, CAMPOSGENERICOS_HISTORICO)
    If lErro <> AD_SQL_SUCESSO Then gError 132212
    
    'Posiciona no primeiro registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 132213
    
    'Se encontrou algum registro => erro de inconsistência
    If lErro = AD_SQL_SUCESSO Then gError 132214
    
    'Fecha comando
    Call Comando_Fechar(lComando)
    
    CamposGenericos_ValidaExclusao_Historico = SUCESSO
    
    Exit Function
    
Erro_CamposGenericos_ValidaExclusao_Historico:

    CamposGenericos_ValidaExclusao_Historico = gErr
    
    Select Case gErr
    
        Case 132208
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 132209, 132210
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFPAG_CAMPOSGENERICOS", gErr)
        
        Case 132211
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_HISTORICO_NFPAG", gErr, lHistorico, lNumero, lFornecedor, iFilial)
            
        Case 132212, 132213
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOSPAG_CAMPOSGENERICOS", gErr)
        
        Case 132214
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_HISTORICO_TITULOSPAG", gErr, lHistorico, lNumero, lFornecedor, iFilial)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179194)
    
    End Select

    'Fecha comando
    Call Comando_Fechar(lComando)

End Function

Public Function CamposGenericos_ValidaExclusao_SubConta() As Long
'Verifica se é válida a exclusão de valores para o campo Requisitante

Dim lComando As Long
Dim lErro As Long
Dim lSubConta As Long
Dim lNumero As Long
Dim lFornecedor As Long
Dim iFilial As Integer

On Error GoTo Erro_CamposGenericos_ValidaExclusao_SubConta
    
    'Abre comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 132215
    
    'Busca Situacoes em NFsPad_Todas que não existam na tabela CamposGenericos Valores
    lErro = Comando_Executar(lComando, "SELECT SubConta, Fornecedor, Filial, NumNotaFiscal FROM NFsPag_Todas WHERE SubConta NOT IN (SELECT CodValor FROM CamposGenericosValores WHERE CodCampo = ?) AND SubConta <> 0", lSubConta, lNumero, lFornecedor, iFilial, CAMPOSGENERICOS_SUBCONTA)
    If lErro <> AD_SQL_SUCESSO Then gError 132216
    
    'Posiciona no primeiro registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 132217
    
    'Se encontrou algum registro => erro de inconsistência
    If lErro = AD_SQL_SUCESSO Then gError 132218
    
    'Busca Situacoes em TitulosPagTodos que não existam na tabela CamposGenericos Valores
    lErro = Comando_Executar(lComando, "SELECT SubConta, Fornecedor, Filial, NumTitulo FROM TitulosPag_Todos WHERE SubConta NOT IN (SELECT CodValor FROM CamposGenericosValores WHERE CodCampo = ?) AND SubConta <> 0", lSubConta, lNumero, lFornecedor, iFilial, CAMPOSGENERICOS_SUBCONTA)
    If lErro <> AD_SQL_SUCESSO Then gError 132219
    
    'Posiciona no primeiro registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 132220
    
    'Se encontrou algum registro => erro de inconsistência
    If lErro = AD_SQL_SUCESSO Then gError 132221
    
    'Fecha comando
    Call Comando_Fechar(lComando)
    
    CamposGenericos_ValidaExclusao_SubConta = SUCESSO
    
    Exit Function
    
Erro_CamposGenericos_ValidaExclusao_SubConta:

    CamposGenericos_ValidaExclusao_SubConta = gErr
    
    Select Case gErr
    
        Case 132215
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 132216, 132217
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFPAG_CAMPOSGENERICOS", gErr)
        
        Case 132218
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_SUBCONTA_NFPAG", gErr, lSubConta, lNumero, lFornecedor, iFilial)
            
        Case 132219, 132220
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOSPAG_CAMPOSGENERICOS", gErr)
        
        Case 132221
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_SUBCONTA_TITULOSPAG", gErr, lSubConta, lNumero, lFornecedor, iFilial)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179195)
    
    End Select

    'Fecha comando
    Call Comando_Fechar(lComando)

End Function

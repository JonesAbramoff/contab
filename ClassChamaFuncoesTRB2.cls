VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassChamaFuncoesTRB"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function ICMSAliqExternas_Le_EstadoOrigem(sSiglaOrigem As String, colAliqExternas As Collection) As Long
'carrega as aliquotas externas que partem de um estado (sSiglaOrigem)

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ICMSAliqExternas_Le_EstadoOrigem
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("TRB_ICMSALIQEXTERNAS_LE_ESTADOORIGEM", X)
    If lErro <> SUCESSO Then Error 24593
    
    
    ICMSAliqExternas_Le_EstadoOrigem = X.ICMSAliqExternas_Le_EstadoOrigem(sSiglaOrigem, colAliqExternas)
    
    Exit Function
    
Erro_ICMSAliqExternas_Le_EstadoOrigem:

    ICMSAliqExternas_Le_EstadoOrigem = Err
    
    Select Case Err

        Case 24593
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146530)
        
    End Select
    
    Exit Function
    
End Function

Function ICMSExcecoes_Le_Estado(sEstado As String, colExcecoes As Collection) As Long
'preenche colExcecoes com objetos ClassICMSExcecoes de um determinado estado, ordenandos pelo campo "prioridade"

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ICMSExcecoes_Le_Estado
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("TRB_ICMSEXCECOES_LE_ESTADO", X)
    If lErro <> SUCESSO Then Error 24594
    
    
    ICMSExcecoes_Le_Estado = X.ICMSExcecoes_Le_Estado(sEstado, colExcecoes)
    
    Exit Function
    
Erro_ICMSExcecoes_Le_Estado:

    ICMSExcecoes_Le_Estado = Err
    
    Select Case Err

        Case 24594
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146531)
        
    End Select
    
    Exit Function
    
End Function

Function TiposTribICMS_Le_Todos(colTiposTribICMS As Collection) As Long
'preenche colTiposTribICMS com objetos ClassTipoTribICMS a partir da tabela TiposTribICMS

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TiposTribICMS_Le_Todos
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("TRB_TIPOSTRIBICMS_LE_TODOS", X)
    If lErro <> SUCESSO Then Error 24595
    
    
    TiposTribICMS_Le_Todos = X.TiposTribICMS_Le_Todos(colTiposTribICMS)
    
    Exit Function
    
Erro_TiposTribICMS_Le_Todos:

    TiposTribICMS_Le_Todos = Err
    
    Select Case Err

        Case 24595
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146532)
        
    End Select
    
    Exit Function
    
End Function

Function TipoTribICMS_Le(objTipoTribICMS As ClassTipoTribICMS)
'Lê o item da tabela TipoTributacaoICMS com o Tipo passado em objTipoTributacaoICMS

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TipoTribICMS_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("TRB_TIPOTRIBICMS_LE", X)
    If lErro <> SUCESSO Then Error 24885
    
    
    TipoTribICMS_Le = X.TipoTribICMS_Le(objTipoTribICMS)
    
    Exit Function
    
Erro_TipoTribICMS_Le:

    TipoTribICMS_Le = Err
    
    Select Case Err

        Case 24885
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146533)
        
    End Select
    
    Exit Function
    
End Function

Function TipoTribIPI_Le(objTipoTribIPI As ClassTipoTribIPI)
'Lê o item da tabela TipoTributacaoIPI com o Tipo passado em objTipoTributacaoIPI

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TipoTribIPI_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("TRB_TIPOTRIBIPI_LE", X)
    If lErro <> SUCESSO Then Error 24886
    
    
    TipoTribIPI_Le = X.TipoTribIPI_Le(objTipoTribIPI)
    
    Exit Function
    
Erro_TipoTribIPI_Le:

    TipoTribIPI_Le = Err
    
    Select Case Err

        Case 24886
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146534)
        
    End Select
    
    Exit Function
    
End Function

Function TiposTribIPI_Le_Todos(colTiposTribIPI As Collection) As Long
'preenche colTiposTribIPI com objetos ClassTipoTribIPI a partir da tabela TiposTribIPI

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TiposTribIPI_Le_Todos
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("TRB_TIPOSTRIBIPI_LE_TODOS", X)
    If lErro <> SUCESSO Then Error 27588
    
    
    TiposTribIPI_Le_Todos = X.TiposTribIPI_Le_Todos(colTiposTribIPI)
    
    Exit Function
    
Erro_TiposTribIPI_Le_Todos:

    TiposTribIPI_Le_Todos = Err
    
    Select Case Err

        Case 27588
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146535)
        
    End Select
    
    Exit Function
    
End Function

Function IPIExcecoes_Le_Todas(colExcecoes As Collection) As Long
'preenche colExcecoes com objetos ClassIPIExcecoes, ordenandos pelo campo "prioridade"
Dim X As Object, lErro As Long
    
On Error GoTo Erro_IPIExcecoes_Le_Todas
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("TRB_IPIEXCECOES_LE_TODAS", X)
    If lErro <> SUCESSO Then Error 27589
    
    
    IPIExcecoes_Le_Todas = X.IPIExcecoes_Le_Todas(colExcecoes)
    
    Exit Function
    
Erro_IPIExcecoes_Le_Todas:

    IPIExcecoes_Le_Todas = Err
    
    Select Case Err

        Case 27589
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146536)
        
    End Select
    
    Exit Function
    
End Function

Function IPIExcecoes_Le_FilialCliente(lCliente As Long, iFilial As Integer, colExcecoes As Collection) As Long
'preenche colExcecoes com objetos ClassIPIExcecoes validos para a filial cliente, ordenandos pelo campo "prioridade"
Dim X As Object, lErro As Long
    
On Error GoTo Erro_IPIExcecoes_Le_FilialCliente
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("TRB_IPIEXCECOES_LE_FILIALCLIENTE", X)
    If lErro <> SUCESSO Then Error 27590
    
    
    IPIExcecoes_Le_FilialCliente = X.IPIExcecoes_Le_FilialCliente(lCliente, iFilial, colExcecoes)
    
    Exit Function
    
Erro_IPIExcecoes_Le_FilialCliente:

    IPIExcecoes_Le_FilialCliente = Err
    
    Select Case Err

        Case 27590
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146537)
        
    End Select
    
    Exit Function
    
End Function

Function TipoTributacao_Le(objTributacaoTipo As ClassTipoDeTributacaoMovto) As Long
'Preenche objTributacaoTipo a partir do bd p/registro identificado por objTributacaoTipo.iTipo
'Retorna 27259 se nao encontrar
Dim X As Object, lErro As Long
    
On Error GoTo Erro_TipoTributacao_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("TRB_TIPOTRIBUTACAO_LE", X)
    If lErro <> SUCESSO Then Error 27669
    
    
    TipoTributacao_Le = X.TipoTributacao_Le(objTributacaoTipo)
    
    Exit Function
    
Erro_TipoTributacao_Le:

    TipoTributacao_Le = Err
    
    Select Case Err

        Case 27669
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146538)
        
    End Select
    
    Exit Function
    
End Function

Function ICMSAliqExternas_Le_EstadoDestino(sSiglaDestino As String, colAliqExternas As Collection) As Long
'carrega as aliquotas externas que chegam a um estado (sSiglaDestino)

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ICMSAliqExternas_Le_EstadoDestino
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("TRB_ICMSALIQEXTERNAS_LE_ESTADODEST", X)
    If lErro <> SUCESSO Then Error 27875
    
    
    ICMSAliqExternas_Le_EstadoDestino = X.ICMSAliqExternas_Le_EstadoDestino(sSiglaDestino, colAliqExternas)
    
    Exit Function
    
Erro_ICMSAliqExternas_Le_EstadoDestino:

    ICMSAliqExternas_Le_EstadoDestino = Err
    
    Select Case Err

        Case 27875
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146539)
        
    End Select
    
    Exit Function
    
End Function

Function ICMSExcecao_Le(objICMSExcecoes As ClassICMSExcecao)
'Lê um item da tabela ICMSExcecoes, a partir dos dados passados em objICMSExcecoes

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ICMSExcecao_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("TRB_ICMSEXCECAO_LE", X)
    If lErro <> SUCESSO Then Error 16933
    
    
    ICMSExcecao_Le = X.ICMSExcecao_Le(objICMSExcecoes)
    
    Exit Function
    
Erro_ICMSExcecao_Le:

    ICMSExcecao_Le = Err
    
    Select Case Err

        Case 16933
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146540)
        
    End Select
    
    Exit Function
    
End Function

Function ICMSExcecao_Grava(objICMSExcecoes As ClassICMSExcecao) As Long
'grava as alterações/inclusões realizadas pelo usuário nos dados de ICMSExcecoes para um registro.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ICMSExcecao_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("TRB_ICMSEXCECAO_GRAVA", X)
    If lErro <> SUCESSO Then Error 16934
    
    
    ICMSExcecao_Grava = X.ICMSExcecao_Grava(objICMSExcecoes)
    
    Exit Function
    
Erro_ICMSExcecao_Grava:

    ICMSExcecao_Grava = Err
    
    Select Case Err

        Case 16934
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146541)
        
    End Select
    
    Exit Function
    
End Function

Function ICMSExcecao_Exclui(objICMSExcecoes As ClassICMSExcecao) As Long
'grava as exclusões realizadas pelo usuário nos dados de ICMSExcecoes para um registro.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ICMSExcecao_Exclui
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("TRB_ICMSEXCECAO_EXCLUI", X)
    If lErro <> SUCESSO Then Error 16935
    
    
    ICMSExcecao_Exclui = X.ICMSExcecao_Exclui(objICMSExcecoes)
    
    Exit Function
    
Erro_ICMSExcecao_Exclui:

    ICMSExcecao_Exclui = Err
    
    Select Case Err

        Case 16935
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146542)
        
    End Select
    
    Exit Function
    
End Function

Function TipoTributacao_Exclui(objTipoTributacao As ClassTipoDeTributacaoMovto) As Long
'Exclui um tipo de tributação paassado como parâmetro
'objTipoDeTributacao.iTipo tem que estar preenchido

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TipoTributacao_Exclui
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("TRB_TIPOTRIBUTACAO_EXCLUI", X)
    If lErro <> SUCESSO Then Error 19741
    
    
    TipoTributacao_Exclui = X.TipoTributacao_Exclui(objTipoTributacao)
    
    Exit Function
    
Erro_TipoTributacao_Exclui:

    TipoTributacao_Exclui = Err
    
    Select Case Err

        Case 19741
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146543)
        
    End Select
    
    Exit Function
    
End Function

Function TipoTributacao_Grava(objTipoTributacao As ClassTipoDeTributacaoMovto) As Long
'Inclui ou altera um tipo de tributação

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TipoTributacao_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("TRB_TIPOTRIBUTACAO_GRAVA", X)
    If lErro <> SUCESSO Then Error 19742
    
    
    TipoTributacao_Grava = X.TipoTributacao_Grava(objTipoTributacao)
    
    Exit Function
    
Erro_TipoTributacao_Grava:

    TipoTributacao_Grava = Err
    
    Select Case Err

        Case 19742
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146544)
        
    End Select
    
    Exit Function
    
End Function

Function PadraoTribEntrada_Grava(objPadraoTribEnt As ClassPadraoTribEnt) As Long
'Inclui ou Altera um Padrão de Tributação Entrada

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PadraoTribEntrada_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("TRB_PADRAOTRIBENTRADA_GRAVA", X)
    If lErro <> SUCESSO Then Error 19756
    
    
    PadraoTribEntrada_Grava = X.PadraoTribEntrada_Grava(objPadraoTribEnt)
    
    Exit Function
    
Erro_PadraoTribEntrada_Grava:

    PadraoTribEntrada_Grava = Err
    
    Select Case Err

        Case 19756
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146545)
        
    End Select
    
    Exit Function
    
End Function

Function PadraoTribEntrada_Exclui(objPadraoTribEnt As ClassPadraoTribEnt) As Long
'Exclui um Padrão de Tributação Entrada existente

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PadraoTribEntrada_Exclui
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("TRB_PADRAOTRIBENTRADA_EXCLUI", X)
    If lErro <> SUCESSO Then Error 19757
    
    
    PadraoTribEntrada_Exclui = X.PadraoTribEntrada_Exclui(objPadraoTribEnt)
    
    Exit Function
    
Erro_PadraoTribEntrada_Exclui:

    PadraoTribEntrada_Exclui = Err
    
    Select Case Err

        Case 19757
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146546)
        
    End Select
    
    Exit Function
    
End Function

Function TipoTributacao_Lock(iTipoTributacaoPadrao As Integer, lComando As Long) As Long
'faz lock shared de um tipo de tributacao para garantir a sua existencia

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TipoTributacao_Lock
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("TRB_TIPOTRIBUTACAO_LOCK", X)
    If lErro <> SUCESSO Then Error 19758
    
    
    TipoTributacao_Lock = X.TipoTributacao_Lock(iTipoTributacaoPadrao, lComando)
    
    Exit Function
    
Erro_TipoTributacao_Lock:

    TipoTributacao_Lock = Err
    
    Select Case Err

        Case 19758
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146547)
        
    End Select
    
    Exit Function
    
End Function

Function PadraoTribEntrada_Le(objPadraoTribEnt As ClassPadraoTribEnt) As Long
'Lê o Padrão de Tributação Entrada passado como parâmetro em objPadraoTribEnt

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PadraoTribEntrada_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("TRB_PADRAOTRIBENTRADA_LE", X)
    If lErro <> SUCESSO Then Error 19759
    
    
    PadraoTribEntrada_Le = X.PadraoTribEntrada_Le(objPadraoTribEnt)
    
    Exit Function
    
Erro_PadraoTribEntrada_Le:

    PadraoTribEntrada_Le = Err
    
    Select Case Err

        Case 19759
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146548)
        
    End Select
    
    Exit Function
    
End Function

Function PadraoTribSaida_Le(objPadraoTribSaida As ClassPadraoTribSaida) As Long
'Lê o Padrão de Tributação Saída passado como parâmetro em objPadraoTribSaida

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PadraoTribSaida_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("TRB_PADRAOTRIBSAIDA_LE", X)
    If lErro <> SUCESSO Then Error 19760
    
    
    PadraoTribSaida_Le = X.PadraoTribSaida_Le(objPadraoTribSaida)
    
    Exit Function
    
Erro_PadraoTribSaida_Le:

    PadraoTribSaida_Le = Err
    
    Select Case Err

        Case 19760
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146549)
        
    End Select
    
    Exit Function
    
End Function

Function PadraoTribSaida_Grava(objPadraoTribSaida As ClassPadraoTribSaida) As Long
'Inclui ou Altera um Padrão de Tributação Saída

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PadraoTribSaida_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("TRB_PADRAOTRIBSAIDA_GRAVA", X)
    If lErro <> SUCESSO Then Error 19761
    
    
    PadraoTribSaida_Grava = X.PadraoTribSaida_Grava(objPadraoTribSaida)
    
    Exit Function
    
Erro_PadraoTribSaida_Grava:

    PadraoTribSaida_Grava = Err
    
    Select Case Err

        Case 19761
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146550)
        
    End Select
    
    Exit Function
    
End Function

Function PadraoTribSaida_Exclui(objPadraoTribSaida As ClassPadraoTribSaida) As Long
'Exclui um Padrão de Tributação Saída existente

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PadraoTribSaida_Exclui
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("TRB_PADRAOTRIBSAIDA_EXCLUI", X)
    If lErro <> SUCESSO Then Error 19762
    
    
    PadraoTribSaida_Exclui = X.PadraoTribSaida_Exclui(objPadraoTribSaida)
    
    Exit Function
    
Erro_PadraoTribSaida_Exclui:

    PadraoTribSaida_Exclui = Err
    
    Select Case Err

        Case 19762
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146551)
        
    End Select
    
    Exit Function
    
End Function

Function IPIExcecao_Le(objIPIExcecoes As ClassIPIExcecao)
'Lê um item da tabela IPIExcecoes, a partir dos dados passados em objIPIExcecoes

Dim X As Object, lErro As Long
    
On Error GoTo Erro_IPIExcecao_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("TRB_IPIEXCECAO_LE", X)
    If lErro <> SUCESSO Then Error 45579
    
    
    IPIExcecao_Le = X.IPIExcecao_Le(objIPIExcecoes)
    
    Exit Function
    
Erro_IPIExcecao_Le:

    IPIExcecao_Le = Err
    
    Select Case Err

        Case 45579
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146552)
        
    End Select
    
    Exit Function
    
End Function

Function IPIExcecao_Grava(objIPIExcecoes As ClassIPIExcecao) As Long
'grava as alterações/inclusões realizadas pelo usuário nos dados de IPIExcecoes para um registro.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_IPIExcecao_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("TRB_IPIEXCECAO_GRAVA", X)
    If lErro <> SUCESSO Then Error 45580
    
    
    IPIExcecao_Grava = X.IPIExcecao_Grava(objIPIExcecoes)
    
    Exit Function
    
Erro_IPIExcecao_Grava:

    IPIExcecao_Grava = Err
    
    Select Case Err

        Case 45580
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146553)
        
    End Select
    
    Exit Function
    
End Function

Function IPIExcecao_Exclui(objIPIExcecoes As ClassIPIExcecao) As Long
'grava as exclusões realizadas pelo usuário nos dados de IPIExcecoes para um registro.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_IPIExcecao_Exclui
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("TRB_IPIEXCECAO_EXCLUI", X)
    If lErro <> SUCESSO Then Error 45581
    
    
    IPIExcecao_Exclui = X.IPIExcecao_Exclui(objIPIExcecoes)
    
    Exit Function
    
Erro_IPIExcecao_Exclui:

    IPIExcecao_Exclui = Err
    
    Select Case Err

        Case 45581
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146554)
        
    End Select
    
    Exit Function
    
End Function


Function NFiscal_GravaTributacao(objNFiscal As ClassNFiscal, lComando1 As Long, lComando2 As Long) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscal_GravaTributacao
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_NFISCAL_GRAVATRIBUTACAO", X)
    If lErro <> SUCESSO Then Error 19683
    
    
    NFiscal_GravaTributacao = X.NFiscal_GravaTributacao(objNFiscal, lComando1, lComando2)

    Exit Function
    
Erro_NFiscal_GravaTributacao:
    
    NFiscal_GravaTributacao = Err
    
    Select Case Err

        Case 19683
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146555)
        
    End Select
    
    Exit Function

End Function

Function NFiscal_ExcluiTributacao(objNFiscal As ClassNFiscal) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscal_ExcluiTributacao
    
    lErro = Rotina_ObterObjetoFuncao("FAT_NFISCAL_EXCLUITRIBUTACAO", X)
    If lErro <> SUCESSO Then Error 19683
    
    NFiscal_ExcluiTributacao = X.NFiscal_ExcluiTributacao(objNFiscal)

    Exit Function
    
Erro_NFiscal_ExcluiTributacao:
    
    NFiscal_ExcluiTributacao = Err
    
    Select Case Err

        Case 19683
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146556)
        
    End Select
    
    Exit Function

End Function

Function NFiscal_GravaTributacaoCompl(objNFiscal As ClassNFiscal, objTribComplNF As ClassTribComplNF, lComando As Long) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscal_GravaTributacaoCompl
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_NFISCAL_GRAVATRIBUTACAOCOMPL", X)
    If lErro <> SUCESSO Then Error 19684
    
    
    NFiscal_GravaTributacaoCompl = X.NFiscal_GravaTributacaoCompl(objNFiscal, objTribComplNF, lComando)

    Exit Function
    
Erro_NFiscal_GravaTributacaoCompl:
    
    NFiscal_GravaTributacaoCompl = Err
    
    Select Case Err

        Case 19684
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146557)
        
    End Select
    
    Exit Function

End Function

Function NFiscal_GravaTributacaoItem(objItemNF As ClassItemNF, lComando As Long, lComandoProd As Long) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscal_GravaTributacaoItem
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_NFISCAL_GRAVATRIBUTACAOITEM", X)
    If lErro <> SUCESSO Then Error 19682
    
    
    NFiscal_GravaTributacaoItem = X.NFiscal_GravaTributacaoItem(objItemNF, lComando, lComandoProd)

    Exit Function
    
Erro_NFiscal_GravaTributacaoItem:
    
    NFiscal_GravaTributacaoItem = Err
    
    Select Case Err

        Case 19682
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146558)
        
    End Select
    
    Exit Function

End Function

Function ObterTribPadraoCli(iTipoTrib As Integer, sTipoDocSigla As String, sNatOp As String, lCliente As Long, iFilialCliente As Integer) As Long
'retorna em iTipoTrib o tipo de tributacao default p/o tipo do docto, natureza de operacao e filial cliente informados
'a pesquisa é feita na tabela PadroesTributacao, pegando um padrao correspondente a uma categoria associada a filialCliente
'Se nao achar, pode retornar um padrao generico, que independa do cliente. Se ainda assim nao achar nenhum registro retorna zero.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ObterTribPadraoCli
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("TRB_OBTERTRIBPADRAOCLI", X)
    If lErro <> SUCESSO Then Error 54815
    
    
    ObterTribPadraoCli = X.ObterTribPadraoCli(iTipoTrib, sTipoDocSigla, sNatOp, lCliente, iFilialCliente)

    Exit Function
    
Erro_ObterTribPadraoCli:
    
    ObterTribPadraoCli = Err
    
    Select Case Err

        Case 54815
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146559)
        
    End Select
    
    Exit Function

End Function

Function ObterTribPadraoProd(iTipoTrib As Integer, sTipoDocSigla As String, sNatOp As String, sProduto As String) As Long
'retorna em iTipoTrib o tipo de tributacao default p/o tipo do docto, natureza de operacao e produtos informados
'a pesquisa é feita na tabela PadroesTribEntrada, pegando um padrao correspondente a uma categoria associada ao produto
'Se nao achar, pode retornar um padrao generico, que independa do produto. Se ainda assim nao achar nenhum registro retorna zero.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ObterTribPadraoProd
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("TRB_OBTERTRIBPADRAOPROD", X)
    If lErro <> SUCESSO Then Error 54816
    
    
    ObterTribPadraoProd = X.ObterTribPadraoProd(iTipoTrib, sTipoDocSigla, sNatOp, sProduto)

    Exit Function
    
Erro_ObterTribPadraoProd:
    
    ObterTribPadraoProd = Err
    
    Select Case Err

        Case 54816
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146560)
        
    End Select
    
    Exit Function

End Function

Function NatOpPadrao_Le_Padrao(objNatOpPadrao As ClassNatOpPadrao) As Long
'obtem dados default p/tipo de operacao e tipo de atividade da empresa (comercio, industria,...)

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NatOpPadrao_Le_Padrao
    
    lErro = Rotina_ObterObjetoFuncao("TRB_NATOPPADRAO_LE_PADRAO", X)
    If lErro <> SUCESSO Then Error 65474
    
    NatOpPadrao_Le_Padrao = X.NatOpPadrao_Le_Padrao(objNatOpPadrao)

    Exit Function
    
Erro_NatOpPadrao_Le_Padrao:
    
    NatOpPadrao_Le_Padrao = Err
    
    Select Case Err

        Case 65474
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146561)
        
    End Select
    
    Exit Function

End Function

'-------------------------- VERSAO 2 ------------------------------------


Function NotaFiscal_Grava_Fis(objNFiscal As ClassNFiscal) As Long
'Grava nas Tabelas de Livros Fiscais a Nota Fiscal

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NotaFiscal_Grava_Fis
    
    lErro = Rotina_ObterObjetoFuncao("FIS_NOTAFISCAL_GRAVA_FIS", X)
    If lErro <> SUCESSO Then gError 69324
    
    NotaFiscal_Grava_Fis = X.NotaFiscal_Grava_Fis(objNFiscal)

    Exit Function
    
Erro_NotaFiscal_Grava_Fis:
    
    NotaFiscal_Grava_Fis = gErr
    
    Select Case gErr

        Case 69324
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 146562)
        
    End Select
    
    Exit Function

End Function

Function NotaFiscal_Cancela_Fis(objNFiscal As ClassNFiscal) As Long
'Cancela nas Tabelas de Livros Fiscais a Nota Fiscal

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NotaFiscal_Cancela_Fis
    
    lErro = Rotina_ObterObjetoFuncao("FIS_NOTAFISCAL_CANCELA_FIS", X)
    If lErro <> SUCESSO Then gError 88024
    
    NotaFiscal_Cancela_Fis = X.NotaFiscal_Cancela_Fis(objNFiscal)

    Exit Function
    
Erro_NotaFiscal_Cancela_Fis:
    
    NotaFiscal_Cancela_Fis = gErr
    
    Select Case gErr

        Case 88024
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 146563)
        
    End Select
    
    Exit Function

End Function

Function NotaFiscal_Exclui_Fis(objNFiscal As ClassNFiscal) As Long
'Exclui nas Tabelas de Livros Fiscais a Nota Fiscal

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NotaFiscal_Exclui_Fis
    
    lErro = Rotina_ObterObjetoFuncao("FIS_NOTAFISCAL_EXCLUI_FIS", X)
    If lErro <> SUCESSO Then gError 69324
    
    NotaFiscal_Exclui_Fis = X.NotaFiscal_Exclui_Fis(objNFiscal)

    Exit Function
    
Erro_NotaFiscal_Exclui_Fis:
    
    NotaFiscal_Exclui_Fis = gErr
    
    Select Case gErr

        Case 69324
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 146564)
        
    End Select
    
    Exit Function

End Function

Function NotaFiscal_Altera_Fis(objNFiscal As ClassNFiscal) As Long
'Trata a alteracao de dados de uma nf em relacao aos livros fiscais
'obs.:atualmente tratando apenas a alteracao das datas de emissao e de entrada

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NotaFiscal_Altera_Fis
    
    lErro = Rotina_ObterObjetoFuncao("FIS_NOTAFISCAL_ALTERA_FIS", X)
    If lErro <> SUCESSO Then gError 88033
    
    NotaFiscal_Altera_Fis = X.NotaFiscal_Altera_Fis(objNFiscal)

    Exit Function
    
Erro_NotaFiscal_Altera_Fis:
    
    NotaFiscal_Altera_Fis = gErr
    
    Select Case gErr

        Case 88033
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 146565)
        
    End Select
    
    Exit Function

End Function


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassESTGravaFox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function NFiscalEntradaCom_Grava(ByVal objNFiscal As ClassNFiscal, ByVal objContabil As ClassContabil, ByVal colPedCompras As Collection, ByVal colReqCompras As Collection) As Long
'Faz a Gravação de Uma Nota Fiscal de Entrada

Dim lErro As Long
Dim lTransacao As Long
Dim lNumNFOficial As Long

On Error GoTo Erro_NFiscalEntradaCom_Grava

    'Abre a Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 65915

    'FOX
    lErro = NFiscal_Grava_Clone(objNFiscal, objContabil, "NFiscalEntrada_Grava_EmTrans", lNumNFOficial)
    If lErro <> SUCESSO Then gError 500027
    'FIM FOX
    
    lErro = CF("NFiscalEntradaCom_Grava_EmTrans", objNFiscal, objContabil)
    If lErro <> SUCESSO Then gError 106565
        
    'FOX
    If lNumNFOficial <> 0 And lNumNFOficial <> objNFiscal.lNumNotaFiscal Then gError 500021
    'FIM FOX
    
    'Faz o Commit da Transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 65925

    NFiscalEntradaCom_Grava = SUCESSO

    Exit Function

Erro_NFiscalEntradaCom_Grava:

    NFiscalEntradaCom_Grava = gErr

    Select Case gErr

        Case 106565, 500021, 500027
        
        Case 65915
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 65925
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149604)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function

Function NFiscalFatEntradaCom_Grava(ByVal objNFiscal As ClassNFiscal, ByVal objContabil As ClassContabil, ByVal colPedCompras As Collection, ByVal colReqCompras As Collection) As Long
'Faz a Gravação de Uma Nota Fiscal de Entrada

Dim lErro As Long
Dim lTransacao As Long
Dim lNumNFOficial As Long

On Error GoTo Erro_NFiscalFatEntradaCom_Grava

    'Abre a Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 66500

    'FOX
    lErro = NFiscal_Grava_Clone(objNFiscal, objContabil, "NFiscalFatEntrada_Grava_EmTrans", lNumNFOficial)
    If lErro <> SUCESSO Then gError 500028
    'FIM FOX
    
    lErro = CF("NFiscalFatEntradaCom_Grava_EmTrans", objNFiscal, objContabil)
    If lErro <> SUCESSO Then gError 106569
        
    'FOX
    If lNumNFOficial <> 0 And lNumNFOficial <> objNFiscal.lNumNotaFiscal Then gError 500022
    'FIM FOX
    
    'Faz o Commit da Transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 66509

    NFiscalFatEntradaCom_Grava = SUCESSO

    Exit Function

Erro_NFiscalFatEntradaCom_Grava:

    NFiscalFatEntradaCom_Grava = gErr

    Select Case gErr

        Case 106569, 500022, 500028
        
        Case 66500
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 66509
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149605)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function

Function NFiscalEntrada_Grava(ByVal objNFiscal As ClassNFiscal, ByVal objContabil As ClassContabil) As Long
'Faz a Gravação de Uma Nota Fiscal de Entrada

Dim lErro As Long
Dim lTransacao As Long
Dim lNumNFOficial As Long

On Error GoTo Erro_NFiscalEntrada_Grava

    'Abre a Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 35043

    'FOX
    lErro = NFiscal_Grava_Clone(objNFiscal, objContabil, "NFiscalEntrada_Grava_EmTrans", lNumNFOficial)
    If lErro <> SUCESSO Then gError 500029
    'FIM FOX
    
    lErro = CF("NFiscalEntrada_Grava_EmTrans", objNFiscal, objContabil)
    If lErro <> SUCESSO Then gError 106564
        
    'FOX
    If lNumNFOficial <> 0 And lNumNFOficial <> objNFiscal.lNumNotaFiscal Then gError 500023
    'FIM FOX
    
    'Faz o Commit da Transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 35079

    NFiscalEntrada_Grava = SUCESSO

    Exit Function

Erro_NFiscalEntrada_Grava:

    NFiscalEntrada_Grava = gErr

    Select Case gErr

        Case 106564, 500023, 500029
        
        Case 35043
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 35079
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149606)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function

Function NFiscalFatEntrada_Grava(ByVal objNFiscal As ClassNFiscal, ByVal objContabil As ClassContabil) As Long

Dim lErro As Long
Dim lTransacao As Long
Dim lNumNFOficial As Long

On Error GoTo Erro_NFiscalFatEntrada_Grava

    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 42665
    
    'FOX
    lErro = NFiscal_Grava_Clone(objNFiscal, objContabil, "NFiscalFatEntrada_Grava_EmTrans", lNumNFOficial)
    If lErro <> SUCESSO Then gError 500030
    'FIM FOX
    
    lErro = CF("NFiscalFatEntrada_Grava_EmTrans", objNFiscal, objContabil)
    If lErro <> SUCESSO Then gError 106568
            
    'FOX
    If lNumNFOficial <> 0 And lNumNFOficial <> objNFiscal.lNumNotaFiscal Then gError 500024
    'FIM FOX
    
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 42684

    NFiscalFatEntrada_Grava = SUCESSO

    Exit Function

Erro_NFiscalFatEntrada_Grava:

    NFiscalFatEntrada_Grava = gErr
    
    Select Case gErr
    
        Case 106568, 500024, 500030
        
        Case 42665
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 42684
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149607)
            
    End Select

    Call Transacao_Rollback

    Exit Function
    
End Function

Function NFiscalEntradaDevolucao_Grava(ByVal objNFiscal As ClassNFiscal, ByVal objContabil As ClassContabil) As Long
'Faz a gravação de Uma Nota Fiscal de Entrada-Devolução

Dim lErro As Long
Dim lTransacao As Long
Dim lNumNFOficial As Long

On Error GoTo Erro_NFiscalEntradaDevolucao_Grava

    'Abre a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 35378

    'FOX
    lErro = NFiscal_Grava_Clone(objNFiscal, objContabil, "NFiscalEntradaDevolucao_Grava_EmTrans", lNumNFOficial)
    If lErro <> SUCESSO Then gError 500031
    'FIM FOX
    
    lErro = CF("NFiscalEntradaDevolucao_Grava_EmTrans", objNFiscal, objContabil)
    If lErro <> SUCESSO Then gError 106566
        
    'FOX
    If lNumNFOficial <> 0 And lNumNFOficial <> objNFiscal.lNumNotaFiscal Then gError 500025
    'FIM FOX
    
    'Faz o Commit na Transacao
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 35443

    NFiscalEntradaDevolucao_Grava = SUCESSO

    Exit Function

Erro_NFiscalEntradaDevolucao_Grava:

    NFiscalEntradaDevolucao_Grava = gErr

    Select Case gErr

        Case 106566, 500025, 500031
        
        Case 35378
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 35443
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149608)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function

Function NFiscalEntradaRemessa_Grava(ByVal objNFiscal As ClassNFiscal, ByVal objContabil As ClassContabil) As Long
'Faz as validações necessárias a gravação de Uma Nota Fiscal de Entrada-Remessa

Dim lErro As Long
Dim lTransacao As Long
Dim lNumNFOficial As Long

On Error GoTo Erro_NFiscalEntradaRemessa_Grava

    'Abre a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 35603
    
    'FOX
    lErro = NFiscal_Grava_Clone(objNFiscal, objContabil, "NFiscalEntradaRemessa_Grava_EmTrans", lNumNFOficial)
    If lErro <> SUCESSO Then gError 500032
    'FIM FOX
    
    lErro = CF("NFiscalEntradaRemessa_Grava_EmTrans", objNFiscal, objContabil)
    If lErro <> SUCESSO Then gError 106567
        
    'FOX
    If lNumNFOficial <> 0 And lNumNFOficial <> objNFiscal.lNumNotaFiscal Then gError 500026
    'FIM FOX
    
    'Faz o Commit na Transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 35624

    NFiscalEntradaRemessa_Grava = SUCESSO
    
    Exit Function

Erro_NFiscalEntradaRemessa_Grava:

    NFiscalEntradaRemessa_Grava = gErr

    Select Case gErr

        Case 106567, 500026, 500032
        
        Case 35603
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 35624
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149609)

    End Select

    Call Transacao_Rollback
    
    Exit Function

End Function

Function NFiscalEntrada_Verifica_Igual(ByVal objNFiscal As ClassNFiscal, ByVal iTipo As Integer, ByVal lComando As Long, ByVal objNFiscalBD As ClassNFiscal) As Long

Dim lErro As Long
Dim lNumIntDoc As Long
Dim vbMsgRes
Dim dtDataEntrada As Date

On Error GoTo Erro_NFiscalEntrada_Verifica_Igual

    If iTipo = DOCINFO_NF_EXTERNA Then

        'Busca Nota Fiscal com os mesmos dados contidos em objNFiscal
        If objNFiscal.iFilialEmpresa <= FOX_DELTA_FILIALREAL_OFICIAL Then
            lErro = Comando_Executar(lComando, "SELECT NumIntDoc, DataEntrada FROM TipoDocInfoNFiscal WHERE FilialEmpresa <= ? AND Fornecedor = ? AND FilialForn = ? AND Cliente = ? AND FilialCli = ? AND DataEmissao = ? AND Serie = ? AND NumNotaFiscal = ? AND Status <> ? AND Tipo = ? ", lNumIntDoc, dtDataEntrada, FOX_DELTA_FILIALREAL_OFICIAL, objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.lCliente, objNFiscal.iFilialCli, objNFiscal.dtDataEmissao, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, STATUS_EXCLUIDO, DOCINFO_NF_EXTERNA)
        Else
            lErro = Comando_Executar(lComando, "SELECT NumIntDoc, DataEntrada FROM TipoDocInfoNFiscal WHERE FilialEmpresa > ? AND Fornecedor = ? AND FilialForn = ? AND Cliente = ? AND FilialCli = ? AND DataEmissao = ? AND Serie = ? AND NumNotaFiscal = ? AND Status <> ? AND Tipo = ? ", lNumIntDoc, dtDataEntrada, FOX_DELTA_FILIALREAL_OFICIAL, objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.lCliente, objNFiscal.iFilialCli, objNFiscal.dtDataEmissao, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, STATUS_EXCLUIDO, DOCINFO_NF_EXTERNA)
        End If
        If lErro <> AD_SQL_SUCESSO Then Error 35382
    
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 35383
        
        If lErro = AD_SQL_SUCESSO Then
        
            objNFiscal.lNumIntDoc = lNumIntDoc
            objNFiscalBD.dtDataEntrada = dtDataEntrada
        
             Error 35384 'Se encontrar --> Erro

        End If

    ElseIf iTipo = DOCINFO_NF_INT_ENTRADA Then

        'Busca Nota Fiscal com os mesmos dados contidos em objNFiscal
        lErro = Comando_Executar(lComando, "SELECT NumIntDoc, DataEntrada FROM TipoDocInfoNFiscal WHERE FilialEmpresa = ? AND Serie = ? AND NumNotaFiscal = ? AND DataEmissao=? AND Status <> ? AND (Tipo = ? OR Tipo =?) ", lNumIntDoc, dtDataEntrada, objNFiscal.iFilialEmpresa, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.dtDataEmissao, STATUS_EXCLUIDO, DOCINFO_NF_INT_ENTRADA, DOCINFO_NF_INT_SAIDA)
        If lErro <> AD_SQL_SUCESSO Then Error 42079
    
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 42081
        
        If lErro = AD_SQL_SUCESSO Then
        
            objNFiscal.lNumIntDoc = lNumIntDoc
            
            objNFiscalBD.dtDataEntrada = dtDataEntrada
            
            Error 42083 'Se encontrar --> Erro

        End If

    End If

    NFiscalEntrada_Verifica_Igual = SUCESSO

    Exit Function

Erro_NFiscalEntrada_Verifica_Igual:

    NFiscalEntrada_Verifica_Igual = Err

    Select Case Err
    
        Case 35382, 35383, 42079, 42081
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL3", Err)

        Case 35384 'Externa Tratar erro de Nota Fiscal Igual
            
        Case 42083 'Interna Tratar erro de Nota Fiscal Igual
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149610)
    
    End Select
    
    Exit Function

End Function

Function EstoqueInicial_Grava1(ByVal objEstoqueProduto As ClassEstoqueProduto, ByVal iAlmoxarifadoPadrao As Integer, ByVal colRastreamento As Collection) As Long
'Grava o estoque inicial do produto

Dim lErro As Long

On Error GoTo Erro_EstoqueInicial_Grava1

    'grava o estoque inicial em transacao
    lErro = CF("EstoqueInicial_Grava_Trans", objEstoqueProduto, iAlmoxarifadoPadrao, colRastreamento)
    If lErro <> SUCESSO Then gError 105120
    
    'Se a filial é nao é de filial oficial entao
    If FilialEmpresa_ConvFRFO(objEstoqueProduto.iFilialEmpresa) <> objEstoqueProduto.iFilialEmpresa Then
    
        objEstoqueProduto.iAlmoxarifado = objEstoqueProduto.iAlmoxarifado - FOX_DELTA_FILIALREAL_OFICIAL
        objEstoqueProduto.iFilialEmpresa = objEstoqueProduto.iFilialEmpresa - FOX_DELTA_FILIALREAL_OFICIAL
    
        'grava o estoque inicial em transacao
        lErro = CF("EstoqueInicial_Grava_Trans", objEstoqueProduto, iAlmoxarifadoPadrao, colRastreamento)
        If lErro <> SUCESSO Then gError 105121
        
    End If
        
    EstoqueInicial_Grava1 = SUCESSO

    Exit Function

Erro_EstoqueInicial_Grava1:

    EstoqueInicial_Grava1 = gErr

    Select Case gErr

        Case 105120, 105121

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149611)

    End Select

    Exit Function

End Function

Function MovEstoque_ReqProducao_Grava0(ByVal objMovEstoque As ClassMovEstoque, ByVal objContabil As ClassContabil) As Long
'Grava o Movimento de Estoque de Requisição de Produção

Dim lErro As Long
Dim objItemMovEst As ClassItemMovEstoque
Dim iFilialEmpresa As Integer

On Error GoTo Erro_MovEstoque_ReqProducao_Grava0

    lErro = CF("MovEstoque_ReqProducao_Grava_Trans", objMovEstoque, objContabil)
    If lErro <> SUCESSO Then gError 105146
    
    'Se a filial é nao é de filial oficial entao
    If FilialEmpresa_ConvFRFO(objMovEstoque.iFilialEmpresa) <> objMovEstoque.iFilialEmpresa Then
    
        objMovEstoque.iFilialEmpresa = objMovEstoque.iFilialEmpresa - FOX_DELTA_FILIALREAL_OFICIAL
    
        For Each objItemMovEst In objMovEstoque.colItens
    
            objItemMovEst.iAlmoxarifado = objItemMovEst.iAlmoxarifado - FOX_DELTA_FILIALREAL_OFICIAL
            objItemMovEst.iFilialEmpresa = objMovEstoque.iFilialEmpresa
    
        Next
    
        iFilialEmpresa = giFilialEmpresa
        giFilialEmpresa = objMovEstoque.iFilialEmpresa
    
        lErro = CF("MovEstoque_ReqProducao_Grava_Trans", objMovEstoque, objContabil)
        If lErro <> SUCESSO Then gError 105147
    
        giFilialEmpresa = iFilialEmpresa
        
    
    End If
    
    MovEstoque_ReqProducao_Grava0 = SUCESSO

    Exit Function

Erro_MovEstoque_ReqProducao_Grava0:

    MovEstoque_ReqProducao_Grava0 = gErr

    Select Case gErr

        Case 105146
        
        Case 105147
            giFilialEmpresa = iFilialEmpresa

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149612)

    End Select

    Exit Function

End Function




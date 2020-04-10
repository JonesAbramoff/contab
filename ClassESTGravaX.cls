VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassESTGrava"
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
Dim lErro1 As Long
Dim lErro2 As Long
Dim vbMsgRes As VbMsgBoxResult
Dim lNumNotaFiscal As Long

On Error GoTo Erro_NFiscalEntradaCom_Grava

    lNumNotaFiscal = objNFiscal.lNumNotaFiscal

    'Abre a Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 65915

    'sport
    lErro1 = NFiscal_Grava_Clone(objNFiscal, objContabil, "NFiscalEntrada_Grava_EmTrans", lNumNFOficial)
'    If lErro <> SUCESSO Then gError 500027
    'FIM sport
    
     If lErro1 = SUCESSO Then
        lErro2 = CF("NFiscalEntradaCom_Grava_EmTrans", objNFiscal, objContabil, colPedCompras, colReqCompras)
    End If
   

    If lErro1 <> SUCESSO Or lErro2 <> SUCESSO Then
    
        'se era uma alteracao e deu erro na exclusao;insercao ==> tenta somente gravar os dados de contabilidade
        If lNumNotaFiscal <> 0 Then
    
            'Dá a Mensagem de aviso que só será gravado a contabilidade
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_ALTERACAO_NFISCAL_INTERNA_CONTAB", objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.dtDataEmissao)
            If vbMsgRes = vbNo Then gError 207819

            Call Transacao_Rollback
    
            lTransacao = Transacao_Abrir()
            If lTransacao = 0 Then gError 207820
    
            lErro = NFiscal_Altera_Clone(objNFiscal, objContabil, "NFiscalEntrada_AlteraDados_EmTrans", lNumNFOficial)
            If lErro <> SUCESSO Then gError 207821
        
            lErro = CF("NFiscalEntrada_AlteraDados_EmTrans", objNFiscal, objContabil)
            If lErro <> SUCESSO Then gError 207822
    
            'Faz o Commit da Transação
            lErro = Transacao_Commit()
            If lErro <> AD_SQL_SUCESSO Then gError 207823
    
        Else
            gError 207824
    
        End If
    
    Else
    
    
        'sport
        If lNumNFOficial <> 0 And lNumNFOficial <> objNFiscal.lNumNotaFiscal Then gError 500021
        'FIM sport
        
        'Faz o Commit da Transação
        lErro = Transacao_Commit()
        If lErro <> AD_SQL_SUCESSO Then gError 65925

    End If


    NFiscalEntradaCom_Grava = SUCESSO

    Exit Function

Erro_NFiscalEntradaCom_Grava:

    NFiscalEntradaCom_Grava = gErr

    Select Case gErr

        Case 65915, 207820
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 65925, 207823
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case 124153, 500021, 500027, 207819, 207821, 207822, 207824
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149613)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function

Function NFiscalFatEntradaCom_Grava(ByVal objNFiscal As ClassNFiscal, ByVal objContabil As ClassContabil, ByVal colPedCompras As Collection, ByVal colReqCompras As Collection) As Long
'Faz a Gravação de Uma Nota Fiscal de Entrada

Dim lErro As Long
Dim lTransacao As Long
Dim lNumNFOficial As Long
Dim lErro1 As Long
Dim lErro2 As Long
Dim vbMsgRes As VbMsgBoxResult
Dim lNumNotaFiscal As Long

On Error GoTo Erro_NFiscalFatEntradaCom_Grava

    lNumNotaFiscal = objNFiscal.lNumNotaFiscal

    'Abre a Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 66500

    Set objNFiscal.objContabil = objContabil

    'sport
    lErro1 = NFiscal_Grava_Clone(objNFiscal, objContabil, "NFiscalFatEntrada_Grava_EmTrans", lNumNFOficial)
    'If lErro <> SUCESSO Then gError 500028
    'FIM sport
    
    If lErro1 = SUCESSO Then
        lErro2 = CF("NFiscalFatEntradaCom_Grava_EmTrans", objNFiscal, objContabil, colPedCompras, colReqCompras)
    End If

    If lErro1 <> SUCESSO Or lErro2 <> SUCESSO Then
    
        'se era uma alteracao e deu erro na exclusao;insercao ==> tenta somente gravar os dados de contabilidade
        If lNumNotaFiscal <> 0 Then
    
            'Dá a Mensagem de aviso que só será gravado a contabilidade
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_ALTERACAO_NFISCAL_INTERNA_CONTAB", objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.dtDataEmissao)
            If vbMsgRes = vbNo Then gError 207825

            Call Transacao_Rollback
    
            lTransacao = Transacao_Abrir()
            If lTransacao = 0 Then gError 207826
    
            lErro = NFiscal_Altera_Clone(objNFiscal, objContabil, "NFiscalEntrada_AlteraDados_EmTrans", lNumNFOficial)
            If lErro <> SUCESSO Then gError 207827
        
            lErro = CF("NFiscalEntrada_AlteraDados_EmTrans", objNFiscal, objContabil)
            If lErro <> SUCESSO Then gError 207828
    
            'Faz o Commit da Transação
            lErro = Transacao_Commit()
            If lErro <> AD_SQL_SUCESSO Then gError 207829
    
        Else
            gError 207830
    
        End If
    
    Else
    
    
        'sport
        If lNumNFOficial <> 0 And lNumNFOficial <> objNFiscal.lNumNotaFiscal Then gError 500022
        'FIM sport
        
        'Faz o Commit da Transação
        lErro = Transacao_Commit()
        If lErro <> AD_SQL_SUCESSO Then gError 66509


    End If


    NFiscalFatEntradaCom_Grava = SUCESSO

    Exit Function

Erro_NFiscalFatEntradaCom_Grava:

    NFiscalFatEntradaCom_Grava = gErr

    Select Case gErr

        Case 66500, 207826
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 66509, 207829
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case 124152, 500022, 500028, 207825, 207827, 207828, 207830
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149614)

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
        If objNFiscal.iFilialEmpresa <= DELTA_FILIALREAL_OFICIAL Then
            lErro = Comando_Executar(lComando, "SELECT NumIntDoc, DataEntrada FROM TipoDocInfoNFiscal WHERE FilialEmpresa <= ? AND Fornecedor = ? AND FilialForn = ? AND Cliente = ? AND FilialCli = ? AND DataEmissao = ? AND Serie = ? AND NumNotaFiscal = ? AND Status <> ? AND Tipo = ? ", lNumIntDoc, dtDataEntrada, DELTA_FILIALREAL_OFICIAL, objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.lCliente, objNFiscal.iFilialCli, objNFiscal.dtDataEmissao, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, STATUS_EXCLUIDO, DOCINFO_NF_EXTERNA)
        Else
            lErro = Comando_Executar(lComando, "SELECT NumIntDoc, DataEntrada FROM TipoDocInfoNFiscal WHERE FilialEmpresa > ? AND Fornecedor = ? AND FilialForn = ? AND Cliente = ? AND FilialCli = ? AND DataEmissao = ? AND Serie = ? AND NumNotaFiscal = ? AND Status <> ? AND Tipo = ? ", lNumIntDoc, dtDataEntrada, DELTA_FILIALREAL_OFICIAL, objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.lCliente, objNFiscal.iFilialCli, objNFiscal.dtDataEmissao, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, STATUS_EXCLUIDO, DOCINFO_NF_EXTERNA)
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149615)
    
    End Select
    
    Exit Function

End Function

Function NFiscalEntrada_Grava(ByVal objNFiscal As ClassNFiscal, ByVal objContabil As ClassContabil) As Long
'Faz a Gravação de Uma Nota Fiscal de Entrada

Dim lErro As Long
Dim lTransacao As Long
Dim lNumNFOficial As Long
Dim lNumNF As Long
Dim lErro1 As Long
Dim lErro2 As Long
Dim vbMsgRes As VbMsgBoxResult
Dim lNumNotaFiscal As Long

On Error GoTo Erro_NFiscalEntrada_Grava

    lNumNotaFiscal = objNFiscal.lNumNotaFiscal

    'Abre a Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 183184

    Set objNFiscal.objContabil = objContabil

'    'temporario
'    lNumNF = objNFiscal.lNumNotaFiscal
'
'    'temporario
'    objNFiscal.lNumNotaFiscal = 0
    
    'sport
    lErro1 = NFiscal_Grava_Clone(objNFiscal, objContabil, "NFiscalEntrada_Grava_EmTrans", lNumNFOficial)
'    If lErro <> SUCESSO Then gError 183185
    'FIM sport
    
    
'    'temporario
'    objNFiscal.lNumNotaFiscal = lNumNF
        
    If lErro1 = SUCESSO Then
        lErro2 = CF("NFiscalEntrada_Grava_EmTrans", objNFiscal, objContabil)
    End If
        
    If lErro1 <> SUCESSO Or lErro2 <> SUCESSO Then
    
        'se era uma alteracao e deu erro na exclusao;insercao ==> tenta somente gravar os dados de contabilidade
        If lNumNotaFiscal <> 0 Then
    
            'Dá a Mensagem de aviso que só será gravado a contabilidade
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_ALTERACAO_NFISCAL_INTERNA_CONTAB", objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.dtDataEmissao)
            If vbMsgRes = vbNo Then gError 207831

            Call Transacao_Rollback
    
            lTransacao = Transacao_Abrir()
            If lTransacao = 0 Then gError 207832
    
            lErro = NFiscal_Altera_Clone(objNFiscal, objContabil, "NFiscalEntrada_AlteraDados_EmTrans", lNumNFOficial)
            If lErro <> SUCESSO Then gError 207833
        
            lErro = CF("NFiscalEntrada_AlteraDados_EmTrans", objNFiscal, objContabil)
            If lErro <> SUCESSO Then gError 207834
    
            'Faz o Commit da Transação
            lErro = Transacao_Commit()
            If lErro <> AD_SQL_SUCESSO Then gError 207835
    
        Else
            gError 207836
    
        End If
    
    Else
    
    
        'sport
        If lNumNFOficial <> 0 And lNumNFOficial <> objNFiscal.lNumNotaFiscal Then gError 183187
        'FIM sport
        
        'Faz o Commit da Transação
        lErro = Transacao_Commit()
        If lErro <> AD_SQL_SUCESSO Then gError 183188

    End If



    NFiscalEntrada_Grava = SUCESSO

    Exit Function

Erro_NFiscalEntrada_Grava:

    NFiscalEntrada_Grava = gErr

    Select Case gErr

        Case 183184, 207832
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 183185 To 183187, 207831, 207833, 207834, 207836
        
        Case 183188, 207835
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183189)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function

Function NFiscalFatEntrada_Grava(ByVal objNFiscal As ClassNFiscal, ByVal objContabil As ClassContabil) As Long
'Faz a Gravação de Uma Nota Fiscal Fatura de Entrada

Dim lErro As Long
Dim lTransacao As Long
Dim lNumNFOficial As Long
Dim lErro1 As Long
Dim lErro2 As Long
Dim vbMsgRes As VbMsgBoxResult
Dim lNumNotaFiscal As Long

On Error GoTo Erro_NFiscalFatEntrada_Grava

    lNumNotaFiscal = objNFiscal.lNumNotaFiscal

    'Abre a Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 183190

    Set objNFiscal.objContabil = objContabil
    
    'sport
    lErro1 = NFiscal_Grava_Clone(objNFiscal, objContabil, "NFiscalFatEntrada_Grava_EmTrans", lNumNFOficial)
'    If lErro <> SUCESSO Then gError 183191
    'FIM sport
    
    If lErro1 = SUCESSO Then
        lErro2 = CF("NFiscalFatEntrada_Grava_EmTrans", objNFiscal, objContabil)
    End If
        
    If lErro1 <> SUCESSO Or lErro2 <> SUCESSO Then
    
        'se era uma alteracao e deu erro na exclusao;insercao ==> tenta somente gravar os dados de contabilidade
        If lNumNotaFiscal <> 0 Then
    
            'Dá a Mensagem de aviso que só será gravado a contabilidade
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_ALTERACAO_NFISCAL_INTERNA_CONTAB", objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.dtDataEmissao)
            If vbMsgRes = vbNo Then gError 207837

            Call Transacao_Rollback
    
            lTransacao = Transacao_Abrir()
            If lTransacao = 0 Then gError 207838
    
            lErro = NFiscal_Altera_Clone(objNFiscal, objContabil, "NFiscalEntrada_AlteraDados_EmTrans", lNumNFOficial)
            If lErro <> SUCESSO Then gError 207839
        
            lErro = CF("NFiscalEntrada_AlteraDados_EmTrans", objNFiscal, objContabil)
            If lErro <> SUCESSO Then gError 207840
    
            'Faz o Commit da Transação
            lErro = Transacao_Commit()
            If lErro <> AD_SQL_SUCESSO Then gError 207841
    
        Else
            gError 207842
    
        End If
    
    Else
    
        'sport
        If lNumNFOficial <> 0 And lNumNFOficial <> objNFiscal.lNumNotaFiscal Then gError 183193
        'FIM sport
        
        'Faz o Commit da Transação
        lErro = Transacao_Commit()
        If lErro <> AD_SQL_SUCESSO Then gError 183194

    End If


    NFiscalFatEntrada_Grava = SUCESSO

    Exit Function

Erro_NFiscalFatEntrada_Grava:

    NFiscalFatEntrada_Grava = gErr

    Select Case gErr

        Case 183190, 207838
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 183194, 207841
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case 183191 To 183193, 207837, 207839, 207840, 207842
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 183195)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function


Function NFiscalEntradaRemessa_Grava(ByVal objNFiscal As ClassNFiscal, ByVal objContabil As ClassContabil) As Long
'Faz a Gravação de Uma Nota Fiscal de Entrada

Dim lErro As Long
Dim lTransacao As Long
Dim lNumNFOficial As Long
Dim lNumNF As Long
Dim lErro1 As Long
Dim lErro2 As Long
Dim vbMsgRes As VbMsgBoxResult
Dim lNumNotaFiscal As Long


On Error GoTo Erro_NFiscalEntradaRemessa_Grava

    lNumNotaFiscal = objNFiscal.lNumNotaFiscal

    'Abre a Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 197961

    lErro1 = NFiscal_Grava_Clone(objNFiscal, objContabil, "NFiscalEntradaRemessa_Grava_EmTrans", lNumNFOficial)
'    If lErro <> SUCESSO Then gError 197962
    
    If lErro1 = SUCESSO Then
        lErro2 = CF("NFiscalEntradaRemessa_Grava_EmTrans", objNFiscal, objContabil)
    End If
    
        
    If lErro1 <> SUCESSO Or lErro2 <> SUCESSO Then
    
        'se era uma alteracao e deu erro na exclusao;insercao ==> tenta somente gravar os dados de contabilidade
        If lNumNotaFiscal <> 0 Then
    
            'Dá a Mensagem de aviso que só será gravado a contabilidade
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_ALTERACAO_NFISCAL_INTERNA_CONTAB", objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.dtDataEmissao)
            If vbMsgRes = vbNo Then gError 207843

            Call Transacao_Rollback
    
            lTransacao = Transacao_Abrir()
            If lTransacao = 0 Then gError 207844
    
            lErro = NFiscal_Altera_Clone(objNFiscal, objContabil, "NFiscalEntrada_AlteraDados_EmTrans", lNumNFOficial)
            If lErro <> SUCESSO Then gError 207845
        
            lErro = CF("NFiscalEntrada_AlteraDados_EmTrans", objNFiscal, objContabil)
            If lErro <> SUCESSO Then gError 207846
    
            'Faz o Commit da Transação
            lErro = Transacao_Commit()
            If lErro <> AD_SQL_SUCESSO Then gError 207847
    
        Else
            gError 207848
    
        End If
    
    Else
    
        If lNumNFOficial <> 0 And lNumNFOficial <> objNFiscal.lNumNotaFiscal Then gError 197964
        
        'Faz o Commit da Transação
        lErro = Transacao_Commit()
        If lErro <> AD_SQL_SUCESSO Then gError 197965

    End If


    NFiscalEntradaRemessa_Grava = SUCESSO

    Exit Function

Erro_NFiscalEntradaRemessa_Grava:

    NFiscalEntradaRemessa_Grava = gErr

    Select Case gErr

        Case 197961, 207844
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 197962 To 197964, 207843, 207845, 207846, 207848
        
        Case 197965, 207847
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197966)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function

Function NFiscalEntradaDevolucao_Grava(ByVal objNFiscal As ClassNFiscal, ByVal objContabil As ClassContabil) As Long
'Faz a Gravação de Uma Nota Fiscal de Entrada

Dim lErro As Long
Dim lTransacao As Long
Dim lNumNFOficial As Long
Dim lNumNF As Long
Dim lErro1 As Long
Dim lErro2 As Long
Dim vbMsgRes As VbMsgBoxResult
Dim lNumNotaFiscal As Long

On Error GoTo Erro_NFiscalEntradaDevolucao_Grava

    lNumNotaFiscal = objNFiscal.lNumNotaFiscal

    'Abre a Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 197967

    lErro1 = NFiscal_Grava_Clone(objNFiscal, objContabil, "NFiscalEntradaDevolucao_Grava_EmTrans", lNumNFOficial)
    
    If lErro1 = SUCESSO Then
        lErro2 = CF("NFiscalEntradaDevolucao_Grava_EmTrans", objNFiscal, objContabil)
    End If
        
    If lErro1 <> SUCESSO Or lErro2 <> SUCESSO Then
    
        'se era uma alteracao e deu erro na exclusao;insercao ==> tenta somente gravar os dados de contabilidade
        If lNumNotaFiscal <> 0 Then
    
            'Dá a Mensagem de aviso que só será gravado a contabilidade
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_ALTERACAO_NFISCAL_INTERNA_CONTAB", objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.dtDataEmissao)
            If vbMsgRes = vbNo Then gError 207849

            Call Transacao_Rollback
    
            lTransacao = Transacao_Abrir()
            If lTransacao = 0 Then gError 207850
    
            lErro = NFiscal_Altera_Clone(objNFiscal, objContabil, "NFiscalEntrada_AlteraDados_EmTrans", lNumNFOficial)
            If lErro <> SUCESSO Then gError 207851
        
            lErro = CF("NFiscalEntrada_AlteraDados_EmTrans", objNFiscal, objContabil)
            If lErro <> SUCESSO Then gError 207852
    
            'Faz o Commit da Transação
            lErro = Transacao_Commit()
            If lErro <> AD_SQL_SUCESSO Then gError 207853
    
        Else
            gError 207854
    
        End If
    
    Else
    
        If lNumNFOficial <> 0 And lNumNFOficial <> objNFiscal.lNumNotaFiscal Then gError 197970
        
        'Faz o Commit da Transação
        lErro = Transacao_Commit()
        If lErro <> AD_SQL_SUCESSO Then gError 197971

    End If


    NFiscalEntradaDevolucao_Grava = SUCESSO

    Exit Function

Erro_NFiscalEntradaDevolucao_Grava:

    NFiscalEntradaDevolucao_Grava = gErr

    Select Case gErr

        Case 197967, 207850
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 197968 To 197970, 207849, 207851, 207852, 207854
        
        Case 197971, 207853
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197972)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function

Function NotaFiscalEntrada_Excluir(ByVal objNFiscal As ClassNFiscal, ByVal objContabil As ClassContabil) As Long
'Faz a exclusao de uma nota fiscal de entrada

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_NotaFiscalEntrada_Excluir
    
    'ABre a transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 197970

    'exclui o clone se houver
    lErro = NFiscal_Ent_Exclui_Clone(objNFiscal, objContabil)
    If lErro <> SUCESSO Then gError 197971

    lErro = CF("NotaFiscalEntrada_Excluir_EmTrans", objNFiscal, objContabil)
    If lErro <> SUCESSO Then gError 197972
    
    'Confirma a transacao
    lErro = Transacao_Commit()
    If lErro <> SUCESSO Then gError 197973
    
    NotaFiscalEntrada_Excluir = SUCESSO

    Exit Function

Erro_NotaFiscalEntrada_Excluir:

    NotaFiscalEntrada_Excluir = gErr

    Select Case gErr

        Case 197970
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 197971, 197972
        
        Case 197973
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197974)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function

Function NotaFiscalEntrada_Cancelar(ByVal objNFiscal As ClassNFiscal, ByVal dtDataCancelamento As Date) As Long
'Faz o cancelamento de uma nota fiscal de entrada

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_NotaFiscalEntrada_Cancelar
    
    'ABre a transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 197975

    'exclui o clone se houver
    lErro = NFiscal_Ent_Cancela_Clone(objNFiscal, dtDataCancelamento)
    If lErro <> SUCESSO Then gError 197976

    lErro = CF("NotaFiscalEntrada_Cancelar_EmTrans", objNFiscal, dtDataCancelamento)
    If lErro <> SUCESSO Then gError 197977
    
    'Confirma a transacao
    lErro = Transacao_Commit()
    If lErro <> SUCESSO Then gError 197978
    
    NotaFiscalEntrada_Cancelar = SUCESSO

    Exit Function

Erro_NotaFiscalEntrada_Cancelar:

    NotaFiscalEntrada_Cancelar = gErr

    Select Case gErr

        Case 197975
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 197976, 197977
        
        Case 197978
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197979)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function

Function NotaFiscalEntrada_Cancelar_EmTrans1(ByVal objNFiscalOficial As ClassNFiscal, ByVal dtDataCancelamento As Date) As Long
'Faz o cancelamento de uma nota fiscal de Entrada

Dim lErro As Long
Dim objNFiscal As New ClassNFiscal

On Error GoTo Erro_NotaFiscalEntrada_Cancelar_EmTrans1
    
    objNFiscal.lNumIntDoc = objNFiscalOficial.lNumIntDoc + 1

    lErro = CF("NFiscal_Le", objNFiscal)
    If lErro <> SUCESSO And lErro <> 31442 Then gError 210483
    
    If lErro <> SUCESSO Then gError 210484
    
    'Lê os itens da nota fiscal
    lErro = CF("NFiscalItens_Le", objNFiscal)
    If lErro <> SUCESSO Then gError 210485
    
    objNFiscal.sMotivoCancel = objNFiscalOficial.sMotivoCancel
    
    'exclui o clone se houver
    lErro = NFiscal_Cancela_Clone(objNFiscal, dtDataCancelamento)
    If lErro <> SUCESSO Then gError 210486
    
    lErro = CF("NotaFiscalEntrada_Cancelar_EmTrans", objNFiscal, dtDataCancelamento)
    If lErro <> SUCESSO Then gError 210487
    
    NotaFiscalEntrada_Cancelar_EmTrans1 = SUCESSO

    Exit Function

Erro_NotaFiscalEntrada_Cancelar_EmTrans1:

    NotaFiscalEntrada_Cancelar_EmTrans1 = gErr

    Select Case gErr

        Case 210483, 210485 To 210487
        
        Case 210484
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA", gErr, objNFiscal.lNumIntDoc)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 210481)

    End Select

    Exit Function

End Function


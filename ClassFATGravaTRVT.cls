VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFATGravaTRV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'GRAVAÇÃO
Public Function TRVOcorrencias_Grava(ByVal objTRVOcorrencias As ClassTRVOcorrencias) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_TRVOcorrencias_Grava

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 189450

    lErro = CF("TRVOcorrencias_Grava_EmTrans", objTRVOcorrencias)
    If lErro <> SUCESSO Then gError 189451

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 189452

    TRVOcorrencias_Grava = SUCESSO

    Exit Function

Erro_TRVOcorrencias_Grava:

    TRVOcorrencias_Grava = gErr

    Select Case gErr

        Case 189450
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 189451

        Case 189452
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189473)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function TRVOcorrencias_Grava_EmTrans(ByVal objTRVOcorrencias As ClassTRVOcorrencias, Optional ByVal bValidaCancVou As Boolean = True, Optional ByVal bIgnoraJaFaturado As Boolean = False) As Long

Dim lErro As Long
Dim alComando(0 To 7) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim lNumIntDocDest As Long
Dim iFormaPagto As Integer
Dim objVoucher As New ClassTRVVouchers
Dim sSerie As String
Dim sTipo As String
Dim lNumVou As Long
Dim iStatusVou As Integer
Dim dValorOcr As Double
Dim dValorOcrAtu As Double
Dim dValorAnterior As Double
Dim BOrcNova As Boolean
Dim bIgnorar As Boolean

On Error GoTo Erro_TRVOcorrencias_Grava_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 189454
    Next
       
    sSerie = String(STRING_TRV_OCR_SERIE, 0)
    sTipo = String(STRING_TRV_OCR_TIPOVOU, 0)

    'Le a tabelaTRVOcorrencias
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc,NumIntDocDestino,FormaPagto, Serie, TipoDoc, NumVou, ValorTotal FROM TRVOcorrencias WHERE Codigo= ? ", _
                0, lNumIntDoc, lNumIntDocDest, iFormaPagto, sSerie, sTipo, lNumVou, dValorAnterior, objTRVOcorrencias.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 189455

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189456

    If lErro = AD_SQL_SUCESSO Then
        BOrcNova = False
    Else
        BOrcNova = True
        dValorAnterior = 0
    End If

    'Le a tabela TRVVoucher
    lErro = Comando_ExecutarPos(alComando(6), "SELECT Status, ValorOcr FROM TRVVouchers WHERE TipoDoc= ? AND NumVou = ? AND TipVou = ? AND Serie = ? ", _
                0, iStatusVou, dValorOcr, TRV_TIPODOC_VOU_TEXTO, objTRVOcorrencias.lNumVou, objTRVOcorrencias.sTipoDoc, objTRVOcorrencias.sSerie)
    If lErro <> AD_SQL_SUCESSO Then gError 194437

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(6))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194438

    If lErro <> AD_SQL_SUCESSO Then gError 194441

    'Trava registro contra alterações/Leituras
    lErro = Comando_LockExclusive(alComando(6))
    If lErro <> SUCESSO Then gError 194439
    
    dValorOcrAtu = dValorOcr + objTRVOcorrencias.dValorTotal - dValorAnterior

    lErro = Comando_ExecutarPos(alComando(7), "UPDATE TRVVouchers SET ValorOcr= ? , TemOcr= ? ", alComando(6), dValorOcrAtu, MARCADO)
    If lErro <> SUCESSO Then gError 194440
    
    bIgnorar = False

    'Se existir => UPDATE, senão => INSERT
    If Not BOrcNova Then
    
        If lNumIntDocDest <> 0 Then
            If Not bIgnoraJaFaturado Then
                If iFormaPagto = FORMAPAGTO_TRV_OCR_FAT Then
                    gError 190310
                Else
                    gError 190311
                End If
            Else
                bIgnorar = True
            End If
        End If
        
        If Not bIgnorar Then

            'Trava registro contra alterações/Leituras
            lErro = Comando_LockExclusive(alComando(0))
            If lErro <> SUCESSO Then gError 189457
            
            objTRVOcorrencias.lNumIntDoc = lNumIntDoc
                
            lErro = Comando_ExecutarPos(alComando(1), "UPDATE TRVOcorrencias SET NumVou= ? , TipoDoc= ? , Serie= ? , Cliente= ? , " & _
                    "FilialCliente= ? , DataEmissao= ? , Observacao= ? , Status= ? , Origem= ? , " & _
                    "Historico= ? , FormaPagto= ? , ValorTotal= ? ", alComando(0), _
                    objTRVOcorrencias.lNumVou, objTRVOcorrencias.sTipoDoc, objTRVOcorrencias.sSerie, objTRVOcorrencias.lCliente, _
                    objTRVOcorrencias.iFilialCliente, objTRVOcorrencias.dtDataEmissao, objTRVOcorrencias.sObservacao, objTRVOcorrencias.iStatus, objTRVOcorrencias.iOrigem, _
                    objTRVOcorrencias.sHistorico, objTRVOcorrencias.iFormaPagto, objTRVOcorrencias.dValorTotal)
            If lErro <> AD_SQL_SUCESSO Then gError 189458
            
            lErro = TRVOcorrenciaDet_Exclui(alComando(2), alComando(3), objTRVOcorrencias)
            If lErro <> SUCESSO Then gError 189473
            
        End If

    Else

        'Obter NumIntDoc
        lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_INT_PROX_TRVOCORRENCIAS", lNumIntDoc)
        If lErro <> SUCESSO Then gError 189453

        objTRVOcorrencias.lNumIntDoc = lNumIntDoc
        
        lErro = Comando_Executar(alComando(1), "INSERT INTO TRVOcorrencias( NumIntDoc, NumVou, TipoDoc, Serie, " & _
                "Codigo, Cliente, FilialCliente, DataEmissao, Observacao, " & _
                "Status, Origem, Historico, FormaPagto, " & _
                "ValorTotal) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                objTRVOcorrencias.lNumIntDoc, objTRVOcorrencias.lNumVou, objTRVOcorrencias.sTipoDoc, objTRVOcorrencias.sSerie, _
                objTRVOcorrencias.lCodigo, objTRVOcorrencias.lCliente, objTRVOcorrencias.iFilialCliente, objTRVOcorrencias.dtDataEmissao, objTRVOcorrencias.sObservacao, _
                objTRVOcorrencias.iStatus, objTRVOcorrencias.iOrigem, objTRVOcorrencias.sHistorico, objTRVOcorrencias.iFormaPagto, _
                objTRVOcorrencias.dValorTotal)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189459

    End If
    
    If Not bIgnorar Then
        lErro = TRVOcorrenciaDet_Insere(alComando(4), objTRVOcorrencias)
        If lErro <> SUCESSO Then gError 189474
    End If
    
    If giVersaoTRV > TRV_VERSAO_FATURAMENTO Then
    
        If Abs(objTRVOcorrencias.dValorTotal - dValorAnterior) > DELTA_VALORMONETARIO Then
        
            objVoucher.lNumVou = objTRVOcorrencias.lNumVou
            objVoucher.sSerie = objTRVOcorrencias.sSerie
            objVoucher.sTipoDoc = TRV_TIPODOC_VOU_TEXTO
            objVoucher.sTipVou = objTRVOcorrencias.sTipoDoc
        
            lErro = CF("TRVVouchers_Le", objVoucher)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 192874
            
            'Tem que gerar a comissão
            If objTRVOcorrencias.iOrigem = INATIVACAO_AUTOMATICA_CODIGO Then
                lErro = CF("TRVVoucher_Gera_Comissao", objVoucher, objTRVOcorrencias.dValorTotal - dValorAnterior, objTRVOcorrencias.dtDataEmissao, "Inativação do Voucher", TRV_TIPO_TRATAMENTO_COMI_NVL)
            Else
                lErro = CF("TRVVoucher_Gera_Comissao", objVoucher, objTRVOcorrencias.dValorTotal - dValorAnterior, objTRVOcorrencias.dtDataEmissao, "Ocorrência código " & CStr(objTRVOcorrencias.lCodigo), TRV_TIPO_TRATAMENTO_COMI_OCR)
            End If
            If lErro <> SUCESSO Then gError 196312
            
        End If
        
    End If
        
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TRVOcorrencias_Grava_EmTrans = SUCESSO

    Exit Function

Erro_TRVOcorrencias_Grava_EmTrans:

    TRVOcorrencias_Grava_EmTrans = gErr

    Select Case gErr

        Case 189453, 189473, 189474, 196312

        Case 189454
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 189455, 189456
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVOCORRENCIAS", gErr)

        Case 189457
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRVOCORRENCIAS", gErr)

        Case 189458
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRVOCORRENCIAS", gErr)

        Case 189459
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRVOCORRENCIAS", gErr)

        Case 190310
            Call Rotina_Erro(vbOKOnly, "ERRO_TRVOCORRENCIAS_FATURADA", gErr)

        Case 190311
            Call Rotina_Erro(vbOKOnly, "ERRO_TRVOCORRENCIAS_DEBITO", gErr)
            
        Case 192874
        
        Case 192875
            Call Rotina_Erro(vbOKOnly, "ERRO_VOUCHER_JA_CANCELADO", gErr)
            
        Case 194437, 194438
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVVOUCHERS", gErr)

        Case 194439
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRVVOUCHERS", gErr)

        Case 194440
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRVVOUCHERS", gErr)

        Case 194441
            Call Rotina_Erro(vbOKOnly, "ERRO_TRVVOUCHERS_NAO_CADASTRADO", gErr, lNumVou, sTipo, sSerie)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189479)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function TRVOcorrenciaDet_Insere(ByVal lComando As Long, ByVal objTRVOcorrencia As ClassTRVOcorrencias) As Long

Dim lErro As Long
Dim lNumIntDoc As Long
Dim objTRVOcorrenciaDet As ClassTRVOcorrenciaDet

On Error GoTo Erro_TRVOcorrenciaDet_Insere

    For Each objTRVOcorrenciaDet In objTRVOcorrencia.colDetalhamento

        'Obter NumIntDoc
        lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_INT_PROX_TRVOCORRENCIADET", lNumIntDoc)
        If lErro <> SUCESSO Then gError 189476
    
        objTRVOcorrenciaDet.lNumIntDoc = lNumIntDoc
        objTRVOcorrenciaDet.lNumIntDocOCR = objTRVOcorrencia.lNumIntDoc
    
        lErro = Comando_Executar(lComando, "INSERT INTO TRVOcorrenciaDet( NumIntDoc, NumIntDocOcr, Tipo, Valor, Seq) VALUES (?,?,?,?,?)", _
                objTRVOcorrenciaDet.lNumIntDoc, objTRVOcorrenciaDet.lNumIntDocOCR, objTRVOcorrenciaDet.iTipo, objTRVOcorrenciaDet.dValor, objTRVOcorrenciaDet.iSeq)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189482

    Next

    TRVOcorrenciaDet_Insere = SUCESSO

    Exit Function

Erro_TRVOcorrenciaDet_Insere:

    TRVOcorrenciaDet_Insere = gErr

    Select Case gErr

        Case 189476

        Case 189477
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRVOCORRENCIADET", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189478)

    End Select

    Exit Function

End Function

Public Function TRVOcorrenciaDet_Exclui(ByVal lComando1 As Long, lComando2 As Long, ByVal objTRVOcorrencia As ClassTRVOcorrencias) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long

On Error GoTo Erro_TRVOcorrenciaDet_Exclui

    'Le a tabela TRVOcorrenciaDet
    lErro = Comando_ExecutarPos(lComando1, "SELECT NumIntDoc FROM TRVOcorrenciaDet WHERE NumIntDocOcr= ? ", _
                0, lNumIntDoc, objTRVOcorrencia.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 189468

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189469

    Do While lErro <> AD_SQL_SEM_DADOS

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(lComando1)
        If lErro <> SUCESSO Then gError 189471
    
        lErro = Comando_ExecutarPos(lComando2, "DELETE FROM TRVOcorrenciaDet", lComando1)
        If lErro <> AD_SQL_SUCESSO Then gError 189472

        'Busca Próximo
        lErro = Comando_BuscarProximo(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189470

    Loop

    TRVOcorrenciaDet_Exclui = SUCESSO

    Exit Function

Erro_TRVOcorrenciaDet_Exclui:

    TRVOcorrenciaDet_Exclui = gErr

    Select Case gErr

        Case 189468, 189469, 189470
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVOCORRENCIADET", gErr)

        Case 189471
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRVOCORRENCIADET", gErr)

        Case 189472
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRVOCORRENCIADET", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189480)

    End Select

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'EXCLUSÃO
Public Function TRVOcorrencias_Exclui(ByVal objTRVOcorrencias As ClassTRVOcorrencias) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_TRVOcorrencias_Exclui

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 189441

    lErro = CF("TRVOcorrencias_Exclui_EmTrans", objTRVOcorrencias)
    If lErro <> SUCESSO Then gError 189442

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 189443

    TRVOcorrencias_Exclui = SUCESSO

    Exit Function

Erro_TRVOcorrencias_Exclui:

    TRVOcorrencias_Exclui = gErr

    Select Case gErr

        Case 189441
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 189442

        Case 189443
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189482)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function TRVOcorrencias_Exclui_EmTrans(ByVal objTRVOcorrencias As ClassTRVOcorrencias, Optional ByVal bInativacaoExclui As Boolean = False, Optional ByVal bImportacao As Boolean = False) As Long

Dim lErro As Long
Dim alComando(0 To 5) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim iTipoDestino As Integer
Dim lNumIntDocDestino As Long
Dim iTemOcr As Integer
Dim dValorOcr As Double
Dim dValorOcrAtu As Double
Dim lNumVou As Long
Dim sTipoDoc As String
Dim sSerie As String
Dim iStatusVou As Integer
Dim iOrigem As Integer
Dim bSemDados As Boolean
Dim objVoucher As New ClassTRVVouchers
Dim dValorTotal As Double
Dim dtDataEmissao As Date

On Error GoTo Erro_TRVOcorrencias_Exclui_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 189444
    Next
    
    sSerie = String(STRING_TRV_OCR_SERIE, 0)
    sTipoDoc = String(STRING_TRV_OCR_TIPOVOU, 0)
    
    'Le a tabela TRVOcorrencias
    If Not bInativacaoExclui Then
        lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc, TipoDocDestino, NumIntDocDestino,Origem, NumVou, TipoDoc, Serie, ValorTotal, DataEmissao FROM TRVOcorrencias WHERE Codigo= ? ", _
                    0, lNumIntDoc, iTipoDestino, lNumIntDocDestino, iOrigem, lNumVou, sTipoDoc, sSerie, dValorTotal, dtDataEmissao, objTRVOcorrencias.lCodigo)
    Else
        lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc, TipoDocDestino, NumIntDocDestino,Origem, NumVou, TipoDoc, Serie, ValorTotal, DataEmissao FROM TRVOcorrencias WHERE Origem = ? AND NumVou = ? AND TipoDoc = ? AND Serie = ? AND Cliente = ? ", _
                    0, lNumIntDoc, iTipoDestino, lNumIntDocDestino, iOrigem, lNumVou, sTipoDoc, sSerie, dValorTotal, dtDataEmissao, INATIVACAO_AUTOMATICA_CODIGO, objTRVOcorrencias.lNumVou, objTRVOcorrencias.sTipoDoc, objTRVOcorrencias.sSerie, objTRVOcorrencias.lCliente)
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 189445

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189446

    'Se não existir => ERRO
    If lErro = AD_SQL_SEM_DADOS Then
    
        lNumVou = objTRVOcorrencias.lNumVou
        sTipoDoc = objTRVOcorrencias.sTipoDoc
        sSerie = objTRVOcorrencias.sSerie
        
        bSemDados = True
        
        If bImportacao And Not bInativacaoExclui Then
            TRVOcorrencias_Exclui_EmTrans = SUCESSO
            Exit Function
        End If
        
    Else
        bSemDados = False
    End If
    
    If bSemDados And Not bInativacaoExclui Then gError 189447
    
    'Le a tabela TRVVoucher
    lErro = Comando_ExecutarPos(alComando(4), "SELECT Status, ValorOcr FROM TRVVouchers WHERE TipoDoc= ? AND NumVou = ? AND TipVou = ? AND Serie = ? ", _
                0, iStatusVou, dValorOcr, TRV_TIPODOC_VOU_TEXTO, lNumVou, sTipoDoc, sSerie)
    If lErro <> AD_SQL_SUCESSO Then gError 194437

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(4))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194438
    
    If lErro <> AD_SQL_SUCESSO Then gError 194441
    
    'Trava registro contra alterações/Leituras
    lErro = Comando_LockExclusive(alComando(4))
    If lErro <> SUCESSO Then gError 194439
    
    dValorOcrAtu = dValorOcr - dValorTotal
    
    If Abs(dValorOcrAtu) > DELTA_VALORMONETARIO Then
        iTemOcr = MARCADO
    Else
        iTemOcr = DESMARCADO
    End If

    If bSemDados Then
    
        If bInativacaoExclui Then
            'Se ainda não tem a OCR de inativação e for uma exclusão de inativação apenas troca o status do voucher
            lErro = Comando_ExecutarPos(alComando(5), "UPDATE TRVVouchers SET Status = ?", alComando(4), STATUS_TRV_VOU_ABERTO)
        End If
    
    ElseIf iOrigem = INATIVACAO_AUTOMATICA_CODIGO Then
        'Se tiver a ocr de inativação cadastrada atualiza os valores no voucher e o status
        lErro = Comando_ExecutarPos(alComando(5), "UPDATE TRVVouchers SET ValorOcr= ? , TemOcr= ? , Status = ?", alComando(4), dValorOcrAtu, iTemOcr, STATUS_TRV_VOU_ABERTO)
    Else
        'Se for só uma exclusão de ocr acerta os valores no voucher
        lErro = Comando_ExecutarPos(alComando(5), "UPDATE TRVVouchers SET ValorOcr= ? , TemOcr= ? ", alComando(4), dValorOcrAtu, iTemOcr)
    End If
    If lErro <> SUCESSO Then gError 194440
    
    
    If giVersaoTRV > TRV_VERSAO_FATURAMENTO Then

        objVoucher.lNumVou = lNumVou
        objVoucher.sSerie = sSerie
        objVoucher.sTipoDoc = TRV_TIPODOC_VOU_TEXTO
        objVoucher.sTipVou = sTipoDoc
    
        lErro = CF("TRVVouchers_Le", objVoucher)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 196381
        
        'Tem que gerar a comissão
        If objTRVOcorrencias.iOrigem = INATIVACAO_AUTOMATICA_CODIGO Then
            lErro = CF("TRVVoucher_Gera_Comissao", objVoucher, -dValorTotal, dtDataEmissao, "Exclusão da inativação do Voucher", TRV_TIPO_TRATAMENTO_COMI_NVL)
        Else
            lErro = CF("TRVVoucher_Gera_Comissao", objVoucher, -dValorTotal, dtDataEmissao, "Exclusão da ocorrência código " & CStr(objTRVOcorrencias.lCodigo), TRV_TIPO_TRATAMENTO_COMI_OCR)
        End If
        If lErro <> SUCESSO Then gError 196382

    End If
    
    'Se não tinha a OCR cadastrada não faz mais nada
    If Not bSemDados Then
    
        objTRVOcorrencias.lNumIntDoc = lNumIntDoc

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 189448
    
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM TRVOcorrencias", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 189449
        
        lErro = TRVOcorrenciaDet_Exclui(alComando(2), alComando(3), objTRVOcorrencias)
        If lErro <> SUCESSO Then gError 189485
        
        lErro = Item_Exclui_Com_Destino(iTipoDestino, lNumIntDocDestino)
        If lErro <> SUCESSO Then gError 192062
        
    End If
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TRVOcorrencias_Exclui_EmTrans = SUCESSO

    Exit Function

Erro_TRVOcorrencias_Exclui_EmTrans:

    TRVOcorrencias_Exclui_EmTrans = gErr

    Select Case gErr

        Case 189444
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 189445, 189446
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVOCORRENCIAS", gErr)

        Case 189447
            Call Rotina_Erro(vbOKOnly, "ERRO_TRVOCORRENCIAS_NAO_CADASTRADO", gErr, objTRVOcorrencias.lCodigo)

        Case 189448
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRVOCORRENCIAS", gErr)

        Case 189449
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRVOCORRENCIAS", gErr)
            
        Case 189485, 192062, 196381, 196382

        Case 194437, 194438
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVVOUCHERS", gErr)

        Case 194439
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRVVOUCHERS", gErr)

        Case 194440
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRVVOUCHERS", gErr)

        Case 194441
            Call Rotina_Erro(vbOKOnly, "ERRO_TRVVOUCHERS_NAO_CADASTRADO", gErr, lNumVou, sTipoDoc, sSerie)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189481)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Private Function Item_Exclui_Com_Destino(ByVal iTipoDestino As Integer, ByVal lNumIntDocDestino As Long) As Long

Dim lErro As Long
Dim objDebitoRecCli As New ClassDebitoRecCli
Dim objContabil As New ClassContabil
Dim objCreditoPagar As New ClassCreditoPagar

On Error GoTo Erro_Item_Exclui_Com_Destino
    
    If lNumIntDocDestino <> 0 Then
    
        Select Case iTipoDestino
    
            Case TRV_TIPO_DOC_DESTINO_DEBCLI
                       
                Call objContabil.Contabil_Inicializa_Contabilidade4(29, MODULO_CONTASARECEBER)
                
                objDebitoRecCli.lNumIntDoc = lNumIntDocDestino
                
                'Verifica se o debito a receber existe
                lErro = CF("DebitoReceber_Le", objDebitoRecCli)
                If lErro <> SUCESSO And lErro <> 17835 Then gError 192059
                
                lErro = CF("DebitoRecCli_Exclui_EmTransacao", objDebitoRecCli, objContabil)
                If lErro <> SUCESSO Then gError 192059
        
            Case TRV_TIPO_DOC_DESTINO_CREDFORN
            
                Call objContabil.Contabil_Inicializa_Contabilidade4(28, MODULO_CONTASAPAGAR)
        
                objCreditoPagar.lNumIntDoc = lNumIntDocDestino
                
                'Verifica se o credito a pagar existe
                lErro = CF("CreditoPagar_Le", objCreditoPagar)
                If lErro <> SUCESSO And lErro <> 17071 Then gError 192060
                
                lErro = CF("CreditoPagar_Exclui_EmTransacao", objCreditoPagar, objContabil)
                If lErro <> SUCESSO Then gError 192060
        
            Case Else
                gError 192058
    
        End Select
        
    End If

    Item_Exclui_Com_Destino = SUCESSO

    Exit Function

Erro_Item_Exclui_Com_Destino:

    Item_Exclui_Com_Destino = gErr

    Select Case gErr

        Case 192059, 192060
        
        Case 192058
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEM_JA_FATURADO", gErr, iTipoDestino, lNumIntDocDestino)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192061)

    End Select

    Exit Function

End Function

Public Function CamposGenericos_ValidaExclusao_TRVOCROrigem() As Long
'Verifica se é válida a exclusão de valores para o campo Origem

Dim alComando(1 To 2) As Long
Dim lErro As Long
Dim lCodigo As Long
Dim iOrigem As Integer
Dim sValor As String
Dim iIndice As Integer

On Error GoTo Erro_CamposGenericos_ValidaExclusao_TRVOCROrigem
    
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 190154
    Next
    
    'Busca nas ocorrências origens que não existam na tabela CamposGenericos Valores
    lErro = Comando_Executar(alComando(1), "SELECT Codigo, Origem FROM TRVOcorrencias WHERE Origem NOT IN (SELECT CodValor FROM CamposGenericosValores WHERE CodCampo = ?) ", lCodigo, iOrigem, CAMPOSGENERICOS_TRVORIGEM_OCR)
    If lErro <> AD_SQL_SUCESSO Then gError 190155
    
    'Posiciona no primeiro registro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190156
    
    'Se encontrou algum registro => erro de inconsistência
    If lErro = AD_SQL_SUCESSO Then gError 190157
    
    sValor = String(STRING_CAMPOSGENERICOSVALORES_VALOR, 0)
    
    'Busca nas ocorrências origens que não existam na tabela CamposGenericos Valores
    lErro = Comando_Executar(alComando(2), "SELECT Valor FROM CamposGenericosValores WHERE CodCampo = ? AND CodValor = ? ", sValor, CAMPOSGENERICOS_TRVORIGEM_OCR, INATIVACAO_AUTOMATICA_CODIGO)
    If lErro <> AD_SQL_SUCESSO Then gError 190555
    
    'Posiciona no primeiro registro
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190556
    
    'O código 1 é obrigatório pois é utilizado nas inativações automáticas
    If lErro <> AD_SQL_SUCESSO Then gError 190557
    
    If sValor <> INATIVACAO_AUTOMATICA_TEXTO Then gError 190558
    
    'Busca nas ocorrências origens que não existam na tabela CamposGenericos Valores
    lErro = Comando_Executar(alComando(2), "SELECT Valor FROM CamposGenericosValores WHERE CodCampo = ? AND CodValor = ? ", sValor, CAMPOSGENERICOS_TRVORIGEM_OCR, IMPORTACAO_OCORRENCIA_CODIGO)
    If lErro <> AD_SQL_SUCESSO Then gError 194430
    
    'Posiciona no primeiro registro
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194431
    
    'O código 2 é obrigatório pois é utilizado nas importações de ocorrências
    If lErro <> AD_SQL_SUCESSO Then gError 194432
    
    If sValor <> IMPORTACAO_OCORRENCIA_TEXTO Then gError 194433
    
    'Fecha comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    CamposGenericos_ValidaExclusao_TRVOCROrigem = SUCESSO
    
    Exit Function
    
Erro_CamposGenericos_ValidaExclusao_TRVOCROrigem:

    CamposGenericos_ValidaExclusao_TRVOCROrigem = gErr
    
    Select Case gErr
    
        Case 190154
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 190155, 190156, 190555, 190556, 194430, 194431
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CAMPOSGENERICOS", gErr, CAMPOSGENERICOS_TRVORIGEM_OCR)
        
        Case 190157
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ORIGEM_OCR", gErr, iOrigem, lCodigo)
            
        Case 190557
            Call Rotina_Erro(vbOKOnly, "ERRO_ORIGEMOCR_CODIGO_OBRIGATORIO", gErr, INATIVACAO_AUTOMATICA_CODIGO)
        
        Case 190558
            Call Rotina_Erro(vbOKOnly, "ERRO_ORIGEMOCR_TEXTO_OBRIGATORIO", gErr, INATIVACAO_AUTOMATICA_CODIGO, INATIVACAO_AUTOMATICA_TEXTO)
            
        Case 194432
            Call Rotina_Erro(vbOKOnly, "ERRO_ORIGEMOCR_CODIGO_OBRIGATORIO", gErr, IMPORTACAO_OCORRENCIA_CODIGO)
        
        Case 194433
            Call Rotina_Erro(vbOKOnly, "ERRO_ORIGEMOCR_TEXTO_OBRIGATORIO", gErr, IMPORTACAO_OCORRENCIA_CODIGO, IMPORTACAO_OCORRENCIA_TEXTO)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190158)
    
    End Select

    'Fecha comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Public Function CamposGenericos_ValidaExclusao_TRVOCRTipo() As Long
'Verifica se é válida a exclusão de valores para o campo Tipo

Dim alComando(1 To 2) As Long
Dim lErro As Long
Dim lCodigo As Long
Dim iTipo As Integer
Dim iTipoAux As Integer
Dim sTextoAux As String
Dim iIndice As Integer
Dim sValor As String

On Error GoTo Erro_CamposGenericos_ValidaExclusao_TRVOCRTipo
    
    'Abre comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 190159
    Next
    
    'Busca nas ocorrências origens que não existam na tabela CamposGenericos Valores
    lErro = Comando_Executar(alComando(1), "SELECT O.Codigo, D.Tipo FROM TRVOcorrenciaDet AS D, TRVOcorrencias AS O WHERE O.NumIntDoc = D.NumIntDocOcr AND D.Tipo NOT IN (SELECT CodValor FROM CamposGenericosValores WHERE CodCampo = ?) ", lCodigo, iTipo, CAMPOSGENERICOS_TRVTIPODET_OCR)
    If lErro <> AD_SQL_SUCESSO Then gError 190160
    
    'Posiciona no primeiro registro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190161
    
    'Se encontrou algum registro => erro de inconsistência
    If lErro = AD_SQL_SUCESSO Then gError 190162
    
    For iIndice = 1 To 6
    
        Select Case iIndice
        
            Case 1
                iTipoAux = INATIVACAO_AUTOMATICA_TIPO_NVL_CODIGO
                sTextoAux = INATIVACAO_AUTOMATICA_TIPO_NVL_TEXTO
            Case 2
                iTipoAux = INATIVACAO_AUTOMATICA_TIPO_TAR_CODIGO
                sTextoAux = INATIVACAO_AUTOMATICA_TIPO_TAR_TEXTO
            Case 3
                iTipoAux = INATIVACAO_AUTOMATICA_TIPO_PIS_CODIGO
                sTextoAux = INATIVACAO_AUTOMATICA_TIPO_PIS_TEXTO
            Case 4
                iTipoAux = INATIVACAO_AUTOMATICA_TIPO_COFINS_CODIGO
                sTextoAux = INATIVACAO_AUTOMATICA_TIPO_COFINS_TEXTO
            Case 5
                iTipoAux = INATIVACAO_AUTOMATICA_TIPO_ISS_CODIGO
                sTextoAux = INATIVACAO_AUTOMATICA_TIPO_ISS_TEXTO
            Case 6
                iTipoAux = IMPORTACAO_OCORRENCIA_TIPO_IMP_CODIGO
                sTextoAux = IMPORTACAO_OCORRENCIA_TIPO_IMP_TEXTO

        End Select
        
        sValor = String(STRING_CAMPOSGENERICOSVALORES_VALOR, 0)
    
        'Busca nas ocorrências origens que não existam na tabela CamposGenericos Valores
        lErro = Comando_Executar(alComando(2), "SELECT Valor FROM CamposGenericosValores WHERE CodCampo = ? AND CodValor = ? ", sValor, CAMPOSGENERICOS_TRVTIPODET_OCR, iTipoAux)
        If lErro <> AD_SQL_SUCESSO Then gError 190559
        
        'Posiciona no primeiro registro
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190560
        
        'O código 1 é obrigatório pois é utilizado nas inativações automáticas
        If lErro <> AD_SQL_SUCESSO Then gError 190561
        
        If sValor <> sTextoAux Then gError 190562
        
    Next
    
    'Fecha comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    CamposGenericos_ValidaExclusao_TRVOCRTipo = SUCESSO
    
    Exit Function
    
Erro_CamposGenericos_ValidaExclusao_TRVOCRTipo:

    CamposGenericos_ValidaExclusao_TRVOCRTipo = gErr
    
    Select Case gErr
    
        Case 190159
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 190160, 190161, 190559, 190560
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CAMPOSGENERICOS", gErr, CAMPOSGENERICOS_TRVTIPODET_OCR)
        
        Case 190162
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TIPO_OCR", gErr, iTipo, lCodigo)
            
        Case 190561
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOOCR_CODIGO_OBRIGATORIO", gErr, iTipoAux)
        
        Case 190562
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOOCR_TEXTO_OBRIGATORIO", gErr, iTipoAux, sTextoAux)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190163)
    
    End Select

    'Fecha comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Public Function CamposGenericos_ValidaExclusao_TRVAporteTipo() As Long
'Verifica se é válida a exclusão de valores para o campo Tipo

Dim lComando As Long
Dim lErro As Long
Dim lCodigo As Long
Dim iTipo As Integer

On Error GoTo Erro_CamposGenericos_ValidaExclusao_TRVAporteTipo
    
    'Abre comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 190550
    
    'Busca nas ocorrências origens que não existam na tabela CamposGenericos Valores
    lErro = Comando_Executar(lComando, "SELECT Codigo, Tipo FROM TRVAportes WHERE Tipo NOT IN (SELECT CodValor FROM CamposGenericosValores WHERE CodCampo = ?) ", lCodigo, iTipo, CAMPOSGENERICOS_TRVTIPOAPORTE_OCR)
    If lErro <> AD_SQL_SUCESSO Then gError 190551
    
    'Posiciona no primeiro registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190552
    
    'Se encontrou algum registro => erro de inconsistência
    If lErro = AD_SQL_SUCESSO Then gError 190553
    
    'Fecha comando
    Call Comando_Fechar(lComando)
    
    CamposGenericos_ValidaExclusao_TRVAporteTipo = SUCESSO
    
    Exit Function
    
Erro_CamposGenericos_ValidaExclusao_TRVAporteTipo:

    CamposGenericos_ValidaExclusao_TRVAporteTipo = gErr
    
    Select Case gErr
    
        Case 190550
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 190551, 190552
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CAMPOSGENERICOS", gErr, CAMPOSGENERICOS_TRVTIPODET_OCR)
        
        Case 190553
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TIPO_APORTE", gErr, iTipo, lCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190554)
    
    End Select

    'Fecha comando
    Call Comando_Fechar(lComando)

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'GRAVAÇÃO
Public Function TRVOcorrencias_Libera(ByVal colOcorrencias As Collection) As Long

Dim lErro As Long
Dim lTransacao As Long
Dim objTRVOcorrencias As ClassTRVOcorrencias
Dim alComando(0 To 4) As Long
Dim iIndice As Integer

On Error GoTo Erro_TRVOcorrencias_Libera

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 190306
    
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 190312
    Next
    
    For Each objTRVOcorrencias In colOcorrencias
    
        lErro = TRVOcorrencias_Grava_EmTrans(objTRVOcorrencias)
        If lErro <> SUCESSO Then gError 190307

        lErro = TRVOcorrencias_Libera_EmTrans(alComando, objTRVOcorrencias)
        If lErro <> SUCESSO Then gError 190307
        
    Next
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 190308

    TRVOcorrencias_Libera = SUCESSO

    Exit Function

Erro_TRVOcorrencias_Libera:

    TRVOcorrencias_Libera = gErr

    Select Case gErr

        Case 190306
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 190307

        Case 190308
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case 190312
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190309)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function TRVOcorrencias_Libera_EmTrans(alComando() As Long, ByVal objTRVOcorrencias As ClassTRVOcorrencias) As Long

Dim lErro As Long
Dim lNumIntDoc As Long
Dim objDebitoRecCli As New ClassDebitoRecCli
Dim colInfoComissao As New colInfoComissao
Dim objContabil As ClassContabil
Dim objCreditoPagar As New ClassCreditoPagar

On Error GoTo Erro_TRVOcorrencias_Libera_EmTrans

    objTRVOcorrencias.iStatus = STATUS_TRV_OCR_LIBERADO

    If objTRVOcorrencias.iFormaPagto = FORMAPAGTO_TRV_OCR_FAT Then
        'Não precisa fazer nada => O Tratamento é dado na tela de Faturamento
    Else
        'Gera o crédito e informa o número do numintdoc do débito
    
        If objTRVOcorrencias.dValorTotal < 0 Then
        
            objDebitoRecCli.iFilialEmpresa = giFilialEmpresa
            objDebitoRecCli.lCliente = objTRVOcorrencias.lCliente
            objDebitoRecCli.iFilial = objTRVOcorrencias.iFilialCliente
            objDebitoRecCli.sSiglaDocumento = "NCR" 'Nota de crédito
            objDebitoRecCli.lNumTitulo = objTRVOcorrencias.lCodigo
            objDebitoRecCli.dtDataEmissao = objTRVOcorrencias.dtDataEmissao
            objDebitoRecCli.dValorTotal = objTRVOcorrencias.dValorTotal
            objDebitoRecCli.dValorProdutos = objTRVOcorrencias.dValorTotal
            
            lErro = CF("DebitoRecCli_Grava_EmTrans", objDebitoRecCli, colInfoComissao, objContabil)
            If lErro <> SUCESSO Then gError 190315
            
            objTRVOcorrencias.lNumIntDocDestino = objDebitoRecCli.lNumIntDoc
            objTRVOcorrencias.iTipoDocDestino = TRV_TIPO_DOC_DESTINO_DEBCLI
            objTRVOcorrencias.lNumDocDestino = objDebitoRecCli.lNumTitulo
        
        Else
        
            objCreditoPagar.iFilialEmpresa = giFilialEmpresa
            objCreditoPagar.lFornecedor = objTRVOcorrencias.lCliente
            objCreditoPagar.iFilial = objTRVOcorrencias.iFilialCliente
            objCreditoPagar.sSiglaDocumento = "NCP" 'Nota de crédito
            objCreditoPagar.lNumTitulo = objTRVOcorrencias.lCodigo
            objCreditoPagar.dtDataEmissao = objTRVOcorrencias.dtDataEmissao
            objCreditoPagar.dValorTotal = Abs(objTRVOcorrencias.dValorTotal)
            objCreditoPagar.dValorProdutos = Abs(objTRVOcorrencias.dValorTotal)
            
            lErro = CF("CreditoPagar_Grava_EmTrans", objCreditoPagar, objContabil)
            If lErro <> SUCESSO Then gError 190325
            
            objTRVOcorrencias.lNumIntDocDestino = objCreditoPagar.lNumIntDoc
            objTRVOcorrencias.iTipoDocDestino = TRV_TIPO_DOC_DESTINO_CREDFORN
            objTRVOcorrencias.lNumDocDestino = objCreditoPagar.lNumTitulo
    
        End If
    
    End If
    
    'Le a tabelaTRVOcorrencias
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM TRVOcorrencias WHERE Codigo= ? ", _
                0, lNumIntDoc, objTRVOcorrencias.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 190316

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190317

    'Se existir => UPDATE, senão => INSERT
    If lErro <> AD_SQL_SUCESSO Then gError 190318

    'Trava registro contra alterações/Leituras
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> SUCESSO Then gError 190319
    
    objTRVOcorrencias.lNumIntDoc = lNumIntDoc

    'Altera o Status da ocorrencia e a forma de pagamento
    lErro = Comando_ExecutarPos(alComando(1), "UPDATE TRVOcorrencias SET Status= ? , FormaPagto= ?, NumIntDocDestino = ?, DataLiberacao = ? , Usuario = ?, TipoDocDestino = ?, ValorTotal = ?, NumTitulo = ? ", alComando(0), objTRVOcorrencias.iStatus, objTRVOcorrencias.iFormaPagto, objTRVOcorrencias.lNumIntDocDestino, Date, gsUsuario, objTRVOcorrencias.iTipoDocDestino, objTRVOcorrencias.dValorTotal, objTRVOcorrencias.lNumDocDestino)
    If lErro <> AD_SQL_SUCESSO Then gError 190320
    
    lErro = TRVOcorrenciaDet_Exclui(alComando(2), alComando(3), objTRVOcorrencias)
    If lErro <> SUCESSO Then gError 190321
    
    lErro = TRVOcorrenciaDet_Insere(alComando(4), objTRVOcorrencias)
    If lErro <> SUCESSO Then gError 190322

    TRVOcorrencias_Libera_EmTrans = SUCESSO

    Exit Function

Erro_TRVOcorrencias_Libera_EmTrans:

    TRVOcorrencias_Libera_EmTrans = gErr

    Select Case gErr
        
        Case 190316, 190317
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVOCORRENCIAS", gErr)
            
        Case 190318
            Call Rotina_Erro(vbOKOnly, "ERRO_TRVOCORRENCIAS_NAO_CADASTRADO", gErr, objTRVOcorrencias.lCodigo)

        Case 190319
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRVOCORRENCIAS", gErr)

        Case 190320
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRVOCORRENCIAS", gErr)
            
        Case 190321, 190322, 190325
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190323)

    End Select

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'GRAVAÇÃO
Public Function TRVAportes_Grava(ByVal objTRVAportes As ClassTRVAportes) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_TRVAportes_Grava

    lErro = TRVAportes_Valida(objTRVAportes)
    If lErro <> SUCESSO Then gError 196478

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 190348

    lErro = CF("TRVAportes_Grava_EmTrans", objTRVAportes)
    If lErro <> SUCESSO Then gError 190349

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 190350

    TRVAportes_Grava = SUCESSO

    Exit Function

Erro_TRVAportes_Grava:

    TRVAportes_Grava = gErr

    Select Case gErr

        Case 190348
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 190349, 196478

        Case 190350
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190351)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function TRVAportes_Grava_EmTrans(ByVal objTRVAportes As ClassTRVAportes) As Long

Dim lErro As Long
Dim alComando(0 To 20) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim iProxCmd As Integer
Dim iMoeda As Integer

On Error GoTo Erro_TRVAportes_Grava_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 190352
    Next

    'Le a tabelaTRVAportes
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc, Moeda FROM TRVAportes WHERE Codigo= ? ", _
                0, lNumIntDoc, iMoeda, objTRVAportes.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 190353

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190354

    'Se existir => UPDATE, senão => INSERT
    If lErro = AD_SQL_SUCESSO Then

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 190355
        
        objTRVAportes.lNumIntDoc = lNumIntDoc
        
        If iMoeda <> objTRVAportes.iMoeda Then gError 192887

        lErro = Comando_ExecutarPos(alComando(1), "UPDATE TRVAportes SET Cliente= ? , FilialCliente= ? , DataEmissao= ? , Observacao= ? , " & _
                "Historico= ? , Tipo= ? , Moeda= ? , PrevValor= ? , PrevDataDe= ? , " & _
                "PrevDataAte= ? ", alComando(0), _
                objTRVAportes.lCliente, objTRVAportes.iFilialCliente, objTRVAportes.dtDataEmissao, objTRVAportes.sObservacao, _
                objTRVAportes.sHistorico, objTRVAportes.iTipo, objTRVAportes.iMoeda, objTRVAportes.dPrevValor, objTRVAportes.dtPrevDataDe, _
                objTRVAportes.dtPrevDataAte)
        If lErro <> AD_SQL_SUCESSO Then gError 190356

    Else

        'Obter NumIntDoc
        lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_INT_PROX_TRVAPORTES", lNumIntDoc)
        If lErro <> SUCESSO Then gError 190357

        objTRVAportes.lNumIntDoc = lNumIntDoc

        lErro = Comando_Executar(alComando(1), "INSERT INTO TRVAportes( NumIntDoc, Codigo, Cliente, FilialCliente, " & _
                "DataEmissao, Observacao, Historico, Tipo, Moeda, " & _
                "PrevValor, PrevDataDe, PrevDataAte) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)", _
                objTRVAportes.lNumIntDoc, objTRVAportes.lCodigo, objTRVAportes.lCliente, objTRVAportes.iFilialCliente, _
                objTRVAportes.dtDataEmissao, objTRVAportes.sObservacao, objTRVAportes.sHistorico, objTRVAportes.iTipo, objTRVAportes.iMoeda, _
                objTRVAportes.dPrevValor, objTRVAportes.dtPrevDataDe, objTRVAportes.dtPrevDataAte)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190358

    End If
    
    iProxCmd = 1
    
    lErro = TRVAportes_Grava_Pagtos(alComando, objTRVAportes, iProxCmd)
    If lErro <> SUCESSO Then gError 190359

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TRVAportes_Grava_EmTrans = SUCESSO

    Exit Function

Erro_TRVAportes_Grava_EmTrans:

    TRVAportes_Grava_EmTrans = gErr

    Select Case gErr

        Case 190352
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 190353, 190354
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVAPORTES", gErr)

        Case 190355
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRVAPORTES", gErr)

        Case 190356
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRVAPORTES", gErr)

        Case 190357, 190359
        
        Case 190358
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRVAPORTES", gErr)
            
        Case 192887
            Call Rotina_Erro(vbOKOnly, "ERRO_TRVAPORTES_TROCA_MOEDA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190360)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function TRVAportes_Grava_Pagtos(alComando() As Long, ByVal objAporte As ClassTRVAportes, iProxCmd As Integer) As Long

Dim lErro As Long
Dim objAporteBD As New ClassTRVAportes
Dim bAchou As Boolean
Dim bExisteDestino As Boolean
Dim lNumTitulo As Long
Dim sDoc As String
Dim iIndice As Integer
Dim colColecaoBD As Collection
Dim colColecao As Collection
Dim objObjetoBD As Object
Dim objObjeto As Object
Dim sPrefixoFunc As String
Dim iSeq As Integer

On Error GoTo Erro_TRVAportes_Grava_Pagtos

    'A tela tem que manter os numintdoc nos pagtos
    
    objAporteBD.lCodigo = objAporte.lCodigo
    
    'Le o aporte do BD
    lErro = CF("TRVAportes_Le_Aux", alComando, objAporteBD, iProxCmd, True, True)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 190365
    
    '1 para cada tipo de pagamento
    'Para evitar código duplicado
    For iIndice = 1 To 4
    
        iSeq = 0
    
        If iIndice = 1 Then
            Set colColecaoBD = objAporteBD.colPagtoCondicionados
            Set colColecao = objAporte.colPagtoCondicionados
            sPrefixoFunc = "TRVAportePagtoCond"
        ElseIf iIndice = 2 Then
            Set colColecaoBD = objAporteBD.colPagtoDireto
            Set colColecao = objAporte.colPagtoDireto
            sPrefixoFunc = "TRVAportePagtoDireto"
        ElseIf iIndice = 3 Then
            Set colColecaoBD = objAporteBD.colPagtoSobreFaturaCond
            Set colColecao = objAporte.colPagtoSobreFaturaCond
            sPrefixoFunc = "TRVAportePagtoFatCond"
        Else
            Set colColecaoBD = objAporteBD.colPagtoSobreFatura
            Set colColecao = objAporte.colPagtoSobreFatura
            sPrefixoFunc = "TRVAportePagtoFat"
        End If
    
        For Each objObjeto In colColecao
        
            iSeq = iSeq + 1
        
            objObjeto.iSeq = iSeq
        
            objObjeto.lNumIntDocAporte = objAporte.lNumIntDoc
        
            If objObjeto.lNumIntDoc = 0 Then
                lErro = CF(sPrefixoFunc & "_Insere", alComando(iProxCmd), objObjeto, objAporte)
                iProxCmd = iProxCmd + 1
                If lErro <> SUCESSO Then gError 190366
            Else
            
                bAchou = False
                For Each objObjetoBD In colColecaoBD
                    If objObjeto.lNumIntDoc = objObjetoBD.lNumIntDoc Then
                        bAchou = True
                        Exit For
                    End If
                Next
                If Not bAchou Then gError 190367 ' NumIntDoc inexistente no BD
                
                If iIndice < 3 Then
                    'Esse código é apenas para reforçar o código da tela que já
                    'não deixa excluir ou alterar um item que já tenha originado outro
                    lErro = CF("Verifica_Existencia_Doc_TRV", objObjetoBD.lNumIntDocDestino, objObjetoBD.iTipoDocDestino, bExisteDestino, lNumTitulo, sDoc)
                    If lErro <> SUCESSO Then gError 190368
                Else
                    bExisteDestino = False
                End If
                
                If Not bExisteDestino Then
                    iProxCmd = iProxCmd + 1
                    lErro = CF(sPrefixoFunc & "_Atualiza", alComando(iProxCmd - 1), alComando(iProxCmd), objObjeto)
                    iProxCmd = iProxCmd + 1
                    If lErro <> SUCESSO Then gError 190369
                End If
            End If
        Next
    
        For Each objObjetoBD In colColecaoBD
            bAchou = False
            For Each objObjeto In colColecao
                If objObjeto.lNumIntDoc = objObjetoBD.lNumIntDoc Then
                    bAchou = True
                    Exit For
                End If
            Next
            'Tinha antes e já não tem mais
            If Not bAchou Then
            
                If iIndice < 3 Then
                    'Esse código é apenas para reforçar o código da tela que já
                    'não deixa excluir ou alterar um item que já tenha originado outro
                    lErro = CF("Verifica_Existencia_Doc_TRV", objObjetoBD.lNumIntDocDestino, objObjetoBD.iTipoDocDestino, bExisteDestino, lNumTitulo, sDoc)
                    If lErro <> SUCESSO Then gError 190370
                
                    If bExisteDestino And objObjetoBD.iTipoDocDestino <> TRV_TIPO_DOC_DESTINO_DEBCLI And objObjetoBD.iTipoDocDestino <> TRV_TIPO_DOC_DESTINO_CREDFORN Then gError 190371
                Else
                    bExisteDestino = False
                End If
                
                
                iProxCmd = iProxCmd + 1
                lErro = CF(sPrefixoFunc & "_Exclui", alComando(iProxCmd - 1), alComando(iProxCmd), objObjetoBD)
                iProxCmd = iProxCmd + 1
                If lErro <> SUCESSO Then gError 190372
    
            End If
        Next
        
    Next
    
    TRVAportes_Grava_Pagtos = SUCESSO

    Exit Function

Erro_TRVAportes_Grava_Pagtos:

    TRVAportes_Grava_Pagtos = gErr

    Select Case gErr
    
        Case 190365, 190366, 190368 To 190370, 190372
        
        Case 190367
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMINTDOC_NAO_CADASTRADO_TABELA", gErr, objObjeto.lNumIntDoc, sPrefixoFunc)
        
        Case 190371
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_EXISTE_DOC_DESTINO", gErr, sDoc)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190373)

    End Select

    Exit Function

End Function

Public Function Verifica_Existencia_Doc_TRV(ByVal lNumIntDoc As Long, ByVal iTipo As Integer, bExiste As Boolean, lNumTitulo As Long, sDoc As String) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim sTabela As String
Dim sSQL As String
Dim sDocAux As String

On Error GoTo Erro_Verifica_Existencia_Doc_TRV

    bExiste = False
    sDoc = ""

    If iTipo <> 0 And lNumIntDoc <> 0 Then

        'Abertura de Comando
        For iIndice = LBound(alComando) To UBound(alComando)
            alComando(iIndice) = Comando_Abrir()
            If alComando(iIndice) = 0 Then gError 190361
        Next
        
        Select Case iTipo
        
            Case TRV_TIPO_DOC_DESTINO_CREDFORN
                sTabela = TRV_TIPO_DOC_DESTINO_CREDFORN_TABELA
                sDocAux = TRV_TIPO_DOC_DESTINO_CREDFORN_TEXTO
                
            Case TRV_TIPO_DOC_DESTINO_DEBCLI
                sTabela = TRV_TIPO_DOC_DESTINO_DEBCLI_TABELA
                sDocAux = TRV_TIPO_DOC_DESTINO_DEBCLI_TEXTO
        
            Case TRV_TIPO_DOC_DESTINO_TITPAG
                sTabela = TRV_TIPO_DOC_DESTINO_TITPAG_TABELA
                sDocAux = TRV_TIPO_DOC_DESTINO_TITPAG_TEXTO
        
            Case TRV_TIPO_DOC_DESTINO_TITREC
                sTabela = TRV_TIPO_DOC_DESTINO_TITREC_TABELA
                sDocAux = TRV_TIPO_DOC_DESTINO_TITREC_TEXTO
        
            Case TRV_TIPO_DOC_DESTINO_NFSPAG
                sTabela = TRV_TIPO_DOC_DESTINO_NFSPAG_TABELA
                sDocAux = TRV_TIPO_DOC_DESTINO_NFSPAG_TEXTO
        
        End Select
        
        sDocAux = sDocAux & " n° "
    
        sSQL = "SELECT NumTitulo FROM " & sTabela & " WHERE NumIntDoc = ? "
    
        lErro = Comando_Executar(alComando(0), sSQL, lNumTitulo, lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 190362
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190363
    
        If lErro <> AD_SQL_SEM_DADOS Then
            bExiste = True
            sDoc = sDocAux & CStr(lNumTitulo)
        End If
    
    End If

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Verifica_Existencia_Doc_TRV = SUCESSO

    Exit Function

Erro_Verifica_Existencia_Doc_TRV:

    Verifica_Existencia_Doc_TRV = gErr

    Select Case gErr

        Case 190361
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 190362 To 190363
            Call Rotina_Erro(vbOKOnly, "ERRO_PREPARACAO_COMANDO_SQL", gErr, sSQL)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190364)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function TRVAportePagtoCond_Atualiza(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal objTRVAportePagtoCond As ClassTRVAportePagtoCond) As Long

Dim lErro As Long
Dim tPagto As typeTRVAportePagtoCond
Dim lNumIntDoc As Long
Dim iStatus As Integer

On Error GoTo Erro_TRVAportePagtoCond_Atualiza

    'Le a tabelaTRVAportePagtoCond
    lErro = Comando_ExecutarPos(lComando1, "SELECT NumIntDoc, Status FROM TRVAportePagtoCond WHERE NumIntDoc = ? ", 0, lNumIntDoc, iStatus, objTRVAportePagtoCond.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 190385

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190386

    If lErro <> AD_SQL_SUCESSO Then gError 190387
    
    'If iStatus <> STATUS_TRV_OCR_BLOQUEADO Then gError 190563
    
    If iStatus = STATUS_TRV_OCR_BLOQUEADO Then
    
        lErro = Comando_ExecutarPos(lComando2, "UPDATE TRVAportePagtoCond SET NumIntDocAporte= ? , Base= ? , DataPagto= ? , " & _
                "NumIntDocDestino= ? , FormaPagto= ? , TipoDocDestino= ? , Percentual= ?, Seq = ? ", lComando1, _
                objTRVAportePagtoCond.lNumIntDocAporte, objTRVAportePagtoCond.iBase, objTRVAportePagtoCond.dtDataPagto, _
                objTRVAportePagtoCond.lNumIntDocDestino, objTRVAportePagtoCond.iFormaPagto, objTRVAportePagtoCond.iTipoDocDestino, objTRVAportePagtoCond.dPercentual, objTRVAportePagtoCond.iSeq)
        
    Else
        'Se já estiver liberado só deixa apontar para o documento de destino
        lErro = Comando_ExecutarPos(lComando2, "UPDATE TRVAportePagtoCond SET NumIntDocDestino= ? , TipoDocDestino= ? ", lComando1, _
                objTRVAportePagtoCond.lNumIntDocDestino, objTRVAportePagtoCond.iTipoDocDestino)
    
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 190388

    TRVAportePagtoCond_Atualiza = SUCESSO

    Exit Function

Erro_TRVAportePagtoCond_Atualiza:

    TRVAportePagtoCond_Atualiza = gErr

    Select Case gErr

        Case 190385, 190386
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVAPORTEPAGTOCOND", gErr)
            
        Case 190387
            Call Rotina_Erro(vbOKOnly, "ERRO_TRVAPORTEPAGTOCOND_NAO_CADASTRADO", gErr, objTRVAportePagtoCond.lNumIntDoc)

        Case 190388
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRVAPORTEPAGTOCOND", gErr)
            
        Case 190563
            Call Rotina_Erro(vbOKOnly, "ERRO_TRVAPORTEPAGTOCOND_STATUS_DIF_BLOQUEADO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190389)

    End Select


    Exit Function

End Function

Public Function TRVAportePagtoCond_Insere(ByVal lComando As Long, ByVal objTRVAportePagtoCond As ClassTRVAportePagtoCond, ByVal objAporte As ClassTRVAportes) As Long

Dim lErro As Long
Dim lNumIntDoc As Long

On Error GoTo Erro_TRVAportePagtoCond_Insere

    'Obter NumIntDoc
    lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_INT_PROX_TRVAPORTEPAGTOCOND", lNumIntDoc)
    If lErro <> SUCESSO Then gError 190382

    objTRVAportePagtoCond.lNumIntDoc = lNumIntDoc
        
    lErro = Comando_Executar(lComando, "INSERT INTO TRVAportePagtoCond( NumIntDoc, NumIntDocAporte, Base, Status, " & _
            "DataPagto, NumIntDocDestino, FormaPagto, TipoDocDestino, Percentual, Valor, Seq) VALUES (?,?,?,?,?,?,?,?,?,?,?)", _
            objTRVAportePagtoCond.lNumIntDoc, objTRVAportePagtoCond.lNumIntDocAporte, objTRVAportePagtoCond.iBase, STATUS_TRV_OCR_BLOQUEADO, _
            objTRVAportePagtoCond.dtDataPagto, objTRVAportePagtoCond.lNumIntDocDestino, objTRVAportePagtoCond.iFormaPagto, objTRVAportePagtoCond.iTipoDocDestino, objTRVAportePagtoCond.dPercentual, objTRVAportePagtoCond.dValor, objTRVAportePagtoCond.iSeq)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190383


    TRVAportePagtoCond_Insere = SUCESSO

    Exit Function

Erro_TRVAportePagtoCond_Insere:

    TRVAportePagtoCond_Insere = gErr

    Select Case gErr
    
        Case 190382

        Case 190383
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRVAPORTEPAGTOCOND", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190384)

    End Select

    Exit Function

End Function

Public Function TRVAportePagtoCond_Exclui(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal objTRVAportePagtoCond As ClassTRVAportePagtoCond) As Long

Dim lErro As Long
Dim lNumIntDoc As Long
Dim iTipoDestino As Integer
Dim lNumIntDocDestino As Long

On Error GoTo Erro_TRVAportePagtoCond_Exclui

    'Le a tabelaTRVAportePagtoCond
    lErro = Comando_ExecutarPos(lComando1, "SELECT NumIntDoc, TipoDocDestino, NumIntDocDestino FROM TRVAportePagtoCond WHERE NumIntDoc = ? ", 0, lNumIntDoc, iTipoDestino, lNumIntDocDestino, objTRVAportePagtoCond.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 190390

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190391

    'Se não existir => ERRO
    If lErro <> AD_SQL_SUCESSO Then gError 190392
    
    lErro = Comando_ExecutarPos(lComando2, "DELETE FROM TRVAportePagtoCond ", lComando1)
    If lErro <> AD_SQL_SUCESSO Then gError 190393
    
    lErro = Item_Exclui_Com_Destino(iTipoDestino, lNumIntDocDestino)
    If lErro <> SUCESSO Then gError 192063

    TRVAportePagtoCond_Exclui = SUCESSO

    Exit Function

Erro_TRVAportePagtoCond_Exclui:

    TRVAportePagtoCond_Exclui = gErr

    Select Case gErr

        Case 190390, 190391
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVAPORTEPAGTOCOND", gErr)
            
        Case 190392
            Call Rotina_Erro(vbOKOnly, "ERRO_TRVAPORTEPAGTOCOND_NAO_CADASTRADO", gErr, objTRVAportePagtoCond.lNumIntDoc)

        Case 190393
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRVAPORTEPAGTOCOND", gErr)
            
        Case 192063

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190394)

    End Select


    Exit Function

End Function

Public Function TRVAportePagtoDireto_Exclui(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal objTRVAportePagtoDireto As ClassTRVAportePagtoDireto) As Long

Dim lErro As Long
Dim lNumIntDoc As Long
Dim iTipoDestino As Integer
Dim lNumIntDocDestino As Long

On Error GoTo Erro_TRVAportePagtoDireto_Exclui

    'Le a tabelaTRVAportePagtoDireto
    lErro = Comando_ExecutarPos(lComando1, "SELECT NumIntDoc, TipoDocDestino, NumIntDocDestino FROM TRVAportePagtoDireto WHERE NumIntDoc= ? ", _
                0, lNumIntDoc, iTipoDestino, lNumIntDocDestino, objTRVAportePagtoDireto.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 190403

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190404

    'Se não existir => ERRO
    If lErro <> AD_SQL_SUCESSO Then gError 190405

    lErro = Comando_ExecutarPos(lComando2, "DELETE FROM TRVAportePagtoDireto  ", lComando1)
    If lErro <> AD_SQL_SUCESSO Then gError 190406
    
    lErro = Item_Exclui_Com_Destino(iTipoDestino, lNumIntDocDestino)
    If lErro <> SUCESSO Then gError 192065

    TRVAportePagtoDireto_Exclui = SUCESSO

    Exit Function

Erro_TRVAportePagtoDireto_Exclui:

    TRVAportePagtoDireto_Exclui = gErr

    Select Case gErr

        Case 190403, 190404
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVAPORTEPAGTODIRETO", gErr)

        Case 190405
            Call Rotina_Erro(vbOKOnly, "ERRO_TRVAPORTEPAGTODIRETO_NAO_CADASTRADO", gErr, objTRVAportePagtoDireto.lNumIntDoc)
        
        Case 190406
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRVAPORTEPAGTODIRETO", gErr)
            
        Case 192065

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190407)

    End Select

    Exit Function

End Function

Public Function TRVAportePagtoDireto_Atualiza(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal objTRVAportePagtoDireto As ClassTRVAportePagtoDireto) As Long

Dim lErro As Long
Dim lNumIntDoc As Long

On Error GoTo Erro_TRVAportePagtoDireto_Atualiza

    'Le a tabelaTRVAportePagtoDireto
    lErro = Comando_ExecutarPos(lComando1, "SELECT NumIntDoc FROM TRVAportePagtoDireto WHERE NumIntDoc= ? ", _
                0, lNumIntDoc, objTRVAportePagtoDireto.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 190398

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190399

    'Se existir => UPDATE, senão => INSERT
    If lErro <> AD_SQL_SUCESSO Then gError 190400

    lErro = Comando_ExecutarPos(lComando2, "UPDATE TRVAportePagtoDireto SET Valor= ? , Vencimento= ? , NumIntDocDestino= ? , FormaPagto= ? , " & _
            "TipoDocDestino= ?, Seq = ?  ", lComando1, _
            objTRVAportePagtoDireto.dValor, objTRVAportePagtoDireto.dtVencimento, objTRVAportePagtoDireto.lNumIntDocDestino, objTRVAportePagtoDireto.iFormaPagto, _
            objTRVAportePagtoDireto.iTipoDocDestino, objTRVAportePagtoDireto.iSeq)
    If lErro <> AD_SQL_SUCESSO Then gError 190401

    TRVAportePagtoDireto_Atualiza = SUCESSO

    Exit Function

Erro_TRVAportePagtoDireto_Atualiza:

    TRVAportePagtoDireto_Atualiza = gErr

    Select Case gErr

        Case 190398, 190399
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVAPORTEPAGTODIRETO", gErr)

        Case 190400
            Call Rotina_Erro(vbOKOnly, "ERRO_TRVAPORTEPAGTODIRETO_NAO_CADASTRADO", gErr, objTRVAportePagtoDireto.lNumIntDoc)
        
        Case 190401
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRVAPORTEPAGTODIRETO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190402)

    End Select

    Exit Function

End Function

Public Function TRVAportePagtoDireto_Insere(ByVal lComando As Long, ByVal objTRVAportePagtoDireto As ClassTRVAportePagtoDireto, ByVal objAporte As ClassTRVAportes) As Long

Dim lErro As Long
Dim lNumIntDoc As Long
Dim objDebitoRecCli As New ClassDebitoRecCli
Dim colInfoComissao As New colInfoComissao
Dim objContabil As ClassContabil
Dim lCodigoAuto As Long
Dim sFiltro As String

On Error GoTo Erro_TRVAportePagtoDireto_Insere

    'Obter NumIntDoc
    lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_INT_PROX_TRVAPORTEPAGTODIRETO", lNumIntDoc)
    If lErro <> SUCESSO Then gError 190395

    objTRVAportePagtoDireto.lNumIntDoc = lNumIntDoc
    
    If objTRVAportePagtoDireto.iFormaPagto = FORMAPAGTO_TRV_OCR_FAT Then
        'Não precisa fazer nada => O Tratamento é dado quando chegar a nota
    Else
        'Gera o crédito e informa o número do numintdoc
          
        lErro = CF("Config_ObterAutomatico2", TRV_TIPO_DOC_DESTINO_DEBCLI_TABELA, "NumTitulo", lCodigoAuto, sFiltro)
        If lErro <> SUCESSO Then gError 190751
           
        objDebitoRecCli.iFilialEmpresa = giFilialEmpresa
        objDebitoRecCli.lCliente = objAporte.lCliente
        objDebitoRecCli.iFilial = objAporte.iFilialCliente
        objDebitoRecCli.sSiglaDocumento = "NCR" 'Nota de crédito
        objDebitoRecCli.lNumTitulo = lCodigoAuto
        objDebitoRecCli.dtDataEmissao = gdtDataAtual
        objDebitoRecCli.dValorTotal = objTRVAportePagtoDireto.dValor
        objDebitoRecCli.dValorProdutos = objTRVAportePagtoDireto.dValor
        
        lErro = CF("DebitoRecCli_Grava_EmTrans", objDebitoRecCli, colInfoComissao, objContabil)
        If lErro <> SUCESSO Then gError 190752
        
        objTRVAportePagtoDireto.lNumIntDocDestino = objDebitoRecCli.lNumIntDoc
        objTRVAportePagtoDireto.iTipoDocDestino = TRV_TIPO_DOC_DESTINO_DEBCLI
    
    End If

    lErro = Comando_Executar(lComando, "INSERT INTO TRVAportePagtoDireto( NumIntDoc, NumIntDocAporte, Valor, Vencimento, " & _
            "NumIntDocDestino, FormaPagto, TipoDocDestino, Seq) VALUES (?,?,?,?,?,?,?,?)", _
            objTRVAportePagtoDireto.lNumIntDoc, objTRVAportePagtoDireto.lNumIntDocAporte, objTRVAportePagtoDireto.dValor, objTRVAportePagtoDireto.dtVencimento, _
            objTRVAportePagtoDireto.lNumIntDocDestino, objTRVAportePagtoDireto.iFormaPagto, objTRVAportePagtoDireto.iTipoDocDestino, objTRVAportePagtoDireto.iSeq)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190396

    TRVAportePagtoDireto_Insere = SUCESSO

    Exit Function

Erro_TRVAportePagtoDireto_Insere:

    TRVAportePagtoDireto_Insere = gErr

    Select Case gErr

        Case 190395, 190751, 190752

        Case 190396
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRVAPORTEPAGTODIRETO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190397)

    End Select


    Exit Function

End Function

Public Function TRVAportePagtoFat_Exclui(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal objTRVAportePagtoFat As ClassTRVAportePagtoFat) As Long

Dim lErro As Long
Dim lNumIntDoc As Long

On Error GoTo Erro_TRVAportePagtoFat_Exclui

    'Le a tabelaTRVAportePagtoFat
    lErro = Comando_ExecutarPos(lComando1, "SELECT NumIntDoc FROM TRVAportePagtoFat WHERE NumIntDoc= ? ", _
                0, lNumIntDoc, objTRVAportePagtoFat.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 190416

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190417

    'Se não existir => ERRO
    If lErro <> AD_SQL_SUCESSO Then gError 190418

    lErro = Comando_ExecutarPos(lComando2, "DELETE FROM TRVAportePagtoFat ", lComando1)
    If lErro <> AD_SQL_SUCESSO Then gError 190419
    
    TRVAportePagtoFat_Exclui = SUCESSO

    Exit Function

Erro_TRVAportePagtoFat_Exclui:

    TRVAportePagtoFat_Exclui = gErr

    Select Case gErr

        Case 190416, 190417
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVAPORTEPAGTOFAT", gErr)

        Case 190418
            Call Rotina_Erro(vbOKOnly, "ERRO_TRVAPORTEPAGTOFAT_NAO_CADASTRADO", gErr, objTRVAportePagtoFat.lNumIntDoc)
        
        Case 190419
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRVAPORTEPAGTOFAT", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190420)

    End Select

    Exit Function

End Function

Public Function TRVAportePagtoFat_Atualiza(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal objTRVAportePagtoFat As ClassTRVAportePagtoFat) As Long

Dim lErro As Long
Dim lNumIntDoc As Long

On Error GoTo Erro_TRVAportePagtoFat_Atualiza

    'Le a tabelaTRVAportePagtoFat
    lErro = Comando_ExecutarPos(lComando1, "SELECT NumIntDoc FROM TRVAportePagtoFat WHERE NumIntDoc= ? ", _
                0, lNumIntDoc, objTRVAportePagtoFat.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 190411

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190412

    'Se não existir => ERRO
    If lErro <> AD_SQL_SUCESSO Then gError 190413

    lErro = Comando_ExecutarPos(lComando2, "UPDATE TRVAportePagtoFat SET Valor= ? , ValidadeDe= ? , ValidadeAte= ? , Saldo= ? , " & _
            "Percentual= ?, Seq = ? ", lComando1, _
            objTRVAportePagtoFat.dValor, objTRVAportePagtoFat.dtValidadeDe, objTRVAportePagtoFat.dtValidadeAte, objTRVAportePagtoFat.dSaldo, _
            objTRVAportePagtoFat.dPercentual, objTRVAportePagtoFat.iSeq)
    If lErro <> AD_SQL_SUCESSO Then gError 190414
    
    TRVAportePagtoFat_Atualiza = SUCESSO

    Exit Function

Erro_TRVAportePagtoFat_Atualiza:

    TRVAportePagtoFat_Atualiza = gErr

    Select Case gErr

        Case 190411, 190412
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVAPORTEPAGTOFAT", gErr)

        Case 190413
            Call Rotina_Erro(vbOKOnly, "ERRO_TRVAPORTEPAGTOFAT_NAO_CADASTRADO", gErr, objTRVAportePagtoFat.lNumIntDoc)
        
        Case 190414
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRVAPORTEPAGTOFAT", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190415)

    End Select

    Exit Function

End Function

Public Function TRVAportePagtoFat_Insere(ByVal lComando As Long, ByVal objTRVAportePagtoFat As ClassTRVAportePagtoFat, ByVal objAporte As ClassTRVAportes) As Long

Dim lErro As Long
Dim lNumIntDoc As Long

On Error GoTo Erro_TRVAportePagtoFat_Insere

    'Obter NumIntDoc
    lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_INT_PROX_TRVAPORTEPAGTOFAT", lNumIntDoc)
    If lErro <> SUCESSO Then gError 190408

    objTRVAportePagtoFat.lNumIntDoc = lNumIntDoc

    lErro = Comando_Executar(lComando, "INSERT INTO TRVAportePagtoFat( NumIntDoc, NumIntDocAporte, Valor, ValidadeDe, " & _
            "ValidadeAte, Saldo, Percentual, Seq) VALUES (?,?,?,?,?,?,?,?)", _
            objTRVAportePagtoFat.lNumIntDoc, objTRVAportePagtoFat.lNumIntDocAporte, objTRVAportePagtoFat.dValor, objTRVAportePagtoFat.dtValidadeDe, _
            objTRVAportePagtoFat.dtValidadeAte, objTRVAportePagtoFat.dSaldo, objTRVAportePagtoFat.dPercentual, objTRVAportePagtoFat.iSeq)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190409

    TRVAportePagtoFat_Insere = SUCESSO

    Exit Function

Erro_TRVAportePagtoFat_Insere:

    TRVAportePagtoFat_Insere = gErr

    Select Case gErr

        Case 190408

        Case 190409
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRVAPORTEPAGTOFAT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190410)

    End Select

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'EXCLUSÃO
Public Function TRVAportes_Exclui(ByVal objTRVAportes As ClassTRVAportes) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_TRVAportes_Exclui

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 190421

    lErro = CF("TRVAportes_Exclui_EmTrans", objTRVAportes)
    If lErro <> SUCESSO Then gError 190422

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 190423

    TRVAportes_Exclui = SUCESSO

    Exit Function

Erro_TRVAportes_Exclui:

    TRVAportes_Exclui = gErr

    Select Case gErr

        Case 190421
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 190422

        Case 190423
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190424)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function TRVAportes_Exclui_EmTrans(ByVal objTRVAportes As ClassTRVAportes) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long

On Error GoTo Erro_TRVAportes_Exclui_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 190425
    Next

    'Le a tabelaTRVAportes
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM TRVAportes WHERE Codigo= ? ", _
                0, lNumIntDoc, objTRVAportes.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 190426

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190427

    'Se não existir => ERRO
    If lErro = AD_SQL_SEM_DADOS Then gError 190428
    
    objTRVAportes.lNumIntDoc = lNumIntDoc

    'Trava registro contra alterações/Leituras
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> SUCESSO Then gError 190429

    lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM TRVAportes", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 190430
    
    lErro = CF("TRVAportePagtoCond_Exclui_Aporte", objTRVAportes)
    If lErro <> SUCESSO Then gError 190431

    lErro = CF("TRVAportePagtoFat_Exclui_Aporte", objTRVAportes)
    If lErro <> SUCESSO Then gError 190432

    lErro = CF("TRVAportePagtoDireto_Exclui_Aporte", objTRVAportes)
    If lErro <> SUCESSO Then gError 190433

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TRVAportes_Exclui_EmTrans = SUCESSO

    Exit Function

Erro_TRVAportes_Exclui_EmTrans:

    TRVAportes_Exclui_EmTrans = gErr

    Select Case gErr

        Case 190425
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 190426, 190427
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVAPORTES", gErr)

        Case 190428
            Call Rotina_Erro(vbOKOnly, "ERRO_TRVAPORTES_NAO_CADASTRADO", gErr, objTRVAportes.lCodigo)

        Case 190429
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRVAPORTES", gErr)

        Case 190430
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRVAPORTES", gErr)
            
        Case 190431 To 190433

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190434)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function TRVAportePagtoCond_Exclui_Aporte(ByVal objAporte As ClassTRVAportes) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Integer
Dim iTipoDestino As Integer
Dim lNumIntDocDestino As Long

On Error GoTo Erro_TRVAportePagtoCond_Exclui_Aporte

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 190435
    Next

    'Le a tabelaTRVAportePagtoCond
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc, TipoDocDestino, NumIntDocDestino FROM TRVAportePagtoCond WHERE NumIntDocAporte = ? ", _
                0, lNumIntDoc, iTipoDestino, lNumIntDocDestino, objAporte.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 190436

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190437

    Do While lErro <> AD_SQL_SEM_DADOS

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 190438
    
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM TRVAportePagtoCond", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 190439
        
        lErro = Item_Exclui_Com_Destino(iTipoDestino, lNumIntDocDestino)
        If lErro <> SUCESSO Then gError 192064

        'Busca Proximo
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190440

    Loop

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TRVAportePagtoCond_Exclui_Aporte = SUCESSO

    Exit Function

Erro_TRVAportePagtoCond_Exclui_Aporte:

    TRVAportePagtoCond_Exclui_Aporte = gErr

    Select Case gErr

        Case 190435
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 190436, 190437, 190440
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVAPORTEPAGTOCOND", gErr)

        Case 190438
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRVAPORTEPAGTOCOND", gErr)

        Case 190439
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRVAPORTEPAGTOCOND", gErr)
            
        Case 192064

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190441)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function TRVAportePagtoDireto_Exclui_Aporte(ByVal objAporte As ClassTRVAportes) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Integer
Dim iTipoDestino As Integer
Dim lNumIntDocDestino As Long

On Error GoTo Erro_TRVAportePagtoDireto_Exclui_Aporte

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 190442
    Next

    'Le a tabelaTRVAportePagtoDireto
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc, TipoDocDestino, NumIntDocDestino FROM TRVAportePagtoDireto WHERE NumIntDocAporte = ? ", _
                0, lNumIntDoc, iTipoDestino, lNumIntDocDestino, objAporte.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 190443

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190444

    Do While lErro <> AD_SQL_SEM_DADOS

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 190445
    
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM TRVAportePagtoDireto", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 190446
        
        lErro = Item_Exclui_Com_Destino(iTipoDestino, lNumIntDocDestino)
        If lErro <> SUCESSO Then gError 192066

        'Busca Proximo
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190447

    Loop

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TRVAportePagtoDireto_Exclui_Aporte = SUCESSO

    Exit Function

Erro_TRVAportePagtoDireto_Exclui_Aporte:

    TRVAportePagtoDireto_Exclui_Aporte = gErr

    Select Case gErr

        Case 190442
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 190443, 190444, 190447
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVAPORTEPAGTODIRETO", gErr)

        Case 190445
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRVAPORTEPAGTODIRETO", gErr)

        Case 190446
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRVAPORTEPAGTODIRETO", gErr)
            
        Case 192066

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190448)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function TRVAportePagtoFat_Exclui_Aporte(ByVal objAporte As ClassTRVAportes) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Integer

On Error GoTo Erro_TRVAportePagtoFat_Exclui_Aporte

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 190449
    Next

    'Le a tabelaTRVAportePagtoFat
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM TRVAportePagtoFat WHERE NumIntDocAporte = ? ", _
                0, lNumIntDoc, objAporte.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 190450

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190451

    Do While lErro <> AD_SQL_SEM_DADOS

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 190452
    
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM TRVAportePagtoFat", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 190453

        'Busca Proximo
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190454

    Loop

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TRVAportePagtoFat_Exclui_Aporte = SUCESSO

    Exit Function

Erro_TRVAportePagtoFat_Exclui_Aporte:

    TRVAportePagtoFat_Exclui_Aporte = gErr

    Select Case gErr

        Case 190449
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 190450, 190451, 190454
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVAPORTEPAGTOFAT", gErr)

        Case 190452
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRVAPORTEPAGTOFAT", gErr)

        Case 190453
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRVAPORTEPAGTOFAT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190455)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function RelOpAcomInadTRV_Prepara(ByVal objRelAcomInadTRV As ClassRelAcomInadTRV, lNumIntRel As Long) As Long
 
Dim lErro As Long
Dim lTransacao As Long
Dim alComando(0 To 5) As Long
Dim iIndice As Integer
Dim iIdxData As Integer
Dim iIdxFilial As Integer
Dim dValorVencimentos As Double
Dim dValorAbertoData As Double
Dim dValorBaixadoData As Double
Dim dValorAtrasadoData As Double
Dim dValorAux As Double
Dim dValorBaixaAux As Double
Dim dValorTBaixaAux As Double
Dim dValorTAux As Double
Dim lNumIntParcela As Long
 
On Error GoTo Erro_RelOpAcomInadTRV_Prepara
 
    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 190533
 
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 190534
    Next
    
    'Obter NumIntRel
    lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_INT_REL_INADTRV", lNumIntRel)
    If lErro <> SUCESSO Then gError 190535
    
    For iIdxFilial = 1 To objRelAcomInadTRV.iNumFiliais
    
        If objRelAcomInadTRV.aiFiliais(iIdxFilial) <> 0 Then
    
            lErro = Comando_Executar(alComando(0), "SELECT SUM(P.Valor)FROM ParcelasRecTodas AS P, TitulosRecTodos as T WHERE P.NumIntTitulo = T.NumIntDoc AND P.Status <> ? AND T.FilialEmpresa = ? AND (T.SiglaDocumento = ? OR ? = ?) AND (P.DataVencimentoReal >= ? OR ? = ?) AND (P.DataVencimentoReal <= ? OR ? = ?) ", _
            dValorVencimentos, STATUS_EXCLUIDO, objRelAcomInadTRV.aiFiliais(iIdxFilial), objRelAcomInadTRV.sTipoDoc, objRelAcomInadTRV.sTipoDoc, "", objRelAcomInadTRV.dtDataVencDe, objRelAcomInadTRV.dtDataVencDe, DATA_NULA, objRelAcomInadTRV.dtDataVencAte, objRelAcomInadTRV.dtDataVencAte, DATA_NULA)
            If lErro <> AD_SQL_SUCESSO Then gError 190536
        
            'Busca Primeiro
            lErro = Comando_BuscarPrimeiro(alComando(0))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190537
        
            lErro = Comando_Executar(alComando(1), "INSERT INTO RelInadTRVCab (NumIntRel, FilialEmpresa, ValorVencimentos,DataVencDe,DataVencAte,TipoDoc) VALUES (?,?,?,?,?,?) ", lNumIntRel, objRelAcomInadTRV.aiFiliais(iIdxFilial), dValorVencimentos, objRelAcomInadTRV.dtDataVencDe, objRelAcomInadTRV.dtDataVencAte, objRelAcomInadTRV.sTipoDoc)
            If lErro <> AD_SQL_SUCESSO Then gError 190538
        
            For iIdxData = 1 To objRelAcomInadTRV.iNumDatas
            
                dValorTAux = 0
                dValorTBaixaAux = 0
            
                lErro = Comando_Executar(alComando(2), "SELECT P.NumIntDoc, P.Valor FROM ParcelasRecTodas AS P, TitulosRecTodos as T WHERE P.NumIntTitulo = T.NumIntDoc AND P.Status <> ? AND T.FilialEmpresa = ? AND (T.SiglaDocumento = ? OR ? = ?) AND (P.DataVencimentoReal >= ? OR ? = ?) AND (P.DataVencimentoReal <= ? OR ? = ?) AND T.DataEmissao < ? AND P.DataVencimentoReal < ? ", _
                lNumIntParcela, dValorAux, STATUS_EXCLUIDO, objRelAcomInadTRV.aiFiliais(iIdxFilial), objRelAcomInadTRV.sTipoDoc, objRelAcomInadTRV.sTipoDoc, "", objRelAcomInadTRV.dtDataVencDe, objRelAcomInadTRV.dtDataVencDe, DATA_NULA, objRelAcomInadTRV.dtDataVencAte, objRelAcomInadTRV.dtDataVencAte, DATA_NULA, objRelAcomInadTRV.adtDatasRef(iIdxData), objRelAcomInadTRV.adtDatasRef(iIdxData))
                If lErro <> AD_SQL_SUCESSO Then gError 190539

                'Busca Primeiro
                lErro = Comando_BuscarPrimeiro(alComando(2))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190540
                
                Do While lErro <> AD_SQL_SEM_DADOS
                
                    'Total de títulos que estavam vencidos na época
                    dValorTAux = dValorTAux + dValorAux

                    'Le as baixas para tirar os que já estavam pagos (na época)
                    lErro = Comando_Executar(alComando(3), "SELECT SUM(BP.ValorBaixado)FROM BaixasParcRec AS BP, BaixasRec AS B WHERE B.NumIntBaixa = BP.NumIntBaixa AND BP.Status <> ? AND BP.NumIntParcela = ? AND B.Data < ? ", dValorBaixaAux, STATUS_EXCLUIDO, lNumIntParcela, objRelAcomInadTRV.adtDatasRef(iIdxData))
                    If lErro <> AD_SQL_SUCESSO Then gError 190541
    
                    'Busca Primeiro
                    lErro = Comando_BuscarPrimeiro(alComando(3))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190542
                    
                    dValorTBaixaAux = dValorTBaixaAux + dValorBaixaAux
                    
                    'Busca Próximo
                    lErro = Comando_BuscarProximo(alComando(2))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190540
                    
                Loop
                
                dValorAtrasadoData = dValorTAux - dValorTBaixaAux
                
                lErro = Comando_Executar(alComando(4), "SELECT SUM(BP.ValorBaixado)FROM ParcelasRecTodas AS P, TitulosRecTodos as T , BaixasRec as B, BaixasParcRec as BP WHERE P.NumIntTitulo = T.NumIntDoc AND BP.NumIntParcela = P.NumIntDoc AND B.NumIntBaixa = BP.NUmIntBaixa AND P.Status <> ? AND Bp.Status <> ? AND T.FilialEmpresa = ? AND (T.SiglaDocumento = ? OR ? = ?) AND (P.DataVencimentoReal >= ? OR ? = ?) AND (P.DataVencimentoReal <= ? OR ? = ?) AND T.DataEmissao < ? AND B.Data < ? ", _
                dValorBaixadoData, STATUS_EXCLUIDO, STATUS_EXCLUIDO, objRelAcomInadTRV.aiFiliais(iIdxFilial), objRelAcomInadTRV.sTipoDoc, objRelAcomInadTRV.sTipoDoc, "", objRelAcomInadTRV.dtDataVencDe, objRelAcomInadTRV.dtDataVencDe, DATA_NULA, objRelAcomInadTRV.dtDataVencAte, objRelAcomInadTRV.dtDataVencAte, DATA_NULA, objRelAcomInadTRV.adtDatasRef(iIdxData), objRelAcomInadTRV.adtDatasRef(iIdxData))
                If lErro <> AD_SQL_SUCESSO Then gError 190543
                           
                'Busca Primeiro
                lErro = Comando_BuscarPrimeiro(alComando(4))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190544

                dValorAbertoData = dValorVencimentos - dValorBaixadoData

                lErro = Comando_Executar(alComando(5), "INSERT INTO RelInadTRVDet (NumIntRel, FilialEmpresa, DataReferencia, ValorAbertoData, ValorBaixadoData, ValorAtrasadoData) VALUES (?,?,?,?,?,?) ", lNumIntRel, objRelAcomInadTRV.aiFiliais(iIdxFilial), objRelAcomInadTRV.adtDatasRef(iIdxData), dValorAbertoData, dValorBaixadoData, dValorAtrasadoData)
                If lErro <> AD_SQL_SUCESSO Then gError 190545
            
            Next
        
        End If
    
    Next

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 190546
 
    RelOpAcomInadTRV_Prepara = SUCESSO
 
    Exit Function
 
Erro_RelOpAcomInadTRV_Prepara:
 
    RelOpAcomInadTRV_Prepara = gErr
 
    Select Case gErr
 
        Case 190533
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
 
        Case 190534
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
 
        Case 190535
        
        Case 190536, 190537, 190539 To 190544
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PARCELASREC", gErr)
            
        Case 190538
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RELINDCAB", gErr)
 
        Case 190545
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RELINDDET", gErr)
 
        Case 190546
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190547)
 
    End Select
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Call Transacao_Rollback
 
    Exit Function
 
End Function

Public Function TRVAportes_Libera(ByVal colAportePagtos As Collection) As Long

Dim lErro As Long
Dim lTransacao As Long
Dim objTRVAportePagtos As ClassTRVAportePagtoCond
Dim alComando(0 To 4) As Long
Dim iIndice As Integer

On Error GoTo Erro_TRVAportes_Libera

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 190746
    
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 190747
    Next
    
    For Each objTRVAportePagtos In colAportePagtos

        lErro = TRVAportes_Libera_EmTrans(alComando, objTRVAportePagtos)
        If lErro <> SUCESSO Then gError 190748
        
    Next
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 190749

    TRVAportes_Libera = SUCESSO

    Exit Function

Erro_TRVAportes_Libera:

    TRVAportes_Libera = gErr

    Select Case gErr

        Case 190746
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 190747
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 190748

        Case 190749
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190750)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function TRVAportes_Libera_EmTrans(alComando() As Long, ByVal objTRVAportePagto As ClassTRVAportePagtoCond) As Long

Dim lErro As Long
Dim lNumIntDoc As Long
Dim objDebitoRecCli As New ClassDebitoRecCli
Dim colInfoComissao As New colInfoComissao
Dim objContabil As ClassContabil
Dim objAporte As New ClassTRVAportes
Dim lCodigoAuto As Long
Dim sFiltro As String
Dim objTRVAportePagtoBD As New ClassTRVAportePagtoCond

On Error GoTo Erro_TRVAportes_Libera_EmTrans

    objTRVAportePagto.iStatus = STATUS_TRV_OCR_LIBERADO

    objAporte.lNumIntDoc = objTRVAportePagto.lNumIntDocAporte
    
    lErro = CF("TRVAportes_Le", objAporte)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 190753

    If objTRVAportePagto.iFormaPagto = FORMAPAGTO_TRV_OCR_FAT Then
        'Não precisa fazer nada => O Tratamento é dado quando chegar a nota
    Else
        'Gera o crédito e informa o número do numintdoc do débito
          
        lErro = CF("Config_ObterAutomatico2", TRV_TIPO_DOC_DESTINO_DEBCLI_TABELA, "NumTitulo", lCodigoAuto, sFiltro)
        If lErro <> SUCESSO Then gError 190754
           
        objDebitoRecCli.iFilialEmpresa = giFilialEmpresa
        objDebitoRecCli.lCliente = objAporte.lCliente
        objDebitoRecCli.iFilial = objAporte.iFilialCliente
        objDebitoRecCli.sSiglaDocumento = "NCR" 'Nota de crédito
        objDebitoRecCli.lNumTitulo = lCodigoAuto
        objDebitoRecCli.dtDataEmissao = gdtDataAtual
        objDebitoRecCli.dValorTotal = objTRVAportePagto.dValor
        objDebitoRecCli.dValorProdutos = objTRVAportePagto.dValor
        
        lErro = CF("DebitoRecCli_Grava_EmTrans", objDebitoRecCli, colInfoComissao, objContabil)
        If lErro <> SUCESSO Then gError 190755
        
        objTRVAportePagto.lNumIntDocDestino = objDebitoRecCli.lNumIntDoc
        objTRVAportePagto.iTipoDocDestino = TRV_TIPO_DOC_DESTINO_DEBCLI
    
    End If
    
    objTRVAportePagtoBD.lNumIntDoc = objTRVAportePagto.lNumIntDoc
    
    lErro = CF("TRVAportePagtoCond_Le_NumIntDoc", alComando(0), objTRVAportePagtoBD, True)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 190756

    'Se não existir => erro
    If lErro <> SUCESSO Then gError 190318

    'Altera o Status da Aporte e a forma de pagamento
    lErro = Comando_ExecutarPos(alComando(1), "UPDATE TRVAportePagtoCond SET Status= ? , FormaPagto= ?, NumIntDocDestino = ?, DataLiberacao = ? , Usuario = ?, TipoDocDestino = ?, Valor = ? ", alComando(0), objTRVAportePagto.iStatus, objTRVAportePagto.iFormaPagto, objTRVAportePagto.lNumIntDocDestino, Date, gsUsuario, objTRVAportePagto.iTipoDocDestino, objTRVAportePagto.dValor)
    If lErro <> AD_SQL_SUCESSO Then gError 190759

    TRVAportes_Libera_EmTrans = SUCESSO

    Exit Function

Erro_TRVAportes_Libera_EmTrans:

    TRVAportes_Libera_EmTrans = gErr

    Select Case gErr
    
        Case 190753 To 190756
         
        Case 190758
            Call Rotina_Erro(vbOKOnly, "ERRO_TRVAPORTEPAGTOCOND_NAO_CADASTRADO", gErr, objTRVAportePagto.lNumIntDoc)

        Case 190759
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRVAPORTEPAGTOCOND", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190760)

    End Select

    Exit Function

End Function

Public Function TRVVouchers_Grava(ByVal objTRVVouchers As ClassTRVVouchers, Optional ByVal iExtraiInfoSigav As Integer = 0, Optional ByVal bExtraiuInfoSigavNovo As Boolean = False) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lNumFat As Long
Dim dValor As Double
Dim iCartao As Integer
Dim objVoucherInfo As New ClassTRVVoucherInfo
Dim dValorBruto As Double
Dim dPercAg As Double
Dim iStatus As Integer

On Error GoTo Erro_TRVVouchers_Grava

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 190857
    Next

    'Le a tabelaTRVVouchers
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumFatCoinfo, Valor, ValorBruto, Cartao, ComissaoAg, Status FROM TRVVouchers WHERE TipVou= ? AND TipoDoc = ? AND NumVou = ? AND Serie = ?", 0, lNumFat, dValor, dValorBruto, iCartao, dPercAg, iStatus, objTRVVouchers.sTipVou, objTRVVouchers.sTipoDoc, objTRVVouchers.lNumVou, objTRVVouchers.sSerie)
    If lErro <> AD_SQL_SUCESSO Then gError 190858

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190859

    'Se existir => UPDATE, senão => INSERT
    If lErro = AD_SQL_SUCESSO Then

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 190860
        
        'Se já foi faturado tem que ignorar a alteração
        'Se não foi faturado e está vindo via importação grava completo
        If lNumFat = 0 And iExtraiInfoSigav = MARCADO Then
        
            'Só extrai as informações do Sigav novamente se alterou o valor ou passou a ser título
            'de cartão
            If iExtraiInfoSigav = MARCADO Then
                If Abs(dValor - objTRVVouchers.dValor) < DELTA_VALORMONETARIO And iCartao = objTRVVouchers.iCartao Then
                    iExtraiInfoSigav = DESMARCADO
                End If
            End If
            
            If objTRVVouchers.iStatus = 0 Then
                objTRVVouchers.iStatus = iStatus
            End If

            lErro = Comando_ExecutarPos(alComando(1), "UPDATE TRVVouchers SET NumIntArq= ? , NumRegArqTexto= ? ,Cliente = ?, Recnsica = ?, Valor= ? , Data= ? , " & _
                    "CondPagto= ? , TipVou= ? , CodGrupo= ? , NumFatCoinfo= ? , " & _
                    "Cartao= ?, Pax = ?, Produto = ?, Moeda = ?, Controle = ?, CiaCart = ?,NumCCred = ?,NumAuto = ?,QuantParc = ?,DiasAntc = ?,Kit = ?, ExtraiInfoSigav = ? , Status = ? ", alComando(0), _
                    objTRVVouchers.lNumIntArq, objTRVVouchers.lNumRegArqTexto, objTRVVouchers.lCliente, objTRVVouchers.lrecnsica, objTRVVouchers.dValor, objTRVVouchers.dtData, _
                    objTRVVouchers.iCondPagto, objTRVVouchers.sTipVou, objTRVVouchers.sCodGrupo, objTRVVouchers.lNumFatCoinfo, _
                    objTRVVouchers.iCartao, objTRVVouchers.iPax, objTRVVouchers.sProduto, objTRVVouchers.sMoeda, objTRVVouchers.sControle, objTRVVouchers.sCiaCart, objTRVVouchers.sNumCCred, objTRVVouchers.lNumAuto, objTRVVouchers.iQuantParc, objTRVVouchers.idiasantc, objTRVVouchers.iKit, iExtraiInfoSigav, objTRVVouchers.iStatus)
        Else
            If iExtraiInfoSigav = DESMARCADO Then
                lErro = Comando_ExecutarPos(alComando(1), "UPDATE TRVVouchers SET Titular = ?, Cambio = ?, ValorCambio = ?, ExtraiInfoSigav = ?, PassageiroNome = ?, PassageiroSobreNome = ?, Emissor = ?, ClienteVou = ?, CliPassageiro = ? ", alComando(0), _
                        objTRVVouchers.sTitular, objTRVVouchers.dCambio, objTRVVouchers.dValorCambio, iExtraiInfoSigav, objTRVVouchers.sPassageiroNome, objTRVVouchers.sPassageiroSobreNome, objTRVVouchers.lEmissor, objTRVVouchers.lClienteVou, objTRVVouchers.lCliPassageiro)
            End If
        End If
        If lErro <> AD_SQL_SUCESSO Then gError 190861

    Else
    
        lErro = Comando_Executar(alComando(1), "INSERT INTO TRVVouchers( NumVou, TipoDoc, Serie, NumIntArq, " & _
                "NumRegArqTexto, Cliente, Recnsica, Valor, Data, " & _
                "CondPagto, TipVou, CodGrupo, NumFatCoinfo, TipoDocDestino, " & _
                "NumIntDocDestino, Cartao, Status,Pax,Produto,Moeda,Controle, CiaCart,NumCCred,NumAuto,QuantParc,DiasAntc,Kit,ExtraiInfoSigav) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                objTRVVouchers.lNumVou, objTRVVouchers.sTipoDoc, objTRVVouchers.sSerie, objTRVVouchers.lNumIntArq, _
                objTRVVouchers.lNumRegArqTexto, objTRVVouchers.lCliente, objTRVVouchers.lrecnsica, objTRVVouchers.dValor, objTRVVouchers.dtData, _
                objTRVVouchers.iCondPagto, objTRVVouchers.sTipVou, objTRVVouchers.sCodGrupo, objTRVVouchers.lNumFatCoinfo, objTRVVouchers.iTipoDocDestino, _
                objTRVVouchers.lNumIntDocDestino, objTRVVouchers.iCartao, STATUS_TRV_VOU_ABERTO, objTRVVouchers.iPax, objTRVVouchers.sProduto, objTRVVouchers.sMoeda, objTRVVouchers.sControle, objTRVVouchers.sCiaCart, objTRVVouchers.sNumCCred, objTRVVouchers.lNumAuto, objTRVVouchers.iQuantParc, objTRVVouchers.idiasantc, objTRVVouchers.iKit, iExtraiInfoSigav)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190862

    End If
    
    '########################################################
    'GERAR AS COMISSÕES OVER, CMR E CMC Bloqueadas tanto na inclusão como na alteração
    'na alteração de vouchers
    If objTRVVouchers.lNumFatCoinfo = 0 Then
        If bExtraiuInfoSigavNovo Then
            lErro = CF("TRVVoucher_Gera_Comissao", objTRVVouchers, objTRVVouchers.dValor, objTRVVouchers.dtData, "Voucher emitido", TRV_TIPO_TRATAMENTO_COMI_VOU_NOVO)
            If lErro <> SUCESSO Then gError 196314
        Else
            'Verifica se tem dados em InfoSigav, e se o valor do voucher mudou
            If Abs(dValor - objTRVVouchers.dValor) < DELTA_VALORMONETARIO Or iCartao <> objTRVVouchers.iCartao Then
                            
                objVoucherInfo.lNumVou = objTRVVouchers.lNumVou
                objVoucherInfo.sSerie = objTRVVouchers.sSerie
                objVoucherInfo.sTipo = objTRVVouchers.sTipVou
                
                lErro = CF("TRVVoucherInfoSigav_Le", objVoucherInfo)
                If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 196315
            
                'Se as informações do voucher está completa
                If lErro = SUCESSO Then
                
                    'Mudou para pagto com cartão ou deixou de ser pagto com cartão
                    If iCartao <> objTRVVouchers.iCartao Then
                        lErro = CF("TRVVoucher_Gera_Comissao", objTRVVouchers, 0, gdtDataAtual, "Alteração no tipo de pagamento do voucher. Cartão Ant.: " & IIf(iCartao = MARCADO, "SIM", "NÃO") & " , Cartão Atual: " & IIf(objTRVVouchers.iCartao = MARCADO, "SIM", "NÃO") & ".", TRV_TIPO_TRATAMENTO_COMI_VOU_TRANSF_CARTAO)
                        If lErro <> SUCESSO Then gError 196316
                    End If
                    
                    'Mudou o valor do voucher
                    If Abs(dValor - objTRVVouchers.dValor) < DELTA_VALORMONETARIO Then
                        'Mudou para pagto com cartão ou deixou de ser pagto com cartão
                        If iCartao <> objTRVVouchers.iCartao Then
                            'Se trocou tipo de pagamento tem que considerar que uma horam vem o líquido outra o bruto
                            If iCartao = MARCADO Then
                                'Se era cartão dValor é o bruto e objTRVVouchers.dValor tem que abater também a comissão
                                lErro = CF("TRVVoucher_Gera_Comissao", objTRVVouchers, (dValor * (1 - dPercAg)) - objTRVVouchers.dValor, gdtDataAtual, "Alteração no valor do voucher Valor Ant.: " & Format(dValor, "STANDARD") & " Valor Atual: " & Format(objTRVVouchers.dValor, "STANDARD") & ".", TRV_TIPO_TRATAMENTO_COMI_VOU_ALTERADO)
                            Else
                                'Se passou a ser cartão bruto = líquido, tem que considerar o bruto anterior com o valor atual
                                lErro = CF("TRVVoucher_Gera_Comissao", objTRVVouchers, dValorBruto - objTRVVouchers.dValor, gdtDataAtual, "Alteração no valor do voucher Valor Ant.: " & Format(dValor, "STANDARD") & " Valor Atual: " & Format(objTRVVouchers.dValor, "STANDARD") & ".", TRV_TIPO_TRATAMENTO_COMI_VOU_ALTERADO)
                            End If
                        Else
                            'Foi só uma altereção de valor
                            lErro = CF("TRVVoucher_Gera_Comissao", objTRVVouchers, dValor - objTRVVouchers.dValor, gdtDataAtual, "Alteração no valor do voucher Valor Ant.: " & Format(dValor, "STANDARD") & " Valor Atual: " & Format(objTRVVouchers.dValor, "STANDARD") & ".", TRV_TIPO_TRATAMENTO_COMI_VOU_ALTERADO)
                        End If
                        If lErro <> SUCESSO Then gError 196316
                    End If
                
                End If
            
            End If
        
        End If
    End If
    '########################################################

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TRVVouchers_Grava = SUCESSO

    Exit Function

Erro_TRVVouchers_Grava:

    TRVVouchers_Grava = gErr

    Select Case gErr

        Case 190857
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 190858, 190859
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVVOUCHERS", gErr)

        Case 190860
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRVVOUCHERS", gErr)

        Case 190861
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRVVOUCHERS", gErr)

        Case 190862
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRVVOUCHERS", gErr)
            
        Case 196314, 196315, 196316

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190856)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function TRVVouchers_Cancela(ByVal objTRVVouchers As ClassTRVVouchers, Optional ByVal bImportacao As Boolean = False, Optional ByVal dValorNVL As Double = 0, Optional ByVal dtDataNVL As Date = DATA_NULA, Optional ByRef lNumIntDocOCR As Long) As Long

Dim lErro As Long
Dim alComando(0 To 7) As Long
Dim iIndice As Integer
Dim lNumFat As Long
Dim iStatus As Integer
Dim lNumIntDocTitRec As Long
Dim lNumIntDocDestino As Long
Dim iTipoDocDestino As Integer
Dim sTipoDoc As String
Dim objOcorrencia As ClassTRVOcorrencias
Dim objOcorrenciaDet As ClassTRVOcorrenciaDet
Dim dValor As Double
Dim lCliente As Long
Dim iFilialEmpresa As Integer
Dim sSiglaDocumento As String
Dim lNumVou As Long
Dim sSerie As String
Dim sTipVou As String
Dim lCodigo As Long
Dim iCartao As Integer
Dim objVoucherInfo As New ClassTRVVoucherInfo
Dim lNumIntDocNVL As Long
Dim dValorOcr As Double
Dim lLeituraVou As Long
Dim bCriaNVL As Boolean
Dim bMarcarCancelado As Boolean
Dim bNVLCartao As Boolean
Dim dValorComissao As Double
Dim bComissaoCorp As Boolean
Dim lNumIntComi As Long

On Error GoTo Erro_TRVVouchers_Cancela

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 190848
    Next
    
    sTipoDoc = String(STRING_TRV_OCR_TIPODOC, 0)
    sSerie = String(STRING_TRV_OCR_SERIE, 0)
    sTipVou = String(STRING_TRV_OCR_TIPOVOU, 0)

    'Le o(s) documento(s) a ser(em) cancelado(s)
    'Se é um DOC NVL (Importação) tenta localizar o documento no cliente (CMC, CMR ou VOU) ... se
    'não localizar é um reembolso de comissão de cartão de crédito
    If objTRVVouchers.sTipoDoc = TRV_TIPODOC_NVL_TEXTO Then
        lErro = Comando_ExecutarPos(alComando(0), "SELECT NumFatCoinfo, Status, NumIntDocDestino,TipoDocDestino, TipoDoc, Valor, Cliente, NumVou, Serie, TipVou, Cartao, NumIntDocNVL, ValorOCR FROM TRVVouchers WHERE TipVou= ? AND NumVou = ? AND Serie = ? AND Cliente = ?", _
                    0, lNumFat, iStatus, lNumIntDocDestino, iTipoDocDestino, sTipoDoc, dValor, lCliente, lNumVou, sSerie, sTipVou, iCartao, lNumIntDocNVL, dValorOcr, objTRVVouchers.sTipVou, objTRVVouchers.lNumVou, objTRVVouchers.sSerie, objTRVVouchers.lCliente)
    Else
        'Se o cancelamento é manual cancela todos os documentos associados ao voucher (VOU, CMC e CMR),
        'Senão cancela o DOC importado
        If Not bImportacao Then
            lErro = Comando_ExecutarPos(alComando(0), "SELECT NumFatCoinfo, Status, NumIntDocDestino,TipoDocDestino, TipoDoc, Valor, Cliente, NumVou, Serie, TipVou, Cartao, NumIntDocNVL, ValorOCR FROM TRVVouchers WHERE TipVou= ? AND NumVou = ? AND Serie = ? ORDER BY TipoDoc DESC", _
                        0, lNumFat, iStatus, lNumIntDocDestino, iTipoDocDestino, sTipoDoc, dValor, lCliente, lNumVou, sSerie, sTipVou, iCartao, lNumIntDocNVL, dValorOcr, objTRVVouchers.sTipVou, objTRVVouchers.lNumVou, objTRVVouchers.sSerie)
        Else
            lErro = Comando_ExecutarPos(alComando(0), "SELECT NumFatCoinfo, Status, NumIntDocDestino,TipoDocDestino, TipoDoc, Valor, Cliente, NumVou, Serie, TipVou, Cartao, NumIntDocNVL, ValorOCR FROM TRVVouchers WHERE TipVou= ? AND NumVou = ? AND Serie = ? AND TipoDoc = ?", _
                        0, lNumFat, iStatus, lNumIntDocDestino, iTipoDocDestino, sTipoDoc, dValor, lCliente, lNumVou, sSerie, sTipVou, iCartao, lNumIntDocNVL, dValorOcr, objTRVVouchers.sTipVou, objTRVVouchers.lNumVou, objTRVVouchers.sSerie, objTRVVouchers.sTipoDoc)
        End If
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 190849

    'Busca Primeiro documento
    lLeituraVou = Comando_BuscarPrimeiro(alComando(0))
    If lLeituraVou <> AD_SQL_SUCESSO And lLeituraVou <> AD_SQL_SEM_DADOS Then gError 190850
    
    bNVLCartao = False
    
    If sTipoDoc = TRV_TIPODOC_VOU_TEXTO And iCartao = MARCADO Then bNVLCartao = True
    
    'Se for comissão de cartão a agência não vai encontrar o documento original
    '(tem que cancelar o voucher (NVL ao passageiro) e gerar o NVL da comissão)
    If lLeituraVou = AD_SQL_SEM_DADOS Then
       
        'Le o voucher
        lErro = Comando_ExecutarPos(alComando(0), "SELECT NumFatCoinfo, Status, NumIntDocDestino,TipoDocDestino, TipoDoc, Valor, Cliente, NumVou, Serie, TipVou, Cartao, NumIntDocNVL, ValorOCR FROM TRVVouchers WHERE TipVou= ? AND NumVou = ? AND Serie = ? AND TipoDoc = ?", _
                    0, lNumFat, iStatus, lNumIntDocDestino, iTipoDocDestino, sTipoDoc, dValor, lCliente, lNumVou, sSerie, sTipVou, iCartao, lNumIntDocNVL, dValorOcr, objTRVVouchers.sTipVou, objTRVVouchers.lNumVou, objTRVVouchers.sSerie, TRV_TIPODOC_VOU_TEXTO)
        If lErro <> AD_SQL_SUCESSO Then gError 190849
     
        'Busca Primeiro documento
        lLeituraVou = Comando_BuscarPrimeiro(alComando(0))
        If lLeituraVou <> AD_SQL_SUCESSO And lLeituraVou <> AD_SQL_SEM_DADOS Then gError 190850
        
        If iCartao = MARCADO Then bNVLCartao = True
       
    End If
       
    'Enquanto existir
    Do While lLeituraVou <> AD_SQL_SEM_DADOS
    
        'Após a versão de comissionamento não vai existir mais cancelamento de CMR e CMC, já está na alteração de comissão
        'do vouchers
        If giVersaoTRV < TRV_VERSAO_COMISSIONAMENTO Or sTipoDoc = TRV_TIPODOC_VOU_TEXTO Then
    
            bCriaNVL = False
            bMarcarCancelado = True
    
            'Trava registro contra alterações/Leituras
            lErro = Comando_LockExclusive(alComando(0))
            If lErro <> SUCESSO Then gError 190852
                       
            'Se já estiver cancelado dá erro se for cancelamento manual,
            'na importação é só ignorar
            If iStatus = STATUS_TRV_VOU_CANCELADO Then
                If Not bImportacao Then gError 192760
            End If
       
            'Se ainda não foi inativado
            If lNumIntDocNVL = 0 Then
       
                'Se for inativação de um voucher já Faturado com a fatura baixada cria o NVL
                If lNumIntDocDestino <> 0 And iTipoDocDestino = TRV_TIPO_DOC_DESTINO_TITREC And sTipoDoc = TRV_TIPODOC_VOU_TEXTO Then
                
                    sSiglaDocumento = String(STRING_SIGLA_DOCUMENTO, 0)
                
                    'Verifica se o título a receber já foi baixado
                    lErro = Comando_Executar(alComando(2), "SELECT NumIntDoc, SiglaDocumento, FilialEmpresa FROM TitulosRecBaixados WHERE NumIntDoc = ?", _
                                lNumIntDocTitRec, sSiglaDocumento, iFilialEmpresa, lNumIntDocDestino)
                    If lErro <> AD_SQL_SUCESSO Then gError 192762
                
                    'Busca Primeiro
                    lErro = Comando_BuscarPrimeiro(alComando(2))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192763
                    
                    If lErro = AD_SQL_SUCESSO Then bCriaNVL = True
                    
                End If
                
                'Se for inativação de comissão cria o NVL se for importacao de NVL ou manual,
                'só não cria para comissão em caso de exclusão, senão vai faturar só
                'o NVL e não a Comissão (O NVL vem em outro registro)
                If (sTipoDoc = TRV_TIPODOC_CMC_TEXTO Or sTipoDoc = TRV_TIPODOC_CMR_TEXTO) Then
                    If objTRVVouchers.sTipoDoc = TRV_TIPODOC_NVL_TEXTO Or Not bImportacao Then
                        bCriaNVL = True
                        bMarcarCancelado = False
                    End If
                End If
                
                'Criação do documento NVL
                If bCriaNVL Then
                
                    Set objOcorrencia = New ClassTRVOcorrencias
                    Set objOcorrenciaDet = New ClassTRVOcorrenciaDet
                    
                    If dtDataNVL = DATA_NULA Then
                        objOcorrencia.dtDataEmissao = gdtDataAtual
                    Else
                        objOcorrencia.dtDataEmissao = dtDataNVL
                    End If
                    
                    If dValorNVL = 0 Or bNVLCartao Then
                        objOcorrencia.dValorTotal = -dValor
                    Else
                        objOcorrencia.dValorTotal = dValorNVL
                    End If
                    
                    objOcorrencia.iFilialCliente = FILIAL_MATRIZ
                    objOcorrencia.iFormaPagto = FORMAPAGTO_TRV_OCR_FAT
                    objOcorrencia.iOrigem = INATIVACAO_AUTOMATICA_CODIGO
                    objOcorrencia.iStatus = STATUS_TRV_OCR_LIBERADO
                    
                    'Tem que pegar o passageiro
                    If bNVLCartao Then
                    
                        objOcorrencia.iStatus = STATUS_TRV_OCR_BLOQUEADO
                    
                        objVoucherInfo.lNumVou = lNumVou
                        objVoucherInfo.sSerie = sSerie
                        objVoucherInfo.sTipo = sTipVou
                        
                        lErro = CF("TRVVoucherInfoSigav_Le", objVoucherInfo)
                        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 192823
            
                        'VERIFICAR SE O PASSAGEIRO ESTÁ PREENCHIDO
                        If objVoucherInfo.lCliPassageiro = 0 Then
                        
                            'CADASTRAR O PASSAGEIRO E ACERTA A INFORMAÇÃO EM VOUCHERINFO
                            lErro = TRV_VoucherInfo_Cadastra_Cliente_EmTrans(objVoucherInfo)
                            If lErro <> SUCESSO Then gError 192824
                
                        End If
                        
                        lCliente = objVoucherInfo.lCliPassageiro
                    
                    End If
    
                    If Not bImportacao Or bNVLCartao Then
                        objOcorrencia.lCliente = lCliente
                    Else
                        objOcorrencia.lCliente = objTRVVouchers.lCliente
                    End If
                    
                    lErro = CF("Config_ObterAutomatico_EmTrans", "FATConfig", "NUM_PROX_TRVOCORRENCIAS", "TRVOcorrencias", "Codigo", lCodigo)
                    If lErro <> SUCESSO Then gError 192764
            
                    objOcorrencia.lCodigo = lCodigo
                    objOcorrencia.lNumVou = lNumVou
                    objOcorrencia.sHistorico = "Reembolso com base no pagto do título " & CStr(lNumFat) & " filial empresa " & CStr(iFilialEmpresa) & " e tipo de documento " & sSiglaDocumento
                    objOcorrencia.sSerie = sSerie
                    objOcorrencia.sTipoDoc = sTipVou
                    objOcorrencia.sObservacao = ""
                    
                    objOcorrenciaDet.iTipo = INATIVACAO_AUTOMATICA_TIPO_NVL_CODIGO
                    objOcorrenciaDet.iSeq = 1
                    objOcorrenciaDet.dValor = objOcorrencia.dValorTotal
                    
                    objOcorrencia.colDetalhamento.Add objOcorrenciaDet
                    
                    If bImportacao Then
                        lErro = TRVOcorrencias_Grava_EmTrans(objOcorrencia, False)
                    Else
                        lErro = TRVOcorrencias_Grava_EmTrans(objOcorrencia)
                    End If
                    If lErro <> SUCESSO Then gError 192765
                    
                    lNumIntDocOCR = objOcorrencia.lNumIntDoc
                                    
                    lErro = Comando_ExecutarPos(alComando(3), "UPDATE TRVVouchers SET NumIntDocNVL = ?", alComando(0), lNumIntDocOCR)
                    If lErro <> AD_SQL_SUCESSO Then gError 192766
    
                    'Se for inativação de cartão tem que gerar o NVL para a comissão também
                    If bNVLCartao And giVersaoTRV < TRV_VERSAO_COMISSIONAMENTO Then
                    
                        dValorComissao = 0
                        
                         'Verifica se o título a receber foi Faturado pelo Corporator
                        lErro = Comando_Executar(alComando(4), "SELECT NumIntDocNFPagComi,ValorComissao FROM TitulosRecTRV WHERE NumIntDocTitRec = ?", _
                                    lNumIntComi, dValorComissao, lNumIntDocDestino)
                        If lErro <> AD_SQL_SUCESSO Then gError 196270
                    
                        'Busca Primeiro
                        lErro = Comando_BuscarPrimeiro(alComando(4))
                        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196271
                        
                        If lErro = AD_SQL_SEM_DADOS Then
                        
                            'buscar a fatura como NFPagar aberto
                            lErro = Comando_Executar(alComando(5), "SELECT  NumIntDoc, ValorTotal FROM NFsPag WHERE FilialEmpresa = ? AND NumNotaFiscal = ? AND Status <> ?", _
                                 lNumIntComi, dValorComissao, FILIAL_MATRIZ, lNumFat, STATUS_EXCLUIDO)
                            If lErro <> AD_SQL_SUCESSO Then gError 196272
                            
                            'Busca o primeiro
                            lErro = Comando_BuscarProximo(alComando(5))
                            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196273
                        
                            If lErro <> AD_SQL_SUCESSO Then
                            
                                'buscar a fatura como titulo a Pagar baixado
                                lErro = Comando_Executar(alComando(6), "SELECT  NumIntDoc, ValorTotal FROM NFsPagBaixadas WHERE FilialEmpresa = ? AND NumNotaFiscal = ? AND Status <> ? ", _
                                     lNumIntComi, dValorComissao, FILIAL_MATRIZ, lNumFat, STATUS_EXCLUIDO)
                                If lErro <> AD_SQL_SUCESSO Then gError 196274
                                
                                'Busca o primeiro
                                lErro = Comando_BuscarProximo(alComando(6))
                                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196275
                                
                            End If
    
                        End If
                        
                        'Se achou o NFsPag
                        If lErro = AD_SQL_SUCESSO Then
                        
                            'Le as comissões associadas a esse item de fatura que não foram geradas
                            lErro = Comando_Executar(alComando(7), "SELECT C1.ValComi FROM CoinfoComissao AS C1 WHERE NumVou = ? AND Serie = ? AND TipVou = ? AND NumIntDocCorporator = ? ORDER BY C1.NumIntArq DESC, C1.NumRegArqTexto DESC", dValorComissao, lNumVou, sSerie, TRV_TIPODOC_VOU_TEXTO, lNumIntComi)
                            If lErro <> AD_SQL_SUCESSO Then gError 196278
                            
                            'Busca a primeira
                            lErro = Comando_BuscarProximo(alComando(7))
                            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196279
                        
                        End If
        
                        If dValorComissao <> 0 Or dValorNVL <> 0 Then
                            
                            Set objOcorrencia = New ClassTRVOcorrencias
                            Set objOcorrenciaDet = New ClassTRVOcorrenciaDet
                            
                            If bImportacao Then
                                objOcorrencia.dtDataEmissao = dtDataNVL
                            Else
                                objOcorrencia.dtDataEmissao = gdtDataAtual
                            End If
                            
                            If Not bComissaoCorp Then
                                objOcorrencia.dValorTotal = dValorNVL
                            Else
                                objOcorrencia.dValorTotal = dValorComissao
                            End If
                            objOcorrencia.iFilialCliente = FILIAL_MATRIZ
                            objOcorrencia.iFormaPagto = FORMAPAGTO_TRV_OCR_FAT
                            objOcorrencia.iOrigem = INATIVACAO_AUTOMATICA_CODIGO
                            objOcorrencia.iStatus = STATUS_TRV_OCR_LIBERADO
                            objOcorrencia.lCliente = objTRVVouchers.lCliente
                            
                            lErro = CF("Config_ObterAutomatico_EmTrans", "FATConfig", "NUM_PROX_TRVOCORRENCIAS", "TRVOcorrencias", "Codigo", lCodigo)
                            If lErro <> SUCESSO Then gError 196276
                    
                            objOcorrencia.lCodigo = lCodigo
                            objOcorrencia.lNumVou = lNumVou
                            objOcorrencia.sHistorico = "Reembolso de comissão de cartão de crédito número " & CStr(lNumFat) & ", Filial Empresa " & CStr(iFilialEmpresa) & "."
                            objOcorrencia.sSerie = sSerie
                            objOcorrencia.sTipoDoc = sTipVou
                            objOcorrencia.sObservacao = ""
                            
                            objOcorrenciaDet.iTipo = INATIVACAO_AUTOMATICA_TIPO_NVL_CODIGO
                            objOcorrenciaDet.iSeq = 1
                            
                            objOcorrenciaDet.dValor = objOcorrencia.dValorTotal
                            
                            objOcorrencia.colDetalhamento.Add objOcorrenciaDet
                            
                            If bImportacao Then
                                lErro = TRVOcorrencias_Grava_EmTrans(objOcorrencia, False)
                            Else
                                If bComissaoCorp Then
                                    lErro = TRVOcorrencias_Grava_EmTrans(objOcorrencia)
                                End If
                            End If
                            If lErro <> SUCESSO Then gError 196277
                            
                            lNumIntDocOCR = objOcorrencia.lNumIntDoc
                            
                        End If
                    End If
                    
                'Se não está criando o NVL e for voucher (voucher excluído) => Estorna as comissões
                Else
                    If sTipoDoc = TRV_TIPODOC_VOU_TEXTO Then
                    
                        'Cancelou o voucher sem gerar o NVL
                        lErro = CF("TRVVoucher_Gera_Comissao", objTRVVouchers, -dValor, gdtDataAtual, "Cancelamento de voucher não faturado.", TRV_TIPO_TRATAMENTO_COMI_NVL)
                        If lErro <> SUCESSO Then gError 196277
                    
                    End If
                
                End If
                
                'Se não foi forçado (não existe documento associado) e tem que marcar como cancelado
                'Marca o item como cancelado
                If bMarcarCancelado Then
                    lErro = Comando_ExecutarPos(alComando(1), "UPDATE TRVVouchers SET Status = ?, DataCanc  = ?, HoraCanc = ?, UsuarioCanc = ?", alComando(0), STATUS_TRV_VOU_CANCELADO, Date, CDbl(Time), gsUsuario)
                    If lErro <> AD_SQL_SUCESSO Then gError 190846
                End If
                
            End If
        
        End If
        
        'Busca Próximo
        lLeituraVou = Comando_BuscarProximo(alComando(0))
        If lLeituraVou <> AD_SQL_SUCESSO And lLeituraVou <> AD_SQL_SEM_DADOS Then gError 192761
        
    Loop
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TRVVouchers_Cancela = SUCESSO

    Exit Function

Erro_TRVVouchers_Cancela:

    TRVVouchers_Cancela = gErr

    Select Case gErr

        Case 190846, 192766
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRVVOUCHERS", gErr)

        Case 190848
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 190849, 190850, 192761
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVVOUCHERS", gErr)

        Case 190852
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRVVOUCHERS", gErr)

        Case 190853
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRVVOUCHERS", gErr)
            
        Case 192760
            Call Rotina_Erro(vbOKOnly, "ERRO_VOUCHER_JA_CANCELADO", gErr)
            
        Case 192762, 192763
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOSRECBAIXADOS", gErr)
            
        Case 192764, 192765, 192823, 192824
        
        Case 196270, 196271
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOSRECTRV", gErr)
        
        Case 196272, 196273
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFSPAG", gErr)
        
        Case 196274, 196275
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFSPAGBAIXADAS", gErr)
            
        Case 196276, 196277

        Case 196278, 196279
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COINFOCOMISSAO", gErr)


        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190847)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function TRVFaturamento_Gera(ByVal objFaturamento As ClassTRVFaturamento, ByVal colArquivosHtml As Collection) As Long

Dim lErro As Long
Dim lTransacao As Long
Dim alComando(0 To 35) As Long
Dim iIndice As Integer
Dim colParcelas As New Collection
Dim lFatIni As Long
Dim lFatFim As Long

On Error GoTo Erro_TRVFaturamento_Gera

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 190986
    
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 190987
    Next

    lErro = TRVFaturamento_Gera_EmTrans(alComando, objFaturamento, colParcelas, colArquivosHtml, lFatIni, lFatFim)
    If lErro <> SUCESSO Then gError 190988

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 190989
  
    Call Rotina_Aviso(vbOKOnly, "AVISO_FAT_TRV_SUCESSO", CStr(lFatIni), CStr(lFatFim))
  
    If objFaturamento.iEnviarPorEmail = MARCADO Then
        Call TRVFaturamento_Envia_PorEmail(objFaturamento, colParcelas)
    End If
    
    Call CF("TRVTitulos_Exporta_Sigav")
    
    TRVFaturamento_Gera = SUCESSO

    Exit Function

Erro_TRVFaturamento_Gera:

    TRVFaturamento_Gera = gErr

    Select Case gErr
        
        Case 190986
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 190987
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 190988, 192211

        Case 190989
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190990)

    End Select
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Private Function TRVFaturamento_Gera_EmTrans(alComando() As Long, ByVal objFaturamento As ClassTRVFaturamento, ByVal colParcelas As Collection, ByVal colArquivosHtml As Collection, lFatIni As Long, lFatFim As Long) As Long

Dim lErro As Long
Dim objTitRec As ClassTituloReceber
Dim objTitPag As ClassTituloPagar
Dim objNFsPag As ClassNFsPag
Dim objInfoFatura As ClassTRVFATInfoFatura
Dim objInfoVoucher As ClassTRVFATInfoVoucher
Dim iTipoDestino As Integer
Dim lNumIntDocDestino As Long
Dim lNumIntDocAux As Long
Dim dSaldo As Double
Dim lProxCR As Long
Dim sConteudo As String
Dim objTRVConfig As ClassTRVConfig
Dim colConfig As New Collection
Dim sArquivo As String
Dim sConteudoHtml As String
Dim iFaturaAtual As Integer
Dim objCliente As ClassCliente
Dim objFilialEmpresa As AdmFiliais
Dim objEndereco As ClassEndereco
Dim objFilialCliente As ClassFilialCliente
Dim bAchou As Boolean
Dim iLinha As Integer
Dim objParcelaRec As ClassParcelaReceber
Dim bAtualizaProxParcela As Boolean
Dim iParcela As Integer
Dim objTRVTitulosExp As ClassTRVTitulosExp
Dim sModeloCP As String
Dim objVouAporte As ClassTRVVoucherAporte
Dim objOCRAporte As ClassTRVOcorrenciaAporte
Dim objFatAporte As ClassTRVFatAporte
Dim objFaturaAporte As ClassTRVFaturaAporte
Dim colParcRecDataVencReal As New Collection
Dim objFrmAguarde As New ClassFrmAguarde
Dim objFrmAguardeTela As New FrmAguarde
Dim sItem As String

On Error GoTo Erro_TRVFaturamento_Gera_EmTrans

    lErro = CF("TRVConfig_Le", TRVCONFIG_DIRETORIO_MODELO_FAT_HTML_CARTAO, EMPRESA_TODA, sModeloCP)
    If lErro <> SUCESSO Then gError 192082

    lErro = CF("TRVConfig_Le_Lockado", alComando(6), TRVCONFIG_PROX_NUM_TITREC, EMPRESA_TODA, sConteudo)
    If lErro <> SUCESSO Then gError 192082
    
    lProxCR = StrParaLong(sConteudo)
    lFatIni = lProxCR
    
    objFrmAguarde.iTotalItens = objFaturamento.colInfoFaturas.Count
    Call objFrmAguardeTela.Inicializa_Progressao(objFrmAguarde)
    
    iFaturaAtual = 0
    For Each objInfoFatura In objFaturamento.colInfoFaturas
    
        Set objInfoFatura.colFatAporte = New Collection
    
        bAtualizaProxParcela = True
    
        iFaturaAtual = iFaturaAtual + 1
    
        If objInfoFatura.iMarcado = MARCADO Then
        
            objFaturamento.iIndiceFaturaAtual = iFaturaAtual
            
            'Tem que ler as informações de cliente, filial e endereço e preencher na coleção
            'do objFaturamento e informar o indice no objInfoFatura, isso para otimizar
            'o código de geração do arquivo html que precisará dessas informações
            'objInfoFatura.iLinhaFatCliente
            'objInfoFatura.iLinhaFatEnderecoCli
            'objInfoFatura.iLinhaFatEnderecoEmp
            'objInfoFatura.iLinhaFatFilial
            'objInfoFatura.iLinhaFatFilialCliente
            
            bAchou = False
            iLinha = 0
            For Each objCliente In objFaturamento.colFatClientes
                iLinha = iLinha + 1
                If objCliente.lCodigo = objInfoFatura.lCliente Then
                    objInfoFatura.iLinhaFatCliente = iLinha
                    objInfoFatura.iLinhaFatFilialCliente = iLinha
                    Set objFilialCliente = objFaturamento.colFatFiliaisCliente.Item(iLinha)
                    iLinha = 0
                    For Each objEndereco In objFaturamento.colFatEnderecos
                        iLinha = iLinha + 1
                        If objEndereco.lCodigo = objFilialCliente.lEndereco Then
                            objInfoFatura.iLinhaFatEnderecoCli = iLinha
                            Exit For
                        End If
                    Next
                    bAchou = True
                    Exit For
                End If
            Next
            If Not bAchou Then
            
                Set objCliente = New ClassCliente
                Set objFilialCliente = New ClassFilialCliente
                Set objEndereco = New ClassEndereco
            
                objCliente.lCodigo = objInfoFatura.lCliente
                objFilialCliente.lCodCliente = objInfoFatura.lCliente
                objFilialCliente.iCodFilial = FILIAL_MATRIZ
                
                lErro = CF("Cliente_Le", objCliente, alComando(14), alComando(15))
                If lErro <> SUCESSO And lErro <> 12293 Then gError 192197
                
                lErro = CF("FilialCliente_Le", objFilialCliente, alComando(16))
                If lErro <> SUCESSO And lErro <> 12567 Then gError 192199

                objFaturamento.colFatClientes.Add objCliente
                objFaturamento.colFatFiliaisCliente.Add objFilialCliente
                
                objInfoFatura.iLinhaFatCliente = objFaturamento.colFatClientes.Count
                objInfoFatura.iLinhaFatFilialCliente = objFaturamento.colFatFiliaisCliente.Count
                
                objEndereco.lCodigo = objFilialCliente.lEndereco
                
                lErro = CF("Endereco_Le", objEndereco, alComando(17))
                If lErro <> SUCESSO Then gError 192200
                
                objFaturamento.colFatEnderecos.Add objEndereco
                
                objInfoFatura.iLinhaFatEnderecoCli = objFaturamento.colFatEnderecos.Count
            
            End If
            
            bAchou = False
            iLinha = 0
            For Each objFilialEmpresa In objFaturamento.colFatFiliais
                iLinha = iLinha + 1
                If objFilialEmpresa.iCodFilial = objInfoFatura.iFilialEmpresa Then
                    objInfoFatura.iLinhaFatFilial = iLinha
                    iLinha = 0
                    For Each objEndereco In objFaturamento.colFatEnderecos
                        iLinha = iLinha + 1
                        If objEndereco.lCodigo = objFilialEmpresa.objEndereco.lCodigo Then
                            objInfoFatura.iLinhaFatEnderecoEmp = iLinha
                            Exit For
                        End If
                    Next
                    bAchou = True
                    Exit For
                End If
            Next
            If Not bAchou Then
            
                Set objFilialEmpresa = New AdmFiliais
            
                objFilialEmpresa.iCodFilial = objInfoFatura.iFilialEmpresa
                
                lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
                If lErro <> SUCESSO Then gError 192198

                objFaturamento.colFatFiliais.Add objFilialEmpresa
                
                objInfoFatura.iLinhaFatFilial = objFaturamento.colFatFiliais.Count
            
                objFaturamento.colFatEnderecos.Add objFilialEmpresa.objEndereco
                
                objInfoFatura.iLinhaFatEnderecoEmp = objFaturamento.colFatEnderecos.Count
            
            End If


            If objInfoFatura.dValor > 0 Then
                
                Set objTitRec = New ClassTituloReceber
                
                iTipoDestino = TRV_TIPO_DOC_DESTINO_TITREC
                
                lErro = TRVFaturamento_Gera_TitRec(alComando(21), alComando(22), alComando(23), alComando(24), alComando(25), alComando(26), alComando(27), alComando(28), alComando(29), alComando(30), objInfoFatura, objTitRec, lProxCR, colParcRecDataVencReal)
                If lErro <> SUCESSO Then gError 192001
                
                For Each objParcelaRec In objTitRec.colParcelaReceber
                    colParcelas.Add objParcelaRec
                Next
                
                objInfoFatura.lNumeroFaturaReal = objTitRec.lNumTitulo
                
                lProxCR = lProxCR + 1
                
                lNumIntDocDestino = objTitRec.lNumIntDoc
    
            Else
            
                'Ocorrências
                If objInfoFatura.iTipoDoc = TRV_TIPODOC_FAT_OCR Then
            
                    Set objTitPag = New ClassTituloPagar
                    
                    iTipoDestino = TRV_TIPO_DOC_DESTINO_TITPAG
                    
                    lErro = TRVFaturamento_Gera_TitPag(objInfoFatura, objTitPag, lProxCR, objCliente, objFaturamento.colFatEnderecos.Item(objInfoFatura.iLinhaFatEnderecoCli), colParcRecDataVencReal)
                    If lErro <> SUCESSO Then gError 192002
                    
                    objInfoFatura.lNumeroFaturaReal = objTitPag.lNumTitulo
                    
                    lProxCR = lProxCR + 1
                    
                    lNumIntDocDestino = objTitPag.lNumIntDoc
                
                'Over
                ElseIf objInfoFatura.iTipoDoc = TRV_TIPODOC_FAT_OVER Then
                   
                     Set objTitPag = New ClassTituloPagar
                    
                    iTipoDestino = TRV_TIPO_DOC_DESTINO_TITPAG
                    
                    lErro = TRVFaturamento_Gera_TitPagOver(objInfoFatura, objTitPag, lProxCR)
                    If lErro <> SUCESSO Then gError 192002
                    
                    objInfoFatura.lNumeroFaturaReal = objTitPag.lNumTitulo
                    
                    lProxCR = lProxCR + 1
                    
                    lNumIntDocDestino = objTitPag.lNumIntDoc
                   
                'Comissão
                Else
                
                    Set objNFsPag = New ClassNFsPag
                    
                    iTipoDestino = TRV_TIPO_DOC_DESTINO_NFSPAG
                    
                    lErro = TRVFaturamento_Gera_NFPag(objInfoFatura, objNFsPag, lProxCR)
                    If lErro <> SUCESSO Then gError 194202
                    
                    objInfoFatura.lNumeroFaturaReal = objNFsPag.lNumNotaFiscal
                    
                    lProxCR = lProxCR + 1
                    
                    lNumIntDocDestino = objNFsPag.lNumIntDoc
                
                End If
            
            End If
            
            Set objTRVTitulosExp = New ClassTRVTitulosExp
            
            objTRVTitulosExp.iExcluido = DESMARCADO
            objTRVTitulosExp.iExportado = DESMARCADO
            objTRVTitulosExp.iTipoDocOrigem = iTipoDestino
            objTRVTitulosExp.lNumIntDocOrigem = lNumIntDocDestino
            objTRVTitulosExp.lNumTitulo = objInfoFatura.lNumeroFaturaReal
            objTRVTitulosExp.iTemQueContabilizar = MARCADO
            objTRVTitulosExp.dValorAporte = objInfoFatura.dValorAporte
            objTRVTitulosExp.dValorAporteCred = objInfoFatura.dValorAporteCred
            objTRVTitulosExp.dValorComissao = objInfoFatura.dValorComissao
            objTRVTitulosExp.dValorCreditos = objInfoFatura.dValorCreditos
            objTRVTitulosExp.dValorDebitos = objInfoFatura.dValorDebitos
            objTRVTitulosExp.dValorDeducoes = objInfoFatura.dValorDeducoes
            objTRVTitulosExp.dValorTarifa = objInfoFatura.dValorTarifa
            
            For Each objFaturaAporte In objInfoFatura.colFatAporteCred
                objTRVTitulosExp.colFatAporteCred.Add objFaturaAporte
            Next
            
            lErro = CF("TRVTitulosExp_Insere", objTRVTitulosExp, alComando(18), alComando(19))
            If lErro <> SUCESSO Then gError 196240
            
            For Each objInfoVoucher In objInfoFatura.colInfoVouchers
            
                If objInfoVoucher.iMarcado = MARCADO Then
            
                    'Se é um voucher ou uma comissão
                    If objInfoVoucher.sTipoDoc = TRV_TIPODOC_VOU_TEXTO Or ((objInfoVoucher.sTipoDoc = TRV_TIPODOC_CMR_TEXTO Or objInfoVoucher.sTipoDoc = TRV_TIPODOC_CMC_TEXTO) And giVersaoTRV < TRV_VERSAO_COMISSIONAMENTO) Then
    
                        lErro = Comando_ExecutarPos(alComando(0), "SELECT Recnsica FROM TRVVouchers WHERE NumVou= ? AND TipoDoc = ? AND Serie = ? AND TipVou = ? ", _
                                    0, lNumIntDocAux, objInfoVoucher.lNumVou, objInfoVoucher.sTipoDoc, objInfoVoucher.sSerie, objInfoVoucher.sTipoVou)
                        If lErro <> AD_SQL_SUCESSO Then gError 192003
                    
                        'Busca Primeiro
                        lErro = Comando_BuscarPrimeiro(alComando(0))
                        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192004
                        
                        If lErro = AD_SQL_SEM_DADOS Then gError 192016
                    
                        'Trava registro contra alterações/Leituras
                        lErro = Comando_LockExclusive(alComando(0))
                        If lErro <> SUCESSO Then gError 192005
                        
                        lErro = Comando_ExecutarPos(alComando(1), "UPDATE TRVVouchers SET TipoDocDestino= ?, NumIntDocDestino=?, ValorAporte = ?, ValorComissao = ?, NumFatCoinfo = ? ", alComando(0), _
                        iTipoDestino, lNumIntDocDestino, objInfoVoucher.dValorAporte, objInfoVoucher.dValorComissao, objInfoFatura.lNumeroFaturaReal)
                        If lErro <> AD_SQL_SUCESSO Then gError 192006
                        
                        'Se é um voucher e a versão é a partir da de comissionamento verifica se
                        'não tem que liberar o OVER (liberação na emissão)
                        If objInfoVoucher.sTipoDoc = TRV_TIPODOC_VOU_TEXTO And giVersaoTRV >= TRV_VERSAO_COMISSIONAMENTO Then
                        
                            sItem = String(255, 0)
                            
                            'Verifica o tipo de tratamento (Se o pagamento é na baixa ou na emissão)
                            lErro = Comando_Executar(alComando(33), "SELECT Item FROM FilialFornecedorCategorias WHERE Fornecedor = ? AND Filial = ? AND Categoria = ? ", sItem, objInfoVoucher.lCliente, FILIAL_MATRIZ, "OVER")
                            If lErro <> AD_SQL_SUCESSO Then gError 182532
                            
                            'Busca a primeira
                            lErro = Comando_BuscarProximo(alComando(33))
                            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182533
                            
                            'Se não encontrou o tipo de tratamento então é porque é na baixa
                            If lErro <> AD_SQL_SUCESSO Then
                                sItem = "Baixa"
                            End If
            
                            'Se não for na baixa libera o over para pagamento
                            If InStr(1, UCase(sItem), "BAIXA") <> 0 Then
                            
                                'Le as comissões over bloqueadas para esse voucher
                                lErro = Comando_ExecutarPos(alComando(34), "SELECT NumIntDoc FROM TRVVoucherInfo WHERE NumVou= ? AND TipoDoc = ? AND Serie = ? AND TipVou = ? AND Status = ? ", _
                                0, lNumIntDocAux, , objInfoVoucher.lNumVou, TRV_TIPODOC_OVER_TEXTO, objInfoVoucher.sSerie, objInfoVoucher.sTipoVou, TRV_VOU_INFO_STATUS_BLOQUEADO)
                                If lErro <> AD_SQL_SUCESSO Then gError 196376
                            
                                'Busca Primeiro
                                lErro = Comando_BuscarPrimeiro(alComando(34))
                                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196377
                        
                                Do While lErro <> AD_SQL_SEM_DADOS
                                
                                    'Trava registro contra alterações/Leituras
                                    lErro = Comando_LockExclusive(alComando(34))
                                    If lErro <> SUCESSO Then gError 196378
                                    
                                    lErro = Comando_ExecutarPos(alComando(35), "UPDATE TRVVoucherInfo SET Status = ?, TipoLiberacao = ?, NumIntDocLiberacao = ? ", alComando(34), TRV_VOU_INFO_STATUS_LIBERADO, TRV_TIPO_LIBERACAO_COMISSAO_FAT, lNumIntDocDestino)
                                    If lErro <> AD_SQL_SUCESSO Then gError 196379
                                
                                    'Busca Próximo
                                    lErro = Comando_BuscarProximo(alComando(34))
                                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196380
                                
                                Loop
                                
                            End If
                            
                        End If
   
                        For Each objVouAporte In objInfoVoucher.colAportes
            
                            'Busca para ver se o aporte já teve a próxima parcela atualizada
                            bAchou = False
                            For Each objFatAporte In objInfoFatura.colFatAporte
                                If objFatAporte.lCodigo = objVouAporte.lCodigoAporte Then
                                    bAchou = True
                                    Exit For
                                End If
                            Next
                            If Not bAchou Then
                                
                                'Acerta a próxima parcela
                                lErro = TRVAportes_Acerta_ProxParcela(alComando(8), alComando(9), objVouAporte.lCodigoAporte, iParcela)
                                If lErro <> SUCESSO Then gError 192021
                                
                                Set objFatAporte = New ClassTRVFatAporte
                                
                                objFatAporte.lCodigo = objVouAporte.lCodigoAporte
                                objFatAporte.iProxParcela = iParcela
                                
                                objInfoFatura.colFatAporte.Add objFatAporte
                                
                                bAtualizaProxParcela = True
                            Else
                                bAtualizaProxParcela = False
                            End If
                            
                            objVouAporte.iParcelaAporte = objFatAporte.iProxParcela
                            objVouAporte.sSerie = objInfoVoucher.sSerie
                            objVouAporte.sTipVou = objInfoVoucher.sTipoVou
                            objVouAporte.lNumVou = objInfoVoucher.lNumVou
                                       
                            'Acerta o saldo do aporte
                            If objVouAporte.iMoedaAporte = MOEDA_REAL Then
                                lErro = TRVAportePagtoFat_Acerta_Saldo(alComando(4), alComando(5), alComando(12), alComando(13), objVouAporte.iTipoPagtoAporte, objVouAporte.lNumIntDocPagtoAporte, objVouAporte.dValorAporte)
                            Else
                                lErro = TRVAportePagtoFat_Acerta_Saldo(alComando(4), alComando(5), alComando(12), alComando(13), objVouAporte.iTipoPagtoAporte, objVouAporte.lNumIntDocPagtoAporte, objVouAporte.dValorAporteUSS)
                            End If
                            If lErro <> SUCESSO Then gError 192021
                            
                            lErro = Comando_Executar(alComando(10), "INSERT INTO TRVVoucherAporte( TipVou, Serie, NumVou, TipoPagtoAporte, " & _
                                    "NumIntDocPagtoAporte, CodigoAporte, ParcelaAporte, ValorAporte, ValorAporteUSS, MoedaAporte) VALUES (?,?,?,?,?,?,?,?,?,?)", _
                                    objVouAporte.sTipVou, objVouAporte.sSerie, objVouAporte.lNumVou, objVouAporte.iTipoPagtoAporte, _
                                    objVouAporte.lNumIntDocPagtoAporte, objVouAporte.lCodigoAporte, objVouAporte.iParcelaAporte, objVouAporte.dValorAporte, objVouAporte.dValorAporteUSS, objVouAporte.iMoedaAporte)
                            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196476
                        
                        Next
                        
                    'Quando é uma comissão faturada com a versão de comissionamento (CMC, CMR, CMCC e OVER)
                    ElseIf ((objInfoVoucher.sTipoDoc = TRV_TIPODOC_CMR_TEXTO Or objInfoVoucher.sTipoDoc = TRV_TIPODOC_CMC_TEXTO Or objInfoVoucher.sTipoDoc = TRV_TIPODOC_CMCC_TEXTO Or objInfoVoucher.sTipoDoc = TRV_TIPODOC_OVER_TEXTO) And giVersaoTRV >= TRV_VERSAO_COMISSIONAMENTO) Then
                    
                        lErro = Comando_ExecutarPos(alComando(31), "SELECT NumIntDoc FROM TRVVoucherInfo WHERE NumIntDoc = ? ", 0, lNumIntDocAux, objInfoVoucher.lNumIntDoc)
                        If lErro <> AD_SQL_SUCESSO Then gError 196371
                    
                        'Busca Primeiro
                        lErro = Comando_BuscarPrimeiro(alComando(31))
                        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196372
                        
                        If lErro = AD_SQL_SEM_DADOS Then gError 196373
                        
                        'Trava registro contra alterações/Leituras
                        lErro = Comando_LockExclusive(alComando(31))
                        If lErro <> SUCESSO Then gError 196374
                        
                        lErro = Comando_ExecutarPos(alComando(32), "UPDATE TRVVoucherInfo SET TipoDocDestino= ?, NumIntDocDestino=?, NumTitulo = ? ", alComando(31), _
                        iTipoDestino, lNumIntDocDestino, objInfoFatura.lNumeroFaturaReal)
                        If lErro <> AD_SQL_SUCESSO Then gError 196375
                    
                    'Caso seja uma ORC ou um NVL
                    Else
                    
                        lErro = Comando_ExecutarPos(alComando(2), "SELECT NumIntDoc FROM TRVOcorrencias WHERE NumIntDoc= ? ", _
                                    0, lNumIntDocAux, objInfoVoucher.lNumIntDoc)
                        If lErro <> AD_SQL_SUCESSO Then gError 192007
                    
                        'Busca Primeiro
                        lErro = Comando_BuscarPrimeiro(alComando(2))
                        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192008
                        
                        If lErro = AD_SQL_SEM_DADOS Then gError 192018
                    
                        'Trava registro contra alterações/Leituras
                        lErro = Comando_LockExclusive(alComando(2))
                        If lErro <> SUCESSO Then gError 192009
                        
                        For Each objOCRAporte In objInfoVoucher.colAportes
                
                            'Acerta o saldo do aporte
                            If objOCRAporte.iMoedaAporte = MOEDA_REAL Then
                                lErro = TRVAportePagtoFat_Acerta_Saldo(alComando(4), alComando(5), alComando(12), alComando(13), objOCRAporte.iTipoPagtoAporte, objOCRAporte.lNumIntDocPagtoAporte, objOCRAporte.dValorAporte)
                            Else
                                lErro = TRVAportePagtoFat_Acerta_Saldo(alComando(4), alComando(5), alComando(12), alComando(13), objOCRAporte.iTipoPagtoAporte, objOCRAporte.lNumIntDocPagtoAporte, objOCRAporte.dValorAporteUSS)
                            End If
                            If lErro <> SUCESSO Then gError 192021
                            
                            lErro = Comando_Executar(alComando(11), "INSERT INTO TRVOcorrenciaAporte( NumIntDocOCR, TipoPagtoAporte, " & _
                                    "NumIntDocPagtoAporte, CodigoAporte, ParcelaAporte, ValorAporte, ValorAporteUSS, MoedaAporte) VALUES (?,?,?,?,?,?,?,?)", _
                                    objOCRAporte.lNumIntDocOCR, objOCRAporte.iTipoPagtoAporte, objOCRAporte.lNumIntDocPagtoAporte, objOCRAporte.lCodigoAporte, objOCRAporte.iParcelaAporte, objOCRAporte.dValorAporte, objOCRAporte.dValorAporteUSS, objOCRAporte.iMoedaAporte)
                            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196476
                            
                        Next
                
                        lErro = Comando_ExecutarPos(alComando(3), "UPDATE TRVOcorrencias SET TipoDocDestino= ?, NumIntDocDestino=?, NumTitulo = ?, ValorAporte = ? ", alComando(2), _
                        iTipoDestino, lNumIntDocDestino, objInfoFatura.lNumeroFaturaReal, objInfoVoucher.dValorAporte)
                        If lErro <> AD_SQL_SUCESSO Then gError 192010
                        
                    End If
                
                End If
        
            Next
            
            'Atualiza o saldo dos créditos (aportes diretos que viram débitos com o cliente)
            For Each objFaturaAporte In objInfoFatura.colFatAporteCred
            
                'Grava a apropriação de saldo do DebitoReceber no BD
                lErro = CF("DebitoRecCli_AlterarSaldo", objFaturaAporte.objCredito, -objFaturaAporte.dValorAporte)
                If lErro <> SUCESSO Then gError 192194
    
            Next
            
            'Gera o documento .html
            sArquivo = objFaturamento.sDiretorio & gsEmpresaTRVHTML & String(6 - Len(CStr(objInfoFatura.lNumeroFaturaReal)), "0") & CStr(objInfoFatura.lNumeroFaturaReal) & ".html"
            
            If objInfoFatura.dValor >= 0 Then
                'Passa o modelo normal com geração de boleto
                lErro = Gera_Arquivo_Html(objFaturamento, objFaturamento.sModelo, sArquivo, sConteudoHtml)
            Else
                'Passa o modelo de cartão e nota de crédito
                lErro = Gera_Arquivo_Html(objFaturamento, sModeloCP, sArquivo, sConteudoHtml)
            End If
            If lErro <> SUCESSO Then gError 192194
            
            colArquivosHtml.Add sArquivo
            
        End If
        
        Call objFrmAguardeTela.ProcessouItem
        
        If objFrmAguarde.iCancelar = MARCADO Then Exit For

    Next
    
    Set objFrmAguardeTela = Nothing
    If objFrmAguarde.iCancelar = MARCADO Then gError 192084
    
    Set objTRVConfig = New ClassTRVConfig
    
    objTRVConfig.sCodigo = TRVCONFIG_PROX_NUM_TITREC
    objTRVConfig.sConteudo = CStr(lProxCR)
    
    lFatFim = lProxCR - 1
    
    lErro = CF("TRVConfig_Grava_Lockado", alComando(6), alComando(7), objTRVConfig)
    If lErro <> SUCESSO Then gError 192084
        
    TRVFaturamento_Gera_EmTrans = SUCESSO

    Exit Function

Erro_TRVFaturamento_Gera_EmTrans:

    TRVFaturamento_Gera_EmTrans = gErr

    Select Case gErr
    
        Case 192001, 192002, 192021, 192082, 192083, 192084, 192194, 192197 To 192200, 194202, 196240, 196476

        Case 192003, 192004
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVVOUCHERS", gErr)
        
        Case 192005
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRVVOUCHERS", gErr)
        
        Case 192006
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRVVOUCHERS", gErr)
            
        Case 192007, 192008
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVOCORRENCIAS", gErr)
        
        Case 192009
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRVOCORRENCIAS", gErr)
        
        Case 192010
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRVOCORRENCIAS", gErr)
            
        Case 192016
            Call Rotina_Erro(vbOKOnly, "ERRO_TRVVOUCHERS_NAO_CADASTRADO2", gErr, objInfoVoucher.lNumIntDoc)

        Case 192018
            Call Rotina_Erro(vbOKOnly, "ERRO_OCORRREMPARCREC_NAO_CADASTRADA", gErr, objInfoVoucher.lNumIntDoc)
            
        Case 196371, 196372, 196376, 196377, 196380
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVVOUCHERINFO", gErr)

        Case 1963713
            Call Rotina_Erro(vbOKOnly, "ERRO_TRVVOUCHERINFO_NAO_CADASTRADO", gErr, objInfoVoucher.lNumIntDoc)

        Case 196378, 196374
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRVVOUCHERINFO", gErr)
        
        Case 196379, 196375
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRVVOUCHERINFO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190991)

    End Select
    
    If Not (objFrmAguardeTela Is Nothing) Then
        Call objFrmAguardeTela.Trata_Erro
    End If

    Exit Function

End Function

Private Function TRVFaturamento_Gera_TitRec(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lComando3 As Long, ByVal lComando4 As Long, ByVal lComando5 As Long, ByVal lComando6 As Long, ByVal lComando7 As Long, ByVal lComando8 As Long, ByVal lComando9 As Long, ByVal lComando10 As Long, ByVal objInfoFatura As ClassTRVFATInfoFatura, ByVal objTitRec As ClassTituloReceber, ByVal lProxNum As Long, ByVal colParcRecDataVencReal As Collection) As Long

Dim lErro As Long
Dim objParcelaReceber As New ClassParcelaReceber
Dim objParcelaReceberAux As ClassParcelaReceber
Dim colParcelaReceber As New colParcelaReceber
Dim colComissaoEmissao As New colComissao
Dim colcolComissao As New colcolComissao
Dim colComissao As New colComissao
Dim colcolDesconto As New colcolDesconto
Dim colDesconto As New colDesconto
Dim objContabil As ClassContabil
Dim objCtb As New ClassImpCoinfoFatCRCtb
Dim objContabAutomatica As New ClassContabAutomatica
Dim dtDataVenctoReal As Date
Dim bAchou As Boolean

On Error GoTo Erro_TRVFaturamento_Gera_TitRec
    
    objTitRec.lCliente = objInfoFatura.lCliente
    objTitRec.iFilial = FILIAL_MATRIZ
    objTitRec.lNumTitulo = lProxNum
    objTitRec.dtDataEmissao = objInfoFatura.dtDataEmissao
    objTitRec.dValor = objInfoFatura.dValor
    objTitRec.dSaldo = objInfoFatura.dValor
    objTitRec.iFilialEmpresa = objInfoFatura.iFilialEmpresa

    objTitRec.sSiglaDocumento = TIPODOC_FATURA_SERVICO_CR
    objTitRec.iNumParcelas = 1

    objParcelaReceber.iNumParcela = 1
    objParcelaReceber.dtDataVencimento = objInfoFatura.dtDataVencimento
    
    bAchou = False
    For Each objParcelaReceberAux In colParcRecDataVencReal
        If objParcelaReceberAux.dtDataVencimento = objParcelaReceber.dtDataVencimento Then
            bAchou = True
            Exit For
        End If
    Next
    If Not bAchou Then

        'Calcula Data Vencimento Real
        lErro = CF("DataVencto_Real", objParcelaReceber.dtDataVencimento, dtDataVenctoReal)
        If lErro <> SUCESSO Then gError 190992
        
        Set objParcelaReceberAux = New ClassParcelaReceber
        
        objParcelaReceberAux.dtDataVencimento = objParcelaReceber.dtDataVencimento
        objParcelaReceberAux.dtDataVencimentoReal = dtDataVenctoReal
        
        colParcRecDataVencReal.Add objParcelaReceberAux
        
    End If
    
    objParcelaReceber.dtDataVencimentoReal = objParcelaReceberAux.dtDataVencimentoReal
    
    'objParcelaReceber.dtDataVencimentoReal = dtDataVenctoReal
    objParcelaReceber.dValor = objInfoFatura.dValor
    objParcelaReceber.dValorOriginal = objInfoFatura.dValor
    objParcelaReceber.dtDataCredito = DATA_NULA
    objParcelaReceber.dtDataDepositoCheque = DATA_NULA
    objParcelaReceber.dtDataEmissaoCheque = DATA_NULA
    objParcelaReceber.dtDataPrevReceb = DATA_NULA
    objParcelaReceber.dtDataProxCobr = DATA_NULA
    objParcelaReceber.dtDataTransacaoCartao = DATA_NULA
    objParcelaReceber.dtDesconto1Ate = DATA_NULA
    objParcelaReceber.dtDesconto2Ate = DATA_NULA
    objParcelaReceber.dtDesconto3Ate = DATA_NULA
    objParcelaReceber.dtValidadeCartao = DATA_NULA

    colParcelaReceber.AddObj objParcelaReceber
    
    colcolDesconto.Add colDesconto
    colcolComissao.Add colComissao
    
    'Grava o Título a receber
    'lErro = CF("TituloReceber_Grava_EmTrans", objTitRec, colComissaoEmissao, colParcelaReceber, colcolComissao, colcolDesconto, objContabil)
    
    lErro = FAT_TituloReceber_Insere_TRV(lComando1, lComando2, lComando3, lComando4, lComando5, lComando6, lComando7, lComando8, lComando9, lComando10, objTitRec, colParcelaReceber)
    If lErro <> SUCESSO Then gError 190993
                        
'    'Coloca a comissão no CSLL
'    objTitRec.dCSLLRetido = objInfoFatura.dValorComissao
'
'    'gravar a contabilizacao
'    lErro = objContabAutomatica.Inicializa_Contab(objCtb, MODULO_CONTASARECEBER, "CR", objTitRec.dtDataEmissao, objTitRec.dtDataEmissao)
'    If lErro <> SUCESSO Then gError 190994
'
'    lErro = objContabAutomatica.GeraContabilizacao(objTitRec, Nothing)
'    If lErro <> SUCESSO Then gError 190995
'
'    lErro = objContabAutomatica.Finaliza_Contab()
'    If lErro <> SUCESSO Then gError 190996
'
'    objTitRec.colParcelaReceber.AddObj objParcelaReceber
'
'    Set objContabAutomatica = Nothing
'    Call objCtb.Encerrando
'    Set objCtb = Nothing
                        
    TRVFaturamento_Gera_TitRec = SUCESSO

    Exit Function

Erro_TRVFaturamento_Gera_TitRec:

    TRVFaturamento_Gera_TitRec = gErr

    Select Case gErr
    
        Case 190992 To 190996

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190991)

    End Select

    Exit Function

End Function

Private Function TRVFaturamento_Gera_TitPag(ByVal objInfoFatura As ClassTRVFATInfoFatura, ByVal objTitPag As ClassTituloPagar, ByVal lProxNum As Long, ByVal objCliente As ClassCliente, ByVal objEndereco As ClassEndereco, ByVal colParcRecDataVencReal As Collection) As Long

Dim lErro As Long
Dim objParcelaPagar As New ClassParcelaPagar
Dim colParcelaPagar As New colParcelaPagar
Dim objContabil As ClassContabil
Dim objCtb As New ClassImpCoinfoFatCPCtb
Dim objContabAutomatica As New ClassContabAutomatica
Dim dtDataVenctoReal As Date
Dim objInfoVoucher As ClassTRVFATInfoVoucher
Dim objParcelaReceberAux As ClassParcelaReceber
Dim bAchou As Boolean

On Error GoTo Erro_TRVFaturamento_Gera_TitPag

    Set objInfoVoucher = objInfoFatura.colInfoVouchers.Item(1)

    If objInfoVoucher.iTipoCliente = TRV_TIPO_CLIENTE_PASSAGEIRO Then
    
        lErro = TRV_Cliente_Grava_Fornecedor(objCliente, objEndereco)
        If lErro <> SUCESSO Then gError 196904
    
    End If
               
    objTitPag.lFornecedor = objInfoFatura.lCliente
    objTitPag.iFilial = FILIAL_MATRIZ
    objTitPag.iFilialPedCompra = FILIAL_MATRIZ
    objTitPag.lNumTitulo = lProxNum
    objTitPag.dtDataEmissao = objInfoFatura.dtDataEmissao
    objTitPag.dValorTotal = Abs(objInfoFatura.dValor)
    objTitPag.dValorProdutos = Abs(objInfoFatura.dValor)
    objTitPag.dSaldo = Abs(objInfoFatura.dValor)
    objTitPag.iFilialEmpresa = objInfoFatura.iFilialEmpresa
    objTitPag.sSiglaDocumento = TIPODOC_FATURA_SERVICO_CP
    objTitPag.iNumParcelas = 1

    objParcelaPagar.iNumParcela = 1
    objParcelaPagar.dtDataVencimento = objInfoFatura.dtDataVencimento
    
    bAchou = False
    For Each objParcelaReceberAux In colParcRecDataVencReal
        If objParcelaReceberAux.dtDataVencimento = objParcelaPagar.dtDataVencimento Then
            bAchou = True
            Exit For
        End If
    Next
    If Not bAchou Then

        'Calcula Data Vencimento Real
        lErro = CF("DataVencto_Real", objParcelaPagar.dtDataVencimento, dtDataVenctoReal)
        If lErro <> SUCESSO Then gError 190997
        
        Set objParcelaReceberAux = New ClassParcelaReceber
        
        objParcelaReceberAux.dtDataVencimento = objParcelaPagar.dtDataVencimento
        objParcelaReceberAux.dtDataVencimentoReal = dtDataVenctoReal
        
        colParcRecDataVencReal.Add objParcelaReceberAux
        
    End If
    
    objParcelaPagar.dtDataVencimentoReal = objParcelaReceberAux.dtDataVencimentoReal

    objParcelaPagar.dValor = Abs(objInfoFatura.dValor)
    objParcelaPagar.iTipoCobranca = TIPO_COBRANCA_CARTEIRA
    objParcelaPagar.iStatus = STATUS_ABERTO

    With objParcelaPagar
        colParcelaPagar.Add .lNumIntDoc, .lNumIntTitulo, .iNumParcela, .iStatus, .dtDataVencimento, .dtDataVencimentoReal, .dSaldo, .dValor, .iPortador, .iProxSeqBaixa, .iTipoCobranca, .iBancoCobrador, .sNossoNumero, .sCodigoDeBarras
    End With

    'Grava o Título a Pagar
    lErro = CF("NFFatPag_Grava_EmTrans", objTitPag, colParcelaPagar, objContabil)
    If lErro <> SUCESSO Then gError 190998
    
'    objTitPag.dCSLLRetido = objInfoFatura.dValorComissao
'
'    objCtb.sModelo = "Reembolso"
'
'    'gravar a contabilizacao
'    lErro = objContabAutomatica.Inicializa_Contab(objCtb, MODULO_CONTASAPAGAR, "CP", objTitPag.dtDataEmissao, objTitPag.dtDataEmissao)
'    If lErro <> SUCESSO Then gError 190999
'
'    lErro = objContabAutomatica.GeraContabilizacao(objTitPag)
'    If lErro <> SUCESSO Then gError 192000
'
'    Set objContabAutomatica = Nothing
'    Call objCtb.Encerrando
'    Set objCtb = Nothing
                    
    TRVFaturamento_Gera_TitPag = SUCESSO

    Exit Function

Erro_TRVFaturamento_Gera_TitPag:

    TRVFaturamento_Gera_TitPag = gErr

    Select Case gErr
    
        Case 190997 To 190999, 192000, 196904

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190991)

    End Select

    Exit Function

End Function

Private Function TRVFaturamento_Gera_NFPag(ByVal objInfoFatura As ClassTRVFATInfoFatura, ByVal objNFPag As ClassNFsPag, ByVal lProxNum As Long) As Long
'Comissões em geral
Dim lErro As Long

Dim objContabil As ClassContabil
Dim objCtb As New ClassImpCoinfoNFPagCtb
Dim objContabAutomatica As New ClassContabAutomatica

On Error GoTo Erro_TRVFaturamento_Gera_NFPag

    objNFPag.lFornecedor = objInfoFatura.lCliente
    objNFPag.iFilial = FILIAL_MATRIZ
    objNFPag.lNumNotaFiscal = lProxNum
    objNFPag.dtDataEmissao = objInfoFatura.dtDataEmissao
    objNFPag.dValorTotal = Abs(objInfoFatura.dValor)
    objNFPag.dValorProdutos = Abs(objInfoFatura.dValor)
    objNFPag.iFilialEmpresa = objInfoFatura.iFilialEmpresa
    objNFPag.dtDataVencimento = objInfoFatura.dtDataVencimento
               
    'Grava o Título a Pagar
    lErro = CF("NFPag_Grava_EmTrans", objNFPag, objContabil)
    If lErro <> SUCESSO Then gError 194203

'    'gravar a contabilizacao
'    lErro = objContabAutomatica.Inicializa_Contab(objCtb, MODULO_CONTASAPAGAR, "CP", objNFPag.dtDataEmissao, objNFPag.dtDataEmissao)
'    If lErro <> SUCESSO Then gError 194204
'
'    lErro = objContabAutomatica.GeraContabilizacao(objNFPag, False)
'    If lErro <> SUCESSO Then gError 194205
'
'    Set objContabAutomatica = Nothing
'    Call objCtb.Encerrando
'    Set objCtb = Nothing
                    
    TRVFaturamento_Gera_NFPag = SUCESSO

    Exit Function

Erro_TRVFaturamento_Gera_NFPag:

    TRVFaturamento_Gera_NFPag = gErr

    Select Case gErr
    
        Case 194203 To 194205

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194206)

    End Select

    Exit Function

End Function

Public Function Trata_Exclusao_Do_Destino(ByVal objDestino As Object) As Long

Dim lErro As Long
Dim alComando(0 To 22) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim dValorAporte As Double
Dim lNumIntDocAportePagto As Long
Dim dSaldo As Double
Dim lCodigo As Long
Dim lNumIntDocAporte As Long
Dim objAporte As ClassTRVAportes
Dim iTipoDestino As Integer
Dim lNumIntDocDestino As Long
Dim sClasse As String
Dim bValido As Boolean
Dim objTRVTitulosExp As ClassTRVTitulosExp
Dim lNumeroFaturaReal As Long
Dim sTipVou As String
Dim sSerie As String
Dim lNumVou As Long
Dim iTipoPagtoAporte As Integer
Dim bMotivoObrigatorio As Boolean
Dim objFaturaTRV As ClassFaturaTRV
Dim dValorAporteUSS As Double
Dim iMoedaAporte As Double
Dim objDebitoRecCli As ClassDebitoRecCli
Dim dValorAporteCred As Double
Dim lNumIntDocCredito As Long
Dim objFaturaAporte As ClassTRVFaturaAporte
Dim lNumIntDocVouInfo As Long

On Error GoTo Erro_Trata_Exclusao_Do_Destino

    bValido = True
    bMotivoObrigatorio = False
   
    sClasse = TypeName(objDestino)
    
    Select Case sClasse
    
        Case TRV_TIPO_DOC_DESTINO_CREDFORN_CLASSE
            iTipoDestino = TRV_TIPO_DOC_DESTINO_CREDFORN
            
        Case TRV_TIPO_DOC_DESTINO_DEBCLI_CLASSE
            iTipoDestino = TRV_TIPO_DOC_DESTINO_DEBCLI
            
        Case TRV_TIPO_DOC_DESTINO_TITPAG_CLASSE
            iTipoDestino = TRV_TIPO_DOC_DESTINO_TITPAG
            lNumeroFaturaReal = objDestino.lNumTitulo
            If objDestino.sSiglaDocumento = TIPODOC_FATURA_SERVICO_CP Then
                bMotivoObrigatorio = True
            End If
            
        Case TRV_TIPO_DOC_DESTINO_TITREC_CLASSE
            iTipoDestino = TRV_TIPO_DOC_DESTINO_TITREC
            lNumeroFaturaReal = objDestino.lNumTitulo
            If objDestino.sSiglaDocumento = TIPODOC_FATURA_SERVICO_CR Or objDestino.sSiglaDocumento = TIPODOC_CARTAO_CRED_DEB Then
                bMotivoObrigatorio = True
            End If
            
            'Muda o status das comissões que foram liberadas no faturamento
            'Não muda as já faturadas
            lErro = TRVVouchers_Comissao_Deslibera(TRV_TIPO_LIBERACAO_COMISSAO_FAT, objDestino.lNumIntDoc)
            If lErro <> SUCESSO Then gError 196393
            
        Case TRV_TIPO_DOC_DESTINO_NFSPAG_CLASSE
            iTipoDestino = TRV_TIPO_DOC_DESTINO_NFSPAG
            lNumeroFaturaReal = objDestino.lNumNotaFiscal
            bMotivoObrigatorio = True
            
        Case Else
            bValido = False
    
    End Select
    
    If bValido Then
    
        If bMotivoObrigatorio Then
        
            If objDestino.objInfoUsu Is Nothing Then gError 196490
            If TypeName(objDestino.objInfoUsu) <> "ClassFaturaTRV" Then gError 196490
        
            Set objFaturaTRV = objDestino.objInfoUsu
            
            If Len(Trim(objFaturaTRV.sMotivo)) = 0 Then gError 196491
        
        End If
    
        lNumIntDocDestino = objDestino.lNumIntDoc
        
        Set objTRVTitulosExp = New ClassTRVTitulosExp
        
        objTRVTitulosExp.iExcluido = MARCADO
        objTRVTitulosExp.iExportado = DESMARCADO
        objTRVTitulosExp.iTipoDocOrigem = iTipoDestino
        objTRVTitulosExp.lNumIntDocOrigem = lNumIntDocDestino
        objTRVTitulosExp.lNumTitulo = lNumeroFaturaReal
            
        If bMotivoObrigatorio Then
            
            objTRVTitulosExp.sMotivo = objFaturaTRV.sMotivo
        
            lErro = CF("TRVTitulosExp_Insere", objTRVTitulosExp)
            If lErro <> SUCESSO Then gError 196241
    
        End If
    
        'Abertura de Comando
        For iIndice = LBound(alComando) To UBound(alComando)
            alComando(iIndice) = Comando_Abrir()
            If alComando(iIndice) = 0 Then gError 192022
        Next
        
        sTipVou = String(STRING_TRV_OCR_TIPOVOU, 0)
        sSerie = String(STRING_TRV_OCR_SERIE, 0)

        lErro = Comando_ExecutarPos(alComando(0), "SELECT Recnsica,TipVou,Serie,NumVou FROM TRVVouchers WHERE TipoDocDestino= ? AND NumIntDocDestino = ? ", _
                    0, lNumIntDoc, sTipVou, sSerie, lNumVou, iTipoDestino, lNumIntDocDestino)
        If lErro <> AD_SQL_SUCESSO Then gError 192023
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192024
        
        Do While lErro = AD_SQL_SUCESSO
    
            'Trava registro contra alterações/Leituras
            lErro = Comando_LockExclusive(alComando(0))
            If lErro <> SUCESSO Then gError 192025
        
            lErro = Comando_ExecutarPos(alComando(1), "UPDATE TRVVouchers SET TipoDocDestino= ?, NumIntDocDestino=?, ValorAporte = ?,  ValorComissao = ?, NumFatCoinfo = ? ", alComando(0), 0, 0, 0, 0, 0)
            If lErro <> AD_SQL_SUCESSO Then gError 192026
            
            lErro = Comando_ExecutarPos(alComando(10), "SELECT TipoPagtoAporte, NumIntDocPagtoAporte, ValorAporte,ValorAporteUSS, MoedaAporte FROM TRVVoucherAporte WHERE TipVou= ? AND Serie = ? AND NumVou = ? ", _
                    0, iTipoPagtoAporte, lNumIntDocAportePagto, dValorAporte, dValorAporteUSS, iMoedaAporte, sTipVou, sSerie, lNumVou)
            If lErro <> AD_SQL_SUCESSO Then gError 192023
    
            'Busca Primeiro
            lErro = Comando_BuscarPrimeiro(alComando(10))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192024
            
            Do While lErro = AD_SQL_SUCESSO
        
                'Trava registro contra alterações/Leituras
                lErro = Comando_LockExclusive(alComando(10))
                If lErro <> SUCESSO Then gError 192025
        
                If Abs(dValorAporte) > DELTA_VALORMONETARIO Then
                
                    'Acerta o saldo do aporte
                    If iMoedaAporte = MOEDA_REAL Then
                        lErro = TRVAportePagtoFat_Acerta_Saldo(alComando(2), alComando(3), alComando(12), alComando(13), iTipoPagtoAporte, lNumIntDocAportePagto, -dValorAporte)
                    Else
                        lErro = TRVAportePagtoFat_Acerta_Saldo(alComando(2), alComando(3), alComando(12), alComando(13), iTipoPagtoAporte, lNumIntDocAportePagto, -dValorAporteUSS)
                    End If
                    If lErro <> SUCESSO Then gError 192027
                    
                End If
                    
                lErro = Comando_ExecutarPos(alComando(11), "DELETE TRVVoucherAporte ", alComando(10))
                If lErro <> AD_SQL_SUCESSO Then gError 192026
                
                'Busca Próximo
                lErro = Comando_BuscarProximo(alComando(10))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192028
                
            Loop
            
            'Busca Próximo
            lErro = Comando_BuscarProximo(alComando(0))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192028
            
        Loop
        
        'Le as ocorrências relacionadas
        lErro = Comando_ExecutarPos(alComando(4), "SELECT NumIntDoc, Codigo FROM TRVOcorrencias WHERE TipoDocDestino= ? AND NumIntDocDestino = ? ", _
                    0, lNumIntDoc, lCodigo, iTipoDestino, lNumIntDocDestino)
        If lErro <> AD_SQL_SUCESSO Then gError 192029
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(4))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192030
        
        Do While lErro = AD_SQL_SUCESSO
        
            If iTipoDestino = TRV_TIPO_DOC_DESTINO_CREDFORN Then gError 192034
            If iTipoDestino = TRV_TIPO_DOC_DESTINO_DEBCLI Then gError 192035
    
            'Trava registro contra alterações/Leituras
            lErro = Comando_LockExclusive(alComando(4))
            If lErro <> SUCESSO Then gError 192031
        
            lErro = Comando_ExecutarPos(alComando(5), "UPDATE TRVOcorrencias SET TipoDocDestino= ?, NumIntDocDestino=?, NumTitulo = ?, ValorAporte = ? ", alComando(4), 0, 0, 0, 0)
            If lErro <> AD_SQL_SUCESSO Then gError 192032
            
            lErro = Comando_ExecutarPos(alComando(14), "SELECT TipoPagtoAporte, NumIntDocPagtoAporte, ValorAporte,ValorAporteUSS, MoedaAporte FROM TRVOcorrenciaAporte WHERE NumIntDocOCR= ? ", _
                    0, iTipoPagtoAporte, lNumIntDocAportePagto, dValorAporte, dValorAporteUSS, iMoedaAporte, lNumIntDoc)
            If lErro <> AD_SQL_SUCESSO Then gError 192023
    
            'Busca Primeiro
            lErro = Comando_BuscarPrimeiro(alComando(14))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192024
            
            Do While lErro = AD_SQL_SUCESSO
        
                'Trava registro contra alterações/Leituras
                lErro = Comando_LockExclusive(alComando(14))
                If lErro <> SUCESSO Then gError 192025
        
                If Abs(dValorAporte) > DELTA_VALORMONETARIO Then
                
                    'Acerta o saldo do aporte
                    If iMoedaAporte = MOEDA_REAL Then
                        lErro = TRVAportePagtoFat_Acerta_Saldo(alComando(2), alComando(3), alComando(12), alComando(13), iTipoPagtoAporte, lNumIntDocAportePagto, -dValorAporte)
                    Else
                        lErro = TRVAportePagtoFat_Acerta_Saldo(alComando(2), alComando(3), alComando(12), alComando(13), iTipoPagtoAporte, lNumIntDocAportePagto, -dValorAporteUSS)
                    End If
                    If lErro <> SUCESSO Then gError 192027
                    
                End If
                    
                lErro = Comando_ExecutarPos(alComando(15), "DELETE TRVOcorrenciaAporte ", alComando(14))
                If lErro <> AD_SQL_SUCESSO Then gError 192026
                
                'Busca Próximo
                lErro = Comando_BuscarProximo(alComando(14))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192028
                
            Loop
            
            'Busca Próximo
            lErro = Comando_BuscarProximo(alComando(4))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192033
            
        Loop
        
        'Le as comissões relacionadas
        lErro = Comando_ExecutarPos(alComando(21), "SELECT NumIntDoc FROM TRVVoucherInfo WHERE TipoDocDestino= ? AND NumIntDocDestino = ? ", _
                    0, lNumIntDocVouInfo, iTipoDestino, lNumIntDocDestino)
        If lErro <> AD_SQL_SUCESSO Then gError 196381
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(21))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196382
        
        Do While lErro = AD_SQL_SUCESSO
    
            'Trava registro contra alterações/Leituras
            lErro = Comando_LockExclusive(alComando(21))
            If lErro <> SUCESSO Then gError 196383

            lErro = Comando_ExecutarPos(alComando(22), "UPDATE TRVVoucherInfo SET TipoDocDestino= ?, NumIntDocDestino=?, NumTitulo = ? ", alComando(21), 0, 0, 0)
            If lErro <> AD_SQL_SUCESSO Then gError 196384
            
            'Busca Próximo
            lErro = Comando_BuscarProximo(alComando(21))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196385
            
        Loop
        
        'Le os aportes diretos
        lErro = Comando_ExecutarPos(alComando(6), "SELECT NumIntDocAporte FROM TRVAportePagtoDireto WHERE TipoDocDestino= ? AND NumIntDocDestino = ? ", _
                    0, lNumIntDocAporte, iTipoDestino, lNumIntDocDestino)
        If lErro <> AD_SQL_SUCESSO Then gError 192036
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(6))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192037
        
        Do While lErro = AD_SQL_SUCESSO
        
            Set objAporte = New ClassTRVAportes
            
            objAporte.lNumIntDoc = lNumIntDocAporte
            
            lErro = CF("TRVAportes_Le", objAporte, False)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 192038
        
            If iTipoDestino = TRV_TIPO_DOC_DESTINO_CREDFORN Then gError 192039
            If iTipoDestino = TRV_TIPO_DOC_DESTINO_DEBCLI Then gError 192040
    
            'Trava registro contra alterações/Leituras
            lErro = Comando_LockExclusive(alComando(6))
            If lErro <> SUCESSO Then gError 192041
        
            lErro = Comando_ExecutarPos(alComando(7), "UPDATE TRVAportePagtoDireto SET TipoDocDestino= ?, NumIntDocDestino=? ", alComando(6), 0, 0)
            If lErro <> AD_SQL_SUCESSO Then gError 192042
            
            'Busca Próximo
            lErro = Comando_BuscarProximo(alComando(6))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192043
            
        Loop
        
        'Le os aportes condicionados
        lErro = Comando_ExecutarPos(alComando(8), "SELECT NumIntDocAporte FROM TRVAportePagtoCond WHERE TipoDocDestino= ? AND NumIntDocDestino = ? ", _
                    0, lNumIntDocAporte, iTipoDestino, lNumIntDocDestino)
        If lErro <> AD_SQL_SUCESSO Then gError 192044
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(8))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192045
        
        Do While lErro = AD_SQL_SUCESSO
        
            Set objAporte = New ClassTRVAportes
            
            objAporte.lNumIntDoc = lNumIntDocAporte
            
            lErro = CF("TRVAportes_Le", objAporte, False)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 192046
        
            If iTipoDestino = TRV_TIPO_DOC_DESTINO_CREDFORN Then gError 192047
            If iTipoDestino = TRV_TIPO_DOC_DESTINO_DEBCLI Then gError 192048
    
            'Trava registro contra alterações/Leituras
            lErro = Comando_LockExclusive(alComando(8))
            If lErro <> SUCESSO Then gError 192049
        
            lErro = Comando_ExecutarPos(alComando(9), "UPDATE TRVAportePagtoCond SET TipoDocDestino= ?, NumIntDocDestino=? ", alComando(8), 0, 0)
            If lErro <> AD_SQL_SUCESSO Then gError 192050
            
            'Busca Próximo
            lErro = Comando_BuscarProximo(alComando(8))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192051
            
        Loop
        
        If iTipoDestino = TRV_TIPO_DOC_DESTINO_TITREC Then
        
            'Le as informações do faturamento
            lErro = Comando_Executar(alComando(16), "SELECT A.ValorAporte, A.NumIntDocCredito FROM TRVTitulosExp AS T, TRVTitulosExpAporte AS A  WHERE T.NumIntDoc = A.NumIntDocTitExp AND T.TipoDocOrigem = ? AND T.NumIntDocOrigem = ? AND T.Excluido = ? AND A.NumIntDocCredito <> ? ", _
                        dValorAporteCred, lNumIntDocCredito, iTipoDestino, lNumIntDocDestino, DESMARCADO, 0)
            If lErro <> AD_SQL_SUCESSO Then gError 192036
        
            'Busca Primeiro
            lErro = Comando_BuscarPrimeiro(alComando(16))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192037
        
            'Se recebeu um aporte via crédito tem que devolver o saldo
            Do While lErro = AD_SQL_SUCESSO
        
                Set objDebitoRecCli = New ClassDebitoRecCli
    
                objDebitoRecCli.lNumIntDoc = lNumIntDocCredito
                objDebitoRecCli.iFilialEmpresa = FILIAL_MATRIZ
                
                'Lê os dados para dos Débitos
                lErro = CF("DebitoReceber_Le", objDebitoRecCli)
                If lErro <> SUCESSO And lErro <> 17835 Then gError 192784
            
                'Altera o saldo do Débito a Receber
                lErro = CF("DebitoRecCli_AlterarSaldo", objDebitoRecCli, dValorAporteCred)
                If lErro <> SUCESSO Then gError 192784
                
                'Busca Próximo
                lErro = Comando_BuscarProximo(alComando(16))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192037
                
            Loop
        
            'Exclui as informações específicas de Títulos a Receber da Travel Ace
            lErro = CF("TitulosRecTRV_Exclui", lNumIntDocDestino, objFaturaTRV)
            If lErro <> SUCESSO Then gError 192784
            
        End If
        
        'Fecha Comando
        For iIndice = LBound(alComando) To UBound(alComando)
            Call Comando_Fechar(alComando(iIndice))
        Next
        
    End If
   

    Trata_Exclusao_Do_Destino = SUCESSO

    Exit Function

Erro_Trata_Exclusao_Do_Destino:

    Trata_Exclusao_Do_Destino = gErr

    Select Case gErr
        
        Case 192022
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 192023, 192024, 192028
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVVOUCHERS", gErr)
        
        Case 192025
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRVVOUCHERS", gErr)
        
        Case 192026
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRVVOUCHERS", gErr)
            
        Case 192027, 192038, 192046, 192784, 196241
            
        Case 192029, 192030, 192033
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVOCORRENCIAS", gErr)
        
        Case 192031
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRVOCORRENCIAS", gErr)
        
        Case 192032
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRVOCORRENCIAS", gErr)
            
        Case 192034
            Call Rotina_Erro(vbOKOnly, "ERRO_CRED_VINC_ORC", gErr, lCodigo)
        
        Case 192035
            Call Rotina_Erro(vbOKOnly, "ERRO_DEB_VINC_OCR", gErr, lCodigo)
                    
        Case 192036, 192037, 192043
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVAPORTEPAGTODIRETO", gErr)
        
        Case 192039
            Call Rotina_Erro(vbOKOnly, "ERRO_CRED_VINC_APORTE", gErr, objAporte.lCodigo)
        
        Case 192040
            Call Rotina_Erro(vbOKOnly, "ERRO_DEB_VINC_APORTE", gErr, objAporte.lCodigo)
        
        Case 192041
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRVAPORTEPAGTODIRETO", gErr)
        
        Case 192042
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRVAPORTEPAGTODIRETO", gErr)
            
        Case 192044, 192045, 192051
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVAPORTEPAGTOCOND", gErr)
        
        Case 192047
            Call Rotina_Erro(vbOKOnly, "ERRO_CRED_VINC_APORTE", gErr, objAporte.lCodigo)
        
        Case 192048
            Call Rotina_Erro(vbOKOnly, "ERRO_DEB_VINC_APORTE", gErr, objAporte.lCodigo)
        
        Case 192049
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRVAPORTEPAGTOCOND", gErr)
        
        Case 192050
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRVAPORTEPAGTOCOND", gErr)
            
        Case 196381, 196382, 196385
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVVOUCHERINFO", gErr)
        
        Case 196383
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRVVOUCHERINFO", gErr)
        
        Case 196384
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRVVOUCHERINFO", gErr)
            
        Case 196393
            
        Case 196490
            Call Rotina_Erro(vbOKOnly, "ERRO_TRV_CANC_FAT_TELA", gErr)
        
        Case 196491
            Call Rotina_Erro(vbOKOnly, "ERRO_MOTIVO_NAO_INFORMADO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192052)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Private Function TRVAportePagtoFat_Acerta_Saldo(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lComando3 As Long, ByVal lComando4 As Long, ByVal iTipoAporte As Integer, ByVal lNumIntDocAportePagto As Long, ByVal dValorUsado As Double) As Long

Dim lErro As Long
Dim dSaldo As Double
Dim dValor As Double

On Error GoTo Erro_TRVAportePagtoFat_Acerta_Saldo

    If lNumIntDocAportePagto <> 0 And dValorUsado <> 0 Then
    
        If iTipoAporte = TRV_TIPO_APORTE_SOBREFATURA Then
   
            'Acerta o saldo do aporte
            lErro = Comando_ExecutarPos(lComando1, "SELECT Valor, Saldo  FROM TRVAportePagtoFat WHERE NumIntDoc= ? ", 0, dValor, dSaldo, lNumIntDocAportePagto)
            If lErro <> AD_SQL_SUCESSO Then gError 192011
        
            'Busca Primeiro
            lErro = Comando_BuscarPrimeiro(lComando1)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192012
            
            If lErro = AD_SQL_SEM_DADOS Then gError 192017
        
            'Trava registro contra alterações/Leituras
            lErro = Comando_LockExclusive(lComando1)
            If lErro <> SUCESSO Then gError 192013
            
            dSaldo = dSaldo - dValorUsado
            
            If dSaldo < -DELTA_VALORMONETARIO2 Then gError 192015
            If dSaldo > dValor + DELTA_VALORMONETARIO Then gError 192019
            
            dSaldo = Arredonda_Moeda(dSaldo, 4)
        
            lErro = Comando_ExecutarPos(lComando2, "UPDATE TRVAportePagtoFat SET Saldo = ? ", lComando1, dSaldo)
            If lErro <> AD_SQL_SUCESSO Then gError 192014
        
        Else
        
            'Acerta o saldo do aporte
            lErro = Comando_ExecutarPos(lComando3, "SELECT Valor FROM TRVAportePagtoFatCond WHERE NumIntDoc= ? ", 0, dValor, lNumIntDocAportePagto)
            If lErro <> AD_SQL_SUCESSO Then gError 192011
        
            'Busca Primeiro
            lErro = Comando_BuscarPrimeiro(lComando3)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192012
            
            If lErro = AD_SQL_SEM_DADOS Then gError 192017
        
            'Trava registro contra alterações/Leituras
            lErro = Comando_LockExclusive(lComando3)
            If lErro <> SUCESSO Then gError 192013
            
            dValor = dValor + dValorUsado
            
            dValor = Arredonda_Moeda(dValor, 4)
        
            lErro = Comando_ExecutarPos(lComando4, "UPDATE TRVAportePagtoFatCond SET Valor = ? ", lComando3, dValor)
            If lErro <> AD_SQL_SUCESSO Then gError 192014
        
        End If
    
    End If

    TRVAportePagtoFat_Acerta_Saldo = SUCESSO

    Exit Function

Erro_TRVAportePagtoFat_Acerta_Saldo:

    TRVAportePagtoFat_Acerta_Saldo = gErr

    Select Case gErr

        Case 192011, 192012
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVAPORTEPAGTOFAT", gErr)
        
        Case 192013
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRVAPORTEPAGTOFAT", gErr)
        
        Case 192014
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRVAPORTEPAGTOFAT", gErr)

        Case 192015
            Call Rotina_Erro(vbOKOnly, "ERRO_APORTE_SALDO_NEGATIVO", gErr, lNumIntDocAportePagto)
            
        Case 192017
            Call Rotina_Erro(vbOKOnly, "ERRO_TRVAPORTEPAGTOFAT_NAO_CADASTRADO", gErr, lNumIntDocAportePagto)

        Case 192019
            Call Rotina_Erro(vbOKOnly, "ERRO_APORTE_SALDO_MAIOR_VALOR", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192020)

    End Select

    Exit Function

End Function

Private Function TRVAportes_Acerta_ProxParcela(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lCodigoAporte As Long, ByRef iParcela As Integer) As Long

Dim lErro As Long
Dim iProxParcela As Integer

On Error GoTo Erro_TRVAportes_Acerta_ProxParcela

     'Acerta o saldo do aporte
     lErro = Comando_ExecutarPos(lComando1, "SELECT ProxParcela FROM TRVAportes WHERE Codigo= ? ", 0, iParcela, lCodigoAporte)
     If lErro <> AD_SQL_SUCESSO Then gError 196470
 
     'Busca Primeiro
     lErro = Comando_BuscarPrimeiro(lComando1)
     If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196471
     
     If lErro = AD_SQL_SEM_DADOS Then gError 196472
 
     'Trava registro contra alterações/Leituras
     lErro = Comando_LockExclusive(lComando1)
     If lErro <> SUCESSO Then gError 196473
     
     iProxParcela = iParcela + 1
 
     lErro = Comando_ExecutarPos(lComando2, "UPDATE TRVAportes SET ProxParcela = ? ", lComando1, iProxParcela)
     If lErro <> AD_SQL_SUCESSO Then gError 196474

    TRVAportes_Acerta_ProxParcela = SUCESSO

    Exit Function

Erro_TRVAportes_Acerta_ProxParcela:

    TRVAportes_Acerta_ProxParcela = gErr

    Select Case gErr

        Case 196470, 196471
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVAPORTES", gErr)
        
        Case 196472
            Call Rotina_Erro(vbOKOnly, "ERRO_TRVAPORTES_NAO_CADASTRADO", gErr, lCodigoAporte)
        
        Case 196473
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRVAPORTES", gErr)
        
        Case 196474
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRVAPORTES", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196475)

    End Select

    Exit Function

End Function

Function TRVConfig_Grava(ByVal colConfigs As Collection) As Long
'grava na tabela TRVConfig uma colecao de registros de configuracao

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_TRVConfig_Grava
    
    'Abrir transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 192072

    lErro = TRVConfig_Grava_Trans(colConfigs)
    If lErro <> SUCESSO Then gError 192073

    'Fechar transacao
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 192074
          
    TRVConfig_Grava = SUCESSO
    
    Exit Function
    
Erro_TRVConfig_Grava:
    
    TRVConfig_Grava = gErr
    
    Select Case gErr
    
        Case 192072
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 192073

        Case 192074
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192075)
            
    End Select
    
    Call Transacao_Rollback
    
    Exit Function
    
End Function

Public Function TRVConfig_Grava_Trans(ByVal colConfigs As Collection) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(1 To 2) As Long
Dim objTRVConfig As ClassTRVConfig
Dim sConteudo As String

On Error GoTo Erro_TRVConfig_Grava_Trans
    
    'abrir comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 192076
    Next
    
   'para cada elemento da colecao
    For Each objTRVConfig In colConfigs
    
        'inicializar buffer
        sConteudo = String(STRING_CONFIG_CONTEUDO, 0)
        
        'ler registo
        lErro = Comando_ExecutarPos(alComando(1), "SELECT Conteudo FROM TRVConfig WHERE Codigo = ? AND FilialEmpresa = ?", 0, sConteudo, objTRVConfig.sCodigo, objTRVConfig.iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 192077
        
        'Lê o primeiro registro
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192078
        
        'Se não encontrou o registro
        If lErro = AD_SQL_SEM_DADOS Then gError 192079
        
        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(1))
        If lErro <> SUCESSO Then gError 192250
            
        lErro = Comando_ExecutarPos(alComando(2), "UPDATE TRVConfig SET Conteudo = ?", alComando(1), objTRVConfig.sConteudo)
        If lErro <> AD_SQL_SUCESSO Then gError 192080

    Next
    
   'fechar comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
          
    TRVConfig_Grava_Trans = SUCESSO
    
    Exit Function
    
Erro_TRVConfig_Grava_Trans:
    
    TRVConfig_Grava_Trans = gErr
    
    Select Case gErr
    
        Case 192076
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 192077, 192078
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVCONFIG", gErr)

        Case 192079
            Call Rotina_Erro(vbOKOnly, "ERRO_TRVCONFIG_INEXISTENTE", gErr, objTRVConfig.sCodigo, objTRVConfig.iFilialEmpresa)
        
        Case 192080
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRVCONFIG", gErr)
            
        Case 192250
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_TRVCONFIG", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192081)
            
    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Public Function TRVConfig_Grava_Lockado(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal objTRVConfig As ClassTRVConfig) As Long

Dim lErro As Long

On Error GoTo Erro_TRVConfig_Grava_Lockado
 
    lErro = Comando_ExecutarPos(lComando2, "UPDATE TRVConfig SET Conteudo = ?", lComando1, objTRVConfig.sConteudo)
    If lErro <> AD_SQL_SUCESSO Then gError 194251
          
    TRVConfig_Grava_Lockado = SUCESSO
    
    Exit Function
    
Erro_TRVConfig_Grava_Lockado:
    
    TRVConfig_Grava_Lockado = gErr
    
    Select Case gErr

        Case 194251
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRVCONFIG", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 194252)
            
    End Select
    
    Exit Function
    
End Function


Function Gera_Arquivo_Html(ByVal objTela As Object, ByVal sModelo As String, ByVal sArquivo As String, sConteudoHtml As String) As Long

Dim lErro As Long
Dim objTS1 As TextStream
Dim objTS2 As TextStream
Dim objFSO As New FileSystemObject
Dim sRegistro As String
Dim bArqAberto1 As Boolean
Dim bArqAberto2 As Boolean
Dim lPos1 As Long
Dim lPos2 As Long
Dim objMnemonicoValor As ClassMnemonicoValor
Dim sValorTexto As String

On Error GoTo Erro_Gera_Arquivo_Html

    If Len(Trim(sModelo)) > 0 Then
    
        bArqAberto1 = False
        Set objTS1 = objFSO.OpenTextFile(sModelo, IOMode.ForReading, False)
        bArqAberto1 = True

        bArqAberto2 = False
        Set objTS2 = objFSO.OpenTextFile(sArquivo, IOMode.ForWriting, True)
        bArqAberto2 = True

        'Até chegar ao fim do arquivo
        Do While Not objTS1.AtEndOfLine
        
            'Busca o próximo registro do arquivo
            sRegistro = objTS1.ReadAll
             
            lPos1 = InStr(1, sRegistro, "ZZZ")
             
            Do While lPos1 <> 0
             
                lPos2 = InStr(lPos1 + 3, sRegistro, "ZZZ")
                 
                Set objMnemonicoValor = New ClassMnemonicoValor
                Set objMnemonicoValor.colValor = New Collection
                 
                objMnemonicoValor.sMnemonico = Mid(sRegistro, lPos1 + 3, lPos2 - lPos1 - 3)
                 
                lErro = objTela.Calcula_Mnemonico(objMnemonicoValor, sValorTexto)
                If lErro <> SUCESSO Then gError 192209
             
                sRegistro = Replace(sRegistro, "ZZZ" & objMnemonicoValor.sMnemonico & "ZZZ", sValorTexto)
                
                lPos1 = InStr(lPos1 + Len(sValorTexto), sRegistro, "ZZZ")
            Loop
            
            objTS2.Write sRegistro
            
            sConteudoHtml = sRegistro

        Loop
        
        objTS1.Close
        bArqAberto1 = False
        objTS2.Close
        bArqAberto2 = False

    End If

    Gera_Arquivo_Html = SUCESSO

    Exit Function

Erro_Gera_Arquivo_Html:

    Gera_Arquivo_Html = gErr

    Select Case gErr
    
        Case 53 'ERRO_ARQUIVO_FTP_NAO_ENCONTRADO
            Call Rotina_Erro(vbOKOnly, "ERRO_ARQUIVO_FTP_NAO_ENCONTRADO", gErr, sModelo)
            
        Case 192209

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192193)

    End Select
    
    'fechar aquivo texto
    If bArqAberto1 Then
        objTS1.Close
    End If

    'fechar aquivo texto
    If bArqAberto2 Then
        objTS2.Close
    End If

    Exit Function

End Function

Function TRVFaturamento_Envia_PorEmail(ByVal objFaturamento As ClassTRVFaturamento, ByVal colParcelas As Collection) As Long

Dim lErro As Long
Dim objParcelaRec As ClassParcelaReceber
Dim objCobrancaSel As New ClassCobrancaPorEmailSel

On Error GoTo Erro_TRVFaturamento_Envia_PorEmail

    For Each objParcelaRec In colParcelas
    
        lErro = CF("ParcelaReceber_Le", objParcelaRec)
        If lErro <> SUCESSO And lErro <> 19147 Then gError 185994

        objCobrancaSel.colParcelas.Add objParcelaRec

    Next
    
    objCobrancaSel.iFaturamento = MARCADO
    objCobrancaSel.lModeloForcado = objFaturamento.lModelo
    
    Call Chama_Tela("CobrancaPorEmail", TIPOTELA_EMAIL_COBRANCA, objCobrancaSel)

    TRVFaturamento_Envia_PorEmail = SUCESSO

    Exit Function

Erro_TRVFaturamento_Envia_PorEmail:

    TRVFaturamento_Envia_PorEmail = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192210)

    End Select
    

    Exit Function

End Function

Public Function TRV_NF_Gera(ByVal objGeracaoNF As ClassTRVGeracaoNF) As Long

Dim lErro As Long
Dim lTransacao As Long
Dim alComando(0 To 55) As Long
Dim iIndice As Integer

On Error GoTo Erro_TRV_NF_Gera

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 192283
    
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 192284
    Next

    lErro = TRV_NF_Gera_EmTrans(alComando, objGeracaoNF)
    If lErro <> SUCESSO Then gError 192285

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
   
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 192286
        
    TRV_NF_Gera = SUCESSO

    Exit Function

Erro_TRV_NF_Gera:

    TRV_NF_Gera = gErr

    Select Case gErr
        
        Case 192283
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 192284
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 192285

        Case 192286
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192287)

    End Select
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Private Function TRV_NF_Gera_EmTrans(alComando() As Long, ByVal objGeracaoNF As ClassTRVGeracaoNF) As Long

Dim lErro As Long
Dim objNF As ClassNFiscal
Dim objItemNF As ClassItemNF
Dim objTituloRec As ClassTituloReceber
Dim iIndice As Integer
Dim dComissao As Double
Dim lNumVou As Long
Dim sSerie As String
Dim sTipoDoc As String
Dim lNumNotaAnt As Long
Dim lNumIntDocItemNF As Long
Dim lNumNota As Long
Dim iFilialEmpresaAnt As Integer
Dim lNumIntDoc As Long
Dim iProxComando As Integer
Dim objNFTRV As ClassNFTRV
Dim iFilialEmpAnt As Integer
Dim objFilialEmpresa As AdmFiliais
Dim colFiliais As New Collection
Dim bAchou As Boolean
Dim objFilialCliente As ClassFilialCliente
Dim objEndereco As ClassEndereco
Dim iProxComandoAnt As Integer
Dim objFrmAguarde As New ClassFrmAguarde
Dim objFrmAguardeTela As New FrmAguarde

On Error GoTo Erro_TRV_NF_Gera_EmTrans

    iFilialEmpresaAnt = giFilialEmpresa
    
    iProxComando = 0
    
    objFrmAguarde.iTotalItens = objGeracaoNF.colNF.Count
    Call objFrmAguardeTela.Inicializa_Progressao(objFrmAguarde)
    
    For Each objNF In objGeracaoNF.colNF
    
        Set objNFTRV = objNF.objInfoUsu
        
        bAchou = False
        For Each objFilialEmpresa In colFiliais
            If objFilialEmpresa.iCodFilial = objNF.iFilialEmpresa Then
                bAchou = True
                Exit For
            End If
        Next
        If Not bAchou Then
        
            Set objFilialEmpresa = New AdmFiliais
        
            objFilialEmpresa.iCodFilial = objNF.iFilialEmpresa
        
            'Le os dados da Filial Empresa
            lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
            If lErro <> SUCESSO And lErro <> 27378 Then gError 196817
            
        End If
        
        Set objNFTRV.objFilialEmpresa = objFilialEmpresa
        
        Set objFilialCliente = New ClassFilialCliente
        Set objEndereco = New ClassEndereco
        
        objFilialCliente.lCodCliente = objNF.lCliente
        objFilialCliente.iCodFilial = objNF.iFilialCli
        
        lErro = CF("FilialCliente_Le", objFilialCliente, alComando(iProxComando))
        If lErro <> SUCESSO And lErro <> 12567 Then gError 196818
                
        objEndereco.lCodigo = objFilialCliente.lEndereco
        
        lErro = CF("Endereco_Le", objEndereco, alComando(iProxComando + 1))
        If lErro <> SUCESSO Then gError 196819
        
        Set objNFTRV.objFilialCliente = objFilialCliente
        Set objNFTRV.objEnderecoCli = objEndereco
        
    Next
    
    iProxComando = iProxComando + 1
    iProxComandoAnt = iProxComando
    
    iIndice = 0
    For Each objNF In objGeracaoNF.colNF
    
        giFilialEmpresa = objNF.iFilialEmpresa

        lNumNotaAnt = objNF.lNumNotaFiscal
        objNF.lNumNotaFiscal = 0
        
        If objNF.iFilialEmpresa = FILIAL_MATRIZ Then
            objNF.iRecibo = MARCADO
        End If
        
        iProxComando = iProxComandoAnt
    
        lErro = NF_NFiscal_Grava_TRV(alComando(), objNF, iProxComando)
        If lErro <> SUCESSO Then gError 192289
        
        If lNumNotaAnt <> objNF.lNumNotaFiscal Then gError 192290
        
        For Each objItemNF In objNF.ColItensNF
        
            iIndice = iIndice + 1
            
            Set objTituloRec = objGeracaoNF.colTitulos.Item(iIndice)
            
            If Abs(objItemNF.dValorTotal - objTituloRec.dValor) > DELTA_VALORMONETARIO Then gError 192299
            
            'Busca os vouchers da fatura para poder calcular a comissão
            lErro = Comando_Executar(alComando(iProxComando + 2), "SELECT X.NumIntDocItemNF, N.NumNotaFiscal FROM TRVItemNotaTitulo AS X, ItensNFiscal AS I, NFiscal AS N WHERE N.NumIntDoc = I.NumIntNF AND N.Status <> 7 AND I.NumIntDoc = X.NumIntDocItemNF AND X.NumIntDocTitulo = ? ", lNumIntDocItemNF, lNumNota, objTituloRec.lNumIntDoc)
            If lErro <> AD_SQL_SUCESSO Then gError 192300
            
            lErro = Comando_BuscarPrimeiro(alComando(iProxComando + 2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192301
            
            If lErro = AD_SQL_SUCESSO Then gError 192302
            
            lErro = Comando_Executar(alComando(iProxComando), "INSERT INTO TRVItemNotaTitulo (NumIntDocItemNF,NumIntDocTitulo) VALUES (?,?)", objItemNF.lNumIntDoc, objTituloRec.lNumIntDoc)
            If lErro <> AD_SQL_SUCESSO Then gError 192291
            
            sSerie = String(STRING_TRV_OCR_SERIE, 0)
            sTipoDoc = String(STRING_TRV_OCR_TIPODOC, 0)

            'Busca as informações de faturamento no Corporator, se não encontrar desce até o voucher para obter a informação da comissão
            lErro = Comando_Executar(alComando(iProxComando + 4), "SELECT ValorComissao FROM TRVTitulosExp WHERE TipoDocOrigem = ? AND NumIntDocOrigem = ? AND Excluido = ? ", dComissao, TRV_TIPO_DOC_DESTINO_TITREC, objTituloRec.lNumIntDoc, DESMARCADO)
            If lErro <> AD_SQL_SUCESSO Then gError 196820

            lErro = Comando_BuscarPrimeiro(alComando(iProxComando + 4))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196821

            If lErro <> AD_SQL_SUCESSO Then

                'Verifica as comissões dadas nos vouchers
                lErro = Comando_Executar(alComando(iProxComando + 1), "SELECT SUM(ValorComissao) FROM TRVVouchers WHERE NumFatCoinfo = ? ", dComissao, objTituloRec.lNumTitulo)
                If lErro <> AD_SQL_SUCESSO Then gError 192292
    
                lErro = Comando_BuscarPrimeiro(alComando(iProxComando + 1))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192293
        
            End If
        
        Next
       
        'Obter NumIntDoc
        lErro = CF("Config_ObterNumInt2", alComando(iProxComando + 5), alComando(iProxComando + 6), "TRVConfig", "NUM_INT_PROX_TRVNFSEXP", lNumIntDoc)
        If lErro <> SUCESSO Then gError 196822
   
        lErro = Comando_Executar(alComando(iProxComando + 3), "INSERT INTO TRVNFsExp( NumIntDoc, Usuario, Data, Hora, " & _
                "NumIntDocNF, NumNF, TemQueContabilizar,ValorComissao) VALUES (?,?,?,?,?,?,?,?)", _
                lNumIntDoc, gsUsuario, Date, CDbl(Time), objNF.lNumIntDoc, objNF.lNumNotaFiscal, MARCADO, dComissao)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196823
    
        Call objFrmAguardeTela.ProcessouItem
        
        If objFrmAguarde.iCancelar = MARCADO Then Exit For

    Next
    
    Set objFrmAguardeTela = Nothing
    If objFrmAguarde.iCancelar = MARCADO Then gError 196784
    
    giFilialEmpresa = iFilialEmpresaAnt

    TRV_NF_Gera_EmTrans = SUCESSO

    Exit Function

Erro_TRV_NF_Gera_EmTrans:

    giFilialEmpresa = iFilialEmpresaAnt

    TRV_NF_Gera_EmTrans = gErr

    Select Case gErr
    
        Case 192289, 192294, 192296 To 192298, 196784
        
        Case 192290
            Call Rotina_Erro(vbOKOnly, "ERRO_NUM_NOTA_DIFERENTE", gErr, lNumNotaAnt, objNF.lNumNotaFiscal)
        
        Case 192291 'INSERT
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRVITEMNOTATITULO", gErr)
        
        Case 192292, 192293
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVVOUCHERS", gErr)
        
        Case 192299
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_ITEM_ERRADO", gErr, Format(objItemNF.dValorTotal, "STANDARD"), Format(objTituloRec.dValor, "STANDARD"))

        Case 192300, 192301
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVITEMNOTATITULO", gErr)
            
        Case 192302
            Call Rotina_Erro(vbOKOnly, "ERRO_TITULO_JA_COM_NOTA", gErr, objTituloRec.lNumTitulo, lNumNota)

        Case 196817 To 196819, 196822
        
        Case 196820, 196821
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVTITULOSEXP", gErr)
            
        Case 196823
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRVNFSEXP", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192288)

    End Select
    
    If Not (objFrmAguardeTela Is Nothing) Then
        Call objFrmAguardeTela.Trata_Erro
    End If

    Exit Function

End Function

Public Function BaixasParcelaRec_Grava_Customizado(ByVal objTituloRec As ClassTituloReceber, ByVal objParcelaRec As ClassParcelaReceber, ByVal objBaixaParcRec As ClassBaixaParcRec) As Long

Dim lErro As Long
Dim alComando(1 To 10) As Long
Dim iIndice As Integer
Dim lNumVou As Long
Dim sSerie As String
Dim sTipoDoc As String
Dim sTipoVou As String
Dim objOcorrencia As ClassTRVOcorrencias
Dim objOcorrenciaDet As ClassTRVOcorrenciaDet
Dim dValor As Double
Dim lCodigo As Long
Dim objVoucherInfo As ClassTRVVoucherInfo
Dim iCartao As Integer
Dim lCliente As Long
Dim dValorComissao As Double
Dim lNumIntComi As Long
Dim lNumIntDocVouInfo As Long
Dim lNumIntDocVouInfoAux As Long
Dim dtDataVou As Date
Dim sConteudo As String

On Error GoTo Erro_BaixasParcelaRec_Grava_Customizado
 
     'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 192308
    Next

    sSerie = String(STRING_TRV_OCR_SERIE, 0)
    sTipoDoc = String(STRING_TRV_OCR_TIPODOC, 0)
    sTipoVou = String(STRING_TRV_OCR_TIPOVOU, 0)

    'Busca os vouchers cancelados para poder gerar as ocorrências de reembolso
    lErro = Comando_ExecutarPos(alComando(1), "SELECT NumVou, Serie, TipoDoc, Valor, TipVou, Cartao FROM TRVVouchers WHERE Status = ? AND TipoDocDestino = ? AND NumIntDocDestino = ? AND NumIntDocNVL = ? ", 0, lNumVou, sSerie, sTipoDoc, dValor, sTipoVou, iCartao, STATUS_CANCELADO, TRV_TIPO_DOC_DESTINO_TITREC, objTituloRec.lNumIntDoc, 0)
    If lErro <> AD_SQL_SUCESSO Then gError 192309
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192310

    'Enquanto existirem vouchers cancelados
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objOcorrencia = New ClassTRVOcorrencias
        Set objOcorrenciaDet = New ClassTRVOcorrenciaDet
        Set objVoucherInfo = New ClassTRVVoucherInfo
        
        objOcorrencia.dtDataEmissao = gdtDataAtual
        objOcorrencia.dValorTotal = -dValor
        objOcorrencia.iFilialCliente = objTituloRec.iFilial
        objOcorrencia.iFormaPagto = FORMAPAGTO_TRV_OCR_FAT
        objOcorrencia.iOrigem = INATIVACAO_AUTOMATICA_CODIGO
        objOcorrencia.iStatus = STATUS_TRV_OCR_BLOQUEADO
        
        lCliente = objTituloRec.lCliente
        
        'Tem que pegar o passageiro
        If iCartao = MARCADO And sTipoDoc = TRV_TIPODOC_VOU_TEXTO Then
        
            objVoucherInfo.lNumVou = lNumVou
            objVoucherInfo.sSerie = sSerie
            objVoucherInfo.sTipo = sTipoVou

            lErro = CF("TRVVoucherInfoSigav_Le", objVoucherInfo)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 192825

            'VERIFICAR SE O PASSAGEIRO ESTÁ PREENCHIDO
            If objVoucherInfo.lCliPassageiro = 0 Then
            
                'CADASTRAR O PASSAGEIRO E ACERTA A INFORMAÇÃO EM VOUCHERINFO
                lErro = TRV_VoucherInfo_Cadastra_Cliente_EmTrans(objVoucherInfo)
                If lErro <> SUCESSO Then gError 192826
    
            End If
            
            lCliente = objVoucherInfo.lCliPassageiro
        
        End If
                           
        objOcorrencia.lCliente = lCliente
        
        lErro = CF("Config_ObterAutomatico_EmTrans", "FATConfig", "NUM_PROX_TRVOCORRENCIAS", "TRVOcorrencias", "Codigo", lCodigo)
        If lErro <> SUCESSO Then gError 192311

        objOcorrencia.lCodigo = lCodigo
        objOcorrencia.lNumVou = lNumVou
        objOcorrencia.sHistorico = "Inativação com base no pagto do título " & CStr(objTituloRec.lNumTitulo) & " filial empresa " & CStr(objTituloRec.iFilialEmpresa) & " e tipo de documento " & objTituloRec.sSiglaDocumento
        objOcorrencia.sSerie = sSerie
        objOcorrencia.sTipoDoc = sTipoVou
        objOcorrencia.sObservacao = ""
        
        objOcorrenciaDet.iTipo = INATIVACAO_AUTOMATICA_TIPO_NVL_CODIGO
        objOcorrenciaDet.iSeq = 1
        objOcorrenciaDet.dValor = -dValor
        
        objOcorrencia.colDetalhamento.Add objOcorrenciaDet
        
        lErro = TRVOcorrencias_Grava_EmTrans(objOcorrencia, False)
        If lErro <> SUCESSO Then gError 192312
        
        lErro = Comando_ExecutarPos(alComando(2), "UPDATE TRVVouchers SET NumIntDocNVL = ?", alComando(1), objOcorrencia.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 192314
        
        'Grava o NVL da comissão em caso de cartão de crédito
        If iCartao = MARCADO And sTipoDoc = TRV_TIPODOC_VOU_TEXTO Then
        
            '[13:07:15] Wagner Luis Carvalhosa de Mello diz :Existe a possibilidade ter vários vouchers em uma fatura de cartão de crédito e só um ser cancelado. Se acontecer o Sigav não me dá informações suficientes para eu saber
            '[14:08:57] TRV - Daniel diz :uh!
            '[14:09:15] TRV - Daniel diz :então estive pensado nisso hoje cedo, e nunca tivemos um caso assim...
               
             'Verifica se o título a receber foi Faturado pelo Corporator
            lErro = Comando_Executar(alComando(4), "SELECT NumIntDocNFPagComi, ValorComissao FROM TitulosRecTRV WHERE NumIntDocTitRec = ?", _
                        lNumIntComi, dValorComissao, objTituloRec.lNumIntDoc)
            If lErro <> AD_SQL_SUCESSO Then gError 196270
        
            'Busca Primeiro
            lErro = Comando_BuscarPrimeiro(alComando(4))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196271
        
            If lErro = AD_SQL_SEM_DADOS Then
            
                'buscar a fatura como NFPagar aberto
                lErro = Comando_Executar(alComando(5), "SELECT NumIntDoc, ValorTotal FROM NFsPag WHERE FilialEmpresa = ? AND NumNotaFiscal = ? AND Status <> ?", _
                     lNumIntComi, dValorComissao, objTituloRec.iFilialEmpresa, objTituloRec.lNumTitulo, STATUS_EXCLUIDO)
                If lErro <> AD_SQL_SUCESSO Then gError 196272
                
                'Busca o primeiro
                lErro = Comando_BuscarProximo(alComando(5))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196273
            
                If lErro <> AD_SQL_SUCESSO Then
                
                    'buscar a fatura como titulo a Pagar baixado
                    lErro = Comando_Executar(alComando(6), "SELECT  NumIntDoc, ValorTotal FROM NFsPagBaixadas WHERE FilialEmpresa = ? AND NumNotaFiscal = ? AND Status <> ? ", _
                         lNumIntComi, dValorComissao, objTituloRec.iFilialEmpresa, objTituloRec.lNumTitulo, STATUS_EXCLUIDO)
                    If lErro <> AD_SQL_SUCESSO Then gError 196274
                    
                    'Busca o primeiro
                    lErro = Comando_BuscarProximo(alComando(6))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196275
                    
                End If
               
            End If
            
            If lErro = AD_SQL_SUCESSO Then
            
                'Le as comissões associadas a esse item de fatura que não foram geradas
                lErro = Comando_Executar(alComando(7), "SELECT C1.ValComi FROM CoinfoComissao AS C1 WHERE NumVou = ? AND Serie = ? AND TipVou = ? AND NumIntDocCorporator = ? ORDER BY C1.NumIntArq DESC, C1.NumRegArqTexto DESC", dValorComissao, lNumVou, sSerie, TRV_TIPODOC_VOU_TEXTO, lNumIntComi)
                If lErro <> AD_SQL_SUCESSO Then gError 196278
                
                'Busca a primeira
                lErro = Comando_BuscarProximo(alComando(7))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196279

                If lErro = AD_SQL_SUCESSO Then
        
                    Set objOcorrencia = New ClassTRVOcorrencias
                    Set objOcorrenciaDet = New ClassTRVOcorrenciaDet
                    
                    objOcorrencia.dtDataEmissao = gdtDataAtual
                    objOcorrencia.dValorTotal = dValorComissao
            
                    objOcorrencia.iFilialCliente = FILIAL_MATRIZ
                    objOcorrencia.iFormaPagto = FORMAPAGTO_TRV_OCR_FAT
                    objOcorrencia.iOrigem = INATIVACAO_AUTOMATICA_CODIGO
                    objOcorrencia.iStatus = STATUS_TRV_OCR_LIBERADO
                    objOcorrencia.lCliente = objVoucherInfo.lCliente
                    
                    lErro = CF("Config_ObterAutomatico_EmTrans", "FATConfig", "NUM_PROX_TRVOCORRENCIAS", "TRVOcorrencias", "Codigo", lCodigo)
                    If lErro <> SUCESSO Then gError 196276
            
                    objOcorrencia.lCodigo = lCodigo
                    objOcorrencia.lNumVou = lNumVou
                    objOcorrencia.sHistorico = "Reembolso de comissão de cartão de crédito número " & CStr(objTituloRec.lNumTitulo) & ", Filial Empresa " & CStr(objTituloRec.iFilialEmpresa) & "."
                    objOcorrencia.sSerie = sSerie
                    objOcorrencia.sTipoDoc = sTipoVou
                    objOcorrencia.sObservacao = ""
                    
                    objOcorrenciaDet.iTipo = INATIVACAO_AUTOMATICA_TIPO_NVL_CODIGO
                    objOcorrenciaDet.iSeq = 1
                    
                    objOcorrenciaDet.dValor = objOcorrencia.dValorTotal
                    
                    objOcorrencia.colDetalhamento.Add objOcorrenciaDet
            
                    lErro = TRVOcorrencias_Grava_EmTrans(objOcorrencia)
                    If lErro <> SUCESSO Then gError 196277
                    
                End If

            End If
        
        End If
    
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192313
        
    Loop
    
    'FAZ A LIBERAÇÃO DA COMISSÃO PELO CORPORATOR
    If giVersaoTRV >= TRV_VERSAO_COMISSIONAMENTO Then
    
        'Busca o histórico relacionado ao voucher para liberá-lo para o comissionamento
        lErro = Comando_Executar(alComando(8), "SELECT I.NumIntDoc, V.Data FROM TRVVouchers AS V, TRVVoucherInfo AS I WHERE V.TipoDoc = ? AND V.NumVou = I.NumVou AND V.Serie = I.Serie AND V.TipVou = I.TipVou AND V.TipoDocDestino = ? AND V.NumIntDocDestino = ? AND I.Status = ? ", _
                                                 lNumIntDocVouInfo, dtDataVou, TRV_TIPODOC_VOU_TEXTO, TRV_TIPO_DOC_DESTINO_TITREC, objTituloRec.lNumIntDoc, TRV_VOU_INFO_STATUS_BLOQUEADO)
        If lErro <> AD_SQL_SUCESSO Then gError 196360
        
        lErro = Comando_BuscarPrimeiro(alComando(8))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196361
    
        'Enquanto existirem regsitros
        Do While lErro <> AD_SQL_SEM_DADOS
        
            'Busca o histórico relacionado ao voucher para liberá-lo para o comissionamento
            lErro = Comando_ExecutarPos(alComando(9), "SELECT NumIntDoc FROM TRVVoucherInfo WHERE NumIntDoc = ? ", 0, lNumIntDocVouInfoAux, lNumIntDocVouInfo)
            If lErro <> AD_SQL_SUCESSO Then gError 196362
            
            lErro = Comando_BuscarPrimeiro(alComando(9))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196363
    
            'Trava registro contra alterações/Leituras
            lErro = Comando_LockExclusive(alComando(9))
            If lErro <> SUCESSO Then gError 196364
        
            lErro = Comando_ExecutarPos(alComando(10), "UPDATE TRVVoucherInfo SET Status = ?, TipoLiberacao = ?, NumIntDocLiberacao = ? ", alComando(9), TRV_VOU_INFO_STATUS_LIBERADO, TRV_TIPO_LIBERACAO_COMISSAO_BAIXA, objBaixaParcRec.lNumIntDoc)
            If lErro <> AD_SQL_SUCESSO Then gError 196365
            
            lErro = Comando_BuscarProximo(alComando(8))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196366
            
        Loop
    
    End If
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    BaixasParcelaRec_Grava_Customizado = SUCESSO

    Exit Function

Erro_BaixasParcelaRec_Grava_Customizado:

    BaixasParcelaRec_Grava_Customizado = gErr

    Select Case gErr

        Case 192308
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 192309, 192310, 192313
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVVOUCHERS", gErr)
            
        Case 192311, 192312, 192825, 192826
        
        Case 192314
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRVVOUCHERS", gErr)
            
        Case 196270, 196271
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOSRECTRV", gErr)
        
        Case 196272, 196273
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFSPAG", gErr)
        
        Case 196274, 196275
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFSPAGBAIXADAS", gErr)
            
        Case 196276, 196277, 192370

        Case 196278, 196279
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COINFOCOMISSAO", gErr)
            
        Case 196360 To 196363, 196366
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVVOUCHERINFO", gErr)

        Case 196364
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_TRVVOUCHERINFO", gErr)

        Case 196365
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRVVOUCHERINFO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192313)

    End Select
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function NFFatPag_Grava(ByVal objTituloPagar As ClassTituloPagar, ByVal colParcelas As colParcelaPagar, ByVal objContabil As ClassContabil) As Long
'Grava Nota Fiscal Fatura a Pagar no BD (nível + alto)

Dim lErro As Long
Dim lTransacao As Long
Dim iIndice As Integer
Dim alComando(0 To 2) As Long
Dim objAporte As New ClassTRVAportes
Dim objAportePagtoCond As New ClassTRVAportePagtoCond
Dim objAportePagtoDireto As New ClassTRVAportePagtoDireto
Dim lNumIntDocPagto As Long

On Error GoTo Erro_NFFatPag_Grava

    'Abre a transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 192368
    
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 192369
    Next
    
    lErro = CF("NFFatPag_Grava_EmTrans", objTituloPagar, colParcelas, objContabil)
    If lErro <> SUCESSO Then gError 192370
    
    If Not (objTituloPagar.objInfoUsu Is Nothing) Then
    
        Set objAporte = objTituloPagar.objInfoUsu
        
        If objAporte.lNumIntDocPagto <> 0 Then
        
            objAportePagtoDireto.iTipoDocDestino = TRV_TIPO_DOC_DESTINO_TITPAG
            objAportePagtoDireto.lNumIntDocDestino = objTituloPagar.lNumIntDoc

            lErro = CF("TRVAportePagtoDireto_Le_NumIntDocDestino", objAportePagtoDireto)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 192377
            
            lNumIntDocPagto = objAportePagtoDireto.lNumIntDoc
            
            If lErro = ERRO_LEITURA_SEM_DADOS Then
            
                objAportePagtoCond.iTipoDocDestino = TRV_TIPO_DOC_DESTINO_TITPAG
                objAportePagtoCond.lNumIntDocDestino = objTituloPagar.lNumIntDoc
            
                lErro = CF("TRVAportePagtoCond_Le_NumIntDocDestino", objAportePagtoCond)
                If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 192378
                
                lNumIntDocPagto = objAportePagtoCond.lNumIntDoc
                
            End If
            
            'Título a pagar já vinculado a outro aporte
            If lErro = SUCESSO Then
                If lNumIntDocPagto <> objAporte.lNumIntDocPagto Then gError 192379
            End If
            If objAporte.iTipoPagto = FORMAPAGTO_TRV_APORTE_TIPOPAGTO_DIRETO Then
                
                objAportePagtoDireto.lNumIntDoc = objAporte.lNumIntDocPagto
                                
                lErro = CF("TRVAportePagtoDireto_Le_NumIntDoc", objAportePagtoDireto)
                If lErro <> SUCESSO Then gError 192371
                
                If objAportePagtoDireto.lNumIntDocDestino <> 0 And objAportePagtoDireto.lNumIntDocDestino <> objTituloPagar.lNumIntDoc Then gError 192380
                                
                objAportePagtoDireto.iTipoDocDestino = TRV_TIPO_DOC_DESTINO_TITPAG
                objAportePagtoDireto.lNumIntDocDestino = objTituloPagar.lNumIntDoc
                                
                lErro = TRVAportePagtoDireto_Atualiza(alComando(0), alComando(1), objAportePagtoDireto)
                If lErro <> SUCESSO Then gError 192372
            Else
                objAportePagtoCond.lNumIntDoc = objAporte.lNumIntDocPagto
            
                lErro = CF("TRVAportePagtoCond_Le_NumIntDoc", alComando(2), objAportePagtoCond)
                If lErro <> SUCESSO Then gError 192373
            
                If objAportePagtoCond.lNumIntDocDestino <> 0 And objAportePagtoCond.lNumIntDocDestino <> objTituloPagar.lNumIntDoc Then gError 192381
                
                objAportePagtoCond.iTipoDocDestino = TRV_TIPO_DOC_DESTINO_TITPAG
                objAportePagtoCond.lNumIntDocDestino = objTituloPagar.lNumIntDoc
            
                lErro = TRVAportePagtoCond_Atualiza(alComando(0), alComando(1), objAportePagtoCond)
                If lErro <> SUCESSO Then gError 192374
            End If
            
        End If
        
    End If
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Encerra transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 192375
    
    NFFatPag_Grava = SUCESSO
    
    Exit Function
    
Erro_NFFatPag_Grava:

    NFFatPag_Grava = gErr
    
    Select Case gErr
    
        Case 192368
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 192369
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 192370 To 192374, 192377, 192378

        Case 192375
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case 192379
            Call Rotina_Erro(vbOKOnly, "ERRO_TITPAG_CAD_OUTRO_APORTE", gErr)
            
        Case 192380, 192381
            Call Rotina_Erro(vbOKOnly, "ERRO_APORTE_JA_ASSOCIADO", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192376)

    End Select
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Call Transacao_Rollback
    
    Exit Function
    
End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'GRAVAÇÃO
Public Function TRVVoucherInfoSigav_Grava(ByVal objTRVVoucherInfo As ClassTRVVoucherInfo) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_TRVVoucherInfoSigav_Grava

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 192608

    lErro = CF("TRVVoucherInfoSigav_Grava_EmTrans", objTRVVoucherInfo)
    If lErro <> SUCESSO Then gError 192609

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 192610

    TRVVoucherInfoSigav_Grava = SUCESSO

    Exit Function

Erro_TRVVoucherInfoSigav_Grava:

    TRVVoucherInfoSigav_Grava = gErr

    Select Case gErr

        Case 192608
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 192609

        Case 192610
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192618)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function TRVVoucherInfoSigav_Grava_EmTrans(ByVal objTRVVoucherInfo As ClassTRVVoucherInfo) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lNumVou As Long
Dim objVoucher As New ClassTRVVouchers
Dim objFSO As New FileSystemObject
Dim objTS As TextStream
Dim bNovo As Boolean

On Error GoTo Erro_TRVVoucherInfoSigav_Grava_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 192611
    Next

    'Le a tabelaTRVVoucherInfoSigav
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumVou FROM TRVVoucherInfoSigav WHERE Tipo= ?  AND Serie= ?  AND NumVou= ? ", _
                0, lNumVou, objTRVVoucherInfo.sTipo, objTRVVoucherInfo.sSerie, objTRVVoucherInfo.lNumVou)
    If lErro <> AD_SQL_SUCESSO Then gError 192612

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192613

    'Se existir => UPDATE, senão => INSERT
    If lErro = AD_SQL_SUCESSO Then
    
        bNovo = False

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 192614

        lErro = Comando_ExecutarPos(alComando(1), "UPDATE TRVVoucherInfoSigav SET DataEmissao= ? , DataRecepcao= ? , Cliente= ? , FornEmissor= ? , " & "Produto= ? , Destino= ? , DataInicio= ? , DataTermino= ? , Vigencia= ? , Idioma= ? , Pax= ? , DestinoVou= ? , Antc= ? , Controle= ? , Convenio= ? , DataPag= ? , Cartao= ? , Pago= ? , NumFat= ? , " & _
                "DataNasc= ? , Sexo= ? , TipoDoc= ? , CartaoFid= ? , Moeda= ? , TarifaUnitaria= ? , Cambio= ? , Valor= ? , TarifaPerc= ? , " & _
                "TarifaValorMoeda= ? , TarifaValorReal= ? , ComissaoPerc= ? , ComissaoValorMoeda= ? , ComissaoValorReal= ? , CartaoPerc= ? , CartaoValorMoeda= ? , CartaoValorReal= ? , OverPerc= ? , OverValorMoeda= ? , " & _
                "OverValorReal= ? , CMRPerc= ? , CMRValorMoeda= ? , CMRValorReal= ? , Cia= ? , Validade= ? , NumeroCC= ? , Titular= ? , ValorCartao= ? , Parcela= ? , " & _
                "Aprovacao= ? , PassageiroSobreNome= ? , PassageiroNome= ? , PassageiroCGC= ? , PassageiroEndereco= ? , PassageiroBairro= ? , PassageiroCidade= ? , PassageiroCEP= ? , PassageiroUF= ? , PassageiroEmail= ? , " & _
                "PassageiroContato= ? , PassageiroTelefone1= ?, PassageiroTelefone2= ?, Grupo= ? ", alComando(0), _
                objTRVVoucherInfo.dtDataEmissao, objTRVVoucherInfo.dtDataRecepcao, objTRVVoucherInfo.lCliente, objTRVVoucherInfo.lFornEmissor, _
                objTRVVoucherInfo.sProduto, objTRVVoucherInfo.sDestino, objTRVVoucherInfo.dtDataInicio, objTRVVoucherInfo.dtDataTermino, objTRVVoucherInfo.sVigencia, _
                objTRVVoucherInfo.sIdioma, objTRVVoucherInfo.iPax, objTRVVoucherInfo.sDestinoVou, objTRVVoucherInfo.iAntc, objTRVVoucherInfo.sControle, _
                objTRVVoucherInfo.sConvenio, objTRVVoucherInfo.dtDataPag, objTRVVoucherInfo.iCartao, objTRVVoucherInfo.iPago, objTRVVoucherInfo.lNumFat, _
                objTRVVoucherInfo.dtDataNasc, objTRVVoucherInfo.sSexo, objTRVVoucherInfo.sTipoDoc, objTRVVoucherInfo.sCartaoFid, _
                objTRVVoucherInfo.sMoeda, objTRVVoucherInfo.dTarifaUnitaria, objTRVVoucherInfo.dCambio, objTRVVoucherInfo.sValor, objTRVVoucherInfo.dTarifaPerc, _
                objTRVVoucherInfo.dTarifaValorMoeda, objTRVVoucherInfo.dTarifaValorReal, objTRVVoucherInfo.dComissaoPerc, objTRVVoucherInfo.dComissaoValorMoeda, objTRVVoucherInfo.dComissaoValorReal, _
                objTRVVoucherInfo.dCartaoPerc, objTRVVoucherInfo.dCartaoValorMoeda, objTRVVoucherInfo.dCartaoValorReal, objTRVVoucherInfo.dOverPerc, objTRVVoucherInfo.dOverValorMoeda, _
                objTRVVoucherInfo.dOverValorReal, objTRVVoucherInfo.dCMRPerc, objTRVVoucherInfo.dCMRValorMoeda, objTRVVoucherInfo.dCMRValorReal, objTRVVoucherInfo.sCia, _
                objTRVVoucherInfo.sValidade, objTRVVoucherInfo.sNumeroCC, objTRVVoucherInfo.sTitular, objTRVVoucherInfo.dValorCartao, objTRVVoucherInfo.lParcela, _
                objTRVVoucherInfo.lAprovacao, objTRVVoucherInfo.sPassageiroSobreNome, objTRVVoucherInfo.sPassageiroNome, objTRVVoucherInfo.sPassageiroCGC, objTRVVoucherInfo.sPassageiroEndereco, _
                objTRVVoucherInfo.sPassageiroBairro, objTRVVoucherInfo.sPassageiroCidade, objTRVVoucherInfo.sPassageiroCEP, objTRVVoucherInfo.sPassageiroUF, objTRVVoucherInfo.sPassageiroEmail, _
                objTRVVoucherInfo.sPassageiroContato, objTRVVoucherInfo.sPassageiroTelefone1, objTRVVoucherInfo.sPassageiroTelefone2, objTRVVoucherInfo.sGrupo)
        If lErro <> AD_SQL_SUCESSO Then gError 192615

    Else
    
        bNovo = True

        lErro = Comando_Executar(alComando(1), "INSERT INTO TRVVoucherInfoSigav( Tipo, Serie, NumVou, DataEmissao, " & _
                "DataRecepcao, Cliente, FornEmissor, Produto, Destino, " & _
                "DataInicio, DataTermino, Vigencia, Idioma, Pax, DestinoVou, Antc, Controle, Convenio, DataPag, " & _
                "Cartao, Pago, NumFat, DataNasc, Sexo, TipoDoc, CartaoFid, Moeda, TarifaUnitaria, " & _
                "Cambio, Valor, TarifaPerc, TarifaValorMoeda, TarifaValorReal, ComissaoPerc, ComissaoValorMoeda, ComissaoValorReal, CartaoPerc, CartaoValorMoeda, " & _
                "CartaoValorReal, OverPerc, OverValorMoeda, OverValorReal, CMRPerc, CMRValorMoeda, CMRValorReal, Cia, Validade, NumeroCC, " & _
                "Titular, ValorCartao, Parcela, Aprovacao, PassageiroSobreNome, PassageiroNome, PassageiroCGC, PassageiroEndereco, PassageiroBairro, PassageiroCidade, " & _
                "PassageiroCEP, PassageiroUF, PassageiroEmail, PassageiroContato, PassageiroTelefone1, PassageiroTelefone2, Grupo) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                objTRVVoucherInfo.sTipo, objTRVVoucherInfo.sSerie, objTRVVoucherInfo.lNumVou, objTRVVoucherInfo.dtDataEmissao, _
                objTRVVoucherInfo.dtDataRecepcao, objTRVVoucherInfo.lCliente, objTRVVoucherInfo.lFornEmissor, objTRVVoucherInfo.sProduto, objTRVVoucherInfo.sDestino, _
                objTRVVoucherInfo.dtDataInicio, objTRVVoucherInfo.dtDataTermino, objTRVVoucherInfo.sVigencia, objTRVVoucherInfo.sIdioma, objTRVVoucherInfo.iPax, _
                objTRVVoucherInfo.sDestinoVou, objTRVVoucherInfo.iAntc, objTRVVoucherInfo.sControle, objTRVVoucherInfo.sConvenio, objTRVVoucherInfo.dtDataPag, objTRVVoucherInfo.iCartao, objTRVVoucherInfo.iPago, objTRVVoucherInfo.lNumFat, objTRVVoucherInfo.dtDataNasc, _
                objTRVVoucherInfo.sSexo, objTRVVoucherInfo.sTipoDoc, objTRVVoucherInfo.sCartaoFid, objTRVVoucherInfo.sMoeda, objTRVVoucherInfo.dTarifaUnitaria, objTRVVoucherInfo.dCambio, objTRVVoucherInfo.sValor, objTRVVoucherInfo.dTarifaPerc, objTRVVoucherInfo.dTarifaValorMoeda, objTRVVoucherInfo.dTarifaValorReal, _
                objTRVVoucherInfo.dComissaoPerc, objTRVVoucherInfo.dComissaoValorMoeda, objTRVVoucherInfo.dComissaoValorReal, objTRVVoucherInfo.dCartaoPerc, objTRVVoucherInfo.dCartaoValorMoeda, objTRVVoucherInfo.dCartaoValorReal, objTRVVoucherInfo.dOverPerc, objTRVVoucherInfo.dOverValorMoeda, objTRVVoucherInfo.dOverValorReal, objTRVVoucherInfo.dCMRPerc, _
                objTRVVoucherInfo.dCMRValorMoeda, objTRVVoucherInfo.dCMRValorReal, objTRVVoucherInfo.sCia, objTRVVoucherInfo.sValidade, objTRVVoucherInfo.sNumeroCC, objTRVVoucherInfo.sTitular, objTRVVoucherInfo.dValorCartao, objTRVVoucherInfo.lParcela, objTRVVoucherInfo.lAprovacao, objTRVVoucherInfo.sPassageiroSobreNome, _
                objTRVVoucherInfo.sPassageiroNome, objTRVVoucherInfo.sPassageiroCGC, objTRVVoucherInfo.sPassageiroEndereco, objTRVVoucherInfo.sPassageiroBairro, objTRVVoucherInfo.sPassageiroCidade, objTRVVoucherInfo.sPassageiroCEP, objTRVVoucherInfo.sPassageiroUF, objTRVVoucherInfo.sPassageiroEmail, objTRVVoucherInfo.sPassageiroContato, objTRVVoucherInfo.sPassageiroTelefone1, objTRVVoucherInfo.sPassageiroTelefone2, objTRVVoucherInfo.sGrupo)
        If lErro <> AD_SQL_SUCESSO Then
        
            Set objTS = objFSO.OpenTextFile("C:\SGE_Log_" & Format(Now, "AAAAMMDDHHMMSS"), ForWriting, True)
        
            With objTRVVoucherInfo
                objTS.WriteLine "Tipo: " & .sTipo & " Tamanho: " & CStr(Len(.sTipo))
                objTS.WriteLine "Serie: " & .sSerie & " Tamanho: " & CStr(Len(.sSerie))
                objTS.WriteLine "NumVou: " & CStr(.lNumVou)
                objTS.WriteLine "DataEmissao: " & CStr(.dtDataEmissao)
                objTS.WriteLine "DataRecepcao: " & CStr(.dtDataRecepcao)
                objTS.WriteLine "Cliente: " & CStr(.lCliente)
                objTS.WriteLine "Emissor: " & CStr(.lFornEmissor)
                objTS.WriteLine "Produto: " & .sProduto & " Tamanho: " & CStr(Len(.sProduto))
                objTS.WriteLine "Destino: " & .sDestino & " Tamanho: " & CStr(Len(.sDestino))
                objTS.WriteLine "DataInicio: " & CStr(.dtDataInicio)
                objTS.WriteLine "DataTermino: " & CStr(.dtDataTermino)
                objTS.WriteLine "Vigencia: " & .sVigencia & " Tamanho: " & CStr(Len(.sVigencia))
                objTS.WriteLine "Idioma: " & .sIdioma & " Tamanho: " & CStr(Len(.sIdioma))
                objTS.WriteLine "Pax: " & CStr(.iPax)
                objTS.WriteLine "DestinoVou: " & .sDestinoVou & " Tamanho: " & CStr(Len(.sDestinoVou))
                objTS.WriteLine "Antc: " & CStr(.iAntc)
                objTS.WriteLine "Controle: " & .sControle & " Tamanho: " & CStr(Len(.sControle))
                objTS.WriteLine "Convenio: " & .sConvenio & " Tamanho: " & CStr(Len(.sConvenio))
                objTS.WriteLine "DataPag: " & CStr(.dtDataPag)
                objTS.WriteLine "Cartao: " & CStr(.iCartao)
                objTS.WriteLine "Pago: " & CStr(.iPago)
                objTS.WriteLine "NumFat: " & CStr(.lNumFat)
                objTS.WriteLine "CliPassageiro: " & CStr(.lCliPassageiro)
                objTS.WriteLine "DataNasc: " & CStr(.dtDataNasc)
                objTS.WriteLine "Sexo: " & .sSexo & " Tamanho: " & CStr(Len(.sSexo))
                objTS.WriteLine "TipoDoc: " & .sTipoDoc & " Tamanho: " & CStr(Len(.sTipoDoc))
                objTS.WriteLine "CartaoFid: " & .sCartaoFid & " Tamanho: " & CStr(Len(.sCartaoFid))
                objTS.WriteLine "Moeda: " & .sMoeda & " Tamanho: " & CStr(Len(.sMoeda))
                objTS.WriteLine "TarifaUnitaria: " & CStr(.dTarifaUnitaria)
                objTS.WriteLine "Cambio: " & CStr(.dCambio)
                objTS.WriteLine "Valor: " & .sValor & " Tamanho: " & CStr(Len(.sValor))
                objTS.WriteLine "TarifaPerc: " & CStr(.dTarifaPerc)
                objTS.WriteLine "TarifaValorMoeda: " & CStr(.dTarifaValorMoeda)
                objTS.WriteLine "TarifaValorReal: " & CStr(.dTarifaValorReal)
                objTS.WriteLine "ComissaoPerc: " & CStr(.dComissaoPerc)
                objTS.WriteLine "ComissaoValorMoeda: " & CStr(.dComissaoValorMoeda)
                objTS.WriteLine "ComissaoValorReal: " & CStr(.dComissaoValorReal)
                objTS.WriteLine "CartaoPerc: " & CStr(.dCartaoPerc)
                objTS.WriteLine "CartaoValorMoeda: " & CStr(.dCartaoValorMoeda)
                objTS.WriteLine "CartaoValorReal: " & CStr(.dCartaoValorReal)
                objTS.WriteLine "OverPerc: " & CStr(.dOverPerc)
                objTS.WriteLine "OverValorMoeda: " & CStr(.dOverValorMoeda)
                objTS.WriteLine "OverValorReal: " & CStr(.dOverValorReal)
                objTS.WriteLine "CMRPerc: " & CStr(.dCMRPerc)
                objTS.WriteLine "CMRValorMoeda: " & CStr(.dCMRValorMoeda)
                objTS.WriteLine "CMRValorReal: " & CStr(.dCMRValorReal)
                objTS.WriteLine "Cia: " & .sCia & " Tamanho: " & CStr(Len(.sCia))
                objTS.WriteLine "Validade: " & .sValidade & " Tamanho: " & CStr(Len(.sValidade))
                objTS.WriteLine "NumeroCC: " & .sNumeroCC & " Tamanho: " & CStr(Len(.sNumeroCC))
                objTS.WriteLine "Titular: " & .sTitular & " Tamanho: " & CStr(Len(.sTitular))
                objTS.WriteLine "ValorCartao: " & CStr(.dValorCartao)
                objTS.WriteLine "Parcela: " & CStr(.lParcela)
                objTS.WriteLine "Aprovacao: " & CStr(.lAprovacao)
                objTS.WriteLine "PassageiroSobreNome: " & .sPassageiroSobreNome & " Tamanho: " & CStr(Len(.sPassageiroSobreNome))
                objTS.WriteLine "PassageiroNome: " & .sPassageiroNome & " Tamanho: " & CStr(Len(.sPassageiroNome))
                objTS.WriteLine "PassageiroCGC: " & .sPassageiroCGC & " Tamanho: " & CStr(Len(.sPassageiroCGC))
                objTS.WriteLine "PassageiroEndereco: " & .sPassageiroEndereco & " Tamanho: " & CStr(Len(.sPassageiroEndereco))
                objTS.WriteLine "PassageiroBairro: " & .sPassageiroBairro & " Tamanho: " & CStr(Len(.sPassageiroBairro))
                objTS.WriteLine "PassageiroCidade: " & .sPassageiroCidade & " Tamanho: " & CStr(Len(.sPassageiroCidade))
                objTS.WriteLine "PassageiroCEP: " & .sPassageiroCEP & " Tamanho: " & CStr(Len(.sPassageiroCEP))
                objTS.WriteLine "PassageiroUF: " & .sPassageiroUF & " Tamanho: " & CStr(Len(.sPassageiroUF))
                objTS.WriteLine "PassageiroEmail: " & .sPassageiroEmail & " Tamanho: " & CStr(Len(.sPassageiroEmail))
                objTS.WriteLine "PassageiroContato: " & .sPassageiroContato & " Tamanho: " & CStr(Len(.sPassageiroContato))
                objTS.WriteLine "PassageiroTelefone1: " & .sPassageiroTelefone1 & " Tamanho: " & CStr(Len(.sPassageiroTelefone1))
                objTS.WriteLine "PassageiroTelefone2: " & .sPassageiroTelefone2 & " Tamanho: " & CStr(Len(.sPassageiroTelefone2))
                objTS.WriteLine "Grupo: " & .sGrupo & " Tamanho: " & CStr(Len(.sGrupo))

            End With
        
            objTS.Close
        
            gError 192616
            
        End If
        
    End If
    
    objVoucher.sSerie = objTRVVoucherInfo.sSerie
    objVoucher.sTipoDoc = TRV_TIPODOC_VOU_TEXTO
    objVoucher.sTipVou = objTRVVoucherInfo.sTipo
    objVoucher.lNumVou = objTRVVoucherInfo.lNumVou
    
    lErro = CF("TRVVouchers_Le", objVoucher)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 192671

    If lErro <> SUCESSO Then gError 194436
        
    objVoucher.sTitular = objTRVVoucherInfo.sTitular
    objVoucher.dCambio = objTRVVoucherInfo.dCambio
    objVoucher.dValorCambio = objTRVVoucherInfo.dTarifaValorMoeda
    objVoucher.sPassageiroNome = objTRVVoucherInfo.sPassageiroNome
    objVoucher.lEmissor = objTRVVoucherInfo.lFornEmissor
    objVoucher.sPassageiroSobreNome = objTRVVoucherInfo.sPassageiroSobreNome
    objVoucher.lClienteVou = objTRVVoucherInfo.lCliente
    objVoucher.lCliPassageiro = objTRVVoucherInfo.lCliPassageiro
    
    lErro = CF("TRVVouchers_Grava", objVoucher, 0, bNovo)
    If lErro <> SUCESSO Then gError 192672
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TRVVoucherInfoSigav_Grava_EmTrans = SUCESSO

    Exit Function

Erro_TRVVoucherInfoSigav_Grava_EmTrans:

    TRVVoucherInfoSigav_Grava_EmTrans = gErr

    Select Case gErr

        Case 192611
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 192612, 192613
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVVOUCHERINFOSIGAV", gErr)

        Case 192614
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRVVOUCHERINFOSIGAV", gErr)

        Case 192615
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRVVOUCHERINFOSIGAV", gErr)

        Case 192616
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRVVOUCHERINFOSIGAV", gErr)

        Case 192671, 192672
        
        Case 194436
            Call Rotina_Erro(vbOKOnly, "ERRO_TRVVOUCHERS_NAO_CADASTRADO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192617)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'EXCLUSÃO
Public Function TRVVoucherInfoSigav_Exclui(ByVal objTRVVoucherInfo As ClassTRVVoucherInfo) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_TRVVoucherInfoSigav_Exclui

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 192599

    lErro = CF("TRVVoucherInfoSigav_Exclui_EmTrans", objTRVVoucherInfo)
    If lErro <> SUCESSO Then gError 192600

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 192601

    TRVVoucherInfoSigav_Exclui = SUCESSO

    Exit Function

Erro_TRVVoucherInfoSigav_Exclui:

    TRVVoucherInfoSigav_Exclui = gErr

    Select Case gErr

        Case 192599
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 192600

        Case 192601
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192620)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function TRVVoucherInfoSigav_Exclui_EmTrans(ByVal objTRVVoucherInfo As ClassTRVVoucherInfo) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim iAux As Integer

On Error GoTo Erro_TRVVoucherInfoSigav_Exclui_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 192602
    Next

    'Le a tabelaTRVVoucherInfoSigav
    lErro = Comando_ExecutarPos(alComando(0), "SELECT 1 FROM TRVVoucherInfoSigav WHERE Tipo= ?  AND Serie= ?  AND NumVou= ? ", _
                0, iAux, objTRVVoucherInfo.sTipo, objTRVVoucherInfo.sSerie, objTRVVoucherInfo.lNumVou)
    If lErro <> AD_SQL_SUCESSO Then gError 192603

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192604

    'Se não existir => ERRO
    If lErro = AD_SQL_SEM_DADOS Then gError 192605

    'Trava registro contra alterações/Leituras
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> SUCESSO Then gError 192606

    lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM TRVVoucherInfoSigav", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 192607

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TRVVoucherInfoSigav_Exclui_EmTrans = SUCESSO

    Exit Function

Erro_TRVVoucherInfoSigav_Exclui_EmTrans:

    TRVVoucherInfoSigav_Exclui_EmTrans = gErr

    Select Case gErr

        Case 192602
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 192603, 192604
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVVOUCHERINFOSIGAV", gErr)

        Case 192605
            Call Rotina_Erro(vbOKOnly, "ERRO_TRVVOUCHERINFOSIGAV_NAO_CADASTRADO", gErr, objTRVVoucherInfo.sTipo, objTRVVoucherInfo.sSerie, objTRVVoucherInfo.lNumVou)

        Case 192606
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRVVOUCHERINFOSIGAV", gErr)

        Case 192607
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRVVOUCHERINFOSIGAV", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192619)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function TRVVoucher_Exclui(ByVal objTRVVoucher As ClassTRVVouchers) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_TRVVoucher_Exclui

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 192646

    lErro = CF("TRVVouchers_Cancela", objTRVVoucher)
    If lErro <> SUCESSO Then gError 192647

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 192648

    TRVVoucher_Exclui = SUCESSO

    Exit Function

Erro_TRVVoucher_Exclui:

    TRVVoucher_Exclui = gErr

    Select Case gErr

        Case 192646
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 192647

        Case 192648
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192649)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function TRV_VoucherInfo_Cadastra_Cliente(ByVal objTRVVoucherInfo As ClassTRVVoucherInfo) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_TRV_VoucherInfo_Cadastra_Cliente

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 192740

    lErro = TRV_VoucherInfo_Cadastra_Cliente_EmTrans(objTRVVoucherInfo)
    If lErro <> SUCESSO Then gError 192741

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 192742

    TRV_VoucherInfo_Cadastra_Cliente = SUCESSO

    Exit Function

Erro_TRV_VoucherInfo_Cadastra_Cliente:

    TRV_VoucherInfo_Cadastra_Cliente = gErr

    Select Case gErr

        Case 192740
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 192741

        Case 192742
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192743)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function TRV_VoucherInfo_Cadastra_Cliente_EmTrans(ByVal objTRVVoucherInfo As ClassTRVVoucherInfo) As Long

Dim lErro As Long
Dim alComando(0 To 5) As Long
Dim iIndice As Integer
Dim colEndereco As New Collection
Dim colEnderecos As New colEndereco
Dim objEndereco As ClassEndereco
Dim objCliente As New ClassCliente
Dim lCodigo As Long
Dim lCliente As Long
Dim iFilial As Integer
Dim objCidades As ClassCidades
Dim objFilialCliente As New ClassFilialCliente
Dim colCategoriaItem As New Collection
Dim lNumVou As Long
Dim iCodigoCidade As Integer
Dim sNumero As String
Dim iTipo As Integer
Dim objFilialCliCategoria As ClassFilialCliCategoria
Dim objClienteTRV As ClassClienteTRV
Dim iFilialEmpresa As Integer
Dim sGrupo As String

On Error GoTo Erro_TRV_VoucherInfo_Cadastra_Cliente_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 192744
    Next
    
    lErro = Comando_Executar(alComando(0), "SELECT F.CodCliente, F.CodFilial, C.Tipo FROM FiliaisClientes AS F, Clientes AS C WHERE F.CodCliente = C.Codigo AND F.CGC = ? AND F.CGC <> ?", lCliente, iFilial, iTipo, objTRVVoucherInfo.sPassageiroCGC, "")
    If lErro <> AD_SQL_SUCESSO Then gError 192745

    lErro = Comando_BuscarProximo(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192746
    
    If lErro <> AD_SQL_SUCESSO Then
    
        lErro = CF("Config_ObterAutomatico_EmTrans", "TRVConfig", TRVCONFIG_PROX_NUM_PASSAGEIRO, "Clientes", "Codigo", lCodigo)
        If lErro <> SUCESSO Then gError 192747
    
        For iIndice = 1 To 3
            Set objEndereco = New ClassEndereco
            colEndereco.Add objEndereco
        Next
        
        objCliente.lCodigo = lCodigo
        objCliente.iAtivo = MARCADO

        sNumero = ""
        Call Formata_String_Numero(objTRVVoucherInfo.sPassageiroCGC, sNumero)
        objCliente.sCgc = sNumero
        
        objCliente.iTipo = TRV_TIPO_CLIENTE_PASSAGEIRO
        objCliente.dtDataPrimeiraCompra = DATA_NULA
        objCliente.dtDataUltChequeDevolvido = DATA_NULA
        objCliente.dtDataUltimaCompra = DATA_NULA
        objCliente.dtDataUltimoProtesto = DATA_NULA
        objCliente.dtDataUltVisita = DATA_NULA
        objCliente.iCondicaoPagto = 1
            
        Set objFilialCliCategoria = New ClassFilialCliCategoria
        Set objClienteTRV = New ClassClienteTRV
        
        objFilialCliCategoria.sCategoria = TRV_CATEGORIA_CONDFAT
        objFilialCliCategoria.sItem = TRV_CATEGORIA_CONDFAT_PASSAGEIRO
        objFilialCliCategoria.iFilial = FILIAL_MATRIZ
        objFilialCliCategoria.lCliente = objCliente.lCodigo
        
        objCliente.colCategoriaItem.Add objFilialCliCategoria
        
        Set objCliente.objInfoUsu = objClienteTRV
        
        sGrupo = String(STRING_TRV_OCR_CODGRUPO, 0)
        
        lErro = Comando_Executar(alComando(3), "SELECT FilialEmpresa, Grupo FROM FiliaisClientesTRV WHERE CodCliente = ? AND CodFilial = ?", iFilialEmpresa, sGrupo, objTRVVoucherInfo.lCliente, FILIAL_MATRIZ)
        If lErro <> AD_SQL_SUCESSO Then gError 192745
    
        lErro = Comando_BuscarProximo(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192746
        
        If lErro <> AD_SQL_SEM_DADOS Then
        
            objClienteTRV.iFilialCoinfo = Filial_Corporator_Retorna_Coinfo(iFilialEmpresa)
            objClienteTRV.iFilialEmpresa = iFilialEmpresa
            objClienteTRV.sGrupo = sGrupo
            
        End If
        
        objClienteTRV.lEmpresaPai = objCliente.lCodigo
               
    Else
    
        If lCliente < FATOR_PAX_CLIENTE And iTipo <> 10 And iTipo <> 12 Then gError 196543
    
        objCliente.lCodigo = lCliente
        
        lErro = CF("Cliente_Le", objCliente)
        If lErro <> SUCESSO And lErro <> 12293 Then gError 192748
        
        lErro = CF("Enderecos_Le_Cliente", colEnderecos, objCliente)
        If lErro <> SUCESSO Then gError 192749

        For Each objEndereco In colEnderecos
            colEndereco.Add objEndereco
        Next
        
        objFilialCliente.lCodCliente = lCliente
        objFilialCliente.iCodFilial = FILIAL_MATRIZ
        
        'Lê as Categoria da Filial do Cliente
        lErro = CF("FilialCliente_Le_Categorias", objFilialCliente, colCategoriaItem)
        If lErro <> SUCESSO Then gError 192750

        Set objCliente.colCategoriaItem = colCategoriaItem
        
        lErro = CF("Cliente_Le_Customizado", objCliente)
        If lErro <> SUCESSO Then gError 192750
    
    End If
    
    If Len(Trim(objTRVVoucherInfo.sPassageiroNome)) = 0 Then gError 192829
    
    objCliente.sNomeReduzido = Left(objTRVVoucherInfo.sPassageiroNome, STRING_CLIENTE_NOME_REDUZIDO - Len(SEPARADOR & CStr(objCliente.lCodigo))) & SEPARADOR & CStr(objCliente.lCodigo)
    objCliente.sRazaoSocial = objTRVVoucherInfo.sPassageiroNome & " \ " & objTRVVoucherInfo.sPassageiroSobreNome
    
    For Each objEndereco In colEndereco

        objEndereco.iCodigoPais = PAIS_BRASIL
        objEndereco.sBairro = objTRVVoucherInfo.sPassageiroBairro
        
        sNumero = ""
        Call Formata_String_Numero(objTRVVoucherInfo.sPassageiroCEP, sNumero)
        If sNumero = "00000000" Then sNumero = ""
        objEndereco.sCEP = sNumero
        objEndereco.sCidade = objTRVVoucherInfo.sPassageiroCidade
        objEndereco.sContato = objTRVVoucherInfo.sPassageiroContato
        objEndereco.sEmail = objTRVVoucherInfo.sPassageiroEmail
        objEndereco.sEndereco = objTRVVoucherInfo.sPassageiroEndereco
        If Len(Trim(objTRVVoucherInfo.sPassageiroUF)) > 0 Then
            objEndereco.sSiglaEstado = objTRVVoucherInfo.sPassageiroUF
        Else
            objEndereco.sSiglaEstado = "SP"
        End If
        objEndereco.sTelefone1 = objTRVVoucherInfo.sPassageiroTelefone1
        objEndereco.sTelefone2 = objTRVVoucherInfo.sPassageiroTelefone2
        
        If Len(Trim(objEndereco.sSiglaEstado)) = 0 Then gError 192828
        
    Next
    
    'Grava as Cidades que não estiverem cadastradas
    For Each objEndereco In colEndereco
    
        Set objCidades = New ClassCidades
        
        objCidades.sDescricao = objEndereco.sCidade
        
        lErro = CF("Cidade_Le_Nome", objCidades)
        If lErro <> SUCESSO And lErro <> ERRO_OBJETO_NAO_CADASTRADO Then gError 192751
        
        If lErro <> SUCESSO Then
        
            lErro = CF("Config_Obter_Inteiro_Automatico_EmTrans", "FATConfig", "NUM_PROX_CIDADECADASTRO", "Cidades", "Codigo", iCodigoCidade)
            If lErro <> SUCESSO Then gError 192752

            objCidades.iCodigo = iCodigoCidade
            objCidades.sDescricao = objEndereco.sCidade
                
            lErro = CF("Cidade_Grava_EmTrans", objCidades)
            If lErro <> SUCESSO Then gError 192753
            
        End If
    
    Next

    lErro = CF("Cliente_Grava_EmTrans", objCliente, colEndereco)
    If lErro <> SUCESSO Then gError 192754
    
    objTRVVoucherInfo.lCliPassageiro = objCliente.lCodigo

    lErro = Comando_ExecutarPos(alComando(1), "SELECT NumVou FROM TRVVoucherInfoSigav WHERE Tipo= ?  AND Serie= ?  AND NumVou= ? ", _
                0, lNumVou, objTRVVoucherInfo.sTipo, objTRVVoucherInfo.sSerie, objTRVVoucherInfo.lNumVou)
    If lErro <> AD_SQL_SUCESSO Then gError 192755

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192756

    If lErro <> AD_SQL_SUCESSO Then gError 192757

    'Trava registro contra alterações/Leituras
    lErro = Comando_LockExclusive(alComando(1))
    If lErro <> AD_SQL_SUCESSO Then gError 192758

    lErro = Comando_ExecutarPos(alComando(2), "UPDATE TRVVoucherInfoSigav SET CliPassageiro= ? ", alComando(1), objCliente.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 192759
    
    lErro = Comando_ExecutarPos(alComando(4), "SELECT NumVou FROM TRVVouchers WHERE NumVou = ? AND TipoDoc = ? AND Serie = ? AND TipVou = ? ", _
                0, lNumVou, objTRVVoucherInfo.lNumVou, TRV_TIPODOC_VOU_TEXTO, objTRVVoucherInfo.sSerie, objTRVVoucherInfo.sTipo)
    If lErro <> AD_SQL_SUCESSO Then gError 192755

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(4))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192756

    If lErro <> AD_SQL_SUCESSO Then gError 192757

    'Trava registro contra alterações/Leituras
    lErro = Comando_LockExclusive(alComando(4))
    If lErro <> AD_SQL_SUCESSO Then gError 192758

    lErro = Comando_ExecutarPos(alComando(5), "UPDATE TRVVouchers SET CliPassageiro= ? ", alComando(4), objCliente.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 192759

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TRV_VoucherInfo_Cadastra_Cliente_EmTrans = SUCESSO

    Exit Function

Erro_TRV_VoucherInfo_Cadastra_Cliente_EmTrans:

    TRV_VoucherInfo_Cadastra_Cliente_EmTrans = gErr

    Select Case gErr

        Case 192744
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 192745, 192746
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLIENTES", gErr)
            
        Case 192747 To 192754
        
        Case 192755, 192756
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVVOUCHERS", gErr)
            
        Case 192757
            Call Rotina_Erro(vbOKOnly, "ERRO_VOUCHER_NAO_CADASTRADO", gErr, objTRVVoucherInfo.lNumVou, objTRVVoucherInfo.sSerie, objTRVVoucherInfo.sTipo)

        Case 192758
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRVVOUCHERS", gErr)

        Case 192759
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRVVOUCHERS", gErr)
            
        Case 192828
            Call Rotina_Erro(vbOKOnly, "ERRO_ESTADO_NAO_CADASTRADO", gErr, objEndereco.sSiglaEstado)

        Case 192829
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_REDUZIDO_NAO_PREENCHIDO", gErr)
            
        Case 196543
            Call Rotina_Erro(vbOKOnly, "ERRO_CNPJ_PERTENCE_A_UMA_AGENGIA", gErr, objTRVVoucherInfo.sPassageiroCGC, objTRVVoucherInfo.lNumVou, lCliente)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192617)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'GRAVAÇÃO
Public Function TitulosRecTRV_Grava(ByVal objTitulosRecTRV As ClassTitulosRecTRV) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Integer

On Error GoTo Erro_TitulosRecTRV_Grava

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 192776
    Next

    'Le a tabelaTitulosRecTRV
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDocTitRec FROM TitulosRecTRV WHERE NumIntDocTitRec= ? ", _
                0, lNumIntDoc, objTitulosRecTRV.lNumIntDocTitRec)
    If lErro <> AD_SQL_SUCESSO Then gError 192777

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192778

    'Se existir => UPDATE, senão => INSERT
    If lErro = AD_SQL_SUCESSO Then

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 192779

        lErro = Comando_ExecutarPos(alComando(1), "UPDATE TitulosRecTRV SET NumIntDocNFPagComi= ? , ValorComissao= ? ", alComando(0), _
                objTitulosRecTRV.lNumIntDocNFPagComi, objTitulosRecTRV.dValorComissao)
        If lErro <> AD_SQL_SUCESSO Then gError 192780

    Else


        lErro = Comando_Executar(alComando(1), "INSERT INTO TitulosRecTRV( NumIntDocTitRec, ValorTarifa, NumIntDocNFPagComi, ValorDeducoes, " & _
                "ValorComissao, ValorBruto) VALUES (?,?,?,?,?,?)", _
                objTitulosRecTRV.lNumIntDocTitRec, objTitulosRecTRV.dValorTarifa, objTitulosRecTRV.lNumIntDocNFPagComi, objTitulosRecTRV.dValorDeducoes, _
                objTitulosRecTRV.dValorComissao, objTitulosRecTRV.dValorBruto)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192781

    End If

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TitulosRecTRV_Grava = SUCESSO

    Exit Function

Erro_TitulosRecTRV_Grava:

    TitulosRecTRV_Grava = gErr

    Select Case gErr

        Case 192776
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 192777, 192778
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOSRECTRV", gErr)

        Case 192779
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TITULOSRECTRV", gErr)

        Case 192780
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TITULOSRECTRV", gErr)

        Case 192781
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TITULOSRECTRV", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192775)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function


Public Function TitulosRecTRV_Exclui(ByVal lNumIntDoc As Long, ByVal objFaturaTRV As ClassFaturaTRV) As Long

Dim lErro As Long
Dim alComando(0 To 2) As Long
Dim iIndice As Integer
Dim lNumIntDocComi As Long
Dim objNFPag As New ClassNFsPag
Dim objContabil As New ClassContabil
Dim lNumNF As Long

On Error GoTo Erro_TitulosRecTRV_Exclui

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 192767
    Next

    'Le a tabelaTitulosRecTRV
    lErro = Comando_Executar(alComando(0), "SELECT N.NumNotaFiscal FROM TRVItemNotaTitulo AS T, NFiscal AS N, ItensNFiscal AS I WHERE T.NumIntDocTitulo = ? AND T.NumIntDocItemNF = I.NumIntDoc AND I.NumIntNF = N.NumIntDoc AND N.Status NOT IN (5,7) ", _
                lNumNF, lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 192768

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192769
    
    'O título está associado a NF
    If lErro = AD_SQL_SUCESSO Then gError 196455
    
    'Le a tabelaTitulosRecTRV
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDocNFPagComi FROM TitulosRecTRV WHERE NumIntDocTitRec= ? ", _
                0, lNumIntDocComi, lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 192768

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192769

    If lErro <> AD_SQL_SEM_DADOS Then

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 192771
    
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM TitulosRecTRV", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 192772
        
        If lNumIntDocComi <> 0 Then
     
            objNFPag.lNumIntDoc = lNumIntDocComi
            
            Set objNFPag.objInfoUsu = objFaturaTRV
     
            'Le a nota de comissão
            lErro = CF("NFPag_Le", objNFPag)
            If lErro <> SUCESSO And lErro <> 18257 Then gError 192782
            
            'Se ela ainda não foi faturada
            If lErro = SUCESSO And objNFPag.lNumIntTitPag = 0 Then

                Call objContabil.Contabil_Inicializa_Contabilidade4(14, MODULO_CONTASAPAGAR)
            
                objNFPag.iExclusao = DESMARCADO
            
                lErro = CF("NFPag_Exclui_EmTransacao", objNFPag, objContabil)
                If lErro <> SUCESSO Then gError 192783
                
                Set objContabil = Nothing
                
            End If
                
        End If
        
    End If

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TitulosRecTRV_Exclui = SUCESSO

    Exit Function

Erro_TitulosRecTRV_Exclui:

    TitulosRecTRV_Exclui = gErr

    Select Case gErr

        Case 192767
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 192768, 192769
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOSRECTRV", gErr)

        Case 192771
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TITULOSRECTRV", gErr)

        Case 192772
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TITULOSRECTRV", gErr)
            
        Case 192782, 192783
        
        Case 196455
            Call Rotina_Erro(vbOKOnly, "ERRO_TITREC_ASSOC_NF", gErr, lNumNF)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192770)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function TRVFaturas_Regera_Html(ByVal lFaturaDe As Long, ByVal lFaturaAte As Long, ByVal sModelo As String, ByVal sDiretorio As String, ByVal sSiglaDoc As String) As Long

Dim lErro As Long
Dim alComando(0 To 5) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim objFaturamento As ClassTRVFaturamento
Dim objTitRec As ClassTituloReceber
Dim objInfoFatura As ClassTRVFATInfoFatura
Dim objInfoVoucher As ClassTRVFATInfoVoucher
Dim objCliente As ClassCliente
Dim objFilialEmpresa As AdmFiliais
Dim objEndereco As ClassEndereco
Dim objFilialCliente As ClassFilialCliente
Dim sSerie As String
Dim sTipVou As String
Dim sTipoDoc As String
Dim lNumVou As Long
Dim lNumFat As Long
Dim dtDataVencimento As Date
Dim dValorAporte As Double
Dim dValorComissao As Double
Dim objVoucher As ClassTRVVouchers
Dim lNumIntDocAporte As Long
Dim lNumIntDocAportePagto As Long
Dim iSeq As Integer
Dim iParcelaAporte As Integer
Dim lCodigoAporte As Long
Dim sArquivo As String
Dim sConteudoHtml As String
Dim dValorTarifa As Double
Dim dValorDeducoes As Double
Dim sSiglaDocumento As String
Dim objTRVVoucherInfo As ClassTRVVoucherInfo
Dim iTipoDocDestino As Integer
Dim iFilialEmpresa As Integer
Dim lCliente As Long
Dim dValor As Double
Dim dtDataEmissao As Date
Dim iCondicaoPagto As Integer
Dim sModeloCP As String
Dim dValorItem As Double
Dim dtDataItem As Date
Dim objForn As ClassFornecedor
Dim dValorItens As Double
Dim dValorCreditos As Double
Dim dValorDebitos As Double
Dim iLinha As Integer
Dim sNomeCliente As String
Dim objVouAporte As ClassTRVVoucherAporte
Dim objOCRAporte As ClassTRVOcorrenciaAporte
Dim iTipoPagtoAporte As Integer
Dim lNumIntDocOCR As Long
Dim dValorAporteCred As Double
Dim lNumIntDocCredito As Long
Dim lCodigoAporteCred As Long
Dim objFaturaAporte As ClassTRVFaturaAporte
Dim iNumParcelas As Integer
Dim iPagtoParcial As Integer

On Error GoTo Erro_TRVFaturas_Regera_Html

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 192830
    Next
    
    lErro = CF("TRVConfig_Le", TRVCONFIG_DIRETORIO_MODELO_FAT_HTML_CARTAO, EMPRESA_TODA, sModeloCP)
    If lErro <> SUCESSO Then gError 192838
    
    sSiglaDocumento = String(STRING_SIGLA_DOCUMENTO, 0)

    'Le a tabela TitulosRecTRV
    lErro = Comando_Executar(alComando(0), "SELECT NumIntDoc,TipoDocDestino, NumTitulo,FilialEmpresa, SiglaDocumento, Cliente, Valor, DataEmissao, DataVencimento, CondicaoPagto FROM TRVFaturas WHERE NumTitulo BETWEEN ? AND ? AND (SiglaDocumento = ? OR ? = '') AND Status <> 5 ORDER BY SiglaDocumento ASC", _
                lNumIntDoc, iTipoDocDestino, lNumFat, iFilialEmpresa, sSiglaDocumento, lCliente, dValor, dtDataEmissao, dtDataVencimento, iCondicaoPagto, lFaturaDe, lFaturaAte, sSiglaDoc, sSiglaDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 192831

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192832

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objFaturamento = New ClassTRVFaturamento
        Set objTitRec = New ClassTituloReceber
        Set objInfoFatura = New ClassTRVFATInfoFatura
        objFaturamento.colInfoFaturas.Add objInfoFatura
        
        objFaturamento.iIndiceFaturaAtual = 1
        
        Select Case iTipoDocDestino
        
            Case TRV_TIPO_DOC_DESTINO_TITREC
        
                'Le a tabela TRVTitulosEXP
                lErro = Comando_Executar(alComando(2), "SELECT A.ValorAporte, A.NumIntDocCredito, A.CodigoAporte, A.ParcelaAporte, A.NumParcelas, A.PagtoParcial FROM TRVTitulosExp AS E, TRVTitulosExpAporte AS A WHERE E.NumIntDoc = A.NumIntDocTitExp AND E.TipoDocOrigem = ? AND E.NumIntDocOrigem = ? AND E.Excluido = ? ", dValorAporteCred, lNumIntDocCredito, lCodigoAporteCred, iParcelaAporte, iNumParcelas, iPagtoParcial, iTipoDocDestino, lNumIntDoc, DESMARCADO)
                If lErro <> AD_SQL_SUCESSO Then gError 192834

                'Busca Primeiro
                lErro = Comando_BuscarPrimeiro(alComando(2))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192835
                                   
                Do While lErro = AD_SQL_SUCESSO
                
                    Set objFaturaAporte = New ClassTRVFaturaAporte
                
                    objInfoFatura.dValorAporteCred = objInfoFatura.dValorAporteCred + dValorAporteCred
                    objInfoFatura.colFatAporteCred.Add objFaturaAporte
                
                    objFaturaAporte.lCodigoAporte = lCodigoAporteCred
                    objFaturaAporte.lNumIntDocCredito = lNumIntDocCredito
                    objFaturaAporte.iParcelaAporte = iParcelaAporte
                    objFaturaAporte.iNumParcelas = iNumParcelas
                    objFaturaAporte.iPagtoParcial = iPagtoParcial
                    objFaturaAporte.dValorAporte = dValorAporteCred
                    
                    'Busca Primeiro
                    lErro = Comando_BuscarPrimeiro(alComando(2))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192835
                    
                Loop
                
                'Le a tabela TitulosRecTRV
                lErro = Comando_Executar(alComando(3), "SELECT ValorTarifa, ValorDeducoes FROM TitulosRecTrv WHERE NumIntDocTitRec = ? ", dValorTarifa, dValorDeducoes, lNumIntDoc)
                If lErro <> AD_SQL_SUCESSO Then gError 192834
            
                'Busca Primeiro
                lErro = Comando_BuscarPrimeiro(alComando(3))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192835
                
                If lErro = AD_SQL_SUCESSO Then
                    objInfoFatura.dValorDeducoes = dValorDeducoes
                    objInfoFatura.dValorTarifa = dValorTarifa
                End If
                
                If sSiglaDocumento = TIPODOC_CARTAO_CRED_DEB Then
                    objInfoFatura.iCartao = MARCADO
                End If
                
            Case TRV_TIPO_DOC_DESTINO_TITPAG, TRV_TIPO_DOC_DESTINO_NFSPAG
            
                dValor = -dValor
                
        End Select
        
        objInfoFatura.dtDataEmissao = dtDataEmissao
        objInfoFatura.dtDataVencimento = dtDataVencimento
        objInfoFatura.dValor = dValor
        objInfoFatura.iCondPagto = iCondicaoPagto
        objInfoFatura.iFilialEmpresa = iFilialEmpresa
        objInfoFatura.iLinha = 1
        objInfoFatura.iLinhaFatCliente = 1
        objInfoFatura.iLinhaFatEnderecoCli = 1
        objInfoFatura.iLinhaFatEnderecoEmp = 2
        objInfoFatura.iLinhaFatFilial = 1
        objInfoFatura.iLinhaFatFilialCliente = 1
        objInfoFatura.iMarcado = MARCADO
        objInfoFatura.lCliente = lCliente
        objInfoFatura.lFatura = 1
        objInfoFatura.lNumeroFaturaReal = lNumFat
        
        sSerie = String(STRING_TRV_OCR_SERIE, 0)
        sTipVou = String(STRING_TRV_OCR_TIPOVOU, 0)
        sTipoDoc = String(STRING_TRV_OCR_TIPODOC, 0)
        sNomeCliente = String(STRING_CLIENTE_NOME_REDUZIDO, 0)
        
        'Le a tabelaTitulosRecTRV
        lErro = Comando_Executar(alComando(1), "SELECT NumIntDoc, Serie, TipVou, TipoDoc,NomeCliente, NumVou, ValorAporte, ValorComissao, NumIntDocPagtoAporteFat, ParcelaAporte, CodigoAporte, Valor, Data FROM DocFaturados WHERE NumTitulo = ? ", _
                    lNumIntDocOCR, sSerie, sTipVou, sTipoDoc, sNomeCliente, lNumVou, dValorAporte, dValorComissao, lNumIntDocAportePagto, iParcelaAporte, lCodigoAporte, dValorItem, dtDataItem, lNumFat)
        If lErro <> AD_SQL_SUCESSO Then gError 192836
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192837

        iSeq = 0
        Do While lErro <> AD_SQL_SEM_DADOS
        
            Set objInfoVoucher = New ClassTRVFATInfoVoucher
            objFaturamento.colInfoVouchers.Add objInfoVoucher
            Set objVoucher = New ClassTRVVouchers
            Set objTRVVoucherInfo = New ClassTRVVoucherInfo
            
            iSeq = iSeq + 1
            
            objVoucher.sSerie = sSerie
            objVoucher.sTipVou = sTipVou
            objVoucher.sTipoDoc = TRV_TIPODOC_VOU_TEXTO
            objVoucher.lNumVou = lNumVou
            
            lErro = CF("TRVVouchers_Le", objVoucher)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 192838
            
            objInfoVoucher.dCambio = objVoucher.dCambio
            objInfoVoucher.iFilialEmpresa = iFilialEmpresa
            objInfoVoucher.iLinha = iSeq
            objInfoVoucher.iMarcado = MARCADO
            objInfoVoucher.iPax = objVoucher.iPax
            objInfoVoucher.lCliente = lCliente
            objInfoVoucher.sNome = sNomeCliente
            objInfoVoucher.lNumIntDoc = lNumIntDocOCR
            
            If sTipoDoc = TRV_TIPODOC_VOU_TEXTO Then
                objInfoVoucher.dValorCambio = objVoucher.dValorCambio
                objInfoVoucher.lCodigoAporte = lCodigoAporte
                objInfoVoucher.lNumIntDocAportePagto = lNumIntDocAportePagto
                objInfoVoucher.dValorComissao = dValorComissao
                objInfoVoucher.dValorBruto = dValorComissao + objVoucher.dValor
                objInfoVoucher.dtDataEmissao = objVoucher.dtData
                objInfoVoucher.dValor = objVoucher.dValor
                objInfoVoucher.dValorAporte = dValorAporte
            Else
                objInfoVoucher.dValorCambio = 0
                objInfoVoucher.lCodigoAporte = 0
                objInfoVoucher.lNumIntDocAportePagto = 0
                objInfoVoucher.dValorComissao = 0
                objInfoVoucher.dValorBruto = dValorItem
                objInfoVoucher.dtDataEmissao = dtDataItem
                objInfoVoucher.dValor = dValorItem
                objInfoVoucher.dValorAporte = 0
            End If
            
            objInfoVoucher.lFatura = 1
            objInfoVoucher.lNumVou = objVoucher.lNumVou
            objInfoVoucher.sCodGrupo = objVoucher.sCodGrupo
            objInfoVoucher.sControle = objVoucher.sControle
            objInfoVoucher.sMoeda = objVoucher.sMoeda
            objInfoVoucher.sProduto = objVoucher.sProduto
            objInfoVoucher.sSerie = sSerie
            objInfoVoucher.sTipoDoc = sTipoDoc
            objInfoVoucher.sTipoVou = sTipVou
            
            If objInfoVoucher.sTipoDoc = TRV_TIPODOC_CMC_TEXTO Or objInfoVoucher.sTipoDoc = TRV_TIPODOC_CMR_TEXTO Then
                objInfoVoucher.sControle = ""
            End If

            If objInfoVoucher.sTipoDoc = TRV_TIPODOC_NVL_TEXTO Then
                objInfoVoucher.dValorAporte = dValorAporte
                objInfoVoucher.lCodigoAporte = lCodigoAporte
                objInfoVoucher.iParcelaAporte = iParcelaAporte
                objInfoVoucher.lNumIntDocAportePagto = lNumIntDocAportePagto
            End If
            
            objInfoVoucher.lCliVou = objInfoVoucher.lCliente
            objInfoVoucher.sNomeCliVou = objInfoVoucher.sNome
            
            objTRVVoucherInfo.lNumVou = objInfoVoucher.lNumVou
            objTRVVoucherInfo.sTipo = objInfoVoucher.sTipoVou
            objTRVVoucherInfo.sSerie = objInfoVoucher.sSerie
                        
            lErro = CF("TRVVoucherInfoSigav_Le", objTRVVoucherInfo)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 194210
            
            If lErro = SUCESSO Then
                            
                objInfoVoucher.sTitular = objTRVVoucherInfo.sPassageiroNome & " " & objTRVVoucherInfo.sPassageiroSobreNome
                            
                If objTRVVoucherInfo.lFornEmissor <> 0 Then

                    Set objForn = New ClassFornecedor

                    objForn.lCodigo = objTRVVoucherInfo.lFornEmissor
                    
                    lErro = CF("Fornecedor_Le", objForn)
                    If lErro <> SUCESSO And lErro <> 12729 Then gError 194210
                    
                    objInfoVoucher.sEmissor = objForn.sRazaoSocial
                
                End If
            
                If objInfoVoucher.sTipoDoc = TRV_TIPODOC_CMC_TEXTO Or objInfoVoucher.sTipoDoc = TRV_TIPODOC_CMR_TEXTO Or (objVoucher.iCartao = MARCADO And objInfoVoucher.sTipoDoc = TRV_TIPODOC_VOU_TEXTO) Or ((objInfoVoucher.sTipoDoc = TRV_TIPODOC_NVL_TEXTO Or objInfoVoucher.sTipoDoc = TRV_TIPODOC_OCR_TEXTO) And objInfoVoucher.lCliente <> objTRVVoucherInfo.lCliente) Then
                    
                    If objTRVVoucherInfo.lCliente <> 0 Then
                    
                        Set objCliente = New ClassCliente
            
                        objCliente.lCodigo = objTRVVoucherInfo.lCliente
                
                        lErro = CF("Cliente_Le", objCliente)
                        If lErro <> SUCESSO And lErro <> 12293 Then gError 192840
                        
                        objInfoVoucher.lCliVou = objTRVVoucherInfo.lCliente
                        objInfoVoucher.sNomeCliVou = objCliente.sRazaoSocial
                        
                        If objInfoVoucher.sTipoDoc = TRV_TIPODOC_CMC_TEXTO Or objInfoVoucher.sTipoDoc = TRV_TIPODOC_CMR_TEXTO Or ((objInfoVoucher.sTipoDoc = TRV_TIPODOC_NVL_TEXTO Or objInfoVoucher.sTipoDoc = TRV_TIPODOC_OCR_TEXTO) And objInfoVoucher.lCliente <> objTRVVoucherInfo.lCliente) Then
                            'Se for comissão deve sair o cliente do voucher como titular
                            objInfoVoucher.sTitular = objCliente.sNomeReduzido
                        End If
                        
                    End If
                    
                End If
        
            End If
            
            'Se for voucher pago com cartão o titular é o titular do cartão
            If objVoucher.iCartao = MARCADO And objInfoVoucher.sTipoDoc = TRV_TIPODOC_VOU_TEXTO Then
                objInfoVoucher.sTitular = objVoucher.sTitular
            End If
            
            objInfoFatura.iParcelaAporte = iParcelaAporte

            objInfoFatura.colInfoVouchers.Add objInfoVoucher
            
            If sTipoDoc = TRV_TIPODOC_VOU_TEXTO Then
                       
                lErro = Comando_ExecutarPos(alComando(4), "SELECT TipoPagtoAporte, NumIntDocPagtoAporte, ValorAporte, ParcelaAporte, CodigoAporte FROM TRVVoucherAporte WHERE TipVou= ? AND Serie = ? AND NumVou = ? ", _
                        0, iTipoPagtoAporte, lNumIntDocAportePagto, dValorAporte, iParcelaAporte, lCodigoAporte, objInfoVoucher.sTipoVou, objInfoVoucher.sSerie, objInfoVoucher.lNumVou)
                If lErro <> AD_SQL_SUCESSO Then gError 196486
        
                'Busca Primeiro
                lErro = Comando_BuscarPrimeiro(alComando(4))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196487
                
                Do While lErro = AD_SQL_SUCESSO
            
                    Set objVouAporte = New ClassTRVVoucherAporte
                    
                    objVouAporte.dValorAporte = dValorAporte
                    objVouAporte.iParcelaAporte = iParcelaAporte
                    objVouAporte.iTipoPagtoAporte = iTipoPagtoAporte
                    objVouAporte.lCodigoAporte = lCodigoAporte
                    objVouAporte.lNumIntDocPagtoAporte = lNumIntDocAportePagto
                    objVouAporte.lNumVou = objInfoVoucher.lNumVou
                    objVouAporte.sSerie = objInfoVoucher.sSerie
                    objVouAporte.sTipVou = objInfoVoucher.sTipoVou

                    objInfoVoucher.colAportes.Add objVouAporte
                    
                    'Busca Próximo
                    lErro = Comando_BuscarProximo(alComando(4))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196488
                    
                Loop
      
            ElseIf sTipoDoc = TRV_TIPODOC_NVL_TEXTO Or sTipoDoc = TRV_TIPODOC_OCR_TEXTO Then
            
                lErro = Comando_ExecutarPos(alComando(5), "SELECT TipoPagtoAporte, NumIntDocPagtoAporte, ValorAporte, ParcelaAporte, CodigoAporte FROM TRVOcorrenciaAporte WHERE NumIntDocOCR = ? ", _
                        0, iTipoPagtoAporte, lNumIntDocAportePagto, dValorAporte, iParcelaAporte, lCodigoAporte, objInfoVoucher.lNumIntDoc)
                If lErro <> AD_SQL_SUCESSO Then gError 196489
        
                'Busca Primeiro
                lErro = Comando_BuscarPrimeiro(alComando(5))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196490
                
                Do While lErro = AD_SQL_SUCESSO
            
                    Set objOCRAporte = New ClassTRVOcorrenciaAporte
                    
                    objOCRAporte.dValorAporte = dValorAporte
                    objOCRAporte.iParcelaAporte = iParcelaAporte
                    objOCRAporte.iTipoPagtoAporte = iTipoPagtoAporte
                    objOCRAporte.lCodigoAporte = lCodigoAporte
                    objOCRAporte.lNumIntDocPagtoAporte = lNumIntDocAportePagto
                    objOCRAporte.lNumIntDocOCR = objInfoVoucher.lNumIntDoc

                    objInfoVoucher.colAportes.Add objOCRAporte
                    
                    'Busca Próximo
                    lErro = Comando_BuscarProximo(alComando(5))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196491
                    
                Loop
            
            End If

            'Busca Proximo
            lErro = Comando_BuscarProximo(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192839
        
        Loop
    
        dValorItens = 0
        dValorComissao = 0
        dValorAporte = 0
        dValorCreditos = 0
        dValorDebitos = 0
        iLinha = iLinha + 1
        
        objInfoFatura.iLinha = iLinha
        
        For Each objInfoVoucher In objInfoFatura.colInfoVouchers
        
            If objInfoVoucher.iMarcado = MARCADO Then
                dValorItens = dValorItens + objInfoVoucher.dValorBruto - objInfoVoucher.dValorAporte - objInfoVoucher.dValorComissao
                dValorComissao = dValorComissao + objInfoVoucher.dValorComissao
                dValorAporte = dValorAporte + objInfoVoucher.dValorAporte
                
                'Se for negativo é um crédito
                If objInfoVoucher.dValor < 0 Then
                    dValorCreditos = dValorCreditos + objInfoVoucher.dValor
                End If
                'Se for um documento fora voucher que acrescebta no valor é um débito
                If objInfoVoucher.dValor > 0 And objInfoVoucher.sTipoDoc <> TRV_TIPODOC_VOU_TEXTO Then
                    dValorDebitos = dValorDebitos + objInfoVoucher.dValor
                End If
                'Se existe um investimento é um crédito
                If objInfoVoucher.dValorAporte > 0 Then
                    dValorCreditos = dValorCreditos - objInfoVoucher.dValorAporte
                End If
                'Se está sendo devolvido um investimento é um débito
                If objInfoVoucher.dValorAporte < 0 Then
                    dValorDebitos = dValorDebitos - objInfoVoucher.dValorAporte
                End If
                
            End If
    
        Next
        
        If dValorItens = 0 Then
            objInfoFatura.iMarcado = DESMARCADO
        End If
        
        objInfoFatura.dValor = dValorItens - objInfoFatura.dValorDeducoes - objInfoFatura.dValorTarifa - objInfoFatura.dValorAporteCred
        objInfoFatura.dValorAporte = dValorAporte
        objInfoFatura.dValorComissao = dValorComissao
        objInfoFatura.dValorCreditos = dValorCreditos - objInfoFatura.dValorAporteCred
        objInfoFatura.dValorDebitos = dValorDebitos
        
        Set objCliente = New ClassCliente
        Set objFilialCliente = New ClassFilialCliente
        Set objEndereco = New ClassEndereco
        Set objFilialEmpresa = New AdmFiliais
    
        objCliente.lCodigo = objInfoFatura.lCliente
        objFilialCliente.lCodCliente = objInfoFatura.lCliente
        objFilialCliente.iCodFilial = FILIAL_MATRIZ
        
        lErro = CF("Cliente_Le", objCliente)
        If lErro <> SUCESSO And lErro <> 12293 Then gError 192840
        
        lErro = CF("FilialCliente_Le", objFilialCliente)
        If lErro <> SUCESSO And lErro <> 12567 Then gError 192841

        objFaturamento.colFatClientes.Add objCliente
        objFaturamento.colFatFiliaisCliente.Add objFilialCliente
                
        objEndereco.lCodigo = objFilialCliente.lEndereco
        
        lErro = CF("Endereco_Le", objEndereco)
        If lErro <> SUCESSO Then gError 192842
        
        objFaturamento.colFatEnderecos.Add objEndereco
                
        objFilialEmpresa.iCodFilial = objInfoFatura.iFilialEmpresa
        
        lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
        If lErro <> SUCESSO Then gError 192843

        objFaturamento.colFatFiliais.Add objFilialEmpresa
        
        objInfoFatura.iLinhaFatFilial = objFaturamento.colFatFiliais.Count
    
        objFaturamento.colFatEnderecos.Add objFilialEmpresa.objEndereco
        
        sArquivo = sDiretorio & gsEmpresaTRVHTML & String(6 - Len(CStr(objInfoFatura.lNumeroFaturaReal)), "0") & CStr(objInfoFatura.lNumeroFaturaReal) & ".html"
        
        If sSiglaDocumento = TIPODOC_CARTAO_CRED_DEB Or objInfoFatura.dValor < 0 Then
            lErro = Gera_Arquivo_Html(objFaturamento, sModeloCP, sArquivo, sConteudoHtml)
        Else
            lErro = Gera_Arquivo_Html(objFaturamento, sModelo, sArquivo, sConteudoHtml)
        End If
        If lErro <> SUCESSO Then gError 192844
    
        'Busca Proximo
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192845
    
    Loop

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TRVFaturas_Regera_Html = SUCESSO

    Exit Function

Erro_TRVFaturas_Regera_Html:

    TRVFaturas_Regera_Html = gErr

    Select Case gErr

        Case 192830
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 192831, 192832, 192845
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOSREC", gErr)
            
        Case 192833, 192838, 192840 To 192844, 194207

        Case 192834, 192835
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PARCELASREC", gErr)
            
        Case 192836, 192837, 192839
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVVOUCHERS", gErr)
            
        Case 196486 To 196488
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVVOUCHERAPORTE", gErr)
        
        Case 196489 To 196491
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVOCORRENCIAAPORTE", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192845)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function Browse_Formata_Campo_Customizado(objBrowse As AdmBrowse, colRegCampo As Collection, objValorCampo As AdmValorCampo, vCampo As Variant, iProximoCampo As Integer, iLinha As Integer, iIndice As Integer, iExercicio As Integer, iPeriodo As Integer, sNomeExterno As String) As Long

Dim lErro As Long

On Error GoTo Erro_Browse_Formata_Campo_Customizado

    Select Case objValorCampo.iTipo

        Case ADM_TIPO_DOUBLE
        
            Select Case objValorCampo.iSubTipo
        
                Case ADM_SUBTIPO_APORTE
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = Format(vCampo, "#,##0.00##")
                    
            End Select

        Case ADM_TIPO_DATE
           

        Case ADM_TIPO_SMALLINT
    
            Select Case objValorCampo.iSubTipo
        
                Case ADM_SUBTIPO_STATUS
'                    Call CF("Converte_Status", vCampo, sStatus)
'                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sStatus

            End Select
    
        Case Else
            
            Select Case objValorCampo.iSubTipo

                Case ADM_SUBTIPO_CGC
                
                    'objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sContaMascarada

            End Select
            
    End Select
    
    Browse_Formata_Campo_Customizado = SUCESSO
    
    Exit Function
    
Erro_Browse_Formata_Campo_Customizado:

    Browse_Formata_Campo_Customizado = gErr

    Select Case gErr
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192874)

    End Select

    Exit Function

End Function

Public Function DocFaturadosLista_BotaoConsulta(ByVal objVoucher As ClassTRVVouchers) As Long

Dim sTela As String
Dim objObjeto As Object

On Error GoTo Erro_DocFaturadosLista_BotaoConsulta

    Select Case objVoucher.iTipoDocDestino
    
        Case TRV_TIPO_DOC_DESTINO_CREDFORN
            sTela = TRV_TIPO_DOC_DESTINO_CREDFORN_TELA
            Set objObjeto = New ClassCreditoPagar
            
        Case TRV_TIPO_DOC_DESTINO_DEBCLI
            sTela = TRV_TIPO_DOC_DESTINO_DEBCLI_TELA
            Set objObjeto = New ClassDebitoRecCli
    
        Case TRV_TIPO_DOC_DESTINO_TITPAG
            sTela = TRV_TIPO_DOC_DESTINO_TITPAG_TELA
            Set objObjeto = New ClassTituloPagar
    
        Case TRV_TIPO_DOC_DESTINO_TITREC
            sTela = TRV_TIPO_DOC_DESTINO_TITREC_TELA
            Set objObjeto = New ClassTituloReceber
    
        Case TRV_TIPO_DOC_DESTINO_NFSPAG
            sTela = TRV_TIPO_DOC_DESTINO_NFSPAG_TELA
            Set objObjeto = New ClassNFsPag
    
    End Select
    
    If Not (objObjeto Is Nothing) Then
    
        objObjeto.lNumIntDoc = objVoucher.lNumIntDocDestino
        
        Call Chama_Tela(sTela, objObjeto)
        
    End If
    
    DocFaturadosLista_BotaoConsulta = SUCESSO

    Exit Function

Erro_DocFaturadosLista_BotaoConsulta:

    DocFaturadosLista_BotaoConsulta = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192879)

    End Select

    Exit Function

End Function

Public Function TRVAPortesPagtoFatHistResLista_BotaoConsulta(ByVal objVoucher As ClassTRVVouchers) As Long

Dim colSelecao As New Collection

On Error GoTo Erro_TRVAPortesPagtoFatHistResLista_BotaoConsulta

    colSelecao.Add objVoucher.lNumIntDocPagtoAporteFat
    colSelecao.Add objVoucher.iTipoPagtoAporte
    colSelecao.Add objVoucher.iTipoDocDestino
    colSelecao.Add objVoucher.lNumIntDocDestino
     
    Call Chama_Tela("TRVAPortesPagtoFatHistLista", colSelecao, Nothing, Nothing, "NumIntDocPagtoAporteFat = ? AND TipoPagtoAporte = ? AND TipoDocDestino = ? AND NumIntDocDestino = ?")

    TRVAPortesPagtoFatHistResLista_BotaoConsulta = SUCESSO

    Exit Function

Erro_TRVAPortesPagtoFatHistResLista_BotaoConsulta:

    TRVAPortesPagtoFatHistResLista_BotaoConsulta = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192879)

    End Select

    Exit Function

End Function

Public Function TRVOcorrencias_Importa(ByVal objImpOcr As ClassImpCoinfoOcor, Optional ByVal bConsiderarFatSigav As Boolean = False) As Long

Dim lErro As Long
Dim alComando(0 To 3) As Long
Dim iIndice As Integer
Dim objOcr As New ClassTRVOcorrencias
Dim objOcrBD As New ClassTRVOcorrencias
Dim objOcrDet As New ClassTRVOcorrenciaDet
Dim lNumFat As Long
Dim lNumIntDocDest As Long
Dim iTipoNumIntDocDest As Integer
Dim iCancelado As Integer
Dim lNumIntDocOCR As Long
Dim iTipoDestino As Integer
Dim iExcluido As Integer

On Error GoTo Erro_TRVOcorrencias_Importa

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 194435
    Next

    objOcr.lCodigo = objImpOcr.lnumocorr + FATOR_COD_OCR_IMPORTACAO
    objOcr.lCliente = objImpOcr.lCodEmp
    objOcr.iFilialCliente = FILIAL_MATRIZ
    objOcr.dtDataEmissao = objImpOcr.dtData
    
    objOcr.dValorTotal = objImpOcr.dValor
    If objImpOcr.idc = DESMARCADO Then objOcr.dValorTotal = -objOcr.dValorTotal
    
    objOcr.sObservacao = ""
    objOcr.iOrigem = IMPORTACAO_OCORRENCIA_CODIGO
    objOcr.sHistorico = Left(objImpOcr.stexto1 & objImpOcr.stexto2 & objImpOcr.stexto3 & objImpOcr.stexto4 & objImpOcr.stexto5 & objImpOcr.stexto6, STRING_TRV_OCR_HISTORICO)
    
    If objImpOcr.iliberado = MARCADO Then
        objOcr.iStatus = STATUS_TRV_OCR_BLOQUEADO
    Else
        objOcr.iStatus = STATUS_TRV_VOU_ABERTO
    End If
    
    objOcr.iFormaPagto = FORMAPAGTO_TRV_OCR_FAT
    objOcr.lNumVou = objImpOcr.lNumVou
    objOcr.sSerie = objImpOcr.sSerie
    objOcr.sTipoDoc = objImpOcr.sTipVou

    objOcrDet.iTipo = IMPORTACAO_OCORRENCIA_TIPO_IMP_CODIGO
    objOcrDet.iSeq = 1
    objOcrDet.dValor = objOcr.dValorTotal
    
    objOcr.colDetalhamento.Add objOcrDet
    
    objOcrBD.lCodigo = objOcr.lCodigo
    
    'Le a ocorrência
    lErro = CF("TRVOcorrencias_Le", objOcrBD)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 194436
    
    If objOcrBD.lNumIntDocDestino = 0 Then
    
        'Se já existe exclui para poder relançar
        If lErro = SUCESSO Then
            lErro = CF("TRVOcorrencias_Exclui_EmTrans", objOcr)
            If lErro <> SUCESSO Then gError 194436
        End If
        
        'Le a tabela TitulosRecTRV
        lErro = Comando_Executar(alComando(0), "SELECT NumFat, Excluido FROM CoinfoItemFatura WHERE TipoDoc = ? AND CodEst = ? AND NumDoc = ? ORDER BY  NumIntArq DESC,  NumRegArqTexto DESC", lNumFat, iExcluido, TRV_TIPODOC_OCR_TEXTO, gsEmpresaTRV, objOcr.lCodigo - FATOR_COD_OCR_IMPORTACAO)
        If lErro <> AD_SQL_SUCESSO Then gError 194437
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194438
        
        If lErro <> AD_SQL_SUCESSO Then
            lNumFat = 0
            iExcluido = 0
        End If
        
        If (lErro = AD_SQL_SUCESSO And iExcluido = DESMARCADO) Or (lErro <> AD_SQL_SUCESSO) Then
        
            lErro = CF("TRVOcorrencias_Grava_EmTrans", objOcr, True, True)
            If lErro <> SUCESSO Then gError 194436
        
            'Le a tabela TitulosRecTRV
            lErro = Comando_Executar(alComando(1), "SELECT Cancelado, TipoDocCorporator, NumIntDocCorporator FROM CoinfoFatura WHERE CodEst = ? AND NumFat = ? AND NumIntDocCorporator <> 0 ORDER BY  NumIntArq DESC,  NumRegArqTexto DESC", iCancelado, iTipoNumIntDocDest, lNumIntDocDest, gsEmpresaTRV, lNumFat)
            If lErro <> AD_SQL_SUCESSO Then gError 194439
        
            'Busca Primeiro
            lErro = Comando_BuscarPrimeiro(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194440
        
            If iCancelado = DESMARCADO Then
            
                lErro = Comando_ExecutarPos(alComando(2), "SELECT NumIntDoc FROM TRVOcorrencias WHERE Codigo= ? ", _
                            0, lNumIntDocOCR, objOcr.lCodigo)
                If lErro <> AD_SQL_SUCESSO Then gError 194441
            
                'Busca Primeiro
                lErro = Comando_BuscarPrimeiro(alComando(2))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194442
                
                If lErro = AD_SQL_SEM_DADOS Then gError 194433
            
                'Trava registro contra alterações/Leituras
                lErro = Comando_LockExclusive(alComando(2))
                If lErro <> SUCESSO Then gError 194434
                
                Select Case iTipoNumIntDocDest
                
                    Case TIPO_TITULO_PAGAR_TRV
                        iTipoDestino = TRV_TIPO_DOC_DESTINO_TITPAG
                        
                    Case TIPO_TITULO_RECEBER_TRV
                        iTipoDestino = TRV_TIPO_DOC_DESTINO_TITREC
                
                    Case TIPO_NF_PAGAR_TRV
                        iTipoDestino = TRV_TIPO_DOC_DESTINO_NFSPAG
                
                End Select
        
                'Importa sem marcar como já faturado (Faturamento em pararelo)
                If Not bConsiderarFatSigav Then
                    lErro = Comando_ExecutarPos(alComando(3), "UPDATE TRVOcorrencias SET Status = ? ", alComando(2), _
                    STATUS_TRV_OCR_LIBERADO)
                Else
                    lErro = Comando_ExecutarPos(alComando(3), "UPDATE TRVOcorrencias SET TipoDocDestino= ?, NumIntDocDestino=?, NumTitulo = ?, Status = ? ", alComando(2), _
                    iTipoDestino, lNumIntDocDest, lNumFat, STATUS_TRV_OCR_LIBERADO)
                End If
                If lErro <> AD_SQL_SUCESSO Then gError 194435
            
            End If
        
        End If
        
        objImpOcr.lNumIntDoc = objOcr.lNumIntDoc
    
    Else
        objImpOcr.lNumIntDoc = objOcrBD.lNumIntDoc
    
    End If
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TRVOcorrencias_Importa = SUCESSO

    Exit Function

Erro_TRVOcorrencias_Importa:

    TRVOcorrencias_Importa = gErr

    Select Case gErr

        Case 194435
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 194436
        
        Case 194437 To 194432
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVOCORRENCIAS", gErr)
    
        Case 194433 'ERRO_OCORRENCIA_NAO_CADASTRADA
            Call Rotina_Erro(vbOKOnly, "ERRO_OCORRENCIA_NAO_CADASTRADA", gErr, objOcr.lCodigo)
        
        Case 194434
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRVOCORRENCIAS", gErr)
        
        Case 194435
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRVOCORRENCIAS", gErr)
               
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194436)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function Endereco_Grava_Customizado(ByVal objEndereco As ClassEndereco) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lCodigo As Long

On Error GoTo Erro_Endereco_Grava_Customizado

    If Not (objEndereco.objInfoUsu Is Nothing) Then

        'Abertura de Comando
        For iIndice = LBound(alComando) To UBound(alComando)
            alComando(iIndice) = Comando_Abrir()
            If alComando(iIndice) = 0 Then gError 196215
        Next
    
        'Le a tabelaTRVOcorrencias
        lErro = Comando_ExecutarPos(alComando(0), "SELECT Codigo FROM Enderecos WHERE Codigo= ? ", 0, lCodigo, objEndereco.lCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 196216
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196217
    
        If lErro <> AD_SQL_SUCESSO Then gError 196218
    
        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 196219
    
        lErro = Comando_ExecutarPos(alComando(1), "UPDATE Enderecos SET Email2= ? ", alComando(0), objEndereco.objInfoUsu.sEmail2)
        If lErro <> AD_SQL_SUCESSO Then gError 196220
            
        'Fecha Comando
        For iIndice = LBound(alComando) To UBound(alComando)
            Call Comando_Fechar(alComando(iIndice))
        Next
        
    End If

    Endereco_Grava_Customizado = SUCESSO

    Exit Function

Erro_Endereco_Grava_Customizado:

    Endereco_Grava_Customizado = gErr

    Select Case gErr
        
        Case 196215
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 196216, 196217
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ENDERECOS", gErr)

        Case 196218
            Call Rotina_Erro(vbOKOnly, "ERRO_ENDERECOS_NAO_CADASTRADO", gErr)

        Case 196219
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_ENDERECOS", gErr)

        Case 196220
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_ENDERECOS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196221)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function TRVTitulosExp_Insere(ByVal objTRVTitulosExp As ClassTRVTitulosExp, Optional lComandoP1 As Long = 0, Optional lComandoP2 As Long = 0) As Long

Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim objFaturaAporte As ClassTRVFaturaAporte

On Error GoTo Erro_TRVTitulosExp_Insere

    If lComandoP1 <> 0 And lComandoP2 <> 0 Then
        alComando(1) = lComandoP1
        alComando(2) = lComandoP2
    Else
        'Abertura de Comando
        For iIndice = LBound(alComando) To UBound(alComando)
            alComando(iIndice) = Comando_Abrir()
            If alComando(iIndice) = 0 Then gError 196236
        Next
    End If

    'Obter NumIntDoc
    lErro = CF("Config_ObterNumInt", "TRVConfig", "NUM_INT_PROX_TRVTITULOSEXP", lNumIntDoc)
    If lErro <> SUCESSO Then gError 196237

    objTRVTitulosExp.lNumIntDoc = lNumIntDoc

    lErro = Comando_Executar(alComando(1), "INSERT INTO TRVTitulosExp( NumIntDoc, Usuario, Data, Hora, " & _
            "TipoDocOrigem, NumIntDocOrigem, NumTitulo, Excluido, Exportado, Motivo, TemQueContabilizar," & _
            "ValorAporteCred,ValorAporte,ValorComissao,ValorCreditos,ValorDebitos,ValorTarifa,ValorDeducoes) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
            objTRVTitulosExp.lNumIntDoc, gsUsuario, Date, CDbl(Time), _
            objTRVTitulosExp.iTipoDocOrigem, objTRVTitulosExp.lNumIntDocOrigem, objTRVTitulosExp.lNumTitulo, objTRVTitulosExp.iExcluido, objTRVTitulosExp.iExportado, objTRVTitulosExp.sMotivo, objTRVTitulosExp.iTemQueContabilizar, _
            objTRVTitulosExp.dValorAporteCred, objTRVTitulosExp.dValorAporte, objTRVTitulosExp.dValorComissao, objTRVTitulosExp.dValorCreditos, objTRVTitulosExp.dValorDebitos, _
            objTRVTitulosExp.dValorTarifa, objTRVTitulosExp.dValorDeducoes)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196238
    
    For Each objFaturaAporte In objTRVTitulosExp.colFatAporteCred
    
        lErro = Comando_Executar(alComando(2), "INSERT INTO TRVTitulosExpAporte( NumIntDocTitExp,CodigoAporte,ParcelaAporte,NumParcelas,ValorAporte,NumIntDocCredito,PagtoParcial) VALUES (?,?,?,?,?,?,?)", _
                objTRVTitulosExp.lNumIntDoc, objFaturaAporte.lCodigoAporte, objFaturaAporte.iParcelaAporte, objFaturaAporte.iNumParcelas, objFaturaAporte.dValorAporte, objFaturaAporte.lNumIntDocCredito, objFaturaAporte.iPagtoParcial)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196238
    
    Next

    If lComandoP1 = 0 Or lComandoP2 = 0 Then
        'Fecha Comando
        For iIndice = LBound(alComando) To UBound(alComando)
            Call Comando_Fechar(alComando(iIndice))
        Next
    End If

    TRVTitulosExp_Insere = SUCESSO

    Exit Function

Erro_TRVTitulosExp_Insere:

    TRVTitulosExp_Insere = gErr

    Select Case gErr

        Case 196236
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 196237

        Case 196238
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRVTITULOSEXP", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196239)

    End Select

    If lComandoP1 = 0 Or lComandoP2 = 0 Then
        'Fecha Comando
        For iIndice = LBound(alComando) To UBound(alComando)
            Call Comando_Fechar(alComando(iIndice))
        Next
    End If

    Exit Function

End Function

Public Function TRVTitulos_Exporta_Sigav() As Long

Dim lErro As Long
Dim lTransacao As Long
Dim alComando(1 To 5) As Long
Dim iIndice As Integer
Dim objTRVTitulosExp As ClassTRVTitulosExp
Dim colTitulos As New Collection
Dim colRegTitulos As New Collection
Dim colRegItens As New Collection
Dim sRegistroTit As String
Dim sRegistroItens As String
Dim dtDataEmissao As Date
Dim dtDataVencimento As Date
Dim lCodEmp As Long
Dim dValor As Double
Dim iCpCr As Integer
Dim objTitPag As ClassTituloPagar
Dim objTitRec As ClassTituloReceber
Dim objNFsPag As ClassNFsPag
Dim dTotalAporte As Double
Dim dValorAporte As Double
Dim dtDataVou As Date
Dim lNumVou As Long
Dim dValorVou As Double
Dim lNumeroDoc As Long
Dim sTipoDoc As String
Dim sSerie As String
Dim sTipVou As String
Dim sCodGrupo As String
Dim sNomeArq1 As String
Dim sNomeArq2 As String
Dim sParteNomeArq As String
Dim vRegistro As Variant
Dim bArq1Aberto As Boolean
Dim bArq2Aberto As Boolean
Dim sDiretorio As String
Dim colParcelaReceber As colParcelaReceber
Dim colParcelasPag As colParcelaPagar
Dim sUsuario As String
Dim dtdatareg As Date
Dim iMarca As Integer
Dim lrecnsica As Long
Dim iQtdPax As Integer
Dim icondpag As Integer
Dim sControle As String
Dim iCartao As Integer
Dim sCodPro As String
Dim sGrupo As String
Dim sValidCC As String
Dim iQuantParc As Integer
Dim dVlCartao As Double
Dim lNumAuto As Long
Dim sCiaCart As String
Dim idiasantc As Integer
Dim imaster As Integer
Dim lnummstr As Long
Dim lKit As Long
Dim lRetorno As Long
Dim tTRVTitulosExp As typeTRVTitulosExp
Dim bIgnora As Boolean

On Error GoTo Erro_TRVTitulos_Exporta_Sigav

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 196243

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 196244
    Next
    
    sParteNomeArq = Format(Now, "YYYYMMDDHHMMSS")
    
    sNomeArq1 = gsEmpresaTRV & "cabecfat" & sParteNomeArq & ".d"
    sNomeArq2 = gsEmpresaTRV & "itemfatu" & sParteNomeArq & ".d"
    
    'Obtém o diretório onde estão os arquivos
    sDiretorio = String(512, 0)
    lRetorno = GetPrivateProfileString("Geral", "dirArqImpCorpFat", "c:\travelace\corporator\", sDiretorio, 512, "ADM100.INI")
    sDiretorio = Left(sDiretorio, lRetorno)
    
    'Alocação de espaço no buffer
    tTRVTitulosExp.sUsuario = String(STRING_USUARIO, 0)

    'Le a tabelaTRVTitulosExp
    lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntDoc, Usuario, Data, Hora, TipoDocOrigem, NumIntDocOrigem, " & _
                "NumTitulo, Excluido, Exportado FROM TRVTitulosExp WHERE Exportado = ? ORDER BY NumIntDoc ", 0, _
                tTRVTitulosExp.lNumIntDoc, tTRVTitulosExp.sUsuario, tTRVTitulosExp.dtData, tTRVTitulosExp.dHora, _
                tTRVTitulosExp.iTipoDocOrigem, tTRVTitulosExp.lNumIntDocOrigem, tTRVTitulosExp.lNumTitulo, tTRVTitulosExp.iExcluido, tTRVTitulosExp.iExportado, 0)
    If lErro <> AD_SQL_SUCESSO Then gError 196245

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196246

    Do While lErro <> AD_SQL_SEM_DADOS
    
        bIgnora = False
    
        Set objTRVTitulosExp = New ClassTRVTitulosExp

        objTRVTitulosExp.lNumIntDoc = tTRVTitulosExp.lNumIntDoc
        objTRVTitulosExp.sUsuario = tTRVTitulosExp.sUsuario
        objTRVTitulosExp.dtData = tTRVTitulosExp.dtData
        objTRVTitulosExp.dHora = tTRVTitulosExp.dHora
        objTRVTitulosExp.iTipoDocOrigem = tTRVTitulosExp.iTipoDocOrigem
        objTRVTitulosExp.lNumIntDocOrigem = tTRVTitulosExp.lNumIntDocOrigem
        objTRVTitulosExp.lNumTitulo = tTRVTitulosExp.lNumTitulo
        objTRVTitulosExp.iExcluido = tTRVTitulosExp.iExcluido
        objTRVTitulosExp.iExportado = tTRVTitulosExp.iExportado
        
        Select Case objTRVTitulosExp.iTipoDocOrigem
        
            Case TRV_TIPO_DOC_DESTINO_TITPAG
            
                Set objTitPag = New ClassTituloPagar
                Set colParcelasPag = New colParcelaPagar
                
                objTitPag.lNumIntDoc = objTRVTitulosExp.lNumIntDocOrigem
                
                lErro = CF("TituloPagar_Le_Todos", objTitPag, True)
                If lErro <> SUCESSO Then gError 196247
                
                'Lê as Parcelas a Pagar vinculadas ao Título
                lErro = CF("ParcelasPagar_Le_Todas", objTitPag, colParcelasPag, True)
                If lErro <> SUCESSO Then gError 196248
        
                iCpCr = MARCADO
                dtDataEmissao = objTitPag.dtDataEmissao
                lCodEmp = objTitPag.lFornecedor
                dValor = -objTitPag.dValorTotal
                dtDataVencimento = colParcelasPag.Item(1).dtDataVencimento
                
                If objTitPag.sSiglaDocumento <> TIPODOC_FATURA_SERVICO_CP Then
                    bIgnora = True
                End If
        
            Case TRV_TIPO_DOC_DESTINO_TITREC
            
                Set objTitRec = New ClassTituloReceber
                Set colParcelaReceber = New colParcelaReceber
                
                objTitRec.lNumIntDoc = objTRVTitulosExp.lNumIntDocOrigem

                'Lê o Título
                lErro = CF("TituloReceber_Le", objTitRec, MARCADO)
                If lErro <> SUCESSO And lErro <> 26061 Then gError 196249
        
                'Se não achou o Título
                If lErro <> SUCESSO Then
        
                    'Lê o Título a Receber baixado
                    lErro = CF("TituloReceberBaixado_Le", objTitRec, MARCADO)
                    If lErro <> SUCESSO And lErro <> 56570 Then gError 196250
        
                End If
        
                'Lê as Parcelas a Receber vinculadas ao Título
                lErro = CF("ParcelasReceber_Le_Todas", objTitRec, colParcelaReceber, True)
                If lErro <> SUCESSO And lErro <> 58990 Then gError 196251
    
                iCpCr = DESMARCADO
                dtDataEmissao = objTitRec.dtDataEmissao
                lCodEmp = objTitRec.lCliente
                dValor = objTitRec.dValor
                dtDataVencimento = colParcelaReceber.Item(1).dtDataVencimento
                
                If objTitRec.sSiglaDocumento <> TIPODOC_FATURA_SERVICO_CR And objTitRec.sSiglaDocumento <> TIPODOC_CARTAO_CRED_DEB Then
                    bIgnora = True
                End If
                
            Case TRV_TIPO_DOC_DESTINO_NFSPAG
            
                Set objNFsPag = New ClassNFsPag
                
                objNFsPag.lNumIntDoc = objTRVTitulosExp.lNumIntDocOrigem
                
                'Lê o Título
                lErro = CF("NFPag_Le_Todos2", objNFsPag, MARCADO)
                If lErro <> SUCESSO And lErro <> 18257 Then gError 196252
        
                iCpCr = MARCADO
                dtDataEmissao = objNFsPag.dtDataEmissao
                lCodEmp = objNFsPag.lFornecedor
                dValor = -objNFsPag.dValorTotal
                dtDataVencimento = objNFsPag.dtDataVencimento
        
        End Select
        
        dTotalAporte = 0
        
        If Not bIgnora Then
        
            sTipoDoc = String(STRING_TRV_OCR_TIPODOC, 0)
            sSerie = String(STRING_TRV_OCR_SERIE, 0)
            sCodGrupo = String(STRING_TRV_OCR_CODGRUPO, 0)
            sTipVou = String(STRING_TRV_OCR_TIPOVOU, 0)
            
            'Le a view DocFaturados
            lErro = Comando_Executar(alComando(2), "SELECT Data, NumVou, Valor, Numero, TipoDoc, Serie, TipVou, CodGrupo, ValorAporte FROM DocFaturados WHERE TipoDocDestino = ? AND NumIntDocDestino = ? ", _
            dtDataVou, lNumVou, dValorVou, lNumeroDoc, sTipoDoc, sSerie, sTipVou, sCodGrupo, dValorAporte, objTRVTitulosExp.iTipoDocOrigem, objTRVTitulosExp.lNumIntDocOrigem)
            If lErro <> AD_SQL_SUCESSO Then gError 196253
        
            'Busca Primeiro
            lErro = Comando_BuscarPrimeiro(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196254
        
            Do While lErro <> AD_SQL_SEM_DADOS
            
                sUsuario = String(255, 0)
                sControle = String(255, 0)
                sGrupo = String(255, 0)
                sValidCC = String(255, 0)
                sCiaCart = String(255, 0)
                sCodPro = String(255, 0)
    
                'Le a tabela TRVTitulosExp
                lErro = Comando_Executar(alComando(3), "SELECT NumDoc, Usuario, DataReg, marca, recnsica,QtdPax,condpag, Controle, cartao, CodPro, Grupo, ValidCC, QuantParc, VlCartao, NumAuto, CiaCart,diasantc,master,nummstr,kit FROM CoinfoItemFatura WHERE TipVou = ? AND TipoDoc = ? AND NumVou = ? AND Serie = ? ORDER BY NumIntArq DESC, NumRegArqTexto DESC", _
                lNumeroDoc, sUsuario, dtdatareg, iMarca, lrecnsica, iQtdPax, icondpag, sControle, iCartao, sCodPro, sGrupo, sValidCC, iQuantParc, dVlCartao, lNumAuto, sCiaCart, idiasantc, imaster, lnummstr, lKit, sTipVou, sTipoDoc, lNumVou, sSerie)
                If lErro <> AD_SQL_SUCESSO Then gError 196255
            
                'Busca Primeiro
                lErro = Comando_BuscarPrimeiro(alComando(3))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196256
            
                dTotalAporte = dTotalAporte + dValorAporte
                
                sRegistroItens = CStr(lCodEmp) & " " & Format(dtDataVou, "dd/mm/yy") & " " & CStr(lNumVou) & " " & _
                """BRL""" & " " & Format(dValorVou, "STANDARD") & " " & """" & gsEmpresaTRV & """" & " " & CStr(objTRVTitulosExp.lNumTitulo) & " " & _
                Format(dtDataVencimento, "dd/mm/yy") & " " & CStr(lNumeroDoc) & " " & """" & sUsuario & """" & " " & Format(dtdatareg, "dd/mm/yy") & " " & _
                IIf(iMarca = 0, "no", "yes") & " " & """" & sTipoDoc & """" & " " & CStr(iQtdPax) & " " & CStr(lrecnsica) & " " & _
                CStr(icondpag) & " " & """" & sSerie & """" & " " & """" & sTipVou & """" & " " & """" & sCodGrupo & """" & " " & """" & sControle & """" & " " & _
                IIf(iCartao = 0, "no", "yes") & " " & """" & sCodPro & """" & " " & """" & sGrupo & """" & " " & "?" & " " & IIf(sValidCC = "", "?", """" & sValidCC & """") & " " & _
                CStr(iQuantParc) & " " & Format(dVlCartao, "STANDARD") & " " & _
                CStr(lNumAuto) & " " & IIf(sCiaCart = "", "?", """" & sCiaCart & """") & _
                " no no ? " & IIf(idiasantc = MARCADO, "yes", "no") & " " & IIf(imaster = MARCADO, "yes", "no") & " " & _
                CStr(lnummstr) & " " & CStr(lKit)
                
                colRegItens.Add sRegistroItens
            
                'Busca Primeiro
                lErro = Comando_BuscarPrimeiro(alComando(2))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196257
            
            Loop
            
            sRegistroTit = CStr(objTRVTitulosExp.lNumTitulo) & " " & Format(dtDataEmissao, "dd/mm/yy") & " " & _
            CStr(lCodEmp) & " " & """BRL""" & " " & Format(dValor, "STANDARD") & " " & """" & gsEmpresaTRV & """" & " " & _
            Format(dtDataVencimento, "dd/mm/yy") & " " & IIf(objTRVTitulosExp.iExcluido = MARCADO, "yes", "no") & " " & _
            IIf(iCpCr = MARCADO, "yes", "no") & " " & """" & gsUsuario & """" & " " & _
            Format(objTRVTitulosExp.dtData, "dd/mm/yy") & " 0 0 " & _
            Format(dTotalAporte / (dValor + dTotalAporte), "STANDARD") & " " & Format(dTotalAporte, "STANDARD") & _
            " " & Format(dValor + dTotalAporte, "STANDARD") & " " & _
            IIf(objTRVTitulosExp.iExcluido = MARCADO, Format(objTRVTitulosExp.dtData, "dd/mm/yy"), "?") & _
            " no ? 0 yes " & Format(Date, "dd/mm/yy")
            
            colTitulos.Add objTRVTitulosExp
            colRegTitulos.Add sRegistroTit
            
        
            'Marca como exportado
            lErro = Comando_ExecutarPos(alComando(4), "UPDATE TRVTitulosExp SET Exportado = ?, DataExportacao = ?, HoraExportacao = ?, UsuarioExportacao = ?, NomeArq = ?", alComando(1), MARCADO, Date, CDbl(Time), gsUsuario, sNomeArq1)
        Else
            'Marca como exportado
            lErro = Comando_ExecutarPos(alComando(4), "UPDATE TRVTitulosExp SET Exportado = ?, DataExportacao = ?, HoraExportacao = ?, UsuarioExportacao = ?, NomeArq = ?", alComando(1), MARCADO, Date, CDbl(Time), gsUsuario, "")
        End If
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196258
        
        'Busca Próximo
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196259
    
    Loop
    
    Open sDiretorio & sNomeArq1 For Output As #1
    bArq1Aberto = True
    
    For Each vRegistro In colRegTitulos
        Print #1, vRegistro
    Next
    
    Close #1
    bArq1Aberto = False
    
    Open sDiretorio & sNomeArq2 For Output As #2
    bArq2Aberto = True
    
    For Each vRegistro In colRegItens
        Print #2, vRegistro
    Next
    
    Close #2
    bArq2Aberto = False
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 196260

    TRVTitulos_Exporta_Sigav = SUCESSO

    Exit Function

Erro_TRVTitulos_Exporta_Sigav:

    TRVTitulos_Exporta_Sigav = gErr

    Select Case gErr

        Case 196243
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 196244
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 196245, 196246, 196259
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVTITULOSEXP", gErr)
            
        Case 196247 To 196252
        
        Case 196253 To 196257
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVVOUCHERS", gErr)

        Case 196258
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRVTITULOSEXP", gErr)

        Case 196260
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196261)

    End Select

    If bArq1Aberto Then Close #1
    If bArq2Aberto Then Close #2

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Function Faturamento_Excluir_PorEmissao(ByVal dtDataEmissao As Date) As Long

Dim lErro As Long
Dim alComando(0 To 7) As Long
Dim iIndice As Integer, iFilialEmpresa As Integer
Dim lNumIntDoc As Long, objTitPag As ClassTituloPagar, objTitRec As ClassTituloReceber
Dim colTP As New Collection, colTR As New Collection

On Error GoTo Erro_Faturamento_Excluir_PorEmissao

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 1001
    Next
        
'    lErro = Comando_Executar(alComando(0), "SELECT NumIntDoc, FilialEmpresa FROM TitulosPag WHERE DataEmissao = ? AND sigladocumento = ? AND Status <> ? ORDER BY NumIntDoc", _
'        lNumIntDoc, iFilialEmpresa, dtDataEmissao, "FSCP", STATUS_EXCLUIDO)
'    If lErro <> AD_SQL_SUCESSO Then gError 1002
'
'    lErro = Comando_BuscarProximo(alComando(0))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 1003
'
'    Do While lErro = AD_SQL_SUCESSO
'
'        Set objTitPag = New ClassTituloPagar
'
'        With objTitPag
'            .lNumIntDoc = lNumIntDoc
'            .iFilialEmpresa = iFilialEmpresa
'            .sSiglaDocumento = "FSCP"
'        End With
'
'        colTP.Add objTitPag
'
'        lErro = Comando_BuscarProximo(alComando(0))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 1005
'
'    Loop
'
'    For Each objTitPag In colTP
'
'        lErro = CF("NFFatPag_Exclui", objTitPag, Nothing)
'        If lErro <> SUCESSO Then gError 1004
'
'    Next
    
    lErro = Comando_Executar(alComando(1), "SELECT NumIntDoc, FilialEmpresa FROM TitulosRec WHERE DataEmissao = ? AND sigladocumento = ? AND Status <> ? ORDER BY NumIntDoc", _
        lNumIntDoc, iFilialEmpresa, dtDataEmissao, "FSCR", STATUS_EXCLUIDO)
    If lErro <> AD_SQL_SUCESSO Then gError 1006
    
    lErro = Comando_BuscarProximo(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 1007
    
    Do While lErro = AD_SQL_SUCESSO
    
        Set objTitRec = New ClassTituloReceber
    
        With objTitRec
            .lNumIntDoc = lNumIntDoc
            .iFilialEmpresa = iFilialEmpresa
            .sSiglaDocumento = "FSCR"
        End With
        
        colTR.Add objTitRec
        
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 1009
    
    Loop
    
    For Each objTitRec In colTR
    
        lErro = CF("TituloReceber_Exclui", objTitRec, Nothing)
        If lErro <> SUCESSO Then gError 1008
        
    Next
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Faturamento_Excluir_PorEmissao = SUCESSO
    
    Exit Function
    
Erro_Faturamento_Excluir_PorEmissao:

    Faturamento_Excluir_PorEmissao = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 1010)

    End Select
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Public Function Cliente_Grava_Customizado(ByVal objCliente As ClassCliente) As Long

Dim lErro As Long
Dim alComando(1 To 8) As Long
Dim iIndice As Integer
Dim lCliente As Integer
Dim objClienteTRV As ClassClienteTRV
Dim objCliEmissoresTRV As ClassCliEmissoresTRV
Dim lNumIntDoc As Long
Dim lNumIntDoc1 As Long
Dim iIndice1 As Integer
Dim objCliEmissoresExcTRV As ClassCliEmissoresExcTRV
Dim objTRVCliRep As ClassTRVClienteRepProd
Dim objTRVCliCor As ClassTRVClienteCorProd

On Error GoTo Erro_Cliente_Grava_Customizado

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 195811
    Next

    lErro = Comando_Executar(alComando(1), "DELETE FROM TRVCliEmissoresExc WHERE NumIntDocCliEmi IN (Select NumIntDoc FROM TRVCliEmissores WHERE Cliente = ?)", objCliente.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 195812

    lErro = Comando_Executar(alComando(2), "DELETE FROM TRVCliEmissores WHERE Cliente = ?", objCliente.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 195813
            
    iIndice = 0

    If Not (objCliente.objInfoUsu Is Nothing) Then
    
        Set objClienteTRV = objCliente.objInfoUsu
        
        If objClienteTRV.colCliEmissoresTRV.Count > 0 Then
        
            lErro = CF("Config_ObterNumInt", "TRVConfig", TRVCONFIG_NUM_INT_PROX_TRVCLIEMISSORES, lNumIntDoc, objClienteTRV.colCliEmissoresTRV.Count)
            If lErro <> SUCESSO Then gError 195814

            For Each objCliEmissoresTRV In objClienteTRV.colCliEmissoresTRV
            
                iIndice = iIndice + 1
                
                objCliEmissoresTRV.lNumIntDoc = lNumIntDoc
                objCliEmissoresTRV.iSeq = iIndice
                
                lNumIntDoc = lNumIntDoc + 1
    
                lErro = Comando_Executar(alComando(3), "INSERT INTO TRVCliEmissores (NumIntDoc, Cliente, Seq, FornEmissor, PercComissao) VALUES (?,?,?,?,?)", _
                         objCliEmissoresTRV.lNumIntDoc, objCliente.lCodigo, objCliEmissoresTRV.iSeq, objCliEmissoresTRV.lFornEmissor, objCliEmissoresTRV.dPercComissao)
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 195815
    
                If objCliEmissoresTRV.colCliEmissoresExcTRV.Count > 0 Then
    
                    lErro = CF("Config_ObterNumInt", "TRVConfig", TRVCONFIG_NUM_INT_PROX_TRVCLIEMISSORESEXC, lNumIntDoc1, objCliEmissoresTRV.colCliEmissoresExcTRV.Count)
                    If lErro <> SUCESSO Then gError 195816
                    
                    iIndice1 = 0
                    
                    For Each objCliEmissoresExcTRV In objCliEmissoresTRV.colCliEmissoresExcTRV
                        
                        iIndice1 = iIndice1 + 1
                        
                        objCliEmissoresExcTRV.lNumIntDoc = lNumIntDoc1
                        objCliEmissoresExcTRV.lNumIntDocCliEmi = objCliEmissoresTRV.lNumIntDoc
                        objCliEmissoresExcTRV.iSeq = iIndice1
                        
                        lNumIntDoc1 = lNumIntDoc1 + 1
                        
                        lErro = Comando_Executar(alComando(4), "INSERT INTO TRVCliEmissoresExc (NumIntDoc, NumIntDocCliEmi, Seq, Produto, PercComissao) VALUES (?,?,?,?,?)", _
                                 objCliEmissoresExcTRV.lNumIntDoc, objCliEmissoresExcTRV.lNumIntDocCliEmi, objCliEmissoresExcTRV.iSeq, objCliEmissoresExcTRV.sProduto, objCliEmissoresExcTRV.dPercComissao)
                        If lErro <> AD_SQL_SUCESSO Then gError 195817
                        
                    Next
    
                End If
    
            Next
    
        End If
    
        lErro = Comando_Executar(alComando(5), "DELETE FROM FiliaisClientesTRV WHERE CodCliente = ?  AND CodFilial = ?", objCliente.lCodigo, FILIAL_MATRIZ)
        If lErro <> AD_SQL_SUCESSO Then gError 195859
        
        If objClienteTRV.iFilialCoinfo = 0 Then
            objClienteTRV.iFilialCoinfo = Filial_Corporator_Retorna_Coinfo(objClienteTRV.iFilialEmpresa)
        End If
    
        lErro = Comando_Executar(alComando(6), "INSERT INTO FiliaisClientesTRV (CodCliente, CodFilial, Representante, ComissaoRep, Correntista, ComissaoCorr, ComissaoAg, Aporte, EmpresaPai, FilialNF, Grupo, FilialEmpresa, FilialCoinfo, FilialFat, CondPagtoCC) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", objCliente.lCodigo, FILIAL_MATRIZ, _
        objClienteTRV.lCodRepresentante, objClienteTRV.dPercComiRep, objClienteTRV.lCodCorrentista, objClienteTRV.dPercComiCorr, objClienteTRV.dPercComiAg, objClienteTRV.iConsiderarAporte, objClienteTRV.lEmpresaPai, objClienteTRV.iFilialNF, objClienteTRV.sGrupo, objClienteTRV.iFilialEmpresa, objClienteTRV.iFilialCoinfo, objClienteTRV.iFilialFat, objClienteTRV.iCondPagtoCC)
        If lErro <> AD_SQL_SUCESSO Then gError 195860
        
        
        lErro = TRVClienteCorProd_Exclui(objCliente)
        If lErro <> SUCESSO Then gError 196955
        
        lErro = TRVClienteRepProd_Exclui(objCliente)
        If lErro <> SUCESSO Then gError 196956
        
        For Each objTRVCliCor In objClienteTRV.colExcCor
            objTRVCliCor.lCliente = objCliente.lCodigo
            lErro = TRVClienteCorProd_Insere(alComando(7), objTRVCliCor)
            If lErro <> SUCESSO Then gError 196957
        Next
        
        For Each objTRVCliRep In objClienteTRV.colExcRep
            objTRVCliRep.lCliente = objCliente.lCodigo
            lErro = TRVClienteRepProd_Insere(alComando(8), objTRVCliRep)
            If lErro <> SUCESSO Then gError 196958
        Next
    
    End If

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Cliente_Grava_Customizado = SUCESSO

    Exit Function

Erro_Cliente_Grava_Customizado:

    Cliente_Grava_Customizado = gErr

    Select Case gErr

        Case 195811
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 195812
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRVCLIEMISSORESEXC", gErr)

        Case 195813
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRVCLIEMISSORES", gErr)
        
        Case 195814, 195816, 196955 To 196958
        
        Case 195815
            Call Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_TRVCLIEMISSORES", gErr)

        Case 195817
            Call Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_TRVCLIEMISSORESEXC", gErr)

        Case 195859
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_FILIAISCLIENTESTRV", gErr)

        Case 195859
            Call Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_FILIAISCLIENTESTRV", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195818)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function Cliente_Exclui_Customizado(ByVal objCliente As ClassCliente) As Long

Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim lCliente As Integer

On Error GoTo Erro_Cliente_Exclui_Customizado

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 195819
    Next

    lErro = Comando_Executar(alComando(1), "DELETE FROM TRVCliEmissoresExc WHERE NumIntDocCliEmi IN (Select NumIntDoc FROM TRVCliEmissores WHERE Cliente = ?)", objCliente.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 195820

    lErro = Comando_Executar(alComando(2), "DELETE FROM TRVCliEmissores WHERE Cliente = ?", objCliente.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 195821

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Cliente_Exclui_Customizado = SUCESSO

    Exit Function

Erro_Cliente_Exclui_Customizado:

    Cliente_Exclui_Customizado = gErr

    Select Case gErr

        Case 195819
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 195820
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRVCLIEMISSORESEXC", gErr)

        Case 195821
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRVCLIEMISSORES", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195822)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function TRVAcordo_Grava(ByVal objTRVAcordo As ClassTRVAcordos) As Long
    
Dim lErro As Long
Dim alComando(0 To 20) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim lTransacao As Long

On Error GoTo Erro_TRVAcordo_Grava

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 197124

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 197125
    Next

    'Le a tabelaTRVAcordos
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM TRVAcordos WHERE Numero = ? ", _
                0, lNumIntDoc, objTRVAcordo.lNumero)
    If lErro <> AD_SQL_SUCESSO Then gError 197126

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197127

    'Se existir => UPDATE, senão => INSERT
    If lErro = AD_SQL_SUCESSO Then

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 197128
        
        objTRVAcordo.lNumIntDoc = lNumIntDoc
        
        lErro = Comando_ExecutarPos(alComando(1), "UPDATE TRVAcordos SET Contrato = ?,  Cliente= ? , FilialCliente= ? , ValidadeDe = ? , ValidadeAte =?,  Observacao= ? , Descricao = ?", alComando(0), _
                objTRVAcordo.sContrato, objTRVAcordo.lCliente, objTRVAcordo.iFilialCliente, objTRVAcordo.dtValidadeDe, objTRVAcordo.dtValidadeAte, objTRVAcordo.sObservacao, objTRVAcordo.sDescricao)
        If lErro <> AD_SQL_SUCESSO Then gError 197129

    Else

        'Obter NumIntDoc
        lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_INT_PROX_TRVACORDO", lNumIntDoc)
        If lErro <> SUCESSO Then gError 197130

        objTRVAcordo.lNumIntDoc = lNumIntDoc

        lErro = Comando_Executar(alComando(1), "INSERT INTO TRVAcordos( NumIntDoc, Numero, Contrato, Cliente, FilialCliente, " & _
                "ValidadeDe, ValidadeAte, Observacao, Descricao) VALUES (?,?,?,?,?,?,?,?,?)", _
                objTRVAcordo.lNumIntDoc, objTRVAcordo.lNumero, objTRVAcordo.sContrato, objTRVAcordo.lCliente, objTRVAcordo.iFilialCliente, _
                objTRVAcordo.dtValidadeDe, objTRVAcordo.dtValidadeAte, objTRVAcordo.sObservacao, objTRVAcordo.sDescricao)
        If lErro <> AD_SQL_SUCESSO Then gError 197131

    End If
    
    lErro = TRVAcordoComissaoTarifa_Exclui(objTRVAcordo)
    If lErro <> SUCESSO Then gError 197132

    lErro = TRVAcordoComissaoTarifa_Inclui(objTRVAcordo)
    If lErro <> SUCESSO Then gError 197133

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 197134

    TRVAcordo_Grava = SUCESSO

    Exit Function

Erro_TRVAcordo_Grava:

    TRVAcordo_Grava = gErr

    Select Case gErr

        Case 197124
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 197125
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 197126, 197127
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVACORDOS", gErr)

        Case 197128
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRVACORDOS", gErr)

        Case 197129
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_TRVACORDOS", gErr)

        Case 197130, 197132, 197133
        
        Case 197131
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRVACORDOS", gErr)
            
        Case 197134
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197135)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Function TRVAcordoComissaoTarifa_Exclui(ByVal objTRVAcordo As ClassTRVAcordos) As Long

Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim lCliente As Integer

On Error GoTo Erro_TRVAcordoComissaoTarifa_Exclui

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 197120
    Next

    lErro = Comando_Executar(alComando(1), "DELETE FROM TRVAcordoComissao WHERE NumIntAcordo IN (Select NumIntDoc FROM TRVAcordos WHERE Numero = ?)", objTRVAcordo.lNumero)
    If lErro <> AD_SQL_SUCESSO Then gError 197121

    lErro = Comando_Executar(alComando(2), "DELETE FROM TRVAcordoTarifaDif WHERE NumIntAcordo IN (Select NumIntDoc FROM TRVAcordos WHERE Numero = ?)", objTRVAcordo.lNumero)
    If lErro <> AD_SQL_SUCESSO Then gError 197122

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TRVAcordoComissaoTarifa_Exclui = SUCESSO

    Exit Function

Erro_TRVAcordoComissaoTarifa_Exclui:

    TRVAcordoComissaoTarifa_Exclui = gErr

    Select Case gErr

        Case 197120
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 197121
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRVACORDOCOMISSAO", gErr)

        Case 197122
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRVACORDOTARIFADIF", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197123)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function TRVAcordoComissaoTarifa_Inclui(ByVal objTRVAcordo As ClassTRVAcordos) As Long

Dim lErro As Long
Dim alComando(1 To 2) As Long

Dim iIndice As Integer
Dim lCliente As Integer
Dim objTRVAcordoComiss As ClassTRVAcordoComissao
Dim objTRVAcordoDif As ClassTRVAcordoTarifaDif
Dim lNumIntDoc As Long

On Error GoTo Erro_TRVAcordoComissaoTarifa_Inclui

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 197114
    Next

    iIndice = 0

    If objTRVAcordo.colTRVAcordoComiss.Count > 0 Then
    
        lErro = CF("Config_ObterNumInt", "TRVConfig", TRVCONFIG_NUM_INT_PROX_TRVACORDOCOMISS, lNumIntDoc, objTRVAcordo.colTRVAcordoComiss.Count)
        If lErro <> SUCESSO Then gError 197115

        For Each objTRVAcordoComiss In objTRVAcordo.colTRVAcordoComiss
        
            iIndice = iIndice + 1
            
            objTRVAcordoComiss.lNumIntDoc = lNumIntDoc
            objTRVAcordoComiss.lNumIntAcordo = objTRVAcordo.lNumIntDoc
            objTRVAcordoComiss.iSeq = iIndice
            
            lNumIntDoc = lNumIntDoc + 1

            lErro = Comando_Executar(alComando(1), "INSERT INTO TRVAcordoComissao (NumIntDoc, NumIntAcordo, Seq, Produto, Destino, PercComissao) VALUES (?,?,?,?,?,?)", _
                     objTRVAcordoComiss.lNumIntDoc, objTRVAcordoComiss.lNumIntAcordo, objTRVAcordoComiss.iSeq, objTRVAcordoComiss.sProduto, objTRVAcordoComiss.iDestino, objTRVAcordoComiss.dPercComissao)
            If lErro <> AD_SQL_SUCESSO Then gError 197116

        Next

    End If

    iIndice = 0

    If objTRVAcordo.colTRVAcordoDif.Count > 0 Then

        lErro = CF("Config_ObterNumInt", "TRVConfig", TRVCONFIG_NUM_INT_PROX_TRVACORDODIF, lNumIntDoc, objTRVAcordo.colTRVAcordoDif.Count)
        If lErro <> SUCESSO Then gError 197117

        For Each objTRVAcordoDif In objTRVAcordo.colTRVAcordoDif
        
            iIndice = iIndice + 1
            
            objTRVAcordoDif.lNumIntDoc = lNumIntDoc
            objTRVAcordoDif.lNumIntAcordo = objTRVAcordo.lNumIntDoc
            objTRVAcordoDif.iSeq = iIndice
            
            lNumIntDoc = lNumIntDoc + 1

            lErro = Comando_Executar(alComando(2), "INSERT INTO TRVAcordoTarifaDif (NumIntDoc, NumIntAcordo, Seq, Produto, Preco, Moeda, Destino,PercComissao) VALUES (?,?,?,?,?,?,?,?)", _
                     objTRVAcordoDif.lNumIntDoc, objTRVAcordoDif.lNumIntAcordo, objTRVAcordoDif.iSeq, objTRVAcordoDif.sProduto, objTRVAcordoDif.dPreco, objTRVAcordoDif.iMoeda, objTRVAcordoDif.iDestino, objTRVAcordoDif.dPercComissao)
            If lErro <> AD_SQL_SUCESSO Then gError 197118

        Next

    End If

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TRVAcordoComissaoTarifa_Inclui = SUCESSO
    
    Exit Function

Erro_TRVAcordoComissaoTarifa_Inclui:

    TRVAcordoComissaoTarifa_Inclui = gErr

    Select Case gErr

        Case 197114
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 197115, 197117

        Case 197116
            Call Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_TRVACORDOCOMISSAO", gErr)

        Case 197118
            Call Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_TRVACORDOTARIFADIF", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197119)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function TRVAcordo_Exclui(ByVal objTRVAcordo As ClassTRVAcordos) As Long

Dim lErro As Long
Dim alComando(1 To 1) As Long
Dim iIndice As Integer
Dim lTransacao As Long


On Error GoTo Erro_TRVAcordo_Exclui

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 197139

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 197140
    Next

    lErro = Comando_Executar(alComando(1), "DELETE FROM TRVAcordos WHERE Numero = ?", objTRVAcordo.lNumero)
    If lErro <> AD_SQL_SUCESSO Then gError 197141

    lErro = TRVAcordoComissaoTarifa_Exclui(objTRVAcordo)
    If lErro <> SUCESSO Then gError 197142

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 197143

    TRVAcordo_Exclui = SUCESSO

    Exit Function

Erro_TRVAcordo_Exclui:

    TRVAcordo_Exclui = gErr

    Select Case gErr

        Case 197139
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 197140
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 197141
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRVACORDOS", gErr)

        Case 197142
        
        Case 197143
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197144)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Desfaz Transação
    Call Transacao_Rollback
    
    Exit Function

End Function

Public Function TRVVoucher_Gera_Comissao(ByVal objVou As ClassTRVVouchers, ByVal dValorBase As Double, ByVal dtData As Date, ByVal sHistorico As String, ByVal iTipoTratamento As Integer, Optional ByVal iStatus As Integer = TRV_VOU_INFO_STATUS_BLOQUEADO, Optional iTipoOcr As Integer = 0) As Long
'Gera a comissão externa relacionada ao voucher

Dim lErro As Long
Dim alComando(0 To 22) As Long
Dim iIndice As Integer
Dim lSeq As Long
Dim iTipoLiberacao As Integer
Dim lNumIntDocLiberacao As Long

On Error GoTo Erro_TRVVoucher_Gera_Comissao

    'Se o voucher foi emitido antes da implantação do comissionamento ou ainda não é a versão de comissionamento
    'se sim, não faz nada
    If gdtDataInicioComisCorp <= objVou.dtData And giVersaoTRV >= TRV_VERSAO_COMISSIONAMENTO Then

        'Abertura de Comando
        For iIndice = LBound(alComando) To UBound(alComando)
            alComando(iIndice) = Comando_Abrir()
            If alComando(iIndice) = 0 Then gError 196280
        Next
        
        lErro = CF("TRVVouchers_Le", objVou)
        If lErro <> SUCESSO And lErro <> erro_leitura_sem_dado Then gError 9999
        
        If lErro <> ERRO_LEITURA_SEM_DADOS Then gError 999999
        
        lErro = TRVVoucher_Gera_Comissao_Obtem_Seq(alComando(0), objVou, lSeq)
        If lErro <> SUCESSO Then gError 9999999
        
        lErro = TRVVoucher_Gera_Comissao_Obtem_Status(alComando(1), objVou, iStatus, iTipoLiberacao, lNumIntDocLiberacao)
        If lErro <> SUCESSO Then gError 9999999
        
        'Se é um voucher novo
        If iTipoTratamento = TRV_TIPO_TRATAMENTO_COMI_VOU_NOVO Then
        
            'Insere as informações iniciais do voucher
            lErro = TRVVoucher_Gera_Comissao_Insere_Novo()
            If lErro <> SUCESSO Then gError 9999999
        
        End If
        
        
        
        
        
        'Fecha Comando
        For iIndice = LBound(alComando) To UBound(alComando)
            Call Comando_Fechar(alComando(iIndice))
        Next
    
    End If
 
    TRVVoucher_Gera_Comissao = SUCESSO

    Exit Function

Erro_TRVVoucher_Gera_Comissao:

    TRVVoucher_Gera_Comissao = gErr

    Select Case gErr

        Case 196280
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 196281, 196282

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196311)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Private Function TRVVoucher_Gera_Comissao_Obtem_Seq(ByVal lComando As Long, ByVal objVou As ClassTRVVouchers, lContador As Long) As Long
'Gera a comissão externa relacionada ao voucher

Dim lErro As Long

On Error GoTo Erro_TRVVoucher_Gera_Comissao_Obtem_Seq

     'Le o sequencial do histórico do Voucher
    lErro = Comando_Executar(lComando, "SELECT MAX(Seq) FROM TRVVoucherInfo WHERE TipVou = ? AND Serie = ? AND NumVou = ? ", _
                lContador, objVou.sTipVou, objVou.sSerie, objVou.lNumVou)
    If lErro <> AD_SQL_SUCESSO Then gError 196301

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196302
    
    If lErro <> AD_SQL_SUCESSO Then lContador = 0
    
    lContador = lContador + 1
 
    TRVVoucher_Gera_Comissao_Obtem_Seq = SUCESSO

    Exit Function

Erro_TRVVoucher_Gera_Comissao_Obtem_Seq:

    TRVVoucher_Gera_Comissao_Obtem_Seq = gErr

    Select Case gErr

        Case 196301, 196302
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVVOUCHERINFO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196311)

    End Select

    Exit Function

End Function

Private Function TRVVoucher_Gera_Comissao_Obtem_Status(ByVal lComando As Long, ByVal objVou As ClassTRVVouchers, iStatus As Integer, iTipoLiberacao As Integer, lNumIntDocLiberacao As Long) As Long
'Gera a comissão externa relacionada ao voucher

Dim lErro As Long
Dim lNumIntDocBaixa As Long

On Error GoTo Erro_TRVVoucher_Gera_Comissao_Obtem_Status

    'Se era para gerar a comissão bloqueada e o voucher já tiver faturado, verifica se o título já está baixado, se sim muda
    'o status para liberado e guarda as informações de o que liberou
    If iStatus = TRV_VOU_INFO_STATUS_BLOQUEADO And objVou.lNumIntDocDestino <> 0 And objVou.iTipoDocDestino = TRV_TIPO_DOC_DESTINO_TITREC Then
        
        'Le o Voucher
        lErro = Comando_Executar(lComando, "SELECT B.NumIntDoc FROM ParcelasRecTodas AS P, BaixasParcRec AS B WHERE P.NumIntTiTulo = ? AND P.NumIntDoc = B.NumIntParcela AND B.Status <> ? ", _
                    lNumIntDocBaixa, objVou.lNumIntDocDestino, STATUS_EXCLUIDO)
        If lErro <> AD_SQL_SUCESSO Then gError 192367
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192368
    
        If lErro = AD_SQL_SUCESSO Then
            iStatus = TRV_VOU_INFO_STATUS_LIBERADO
            iTipoLiberacao = TRV_TIPO_LIBERACAO_COMISSAO_BAIXA
            lNumIntDocLiberacao = lNumIntDocBaixa
        End If
    
    End If
 
    TRVVoucher_Gera_Comissao_Obtem_Status = SUCESSO

    Exit Function

Erro_TRVVoucher_Gera_Comissao_Obtem_Status:

    TRVVoucher_Gera_Comissao_Obtem_Status = gErr

    Select Case gErr

        Case 192367, 192368
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BAIXASPARCREC", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196311)

    End Select

    Exit Function

End Function

Private Function TRVVoucher_Gera_Comissao_Insere_Novo(ByVal objVou As ClassTRVVouchers, ByVal dValorBase As Double, ByVal dtData As Date, ByVal sHistorico As String, ByVal iTipoTratamento As Integer, Optional ByVal iStatus As Integer = TRV_VOU_INFO_STATUS_BLOQUEADO, Optional iTipoOcr As Integer = 0) As Long
'Gera a comissão externa relacionada ao voucher

Dim lErro As Long
Dim lCodDestino As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim lCliCMR As Long
Dim dPercCMR As Double
Dim lCliCMC As Long
Dim dPercCMC As Double
Dim dPercAgencia As Double
Dim dPercOVER As Double
Dim lNumVou As Long

On Error GoTo Erro_TRVVoucher_Gera_Comissao_Insere_Novo
   
    'Le as informações padrões, depois acerta as exceções
    'Le o representante e o correntista e o percentual de comissao
    lErro = Comando_Executar(lComando1, "SELECT Representante, ComissaoRep, Correntista, ComissaoCorr, ComissaoAg FROM FiliaisClientesTRV WHERE CodCliente = ? ", _
                lCliCMR, dPercCMR, lCliCMC, dPercCMC, dPercAgencia, objVou.lClienteVou)
    If lErro <> AD_SQL_SUCESSO Then gError 196281

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196282
    
    If lErro = AD_SQL_SUCESSO Then
        objVou.lRepresentante = lCliCMR
        objVou.dComissaoRep = dPercCMR
        objVou.lCorrentista = lCliCMC
        objVou.dComissaoCorr = dPercCMC
        objVou.dComissaoAg = dPercAgencia
    End If
    
    lErro = CF("Produto_Formata", objVou.sProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 192650
    
    'Exceção de Representante por produto
    'Le o representante e o percentual de comissao
    lErro = Comando_Executar(lComando2, "SELECT Representante, PercComis FROM TRVClienteRepProd WHERE Cliente = ? AND Produto = ? ", _
                lCliCMR, dPercCMR, objVou.lClienteVou, sProdutoFormatado)
    If lErro <> AD_SQL_SUCESSO Then gError 196281

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando2)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196282
    
    If lErro = AD_SQL_SUCESSO Then
        objVou.lRepresentante = lCliCMR
        objVou.dComissaoRep = dPercCMR
    End If
    
    'Exceção de Correntista por produto
    'Le o correntista e o percentual de comissao
    lErro = Comando_Executar(lComando3, "SELECT Correntista, PercComis FROM TRVClienteCorProd WHERE Cliente = ? AND Produto = ? ", _
                lCliCMC, dPercCMC, objVou.lClienteVou, sProdutoFormatado)
    If lErro <> AD_SQL_SUCESSO Then gError 196281

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando3)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196282
    
    If lErro = AD_SQL_SUCESSO Then
        objVou.lCorrentista = lCliCMC
        objVou.dComissaoCorr = dPercCMC
    End If
    
    'Le o destino do voucher e pega o valor do campo no corporator
    lErro = TRVVoucher_Gera_Comissao_Obtem_Destino(lComando4, lComando5, objVou, lCodDestino)
    If lErro <> SUCESSO Then gError 99999
    
    'Percentual da agência exceção nível 1
    'Le o percentual de comissão da agência no acordo
    lErro = Comando_Executar(lComando6, "SELECT C.PercComissao FROM TRVAcordoComissao AS C, TRVAcordos AS A WHERE C.NumIntAcordo = A.NumIntDoc AND A.Cliente = ? AND A.FilialCliente = ? AND (A.ValidadeDe >= ? OR A.ValidadeDe = ?) AND (A.ValidadeAte <= ? OR A.ValidadeAte = ?) AND C.Produto = ? ", _
                dPercAgencia, objVou.lClienteVou, FILIAL_MATRIZ, objVou.dtData, DATA_NULA, objVou.dtData, DATA_NULA, sProdutoFormatado)
    If lErro <> AD_SQL_SUCESSO Then gError 196356

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando6)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196357
    
    'Se existir acordo o percentual de comissão da agência vem do acordo
    If lErro = AD_SQL_SUCESSO Then
        objVou.dComissaoAg = dPercAgencia
    End If

    'Percentual da agência exceção nível 2
    'Le a tabela diferenciada por produto\destino
    lErro = Comando_Executar(lComando7, "SELECT T.PercComissao FROM TRVAcordoTarifaDif AS T WHERE T.NumIntAcordo = A.NumIntDoc AND A.Cliente = ? AND A.FilialCliente = ? AND (A.ValidadeDe >= ? OR A.ValidadeDe = ?) AND (A.ValidadeAte <= ? OR A.ValidadeAte = ?) AND T.Produto = ? AND Destino = ? ", _
                dPercAgencia, objVou.lClienteVou, FILIAL_MATRIZ, objVou.dtData, DATA_NULA, objVou.dtData, DATA_NULA, sProdutoFormatado, lCodDestino)
    If lErro <> AD_SQL_SUCESSO Then gError 196358

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando7)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196359
    
    'Se tem tarifa diferenciada não tem comissão para agência
    If lErro = AD_SQL_SUCESSO Then
        objVou.dComissaoAg = dPercAgencia
    End If
      
    'Se for venda com cartão a comissão deve ser repassada para agência
    If objVou.iCartao = MARCADO Then
        objVou.dValorBruto = objVou.dValor
    Else
        'O valor bruto é o valor líquido acrescido da comissão descontada
        objVou.dValorBruto = objVou.dValor / (1 - objVou.dComissaoAg)
    End If

    'Se tem emissor
    If objVou.lEmissor <> 0 Then
    
        'Le o Emissor e sua comissão
        lErro = Comando_Executar(lComando8, "SELECT NumIntDoc, PercComissao FROM TRVCliEmissores WHERE Cliente = ? AND FornEmissor = ? ", _
                     lNumIntDocCliEmi, dPercOVER, objVou.lClienteVou, objVou.lEmissor)
        If lErro <> AD_SQL_SUCESSO Then gError 196289
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(lComando8)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196290
    
        If lErro <> AD_SQL_SUCESSO Then gError 196291 'Esse emissor não está associado a esse cliente
        
        objVou.dComissaoEmissor = dPercOVER
        
        'Over exceção nível 1
        'Le as exceções
        lErro = Comando_Executar(lComando9, "SELECT PercComissao FROM TRVCliEmissoresExc WHERE NumIntDocCliEmi = ? AND Produto = ? ", _
                     dPercOVER, lNumIntDocCliEmi, sProdutoFormatado)
        If lErro <> AD_SQL_SUCESSO Then gError 196292
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(lComando9)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196293
            
        If lErro = AD_SQL_SUCESSO Then
            objVou.dComissaoEmissor = dPercOVER
        End If

    End If
    
    
    'Le o voucher para o update
    lErro = Comando_Executar(lComando10, "SELECT NumVou FROM TRVVouchers WHERE NumVou = ? AND TipoDoc = ? AND  Serie = ? AND TipVou= ?  ", _
                 lNumVou, objVou.lNumVou, objVou.sTipoDoc, objVou.sSerie, objVou.sTipVou)
    If lErro <> AD_SQL_SUCESSO Then gError 196292

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando10)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196293
    
    'Trava registro contra alterações/Leituras
    lErro = Comando_LockExclusive(lComando10)
    If lErro <> SUCESSO Then gError 196311
    
    'Marca as informações pertinentes a comissão no voucher, a partir daí esses
    'serão os valores a serem considerados mesmo se houverem alterações no cadastro de clientes
    '(informações históricas - qual era o representante na emissão do voucher, qual o % de comissão, etc ...)
    lErro = Comando_ExecutarPos(lComando11, "UPDATE TRVVouchers SET Representante = ?, ComissaoRep = ?, Correntista = ?, ComissaoCorr = ?, ComissaoAg = ?, ComissaoEmissor = ?, ValorBruto = ? ", _
                lComando10, objVou.lRepresentante, objVou.dComissaoRep, objVou.lCorrentista, objVou.dComissaoCorr, objVou.dComissaoAg, objVou.dComissaoEmissor, objVou.dValorBruto)
    If lErro <> AD_SQL_SUCESSO Then gError 196294
 
    TRVVoucher_Gera_Comissao_Insere_Novo = SUCESSO

    Exit Function

Erro_TRVVoucher_Gera_Comissao_Insere_Novo:

    TRVVoucher_Gera_Comissao_Insere_Novo = gErr

    Select Case gErr


        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196311)

    End Select

    Exit Function

End Function

Private Function TRVVoucher_Gera_Comissao_Obtem_Destino(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal objVou As ClassTRVVouchers, lCodDestino As Long) As Long
'Gera a comissão externa relacionada ao voucher

Dim lErro As Long
Dim sDestino As String

On Error GoTo Erro_TRVVoucher_Gera_Comissao_Obtem_Destino

    sDestino = String(255, 0)
    
    'Le o Voucher (Informações extraídas) para pegar o destino
    lErro = Comando_Executar(lComando1, "SELECT Destino FROM TRVVoucherInfoSigav WHERE Tipo = ? AND Serie = ? AND NumVou = ? ", _
                 sDestino, objVou.sTipVou, objVou.sSerie, objVou.lNumVou)
    If lErro <> AD_SQL_SUCESSO Then gError 196286

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196287

    If lErro <> AD_SQL_SUCESSO Then gError 196288 ' Não houve a importação e por isso não tem de onde tirar o emissor
        
    'Le o código do destino
    lErro = Comando_Executar(lComando2, "SELECT CodValor FROM CamposGenericosValores WHERE CodCampo = ? AND Valor = ? ", _
                lCodDestino, CAMPOSGENERICOS_DESTINO_VIAGEM, Trim(Left(sDestino, 9)))
    If lErro <> AD_SQL_SUCESSO Then gError 196350

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando2)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196351
    
    If lErro <> SUCESSO Then gError 196352 'Destino não cadastrado
 
    TRVVoucher_Gera_Comissao_Obtem_Destino = SUCESSO

    Exit Function

Erro_TRVVoucher_Gera_Comissao_Obtem_Destino:

    TRVVoucher_Gera_Comissao_Obtem_Destino = gErr

    Select Case gErr

        Case 196288
            Call Rotina_Erro(vbOKOnly, "ERRO_VOUCHER_SEM_DADOS_SIGAV", gErr)
            
        Case 196286, 196287
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVVOUCHERINFOSIGAV", gErr)

        Case 196350, 196351
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CAMPOSGENERICOSVALORES", gErr)
            
        Case 196352
            Call Rotina_Erro(vbOKOnly, "ERRO_TRV_DESTINO_NAO_CADASTRADO", gErr, Trim(Left(sDestino, 9)))

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196311)

    End Select

    Exit Function

End Function

Public Function TRVVoucher_Gera_Comissao2(ByVal objVou As ClassTRVVouchers, ByVal dValorBase As Double, ByVal dtData As Date, ByVal sHistorico As String, ByVal iTipoTratamento As Integer, Optional ByVal iStatus As Integer = TRV_VOU_INFO_STATUS_BLOQUEADO) As Long
'Gera a comissão externa relacionada ao voucher

Dim lErro As Long
Dim alComando(0 To 22) As Long
Dim iIndice As Integer
Dim lCliCMR As Long
Dim dPercCMR As Double
Dim lCliCMC As Long
Dim dPercCMC As Double
Dim lFornOVER As Long
Dim dPercOVER As Double
Dim dPercAgencia As Double
Dim dValorBaseBruto As Double
Dim dValorCMR As Double
Dim dValorCMC As Double
Dim dValorOVER As Double
Dim iCartao As Integer
Dim lNumIntDocCliEmi As Long
Dim dPercOVERExc As Double
Dim lNumIntDoc As Long
Dim sTipoDoc As String
Dim lContador As Long
Dim dValorBrutoBase As Double
Dim lCliente As Long
Dim sDestino As String
Dim lCodDestino As Long
Dim lNumIntAcordo As Long
Dim lNumIntTarfDif As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim lNumIntDocBaixa As Long
Dim dPercAgenciaExc As Double
Dim sConteudo As String
Dim sItem As String
Dim iStatusOver As Integer
Dim idiasantc As Integer
Dim dValorBrutoVou As Double
Dim dValorComiAgencia As Double
Dim iTipoLiberacao As Integer
Dim lNumIntDocLiberacao As Long
Dim iTipoLiberacaoOver As Integer
Dim lNumIntDocLiberacaoOver As Long
Dim lNumIntDocInfoAux As Long
Dim sTipoDocInfoAux As String

On Error GoTo Erro_TRVVoucher_Gera_Comissao

    'Se o voucher foi emitido antes da implantação do comissionamento ou ainda não é a versão de comissionamento
    'se sim, não faz nada
    If gdtDataInicioComisCorp <= objVou.dtData And giVersaoTRV >= TRV_VERSAO_COMISSIONAMENTO Then

        'Abertura de Comando
        For iIndice = LBound(alComando) To UBound(alComando)
            alComando(iIndice) = Comando_Abrir()
            If alComando(iIndice) = 0 Then gError 196280
        Next
        
         'Le o sequencial do histórico do Voucher
        lErro = Comando_Executar(alComando(11), "SELECT MAX(Seq) FROM TRVVoucherInfo WHERE TipVou = ? AND Serie = ? AND NumVou = ? ", _
                    lContador, objVou.sTipVou, objVou.sSerie, objVou.lNumVou)
        If lErro <> AD_SQL_SUCESSO Then gError 196301
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(11))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196302
        
        'Se era para gerar a comissão bloqueada e o voucher já tiver faturado, verifica se o título já está baixado, se sim muda
        'o status para liberado
        If iStatus = TRV_VOU_INFO_STATUS_BLOQUEADO And objVou.lNumIntDocDestino <> 0 And objVou.iTipoDocDestino = TRV_TIPO_DOC_DESTINO_TITREC Then
            
            'Le o Voucher
            lErro = Comando_Executar(alComando(16), "SELECT B.NumIntDoc FROM ParcelasRecTodas AS P, BaixasParcRec AS B WHERE P.NumIntTiTulo = ? AND P.NumIntDoc = B.NumIntParcela AND B.Status <> ? ", _
                        lNumIntDocBaixa, objVou.lNumIntDocDestino, STATUS_EXCLUIDO)
            If lErro <> AD_SQL_SUCESSO Then gError 192367
        
            'Busca Primeiro
            lErro = Comando_BuscarPrimeiro(alComando(16))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192368
        
            If lErro = AD_SQL_SUCESSO Then
                iStatus = TRV_VOU_INFO_STATUS_LIBERADO
                iTipoLiberacao = TRV_TIPO_LIBERACAO_COMISSAO_BAIXA
                lNumIntDocLiberacao = lNumIntDocBaixa
            End If
        
        End If
    
        'Se é um voucher novo
        If iTipoTratamento = TRV_TIPO_TRATAMENTO_COMI_VOU_NOVO Then
            
            sDestino = String(255, 0)
            
            'Le o Voucher (Informações extraídas) para pegar o emissor, o cliente e o destino
            lErro = Comando_Executar(alComando(4), "SELECT FornEmissor, Cliente, Destino FROM TRVVoucherInfoSigav WHERE Tipo = ? AND Serie = ? AND NumVou = ? ", _
                         lFornOVER, lCliente, sDestino, objVou.sTipVou, objVou.sSerie, objVou.lNumVou)
            If lErro <> AD_SQL_SUCESSO Then gError 196286
        
            'Busca Primeiro
            lErro = Comando_BuscarPrimeiro(alComando(4))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196287
        
            If lErro <> AD_SQL_SUCESSO Then gError 196288 ' Não houve a importação e por isso não tem de onde tirar o emissor
            
            'Le o representante e o correntista e o percentual de comissao
            lErro = Comando_Executar(alComando(0), "SELECT Representante, ComissaoRep, Correntista, ComissaoCorr, ComissaoAg FROM FiliaisClientesTRV WHERE CodCliente = ? ", _
                        lCliCMR, dPercCMR, lCliCMC, dPercCMC, dPercAgencia, lCliente)
            If lErro <> AD_SQL_SUCESSO Then gError 196281
        
            'Busca Primeiro
            lErro = Comando_BuscarPrimeiro(alComando(0))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196282
            
            'Le o código do destino
            lErro = Comando_Executar(alComando(12), "SELECT CodValor FROM CamposGenericosValores WHERE CodCampo = ? AND Valor = ? ", _
                        lCodDestino, CAMPOSGENERICOS_DESTINO_VIAGEM, Trim(Left(sDestino, 9)))
            If lErro <> AD_SQL_SUCESSO Then gError 196350
        
            'Busca Primeiro
            lErro = Comando_BuscarPrimeiro(alComando(12))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196351
            
            If lErro <> SUCESSO Then gError 196352 'Destino não cadastrado
    
            'Le o acordo
            lErro = Comando_Executar(alComando(13), "SELECT NumIntDoc FROM TRVAcordos AS A WHERE A.Cliente = ? AND A.FilialCliente = ? AND (A.ValidadeDe >= ? OR A.ValidadeDe = ?) AND (A.ValidadeAte <= ? OR A.ValidadeAte = ?) ", _
                        lNumIntAcordo, lCliente, FILIAL_MATRIZ, objVou.dtData, DATA_NULA, objVou.dtData, DATA_NULA)
            If lErro <> AD_SQL_SUCESSO Then gError 196353
        
            'Busca Primeiro
            lErro = Comando_BuscarPrimeiro(alComando(13))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196354
            
            'Se existe o acordo verifica se tem tabela diferenciada ou comissão diferenciada
            If lErro = AD_SQL_SUCESSO Then
            
                lErro = CF("Produto_Formata", objVou.sProduto, sProdutoFormatado, iProdutoPreenchido)
                If lErro <> SUCESSO Then gError 196355
        
                'Le o percentual de comissão da agência no acordo
                lErro = Comando_Executar(alComando(14), "SELECT PercComissao FROM TRVAcordoComissao AS C WHERE C.NumIntAcordo = ? AND (C.Destino = 0 OR C.Destino = ?) AND C.Produto = ? ", _
                            dPercAgenciaExc, lNumIntAcordo, lCodDestino, sProdutoFormatado)
                If lErro <> AD_SQL_SUCESSO Then gError 196356
            
                'Busca Primeiro
                lErro = Comando_BuscarPrimeiro(alComando(14))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196357
                
                'Se existir acordo o percentual de comissão da agência vem do acordo
                If lErro = AD_SQL_SUCESSO Then
                    dPercAgencia = dPercAgenciaExc
                End If
    
                'Le a tabela diferenciada
                lErro = Comando_Executar(alComando(15), "SELECT NumIntDoc FROM TRVAcordoTarifaDif AS T WHERE T.NumIntAcordo = ? AND T.Produto = ? ", _
                            lNumIntTarfDif, lNumIntAcordo, sProdutoFormatado)
                If lErro <> AD_SQL_SUCESSO Then gError 196358
            
                'Busca Primeiro
                lErro = Comando_BuscarPrimeiro(alComando(15))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196359
                
                'Se tem tarifa diferenciada não tem comissão para agência
                If lErro = AD_SQL_SUCESSO Then
                    dPercAgencia = 0
                End If
            
            End If
                           
            'Le o Voucher que vai ter as informações atualizadas
            lErro = Comando_ExecutarPos(alComando(1), "SELECT Cartao, DiasAntc FROM TRVVouchers WHERE TipoDoc = ? AND TipVou = ? AND Serie = ? AND NumVou = ? ", _
                        0, iCartao, idiasantc, objVou.sTipoDoc, objVou.sTipVou, objVou.sSerie, objVou.lNumVou)
            If lErro <> AD_SQL_SUCESSO Then gError 196283
        
            'Busca Primeiro
            lErro = Comando_BuscarPrimeiro(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196284
            
            If lErro <> AD_SQL_SUCESSO Then gError 196285 'Voucher não cadastrado
                           
            'Se for venda com cartão a comissão deve ser repassada para agência
            'REGRA USADA
            '1 - tarifa diferenciada = 0% de comissão para agência
            '2 - Tem acordo válido para o produto\destino = comissão do acordo
            '3 - Demais = comissão do cadastro do cliente
            If iCartao = MARCADO Then
                dValorBrutoBase = dValorBase
            Else
                'O valor bruto é o valor líquido acrescido da comissão descontada
                dValorBrutoBase = dValorBase / (1 - dPercAgencia)
            End If
            
            dValorBrutoVou = dValorBrutoBase
    
            'Se tem emissor
            If lFornOVER <> 0 Then
            
                'Le o Emissor
                lErro = Comando_Executar(alComando(5), "SELECT NumIntDoc, PercComissao FROM TRVCliEmissores WHERE Cliente = ? AND FornEmissor = ? ", _
                             lNumIntDocCliEmi, dPercOVER, lCliente, lFornOVER)
                If lErro <> AD_SQL_SUCESSO Then gError 196289
            
                'Busca Primeiro
                lErro = Comando_BuscarPrimeiro(alComando(5))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196290
            
                If lErro <> AD_SQL_SUCESSO Then gError 196291 'Esse emissor não está associado a esse cliente
                
                'Le as exceções
                lErro = Comando_Executar(alComando(6), "SELECT PercComissao FROM TRVCliEmissoresExc WHERE NumIntDocCliEmi = ? AND Produto = ? ", _
                             dPercOVERExc, lNumIntDocCliEmi, objVou.sProduto)
                If lErro <> AD_SQL_SUCESSO Then gError 196292
            
                'Busca Primeiro
                lErro = Comando_BuscarPrimeiro(alComando(6))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196293
                    
                If lErro = AD_SQL_SUCESSO Then
                    dPercOVER = dPercOVERExc
                End If
    
            End If
            
            'Trava registro contra alterações/Leituras
            lErro = Comando_LockExclusive(alComando(1))
            If lErro <> SUCESSO Then gError 196311
            
            'Marca as informações pertinentes a comissão no voucher, a partir daí esses
            'serão os valores a serem considerados mesmo se houverem alterações no cadastro de clientes
            '(informações históricas - qual era o representante na emissão do voucher, qual o % de comissão, etc ...)
            lErro = Comando_ExecutarPos(alComando(2), "UPDATE TRVVouchers SET Representante = ?, ComissaoRep = ?, Correntista = ?, ComissaoCorr = ?, ComissaoAg = ?, Emissor = ?, ComissaoEmissor = ?, ValorBruto = ?, ClienteVou = ? ", _
                        alComando(1), lCliCMR, dPercCMR, lCliCMC, dPercCMC, dPercAgencia, lFornOVER, dPercOVER, dValorBrutoBase, lCliente)
            If lErro <> AD_SQL_SUCESSO Then gError 196294
    
        Else
            
            'Le o Voucher
            lErro = Comando_Executar(alComando(3), "SELECT DiasAntc, Cartao, Representante, ComissaoRep, Correntista, ComissaoCorr, ComissaoAg, Emissor, ComissaoEmissor, ValorBruto FROM TRVVouchers WHERE TipoDoc = ? AND TipVou = ? AND Serie = ? AND NumVou = ? ", _
                         idiasantc, iCartao, lCliCMR, dPercCMR, lCliCMC, dPercCMC, dPercAgencia, lFornOVER, dPercOVER, dValorBrutoVou, objVou.sTipoDoc, objVou.sTipVou, objVou.sSerie, objVou.lNumVou)
            If lErro <> AD_SQL_SUCESSO Then gError 196299
        
            'Busca Primeiro
            lErro = Comando_BuscarPrimeiro(alComando(3))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196300
            
            'PS-> Não pode ler o ValorBruto no voucher, pois nesse caso não é a insersão dele por inteiro
    
            If iCartao = MARCADO Then
                dValorBrutoBase = dValorBase
            Else
                'O valor bruto é o valor líquido acrescido da comissão descontada
                dValorBrutoBase = dValorBase / (1 - dPercAgencia)
            End If
            
        End If
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        'Se for tratamento para OCR ou NVL
        If iTipoTratamento = TRV_TIPO_TRATAMENTO_COMI_OCR Or iTipoTratamento = TRV_TIPO_TRATAMENTO_COMI_NVL Then
    
            If iTipoTratamento = TRV_TIPO_TRATAMENTO_COMI_OCR Then
                sTipoDoc = TRV_TIPODOC_OCR_TEXTO
            Else
                sTipoDoc = TRV_TIPODOC_NVL_TEXTO
            End If
    
            'Obter NumIntDoc
            lErro = CF("Config_ObterNumInt", "TRVConfig", "NUM_INT_PROX_TRVVOURCHERINFO", lNumIntDoc)
            If lErro <> SUCESSO Then gError 196309
            
            lContador = lContador + 1
            
            'Insere o OCR ou NVL
            lErro = Comando_Executar(alComando(7), "INSERT INTO TRVVoucherINFO (NumIntDoc,TipoDoc,TipVou,Serie,NumVou,Seq,Data,TipoDocDestino,NumIntDocDestino,NumIntDocComiInt,Valor,Historico,NumTitulo,Status,TipoLiberacao,NumIntDocLiberacao) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) ", _
                         lNumIntDoc, sTipoDoc, objVou.sTipVou, objVou.sSerie, objVou.lNumVou, lContador, dtData, 0, 0, 0, Arredonda_Moeda(dValorBrutoBase), sHistorico, 0, iStatus, iTipoLiberacao, lNumIntDocLiberacao)
            If lErro <> AD_SQL_SUCESSO Then gError 196310
                
        End If
        
        'Se mexeu no valor do voucher tem que inserir o valor bruto e a comissão
        If iTipoTratamento = TRV_TIPO_TRATAMENTO_COMI_VOU_ALTERADO Or iTipoTratamento = TRV_TIPO_TRATAMENTO_COMI_VOU_NOVO Then
        
            'Obter NumIntDoc
            lErro = CF("Config_ObterNumInt", "TRVConfig", "NUM_INT_PROX_TRVVOURCHERINFO", lNumIntDoc)
            If lErro <> SUCESSO Then gError 196295
            
            lContador = lContador + 1
            
            'Insere o BRUTO
            lErro = Comando_Executar(alComando(8), "INSERT INTO TRVVoucherInfo (NumIntDoc,TipoDoc,TipVou,Serie,NumVou,Seq,Data,TipoDocDestino,NumIntDocDestino,NumIntDocComiInt,Valor,Historico,NumTitulo, Status,TipoLiberacao,NumIntDocLiberacao) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) ", _
                         lNumIntDoc, TRV_TIPODOC_BRUTO_TEXTO, objVou.sTipVou, objVou.sSerie, objVou.lNumVou, lContador, dtData, 0, 0, 0, Arredonda_Moeda(dValorBrutoBase), sHistorico, 0, iStatus, iTipoLiberacao, lNumIntDocLiberacao)
            If lErro <> AD_SQL_SUCESSO Then gError 196296
            
        End If
        
        'Se mexeu no valor do voucher ou gerou uma OCR ou NVL tem que acertar a comissão
        If iTipoTratamento = TRV_TIPO_TRATAMENTO_COMI_OCR Or iTipoTratamento = TRV_TIPO_TRATAMENTO_COMI_NVL Or iTipoTratamento = TRV_TIPO_TRATAMENTO_COMI_VOU_ALTERADO Or iTipoTratamento = TRV_TIPO_TRATAMENTO_COMI_VOU_NOVO Then
            
            'Obter NumIntDoc
            lErro = CF("Config_ObterNumInt", "TRVConfig", "NUM_INT_PROX_TRVVOURCHERINFO", lNumIntDoc)
            If lErro <> SUCESSO Then gError 196297
            
            lContador = lContador + 1
            
            If iCartao = DESMARCADO Then
                'Insere a CMA (Comissão da agência)
                If Abs(Arredonda_Moeda(dValorBrutoBase - dValorBase)) > DELTA_VALORMONETARIO Then
                    lErro = Comando_Executar(alComando(9), "INSERT INTO TRVVoucherInfo (NumIntDoc,TipoDoc,TipVou,Serie,NumVou,Seq,Data,TipoDocDestino,NumIntDocDestino,NumIntDocComiInt,Valor,Historico,NumTitulo, Status,TipoLiberacao,NumIntDocLiberacao) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) ", _
                                 lNumIntDoc, TRV_TIPODOC_CMA_TEXTO, objVou.sTipVou, objVou.sSerie, objVou.lNumVou, lContador, dtData, 0, 0, 0, Arredonda_Moeda(dValorBrutoBase - dValorBase), sHistorico, 0, iStatus, iTipoLiberacao, lNumIntDocLiberacao)
                End If
            Else
                If Abs(Arredonda_Moeda(dValorBrutoBase * dPercAgencia)) > DELTA_VALORMONETARIO Then
                    'Insere a CMCC (Comissão da agência em venda sobre cartão de crédito)
                    lErro = Comando_Executar(alComando(9), "INSERT INTO TRVVoucherInfo (NumIntDoc,TipoDoc,TipVou,Serie,NumVou,Seq,Data,TipoDocDestino,NumIntDocDestino,NumIntDocComiInt,Valor,Historico,NumTitulo, Status,TipoLiberacao,NumIntDocLiberacao) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) ", _
                                 lNumIntDoc, TRV_TIPODOC_CMCC_TEXTO, objVou.sTipVou, objVou.sSerie, objVou.lNumVou, lContador, dtData, 0, 0, 0, Arredonda_Moeda(dValorBrutoBase * dPercAgencia), sHistorico, 0, iStatus, iTipoLiberacao, lNumIntDocLiberacao)
                End If
            End If
            If lErro <> AD_SQL_SUCESSO Then gError 196298
            
        End If
        
        'Se o voucher passou a ser de cartão ou deixou de ser
        If iTipoTratamento = TRV_TIPO_TRATAMENTO_COMI_VOU_TRANSF_CARTAO Then
            
            'Deixou de ser cartão
            If iCartao = DESMARCADO Then
            
                dValorComiAgencia = 0
            
                'Le a TRVVOucherInfo para pegar o saldo do que foi lançado de comissão de cartão para estornar e lançar para agência
                lErro = Comando_Executar(alComando(18), "SELECT SUM(Valor) FROM TRVVoucherInfo WHERE TipoDoc = ? AND TipVou = ? AND Serie = ? AND NumVou = ? ", _
                             dValorComiAgencia, TRV_TIPODOC_CMCC_TEXTO, objVou.sTipVou, objVou.sSerie, objVou.lNumVou)
                If lErro <> AD_SQL_SUCESSO Then gError 196299
            
                'Busca Primeiro
                lErro = Comando_BuscarPrimeiro(alComando(18))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196300
     
                If Abs(Arredonda_Moeda(dValorComiAgencia)) > DELTA_VALORMONETARIO Then
                    
                    'Obter NumIntDoc
                    lErro = CF("Config_ObterNumInt", "TRVConfig", "NUM_INT_PROX_TRVVOURCHERINFO", lNumIntDoc)
                    If lErro <> SUCESSO Then gError 196297
                    
                    lContador = lContador + 1
                    
                    'Insere o estorno da CMCC (Comissão da agência em venda sobre cartão de crédito)
                    lErro = Comando_Executar(alComando(9), "INSERT INTO TRVVoucherInfo (NumIntDoc,TipoDoc,TipVou,Serie,NumVou,Seq,Data,TipoDocDestino,NumIntDocDestino,NumIntDocComiInt,Valor,Historico,NumTitulo, Status,TipoLiberacao,NumIntDocLiberacao) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) ", _
                                 lNumIntDoc, TRV_TIPODOC_CMCC_TEXTO, objVou.sTipVou, objVou.sSerie, objVou.lNumVou, lContador, dtData, 0, 0, 0, -Arredonda_Moeda(dValorComiAgencia), sHistorico, 0, iStatus, iTipoLiberacao, lNumIntDocLiberacao)
                    If lErro <> AD_SQL_SUCESSO Then gError 196298
                            
                    'Obter NumIntDoc
                    lErro = CF("Config_ObterNumInt", "TRVConfig", "NUM_INT_PROX_TRVVOURCHERINFO", lNumIntDoc)
                    If lErro <> SUCESSO Then gError 196297
                    
                    lContador = lContador + 1
                    
                    'Insere a CMA (Comissão da agência) (Comissão da agência em venda sobre cartão de crédito)
                    lErro = Comando_Executar(alComando(9), "INSERT INTO TRVVoucherInfo (NumIntDoc,TipoDoc,TipVou,Serie,NumVou,Seq,Data,TipoDocDestino,NumIntDocDestino,NumIntDocComiInt,Valor,Historico,NumTitulo, Status,TipoLiberacao,NumIntDocLiberacao) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) ", _
                                 lNumIntDoc, TRV_TIPODOC_CMA_TEXTO, objVou.sTipVou, objVou.sSerie, objVou.lNumVou, lContador, dtData, 0, 0, 0, Arredonda_Moeda(dValorComiAgencia), sHistorico, 0, iStatus, iTipoLiberacao, lNumIntDocLiberacao)
                    If lErro <> AD_SQL_SUCESSO Then gError 196298
                
                End If
            
            'Passou a ser venda de cartão
            Else
            
                dValorComiAgencia = 0
            
                'Le a TRVVOucherInfo para pegar o saldo do que foi lançado de comissão da agência para estornar e lançar como cartão
                lErro = Comando_Executar(alComando(19), "SELECT SUM(Valor) FROM TRVVoucherInfo WHERE TipoDoc = ? AND TipVou = ? AND Serie = ? AND NumVou = ? ", _
                             dValorComiAgencia, TRV_TIPODOC_CMA_TEXTO, objVou.sTipVou, objVou.sSerie, objVou.lNumVou)
                If lErro <> AD_SQL_SUCESSO Then gError 196299
            
                'Busca Primeiro
                lErro = Comando_BuscarPrimeiro(alComando(19))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196300
            
                If Abs(Arredonda_Moeda(dValorComiAgencia)) > DELTA_VALORMONETARIO Then
                    
                    'Obter NumIntDoc
                    lErro = CF("Config_ObterNumInt", "TRVConfig", "NUM_INT_PROX_TRVVOURCHERINFO", lNumIntDoc)
                    If lErro <> SUCESSO Then gError 196297
                    
                    lContador = lContador + 1
                    
                    'Insere o estorno da CMA (Comissão da agência)
                    lErro = Comando_Executar(alComando(9), "INSERT INTO TRVVoucherInfo (NumIntDoc,TipoDoc,TipVou,Serie,NumVou,Seq,Data,TipoDocDestino,NumIntDocDestino,NumIntDocComiInt,Valor,Historico,NumTitulo, Status,TipoLiberacao,NumIntDocLiberacao) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) ", _
                                 lNumIntDoc, TRV_TIPODOC_CMA_TEXTO, objVou.sTipVou, objVou.sSerie, objVou.lNumVou, lContador, dtData, 0, 0, 0, -Arredonda_Moeda(dValorComiAgencia), sHistorico, 0, iStatus, iTipoLiberacao, lNumIntDocLiberacao)
                    If lErro <> AD_SQL_SUCESSO Then gError 196298
                            
                    'Obter NumIntDoc
                    lErro = CF("Config_ObterNumInt", "TRVConfig", "NUM_INT_PROX_TRVVOURCHERINFO", lNumIntDoc)
                    If lErro <> SUCESSO Then gError 196297
                    
                    lContador = lContador + 1
                    
                    'Insere a CMCC (Comissão da agência em venda sobre cartão de crédito)
                    lErro = Comando_Executar(alComando(9), "INSERT INTO TRVVoucherInfo (NumIntDoc,TipoDoc,TipVou,Serie,NumVou,Seq,Data,TipoDocDestino,NumIntDocDestino,NumIntDocComiInt,Valor,Historico,NumTitulo, Status,TipoLiberacao,NumIntDocLiberacao) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) ", _
                                 lNumIntDoc, TRV_TIPODOC_CMCC_TEXTO, objVou.sTipVou, objVou.sSerie, objVou.lNumVou, lContador, dtData, 0, 0, 0, Arredonda_Moeda(dValorComiAgencia), sHistorico, 0, iStatus, iTipoLiberacao, lNumIntDocLiberacao)
                    If lErro <> AD_SQL_SUCESSO Then gError 196298
                
                End If
                
            End If
        
        'Se não foi só uma alteração no campo Cartão
        Else
           
            'CMR
            'Se tem representante
            If lCliCMR <> 0 Then
            
                dValorCMR = dValorBrutoBase * dPercCMR
                
                If bInsereCMR And dValorCMR <> 0 Then
                
                    'Obter NumIntDoc
                    lErro = CF("Config_ObterNumInt", "TRVConfig", "NUM_INT_PROX_TRVVOURCHERINFO", lNumIntDoc)
                    If lErro <> SUCESSO Then gError 196303
                    
                    lContador = lContador + 1
                    
                    'Insere o CMR
                    lErro = Comando_Executar(alComando(7), "INSERT INTO TRVVoucherINFO (NumIntDoc,TipoDoc,TipVou,Serie,NumVou,Seq,Data,TipoDocDestino,NumIntDocDestino,NumIntDocComiInt,Valor,Historico,NumTitulo,Status,TipoLiberacao,NumIntDocLiberacao) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) ", _
                                 lNumIntDoc, TRV_TIPODOC_CMR_TEXTO, objVou.sTipVou, objVou.sSerie, objVou.lNumVou, lContador, dtData, 0, 0, 0, Arredonda_Moeda(dValorCMR), sHistorico, 0, iStatus, iTipoLiberacao, lNumIntDocLiberacao)
                    If lErro <> AD_SQL_SUCESSO Then gError 196304
                    
                End If
                
            End If
            
            'CMC
            'Se tem representante
            If lCliCMC <> 0 Then
            
                dValorCMC = dValorBrutoBase * dPercCMC
                
                If bInsereCMC And dValorCMC <> 0 Then
                
                    'Obter NumIntDoc
                    lErro = CF("Config_ObterNumInt", "TRVConfig", "NUM_INT_PROX_TRVVOURCHERINFO", lNumIntDoc)
                    If lErro <> SUCESSO Then gError 196305
                    
                    lContador = lContador + 1
                    
                    'Insere o CMC
                    lErro = Comando_Executar(alComando(7), "INSERT INTO TRVVoucherINFO (NumIntDoc,TipoDoc,TipVou,Serie,NumVou,Seq,Data,TipoDocDestino,NumIntDocDestino,NumIntDocComiInt,Valor,Historico,NumTitulo,Status,TipoLiberacao,NumIntDocLiberacao) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) ", _
                                 lNumIntDoc, TRV_TIPODOC_CMC_TEXTO, objVou.sTipVou, objVou.sSerie, objVou.lNumVou, lContador, dtData, 0, 0, 0, Arredonda_Moeda(dValorCMC), sHistorico, 0, iStatus, iTipoLiberacao, lNumIntDocLiberacao)
                    If lErro <> AD_SQL_SUCESSO Then gError 196306
                    
                End If
                
            End If
            
            'OVER (Não tem over quando for dias antecipados)
            If lFornOVER <> 0 And idiasantc = DESMARCADO Then
            
                dValorOVER = dValorBrutoBase * dPercOVER
                
                If bInsereOVER And dValorOVER <> 0 Then
                
                    'Obter NumIntDoc
                    lErro = CF("Config_ObterNumInt", "TRVConfig", "NUM_INT_PROX_TRVVOURCHERINFO", lNumIntDoc)
                    If lErro <> SUCESSO Then gError 196307
                    
                    lContador = lContador + 1
                    
                    iStatusOver = iStatus
                    iTipoLiberacaoOver = iTipoLiberacao
                    lNumIntDocLiberacaoOver = lNumIntDocLiberacao
    
                    'Se é para gravar bloqueado e já foi faturado, verifica se a liberação não é na emissão
                    If iStatus <> TRV_VOU_INFO_STATUS_ANTIGA And objVou.lNumIntDocDestino <> 0 And objVou.iTipoDocDestino = TRV_TIPO_DOC_DESTINO_TITREC Then
        
                        sItem = String(255, 0)
        
                        'Verifica o tipo de tratamento (Se o pagamento é na baixa ou na emissão)
                        lErro = Comando_Executar(alComando(17), "SELECT Item FROM FilialFornecedorCategorias WHERE Fornecedor = ? AND Filial = ? AND Categoria = ? ", sItem, lFornOVER, FILIAL_MATRIZ, "OVER")
                        If lErro <> AD_SQL_SUCESSO Then gError 182532
                        
                        'Busca a primeira
                        lErro = Comando_BuscarProximo(alComando(17))
                        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182533
                        
                        'Se não encontrou o tipo de tratamento então é porque é na baixa
                        If lErro <> AD_SQL_SUCESSO Then
                            sItem = "Baixa"
                        End If
        
                        'Se não for na baixa libera o over para pagamento
                        If InStr(1, UCase(sItem), "BAIXA") <> 0 Then
                            iStatusOver = TRV_VOU_INFO_STATUS_LIBERADO
                            iTipoLiberacaoOver = TRV_TIPO_LIBERACAO_COMISSAO_FAT
                            lNumIntDocLiberacaoOver = objVou.lNumIntDocDestino
                        End If
        
                    End If
                    
                    'Insere o OVER
                    lErro = Comando_Executar(alComando(10), "INSERT INTO TRVVoucherINFO (NumIntDoc,TipoDoc,TipVou,Serie,NumVou,Seq,Data,TipoDocDestino,NumIntDocDestino,NumIntDocComiInt,Valor,Historico,NumTitulo,Status,TipoLiberacao,NumIntDocLiberacao) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) ", _
                                 lNumIntDoc, TRV_TIPODOC_OVER_TEXTO, objVou.sTipVou, objVou.sSerie, objVou.lNumVou, lContador, dtData, 0, 0, 0, Arredonda_Moeda(dValorOVER), sHistorico, 0, iStatusOver, iTipoLiberacaoOver, lNumIntDocLiberacaoOver)
                    If lErro <> AD_SQL_SUCESSO Then gError 196308
                    
                End If
            
            End If
            
        End If
        
        sTipoDocInfoAux = String(STRING_TRV_VOU_INFO_TIPODOC, 0)
        
        'Le a tabelaTRVVoucherInfo para pegar todos registros bloqueados que tenha algo liberado de mesmo tipo
        '(Para evitar que algo que tenha sido liberado e faturado e por isso
        'não conseguiu ser bloqueado e por isso continue liberado e seu estorno bloqueado)
        lErro = Comando_ExecutarPos(alComando(20), "SELECT NumIntDoc, TipoDoc FROM TRVVoucherInfo WHERE TipVou = ? AND Serie = ? AND NumVou = ? AND Status = ? ", 0, _
                    lNumIntDocInfoAux, sTipoDocInfoAux, objVou.sTipVou, objVou.sSerie, objVou.lNumVou, TRV_VOU_INFO_STATUS_BLOQUEADO)
        If lErro <> AD_SQL_SUCESSO Then gError 196395
        
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(20))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196396
        
        'Enquanto existir
        Do While lErro <> AD_SQL_SEM_DADOS
    
            'Busca para ver se já tem algo liberado para o documento
            lErro = Comando_Executar(alComando(22), "SELECT TipoLiberacao, NumIntDocLiberacao FROM TRVVoucherInfo WHERE TipoDoc = ? AND TipVou = ? AND Serie = ? AND NumVou = ? AND Status = ?  ", _
                        iTipoLiberacao, lNumIntDocLiberacao, sTipoDocInfoAux, objVou.sTipVou, objVou.sSerie, objVou.lNumVou, TRV_VOU_INFO_STATUS_LIBERADO)
            If lErro <> AD_SQL_SUCESSO Then gError 196400
            
            'Busca Primeiro
            lErro = Comando_BuscarPrimeiro(alComando(22))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196401
            
            If lErro = AD_SQL_SUCESSO Then
                
                'Trava registro contra alterações/Leituras
                lErro = Comando_LockExclusive(alComando(20))
                If lErro <> SUCESSO Then gError 196397
        
                lErro = Comando_ExecutarPos(alComando(21), "UPDATE TRVVoucherInfo SET Status = ?, Liberacao = ?, NumIntDocLiberacao = ?", alComando(20), TRV_VOU_INFO_STATUS_LIBERADO, iTipoLiberacao, lNumIntDocLiberacao)
                If lErro <> AD_SQL_SUCESSO Then gError 196398
                
            End If
            
            'Busca Próximo
            lErro = Comando_BuscarProximo(alComando(20))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196399
            
        Loop

        'Fecha Comando
        For iIndice = LBound(alComando) To UBound(alComando)
            Call Comando_Fechar(alComando(iIndice))
        Next
    
    End If

    TRVVoucher_Gera_Comissao = SUCESSO

    Exit Function

Erro_TRVVoucher_Gera_Comissao:

    TRVVoucher_Gera_Comissao = gErr

    Select Case gErr

        Case 196280
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 196281, 196282
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISCLIENTESTRV", gErr)

        Case 196283, 196284, 196299, 196300
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVVOUCHERS", gErr)
            
        Case 196285
            Call Rotina_Erro(vbOKOnly, "ERRO_VOUCHER_NAO_CADASTRADO", gErr, objVou.lNumVou, objVou.sSerie, objVou.sTipVou)

        Case 196286, 196287
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVVOUCHERINFOSIGAV", gErr)
            
        Case 196288
            Call Rotina_Erro(vbOKOnly, "ERRO_VOUCHER_SEM_DADOS_SIGAV", gErr)
            
        Case 196289, 196290
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVCLIEMISSORES", gErr)
            
        Case 196291
            Call Rotina_Erro(vbOKOnly, "ERRO_TRVCLIEMISSORES_NAO_CADASTRADO", gErr, lFornOVER, objVou.lCliente)

        Case 196292, 196293
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVCLIEMISSORESEXC", gErr)
            
        Case 196294
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRVVOUCHERS", gErr)
        
        Case 196295, 196297, 196303, 196305, 196307, 196309, 196355, 192367
        
        Case 196296, 196298, 196304, 196306, 196308, 196310
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRVVOUCHERINFO", gErr)
        
        Case 196301, 196302
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVVOUCHERINFO", gErr)

        Case 196311
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRVVOUCHERS", gErr)

        Case 196350, 196351
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CAMPOSGENERICOSVALORES", gErr)
            
        Case 196352
            Call Rotina_Erro(vbOKOnly, "ERRO_TRV_DESTINO_NAO_CADASTRADO", gErr, Trim(Left(sDestino, 9)))

        Case 196353, 196354
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVACORDOS", gErr)

        Case 196356, 196357
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVACORDOCOMISSAO", gErr)

        Case 196358, 196359
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVACORDOTARIFADIF", gErr)
            
        Case 192368, 192369
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BAIXASPARCREC", gErr)

        Case 196395, 196396, 196399, 196400, 196401
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVVOUCHERINFO", gErr)

        Case 196397
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRVVOUCHERINFO", gErr)

        Case 196398
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRVVOUCHERINFO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196311)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function TRVCliente_Grava(ByVal objCliente As ClassCliente) As Long

Dim lErro As Long
Dim alComando(0 To 7) As Long
Dim iIndice As Integer
Dim objClienteTRV As New ClassClienteTRV
Dim objCliEmiTRV As ClassCliEmissoresTRV
Dim objCliEmiExcTRV As ClassCliEmissoresExcTRV
Dim lEmpPai As Long
Dim lCorrent As Long
Dim dComiss As Double
Dim lCodRep As Long
Dim iAporte As Integer
Dim dPercComiCorr As Double
Dim dPercComiRep As Double
Dim lEmissor As Long
Dim lEmissorAnt As Long
Dim dComissEmissor As Double
Dim iSeq As Integer
Dim sProduto As String
Dim dComissEmissorExc As Double
Dim scodgru As String
Dim iFilialEmp As Integer
Dim iFilialCoinfo As Integer
Dim iFilialNF As Integer
Dim iFilialFat As Integer
Dim iCondPagtoCC As Integer

On Error GoTo Erro_TRVCliente_Grava

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 196320
    Next
    
    scodgru = String(255, 0)
    
    'Le a empresa do Sigav
    lErro = Comando_Executar(alComando(0), "SELECT emppai, corrent,comis, codRep, aporte, codgru FROM CoinfoEmpresaAtual WHERE CodEmp = ? ", _
                lEmpPai, lCorrent, dComiss, lCodRep, iAporte, scodgru, objCliente.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 196321

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196322
    
    If lErro = SUCESSO Then 'gError 196323
    
        'Le a empresa do Sigav
        lErro = Comando_Executar(alComando(6), "SELECT Filial FROM CoinfoGrupoAtual WHERE Codigo = ? ", _
                    iFilialCoinfo, scodgru)
        If lErro <> AD_SQL_SUCESSO Then gError 196321
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(6))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196322
        
        If lErro <> AD_SQL_SUCESSO Then
            iFilialCoinfo = 1
            iFilialEmp = 1
        Else
            iFilialEmp = Filial_Coinfo_Retorna_Corporator(iFilialCoinfo)
        End If
    
        'Se tem correntista
        If lCorrent <> 0 Then
        
            'Le a empresa do Sigav (Correntista)
            lErro = Comando_Executar(alComando(1), "SELECT comis FROM CoinfoEmpresa WHERE CodEmp = ? ORDER BY NumIntArq DESC, NumRegArqTexto DESC", _
                        dPercComiCorr, lCorrent)
            If lErro <> AD_SQL_SUCESSO Then gError 196324
        
            'Busca Primeiro
            lErro = Comando_BuscarPrimeiro(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196325
            
            If lErro <> SUCESSO Then gError 196326
        
        End If
        
        'Se tem representante
        If lCodRep <> 0 Then
        
            'Le a empresa do Sigav (Representante)
            lErro = Comando_Executar(alComando(0), "SELECT comis FROM CoinfoEmpresa WHERE CodEmp = ? ORDER BY NumIntArq DESC, NumRegArqTexto DESC", _
                        dPercComiRep, lCodRep)
            If lErro <> AD_SQL_SUCESSO Then gError 196327
        
            'Busca Primeiro
            lErro = Comando_BuscarPrimeiro(alComando(0))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196328
            
            If lErro <> SUCESSO Then gError 196329
        
        End If
        
        objClienteTRV.dPercComiAg = (dComiss) / 100
        If lCorrent <> 0 Then objClienteTRV.dPercComiCorr = (dPercComiCorr - dComiss) / 100
        If lCodRep <> 0 Then objClienteTRV.dPercComiRep = (dPercComiRep - dComiss) / 100
        objClienteTRV.iConsiderarAporte = iAporte
        objClienteTRV.lCodCorrentista = lCorrent
        objClienteTRV.lCodRepresentante = lCodRep
        objClienteTRV.lEmpresaPai = lEmpPai
        objClienteTRV.sGrupo = scodgru
        objClienteTRV.iFilialCoinfo = iFilialCoinfo
        objClienteTRV.iFilialEmpresa = iFilialEmp
        
        If objClienteTRV.dPercComiCorr < 0 Then objClienteTRV.dPercComiCorr = 0
        If objClienteTRV.dPercComiRep < 0 Then objClienteTRV.dPercComiRep = 0
        
        Set objCliente.objInfoUsu = objClienteTRV
        
        'Le os emissores
        lErro = Comando_Executar(alComando(3), "SELECT CodEmi, percover FROM CoinfoEmissor WHERE CodEmp = ? ORDER BY NumIntArq DESC, NumRegArqTexto DESC", _
                    lEmissor, dComissEmissor, objCliente.lCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 196330
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196331
        
        lEmissorAnt = 0
        Do While lErro <> AD_SQL_SEM_DADOS
        
            If lEmissor <> lEmissorAnt Then
                
                lEmissorAnt = lEmissor
                
                Set objCliEmiTRV = New ClassCliEmissoresTRV
                
                objCliEmiTRV.dPercComissao = dComissEmissor / 100
                objCliEmiTRV.lCliente = objCliente.lCodigo
                objCliEmiTRV.lFornEmissor = lEmissor + FATOR_SOMA_COD_EMISSOR
                objCliEmiTRV.iSeq = objClienteTRV.colCliEmissoresTRV.Count + 1
                
                objClienteTRV.colCliEmissoresTRV.Add objCliEmiTRV
                
            End If
        
             'Busca Próximo
            lErro = Comando_BuscarProximo(alComando(3))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196332
               
        Loop
        
        For Each objCliEmiTRV In objClienteTRV.colCliEmissoresTRV
        
            sProduto = String(STRING_PRODUTO, 0)
        
            'Le as excessões já cadastradas
            lErro = Comando_Executar(alComando(4), "SELECT EE.Seq, EE.Produto, EE.PercComissao FROM TRVCliEmissoresExc AS EE, TRVCliEmissores AS E WHERE EE.NumIntDocCliEmi = E.NumIntDoc AND E.Cliente = ? AND E.FornEmissor = ?", _
                        iSeq, sProduto, dComissEmissorExc, objCliente.lCodigo, objCliEmiTRV.lFornEmissor)
            If lErro <> AD_SQL_SUCESSO Then gError 196333
        
            'Busca Primeiro
            lErro = Comando_BuscarPrimeiro(alComando(4))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196334
            
            Do While lErro <> AD_SQL_SEM_DADOS
            
                Set objCliEmiExcTRV = New ClassCliEmissoresExcTRV
                
                objCliEmiExcTRV.dPercComissao = dComissEmissorExc
                objCliEmiExcTRV.iSeq = iSeq
                objCliEmiExcTRV.sProduto = sProduto
                
                objCliEmiTRV.colCliEmissoresExcTRV.Add objCliEmiExcTRV
            
                'Busca Próximo
                lErro = Comando_BuscarProximo(alComando(4))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196335
            
            Loop
        
        Next
        
        lErro = Comando_Executar(alComando(7), "SELECT FilialNF,FilialFat,CondPagtoCC FROM FiliaisClientesTRV WHERE CodCliente = ?  AND CodFilial = ? ", _
                    iFilialNF, iFilialFat, iCondPagtoCC, objCliente.lCodigo, FILIAL_MATRIZ)
        If lErro <> AD_SQL_SUCESSO Then gError 196333
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(7))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196334
        
        If lErro <> AD_SQL_SEM_DADOS Then
            objClienteTRV.iFilialNF = iFilialNF
            objClienteTRV.iFilialFat = iFilialFat
            objClienteTRV.iCondPagtoCC = iCondPagtoCC
        End If
        
        lErro = CF("TRVClienteCorProd_Le", objCliente)
        If lErro <> SUCESSO Then gError 196951
        
        lErro = CF("TRVClienteRepProd_Le", objCliente)
        If lErro <> SUCESSO Then gError 196952
        
        lErro = Cliente_Grava_Customizado(objCliente)
        If lErro <> SUCESSO Then gError 196336
        
    End If
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TRVCliente_Grava = SUCESSO

    Exit Function

Erro_TRVCliente_Grava:

    TRVCliente_Grava = gErr

    Select Case gErr

        Case 196320
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 196321, 196322, 196324, 196325, 196327, 196328
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COINFOEMPRESA", gErr)
            
        Case 196323
            Call Rotina_Erro(vbOKOnly, "ERRO_COINFOEMPRESA_NAO_CADASTRADO", gErr, objCliente.lCodigo)

        Case 196326
            Call Rotina_Erro(vbOKOnly, "ERRO_COINFOEMPRESA_NAO_CADASTRADO", gErr, lCorrent)

        Case 196329
            Call Rotina_Erro(vbOKOnly, "ERRO_COINFOEMPRESA_NAO_CADASTRADO", gErr, lCodRep)

        Case 196330, 196331, 196332
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COINFOEMISSOR", gErr)

        Case 196333, 196334, 196335
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVCLIEMISSOR", gErr)
            
        Case 196336, 196951, 196952

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196337)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function RPS_Obtem_Discriminacao_Cust(ByVal objNF As ClassNFiscal, sDiscriminacao As String)

Dim lErro As Long
Dim objItemNF As ClassItemNF
Dim iPOS As Integer
Dim sParte As String

On Error GoTo Erro_RPS_Obtem_Discriminacao_Cust

    sDiscriminacao = ""
    
    For Each objItemNF In objNF.ColItensNF
    
'        sParte = objItemNF.sDescricaoItem
'
'        iPOS = InStr(1, sParte, " ")
'
'        Do While iPOS <> 0
'
'            sParte = Mid(sParte, iPOS + 1)
'
'            iPOS = InStr(1, sParte, " ")
'
'        Loop
'
'        sDiscriminacao = sDiscriminacao & IIf(Len(Trim(sDiscriminacao)) > 0, "||", "") & "Prestação de serviço de assistência ao viajante, referente a fatura Nro. " & sParte
    
        sDiscriminacao = sDiscriminacao & IIf(Len(Trim(sDiscriminacao)) > 0, "||", "") & objItemNF.sDescricaoItem
    
    Next
    
    RPS_Obtem_Discriminacao_Cust = SUCESSO

    Exit Function

Erro_RPS_Obtem_Discriminacao_Cust:

    RPS_Obtem_Discriminacao_Cust = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196337)

    End Select

    Exit Function
    
End Function

Public Function TRVAportePagtoFatCond_Exclui(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal objTRVAportePagtoFatCond As ClassTRVAportePagtoFatC) As Long

Dim lErro As Long
Dim lNumIntDoc As Long

On Error GoTo Erro_TRVAportePagtoFatCond_Exclui

    'Le a tabelaTRVAportePagtoFatCond
    lErro = Comando_ExecutarPos(lComando1, "SELECT NumIntDoc FROM TRVAportePagtoFatCond WHERE NumIntDoc= ? ", _
                0, lNumIntDoc, objTRVAportePagtoFatCond.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 190416

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190417

    'Se não existir => ERRO
    If lErro <> AD_SQL_SUCESSO Then gError 190418

    lErro = Comando_ExecutarPos(lComando2, "DELETE FROM TRVAportePagtoFatCond ", lComando1)
    If lErro <> AD_SQL_SUCESSO Then gError 190419
    
    TRVAportePagtoFatCond_Exclui = SUCESSO

    Exit Function

Erro_TRVAportePagtoFatCond_Exclui:

    TRVAportePagtoFatCond_Exclui = gErr

    Select Case gErr

        Case 190416, 190417
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVAPORTEPAGTOFATCOND", gErr)

        Case 190418
            Call Rotina_Erro(vbOKOnly, "ERRO_TRVAPORTEPAGTOFATCOND_NAO_CADASTRADO", gErr, objTRVAportePagtoFatCond.lNumIntDoc)
        
        Case 190419
            Call Rotina_Erro(vbOKOnly, "ERRO_DELETE_TRVAPORTEPAGTOFATCOND", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190420)

    End Select

    Exit Function

End Function

Public Function TRVAportePagtoFatCond_Atualiza(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal objTRVAportePagtoFatCond As ClassTRVAportePagtoFatC) As Long

Dim lErro As Long
Dim lNumIntDoc As Long

On Error GoTo Erro_TRVAportePagtoFatCond_Atualiza

    'Le a tabelaTRVAportePagtoFatCond
    lErro = Comando_ExecutarPos(lComando1, "SELECT NumIntDoc FROM TRVAportePagtoFatCond WHERE NumIntDoc= ? ", _
                0, lNumIntDoc, objTRVAportePagtoFatCond.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 190411

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190412

    'Se não existir => ERRO
    If lErro <> AD_SQL_SUCESSO Then gError 190413

    lErro = Comando_ExecutarPos(lComando2, "UPDATE TRVAportePagtoFatCond SET Valor= ? , ValidadeDe= ? , ValidadeAte= ? , " & _
            "Percentual= ?, Seq = ? ", lComando1, _
            objTRVAportePagtoFatCond.dValor, objTRVAportePagtoFatCond.dtValidadeDe, objTRVAportePagtoFatCond.dtValidadeAte, _
            objTRVAportePagtoFatCond.dPercentual, objTRVAportePagtoFatCond.iSeq)
    If lErro <> AD_SQL_SUCESSO Then gError 190414
    
    TRVAportePagtoFatCond_Atualiza = SUCESSO

    Exit Function

Erro_TRVAportePagtoFatCond_Atualiza:

    TRVAportePagtoFatCond_Atualiza = gErr

    Select Case gErr

        Case 190411, 190412
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVAPORTEPAGTOFATCOND", gErr)

        Case 190413
            Call Rotina_Erro(vbOKOnly, "ERRO_TRVAPORTEPAGTOFATCOND_NAO_CADASTRADO", gErr, objTRVAportePagtoFatCond.lNumIntDoc)
        
        Case 190414
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRVAPORTEPAGTOFATCOND", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190415)

    End Select

    Exit Function

End Function

Public Function TRVAportePagtoFatCond_Insere(ByVal lComando As Long, ByVal objTRVAportePagtoFatCond As ClassTRVAportePagtoFatC, ByVal objAporte As ClassTRVAportes) As Long

Dim lErro As Long
Dim lNumIntDoc As Long

On Error GoTo Erro_TRVAportePagtoFatCond_Insere

    'Obter NumIntDoc
    lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_INT_PROX_TRVAPORTEPAGTOFATCOND", lNumIntDoc)
    If lErro <> SUCESSO Then gError 190408

    objTRVAportePagtoFatCond.lNumIntDoc = lNumIntDoc

    lErro = Comando_Executar(lComando, "INSERT INTO TRVAportePagtoFatCond( NumIntDoc, NumIntDocAporte, Valor, ValidadeDe, " & _
            "ValidadeAte, Percentual, Seq) VALUES (?,?,?,?,?,?,?)", _
            objTRVAportePagtoFatCond.lNumIntDoc, objTRVAportePagtoFatCond.lNumIntDocAporte, objTRVAportePagtoFatCond.dValor, objTRVAportePagtoFatCond.dtValidadeDe, _
            objTRVAportePagtoFatCond.dtValidadeAte, objTRVAportePagtoFatCond.dPercentual, objTRVAportePagtoFatCond.iSeq)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190409

    TRVAportePagtoFatCond_Insere = SUCESSO

    Exit Function

Erro_TRVAportePagtoFatCond_Insere:

    TRVAportePagtoFatCond_Insere = gErr

    Select Case gErr

        Case 190408

        Case 190409
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRVAPORTEPagtoFatCond", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190410)

    End Select

    Exit Function

End Function

Public Function TRVAportes_Valida(ByVal objTRVAportes As ClassTRVAportes) As Long

Dim lErro As Long
Dim alComando(0 To 0) As Long
Dim iIndice As Integer
Dim iIndice2 As Integer
Dim objPagtoFat As ClassTRVAportePagtoFat
Dim objPagtoFatAux As ClassTRVAportePagtoFat
Dim sTexto As String
Dim sTextoFinal As String
Dim lCodigo As Long
Dim dValor As Double
Dim dSaldo As Double
Dim dtValidadeDe As Date
Dim dtValidadeAte As Date
Dim iSeq As Integer
Dim vbResult As VbMsgBoxResult
Dim bValidadeRepetida

On Error GoTo Erro_TRVAportes_Valida

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 196480
    Next
    
    iIndice = 0
    For Each objPagtoFat In objTRVAportes.colPagtoSobreFatura
        bValidadeRepetida = False
        iIndice = iIndice + 1
        For iIndice2 = iIndice + 1 To objTRVAportes.colPagtoSobreFatura.Count
            Set objPagtoFatAux = objTRVAportes.colPagtoSobreFatura.Item(iIndice2)
            If Not ((objPagtoFatAux.dtValidadeDe < objPagtoFat.dtValidadeDe And objPagtoFatAux.dtValidadeAte < objPagtoFat.dtValidadeDe) Or (objPagtoFatAux.dtValidadeDe > objPagtoFat.dtValidadeAte And objPagtoFatAux.dtValidadeAte > objPagtoFat.dtValidadeAte)) Then
                bValidadeRepetida = True
                sTexto = sTexto & " e Aporte: " & objTRVAportes.lCodigo & " Valor: " & Format(objPagtoFatAux.dValor, "STANDARD") & " Saldo: " & Format(objPagtoFatAux.dSaldo, "STANDARD") & " Validade De: " & Format(objPagtoFatAux.dtValidadeDe, "dd/mm/yyyy") & " Validade De: " & Format(objPagtoFatAux.dtValidadeAte, "dd/mm/yyyy") & " Linha : " & CStr(objPagtoFatAux.iSeq)
            End If
        Next
    
        'Le a tabelaTRVAportes
        lErro = Comando_Executar(alComando(0), "SELECT A.Codigo, P.Valor, P.Saldo, P.Validadede, P.ValidadeAte, P.Seq FROM TRVAportes AS A, TRVAportePagtoFat AS P WHERE A.NumintDoc = P.NumIntDocAporte AND NOT ((ValidadeDe < ? AND ValidadeAte < ?) OR (ValidadeDe > ? AND ValidadeAte > ?)) AND Cliente = ? AND Codigo <> ? ", lCodigo, dValor, dSaldo, dtValidadeDe, dtValidadeAte, iSeq, objPagtoFat.dtValidadeDe, objPagtoFat.dtValidadeDe, objPagtoFat.dtValidadeAte, objPagtoFat.dtValidadeAte, objTRVAportes.lCliente, objTRVAportes.lCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 196481
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196482
        
        Do While lErro <> AD_SQL_SEM_DADOS
            
            bValidadeRepetida = True
            sTexto = sTexto & " Aporte: " & lCodigo & " Valor: " & Format(dValor, "STANDARD") & " Saldo: " & Format(dSaldo, "STANDARD") & " Validade De: " & Format(dtValidadeDe, "dd/mm/yyyy") & " Validade De: " & Format(dtValidadeAte, "dd/mm/yyyy") & " Linha : " & CStr(iSeq)
        
            'Busca Próximo
            lErro = Comando_BuscarProximo(alComando(0))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196483

        Loop
        
        If bValidadeRepetida Then
            sTextoFinal = sTextoFinal & " Aporte: " & objTRVAportes.lCodigo & " Valor: " & Format(objPagtoFat.dValor, "STANDARD") & " Saldo: " & Format(objPagtoFat.dSaldo, "STANDARD") & " Validade De: " & Format(objPagtoFat.dtValidadeDe, "dd/mm/yyyy") & " Validade De: " & Format(objPagtoFat.dtValidadeAte, "dd/mm/yyyy") & " Linha : " & CStr(objPagtoFat.iSeq) & sTexto
        End If
    
    Next
    
    If Len(Trim(sTextoFinal)) > 0 Then
    
        vbResult = Rotina_Aviso(vbYesNo, "AVISO_PAGTOAPORTE_VALIDADEREPETIDA", objTRVAportes.lCliente, sTextoFinal)
        If vbResult = vbNo Then gError 196484
    
    End If
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TRVAportes_Valida = SUCESSO

    Exit Function

Erro_TRVAportes_Valida:

    TRVAportes_Valida = gErr

    Select Case gErr

        Case 196480
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 196481 To 196483
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVAPORTES", gErr)
        
        Case 196484

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196485)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Private Function NF_NFiscal_Grava_TRV(alComando() As Long, ByVal objNFiscal As ClassNFiscal, iProxComando As Integer) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim objItemNF As ClassItemNF
Dim tSerie As typeSerie
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim objFilCliFilEmp As New ClassFilCliFilEmp
Dim objRPS As New ClassRPS
Dim objNFTRV As ClassNFTRV
Dim tFilCliFilEmp As typeFilCliFilEmp

On Error GoTo Erro_NF_NFiscal_Grava_TRV
    
    Set objNFTRV = objNFiscal.objInfoUsu

    lErro = Comando_ExecutarPos(alComando(iProxComando + 12), "SELECT ProxNumNFiscal FROM Serie WHERE Serie = ? AND FilialEmpresa = ?", 0, tSerie.lProxNumNFiscal, objNFiscal.sSerie, objNFiscal.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 196795

    lErro = Comando_BuscarPrimeiro(alComando(iProxComando + 12))
    If lErro <> AD_SQL_SUCESSO Then gError 196796
    
    lErro = Comando_LockExclusive(alComando(iProxComando + 12))
    If lErro <> AD_SQL_SUCESSO Then gError 196797
    
    objNFiscal.lNumNotaFiscal = tSerie.lProxNumNFiscal
    
    tSerie.lProxNumNFiscal = tSerie.lProxNumNFiscal + 1
    
    lErro = Comando_ExecutarPos(alComando(iProxComando + 13), "UPDATE Serie SET ProxNumNFiscal = ?", alComando(iProxComando + 12), tSerie.lProxNumNFiscal)
    If lErro <> AD_SQL_SUCESSO Then gError 196798
    
    'Obtem o Número Interno da nova Nota Fiscal
    lErro = CF("NFiscal_Automatico1", alComando(iProxComando + 1), alComando(iProxComando + 2), alComando(iProxComando + 3), lNumIntDoc)
    If lErro <> SUCESSO Then gError 196799

    objNFiscal.lNumIntDoc = lNumIntDoc

    With objNFiscal

        'Insere a nova Nota Fiscal de Saida no BD
        lErro = Comando_Executar(alComando(iProxComando + 4), "INSERT INTO NFiscal (DataAtualizacao, NumIntDoc, Status, FilialEmpresa, Serie, NumNotaFiscal, Cliente, FilialCli, FilialEntrega, Fornecedor, FilialForn, DataEmissao, DataEntrada, DataSaida, DataVencimento, DataReferencia,NumPedidoVenda, NumPedidoTerc, ClasseDocCPR, NumIntDocCPR, ValorTotal, ValorProdutos, ValorFrete, ValorSeguro, ValorOutrasDespesas, ValorDesconto, CodTransportadora, MensagemNota, TabelaPreco, TipoNFiscal, NaturezaOp, PesoLiq, PesoBruto, NumIntTrib, Placa, PlacaUF, VolumeQuant, VolumeEspecie, VolumeMarca, Canal, NumIntNotaOriginal,FilialPedido, VolumeNumero, FreteRespons, HoraSaida, CodTranspRedesp, DetPagFrete, Moeda, TaxaMoeda, VolumeTotal, SemDataSaida, NaoImpCobranca, MensagemCorpoNota,RPS) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
            gdtDataHoje, .lNumIntDoc, .iStatus, .iFilialEmpresa, .sSerie, .lNumNotaFiscal, .lCliente, .iFilialCli, .iFilialEntrega, .lFornecedor, .iFilialForn, .dtDataEmissao, .dtDataEntrada, .dtDataSaida, .dtDataVencimento, .dtDataReferencia, .lNumPedidoVenda, .sNumPedidoTerc, .iClasseDocCPR, .lNumIntDocCPR, .dValorTotal, .dValorProdutos, .dValorFrete, .dValorSeguro, .dValorOutrasDespesas, .dValorDesconto, .iCodTransportadora, .sMensagemNota, .iTabelaPreco, .iTipoNFiscal, .sNaturezaOp, .dPesoLiq, .dPesoBruto, .lNumIntTrib, .sPlaca, .sPlacaUF, .lVolumeQuant, .lVolumeEspecie, .lVolumeMarca, .iCanal, .lNumIntNotaOriginal, .iFilialPedido, .sVolumeNumero, .iFreteRespons, CDbl(.dtHoraSaida), .iCodTranspRedesp, .iDetPagFrete, .iMoeda, .dTaxaMoeda, .dVolumeTotal, .iSemDataSaida, .iNaoImpCobranca, .sMensagemCorpoNota, .iRecibo)
        If lErro <> AD_SQL_SUCESSO Then gError 196800
        
    End With

    'grava a tributacao da nota como um todo
    lErro = CF("NFiscal_GravaTributacao", objNFiscal, alComando(iProxComando + 5), alComando(iProxComando + 6))
    If lErro <> SUCESSO Then gError 196801
   
    'Obtém o Número Interno dos itens da NF
    lErro = CF("CRFATConfig_ObterNumInt", alComando(iProxComando + 7), alComando(iProxComando + 8), "NUM_PROX_ITEM_NOTA_FISCAL", lNumIntDoc, objNFiscal.ColItensNF.Count)
    If lErro <> SUCESSO Then gError 196802
   
    'Para cada item da Nota
    For Each objItemNF In objNFiscal.ColItensNF

        objItemNF.lNumIntNF = objNFiscal.lNumIntDoc
        objItemNF.lNumIntDoc = lNumIntDoc

        With objItemNF
        
            'Insere o item no Banco de dados
            lErro = Comando_Executar(alComando(iProxComando + 9), "INSERT INTO ItensNFiscal (NumIntNF, Item, Status, Produto, UnidadeMed, Quantidade, PrecoUnitario, PercDesc, ValorDesconto, DataEntrega, DescricaoItem, ValorAbatComissao, NumIntPedVenda, NumIntItemPedVenda, NumIntDoc, NumIntTrib, NumIntDocOrig, Almoxarifado) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", .lNumIntNF, .iItem, STATUS_LANCADO, .sProduto, .sUnidadeMed, .dQuantidade, .dPrecoUnitario, .dPercDesc, .dValorDesconto, .dtDataEntrega, .sDescricaoItem, .dValorAbatComissao, .lNumIntPedVenda, .lNumIntItemPedVenda, lNumIntDoc, .lNumIntTrib, 0, .iAlmoxarifado)
            If lErro <> AD_SQL_SUCESSO Then gError 196803
            
        End With
        
        'grava a tributacao relativa a cada item da nota
        lErro = CF("NFiscal_GravaTributacaoItem", objItemNF, alComando(iProxComando + 10), alComando(iProxComando + 11))
        If lErro <> SUCESSO Then gError 196804
    
        lNumIntDoc = lNumIntDoc + 1
    
    Next
    
    Set objTipoDocInfo = objNFTRV.objTipoDocInfo
    
    'Trata-se de uma nota fiscal de venda
    If (objTipoDocInfo.iFaturamento = TIPODOCINFO_FATURAVEL) And (objTipoDocInfo.iComplementar = DOCINFO_NORMAL) Then
        
        objFilCliFilEmp.lCodCliente = objNFiscal.lCliente
        objFilCliFilEmp.iFilialEmpresa = objNFiscal.iFilialEmpresa
        objFilCliFilEmp.iCodFilial = objNFiscal.iFilialCli
        
        'Lê a tabela FilCliFilEmp
        lErro = Comando_ExecutarPos(alComando(iProxComando + 14), "SELECT NumeroCompras, MediaCompra, ValorAcumuladoCompras  FROM FilialClienteFilEmp WHERE FilialEmpresa = ? AND CodCliente =? AND CodFilial = ?", 0, tFilCliFilEmp.lNumeroCompras, tFilCliFilEmp.dMediaCompra, tFilCliFilEmp.dValorAcumuladoCompras, objFilCliFilEmp.iFilialEmpresa, objFilCliFilEmp.lCodCliente, objFilCliFilEmp.iCodFilial)
        If lErro <> AD_SQL_SUCESSO Then gError 196805
    
        lErro = Comando_BuscarPrimeiro(alComando(iProxComando + 14))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196806
    
        'Não Achou a FilCliFilEmp
        If lErro = AD_SQL_SEM_DADOS Then
    
            'Insere na tabela de FilialClienteFilEmp
            lErro = Comando_Executar(alComando(iProxComando + 16), "INSERT INTO FilialClienteFilEmp (FilialEmpresa, CodCliente, CodFilial, NumeroCompras ,MediaCompra, ValorAcumuladoCompras) VALUES (?,?,?,?,?,?)", objNFiscal.iFilialEmpresa, objNFiscal.lCliente, objNFiscal.iFilialCli, 1, objNFiscal.dValorTotal, objNFiscal.dValorTotal)
            If lErro <> AD_SQL_SUCESSO Then gError 196807
        'se encontrou
        Else
            
            'Lock da FilCliFilEmp
            lErro = Comando_LockExclusive(alComando(iProxComando + 14))
            If lErro <> AD_SQL_SUCESSO Then gError 196808
            
            'recalcula os valores
            objFilCliFilEmp.dValorAcumuladoCompras = tFilCliFilEmp.dValorAcumuladoCompras + objNFiscal.dValorTotal
            objFilCliFilEmp.lNumeroCompras = tFilCliFilEmp.lNumeroCompras + 1
            objFilCliFilEmp.dMediaCompra = tFilCliFilEmp.dValorAcumuladoCompras / objFilCliFilEmp.lNumeroCompras
            
            'Atualiza a tabela de FilialClienteFilEmp
            lErro = Comando_ExecutarPos(alComando(iProxComando + 15), "UPDATE FilialClienteFilEmp SET NumeroCompras =?,MediaCompra=?, ValorAcumuladoCompras=?", alComando(iProxComando + 14), objFilCliFilEmp.lNumeroCompras, objFilCliFilEmp.dMediaCompra, objFilCliFilEmp.dValorAcumuladoCompras)
            If lErro <> AD_SQL_SUCESSO Then gError 196809
        End If
        
    
    End If

    'verifica se é Notafiscal de venda
    If (objTipoDocInfo.iFaturamento = TIPODOCINFO_FATURAMENTO_SIM Or objTipoDocInfo.iFaturamento = TIPODOCINFO_FATURAMENTO_DEV) And (objTipoDocInfo.iComplementar = DOCINFO_NORMAL Or (objTipoDocInfo.iCodigo = DOCINFO_NFISCP Or objTipoDocInfo.iCodigo = DOCINFO_NFISFCP)) Then
        
        'Grava a estatistica do Dia do produto
        lErro = NF_SldDiaFat_Grava_Est_TRV(alComando(iProxComando + 17), alComando(iProxComando + 18), alComando(iProxComando + 19), objNFiscal, objTipoDocInfo)
        If lErro <> SUCESSO Then gError 196810
        
        'Grava a estatística do Mês do produto
        lErro = NF_SldMesFat_Grava_Est_TRV(alComando(iProxComando + 20), alComando(iProxComando + 21), alComando(iProxComando + 22), alComando(iProxComando + 23), alComando(iProxComando + 24), alComando(iProxComando + 25), objNFiscal, objTipoDocInfo)
        If lErro <> SUCESSO Then gError 196811
    
    End If

    'Verifica se o modulo de Livros Fiscais está Ativo
    If gcolModulo.Ativo(MODULO_LIVROSFISCAIS) = MODULO_ATIVO Then
                        
        lErro = NF_NotaFiscal_Grava_Fis_TRV(alComando(iProxComando + 26), alComando(iProxComando + 27), alComando(iProxComando + 28), alComando(iProxComando + 29), alComando(iProxComando + 30), alComando(iProxComando + 31), alComando(iProxComando + 32), alComando(iProxComando + 33), alComando(iProxComando + 36), alComando(iProxComando + 37), alComando(iProxComando + 39), alComando(iProxComando + 40), alComando(iProxComando + 41), alComando(iProxComando + 42), objNFiscal)
        If lErro <> SUCESSO Then gError 196812
    
    End If
    
    If objNFiscal.iRecibo = MARCADO Then

        lErro = NF_Converte_RPS_TRV(objNFiscal, objRPS)
        If lErro <> SUCESSO Then gError 196813
        
        If objRPS.dValorServico > 0 Then
        
            'Obter NumIntDoc
            lErro = CF("Config_ObterNumInt2", alComando(iProxComando + 34), alComando(iProxComando + 38), "FATConfig", "NUM_INT_PROX_RPS", lNumIntDoc)
            If lErro <> SUCESSO Then gError 196814
    
            objRPS.lNumIntDoc = lNumIntDoc
    
            lErro = Comando_Executar(alComando(iProxComando + 35), "INSERT INTO RPS( NumIntDoc, Tipo, Serie, Numero, DataEmissao, FilialEmpresa, " & _
                    "Situacao, ValorServico, ValorDeducao, CodigoServico, Aliquota, " & _
                    "ISSRetido, TipoCGC, CGC, InscricaoMunicipal, InscricaoEstadual, " & _
                    "RazaoSocial, Endereco, EndNumero, EndComplemento, Bairro, " & _
                    "Cidade, UF, CEP, Email, Discriminacao, NumIntDocNF, Cliente, FilialCliente,DataUltAlteracao,HoraUltAlteracao, Enviado) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                    objRPS.lNumIntDoc, objRPS.sTipo, objRPS.sSerie, objRPS.lNumero, objRPS.dtDataEmissao, objRPS.iFilialEmpresa, _
                    objRPS.sSituacao, objRPS.dValorServico, objRPS.dValorDeducao, objRPS.lCodigoServico, objRPS.dAliquota, _
                    objRPS.iISSRetido, objRPS.iTipoCGC, objRPS.sCgc, objRPS.sInscricaoMunicipal, objRPS.sInscricaoEstadual, _
                    objRPS.sRazaoSocial, objRPS.sEndereco, objRPS.sEndNumero, objRPS.sEndComplemento, objRPS.sBairro, _
                    objRPS.sCidade, objRPS.sUF, objRPS.sCEP, objRPS.sEmail, objRPS.sDiscriminacao, objRPS.lNumIntDocNF, objRPS.lCliente, objRPS.iFilialCliente, Date, CDbl(Time), DESMARCADO)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then
                gError 196815
            End If
            
        End If
        
    End If
    
    iProxComando = iProxComando + 43
    
    NF_NFiscal_Grava_TRV = SUCESSO

    Exit Function

Erro_NF_NFiscal_Grava_TRV:

    NF_NFiscal_Grava_TRV = gErr

    Select Case gErr

        Case 196795, 196796
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SERIE1", gErr, objNFiscal.sSerie)

        Case 196797
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_SERIE", gErr)
        
        Case 196798
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_SERIE", gErr, objNFiscal.sSerie)
            
        Case 196799, 196801, 196802, 196804, 196810 To 196814
        
        Case 196800
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_NFISCAL_SAIDA", gErr, objNFiscal.lCliente, objNFiscal.iFilialCli, objNFiscal.iTipoNFiscal, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal)
        
        Case 196803
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITENSNFISCAL", gErr, objNFiscal.lNumNotaFiscal)

        Case 196805, 196806
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIALCLIENTEFILEMP", gErr)
        
        Case 196807
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERIR_FILIALCLIENTEFILEMP", gErr)
        
        Case 196808
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_FILIALCLIENTEFILEMP", gErr)
            
        Case 196809
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_FILIALCLIENTEFILEMP", gErr)
        
        Case 196815
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RPS", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196816)

    End Select

    Exit Function
    
End Function

Private Function NF_SldDiaFat_Grava_Est_TRV(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lComando3 As Long, ByVal objNFiscal As ClassNFiscal, ByVal objTipoDocInfo As ClassTipoDocInfo) As Long
'Grava a estatística do Dia do produto

Dim lErro As Long
Dim objItemNF As New ClassItemNF
Dim objSldDiaFat As New ClassSldDiaFat
Dim objProduto As New ClassProduto
Dim iIndice As Integer

On Error GoTo Erro_NF_SldDiaFat_Grava_Est_TRV
            
    'Para cada objItemNf
    For Each objItemNF In objNFiscal.ColItensNF
                       
        objSldDiaFat.iFilialEmpresa = objNFiscal.iFilialEmpresa
        objSldDiaFat.dtData = objNFiscal.dtDataEmissao
        objSldDiaFat.sProduto = objItemNF.sProduto
            
        'Lê todos os atributos na tabela SldDiaFat e faza lock exclusive
        lErro = CF("SldDiaFat_Le_Lock", lComando1, objSldDiaFat)
        If lErro <> SUCESSO And lErro <> 52729 Then gError 196824
                        
        If (objTipoDocInfo.iFaturamento = TIPODOCINFO_FATURAMENTO_SIM) Then
        
            objSldDiaFat.dValorFaturado = objItemNF.dValorTotal
            objSldDiaFat.dTotalDescontos = objItemNF.dValorDesconto
            
            'Se for uma Nota de complemento de preço
            If objTipoDocInfo.iComplementar = DOCINFO_COMPLEMENTO Then
                'Não altera a quantidade
                objSldDiaFat.dQuantFaturada = 0
            Else
                objSldDiaFat.dQuantFaturada = objItemNF.dQuantidade
            End If
        
        End If
                        
        If lErro = SUCESSO Then 'Registro SldDiaFat existe no BD
                                    
            'Verifica se o documento é TIPODOCINFO_FATURAMENTO_SIM
            If objTipoDocInfo.iFaturamento = TIPODOCINFO_FATURAMENTO_SIM Then
            
                'Atualiza os campos referentes a notas de faturamneto em SldDiaFat
                lErro = Comando_ExecutarPos(lComando2, "UPDATE SldDiaFat SET ValorFaturado = ValorFaturado + ?, QuantFaturada = QuantFaturada + ?, TotalDescontos = TotalDescontos + ?, ValorDevolvido = ?, QuantDevolvida = ?", lComando1, objSldDiaFat.dValorFaturado, objSldDiaFat.dQuantFaturada, objSldDiaFat.dTotalDescontos, objSldDiaFat.dValorDevolvido, objSldDiaFat.dQuantDevolvida)
                If lErro <> AD_SQL_SUCESSO Then gError 196825
        
            End If
        
        Else 'Registro SldDiaFat não existe no BD

            If objNFiscal.iFilialEmpresa = 0 Then objNFiscal.iFilialEmpresa = giFilialEmpresa
            
            lErro = Comando_Executar(lComando3, "INSERT INTO SldDiaFat (FilialEmpresa, Data, Produto, ValorFaturado , QuantFaturada, TotalDescontos, ValorDevolvido, QuantDevolvida) VALUES (?,?,?,?,?,?,?,?)", objNFiscal.iFilialEmpresa, objNFiscal.dtDataEmissao, objItemNF.sProduto, objSldDiaFat.dValorFaturado, objSldDiaFat.dQuantFaturada, objSldDiaFat.dTotalDescontos, objSldDiaFat.dValorDevolvido, objSldDiaFat.dQuantDevolvida)
            If lErro <> AD_SQL_SUCESSO Then gError 196826
                    
        End If
        
    Next
    
    NF_SldDiaFat_Grava_Est_TRV = SUCESSO
    
    Exit Function
    
Erro_NF_SldDiaFat_Grava_Est_TRV:
    
    NF_SldDiaFat_Grava_Est_TRV = gErr

    Select Case gErr

        Case 196824
        
        Case 196825
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_SLDDIAFAT", gErr, objSldDiaFat.iFilialEmpresa, objSldDiaFat.sProduto, objSldDiaFat.dtData)
        
        Case 196826
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SLDDIAFAT", gErr, objNFiscal.iFilialEmpresa, objItemNF.sProduto, objNFiscal.dtDataEmissao)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196827)

    End Select
    
    Exit Function

End Function
        
Private Function NF_SldMesFat_Grava_Est_TRV(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lComando3 As Long, ByVal lComando4 As Long, ByVal lComando5 As Long, ByVal lComando6 As Long, ByVal objNFiscal As ClassNFiscal, ByVal objTipoDocInfo As ClassTipoDocInfo) As Long
'Grava a estatística do Mês do produto

Dim lErro As Long
Dim objItemNF As ClassItemNF
Dim objSldMesFat As New ClassSldMesFat
Dim iMes As Integer
Dim sMes As String
Dim iIndice As Integer

On Error GoTo Erro_NF_SldMesFat_Grava_Est_TRV
            
    'Para cada objItemNf
    For Each objItemNF In objNFiscal.ColItensNF
                        
        objSldMesFat.iFilialEmpresa = objNFiscal.iFilialEmpresa
        objSldMesFat.iAno = Year(objNFiscal.dtDataEmissao)
        objSldMesFat.sProduto = objItemNF.sProduto
        
        iMes = Month(objNFiscal.dtDataEmissao)
        sMes = CStr(iMes)
                                    
        'Lê todos os atributos na tabela SldMesFat e faza lock exclusive
        lErro = CF("SldMesFat_Le_Lock_Mes", lComando1, objSldMesFat, iMes)
        If lErro <> SUCESSO And lErro <> 59264 Then gError 196828
                        
        If (objTipoDocInfo.iFaturamento = TIPODOCINFO_FATURAMENTO_SIM) Then
                    
            objSldMesFat.dValorFaturado(iMes) = objItemNF.dValorTotal
            objSldMesFat.dTotalDescontos(iMes) = objItemNF.dValorDesconto
        
            'Se for uma Nota de complemento de preço
            If objTipoDocInfo.iComplementar = DOCINFO_COMPLEMENTO Then
                'Não altera a quantidade
                objSldMesFat.dQuantFaturada(iMes) = 0
            Else
                objSldMesFat.dQuantFaturada(iMes) = objItemNF.dQuantidade
            End If
        
        End If
        
        If lErro = SUCESSO Then
        
            'Verifica se o tipo de faturamento é TIPODOCINFO_FATURAMENTO_SIM
            If objTipoDocInfo.iFaturamento = TIPODOCINFO_FATURAMENTO_SIM Then
        
                'Atualiza os campos referentes a Notas de Faturamento em SldMesFat
                lErro = Comando_ExecutarPos(lComando2, "UPDATE SldMesFat SET ValorFaturado" & sMes & " = ValorFaturado" & sMes & " + ?, QuantFaturada" & sMes & " = QuantFaturada" & sMes & " + ?, TotalDescontos" & sMes & " = TotalDescontos" & sMes & " + ?, ValorDevolvido" & sMes & " = ?, QuantDevolvida" & sMes & " = ?", lComando1, objSldMesFat.dValorFaturado(iMes), objSldMesFat.dQuantFaturada(iMes), objSldMesFat.dTotalDescontos(iMes), objSldMesFat.dValorDevolvido(iMes), objSldMesFat.dQuantDevolvida(iMes))
                If lErro <> AD_SQL_SUCESSO Then gError 196829
            
            End If
                    
        Else
            
            'Atualiza a tabela de SldMesFat
            lErro = Comando_Executar(lComando3, "INSERT INTO SldMesFat (FilialEmpresa, Ano, Produto, ValorFaturado" & sMes & ", QuantFaturada" & sMes & ", TotalDescontos" & sMes & ", ValorDevolvido" & sMes & ", QuantDevolvida" & sMes & ") VALUES (?,?,?,?,?,?,?,?)", objSldMesFat.iFilialEmpresa, objSldMesFat.iAno, objSldMesFat.sProduto, objSldMesFat.dValorFaturado(iMes), objSldMesFat.dQuantFaturada(iMes), objSldMesFat.dTotalDescontos(iMes), objSldMesFat.dValorDevolvido(iMes), objSldMesFat.dQuantDevolvida(iMes))
            If lErro <> AD_SQL_SUCESSO Then gError 196830
            
        End If
        
        If objTipoDocInfo.iFaturamento = TIPODOCINFO_FATURAMENTO_SIM Then
            
            lErro = NF_SldMesFatFilCli_Grava_Est_TRV(lComando4, lComando5, lComando6, objSldMesFat.iFilialEmpresa, objSldMesFat.iAno, objSldMesFat.sProduto, objNFiscal.lCliente, objNFiscal.iFilialCli, objSldMesFat.dQuantFaturada(iMes), objSldMesFat.dValorFaturado(iMes), 0, 0, sMes)
            If lErro <> SUCESSO Then gError 196831
            
        End If
        
    Next
            
    NF_SldMesFat_Grava_Est_TRV = SUCESSO
    
    Exit Function
    
Erro_NF_SldMesFat_Grava_Est_TRV:
    
    NF_SldMesFat_Grava_Est_TRV = gErr

    Select Case gErr

        Case 196828
        
        Case 196829
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_SLDMESFAT", gErr, objSldMesFat.iAno, objSldMesFat.iFilialEmpresa, objSldMesFat.sProduto)
        
        Case 196830
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SLDMESFAT", gErr, objSldMesFat.iFilialEmpresa, objSldMesFat.sProduto, objSldMesFat.iAno)
        
        Case 196831
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196832)

    End Select
    
    Exit Function

End Function

Private Function NF_SldMesFatFilCli_Grava_Est_TRV(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lComando3 As Long, iFilialEmpresa As Integer, iAno As Integer, sCodProduto As String, lCliente As Long, iFilialCliente As Integer, dQtdFaturada As Double, dValorFaturado As Double, dQtdDevolvida As Double, dValorDevolvido As Double, sMes As String)

Dim lErro As Long
Dim iFilial As Integer

On Error GoTo Erro_NF_SldMesFatFilCli_Grava_Est_TRV
    
    lErro = Comando_ExecutarPos(lComando1, "SELECT FilialEmpresa FROM SldMesFatFilCli WHERE FilialEmpresa = ? AND Ano = ? AND Produto = ? AND Cliente = ? AND Filial = ?", 0, iFilial, iFilialEmpresa, iAno, sCodProduto, lCliente, iFilialCliente)
    If lErro <> AD_SQL_SUCESSO Then gError 196833
    
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196834
    
    If lErro = AD_SQL_SUCESSO Then
    
        'Lock
        lErro = Comando_LockExclusive(lComando1)
        If lErro <> AD_SQL_SUCESSO Then gError 196835
        
        'Atualiza
        lErro = Comando_ExecutarPos(lComando2, "UPDATE SldMesFatFilCli SET QtdFaturada" & sMes & " = QtdFaturada" & sMes & " + ?, ValorFaturado" & sMes & " = ValorFaturado" & sMes & " + ?, QtdDevolvida" & sMes & " = QtdDevolvida" & sMes & " + ?, ValorDevolvido" & sMes & " = ValorDevolvido" & sMes & " + ? ", lComando1, dQtdFaturada, dValorFaturado, dQtdDevolvida, dValorDevolvido)
        If lErro <> AD_SQL_SUCESSO Then gError 196836
    
    Else
    
        lErro = Comando_Executar(lComando3, "INSERT INTO SldMesFatFilCli (FilialEmpresa, Ano, Produto, Cliente, Filial, QtdFaturada" & sMes & ", ValorFaturado" & sMes & ", QtdDevolvida" & sMes & ", ValorDevolvido" & sMes & " ) VALUES (?,?,?,?,?,?,?,?,?)", iFilialEmpresa, iAno, sCodProduto, lCliente, iFilialCliente, dQtdFaturada, dValorFaturado, dQtdDevolvida, dValorDevolvido)
        If lErro <> AD_SQL_SUCESSO Then gError 196837
    
    End If

    NF_SldMesFatFilCli_Grava_Est_TRV = SUCESSO
    
    Exit Function
    
Erro_NF_SldMesFatFilCli_Grava_Est_TRV:

    NF_SldMesFatFilCli_Grava_Est_TRV = gErr
    
    Select Case gErr
  
        Case 196833, 196834
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESFATFILCLI", gErr, iAno, iFilialEmpresa, sCodProduto, lCliente, iFilialCliente)
        
        Case 196835
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_SLDMESFATFILCLI", gErr, iAno, iFilialEmpresa, , sCodProduto, lCliente, iFilialCliente)
        
        Case 196836
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_SLDMESFATFILCLI", gErr, iAno, iFilialEmpresa, sCodProduto, lCliente, iFilialCliente)
        
        Case 196837
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SLDMESFATFILCLI", gErr, iAno, iFilialEmpresa, sCodProduto, lCliente, iFilialCliente)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196838)

    End Select
    
    Exit Function

End Function

Private Function NF_NotaFiscal_Grava_Fis_TRV(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lComando3 As Long, ByVal lComando4 As Long, ByVal lComando5 As Long, ByVal lComando6 As Long, ByVal lComando7 As Long, ByVal lComando8 As Long, ByVal lComando9 As Long, ByVal lComando10 As Long, ByVal lComando11 As Long, ByVal lComando12 As Long, ByVal lComando13 As Long, ByVal lComando14 As Long, ByVal objNFiscal As ClassNFiscal) As Long
'Grava nas Tabelas de Livros Fiscais a Nota Fiscal

Dim lErro As Long
Dim objLivRegES As New ClassLivRegES
Dim objLivRegESEmitentes As New ClassLivRegESEmitentes
Dim colLivRegESItemNF As New Collection
Dim colLivRegESLinha As New Collection
Dim objNFTRV As ClassNFTRV

On Error GoTo Erro_NF_NotaFiscal_Grava_Fis_TRV

    Set objNFTRV = objNFiscal.objInfoUsu
    
    'só vou gerar info para notas que interessem p/livros de ICMS/IPI
    If objNFTRV.objTipoDocInfo.iModeloArqICMS <> 0 Then
  
        'Gera o Registro de LivRegES
        lErro = NF_Saida_Grava_LivRegES_TRV(lComando1, lComando2, lComando3, objNFiscal, objLivRegES, objNFTRV.objTipoDocInfo, objNFTRV.objCliente, objNFTRV.objFilialCliente, objNFTRV.objEnderecoCli)
        If lErro <> SUCESSO Then gError 196838
        
        'Gera as linhas para livros fiscais
        lErro = NF_NotaFiscal_GeraLinhas_Fis_TRV(lComando4, lComando5, lComando6, lComando7, lComando8, lComando11, lComando12, objNFTRV.objTipoDocInfo.iModeloArqICMS, colLivRegESLinha, colLivRegESItemNF, objNFiscal, objLivRegES)
        If lErro <> SUCESSO Then gError 196839
        
        'Grava as Coleções, que estão ordenadas e completadas com o Complemento
        lErro = NF_Grava_LivRegES_Colecao_TRV(lComando9, lComando10, lComando13, lComando14, colLivRegESLinha, colLivRegESItemNF)
        If lErro <> SUCESSO Then gError 196840

    End If
    
    NF_NotaFiscal_Grava_Fis_TRV = SUCESSO
    
    Exit Function
    
Erro_NF_NotaFiscal_Grava_Fis_TRV:

    NF_NotaFiscal_Grava_Fis_TRV = gErr
    
    Select Case gErr
        
        Case 196838, 196839, 196840
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196841)

    End Select
    
    Exit Function
        
End Function

Private Function NF_Saida_Grava_LivRegES_TRV(ByVal lComando As Long, ByVal lComando2 As Long, ByVal lComando3 As Long, ByVal objNFiscal As ClassNFiscal, ByVal objLivRegES As ClassLivRegES, ByVal objTipoDocInfo As ClassTipoDocInfo, ByVal objCliente As ClassCliente, ByVal objFilialCliente As ClassFilialCliente, ByVal objEndereco As ClassEndereco) As Long
'Preenche o objLivRegES com os dados da Nota Fiscal

Dim lErro As Long
Dim lNumIntDoc As Long
Dim objFilialEmpresa As New AdmFiliais
Dim sInscricaoAux As String

On Error GoTo Erro_NF_Saida_Grava_LivRegES_TRV
    
    objLivRegES.iFilialEmpresa = objNFiscal.iFilialEmpresa
    objLivRegES.dtData = objNFiscal.dtDataEmissao
    objLivRegES.dtDataEmissao = objNFiscal.dtDataEmissao
    objLivRegES.iDestinatario = objTipoDocInfo.iDestinatario
    objLivRegES.iEmitente = objTipoDocInfo.iEmitente
    objLivRegES.iModelo = objTipoDocInfo.iModeloArqICMS
    objLivRegES.iOrigem = objTipoDocInfo.iOrigem
    objLivRegES.iSituacao = SITUACAO_NORMAL
    objLivRegES.iStatusLivro = STATUS_FIS_ORIGINAL
    objLivRegES.iTipo = TIPO_REGES_SAIDA
    objLivRegES.iFreteRespons = objNFiscal.iFreteRespons
    
    'inclui este teste em funcao do conhecimentos de transporte
    If Not (objNFiscal.objTributacaoNF Is Nothing) Then
        objLivRegES.dPISValor = objNFiscal.objTributacaoNF.dPISCredito
        objLivRegES.dCOFINSValor = objNFiscal.objTributacaoNF.dCOFINSCredito
    Else
        objLivRegES.dPISValor = 0
        objLivRegES.dCOFINSValor = 0
    End If
    
    'Gera número automático para o Item de apuração ICMS
    lErro = CF("Config_ObterNumInt2", lComando2, lComando3, "FISConfig", "FISC_NUM_INT_LIVREGES", lNumIntDoc)
    If lErro <> SUCESSO Then gError 196842
    
    objLivRegES.lNumIntDocOrigem = objNFiscal.lNumIntDoc
    objLivRegES.lNumIntDoc = lNumIntDoc
    objLivRegES.iTipoNumIntDocOrigem = TIPO_NUMINTDOC_ORIGEM_NOTAFISCAL
    objLivRegES.lNumIntNF = objNFiscal.lNumIntDoc
    objLivRegES.lNumNotaFiscal = objNFiscal.lNumNotaFiscal
    objLivRegES.sSerie = objNFiscal.sSerie
        
    'Arrumar uma solução viavel
    objLivRegES.sCgc = objFilialCliente.sCgc
    Call Formata_String_Numero(objFilialCliente.sInscricaoEstadual, sInscricaoAux)
    objLivRegES.sInscricaoEstadual = sInscricaoAux
    objLivRegES.sNome = objCliente.sRazaoSocial
    
    If objEndereco.iCodigoPais <> PAIS_BRASIL Then
        objLivRegES.sUF = "EX"
    Else
        objLivRegES.sUF = objEndereco.sSiglaEstado
    End If
    
    'Faz a inserção na Tabela de LivResES
    lErro = Comando_Executar(lComando, "INSERT INTO LivRegES (NumIntDoc, TipoNumIntDocOrigem, NumIntDocOrigem, StatusLivro, Tipo, CGC, InscricaoEstadual, Nome, Data, DataEmissao, UF, Serie, NumNotaFiscal, Situacao, Modelo, Emitente, Destinatario, Origem, NumIntEmitente, FilialEmpresa, CIF_FOB, PISValor, COFINSValor) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", objLivRegES.lNumIntDoc, objLivRegES.iTipoNumIntDocOrigem, objLivRegES.lNumIntDocOrigem, objLivRegES.iStatusLivro, objLivRegES.iTipo, objLivRegES.sCgc, objLivRegES.sInscricaoEstadual, objLivRegES.sNome, objLivRegES.dtData, objLivRegES.dtDataEmissao, objLivRegES.sUF, objLivRegES.sSerie, objLivRegES.lNumNotaFiscal, objLivRegES.iSituacao, objLivRegES.iModelo, objLivRegES.iEmitente, objLivRegES.iDestinatario, objLivRegES.iOrigem, objLivRegES.lNumIntEmitente, objLivRegES.iFilialEmpresa, objLivRegES.iFreteRespons, objLivRegES.dPISValor, objLivRegES.dCOFINSValor)
    If lErro <> AD_SQL_SUCESSO Then gError 196843
                                
    NF_Saida_Grava_LivRegES_TRV = SUCESSO
    
    Exit Function
    
Erro_NF_Saida_Grava_LivRegES_TRV:

    NF_Saida_Grava_LivRegES_TRV = gErr

    Select Case gErr

        Case 196842
        
        Case 196843
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_LIVREGES", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196844)

    End Select
    
    Exit Function

End Function

Private Function NF_NotaFiscal_GeraLinhas_Fis_TRV(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lComando3 As Long, ByVal lComando4 As Long, ByVal lComando5 As Long, ByVal lComando6 As Long, ByVal lComando7 As Long, ByVal iModeloArqICMS As Integer, ByVal colLivRegESLinha As Collection, ByVal colLivRegESItemNF As Collection, ByVal objNFiscal As ClassNFiscal, ByVal objLivRegES As ClassLivRegES) As Long
'Gera as linhas para livros fiscais de acordo com o valor passado iModeloArqICMS
'Em geral, iModeloArqICMS será 1, o que indica que as linhas serão geradas no Case Else
'Caso o valor seja diferente de 1, significa que será gerada uma linha "especial"

Dim lErro As Long
Dim objItemNF As New ClassItemNF
Dim objLivRegESCadProd As New ClassLivRegESCadProd
Dim objLivRegESLinha As ClassLivRegESLinha
Dim objNFiscalOrig As New ClassNFiscal

On Error GoTo Erro_NF_NotaFiscal_GeraLinhas_Fis_TRV
            
    'Varre os Itens para gerar LivRegLinha, LivRegESNF, LivRegESCadProd
    For Each objItemNF In objNFiscal.ColItensNF
        
        'Grava os dados da Nota para a coleção colLivRegESLinha
        lErro = NF_Move_Tela_LivRegESLinha_Obj_TRV(objNFiscal, objItemNF.objTributacaoItemNF, objItemNF.dValorTotal, colLivRegESLinha, objLivRegES, objItemNF.dValorDesconto, objItemNF.dValorDescGlobal)
        If lErro <> SUCESSO Then gError 196845
        
        'Move os dados do Item NF para a coleção colLivRegESCadProd
        lErro = NF_Grava_LivRegESCadProd_TRV(lComando1, lComando2, lComando3, lComando4, lComando6, objLivRegES.dtData, objItemNF, objLivRegESCadProd, objNFiscal)
        If lErro <> SUCESSO Then gError 196846
        
        'Move os dados do Item NF para a coleção colLivRegESItemNF
        lErro = NF_Move_Memoria_LivRegESItemNF_TRV(lComando5, lComando7, objItemNF, objNFiscal, objLivRegES, objLivRegESCadProd, colLivRegESItemNF)
        If lErro <> SUCESSO Then gError 196847
                
    Next

    NF_NotaFiscal_GeraLinhas_Fis_TRV = SUCESSO
    
    Exit Function

Erro_NF_NotaFiscal_GeraLinhas_Fis_TRV:

    NF_NotaFiscal_GeraLinhas_Fis_TRV = gErr
    
    Select Case gErr
    
        Case 196845 To 196847
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196848)
    
    End Select
    
    Exit Function
        
End Function

Private Function NF_Grava_LivRegES_Colecao_TRV(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lComando3 As Long, ByVal lComando4 As Long, ByVal colLivRegESLinha As Collection, ByVal colLivRegESItemNF As Collection) As Long

Dim lErro As Long
Dim objLivRegESItemNF As New ClassLivRegESItemNF
Dim objLivRegESLinha As New ClassLivRegESLinha
Dim lNumIntDoc As Long

On Error GoTo Erro_NF_Grava_LivRegES_Colecao_TRV
    
    For Each objLivRegESLinha In colLivRegESLinha
        
        'Gera número automático para o Item de apuração ICMS
        lErro = CF("Config_ObterNumInt2", lComando3, lComando4, "FISConfig", "FISC_NUM_INT_LIVREGESLINHA", lNumIntDoc)
        If lErro <> SUCESSO Then gError 196849
    
        objLivRegESLinha.lNumIntDoc = lNumIntDoc
        
        'insere as linhas do Livro Fiscal desta Nota Fiscal
        lErro = Comando_Executar(lComando1, "INSERT INTO LivRegESLinha (NumIntDoc, NumIntDocRegES, NaturezaOp, ValorTotal, ValorICMSBase, ValorICMS, ValorICMSIsentoNaoTrib, ValorICMSOutras, ValorICMSSubstBase, ValorICMSSubstRet, AliquotaICMS, ValorDespAcess, ValorIPI, AliquotaIPI, ValorIPIBase, ValorIPIIsentoNaoTrib, ValorIPIOutras, ClassifContabil, ValorContabil, ObservacaoLivFisc, CodigoICMS, CodigoIPI) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
            objLivRegESLinha.lNumIntDoc, objLivRegESLinha.lNumIntDocRegES, objLivRegESLinha.sNaturezaOp, Arredonda_Moeda(objLivRegESLinha.dValorTotal), Arredonda_Moeda(objLivRegESLinha.dValorICMSBase), Arredonda_Moeda(objLivRegESLinha.dValorICMS), Arredonda_Moeda(objLivRegESLinha.dValorICMSIsentoNaoTrib), Arredonda_Moeda(objLivRegESLinha.dValorICMSOutras), Arredonda_Moeda(objLivRegESLinha.dValorICMSSubstBase), Arredonda_Moeda(objLivRegESLinha.dValorICMSSubstRet), objLivRegESLinha.dAliquotaICMS, Arredonda_Moeda(objLivRegESLinha.dValorDespAcess), Arredonda_Moeda(objLivRegESLinha.dValorIPI), objLivRegESLinha.dAliquotaIPI, Arredonda_Moeda(objLivRegESLinha.dValorIPIBase), _
            Arredonda_Moeda(objLivRegESLinha.dValorIPIIsentoNaoTrib), Arredonda_Moeda(objLivRegESLinha.dValorIPIOutras), objLivRegESLinha.sClassifContabil, Arredonda_Moeda(objLivRegESLinha.dValorContabil), objLivRegESLinha.sObservacaoLivFisc, objLivRegESLinha.iCodigoICMS, objLivRegESLinha.iCodigoIPI)
        If lErro <> AD_SQL_SUCESSO Then gError 196850
    
    Next
    
    For Each objLivRegESItemNF In colLivRegESItemNF
        
        'insere as linhas do Livro Fiscal desta Nota Fiscal
        lErro = Comando_Executar(lComando2, "INSERT INTO LivRegESItemNF (NumIntDoc, NumIntDocRegES, NumItem, CFOP, NumIntCadProd, Quantidade, ValorProduto, ValorDescontoDespAcess, ValorICMSBase, ValorICMSSubstBase, ValorIPI, AliquotaICMS, TipoTribICMS, TipoTribIPI, ValorICMS, ValorIPIBase, AliquotaIPI, RedBaseICMS, RedBaseIPI, TipoTributacao, AliquotaSubst, ValorSubst) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
        objLivRegESItemNF.lNumIntDoc, objLivRegESItemNF.lNumIntDocRegES, objLivRegESItemNF.iNumItem, objLivRegESItemNF.sCFOP, objLivRegESItemNF.lNumIntCadProd, objLivRegESItemNF.dQuantidade, objLivRegESItemNF.dValorProduto, objLivRegESItemNF.dValorDescontoDespAcess, objLivRegESItemNF.dValorICMSBase, objLivRegESItemNF.dValorICMSSubstBase, objLivRegESItemNF.dValorIPI, objLivRegESItemNF.dAliquotaICMS, objLivRegESItemNF.iTipoTribICMS, objLivRegESItemNF.iTipoTribIPI, objLivRegESItemNF.dValorICMS, objLivRegESItemNF.dValorIPIBase, objLivRegESItemNF.dAliquotaIPI, objLivRegESItemNF.dRedBaseICMS, objLivRegESItemNF.dRedBaseIPI, objLivRegESItemNF.iTipoTributacao, objLivRegESItemNF.dAliquotaSubst, objLivRegESItemNF.dValorSubst)
        If lErro <> AD_SQL_SUCESSO Then gError 196851
    
    Next
    
    NF_Grava_LivRegES_Colecao_TRV = SUCESSO
    
    Exit Function
    
Erro_NF_Grava_LivRegES_Colecao_TRV:

    NF_Grava_LivRegES_Colecao_TRV = gErr
    
    Select Case gErr
        
        Case 196849

        Case 196850
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_LIVREGESLINHA", gErr)
        
        Case 196851
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_LIVREGESITEMNF", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196852)

    End Select
    
    Exit Function

End Function

Private Function NF_Move_Tela_LivRegESLinha_Obj_TRV(ByVal objNFiscal As ClassNFiscal, ByVal objTributacaoObj As Object, ByVal dValorObj As Double, ByVal colLivRegESLinha As Collection, ByVal objLivRegES As ClassLivRegES, ByVal dValorDesconto As Double, ByVal dValorDescGlobal As Double) As Long

Dim iEncontrou As Integer
Dim objLivRegESLinha As New ClassLivRegESLinha
Dim objTipoTribICMS As ClassTipoTribICMS
Dim objTipoTribIPI As ClassTipoTribIPI
Dim lErro As Long, dAliquotaIPI As Double, dAliquotaICMS As Double
Dim objProduto As New ClassProduto, dValorIPIBase As Double
Dim dValorICMSBase As Double, dValorICMSOutras As Double
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim lNumIntDoc As Long, dDifBaseObj As Double
Dim objTipoDeTributacao As ClassTipoDeTributacaoMovto, dICMSOutrasAux As Double
Dim sContaContabil As String, dICMSValRedBase As Double, dIPIValRedBase As Double
Dim objNFTRV As ClassNFTRV

On Error GoTo Erro_NF_Move_Tela_LivRegESLinha_Obj_TRV

    Set objNFTRV = objNFiscal.objInfoUsu
    Set objTipoTribICMS = objNFTRV.objTipoTribICMS
    Set objTipoTribIPI = objNFTRV.objTipoTribIPI
    Set objTipoDeTributacao = objNFTRV.objTipoDeTributacaoMovto

    If objTributacaoObj.dICMSCredito <> 0 Then
        dICMSValRedBase = Round(objTributacaoObj.dICMSBase * objTributacaoObj.dICMSPercRedBase, 4)
        dValorICMSBase = Round(objTributacaoObj.dICMSBase - dICMSValRedBase, 4)
        dAliquotaICMS = objTributacaoObj.dICMSAliquota
    End If
    
    If objTributacaoObj.dIPICredito <> 0 Then
        dIPIValRedBase = Round(objTributacaoObj.dIPIBaseCalculo * objTributacaoObj.dIPIPercRedBase, 4)
        dValorIPIBase = Round(objTributacaoObj.dIPIBaseCalculo - dIPIValRedBase, 4)
        dAliquotaIPI = objTributacaoObj.dIPIAliquota
    End If

    'Trata o Preenchimento do objLivRegESLinha
    For Each objLivRegESLinha In colLivRegESLinha
    
        'verifica se já existe na colecao item com as mesmas Natureza, Aliquotas (IPI e ICMS), conta contabil e colunas no livro (IPI e ICMS)
        If objLivRegESLinha.sNaturezaOp = objTributacaoObj.sNaturezaOp And objLivRegESLinha.dAliquotaICMS = dAliquotaICMS And objLivRegESLinha.dAliquotaIPI = dAliquotaIPI And objLivRegESLinha.sClassifContabil = objNFTRV.objProduto.sContaContabil Then
            
            iEncontrou = 1
                    
            Exit For
        
        End If
    
    Next
    
    'Se não encontrou nenhuma Linha com a mesma Natureza e Alíquota
    If iEncontrou = 0 Then
        
        'insere uma Linha
        Set objLivRegESLinha = New ClassLivRegESLinha
        
        objLivRegESLinha.lNumIntDocRegES = objLivRegES.lNumIntDoc
        
        objLivRegESLinha.dAliquotaICMS = dAliquotaICMS
        objLivRegESLinha.sNaturezaOp = objTributacaoObj.sNaturezaOp
        objLivRegESLinha.dAliquotaIPI = dAliquotaIPI
        objLivRegESLinha.iCodigoICMS = objTipoTribICMS.iColunaNoLivroSaida
        objLivRegESLinha.iCodigoIPI = objTipoTribIPI.iColunaNoLivro
        objLivRegESLinha.sClassifContabil = objNFTRV.objProduto.sContaContabil
        
        colLivRegESLinha.Add objLivRegESLinha
        
    End If
            
    objLivRegESLinha.dValorContabil = Round(objLivRegESLinha.dValorContabil + dValorObj + objTributacaoObj.dIPIValor - dValorDescGlobal, 4)
    objLivRegESLinha.dValorTotal = Round(objLivRegESLinha.dValorTotal + dValorObj + objTributacaoObj.dIPIValor - dValorDescGlobal, 4)
    objLivRegESLinha.dValorICMSBase = Round(objLivRegESLinha.dValorICMSBase + dValorICMSBase, 4)
    objLivRegESLinha.dValorIPIBase = Round(objLivRegESLinha.dValorIPIBase + dValorIPIBase, 4)
            
    If objTributacaoObj.dICMSCredito <> 0 Then
        objLivRegESLinha.dValorICMS = Round(objLivRegESLinha.dValorICMS + objTributacaoObj.dICMSValor, 4)
        objLivRegESLinha.dValorICMSIsentoNaoTrib = Round(objLivRegESLinha.dValorICMSIsentoNaoTrib + dICMSValRedBase, 4)
        If objTributacaoObj.dICMSBase < (dValorObj - dValorDescGlobal) Then
            objLivRegESLinha.dValorICMSOutras = Round(objLivRegESLinha.dValorICMSOutras + dValorObj - objTributacaoObj.dICMSBase - dValorDescGlobal, 4)
        End If

    ElseIf objTipoTribICMS.iColunaNoLivroSaida = TIPO_ISENTO_NAO_TRIBUTADO Then
        objLivRegESLinha.dValorICMSIsentoNaoTrib = Round(objLivRegESLinha.dValorICMSIsentoNaoTrib + IIf(objTributacaoObj.dICMSBase <= (dValorObj - dValorDescGlobal), dValorObj - dValorDescGlobal, objTributacaoObj.dICMSBase), 4)
    Else
        dICMSOutrasAux = IIf(objTributacaoObj.dICMSBase <= (dValorObj - dValorDescGlobal), dValorObj - dValorDescGlobal, objTributacaoObj.dICMSBase)
        If objTipoDeTributacao.iICMSBaseComIPI And dICMSOutrasAux > objTributacaoObj.dIPIValor Then
            dICMSOutrasAux = dICMSOutrasAux - objTributacaoObj.dIPIValor
        End If
        objLivRegESLinha.dValorICMSOutras = Round(objLivRegESLinha.dValorICMSOutras + dICMSOutrasAux, 4)
    End If
    
    If gobjCRFAT.iIPIAbateDesconto <> 0 Then
    
        If objTributacaoObj.dIPICredito <> 0 Then
            objLivRegESLinha.dValorIPI = Round(objLivRegESLinha.dValorIPI + objTributacaoObj.dIPIValor, 4)
            objLivRegESLinha.dValorIPIIsentoNaoTrib = Round(objLivRegESLinha.dValorIPIIsentoNaoTrib + dIPIValRedBase, 4)
            If objTributacaoObj.dIPIBaseCalculo < (dValorObj - dValorDescGlobal) Then
                objLivRegESLinha.dValorIPIOutras = Round(objLivRegESLinha.dValorIPIOutras + dValorObj - objTributacaoObj.dIPIBaseCalculo - dValorDescGlobal, 4)
            End If
        ElseIf objTipoTribIPI.iColunaNoLivro = TIPO_ISENTO_NAO_TRIBUTADO Then
            objLivRegESLinha.dValorIPIIsentoNaoTrib = Round(objLivRegESLinha.dValorIPIIsentoNaoTrib + dValorObj - dValorDescGlobal, 4)
        Else
            objLivRegESLinha.dValorIPIOutras = Round(objLivRegESLinha.dValorIPIOutras + dValorObj - dValorDescGlobal, 4)
        End If

    Else
    
        If objTributacaoObj.dIPICredito <> 0 Then
            objLivRegESLinha.dValorIPI = Round(objLivRegESLinha.dValorIPI + objTributacaoObj.dIPIValor, 4)
            objLivRegESLinha.dValorIPIIsentoNaoTrib = Round(objLivRegESLinha.dValorIPIIsentoNaoTrib + dIPIValRedBase, 4)
            If objTributacaoObj.dIPIBaseCalculo < (dValorObj) Then
                objLivRegESLinha.dValorIPIOutras = Round(objLivRegESLinha.dValorIPIOutras + dValorObj - objTributacaoObj.dIPIBaseCalculo, 4)
            End If
        ElseIf objTipoTribIPI.iColunaNoLivro = TIPO_ISENTO_NAO_TRIBUTADO Then
            objLivRegESLinha.dValorIPIIsentoNaoTrib = Round(objLivRegESLinha.dValorIPIIsentoNaoTrib + dValorObj, 4)
        Else
            objLivRegESLinha.dValorIPIOutras = Round(objLivRegESLinha.dValorIPIOutras + dValorObj - dValorDescGlobal, 4) 'alterado em 13/06/07 para abater o desconto
        End If
    
    End If
    
    NF_Move_Tela_LivRegESLinha_Obj_TRV = SUCESSO
     
    Exit Function
    
Erro_NF_Move_Tela_LivRegESLinha_Obj_TRV:

    NF_Move_Tela_LivRegESLinha_Obj_TRV = gErr
     
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196853)
     
    End Select
     
    Exit Function
    
End Function

Private Function NF_Grava_LivRegESCadProd_TRV(ByVal lComando As Long, ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lComando3 As Long, ByVal lComando4 As Long, ByVal dtData As Date, ByVal objItemNF As ClassItemNF, objLivRegESCadProd As ClassLivRegESCadProd, ByVal objNF As ClassNFiscal) As Long
'Preenhe o objLivRegES com os dados do item da Nota Fiscal

Dim lErro As Long
Dim lNumIntDoc As Long
Dim objProduto As New ClassProduto
Dim objTipoTribICMS As New ClassTipoTribICMS
Dim dtDataInicialAnterior As Date
Dim sCodigoIpiAux As String
Dim objNFTRV As ClassNFTRV

On Error GoTo Erro_NF_Grava_LivRegESCadProd_TRV

    Set objNFTRV = objNF.objInfoUsu

    'Lê no BD para ver se este Produto já não está cadastrado na tabela LivResESCadProd
    lErro = Comando_ExecutarPos(lComando, "SELECT NumIntDoc, DataInicial FROM LivRegESCadProd WHERE Produto = ? ORDER BY DataInicial", 0, lNumIntDoc, dtDataInicialAnterior, objItemNF.sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 196854

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196855

    'Se não estiver então preenche o objeto
    If lErro = AD_SQL_SEM_DADOS Then
    
        Set objLivRegESCadProd = New ClassLivRegESCadProd

        objLivRegESCadProd.dAliquotaICMS = 0
        objLivRegESCadProd.dAliquotaIPI = 0
        objLivRegESCadProd.dBaseCalculoICMSSubst = 0
        objLivRegESCadProd.dReducaoBaseCalculoICMS = 0
        objLivRegESCadProd.dtDataFinal = DATA_NULA
        objLivRegESCadProd.dtDataInicial = dtData
        
        'Gera número automático para o Item de apuração ICMS
        lErro = CF("Config_ObterNumInt2", lComando3, lComando4, "FISConfig", "FISC_NUM_INT_LIVREGES_CADPROD", lNumIntDoc)
        If lErro <> SUCESSO Then gError 196856
        
        objLivRegESCadProd.lNumIntDoc = lNumIntDoc
        
        objLivRegESCadProd.sDescricao = Left(objItemNF.sDescricaoItem, 50)
        objLivRegESCadProd.sProduto = objItemNF.sProduto
        objLivRegESCadProd.sSiglaUM = objItemNF.sUnidadeMed
        
        Set objProduto = objNFTRV.objProduto
        Set objTipoTribICMS = objNFTRV.objTipoTribICMS
            
        objLivRegESCadProd.sCodigoNCM = Mid(objProduto.sIPICodigo, 1, 8)
        
        objLivRegESCadProd.sSituacaoTrib = CStr(objProduto.iOrigemMercadoria) & Left(CStr(objTipoTribICMS.iTipoTribCST), 1)
        
        'Faz a inserção na Tabela de LivResESCadProd
        lErro = Comando_Executar(lComando1, "INSERT INTO LivRegESCadProd ( NumIntDoc, Produto, CodigoNCM, Descricao, SiglaUM, SituacaoTrib, AliquotaIPI, AliquotaICMS, ReducaoBaseCalculoICMS, BaseCalculoICMSSubst, DataInicial) VALUES (?,?,?,?,?,?,?,?,?,?,?)", objLivRegESCadProd.lNumIntDoc, objLivRegESCadProd.sProduto, objLivRegESCadProd.sCodigoNCM, objLivRegESCadProd.sDescricao, objLivRegESCadProd.sSiglaUM, objLivRegESCadProd.sSituacaoTrib, objLivRegESCadProd.dAliquotaIPI, objLivRegESCadProd.dAliquotaICMS, objLivRegESCadProd.dReducaoBaseCalculoICMS, objLivRegESCadProd.dBaseCalculoICMSSubst, objLivRegESCadProd.dtDataInicial)
        If lErro <> AD_SQL_SUCESSO Then gError 196857
        
    Else
    
        If dtDataInicialAnterior > dtData Then
        
            'Atualiza a data inicial
            lErro = Comando_ExecutarPos(lComando2, "UPDATE LivRegESCadProd SET DataInicial = ?", lComando, dtData)
            If lErro <> AD_SQL_SUCESSO Then gError 196858
        
        End If
    
    End If
    
    objLivRegESCadProd.lNumIntDoc = lNumIntDoc
    
    NF_Grava_LivRegESCadProd_TRV = SUCESSO
    
    Exit Function
    
Erro_NF_Grava_LivRegESCadProd_TRV:

    NF_Grava_LivRegESCadProd_TRV = gErr

    Select Case gErr

        Case 196854, 196855
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LIVREGESCADPROD", gErr)
        
        Case 196856

        Case 196857
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_LIVREGESCADPROD", gErr)
        
        Case 196858
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_LIVREGESCADPROD", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196859)

    End Select
    
    Exit Function

End Function

Private Function NF_Move_Memoria_LivRegESItemNF_TRV(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal objItemNF As ClassItemNF, ByVal objNFiscal As ClassNFiscal, ByVal objLivRegES As ClassLivRegES, ByVal objLivRegESCadProd As ClassLivRegESCadProd, ByVal colLivRegESItemNF As Collection) As Long
'Preenhe o objLivRegES com os dados do item da Nota Fiscal

Dim lErro As Long
Dim lNumIntDoc As Long
Dim objLivRegESItemNF As New ClassLivRegESItemNF

On Error GoTo Erro_NF_Move_Memoria_LivRegESItemNF_TRV

    Set objLivRegESItemNF = New ClassLivRegESItemNF
    
    With objItemNF.objTributacaoItemNF
        objLivRegESItemNF.dAliquotaICMS = .dICMSAliquota
        objLivRegESItemNF.dQuantidade = objItemNF.dQuantidade
        objLivRegESItemNF.dValorICMSBase = .dICMSBase
        objLivRegESItemNF.dValorICMSSubstBase = .dICMSSubstBase
        objLivRegESItemNF.dValorIPI = .dIPIValor
        objLivRegESItemNF.dValorProduto = Round(objItemNF.dPrecoUnitario * objItemNF.dQuantidade, 2)
        objLivRegESItemNF.iNumItem = objItemNF.iItem
        objLivRegESItemNF.dValorDescontoDespAcess = Round(objItemNF.dValorDesconto + objItemNF.dValorDescGlobal, 2)
        objLivRegESItemNF.iTipoTribICMS = .iICMSTipo
        objLivRegESItemNF.iTipoTribIPI = .iIPITipo
        objLivRegESItemNF.dAliquotaIPI = .dIPIAliquota
        objLivRegESItemNF.dRedBaseICMS = .dICMSPercRedBase
        objLivRegESItemNF.dRedBaseIPI = .dIPIPercRedBase
        objLivRegESItemNF.iTipoTributacao = .iTipoTributacao
        objLivRegESItemNF.dAliquotaSubst = .dICMSSubstAliquota
        objLivRegESItemNF.dValorSubst = .dICMSSubstValor
        objLivRegESItemNF.dValorICMS = .dICMSValor
        objLivRegESItemNF.dValorIPIBase = .dIPIBaseCalculo
    End With
    
    'Gera número automático para o Item de apuração ICMS
    lErro = CF("Config_ObterNumInt2", lComando1, lComando2, "FISConfig", "FISC_NUM_INT_LIVREGES_ITEMNF", lNumIntDoc)
    If lErro <> SUCESSO Then gError 196860
    
    objLivRegESItemNF.lNumIntDoc = lNumIntDoc
    objLivRegESItemNF.sCFOP = objItemNF.objTributacaoItemNF.sNaturezaOp
    
    objLivRegESItemNF.lNumIntCadProd = objLivRegESCadProd.lNumIntDoc
    objLivRegESItemNF.lNumIntDocRegES = objLivRegES.lNumIntDoc
        
    colLivRegESItemNF.Add objLivRegESItemNF
    
    NF_Move_Memoria_LivRegESItemNF_TRV = SUCESSO
    
    Exit Function
    
Erro_NF_Move_Memoria_LivRegESItemNF_TRV:

    NF_Move_Memoria_LivRegESItemNF_TRV = gErr

    Select Case gErr
        
        Case 196860
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196861)

    End Select
    
    Exit Function

End Function

Public Function NF_Converte_RPS_TRV(ByVal objNF As ClassNFiscal, ByVal objRPS As ClassRPS) As Long

Dim lErro As Long
Dim objItemNF As ClassItemNF
Dim sMensagem As String
Dim objProduto As ClassProduto
Dim dValorServico As Double
Dim objProdutoCategoria As ClassProdutoCategoria
Dim bAchou As Boolean
Dim lCodServico As Long
Dim bTemProdutoNaoServico As Boolean
Dim iIndice As Integer
Dim objNFTRV As ClassNFTRV
Dim objCliente As ClassCliente
Dim objFilialCliente As ClassFilialCliente
Dim objEndereco As ClassEndereco
Dim sStringAux As String

On Error GoTo Erro_NF_Converte_RPS_TRV

    Set objNFTRV = objNF.objInfoUsu

    objRPS.dtDataEmissao = objNF.dtDataEmissao
    objRPS.iFilialCliente = objNF.iFilialCli
    objRPS.iFilialEmpresa = objNF.iFilialEmpresa
    objRPS.lCliente = objNF.lCliente
    objRPS.lNumero = objNF.lNumNotaFiscal
    objRPS.lNumIntDocNF = objNF.lNumIntDoc
    objRPS.sSerie = objNF.sSerie
    
    Set objProduto = objNFTRV.objProduto

    sMensagem = ""
    
    bTemProdutoNaoServico = False
    For Each objItemNF In objNF.ColItensNF

        If objProduto.iNatureza = NATUREZA_PROD_SERVICO Then
        
            dValorServico = dValorServico + objItemNF.dValorTotal
                    
            sMensagem = sMensagem & IIf(Len(Trim(sMensagem)) > 0, "||", "") & objItemNF.sDescricaoItem
            
            bAchou = False
            For Each objProdutoCategoria In objProduto.colCategoriaItem
                
                If objProdutoCategoria.sCategoria = PRODUTO_CAT_COD_SERVICO_NFE Then
                
                    If Not IsNumeric(objProdutoCategoria.sItem) Then gError 196863 'Código do serviço não é numérico
                    
                    If lCodServico <> StrParaInt(objProdutoCategoria.sItem) And lCodServico <> 0 Then gError 196864 'Não pode existir diferentes códigos de serviço no mesmo recibo
                    
                    lCodServico = StrParaInt(objProdutoCategoria.sItem)
                    bAchou = True
                    
                    Exit For
                End If
            
            Next
            
            If Not bAchou Then gError 196865 'Produto sem a informação de código do serviço
            
        Else
            bTemProdutoNaoServico = True
        End If
        
    Next
    
    objRPS.sDiscriminacao = sMensagem

    objRPS.dAliquota = objNF.objTributacaoNF.dISSAliquota
    objRPS.dValorDeducao = 0
    objRPS.dValorServico = dValorServico
    
    If bTemProdutoNaoServico Then
        objRPS.sTipo = "RPS-M"
    Else
        objRPS.sTipo = "RPS"
    End If
    
    If objNF.objTributacaoNF.dISSRetido > 0 Then
        objRPS.iISSRetido = 1
    Else
        objRPS.iISSRetido = 2
    End If
    
    objRPS.lCodigoServico = lCodServico
       
    If objNF.iStatus <> STATUS_CANCELADO Then
        objRPS.sSituacao = "T"
    Else
        objRPS.sSituacao = "C"
    End If
                    
    Set objCliente = objNFTRV.objCliente
    Set objFilialCliente = objNFTRV.objFilialCliente
    Set objEndereco = objNFTRV.objEnderecoCli
    
    objRPS.sBairro = objEndereco.sBairro
    
    sStringAux = ""
    Call Formata_String_Numero(objEndereco.sCEP, sStringAux)
    objRPS.sCEP = sStringAux
    
    sStringAux = ""
    Call Formata_String_Numero(objFilialCliente.sCgc, sStringAux)
    objRPS.sCgc = sStringAux
    
    objRPS.sCidade = objEndereco.sCidade
    objRPS.sEmail = objEndereco.sEmail
    objRPS.sEndereco = objEndereco.sEndereco
    
    sStringAux = ""
    Call Formata_String_Numero(objFilialCliente.sInscricaoEstadual, sStringAux)
    objRPS.sInscricaoEstadual = sStringAux
    
    sStringAux = ""
    Call Formata_String_Numero(objFilialCliente.sInscricaoMunicipal, sStringAux)
    objRPS.sInscricaoMunicipal = sStringAux
    
    objRPS.sRazaoSocial = objCliente.sRazaoSocial
    objRPS.sUF = objEndereco.sSiglaEstado
    
    If Len(Trim(objRPS.sCgc)) = 0 Then
        objRPS.iTipoCGC = 3
    ElseIf Len(Trim(objRPS.sCgc)) < 14 Then
        objRPS.iTipoCGC = 1
    Else
        objRPS.iTipoCGC = 2
    End If
    
    NF_Converte_RPS_TRV = SUCESSO

    Exit Function

Erro_NF_Converte_RPS_TRV:

    NF_Converte_RPS_TRV = gErr

    Select Case gErr
           
        Case 196863
            Call Rotina_Erro(vbOKOnly, "ERRO_RPS_CODSERVICO_NAO_NUMERICO", gErr, PRODUTO_CAT_COD_SERVICO_NFE, objProduto.sCodigo, objProdutoCategoria.sItem)
        
        Case 196864
            Call Rotina_Erro(vbOKOnly, "ERRO_RPS_VARIOS_CODSERVICO", gErr)
        
        Case 196865
            Call Rotina_Erro(vbOKOnly, "ERRO_RPS_CODSERVICO_NAO_ENCONTRADO", gErr, objProduto.sCodigo, PRODUTO_CAT_COD_SERVICO_NFE)
         
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196866)

    End Select

    Exit Function

End Function

Function FAT_TituloReceber_Insere_TRV(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lComando3 As Long, ByVal lComando4 As Long, ByVal lComando5 As Long, ByVal lComando6 As Long, ByVal lComando7 As Long, ByVal lComando8 As Long, ByVal lComando9 As Long, ByVal lComando10 As Long, ByVal objTituloReceber As ClassTituloReceber, ByVal colParcelaReceber As colParcelaReceber) As Long
'Insere o Título à Receber na Tabela de Titulos a Receber
    
Dim lErro As Long
Dim lNumIntDoc As Long
Dim objParcelaReceber As ClassParcelaReceber
Dim lNumeroFatura As Long

On Error GoTo Erro_FAT_TituloReceber_Insere_TRV
    
    'Verifica a existencia de um Título no BD com as mesmas características
    lErro = CF("TituloReceber_Testa_Repeticao", objTituloReceber, lComando1)
    If lErro <> SUCESSO Then gError 196890
    
    'Gera o número automático
    lErro = CF("Config_ObterNumInt2", lComando3, lComando4, "CPRConfig", "NUM_PROX_TIT_REC", lNumIntDoc)
    If lErro <> SUCESSO Then gError 196891
    
    objTituloReceber.lNumIntDoc = lNumIntDoc
    
    'Verifica se Cliente e Filial estão cadastrados
    lErro = CF("FilialCliente_Lock1", lComando5, objTituloReceber.lCliente, objTituloReceber.iFilial)
    If lErro <> SUCESSO And lErro <> 43050 Then gError 196892
    
    'Não encontrou a Filial Cliente --> erro
    If lErro = 43050 Then gError 196893

    If objTituloReceber.iFilialEmpresa = 0 Then objTituloReceber.iFilialEmpresa = giFilialEmpresa
    
    'Insere o Título na tabela TitulosRec
    lErro = Comando_Executar(lComando2, "INSERT INTO TitulosRec (NumIntDoc, FilialEmpresa, Cliente, Filial, SiglaDocumento, NumTitulo, DataEmissao, Status, Saldo, NumParcelas, Valor, ValorIRRF, ValorISS, ISSRetido, ValorINSS, INSSRetido, PercJurosDiarios, PercMulta, Observacao, CondicaoPagto, DataRegistro, PISRetido, COFINSRetido, CSLLRetido, Moeda, Ccl, Natureza, CpoGenerico1, ReajustePeriodicidade, ReajusteBase, ReajustadoAte) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                                        lNumIntDoc, objTituloReceber.iFilialEmpresa, objTituloReceber.lCliente, objTituloReceber.iFilial, objTituloReceber.sSiglaDocumento, objTituloReceber.lNumTitulo, objTituloReceber.dtDataEmissao, STATUS_LANCADO, objTituloReceber.dSaldo, objTituloReceber.iNumParcelas, objTituloReceber.dValor, objTituloReceber.dValorIRRF, objTituloReceber.dValorISS, objTituloReceber.dISSRetido, objTituloReceber.dValorINSS, objTituloReceber.iINSSRetido, objTituloReceber.dPercJurosDiario, objTituloReceber.dPercMulta, objTituloReceber.sObservacao, objTituloReceber.iCondicaoPagto, objTituloReceber.dtDataRegistro, objTituloReceber.dPISRetido, objTituloReceber.dCOFINSRetido, objTituloReceber.dCSLLRetido, objTituloReceber.iMoeda, objTituloReceber.sCcl, objTituloReceber.sNatureza, objTituloReceber.lCpoGenerico1, objTituloReceber.iReajustePeriodicidade, objTituloReceber.dtReajusteBase, objTituloReceber.dtReajustadoAte)
    If lErro <> AD_SQL_SUCESSO Then gError 196894

    For Each objParcelaReceber In colParcelaReceber
        objParcelaReceber.lNumIntTitulo = lNumIntDoc
    Next
    
    'Grava as Parcelas a Receber
    lErro = FAT_ParcelasReceber_Insere_TRV(lComando6, lComando7, lComando8, lComando9, lComando10, colParcelaReceber)
    If lErro <> SUCESSO Then gError 196895
    
    FAT_TituloReceber_Insere_TRV = SUCESSO
    
    Exit Function
    
Erro_FAT_TituloReceber_Insere_TRV:

    FAT_TituloReceber_Insere_TRV = gErr
    
    Select Case gErr
    
        Case 196890 To 196892, 196895
    
        Case 196893
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_CADASTRADA", gErr, objTituloReceber.iFilial, objTituloReceber.lCliente)

        Case 196894
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TITULOS_REC", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196896)
            
    End Select
    
    Exit Function

End Function

Private Function FAT_ParcelasReceber_Insere_TRV(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lComando3 As Long, ByVal lComando4 As Long, ByVal lComando5 As Long, ByVal colParcelaReceber As colParcelaReceber, Optional ByVal iNumPrimeiraParcela As Integer = 1) As Long
'Insere as Parcelas de um Titulo passadas por parâmetro
'iNumPrimeiraParcela é util na alteracao de titulo onde apenas as parcelas novas sao inseridas

Dim lErro As Long
Dim iIndice As Integer
Dim objCarteiraCobrador As New ClassCarteiraCobrador
Dim lNumIntDoc As Long
Dim lNumIntCom As Long
Dim colDesconto As colDesconto

On Error GoTo Erro_FAT_ParcelasReceber_Insere_TRV

    'Gera número interno automático da 1a parcela e reserva numeros para outras parcelas, se houverem
    lErro = CF("Config_ObterNumInt2", lComando4, lComando5, "CPRConfig", "NUM_PROX_PARC_REC", lNumIntDoc, colParcelaReceber.Count - iNumPrimeiraParcela + 1)
    If lErro <> SUCESSO Then gError 196897

    'Para cada Parcela
    For iIndice = iNumPrimeiraParcela To colParcelaReceber.Count
    
        colParcelaReceber(iIndice).lNumIntDoc = lNumIntDoc
    
        'Armazena iProxSeqBaixa e Saldo
        colParcelaReceber(iIndice).iProxSeqBaixa = 1
        colParcelaReceber(iIndice).iProxSeqOcorr = 1
        colParcelaReceber(iIndice).dSaldo = colParcelaReceber(iIndice).dValor
        
        If colParcelaReceber(iIndice).iCobrador = 0 Then
            objCarteiraCobrador.iCobrador = COBRADOR_PROPRIA_EMPRESA
        Else
            objCarteiraCobrador.iCobrador = colParcelaReceber(iIndice).iCobrador
        End If
        
        If colParcelaReceber(iIndice).iCarteiraCobranca = 0 Then
            objCarteiraCobrador.iCodCarteiraCobranca = CARTEIRA_CARTEIRA
        Else
            objCarteiraCobrador.iCodCarteiraCobranca = colParcelaReceber(iIndice).iCarteiraCobranca
        End If

        'Insere a Parcela no BD
        With colParcelaReceber(iIndice)
            lErro = Comando_Executar(lComando1, "INSERT INTO ParcelasRec (NumIntDoc,NumIntTitulo,NumParcela,Status,DataVencimento,DataVencimentoReal,Saldo,Valor,Cobrador,CarteiraCobranca,NumTitCobrador,ProxSeqBaixa,ProxSeqOcorr,Desconto1Codigo, Desconto1Ate, Desconto1Valor,Desconto2Codigo, Desconto2Ate, Desconto2Valor,Desconto3Codigo, Desconto3Ate, Desconto3Valor, Aceite, Descontada, Previsao, Observacao, ValorOriginal, TipoPagto, CodConta, DataCredito, DataEmissaoCheque, BancoCheque, AgenciaCheque, ContaCorrenteCheque, NumeroCheque, DataDepositoCheque, AdmMeioPagto, Parcelamento, NumeroCartao, ValidadeCartao, AprovacaoCartao, DataTransacaoCartao) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
            .lNumIntDoc, .lNumIntTitulo, .iNumParcela, STATUS_ABERTO, .dtDataVencimento, .dtDataVencimentoReal, .dSaldo, .dValor, objCarteiraCobrador.iCobrador, objCarteiraCobrador.iCodCarteiraCobranca, .sNumTitCobrador, .iProxSeqBaixa, .iProxSeqOcorr, .iDesconto1Codigo, .dtDesconto1Ate, .dDesconto1Valor, .iDesconto2Codigo, .dtDesconto2Ate, .dDesconto2Valor, .iDesconto3Codigo, .dtDesconto3Ate, .dDesconto3Valor, .iAceite, .iDescontada, .iPrevisao, .sObservacao, .dValorOriginal, .iTipoPagto, .iCodConta, .dtDataCredito, .dtDataEmissaoCheque, .iBancoCheque, .sAgenciaCheque, .sContaCorrenteCheque, .lNumeroCheque, .dtDataDepositoCheque, .iAdmMeioPagto, .iParcelamento, .sNumeroCartao, .dtValidadeCartao, .sAprovacaoCartao, .dtDataTransacaoCartao)
        End With
        If lErro <> AD_SQL_SUCESSO Then gError 196898

        lNumIntDoc = lNumIntDoc + 1
        
        lErro = CF("CarteiraCobrador_AtualizaSaldo", objCarteiraCobrador, 1, colParcelaReceber(iIndice).dSaldo, lComando2, lComando3)
        If lErro <> SUCESSO Then gError 196899
    
    Next
    
    FAT_ParcelasReceber_Insere_TRV = SUCESSO
    
    Exit Function

Erro_FAT_ParcelasReceber_Insere_TRV:

    FAT_ParcelasReceber_Insere_TRV = gErr
    
    Select Case gErr
    
        Case 196897, 196899
        
        Case 196898
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PARCELAS_REC", gErr, colParcelaReceber(iIndice).lNumIntTitulo, iIndice)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196900)
            
    End Select
    
    Exit Function
    
End Function

Private Function TRV_Cliente_Grava_Fornecedor(ByVal objCli As ClassCliente, ByVal objEndereco As ClassEndereco) As Long

Dim lErro As Long
Dim objForn As New ClassFornecedor
Dim objEnderecoForn As New ClassEndereco

On Error GoTo Erro_TRV_Cliente_Grava_Fornecedor

    objForn.lCodigo = objCli.lCodigo
    
    'faz a leitura para carregar info já existente que nao tenha vindo na importacao
    lErro = CF("Fornecedor_Le", objForn)
    If lErro <> SUCESSO And lErro <> 12729 Then gError 196901

    If lErro <> SUCESSO Then
    
        With objEnderecoForn
            .sEndereco = objEndereco.sEndereco
            .sBairro = objEndereco.sBairro
            .sCidade = objEndereco.sCidade
            .sSiglaEstado = objEndereco.sSiglaEstado
            .sCEP = objEndereco.sCEP
            .sTelefone1 = objEndereco.sTelefone1
            .sTelefone2 = objEndereco.sTelefone2
            .sFax = objEndereco.sFax
            .sContato = objEndereco.sContato
            .sEmail = objEndereco.sEmail
            .iCodigoPais = objEndereco.iCodigoPais
        End With
    
        With objForn
            .sNomeReduzido = objCli.sNomeReduzido
            .sRazaoSocial = objCli.sRazaoSocial
            .sCgc = objCli.sCgc
            .sInscricaoEstadual = objCli.sInscricaoEstadual
            .dDesconto = objCli.dDesconto
            .iTipo = objCli.iTipo
            .iAtivo = objCli.iAtivo
            .sInscricaoMunicipal = objCli.sInscricaoMunicipal
        End With
        
        lErro = CF("Fornecedor_Grava_EmTrans", objForn, objEnderecoForn)
        If lErro <> SUCESSO Then gError 196902
    
    End If
    
    TRV_Cliente_Grava_Fornecedor = SUCESSO
    
    Exit Function

Erro_TRV_Cliente_Grava_Fornecedor:

    TRV_Cliente_Grava_Fornecedor = gErr
    
    Select Case gErr
    
        Case 196901, 196902

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196903)
            
    End Select
    
    Exit Function
    
End Function

Private Function TRVFaturamento_Gera_TitPagOver(ByVal objInfoFatura As ClassTRVFATInfoFatura, ByVal objTitPag As ClassTituloPagar, ByVal lProxNum As Long) As Long

Dim lErro As Long
Dim objParcelaPagar As New ClassParcelaPagar
Dim colParcelaPagar As New colParcelaPagar
Dim objContabil As ClassContabil
Dim objCtb As New ClassImpCoinfoFatCPCtb
Dim objContabAutomatica As New ClassContabAutomatica
Dim dtDataVenctoReal As Date
Dim dtDataVencimento As Date

On Error GoTo Erro_TRVFaturamento_Gera_TitPagOver
               
    objTitPag.lFornecedor = objInfoFatura.lCliente
    objTitPag.iFilial = FILIAL_MATRIZ
    objTitPag.iFilialPedCompra = FILIAL_MATRIZ
    objTitPag.lNumTitulo = lProxNum
    objTitPag.dtDataEmissao = objInfoFatura.dtDataEmissao
    objTitPag.dValorTotal = Abs(objInfoFatura.dValor)
    objTitPag.dValorProdutos = Abs(objInfoFatura.dValor)
    objTitPag.dSaldo = Abs(objInfoFatura.dValor)
    objTitPag.iFilialEmpresa = objInfoFatura.iFilialEmpresa
    objTitPag.sSiglaDocumento = TIPODOC_FATURA_OVER
    objTitPag.iNumParcelas = 1
    
'    'O Vencimento ocorre na sexta da semana seguinte
'    Select Case Weekday(dtDataAux)
'
'        Case vbMonday
'            dtDataVencimento = DateAdd("d", 11, objTitPag.dtDataEmissao)
'        Case vbTuesday
'            dtDataVencimento = DateAdd("d", 10, objTitPag.dtDataEmissao)
'        Case vbWednesday
'            dtDataVencimento = DateAdd("d", 9, objTitPag.dtDataEmissao)
'        Case vbThursday
'            dtDataVencimento = DateAdd("d", 8, objTitPag.dtDataEmissao)
'        Case vbFriday
'            dtDataVencimento = DateAdd("d", 7, objTitPag.dtDataEmissao)
'        Case vbSaturday
'            dtDataVencimento = DateAdd("d", 6, objTitPag.dtDataEmissao)
'        Case vbSunday
'            dtDataVencimento = DateAdd("d", 5, objTitPag.dtDataEmissao)
'
'    End Select

    objParcelaPagar.iNumParcela = 1
    objParcelaPagar.dtDataVencimento = objInfoFatura.dtDataVencimento
    
    'Calcula Data Vencimento Real
    lErro = CF("DataVencto_Real", objInfoFatura.dtDataVencimento, dtDataVenctoReal)
    If lErro <> SUCESSO Then gError 190997
    
    objParcelaPagar.dtDataVencimentoReal = dtDataVenctoReal
    objParcelaPagar.dValor = objTitPag.dValorTotal
    objParcelaPagar.iTipoCobranca = TIPO_COBRANCA_CARTEIRA
    objParcelaPagar.iStatus = STATUS_ABERTO

    With objParcelaPagar
        colParcelaPagar.Add .lNumIntDoc, .lNumIntTitulo, .iNumParcela, .iStatus, .dtDataVencimento, .dtDataVencimentoReal, .dSaldo, .dValor, .iPortador, .iProxSeqBaixa, .iTipoCobranca, .iBancoCobrador, .sNossoNumero, .sCodigoDeBarras
    End With

    'Grava o Título a Pagar
    lErro = CF("NFFatPag_Grava_EmTrans", objTitPag, colParcelaPagar, objContabil)
    If lErro <> SUCESSO Then gError 190998
    
    objCtb.sModelo = "Comissão Over"
    
    'gravar a contabilizacao
    lErro = objContabAutomatica.Inicializa_Contab(objCtb, MODULO_CONTASAPAGAR, "CP", objTitPag.dtDataEmissao, objTitPag.dtDataEmissao)
    If lErro <> SUCESSO Then gError 190999
    
    lErro = objContabAutomatica.GeraContabilizacao(objTitPag)
    If lErro <> SUCESSO Then gError 192000

    Set objContabAutomatica = Nothing
    Call objCtb.Encerrando
    Set objCtb = Nothing
                    
    TRVFaturamento_Gera_TitPagOver = SUCESSO

    Exit Function

Erro_TRVFaturamento_Gera_TitPagOver:

    TRVFaturamento_Gera_TitPagOver = gErr

    Select Case gErr
    
        Case 190997 To 190999, 192000

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190991)

    End Select

    Exit Function

End Function

Public Function TRVVouchers_Comissao_Deslibera(ByVal iTipoLiberacao As Integer, ByVal lNumIntDocLiberacao As Long) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim tVouInfo As typeTRVVoucherInfoN

On Error GoTo Erro_TRVVouchers_Comissao_Deslibera

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 196386
    Next
    
    'Alocação de espaço no buffer
    tVouInfo.sTipoDoc = String(STRING_TRV_VOU_INFO_TIPODOC, 0)
    tVouInfo.sTipVou = String(STRING_TRV_VOU_INFO_TIPVOU, 0)
    tVouInfo.sSerie = String(STRING_TRV_VOU_INFO_SERIE, 0)
    tVouInfo.sHistorico = String(STRING_TRV_VOU_HISTORICO, 0)

    'Le a tabelaTRVVoucherInfo para pegar todos registros liberados e não faturados
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc, TipoDoc, TipVou, Serie, NumVou, Seq, " & _
                "Data, TipoDocDestino, NumIntDocDestino, NumIntDocComiInt, Valor, Historico, NumTitulo, Status FROM TRVVoucherInfo WHERE TipoLiberacao = ? AND NumIntDocLiberacao = ? AND NumIntDocDestino = ? ", 0, _
                tVouInfo.lNumIntDoc, tVouInfo.sTipoDoc, tVouInfo.sTipVou, tVouInfo.sSerie, _
                tVouInfo.lNumVou, tVouInfo.iSeq, tVouInfo.dtData, tVouInfo.iTipoDocDestino, tVouInfo.lNumIntDocDestino, _
                tVouInfo.lNumIntDocComiInt, tVouInfo.dValor, tVouInfo.sHistorico, tVouInfo.lNumTitulo, tVouInfo.iStatus, iTipoLiberacao, lNumIntDocLiberacao, 0)
    If lErro <> AD_SQL_SUCESSO Then gError 196387
    
    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196388
    
    'Enquanto existir
    Do While lErro <> AD_SQL_SEM_DADOS

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 196389

        lErro = Comando_ExecutarPos(alComando(1), "UPDATE TRVVoucherInfo SET Status = ?, TipoLiberacao = ?, NumIntDocLiberacao = ?", alComando(0), TRV_VOU_INFO_STATUS_BLOQUEADO, 0, 0)
        If lErro <> AD_SQL_SUCESSO Then gError 196390
        
        'Busca Próximo
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196391
        
    Loop
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TRVVouchers_Comissao_Deslibera = SUCESSO

    Exit Function

Erro_TRVVouchers_Comissao_Deslibera:

    TRVVouchers_Comissao_Deslibera = gErr

    Select Case gErr

        Case 196386
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 196387, 196388, 196391
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVVOUCHERINFO", gErr)

        Case 196389
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRVVOUCHERINFO", gErr)

        Case 196390
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRVVOUCHERINFO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196392)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function BaixaParcRec_Cancelar_Customizado(ByVal lNumIntDocBaixaParcRec As Long) As Long
    BaixaParcRec_Cancelar_Customizado = TRVVouchers_Comissao_Deslibera(TRV_TIPO_LIBERACAO_COMISSAO_BAIXA, lNumIntDocBaixaParcRec)
End Function

Function TRVGerComiInt_Exclui(objTRVGerComiInt As ClassTRVGerComiInt) As Long

Dim lErro As Long
Dim alComando(1 To 7) As Long
Dim iIndice As Integer
Dim lTransacao As Long
Dim lNumIntDoc As Long
Dim lNumReg As Long

On Error GoTo Erro_TRVGerComiInt_Exclui

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 197438
    
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 197439
    Next
    
    lErro = Comando_Executar(alComando(1), "SELECT NumIntDoc FROM TRVGerComiInt WHERE Codigo = ?", lNumIntDoc, objTRVGerComiInt.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 197440
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197441

    If lErro = AD_SQL_SEM_DADOS Then gError 197442

    lErro = Comando_LockExclusive(alComando(1))
    If lErro <> AD_SQL_SUCESSO Then gError 197443

    lErro = Comando_Executar(alComando(2), "SELECT Count(*) FROM Comissoes, TRVGerComiIntDet, TRVGerComiInt WHERE TRVGerComiInt.Codigo = ? AND Comissoes.NumIntCom = TRVGerComiIntDet.NumIntDocComi AND TRVGerComiInt.NumIntDoc = TRVGerComiIntDet.NumIntDocGerComi AND Comissoes.DataBaixa <> ? ", lNumReg, objTRVGerComiInt.lCodigo, DATA_NULA)
    If lErro <> AD_SQL_SUCESSO Then gError 197444
    
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO Then gError 197445

    'se ja existir comissao baixada ==> nao pode exluir
    If lNumReg <> 0 Then gError 197446

    lErro = Comando_Executar(alComando(3), "DELETE FROM Comissoes WHERE NumIntCom IN (SELECT TRVGerComiIntDet.NumIntDocComi FROM TRVGerComiInt, TRVGerComiIntDet WHERE TRVGerComiInt.NumIntDoc = TRVGerComiIntDet.NumIntDocGerComi AND TRVGerComiInt.Codigo = ?)", objTRVGerComiInt.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 197447

    lErro = Comando_Executar(alComando(4), "DELETE FROM TRVVoucherInfoGer WHERE NumIntVoucherInfo IN (SELECT NumIntDoc FROM TRVVoucherInfo, TRVGerComiInt WHERE TRVVoucherInfo.NumIntDocComiInt = TRVGerComiInt.NumIntDoc AND TRVGerComiInt.Codigo = ?)", objTRVGerComiInt.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 197448

    lErro = Comando_Executar(alComando(5), "UPDATE TRVVoucherInfo SET  NumIntDocComiInt = 0 WHERE NumIntComiInt IN (SELECT NumIntDoc FROM TRVGerComiInt WHERE TRVGerComiInt.Codigo = ?)", objTRVGerComiInt.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 197449

    lErro = Comando_Executar(alComando(6), "DELETE FROM TRVGerComiIntDet WHERE NumIntDocGerComi IN (SELECT NumIntDoc FROM TRVGerComiInt WHERE TRVGerComiInt.Codigo = ?)", objTRVGerComiInt.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 197450

    lErro = Comando_Executar(alComando(7), "DELETE FROM TRVGerComiInt WHERE NumIntDoc = ?", objTRVGerComiInt.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 197451

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 197452

    TRVGerComiInt_Exclui = SUCESSO

    Exit Function

Erro_TRVGerComiInt_Exclui:

    TRVGerComiInt_Exclui = gErr

    Select Case gErr

        Case 197438
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 197439
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 197440, 197441
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVGERCOMIINT", gErr)

        Case 197442
            Call Rotina_Erro(vbOKOnly, "ERRO_TRVGERCOMIINT_NAO_CADASTRADO", gErr, objTRVGerComiInt.lCodigo)

        Case 197443
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_TRVGERCOMIINT", gErr, objTRVGerComiInt.lCodigo)

        Case 197444, 197445
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMISSOES3", gErr)

        Case 197446
            Call Rotina_Erro(vbOKOnly, "ERRO_EXC_TRVGERCOMIINT_COMIS_BAIXADA", gErr, objTRVGerComiInt.lCodigo)
        
        Case 197447
            Call Rotina_Erro(vbOKOnly, "ERRO_EXC_COMISSOES_TRVGERCOMIINT", gErr, objTRVGerComiInt.lCodigo)

        Case 197448
            Call Rotina_Erro(vbOKOnly, "ERRO_EXC_TRVVOUCHERINFOGER", gErr, objTRVGerComiInt.lCodigo)

        Case 197449
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_TRVVOUCHERINFO1", gErr, objTRVGerComiInt.lCodigo)

        Case 197450
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRVGERCOMIINTDET", gErr, objTRVGerComiInt.lCodigo)

        Case 197451
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRVGERCOMIINT", gErr, objTRVGerComiInt.lCodigo)

        Case 197452
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197453)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Desfaz Transação
    Call Transacao_Rollback
    
    Exit Function

End Function

Public Function TRVClienteRepProd_Exclui(ByVal objCliente As ClassCliente) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lCliente As Integer

On Error GoTo Erro_TRVClienteRepProd_Exclui

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 196931
    Next

    'Le a tabelaTRVClienteRepProd
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Cliente FROM TRVClienteRepProd WHERE Cliente= ? ", _
                0, lCliente, objCliente.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 196932

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196933

    'Se não existir => ERRO
    Do While lErro <> AD_SQL_SEM_DADOS

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 196934
    
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM TRVClienteRepProd", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 196935
        
        'Busca Próximo
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196936
        
    Loop

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TRVClienteRepProd_Exclui = SUCESSO

    Exit Function

Erro_TRVClienteRepProd_Exclui:

    TRVClienteRepProd_Exclui = gErr

    Select Case gErr

        Case 196931
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 196932, 196933, 196936
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVCLIENTEREPPROD", gErr)
        
        Case 196934
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRVCLIENTEREPPROD", gErr)

        Case 196935
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRVCLIENTEREPPROD", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196937)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function TRVClienteCorProd_Exclui(ByVal objCliente As ClassCliente) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lCliente As Integer

On Error GoTo Erro_TRVClienteCorProd_Exclui

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 196938
    Next

    'Le a tabela TRVClienteRepProd
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Cliente FROM TRVClienteCorProd WHERE Cliente= ? ", _
                0, lCliente, objCliente.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 196939

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196940

    'Se não existir => ERRO
    Do While lErro <> AD_SQL_SEM_DADOS

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 196941
    
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM TRVClienteCorProd", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 196942
        
        'Busca Próximo
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196936
        
    Loop

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TRVClienteCorProd_Exclui = SUCESSO

    Exit Function

Erro_TRVClienteCorProd_Exclui:

    TRVClienteCorProd_Exclui = gErr

    Select Case gErr

        Case 196938
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 196939, 196940, 196943
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVCLIENTECORPROD", gErr)
        
        Case 196941
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRVCLIENTECORPROD", gErr)

        Case 196942
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRVCLIENTECORPROD", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196944)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function TRVClienteRepProd_Insere(ByVal lComando As Long, ByVal objTRVClienteRepProd As ClassTRVClienteRepProd) As Long

Dim lErro As Long

On Error GoTo Erro_TRVClienteRepProd_Insere

    lErro = Comando_Executar(lComando, "INSERT INTO TRVClienteRepProd( Cliente, Seq, Representante, Produto, " & _
            "PercComis) VALUES (?,?,?,?,?)", _
            objTRVClienteRepProd.lCliente, objTRVClienteRepProd.iSeq, objTRVClienteRepProd.lRepresentante, objTRVClienteRepProd.sProduto, _
            objTRVClienteRepProd.dPercComis)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196946

    TRVClienteRepProd_Insere = SUCESSO

    Exit Function

Erro_TRVClienteRepProd_Insere:

    TRVClienteRepProd_Insere = gErr

    Select Case gErr
        
        Case 196946
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRVCLIENTEREPPROD", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196947)

    End Select

    Exit Function

End Function

Public Function TRVClienteCorProd_Insere(ByVal lComando As Long, ByVal objTRVClienteCorProd As ClassTRVClienteCorProd) As Long

Dim lErro As Long

On Error GoTo Erro_TRVClienteCorProd_Insere

    lErro = Comando_Executar(lComando, "INSERT INTO TRVClienteCorProd( Cliente, Seq, Correntista, Produto, " & _
            "PercComis) VALUES (?,?,?,?,?)", _
            objTRVClienteCorProd.lCliente, objTRVClienteCorProd.iSeq, objTRVClienteCorProd.lCorrentista, objTRVClienteCorProd.sProduto, _
            objTRVClienteCorProd.dPercComis)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196949

    TRVClienteCorProd_Insere = SUCESSO

    Exit Function

Erro_TRVClienteCorProd_Insere:

    TRVClienteCorProd_Insere = gErr

    Select Case gErr
    
        Case 196949
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRVCLIENTECORPROD", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196950)

    End Select

    Exit Function

End Function


'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'GRAVAÇÃO
Public Function TRVTiposOcorrencia_Grava(ByVal objTRVTiposOcorrencia As ClassTRVTiposOcorrencia) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_TRVTiposOcorrencia_Grava

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 198017

    lErro = CF("TRVTiposOcorrencia_Grava_EmTrans", objTRVTiposOcorrencia)
    If lErro <> SUCESSO Then gError 198018

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 198019

    TRVTiposOcorrencia_Grava = SUCESSO

    Exit Function

Erro_TRVTiposOcorrencia_Grava:

    TRVTiposOcorrencia_Grava = gErr

    Select Case gErr

        Case 198017
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 198018

        Case 198019
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198020)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function TRVTiposOcorrencia_Grava_EmTrans(ByVal objTRVTiposOcorrencia As ClassTRVTiposOcorrencia) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim iCodigo As Integer

On Error GoTo Erro_TRVTiposOcorrencia_Grava_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 198021
    Next

    'Le a tabelaTRVTiposOcorrencia
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Codigo FROM TRVTiposOcorrencia WHERE Codigo= ? ", _
                0, iCodigo, objTRVTiposOcorrencia.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 198022

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 198023

    'Se existir => UPDATE, senão => INSERT
    If lErro = AD_SQL_SUCESSO Then

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 198024

        lErro = Comando_ExecutarPos(alComando(1), "UPDATE TRVTiposOcorrencia SET Descricao= ? , EstornaAporteVou= ? , ConsideraComisInt= ? , AlteraComiVou= ? , " & _
                "AlteraCMCC= ? , AlteraCMC= ? , AlteraCMR= ? , AlteraOVER= ? , AlteraCMA= ? , " & _
                "AceitaVlrPositivo= ? , AceitaVlrNegativo= ? ", alComando(0), _
                objTRVTiposOcorrencia.sDescricao, objTRVTiposOcorrencia.iEstornaAporteVou, objTRVTiposOcorrencia.iConsideraComisInt, objTRVTiposOcorrencia.iAlteraComiVou, _
                objTRVTiposOcorrencia.iAlteraCMCC, objTRVTiposOcorrencia.iAlteraCMC, objTRVTiposOcorrencia.iAlteraCMR, objTRVTiposOcorrencia.iAlteraOVER, objTRVTiposOcorrencia.iAlteraCMA, _
                objTRVTiposOcorrencia.iAceitaVlrPositivo, objTRVTiposOcorrencia.iAceitaVlrNegativo)
        If lErro <> AD_SQL_SUCESSO Then gError 198025

    Else


        lErro = Comando_Executar(alComando(1), "INSERT INTO TRVTiposOcorrencia( Codigo, Descricao, EstornaAporteVou, ConsideraComisInt, " & _
                "AlteraComiVou, AlteraCMCC, AlteraCMC, AlteraCMR, AlteraOVER, " & _
                "AlteraCMA, AceitaVlrPositivo, AceitaVlrNegativo) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)", _
                objTRVTiposOcorrencia.iCodigo, objTRVTiposOcorrencia.sDescricao, objTRVTiposOcorrencia.iEstornaAporteVou, objTRVTiposOcorrencia.iConsideraComisInt, _
                objTRVTiposOcorrencia.iAlteraComiVou, objTRVTiposOcorrencia.iAlteraCMCC, objTRVTiposOcorrencia.iAlteraCMC, objTRVTiposOcorrencia.iAlteraCMR, objTRVTiposOcorrencia.iAlteraOVER, _
                objTRVTiposOcorrencia.iAlteraCMA, objTRVTiposOcorrencia.iAceitaVlrPositivo, objTRVTiposOcorrencia.iAceitaVlrNegativo)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 198026

    End If

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TRVTiposOcorrencia_Grava_EmTrans = SUCESSO

    Exit Function

Erro_TRVTiposOcorrencia_Grava_EmTrans:

    TRVTiposOcorrencia_Grava_EmTrans = gErr

    Select Case gErr

        Case 198021
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 198022, 198023
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVTIPOSOCORRENCIA", gErr)

        Case 198024
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRVTIPOSOCORRENCIA", gErr)

        Case 198025
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRVTIPOSOCORRENCIA", gErr)

        Case 198026
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRVTIPOSOCORRENCIA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198027)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'EXCLUSÃO
Public Function TRVTiposOcorrencia_Exclui(ByVal objTRVTiposOcorrencia As ClassTRVTiposOcorrencia) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_TRVTiposOcorrencia_Exclui

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 198006

    lErro = CF("TRVTiposOcorrencia_Exclui_EmTrans", objTRVTiposOcorrencia)
    If lErro <> SUCESSO Then gError 198007

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 198008

    TRVTiposOcorrencia_Exclui = SUCESSO

    Exit Function

Erro_TRVTiposOcorrencia_Exclui:

    TRVTiposOcorrencia_Exclui = gErr

    Select Case gErr

        Case 198006
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 198007

        Case 198008
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198009)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function TRVTiposOcorrencia_Exclui_EmTrans(ByVal objTRVTiposOcorrencia As ClassTRVTiposOcorrencia) As Long

Dim lErro As Long
Dim alComando(0 To 2) As Long
Dim iIndice As Integer
Dim iCodigo As Integer
Dim lCodigoOcr As Long

On Error GoTo Erro_TRVTiposOcorrencia_Exclui_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 198010
    Next

    'Le a tabelaTRVTiposOcorrencia
    lErro = Comando_ExecutarPos(alComando(2), "SELECT Codigo FROM TRVOcorrencias WHERE Origem= ? ", _
                0, lCodigoOcr, objTRVTiposOcorrencia.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 198089

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 198090

    'Se não existir => ERRO
    If lErro <> AD_SQL_SEM_DADOS Then gError 198091

    'Le a tabelaTRVTiposOcorrencia
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Codigo FROM TRVTiposOcorrencia WHERE Codigo= ? ", _
                0, iCodigo, objTRVTiposOcorrencia.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 198011

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 198012

    'Se não existir => ERRO
    If lErro = AD_SQL_SEM_DADOS Then gError 198013

    'Trava registro contra alterações/Leituras
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> SUCESSO Then gError 198014

    lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM TRVTiposOcorrencia", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 198015

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TRVTiposOcorrencia_Exclui_EmTrans = SUCESSO

    Exit Function

Erro_TRVTiposOcorrencia_Exclui_EmTrans:

    TRVTiposOcorrencia_Exclui_EmTrans = gErr

    Select Case gErr

        Case 198010
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 198011, 198012
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVTIPOSOCORRENCIA", gErr)

        Case 198013
            Call Rotina_Erro(vbOKOnly, "ERRO_TRVTIPOSOCORRENCIA_NAO_CADASTRADO", gErr, objTRVTiposOcorrencia.iCodigo)

        Case 198014
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRVTIPOSOCORRENCIA", gErr)

        Case 198015
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRVTIPOSOCORRENCIA", gErr)
            
        Case 198089, 198090
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVOCORRENCIAS", gErr)

        Case 198091
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_ASSOC_OCR", gErr, objTRVTiposOcorrencia.iCodigo, lCodigoOcr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198016)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFATGravaTRP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'GRAVAÇÃO
Public Function TRPOcorrencias_Grava(ByVal objOcr As ClassTRPOcorrencias) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_TRPOcorrencias_Grava

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 189450

    lErro = CF("TRPOcorrencias_Grava_EmTrans", objOcr)
    If lErro <> SUCESSO Then gError 189451

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 189452

    TRPOcorrencias_Grava = SUCESSO

    Exit Function

Erro_TRPOcorrencias_Grava:

    TRPOcorrencias_Grava = gErr

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

Public Function TRPOcorrencias_Grava_EmTrans(ByVal objOcr As ClassTRPOcorrencias) As Long

Dim lErro As Long
Dim alComando(0 To 7) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim lNumIntDocDest As Long
Dim iFormaPagto As Integer
Dim objVoucher As New ClassTRPVouchers
Dim sSerie As String
Dim sTipo As String
Dim lNumVou As Long
Dim iStatusVou As Integer
Dim dValorOcr As Double
Dim dValorOcrAtu As Double
Dim dValorAnterior As Double
Dim BOrcNova As Boolean
Dim iOrigem As Integer
Dim dValorOCRBrutoVou As Double
Dim dValorOCRCMAVou As Double
Dim dValorBrutoVou As Double
Dim dValorCMAVou As Double
Dim iCartao As Integer

On Error GoTo Erro_TRPOcorrencias_Grava_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 189454
    Next
       
    sSerie = String(STRING_TRP_OCR_SERIE, 0)
    sTipo = String(STRING_TRP_OCR_TIPOVOU, 0)

    'Le a tabelaTRPOcorrencias
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc,NumIntDocDestino,FormaPagto, Serie, TipoDoc, NumVou, ValorTotal, Origem,ValorOCRBrutoVou,ValorOCRCMAVou FROM TRPOcorrencias WHERE Codigo= ? ", _
                0, lNumIntDoc, lNumIntDocDest, iFormaPagto, sSerie, sTipo, lNumVou, dValorAnterior, iOrigem, dValorOCRBrutoVou, dValorOCRCMAVou, objOcr.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 189455

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189456

    If lErro = AD_SQL_SUCESSO Then
        BOrcNova = False
        If UCase(sSerie) <> UCase(objOcr.sSerie) Or UCase(sTipo) <> UCase(objOcr.sTipoDoc) Or lNumVou <> objOcr.lNumVou Then gError 202790
    Else
        BOrcNova = True
        dValorAnterior = 0
    End If
    
    'Le a tabela TRPVoucher
    lErro = Comando_ExecutarPos(alComando(6), "SELECT Status, ValorOcr, ValorComissao, ValorBruto, Cartao FROM TRPVouchers WHERE NumVou = ? AND Serie = ? AND TipVou = ?  ", _
                0, iStatusVou, dValorOcr, dValorCMAVou, dValorBrutoVou, iCartao, objOcr.lNumVou, objOcr.sSerie, objOcr.sTipoDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 194437

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(6))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194438

    If lErro <> AD_SQL_SUCESSO Then gError 194441

    'Trava registro contra alterações/Leituras
    lErro = Comando_LockExclusive(alComando(6))
    If lErro <> SUCESSO Then gError 194439
    
    If iCartao = MARCADO Then dValorCMAVou = 0
    
    dValorOcrAtu = dValorOcr + objOcr.dValorTotal - dValorAnterior

    lErro = Comando_ExecutarPos(alComando(7), "UPDATE TRPVouchers SET ValorOcr= ? , TemOcr= ? ", alComando(6), dValorOcrAtu, MARCADO)
    If lErro <> SUCESSO Then gError 194440
    
    'Se existir => UPDATE, senão => INSERT
    If Not BOrcNova Then
    
        If iOrigem <> objOcr.iOrigem Then gError 198220
        If Abs(dValorOCRBrutoVou - objOcr.dValorOCRBrutoVou) > DELTA_VALORMONETARIO Then gError 200700
        If Abs(dValorOCRCMAVou - objOcr.dValorOCRCMAVou) > DELTA_VALORMONETARIO Then gError 200700
    
        If lNumIntDocDest <> 0 Then
            If iFormaPagto = FORMAPAGTO_TRP_OCR_FAT Then
                gError 190310
            Else
                gError 190311
            End If
        End If
        
        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 189457
        
        objOcr.lNumIntDoc = lNumIntDoc
            
        lErro = Comando_ExecutarPos(alComando(1), "UPDATE TRPOcorrencias SET NumVou= ? , TipoDoc= ? , Serie= ? , Cliente= ? , " & _
                "FilialCliente= ? , DataEmissao= ? , Observacao= ? , Status= ? , Origem= ? , " & _
                "Historico= ? , FormaPagto= ? , ValorTotal= ? ", alComando(0), _
                objOcr.lNumVou, objOcr.sTipoDoc, objOcr.sSerie, objOcr.lCliente, _
                objOcr.iFilialCliente, objOcr.dtDataEmissao, objOcr.sObservacao, objOcr.iStatus, objOcr.iOrigem, _
                objOcr.sHistorico, objOcr.iFormaPagto, objOcr.dValorTotal)
        If lErro <> AD_SQL_SUCESSO Then gError 189458
        
        lErro = TRPOcorrenciaDet_Exclui(alComando(2), alComando(3), objOcr)
        If lErro <> SUCESSO Then gError 189473

    Else

        'Obter NumIntDoc
        lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_INT_PROX_TRPOCORRENCIAS", lNumIntDoc)
        If lErro <> SUCESSO Then gError 189453

        objOcr.lNumIntDoc = lNumIntDoc
        
        If objOcr.iOrigem = INATIVACAO_AUTOMATICA_CODIGO Then
            objOcr.dValorOCRBrutoVou = -dValorBrutoVou
            objOcr.dValorOCRCMAVou = -dValorCMAVou
        End If
        
        lErro = Comando_Executar(alComando(1), "INSERT INTO TRPOcorrencias( NumIntDoc, NumVou, TipoDoc, Serie, " & _
                "Codigo, Cliente, FilialCliente, DataEmissao, Observacao, " & _
                "Status, Origem, Historico, FormaPagto, " & _
                "ValorTotal,ValorOCRBrutoVou,ValorOCRCMAVou) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                objOcr.lNumIntDoc, objOcr.lNumVou, objOcr.sTipoDoc, objOcr.sSerie, _
                objOcr.lCodigo, objOcr.lCliente, objOcr.iFilialCliente, objOcr.dtDataEmissao, objOcr.sObservacao, _
                objOcr.iStatus, objOcr.iOrigem, objOcr.sHistorico, objOcr.iFormaPagto, _
                objOcr.dValorTotal, objOcr.dValorOCRBrutoVou, objOcr.dValorOCRCMAVou)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189459

    End If
    
    lErro = TRPOcorrenciaDet_Insere(alComando(4), objOcr)
    If lErro <> SUCESSO Then gError 189474
       
    If BOrcNova Then
    
        objVoucher.lNumVou = objOcr.lNumVou
        objVoucher.sSerie = objOcr.sSerie
        objVoucher.sTipVou = objOcr.sTipoDoc
        
        lErro = CF("TRPVouchers_Le", objVoucher)
        If lErro <> SUCESSO Then gError 196312

        'Tem que gerar a comissão
        If objOcr.iOrigem = INATIVACAO_AUTOMATICA_CODIGO Then
            lErro = TRPVou_Gera_Comis(objVoucher, objOcr.dValorTotal - dValorAnterior, TRP_TIPO_VALOR_BASE_LIQ, objOcr.dtDataEmissao, "Inativação do Voucher", TRP_TIPO_TRATAMENTO_COMI_NVL, objOcr)
        Else
            lErro = TRPVou_Gera_Comis(objVoucher, objOcr.dValorTotal - dValorAnterior, TRP_TIPO_VALOR_BASE_LIQ, objOcr.dtDataEmissao, "Ocorrência código " & CStr(objOcr.lCodigo), TRP_TIPO_TRATAMENTO_COMI_OCR, objOcr)
        End If
        If lErro <> SUCESSO Then gError 196312
        
    End If
        
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TRPOcorrencias_Grava_EmTrans = SUCESSO

    Exit Function

Erro_TRPOcorrencias_Grava_EmTrans:

    TRPOcorrencias_Grava_EmTrans = gErr

    Select Case gErr

        Case 189453, 189473, 189474, 196312

        Case 189454
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 189455, 189456
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPOCORRENCIAS", gErr)

        Case 189457
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRPOCORRENCIAS", gErr)

        Case 189458
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRPOCORRENCIAS", gErr)

        Case 189459
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRPOCORRENCIAS", gErr)

        Case 190310
            Call Rotina_Erro(vbOKOnly, "ERRO_TRPOCORRENCIAS_FATURADA", gErr)

        Case 190311
            Call Rotina_Erro(vbOKOnly, "ERRO_TRPOCORRENCIAS_DEBITO", gErr)
            
        Case 192874
        
        Case 192875
            Call Rotina_Erro(vbOKOnly, "ERRO_VOUCHER_JA_CANCELADO", gErr)
            
        Case 194437, 194438
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPVOUCHERS", gErr)

        Case 194439
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRPVOUCHERS", gErr)

        Case 194440
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRPVOUCHERS", gErr)

        Case 194441
            Call Rotina_Erro(vbOKOnly, "ERRO_TRPVOUCHERS_NAO_CADASTRADO", gErr, lNumVou, sTipo, sSerie)

        Case 198220
            Call Rotina_Erro(vbOKOnly, "ERRO_TRP_TIPOOCR_NAO_ALTERAVEL", gErr)
            
        Case 200700
            Call Rotina_Erro(vbOKOnly, "ERRO_OCR_VALOR_ALTERADO", gErr)

        Case 202790
            Call Rotina_Erro(vbOKOnly, "ERRO_ORC_VOU_ALTERADO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189479)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function TRPOcorrenciaDet_Insere(ByVal lComando As Long, ByVal objTRPOcorrencia As ClassTRPOcorrencias) As Long

Dim lErro As Long
Dim lNumIntDoc As Long
Dim objTRPOcorrenciaDet As ClassTRPOcorrenciaDet

On Error GoTo Erro_TRPOcorrenciaDet_Insere

    For Each objTRPOcorrenciaDet In objTRPOcorrencia.colDetalhamento

        'Obter NumIntDoc
        lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_INT_PROX_TRPOCORRENCIADET", lNumIntDoc)
        If lErro <> SUCESSO Then gError 189476
    
        objTRPOcorrenciaDet.lNumIntDoc = lNumIntDoc
        objTRPOcorrenciaDet.lNumIntDocOCR = objTRPOcorrencia.lNumIntDoc
    
        lErro = Comando_Executar(lComando, "INSERT INTO TRPOcorrenciaDet( NumIntDoc, NumIntDocOcr, Tipo, Valor, Seq) VALUES (?,?,?,?,?)", _
                objTRPOcorrenciaDet.lNumIntDoc, objTRPOcorrenciaDet.lNumIntDocOCR, objTRPOcorrenciaDet.iTipo, objTRPOcorrenciaDet.dValor, objTRPOcorrenciaDet.iSeq)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189482

    Next

    TRPOcorrenciaDet_Insere = SUCESSO

    Exit Function

Erro_TRPOcorrenciaDet_Insere:

    TRPOcorrenciaDet_Insere = gErr

    Select Case gErr

        Case 189476

        Case 189477
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRPOCORRENCIADET", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189478)

    End Select

    Exit Function

End Function

Public Function TRPOcorrenciaDet_Exclui(ByVal lComando1 As Long, lComando2 As Long, ByVal objTRPOcorrencia As ClassTRPOcorrencias) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long

On Error GoTo Erro_TRPOcorrenciaDet_Exclui

    'Le a tabela TRPOcorrenciaDet
    lErro = Comando_ExecutarPos(lComando1, "SELECT NumIntDoc FROM TRPOcorrenciaDet WHERE NumIntDocOcr= ? ", _
                0, lNumIntDoc, objTRPOcorrencia.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 189468

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189469

    Do While lErro <> AD_SQL_SEM_DADOS

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(lComando1)
        If lErro <> SUCESSO Then gError 189471
    
        lErro = Comando_ExecutarPos(lComando2, "DELETE FROM TRPOcorrenciaDet", lComando1)
        If lErro <> AD_SQL_SUCESSO Then gError 189472

        'Busca Próximo
        lErro = Comando_BuscarProximo(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189470

    Loop

    TRPOcorrenciaDet_Exclui = SUCESSO

    Exit Function

Erro_TRPOcorrenciaDet_Exclui:

    TRPOcorrenciaDet_Exclui = gErr

    Select Case gErr

        Case 189468, 189469, 189470
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPOCORRENCIADET", gErr)

        Case 189471
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRPOCORRENCIADET", gErr)

        Case 189472
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRPOCORRENCIADET", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189480)

    End Select

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'EXCLUSÃO
Public Function TRPOcorrencias_Exclui(ByVal objOcr As ClassTRPOcorrencias) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_TRPOcorrencias_Exclui

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 189441

    lErro = CF("TRPOcorrencias_Exclui_EmTrans", objOcr)
    If lErro <> SUCESSO Then gError 189442

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 189443

    TRPOcorrencias_Exclui = SUCESSO

    Exit Function

Erro_TRPOcorrencias_Exclui:

    TRPOcorrencias_Exclui = gErr

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

Public Function TRPOcorrencias_Exclui_EmTrans(ByVal objOcr As ClassTRPOcorrencias, Optional ByVal bInativacaoExclui As Boolean = False) As Long

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
Dim objVoucher As New ClassTRPVouchers
Dim dValorTotal As Double
Dim dtDataEmissao As Date
Dim lCliente As Long
Dim objOCRBD As New ClassTRPOcorrencias

On Error GoTo Erro_TRPOcorrencias_Exclui_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 189444
    Next
    
    sSerie = String(STRING_TRP_OCR_SERIE, 0)
    sTipoDoc = String(STRING_TRP_OCR_TIPOVOU, 0)
    
    'Le a tabela TRPOcorrencias
    If Not bInativacaoExclui Then
        If objOcr.lNumIntDoc <> 0 Then
            lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc, TipoDocDestino, NumIntDocDestino,Origem, NumVou, TipoDoc, Serie, ValorTotal, DataEmissao, Cliente FROM TRPOcorrencias WHERE NumIntDoc= ? ", _
                        0, lNumIntDoc, iTipoDestino, lNumIntDocDestino, iOrigem, lNumVou, sTipoDoc, sSerie, dValorTotal, dtDataEmissao, lCliente, objOcr.lNumIntDoc)
        Else
            lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc, TipoDocDestino, NumIntDocDestino,Origem, NumVou, TipoDoc, Serie, ValorTotal, DataEmissao, Cliente FROM TRPOcorrencias WHERE Codigo= ? ", _
                        0, lNumIntDoc, iTipoDestino, lNumIntDocDestino, iOrigem, lNumVou, sTipoDoc, sSerie, dValorTotal, dtDataEmissao, lCliente, objOcr.lCodigo)
        End If
    Else
        lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc, TipoDocDestino, NumIntDocDestino,Origem, NumVou, TipoDoc, Serie, ValorTotal, DataEmissao, Cliente FROM TRPOcorrencias WHERE Origem = ? AND NumVou = ? AND TipoDoc = ? AND Serie = ? AND Cliente = ? ", _
                    0, lNumIntDoc, iTipoDestino, lNumIntDocDestino, iOrigem, lNumVou, sTipoDoc, sSerie, dValorTotal, dtDataEmissao, lCliente, INATIVACAO_AUTOMATICA_CODIGO, objOcr.lNumVou, objOcr.sTipoDoc, objOcr.sSerie, objOcr.lCliente)
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 189445

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189446

    'Se não existir => ERRO
    If lErro = AD_SQL_SEM_DADOS Then
    
        lNumVou = objOcr.lNumVou
        sTipoDoc = objOcr.sTipoDoc
        sSerie = objOcr.sSerie
        
        bSemDados = True
        
    Else
        bSemDados = False
    End If
    
    If bSemDados And Not bInativacaoExclui Then gError 189447
    
    'Le a tabela TRPVoucher
    lErro = Comando_ExecutarPos(alComando(4), "SELECT Status, ValorOcr FROM TRPVouchers WHERE NumVou = ? AND Serie = ? AND TipVou = ? ", _
                0, iStatusVou, dValorOcr, lNumVou, sSerie, sTipoDoc)
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
            lErro = Comando_ExecutarPos(alComando(5), "UPDATE TRPVouchers SET Status = ?, NumIntDocNVL = ?", alComando(4), STATUS_TRP_VOU_ABERTO, 0)
        End If
    
    ElseIf iOrigem = INATIVACAO_AUTOMATICA_CODIGO Then
        'Se tiver a ocr de inativação cadastrada atualiza os valores no voucher e o status
        lErro = Comando_ExecutarPos(alComando(5), "UPDATE TRPVouchers SET ValorOcr= ? , TemOcr= ? , Status = ?, NumIntDocNVL = ?", alComando(4), dValorOcrAtu, iTemOcr, STATUS_TRP_VOU_ABERTO, 0)
    Else
        'Se for só uma exclusão de ocr acerta os valores no voucher
        lErro = Comando_ExecutarPos(alComando(5), "UPDATE TRPVouchers SET ValorOcr= ? , TemOcr= ? ", alComando(4), dValorOcrAtu, iTemOcr)
    End If
    If lErro <> SUCESSO Then gError 194440

    objVoucher.lNumVou = lNumVou
    objVoucher.sSerie = sSerie
    objVoucher.sTipVou = sTipoDoc

    lErro = CF("TRPVouchers_Le", objVoucher)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 196382

    objOcr.lNumIntDoc = lNumIntDoc
    objOcr.iOrigem = iOrigem
    objOcr.lCliente = lCliente
    
    'Tem que gerar a comissão
    If objOcr.iOrigem = INATIVACAO_AUTOMATICA_CODIGO Then
        lErro = TRPVou_Gera_Comis(objVoucher, -dValorTotal, TRP_TIPO_VALOR_BASE_LIQ, dtDataEmissao, "Exclusão da inativação do Voucher", TRP_TIPO_TRATAMENTO_COMI_NVL_EXCLUSAO, objOcr)
    Else
        lErro = TRPVou_Gera_Comis(objVoucher, -dValorTotal, TRP_TIPO_VALOR_BASE_LIQ, dtDataEmissao, "Exclusão da ocorrência código " & CStr(objOcr.lCodigo), TRP_TIPO_TRATAMENTO_COMI_OCR_EXCLUSAO, objOcr)
    End If
    If lErro <> SUCESSO Then gError 196382
    
    'Se não tinha a OCR cadastrada não faz mais nada
    If Not bSemDados Then
    
        objOcr.lNumIntDoc = lNumIntDoc

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 189448
    
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM TRPOcorrencias", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 189449
        
        lErro = TRPOcorrenciaDet_Exclui(alComando(2), alComando(3), objOcr)
        If lErro <> SUCESSO Then gError 189485
        
        lErro = Item_Exclui_Com_Destino(iTipoDestino, lNumIntDocDestino)
        If lErro <> SUCESSO Then gError 192062
        
        objOCRBD.lCodigo = objOcr.lCodigo
        objOCRBD.lNumIntDoc = objOcr.lNumIntDoc
        
        lErro = CF("TRPOcorrencias_Le", objOCRBD)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 192062
    
    End If
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TRPOcorrencias_Exclui_EmTrans = SUCESSO

    Exit Function

Erro_TRPOcorrencias_Exclui_EmTrans:

    TRPOcorrencias_Exclui_EmTrans = gErr

    Select Case gErr

        Case 189444
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 189445, 189446
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPOCORRENCIAS", gErr)

        Case 189447
            Call Rotina_Erro(vbOKOnly, "ERRO_TRPOCORRENCIAS_NAO_CADASTRADO", gErr, objOcr.lCodigo)

        Case 189448
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRPOCORRENCIAS", gErr)

        Case 189449
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRPOCORRENCIAS", gErr)
            
        Case 189485, 192062, 196381, 196382

        Case 194437, 194438
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPVOUCHERS", gErr)

        Case 194439
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRPVOUCHERS", gErr)

        Case 194440
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRPVOUCHERS", gErr)

        Case 194441
            Call Rotina_Erro(vbOKOnly, "ERRO_TRPVOUCHERS_NAO_CADASTRADO", gErr, lNumVou, sTipoDoc, sSerie)

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
    
            Case TRP_TIPO_DOC_DESTINO_DEBCLI
                       
                Call objContabil.Contabil_Inicializa_Contabilidade4(29, MODULO_CONTASARECEBER)
                
                objDebitoRecCli.lNumIntDoc = lNumIntDocDestino
                
                'Verifica se o debito a receber existe
                lErro = CF("DebitoReceber_Le", objDebitoRecCli)
                If lErro <> SUCESSO And lErro <> 17835 Then gError 192059
                
                lErro = CF("DebitoRecCli_Exclui_EmTransacao", objDebitoRecCli, objContabil)
                If lErro <> SUCESSO Then gError 192059
        
            Case TRP_TIPO_DOC_DESTINO_CREDFORN
            
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

Public Function CamposGenericos_ValidaExclusao_TRPOCROrigem() As Long
'Verifica se é válida a exclusão de valores para o campo Origem

Dim alComando(1 To 2) As Long
Dim lErro As Long
Dim lCodigo As Long
Dim iOrigem As Integer
Dim sValor As String
Dim iIndice As Integer

On Error GoTo Erro_CamposGenericos_ValidaExclusao_TRPOCROrigem
    
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 190154
    Next
    
    'Busca nas ocorrências origens que não existam na tabela CamposGenericos Valores
    lErro = Comando_Executar(alComando(1), "SELECT Codigo, Origem FROM TRPOcorrencias WHERE Origem NOT IN (SELECT CodValor FROM CamposGenericosValores WHERE CodCampo = ?) ", lCodigo, iOrigem, CAMPOSGENERICOS_TRPORIGEM_OCR)
    If lErro <> AD_SQL_SUCESSO Then gError 190155
    
    'Posiciona no primeiro registro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190156
    
    'Se encontrou algum registro => erro de inconsistência
    If lErro = AD_SQL_SUCESSO Then gError 190157
    
    sValor = String(STRING_CAMPOSGENERICOSVALORES_VALOR, 0)
    
    'Busca nas ocorrências origens que não existam na tabela CamposGenericos Valores
    lErro = Comando_Executar(alComando(2), "SELECT Valor FROM CamposGenericosValores WHERE CodCampo = ? AND CodValor = ? ", sValor, CAMPOSGENERICOS_TRPORIGEM_OCR, INATIVACAO_AUTOMATICA_CODIGO)
    If lErro <> AD_SQL_SUCESSO Then gError 190555
    
    'Posiciona no primeiro registro
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190556
    
    'O código 1 é obrigatório pois é utilizado nas inativações automáticas
    If lErro <> AD_SQL_SUCESSO Then gError 190557
    
    If sValor <> INATIVACAO_AUTOMATICA_TEXTO Then gError 190558
     
    'Fecha comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    CamposGenericos_ValidaExclusao_TRPOCROrigem = SUCESSO
    
    Exit Function
    
Erro_CamposGenericos_ValidaExclusao_TRPOCROrigem:

    CamposGenericos_ValidaExclusao_TRPOCROrigem = gErr
    
    Select Case gErr
    
        Case 190154
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 190155, 190156, 190555, 190556, 194430, 194431
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CAMPOSGENERICOS", gErr, CAMPOSGENERICOS_TRPORIGEM_OCR)
        
        Case 190157
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ORIGEM_OCR", gErr, iOrigem, lCodigo)
            
        Case 190557
            Call Rotina_Erro(vbOKOnly, "ERRO_ORIGEMOCR_CODIGO_OBRIGATORIO", gErr, INATIVACAO_AUTOMATICA_CODIGO)
        
        Case 190558
            Call Rotina_Erro(vbOKOnly, "ERRO_ORIGEMOCR_TEXTO_OBRIGATORIO", gErr, INATIVACAO_AUTOMATICA_CODIGO, INATIVACAO_AUTOMATICA_TEXTO)
            
         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190158)
    
    End Select

    'Fecha comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Public Function CamposGenericos_ValidaExclusao_TRPOCRTipo() As Long
'Verifica se é válida a exclusão de valores para o campo Tipo

Dim alComando(1 To 2) As Long
Dim lErro As Long
Dim lCodigo As Long
Dim iTipo As Integer
Dim iTipoAux As Integer
Dim sTextoAux As String
Dim iIndice As Integer
Dim sValor As String

On Error GoTo Erro_CamposGenericos_ValidaExclusao_TRPOCRTipo
    
    'Abre comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 190159
    Next
    
    'Busca nas ocorrências origens que não existam na tabela CamposGenericos Valores
    lErro = Comando_Executar(alComando(1), "SELECT O.Codigo, D.Tipo FROM TRPOcorrenciaDet AS D, TRPOcorrencias AS O WHERE O.NumIntDoc = D.NumIntDocOcr AND D.Tipo NOT IN (SELECT CodValor FROM CamposGenericosValores WHERE CodCampo = ?) ", lCodigo, iTipo, CAMPOSGENERICOS_TRPTIPODET_OCR)
    If lErro <> AD_SQL_SUCESSO Then gError 190160
    
    'Posiciona no primeiro registro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190161
    
    'Se encontrou algum registro => erro de inconsistência
    If lErro = AD_SQL_SUCESSO Then gError 190162
    
    For iIndice = 1 To 5
    
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

        End Select
        
        sValor = String(STRING_CAMPOSGENERICOSVALORES_VALOR, 0)
    
        'Busca nas ocorrências origens que não existam na tabela CamposGenericos Valores
        lErro = Comando_Executar(alComando(2), "SELECT Valor FROM CamposGenericosValores WHERE CodCampo = ? AND CodValor = ? ", sValor, CAMPOSGENERICOS_TRPTIPODET_OCR, iTipoAux)
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
    
    CamposGenericos_ValidaExclusao_TRPOCRTipo = SUCESSO
    
    Exit Function
    
Erro_CamposGenericos_ValidaExclusao_TRPOCRTipo:

    CamposGenericos_ValidaExclusao_TRPOCRTipo = gErr
    
    Select Case gErr
    
        Case 190159
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 190160, 190161, 190559, 190560
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CAMPOSGENERICOS", gErr, CAMPOSGENERICOS_TRPTIPODET_OCR)
        
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

Public Function CamposGenericos_ValidaExclusao_TRPAporteTipo() As Long
'Verifica se é válida a exclusão de valores para o campo Tipo

Dim lComando As Long
Dim lErro As Long
Dim lCodigo As Long
Dim iTipo As Integer

On Error GoTo Erro_CamposGenericos_ValidaExclusao_TRPAporteTipo
    
    'Abre comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 190550
    
    'Busca nas ocorrências origens que não existam na tabela CamposGenericos Valores
    lErro = Comando_Executar(lComando, "SELECT Codigo, Tipo FROM TRPAportes WHERE Tipo NOT IN (SELECT CodValor FROM CamposGenericosValores WHERE CodCampo = ?) ", lCodigo, iTipo, CAMPOSGENERICOS_TRPTIPOAPORTE_OCR)
    If lErro <> AD_SQL_SUCESSO Then gError 190551
    
    'Posiciona no primeiro registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190552
    
    'Se encontrou algum registro => erro de inconsistência
    If lErro = AD_SQL_SUCESSO Then gError 190553
    
    'Fecha comando
    Call Comando_Fechar(lComando)
    
    CamposGenericos_ValidaExclusao_TRPAporteTipo = SUCESSO
    
    Exit Function
    
Erro_CamposGenericos_ValidaExclusao_TRPAporteTipo:

    CamposGenericos_ValidaExclusao_TRPAporteTipo = gErr
    
    Select Case gErr
    
        Case 190550
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 190551, 190552
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CAMPOSGENERICOS", gErr, CAMPOSGENERICOS_TRPTIPODET_OCR)
        
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
Public Function TRPOcorrencias_Libera(ByVal colOcorrencias As Collection) As Long

Dim lErro As Long
Dim lTransacao As Long
Dim objOcr As ClassTRPOcorrencias
Dim alComando(0 To 4) As Long
Dim iIndice As Integer

On Error GoTo Erro_TRPOcorrencias_Libera

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 190306
    
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 190312
    Next
    
    For Each objOcr In colOcorrencias
    
        lErro = TRPOcorrencias_Grava_EmTrans(objOcr)
        If lErro <> SUCESSO Then gError 190307

        lErro = TRPOcorrencias_Libera_EmTrans(alComando, objOcr)
        If lErro <> SUCESSO Then gError 190307
        
    Next
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 190308

    TRPOcorrencias_Libera = SUCESSO

    Exit Function

Erro_TRPOcorrencias_Libera:

    TRPOcorrencias_Libera = gErr

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

Public Function TRPOcorrencias_Libera_EmTrans(alComando() As Long, ByVal objOcr As ClassTRPOcorrencias) As Long

Dim lErro As Long
Dim lNumIntDoc As Long
Dim objDebitoRecCli As New ClassDebitoRecCli
Dim colInfoComissao As New colInfoComissao
Dim objContabil As ClassContabil
Dim objCreditoPagar As New ClassCreditoPagar

On Error GoTo Erro_TRPOcorrencias_Libera_EmTrans

    objOcr.iStatus = STATUS_TRP_OCR_LIBERADO

    If objOcr.iFormaPagto = FORMAPAGTO_TRP_OCR_FAT Then
        'Não precisa fazer nada => O Tratamento é dado na tela de Faturamento
    Else
        'Gera o crédito e informa o número do numintdoc do débito
    
        If objOcr.dValorTotal < 0 Then
        
            objDebitoRecCli.iFilialEmpresa = giFilialEmpresa
            objDebitoRecCli.lCliente = objOcr.lCliente
            objDebitoRecCli.iFilial = objOcr.iFilialCliente
            objDebitoRecCli.sSiglaDocumento = "NCR" 'Nota de crédito
            objDebitoRecCli.lNumTitulo = objOcr.lCodigo
            objDebitoRecCli.dtDataEmissao = objOcr.dtDataEmissao
            objDebitoRecCli.dValorTotal = Abs(objOcr.dValorTotal)
            objDebitoRecCli.dValorProdutos = Abs(objOcr.dValorTotal)
            
            lErro = CF("DebitoRecCli_Grava_EmTrans", objDebitoRecCli, colInfoComissao, objContabil)
            If lErro <> SUCESSO Then gError 190315
            
            objOcr.lNumIntDocDestino = objDebitoRecCli.lNumIntDoc
            objOcr.iTipoDocDestino = TRP_TIPO_DOC_DESTINO_DEBCLI
            objOcr.lNumDocDestino = objDebitoRecCli.lNumTitulo
        
        Else
        
            objCreditoPagar.iFilialEmpresa = giFilialEmpresa
            objCreditoPagar.lFornecedor = objOcr.lCliente
            objCreditoPagar.iFilial = objOcr.iFilialCliente
            objCreditoPagar.sSiglaDocumento = "NCP" 'Nota de crédito
            objCreditoPagar.lNumTitulo = objOcr.lCodigo
            objCreditoPagar.dtDataEmissao = objOcr.dtDataEmissao
            objCreditoPagar.dValorTotal = Abs(objOcr.dValorTotal)
            objCreditoPagar.dValorProdutos = Abs(objOcr.dValorTotal)
            
            lErro = CF("CreditoPagar_Grava_EmTrans", objCreditoPagar, objContabil)
            If lErro <> SUCESSO Then gError 190325
            
            objOcr.lNumIntDocDestino = objCreditoPagar.lNumIntDoc
            objOcr.iTipoDocDestino = TRP_TIPO_DOC_DESTINO_CREDFORN
            objOcr.lNumDocDestino = objCreditoPagar.lNumTitulo
    
        End If
    
    End If
    
    'Le a tabelaTRPOcorrencias
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM TRPOcorrencias WHERE Codigo= ? ", _
                0, lNumIntDoc, objOcr.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 190316

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190317

    'Se existir => UPDATE, senão => INSERT
    If lErro <> AD_SQL_SUCESSO Then gError 190318

    'Trava registro contra alterações/Leituras
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> SUCESSO Then gError 190319
    
    objOcr.lNumIntDoc = lNumIntDoc

    'Altera o Status da ocorrencia e a forma de pagamento
    lErro = Comando_ExecutarPos(alComando(1), "UPDATE TRPOcorrencias SET Status= ? , FormaPagto= ?, NumIntDocDestino = ?, DataLiberacao = ? , Usuario = ?, TipoDocDestino = ?, ValorTotal = ?, NumTitulo = ? ", alComando(0), objOcr.iStatus, objOcr.iFormaPagto, objOcr.lNumIntDocDestino, Date, gsUsuario, objOcr.iTipoDocDestino, objOcr.dValorTotal, objOcr.lNumDocDestino)
    If lErro <> AD_SQL_SUCESSO Then gError 190320
    
    lErro = TRPOcorrenciaDet_Exclui(alComando(2), alComando(3), objOcr)
    If lErro <> SUCESSO Then gError 190321
    
    lErro = TRPOcorrenciaDet_Insere(alComando(4), objOcr)
    If lErro <> SUCESSO Then gError 190322

    TRPOcorrencias_Libera_EmTrans = SUCESSO

    Exit Function

Erro_TRPOcorrencias_Libera_EmTrans:

    TRPOcorrencias_Libera_EmTrans = gErr

    Select Case gErr
        
        Case 190316, 190317
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPOCORRENCIAS", gErr)
            
        Case 190318
            Call Rotina_Erro(vbOKOnly, "ERRO_TRPOCORRENCIAS_NAO_CADASTRADO", gErr, objOcr.lCodigo)

        Case 190319
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRPOCORRENCIAS", gErr)

        Case 190320
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRPOCORRENCIAS", gErr)
            
        Case 190321, 190322, 190325
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190323)

    End Select

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'GRAVAÇÃO
Public Function TRPAportes_Grava(ByVal objTRPAportes As ClassTRPAportes) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_TRPAportes_Grava

    lErro = TRPAportes_Valida(objTRPAportes)
    If lErro <> SUCESSO Then gError 196478

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 190348

    lErro = CF("TRPAportes_Grava_EmTrans", objTRPAportes)
    If lErro <> SUCESSO Then gError 190349

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 190350

    TRPAportes_Grava = SUCESSO

    Exit Function

Erro_TRPAportes_Grava:

    TRPAportes_Grava = gErr

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

Public Function TRPAportes_Grava_EmTrans(ByVal objTRPAportes As ClassTRPAportes) As Long

Dim lErro As Long
Dim alComando(0 To 20) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim iProxCmd As Integer
Dim iMoeda As Integer

On Error GoTo Erro_TRPAportes_Grava_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 190352
    Next

    'Le a tabelaTRPAportes
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc, Moeda FROM TRPAportes WHERE Codigo= ? ", _
                0, lNumIntDoc, iMoeda, objTRPAportes.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 190353

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190354

    'Se existir => UPDATE, senão => INSERT
    If lErro = AD_SQL_SUCESSO Then

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 190355
        
        objTRPAportes.lNumIntDoc = lNumIntDoc
        
        If iMoeda <> objTRPAportes.iMoeda Then gError 192887

        lErro = Comando_ExecutarPos(alComando(1), "UPDATE TRPAportes SET Cliente= ? , FilialCliente= ? , DataEmissao= ? , Observacao= ? , " & _
                "Historico= ? , Tipo= ? , Moeda= ? , PrevValor= ? , PrevDataDe= ? , " & _
                "PrevDataAte= ? ", alComando(0), _
                objTRPAportes.lCliente, objTRPAportes.iFilialCliente, objTRPAportes.dtDataEmissao, objTRPAportes.sObservacao, _
                objTRPAportes.sHistorico, objTRPAportes.iTipo, objTRPAportes.iMoeda, objTRPAportes.dPrevValor, objTRPAportes.dtPrevDataDe, _
                objTRPAportes.dtPrevDataAte)
        If lErro <> AD_SQL_SUCESSO Then gError 190356

    Else

        'Obter NumIntDoc
        lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_INT_PROX_TRPAPORTES", lNumIntDoc)
        If lErro <> SUCESSO Then gError 190357

        objTRPAportes.lNumIntDoc = lNumIntDoc

        lErro = Comando_Executar(alComando(1), "INSERT INTO TRPAportes( NumIntDoc, Codigo, Cliente, FilialCliente, " & _
                "DataEmissao, Observacao, Historico, Tipo, Moeda, " & _
                "PrevValor, PrevDataDe, PrevDataAte) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)", _
                objTRPAportes.lNumIntDoc, objTRPAportes.lCodigo, objTRPAportes.lCliente, objTRPAportes.iFilialCliente, _
                objTRPAportes.dtDataEmissao, objTRPAportes.sObservacao, objTRPAportes.sHistorico, objTRPAportes.iTipo, objTRPAportes.iMoeda, _
                objTRPAportes.dPrevValor, objTRPAportes.dtPrevDataDe, objTRPAportes.dtPrevDataAte)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190358

    End If
    
    iProxCmd = 1
    
    lErro = TRPAportes_Grava_Pagtos(alComando, objTRPAportes, iProxCmd)
    If lErro <> SUCESSO Then gError 190359

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TRPAportes_Grava_EmTrans = SUCESSO

    Exit Function

Erro_TRPAportes_Grava_EmTrans:

    TRPAportes_Grava_EmTrans = gErr

    Select Case gErr

        Case 190352
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 190353, 190354
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPAPORTES", gErr)

        Case 190355
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRPAPORTES", gErr)

        Case 190356
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRPAPORTES", gErr)

        Case 190357, 190359
        
        Case 190358
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRPAPORTES", gErr)
            
        Case 192887
            Call Rotina_Erro(vbOKOnly, "ERRO_TRPAPORTES_TROCA_MOEDA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190360)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function TRPAportes_Grava_Pagtos(alComando() As Long, ByVal objAporte As ClassTRPAportes, iProxCmd As Integer) As Long

Dim lErro As Long
Dim objAporteBD As New ClassTRPAportes
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

On Error GoTo Erro_TRPAportes_Grava_Pagtos

    'A tela tem que manter os numintdoc nos pagtos
    
    objAporteBD.lCodigo = objAporte.lCodigo
    
    'Le o aporte do BD
    lErro = CF("TRPAportes_Le_Aux", alComando, objAporteBD, iProxCmd, True, True)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 190365
    
    '1 para cada tipo de pagamento
    'Para evitar código duplicado
    For iIndice = 1 To 4
    
        iSeq = 0
    
        If iIndice = 1 Then
            Set colColecaoBD = objAporteBD.colPagtoCondicionados
            Set colColecao = objAporte.colPagtoCondicionados
            sPrefixoFunc = "TRPAportePagtoCond"
        ElseIf iIndice = 2 Then
            Set colColecaoBD = objAporteBD.colPagtoDireto
            Set colColecao = objAporte.colPagtoDireto
            sPrefixoFunc = "TRPAportePagtoDireto"
        ElseIf iIndice = 3 Then
            Set colColecaoBD = objAporteBD.colPagtoSobreFaturaCond
            Set colColecao = objAporte.colPagtoSobreFaturaCond
            sPrefixoFunc = "TRPAportePagtoFatCond"
        Else
            Set colColecaoBD = objAporteBD.colPagtoSobreFatura
            Set colColecao = objAporte.colPagtoSobreFatura
            sPrefixoFunc = "TRPAportePagtoFat"
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
                    lErro = CF("Verifica_Existencia_Doc_TRP", objObjetoBD.lNumIntDocDestino, objObjetoBD.iTipoDocDestino, bExisteDestino, lNumTitulo, sDoc)
                    If lErro <> SUCESSO Then gError 190368
                Else
                    bExisteDestino = False
                End If
                
                If Not bExisteDestino Then
                    iProxCmd = iProxCmd + 1
                    lErro = CF(sPrefixoFunc & "_Atualiza", alComando(iProxCmd - 1), alComando(iProxCmd), objObjeto, objAporte)
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
                    lErro = CF("Verifica_Existencia_Doc_TRP", objObjetoBD.lNumIntDocDestino, objObjetoBD.iTipoDocDestino, bExisteDestino, lNumTitulo, sDoc)
                    If lErro <> SUCESSO Then gError 190370
                
                    If bExisteDestino And objObjetoBD.iTipoDocDestino <> TRP_TIPO_DOC_DESTINO_DEBCLI And objObjetoBD.iTipoDocDestino <> TRP_TIPO_DOC_DESTINO_CREDFORN Then gError 190371
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
    
    TRPAportes_Grava_Pagtos = SUCESSO

    Exit Function

Erro_TRPAportes_Grava_Pagtos:

    TRPAportes_Grava_Pagtos = gErr

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

Public Function Verifica_Existencia_Doc_TRP(ByVal lNumIntDoc As Long, ByVal iTipo As Integer, bExiste As Boolean, lNumTitulo As Long, sDoc As String) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim sTabela As String
Dim sSQL As String
Dim sDocAux As String

On Error GoTo Erro_Verifica_Existencia_Doc_TRP

    bExiste = False
    sDoc = ""

    If iTipo <> 0 And lNumIntDoc <> 0 Then

        'Abertura de Comando
        For iIndice = LBound(alComando) To UBound(alComando)
            alComando(iIndice) = Comando_Abrir()
            If alComando(iIndice) = 0 Then gError 190361
        Next
        
        Select Case iTipo
        
            Case TRP_TIPO_DOC_DESTINO_CREDFORN
                sTabela = TRP_TIPO_DOC_DESTINO_CREDFORN_TABELA
                sDocAux = TRP_TIPO_DOC_DESTINO_CREDFORN_TEXTO
                
            Case TRP_TIPO_DOC_DESTINO_DEBCLI
                sTabela = TRP_TIPO_DOC_DESTINO_DEBCLI_TABELA
                sDocAux = TRP_TIPO_DOC_DESTINO_DEBCLI_TEXTO
        
            Case TRP_TIPO_DOC_DESTINO_TITPAG
                sTabela = TRP_TIPO_DOC_DESTINO_TITPAG_TABELA
                sDocAux = TRP_TIPO_DOC_DESTINO_TITPAG_TEXTO
        
            Case TRP_TIPO_DOC_DESTINO_TITREC
                sTabela = TRP_TIPO_DOC_DESTINO_TITREC_TABELA
                sDocAux = TRP_TIPO_DOC_DESTINO_TITREC_TEXTO
        
            Case TRP_TIPO_DOC_DESTINO_NFSPAG
                sTabela = TRP_TIPO_DOC_DESTINO_NFSPAG_TABELA
                sDocAux = TRP_TIPO_DOC_DESTINO_NFSPAG_TEXTO
        
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

    Verifica_Existencia_Doc_TRP = SUCESSO

    Exit Function

Erro_Verifica_Existencia_Doc_TRP:

    Verifica_Existencia_Doc_TRP = gErr

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

Public Function TRPAportePagtoCond_Atualiza(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal objTRPAportePagtoCond As ClassTRPAportePagtoCond, ByVal objAporte As ClassTRPAportes) As Long

Dim lErro As Long
Dim tPagto As typeTRPAportePagtoCond
Dim lNumIntDoc As Long
Dim iStatus As Integer

On Error GoTo Erro_TRPAportePagtoCond_Atualiza

    'Le a tabelaTRPAportePagtoCond
    lErro = Comando_ExecutarPos(lComando1, "SELECT NumIntDoc, Status FROM TRPAportePagtoCond WHERE NumIntDoc = ? ", 0, lNumIntDoc, iStatus, objTRPAportePagtoCond.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 190385

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190386

    If lErro <> AD_SQL_SUCESSO Then gError 190387
    
    'If iStatus <> STATUS_TRP_OCR_BLOQUEADO Then gError 190563
    
    If iStatus = STATUS_TRP_OCR_BLOQUEADO Then
    
        lErro = Comando_ExecutarPos(lComando2, "UPDATE TRPAportePagtoCond SET NumIntDocAporte= ? , Base= ? , DataPagto= ? , " & _
                "NumIntDocDestino= ? , FormaPagto= ? , TipoDocDestino= ? , Percentual= ?, Seq = ? ", lComando1, _
                objTRPAportePagtoCond.lNumIntDocAporte, objTRPAportePagtoCond.iBase, objTRPAportePagtoCond.dtDataPagto, _
                objTRPAportePagtoCond.lNumIntDocDestino, objTRPAportePagtoCond.iFormaPagto, objTRPAportePagtoCond.iTipoDocDestino, objTRPAportePagtoCond.dPercentual, objTRPAportePagtoCond.iSeq)
        
    Else
        'Se já estiver liberado só deixa apontar para o documento de destino
        lErro = Comando_ExecutarPos(lComando2, "UPDATE TRPAportePagtoCond SET NumIntDocDestino= ? , TipoDocDestino= ? ", lComando1, _
                objTRPAportePagtoCond.lNumIntDocDestino, objTRPAportePagtoCond.iTipoDocDestino)
    
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 190388

    TRPAportePagtoCond_Atualiza = SUCESSO

    Exit Function

Erro_TRPAportePagtoCond_Atualiza:

    TRPAportePagtoCond_Atualiza = gErr

    Select Case gErr

        Case 190385, 190386
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPAPORTEPAGTOCOND", gErr)
            
        Case 190387
            Call Rotina_Erro(vbOKOnly, "ERRO_TRPAPORTEPAGTOCOND_NAO_CADASTRADO", gErr, objTRPAportePagtoCond.lNumIntDoc)

        Case 190388
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRPAPORTEPAGTOCOND", gErr)
            
        Case 190563
            Call Rotina_Erro(vbOKOnly, "ERRO_TRPAPORTEPAGTOCOND_STATUS_DIF_BLOQUEADO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190389)

    End Select


    Exit Function

End Function

Public Function TRPAportePagtoCond_Insere(ByVal lComando As Long, ByVal objTRPAportePagtoCond As ClassTRPAportePagtoCond, ByVal objAporte As ClassTRPAportes) As Long

Dim lErro As Long
Dim lNumIntDoc As Long

On Error GoTo Erro_TRPAportePagtoCond_Insere

    'Obter NumIntDoc
    lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_INT_PROX_TRPAPORTEPAGTOCOND", lNumIntDoc)
    If lErro <> SUCESSO Then gError 190382

    objTRPAportePagtoCond.lNumIntDoc = lNumIntDoc
        
    lErro = Comando_Executar(lComando, "INSERT INTO TRPAportePagtoCond( NumIntDoc, NumIntDocAporte, Base, Status, " & _
            "DataPagto, NumIntDocDestino, FormaPagto, TipoDocDestino, Percentual, Valor, Seq) VALUES (?,?,?,?,?,?,?,?,?,?,?)", _
            objTRPAportePagtoCond.lNumIntDoc, objTRPAportePagtoCond.lNumIntDocAporte, objTRPAportePagtoCond.iBase, STATUS_TRP_OCR_BLOQUEADO, _
            objTRPAportePagtoCond.dtDataPagto, objTRPAportePagtoCond.lNumIntDocDestino, objTRPAportePagtoCond.iFormaPagto, objTRPAportePagtoCond.iTipoDocDestino, objTRPAportePagtoCond.dPercentual, objTRPAportePagtoCond.dValor, objTRPAportePagtoCond.iSeq)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190383


    TRPAportePagtoCond_Insere = SUCESSO

    Exit Function

Erro_TRPAportePagtoCond_Insere:

    TRPAportePagtoCond_Insere = gErr

    Select Case gErr
    
        Case 190382

        Case 190383
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRPAPORTEPAGTOCOND", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190384)

    End Select

    Exit Function

End Function

Public Function TRPAportePagtoCond_Exclui(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal objTRPAportePagtoCond As ClassTRPAportePagtoCond) As Long

Dim lErro As Long
Dim lNumIntDoc As Long
Dim iTipoDestino As Integer
Dim lNumIntDocDestino As Long

On Error GoTo Erro_TRPAportePagtoCond_Exclui

    'Le a tabelaTRPAportePagtoCond
    lErro = Comando_ExecutarPos(lComando1, "SELECT NumIntDoc, TipoDocDestino, NumIntDocDestino FROM TRPAportePagtoCond WHERE NumIntDoc = ? ", 0, lNumIntDoc, iTipoDestino, lNumIntDocDestino, objTRPAportePagtoCond.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 190390

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190391

    'Se não existir => ERRO
    If lErro <> AD_SQL_SUCESSO Then gError 190392
    
    lErro = Comando_ExecutarPos(lComando2, "DELETE FROM TRPAportePagtoCond ", lComando1)
    If lErro <> AD_SQL_SUCESSO Then gError 190393
    
    lErro = Item_Exclui_Com_Destino(iTipoDestino, lNumIntDocDestino)
    If lErro <> SUCESSO Then gError 192063

    TRPAportePagtoCond_Exclui = SUCESSO

    Exit Function

Erro_TRPAportePagtoCond_Exclui:

    TRPAportePagtoCond_Exclui = gErr

    Select Case gErr

        Case 190390, 190391
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPAPORTEPAGTOCOND", gErr)
            
        Case 190392
            Call Rotina_Erro(vbOKOnly, "ERRO_TRPAPORTEPAGTOCOND_NAO_CADASTRADO", gErr, objTRPAportePagtoCond.lNumIntDoc)

        Case 190393
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRPAPORTEPAGTOCOND", gErr)
            
        Case 192063

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190394)

    End Select


    Exit Function

End Function

Public Function TRPAportePagtoDireto_Exclui(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal objTRPAportePagtoDireto As ClassTRPAportePagtoDireto) As Long

Dim lErro As Long
Dim lNumIntDoc As Long
Dim iTipoDestino As Integer
Dim lNumIntDocDestino As Long

On Error GoTo Erro_TRPAportePagtoDireto_Exclui

    'Le a tabelaTRPAportePagtoDireto
    lErro = Comando_ExecutarPos(lComando1, "SELECT NumIntDoc, TipoDocDestino, NumIntDocDestino FROM TRPAportePagtoDireto WHERE NumIntDoc= ? ", _
                0, lNumIntDoc, iTipoDestino, lNumIntDocDestino, objTRPAportePagtoDireto.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 190403

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190404

    'Se não existir => ERRO
    If lErro <> AD_SQL_SUCESSO Then gError 190405

    lErro = Comando_ExecutarPos(lComando2, "DELETE FROM TRPAportePagtoDireto  ", lComando1)
    If lErro <> AD_SQL_SUCESSO Then gError 190406
    
    lErro = Item_Exclui_Com_Destino(iTipoDestino, lNumIntDocDestino)
    If lErro <> SUCESSO Then gError 192065

    TRPAportePagtoDireto_Exclui = SUCESSO

    Exit Function

Erro_TRPAportePagtoDireto_Exclui:

    TRPAportePagtoDireto_Exclui = gErr

    Select Case gErr

        Case 190403, 190404
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPAPORTEPAGTODIRETO", gErr)

        Case 190405
            Call Rotina_Erro(vbOKOnly, "ERRO_TRPAPORTEPAGTODIRETO_NAO_CADASTRADO", gErr, objTRPAportePagtoDireto.lNumIntDoc)
        
        Case 190406
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRPAPORTEPAGTODIRETO", gErr)
            
        Case 192065

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190407)

    End Select

    Exit Function

End Function

Public Function TRPAportePagtoDireto_Atualiza(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal objTRPAportePagtoDireto As ClassTRPAportePagtoDireto, ByVal objAporte As ClassTRPAportes) As Long

Dim lErro As Long
Dim lNumIntDoc As Long
Dim lNumIntDocDestino As Long
Dim objDebitoRecCli As New ClassDebitoRecCli
Dim colInfoComissao As New colInfoComissao
Dim objContabil As ClassContabil
Dim lCodigoAuto As Long
Dim sFiltro As String

On Error GoTo Erro_TRPAportePagtoDireto_Atualiza

    'Le a tabelaTRPAportePagtoDireto
    lErro = Comando_ExecutarPos(lComando1, "SELECT NumIntDoc,NumIntDocDestino FROM TRPAportePagtoDireto WHERE NumIntDoc= ? ", _
                0, lNumIntDoc, lNumIntDocDestino, objTRPAportePagtoDireto.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 190398

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190399

    'Se existir => UPDATE, senão => INSERT
    If lErro <> AD_SQL_SUCESSO Then gError 190400

    If objTRPAportePagtoDireto.iFormaPagto = FORMAPAGTO_TRP_OCR_CRED And lNumIntDocDestino = 0 Then
        'Gera o crédito e informa o número do numintdoc
          
        lErro = CF("Config_ObterAutomatico2", TRP_TIPO_DOC_DESTINO_DEBCLI_TABELA, "NumTitulo", lCodigoAuto, sFiltro)
        If lErro <> SUCESSO Then gError 190751
           
        objDebitoRecCli.iFilialEmpresa = giFilialEmpresa
        objDebitoRecCli.lCliente = objAporte.lCliente
        objDebitoRecCli.iFilial = objAporte.iFilialCliente
        objDebitoRecCli.sSiglaDocumento = "NCR" 'Nota de crédito
        objDebitoRecCli.lNumTitulo = lCodigoAuto
        objDebitoRecCli.dtDataEmissao = gdtDataAtual
        objDebitoRecCli.dValorTotal = objTRPAportePagtoDireto.dValor
        objDebitoRecCli.dValorProdutos = objTRPAportePagtoDireto.dValor
        
        lErro = CF("DebitoRecCli_Grava_EmTrans", objDebitoRecCli, colInfoComissao, objContabil)
        If lErro <> SUCESSO Then gError 190752
        
        objTRPAportePagtoDireto.lNumIntDocDestino = objDebitoRecCli.lNumIntDoc
        objTRPAportePagtoDireto.iTipoDocDestino = TRP_TIPO_DOC_DESTINO_DEBCLI
    
    End If
    
    lErro = Comando_ExecutarPos(lComando2, "UPDATE TRPAportePagtoDireto SET Valor= ? , Vencimento= ? , NumIntDocDestino= ? , FormaPagto= ? , " & _
            "TipoDocDestino= ?, Seq = ?  ", lComando1, _
            objTRPAportePagtoDireto.dValor, objTRPAportePagtoDireto.dtVencimento, objTRPAportePagtoDireto.lNumIntDocDestino, objTRPAportePagtoDireto.iFormaPagto, _
            objTRPAportePagtoDireto.iTipoDocDestino, objTRPAportePagtoDireto.iSeq)
    If lErro <> AD_SQL_SUCESSO Then gError 190401

    TRPAportePagtoDireto_Atualiza = SUCESSO

    Exit Function

Erro_TRPAportePagtoDireto_Atualiza:

    TRPAportePagtoDireto_Atualiza = gErr

    Select Case gErr

        Case 190398, 190399
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPAPORTEPAGTODIRETO", gErr)

        Case 190400
            Call Rotina_Erro(vbOKOnly, "ERRO_TRPAPORTEPAGTODIRETO_NAO_CADASTRADO", gErr, objTRPAportePagtoDireto.lNumIntDoc)
        
        Case 190401
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRPAPORTEPAGTODIRETO", gErr)
            
        Case 190751, 190752

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190402)

    End Select

    Exit Function

End Function

Public Function TRPAportePagtoDireto_Insere(ByVal lComando As Long, ByVal objTRPAportePagtoDireto As ClassTRPAportePagtoDireto, ByVal objAporte As ClassTRPAportes) As Long

Dim lErro As Long
Dim lNumIntDoc As Long
Dim objDebitoRecCli As New ClassDebitoRecCli
Dim colInfoComissao As New colInfoComissao
Dim objContabil As ClassContabil
Dim lCodigoAuto As Long
Dim sFiltro As String

On Error GoTo Erro_TRPAportePagtoDireto_Insere

    'Obter NumIntDoc
    lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_INT_PROX_TRPAPORTEPAGTODIRETO", lNumIntDoc)
    If lErro <> SUCESSO Then gError 190395

    objTRPAportePagtoDireto.lNumIntDoc = lNumIntDoc
    
    If objTRPAportePagtoDireto.iFormaPagto = FORMAPAGTO_TRP_OCR_FAT Then
        'Não precisa fazer nada => O Tratamento é dado quando chegar a nota
    Else
        'Gera o crédito e informa o número do numintdoc
          
        lErro = CF("Config_ObterAutomatico2", TRP_TIPO_DOC_DESTINO_DEBCLI_TABELA, "NumTitulo", lCodigoAuto, sFiltro)
        If lErro <> SUCESSO Then gError 190751
           
        objDebitoRecCli.iFilialEmpresa = giFilialEmpresa
        objDebitoRecCli.lCliente = objAporte.lCliente
        objDebitoRecCli.iFilial = objAporte.iFilialCliente
        objDebitoRecCli.sSiglaDocumento = "NCR" 'Nota de crédito
        objDebitoRecCli.lNumTitulo = lCodigoAuto
        objDebitoRecCli.dtDataEmissao = gdtDataAtual
        objDebitoRecCli.dValorTotal = objTRPAportePagtoDireto.dValor
        objDebitoRecCli.dValorProdutos = objTRPAportePagtoDireto.dValor
        
        lErro = CF("DebitoRecCli_Grava_EmTrans", objDebitoRecCli, colInfoComissao, objContabil)
        If lErro <> SUCESSO Then gError 190752
        
        objTRPAportePagtoDireto.lNumIntDocDestino = objDebitoRecCli.lNumIntDoc
        objTRPAportePagtoDireto.iTipoDocDestino = TRP_TIPO_DOC_DESTINO_DEBCLI
    
    End If

    lErro = Comando_Executar(lComando, "INSERT INTO TRPAportePagtoDireto( NumIntDoc, NumIntDocAporte, Valor, Vencimento, " & _
            "NumIntDocDestino, FormaPagto, TipoDocDestino, Seq) VALUES (?,?,?,?,?,?,?,?)", _
            objTRPAportePagtoDireto.lNumIntDoc, objTRPAportePagtoDireto.lNumIntDocAporte, objTRPAportePagtoDireto.dValor, objTRPAportePagtoDireto.dtVencimento, _
            objTRPAportePagtoDireto.lNumIntDocDestino, objTRPAportePagtoDireto.iFormaPagto, objTRPAportePagtoDireto.iTipoDocDestino, objTRPAportePagtoDireto.iSeq)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190396

    TRPAportePagtoDireto_Insere = SUCESSO

    Exit Function

Erro_TRPAportePagtoDireto_Insere:

    TRPAportePagtoDireto_Insere = gErr

    Select Case gErr

        Case 190395, 190751, 190752

        Case 190396
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRPAPORTEPAGTODIRETO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190397)

    End Select


    Exit Function

End Function

Public Function TRPAportePagtoFat_Exclui(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal objTRPAportePagtoFat As ClassTRPAportePagtoFat) As Long

Dim lErro As Long
Dim lNumIntDoc As Long

On Error GoTo Erro_TRPAportePagtoFat_Exclui

    'Le a tabelaTRPAportePagtoFat
    lErro = Comando_ExecutarPos(lComando1, "SELECT NumIntDoc FROM TRPAportePagtoFat WHERE NumIntDoc= ? ", _
                0, lNumIntDoc, objTRPAportePagtoFat.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 190416

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190417

    'Se não existir => ERRO
    If lErro <> AD_SQL_SUCESSO Then gError 190418

    lErro = Comando_ExecutarPos(lComando2, "DELETE FROM TRPAportePagtoFat ", lComando1)
    If lErro <> AD_SQL_SUCESSO Then gError 190419
    
    TRPAportePagtoFat_Exclui = SUCESSO

    Exit Function

Erro_TRPAportePagtoFat_Exclui:

    TRPAportePagtoFat_Exclui = gErr

    Select Case gErr

        Case 190416, 190417
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPAPORTEPAGTOFAT", gErr)

        Case 190418
            Call Rotina_Erro(vbOKOnly, "ERRO_TRPAPORTEPAGTOFAT_NAO_CADASTRADO", gErr, objTRPAportePagtoFat.lNumIntDoc)
        
        Case 190419
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRPAPORTEPAGTOFAT", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190420)

    End Select

    Exit Function

End Function

Public Function TRPAportePagtoFat_Atualiza(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal objTRPAportePagtoFat As ClassTRPAportePagtoFat, ByVal objAporte As ClassTRPAportes) As Long

Dim lErro As Long
Dim lNumIntDoc As Long

On Error GoTo Erro_TRPAportePagtoFat_Atualiza

    'Le a tabelaTRPAportePagtoFat
    lErro = Comando_ExecutarPos(lComando1, "SELECT NumIntDoc FROM TRPAportePagtoFat WHERE NumIntDoc= ? ", _
                0, lNumIntDoc, objTRPAportePagtoFat.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 190411

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190412

    'Se não existir => ERRO
    If lErro <> AD_SQL_SUCESSO Then gError 190413

    lErro = Comando_ExecutarPos(lComando2, "UPDATE TRPAportePagtoFat SET Valor= ? , ValidadeDe= ? , ValidadeAte= ? , Saldo= ? , " & _
            "Percentual= ?, Seq = ? ", lComando1, _
            objTRPAportePagtoFat.dValor, objTRPAportePagtoFat.dtValidadeDe, objTRPAportePagtoFat.dtValidadeAte, objTRPAportePagtoFat.dSaldo, _
            objTRPAportePagtoFat.dPercentual, objTRPAportePagtoFat.iSeq)
    If lErro <> AD_SQL_SUCESSO Then gError 190414
    
    TRPAportePagtoFat_Atualiza = SUCESSO

    Exit Function

Erro_TRPAportePagtoFat_Atualiza:

    TRPAportePagtoFat_Atualiza = gErr

    Select Case gErr

        Case 190411, 190412
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPAPORTEPAGTOFAT", gErr)

        Case 190413
            Call Rotina_Erro(vbOKOnly, "ERRO_TRPAPORTEPAGTOFAT_NAO_CADASTRADO", gErr, objTRPAportePagtoFat.lNumIntDoc)
        
        Case 190414
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRPAPORTEPAGTOFAT", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190415)

    End Select

    Exit Function

End Function

Public Function TRPAportePagtoFat_Insere(ByVal lComando As Long, ByVal objTRPAportePagtoFat As ClassTRPAportePagtoFat, ByVal objAporte As ClassTRPAportes) As Long

Dim lErro As Long
Dim lNumIntDoc As Long

On Error GoTo Erro_TRPAportePagtoFat_Insere

    'Obter NumIntDoc
    lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_INT_PROX_TRPAPORTEPAGTOFAT", lNumIntDoc)
    If lErro <> SUCESSO Then gError 190408

    objTRPAportePagtoFat.lNumIntDoc = lNumIntDoc

    lErro = Comando_Executar(lComando, "INSERT INTO TRPAportePagtoFat( NumIntDoc, NumIntDocAporte, Valor, ValidadeDe, " & _
            "ValidadeAte, Saldo, Percentual, Seq) VALUES (?,?,?,?,?,?,?,?)", _
            objTRPAportePagtoFat.lNumIntDoc, objTRPAportePagtoFat.lNumIntDocAporte, objTRPAportePagtoFat.dValor, objTRPAportePagtoFat.dtValidadeDe, _
            objTRPAportePagtoFat.dtValidadeAte, objTRPAportePagtoFat.dSaldo, objTRPAportePagtoFat.dPercentual, objTRPAportePagtoFat.iSeq)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190409

    TRPAportePagtoFat_Insere = SUCESSO

    Exit Function

Erro_TRPAportePagtoFat_Insere:

    TRPAportePagtoFat_Insere = gErr

    Select Case gErr

        Case 190408

        Case 190409
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRPAPORTEPAGTOFAT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190410)

    End Select

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'EXCLUSÃO
Public Function TRPAportes_Exclui(ByVal objTRPAportes As ClassTRPAportes) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_TRPAportes_Exclui

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 190421

    lErro = CF("TRPAportes_Exclui_EmTrans", objTRPAportes)
    If lErro <> SUCESSO Then gError 190422

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 190423

    TRPAportes_Exclui = SUCESSO

    Exit Function

Erro_TRPAportes_Exclui:

    TRPAportes_Exclui = gErr

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

Public Function TRPAportes_Exclui_EmTrans(ByVal objTRPAportes As ClassTRPAportes) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long

On Error GoTo Erro_TRPAportes_Exclui_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 190425
    Next

    'Le a tabelaTRPAportes
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM TRPAportes WHERE Codigo= ? ", _
                0, lNumIntDoc, objTRPAportes.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 190426

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190427

    'Se não existir => ERRO
    If lErro = AD_SQL_SEM_DADOS Then gError 190428
    
    objTRPAportes.lNumIntDoc = lNumIntDoc

    'Trava registro contra alterações/Leituras
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> SUCESSO Then gError 190429

    lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM TRPAportes", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 190430
    
    lErro = CF("TRPAportePagtoCond_Exclui_Aporte", objTRPAportes)
    If lErro <> SUCESSO Then gError 190431

    lErro = CF("TRPAportePagtoFat_Exclui_Aporte", objTRPAportes)
    If lErro <> SUCESSO Then gError 190432

    lErro = CF("TRPAportePagtoDireto_Exclui_Aporte", objTRPAportes)
    If lErro <> SUCESSO Then gError 190433

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TRPAportes_Exclui_EmTrans = SUCESSO

    Exit Function

Erro_TRPAportes_Exclui_EmTrans:

    TRPAportes_Exclui_EmTrans = gErr

    Select Case gErr

        Case 190425
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 190426, 190427
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPAPORTES", gErr)

        Case 190428
            Call Rotina_Erro(vbOKOnly, "ERRO_TRPAPORTES_NAO_CADASTRADO", gErr, objTRPAportes.lCodigo)

        Case 190429
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRPAPORTES", gErr)

        Case 190430
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRPAPORTES", gErr)
            
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

Public Function TRPAportePagtoCond_Exclui_Aporte(ByVal objAporte As ClassTRPAportes) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Integer
Dim iTipoDestino As Integer
Dim lNumIntDocDestino As Long

On Error GoTo Erro_TRPAportePagtoCond_Exclui_Aporte

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 190435
    Next

    'Le a tabelaTRPAportePagtoCond
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc, TipoDocDestino, NumIntDocDestino FROM TRPAportePagtoCond WHERE NumIntDocAporte = ? ", _
                0, lNumIntDoc, iTipoDestino, lNumIntDocDestino, objAporte.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 190436

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190437

    Do While lErro <> AD_SQL_SEM_DADOS

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 190438
    
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM TRPAportePagtoCond", alComando(0))
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

    TRPAportePagtoCond_Exclui_Aporte = SUCESSO

    Exit Function

Erro_TRPAportePagtoCond_Exclui_Aporte:

    TRPAportePagtoCond_Exclui_Aporte = gErr

    Select Case gErr

        Case 190435
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 190436, 190437, 190440
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPAPORTEPAGTOCOND", gErr)

        Case 190438
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRPAPORTEPAGTOCOND", gErr)

        Case 190439
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRPAPORTEPAGTOCOND", gErr)
            
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

Public Function TRPAportePagtoDireto_Exclui_Aporte(ByVal objAporte As ClassTRPAportes) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Integer
Dim iTipoDestino As Integer
Dim lNumIntDocDestino As Long

On Error GoTo Erro_TRPAportePagtoDireto_Exclui_Aporte

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 190442
    Next

    'Le a tabelaTRPAportePagtoDireto
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc, TipoDocDestino, NumIntDocDestino FROM TRPAportePagtoDireto WHERE NumIntDocAporte = ? ", _
                0, lNumIntDoc, iTipoDestino, lNumIntDocDestino, objAporte.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 190443

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190444

    Do While lErro <> AD_SQL_SEM_DADOS

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 190445
    
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM TRPAportePagtoDireto", alComando(0))
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

    TRPAportePagtoDireto_Exclui_Aporte = SUCESSO

    Exit Function

Erro_TRPAportePagtoDireto_Exclui_Aporte:

    TRPAportePagtoDireto_Exclui_Aporte = gErr

    Select Case gErr

        Case 190442
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 190443, 190444, 190447
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPAPORTEPAGTODIRETO", gErr)

        Case 190445
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRPAPORTEPAGTODIRETO", gErr)

        Case 190446
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRPAPORTEPAGTODIRETO", gErr)
            
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

Public Function TRPAportePagtoFat_Exclui_Aporte(ByVal objAporte As ClassTRPAportes) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Integer

On Error GoTo Erro_TRPAportePagtoFat_Exclui_Aporte

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 190449
    Next

    'Le a tabelaTRPAportePagtoFat
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM TRPAportePagtoFat WHERE NumIntDocAporte = ? ", _
                0, lNumIntDoc, objAporte.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 190450

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190451

    Do While lErro <> AD_SQL_SEM_DADOS

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 190452
    
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM TRPAportePagtoFat", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 190453

        'Busca Proximo
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190454

    Loop

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TRPAportePagtoFat_Exclui_Aporte = SUCESSO

    Exit Function

Erro_TRPAportePagtoFat_Exclui_Aporte:

    TRPAportePagtoFat_Exclui_Aporte = gErr

    Select Case gErr

        Case 190449
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 190450, 190451, 190454
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPAPORTEPAGTOFAT", gErr)

        Case 190452
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRPAPORTEPAGTOFAT", gErr)

        Case 190453
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRPAPORTEPAGTOFAT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190455)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function TRPAportes_Libera(ByVal colAportePagtos As Collection) As Long

Dim lErro As Long
Dim lTransacao As Long
Dim objTRPAportePagtos As ClassTRPAportePagtoCond
Dim alComando(0 To 4) As Long
Dim iIndice As Integer

On Error GoTo Erro_TRPAportes_Libera

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 190746
    
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 190747
    Next
    
    For Each objTRPAportePagtos In colAportePagtos

        lErro = TRPAportes_Libera_EmTrans(alComando, objTRPAportePagtos)
        If lErro <> SUCESSO Then gError 190748
        
    Next
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 190749

    TRPAportes_Libera = SUCESSO

    Exit Function

Erro_TRPAportes_Libera:

    TRPAportes_Libera = gErr

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

Public Function TRPAportes_Libera_EmTrans(alComando() As Long, ByVal objTRPAportePagto As ClassTRPAportePagtoCond) As Long

Dim lErro As Long
Dim lNumIntDoc As Long
Dim objDebitoRecCli As New ClassDebitoRecCli
Dim colInfoComissao As New colInfoComissao
Dim objContabil As ClassContabil
Dim objAporte As New ClassTRPAportes
Dim lCodigoAuto As Long
Dim sFiltro As String
Dim objTRPAportePagtoBD As New ClassTRPAportePagtoCond

On Error GoTo Erro_TRPAportes_Libera_EmTrans

    objTRPAportePagto.iStatus = STATUS_TRP_OCR_LIBERADO

    objAporte.lNumIntDoc = objTRPAportePagto.lNumIntDocAporte
    
    lErro = CF("TRPAportes_Le", objAporte)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 190753

    If objTRPAportePagto.iFormaPagto = FORMAPAGTO_TRP_OCR_FAT Then
        'Não precisa fazer nada => O Tratamento é dado quando chegar a nota
    Else
        'Gera o crédito e informa o número do numintdoc do débito
          
        lErro = CF("Config_ObterAutomatico2", TRP_TIPO_DOC_DESTINO_DEBCLI_TABELA, "NumTitulo", lCodigoAuto, sFiltro)
        If lErro <> SUCESSO Then gError 190754
           
        objDebitoRecCli.iFilialEmpresa = giFilialEmpresa
        objDebitoRecCli.lCliente = objAporte.lCliente
        objDebitoRecCli.iFilial = objAporte.iFilialCliente
        objDebitoRecCli.sSiglaDocumento = "NCR" 'Nota de crédito
        objDebitoRecCli.lNumTitulo = lCodigoAuto
        objDebitoRecCli.dtDataEmissao = gdtDataAtual
        objDebitoRecCli.dValorTotal = objTRPAportePagto.dValor
        objDebitoRecCli.dValorProdutos = objTRPAportePagto.dValor
        
        lErro = CF("DebitoRecCli_Grava_EmTrans", objDebitoRecCli, colInfoComissao, objContabil)
        If lErro <> SUCESSO Then gError 190755
        
        objTRPAportePagto.lNumIntDocDestino = objDebitoRecCli.lNumIntDoc
        objTRPAportePagto.iTipoDocDestino = TRP_TIPO_DOC_DESTINO_DEBCLI
    
    End If
    
    objTRPAportePagtoBD.lNumIntDoc = objTRPAportePagto.lNumIntDoc
    
    lErro = CF("TRPAportePagtoCond_Le_NumIntDoc", alComando(0), objTRPAportePagtoBD, True)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 190756

    'Se não existir => erro
    If lErro <> SUCESSO Then gError 190318

    'Altera o Status da Aporte e a forma de pagamento
    lErro = Comando_ExecutarPos(alComando(1), "UPDATE TRPAportePagtoCond SET Status= ? , FormaPagto= ?, NumIntDocDestino = ?, DataLiberacao = ? , Usuario = ?, TipoDocDestino = ?, Valor = ? ", alComando(0), objTRPAportePagto.iStatus, objTRPAportePagto.iFormaPagto, objTRPAportePagto.lNumIntDocDestino, Date, gsUsuario, objTRPAportePagto.iTipoDocDestino, objTRPAportePagto.dValor)
    If lErro <> AD_SQL_SUCESSO Then gError 190759

    TRPAportes_Libera_EmTrans = SUCESSO

    Exit Function

Erro_TRPAportes_Libera_EmTrans:

    TRPAportes_Libera_EmTrans = gErr

    Select Case gErr
    
        Case 190753 To 190756
         
        Case 190758
            Call Rotina_Erro(vbOKOnly, "ERRO_TRPAPORTEPAGTOCOND_NAO_CADASTRADO", gErr, objTRPAportePagto.lNumIntDoc)

        Case 190759
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRPAPORTEPAGTOCOND", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190760)

    End Select

    Exit Function

End Function

Public Function TRPVouchers_Grava(ByVal objVou As ClassTRPVouchers, Optional ByVal iTipoGravacao As Integer = TRP_TIPO_GRAVACAO_VOU_MANUTENCAO) As Long

Dim lErro As Long
Dim alComando(0 To 26) As Long
Dim iIndice As Integer
Dim lNumFat As Long
Dim dValor As Double
Dim iCartao As Integer
Dim dValorBruto As Double
Dim dPercAg As Double
Dim iStatus As Integer
Dim iDestino As Integer
Dim iDiasAntc As Integer
Dim sProduto As String
Dim lEmissor As Long
Dim lCorrentista As Long
Dim lRepresentante As Long
Dim lClienteComissao As Long
Dim dComissaoRep As Double
Dim dComissaoCorr As Double
Dim dComissaoAg As Double
Dim dComissaoEmissor As Double
Dim lClienteVou As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objVouExp As New ClassTRPVouExp
Dim lCodEndereco As Long
Dim iPromotor As Integer
Dim iGeraComissao As Integer
Dim dtDataVigenciaDe As Date
Dim objVouBD As New ClassTRPVouchers
Dim dBrutoReal As Double
Dim dCMAReal As Double
Dim dCMCReal As Double
Dim dCMCCReal As Double
Dim dCMRReal As Double
Dim dCMEReal As Double
Dim objVouInfo As ClassTRPVoucherInfo
Dim lClienteFat As Long

On Error GoTo Erro_TRPVouchers_Grava

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 190857
    Next
    
    If objVou.iCartao = DESMARCADO Then
        objVou.lClienteFat = objVou.lClienteVou
    Else
                
        'Le o Cliente
        lErro = Comando_Executar(alComando(26), "SELECT R.Cliente FROM AdmMeioPagto AS P, Redes AS R WHERE P.FilialEmpresa = 1 AND  P.Nome = ? AND P.Rede = R.Codigo", _
                    lClienteFat, objVou.sCiaCart)
        If lErro <> AD_SQL_SUCESSO Then gError 190858
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(26))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190858
    
        objVou.lClienteFat = lClienteFat
    
    End If
    
    objVou.lClienteComissao = objVou.lClienteVou
    
    sProduto = String(STRING_PRODUTO, 0)
    
    'Le a tabelaTRPVouchers
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumFat, Valor, ValorBruto, Cartao, ComissaoAg, Status, Destino, DiasAntc, Produto, Emissor, Correntista, Representante,ClienteVou, ClienteVou,ComissaoRep,ComissaoCorr,ComissaoAg,ComissaoEmissor, GeraComissao, DataVigenciaDe FROM TRPVouchers WHERE NumVou = ? AND Serie = ? AND TipVou= ? ", 0, _
    lNumFat, dValor, dValorBruto, iCartao, dPercAg, iStatus, iDestino, iDiasAntc, sProduto, lEmissor, lCorrentista, lRepresentante, lClienteVou, lClienteComissao, dComissaoRep, dComissaoCorr, dComissaoAg, dComissaoEmissor, iGeraComissao, dtDataVigenciaDe, objVou.lNumVou, objVou.sSerie, objVou.sTipVou)
    If lErro <> AD_SQL_SUCESSO Then gError 190858

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190859
    
    If lClienteComissao = 0 Then lClienteComissao = lClienteVou
    
    objVouBD.sSerie = objVou.sSerie
    objVouBD.sTipVou = objVou.sTipVou
    objVouBD.lNumVou = objVou.lNumVou
    
    lErro = CF("TRPVouchers_Le", objVouBD)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 190860
    
    'Se existir => UPDATE, senão => INSERT
    If lErro = AD_SQL_SUCESSO Then

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 190860
        
        If iTipoGravacao = TRP_TIPO_GRAVACAO_VOU_MANUTENCAO Then
        
            'If lNumFat <> 0 Then gError 200267
            If iStatus = STATUS_TRP_VOU_CANCELADO Then gError 200419 'Não pode mexer em voucher cancelado
            If iGeraComissao = MARCADO Then gError 200413 'É necessário executar a rotina de geração de comissão primeiro
            If Abs(dComissaoAg - objVou.dComissaoAg) > DELTA_VALORMONETARIO And lNumFat <> 0 And objVou.iCartao = DESMARCADO Then gError 200276
            If objVou.dComissaoEmissor > DELTA_VALORMONETARIO And objVou.iDiasAntc = MARCADO Then gError 200277
            'If dtDataVigenciaDe <= gdtDataAtual Then gError 200420 'Não pode mexer em voucher que já entrou em vigência
               
            lErro = TRPVouHistorico_Grava(objVou, objVouBD)
            If lErro <> SUCESSO Then gError 190860
               
            '=> GRAVAR ENDERECO
            lErro = CF("Endereco_Grava", objVou.objEnderecoPax)
            If lErro <> SUCESSO Then gError 200269
            
            '=> GRAVAR PAX
            lErro = TRPVouPassageiros_Exclui(alComando(16), alComando(17), objVou)
            If lErro <> SUCESSO Then gError 200399
            
            lErro = TRPVouPassageiros_Inserir(alComando(15), objVou)
            If lErro <> SUCESSO Then gError 200399
            
            '=> GRAVAR CONTATOS
            lErro = TRPVouContatos_Exclui(alComando(18), alComando(19), objVou)
            If lErro <> SUCESSO Then gError 200399
            
            lErro = TRPVouContatos_Inserir(alComando(20), objVou)
            If lErro <> SUCESSO Then gError 200399
            
            '=> GRAVAR VENDEDORES
            lErro = TRPVouVendedores_Exclui(alComando(21), alComando(22), objVou)
            If lErro <> SUCESSO Then gError 200399
            
            lErro = TRPVouVendedores_Grava(objVou, alComando(10), alComando(11), alComando(12), alComando(13), alComando(14), alComando(23), alComando(24), alComando(25), True)
            If lErro <> SUCESSO Then gError 200399
            
            'Se não entrou em vigência e não foi faturado grava completo, senão só alguns campos
            If (lNumFat = 0 And dtDataVigenciaDe >= Date) Or objVou.dtDataLibManut <> DATA_NULA Then

                lErro = Comando_ExecutarPos(alComando(1), "UPDATE TRPVouchers SET ClienteFat= ? , Data= ?,  " & _
                        "Cartao= ? , Titular= ? , Produto= ? , Pax= ? , Moeda= ? , ValorCambio= ? , Cambio= ? , Controle= ? , CiaCart= ? , " & _
                        "NumCCred= ? , NumAuto= ? , QuantParc= ? , DiasAntc= ? , Kit= ? , " & _
                        "ValorBruto= ? , PassageiroNome= ? , PassageiroSobrenome= ? , ClienteVou= ?,  " & _
                        "DataVigenciaDe= ? , DataVigenciaAte= ? , ClienteComissao= ? , " & _
                        "ciaaerea= ? , Aeroportos= ? , EnderecoPaxTitular= ? , Destino= ? , TitularCPF= ? , TarifaUnitaria= ? , Vigencia= ? , TarifaUnitariaFolheto= ?, PassageiroCGC = ?, PassageiroDataNasc = ?, Idioma = ?, Grupo = ?, UsuarioManut = ?, DataUltimaManut = ?, HoraUltimaManut = ?, DataAutoCC = ?, ValidadeCC = ?, " & _
                        "Emissor = ?,  ValorComissao = ?, Promotor = ?, Representante = ?, ComissaoRep = ?, Correntista = ?, ComissaoCorr = ?, ComissaoAg = ?, ComissaoEmissor = ?, ComissaoProm = ?, CodSegurancaCC = ?, ImprimirValor = ?, Observacao = ?, UsuarioLibManut= ?, DataLibManut = ?, HoraLibManut = ?  ", alComando(0), _
                        objVou.lClienteFat, objVou.dtData, _
                        objVou.iCartao, objVou.sTitular, objVou.sProduto, objVou.iPax, objVou.iMoeda, objVou.dValorCambio, objVou.dCambio, objVou.sControle, objVou.sCiaCart, _
                        objVou.sNumCCred, objVou.lNumAuto, objVou.iQuantParc, objVou.iDiasAntc, objVou.iKit, _
                        objVou.dValorBruto, objVou.sPassageiroNome, objVou.sPassageiroSobreNome, objVou.lClienteVou, _
                        objVou.dtDataVigenciaDe, objVou.dtDataVigenciaAte, objVou.lClienteComissao, _
                        objVou.sCiaaerea, objVou.sAeroportos, objVou.lEnderecoPaxTitular, _
                        objVou.iDestino, objVou.sTitularCPF, objVou.dTarifaUnitaria, objVou.iVigencia, objVou.dTarifaUnitariaFolheto, objVou.sPassageiroCGC, objVou.dtPassageiroDataNasc, objVou.iIdioma, objVou.iGrupo, gsUsuario, Date, CDbl(Time), objVou.dtDataAutoCC, objVou.sValidadeCC, _
                        objVou.lEmissor, Arredonda_Moeda(objVou.dValorBruto * objVou.dComissaoAg), objVou.iPromotor, objVou.lRepresentante, objVou.dComissaoRep, objVou.lCorrentista, objVou.dComissaoCorr, objVou.dComissaoAg, objVou.dComissaoEmissor, objVou.dComissaoProm, objVou.iCodSegurancaCC, objVou.iImprimirValor, objVou.sObservacao, objVou.sUsuarioLibManut, objVou.dtDataLibManut, objVou.dHoraLibManut)
                    If lErro <> AD_SQL_SUCESSO Then gError 190861
            Else
                'Se estiver faturado também não deixa alterar para cartão
                If lNumFat <> 0 Then
                    lErro = Comando_ExecutarPos(alComando(1), "UPDATE TRPVouchers SET Controle= ? , Kit= ? , Idioma = ?, Grupo = ?, UsuarioManut = ?, DataUltimaManut = ?, HoraUltimaManut = ?,Emissor = ?, Promotor = ?, Representante = ?, ComissaoRep = ?, Correntista = ?, ComissaoCorr = ?, ComissaoAg = ?, ComissaoEmissor = ?, ComissaoProm = ?, ImprimirValor = ?, Observacao = ? ", alComando(0), _
                            objVou.sControle, objVou.iKit, objVou.iIdioma, objVou.iGrupo, gsUsuario, Date, CDbl(Time), objVou.lEmissor, objVou.iPromotor, objVou.lRepresentante, objVou.dComissaoRep, objVou.lCorrentista, objVou.dComissaoCorr, objVou.dComissaoAg, objVou.dComissaoEmissor, objVou.dComissaoProm, objVou.iImprimirValor, objVou.sObservacao)
                    If lErro <> AD_SQL_SUCESSO Then gError 190861
                Else
                    lErro = Comando_ExecutarPos(alComando(1), "UPDATE TRPVouchers SET Cartao= ? , Titular= ? , NumCCred= ? , NumAuto= ? , QuantParc= ?, Controle = ? , Kit= ? , " & _
                            "TitularCPF= ? , Idioma = ?, Grupo = ?, UsuarioManut = ?, DataUltimaManut = ?, HoraUltimaManut = ?, DataAutoCC = ?, ValidadeCC = ?, Emissor = ?, ClienteComissao= ?,  ValorComissao = ?, Promotor = ?, Representante = ?, ComissaoRep = ?, Correntista = ?, ComissaoCorr = ?, ComissaoAg = ?, ComissaoEmissor = ?, ComissaoProm = ?, CodSegurancaCC = ?, ImprimirValor = ?, CiaCart = ?, ClienteFat = ?, Observacao = ? ", alComando(0), _
                            objVou.iCartao, objVou.sTitular, objVou.sNumCCred, objVou.lNumAuto, objVou.iQuantParc, objVou.sControle, objVou.iKit, _
                            objVou.sTitularCPF, objVou.iIdioma, objVou.iGrupo, gsUsuario, Date, CDbl(Time), objVou.dtDataAutoCC, objVou.sValidadeCC, objVou.lEmissor, objVou.lClienteComissao, objVou.dValorBruto - objVou.dValor, objVou.iPromotor, objVou.lRepresentante, objVou.dComissaoRep, objVou.lCorrentista, objVou.dComissaoCorr, objVou.dComissaoAg, objVou.dComissaoEmissor, objVou.dComissaoProm, objVou.iCodSegurancaCC, objVou.iImprimirValor, objVou.sCiaCart, objVou.lClienteFat, objVou.sObservacao)
                    If lErro <> AD_SQL_SUCESSO Then gError 190861
                End If
            End If

            
            '###############################################
        
        ElseIf iTipoGravacao = TRP_TIPO_GRAVACAO_VOU_COMISSAO Then
        
                If Abs(dComissaoAg - objVou.dComissaoAg) > DELTA_VALORMONETARIO And lNumFat <> 0 And objVou.iCartao = DESMARCADO Then gError 200276
                If objVou.dComissaoEmissor > DELTA_VALORMONETARIO And objVou.iDiasAntc = MARCADO Then gError 200277
                If gdtDataInicioComisCorp > objVou.dtData Then gError 200278
                If iGeraComissao = MARCADO Then gError 200413 'É necessário executar a rotina de geração de comissão primeiro
        
                lErro = TRPVouHistorico_Grava(objVou, objVouBD)
                If lErro <> SUCESSO Then gError 190860
        
                lErro = Comando_ExecutarPos(alComando(1), "UPDATE TRPVouchers SET Emissor = ?, ValorComissao = ?, Promotor = ?, Representante = ?, ComissaoRep = ?, Correntista = ?, ComissaoCorr = ?, ComissaoAg = ?, ComissaoEmissor = ?  ", alComando(0), _
                        objVou.lEmissor, Arredonda_Moeda(objVou.dValorBruto * objVou.dComissaoAg), objVou.iPromotor, objVou.lRepresentante, objVou.dComissaoRep, objVou.lCorrentista, objVou.dComissaoCorr, objVou.dComissaoAg, objVou.dComissaoEmissor)
                If lErro <> AD_SQL_SUCESSO Then gError 190861
        
        ElseIf iTipoGravacao = TRP_TIPO_GRAVACAO_ACERTO_VALOR_LIQ Then
            
                objVou.dValorComissao = Arredonda_Moeda(objVou.dValorBruto * objVou.dComissaoAg)
                
                If objVou.iCartao = MARCADO Then
                    objVou.dValor = Arredonda_Moeda(objVou.dValorBruto)
                Else
                    objVou.dValor = Arredonda_Moeda(objVou.dValorBruto - objVou.dValorComissao)
                End If
                
                lErro = CF("TRPVoucherInfo_Le", objVouBD)
                If lErro <> SUCESSO Then gError 200269
                
                dBrutoReal = 0
                dCMAReal = 0
                dCMCReal = 0
                dCMCCReal = 0
                dCMRReal = 0
                dCMEReal = 0
                For Each objVouInfo In objVouBD.colTRPVoucherInfo
                    If objVouInfo.sTipoDoc = TRP_TIPODOC_CMCC_TEXTO Then
                        dCMCCReal = dCMCCReal + objVouInfo.dValor
                    End If
                    If objVouInfo.sTipoDoc = TRP_TIPODOC_CMA_TEXTO Then
                        dCMAReal = dCMAReal + objVouInfo.dValor
                    End If
                    If objVouInfo.sTipoDoc = TRP_TIPODOC_BRUTO_TEXTO Then
                        dBrutoReal = dBrutoReal + objVouInfo.dValor
                    End If
                    If objVouInfo.sTipoDoc = TRP_TIPODOC_CMC_TEXTO Then
                        dCMCReal = dCMCReal + objVouInfo.dValor
                    End If
                    If objVouInfo.sTipoDoc = TRP_TIPODOC_CMR_TEXTO Then
                        dCMRReal = dCMRReal + objVouInfo.dValor
                    End If
                    If objVouInfo.sTipoDoc = TRP_TIPODOC_OVER_TEXTO Then
                        dCMEReal = dCMEReal + objVouInfo.dValor
                    End If
                Next
        
                lErro = Comando_ExecutarPos(alComando(1), "UPDATE TRPVouchers SET Valor = ?, ValorComissao = ?, ValorBrutoComOCR = ?, ValorCMAComOCR = ?, ValorCMC = ?, ValorCMR = ?, ValorCMCC = ?, ValorCME = ?", alComando(0), _
                        objVou.dValor, objVou.dValorComissao, Arredonda_Moeda(dBrutoReal), Arredonda_Moeda(dCMAReal), Arredonda_Moeda(dCMCReal), Arredonda_Moeda(dCMRReal), Arredonda_Moeda(dCMCCReal), Arredonda_Moeda(dCMEReal))
                If lErro <> AD_SQL_SUCESSO Then gError 190861
        
        End If

    Else
    
        '=> GRAVAR PAX
        lErro = TRPVouPassageiros_Inserir(alComando(15), objVou)
        If lErro <> SUCESSO Then gError 200399
        
        '=> GRAVAR CONTATOS
        lErro = TRPVouContatos_Inserir(alComando(20), objVou)
        If lErro <> SUCESSO Then gError 200399
            
        '=> GRAVAR VENDEDORES
        lErro = TRPVouVendedores_Grava(objVou, alComando(10), alComando(11), alComando(12), alComando(13), alComando(14), alComando(23), alComando(24), alComando(25))
        If lErro <> SUCESSO Then gError 200399
    
        '=> GRAVAR ENDERECO
        lErro = CF("Endereco_Automatico", lCodEndereco)
        If lErro <> SUCESSO Then gError 200268

        objVou.lEnderecoPaxTitular = lCodEndereco
        objVou.objEnderecoPax.lCodigo = lCodEndereco

        lErro = CF("Endereco_Grava", objVou.objEnderecoPax)
        If lErro <> SUCESSO Then gError 200269
    
        lErro = Comando_Executar(alComando(1), "INSERT INTO TRPVouchers( NumIntDoc, NumVou, Serie, ClienteFat, " & _
                "Valor, Data, TipVou, NumFat, TipoDocDestino, " & _
                "NumIntDocDestino, Cartao, Status, NumIntDocNVL, ValorAporte, " & _
                "ValorComissao, Titular, Produto, Pax, Moeda, " & _
                "ValorCambio, Cambio, Controle, CiaCart, NumCCred, " & _
                "NumAuto, QuantParc, DiasAntc, Kit, ValorOcr, " & _
                "TemOcr, Representante, ComissaoRep, Correntista, ComissaoCorr, " & _
                "Emissor, ComissaoEmissor, ComissaoAg, ValorBruto, PassageiroNome, " & _
                "PassageiroSobrenome, ClienteVou, CliPassageiro, DataCanc, DataVigenciaDe, " & _
                "DataVigenciaAte, HoraCanc, UsuarioCanc, ClienteComissao, ciaaerea, " & _
                "Aeroportos, EnderecoPaxTitular, TemQueContabilizar, Promotor, Destino, " & _
                "TitularCPF, TarifaUnitaria, Vigencia, TarifaUnitariaFolheto, PassageiroCGC,PassageiroDataNasc, Grupo, Idioma, DataAutoCC, ValidadeCC,CodSegurancaCC, ImprimirValor, Observacao) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                objVou.lNumIntDoc, objVou.lNumVou, objVou.sSerie, objVou.lClienteFat, _
                objVou.dValor, objVou.dtData, objVou.sTipVou, objVou.lNumFat, objVou.iTipoDocDestino, _
                objVou.lNumIntDocDestino, objVou.iCartao, objVou.iStatus, objVou.lNumIntDocNVL, objVou.dValorAporte, _
                objVou.dValorComissao, objVou.sTitular, objVou.sProduto, objVou.iPax, objVou.iMoeda, _
                objVou.dValorCambio, objVou.dCambio, objVou.sControle, objVou.sCiaCart, objVou.sNumCCred, _
                objVou.lNumAuto, objVou.iQuantParc, objVou.iDiasAntc, objVou.iKit, objVou.dValorOcr, _
                objVou.iTemOcr, objVou.lRepresentante, objVou.dComissaoRep, objVou.lCorrentista, objVou.dComissaoCorr, _
                objVou.lEmissor, objVou.dComissaoEmissor, objVou.dComissaoAg, objVou.dValorBruto, objVou.sPassageiroNome, _
                objVou.sPassageiroSobreNome, objVou.lClienteVou, objVou.lCliPassageiro, DATA_NULA, objVou.dtDataVigenciaDe, _
                objVou.dtDataVigenciaAte, objVou.dHoraCanc, objVou.sUsuarioCanc, objVou.lClienteComissao, objVou.sCiaaerea, _
                objVou.sAeroportos, objVou.lEnderecoPaxTitular, objVou.iTemQueContabilizar, objVou.iPromotor, objVou.iDestino, _
                objVou.sTitularCPF, objVou.dTarifaUnitaria, objVou.iVigencia, objVou.dTarifaUnitariaFolheto, objVou.sPassageiroCGC, objVou.dtPassageiroDataNasc, objVou.iGrupo, objVou.iIdioma, objVou.dtDataAutoCC, objVou.sValidadeCC, objVou.iCodSegurancaCC, objVou.iImprimirValor, objVou.sObservacao)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190862
            
    End If
    
    'Se não está só ajustando os valores após a geração da comissão
    If iTipoGravacao <> TRP_TIPO_GRAVACAO_ACERTO_VALOR_LIQ Then

        If objVou.lClienteComissao = 0 Then objVou.lClienteComissao = objVou.lClienteVou
                
        'Se está emitindo gera a comissão do zero
        If iTipoGravacao = TRP_TIPO_GRAVACAO_VOU_EMISSAO Then
            If objVou.dValorBruto <> 0 Then
                lErro = TRPVou_Gera_Comis(objVou, objVou.dValorBruto, TRP_TIPO_VALOR_BASE_BRU, objVou.dtData, "Voucher emitido", TRP_TIPO_TRATAMENTO_COMI_VOU_NOVO)
                If lErro <> SUCESSO Then gError 196314
            End If
        Else
        
            'Mudou o valor do voucher
            If Abs(dValorBruto - objVou.dValorBruto) > DELTA_VALORMONETARIO Then
                objVouBD.dValorBruto = objVou.dValorBruto
                lErro = TRPVou_Gera_Comis(objVouBD, objVou.dValorBruto - dValorBruto, TRP_TIPO_VALOR_BASE_BRU, gdtDataAtual, "Alteração no valor do voucher Valor Ant.: " & Format(dValorBruto, "STANDARD") & " Valor Atual: " & Format(objVou.dValorBruto, "STANDARD") & ".", TRP_TIPO_TRATAMENTO_COMI_VOU_VLR_ALTERADO)
                objVouBD.dValorBruto = dValorBruto
                If lErro <> SUCESSO Then gError 196316
            End If
        
            'Mudou para pagto com cartão ou deixou de ser pagto com cartão
            If iCartao <> objVou.iCartao Then
                objVouBD.iCartao = objVou.iCartao
                lErro = TRPVou_Gera_Comis(objVouBD, 0, TRP_TIPO_VALOR_BASE_BRU, gdtDataAtual, "Alteração no tipo de pagamento do voucher. Cartão Ant.: " & IIf(iCartao = MARCADO, "SIM", "NÃO") & " , Cartão Atual: " & IIf(objVou.iCartao = MARCADO, "SIM", "NÃO") & ".", TRP_TIPO_TRATAMENTO_COMI_VOU_TRANSF_CARTAO)
                objVouBD.iCartao = iCartao
                If lErro <> SUCESSO Then gError 196316
            End If
                    
            'Mudou para pagto antecipado ou deixou de ser pagto antecipado
            If iDiasAntc <> objVou.iDiasAntc Then
                objVouBD.iDiasAntc = objVou.iDiasAntc
                If objVou.iDiasAntc = MARCADO Then
                    'Se passou a ser dia anteciado tem que estornar a comissão do emissor
                    lErro = TRPVou_Gera_Comis(objVou, 0, TRP_TIPO_VALOR_BASE_PER, gdtDataAtual, "Alteração no tipo de pagamento do voucher. DiasAntc Ant.: " & IIf(iDiasAntc = MARCADO, "SIM", "NÃO") & " , DiasAntc Atual: " & IIf(objVou.iDiasAntc = MARCADO, "SIM", "NÃO") & ".", TRP_TIPO_TRATAMENTO_COMI_VOU_TRANSF_EMI)
                    If lErro <> SUCESSO Then gError 196316
                Else
                    'Se passou deixou de ser dia anteciado tem que lançar a comissão do emissor
                    lErro = TRPVou_Gera_Comis(objVou, objVou.dComissaoEmissor, TRP_TIPO_VALOR_BASE_PER, gdtDataAtual, "Alteração no tipo de pagamento do voucher. DiasAntc Ant.: " & IIf(iDiasAntc = MARCADO, "SIM", "NÃO") & " , DiasAntc Atual: " & IIf(objVou.iDiasAntc = MARCADO, "SIM", "NÃO") & ".", TRP_TIPO_TRATAMENTO_COMI_OVER_ALT_PERCCOMI)
                    If lErro <> SUCESSO Then gError 196316
                End If
                objVouBD.iDiasAntc = iDiasAntc
            End If
        
            'Se foi uma alteração específica de comissão
            If iTipoGravacao = TRP_TIPO_GRAVACAO_VOU_COMISSAO Or iTipoGravacao = TRP_TIPO_GRAVACAO_VOU_MANUTENCAO Then
                               
                'Mudou de emissor
                If lEmissor <> objVou.lEmissor And dComissaoEmissor <> 0 Then
                    lErro = TRPVou_Gera_Comis(objVou, objVou.dComissaoEmissor / dComissaoEmissor, TRP_TIPO_VALOR_BASE_PER, gdtDataAtual, "Trocou o emissor. Emissor Ant.: " & CStr(lEmissor) & " , Emissor Atual: " & CStr(objVou.lEmissor) & ".", TRP_TIPO_TRATAMENTO_COMI_VOU_TRANSF_EMI)
                    If lErro <> SUCESSO Then gError 196316
                End If
                
                'Mudou de Representante
                If lRepresentante <> objVou.lRepresentante And dComissaoRep <> 0 Then
                    lErro = TRPVou_Gera_Comis(objVou, objVou.dComissaoRep / dComissaoRep, TRP_TIPO_VALOR_BASE_PER, gdtDataAtual, "Trocou o representante. Representante Ant.: " & CStr(lRepresentante) & " , Representante Atual: " & CStr(objVou.lRepresentante) & ".", TRP_TIPO_TRATAMENTO_COMI_VOU_TRANSF_REP)
                    If lErro <> SUCESSO Then gError 196316
                End If
                
                'Mudou o correntista
                If lCorrentista <> objVou.lCorrentista And dComissaoCorr <> 0 Then
                    lErro = TRPVou_Gera_Comis(objVou, objVou.dComissaoCorr / dComissaoCorr, TRP_TIPO_VALOR_BASE_PER, gdtDataAtual, "Trocou o correntista. Correntista Ant.: " & CStr(lCorrentista) & " , Correntista Atual: " & CStr(objVou.lCorrentista) & ".", TRP_TIPO_TRATAMENTO_COMI_VOU_TRANSF_COR)
                    If lErro <> SUCESSO Then gError 196316
                End If
                        
                'Mudou o cliente
'                If lClienteComissao <> objVou.lClienteComissao And dComissaoAg = 0 Then
'                    lErro = TRPVou_Gera_Comis(objVou, 0, TRP_TIPO_VALOR_BASE_PER, gdtDataAtual, "Trocou o cliente. Cliente Ant.: " & CStr(lClienteComissao) & " , Cliente Atual: " & CStr(objVou.lClienteComissao) & ".", TRP_TIPO_TRATAMENTO_COMI_VOU_TRANSF_CLI)
'                    If lErro <> SUCESSO Then gError 196316
'                End If
                If lClienteComissao <> objVou.lClienteComissao And dComissaoAg <> 0 Then
                    lErro = TRPVou_Gera_Comis(objVou, objVou.dComissaoAg / dComissaoAg, TRP_TIPO_VALOR_BASE_PER, gdtDataAtual, "Trocou o cliente. Cliente Ant.: " & CStr(lClienteComissao) & " , Cliente Atual: " & CStr(objVou.lClienteComissao) & ".", TRP_TIPO_TRATAMENTO_COMI_VOU_TRANSF_CLI)
                    If lErro <> SUCESSO Then gError 196316
                End If
                
                lErro = CF("TRPVoucherInfo_Le", objVouBD)
                If lErro <> SUCESSO Then gError 200269
        
                dBrutoReal = 0
                dCMAReal = 0
                dCMCReal = 0
                dCMCCReal = 0
                dCMRReal = 0
                dCMEReal = 0
                For Each objVouInfo In objVouBD.colTRPVoucherInfo
                    If objVouInfo.sTipoDoc = TRP_TIPODOC_CMCC_TEXTO Then
                        dCMCCReal = dCMCCReal + objVouInfo.dValor
                    End If
                    If objVouInfo.sTipoDoc = TRP_TIPODOC_CMA_TEXTO Then
                        dCMAReal = dCMAReal + objVouInfo.dValor
                    End If
                    If objVouInfo.sTipoDoc = TRP_TIPODOC_BRUTO_TEXTO Then
                        dBrutoReal = dBrutoReal + objVouInfo.dValor
                    End If
                    If objVouInfo.sTipoDoc = TRP_TIPODOC_CMC_TEXTO Then
                        dCMCReal = dCMCReal + objVouInfo.dValor
                    End If
                    If objVouInfo.sTipoDoc = TRP_TIPODOC_CMR_TEXTO Then
                        dCMRReal = dCMRReal + objVouInfo.dValor
                    End If
                    If objVouInfo.sTipoDoc = TRP_TIPODOC_OVER_TEXTO Then
                        dCMEReal = dCMEReal + objVouInfo.dValor
                    End If
                Next
                               
                'Mudou o % de comissão do Representante
                If Abs(dComissaoRep - objVou.dComissaoRep) > DELTA_VALORMONETARIO2 And (lRepresentante = objVou.lRepresentante Or dComissaoRep = 0) Then
                    lErro = TRPVou_Gera_Comis(objVou, (dBrutoReal * objVou.dComissaoRep) - dCMRReal, TRP_TIPO_VALOR_BASE_LIQ, gdtDataAtual, "Alteração no % de comissão do representante. % Ant.: " & Format(dComissaoRep, "PERCENT") & " , % Atual: " & Format(objVou.dComissaoRep, "PERCENT") & ".", TRP_TIPO_TRATAMENTO_COMI_CMR_ALT_PERCCOMI)
                    If lErro <> SUCESSO Then gError 196316
                End If
                
                'Mudou o % de comissão do correntista
                If Abs(dComissaoCorr - objVou.dComissaoCorr) > DELTA_VALORMONETARIO2 And (lCorrentista = objVou.lCorrentista Or dComissaoCorr = 0) Then
                    lErro = TRPVou_Gera_Comis(objVou, (dBrutoReal * objVou.dComissaoCorr) - dCMCReal, TRP_TIPO_VALOR_BASE_LIQ, gdtDataAtual, "Alteração no % de comissão do correntista. % Ant.: " & Format(dComissaoCorr, "PERCENT") & " , % Atual: " & Format(objVou.dComissaoCorr, "PERCENT") & ".", TRP_TIPO_TRATAMENTO_COMI_CMC_ALT_PERCCOMI)
                    If lErro <> SUCESSO Then gError 196316
                End If
                
                'Mudou o % de comissão do emissor
                If Abs(dComissaoEmissor - objVou.dComissaoEmissor) > DELTA_VALORMONETARIO2 And (lEmissor = objVou.lEmissor Or dComissaoEmissor = 0) Then
                    lErro = TRPVou_Gera_Comis(objVou, (dBrutoReal * objVou.dComissaoEmissor) - dCMEReal, TRP_TIPO_VALOR_BASE_LIQ, gdtDataAtual, "Alteração no % de comissão do emissor. % Ant.: " & Format(dComissaoEmissor, "PERCENT") & " , % Atual: " & Format(objVou.dComissaoEmissor, "PERCENT") & ".", TRP_TIPO_TRATAMENTO_COMI_OVER_ALT_PERCCOMI)
                    If lErro <> SUCESSO Then gError 196316
                End If
                
                'Mudou o % de comissão do agência
                If Abs(dComissaoAg - objVou.dComissaoAg) > DELTA_VALORMONETARIO2 And (lClienteComissao = objVou.lClienteComissao Or dComissaoAg = 0) Then
                    lErro = TRPVou_Gera_Comis(objVou, (dBrutoReal * objVou.dComissaoAg) - (dCMAReal + dCMCCReal), TRP_TIPO_VALOR_BASE_LIQ, gdtDataAtual, "Alteração no % de comissão da agência. % Ant.: " & Format(dComissaoAg, "PERCENT") & " , % Atual: " & Format(objVou.dComissaoAg, "PERCENT") & ".", TRP_TIPO_TRATAMENTO_COMI_CMA_CMCC_ALT_PERCCOMI)
                    If lErro <> SUCESSO Then gError 196316
                End If
            End If

        End If
    End If
    '########################################################

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TRPVouchers_Grava = SUCESSO

    Exit Function

Erro_TRPVouchers_Grava:

    TRPVouchers_Grava = gErr

    Select Case gErr

        Case 190857
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 190858, 190859
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPVOUCHERS", gErr)

        Case 190860
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRPVOUCHERS", gErr)

        Case 190861
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRPVOUCHERS", gErr)

        Case 190862
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRPVOUCHERS", gErr)
            
        Case 196314, 196315, 196316, 200268, 200269, 200399
        
        Case 200267
            Call Rotina_Erro(vbOKOnly, "ERRO_VOUCHER_FATURADO", gErr)
            
        Case 200276
            Call Rotina_Erro(vbOKOnly, "ERRO_TRP_CMA_NAO_ALT_VOU_FAT", gErr)
            
        Case 200277
            Call Rotina_Erro(vbOKOnly, "ERRO_TRP_OVER_NAO_ALT_VOU_ANTC", gErr)
            
        Case 200278
            Call Rotina_Erro(vbOKOnly, "ERRO_TRP_COMISSAO_NAO_ALTERAVEL_DATA", gErr, gdtDataInicioComisCorp)

        Case 200413
            Call Rotina_Erro(vbOKOnly, "ERRO_TRP_VOU_SEM_COMI", gErr)
            
        Case 200419
            Call Rotina_Erro(vbOKOnly, "ERRO_MANUT_VOU_CANCELADO", gErr)
        
        Case 200420
            Call Rotina_Erro(vbOKOnly, "ERRO_MANUT_VOU_EM_VIG", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190856)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function TRPVouchers_Cancela(ByVal objVou As ClassTRPVouchers) As Long

Dim lErro As Long
Dim alComando(0 To 13) As Long
Dim iIndice As Integer
Dim lNumFat As Long
Dim iStatus As Integer
Dim lNumIntDocTitRec As Long
Dim lNumIntDocDestino As Long
Dim iTipoDocDestino As Integer
Dim sTipoDoc As String
Dim objOcorrencia As ClassTRPOcorrencias
Dim objOcorrenciaDet As ClassTRPOcorrenciaDet
Dim dValor As Double
Dim lCliente As Long
Dim iFilialEmpresa As Integer
Dim sSiglaDocumento As String
Dim lNumVou As Long
Dim sSerie As String
Dim sTipVou As String
Dim lCodigo As Long
Dim iCartao As Integer
Dim lNumIntDocNVL As Long
Dim dValorOcr As Double
Dim dValorComissao As Double
Dim bComissaoCorp As Boolean
Dim lNumIntComi As Long
Dim dValorBruto As Double
Dim sTipoDocAux As String
Dim objVouExp As New ClassTRPVouExp
Dim lClienteVou As Long
Dim lCliPassageiro As Long
Dim iGeraComissao As Integer

On Error GoTo Erro_TRPVouchers_Cancela

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 190848
    Next
    
    sTipoDoc = String(STRING_TRP_OCR_TIPODOC, 0)
    sSerie = String(STRING_TRP_OCR_SERIE, 0)
    sTipVou = String(STRING_TRP_OCR_TIPOVOU, 0)

    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumFat, Status, NumIntDocDestino,TipoDocDestino, Valor, ClienteFat, NumVou, Serie, TipVou, Cartao, NumIntDocNVL, ValorOCR, ValorBruto, ClienteVou, CliPassageiro,GeraComissao FROM TRPVouchers WHERE NumVou = ? AND Serie = ? AND TipVou= ?", _
                        0, lNumFat, iStatus, lNumIntDocDestino, iTipoDocDestino, dValor, lCliente, lNumVou, sSerie, sTipVou, iCartao, lNumIntDocNVL, dValorOcr, dValorBruto, lClienteVou, lCliPassageiro, iGeraComissao, objVou.lNumVou, objVou.sSerie, objVou.sTipVou)
    If lErro <> AD_SQL_SUCESSO Then gError 190849

    'Busca Primeiro documento
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190850
    
    'Se existir
    If lErro <> AD_SQL_SEM_DADOS Then

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 190852
                   
        'Se já estiver cancelado dá erro se for cancelamento manual,
        'na importação é só ignorar
        If iStatus = STATUS_TRP_VOU_CANCELADO Then gError 192760
        If Abs(dValorOcr) > DELTA_VALORMONETARIO Then gError 200710
        If iGeraComissao = MARCADO Then gError 200413
   
        'Se ainda não foi inativado
        If lNumIntDocNVL = 0 Then

            'Criação do documento NVL
            If lNumIntDocDestino <> 0 Then
            
                Set objOcorrencia = New ClassTRPOcorrencias
                Set objOcorrenciaDet = New ClassTRPOcorrenciaDet
                
                objOcorrencia.dtDataEmissao = Date
                objOcorrencia.dValorTotal = -dValor
                objOcorrencia.iFilialCliente = FILIAL_MATRIZ
                objOcorrencia.iFormaPagto = FORMAPAGTO_TRP_OCR_FAT
                objOcorrencia.iOrigem = INATIVACAO_AUTOMATICA_CODIGO
                objOcorrencia.iStatus = STATUS_TRP_OCR_BLOQUEADO
                
                'Tem que pegar o passageiro
                If iCartao = MARCADO Then
                
                    'VERIFICAR SE O PASSAGEIRO ESTÁ PREENCHIDO
                    If lCliPassageiro = 0 Then
                    
                        'CADASTRAR O PASSAGEIRO E ACERTA A INFORMAÇÃO EM VOUCHERINFO
                        lErro = TRP_Voucher_Cadastra_Cliente_EmTrans(objVou)
                        If lErro <> SUCESSO Then gError 192824
                        
                    Else
                        objVou.lCliPassageiro = lCliPassageiro
                    End If
                    
                    objOcorrencia.lCliente = objVou.lCliPassageiro
                
                Else
                    objOcorrencia.lCliente = objVou.lClienteVou
                End If
               
                'Se for inativação de um voucher já Faturado cria o NVL
                If lNumIntDocDestino <> 0 And iTipoDocDestino = TRP_TIPO_DOC_DESTINO_TITREC And sTipoDoc = TRP_TIPODOC_VOU_TEXTO Then
                
                    sSiglaDocumento = String(STRING_SIGLA_DOCUMENTO, 0)
                
                    'Le o título a receber
                    lErro = Comando_Executar(alComando(2), "SELECT NumIntDoc, SiglaDocumento, FilialEmpresa FROM TitulosRecTodos WHERE NumIntDoc = ? ", _
                                lNumIntDocTitRec, sSiglaDocumento, iFilialEmpresa, lNumIntDocDestino)
                    If lErro <> AD_SQL_SUCESSO Then gError 192762
                
                    'Busca Primeiro
                    lErro = Comando_BuscarPrimeiro(alComando(2))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192763
                                        
                End If
        
                lErro = CF("Config_ObterAutomatico_EmTrans", "FATConfig", "NUM_PROX_TRPOCORRENCIAS", "TRPOcorrencias", "Codigo", lCodigo)
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
                
                lErro = TRPOcorrencias_Grava_EmTrans(objOcorrencia)
                If lErro <> SUCESSO Then gError 192765
                                                    
                lErro = Comando_ExecutarPos(alComando(3), "UPDATE TRPVouchers SET NumIntDocNVL = ?", alComando(0), objOcorrencia.lNumIntDoc)
                If lErro <> AD_SQL_SUCESSO Then gError 192766
                    
            'Se não está criando o NVL e for voucher (voucher excluído) => Estorna as comissões
            Else
                
                'Cancelou o voucher sem gerar o NVL
                lErro = TRPVou_Gera_Comis(objVou, -dValorBruto, TRP_TIPO_VALOR_BASE_BRU, gdtDataAtual, "Cancelamento de voucher não faturado.", TRP_TIPO_TRATAMENTO_COMI_VOU_VLR_ALTERADO)
                If lErro <> SUCESSO Then gError 196277
            
            End If
            
            lErro = Comando_ExecutarPos(alComando(1), "UPDATE TRPVouchers SET Status = ?, DataCanc  = ?, HoraCanc = ?, UsuarioCanc = ?", alComando(0), STATUS_TRP_VOU_CANCELADO, Date, CDbl(Time), gsUsuario)
            If lErro <> AD_SQL_SUCESSO Then gError 190846
            
        End If
    
    End If
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TRPVouchers_Cancela = SUCESSO

    Exit Function

Erro_TRPVouchers_Cancela:

    TRPVouchers_Cancela = gErr

    Select Case gErr

        Case 190846, 192766
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRPVOUCHERS", gErr)

        Case 190848
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 190849, 190850, 192761
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPVOUCHERS", gErr)

        Case 190852
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRPVOUCHERS", gErr)

        Case 190853
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRPVOUCHERS", gErr)
            
        Case 192760
            Call Rotina_Erro(vbOKOnly, "ERRO_VOUCHER_JA_CANCELADO", gErr)
            
        Case 192764, 192765, 192824, 200195, 196277
        
        Case 200710
            Call Rotina_Erro(vbOKOnly, "ERRO_CANC_VOU_COM_OCR", gErr)

        Case 200413
            Call Rotina_Erro(vbOKOnly, "ERRO_TRP_VOU_SEM_COMI", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190847)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function TRPFaturamento_Gera(ByVal objFaturamento As ClassTRPFaturamento, ByVal colArquivosHtml As Collection) As Long

Dim lErro As Long
Dim lTransacao As Long
Dim alComando(0 To 42) As Long
Dim iIndice As Integer
Dim colParcelas As New Collection
Dim lFatIni As Long
Dim lFatFim As Long
Dim colNFsPag As New Collection

On Error GoTo Erro_TRPFaturamento_Gera

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 190986
    
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 190987
    Next

    lErro = TRPFaturamento_Gera_EmTrans(alComando, objFaturamento, colParcelas, colArquivosHtml, lFatIni, lFatFim, colNFsPag)
    If lErro <> SUCESSO Then gError 190988

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 190989
  
    Call Rotina_Aviso(vbOKOnly, "AVISO_FAT_TRP_SUCESSO", CStr(lFatIni), CStr(lFatFim))
  
    If objFaturamento.iEnviarPorEmail = MARCADO Then
        
        If colParcelas.Count > 0 Then
            Call TRPFaturamento_Envia_PorEmail(objFaturamento, colParcelas)
        End If
        
        If colNFsPag.Count > 0 Then
            Call TRPFaturamento_Envia_PorEmail_CP(objFaturamento, colNFsPag)
        End If
        
    End If
    
    TRPFaturamento_Gera = SUCESSO

    Exit Function

Erro_TRPFaturamento_Gera:

    TRPFaturamento_Gera = gErr

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

Private Function TRPFaturamento_Gera_EmTrans(alComando() As Long, ByVal objFaturamento As ClassTRPFaturamento, ByVal colParcelas As Collection, ByVal colArquivosHtml As Collection, lFatIni As Long, lFatFim As Long, ByVal colNFsPag As Collection) As Long

Dim lErro As Long
Dim objTitRec As ClassTituloReceber
Dim objTitPag As ClassTituloPagar
Dim objNFsPag As ClassNFsPag
Dim objInfoFatura As ClassTRPFATInfoFatura
Dim objInfoVoucher As ClassTRPFATInfoVoucher
Dim iTipoDestino As Integer
Dim lNumIntDocDestino As Long
Dim lNumIntDocAux As Long
Dim dSaldo As Double
Dim lProxCR As Long
Dim sConteudo As String
Dim objTRPConfig As ClassTRPConfig
Dim colConfig As New Collection
Dim sArquivo As String
Dim sConteudoHtml As String
Dim iFaturaAtual As Integer
Dim objCliente As ClassCliente
Dim objForn As ClassFornecedor
Dim objFilialEmpresa As AdmFiliais
Dim objEndereco As ClassEndereco
Dim objFilialCliente As ClassFilialCliente
Dim objFilialForn As ClassFilialFornecedor
Dim bAchou As Boolean
Dim iLinha As Integer, iStatusAux As Integer
Dim objParcelaRec As ClassParcelaReceber
Dim bAtualizaProxParcela As Boolean
Dim iParcela As Integer
Dim objTRPTitulosExp As ClassTRPTitulosExp
Dim sModeloCP As String
Dim objVouAporte As ClassTRPVoucherAporte
Dim objOCRAporte As ClassTRPOcorrenciaAporte
Dim objFatAporte As ClassTRPFatAporte
Dim objFaturaAporte As ClassTRPFaturaAporte
Dim colParcRecDataVencReal As New Collection
Dim objFrmAguarde As New ClassFrmAguarde
Dim objFrmAguardeTela As New FrmAguarde
Dim sItem As String
Dim objVouExp As ClassTRPVouExp
Dim sTipoDoc As String
Dim colTipos As Collection
Dim vValor As Variant
Dim colCampos As Collection
Dim colSaida As Collection
Dim objTRPTitulos As ClassTRPTitulos
Dim dValorAux As Double
Dim iCartaoAux As Integer
Dim lNumFatAux As Long
Dim sDocAux As String
Dim iNumFatMarcadas As Integer

On Error GoTo Erro_TRPFaturamento_Gera_EmTrans

    lErro = CF("TRPConfig_Le", TRPCONFIG_DIRETORIO_MODELO_FAT_HTML_CARTAO, EMPRESA_TODA, sModeloCP)
    If lErro <> SUCESSO Then gError 192082

    lErro = CF("TRPConfig_Le_Lockado", alComando(6), TRPCONFIG_PROX_NUM_TITREC, EMPRESA_TODA, sConteudo)
    If lErro <> SUCESSO Then gError 192082
    
    lProxCR = StrParaLong(sConteudo)
    lFatIni = lProxCR
    
    iNumFatMarcadas = 0
    For Each objInfoFatura In objFaturamento.colInfoFaturas
        If objInfoFatura.iMarcado = MARCADO Then
            iNumFatMarcadas = iNumFatMarcadas + 1
        End If
    Next
    
    objFrmAguarde.iTotalItens = iNumFatMarcadas
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
                
                If lErro = SUCESSO Then
                
                    lErro = CF("FilialCliente_Le", objFilialCliente, alComando(16))
                    If lErro <> SUCESSO And lErro <> 12567 Then gError 192199
                    
                Else
                    
                    Set objForn = New ClassFornecedor
                    Set objFilialForn = New ClassFilialFornecedor
                    
                    objForn.lCodigo = objInfoFatura.lCliente
                    objFilialForn.lCodFornecedor = objInfoFatura.lCliente
                    objFilialForn.iCodFilial = FILIAL_MATRIZ
                    
                    lErro = CF("Fornecedor_Le", objForn)
                    If lErro <> SUCESSO And lErro <> 12729 Then gError 192197

                    lErro = CF("FilialFornecedor_Le", objFilialForn)
                    If lErro <> SUCESSO And lErro <> 12929 Then gError 192199
                    
                    objCliente.sNomeReduzido = objForn.sNomeReduzido
                    objCliente.sRazaoSocial = objForn.sRazaoSocial
                    
                    objFilialCliente.lEndereco = objFilialForn.lEndereco
                    objFilialCliente.sCgc = objFilialForn.sCgc
                    objFilialCliente.sInscricaoEstadual = objFilialForn.sInscricaoEstadual
                    objFilialCliente.sInscricaoMunicipal = objFilialForn.sInscricaoMunicipal
                    objFilialCliente.sNome = objFilialForn.sNome

                End If
                
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
            
            sTipoDoc = ""
            Set colTipos = New Collection
            For Each objInfoVoucher In objInfoFatura.colInfoVouchers
                If objInfoVoucher.iMarcado = MARCADO Then
                    bAchou = False
                    For Each objTRPTitulos In colTipos
                        If objInfoVoucher.sTipoDoc = objTRPTitulos.sTipoDoc Then
                            bAchou = True
                            Exit For
                        End If
                    Next
                    If Not bAchou Then
                        Set objTRPTitulos = New ClassTRPTitulos
                        objTRPTitulos.sTipoDoc = objInfoVoucher.sTipoDoc
                        colTipos.Add objTRPTitulos
                    End If
                End If
            Next
            Set colCampos = New Collection
            Set colSaida = New Collection
            
            colCampos.Add "sTipoDoc"
            
            lErro = Ordena_Colecao(colTipos, colSaida, colCampos)
            If lErro <> SUCESSO Then gError 200226
            
            For Each objTRPTitulos In colTipos
                If Len(Trim(sTipoDoc)) > 0 Then sTipoDoc = sTipoDoc & SEPARADOR
                sTipoDoc = sTipoDoc & objTRPTitulos.sTipoDoc
            Next

            If objInfoFatura.dValor > 0 Then
                
                Set objTitRec = New ClassTituloReceber
                
                iTipoDestino = TRP_TIPO_DOC_DESTINO_TITREC
                
                lErro = TRPFaturamento_Gera_TitRec(alComando(21), alComando(22), alComando(23), alComando(24), alComando(25), alComando(26), alComando(27), alComando(28), alComando(29), alComando(30), objInfoFatura, objTitRec, lProxCR, colParcRecDataVencReal)
                If lErro <> SUCESSO Then gError 192001
                
                For Each objParcelaRec In objTitRec.colParcelaReceber
                    colParcelas.Add objParcelaRec
                Next
                
                objInfoFatura.lNumeroFaturaReal = objTitRec.lNumTitulo
                
                lProxCR = lProxCR + 1
                
                lNumIntDocDestino = objTitRec.lNumIntDoc
    
            Else
            
                'Ocorrências
                If objInfoFatura.iTipoDoc = TRP_TIPODOC_FAT_OCR Then
            
                    Set objTitPag = New ClassTituloPagar
                    
                    iTipoDestino = TRP_TIPO_DOC_DESTINO_TITPAG
                    
                    lErro = TRPFaturamento_Gera_TitPag(objInfoFatura, objTitPag, lProxCR, objCliente, objFaturamento.colFatEnderecos.Item(objInfoFatura.iLinhaFatEnderecoCli), colParcRecDataVencReal)
                    If lErro <> SUCESSO Then gError 192002
                    
                    objInfoFatura.lNumeroFaturaReal = objTitPag.lNumTitulo
                    
                    lProxCR = lProxCR + 1
                    
                    lNumIntDocDestino = objTitPag.lNumIntDoc
                
                'Over
                ElseIf objInfoFatura.iTipoDoc = TRP_TIPODOC_FAT_OVER Then
                   
                     Set objTitPag = New ClassTituloPagar
                    
                    iTipoDestino = TRP_TIPO_DOC_DESTINO_TITPAG
                    
                    lErro = TRPFaturamento_Gera_TitPagOver(objInfoFatura, objTitPag, lProxCR)
                    If lErro <> SUCESSO Then gError 192002
                    
                    objInfoFatura.lNumeroFaturaReal = objTitPag.lNumTitulo
                    
                    lProxCR = lProxCR + 1
                    
                    lNumIntDocDestino = objTitPag.lNumIntDoc
                   
                'Comissão
                Else
                
                    Set objNFsPag = New ClassNFsPag
                    
                    iTipoDestino = TRP_TIPO_DOC_DESTINO_NFSPAG
                    
                    lErro = TRPFaturamento_Gera_NFPag(objInfoFatura, objNFsPag, lProxCR)
                    If lErro <> SUCESSO Then gError 194202
                    
                    objInfoFatura.lNumeroFaturaReal = objNFsPag.lNumNotaFiscal
                    
                    colNFsPag.Add objNFsPag
                    
                    lProxCR = lProxCR + 1
                    
                    lNumIntDocDestino = objNFsPag.lNumIntDoc
                
                End If
            
            End If
            
            Set objTRPTitulosExp = New ClassTRPTitulosExp
            
            objTRPTitulosExp.iExcluido = DESMARCADO
            objTRPTitulosExp.iExportado = DESMARCADO
            objTRPTitulosExp.iTipoDocOrigem = iTipoDestino
            objTRPTitulosExp.lNumIntDocOrigem = lNumIntDocDestino
            objTRPTitulosExp.lNumTitulo = objInfoFatura.lNumeroFaturaReal
            objTRPTitulosExp.iTemQueContabilizar = MARCADO
            objTRPTitulosExp.dValorAporte = objInfoFatura.dValorAporte
            objTRPTitulosExp.dValorAporteCred = objInfoFatura.dValorAporteCred
            objTRPTitulosExp.dValorComissao = objInfoFatura.dValorComissao
            objTRPTitulosExp.dValorCreditos = objInfoFatura.dValorCreditos
            objTRPTitulosExp.dValorDebitos = objInfoFatura.dValorDebitos
            objTRPTitulosExp.dValorDeducoes = objInfoFatura.dValorDeducoes
            objTRPTitulosExp.dValorTarifa = objInfoFatura.dValorTarifa
            
            For Each objFaturaAporte In objInfoFatura.colFatAporteCred
                objTRPTitulosExp.colFatAporteCred.Add objFaturaAporte
            Next
            
            lErro = CF("TRPTitulosExp_Insere", objTRPTitulosExp, alComando(18), alComando(19))
            If lErro <> SUCESSO Then gError 196240
            
            Set objTRPTitulos = New ClassTRPTitulos

            objTRPTitulos.iTipoDocOrigem = iTipoDestino
            objTRPTitulos.lNumIntDocOrigem = lNumIntDocDestino
            objTRPTitulos.sTipoDoc = sTipoDoc
            
            lErro = CF("TRPTitulos_Insere", objTRPTitulos, alComando(41))
            If lErro <> SUCESSO Then gError 196240
            
            For Each objInfoVoucher In objInfoFatura.colInfoVouchers
            
                If objInfoVoucher.iMarcado = MARCADO Then
            
                    'Se é um voucher
                    If objInfoVoucher.sTipoDoc = TRP_TIPODOC_VOU_TEXTO Then
    
                        lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc, Valor, Cartao, NumFat, Status FROM TRPVouchers WHERE NumVou= ? AND Serie = ? AND TipVou = ? ", _
                                    0, lNumIntDocAux, dValorAux, iCartaoAux, lNumFatAux, iStatusAux, objInfoVoucher.lNumVou, objInfoVoucher.sSerie, objInfoVoucher.sTipoVou)
                        If lErro <> AD_SQL_SUCESSO Then gError 192003
                    
                        'Busca Primeiro
                        lErro = Comando_BuscarPrimeiro(alComando(0))
                        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192004
                        
                        If lErro = AD_SQL_SEM_DADOS Then gError 192016
                    
                        'Trava registro contra alterações/Leituras
                        lErro = Comando_LockExclusive(alComando(0))
                        If lErro <> SUCESSO Then gError 192005
                        
                        If Abs(dValorAux - objInfoVoucher.dValor) > DELTA_VALORMONETARIO Then
                            sDocAux = "VOU " & objInfoVoucher.sTipoVou & " " & objInfoVoucher.sSerie & " " & CStr(objInfoVoucher.lNumVou) & " Vlr.Ant: " & Format(objInfoVoucher.dValor, "STANDARD") & " " & " Vlr.Atual: " & Format(dValorAux, "STANDARD")
                            gError 200382
                        End If
                         If iCartaoAux = MARCADO Then
                            sDocAux = "VOU " & objInfoVoucher.sTipoVou & " " & objInfoVoucher.sSerie & " " & CStr(objInfoVoucher.lNumVou) & " Cartão: Sim"
                            gError 200382
                        End If
                        If lNumFatAux <> 0 Then
                            sDocAux = "VOU " & objInfoVoucher.sTipoVou & " " & objInfoVoucher.sSerie & " " & CStr(objInfoVoucher.lNumVou) & " Fatura: " & CStr(lNumFatAux)
                            gError 200382
                        End If
                        If iStatusAux = STATUS_CANCELADO Then
                            sDocAux = "VOU " & objInfoVoucher.sTipoVou & " " & objInfoVoucher.sSerie & " " & CStr(objInfoVoucher.lNumVou) & " Cancelado: Sim"
                            gError 200382
                        End If
                        
                        lErro = Comando_ExecutarPos(alComando(1), "UPDATE TRPVouchers SET TipoDocDestino= ?, NumIntDocDestino=?, ValorAporte = ?, NumFat = ? ", alComando(0), _
                        iTipoDestino, lNumIntDocDestino, objInfoVoucher.dValorAporte, objInfoFatura.lNumeroFaturaReal)
                        If lErro <> AD_SQL_SUCESSO Then gError 192006
                        
'                        Set objVouExp = New ClassTRPVouExp
'                        objVouExp.sSerie = objInfoVoucher.sSerie
'                        objVouExp.sTipVou = objInfoVoucher.sTipoVou
'                        objVouExp.lNumVou = objInfoVoucher.lNumVou
'                        objVouExp.lNumFat = objInfoFatura.lNumeroFaturaReal
'                        objVouExp.dtdatafat = objInfoFatura.dtDataEmissao
'                        objVouExp.iTransacao = TRP_EXPORT_VOU_TRANS_FATURAMENTO
'                        objVouExp.lCliente = objInfoFatura.lCliente
'                        lErro = TRPVouExp_Insere(alComando(36), alComando(37), alComando(38), alComando(39), alComando(40), alComando(42), objVouExp)
'                        If lErro <> SUCESSO Then gError 200196
                        
                        'Verifica se não tem que liberar o OVER (liberação na emissão)
                        sItem = String(255, 0)
                        
                        'Verifica o tipo de tratamento (Se o pagamento é na baixa ou na emissão)
                        lErro = Comando_Executar(alComando(33), "SELECT Item FROM FilialFornecedorCategorias WHERE Fornecedor = ? AND Filial = ? AND Categoria = ? AND Item = ?", sItem, objInfoVoucher.lCliente, FILIAL_MATRIZ, TRP_CATEGORIA_LIBCME, TRP_CATEGORIA_LIBCME_FAT)
                        If lErro <> AD_SQL_SUCESSO Then gError 182532
                        
                        'Busca a primeira
                        lErro = Comando_BuscarProximo(alComando(33))
                        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182533
                               
                        'Se não for na baixa libera o over para pagamento
                        If lErro = AD_SQL_SUCESSO Then
                            
                            'Le as comissões over bloqueadas para esse voucher
                            lErro = Comando_ExecutarPos(alComando(34), "SELECT NumIntDoc FROM TRPVoucherInfo WHERE NumVou= ? AND TipoDoc = ? AND Serie = ? AND TipVou = ? AND Status = ? ", _
                            0, lNumIntDocAux, objInfoVoucher.lNumVou, TRP_TIPODOC_OVER_TEXTO, objInfoVoucher.sSerie, objInfoVoucher.sTipoVou, TRP_VOU_INFO_STATUS_BLOQUEADO)
                            If lErro <> AD_SQL_SUCESSO Then gError 196376
                        
                            'Busca Primeiro
                            lErro = Comando_BuscarPrimeiro(alComando(34))
                            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196377
                    
                            Do While lErro <> AD_SQL_SEM_DADOS
                                                        
                                'Trava registro contra alterações/Leituras
                                lErro = Comando_LockExclusive(alComando(34))
                                If lErro <> SUCESSO Then gError 196378
                                
                                lErro = Comando_ExecutarPos(alComando(35), "UPDATE TRPVoucherInfo SET Status = ?, TipoLiberacao = ?, NumIntDocLiberacao = ? ", alComando(34), TRP_VOU_INFO_STATUS_LIBERADO, TRP_TIPO_LIBERACAO_COMISSAO_FAT, lNumIntDocDestino)
                                If lErro <> AD_SQL_SUCESSO Then gError 196379
                            
                                'Busca Próximo
                                lErro = Comando_BuscarProximo(alComando(34))
                                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196380
                            
                            Loop
                            
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
                                lErro = TRPAportes_Acerta_ProxParcela(alComando(8), alComando(9), objVouAporte.lCodigoAporte, iParcela)
                                If lErro <> SUCESSO Then gError 192021
                                
                                Set objFatAporte = New ClassTRPFatAporte
                                
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
                                lErro = TRPAportePagtoFat_Acerta_Saldo(alComando(4), alComando(5), alComando(12), alComando(13), objVouAporte.iTipoPagtoAporte, objVouAporte.lNumIntDocPagtoAporte, objVouAporte.dValorAporte)
                            Else
                                lErro = TRPAportePagtoFat_Acerta_Saldo(alComando(4), alComando(5), alComando(12), alComando(13), objVouAporte.iTipoPagtoAporte, objVouAporte.lNumIntDocPagtoAporte, objVouAporte.dValorAporteUSS)
                            End If
                            If lErro <> SUCESSO Then gError 192021
                            
                            lErro = Comando_Executar(alComando(10), "INSERT INTO TRPVoucherAporte( TipVou, Serie, NumVou, TipoPagtoAporte, " & _
                                    "NumIntDocPagtoAporte, CodigoAporte, ParcelaAporte, ValorAporte, ValorAporteUSS, MoedaAporte) VALUES (?,?,?,?,?,?,?,?,?,?)", _
                                    objVouAporte.sTipVou, objVouAporte.sSerie, objVouAporte.lNumVou, objVouAporte.iTipoPagtoAporte, _
                                    objVouAporte.lNumIntDocPagtoAporte, objVouAporte.lCodigoAporte, objVouAporte.iParcelaAporte, objVouAporte.dValorAporte, objVouAporte.dValorAporteUSS, objVouAporte.iMoedaAporte)
                            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196476
                        
                        Next
                        
                    'Quando é uma comissão faturada com a versão de comissionamento (CMC, CMR, CMCC e OVER)
                    ElseIf objInfoVoucher.sTipoDoc = TRP_TIPODOC_CMR_TEXTO Or objInfoVoucher.sTipoDoc = TRP_TIPODOC_CMC_TEXTO Or objInfoVoucher.sTipoDoc = TRP_TIPODOC_CMCC_TEXTO Or objInfoVoucher.sTipoDoc = TRP_TIPODOC_OVER_TEXTO Then
                    
                        lErro = Comando_ExecutarPos(alComando(31), "SELECT NumIntDoc, NumTitulo, Status FROM TRPVoucherInfo WHERE NumIntDoc = ? ", 0, lNumIntDocAux, lNumFatAux, iStatusAux, objInfoVoucher.lNumIntDoc)
                        If lErro <> AD_SQL_SUCESSO Then gError 196371
                    
                        'Busca Primeiro
                        lErro = Comando_BuscarPrimeiro(alComando(31))
                        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196372
                        
                        If lErro = AD_SQL_SEM_DADOS Then gError 196373
                        
                        'Trava registro contra alterações/Leituras
                        lErro = Comando_LockExclusive(alComando(31))
                        If lErro <> SUCESSO Then gError 196374
                        
                        If lNumFatAux <> 0 Then
                            sDocAux = objInfoVoucher.sTipoDoc & " " & objInfoVoucher.sTipoVou & " " & objInfoVoucher.sSerie & " " & CStr(objInfoVoucher.lNumVou) & " Fatura: " & CStr(lNumFatAux)
                            gError 200382
                        End If
                        If iStatusAux <> TRP_VOU_INFO_STATUS_LIBERADO Then
                            sDocAux = objInfoVoucher.sTipoDoc & " " & objInfoVoucher.sTipoVou & " " & objInfoVoucher.sSerie & " " & CStr(objInfoVoucher.lNumVou) & " Liberado: Não"
                            gError 200382
                        End If
                        
                        lErro = Comando_ExecutarPos(alComando(32), "UPDATE TRPVoucherInfo SET TipoDocDestino= ?, NumIntDocDestino=?, NumTitulo = ? ", alComando(31), _
                        iTipoDestino, lNumIntDocDestino, objInfoFatura.lNumeroFaturaReal)
                        If lErro <> AD_SQL_SUCESSO Then gError 196375
                    
                    'Caso seja uma OCR ou um NVL
                    Else
                    
                        lErro = Comando_ExecutarPos(alComando(2), "SELECT NumIntDoc, ValorTotal, NumTitulo, Status  FROM TRPOcorrencias WHERE NumIntDoc= ? ", _
                                    0, lNumIntDocAux, dValorAux, lNumFatAux, iStatusAux, objInfoVoucher.lNumIntDoc)
                        If lErro <> AD_SQL_SUCESSO Then gError 192007
                    
                        'Busca Primeiro
                        lErro = Comando_BuscarPrimeiro(alComando(2))
                        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192008
                        
                        If lErro = AD_SQL_SEM_DADOS Then gError 192018
                    
                        'Trava registro contra alterações/Leituras
                        lErro = Comando_LockExclusive(alComando(2))
                        If lErro <> SUCESSO Then gError 192009
                        
                       If lNumFatAux <> 0 Then
                            sDocAux = objInfoVoucher.sTipoDoc & " " & CStr(objInfoVoucher.lNumero) & SEPARADOR & objInfoVoucher.sTipoVou & " " & objInfoVoucher.sSerie & " " & CStr(objInfoVoucher.lNumVou) & " Fatura: " & CStr(lNumFatAux)
                            gError 200382
                        End If
                        If iStatusAux <> STATUS_TRP_OCR_LIBERADO Then
                            sDocAux = objInfoVoucher.sTipoDoc & " " & CStr(objInfoVoucher.lNumero) & SEPARADOR & objInfoVoucher.sTipoVou & " " & objInfoVoucher.sSerie & " " & CStr(objInfoVoucher.lNumVou) & " Liberado: Não"
                            gError 200382
                        End If
                        If Abs(dValorAux - objInfoVoucher.dValor) > DELTA_VALORMONETARIO Then
                            sDocAux = objInfoVoucher.sTipoDoc & " " & CStr(objInfoVoucher.lNumero) & SEPARADOR & "VOU " & objInfoVoucher.sTipoVou & " " & objInfoVoucher.sSerie & " " & CStr(objInfoVoucher.lNumVou) & " Vlr.Ant: " & Format(objInfoVoucher.dValor, "STANDARD") & " " & " Vlr.Atual: " & Format(dValorAux, "STANDARD")
                            gError 200382
                        End If
                        
                        For Each objOCRAporte In objInfoVoucher.colAportes
                
                            'Acerta o saldo do aporte
                            If objOCRAporte.iMoedaAporte = MOEDA_REAL Then
                                lErro = TRPAportePagtoFat_Acerta_Saldo(alComando(4), alComando(5), alComando(12), alComando(13), objOCRAporte.iTipoPagtoAporte, objOCRAporte.lNumIntDocPagtoAporte, objOCRAporte.dValorAporte)
                            Else
                                lErro = TRPAportePagtoFat_Acerta_Saldo(alComando(4), alComando(5), alComando(12), alComando(13), objOCRAporte.iTipoPagtoAporte, objOCRAporte.lNumIntDocPagtoAporte, objOCRAporte.dValorAporteUSS)
                            End If
                            If lErro <> SUCESSO Then gError 192021
                            
                            lErro = Comando_Executar(alComando(11), "INSERT INTO TRPOcorrenciaAporte( NumIntDocOCR, TipoPagtoAporte, " & _
                                    "NumIntDocPagtoAporte, CodigoAporte, ParcelaAporte, ValorAporte, ValorAporteUSS, MoedaAporte) VALUES (?,?,?,?,?,?,?,?)", _
                                    objOCRAporte.lNumIntDocOCR, objOCRAporte.iTipoPagtoAporte, objOCRAporte.lNumIntDocPagtoAporte, objOCRAporte.lCodigoAporte, objOCRAporte.iParcelaAporte, objOCRAporte.dValorAporte, objOCRAporte.dValorAporteUSS, objOCRAporte.iMoedaAporte)
                            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196476
                            
                        Next
                
                        lErro = Comando_ExecutarPos(alComando(3), "UPDATE TRPOcorrencias SET TipoDocDestino= ?, NumIntDocDestino=?, NumTitulo = ?, ValorAporte = ? ", alComando(2), _
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
            sArquivo = objFaturamento.sDiretorio & gsEmpresaTRP & String(6 - Len(CStr(objInfoFatura.lNumeroFaturaReal)), "0") & CStr(objInfoFatura.lNumeroFaturaReal) & ".html"
            
            If objInfoFatura.dValor >= 0 Then
                'Passa o modelo normal com geração de boleto
                lErro = Gera_Arquivo_Html(objFaturamento, objFaturamento.sModelo, sArquivo, sConteudoHtml)
            Else
                'Passa o modelo de cartão e nota de crédito
                lErro = Gera_Arquivo_Html(objFaturamento, sModeloCP, sArquivo, sConteudoHtml)
            End If
            If lErro <> SUCESSO Then gError 192194
            
            colArquivosHtml.Add sArquivo
            
            Call objFrmAguardeTela.ProcessouItem
            
            If objFrmAguarde.iCancelar = MARCADO Then Exit For
            
        End If

    Next
    
    Set objFrmAguardeTela = Nothing
    If objFrmAguarde.iCancelar = MARCADO Then gError 192084
    
    Set objTRPConfig = New ClassTRPConfig
    
    objTRPConfig.sCodigo = TRPCONFIG_PROX_NUM_TITREC
    objTRPConfig.sConteudo = CStr(lProxCR)
    
    lFatFim = lProxCR - 1
    
    lErro = CF("TRPConfig_Grava_Lockado", alComando(6), alComando(7), objTRPConfig)
    If lErro <> SUCESSO Then gError 192084
        
    TRPFaturamento_Gera_EmTrans = SUCESSO

    Exit Function

Erro_TRPFaturamento_Gera_EmTrans:

    TRPFaturamento_Gera_EmTrans = gErr

    Select Case gErr
    
        Case 192001, 192002, 192021, 192082, 192083, 192084, 192194, 192197 To 192200, 194202, 196240, 196476, 200196, 200226

        Case 192003, 192004
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPVOUCHERS", gErr)
        
        Case 192005
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRPVOUCHERS", gErr)
        
        Case 192006
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRPVOUCHERS", gErr)
            
        Case 192007, 192008
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPOCORRENCIAS", gErr)
        
        Case 192009
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRPOCORRENCIAS", gErr)
        
        Case 192010
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRPOCORRENCIAS", gErr)
            
        Case 192016
            Call Rotina_Erro(vbOKOnly, "ERRO_TRPVOUCHERS_NAO_CADASTRADO2", gErr, objInfoVoucher.lNumIntDoc)

        Case 192018
            Call Rotina_Erro(vbOKOnly, "ERRO_OCORRREMPARCREC_NAO_CADASTRADA", gErr, objInfoVoucher.lNumIntDoc)
            
        Case 196371, 196372, 196376, 196377, 196380
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPVOUCHERINFO", gErr)

        Case 1963713
            Call Rotina_Erro(vbOKOnly, "ERRO_TRPVOUCHERINFO_NAO_CADASTRADO", gErr, objInfoVoucher.lNumIntDoc)

        Case 196378, 196374
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRPVOUCHERINFO", gErr)
        
        Case 196379, 196375
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRPVOUCHERINFO", gErr)

        Case 200382
            Call Rotina_Erro(vbOKOnly, "ERRO_VOU_VALOR_ALTERADO", gErr, sDocAux)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190991)

    End Select
    
    If Not (objFrmAguardeTela Is Nothing) Then
        Call objFrmAguardeTela.Trata_Erro
    End If

    Exit Function

End Function

Private Function TRPFaturamento_Gera_TitRec(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lComando3 As Long, ByVal lComando4 As Long, ByVal lComando5 As Long, ByVal lComando6 As Long, ByVal lComando7 As Long, ByVal lComando8 As Long, ByVal lComando9 As Long, ByVal lComando10 As Long, ByVal objInfoFatura As ClassTRPFATInfoFatura, ByVal objTitRec As ClassTituloReceber, ByVal lProxNum As Long, ByVal colParcRecDataVencReal As Collection) As Long

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
Dim objCtb As New ClassTRPCTBFatCR
Dim objContabAutomatica As New ClassContabAutomatica
Dim dtDataVenctoReal As Date
Dim bAchou As Boolean

On Error GoTo Erro_TRPFaturamento_Gera_TitRec
    
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
    
    lErro = FAT_TituloReceber_Insere_TRP(lComando1, lComando2, lComando3, lComando4, lComando5, lComando6, lComando7, lComando8, lComando9, lComando10, objTitRec, colParcelaReceber)
    If lErro <> SUCESSO Then gError 190993
                        
    'gravar a contabilizacao
    lErro = objContabAutomatica.Inicializa_Contab(objCtb, MODULO_CONTASARECEBER, "CR", objTitRec.dtDataEmissao, objTitRec.dtDataEmissao)
    If lErro <> SUCESSO Then gError 190994

    lErro = objContabAutomatica.GeraContabilizacao(objTitRec, Nothing)
    If lErro <> SUCESSO Then gError 190995

    lErro = objContabAutomatica.Finaliza_Contab()
    If lErro <> SUCESSO Then gError 190996

    objTitRec.colParcelaReceber.AddObj objParcelaReceber

    Set objContabAutomatica = Nothing
    Call objCtb.Encerrando
    Set objCtb = Nothing
                        
    TRPFaturamento_Gera_TitRec = SUCESSO

    Exit Function

Erro_TRPFaturamento_Gera_TitRec:

    TRPFaturamento_Gera_TitRec = gErr

    Select Case gErr
    
        Case 190992 To 190996

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190991)

    End Select

    Exit Function

End Function

Private Function TRPFaturamento_Gera_TitPag(ByVal objInfoFatura As ClassTRPFATInfoFatura, ByVal objTitPag As ClassTituloPagar, ByVal lProxNum As Long, ByVal objCliente As ClassCliente, ByVal objEndereco As ClassEndereco, ByVal colParcRecDataVencReal As Collection) As Long

Dim lErro As Long
Dim objParcelaPagar As New ClassParcelaPagar
Dim colParcelaPagar As New colParcelaPagar
Dim objContabil As ClassContabil
Dim objCtb As New ClassTRPCTBFatCP
Dim objContabAutomatica As New ClassContabAutomatica
Dim dtDataVenctoReal As Date
Dim objInfoVoucher As ClassTRPFATInfoVoucher
Dim objParcelaReceberAux As ClassParcelaReceber
Dim bAchou As Boolean
Dim objForn As New ClassFornecedor

On Error GoTo Erro_TRPFaturamento_Gera_TitPag

    Set objInfoVoucher = objInfoFatura.colInfoVouchers.Item(1)

    If objInfoVoucher.iTipoCliente = TRP_TIPO_CLIENTE_PASSAGEIRO Then
    
        lErro = TRP_Cliente_Grava_Fornecedor(objCliente, objEndereco)
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
    objParcelaPagar.iTipoCobranca = TIPO_COBRANCA_DOC
    objParcelaPagar.iStatus = STATUS_ABERTO

    objForn.lCodigo = objInfoFatura.lCliente
    
    lErro = CF("Fornecedor_Le", objForn)
    If lErro <> SUCESSO And lErro <> 12729 Then gError 190999
    
    If lErro = SUCESSO Then
        If objForn.iBanco = BANCO_REAL Then
            objParcelaPagar.iTipoCobranca = TIPO_COBRANCA_DEP_CONTA
        End If
    End If

    With objParcelaPagar
        colParcelaPagar.Add .lNumIntDoc, .lNumIntTitulo, .iNumParcela, .iStatus, .dtDataVencimento, .dtDataVencimentoReal, .dSaldo, .dValor, .iPortador, .iProxSeqBaixa, .iTipoCobranca, .iBancoCobrador, .sNossoNumero, .sCodigoDeBarras
    End With

    'Grava o Título a Pagar
    lErro = CF("NFFatPag_Grava_EmTrans", objTitPag, colParcelaPagar, objContabil)
    If lErro <> SUCESSO Then gError 190998

    'gravar a contabilizacao
    lErro = objContabAutomatica.Inicializa_Contab(objCtb, MODULO_CONTASAPAGAR, "CP", objTitPag.dtDataEmissao, objTitPag.dtDataEmissao)
    If lErro <> SUCESSO Then gError 190999

    lErro = objContabAutomatica.GeraContabilizacao(objTitPag)
    If lErro <> SUCESSO Then gError 192000

    Set objContabAutomatica = Nothing
    Call objCtb.Encerrando
    Set objCtb = Nothing
                    
    TRPFaturamento_Gera_TitPag = SUCESSO

    Exit Function

Erro_TRPFaturamento_Gera_TitPag:

    TRPFaturamento_Gera_TitPag = gErr

    Select Case gErr
    
        Case 190997 To 190999, 192000, 196904

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190991)

    End Select

    Exit Function

End Function

Private Function TRPFaturamento_Gera_NFPag(ByVal objInfoFatura As ClassTRPFATInfoFatura, ByVal objNFPag As ClassNFsPag, ByVal lProxNum As Long) As Long
'Comissões em geral
Dim lErro As Long

Dim objContabil As ClassContabil
Dim objCtb As New ClassTRPCTBNFPag
Dim objContabAutomatica As New ClassContabAutomatica

On Error GoTo Erro_TRPFaturamento_Gera_NFPag

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

    'gravar a contabilizacao
    lErro = objContabAutomatica.Inicializa_Contab(objCtb, MODULO_CONTASAPAGAR, "CP", objNFPag.dtDataEmissao, objNFPag.dtDataEmissao)
    If lErro <> SUCESSO Then gError 194204

    lErro = objContabAutomatica.GeraContabilizacao(objNFPag, False)
    If lErro <> SUCESSO Then gError 194205

    Set objContabAutomatica = Nothing
    Call objCtb.Encerrando
    Set objCtb = Nothing
                    
    TRPFaturamento_Gera_NFPag = SUCESSO

    Exit Function

Erro_TRPFaturamento_Gera_NFPag:

    TRPFaturamento_Gera_NFPag = gErr

    Select Case gErr
    
        Case 194203 To 194205

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194206)

    End Select

    Exit Function

End Function

Public Function Trata_Exclusao_Do_Destino(ByVal objDestino As Object) As Long

Dim lErro As Long
Dim alComando(0 To 30) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim dValorAporte As Double
Dim lNumIntDocAportePagto As Long
Dim dSaldo As Double
Dim lCodigo As Long
Dim lNumIntDocAporte As Long
Dim objAporte As ClassTRPAportes
Dim iTipoDestino As Integer
Dim lNumIntDocDestino As Long
Dim sClasse As String
Dim bValido As Boolean
Dim objTRPTitulosExp As ClassTRPTitulosExp
Dim lNumeroFaturaReal As Long
Dim sTipVou As String
Dim sSerie As String
Dim lNumVou As Long
Dim iTipoPagtoAporte As Integer
Dim bMotivoObrigatorio As Boolean
Dim objFaturaTRP As ClassFaturaTRP
Dim dValorAporteUSS As Double
Dim iMoedaAporte As Double
Dim objDebitoRecCli As ClassDebitoRecCli
Dim dValorAporteCred As Double
Dim lNumIntDocCredito As Long
Dim objFaturaAporte As ClassTRPFaturaAporte
Dim lNumIntDocVouInfo As Long
Dim objVouExp As ClassTRPVouExp
Dim dtDataCanc As Date
Dim lClienteVou As Long
Dim objTRPTitulos As ClassTRPTitulos

On Error GoTo Erro_Trata_Exclusao_Do_Destino

    bValido = True
    bMotivoObrigatorio = False
   
    sClasse = TypeName(objDestino)
    
    Select Case sClasse
    
        Case TRP_TIPO_DOC_DESTINO_CREDFORN_CLASSE
            iTipoDestino = TRP_TIPO_DOC_DESTINO_CREDFORN
            
        Case TRP_TIPO_DOC_DESTINO_DEBCLI_CLASSE
            iTipoDestino = TRP_TIPO_DOC_DESTINO_DEBCLI
            
        Case TRP_TIPO_DOC_DESTINO_TITPAG_CLASSE
            iTipoDestino = TRP_TIPO_DOC_DESTINO_TITPAG
            lNumeroFaturaReal = objDestino.lNumTitulo
            If objDestino.sSiglaDocumento = TIPODOC_FATURA_SERVICO_CP Or objDestino.sSiglaDocumento = TIPODOC_FATURA_OVER Then
                bMotivoObrigatorio = True
            End If
            
        Case TRP_TIPO_DOC_DESTINO_TITREC_CLASSE
            iTipoDestino = TRP_TIPO_DOC_DESTINO_TITREC
            lNumeroFaturaReal = objDestino.lNumTitulo
            If objDestino.sSiglaDocumento = TIPODOC_FATURA_SERVICO_CR Or objDestino.sSiglaDocumento = TIPODOC_CARTAO_CRED_DEB Then
                bMotivoObrigatorio = True
            End If
            
            'Muda o status das comissões que foram liberadas no faturamento
            'Não muda as já faturadas
            lErro = TRPVouchers_Comissao_Deslibera(TRP_TIPO_LIBERACAO_COMISSAO_FAT, objDestino.lNumIntDoc)
            If lErro <> SUCESSO Then gError 196393
            
        Case TRP_TIPO_DOC_DESTINO_NFSPAG_CLASSE
            iTipoDestino = TRP_TIPO_DOC_DESTINO_NFSPAG
            lNumeroFaturaReal = objDestino.lNumNotaFiscal
            bMotivoObrigatorio = True
            
        Case Else
            bValido = False
    
    End Select
    
    If bValido Then
    
        If bMotivoObrigatorio Then
        
            If objDestino.objInfoUsu Is Nothing Then gError 196490
            If TypeName(objDestino.objInfoUsu) <> "ClassFaturaTRP" Then gError 196490
        
            Set objFaturaTRP = objDestino.objInfoUsu
            
            If Len(Trim(objFaturaTRP.sMotivo)) = 0 Then gError 196491
        
        End If
    
        lNumIntDocDestino = objDestino.lNumIntDoc
        
        Set objTRPTitulosExp = New ClassTRPTitulosExp
        
        objTRPTitulosExp.iExcluido = MARCADO
        objTRPTitulosExp.iExportado = DESMARCADO
        objTRPTitulosExp.iTipoDocOrigem = iTipoDestino
        objTRPTitulosExp.lNumIntDocOrigem = lNumIntDocDestino
        objTRPTitulosExp.lNumTitulo = lNumeroFaturaReal
            
        If bMotivoObrigatorio Then
            
            objTRPTitulosExp.sMotivo = objFaturaTRP.sMotivo
        
            lErro = CF("TRPTitulosExp_Insere", objTRPTitulosExp)
            If lErro <> SUCESSO Then gError 196241
    
        End If
    
        'Abertura de Comando
        For iIndice = LBound(alComando) To UBound(alComando)
            alComando(iIndice) = Comando_Abrir()
            If alComando(iIndice) = 0 Then gError 192022
        Next
        
        sTipVou = String(STRING_TRP_OCR_TIPOVOU, 0)
        sSerie = String(STRING_TRP_OCR_SERIE, 0)

        lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc, TipVou,Serie,NumVou,DataCanc, ClienteVou FROM TRPVouchers WHERE TipoDocDestino= ? AND NumIntDocDestino = ? ", _
                    0, lNumIntDoc, sTipVou, sSerie, lNumVou, dtDataCanc, lClienteVou, iTipoDestino, lNumIntDocDestino)
        If lErro <> AD_SQL_SUCESSO Then gError 192023
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192024
        
        Do While lErro = AD_SQL_SUCESSO
    
            'Trava registro contra alterações/Leituras
            lErro = Comando_LockExclusive(alComando(0))
            If lErro <> SUCESSO Then gError 192025
            
'            Set objVouExp = New ClassTRPVouExp
'            objVouExp.sSerie = sSerie
'            objVouExp.sTipVou = sTipVou
'            objVouExp.lNumVou = lNumVou
'            objVouExp.dtDataCanc = dtDataCanc
'            objVouExp.iTransacao = TRP_EXPORT_VOU_TRANS_CANC_FAT
'            objVouExp.lCliente = lClienteVou
'            lErro = TRPVouExp_Insere(alComando(23), alComando(24), alComando(25), alComando(26), alComando(27), alComando(30), objVouExp)
'            If lErro <> SUCESSO Then gError 200197
        
            lErro = Comando_ExecutarPos(alComando(1), "UPDATE TRPVouchers SET TipoDocDestino= ?, NumIntDocDestino=?, ValorAporte = ?,  NumFat = ? ", alComando(0), 0, 0, 0, 0)
            If lErro <> AD_SQL_SUCESSO Then gError 192026
            
            lErro = Comando_ExecutarPos(alComando(10), "SELECT TipoPagtoAporte, NumIntDocPagtoAporte, ValorAporte,ValorAporteUSS, MoedaAporte FROM TRPVoucherAporte WHERE TipVou= ? AND Serie = ? AND NumVou = ? ", _
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
                        lErro = TRPAportePagtoFat_Acerta_Saldo(alComando(2), alComando(3), alComando(12), alComando(13), iTipoPagtoAporte, lNumIntDocAportePagto, -dValorAporte)
                    Else
                        lErro = TRPAportePagtoFat_Acerta_Saldo(alComando(2), alComando(3), alComando(12), alComando(13), iTipoPagtoAporte, lNumIntDocAportePagto, -dValorAporteUSS)
                    End If
                    If lErro <> SUCESSO Then gError 192027
                    
                End If
                    
                lErro = Comando_ExecutarPos(alComando(11), "DELETE TRPVoucherAporte ", alComando(10))
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
        lErro = Comando_ExecutarPos(alComando(4), "SELECT NumIntDoc, Codigo FROM TRPOcorrencias WHERE TipoDocDestino= ? AND NumIntDocDestino = ? ", _
                    0, lNumIntDoc, lCodigo, iTipoDestino, lNumIntDocDestino)
        If lErro <> AD_SQL_SUCESSO Then gError 192029
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(4))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192030
        
        Do While lErro = AD_SQL_SUCESSO
        
            If iTipoDestino = TRP_TIPO_DOC_DESTINO_CREDFORN Then gError 192034
            If iTipoDestino = TRP_TIPO_DOC_DESTINO_DEBCLI Then gError 192035
    
            'Trava registro contra alterações/Leituras
            lErro = Comando_LockExclusive(alComando(4))
            If lErro <> SUCESSO Then gError 192031
        
            lErro = Comando_ExecutarPos(alComando(5), "UPDATE TRPOcorrencias SET TipoDocDestino= ?, NumIntDocDestino=?, NumTitulo = ?, ValorAporte = ? ", alComando(4), 0, 0, 0, 0)
            If lErro <> AD_SQL_SUCESSO Then gError 192032
            
            lErro = Comando_ExecutarPos(alComando(14), "SELECT TipoPagtoAporte, NumIntDocPagtoAporte, ValorAporte,ValorAporteUSS, MoedaAporte FROM TRPOcorrenciaAporte WHERE NumIntDocOCR= ? ", _
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
                        lErro = TRPAportePagtoFat_Acerta_Saldo(alComando(2), alComando(3), alComando(12), alComando(13), iTipoPagtoAporte, lNumIntDocAportePagto, -dValorAporte)
                    Else
                        lErro = TRPAportePagtoFat_Acerta_Saldo(alComando(2), alComando(3), alComando(12), alComando(13), iTipoPagtoAporte, lNumIntDocAportePagto, -dValorAporteUSS)
                    End If
                    If lErro <> SUCESSO Then gError 192027
                    
                End If
                    
                lErro = Comando_ExecutarPos(alComando(15), "DELETE TRPOcorrenciaAporte ", alComando(14))
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
        lErro = Comando_ExecutarPos(alComando(21), "SELECT NumIntDoc FROM TRPVoucherInfo WHERE TipoDocDestino= ? AND NumIntDocDestino = ? ", _
                    0, lNumIntDocVouInfo, iTipoDestino, lNumIntDocDestino)
        If lErro <> AD_SQL_SUCESSO Then gError 196381
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(21))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196382
        
        Do While lErro = AD_SQL_SUCESSO
    
            'Trava registro contra alterações/Leituras
            lErro = Comando_LockExclusive(alComando(21))
            If lErro <> SUCESSO Then gError 196383

            lErro = Comando_ExecutarPos(alComando(22), "UPDATE TRPVoucherInfo SET TipoDocDestino= ?, NumIntDocDestino=?, NumTitulo = ? ", alComando(21), 0, 0, 0)
            If lErro <> AD_SQL_SUCESSO Then gError 196384
            
            'Busca Próximo
            lErro = Comando_BuscarProximo(alComando(21))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196385
            
        Loop
        
        'Le os aportes diretos
        lErro = Comando_ExecutarPos(alComando(6), "SELECT NumIntDocAporte FROM TRPAportePagtoDireto WHERE TipoDocDestino= ? AND NumIntDocDestino = ? ", _
                    0, lNumIntDocAporte, iTipoDestino, lNumIntDocDestino)
        If lErro <> AD_SQL_SUCESSO Then gError 192036
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(6))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192037
        
        Do While lErro = AD_SQL_SUCESSO
        
            Set objAporte = New ClassTRPAportes
            
            objAporte.lNumIntDoc = lNumIntDocAporte
            
            lErro = CF("TRPAportes_Le", objAporte, False)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 192038
        
            If iTipoDestino = TRP_TIPO_DOC_DESTINO_CREDFORN Then gError 192039
            If iTipoDestino = TRP_TIPO_DOC_DESTINO_DEBCLI Then gError 192040
    
            'Trava registro contra alterações/Leituras
            lErro = Comando_LockExclusive(alComando(6))
            If lErro <> SUCESSO Then gError 192041
        
            lErro = Comando_ExecutarPos(alComando(7), "UPDATE TRPAportePagtoDireto SET TipoDocDestino= ?, NumIntDocDestino=? ", alComando(6), 0, 0)
            If lErro <> AD_SQL_SUCESSO Then gError 192042
            
            'Busca Próximo
            lErro = Comando_BuscarProximo(alComando(6))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192043
            
        Loop
        
        'Le os aportes condicionados
        lErro = Comando_ExecutarPos(alComando(8), "SELECT NumIntDocAporte FROM TRPAportePagtoCond WHERE TipoDocDestino= ? AND NumIntDocDestino = ? ", _
                    0, lNumIntDocAporte, iTipoDestino, lNumIntDocDestino)
        If lErro <> AD_SQL_SUCESSO Then gError 192044
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(8))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192045
        
        Do While lErro = AD_SQL_SUCESSO
        
            Set objAporte = New ClassTRPAportes
            
            objAporte.lNumIntDoc = lNumIntDocAporte
            
            lErro = CF("TRPAportes_Le", objAporte, False)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 192046
        
            If iTipoDestino = TRP_TIPO_DOC_DESTINO_CREDFORN Then gError 192047
            If iTipoDestino = TRP_TIPO_DOC_DESTINO_DEBCLI Then gError 192048
    
            'Trava registro contra alterações/Leituras
            lErro = Comando_LockExclusive(alComando(8))
            If lErro <> SUCESSO Then gError 192049
        
            lErro = Comando_ExecutarPos(alComando(9), "UPDATE TRPAportePagtoCond SET TipoDocDestino= ?, NumIntDocDestino=? ", alComando(8), 0, 0)
            If lErro <> AD_SQL_SUCESSO Then gError 192050
            
            'Busca Próximo
            lErro = Comando_BuscarProximo(alComando(8))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192051
            
        Loop
        
        If iTipoDestino = TRP_TIPO_DOC_DESTINO_TITREC Then
        
            'Le as informações do faturamento
            lErro = Comando_Executar(alComando(16), "SELECT A.ValorAporte, A.NumIntDocCredito FROM TRPTitulosExp AS T, TRPTitulosExpAporte AS A  WHERE T.NumIntDoc = A.NumIntDocTitExp AND T.TipoDocOrigem = ? AND T.NumIntDocOrigem = ? AND T.Excluido = ? AND A.NumIntDocCredito <> ? ", _
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
            lErro = CF("TitulosRecTRP_Exclui", lNumIntDocDestino, objFaturaTRP)
            If lErro <> SUCESSO Then gError 192784
            
        End If
        
        Set objTRPTitulos = New ClassTRPTitulos
        
        objTRPTitulos.iTipoDocOrigem = iTipoDestino
        objTRPTitulos.lNumIntDocOrigem = lNumIntDocDestino
        
        lErro = CF("TRPTitulos_Exclui", objTRPTitulos, alComando(28), alComando(29))
        If lErro <> SUCESSO Then gError 196240
        
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
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPVOUCHERS", gErr)
        
        Case 192025
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRPVOUCHERS", gErr)
        
        Case 192026
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRPVOUCHERS", gErr)
            
        Case 192027, 192038, 192046, 192784, 196241, 200197
            
        Case 192029, 192030, 192033
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPOCORRENCIAS", gErr)
        
        Case 192031
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRPOCORRENCIAS", gErr)
        
        Case 192032
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRPOCORRENCIAS", gErr)
            
        Case 192034
            Call Rotina_Erro(vbOKOnly, "ERRO_CRED_VINC_ORC", gErr, lCodigo)
        
        Case 192035
            Call Rotina_Erro(vbOKOnly, "ERRO_DEB_VINC_OCR", gErr, lCodigo)
                    
        Case 192036, 192037, 192043
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPAPORTEPAGTODIRETO", gErr)
        
        Case 192039
            Call Rotina_Erro(vbOKOnly, "ERRO_CRED_VINC_APORTE", gErr, objAporte.lCodigo)
        
        Case 192040
            Call Rotina_Erro(vbOKOnly, "ERRO_DEB_VINC_APORTE", gErr, objAporte.lCodigo)
        
        Case 192041
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRPAPORTEPAGTODIRETO", gErr)
        
        Case 192042
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRPAPORTEPAGTODIRETO", gErr)
            
        Case 192044, 192045, 192051
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPAPORTEPAGTOCOND", gErr)
        
        Case 192047
            Call Rotina_Erro(vbOKOnly, "ERRO_CRED_VINC_APORTE", gErr, objAporte.lCodigo)
        
        Case 192048
            Call Rotina_Erro(vbOKOnly, "ERRO_DEB_VINC_APORTE", gErr, objAporte.lCodigo)
        
        Case 192049
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRPAPORTEPAGTOCOND", gErr)
        
        Case 192050
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRPAPORTEPAGTOCOND", gErr)
            
        Case 196381, 196382, 196385
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPVOUCHERINFO", gErr)
        
        Case 196383
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRPVOUCHERINFO", gErr)
        
        Case 196384
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRPVOUCHERINFO", gErr)
            
        Case 196393
            
        Case 196490
            Call Rotina_Erro(vbOKOnly, "ERRO_TRP_CANC_FAT_TELA", gErr)
        
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

Private Function TRPAportePagtoFat_Acerta_Saldo(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lComando3 As Long, ByVal lComando4 As Long, ByVal iTipoAporte As Integer, ByVal lNumIntDocAportePagto As Long, ByVal dValorUsado As Double) As Long

Dim lErro As Long
Dim dSaldo As Double
Dim dValor As Double

On Error GoTo Erro_TRPAportePagtoFat_Acerta_Saldo

    If lNumIntDocAportePagto <> 0 And dValorUsado <> 0 Then
    
        If iTipoAporte = TRP_TIPO_APORTE_SOBREFATURA Then
   
            'Acerta o saldo do aporte
            lErro = Comando_ExecutarPos(lComando1, "SELECT Valor, Saldo  FROM TRPAportePagtoFat WHERE NumIntDoc= ? ", 0, dValor, dSaldo, lNumIntDocAportePagto)
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
        
            lErro = Comando_ExecutarPos(lComando2, "UPDATE TRPAportePagtoFat SET Saldo = ? ", lComando1, dSaldo)
            If lErro <> AD_SQL_SUCESSO Then gError 192014
        
        Else
        
            'Acerta o saldo do aporte
            lErro = Comando_ExecutarPos(lComando3, "SELECT Valor FROM TRPAportePagtoFatCond WHERE NumIntDoc= ? ", 0, dValor, lNumIntDocAportePagto)
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
        
            lErro = Comando_ExecutarPos(lComando4, "UPDATE TRPAportePagtoFatCond SET Valor = ? ", lComando3, dValor)
            If lErro <> AD_SQL_SUCESSO Then gError 192014
        
        End If
    
    End If

    TRPAportePagtoFat_Acerta_Saldo = SUCESSO

    Exit Function

Erro_TRPAportePagtoFat_Acerta_Saldo:

    TRPAportePagtoFat_Acerta_Saldo = gErr

    Select Case gErr

        Case 192011, 192012
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPAPORTEPAGTOFAT", gErr)
        
        Case 192013
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRPAPORTEPAGTOFAT", gErr)
        
        Case 192014
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRPAPORTEPAGTOFAT", gErr)

        Case 192015
            Call Rotina_Erro(vbOKOnly, "ERRO_APORTE_SALDO_NEGATIVO", gErr, lNumIntDocAportePagto)
            
        Case 192017
            Call Rotina_Erro(vbOKOnly, "ERRO_TRPAPORTEPAGTOFAT_NAO_CADASTRADO", gErr, lNumIntDocAportePagto)

        Case 192019
            Call Rotina_Erro(vbOKOnly, "ERRO_APORTE_SALDO_MAIOR_VALOR", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192020)

    End Select

    Exit Function

End Function

Private Function TRPAportes_Acerta_ProxParcela(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lCodigoAporte As Long, ByRef iParcela As Integer) As Long

Dim lErro As Long
Dim iProxParcela As Integer

On Error GoTo Erro_TRPAportes_Acerta_ProxParcela

     'Acerta o saldo do aporte
     lErro = Comando_ExecutarPos(lComando1, "SELECT ProxParcela FROM TRPAportes WHERE Codigo= ? ", 0, iParcela, lCodigoAporte)
     If lErro <> AD_SQL_SUCESSO Then gError 196470
 
     'Busca Primeiro
     lErro = Comando_BuscarPrimeiro(lComando1)
     If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196471
     
     If lErro = AD_SQL_SEM_DADOS Then gError 196472
 
     'Trava registro contra alterações/Leituras
     lErro = Comando_LockExclusive(lComando1)
     If lErro <> SUCESSO Then gError 196473
     
     iProxParcela = iParcela + 1
 
     lErro = Comando_ExecutarPos(lComando2, "UPDATE TRPAportes SET ProxParcela = ? ", lComando1, iProxParcela)
     If lErro <> AD_SQL_SUCESSO Then gError 196474

    TRPAportes_Acerta_ProxParcela = SUCESSO

    Exit Function

Erro_TRPAportes_Acerta_ProxParcela:

    TRPAportes_Acerta_ProxParcela = gErr

    Select Case gErr

        Case 196470, 196471
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPAPORTES", gErr)
        
        Case 196472
            Call Rotina_Erro(vbOKOnly, "ERRO_TRPAPORTES_NAO_CADASTRADO", gErr, lCodigoAporte)
        
        Case 196473
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRPAPORTES", gErr)
        
        Case 196474
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRPAPORTES", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196475)

    End Select

    Exit Function

End Function

Function TRPConfig_Grava(ByVal colConfigs As Collection) As Long
'grava na tabela TRPConfig uma colecao de registros de configuracao

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_TRPConfig_Grava
    
    'Abrir transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 192072

    lErro = TRPConfig_Grava_Trans(colConfigs)
    If lErro <> SUCESSO Then gError 192073

    'Fechar transacao
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 192074
          
    TRPConfig_Grava = SUCESSO
    
    Exit Function
    
Erro_TRPConfig_Grava:
    
    TRPConfig_Grava = gErr
    
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

Public Function TRPConfig_Grava_Trans(ByVal colConfigs As Collection) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(1 To 2) As Long
Dim objTRPConfig As ClassTRPConfig
Dim sConteudo As String

On Error GoTo Erro_TRPConfig_Grava_Trans
    
    'abrir comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 192076
    Next
    
   'para cada elemento da colecao
    For Each objTRPConfig In colConfigs
    
        'inicializar buffer
        sConteudo = String(STRING_CONFIG_CONTEUDO, 0)
        
        'ler registo
        lErro = Comando_ExecutarPos(alComando(1), "SELECT Conteudo FROM TRPConfig WHERE Codigo = ? AND FilialEmpresa = ?", 0, sConteudo, objTRPConfig.sCodigo, objTRPConfig.iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 192077
        
        'Lê o primeiro registro
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192078
        
        'Se não encontrou o registro
        If lErro = AD_SQL_SEM_DADOS Then gError 192079
        
        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(1))
        If lErro <> SUCESSO Then gError 192250
            
        lErro = Comando_ExecutarPos(alComando(2), "UPDATE TRPConfig SET Conteudo = ?", alComando(1), objTRPConfig.sConteudo)
        If lErro <> AD_SQL_SUCESSO Then gError 192080

    Next
    
   'fechar comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
          
    TRPConfig_Grava_Trans = SUCESSO
    
    Exit Function
    
Erro_TRPConfig_Grava_Trans:
    
    TRPConfig_Grava_Trans = gErr
    
    Select Case gErr
    
        Case 192076
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 192077, 192078
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPCONFIG", gErr)

        Case 192079
            Call Rotina_Erro(vbOKOnly, "ERRO_TRPCONFIG_INEXISTENTE", gErr, objTRPConfig.sCodigo, objTRPConfig.iFilialEmpresa)
        
        Case 192080
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRPCONFIG", gErr)
            
        Case 192250
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_TRPCONFIG", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192081)
            
    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Public Function TRPConfig_Grava_Lockado(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal objTRPConfig As ClassTRPConfig) As Long

Dim lErro As Long

On Error GoTo Erro_TRPConfig_Grava_Lockado
 
    lErro = Comando_ExecutarPos(lComando2, "UPDATE TRPConfig SET Conteudo = ?", lComando1, objTRPConfig.sConteudo)
    If lErro <> AD_SQL_SUCESSO Then gError 194251
          
    TRPConfig_Grava_Lockado = SUCESSO
    
    Exit Function
    
Erro_TRPConfig_Grava_Lockado:
    
    TRPConfig_Grava_Lockado = gErr
    
    Select Case gErr

        Case 194251
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRPCONFIG", gErr)

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

Function TRPFaturamento_Envia_PorEmail(ByVal objFaturamento As ClassTRPFaturamento, ByVal colParcelas As Collection) As Long

Dim lErro As Long
Dim objParcelaRec As ClassParcelaReceber
Dim objCobrancaSel As New ClassCobrancaPorEmailSel

On Error GoTo Erro_TRPFaturamento_Envia_PorEmail

    For Each objParcelaRec In colParcelas
    
        lErro = CF("ParcelaReceber_Le", objParcelaRec)
        If lErro <> SUCESSO And lErro <> 19147 Then gError 185994

        objCobrancaSel.colParcelas.Add objParcelaRec

    Next
    
    objCobrancaSel.iFaturamento = MARCADO
    objCobrancaSel.lModeloForcado = objFaturamento.lModelo
    
    Call Chama_Tela("CobrancaPorEmail", TIPOTELA_EMAIL_AVISO_COBRANCA, objCobrancaSel)

    TRPFaturamento_Envia_PorEmail = SUCESSO

    Exit Function

Erro_TRPFaturamento_Envia_PorEmail:

    TRPFaturamento_Envia_PorEmail = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192210)

    End Select
    

    Exit Function

End Function

Function TRPFaturamento_Envia_PorEmail_CP(ByVal objFaturamento As ClassTRPFaturamento, ByVal colNFsPag As Collection) As Long

Dim lErro As Long
Dim objNFsPag As ClassNFsPag
Dim objCobrancaSel As New ClassCobrancaPorEmailSel

On Error GoTo Erro_TRPFaturamento_Envia_PorEmail_CP

    For Each objNFsPag In colNFsPag
        objCobrancaSel.colParcelas.Add objNFsPag
    Next
    
    objCobrancaSel.iFaturamento = MARCADO
    objCobrancaSel.lModeloForcado = objFaturamento.lModelo
    
    Call Chama_Tela("CobrancaFaturaPorEmail", TIPOTELA_EMAIL_COBRANCA_FATURA, objCobrancaSel)

    TRPFaturamento_Envia_PorEmail_CP = SUCESSO

    Exit Function

Erro_TRPFaturamento_Envia_PorEmail_CP:

    TRPFaturamento_Envia_PorEmail_CP = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192210)

    End Select
    

    Exit Function

End Function

Public Function TRP_NF_Gera(ByVal objGeracaoNF As ClassTRPGeracaoNF) As Long

Dim lErro As Long
Dim lTransacao As Long
Dim alComando(0 To 55) As Long
Dim iIndice As Integer

On Error GoTo Erro_TRP_NF_Gera

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 192283
    
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 192284
    Next

    lErro = TRP_NF_Gera_EmTrans(alComando, objGeracaoNF)
    If lErro <> SUCESSO Then gError 192285

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
   
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 192286
        
    TRP_NF_Gera = SUCESSO

    Exit Function

Erro_TRP_NF_Gera:

    TRP_NF_Gera = gErr

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

Private Function TRP_NF_Gera_EmTrans(alComando() As Long, ByVal objGeracaoNF As ClassTRPGeracaoNF) As Long

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
Dim objNFTRP As ClassNFTRP
Dim iFilialEmpAnt As Integer
Dim objFilialEmpresa As AdmFiliais
Dim colFiliais As New Collection
Dim bAchou As Boolean
Dim objFilialCliente As ClassFilialCliente
Dim objEndereco As ClassEndereco
Dim iProxComandoAnt As Integer
Dim objFrmAguarde As New ClassFrmAguarde
Dim objFrmAguardeTela As New FrmAguarde

On Error GoTo Erro_TRP_NF_Gera_EmTrans

    iFilialEmpresaAnt = giFilialEmpresa
    
    iProxComando = 0
    
    objFrmAguarde.iTotalItens = objGeracaoNF.colNF.Count
    Call objFrmAguardeTela.Inicializa_Progressao(objFrmAguarde)
    
    For Each objNF In objGeracaoNF.colNF
    
        Set objNFTRP = objNF.objInfoUsu
        
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
        
        Set objNFTRP.objFilialEmpresa = objFilialEmpresa
        
        Set objFilialCliente = New ClassFilialCliente
        Set objEndereco = New ClassEndereco
        
        objFilialCliente.lCodCliente = objNF.lCliente
        objFilialCliente.iCodFilial = objNF.iFilialCli
        
        lErro = CF("FilialCliente_Le", objFilialCliente, alComando(iProxComando))
        If lErro <> SUCESSO And lErro <> 12567 Then gError 196818
                
        objEndereco.lCodigo = objFilialCliente.lEndereco
        
        lErro = CF("Endereco_Le", objEndereco, alComando(iProxComando + 1))
        If lErro <> SUCESSO Then gError 196819
        
        Set objNFTRP.objFilialCliente = objFilialCliente
        Set objNFTRP.objEnderecoCli = objEndereco
        
    Next
    
    iProxComando = iProxComando + 1
    iProxComandoAnt = iProxComando
    
    iIndice = 0
    For Each objNF In objGeracaoNF.colNF
    
        giFilialEmpresa = objNF.iFilialEmpresa

        lNumNotaAnt = objNF.lNumNotaFiscal
        objNF.lNumNotaFiscal = 0
        
        If objNF.iFilialEmpresa = 2 Then
            objNF.iRecibo = MARCADO
        End If
        
        iProxComando = iProxComandoAnt
    
        lErro = NF_NFiscal_Grava_TRP(alComando(), objNF, iProxComando)
        If lErro <> SUCESSO Then gError 192289
        
        If lNumNotaAnt <> objNF.lNumNotaFiscal Then gError 192290
        
        For Each objItemNF In objNF.ColItensNF
        
            iIndice = iIndice + 1
            
            Set objTituloRec = objGeracaoNF.colTitulos.Item(iIndice)
            
            If Abs(objItemNF.dValorTotal - objTituloRec.dValor) > DELTA_VALORMONETARIO Then gError 192299
            
            'Busca os vouchers da fatura para poder calcular a comissão
            lErro = Comando_Executar(alComando(iProxComando + 2), "SELECT X.NumIntDocItemNF, N.NumNotaFiscal FROM TRPItemNotaTitulo AS X, ItensNFiscal AS I, NFiscal AS N WHERE N.NumIntDoc = I.NumIntNF AND N.Status <> 7 AND I.NumIntDoc = X.NumIntDocItemNF AND X.NumIntDocTitulo = ? ", lNumIntDocItemNF, lNumNota, objTituloRec.lNumIntDoc)
            If lErro <> AD_SQL_SUCESSO Then gError 192300
            
            lErro = Comando_BuscarPrimeiro(alComando(iProxComando + 2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192301
            
            If lErro = AD_SQL_SUCESSO Then gError 192302
            
            lErro = Comando_Executar(alComando(iProxComando), "INSERT INTO TRPItemNotaTitulo (NumIntDocItemNF,NumIntDocTitulo) VALUES (?,?)", objItemNF.lNumIntDoc, objTituloRec.lNumIntDoc)
            If lErro <> AD_SQL_SUCESSO Then gError 192291
            
            sSerie = String(STRING_TRP_OCR_SERIE, 0)
            sTipoDoc = String(STRING_TRP_OCR_TIPODOC, 0)

            'Busca as informações de faturamento no Corporator, se não encontrar desce até o voucher para obter a informação da comissão
            lErro = Comando_Executar(alComando(iProxComando + 4), "SELECT ValorComissao FROM TRPTitulosExp WHERE TipoDocOrigem = ? AND NumIntDocOrigem = ? AND Excluido = ? ", dComissao, TRP_TIPO_DOC_DESTINO_TITREC, objTituloRec.lNumIntDoc, DESMARCADO)
            If lErro <> AD_SQL_SUCESSO Then gError 196820

            lErro = Comando_BuscarPrimeiro(alComando(iProxComando + 4))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196821

            If lErro <> AD_SQL_SUCESSO Then

                'Verifica as comissões dadas nos vouchers
                lErro = Comando_Executar(alComando(iProxComando + 1), "SELECT SUM(ValorComissao) FROM TRPVouchers WHERE NumFat = ? ", dComissao, objTituloRec.lNumTitulo)
                If lErro <> AD_SQL_SUCESSO Then gError 192292
    
                lErro = Comando_BuscarPrimeiro(alComando(iProxComando + 1))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192293
        
            End If
        
        Next
       
        'Obter NumIntDoc
        lErro = CF("Config_ObterNumInt2", alComando(iProxComando + 5), alComando(iProxComando + 6), "TRPConfig", "NUM_INT_PROX_TRPNFSEXP", lNumIntDoc)
        If lErro <> SUCESSO Then gError 196822
   
        lErro = Comando_Executar(alComando(iProxComando + 3), "INSERT INTO TRPNFsExp( NumIntDoc, Usuario, Data, Hora, " & _
                "NumIntDocNF, NumNF, TemQueContabilizar,ValorComissao) VALUES (?,?,?,?,?,?,?,?)", _
                lNumIntDoc, gsUsuario, Date, CDbl(Time), objNF.lNumIntDoc, objNF.lNumNotaFiscal, MARCADO, dComissao)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196823
    
        Call objFrmAguardeTela.ProcessouItem
        
        If objFrmAguarde.iCancelar = MARCADO Then Exit For

    Next
    
    Set objFrmAguardeTela = Nothing
    If objFrmAguarde.iCancelar = MARCADO Then gError 196784
    
    giFilialEmpresa = iFilialEmpresaAnt

    TRP_NF_Gera_EmTrans = SUCESSO

    Exit Function

Erro_TRP_NF_Gera_EmTrans:

    giFilialEmpresa = iFilialEmpresaAnt

    TRP_NF_Gera_EmTrans = gErr

    Select Case gErr
    
        Case 192289, 192294, 192296 To 192298, 196784
        
        Case 192290
            Call Rotina_Erro(vbOKOnly, "ERRO_NUM_NOTA_DIFERENTE", gErr, lNumNotaAnt, objNF.lNumNotaFiscal)
        
        Case 192291 'INSERT
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRPITEMNOTATITULO", gErr)
        
        Case 192292, 192293
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPVOUCHERS", gErr)
        
        Case 192299
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_ITEM_ERRADO", gErr, Format(objItemNF.dValorTotal, "STANDARD"), Format(objTituloRec.dValor, "STANDARD"))

        Case 192300, 192301
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPITEMNOTATITULO", gErr)
            
        Case 192302
            Call Rotina_Erro(vbOKOnly, "ERRO_TITULO_JA_COM_NOTA", gErr, objTituloRec.lNumTitulo, lNumNota)

        Case 196817 To 196819, 196822
        
        Case 196820, 196821
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPTITULOSEXP", gErr)
            
        Case 196823
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRPNFSEXP", gErr)

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
Dim alComando(1 To 16) As Long
Dim iIndice As Integer
Dim lNumVou As Long
Dim sSerie As String
Dim sTipoVou As String
Dim objOcorrencia As ClassTRPOcorrencias
Dim objOcorrenciaDet As ClassTRPOcorrenciaDet
Dim dValor As Double
Dim lCodigo As Long
Dim objVou As ClassTRPVouchers
Dim iCartao As Integer
Dim lCliente As Long
Dim dValorComissao As Double
Dim lNumIntComi As Long
Dim lNumIntDocVouInfo As Long
Dim lNumIntDocVouInfoAux As Long
Dim dtDataVou As Date
Dim sConteudo As String
Dim lNumIntDocNVL As Long
Dim iStatus As Integer
Dim objVouExp As ClassTRPVouExp
Dim dtDataCanc As Date
Dim lClienteVou As Long
Dim lCliPassageiro As Long

On Error GoTo Erro_BaixasParcelaRec_Grava_Customizado

     'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 192308
    Next

    sSerie = String(STRING_TRP_OCR_SERIE, 0)
    sTipoVou = String(STRING_TRP_OCR_TIPOVOU, 0)

    'Busca os vouchers cancelados para poder gerar as ocorrências de reembolso
    lErro = Comando_ExecutarPos(alComando(1), "SELECT NumVou, Serie, Valor, TipVou, Cartao,Status,NumIntDocNVL, DataCanc, ClienteVou, CliPassageiro FROM TRPVouchers WHERE TipoDocDestino = ? AND NumIntDocDestino = ? ", 0, lNumVou, sSerie, dValor, sTipoVou, iCartao, iStatus, lNumIntDocNVL, dtDataCanc, lClienteVou, lCliPassageiro, TRP_TIPO_DOC_DESTINO_TITREC, objTituloRec.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 192309
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192310

    'Enquanto existirem vouchers
    Do While lErro <> AD_SQL_SEM_DADOS
    
        'Se o voucher foi cancelado  e ainda não teve estorno - > estorna
        If iStatus = STATUS_CANCELADO And lNumIntDocNVL = 0 Then
        
            Set objOcorrencia = New ClassTRPOcorrencias
            Set objOcorrenciaDet = New ClassTRPOcorrenciaDet
            
            objOcorrencia.dtDataEmissao = gdtDataAtual
            objOcorrencia.dValorTotal = -dValor
            objOcorrencia.iFilialCliente = objTituloRec.iFilial
            objOcorrencia.iFormaPagto = FORMAPAGTO_TRP_OCR_FAT
            objOcorrencia.iOrigem = INATIVACAO_AUTOMATICA_CODIGO
            objOcorrencia.iStatus = STATUS_TRP_OCR_BLOQUEADO
            
            lCliente = objTituloRec.lCliente
            
            'Tem que pegar o passageiro
            If iCartao = MARCADO Then
            
                Set objVou = New ClassTRPVouchers
            
                objVou.lNumVou = lNumVou
                objVou.sSerie = sSerie
                objVou.sTipVou = sTipoVou
        
                'VERIFICAR SE O PASSAGEIRO ESTÁ PREENCHIDO
                If lCliPassageiro = 0 Then
                
                    'CADASTRAR O PASSAGEIRO E ACERTA A INFORMAÇÃO EM VOUCHERINFO
                    lErro = TRP_Voucher_Cadastra_Cliente_EmTrans(objVou)
                    If lErro <> SUCESSO Then gError 192826
                    
                Else
                    objVou.lCliPassageiro = lCliente
                End If
                
                lCliente = objVou.lCliPassageiro
            
            End If
                               
            objOcorrencia.lCliente = lCliente
            
            lErro = CF("Config_ObterAutomatico_EmTrans", "FATConfig", "NUM_PROX_TRPOCORRENCIAS", "TRPOcorrencias", "Codigo", lCodigo)
            If lErro <> SUCESSO Then gError 192311
    
            objOcorrencia.lCodigo = lCodigo
            objOcorrencia.lNumVou = lNumVou
            objOcorrencia.sHistorico = "Reembolso com base no pagto do título " & CStr(objTituloRec.lNumTitulo) & " filial empresa " & CStr(objTituloRec.iFilialEmpresa) & " e tipo de documento " & objTituloRec.sSiglaDocumento
            objOcorrencia.sSerie = sSerie
            objOcorrencia.sTipoDoc = sTipoVou
            objOcorrencia.sObservacao = ""
            
            objOcorrenciaDet.iTipo = INATIVACAO_AUTOMATICA_TIPO_NVL_CODIGO
            objOcorrenciaDet.iSeq = 1
            objOcorrenciaDet.dValor = -dValor
            
            objOcorrencia.colDetalhamento.Add objOcorrenciaDet
            
            lErro = TRPOcorrencias_Grava_EmTrans(objOcorrencia)
            If lErro <> SUCESSO Then gError 192312
            
            lErro = Comando_ExecutarPos(alComando(2), "UPDATE TRPVouchers SET NumIntDocNVL = ?", alComando(1), objOcorrencia.lNumIntDoc)
            If lErro <> AD_SQL_SUCESSO Then gError 192314
            
        End If
    
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192313
        
    Loop
    
    'Busca o histórico relacionado ao voucher para liberá-lo para o comissionamento
    lErro = Comando_Executar(alComando(8), "SELECT I.NumIntDoc, V.Data FROM TRPVouchers AS V, TRPVoucherInfo AS I WHERE V.NumVou = I.NumVou AND V.Serie = I.Serie AND V.TipVou = I.TipVou AND V.TipoDocDestino = ? AND V.NumIntDocDestino = ? AND I.Status = ? ", _
                                             lNumIntDocVouInfo, dtDataVou, TRP_TIPO_DOC_DESTINO_TITREC, objTituloRec.lNumIntDoc, TRP_VOU_INFO_STATUS_BLOQUEADO)
    If lErro <> AD_SQL_SUCESSO Then gError 196360
    
    lErro = Comando_BuscarPrimeiro(alComando(8))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196361

    'Enquanto existirem registros
    Do While lErro <> AD_SQL_SEM_DADOS
    
        'Busca o histórico relacionado ao voucher para liberá-lo para o comissionamento
        lErro = Comando_ExecutarPos(alComando(9), "SELECT NumIntDoc FROM TRPVoucherInfo WHERE NumIntDoc = ? ", 0, lNumIntDocVouInfoAux, lNumIntDocVouInfo)
        If lErro <> AD_SQL_SUCESSO Then gError 196362
        
        lErro = Comando_BuscarPrimeiro(alComando(9))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196363

'        'Trava registro contra alterações/Leituras
'        lErro = Comando_LockExclusive(alComando(9))
'        If lErro <> SUCESSO Then gError 196364
    
        lErro = Comando_ExecutarPos(alComando(10), "UPDATE TRPVoucherInfo SET Status = ?, TipoLiberacao = ?, NumIntDocLiberacao = ? ", alComando(9), TRP_VOU_INFO_STATUS_LIBERADO, TRP_TIPO_LIBERACAO_COMISSAO_BAIXA, objBaixaParcRec.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 196365
        
        lErro = Comando_BuscarProximo(alComando(8))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196366
        
    Loop
    
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
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPVOUCHERS", gErr)
            
        Case 192311, 192312, 192825, 192826, 200198
        
        Case 192314
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRPVOUCHERS", gErr)
            
        Case 196270, 196271
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOSRECTRP", gErr)
        
        Case 196272, 196273
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFSPAG", gErr)
        
        Case 196274, 196275
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFSPAGBAIXADAS", gErr)
            
        Case 196276, 196277, 192370

        Case 196278, 196279
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COINFOCOMISSAO", gErr)
            
        Case 196360 To 196363, 196366
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPVOUCHERINFO", gErr)

        Case 196364
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_TRPVOUCHERINFO", gErr)

        Case 196365
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRPVOUCHERINFO", gErr)

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
Dim objAporte As New ClassTRPAportes
Dim objAportePagtoCond As New ClassTRPAportePagtoCond
Dim objAportePagtoDireto As New ClassTRPAportePagtoDireto
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
        
            objAportePagtoDireto.iTipoDocDestino = TRP_TIPO_DOC_DESTINO_TITPAG
            objAportePagtoDireto.lNumIntDocDestino = objTituloPagar.lNumIntDoc

            lErro = CF("TRPAportePagtoDireto_Le_NumIntDocDestino", objAportePagtoDireto)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 192377
            
            lNumIntDocPagto = objAportePagtoDireto.lNumIntDoc
            
            If lErro = ERRO_LEITURA_SEM_DADOS Then
            
                objAportePagtoCond.iTipoDocDestino = TRP_TIPO_DOC_DESTINO_TITPAG
                objAportePagtoCond.lNumIntDocDestino = objTituloPagar.lNumIntDoc
            
                lErro = CF("TRPAportePagtoCond_Le_NumIntDocDestino", objAportePagtoCond)
                If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 192378
                
                lNumIntDocPagto = objAportePagtoCond.lNumIntDoc
                
            End If
            
            'Título a pagar já vinculado a outro aporte
            If lErro = SUCESSO Then
                If lNumIntDocPagto <> objAporte.lNumIntDocPagto Then gError 192379
            End If
            If objAporte.iTipoPagto = FORMAPAGTO_TRP_APORTE_TIPOPAGTO_DIRETO Then
                
                objAportePagtoDireto.lNumIntDoc = objAporte.lNumIntDocPagto
                                
                lErro = CF("TRPAportePagtoDireto_Le_NumIntDoc", objAportePagtoDireto)
                If lErro <> SUCESSO Then gError 192371
                
                If objAportePagtoDireto.lNumIntDocDestino <> 0 And objAportePagtoDireto.lNumIntDocDestino <> objTituloPagar.lNumIntDoc Then gError 192380
                                
                objAportePagtoDireto.iTipoDocDestino = TRP_TIPO_DOC_DESTINO_TITPAG
                objAportePagtoDireto.lNumIntDocDestino = objTituloPagar.lNumIntDoc
                                
                lErro = TRPAportePagtoDireto_Atualiza(alComando(0), alComando(1), objAportePagtoDireto, objAporte)
                If lErro <> SUCESSO Then gError 192372
            Else
                objAportePagtoCond.lNumIntDoc = objAporte.lNumIntDocPagto
            
                lErro = CF("TRPAportePagtoCond_Le_NumIntDoc", alComando(2), objAportePagtoCond)
                If lErro <> SUCESSO Then gError 192373
            
                If objAportePagtoCond.lNumIntDocDestino <> 0 And objAportePagtoCond.lNumIntDocDestino <> objTituloPagar.lNumIntDoc Then gError 192381
                
                objAportePagtoCond.iTipoDocDestino = TRP_TIPO_DOC_DESTINO_TITPAG
                objAportePagtoCond.lNumIntDocDestino = objTituloPagar.lNumIntDoc
            
                lErro = TRPAportePagtoCond_Atualiza(alComando(0), alComando(1), objAportePagtoCond, objAporte)
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

Public Function TRP_Voucher_Cadastra_Cliente(ByVal objVou As ClassTRPVouchers) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_TRP_Voucher_Cadastra_Cliente

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 192740

    lErro = TRP_Voucher_Cadastra_Cliente_EmTrans(objVou)
    If lErro <> SUCESSO Then gError 192741

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 192742

    TRP_Voucher_Cadastra_Cliente = SUCESSO

    Exit Function

Erro_TRP_Voucher_Cadastra_Cliente:

    TRP_Voucher_Cadastra_Cliente = gErr

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

Public Function TRP_Voucher_Cadastra_Cliente_EmTrans(ByVal objVou As ClassTRPVouchers) As Long

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
Dim objClienteTRP As ClassClienteTRP
Dim iFilialEmpresa As Integer
Dim sNomePax As String
Dim sRazaoPax As String
Dim sCPFPax As String

On Error GoTo Erro_TRP_Voucher_Cadastra_Cliente_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 192744
    Next
    
    lErro = CF("TRPVouchers_Le", objVou)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 200270
    
    sNumero = ""
    If objVou.iCartao = DESMARCADO Then
        Call Formata_String_Numero(objVou.sPassageiroCGC, sNumero)
        sNomePax = objVou.sPassageiroNome
        sRazaoPax = objVou.sPassageiroNome & " " & objVou.sPassageiroSobreNome
    Else
        Call Formata_String_Numero(objVou.sTitularCPF, sNumero)
        sNomePax = objVou.sTitular
        sRazaoPax = objVou.sTitular
    End If
    If sCPFPax = "00000000000" Then sCPFPax = ""
    sCPFPax = sNumero
    
    lErro = Comando_Executar(alComando(0), "SELECT F.CodCliente, F.CodFilial, C.Tipo FROM FiliaisClientes AS F, Clientes AS C WHERE F.CodCliente = C.Codigo AND F.CGC = ? AND F.CGC <> ?", lCliente, iFilial, iTipo, sCPFPax, "")
    If lErro <> AD_SQL_SUCESSO Then gError 192745

    lErro = Comando_BuscarProximo(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192746
    
    If lErro <> AD_SQL_SUCESSO Then
    
        lErro = CF("Config_ObterAutomatico_EmTrans", "TRPConfig", TRPCONFIG_PROX_NUM_PASSAGEIRO, "Clientes", "Codigo", lCodigo)
        If lErro <> SUCESSO Then gError 192747
    
        If lCodigo < FAIXA_TRP_TIPO_CLIENTE_PASSAGEIRO_DE Or lCodigo > FAIXA_TRP_TIPO_CLIENTE_PASSAGEIRO_ATE Then gError 200298 'Não está na faixa reservada para passageiros
    
        For iIndice = 1 To 3
            objVou.objEnderecoPax.lCodigo = 0
            colEndereco.Add objVou.objEnderecoPax
        Next
        
        objCliente.lCodigo = lCodigo
        objCliente.iAtivo = MARCADO

        objCliente.sCgc = sCPFPax
        
        objCliente.iTipo = TRP_TIPO_CLIENTE_PASSAGEIRO
        objCliente.dtDataPrimeiraCompra = DATA_NULA
        objCliente.dtDataUltChequeDevolvido = DATA_NULA
        objCliente.dtDataUltimaCompra = DATA_NULA
        objCliente.dtDataUltimoProtesto = DATA_NULA
        objCliente.dtDataUltVisita = DATA_NULA
        objCliente.iCondicaoPagto = 1
            
        Set objFilialCliCategoria = New ClassFilialCliCategoria
        Set objClienteTRP = New ClassClienteTRP
        
        objFilialCliCategoria.sCategoria = TRP_CATEGORIA_CONDFAT
        objFilialCliCategoria.sItem = TRP_CATEGORIA_CONDFAT_PASSAGEIRO
        objFilialCliCategoria.iFilial = FILIAL_MATRIZ
        objFilialCliCategoria.lCliente = objCliente.lCodigo
        
        objCliente.colCategoriaItem.Add objFilialCliCategoria
        
        Set objCliente.objInfoUsu = objClienteTRP
                
        lErro = Comando_Executar(alComando(3), "SELECT FilialEmpresa FROM FiliaisClientesTRP WHERE CodCliente = ? AND CodFilial = ?", iFilialEmpresa, objVou.lClienteVou, FILIAL_MATRIZ)
        If lErro <> AD_SQL_SUCESSO Then gError 192745
    
        lErro = Comando_BuscarProximo(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192746
        
        If lErro <> AD_SQL_SEM_DADOS Then
        
            objClienteTRP.iFilialEmpresa = iFilialEmpresa
            
        End If
        
        objClienteTRP.lEmpresaPai = objCliente.lCodigo
               
    Else
    
        If iTipo <> TRP_TIPO_CLIENTE_PASSAGEIRO Then gError 196543
    
        objCliente.lCodigo = lCliente
        
        lErro = CF("Cliente_Le", objCliente)
        If lErro <> SUCESSO And lErro <> 12293 Then gError 192748
        
        lErro = CF("Enderecos_Le_Cliente", colEnderecos, objCliente)
        If lErro <> SUCESSO Then gError 192749

        For Each objEndereco In colEnderecos
            objEndereco.sBairro = objVou.objEnderecoPax.sBairro
            objEndereco.sCEP = objVou.objEnderecoPax.sCEP
            objEndereco.sCidade = objVou.objEnderecoPax.sCidade
            objEndereco.sEmail = objVou.objEnderecoPax.sEmail
            objEndereco.sEndereco = objVou.objEnderecoPax.sEndereco
            objEndereco.sFax = objVou.objEnderecoPax.sFax
            objEndereco.sSiglaEstado = objVou.objEnderecoPax.sSiglaEstado
            objEndereco.sTelefone1 = objVou.objEnderecoPax.sTelefone1
            objEndereco.sTelefone2 = objVou.objEnderecoPax.sTelefone2
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
    
    If Len(Trim(objVou.sPassageiroNome)) = 0 Then gError 192829
    
    objCliente.sNomeReduzido = left(sNomePax, STRING_CLIENTE_NOME_REDUZIDO - Len(SEPARADOR & CStr(objCliente.lCodigo))) & SEPARADOR & CStr(objCliente.lCodigo)
    objCliente.sRazaoSocial = sRazaoPax
       
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
    
    objVou.lCliPassageiro = objCliente.lCodigo
   
    lErro = Comando_ExecutarPos(alComando(4), "SELECT NumVou FROM TRPVouchers WHERE NumVou = ? AND Serie = ? AND TipVou = ? ", _
                0, lNumVou, objVou.lNumVou, objVou.sSerie, objVou.sTipVou)
    If lErro <> AD_SQL_SUCESSO Then gError 192755

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(4))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192756

    If lErro <> AD_SQL_SUCESSO Then gError 192757

    'Trava registro contra alterações/Leituras
    lErro = Comando_LockExclusive(alComando(4))
    If lErro <> AD_SQL_SUCESSO Then gError 192758

    lErro = Comando_ExecutarPos(alComando(5), "UPDATE TRPVouchers SET CliPassageiro= ? ", alComando(4), objCliente.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 192759

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TRP_Voucher_Cadastra_Cliente_EmTrans = SUCESSO

    Exit Function

Erro_TRP_Voucher_Cadastra_Cliente_EmTrans:

    TRP_Voucher_Cadastra_Cliente_EmTrans = gErr

    Select Case gErr

        Case 192744
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 192745, 192746
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLIENTES", gErr)
            
        Case 192747 To 192754
        
        Case 192755, 192756
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPVOUCHERS", gErr)
            
        Case 192757
            Call Rotina_Erro(vbOKOnly, "ERRO_VOUCHER_NAO_CADASTRADO", gErr, objVou.lNumVou, objVou.sSerie, objVou.sTipVou)

        Case 192758
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRPVOUCHERS", gErr)

        Case 192759
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRPVOUCHERS", gErr)
            
        Case 192828
            Call Rotina_Erro(vbOKOnly, "ERRO_ESTADO_NAO_CADASTRADO", gErr, objEndereco.sSiglaEstado)

        Case 192829
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_REDUZIDO_NAO_PREENCHIDO", gErr)
            
        Case 196543
            Call Rotina_Erro(vbOKOnly, "ERRO_CNPJ_PERTENCE_A_UMA_AGENGIA", gErr, objVou.sPassageiroCGC, objVou.lNumVou, lCliente)

        Case 200298
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_FORA_FAIXA_PAX", gErr, FAIXA_TRP_TIPO_CLIENTE_PASSAGEIRO_DE, FAIXA_TRP_TIPO_CLIENTE_PASSAGEIRO_ATE)

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
Public Function TitulosRecTRP_Grava(ByVal objTitulosRecTRP As ClassTitulosRecTRP) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Integer

On Error GoTo Erro_TitulosRecTRP_Grava

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 192776
    Next

    'Le a tabelaTitulosRecTRP
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDocTitRec FROM TitulosRecTRP WHERE NumIntDocTitRec= ? ", _
                0, lNumIntDoc, objTitulosRecTRP.lNumIntDocTitRec)
    If lErro <> AD_SQL_SUCESSO Then gError 192777

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192778

    'Se existir => UPDATE, senão => INSERT
    If lErro = AD_SQL_SUCESSO Then

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 192779

        lErro = Comando_ExecutarPos(alComando(1), "UPDATE TitulosRecTRP SET ValorComissao= ? ", alComando(0), _
                objTitulosRecTRP.dValorComissao)
        If lErro <> AD_SQL_SUCESSO Then gError 192780

    Else


        lErro = Comando_Executar(alComando(1), "INSERT INTO TitulosRecTRP( NumIntDocTitRec, ValorTarifa,  ValorDeducoes, " & _
                "ValorComissao, ValorBruto) VALUES (?,?,?,?,?)", _
                objTitulosRecTRP.lNumIntDocTitRec, objTitulosRecTRP.dValorTarifa, objTitulosRecTRP.dValorDeducoes, _
                objTitulosRecTRP.dValorComissao, objTitulosRecTRP.dValorBruto)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192781

    End If

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TitulosRecTRP_Grava = SUCESSO

    Exit Function

Erro_TitulosRecTRP_Grava:

    TitulosRecTRP_Grava = gErr

    Select Case gErr

        Case 192776
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 192777, 192778
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOSRECTRP", gErr)

        Case 192779
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TITULOSRECTRP", gErr)

        Case 192780
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TITULOSRECTRP", gErr)

        Case 192781
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TITULOSRECTRP", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192775)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function TitulosRecTRP_Exclui(ByVal lNumIntDoc As Long, ByVal objFaturaTRP As ClassFaturaTRP) As Long

Dim lErro As Long
Dim alComando(0 To 2) As Long
Dim iIndice As Integer
Dim lNumIntDocAux As Long
Dim objNFPag As New ClassNFsPag
Dim objContabil As New ClassContabil
Dim lNumNF As Long

On Error GoTo Erro_TitulosRecTRP_Exclui

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 192767
    Next

    'Le a tabelaTitulosRecTRP
    lErro = Comando_Executar(alComando(0), "SELECT N.NumNotaFiscal FROM TRPItemNotaTitulo AS T, NFiscal AS N, ItensNFiscal AS I WHERE T.NumIntDocTitulo = ? AND T.NumIntDocItemNF = I.NumIntDoc AND I.NumIntNF = N.NumIntDoc AND N.Status NOT IN (5,7) ", _
                lNumNF, lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 192768

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192769
    
    'O título está associado a NF
    If lErro = AD_SQL_SUCESSO Then gError 196455
    
    'Le a tabelaTitulosRecTRP
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDocTitRec FROM TitulosRecTRP WHERE NumIntDocTitRec= ? ", _
                0, lNumIntDocAux, lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 192768

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192769

    If lErro <> AD_SQL_SEM_DADOS Then

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 192771
    
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM TitulosRecTRP", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 192772
        
    End If

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TitulosRecTRP_Exclui = SUCESSO

    Exit Function

Erro_TitulosRecTRP_Exclui:

    TitulosRecTRP_Exclui = gErr

    Select Case gErr

        Case 192767
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 192768, 192769
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOSRECTRP", gErr)

        Case 192771
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TITULOSRECTRP", gErr)

        Case 192772
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TITULOSRECTRP", gErr)
        
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

Public Function TRPFaturas_Regera_Html(ByVal lFaturaDe As Long, ByVal lFaturaAte As Long, ByVal sModelo As String, ByVal sDiretorio As String, ByVal sSiglaDoc As String, Optional ByVal bAcompanhamento As Boolean = True) As Long

Dim lErro As Long
Dim alComando(0 To 7) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim objFaturamento As ClassTRPFaturamento
Dim objTitRec As ClassTituloReceber
Dim objInfoFatura As ClassTRPFATInfoFatura
Dim objInfoVoucher As ClassTRPFATInfoVoucher
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
Dim objVoucher As ClassTRPVouchers
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
Dim objVouAporte As ClassTRPVoucherAporte
Dim objOCRAporte As ClassTRPOcorrenciaAporte
Dim iTipoPagtoAporte As Integer
Dim lNumIntDocOCR As Long
Dim dValorAporteCred As Double
Dim lNumIntDocCredito As Long
Dim lCodigoAporteCred As Long
Dim objFaturaAporte As ClassTRPFaturaAporte
Dim iNumParcelas As Integer
Dim iPagtoParcial As Integer
Dim lCount As Long
Dim objFrmAguarde As New ClassFrmAguarde
Dim objFrmAguardeTela As New FrmAguarde
Dim objFilialForn As ClassFilialFornecedor

On Error GoTo Erro_TRPFaturas_Regera_Html

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 192830
    Next
    
    'Le a tabela TitulosRecTRP
    lErro = Comando_Executar(alComando(6), "SELECT COUNT(NumIntDoc) FROM TRPFaturas WHERE NumTitulo BETWEEN ? AND ? AND (SiglaDocumento = ? OR ? = '') AND Status <> 5 ", _
                lCount, lFaturaDe, lFaturaAte, sSiglaDoc, sSiglaDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 192831

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(6))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192832
    
    If bAcompanhamento Then
        objFrmAguarde.iTotalItens = lCount
        Call objFrmAguardeTela.Inicializa_Progressao(objFrmAguarde)
    End If
    
    lErro = CF("TRPConfig_Le", TRPCONFIG_DIRETORIO_MODELO_FAT_HTML_CARTAO, EMPRESA_TODA, sModeloCP)
    If lErro <> SUCESSO Then gError 192838
    
    sSiglaDocumento = String(STRING_SIGLA_DOCUMENTO, 0)

    'Le a tabela TitulosRecTRP
    lErro = Comando_Executar(alComando(0), "SELECT NumIntDoc,TipoDocDestino, NumTitulo,FilialEmpresa, SiglaDocumento, Cliente, Valor, DataEmissao, DataVencimento, CondicaoPagto FROM TRPFaturas WHERE NumTitulo BETWEEN ? AND ? AND (SiglaDocumento = ? OR ? = '') AND Status <> 5 ORDER BY SiglaDocumento ASC", _
                lNumIntDoc, iTipoDocDestino, lNumFat, iFilialEmpresa, sSiglaDocumento, lCliente, dValor, dtDataEmissao, dtDataVencimento, iCondicaoPagto, lFaturaDe, lFaturaAte, sSiglaDoc, sSiglaDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 192831

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192832

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objFaturamento = New ClassTRPFaturamento
        Set objTitRec = New ClassTituloReceber
        Set objInfoFatura = New ClassTRPFATInfoFatura
        objFaturamento.colInfoFaturas.Add objInfoFatura
        
        objFaturamento.iIndiceFaturaAtual = 1
        
        Select Case iTipoDocDestino
        
            Case TRP_TIPO_DOC_DESTINO_TITREC
        
                'Le a tabela TRPTitulosEXP
                lErro = Comando_Executar(alComando(2), "SELECT A.ValorAporte, A.NumIntDocCredito, A.CodigoAporte, A.ParcelaAporte, A.NumParcelas, A.PagtoParcial FROM TRPTitulosExp AS E, TRPTitulosExpAporte AS A WHERE E.NumIntDoc = A.NumIntDocTitExp AND E.TipoDocOrigem = ? AND E.NumIntDocOrigem = ? AND E.Excluido = ? ", dValorAporteCred, lNumIntDocCredito, lCodigoAporteCred, iParcelaAporte, iNumParcelas, iPagtoParcial, iTipoDocDestino, lNumIntDoc, DESMARCADO)
                If lErro <> AD_SQL_SUCESSO Then gError 192834

                'Busca Primeiro
                lErro = Comando_BuscarPrimeiro(alComando(2))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192835
                                   
                Do While lErro = AD_SQL_SUCESSO
                
                    Set objFaturaAporte = New ClassTRPFaturaAporte
                
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
                
                'Le a tabela TitulosRecTRP
                lErro = Comando_Executar(alComando(3), "SELECT ValorTarifa, ValorDeducoes FROM TitulosRecTRP WHERE NumIntDocTitRec = ? ", dValorTarifa, dValorDeducoes, lNumIntDoc)
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
                
            Case TRP_TIPO_DOC_DESTINO_TITPAG, TRP_TIPO_DOC_DESTINO_NFSPAG
            
                dValor = -dValor
                
                If sSiglaDocumento = TIPODOC_FATURA_OVER Then
                    objInfoFatura.iTipoDoc = TRP_TIPODOC_FAT_OVER
                End If
                
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
        
        sSerie = String(STRING_TRP_OCR_SERIE, 0)
        sTipVou = String(STRING_TRP_OCR_TIPOVOU, 0)
        sTipoDoc = String(STRING_TRP_OCR_TIPODOC, 0)
        sNomeCliente = String(STRING_CLIENTE_NOME_REDUZIDO, 0)
        
        'Le a tabelaTitulosRecTRP
        lErro = Comando_Executar(alComando(1), "SELECT NumIntDoc, Serie, TipVou, TipoDoc,NomeCliente, NumVou, ValorAporte, ValorComissao, Valor, Data FROM DocFaturados WHERE NumTitulo = ? ", _
                    lNumIntDocOCR, sSerie, sTipVou, sTipoDoc, sNomeCliente, lNumVou, dValorAporte, dValorComissao, dValorItem, dtDataItem, lNumFat)
        If lErro <> AD_SQL_SUCESSO Then gError 192836
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192837

        iSeq = 0
        Do While lErro <> AD_SQL_SEM_DADOS
        
            Set objInfoVoucher = New ClassTRPFATInfoVoucher
            objFaturamento.colInfoVouchers.Add objInfoVoucher
            Set objVoucher = New ClassTRPVouchers
            
            iSeq = iSeq + 1
            
            objVoucher.sSerie = sSerie
            objVoucher.sTipVou = sTipVou
            objVoucher.lNumVou = lNumVou
            
            lErro = CF("TRPVouchers_Le", objVoucher)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 192838
            
            objInfoVoucher.dCambio = objVoucher.dCambio
            objInfoVoucher.iFilialEmpresa = iFilialEmpresa
            objInfoVoucher.iLinha = iSeq
            objInfoVoucher.iMarcado = MARCADO
            objInfoVoucher.iPax = objVoucher.iPax
            objInfoVoucher.lCliente = lCliente
            objInfoVoucher.sNome = sNomeCliente
            objInfoVoucher.lNumIntDoc = lNumIntDocOCR
            
            If sTipoDoc = TRP_TIPODOC_VOU_TEXTO Then
                objInfoVoucher.dValorCambio = objVoucher.dValorCambio
                objInfoVoucher.lCodigoAporte = lCodigoAporte
                objInfoVoucher.lNumIntDocAportePagto = lNumIntDocAportePagto
                objInfoVoucher.dValorComissao = objVoucher.dValorComissao
                objInfoVoucher.dValorBruto = objVoucher.dValorBruto
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
            objInfoVoucher.sControle = objVoucher.sControle
            objInfoVoucher.iMoeda = objVoucher.iMoeda
            objInfoVoucher.sProduto = objVoucher.sProduto
            objInfoVoucher.sSerie = sSerie
            objInfoVoucher.sTipoDoc = sTipoDoc
            objInfoVoucher.sTipoVou = sTipVou
            
            If objInfoVoucher.sTipoDoc = TRP_TIPODOC_CMC_TEXTO Or objInfoVoucher.sTipoDoc = TRP_TIPODOC_CMR_TEXTO Then
                objInfoVoucher.sControle = ""
            End If

            If objInfoVoucher.sTipoDoc = TRP_TIPODOC_NVL_TEXTO Then
                objInfoVoucher.dValorAporte = dValorAporte
                objInfoVoucher.lCodigoAporte = lCodigoAporte
                objInfoVoucher.iParcelaAporte = iParcelaAporte
                objInfoVoucher.lNumIntDocAportePagto = lNumIntDocAportePagto
            End If
            
            objInfoVoucher.lCliVou = objInfoVoucher.lCliente
            objInfoVoucher.sNomeCliVou = objInfoVoucher.sNome
                                       
            objInfoVoucher.sTitular = objVoucher.sPassageiroNome & " " & objVoucher.sPassageiroSobreNome
                            
            If objVoucher.lEmissor <> 0 Then

                Set objForn = New ClassFornecedor

                objForn.lCodigo = objVoucher.lEmissor
                
                lErro = CF("Fornecedor_Le", objForn)
                If lErro <> SUCESSO And lErro <> 12729 Then gError 194210
                
                objInfoVoucher.sEmissor = objForn.sRazaoSocial
            
            End If
        
            If objInfoVoucher.sTipoDoc = TRP_TIPODOC_CMC_TEXTO Or objInfoVoucher.sTipoDoc = TRP_TIPODOC_OVER_TEXTO Or objInfoVoucher.sTipoDoc = TRP_TIPODOC_CMR_TEXTO Or (objVoucher.iCartao = MARCADO And objInfoVoucher.sTipoDoc = TRP_TIPODOC_VOU_TEXTO) Or ((objInfoVoucher.sTipoDoc = TRP_TIPODOC_NVL_TEXTO Or objInfoVoucher.sTipoDoc = TRP_TIPODOC_OCR_TEXTO) And objInfoVoucher.lCliente <> objVoucher.lClienteVou) Then
                
                If objVoucher.lClienteVou <> 0 Then
                
                    Set objCliente = New ClassCliente
        
                    objCliente.lCodigo = objVoucher.lClienteVou
            
                    lErro = CF("Cliente_Le", objCliente)
                    If lErro <> SUCESSO And lErro <> 12293 Then gError 192840
                    
                    objInfoVoucher.lCliVou = objVoucher.lClienteVou
                    objInfoVoucher.sNomeCliVou = objCliente.sRazaoSocial
                    
                    If objInfoVoucher.sTipoDoc = TRP_TIPODOC_CMC_TEXTO Or objInfoVoucher.sTipoDoc = TRP_TIPODOC_CMR_TEXTO Or ((objInfoVoucher.sTipoDoc = TRP_TIPODOC_NVL_TEXTO Or objInfoVoucher.sTipoDoc = TRP_TIPODOC_OCR_TEXTO) And objInfoVoucher.lCliente <> objVoucher.lClienteVou) Then
                        'Se for comissão deve sair o cliente do voucher como titular
                        objInfoVoucher.sTitular = objCliente.sNomeReduzido
                    End If
                    
                End If
                
            End If
            
            'Se for voucher pago com cartão o titular é o titular do cartão
            If objVoucher.iCartao = MARCADO And objInfoVoucher.sTipoDoc = TRP_TIPODOC_VOU_TEXTO Then
                objInfoVoucher.sTitular = objVoucher.sTitular
            End If
            
            objInfoFatura.iParcelaAporte = iParcelaAporte

            objInfoFatura.colInfoVouchers.Add objInfoVoucher
            
            If sTipoDoc = TRP_TIPODOC_VOU_TEXTO Then
                       
                lErro = Comando_Executar(alComando(4), "SELECT TipoPagtoAporte, NumIntDocPagtoAporte, ValorAporte, ParcelaAporte, CodigoAporte FROM TRPVoucherAporte WHERE TipVou= ? AND Serie = ? AND NumVou = ? ", _
                         iTipoPagtoAporte, lNumIntDocAportePagto, dValorAporte, iParcelaAporte, lCodigoAporte, objInfoVoucher.sTipoVou, objInfoVoucher.sSerie, objInfoVoucher.lNumVou)
                If lErro <> AD_SQL_SUCESSO Then gError 196486
        
                'Busca Primeiro
                lErro = Comando_BuscarPrimeiro(alComando(4))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196487
                
                Do While lErro = AD_SQL_SUCESSO
            
                    Set objVouAporte = New ClassTRPVoucherAporte
                    
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
      
            ElseIf sTipoDoc = TRP_TIPODOC_NVL_TEXTO Or sTipoDoc = TRP_TIPODOC_OCR_TEXTO Then
            
                lErro = Comando_ExecutarPos(alComando(5), "SELECT TipoPagtoAporte, NumIntDocPagtoAporte, ValorAporte, ParcelaAporte, CodigoAporte FROM TRPOcorrenciaAporte WHERE NumIntDocOCR = ? ", _
                        0, iTipoPagtoAporte, lNumIntDocAportePagto, dValorAporte, iParcelaAporte, lCodigoAporte, objInfoVoucher.lNumIntDoc)
                If lErro <> AD_SQL_SUCESSO Then gError 196489
        
                'Busca Primeiro
                lErro = Comando_BuscarPrimeiro(alComando(5))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196490
                
                Do While lErro = AD_SQL_SUCESSO
            
                    Set objOCRAporte = New ClassTRPOcorrenciaAporte
                    
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
        
        Call objInfoFatura.Agrupa_Itens
        
        objInfoFatura.iLinha = iLinha
        
        For Each objInfoVoucher In objInfoFatura.colInfoVouchersAG
        
            If objInfoVoucher.iMarcado = MARCADO Then
                dValorItens = dValorItens + objInfoVoucher.dValor - objInfoVoucher.dValorAporte
                dValorComissao = dValorComissao + objInfoVoucher.dValorComissao
                dValorAporte = dValorAporte + objInfoVoucher.dValorAporte
                
                'Se for negativo é um crédito
                If objInfoVoucher.dValor < 0 Then
                    dValorCreditos = dValorCreditos + objInfoVoucher.dValor
                End If
                'Se for um documento fora voucher que acrescebta no valor é um débito
                If objInfoVoucher.dValor > 0 And objInfoVoucher.sTipoDoc <> TRP_TIPODOC_VOU_TEXTO Then
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
        
        objInfoFatura.dValor = dValor
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
        
        If lErro = SUCESSO Then
        
            lErro = CF("FilialCliente_Le", objFilialCliente, alComando(7))
            If lErro <> SUCESSO And lErro <> 12567 Then gError 192841
            
        Else
            
            Set objForn = New ClassFornecedor
            Set objFilialForn = New ClassFilialFornecedor
            
            objForn.lCodigo = objInfoFatura.lCliente
            objFilialForn.lCodFornecedor = objInfoFatura.lCliente
            objFilialForn.iCodFilial = FILIAL_MATRIZ
            
            lErro = CF("Fornecedor_Le", objForn)
            If lErro <> SUCESSO And lErro <> 12729 Then gError 192841

            lErro = CF("FilialFornecedor_Le", objFilialForn)
            If lErro <> SUCESSO And lErro <> 12929 Then gError 192841
            
            objCliente.sNomeReduzido = objForn.sNomeReduzido
            objCliente.sRazaoSocial = objForn.sRazaoSocial
            
            objFilialCliente.lEndereco = objFilialForn.lEndereco
            objFilialCliente.sCgc = objFilialForn.sCgc
            objFilialCliente.sInscricaoEstadual = objFilialForn.sInscricaoEstadual
            objFilialCliente.sInscricaoMunicipal = objFilialForn.sInscricaoMunicipal
            objFilialCliente.sNome = objFilialForn.sNome

        End If

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
        
        sArquivo = sDiretorio & gsEmpresaTRP & String(6 - Len(CStr(objInfoFatura.lNumeroFaturaReal)), "0") & CStr(objInfoFatura.lNumeroFaturaReal) & ".html"
        
        If sSiglaDocumento = TIPODOC_CARTAO_CRED_DEB Or objInfoFatura.dValor < 0 Then
            lErro = Gera_Arquivo_Html(objFaturamento, sModeloCP, sArquivo, sConteudoHtml)
        Else
            lErro = Gera_Arquivo_Html(objFaturamento, sModelo, sArquivo, sConteudoHtml)
        End If
        If lErro <> SUCESSO Then gError 192844
    
        'Busca Proximo
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192845
    
        If bAcompanhamento Then
            Call objFrmAguardeTela.ProcessouItem
            
            If objFrmAguarde.iCancelar = MARCADO Then Exit Do
        End If

    Loop
    
    If bAcompanhamento Then
        Set objFrmAguardeTela = Nothing
        If objFrmAguarde.iCancelar = MARCADO Then gError 192844
    End If

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TRPFaturas_Regera_Html = SUCESSO

    Exit Function

Erro_TRPFaturas_Regera_Html:

    TRPFaturas_Regera_Html = gErr

    Select Case gErr

        Case 192830
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 192831, 192832, 192845
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOSREC", gErr)
            
        Case 192833, 192838, 192840 To 192844, 194207

        Case 192834, 192835
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PARCELASREC", gErr)
            
        Case 192836, 192837, 192839
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPVOUCHERS", gErr)
            
        Case 196486 To 196488
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPVOUCHERAPORTE", gErr)
        
        Case 196489 To 196491
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPOCORRENCIAAPORTE", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192845)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    If Not (objFrmAguardeTela Is Nothing) Then
        Call objFrmAguardeTela.Trata_Erro
    End If

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

Public Function DocFaturadosLista_BotaoConsulta(ByVal objVoucher As ClassTRPVouchers) As Long

Dim sTela As String
Dim objObjeto As Object

On Error GoTo Erro_DocFaturadosLista_BotaoConsulta

    Select Case objVoucher.iTipoDocDestino
    
        Case TRP_TIPO_DOC_DESTINO_CREDFORN
            sTela = TRP_TIPO_DOC_DESTINO_CREDFORN_TELA
            Set objObjeto = New ClassCreditoPagar
            
        Case TRP_TIPO_DOC_DESTINO_DEBCLI
            sTela = TRP_TIPO_DOC_DESTINO_DEBCLI_TELA
            Set objObjeto = New ClassDebitoRecCli
    
        Case TRP_TIPO_DOC_DESTINO_TITPAG
            sTela = TRP_TIPO_DOC_DESTINO_TITPAG_TELA
            Set objObjeto = New ClassTituloPagar
    
        Case TRP_TIPO_DOC_DESTINO_TITREC
            sTela = TRP_TIPO_DOC_DESTINO_TITREC_TELA
            Set objObjeto = New ClassTituloReceber
    
        Case TRP_TIPO_DOC_DESTINO_NFSPAG
            sTela = TRP_TIPO_DOC_DESTINO_NFSPAG_TELA
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

Public Function TRPAPortesPagtoFatHistResLista_BotaoConsulta(ByVal objVoucher As ClassTRPVouchers) As Long

Dim colSelecao As New Collection

On Error GoTo Erro_TRPAPortesPagtoFatHistResLista_BotaoConsulta

    colSelecao.Add objVoucher.lNumIntDocPagtoAporteFat
    colSelecao.Add objVoucher.iTipoPagtoAporte
    colSelecao.Add objVoucher.iTipoDocDestino
    colSelecao.Add objVoucher.lNumIntDocDestino
     
    Call Chama_Tela("TRPAPortesPagtoFatHistLista", colSelecao, Nothing, Nothing, "NumIntDocPagtoAporteFat = ? AND TipoPagtoAporte = ? AND TipoDocDestino = ? AND NumIntDocDestino = ?")

    TRPAPortesPagtoFatHistResLista_BotaoConsulta = SUCESSO

    Exit Function

Erro_TRPAPortesPagtoFatHistResLista_BotaoConsulta:

    TRPAPortesPagtoFatHistResLista_BotaoConsulta = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192879)

    End Select

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
    
        'Le a tabelaTRPOcorrencias
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

Public Function TRPTitulosExp_Insere(ByVal objTRPTitulosExp As ClassTRPTitulosExp, Optional lComandoP1 As Long = 0, Optional lComandoP2 As Long = 0) As Long

Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim objFaturaAporte As ClassTRPFaturaAporte

On Error GoTo Erro_TRPTitulosExp_Insere

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
    lErro = CF("Config_ObterNumInt", "TRPConfig", "NUM_INT_PROX_TRPTITULOSEXP", lNumIntDoc)
    If lErro <> SUCESSO Then gError 196237

    objTRPTitulosExp.lNumIntDoc = lNumIntDoc

    lErro = Comando_Executar(alComando(1), "INSERT INTO TRPTitulosExp( NumIntDoc, Usuario, Data, Hora, " & _
            "TipoDocOrigem, NumIntDocOrigem, NumTitulo, Excluido, Exportado, Motivo, TemQueContabilizar," & _
            "ValorAporteCred,ValorAporte,ValorComissao,ValorCreditos,ValorDebitos,ValorTarifa,ValorDeducoes) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
            objTRPTitulosExp.lNumIntDoc, gsUsuario, Date, CDbl(Time), _
            objTRPTitulosExp.iTipoDocOrigem, objTRPTitulosExp.lNumIntDocOrigem, objTRPTitulosExp.lNumTitulo, objTRPTitulosExp.iExcluido, objTRPTitulosExp.iExportado, objTRPTitulosExp.sMotivo, objTRPTitulosExp.iTemQueContabilizar, _
            objTRPTitulosExp.dValorAporteCred, objTRPTitulosExp.dValorAporte, objTRPTitulosExp.dValorComissao, objTRPTitulosExp.dValorCreditos, objTRPTitulosExp.dValorDebitos, _
            objTRPTitulosExp.dValorTarifa, objTRPTitulosExp.dValorDeducoes)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196238
    
    For Each objFaturaAporte In objTRPTitulosExp.colFatAporteCred
    
        lErro = Comando_Executar(alComando(2), "INSERT INTO TRPTitulosExpAporte( NumIntDocTitExp,CodigoAporte,ParcelaAporte,NumParcelas,ValorAporte,NumIntDocCredito,PagtoParcial) VALUES (?,?,?,?,?,?,?)", _
                objTRPTitulosExp.lNumIntDoc, objFaturaAporte.lCodigoAporte, objFaturaAporte.iParcelaAporte, objFaturaAporte.iNumParcelas, objFaturaAporte.dValorAporte, objFaturaAporte.lNumIntDocCredito, objFaturaAporte.iPagtoParcial)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196238
    
    Next

    If lComandoP1 = 0 Or lComandoP2 = 0 Then
        'Fecha Comando
        For iIndice = LBound(alComando) To UBound(alComando)
            Call Comando_Fechar(alComando(iIndice))
        Next
    End If

    TRPTitulosExp_Insere = SUCESSO

    Exit Function

Erro_TRPTitulosExp_Insere:

    TRPTitulosExp_Insere = gErr

    Select Case gErr

        Case 196236
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 196237

        Case 196238
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRPTITULOSEXP", gErr)

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

Public Function TRPTitulos_Insere(ByVal objTRPTitulos As ClassTRPTitulos, Optional lComandoP1 As Long = 0) As Long

Dim lErro As Long
Dim alComando(1 To 1) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long

On Error GoTo Erro_TRPTitulos_Insere

    If lComandoP1 <> 0 Then
        alComando(1) = lComandoP1
    Else
        'Abertura de Comando
        For iIndice = LBound(alComando) To UBound(alComando)
            alComando(iIndice) = Comando_Abrir()
            If alComando(iIndice) = 0 Then gError 200228
        Next
    End If

    'Obter NumIntDoc
    lErro = CF("Config_ObterNumInt", "TRPConfig", "NUM_INT_PROX_TRPTITULOS", lNumIntDoc)
    If lErro <> SUCESSO Then gError 200229

    objTRPTitulos.lNumIntDoc = lNumIntDoc

    lErro = Comando_Executar(alComando(1), "INSERT INTO TRPTitulos( NumIntDoc, TipoDocOrigem, NumIntDocOrigem, TipoDoc) VALUES (?,?,?,?)", _
            objTRPTitulos.lNumIntDoc, objTRPTitulos.iTipoDocOrigem, objTRPTitulos.lNumIntDocOrigem, objTRPTitulos.sTipoDoc)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200230
    

    If lComandoP1 = 0 Then
        'Fecha Comando
        For iIndice = LBound(alComando) To UBound(alComando)
            Call Comando_Fechar(alComando(iIndice))
        Next
    End If

    TRPTitulos_Insere = SUCESSO

    Exit Function

Erro_TRPTitulos_Insere:

    TRPTitulos_Insere = gErr

    Select Case gErr

        Case 200228
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 200229

        Case 200230
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRPTITULOS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200231)

    End Select

    If lComandoP1 = 0 Then
        'Fecha Comando
        For iIndice = LBound(alComando) To UBound(alComando)
            Call Comando_Fechar(alComando(iIndice))
        Next
    End If

    Exit Function

End Function

Public Function Cliente_Grava_Customizado(ByVal objCliente As ClassCliente) As Long

Dim lErro As Long
Dim alComando(1 To 9) As Long
Dim iIndice As Integer
Dim lCliente As Long
Dim objClienteTRP As ClassClienteTRP
Dim objCliEmissoresTRP As ClassCliEmissoresTRP
Dim lNumIntDoc As Long
Dim lNumIntDoc1 As Long
Dim iIndice1 As Integer
Dim objCliEmissoresExcTRP As ClassCliEmissoresExcTRP
Dim objTRPCliRep As ClassTRPClienteRepProd
Dim objTRPCliCor As ClassTRPClienteCorProd
Dim objExcComisCli As ClassTRPExcComissaoCli
Dim colClientesContatos As New Collection
Dim colClientesContatosAux As New Collection
Dim objClienteContato As New ClassClienteContatos
Dim objClienteContatoAux As ClassClienteContatos
Dim bAchou As Boolean
Dim objEndereco As ClassEndereco
Dim colEnderecos As New colEndereco
Dim objClienteBD As New ClassCliente
Dim objForn As New ClassFornecedor
Dim objEnderecoForn As New ClassEndereco

On Error GoTo Erro_Cliente_Grava_Customizado

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 195811
    Next

    lErro = Comando_Executar(alComando(1), "DELETE FROM TRPCliEmissoresExc WHERE NumIntDocCliEmi IN (Select NumIntDoc FROM TRPCliEmissores WHERE Cliente = ?)", objCliente.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 195812

    'Não vai excluir nem incluir só alterar o percentual de comissão no cadastro do cliente, as demais informações do emissor
    'será em tela especícica para isso

'    lErro = Comando_Executar(alComando(2), "DELETE FROM TRPCliEmissores WHERE Cliente = ?", objCliente.lCodigo)
'    If lErro <> AD_SQL_SUCESSO Then gError 195813
            
    iIndice = 0

    If Not (objCliente.objInfoUsu Is Nothing) Then
    
        Set objClienteTRP = objCliente.objInfoUsu
        
        If objClienteTRP.colCliEmissoresTRP.Count > 0 Then
        
'            lErro = CF("Config_ObterNumInt", "TRPConfig", TRPCONFIG_NUM_INT_PROX_TRPCLIEMISSORES, lNumIntDoc, objClienteTRP.colCliEmissoresTRP.Count)
'            If lErro <> SUCESSO Then gError 195814

            For Each objCliEmissoresTRP In objClienteTRP.colCliEmissoresTRP
            
'                iIndice = iIndice + 1
'
'                objCliEmissoresTRP.lNumIntDoc = lNumIntDoc
'                objCliEmissoresTRP.iSeq = iIndice
'
'                lNumIntDoc = lNumIntDoc + 1
    
                lErro = Comando_Executar(alComando(3), "UPDATE TRPCliEmissores SET PercComissao = ? WHERE FornEmissor = ? ", _
                         objCliEmissoresTRP.dPercComissao, objCliEmissoresTRP.lFornEmissor)
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 195815
    
                If objCliEmissoresTRP.colCliEmissoresExcTRP.Count > 0 Then
    
                    lErro = CF("Config_ObterNumInt", "TRPConfig", TRPCONFIG_NUM_INT_PROX_TRPCLIEMISSORESEXC, lNumIntDoc1, objCliEmissoresTRP.colCliEmissoresExcTRP.Count)
                    If lErro <> SUCESSO Then gError 195816

                    iIndice1 = 0
                    
                    For Each objCliEmissoresExcTRP In objCliEmissoresTRP.colCliEmissoresExcTRP

                        iIndice1 = iIndice1 + 1

                        objCliEmissoresExcTRP.lNumIntDoc = lNumIntDoc1
                        objCliEmissoresExcTRP.lNumIntDocCliEmi = objCliEmissoresTRP.lNumIntDoc
                        objCliEmissoresExcTRP.iSeq = iIndice1

                        lNumIntDoc1 = lNumIntDoc1 + 1
                        
                        lErro = Comando_Executar(alComando(4), "INSERT INTO TRPCliEmissoresExc (NumIntDoc, NumIntDocCliEmi, Seq, Produto, PercComissao) VALUES (?,?,?,?,?)", _
                                 objCliEmissoresExcTRP.lNumIntDoc, objCliEmissoresExcTRP.lNumIntDocCliEmi, objCliEmissoresExcTRP.iSeq, objCliEmissoresExcTRP.sProduto, objCliEmissoresExcTRP.dPercComissao)
                        If lErro <> AD_SQL_SUCESSO Then gError 195817
                        
                    Next
    
                End If
    
            Next
    
        End If
    
        lErro = Comando_Executar(alComando(5), "DELETE FROM FiliaisClientesTRP WHERE CodCliente = ?  AND CodFilial = ?", objCliente.lCodigo, FILIAL_MATRIZ)
        If lErro <> AD_SQL_SUCESSO Then gError 195859
    
        lErro = Comando_Executar(alComando(6), "INSERT INTO FiliaisClientesTRP (CodCliente, CodFilial, Representante, ComissaoRep, Correntista, ComissaoCorr, ComissaoAg, Aporte, EmpresaPai, FilialNF,  FilialEmpresa, FilialFat, CondPagtoCC) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)", objCliente.lCodigo, FILIAL_MATRIZ, _
        objClienteTRP.lCodRepresentante, objClienteTRP.dPercComiRep, objClienteTRP.lCodCorrentista, objClienteTRP.dPercComiCorr, objClienteTRP.dPercComiAg, objClienteTRP.iConsiderarAporte, objClienteTRP.lEmpresaPai, objClienteTRP.iFilialNF, objClienteTRP.iFilialEmpresa, objClienteTRP.iFilialFat, objClienteTRP.iCondPagtoCC)
        If lErro <> AD_SQL_SUCESSO Then gError 195860
        
        lErro = TRPClienteCorProd_Exclui(objCliente)
        If lErro <> SUCESSO Then gError 196955
        
        lErro = TRPClienteRepProd_Exclui(objCliente)
        If lErro <> SUCESSO Then gError 196956
        
        If objCliente.lCodigo = objClienteTRP.lEmpresaPai Then
            lErro = TRPExcComissaoCli_Exclui(objCliente)
            If lErro <> SUCESSO Then gError 196956
        End If
                
        For Each objTRPCliCor In objClienteTRP.colExcCor
            objTRPCliCor.lCliente = objCliente.lCodigo
            lErro = TRPClienteCorProd_Insere(alComando(7), objTRPCliCor)
            If lErro <> SUCESSO Then gError 196957
        Next
        
        For Each objTRPCliRep In objClienteTRP.colExcRep
            objTRPCliRep.lCliente = objCliente.lCodigo
            lErro = TRPClienteRepProd_Insere(alComando(8), objTRPCliRep)
            If lErro <> SUCESSO Then gError 196958
        Next
        
        If objCliente.lCodigo = objClienteTRP.lEmpresaPai Then
            For Each objExcComisCli In objClienteTRP.colExcComissaoCli
                objExcComisCli.lCliente = objCliente.lCodigo
                lErro = TRPExcComissaoCli_Insere(alComando(9), objExcComisCli)
                If lErro <> SUCESSO Then gError 196957
            Next
        End If
    
    End If
        
    objClienteBD.lCodigo = objCliente.lCodigo
    
    lErro = CF("Cliente_Le", objClienteBD)
    If lErro <> SUCESSO And lErro <> 12293 Then gError 196957
    
    objClienteContato.lCliente = objCliente.lCodigo
    objClienteContato.iFilialCliente = FILIAL_MATRIZ
    
    'Le os contatos já cadastrados
    lErro = CF("ClienteContatos_Le_Todos", colClientesContatosAux, objClienteContato)
    If lErro <> SUCESSO And lErro <> 102574 Then gError 196957
    
    bAchou = False
    For Each objClienteContatoAux In colClientesContatosAux
        For Each objClienteContato In colClientesContatos
            If Compara_Telefone(objClienteContato.sTelefone, objClienteContatoAux.sTelefone) And Compara_Contato(objClienteContato.sContato, objClienteContatoAux.sContato) Then
                bAchou = True
                If objClienteContato.sContato = "Sem nome" And objClienteContatoAux.sContato <> "Sem nome" Then
                    objClienteContato.sContato = objClienteContatoAux.sContato
                    objClienteContato.sCargo = objClienteContatoAux.sCargo
                    objClienteContato.sEmail = objClienteContatoAux.sEmail
                    objClienteContato.sFax = objClienteContatoAux.sFax
                    objClienteContato.sSetor = objClienteContatoAux.sSetor
                End If
                Exit For
            End If
        Next
        If Not bAchou Then
            objClienteContatoAux.iPadrao = DESMARCADO
            colClientesContatos.Add objClienteContatoAux
        End If
    Next
    
    'ACERTA OS CONTATOS
    
    'Lê os dados dos tres tipos de enderecos
    lErro = CF("Enderecos_Le_Cliente", colEnderecos, objClienteBD)
    If lErro <> SUCESSO Then gError 196957
    
    For Each objEndereco In colEnderecos
        lErro = CF("Endereco_Le_Customizado", objEndereco)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 196957
    Next
    
    bAchou = False
    iIndice = 0
    For Each objEndereco In colEnderecos
        iIndice = iIndice + 1
        For Each objClienteContato In colClientesContatos
            If Compara_Telefone(objClienteContato.sTelefone, objEndereco.sTelefone1) And Compara_Contato(objClienteContato.sContato, objEndereco.sContato) Then
                bAchou = True
                If iIndice = 1 Then
                    objClienteContato.iPadrao = MARCADO
                End If
                Exit For
            End If
        Next
        If Not bAchou Then
        
            If Len(Trim(objEndereco.sTelefone1)) > 0 Then
            
                Set objClienteContato = New ClassClienteContatos
                
                If Len(Trim(objEndereco.sContato)) = 0 Then
                    objClienteContato.sContato = "Sem nome"
                Else
                    objClienteContato.sContato = objEndereco.sContato
                End If
                objClienteContato.sTelefone = objEndereco.sTelefone1
                objClienteContato.sEmail = objEndereco.sEmail
                objClienteContato.sFax = objEndereco.sFax
                
                If iIndice = 1 Then
                    objClienteContato.iPadrao = MARCADO
                End If
                
                colClientesContatos.Add objClienteContato
                
            End If
            
        End If
    Next
    
    bAchou = False
    For Each objEndereco In colEnderecos
        For Each objClienteContato In colClientesContatos
            If Compara_Telefone(objClienteContato.sTelefone, objEndereco.sTelefone2) And Compara_Contato(objClienteContato.sContato, objEndereco.sContato) Then
                bAchou = True
                Exit For
            End If
        Next
        If Not bAchou Then
        
            If Len(Trim(objEndereco.sTelefone2)) > 0 Then
            
                Set objClienteContato = New ClassClienteContatos
                
                If Len(Trim(objEndereco.sContato)) = 0 Then
                    objClienteContato.sContato = "Sem nome"
                Else
                    objClienteContato.sContato = objEndereco.sContato
                End If
                objClienteContato.sTelefone = objEndereco.sTelefone2
                objClienteContato.sEmail = objEndereco.objInfoUsu.sEmail2
                objClienteContato.sFax = objEndereco.sFax
                
                colClientesContatos.Add objClienteContato
            
            End If
            
        End If
    Next
    
    iIndice = 0
    For Each objClienteContato In colClientesContatos
        iIndice = iIndice + 1
        objClienteContato.iCodigo = iIndice
        objClienteContato.lCliente = objCliente.lCodigo
        objClienteContato.iFilialCliente = FILIAL_MATRIZ
    Next
    
    If colClientesContatos.Count > 0 Then
        lErro = CF("ClienteContatos_Grava_EmTrans", colClientesContatos)
        If lErro <> SUCESSO Then gError 196957
    End If

    'Grava as informações na tabela de fornecedores também
    objForn.lCodigo = objCliente.lCodigo
    
    'faz a leitura para carregar info já existente que nao tenha vindo na importacao
    lErro = CF("Fornecedor_Le", objForn)
    If lErro <> SUCESSO And lErro <> 12729 Then gError 200303
    
    With objEnderecoForn
        .sEndereco = colEnderecos(1).sEndereco
        .sBairro = colEnderecos(1).sBairro
        .sCidade = colEnderecos(1).sCidade
        .sSiglaEstado = colEnderecos(1).sSiglaEstado
        .sCEP = colEnderecos(1).sCEP
        .sTelefone1 = colEnderecos(1).sTelefone1
        .sTelefone2 = colEnderecos(1).sTelefone2
        .sFax = colEnderecos(1).sFax
        .sContato = colEnderecos(1).sContato
        .sEmail = colEnderecos(1).sEmail
        .iCodigoPais = colEnderecos(1).iCodigoPais
    End With
            
    With objForn
        .sNomeReduzido = objCliente.sNomeReduzido
        .sRazaoSocial = objCliente.sRazaoSocial
        .sCgc = objCliente.sCgc
        .sInscricaoEstadual = objCliente.sInscricaoEstadual
        .iTipo = objCliente.iTipo
        .iAtivo = objCliente.iAtivo
        
        If Not (objCliente.objInfoUsu Is Nothing) Then
            .iBanco = objClienteTRP.iBanco
            .sAgencia = objClienteTRP.sAgencia
            .sContaCorrente = objClienteTRP.sContaCorrente
            .iCondicaoPagto = objClienteTRP.iCondPagtoCC
        End If
    
        .sInscricaoMunicipal = objCliente.sInscricaoMunicipal
    End With
    
    lErro = CF("Fornecedor_Grava_EmTrans", objForn, objEnderecoForn)
    If lErro <> SUCESSO Then gError 200304

    lErro = CF("TRPUsuarios_Grava", objClienteTRP.objUsuarioWeb)
    If lErro <> SUCESSO Then gError 200304

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
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRPCLIEMISSORESEXC", gErr)

        Case 195813
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRPCLIEMISSORES", gErr)
        
        Case 195814, 195816, 196955 To 196958, 200303, 200304
        
        Case 195815
            Call Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_TRPCLIEMISSORES", gErr)

        Case 195817
            Call Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_TRPCLIEMISSORESEXC", gErr)

        Case 195859
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_FILIAISCLIENTESTRP", gErr)

        Case 195859
            Call Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_FILIAISCLIENTESTRP", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195818)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function Cliente_Grava_Customizado2(ByVal objCliente As ClassCliente, ByVal bClienteNovo As Boolean) As Long
    Cliente_Grava_Customizado2 = SUCESSO
End Function

Public Function Cliente_Exclui_Customizado(ByVal objCliente As ClassCliente) As Long

Dim lErro As Long
Dim alComando(0 To 2) As Long
Dim iIndice As Integer
Dim lCliente As Long
Dim objForn As New ClassFornecedor

On Error GoTo Erro_Cliente_Exclui_Customizado

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 195819
    Next
    
    'Le a tabela FiliaisClienteTRP
    lErro = Comando_Executar(alComando(0), "SELECT CodCliente FROM FiliaisClientesTRP WHERE CodCliente <> ? AND EmpresaPai= ? ", lCliente, objCliente.lCodigo, objCliente.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 196955

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196955

    If lErro = AD_SQL_SUCESSO Then gError 196956

'    lErro = Comando_Executar(alComando(1), "DELETE FROM TRPCliEmissoresExc WHERE NumIntDocCliEmi IN (Select NumIntDoc FROM TRPCliEmissores WHERE Cliente = ?)", objCliente.lCodigo)
'    If lErro <> AD_SQL_SUCESSO Then gError 195820
'
'    lErro = Comando_Executar(alComando(2), "DELETE FROM TRPCliEmissores WHERE Cliente = ?", objCliente.lCodigo)
'    If lErro <> AD_SQL_SUCESSO Then gError 195821
    
    lErro = TRPClienteCorProd_Exclui(objCliente)
    If lErro <> SUCESSO Then gError 196955
    
    lErro = TRPClienteRepProd_Exclui(objCliente)
    If lErro <> SUCESSO Then gError 196956
    
    objForn.lCodigo = objCliente.lCodigo
    
    lErro = CF("Fornecedor_Exclui_EmTrans", objForn)
    If lErro <> SUCESSO Then gError 196956
    
'    If objCliente.lCodigo = objCliente.objInfoUsu.lEmpresaPai Then
'        lErro = TRPExcComissaoCli_Exclui(objCliente)
'        If lErro <> SUCESSO Then gError 196956
'    End If

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
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRPCLIEMISSORESEXC", gErr)

        Case 195821
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRPCLIEMISSORES", gErr)
            
        Case 196955
        
        Case 196956
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CLIENTE_PAI", gErr, objCliente.lCodigo, lCliente)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195822)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function TRPAcordo_Grava(ByVal objTRPAcordo As ClassTRPAcordos, Optional ByVal bAbreTrans As Boolean = True) As Long
    
Dim lErro As Long
Dim alComando(0 To 5) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim lTransacao As Long

On Error GoTo Erro_TRPAcordo_Grava

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 197125
    Next

    'Abertura de transação
    If bAbreTrans Then
        lTransacao = Transacao_Abrir()
        If lTransacao = 0 Then gError 197124
    End If
    
    'Le a tabelaTRPAcordos
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM TRPAcordos WHERE Numero = ? ", _
                0, lNumIntDoc, objTRPAcordo.lNumero)
    If lErro <> AD_SQL_SUCESSO Then gError 197126

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197127

    'Se existir => UPDATE, senão => INSERT
    If lErro = AD_SQL_SUCESSO Then

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 197128
        
        objTRPAcordo.lNumIntDoc = lNumIntDoc
        
        lErro = Comando_ExecutarPos(alComando(1), "UPDATE TRPAcordos SET Contrato = ?,  Cliente= ? , FilialCliente= ? , ValidadeDe = ? , ValidadeAte =?,  Observacao= ? , Descricao = ?", alComando(0), _
                objTRPAcordo.sContrato, objTRPAcordo.lCliente, objTRPAcordo.iFilialCliente, objTRPAcordo.dtValidadeDe, objTRPAcordo.dtValidadeAte, objTRPAcordo.sObservacao, objTRPAcordo.sDescricao)
        If lErro <> AD_SQL_SUCESSO Then gError 197129

    Else

        'Obter NumIntDoc
        lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_INT_PROX_TRPACORDO", lNumIntDoc)
        If lErro <> SUCESSO Then gError 197130

        objTRPAcordo.lNumIntDoc = lNumIntDoc

        lErro = Comando_Executar(alComando(1), "INSERT INTO TRPAcordos( NumIntDoc, Numero, Contrato, Cliente, FilialCliente, " & _
                "ValidadeDe, ValidadeAte, Observacao, Descricao) VALUES (?,?,?,?,?,?,?,?,?)", _
                objTRPAcordo.lNumIntDoc, objTRPAcordo.lNumero, objTRPAcordo.sContrato, objTRPAcordo.lCliente, objTRPAcordo.iFilialCliente, _
                objTRPAcordo.dtValidadeDe, objTRPAcordo.dtValidadeAte, objTRPAcordo.sObservacao, objTRPAcordo.sDescricao)
        If lErro <> AD_SQL_SUCESSO Then gError 197131

    End If
    
    lErro = TRPAcordoComissaoTarifa_Exclui(objTRPAcordo)
    If lErro <> SUCESSO Then gError 197132

    lErro = TRPAcordoComissaoTarifa_Inclui(objTRPAcordo)
    If lErro <> SUCESSO Then gError 197133

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    If bAbreTrans Then
        'Confirma a transação
        lErro = Transacao_Commit()
        If lErro <> AD_SQL_SUCESSO Then gError 197134
    End If

    TRPAcordo_Grava = SUCESSO

    Exit Function

Erro_TRPAcordo_Grava:

    TRPAcordo_Grava = gErr

    Select Case gErr

        Case 197124
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 197125
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 197126, 197127
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPACORDOS", gErr)

        Case 197128
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRPACORDOS", gErr)

        Case 197129
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_TRPACORDOS", gErr)

        Case 197130, 197132, 197133
        
        Case 197131
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRPACORDOS", gErr)
            
        Case 197134
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case 198626 To 198629
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIALCLIENTECATEGORIAS", gErr)
        
        Case 198630, 198632
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_FILIALCLIENTECATEGORIAS", gErr)

        Case 198631, 198633
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_FILIALCLIENTECATEGORIAS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197135)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    If bAbreTrans Then
        'Desfaz Transação
        Call Transacao_Rollback
    End If

    Exit Function

End Function

Function TRPAcordoComissaoTarifa_Exclui(ByVal objTRPAcordo As ClassTRPAcordos) As Long

Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim lCliente As Long

On Error GoTo Erro_TRPAcordoComissaoTarifa_Exclui

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 197120
    Next

    lErro = Comando_Executar(alComando(1), "DELETE FROM TRPAcordoComissao WHERE NumIntAcordo IN (SELECT NumIntDoc FROM TRPAcordos WHERE Numero = ?)", objTRPAcordo.lNumero)
    If lErro <> AD_SQL_SUCESSO Then gError 197121

    lErro = Comando_Executar(alComando(2), "DELETE FROM TRPAcordoTarifaDif WHERE NumIntAcordoComis IN (SELECT NumIntDoc FROM TRPAcordoComissao WHERE NumIntAcordo IN (SELECT NumIntDoc FROM TRPAcordos WHERE Numero = ?))", objTRPAcordo.lNumero)
    If lErro <> AD_SQL_SUCESSO Then gError 197122

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TRPAcordoComissaoTarifa_Exclui = SUCESSO

    Exit Function

Erro_TRPAcordoComissaoTarifa_Exclui:

    TRPAcordoComissaoTarifa_Exclui = gErr

    Select Case gErr

        Case 197120
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 197121
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRPACORDOCOMISSAO", gErr)

        Case 197122
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRPACORDOTARIFADIF", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197123)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function TRPAcordoComissaoTarifa_Inclui(ByVal objTRPAcordo As ClassTRPAcordos) As Long

Dim lErro As Long
Dim alComando(1 To 3) As Long
Dim iIndice As Integer
Dim lCliente As Long
Dim objTRPAcordoComiss As ClassTRPAcordoComissao
Dim objTarifa As ClassTRPAcordoTarifaDif
Dim lNumIntDoc As Long
Dim lNumIntDoc2 As Long

On Error GoTo Erro_TRPAcordoComissaoTarifa_Inclui

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 197114
    Next

    iIndice = 0

    If objTRPAcordo.colTRPAcordoComiss.Count > 0 Then
    
        lErro = CF("Config_ObterNumInt", "FATConfig", TRPCONFIG_NUM_INT_PROX_TRPACORDOCOMISS, lNumIntDoc, objTRPAcordo.colTRPAcordoComiss.Count)
        If lErro <> SUCESSO Then gError 197115

        For Each objTRPAcordoComiss In objTRPAcordo.colTRPAcordoComiss
        
            iIndice = iIndice + 1
            
            objTRPAcordoComiss.lNumIntDoc = lNumIntDoc
            objTRPAcordoComiss.lNumIntAcordo = objTRPAcordo.lNumIntDoc
            objTRPAcordoComiss.iSeq = iIndice
            
            lNumIntDoc = lNumIntDoc + 1

            lErro = Comando_Executar(alComando(1), "INSERT INTO TRPAcordoComissao (NumIntDoc, NumIntAcordo, Seq, Produto, Destino, PercComissao) VALUES (?,?,?,?,?,?)", _
                     objTRPAcordoComiss.lNumIntDoc, objTRPAcordoComiss.lNumIntAcordo, objTRPAcordoComiss.iSeq, objTRPAcordoComiss.sProduto, objTRPAcordoComiss.iDestino, objTRPAcordoComiss.dPercComissao)
            If lErro <> AD_SQL_SUCESSO Then gError 197116
            
            lErro = CF("Config_ObterNumInt", "FATConfig", TRPCONFIG_NUM_INT_PROX_TRPACORDODIF, lNumIntDoc2, objTRPAcordoComiss.colTRPAcordoDif.Count)
            If lErro <> SUCESSO Then gError 197115
            
            For Each objTarifa In objTRPAcordoComiss.colTRPAcordoDif
            
                objTarifa.lNumIntDoc = lNumIntDoc2
                objTarifa.lNumIntAcordoComis = objTRPAcordoComiss.lNumIntDoc
            
                lNumIntDoc2 = lNumIntDoc2 + 1
            
                lErro = Comando_Executar(alComando(3), "INSERT INTO TRPAcordoTarifaDif (NumIntDoc, NumIntAcordoComis, DiasDe, DiasAte, Diario, Valor, ValorAdicional) VALUES (?,?,?,?,?,?,?)", _
                         objTarifa.lNumIntDoc, objTarifa.lNumIntAcordoComis, objTarifa.iDiasDe, objTarifa.iDiasAte, objTarifa.iDiario, objTarifa.dValor, objTarifa.dValorAdicional)
                If lErro <> AD_SQL_SUCESSO Then gError 197116
            
            Next

        Next

    End If

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TRPAcordoComissaoTarifa_Inclui = SUCESSO
    
    Exit Function

Erro_TRPAcordoComissaoTarifa_Inclui:

    TRPAcordoComissaoTarifa_Inclui = gErr

    Select Case gErr

        Case 197114
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 197115, 197117

        Case 197116
            Call Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_TRPACORDOCOMISSAO", gErr)
           
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197119)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function TRPAcordo_Exclui(ByVal objTRPAcordo As ClassTRPAcordos) As Long

Dim lErro As Long
Dim alComando(1 To 1) As Long
Dim iIndice As Integer
Dim lTransacao As Long

On Error GoTo Erro_TRPAcordo_Exclui

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 197139

    lErro = CF("TRPAcordo_Exclui_EmTrans", objTRPAcordo)
    If lErro <> SUCESSO Then gError 197142

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 197143

    TRPAcordo_Exclui = SUCESSO

    Exit Function

Erro_TRPAcordo_Exclui:

    TRPAcordo_Exclui = gErr

    Select Case gErr

        Case 197139
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 197142
        
        Case 197143
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197144)

    End Select


    'Desfaz Transação
    Call Transacao_Rollback
    
    Exit Function

End Function

Function TRPAcordo_Exclui_EmTrans(ByVal objTRPAcordo As ClassTRPAcordos) As Long

Dim lErro As Long
Dim alComando(1 To 1) As Long
Dim iIndice As Integer

On Error GoTo Erro_TRPAcordo_Exclui_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 197140
    Next

    lErro = Comando_Executar(alComando(1), "DELETE FROM TRPAcordos WHERE Numero = ?", objTRPAcordo.lNumero)
    If lErro <> AD_SQL_SUCESSO Then gError 197141

    lErro = TRPAcordoComissaoTarifa_Exclui(objTRPAcordo)
    If lErro <> SUCESSO Then gError 197142

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TRPAcordo_Exclui_EmTrans = SUCESSO

    Exit Function

Erro_TRPAcordo_Exclui_EmTrans:

    TRPAcordo_Exclui_EmTrans = gErr

    Select Case gErr


        Case 197140
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 197141
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRPACORDOS", gErr)

        Case 197142
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197144)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Public Function TRPVou_Gera_Comis(ByVal objVou As ClassTRPVouchers, ByVal dValorBase As Double, ByVal iTipoValorBase As Integer, ByVal dtData As Date, ByVal sHistorico As String, ByVal iTipoTratamento As Integer, Optional ByVal objOcr As ClassTRPOcorrencias = Nothing) As Long
'Gera a comissão externa relacionada ao voucher

Dim lErro As Long
Dim alComando(0 To 41) As Long
Dim iIndice As Integer
Dim lContador As Long
Dim iTipoLiberacao As Integer
Dim lNumIntDocLiberacao As Long
Dim lNumIntDocOCR As Long
Dim objTipoOCR As New ClassTRPTiposOcorrencia
Dim iStatus As Integer
Dim dPercValVou As Double
Dim dValorBrutoBase As Double
Dim sTipoDoc As String
Dim dPercBase As Double
Dim iTipoLiberacaoOver As Integer
Dim lNumIntDocLiberacaoOver As Long
Dim iStatusOver As Integer
Dim iTipoLiberacaoCMCC As Integer
Dim lNumIntDocLiberacaoCMCC As Long
Dim iStatusCMCC As Integer

On Error GoTo Erro_TRPVou_Gera_Comis
                 
     'Se o voucher foi emitido antes da implantação do comissionamento
     If gdtDataInicioComisCorp <= objVou.dtData Then
 
         'Abertura de Comando
         For iIndice = LBound(alComando) To UBound(alComando)
             alComando(iIndice) = Comando_Abrir()
             If alComando(iIndice) = 0 Then gError 198158
         Next
         
         iStatus = TRP_VOU_INFO_STATUS_BLOQUEADO
                 
         'Se for uma ocorrência pega o numintdoc e o tipo
         If Not (objOcr Is Nothing) Then
         
             lNumIntDocOCR = objOcr.lNumIntDoc
             
             objTipoOCR.iCodigo = objOcr.iOrigem
             
             'Le o tipo da ocorrência para dar o tratamento adequado
             lErro = CF("TRPTiposOcorrencia_Le", objTipoOCR)
             If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 198159
         
             If lErro <> SUCESSO Then gError 198160
             
         End If
         
         'Pega o último sequencial do último item gerado para esse voucher
         lErro = TRPVou_Gera_Comis_Obtem_Seq(alComando(0), objVou, lContador)
         If lErro <> SUCESSO Then gError 198161
         
         'Verifica se o a receber já está baixado, se sim coloca o status como liberado e guarda o tipo de liberação e o numintbaixa
         lErro = TRPVou_Gera_Comis_Obtem_Status(alComando(1), alComando(41), objVou, iStatus, iTipoLiberacao, lNumIntDocLiberacao, iStatusOver, iTipoLiberacaoOver, lNumIntDocLiberacaoOver, iStatusCMCC, iTipoLiberacaoCMCC, lNumIntDocLiberacaoCMCC)
         If lErro <> SUCESSO Then gError 198162
         
         'Se é um voucher novo
         If iTipoTratamento = TRP_TIPO_TRATAMENTO_COMI_VOU_NOVO Then
         
             'Insere as informações iniciais do voucher e preenche o obj com as informações sobre o comissionamento
             'Aqui só acerta TRPVoucher sem gerar nada em TRPVoucherInfo
             lErro = TRPVou_Gera_Comis_Insere_Novo(alComando(2), alComando(3), alComando(4), alComando(5), alComando(6), alComando(7), alComando(8), alComando(9), alComando(10), alComando(11), alComando(12), objVou)
             If lErro <> SUCESSO Then gError 198163
         
         End If
         
         'Calcula o bruto base com base na comissão da agência e o valor líquido base passado para essa função
         Select Case iTipoValorBase
         
             'O valor passado é o líquido (Ocorrências - Inclusão, Alteração ou Exclusão)
             Case TRP_TIPO_VALOR_BASE_LIQ
             
                 If objVou.iCartao = MARCADO Then
                     dValorBrutoBase = dValorBase
                 Else
                     'O valor bruto é o valor líquido acrescido da comissão descontada
                     dValorBrutoBase = dValorBase / (1 - objVou.dComissaoAg)
                 End If
                 dPercBase = 0
             
             'O Valor passado já é o bruto (Vouchers - Inclusão ou alteração de valor)
             Case TRP_TIPO_VALOR_BASE_BRU
                 dValorBrutoBase = dValorBase
                 dPercBase = 0
             
             'O valor passado é um acerto de percentual
             Case TRP_TIPO_VALOR_BASE_PER
                 dValorBrutoBase = objVou.dValorBruto
                 dPercBase = dValorBase
                 
             Case Else
                 gError 198164
             
         End Select
         
         'Se for tratamento para OCR ou NVL
         'lança as informações conforme o tipo de OCR
         If iTipoTratamento = TRP_TIPO_TRATAMENTO_COMI_OCR Or iTipoTratamento = TRP_TIPO_TRATAMENTO_COMI_NVL Then
             
            If objOcr Is Nothing Then gError 198165
             
            'Percentual que o valor considerado tem em cima do voucher
            If objVou.dValorBruto <> 0 Then
                dPercValVou = -objOcr.dValorOCRBrutoVou / objVou.dValorBruto
            Else
                dPercValVou = 0
            End If
                
            lErro = TRPVou_Gera_Comis_Insere_Info(alComando(19), alComando(20), alComando(21), objVou, TRP_TIPODOC_BRUTO_TEXTO, dtData, objOcr.dValorOCRBrutoVou, sHistorico, TRP_VOU_INFO_STATUS_BLOQUEADO, 0, 0, objOcr.lNumIntDoc, TRP_CLIENTEINFO_TIPO_CLIENTE, objVou.lClienteVou, TRP_TIPO_ESTORNO_LANC_NORMAL, 0, lContador)
            If lErro <> AD_SQL_SUCESSO Then gError 198166
            
            lErro = TRPVou_Gera_Comis_Insere_Info(alComando(19), alComando(20), alComando(21), objVou, TRP_TIPODOC_CMA_TEXTO, dtData, objOcr.dValorOCRCMAVou, sHistorico, TRP_VOU_INFO_STATUS_BLOQUEADO, 0, 0, objOcr.lNumIntDoc, TRP_CLIENTEINFO_TIPO_CLIENTE, objVou.lClienteVou, TRP_TIPO_ESTORNO_LANC_NORMAL, 0, lContador)
            If lErro <> AD_SQL_SUCESSO Then gError 198166
             
             If objTipoOCR.iAlteraCMC = MARCADO Then
                 lErro = TRPVou_Gera_Comis_Trata_Estorno(alComando(13), alComando(14), alComando(15), alComando(16), alComando(17), objVou, dtData, sHistorico, TRP_TIPODOC_CMC_TEXTO, lContador, , , lNumIntDocOCR, dPercValVou)
                 If lErro <> SUCESSO Then gError 198168
             End If
 
             If objTipoOCR.iAlteraCMCC = MARCADO Then
                 lErro = TRPVou_Gera_Comis_Trata_Estorno(alComando(13), alComando(14), alComando(15), alComando(16), alComando(17), objVou, dtData, sHistorico, TRP_TIPODOC_CMCC_TEXTO, lContador, , , lNumIntDocOCR, dPercValVou)
                 If lErro <> SUCESSO Then gError 198169
             End If
             
             If objTipoOCR.iAlteraCMR = MARCADO Then
                 lErro = TRPVou_Gera_Comis_Trata_Estorno(alComando(13), alComando(14), alComando(15), alComando(16), alComando(17), objVou, dtData, sHistorico, TRP_TIPODOC_CMR_TEXTO, lContador, , , lNumIntDocOCR, dPercValVou)
                 If lErro <> SUCESSO Then gError 198170
             End If
             
             If objTipoOCR.iAlteraOVER = MARCADO Then
                 lErro = TRPVou_Gera_Comis_Trata_Estorno(alComando(13), alComando(14), alComando(15), alComando(16), alComando(17), objVou, dtData, sHistorico, TRP_TIPODOC_OVER_TEXTO, lContador, , , lNumIntDocOCR, dPercValVou)
                 If lErro <> SUCESSO Then gError 198171
             End If
                 
         End If
                  
         'Se for tratamento para exclusão de OCR ou NVL
         'Apenas relança documento gerados pelo OCR anterior com sinal invertido
         If iTipoTratamento = TRP_TIPO_TRATAMENTO_COMI_OCR_EXCLUSAO Or iTipoTratamento = TRP_TIPO_TRATAMENTO_COMI_NVL_EXCLUSAO Then
         
             'Gera o estorno através do NumIntDocOCR e do voucher que está sendo mexido
             lErro = TRPVou_Gera_Comis_Trata_Estorno2(alComando(22), alComando(23), alComando(24), alComando(25), alComando(26), objVou, dtData, sHistorico, lNumIntDocOCR, lContador)
             If lErro <> SUCESSO Then gError 198172
         
         End If
         
         'Se mexeu no valor do voucher ou ele é novo tem que inserir o valor bruto e a comissão
         If iTipoTratamento = TRP_TIPO_TRATAMENTO_COMI_VOU_VLR_ALTERADO Or iTipoTratamento = TRP_TIPO_TRATAMENTO_COMI_VOU_NOVO Then
         
             'Insere o registro do valor bruto
             lErro = TRPVou_Gera_Comis_Insere_Info(alComando(19), alComando(20), alComando(21), objVou, TRP_TIPODOC_BRUTO_TEXTO, dtData, dValorBrutoBase, sHistorico, iStatus, iTipoLiberacao, lNumIntDocLiberacao, 0, TRP_CLIENTEINFO_TIPO_CLIENTE, objVou.lClienteVou, TRP_TIPO_ESTORNO_LANC_NORMAL, 0, lContador)
             If lErro <> AD_SQL_SUCESSO Then gError 198173
             
             If objVou.iCartao = DESMARCADO Then
             
                 If Abs(Arredonda_Moeda(dValorBrutoBase * objVou.dComissaoAg)) > DELTA_VALORMONETARIO Then
                     'Insere a CMA (Comissão da agência)
                     lErro = TRPVou_Gera_Comis_Insere_Info(alComando(19), alComando(20), alComando(21), objVou, TRP_TIPODOC_CMA_TEXTO, dtData, Arredonda_Moeda(dValorBrutoBase * objVou.dComissaoAg), sHistorico, iStatus, iTipoLiberacao, lNumIntDocLiberacao, 0, TRP_CLIENTEINFO_TIPO_CLIENTE, objVou.lClienteVou, TRP_TIPO_ESTORNO_LANC_NORMAL, 0, lContador)
                     If lErro <> AD_SQL_SUCESSO Then gError 198174
                 End If
 
             Else
                 If Abs(Arredonda_Moeda(dValorBrutoBase * objVou.dComissaoAg)) > DELTA_VALORMONETARIO Then
                     'Insere a CMCC (Comissão da agência em venda sobre cartão de crédito)
                     lErro = TRPVou_Gera_Comis_Insere_Info(alComando(19), alComando(20), alComando(21), objVou, TRP_TIPODOC_CMCC_TEXTO, dtData, Arredonda_Moeda(dValorBrutoBase * objVou.dComissaoAg), sHistorico, iStatusCMCC, iTipoLiberacaoCMCC, lNumIntDocLiberacaoCMCC, 0, TRP_CLIENTEINFO_TIPO_CLIENTE, objVou.lClienteVou, TRP_TIPO_ESTORNO_LANC_NORMAL, 0, lContador)
                     If lErro <> AD_SQL_SUCESSO Then gError 198175
                 End If
             End If
             
             'Lança as comissões de representante de correntista e over
             lErro = TRPVou_Gera_Comis_Lanca_Comis(alComando(31), alComando(32), alComando(33), alComando(34), objVou, dValorBrutoBase, dtData, sHistorico, iStatus, iTipoLiberacao, lNumIntDocLiberacao, lContador, iStatusOver, iTipoLiberacaoOver, lNumIntDocLiberacaoOver)
             If lErro <> AD_SQL_SUCESSO Then gError 198176
             
         End If
         
         'Se o voucher passou a ser de cartão ou deixou de ser
         'Tem que estornar um tipo e relançar no outro
         If iTipoTratamento = TRP_TIPO_TRATAMENTO_COMI_VOU_TRANSF_CARTAO Then
 
             'Deixou de ser cartão
             If objVou.iCartao = DESMARCADO Then
                 
                 'Estorna a comissão de cartão e lança a comissão retida na agência
                 lErro = TRPVou_Gera_Comis_Trata_Estorno(alComando(26), alComando(27), alComando(28), alComando(29), alComando(30), objVou, dtData, sHistorico, TRP_TIPODOC_CMCC_TEXTO, lContador, MARCADO, TRP_TIPODOC_CMA_TEXTO)
                 If lErro <> SUCESSO Then gError 198177
             
             'Passou a ser venda de cartão
             Else
                 
                 'Estorna a comissão retida na agência e lança a comissão de cartão
                 lErro = TRPVou_Gera_Comis_Trata_Estorno(alComando(26), alComando(27), alComando(28), alComando(29), alComando(30), objVou, dtData, sHistorico, TRP_TIPODOC_CMA_TEXTO, lContador, MARCADO, TRP_TIPODOC_CMCC_TEXTO)
                 If lErro <> SUCESSO Then gError 198178
             End If
 
         End If
         
         'Se o voucher passou a ter outro representante
         If iTipoTratamento = TRP_TIPO_TRATAMENTO_COMI_VOU_TRANSF_REP Then
 
             lErro = TRPVou_Gera_Comis_Trata_Estorno(alComando(26), alComando(27), alComando(28), alComando(29), alComando(30), objVou, dtData, sHistorico, TRP_TIPODOC_CMR_TEXTO, lContador, MARCADO, , , , dPercBase, objVou.lRepresentante)
             If lErro <> SUCESSO Then gError 198179
 
         End If
         
         'Se o voucher passou a ter outro correntista
         If iTipoTratamento = TRP_TIPO_TRATAMENTO_COMI_VOU_TRANSF_COR Then
 
             lErro = TRPVou_Gera_Comis_Trata_Estorno(alComando(26), alComando(27), alComando(28), alComando(29), alComando(30), objVou, dtData, sHistorico, TRP_TIPODOC_CMC_TEXTO, lContador, MARCADO, , , , dPercBase, objVou.lCorrentista)
             If lErro <> SUCESSO Then gError 198180
 
         End If
         
         'Se o voucher passou a ter outro emissor
         If iTipoTratamento = TRP_TIPO_TRATAMENTO_COMI_VOU_TRANSF_EMI Then
 
             lErro = TRPVou_Gera_Comis_Trata_Estorno(alComando(26), alComando(27), alComando(28), alComando(29), alComando(30), objVou, dtData, sHistorico, TRP_TIPODOC_OVER_TEXTO, lContador, MARCADO, , , , dPercBase, objVou.lEmissor)
             If lErro <> SUCESSO Then gError 198181
 
         End If
         
         'Se o voucher passou a ter outra agência
         If iTipoTratamento = TRP_TIPO_TRATAMENTO_COMI_VOU_TRANSF_CLI Then
 
            'Estorna o bruto
            lErro = TRPVou_Gera_Comis_Trata_Estorno(alComando(26), alComando(27), alComando(28), alComando(29), alComando(30), objVou, dtData, sHistorico, TRP_TIPODOC_BRUTO_TEXTO, lContador, MARCADO, , , , 1, objVou.lClienteComissao)
            If lErro <> SUCESSO Then gError 198182
            
            'Estorna a comissão (CMA ou CMCC)
            If objVou.iCartao = MARCADO Then
                 
                 lErro = TRPVou_Gera_Comis_Trata_Estorno(alComando(26), alComando(27), alComando(28), alComando(29), alComando(30), objVou, dtData, sHistorico, TRP_TIPODOC_CMCC_TEXTO, lContador, MARCADO, , , , dPercBase, objVou.lClienteComissao)
                 If lErro <> SUCESSO Then gError 198182
             Else
                 
                 lErro = TRPVou_Gera_Comis_Trata_Estorno(alComando(26), alComando(27), alComando(28), alComando(29), alComando(30), objVou, dtData, sHistorico, TRP_TIPODOC_CMA_TEXTO, lContador, MARCADO, , , , dPercBase, objVou.lClienteComissao)
                 If lErro <> SUCESSO Then gError 198183
             End If
 
         End If
         
         'Se alterou o percentual de comissão do representante
         If iTipoTratamento = TRP_TIPO_TRATAMENTO_COMI_CMR_ALT_PERCCOMI Then
         
             If objVou.lRepresentante <> 0 Then
         
                If iTipoValorBase = TRP_TIPO_VALOR_BASE_PER Then
                    lErro = TRPVou_Gera_Comis_Insere_Info(alComando(19), alComando(20), alComando(21), objVou, TRP_TIPODOC_CMR_TEXTO, dtData, objVou.dValorBruto * dPercBase, sHistorico, iStatus, iTipoLiberacao, lNumIntDocLiberacao, 0, TRP_CLIENTEINFO_TIPO_CLIENTE, objVou.lRepresentante, TRP_TIPO_ESTORNO_LANC_NORMAL, 0, lContador)
                    If lErro <> AD_SQL_SUCESSO Then gError 198184
                Else
                    lErro = TRPVou_Gera_Comis_Insere_Info(alComando(19), alComando(20), alComando(21), objVou, TRP_TIPODOC_CMR_TEXTO, dtData, dValorBase, sHistorico, iStatus, iTipoLiberacao, lNumIntDocLiberacao, 0, TRP_CLIENTEINFO_TIPO_CLIENTE, objVou.lRepresentante, TRP_TIPO_ESTORNO_LANC_NORMAL, 0, lContador)
                    If lErro <> AD_SQL_SUCESSO Then gError 198184
                End If
             
             End If
         
         End If
 
         'Se alterou o percentual de comissão do correntista
         If iTipoTratamento = TRP_TIPO_TRATAMENTO_COMI_CMC_ALT_PERCCOMI Then
         
             If objVou.lCorrentista <> 0 Then
         
                If iTipoValorBase = TRP_TIPO_VALOR_BASE_PER Then
                    lErro = TRPVou_Gera_Comis_Insere_Info(alComando(19), alComando(20), alComando(21), objVou, TRP_TIPODOC_CMC_TEXTO, dtData, objVou.dValorBruto * dPercBase, sHistorico, iStatus, iTipoLiberacao, lNumIntDocLiberacao, 0, TRP_CLIENTEINFO_TIPO_CLIENTE, objVou.lCorrentista, TRP_TIPO_ESTORNO_LANC_NORMAL, 0, lContador)
                    If lErro <> AD_SQL_SUCESSO Then gError 198185
                Else
                    lErro = TRPVou_Gera_Comis_Insere_Info(alComando(19), alComando(20), alComando(21), objVou, TRP_TIPODOC_CMC_TEXTO, dtData, dValorBase, sHistorico, iStatus, iTipoLiberacao, lNumIntDocLiberacao, 0, TRP_CLIENTEINFO_TIPO_CLIENTE, objVou.lCorrentista, TRP_TIPO_ESTORNO_LANC_NORMAL, 0, lContador)
                    If lErro <> AD_SQL_SUCESSO Then gError 198185
                End If
             
             End If
         
         End If
         
         'Se alterou o percentual de comissão do emissor
         If iTipoTratamento = TRP_TIPO_TRATAMENTO_COMI_OVER_ALT_PERCCOMI Then
         
             If objVou.lEmissor <> 0 And objVou.iDiasAntc = DESMARCADO Then
         
                If iTipoValorBase = TRP_TIPO_VALOR_BASE_PER Then
                    lErro = TRPVou_Gera_Comis_Insere_Info(alComando(19), alComando(20), alComando(21), objVou, TRP_TIPODOC_OVER_TEXTO, dtData, objVou.dValorBruto * dPercBase, sHistorico, iStatusOver, iTipoLiberacaoOver, lNumIntDocLiberacaoOver, 0, TRP_CLIENTEINFO_TIPO_FORNECEDOR, objVou.lEmissor, TRP_TIPO_ESTORNO_LANC_NORMAL, 0, lContador)
                    If lErro <> AD_SQL_SUCESSO Then gError 198186
                Else
                    lErro = TRPVou_Gera_Comis_Insere_Info(alComando(19), alComando(20), alComando(21), objVou, TRP_TIPODOC_OVER_TEXTO, dtData, dValorBase, sHistorico, iStatusOver, iTipoLiberacaoOver, lNumIntDocLiberacaoOver, 0, TRP_CLIENTEINFO_TIPO_FORNECEDOR, objVou.lEmissor, TRP_TIPO_ESTORNO_LANC_NORMAL, 0, lContador)
                    If lErro <> AD_SQL_SUCESSO Then gError 198186
                End If
                
             End If
         
         End If
         
         'Se alterou o percentual de comissão da agência
         If iTipoTratamento = TRP_TIPO_TRATAMENTO_COMI_CMA_CMCC_ALT_PERCCOMI Then
         
             If objVou.iCartao = MARCADO Then
         
                If iTipoValorBase = TRP_TIPO_VALOR_BASE_PER Then
                    lErro = TRPVou_Gera_Comis_Insere_Info(alComando(19), alComando(20), alComando(21), objVou, TRP_TIPODOC_CMCC_TEXTO, dtData, objVou.dValorBruto * dPercBase, sHistorico, iStatusCMCC, iTipoLiberacaoCMCC, lNumIntDocLiberacaoCMCC, 0, TRP_CLIENTEINFO_TIPO_CLIENTE, objVou.lClienteVou, TRP_TIPO_ESTORNO_LANC_NORMAL, 0, lContador)
                    If lErro <> AD_SQL_SUCESSO Then gError 198187
                Else
                    lErro = TRPVou_Gera_Comis_Insere_Info(alComando(19), alComando(20), alComando(21), objVou, TRP_TIPODOC_CMCC_TEXTO, dtData, dValorBase, sHistorico, iStatusCMCC, iTipoLiberacaoCMCC, lNumIntDocLiberacaoCMCC, 0, TRP_CLIENTEINFO_TIPO_CLIENTE, objVou.lClienteVou, TRP_TIPO_ESTORNO_LANC_NORMAL, 0, lContador)
                    If lErro <> AD_SQL_SUCESSO Then gError 198187
                End If
             Else
       
                If iTipoValorBase = TRP_TIPO_VALOR_BASE_PER Then
                    lErro = TRPVou_Gera_Comis_Insere_Info(alComando(19), alComando(20), alComando(21), objVou, TRP_TIPODOC_CMA_TEXTO, dtData, objVou.dValorBruto * dPercBase, sHistorico, iStatus, iTipoLiberacao, lNumIntDocLiberacao, 0, TRP_CLIENTEINFO_TIPO_CLIENTE, objVou.lClienteVou, TRP_TIPO_ESTORNO_LANC_NORMAL, 0, lContador)
                    If lErro <> AD_SQL_SUCESSO Then gError 198188
                Else
                    lErro = TRPVou_Gera_Comis_Insere_Info(alComando(19), alComando(20), alComando(21), objVou, TRP_TIPODOC_CMA_TEXTO, dtData, dValorBase, sHistorico, iStatus, iTipoLiberacao, lNumIntDocLiberacao, 0, TRP_CLIENTEINFO_TIPO_CLIENTE, objVou.lClienteVou, TRP_TIPO_ESTORNO_LANC_NORMAL, 0, lContador)
                    If lErro <> AD_SQL_SUCESSO Then gError 198188
                End If
                 
             End If
         
         End If
         
         'Acerta os valores finais de comissão e valor líquido no voucher
         lErro = CF("TRPVouchers_Grava", objVou, TRP_TIPO_GRAVACAO_ACERTO_VALOR_LIQ)
         If lErro <> SUCESSO Then gError 198188
         
         'Fecha Comando
         For iIndice = LBound(alComando) To UBound(alComando)
             Call Comando_Fechar(alComando(iIndice))
         Next
     
    End If
 
    TRPVou_Gera_Comis = SUCESSO

    Exit Function

Erro_TRPVou_Gera_Comis:

    TRPVou_Gera_Comis = gErr

    Select Case gErr

        Case 198156, 198159, 198161 To 198163, 198166 To 198188
            
        Case 198157 'Voucher não cadastrado
            Call Rotina_Erro(vbOKOnly, "ERRO_VOUCHER_NAO_CADASTRADO", gErr, objVou.lNumVou, objVou.sSerie, objVou.sTipVou)
        
        Case 198158
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 198160 'Tipo de ocorrência não cadastrado
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_TRPTIPOSOCORRENCIA_NAO_PREENCHIDO", gErr)
        
        Case 198164 'Tipo de valor base desconhecido iTipoValorBase
            Call Rotina_Erro(vbOKOnly, "ERRO_TRP_TIPOVALOR_DESCONHECIDO", gErr, iTipoValorBase)
        
        Case 198165 'Ocorrência não passada
            Call Rotina_Erro(vbOKOnly, "ERRO_TRP_OCR_NAO_INFORMADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198189)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Private Function TRPVou_Gera_Comis_Obtem_Seq(ByVal lComando As Long, ByVal objVou As ClassTRPVouchers, lContador As Long) As Long
'Gera a comissão externa relacionada ao voucher

Dim lErro As Long

On Error GoTo Erro_TRPVou_Gera_Comis_Obtem_Seq

     'Le o sequencial do histórico do Voucher
    lErro = Comando_Executar(lComando, "SELECT MAX(Seq) FROM TRPVoucherInfo WHERE TipVou = ? AND Serie = ? AND NumVou = ? ", _
                lContador, objVou.sTipVou, objVou.sSerie, objVou.lNumVou)
    If lErro <> AD_SQL_SUCESSO Then gError 198095

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 198096
    
    If lErro <> AD_SQL_SUCESSO Then lContador = 0
 
    TRPVou_Gera_Comis_Obtem_Seq = SUCESSO

    Exit Function

Erro_TRPVou_Gera_Comis_Obtem_Seq:

    TRPVou_Gera_Comis_Obtem_Seq = gErr

    Select Case gErr

        Case 198095, 198096
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPVOUCHERINFO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198097)

    End Select

    Exit Function

End Function

Private Function TRPVou_Gera_Comis_Insere_Info(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lComando3 As Long, ByVal objVou As ClassTRPVouchers, ByVal sTipoDoc As String, ByVal dtData As Date, ByVal dValor As Double, ByVal sHistorico As String, ByVal iStatus As Integer, ByVal iTipoLiberacao As Integer, ByVal lNumIntDocLiberacao As Long, ByVal lNumIntDocOCR As Long, ByVal iTipoCliForn As Integer, ByVal lCliForn As Long, ByVal iEstorno As Integer, ByVal lNumIntDocEstorno As Long, lContador As Long, Optional lNumInDocInfo As Long) As Long
'Gera a comissão externa relacionada ao voucher

Dim lErro As Long
Dim lNumIntDoc As Long
Dim objCtb As New ClassTRPCTBVouInfo
Dim objContabAutomatica As New ClassContabAutomatica

On Error GoTo Erro_TRPVou_Gera_Comis_Insere_Info

    If Abs(dValor) > DELTA_VALORMONETARIO Then

        'Obter NumIntDoc
        lErro = CF("Config_ObterNumInt2", lComando1, lComando2, "TRPConfig", "NUM_INT_PROX_TRPVOURCHERINFO", lNumIntDoc)
        If lErro <> SUCESSO Then gError 198153
        
        lContador = lContador + 1
        
        'Insere em TRPVoucherInfo
        lErro = Comando_Executar(lComando3, "INSERT INTO TRPVoucherInfo (NumIntDoc,TipoDoc,TipVou,Serie,NumVou,Seq,Data,TipoDocDestino,NumIntDocDestino,NumIntDocComiInt,Valor,Historico,NumTitulo, Status,TipoLiberacao,NumIntDocLiberacao, NumIntDocOcr,TipoCliForn,CliForn, Estorno, NumIntDocEstorno, DataRegistro, HoraRegistro, Usuario) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) ", _
                     lNumIntDoc, sTipoDoc, objVou.sTipVou, objVou.sSerie, objVou.lNumVou, lContador, dtData, 0, 0, 0, Arredonda_Moeda(dValor), sHistorico, 0, iStatus, iTipoLiberacao, lNumIntDocLiberacao, lNumIntDocOCR, iTipoCliForn, lCliForn, iEstorno, lNumIntDocEstorno, Date, CDbl(Time), gsUsuario)
        If lErro <> AD_SQL_SUCESSO Then gError 198154
        
'        'gravar a contabilizacao
'        lErro = objContabAutomatica.Inicializa_Contab(objCtb, MODULO_FATURAMENTO, "FAT", objVou.dtData, objVou.dtData)
'        If lErro = SUCESSO Then gError 200271
'
'        lErro = objContabAutomatica.GeraContabilizacao(objVou)
'        If lErro <> SUCESSO Then gError 200271
'
'        lErro = objContabAutomatica.Finaliza_Contab()
'        If lErro <> SUCESSO Then gError 200272
'
'        Set objContabAutomatica = Nothing
'        Call objCtb.Encerrando
'        Set objCtb = Nothing
           
    End If
            
    lNumInDocInfo = lNumIntDoc
            
    TRPVou_Gera_Comis_Insere_Info = SUCESSO

    Exit Function

Erro_TRPVou_Gera_Comis_Insere_Info:

    TRPVou_Gera_Comis_Insere_Info = gErr

    Select Case gErr

        Case 198153, 200271, 200272
            
        Case 198154
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERT_TRPVOUCHERINFO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198155)

    End Select

    Exit Function

End Function

Private Function TRPVou_Gera_Comis_Obtem_Status(ByVal lComando As Long, ByVal lComando1 As Long, ByVal objVou As ClassTRPVouchers, iStatus As Integer, iTipoLiberacao As Integer, lNumIntDocLiberacao As Long, iStatusOver As Integer, iTipoLiberacaoOver As Integer, lNumIntDocLiberacaoOver As Long, iStatusCMCC As Integer, iTipoLiberacaoCMCC As Integer, lNumIntDocLiberacaoCMCC As Long) As Long
'Gera a comissão externa relacionada ao voucher

Dim lErro As Long
Dim lNumIntDocBaixa As Long
Dim sItem As String

On Error GoTo Erro_TRPVou_Gera_Comis_Obtem_Status

    'Se era para gerar a comissão bloqueada e o voucher já tiver faturado, verifica se o título já está baixado, se sim muda
    'o status para liberado e guarda as informações de o que liberou
    If iStatus = TRP_VOU_INFO_STATUS_BLOQUEADO And objVou.lNumIntDocDestino <> 0 And objVou.iTipoDocDestino = TRP_TIPO_DOC_DESTINO_TITREC Then
        
        'Le o Voucher
        lErro = Comando_Executar(lComando, "SELECT B.NumIntDoc FROM ParcelasRecTodas AS P, BaixasParcRec AS B WHERE P.NumIntTiTulo = ? AND P.NumIntDoc = B.NumIntParcela AND B.Status <> ? ", _
                    lNumIntDocBaixa, objVou.lNumIntDocDestino, STATUS_EXCLUIDO)
        If lErro <> AD_SQL_SUCESSO Then gError 198150
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 198151
    
        If lErro = AD_SQL_SUCESSO Then
            iStatus = TRP_VOU_INFO_STATUS_LIBERADO
            iTipoLiberacao = TRP_TIPO_LIBERACAO_COMISSAO_BAIXA
            lNumIntDocLiberacao = lNumIntDocBaixa
        End If
    
    End If
    
    iStatusCMCC = iStatus
    iTipoLiberacaoCMCC = iTipoLiberacao
    lNumIntDocLiberacaoCMCC = lNumIntDocLiberacao
    
    If objVou.iCartao = MARCADO And objVou.iQuantParc = 1 Then
        iStatusCMCC = TRP_VOU_INFO_STATUS_LIBERADO
        iTipoLiberacaoCMCC = TRP_TIPO_LIBERACAO_COMISSAO_EMI
        lNumIntDocLiberacaoCMCC = 0
    End If
    
    iStatusOver = iStatus
    iTipoLiberacaoOver = iTipoLiberacao
    lNumIntDocLiberacaoOver = lNumIntDocLiberacao
    
    sItem = String(255, 0)

    'Verifica o tipo de tratamento (Se o pagamento é na baixa ou na emissão)
    lErro = Comando_Executar(lComando1, "SELECT Item FROM FilialFornecedorCategorias WHERE Fornecedor = ? AND Filial = ? AND Categoria = ? ", sItem, objVou.lEmissor, FILIAL_MATRIZ, TRP_CATEGORIA_LIBCME)
    If lErro <> AD_SQL_SUCESSO Then gError 198117
    
    'Busca a primeira
    lErro = Comando_BuscarProximo(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 198118
    
    If lErro = AD_SQL_SUCESSO Then
    
        'Se for no faturamento e estiver faturado
        If sItem = TRP_CATEGORIA_LIBCME_FAT And objVou.lNumIntDocDestino <> 0 And objVou.iTipoDocDestino = TRP_TIPO_DOC_DESTINO_TITREC Then
            iStatusOver = TRP_VOU_INFO_STATUS_LIBERADO
            iTipoLiberacaoOver = TRP_TIPO_LIBERACAO_COMISSAO_FAT
            lNumIntDocLiberacaoOver = objVou.lNumIntDocDestino
            
        'Se for na emissão do voucher
        ElseIf sItem = TRP_CATEGORIA_LIBCME_EMI Then
            iStatusOver = TRP_VOU_INFO_STATUS_LIBERADO
            iTipoLiberacaoOver = TRP_TIPO_LIBERACAO_COMISSAO_EMI
            lNumIntDocLiberacaoOver = 0
        End If
    End If
 
    TRPVou_Gera_Comis_Obtem_Status = SUCESSO

    Exit Function

Erro_TRPVou_Gera_Comis_Obtem_Status:

    TRPVou_Gera_Comis_Obtem_Status = gErr

    Select Case gErr

        Case 198117, 198118
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIALFORNECEDORCATEGORIAS2", gErr, objVou.lEmissor)

        Case 198150, 198151
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BAIXASPARCREC", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198152)

    End Select

    Exit Function

End Function

Public Function TRPVou_Le_Dados_Comis_Cliente(ByVal objVou As ClassTRPVouchers) As Long

Dim lErro As Long
Dim alComando(0 To 10) As Long
Dim iIndice As Integer

On Error GoTo Erro_TRPVou_Le_Dados_Comis_Cliente

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 198190
    Next
    
    lErro = TRPVou_Gera_Comis_Insere_Novo(alComando(0), alComando(1), alComando(2), alComando(3), alComando(4), alComando(5), alComando(6), alComando(7), alComando(8), alComando(9), alComando(10), objVou, True)
    If lErro <> SUCESSO Then gError 198191
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    TRPVou_Le_Dados_Comis_Cliente = SUCESSO

    Exit Function

Erro_TRPVou_Le_Dados_Comis_Cliente:

    TRPVou_Le_Dados_Comis_Cliente = gErr

    Select Case gErr

        Case 198190
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 198191

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198192)

    End Select
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function TRPVoucher_Exclui_Comissao(ByVal objVou As ClassTRPVouchers) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim lTransacao As Long

On Error GoTo Erro_TRPVoucher_Exclui_Comissao

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 198203
   
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 198204
    Next
    
    'Le a tabelaTRPTiposOcorrencia
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM TRPVoucherInfo WHERE TipVou  = ? AND Serie  = ? AND NumVou = ? ", _
                0, lNumIntDoc, objVou.sTipVou, objVou.sSerie, objVou.lNumVou)
    If lErro <> AD_SQL_SUCESSO Then gError 198205

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 198206

    Do While lErro <> AD_SQL_SEM_DADOS

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 198207
    
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM TRPVoucherInfo", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 198208
        
        'Busca Próximo
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 198209
        
    Loop
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 198210
    
    TRPVoucher_Exclui_Comissao = SUCESSO

    Exit Function

Erro_TRPVoucher_Exclui_Comissao:

    TRPVoucher_Exclui_Comissao = gErr

    Select Case gErr

        Case 198203
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 198204
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 198205, 198206, 198209
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPVOUCHERINFO", gErr)
            
        Case 198207
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRPVOUCHERINFO", gErr)
        
        Case 198208
            Call Rotina_Erro(vbOKOnly, "ERRO_DELETE_TRPVOUCHERINFO", gErr)
        
        Case 198210
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198211)

    End Select
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Desfaz Transação
    Call Transacao_Rollback
    
    Exit Function

End Function

Public Function TRPVouComi_Grava(ByVal objVou As ClassTRPVouchers) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_TRPVouComi_Grava

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 198212
   
    'Grava o/a TRPTiposOcorrencia no Banco de Dados
    lErro = TRPVouchers_Grava(objVou, TRP_TIPO_GRAVACAO_VOU_COMISSAO)
    If lErro <> SUCESSO Then gError 198213
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 198214
    
    TRPVouComi_Grava = SUCESSO

    Exit Function

Erro_TRPVouComi_Grava:

    TRPVouComi_Grava = gErr

    Select Case gErr

        Case 198212
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 198213

        Case 198214
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198215)

    End Select
    
    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function TRPVouManu_Grava(ByVal objVou As ClassTRPVouchers) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_TRPVouManu_Grava

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 198212
    
    'Grava
    lErro = TRPVouchers_Grava(objVou, TRP_TIPO_GRAVACAO_VOU_MANUTENCAO)
    If lErro <> SUCESSO Then gError 198213
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 198214
    
    TRPVouManu_Grava = SUCESSO

    Exit Function

Erro_TRPVouManu_Grava:

    TRPVouManu_Grava = gErr

    Select Case gErr

        Case 198212
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 198213

        Case 198214
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198215)

    End Select
    
    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Private Function TRPVou_Gera_Comis_Insere_Novo(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lComando3 As Long, ByVal lComando4 As Long, ByVal lComando5 As Long, ByVal lComando6 As Long, ByVal lComando7 As Long, ByVal lComando8 As Long, ByVal lComando9 As Long, ByVal lComando10 As Long, ByVal lComando11 As Long, ByVal objVou As ClassTRPVouchers, Optional ByVal bSoLe As Boolean = False) As Long
'Gera a comissão externa relacionada ao voucher

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim lCliCMR As Long
Dim dPercCMR As Double
Dim lCliCMC As Long
Dim dPercCMC As Double
Dim dPercAgencia As Double
Dim dPercOVER As Double
Dim lNumVou As Long
Dim lNumIntDocCliEmi As Long

On Error GoTo Erro_TRPVou_Gera_Comis_Insere_Novo
   
    'Le as informações padrões, depois acerta as exceções
    'Le o representante e o correntista e o percentual de comissao
    lErro = Comando_Executar(lComando1, "SELECT Representante, ComissaoRep, Correntista, ComissaoCorr, ComissaoAg FROM FiliaisClientesTRP WHERE CodCliente = ? ", _
                lCliCMR, dPercCMR, lCliCMC, dPercCMC, dPercAgencia, objVou.lClienteVou)
    If lErro <> AD_SQL_SUCESSO Then gError 198128

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 198129
    
    If lErro = AD_SQL_SUCESSO Then
        objVou.lRepresentante = lCliCMR
        objVou.dComissaoRep = dPercCMR
        objVou.lCorrentista = lCliCMC
        objVou.dComissaoCorr = dPercCMC
        objVou.dComissaoAg = dPercAgencia
    End If
    
    lErro = CF("Produto_Formata", objVou.sProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 198130
    
    'Exceção de Representante por produto
    'Le o representante e o percentual de comissao
    lErro = Comando_Executar(lComando2, "SELECT Representante, PercComis FROM TRPClienteRepProd WHERE Cliente = ? AND Produto = ? ", _
                lCliCMR, dPercCMR, objVou.lClienteVou, sProdutoFormatado)
    If lErro <> AD_SQL_SUCESSO Then gError 198131

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando2)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 198132
    
    If lErro = AD_SQL_SUCESSO Then
        objVou.lRepresentante = lCliCMR
        objVou.dComissaoRep = dPercCMR
    End If
    
    'Exceção de Correntista por produto
    'Le o correntista e o percentual de comissao
    lErro = Comando_Executar(lComando3, "SELECT Correntista, PercComis FROM TRPClienteCorProd WHERE Cliente = ? AND Produto = ? ", _
                lCliCMC, dPercCMC, objVou.lClienteVou, sProdutoFormatado)
    If lErro <> AD_SQL_SUCESSO Then gError 198133

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando3)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 198134
    
    If lErro = AD_SQL_SUCESSO Then
        objVou.lCorrentista = lCliCMC
        objVou.dComissaoCorr = dPercCMC
    End If
    
    'Percentual da agência exceção nível 1
    'Le o percentual de comissão da agência no acordo
    lErro = Comando_Executar(lComando6, "SELECT C.PercComissao FROM TRPExcComissaoCli AS C, FiliaisClientesTRP AS F WHERE C.Cliente = F.EmpresaPai AND F.CodCliente = ? AND C.Produto = ? ", _
                dPercAgencia, objVou.lClienteVou, sProdutoFormatado)
    If lErro <> AD_SQL_SUCESSO Then gError 198136

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando6)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 198137
    
    'Se existir acordo o percentual de comissão da agência vem do acordo
    If lErro = AD_SQL_SUCESSO Then
        objVou.dComissaoAg = dPercAgencia
    End If

    'Percentual da agência exceção nível 2
    'Le a tabela de acordos por produto\destino
    lErro = Comando_Executar(lComando7, "SELECT C.PercComissao FROM TRPAcordoComissao AS C, TRPAcordos AS A, FiliaisClientesTRP AS F WHERE C.NumIntAcordo = A.NumIntDoc AND A.Cliente = F.EmpresaPai AND F.CodCliente = ? AND A.FilialCliente = ? AND (A.ValidadeDe <= ? OR A.ValidadeDe = ?) AND (A.ValidadeAte >= ? OR A.ValidadeAte = ?) AND C.Produto = ?  AND C.Destino = ?", _
                dPercAgencia, objVou.lClienteVou, FILIAL_MATRIZ, objVou.dtData, DATA_NULA, objVou.dtData, DATA_NULA, sProdutoFormatado, objVou.iDestino)
    If lErro <> AD_SQL_SUCESSO Then gError 198138

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando7)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 198139
    
    'Se tem tarifa diferenciada não tem comissão para agência
    If lErro = AD_SQL_SUCESSO Then
        objVou.dComissaoAg = dPercAgencia
    End If

    'Se tem emissor
    If objVou.lEmissor <> 0 Then
    
        'Le o Emissor e sua comissão
        lErro = Comando_Executar(lComando8, "SELECT NumIntDoc, PercComissao FROM TRPCliEmissores WHERE Cliente = ? AND FornEmissor = ? ", _
                     lNumIntDocCliEmi, dPercOVER, objVou.lClienteVou, objVou.lEmissor)
        If lErro <> AD_SQL_SUCESSO Then gError 198140
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(lComando8)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 198141
    
        If lErro <> AD_SQL_SUCESSO Then gError 198142 'Esse emissor não está associado a esse cliente
        
        objVou.dComissaoEmissor = dPercOVER
        
        'Over exceção nível 1
        'Le as exceções
        lErro = Comando_Executar(lComando9, "SELECT PercComissao FROM TRPCliEmissoresExc WHERE NumIntDocCliEmi = ? AND Produto = ? ", _
                     dPercOVER, lNumIntDocCliEmi, sProdutoFormatado)
        If lErro <> AD_SQL_SUCESSO Then gError 198143
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(lComando9)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 198144
            
        If lErro = AD_SQL_SUCESSO Then
            objVou.dComissaoEmissor = dPercOVER
        End If

    End If
    
    If Not bSoLe Then
    
        'Le o voucher para o update
        lErro = Comando_ExecutarPos(lComando10, "SELECT NumVou FROM TRPVouchers WHERE NumVou = ? AND Serie = ? AND TipVou= ?  ", 0, _
                     lNumVou, objVou.lNumVou, objVou.sSerie, objVou.sTipVou)
        If lErro <> AD_SQL_SUCESSO Then gError 198145
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(lComando10)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 198146
        
        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(lComando10)
        If lErro <> SUCESSO Then gError 198147
        
        'Marca as informações pertinentes a comissão no voucher, a partir daí esses
        'serão os valores a serem considerados mesmo se houverem alterações no cadastro de clientes
        '(informações históricas - qual era o representante na emissão do voucher, qual o % de comissão, etc ...)
        lErro = Comando_ExecutarPos(lComando11, "UPDATE TRPVouchers SET Representante = ?, ComissaoRep = ?, Correntista = ?, ComissaoCorr = ?, ComissaoAg = ?, ComissaoEmissor = ? ", _
                    lComando10, objVou.lRepresentante, objVou.dComissaoRep, objVou.lCorrentista, objVou.dComissaoCorr, objVou.dComissaoAg, objVou.dComissaoEmissor)
        If lErro <> AD_SQL_SUCESSO Then gError 198148
        
    End If
 
    TRPVou_Gera_Comis_Insere_Novo = SUCESSO

    Exit Function

Erro_TRPVou_Gera_Comis_Insere_Novo:

    TRPVou_Gera_Comis_Insere_Novo = gErr

    Select Case gErr

        Case 198128, 198129
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISCLIENTESTRP", gErr)
        
        Case 198130, 198135
        
        Case 198131, 198132
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPCLIENTEREPPROD", gErr)

        Case 198133, 198134
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPCLIENTECORPROD", gErr)
            
        Case 198136 To 198139
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPACORDOS", gErr)

        Case 198140, 198141
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPCLIEMISSORES", gErr)
            
        Case 198142
            Call Rotina_Erro(vbOKOnly, "ERRO_TRPCLIEMISSORES_NAO_CADASTRADO", gErr, objVou.lEmissor, objVou.lClienteVou)
        
        Case 198143, 198144
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPCLIEMISSORESEXC", gErr)
            
        Case 198145, 198146
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPVOUCHERS", gErr)
            
        Case 198147
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRPVOUCHERS", gErr)
            
        Case 198148
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRPVOUCHERS", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198149)

    End Select

    Exit Function

End Function

Private Function TRPVou_Gera_Comis_Trata_Estorno(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lComando3 As Long, ByVal lComando4 As Long, ByVal lComando5 As Long, ByVal objVou As ClassTRPVouchers, ByVal dtData As Date, ByVal sHistorico As String, ByVal sTipoDocEstornado As String, lContador As Long, Optional ByVal iRelancarOutroTipo As Integer = DESMARCADO, Optional ByVal sTipoRelancamento As String = "", Optional ByVal lNumIntDocOCR As Long = 0, Optional ByVal dFatorEstorno As Double = 1, Optional ByVal dFatorRelanc As Double = 1, Optional ByVal lNovoCliForn As Long = 0) As Long
'Gera a comissão externa relacionada ao voucher

Dim lErro As Long
Dim lNumIntDoc As Long
Dim dValorEstornar As Double
Dim dValorEstornarRelanc As Double
Dim dValorEstornarAux As Double
Dim lCliForn As Long
Dim iStatus As Integer
Dim iTipoCliForn As Integer
Dim iTipoLiberacao As Integer
Dim lNumIntDocLiberacao As Long
Dim lNumIntDocVouInfo As Long
Dim lNumIntDocEstorno As Long
Dim lSeq As Long
Dim objVouInfo As ClassTRPVoucherInfo
Dim objVouInfoAux As ClassTRPVoucherInfo
Dim colVouInfoItens As Collection
Dim colcolVouInfoItens As New Collection
Dim colVouInfoGrupado As New Collection
Dim bAchou As Boolean
Dim iIndice As Integer

On Error GoTo Erro_TRPVou_Gera_Comis_Trata_Estorno

    'Le a TRPVoucherInfo para pegar o saldo do que foi lançado de comissão para estornar e relançar se necessário
    'Só le o que não é um documento estornado ou um documento estornador e gera os estornos
    lErro = Comando_ExecutarPos(lComando1, "SELECT NumIntDoc, Valor, TipoCliForn, CliForn,Status,TipoLiberacao,NumIntDocLiberacao FROM TRPVoucherInfo WHERE TipoDoc = ? AND TipVou = ? AND Serie = ? AND NumVou = ? ", 0, _
                 lNumIntDocVouInfo, dValorEstornarAux, iTipoCliForn, lCliForn, iStatus, iTipoLiberacao, lNumIntDocLiberacao, sTipoDocEstornado, objVou.sTipVou, objVou.sSerie, objVou.lNumVou)
    If lErro <> AD_SQL_SUCESSO Then gError 198100

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 198101

    Do While lErro <> AD_SQL_SEM_DADOS
    
        lErro = Comando_LockExclusive(lComando1)
        If lErro <> AD_SQL_SUCESSO Then gError 198102
        
        Set objVouInfo = New ClassTRPVoucherInfo
        
        objVouInfo.lNumIntDoc = lNumIntDocVouInfo
        objVouInfo.dValor = dValorEstornarAux
        objVouInfo.iTipoCliForn = iTipoCliForn
        objVouInfo.lCliForn = lCliForn
        objVouInfo.iStatus = iStatus
        objVouInfo.iTipoLiberacao = iTipoLiberacao
        objVouInfo.lNumIntDocLiberacao = lNumIntDocLiberacao
        
        bAchou = False
        iIndice = 0
        For Each objVouInfoAux In colVouInfoGrupado
            iIndice = iIndice + 1
            If 1 = 1 And _
                        objVouInfo.iTipoCliForn = objVouInfoAux.iTipoCliForn And _
                        objVouInfo.lCliForn = objVouInfoAux.lCliForn And _
                        objVouInfo.iStatus = objVouInfoAux.iStatus And _
                        objVouInfo.iTipoLiberacao = objVouInfoAux.iTipoLiberacao And _
                        objVouInfo.lNumIntDocLiberacao = objVouInfoAux.lNumIntDocLiberacao Then
                bAchou = True
                objVouInfoAux.dValor = objVouInfoAux.dValor + objVouInfo.dValor
                Set colVouInfoItens = colcolVouInfoItens.Item(iIndice)
                colVouInfoItens.Add objVouInfo
                Exit For
            End If
        Next
        If Not bAchou Then
            colVouInfoGrupado.Add objVouInfo
            Set colVouInfoItens = New Collection
            colcolVouInfoItens.Add colVouInfoItens
            colVouInfoItens.Add objVouInfo
        End If
        
        'Busca Próximo
        lErro = Comando_BuscarProximo(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 198106
    
    Loop

    iIndice = 0
    For Each objVouInfo In colVouInfoGrupado
    
        iIndice = iIndice + 1
    
        dValorEstornar = objVouInfo.dValor
        dValorEstornarRelanc = objVouInfo.dValor
        If lNovoCliForn = 0 Then
            lNovoCliForn = objVouInfo.lCliForn
        End If
        If sTipoRelancamento = "" Then
            sTipoRelancamento = sTipoDocEstornado
        End If
       
        dValorEstornar = dValorEstornar * dFatorEstorno
                
        'Insere o estorno
        If dFatorEstorno = 1 Then
            lErro = TRPVou_Gera_Comis_Insere_Info(lComando2, lComando3, lComando4, objVou, sTipoDocEstornado, dtData, -dValorEstornar, sHistorico, objVouInfo.iStatus, objVouInfo.iTipoLiberacao, objVouInfo.lNumIntDocLiberacao, lNumIntDocOCR, objVouInfo.iTipoCliForn, objVouInfo.lCliForn, TRP_TIPO_ESTORNO_LANC_ESTORNADOR, 0, lContador, lNumIntDocEstorno)
        Else
            lErro = TRPVou_Gera_Comis_Insere_Info(lComando2, lComando3, lComando4, objVou, sTipoDocEstornado, dtData, -dValorEstornar, sHistorico, objVouInfo.iStatus, objVouInfo.iTipoLiberacao, objVouInfo.lNumIntDocLiberacao, lNumIntDocOCR, objVouInfo.iTipoCliForn, objVouInfo.lCliForn, TRP_TIPO_ESTORNO_LANC_NORMAL, 0, lContador)
            lNumIntDocEstorno = 0
        End If
        If lErro <> AD_SQL_SUCESSO Then gError 198103
                
        If iRelancarOutroTipo = MARCADO Then
        
            dValorEstornarRelanc = dValorEstornarRelanc * dFatorRelanc
        
            'Relança o documento com outro tipo
            lErro = TRPVou_Gera_Comis_Insere_Info(lComando2, lComando3, lComando4, objVou, sTipoRelancamento, dtData, dValorEstornarRelanc, sHistorico, objVouInfo.iStatus, objVouInfo.iTipoLiberacao, objVouInfo.lNumIntDocLiberacao, lNumIntDocOCR, objVouInfo.iTipoCliForn, lNovoCliForn, TRP_TIPO_ESTORNO_LANC_RELANCAMENTO, 0, lContador)
            If lErro <> AD_SQL_SUCESSO Then gError 198104
    
        End If
        
        Set colVouInfoItens = colcolVouInfoItens.Item(iIndice)
        
        If lNumIntDocEstorno <> 0 Then
        
            For Each objVouInfoAux In colVouInfoItens
                'Marca que o documento foi estornado e quem estornou
                lErro = Comando_Executar(lComando5, "UPDATE TRPVoucherInfo SET Estorno = ?, NumIntDocEstorno = ? WHERE NumIntDoc = ?", TRP_TIPO_ESTORNO_LANC_ESTORNADO, lNumIntDocEstorno, objVouInfoAux.lNumIntDoc)
                If lErro <> AD_SQL_SUCESSO Then gError 198105
            Next
            
        End If
        
    Next
 
    TRPVou_Gera_Comis_Trata_Estorno = SUCESSO

    Exit Function

Erro_TRPVou_Gera_Comis_Trata_Estorno:

    TRPVou_Gera_Comis_Trata_Estorno = gErr

    Select Case gErr
    
        Case 198100, 198101, 198106
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPVOUCHERINFO", gErr)

        Case 198102
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRPVOUCHERINFO", gErr)
            
        Case 198103, 198104

        Case 198105
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRPVOUCHERINFO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198107)

    End Select

    Exit Function

End Function

Private Function TRPVou_Gera_Comis_Trata_Estorno2(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lComando3 As Long, ByVal lComando4 As Long, ByVal lComando5 As Long, ByVal objVou As ClassTRPVouchers, ByVal dtData As Date, ByVal sHistorico As String, ByVal lNumIntDocOCR As Long, lContador As Long) As Long
'Gera a comissão externa relacionada ao voucher

Dim lErro As Long
Dim lNumIntDoc As Long
Dim dValorEstornar As Double
Dim sTipoDocEstornado
Dim iStatus As Integer
Dim iTipoLiberacao As Integer
Dim lNumIntDocLiberacao As Long
Dim iTipoCliForn As Integer
Dim lCliForn As Long
Dim lNumIntDocEstornador As Long

On Error GoTo Erro_TRPVou_Gera_Comis_Trata_Estorno2

    sTipoDocEstornado = String(STRING_TRP_VOU_INFO_TIPODOC, 0)

    'Le a TRPVOucherInfo para pegar o saldo do que foi lançado de comissão de cartão para estornar e lançar para agência
    lErro = Comando_ExecutarPos(lComando1, "SELECT NumIntDoc, Valor, TipoDoc, Status, TipoLiberacao, NumIntDocLiberacao, TipoCliForn, CliForn FROM TRPVoucherInfo WHERE TipVou = ? AND Serie = ? AND NumVou = ? AND NumIntDocOCR = ? ", 0, _
                 lNumIntDoc, dValorEstornar, sTipoDocEstornado, iStatus, iTipoLiberacao, lNumIntDocLiberacao, iTipoCliForn, lCliForn, objVou.sTipVou, objVou.sSerie, objVou.lNumVou, lNumIntDocOCR)
    If lErro <> AD_SQL_SUCESSO Then gError 198108

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 198109

    Do While lErro <> AD_SQL_SEM_DADOS
        
        lErro = Comando_LockExclusive(lComando1)
        If lErro <> AD_SQL_SUCESSO Then gError 198110

        'Relança o documento com outro tipo
        lErro = TRPVou_Gera_Comis_Insere_Info(lComando2, lComando3, lComando4, objVou, sTipoDocEstornado, dtData, -dValorEstornar, sHistorico, iStatus, iTipoLiberacao, lNumIntDocLiberacao, 0, iTipoCliForn, lCliForn, TRP_TIPO_ESTORNO_LANC_ESTORNADOR, lNumIntDoc, lContador, lNumIntDocEstornador)
        If lErro <> AD_SQL_SUCESSO Then gError 198111
        
        'Marca que o documento foi estornado e quem estornou
        lErro = Comando_ExecutarPos(lComando5, "UPDATE TRPVoucherInfo SET Estorno = ?, NumIntDocEstorno = ?", lComando1, TRP_TIPO_ESTORNO_LANC_ESTORNADO, lNumIntDocEstornador)
        If lErro <> AD_SQL_SUCESSO Then gError 198112
        
        'Busca Próximo
        lErro = Comando_BuscarProximo(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 198113
    
    Loop
 
    TRPVou_Gera_Comis_Trata_Estorno2 = SUCESSO

    Exit Function

Erro_TRPVou_Gera_Comis_Trata_Estorno2:

    TRPVou_Gera_Comis_Trata_Estorno2 = gErr

    Select Case gErr

        Case 198108, 198109, 198113
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPVOUCHERINFO", gErr)

        Case 198110
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRPVOUCHERINFO", gErr)
            
        Case 198111

        Case 198112
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRPVOUCHERINFO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198114)

    End Select

    Exit Function

End Function

Public Function TRPVou_Gera_Comis_Lanca_Comis(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lComando3 As Long, ByVal lComando4 As Long, ByVal objVou As ClassTRPVouchers, ByVal dValorBrutoBase As Double, ByVal dtData As Date, ByVal sHistorico As String, ByVal iStatus As Integer, ByVal iTipoLiberacao As Integer, ByVal lNumIntDocLiberacao As Long, lContador As Long, ByVal iStatusOver As Integer, ByVal iTipoLiberacaoOver As Integer, ByVal lNumIntDocLiberacaoOver As Long) As Long
'Gera a comissão externa relacionada ao voucher

Dim lErro As Long
Dim dValorCMR As Double
Dim dValorCMC As Double
Dim dValorOVER As Double
Dim sItem As String

On Error GoTo Erro_TRPVou_Gera_Comis_Lanca_Comis

    'CMR
    'Se tem representante
    If objVou.lRepresentante <> 0 Then
    
        dValorCMR = dValorBrutoBase * objVou.dComissaoRep
        
        lErro = TRPVou_Gera_Comis_Insere_Info(lComando1, lComando2, lComando3, objVou, TRP_TIPODOC_CMR_TEXTO, dtData, dValorCMR, sHistorico, iStatus, iTipoLiberacao, lNumIntDocLiberacao, 0, TRP_CLIENTEINFO_TIPO_CLIENTE, objVou.lRepresentante, TRP_TIPO_ESTORNO_LANC_NORMAL, 0, lContador)
        If lErro <> AD_SQL_SUCESSO Then gError 198115
        
    End If
        
    'CMC
    'Se tem representante
    If objVou.lCorrentista <> 0 Then
            
        dValorCMC = dValorBrutoBase * objVou.dComissaoCorr
            
        lErro = TRPVou_Gera_Comis_Insere_Info(lComando1, lComando2, lComando3, objVou, TRP_TIPODOC_CMC_TEXTO, dtData, dValorCMC, sHistorico, iStatus, iTipoLiberacao, lNumIntDocLiberacao, 0, TRP_CLIENTEINFO_TIPO_CLIENTE, objVou.lCorrentista, TRP_TIPO_ESTORNO_LANC_NORMAL, 0, lContador)
        If lErro <> AD_SQL_SUCESSO Then gError 198116
        
    End If
        
    'OVER (Não tem over quando for dias antecipados)
    If objVou.lEmissor <> 0 And objVou.iDiasAntc = DESMARCADO Then
    
        dValorOVER = dValorBrutoBase * objVou.dComissaoEmissor
                              
        lErro = TRPVou_Gera_Comis_Insere_Info(lComando1, lComando2, lComando3, objVou, TRP_TIPODOC_OVER_TEXTO, dtData, dValorOVER, sHistorico, iStatusOver, iTipoLiberacaoOver, lNumIntDocLiberacaoOver, 0, TRP_CLIENTEINFO_TIPO_FORNECEDOR, objVou.lEmissor, TRP_TIPO_ESTORNO_LANC_NORMAL, 0, lContador)
        If lErro <> AD_SQL_SUCESSO Then gError 198119
    
    End If
    
    TRPVou_Gera_Comis_Lanca_Comis = SUCESSO

    Exit Function

Erro_TRPVou_Gera_Comis_Lanca_Comis:

    TRPVou_Gera_Comis_Lanca_Comis = gErr

    Select Case gErr
    
        Case 198115, 198116, 198119
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198120)

    End Select

    Exit Function

End Function

Public Function TRPAportePagtoFatCond_Exclui(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal objTRPAportePagtoFatCond As ClassTRPAportePagtoFatC) As Long

Dim lErro As Long
Dim lNumIntDoc As Long

On Error GoTo Erro_TRPAportePagtoFatCond_Exclui

    'Le a tabelaTRPAportePagtoFatCond
    lErro = Comando_ExecutarPos(lComando1, "SELECT NumIntDoc FROM TRPAportePagtoFatCond WHERE NumIntDoc= ? ", _
                0, lNumIntDoc, objTRPAportePagtoFatCond.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 190416

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190417

    'Se não existir => ERRO
    If lErro <> AD_SQL_SUCESSO Then gError 190418

    lErro = Comando_ExecutarPos(lComando2, "DELETE FROM TRPAportePagtoFatCond ", lComando1)
    If lErro <> AD_SQL_SUCESSO Then gError 190419
    
    TRPAportePagtoFatCond_Exclui = SUCESSO

    Exit Function

Erro_TRPAportePagtoFatCond_Exclui:

    TRPAportePagtoFatCond_Exclui = gErr

    Select Case gErr

        Case 190416, 190417
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPAPORTEPAGTOFATCOND", gErr)

        Case 190418
            Call Rotina_Erro(vbOKOnly, "ERRO_TRPAPORTEPAGTOFATCOND_NAO_CADASTRADO", gErr, objTRPAportePagtoFatCond.lNumIntDoc)
        
        Case 190419
            Call Rotina_Erro(vbOKOnly, "ERRO_DELETE_TRPAPORTEPAGTOFATCOND", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190420)

    End Select

    Exit Function

End Function

Public Function TRPAportePagtoFatCond_Atualiza(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal objTRPAportePagtoFatCond As ClassTRPAportePagtoFatC, ByVal objAporte As ClassTRPAportes) As Long

Dim lErro As Long
Dim lNumIntDoc As Long

On Error GoTo Erro_TRPAportePagtoFatCond_Atualiza

    'Le a tabelaTRPAportePagtoFatCond
    lErro = Comando_ExecutarPos(lComando1, "SELECT NumIntDoc FROM TRPAportePagtoFatCond WHERE NumIntDoc= ? ", _
                0, lNumIntDoc, objTRPAportePagtoFatCond.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 190411

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190412

    'Se não existir => ERRO
    If lErro <> AD_SQL_SUCESSO Then gError 190413

    lErro = Comando_ExecutarPos(lComando2, "UPDATE TRPAportePagtoFatCond SET Valor= ? , ValidadeDe= ? , ValidadeAte= ? , " & _
            "Percentual= ?, Seq = ? ", lComando1, _
            objTRPAportePagtoFatCond.dValor, objTRPAportePagtoFatCond.dtValidadeDe, objTRPAportePagtoFatCond.dtValidadeAte, _
            objTRPAportePagtoFatCond.dPercentual, objTRPAportePagtoFatCond.iSeq)
    If lErro <> AD_SQL_SUCESSO Then gError 190414
    
    TRPAportePagtoFatCond_Atualiza = SUCESSO

    Exit Function

Erro_TRPAportePagtoFatCond_Atualiza:

    TRPAportePagtoFatCond_Atualiza = gErr

    Select Case gErr

        Case 190411, 190412
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPAPORTEPAGTOFATCOND", gErr)

        Case 190413
            Call Rotina_Erro(vbOKOnly, "ERRO_TRPAPORTEPAGTOFATCOND_NAO_CADASTRADO", gErr, objTRPAportePagtoFatCond.lNumIntDoc)
        
        Case 190414
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRPAPORTEPAGTOFATCOND", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190415)

    End Select

    Exit Function

End Function

Public Function TRPAportePagtoFatCond_Insere(ByVal lComando As Long, ByVal objTRPAportePagtoFatCond As ClassTRPAportePagtoFatC, ByVal objAporte As ClassTRPAportes) As Long

Dim lErro As Long
Dim lNumIntDoc As Long

On Error GoTo Erro_TRPAportePagtoFatCond_Insere

    'Obter NumIntDoc
    lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_INT_PROX_TRPAPORTEPAGTOFATCOND", lNumIntDoc)
    If lErro <> SUCESSO Then gError 190408

    objTRPAportePagtoFatCond.lNumIntDoc = lNumIntDoc

    lErro = Comando_Executar(lComando, "INSERT INTO TRPAportePagtoFatCond( NumIntDoc, NumIntDocAporte, Valor, ValidadeDe, " & _
            "ValidadeAte, Percentual, Seq) VALUES (?,?,?,?,?,?,?)", _
            objTRPAportePagtoFatCond.lNumIntDoc, objTRPAportePagtoFatCond.lNumIntDocAporte, objTRPAportePagtoFatCond.dValor, objTRPAportePagtoFatCond.dtValidadeDe, _
            objTRPAportePagtoFatCond.dtValidadeAte, objTRPAportePagtoFatCond.dPercentual, objTRPAportePagtoFatCond.iSeq)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190409

    TRPAportePagtoFatCond_Insere = SUCESSO

    Exit Function

Erro_TRPAportePagtoFatCond_Insere:

    TRPAportePagtoFatCond_Insere = gErr

    Select Case gErr

        Case 190408

        Case 190409
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRPAPORTEPagtoFatCond", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190410)

    End Select

    Exit Function

End Function

Public Function TRPAportes_Valida(ByVal objTRPAportes As ClassTRPAportes) As Long

Dim lErro As Long
Dim alComando(0 To 0) As Long
Dim iIndice As Integer
Dim iIndice2 As Integer
Dim objPagtoFat As ClassTRPAportePagtoFat
Dim objPagtoFatAux As ClassTRPAportePagtoFat
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

On Error GoTo Erro_TRPAportes_Valida

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 196480
    Next
    
    iIndice = 0
    For Each objPagtoFat In objTRPAportes.colPagtoSobreFatura
        bValidadeRepetida = False
        iIndice = iIndice + 1
        For iIndice2 = iIndice + 1 To objTRPAportes.colPagtoSobreFatura.Count
            Set objPagtoFatAux = objTRPAportes.colPagtoSobreFatura.Item(iIndice2)
            If Not ((objPagtoFatAux.dtValidadeDe < objPagtoFat.dtValidadeDe And objPagtoFatAux.dtValidadeAte < objPagtoFat.dtValidadeDe) Or (objPagtoFatAux.dtValidadeDe > objPagtoFat.dtValidadeAte And objPagtoFatAux.dtValidadeAte > objPagtoFat.dtValidadeAte)) Then
                bValidadeRepetida = True
                sTexto = sTexto & " e Aporte: " & objTRPAportes.lCodigo & " Valor: " & Format(objPagtoFatAux.dValor, "STANDARD") & " Saldo: " & Format(objPagtoFatAux.dSaldo, "STANDARD") & " Validade De: " & Format(objPagtoFatAux.dtValidadeDe, "dd/mm/yyyy") & " Validade De: " & Format(objPagtoFatAux.dtValidadeAte, "dd/mm/yyyy") & " Linha : " & CStr(objPagtoFatAux.iSeq)
            End If
        Next
    
        'Le a tabelaTRPAportes
        lErro = Comando_Executar(alComando(0), "SELECT A.Codigo, P.Valor, P.Saldo, P.Validadede, P.ValidadeAte, P.Seq FROM TRPAportes AS A, TRPAportePagtoFat AS P WHERE A.NumintDoc = P.NumIntDocAporte AND NOT ((ValidadeDe < ? AND ValidadeAte < ?) OR (ValidadeDe > ? AND ValidadeAte > ?)) AND Cliente = ? AND Codigo <> ? ", lCodigo, dValor, dSaldo, dtValidadeDe, dtValidadeAte, iSeq, objPagtoFat.dtValidadeDe, objPagtoFat.dtValidadeDe, objPagtoFat.dtValidadeAte, objPagtoFat.dtValidadeAte, objTRPAportes.lCliente, objTRPAportes.lCodigo)
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
            sTextoFinal = sTextoFinal & " Aporte: " & objTRPAportes.lCodigo & " Valor: " & Format(objPagtoFat.dValor, "STANDARD") & " Saldo: " & Format(objPagtoFat.dSaldo, "STANDARD") & " Validade De: " & Format(objPagtoFat.dtValidadeDe, "dd/mm/yyyy") & " Validade De: " & Format(objPagtoFat.dtValidadeAte, "dd/mm/yyyy") & " Linha : " & CStr(objPagtoFat.iSeq) & sTexto
        End If
    
    Next
    
    If Len(Trim(sTextoFinal)) > 0 Then
    
        vbResult = Rotina_Aviso(vbYesNo, "AVISO_PAGTOAPORTE_VALIDADEREPETIDA", objTRPAportes.lCliente, sTextoFinal)
        If vbResult = vbNo Then gError 196484
    
    End If
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TRPAportes_Valida = SUCESSO

    Exit Function

Erro_TRPAportes_Valida:

    TRPAportes_Valida = gErr

    Select Case gErr

        Case 196480
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 196481 To 196483
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPAPORTES", gErr)
        
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

Private Function NF_NFiscal_Grava_TRP(alComando() As Long, ByVal objNFiscal As ClassNFiscal, iProxComando As Integer) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim objItemNF As ClassItemNF
Dim tSerie As typeSerie
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim objFilCliFilEmp As New ClassFilCliFilEmp
Dim objRPS As New ClassRPS
Dim objNFTRP As ClassNFTRP
Dim tFilCliFilEmp As typeFilCliFilEmp

On Error GoTo Erro_NF_NFiscal_Grava_TRP
    
    Set objNFTRP = objNFiscal.objInfoUsu

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
    
    Set objTipoDocInfo = objNFTRP.objTipoDocInfo
    
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
        lErro = NF_SldDiaFat_Grava_Est_TRP(alComando(iProxComando + 17), alComando(iProxComando + 18), alComando(iProxComando + 19), objNFiscal, objTipoDocInfo)
        If lErro <> SUCESSO Then gError 196810
        
        'Grava a estatística do Mês do produto
        lErro = NF_SldMesFat_Grava_Est_TRP(alComando(iProxComando + 20), alComando(iProxComando + 21), alComando(iProxComando + 22), alComando(iProxComando + 23), alComando(iProxComando + 24), alComando(iProxComando + 25), objNFiscal, objTipoDocInfo)
        If lErro <> SUCESSO Then gError 196811
    
    End If

    'Verifica se o modulo de Livros Fiscais está Ativo
    If gcolModulo.Ativo(MODULO_LIVROSFISCAIS) = MODULO_ATIVO Then
                        
        lErro = NF_NotaFiscal_Grava_Fis_TRP(alComando(iProxComando + 26), alComando(iProxComando + 27), alComando(iProxComando + 28), alComando(iProxComando + 29), alComando(iProxComando + 30), alComando(iProxComando + 31), alComando(iProxComando + 32), alComando(iProxComando + 33), alComando(iProxComando + 36), alComando(iProxComando + 37), alComando(iProxComando + 39), alComando(iProxComando + 40), alComando(iProxComando + 41), alComando(iProxComando + 42), objNFiscal)
        If lErro <> SUCESSO Then gError 196812
    
    End If
    
    If objNFiscal.iRecibo = MARCADO Then

        lErro = NF_Converte_RPS_TRP(objNFiscal, objRPS)
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
                    "Cidade, UF, CEP, Email, Discriminacao, NumIntDocNF, Cliente, FilialCliente,DataUltAlteracao,HoraUltAlteracao, Enviado, ValorCofins, ValorCSLL, ValorINSS, ValorIRPJ, ValorPIS) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                    objRPS.lNumIntDoc, objRPS.sTipo, objRPS.sSerie, objRPS.lNumero, objRPS.dtDataEmissao, objRPS.iFilialEmpresa, _
                    objRPS.sSituacao, objRPS.dValorServico, objRPS.dValorDeducao, objRPS.lCodigoServico, objRPS.dAliquota, _
                    objRPS.iISSRetido, objRPS.iTipoCGC, objRPS.sCgc, objRPS.sInscricaoMunicipal, objRPS.sInscricaoEstadual, _
                    objRPS.sRazaoSocial, objRPS.sEndereco, objRPS.sEndNumero, objRPS.sEndComplemento, objRPS.sBairro, _
                    objRPS.sCidade, objRPS.sUF, objRPS.sCEP, objRPS.sEmail, objRPS.sDiscriminacao, objRPS.lNumIntDocNF, objRPS.lCliente, objRPS.iFilialCliente, Date, CDbl(Time), DESMARCADO, objRPS.dValorCofins, objRPS.dValorCSLL, objRPS.dValorINSS, objRPS.dValorIRPJ, objRPS.dValorPIS)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then
                gError 196815
            End If
            
        End If
        
    End If
    
    iProxComando = iProxComando + 43
    
    NF_NFiscal_Grava_TRP = SUCESSO

    Exit Function

Erro_NF_NFiscal_Grava_TRP:

    NF_NFiscal_Grava_TRP = gErr

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
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RPS2", gErr, objRPS.lNumero, objRPS.lCliente)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196816)

    End Select

    Exit Function
    
End Function

Private Function NF_SldDiaFat_Grava_Est_TRP(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lComando3 As Long, ByVal objNFiscal As ClassNFiscal, ByVal objTipoDocInfo As ClassTipoDocInfo) As Long
'Grava a estatística do Dia do produto

Dim lErro As Long
Dim objItemNF As New ClassItemNF
Dim objSldDiaFat As New ClassSldDiaFat
Dim objProduto As New ClassProduto
Dim iIndice As Integer

On Error GoTo Erro_NF_SldDiaFat_Grava_Est_TRP
            
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
    
    NF_SldDiaFat_Grava_Est_TRP = SUCESSO
    
    Exit Function
    
Erro_NF_SldDiaFat_Grava_Est_TRP:
    
    NF_SldDiaFat_Grava_Est_TRP = gErr

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
        
Private Function NF_SldMesFat_Grava_Est_TRP(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lComando3 As Long, ByVal lComando4 As Long, ByVal lComando5 As Long, ByVal lComando6 As Long, ByVal objNFiscal As ClassNFiscal, ByVal objTipoDocInfo As ClassTipoDocInfo) As Long
'Grava a estatística do Mês do produto

Dim lErro As Long
Dim objItemNF As ClassItemNF
Dim objSldMesFat As New ClassSldMesFat
Dim iMes As Integer
Dim sMes As String
Dim iIndice As Integer

On Error GoTo Erro_NF_SldMesFat_Grava_Est_TRP
            
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
            
            lErro = NF_SldMesFatFilCli_Grava_Est_TRP(lComando4, lComando5, lComando6, objSldMesFat.iFilialEmpresa, objSldMesFat.iAno, objSldMesFat.sProduto, objNFiscal.lCliente, objNFiscal.iFilialCli, objSldMesFat.dQuantFaturada(iMes), objSldMesFat.dValorFaturado(iMes), 0, 0, sMes)
            If lErro <> SUCESSO Then gError 196831
            
        End If
        
    Next
            
    NF_SldMesFat_Grava_Est_TRP = SUCESSO
    
    Exit Function
    
Erro_NF_SldMesFat_Grava_Est_TRP:
    
    NF_SldMesFat_Grava_Est_TRP = gErr

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

Private Function NF_SldMesFatFilCli_Grava_Est_TRP(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lComando3 As Long, iFilialEmpresa As Integer, iAno As Integer, sCodProduto As String, lCliente As Long, iFilialCliente As Integer, dQtdFaturada As Double, dValorFaturado As Double, dQtdDevolvida As Double, dValorDevolvido As Double, sMes As String)

Dim lErro As Long
Dim iFilial As Integer

On Error GoTo Erro_NF_SldMesFatFilCli_Grava_Est_TRP
    
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

    NF_SldMesFatFilCli_Grava_Est_TRP = SUCESSO
    
    Exit Function
    
Erro_NF_SldMesFatFilCli_Grava_Est_TRP:

    NF_SldMesFatFilCli_Grava_Est_TRP = gErr
    
    Select Case gErr
  
        Case 196833, 196834
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESFATFILCLI", gErr, iAno, iFilialEmpresa, sCodProduto, lCliente, iFilialCliente)
        
        Case 196835
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_SLDMESFATFILCLI", gErr, iAno, iFilialEmpresa, sCodProduto, lCliente, iFilialCliente)
        
        Case 196836
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_SLDMESFATFILCLI", gErr, iAno, iFilialEmpresa, sCodProduto, lCliente, iFilialCliente)
        
        Case 196837
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SLDMESFATFILCLI", gErr, iAno, iFilialEmpresa, sCodProduto, lCliente, iFilialCliente)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196838)

    End Select
    
    Exit Function

End Function

Private Function NF_NotaFiscal_Grava_Fis_TRP(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lComando3 As Long, ByVal lComando4 As Long, ByVal lComando5 As Long, ByVal lComando6 As Long, ByVal lComando7 As Long, ByVal lComando8 As Long, ByVal lComando9 As Long, ByVal lComando10 As Long, ByVal lComando11 As Long, ByVal lComando12 As Long, ByVal lComando13 As Long, ByVal lComando14 As Long, ByVal objNFiscal As ClassNFiscal) As Long
'Grava nas Tabelas de Livros Fiscais a Nota Fiscal

Dim lErro As Long
Dim objLivRegES As New ClassLivRegES
Dim objLivRegESEmitentes As New ClassLivRegESEmitentes
Dim colLivRegESItemNF As New Collection
Dim colLivRegESLinha As New Collection
Dim objNFTRP As ClassNFTRP

On Error GoTo Erro_NF_NotaFiscal_Grava_Fis_TRP

    Set objNFTRP = objNFiscal.objInfoUsu
    
    'só vou gerar info para notas que interessem p/livros de ICMS/IPI
    If objNFTRP.objTipoDocInfo.iModeloArqICMS <> 0 Then
  
        'Gera o Registro de LivRegES
        lErro = NF_Saida_Grava_LivRegES_TRP(lComando1, lComando2, lComando3, objNFiscal, objLivRegES, objNFTRP.objTipoDocInfo, objNFTRP.objCliente, objNFTRP.objFilialCliente, objNFTRP.objEnderecoCli)
        If lErro <> SUCESSO Then gError 196838
        
        'Gera as linhas para livros fiscais
        lErro = NF_NotaFiscal_GeraLinhas_Fis_TRP(lComando4, lComando5, lComando6, lComando7, lComando8, lComando11, lComando12, objNFTRP.objTipoDocInfo.iModeloArqICMS, colLivRegESLinha, colLivRegESItemNF, objNFiscal, objLivRegES)
        If lErro <> SUCESSO Then gError 196839
        
        'Grava as Coleções, que estão ordenadas e completadas com o Complemento
        lErro = NF_Grava_LivRegES_Colecao_TRP(lComando9, lComando10, lComando13, lComando14, colLivRegESLinha, colLivRegESItemNF)
        If lErro <> SUCESSO Then gError 196840

    End If
    
    NF_NotaFiscal_Grava_Fis_TRP = SUCESSO
    
    Exit Function
    
Erro_NF_NotaFiscal_Grava_Fis_TRP:

    NF_NotaFiscal_Grava_Fis_TRP = gErr
    
    Select Case gErr
        
        Case 196838, 196839, 196840
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196841)

    End Select
    
    Exit Function
        
End Function

Private Function NF_Saida_Grava_LivRegES_TRP(ByVal lComando As Long, ByVal lComando2 As Long, ByVal lComando3 As Long, ByVal objNFiscal As ClassNFiscal, ByVal objLivRegES As ClassLivRegES, ByVal objTipoDocInfo As ClassTipoDocInfo, ByVal objCliente As ClassCliente, ByVal objFilialCliente As ClassFilialCliente, ByVal objEndereco As ClassEndereco) As Long
'Preenche o objLivRegES com os dados da Nota Fiscal

Dim lErro As Long
Dim lNumIntDoc As Long
Dim objFilialEmpresa As New AdmFiliais
Dim sInscricaoAux As String

On Error GoTo Erro_NF_Saida_Grava_LivRegES_TRP
    
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
                                
    NF_Saida_Grava_LivRegES_TRP = SUCESSO
    
    Exit Function
    
Erro_NF_Saida_Grava_LivRegES_TRP:

    NF_Saida_Grava_LivRegES_TRP = gErr

    Select Case gErr

        Case 196842
        
        Case 196843
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_LIVREGES", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196844)

    End Select
    
    Exit Function

End Function

Private Function NF_NotaFiscal_GeraLinhas_Fis_TRP(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lComando3 As Long, ByVal lComando4 As Long, ByVal lComando5 As Long, ByVal lComando6 As Long, ByVal lComando7 As Long, ByVal iModeloArqICMS As Integer, ByVal colLivRegESLinha As Collection, ByVal colLivRegESItemNF As Collection, ByVal objNFiscal As ClassNFiscal, ByVal objLivRegES As ClassLivRegES) As Long
'Gera as linhas para livros fiscais de acordo com o valor passado iModeloArqICMS
'Em geral, iModeloArqICMS será 1, o que indica que as linhas serão geradas no Case Else
'Caso o valor seja diferente de 1, significa que será gerada uma linha "especial"

Dim lErro As Long
Dim objItemNF As New ClassItemNF
Dim objLivRegESCadProd As New ClassLivRegESCadProd
Dim objLivRegESLinha As ClassLivRegESLinha
Dim objNFiscalOrig As New ClassNFiscal

On Error GoTo Erro_NF_NotaFiscal_GeraLinhas_Fis_TRP
            
    'Varre os Itens para gerar LivRegLinha, LivRegESNF, LivRegESCadProd
    For Each objItemNF In objNFiscal.ColItensNF
        
        'Grava os dados da Nota para a coleção colLivRegESLinha
        lErro = NF_Move_Tela_LivRegESLinha_Obj_TRP(objNFiscal, objItemNF.objTributacaoItemNF, objItemNF.dValorTotal, colLivRegESLinha, objLivRegES, objItemNF.dValorDesconto, objItemNF.dValorDescGlobal)
        If lErro <> SUCESSO Then gError 196845
        
        'Move os dados do Item NF para a coleção colLivRegESCadProd
        lErro = NF_Grava_LivRegESCadProd_TRP(lComando1, lComando2, lComando3, lComando4, lComando6, objLivRegES.dtData, objItemNF, objLivRegESCadProd, objNFiscal)
        If lErro <> SUCESSO Then gError 196846
        
        'Move os dados do Item NF para a coleção colLivRegESItemNF
        lErro = NF_Move_Memoria_LivRegESItemNF_TRP(lComando5, lComando7, objItemNF, objNFiscal, objLivRegES, objLivRegESCadProd, colLivRegESItemNF)
        If lErro <> SUCESSO Then gError 196847
                
    Next

    NF_NotaFiscal_GeraLinhas_Fis_TRP = SUCESSO
    
    Exit Function

Erro_NF_NotaFiscal_GeraLinhas_Fis_TRP:

    NF_NotaFiscal_GeraLinhas_Fis_TRP = gErr
    
    Select Case gErr
    
        Case 196845 To 196847
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196848)
    
    End Select
    
    Exit Function
        
End Function

Private Function NF_Grava_LivRegES_Colecao_TRP(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lComando3 As Long, ByVal lComando4 As Long, ByVal colLivRegESLinha As Collection, ByVal colLivRegESItemNF As Collection) As Long

Dim lErro As Long
Dim objLivRegESItemNF As New ClassLivRegESItemNF
Dim objLivRegESLinha As New ClassLivRegESLinha
Dim lNumIntDoc As Long

On Error GoTo Erro_NF_Grava_LivRegES_Colecao_TRP
    
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
    
    NF_Grava_LivRegES_Colecao_TRP = SUCESSO
    
    Exit Function
    
Erro_NF_Grava_LivRegES_Colecao_TRP:

    NF_Grava_LivRegES_Colecao_TRP = gErr
    
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

Private Function NF_Move_Tela_LivRegESLinha_Obj_TRP(ByVal objNFiscal As ClassNFiscal, ByVal objTributacaoObj As Object, ByVal dValorObj As Double, ByVal colLivRegESLinha As Collection, ByVal objLivRegES As ClassLivRegES, ByVal dValorDesconto As Double, ByVal dValorDescGlobal As Double) As Long

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
Dim objNFTRP As ClassNFTRP

On Error GoTo Erro_NF_Move_Tela_LivRegESLinha_Obj_TRP

    Set objNFTRP = objNFiscal.objInfoUsu
    Set objTipoTribICMS = objNFTRP.objTipoTribICMS
    Set objTipoTribIPI = objNFTRP.objTipoTribIPI
    Set objTipoDeTributacao = objNFTRP.objTipoDeTributacaoMovto

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
        If objLivRegESLinha.sNaturezaOp = objTributacaoObj.sNaturezaOp And objLivRegESLinha.dAliquotaICMS = dAliquotaICMS And objLivRegESLinha.dAliquotaIPI = dAliquotaIPI And objLivRegESLinha.sClassifContabil = objNFTRP.objProduto.sContaContabil Then
            
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
        objLivRegESLinha.sClassifContabil = objNFTRP.objProduto.sContaContabil
        
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
    
    NF_Move_Tela_LivRegESLinha_Obj_TRP = SUCESSO
     
    Exit Function
    
Erro_NF_Move_Tela_LivRegESLinha_Obj_TRP:

    NF_Move_Tela_LivRegESLinha_Obj_TRP = gErr
     
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196853)
     
    End Select
     
    Exit Function
    
End Function

Private Function NF_Grava_LivRegESCadProd_TRP(ByVal lComando As Long, ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lComando3 As Long, ByVal lComando4 As Long, ByVal dtData As Date, ByVal objItemNF As ClassItemNF, objLivRegESCadProd As ClassLivRegESCadProd, ByVal objNF As ClassNFiscal) As Long
'Preenhe o objLivRegES com os dados do item da Nota Fiscal

Dim lErro As Long
Dim lNumIntDoc As Long
Dim objProduto As New ClassProduto
Dim objTipoTribICMS As New ClassTipoTribICMS
Dim dtDataInicialAnterior As Date
Dim sCodigoIpiAux As String
Dim objNFTRP As ClassNFTRP

On Error GoTo Erro_NF_Grava_LivRegESCadProd_TRP

    Set objNFTRP = objNF.objInfoUsu

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
        
        objLivRegESCadProd.sDescricao = left(objItemNF.sDescricaoItem, 50)
        objLivRegESCadProd.sProduto = objItemNF.sProduto
        objLivRegESCadProd.sSiglaUM = objItemNF.sUnidadeMed
        
        Set objProduto = objNFTRP.objProduto
        Set objTipoTribICMS = objNFTRP.objTipoTribICMS
            
        objLivRegESCadProd.sCodigoNCM = Mid(objProduto.sIPICodigo, 1, 8)
        
        objLivRegESCadProd.sSituacaoTrib = CStr(objProduto.iOrigemMercadoria) & left(CStr(objTipoTribICMS.iTipoTribCST), 1)
        
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
    
    NF_Grava_LivRegESCadProd_TRP = SUCESSO
    
    Exit Function
    
Erro_NF_Grava_LivRegESCadProd_TRP:

    NF_Grava_LivRegESCadProd_TRP = gErr

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

Private Function NF_Move_Memoria_LivRegESItemNF_TRP(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal objItemNF As ClassItemNF, ByVal objNFiscal As ClassNFiscal, ByVal objLivRegES As ClassLivRegES, ByVal objLivRegESCadProd As ClassLivRegESCadProd, ByVal colLivRegESItemNF As Collection) As Long
'Preenhe o objLivRegES com os dados do item da Nota Fiscal

Dim lErro As Long
Dim lNumIntDoc As Long
Dim objLivRegESItemNF As New ClassLivRegESItemNF

On Error GoTo Erro_NF_Move_Memoria_LivRegESItemNF_TRP

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
    
    NF_Move_Memoria_LivRegESItemNF_TRP = SUCESSO
    
    Exit Function
    
Erro_NF_Move_Memoria_LivRegESItemNF_TRP:

    NF_Move_Memoria_LivRegESItemNF_TRP = gErr

    Select Case gErr
        
        Case 196860
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196861)

    End Select
    
    Exit Function

End Function

Public Function NF_Converte_RPS_TRP(ByVal objNF As ClassNFiscal, ByVal objRPS As ClassRPS) As Long

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
Dim objNFTRP As ClassNFTRP
Dim objCliente As ClassCliente
Dim objFilialCliente As ClassFilialCliente
Dim objEndereco As ClassEndereco
Dim sStringAux As String
Dim dAliquotaPadrao As Double, bPossuiAliqDif As Boolean, dAliquotaMedia As Double

On Error GoTo Erro_NF_Converte_RPS_TRP

    Set objNFTRP = objNF.objInfoUsu

    objRPS.dtDataEmissao = objNF.dtDataEmissao
    objRPS.iFilialCliente = objNF.iFilialCli
    objRPS.iFilialEmpresa = objNF.iFilialEmpresa
    objRPS.lCliente = objNF.lCliente
    objRPS.lNumero = objNF.lNumNotaFiscal
    objRPS.lNumIntDocNF = objNF.lNumIntDoc
    objRPS.sSerie = objNF.sSerie
    
    Set objProduto = objNFTRP.objProduto

    sMensagem = ""
    
    dAliquotaPadrao = -1
    dAliquotaMedia = 0
    bPossuiAliqDif = False
    bTemProdutoNaoServico = False
    For Each objItemNF In objNF.ColItensNF

        If objProduto.iNatureza = NATUREZA_PROD_SERVICO Then
        
            If dAliquotaPadrao = -1 Then
                dAliquotaPadrao = objItemNF.objTributacao.dISSAliquota
            Else
                If Abs(dAliquotaPadrao = objItemNF.objTributacao.dISSAliquota) > DELTA_VALORMONETARIO Then
                    bPossuiAliqDif = True
                End If
            End If
            dAliquotaMedia = dAliquotaMedia + objItemNF.objTributacao.dISSAliquota * objItemNF.dValorTotal
        
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
    
    dAliquotaMedia = dAliquotaMedia / dValorServico
    
    objRPS.sDiscriminacao = sMensagem

    If bPossuiAliqDif Then
        objRPS.dAliquota = dAliquotaMedia
    Else
        objRPS.dAliquota = dAliquotaPadrao
    End If
    objRPS.dValorDeducao = 0
    objRPS.dValorServico = dValorServico
    
    objRPS.dValorCofins = objNF.objTributacaoNF.dCOFINSRetido
    objRPS.dValorCSLL = objNF.objTributacaoNF.dCSLLRetido
    objRPS.dValorINSS = IIf(objNF.objTributacaoNF.iINSSRetido = MARCADO, objNF.objTributacaoNF.dValorINSS, 0)
    objRPS.dValorIRPJ = objNF.objTributacaoNF.dIRRFValor
    objRPS.dValorPIS = objNF.objTributacaoNF.dPISRetido
    
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
                    
    Set objCliente = objNFTRP.objCliente
    Set objFilialCliente = objNFTRP.objFilialCliente
    Set objEndereco = objNFTRP.objEnderecoCli
    
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
    
    NF_Converte_RPS_TRP = SUCESSO

    Exit Function

Erro_NF_Converte_RPS_TRP:

    NF_Converte_RPS_TRP = gErr

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

Function FAT_TituloReceber_Insere_TRP(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lComando3 As Long, ByVal lComando4 As Long, ByVal lComando5 As Long, ByVal lComando6 As Long, ByVal lComando7 As Long, ByVal lComando8 As Long, ByVal lComando9 As Long, ByVal lComando10 As Long, ByVal objTituloReceber As ClassTituloReceber, ByVal colParcelaReceber As colParcelaReceber) As Long
'Insere o Título à Receber na Tabela de Titulos a Receber
    
Dim lErro As Long
Dim lNumIntDoc As Long
Dim objParcelaReceber As ClassParcelaReceber
Dim lNumeroFatura As Long

On Error GoTo Erro_FAT_TituloReceber_Insere_TRP
    
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
    lErro = FAT_ParcelasReceber_Insere_TRP(lComando6, lComando7, lComando8, lComando9, lComando10, colParcelaReceber)
    If lErro <> SUCESSO Then gError 196895
    
    FAT_TituloReceber_Insere_TRP = SUCESSO
    
    Exit Function
    
Erro_FAT_TituloReceber_Insere_TRP:

    FAT_TituloReceber_Insere_TRP = gErr
    
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

Private Function FAT_ParcelasReceber_Insere_TRP(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lComando3 As Long, ByVal lComando4 As Long, ByVal lComando5 As Long, ByVal colParcelaReceber As colParcelaReceber, Optional ByVal iNumPrimeiraParcela As Integer = 1) As Long
'Insere as Parcelas de um Titulo passadas por parâmetro
'iNumPrimeiraParcela é util na alteracao de titulo onde apenas as parcelas novas sao inseridas

Dim lErro As Long
Dim iIndice As Integer
Dim objCarteiraCobrador As New ClassCarteiraCobrador
Dim lNumIntDoc As Long
Dim lNumIntCom As Long
Dim colDesconto As colDesconto

On Error GoTo Erro_FAT_ParcelasReceber_Insere_TRP

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
    
    FAT_ParcelasReceber_Insere_TRP = SUCESSO
    
    Exit Function

Erro_FAT_ParcelasReceber_Insere_TRP:

    FAT_ParcelasReceber_Insere_TRP = gErr
    
    Select Case gErr
    
        Case 196897, 196899
        
        Case 196898
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PARCELAS_REC", gErr, colParcelaReceber(iIndice).lNumIntTitulo, iIndice)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196900)
            
    End Select
    
    Exit Function
    
End Function

Private Function TRP_Cliente_Grava_Fornecedor(ByVal objCli As ClassCliente, ByVal objEndereco As ClassEndereco) As Long

Dim lErro As Long
Dim objForn As New ClassFornecedor
Dim objEnderecoForn As New ClassEndereco

On Error GoTo Erro_TRP_Cliente_Grava_Fornecedor

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
    
    TRP_Cliente_Grava_Fornecedor = SUCESSO
    
    Exit Function

Erro_TRP_Cliente_Grava_Fornecedor:

    TRP_Cliente_Grava_Fornecedor = gErr
    
    Select Case gErr
    
        Case 196901, 196902

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196903)
            
    End Select
    
    Exit Function
    
End Function

Private Function TRPFaturamento_Gera_TitPagOver(ByVal objInfoFatura As ClassTRPFATInfoFatura, ByVal objTitPag As ClassTituloPagar, ByVal lProxNum As Long) As Long

Dim lErro As Long
Dim objParcelaPagar As New ClassParcelaPagar
Dim colParcelaPagar As New colParcelaPagar
Dim objContabil As ClassContabil
Dim objCtb As New ClassTRPCTBFatCP
Dim objContabAutomatica As New ClassContabAutomatica
Dim dtDataVenctoReal As Date
Dim dtDataVencimento As Date

On Error GoTo Erro_TRPFaturamento_Gera_TitPagOver
               
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
                    
    TRPFaturamento_Gera_TitPagOver = SUCESSO

    Exit Function

Erro_TRPFaturamento_Gera_TitPagOver:

    TRPFaturamento_Gera_TitPagOver = gErr

    Select Case gErr
    
        Case 190997 To 190999, 192000

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190991)

    End Select

    Exit Function

End Function

Public Function TRPVouchers_Comissao_Deslibera(ByVal iTipoLiberacao As Integer, ByVal lNumIntDocLiberacao As Long) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim tVouInfo As typeTRPVoucherInfo

On Error GoTo Erro_TRPVouchers_Comissao_Deslibera

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 196386
    Next
    
    'Alocação de espaço no buffer
    tVouInfo.sTipoDoc = String(STRING_TRP_VOU_INFO_TIPODOC, 0)
    tVouInfo.sTipVou = String(STRING_TRP_VOU_INFO_TIPVOU, 0)
    tVouInfo.sSerie = String(STRING_TRP_VOU_INFO_SERIE, 0)
    tVouInfo.sHistorico = String(STRING_TRP_VOU_HISTORICO, 0)

    'Le a tabelaTRPVoucherInfo para pegar todos registros liberados e não faturados
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc, TipoDoc, TipVou, Serie, NumVou, Seq, " & _
                "Data, TipoDocDestino, NumIntDocDestino, NumIntDocComiInt, Valor, Historico, NumTitulo, Status FROM TRPVoucherInfo WHERE TipoLiberacao = ? AND NumIntDocLiberacao = ? AND NumIntDocDestino = ? ", 0, _
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

        lErro = Comando_ExecutarPos(alComando(1), "UPDATE TRPVoucherInfo SET Status = ?, TipoLiberacao = ?, NumIntDocLiberacao = ?", alComando(0), TRP_VOU_INFO_STATUS_BLOQUEADO, 0, 0)
        If lErro <> AD_SQL_SUCESSO Then gError 196390
        
        'Busca Próximo
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196391
        
    Loop
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TRPVouchers_Comissao_Deslibera = SUCESSO

    Exit Function

Erro_TRPVouchers_Comissao_Deslibera:

    TRPVouchers_Comissao_Deslibera = gErr

    Select Case gErr

        Case 196386
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 196387, 196388, 196391
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPVOUCHERINFO", gErr)

        Case 196389
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRPVOUCHERINFO", gErr)

        Case 196390
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRPVOUCHERINFO", gErr)

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
    
Dim lErro As Long
Dim alComando(0 To 6) As Long
Dim iIndice As Integer
Dim sSQL As String
Dim sTipVou As String
Dim sSerie As String
Dim lNumVou As Long
Dim lNumFat As Long
Dim dtDataCanc As Date
Dim dtdatafat As Date
Dim objVouExp As ClassTRPVouExp
Dim lClienteVou As Long
Dim iNumParcela As Integer

On Error GoTo Erro_BaixaParcRec_Cancelar_Customizado

'    'Abertura de Comando
'    For iIndice = LBound(alComando) To UBound(alComando)
'        alComando(iIndice) = Comando_Abrir()
'        If alComando(iIndice) = 0 Then gError 200200
'    Next
'
'    'Alocação de espaço no buffer
'    sTipVou = String(STRING_TRP_VOU_INFO_TIPVOU, 0)
'    sSerie = String(STRING_TRP_VOU_INFO_SERIE, 0)
'
'    sSQL = "SELECT V.TipVou, V.Serie, V.NumVou, V.DataCanc, T.NumTitulo, T.DataEmissao, V.ClienteVou, P.NumParcela " & _
'        "FROM ParcelasRecTodas AS P, BaixasParcRec AS B, TRPVouchers As V , TitulosRecTodos AS T " & _
'        "WHERE B.NumIntDoc = ? And B.NumIntParcela = P.NumIntDoc And P.NumintTitulo = T.NumIntDoc And V.NumIntDocDestino = T.NumIntDoc And V.TipoDocDestino = 3"
'
'    'Le os vouchers
'    lErro = Comando_Executar(alComando(0), sSQL, sTipVou, sSerie, lNumVou, dtDataCanc, lNumFat, dtdatafat, lClienteVou, iNumParcela, lNumIntDocBaixaParcRec)
'    If lErro <> AD_SQL_SUCESSO Then gError 200201
'
'    'Busca Primeiro
'    lErro = Comando_BuscarPrimeiro(alComando(0))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200202
'
'    Do While lErro <> AD_SQL_SEM_DADOS
'
'        If iNumParcela = 1 Then
'            Set objVouExp = New ClassTRPVouExp
'            objVouExp.sSerie = sSerie
'            objVouExp.sTipVou = sTipVou
'            objVouExp.lNumVou = lNumVou
'            objVouExp.dtDataCanc = dtDataCanc
'            objVouExp.lNumFat = lNumFat
'            objVouExp.dtdatafat = dtdatafat
'            objVouExp.iTransacao = TRP_EXPORT_VOU_TRANS_CANC_PAGTO
'            objVouExp.lCliente = lClienteVou
'            lErro = TRPVouExp_Insere(alComando(1), alComando(2), alComando(3), alComando(4), alComando(5), alComando(6), objVouExp)
'            If lErro <> SUCESSO Then gError 200204
'        End If
'
'        'Busca Próximo
'        lErro = Comando_BuscarProximo(alComando(0))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200205
'
'    Loop
    
    lErro = TRPVouchers_Comissao_Deslibera(TRP_TIPO_LIBERACAO_COMISSAO_BAIXA, lNumIntDocBaixaParcRec)
    If lErro <> SUCESSO Then gError 200206

    BaixaParcRec_Cancelar_Customizado = SUCESSO

    Exit Function

Erro_BaixaParcRec_Cancelar_Customizado:

    BaixaParcRec_Cancelar_Customizado = gErr

    Select Case gErr

        Case 200200
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 200201, 200202, 200205
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPVOUCHERINFO", gErr)
            
        Case 200204, 200206

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200207)

    End Select

'    'Fecha Comando
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next

    Exit Function

End Function

Function TRPGerComiInt_Exclui(objTRPGerComiInt As ClassTRPGerComiInt) As Long

Dim lErro As Long
Dim alComando(1 To 8) As Long
Dim iIndice As Integer
Dim lTransacao As Long
Dim lNumIntDoc As Long
Dim lNumReg As Long

On Error GoTo Erro_TRPGerComiInt_Exclui

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 197438
    
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 197439
    Next
    
    lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntDoc FROM TRPGerComiInt WHERE DataGeracao = ?", 0, lNumIntDoc, objTRPGerComiInt.dtDataGeracao)
    If lErro <> AD_SQL_SUCESSO Then gError 197440
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197441

    If lErro = AD_SQL_SEM_DADOS Then gError 197442

    lErro = Comando_LockExclusive(alComando(1))
    If lErro <> AD_SQL_SUCESSO Then gError 197443
    
    objTRPGerComiInt.lNumIntDoc = lNumIntDoc

    lErro = Comando_Executar(alComando(2), "SELECT Count(*) FROM Comissoes, TRPGerComiIntDet, TRPGerComiInt WHERE TRPGerComiInt.NumIntDoc = ? AND Comissoes.NumIntCom = TRPGerComiIntDet.NumIntDocComi AND TRPGerComiInt.NumIntDoc = TRPGerComiIntDet.NumIntDocGerComi AND Comissoes.DataBaixa <> ? ", lNumReg, objTRPGerComiInt.lNumIntDoc, DATA_NULA)
    If lErro <> AD_SQL_SUCESSO Then gError 197444
    
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO Then gError 197445

    'se ja existir comissao baixada ==> nao pode exluir
    If lNumReg <> 0 Then gError 197446

    lErro = Comando_Executar(alComando(3), "DELETE FROM Comissoes WHERE NumIntCom IN (SELECT NumIntDocComi FROM TRPGerComiIntDet WHERE NumIntDocGerComi = ?)", objTRPGerComiInt.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 197447

    lErro = Comando_Executar(alComando(4), "DELETE FROM TRPVoucherInfoGer WHERE NumIntComiIntDet IN (SELECT NumIntDoc FROM TRPGerComiIntDet WHERE NumIntDocGerComi = ?)", objTRPGerComiInt.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 197448

    lErro = Comando_Executar(alComando(5), "UPDATE TRPVoucherInfo SET  NumIntDocComiInt = 0 WHERE NumIntDocComiInt = ?", objTRPGerComiInt.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 197449

    lErro = Comando_Executar(alComando(6), "DELETE FROM TRPGerComiIntDet WHERE NumIntDocGerComi = ?", objTRPGerComiInt.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 197450

    lErro = Comando_ExecutarPos(alComando(7), "DELETE FROM TRPGerComiInt ", alComando(1))
    If lErro <> AD_SQL_SUCESSO Then gError 197451

    lErro = Comando_Executar(alComando(8), "DELETE FROM TRPGerComiIntDetVou WHERE NumIntComiIntDet IN (SELECT NumIntDoc FROM TRPGerComiIntDet WHERE NumIntDocGerComi = ?)", objTRPGerComiInt.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 197448

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 197452

    TRPGerComiInt_Exclui = SUCESSO

    Exit Function

Erro_TRPGerComiInt_Exclui:

    TRPGerComiInt_Exclui = gErr

    Select Case gErr

        Case 197438
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 197439
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 197440, 197441
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPGERCOMIINT", gErr)

        Case 197442
            Call Rotina_Erro(vbOKOnly, "ERRO_TRPGERCOMIINT_NAO_CADASTRADO", gErr, objTRPGerComiInt.lCodigo)

        Case 197443
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_TRPGERCOMIINT", gErr, objTRPGerComiInt.lCodigo)

        Case 197444, 197445
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMISSOES3", gErr)

        Case 197446
            Call Rotina_Erro(vbOKOnly, "ERRO_EXC_TRPGERCOMIINT_COMIS_BAIXADA", gErr, objTRPGerComiInt.lCodigo)
        
        Case 197447
            Call Rotina_Erro(vbOKOnly, "ERRO_EXC_COMISSOES_TRPGERCOMIINT", gErr, objTRPGerComiInt.lCodigo)

        Case 197448
            Call Rotina_Erro(vbOKOnly, "ERRO_EXC_TRPVOUCHERINFOGER", gErr, objTRPGerComiInt.lCodigo)

        Case 197449
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_TRPVOUCHERINFO1", gErr, objTRPGerComiInt.lCodigo)

        Case 197450
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRPGERCOMIINTDET", gErr, objTRPGerComiInt.lCodigo)

        Case 197451
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRPGERCOMIINT", gErr, objTRPGerComiInt.lCodigo)

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

Public Function TRPClienteRepProd_Exclui(ByVal objCliente As ClassCliente) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lCliente As Long

On Error GoTo Erro_TRPClienteRepProd_Exclui

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 196931
    Next

    'Le a tabelaTRPClienteRepProd
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Cliente FROM TRPClienteRepProd WHERE Cliente= ? ", _
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
    
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM TRPClienteRepProd", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 196935
        
        'Busca Próximo
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196936
        
    Loop

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TRPClienteRepProd_Exclui = SUCESSO

    Exit Function

Erro_TRPClienteRepProd_Exclui:

    TRPClienteRepProd_Exclui = gErr

    Select Case gErr

        Case 196931
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 196932, 196933, 196936
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPCLIENTEREPPROD", gErr)
        
        Case 196934
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRPCLIENTEREPPROD", gErr)

        Case 196935
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRPCLIENTEREPPROD", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196937)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function TRPExcComissaoCli_Exclui(ByVal objCliente As ClassCliente) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lCliente As Long

On Error GoTo Erro_TRPExcComissaoCli_Exclui

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 196931
    Next

    'Le a tabelaTRPExcComissaoCli
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Cliente FROM TRPExcComissaoCli WHERE Cliente= ? ", _
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
    
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM TRPExcComissaoCli", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 196935
        
        'Busca Próximo
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196936
        
    Loop

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TRPExcComissaoCli_Exclui = SUCESSO

    Exit Function

Erro_TRPExcComissaoCli_Exclui:

    TRPExcComissaoCli_Exclui = gErr

    Select Case gErr

        Case 196931
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 196932, 196933, 196936
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPEXCCOMISSAOCLI", gErr)
        
        Case 196934
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRPEXCCOMISSAOCLI", gErr)

        Case 196935
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRPEXCCOMISSAOCLI", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196937)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function TRPClienteCorProd_Exclui(ByVal objCliente As ClassCliente) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lCliente As Long

On Error GoTo Erro_TRPClienteCorProd_Exclui

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 196938
    Next

    'Le a tabela TRPClienteRepProd
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Cliente FROM TRPClienteCorProd WHERE Cliente= ? ", _
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
    
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM TRPClienteCorProd", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 196942
        
        'Busca Próximo
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196936
        
    Loop

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TRPClienteCorProd_Exclui = SUCESSO

    Exit Function

Erro_TRPClienteCorProd_Exclui:

    TRPClienteCorProd_Exclui = gErr

    Select Case gErr

        Case 196938
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 196939, 196940, 196943
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPCLIENTECORPROD", gErr)
        
        Case 196941
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRPCLIENTECORPROD", gErr)

        Case 196942
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRPCLIENTECORPROD", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196944)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function TRPClienteRepProd_Insere(ByVal lComando As Long, ByVal objTRPClienteRepProd As ClassTRPClienteRepProd) As Long

Dim lErro As Long

On Error GoTo Erro_TRPClienteRepProd_Insere

    lErro = Comando_Executar(lComando, "INSERT INTO TRPClienteRepProd( Cliente, Seq, Representante, Produto, " & _
            "PercComis) VALUES (?,?,?,?,?)", _
            objTRPClienteRepProd.lCliente, objTRPClienteRepProd.iSeq, objTRPClienteRepProd.lRepresentante, objTRPClienteRepProd.sProduto, _
            objTRPClienteRepProd.dPercComis)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196946

    TRPClienteRepProd_Insere = SUCESSO

    Exit Function

Erro_TRPClienteRepProd_Insere:

    TRPClienteRepProd_Insere = gErr

    Select Case gErr
        
        Case 196946
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRPCLIENTEREPPROD", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196947)

    End Select

    Exit Function

End Function

Public Function TRPExcComissaoCli_Insere(ByVal lComando As Long, ByVal objTRPExcComissaoCli As ClassTRPExcComissaoCli) As Long

Dim lErro As Long

On Error GoTo Erro_TRPExcComissaoCli_Insere

    lErro = Comando_Executar(lComando, "INSERT INTO TRPExcComissaoCli( Cliente, Seq, Produto, " & _
            "PercComissao) VALUES (?,?,?,?)", _
            objTRPExcComissaoCli.lCliente, objTRPExcComissaoCli.iSeq, objTRPExcComissaoCli.sProduto, _
            objTRPExcComissaoCli.dPercComissao)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196946

    TRPExcComissaoCli_Insere = SUCESSO

    Exit Function

Erro_TRPExcComissaoCli_Insere:

    TRPExcComissaoCli_Insere = gErr

    Select Case gErr
        
        Case 196946
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRPEXCCOMISSAOCLI", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196947)

    End Select

    Exit Function

End Function

Public Function TRPClienteCorProd_Insere(ByVal lComando As Long, ByVal objTRPClienteCorProd As ClassTRPClienteCorProd) As Long

Dim lErro As Long

On Error GoTo Erro_TRPClienteCorProd_Insere

    lErro = Comando_Executar(lComando, "INSERT INTO TRPClienteCorProd( Cliente, Seq, Correntista, Produto, " & _
            "PercComis) VALUES (?,?,?,?,?)", _
            objTRPClienteCorProd.lCliente, objTRPClienteCorProd.iSeq, objTRPClienteCorProd.lCorrentista, objTRPClienteCorProd.sProduto, _
            objTRPClienteCorProd.dPercComis)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196949

    TRPClienteCorProd_Insere = SUCESSO

    Exit Function

Erro_TRPClienteCorProd_Insere:

    TRPClienteCorProd_Insere = gErr

    Select Case gErr
    
        Case 196949
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRPCLIENTECORPROD", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196950)

    End Select

    Exit Function

End Function


'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'GRAVAÇÃO
Public Function TRPTiposOcorrencia_Grava(ByVal objTRPTiposOcorrencia As ClassTRPTiposOcorrencia) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_TRPTiposOcorrencia_Grava

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 198017

    lErro = CF("TRPTiposOcorrencia_Grava_EmTrans", objTRPTiposOcorrencia)
    If lErro <> SUCESSO Then gError 198018

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 198019

    TRPTiposOcorrencia_Grava = SUCESSO

    Exit Function

Erro_TRPTiposOcorrencia_Grava:

    TRPTiposOcorrencia_Grava = gErr

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

Public Function TRPTiposOcorrencia_Grava_EmTrans(ByVal objTRPTiposOcorrencia As ClassTRPTiposOcorrencia) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim iCodigo As Integer

On Error GoTo Erro_TRPTiposOcorrencia_Grava_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 198021
    Next

    'Le a tabelaTRPTiposOcorrencia
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Codigo FROM TRPTiposOcorrencia WHERE Codigo= ? ", _
                0, iCodigo, objTRPTiposOcorrencia.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 198022

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 198023

    'Se existir => UPDATE, senão => INSERT
    If lErro = AD_SQL_SUCESSO Then

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 198024

        lErro = Comando_ExecutarPos(alComando(1), "UPDATE TRPTiposOcorrencia SET Descricao= ? , EstornaAporteVou= ? , ConsideraComisInt= ? , AlteraComiVou= ? , " & _
                "AlteraCMCC= ? , AlteraCMC= ? , AlteraCMR= ? , AlteraOVER= ? , AlteraCMA= ? , " & _
                "AceitaVlrPositivo= ? , AceitaVlrNegativo= ?, IncideSobre = ? ", alComando(0), _
                objTRPTiposOcorrencia.sDescricao, objTRPTiposOcorrencia.iEstornaAporteVou, objTRPTiposOcorrencia.iConsideraComisInt, objTRPTiposOcorrencia.iAlteraComiVou, _
                objTRPTiposOcorrencia.iAlteraCMCC, objTRPTiposOcorrencia.iAlteraCMC, objTRPTiposOcorrencia.iAlteraCMR, objTRPTiposOcorrencia.iAlteraOVER, objTRPTiposOcorrencia.iAlteraCMA, _
                objTRPTiposOcorrencia.iAceitaVlrPositivo, objTRPTiposOcorrencia.iAceitaVlrNegativo, objTRPTiposOcorrencia.iIncideSobre)
        If lErro <> AD_SQL_SUCESSO Then gError 198025

    Else


        lErro = Comando_Executar(alComando(1), "INSERT INTO TRPTiposOcorrencia( Codigo, Descricao, EstornaAporteVou, ConsideraComisInt, " & _
                "AlteraComiVou, AlteraCMCC, AlteraCMC, AlteraCMR, AlteraOVER, " & _
                "AlteraCMA, AceitaVlrPositivo, AceitaVlrNegativo, IncideSobre) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                objTRPTiposOcorrencia.iCodigo, objTRPTiposOcorrencia.sDescricao, objTRPTiposOcorrencia.iEstornaAporteVou, objTRPTiposOcorrencia.iConsideraComisInt, _
                objTRPTiposOcorrencia.iAlteraComiVou, objTRPTiposOcorrencia.iAlteraCMCC, objTRPTiposOcorrencia.iAlteraCMC, objTRPTiposOcorrencia.iAlteraCMR, objTRPTiposOcorrencia.iAlteraOVER, _
                objTRPTiposOcorrencia.iAlteraCMA, objTRPTiposOcorrencia.iAceitaVlrPositivo, objTRPTiposOcorrencia.iAceitaVlrNegativo, objTRPTiposOcorrencia.iIncideSobre)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 198026

    End If

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TRPTiposOcorrencia_Grava_EmTrans = SUCESSO

    Exit Function

Erro_TRPTiposOcorrencia_Grava_EmTrans:

    TRPTiposOcorrencia_Grava_EmTrans = gErr

    Select Case gErr

        Case 198021
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 198022, 198023
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPTIPOSOCORRENCIA", gErr)

        Case 198024
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRPTIPOSOCORRENCIA", gErr)

        Case 198025
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRPTIPOSOCORRENCIA", gErr)

        Case 198026
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRPTIPOSOCORRENCIA", gErr)

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
Public Function TRPTiposOcorrencia_Exclui(ByVal objTRPTiposOcorrencia As ClassTRPTiposOcorrencia) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_TRPTiposOcorrencia_Exclui

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 198006

    lErro = CF("TRPTiposOcorrencia_Exclui_EmTrans", objTRPTiposOcorrencia)
    If lErro <> SUCESSO Then gError 198007

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 198008

    TRPTiposOcorrencia_Exclui = SUCESSO

    Exit Function

Erro_TRPTiposOcorrencia_Exclui:

    TRPTiposOcorrencia_Exclui = gErr

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

Public Function TRPTiposOcorrencia_Exclui_EmTrans(ByVal objTRPTiposOcorrencia As ClassTRPTiposOcorrencia) As Long

Dim lErro As Long
Dim alComando(0 To 2) As Long
Dim iIndice As Integer
Dim iCodigo As Integer
Dim lCodigoOcr As Long

On Error GoTo Erro_TRPTiposOcorrencia_Exclui_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 198010
    Next

    'Le a tabelaTRPTiposOcorrencia
    lErro = Comando_ExecutarPos(alComando(2), "SELECT Codigo FROM TRPOcorrencias WHERE Origem= ? ", _
                0, lCodigoOcr, objTRPTiposOcorrencia.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 198089

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 198090

    'Se existir => ERRO
    If lErro <> AD_SQL_SEM_DADOS Then gError 198091

    'Le a tabelaTRPTiposOcorrencia
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Codigo FROM TRPTiposOcorrencia WHERE Codigo= ? ", _
                0, iCodigo, objTRPTiposOcorrencia.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 198011

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 198012

    'Se não existir => ERRO
    If lErro = AD_SQL_SEM_DADOS Then gError 198013

    'Trava registro contra alterações/Leituras
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> SUCESSO Then gError 198014

    lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM TRPTiposOcorrencia", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 198015

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TRPTiposOcorrencia_Exclui_EmTrans = SUCESSO

    Exit Function

Erro_TRPTiposOcorrencia_Exclui_EmTrans:

    TRPTiposOcorrencia_Exclui_EmTrans = gErr

    Select Case gErr

        Case 198010
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 198011, 198012
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPTIPOSOCORRENCIA", gErr)

        Case 198013
            Call Rotina_Erro(vbOKOnly, "ERRO_TRPTIPOSOCORRENCIA_NAO_CADASTRADO", gErr, objTRPTiposOcorrencia.iCodigo)

        Case 198014
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRPTIPOSOCORRENCIA", gErr)

        Case 198015
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRPTIPOSOCORRENCIA", gErr)
            
        Case 198089, 198090
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPOCORRENCIAS", gErr)

        Case 198091
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_ASSOC_OCR", gErr, objTRPTiposOcorrencia.iCodigo, lCodigoOcr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198016)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function TRPTitulos_Exclui(ByVal objTRPTitulos As ClassTRPTitulos, ByVal lComando1 As Long, ByVal lComando2 As Long) As Long

Dim lErro As Long
Dim lNumIntDoc As Long

On Error GoTo Erro_TRPTitulos_Exclui

    'Le a tabelaTRPTiposOcorrencia
    lErro = Comando_ExecutarPos(lComando1, "SELECT NumIntDoc FROM TRPTitulos WHERE TipoDocOrigem = ? AND NumIntDocOrigem = ? ", _
                0, lNumIntDoc, objTRPTitulos.iTipoDocOrigem, objTRPTitulos.lNumIntDocOrigem)
    If lErro <> AD_SQL_SUCESSO Then gError 200232

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200233

    'Se não existir => ERRO
    If lErro <> AD_SQL_SEM_DADOS Then

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(lComando1)
        If lErro <> SUCESSO Then gError 200234
    
        lErro = Comando_ExecutarPos(lComando2, "DELETE FROM TRPTitulos", lComando1)
        If lErro <> AD_SQL_SUCESSO Then gError 200235
        
    End If

    TRPTitulos_Exclui = SUCESSO

    Exit Function

Erro_TRPTitulos_Exclui:

    TRPTitulos_Exclui = gErr

    Select Case gErr

        Case 200232, 200233
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPTITULOS", gErr)

        Case 200234
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRPTITULOS", gErr)

        Case 200235
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRPTITULOS", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200236)

    End Select

    Exit Function

End Function

Function TRPEmissor_Automatico(lCodigo As Long) As Long
'retorna o número do proximo Fornecedor disponivel

    TRPEmissor_Automatico = CF("Config_ObterAutomatico", "TRPConfig", "NUM_PROX_EMISSOR", "TRPCliEmissores", "Codigo", lCodigo)

End Function

Function TRPEmissores_Grava(ByVal objForn As ClassFornecedor, ByVal objEnd As ClassEndereco, ByVal objEmi As ClassCliEmissoresTRP) As Long
'Grava um novo Fornecedor ou atualiza um já existente

Dim lErro As Long
Dim lTransacao As Long
Dim alComando(0 To 9) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim lFornEmissor As Long
Dim iMaxSeq As Integer
Dim lCliente As Long
Dim objEmiExc As ClassCliEmissoresExcTRP
Dim lNumIntDoc1 As Long
Dim iIndice1 As Integer

On Error GoTo Erro_TRPEmissores_Grava

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 200351
    
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 200352
    Next

    'Le a tabelaTRPVouchers
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Seq FROM TRPCliEmissores WHERE Cliente = ? ORDER BY Seq DESC ", 0, iMaxSeq, objEmi.lCliente)
    If lErro <> AD_SQL_SUCESSO Then gError 200353

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200354
       
    'Le a tabelaTRPVouchers
    lErro = Comando_ExecutarPos(alComando(9), "SELECT NumIntDoc, FornEmissor, Cliente FROM TRPCliEmissores WHERE Codigo = ? ", 0, lNumIntDoc, lFornEmissor, lCliente, objEmi.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 200355

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(9))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200356
    
    'Se existir => UPDATE, senão => INSERT
    If lErro = AD_SQL_SUCESSO Then

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(9))
        If lErro <> SUCESSO Then gError 200357
        
        objEmi.lNumIntDoc = lNumIntDoc
        objEmi.lFornEmissor = lFornEmissor

        'Se trocou de cliente acerta o Seq também
        If objEmi.lCliente <> lCliente Then gError 200715
            
'            objEmi.iSeq = iMaxSeq + 1
'
'            lErro = Comando_ExecutarPos(alComando(1), "UPDATE TRPCliEmissores SET Cliente = ?, Seq = ?, PercComissao = ?, EmissorSuperior = ?, Hierarquia = ? ", alComando(0), _
'                    objEmi.lCliente, objEmi.iSeq, objEmi.dPercComissao, objEmi.lSuperior, objEmi.sHierarquia)
'        Else
            lErro = Comando_ExecutarPos(alComando(1), "UPDATE TRPCliEmissores SET PercComissao = ?, EmissorSuperior = ? ", alComando(9), _
                    objEmi.dPercComissao, objEmi.lSuperior)
'        End If
        If lErro <> AD_SQL_SUCESSO Then gError 200358

    Else
    
        lErro = CF("Config_ObterAutomatico_EmTrans", "TRPConfig", "NUM_PROX_EMISSORFORN", "Fornecedores", "Codigo", lFornEmissor)
        If lErro <> SUCESSO Then gError 200359
        
        If lFornEmissor < FAIXA_TRP_TIPO_CLIENTE_EMISSOR_DE Or lFornEmissor > FAIXA_TRP_TIPO_CLIENTE_EMISSOR_ATE Then gError 200360

        lErro = CF("Config_ObterNumInt", "TRPConfig", TRPCONFIG_NUM_INT_PROX_TRPCLIEMISSORES, lNumIntDoc)
        If lErro <> SUCESSO Then gError 200361

        objEmi.lFornEmissor = lFornEmissor
        objEmi.lNumIntDoc = lNumIntDoc
        objEmi.iSeq = iMaxSeq + 1

        lErro = Comando_Executar(alComando(1), "INSERT INTO TRPCliEmissores(NumIntDoc, Cliente, Seq, FornEmissor, PercComissao, EmissorSuperior, Codigo) VALUES (?,?,?,?,?,?,?) ", _
        objEmi.lNumIntDoc, objEmi.lCliente, objEmi.iSeq, objEmi.lFornEmissor, objEmi.dPercComissao, objEmi.lSuperior, objEmi.lCodigo)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200362

    End If

    lErro = Comando_Executar(alComando(2), "DELETE FROM TRPCliEmissoresExc WHERE NumIntDocCliEmi  = ?", objEmi.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 200363
    
    If objEmi.colCliEmissoresExcTRP.Count > 0 Then

        lErro = CF("Config_ObterNumInt", "TRPConfig", TRPCONFIG_NUM_INT_PROX_TRPCLIEMISSORESEXC, lNumIntDoc1, objEmi.colCliEmissoresExcTRP.Count)
        If lErro <> SUCESSO Then gError 200364

        iIndice1 = 0
        
        For Each objEmiExc In objEmi.colCliEmissoresExcTRP

            iIndice1 = iIndice1 + 1

            objEmiExc.lNumIntDoc = lNumIntDoc1
            objEmiExc.lNumIntDocCliEmi = objEmi.lNumIntDoc
            objEmiExc.iSeq = iIndice1

            lNumIntDoc1 = lNumIntDoc1 + 1
            
            lErro = Comando_Executar(alComando(3), "INSERT INTO TRPCliEmissoresExc (NumIntDoc, NumIntDocCliEmi, Seq, Produto, PercComissao) VALUES (?,?,?,?,?)", _
                     objEmiExc.lNumIntDoc, objEmiExc.lNumIntDocCliEmi, objEmiExc.iSeq, objEmiExc.sProduto, objEmiExc.dPercComissao)
            If lErro <> AD_SQL_SUCESSO Then gError 200365
            
        Next

    End If
    
    lErro = TRPEmissores_Monta_Hierarquia(alComando(5), alComando(6), alComando(7), objEmi)
    If lErro <> SUCESSO Then gError 200359

    objForn.lCodigo = objEmi.lFornEmissor
    objForn.iTipo = TRP_TIPO_CLIENTE_EMISSOR

    lErro = CF("Fornecedor_Grava_EmTrans", objForn, objEnd)
    If lErro <> SUCESSO Then gError 200366
    
    lErro = CF("TRPUsuarios_Grava", objEmi.objUsuarioWeb)
    If lErro <> SUCESSO Then gError 200376

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 200367

    TRPEmissores_Grava = SUCESSO

    Exit Function

Erro_TRPEmissores_Grava:

    TRPEmissores_Grava = gErr

    Select Case gErr

        Case 200351
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 200352
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 200353 To 200356
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPCLIEMISSORES", gErr)
            
        Case 200357 'LOCK
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUIVE_TRPCLIEMISSORES", gErr)
        
        Case 200358 'UPDATE
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRPCLIEMISSORES", gErr)
        
        Case 200359, 200361, 200364, 200366, 200376
        
        Case 200715
            Call Rotina_Erro(vbOKOnly, "ERRO_TROCA_CLIENTE_EMISSOR", gErr)
        
        Case 200360
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_FORA_FAIXA_EMI", gErr, FAIXA_TRP_TIPO_CLIENTE_EMISSOR_DE, FAIXA_TRP_TIPO_CLIENTE_EMISSOR_ATE)

        Case 200362
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERT_TRPCLIEMISSORES", gErr)
            
        Case 200363
            Call Rotina_Erro(vbOKOnly, "ERRO_DELETE_TRPCLIEMISSORESEXC", gErr)
        
        Case 200365
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERT_TRPCLIEMISSORESEXC", gErr)

        Case 200367
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT_TRANSACAO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200368)

    End Select

    Call Transacao_Rollback
    
    Exit Function

End Function

Public Function TRPUsuarios_Grava(ByVal objUsuWeb As ClassTRPUsuarios) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lCodigo As Integer

On Error GoTo Erro_TRPUsuarios_Grava

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 200369
    Next

    'Le a tabelaTRPUsuarios
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Codigo FROM TRPUsuarios WHERE TipoUsuario= ?  AND Codigo= ? ", _
                0, lCodigo, objUsuWeb.iTipoUsuario, objUsuWeb.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 200370

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200371

    'Se existir => UPDATE, senão => INSERT
    If lErro = AD_SQL_SUCESSO Then

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 200372

        lErro = Comando_ExecutarPos(alComando(1), "UPDATE TRPUsuarios SET Login= ? , Senha= ? , AlteraSenhaProxLog= ? , GrupoAcesso= ? ", alComando(0), _
                objUsuWeb.sLogin, objUsuWeb.sSenha, objUsuWeb.iAlteraSenhaProxLog, objUsuWeb.sGrupoAcesso)
        If lErro <> AD_SQL_SUCESSO Then gError 200373

    Else


        lErro = Comando_Executar(alComando(1), "INSERT INTO TRPUsuarios( TipoUsuario, Codigo, Login, Senha, " & _
                "AlteraSenhaProxLog, GrupoAcesso) VALUES (?,?,?,?,?,?)", _
                objUsuWeb.iTipoUsuario, objUsuWeb.lCodigo, objUsuWeb.sLogin, objUsuWeb.sSenha, _
                objUsuWeb.iAlteraSenhaProxLog, objUsuWeb.sGrupoAcesso)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200374

    End If

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TRPUsuarios_Grava = SUCESSO

    Exit Function

Erro_TRPUsuarios_Grava:

    TRPUsuarios_Grava = gErr

    Select Case gErr

        Case 200369
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 200370, 200371
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPUSUARIOS", gErr)

        Case 200372
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRPUSUARIOS", gErr)

        Case 200373
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRPUSUARIOS", gErr)

        Case 200374
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRPUSUARIOS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200375)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function TRPVouchers_GerComi(ByVal objTela As Object, ByVal colVouchers As Collection) As Long

Dim lErro As Long
Dim alComando(1 To 11) As Long
Dim iIndice As Integer
Dim objVou As ClassTRPVouchers
Dim lTransacao As Long
Dim lClienteVou As Long
Dim lClienteFat As Long

On Error GoTo Erro_TRPVouchers_GerComi
   
    For Each objVou In colVouchers

        'Inicia a Transacao
        lTransacao = Transacao_Abrir()
        If lTransacao = 0 Then gError 200406

        'Abertura de Comando
        For iIndice = LBound(alComando) To UBound(alComando)
            alComando(iIndice) = Comando_Abrir()
            If alComando(iIndice) = 0 Then gError 200407
        Next

        lErro = CF("TRPVouchers_Le", objVou, False, True, alComando(1))
        If lErro <> SUCESSO Then gError 200408
        
        If objVou.iGeraComissao = MARCADO Then
        
            If objVou.iCartao = DESMARCADO Then
                objVou.lClienteFat = objVou.lClienteVou
            Else
                        
                'Le o Cliente
                lErro = Comando_Executar(alComando(8), "SELECT R.Cliente FROM AdmMeioPagto AS P, Redes AS R WHERE P.FilialEmpresa = 1 AND  P.Nome = ? AND P.Rede = R.Codigo", _
                            lClienteFat, objVou.sCiaCart)
                If lErro <> AD_SQL_SUCESSO Then gError 200408
            
                'Busca Primeiro
                lErro = Comando_BuscarPrimeiro(alComando(8))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200408
            
                objVou.lClienteFat = lClienteFat
            
            End If
            
            objVou.lClienteComissao = objVou.lClienteVou
        
            lErro = Comando_ExecutarPos(alComando(2), "UPDATE TRPVouchers SET GeraComissao= ?, ClienteFat = ?, ClienteComissao = ? ", alComando(1), _
                    DESMARCADO, objVou.lClienteFat, objVou.lClienteComissao)
            If lErro <> AD_SQL_SUCESSO Then gError 200410
        
            'Gera a comissão e contabiliza
            lErro = TRPVou_Gera_Comis(objVou, objVou.dValorBruto, TRP_TIPO_VALOR_BASE_BRU, objVou.dtData, "Voucher emitido", TRP_TIPO_TRATAMENTO_COMI_VOU_NOVO)
            If lErro <> SUCESSO Then gError 200409
            
            'Guarda a hierarquia de vendedores
            lErro = TRPVouVendedores_Grava(objVou, alComando(3), alComando(4), alComando(5), alComando(6), alComando(7), alComando(9), alComando(10), alComando(11))
            If lErro <> SUCESSO Then gError 200409
        
            lErro = Comando_ExecutarPos(alComando(2), "UPDATE TRPVouchers SET Promotor = ?, ComissaoProm = ? ", alComando(1), _
                    objVou.iPromotor, objVou.dComissaoProm)
            If lErro <> AD_SQL_SUCESSO Then gError 200410
        
        End If
            
        'Fecha Comando
        For iIndice = LBound(alComando) To UBound(alComando)
            Call Comando_Fechar(alComando(iIndice))
        Next
        
        lErro = TRPVouHistorico_Grava(objVou, objVou)
        If lErro <> SUCESSO Then gError 200409
    
        'Confirma a transação
        lErro = Transacao_Commit()
        If lErro <> AD_SQL_SUCESSO Then gError 200411
        
        lErro = objTela.Processa_Registro
        If lErro <> SUCESSO Then gError 200408
                        
    Next
    
    TRPVouchers_GerComi = SUCESSO

    Exit Function

Erro_TRPVouchers_GerComi:

    TRPVouchers_GerComi = gErr

    Select Case gErr

        Case 200406
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 200407
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 200408, 200409

        Case 200410
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRPVOUCHERS", gErr)

        Case 200411
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200412)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Call Transacao_Rollback
    
    Exit Function

End Function

Public Function TRPVouVendedores_Remonta(ByVal dtDataDe As Date, ByVal dtDataAte As Date, ByVal lCliente As Long) As Long

Dim lErro As Long
Dim alComando(1 To 15) As Long
Dim iIndice As Integer
Dim objVou As ClassTRPVouchers
Dim lTransacao As Long
Dim lNumVou As Long
Dim sSerie As String
Dim sTipVou As String
Dim objFrmAguarde As New ClassFrmAguarde
Dim objFrmAguardeTela As New FrmAguarde
Dim lCount As Long

On Error GoTo Erro_TRPVouVendedores_Remonta
   
    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 200794

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 200795
    Next

    'Le a tabelaTRPVouchers
    If lCliente <> 0 Then
        lErro = Comando_Executar(alComando(12), "SELECT Count(*) FROM TRPVouchers WHERE Data BETWEEN  ? AND  ? AND GeraComissao = 0 AND ClienteVou = ?", _
                     lCount, dtDataDe, dtDataAte, lCliente)
    Else
        lErro = Comando_Executar(alComando(12), "SELECT Count(*) FROM TRPVouchers WHERE Data BETWEEN  ? AND  ? AND GeraComissao = 0", _
                     lCount, dtDataDe, dtDataAte)
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 200796

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(12))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200797
    
    objFrmAguarde.iTotalItens = lCount
    Call objFrmAguardeTela.Inicializa_Progressao(objFrmAguarde)

    sSerie = String(STRING_TRP_OCR_SERIE, 0)
    sTipVou = String(STRING_TRP_OCR_TIPOVOU, 0)

    'Le a tabelaTRPVouchers
    If lCliente <> 0 Then
        lErro = Comando_Executar(alComando(2), "SELECT NumVou, Serie, TipVou FROM TRPVouchers WHERE Data BETWEEN  ? AND  ? AND GeraComissao = 0 AND ClienteVou = ?", _
                     lNumVou, sSerie, sTipVou, dtDataDe, dtDataAte, lCliente)
    Else
        lErro = Comando_Executar(alComando(2), "SELECT NumVou, Serie, TipVou FROM TRPVouchers WHERE Data BETWEEN  ? AND  ? AND GeraComissao = 0", _
                     lNumVou, sSerie, sTipVou, dtDataDe, dtDataAte)
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 200798

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200799
    
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objVou = New ClassTRPVouchers
        
        objVou.sSerie = sSerie
        objVou.sTipVou = sTipVou
        objVou.lNumVou = lNumVou
            
        lErro = CF("TRPVouchers_Le", objVou, False, True, alComando(1))
        If lErro <> SUCESSO Then gError 200800
        
        lErro = TRPVouVendedores_Exclui(alComando(13), alComando(14), objVou)
        If lErro <> SUCESSO Then gError 200801
        
        'Guarda a hierarquia de vendedores
        lErro = TRPVouVendedores_Grava(objVou, alComando(3), alComando(4), alComando(5), alComando(6), alComando(7), alComando(9), alComando(10), alComando(11))
        If lErro <> SUCESSO Then gError 200802
    
        lErro = Comando_ExecutarPos(alComando(15), "UPDATE TRPVouchers SET Promotor = ?, ComissaoProm = ? ", alComando(1), _
                objVou.iPromotor, objVou.dComissaoProm)
        If lErro <> AD_SQL_SUCESSO Then gError 200803
        
        'Busca Próximo
        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200804
        
        Call objFrmAguardeTela.ProcessouItem
        
        If objFrmAguarde.iCancelar = MARCADO Then Exit Do

    Loop
    
    Set objFrmAguardeTela = Nothing
    If objFrmAguarde.iCancelar = MARCADO Then gError 200805
            
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 200806
    
    TRPVouVendedores_Remonta = SUCESSO

    Exit Function

Erro_TRPVouVendedores_Remonta:

    TRPVouVendedores_Remonta = gErr

    Select Case gErr

        Case 200794
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 200795
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 200796 To 200799, 200804
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPVOUCHERS", gErr)
        
        Case 200800 To 200802, 200805

        Case 200803
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRPVOUCHERS", gErr)

        Case 200806
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200807)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Call Transacao_Rollback

    If Not (objFrmAguardeTela Is Nothing) Then
        Call objFrmAguardeTela.Trata_Erro
    End If
    
    Exit Function

End Function

Private Function TRPVouVendedores_Grava(ByVal objVou As ClassTRPVouchers, ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lComando3 As Long, ByVal lComando4 As Long, ByVal lComando5 As Long, ByVal lComando6 As Long, ByVal lComando7 As Long, ByVal lComando8 As Long, Optional ByVal bSoGrava As Boolean = False) As Long

Dim lErro As Long
Dim lNumIntDoc As Long
Dim iPromotor As Integer
Dim iRegiao As Integer
Dim colVend As New Collection
Dim objVend As ClassVendedor
Dim objVendPro As ClassVendedor
Dim dPercComissao As Double
Dim dValorAporteDe As Double
Dim iMoedaAporteDe As Integer
Dim dPercComissaoAporteDe As Double
Dim dValor As Double
Dim dValorAporteRS As Double
Dim dValorAporteUSS As Double
Dim lEmpresaPai As Long
Dim iMoeda As Integer

On Error GoTo Erro_TRPVouVendedores_Grava

    If Not bSoGrava Then
       
        'Le a Filiais Clientes
        lErro = Comando_Executar(lComando2, "SELECT F1.Vendedor, F1.Regiao, F2.EmpresaPai FROM FiliaisClientes AS F1, FiliaisClientesTRP AS F2 WHERE F1.CodCliente = F2.CodCliente AND F1.CodFilial = F2.CodFilial AND F1.CodCliente = ? AND F1.CodFilial = ?", _
        iPromotor, iRegiao, lEmpresaPai, objVou.lClienteVou, FILIAL_MATRIZ)
        If lErro <> AD_SQL_SUCESSO Then gError 200393
        
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(lComando2)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200394
        
        'Le os investimentos com desconto em fatura
        lErro = Comando_Executar(lComando6, "SELECT Moeda, Valor FROM TRPAportes, TRPAportePagtoFat WHERE Cliente = ? AND TRPAportes.NumIntDoc = NumIntDocAporte AND ValidadeDe <= ? AND ValidadeAte >= ? ", iMoeda, dValor, lEmpresaPai, objVou.dtData, objVou.dtData)
        If lErro <> AD_SQL_SUCESSO Then gError 200767
      
        lErro = Comando_BuscarPrimeiro(lComando6)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200768
    
        Do While lErro = AD_SQL_SUCESSO
        
            If iMoeda = MOEDA_DOLAR Then
                dValorAporteRS = dValorAporteRS + dValor * objVou.dCambio
                dValorAporteUSS = dValorAporteUSS + dValor
            Else
                dValorAporteRS = dValorAporteRS + dValor
                dValorAporteUSS = dValorAporteUSS + dValor / objVou.dCambio
            End If
                
            lErro = Comando_BuscarProximo(lComando6)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200769
        
        Loop
    
        lErro = Comando_Executar(lComando7, "SELECT Moeda, Valor FROM TRPAportes, TRPAportePagtoDireto WHERE Cliente = ? AND TRPAportes.NumIntDoc = NumIntDocAporte AND DataEmissao <= ? AND DataEmissao+365 >= ? ", iMoeda, dValor, lEmpresaPai, objVou.dtData, objVou.dtData)
        If lErro <> AD_SQL_SUCESSO Then gError 200770
    
        lErro = Comando_BuscarPrimeiro(lComando7)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200771
    
        Do While lErro = AD_SQL_SUCESSO
    
            If iMoeda = MOEDA_DOLAR Then
                dValorAporteRS = dValorAporteRS + dValor * objVou.dCambio
                dValorAporteUSS = dValorAporteUSS + dValor
            Else
                dValorAporteRS = dValorAporteRS + dValor
                dValorAporteUSS = dValorAporteUSS + dValor / objVou.dCambio
            End If
    
            lErro = Comando_BuscarProximo(lComando7)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200772
    
        Loop
    
        objVou.iPromotor = iPromotor
        
        Set objVendPro = New ClassVendedor
        objVendPro.iCodigo = iPromotor
        colVend.Add objVendPro
       
        lErro = TRPVendedor_Le_Recursivo2(lComando3, lComando8, colVend, iRegiao, 1)
        If lErro <> SUCESSO Then gError 200395
    
        For Each objVend In colVend
        
            'Le o Vendedor
            lErro = Comando_Executar(lComando1, "SELECT PerCcomissao FROM Vendedores WHERE Codigo = ? ", _
            dPercComissao, objVend.iCodigo)
            If lErro <> AD_SQL_SUCESSO Then gError 200391
            
            'Busca Primeiro
            lErro = Comando_BuscarPrimeiro(lComando1)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200392
    
            'Le a redução por aporte
            lErro = Comando_Executar(lComando5, "SELECT ValorDe, Moeda, PercComissaoMax FROM VendedorTRPComiRed WHERE Vendedor = ? ", _
            dValorAporteDe, iMoedaAporteDe, dPercComissaoAporteDe, iPromotor)
            If lErro <> AD_SQL_SUCESSO Then gError 200391
            
            'Busca Primeiro
            lErro = Comando_BuscarPrimeiro(lComando5)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200392
            
            'Verifica se o investimento está dentro da redução
            If iMoedaAporteDe = MOEDA_DOLAR Then
                If (dValorAporteDe - DELTA_VALORMONETARIO <= dValorAporteUSS) And dValorAporteDe > 0 Then
                    dPercComissao = dPercComissaoAporteDe
                End If
            Else
                If (dValorAporteDe - DELTA_VALORMONETARIO <= dValorAporteRS) And dValorAporteDe > 0 Then
                    dPercComissao = dPercComissaoAporteDe
                End If
            End If
            
            objVend.dPercComissao = dPercComissao
    
        Next
        
        objVou.dComissaoProm = colVend.Item(1).dPercComissao
        colVend.Remove (1)
        
    Else
        For Each objVend In objVou.colVendedores
            colVend.Add objVend
        Next
    End If
    
    For Each objVend In colVend
    
        lErro = Comando_Executar(lComando4, "INSERT INTO TRPVouVendedores( TipVou, Serie, NumVou, Vendedor, PercComissao) VALUES (?,?,?,?,?)", objVou.sTipVou, objVou.sSerie, objVou.lNumVou, objVend.iCodigo, objVend.dPercComissao)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200396

    Next

    TRPVouVendedores_Grava = SUCESSO

    Exit Function

Erro_TRPVouVendedores_Grava:

    TRPVouVendedores_Grava = gErr

    Select Case gErr

        Case 200391, 200392
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS", gErr)

        Case 200393, 200394
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISCLIENTES", gErr)
            
        Case 200395

        Case 200396
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERT_TRPVOUVENDEDORES", gErr)
        
        Case 200400, 200401
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIALCLIENTECATEGORIAS", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200397)

    End Select

    Exit Function

End Function

Public Function TRPVendedor_Le_Recursivo(ByVal lComando As Long, ByVal colVendedores As Collection, ByVal iVendedorFilho As Integer, ByVal iNivel As Integer) As Long

Dim lErro As Long
Dim iVendedorPai  As Integer
Dim objVendedor As ClassVendedor

On Error GoTo Erro_TRPVendedor_Le_Recursivo

    If iVendedorFilho <> 0 Then

        If iNivel > 20 Then gError 200385
       
        'Le a tabelaTRPProdTarifa
        lErro = Comando_Executar(lComando, "SELECT Superior FROM VendedorTRP WHERE Vendedor = ? ", _
                    iVendedorPai, iVendedorFilho)
        If lErro <> AD_SQL_SUCESSO Then gError 200386
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200387
    
        'Sem Dados
        If lErro <> AD_SQL_SUCESSO Then gError 200388
        
        If iVendedorPai <> 0 Then
        
            Set objVendedor = New ClassVendedor
            
            objVendedor.iCodigo = iVendedorPai
            
            colVendedores.Add objVendedor
        
            lErro = TRPVendedor_Le_Recursivo(lComando, colVendedores, iVendedorPai, iNivel + 1)
            If lErro <> SUCESSO Then gError 200389
        
        End If
    
    End If
    
    TRPVendedor_Le_Recursivo = SUCESSO

    Exit Function

Erro_TRPVendedor_Le_Recursivo:

    TRPVendedor_Le_Recursivo = gErr

    Select Case gErr
    
        Case 200385 'ERRO_LIMITE_RECURSIVIDADE
            Call Rotina_Erro(vbOKOnly, "ERRO_LIMITE_RECURSIVIDADE", gErr)
    
        Case 200386, 200387
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_VENDEDORES", gErr)

        Case 200388
            Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO", gErr, iVendedorFilho)
            
        Case 200389
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200390)

    End Select

    Exit Function

End Function

Public Function TRPVendedor_Le_Recursivo2(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal colVendedores As Collection, ByVal iRegiaoFilha As Integer, ByVal iNivel As Integer) As Long

Dim lErro As Long
Dim iRegiaoPai  As Integer
Dim objVendedor As ClassVendedor
Dim iVendedor As Integer
Dim bAchou As Boolean

On Error GoTo Erro_TRPVendedor_Le_Recursivo2

    If iRegiaoFilha <> 0 Then

        If iNivel > 20 Then gError 200785
       
        'Le a região de Venda
        lErro = Comando_Executar(lComando1, "SELECT CodRegiaoPai FROM TRPRegiaoVenda WHERE Codigo = ? ", _
                    iRegiaoPai, iRegiaoFilha)
        If lErro <> AD_SQL_SUCESSO Then gError 200786
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200787
    
        'Sem Dados
        If lErro <> AD_SQL_SUCESSO Then gError 200788
        
        'Le os vendedores que recebem pela região de venda
        lErro = Comando_Executar(lComando2, "SELECT V1.Codigo FROM Vendedores AS V1, VendedorTRP AS V2, VendedorTRPRegVenda AS V3 WHERE V1.Codigo = V2.Vendedor AND V1.Codigo = V3.Vendedor AND V3.RegiaoVenda = ? ORDER BY V2.Cargo ", _
                    iVendedor, iRegiaoFilha)
        If lErro <> AD_SQL_SUCESSO Then gError 200789
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(lComando2)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200790
        
        Do While lErro <> AD_SQL_SEM_DADOS
        
            bAchou = False
            For Each objVendedor In colVendedores
                If objVendedor.iCodigo = iVendedor Then
                    bAchou = True
                    Exit For
                End If
            Next
            If Not bAchou Then
            
                Set objVendedor = New ClassVendedor
                
                objVendedor.iCodigo = iVendedor
                
                colVendedores.Add objVendedor
            
            End If
        
            'Busca Próximo
            lErro = Comando_BuscarProximo(lComando2)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200791
        
        Loop
        
        If iRegiaoPai <> 0 Then
        
            lErro = TRPVendedor_Le_Recursivo2(lComando1, lComando2, colVendedores, iRegiaoPai, iNivel + 1)
            If lErro <> SUCESSO Then gError 200792
        
        End If
    
    End If
    
    TRPVendedor_Le_Recursivo2 = SUCESSO

    Exit Function

Erro_TRPVendedor_Le_Recursivo2:

    TRPVendedor_Le_Recursivo2 = gErr

    Select Case gErr
    
        Case 200785 'ERRO_LIMITE_RECURSIVIDADE
            Call Rotina_Erro(vbOKOnly, "ERRO_LIMITE_RECURSIVIDADE", gErr)
    
        Case 200786, 200787
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REGIOESDEVENDA", gErr)
    
        Case 200788
            Call Rotina_Erro(vbOKOnly, "ERRO_REGIAO_NAO_CADASTRADA", gErr, iRegiaoFilha)
            
        Case 200789, 200790, 200791
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_VENDEDORES", gErr)
            
        Case 200792
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200793)

    End Select

    Exit Function

End Function

Public Function TRPVouVendedores_Exclui(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal objVou As ClassTRPVouchers) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim iVend As Integer

On Error GoTo Erro_TRPVouVendedores_Exclui

    'Le a Vendedores do Voucher
    lErro = Comando_ExecutarPos(lComando1, "SELECT Vendedor FROM TRPVouVendedores WHERE TipVou = ? AND Serie = ? AND NumVou = ? ", _
                0, iVend, objVou.sTipVou, objVou.sSerie, objVou.lNumVou)
    If lErro <> AD_SQL_SUCESSO Then gError 200425

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200426

    Do While lErro <> AD_SQL_SEM_DADOS

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(lComando1)
        If lErro <> SUCESSO Then gError 200427
    
        lErro = Comando_ExecutarPos(lComando2, "DELETE FROM TRPVouVendedores", lComando1)
        If lErro <> AD_SQL_SUCESSO Then gError 200428

        'Busca Próximo
        lErro = Comando_BuscarProximo(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200429

    Loop

    TRPVouVendedores_Exclui = SUCESSO

    Exit Function

Erro_TRPVouVendedores_Exclui:

    TRPVouVendedores_Exclui = gErr

    Select Case gErr

        Case 200425, 200426, 200429
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPVOUVENDEDORES", gErr)

        Case 200427
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRPVOUVENDEDORES", gErr)

        Case 200428
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRPVOUVENDEDORES", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200430)

    End Select

    Exit Function

End Function

Public Function TRPVouContatos_Exclui(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal objVou As ClassTRPVouchers) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim iSeq As Integer

On Error GoTo Erro_TRPVouContatos_Exclui

    'Le a Contatos do Voucher
    lErro = Comando_ExecutarPos(lComando1, "SELECT Seq FROM TRPVouContatos WHERE TipoVou = ? AND Serie = ? AND NumVou = ? ", _
                0, iSeq, objVou.sTipVou, objVou.sSerie, objVou.lNumVou)
    If lErro <> AD_SQL_SUCESSO Then gError 200431

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200432

    Do While lErro <> AD_SQL_SEM_DADOS

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(lComando1)
        If lErro <> SUCESSO Then gError 200433
    
        lErro = Comando_ExecutarPos(lComando2, "DELETE FROM TRPVouContatos", lComando1)
        If lErro <> AD_SQL_SUCESSO Then gError 200434

        'Busca Próximo
        lErro = Comando_BuscarProximo(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200435

    Loop

    TRPVouContatos_Exclui = SUCESSO

    Exit Function

Erro_TRPVouContatos_Exclui:

    TRPVouContatos_Exclui = gErr

    Select Case gErr

        Case 200431, 200432, 200435
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPVOUCONTATOS", gErr)

        Case 200433
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRPVOUCONTATOS", gErr)

        Case 200434
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRPVOUCONTATOS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200436)

    End Select

    Exit Function

End Function

Public Function TRPVouPassageiros_Exclui(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal objVou As ClassTRPVouchers) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim iSeq As Integer

On Error GoTo Erro_TRPVouPassageiros_Exclui

    'Le a Passageiros do Voucher
    lErro = Comando_ExecutarPos(lComando1, "SELECT Seq FROM TRPVouPassageiros WHERE TipoVou = ? AND Serie = ? AND NumVou = ? ", _
                0, iSeq, objVou.sTipVou, objVou.sSerie, objVou.lNumVou)
    If lErro <> AD_SQL_SUCESSO Then gError 200437

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200438

    Do While lErro <> AD_SQL_SEM_DADOS

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(lComando1)
        If lErro <> SUCESSO Then gError 200439
    
        lErro = Comando_ExecutarPos(lComando2, "DELETE FROM TRPVouPassageiros", lComando1)
        If lErro <> AD_SQL_SUCESSO Then gError 200440

        'Busca Próximo
        lErro = Comando_BuscarProximo(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200441

    Loop

    TRPVouPassageiros_Exclui = SUCESSO

    Exit Function

Erro_TRPVouPassageiros_Exclui:

    TRPVouPassageiros_Exclui = gErr

    Select Case gErr

        Case 200437, 200438, 200441
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPVOUPASSAGEIROS", gErr)

        Case 200439
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRPVOUPASSAGEIROS", gErr)

        Case 200440
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRPVOUPASSAGEIROS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200442)

    End Select

    Exit Function

End Function

Public Function TRPVouContatos_Inserir(ByVal lComando1 As Long, ByVal objVou As ClassTRPVouchers) As Long

Dim lErro As Long
Dim objContato As ClassTRPVouContatos
Dim iSeq As Integer

On Error GoTo Erro_TRPVouContatos_Inserir

    For Each objContato In objVou.colContatos

        iSeq = iSeq + 1

        'Le a Contatos do Voucher
        lErro = Comando_Executar(lComando1, "INSERT INTO TRPVouContatos (TipoVou,Serie,NumVou,Seq,Nome,Telefone) VALUES (?,?,?,?,?,?) ", _
                    objVou.sTipVou, objVou.sSerie, objVou.lNumVou, iSeq, objContato.sNome, objContato.sTelefone)
        If lErro <> AD_SQL_SUCESSO Then gError 200443

    Next

    TRPVouContatos_Inserir = SUCESSO

    Exit Function

Erro_TRPVouContatos_Inserir:

    TRPVouContatos_Inserir = gErr

    Select Case gErr

        Case 200443
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERT_TRPVOUCONTATOS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200444)

    End Select

    Exit Function

End Function

Public Function TRPVouPassageiros_Inserir(ByVal lComando1 As Long, ByVal objVou As ClassTRPVouchers) As Long

Dim lErro As Long
Dim objPax As ClassTRPVouPassageiros
'Dim iSeq As Integer
Dim iQtdPax As Integer

On Error GoTo Erro_TRPVouPassageiros_Inserir

    iQtdPax = 0
    For Each objPax In objVou.colPassageiros

        'iSeq = iSeq + 1
        
        If objPax.iStatus <> STATUS_CANCELADO Then
            iQtdPax = iQtdPax + 1
        End If

        'Le a Passageiros do Voucher
        lErro = Comando_Executar(lComando1, "INSERT INTO TRPVouPassageiros (TipoVou,Serie,NumVou,Seq,Nome,DataNascimento,TipoDocumento,NumeroDocumento,Sexo,ValorPago,Status,Titular,ValorPagoEmi, PrimeiroNome, SobreNome) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) ", _
                    objVou.sTipVou, objVou.sSerie, objVou.lNumVou, objPax.iSeq, objPax.sNome, objPax.dtDataNascimento, objPax.sTipoDocumento, objPax.sNumeroDocumento, objPax.sSexo, objPax.dValorPago, objPax.iStatus, objPax.iTitular, objPax.dValorPagoEmi, objPax.sPrimeiroNome, objPax.sSobreNome)
        If lErro <> AD_SQL_SUCESSO Then gError 200445

    Next
    
    objVou.iPax = iQtdPax

    TRPVouPassageiros_Inserir = SUCESSO

    Exit Function

Erro_TRPVouPassageiros_Inserir:

    TRPVouPassageiros_Inserir = gErr

    Select Case gErr

        Case 200445
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERT_TRPVOUPASSAGEIROS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200446)

    End Select

    Exit Function

End Function

Function TiposDeCliente_Grava(ByVal objTipoCliente As ClassTipoCliente) As Long
'Inclui ou altera um Tipo de Cliente
'Tabelas:TiposDeCliente e TipoDeClienteCategoria

Dim lErro As Long
Dim lTransacao As Long
Dim objTipoFornecedor As New ClassTipoFornecedor

On Error GoTo Erro_TiposDeCliente_Grava

    'Inicia a Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 200631

    lErro = Tipo_Valida_Descricao(objTipoCliente.iCodigo, objTipoCliente.sDescricao)
    If lErro <> SUCESSO Then gError 200632

    lErro = CF("TiposDeCliente_Grava_EmTrans", objTipoCliente)
    If lErro <> SUCESSO Then gError 200633
    
    objTipoFornecedor.iCodigo = objTipoCliente.iCodigo
    
    lErro = CF("TipoFornecedor_Le", objTipoFornecedor)
    If lErro <> SUCESSO And lErro <> 12765 Then gError 200634
    
    objTipoFornecedor.sDescricao = objTipoCliente.sDescricao

    lErro = CF("TipoFornecedor_Grava_EmTrans", objTipoFornecedor)
    If lErro <> SUCESSO Then gError 200635

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 200636

    TiposDeCliente_Grava = SUCESSO

    Exit Function

Erro_TiposDeCliente_Grava:

    TiposDeCliente_Grava = gErr

    Select Case gErr

        Case 200631
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 200631 To 200635

        Case 200636
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200637)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function

Function TipoFornecedor_Grava(ByVal objTipoFornecedor As ClassTipoFornecedor) As Long
'Atualiza ou insere um novo registro na tabela TiposDeFornecedor

Dim lErro As Long
Dim lTransacao As Long
Dim objTipoCliente As New ClassTipoCliente

On Error GoTo Erro_TipoFornecedor_Grava

    'inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 200638
    
    lErro = Tipo_Valida_Descricao(objTipoFornecedor.iCodigo, objTipoFornecedor.sDescricao)
    If lErro <> SUCESSO Then gError 200639

    lErro = CF("TipoFornecedor_Grava_EmTrans", objTipoFornecedor)
    If lErro <> SUCESSO Then gError 200640
    
    objTipoCliente.iCodigo = objTipoFornecedor.iCodigo
    
    lErro = CF("TipoCliente_Le", objTipoCliente)
    If lErro <> SUCESSO And lErro <> 19062 Then gError 200641
  
    objTipoCliente.sDescricao = objTipoFornecedor.sDescricao

    lErro = CF("TiposDeCliente_Grava_EmTrans", objTipoCliente)
    If lErro <> SUCESSO Then gError 200642

    'Confirma a Transacao
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 200643

    TipoFornecedor_Grava = SUCESSO

    Exit Function

Erro_TipoFornecedor_Grava:

    TipoFornecedor_Grava = gErr

    Select Case gErr

        Case 200638
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 200639 To 200642

        Case 200643
             Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200644)

    End Select

    Call Transacao_Rollback
    
    Exit Function

End Function

Private Function Tipo_Valida_Descricao(ByVal iTipo As Integer, ByVal sDescricao As String) As Long

Dim sDescricaoOriginal As String

On Error GoTo Erro_Tipo_Valida_Descricao

    Select Case iTipo
    
        Case TRP_TIPO_CLIENTE_EMISSOR
            sDescricaoOriginal = TRP_TIPO_CLIENTE_EMISSOR_TEXTO
        Case TRP_TIPO_CLIENTE_FORNECEDORES
            sDescricaoOriginal = TRP_TIPO_CLIENTE_FORNECEDORES_TEXTO
        Case TRP_TIPO_CLIENTE_PASSAGEIRO
            sDescricaoOriginal = TRP_TIPO_CLIENTE_PASSAGEIRO_TEXTO
        Case Else
            sDescricaoOriginal = sDescricao

    End Select
    
    If UCase(sDescricao) <> UCase(sDescricaoOriginal) Then gError 200629

    Tipo_Valida_Descricao = SUCESSO

    Exit Function

Erro_Tipo_Valida_Descricao:

    Tipo_Valida_Descricao = gErr

    Select Case gErr

        Case 200629
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_DESCRICAO_NAO_ALTERAVEL", gErr, iTipo, sDescricaoOriginal)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200630)

    End Select

    
    Exit Function
    
End Function

Public Function TRPVou_Le_Comissao(ByVal objVou As ClassTRPVouchers) As Long

Dim lErro As Long
Dim alComando(1 To 11) As Long
Dim iIndice As Integer

On Error GoTo Erro_TRPVou_Le_Comissao

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 200647
    Next

    lErro = TRPVou_Gera_Comis_Insere_Novo(alComando(1), alComando(2), alComando(3), alComando(4), alComando(5), alComando(6), alComando(7), alComando(8), alComando(9), alComando(10), alComando(11), objVou, True)
    If lErro <> SUCESSO Then gError 200648
        
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    TRPVou_Le_Comissao = SUCESSO

    Exit Function

Erro_TRPVou_Le_Comissao:

    TRPVou_Le_Comissao = gErr

    Select Case gErr

        Case 200647
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 200648

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200649)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Public Function TRPFatCartao_Grava(ByVal colAuto As Collection) As Long
'Atualiza os Títulos de Cartão de Crédito a partir dos vouchers pagos com cartão

Dim lErro As Long
Dim lTransacao As Long
Dim alComando(1 To 7) As Long
Dim iIndice As Integer
Dim sConteudo As String
Dim lNumTituloIni As Long
Dim dValor As Double
Dim objCtb As ClassTRPCTBFatCR
Dim objContabAutomatica As ClassContabAutomatica
Dim objTRPConfig As ClassTRPConfig
Dim lProxTitRec As Long
Dim colConfig As Collection
Dim objTitRecTRP As ClassTitulosRecTRP
Dim sModeloFat As String
Dim sNomeDiretorio As String
Dim objTRPTitulosExp As ClassTRPTitulosExp
Dim objVoucherInfo As ClassTRPVoucherInfo
Dim objVouExp As ClassTRPVouExp
Dim objTRPTitulos As ClassTRPTitulos
Dim objVouAutoCC As ClassTRPVouAutoCC
Dim objVou As ClassTRPFATInfoVoucher
Dim objFrmAguarde As New ClassFrmAguarde
Dim objFrmAguardeTela As New FrmAguarde
Dim dValorAux As Double
Dim iCartaoAux As Integer
Dim lNumFatAux As Long
Dim sDocAux As String
Dim sItem As String
Dim objTitRec As ClassTituloReceber
Dim lNumIntDocAux As Long

On Error GoTo Erro_TRPFatCartao_Grava
    
    'Abre o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 200672
    Next
    
    objFrmAguarde.iTotalItens = colAuto.Count
    Call objFrmAguardeTela.Inicializa_Progressao(objFrmAguarde)

    lErro = CF("TRPConfig_Le", TRPCONFIG_DIRETORIO_MODELO_FAT_HTML_CARTAO, EMPRESA_TODA, sModeloFat)
    If lErro <> SUCESSO Then gError 200673
    
    lErro = CF("TRPConfig_Le", TRPCONFIG_DIRETORIO_FAT_HTML, EMPRESA_TODA, sNomeDiretorio)
    If lErro <> SUCESSO Then gError 200674
  
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 200675
  
    lErro = CF("TRPConfig_Le_Lockado", alComando(6), TRPCONFIG_PROX_NUM_TITREC, EMPRESA_TODA, sConteudo)
    If lErro <> SUCESSO Then gError 200676
    
    lProxTitRec = StrParaLong(sConteudo)
    lNumTituloIni = lProxTitRec
  
    For Each objVouAutoCC In colAuto
       
        lProxTitRec = lProxTitRec + 1
        
        Set objTitRec = New ClassTituloReceber
        
        objVouAutoCC.lNumTitulo = lProxTitRec
        
        'Gera o parcelamento do título e grava
        lErro = TRPFatCartao_Grava_TitRec(objVouAutoCC)
        If lErro <> SUCESSO Then gError 200677
                
        Set objTitRec = New ClassTituloReceber
                
        objTitRec.lNumIntDoc = objVouAutoCC.lNumIntTitulo
        
        lErro = CF("TituloReceber_Le", objTitRec)
        If lErro <> SUCESSO And lErro <> 26061 Then gError 200678
               
        Set objContabAutomatica = New ClassContabAutomatica
        Set objCtb = New ClassTRPCTBFatCR
        
        'gravar a contabilizacao
        lErro = objContabAutomatica.Inicializa_Contab(objCtb, MODULO_CONTASARECEBER, "CR", objTitRec.dtDataEmissao, objTitRec.dtDataEmissao)
        If lErro <> SUCESSO Then gError 200679
        
        lErro = objContabAutomatica.GeraContabilizacao(objTitRec)
        If lErro <> SUCESSO Then gError 200680
        
        Set objContabAutomatica = Nothing
        Call objCtb.Encerrando
        Set objCtb = Nothing
        
        Set objTitRecTRP = New ClassTitulosRecTRP
        
        objTitRecTRP.lNumIntDocTitRec = objTitRec.lNumIntDoc
        objTitRecTRP.dValorBruto = objVouAutoCC.dValorB
        objTitRecTRP.dValorDeducoes = objVouAutoCC.dValorDeducoes
        objTitRecTRP.dValorTarifa = objVouAutoCC.dValorTarifa
                
        'Atualiza a informação da Fatura no Voucher
        For Each objVou In objVouAutoCC.colVou
        
            lErro = Comando_ExecutarPos(alComando(1), "SELECT NumFat, Valor, Cartao FROM TRPVouchers WHERE NumVou = ? AND Serie = ? AND TipVou = ? ", 0, lNumFatAux, dValorAux, iCartaoAux, objVou.lNumVou, objVou.sSerie, objVou.sTipoVou)
            If lErro <> AD_SQL_SUCESSO Then gError 200681
            
            'Busca o primeiro
            lErro = Comando_BuscarPrimeiro(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200682
            
            If Abs(dValorAux - objVou.dValor) > DELTA_VALORMONETARIO Then
                sDocAux = "VOU " & objVou.sTipoVou & " " & objVou.sSerie & " " & CStr(objVou.lNumVou) & " Vlr.Ant: " & Format(objVou.dValor, "STANDARD") & " " & " Vlr.Atual: " & Format(dValorAux, "STANDARD")
                gError 200683
            End If
             If iCartaoAux = DESMARCADO Then
                sDocAux = "VOU " & objVou.sTipoVou & " " & objVou.sSerie & " " & CStr(objVou.lNumVou) & " Cartão: Não"
                gError 200683
            End If
            If lNumFatAux <> 0 Then
                sDocAux = "VOU " & objVou.sTipoVou & " " & objVou.sSerie & " " & CStr(objVou.lNumVou) & " Fatura: " & CStr(lNumFatAux)
                gError 200683
            End If
                        
            'Marca o arquivo como atualizado
            lErro = Comando_ExecutarPos(alComando(2), "UPDATE TRPVouchers SET NumFat = ?, TipoDocDestino = ?, NumIntDocDestino = ?", alComando(1), objTitRec.lNumTitulo, TRP_TIPO_DOC_DESTINO_TITREC, objTitRec.lNumIntDoc)
            If lErro <> SUCESSO Then gError 200685
            
            'Verifica se não tem que liberar o OVER (liberação na emissão)
            sItem = String(255, 0)
            
            'Verifica o tipo de tratamento (Se o pagamento é na baixa ou na emissão)
            lErro = Comando_Executar(alComando(3), "SELECT Item FROM FilialFornecedorCategorias WHERE Fornecedor = ? AND Filial = ? AND Categoria = ? AND Item = ?", sItem, objVou.lCliente, FILIAL_MATRIZ, TRP_CATEGORIA_LIBCME, TRP_CATEGORIA_LIBCME_FAT)
            If lErro <> AD_SQL_SUCESSO Then gError 182532
            
            'Busca a primeira
            lErro = Comando_BuscarProximo(alComando(3))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182533
                   
            'Se não for na baixa libera o over para pagamento
            If lErro = AD_SQL_SUCESSO Then
                
                'Le as comissões over bloqueadas para esse voucher
                lErro = Comando_ExecutarPos(alComando(4), "SELECT NumIntDoc FROM TRPVoucherInfo WHERE NumVou= ? AND TipoDoc = ? AND Serie = ? AND TipVou = ? AND Status = ? ", _
                0, lNumIntDocAux, objVou.lNumVou, TRP_TIPODOC_OVER_TEXTO, objVou.sSerie, objVou.sTipoVou, TRP_VOU_INFO_STATUS_BLOQUEADO)
                If lErro <> AD_SQL_SUCESSO Then gError 196376
            
                'Busca Primeiro
                lErro = Comando_BuscarPrimeiro(alComando(4))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196377
        
                Do While lErro <> AD_SQL_SEM_DADOS
                    
                    lErro = Comando_ExecutarPos(alComando(5), "UPDATE TRPVoucherInfo SET Status = ?, TipoLiberacao = ?, NumIntDocLiberacao = ? ", alComando(4), TRP_VOU_INFO_STATUS_LIBERADO, TRP_TIPO_LIBERACAO_COMISSAO_FAT, objTitRec.lNumIntDoc)
                    If lErro <> AD_SQL_SUCESSO Then gError 196379
                
                    'Busca Próximo
                    lErro = Comando_BuscarProximo(alComando(4))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196380
                
                Loop
                
            End If
                    
        Next

        lErro = CF("TitulosRecTRP_Grava", objTitRecTRP)
        If lErro <> SUCESSO Then gError 200686
        
        Set objTRPTitulosExp = New ClassTRPTitulosExp
        
        objTRPTitulosExp.iExcluido = DESMARCADO
        objTRPTitulosExp.iExportado = DESMARCADO
        objTRPTitulosExp.iTipoDocOrigem = TRP_TIPO_DOC_DESTINO_TITREC
        objTRPTitulosExp.lNumIntDocOrigem = objTitRec.lNumIntDoc
        objTRPTitulosExp.lNumTitulo = objTitRec.lNumTitulo
        objTRPTitulosExp.dValorDeducoes = objVouAutoCC.dValorDeducoes
        objTRPTitulosExp.dValorTarifa = objVouAutoCC.dValorTarifa
        
        lErro = CF("TRPTitulosExp_Insere", objTRPTitulosExp)
        If lErro <> SUCESSO Then gError 200687
        
        Set objTRPTitulos = New ClassTRPTitulos

        objTRPTitulos.iTipoDocOrigem = TRP_TIPO_DOC_DESTINO_TITREC
        objTRPTitulos.lNumIntDocOrigem = objTitRec.lNumIntDoc
        objTRPTitulos.sTipoDoc = TRP_TIPODOC_VOU_TEXTO
        
        lErro = CF("TRPTitulos_Insere", objTRPTitulos)
        If lErro <> SUCESSO Then gError 200688
   
        Call objFrmAguardeTela.ProcessouItem
        
        If objFrmAguarde.iCancelar = MARCADO Then Exit For

    Next
    
    Set objFrmAguardeTela = Nothing
    If objFrmAguarde.iCancelar = MARCADO Then gError 192084
    
    Set objTRPConfig = New ClassTRPConfig
    
    objTRPConfig.sCodigo = TRPCONFIG_PROX_NUM_TITREC
    objTRPConfig.sConteudo = CStr(lProxTitRec + 1)
    
    'lErro = CF("TRPConfig_Grava_Trans", colConfig)
    lErro = CF("TRPConfig_Grava_Lockado", alComando(6), alComando(7), objTRPConfig)
    If lErro <> SUCESSO Then gError 200690
    
    'gera o documento html
    lErro = CF("TRPFaturas_Regera_Html", lNumTituloIni, lProxTitRec, sModeloFat, sNomeDiretorio, objTitRec.sSiglaDocumento, False)
    If lErro <> SUCESSO Then gError 200689
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 200691
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
        
    TRPFatCartao_Grava = SUCESSO
     
    Exit Function
    
Erro_TRPFatCartao_Grava:

    TRPFatCartao_Grava = gErr
    
    Select Case gErr
          
        Case 200672
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 200673, 200674, 200676 To 200680, 200686 To 200690
        
        Case 200675
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
                
        Case 200681, 200682
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPVOUCHERS", gErr)
            
        Case 200683
            Call Rotina_Erro(vbOKOnly, "ERRO_VOU_VALOR_ALTERADO", gErr, sDocAux)
  
        Case 200685
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRPVOUCHERS", gErr)
                
        Case 200691
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182722)
     
    End Select
    
    'Desfaz a transação
    Call Transacao_Rollback
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    If Not (objFrmAguardeTela Is Nothing) Then
        Call objFrmAguardeTela.Trata_Erro
    End If
        
    Exit Function

End Function

Function TRPFatCartao_Grava_TitRec(objVouAutoCC As ClassTRPVouAutoCC) As Long

Dim lErro As Long
Dim colAdmMeioPagto As New Collection
Dim objAdmMeioPagto As New ClassAdmMeioPagto
Dim objRede As New ClassRede
Dim objAdmMeioPagtoCondPagto As New ClassAdmMeioPagtoCondPagto
Dim objTitRecCartao As New ClassTituloReceber
Dim colParcelasCartao As New colParcelaReceber
Dim iAchou As Integer
Dim objAdmMeioPagtoParcela As ClassAdmMeioPagtoParcelas
Dim colComissaoEmissao As New colComissao
Dim colcolComissao As New colcolComissao
Dim colComissaoFantasma As colComissao
Dim dtDataVencimento1 As Date
Dim dValor As Double
Dim lNumTitOriginal As Long
Dim lNumAuto As Long
Dim objParcelaRecCartao As ClassParcelaReceber
Dim dtDataVencimentoReal As Date
Dim iIndice As Integer
Dim dValorPresente As Double
Dim dValorTaxa As Double, dDeducaoValorPresente As Double, objParcelaRecCartao1 As ClassParcelaReceber

On Error GoTo Erro_TRPFatCartao_Grava_TitRec

    lErro = CF("AdmMeioPagto_Le_Todas1", FILIAL_MATRIZ, colAdmMeioPagto)
    If lErro <> SUCESSO Then gError 183111
    
    For Each objAdmMeioPagto In colAdmMeioPagto
        If objAdmMeioPagto.sNome = objVouAutoCC.sBandeira Then
            objAdmMeioPagto.iFilialEmpresa = FILIAL_MATRIZ
            iAchou = 1
            Exit For
        End If
    Next
    
    If iAchou = 0 Then gError 183112
    
    objRede.iCodigo = objAdmMeioPagto.iRede
    
    lErro = CF("Rede_Le", objRede)
    If lErro <> SUCESSO And lErro <> 104244 Then gError 183113
    
    If lErro <> SUCESSO Then gError 183114
    
    Set colComissaoFantasma = New colComissao
    Set colcolComissao = New colcolComissao

    'Lê em colCodPagtoLoja as Codiçoes de Pagamento relacionados a Admnistradora
    lErro = CF("AdmMeioPagtoCondPagto_Le", objAdmMeioPagto)
    If lErro <> SUCESSO And lErro <> 104086 Then gError 183120

    iAchou = 0

    For Each objAdmMeioPagtoCondPagto In objAdmMeioPagto.colCondPagtoLoja
        If objAdmMeioPagtoCondPagto.iNumParcelas = objVouAutoCC.iQuantParc Then
            iAchou = 1
            Exit For
        End If
    Next

    If iAchou = 0 Then gError 183121

    'se a coleção retornar vazia-> erro
    If objAdmMeioPagtoCondPagto.colParcelas.Count = 0 Then gError 183122

    objTitRecCartao.lNumTitulo = objVouAutoCC.lNumTitulo
    objTitRecCartao.sSiglaDocumento = TIPODOC_CARTAO_CRED_DEB
    objTitRecCartao.dtDataEmissao = objVouAutoCC.dtDataAutoCC
    dValor = objVouAutoCC.dValorB * (1 - objAdmMeioPagtoCondPagto.dTaxa - objAdmMeioPagtoCondPagto.dDesconto)
    objTitRecCartao.iNumParcelas = objAdmMeioPagtoCondPagto.iParcelasRecebto
    objTitRecCartao.iStatus = STATUS_LANCADO
    objTitRecCartao.dtDataRegistro = Date
    objTitRecCartao.lCliente = objRede.lCliente
    objTitRecCartao.iFilial = objRede.iFilialCli
    objTitRecCartao.iFilialEmpresa = objVouAutoCC.iFilialEmpresa
    objTitRecCartao.iMoeda = MOEDA_REAL
    objTitRecCartao.dtReajusteBase = DATA_NULA
    objTitRecCartao.dtReajustadoAte = DATA_NULA

    dtDataVencimento1 = objVouAutoCC.dtDataAutoCC

    dValorTaxa = objVouAutoCC.dValorB * (objAdmMeioPagtoCondPagto.dTaxa - objAdmMeioPagtoCondPagto.dDesconto)

    'varre a coleção de parcelas "modelo do admmeiopagtocondpagto
    For Each objAdmMeioPagtoParcela In objAdmMeioPagtoCondPagto.colParcelas
        
        'cria uma coleção de comissão fantasmas para a parcela em questão
        colcolComissao.Add colComissaoFantasma

        'cria uma nova parcela a receber
        Set objParcelaRecCartao = New ClassParcelaReceber
        
        'o valor da parcela é o valor total multiplicado pelo percentual da parcela
        objParcelaRecCartao.dValor = dValor * objAdmMeioPagtoParcela.dPercRecebimento
        objParcelaRecCartao.iNumParcela = objAdmMeioPagtoParcela.iParcela
        
        dtDataVencimento1 = dtDataVencimento1 + objAdmMeioPagtoParcela.iIntervaloRecebimento
        objParcelaRecCartao.dtDataVencimento = dtDataVencimento1

        'calcula seu vencimento real
        lErro = CF("DataVencto_Real", dtDataVencimento1, dtDataVencimentoReal)
        If lErro <> SUCESSO Then gError 183124

        objParcelaRecCartao.dtDataVencimentoReal = dtDataVencimentoReal
        objParcelaRecCartao.iStatus = STATUS_ABERTO
        objParcelaRecCartao.iCobrador = COBRADOR_PROPRIA_EMPRESA
        objParcelaRecCartao.iCarteiraCobranca = CARTEIRA_CARTEIRA
        
        'adiciona à coleção de parcelas a receber
        With objParcelaRecCartao
        '#################################################
        'ALTERADO POR WAGNER
            colParcelasCartao.Add 0, 0, .iNumParcela, .iStatus, .dtDataVencimento, .dtDataVencimentoReal, .dSaldo, .dValor, .iProxSeqBaixa, .iCarteiraCobranca, .iCobrador, .sNumTitCobrador, 0, 0, 0, 0, 0, 0, .iDesconto1Codigo, DATA_NULA, .dDesconto1Valor, .iDesconto2Codigo, DATA_NULA, .dDesconto2Valor, .iDesconto3Codigo, DATA_NULA, .dDesconto3Valor, .lNumIntCheque, .iAceite, .iDescontada, .iProxSeqOcorr, .iPrevisao, .sObservacao, .dValor
        '#################################################
        End With

        objTitRecCartao.dValor = objTitRecCartao.dValor + objParcelaRecCartao.dValor

    Next

    If objAdmMeioPagto.dTaxaVista > 0 Then
    
        Set colComissaoFantasma = New colComissao
        Set colcolComissao = New colcolComissao
    
        'cria uma coleção de comissão fantasmas para a parcela em questão
        colcolComissao.Add colComissaoFantasma
    
        'cria uma nova parcela a receber
        Set objParcelaRecCartao = New ClassParcelaReceber

        objParcelaRecCartao.iNumParcela = 1
        objTitRecCartao.iNumParcelas = 1

        objParcelaRecCartao.dtDataVencimento = dtDataVencimento1

        'calcula seu vencimento real
        lErro = CF("DataVencto_Real", dtDataVencimento1, dtDataVencimentoReal)
        If lErro <> SUCESSO Then gError 183124

        objParcelaRecCartao.dtDataVencimentoReal = dtDataVencimentoReal
        objParcelaRecCartao.iStatus = STATUS_ABERTO
        objParcelaRecCartao.iCobrador = COBRADOR_PROPRIA_EMPRESA
        objParcelaRecCartao.iCarteiraCobranca = CARTEIRA_CARTEIRA
        
        '******     INICIO DO CÁLCULO DO VALOR PRESENTE  **********
        dValorPresente = 0
    
        For Each objParcelaRecCartao1 In colParcelasCartao
            
            dValorPresente = dValorPresente + (objParcelaRecCartao1.dValor / ((1 + objAdmMeioPagto.dTaxaVista) ^ ((objParcelaRecCartao1.dtDataVencimento - objParcelaRecCartao.dtDataVencimento) / 30)))
            dDeducaoValorPresente = dDeducaoValorPresente + objParcelaRecCartao1.dValor - (objParcelaRecCartao1.dValor / ((1 + objAdmMeioPagto.dTaxaVista) ^ ((objParcelaRecCartao1.dtDataVencimento - objParcelaRecCartao.dtDataVencimento) / 30)))
        Next
    
        objParcelaRecCartao.dValor = dValorPresente
    
        '***********  FIM DO CÁLCULO DO VALOR PRESENTE  ***********
        
        Set colParcelasCartao = New colParcelaReceber
        
        'adiciona à coleção de parcelas a receber
        With objParcelaRecCartao
        '#################################################
        'ALTERADO POR WAGNER
            colParcelasCartao.Add 0, 0, .iNumParcela, .iStatus, .dtDataVencimento, .dtDataVencimentoReal, .dSaldo, .dValor, .iProxSeqBaixa, .iCarteiraCobranca, .iCobrador, .sNumTitCobrador, 0, 0, 0, 0, 0, 0, .iDesconto1Codigo, DATA_NULA, .dDesconto1Valor, .iDesconto2Codigo, DATA_NULA, .dDesconto2Valor, .iDesconto3Codigo, DATA_NULA, .dDesconto3Valor, .lNumIntCheque, .iAceite, .iDescontada, .iProxSeqOcorr, .iPrevisao, .sObservacao, .dValor
        '#################################################
        End With

        objTitRecCartao.dValor = objParcelaRecCartao.dValor

    End If


    objTitRecCartao.dSaldo = Arredonda_Moeda(objTitRecCartao.dValor, 2)

    dValor = 0

    For iIndice = 1 To colParcelasCartao.Count
        
        Set objParcelaRecCartao = colParcelasCartao.Item(iIndice)
        
        If iIndice = colParcelasCartao.Count Then
            objParcelaRecCartao.dValor = objTitRecCartao.dSaldo - dValor
        Else
        
            objParcelaRecCartao.dValor = Round(objParcelaRecCartao.dValor, 2)
            dValor = dValor + objParcelaRecCartao.dValor
        End If
    
        objParcelaRecCartao.dSaldo = objParcelaRecCartao.dValor
    
    Next

    objTitRecCartao.dSaldo = objTitRecCartao.dValor

    lErro = CF("TituloReceber_Insere", objTitRecCartao, colParcelasCartao, colComissaoEmissao, colcolComissao)
    If lErro <> SUCESSO Then gError 183125
    
    objVouAutoCC.lNumIntTitulo = objTitRecCartao.lNumIntDoc
    objVouAutoCC.dValorTarifa = dValorTaxa
    objVouAutoCC.dValorDeducoes = dDeducaoValorPresente
    objVouAutoCC.dValorL = objTitRecCartao.dValor

    TRPFatCartao_Grava_TitRec = SUCESSO

    Exit Function

Erro_TRPFatCartao_Grava_TitRec:

    TRPFatCartao_Grava_TitRec = gErr

    Select Case gErr
            
        Case 183111, 183113, 183120, 183123 To 183125

        Case 183112
            Call Rotina_Erro(vbOKOnly, "ERRO_CARTAO_INEXISTENTE", gErr, objVouAutoCC.sBandeira)
        
        Case 183114
            Call Rotina_Erro(vbOKOnly, "ERRO_REDE_NAO_ENCONTRADA", gErr, objRede.iCodigo)

        Case 183121
            Call Rotina_Erro(vbOKOnly, "ERRO_PARCELAMENTO_INEXISTENTE", gErr, objVouAutoCC.iQuantParc, objVouAutoCC.sBandeira)

        Case 183122
            Call Rotina_Erro(vbOKOnly, "ERRO_ADMMEIOPAGTOPARCELAS_NAOENCONTRADA", gErr, objAdmMeioPagtoCondPagto.iFilialEmpresa, objAdmMeioPagtoCondPagto.iAdmMeioPagto, objAdmMeioPagtoCondPagto.iParcelamento)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 183126)

    End Select

    Exit Function

End Function

Public Function TRPVoucher_Exclui(ByVal objTRPVoucher As ClassTRPVouchers) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_TRPVoucher_Exclui

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 192646

    lErro = CF("TRPVouchers_Cancela", objTRPVoucher)
    If lErro <> SUCESSO Then gError 192647

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 192648

    TRPVoucher_Exclui = SUCESSO

    Exit Function

Erro_TRPVoucher_Exclui:

    TRPVoucher_Exclui = gErr

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

Public Function TRPVoucher_Reativa(ByVal objTRPVoucher As ClassTRPVouchers) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_TRPVoucher_Reativa

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 192646

    lErro = TRPVouchers_Reativa_EmTrans(objTRPVoucher)
    If lErro <> SUCESSO Then gError 192647

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 192648

    TRPVoucher_Reativa = SUCESSO

    Exit Function

Erro_TRPVoucher_Reativa:

    TRPVoucher_Reativa = gErr

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

Public Function TRPVouchers_Reativa_EmTrans(ByVal objVou As ClassTRPVouchers) As Long

Dim lErro As Long
Dim alComando(0 To 13) As Long
Dim iIndice As Integer
Dim lNumFat As Long
Dim iStatus As Integer
Dim lNumIntDocTitRec As Long
Dim lNumIntDocDestino As Long
Dim iTipoDocDestino As Integer
Dim sTipoDoc As String
Dim objOcorrencia As ClassTRPOcorrencias
Dim objOcorrenciaDet As ClassTRPOcorrenciaDet
Dim dValor As Double
Dim lCliente As Long
Dim iFilialEmpresa As Integer
Dim sSiglaDocumento As String
Dim lNumVou As Long
Dim sSerie As String
Dim sTipVou As String
Dim lCodigo As Long
Dim iCartao As Integer
Dim lNumIntDocNVL As Long
Dim dValorOcr As Double
Dim dValorComissao As Double
Dim bComissaoCorp As Boolean
Dim lNumIntComi As Long
Dim dValorBruto As Double
Dim sTipoDocAux As String
Dim objVouExp As New ClassTRPVouExp
Dim lClienteVou As Long
Dim lCliPassageiro As Long

On Error GoTo Erro_TRPVouchers_Reativa_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 190848
    Next
    
    sTipoDoc = String(STRING_TRP_OCR_TIPODOC, 0)
    sSerie = String(STRING_TRP_OCR_SERIE, 0)
    sTipVou = String(STRING_TRP_OCR_TIPOVOU, 0)

    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumFat, Status, NumIntDocDestino,TipoDocDestino, Valor, ClienteFat, NumVou, Serie, TipVou, Cartao, NumIntDocNVL, ValorOCR, ValorBruto, ClienteVou, CliPassageiro FROM TRPVouchers WHERE NumVou = ? AND Serie = ? AND TipVou= ?", _
                        0, lNumFat, iStatus, lNumIntDocDestino, iTipoDocDestino, dValor, lCliente, lNumVou, sSerie, sTipVou, iCartao, lNumIntDocNVL, dValorOcr, dValorBruto, lClienteVou, lCliPassageiro, objVou.lNumVou, objVou.sSerie, objVou.sTipVou)
    If lErro <> AD_SQL_SUCESSO Then gError 190849

    'Busca Primeiro documento
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190850
    
    'Se existir
    If lErro <> AD_SQL_SEM_DADOS Then

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 190852
                   
        'Se já estiver ativo dá erro
        If iStatus = STATUS_TRP_VOU_ABERTO Then gError 192760
   
        'Se ainda não foi inativado
        If lNumIntDocNVL <> 0 Then
           
            Set objOcorrencia = New ClassTRPOcorrencias
                
            objOcorrencia.lNumIntDoc = lNumIntDocNVL
                
            lErro = TRPOcorrencias_Exclui_EmTrans(objOcorrencia)
            If lErro <> SUCESSO Then gError 192765
                                                    
            lErro = Comando_ExecutarPos(alComando(3), "UPDATE TRPVouchers SET NumIntDocNVL = ?", alComando(0), 0)
            If lErro <> AD_SQL_SUCESSO Then gError 192766
                    
        'Se não está criando o NVL e for voucher (voucher excluído) => Estorna as comissões
        Else
            
            'Reativou
            lErro = TRPVou_Gera_Comis(objVou, objVou.dValorBruto, TRP_TIPO_VALOR_BASE_BRU, gdtDataAtual, "Reativação do voucher.", TRP_TIPO_TRATAMENTO_COMI_VOU_VLR_ALTERADO)
            If lErro <> SUCESSO Then gError 196277
        
        End If
            
        lErro = Comando_ExecutarPos(alComando(1), "UPDATE TRPVouchers SET Status = ?, DataCanc  = ?, HoraCanc = ?, UsuarioCanc = ?", alComando(0), STATUS_TRP_VOU_ABERTO, DATA_NULA, 0, "")
        If lErro <> AD_SQL_SUCESSO Then gError 190846
                
    End If
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TRPVouchers_Reativa_EmTrans = SUCESSO

    Exit Function

Erro_TRPVouchers_Reativa_EmTrans:

    TRPVouchers_Reativa_EmTrans = gErr

    Select Case gErr

        Case 190846, 192766
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRPVOUCHERS", gErr)

        Case 190848
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 190849, 190850, 192761
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPVOUCHERS", gErr)

        Case 190852
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRPVOUCHERS", gErr)

        Case 190853
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRPVOUCHERS", gErr)
            
        Case 192760
            Call Rotina_Erro(vbOKOnly, "ERRO_VOUCHER_JA_CANCELADO", gErr)
            
        Case 192764, 192765, 192824, 200195, 196277
        
        Case 200710
            Call Rotina_Erro(vbOKOnly, "ERRO_CANC_VOU_COM_OCR", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190847)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function TRPEmissores_Monta_Hierarquia(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lComando3 As Long, ByVal objEmi As ClassCliEmissoresTRP) As Long

Dim sHierarquia As String
Dim lErro As Long
Dim lEmi As Long

On Error GoTo Erro_TRPEmissores_Monta_Hierarquia

    lErro = Comando_ExecutarPos(lComando1, "SELECT Codigo FROM TRPCliEmissores WHERE Cliente = ? ORDER BY Seq ", 0, lEmi, objEmi.lCliente)
    If lErro <> AD_SQL_SUCESSO Then gError 200353

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200354

    Do While lErro <> AD_SQL_SEM_DADOS
    
        sHierarquia = ""

        lErro = TRPEmissores_Le_Recursivo(lComando2, sHierarquia, lEmi, 1)
        If lErro <> SUCESSO Then gError 200389
    
        objEmi.sHierarquia = sHierarquia & CStr(lEmi) & "."
        
        lErro = Comando_ExecutarPos(lComando3, "UPDATE TRPCliEmissores SET Hierarquia = ? ", lComando1, objEmi.sHierarquia)
        If lErro <> AD_SQL_SUCESSO Then gError 200353
        
        'Busca Próximo
        lErro = Comando_BuscarProximo(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200354
        
    Loop

    TRPEmissores_Monta_Hierarquia = SUCESSO

    Exit Function

Erro_TRPEmissores_Monta_Hierarquia:

    TRPEmissores_Monta_Hierarquia = gErr

    Select Case gErr

        Case 200389

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190847)

    End Select

    Exit Function
    
End Function

Function TRPEmissores_Le_Recursivo(ByVal lComando As Long, ByRef sHierarquia As String, ByVal lFilho As Long, ByVal iNivel As Integer) As Long

Dim lErro As Long
Dim lPai As Long

On Error GoTo Erro_TRPEmissores_Le_Recursivo

    If lFilho <> 0 Then

        If iNivel > 20 Then gError 200385
       
        'Le a tabelaTRPProdTarifa
        lErro = Comando_Executar(lComando, "SELECT EmissorSuperior FROM TRPCliEmissores WHERE Codigo = ? ", _
                    lPai, lFilho)
        If lErro <> AD_SQL_SUCESSO Then gError 200386
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200387
    
        'Sem Dados
        'If lErro <> AD_SQL_SUCESSO Then gError 200388
        
        If lPai <> 0 Then
        
            lErro = TRPEmissores_Le_Recursivo(lComando, sHierarquia, lPai, iNivel + 1)
            If lErro <> SUCESSO Then gError 200389
        
            sHierarquia = sHierarquia & CStr(lPai) & "."
        
        End If
    
    End If

    TRPEmissores_Le_Recursivo = SUCESSO

    Exit Function

Erro_TRPEmissores_Le_Recursivo:

    TRPEmissores_Le_Recursivo = gErr

    Select Case gErr

        Case 200385 'ERRO_LIMITE_RECURSIVIDADE
            Call Rotina_Erro(vbOKOnly, "ERRO_LIMITE_RECURSIVIDADE", gErr)
    
        Case 200386, 200387
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPEMISSORES", gErr)

        Case 200388
            Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO", gErr, lFilho)
            
        Case 200389
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190847)

    End Select

    Exit Function
    
End Function

Function TRPGerComiInt_Gera(objTRPGerComiInt As ClassTRPGerComiInt) As Long

Dim lErro As Long
Dim alComando(0 To 20) As Long
Dim iIndice As Integer
Dim lTransacao As Long
Dim lCount As Long
Dim colVouchers As New Collection
Dim colVendedores As New Collection
Dim sCabecArqResVou As String
Dim sCabecArqResVend As String
Dim sCabecArqDetVend As String
Dim objVou As ClassTRPVouchers
Dim objVouInfo As ClassTRPVoucherInfo
Dim tVou As typeTRPVouchers
Dim iVendedor As Integer
Dim sNomeVendedor As String
Dim iCargoVendedor As Integer
Dim objVend As ClassVendedor
Dim objVendAux As ClassVendedor
Dim objVendTRP As ClassVendedorTRP
Dim bArqAberto As Boolean
Dim objComissao As ClassComissao
Dim colComissao As New Collection
Dim sNomeProm As String
Dim sNomeDir As String
Dim sNomeOut As String
Dim sNomeSup As String
Dim sNomeGer As String
Dim sDescProduto As String
Dim sNomeCliente As String
Dim bAchou As Boolean
Dim lNumIntTRPGerComiInt As Long
Dim lNumIntTRPGerComiIntDet As Long
Dim lNumIntCom As Long
Dim sSigla As String
Dim objFrmAguarde As New ClassFrmAguarde
Dim objFrmAguardeTela As New FrmAguarde

On Error GoTo Erro_TRPGerComiInt_Gera

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 200820
    
    sCabecArqResVou = "Nome Executivo de Contas;Nome Outros;Nome Supervisor;Nome Gerente;Nome Diretor;Cliente;Nome Cliente;Produto;Voucher;Emissão do Voucher;Valor Bruto; Valor Base Comissão"
    sCabecArqResVend = "Vendedor;Nome Vendedor;Valor Base Comissão; Valor da Comissão"
    sCabecArqDetVend = "Vendedor;Nome Vendedor;Cliente;Nome Cliente;Produto;Voucher;Emissão do Voucher;Valor Bruto; Valor Base Comissão"
    
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 200821
    Next
    
    If objTRPGerComiInt.iPrevia = DESMARCADO Then
    
        sSigla = ""
    
        lErro = CF("Config_ObterNumInt", "TRPConfig", "NUM_INT_PROX_TRPGERCOMIINT", lNumIntTRPGerComiInt)
        If lErro <> SUCESSO Then gError 200822
        
        objTRPGerComiInt.lNumIntDoc = lNumIntTRPGerComiInt
     
        lErro = Comando_Executar(alComando(4), "INSERT INTO TRPGerComiInt (NumIntDoc, DataGeracao, HoraGeracao, Usuario, DataEmiAte, Diretorio, Previa) VALUES (?,?,?,?,?,?,?)", _
        objTRPGerComiInt.lNumIntDoc, objTRPGerComiInt.dtDataGeracao, objTRPGerComiInt.dHoraGeracao, objTRPGerComiInt.sUsuario, objTRPGerComiInt.dtDataEmiAte, objTRPGerComiInt.sDiretorio, objTRPGerComiInt.iPrevia)
        If lErro <> AD_SQL_SUCESSO Then gError 200823
    
    Else
        sSigla = "PREV_"
    End If
    
    sNomeVendedor = String(STRING_VENDEDOR_NOME_REDUZIDO, 0)
    
    lErro = Comando_Executar(alComando(3), "SELECT V2.Vendedor, V1.NomeReduzido, V2.Cargo FROM Vendedores AS V1, VendedorTRP AS V2 WHERE V1.Codigo = V2.Vendedor ORDER BY V2.Vendedor ", _
                                               iVendedor, sNomeVendedor, iCargoVendedor)
    If lErro <> AD_SQL_SUCESSO Then gError 200824

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200825

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objVend = New ClassVendedor
        Set objVendTRP = New ClassVendedorTRP
        Set objVend.objInfoUsu = objVendTRP
        Set objComissao = New ClassComissao
        
        objVend.iCodigo = iVendedor
        objVend.sNomeReduzido = sNomeVendedor
        objVendTRP.iCargo = iCargoVendedor
        
        objComissao.iCodVendedor = iVendedor
        objComissao.iFilialEmpresa = giFilialEmpresa
        objComissao.iStatus = STATUS_LIBERADO
        objComissao.iTipoTitulo = TIPO_COMISSAO_TRP
        objComissao.dtDataGeracao = Date
        
        colVendedores.Add objVend
        colComissao.Add objComissao
    
        'Busca Próximo
        lErro = Comando_BuscarProximo(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200826
    
    Loop
    
    lErro = Comando_Executar(alComando(0), "SELECT COUNT(*) FROM TRPVouchers AS V WHERE V.Data <= ? AND V.GeraComissao = 0 AND EXISTS (SELECT I.NumIntDoc FROM TRPVoucherInfo AS I WHERE I.TipVou = V.TipVou AND I.Serie = V.Serie AND I.NumVou = I.NumVou AND I.NumIntDocComiInt = 0)", lCount, objTRPGerComiInt.dtDataEmiAte)
    If lErro <> AD_SQL_SUCESSO Then gError 200827

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200828
    
    objFrmAguarde.iTotalItens = lCount + (lCount * colVendedores.Count) + colVendedores.Count
    Call objFrmAguardeTela.Inicializa_Progressao(objFrmAguarde)
    
    Open objTRPGerComiInt.sDiretorio & sSigla & "CM_" & Format(Date, "yyyymmdd") & "_RES_VOU.csv" For Output As #1
    bArqAberto = True
    
    Print #1, sCabecArqResVou
    
    tVou.sSerie = String(STRING_TRP_OCR_SERIE, 0)
    tVou.sTipVou = String(STRING_TRP_OCR_TIPOVOU, 0)
    tVou.sProduto = String(STRING_PRODUTO, 0)
    sDescProduto = String(STRING_PRODUTO_DESCRICAO, 0)
    sNomeCliente = String(STRING_CLIENTE_NOME_REDUZIDO, 0)

    lErro = Comando_Executar(alComando(1), "SELECT V.TipVou, V.Serie, V.NumVou, V.Data, V.ClienteVou, V.Promotor, V.ComissaoProm,V.NumFat, V.Cartao, V.Status, V.Produto,V.ValorBruto, V.ValorBrutoComOCR, V.ValorCMAComOCR, V.ValorCMC, V.ValorCMR, V.ValorCMCC, V.ValorCME, P.Descricao, C.NomeReduzido FROM TRPVouchers AS V, Clientes AS C, Produtos AS P WHERE V.Produto = P.Codigo AND V.ClienteVou = C.Codigo AND V.Data <= ? AND V.GeraComissao = 0 AND EXISTS (SELECT I.NumIntDoc FROM TRPVoucherInfo AS I WHERE I.TipVou = V.TipVou AND I.Serie = V.Serie AND I.NumVou = I.NumVou AND I.NumIntDocComiInt = 0) ORDER BY NumVou ", _
                                               tVou.sTipVou, tVou.sSerie, tVou.lNumVou, tVou.dtData, tVou.lClienteVou, tVou.iPromotor, tVou.dComissaoProm, tVou.lNumFat, tVou.iCartao, tVou.iStatus, tVou.sProduto, tVou.dValorBruto, tVou.dValorBrutoComOCR, tVou.dValorCMAComOCR, tVou.dValorCMC, tVou.dValorCMR, tVou.dValorCMCC, tVou.dValorCME, sDescProduto, sNomeCliente, _
                                               objTRPGerComiInt.dtDataEmiAte)
    If lErro <> AD_SQL_SUCESSO Then gError 200829

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200830

    Do While lErro <> AD_SQL_SEM_DADOS
        
        Set objVou = New ClassTRPVouchers
        
        objVou.sTipVou = tVou.sTipVou
        objVou.sSerie = tVou.sSerie
        objVou.lNumVou = tVou.lNumVou
        objVou.dtData = tVou.dtData
        objVou.lClienteVou = tVou.lClienteVou
        objVou.iPromotor = tVou.iPromotor
        objVou.dComissaoProm = tVou.dComissaoProm
        objVou.lNumFat = tVou.lNumFat
        objVou.iCartao = tVou.iCartao
        objVou.iStatus = tVou.iStatus
        objVou.sProduto = tVou.sProduto & SEPARADOR & sDescProduto
        objVou.dValorBruto = tVou.dValorBruto
        objVou.dValorBrutoComOCR = tVou.dValorBrutoComOCR
        objVou.dValorCMAComOCR = tVou.dValorCMAComOCR
        objVou.dValorCMC = tVou.dValorCMC
        objVou.dValorCMR = tVou.dValorCMR
        objVou.dValorCMCC = tVou.dValorCMCC
        objVou.dValorCME = tVou.dValorCME
        objVou.sAeroportos = sNomeCliente
        
        lErro = CF("TRPVoucherInfo_Le", objVou, alComando(12))
        If lErro <> SUCESSO Then gError 200843
        
        'Calcula o Valor Net sem ter tido comissão gerada
        objVou.dValor = 0
        For Each objVouInfo In objVou.colTRPVoucherInfo
            If objVouInfo.lNumIntDocComiInt = 0 Then
                If objVouInfo.sTipoDoc = TRP_TIPODOC_BRUTO_TEXTO Then
                    objVou.dValor = objVou.dValor + objVouInfo.dValor
                Else
                    objVou.dValor = objVou.dValor - objVouInfo.dValor
                End If
            End If
        Next
        
        'objVou.dValor = objVou.dValorBrutoComOCR - objVou.dValorCMAComOCR - objVou.dValorCMC - objVou.dValorCMR - objVou.dValorCMCC - objVou.dValorCME
        
        If objVou.iPromotor <> 0 Then
            Set objVend = New ClassVendedor
            
            objVend.iCodigo = objVou.iPromotor
            objVend.dPercComissao = objVou.dComissaoProm
            
            objVou.colVendedores.Add objVend
        End If
        
        lErro = CF("TRPVouVendedores_Le", alComando(2), objVou)
        If lErro <> SUCESSO Then gError 200831
        
        colVouchers.Add objVou
        
        sNomeProm = ""
        sNomeDir = ""
        sNomeOut = ""
        sNomeSup = ""
        sNomeGer = ""
        
        For Each objVendAux In objVou.colVendedores
            If objVendAux.iCodigo <> 0 Then
                For Each objVend In colVendedores
                    If objVend.iCodigo = objVendAux.iCodigo Then
                        Exit For
                    End If
                Next
                If objVend.iCodigo <> 0 Then
                    Select Case objVend.objInfoUsu.iCargo
                        Case VENDEDOR_CARGO_PROMOTOR
                            sNomeProm = IIf(Len(Trim(sNomeProm)) = 0, "", sNomeProm & ", ") & objVend.iCodigo & SEPARADOR & objVend.sNomeReduzido
                        Case VENDEDOR_CARGO_SUPERVISOR
                            sNomeSup = IIf(Len(Trim(sNomeSup)) = 0, "", sNomeSup & ", ") & objVend.iCodigo & SEPARADOR & objVend.sNomeReduzido
                        Case VENDEDOR_CARGO_GERENTE
                            sNomeGer = IIf(Len(Trim(sNomeGer)) = 0, "", sNomeGer & ", ") & objVend.iCodigo & SEPARADOR & objVend.sNomeReduzido
                        Case VENDEDOR_CARGO_DIRETOR
                            sNomeDir = IIf(Len(Trim(sNomeDir)) = 0, "", sNomeDir & ", ") & objVend.iCodigo & SEPARADOR & objVend.sNomeReduzido
                        Case Else
                            sNomeOut = IIf(Len(Trim(sNomeOut)) = 0, "", sNomeOut & ", ") & objVend.iCodigo & SEPARADOR & objVend.sNomeReduzido
                    End Select
                End If
            End If
        Next
        
        Print #1, sNomeProm & ";" & sNomeOut & ";" & sNomeSup & ";" & sNomeGer & ";" & sNomeDir & ";" & CStr(objVou.lClienteVou) & ";" & objVou.sAeroportos & ";" & objVou.sProduto & ";" & objVou.sTipVou & objVou.sSerie & Format(objVou.lNumVou, "000000") & ";" & Format(objVou.dtData, "dd/mm/yyyy") & ";" & Format(objVou.dValorBruto, "STANDARD") & ";" & Format(objVou.dValor, "STANDARD")

        If objTRPGerComiInt.iPrevia = DESMARCADO Then

            lErro = Comando_Executar(alComando(5), "UPDATE TRPVoucherInfo SET NumIntDocComiInt = ? WHERE TipVou = ? AND Serie = ? AND NumVou = ?", objTRPGerComiInt.lNumIntDoc, objVou.sTipVou, objVou.sSerie, objVou.lNumVou)
            If lErro <> AD_SQL_SUCESSO Then gError 200832
        
        End If

        'Busca Próximo
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200833

        Call objFrmAguardeTela.ProcessouItem
        
        If objFrmAguarde.iCancelar = MARCADO Then Exit Do

    Loop
    
    If objFrmAguarde.iCancelar = MARCADO Then gError 200837
    
    Close #1
    bArqAberto = False
    
    iIndice = 0
    For Each objVend In colVendedores
        
        iIndice = iIndice + 1
        Set objComissao = colComissao.Item(iIndice)
        
        lErro = CF("Config_ObterNumInt", "TRPConfig", "NUM_INT_PROX_TRPGERCOMIINTDET", lNumIntTRPGerComiIntDet)
        If lErro <> SUCESSO Then gError 200834
        
        objComissao.lNumIntDoc = lNumIntTRPGerComiIntDet
    
        Open objTRPGerComiInt.sDiretorio & sSigla & "CM_" & Format(Date, "yyyymmdd") & "_DET_VEND_COD_" & Format(objVend.iCodigo, "0000") & ".csv" For Output As #1
        bArqAberto = True
        
        Print #1, sCabecArqDetVend
    
        For Each objVou In colVouchers
            
            bAchou = False
            For Each objVendAux In objVou.colVendedores
                If objVendAux.iCodigo = objVend.iCodigo Then
                    bAchou = True
                    Exit For
                End If
            Next
            
            If bAchou Then
            
                If objTRPGerComiInt.iPrevia = DESMARCADO Then
                 
                    lErro = Comando_Executar(alComando(6), "INSERT INTO TRPGerComiIntDetVou (NumIntComiIntDet, TipVou, Serie, NumVou, ValorBase, PercComi) VALUES (?,?,?,?,?,?)", _
                    lNumIntTRPGerComiIntDet, objVou.sTipVou, objVou.sSerie, objVou.lNumVou, objVou.dValor, objVendAux.dPercComissao)
                    If lErro <> AD_SQL_SUCESSO Then gError 200835
                
                    lErro = Comando_Executar(alComando(11), "INSERT INTO TRPVoucherInfoGer (NumIntComiIntDet, NumIntVoucherInfo) (SELECT ?, NumIntDoc FROM TRPVoucherInfo WHERE TipVou = ? AND Serie = ? AND NumVou = ?)", _
                    lNumIntTRPGerComiIntDet, objVou.sTipVou, objVou.sSerie, objVou.lNumVou)
                    If lErro <> AD_SQL_SUCESSO Then gError 200836
                
                End If
            
                objComissao.dValor = objComissao.dValor + Arredonda_Moeda(objVou.dValor * objVendAux.dPercComissao)
                objComissao.dValorBase = objComissao.dValorBase + objVou.dValor
        
                Print #1, CStr(objVendAux.iCodigo) & ";" & objVend.sNomeReduzido; ";" & CStr(objVou.lClienteVou) & ";" & objVou.sAeroportos & ";" & objVou.sProduto & ";" & objVou.sTipVou & objVou.sSerie & Format(objVou.lNumVou, "000000") & ";" & Format(objVou.dtData, "dd/mm/yyyy") & ";" & Format(objVou.dValorBruto, "STANDARD") & ";" & Format(objVou.dValor, "STANDARD")
        
            End If
        
            If objComissao.dValorBase > 0 Then
                objComissao.dPercentual = Arredonda_Moeda(objComissao.dValor / objComissao.dValorBase, 4)
            End If
                   
            Call objFrmAguardeTela.ProcessouItem
            
            If objFrmAguarde.iCancelar = MARCADO Then Exit For
        
        Next
        
        Close #1
        bArqAberto = False
        
        If objFrmAguarde.iCancelar = MARCADO Then Exit For
        
    Next
        
    If objFrmAguarde.iCancelar = MARCADO Then gError 200837
    
    Open objTRPGerComiInt.sDiretorio & sSigla & "CM_" & Format(Date, "yyyymmdd") & "_RES_VEND.csv" For Output As #1
    bArqAberto = True
    
    Print #1, sCabecArqResVend
        
    iIndice = 0
    For Each objVend In colVendedores
        
        iIndice = iIndice + 1
        Set objComissao = colComissao.Item(iIndice)
        
        lNumIntTRPGerComiIntDet = objComissao.lNumIntDoc
       
        If objTRPGerComiInt.iPrevia = DESMARCADO Then
         
            lNumIntCom = 0
            
            If objComissao.dValor <> 0 Then
                
                'Gera o próximo número interno para a nova Comissão
                lErro = CF("Comissao_Automatico1", alComando(9), alComando(10), lNumIntCom)
                If lErro <> SUCESSO Then gError 200838
                
                lErro = Comando_Executar(alComando(7), "INSERT INTO Comissoes (NumIntCom, FilialEmpresa, Status, TipoTitulo, NumIntDoc, CodVendedor, DataGeracao, DataBaixa, Percentual, ValorBase, Valor) VALUES (?,?,?,?,?,?,?,?,?,?,?)", _
                lNumIntCom, objComissao.iFilialEmpresa, objComissao.iStatus, objComissao.iTipoTitulo, lNumIntTRPGerComiIntDet, objComissao.iCodVendedor, objComissao.dtDataGeracao, DATA_NULA, objComissao.dPercentual, objComissao.dValorBase, objComissao.dValor)
                If lErro <> AD_SQL_SUCESSO Then gError 200839
                
            End If
        
            lErro = Comando_Executar(alComando(8), "INSERT INTO TRPGerComiIntDet (NumIntDoc, NumIntDocGerComi, Seq, NumIntDocComi, ValorBase, ValorComissao, Vendedor) VALUES (?,?,?,?,?,?,?)", _
            lNumIntTRPGerComiIntDet, objTRPGerComiInt.lNumIntDoc, iIndice, lNumIntCom, objComissao.dValorBase, objComissao.dValor, objComissao.iCodVendedor)
            If lErro <> AD_SQL_SUCESSO Then gError 200840
        
        End If
        
        Print #1, CStr(objVend.iCodigo) & ";" & objVend.sNomeReduzido; ";" & Format(objComissao.dValorBase, "STANDARD") & ";" & Format(objComissao.dValor, "STANDARD")
        
        Call objFrmAguardeTela.ProcessouItem
        
        If objFrmAguarde.iCancelar = MARCADO Then Exit For
    
    Next
    
    Close #1
    bArqAberto = False
    
    Set objFrmAguardeTela = Nothing
    If objFrmAguarde.iCancelar = MARCADO Then gError 200837
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 200841

    TRPGerComiInt_Gera = SUCESSO

    Exit Function

Erro_TRPGerComiInt_Gera:

    TRPGerComiInt_Gera = gErr

    Select Case gErr

        Case 200820
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 200821
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 200822, 200831, 200834, 200837, 200838, 200843
        
        Case 200823
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRPGERCOMIINT", gErr)
            
        Case 200824 To 200826
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_VENDEDORES", gErr)
            
        Case 200827 To 200830, 200833
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPVOUCHERS", gErr)
            
        Case 200832
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRPVOUCHERINFO", gErr)

        Case 200835
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRPGERCOMIINTDETVOU", gErr)

        Case 200836
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRPVOUCHERINFOGER", gErr)

        Case 200839
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_COMISSOES", gErr)

        Case 200840
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRPGERCOMIINTDET", gErr)

        Case 200841
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200842)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Desfaz Transação
    Call Transacao_Rollback
 
    If Not (objFrmAguardeTela Is Nothing) Then
        Call objFrmAguardeTela.Trata_Erro
    End If

    If bArqAberto Then Close #1
    
    Exit Function

End Function

Private Function TRPVouHistorico_Grava(ByVal objVouTela As ClassTRPVouchers, ByVal objVouBD As ClassTRPVouchers)

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lNumVouAux As Long
Dim iEmVigencia As Integer
Dim iFaturado As Integer

On Error GoTo Erro_TRPVouHistorico_Grava

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 202731
    Next

    If Date > objVouBD.dtDataVigenciaDe Then
        iEmVigencia = MARCADO
    Else
        iEmVigencia = DESMARCADO
    End If
    If objVouBD.lNumFat <> 0 Then
        iFaturado = MARCADO
    Else
        iFaturado = DESMARCADO
    End If
        
    'Le a tabela TRPVouHistorico
    lErro = Comando_Executar(alComando(0), "SELECT NumVou FROM TRPVouHistorico WHERE TipVou = ? AND Serie= ? AND NumVou = ? ", _
                lNumVouAux, objVouTela.sTipVou, objVouTela.sSerie, objVouTela.lNumVou)
    If lErro <> AD_SQL_SUCESSO Then gError 202732

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202733

    'Se existir => TESTA ALTERAÇÃO
    If lErro = AD_SQL_SUCESSO Then

        'Se alterou algo insere
        If objVouTela.sProduto <> objVouBD.sProduto Or objVouTela.lClienteVou <> objVouBD.lClienteVou Or _
            objVouTela.dtDataVigenciaDe <> objVouBD.dtDataVigenciaDe Or objVouTela.dtDataVigenciaAte <> objVouBD.dtDataVigenciaAte Or _
            Abs(objVouTela.dComissaoAg - objVouBD.dComissaoAg) > DELTA_VALORMONETARIO Or _
            Abs(objVouTela.dComissaoEmissor - objVouBD.dComissaoEmissor) > DELTA_VALORMONETARIO Or _
            Abs(objVouTela.dValorBruto - objVouBD.dValorBruto) > DELTA_VALORMONETARIO Or _
            Abs(objVouTela.dValor - objVouBD.dValor) > DELTA_VALORMONETARIO Or _
            Abs(objVouTela.dCambio - objVouBD.dCambio) > DELTA_VALORMONETARIO Or _
            Abs(objVouTela.dComissaoRep - objVouBD.dComissaoRep) > DELTA_VALORMONETARIO Or _
            Abs(objVouTela.dComissaoCorr - objVouBD.dComissaoCorr) > DELTA_VALORMONETARIO Or _
            objVouTela.lEmissor <> objVouBD.lEmissor Or objVouTela.iDestino <> objVouBD.iDestino Or _
            objVouTela.iMoeda <> objVouBD.iMoeda Or objVouTela.iCartao <> objVouBD.iCartao Or _
            objVouTela.sTitular <> objVouBD.sTitular Or objVouTela.sTitularCPF <> objVouBD.sTitularCPF Or _
            objVouTela.sCiaCart <> objVouBD.sCiaCart Or objVouTela.sNumCCred <> objVouBD.sNumCCred Or _
            objVouTela.iQuantParc <> objVouBD.iQuantParc Or objVouTela.lNumAuto <> objVouBD.lNumAuto Or _
            objVouTela.dtDataAutoCC <> objVouBD.dtDataAutoCC Or objVouTela.lRepresentante <> objVouBD.lRepresentante Or _
            objVouTela.lCorrentista <> objVouBD.lCorrentista _
        Then

            With objVouTela
                lErro = Comando_Executar(alComando(1), "INSERT INTO TRPVouHistorico(TipVou,Serie,NumVou,DataLog,HoraLog,UsuarioLog,DataLibManut,HoraLibManut,UsuarioLibManut,EmVigencia,Faturado,Produto,DataVigenciaDe,DataVigenciaAte,Cliente,PercComissaoAg,Emissor,PercComissaoEmi,Destino,Moeda,Cartao,TitularCC,TitularCCCPF,BandeiraCC,NumCCred,ParcelamentoCC,NumAutorizacao,DataAutorizacao,ValorBrutoRS,ValorFaturavelRS,Cambio,Representante,PercComissaoRepr,Correntista,PercComissaoCorr) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                    .sTipVou, .sSerie, .lNumVou, Date, CDbl(Time), gsUsuario, .dtDataLibManut, .dHoraLibManut, .sUsuarioLibManut, iEmVigencia, iFaturado, .sProduto, .dtDataVigenciaDe, .dtDataVigenciaAte, .lClienteVou, .dComissaoAg, .lEmissor, .dComissaoEmissor, .iDestino, .iMoeda, .iCartao, .sTitular, .sTitularCPF, .sCiaCart, .sNumCCred, .iQuantParc, CStr(.lNumAuto), .dtDataAutoCC, .dValorBruto, .dValor, .dCambio, .lRepresentante, .dComissaoRep, .lCorrentista, .dComissaoCorr)
            End With
            If lErro <> AD_SQL_SUCESSO Then gError 202734
            
        End If

    Else
    
        With objVouTela
            lErro = Comando_Executar(alComando(1), "INSERT INTO TRPVouHistorico(TipVou,Serie,NumVou,DataLog,HoraLog,UsuarioLog,DataLibManut,HoraLibManut,UsuarioLibManut,EmVigencia,Faturado,Produto,DataVigenciaDe,DataVigenciaAte,Cliente,PercComissaoAg,Emissor,PercComissaoEmi,Destino,Moeda,Cartao,TitularCC,TitularCCCPF,BandeiraCC,NumCCred,ParcelamentoCC,NumAutorizacao,DataAutorizacao,ValorBrutoRS,ValorFaturavelRS,Cambio,Representante,PercComissaoRepr,Correntista,PercComissaoCorr) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                .sTipVou, .sSerie, .lNumVou, Date, CDbl(Time), gsUsuario, .dtDataLibManut, .dHoraLibManut, .sUsuarioLibManut, iEmVigencia, iFaturado, .sProduto, .dtDataVigenciaDe, .dtDataVigenciaAte, .lClienteVou, .dComissaoAg, .lEmissor, .dComissaoEmissor, .iDestino, .iMoeda, .iCartao, .sTitular, .sTitularCPF, .sCiaCart, .sNumCCred, .iQuantParc, CStr(.lNumAuto), .dtDataAutoCC, .dValorBruto, .dValor, .dCambio, .lRepresentante, .dComissaoRep, .lCorrentista, .dComissaoCorr)
        End With
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202735

    End If

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TRPVouHistorico_Grava = SUCESSO

    Exit Function

Erro_TRPVouHistorico_Grava:

    TRPVouHistorico_Grava = gErr

    Select Case gErr

        Case 202731
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 202732, 202733
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPVOUHISTORICO", gErr)

        Case 202734, 202735
            Call Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_TRPVOUHISTORICO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202736)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function
    
End Function

Public Function TitRec_Testa_Repeticao_Insert_Cust(ByVal objTitRec As ClassTituloReceber)

On Error GoTo Erro_TitRec_Testa_Repeticao_Insert_Cust

    If objTitRec.sSiglaDocumento = TIPODOC_FATURA_SERVICO_CR Then gError 202991

    TitRec_Testa_Repeticao_Insert_Cust = SUCESSO

    Exit Function

Erro_TitRec_Testa_Repeticao_Insert_Cust:

    TitRec_Testa_Repeticao_Insert_Cust = gErr

    Select Case gErr

        Case 202991
            Call Rotina_Erro(vbOKOnly, "ERRO_TITREC_FSCR_NAOTRPFAT", gErr, TIPODOC_CARTAO_CRED_DEB, TIPODOC_FATURA_SERVICO_CR)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202990)

    End Select
    
End Function

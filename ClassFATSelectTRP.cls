VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFATSelectTRP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function Comando_BindVarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_BindVar" (ByVal lComando As Long, lpVar As Variant) As Long
Private Declare Function Comando_PrepararInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Preparar" (ByVal lComando As Long, ByVal lpSQLStmt As String) As Long
Private Declare Function Comando_ExecutarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Executar" (ByVal lComando As Long) As Long


'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function TRPOcorrencias_Le(ByVal objOcr As ClassTRPOcorrencias) As Long

Dim lErro As Long
Dim lComando As Long
Dim tTRPOcorrencias As typeTRPOcorrencias

On Error GoTo Erro_TRPOcorrencias_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 189437

    'Alocação de espaço no buffer
    tTRPOcorrencias.sTipoDoc = String(STRING_TRP_OCR_TIPODOC, 0)
    tTRPOcorrencias.sSerie = String(STRING_TRP_OCR_SERIE, 0)
    tTRPOcorrencias.sObservacao = String(STRING_TRP_OCR_OBS, 0)
    tTRPOcorrencias.sHistorico = String(STRING_TRP_OCR_HISTORICO, 0)

    'Le a tabelaTRPOcorrencias
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumVou, TipoDoc, Serie, Codigo, Cliente, " & _
                "FilialCliente, DataEmissao, Observacao, Status, Origem, Historico, FormaPagto, " & _
                "NumIntDocDestino, ValorTotal, NumTitulo, TipoDocDestino,ValorOCRBrutoVou,ValorOCRCMAVou  FROM TRPOcorrencias WHERE Codigo= ? ", _
                tTRPOcorrencias.lNumIntDoc, tTRPOcorrencias.lNumVou, tTRPOcorrencias.sTipoDoc, tTRPOcorrencias.sSerie, _
                tTRPOcorrencias.lCodigo, tTRPOcorrencias.lCliente, tTRPOcorrencias.iFilialCliente, tTRPOcorrencias.dtDataEmissao, tTRPOcorrencias.sObservacao, _
                tTRPOcorrencias.iStatus, tTRPOcorrencias.iOrigem, tTRPOcorrencias.sHistorico, tTRPOcorrencias.iFormaPagto, tTRPOcorrencias.lNumIntDocDestino, _
                tTRPOcorrencias.dValorTotal, tTRPOcorrencias.lNumDocDestino, tTRPOcorrencias.iTipoDocDestino, tTRPOcorrencias.dValorOCRBrutoVou, tTRPOcorrencias.dValorOCRCMAVou, _
                objOcr.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 189438

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189439

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objOcr.lNumIntDoc = tTRPOcorrencias.lNumIntDoc
    objOcr.lNumVou = tTRPOcorrencias.lNumVou
    objOcr.sTipoDoc = tTRPOcorrencias.sTipoDoc
    objOcr.sSerie = tTRPOcorrencias.sSerie
    objOcr.lCodigo = tTRPOcorrencias.lCodigo
    objOcr.lCliente = tTRPOcorrencias.lCliente
    objOcr.iFilialCliente = tTRPOcorrencias.iFilialCliente
    objOcr.dtDataEmissao = tTRPOcorrencias.dtDataEmissao
    objOcr.sObservacao = tTRPOcorrencias.sObservacao
    objOcr.iStatus = tTRPOcorrencias.iStatus
    objOcr.iOrigem = tTRPOcorrencias.iOrigem
    objOcr.sHistorico = tTRPOcorrencias.sHistorico
    objOcr.iFormaPagto = tTRPOcorrencias.iFormaPagto
    objOcr.lNumIntDocDestino = tTRPOcorrencias.lNumIntDocDestino
    objOcr.dValorTotal = tTRPOcorrencias.dValorTotal
    objOcr.lNumDocDestino = tTRPOcorrencias.lNumDocDestino
    objOcr.iTipoDocDestino = tTRPOcorrencias.iTipoDocDestino
    objOcr.dValorOCRBrutoVou = tTRPOcorrencias.dValorOCRBrutoVou
    objOcr.dValorOCRCMAVou = tTRPOcorrencias.dValorOCRCMAVou
    
    lErro = TRPOcorrenciaDet_Le(objOcr)
    If lErro <> SUCESSO Then gError 189484

    'Fecha Comando
    Call Comando_Fechar(lComando)

    TRPOcorrencias_Le = SUCESSO

    Exit Function

Erro_TRPOcorrencias_Le:

    TRPOcorrencias_Le = gErr

    Select Case gErr

        Case 189437
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 189438, 189439
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPOCORRENCIAS", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora
        
        Case 189484

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189440)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Private Function TRPOcorrenciaDet_Le(ByVal objOcr As ClassTRPOcorrencias) As Long

Dim lErro As Long
Dim lComando As Long
Dim tTRPOcorrenciaDet As typeTRPOcorrenciaDet
Dim objTRPOcorrenciaDet As ClassTRPOcorrenciaDet

On Error GoTo Erro_TRPOcorrenciaDet_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 189460

    'Alocação de espaço no buffer

    'Le a tabelaTRPOcorrenciaDet
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntDocOcr, Tipo, Valor, Seq FROM TRPOcorrenciaDet WHERE NumIntDocOcr= ? ORDER BY Seq", _
                tTRPOcorrenciaDet.lNumIntDoc, tTRPOcorrenciaDet.lNumIntDocOCR, tTRPOcorrenciaDet.iTipo, tTRPOcorrenciaDet.dValor, tTRPOcorrenciaDet.iSeq, _
                objOcr.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 189461

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189462

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objTRPOcorrenciaDet = New ClassTRPOcorrenciaDet

        objTRPOcorrenciaDet.lNumIntDoc = tTRPOcorrenciaDet.lNumIntDoc
        objTRPOcorrenciaDet.lNumIntDocOCR = tTRPOcorrenciaDet.lNumIntDocOCR
        objTRPOcorrenciaDet.iTipo = tTRPOcorrenciaDet.iTipo
        objTRPOcorrenciaDet.dValor = tTRPOcorrenciaDet.dValor
        objTRPOcorrenciaDet.iSeq = tTRPOcorrenciaDet.iSeq
        
        objOcr.colDetalhamento.Add objTRPOcorrenciaDet

        'Busca Próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189463

    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    TRPOcorrenciaDet_Le = SUCESSO

    Exit Function

Erro_TRPOcorrenciaDet_Le:

    TRPOcorrenciaDet_Le = gErr

    Select Case gErr

        Case 189460
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 189461, 189462, 189463
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPOCORRENCIADET", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189483)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function Carrega_Combo_Historico(ByVal objCombo As Object, ByVal sTabela As String, ByVal iTamanho As Integer, Optional ByVal sCampo As String = "Historico", Optional ByVal sFiltro As String = "") As Long

Dim lErro As Long
Dim lComando As Long
Dim sHistorico As String
Dim sSQL As String

On Error GoTo Erro_Carrega_Combo_Historico

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 190147

    'Alocação de espaço no buffer
    sHistorico = String(iTamanho, 0)
    
    If Len(Trim(sFiltro)) > 0 Then
        sSQL = "SELECT DISTINCT " & sCampo & " FROM " & sTabela & " WHERE " & sFiltro & " ORDER BY " & sCampo
    Else
        sSQL = "SELECT DISTINCT " & sCampo & " FROM " & sTabela & " ORDER BY " & sCampo
    End If

    'Le a tabelaTRPOcorrencias
    lErro = Comando_Executar(lComando, sSQL, sHistorico)
    If lErro <> AD_SQL_SUCESSO Then gError 190148

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190149

    objCombo.Clear

    Do While lErro <> AD_SQL_SEM_DADOS
    
        objCombo.AddItem sHistorico
  
        'Busca Próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190150
    
    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Carrega_Combo_Historico = SUCESSO

    Exit Function

Erro_Carrega_Combo_Historico:

    Carrega_Combo_Historico = gErr

    Select Case gErr

        Case 190147
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 190148 To 190150
            Call Rotina_Erro(vbOKOnly, "ERRO_PREPARACAO_COMANDO_SQL", gErr, sSQL)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190151)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function Carrega_Combo_Data(ByVal objCombo As Object, ByVal sTabela As String, ByVal sCampo As String, Optional ByVal sFiltro As String = "") As Long

Dim lErro As Long
Dim lComando As Long
Dim dtData As Date
Dim sSQL As String

On Error GoTo Erro_Carrega_Combo_Data

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 190147
    
    If Len(Trim(sFiltro)) > 0 Then
        sSQL = "SELECT DISTINCT " & sCampo & " FROM " & sTabela & " WHERE " & sFiltro & " ORDER BY " & sCampo
    Else
        sSQL = "SELECT DISTINCT " & sCampo & " FROM " & sTabela & " ORDER BY " & sCampo
    End If

    'Le a tabelaTRPOcorrencias
    lErro = Comando_Executar(lComando, sSQL, dtData)
    If lErro <> AD_SQL_SUCESSO Then gError 190148

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190149

    objCombo.Clear

    Do While lErro <> AD_SQL_SEM_DADOS
    
        objCombo.AddItem Format(dtData, "dd/mm/yyyy")
  
        'Busca Próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190150
    
    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Carrega_Combo_Data = SUCESSO

    Exit Function

Erro_Carrega_Combo_Data:

    Carrega_Combo_Data = gErr

    Select Case gErr

        Case 190147
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 190148 To 190150
            Call Rotina_Erro(vbOKOnly, "ERRO_PREPARACAO_COMANDO_SQL", gErr, sSQL)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190151)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function Carrega_Combo_WebProduto(ByVal objCombo As Object, ByVal iDestino As Integer) As Long

Dim lErro As Long
Dim lComando As Long
Dim sCodigo As String
Dim sDescricao As String
Dim sSQL As String

On Error GoTo Erro_Carrega_Combo_WebProduto

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 190147

    'Alocação de espaço no buffer
    sCodigo = String(STRING_PRODUTO, 0)
    sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)

    sSQL = "SELECT Produto, Descricao FROM WebProdutoDestino WHERE Destino = ? ORDER BY Produto"

    'Le a tabelaTRPOcorrencias
    lErro = Comando_Executar(lComando, sSQL, sCodigo, sDescricao, iDestino)
    If lErro <> AD_SQL_SUCESSO Then gError 190148

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190149

    objCombo.Clear

    Do While lErro <> AD_SQL_SEM_DADOS
    
        objCombo.AddItem Trim(sCodigo) & SEPARADOR & sDescricao
  
        'Busca Próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190150
    
    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Carrega_Combo_WebProduto = SUCESSO

    Exit Function

Erro_Carrega_Combo_WebProduto:

    Carrega_Combo_WebProduto = gErr

    Select Case gErr

        Case 190147
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 190148 To 190150
            Call Rotina_Erro(vbOKOnly, "ERRO_PREPARACAO_COMANDO_SQL", gErr, sSQL)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190151)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function TRPOcorrencias_Le_Bloqueios(ByVal objTRPLiberaOCRSel As ClassTRPLiberaOCRSel) As Long
 
Dim lErro As Long
Dim iIndice As Integer
Dim sSQL As String
Dim sWhere As String
Dim vlCodigo As Variant
Dim viStatus As Variant
Dim alComando(1 To 1) As Long
Dim objOcorrencia As ClassTRPOcorrencias
Dim vlCodigoDe As Variant
Dim vlCodigoAte As Variant
Dim vdtEmissaoDe As Variant
Dim vdtEmissaoAte As Variant

On Error GoTo Erro_TRPOcorrencias_Le_Bloqueios
    
    vlCodigoDe = objTRPLiberaOCRSel.lCodigoDe
    vlCodigoAte = objTRPLiberaOCRSel.lCodigoAte
    vdtEmissaoDe = objTRPLiberaOCRSel.dtDataEmissaoDe
    vdtEmissaoAte = objTRPLiberaOCRSel.dtDataEmissaoAte
    vlCodigo = CLng(0)
    viStatus = STATUS_TRP_OCR_BLOQUEADO
 
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 190244
    Next
    
    sSQL = "SELECT O.Codigo FROM TRPOcorrencias AS O WHERE O.Status = ? "
    
    If objTRPLiberaOCRSel.iPago = TRP_VOU_NAO_PAGO Then
        sWhere = sWhere & "AND NOT EXISTS (SELECT T.NumIntDoc FROM TRPVouchers AS V, TitulosRecTodos AS T WHERE V.TipoDocDestino =3 AND V.NumIntDocDestino = T.NumIntDoc AND T.Saldo + 0.01 < T.Valor AND O.NumVou = V.NumVou AND O.TipoDoc = V.TipVou AND O.Serie = V.Serie)"
    ElseIf objTRPLiberaOCRSel.iPago = TRP_VOU_PAGO Then
        sWhere = sWhere & "AND EXISTS (SELECT T.NumIntDoc FROM TRPVouchers AS V, TitulosRecTodos AS T WHERE V.TipoDocDestino =3 AND V.NumIntDocDestino = T.NumIntDoc AND T.Saldo + 0.01 < T.Valor AND O.NumVou = V.NumVou AND O.TipoDoc = V.TipVou AND O.Serie = V.Serie)"
    End If
 
    If vlCodigoDe <> 0 Then
        sWhere = sWhere & "AND O.Codigo >= ? "
    End If
    If vlCodigoAte <> 0 Then
        sWhere = sWhere & "AND O.Codigo <= ? "
    End If

    If vdtEmissaoDe <> DATA_NULA Then
        sWhere = sWhere & "AND O.DataEmissao >= ? "
    End If
    If vdtEmissaoAte <> DATA_NULA Then
        sWhere = sWhere & "AND O.DataEmissao <= ? "
    End If
    
    sSQL = sSQL & sWhere
    
    sSQL = sSQL & "ORDER BY O.Codigo "
 
    lErro = Comando_PrepararInt(alComando(1), sSQL)
    If (lErro <> AD_SQL_SUCESSO) Then gError 190245
 
    lErro = Comando_BindVarInt(alComando(1), vlCodigo)
    If (lErro <> AD_SQL_SUCESSO) Then gError 190246
 
    lErro = Comando_BindVarInt(alComando(1), viStatus)
    If (lErro <> AD_SQL_SUCESSO) Then gError 190247
 
    If vlCodigoDe <> 0 Then
        lErro = Comando_BindVarInt(alComando(1), vlCodigoDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError 190248
    End If
    
    If vlCodigoAte <> 0 Then
        lErro = Comando_BindVarInt(alComando(1), vlCodigoAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError 190249
    End If
 
    If vdtEmissaoDe <> DATA_NULA Then
        lErro = Comando_BindVarInt(alComando(1), vdtEmissaoDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError 190250
    End If
    
    If vdtEmissaoAte <> DATA_NULA Then
        lErro = Comando_BindVarInt(alComando(1), vdtEmissaoAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError 190251
    End If
    
    lErro = Comando_ExecutarInt(alComando(1))
    If (lErro <> AD_SQL_SUCESSO) Then gError 190252
 
    'Busca o primeiro registro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190253
        
    'Para cada Produto
    Do While lErro = SUCESSO
    
        Set objOcorrencia = New ClassTRPOcorrencias

        objOcorrencia.lCodigo = vlCodigo
        
        'Lê o produto
        lErro = TRPOcorrencias_Le(objOcorrencia)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 190254
        
        objTRPLiberaOCRSel.colOcorrenciais.Add objOcorrencia
     
        'Busca o próximo registro
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190255
 
    Loop
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
 
    TRPOcorrencias_Le_Bloqueios = SUCESSO
 
    Exit Function
 
Erro_TRPOcorrencias_Le_Bloqueios:
 
    TRPOcorrencias_Le_Bloqueios = gErr
 
    Select Case gErr
 
        Case 190244
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 190245 To 190253, 190255
            Call Rotina_Erro(vbOKOnly, "ERRO_PREPARACAO_COMANDO_SQL", gErr, sSQL)
 
        Case 190254
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190255)
 
    End Select
 
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
 
    Exit Function
 
End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA

Public Function TRPAportes_Le(ByVal objTRPAportes As ClassTRPAportes, Optional ByVal bLePagtos As Boolean = True) As Long

Dim lErro As Long
Dim alComando(0 To 5) As Long
Dim iIndice As Integer
Dim iProxCmd As Integer

On Error GoTo Erro_TRPAportes_Le

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 190376
    Next
    
    lErro = CF("TRPAportes_Le_Aux", alComando, objTRPAportes, iProxCmd, False, bLePagtos)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 190377
    
    If lErro = ERRO_LEITURA_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TRPAportes_Le = SUCESSO

    Exit Function

Erro_TRPAportes_Le:

    TRPAportes_Le = gErr

    Select Case gErr

        Case 190376
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 190377
        
        Case ERRO_LEITURA_SEM_DADOS
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190378)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function TRPAportes_Le_Aux(alComando() As Long, ByVal objTRPAportes As ClassTRPAportes, iProxCmd As Integer, Optional ByVal bLock As Boolean = False, Optional ByVal bLePagtos As Boolean = True) As Long

Dim lErro As Long
Dim tTRPAportes As typeTRPAportes
Dim lChave As Long
Dim sCampoChave As String

On Error GoTo Erro_TRPAportes_Le_Aux

    'Alocação de espaço no buffer
    tTRPAportes.sObservacao = String(STRING_TRP_APORTE_OBS, 0)
    tTRPAportes.sHistorico = String(STRING_TRP_APORTE_HISTORICO, 0)
    
    If objTRPAportes.lNumIntDoc <> 0 Then
        lChave = objTRPAportes.lNumIntDoc
        sCampoChave = "NumIntDoc"
    Else
        lChave = objTRPAportes.lCodigo
        sCampoChave = "Codigo"
    End If

    'Le a tabelaTRPAportes
    If bLock Then

        lErro = Comando_ExecutarPos(alComando(iProxCmd), "SELECT NumIntDoc, Codigo, Cliente, FilialCliente, DataEmissao, Observacao, " & _
                    "Historico, Tipo, Moeda, PrevValor, PrevDataDe, PrevDataAte, ProxParcela FROM TRPAportes WHERE " & sCampoChave & " = ? ", 0, _
                    tTRPAportes.lNumIntDoc, tTRPAportes.lCodigo, tTRPAportes.lCliente, tTRPAportes.iFilialCliente, _
                    tTRPAportes.dtDataEmissao, tTRPAportes.sObservacao, tTRPAportes.sHistorico, tTRPAportes.iTipo, tTRPAportes.iMoeda, _
                    tTRPAportes.dPrevValor, tTRPAportes.dtPrevDataDe, tTRPAportes.dtPrevDataAte, tTRPAportes.iProxParcela, _
                    lChave)
    Else
        lErro = Comando_Executar(alComando(iProxCmd), "SELECT NumIntDoc, Codigo, Cliente, FilialCliente, DataEmissao, Observacao, " & _
                    "Historico, Tipo, Moeda, PrevValor, PrevDataDe, PrevDataAte, ProxParcela FROM TRPAportes WHERE " & sCampoChave & " = ? ", _
                    tTRPAportes.lNumIntDoc, tTRPAportes.lCodigo, tTRPAportes.lCliente, tTRPAportes.iFilialCliente, _
                    tTRPAportes.dtDataEmissao, tTRPAportes.sObservacao, tTRPAportes.sHistorico, tTRPAportes.iTipo, tTRPAportes.iMoeda, _
                    tTRPAportes.dPrevValor, tTRPAportes.dtPrevDataDe, tTRPAportes.dtPrevDataAte, tTRPAportes.iProxParcela, _
                    lChave)
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 190327

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(iProxCmd))
    iProxCmd = iProxCmd + 1
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190328

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS
    
    If bLock Then
        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(iProxCmd - 1))
        If lErro <> SUCESSO Then gError 190374
    End If

    objTRPAportes.lNumIntDoc = tTRPAportes.lNumIntDoc
    objTRPAportes.lCodigo = tTRPAportes.lCodigo
    objTRPAportes.lCliente = tTRPAportes.lCliente
    objTRPAportes.iFilialCliente = tTRPAportes.iFilialCliente
    objTRPAportes.dtDataEmissao = tTRPAportes.dtDataEmissao
    objTRPAportes.sObservacao = tTRPAportes.sObservacao
    objTRPAportes.sHistorico = tTRPAportes.sHistorico
    objTRPAportes.iTipo = tTRPAportes.iTipo
    objTRPAportes.iMoeda = tTRPAportes.iMoeda
    objTRPAportes.dPrevValor = tTRPAportes.dPrevValor
    objTRPAportes.dtPrevDataDe = tTRPAportes.dtPrevDataDe
    objTRPAportes.dtPrevDataAte = tTRPAportes.dtPrevDataAte
    objTRPAportes.iProxParcela = tTRPAportes.iProxParcela
    
    If bLePagtos Then
    
        lErro = TRPAportePagtoCond_Le(alComando(iProxCmd), objTRPAportes, bLock)
        iProxCmd = iProxCmd + 1
        If lErro <> SUCESSO Then gError 190329
    
        lErro = TRPAportePagtoDireto_Le(alComando(iProxCmd), objTRPAportes, bLock)
        iProxCmd = iProxCmd + 1
        If lErro <> SUCESSO Then gError 190330
    
        lErro = TRPAportePagtoFat_Le(alComando(iProxCmd), objTRPAportes, bLock)
        iProxCmd = iProxCmd + 1
        If lErro <> SUCESSO Then gError 190331
        
        lErro = TRPAportePagtoFatCond_Le(alComando(iProxCmd), objTRPAportes, bLock)
        iProxCmd = iProxCmd + 1
        If lErro <> SUCESSO Then gError 190331

    End If

    TRPAportes_Le_Aux = SUCESSO

    Exit Function

Erro_TRPAportes_Le_Aux:

    TRPAportes_Le_Aux = gErr

    Select Case gErr

        Case 190326
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 190327, 190328
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPAPORTES", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora
        
        Case 190329 To 190321
        
        Case 190374
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRPAPORTES", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190332)

    End Select

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Private Function TRPAportePagtoCond_Le(ByVal lComando As Long, ByVal objTRPAporte As ClassTRPAportes, Optional ByVal bLock As Boolean = False) As Long

Dim lErro As Long
Dim tTRPAportePagtoCond As typeTRPAportePagtoCond
Dim objTRPAportePagtoCond As ClassTRPAportePagtoCond

On Error GoTo Erro_TRPAportePagtoCond_Le

    'Le a tabela TRPAportePagtoCond
    If bLock Then
        lErro = Comando_ExecutarPos(lComando, "SELECT NumIntDoc, NumIntDocAporte, Base, Status, DataPagto, NumIntDocDestino, " & _
                    "FormaPagto, TipoDocDestino, Percentual, Valor FROM TRPAportePagtoCond WHERE NumIntDocAporte = ? ORDER BY Seq ", 0, _
                    tTRPAportePagtoCond.lNumIntDoc, tTRPAportePagtoCond.lNumIntDocAporte, tTRPAportePagtoCond.iBase, tTRPAportePagtoCond.iStatus, _
                    tTRPAportePagtoCond.dtDataPagto, tTRPAportePagtoCond.lNumIntDocDestino, tTRPAportePagtoCond.iFormaPagto, tTRPAportePagtoCond.iTipoDocDestino, tTRPAportePagtoCond.dPercentual, tTRPAportePagtoCond.dValor, _
                    objTRPAporte.lNumIntDoc)
    Else
        lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntDocAporte, Base, Status, DataPagto, NumIntDocDestino, " & _
                    "FormaPagto, TipoDocDestino, Percentual, Valor FROM TRPAportePagtoCond WHERE NumIntDocAporte = ? ORDER BY Seq ", _
                    tTRPAportePagtoCond.lNumIntDoc, tTRPAportePagtoCond.lNumIntDocAporte, tTRPAportePagtoCond.iBase, tTRPAportePagtoCond.iStatus, _
                    tTRPAportePagtoCond.dtDataPagto, tTRPAportePagtoCond.lNumIntDocDestino, tTRPAportePagtoCond.iFormaPagto, tTRPAportePagtoCond.iTipoDocDestino, tTRPAportePagtoCond.dPercentual, tTRPAportePagtoCond.dValor, _
                    objTRPAporte.lNumIntDoc)
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 190334

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190335

    Do While lErro <> AD_SQL_SEM_DADOS
    
        If bLock Then
            'Trava registro contra alterações/Leituras
            lErro = Comando_LockExclusive(lComando)
            If lErro <> SUCESSO Then gError 190379
        End If
    
        Set objTRPAportePagtoCond = New ClassTRPAportePagtoCond

        objTRPAportePagtoCond.lNumIntDoc = tTRPAportePagtoCond.lNumIntDoc
        objTRPAportePagtoCond.lNumIntDocAporte = tTRPAportePagtoCond.lNumIntDocAporte
        objTRPAportePagtoCond.iBase = tTRPAportePagtoCond.iBase
        objTRPAportePagtoCond.iStatus = tTRPAportePagtoCond.iStatus
        objTRPAportePagtoCond.dtDataPagto = tTRPAportePagtoCond.dtDataPagto
        objTRPAportePagtoCond.lNumIntDocDestino = tTRPAportePagtoCond.lNumIntDocDestino
        objTRPAportePagtoCond.iFormaPagto = tTRPAportePagtoCond.iFormaPagto
        objTRPAportePagtoCond.iTipoDocDestino = tTRPAportePagtoCond.iTipoDocDestino
        objTRPAportePagtoCond.dPercentual = tTRPAportePagtoCond.dPercentual
        objTRPAportePagtoCond.dValor = tTRPAportePagtoCond.dValor
        
        objTRPAporte.colPagtoCondicionados.Add objTRPAportePagtoCond
        
        'Busca Próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190336
        
    Loop

    TRPAportePagtoCond_Le = SUCESSO

    Exit Function

Erro_TRPAportePagtoCond_Le:

    TRPAportePagtoCond_Le = gErr

    Select Case gErr
        
        Case 190334, 190335, 190336
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPAPORTEPAGTOCOND", gErr)
        
        Case 190379
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRPAPORTEPAGTOCOND", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190337)

    End Select

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Private Function TRPAportePagtoDireto_Le(ByVal lComando As Long, ByVal objTRPAporte As ClassTRPAportes, Optional ByVal bLock As Boolean = False) As Long

Dim lErro As Long
Dim tTRPAportePagtoDireto As typeTRPAportePagtoDireto
Dim objTRPAportePagtoDireto As ClassTRPAportePagtoDireto

On Error GoTo Erro_TRPAportePagtoDireto_Le

    'Le a tabelaTRPAportePagtoDireto
    If bLock Then
        lErro = Comando_ExecutarPos(lComando, "SELECT NumIntDoc, NumIntDocAporte, Valor, Vencimento, NumIntDocDestino, FormaPagto, " & _
                    "TipoDocDestino FROM TRPAportePagtoDireto WHERE NumIntDocAporte= ? ORDER BY Seq ", 0, _
                    tTRPAportePagtoDireto.lNumIntDoc, tTRPAportePagtoDireto.lNumIntDocAporte, tTRPAportePagtoDireto.dValor, tTRPAportePagtoDireto.dtVencimento, _
                    tTRPAportePagtoDireto.lNumIntDocDestino, tTRPAportePagtoDireto.iFormaPagto, tTRPAportePagtoDireto.iTipoDocDestino, _
                    objTRPAporte.lNumIntDoc)
    Else
        lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntDocAporte, Valor, Vencimento, NumIntDocDestino, FormaPagto, " & _
                    "TipoDocDestino FROM TRPAportePagtoDireto WHERE NumIntDocAporte= ? ORDER BY Seq ", _
                    tTRPAportePagtoDireto.lNumIntDoc, tTRPAportePagtoDireto.lNumIntDocAporte, tTRPAportePagtoDireto.dValor, tTRPAportePagtoDireto.dtVencimento, _
                    tTRPAportePagtoDireto.lNumIntDocDestino, tTRPAportePagtoDireto.iFormaPagto, tTRPAportePagtoDireto.iTipoDocDestino, _
                    objTRPAporte.lNumIntDoc)
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 190339

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190340

    Do While lErro <> AD_SQL_SEM_DADOS
    
        If bLock Then
            'Trava registro contra alterações/Leituras
            lErro = Comando_LockExclusive(lComando)
            If lErro <> SUCESSO Then gError 190380
        End If
    
        Set objTRPAportePagtoDireto = New ClassTRPAportePagtoDireto

        objTRPAportePagtoDireto.lNumIntDoc = tTRPAportePagtoDireto.lNumIntDoc
        objTRPAportePagtoDireto.lNumIntDocAporte = tTRPAportePagtoDireto.lNumIntDocAporte
        objTRPAportePagtoDireto.dValor = tTRPAportePagtoDireto.dValor
        objTRPAportePagtoDireto.dtVencimento = tTRPAportePagtoDireto.dtVencimento
        objTRPAportePagtoDireto.lNumIntDocDestino = tTRPAportePagtoDireto.lNumIntDocDestino
        objTRPAportePagtoDireto.iFormaPagto = tTRPAportePagtoDireto.iFormaPagto
        objTRPAportePagtoDireto.iTipoDocDestino = tTRPAportePagtoDireto.iTipoDocDestino

        objTRPAporte.colPagtoDireto.Add objTRPAportePagtoDireto
    
        'Busca Próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190341
    
    Loop

    TRPAportePagtoDireto_Le = SUCESSO

    Exit Function

Erro_TRPAportePagtoDireto_Le:

    TRPAportePagtoDireto_Le = gErr

    Select Case gErr

        Case 190339, 190340, 190341
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPAPORTEPAGTODIRETO", gErr)

        Case 190380
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRPAPORTEPAGTODIRETO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190342)

    End Select

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Private Function TRPAportePagtoFat_Le(ByVal lComando As Long, ByVal objTRPAporte As ClassTRPAportes, Optional ByVal bLock As Boolean = False) As Long

Dim lErro As Long
Dim tTRPAportePagtoFat As typeTRPAportePagtoFat
Dim objTRPAportePagtoFat As ClassTRPAportePagtoFat

On Error GoTo Erro_TRPAportePagtoFat_Le

    'Le a tabelaTRPAportePagtoFat
    If bLock Then
        lErro = Comando_ExecutarPos(lComando, "SELECT NumIntDoc, NumIntDocAporte, Valor, ValidadeDe, ValidadeAte, Saldo, " & _
                    "Percentual FROM TRPAportePagtoFat WHERE NumIntDocAporte= ? ORDER BY Seq ", 0, _
                    tTRPAportePagtoFat.lNumIntDoc, tTRPAportePagtoFat.lNumIntDocAporte, tTRPAportePagtoFat.dValor, tTRPAportePagtoFat.dtValidadeDe, _
                    tTRPAportePagtoFat.dtValidadeAte, tTRPAportePagtoFat.dSaldo, tTRPAportePagtoFat.dPercentual, _
                    objTRPAporte.lNumIntDoc)
    Else
        lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntDocAporte, Valor, ValidadeDe, ValidadeAte, Saldo, " & _
                    "Percentual FROM TRPAportePagtoFat WHERE NumIntDocAporte= ? ORDER BY Seq ", _
                    tTRPAportePagtoFat.lNumIntDoc, tTRPAportePagtoFat.lNumIntDocAporte, tTRPAportePagtoFat.dValor, tTRPAportePagtoFat.dtValidadeDe, _
                    tTRPAportePagtoFat.dtValidadeAte, tTRPAportePagtoFat.dSaldo, tTRPAportePagtoFat.dPercentual, _
                    objTRPAporte.lNumIntDoc)
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 190344

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190345

    Do While lErro <> AD_SQL_SEM_DADOS
    
        If bLock Then
            'Trava registro contra alterações/Leituras
            lErro = Comando_LockExclusive(lComando)
            If lErro <> SUCESSO Then gError 190381
        End If
    
        Set objTRPAportePagtoFat = New ClassTRPAportePagtoFat

        objTRPAportePagtoFat.lNumIntDoc = tTRPAportePagtoFat.lNumIntDoc
        objTRPAportePagtoFat.lNumIntDocAporte = tTRPAportePagtoFat.lNumIntDocAporte
        objTRPAportePagtoFat.dValor = tTRPAportePagtoFat.dValor
        objTRPAportePagtoFat.dtValidadeDe = tTRPAportePagtoFat.dtValidadeDe
        objTRPAportePagtoFat.dtValidadeAte = tTRPAportePagtoFat.dtValidadeAte
        objTRPAportePagtoFat.dSaldo = tTRPAportePagtoFat.dSaldo
        objTRPAportePagtoFat.dPercentual = tTRPAportePagtoFat.dPercentual
        
        objTRPAporte.colPagtoSobreFatura.Add objTRPAportePagtoFat

        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190346

    Loop

    TRPAportePagtoFat_Le = SUCESSO

    Exit Function

Erro_TRPAportePagtoFat_Le:

    TRPAportePagtoFat_Le = gErr

    Select Case gErr

        Case 190344, 190345, 190346
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPAPORTEPAGTOFAT", gErr)

        Case 190381
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRPAPORTEPAGTOFAT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190347)

    End Select

    Exit Function

End Function

Public Function Vouchers_Le_Periodo_Cliente(ByVal lCliente As Long, ByVal dtDataDe As Date, ByVal dtDataAte As Date, dValor As Double, dValorUSS As Double) As Long

Dim lErro As Long
Dim lComando As Long

On Error GoTo Erro_Vouchers_Le_Periodo_Cliente

    dValor = 0
    dValorUSS = 0

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 190709

    'Le a tabelaTRPOcorrencias
    lErro = Comando_Executar(lComando, "SELECT SUM(V.Valor), SUM(V.Valor / V.Cambio) FROM TRPVouchers AS V, FiliaisClientesTRP AS C WHERE C.CodCliente = V.Cliente AND C.EmpresaPai = ? AND V.Data BETWEEN ? AND ? AND  V.Status <> ? AND V.TipoDoc = ? AND C.Aporte = ? AND V.Cambio<>0 ", dValor, dValorUSS, lCliente, dtDataDe, dtDataAte, STATUS_CANCELADO, TRP_TIPODOC_VOU_TEXTO, MARCADO)
    If lErro <> AD_SQL_SUCESSO Then gError 190710

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190711

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS
    
    dValor = Arredonda_Moeda(dValor)
    dValorUSS = Arredonda_Moeda(dValorUSS)

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Vouchers_Le_Periodo_Cliente = SUCESSO

    Exit Function

Erro_Vouchers_Le_Periodo_Cliente:

    Vouchers_Le_Periodo_Cliente = gErr

    Select Case gErr

        Case 190709
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 190710, 190711
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPVOUCHERS", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190712)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function Vouchers_Le_Periodo_Cliente2(ByVal lCliente As Long, ByVal dPrevisaoUSS As Double, ByVal dtDataPrevDe As Date, ByVal dtDataPrevAte As Date, iCumpriuAMeta As Integer, dValorTotalRS As Double, dValorTotalUSS As Double, lNumVouAnt As Long, dtDataVouAnt As Date, dValorUSSAnt As Double, dValorRSAnt As Double, Optional lComandoP As Long = 0) As Long

Dim lErro As Long
Dim lComando As Long
Dim dValor As Double
Dim dCambio As Double
Dim dtDataVou As Date
Dim lNumVou As Long
Dim sTipVou As String
Dim sSerie As String

On Error GoTo Erro_Vouchers_Le_Periodo_Cliente2

    iCumpriuAMeta = DESMARCADO
    dValorTotalRS = 0
    dValorTotalUSS = 0
    lNumVouAnt = 0
    dtDataVouAnt = DATA_NULA
    dValorUSSAnt = 0
    dValorRSAnt = 0

    If lComandoP <> 0 Then
        lComando = lComandoP
    Else
        'Executa a abertura do Comando
        lComando = Comando_Abrir()
        If lComando = 0 Then gError 190709
    End If
    
    sTipVou = String(STRING_TRP_OCR_TIPOVOU, 0)
    sSerie = String(STRING_TRP_OCR_SERIE, 0)

    'Le a tabelaTRPOcorrencias
    lErro = Comando_Executar(lComando, "SELECT V.TipVou, V.Serie, V.NumVou, V.Valor, V.Cambio, V.Data FROM TRPVouchers AS V, FiliaisClientesTRP AS C WHERE V.Cliente = C.CodCliente AND C.EmpresaPai = ? AND V.Data BETWEEN ? AND ? AND  V.Status <> ? AND C.Aporte = ? AND V.Cambio <>0 ORDER BY V.Data, V.NumVou ", sTipVou, sSerie, lNumVou, dValor, dCambio, dtDataVou, lCliente, dtDataPrevDe, dtDataPrevAte, STATUS_CANCELADO, MARCADO)
    If lErro <> AD_SQL_SUCESSO Then gError 190710

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190711

    'Sem Dados
    Do While lErro <> AD_SQL_SEM_DADOS
    
        dValorTotalRS = dValorTotalRS + dValor
        dValorTotalUSS = dValorTotalUSS + Arredonda_Moeda(dValor / dCambio)
        
        If dValorTotalUSS > dPrevisaoUSS And iCumpriuAMeta = DESMARCADO Then
            iCumpriuAMeta = MARCADO
            lNumVouAnt = lNumVou
            dtDataVouAnt = dtDataVou
            dValorUSSAnt = dValorTotalUSS
            dValorRSAnt = dValorTotalRS
        End If

        'Busca Próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190711

    Loop

    If lComandoP = 0 Then
        'Fecha Comando
        Call Comando_Fechar(lComando)
    End If

    Vouchers_Le_Periodo_Cliente2 = SUCESSO

    Exit Function

Erro_Vouchers_Le_Periodo_Cliente2:

    Vouchers_Le_Periodo_Cliente2 = gErr

    Select Case gErr

        Case 190709
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 190710, 190711
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPVOUCHERS", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190712)

    End Select

    If lComandoP = 0 Then
        'Fecha Comando
        Call Comando_Fechar(lComando)
    End If

    Exit Function

End Function

Public Function TRPAportePagtoCond_Le_NumIntDoc(ByVal lComando As Long, ByVal objTRPAportePagto As ClassTRPAportePagtoCond, Optional ByVal bLock As Boolean = False) As Long

Dim lErro As Long
Dim tTRPAportePagtoCond As typeTRPAportePagtoCond

On Error GoTo Erro_TRPAportePagtoCond_Le_NumIntDoc

    'Le a tabela TRPAportePagtoCond
    If bLock Then
        lErro = Comando_ExecutarPos(lComando, "SELECT NumIntDoc, NumIntDocAporte, Base, Status, DataPagto, NumIntDocDestino, " & _
                    "FormaPagto, TipoDocDestino, Percentual, Valor FROM TRPAportePagtoCond WHERE NumIntDoc = ? ", 0, _
                    tTRPAportePagtoCond.lNumIntDoc, tTRPAportePagtoCond.lNumIntDocAporte, tTRPAportePagtoCond.iBase, tTRPAportePagtoCond.iStatus, _
                    tTRPAportePagtoCond.dtDataPagto, tTRPAportePagtoCond.lNumIntDocDestino, tTRPAportePagtoCond.iFormaPagto, tTRPAportePagtoCond.iTipoDocDestino, tTRPAportePagtoCond.dPercentual, tTRPAportePagtoCond.dValor, _
                    objTRPAportePagto.lNumIntDoc)
    Else
        lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntDocAporte, Base, Status, DataPagto, NumIntDocDestino, " & _
                    "FormaPagto, TipoDocDestino, Percentual, Valor FROM TRPAportePagtoCond WHERE NumIntDoc = ? ", _
                    tTRPAportePagtoCond.lNumIntDoc, tTRPAportePagtoCond.lNumIntDocAporte, tTRPAportePagtoCond.iBase, tTRPAportePagtoCond.iStatus, _
                    tTRPAportePagtoCond.dtDataPagto, tTRPAportePagtoCond.lNumIntDocDestino, tTRPAportePagtoCond.iFormaPagto, tTRPAportePagtoCond.iTipoDocDestino, tTRPAportePagtoCond.dPercentual, tTRPAportePagtoCond.dValor, _
                    objTRPAportePagto.lNumIntDoc)
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 190720

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190721

    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS
    
    If bLock Then
        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(lComando)
        If lErro <> SUCESSO Then gError 190722
    End If

    objTRPAportePagto.lNumIntDoc = tTRPAportePagtoCond.lNumIntDoc
    objTRPAportePagto.lNumIntDocAporte = tTRPAportePagtoCond.lNumIntDocAporte
    objTRPAportePagto.iBase = tTRPAportePagtoCond.iBase
    objTRPAportePagto.iStatus = tTRPAportePagtoCond.iStatus
    objTRPAportePagto.dtDataPagto = tTRPAportePagtoCond.dtDataPagto
    objTRPAportePagto.lNumIntDocDestino = tTRPAportePagtoCond.lNumIntDocDestino
    objTRPAportePagto.iFormaPagto = tTRPAportePagtoCond.iFormaPagto
    objTRPAportePagto.iTipoDocDestino = tTRPAportePagtoCond.iTipoDocDestino
    objTRPAportePagto.dPercentual = tTRPAportePagtoCond.dPercentual
    objTRPAportePagto.dValor = tTRPAportePagtoCond.dValor
    
    TRPAportePagtoCond_Le_NumIntDoc = SUCESSO

    Exit Function

Erro_TRPAportePagtoCond_Le_NumIntDoc:

    TRPAportePagtoCond_Le_NumIntDoc = gErr

    Select Case gErr
        
        Case 190720, 190721
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPAPORTEPAGTOCOND", gErr)
        
        Case 190722
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRPAPORTEPAGTOCOND", gErr)
        
        Case ERRO_LEITURA_SEM_DADOS
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190723)

    End Select

    Exit Function

End Function

Public Function TRPAportes_Le_Bloqueios(ByVal objTRPLiberaAporteSel As ClassTRPLiberaAporteSel) As Long
 
Dim lErro As Long
Dim iIndice As Integer
Dim sSQL As String
Dim sWhere As String
Dim vlNumIntDoc As Variant
Dim viStatus As Variant
Dim alComando(1 To 2) As Long
Dim objAportePagto As ClassTRPAportePagtoCond
Dim vlCodigoDe As Variant
Dim vlCodigoAte As Variant
Dim vdtPagtoDe As Variant
Dim vdtPagtoAte As Variant

On Error GoTo Erro_TRPAportes_Le_Bloqueios
    
    vlCodigoDe = objTRPLiberaAporteSel.lCodigoDe
    vlCodigoAte = objTRPLiberaAporteSel.lCodigoAte
    vdtPagtoDe = objTRPLiberaAporteSel.dtDataPagtoDe
    vdtPagtoAte = objTRPLiberaAporteSel.dtDataPagtoAte
    vlNumIntDoc = CLng(0)
    viStatus = STATUS_TRP_OCR_BLOQUEADO
 
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 190724
    Next
    
    sSQL = "SELECT P.NumIntDoc FROM TRPAportes AS A, TRPAportePagtoCond AS P  WHERE A.NumIntDoc = P.NumIntDocAporte AND Status = ? "
 
    If vlCodigoDe <> 0 Then
        sWhere = sWhere & "AND A.Codigo >= ? "
    End If
    If vlCodigoAte <> 0 Then
        sWhere = sWhere & "AND A.Codigo <= ? "
    End If

    If vdtPagtoDe <> DATA_NULA Then
        sWhere = sWhere & "AND P.DataPagto >= ? "
    End If
    If vdtPagtoAte <> DATA_NULA Then
        sWhere = sWhere & "AND P.DataPagto <= ? "
    End If
    
    sSQL = sSQL & sWhere
    
    sSQL = sSQL & "ORDER BY P.Seq "
 
    lErro = Comando_PrepararInt(alComando(1), sSQL)
    If (lErro <> AD_SQL_SUCESSO) Then gError 190725
 
    lErro = Comando_BindVarInt(alComando(1), vlNumIntDoc)
    If (lErro <> AD_SQL_SUCESSO) Then gError 190726
 
    lErro = Comando_BindVarInt(alComando(1), viStatus)
    If (lErro <> AD_SQL_SUCESSO) Then gError 190727
 
    If vlCodigoDe <> 0 Then
        lErro = Comando_BindVarInt(alComando(1), vlCodigoDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError 190728
    End If
    
    If vlCodigoAte <> 0 Then
        lErro = Comando_BindVarInt(alComando(1), vlCodigoAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError 190729
    End If
 
    If vdtPagtoDe <> DATA_NULA Then
        lErro = Comando_BindVarInt(alComando(1), vdtPagtoDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError 190730
    End If
    
    If vdtPagtoAte <> DATA_NULA Then
        lErro = Comando_BindVarInt(alComando(1), vdtPagtoAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError 190731
    End If
    
    lErro = Comando_ExecutarInt(alComando(1))
    If (lErro <> AD_SQL_SUCESSO) Then gError 190732
 
    'Busca o primeiro registro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190733
        
    'Para cada Produto
    Do While lErro = SUCESSO
    
        Set objAportePagto = New ClassTRPAportePagtoCond

        objAportePagto.lNumIntDoc = vlNumIntDoc
        
        'Lê o pagamento
        lErro = CF("TRPAportePagtoCond_Le_NumIntDoc", alComando(2), objAportePagto, False)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 190735
        
        objTRPLiberaAporteSel.colAportes.Add objAportePagto
     
        'Busca o próximo registro
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190734
 
    Loop
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
 
    TRPAportes_Le_Bloqueios = SUCESSO
 
    Exit Function
 
Erro_TRPAportes_Le_Bloqueios:
 
    TRPAportes_Le_Bloqueios = gErr
 
    Select Case gErr
 
        Case 190724
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 190725 To 190734
            Call Rotina_Erro(vbOKOnly, "ERRO_PREPARACAO_COMANDO_SQL", gErr, sSQL)
 
        Case 190735
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190736)
 
    End Select
 
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
 
    Exit Function
 
End Function

Public Function Carrega_Combo_FormaPagto(ByVal objCombo As Object) As Long

Dim lErro As Long

On Error GoTo Erro_Carrega_Combo_FormaPagto

    objCombo.Clear
    
    objCombo.AddItem FORMAPAGTO_TRP_OCR_FAT & SEPARADOR & FORMAPAGTO_TRP_OCR_FAT_TEXTO
    objCombo.ItemData(objCombo.NewIndex) = FORMAPAGTO_TRP_OCR_FAT

    objCombo.AddItem FORMAPAGTO_TRP_OCR_CRED & SEPARADOR & FORMAPAGTO_TRP_OCR_CRED_TEXTO
    objCombo.ItemData(objCombo.NewIndex) = FORMAPAGTO_TRP_OCR_CRED

    Carrega_Combo_FormaPagto = SUCESSO

    Exit Function

Erro_Carrega_Combo_FormaPagto:

    Carrega_Combo_FormaPagto = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190738)

    End Select

    Exit Function

End Function

Public Function Carrega_Combo_Base(ByVal objCombo As Object) As Long

Dim lErro As Long

On Error GoTo Erro_Carrega_Combo_Base

    objCombo.Clear

    objCombo.AddItem BASE_TRP_APORTE_REAL & SEPARADOR & BASE_TRP_APORTE_REAL_TEXTO
    objCombo.ItemData(objCombo.NewIndex) = BASE_TRP_APORTE_REAL

    objCombo.AddItem BASE_TRP_APORTE_REALACIMAPREV & SEPARADOR & BASE_TRP_APORTE_REALACIMAPREV_TEXTO
    objCombo.ItemData(objCombo.NewIndex) = BASE_TRP_APORTE_REALACIMAPREV

    Carrega_Combo_Base = SUCESSO

    Exit Function

Erro_Carrega_Combo_Base:

    Carrega_Combo_Base = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190737)

    End Select

    Exit Function

End Function

Public Function Carrega_Combo_TipoDoc(ByVal objCombo As Object) As Long

Dim lErro As Long

On Error GoTo Erro_Carrega_Combo_TipoDoc

    objCombo.Clear

    objCombo.AddItem TRP_TIPODOC_VOU_TEXTO
    objCombo.ItemData(objCombo.NewIndex) = TRP_TIPODOC_VOU

    objCombo.AddItem TRP_TIPODOC_OCR_TEXTO
    objCombo.ItemData(objCombo.NewIndex) = TRP_TIPODOC_OCR

    objCombo.AddItem TRP_TIPODOC_NVL_TEXTO
    objCombo.ItemData(objCombo.NewIndex) = TRP_TIPODOC_NVL

    objCombo.AddItem TRP_TIPODOC_CMR_TEXTO
    objCombo.ItemData(objCombo.NewIndex) = TRP_TIPODOC_CMR

    objCombo.AddItem TRP_TIPODOC_CMC_TEXTO
    objCombo.ItemData(objCombo.NewIndex) = TRP_TIPODOC_CMC
    
    objCombo.AddItem TRP_TIPODOC_OVER_TEXTO
    objCombo.ItemData(objCombo.NewIndex) = TRP_TIPODOC_OVER
    
    objCombo.AddItem TRP_TIPODOC_CMCC_TEXTO
    objCombo.ItemData(objCombo.NewIndex) = TRP_TIPODOC_CMCC
    
    Carrega_Combo_TipoDoc = SUCESSO

    Exit Function

Erro_Carrega_Combo_TipoDoc:

    Carrega_Combo_TipoDoc = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190828)

    End Select

    Exit Function

End Function

Public Function TRPFaturamento_Le_Dados(ByVal objFaturamento As ClassTRPFaturamento) As Long
'Le os itens a serem faturados com os filtros passados
 
Dim lErro As Long
Dim iIndice As Integer
Dim sSQL As String
Dim sWhere As String
Dim sWhereAux As String
Dim alComando(1 To 10) As Long
Dim vValor As Variant
Dim vValorAux(1 To 200) As Variant
Dim vlNumIntDoc As Variant
Dim vsTipoDoc As Variant
Dim vsSerie As Variant
Dim vlCliente As Variant
Dim viCondPagto As Variant
Dim vdValor As Variant
Dim vdtData As Variant
Dim vlNumVou As Variant
Dim vlNumero As Variant
Dim vlClienteAporte As Variant
Dim viAporte As Variant
Dim vsNome As Variant
Dim vsTipoVou As Variant
Dim objCliente As ClassCliente
Dim objFatCondPagto  As ClassTRPFatCondPagto
Dim objCondPagto  As ClassCondicaoPagto
Dim objInfoVoucher As ClassTRPFATInfoVoucher
Dim iProx As Integer
Dim objVoucher As ClassTRPVouchers

Dim vdCambio As Variant
Dim vdValorCambio As Variant
Dim viPax As Variant
Dim vsControle As Variant
Dim viMoeda As Variant
Dim vsProduto As Variant
Dim vdValorAporte As Variant
Dim viCartao As Variant
Dim vsTitular As Variant
Dim vsPassageiroNome As Variant
Dim vsPassageiroSobrenome As Variant
Dim vlEmissor As Variant
Dim vlClienteVou As Variant
Dim vlCliPassageiro As Variant
Dim viFilialEmpresa As Variant
Dim viFilialFat As Variant
Dim vsNomeCliVou As Variant
Dim vsNomeEmissor As Variant
Dim viFaturaJuntoOCR As Variant
Dim viUmVouPorFat As Variant
Dim viTipoCliente As Variant
Dim viTipoOCR As Variant
Dim vdValorVou As Variant
Dim vdValorComiVou As Variant

On Error GoTo Erro_TRPFaturamento_Le_Dados
 
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 190921
    Next
        
    vlNumIntDoc = CLng(0)
    vsTipoDoc = String(STRING_TRP_OCR_TIPODOC, 0)
    vsTipoVou = String(STRING_TRP_OCR_TIPOVOU, 0)
    vsSerie = String(STRING_TRP_OCR_SERIE, 0)
    vlCliente = CLng(0)
    vdValor = CDbl(0)
    vdtData = DATA_NULA
    vlNumVou = CLng(0)
    vlNumero = CLng(0)
    viCondPagto = CInt(0)
    vlClienteAporte = CLng(0)
    viAporte = CInt(0)
    vsNome = String(STRING_CLIENTE_NOME_REDUZIDO, 0)
    
    vdCambio = CDbl(0)
    vdValorCambio = CDbl(0)
    viPax = CInt(0)
    vsControle = String(STRING_TRP_VOU_CONTROLE, 0)
    viMoeda = CInt(0)
    vsProduto = String(STRING_PRODUTO, 0)
    vdValorAporte = CDbl(0)
    viCartao = CInt(0)
    vsTitular = String(STRING_TRP_VOU_TITULAR, 0)
    vsPassageiroNome = String(STRING_TRP_VOU_PAXNOME, 0)
    vsPassageiroSobrenome = String(STRING_TRP_VOU_PAXNOME, 0)
    vlEmissor = CLng(0)
    vlClienteVou = CLng(0)
    vlCliPassageiro = CLng(0)
    viFilialEmpresa = CInt(0)
    viFilialFat = CInt(0)
    vsNomeCliVou = String(STRING_CLIENTE_NOME_REDUZIDO, 0)
    vsNomeEmissor = String(STRING_FORNECEDOR_NOME_REDUZIDO, 0)
    viFaturaJuntoOCR = CInt(0)
    viUmVouPorFat = CInt(0)
    viTipoCliente = CInt(0)
    viTipoOCR = CInt(0)
    vdValorVou = CDbl(0)
    vdValorComiVou = CDbl(0)
      
    sSQL = "SELECT NumIntDoc, TipVou, TipoDoc, Serie , Cliente, Valor, Data, NumVou, Numero, CondPagto, EmpPai, Aporte, NomeCliente, " & _
    "Cambio, ValorCambio, Pax, Controle, Moeda, Produto, ValorAporte, Cartao, Titular, PassageiroNome, PassageiroSobrenome, Emissor, ClienteVou, CliPassageiro, FilialEmpresa, FilialFat, NomeClienteVou, NomeEmissor,FaturaJuntoOCR,UmVouPorFat, TipoCliente,TipoOCR,ValorVou, ValorVouComissao FROM DocsParaFat WHERE 1 = 1 "

    iProx = 1

    If objFaturamento.colFiliais.Count > 0 Then
        sWhere = sWhere & "AND ("
        sWhereAux = ""
        For Each vValor In objFaturamento.colFiliais
            If vValor <> 0 Then
                If Len(Trim(sWhereAux)) > 0 Then sWhereAux = sWhereAux & " AND "
                sWhereAux = sWhereAux & "FilialEmpresa <> ? "
            End If
        Next
        sWhere = sWhere & sWhereAux & ") "
    End If
    
    For Each objCliente In objFaturamento.colExcClientes
        sWhere = sWhere & "AND Cliente <> ? "
    Next
    
    For Each objVoucher In objFaturamento.colExcVouchers
        sWhere = sWhere & "AND NOT (TipVou = ? AND Serie = ? AND NumVou = ?) "
    Next
    
    sWhere = sWhere & "AND ("
    sWhereAux = ""
    For Each objFatCondPagto In objFaturamento.colCondPagtos
        If Len(Trim(sWhereAux)) > 0 Then sWhereAux = sWhereAux & " OR "
        sWhereAux = sWhereAux & " (CondPagto = ? "
        If objFatCondPagto.dtDataVouDe <> DATA_NULA Then
            sWhereAux = sWhereAux & "AND Data >= ? "
        End If
        If objFatCondPagto.dtDataVouAte <> DATA_NULA Then
            sWhereAux = sWhereAux & "AND Data <= ? "
        End If
        sWhereAux = sWhereAux & ") "
    Next
    sWhere = sWhere & sWhereAux & ") "
    
    If objFaturamento.colTiposDoc.Count > 0 Then
        sWhere = sWhere & "AND ("
        sWhereAux = ""
        For Each vValor In objFaturamento.colTiposDoc
            If Len(Trim(sWhereAux)) > 0 Then sWhereAux = sWhereAux & " AND "
            sWhereAux = sWhereAux & " TipoDoc <> ? "
        Next
        sWhere = sWhere & sWhereAux & ") "
    End If
    
    If objFaturamento.colTiposFat.Count > 0 Then
        sWhere = sWhere & "AND ("
        sWhereAux = ""
        For Each vValor In objFaturamento.colTiposFat
            If Len(Trim(sWhereAux)) > 0 Then sWhereAux = sWhereAux & " AND "
            sWhereAux = sWhereAux & " TipoFat <> ? "
        Next
        sWhere = sWhere & sWhereAux & ") "
    End If
    
    If objFaturamento.lCliente <> 0 Then
        sWhere = sWhere & "AND Cliente = ? "
    End If
    
    sSQL = sSQL & sWhere & " ORDER BY Data"
    
    lErro = Comando_PrepararInt(alComando(1), sSQL)
    If (lErro <> AD_SQL_SUCESSO) Then gError 190922
 
    lErro = Comando_BindVarInt(alComando(1), vlNumIntDoc)
    If (lErro <> AD_SQL_SUCESSO) Then gError 190923

    lErro = Comando_BindVarInt(alComando(1), vsTipoVou)
    If (lErro <> AD_SQL_SUCESSO) Then gError 190924

    lErro = Comando_BindVarInt(alComando(1), vsTipoDoc)
    If (lErro <> AD_SQL_SUCESSO) Then gError 190924
    
    lErro = Comando_BindVarInt(alComando(1), vsSerie)
    If (lErro <> AD_SQL_SUCESSO) Then gError 190925
    
    lErro = Comando_BindVarInt(alComando(1), vlCliente)
    If (lErro <> AD_SQL_SUCESSO) Then gError 190926
    
    lErro = Comando_BindVarInt(alComando(1), vdValor)
    If (lErro <> AD_SQL_SUCESSO) Then gError 190927
    
    lErro = Comando_BindVarInt(alComando(1), vdtData)
    If (lErro <> AD_SQL_SUCESSO) Then gError 190928
        
    lErro = Comando_BindVarInt(alComando(1), vlNumVou)
    If (lErro <> AD_SQL_SUCESSO) Then gError 190930
    
    lErro = Comando_BindVarInt(alComando(1), vlNumero)
    If (lErro <> AD_SQL_SUCESSO) Then gError 190931
    
    lErro = Comando_BindVarInt(alComando(1), viCondPagto)
    If (lErro <> AD_SQL_SUCESSO) Then gError 190933
    
    lErro = Comando_BindVarInt(alComando(1), vlClienteAporte)
    If (lErro <> AD_SQL_SUCESSO) Then gError 190973
    
    lErro = Comando_BindVarInt(alComando(1), viAporte)
    If (lErro <> AD_SQL_SUCESSO) Then gError 190974
    
    lErro = Comando_BindVarInt(alComando(1), vsNome)
    If (lErro <> AD_SQL_SUCESSO) Then gError 190975
    
    lErro = Comando_BindVarInt(alComando(1), vdCambio)
    If (lErro <> AD_SQL_SUCESSO) Then gError 190975
    
    lErro = Comando_BindVarInt(alComando(1), vdValorCambio)
    If (lErro <> AD_SQL_SUCESSO) Then gError 190975
    
    lErro = Comando_BindVarInt(alComando(1), viPax)
    If (lErro <> AD_SQL_SUCESSO) Then gError 190975
    
    lErro = Comando_BindVarInt(alComando(1), vsControle)
    If (lErro <> AD_SQL_SUCESSO) Then gError 190975
    
    lErro = Comando_BindVarInt(alComando(1), viMoeda)
    If (lErro <> AD_SQL_SUCESSO) Then gError 190975
    
    lErro = Comando_BindVarInt(alComando(1), vsProduto)
    If (lErro <> AD_SQL_SUCESSO) Then gError 190975
    
    lErro = Comando_BindVarInt(alComando(1), vdValorAporte)
    If (lErro <> AD_SQL_SUCESSO) Then gError 190975
    
    lErro = Comando_BindVarInt(alComando(1), viCartao)
    If (lErro <> AD_SQL_SUCESSO) Then gError 190975
    
    lErro = Comando_BindVarInt(alComando(1), vsTitular)
    If (lErro <> AD_SQL_SUCESSO) Then gError 190975
    
    lErro = Comando_BindVarInt(alComando(1), vsPassageiroNome)
    If (lErro <> AD_SQL_SUCESSO) Then gError 190975
    
    lErro = Comando_BindVarInt(alComando(1), vsPassageiroSobrenome)
    If (lErro <> AD_SQL_SUCESSO) Then gError 190975
    
    lErro = Comando_BindVarInt(alComando(1), vlEmissor)
    If (lErro <> AD_SQL_SUCESSO) Then gError 190975
    
    lErro = Comando_BindVarInt(alComando(1), vlClienteVou)
    If (lErro <> AD_SQL_SUCESSO) Then gError 190975
    
    lErro = Comando_BindVarInt(alComando(1), vlCliPassageiro)
    If (lErro <> AD_SQL_SUCESSO) Then gError 190975
    
    lErro = Comando_BindVarInt(alComando(1), viFilialEmpresa)
    If (lErro <> AD_SQL_SUCESSO) Then gError 190975
    
    lErro = Comando_BindVarInt(alComando(1), viFilialFat)
    If (lErro <> AD_SQL_SUCESSO) Then gError 190975

    lErro = Comando_BindVarInt(alComando(1), vsNomeCliVou)
    If (lErro <> AD_SQL_SUCESSO) Then gError 190975
    
    lErro = Comando_BindVarInt(alComando(1), vsNomeEmissor)
    If (lErro <> AD_SQL_SUCESSO) Then gError 190975
    
    lErro = Comando_BindVarInt(alComando(1), viFaturaJuntoOCR)
    If (lErro <> AD_SQL_SUCESSO) Then gError 190975
    
    lErro = Comando_BindVarInt(alComando(1), viUmVouPorFat)
    If (lErro <> AD_SQL_SUCESSO) Then gError 190975
    
    lErro = Comando_BindVarInt(alComando(1), viTipoCliente)
    If (lErro <> AD_SQL_SUCESSO) Then gError 190975
    
    lErro = Comando_BindVarInt(alComando(1), viTipoOCR)
    If (lErro <> AD_SQL_SUCESSO) Then gError 190975
    
    lErro = Comando_BindVarInt(alComando(1), vdValorVou)
    If (lErro <> AD_SQL_SUCESSO) Then gError 190975
    
    lErro = Comando_BindVarInt(alComando(1), vdValorComiVou)
    If (lErro <> AD_SQL_SUCESSO) Then gError 190975
    
    For Each vValor In objFaturamento.colFiliais
        If vValor <> 0 Then
        
            iProx = iProx + 1
            vValorAux(iProx) = vValor
        
            lErro = Comando_BindVarInt(alComando(1), vValorAux(iProx))
            If (lErro <> AD_SQL_SUCESSO) Then gError 190934
            
        End If
    Next
    
    For Each objCliente In objFaturamento.colExcClientes
    
        iProx = iProx + 1
        vValorAux(iProx) = objCliente.lCodigo
        
        lErro = Comando_BindVarInt(alComando(1), vValorAux(iProx))
        If (lErro <> AD_SQL_SUCESSO) Then gError 190935
        
    Next
    
    For Each objVoucher In objFaturamento.colExcVouchers
    
        iProx = iProx + 1
        vValorAux(iProx) = objVoucher.sTipVou
        
        lErro = Comando_BindVarInt(alComando(1), vValorAux(iProx))
        If (lErro <> AD_SQL_SUCESSO) Then gError 190935
        
        iProx = iProx + 1
        vValorAux(iProx) = objVoucher.sSerie
        
        lErro = Comando_BindVarInt(alComando(1), vValorAux(iProx))
        If (lErro <> AD_SQL_SUCESSO) Then gError 190935
        
        iProx = iProx + 1
        vValorAux(iProx) = objVoucher.lNumVou
        
        lErro = Comando_BindVarInt(alComando(1), vValorAux(iProx))
        If (lErro <> AD_SQL_SUCESSO) Then gError 190935
        
    Next
    
    For Each objFatCondPagto In objFaturamento.colCondPagtos
    
        iProx = iProx + 1
        vValorAux(iProx) = objFatCondPagto.objCondPagtos.iCodigo
        
        lErro = Comando_BindVarInt(alComando(1), vValorAux(iProx))
        If (lErro <> AD_SQL_SUCESSO) Then gError 190936
        
        If objFatCondPagto.dtDataVouDe <> DATA_NULA Then
        
            iProx = iProx + 1
            vValorAux(iProx) = objFatCondPagto.dtDataVouDe
            
            lErro = Comando_BindVarInt(alComando(1), vValorAux(iProx))
            If (lErro <> AD_SQL_SUCESSO) Then gError 190937

        End If
        If objFatCondPagto.dtDataVouAte <> DATA_NULA Then
            
            iProx = iProx + 1
            vValorAux(iProx) = objFatCondPagto.dtDataVouAte
             
            lErro = Comando_BindVarInt(alComando(1), vValorAux(iProx))
            If (lErro <> AD_SQL_SUCESSO) Then gError 190938
        
        End If
    Next
    
    For Each vValor In objFaturamento.colTiposDoc
    
        iProx = iProx + 1
        vValorAux(iProx) = vValor
    
        lErro = Comando_BindVarInt(alComando(1), vValorAux(iProx))
        If (lErro <> AD_SQL_SUCESSO) Then gError 190939
        
    Next
    
    For Each vValor In objFaturamento.colTiposFat
    
        iProx = iProx + 1
        vValorAux(iProx) = vValor
    
        lErro = Comando_BindVarInt(alComando(1), vValorAux(iProx))
        If (lErro <> AD_SQL_SUCESSO) Then gError 194208
        
    Next
    
    If objFaturamento.lCliente <> 0 Then
            
        iProx = iProx + 1
        vValorAux(iProx) = objFaturamento.lCliente
        
        lErro = Comando_BindVarInt(alComando(1), vValorAux(iProx))
        If (lErro <> AD_SQL_SUCESSO) Then gError 190940
    
    End If
    
    lErro = Comando_ExecutarInt(alComando(1))
    If (lErro <> AD_SQL_SUCESSO) Then gError 190941
 
    'Busca o primeiro registro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190942
        
    'Para cada Produto
    Do While lErro = SUCESSO
    
        Set objInfoVoucher = New ClassTRPFATInfoVoucher
    
        objInfoVoucher.dtDataEmissao = vdtData
        objInfoVoucher.dValor = vdValor
        objInfoVoucher.iCondPagto = viCondPagto
        objInfoVoucher.iFilialEmpresa = viFilialEmpresa
        objInfoVoucher.lCliente = vlCliente
        objInfoVoucher.lNumVou = vlNumVou
        objInfoVoucher.sSerie = vsSerie
        objInfoVoucher.sTipoDoc = vsTipoDoc
        objInfoVoucher.sTipoVou = vsTipoVou
        objInfoVoucher.lNumIntDoc = vlNumIntDoc
        objInfoVoucher.lNumero = vlNumero
        objInfoVoucher.iMarcado = MARCADO
        objInfoVoucher.lClienteAporte = vlClienteAporte
        objInfoVoucher.iaporte = viAporte
        objInfoVoucher.sNome = vsNome
        
        objInfoVoucher.dCambio = vdCambio
        objInfoVoucher.dValorCambio = vdValorCambio
        objInfoVoucher.iPax = viPax
        objInfoVoucher.sControle = vsControle
        objInfoVoucher.iMoeda = viMoeda
        objInfoVoucher.sProduto = vsProduto
        objInfoVoucher.dValorAporteVou = vdValorAporte
        objInfoVoucher.iCartao = viCartao
        objInfoVoucher.sTitularVou = vsTitular
        objInfoVoucher.sTitular = vsTitular
        objInfoVoucher.sPassageiroNome = vsPassageiroNome
        objInfoVoucher.sPassageiroSobreNome = vsPassageiroSobrenome
        objInfoVoucher.lEmissor = vlEmissor
        objInfoVoucher.lCliVou = vlClienteVou
        objInfoVoucher.lCliPassageiro = vlCliPassageiro
        objInfoVoucher.iFilialEmpresa = viFilialEmpresa
        objInfoVoucher.iFilialFat = viFilialFat
        objInfoVoucher.sNomeCliVou = vsNomeCliVou
        objInfoVoucher.sEmissor = vsNomeEmissor
        objInfoVoucher.iFaturaJuntoOcr = viFaturaJuntoOCR
        objInfoVoucher.iUmVouPorFat = viUmVouPorFat
        objInfoVoucher.iTipoCliente = viTipoCliente
        
        objInfoVoucher.iTipoOCR = viTipoOCR
        objInfoVoucher.dValorVou = vdValorVou
                
        If objInfoVoucher.sTipoDoc = TRP_TIPODOC_VOU_TEXTO Then objInfoVoucher.dValorComissao = vdValorComiVou
                        
        objFaturamento.colInfoVouchers.Add objInfoVoucher
    
        'Busca o próximo registro
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190943
    
    Loop
    
    'Le as informações dos aportes
    lErro = TRPFaturamento_Le_Dados2(objFaturamento, alComando(4), alComando(5), alComando(6), alComando(7), alComando(8))
    If lErro <> SUCESSO Then gError 190948
    
    'Calcula as informações das Faturas
    lErro = TRPFaturamento_Calcula_Dados1(objFaturamento, alComando(9), alComando(10))
    If lErro <> SUCESSO Then gError 190944
    
    'Calcula as informações das filiais
    lErro = TRPFaturamento_Calcula_Dados2(objFaturamento, alComando(2))
    If lErro <> SUCESSO Then gError 190946
    
    'Calcula as informações dos clientes
    lErro = TRPFaturamento_Calcula_Dados3(objFaturamento, alComando(3))
    If lErro <> SUCESSO Then gError 190947
    
    'Acerta o obj conforme as marcações (redistribui o aporte, acerta o valor das faturas, número de itens selecionados, etc ...)
    lErro = CF("TRPFaturamento_Atualiza", objFaturamento)
    If lErro <> SUCESSO Then gError 190949
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
 
    TRPFaturamento_Le_Dados = SUCESSO
 
    Exit Function
 
Erro_TRPFaturamento_Le_Dados:
 
    TRPFaturamento_Le_Dados = gErr
 
    Select Case gErr
 
        Case 190920
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 190921 To 190943, 190973, 190974, 190975, 194208
            Call Rotina_Erro(vbOKOnly, "ERRO_PREPARACAO_COMANDO_SQL", gErr, sSQL)
            
        Case 190944, 190946, 190947, 190948, 190949
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190945)
 
    End Select
 
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
 
    Exit Function
 
End Function

Private Function TRPFaturamento_Calcula_Dados1(ByVal objFaturamento As ClassTRPFaturamento, ByVal lComando1 As Long, ByVal lComando2 As Long) As Long
'Com base nos itens a serem faturados calcula as demais informações

Dim lErro As Long
Dim iIndice As Integer
Dim objInfoVoucher As ClassTRPFATInfoVoucher
Dim objInfoVoucherAux As ClassTRPFATInfoVoucher
Dim colCampos As New Collection
Dim colItensOrdenados As New Collection
Dim lClienteAnt As Long
Dim iCondPagtoAnt As Integer
Dim iFilialEmpresaAnt As Integer
Dim iTipoDocAnt As Integer
Dim colAuxiliar As New Collection
Dim bUltimo As Boolean
Dim objInfoFatura As ClassTRPFATInfoFatura
Dim lSeq As Long
Dim objFatCondPagto  As ClassTRPFatCondPagto
Dim dValor As Double
Dim dComissao As Double
Dim bFaturarIndividualmente As Boolean
Dim bTrataRegistro As Boolean
Dim iTipoDoc As Integer
Dim bOCRCliMesmaFat As Boolean
Dim objFilialCliente As ClassFilialCliente
Dim colCategoriaItem As New Collection
Dim objFilCliCat As ClassFilialCliCategoria
Dim objForn As ClassFornecedor
Dim objCliente As ClassCliente
Dim objVouAporte As ClassTRPVoucherAporte
Dim objOCRAporte As ClassTRPOcorrenciaAporte
Dim colVouAporte As Collection
Dim objOcr As ClassTRPOcorrencias
Dim objTipoOCR As ClassTRPTiposOcorrencia
Dim dFatorApOCR As Double

On Error GoTo Erro_TRPFaturamento_Calcula_Dados1

    bUltimo = False
    lSeq = 0

    If objFaturamento.colInfoVouchers.Count > 0 Then
    
        For Each objInfoVoucher In objFaturamento.colInfoVouchers
        
            If objInfoVoucher.iFilialFat <> 0 Then
                objInfoVoucher.iFilialEmpresa = objInfoVoucher.iFilialFat
            End If
                    
            Select Case objInfoVoucher.sTipoDoc

                Case TRP_TIPODOC_OVER_TEXTO
                    objInfoVoucher.iTipoDoc = TRP_TIPODOC_FAT_OVER
                    
                Case TRP_TIPODOC_OCR_TEXTO, TRP_TIPODOC_NVL_TEXTO
                    objInfoVoucher.iTipoDoc = TRP_TIPODOC_FAT_OCR
                
                Case Else
                    objInfoVoucher.iTipoDoc = TRP_TIPODOC_FAT_OUTROS
            
            End Select
            
            'O Valor do Cambio vem o do voucher que é o único que pode estar em dolar
            If objInfoVoucher.sTipoDoc <> TRP_TIPODOC_VOU_TEXTO Then
                objInfoVoucher.dValorCambio = objInfoVoucher.dValor
                objInfoVoucher.dCambio = 1
            End If
            
            'Os representantes e correntistas não precisam receber o controle preenchido pela agência
            If objInfoVoucher.sTipoDoc = TRP_TIPODOC_CMC_TEXTO Or objInfoVoucher.sTipoDoc = TRP_TIPODOC_CMR_TEXTO Then
                objInfoVoucher.sControle = ""
            End If
                
            'Pega o estorno do aporte (só para NVL do mesmo cliente do voucher, de clientes diferentes é estorno de comissão que não tem aporte)
            If (objInfoVoucher.sTipoDoc = TRP_TIPODOC_NVL_TEXTO Or objInfoVoucher.sTipoDoc = TRP_TIPODOC_OCR_TEXTO) And objInfoVoucher.lCliente = objInfoVoucher.lCliVou Then
                
                'Tem que ler o tipo para ver se deve ou não estornar
                Set objTipoOCR = New ClassTRPTiposOcorrencia
                
                objTipoOCR.iCodigo = objInfoVoucher.iTipoOCR
                
                'Le o tipo de ocorrência
                lErro = CF("TRPTiposOcorrencia_Le", objTipoOCR)
                If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 198230
                
                'Tipo de ocorrência não cadastrada
                If lErro <> SUCESSO Then gError 198231
                
                If objTipoOCR.iEstornaAporteVou = MARCADO And objInfoVoucher.dValor < 0 And objInfoVoucher.dValorVou > 0 Then
                
                    If objInfoVoucher.sTipoDoc = TRP_TIPODOC_NVL_TEXTO Then
                        dFatorApOCR = 1
                    Else
                        dFatorApOCR = Abs(objInfoVoucher.dValor / objInfoVoucher.dValorVou)
                    End If
                
                    'Tem que pegar de TRPVoucherAporte agora
                    objInfoVoucher.dValorAporte = -objInfoVoucher.dValorAporteVou * dFatorApOCR
                    
                    Set objVouAporte = New ClassTRPVoucherAporte
                    Set colVouAporte = New Collection
                
                    objVouAporte.sTipVou = objInfoVoucher.sTipoVou
                    objVouAporte.sSerie = objInfoVoucher.sSerie
                    objVouAporte.lNumVou = objInfoVoucher.lNumVou
                
                    'Le os aportes do vouchere inativado
                    lErro = TRPVoucherAporte_Le_Voucher(objVouAporte, colVouAporte)
                    If lErro <> SUCESSO Then gError 196466
                    
                    'Coloca os aportes na inativação com o valor com sinal invertido para o estorno
                    For Each objVouAporte In colVouAporte
                    
                        Set objOCRAporte = New ClassTRPOcorrenciaAporte
                        
                        objOCRAporte.dValorAporte = -objVouAporte.dValorAporte * dFatorApOCR
                        objOCRAporte.dValorAporteUSS = -objVouAporte.dValorAporteUSS * dFatorApOCR
                           
                        objOCRAporte.iParcelaAporte = objVouAporte.iParcelaAporte
                        objOCRAporte.iTipoPagtoAporte = objVouAporte.iTipoPagtoAporte
                        objOCRAporte.lCodigoAporte = objVouAporte.lCodigoAporte
                        objOCRAporte.lNumIntDocOCR = objInfoVoucher.lNumIntDoc
                        objOCRAporte.lNumIntDocPagtoAporte = objVouAporte.lNumIntDocPagtoAporte
                        objOCRAporte.iMoedaAporte = objVouAporte.iMoedaAporte
                
                        objInfoVoucher.colAportes.Add objOCRAporte
                    
                    Next
                    
                End If
            
            End If
            
            'Se for voucher pago com cartão o titular é o titular do cartão
            If objInfoVoucher.iCartao = MARCADO And objInfoVoucher.sTipoDoc = TRP_TIPODOC_VOU_TEXTO Then
                objInfoVoucher.sTitular = objInfoVoucher.sTitularVou
            ElseIf objInfoVoucher.sTipoDoc = TRP_TIPODOC_CMC_TEXTO Or objInfoVoucher.sTipoDoc = TRP_TIPODOC_CMR_TEXTO Or ((objInfoVoucher.sTipoDoc = TRP_TIPODOC_NVL_TEXTO Or objInfoVoucher.sTipoDoc = TRP_TIPODOC_OCR_TEXTO) And objInfoVoucher.lCliente <> objInfoVoucher.lCliVou) Then
            'Se for comissão deve sair o cliente do voucher como titular
                objInfoVoucher.sTitular = objInfoVoucher.sNomeCliVou
            Else
            'Senão é o passageiro
                objInfoVoucher.sTitular = objInfoVoucher.sPassageiroNome & " " & objInfoVoucher.sPassageiroSobreNome
            End If
            
            objInfoVoucher.dValorBruto = objInfoVoucher.dValor + objInfoVoucher.dValorComissao
            
        Next
        
        'Ordena os documentos de forma que seja possível dividir em faturas
        colCampos.Add "iFilialEmpresa"
        colCampos.Add "lCliente"
        colCampos.Add "iCondPagto"
        colCampos.Add "iTipoDoc"
        
        lErro = Ordena_Colecao(objFaturamento.colInfoVouchers, colItensOrdenados, colCampos)
        If lErro <> SUCESSO Then gError 190962
        
        For iIndice = 1 To colItensOrdenados.Count + 1
               
            If iIndice = colItensOrdenados.Count + 1 Then bUltimo = True
            
            If Not bUltimo Then
                Set objInfoVoucher = colItensOrdenados.Item(iIndice)
            End If
            
            If objInfoVoucher.iFaturaJuntoOcr = MARCADO Then
                bOCRCliMesmaFat = True
            Else
                bOCRCliMesmaFat = False
            End If
            
            If iIndice = 1 Then
                lClienteAnt = objInfoVoucher.lCliente
                iCondPagtoAnt = objInfoVoucher.iCondPagto
                iFilialEmpresaAnt = objInfoVoucher.iFilialEmpresa
                iTipoDocAnt = objInfoVoucher.iTipoDoc
            End If

            If iIndice = 1 Then
                bTrataRegistro = False
            Else
                bTrataRegistro = (lClienteAnt <> objInfoVoucher.lCliente Or iFilialEmpresaAnt <> objInfoVoucher.iFilialEmpresa Or iCondPagtoAnt <> objInfoVoucher.iCondPagto Or bUltimo Or (iTipoDocAnt <> objInfoVoucher.iTipoDoc And ((Not bOCRCliMesmaFat) Or objInfoVoucher.iTipoDoc = TRP_TIPODOC_FAT_OVER)))
            
                If objFaturamento.iFaturarIndividualmente = MARCADO Or objInfoVoucher.iUmVouPorFat = MARCADO Then bTrataRegistro = True
            End If
            
            If bTrataRegistro Then
                
                lClienteAnt = objInfoVoucher.lCliente
                iCondPagtoAnt = objInfoVoucher.iCondPagto
                iFilialEmpresaAnt = objInfoVoucher.iFilialEmpresa
                iTipoDoc = iTipoDocAnt
                iTipoDocAnt = objInfoVoucher.iTipoDoc
               
                lSeq = lSeq + 1
                
                Set objInfoFatura = New ClassTRPFATInfoFatura
                
                Set objInfoFatura.colInfoVouchers = colAuxiliar
                
                dValor = 0
                For Each objInfoVoucherAux In objInfoFatura.colInfoVouchers
                    objInfoVoucherAux.lFatura = lSeq
                    dValor = dValor + objInfoVoucherAux.dValor
                Next
                
                For Each objInfoVoucherAux In objInfoFatura.colInfoVouchers
                    Exit For
                Next
                
                For Each objFatCondPagto In objFaturamento.colCondPagtos
                    If objInfoVoucherAux.iCondPagto = objFatCondPagto.objCondPagtos.iCodigo Then
                        Exit For
                    End If
                Next
                
                objInfoFatura.dtDataEmissao = objFatCondPagto.dtDataEmissao
                objInfoFatura.dtDataVencimento = objFatCondPagto.dtDataVencimento
                objInfoFatura.dValor = dValor 'Voucher - Comissão (Depois tem que calcular e abater os aportes)
                objInfoFatura.iCondPagto = objInfoVoucherAux.iCondPagto
                objInfoFatura.iFilialEmpresa = objInfoVoucherAux.iFilialEmpresa
                objInfoFatura.iMarcado = MARCADO
                objInfoFatura.lCliente = objInfoVoucherAux.lCliente
                objInfoFatura.lFatura = lSeq
                objInfoFatura.iTipoDoc = iTipoDoc
                
                objFaturamento.colInfoFaturas.Add objInfoFatura
            
                Set colAuxiliar = New Collection
                
            End If
            
            colAuxiliar.Add objInfoVoucher
        
        Next
    
    End If

    TRPFaturamento_Calcula_Dados1 = SUCESSO

    Exit Function

Erro_TRPFaturamento_Calcula_Dados1:

    TRPFaturamento_Calcula_Dados1 = gErr

    Select Case gErr
    
        Case 190960 To 190962, 192725, 194201, 194210, 196466, 198230
        
        Case 198231 'Tipo de ocorrência não cadastrado
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_TRPTIPOSOCORRENCIA_NAO_PREENCHIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190963)

    End Select


    Exit Function

End Function

Private Function TRPFaturamento_Calcula_Dados2(ByVal objFaturamento As ClassTRPFaturamento, ByVal lComando As Long) As Long
'Com base nos itens a serem faturados calcula as demais informações

Dim lErro As Long
Dim objCodigoNome As New AdmCodigoNome
Dim colCodigoNome As New AdmColCodigoNome
Dim objInfoFilial As ClassTRPFATInfoFilial
Dim iQtdItens As Integer
Dim dValorItens As Double

On Error GoTo Erro_TRPFaturamento_Calcula_Dados2

    'Lê o Código e o Nome de Toda FilialEmpresa do BD
    lErro = CF("Cod_Nomes_Le_FilEmp", colCodigoNome)
    If lErro <> SUCESSO Then gError 190980
    
    For Each objCodigoNome In colCodigoNome
       
        Set objInfoFilial = New ClassTRPFATInfoFilial
        
        objInfoFilial.iFilialEmpresa = objCodigoNome.iCodigo
        objInfoFilial.sNome = objCodigoNome.sNome
        
        lErro = Comando_Executar(lComando, "SELECT COUNT(*), SUM(Valor) FROM DocsParaFatAux WHERE FilialEmpresa = ? ", iQtdItens, dValorItens, objInfoFilial.iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 190964
        
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190965
                    
        objInfoFilial.iNumItens = iQtdItens
        objInfoFilial.dValor = dValorItens
        
        objFaturamento.colInfoFiliais.Add objInfoFilial
    
    Next

    TRPFaturamento_Calcula_Dados2 = SUCESSO

    Exit Function

Erro_TRPFaturamento_Calcula_Dados2:

    TRPFaturamento_Calcula_Dados2 = gErr

    Select Case gErr
    
        Case 190964, 190965
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_DOCSPARAFAT", gErr)
            
        Case 190980

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190966)

    End Select

    Exit Function

End Function

Private Function TRPFaturamento_Calcula_Dados3(ByVal objFaturamento As ClassTRPFaturamento, ByVal lComando As Long) As Long
'Com base nos itens a serem faturados calcula as demais informações

Dim lErro As Long
Dim objInfoFatura As ClassTRPFATInfoFatura
Dim objInfoFaturaAux As ClassTRPFATInfoFatura
Dim bUltimo As Boolean
Dim colCampos As New Collection
Dim colItensOrdenados As New Collection
Dim lClienteAnt As Long
Dim objInfoCliente As ClassTRPFATInfoCliente
Dim colAuxiliar As New Collection
Dim iQtdItens As Integer
Dim dValorItens As Double
Dim iIndice As Integer
Dim iFilialEmpresaAnt As Integer

On Error GoTo Erro_TRPFaturamento_Calcula_Dados3

    bUltimo = False
    
    If objFaturamento.colInfoFaturas.Count > 0 Then
    
        'Já está com essa ordenação, não precisa odenar de novo
        Set colItensOrdenados = objFaturamento.colInfoFaturas
        
        For iIndice = 1 To colItensOrdenados.Count + 1
        
            If iIndice = colItensOrdenados.Count + 1 Then bUltimo = True
            
            If Not bUltimo Then
                Set objInfoFatura = colItensOrdenados.Item(iIndice)
            End If
            
            If iIndice = 1 Then
                lClienteAnt = objInfoFatura.lCliente
                iFilialEmpresaAnt = objInfoFatura.iFilialEmpresa
            End If
            
            If lClienteAnt <> objInfoFatura.lCliente Or bUltimo Or iFilialEmpresaAnt <> objInfoFatura.iFilialEmpresa Then
            
                lClienteAnt = objInfoFatura.lCliente
                iFilialEmpresaAnt = objInfoFatura.iFilialEmpresa
                
                Set objInfoCliente = New ClassTRPFATInfoCliente
                
                Set objInfoCliente.colInfoFaturas = colAuxiliar
                                       
                For Each objInfoFaturaAux In objInfoCliente.colInfoFaturas
                    Exit For
                Next
    
                lErro = Comando_Executar(lComando, "SELECT COUNT(*), SUM(Valor) FROM DocsParaFatAux WHERE Cliente = ? ", iQtdItens, dValorItens, objInfoFaturaAux.lCliente)
                If lErro <> AD_SQL_SUCESSO Then gError 190968
                
                lErro = Comando_BuscarPrimeiro(lComando)
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190969
    
                objInfoCliente.iNumItens = iQtdItens
                objInfoCliente.dValor = dValorItens
                objInfoCliente.iMarcado = MARCADO
                objInfoCliente.lCliente = objInfoFaturaAux.lCliente
                objInfoCliente.sNome = objInfoFaturaAux.colInfoVouchers.Item(1).sNome
                objInfoCliente.iFilialEmpresa = objInfoFaturaAux.iFilialEmpresa
                
                If objInfoFaturaAux.iTipoDoc = TRP_TIPODOC_FAT_OVER Then
                    objInfoCliente.iTipo = TRP_CLIENTEINFO_TIPO_FORNECEDOR
                Else
                    objInfoCliente.iTipo = TRP_CLIENTEINFO_TIPO_CLIENTE
                End If
                
                objFaturamento.colInfoClientes.Add objInfoCliente
            
                Set colAuxiliar = New Collection
                
            End If
            
            colAuxiliar.Add objInfoFatura
        
        Next
        
    End If
    
    TRPFaturamento_Calcula_Dados3 = SUCESSO

    Exit Function

Erro_TRPFaturamento_Calcula_Dados3:

    TRPFaturamento_Calcula_Dados3 = gErr

    Select Case gErr
    
        Case 190967
        
        Case 190968, 190969
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_DOCSPARAFAT", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190959)

    End Select


    Exit Function

End Function

Public Function TRPFaturamento_Atualiza(ByVal objFaturamento As ClassTRPFaturamento) As Long
'Com base nos itens a serem faturados calcula as demais informações

Dim lErro As Long
Dim objInfoFilial As ClassTRPFATInfoFilial
Dim objInfoVoucher As ClassTRPFATInfoVoucher
Dim objInfoFatura As ClassTRPFATInfoFatura
Dim objInfoCliente As ClassTRPFATInfoCliente
Dim iNumItens As Integer
Dim iNumFat As Integer
Dim dValorItens As Double
Dim dValorComissao As Double
Dim colItens As New Collection
Dim colItensOrdenados As New Collection
Dim colCampos As New Collection
Dim objFatAporte As ClassTRPFatAporte
Dim objFatAporteCred As ClassTRPFatAporteCred
Dim dValorAporte As Double
Dim iLinha As Integer
Dim dValorAporteAux As Double
Dim dValorCreditos As Double
Dim dValorDebitos As Double
Dim objVouAporte As ClassTRPVoucherAporte
Dim dValorAporteUSS As Double
Dim objFaturaAporte As ClassTRPFaturaAporte

On Error GoTo Erro_TRPFaturamento_Atualiza

    'Zera a utilização dos aportes
    For Each objInfoVoucher In objFaturamento.colInfoVouchers
        'Se é um voucher
        If objInfoVoucher.sTipoDoc = TRP_TIPODOC_VOU_TEXTO Then
            objInfoVoucher.dValorAporte = 0
            Set objInfoVoucher.colAportes = New Collection
        End If
        If objInfoVoucher.iMarcado = MARCADO Then
            colItens.Add objInfoVoucher
        End If
    Next
    For Each objFatAporte In objFaturamento.colAportes
        If objFatAporte.iTipoPagtoAporte = TRP_TIPO_APORTE_SOBREFATURA Then
            objFatAporte.dSaldoAtual = objFatAporte.objAportePagto.dSaldo
        End If
    Next
    For Each objFatAporteCred In objFaturamento.colAportesCred
        objFatAporteCred.dSaldoAtual = objFatAporteCred.dSaldoAnt
    Next
    For Each objInfoFatura In objFaturamento.colInfoFaturas
        Set objInfoFatura.colFatAporteCred = New Collection
        objInfoFatura.dValorAporteCred = 0
    Next

    'Ordena pela data para poder redistrubir o aporte
    colCampos.Add "dtDataEmissao"
    colCampos.Add "lNumVou"
    
    lErro = Ordena_Colecao(colItens, colItensOrdenados, colCampos)
    If lErro <> SUCESSO Then gError 190967
    
    For Each objInfoVoucher In colItensOrdenados
    
        'Se é um voucher
        If objInfoVoucher.sTipoDoc = TRP_TIPODOC_VOU_TEXTO Then

            For Each objFatAporte In objFaturamento.colAportes
            
                'Se é o mesmo cliente e o aporte ainda tem saldo disponível ou é por cumprimento de meta
                'e o voucher em questão é posteior ao cumprimento dessa meta
                If objFatAporte.lCliente = objInfoVoucher.lClienteAporte And objInfoVoucher.iaporte = MARCADO And (objFatAporte.dSaldoAtual > 0 Or (objFatAporte.iTipoPagtoAporte = TRP_TIPO_APORTE_SOBREFATURA_COND And (objInfoVoucher.dtDataEmissao > objFatAporte.dtDataVouAntMeta Or (objInfoVoucher.dtDataEmissao = objFatAporte.dtDataVouAntMeta And objInfoVoucher.lNumVou >= objFatAporte.lNumVouAntMeta)))) Then
                
                    'Se não passou da validade
                    If (objInfoVoucher.dtDataEmissao <= objFatAporte.objAportePagto.dtValidadeAte Or objFatAporte.objAportePagto.dtValidadeAte = DATA_NULA) Then
                    
                        'Se já começou a ser válido
                        If (objInfoVoucher.dtDataEmissao >= objFatAporte.objAportePagto.dtValidadeDe Or objFatAporte.objAportePagto.dtValidadeDe = DATA_NULA) Then
                        
                            dValorAporte = objInfoVoucher.dValor * objFatAporte.objAportePagto.dPercentual
                            
                            'Se é por cumprimento de meta
                            If objFatAporte.iTipoPagtoAporte = TRP_TIPO_APORTE_SOBREFATURA_COND Then
                                'Se é o primeiro voucher que deveria receber a meta
                                If objInfoVoucher.lNumVou = objFatAporte.lNumVouAntMeta Then
                                    If objInfoVoucher.dCambio = 0 Then gError 192862
                                    dValorAporte = (((objFatAporte.dRealizadoAntMetaUSS - objFatAporte.dPrevisaoUSS) * objInfoVoucher.dCambio)) * objFatAporte.objAportePagto.dPercentual
                                End If
                            End If
                            
                            If objFatAporte.iMoeda = MOEDA_DOLAR Then
                                If objInfoVoucher.dCambio = 0 Then gError 192862
                                If dValorAporte / objInfoVoucher.dCambio > objFatAporte.dSaldoAtual And objFatAporte.iTipoPagtoAporte = TRP_TIPO_APORTE_SOBREFATURA Then
                                    dValorAporte = objFatAporte.dSaldoAtual * objInfoVoucher.dCambio
                                End If
                            Else
                                If dValorAporte > objFatAporte.dSaldoAtual And objFatAporte.iTipoPagtoAporte = TRP_TIPO_APORTE_SOBREFATURA Then
                                    dValorAporte = objFatAporte.dSaldoAtual
                                End If
                            End If
                            
                            If objFatAporte.iMoeda = MOEDA_DOLAR Then
                                dValorAporteUSS = Arredonda_Moeda(dValorAporte / objInfoVoucher.dCambio, 4)
                            Else
                                dValorAporteUSS = 0
                            End If
                            dValorAporte = Arredonda_Moeda(dValorAporte, 4)
                        
                            objInfoVoucher.dValorAporte = objInfoVoucher.dValorAporte + dValorAporte
                            
                            Set objVouAporte = New ClassTRPVoucherAporte
                            
                            objVouAporte.lNumIntDocPagtoAporte = objFatAporte.objAportePagto.lNumIntDoc
                            objVouAporte.lCodigoAporte = objFatAporte.lCodigo
                            objVouAporte.dValorAporte = dValorAporte
                            objVouAporte.iTipoPagtoAporte = objFatAporte.iTipoPagtoAporte
                            objVouAporte.iMoedaAporte = objFatAporte.iMoeda
                            objVouAporte.dValorAporteUSS = dValorAporteUSS
                            
                            If Abs(objVouAporte.dValorAporte) > DELTA_VALORMONETARIO Then
                                objInfoVoucher.colAportes.Add objVouAporte
                            End If
                            
                            If objFatAporte.iMoeda = MOEDA_DOLAR Then
                                objFatAporte.dSaldoAtual = objFatAporte.dSaldoAtual - dValorAporteUSS
                            Else
                                objFatAporte.dSaldoAtual = objFatAporte.dSaldoAtual - dValorAporte
                            End If

                        End If
                    
                    End If
                
                End If
        
            Next
            
        End If
        
    Next
                
    iLinha = 0
    For Each objInfoFatura In objFaturamento.colInfoFaturas
    
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
            
                dValorItens = dValorItens + objInfoVoucher.dValorBruto - objInfoVoucher.dValorAporte - objInfoVoucher.dValorComissao
                dValorComissao = dValorComissao + objInfoVoucher.dValorComissao
                dValorAporte = dValorAporte + objInfoVoucher.dValorAporte
                
                'Se for negativo é um crédito
                If objInfoVoucher.dValor < 0 Then
                    dValorCreditos = dValorCreditos + objInfoVoucher.dValor
                End If
                'Se for um documento fora voucher que acrescenta no valor é um débito
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
        
        If (Abs(dValorItens) < DELTA_VALORMONETARIO) Or (objInfoFatura.iTipoDoc = TRP_TIPODOC_FAT_OVER And dValorItens > -DELTA_VALORMONETARIO) Then
            objInfoFatura.iMarcado = DESMARCADO
            For Each objInfoVoucher In objInfoFatura.colInfoVouchers
                objInfoVoucher.iMarcado = DESMARCADO
            Next
        End If
        
        objInfoFatura.dValor = dValorItens
        objInfoFatura.dValorAporte = dValorAporte
        objInfoFatura.dValorComissao = dValorComissao
        objInfoFatura.dValorCreditos = dValorCreditos
        objInfoFatura.dValorDebitos = dValorDebitos
    
    Next

    iLinha = 0
    For Each objInfoVoucher In objFaturamento.colInfoVouchers
        iLinha = iLinha + 1
        objInfoVoucher.iLinha = iLinha
    Next
    
    'distribui primeiro da empresa pai
    For Each objFatAporteCred In objFaturamento.colAportesCred
        'Distribui os créditos via aporte parcela única
        For Each objInfoFatura In objFaturamento.colInfoFaturas
            'Se o cliente é a empresa pai do aporte
            If objInfoFatura.lCliente = objFatAporteCred.lCliente And objFatAporteCred.dSaldoAtual > 0 And objInfoFatura.dtDataEmissao >= objFatAporteCred.dtDataVenc And objInfoFatura.dValor > 0 Then
            
                Set objFaturaAporte = New ClassTRPFaturaAporte
            
                'Dá o crédito, mas deixa a fatura com no mínimo 1 centavo
                If objInfoFatura.dValor >= objFatAporteCred.dSaldoAtual + 0.01 Then
                    objFaturaAporte.dValorAporte = objFatAporteCred.dSaldoAtual
                    objFaturaAporte.iPagtoParcial = DESMARCADO
                Else
                    objFaturaAporte.dValorAporte = objInfoFatura.dValor - 0.01
                    objFaturaAporte.iPagtoParcial = MARCADO
                End If
                
                'Acerta as informações de valores na fatura
                objInfoFatura.dValor = objInfoFatura.dValor - objFaturaAporte.dValorAporte
                objInfoFatura.dValorCreditos = objInfoFatura.dValorCreditos - objFaturaAporte.dValorAporte
                objInfoFatura.dValorAporteCred = objInfoFatura.dValorAporteCred + objFaturaAporte.dValorAporte
                
                'Acerta o saldo dos créditos sobre fatura
                objFatAporteCred.dSaldoAtual = objFatAporteCred.dSaldoAtual - objFaturaAporte.dValorAporte
                
                'Associa crédito com a fatura
                objFaturaAporte.lNumIntDocCredito = objFatAporteCred.lNumIntDocCredito
                objFaturaAporte.lCodigoAporte = objFatAporteCred.lCodigo
                objFaturaAporte.iNumParcelas = objFatAporteCred.iNumParcelas
                objFaturaAporte.iParcelaAporte = objFatAporteCred.iParcelaAporte
                Set objFaturaAporte.objCredito = objFatAporteCred.objCredito
            
                objInfoFatura.colFatAporteCred.Add objFaturaAporte
            End If
        Next
    Next
    'Se sobra distribui para as empresas filhas
    For Each objFatAporteCred In objFaturamento.colAportesCred
    
        If objFatAporteCred.dSaldoAtual > 0 Then
        
            'Distribui os créditos via aporte parcela única
            For Each objInfoFatura In objFaturamento.colInfoFaturas
            
                'Se o cliente é uma empresa filha
                If objInfoFatura.colInfoVouchers.Item(1).sTipoDoc = TRP_TIPODOC_VOU_TEXTO And objInfoFatura.colInfoVouchers.Item(1).lClienteAporte = objFatAporteCred.lCliente And (objInfoFatura.lCliente <> objFatAporteCred.lCliente) And objFatAporteCred.dSaldoAtual > 0 And objInfoFatura.dtDataEmissao >= objFatAporteCred.dtDataVenc Then
            
                    Set objFaturaAporte = New ClassTRPFaturaAporte
                
                    'Dá o crédito, mas deixa a fatura com no mínimo 1 centavo
                    If objInfoFatura.dValor >= objFatAporteCred.dSaldoAtual + 0.01 Then
                        objFaturaAporte.dValorAporte = objFatAporteCred.dSaldoAtual
                        objFaturaAporte.iPagtoParcial = DESMARCADO
                    Else
                        objFaturaAporte.dValorAporte = objInfoFatura.dValor - 0.01
                        objFaturaAporte.iPagtoParcial = MARCADO
                    End If
                    
                    'Acerta as informações de valores na fatura
                    objInfoFatura.dValor = objInfoFatura.dValor - objFaturaAporte.dValorAporte
                    objInfoFatura.dValorCreditos = objInfoFatura.dValorCreditos - objFaturaAporte.dValorAporte
                    objInfoFatura.dValorAporteCred = objInfoFatura.dValorAporteCred + objFaturaAporte.dValorAporte
                    
                    'Acerta o saldo dos créditos sobre fatura
                    objFatAporteCred.dSaldoAtual = objFatAporteCred.dSaldoAtual - objFaturaAporte.dValorAporte
                    
                    'Associa crédito com a fatura
                    objFaturaAporte.lNumIntDocCredito = objFatAporteCred.lNumIntDocCredito
                    objFaturaAporte.lCodigoAporte = objFatAporteCred.lCodigo
                    objFaturaAporte.iNumParcelas = objFatAporteCred.iNumParcelas
                    objFaturaAporte.iParcelaAporte = objFatAporteCred.iParcelaAporte
                    Set objFaturaAporte.objCredito = objFatAporteCred.objCredito
                
                    objInfoFatura.colFatAporteCred.Add objFaturaAporte
                End If
            Next
            
        End If
    Next
    
    'Calcula os valor selecionado por filial
    For Each objInfoFilial In objFaturamento.colInfoFiliais
    
        iNumItens = 0
        dValorItens = 0
        iNumFat = 0
        
        For Each objInfoFatura In objFaturamento.colInfoFaturas
        
            If objInfoFatura.iFilialEmpresa = objInfoFilial.iFilialEmpresa Then
            
                If objInfoFatura.iMarcado = MARCADO Then
                    iNumFat = iNumFat + 1
                    iNumItens = iNumItens + objInfoFatura.colInfoVouchers.Count
                    dValorItens = dValorItens + objInfoFatura.dValor
                End If
            
            End If
    
        Next
        
        objInfoFilial.dValorS = dValorItens
        objInfoFilial.iNumItensS = iNumItens
        objInfoFilial.iNumFatS = iNumFat
    
    Next
    
    'Calcula os valor selecionado por cliente
    iLinha = 0
    For Each objInfoCliente In objFaturamento.colInfoClientes
    
        iNumItens = 0
        dValorItens = 0
        iLinha = iLinha + 1
        
        objInfoCliente.iLinha = iLinha
        
        For Each objInfoFatura In objFaturamento.colInfoFaturas
        
            If objInfoFatura.lCliente = objInfoCliente.lCliente Then
            
                If objInfoFatura.iMarcado = MARCADO Then
                    iNumFat = iNumFat + 1
                    iNumItens = iNumItens + objInfoFatura.colInfoVouchers.Count
                    dValorItens = dValorItens + objInfoFatura.dValor
                End If
            
            End If
    
        Next
        
        objInfoCliente.dValorS = dValorItens
        objInfoCliente.iNumItensS = iNumItens
    
    Next
    
    TRPFaturamento_Atualiza = SUCESSO

    Exit Function

Erro_TRPFaturamento_Atualiza:

    TRPFaturamento_Atualiza = gErr

    Select Case gErr
    
        Case 192862
            Call Rotina_Erro(vbOKOnly, "ERRO_APORTE_DOLAR_VOUCHER_SEM_CAMBIO", gErr, objInfoVoucher.sTipoVou, objInfoVoucher.sSerie, objInfoVoucher.lNumVou)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190958)

    End Select

    Exit Function

End Function

Private Function TRPFaturamento_Le_Dados2(ByVal objFaturamento As ClassTRPFaturamento, ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lComando3 As Long, ByVal lComando4 As Long, ByVal lComando5 As Long) As Long
'Le os aportes sobre fatura

Dim lErro As Long
Dim objAportePagto As Object
Dim objAportePagtoCond As ClassTRPAportePagtoFatC
Dim objFatAporte As ClassTRPFatAporte
Dim lNumIntDoc As Long
Dim lCliente As Long
Dim iProxParcela As Integer
Dim lCodigo As Long
Dim iMoeda As Integer
Dim iTipoPagtoAporte As Integer
Dim dPrevValor As Double
Dim dtPrevDataDe As Date
Dim dtPrevDataAte As Date
Dim bConsiderar As Boolean
Dim iCumpriuAMeta As Integer
Dim dValorTotalRS As Double
Dim dValorTotalUSS As Double
Dim lNumVouAnt As Long
Dim dtDataVouAnt As Date
Dim dValorUSSAnt As Double
Dim dValorRSAnt As Double
Dim dValor As Double
Dim dSaldo As Double
Dim dtVencimento As Date
Dim lNumIntDocDestino As Long
Dim lNumIntDocAporte As Long
Dim objFatAporteCred As ClassTRPFatAporteCred
Dim objDebitoRecCli As ClassDebitoRecCli
Dim iNumParcelas As Integer
Dim iParcela As Integer

On Error GoTo Erro_TRPFaturamento_Le_Dados2
            
    'Le os investimento sobre fatura (os que são descontados via percentual voucher a voucher)
    lErro = Comando_Executar(lComando1, "SELECT TipoPagtoAporte, NumIntDocPagtoAporte, ProxParcela, Cliente, CodigoAporte, Moeda, PrevValor, PrevDataDe, PrevDataAte FROM TRPAportesPagtoParaFat ORDER BY Cliente, CodigoAporte", iTipoPagtoAporte, lNumIntDoc, iProxParcela, lCliente, lCodigo, iMoeda, dPrevValor, dtPrevDataDe, dtPrevDataAte)
    If lErro <> AD_SQL_SUCESSO Then gError 190953
    
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190954
    
    Do While lErro <> AD_SQL_SEM_DADOS
    
        bConsiderar = True
    
        Set objFatAporte = New ClassTRPFatAporte
    
        objFatAporte.lCliente = lCliente
        objFatAporte.iProxParcela = iProxParcela
        objFatAporte.lCodigo = lCodigo
        objFatAporte.iMoeda = iMoeda
        objFatAporte.iTipoPagtoAporte = iTipoPagtoAporte
        objFatAporte.dPrevisaoUSS = dPrevValor
        objFatAporte.dtPrevDataDe = dtPrevDataDe
        objFatAporte.dtPrevDataAte = dtPrevDataAte

        If iTipoPagtoAporte = TRP_TIPO_APORTE_SOBREFATURA Then
        
            Set objAportePagto = New ClassTRPAportePagtoFat
            
            objAportePagto.lNumIntDoc = lNumIntDoc
            
            lErro = TRPAportePagtoFat_Le_NumIntDoc(lComando2, objAportePagto)
        Else
        
            Set objAportePagto = New ClassTRPAportePagtoFatC
            
            objAportePagto.lNumIntDoc = lNumIntDoc
    
            lErro = TRPAportePagtoFatCond_Le_NumIntDoc(lComando3, objAportePagto)
        End If
        If lErro <> SUCESSO Then gError 190955
               
        Set objFatAporte.objAportePagto = objAportePagto
               
        If iTipoPagtoAporte = TRP_TIPO_APORTE_SOBREFATURA Then
            objFatAporte.dSaldoAtual = objAportePagto.dSaldo
        Else
        
            'Lê o realizado caso seja um aporte sobre fatura por cumprimento de meta
            lErro = Vouchers_Le_Periodo_Cliente2(lCliente, dPrevValor, dtPrevDataDe, dtPrevDataAte, iCumpriuAMeta, dValorTotalRS, dValorTotalUSS, lNumVouAnt, dtDataVouAnt, dValorUSSAnt, dValorRSAnt, lComando5)
            If lErro <> SUCESSO Then gError 190955
            
            If iCumpriuAMeta = DESMARCADO Then bConsiderar = False
            
            objFatAporte.dRealizadoAntMetaRS = dValorRSAnt
            objFatAporte.dRealizadoAntMetaUSS = dValorUSSAnt
            objFatAporte.dRealizadoRS = dValorTotalRS
            objFatAporte.dRealizadoUSS = dValorTotalUSS
            objFatAporte.dtDataVouAntMeta = dtDataVouAnt
            objFatAporte.lNumVouAntMeta = lNumVouAnt
        
        End If
    
        If bConsiderar Then objFaturamento.colAportes.Add objFatAporte
    
        lErro = Comando_BuscarProximo(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190956
    
    Loop

    'Le os investimento diretos via crédito (os pagtos direto por crédito abatem em parcela única no faturamento, salvo quando o valor da fatura não permitir o crédito de forma única)
    lErro = Comando_Executar(lComando4, "SELECT TipoPagtoAporte, NumIntDocPagto, NumIntDocAporte, Cliente, ValorTotal, Saldo, Vencimento, NumIntDocDestino, Codigo, ParcMax, Parcela FROM TRPAporteCred ORDER BY Cliente, Codigo", _
    iTipoPagtoAporte, lNumIntDoc, lNumIntDocAporte, lCliente, dValor, dSaldo, dtVencimento, lNumIntDocDestino, lCodigo, iNumParcelas, iParcela)
    If lErro <> AD_SQL_SUCESSO Then gError 190953
    
    lErro = Comando_BuscarPrimeiro(lComando4)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190954
    
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objFatAporteCred = New ClassTRPFatAporteCred
        Set objDebitoRecCli = New ClassDebitoRecCli
    
        objFatAporteCred.dSaldoAnt = dSaldo
        objFatAporteCred.dSaldoAtual = dSaldo
        objFatAporteCred.dtDataVenc = dtVencimento
        objFatAporteCred.dValor = dValor
        objFatAporteCred.iTipoPagtoAporte = iTipoPagtoAporte
        objFatAporteCred.lCliente = lCliente
        objFatAporteCred.lCodigo = lCodigo
        objFatAporteCred.lNumIntDocAporte = lNumIntDocAporte
        objFatAporteCred.lNumIntDocCredito = lNumIntDocDestino
        objFatAporteCred.lNumIntDocPagto = lNumIntDoc
        objFatAporteCred.iNumParcelas = iNumParcelas
        objFatAporteCred.iParcelaAporte = iParcela
        
        objDebitoRecCli.lNumIntDoc = lNumIntDocDestino
        objDebitoRecCli.iFilialEmpresa = FILIAL_MATRIZ
        
        'Lê os dados para dos Débitos
        lErro = CF("DebitoReceber_Le", objDebitoRecCli)
        If lErro <> SUCESSO And lErro <> 17835 Then gError 190955
        
        Set objFatAporteCred.objCredito = objDebitoRecCli
    
        objFaturamento.colAportesCred.Add objFatAporteCred
    
        lErro = Comando_BuscarProximo(lComando4)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190956
    
    Loop

    TRPFaturamento_Le_Dados2 = SUCESSO

    Exit Function

Erro_TRPFaturamento_Le_Dados2:

    TRPFaturamento_Le_Dados2 = gErr

    Select Case gErr
    
        Case 190953, 190954, 190956
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPAPORTEPAGTOFAT", gErr)
            
        Case 190955

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190957)

    End Select

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function TRPVouchers_Le(ByVal objVou As ClassTRPVouchers, Optional ByVal bLeTudo As Boolean = True, Optional ByVal bLocka As Boolean = False, Optional ByVal lComando As Long = 0) As Long

Dim lErro As Long
Dim alComando(1 To 4) As Long
Dim iIndice As Integer
Dim tVou As typeTRPVouchers
Dim objEndereco As New ClassEndereco

On Error GoTo Erro_TRPVouchers_Le

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
    
        If lComando <> 0 And iIndice = 1 Then
            alComando(iIndice) = lComando
        Else
            alComando(iIndice) = Comando_Abrir()
            If alComando(iIndice) = 0 Then gError 190841
        End If
    Next

    'Alocação de espaço no buffer
    tVou.sSerie = String(STRING_TRP_OCR_SERIE, 0)
    tVou.sTipVou = String(STRING_TRP_OCR_TIPOVOU, 0)
    tVou.sTitular = String(STRING_TRP_VOU_TITULAR, 0)
    tVou.sProduto = String(STRING_PRODUTO, 0)
    tVou.sControle = String(STRING_TRP_VOU_CONTROLE, 0)
    tVou.sCiaCart = String(STRING_TRP_VOU_CIACART, 0)
    tVou.sNumCCred = String(STRING_TRP_VOU_NUMCCRED, 0)
    tVou.sPassageiroNome = String(STRING_TRP_VOU_PAXNOME, 0)
    tVou.sPassageiroSobreNome = String(STRING_TRP_VOU_PAXNOME, 0)
    tVou.sUsuarioCanc = String(STRING_USUARIO_CODIGO, 0)
    tVou.sUsuarioManut = String(STRING_USUARIO_CODIGO, 0)
    tVou.sCiaaerea = String(STRING_TRP_VOU_CIAAEREA, 0)
    tVou.sAeroportos = String(STRING_TRP_VOU_AEROPORTOS, 0)
    tVou.sTitularCPF = String(STRING_CGC, 0)
    tVou.sPassageiroCGC = String(STRING_CGC, 0)
    tVou.sValidadeCC = String(STRING_TRP_VOU_VALIDADECC, 0)
    tVou.sObservacao = String(STRING_OBSERVACAO_OBSERVACAO, 0)

    'Le a tabelaTRPVouchers
    If Not bLocka Then
        lErro = Comando_Executar(alComando(1), "SELECT NumIntDoc, NumVou, Serie, ClienteFat, Valor, Data, " & _
                    "TipVou, NumFat, TipoDocDestino, NumIntDocDestino, Cartao, Status, NumIntDocNVL, " & _
                    "ValorAporte, ValorComissao, Titular, Produto, Pax, Moeda, ValorCambio, " & _
                    "Cambio, Controle, CiaCart, NumCCred, NumAuto, QuantParc, DiasAntc, " & _
                    "Kit, ValorOcr, TemOcr, Representante, ComissaoRep, Correntista, ComissaoCorr, " & _
                    "Emissor, ComissaoEmissor, ComissaoAg, ValorBruto, PassageiroNome, PassageiroSobrenome, ClienteVou, " & _
                    "CliPassageiro, DataCanc, DataVigenciaDe, DataVigenciaAte, HoraCanc, UsuarioCanc, ClienteComissao, " & _
                    "ciaaerea, Aeroportos, EnderecoPaxTitular, TemQueContabilizar, Promotor, Destino, TitularCPF, " & _
                    "TarifaUnitaria, Vigencia, TarifaUnitariaFolheto, PassageiroCGC, PassageiroDataNasc, GeraComissao, CancelaComissao, Idioma, Grupo, UsuarioManut, DataUltimaManut, HoraUltimaManut, DataAutoCC, ValidadeCC,ImprimirValor,CodSegurancaCC,ValorBrutoComOCR,ValorCMAComOCR,ValorCMC,ValorCMR,ValorCMCC,ValorCME,ComissaoProm, Observacao FROM TRPVouchers WHERE NumVou= ?  AND Serie= ?  AND TipVou= ? ", _
                    tVou.lNumIntDoc, tVou.lNumVou, tVou.sSerie, tVou.lClienteFat, _
                    tVou.dValor, tVou.dtData, tVou.sTipVou, tVou.lNumFat, tVou.iTipoDocDestino, _
                    tVou.lNumIntDocDestino, tVou.iCartao, tVou.iStatus, tVou.lNumIntDocNVL, tVou.dValorAporte, _
                    tVou.dValorComissao, tVou.sTitular, tVou.sProduto, tVou.iPax, tVou.iMoeda, _
                    tVou.dValorCambio, tVou.dCambio, tVou.sControle, tVou.sCiaCart, tVou.sNumCCred, _
                    tVou.lNumAuto, tVou.iQuantParc, tVou.iDiasAntc, tVou.iKit, tVou.dValorOcr, _
                    tVou.iTemOcr, tVou.lRepresentante, tVou.dComissaoRep, tVou.lCorrentista, tVou.dComissaoCorr, _
                    tVou.lEmissor, tVou.dComissaoEmissor, tVou.dComissaoAg, tVou.dValorBruto, tVou.sPassageiroNome, _
                    tVou.sPassageiroSobreNome, tVou.lClienteVou, tVou.lCliPassageiro, tVou.dtDataCanc, tVou.dtDataVigenciaDe, _
                    tVou.dtDataVigenciaAte, tVou.dHoraCanc, tVou.sUsuarioCanc, tVou.lClienteComissao, tVou.sCiaaerea, _
                    tVou.sAeroportos, tVou.lEnderecoPaxTitular, tVou.iTemQueContabilizar, tVou.iPromotor, tVou.iDestino, _
                    tVou.sTitularCPF, tVou.dTarifaUnitaria, tVou.iVigencia, tVou.dTarifaUnitariaFolheto, tVou.sPassageiroCGC, tVou.dtPassageiroDataNasc, tVou.iGeraComissao, tVou.iCancelaComissao, tVou.iIdioma, tVou.iGrupo, tVou.sUsuarioManut, tVou.dtDataUltimaManut, tVou.dHoraUltimaManut, tVou.dtDataAutoCC, tVou.sValidadeCC, tVou.iImprimirValor, tVou.iCodSegurancaCC, tVou.dValorBrutoComOCR, tVou.dValorCMAComOCR, tVou.dValorCMC, tVou.dValorCMR, tVou.dValorCMCC, tVou.dValorCME, tVou.dComissaoProm, tVou.sObservacao, _
                    objVou.lNumVou, objVou.sSerie, objVou.sTipVou)
    Else
        lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntDoc, NumVou, Serie, ClienteFat, Valor, Data, " & _
                    "TipVou, NumFat, TipoDocDestino, NumIntDocDestino, Cartao, Status, NumIntDocNVL, " & _
                    "ValorAporte, ValorComissao, Titular, Produto, Pax, Moeda, ValorCambio, " & _
                    "Cambio, Controle, CiaCart, NumCCred, NumAuto, QuantParc, DiasAntc, " & _
                    "Kit, ValorOcr, TemOcr, Representante, ComissaoRep, Correntista, ComissaoCorr, " & _
                    "Emissor, ComissaoEmissor, ComissaoAg, ValorBruto, PassageiroNome, PassageiroSobrenome, ClienteVou, " & _
                    "CliPassageiro, DataCanc, DataVigenciaDe, DataVigenciaAte, HoraCanc, UsuarioCanc, ClienteComissao, " & _
                    "ciaaerea, Aeroportos, EnderecoPaxTitular, TemQueContabilizar, Promotor, Destino, TitularCPF, " & _
                    "TarifaUnitaria, Vigencia, TarifaUnitariaFolheto, PassageiroCGC, PassageiroDataNasc, GeraComissao, CancelaComissao, Idioma, Grupo, UsuarioManut, DataUltimaManut, HoraUltimaManut, DataAutoCC, ValidadeCC,ImprimirValor,CodSegurancaCC,ValorBrutoComOCR,ValorCMAComOCR,ValorCMC,ValorCMR,ValorCMCC,ValorCME,ComissaoProm, Observacao FROM TRPVouchers WHERE NumVou= ?  AND Serie= ?  AND TipVou= ? ", 0, _
                    tVou.lNumIntDoc, tVou.lNumVou, tVou.sSerie, tVou.lClienteFat, _
                    tVou.dValor, tVou.dtData, tVou.sTipVou, tVou.lNumFat, tVou.iTipoDocDestino, _
                    tVou.lNumIntDocDestino, tVou.iCartao, tVou.iStatus, tVou.lNumIntDocNVL, tVou.dValorAporte, _
                    tVou.dValorComissao, tVou.sTitular, tVou.sProduto, tVou.iPax, tVou.iMoeda, _
                    tVou.dValorCambio, tVou.dCambio, tVou.sControle, tVou.sCiaCart, tVou.sNumCCred, _
                    tVou.lNumAuto, tVou.iQuantParc, tVou.iDiasAntc, tVou.iKit, tVou.dValorOcr, _
                    tVou.iTemOcr, tVou.lRepresentante, tVou.dComissaoRep, tVou.lCorrentista, tVou.dComissaoCorr, _
                    tVou.lEmissor, tVou.dComissaoEmissor, tVou.dComissaoAg, tVou.dValorBruto, tVou.sPassageiroNome, _
                    tVou.sPassageiroSobreNome, tVou.lClienteVou, tVou.lCliPassageiro, tVou.dtDataCanc, tVou.dtDataVigenciaDe, _
                    tVou.dtDataVigenciaAte, tVou.dHoraCanc, tVou.sUsuarioCanc, tVou.lClienteComissao, tVou.sCiaaerea, _
                    tVou.sAeroportos, tVou.lEnderecoPaxTitular, tVou.iTemQueContabilizar, tVou.iPromotor, tVou.iDestino, _
                    tVou.sTitularCPF, tVou.dTarifaUnitaria, tVou.iVigencia, tVou.dTarifaUnitariaFolheto, tVou.sPassageiroCGC, tVou.dtPassageiroDataNasc, tVou.iGeraComissao, tVou.iCancelaComissao, tVou.iIdioma, tVou.iGrupo, tVou.sUsuarioManut, tVou.dtDataUltimaManut, tVou.dHoraUltimaManut, tVou.dtDataAutoCC, tVou.sValidadeCC, tVou.iImprimirValor, tVou.iCodSegurancaCC, tVou.dValorBrutoComOCR, tVou.dValorCMAComOCR, tVou.dValorCMC, tVou.dValorCMR, tVou.dValorCMCC, tVou.dValorCME, tVou.dComissaoProm, tVou.sObservacao, _
                    objVou.lNumVou, objVou.sSerie, objVou.sTipVou)
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 190842

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190843

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS
    
    If bLocka Then
        lErro = Comando_LockExclusive(lComando)
        If lErro <> SUCESSO Then gError 190843
    End If

    objVou.lNumVou = tVou.lNumVou
    objVou.lNumIntDoc = tVou.lNumIntDoc
    objVou.lNumVou = tVou.lNumVou
    objVou.sSerie = tVou.sSerie
    objVou.lClienteFat = tVou.lClienteFat
    objVou.dValor = tVou.dValor
    objVou.dtData = tVou.dtData
    objVou.sTipVou = tVou.sTipVou
    objVou.lNumFat = tVou.lNumFat
    objVou.iTipoDocDestino = tVou.iTipoDocDestino
    objVou.lNumIntDocDestino = tVou.lNumIntDocDestino
    objVou.iCartao = tVou.iCartao
    objVou.iStatus = tVou.iStatus
    objVou.lNumIntDocNVL = tVou.lNumIntDocNVL
    objVou.dValorAporte = tVou.dValorAporte
    objVou.dValorComissao = tVou.dValorComissao
    objVou.sTitular = tVou.sTitular
    objVou.sProduto = tVou.sProduto
    objVou.iPax = tVou.iPax
    objVou.iMoeda = tVou.iMoeda
    objVou.dValorCambio = tVou.dValorCambio
    objVou.dCambio = tVou.dCambio
    objVou.sControle = tVou.sControle
    objVou.sCiaCart = tVou.sCiaCart
    objVou.sNumCCred = tVou.sNumCCred
    objVou.lNumAuto = tVou.lNumAuto
    objVou.iQuantParc = tVou.iQuantParc
    objVou.iDiasAntc = tVou.iDiasAntc
    objVou.iKit = tVou.iKit
    objVou.dValorOcr = tVou.dValorOcr
    objVou.iTemOcr = tVou.iTemOcr
    objVou.lRepresentante = tVou.lRepresentante
    objVou.dComissaoRep = tVou.dComissaoRep
    objVou.lCorrentista = tVou.lCorrentista
    objVou.dComissaoCorr = tVou.dComissaoCorr
    objVou.lEmissor = tVou.lEmissor
    objVou.dComissaoEmissor = tVou.dComissaoEmissor
    objVou.dComissaoAg = tVou.dComissaoAg
    objVou.dValorBruto = tVou.dValorBruto
    objVou.sPassageiroNome = tVou.sPassageiroNome
    objVou.sPassageiroSobreNome = tVou.sPassageiroSobreNome
    objVou.lClienteVou = tVou.lClienteVou
    objVou.lCliPassageiro = tVou.lCliPassageiro
    objVou.dtDataCanc = tVou.dtDataCanc
    objVou.dtDataVigenciaDe = tVou.dtDataVigenciaDe
    objVou.dtDataVigenciaAte = tVou.dtDataVigenciaAte
    objVou.dHoraCanc = tVou.dHoraCanc
    objVou.sUsuarioCanc = tVou.sUsuarioCanc
    objVou.lClienteComissao = tVou.lClienteComissao
    objVou.sCiaaerea = tVou.sCiaaerea
    objVou.sAeroportos = tVou.sAeroportos
    objVou.lEnderecoPaxTitular = tVou.lEnderecoPaxTitular
    objVou.iTemQueContabilizar = tVou.iTemQueContabilizar
    objVou.iPromotor = tVou.iPromotor
    objVou.iDestino = tVou.iDestino
    objVou.sTitularCPF = tVou.sTitularCPF
    objVou.dTarifaUnitaria = tVou.dTarifaUnitaria
    objVou.iVigencia = tVou.iVigencia
    objVou.dTarifaUnitariaFolheto = tVou.dTarifaUnitariaFolheto
    objVou.sPassageiroCGC = tVou.sPassageiroCGC
    objVou.dtPassageiroDataNasc = tVou.dtPassageiroDataNasc
    objVou.iGeraComissao = tVou.iGeraComissao
    objVou.iCancelaComissao = tVou.iCancelaComissao
    objVou.iIdioma = tVou.iIdioma
    objVou.iGrupo = tVou.iGrupo
    objVou.sUsuarioManut = tVou.sUsuarioManut
    objVou.dtDataUltimaManut = tVou.dtDataUltimaManut
    objVou.dHoraUltimaManut = tVou.dHoraUltimaManut
    objVou.dtDataAutoCC = tVou.dtDataAutoCC
    objVou.sValidadeCC = tVou.sValidadeCC
    objVou.sObservacao = tVou.sObservacao
    
    objVou.iImprimirValor = tVou.iImprimirValor
    objVou.iCodSegurancaCC = tVou.iCodSegurancaCC
    objVou.dValorBrutoComOCR = tVou.dValorBrutoComOCR
    objVou.dValorCMAComOCR = tVou.dValorCMAComOCR
    objVou.dValorCMC = tVou.dValorCMC
    objVou.dValorCMR = tVou.dValorCMR
    objVou.dValorCMCC = tVou.dValorCMCC
    objVou.dValorCME = tVou.dValorCME
    objVou.dComissaoProm = tVou.dComissaoProm

    If objVou.lClienteComissao = 0 Then objVou.lClienteComissao = objVou.lClienteVou
    
    If bLeTudo Then
    
        lErro = TRPVouContatos_Le(alComando(2), objVou)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 200322
    
        lErro = TRPVouPassageiros_Le(alComando(3), objVou)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 200323
    
        lErro = TRPVouVendedores_Le(alComando(4), objVou)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 200323
    
        objEndereco.lCodigo = objVou.lEnderecoPaxTitular
        
        lErro = CF("Endereco_Le", objEndereco)
        If lErro <> SUCESSO And lErro <> 12309 Then gError 200272
        
        lErro = CF("Endereco_Le_Customizado", objEndereco)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 200273
    
        Set objVou.objEnderecoPax = objEndereco
        
    End If

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        If lComando = 0 Or iIndice <> 1 Then
            Call Comando_Fechar(alComando(iIndice))
        End If
    Next

    TRPVouchers_Le = SUCESSO

    Exit Function

Erro_TRPVouchers_Le:

    TRPVouchers_Le = gErr

    Select Case gErr

        Case 190841
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 190842, 190843
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPVOUCHERS", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora
        
        Case 200272, 200273, 200322, 200323

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190844)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        If lComando = 0 Or iIndice <> 1 Then
            Call Comando_Fechar(alComando(iIndice))
        End If
    Next

    Exit Function

End Function

Private Function TRPAportePagtoFat_Le_NumIntDoc(ByVal lComando As Long, ByVal objTRPAportePagtoFat As ClassTRPAportePagtoFat) As Long

Dim lErro As Long
Dim tTRPAportePagtoFat As typeTRPAportePagtoFat

On Error GoTo Erro_TRPAportePagtoFat_Le_NumIntDoc

    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntDocAporte, Valor, ValidadeDe, ValidadeAte, Saldo, " & _
                    "Percentual FROM TRPAportePagtoFat WHERE NumIntDoc= ? ORDER BY Seq ", _
                    tTRPAportePagtoFat.lNumIntDoc, tTRPAportePagtoFat.lNumIntDocAporte, tTRPAportePagtoFat.dValor, tTRPAportePagtoFat.dtValidadeDe, _
                    tTRPAportePagtoFat.dtValidadeAte, tTRPAportePagtoFat.dSaldo, tTRPAportePagtoFat.dPercentual, _
                    objTRPAportePagtoFat.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 190950

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190951

    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS
    
    objTRPAportePagtoFat.lNumIntDoc = tTRPAportePagtoFat.lNumIntDoc
    objTRPAportePagtoFat.lNumIntDocAporte = tTRPAportePagtoFat.lNumIntDocAporte
    objTRPAportePagtoFat.dValor = tTRPAportePagtoFat.dValor
    objTRPAportePagtoFat.dtValidadeDe = tTRPAportePagtoFat.dtValidadeDe
    objTRPAportePagtoFat.dtValidadeAte = tTRPAportePagtoFat.dtValidadeAte
    objTRPAportePagtoFat.dSaldo = tTRPAportePagtoFat.dSaldo
    objTRPAportePagtoFat.dPercentual = tTRPAportePagtoFat.dPercentual

    TRPAportePagtoFat_Le_NumIntDoc = SUCESSO

    Exit Function

Erro_TRPAportePagtoFat_Le_NumIntDoc:

    TRPAportePagtoFat_Le_NumIntDoc = gErr

    Select Case gErr

        Case 190950, 190951
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPAPORTEPAGTOFAT", gErr)

        Case ERRO_LEITURA_SEM_DADOS
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190952)

    End Select

    Exit Function

End Function

Function TRPConfig_Le(ByVal sCodigo As String, ByVal iFilialEmpresa As Integer, sConteudo As String) As Long
'Lê o Conteudo da Tabela de TRPConfig de acordo com o código passado

Dim lErro As Long
Dim lComando As Long
On Error GoTo Erro_TRPConfig_Le
    
    'Abrir comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 192067

    sConteudo = String(STRING_CONFIG_CONTEUDO, 0)
        
    'ler registo
    lErro = Comando_Executar(lComando, "SELECT Conteudo FROM TRPConfig WHERE Codigo = ? AND FilialEmpresa = ?", sConteudo, sCodigo, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 192068
    
    'Lê o primeiro registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192069
    
    'Se não encontrou o registro
    If lErro = AD_SQL_SEM_DADOS Then gError 192070
    
    'Fechar comando
    Call Comando_Fechar(lComando)

    TRPConfig_Le = SUCESSO
    
    Exit Function

Erro_TRPConfig_Le:
        
    TRPConfig_Le = gErr
    
    Select Case gErr
    
        Case 192067
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 192068, 192069
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPCONFIG", gErr)

        Case 192070
            Call Rotina_Erro(vbOKOnly, "ERRO_TRPCONFIG_INEXISTENTE", gErr, sCodigo, iFilialEmpresa)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 192071)
     
    End Select
     
    'Fechar comando
    Call Comando_Fechar(lComando)
 
    Exit Function
    
End Function

Function TRPConfig_Le_Lockado(ByVal lComando As Long, ByVal sCodigo As String, ByVal iFilialEmpresa As Integer, sConteudo As String) As Long
'Lê o Conteudo da Tabela de TRPConfig de acordo com o código passado

Dim lErro As Long

On Error GoTo Erro_TRPConfig_Le_Lockado

    sConteudo = String(STRING_CONFIG_CONTEUDO, 0)
        
    'ler registo
    lErro = Comando_ExecutarPos(lComando, "SELECT Conteudo FROM TRPConfig WHERE Codigo = ? AND FilialEmpresa = ?", 0, sConteudo, sCodigo, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 194253
    
    'Lê o primeiro registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194254
    
    'Se não encontrou o registro
    If lErro = AD_SQL_SEM_DADOS Then gError 194255

    TRPConfig_Le_Lockado = SUCESSO
    
    Exit Function

Erro_TRPConfig_Le_Lockado:
        
    TRPConfig_Le_Lockado = gErr
    
    Select Case gErr
    
        Case 194253, 194254
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPCONFIG", gErr)

        Case 194255
            Call Rotina_Erro(vbOKOnly, "ERRO_TRPCONFIG_INEXISTENTE", gErr, sCodigo, iFilialEmpresa)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194256)
     
    End Select
 
    Exit Function
    
End Function

Public Function TRPGeracaoNF_Le_Dados(ByVal objGeracaoNF As ClassTRPGeracaoNF) As Long
'Le os itens a serem faturados com os filtros passados
 
Dim lErro As Long
Dim iIndice As Integer
Dim sSQL As String
Dim sWhere As String
Dim sWhereAux As String
Dim alComando(1 To 7) As Long
Dim vValor As Variant
Dim vValorAux(1 To 100)
Dim iProx As Integer
Dim vdtDataDe As Variant
Dim vdtDataAte As Variant
Dim vdtDataBaixaDe As Variant
Dim vdtDataBaixaAte As Variant
Dim vlTituloDe As Variant
Dim vlTituloAte As Variant
Dim vlNumIntDoc As Variant
Dim viStatus As Variant
Dim vsTipoDoc As Variant
Dim objTitulosRec As ClassTituloReceber
Dim objTitulosRecTRP As ClassTitulosRecTRP
Dim objNFFilial As ClassTRPGeracaoNFFiliais

Dim viFilialEmpresa As Variant
Dim vlNumTitulo As Variant
Dim vdtDataEmissao As Variant
Dim vdValor As Variant
Dim vsNomeCliente As Variant
Dim vlCliente As Variant
Dim vlClienteVou As Variant
Dim vlCliPassageiro As Variant
Dim viNFAoPassageiro As Variant
Dim viUmaNFPorFatura As Variant
Dim vsSiglaDocumento As Variant

On Error GoTo Erro_TRPGeracaoNF_Le_Dados
 
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_AbrirExt(GL_lConexaoRel)
        If alComando(iIndice) = 0 Then gError 192220
    Next
    
    vlNumIntDoc = CLng(0)
    viStatus = STATUS_BAIXADO
    vdtDataDe = objGeracaoNF.dtDataDe
    vdtDataAte = objGeracaoNF.dtDataAte
    vdtDataBaixaDe = objGeracaoNF.dtDataBaixaDe
    vdtDataBaixaAte = objGeracaoNF.dtDataBaixaAte
    vsTipoDoc = objGeracaoNF.sTipoDoc
    vlTituloDe = objGeracaoNF.lTituloDe
    vlTituloAte = objGeracaoNF.lTituloAte
    
    viFilialEmpresa = CInt(0)
    vlNumTitulo = CLng(0)
    vdtDataEmissao = DATA_NULA
    vdValor = CDbl(0)
    vsNomeCliente = String(STRING_CLIENTE_NOME_REDUZIDO, 0)
    vlCliente = CLng(0)
    vlClienteVou = CLng(0)
    vlCliPassageiro = CLng(0)
    vsSiglaDocumento = String(STRING_SIGLA_DOCUMENTO, 0)
    viNFAoPassageiro = CInt(0)
    viUmaNFPorFatura = CInt(0)

    sSQL = "SELECT NumIntDoc,FilialEmpresa, NumTitulo, DataEmissao, Valor, NomeCliente, Cliente, ClienteVou, CliPassageiro, SiglaDocumento, NFAoPassageiro, UmaNFPorFatura FROM TitulosSemNota WHERE "

    sWhere = " ("
    sWhereAux = ""
    For Each objNFFilial In objGeracaoNF.colFiliais
        If Len(Trim(sWhereAux)) > 0 Then sWhereAux = sWhereAux & " OR "
        sWhereAux = sWhereAux & "FilialEmpresa = ? "
    Next
    sWhere = sWhere & sWhereAux & ") "
    
    If objGeracaoNF.dtDataDe <> DATA_NULA Then
        sWhere = sWhere & "AND DataEmissao >= ? "
    End If
    
    If objGeracaoNF.dtDataAte <> DATA_NULA Then
        sWhere = sWhere & "AND DataEmissao <= ? "
    End If
    
    If objGeracaoNF.dtDataBaixaDe <> DATA_NULA Then
        sWhere = sWhere & "AND DataBaixa >= ? "
    End If
    
    If objGeracaoNF.dtDataBaixaAte <> DATA_NULA Then
        sWhere = sWhere & "AND DataBaixa <= ? "
    End If
    
    If objGeracaoNF.lTituloDe <> 0 Then
        sWhere = sWhere & "AND NumTitulo >= ? "
    End If
    
    If objGeracaoNF.lTituloAte <> 0 Then
        sWhere = sWhere & "AND NumTitulo <= ? "
    End If
    
    If objGeracaoNF.iTipo = TRP_GERACAONF_TITULOS_BAIXADO Then
        sWhere = sWhere & "AND Status = ? "
    End If
    
    If objGeracaoNF.sTipoDoc <> "" Then
        sWhere = sWhere & "AND SiglaDocumento = ? "
    End If
    
    objGeracaoNF.sFiltro = sWhere
    
    sSQL = sSQL & sWhere & " ORDER BY FilialEmpresa, Cliente"
    
    lErro = Comando_PrepararInt(alComando(1), sSQL)
    If (lErro <> AD_SQL_SUCESSO) Then gError 192221
 
    lErro = Comando_BindVarInt(alComando(1), vlNumIntDoc)
    If (lErro <> AD_SQL_SUCESSO) Then gError 192222

       
    lErro = Comando_BindVarInt(alComando(1), viFilialEmpresa)
    If (lErro <> AD_SQL_SUCESSO) Then gError 192222
       
    lErro = Comando_BindVarInt(alComando(1), vlNumTitulo)
    If (lErro <> AD_SQL_SUCESSO) Then gError 192222
       
    lErro = Comando_BindVarInt(alComando(1), vdtDataEmissao)
    If (lErro <> AD_SQL_SUCESSO) Then gError 192222
       
    lErro = Comando_BindVarInt(alComando(1), vdValor)
    If (lErro <> AD_SQL_SUCESSO) Then gError 192222
       
    lErro = Comando_BindVarInt(alComando(1), vsNomeCliente)
    If (lErro <> AD_SQL_SUCESSO) Then gError 192222
       
    lErro = Comando_BindVarInt(alComando(1), vlCliente)
    If (lErro <> AD_SQL_SUCESSO) Then gError 192222
      
    lErro = Comando_BindVarInt(alComando(1), vlClienteVou)
    If (lErro <> AD_SQL_SUCESSO) Then gError 192222
      
    lErro = Comando_BindVarInt(alComando(1), vlCliPassageiro)
    If (lErro <> AD_SQL_SUCESSO) Then gError 192222
      
    lErro = Comando_BindVarInt(alComando(1), vsSiglaDocumento)
    If (lErro <> AD_SQL_SUCESSO) Then gError 192222
      
    lErro = Comando_BindVarInt(alComando(1), viNFAoPassageiro)
    If (lErro <> AD_SQL_SUCESSO) Then gError 192222
      
    lErro = Comando_BindVarInt(alComando(1), viUmaNFPorFatura)
    If (lErro <> AD_SQL_SUCESSO) Then gError 192222
     
    For Each objNFFilial In objGeracaoNF.colFiliais
        
        iProx = iProx + 1
        vValorAux(iProx) = objNFFilial.iFilialEmpresa
    
        lErro = Comando_BindVarInt(alComando(1), vValorAux(iProx))
        If (lErro <> AD_SQL_SUCESSO) Then gError 192223
        
        objGeracaoNF.colSelecao.Add vValorAux(iProx)

    Next
    
    If objGeracaoNF.dtDataDe <> DATA_NULA Then
        lErro = Comando_BindVarInt(alComando(1), vdtDataDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError 192224
    
        objGeracaoNF.colSelecao.Add vdtDataDe
    End If
    
    If objGeracaoNF.dtDataAte <> DATA_NULA Then
        lErro = Comando_BindVarInt(alComando(1), vdtDataAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError 192225
    
        objGeracaoNF.colSelecao.Add vdtDataAte
    End If
    
    If objGeracaoNF.dtDataBaixaDe <> DATA_NULA Then
        lErro = Comando_BindVarInt(alComando(1), vdtDataBaixaDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError 192224
    
        objGeracaoNF.colSelecao.Add vdtDataBaixaDe
    End If
    
    If objGeracaoNF.dtDataBaixaAte <> DATA_NULA Then
        lErro = Comando_BindVarInt(alComando(1), vdtDataBaixaAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError 192225
    
        objGeracaoNF.colSelecao.Add vdtDataBaixaAte
    End If
    
    If objGeracaoNF.lTituloDe <> 0 Then
        lErro = Comando_BindVarInt(alComando(1), vlTituloDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError 192224
    
        objGeracaoNF.colSelecao.Add vlTituloDe
    End If
    
    If objGeracaoNF.lTituloAte <> 0 Then
        lErro = Comando_BindVarInt(alComando(1), vlTituloAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError 192225
    
        objGeracaoNF.colSelecao.Add vlTituloAte
    End If
    
    If objGeracaoNF.iTipo = TRP_GERACAONF_TITULOS_BAIXADO Then
        lErro = Comando_BindVarInt(alComando(1), viStatus)
        If (lErro <> AD_SQL_SUCESSO) Then gError 192226
    
        objGeracaoNF.colSelecao.Add viStatus
    End If
    
    If objGeracaoNF.sTipoDoc <> "" Then
        lErro = Comando_BindVarInt(alComando(1), vsTipoDoc)
        If (lErro <> AD_SQL_SUCESSO) Then gError 192306
    
        objGeracaoNF.colSelecao.Add vsTipoDoc
    End If
    
    lErro = Comando_ExecutarInt(alComando(1))
    If (lErro <> AD_SQL_SUCESSO) Then gError 192227
 
    'Busca o primeiro registro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192228
        
    'Para cada Produto
    Do While lErro = SUCESSO
    
        Set objTitulosRec = New ClassTituloReceber
        Set objTitulosRecTRP = New ClassTitulosRecTRP
    
        objTitulosRec.lNumIntDoc = vlNumIntDoc

        objTitulosRec.iFilialEmpresa = viFilialEmpresa
        objTitulosRec.lNumTitulo = vlNumTitulo
        objTitulosRec.dtDataEmissao = vdtDataEmissao
        objTitulosRec.dValor = vdValor
        objTitulosRec.sSiglaDocumento = vsSiglaDocumento
        objTitulosRec.lCliente = vlCliente
        objTitulosRec.iFilial = FILIAL_MATRIZ
        
        objTitulosRecTRP.lClienteVou = vlClienteVou
        objTitulosRecTRP.sNomeCliente = vsNomeCliente
        objTitulosRecTRP.lCliPassageiro = vlCliPassageiro
        objTitulosRecTRP.iNFAoPassageiro = viNFAoPassageiro
        objTitulosRecTRP.iUmaNFPorFatura = viUmaNFPorFatura
        
        If objTitulosRec.sSiglaDocumento = TIPODOC_CARTAO_CRED_DEB Then
            objTitulosRecTRP.iNFAoPassageiro = MARCADO
        End If

        Set objTitulosRec.objInfoUsu = objTitulosRecTRP
        
        objGeracaoNF.colTitulos.Add objTitulosRec
    
        'Busca o próximo registro
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192229
    
    Loop
    
    lErro = TRPGeracaoNF_Calcula_Dados(alComando(2), alComando(3), alComando(4), alComando(5), alComando(6), alComando(7), objGeracaoNF)
    If lErro <> SUCESSO Then gError 192230
        
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
 
    TRPGeracaoNF_Le_Dados = SUCESSO
 
    Exit Function
 
Erro_TRPGeracaoNF_Le_Dados:
 
    TRPGeracaoNF_Le_Dados = gErr
 
    Select Case gErr
 
        Case 192220
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 192221 To 192229
            Call Rotina_Erro(vbOKOnly, "ERRO_PREPARACAO_COMANDO_SQL", gErr, sSQL)
            
        Case 192230, 192233, 192234, 192306
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192231)
 
    End Select
 
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
 
    Exit Function
 
End Function

Private Function TRPGeracaoNF_Calcula_Dados(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lComando3 As Long, ByVal lComando4 As Long, ByVal lComando5 As Long, ByVal lComando6 As Long, ByVal objGeracaoNF As ClassTRPGeracaoNF) As Long
 
Dim lErro As Long
Dim objTitulosRec As ClassTituloReceber
Dim objTitulosRecTRP As ClassTitulosRecTRP
Dim objNF As ClassNFiscal
Dim objItemNF As ClassItemNF
Dim objProduto As New ClassProduto
Dim objTribTab As ClassTribTab
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim lClienteAnt As Long
Dim iFilialEmpresa As Integer
Dim iFilialEmpresaAnt As Long
Dim iItemNF As Integer
Dim iNumNotas As Integer
Dim iNumTitulo As Integer
Dim dValor As Double
Dim iIndice As Integer
Dim objGeracaoNFItem As ClassTRPGeracaoNFItem
Dim objSerie As ClassSerie
Dim lCliente As Long
Dim iFilialCliente As Integer
Dim objNFTRP As ClassNFTRP
Dim bPrimeiroItem As Boolean
Dim objTipoTribICMS As New ClassTipoTribICMS
Dim objTipoTribIPI As New ClassTipoTribIPI
Dim objTipoDeTributacao As New ClassTipoDeTributacaoMovto

Const NATUREZAOP_TRP = "5001"
Const TIPO_TRIBUTACAO_TRP = 3
Const SERIE_TRP = "1"
Const PRODUTO_TRP = "FAT       "
Dim colSaida As New Collection
Dim colCampos As New Collection

On Error GoTo Erro_TRPGeracaoNF_Calcula_Dados

    bPrimeiroItem = True

    'Le o tipo de nota que vai ser usado
    objTipoDocInfo.iCodigo = DOCINFO_NFISS
    
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 192235
    
    'Le o produto que vai ser usado nos itens da Nota
    objProduto.sCodigo = PRODUTO_TRP
    
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 192236
    
    'Se o produto não existe => Erro
    If lErro <> SUCESSO Then gError 192237
    
    lErro = CF("Produto_Le_Categorias", objProduto, objProduto.colCategoriaItem)
    If lErro <> SUCESSO Then gError 192236
        
    'Se for uma nota fiscal de serviço
    If objTipoDocInfo.iCodigo = DOCINFO_NFISS Then
        If Not (objProduto.iCompras = PRODUTO_PRODUZIVEL And objProduto.iFaturamento = PRODUTO_VENDAVEL And objProduto.iControleEstoque = PRODUTO_CONTROLE_SEM_ESTOQUE) Then gError 192238
    End If

    '#########################################################################################
    'Preenche nota e itens
    For Each objTitulosRec In objGeracaoNF.colTitulos
    
        'Verifica se o cliente não deve ser um passageiro
        lErro = TRP_GeracaoNF_Obtem_Cliente(lComando1, lComando2, objGeracaoNF, objTitulosRec, lCliente, iFilialCliente)
        If lErro <> SUCESSO Then gError 192727
        
        objTitulosRec.lCliente = lCliente
        objTitulosRec.iFilial = iFilialCliente
            
    Next
    
    colCampos.Add "iFilialEmpresa"
    colCampos.Add "lCliente"
    colCampos.Add "iFilial"
    
    Call Ordena_Colecao(objGeracaoNF.colTitulos, colSaida, colCampos)
    
    For iIndice = colSaida.Count To 1 Step -1
        objGeracaoNF.colTitulos.Remove (iIndice)
    Next
    
    For Each objTitulosRec In colSaida
        objGeracaoNF.colTitulos.Add objTitulosRec
    Next
        
    For Each objTitulosRec In objGeracaoNF.colTitulos
        
        lCliente = objTitulosRec.lCliente
        iFilialCliente = objTitulosRec.iFilial
        
        Set objTitulosRecTRP = objTitulosRec.objInfoUsu
        
        iFilialEmpresa = objTitulosRec.iFilialEmpresa
        
        If lClienteAnt <> lCliente Or iFilialEmpresaAnt <> iFilialEmpresa Or (objTitulosRecTRP.iUmaNFPorFatura = MARCADO Or objGeracaoNF.iGerarNFParaCadaFat = MARCADO) Then
        
            Set objNF = New ClassNFiscal
            Set objNFTRP = New ClassNFTRP
            Set objNF.objInfoUsu = objNFTRP
            
            Set objNFTRP.objProduto = objProduto
            Set objNFTRP.objTipoDocInfo = objTipoDocInfo
            
            objNF.iFilialEmpresa = iFilialEmpresa
            objNF.sSerie = SERIE_TRP
                       
            objNF.lCliente = lCliente
            objNF.iFilialCli = iFilialCliente
            objNF.iFilialEntrega = iFilialCliente
            
            objNF.dtDataEmissao = objGeracaoNF.dtDataEmissao
            objNF.iTipoNFiscal = objTipoDocInfo.iCodigo
            objNF.iTipoDocInfo = objTipoDocInfo.iTipo
            objNF.iStatus = STATUS_LANCADO
            
            'Obtém dados fixos
            objNF.iNaoVerificaCredito = NAO_VERIFICA_CREDITO_CLIENTE
            objNF.dtDataReferencia = DATA_NULA
            objNF.iFreteRespons = FRETE_EMITENTE
            objNF.dtDataRegistro = Date
            objNF.dtDataEntrada = DATA_NULA
            objNF.dtDataVencimento = DATA_NULA
            objNF.sNaturezaOp = NATUREZAOP_TRP
                   
            'Obtém dados do Registro
            objNF.dtDataSaida = objGeracaoNF.dtDataEmissao
                        
            'Inicializa o rastreamento
            Set objNF.objRastreamento = New ClassRastreamento
        
            lClienteAnt = objTitulosRec.lCliente
            iFilialEmpresaAnt = iFilialEmpresa
            
            objGeracaoNF.colNF.Add objNF
            objGeracaoNF.colNFFilialO.Add iFilialEmpresa
            
            iItemNF = 0
            
        End If
        
        iItemNF = iItemNF + 1
        
        Set objItemNF = New ClassItemNF
    
        'Obtém dados fixos
        objItemNF.dtDataEntrega = DATA_NULA

        objItemNF.sProduto = objProduto.sCodigo
        objItemNF.dQuantidade = 1
        objItemNF.dPrecoUnitario = objTitulosRec.dValor

        objItemNF.sUnidadeMed = objProduto.sSiglaUMVenda
        objItemNF.sUMEstoque = objProduto.sSiglaUMEstoque
        objItemNF.sUMVenda = objProduto.sSiglaUMVenda
        objItemNF.iClasseUM = objProduto.iClasseUM
        
        objItemNF.iItem = iItemNF
    
        objNF.ColItensNF.Add1 objItemNF
                
        objNF.dValorTotal = objNF.dValorTotal + objTitulosRec.dValor
        objNF.dValorProdutos = objNF.dValorTotal
    
    Next
    '#########################################################################################
          
    '#########################################################################################
    'Calcula as informações por filial antes da redis.
    iFilialEmpresaAnt = 0
    iIndice = 0
    For Each objNF In objGeracaoNF.colNF
 
        iIndice = iIndice + 1
        
        If iFilialEmpresaAnt <> objNF.iFilialEmpresa Then
        
            If iIndice <> 1 Then
            
                Set objGeracaoNFItem = New ClassTRPGeracaoNFItem
                
                objGeracaoNFItem.dValor = dValor
                objGeracaoNFItem.iFilialEmpresa = iFilialEmpresaAnt
                
                objGeracaoNF.colItens.Add objGeracaoNFItem
            
            End If
 
            iNumNotas = 0
            iNumTitulo = 0
            dValor = 0
            iFilialEmpresaAnt = objNF.iFilialEmpresa
        End If
        
        iNumNotas = iNumNotas + 1
        iNumTitulo = iNumTitulo + objNF.ColItensNF.Count
        dValor = dValor + objNF.dValorTotal
        
    Next
    
    Set objGeracaoNFItem = New ClassTRPGeracaoNFItem
    
    objGeracaoNFItem.dValor = dValor
    objGeracaoNFItem.iFilialEmpresa = iFilialEmpresaAnt
    
    objGeracaoNF.colItens.Add objGeracaoNFItem
    '#########################################################################################
    
    'Redis. as NFs pelas filiais
    lErro = TRPGeracaoNF_Calcula_Dados2(lComando3, lComando4, lComando5, lComando6, objGeracaoNF)
    If lErro <> SUCESSO Then gError 192241
    
    '#########################################################################################
    'Acerta a parte tributária
    For Each objNF In objGeracaoNF.colNF
    
        'Inicializa a Parte tributária da Nota
        Call objNF.Inicializa_Tributacao
        Call objNF.objTributacaoNF.Coloca_Auto
        
        objNF.objTributacaoNF.iTipoTributacao = TIPO_TRIBUTACAO_TRP
        objNF.objTributacaoNF.iNaturezaOpManual = VAR_PREENCH_MANUAL
        objNF.objTributacaoNF.sNaturezaOpInterna = NATUREZAOP_TRP
        objNF.objTributacaoNF.iNaturezaOpManual = VAR_PREENCH_MANUAL
                
        For Each objItemNF In objNF.ColItensNF
        
            'Inicializa parte tributária do item da Nota
            Call objItemNF.Inicializa_Tributacao
            Call objItemNF.objTributacaoItemNF.Coloca_Auto
           
            objItemNF.objTributacaoItemNF.iTipoTributacao = TIPO_TRIBUTACAO_TRP
            objItemNF.objTributacaoItemNF.iTipoTributacaoManual = VAR_PREENCH_MANUAL
            objItemNF.objTributacaoItemNF.sNaturezaOp = objNF.objTributacaoNF.sNaturezaOpInterna
            objItemNF.objTributacaoItemNF.iNaturezaOpManual = VAR_PREENCH_MANUAL
            
            If bPrimeiroItem Then
            
                bPrimeiroItem = False
            
                objTipoDeTributacao.iTipo = objItemNF.objTributacaoItemNF.iTipoTributacao
            
                lErro = CF("TipoTributacao_Le", objTipoDeTributacao)
                If lErro <> SUCESSO And lErro <> 27259 Then gError 196790
               
                'Lê o Tipo de Tributação para pegar a coluna no Livro Fiscal
                objTipoTribICMS.iTipo = objItemNF.objTributacaoItemNF.iICMSTipo
                
                lErro = CF("TipoTribICMS_Le", objTipoTribICMS)
                If lErro <> SUCESSO And lErro <> 21534 Then gError 196791

                                    
                'Lê o Tipo de Tributação para pegar a coluna no Livro Fiscal
                objTipoTribIPI.iTipo = objItemNF.objTributacaoItemNF.iIPITipo
            
                lErro = CF("TipoTribIPI_Le", objTipoTribIPI)
                If lErro <> SUCESSO And lErro <> 21620 Then gError 196792
            
            End If
    
        Next
        
        Set objTribTab = New ClassTribTab
        
        lErro = objTribTab.TributacaoNF_Reset(objNF)
        If lErro <> SUCESSO Then gError 192239
        
        'Atualiza a Tributação
        Call objTribTab.TipoNFiscal_Definir(objTipoDocInfo.iCodigo, objTipoDocInfo.sSigla)
        Call objTribTab.DataEmissao_Alterada(objNF.dtDataEmissao)
        Call objTribTab.Serie_Alterada(objNF.sSerie)
    
        objTribTab.objNFTributacao.iISSInclusoManual = VAR_PREENCH_AUTOMATICO
        
        lErro = gobjTributacao.AtualizaImpostos(objTribTab.mvarobjNFTributacao, 0)
        If lErro <> SUCESSO Then gError 192240
        
        lErro = objTribTab.PreencherNF(objNF)
        If lErro <> SUCESSO Then gError 192241
        
        Set objTribTab = Nothing
        
        Set objNFTRP.objTipoDeTributacaoMovto = objTipoDeTributacao
        Set objNFTRP.objTipoTribIPI = objTipoTribIPI
        Set objNFTRP.objTipoTribICMS = objTipoTribICMS
        
        objNF.sNaturezaOp = objNF.objTributacaoNF.sNaturezaOpInterna
        
        'acertar valor total considerando os tributos, retenções,...
        objNF.dValorTotal = objNF.dValorTotal + objNF.objTributacaoNF.dICMSSubstValor + objNF.objTributacaoNF.dIPIValor + objNF.dValorSeguro + objNF.dValorOutrasDespesas + IIf(objNF.objTributacaoNF.iISSIncluso = 0, objNF.objTributacaoNF.dISSValor, 0)
    
    Next
    '#########################################################################################
    
    '#########################################################################################
    'Acerta a faixa da numeração das Nfs por filial
    For Each objGeracaoNFItem In objGeracaoNF.colItens
    
        Set objSerie = New ClassSerie
        
        objSerie.iFilialEmpresa = objGeracaoNFItem.iFilialEmpresa
        objSerie.sSerie = SERIE_TRP
        
        'Lê a Serie no BD
        lErro = CF("Serie_Le", objSerie)
        If lErro <> SUCESSO Then gError 192242
        
        If objGeracaoNFItem.iNumNF <> 0 Then
            objGeracaoNFItem.lNFDe = objSerie.lProxNumNFiscal
            objGeracaoNFItem.lNFAte = objSerie.lProxNumNFiscal + objGeracaoNFItem.iNumNF - 1
        Else
            objGeracaoNFItem.lNFDe = 0
            objGeracaoNFItem.lNFAte = 0
        End If
        
        objGeracaoNFItem.lNFProx = objSerie.lProxNumNFiscal
        
    Next
    '#########################################################################################
    
    '#########################################################################################
    'Acerta a numeração das notas fiscais
    For Each objNF In objGeracaoNF.colNF

        For Each objGeracaoNFItem In objGeracaoNF.colItens

            If objNF.iFilialEmpresa = objGeracaoNFItem.iFilialEmpresa Then
                objNF.lNumNotaFiscal = objGeracaoNFItem.lNFProx
                objGeracaoNFItem.lNFProx = objGeracaoNFItem.lNFProx + 1
                Exit For
            End If
            
        Next

    Next
    '#########################################################################################
    
    TRPGeracaoNF_Calcula_Dados = SUCESSO
 
    Exit Function
 
Erro_TRPGeracaoNF_Calcula_Dados:
 
    TRPGeracaoNF_Calcula_Dados = gErr
 
    Select Case gErr
    
        Case 192235, 192236, 192239 To 192242, 192727, 196790, 196791, 196792
        
        Case 192237
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
        
        Case 192238
             Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NOTA_SERVICO", gErr, objProduto.sCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192232)
 
    End Select
 
    Exit Function
 
End Function

Public Function TRPGeracaoNF_Calcula_Dados2(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lComando3 As Long, ByVal lComando4 As Long, ByVal objGeracaoNF As ClassTRPGeracaoNF) As Long
'Redistribui as NFs conforme percentual
 
Dim lErro As Long
Dim iIndice As Integer
Dim objGeracaoNFItem As ClassTRPGeracaoNFItem
Dim iFilialO As Integer
Dim iFilialD As Integer
Dim objNFFilial As ClassTRPGeracaoNFFiliais
Dim dValorARedis As Double
Dim dValorFaltaRedis As Double
Dim dValorPodeRedis As Double
Dim objCliente As ClassCliente
Dim objClienteTRP As ClassClienteTRP
Dim colNFPodeRedisFilial As Collection
Dim dFatorBusca As Double
Dim iFatorBusca As Integer
Dim objNF As ClassNFiscal
Dim lClienteAnt As Long
Dim colNFPodeRedisFilialOrd As Collection
Dim colCampos As Collection
Dim bChegou As Boolean
Dim iTentativa As Integer
Dim dValor As Double
Dim iNumNotas As Integer
Dim iNumTitulo As Integer
Dim bAchou As Boolean
Dim iIndiceTit As Integer
Dim objItemNF As ClassItemNF
Dim objTituloRec As ClassTituloReceber
Dim objTituloRecTRP As ClassTitulosRecTRP
Dim objNFTRP As ClassNFTRP

Const MARGEM_ERRO = 0.01
Const FILIAL_REDIRECIONAMENTO = 6
Const NUM_TENTATIVAS = 3

On Error GoTo Erro_TRPGeracaoNF_Calcula_Dados2

    'Redis os clientes fixos
    For Each objNF In objGeracaoNF.colNF
    
        If lClienteAnt <> objNF.lCliente Then

            lClienteAnt = objNF.lCliente

            Set objCliente = New ClassCliente
            
            objCliente.lCodigo = objNF.lCliente
            
            lErro = CF("Cliente_Le", objCliente, lComando1)
            If lErro <> SUCESSO Then gError 196530
            
            lErro = CF("Cliente_Le_Customizado", objCliente, False, lComando2, lComando3, lComando4)
            If lErro <> SUCESSO Then gError 196531
            
        End If
        
        If Not (objCliente.objInfoUsu Is Nothing) Then
            
            Set objClienteTRP = objCliente.objInfoUsu
            
            If objClienteTRP.iFilialNF <> 0 Then
                objNF.iFilialEmpresa = objClienteTRP.iFilialNF
            End If
        
        End If
        
        Set objNFTRP = objNF.objInfoUsu
        Set objNFTRP.objCliente = objCliente
        
        objGeracaoNF.colNFClientes.Add objCliente

    Next
    
    'para cada filial
    For Each objNFFilial In objGeracaoNF.colFiliais
        
        bAchou = False
        For Each objGeracaoNFItem In objGeracaoNF.colItens
            If objGeracaoNFItem.iFilialEmpresa = objNFFilial.iFilialEmpresa Then
                bAchou = True
                Exit For
            End If
        Next
        If Not bAchou Then
            
            Set objGeracaoNFItem = New ClassTRPGeracaoNFItem
            
            objGeracaoNFItem.iFilialEmpresa = objNFFilial.iFilialEmpresa
            
            objGeracaoNF.colItens.Add objGeracaoNFItem
        
        End If
            
    Next

    'para cada filial
    For Each objNFFilial In objGeracaoNF.colFiliais
        
        Set colNFPodeRedisFilial = New Collection
        dValorARedis = 0
        dValorFaltaRedis = 0
        dValorPodeRedis = 0
    
        If objNFFilial.dPercentual > 0 Then
        
            bAchou = False
            For Each objGeracaoNFItem In objGeracaoNF.colItens
                If objGeracaoNFItem.iFilialEmpresa = objNFFilial.iFilialEmpresa Then
                    bAchou = True
                    Exit For
                End If
            Next
            
            'Se tem Nfs para filial
            If bAchou Then
        
                dValorARedis = objNFFilial.dPercentual * objGeracaoNFItem.dValor
                dValorFaltaRedis = dValorARedis
        
                iIndice = 0
                iIndiceTit = 0
                For Each objNF In objGeracaoNF.colNF
                
                    iIndiceTit = iIndiceTit + objNF.ColItensNF.Count
                    Set objItemNF = objNF.ColItensNF.Item(objNF.ColItensNF.Count)
                    Set objTituloRec = objGeracaoNF.colTitulos.Item(iIndiceTit)
                    
                    If Abs(objItemNF.dValorTotal - objTituloRec.dValor) > DELTA_VALORMONETARIO Then gError 192299
                    
                    Set objTituloRecTRP = objTituloRec.objInfoUsu
            
                    iIndice = iIndice + 1
                    If objNF.iFilialEmpresa = objNFFilial.iFilialEmpresa Then
                        
                        Set objCliente = objGeracaoNF.colNFClientes.Item(iIndice)
                        iFilialO = objGeracaoNF.colNFFilialO.Item(iIndice)
                        
                        'Se não é venda direta, se não é passageiro (exceto cartão) e se não é CVC e se não trocou a filial antes
                        If objCliente.iTipo <> 1 And (objCliente.iTipo <> 10 Or objTituloRec.sSiglaDocumento = TIPODOC_CARTAO_CRED_DEB) And InStr(1, objCliente.sNomeReduzido, "CVC") = 0 And InStr(1, objTituloRecTRP.sNomeCliente, "CVC") = 0 And iFilialO = objNF.iFilialEmpresa Then
                            colNFPodeRedisFilial.Add objNF
                            dValorPodeRedis = dValorPodeRedis + objNF.dValorTotal
                        End If
                    
                    End If
            
                Next
                
                If dValorPodeRedis > 0 Then
                
                    dFatorBusca = dValorPodeRedis / dValorARedis
                    
                    Set colCampos = New Collection
                    Set colNFPodeRedisFilialOrd = New Collection
                    
                    colCampos.Add "dValorTotal"
                    
                    lErro = Ordena_Colecao(colNFPodeRedisFilial, colNFPodeRedisFilialOrd, colCampos)
                    If lErro <> SUCESSO Then gError 190962
                    
                    bChegou = False
                    For iTentativa = 1 To NUM_TENTATIVAS
                    
                        dFatorBusca = dFatorBusca + (1 - iTentativa)
                        If dFatorBusca < 0 Then Exit For
                    
                        iFatorBusca = 0
                        For iIndice = colNFPodeRedisFilialOrd.Count To 1 Step -1
                            iFatorBusca = iFatorBusca + 1
                            
                            If iFatorBusca > dFatorBusca Then
                            
                                iFatorBusca = 0
                                
                                Set objNF = colNFPodeRedisFilialOrd.Item(iIndice)
                                
                                If dValorFaltaRedis + MARGEM_ERRO * dValorARedis > objNF.dValorTotal Then
                                
                                    dValorFaltaRedis = dValorFaltaRedis - objNF.dValorTotal
                                    objNF.iFilialEmpresa = FILIAL_REDIRECIONAMENTO
                                    
                                    colNFPodeRedisFilialOrd.Remove iIndice
                                
                                End If
                            
                            End If
                            
                            If Abs(dValorFaltaRedis) < MARGEM_ERRO * dValorARedis Then
                                bChegou = True
                                Exit For
                            End If
                            
                        Next
                        
                        If bChegou Then Exit For
                        
                    Next
                    
                End If
                
            End If
            
        End If
        
    Next
    
    iIndiceTit = 0
    For Each objNF In objGeracaoNF.colNF
        For Each objItemNF In objNF.ColItensNF
            iIndiceTit = iIndiceTit + 1
            Set objTituloRec = objGeracaoNF.colTitulos.Item(iIndiceTit)
            If objNF.iFilialEmpresa = 1 Or objNF.iFilialEmpresa = 2 Then
                objItemNF.sDescricaoItem = "Prestação de serviço de assistência ao viajante, referente a fatura Nro. " & CStr(objTituloRec.lNumTitulo)
            Else
                objItemNF.sDescricaoItem = "Prestação de Serv. de assistência ao viajante, ref. fatura Nro " & CStr(objTituloRec.lNumTitulo)
            End If
        Next
    Next
     
    'Acerta a filial de destino
    For Each objNF In objGeracaoNF.colNF
        objGeracaoNF.colNFFilialD.Add objNF.iFilialEmpresa
    Next

    For Each objGeracaoNFItem In objGeracaoNF.colItens
 
        iIndice = 0
        dValor = 0
        iNumNotas = 0
        iNumTitulo = 0
        For Each objNF In objGeracaoNF.colNF
            
            iIndice = iIndice + 1
            If objNF.iFilialEmpresa = objGeracaoNFItem.iFilialEmpresa Then
        
                iNumNotas = iNumNotas + 1
                iNumTitulo = iNumTitulo + objNF.ColItensNF.Count
                dValor = dValor + objNF.dValorTotal
                
            End If
        
        Next
        
        objGeracaoNFItem.dValorR = dValor
        objGeracaoNFItem.iNumNF = iNumNotas
        objGeracaoNFItem.iNumTitulos = iNumTitulo
                
    Next
                      
    TRPGeracaoNF_Calcula_Dados2 = SUCESSO
 
    Exit Function
 
Erro_TRPGeracaoNF_Calcula_Dados2:
 
    TRPGeracaoNF_Calcula_Dados2 = gErr
 
    Select Case gErr
    
        Case 196530, 196531
            
        Case 192299
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_ITEM_ERRADO", gErr, Format(objItemNF.dValorTotal, "STANDARD"), Format(objTituloRec.dValor, "STANDARD"))
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196532)
 
    End Select
 
    Exit Function
 
End Function

Public Function TRPAportePagtoCond_Le_NumIntDocDestino(ByVal objTRPAportePagto As ClassTRPAportePagtoCond) As Long

Dim lErro As Long
Dim lComando As Long
Dim tTRPAportePagtoCond As typeTRPAportePagtoCond

On Error GoTo Erro_TRPAportePagtoCond_Le_NumIntDocDestino

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 192341

    'Le a tabela TRPAportePagtoCond
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntDocAporte, Base, Status, DataPagto, NumIntDocDestino, " & _
                "FormaPagto, TipoDocDestino, Percentual, Valor FROM TRPAportePagtoCond WHERE NumIntDocDestino = ? AND TipoDocDestino = ?  ", _
                tTRPAportePagtoCond.lNumIntDoc, tTRPAportePagtoCond.lNumIntDocAporte, tTRPAportePagtoCond.iBase, tTRPAportePagtoCond.iStatus, _
                tTRPAportePagtoCond.dtDataPagto, tTRPAportePagtoCond.lNumIntDocDestino, tTRPAportePagtoCond.iFormaPagto, tTRPAportePagtoCond.iTipoDocDestino, tTRPAportePagtoCond.dPercentual, tTRPAportePagtoCond.dValor, _
                objTRPAportePagto.lNumIntDocDestino, objTRPAportePagto.iTipoDocDestino)
    If lErro <> AD_SQL_SUCESSO Then gError 192342

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192343

    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS
    
    objTRPAportePagto.lNumIntDoc = tTRPAportePagtoCond.lNumIntDoc
    objTRPAportePagto.lNumIntDocAporte = tTRPAportePagtoCond.lNumIntDocAporte
    objTRPAportePagto.iBase = tTRPAportePagtoCond.iBase
    objTRPAportePagto.iStatus = tTRPAportePagtoCond.iStatus
    objTRPAportePagto.dtDataPagto = tTRPAportePagtoCond.dtDataPagto
    objTRPAportePagto.lNumIntDocDestino = tTRPAportePagtoCond.lNumIntDocDestino
    objTRPAportePagto.iFormaPagto = tTRPAportePagtoCond.iFormaPagto
    objTRPAportePagto.iTipoDocDestino = tTRPAportePagtoCond.iTipoDocDestino
    objTRPAportePagto.dPercentual = tTRPAportePagtoCond.dPercentual
    objTRPAportePagto.dValor = tTRPAportePagtoCond.dValor
    
    Call Comando_Fechar(lComando)
    
    TRPAportePagtoCond_Le_NumIntDocDestino = SUCESSO

    Exit Function

Erro_TRPAportePagtoCond_Le_NumIntDocDestino:

    TRPAportePagtoCond_Le_NumIntDocDestino = gErr

    Select Case gErr
    
        Case 192341
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 192342, 192343
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPAPORTEPAGTOCOND", gErr)

        Case ERRO_LEITURA_SEM_DADOS
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192344)

    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function TRPAportePagtoDireto_Le_NumIntDocDestino(ByVal objTRPAportePagtoDireto As ClassTRPAportePagtoDireto) As Long

Dim lErro As Long
Dim tTRPAportePagtoDireto As typeTRPAportePagtoDireto
Dim lComando As Long

On Error GoTo Erro_TRPAportePagtoDireto_Le_NumIntDocDestino

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 192345
    
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntDocAporte, Valor, Vencimento, NumIntDocDestino, FormaPagto, " & _
                    "TipoDocDestino FROM TRPAportePagtoDireto WHERE NumIntDocDestino= ? AND TipoDocDestino = ? ", _
                    tTRPAportePagtoDireto.lNumIntDoc, tTRPAportePagtoDireto.lNumIntDocAporte, tTRPAportePagtoDireto.dValor, tTRPAportePagtoDireto.dtVencimento, _
                    tTRPAportePagtoDireto.lNumIntDocDestino, tTRPAportePagtoDireto.iFormaPagto, tTRPAportePagtoDireto.iTipoDocDestino, _
                    objTRPAportePagtoDireto.lNumIntDocDestino, objTRPAportePagtoDireto.iTipoDocDestino)
    If lErro <> AD_SQL_SUCESSO Then gError 192346

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192347
    
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objTRPAportePagtoDireto.lNumIntDoc = tTRPAportePagtoDireto.lNumIntDoc
    objTRPAportePagtoDireto.lNumIntDocAporte = tTRPAportePagtoDireto.lNumIntDocAporte
    objTRPAportePagtoDireto.dValor = tTRPAportePagtoDireto.dValor
    objTRPAportePagtoDireto.dtVencimento = tTRPAportePagtoDireto.dtVencimento
    objTRPAportePagtoDireto.lNumIntDocDestino = tTRPAportePagtoDireto.lNumIntDocDestino
    objTRPAportePagtoDireto.iFormaPagto = tTRPAportePagtoDireto.iFormaPagto
    objTRPAportePagtoDireto.iTipoDocDestino = tTRPAportePagtoDireto.iTipoDocDestino
    
    Call Comando_Fechar(lComando)

    TRPAportePagtoDireto_Le_NumIntDocDestino = SUCESSO

    Exit Function

Erro_TRPAportePagtoDireto_Le_NumIntDocDestino:

    TRPAportePagtoDireto_Le_NumIntDocDestino = gErr

    Select Case gErr
    
        Case 192345
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 192346, 192347
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPAPORTEPAGTODIRETO", gErr)
            
        Case ERRO_LEITURA_SEM_DADOS

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192348)

    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function TRPAportePagtoDireto_Le_NumIntDoc(ByVal objTRPAportePagtoDireto As ClassTRPAportePagtoDireto) As Long

Dim lErro As Long
Dim tTRPAportePagtoDireto As typeTRPAportePagtoDireto
Dim lComando As Long

On Error GoTo Erro_TRPAportePagtoDireto_Le_NumIntDoc

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 192364
    
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntDocAporte, Valor, Vencimento, NumIntDocDestino, FormaPagto, " & _
                    "TipoDocDestino FROM TRPAportePagtoDireto WHERE NumIntDoc = ? ", _
                    tTRPAportePagtoDireto.lNumIntDoc, tTRPAportePagtoDireto.lNumIntDocAporte, tTRPAportePagtoDireto.dValor, tTRPAportePagtoDireto.dtVencimento, _
                    tTRPAportePagtoDireto.lNumIntDocDestino, tTRPAportePagtoDireto.iFormaPagto, tTRPAportePagtoDireto.iTipoDocDestino, _
                    objTRPAportePagtoDireto.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 192365

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192366
    
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objTRPAportePagtoDireto.lNumIntDoc = tTRPAportePagtoDireto.lNumIntDoc
    objTRPAportePagtoDireto.lNumIntDocAporte = tTRPAportePagtoDireto.lNumIntDocAporte
    objTRPAportePagtoDireto.dValor = tTRPAportePagtoDireto.dValor
    objTRPAportePagtoDireto.dtVencimento = tTRPAportePagtoDireto.dtVencimento
    objTRPAportePagtoDireto.lNumIntDocDestino = tTRPAportePagtoDireto.lNumIntDocDestino
    objTRPAportePagtoDireto.iFormaPagto = tTRPAportePagtoDireto.iFormaPagto
    objTRPAportePagtoDireto.iTipoDocDestino = tTRPAportePagtoDireto.iTipoDocDestino
    
    Call Comando_Fechar(lComando)

    TRPAportePagtoDireto_Le_NumIntDoc = SUCESSO

    Exit Function

Erro_TRPAportePagtoDireto_Le_NumIntDoc:

    TRPAportePagtoDireto_Le_NumIntDoc = gErr

    Select Case gErr
    
        Case 192364
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 192365, 192366
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPAPORTEPAGTODIRETO", gErr)
            
        Case ERRO_LEITURA_SEM_DADOS

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192367)

    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function TRPVouchers_Le_Periodo(ByVal dtDataDe As Date, ByVal dtDataAte As Date, ByVal colVouchers As Collection) As Long

Dim lErro As Long
Dim lComando As Long
Dim tVou As typeTRPVouchers
Dim objVou As ClassTRPVouchers

On Error GoTo Erro_TRPVouchers_Le_Periodo

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 192699

    'Alocação de espaço no buffer
    tVou.sSerie = String(STRING_TRP_OCR_SERIE, 0)
    tVou.sTipVou = String(STRING_TRP_OCR_TIPOVOU, 0)
    tVou.sTitular = String(STRING_TRP_VOU_TITULAR, 0)
    tVou.sProduto = String(STRING_PRODUTO, 0)
    tVou.sControle = String(STRING_TRP_VOU_CONTROLE, 0)
    tVou.sCiaCart = String(STRING_TRP_VOU_CIACART, 0)
    tVou.sNumCCred = String(STRING_TRP_VOU_NUMCCRED, 0)
    tVou.sPassageiroNome = String(STRING_TRP_VOU_PAXNOME, 0)
    tVou.sPassageiroSobreNome = String(STRING_TRP_VOU_PAXNOME, 0)
    tVou.sUsuarioCanc = String(STRING_USUARIO_CODIGO, 0)
    tVou.sCiaaerea = String(STRING_TRP_VOU_CIAAEREA, 0)
    tVou.sAeroportos = String(STRING_TRP_VOU_AEROPORTOS, 0)
    tVou.sTitularCPF = String(STRING_CGC, 0)
    tVou.sPassageiroCGC = String(STRING_CGC, 0)

    'Le a tabelaTRPVouchers
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumVou, Serie, ClienteFat, Valor, Data, " & _
                "TipVou, NumFat, TipoDocDestino, NumIntDocDestino, Cartao, Status, NumIntDocNVL, " & _
                "ValorAporte, ValorComissao, Titular, Produto, Pax, Moeda, ValorCambio, " & _
                "Cambio, Controle, CiaCart, NumCCred, NumAuto, QuantParc, DiasAntc, " & _
                "Kit, ValorOcr, TemOcr, Representante, ComissaoRep, Correntista, ComissaoCorr, " & _
                "Emissor, ComissaoEmissor, ComissaoAg, ValorBruto, PassageiroNome, PassageiroSobrenome, ClienteVou, " & _
                "CliPassageiro, DataCanc, DataVigenciaDe, DataVigenciaAte, HoraCanc, UsuarioCanc, ClienteComissao, " & _
                "ciaaerea, Aeroportos, EnderecoPaxTitular, TemQueContabilizar, Promotor, Destino, TitularCPF, " & _
                "TarifaUnitaria, Vigencia, TarifaUnitariaFolheto, PassageiroCGC, PassageiroDataNasc FROM TRPVouchers WHERE Data BETWEEN  ? AND  ? ", _
                tVou.lNumIntDoc, tVou.lNumVou, tVou.sSerie, tVou.lClienteFat, _
                tVou.dValor, tVou.dtData, tVou.sTipVou, tVou.lNumFat, tVou.iTipoDocDestino, _
                tVou.lNumIntDocDestino, tVou.iCartao, tVou.iStatus, tVou.lNumIntDocNVL, tVou.dValorAporte, _
                tVou.dValorComissao, tVou.sTitular, tVou.sProduto, tVou.iPax, tVou.iMoeda, _
                tVou.dValorCambio, tVou.dCambio, tVou.sControle, tVou.sCiaCart, tVou.sNumCCred, _
                tVou.lNumAuto, tVou.iQuantParc, tVou.iDiasAntc, tVou.iKit, tVou.dValorOcr, _
                tVou.iTemOcr, tVou.lRepresentante, tVou.dComissaoRep, tVou.lCorrentista, tVou.dComissaoCorr, _
                tVou.lEmissor, tVou.dComissaoEmissor, tVou.dComissaoAg, tVou.dValorBruto, tVou.sPassageiroNome, _
                tVou.sPassageiroSobreNome, tVou.lClienteVou, tVou.lCliPassageiro, tVou.dtDataCanc, tVou.dtDataVigenciaDe, _
                tVou.dtDataVigenciaAte, tVou.dHoraCanc, tVou.sUsuarioCanc, tVou.lClienteComissao, tVou.sCiaaerea, _
                tVou.sAeroportos, tVou.lEnderecoPaxTitular, tVou.iTemQueContabilizar, tVou.iPromotor, tVou.iDestino, _
                tVou.sTitularCPF, tVou.dTarifaUnitaria, tVou.iVigencia, tVou.dTarifaUnitariaFolheto, tVou.sPassageiroCGC, tVou.dtPassageiroDataNasc, _
                dtDataDe, dtDataAte)
    If lErro <> AD_SQL_SUCESSO Then gError 192700

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192701

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objVou = New ClassTRPVouchers

        objVou.lNumVou = tVou.lNumVou
        objVou.lNumIntDoc = tVou.lNumIntDoc
        objVou.lNumVou = tVou.lNumVou
        objVou.sSerie = tVou.sSerie
        objVou.lClienteFat = tVou.lClienteFat
        objVou.dValor = tVou.dValor
        objVou.dtData = tVou.dtData
        objVou.sTipVou = tVou.sTipVou
        objVou.lNumFat = tVou.lNumFat
        objVou.iTipoDocDestino = tVou.iTipoDocDestino
        objVou.lNumIntDocDestino = tVou.lNumIntDocDestino
        objVou.iCartao = tVou.iCartao
        objVou.iStatus = tVou.iStatus
        objVou.lNumIntDocNVL = tVou.lNumIntDocNVL
        objVou.dValorAporte = tVou.dValorAporte
        objVou.dValorComissao = tVou.dValorComissao
        objVou.sTitular = tVou.sTitular
        objVou.sProduto = tVou.sProduto
        objVou.iPax = tVou.iPax
        objVou.iMoeda = tVou.iMoeda
        objVou.dValorCambio = tVou.dValorCambio
        objVou.dCambio = tVou.dCambio
        objVou.sControle = tVou.sControle
        objVou.sCiaCart = tVou.sCiaCart
        objVou.sNumCCred = tVou.sNumCCred
        objVou.lNumAuto = tVou.lNumAuto
        objVou.iQuantParc = tVou.iQuantParc
        objVou.iDiasAntc = tVou.iDiasAntc
        objVou.iKit = tVou.iKit
        objVou.dValorOcr = tVou.dValorOcr
        objVou.iTemOcr = tVou.iTemOcr
        objVou.lRepresentante = tVou.lRepresentante
        objVou.dComissaoRep = tVou.dComissaoRep
        objVou.lCorrentista = tVou.lCorrentista
        objVou.dComissaoCorr = tVou.dComissaoCorr
        objVou.lEmissor = tVou.lEmissor
        objVou.dComissaoEmissor = tVou.dComissaoEmissor
        objVou.dComissaoAg = tVou.dComissaoAg
        objVou.dValorBruto = tVou.dValorBruto
        objVou.sPassageiroNome = tVou.sPassageiroNome
        objVou.sPassageiroSobreNome = tVou.sPassageiroSobreNome
        objVou.lClienteVou = tVou.lClienteVou
        objVou.lCliPassageiro = tVou.lCliPassageiro
        objVou.dtDataCanc = tVou.dtDataCanc
        objVou.dtDataVigenciaDe = tVou.dtDataVigenciaDe
        objVou.dtDataVigenciaAte = tVou.dtDataVigenciaAte
        objVou.dHoraCanc = tVou.dHoraCanc
        objVou.sUsuarioCanc = tVou.sUsuarioCanc
        objVou.lClienteComissao = tVou.lClienteComissao
        objVou.sCiaaerea = tVou.sCiaaerea
        objVou.sAeroportos = tVou.sAeroportos
        objVou.lEnderecoPaxTitular = tVou.lEnderecoPaxTitular
        objVou.iTemQueContabilizar = tVou.iTemQueContabilizar
        objVou.iPromotor = tVou.iPromotor
        objVou.iDestino = tVou.iDestino
        objVou.sTitularCPF = tVou.sTitularCPF
        objVou.dTarifaUnitaria = tVou.dTarifaUnitaria
        objVou.iVigencia = tVou.iVigencia
        objVou.dTarifaUnitariaFolheto = tVou.dTarifaUnitariaFolheto
        objVou.sPassageiroCGC = tVou.sPassageiroCGC
        objVou.dtPassageiroDataNasc = tVou.dtPassageiroDataNasc
        
        colVouchers.Add objVou
        
        'Busca Proximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192702
        
    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    TRPVouchers_Le_Periodo = SUCESSO

    Exit Function

Erro_TRPVouchers_Le_Periodo:

    TRPVouchers_Le_Periodo = gErr

    Select Case gErr

        Case 192699
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 192700, 192701, 192702
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPVOUCHERS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192703)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function TRPVouchers_Le_GerComi(ByVal colVouchers As Collection) As Long

Dim lErro As Long
Dim lComando As Long
Dim tVou As typeTRPVouchers
Dim objVou As ClassTRPVouchers

On Error GoTo Erro_TRPVouchers_Le_GerComi

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 192699

    'Alocação de espaço no buffer
    tVou.sSerie = String(STRING_TRP_OCR_SERIE, 0)
    tVou.sTipVou = String(STRING_TRP_OCR_TIPOVOU, 0)

    'Le a tabelaTRPVouchers
    lErro = Comando_Executar(lComando, "SELECT NumVou, Serie, TipVou FROM TRPVouchers WHERE GeraComissao = ? ", _
                tVou.lNumVou, tVou.sSerie, tVou.sTipVou, MARCADO)
    If lErro <> AD_SQL_SUCESSO Then gError 192700

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192701

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objVou = New ClassTRPVouchers

        objVou.lNumVou = tVou.lNumVou
        objVou.sSerie = tVou.sSerie
        objVou.sTipVou = tVou.sTipVou
        
        colVouchers.Add objVou
        
        'Busca Proximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192702
        
    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    TRPVouchers_Le_GerComi = SUCESSO

    Exit Function

Erro_TRPVouchers_Le_GerComi:

    TRPVouchers_Le_GerComi = gErr

    Select Case gErr

        Case 192699
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 192700, 192701, 192702
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPVOUCHERS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192703)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Private Function TRP_Cliente_Tipo_Faturamento(ByVal lCliente As Long, bFaturarIndividualmente As Boolean) As Long

Dim lErro As Long
Dim objFilialCliente As New ClassFilialCliente
Dim colCategoriaItem As New Collection
Dim iIndice As Integer

On Error GoTo Erro_TRP_Cliente_Tipo_Faturamento

    bFaturarIndividualmente = False

    'Lê as categoria deste cliente
    objFilialCliente.lCodCliente = lCliente
    objFilialCliente.iCodFilial = FILIAL_MATRIZ
    
    lErro = CF("FilialCliente_Le_Categorias", objFilialCliente, colCategoriaItem)
    If lErro <> SUCESSO Then gError 192726
    
    If colCategoriaItem.Count > 0 Then
        For iIndice = 1 To colCategoriaItem.Count
            If UCase(colCategoriaItem.Item(iIndice).sCategoria) = TRP_CATEGORIA_FATURAMENTO Then
                If UCase(colCategoriaItem.Item(iIndice).sItem) = TRP_CATEGORIA_FATURAMENTO_POR_VOUCHER Then
                    bFaturarIndividualmente = True
                End If
                Exit For
            End If
        Next
    End If
 
    TRP_Cliente_Tipo_Faturamento = SUCESSO

    Exit Function

Erro_TRP_Cliente_Tipo_Faturamento:

    TRP_Cliente_Tipo_Faturamento = gErr

    Select Case gErr
    
        Case 192726

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192727)

    End Select


    Exit Function

End Function

Private Function TRP_GeracaoNF_Obtem_Cliente(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal objGeracaoNF As ClassTRPGeracaoNF, ByVal objTitRec As ClassTituloReceber, lCliente As Long, iFilialCliente As Integer) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim bAoPassageiro As Boolean
Dim lNumVou As Long
Dim sTipo As String
Dim sSerie As String
Dim objVoucher As New ClassTRPVouchers
Dim objTitRecTRP As ClassTitulosRecTRP

On Error GoTo Erro_TRP_GeracaoNF_Obtem_Cliente

    Set objTitRecTRP = objTitRec.objInfoUsu
    
    lCliente = objTitRec.lCliente
    iFilialCliente = objTitRec.iFilial

    If objTitRecTRP.iNFAoPassageiro = MARCADO Then
        bAoPassageiro = True
    Else
        bAoPassageiro = False
    End If
    
    'Se o cliente da NF deve ser o passageiro e o mesmo não está cadastrado, tem
    'que obter as informações dele e fazer o cadastro
    If bAoPassageiro And objTitRecTRP.lCliPassageiro = 0 Then
    
        'TEM QUE OBTER UM DOS VOUCHERS
        sTipo = String(STRING_TRP_OCR_TIPOVOU, 0)
        sSerie = String(STRING_TRP_OCR_SERIE, 0)
        
        'Le a tabelaTRPVouchers
        lErro = Comando_Executar(lComando1, "SELECT NumVou, TipVou, Serie FROM DocFaturados WHERE TipoDocDestino = ? AND NumIntDocDestino = ? ", lNumVou, sTipo, sSerie, TRP_TIPO_DOC_DESTINO_TITREC, objTitRec.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 192732
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192733
    
        If lErro <> SUCESSO Then gError 192734
        
        objVoucher.sSerie = sSerie
        objVoucher.sTipVou = sTipo
        objVoucher.lNumVou = lNumVou
        
        lErro = CF("TRPVouchers_Le", objVoucher)
        If lErro <> SUCESSO Then gError 200275
        
        'VERIFICAR SE O PASSAGEIRO ESTÁ PREENCHIDO
        If objVoucher.lCliPassageiro = 0 Then
        
            'CADASTRAR O PASSAGEIRO E ACERTA A INFORMAÇÃO EM VOUCHERINFO
            lErro = CF("TRP_Voucher_Cadastra_Cliente", objVoucher)
            If lErro <> SUCESSO Then gError 192738

        End If
        
        objTitRecTRP.lCliPassageiro = objVoucher.lCliPassageiro
    
    End If
    
    If bAoPassageiro Then
        lCliente = objTitRecTRP.lCliPassageiro
        iFilialCliente = FILIAL_MATRIZ
    End If
 
    TRP_GeracaoNF_Obtem_Cliente = SUCESSO

    Exit Function

Erro_TRP_GeracaoNF_Obtem_Cliente:

    TRP_GeracaoNF_Obtem_Cliente = gErr

    Select Case gErr

        Case 192731, 192735 To 192738, 192728, 200275
        
        Case 192732, 192733
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPVOUCHERS", gErr)
            
        Case 192734
            Call Rotina_Erro(vbOKOnly, "ERRO_VOUCHER_NAO_CADASTRADO2", gErr, objTitRec.lNumTitulo)
       
        Case 192729
            Call Rotina_Erro(vbOKOnly, "ERRO_SENHA_NAO_PREENCHIDA", gErr)
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192739)

    End Select
    
    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function TitulosRecTRP_Le(ByVal objTitulosRecTRP As ClassTitulosRecTRP) As Long

Dim lErro As Long
Dim lComando As Long
Dim tTitulosRecTRP As typeTitulosRecTRP

On Error GoTo Erro_TitulosRecTRP_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 192760

    'Alocação de espaço no buffer

    'Le a tabelaTitulosRecTRP
    lErro = Comando_Executar(lComando, "SELECT NumIntDocTitRec, ValorTarifa,  ValorDeducoes, ValorComissao, ValorBruto FROM TitulosRecTRP WHERE NumIntDocTitRec= ? ", _
                tTitulosRecTRP.lNumIntDocTitRec, tTitulosRecTRP.dValorTarifa, tTitulosRecTRP.dValorDeducoes, _
                tTitulosRecTRP.dValorComissao, tTitulosRecTRP.dValorBruto, _
                objTitulosRecTRP.lNumIntDocTitRec)
    If lErro <> AD_SQL_SUCESSO Then gError 192761

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192762

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objTitulosRecTRP.lNumIntDocTitRec = tTitulosRecTRP.lNumIntDocTitRec
    objTitulosRecTRP.dValorTarifa = tTitulosRecTRP.dValorTarifa
    objTitulosRecTRP.dValorDeducoes = tTitulosRecTRP.dValorDeducoes
    objTitulosRecTRP.dValorComissao = tTitulosRecTRP.dValorComissao
    objTitulosRecTRP.dValorBruto = tTitulosRecTRP.dValorBruto

    'Fecha Comando
    Call Comando_Fechar(lComando)

    TitulosRecTRP_Le = SUCESSO

    Exit Function

Erro_TitulosRecTRP_Le:

    TitulosRecTRP_Le = gErr

    Select Case gErr

        Case 192760
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 192761, 192762
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOSRECTRP", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192763)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function Endereco_Le_Customizado(ByVal objEndereco As ClassEndereco) As Long

Dim lErro As Long
Dim lComando As Long
Dim sEmail2 As String
Dim objEnderecoTRP As New ClassEnderecoTRP

On Error GoTo Erro_Endereco_Le_Customizado

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 196211

    'Alocação de espaço no buffer
    sEmail2 = String(STRING_EMAIL, 0)

    'Le a tabelaTRPOcorrencias
    lErro = Comando_Executar(lComando, "SELECT Email2 FROM Enderecos WHERE Codigo= ? ", sEmail2, objEndereco.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 196212

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196213

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS
    
    objEnderecoTRP.sEmail2 = sEmail2

    Set objEndereco.objInfoUsu = objEnderecoTRP

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Endereco_Le_Customizado = SUCESSO

    Exit Function

Erro_Endereco_Le_Customizado:

    Endereco_Le_Customizado = gErr

    Select Case gErr

        Case 196211
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 196212, 196213
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ENDERECOS", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196214)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function Endereco_Trata_Customizacao(ByVal objEndereco As ClassEndereco) As Long

Dim lErro As Long

On Error GoTo Erro_Endereco_Trata_Customizacao

    lErro = CF("Endereco_Le_Customizado", objEndereco)
    If lErro <> SUCESSO Then gError 196233
    
    If Len(Trim(objEndereco.objInfoUsu.sEmail2)) <> 0 Then
        If Len(Trim(objEndereco.sEmail)) > 0 Then
            objEndereco.sEmail = objEndereco.sEmail & ";" & objEndereco.objInfoUsu.sEmail2
        Else
            objEndereco.sEmail = objEndereco.objInfoUsu.sEmail2
        End If
    End If
    
    Endereco_Trata_Customizacao = SUCESSO

    Exit Function

Erro_Endereco_Trata_Customizacao:

    Endereco_Trata_Customizacao = gErr

    Select Case gErr
    
        Case 196233

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196234)

    End Select

    Exit Function
    
End Function

Public Function Cliente_Le_Customizado(ByVal objCliente As ClassCliente, Optional ByVal bLeTudo As Boolean = False, Optional ByVal lComandoP1 As Long, Optional ByVal lComandoP2 As Long, Optional ByVal lComandoP3 As Long) As Long

Dim lErro As Long
Dim tCliEmissoresTRP As typeCliEmissoresTRP
Dim tCliEmissoresExcTRP As typeCliEmissoresExcTRP
Dim objCliEmissoresTRP As ClassCliEmissoresTRP
Dim objCliEmissoresExcTRP As ClassCliEmissoresExcTRP
Dim objClienteTRP As New ClassClienteTRP
Dim alComando(1 To 3) As Long
Dim iIndice As Integer
Dim tFiliaisClientesTRP As typeFiliaisClientesTRP
Dim objForn As New ClassFornecedor

On Error GoTo Erro_Cliente_Le_Customizado

    If lComandoP1 <> 0 And lComandoP2 <> 0 And lComandoP3 <> 0 Then
        alComando(1) = lComandoP1
        alComando(2) = lComandoP2
        alComando(3) = lComandoP3
    Else
        'Abre o comando
        For iIndice = LBound(alComando) To UBound(alComando)
            alComando(iIndice) = Comando_Abrir()
            If alComando(iIndice) = 0 Then gError 195774
        Next
    End If
    
    If bLeTudo Then

        'Alocação de espaço no buffer
    
        'Le a tabela TRPCliEmissores
        lErro = Comando_Executar(alComando(1), "SELECT NumIntDoc, Cliente, Seq, FornEmissor, PercComissao, Codigo FROM TRPCliEmissores WHERE Cliente= ? ", _
                    tCliEmissoresTRP.lNumIntDoc, tCliEmissoresTRP.lCliente, tCliEmissoresTRP.iSeq, tCliEmissoresTRP.lFornEmissor, tCliEmissoresTRP.dPercComissao, tCliEmissoresTRP.lCodigo, _
                    objCliente.lCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 195775
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 195776
    
        Do While lErro = AD_SQL_SUCESSO
    
            Set objCliEmissoresTRP = New ClassCliEmissoresTRP
    
            objCliEmissoresTRP.lNumIntDoc = tCliEmissoresTRP.lNumIntDoc
            objCliEmissoresTRP.lCliente = tCliEmissoresTRP.lCliente
            objCliEmissoresTRP.iSeq = tCliEmissoresTRP.iSeq
            objCliEmissoresTRP.lFornEmissor = tCliEmissoresTRP.lFornEmissor
            objCliEmissoresTRP.dPercComissao = tCliEmissoresTRP.dPercComissao
            objCliEmissoresTRP.lCodigo = tCliEmissoresTRP.lCodigo
            
            objClienteTRP.colCliEmissoresTRP.Add objCliEmissoresTRP
            
            tCliEmissoresExcTRP.sProduto = String(STRING_PRODUTO, 0)
            
            'Le a tabela CliEmissoresExcTRP
            lErro = Comando_Executar(alComando(2), "SELECT NumIntDoc, NumIntDocCliEmi, Seq, Produto, PercComissao FROM TRPCliEmissoresExc WHERE NumIntDocCliEmi = ? ", _
                        tCliEmissoresExcTRP.lNumIntDoc, tCliEmissoresExcTRP.lNumIntDocCliEmi, tCliEmissoresExcTRP.iSeq, tCliEmissoresExcTRP.sProduto, tCliEmissoresExcTRP.dPercComissao, _
                        objCliEmissoresTRP.lNumIntDoc)
            If lErro <> AD_SQL_SUCESSO Then gError 195777
        
            'Busca Primeiro
            lErro = Comando_BuscarPrimeiro(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 195778
            
            Do While lErro = AD_SQL_SUCESSO
            
                Set objCliEmissoresExcTRP = New ClassCliEmissoresExcTRP
        
                objCliEmissoresExcTRP.lNumIntDoc = tCliEmissoresExcTRP.lNumIntDoc
                objCliEmissoresExcTRP.lNumIntDocCliEmi = tCliEmissoresExcTRP.lNumIntDocCliEmi
                objCliEmissoresExcTRP.iSeq = tCliEmissoresExcTRP.iSeq
                objCliEmissoresExcTRP.sProduto = tCliEmissoresExcTRP.sProduto
                objCliEmissoresExcTRP.dPercComissao = tCliEmissoresExcTRP.dPercComissao
            
                objCliEmissoresTRP.colCliEmissoresExcTRP.Add objCliEmissoresExcTRP
                
                lErro = Comando_BuscarProximo(alComando(2))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 195779
            
            Loop
            
            lErro = Comando_BuscarProximo(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 195780
            
        Loop
        
    End If
        
    tFiliaisClientesTRP.sGrupo = String(STRING_TRP_OCR_CODGRUPO, 0)
    
    'Le a tabela FiliaisClientesTRP
    lErro = Comando_Executar(alComando(3), "SELECT Representante, ComissaoRep, Correntista, ComissaoCorr, ComissaoAg, Aporte, EmpresaPai, FilialNF, FilialFAT, FilialEmpresa,  CondPagtoCC FROM FiliaisClientesTRP WHERE CodCliente= ? AND CodFilial = ?", _
                tFiliaisClientesTRP.lRepresentante, tFiliaisClientesTRP.dPercComiRep, tFiliaisClientesTRP.lCorrentista, tFiliaisClientesTRP.dPercComiCorr, tFiliaisClientesTRP.dPercComiAg, tFiliaisClientesTRP.iConsiderarAporte, _
                tFiliaisClientesTRP.lEmpresaPai, tFiliaisClientesTRP.iFilialNF, tFiliaisClientesTRP.iFilialFat, tFiliaisClientesTRP.iFilialEmpresa, tFiliaisClientesTRP.iCondPagtoCC, objCliente.lCodigo, FILIAL_MATRIZ)
    If lErro <> AD_SQL_SUCESSO Then gError 195861

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 195862
    
    objClienteTRP.lCodRepresentante = tFiliaisClientesTRP.lRepresentante
    objClienteTRP.dPercComiRep = tFiliaisClientesTRP.dPercComiRep
    objClienteTRP.lCodCorrentista = tFiliaisClientesTRP.lCorrentista
    objClienteTRP.dPercComiCorr = tFiliaisClientesTRP.dPercComiCorr
    objClienteTRP.dPercComiAg = tFiliaisClientesTRP.dPercComiAg
    objClienteTRP.iConsiderarAporte = tFiliaisClientesTRP.iConsiderarAporte
    objClienteTRP.lEmpresaPai = tFiliaisClientesTRP.lEmpresaPai
    objClienteTRP.iFilialNF = tFiliaisClientesTRP.iFilialNF
    objClienteTRP.iFilialFat = tFiliaisClientesTRP.iFilialFat
    objClienteTRP.iFilialEmpresa = tFiliaisClientesTRP.iFilialEmpresa
    objClienteTRP.iCondPagtoCC = tFiliaisClientesTRP.iCondPagtoCC
    
    Set objCliente.objInfoUsu = objClienteTRP
    
    If bLeTudo Then
        
        lErro = CF("TRPClienteCorProd_Le", objCliente)
        If lErro <> SUCESSO Then gError 196953
        
        lErro = CF("TRPClienteRepProd_Le", objCliente)
        If lErro <> SUCESSO Then gError 196954
        
        lErro = CF("TRPExcComissaoCli_Le", objCliente)
        If lErro <> SUCESSO Then gError 196954
        
        objForn.lCodigo = objCliente.lCodigo
    
        lErro = CF("Fornecedor_Le", objForn)
        If lErro <> SUCESSO And lErro <> 12729 Then gError 196954
        
        objClienteTRP.iBanco = objForn.iBanco
        objClienteTRP.sAgencia = objForn.sAgencia
        objClienteTRP.sContaCorrente = objForn.sContaCorrente
    
    End If
    
    objClienteTRP.objUsuarioWeb.iTipoUsuario = TRP_USUARIO_CLIENTE
    objClienteTRP.objUsuarioWeb.lCodigo = objCliente.lCodigo
    
    lErro = CF("TRPUsuarios_Le", objClienteTRP.objUsuarioWeb)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 196954

    If lComandoP1 = 0 Or lComandoP2 = 0 Or lComandoP3 = 0 Then
        'Fecha Comando
        For iIndice = LBound(alComando) To UBound(alComando)
            Call Comando_Fechar(alComando(iIndice))
        Next
    End If

    Cliente_Le_Customizado = SUCESSO

    Exit Function

Erro_Cliente_Le_Customizado:

    Cliente_Le_Customizado = gErr

    Select Case gErr

        Case 195774
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 195775, 195776, 195780
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPCLIEMISSORES", gErr)

        Case 195777, 195778, 195779
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPCLIEMISSORESEXC", gErr)
        
        Case 195861, 195862
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISCLIENTESTRP", gErr)
            
        Case 196953, 196954
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195781)

    End Select

    If lComandoP1 = 0 Or lComandoP2 = 0 Or lComandoP3 = 0 Then
        'Fecha Comando
        For iIndice = LBound(alComando) To UBound(alComando)
            Call Comando_Fechar(alComando(iIndice))
        Next
    End If

    Exit Function

End Function

Public Function TRPAcordo_Le(ByVal objTRPAcordo As ClassTRPAcordos) As Long

Dim lErro As Long
Dim tTRPAcordo As typeTRPAcordos
Dim tTRPAcordoComiss As typeTRPAcordoComissao
Dim tTaf As typeTRPAcordoTarifaDif
Dim objTRPAcordoComiss As ClassTRPAcordoComissao
Dim objTaf As ClassTRPAcordoTarifaDif
Dim alComando(1 To 3) As Long
Dim iIndice As Integer

On Error GoTo Erro_TRPAcordo_Le

    'Abre o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 197145
    Next

    tTRPAcordo.sContrato = String(STRING_TRP_ACORDO_CONTRATO, 0)
    tTRPAcordo.sObservacao = String(STRING_TRP_ACORDO_OBS, 0)
    tTRPAcordo.sDescricao = String(STRING_TRP_ACORDO_DESC, 0)

    'Le a tabela TRPAcordo
    lErro = Comando_Executar(alComando(1), "SELECT NumIntDoc, Contrato, Cliente, FilialCliente, ValidadeDe, ValidadeAte, " & _
                              "Observacao, Descricao FROM TRPAcordos WHERE Numero = ? ", _
                tTRPAcordo.lNumIntDoc, tTRPAcordo.sContrato, tTRPAcordo.lCliente, tTRPAcordo.iFilialCliente, tTRPAcordo.dtValidadeDe, _
                tTRPAcordo.dtValidadeAte, tTRPAcordo.sObservacao, tTRPAcordo.sDescricao, objTRPAcordo.lNumero)
    If lErro <> AD_SQL_SUCESSO Then gError 197146

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197147

    If lErro <> AD_SQL_SUCESSO Then gError 197148

    objTRPAcordo.lNumIntDoc = tTRPAcordo.lNumIntDoc
    objTRPAcordo.sContrato = tTRPAcordo.sContrato
    objTRPAcordo.lCliente = tTRPAcordo.lCliente
    objTRPAcordo.iFilialCliente = tTRPAcordo.iFilialCliente
    objTRPAcordo.dtValidadeDe = tTRPAcordo.dtValidadeDe
    objTRPAcordo.dtValidadeAte = tTRPAcordo.dtValidadeAte
    objTRPAcordo.sObservacao = tTRPAcordo.sObservacao
    objTRPAcordo.sDescricao = tTRPAcordo.sDescricao

    tTRPAcordoComiss.sProduto = String(STRING_PRODUTO, 0)

    lErro = Comando_Executar(alComando(2), "SELECT NumIntDoc, Seq, Produto, Destino, PercComissao FROM TRPAcordoComissao WHERE NumIntAcordo = ? ORDER BY Seq", _
                                           tTRPAcordoComiss.lNumIntDoc, tTRPAcordoComiss.iSeq, tTRPAcordoComiss.sProduto, tTRPAcordoComiss.iDestino, tTRPAcordoComiss.dPercComissao, objTRPAcordo.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 197149

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197150

    Do While lErro = AD_SQL_SUCESSO
    
        Set objTRPAcordoComiss = New ClassTRPAcordoComissao
    
        objTRPAcordoComiss.lNumIntDoc = tTRPAcordoComiss.lNumIntDoc
        objTRPAcordoComiss.lNumIntAcordo = objTRPAcordo.lNumIntDoc
        objTRPAcordoComiss.iSeq = tTRPAcordoComiss.iSeq
        objTRPAcordoComiss.sProduto = tTRPAcordoComiss.sProduto
        objTRPAcordoComiss.iDestino = tTRPAcordoComiss.iDestino
        objTRPAcordoComiss.dPercComissao = tTRPAcordoComiss.dPercComissao
    
        objTRPAcordo.colTRPAcordoComiss.Add objTRPAcordoComiss
    
        lErro = Comando_Executar(alComando(3), "SELECT NumIntDoc,NumIntAcordoComis,DiasDe,DiasAte,Diario,Valor,ValorAdicional FROM TRPAcordoTarifaDif WHERE NumIntAcordoComis = ? ORDER BY DiasDe", _
                                               tTaf.lNumIntDoc, tTaf.lNumIntAcordoComis, tTaf.iDiasDe, tTaf.iDiasAte, tTaf.iDiario, tTaf.dValor, tTaf.dValorAdicional, objTRPAcordoComiss.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 197149
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197150
    
        Do While lErro = AD_SQL_SUCESSO
        
            Set objTaf = New ClassTRPAcordoTarifaDif
    
            objTaf.dValor = tTaf.dValor
            objTaf.dValorAdicional = tTaf.dValorAdicional
            objTaf.iDiario = tTaf.iDiario
            objTaf.iDiasAte = tTaf.iDiasAte
            objTaf.iDiasDe = tTaf.iDiasDe
            objTaf.lNumIntAcordoComis = tTaf.lNumIntAcordoComis
            objTaf.lNumIntDoc = tTaf.lNumIntDoc
    
            objTRPAcordoComiss.colTRPAcordoDif.Add objTaf
    
            lErro = Comando_BuscarProximo(alComando(3))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197151
    
        Loop
    
        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197151
    
    Loop

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TRPAcordo_Le = SUCESSO

    Exit Function

Erro_TRPAcordo_Le:

    TRPAcordo_Le = gErr

    Select Case gErr

        Case 197145
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 197146, 197147
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPACORDOS", gErr)

        Case 197148

        Case 197149, 197150, 197151
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPACORDOCOMISS", gErr)
        
        Case 197152, 197153, 197154
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPACORDODIF", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197155)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function TRPVoucherInfo_Le(ByVal objTRPVoucher As ClassTRPVouchers, Optional ByVal lComandoP As Long = 0) As Long

Dim lErro As Long
Dim lComando As Long
Dim tTRPVoucherInfo As typeTRPVoucherInfo
Dim objTRPVoucherInfo As ClassTRPVoucherInfo

On Error GoTo Erro_TRPVoucherInfo_Le

    'Executa a abertura do Comando
    If lComandoP = 0 Then
        lComando = Comando_Abrir()
        If lComando = 0 Then gError 197207
    Else
        lComando = lComandoP
    End If

    'Alocação de espaço no buffer
    tTRPVoucherInfo.sTipoDoc = String(STRING_TRP_VOU_INFO_TIPODOC, 0)
    tTRPVoucherInfo.sTipVou = String(STRING_TRP_VOU_INFO_TIPVOU, 0)
    tTRPVoucherInfo.sSerie = String(STRING_TRP_VOU_INFO_SERIE, 0)
    tTRPVoucherInfo.sHistorico = String(STRING_TRP_VOU_HISTORICO, 0)

    'Le a tabelaTRPVoucherInfo
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, TipoDoc, TipVou, Serie, NumVou, Seq, " & _
                "Data, TipoDocDestino, NumIntDocDestino, NumIntDocComiInt, Valor, Historico, NumTitulo, Status,TipoLiberacao , Manual, TipoCliForn, Estorno, NumIntDocLiberacao, CliForn, NumIntDocOCR,NumIntDocEstorno FROM TRPVoucherInfo WHERE TipVou= ?  AND Serie= ?  AND NumVou= ? ORDER BY TipoDoc, Seq", _
                tTRPVoucherInfo.lNumIntDoc, tTRPVoucherInfo.sTipoDoc, tTRPVoucherInfo.sTipVou, tTRPVoucherInfo.sSerie, _
                tTRPVoucherInfo.lNumVou, tTRPVoucherInfo.iSeq, tTRPVoucherInfo.dtData, tTRPVoucherInfo.iTipoDocDestino, tTRPVoucherInfo.lNumIntDocDestino, _
                tTRPVoucherInfo.lNumIntDocComiInt, tTRPVoucherInfo.dValor, tTRPVoucherInfo.sHistorico, tTRPVoucherInfo.lNumTitulo, tTRPVoucherInfo.iStatus, _
                tTRPVoucherInfo.iTipoLiberacao, tTRPVoucherInfo.iManual, tTRPVoucherInfo.iTipoCliForn, tTRPVoucherInfo.iEstorno, tTRPVoucherInfo.lNumIntDocLiberacao, _
                tTRPVoucherInfo.lCliForn, tTRPVoucherInfo.lNumIntDocOCR, tTRPVoucherInfo.lNumIntDocEstorno, _
                objTRPVoucher.sTipVou, objTRPVoucher.sSerie, objTRPVoucher.lNumVou)
    If lErro <> AD_SQL_SUCESSO Then gError 197208

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197209

    Do While lErro = AD_SQL_SUCESSO

        Set objTRPVoucherInfo = New ClassTRPVoucherInfo

        objTRPVoucherInfo.lNumIntDoc = tTRPVoucherInfo.lNumIntDoc
        objTRPVoucherInfo.sTipoDoc = tTRPVoucherInfo.sTipoDoc
        objTRPVoucherInfo.sTipVou = tTRPVoucherInfo.sTipVou
        objTRPVoucherInfo.sSerie = tTRPVoucherInfo.sSerie
        objTRPVoucherInfo.lNumVou = tTRPVoucherInfo.lNumVou
        objTRPVoucherInfo.iSeq = tTRPVoucherInfo.iSeq
        objTRPVoucherInfo.dtData = tTRPVoucherInfo.dtData
        objTRPVoucherInfo.iTipoDocDestino = tTRPVoucherInfo.iTipoDocDestino
        objTRPVoucherInfo.lNumIntDocDestino = tTRPVoucherInfo.lNumIntDocDestino
        objTRPVoucherInfo.lNumIntDocComiInt = tTRPVoucherInfo.lNumIntDocComiInt
        objTRPVoucherInfo.dValor = tTRPVoucherInfo.dValor
        objTRPVoucherInfo.sHistorico = tTRPVoucherInfo.sHistorico
        objTRPVoucherInfo.lNumTitulo = tTRPVoucherInfo.lNumTitulo
        objTRPVoucherInfo.iStatus = tTRPVoucherInfo.iStatus
        objTRPVoucherInfo.iTipoLiberacao = tTRPVoucherInfo.iTipoLiberacao
        objTRPVoucherInfo.iManual = tTRPVoucherInfo.iManual
        objTRPVoucherInfo.iTipoCliForn = tTRPVoucherInfo.iTipoCliForn
        objTRPVoucherInfo.iEstorno = tTRPVoucherInfo.iEstorno
        objTRPVoucherInfo.lNumIntDocLiberacao = tTRPVoucherInfo.lNumIntDocLiberacao
        objTRPVoucherInfo.lCliForn = tTRPVoucherInfo.lCliForn
        objTRPVoucherInfo.lNumIntDocOCR = tTRPVoucherInfo.lNumIntDocOCR
        objTRPVoucherInfo.lNumIntDocEstorno = tTRPVoucherInfo.lNumIntDocEstorno

        objTRPVoucher.colTRPVoucherInfo.Add objTRPVoucherInfo

        'Busca Proximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197210

    Loop

    'Fecha Comando
    If lComandoP = 0 Then Call Comando_Fechar(lComando)

    TRPVoucherInfo_Le = SUCESSO

    Exit Function

Erro_TRPVoucherInfo_Le:

    TRPVoucherInfo_Le = gErr

    Select Case gErr

        Case 197207
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 197208 To 197210
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPVOUCHERINFO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197211)

    End Select

    'Fecha Comando
    If lComandoP = 0 Then Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function TRPVoucher_Le_ComisInt(ByVal objVou As ClassTRPVouchers) As Long

Dim lErro As Long
Dim lComando As Long
Dim objTRPGerComiIntDet As ClassTRPGerComiIntDet
Dim tTRPGerComiIntDet As typeTRPGerComiIntDet

On Error GoTo Erro_TRPVoucher_Le_ComisInt

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 197217

    'Alocação de espaço no buffer
    tTRPGerComiIntDet.sNomeReduzidoVendedor = String(STRING_VENDEDOR_NOME_REDUZIDO, 0)

    'Le a tabelaTRPVoucherInfo
    lErro = Comando_Executar(lComando, "SELECT V.Codigo, V.NomeReduzido, DV.ValorBase, DV.PercComi, G.DataGeracao FROM TRPGerComiIntDetVou AS DV, TRPGerComiInt AS G, TRPGerComiIntDet AS D, Vendedores AS V WHERE DV.NumIntComiIntDet = D.NumIntDoc AND G.NumIntDoc = D.NumIntDocGerComi AND D.Vendedor = V.Codigo AND DV.TipVou = ? AND DV.Serie = ? AND DV.NumVou = ? ORDER BY D.Vendedor, G.DataGeracao", _
               tTRPGerComiIntDet.iVendedor, tTRPGerComiIntDet.sNomeReduzidoVendedor, tTRPGerComiIntDet.dValorBase, tTRPGerComiIntDet.dPercComissao, tTRPGerComiIntDet.dtDataGeracao, objVou.sTipVou, objVou.sSerie, objVou.lNumVou)
    If lErro <> AD_SQL_SUCESSO Then gError 197218

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197219

    Do While lErro = AD_SQL_SUCESSO

        Set objTRPGerComiIntDet = New ClassTRPGerComiIntDet

        objTRPGerComiIntDet.iVendedor = tTRPGerComiIntDet.iVendedor
        objTRPGerComiIntDet.sNomeReduzidoVendedor = tTRPGerComiIntDet.sNomeReduzidoVendedor
        objTRPGerComiIntDet.dValorBase = tTRPGerComiIntDet.dValorBase
        objTRPGerComiIntDet.dValorComissao = Arredonda_Moeda(tTRPGerComiIntDet.dValorBase * tTRPGerComiIntDet.dPercComissao)
        objTRPGerComiIntDet.dPercComissao = tTRPGerComiIntDet.dPercComissao
        objTRPGerComiIntDet.dtDataGeracao = tTRPGerComiIntDet.dtDataGeracao

        objVou.colTRPGerComiIntDet.Add objTRPGerComiIntDet

        'Busca Proximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197220

    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    TRPVoucher_Le_ComisInt = SUCESSO

    Exit Function

Erro_TRPVoucher_Le_ComisInt:

    TRPVoucher_Le_ComisInt = gErr

    Select Case gErr

        Case 197217
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 197218 To 197220
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPGERCOMIINTDET", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197221)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Private Function TRPAportePagtoFatCond_Le(ByVal lComando As Long, ByVal objTRPAporte As ClassTRPAportes, Optional ByVal bLock As Boolean = False) As Long

Dim lErro As Long
Dim tTRPAportePagtoFatCond As typeTRPAportePagtoFatCond
Dim objTRPAportePagtoFatCond As ClassTRPAportePagtoFatC

On Error GoTo Erro_TRPAportePagtoFatCond_Le

    'Le a tabelaTRPAportePagtoFatCond
    If bLock Then
        lErro = Comando_ExecutarPos(lComando, "SELECT NumIntDoc, NumIntDocAporte, Valor, ValidadeDe, ValidadeAte,  " & _
                    "Percentual FROM TRPAportePagtoFatCond WHERE NumIntDocAporte= ? ORDER BY Seq ", 0, _
                    tTRPAportePagtoFatCond.lNumIntDoc, tTRPAportePagtoFatCond.lNumIntDocAporte, tTRPAportePagtoFatCond.dValor, tTRPAportePagtoFatCond.dtValidadeDe, _
                    tTRPAportePagtoFatCond.dtValidadeAte, tTRPAportePagtoFatCond.dPercentual, _
                    objTRPAporte.lNumIntDoc)
    Else
        lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntDocAporte, Valor, ValidadeDe, ValidadeAte,  " & _
                    "Percentual FROM TRPAportePagtoFatCond WHERE NumIntDocAporte= ? ORDER BY Seq ", _
                    tTRPAportePagtoFatCond.lNumIntDoc, tTRPAportePagtoFatCond.lNumIntDocAporte, tTRPAportePagtoFatCond.dValor, tTRPAportePagtoFatCond.dtValidadeDe, _
                    tTRPAportePagtoFatCond.dtValidadeAte, tTRPAportePagtoFatCond.dPercentual, _
                    objTRPAporte.lNumIntDoc)
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 190344

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190345

    Do While lErro <> AD_SQL_SEM_DADOS
    
        If bLock Then
            'Trava registro contra alterações/Leituras
            lErro = Comando_LockExclusive(lComando)
            If lErro <> SUCESSO Then gError 190381
        End If
    
        Set objTRPAportePagtoFatCond = New ClassTRPAportePagtoFatC

        objTRPAportePagtoFatCond.lNumIntDoc = tTRPAportePagtoFatCond.lNumIntDoc
        objTRPAportePagtoFatCond.lNumIntDocAporte = tTRPAportePagtoFatCond.lNumIntDocAporte
        objTRPAportePagtoFatCond.dValor = tTRPAportePagtoFatCond.dValor
        objTRPAportePagtoFatCond.dtValidadeDe = tTRPAportePagtoFatCond.dtValidadeDe
        objTRPAportePagtoFatCond.dtValidadeAte = tTRPAportePagtoFatCond.dtValidadeAte
        objTRPAportePagtoFatCond.dPercentual = tTRPAportePagtoFatCond.dPercentual
        
        objTRPAporte.colPagtoSobreFaturaCond.Add objTRPAportePagtoFatCond

        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190346

    Loop

    TRPAportePagtoFatCond_Le = SUCESSO

    Exit Function

Erro_TRPAportePagtoFatCond_Le:

    TRPAportePagtoFatCond_Le = gErr

    Select Case gErr

        Case 190344, 190345, 190346
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPAPORTEPAGTOFATCOND", gErr)

        Case 190381
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRPAPORTEPAGTOFATCOND", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190347)

    End Select

    Exit Function

End Function

Private Function TRPAportePagtoFatCond_Le_NumIntDoc(ByVal lComando As Long, ByVal objTRPAportePagtoFatCond As ClassTRPAportePagtoFatC) As Long

Dim lErro As Long
Dim tTRPAportePagtoFatCond As typeTRPAportePagtoFatCond

On Error GoTo Erro_TRPAportePagtoFatCond_Le_NumIntDoc

    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntDocAporte, Valor, ValidadeDe, ValidadeAte,  " & _
                    "Percentual FROM TRPAportePagtoFatCond WHERE NumIntDoc= ? ORDER BY Seq ", _
                    tTRPAportePagtoFatCond.lNumIntDoc, tTRPAportePagtoFatCond.lNumIntDocAporte, tTRPAportePagtoFatCond.dValor, tTRPAportePagtoFatCond.dtValidadeDe, _
                    tTRPAportePagtoFatCond.dtValidadeAte, tTRPAportePagtoFatCond.dPercentual, _
                    objTRPAportePagtoFatCond.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 190950

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190951

    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS
    
    objTRPAportePagtoFatCond.lNumIntDoc = tTRPAportePagtoFatCond.lNumIntDoc
    objTRPAportePagtoFatCond.lNumIntDocAporte = tTRPAportePagtoFatCond.lNumIntDocAporte
    objTRPAportePagtoFatCond.dValor = tTRPAportePagtoFatCond.dValor
    objTRPAportePagtoFatCond.dtValidadeDe = tTRPAportePagtoFatCond.dtValidadeDe
    objTRPAportePagtoFatCond.dtValidadeAte = tTRPAportePagtoFatCond.dtValidadeAte
    objTRPAportePagtoFatCond.dPercentual = tTRPAportePagtoFatCond.dPercentual

    TRPAportePagtoFatCond_Le_NumIntDoc = SUCESSO

    Exit Function

Erro_TRPAportePagtoFatCond_Le_NumIntDoc:

    TRPAportePagtoFatCond_Le_NumIntDoc = gErr

    Select Case gErr

        Case 190950, 190951
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPAPORTEPagtoFatCond", gErr)

        Case ERRO_LEITURA_SEM_DADOS
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190952)

    End Select

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Private Function TRPVoucherAporte_Le_Voucher(ByVal objVouAporte As ClassTRPVoucherAporte, ByVal colVoucherAporte As Collection) As Long

Dim lErro As Long
Dim lComando As Long
Dim tTRPVoucherAporte As typeTRPVoucherAporte
Dim objTRPVoucherAporte As ClassTRPVoucherAporte

On Error GoTo Erro_TRPVoucherAporte_Le_Voucher

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 196461

    'Alocação de espaço no buffer
    tTRPVoucherAporte.sTipVou = String(STRING_TRP_OCR_TIPOVOU, 0)
    tTRPVoucherAporte.sSerie = String(STRING_TRP_OCR_SERIE, 0)

    'Le a tabelaTRPVoucherAporte
    lErro = Comando_Executar(lComando, "SELECT TipVou, Serie, NumVou, TipoPagtoAporte, NumIntDocPagtoAporte, CodigoAporte, " & _
                "ParcelaAporte, ValorAporte, ValorAporteUSS, MoedaAporte FROM TRPVoucherAporte WHERE TipVou= ?  AND Serie= ?  AND NumVou= ? ", _
                tTRPVoucherAporte.sTipVou, tTRPVoucherAporte.sSerie, tTRPVoucherAporte.lNumVou, tTRPVoucherAporte.iTipoPagtoAporte, _
                tTRPVoucherAporte.lNumIntDocPagtoAporte, tTRPVoucherAporte.lCodigoAporte, tTRPVoucherAporte.iParcelaAporte, tTRPVoucherAporte.dValorAporte, tTRPVoucherAporte.dValorAporteUSS, tTRPVoucherAporte.iMoedaAporte, _
                objVouAporte.sTipVou, objVouAporte.sSerie, objVouAporte.lNumVou)
    If lErro <> AD_SQL_SUCESSO Then gError 196462

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196463

    'Sem Dados
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objTRPVoucherAporte = New ClassTRPVoucherAporte

        objTRPVoucherAporte.sTipVou = tTRPVoucherAporte.sTipVou
        objTRPVoucherAporte.sSerie = tTRPVoucherAporte.sSerie
        objTRPVoucherAporte.lNumVou = tTRPVoucherAporte.lNumVou
        objTRPVoucherAporte.iTipoPagtoAporte = tTRPVoucherAporte.iTipoPagtoAporte
        objTRPVoucherAporte.lNumIntDocPagtoAporte = tTRPVoucherAporte.lNumIntDocPagtoAporte
        objTRPVoucherAporte.lCodigoAporte = tTRPVoucherAporte.lCodigoAporte
        objTRPVoucherAporte.iParcelaAporte = tTRPVoucherAporte.iParcelaAporte
        objTRPVoucherAporte.dValorAporte = tTRPVoucherAporte.dValorAporte
        objTRPVoucherAporte.dValorAporteUSS = tTRPVoucherAporte.dValorAporteUSS
        objTRPVoucherAporte.iMoedaAporte = tTRPVoucherAporte.iMoedaAporte
        
        colVoucherAporte.Add objTRPVoucherAporte
        
        'Busca Próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196464
        
    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    TRPVoucherAporte_Le_Voucher = SUCESSO

    Exit Function

Erro_TRPVoucherAporte_Le_Voucher:

    TRPVoucherAporte_Le_Voucher = gErr

    Select Case gErr

        Case 196461
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 196462, 196463, 196464
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPVOUCHERAPORTE", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196465)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function TRPFaturas_Le(ByVal lFatura As Long, ByRef objDestino As Object, ByRef iTipoDocDestino As Integer, Optional ByVal bLeExcluidas As Boolean = False, Optional ByVal bLeTudo As Boolean = False) As Long

Dim lErro As Long
Dim alComando(0 To 6) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim objTitRec As New ClassTituloReceber
Dim objTitPag As New ClassTituloPagar
Dim objNFsPag As New ClassNFsPag
Dim iLeExcluidas As Integer
Dim objFaturaTRP As New ClassFaturaTRP
Dim sUsuario As String
Dim dHora As Double
Dim dtData As Date
Dim iStatus As Integer
Dim dValorAporte As Double
Dim lNumVou As Long
Dim dValor As Double
Dim dHoraRegPagto As Double
Dim dtDataRegPagto As Date
Dim dtDataPagto As Date
Dim lCliente As Long
Dim dValorAporteCred As Double
Dim sMotivoCanc As String
Dim sTipVou As String
Dim sSerie As String

On Error GoTo Erro_TRPFaturas_Le

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 196493
    Next
    
    If bLeExcluidas Then
        iLeExcluidas = MARCADO
    Else
        iLeExcluidas = DESMARCADO
    End If
        
    'Le a tabela TitulosRecTRP
    If bLeExcluidas Then
        lErro = Comando_Executar(alComando(0), "SELECT NumIntDoc,TipoDocDestino, Status FROM TRPFaturas WHERE NumTitulo = ? ORDER BY SiglaDocumento DESC", _
                    lNumIntDoc, iTipoDocDestino, iStatus, lFatura)
    Else
        lErro = Comando_Executar(alComando(0), "SELECT NumIntDoc,TipoDocDestino, Status FROM TRPFaturas WHERE NumTitulo = ? AND Status <> 5 ORDER BY SiglaDocumento DESC", _
                    lNumIntDoc, iTipoDocDestino, iStatus, lFatura)
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 196494

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196495

    If lErro <> AD_SQL_SEM_DADOS Then
    
        objFaturaTRP.lNumFat = lFatura
        
        Select Case iTipoDocDestino
        
            Case TRP_TIPO_DOC_DESTINO_TITREC
            
                objTitRec.lNumIntDoc = lNumIntDoc
        
                'Lê o Título
                lErro = CF("TituloReceber_Le", objTitRec, iLeExcluidas)
                If lErro <> SUCESSO And lErro <> 26061 Then gError 196496
        
                'Se não achou o Título
                If lErro <> SUCESSO Then
        
                    'Lê o Título a Receber baixado
                    lErro = CF("TituloReceberBaixado_Le", objTitRec, iLeExcluidas)
                    If lErro <> SUCESSO And lErro <> 56570 Then gError 196497
        
                End If
                
                Set objTitRec.objInfoUsu = objFaturaTRP
                
                Set objDestino = objTitRec
                
                dValor = objTitRec.dValor
                
                If bLeTudo Then
                
                    lErro = Comando_Executar(alComando(4), "SELECT B.Data, B.DataRegistro, B.HoraRegistro FROM BaixasRec AS B, BaixasParcRec as BP, ParcelasRecTodas AS P, TitulosRecTodos AS T WHERE T.NumIntDoc = P.NumIntTitulo AND P.NumIntDoc = BP.NumIntParcela AND BP.NumIntBaixa = B.NumIntBaixa AND T.NumIntDoc = ? AND BP.ValorBaixado > 0.001 AND BP.Status <> 5 ORDER BY Data", _
                                   dtDataPagto, dtDataRegPagto, dHoraRegPagto, objTitRec.lNumIntDoc)
                    If lErro <> AD_SQL_SUCESSO Then gError 196494
                
                    'Busca Primeiro
                    lErro = Comando_BuscarPrimeiro(alComando(4))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196495
                    
                    If lErro = AD_SQL_SUCESSO Then
                        objFaturaTRP.dHoraRegPagto = dHoraRegPagto
                        objFaturaTRP.dtDataRegPagto = dtDataRegPagto
                        objFaturaTRP.dtDataPagto = dtDataPagto
                    End If
                    
                End If
                
            Case TRP_TIPO_DOC_DESTINO_TITPAG
            
                objTitPag.lNumIntDoc = lNumIntDoc
                
                lErro = CF("TituloPagar_Le_Todos", objTitPag, bLeExcluidas)
                If lErro <> SUCESSO Then gError 196498
                
                Set objTitPag.objInfoUsu = objFaturaTRP
                
                Set objDestino = objTitPag
                
                dValor = objTitPag.dValorTotal
                
                If bLeTudo Then
                
                    lErro = Comando_Executar(alComando(5), "SELECT B.Data, B.DataRegistro FROM BaixasPag AS B, BaixasParcPag as BP, ParcelasPagTodas AS P, TitulosPagTodos AS T WHERE T.NumIntDoc = P.NumIntTitulo AND P.NumIntDoc = BP.NumIntParcela AND BP.NumIntBaixa = B.NumIntBaixa AND T.NumIntDoc = ? AND BP.ValorBaixado > 0.001 AND BP.Status <> 5 ORDER BY Data", _
                                  dtDataPagto, dtDataRegPagto, objTitPag.lNumIntDoc)
                    If lErro <> AD_SQL_SUCESSO Then gError 196494
                
                    'Busca Primeiro
                    lErro = Comando_BuscarPrimeiro(alComando(5))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196495
                    
                    If lErro = AD_SQL_SUCESSO Then
                        objFaturaTRP.dtDataRegPagto = dtDataRegPagto
                        objFaturaTRP.dtDataPagto = dtDataPagto
                    End If
                    
                End If
                
            Case TRP_TIPO_DOC_DESTINO_NFSPAG
            
                 objNFsPag.lNumIntDoc = lNumIntDoc
                
                'Lê o Título
                lErro = CF("NFPag_Le_Todos2", objNFsPag, iLeExcluidas)
                If lErro <> SUCESSO And lErro <> 18257 Then gError 196499
                
                Set objNFsPag.objInfoUsu = objFaturaTRP
                
                Set objDestino = objNFsPag
                
                dValor = objNFsPag.dValorTotal
                
                If bLeTudo Then
                
                    lErro = Comando_Executar(alComando(6), "SELECT B.Data, B.DataRegistro FROM BaixasPag AS B, BaixasParcPag as BP, ParcelasPagTodas AS P, TitulosPagTodos AS T, NFsPag_Todas AS N WHERE T.NumIntDoc = P.NumIntTitulo AND P.NumIntDoc = BP.NumIntParcela AND BP.NumIntBaixa = B.NumIntBaixa AND N.NumIntTitPag = T.NumIntDoc AND N.NumIntDoc = ? AND BP.ValorBaixado > 0.001 AND BP.Status <> 5 ORDER BY Data", _
                                  dtDataPagto, dtDataRegPagto, objNFsPag.lNumIntDoc)
                    If lErro <> AD_SQL_SUCESSO Then gError 196494
                
                    'Busca Primeiro
                    lErro = Comando_BuscarPrimeiro(alComando(6))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196495
                    
                    If lErro = AD_SQL_SUCESSO Then
                        objFaturaTRP.dtDataRegPagto = dtDataRegPagto
                        objFaturaTRP.dtDataPagto = dtDataPagto
                    End If
                    
                End If
               
        End Select

    End If
    
    dValorAporteCred = 0
    
    If bLeTudo Then
    
        sUsuario = String(STRING_USUARIO, 0)
    
        lErro = Comando_Executar(alComando(1), "SELECT Usuario, Data, Hora,ValorAporteCred, ValorAporte FROM TRPTitulosExp WHERE TipoDocOrigem = ? AND NumIntDocOrigem = ? AND Excluido = ?", _
                     sUsuario, dtData, dHora, dValorAporteCred, dValorAporte, iTipoDocDestino, lNumIntDoc, DESMARCADO)
        If lErro <> AD_SQL_SUCESSO Then gError 196494
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196495
    
        If lErro = AD_SQL_SUCESSO Then
        
            objFaturaTRP.sUsuarioGerFat = sUsuario
            objFaturaTRP.dHoraGerFat = dHora
            objFaturaTRP.dtDataGerFat = dtData

            dValorAporte = dValorAporte + dValorAporteCred
        
            objFaturaTRP.dValorAporte = dValorAporte
            objFaturaTRP.dPercAporte = Abs(dValorAporte) / (dValor + Abs(dValorAporte))

        End If
    
        If iStatus = STATUS_EXCLUIDO Then
        
            sUsuario = String(STRING_USUARIO, 0)
            sMotivoCanc = String(STRING_MOTIVOCANCEL, 0)
        
            lErro = Comando_Executar(alComando(1), "SELECT Usuario, Data, Hora,Motivo FROM TRPTitulosExp WHERE TipoDocOrigem = ? AND NumIntDocOrigem = ? AND Excluido = ?", _
                         sUsuario, dtData, dHora, sMotivoCanc, iTipoDocDestino, lNumIntDoc, MARCADO)
            If lErro <> AD_SQL_SUCESSO Then gError 196494
        
            'Busca Primeiro
            lErro = Comando_BuscarPrimeiro(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196495
        
            If lErro = AD_SQL_SUCESSO Then
                objFaturaTRP.sUsuarioCanc = sUsuario
                objFaturaTRP.dHoraCanc = dHora
                objFaturaTRP.dtDataCanc = dtData
                objFaturaTRP.sMotivo = sMotivoCanc
            End If
            
        End If
        
        sTipVou = String(STRING_TRP_OCR_TIPOVOU, 0)
        sSerie = String(STRING_TRP_OCR_SERIE, 0)
        
        lErro = Comando_Executar(alComando(2), "SELECT TipVou, Serie, NumVou FROM DocFaturados WHERE TipoDocDestino = ? AND NumIntDocDestino = ?", _
                     sTipVou, sSerie, lNumVou, iTipoDocDestino, lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 196494
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196495
                
        lErro = Comando_Executar(alComando(3), "SELECT ClienteVou FROM TRPVouchers WHERE NumVou = ? AND Serie = ? AND TipVou = ?", _
                     lCliente, lNumVou, sSerie, sTipVou)
        If lErro <> AD_SQL_SUCESSO Then gError 196494
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196495
        
        If lErro = AD_SQL_SUCESSO Then
            objFaturaTRP.lClienteVou = lCliente
        End If
       
    End If

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TRPFaturas_Le = SUCESSO

    Exit Function

Erro_TRPFaturas_Le:

    TRPFaturas_Le = gErr

    Select Case gErr

        Case 196493
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 196494, 196495
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOSREC", gErr)
            
        Case 196496 To 196499

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196500)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function TitulosRecTRP_Le_NF(ByVal lNumIntDoc As Long, ByVal objNF As ClassNFiscal, iItem As Integer) As Long

Dim lErro As Long
Dim alComando(0 To 0) As Long
Dim iIndice As Integer
Dim lNumIntNF As Long

On Error GoTo Erro_TitulosRecTRP_Le_NF

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 196760
    Next

    'Le a tabelaTitulosRecTRP
    lErro = Comando_Executar(alComando(0), "SELECT N.NumIntDoc, I.Item FROM TRPItemNotaTitulo AS T, NFiscal AS N, ItensNFiscal AS I WHERE T.NumIntDocTitulo = ? AND T.NumIntDocItemNF = I.NumIntDoc AND I.NumIntNF = N.NumIntDoc AND N.Status NOT IN (5,7) ", _
                lNumIntNF, iItem, lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 196761

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196762
    
    If lErro <> AD_SQL_SUCESSO Then gError ERRO_LEITURA_SEM_DADOS
    
    objNF.lNumIntDoc = lNumIntNF
    
    lErro = CF("NFiscal_Le", objNF)
    If lErro <> SUCESSO Then gError 196763
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TitulosRecTRP_Le_NF = SUCESSO

    Exit Function

Erro_TitulosRecTRP_Le_NF:

    TitulosRecTRP_Le_NF = gErr

    Select Case gErr

        Case 196760
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 196761, 196762
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOSRECTRP", gErr)
            
        Case ERRO_LEITURA_SEM_DADOS
            
        Case 196763

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196764)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function TRPClienteRepProd_Le(ByVal objCliente As ClassCliente) As Long

Dim lErro As Long
Dim lComando As Long
Dim tTRPClienteRepProd As typeTRPClienteRepProd
Dim objTRPClienteRepProd As ClassTRPClienteRepProd
Dim objClienteTRP As ClassClienteTRP

On Error GoTo Erro_TRPClienteRepProd_Le

    If Not (objCliente.objInfoUsu Is Nothing) Then
    
        Set objClienteTRP = objCliente.objInfoUsu

        'Executa a abertura do Comando
        lComando = Comando_Abrir()
        If lComando = 0 Then gError 196920
    
        'Alocação de espaço no buffer
        tTRPClienteRepProd.sProduto = String(STRING_PRODUTO, 0)
    
        'Le a tabelaTRPClienteRepProd
        lErro = Comando_Executar(lComando, "SELECT Cliente, Seq, Representante, Produto, PercComis FROM TRPClienteRepProd WHERE Cliente= ? ORDER BY Seq ", _
                    tTRPClienteRepProd.lCliente, tTRPClienteRepProd.iSeq, tTRPClienteRepProd.lRepresentante, tTRPClienteRepProd.sProduto, _
                    tTRPClienteRepProd.dPercComis, _
                    objCliente.lCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 196921
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196922
    
        'Sem Dados
        Do While lErro <> AD_SQL_SEM_DADOS
        
            Set objTRPClienteRepProd = New ClassTRPClienteRepProd
    
            objTRPClienteRepProd.lCliente = tTRPClienteRepProd.lCliente
            objTRPClienteRepProd.iSeq = tTRPClienteRepProd.iSeq
            objTRPClienteRepProd.lRepresentante = tTRPClienteRepProd.lRepresentante
            objTRPClienteRepProd.sProduto = tTRPClienteRepProd.sProduto
            objTRPClienteRepProd.dPercComis = tTRPClienteRepProd.dPercComis
            
            objClienteTRP.colExcRep.Add objTRPClienteRepProd
            
            'Busca Próximo
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196923
            
        Loop
    
        'Fecha Comando
        Call Comando_Fechar(lComando)
        
    End If

    TRPClienteRepProd_Le = SUCESSO

    Exit Function

Erro_TRPClienteRepProd_Le:

    TRPClienteRepProd_Le = gErr

    Select Case gErr

        Case 196920
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 196921 To 196923
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPCLIENTEREPPROD", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196924)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function TRPClienteCorProd_Le(ByVal objCliente As ClassCliente) As Long

Dim lErro As Long
Dim lComando As Long
Dim tTRPClienteCorProd As typeTRPClienteCorProd
Dim objTRPClienteCorProd As ClassTRPClienteCorProd
Dim objClienteTRP As ClassClienteTRP

On Error GoTo Erro_TRPClienteCorProd_Le

    If Not (objCliente.objInfoUsu Is Nothing) Then
    
        Set objClienteTRP = objCliente.objInfoUsu

        'Executa a abertura do Comando
        lComando = Comando_Abrir()
        If lComando = 0 Then gError 196926
    
        'Alocação de espaço no buffer
        tTRPClienteCorProd.sProduto = String(STRING_PRODUTO, 0)
    
        'Le a tabelaTRPClienteCorProdse
        lErro = Comando_Executar(lComando, "SELECT Cliente, Seq, Correntista, Produto, PercComis FROM TRPClienteCorProd WHERE Cliente= ? ORDER BY Seq ", _
                    tTRPClienteCorProd.lCliente, tTRPClienteCorProd.iSeq, tTRPClienteCorProd.lCorrentista, tTRPClienteCorProd.sProduto, _
                    tTRPClienteCorProd.dPercComis, _
                    objCliente.lCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 196927
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196928
    
        'Sem Dados
        Do While lErro <> AD_SQL_SEM_DADOS
        
            Set objTRPClienteCorProd = New ClassTRPClienteCorProd
    
            objTRPClienteCorProd.lCliente = tTRPClienteCorProd.lCliente
            objTRPClienteCorProd.iSeq = tTRPClienteCorProd.iSeq
            objTRPClienteCorProd.lCorrentista = tTRPClienteCorProd.lCorrentista
            objTRPClienteCorProd.sProduto = tTRPClienteCorProd.sProduto
            objTRPClienteCorProd.dPercComis = tTRPClienteCorProd.dPercComis
            
            objClienteTRP.colExcCor.Add objTRPClienteCorProd
            
            'Busca Próximo
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196929
            
        Loop
    
        'Fecha Comando
        Call Comando_Fechar(lComando)
        
    End If

    TRPClienteCorProd_Le = SUCESSO

    Exit Function

Erro_TRPClienteCorProd_Le:

    TRPClienteCorProd_Le = gErr

    Select Case gErr

        Case 196926
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 196927 To 196929
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPCLIENTECorPROD", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196930)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function TRPExcComissaoCli_Le(ByVal objCliente As ClassCliente) As Long

Dim lErro As Long
Dim lComando As Long
Dim tTRPExcComisCli As typeTRPExcComissaoCli
Dim objTRPExcComisCli As ClassTRPExcComissaoCli
Dim objClienteTRP As ClassClienteTRP

On Error GoTo Erro_TRPExcComissaoCli_Le

    If Not (objCliente.objInfoUsu Is Nothing) Then
    
        Set objClienteTRP = objCliente.objInfoUsu

        'Executa a abertura do Comando
        lComando = Comando_Abrir()
        If lComando = 0 Then gError 196926
    
        'Alocação de espaço no buffer
        tTRPExcComisCli.sProduto = String(STRING_PRODUTO, 0)
    
        'Le a tabelaTRPExcComisClise
        lErro = Comando_Executar(lComando, "SELECT C.Cliente, C.Seq, C.Produto, C.PercComissao FROM TRPExcComissaoCli AS C, FiliaisClientesTRP AS F WHERE F.EmpresaPai = C.Cliente AND F.CodCliente= ? ORDER BY C.Seq ", _
                    tTRPExcComisCli.lCliente, tTRPExcComisCli.iSeq, tTRPExcComisCli.sProduto, tTRPExcComisCli.dPercComissao, _
                    objCliente.lCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 196927
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196928
    
        'Sem Dados
        Do While lErro <> AD_SQL_SEM_DADOS
        
            Set objTRPExcComisCli = New ClassTRPExcComissaoCli
    
            objTRPExcComisCli.lCliente = tTRPExcComisCli.lCliente
            objTRPExcComisCli.iSeq = tTRPExcComisCli.iSeq
            objTRPExcComisCli.sProduto = tTRPExcComisCli.sProduto
            objTRPExcComisCli.dPercComissao = tTRPExcComisCli.dPercComissao
            
            objClienteTRP.colExcComissaoCli.Add objTRPExcComisCli
            
            'Busca Próximo
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196929
            
        Loop
    
        'Fecha Comando
        Call Comando_Fechar(lComando)
        
    End If

    TRPExcComissaoCli_Le = SUCESSO

    Exit Function

Erro_TRPExcComissaoCli_Le:

    TRPExcComissaoCli_Le = gErr

    Select Case gErr

        Case 196926
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 196927 To 196929
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPEXCCOMISSAOCLI", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196930)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function TRPTiposOcorrencia_Le(ByVal objTRPTiposOcorrencia As ClassTRPTiposOcorrencia) As Long

Dim lErro As Long
Dim lComando As Long
Dim tTRPTiposOcorrencia As typeTRPTiposOcorrencia

On Error GoTo Erro_TRPTiposOcorrencia_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 198081

    'Alocação de espaço no buffer
    tTRPTiposOcorrencia.sDescricao = String(STRING_TRP_TIPOOCR_DESCRICAO, 0)

    'Le a tabelaTRPTiposOcorrencia
    lErro = Comando_Executar(lComando, "SELECT Codigo, Descricao, EstornaAporteVou, ConsideraComisInt, AlteraComiVou, AlteraCMCC, " & _
                "AlteraCMC, AlteraCMR, AlteraOVER, AlteraCMA, AceitaVlrPositivo, AceitaVlrNegativo, IncideSobre FROM TRPTiposOcorrencia WHERE Codigo= ? ", _
                tTRPTiposOcorrencia.iCodigo, tTRPTiposOcorrencia.sDescricao, tTRPTiposOcorrencia.iEstornaAporteVou, tTRPTiposOcorrencia.iConsideraComisInt, _
                tTRPTiposOcorrencia.iAlteraComiVou, tTRPTiposOcorrencia.iAlteraCMCC, tTRPTiposOcorrencia.iAlteraCMC, tTRPTiposOcorrencia.iAlteraCMR, tTRPTiposOcorrencia.iAlteraOVER, _
                tTRPTiposOcorrencia.iAlteraCMA, tTRPTiposOcorrencia.iAceitaVlrPositivo, tTRPTiposOcorrencia.iAceitaVlrNegativo, tTRPTiposOcorrencia.iIncideSobre, _
                objTRPTiposOcorrencia.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 198082

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 198083

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objTRPTiposOcorrencia.iCodigo = tTRPTiposOcorrencia.iCodigo
    objTRPTiposOcorrencia.sDescricao = tTRPTiposOcorrencia.sDescricao
    objTRPTiposOcorrencia.iEstornaAporteVou = tTRPTiposOcorrencia.iEstornaAporteVou
    objTRPTiposOcorrencia.iConsideraComisInt = tTRPTiposOcorrencia.iConsideraComisInt
    objTRPTiposOcorrencia.iAlteraComiVou = tTRPTiposOcorrencia.iAlteraComiVou
    objTRPTiposOcorrencia.iAlteraCMCC = tTRPTiposOcorrencia.iAlteraCMCC
    objTRPTiposOcorrencia.iAlteraCMC = tTRPTiposOcorrencia.iAlteraCMC
    objTRPTiposOcorrencia.iAlteraCMR = tTRPTiposOcorrencia.iAlteraCMR
    objTRPTiposOcorrencia.iAlteraOVER = tTRPTiposOcorrencia.iAlteraOVER
    objTRPTiposOcorrencia.iAlteraCMA = tTRPTiposOcorrencia.iAlteraCMA
    objTRPTiposOcorrencia.iAceitaVlrPositivo = tTRPTiposOcorrencia.iAceitaVlrPositivo
    objTRPTiposOcorrencia.iAceitaVlrNegativo = tTRPTiposOcorrencia.iAceitaVlrNegativo
    objTRPTiposOcorrencia.iIncideSobre = tTRPTiposOcorrencia.iIncideSobre

    'Fecha Comando
    Call Comando_Fechar(lComando)

    TRPTiposOcorrencia_Le = SUCESSO

    Exit Function

Erro_TRPTiposOcorrencia_Le:

    TRPTiposOcorrencia_Le = gErr

    Select Case gErr

        Case 198081
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 198082, 198083
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPTIPOSOCORRENCIA", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198084)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function TRPCarrega_TipoOrigemOCR(ByVal objCombo As Object) As Long

Dim lErro As Long
Dim lComando As Long
Dim tTRPTiposOcorrencia As typeTRPTiposOcorrencia

On Error GoTo Erro_TRPCarrega_TipoOrigemOCR

    objCombo.Clear

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 198001

    'Alocação de espaço no buffer
    tTRPTiposOcorrencia.sDescricao = String(STRING_TRP_TIPOOCR_DESCRICAO, 0)

    'Le a tabelaTRPTiposOcorrencia
    lErro = Comando_Executar(lComando, "SELECT Codigo, Descricao FROM TRPTiposOcorrencia ORDER BY Codigo ", _
                tTRPTiposOcorrencia.iCodigo, tTRPTiposOcorrencia.sDescricao)
    If lErro <> AD_SQL_SUCESSO Then gError 198002

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 198003
    
    Do While lErro <> AD_SQL_SEM_DADOS
    
        objCombo.AddItem tTRPTiposOcorrencia.iCodigo & SEPARADOR & tTRPTiposOcorrencia.sDescricao
        objCombo.ItemData(objCombo.NewIndex) = tTRPTiposOcorrencia.iCodigo
    
        'Busca Próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 198004

    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    TRPCarrega_TipoOrigemOCR = SUCESSO

    Exit Function

Erro_TRPCarrega_TipoOrigemOCR:

    TRPCarrega_TipoOrigemOCR = gErr

    Select Case gErr

        Case 198001
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 198002, 198003, 198004
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPTIPOSOCORRENCIA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198005)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function TRPGerComiInt_Le(ByVal objTRPGerComiInt As ClassTRPGerComiInt) As Long

Dim lErro As Long
Dim lComando As Long
Dim tTRPGerComiInt As typeTRPGerComiInt

On Error GoTo Erro_TRPGerComiInt_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 197892

    'Alocação de espaço no buffer
    tTRPGerComiInt.sUsuario = String(STRING_USUARIO_CODIGO, 0)
    tTRPGerComiInt.sDiretorio = String(STRING_NOME_ARQUIVO, 0)

    'Le a tabelaTRPVoucherInfo
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, DataGeracao, HoraGeracao, Usuario, DataEmiAte, Diretorio, Previa FROM TRPGerComiInt WHERE DataGeracao = ? ", _
              tTRPGerComiInt.lNumIntDoc, tTRPGerComiInt.dtDataGeracao, tTRPGerComiInt.dHoraGeracao, tTRPGerComiInt.sUsuario, tTRPGerComiInt.dtDataEmiAte, tTRPGerComiInt.sDiretorio, tTRPGerComiInt.iPrevia, objTRPGerComiInt.dtDataGeracao)
    If lErro <> AD_SQL_SUCESSO Then gError 197893

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197894

    If lErro <> AD_SQL_SUCESSO Then gError 197895

    objTRPGerComiInt.dHoraGeracao = tTRPGerComiInt.dHoraGeracao
    objTRPGerComiInt.sUsuario = tTRPGerComiInt.sUsuario
    objTRPGerComiInt.dtDataEmiAte = tTRPGerComiInt.dtDataEmiAte
    objTRPGerComiInt.sDiretorio = tTRPGerComiInt.sDiretorio
    objTRPGerComiInt.iPrevia = tTRPGerComiInt.iPrevia
    objTRPGerComiInt.lNumIntDoc = tTRPGerComiInt.lNumIntDoc

    'Fecha Comando
    Call Comando_Fechar(lComando)

    TRPGerComiInt_Le = SUCESSO

    Exit Function

Erro_TRPGerComiInt_Le:

    TRPGerComiInt_Le = gErr

    Select Case gErr

        Case 197892
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 197893, 197894
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPGERCOMIINT", gErr)

        Case 197895

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197896)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function TRPProdTarifa_Le(ByVal colTarifas As Collection, Optional ByVal sProduto As String = "") As Long

Dim lErro As Long
Dim lComando As Long
Dim tTRPProdTarifa As typeTRPProdTarifa
Dim objTRPProdTarifa As ClassTRPProdTarifa
Dim sProdAnt As String
Dim objTRPProdTarifaCalc As ClassTRPProdTafCalc

On Error GoTo Erro_TRPProdTarifa_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 200240

    'Alocação de espaço no buffer
    tTRPProdTarifa.sProduto = String(STRING_PRODUTO, 0)

    'Le a tabelaTRPProdTarifa
    If sProduto = "" Then
        lErro = Comando_Executar(lComando, "SELECT Produto, DiasDe, DiasAte, Diario, Valor, " & _
                    "ValorAdicional, DataVigencia FROM TRPProdTarifa ORDER BY Produto, DataVigencia, DiasDe ", _
                     tTRPProdTarifa.sProduto, tTRPProdTarifa.iDiasDe, tTRPProdTarifa.iDiasAte, _
                    tTRPProdTarifa.iDiario, tTRPProdTarifa.dValor, tTRPProdTarifa.dValorAdicional, tTRPProdTarifa.dtVigencia)
    Else
        lErro = Comando_Executar(lComando, "SELECT Produto, DiasDe, DiasAte, Diario, Valor, " & _
                    "ValorAdicional, DataVigencia FROM TRPProdTarifa WHERE Produto = ? ORDER BY Produto, DataVigencia, DiasDe ", _
                     tTRPProdTarifa.sProduto, tTRPProdTarifa.iDiasDe, tTRPProdTarifa.iDiasAte, _
                    tTRPProdTarifa.iDiario, tTRPProdTarifa.dValor, tTRPProdTarifa.dValorAdicional, tTRPProdTarifa.dtVigencia, sProduto)
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 200241

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200242

    'Sem Dados
    Do While lErro <> AD_SQL_SEM_DADOS
    
        If UCase(sProdAnt) <> UCase(tTRPProdTarifa.sProduto) Then
            sProdAnt = tTRPProdTarifa.sProduto
            Set objTRPProdTarifaCalc = New ClassTRPProdTafCalc
            objTRPProdTarifaCalc.sProduto = tTRPProdTarifa.sProduto
            colTarifas.Add objTRPProdTarifaCalc
        End If
    
        Set objTRPProdTarifa = New ClassTRPProdTarifa

        objTRPProdTarifa.dtVigencia = tTRPProdTarifa.dtVigencia
        objTRPProdTarifa.sProduto = tTRPProdTarifa.sProduto
        objTRPProdTarifa.iDiasDe = tTRPProdTarifa.iDiasDe
        objTRPProdTarifa.iDiasAte = tTRPProdTarifa.iDiasAte
        objTRPProdTarifa.iDiario = tTRPProdTarifa.iDiario
        objTRPProdTarifa.dValor = tTRPProdTarifa.dValor
        objTRPProdTarifa.dValorAdicional = tTRPProdTarifa.dValorAdicional
        
        objTRPProdTarifaCalc.colTarifas.Add objTRPProdTarifa

        'Busca Próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200243

    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    TRPProdTarifa_Le = SUCESSO

    Exit Function

Erro_TRPProdTarifa_Le:

    TRPProdTarifa_Le = gErr

    Select Case gErr

        Case 200240
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 200241, 200242, 200243
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPPRODTARIFA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200244)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function TRPAcordoTarifa_Le(ByVal colTarifas As Collection, ByVal sProduto As String, ByVal iDestino As Integer, ByVal lCliente As Long, ByVal dtData As Date) As Long

Dim lErro As Long
Dim lComando As Long
Dim tTRPAcordoTarifa As typeTRPAcordoTarifaDif
Dim objTRPAcordoTarifa As ClassTRPAcordoTarifaDif

On Error GoTo Erro_TRPAcordoTarifa_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 200240

    'Le a tabelaTRPAcordoTarifa
    lErro = Comando_Executar(lComando, "SELECT A.DiasDe, A.DiasAte, A.Diario, A.Valor, A.ValorAdicional FROM WebTarifasAcordo AS A, FiliaisClientesTRP AS F " & _
                                        "WHERE A.Produto = ? AND A.Destino = ? AND A.ValidadeDe <= ? AND A.ValidadeAte >= ? AND A.Cliente = F.EmpresaPai AND F.CodCliente = ? " & _
                                        "ORDER BY A.DiasDe ", _
                 tTRPAcordoTarifa.iDiasDe, tTRPAcordoTarifa.iDiasAte, _
                tTRPAcordoTarifa.iDiario, tTRPAcordoTarifa.dValor, tTRPAcordoTarifa.dValorAdicional, sProduto, iDestino, dtData, dtData, lCliente)
    If lErro <> AD_SQL_SUCESSO Then gError 200241

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200242

    'Sem Dados
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objTRPAcordoTarifa = New ClassTRPAcordoTarifaDif

        objTRPAcordoTarifa.iDiasDe = tTRPAcordoTarifa.iDiasDe
        objTRPAcordoTarifa.iDiasAte = tTRPAcordoTarifa.iDiasAte
        objTRPAcordoTarifa.iDiario = tTRPAcordoTarifa.iDiario
        objTRPAcordoTarifa.dValor = tTRPAcordoTarifa.dValor
        objTRPAcordoTarifa.dValorAdicional = tTRPAcordoTarifa.dValorAdicional
        
        colTarifas.Add objTRPAcordoTarifa

        'Busca Próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200243

    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    TRPAcordoTarifa_Le = SUCESSO

    Exit Function

Erro_TRPAcordoTarifa_Le:

    TRPAcordoTarifa_Le = gErr

    Select Case gErr

        Case 200240
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 200241, 200242, 200243
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPACORDOTARIFADIF", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200244)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function TRPUsuarios_Le(ByVal objTRPUsuarios As ClassTRPUsuarios) As Long

Dim lErro As Long
Dim lComando As Long
Dim tTRPUsuarios As typeTRPUsuarios

On Error GoTo Erro_TRPUsuarios_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 200310

    'Alocação de espaço no buffer
    tTRPUsuarios.sLogin = String(STRING_WEB_LOGIN, 0)
    tTRPUsuarios.sSenha = String(STRING_WEB_SENHA, 0)
    tTRPUsuarios.sGrupoAcesso = String(STRING_WEB_GRUPOACESSO, 0)

    'Le a tabelaTRPUsuarios
    lErro = Comando_Executar(lComando, "SELECT TipoUsuario, Codigo, Login, Senha, AlteraSenhaProxLog, GrupoAcesso FROM TRPUsuarios WHERE TipoUsuario= ?  AND Codigo= ? ", _
                tTRPUsuarios.iTipoUsuario, tTRPUsuarios.lCodigo, tTRPUsuarios.sLogin, tTRPUsuarios.sSenha, _
                tTRPUsuarios.iAlteraSenhaProxLog, tTRPUsuarios.sGrupoAcesso, _
                objTRPUsuarios.iTipoUsuario, objTRPUsuarios.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 200311

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200312

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objTRPUsuarios.iTipoUsuario = tTRPUsuarios.iTipoUsuario
    objTRPUsuarios.lCodigo = tTRPUsuarios.lCodigo
    objTRPUsuarios.sLogin = tTRPUsuarios.sLogin
    objTRPUsuarios.sSenha = tTRPUsuarios.sSenha
    objTRPUsuarios.iAlteraSenhaProxLog = tTRPUsuarios.iAlteraSenhaProxLog
    objTRPUsuarios.sGrupoAcesso = tTRPUsuarios.sGrupoAcesso

    'Fecha Comando
    Call Comando_Fechar(lComando)

    TRPUsuarios_Le = SUCESSO

    Exit Function

Erro_TRPUsuarios_Le:

    TRPUsuarios_Le = gErr

    Select Case gErr

        Case 200310
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 200311, 200312
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPUSUARIOS", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200313)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function TRPVouPassageiros_Le(ByVal lComando As Long, ByVal objVou As ClassTRPVouchers) As Long

Dim lErro As Long
Dim tPax As typeTRPVouPassageiros
Dim objPax As ClassTRPVouPassageiros
Dim bAbriuComando As Boolean

On Error GoTo Erro_TRPVouPassageiros_Le

    bAbriuComando = False
    If lComando = 0 Then
        bAbriuComando = True
        'Executa a abertura do Comando
        lComando = Comando_Abrir()
        If lComando = 0 Then gError 200310
    End If

    'Alocação de espaço no buffer
    tPax.sTipoVou = String(STRING_TRP_OCR_TIPOVOU, 0)
    tPax.sSerie = String(STRING_TRP_OCR_SERIE, 0)
    tPax.sNome = String(STRING_TRP_VOU_PAXNOME, 0)
    tPax.sTipoDocumento = String(STRING_TRP_VOU_PAXNOME, 0)
    tPax.sNumeroDocumento = String(STRING_TRP_TAMANHO_OUTROS, 0)
    tPax.sPrimeiroNome = String(STRING_TRP_TAMANHO_OUTROS, 0)
    tPax.sSobreNome = String(STRING_TRP_TAMANHO_OUTROS, 0)
    tPax.sSexo = String(1, 0)

    'Le a tabelaTRPVouPassageiros
    lErro = Comando_Executar(lComando, "SELECT TipoVou, Serie, NumVou, Seq, Nome, DataNascimento, " & _
                "TipoDocumento, NumeroDocumento, Sexo, ValorPago, Status, Titular, ValorPagoEmi, PrimeiroNome, Sobrenome FROM TRPVouPassageiros WHERE TipoVou= ?  AND Serie= ?  AND NumVou= ? ORDER BY Seq ", _
                tPax.sTipoVou, tPax.sSerie, tPax.lNumVou, tPax.iSeq, _
                tPax.sNome, tPax.dtDataNascimento, tPax.sTipoDocumento, tPax.sNumeroDocumento, tPax.sSexo, _
                tPax.dValorPago, tPax.iStatus, tPax.iTitular, tPax.dValorPagoEmi, tPax.sPrimeiroNome, tPax.sSobreNome, _
                objVou.sTipVou, objVou.sSerie, objVou.lNumVou)
    If lErro <> AD_SQL_SUCESSO Then gError 200314

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200315

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objPax = New ClassTRPVouPassageiros

        objPax.sTipoVou = tPax.sTipoVou
        objPax.sSerie = tPax.sSerie
        objPax.lNumVou = tPax.lNumVou
        objPax.iSeq = tPax.iSeq
        objPax.sNome = tPax.sNome
        objPax.dtDataNascimento = tPax.dtDataNascimento
        objPax.sTipoDocumento = tPax.sTipoDocumento
        objPax.sNumeroDocumento = tPax.sNumeroDocumento
        objPax.sSexo = tPax.sSexo
        objPax.dValorPago = tPax.dValorPago
        objPax.dValorPagoEmi = tPax.dValorPagoEmi
        objPax.iStatus = tPax.iStatus
        objPax.iTitular = tPax.iTitular
        objPax.sPrimeiroNome = tPax.sPrimeiroNome
        objPax.sSobreNome = tPax.sSobreNome

        objVou.colPassageiros.Add objPax

        'Busca Próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200316
        
    Loop
    
    If bAbriuComando Then Call Comando_Fechar(lComando)
    
    TRPVouPassageiros_Le = SUCESSO

    Exit Function

Erro_TRPVouPassageiros_Le:

    TRPVouPassageiros_Le = gErr

    Select Case gErr

        Case 200310
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 200314 To 200316
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPVOUPASSAGEIROS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200317)

    End Select

    If bAbriuComando Then Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function TRPVouVendedores_Le(ByVal lComando As Long, ByVal objVou As ClassTRPVouchers) As Long

Dim lErro As Long
Dim objVend As ClassVendedor
Dim iVendedor As Integer
Dim dPercComissao As Double

On Error GoTo Erro_TRPVouVendedores_Le

    'Le a tabelaTRPVouVendedores
    lErro = Comando_Executar(lComando, "SELECT Vendedor, PercComissao FROM TRPVouVendedores WHERE TipVou= ?  AND Serie= ?  AND NumVou= ? ", _
                iVendedor, dPercComissao, objVou.sTipVou, objVou.sSerie, objVou.lNumVou)
    If lErro <> AD_SQL_SUCESSO Then gError 200314

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200315

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objVend = New ClassVendedor

        objVend.iCodigo = iVendedor
        objVend.dPercComissao = dPercComissao

        objVou.colVendedores.Add objVend

        'Busca Próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200316
        
    Loop
    
    TRPVouVendedores_Le = SUCESSO

    Exit Function

Erro_TRPVouVendedores_Le:

    TRPVouVendedores_Le = gErr

    Select Case gErr

        Case 200314 To 200316
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPVOUVENDEDORES", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200317)

    End Select

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function TRPVouContatos_Le(ByVal lComando As Long, ByVal objVou As ClassTRPVouchers) As Long

Dim lErro As Long
Dim tContatos As typeTRPVouContatos
Dim objContatos As ClassTRPVouContatos

On Error GoTo Erro_TRPVouContatos_Le

    'Alocação de espaço no buffer
    tContatos.sTipoVou = String(STRING_TRP_OCR_TIPOVOU, 0)
    tContatos.sSerie = String(STRING_TRP_OCR_SERIE, 0)
    tContatos.sNome = String(STRING_TRP_VOU_PAXNOME, 0)
    tContatos.sTelefone = String(STRING_TELEFONE, 0)

    'Le a tabelaTRPVouContatos
    lErro = Comando_Executar(lComando, "SELECT TipoVou, Serie, NumVou, Seq, Nome, Telefone FROM TRPVouContatos WHERE TipoVou= ?  AND Serie= ?  AND NumVou= ? ORDER BY Seq ", _
                tContatos.sTipoVou, tContatos.sSerie, tContatos.lNumVou, tContatos.iSeq, _
                tContatos.sNome, tContatos.sTelefone, _
                objVou.sTipVou, objVou.sSerie, objVou.lNumVou)
    If lErro <> AD_SQL_SUCESSO Then gError 200318

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200319

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objContatos = New ClassTRPVouContatos

        objContatos.sTipoVou = tContatos.sTipoVou
        objContatos.sSerie = tContatos.sSerie
        objContatos.lNumVou = tContatos.lNumVou
        objContatos.iSeq = tContatos.iSeq
        objContatos.sNome = tContatos.sNome
        objContatos.sTelefone = tContatos.sTelefone
        
        objVou.colContatos.Add objContatos
        
        'Busca próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 2003120
        
    Loop

    TRPVouContatos_Le = SUCESSO

    Exit Function

Erro_TRPVouContatos_Le:

    TRPVouContatos_Le = gErr

    Select Case gErr

        Case 200318 To 200320
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPVOUCONTATOS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200321)

    End Select

    Exit Function

End Function

Public Function TRPEmissores_Le_Forn(ByVal objEmi As ClassCliEmissoresTRP) As Long

Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim tEmi As typeCliEmissoresTRP
Dim tEmiExc As typeCliEmissoresExcTRP
Dim objEmiExc As ClassCliEmissoresExcTRP
Dim objUsuWeb As New ClassTRPUsuarios

On Error GoTo Erro_TRPEmissores_Le_Forn

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 200324
    Next

    'Le a tabelaTRPVouchers
    lErro = Comando_Executar(alComando(1), "SELECT NumIntDoc, Cliente, Seq, PercComissao, EmissorSuperior, Codigo FROM TRPCliEmissores WHERE FornEmissor = ? ", _
    tEmi.lNumIntDoc, tEmi.lCliente, tEmi.iSeq, tEmi.dPercComissao, tEmi.lEmissorSuperior, tEmi.lCodigo, objEmi.lFornEmissor)
    If lErro <> AD_SQL_SUCESSO Then gError 200325

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200326

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objEmi.dPercComissao = tEmi.dPercComissao
    objEmi.iSeq = tEmi.iSeq
    objEmi.lCliente = tEmi.lCliente
    objEmi.lCodigo = tEmi.lCodigo
    objEmi.lNumIntDoc = tEmi.lNumIntDoc
    objEmi.lSuperior = tEmi.lEmissorSuperior

    tEmiExc.sProduto = String(STRING_PRODUTO, 0)
    
    'Le a tabela CliEmissoresExcTRP
    lErro = Comando_Executar(alComando(2), "SELECT NumIntDoc, NumIntDocCliEmi, Seq, Produto, PercComissao FROM TRPCliEmissoresExc WHERE NumIntDocCliEmi = ? ", _
                tEmiExc.lNumIntDoc, tEmiExc.lNumIntDocCliEmi, tEmiExc.iSeq, tEmiExc.sProduto, tEmiExc.dPercComissao, _
                objEmi.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 200327

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200328
    
    Do While lErro = AD_SQL_SUCESSO
    
        Set objEmiExc = New ClassCliEmissoresExcTRP

        objEmiExc.lNumIntDoc = tEmiExc.lNumIntDoc
        objEmiExc.lNumIntDocCliEmi = tEmiExc.lNumIntDocCliEmi
        objEmiExc.iSeq = tEmiExc.iSeq
        objEmiExc.sProduto = tEmiExc.sProduto
        objEmiExc.dPercComissao = tEmiExc.dPercComissao
    
        objEmi.colCliEmissoresExcTRP.Add objEmiExc
        
        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200329
    
    Loop
    
    objUsuWeb.iTipoUsuario = TRP_USUARIO_EMISSOR
    objUsuWeb.lCodigo = objEmi.lCodigo
    
    lErro = TRPUsuarios_Le(objUsuWeb)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 200331
    
    Set objEmi.objUsuarioWeb = objUsuWeb
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TRPEmissores_Le_Forn = SUCESSO

    Exit Function

Erro_TRPEmissores_Le_Forn:

    TRPEmissores_Le_Forn = gErr

    Select Case gErr

        Case 200324
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 200325, 200326
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPCLIEMISSORES", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora
        
        Case 200327, 200328, 200329
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPCLIEMISSORESEXC", gErr)
            
        Case 200331

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200330)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function TRPEmissores_Le(ByVal objEmi As ClassCliEmissoresTRP) As Long

Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim tEmi As typeCliEmissoresTRP
Dim tEmiExc As typeCliEmissoresExcTRP
Dim objEmiExc As ClassCliEmissoresExcTRP
Dim objUsuWeb As New ClassTRPUsuarios

On Error GoTo Erro_TRPEmissores_Le

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 200324
    Next

    'Le a tabelaTRPVouchers
    lErro = Comando_Executar(alComando(1), "SELECT NumIntDoc, Cliente, Seq, PercComissao, EmissorSuperior, FornEmissor FROM TRPCliEmissores WHERE Codigo = ? ", _
    tEmi.lNumIntDoc, tEmi.lCliente, tEmi.iSeq, tEmi.dPercComissao, tEmi.lEmissorSuperior, tEmi.lFornEmissor, objEmi.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 200325

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200326

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objEmi.dPercComissao = tEmi.dPercComissao
    objEmi.iSeq = tEmi.iSeq
    objEmi.lCliente = tEmi.lCliente
    objEmi.lFornEmissor = tEmi.lFornEmissor
    objEmi.lNumIntDoc = tEmi.lNumIntDoc
    objEmi.lSuperior = tEmi.lEmissorSuperior

    tEmiExc.sProduto = String(STRING_PRODUTO, 0)
    
    'Le a tabela CliEmissoresExcTRP
    lErro = Comando_Executar(alComando(2), "SELECT NumIntDoc, NumIntDocCliEmi, Seq, Produto, PercComissao FROM TRPCliEmissoresExc WHERE NumIntDocCliEmi = ? ", _
                tEmiExc.lNumIntDoc, tEmiExc.lNumIntDocCliEmi, tEmiExc.iSeq, tEmiExc.sProduto, tEmiExc.dPercComissao, _
                objEmi.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 200327

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200328
    
    Do While lErro = AD_SQL_SUCESSO
    
        Set objEmiExc = New ClassCliEmissoresExcTRP

        objEmiExc.lNumIntDoc = tEmiExc.lNumIntDoc
        objEmiExc.lNumIntDocCliEmi = tEmiExc.lNumIntDocCliEmi
        objEmiExc.iSeq = tEmiExc.iSeq
        objEmiExc.sProduto = tEmiExc.sProduto
        objEmiExc.dPercComissao = tEmiExc.dPercComissao
    
        objEmi.colCliEmissoresExcTRP.Add objEmiExc
        
        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200329
    
    Loop
    
    objUsuWeb.iTipoUsuario = TRP_USUARIO_EMISSOR
    objUsuWeb.lCodigo = objEmi.lCodigo
    
    lErro = CF("TRPUsuarios_Le", objUsuWeb)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 200331
    
    Set objEmi.objUsuarioWeb = objUsuWeb
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TRPEmissores_Le = SUCESSO

    Exit Function

Erro_TRPEmissores_Le:

    TRPEmissores_Le = gErr

    Select Case gErr

        Case 200324
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 200325, 200326
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPCLIEMISSORES", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora
        
        Case 200327, 200328, 200329
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPCLIEMISSORESEXC", gErr)
            
        Case 200331

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200330)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function TRPUsuarios_Testa_Login(ByVal objTRPUsuarios As ClassTRPUsuarios) As Long

Dim lErro As Long
Dim lComando As Long
Dim lCodigoAux As Long

On Error GoTo Erro_TRPUsuarios_Testa_Login

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 200349

    'Le a tabelaTRPUsuarios
    lErro = Comando_Executar(lComando, "SELECT Codigo FROM TRPUsuarios WHERE (TipoUsuario <> ?  OR Codigo <> ?) AND Login = ? ", _
                lCodigoAux, objTRPUsuarios.iTipoUsuario, objTRPUsuarios.lCodigo, objTRPUsuarios.sLogin)
    If lErro <> AD_SQL_SUCESSO Then gError 200350

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200351

    'Sem Dados
    If lErro <> AD_SQL_SEM_DADOS Then gError 200352

    'Fecha Comando
    Call Comando_Fechar(lComando)

    TRPUsuarios_Testa_Login = SUCESSO

    Exit Function

Erro_TRPUsuarios_Testa_Login:

    TRPUsuarios_Testa_Login = gErr

    Select Case gErr

        Case 200349
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 200350, 200351
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPUSUARIOS", gErr)

        Case 200352
            Call Rotina_Erro(vbOKOnly, "ERRO_LOGIN_JA_UTILIZADO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200353)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function TRPDestino_Le(ByVal iCodigo As Integer, sSigla As String, sDescricao As String) As Long

Dim lErro As Long
Dim lComando As Long

On Error GoTo Erro_TRPDestino_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 200414

    'Alocação de espaço no buffer
    sSigla = String(STRING_CAMPOSGENERICOSVALORES_VALOR, 0)
    sDescricao = String(STRING_CAMPOSGENERICOS_DESCRICAO, 0)
    
    'Le a tabelaTRPOcorrenciaDet
    lErro = Comando_Executar(lComando, "SELECT Valor, Complemento1 FROM CamposGenericosValores WHERE CodCampo = 24 AND CodValor = ?", sSigla, sDescricao, iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 200415

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200416

    If lErro = AD_SQL_SUCESSO Then gError ERRO_LEITURA_SEM_DADOS

    'Fecha Comando
    Call Comando_Fechar(lComando)

    TRPDestino_Le = SUCESSO

    Exit Function

Erro_TRPDestino_Le:

    TRPDestino_Le = gErr

    Select Case gErr

        Case 200414
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 200415, 200416
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CAMPOSGENERICOSVALORES", gErr)
        
        Case ERRO_LEITURA_SEM_DADOS
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200417)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function TRPCambio_Le(ByVal iMoeda As Integer, ByVal dtData As Date, dCambio As Double) As Long

Dim lErro As Long
Dim lComando As Long

On Error GoTo Erro_TRPCambio_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 200421
   
    'Le a tabelaTRPOcorrenciaDet
    lErro = Comando_Executar(lComando, "SELECT Cambio FROM WebCambio WHERE Moeda = ? AND Data <= ? ORDER BY Data DESC", dCambio, iMoeda, dtData)
    If lErro <> AD_SQL_SUCESSO Then gError 200422

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200423

    If lErro = AD_SQL_SUCESSO Then gError ERRO_LEITURA_SEM_DADOS

    'Fecha Comando
    Call Comando_Fechar(lComando)

    TRPCambio_Le = SUCESSO

    Exit Function

Erro_TRPCambio_Le:

    TRPCambio_Le = gErr

    Select Case gErr

        Case 200421
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 200422, 200423
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COTACAOMOEDA", gErr)
        
        Case ERRO_LEITURA_SEM_DADOS
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200424)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function TRPFatCartao_Le(ByVal colAuto As Collection) As Long
'Atualiza os Títulos de Cartão de Crédito a partir dos vouchers pagos com cartão

Dim lErro As Long
Dim lTransacao As Long
Dim alComando(0 To 0) As Long
Dim iIndice As Integer
Dim lNumAuto As Long
Dim lNumAutoAnt As Long
Dim sCia As String
Dim sCiaAnt As String
Dim sNumCCred As String
Dim sNumCCredAnt As String
Dim dtDataAuto As Date
Dim dtDataAutoAnt As Date
Dim lNumVou As Long
Dim sSerie As String
Dim sTipo As String
Dim dtData As Date
Dim lClienteVou As Long
Dim sClienteVou As String
Dim dValor As Double
Dim dValorComissao As Double
Dim sNomePax As String
Dim sSobrenomePax As String
Dim sTitular As String
Dim sTitularCPF As String
Dim iCodSeguranca As Integer
Dim lClienteFat As Long
Dim sClienteFat As String
Dim iNumParc As Integer
Dim dTaxa As Double
Dim sValidadeCC As String
Dim objVouAutoCC As ClassTRPVouAutoCC
Dim objVou As ClassTRPFATInfoVoucher
Dim sSQL As String
Dim iFilialEmpresa As Integer

On Error GoTo Erro_TRPFatCartao_Le
    
    'Abre o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 200668
    Next
    
    sSQL = "SELECT V.CiaCart, V.NumCCred, V.DataAutoCC, V.NumAuto, V.NumVou, V.Serie, V.TipVou, V.Data, V.ClienteVou, C.NomeReduzido, " & _
                    "V.Valor, V.ValorComissao, V.PassageiroNome, V.PassageiroSobrenome, V.Titular, V.TitularCPF, V.CodSegurancaCC, " & _
                    "R.Cliente , A.NomeReduzido, V.QuantParc, P.Taxa, V.ValidadeCC, T.FilialEmpresa " & _
                    "FROM TRPVouchers AS V, Clientes AS C, AdmMeioPagto AS M, Redes AS R, Clientes AS A, AdmMeioPagtoCondPagto AS P, FiliaisClientesTRP AS T " & _
                    "WHERE V.Cartao = 1 AND V.NumFat = 0 AND V.GeraComissao = 0 AND V.CiaCart <> '' AND V.NumAuto <> 0 AND " & _
                    "V.ClienteVou = C.Codigo AND V.CiaCart = M.Nome AND M.Rede = R.Codigo AND A.Codigo = R.Cliente AND " & _
                    "P.AdmMeioPagto = M.Codigo AND P.NumParcelas = V.QuantParc AND T.CodCliente = C.Codigo " & _
                    "ORDER BY V.CiaCart, V.NumCCred, V.DataAutoCC, V.NumAuto, V.NumVou"
    
    sCia = String(STRING_TRP_VOU_CIACART, 0)
    sNumCCred = String(STRING_TRP_VOU_NUMCCRED, 0)
    sSerie = String(STRING_TRP_OCR_SERIE, 0)
    sTipo = String(STRING_TRP_OCR_TIPOVOU, 0)
    sClienteVou = String(STRING_CLIENTE_NOME_REDUZIDO, 0)
    sNomePax = String(STRING_TRP_VOU_PAXNOME, 0)
    sSobrenomePax = String(STRING_TRP_VOU_PAXNOME, 0)
    sTitular = String(STRING_TRP_VOU_TITULAR, 0)
    sTitularCPF = String(STRING_CGC, 0)
    sClienteFat = String(STRING_CLIENTE_NOME_REDUZIDO, 0)
    sValidadeCC = String(STRING_TRP_VOU_VALIDADECC, 0)
    
    'Lê vouchers pagos com cartão de crédito
    lErro = Comando_Executar(alComando(0), sSQL, sCia, sNumCCred, dtDataAuto, lNumAuto, lNumVou, sSerie, sTipo, dtData, lClienteVou, sClienteVou, _
    dValor, dValorComissao, sNomePax, sSobrenomePax, sTitular, sTitularCPF, iCodSeguranca, lClienteFat, sClienteFat, iNumParc, dTaxa, sValidadeCC, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 200669
    
    'Busca o primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200670
        
    'Enquanto existirem arquivos
    Do While lErro = AD_SQL_SUCESSO
    
    
        'Se modificou o número do cartão de crédito, o administradora, a data ou o número de autorização deve gerar um novo título
        If sCiaAnt <> sCia Or sNumCCredAnt <> sNumCCred Or dtDataAuto <> dtDataAutoAnt Or lNumAutoAnt <> lNumAuto Then
            
            sCiaAnt = sCia
            sNumCCredAnt = sNumCCred
            dtDataAutoAnt = dtDataAuto
            lNumAutoAnt = lNumAuto
            
            Set objVouAutoCC = New ClassTRPVouAutoCC
            
            objVouAutoCC.dTarifa = dTaxa
            objVouAutoCC.dtDataAutoCC = dtDataAuto
            objVouAutoCC.iCodSeguranca = iCodSeguranca
            objVouAutoCC.iQuantParc = iNumParc
            objVouAutoCC.lClienteFat = lClienteFat
            objVouAutoCC.lNumAuto = lNumAuto
            objVouAutoCC.sBandeira = sCia
            objVouAutoCC.sClienteFat = sClienteFat
            objVouAutoCC.sNumCCred = sNumCCred
            objVouAutoCC.sTitular = sTitular
            objVouAutoCC.sTitularCPF = sTitularCPF
            objVouAutoCC.sValidadeCC = sValidadeCC
            objVouAutoCC.iFilialEmpresa = iFilialEmpresa
            
            colAuto.Add objVouAutoCC
            
        End If
    
        Set objVou = New ClassTRPFATInfoVoucher
        
        objVou.lNumVou = lNumVou
        objVou.sSerie = sSerie
        objVou.sTipoVou = sTipo
        objVou.dValor = dValor
        objVou.dValorComissao = dValorComissao
        objVou.lCliente = lClienteVou
        objVou.sNomeCliVou = sClienteVou
        objVou.sPassageiroNome = sNomePax
        objVou.sPassageiroSobreNome = sSobrenomePax
        objVou.dtDataEmissao = dtData
        
        objVouAutoCC.colVou.Add objVou
    
        'Busca o próximo arquivo
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200671
        
    Loop
    
    For Each objVouAutoCC In colAuto
        objVouAutoCC.dValorB = 0
        objVouAutoCC.dValorL = 0
        For Each objVou In objVouAutoCC.colVou
            objVouAutoCC.dValorB = objVouAutoCC.dValorB + objVou.dValor
            objVouAutoCC.dValorL = Arredonda_Moeda(objVouAutoCC.dValorL + (objVou.dValor * (1 - objVouAutoCC.dTarifa)))
        Next
    Next
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
        
    TRPFatCartao_Le = SUCESSO
     
    Exit Function
    
Erro_TRPFatCartao_Le:

    TRPFatCartao_Le = gErr
    
    Select Case gErr
          
        Case 200668
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 200669 To 200671
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPVOUCHERS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182722)
     
    End Select

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
        
    Exit Function

End Function


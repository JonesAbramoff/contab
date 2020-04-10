VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFATSelectTRV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function Comando_BindVarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_BindVar" (ByVal lComando As Long, lpVar As Variant) As Long
Private Declare Function Comando_PrepararInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Preparar" (ByVal lComando As Long, ByVal lpSQLStmt As String) As Long
Private Declare Function Comando_ExecutarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Executar" (ByVal lComando As Long) As Long
Private Declare Function Conexao_AbrirExt Lib "ADSQLMN.DLL" Alias "AD_Conexao_Abrir" (ByVal driver_sql As Integer, ByVal lpParamIn As String, ByVal ParamLenIn As Integer, ByVal lpParamOut As String, lpParamLenOut As Integer) As Long
Private Declare Function Conexao_FecharExt Lib "ADSQLMN.DLL" Alias "AD_Conexao_Fechar" (ByVal lConexao As Long) As Long


'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function TRVOcorrencias_Le(ByVal objTRVOcorrencias As ClassTRVOcorrencias) As Long

Dim lErro As Long
Dim lComando As Long
Dim tTRVOcorrencias As typeTRVOcorrencias

On Error GoTo Erro_TRVOcorrencias_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 189437

    'Alocação de espaço no buffer
    tTRVOcorrencias.sTipoDoc = String(STRING_TRV_OCR_TIPODOC, 0)
    tTRVOcorrencias.sSerie = String(STRING_TRV_OCR_SERIE, 0)
    tTRVOcorrencias.sObservacao = String(STRING_TRV_OCR_OBS, 0)
    tTRVOcorrencias.sHistorico = String(STRING_TRV_OCR_HISTORICO, 0)

    'Le a tabelaTRVOcorrencias
    If objTRVOcorrencias.lNumIntDoc <> 0 Then
        lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumVou, TipoDoc, Serie, Codigo, Cliente, " & _
                    "FilialCliente, DataEmissao, Observacao, Status, Origem, Historico, FormaPagto, " & _
                    "NumIntDocDestino, ValorTotal, NumTitulo, TipoDocDestino,ValorOCRBrutoVou,ValorOCRCMAVou FROM TRVOcorrencias WHERE NumIntDoc= ? ", _
                    tTRVOcorrencias.lNumIntDoc, tTRVOcorrencias.lNumVou, tTRVOcorrencias.sTipoDoc, tTRVOcorrencias.sSerie, _
                    tTRVOcorrencias.lCodigo, tTRVOcorrencias.lCliente, tTRVOcorrencias.iFilialCliente, tTRVOcorrencias.dtDataEmissao, tTRVOcorrencias.sObservacao, _
                    tTRVOcorrencias.iStatus, tTRVOcorrencias.iOrigem, tTRVOcorrencias.sHistorico, tTRVOcorrencias.iFormaPagto, tTRVOcorrencias.lNumIntDocDestino, _
                    tTRVOcorrencias.dValorTotal, tTRVOcorrencias.lNumDocDestino, tTRVOcorrencias.iTipoDocDestino, tTRVOcorrencias.dValorOCRBrutoVou, tTRVOcorrencias.dValorOCRCMAVou, _
                    objTRVOcorrencias.lNumIntDoc)
    ElseIf objTRVOcorrencias.lID <> 0 Then
        lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumVou, TipoDoc, Serie, Codigo, Cliente, " & _
                    "FilialCliente, DataEmissao, Observacao, Status, Origem, Historico, FormaPagto, " & _
                    "NumIntDocDestino, ValorTotal, NumTitulo, TipoDocDestino,ValorOCRBrutoVou,ValorOCRCMAVou FROM TRVOcorrencias WHERE ID = ? ", _
                    tTRVOcorrencias.lNumIntDoc, tTRVOcorrencias.lNumVou, tTRVOcorrencias.sTipoDoc, tTRVOcorrencias.sSerie, _
                    tTRVOcorrencias.lCodigo, tTRVOcorrencias.lCliente, tTRVOcorrencias.iFilialCliente, tTRVOcorrencias.dtDataEmissao, tTRVOcorrencias.sObservacao, _
                    tTRVOcorrencias.iStatus, tTRVOcorrencias.iOrigem, tTRVOcorrencias.sHistorico, tTRVOcorrencias.iFormaPagto, tTRVOcorrencias.lNumIntDocDestino, _
                    tTRVOcorrencias.dValorTotal, tTRVOcorrencias.lNumDocDestino, tTRVOcorrencias.iTipoDocDestino, tTRVOcorrencias.dValorOCRBrutoVou, tTRVOcorrencias.dValorOCRCMAVou, _
                    objTRVOcorrencias.lID)
    Else
        lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumVou, TipoDoc, Serie, Codigo, Cliente, " & _
                    "FilialCliente, DataEmissao, Observacao, Status, Origem, Historico, FormaPagto, " & _
                    "NumIntDocDestino, ValorTotal, NumTitulo, TipoDocDestino,ValorOCRBrutoVou,ValorOCRCMAVou FROM TRVOcorrencias WHERE Codigo= ? ", _
                    tTRVOcorrencias.lNumIntDoc, tTRVOcorrencias.lNumVou, tTRVOcorrencias.sTipoDoc, tTRVOcorrencias.sSerie, _
                    tTRVOcorrencias.lCodigo, tTRVOcorrencias.lCliente, tTRVOcorrencias.iFilialCliente, tTRVOcorrencias.dtDataEmissao, tTRVOcorrencias.sObservacao, _
                    tTRVOcorrencias.iStatus, tTRVOcorrencias.iOrigem, tTRVOcorrencias.sHistorico, tTRVOcorrencias.iFormaPagto, tTRVOcorrencias.lNumIntDocDestino, _
                    tTRVOcorrencias.dValorTotal, tTRVOcorrencias.lNumDocDestino, tTRVOcorrencias.iTipoDocDestino, tTRVOcorrencias.dValorOCRBrutoVou, tTRVOcorrencias.dValorOCRCMAVou, _
                    objTRVOcorrencias.lCodigo)
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 189438

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189439

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objTRVOcorrencias.lNumIntDoc = tTRVOcorrencias.lNumIntDoc
    objTRVOcorrencias.lNumVou = tTRVOcorrencias.lNumVou
    objTRVOcorrencias.sTipoDoc = tTRVOcorrencias.sTipoDoc
    objTRVOcorrencias.sSerie = tTRVOcorrencias.sSerie
    objTRVOcorrencias.lCodigo = tTRVOcorrencias.lCodigo
    objTRVOcorrencias.lCliente = tTRVOcorrencias.lCliente
    objTRVOcorrencias.iFilialCliente = tTRVOcorrencias.iFilialCliente
    objTRVOcorrencias.dtDataEmissao = tTRVOcorrencias.dtDataEmissao
    objTRVOcorrencias.sObservacao = tTRVOcorrencias.sObservacao
    objTRVOcorrencias.iStatus = tTRVOcorrencias.iStatus
    objTRVOcorrencias.iOrigem = tTRVOcorrencias.iOrigem
    objTRVOcorrencias.sHistorico = tTRVOcorrencias.sHistorico
    objTRVOcorrencias.iFormaPagto = tTRVOcorrencias.iFormaPagto
    objTRVOcorrencias.lNumIntDocDestino = tTRVOcorrencias.lNumIntDocDestino
    objTRVOcorrencias.dValorTotal = tTRVOcorrencias.dValorTotal
    objTRVOcorrencias.lNumDocDestino = tTRVOcorrencias.lNumDocDestino
    objTRVOcorrencias.iTipoDocDestino = tTRVOcorrencias.iTipoDocDestino
    objTRVOcorrencias.dValorOCRBrutoVou = tTRVOcorrencias.dValorOCRBrutoVou
    objTRVOcorrencias.dValorOCRCMAVou = tTRVOcorrencias.dValorOCRCMAVou
    
    lErro = TRVOcorrenciaDet_Le(objTRVOcorrencias)
    If lErro <> SUCESSO Then gError 189484

    'Fecha Comando
    Call Comando_Fechar(lComando)

    TRVOcorrencias_Le = SUCESSO

    Exit Function

Erro_TRVOcorrencias_Le:

    TRVOcorrencias_Le = gErr

    Select Case gErr

        Case 189437
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 189438, 189439
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVOCORRENCIAS", gErr)

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
Private Function TRVOcorrenciaDet_Le(ByVal objTRVOcorrencias As ClassTRVOcorrencias) As Long

Dim lErro As Long
Dim lComando As Long
Dim tTRVOcorrenciaDet As typeTRVOcorrenciaDet
Dim objTRVOcorrenciaDet As ClassTRVOcorrenciaDet

On Error GoTo Erro_TRVOcorrenciaDet_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 189460

    'Alocação de espaço no buffer

    'Le a tabelaTRVOcorrenciaDet
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntDocOcr, Tipo, Valor, Seq FROM TRVOcorrenciaDet WHERE NumIntDocOcr= ? ORDER BY Seq", _
                tTRVOcorrenciaDet.lNumIntDoc, tTRVOcorrenciaDet.lNumIntDocOCR, tTRVOcorrenciaDet.iTipo, tTRVOcorrenciaDet.dValor, tTRVOcorrenciaDet.iSeq, _
                objTRVOcorrencias.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 189461

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189462

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objTRVOcorrenciaDet = New ClassTRVOcorrenciaDet

        objTRVOcorrenciaDet.lNumIntDoc = tTRVOcorrenciaDet.lNumIntDoc
        objTRVOcorrenciaDet.lNumIntDocOCR = tTRVOcorrenciaDet.lNumIntDocOCR
        objTRVOcorrenciaDet.iTipo = tTRVOcorrenciaDet.iTipo
        objTRVOcorrenciaDet.dValor = tTRVOcorrenciaDet.dValor
        objTRVOcorrenciaDet.iSeq = tTRVOcorrenciaDet.iSeq
        
        objTRVOcorrencias.colDetalhamento.Add objTRVOcorrenciaDet

        'Busca Próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189463

    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    TRVOcorrenciaDet_Le = SUCESSO

    Exit Function

Erro_TRVOcorrenciaDet_Le:

    TRVOcorrenciaDet_Le = gErr

    Select Case gErr

        Case 189460
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 189461, 189462, 189463
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVOCORRENCIADET", gErr)
        
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
    
    If sFiltro = "" Then
        sSQL = "SELECT DISTINCT " & sCampo & " FROM " & sTabela & " ORDER BY " & sCampo
    Else
        sSQL = "SELECT DISTINCT " & sCampo & " FROM " & sTabela & " WHERE " & sFiltro & " ORDER BY " & sCampo
    End If

    'Le a tabelaTRVOcorrencias
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

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function VouchersTodos_Le(ByVal objImpCoinfoItemFat As ClassImpCoinfoItemFat) As Long

Dim lErro As Long
Dim lComando As Long
Dim tImpCoinfoItemFat As typeImpCoinfoItemFat

On Error GoTo Erro_VouchersTodos_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 190152

    'Alocação de espaço no buffer
    tImpCoinfoItemFat.sMoeda = String(255, 0)
    tImpCoinfoItemFat.scodest = String(255, 0)
    tImpCoinfoItemFat.sUsuario = String(255, 0)
    tImpCoinfoItemFat.sTipoDoc = String(255, 0)
    tImpCoinfoItemFat.sSerie = String(255, 0)
    tImpCoinfoItemFat.sTipVou = String(255, 0)
    tImpCoinfoItemFat.scodgru = String(255, 0)
    tImpCoinfoItemFat.sControle = String(255, 0)
    tImpCoinfoItemFat.sCodPro = String(255, 0)
    tImpCoinfoItemFat.sNumCCred = String(255, 0)
    tImpCoinfoItemFat.sValidCC = String(255, 0)
    tImpCoinfoItemFat.sCiaCart = String(255, 0)
    tImpCoinfoItemFat.sGrupo = String(255, 0)
    tImpCoinfoItemFat.sNumAuto = String(255, 0)

    'Le a tabelaVouchersTodos
    lErro = Comando_Executar(lComando, "SELECT NumIntArq, NumRegArqTexto, codemp, data, numvou, moeda, " & _
                "valor, codest, numfat, dataven, numdoc, usuario, datareg, " & _
                "marca, tipodoc, recnsica, condpag, Serie, TipVou, codgru, " & _
                "Controle, cartao, CodPro, NumCCred, ValidCC, QuantParc, VlCartao, " & _
                "NumAuto, CiaCart, Exportado, DataExp, Excluido, QtdPax, Grupo, diasantc, master, " & _
                "nummstr, NumIntDoc FROM VouchersTodos WHERE Numvou= ?  AND tipodoc= ?  AND Serie= ?  AND Empresa= ? AND TipVou = ? ", _
                tImpCoinfoItemFat.lNumIntArq, tImpCoinfoItemFat.lNumRegArqTexto, tImpCoinfoItemFat.lcodemp, tImpCoinfoItemFat.dtData, _
                tImpCoinfoItemFat.lNumVou, tImpCoinfoItemFat.sMoeda, tImpCoinfoItemFat.dValor, tImpCoinfoItemFat.scodest, tImpCoinfoItemFat.lNumFat, _
                tImpCoinfoItemFat.dtdataven, tImpCoinfoItemFat.lnumdoc, tImpCoinfoItemFat.sUsuario, tImpCoinfoItemFat.dtdatareg, tImpCoinfoItemFat.iMarca, _
                tImpCoinfoItemFat.sTipoDoc, tImpCoinfoItemFat.lrecnsica, tImpCoinfoItemFat.icondpag, tImpCoinfoItemFat.sSerie, tImpCoinfoItemFat.sTipVou, _
                tImpCoinfoItemFat.scodgru, tImpCoinfoItemFat.sControle, tImpCoinfoItemFat.iCartao, tImpCoinfoItemFat.sCodPro, tImpCoinfoItemFat.sNumCCred, _
                tImpCoinfoItemFat.sValidCC, tImpCoinfoItemFat.iQuantParc, tImpCoinfoItemFat.dVlCartao, tImpCoinfoItemFat.sNumAuto, tImpCoinfoItemFat.sCiaCart, _
                tImpCoinfoItemFat.iExportado, tImpCoinfoItemFat.dtDataExp, tImpCoinfoItemFat.iExcluido, tImpCoinfoItemFat.iQtdPax, tImpCoinfoItemFat.sGrupo, _
                tImpCoinfoItemFat.idiasantc, tImpCoinfoItemFat.imaster, tImpCoinfoItemFat.lnummstr, tImpCoinfoItemFat.lNumIntDoc, _
                objImpCoinfoItemFat.lNumVou, objImpCoinfoItemFat.sTipoDoc, objImpCoinfoItemFat.sSerie, glEmpresa, objImpCoinfoItemFat.sTipVou)
    If lErro <> AD_SQL_SUCESSO Then gError 190153

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190154

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objImpCoinfoItemFat.lNumIntArq = tImpCoinfoItemFat.lNumIntArq
    objImpCoinfoItemFat.lNumRegArqTexto = tImpCoinfoItemFat.lNumRegArqTexto
    objImpCoinfoItemFat.lcodemp = tImpCoinfoItemFat.lcodemp
    objImpCoinfoItemFat.dtData = tImpCoinfoItemFat.dtData
    objImpCoinfoItemFat.lNumVou = tImpCoinfoItemFat.lNumVou
    objImpCoinfoItemFat.sMoeda = tImpCoinfoItemFat.sMoeda
    objImpCoinfoItemFat.dValor = tImpCoinfoItemFat.dValor
    objImpCoinfoItemFat.scodest = tImpCoinfoItemFat.scodest
    objImpCoinfoItemFat.lNumFat = tImpCoinfoItemFat.lNumFat
    objImpCoinfoItemFat.dtdataven = tImpCoinfoItemFat.dtdataven
    objImpCoinfoItemFat.lnumdoc = tImpCoinfoItemFat.lnumdoc
    objImpCoinfoItemFat.sUsuario = tImpCoinfoItemFat.sUsuario
    objImpCoinfoItemFat.dtdatareg = tImpCoinfoItemFat.dtdatareg
    objImpCoinfoItemFat.iMarca = tImpCoinfoItemFat.iMarca
    objImpCoinfoItemFat.sTipoDoc = tImpCoinfoItemFat.sTipoDoc
    objImpCoinfoItemFat.lrecnsica = tImpCoinfoItemFat.lrecnsica
    objImpCoinfoItemFat.icondpag = tImpCoinfoItemFat.icondpag
    objImpCoinfoItemFat.sSerie = tImpCoinfoItemFat.sSerie
    objImpCoinfoItemFat.sTipVou = tImpCoinfoItemFat.sTipVou
    objImpCoinfoItemFat.scodgru = tImpCoinfoItemFat.scodgru
    objImpCoinfoItemFat.sControle = tImpCoinfoItemFat.sControle
    objImpCoinfoItemFat.iCartao = tImpCoinfoItemFat.iCartao
    objImpCoinfoItemFat.sCodPro = tImpCoinfoItemFat.sCodPro
    objImpCoinfoItemFat.sNumCCred = tImpCoinfoItemFat.sNumCCred
    objImpCoinfoItemFat.sValidCC = tImpCoinfoItemFat.sValidCC
    objImpCoinfoItemFat.iQuantParc = tImpCoinfoItemFat.iQuantParc
    objImpCoinfoItemFat.dVlCartao = tImpCoinfoItemFat.dVlCartao
    'objImpCoinfoItemFat.lNumAuto = tImpCoinfoItemFat.lNumAuto
    objImpCoinfoItemFat.sNumAuto = tImpCoinfoItemFat.sNumAuto
    objImpCoinfoItemFat.sCiaCart = tImpCoinfoItemFat.sCiaCart
    objImpCoinfoItemFat.iExportado = tImpCoinfoItemFat.iExportado
    objImpCoinfoItemFat.dtDataExp = tImpCoinfoItemFat.dtDataExp
    objImpCoinfoItemFat.iExcluido = tImpCoinfoItemFat.iExcluido
    objImpCoinfoItemFat.iQtdPax = tImpCoinfoItemFat.iQtdPax
    objImpCoinfoItemFat.sGrupo = tImpCoinfoItemFat.sGrupo
    objImpCoinfoItemFat.idiasantc = tImpCoinfoItemFat.idiasantc
    objImpCoinfoItemFat.imaster = tImpCoinfoItemFat.imaster
    objImpCoinfoItemFat.lnummstr = tImpCoinfoItemFat.lnummstr
    objImpCoinfoItemFat.lNumIntDoc = tImpCoinfoItemFat.lNumIntDoc

    'Fecha Comando
    Call Comando_Fechar(lComando)

    VouchersTodos_Le = SUCESSO

    Exit Function

Erro_VouchersTodos_Le:

    VouchersTodos_Le = gErr

    Select Case gErr

        Case 190152
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 190153, 190154
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_VOUCHERSTODOS", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190155)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function TRVOcorrencias_Le_Bloqueios(ByVal objTRVLiberaOCRSel As ClassTRVLiberaOCRSel) As Long
 
Dim lErro As Long
Dim iIndice As Integer
Dim sSQL As String
Dim sWhere As String
Dim vlCodigo As Variant
Dim viStatus As Variant
Dim alComando(1 To 1) As Long
Dim objOcorrencia As ClassTRVOcorrencias
Dim vlCodigoDe As Variant
Dim vlCodigoAte As Variant
Dim vdtEmissaoDe As Variant
Dim vdtEmissaoAte As Variant

On Error GoTo Erro_TRVOcorrencias_Le_Bloqueios
    
    vlCodigoDe = objTRVLiberaOCRSel.lCodigoDe
    vlCodigoAte = objTRVLiberaOCRSel.lCodigoAte
    vdtEmissaoDe = objTRVLiberaOCRSel.dtDataEmissaoDe
    vdtEmissaoAte = objTRVLiberaOCRSel.dtDataEmissaoAte
    vlCodigo = CLng(0)
    viStatus = STATUS_TRV_OCR_BLOQUEADO
 
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 190244
    Next
    
    sSQL = "SELECT O.Codigo FROM TRVOcorrencias AS O WHERE O.Status = ? "
    
    If objTRVLiberaOCRSel.iPago = TRV_VOU_NAO_PAGO Then
        sWhere = sWhere & "AND NOT EXISTS (SELECT T.NumIntDoc FROM TrvVouchers AS V, TitulosRecTodos AS T WHERE V.TipoDoc = 'VOU' AND V.TipoDocDestino =3 AND V.NumIntDocDestino = T.NumIntDoc AND T.Saldo + 0.01 < T.Valor AND O.NumVou = V.NumVou AND O.TipoDoc = V.TipVou AND O.Serie = V.Serie)"
    ElseIf objTRVLiberaOCRSel.iPago = TRV_VOU_PAGO Then
        sWhere = sWhere & "AND EXISTS (SELECT T.NumIntDoc FROM TrvVouchers AS V, TitulosRecTodos AS T WHERE V.TipoDoc = 'VOU' AND V.TipoDocDestino =3 AND V.NumIntDocDestino = T.NumIntDoc AND T.Saldo + 0.01 < T.Valor AND O.NumVou = V.NumVou AND O.TipoDoc = V.TipVou AND O.Serie = V.Serie)"
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
    
        Set objOcorrencia = New ClassTRVOcorrencias

        objOcorrencia.lCodigo = vlCodigo
        
        'Lê o produto
        lErro = TRVOcorrencias_Le(objOcorrencia)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 190254
        
        objTRVLiberaOCRSel.colOcorrenciais.Add objOcorrencia
     
        'Busca o próximo registro
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190255
 
    Loop
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
 
    TRVOcorrencias_Le_Bloqueios = SUCESSO
 
    Exit Function
 
Erro_TRVOcorrencias_Le_Bloqueios:
 
    TRVOcorrencias_Le_Bloqueios = gErr
 
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

Public Function TRVAportes_Le(ByVal objTRVAportes As ClassTRVAportes, Optional ByVal bLePagtos As Boolean = True) As Long

Dim lErro As Long
Dim alComando(0 To 5) As Long
Dim iIndice As Integer
Dim iProxCmd As Integer

On Error GoTo Erro_TRVAportes_Le

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 190376
    Next
    
    lErro = CF("TRVAportes_Le_Aux", alComando, objTRVAportes, iProxCmd, False, bLePagtos)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 190377
    
    If lErro = ERRO_LEITURA_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TRVAportes_Le = SUCESSO

    Exit Function

Erro_TRVAportes_Le:

    TRVAportes_Le = gErr

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

Public Function TRVAportes_Le_Aux(alComando() As Long, ByVal objTRVAportes As ClassTRVAportes, iProxCmd As Integer, Optional ByVal bLock As Boolean = False, Optional ByVal bLePagtos As Boolean = True) As Long

Dim lErro As Long
Dim tTRVAportes As typeTRVAportes
Dim lChave As Long
Dim sCampoChave As String

On Error GoTo Erro_TRVAportes_Le_Aux

    'Alocação de espaço no buffer
    tTRVAportes.sObservacao = String(STRING_TRV_APORTE_OBS, 0)
    tTRVAportes.sHistorico = String(STRING_TRV_APORTE_HISTORICO, 0)
    
    If objTRVAportes.lNumIntDoc <> 0 Then
        lChave = objTRVAportes.lNumIntDoc
        sCampoChave = "NumIntDoc"
    Else
        lChave = objTRVAportes.lCodigo
        sCampoChave = "Codigo"
    End If

    'Le a tabelaTRVAportes
    If bLock Then

        lErro = Comando_ExecutarPos(alComando(iProxCmd), "SELECT NumIntDoc, Codigo, Cliente, FilialCliente, DataEmissao, Observacao, " & _
                    "Historico, Tipo, Moeda, PrevValor, PrevDataDe, PrevDataAte, ProxParcela FROM TRVAportes WHERE " & sCampoChave & " = ? ", 0, _
                    tTRVAportes.lNumIntDoc, tTRVAportes.lCodigo, tTRVAportes.lCliente, tTRVAportes.iFilialCliente, _
                    tTRVAportes.dtDataEmissao, tTRVAportes.sObservacao, tTRVAportes.sHistorico, tTRVAportes.iTipo, tTRVAportes.iMoeda, _
                    tTRVAportes.dPrevValor, tTRVAportes.dtPrevDataDe, tTRVAportes.dtPrevDataAte, tTRVAportes.iProxParcela, _
                    lChave)
    Else
        lErro = Comando_Executar(alComando(iProxCmd), "SELECT NumIntDoc, Codigo, Cliente, FilialCliente, DataEmissao, Observacao, " & _
                    "Historico, Tipo, Moeda, PrevValor, PrevDataDe, PrevDataAte, ProxParcela FROM TRVAportes WHERE " & sCampoChave & " = ? ", _
                    tTRVAportes.lNumIntDoc, tTRVAportes.lCodigo, tTRVAportes.lCliente, tTRVAportes.iFilialCliente, _
                    tTRVAportes.dtDataEmissao, tTRVAportes.sObservacao, tTRVAportes.sHistorico, tTRVAportes.iTipo, tTRVAportes.iMoeda, _
                    tTRVAportes.dPrevValor, tTRVAportes.dtPrevDataDe, tTRVAportes.dtPrevDataAte, tTRVAportes.iProxParcela, _
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

    objTRVAportes.lNumIntDoc = tTRVAportes.lNumIntDoc
    objTRVAportes.lCodigo = tTRVAportes.lCodigo
    objTRVAportes.lCliente = tTRVAportes.lCliente
    objTRVAportes.iFilialCliente = tTRVAportes.iFilialCliente
    objTRVAportes.dtDataEmissao = tTRVAportes.dtDataEmissao
    objTRVAportes.sObservacao = tTRVAportes.sObservacao
    objTRVAportes.sHistorico = tTRVAportes.sHistorico
    objTRVAportes.iTipo = tTRVAportes.iTipo
    objTRVAportes.iMoeda = tTRVAportes.iMoeda
    objTRVAportes.dPrevValor = tTRVAportes.dPrevValor
    objTRVAportes.dtPrevDataDe = tTRVAportes.dtPrevDataDe
    objTRVAportes.dtPrevDataAte = tTRVAportes.dtPrevDataAte
    objTRVAportes.iProxParcela = tTRVAportes.iProxParcela
    
    If bLePagtos Then
    
        lErro = TRVAportePagtoCond_Le(alComando(iProxCmd), objTRVAportes, bLock)
        iProxCmd = iProxCmd + 1
        If lErro <> SUCESSO Then gError 190329
    
        lErro = TRVAportePagtoDireto_Le(alComando(iProxCmd), objTRVAportes, bLock)
        iProxCmd = iProxCmd + 1
        If lErro <> SUCESSO Then gError 190330
    
        lErro = TRVAportePagtoFat_Le(alComando(iProxCmd), objTRVAportes, bLock)
        iProxCmd = iProxCmd + 1
        If lErro <> SUCESSO Then gError 190331
        
        lErro = TRVAportePagtoFatCond_Le(alComando(iProxCmd), objTRVAportes, bLock)
        iProxCmd = iProxCmd + 1
        If lErro <> SUCESSO Then gError 190331

    End If

    TRVAportes_Le_Aux = SUCESSO

    Exit Function

Erro_TRVAportes_Le_Aux:

    TRVAportes_Le_Aux = gErr

    Select Case gErr

        Case 190326
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 190327, 190328
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVAPORTES", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora
        
        Case 190329 To 190321
        
        Case 190374
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRVAPORTES", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190332)

    End Select

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Private Function TRVAportePagtoCond_Le(ByVal lComando As Long, ByVal objTRVAporte As ClassTRVAportes, Optional ByVal bLock As Boolean = False) As Long

Dim lErro As Long
Dim tTRVAportePagtoCond As typeTRVAportePagtoCond
Dim objTRVAportePagtoCond As ClassTRVAportePagtoCond

On Error GoTo Erro_TRVAportePagtoCond_Le

    'Le a tabela TRVAportePagtoCond
    If bLock Then
        lErro = Comando_ExecutarPos(lComando, "SELECT NumIntDoc, NumIntDocAporte, Base, Status, DataPagto, NumIntDocDestino, " & _
                    "FormaPagto, TipoDocDestino, Percentual, Valor FROM TRVAportePagtoCond WHERE NumIntDocAporte = ? ORDER BY Seq ", 0, _
                    tTRVAportePagtoCond.lNumIntDoc, tTRVAportePagtoCond.lNumIntDocAporte, tTRVAportePagtoCond.iBase, tTRVAportePagtoCond.iStatus, _
                    tTRVAportePagtoCond.dtDataPagto, tTRVAportePagtoCond.lNumIntDocDestino, tTRVAportePagtoCond.iFormaPagto, tTRVAportePagtoCond.iTipoDocDestino, tTRVAportePagtoCond.dPercentual, tTRVAportePagtoCond.dValor, _
                    objTRVAporte.lNumIntDoc)
    Else
        lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntDocAporte, Base, Status, DataPagto, NumIntDocDestino, " & _
                    "FormaPagto, TipoDocDestino, Percentual, Valor FROM TRVAportePagtoCond WHERE NumIntDocAporte = ? ORDER BY Seq ", _
                    tTRVAportePagtoCond.lNumIntDoc, tTRVAportePagtoCond.lNumIntDocAporte, tTRVAportePagtoCond.iBase, tTRVAportePagtoCond.iStatus, _
                    tTRVAportePagtoCond.dtDataPagto, tTRVAportePagtoCond.lNumIntDocDestino, tTRVAportePagtoCond.iFormaPagto, tTRVAportePagtoCond.iTipoDocDestino, tTRVAportePagtoCond.dPercentual, tTRVAportePagtoCond.dValor, _
                    objTRVAporte.lNumIntDoc)
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
    
        Set objTRVAportePagtoCond = New ClassTRVAportePagtoCond

        objTRVAportePagtoCond.lNumIntDoc = tTRVAportePagtoCond.lNumIntDoc
        objTRVAportePagtoCond.lNumIntDocAporte = tTRVAportePagtoCond.lNumIntDocAporte
        objTRVAportePagtoCond.iBase = tTRVAportePagtoCond.iBase
        objTRVAportePagtoCond.iStatus = tTRVAportePagtoCond.iStatus
        objTRVAportePagtoCond.dtDataPagto = tTRVAportePagtoCond.dtDataPagto
        objTRVAportePagtoCond.lNumIntDocDestino = tTRVAportePagtoCond.lNumIntDocDestino
        objTRVAportePagtoCond.iFormaPagto = tTRVAportePagtoCond.iFormaPagto
        objTRVAportePagtoCond.iTipoDocDestino = tTRVAportePagtoCond.iTipoDocDestino
        objTRVAportePagtoCond.dPercentual = tTRVAportePagtoCond.dPercentual
        objTRVAportePagtoCond.dValor = tTRVAportePagtoCond.dValor
        
        objTRVAporte.colPagtoCondicionados.Add objTRVAportePagtoCond
        
        'Busca Próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190336
        
    Loop

    TRVAportePagtoCond_Le = SUCESSO

    Exit Function

Erro_TRVAportePagtoCond_Le:

    TRVAportePagtoCond_Le = gErr

    Select Case gErr
        
        Case 190334, 190335, 190336
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVAPORTEPAGTOCOND", gErr)
        
        Case 190379
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRVAPORTEPAGTOCOND", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190337)

    End Select

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Private Function TRVAportePagtoDireto_Le(ByVal lComando As Long, ByVal objTRVAporte As ClassTRVAportes, Optional ByVal bLock As Boolean = False) As Long

Dim lErro As Long
Dim tTRVAportePagtoDireto As typeTRVAportePagtoDireto
Dim objTRVAportePagtoDireto As ClassTRVAportePagtoDireto

On Error GoTo Erro_TRVAportePagtoDireto_Le

    'Le a tabelaTRVAportePagtoDireto
    If bLock Then
        lErro = Comando_ExecutarPos(lComando, "SELECT NumIntDoc, NumIntDocAporte, Valor, Vencimento, NumIntDocDestino, FormaPagto, " & _
                    "TipoDocDestino FROM TRVAportePagtoDireto WHERE NumIntDocAporte= ? ORDER BY Seq ", 0, _
                    tTRVAportePagtoDireto.lNumIntDoc, tTRVAportePagtoDireto.lNumIntDocAporte, tTRVAportePagtoDireto.dValor, tTRVAportePagtoDireto.dtVencimento, _
                    tTRVAportePagtoDireto.lNumIntDocDestino, tTRVAportePagtoDireto.iFormaPagto, tTRVAportePagtoDireto.iTipoDocDestino, _
                    objTRVAporte.lNumIntDoc)
    Else
        lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntDocAporte, Valor, Vencimento, NumIntDocDestino, FormaPagto, " & _
                    "TipoDocDestino FROM TRVAportePagtoDireto WHERE NumIntDocAporte= ? ORDER BY Seq ", _
                    tTRVAportePagtoDireto.lNumIntDoc, tTRVAportePagtoDireto.lNumIntDocAporte, tTRVAportePagtoDireto.dValor, tTRVAportePagtoDireto.dtVencimento, _
                    tTRVAportePagtoDireto.lNumIntDocDestino, tTRVAportePagtoDireto.iFormaPagto, tTRVAportePagtoDireto.iTipoDocDestino, _
                    objTRVAporte.lNumIntDoc)
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
    
        Set objTRVAportePagtoDireto = New ClassTRVAportePagtoDireto

        objTRVAportePagtoDireto.lNumIntDoc = tTRVAportePagtoDireto.lNumIntDoc
        objTRVAportePagtoDireto.lNumIntDocAporte = tTRVAportePagtoDireto.lNumIntDocAporte
        objTRVAportePagtoDireto.dValor = tTRVAportePagtoDireto.dValor
        objTRVAportePagtoDireto.dtVencimento = tTRVAportePagtoDireto.dtVencimento
        objTRVAportePagtoDireto.lNumIntDocDestino = tTRVAportePagtoDireto.lNumIntDocDestino
        objTRVAportePagtoDireto.iFormaPagto = tTRVAportePagtoDireto.iFormaPagto
        objTRVAportePagtoDireto.iTipoDocDestino = tTRVAportePagtoDireto.iTipoDocDestino

        objTRVAporte.colPagtoDireto.Add objTRVAportePagtoDireto
    
        'Busca Próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190341
    
    Loop

    TRVAportePagtoDireto_Le = SUCESSO

    Exit Function

Erro_TRVAportePagtoDireto_Le:

    TRVAportePagtoDireto_Le = gErr

    Select Case gErr

        Case 190339, 190340, 190341
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVAPORTEPAGTODIRETO", gErr)

        Case 190380
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRVAPORTEPAGTODIRETO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190342)

    End Select

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Private Function TRVAportePagtoFat_Le(ByVal lComando As Long, ByVal objTRVAporte As ClassTRVAportes, Optional ByVal bLock As Boolean = False) As Long

Dim lErro As Long
Dim tTRVAportePagtoFat As typeTRVAportePagtoFat
Dim objTRVAportePagtoFat As ClassTRVAportePagtoFat

On Error GoTo Erro_TRVAportePagtoFat_Le

    'Le a tabelaTRVAportePagtoFat
    If bLock Then
        lErro = Comando_ExecutarPos(lComando, "SELECT NumIntDoc, NumIntDocAporte, Valor, ValidadeDe, ValidadeAte, Saldo, " & _
                    "Percentual FROM TRVAportePagtoFat WHERE NumIntDocAporte= ? ORDER BY Seq ", 0, _
                    tTRVAportePagtoFat.lNumIntDoc, tTRVAportePagtoFat.lNumIntDocAporte, tTRVAportePagtoFat.dValor, tTRVAportePagtoFat.dtValidadeDe, _
                    tTRVAportePagtoFat.dtValidadeAte, tTRVAportePagtoFat.dSaldo, tTRVAportePagtoFat.dPercentual, _
                    objTRVAporte.lNumIntDoc)
    Else
        lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntDocAporte, Valor, ValidadeDe, ValidadeAte, Saldo, " & _
                    "Percentual FROM TRVAportePagtoFat WHERE NumIntDocAporte= ? ORDER BY Seq ", _
                    tTRVAportePagtoFat.lNumIntDoc, tTRVAportePagtoFat.lNumIntDocAporte, tTRVAportePagtoFat.dValor, tTRVAportePagtoFat.dtValidadeDe, _
                    tTRVAportePagtoFat.dtValidadeAte, tTRVAportePagtoFat.dSaldo, tTRVAportePagtoFat.dPercentual, _
                    objTRVAporte.lNumIntDoc)
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
    
        Set objTRVAportePagtoFat = New ClassTRVAportePagtoFat

        objTRVAportePagtoFat.lNumIntDoc = tTRVAportePagtoFat.lNumIntDoc
        objTRVAportePagtoFat.lNumIntDocAporte = tTRVAportePagtoFat.lNumIntDocAporte
        objTRVAportePagtoFat.dValor = tTRVAportePagtoFat.dValor
        objTRVAportePagtoFat.dtValidadeDe = tTRVAportePagtoFat.dtValidadeDe
        objTRVAportePagtoFat.dtValidadeAte = tTRVAportePagtoFat.dtValidadeAte
        objTRVAportePagtoFat.dSaldo = tTRVAportePagtoFat.dSaldo
        objTRVAportePagtoFat.dPercentual = tTRVAportePagtoFat.dPercentual
        
        objTRVAporte.colPagtoSobreFatura.Add objTRVAportePagtoFat

        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190346

    Loop

    TRVAportePagtoFat_Le = SUCESSO

    Exit Function

Erro_TRVAportePagtoFat_Le:

    TRVAportePagtoFat_Le = gErr

    Select Case gErr

        Case 190344, 190345, 190346
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVAPORTEPAGTOFAT", gErr)

        Case 190381
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRVAPORTEPAGTOFAT", gErr)

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

    'Le a tabelaTRVOcorrencias
    lErro = Comando_Executar(lComando, "SELECT SUM(V.Valor), SUM(V.Valor / V.Cambio) FROM TRVVouchers AS V, FiliaisClientesTRV AS C WHERE C.CodCliente = V.Cliente AND C.EmpresaPai = ? AND V.Data BETWEEN ? AND ? AND  V.Status <> ? AND V.TipoDoc = ? AND C.Aporte = ? AND V.Cambio<>0 ", dValor, dValorUSS, lCliente, dtDataDe, dtDataAte, STATUS_CANCELADO, TRV_TIPODOC_VOU_TEXTO, MARCADO)
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
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVVOUCHERS", gErr)

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
    
    sTipVou = String(STRING_TRV_OCR_TIPOVOU, 0)
    sSerie = String(STRING_TRV_OCR_SERIE, 0)

    'Le a tabela o realizado de vendas
    lErro = Comando_Executar(lComando, "SELECT V.TipVou, V.Serie, V.NumVou, V.Valor, V.Cambio, V.Data FROM TRVVouchers AS V, FiliaisClientesTRV AS C WHERE V.ClienteVou = C.CodCliente AND V.Data BETWEEN ? AND ? AND  V.Status <> ? AND V.TipoDoc = ? AND V.Cambio<>0 AND C.EmpresaPai = ? AND  C.Aporte = ? ORDER BY V.Data, V.NumVou", sTipVou, sSerie, lNumVou, dValor, dCambio, dtDataVou, dtDataPrevDe, dtDataPrevAte, STATUS_CANCELADO, TRV_TIPODOC_VOU_TEXTO, lCliente, MARCADO)
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
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVVOUCHERS", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190712)

    End Select

    If lComandoP = 0 Then
        'Fecha Comando
        Call Comando_Fechar(lComando)
    End If

    Exit Function

End Function

Public Function TRVAportePagtoCond_Le_NumIntDoc(ByVal lComando As Long, ByVal objTRVAportePagto As ClassTRVAportePagtoCond, Optional ByVal bLock As Boolean = False) As Long

Dim lErro As Long
Dim tTRVAportePagtoCond As typeTRVAportePagtoCond

On Error GoTo Erro_TRVAportePagtoCond_Le_NumIntDoc

    'Le a tabela TRVAportePagtoCond
    If bLock Then
        lErro = Comando_ExecutarPos(lComando, "SELECT NumIntDoc, NumIntDocAporte, Base, Status, DataPagto, NumIntDocDestino, " & _
                    "FormaPagto, TipoDocDestino, Percentual, Valor FROM TRVAportePagtoCond WHERE NumIntDoc = ? ", 0, _
                    tTRVAportePagtoCond.lNumIntDoc, tTRVAportePagtoCond.lNumIntDocAporte, tTRVAportePagtoCond.iBase, tTRVAportePagtoCond.iStatus, _
                    tTRVAportePagtoCond.dtDataPagto, tTRVAportePagtoCond.lNumIntDocDestino, tTRVAportePagtoCond.iFormaPagto, tTRVAportePagtoCond.iTipoDocDestino, tTRVAportePagtoCond.dPercentual, tTRVAportePagtoCond.dValor, _
                    objTRVAportePagto.lNumIntDoc)
    Else
        lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntDocAporte, Base, Status, DataPagto, NumIntDocDestino, " & _
                    "FormaPagto, TipoDocDestino, Percentual, Valor FROM TRVAportePagtoCond WHERE NumIntDoc = ? ", _
                    tTRVAportePagtoCond.lNumIntDoc, tTRVAportePagtoCond.lNumIntDocAporte, tTRVAportePagtoCond.iBase, tTRVAportePagtoCond.iStatus, _
                    tTRVAportePagtoCond.dtDataPagto, tTRVAportePagtoCond.lNumIntDocDestino, tTRVAportePagtoCond.iFormaPagto, tTRVAportePagtoCond.iTipoDocDestino, tTRVAportePagtoCond.dPercentual, tTRVAportePagtoCond.dValor, _
                    objTRVAportePagto.lNumIntDoc)
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

    objTRVAportePagto.lNumIntDoc = tTRVAportePagtoCond.lNumIntDoc
    objTRVAportePagto.lNumIntDocAporte = tTRVAportePagtoCond.lNumIntDocAporte
    objTRVAportePagto.iBase = tTRVAportePagtoCond.iBase
    objTRVAportePagto.iStatus = tTRVAportePagtoCond.iStatus
    objTRVAportePagto.dtDataPagto = tTRVAportePagtoCond.dtDataPagto
    objTRVAportePagto.lNumIntDocDestino = tTRVAportePagtoCond.lNumIntDocDestino
    objTRVAportePagto.iFormaPagto = tTRVAportePagtoCond.iFormaPagto
    objTRVAportePagto.iTipoDocDestino = tTRVAportePagtoCond.iTipoDocDestino
    objTRVAportePagto.dPercentual = tTRVAportePagtoCond.dPercentual
    objTRVAportePagto.dValor = tTRVAportePagtoCond.dValor
    
    TRVAportePagtoCond_Le_NumIntDoc = SUCESSO

    Exit Function

Erro_TRVAportePagtoCond_Le_NumIntDoc:

    TRVAportePagtoCond_Le_NumIntDoc = gErr

    Select Case gErr
        
        Case 190720, 190721
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVAPORTEPAGTOCOND", gErr)
        
        Case 190722
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRVAPORTEPAGTOCOND", gErr)
        
        Case ERRO_LEITURA_SEM_DADOS
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190723)

    End Select

    Exit Function

End Function

Public Function TRVAportes_Le_Bloqueios(ByVal objTRVLiberaAporteSel As ClassTRVLiberaAporteSel) As Long
 
Dim lErro As Long
Dim iIndice As Integer
Dim sSQL As String
Dim sWhere As String
Dim vlNumIntDoc As Variant
Dim viStatus As Variant
Dim alComando(1 To 2) As Long
Dim objAportePagto As ClassTRVAportePagtoCond
Dim vlCodigoDe As Variant
Dim vlCodigoAte As Variant
Dim vdtPagtoDe As Variant
Dim vdtPagtoAte As Variant

On Error GoTo Erro_TRVAportes_Le_Bloqueios
    
    vlCodigoDe = objTRVLiberaAporteSel.lCodigoDe
    vlCodigoAte = objTRVLiberaAporteSel.lCodigoAte
    vdtPagtoDe = objTRVLiberaAporteSel.dtDataPagtoDe
    vdtPagtoAte = objTRVLiberaAporteSel.dtDataPagtoAte
    vlNumIntDoc = CLng(0)
    viStatus = STATUS_TRV_OCR_BLOQUEADO
 
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 190724
    Next
    
    sSQL = "SELECT P.NumIntDoc FROM TRVAportes AS A, TRVAportePagtoCond AS P  WHERE A.NumIntDoc = P.NumIntDocAporte AND Status = ? "
 
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
    
        Set objAportePagto = New ClassTRVAportePagtoCond

        objAportePagto.lNumIntDoc = vlNumIntDoc
        
        'Lê o pagamento
        lErro = CF("TRVAportePagtoCond_Le_NumIntDoc", alComando(2), objAportePagto, False)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 190735
        
        objTRVLiberaAporteSel.colAportes.Add objAportePagto
     
        'Busca o próximo registro
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190734
 
    Loop
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
 
    TRVAportes_Le_Bloqueios = SUCESSO
 
    Exit Function
 
Erro_TRVAportes_Le_Bloqueios:
 
    TRVAportes_Le_Bloqueios = gErr
 
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
    
    objCombo.AddItem FORMAPAGTO_TRV_OCR_FAT & SEPARADOR & FORMAPAGTO_TRV_OCR_FAT_TEXTO
    objCombo.ItemData(objCombo.NewIndex) = FORMAPAGTO_TRV_OCR_FAT

    objCombo.AddItem FORMAPAGTO_TRV_OCR_CRED & SEPARADOR & FORMAPAGTO_TRV_OCR_CRED_TEXTO
    objCombo.ItemData(objCombo.NewIndex) = FORMAPAGTO_TRV_OCR_CRED

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

    objCombo.AddItem BASE_TRV_APORTE_REAL & SEPARADOR & BASE_TRV_APORTE_REAL_TEXTO
    objCombo.ItemData(objCombo.NewIndex) = BASE_TRV_APORTE_REAL

    objCombo.AddItem BASE_TRV_APORTE_REALACIMAPREV & SEPARADOR & BASE_TRV_APORTE_REALACIMAPREV_TEXTO
    objCombo.ItemData(objCombo.NewIndex) = BASE_TRV_APORTE_REALACIMAPREV

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

    objCombo.AddItem TRV_TIPODOC_VOU_TEXTO
    objCombo.ItemData(objCombo.NewIndex) = TRV_TIPODOC_VOU

    objCombo.AddItem TRV_TIPODOC_OCR_TEXTO
    objCombo.ItemData(objCombo.NewIndex) = TRV_TIPODOC_OCR

    objCombo.AddItem TRV_TIPODOC_NVL_TEXTO
    objCombo.ItemData(objCombo.NewIndex) = TRV_TIPODOC_NVL

    objCombo.AddItem TRV_TIPODOC_CMR_TEXTO
    objCombo.ItemData(objCombo.NewIndex) = TRV_TIPODOC_CMR

    objCombo.AddItem TRV_TIPODOC_CMC_TEXTO
    objCombo.ItemData(objCombo.NewIndex) = TRV_TIPODOC_CMC
    
    If giVersaoTRV >= TRV_VERSAO_COMISSIONAMENTO Then

        objCombo.AddItem TRV_TIPODOC_OVER_TEXTO
        objCombo.ItemData(objCombo.NewIndex) = TRV_TIPODOC_OVER
        
        objCombo.AddItem TRV_TIPODOC_CMCC_TEXTO
        objCombo.ItemData(objCombo.NewIndex) = TRV_TIPODOC_CMCC
    
    End If

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

Public Function TRVFaturamento_Le_Dados(ByVal objFaturamento As ClassTRVFaturamento) As Long
'Le os itens a serem faturados com os filtros passados
 
Dim lErro As Long
Dim iIndice As Integer
Dim sSQL As String
Dim sWhere As String
Dim sWhereAux As String
Dim alComando(1 To 12) As Long
Dim vValor As Variant
Dim vValorAux(1 To 200) As Variant
Dim vlNumIntDoc As Variant
Dim vsTipoDoc As Variant
Dim vsSerie As Variant
Dim vlCliente As Variant
Dim viCondPagto As Variant
Dim vdValor As Variant
Dim vdtData As Variant
Dim vsCodGrupo As Variant
Dim vlNumVou As Variant
Dim vlNumero As Variant
Dim vlClienteAporte As Variant
Dim viAporte As Variant
Dim viFilialCoinfo As Variant
Dim vsNome As Variant
Dim vsTipoVou As Variant
Dim objCliente As ClassCliente
Dim objFatCondPagto  As ClassTRVFatCondPagto
Dim objCondPagto  As ClassCondicaoPagto
Dim objInfoVoucher As ClassTRVFATInfoVoucher
Dim iProx As Integer
Dim objVoucher As ClassTRVVouchers
Dim viTipoGeracao As Variant

Dim vdCambio As Variant
Dim vdValorCambio As Variant
Dim viPax As Variant
Dim vsControle As Variant
Dim vsMoeda As Variant
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
Dim vdValorComisVou As Variant
Dim viPF As Variant
Dim bTemVou As Boolean
Dim sConteudo As String
Dim bArq1Aberto As Boolean, bGeraLog As Boolean, sDirLog As String, lCount As Long
Dim dtDataDeAux As Date, dtDataAteAux As Date, bDataUnica As Boolean

On Error GoTo Erro_TRVFaturamento_Le_Dados
 
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 190921
    Next
    
    bArq1Aberto = False

    lErro = CF("TRVConfig_Le", TRVCONFIG_GERA_LOG, EMPRESA_TODA, sConteudo)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If StrParaInt(sConteudo) = MARCADO Then
        bGeraLog = True
    
        lErro = CF("TRVConfig_Le", TRVCONFIG_DIR_LOG, EMPRESA_TODA, sDirLog)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        Open sDirLog & "FAT_LE_" & gsUsuario & "_" & Format(Date, "yyyymmdd") & "_" & Format(Now, "hhmmss") & ".txt" For Output As #1
        Open sDirLog & "DET_FAT_LE_" & gsUsuario & "_" & Format(Date, "yyyymmdd") & "_" & Format(Now, "hhmmss") & ".txt" For Output As #2
        bArq1Aberto = True
    Else
        bGeraLog = False
    End If
    
If bGeraLog Then Print #1, "1 - INICIANDO O PROCESSO " & CStr(Now)
    
    bTemVou = False
    
    vlNumIntDoc = CLng(0)
    vsTipoDoc = String(STRING_TRV_OCR_TIPODOC, 0)
    vsTipoVou = String(STRING_TRV_OCR_TIPOVOU, 0)
    vsSerie = String(STRING_TRV_OCR_SERIE, 0)
    vlCliente = CLng(0)
    vdValor = CDbl(0)
    vdtData = DATA_NULA
    vsCodGrupo = String(STRING_TRV_OCR_CODGRUPO, 0)
    vlNumVou = CLng(0)
    vlNumero = CLng(0)
    viFilialCoinfo = CInt(0)
    viCondPagto = CInt(0)
    vlClienteAporte = CLng(0)
    viAporte = CInt(0)
    vsNome = String(STRING_CLIENTE_NOME_REDUZIDO, 0)
    viTipoGeracao = CInt(0)
    
    vdCambio = CDbl(0)
    vdValorCambio = CDbl(0)
    viPax = CInt(0)
    vsControle = String(STRING_TRV_VOU_CONTROLE, 0)
    vsMoeda = String(STRING_TRV_VOU_MOEADA, 0)
    vsProduto = String(STRING_PRODUTO, 0)
    vdValorAporte = CDbl(0)
    viCartao = CInt(0)
    vsTitular = String(STRING_TRV_VOU_TITULAR, 0)
    vsPassageiroNome = String(STRING_TRV_VOU_PAXNOME, 0)
    vsPassageiroSobrenome = String(STRING_TRV_VOU_PAXNOME, 0)
    vlEmissor = CLng(0)
    vlClienteVou = CLng(0)
    vlCliPassageiro = CLng(0)
    viFilialEmpresa = CInt(0)
    viFilialFat = CInt(0)
    vsNomeCliVou = String(STRING_CLIENTE_NOME_REDUZIDO, 0)
    vsNomeEmissor = String(STRING_FORNECEDOR_RAZAO_SOC, 0)
    viFaturaJuntoOCR = CInt(0)
    viUmVouPorFat = CInt(0)
    viTipoCliente = CInt(0)
    viTipoOCR = CInt(0)
    vdValorVou = CDbl(0)
    vdValorComisVou = CDbl(0)
    viPF = CInt(0)
    
    If objFaturamento.iEmpresa = TRV_EMPRESA_AMBOS Then
        sSQL = "SELECT NumIntDoc, TipVou, TipoDoc, Serie , Cliente, Valor, Data, CodGrupo, NumVou, Numero, FilialCoinfo,CondPagto, EmpPai, Aporte, NomeCliente, " & _
        "Cambio, ValorCambio, Pax, Controle, Moeda, Produto, ValorAporte, Cartao, Titular, PassageiroNome, PassageiroSobrenome, Emissor, ClienteVou, CliPassageiro, FilialEmpresa, FilialFat, NomeClienteVou, NomeEmissor,FaturaJuntoOCR,UmVouPorFat, TipoCliente,TipoGeracao,TipoOCR,ValorVou,ValorComisVou, PF FROM DocsParaFat WHERE 1 = 1 "
    ElseIf objFaturamento.iEmpresa = TRV_EMPRESA_MY Then
        sSQL = "SELECT NumIntDoc, TipVou, TipoDoc, Serie , Cliente, Valor, Data, CodGrupo, NumVou, Numero, FilialCoinfo,CondPagto, EmpPai, Aporte, NomeCliente, " & _
        "Cambio, ValorCambio, Pax, Controle, Moeda, Produto, ValorAporte, Cartao, Titular, PassageiroNome, PassageiroSobrenome, Emissor, ClienteVou, CliPassageiro, FilialEmpresa, FilialFat, NomeClienteVou, NomeEmissor,FaturaJuntoOCR,UmVouPorFat, TipoCliente,TipoGeracao,TipoOCR,ValorVou,ValorComisVou, PF FROM DocsParaFat WHERE Cliente >= 1000000 "
    Else
        sSQL = "SELECT NumIntDoc, TipVou, TipoDoc, Serie , Cliente, Valor, Data, CodGrupo, NumVou, Numero, FilialCoinfo,CondPagto, EmpPai, Aporte, NomeCliente, " & _
        "Cambio, ValorCambio, Pax, Controle, Moeda, Produto, ValorAporte, Cartao, Titular, PassageiroNome, PassageiroSobrenome, Emissor, ClienteVou, CliPassageiro, FilialEmpresa, FilialFat, NomeClienteVou, NomeEmissor,FaturaJuntoOCR,UmVouPorFat, TipoCliente,TipoGeracao,TipoOCR,ValorVou,ValorComisVou, PF FROM DocsParaFat WHERE Cliente < 1000000 "
    End If
    iProx = 1

    If objFaturamento.colFiliais.Count > 0 Then
        sWhere = sWhere & "AND ("
        sWhereAux = ""
        For Each vValor In objFaturamento.colFiliais
            If vValor <> 0 Then
                If Len(Trim(sWhereAux)) > 0 Then sWhereAux = sWhereAux & " AND "
                sWhereAux = sWhereAux & "FilialCoinfo <> ? "
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
    
    '15/04/2010
    If objFaturamento.colTiposDoc.Count > 0 Then
        sWhere = sWhere & "AND ("
        sWhereAux = ""
        For Each vValor In objFaturamento.colTiposDoc
            If Len(Trim(sWhereAux)) > 0 Then sWhereAux = sWhereAux & " OR "
            sWhereAux = sWhereAux & " TipoDoc = ? "
        Next
        sWhere = sWhere & sWhereAux & ") "
    End If
    
    dtDataDeAux = DATA_NULA
    dtDataAteAux = DATA_NULA
    bDataUnica = True
    For Each objFatCondPagto In objFaturamento.colCondPagtos
        If dtDataDeAux = DATA_NULA Then dtDataDeAux = objFatCondPagto.dtDataVouDe
        If dtDataAteAux = DATA_NULA Then dtDataAteAux = objFatCondPagto.dtDataVouAte
        If dtDataDeAux <> objFatCondPagto.dtDataVouDe Or dtDataAteAux <> objFatCondPagto.dtDataVouAte Then
            bDataUnica = False
            Exit For
        End If
    Next
    If bDataUnica Then
        sWhere = sWhere & " AND CondPagto <> 0 "
        If dtDataDeAux <> DATA_NULA Then
            sWhere = sWhere & "AND Data >= ? "
        End If
        If dtDataAteAux <> DATA_NULA Then
            sWhere = sWhere & "AND Data <= ? "
        End If
    Else
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
    
    sSQL = sSQL & sWhere '& " ORDER BY Data"
    
If bGeraLog Then Print #1, "      SQL: " & sSQL
    
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
    
    lErro = Comando_BindVarInt(alComando(1), vsCodGrupo)
    If (lErro <> AD_SQL_SUCESSO) Then gError 190929
    
    lErro = Comando_BindVarInt(alComando(1), vlNumVou)
    If (lErro <> AD_SQL_SUCESSO) Then gError 190930
    
    lErro = Comando_BindVarInt(alComando(1), vlNumero)
    If (lErro <> AD_SQL_SUCESSO) Then gError 190931
    
    lErro = Comando_BindVarInt(alComando(1), viFilialCoinfo)
    If (lErro <> AD_SQL_SUCESSO) Then gError 190932
    
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
    
    lErro = Comando_BindVarInt(alComando(1), vsMoeda)
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
    
    lErro = Comando_BindVarInt(alComando(1), viTipoGeracao)
    If (lErro <> AD_SQL_SUCESSO) Then gError 190975
    
    lErro = Comando_BindVarInt(alComando(1), viTipoOCR)
    If (lErro <> AD_SQL_SUCESSO) Then gError 190975
    
    lErro = Comando_BindVarInt(alComando(1), vdValorVou)
    If (lErro <> AD_SQL_SUCESSO) Then gError 190975
    
    lErro = Comando_BindVarInt(alComando(1), vdValorComisVou)
    If (lErro <> AD_SQL_SUCESSO) Then gError 190975
    
    lErro = Comando_BindVarInt(alComando(1), viPF)
    If (lErro <> AD_SQL_SUCESSO) Then gError 190975
    
    For Each vValor In objFaturamento.colFiliais
        If vValor <> 0 Then
        
            iProx = iProx + 1
            vValorAux(iProx) = Filial_Corporator_Retorna_Coinfo(vValor)
        
            lErro = Comando_BindVarInt(alComando(1), vValorAux(iProx))
            If (lErro <> AD_SQL_SUCESSO) Then gError 190934
If bGeraLog Then Print #1, "         PARAM" & CStr(iProx - 1) & ": " & vValorAux(iProx)
            
        End If
    Next
    
    For Each objCliente In objFaturamento.colExcClientes
    
        iProx = iProx + 1
        vValorAux(iProx) = objCliente.lCodigo
        
        lErro = Comando_BindVarInt(alComando(1), vValorAux(iProx))
        If (lErro <> AD_SQL_SUCESSO) Then gError 190935
If bGeraLog Then Print #1, "         PARAM" & CStr(iProx - 1) & ": " & vValorAux(iProx)
        
    Next
    
    For Each objVoucher In objFaturamento.colExcVouchers
    
        iProx = iProx + 1
        vValorAux(iProx) = objVoucher.sTipVou
        
        lErro = Comando_BindVarInt(alComando(1), vValorAux(iProx))
        If (lErro <> AD_SQL_SUCESSO) Then gError 190935
        
If bGeraLog Then Print #1, "         PARAM" & CStr(iProx - 1) & ": " & vValorAux(iProx)
        
        iProx = iProx + 1
        vValorAux(iProx) = objVoucher.sSerie
        
        lErro = Comando_BindVarInt(alComando(1), vValorAux(iProx))
        If (lErro <> AD_SQL_SUCESSO) Then gError 190935
        
If bGeraLog Then Print #1, "         PARAM" & CStr(iProx) & ": " & vValorAux(iProx)
        
        iProx = iProx + 1
        vValorAux(iProx) = objVoucher.lNumVou
        
        lErro = Comando_BindVarInt(alComando(1), vValorAux(iProx))
        If (lErro <> AD_SQL_SUCESSO) Then gError 190935
        
If bGeraLog Then Print #1, "         PARAM" & CStr(iProx - 1) & ": " & vValorAux(iProx)
        
    Next
    
    '15/04/2010
    For Each vValor In objFaturamento.colTiposDoc
    
        iProx = iProx + 1
        vValorAux(iProx) = vValor
    
        lErro = Comando_BindVarInt(alComando(1), vValorAux(iProx))
        If (lErro <> AD_SQL_SUCESSO) Then gError 190939
        
If bGeraLog Then Print #1, "         PARAM" & CStr(iProx - 1) & ": " & vValorAux(iProx)
        
    Next
    
    If bDataUnica Then

        If dtDataDeAux <> DATA_NULA Then
        
            iProx = iProx + 1
            vValorAux(iProx) = dtDataDeAux
            
            lErro = Comando_BindVarInt(alComando(1), vValorAux(iProx))
            If (lErro <> AD_SQL_SUCESSO) Then gError 190937

If bGeraLog Then Print #1, "         PARAM" & CStr(iProx - 1) & ": " & vValorAux(iProx)

        End If
        If dtDataAteAux <> DATA_NULA Then
            
            iProx = iProx + 1
            vValorAux(iProx) = dtDataAteAux
             
            lErro = Comando_BindVarInt(alComando(1), vValorAux(iProx))
            If (lErro <> AD_SQL_SUCESSO) Then gError 190938
    
If bGeraLog Then Print #1, "         PARAM" & CStr(iProx - 1) & ": " & vValorAux(iProx)
    
        End If
            
    Else
        
        For Each objFatCondPagto In objFaturamento.colCondPagtos
        
            iProx = iProx + 1
            vValorAux(iProx) = objFatCondPagto.objCondPagtos.iCodigo
            
            lErro = Comando_BindVarInt(alComando(1), vValorAux(iProx))
            If (lErro <> AD_SQL_SUCESSO) Then gError 190936
        
If bGeraLog Then Print #1, "         PARAM" & CStr(iProx - 1) & ": " & vValorAux(iProx)
        
            If objFatCondPagto.dtDataVouDe <> DATA_NULA Then
            
                iProx = iProx + 1
                vValorAux(iProx) = objFatCondPagto.dtDataVouDe
                
                lErro = Comando_BindVarInt(alComando(1), vValorAux(iProx))
                If (lErro <> AD_SQL_SUCESSO) Then gError 190937

If bGeraLog Then Print #1, "         PARAM" & CStr(iProx - 1) & ": " & vValorAux(iProx)

            End If
            If objFatCondPagto.dtDataVouAte <> DATA_NULA Then
                
                iProx = iProx + 1
                vValorAux(iProx) = objFatCondPagto.dtDataVouAte
                 
                lErro = Comando_BindVarInt(alComando(1), vValorAux(iProx))
                If (lErro <> AD_SQL_SUCESSO) Then gError 190938
        
If bGeraLog Then Print #1, "         PARAM" & CStr(iProx - 1) & ": " & vValorAux(iProx)
        
            End If
        Next
    End If
    
    For Each vValor In objFaturamento.colTiposFat
    
        iProx = iProx + 1
        vValorAux(iProx) = vValor
    
        lErro = Comando_BindVarInt(alComando(1), vValorAux(iProx))
        If (lErro <> AD_SQL_SUCESSO) Then gError 194208
        
If bGeraLog Then Print #1, "         PARAM" & CStr(iProx - 1) & ": " & vValorAux(iProx)
        
    Next
    
    If objFaturamento.lCliente <> 0 Then
            
        iProx = iProx + 1
        vValorAux(iProx) = objFaturamento.lCliente
        
        lErro = Comando_BindVarInt(alComando(1), vValorAux(iProx))
        If (lErro <> AD_SQL_SUCESSO) Then gError 190940
    
If bGeraLog Then Print #1, "         PARAM" & CStr(iProx - 1) & ": " & vValorAux(iProx)
    
    End If
    
If bGeraLog Then Print #1, "2 - VAI EXECUTAR A LEITURA " & CStr(Now)
    
    lErro = Comando_ExecutarInt(alComando(1))
    If (lErro <> AD_SQL_SUCESSO) Then gError 190941
 
    'Busca o primeiro registro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190942
    
    lCount = 0
    'Para cada Produto
    Do While lErro = SUCESSO
    
        lCount = lCount + 1
    
If bGeraLog Then Print #1, "3." & CStr(lCount) & " - LEU O " & vsTipoDoc & " " & CStr(vlNumVou) & " " & CStr(Now)
    
        Set objInfoVoucher = New ClassTRVFATInfoVoucher
    
        objInfoVoucher.dtDataEmissao = vdtData
        objInfoVoucher.dValor = vdValor
        objInfoVoucher.iCondPagto = viCondPagto
        objInfoVoucher.iFilialEmpresa = Filial_Coinfo_Retorna_Corporator(viFilialCoinfo)
        objInfoVoucher.lCliente = vlCliente
        objInfoVoucher.lNumVou = vlNumVou
        objInfoVoucher.sSerie = vsSerie
        objInfoVoucher.sTipoDoc = vsTipoDoc
        objInfoVoucher.sTipoVou = vsTipoVou
        objInfoVoucher.lNumIntDoc = vlNumIntDoc
        objInfoVoucher.sCodGrupo = vsCodGrupo
        objInfoVoucher.lNumero = vlNumero
        objInfoVoucher.iMarcado = MARCADO
        objInfoVoucher.lClienteAporte = vlClienteAporte
        objInfoVoucher.iAporte = viAporte
        objInfoVoucher.sNome = vsNome
        
        If objInfoVoucher.sTipoDoc = "VOU" Then bTemVou = True
        
        objInfoVoucher.dCambio = vdCambio
        objInfoVoucher.dValorCambio = vdValorCambio
        objInfoVoucher.iPax = viPax
        objInfoVoucher.sControle = vsControle
        objInfoVoucher.sMoeda = vsMoeda
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
        objInfoVoucher.iPF = viPF
                
        If objInfoVoucher.sTipoDoc = TRV_TIPODOC_VOU_TEXTO Then
            objInfoVoucher.dValorBruto = objInfoVoucher.dValor + vdValorComisVou
            objInfoVoucher.dValorComissao = vdValorComisVou
        Else
            objInfoVoucher.dValorBruto = objInfoVoucher.dValor
        End If
        
        objInfoVoucher.iTipoGeracao = viTipoGeracao
        
        lErro = TRVFaturamento_Obtem_Cliente(alComando(11), alComando(12), objInfoVoucher)
        If lErro <> SUCESSO Then gError 190948
        
        objFaturamento.colInfoVouchers.Add objInfoVoucher
    
        'Busca o próximo registro
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190943
    
    Loop
    
If bGeraLog Then Print #1, "4 - VAI EXECUTAR A FUNÇÃO TRVFaturamento_Le_Dados2 " & CStr(Now)
    
    'Le as informações dos aportes
    lErro = TRVFaturamento_Le_Dados2(objFaturamento, alComando(4), alComando(5), alComando(6), alComando(7), alComando(8), bTemVou)
    If lErro <> SUCESSO Then gError 190948
    
If bGeraLog Then Print #1, "5 - VAI EXECUTAR A FUNÇÃO TRVFaturamento_Calcula_Dados1 " & CStr(Now)
    
    'Calcula as informações das Faturas
    lErro = TRVFaturamento_Calcula_Dados1(objFaturamento, alComando(9), alComando(10), bGeraLog)
    If lErro <> SUCESSO Then gError 190944
    
If bGeraLog Then Print #1, "6 - VAI EXECUTAR A FUNÇÃO TRVFaturamento_Calcula_Dados2 " & CStr(Now)
    
    'Calcula as informações das filiais
    lErro = TRVFaturamento_Calcula_Dados2(objFaturamento, alComando(2))
    If lErro <> SUCESSO Then gError 190946
    
If bGeraLog Then Print #1, "7 - VAI EXECUTAR A FUNÇÃO TRVFaturamento_Calcula_Dados3 " & CStr(Now)
    
    'Calcula as informações dos clientes
    lErro = TRVFaturamento_Calcula_Dados3(objFaturamento, alComando(3))
    If lErro <> SUCESSO Then gError 190947
    
If bGeraLog Then Print #1, "8 - VAI EXECUTAR A FUNÇÃO TRVFaturamento_Atualiza " & CStr(Now)
    
    'Acerta o obj conforme as marcações (redistribui o aporte, acerta o valor das faturas, número de itens selecionados, etc ...)
    lErro = CF("TRVFaturamento_Atualiza", objFaturamento)
    If lErro <> SUCESSO Then gError 190949
    
If bGeraLog Then Print #1, "9 - FINALIZOU A LEITURA " & CStr(Now)

    If bArq1Aberto Then
        Close #1
        Close #2
    End If
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
 
    TRVFaturamento_Le_Dados = SUCESSO
 
    Exit Function
 
Erro_TRVFaturamento_Le_Dados:
 
    TRVFaturamento_Le_Dados = gErr
 
    Select Case gErr
 
        Case 190920
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 190921 To 190943, 190973, 190974, 190975, 194208
            Call Rotina_Erro(vbOKOnly, "ERRO_PREPARACAO_COMANDO_SQL", gErr, sSQL)
            
        Case 190944, 190946, 190947, 190948, 190949
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190945)
 
    End Select
    
    If bArq1Aberto Then
        Close #1
        Close #2
    End If
 
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
 
    Exit Function
 
End Function

Private Function TRVFaturamento_Calcula_Dados1(ByVal objFaturamento As ClassTRVFaturamento, ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal bGeraLog As Boolean) As Long
'Com base nos itens a serem faturados calcula as demais informações

Dim lErro As Long
Dim lIndice As Long
Dim objInfoVoucher As ClassTRVFATInfoVoucher
Dim objInfoVoucherAux As ClassTRVFATInfoVoucher
Dim colCampos As New Collection
Dim colItensOrdenados As New Collection
Dim lClienteAnt As Long
Dim iCondPagtoAnt As Integer
Dim iFilialEmpresaAnt As Integer
Dim iTipoDocAnt As Integer
Dim colAuxiliar As New Collection
Dim bUltimo As Boolean
Dim objInfoFatura As ClassTRVFATInfoFatura
Dim objInfoFaturaOver As ClassTRVFATInfoFatura
Dim lSeq As Long
Dim objFatCondPagto  As ClassTRVFatCondPagto
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
Dim objVouAporte As ClassTRVVoucherAporte
Dim objOCRAporte As ClassTRVOcorrenciaAporte
Dim colVouAporte As Collection
Dim objOCR As ClassTRVOcorrencias
Dim objTipoOCR As ClassTRVTiposOcorrencia
Dim dFatorApOCR As Double
Dim lNumVouAnt As Long, sConteudo As String
Dim lFornOver As Long, bTemOver As Boolean
Dim objEmi As ClassCliEmissoresTRV, bAchouCondPagtoVou As Boolean
'Dim asReg(1 To 100000) As String, lX As Long

On Error GoTo Erro_TRVFaturamento_Calcula_Dados1

    bUltimo = False
    lSeq = 0
    
    'lX = lX + 1
    'asReg(lX) = "01 - Início TRVFaturamento_Calcula_Dados1"
    If bGeraLog Then Print #2, "01 - Início TRVFaturamento_Calcula_Dados1" & " " & CStr(Now)

    If objFaturamento.colInfoVouchers.Count > 0 Then
    
        For Each objInfoVoucher In objFaturamento.colInfoVouchers
        
            'lX = lX + 1
            'asReg(lX) = "02 - " & objInfoVoucher.sTipoDoc & " " & objInfoVoucher.sTipoVou & objInfoVoucher.sSerie & CStr(objInfoVoucher.lNumVou) & " Valor: " & Format(objInfoVoucher.dValor, "STANDARD") & " Cambio: " & Format(objInfoVoucher.dCambio, "STANDARD") & " Valor no cambio: " & Format(objInfoVoucher.dValorCambio, "STANDARD")
            If bGeraLog Then Print #2, "02 - " & objInfoVoucher.sTipoDoc & " " & objInfoVoucher.sTipoVou & objInfoVoucher.sSerie & CStr(objInfoVoucher.lNumVou) & " Valor: " & Format(objInfoVoucher.dValor, "STANDARD") & " Cambio: " & Format(objInfoVoucher.dCambio, "STANDARD") & " Valor no cambio: " & Format(objInfoVoucher.dValorCambio, "STANDARD") & " " & CStr(Now)

            If objInfoVoucher.iFilialFat <> 0 Then
                objInfoVoucher.iFilialEmpresa = objInfoVoucher.iFilialFat
            End If
                    
            Select Case objInfoVoucher.sTipoDoc

                Case TRV_TIPODOC_OVER_TEXTO
                    objInfoVoucher.iTipoDoc = TRV_TIPODOC_FAT_OVER
                    
                Case TRV_TIPODOC_OCR_TEXTO, TRV_TIPODOC_NVL_TEXTO
                    objInfoVoucher.iTipoDoc = TRV_TIPODOC_FAT_OCR
                
                Case Else
                    objInfoVoucher.iTipoDoc = TRV_TIPODOC_FAT_OUTROS
            
            End Select
            
            'O Valor do Cambio vem o do voucher que é o único que pode estar em dolar
            If objInfoVoucher.sTipoDoc <> TRV_TIPODOC_VOU_TEXTO Then
                objInfoVoucher.dValorCambio = objInfoVoucher.dValor
                objInfoVoucher.dCambio = 1
            End If
            
            'Os representantes e correntistas não precisam receber o controle preenchido pela agência
            If objInfoVoucher.sTipoDoc = TRV_TIPODOC_CMC_TEXTO Or objInfoVoucher.sTipoDoc = TRV_TIPODOC_CMR_TEXTO Then
                objInfoVoucher.sControle = ""
            End If
                
            'Pega o estorno do aporte (só para NVL do mesmo cliente do voucher, de clientes diferentes é estorno de comissão que não tem aporte)
            If (objInfoVoucher.sTipoDoc = TRV_TIPODOC_NVL_TEXTO Or objInfoVoucher.sTipoDoc = TRV_TIPODOC_OCR_TEXTO) And objInfoVoucher.lCliente = objInfoVoucher.lCliVou Then
                
                'Tem que ler o tipo para ver se deve ou não estornar
                Set objTipoOCR = New ClassTRVTiposOcorrencia
                
                objTipoOCR.iCodigo = objInfoVoucher.iTipoOCR
                
                'Le o tipo de ocorrência
                lErro = CF("TRVTiposOcorrencia_Le", objTipoOCR)
                If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 198230
                
                'Tipo de ocorrência não cadastrada
                If lErro <> SUCESSO Then gError 198231
                
                If objTipoOCR.iEstornaAporteVou = MARCADO And objInfoVoucher.dValor < 0 And objInfoVoucher.dValorVou > 0 Then
                
                    If objInfoVoucher.sTipoDoc = TRV_TIPODOC_NVL_TEXTO Then
                        dFatorApOCR = 1
                    Else
                        dFatorApOCR = Abs(objInfoVoucher.dValor / objInfoVoucher.dValorVou)
                    End If
                
                    'Tem que pegar de TRVVoucherAporte agora
                    objInfoVoucher.dValorAporte = -objInfoVoucher.dValorAporteVou * dFatorApOCR
                    
                    Set objVouAporte = New ClassTRVVoucherAporte
                    Set colVouAporte = New Collection
                
                    objVouAporte.sTipVou = objInfoVoucher.sTipoVou
                    objVouAporte.sSerie = objInfoVoucher.sSerie
                    objVouAporte.lNumVou = objInfoVoucher.lNumVou
                
                    'Le os aportes do voucher inativado
                    lErro = TRVVoucherAporte_Le_Voucher(objVouAporte, colVouAporte)
                    If lErro <> SUCESSO Then gError 196466
                    
                    'Coloca os aportes na inativação com o valor com sinal invertido para o estorno
                    For Each objVouAporte In colVouAporte
                    
                        Set objOCRAporte = New ClassTRVOcorrenciaAporte
                        
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
            If objInfoVoucher.iCartao = MARCADO And objInfoVoucher.sTipoDoc = TRV_TIPODOC_VOU_TEXTO Then
                objInfoVoucher.sTitular = objInfoVoucher.sTitularVou
            ElseIf objInfoVoucher.sTipoDoc = TRV_TIPODOC_CMC_TEXTO Or objInfoVoucher.sTipoDoc = TRV_TIPODOC_CMR_TEXTO Or ((objInfoVoucher.sTipoDoc = TRV_TIPODOC_NVL_TEXTO Or objInfoVoucher.sTipoDoc = TRV_TIPODOC_OCR_TEXTO) And objInfoVoucher.lCliente <> objInfoVoucher.lCliVou) Then
            'Se for comissão deve sair o cliente do voucher como titular
                objInfoVoucher.sTitular = objInfoVoucher.sNomeCliVou
            Else
            'Senão é o passageiro
                objInfoVoucher.sTitular = objInfoVoucher.sPassageiroNome & " " & objInfoVoucher.sPassageiroSobreNome
            End If
           
        Next
        
        'lX = lX + 1
        'asReg(lX) = "03 - Ordenação Count = " & CStr(objFaturamento.colInfoVouchers.Count)
        If bGeraLog Then Print #2, "03 - Ordenação Count = " & CStr(objFaturamento.colInfoVouchers.Count) & " " & CStr(Now)
        
        'Ordena os documentos de forma que seja possível dividir em faturas
        colCampos.Add "iFilialEmpresa"
        colCampos.Add "lCliente"
        colCampos.Add "iCondPagto"
        colCampos.Add "iTipoDoc"
        colCampos.Add "lNumVou"
        
        lErro = Ordena_Colecao(objFaturamento.colInfoVouchers, colItensOrdenados, colCampos)
        If lErro <> SUCESSO Then gError 190962
        
        For lIndice = 1 To colItensOrdenados.Count + 1
               
            'lX = lX + 1
            'asReg(lX) = "04 - Ordenação Indice = " & CStr(lIndice)
            If bGeraLog Then Print #2, "04 - Ordenação Indice = " & CStr(lIndice) & " " & CStr(Now)
               
            If lIndice = colItensOrdenados.Count + 1 Then bUltimo = True
            
            If Not bUltimo Then
                Set objInfoVoucher = colItensOrdenados.Item(lIndice)
            End If
            
            If objInfoVoucher.iFaturaJuntoOcr = MARCADO Then
                bOCRCliMesmaFat = True
            Else
                bOCRCliMesmaFat = False
            End If
            
            If lIndice = 1 Then
                lClienteAnt = objInfoVoucher.lCliente
                iCondPagtoAnt = objInfoVoucher.iCondPagto
                iFilialEmpresaAnt = objInfoVoucher.iFilialEmpresa
                iTipoDocAnt = objInfoVoucher.iTipoDoc
                lNumVouAnt = objInfoVoucher.lNumVou
            End If

            If lIndice = 1 Then
                bTrataRegistro = False
            Else
                'bTrataRegistro = (lClienteAnt <> objInfoVoucher.lCliente Or iFilialEmpresaAnt <> objInfoVoucher.iFilialEmpresa Or iCondPagtoAnt <> objInfoVoucher.iCondPagto Or bUltimo Or (iTipoDocAnt <> objInfoVoucher.iTipoDoc And ((Not bOCRCliMesmaFat) Or objInfoVoucher.iTipoDoc = TRV_TIPODOC_FAT_OVER)))
                bTrataRegistro = (lClienteAnt <> objInfoVoucher.lCliente Or iFilialEmpresaAnt <> objInfoVoucher.iFilialEmpresa Or bUltimo Or (iTipoDocAnt <> objInfoVoucher.iTipoDoc And ((Not bOCRCliMesmaFat) Or objInfoVoucher.iTipoDoc = TRV_TIPODOC_FAT_OVER)))
            
                'Se é para faturar os vouchers individualmente e trocou de voucher
                If (objFaturamento.iFaturarIndividualmente = MARCADO Or objInfoVoucher.iUmVouPorFat = MARCADO) And lNumVouAnt <> objInfoVoucher.lNumVou Then bTrataRegistro = True
            End If
            
            If bTrataRegistro Then
            
                'lX = lX + 1
                'asReg(lX) = "05 - Preenchendo a fatura "
                If bGeraLog Then Print #2, "05 - Preenchendo a fatura " & " " & CStr(Now)
                
                lClienteAnt = objInfoVoucher.lCliente
                iCondPagtoAnt = objInfoVoucher.iCondPagto
                iFilialEmpresaAnt = objInfoVoucher.iFilialEmpresa
                iTipoDoc = iTipoDocAnt
                iTipoDocAnt = objInfoVoucher.iTipoDoc
                lNumVouAnt = objInfoVoucher.lNumVou
              
                lSeq = lSeq + 1
                
                Set objInfoFatura = New ClassTRVFATInfoFatura
                
                Set objInfoFatura.colInfoVouchers = colAuxiliar
                
                dValor = 0
                For Each objInfoVoucherAux In objInfoFatura.colInfoVouchers
                    objInfoVoucherAux.lFatura = lSeq
                    dValor = dValor + objInfoVoucherAux.dValor
                Next
                
                bAchouCondPagtoVou = False
                For Each objInfoVoucherAux In objInfoFatura.colInfoVouchers
                    If objInfoVoucherAux.sTipoDoc = TRV_TIPODOC_VOU_TEXTO Then
                        bAchouCondPagtoVou = True
                        Exit For
                    End If
                Next
                
                If Not bAchouCondPagtoVou Then
                    For Each objInfoVoucherAux In objInfoFatura.colInfoVouchers
                        Exit For
                    Next
                End If
                
                For Each objFatCondPagto In objFaturamento.colCondPagtos
                    If objInfoVoucherAux.iCondPagto = objFatCondPagto.objCondPagtos.iCodigo Then
                        Exit For
                    End If
                Next
                
                'lX = lX + 1
                'asReg(lX) = "07 - " & objInfoVoucherAux.sTipoDoc & " " & objInfoVoucherAux.sTipoVou & objInfoVoucherAux.sSerie & CStr(objInfoVoucherAux.lNumVou) & " Valor: " & Format(objInfoVoucherAux.dValor, "STANDARD") & " Cambio: " & Format(objInfoVoucherAux.dCambio, "STANDARD") & " Valor no cambio: " & Format(objInfoVoucherAux.dValorCambio, "STANDARD") & " CondPagto: " & CStr(objInfoVoucherAux.iCondPagto)
                If bGeraLog Then Print #2, "06 - " & objInfoVoucherAux.sTipoDoc & " " & objInfoVoucherAux.sTipoVou & objInfoVoucherAux.sSerie & CStr(objInfoVoucherAux.lNumVou) & " Valor: " & Format(objInfoVoucherAux.dValor, "STANDARD") & " Cambio: " & Format(objInfoVoucherAux.dCambio, "STANDARD") & " Valor no cambio: " & Format(objInfoVoucherAux.dValorCambio, "STANDARD") & " CondPagto: " & CStr(objInfoVoucherAux.iCondPagto) & " " & CStr(Now)
                
                objInfoFatura.dtDataEmissao = objFatCondPagto.dtDataEmissao
                objInfoFatura.dtDataVencimento = objFatCondPagto.dtDataVencimento
                objInfoFatura.dValor = dValor 'Voucher - Comissão (Depois tem que calcular e abater os aportes)
                objInfoFatura.iCondPagto = objInfoVoucherAux.iCondPagto
                objInfoFatura.iFilialEmpresa = objInfoVoucherAux.iFilialEmpresa
                objInfoFatura.iPF = objInfoVoucherAux.iPF
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
    
    'lX = lX + 1
    'asReg(lX) = "08 - Tratamento especial para OVER"
    If bGeraLog Then Print #2, "07 - Tratamento especial para OVER" & " " & CStr(Now)
    
    'Tratamento especial para o OVER
    Set objInfoFaturaOver = New ClassTRVFATInfoFatura
    bTemOver = False
    For lIndice = objFaturamento.colInfoFaturas.Count To 1 Step -1
        Set objInfoFatura = objFaturamento.colInfoFaturas.Item(lIndice)
        If objInfoFatura.iTipoDoc = TRV_TIPODOC_FAT_OVER Then
        
            'lX = lX + 1
            'asReg(lX) = "09 - Fatura Indice = " & CStr(lIndice) & " Emissor = " & CStr(objInfoFatura.lCliente) & " Valor: " & Format(objInfoFatura.dValor, "STANDARD")
            If bGeraLog Then Print #2, "08 - Fatura Indice = " & CStr(lIndice) & " Emissor = " & CStr(objInfoFatura.lCliente) & " Valor: " & Format(objInfoFatura.dValor, "STANDARD") & " " & CStr(Now)
        
            Set objForn = New ClassFornecedor
            objForn.lCodigo = objInfoFatura.lCliente
            lErro = CF("Fornecedor_Le", objForn)
            If lErro <> SUCESSO And lErro <> 12729 Then gError ERRO_SEM_MENSAGEM
            
            Set objEmi = New ClassCliEmissoresTRV
            objEmi.lFornEmissor = objInfoFatura.lCliente
            lErro = TRVCliEmissores_Le(objEmi)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
        
            bTemOver = True
            'Se tiver valor a pagar e o emissor tiver o CPF computa para Inc Club
            If objInfoFatura.dValor < -0.01 And Len(Trim(objForn.sCgc)) > 0 Then
                If objInfoFaturaOver.colInfoVouchers.Count = 0 Then
                    objInfoFaturaOver.dtDataEmissao = objInfoFatura.dtDataEmissao
                    objInfoFaturaOver.dtDataVencimento = objInfoFatura.dtDataVencimento
                    objInfoFaturaOver.iCondPagto = objInfoFatura.iCondPagto
                    objInfoFaturaOver.iFilialEmpresa = objInfoFatura.iFilialEmpresa
                    objInfoFaturaOver.iMarcado = MARCADO
                    objInfoFaturaOver.iTipoDoc = TRV_TIPODOC_FAT_OVER
                End If
                objInfoFaturaOver.dValor = objInfoFaturaOver.dValor + objInfoFatura.dValor
                For Each objInfoVoucher In objInfoFatura.colInfoVouchers
                    objInfoVoucher.sControle = objInfoVoucher.sEmissor
                    objInfoVoucher.sEmissorCPF = objForn.sCgc
                    objInfoVoucher.sEmissorCartao = objEmi.sNumCartao
                    objInfoVoucher.sEmissorNomeCompl = objForn.sRazaoSocial
                    objInfoFaturaOver.colInfoVouchers.Add objInfoVoucher
                Next
            Else
                For Each objInfoVoucher In objInfoFatura.colInfoVouchers
                    'Se tinha valor a pagar joga na lista de pendentes por conta do CPF
                    If objInfoFatura.dValor < -0.01 Then
                        objInfoVoucher.sEmissorCPF = objForn.sCgc
                        objInfoVoucher.sEmissorNomeCompl = objForn.sRazaoSocial
                        objInfoVoucher.sEmissorCartao = objEmi.sNumCartao
                        objInfoFaturaOver.colInfoVouchersPen.Add objInfoVoucher
                    End If
                    'Marca para remover dos itens a serem faturados
                    objInfoVoucher.iMarcado = -1
                Next
            End If
            objFaturamento.colInfoFaturas.Remove (lIndice)
        End If
    Next
    
    'lX = lX + 1
    'asReg(lX) = "10 - Finalizando"
    If bGeraLog Then Print #2, "09 - Finalizando" & " " & CStr(Now)
    
    If bTemOver Then
        
        For lIndice = objFaturamento.colInfoVouchers.Count To 1 Step -1
            Set objInfoVoucher = objFaturamento.colInfoVouchers.Item(lIndice)
            If objInfoVoucher.iMarcado = -1 Then
                objFaturamento.colInfoVouchers.Remove (lIndice)
            End If
        Next
                    
        If objInfoFaturaOver.colInfoVouchers.Count > 0 Then
        
            lErro = CF("TRVConfig_Le", TRVCONFIG_CLIENTE_OVER, EMPRESA_TODA, sConteudo)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            lFornOver = StrParaLong(sConteudo)
            
            If lFornOver = 0 Then gError 206589
               
            objInfoFaturaOver.lCliente = lFornOver
            objFaturamento.colInfoFaturas.Add objInfoFaturaOver
            lSeq = 0
            For Each objInfoFatura In objFaturamento.colInfoFaturas
                lSeq = lSeq + 1
                objInfoFatura.lFatura = lSeq
                For Each objInfoVoucher In objInfoFatura.colInfoVouchers
                    objInfoVoucher.lFatura = lSeq
                Next
            Next
        End If
    End If
    'Tratamento especial para o OVER
    
    'lX = lX + 1
    'asReg(lX) = "11 - Fim"
    If bGeraLog Then Print #2, "10 - Fim TRVFaturamento_Calcula_Dados1" & " " & CStr(Now)

    TRVFaturamento_Calcula_Dados1 = SUCESSO

    Exit Function

Erro_TRVFaturamento_Calcula_Dados1:

    TRVFaturamento_Calcula_Dados1 = gErr

    Select Case gErr
    
        Case 190960 To 190962, 192725, 194201, 194210, 196466, 198230
        
        Case 198231 'Tipo de ocorrência não cadastrado
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_TRVTIPOSOCORRENCIA_NAO_PREENCHIDO", gErr)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190963)

    End Select
    
    'Call Grava_Log(asReg())

    Exit Function

End Function

Private Function TRVFaturamento_Calcula_Dados2(ByVal objFaturamento As ClassTRVFaturamento, ByVal lComando As Long) As Long
'Com base nos itens a serem faturados calcula as demais informações

Dim lErro As Long
Dim objCodigoNome As New AdmCodigoNome
Dim colCodigoNome As New AdmColCodigoNome
Dim objInfoFilial As ClassTRVFATInfoFilial
Dim lQtdItens As Long
Dim dValorItens As Double
Dim iFilialEmpresaAux As Integer

On Error GoTo Erro_TRVFaturamento_Calcula_Dados2

    'Lê o Código e o Nome de Toda FilialEmpresa do BD
    lErro = CF("Cod_Nomes_Le_FilEmp", colCodigoNome)
    If lErro <> SUCESSO Then gError 190980

    For Each objCodigoNome In colCodigoNome
        If objCodigoNome.iCodigo < Abs(giFilialAuxiliar) Then
            Set objInfoFilial = New ClassTRVFATInfoFilial
            objInfoFilial.iFilialEmpresa = objCodigoNome.iCodigo
            objInfoFilial.sNome = objCodigoNome.sNome
            objFaturamento.colInfoFiliais.Add objInfoFilial
        End If
    Next
    
    If objFaturamento.iOtimizar = DESMARCADO Then

        lErro = Comando_Executar(lComando, "SELECT FilialEmpresa, COUNT(*), SUM(Valor) FROM DocsParaFatAux GROUP BY FilialEmpresa ORDER BY FilialEmpresa ", iFilialEmpresaAux, lQtdItens, dValorItens)
        If lErro <> AD_SQL_SUCESSO Then gError 190964
        
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190965
    
        Do While lErro <> AD_SQL_SEM_DADOS
        
            For Each objInfoFilial In objFaturamento.colInfoFiliais
                If objInfoFilial.iFilialEmpresa = iFilialEmpresaAux Then
                    objInfoFilial.lNumItens = lQtdItens
                    objInfoFilial.dValor = dValorItens
                    Exit For
                End If
            Next
        
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190965
        
        Loop
        
    End If

    TRVFaturamento_Calcula_Dados2 = SUCESSO

    Exit Function

Erro_TRVFaturamento_Calcula_Dados2:

    TRVFaturamento_Calcula_Dados2 = gErr

    Select Case gErr
    
        Case 190964, 190965
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_DOCSPARAFAT", gErr)
            
        Case 190980

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190966)

    End Select

    Exit Function

End Function

Private Function TRVFaturamento_Calcula_Dados3(ByVal objFaturamento As ClassTRVFaturamento, ByVal lComando As Long) As Long
'Com base nos itens a serem faturados calcula as demais informações

Dim lErro As Long
Dim objInfoFatura As ClassTRVFATInfoFatura
Dim objInfoFaturaAux As ClassTRVFATInfoFatura
Dim bUltimo As Boolean
Dim colCampos As New Collection
Dim colItensOrdenados As New Collection
Dim lClienteAnt As Long
Dim lClienteAux As Long
Dim objInfoCliente As ClassTRVFATInfoCliente
Dim colAuxiliar As New Collection
Dim lQtdItens As Long
Dim dValorItens As Double
Dim iIndice As Integer
Dim iFilialEmpresaAnt As Integer
Dim objForn As ClassFornecedor, bSoOVER As Boolean

On Error GoTo Erro_TRVFaturamento_Calcula_Dados3

    bUltimo = False
    bSoOVER = True
    
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
                
                Set objInfoCliente = New ClassTRVFATInfoCliente
                
                Set objInfoCliente.colInfoFaturas = colAuxiliar
                                       
                For Each objInfoFaturaAux In objInfoCliente.colInfoFaturas
                    Exit For
                Next

                objInfoCliente.iMarcado = MARCADO
                objInfoCliente.lCliente = objInfoFaturaAux.lCliente
                objInfoCliente.iFilialEmpresa = objInfoFaturaAux.iFilialEmpresa
                
                If objInfoFaturaAux.iTipoDoc = TRV_TIPODOC_FAT_OVER Then
                    'Le o nome da empresa que faz o pagamento do OVER (Não vem do select porque é trocada depois)
                    Set objForn = New ClassFornecedor
                    objForn.lCodigo = objInfoCliente.lCliente
                    lErro = CF("Fornecedor_Le", objForn)
                    If lErro <> SUCESSO And lErro <> 12729 Then gError 190967
                    objInfoCliente.sNome = objForn.sNomeReduzido
                Else
                    bSoOVER = False
                    objInfoCliente.sNome = objInfoFaturaAux.colInfoVouchers.Item(1).sNome
                End If
                If objInfoFaturaAux.iTipoDoc = TRV_TIPODOC_FAT_OVER Then
                    objInfoCliente.iTipo = TRV_CLIENTEINFO_TIPO_FORNECEDOR
                Else
                    objInfoCliente.iTipo = TRV_CLIENTEINFO_TIPO_CLIENTE
                End If
                
                objFaturamento.colInfoClientes.Add objInfoCliente
            
                Set colAuxiliar = New Collection
                
            End If
            
            colAuxiliar.Add objInfoFatura
        
        Next
        
    End If
    
    If objFaturamento.colInfoClientes.Count <> 0 And (Not bSoOVER) And objFaturamento.iOtimizar = DESMARCADO Then
    
        If objFaturamento.colInfoClientes.Count = 1 Then
        
            Set objInfoCliente = objFaturamento.colInfoClientes.Item(1)
        
            lErro = Comando_Executar(lComando, "SELECT COUNT(*), SUM(Valor) FROM DocsParaFatAux WHERE Cliente = ? ", lQtdItens, dValorItens, objInfoCliente.lCliente)
            If lErro <> AD_SQL_SUCESSO Then gError 190968
        
            lErro = Comando_BuscarPrimeiro(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190969
        
            objInfoCliente.lNumItens = lQtdItens
            objInfoCliente.dValor = dValorItens
        
        Else
            
            lErro = Comando_Executar(lComando, "SELECT Cliente, COUNT(*), SUM(Valor) FROM DocsParaFatAux GROUP BY Cliente ORDER BY Cliente ", lClienteAux, lQtdItens, dValorItens)
            If lErro <> AD_SQL_SUCESSO Then gError 190968
        
            lErro = Comando_BuscarPrimeiro(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190969
            
            Do While lErro <> AD_SQL_SEM_DADOS
            
                For Each objInfoCliente In objFaturamento.colInfoClientes
                    If objInfoCliente.lCliente = lClienteAux Then
                        objInfoCliente.lNumItens = lQtdItens
                        objInfoCliente.dValor = dValorItens
                        Exit For
                    End If
                Next
            
                lErro = Comando_BuscarProximo(lComando)
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190969
            
            Loop
            
        End If
        
    End If
    
    TRVFaturamento_Calcula_Dados3 = SUCESSO

    Exit Function

Erro_TRVFaturamento_Calcula_Dados3:

    TRVFaturamento_Calcula_Dados3 = gErr

    Select Case gErr
    
        Case 190967
        
        Case 190968, 190969
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_DOCSPARAFAT", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190959)

    End Select


    Exit Function

End Function

Private Function Verifica_Validade_Aporte(ByVal lComando As Long, ByVal lClienteAporte As Long, ByVal lClienteVoucher As Long, bValido As Boolean) As Long
'Verifica se o cliente tem direito ao aporte

Dim lErro As Long
Dim lClientePai As Long
Dim iAporte As Integer

On Error GoTo Erro_Verifica_Validade_Aporte
    
    bValido = False

    lErro = Comando_Executar(lComando, "SELECT EmpPai, Aporte FROM CoinfoEmpresaAtual WHERE CodEmp = ? ", lClientePai, iAporte, lClienteVoucher)
    If lErro <> AD_SQL_SUCESSO Then gError 190970
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190971
    
    If lClientePai = lClienteAporte And iAporte = MARCADO Then bValido = True

    Verifica_Validade_Aporte = SUCESSO

    Exit Function

Erro_Verifica_Validade_Aporte:

    Verifica_Validade_Aporte = gErr

    Select Case gErr
    
        Case 190970, 190971
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLIENTES", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190972)

    End Select

    Exit Function

End Function

Public Function TRVFaturamento_Atualiza(ByVal objFaturamento As ClassTRVFaturamento) As Long
'Com base nos itens a serem faturados calcula as demais informações

Dim lErro As Long
Dim objInfoFilial As ClassTRVFATInfoFilial
Dim objInfoVoucher As ClassTRVFATInfoVoucher
Dim objInfoFatura As ClassTRVFATInfoFatura
Dim objInfoCliente As ClassTRVFATInfoCliente
Dim lNumItens As Long
Dim lNumFat As Long
Dim dValorItens As Double
Dim dValorComissao As Double
Dim colItens As New Collection
Dim colItensOrdenados As New Collection
Dim colCampos As New Collection
Dim objFatAporte As ClassTRVFatAporte
Dim objFatAporteCred As ClassTRVFatAporteCred
Dim dValorAporte As Double
Dim lLinha As Long
Dim dValorAporteAux As Double
Dim dValorCreditos As Double
Dim dValorDebitos As Double
Dim objVouAporte As ClassTRVVoucherAporte
Dim dValorAporteUSS As Double
Dim objFaturaAporte As ClassTRVFaturaAporte
Dim sConteudo As String, dFornOverPercComiss As Double
Dim iNumEmissorSemCPF As Integer
Dim colEmiSemCPF As Collection, bAchou As Boolean
Dim objInfoVouAux As ClassTRVFATInfoVoucher, dValorCartaoOver As Double
Dim dValorItensAux As Double, dValorItem As Double

On Error GoTo Erro_TRVFaturamento_Atualiza

    'Zera a utilização dos aportes
    For Each objInfoVoucher In objFaturamento.colInfoVouchers
        'Se é um voucher
        If objInfoVoucher.sTipoDoc = TRV_TIPODOC_VOU_TEXTO Then
            objInfoVoucher.dValorAporte = 0
            Set objInfoVoucher.colAportes = New Collection
        End If
        If objInfoVoucher.iMarcado = MARCADO Then
            colItens.Add objInfoVoucher
        End If
    Next
    For Each objFatAporte In objFaturamento.colAportes
        If objFatAporte.iTipoPagtoAporte = TRV_TIPO_APORTE_SOBREFATURA Then
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
        If objInfoVoucher.sTipoDoc = TRV_TIPODOC_VOU_TEXTO Then

            For Each objFatAporte In objFaturamento.colAportes
            
                'Se é o mesmo cliente e o aporte ainda tem saldo disponível ou é por cumprimento de meta
                'e o voucher em questão é posteior ao cumprimento dessa meta
                If objFatAporte.lCliente = objInfoVoucher.lClienteAporte And objInfoVoucher.iAporte = MARCADO And (objFatAporte.dSaldoAtual > 0 Or (objFatAporte.iTipoPagtoAporte = TRV_TIPO_APORTE_SOBREFATURA_COND And (objInfoVoucher.dtDataEmissao > objFatAporte.dtDataVouAntMeta Or (objInfoVoucher.dtDataEmissao = objFatAporte.dtDataVouAntMeta And objInfoVoucher.lNumVou >= objFatAporte.lNumVouAntMeta)))) Then
                
                    'Se não passou da validade
                    If (objInfoVoucher.dtDataEmissao <= objFatAporte.objAportePagto.dtValidadeAte Or objFatAporte.objAportePagto.dtValidadeAte = DATA_NULA) Then
                    
                        'Se já começou a ser válido
                        If (objInfoVoucher.dtDataEmissao >= objFatAporte.objAportePagto.dtValidadeDe Or objFatAporte.objAportePagto.dtValidadeDe = DATA_NULA) Then
                        
                            dValorAporte = objInfoVoucher.dValor * objFatAporte.objAportePagto.dPercentual
                            
                            'Se é por cumprimento de meta
                            If objFatAporte.iTipoPagtoAporte = TRV_TIPO_APORTE_SOBREFATURA_COND Then
                                'Se é o primeiro voucher que deveria receber a meta
                                If objInfoVoucher.lNumVou = objFatAporte.lNumVouAntMeta Then
                                    If objInfoVoucher.dCambio = 0 Then gError 192862
                                    dValorAporte = (((objFatAporte.dRealizadoAntMetaUSS - objFatAporte.dPrevisaoUSS) * objInfoVoucher.dCambio)) * objFatAporte.objAportePagto.dPercentual
                                End If
                            End If
                            
                            If objFatAporte.iMoeda = MOEDA_DOLAR Then
                                If objInfoVoucher.dCambio = 0 Then gError 192862
                                If dValorAporte / objInfoVoucher.dCambio > objFatAporte.dSaldoAtual And objFatAporte.iTipoPagtoAporte = TRV_TIPO_APORTE_SOBREFATURA Then
                                    dValorAporte = objFatAporte.dSaldoAtual * objInfoVoucher.dCambio
                                End If
                            Else
                                If dValorAporte > objFatAporte.dSaldoAtual And objFatAporte.iTipoPagtoAporte = TRV_TIPO_APORTE_SOBREFATURA Then
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
                            
                            Set objVouAporte = New ClassTRVVoucherAporte
                            
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
        
    lLinha = 0
    For Each objInfoFatura In objFaturamento.colInfoFaturas
    
        dValorItens = 0
        dValorComissao = 0
        dValorAporte = 0
        dValorCreditos = 0
        dValorDebitos = 0
        lLinha = lLinha + 1
        
        objInfoFatura.lLinha = lLinha
        
        Call objInfoFatura.Agrupa_Itens
        
        'Tira uma prévia do valor da fatura para saber se é a pagar ou a receber
        dValorItensAux = 0
        For Each objInfoVoucher In objInfoFatura.colInfoVouchersAG
            If objInfoVoucher.iMarcado = MARCADO Then
                dValorItensAux = dValorItensAux + objInfoVoucher.dValorBruto - objInfoVoucher.dValorAporte - objInfoVoucher.dValorComissao
            End If
        Next
        
        For Each objInfoVoucher In objInfoFatura.colInfoVouchersAG
        
            dValorItem = objInfoVoucher.dValorBruto - objInfoVoucher.dValorAporte - objInfoVoucher.dValorComissao
        
            'Se for a fatura for a pagar, com o documento a pagar emitido antes do prazo mínimo desmarca
            If dValorItensAux < 0 And dValorItem < 0 And DateAdd("d", objFaturamento.iPrazoMinPagto, objInfoVoucher.dtDataEmissao) > Date Then
                'Se for para pessoa física não retém o pagamento
                If objInfoVoucher.sTipoDoc <> TRV_TIPODOC_OVER_TEXTO And objInfoFatura.iPF = DESMARCADO Then
                    objInfoVoucher.iMarcado = DESMARCADO
                    'Percorre e desmarca os itens de forma desagrupada para as comissões
                    If objInfoVoucher.sTipoDoc <> "VOU" And objInfoVoucher.sTipoDoc <> "NVL" And objInfoVoucher.sTipoDoc <> "OCR" Then
                        For Each objInfoVouAux In objInfoFatura.colInfoVouchers
                            If objInfoVouAux.sTipoDoc = objInfoVoucher.sTipoDoc And objInfoVouAux.sTipoVou = objInfoVoucher.sTipoVou And objInfoVouAux.sSerie = objInfoVoucher.sSerie And objInfoVouAux.lNumVou = objInfoVoucher.lNumVou And objInfoVouAux.lCliente = objInfoVoucher.lCliente Then
                                objInfoVouAux.iMarcado = DESMARCADO
                            End If
                        Next
                    End If
                    dValorItensAux = dValorItensAux - dValorItem 'Vai abatendo até a fatura ficar a receber
                End If
            End If
        
            If objInfoVoucher.iMarcado = MARCADO Then
            
                dValorItens = dValorItens + objInfoVoucher.dValorBruto - objInfoVoucher.dValorAporte - objInfoVoucher.dValorComissao
                dValorComissao = dValorComissao + objInfoVoucher.dValorComissao
                dValorAporte = dValorAporte + objInfoVoucher.dValorAporte
                
                'Se for negativo é um crédito
                If objInfoVoucher.dValor < 0 Then
                    dValorCreditos = dValorCreditos + objInfoVoucher.dValor
                End If
                'Se for um documento fora voucher que acrescenta no valor é um débito
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
        
        If (Abs(dValorItens) < DELTA_VALORMONETARIO) Or (objInfoFatura.iTipoDoc = TRV_TIPODOC_FAT_OVER And dValorItens > -DELTA_VALORMONETARIO) Then
            objInfoFatura.iMarcado = DESMARCADO
            For Each objInfoVoucher In objInfoFatura.colInfoVouchers
                objInfoVoucher.iMarcado = DESMARCADO
            Next
        End If
        
        objInfoFatura.dValor = Arredonda_Moeda(dValorItens)
        objInfoFatura.dValorAporte = Arredonda_Moeda(dValorAporte)
        objInfoFatura.dValorComissao = Arredonda_Moeda(dValorComissao)
        objInfoFatura.dValorCreditos = Arredonda_Moeda(dValorCreditos)
        objInfoFatura.dValorDebitos = Arredonda_Moeda(dValorDebitos)
    
    Next

    lLinha = 0
    For Each objInfoVoucher In objFaturamento.colInfoVouchers
        lLinha = lLinha + 1
        objInfoVoucher.lLinha = lLinha
    Next
    
    'distribui primeiro da empresa pai
    For Each objFatAporteCred In objFaturamento.colAportesCred
        'Distribui os créditos via aporte parcela única
        For Each objInfoFatura In objFaturamento.colInfoFaturas
            'Se o cliente é a empresa pai do aporte
            If objInfoFatura.lCliente = objFatAporteCred.lCliente And objFatAporteCred.dSaldoAtual > 0 And objInfoFatura.dtDataEmissao >= objFatAporteCred.dtDataVenc And objInfoFatura.dValor > 0 Then
            
                Set objFaturaAporte = New ClassTRVFaturaAporte
            
                'Dá o crédito, mas deixa a fatura com no mínimo 1 centavo
                If objInfoFatura.dValor >= objFatAporteCred.dSaldoAtual + 0.01 Then
                    objFaturaAporte.dValorAporte = Arredonda_Moeda(objFatAporteCred.dSaldoAtual)
                    objFaturaAporte.iPagtoParcial = DESMARCADO
                Else
                    objFaturaAporte.dValorAporte = Arredonda_Moeda(objInfoFatura.dValor - 0.01)
                    objFaturaAporte.iPagtoParcial = MARCADO
                End If
                
                'Acerta as informações de valores na fatura
                objInfoFatura.dValor = objInfoFatura.dValor - objFaturaAporte.dValorAporte
                objInfoFatura.dValorCreditos = objInfoFatura.dValorCreditos - objFaturaAporte.dValorAporte
                objInfoFatura.dValorAporteCred = objInfoFatura.dValorAporteCred + objFaturaAporte.dValorAporte
                
                'Acerta o saldo dos créditos sobre fatura
                objFatAporteCred.dSaldoAtual = Arredonda_Moeda(objFatAporteCred.dSaldoAtual - objFaturaAporte.dValorAporte)
                
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
                If objInfoFatura.colInfoVouchers.Item(1).sTipoDoc = TRV_TIPODOC_VOU_TEXTO And objInfoFatura.colInfoVouchers.Item(1).lClienteAporte = objFatAporteCred.lCliente And objInfoFatura.colInfoVouchers.Item(1).iAporte = MARCADO And (objInfoFatura.lCliente <> objFatAporteCred.lCliente) And objFatAporteCred.dSaldoAtual > 0 And objInfoFatura.dtDataEmissao >= objFatAporteCred.dtDataVenc Then
            
                    Set objFaturaAporte = New ClassTRVFaturaAporte
                
                    'Dá o crédito, mas deixa a fatura com no mínimo 1 centavo
                    If objInfoFatura.dValor >= objFatAporteCred.dSaldoAtual + 0.01 Then
                        objFaturaAporte.dValorAporte = Arredonda_Moeda(objFatAporteCred.dSaldoAtual)
                        objFaturaAporte.iPagtoParcial = DESMARCADO
                    Else
                        objFaturaAporte.dValorAporte = Arredonda_Moeda(objInfoFatura.dValor - 0.01)
                        objFaturaAporte.iPagtoParcial = MARCADO
                    End If
                    
                    'Acerta as informações de valores na fatura
                    objInfoFatura.dValor = objInfoFatura.dValor - objFaturaAporte.dValorAporte
                    objInfoFatura.dValorCreditos = objInfoFatura.dValorCreditos - objFaturaAporte.dValorAporte
                    objInfoFatura.dValorAporteCred = objInfoFatura.dValorAporteCred + objFaturaAporte.dValorAporte
                    
                    'Acerta o saldo dos créditos sobre fatura
                    objFatAporteCred.dSaldoAtual = Arredonda_Moeda(objFatAporteCred.dSaldoAtual - objFaturaAporte.dValorAporte)
                    
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
    
    For Each objInfoFatura In objFaturamento.colInfoFaturas
    
        'Se for uma fatura de Over
        If objInfoFatura.iTipoDoc = TRV_TIPODOC_FAT_OVER Then
        
            'verifica os emissores sem o número do cartão
            Set colEmiSemCPF = New Collection
            For Each objInfoVoucher In objInfoFatura.colInfoVouchers
                If objInfoVoucher.iMarcado = MARCADO And Len(Trim(objInfoVoucher.sEmissorCartao)) = 0 Then
                    bAchou = False
                    For Each objInfoVouAux In colEmiSemCPF
                        If objInfoVouAux.lCliente = objInfoVoucher.lCliente Then
                            bAchou = True
                            Exit For
                        End If
                    Next
                    If Not bAchou Then
                        Set objInfoVouAux = New ClassTRVFATInfoVoucher
                        objInfoVouAux.lCliente = objInfoVoucher.lCliente
                        colEmiSemCPF.Add objInfoVouAux
                    End If
                End If
            Next
            
            iNumEmissorSemCPF = colEmiSemCPF.Count

            lErro = CF("TRVConfig_Le", TRVCONFIG_PERC_COMIS_CLI_OVER, EMPRESA_TODA, sConteudo)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            dFornOverPercComiss = StrParaDbl(sConteudo)
            
            lErro = CF("TRVConfig_Le", TRVCONFIG_TAR_CARTAO_NOVO_CLI_OVER, EMPRESA_TODA, sConteudo)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            dValorCartaoOver = StrParaDbl(sConteudo)
        
            objInfoFatura.iQtdEmiSemCartao = iNumEmissorSemCPF
            objInfoFatura.dValorTarifaCartaoOver = dValorCartaoOver
            
            'A tarifa cobrada é 2% em cima dos vouchers + 2,50 por cartão novo
            objInfoFatura.dValorTarifa = Arredonda_Moeda(-objInfoFatura.dValor * dFornOverPercComiss + (dValorCartaoOver * iNumEmissorSemCPF))
            objInfoFatura.dValor = Arredonda_Moeda(objInfoFatura.dValor - objInfoFatura.dValorTarifa)
        
        End If
    Next
    
    'Calcula os valor selecionado por filial
    For Each objInfoFilial In objFaturamento.colInfoFiliais
    
        lNumItens = 0
        dValorItens = 0
        lNumFat = 0
        
        For Each objInfoFatura In objFaturamento.colInfoFaturas
        
            If objInfoFatura.iFilialEmpresa = objInfoFilial.iFilialEmpresa Then
            
                If objInfoFatura.iMarcado = MARCADO Then
                    lNumFat = lNumFat + 1
                    lNumItens = lNumItens + objInfoFatura.colInfoVouchers.Count
                    dValorItens = dValorItens + objInfoFatura.dValor
                End If
            
            End If
    
        Next
        
        objInfoFilial.dValorS = dValorItens
        objInfoFilial.lNumItensS = lNumItens
        objInfoFilial.lNumFatS = lNumFat
    
    Next
    
    'Calcula os valor selecionado por cliente
    lLinha = 0
    For Each objInfoCliente In objFaturamento.colInfoClientes
    
        lNumItens = 0
        dValorItens = 0
        lLinha = lLinha + 1
        
        objInfoCliente.lLinha = lLinha
        
        For Each objInfoFatura In objFaturamento.colInfoFaturas
        
            If objInfoFatura.lCliente = objInfoCliente.lCliente Then
            
                If objInfoFatura.iMarcado = MARCADO Then
                    lNumFat = lNumFat + 1
                    lNumItens = lNumItens + objInfoFatura.colInfoVouchers.Count
                    dValorItens = dValorItens + objInfoFatura.dValor
                End If
            
            End If
    
        Next
        
        objInfoCliente.dValorS = dValorItens
        objInfoCliente.lNumItensS = lNumItens
    
    Next
    
    TRVFaturamento_Atualiza = SUCESSO

    Exit Function

Erro_TRVFaturamento_Atualiza:

    TRVFaturamento_Atualiza = gErr

    Select Case gErr
    
        Case 192862
            Call Rotina_Erro(vbOKOnly, "ERRO_APORTE_DOLAR_VOUCHER_SEM_CAMBIO", gErr, objInfoVoucher.sTipoVou, objInfoVoucher.sSerie, objInfoVoucher.lNumVou)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190958)

    End Select

    Exit Function

End Function

Public Function Coinfo_Obtem_Comissao_Cliente2(ByVal lNumVou As Long, ByVal sSerie As String, ByVal sTipVou As String, ByVal scodest As String, dComissao As Double, Optional lComandoP1 As Long = 0, Optional lComandoP2 As Long = 0) As Long
'Obtém a FilialEmpresa com base no Cliente (lcodemp)

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim dValComi As Double
Dim dtDataEmissao As Date

On Error GoTo Erro_Coinfo_Obtem_Comissao_Cliente2

    dComissao = 0

    If lComandoP1 <> 0 And lComandoP2 <> 0 Then
        alComando(0) = lComandoP1
        alComando(1) = lComandoP1
    Else
        'Abre o comando
        For iIndice = LBound(alComando) To UBound(alComando)
            alComando(iIndice) = Comando_Abrir()
            If alComando(iIndice) = 0 Then gError 130803
        Next
    End If
    
    'Le a comissão do voucher
    lErro = Comando_Executar(alComando(0), "SELECT Data, ValorComissao FROM TRVVouchers WHERE NumVou = ? AND TipoDoc = ?  AND Serie = ? AND TipVou = ?  ", _
        dtDataEmissao, dValComi, lNumVou, TRV_TIPODOC_VOU_TEXTO, sSerie, sTipVou) 'NumVou, TipoDoc, Serie, TipVou
    If lErro <> AD_SQL_SUCESSO Then gError 130805
    
    'Busca o primeiro
    lErro = Comando_BuscarProximo(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130806
    
'    If gdtDataInicioComisCorp > dtDataEmissao Or giVersaoTRV < TRV_VERSAO_COMISSIONAMENTO Then
'
'        dValComi = 0
'
'        'Le a comissão do voucher
'        lErro = Comando_Executar(alComando(1), "SELECT C.ValComi FROM CoinfoComissao AS C WHERE C.Serie = ? AND C.NumVou = ? AND C.TipVou = ? ORDER BY C.NumIntArq DESC, C.NumRegArqTexto DESC ", _
'            dValComi, sSerie, lNumVou, sTipVou)
'        If lErro <> AD_SQL_SUCESSO Then gError 130805
'
'        'Busca o primeiro
'        lErro = Comando_BuscarProximo(alComando(1))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130806
'
'    End If
    
    dComissao = dValComi
            
    If lComandoP1 = 0 Or lComandoP2 = 0 Then
        'Fecha o comando
        For iIndice = LBound(alComando) To UBound(alComando)
            Call Comando_Fechar(alComando(iIndice))
        Next
    End If
    
    Coinfo_Obtem_Comissao_Cliente2 = SUCESSO
     
    Exit Function
    
Erro_Coinfo_Obtem_Comissao_Cliente2:

    Coinfo_Obtem_Comissao_Cliente2 = gErr
     
    Select Case gErr

        Case 130803
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 130805, 130806
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COINFOARQSLIDOS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 130558)
     
    End Select
    
    If lComandoP1 = 0 Or lComandoP2 = 0 Then
        'Fecha o comando
        For iIndice = LBound(alComando) To UBound(alComando)
            Call Comando_Fechar(alComando(iIndice))
        Next
    End If
     
    Exit Function

End Function

Private Function TRVFaturamento_Le_Dados2(ByVal objFaturamento As ClassTRVFaturamento, ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lComando3 As Long, ByVal lComando4 As Long, ByVal lComando5 As Long, ByVal bTemVou As Boolean) As Long
'Le os aportes sobre fatura

Dim lErro As Long
Dim objAportePagto As Object
Dim objAportePagtoCond As ClassTRVAportePagtoFatC
Dim objFatAporte As ClassTRVFatAporte
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
Dim objFatAporteCred As ClassTRVFatAporteCred
Dim objDebitoRecCli As ClassDebitoRecCli
Dim iNumParcelas As Integer
Dim iParcela As Integer

On Error GoTo Erro_TRVFaturamento_Le_Dados2

    If bTemVou Then
                
        'Le os investimento sobre fatura (os que são descontados via percentual voucher a voucher)
        lErro = Comando_Executar(lComando1, "SELECT TipoPagtoAporte, NumIntDocPagtoAporte, ProxParcela, Cliente, CodigoAporte, Moeda, PrevValor, PrevDataDe, PrevDataAte FROM TRVAportesPagtoParaFat ORDER BY Cliente, CodigoAporte", iTipoPagtoAporte, lNumIntDoc, iProxParcela, lCliente, lCodigo, iMoeda, dPrevValor, dtPrevDataDe, dtPrevDataAte)
        If lErro <> AD_SQL_SUCESSO Then gError 190953
        
        lErro = Comando_BuscarPrimeiro(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190954
        
        Do While lErro <> AD_SQL_SEM_DADOS
        
            bConsiderar = True
        
            Set objFatAporte = New ClassTRVFatAporte
        
            objFatAporte.lCliente = lCliente
            objFatAporte.iProxParcela = iProxParcela
            objFatAporte.lCodigo = lCodigo
            objFatAporte.iMoeda = iMoeda
            objFatAporte.iTipoPagtoAporte = iTipoPagtoAporte
            objFatAporte.dPrevisaoUSS = dPrevValor
            objFatAporte.dtPrevDataDe = dtPrevDataDe
            objFatAporte.dtPrevDataAte = dtPrevDataAte
    
            If iTipoPagtoAporte = TRV_TIPO_APORTE_SOBREFATURA Then
            
                Set objAportePagto = New ClassTRVAportePagtoFat
                
                objAportePagto.lNumIntDoc = lNumIntDoc
                
                lErro = TRVAportePagtoFat_Le_NumIntDoc(lComando2, objAportePagto)
            Else
            
                Set objAportePagto = New ClassTRVAportePagtoFatC
                
                objAportePagto.lNumIntDoc = lNumIntDoc
        
                lErro = TRVAportePagtoFatCond_Le_NumIntDoc(lComando3, objAportePagto)
            End If
            If lErro <> SUCESSO Then gError 190955
                   
            Set objFatAporte.objAportePagto = objAportePagto
                   
            If iTipoPagtoAporte = TRV_TIPO_APORTE_SOBREFATURA Then
                objFatAporte.dSaldoAtual = objAportePagto.dSaldo
            Else
                'Se já prescreveu a 3 meses não considera
                If DateDiff("d", dtPrevDataAte, gdtDataAtual) < 90 Then
                
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
                    
                Else
                    bConsiderar = False
                End If
            
            End If
        
            If bConsiderar Then objFaturamento.colAportes.Add objFatAporte
        
            lErro = Comando_BuscarProximo(lComando1)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190956
        
        Loop
        
    End If

    'Le os investimento diretos via crédito (os pagtos direto por crédito abatem em parcela única no faturamento, salvo quando o valor da fatura não permitir o crédito de forma única)
    lErro = Comando_Executar(lComando4, "SELECT TipoPagtoAporte, NumIntDocPagto, NumIntDocAporte, Cliente, ValorTotal, Saldo, Vencimento, NumIntDocDestino, Codigo, ParcMax, Parcela FROM TRVAporteCred ORDER BY Cliente, Codigo", _
    iTipoPagtoAporte, lNumIntDoc, lNumIntDocAporte, lCliente, dValor, dSaldo, dtVencimento, lNumIntDocDestino, lCodigo, iNumParcelas, iParcela)
    If lErro <> AD_SQL_SUCESSO Then gError 190953
    
    lErro = Comando_BuscarPrimeiro(lComando4)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190954
    
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objFatAporteCred = New ClassTRVFatAporteCred
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

    TRVFaturamento_Le_Dados2 = SUCESSO

    Exit Function

Erro_TRVFaturamento_Le_Dados2:

    TRVFaturamento_Le_Dados2 = gErr

    Select Case gErr
    
        Case 190953, 190954, 190956
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVAPORTEPAGTOFAT", gErr)
            
        Case 190955

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190957)

    End Select

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function TRVVouchers_Le(ByVal objTRVVouchers As ClassTRVVouchers, Optional ByVal lComandoP1 As Long = 0, Optional ByVal lComandoP2 As Long = 0) As Long

Dim lErro As Long
Dim lComando As Long
Dim tTRVVouchers As typeTRVVouchers
Dim lComandoV As Long

On Error GoTo Erro_TRVVouchers_Le

    If lComandoP1 <> 0 Then
        lComando = lComandoP1
    Else
        'Executa a abertura do Comando
        lComando = Comando_Abrir()
        If lComando = 0 Then gError 190841
    End If
    
    If lComandoP2 <> 0 Then
        lComandoV = lComandoP2
    Else
        lComandoV = Comando_Abrir()
        If lComandoV = 0 Then gError 190841
    End If

    'Alocação de espaço no buffer
    tTRVVouchers.sTipoDoc = String(STRING_TRV_OCR_TIPODOC, 0)
    tTRVVouchers.sSerie = String(STRING_TRV_OCR_SERIE, 0)
    tTRVVouchers.sTipVou = String(STRING_TRV_OCR_TIPOVOU, 0)
    tTRVVouchers.sCodGrupo = String(STRING_TRV_OCR_CODGRUPO, 0)
    tTRVVouchers.sControle = String(STRING_TRV_VOU_CONTROLE, 0)
    tTRVVouchers.sMoeda = String(STRING_TRV_VOU_MOEADA, 0)
    tTRVVouchers.sProduto = String(STRING_TRV_VOU_PRODUTO, 0)
    tTRVVouchers.sTitular = String(STRING_TRV_VOU_TITULAR, 0)
    tTRVVouchers.sCiaCart = String(STRING_TRV_VOU_CIACART, 0)
    tTRVVouchers.sNumCCred = String(STRING_TRV_VOU_NUMCCRED, 0)
    tTRVVouchers.sPassageiroNome = String(STRING_TRV_VOU_PAXNOME, 0)
    tTRVVouchers.sPassageiroSobreNome = String(STRING_TRV_VOU_PAXNOME, 0)
    tTRVVouchers.sDestino = String(STRING_TRV_VOU_DESTINO, 0)
    tTRVVouchers.sUsuarioCanc = String(STRING_USUARIO_NOMEREDUZIDO, 0)
    tTRVVouchers.sTitularCPF = String(STRING_CGC, 0)
    tTRVVouchers.sUsuarioWeb = String(STRING_TRV_VOU_USUARIOWEB, 0)
    tTRVVouchers.sNumAuto = String(STRING_TRV_VOU_NUMAUTO, 0)
    tTRVVouchers.sValidadeCC = String(STRING_TRV_VOU_VALIDADECC, 0)

    lErro = Comando_Executar(lComando, "SELECT NumVou, TipoDoc, Serie, NumIntArq, NumRegArqTexto, Cliente, " & _
                "Recnsica, Valor, Data, CondPagto, TipVou, CodGrupo, NumFatCoinfo, " & _
                "TipoDocDestino, NumIntDocDestino, Cartao, Status, NumIntDocNVL, NumIntDocPagtoAporteFat, ValorAporte, " & _
                "ValorComissao, Titular, Produto, Pax, Moeda, ValorCambio, Cambio, " & _
                "Controle, CiaCart, NumCCred, NumAuto, QuantParc, DiasAntc, Kit, " & _
                "CodigoAporte, ParcelaAporte, ExtraiInfoSigav, ValorOcr, TemOcr, Representante, ComissaoRep, " & _
                "Correntista, ComissaoCorr, Emissor, ComissaoEmissor, ComissaoAg, ValorBruto, PassageiroNome, " & _
                "PassageiroSobrenome, ClienteVou, CliPassageiro, DataCanc, HoraCanc, UsuarioCanc, ClienteComissao, " & _
                "Promotor, Destino, TitularCPF, TarifaUnitaria, Vigencia, UsuarioWeb,ValorBaseComis, ValorBrutoComOCR,ValorCMAComOCR,ValorCMC,ValorCMR,ValorCMCC,ValorCME,ValidadeCC, NumIntDoc,NumBoleto,ValorBoleto,DataVencBoleto,TrataBoleto FROM TrvVouchers WHERE NumVou= ?  AND TipoDoc= ?  AND Serie= ?  AND TipVou= ? ", _
                tTRVVouchers.lNumVou, tTRVVouchers.sTipoDoc, tTRVVouchers.sSerie, tTRVVouchers.lNumIntArq, _
                tTRVVouchers.lNumRegArqTexto, tTRVVouchers.lCliente, tTRVVouchers.lrecnsica, tTRVVouchers.dValor, tTRVVouchers.dtData, _
                tTRVVouchers.iCondPagto, tTRVVouchers.sTipVou, tTRVVouchers.sCodGrupo, tTRVVouchers.lNumFatCoinfo, tTRVVouchers.iTipoDocDestino, _
                tTRVVouchers.lNumIntDocDestino, tTRVVouchers.iCartao, tTRVVouchers.iStatus, tTRVVouchers.lNumIntDocNVL, tTRVVouchers.lNumIntDocPagtoAporteFat, _
                tTRVVouchers.dValorAporte, tTRVVouchers.dValorComissao, tTRVVouchers.sTitular, tTRVVouchers.sProduto, tTRVVouchers.iPax, _
                tTRVVouchers.sMoeda, tTRVVouchers.dValorCambio, tTRVVouchers.dCambio, tTRVVouchers.sControle, tTRVVouchers.sCiaCart, _
                tTRVVouchers.sNumCCred, tTRVVouchers.sNumAuto, tTRVVouchers.iQuantParc, tTRVVouchers.idiasantc, tTRVVouchers.iKit, _
                tTRVVouchers.lCodigoAporte, tTRVVouchers.iParcelaAporte, tTRVVouchers.iExtraiInfoSigav, tTRVVouchers.dValorOcr, tTRVVouchers.iTemOcr, _
                tTRVVouchers.lRepresentante, tTRVVouchers.dComissaoRep, tTRVVouchers.lCorrentista, tTRVVouchers.dComissaoCorr, tTRVVouchers.lEmissor, _
                tTRVVouchers.dComissaoEmissor, tTRVVouchers.dComissaoAg, tTRVVouchers.dValorBruto, tTRVVouchers.sPassageiroNome, tTRVVouchers.sPassageiroSobreNome, _
                tTRVVouchers.lClienteVou, tTRVVouchers.lCliPassageiro, tTRVVouchers.dtDataCanc, tTRVVouchers.dHoraCanc, tTRVVouchers.sUsuarioCanc, _
                tTRVVouchers.lClienteComissao, tTRVVouchers.lPromotor, tTRVVouchers.sDestino, tTRVVouchers.sTitularCPF, tTRVVouchers.dTarifaUnitaria, tTRVVouchers.iVigencia, tTRVVouchers.sUsuarioWeb, tTRVVouchers.dValorBaseComis, tTRVVouchers.dValorBrutoComOCR, tTRVVouchers.dValorCMAComOCR, tTRVVouchers.dValorCMC, tTRVVouchers.dValorCMR, tTRVVouchers.dValorCMCC, tTRVVouchers.dValorCME, tTRVVouchers.sValidadeCC, tTRVVouchers.lNumIntDoc, tTRVVouchers.lNumBoleto, tTRVVouchers.dValorBoleto, tTRVVouchers.dtDataVencBoleto, tTRVVouchers.iTrataBoleto, _
                objTRVVouchers.lNumVou, objTRVVouchers.sTipoDoc, objTRVVouchers.sSerie, objTRVVouchers.sTipVou)
    If lErro <> AD_SQL_SUCESSO Then gError 190842

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190843

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objTRVVouchers.lNumVou = tTRVVouchers.lNumVou
    objTRVVouchers.sTipoDoc = tTRVVouchers.sTipoDoc
    objTRVVouchers.sSerie = tTRVVouchers.sSerie
    objTRVVouchers.lNumIntArq = tTRVVouchers.lNumIntArq
    objTRVVouchers.lNumRegArqTexto = tTRVVouchers.lNumRegArqTexto
    objTRVVouchers.lCliente = tTRVVouchers.lCliente
    objTRVVouchers.lrecnsica = tTRVVouchers.lrecnsica
    objTRVVouchers.dValor = tTRVVouchers.dValor
    objTRVVouchers.dtData = tTRVVouchers.dtData
    objTRVVouchers.iCondPagto = tTRVVouchers.iCondPagto
    objTRVVouchers.sTipVou = tTRVVouchers.sTipVou
    objTRVVouchers.sCodGrupo = tTRVVouchers.sCodGrupo
    objTRVVouchers.lNumFatCoinfo = tTRVVouchers.lNumFatCoinfo
    objTRVVouchers.iTipoDocDestino = tTRVVouchers.iTipoDocDestino
    objTRVVouchers.lNumIntDocDestino = tTRVVouchers.lNumIntDocDestino
    objTRVVouchers.iCartao = tTRVVouchers.iCartao
    objTRVVouchers.iStatus = tTRVVouchers.iStatus
    objTRVVouchers.lNumIntDocNVL = tTRVVouchers.lNumIntDocNVL
    objTRVVouchers.lNumIntDocPagtoAporteFat = tTRVVouchers.lNumIntDocPagtoAporteFat
    objTRVVouchers.dValorAporte = tTRVVouchers.dValorAporte
    objTRVVouchers.dValorComissao = tTRVVouchers.dValorComissao
    objTRVVouchers.sTitular = tTRVVouchers.sTitular
    objTRVVouchers.sProduto = tTRVVouchers.sProduto
    objTRVVouchers.iPax = tTRVVouchers.iPax
    objTRVVouchers.sMoeda = tTRVVouchers.sMoeda
    objTRVVouchers.dValorCambio = tTRVVouchers.dValorCambio
    objTRVVouchers.dCambio = tTRVVouchers.dCambio
    objTRVVouchers.sControle = tTRVVouchers.sControle
    objTRVVouchers.sCiaCart = tTRVVouchers.sCiaCart
    objTRVVouchers.sNumCCred = tTRVVouchers.sNumCCred
    'objTRVVouchers.lNumAuto = tTRVVouchers.lNumAuto
    objTRVVouchers.sNumAuto = tTRVVouchers.sNumAuto
    objTRVVouchers.iQuantParc = tTRVVouchers.iQuantParc
    objTRVVouchers.idiasantc = tTRVVouchers.idiasantc
    objTRVVouchers.iKit = tTRVVouchers.iKit
    objTRVVouchers.lCodigoAporte = tTRVVouchers.lCodigoAporte
    objTRVVouchers.iParcelaAporte = tTRVVouchers.iParcelaAporte
    objTRVVouchers.iExtraiInfoSigav = tTRVVouchers.iExtraiInfoSigav
    objTRVVouchers.dValorOcr = tTRVVouchers.dValorOcr
    objTRVVouchers.iTemOcr = tTRVVouchers.iTemOcr
    objTRVVouchers.lRepresentante = tTRVVouchers.lRepresentante
    objTRVVouchers.dComissaoRep = tTRVVouchers.dComissaoRep
    objTRVVouchers.lCorrentista = tTRVVouchers.lCorrentista
    objTRVVouchers.dComissaoCorr = tTRVVouchers.dComissaoCorr
    objTRVVouchers.lEmissor = tTRVVouchers.lEmissor
    objTRVVouchers.dComissaoEmissor = tTRVVouchers.dComissaoEmissor
    objTRVVouchers.dComissaoAg = tTRVVouchers.dComissaoAg
    objTRVVouchers.dValorBruto = tTRVVouchers.dValorBruto
    objTRVVouchers.sPassageiroNome = tTRVVouchers.sPassageiroNome
    objTRVVouchers.sPassageiroSobreNome = tTRVVouchers.sPassageiroSobreNome
    objTRVVouchers.lClienteVou = tTRVVouchers.lClienteVou
    objTRVVouchers.lCliPassageiro = tTRVVouchers.lCliPassageiro
    objTRVVouchers.dtDataCanc = tTRVVouchers.dtDataCanc
    objTRVVouchers.dHoraCanc = tTRVVouchers.dHoraCanc
    objTRVVouchers.sUsuarioCanc = tTRVVouchers.sUsuarioCanc
    objTRVVouchers.lClienteComissao = tTRVVouchers.lClienteComissao
    objTRVVouchers.lPromotor = tTRVVouchers.lPromotor
    objTRVVouchers.sDestino = tTRVVouchers.sDestino
    objTRVVouchers.sTitularCPF = tTRVVouchers.sTitularCPF
    objTRVVouchers.dTarifaUnitaria = tTRVVouchers.dTarifaUnitaria
    objTRVVouchers.iVigencia = tTRVVouchers.iVigencia
    objTRVVouchers.sUsuarioWeb = tTRVVouchers.sUsuarioWeb
    objTRVVouchers.dValorBaseComis = tTRVVouchers.dValorBaseComis
    objTRVVouchers.dValorBrutoComOCR = tTRVVouchers.dValorBrutoComOCR
    objTRVVouchers.dValorCMAComOCR = tTRVVouchers.dValorCMAComOCR
    objTRVVouchers.dValorCMC = tTRVVouchers.dValorCMC
    objTRVVouchers.dValorCMR = tTRVVouchers.dValorCMR
    objTRVVouchers.dValorCMCC = tTRVVouchers.dValorCMCC
    objTRVVouchers.dValorCME = tTRVVouchers.dValorCME
    objTRVVouchers.sValidadeCC = tTRVVouchers.sValidadeCC
    objTRVVouchers.lNumIntDoc = tTRVVouchers.lNumIntDoc
    
    objTRVVouchers.lNumBoleto = tTRVVouchers.lNumBoleto
    objTRVVouchers.dtDataVencBoleto = tTRVVouchers.dtDataVencBoleto
    objTRVVouchers.dValorBoleto = tTRVVouchers.dValorBoleto
    objTRVVouchers.iTrataBoleto = tTRVVouchers.iTrataBoleto
    
    If objTRVVouchers.lClienteComissao = 0 Then objTRVVouchers.lClienteComissao = objTRVVouchers.lClienteVou

    lErro = TRVVouVendedores_Le(objTRVVouchers.colVendedores, objTRVVouchers, lComandoV)
    If lErro <> SUCESSO Then gError 190844

    'Fecha Comando
    If lComandoP1 = 0 Then Call Comando_Fechar(lComando)
    If lComandoP2 = 0 Then Call Comando_Fechar(lComandoV)

    TRVVouchers_Le = SUCESSO

    Exit Function

Erro_TRVVouchers_Le:

    TRVVouchers_Le = gErr

    Select Case gErr

        Case 190841
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 190842, 190843
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVVOUCHERS", gErr)

        Case 190844

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190844)

    End Select

    'Fecha Comando
    If lComandoP1 = 0 Then Call Comando_Fechar(lComando)
    If lComandoP2 = 0 Then Call Comando_Fechar(lComandoV)

    Exit Function

End Function

Private Function TRVAportePagtoFat_Le_NumIntDoc(ByVal lComando As Long, ByVal objTRVAportePagtoFat As ClassTRVAportePagtoFat) As Long

Dim lErro As Long
Dim tTRVAportePagtoFat As typeTRVAportePagtoFat

On Error GoTo Erro_TRVAportePagtoFat_Le_NumIntDoc

    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntDocAporte, Valor, ValidadeDe, ValidadeAte, Saldo, " & _
                    "Percentual FROM TRVAportePagtoFat WHERE NumIntDoc= ? ORDER BY Seq ", _
                    tTRVAportePagtoFat.lNumIntDoc, tTRVAportePagtoFat.lNumIntDocAporte, tTRVAportePagtoFat.dValor, tTRVAportePagtoFat.dtValidadeDe, _
                    tTRVAportePagtoFat.dtValidadeAte, tTRVAportePagtoFat.dSaldo, tTRVAportePagtoFat.dPercentual, _
                    objTRVAportePagtoFat.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 190950

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190951

    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS
    
    objTRVAportePagtoFat.lNumIntDoc = tTRVAportePagtoFat.lNumIntDoc
    objTRVAportePagtoFat.lNumIntDocAporte = tTRVAportePagtoFat.lNumIntDocAporte
    objTRVAportePagtoFat.dValor = tTRVAportePagtoFat.dValor
    objTRVAportePagtoFat.dtValidadeDe = tTRVAportePagtoFat.dtValidadeDe
    objTRVAportePagtoFat.dtValidadeAte = tTRVAportePagtoFat.dtValidadeAte
    objTRVAportePagtoFat.dSaldo = tTRVAportePagtoFat.dSaldo
    objTRVAportePagtoFat.dPercentual = tTRVAportePagtoFat.dPercentual

    TRVAportePagtoFat_Le_NumIntDoc = SUCESSO

    Exit Function

Erro_TRVAportePagtoFat_Le_NumIntDoc:

    TRVAportePagtoFat_Le_NumIntDoc = gErr

    Select Case gErr

        Case 190950, 190951
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVAPORTEPAGTOFAT", gErr)

        Case ERRO_LEITURA_SEM_DADOS
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190952)

    End Select

    Exit Function

End Function

Function TRVConfig_Le(ByVal sCodigo As String, ByVal iFilialEmpresa As Integer, sConteudo As String) As Long
'Lê o Conteudo da Tabela de TRVConfig de acordo com o código passado

Dim lErro As Long
Dim lComando As Long
On Error GoTo Erro_TRVConfig_Le
    
    'Abrir comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 192067

    sConteudo = String(STRING_CONFIG_CONTEUDO, 0)
        
    'ler registo
    lErro = Comando_Executar(lComando, "SELECT Conteudo FROM TRVConfig WHERE Codigo = ? AND FilialEmpresa = ?", sConteudo, sCodigo, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 192068
    
    'Lê o primeiro registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192069
    
    'Se não encontrou o registro
    If lErro = AD_SQL_SEM_DADOS Then gError 192070
    
    'Fechar comando
    Call Comando_Fechar(lComando)

    TRVConfig_Le = SUCESSO
    
    Exit Function

Erro_TRVConfig_Le:
        
    TRVConfig_Le = gErr
    
    Select Case gErr
    
        Case 192067
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 192068, 192069
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVCONFIG", gErr)

        Case 192070
            Call Rotina_Erro(vbOKOnly, "ERRO_TRVCONFIG_INEXISTENTE", gErr, sCodigo, iFilialEmpresa)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 192071)
     
    End Select
     
    'Fechar comando
    Call Comando_Fechar(lComando)
 
    Exit Function
    
End Function

Function TRVConfig_Le_Lockado(ByVal lComando As Long, ByVal sCodigo As String, ByVal iFilialEmpresa As Integer, sConteudo As String) As Long
'Lê o Conteudo da Tabela de TRVConfig de acordo com o código passado

Dim lErro As Long

On Error GoTo Erro_TRVConfig_Le_Lockado

    sConteudo = String(STRING_CONFIG_CONTEUDO, 0)
        
    'ler registo
    lErro = Comando_ExecutarPos(lComando, "SELECT Conteudo FROM TRVConfig WHERE Codigo = ? AND FilialEmpresa = ?", 0, sConteudo, sCodigo, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 194253
    
    'Lê o primeiro registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194254
    
    'Se não encontrou o registro
    If lErro = AD_SQL_SEM_DADOS Then gError 194255
    
    lErro = Comando_LockExclusive(lComando)
    If lErro <> SUCESSO Then gError 194254

    TRVConfig_Le_Lockado = SUCESSO
    
    Exit Function

Erro_TRVConfig_Le_Lockado:
        
    TRVConfig_Le_Lockado = gErr
    
    Select Case gErr
    
        Case 194253, 194254
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVCONFIG", gErr)

        Case 194255
            Call Rotina_Erro(vbOKOnly, "ERRO_TRVCONFIG_INEXISTENTE", gErr, sCodigo, iFilialEmpresa)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194256)
     
    End Select
 
    Exit Function
    
End Function

Public Function TRVGeracaoNF_Le_Dados(ByVal objGeracaoNF As ClassTRVGeracaoNF) As Long
'Le os itens a serem faturados com os filtros passados
 
Dim lErro As Long
Dim iIndice As Integer
Dim sSQL As String
Dim sWhere As String
Dim sWhereAux As String
Dim alComando(1 To 8) As Long
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
Dim objTitulosRecTRV As ClassTitulosRecTRV
Dim objNFFilial As ClassTRVGeracaoNFFiliais

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
Dim bGeraLog As Boolean, bArq1Aberto As Boolean, lSeq As Long, sConteudo As String, sDirLog As String

On Error GoTo Erro_TRVGeracaoNF_Le_Dados

    bArq1Aberto = False

    lErro = CF("TRVConfig_Le", TRVCONFIG_GERA_LOG, EMPRESA_TODA, sConteudo)
    If lErro <> SUCESSO Then gError 192082

    If StrParaInt(sConteudo) = MARCADO Then
        bGeraLog = True
    
        lErro = CF("TRVConfig_Le", TRVCONFIG_DIR_LOG, EMPRESA_TODA, sDirLog)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        Open sDirLog & "NF_" & gsUsuario & "_" & Format(Date, "yyyymmdd") & "_" & Format(Now, "hhmmss") & ".txt" For Output As #1
        bArq1Aberto = True
    Else
        bGeraLog = False
    End If
 
If bGeraLog Then Print #1, "1 - INICIANDO O PROCESSO " & CStr(Now)
 
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_AbrirExt(GL_lConexaoBrowse)
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

    If objGeracaoNF.iEmpresa = TRV_EMPRESA_AMBOS Then
        sSQL = "SELECT NumIntDoc,FilialEmpresa, NumTitulo, DataEmissao, Valor, NomeCliente, Cliente, ClienteVou, CliPassageiro, SiglaDocumento, NFAoPassageiro, UmaNFPorFatura FROM TitulosSemNota WHERE "
    ElseIf objGeracaoNF.iEmpresa = TRV_EMPRESA_MY Then
        sSQL = "SELECT NumIntDoc,FilialEmpresa, NumTitulo, DataEmissao, Valor, NomeCliente, Cliente, ClienteVou, CliPassageiro, SiglaDocumento, NFAoPassageiro, UmaNFPorFatura FROM TitulosSemNota WHERE Cliente >= 1000000 AND "
    Else
        sSQL = "SELECT NumIntDoc,FilialEmpresa, NumTitulo, DataEmissao, Valor, NomeCliente, Cliente, ClienteVou, CliPassageiro, SiglaDocumento, NFAoPassageiro, UmaNFPorFatura FROM TitulosSemNota WHERE Cliente < 1000000 AND "
    End If

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
    
    If objGeracaoNF.iTipo = TRV_GERACAONF_TITULOS_BAIXADO Then
        sWhere = sWhere & "AND Status = ? "
    End If
    
    If objGeracaoNF.sTipoDoc <> "" Then
        sWhere = sWhere & "AND SiglaDocumento = ? "
    End If
    
    objGeracaoNF.sFiltro = sWhere
    
    sSQL = sSQL & sWhere & " ORDER BY FilialEmpresa, Cliente"
    
If bGeraLog Then Print #1, "2 - EXECURANDO SQL " & CStr(Now) & " " & sSQL
    
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
    
    If objGeracaoNF.iTipo = TRV_GERACAONF_TITULOS_BAIXADO Then
        lErro = Comando_BindVarInt(alComando(1), viStatus)
        If (lErro <> AD_SQL_SUCESSO) Then gError 192226
    
        objGeracaoNF.colSelecao.Add viStatus
    End If
    
    If objGeracaoNF.sTipoDoc <> "" Then
        lErro = Comando_BindVarInt(alComando(1), vsTipoDoc)
        If (lErro <> AD_SQL_SUCESSO) Then gError 192306
    
        objGeracaoNF.colSelecao.Add vsTipoDoc
    End If
    
If bGeraLog Then Print #1, "3 - Comando_ExecutarInt " & CStr(Now)
    
    lErro = Comando_ExecutarInt(alComando(1))
    If (lErro <> AD_SQL_SUCESSO) Then gError 192227
 
If bGeraLog Then Print #1, "3 - Comando_BuscarPrimeiro " & CStr(Now)

    'Busca o primeiro registro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192228
        
    'Para cada fatura
    lSeq = 0
    Do While lErro = SUCESSO
    
        lSeq = lSeq + 1
    
If bGeraLog Then Print #1, "4." & CStr(lSeq) & " - No: " & CStr(vlNumTitulo) & " FE: " & CStr(viFilialEmpresa) & " " & CStr(Now)
    
        Set objTitulosRec = New ClassTituloReceber
        Set objTitulosRecTRV = New ClassTitulosRecTRV
    
        objTitulosRec.lNumIntDoc = vlNumIntDoc

        objTitulosRec.iFilialEmpresa = viFilialEmpresa
        objTitulosRec.lNumTitulo = vlNumTitulo
        objTitulosRec.dtDataEmissao = vdtDataEmissao
        objTitulosRec.dValor = vdValor
        objTitulosRec.sSiglaDocumento = vsSiglaDocumento
        objTitulosRec.lCliente = vlCliente
        objTitulosRec.iFilial = FILIAL_MATRIZ
        
        objTitulosRecTRV.lClienteVou = vlClienteVou
        objTitulosRecTRV.sNomeCliente = vsNomeCliente
        objTitulosRecTRV.lCliPassageiro = vlCliPassageiro
        objTitulosRecTRV.iNFAoPassageiro = viNFAoPassageiro
        objTitulosRecTRV.iUmaNFPorFatura = viUmaNFPorFatura
        
        If objTitulosRec.sSiglaDocumento = TIPODOC_CARTAO_CRED_DEB Then
            objTitulosRecTRV.iNFAoPassageiro = MARCADO
        End If

        Set objTitulosRec.objInfoUsu = objTitulosRecTRV
        
        objGeracaoNF.colTitulos.Add objTitulosRec
    
If bGeraLog Then Print #1, "5." & CStr(lSeq) & " - Comando_BuscarProximo " & CStr(Now)
    
        'Busca o próximo registro
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192229
    
    Loop
    
If bGeraLog Then Print #1, "6 - TRVGeracaoNF_Calcula_Dados " & CStr(Now)
    
    lErro = TRVGeracaoNF_Calcula_Dados(alComando(2), alComando(3), alComando(4), alComando(5), alComando(6), alComando(7), alComando(8), objGeracaoNF, bGeraLog)
    If lErro <> SUCESSO Then gError 192230
        
If bGeraLog Then Print #1, "7 - Comando_Fechar " & CStr(Now)
        
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
 
If bGeraLog Then Print #1, "8 - FIM DO PROCESSO " & CStr(Now)

    If bArq1Aberto Then Close #1
 
    TRVGeracaoNF_Le_Dados = SUCESSO
 
    Exit Function
 
Erro_TRVGeracaoNF_Le_Dados:
 
    TRVGeracaoNF_Le_Dados = gErr
 
If bGeraLog Then Print #1, "9 - ERRO: " & CStr(gErr) & " " & CStr(Now)
 
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
    
    If bArq1Aberto Then Close #1
 
    Exit Function
 
End Function

Private Function TRVGeracaoNF_Calcula_Dados(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lComando3 As Long, ByVal lComando4 As Long, ByVal lComando5 As Long, ByVal lComando6 As Long, ByVal lComando7 As Long, ByVal objGeracaoNF As ClassTRVGeracaoNF, Optional ByVal bGeraLog As Boolean = False) As Long
 
Dim lErro As Long
Dim objTitulosRec As ClassTituloReceber
Dim objTitulosRecTRV As ClassTitulosRecTRV
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
Dim objGeracaoNFItem As ClassTRVGeracaoNFItem
Dim objSerie As ClassSerie
Dim lCliente As Long
Dim iFilialCliente As Integer
Dim objNFTRV As ClassNFTRV
Dim bPrimeiroItem As Boolean
Dim objTipoTribICMS As New ClassTipoTribICMS
Dim objTipoTribIPI As New ClassTipoTribIPI
Dim objTipoDeTributacao As New ClassTipoDeTributacaoMovto
Dim colCli As New Collection, objFilialClienteAux As ClassFilialCliente, colEnd As New Collection, objEndereco As ClassEndereco
Dim objFilialCliente As ClassFilialCliente, bAchou As Boolean, sTextoCliSemCGC As String, iAux As Integer

Const NATUREZAOP_TRV = "5949"
Const TIPO_TRIBUTACAO_TRV = 3
Const SERIE_TRV = "1"
Const PRODUTO_TRV = "FAT       "
Dim colSaida As New Collection
Dim colCampos As New Collection
Dim giFilialEmpresAux As Integer
Dim dPercTotImp As Double
Dim sUFFilialEmpresa As String, dTotTribAliqFedAux As Double, dTotTribAliqEstAux As Double, dTotTribAliqMunicAux As Double
Dim objFilEmp As AdmFiliais, colFilEmp As New Collection, bAchouFilEmp As Boolean

On Error GoTo Erro_TRVGeracaoNF_Calcula_Dados

If bGeraLog Then Print #1, "6.1 - INICIANDO TRVGeracaoNF_Calcula_Dados " & CStr(Now)

    bPrimeiroItem = True
    giFilialEmpresAux = giFilialEmpresa

    'Le o tipo de nota que vai ser usado
    objTipoDocInfo.iCodigo = DOCINFO_NFISS
    
If bGeraLog Then Print #1, "6.2 - TipoDocInfo_Le_Codigo " & CStr(Now)
    
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 192235
    
    'Le o produto que vai ser usado nos itens da Nota
    objProduto.sCodigo = PRODUTO_TRV
    
If bGeraLog Then Print #1, "6.3 - Produto_Le " & CStr(Now)
    
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 192236
    
    'Se o produto não existe => Erro
    If lErro <> SUCESSO Then gError 192237
    
If bGeraLog Then Print #1, "6.4 - Produto_Le_Categorias " & CStr(Now)
    
    lErro = CF("Produto_Le_Categorias", objProduto, objProduto.colCategoriaItem)
    If lErro <> SUCESSO Then gError 192236
        
    'Se for uma nota fiscal de serviço
    If objTipoDocInfo.iCodigo = DOCINFO_NFISS Then
        If Not (objProduto.iCompras = PRODUTO_PRODUZIVEL And objProduto.iFaturamento = PRODUTO_VENDAVEL And objProduto.iControleEstoque = PRODUTO_CONTROLE_SEM_ESTOQUE) Then gError 192238
    End If
    
'If bGeraLog Then Print #1, "6.5 - IBPTax_ObterPerc " & CStr(Now)
'
'    lErro = CF("IBPTax_ObterPerc", giFilialEmpresa, objGeracaoNF.dtDataEmissao, objProduto.sCodigo, objProduto.iOrigemMercadoria, dPercTotImp, mvarsUFFilialEmpresa, dTotTribAliqFedAux, dTotTribAliqEstAux, dTotTribAliqMunicAux)
'    If lErro <> SUCESSO And lErro <> ERRO_ITEM_NAO_CADASTRADO Then gError ERRO_SEM_MENSAGEM

    '#########################################################################################
    'Preenche nota e itens
    For iIndice = objGeracaoNF.colTitulos.Count To 1 Step -1
    
        Set objTitulosRec = objGeracaoNF.colTitulos.Item(iIndice)
    
If bGeraLog Then Print #1, "6.6." & CStr(iIndice) & " - No: " & CStr(objTitulosRec.lNumTitulo) & " " & CStr(Now)
    
        'Verifica se o cliente não deve ser um passageiro
        lErro = TRV_GeracaoNF_Obtem_Cliente(lComando1, lComando2, objGeracaoNF, objTitulosRec, lCliente, iFilialCliente)
        If lErro <> SUCESSO Then gError 192727
        
        objTitulosRec.lCliente = lCliente
        objTitulosRec.iFilial = iFilialCliente
        
        Set objFilialCliente = New ClassFilialCliente
        objFilialCliente.lCodCliente = lCliente
        objFilialCliente.iCodFilial = iFilialCliente
        
If bGeraLog Then Print #1, "6.7." & CStr(iIndice) & " - Cliente: " & CStr(lCliente) & " " & CStr(Now)
        
        bAchou = False
        iAux = 0
        For Each objFilialClienteAux In colCli
            iAux = iAux + 1
            Set objEndereco = colEnd(iAux)
            If objFilialCliente.lCodCliente = objFilialClienteAux.lCodCliente And objFilialCliente.iCodFilial = objFilialClienteAux.iCodFilial Then
                bAchou = True
                objFilialClienteAux.sObservacao = objFilialClienteAux.sObservacao & "|" & CStr(objTitulosRec.lNumTitulo)
                If Len(Trim(objFilialClienteAux.sCgc)) < 11 And objEndereco.iCodigoPais = PAIS_BRASIL And Len(Trim(objEndereco.sSiglaEstado)) <> "" And Len(Trim(objEndereco.sSiglaEstado)) <> "EX" Then
                    objGeracaoNF.colTitulos.Remove iIndice
                End If
                Exit For
            End If
        Next
        If Not bAchou Then
            
If bGeraLog Then Print #1, "6.8." & CStr(iIndice) & " - FilialCliente_Le " & CStr(Now)
            
            lErro = CF("FilialCliente_Le", objFilialCliente, lComando7)
            If lErro <> SUCESSO And lErro <> 12567 Then gError 192727
            
            Set objEndereco = New ClassEndereco
            
            objEndereco.lCodigo = objFilialCliente.lEndereco
            
If bGeraLog Then Print #1, "6.9." & CStr(iIndice) & " - Endereco_Le " & CStr(Now)
            
            lErro = CF("Endereco_Le", objEndereco)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            colCli.Add objFilialCliente
            colEnd.Add objEndereco
            
            objFilialCliente.sObservacao = CStr(objTitulosRec.lNumTitulo)
            If Len(Trim(objFilialCliente.sCgc)) < 11 And objEndereco.iCodigoPais = PAIS_BRASIL And Len(Trim(objEndereco.sSiglaEstado)) <> "" And Len(Trim(objEndereco.sSiglaEstado)) <> "EX" Then
                objGeracaoNF.colTitulos.Remove iIndice
            End If
        End If
    
    Next
    
If bGeraLog Then Print #1, "6.10 - FIM DO LOOP " & CStr(Now)
    
    sTextoCliSemCGC = ""
    bAchou = False
    iAux = 0
    For Each objFilialCliente In colCli
        iAux = iAux + 1
        Set objEndereco = colEnd(iAux)
        If Len(Trim(objFilialCliente.sCgc)) < 11 And objEndereco.iCodigoPais = PAIS_BRASIL And Len(Trim(objEndereco.sSiglaEstado)) <> "" And Len(Trim(objEndereco.sSiglaEstado)) <> "EX" Then
            bAchou = True
            sTextoCliSemCGC = sTextoCliSemCGC & IIf(Len(Trim(sTextoCliSemCGC)) > 0, ";", "") & CStr(objFilialCliente.lCodCliente) & "(" & CStr(objFilialCliente.sObservacao) & ")"
        End If
    Next
    
If bGeraLog Then Print #1, "6.11 - sTextoCliSemCGC: " & sTextoCliSemCGC & " " & CStr(Now)
    
    If bAchou Then 'gError 209340
        Call Rotina_Aviso(vbOKOnly, "ERRO_TRV_CLI_SEM_CGC", sTextoCliSemCGC)
    End If
    
    colCampos.Add "iFilialEmpresa"
    colCampos.Add "lCliente"
    colCampos.Add "iFilial"
    colCampos.Add "lNumTitulo"
    
If bGeraLog Then Print #1, "6.12 - Ordena_Colecao " & CStr(Now)
    
    Call Ordena_Colecao(objGeracaoNF.colTitulos, colSaida, colCampos)
    
    For iIndice = colSaida.Count To 1 Step -1
        objGeracaoNF.colTitulos.Remove (iIndice)
    Next
    
    For Each objTitulosRec In colSaida
        objGeracaoNF.colTitulos.Add objTitulosRec
    Next
        
    iIndice = 0
    For Each objTitulosRec In objGeracaoNF.colTitulos
    
        iIndice = iIndice + 1
        
If bGeraLog Then Print #1, "6.13." & CStr(iIndice) & " - No: " & CStr(objTitulosRec.lNumTitulo) & " " & CStr(Now)
        
        lCliente = objTitulosRec.lCliente
        iFilialCliente = objTitulosRec.iFilial
        
        Set objTitulosRecTRV = objTitulosRec.objInfoUsu
        
        iFilialEmpresa = objTitulosRec.iFilialEmpresa
        
        If lClienteAnt <> lCliente Or iFilialEmpresaAnt <> iFilialEmpresa Or (objTitulosRecTRV.iUmaNFPorFatura = MARCADO Or objGeracaoNF.iGerarNFParaCadaFat = MARCADO) Then
        
            Set objNF = New ClassNFiscal
            Set objNFTRV = New ClassNFTRV
            Set objNF.objInfoUsu = objNFTRV
            
            Set objNFTRV.objProduto = objProduto
            Set objNFTRV.objTipoDocInfo = objTipoDocInfo
            
            objNF.iFilialEmpresa = iFilialEmpresa
            objNF.sSerie = SERIE_TRV
                       
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
            objNF.sNaturezaOp = NATUREZAOP_TRV
                   
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
          
If bGeraLog Then Print #1, "6.14 - FIM DO LOOP " & CStr(Now)
          
    '#########################################################################################
    'Calcula as informações por filial antes da redis.
    iFilialEmpresaAnt = 0
    iIndice = 0
    For Each objNF In objGeracaoNF.colNF
 
        iIndice = iIndice + 1
        
If bGeraLog Then Print #1, "6.15." & CStr(iIndice) & " - LOOP NF " & CStr(Now)
        
        If iFilialEmpresaAnt <> objNF.iFilialEmpresa Then
        
            If iIndice <> 1 Then
            
                Set objGeracaoNFItem = New ClassTRVGeracaoNFItem
                
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
    
If bGeraLog Then Print #1, "6.16 - FIM DO LOOP " & CStr(Now)
    
    Set objGeracaoNFItem = New ClassTRVGeracaoNFItem
    
    objGeracaoNFItem.dValor = dValor
    objGeracaoNFItem.iFilialEmpresa = iFilialEmpresaAnt
    
    objGeracaoNF.colItens.Add objGeracaoNFItem
    '#########################################################################################
    
If bGeraLog Then Print #1, "6.17 - TRVGeracaoNF_Calcula_Dados2 " & CStr(Now)
    
    'Redis. as NFs pelas filiais
    lErro = TRVGeracaoNF_Calcula_Dados2(lComando3, lComando4, lComando5, lComando6, objGeracaoNF, bGeraLog)
    If lErro <> SUCESSO Then gError 192241
    
    '#########################################################################################
    'Acerta a parte tributária
    iIndice = 0
    For Each objNF In objGeracaoNF.colNF
    
    iIndice = iIndice + 1
    
If bGeraLog Then Print #1, "6.18." & CStr(iIndice) & " - LOOP NF Inicializa_Tributacao " & CStr(Now)
    
        giFilialEmpresa = objNF.iFilialEmpresa
    
        'Inicializa a Parte tributária da Nota
        Call objNF.Inicializa_Tributacao
        Call objNF.objTributacaoNF.Coloca_Auto
        
        objNF.objTributacaoNF.iTipoTributacao = TIPO_TRIBUTACAO_TRV
        objNF.objTributacaoNF.iTipoTributacaoManual = VAR_PREENCH_AUTOMATICO
        objNF.objTributacaoNF.iNaturezaOpManual = VAR_PREENCH_MANUAL
        objNF.objTributacaoNF.sNaturezaOpInterna = NATUREZAOP_TRV
        objNF.objTributacaoNF.iNaturezaOpManual = VAR_PREENCH_AUTOMATICO
        
        objNF.objTributacaoNF.iTotTribTipo = 1
        objNF.objTributacaoNF.iTotTribTipoManual = VAR_PREENCH_MANUAL
        
If bGeraLog Then Print #1, "6.19 - IBPTax_ObterPerc " & CStr(Now)

        bAchouFilEmp = False
        For Each objFilEmp In colFilEmp
            If objNF.iFilialEmpresa = objFilEmp.iCodFilial Then
                bAchouFilEmp = True
                dPercTotImp = objFilEmp.dAliquotaPadraoContribPrev
            End If
        Next
        If Not bAchouFilEmp Then
        
            Set objFilEmp = New AdmFiliais
            
            objFilEmp.iCodFilial = objNF.iFilialEmpresa
            
            lErro = CF("FilialEmpresa_Le", objFilEmp)
            If lErro <> SUCESSO And lErro <> 27378 Then gError ERRO_SEM_MENSAGEM
       
            sUFFilialEmpresa = objFilEmp.objEndereco.sSiglaEstado
   
            lErro = CF("IBPTax_ObterPerc", giFilialEmpresa, objGeracaoNF.dtDataEmissao, objProduto, objProduto.iOrigemMercadoria, dPercTotImp, sUFFilialEmpresa, dTotTribAliqFedAux, dTotTribAliqEstAux, dTotTribAliqMunicAux)
            If lErro <> SUCESSO And lErro <> ERRO_ITEM_NAO_CADASTRADO Then gError ERRO_SEM_MENSAGEM
            
            objFilEmp.dAliquotaPadraoContribPrev = dPercTotImp
        
            colFilEmp.Add objFilEmp
        
        End If
        
        If dPercTotImp > DELTA_VALORMONETARIO Then
            If objNF.sMensagemNota <> "" Then objNF.sMensagemNota = objNF.sMensagemNota & vbNewLine
            objNF.sMensagemNota = objNF.sMensagemNota & "Total de impostos pagos: R$ " & Format(objNF.dValorTotal * dPercTotImp, "STANDARD") & " (" & Format(dPercTotImp, "PERCENT") & ") Fonte: IBPT"
        End If
                
        For Each objItemNF In objNF.ColItensNF
        
            'Inicializa parte tributária do item da Nota
            Call objItemNF.Inicializa_Tributacao
            Call objItemNF.objTributacaoItemNF.Coloca_Auto
           
            objItemNF.objTributacaoItemNF.iTipoTributacao = TIPO_TRIBUTACAO_TRV
            objItemNF.objTributacaoItemNF.iTipoTributacaoManual = VAR_PREENCH_AUTOMATICO
            objItemNF.objTributacaoItemNF.sNaturezaOp = objNF.objTributacaoNF.sNaturezaOpInterna
            objItemNF.objTributacaoItemNF.iNaturezaOpManual = VAR_PREENCH_AUTOMATICO
            
            objItemNF.objTributacaoItemNF.sProduto = objProduto.sCodigo
            objItemNF.objTributacaoItemNF.iExTIPI = objProduto.iExTIPI
            objItemNF.objTributacaoItemNF.sGenero = objProduto.sGenero
            objItemNF.objTributacaoItemNF.iProdutoEspecifico = objProduto.iProdutoEspecifico
            objItemNF.objTributacaoItemNF.sUMTrib = objProduto.sSiglaUMTrib
            objItemNF.objTributacaoItemNF.sIPICodProduto = objProduto.sIPICodigo

            objItemNF.objTributacaoItemNF.dQtdTrib = objItemNF.dQuantidade
            objItemNF.objTributacaoItemNF.dValorUnitTrib = objItemNF.dPrecoUnitario
            objItemNF.objTributacaoItemNF.dDescontoGrid = objItemNF.dValorDesconto
            objItemNF.objTributacaoItemNF.dPrecoTotal = objItemNF.dValorTotal
            objItemNF.objTributacaoItemNF.sProdutoDescricao = objProduto.sDescricao
            objItemNF.objTributacaoItemNF.sISSQN = objProduto.sISSQN
            
'            Select Case objNF.iFilialEmpresa
'                Case 1
'                    objItemNF.objTributacaoItemNF.sISSCidadeIBGE = "4314902"
'                Case 2
'                    objItemNF.objTributacaoItemNF.sISSCidadeIBGE = "1200401"
'                Case 3
'                    objItemNF.objTributacaoItemNF.sISSCidadeIBGE = "4314902"
'                Case 4
'                    objItemNF.objTributacaoItemNF.sISSCidadeIBGE = "5300108"
'                Case 5
'                    objItemNF.objTributacaoItemNF.sISSCidadeIBGE = "4314902"
'                Case 6
'                    objItemNF.objTributacaoItemNF.sISSCidadeIBGE = "2704302"
'            End Select
            
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

If bGeraLog Then Print #1, "6.19." & CStr(iIndice) & " - TributacaoNF_Reset " & CStr(Now)
        
        lErro = objTribTab.TributacaoNF_Reset(objNF)
        If lErro <> SUCESSO Then gError 192239
        
        'Atualiza a Tributação
        Call objTribTab.TipoNFiscal_Definir(objTipoDocInfo.iCodigo, objTipoDocInfo.sSigla)
        Call objTribTab.DataEmissao_Alterada(objNF.dtDataEmissao)
        Call objTribTab.Serie_Alterada(objNF.sSerie)
        
If bGeraLog Then Print #1, "6.20." & CStr(iIndice) & " - AtualizaImpostos " & CStr(Now)
        
        lErro = gobjTributacao.AtualizaImpostos(objTribTab.mvarobjNFTributacao, 0)
        If lErro <> SUCESSO Then gError 192240
        
If bGeraLog Then Print #1, "6.21." & CStr(iIndice) & " - PreencherNF " & CStr(Now)
        
        lErro = objTribTab.PreencherNF(objNF)
        If lErro <> SUCESSO Then gError 192241
        
        Set objTribTab = Nothing
        
        Set objNFTRV.objTipoDeTributacaoMovto = objTipoDeTributacao
        Set objNFTRV.objTipoTribIPI = objTipoTribIPI
        Set objNFTRV.objTipoTribICMS = objTipoTribICMS
        
        objNF.sNaturezaOp = objNF.objTributacaoNF.sNaturezaOpInterna
        
        'acertar valor total considerando os tributos, retenções,...
        objNF.dValorTotal = objNF.dValorTotal + objNF.objTributacaoNF.dICMSSubstValor + objNF.objTributacaoNF.dIPIValor + objNF.dValorSeguro + objNF.dValorOutrasDespesas + IIf(objNF.objTributacaoNF.iISSIncluso = 0, objNF.objTributacaoNF.dISSValor, 0)
    
    Next
    
If bGeraLog Then Print #1, "6.22 - FIM LOOP NF " & CStr(Now)
    
    giFilialEmpresa = giFilialEmpresAux
    '#########################################################################################
    
    '#########################################################################################
    'Acerta a faixa da numeração das Nfs por filial
    For Each objGeracaoNFItem In objGeracaoNF.colItens
    
        Set objSerie = New ClassSerie
        
        objSerie.iFilialEmpresa = objGeracaoNFItem.iFilialEmpresa
        objSerie.sSerie = SERIE_TRV
        
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
    
    giFilialEmpresa = giFilialEmpresAux
        
If bGeraLog Then Print #1, "6.23 - FIM TRVGeracaoNF_Calcula_Dados " & CStr(Now)
        
    TRVGeracaoNF_Calcula_Dados = SUCESSO
 
    Exit Function
 
Erro_TRVGeracaoNF_Calcula_Dados:

    giFilialEmpresa = giFilialEmpresAux
 
    TRVGeracaoNF_Calcula_Dados = gErr
 
    Select Case gErr
    
        Case 192235, 192236, 192239 To 192242, 192727, 196790, 196791, 196792
        
        Case ERRO_SEM_MENSAGEM
        
        Case 192237
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
        
        Case 192238
             Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NOTA_SERVICO", gErr, objProduto.sCodigo)
             
        Case 209340
             Call Rotina_Erro(vbOKOnly, "ERRO_TRV_CLI_SEM_CGC", gErr, sTextoCliSemCGC)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192232)
 
    End Select
 
    Exit Function
 
End Function

Public Function TRVGeracaoNF_Calcula_Dados2(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lComando3 As Long, ByVal lComando4 As Long, ByVal objGeracaoNF As ClassTRVGeracaoNF, Optional ByVal bGeraLog As Boolean) As Long
'Redistribui as NFs conforme percentual
 
Dim lErro As Long
Dim iIndice As Integer
Dim objGeracaoNFItem As ClassTRVGeracaoNFItem
Dim iFilialO As Integer
Dim iFilialD As Integer
Dim objNFFilial As ClassTRVGeracaoNFFiliais
Dim dValorARedis As Double
Dim dValorFaltaRedis As Double
Dim dValorPodeRedis As Double
Dim objCliente As ClassCliente
Dim objClienteTRV As ClassClienteTRV
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
Dim objTituloRecTRV As ClassTitulosRecTRV
Dim objNFTRV As ClassNFTRV
Dim alCount(1 To 7) As Long
Dim objMensagem As ClassMensagem

Const MARGEM_ERRO = 0.01
Const FILIAL_BH = 6
Const NUM_TENTATIVAS = 3

On Error GoTo Erro_TRVGeracaoNF_Calcula_Dados2

If bGeraLog Then Print #1, "6.17.1 - INICIANDO TRVGeracaoNF_Calcula_Dados2 " & CStr(Now)

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
            
            Set objClienteTRV = objCliente.objInfoUsu
            
            If objClienteTRV.iFilialNF <> 0 Then
                objNF.iFilialEmpresa = objClienteTRV.iFilialNF
            End If
        
        End If
        
        Set objNFTRV = objNF.objInfoUsu
        Set objNFTRV.objCliente = objCliente
        
        If objCliente.iCodMensagem <> 0 Then
        
            Set objMensagem = New ClassMensagem
        
            objMensagem.iCodigo = objCliente.iCodMensagem
            
            'Lê a Mensagem "default" p/ o Cliente
            lErro = CF("Mensagem_Le", objMensagem)
            If lErro <> SUCESSO And lErro <> 19234 Then gError 196531
            
            If objNF.sMensagemNota <> "" Then objNF.sMensagemNota = objNF.sMensagemNota & vbNewLine
            objNF.sMensagemNota = objNF.sMensagemNota & objMensagem.sDescricao
        
        End If
        
        objGeracaoNF.colNFClientes.Add objCliente

    Next
    
If bGeraLog Then Print #1, "6.17.2 - For Each objNFFilial In objGeracaoNF.colFiliais " & CStr(Now)
    
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
            
            Set objGeracaoNFItem = New ClassTRVGeracaoNFItem
            
            objGeracaoNFItem.iFilialEmpresa = objNFFilial.iFilialEmpresa
            
            objGeracaoNF.colItens.Add objGeracaoNFItem
        
        End If
            
    Next

If bGeraLog Then Print #1, "6.17.3 - For Each objNFFilial In objGeracaoNF.colFiliais " & CStr(Now)

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
                    
                    Set objTituloRecTRV = objTituloRec.objInfoUsu
            
                    iIndice = iIndice + 1
                    If objNF.iFilialEmpresa = objNFFilial.iFilialEmpresa Then
                        
                        Set objCliente = objGeracaoNF.colNFClientes.Item(iIndice)
                        iFilialO = objGeracaoNF.colNFFilialO.Item(iIndice)
                        
                        'Se não é venda direta, se não é passageiro (exceto cartão) e se não é CVC e se não trocou a filial antes
                        If objCliente.iTipo <> 1 And (objCliente.iTipo <> 10 Or objTituloRec.sSiglaDocumento = TIPODOC_CARTAO_CRED_DEB) And InStr(1, objCliente.sNomeReduzido, "CVC") = 0 And InStr(1, objTituloRecTRV.sNomeCliente, "CVC") = 0 And iFilialO = objNF.iFilialEmpresa Then
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
                                    objNF.iFilialEmpresa = FILIAL_BH
                                    
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
    
If bGeraLog Then Print #1, "6.17.4 - For Each objNF In objGeracaoNF.colNF " & CStr(Now)
    
    iIndiceTit = 0
    For Each objNF In objGeracaoNF.colNF
        For Each objItemNF In objNF.ColItensNF
            iIndiceTit = iIndiceTit + 1
            Set objTituloRec = objGeracaoNF.colTitulos.Item(iIndiceTit)
            If objNF.iFilialEmpresa = 1 Or objNF.iFilialEmpresa = 2 Or objNF.iFilialEmpresa = 6 Then
                objItemNF.sDescricaoItem = "Prestação de serviço de assistência ao viajante, referente a(s) fatura(s) Nro. " & CStr(objTituloRec.lNumTitulo)
            Else
                objItemNF.sDescricaoItem = "Prestação de Serv. de assistência ao viajante, ref. fatura(s) Nro. " & CStr(objTituloRec.lNumTitulo)
            End If
        Next
    Next
     
    'Acerta a filial de destino
    For Each objNF In objGeracaoNF.colNF
        objGeracaoNF.colNFFilialD.Add objNF.iFilialEmpresa
    Next

If bGeraLog Then Print #1, "6.17.5 - For Each objNF In objGeracaoNF.colNF " & CStr(Now)

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
'
'    For Each objNF In objGeracaoNF.colNF
'        If objNF.iFilialEmpresa = 1 Then
'            alCount(1) = alCount(1) + 1
'        ElseIf objNF.iFilialEmpresa = 2 Then
'            alCount(2) = alCount(2) + 1
'        ElseIf objNF.iFilialEmpresa = 3 Then
'            alCount(3) = alCount(3) + 1
'        ElseIf objNF.iFilialEmpresa = 4 Then
'            alCount(4) = alCount(4) + 1
'        ElseIf objNF.iFilialEmpresa = 5 Then
'            alCount(5) = alCount(5) + 1
'        ElseIf objNF.iFilialEmpresa = 6 Then
'            alCount(6) = alCount(6) + 1
'        Else
'            alCount(7) = alCount(7) + 1
'            MsgBox "NF Filial " & CStr(objNF.iFilialEmpresa) & " Cliente " & objNF.lCliente & " Item: " & objNF.ColItensNF.Item(1).sDescricaoItem
'        End If
'    Next
'    MsgBox "F1 = " & CStr(alCount(1)) & " | F2 = " & CStr(alCount(2)) & " | F3 = " & CStr(alCount(3)) & " | F4 = " & CStr(alCount(4)) & " | F5 = " & CStr(alCount(5)) & " | F6 = " & CStr(alCount(6)) & " | F7 = " & CStr(alCount(7))
'
    TRVGeracaoNF_Calcula_Dados2 = SUCESSO
 
    Exit Function
 
Erro_TRVGeracaoNF_Calcula_Dados2:
 
    TRVGeracaoNF_Calcula_Dados2 = gErr
 
    Select Case gErr
    
        Case 196530, 196531
            
        Case 192299
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_ITEM_ERRADO", gErr, Format(objItemNF.dValorTotal, "STANDARD"), Format(objTituloRec.dValor, "STANDARD"))
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196532)
 
    End Select
 
    Exit Function
 
End Function

Public Function TRVAportePagtoCond_Le_NumIntDocDestino(ByVal objTRVAportePagto As ClassTRVAportePagtoCond) As Long

Dim lErro As Long
Dim lComando As Long
Dim tTRVAportePagtoCond As typeTRVAportePagtoCond

On Error GoTo Erro_TRVAportePagtoCond_Le_NumIntDocDestino

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 192341

    'Le a tabela TRVAportePagtoCond
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntDocAporte, Base, Status, DataPagto, NumIntDocDestino, " & _
                "FormaPagto, TipoDocDestino, Percentual, Valor FROM TRVAportePagtoCond WHERE NumIntDocDestino = ? AND TipoDocDestino = ?  ", _
                tTRVAportePagtoCond.lNumIntDoc, tTRVAportePagtoCond.lNumIntDocAporte, tTRVAportePagtoCond.iBase, tTRVAportePagtoCond.iStatus, _
                tTRVAportePagtoCond.dtDataPagto, tTRVAportePagtoCond.lNumIntDocDestino, tTRVAportePagtoCond.iFormaPagto, tTRVAportePagtoCond.iTipoDocDestino, tTRVAportePagtoCond.dPercentual, tTRVAportePagtoCond.dValor, _
                objTRVAportePagto.lNumIntDocDestino, objTRVAportePagto.iTipoDocDestino)
    If lErro <> AD_SQL_SUCESSO Then gError 192342

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192343

    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS
    
    objTRVAportePagto.lNumIntDoc = tTRVAportePagtoCond.lNumIntDoc
    objTRVAportePagto.lNumIntDocAporte = tTRVAportePagtoCond.lNumIntDocAporte
    objTRVAportePagto.iBase = tTRVAportePagtoCond.iBase
    objTRVAportePagto.iStatus = tTRVAportePagtoCond.iStatus
    objTRVAportePagto.dtDataPagto = tTRVAportePagtoCond.dtDataPagto
    objTRVAportePagto.lNumIntDocDestino = tTRVAportePagtoCond.lNumIntDocDestino
    objTRVAportePagto.iFormaPagto = tTRVAportePagtoCond.iFormaPagto
    objTRVAportePagto.iTipoDocDestino = tTRVAportePagtoCond.iTipoDocDestino
    objTRVAportePagto.dPercentual = tTRVAportePagtoCond.dPercentual
    objTRVAportePagto.dValor = tTRVAportePagtoCond.dValor
    
    Call Comando_Fechar(lComando)
    
    TRVAportePagtoCond_Le_NumIntDocDestino = SUCESSO

    Exit Function

Erro_TRVAportePagtoCond_Le_NumIntDocDestino:

    TRVAportePagtoCond_Le_NumIntDocDestino = gErr

    Select Case gErr
    
        Case 192341
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 192342, 192343
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVAPORTEPAGTOCOND", gErr)

        Case ERRO_LEITURA_SEM_DADOS
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192344)

    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function TRVAportePagtoDireto_Le_NumIntDocDestino(ByVal objTRVAportePagtoDireto As ClassTRVAportePagtoDireto) As Long

Dim lErro As Long
Dim tTRVAportePagtoDireto As typeTRVAportePagtoDireto
Dim lComando As Long

On Error GoTo Erro_TRVAportePagtoDireto_Le_NumIntDocDestino

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 192345
    
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntDocAporte, Valor, Vencimento, NumIntDocDestino, FormaPagto, " & _
                    "TipoDocDestino FROM TRVAportePagtoDireto WHERE NumIntDocDestino= ? AND TipoDocDestino = ? ", _
                    tTRVAportePagtoDireto.lNumIntDoc, tTRVAportePagtoDireto.lNumIntDocAporte, tTRVAportePagtoDireto.dValor, tTRVAportePagtoDireto.dtVencimento, _
                    tTRVAportePagtoDireto.lNumIntDocDestino, tTRVAportePagtoDireto.iFormaPagto, tTRVAportePagtoDireto.iTipoDocDestino, _
                    objTRVAportePagtoDireto.lNumIntDocDestino, objTRVAportePagtoDireto.iTipoDocDestino)
    If lErro <> AD_SQL_SUCESSO Then gError 192346

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192347
    
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objTRVAportePagtoDireto.lNumIntDoc = tTRVAportePagtoDireto.lNumIntDoc
    objTRVAportePagtoDireto.lNumIntDocAporte = tTRVAportePagtoDireto.lNumIntDocAporte
    objTRVAportePagtoDireto.dValor = tTRVAportePagtoDireto.dValor
    objTRVAportePagtoDireto.dtVencimento = tTRVAportePagtoDireto.dtVencimento
    objTRVAportePagtoDireto.lNumIntDocDestino = tTRVAportePagtoDireto.lNumIntDocDestino
    objTRVAportePagtoDireto.iFormaPagto = tTRVAportePagtoDireto.iFormaPagto
    objTRVAportePagtoDireto.iTipoDocDestino = tTRVAportePagtoDireto.iTipoDocDestino
    
    Call Comando_Fechar(lComando)

    TRVAportePagtoDireto_Le_NumIntDocDestino = SUCESSO

    Exit Function

Erro_TRVAportePagtoDireto_Le_NumIntDocDestino:

    TRVAportePagtoDireto_Le_NumIntDocDestino = gErr

    Select Case gErr
    
        Case 192345
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 192346, 192347
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVAPORTEPAGTODIRETO", gErr)
            
        Case ERRO_LEITURA_SEM_DADOS

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192348)

    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function TRVAportePagtoDireto_Le_NumIntDoc(ByVal objTRVAportePagtoDireto As ClassTRVAportePagtoDireto) As Long

Dim lErro As Long
Dim tTRVAportePagtoDireto As typeTRVAportePagtoDireto
Dim lComando As Long

On Error GoTo Erro_TRVAportePagtoDireto_Le_NumIntDoc

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 192364
    
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntDocAporte, Valor, Vencimento, NumIntDocDestino, FormaPagto, " & _
                    "TipoDocDestino FROM TRVAportePagtoDireto WHERE NumIntDoc = ? ", _
                    tTRVAportePagtoDireto.lNumIntDoc, tTRVAportePagtoDireto.lNumIntDocAporte, tTRVAportePagtoDireto.dValor, tTRVAportePagtoDireto.dtVencimento, _
                    tTRVAportePagtoDireto.lNumIntDocDestino, tTRVAportePagtoDireto.iFormaPagto, tTRVAportePagtoDireto.iTipoDocDestino, _
                    objTRVAportePagtoDireto.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 192365

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192366
    
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objTRVAportePagtoDireto.lNumIntDoc = tTRVAportePagtoDireto.lNumIntDoc
    objTRVAportePagtoDireto.lNumIntDocAporte = tTRVAportePagtoDireto.lNumIntDocAporte
    objTRVAportePagtoDireto.dValor = tTRVAportePagtoDireto.dValor
    objTRVAportePagtoDireto.dtVencimento = tTRVAportePagtoDireto.dtVencimento
    objTRVAportePagtoDireto.lNumIntDocDestino = tTRVAportePagtoDireto.lNumIntDocDestino
    objTRVAportePagtoDireto.iFormaPagto = tTRVAportePagtoDireto.iFormaPagto
    objTRVAportePagtoDireto.iTipoDocDestino = tTRVAportePagtoDireto.iTipoDocDestino
    
    Call Comando_Fechar(lComando)

    TRVAportePagtoDireto_Le_NumIntDoc = SUCESSO

    Exit Function

Erro_TRVAportePagtoDireto_Le_NumIntDoc:

    TRVAportePagtoDireto_Le_NumIntDoc = gErr

    Select Case gErr
    
        Case 192364
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 192365, 192366
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVAPORTEPAGTODIRETO", gErr)
            
        Case ERRO_LEITURA_SEM_DADOS

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192367)

    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function TRVVoucherInfoSigav_Le(ByVal objTRVVoucherInfo As ClassTRVVoucherInfo, Optional ByVal lComandoP As Long = 0) As Long

Dim lErro As Long
Dim lComando As Long
Dim tTRVVoucherInfo As typeTRVVoucherInfo

On Error GoTo Erro_TRVVoucherInfoSigav_Le

    If lComandoP <> 0 Then
        lComando = lComandoP
    Else
        'Executa a abertura do Comando
        lComando = Comando_Abrir()
        If lComando = 0 Then gError 192595
    End If

    'Alocação de espaço no buffer
    tTRVVoucherInfo.sTipo = String(STRING_TRV_OCR_TIPOVOU, 0)
    tTRVVoucherInfo.sSerie = String(STRING_TRV_OCR_SERIE, 0)
    tTRVVoucherInfo.sProduto = String(STRING_PRODUTO, 0)
    tTRVVoucherInfo.sDestino = String(STRING_TRV_TAMANHO_OUTROS, 0)
    tTRVVoucherInfo.sVigencia = String(STRING_TRV_TAMANHO_OUTROS, 0)
    tTRVVoucherInfo.sIdioma = String(STRING_TRV_TAMANHO_OUTROS, 0)
    tTRVVoucherInfo.sDestinoVou = String(STRING_TRV_TAMANHO_OUTROS, 0)
    tTRVVoucherInfo.sControle = String(STRING_TRV_TAMANHO_OUTROS, 0)
    tTRVVoucherInfo.sConvenio = String(STRING_TRV_TAMANHO_OUTROS, 0)
    tTRVVoucherInfo.sSexo = String(STRING_TRV_TAMANHO_OUTROS, 0)
    tTRVVoucherInfo.sTipoDoc = String(STRING_TRV_TAMANHO_OUTROS, 0)
    tTRVVoucherInfo.sCartaoFid = String(STRING_TRV_TAMANHO_OUTROS, 0)
    tTRVVoucherInfo.sMoeda = String(STRING_TRV_TAMANHO_OUTROS, 0)
    tTRVVoucherInfo.sValor = String(STRING_TRV_TAMANHO_OUTROS, 0)
    tTRVVoucherInfo.sCia = String(STRING_TRV_TAMANHO_OUTROS, 0)
    tTRVVoucherInfo.sValidade = String(STRING_TRV_TAMANHO_OUTROS, 0)
    tTRVVoucherInfo.sNumeroCC = String(STRING_TRV_TAMANHO_OUTROS, 0)
    tTRVVoucherInfo.sTitular = String(STRING_TRV_TAMANHO_OUTROS, 0)
    tTRVVoucherInfo.sPassageiroSobreNome = String(STRING_TRV_TAMANHO_OUTROS, 0)
    tTRVVoucherInfo.sPassageiroNome = String(STRING_TRV_TAMANHO_OUTROS, 0)
    tTRVVoucherInfo.sPassageiroCGC = String(STRING_TRV_TAMANHO_OUTROS, 0)
    tTRVVoucherInfo.sPassageiroEndereco = String(STRING_TRV_TAMANHO_OUTROS, 0)
    tTRVVoucherInfo.sPassageiroBairro = String(STRING_TRV_TAMANHO_OUTROS, 0)
    tTRVVoucherInfo.sPassageiroCidade = String(STRING_TRV_TAMANHO_OUTROS, 0)
    tTRVVoucherInfo.sPassageiroCEP = String(STRING_TRV_TAMANHO_OUTROS, 0)
    tTRVVoucherInfo.sPassageiroUF = String(STRING_TRV_TAMANHO_OUTROS, 0)
    tTRVVoucherInfo.sPassageiroEmail = String(STRING_TRV_TAMANHO_OUTROS, 0)
    tTRVVoucherInfo.sPassageiroContato = String(STRING_TRV_TAMANHO_OUTROS, 0)
    tTRVVoucherInfo.sPassageiroTelefone1 = String(STRING_TRV_TAMANHO_OUTROS, 0)
    tTRVVoucherInfo.sPassageiroTelefone2 = String(STRING_TRV_TAMANHO_OUTROS, 0)
    tTRVVoucherInfo.sGrupo = String(STRING_TRV_TAMANHO_OUTROS, 0)
    tTRVVoucherInfo.sTitularCPF = String(STRING_TRV_TAMANHO_OUTROS, 0)
    tTRVVoucherInfo.sAprovacao = String(STRING_TRV_VOU_NUMAUTO, 0)

    'Le a tabelaTRVVoucherInfoSigav
    lErro = Comando_Executar(lComando, "SELECT Tipo, Serie, NumVou, DataEmissao, DataRecepcao, Cliente, " & _
                "FornEmissor, Produto, Destino, DataInicio, DataTermino, Vigencia, Idioma, " & _
                "Pax, DestinoVou, Antc, Controle, Convenio, DataPag, Cartao, " & _
                "Pago, NumFat, CliPassageiro, DataNasc, Sexo, TipoDoc, CartaoFid, " & _
                "Moeda, TarifaUnitaria, Cambio, Valor, TarifaPerc, TarifaValorMoeda, TarifaValorReal, " & _
                "ComissaoPerc, ComissaoValorMoeda, ComissaoValorReal, CartaoPerc, CartaoValorMoeda, CartaoValorReal, OverPerc, " & _
                "OverValorMoeda, OverValorReal, CMRPerc, CMRValorMoeda, CMRValorReal, Cia, Validade, " & _
                "NumeroCC, Titular, ValorCartao, Parcela, Aprovacao, PassageiroSobreNome, PassageiroNome, " & _
                "PassageiroCGC, PassageiroEndereco, PassageiroBairro, PassageiroCidade, PassageiroCEP, PassageiroUF, PassageiroEmail, " & _
                "PassageiroContato, PassageiroTelefone1,PassageiroTelefone2, Grupo, TitularCPF FROM TRVVoucherInfoSigav WHERE Tipo= ?  AND Serie= ?  AND NumVou= ? ", _
                tTRVVoucherInfo.sTipo, tTRVVoucherInfo.sSerie, tTRVVoucherInfo.lNumVou, tTRVVoucherInfo.dtDataEmissao, _
                tTRVVoucherInfo.dtDataRecepcao, tTRVVoucherInfo.lCliente, tTRVVoucherInfo.lFornEmissor, tTRVVoucherInfo.sProduto, tTRVVoucherInfo.sDestino, _
                tTRVVoucherInfo.dtDataInicio, tTRVVoucherInfo.dtDataTermino, tTRVVoucherInfo.sVigencia, tTRVVoucherInfo.sIdioma, tTRVVoucherInfo.iPax, _
                tTRVVoucherInfo.sDestinoVou, tTRVVoucherInfo.iAntc, tTRVVoucherInfo.sControle, tTRVVoucherInfo.sConvenio, tTRVVoucherInfo.dtDataPag, _
                tTRVVoucherInfo.iCartao, tTRVVoucherInfo.iPago, tTRVVoucherInfo.lNumFat, tTRVVoucherInfo.lCliPassageiro, tTRVVoucherInfo.dtDataNasc, _
                tTRVVoucherInfo.sSexo, tTRVVoucherInfo.sTipoDoc, tTRVVoucherInfo.sCartaoFid, tTRVVoucherInfo.sMoeda, tTRVVoucherInfo.dTarifaUnitaria, _
                tTRVVoucherInfo.dCambio, tTRVVoucherInfo.sValor, tTRVVoucherInfo.dTarifaPerc, tTRVVoucherInfo.dTarifaValorMoeda, tTRVVoucherInfo.dTarifaValorReal, _
                tTRVVoucherInfo.dComissaoPerc, tTRVVoucherInfo.dComissaoValorMoeda, tTRVVoucherInfo.dComissaoValorReal, tTRVVoucherInfo.dCartaoPerc, tTRVVoucherInfo.dCartaoValorMoeda, _
                tTRVVoucherInfo.dCartaoValorReal, tTRVVoucherInfo.dOverPerc, tTRVVoucherInfo.dOverValorMoeda, tTRVVoucherInfo.dOverValorReal, tTRVVoucherInfo.dCMRPerc, _
                tTRVVoucherInfo.dCMRValorMoeda, tTRVVoucherInfo.dCMRValorReal, tTRVVoucherInfo.sCia, tTRVVoucherInfo.sValidade, tTRVVoucherInfo.sNumeroCC, _
                tTRVVoucherInfo.sTitular, tTRVVoucherInfo.dValorCartao, tTRVVoucherInfo.lParcela, tTRVVoucherInfo.sAprovacao, tTRVVoucherInfo.sPassageiroSobreNome, _
                tTRVVoucherInfo.sPassageiroNome, tTRVVoucherInfo.sPassageiroCGC, tTRVVoucherInfo.sPassageiroEndereco, tTRVVoucherInfo.sPassageiroBairro, tTRVVoucherInfo.sPassageiroCidade, _
                tTRVVoucherInfo.sPassageiroCEP, tTRVVoucherInfo.sPassageiroUF, tTRVVoucherInfo.sPassageiroEmail, tTRVVoucherInfo.sPassageiroContato, tTRVVoucherInfo.sPassageiroTelefone1, tTRVVoucherInfo.sPassageiroTelefone2, tTRVVoucherInfo.sGrupo, tTRVVoucherInfo.sTitularCPF, _
                objTRVVoucherInfo.sTipo, objTRVVoucherInfo.sSerie, objTRVVoucherInfo.lNumVou)
    If lErro <> AD_SQL_SUCESSO Then gError 192596

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192597

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    Call TRVVoucherInfo_Type_Transfere_Obj(objTRVVoucherInfo, tTRVVoucherInfo)

    'Fecha Comando
    If lComandoP = 0 Then Call Comando_Fechar(lComando)

    TRVVoucherInfoSigav_Le = SUCESSO

    Exit Function

Erro_TRVVoucherInfoSigav_Le:

    TRVVoucherInfoSigav_Le = gErr

    Select Case gErr

        Case 192595
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 192596, 192597
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVVOUCHERINFOSIGAV", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192598)

    End Select

    'Fecha Comando
    If lComandoP = 0 Then Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function TRVVouchers_Le_Periodo(ByVal dtDataDe As Date, ByVal dtDataAte As Date, ByVal colVouchers As Collection, Optional ByVal bTodosSemInfo As Boolean = False, Optional ByVal iSistemaIntegradoRotina As Integer = SISTEMA_INTEGRADO_SIGAV) As Long

Dim lErro As Long
Dim lComando As Long
Dim tTRVVouchers As typeTRVVouchers
Dim objTRVVouchers As ClassTRVVouchers

On Error GoTo Erro_TRVVouchers_Le_Periodo

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 192699

    'Alocação de espaço no buffer
    tTRVVouchers.sTipoDoc = String(STRING_TRV_OCR_TIPODOC, 0)
    tTRVVouchers.sSerie = String(STRING_TRV_OCR_SERIE, 0)
    tTRVVouchers.sTipVou = String(STRING_TRV_OCR_TIPOVOU, 0)
    tTRVVouchers.sCodGrupo = String(STRING_TRV_OCR_CODGRUPO, 0)
    tTRVVouchers.sControle = String(STRING_TRV_VOU_CONTROLE, 0)
    tTRVVouchers.sMoeda = String(STRING_TRV_VOU_MOEADA, 0)
    tTRVVouchers.sProduto = String(STRING_TRV_VOU_PRODUTO, 0)
    tTRVVouchers.sTitular = String(STRING_TRV_VOU_TITULAR, 0)
    tTRVVouchers.sCiaCart = String(STRING_TRV_VOU_CIACART, 0)
    tTRVVouchers.sNumCCred = String(STRING_TRV_VOU_NUMCCRED, 0)
    tTRVVouchers.sNumAuto = String(STRING_TRV_VOU_NUMAUTO, 0)
    
    'Le a tabelaTRVVouchers
    If Not bTodosSemInfo Then
        lErro = Comando_Executar(lComando, "SELECT NumVou, TipoDoc, Serie, NumIntArq, NumRegArqTexto, Cliente, " & _
                    "Recnsica, Valor, Data, CondPagto, TipVou, CodGrupo, NumFatCoinfo, " & _
                    "TipoDocDestino, NumIntDocDestino, Cartao, Status, NumIntDocNVL, Titular, Controle, Moeda, Produto, Pax, Cambio, ValorCambio, CiaCart,NumCCred,NumAuto,QuantParc,DiasAntc,Kit FROM TRVVouchers WHERE Data BETWEEN ? AND ? ", _
                    tTRVVouchers.lNumVou, tTRVVouchers.sTipoDoc, tTRVVouchers.sSerie, tTRVVouchers.lNumIntArq, _
                    tTRVVouchers.lNumRegArqTexto, tTRVVouchers.lCliente, tTRVVouchers.lrecnsica, tTRVVouchers.dValor, tTRVVouchers.dtData, _
                    tTRVVouchers.iCondPagto, tTRVVouchers.sTipVou, tTRVVouchers.sCodGrupo, tTRVVouchers.lNumFatCoinfo, tTRVVouchers.iTipoDocDestino, _
                    tTRVVouchers.lNumIntDocDestino, tTRVVouchers.iCartao, tTRVVouchers.iStatus, tTRVVouchers.lNumIntDocNVL, tTRVVouchers.sTitular, _
                    tTRVVouchers.sControle, tTRVVouchers.sMoeda, tTRVVouchers.sProduto, tTRVVouchers.iPax, tTRVVouchers.dCambio, tTRVVouchers.dValorCambio, tTRVVouchers.sCiaCart, tTRVVouchers.sNumCCred, tTRVVouchers.sNumAuto, tTRVVouchers.iQuantParc, tTRVVouchers.idiasantc, tTRVVouchers.iKit, dtDataDe, dtDataAte)
    Else
        lErro = Comando_Executar(lComando, "SELECT NumVou, TipoDoc, Serie, NumIntArq, NumRegArqTexto, Cliente, " & _
                    "Recnsica, Valor, Data, CondPagto, TipVou, CodGrupo, NumFatCoinfo, " & _
                    "TipoDocDestino, NumIntDocDestino, Cartao, Status, NumIntDocNVL, Titular, Controle, Moeda, Produto, Pax, Cambio, ValorCambio, CiaCart,NumCCred,NumAuto,QuantParc,DiasAntc,Kit FROM TRVVouchers WHERE ((NOT EXISTS (SELECT I.NumVou FROM TRVVoucherInfoSigav AS I WHERE TRVVouchers.NumVou = I.NumVou AND  TRVVouchers.Serie = I.Serie AND  TRVVouchers.TipVou = I.Tipo )) OR ExtraiInfoSigav = ?) AND TipoDoc = ? AND (NumFatCoinfo <> 0 OR Status <> 7 OR Sistema = 1) AND Sistema = ? ORDER BY Data, NumVou DESC ", _
                    tTRVVouchers.lNumVou, tTRVVouchers.sTipoDoc, tTRVVouchers.sSerie, tTRVVouchers.lNumIntArq, _
                    tTRVVouchers.lNumRegArqTexto, tTRVVouchers.lCliente, tTRVVouchers.lrecnsica, tTRVVouchers.dValor, tTRVVouchers.dtData, _
                    tTRVVouchers.iCondPagto, tTRVVouchers.sTipVou, tTRVVouchers.sCodGrupo, tTRVVouchers.lNumFatCoinfo, tTRVVouchers.iTipoDocDestino, _
                    tTRVVouchers.lNumIntDocDestino, tTRVVouchers.iCartao, tTRVVouchers.iStatus, tTRVVouchers.lNumIntDocNVL, tTRVVouchers.sTitular, _
                    tTRVVouchers.sControle, tTRVVouchers.sMoeda, tTRVVouchers.sProduto, tTRVVouchers.iPax, tTRVVouchers.dCambio, tTRVVouchers.dValorCambio, tTRVVouchers.sCiaCart, tTRVVouchers.sNumCCred, tTRVVouchers.sNumAuto, tTRVVouchers.iQuantParc, tTRVVouchers.idiasantc, tTRVVouchers.iKit, MARCADO, TRV_TIPODOC_VOU_TEXTO, iSistemaIntegradoRotina)
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 192700

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192701

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objTRVVouchers = New ClassTRVVouchers

        objTRVVouchers.lNumVou = tTRVVouchers.lNumVou
        objTRVVouchers.sTipoDoc = tTRVVouchers.sTipoDoc
        objTRVVouchers.sSerie = tTRVVouchers.sSerie
        objTRVVouchers.lNumIntArq = tTRVVouchers.lNumIntArq
        objTRVVouchers.lNumRegArqTexto = tTRVVouchers.lNumRegArqTexto
        objTRVVouchers.lCliente = tTRVVouchers.lCliente
        objTRVVouchers.lrecnsica = tTRVVouchers.lrecnsica
        objTRVVouchers.dValor = tTRVVouchers.dValor
        objTRVVouchers.dtData = tTRVVouchers.dtData
        objTRVVouchers.iCondPagto = tTRVVouchers.iCondPagto
        objTRVVouchers.sTipVou = tTRVVouchers.sTipVou
        objTRVVouchers.sCodGrupo = tTRVVouchers.sCodGrupo
        objTRVVouchers.lNumFatCoinfo = tTRVVouchers.lNumFatCoinfo
        objTRVVouchers.iTipoDocDestino = tTRVVouchers.iTipoDocDestino
        objTRVVouchers.lNumIntDocDestino = tTRVVouchers.lNumIntDocDestino
        objTRVVouchers.iCartao = tTRVVouchers.iCartao
        objTRVVouchers.iStatus = tTRVVouchers.iStatus
        objTRVVouchers.lNumIntDocNVL = tTRVVouchers.lNumIntDocNVL
        objTRVVouchers.sTitular = tTRVVouchers.sTitular
        objTRVVouchers.sMoeda = tTRVVouchers.sMoeda
        objTRVVouchers.sProduto = tTRVVouchers.sProduto
        objTRVVouchers.sControle = tTRVVouchers.sControle
        objTRVVouchers.iPax = tTRVVouchers.iPax
        objTRVVouchers.dValorCambio = tTRVVouchers.dValorCambio
        objTRVVouchers.dCambio = tTRVVouchers.dCambio
        objTRVVouchers.sCiaCart = tTRVVouchers.sCiaCart
        objTRVVouchers.sNumCCred = tTRVVouchers.sNumCCred
        objTRVVouchers.sNumAuto = tTRVVouchers.sNumAuto
        objTRVVouchers.iQuantParc = tTRVVouchers.iQuantParc
        objTRVVouchers.idiasantc = tTRVVouchers.idiasantc
        objTRVVouchers.iKit = tTRVVouchers.iKit
        
        colVouchers.Add objTRVVouchers
        
        'Busca Proximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192702
        
    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    TRVVouchers_Le_Periodo = SUCESSO

    Exit Function

Erro_TRVVouchers_Le_Periodo:

    TRVVouchers_Le_Periodo = gErr

    Select Case gErr

        Case 192699
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 192700, 192701, 192702
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVVOUCHERS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192703)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Private Function TRV_Cliente_Tipo_Faturamento(ByVal lCliente As Long, bFaturarIndividualmente As Boolean) As Long

Dim lErro As Long
Dim objFilialCliente As New ClassFilialCliente
Dim colCategoriaItem As New Collection
Dim iIndice As Integer

On Error GoTo Erro_TRV_Cliente_Tipo_Faturamento

    bFaturarIndividualmente = False

    'Lê as categoria deste cliente
    objFilialCliente.lCodCliente = lCliente
    objFilialCliente.iCodFilial = FILIAL_MATRIZ
    
    lErro = CF("FilialCliente_Le_Categorias", objFilialCliente, colCategoriaItem)
    If lErro <> SUCESSO Then gError 192726
    
    If colCategoriaItem.Count > 0 Then
        For iIndice = 1 To colCategoriaItem.Count
            If UCase(colCategoriaItem.Item(iIndice).sCategoria) = TRV_CATEGORIA_FATURAMENTO Then
                If UCase(colCategoriaItem.Item(iIndice).sItem) = TRV_CATEGORIA_FATURAMENTO_POR_VOUCHER Then
                    bFaturarIndividualmente = True
                End If
                Exit For
            End If
        Next
    End If
 
    TRV_Cliente_Tipo_Faturamento = SUCESSO

    Exit Function

Erro_TRV_Cliente_Tipo_Faturamento:

    TRV_Cliente_Tipo_Faturamento = gErr

    Select Case gErr
    
        Case 192726

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192727)

    End Select


    Exit Function

End Function

Private Function TRV_GeracaoNF_Obtem_Cliente(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal objGeracaoNF As ClassTRVGeracaoNF, ByVal objTitRec As ClassTituloReceber, lCliente As Long, iFilialCliente As Integer) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim bAoPassageiro As Boolean
Dim lNumVou As Long
Dim sTipo As String
Dim sSerie As String
Dim objVoucherInfo As New ClassTRVVoucherInfo
Dim objSenha As New ClassSenha
Dim objTitRecTRV As ClassTitulosRecTRV
Dim lCliPassageiro As Long

On Error GoTo Erro_TRV_GeracaoNF_Obtem_Cliente

    Set objTitRecTRV = objTitRec.objInfoUsu
    
    lCliente = objTitRec.lCliente
    iFilialCliente = objTitRec.iFilial
    
    If objTitRecTRV.iNFAoPassageiro = MARCADO Then
        bAoPassageiro = True
    Else
        bAoPassageiro = False
    End If
    
    'Se o cliente da NF deve ser o passageiro e o mesmo não está cadastrado, tem
    'que obter as informações dele e fazer o cadastro
    If bAoPassageiro And objTitRecTRV.lCliPassageiro = 0 Then
    
        'TEM QUE OBTER UM DOS VOUCHERS
        sTipo = String(STRING_TRV_OCR_TIPOVOU, 0)
        sSerie = String(STRING_TRV_OCR_SERIE, 0)
        
        'Le a tabelaTRVVouchers
        lErro = Comando_Executar(lComando1, "SELECT NumVou, TipVou, Serie, CliPassageiro FROM TRVVouchers WHERE TipoDocDestino = ? AND NumIntDocDestino = ? ", lNumVou, sTipo, sSerie, lCliPassageiro, TRV_TIPO_DOC_DESTINO_TITREC, objTitRec.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 192732
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192733
    
        If lErro <> SUCESSO Then gError 192734
        
        objVoucherInfo.lNumVou = lNumVou
        objVoucherInfo.sSerie = sSerie
        objVoucherInfo.sTipo = sTipo
        objVoucherInfo.lCliPassageiro = lCliPassageiro
        
        'VERIFICAR SE O PASSAGEIRO ESTÁ PREENCHIDO
        If objVoucherInfo.lCliPassageiro = 0 Then
            
'            If objGeracaoNF.iExtrairDados = MARCADO Then
'                lErro = ERRO_LEITURA_SEM_DADOS
'            Else
            lErro = CF("TRVVoucherInfoSigav_Le", objVoucherInfo, lComando2)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 192735
'            End If
'
'            If lErro = ERRO_LEITURA_SEM_DADOS Then
'
'                If Len(Trim(objGeracaoNF.sSenhaSigav)) = 0 Then
'
'                    Load SigavSenha
'                    lErro = SigavSenha.Trata_Parametros(objSenha)
'                    If lErro <> SUCESSO Then gError 192728
'                    SigavSenha.Show vbModal
'
'                    If Len(Trim(objSenha.sSenha)) = 0 Then gError 192729
'
'                    objGeracaoNF.sSenhaSigav = objSenha.sSenha
'
'                End If
'
'                lErro = Obter_Dados_Sigav(objVoucherInfo, objGeracaoNF.sSenhaSigav)
'                If lErro <> SUCESSO Then gError 192736
'
'                lErro = CF("TRVVoucherInfoSigav_Grava", objVoucherInfo)
'                If lErro <> SUCESSO Then gError 192737
'
'            End If
                            
            'CADASTRAR O PASSAGEIRO E ACERTA A INFORMAÇÃO EM VOUCHERINFO
            lErro = CF("TRV_VoucherInfo_Cadastra_Cliente", objVoucherInfo)
            If lErro <> SUCESSO Then gError 192738

        End If
        
        objTitRecTRV.lCliPassageiro = objVoucherInfo.lCliPassageiro
    
    End If
    
    If bAoPassageiro Then
        lCliente = objTitRecTRV.lCliPassageiro
        iFilialCliente = FILIAL_MATRIZ
    End If
 
    TRV_GeracaoNF_Obtem_Cliente = SUCESSO

    Exit Function

Erro_TRV_GeracaoNF_Obtem_Cliente:

    TRV_GeracaoNF_Obtem_Cliente = gErr

    Select Case gErr

        Case 192731, 192735 To 192738, 192728
        
        Case 192732, 192733
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVVOUCHERS", gErr)
            
        Case 192734
            Call Rotina_Erro(vbOKOnly, "ERRO_VOUCHER_NAO_CADASTRADO2", gErr, objTitRec.lNumTitulo)
       
        Case 192729
            Call Rotina_Erro(vbOKOnly, "ERRO_SENHA_NAO_PREENCHIDA", gErr)
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192739)

    End Select
    
    Exit Function

End Function

Public Function Coinfo_Obtem_FilialEmp_Cliente(ByVal lcodemp As Long, iFilialEmpCorporator As Integer, iFilialEmpCoinfo As Integer) As Long
'Obtém a FilialEmpresa com base no Cliente (lcodemp)

Dim lErro As Long
Dim alComando(0 To 0) As Long
Dim iIndice As Integer
Dim iFilialCorp As Integer
Dim iFilialCoinfo As Integer
Dim lEmpresa As Long

On Error GoTo Erro_Coinfo_Obtem_FilialEmp_Cliente

    'Abre o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 130803
    Next
    
    lEmpresa = lcodemp
    
    'Le o Grupo
    lErro = Comando_Executar(alComando(0), "SELECT FilialEmpresa, FilialCoinfo FROM FiliaisClientesTRV WHERE CodCliente = ? ", _
        iFilialCorp, iFilialCoinfo, lEmpresa)
'    lErro = Comando_Executar(alComando(0), "SELECT Filial FROM CoInfoEmpresaGrupo WHERE CodEmp = ? ", _
'        iFilialCoinfo, lEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 130805
    
    'Busca o primeiro
    lErro = Comando_BuscarProximo(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130806
    
    If lErro = SUCESSO Then
        iFilialEmpCorporator = iFilialCorp
        iFilialEmpCoinfo = iFilialCoinfo
    Else
        iFilialEmpCorporator = FILIAL_MATRIZ
        iFilialEmpCoinfo = FILIAL_MATRIZ
    End If
    
    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Coinfo_Obtem_FilialEmp_Cliente = SUCESSO
     
    Exit Function
    
Erro_Coinfo_Obtem_FilialEmp_Cliente:

    Coinfo_Obtem_FilialEmp_Cliente = gErr
     
    Select Case gErr

        Case 130803
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 130805, 130806
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COINFOARQSLIDOS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 130558)
     
    End Select
     
    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function TitulosRecTRV_Le(ByVal objTitulosRecTRV As ClassTitulosRecTRV) As Long

Dim lErro As Long
Dim lComando As Long
Dim tTitulosRecTRV As typeTitulosRecTRV

On Error GoTo Erro_TitulosRecTRV_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 192760

    'Alocação de espaço no buffer

    'Le a tabelaTitulosRecTRV
    lErro = Comando_Executar(lComando, "SELECT NumIntDocTitRec, ValorTarifa, NumIntDocNFPagComi, ValorDeducoes, ValorComissao, ValorBruto FROM TitulosRecTRV WHERE NumIntDocTitRec= ? ", _
                tTitulosRecTRV.lNumIntDocTitRec, tTitulosRecTRV.dValorTarifa, tTitulosRecTRV.lNumIntDocNFPagComi, tTitulosRecTRV.dValorDeducoes, _
                tTitulosRecTRV.dValorComissao, tTitulosRecTRV.dValorBruto, _
                objTitulosRecTRV.lNumIntDocTitRec)
    If lErro <> AD_SQL_SUCESSO Then gError 192761

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192762

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objTitulosRecTRV.lNumIntDocTitRec = tTitulosRecTRV.lNumIntDocTitRec
    objTitulosRecTRV.dValorTarifa = tTitulosRecTRV.dValorTarifa
    objTitulosRecTRV.lNumIntDocNFPagComi = tTitulosRecTRV.lNumIntDocNFPagComi
    objTitulosRecTRV.dValorDeducoes = tTitulosRecTRV.dValorDeducoes
    objTitulosRecTRV.dValorComissao = tTitulosRecTRV.dValorComissao
    objTitulosRecTRV.dValorBruto = tTitulosRecTRV.dValorBruto

    'Fecha Comando
    Call Comando_Fechar(lComando)

    TitulosRecTRV_Le = SUCESSO

    Exit Function

Erro_TitulosRecTRV_Le:

    TitulosRecTRV_Le = gErr

    Select Case gErr

        Case 192760
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 192761, 192762
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOSRECTRV", gErr)

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
Dim objEnderecoTRV As New ClassEnderecoTRV

On Error GoTo Erro_Endereco_Le_Customizado

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 196211

    'Alocação de espaço no buffer
    sEmail2 = String(STRING_EMAIL, 0)

    'Le a tabelaTRVOcorrencias
    lErro = Comando_Executar(lComando, "SELECT Email2 FROM Enderecos WHERE Codigo= ? ", sEmail2, objEndereco.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 196212

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196213

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS
    
    objEnderecoTRV.sEmail2 = sEmail2

    Set objEndereco.objInfoUsu = objEnderecoTRV

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
Dim tCliEmissoresTRV As typeCliEmissoresTRV
Dim tCliEmissoresExcTRV As typeCliEmissoresExcTRV
Dim objCliEmissoresTRV As ClassCliEmissoresTRV
Dim objCliEmissoresExcTRV As ClassCliEmissoresExcTRV
Dim objClienteTRV As New ClassClienteTRV
Dim alComando(1 To 4) As Long
Dim iIndice As Integer
Dim tFiliaisClientesTRV As typeFiliaisClientesTRV

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
        tCliEmissoresTRV.sCPF = String(STRING_CGC, 0)
        tCliEmissoresTRV.sNumCartao = String(255, 0)
    
        'Le a tabela TRVCliEmissores
        lErro = Comando_Executar(alComando(1), "SELECT E.NumIntDoc, E.Cliente, E.Seq, E.FornEmissor, E.PercComissao, E.Cargo, E.NumCartao, E.PercCI, F.CGC FROM TRVCliEmissores As E, FiliaisFornecedores As F WHERE E.FornEmissor = F.CodFornecedor AND E.Cliente= ? ", _
                    tCliEmissoresTRV.lNumIntDoc, tCliEmissoresTRV.lCliente, tCliEmissoresTRV.iSeq, tCliEmissoresTRV.lFornEmissor, tCliEmissoresTRV.dPercComissao, tCliEmissoresTRV.lCargo, tCliEmissoresTRV.sNumCartao, tCliEmissoresTRV.dPercCI, tCliEmissoresTRV.sCPF, _
                    objCliente.lCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 195775
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 195776
    
        Do While lErro = AD_SQL_SUCESSO
    
            Set objCliEmissoresTRV = New ClassCliEmissoresTRV
    
            objCliEmissoresTRV.lNumIntDoc = tCliEmissoresTRV.lNumIntDoc
            objCliEmissoresTRV.lCliente = tCliEmissoresTRV.lCliente
            objCliEmissoresTRV.iSeq = tCliEmissoresTRV.iSeq
            objCliEmissoresTRV.lFornEmissor = tCliEmissoresTRV.lFornEmissor
            objCliEmissoresTRV.dPercComissao = tCliEmissoresTRV.dPercComissao
            objCliEmissoresTRV.dPercCI = tCliEmissoresTRV.dPercCI
            objCliEmissoresTRV.sCPF = tCliEmissoresTRV.sCPF
            objCliEmissoresTRV.sNumCartao = tCliEmissoresTRV.sNumCartao
            objCliEmissoresTRV.lCargo = tCliEmissoresTRV.lCargo
            
            objClienteTRV.colCliEmissoresTRV.Add objCliEmissoresTRV
            
            tCliEmissoresExcTRV.sProduto = String(STRING_PRODUTO, 0)
            
            'Le a tabela CliEmissoresExcTRV
            lErro = Comando_Executar(alComando(2), "SELECT NumIntDoc, NumIntDocCliEmi, Seq, Produto, PercComissao FROM TRVCliEmissoresExc WHERE NumIntDocCliEmi = ? ", _
                        tCliEmissoresExcTRV.lNumIntDoc, tCliEmissoresExcTRV.lNumIntDocCliEmi, tCliEmissoresExcTRV.iSeq, tCliEmissoresExcTRV.sProduto, tCliEmissoresExcTRV.dPercComissao, _
                        objCliEmissoresTRV.lNumIntDoc)
            If lErro <> AD_SQL_SUCESSO Then gError 195777
        
            'Busca Primeiro
            lErro = Comando_BuscarPrimeiro(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 195778
            
            Do While lErro = AD_SQL_SUCESSO
            
                Set objCliEmissoresExcTRV = New ClassCliEmissoresExcTRV
        
                objCliEmissoresExcTRV.lNumIntDoc = tCliEmissoresExcTRV.lNumIntDoc
                objCliEmissoresExcTRV.lNumIntDocCliEmi = tCliEmissoresExcTRV.lNumIntDocCliEmi
                objCliEmissoresExcTRV.iSeq = tCliEmissoresExcTRV.iSeq
                objCliEmissoresExcTRV.sProduto = tCliEmissoresExcTRV.sProduto
                objCliEmissoresExcTRV.dPercComissao = tCliEmissoresExcTRV.dPercComissao
            
                objCliEmissoresTRV.colCliEmissoresExcTRV.Add objCliEmissoresExcTRV
                
                lErro = Comando_BuscarProximo(alComando(2))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 195779
            
            Loop
            
            lErro = Comando_BuscarProximo(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 195780
            
        Loop
        
    End If
        
    tFiliaisClientesTRV.sGrupo = String(STRING_TRV_OCR_CODGRUPO, 0)
    
    'Le a tabela FiliaisClientesTRV
    lErro = Comando_Executar(alComando(3), "SELECT Representante, ComissaoRep, Correntista, ComissaoCorr, ComissaoAg, Aporte, EmpresaPai, FilialNF, FilialFAT, FilialEmpresa, FilialCoinfo, Grupo, CondPagtoCC, PercFatorDevCMCC FROM FiliaisClientesTRV WHERE CodCliente= ? AND CodFilial = ?", _
                tFiliaisClientesTRV.lRepresentante, tFiliaisClientesTRV.dPercComiRep, tFiliaisClientesTRV.lCorrentista, tFiliaisClientesTRV.dPercComiCorr, tFiliaisClientesTRV.dPercComiAg, tFiliaisClientesTRV.iConsiderarAporte, _
                tFiliaisClientesTRV.lEmpresaPai, tFiliaisClientesTRV.iFilialNF, tFiliaisClientesTRV.iFilialFat, tFiliaisClientesTRV.iFilialEmpresa, tFiliaisClientesTRV.iFilialCoinfo, tFiliaisClientesTRV.sGrupo, tFiliaisClientesTRV.iCondPagtoCC, tFiliaisClientesTRV.dPercFatorDevCMCC, objCliente.lCodigo, FILIAL_MATRIZ)
    If lErro <> AD_SQL_SUCESSO Then gError 195861

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 195862
    
    objClienteTRV.lCodRepresentante = tFiliaisClientesTRV.lRepresentante
    objClienteTRV.dPercComiRep = tFiliaisClientesTRV.dPercComiRep
    objClienteTRV.lCodCorrentista = tFiliaisClientesTRV.lCorrentista
    objClienteTRV.dPercComiCorr = tFiliaisClientesTRV.dPercComiCorr
    objClienteTRV.dPercComiAg = tFiliaisClientesTRV.dPercComiAg
    objClienteTRV.iConsiderarAporte = tFiliaisClientesTRV.iConsiderarAporte
    objClienteTRV.lEmpresaPai = tFiliaisClientesTRV.lEmpresaPai
    objClienteTRV.iFilialNF = tFiliaisClientesTRV.iFilialNF
    objClienteTRV.iFilialFat = tFiliaisClientesTRV.iFilialFat
    objClienteTRV.iFilialEmpresa = tFiliaisClientesTRV.iFilialEmpresa
    objClienteTRV.iFilialCoinfo = tFiliaisClientesTRV.iFilialCoinfo
    objClienteTRV.sGrupo = tFiliaisClientesTRV.sGrupo
    objClienteTRV.iCondPagtoCC = tFiliaisClientesTRV.iCondPagtoCC
    objClienteTRV.dPercFatorDevCMCC = tFiliaisClientesTRV.dPercFatorDevCMCC
    
    Set objCliente.objInfoUsu = objClienteTRV
    
    If bLeTudo Then
        
        lErro = CF("TRVClienteCorProd_Le", objCliente)
        If lErro <> SUCESSO Then gError 196953
        
        lErro = CF("TRVClienteRepProd_Le", objCliente)
        If lErro <> SUCESSO Then gError 196954
        
        lErro = CF("TRVExcComissaoCli_Le", objCliente)
        If lErro <> SUCESSO Then gError 196954
    
        lErro = TRVCliDataCallCenter_Le(objCliente)
        If lErro <> SUCESSO Then gError 196954
    
    End If

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
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVCLIEMISSORES", gErr)

        Case 195777, 195778, 195779
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVCLIEMISSORESEXC", gErr)
        
        Case 195861, 195862
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISCLIENTESTRV", gErr)
            
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

Public Function TRVAcordo_Le(ByVal objTRVAcordo As ClassTRVAcordos) As Long

Dim lErro As Long
Dim tTRVAcordo As typeTRVAcordos
Dim tTRVAcordoComiss As typeTRVAcordoComissao
Dim tTRVAcordoDif As typeTRVAcordoTarifaDif
Dim objTRVAcordoComiss As ClassTRVAcordoComissao
Dim objTRVAcordoDif As ClassTRVAcordoTarifaDif
Dim alComando(1 To 3) As Long
Dim iIndice As Integer

On Error GoTo Erro_TRVAcordo_Le

    'Abre o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 197145
    Next

    tTRVAcordo.sContrato = String(STRING_TRV_ACORDO_CONTRATO, 0)
    tTRVAcordo.sObservacao = String(STRING_TRV_ACORDO_OBS, 0)
    tTRVAcordo.sDescricao = String(STRING_TRV_ACORDO_DESC, 0)

    'Le a tabela TRVAcordo
    lErro = Comando_Executar(alComando(1), "SELECT NumIntDoc, Contrato, Cliente, FilialCliente, ValidadeDe, ValidadeAte, " & _
                              "Observacao, Descricao FROM TRVAcordos WHERE Numero = ? ", _
                tTRVAcordo.lNumIntDoc, tTRVAcordo.sContrato, tTRVAcordo.lCliente, tTRVAcordo.iFilialCliente, tTRVAcordo.dtValidadeDe, _
                tTRVAcordo.dtValidadeAte, tTRVAcordo.sObservacao, tTRVAcordo.sDescricao, objTRVAcordo.lNumero)
    If lErro <> AD_SQL_SUCESSO Then gError 197146

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197147

    If lErro <> AD_SQL_SUCESSO Then gError 197148

    objTRVAcordo.lNumIntDoc = tTRVAcordo.lNumIntDoc
    objTRVAcordo.sContrato = tTRVAcordo.sContrato
    objTRVAcordo.lCliente = tTRVAcordo.lCliente
    objTRVAcordo.iFilialCliente = tTRVAcordo.iFilialCliente
    objTRVAcordo.dtValidadeDe = tTRVAcordo.dtValidadeDe
    objTRVAcordo.dtValidadeAte = tTRVAcordo.dtValidadeAte
    objTRVAcordo.sObservacao = tTRVAcordo.sObservacao
    objTRVAcordo.sDescricao = tTRVAcordo.sDescricao

    tTRVAcordoComiss.sProduto = String(STRING_PRODUTO, 0)

    lErro = Comando_Executar(alComando(2), "SELECT NumIntDoc, Seq, Produto, Destino, PercComissao FROM TRVAcordoComissao WHERE NumIntAcordo = ? ORDER BY Seq", _
                                           tTRVAcordoComiss.lNumIntDoc, tTRVAcordoComiss.iSeq, tTRVAcordoComiss.sProduto, tTRVAcordoComiss.iDestino, tTRVAcordoComiss.dPercComissao, objTRVAcordo.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 197149

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197150

    Do While lErro = AD_SQL_SUCESSO
    
        Set objTRVAcordoComiss = New ClassTRVAcordoComissao
    
        objTRVAcordoComiss.iSeq = tTRVAcordoComiss.iSeq
        objTRVAcordoComiss.sProduto = tTRVAcordoComiss.sProduto
        objTRVAcordoComiss.iDestino = tTRVAcordoComiss.iDestino
        objTRVAcordoComiss.dPercComissao = tTRVAcordoComiss.dPercComissao
    
        objTRVAcordo.colTRVAcordoComiss.Add objTRVAcordoComiss
    
        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197151
    
    
    Loop

    tTRVAcordoDif.sProduto = String(STRING_PRODUTO, 0)


    lErro = Comando_Executar(alComando(3), "SELECT NumIntDoc, Seq, Produto, Preco, Moeda, Destino, PercComissao FROM TRVAcordoTarifaDif WHERE NumIntAcordo = ? ORDER BY Seq", _
                tTRVAcordoDif.lNumIntDoc, tTRVAcordoDif.iSeq, tTRVAcordoDif.sProduto, tTRVAcordoDif.dPreco, tTRVAcordoDif.iMoeda, tTRVAcordoDif.iDestino, tTRVAcordoDif.dPercComissao, objTRVAcordo.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 197152

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197153

    Do While lErro = AD_SQL_SUCESSO
    
        Set objTRVAcordoDif = New ClassTRVAcordoTarifaDif
    
        objTRVAcordoDif.iSeq = tTRVAcordoDif.iSeq
        objTRVAcordoDif.sProduto = tTRVAcordoDif.sProduto
        objTRVAcordoDif.dPreco = tTRVAcordoDif.dPreco
        objTRVAcordoDif.iMoeda = tTRVAcordoDif.iMoeda
        objTRVAcordoDif.iDestino = tTRVAcordoDif.iDestino
        objTRVAcordoDif.dPercComissao = tTRVAcordoDif.dPercComissao
    
        objTRVAcordo.colTRVAcordoDif.Add objTRVAcordoDif
    
        lErro = Comando_BuscarProximo(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197154
    
    Loop

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TRVAcordo_Le = SUCESSO

    Exit Function

Erro_TRVAcordo_Le:

    TRVAcordo_Le = gErr

    Select Case gErr

        Case 197145
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 197146, 197147
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVACORDOS", gErr)

        Case 197148

        Case 197149, 197150, 197151
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVACORDOCOMISS", gErr)
        
        Case 197152, 197153, 197154
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVACORDODIF", gErr)
        
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
Public Function TRVVoucherInfo_Le(ByVal objTRVVoucher As ClassTRVVouchers, Optional ByVal lComandoP As Long = 0, Optional ByVal bLeInfoAuxiliar As Boolean = False) As Long

Dim lErro As Long
Dim lComando As Long
Dim lComandoAux As Long
Dim tTRVVoucherInfoN As typeTRVVoucherInfoN
Dim objTRVVoucherInfoN As ClassTRVVoucherInfoN
Dim lNC As Long, dtNCData As Date, dNCValor As Double, lNCForn As Long, sNCFornNome As String
Dim lNumIntDocNF As Long, lNF As Long, dtNFData As Date, dNFValor As Double, lNFForn As Long, sNFFornNome As String

On Error GoTo Erro_TRVVoucherInfo_Le

    'Executa a abertura do Comando
    If lComandoP <> 0 Then
        lComando = lComandoP
    Else
        lComando = Comando_Abrir()
        If lComando = 0 Then gError 197207
    End If
    
    If bLeInfoAuxiliar Then
        lComandoAux = Comando_Abrir()
        If lComandoAux = 0 Then gError 197207
    End If

    'Alocação de espaço no buffer
    tTRVVoucherInfoN.sTipoDoc = String(STRING_TRV_VOU_INFO_TIPODOC, 0)
    tTRVVoucherInfoN.sTipVou = String(STRING_TRV_VOU_INFO_TIPVOU, 0)
    tTRVVoucherInfoN.sSerie = String(STRING_TRV_VOU_INFO_SERIE, 0)
    tTRVVoucherInfoN.sHistorico = String(STRING_TRV_VOU_HISTORICO, 0)
    tTRVVoucherInfoN.sUsuario = String(STRING_USUARIO, 0)

    'Le a tabelaTRVVoucherInfo
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, TipoDoc, TipVou, Serie, NumVou, Seq, " & _
                "Data, TipoDocDestino, NumIntDocDestino, NumIntDocComiInt, Valor, Historico, NumTitulo, Status,TipoLiberacao , Manual, TipoCliForn, Estorno, NumIntDocLiberacao, CliForn, NumIntDocOCR,NumIntDocEstorno,DataRegistro,HoraRegistro,Usuario,Indireta FROM TRVVoucherInfo WHERE TipVou= ?  AND Serie= ?  AND NumVou= ? ORDER BY TipoDoc, Seq", _
                tTRVVoucherInfoN.lNumIntDoc, tTRVVoucherInfoN.sTipoDoc, tTRVVoucherInfoN.sTipVou, tTRVVoucherInfoN.sSerie, _
                tTRVVoucherInfoN.lNumVou, tTRVVoucherInfoN.iSeq, tTRVVoucherInfoN.dtData, tTRVVoucherInfoN.iTipoDocDestino, tTRVVoucherInfoN.lNumIntDocDestino, _
                tTRVVoucherInfoN.lNumIntDocComiInt, tTRVVoucherInfoN.dValor, tTRVVoucherInfoN.sHistorico, tTRVVoucherInfoN.lNumTitulo, tTRVVoucherInfoN.iStatus, _
                tTRVVoucherInfoN.iTipoLiberacao, tTRVVoucherInfoN.iManual, tTRVVoucherInfoN.iTipoCliForn, tTRVVoucherInfoN.iEstorno, tTRVVoucherInfoN.lNumIntDocLiberacao, _
                tTRVVoucherInfoN.lCliForn, tTRVVoucherInfoN.lNumIntDocOCR, tTRVVoucherInfoN.lNumIntDocEstorno, tTRVVoucherInfoN.dtDataRegistro, tTRVVoucherInfoN.dHoraRegistro, tTRVVoucherInfoN.sUsuario, tTRVVoucherInfoN.iIndireta, _
                objTRVVoucher.sTipVou, objTRVVoucher.sSerie, objTRVVoucher.lNumVou)
    If lErro <> AD_SQL_SUCESSO Then gError 197208

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197209

    Do While lErro = AD_SQL_SUCESSO

        Set objTRVVoucherInfoN = New ClassTRVVoucherInfoN

        objTRVVoucherInfoN.lNumIntDoc = tTRVVoucherInfoN.lNumIntDoc
        objTRVVoucherInfoN.sTipoDoc = tTRVVoucherInfoN.sTipoDoc
        objTRVVoucherInfoN.sTipVou = tTRVVoucherInfoN.sTipVou
        objTRVVoucherInfoN.sSerie = tTRVVoucherInfoN.sSerie
        objTRVVoucherInfoN.lNumVou = tTRVVoucherInfoN.lNumVou
        objTRVVoucherInfoN.iSeq = tTRVVoucherInfoN.iSeq
        objTRVVoucherInfoN.dtData = tTRVVoucherInfoN.dtData
        objTRVVoucherInfoN.iTipoDocDestino = tTRVVoucherInfoN.iTipoDocDestino
        objTRVVoucherInfoN.lNumIntDocDestino = tTRVVoucherInfoN.lNumIntDocDestino
        objTRVVoucherInfoN.lNumIntDocComiInt = tTRVVoucherInfoN.lNumIntDocComiInt
        objTRVVoucherInfoN.dValor = tTRVVoucherInfoN.dValor
        objTRVVoucherInfoN.sHistorico = tTRVVoucherInfoN.sHistorico
        objTRVVoucherInfoN.lNumTitulo = tTRVVoucherInfoN.lNumTitulo
        objTRVVoucherInfoN.iStatus = tTRVVoucherInfoN.iStatus
        objTRVVoucherInfoN.iTipoLiberacao = tTRVVoucherInfoN.iTipoLiberacao
        objTRVVoucherInfoN.iManual = tTRVVoucherInfoN.iManual
        objTRVVoucherInfoN.iTipoCliForn = tTRVVoucherInfoN.iTipoCliForn
        objTRVVoucherInfoN.iEstorno = tTRVVoucherInfoN.iEstorno
        objTRVVoucherInfoN.lNumIntDocLiberacao = tTRVVoucherInfoN.lNumIntDocLiberacao
        objTRVVoucherInfoN.lCliForn = tTRVVoucherInfoN.lCliForn
        objTRVVoucherInfoN.lNumIntDocOCR = tTRVVoucherInfoN.lNumIntDocOCR
        objTRVVoucherInfoN.lNumIntDocEstorno = tTRVVoucherInfoN.lNumIntDocEstorno
        
        objTRVVoucherInfoN.dtDataRegistro = tTRVVoucherInfoN.dtDataRegistro
        objTRVVoucherInfoN.dHoraRegistro = tTRVVoucherInfoN.dHoraRegistro
        objTRVVoucherInfoN.iIndireta = tTRVVoucherInfoN.iIndireta
        objTRVVoucherInfoN.sUsuario = tTRVVoucherInfoN.sUsuario
        objTRVVoucherInfoN.dtDataNC = DATA_NULA
        objTRVVoucherInfoN.dtDataNF = DATA_NULA

        
        If bLeInfoAuxiliar Then
        
            If objTRVVoucherInfoN.iTipoDocDestino = TRV_TIPO_DOC_DESTINO_NFSPAG And objTRVVoucherInfoN.lNumIntDocDestino <> 0 Then
            
                sNCFornNome = String(STRING_FORNECEDOR_NOME_REDUZIDO, 0)
                sNFFornNome = String(STRING_FORNECEDOR_NOME_REDUZIDO, 0)
            
                lErro = Comando_Executar(lComandoAux, "SELECT NC,NCData,NCValor,NCForn,NCFornNome,NumIntDocNF,NF,NFData,NFValor,NFForn,NFFornNome FROM TRVNCNF WHERE NumIntDocNC = ?", lNC, dtNCData, dNCValor, lNCForn, sNCFornNome, lNumIntDocNF, lNF, dtNFData, dNFValor, lNFForn, sNFFornNome, objTRVVoucherInfoN.lNumIntDocDestino)
                If lErro <> AD_SQL_SUCESSO Then gError 197208
            
                'Busca Primeiro
                lErro = Comando_BuscarPrimeiro(lComandoAux)
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197209
            
                If lErro = AD_SQL_SUCESSO Then
                
                    objTRVVoucherInfoN.lNumNF = lNF
                    objTRVVoucherInfoN.lNumIntDocNF = lNumIntDocNF
                    objTRVVoucherInfoN.dtDataNC = dtNCData
                    objTRVVoucherInfoN.dtDataNF = dtNFData
                    objTRVVoucherInfoN.dValorNC = dNCValor
                    objTRVVoucherInfoN.dValorNF = dNFValor
                    objTRVVoucherInfoN.lFornNC = lNCForn
                    objTRVVoucherInfoN.lFornNF = lNFForn
                    objTRVVoucherInfoN.sFornNC = sNCFornNome
                    objTRVVoucherInfoN.sFornNF = sNFFornNome
                
                End If
            
            End If
            
        End If

        objTRVVoucher.colTRVVoucherInfo.Add objTRVVoucherInfoN

        'Busca Proximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197210

    Loop

    'Fecha Comando
    If lComandoP = 0 Then Call Comando_Fechar(lComando)
    If bLeInfoAuxiliar Then Call Comando_Fechar(lComandoAux)
    
    TRVVoucherInfo_Le = SUCESSO

    Exit Function

Erro_TRVVoucherInfo_Le:

    TRVVoucherInfo_Le = gErr

    Select Case gErr

        Case 197207
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 197208 To 197210
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVVOUCHERINFO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197211)

    End Select

    'Fecha Comando
    If lComandoP = 0 Then Call Comando_Fechar(lComando)
    If bLeInfoAuxiliar Then Call Comando_Fechar(lComandoAux)

    Exit Function

End Function


Public Function TRVGerComiIntDet_Le_Promotor(ByVal objTRVVoucher As ClassTRVVouchers) As Long

Dim lErro As Long
Dim lComando As Long
Dim objTRVGerComiIntDet As ClassTRVGerComiIntDet
Dim tTRVGerComiIntDet As typeTRVGerComiIntDet

On Error GoTo Erro_TRVGerComiIntDet_Le_Promotor

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 197217

    'Alocação de espaço no buffer
    tTRVGerComiIntDet.sNomeReduzidoVendedor = String(STRING_VENDEDOR_NOME_REDUZIDO, 0)

    'Le a tabelaTRVVoucherInfo
    lErro = Comando_Executar(lComando, "SELECT DISTINCT Vendedores.NomeReduzido, TRVGerComiIntDet.ValorBase, TRVGerComiIntDet.ValorComissao, TRVGerComiInt.DataGeracao FROM TRVVoucherInfo, TRVVoucherInfoGer, TRVGerComiInt, TRVGerComiIntDet, VendedorTRV, Vendedores WHERE TRVVoucherInfo.TipVou= ?  AND TRVVoucherInfo.Serie= ?  AND TRVVoucherInfo.NumVou= ? AND TRVVoucherInfo.NumIntDoc = TRVGerComiIntDet.NumIntDoc " & _
                                        " AND TRVGerComiIntDet.NumIntDocGerComi = TRVGerComiInt.NumIntDoc AND TRVGerComiIntDet.Vendedor = VendedorTRV.Vendedor AND TRVGerComiIntDet.Vendedor = Vendedores.Codigo AND VendedorTRV.Cargo = ? ORDER BY TRVGerComiInt.DataGeracao", _
                tTRVGerComiIntDet.sNomeReduzidoVendedor, tTRVGerComiIntDet.dValorBase, tTRVGerComiIntDet.dValorComissao, tTRVGerComiIntDet.dtDataGeracao, objTRVVoucher.sTipVou, objTRVVoucher.sSerie, objTRVVoucher.lNumVou, VENDEDOR_CARGO_PROMOTOR)
    If lErro <> AD_SQL_SUCESSO Then gError 197218

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197219

    Do While lErro = AD_SQL_SUCESSO

        Set objTRVGerComiIntDet = New ClassTRVGerComiIntDet

        objTRVGerComiIntDet.sNomeReduzidoVendedor = tTRVGerComiIntDet.sNomeReduzidoVendedor
        objTRVGerComiIntDet.dValorBase = tTRVGerComiIntDet.dValorBase
        objTRVGerComiIntDet.dValorComissao = tTRVGerComiIntDet.dValorComissao
        objTRVGerComiIntDet.dtDataGeracao = tTRVGerComiIntDet.dtDataGeracao

        objTRVVoucher.colTRVGerComiIntDet.Add objTRVGerComiIntDet

        'Busca Proximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197220

    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    TRVGerComiIntDet_Le_Promotor = SUCESSO

    Exit Function

Erro_TRVGerComiIntDet_Le_Promotor:

    TRVGerComiIntDet_Le_Promotor = gErr

    Select Case gErr

        Case 197217
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 197218 To 197220
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVGERCOMIINTDET", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197221)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Private Function TRVAportePagtoFatCond_Le(ByVal lComando As Long, ByVal objTRVAporte As ClassTRVAportes, Optional ByVal bLock As Boolean = False) As Long

Dim lErro As Long
Dim tTRVAportePagtoFatCond As typeTRVAportePagtoFatCond
Dim objTRVAportePagtoFatCond As ClassTRVAportePagtoFatC

On Error GoTo Erro_TRVAportePagtoFatCond_Le

    'Le a tabelaTRVAportePagtoFatCond
    If bLock Then
        lErro = Comando_ExecutarPos(lComando, "SELECT NumIntDoc, NumIntDocAporte, Valor, ValidadeDe, ValidadeAte,  " & _
                    "Percentual FROM TRVAportePagtoFatCond WHERE NumIntDocAporte= ? ORDER BY Seq ", 0, _
                    tTRVAportePagtoFatCond.lNumIntDoc, tTRVAportePagtoFatCond.lNumIntDocAporte, tTRVAportePagtoFatCond.dValor, tTRVAportePagtoFatCond.dtValidadeDe, _
                    tTRVAportePagtoFatCond.dtValidadeAte, tTRVAportePagtoFatCond.dPercentual, _
                    objTRVAporte.lNumIntDoc)
    Else
        lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntDocAporte, Valor, ValidadeDe, ValidadeAte,  " & _
                    "Percentual FROM TRVAportePagtoFatCond WHERE NumIntDocAporte= ? ORDER BY Seq ", _
                    tTRVAportePagtoFatCond.lNumIntDoc, tTRVAportePagtoFatCond.lNumIntDocAporte, tTRVAportePagtoFatCond.dValor, tTRVAportePagtoFatCond.dtValidadeDe, _
                    tTRVAportePagtoFatCond.dtValidadeAte, tTRVAportePagtoFatCond.dPercentual, _
                    objTRVAporte.lNumIntDoc)
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
    
        Set objTRVAportePagtoFatCond = New ClassTRVAportePagtoFatC

        objTRVAportePagtoFatCond.lNumIntDoc = tTRVAportePagtoFatCond.lNumIntDoc
        objTRVAportePagtoFatCond.lNumIntDocAporte = tTRVAportePagtoFatCond.lNumIntDocAporte
        objTRVAportePagtoFatCond.dValor = tTRVAportePagtoFatCond.dValor
        objTRVAportePagtoFatCond.dtValidadeDe = tTRVAportePagtoFatCond.dtValidadeDe
        objTRVAportePagtoFatCond.dtValidadeAte = tTRVAportePagtoFatCond.dtValidadeAte
        objTRVAportePagtoFatCond.dPercentual = tTRVAportePagtoFatCond.dPercentual
        
        objTRVAporte.colPagtoSobreFaturaCond.Add objTRVAportePagtoFatCond

        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190346

    Loop

    TRVAportePagtoFatCond_Le = SUCESSO

    Exit Function

Erro_TRVAportePagtoFatCond_Le:

    TRVAportePagtoFatCond_Le = gErr

    Select Case gErr

        Case 190344, 190345, 190346
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVAPORTEPAGTOFATCOND", gErr)

        Case 190381
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TRVAPORTEPAGTOFATCOND", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190347)

    End Select

    Exit Function

End Function

Private Function TRVAportePagtoFatCond_Le_NumIntDoc(ByVal lComando As Long, ByVal objTRVAportePagtoFatCond As ClassTRVAportePagtoFatC) As Long

Dim lErro As Long
Dim tTRVAportePagtoFatCond As typeTRVAportePagtoFatCond

On Error GoTo Erro_TRVAportePagtoFatCond_Le_NumIntDoc

    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntDocAporte, Valor, ValidadeDe, ValidadeAte,  " & _
                    "Percentual FROM TRVAportePagtoFatCond WHERE NumIntDoc= ? ORDER BY Seq ", _
                    tTRVAportePagtoFatCond.lNumIntDoc, tTRVAportePagtoFatCond.lNumIntDocAporte, tTRVAportePagtoFatCond.dValor, tTRVAportePagtoFatCond.dtValidadeDe, _
                    tTRVAportePagtoFatCond.dtValidadeAte, tTRVAportePagtoFatCond.dPercentual, _
                    objTRVAportePagtoFatCond.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 190950

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190951

    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS
    
    objTRVAportePagtoFatCond.lNumIntDoc = tTRVAportePagtoFatCond.lNumIntDoc
    objTRVAportePagtoFatCond.lNumIntDocAporte = tTRVAportePagtoFatCond.lNumIntDocAporte
    objTRVAportePagtoFatCond.dValor = tTRVAportePagtoFatCond.dValor
    objTRVAportePagtoFatCond.dtValidadeDe = tTRVAportePagtoFatCond.dtValidadeDe
    objTRVAportePagtoFatCond.dtValidadeAte = tTRVAportePagtoFatCond.dtValidadeAte
    objTRVAportePagtoFatCond.dPercentual = tTRVAportePagtoFatCond.dPercentual

    TRVAportePagtoFatCond_Le_NumIntDoc = SUCESSO

    Exit Function

Erro_TRVAportePagtoFatCond_Le_NumIntDoc:

    TRVAportePagtoFatCond_Le_NumIntDoc = gErr

    Select Case gErr

        Case 190950, 190951
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVAPORTEPagtoFatCond", gErr)

        Case ERRO_LEITURA_SEM_DADOS
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190952)

    End Select

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Private Function TRVVoucherAporte_Le_Voucher(ByVal objVouAporte As ClassTRVVoucherAporte, ByVal colVoucherAporte As Collection) As Long

Dim lErro As Long
Dim lComando As Long
Dim tTRVVoucherAporte As typeTRVVoucherAporte
Dim objTRVVoucherAporte As ClassTRVVoucherAporte

On Error GoTo Erro_TRVVoucherAporte_Le_Voucher

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 196461

    'Alocação de espaço no buffer
    tTRVVoucherAporte.sTipVou = String(STRING_TRV_OCR_TIPOVOU, 0)
    tTRVVoucherAporte.sSerie = String(STRING_TRV_OCR_SERIE, 0)

    'Le a tabelaTRVVoucherAporte
    lErro = Comando_Executar(lComando, "SELECT TipVou, Serie, NumVou, TipoPagtoAporte, NumIntDocPagtoAporte, CodigoAporte, " & _
                "ParcelaAporte, ValorAporte, ValorAporteUSS, MoedaAporte FROM TRVVoucherAporte WHERE TipVou= ?  AND Serie= ?  AND NumVou= ? ", _
                tTRVVoucherAporte.sTipVou, tTRVVoucherAporte.sSerie, tTRVVoucherAporte.lNumVou, tTRVVoucherAporte.iTipoPagtoAporte, _
                tTRVVoucherAporte.lNumIntDocPagtoAporte, tTRVVoucherAporte.lCodigoAporte, tTRVVoucherAporte.iParcelaAporte, tTRVVoucherAporte.dValorAporte, tTRVVoucherAporte.dValorAporteUSS, tTRVVoucherAporte.iMoedaAporte, _
                objVouAporte.sTipVou, objVouAporte.sSerie, objVouAporte.lNumVou)
    If lErro <> AD_SQL_SUCESSO Then gError 196462

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196463

    'Sem Dados
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objTRVVoucherAporte = New ClassTRVVoucherAporte

        objTRVVoucherAporte.sTipVou = tTRVVoucherAporte.sTipVou
        objTRVVoucherAporte.sSerie = tTRVVoucherAporte.sSerie
        objTRVVoucherAporte.lNumVou = tTRVVoucherAporte.lNumVou
        objTRVVoucherAporte.iTipoPagtoAporte = tTRVVoucherAporte.iTipoPagtoAporte
        objTRVVoucherAporte.lNumIntDocPagtoAporte = tTRVVoucherAporte.lNumIntDocPagtoAporte
        objTRVVoucherAporte.lCodigoAporte = tTRVVoucherAporte.lCodigoAporte
        objTRVVoucherAporte.iParcelaAporte = tTRVVoucherAporte.iParcelaAporte
        objTRVVoucherAporte.dValorAporte = tTRVVoucherAporte.dValorAporte
        objTRVVoucherAporte.dValorAporteUSS = tTRVVoucherAporte.dValorAporteUSS
        objTRVVoucherAporte.iMoedaAporte = tTRVVoucherAporte.iMoedaAporte
        
        colVoucherAporte.Add objTRVVoucherAporte
        
        'Busca Próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196464
        
    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    TRVVoucherAporte_Le_Voucher = SUCESSO

    Exit Function

Erro_TRVVoucherAporte_Le_Voucher:

    TRVVoucherAporte_Le_Voucher = gErr

    Select Case gErr

        Case 196461
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 196462, 196463, 196464
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVVOUCHERAPORTE", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196465)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function TRVFaturas_Le(ByVal lFatura As Long, ByRef objDestino As Object, ByRef iTipoDocDestino As Integer, Optional ByVal bLeExcluidas As Boolean = False, Optional ByVal bLeTudo As Boolean = False) As Long

Dim lErro As Long
Dim alComando(0 To 7) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim objTitRec As New ClassTituloReceber
Dim objTitPag As New ClassTituloPagar
Dim objNFsPag As New ClassNFsPag
Dim iLeExcluidas As Integer
Dim objFaturaTRV As New ClassFaturaTRV
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
Dim sSerie As String
Dim sTipVou As String
Dim sTipoDoc As String

On Error GoTo Erro_TRVFaturas_Le

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
        
    'Le a tabela TitulosRecTRV
    If bLeExcluidas Then
        lErro = Comando_Executar(alComando(0), "SELECT NumIntDoc,TipoDocDestino, Status FROM TRVFaturas WHERE NumTitulo = ? ORDER BY SiglaDocumento DESC", _
                    lNumIntDoc, iTipoDocDestino, iStatus, lFatura)
    Else
        lErro = Comando_Executar(alComando(0), "SELECT NumIntDoc,TipoDocDestino, Status FROM TRVFaturas WHERE NumTitulo = ? AND Status <> 5 ORDER BY SiglaDocumento DESC", _
                    lNumIntDoc, iTipoDocDestino, iStatus, lFatura)
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 196494

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196495

    If lErro <> AD_SQL_SEM_DADOS Then
    
        objFaturaTRV.lNumFat = lFatura
        
        Select Case iTipoDocDestino
        
            Case TRV_TIPO_DOC_DESTINO_TITREC
            
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
                
                Set objTitRec.objInfoUsu = objFaturaTRV
                
                Set objDestino = objTitRec
                
                dValor = objTitRec.dValor
                
                If bLeTudo Then
                
                    lErro = Comando_Executar(alComando(4), "SELECT B.Data, B.DataRegistro, B.HoraRegistro FROM BaixasRec AS B, BaixasParcRec as BP, ParcelasRecTodas AS P, TitulosRecTodos AS T WHERE T.NumIntDoc = P.NumIntTitulo AND P.NumIntDoc = BP.NumIntParcela AND BP.NumIntBaixa = B.NumIntBaixa AND T.NumIntDoc = ? AND B.Status <> 5 AND BP.Status <> 5 ORDER BY Data", _
                                   dtDataPagto, dtDataRegPagto, dHoraRegPagto, objTitRec.lNumIntDoc)
                    If lErro <> AD_SQL_SUCESSO Then gError 196494
                
                    'Busca Primeiro
                    lErro = Comando_BuscarPrimeiro(alComando(4))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196495
                    
                    If lErro = AD_SQL_SUCESSO Then
                        objFaturaTRV.dHoraRegPagto = dHoraRegPagto
                        objFaturaTRV.dtDataRegPagto = dtDataRegPagto
                        objFaturaTRV.dtDataPagto = dtDataPagto
                    End If
                    
                End If
                
            Case TRV_TIPO_DOC_DESTINO_TITPAG
            
                objTitPag.lNumIntDoc = lNumIntDoc
                
                lErro = CF("TituloPagar_Le_Todos", objTitPag, bLeExcluidas)
                If lErro <> SUCESSO Then gError 196498
                
                Set objTitPag.objInfoUsu = objFaturaTRV
                
                Set objDestino = objTitPag
                
                dValor = objTitPag.dValorTotal
                
                If bLeTudo Then
                
                    lErro = Comando_Executar(alComando(5), "SELECT B.Data, B.DataRegistro FROM BaixasPag AS B, BaixasParcPag as BP, ParcelasPagTodas AS P, TitulosPagTodos AS T WHERE T.NumIntDoc = P.NumIntTitulo AND P.NumIntDoc = BP.NumIntParcela AND BP.NumIntBaixa = B.NumIntBaixa AND T.NumIntDoc = ? AND B.Status <> 5 AND BP.Status <> 5 ORDER BY Data", _
                                  dtDataPagto, dtDataRegPagto, objTitPag.lNumIntDoc)
                    If lErro <> AD_SQL_SUCESSO Then gError 196494
                
                    'Busca Primeiro
                    lErro = Comando_BuscarPrimeiro(alComando(5))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196495
                    
                    If lErro = AD_SQL_SUCESSO Then
                        objFaturaTRV.dtDataRegPagto = dtDataRegPagto
                        objFaturaTRV.dtDataPagto = dtDataPagto
                    End If
                    
                End If
                
            Case TRV_TIPO_DOC_DESTINO_NFSPAG
            
                 objNFsPag.lNumIntDoc = lNumIntDoc
                
                'Lê o Título
                lErro = CF("NFPag_Le_Todos2", objNFsPag, iLeExcluidas)
                If lErro <> SUCESSO And lErro <> 18257 Then gError 196499
                
                Set objNFsPag.objInfoUsu = objFaturaTRV
                
                Set objDestino = objNFsPag
                
                dValor = objNFsPag.dValorTotal
                
                If bLeTudo Then
                
                    lErro = Comando_Executar(alComando(6), "SELECT B.Data, B.DataRegistro FROM BaixasPag AS B, BaixasParcPag as BP, ParcelasPagTodas AS P, TitulosPagTodos AS T, NFsPag_Todas AS N WHERE T.NumIntDoc = P.NumIntTitulo AND P.NumIntDoc = BP.NumIntParcela AND BP.NumIntBaixa = B.NumIntBaixa AND N.NumIntTitPag = T.NumIntDoc AND N.NumIntDoc = ? AND B.Status <> 5 AND BP.Status <> 5 ORDER BY Data", _
                                  dtDataPagto, dtDataRegPagto, objNFsPag.lNumIntDoc)
                    If lErro <> AD_SQL_SUCESSO Then gError 196494
                
                    'Busca Primeiro
                    lErro = Comando_BuscarPrimeiro(alComando(6))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196495
                    
                    If lErro = AD_SQL_SUCESSO Then
                        objFaturaTRV.dtDataRegPagto = dtDataRegPagto
                        objFaturaTRV.dtDataPagto = dtDataPagto
                    End If
                    
                End If
               
        End Select

    End If
    
    dValorAporteCred = 0
    
    If bLeTudo Then
    
        sTipoDoc = String(STRING_TRVTITULOS_TIPODOC, 0)
    
        lErro = Comando_Executar(alComando(7), "SELECT TipoDoc FROM TRVTitulos WHERE TipoDocOrigem = ? AND NumIntDocOrigem = ? ", _
                     sTipoDoc, iTipoDocDestino, lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 196494
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(7))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196495
        
        objFaturaTRV.sTipoDoc = sTipoDoc
    
        sUsuario = String(STRING_USUARIO, 0)
    
        lErro = Comando_Executar(alComando(1), "SELECT Usuario, Data, Hora,ValorAporteCred, ValorAporte FROM TRVTitulosExp WHERE TipoDocOrigem = ? AND NumIntDocOrigem = ? AND Excluido = ?", _
                     sUsuario, dtData, dHora, dValorAporteCred, dValorAporte, iTipoDocDestino, lNumIntDoc, DESMARCADO)
        If lErro <> AD_SQL_SUCESSO Then gError 196494
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196495
    
        If lErro = AD_SQL_SUCESSO Then
        
            objFaturaTRV.sUsuarioGerFat = sUsuario
            objFaturaTRV.dHoraGerFat = dHora
            objFaturaTRV.dtDataGerFat = dtData

            dValorAporte = dValorAporte + dValorAporteCred
        
            objFaturaTRV.dValorAporte = dValorAporte
            objFaturaTRV.dPercAporte = Abs(dValorAporte) / (dValor + Abs(dValorAporte))

        End If
    
        If iStatus = STATUS_EXCLUIDO Then
        
            sUsuario = String(STRING_USUARIO, 0)
            sMotivoCanc = String(STRING_MOTIVOCANCEL, 0)
        
            lErro = Comando_Executar(alComando(1), "SELECT Usuario, Data, Hora,Motivo FROM TRVTitulosExp WHERE TipoDocOrigem = ? AND NumIntDocOrigem = ? AND Excluido = ?", _
                         sUsuario, dtData, dHora, sMotivoCanc, iTipoDocDestino, lNumIntDoc, MARCADO)
            If lErro <> AD_SQL_SUCESSO Then gError 196494
        
            'Busca Primeiro
            lErro = Comando_BuscarPrimeiro(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196495
        
            If lErro = AD_SQL_SUCESSO Then
                objFaturaTRV.sUsuarioCanc = sUsuario
                objFaturaTRV.dHoraCanc = dHora
                objFaturaTRV.dtDataCanc = dtData
                objFaturaTRV.sMotivo = sMotivoCanc
            End If
            
        End If
        
        sSerie = String(STRING_TRV_VOU_INFO_SERIE, 0)
        sTipVou = String(STRING_TRV_VOU_INFO_TIPVOU, 0)
        
        lErro = Comando_Executar(alComando(2), "SELECT NumVou, Serie, TipVou FROM DocFaturados WHERE TipoDocDestino = ? AND NumIntDocDestino = ?", _
                     lNumVou, sSerie, sTipVou, iTipoDocDestino, lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 196494
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196495
                
        lErro = Comando_Executar(alComando(3), "SELECT Cliente FROM TRVVoucherInfoSigav WHERE Tipo = ? AND Serie = ? AND NumVou = ?", _
                     lCliente, sTipVou, sSerie, lNumVou)
        If lErro <> AD_SQL_SUCESSO Then gError 196494
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196495
        
        If lErro = AD_SQL_SUCESSO Then
            objFaturaTRV.lClienteVou = lCliente
        End If
        
       
    End If

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TRVFaturas_Le = SUCESSO

    Exit Function

Erro_TRVFaturas_Le:

    TRVFaturas_Le = gErr

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

Public Function TitulosRecTRV_Le_NF(ByVal lNumIntDoc As Long, ByVal objNF As ClassNFiscal, iItem As Integer) As Long

Dim lErro As Long
Dim alComando(0 To 0) As Long
Dim iIndice As Integer
Dim lNumIntNF As Long

On Error GoTo Erro_TitulosRecTRV_Le_NF

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 196760
    Next

    'Le a tabelaTitulosRecTRV
    lErro = Comando_Executar(alComando(0), "SELECT N.NumIntDoc, I.Item FROM TRVItemNotaTitulo AS T, NFiscal AS N, ItensNFiscal AS I WHERE T.NumIntDocTitulo = ? AND T.NumIntDocItemNF = I.NumIntDoc AND I.NumIntNF = N.NumIntDoc AND N.Status NOT IN (5,7) ", _
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

    TitulosRecTRV_Le_NF = SUCESSO

    Exit Function

Erro_TitulosRecTRV_Le_NF:

    TitulosRecTRV_Le_NF = gErr

    Select Case gErr

        Case 196760
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 196761, 196762
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOSRECTRV", gErr)
            
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
Public Function TRVClienteRepProd_Le(ByVal objCliente As ClassCliente) As Long

Dim lErro As Long
Dim lComando As Long
Dim tTRVClienteRepProd As typeTRVClienteRepProd
Dim objTRVClienteRepProd As ClassTRVClienteRepProd
Dim objClienteTRV As ClassClienteTRV

On Error GoTo Erro_TRVClienteRepProd_Le

    If Not (objCliente.objInfoUsu Is Nothing) Then
    
        Set objClienteTRV = objCliente.objInfoUsu

        'Executa a abertura do Comando
        lComando = Comando_Abrir()
        If lComando = 0 Then gError 196920
    
        'Alocação de espaço no buffer
        tTRVClienteRepProd.sProduto = String(STRING_PRODUTO, 0)
    
        'Le a tabelaTRVClienteRepProd
        lErro = Comando_Executar(lComando, "SELECT Cliente, Seq, Representante, Produto, PercComis FROM TRVClienteRepProd WHERE Cliente= ? ORDER BY Seq ", _
                    tTRVClienteRepProd.lCliente, tTRVClienteRepProd.iSeq, tTRVClienteRepProd.lRepresentante, tTRVClienteRepProd.sProduto, _
                    tTRVClienteRepProd.dPercComis, _
                    objCliente.lCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 196921
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196922
    
        'Sem Dados
        Do While lErro <> AD_SQL_SEM_DADOS
        
            Set objTRVClienteRepProd = New ClassTRVClienteRepProd
    
            objTRVClienteRepProd.lCliente = tTRVClienteRepProd.lCliente
            objTRVClienteRepProd.iSeq = tTRVClienteRepProd.iSeq
            objTRVClienteRepProd.lRepresentante = tTRVClienteRepProd.lRepresentante
            objTRVClienteRepProd.sProduto = tTRVClienteRepProd.sProduto
            objTRVClienteRepProd.dPercComis = tTRVClienteRepProd.dPercComis
            
            objClienteTRV.colExcRep.Add objTRVClienteRepProd
            
            'Busca Próximo
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196923
            
        Loop
    
        'Fecha Comando
        Call Comando_Fechar(lComando)
        
    End If

    TRVClienteRepProd_Le = SUCESSO

    Exit Function

Erro_TRVClienteRepProd_Le:

    TRVClienteRepProd_Le = gErr

    Select Case gErr

        Case 196920
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 196921 To 196923
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVCLIENTEREPPROD", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196924)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function TRVClienteCorProd_Le(ByVal objCliente As ClassCliente) As Long

Dim lErro As Long
Dim lComando As Long
Dim tTRVClienteCorProd As typeTRVClienteCorProd
Dim objTRVClienteCorProd As ClassTRVClienteCorProd
Dim objClienteTRV As ClassClienteTRV

On Error GoTo Erro_TRVClienteCorProd_Le

    If Not (objCliente.objInfoUsu Is Nothing) Then
    
        Set objClienteTRV = objCliente.objInfoUsu

        'Executa a abertura do Comando
        lComando = Comando_Abrir()
        If lComando = 0 Then gError 196926
    
        'Alocação de espaço no buffer
        tTRVClienteCorProd.sProduto = String(STRING_PRODUTO, 0)
    
        'Le a tabelaTRVClienteCorProdse
        lErro = Comando_Executar(lComando, "SELECT Cliente, Seq, Correntista, Produto, PercComis FROM TRVClienteCorProd WHERE Cliente= ? ORDER BY Seq ", _
                    tTRVClienteCorProd.lCliente, tTRVClienteCorProd.iSeq, tTRVClienteCorProd.lCorrentista, tTRVClienteCorProd.sProduto, _
                    tTRVClienteCorProd.dPercComis, _
                    objCliente.lCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 196927
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196928
    
        'Sem Dados
        Do While lErro <> AD_SQL_SEM_DADOS
        
            Set objTRVClienteCorProd = New ClassTRVClienteCorProd
    
            objTRVClienteCorProd.lCliente = tTRVClienteCorProd.lCliente
            objTRVClienteCorProd.iSeq = tTRVClienteCorProd.iSeq
            objTRVClienteCorProd.lCorrentista = tTRVClienteCorProd.lCorrentista
            objTRVClienteCorProd.sProduto = tTRVClienteCorProd.sProduto
            objTRVClienteCorProd.dPercComis = tTRVClienteCorProd.dPercComis
            
            objClienteTRV.colExcCor.Add objTRVClienteCorProd
            
            'Busca Próximo
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196929
            
        Loop
    
        'Fecha Comando
        Call Comando_Fechar(lComando)
        
    End If

    TRVClienteCorProd_Le = SUCESSO

    Exit Function

Erro_TRVClienteCorProd_Le:

    TRVClienteCorProd_Le = gErr

    Select Case gErr

        Case 196926
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 196927 To 196929
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVCLIENTECorPROD", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196930)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function TRVExcComissaoCli_Le(ByVal objCliente As ClassCliente) As Long

Dim lErro As Long
Dim lComando As Long
Dim tTRVExcComisCli As typeTRVExcComissaoCli
Dim objTRVExcComisCli As ClassTRVExcComissaoCli
Dim objClienteTRV As ClassClienteTRV

On Error GoTo Erro_TRVExcComissaoCli_Le

    If Not (objCliente.objInfoUsu Is Nothing) Then
    
        Set objClienteTRV = objCliente.objInfoUsu

        'Executa a abertura do Comando
        lComando = Comando_Abrir()
        If lComando = 0 Then gError 196926
    
        'Alocação de espaço no buffer
        tTRVExcComisCli.sProduto = String(STRING_PRODUTO, 0)
    
        'Le a tabelaTRVExcComisClise
        lErro = Comando_Executar(lComando, "SELECT C.Cliente, C.Seq, C.Produto, C.PercComissao FROM TRVExcComissaoCli AS C, FiliaisClientesTRV AS F WHERE F.EmpresaPai = C.Cliente AND F.CodCliente= ? ORDER BY C.Seq ", _
                    tTRVExcComisCli.lCliente, tTRVExcComisCli.iSeq, tTRVExcComisCli.sProduto, tTRVExcComisCli.dPercComissao, _
                    objCliente.lCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 196927
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196928
    
        'Sem Dados
        Do While lErro <> AD_SQL_SEM_DADOS
        
            Set objTRVExcComisCli = New ClassTRVExcComissaoCli
    
            objTRVExcComisCli.lCliente = tTRVExcComisCli.lCliente
            objTRVExcComisCli.iSeq = tTRVExcComisCli.iSeq
            objTRVExcComisCli.sProduto = tTRVExcComisCli.sProduto
            objTRVExcComisCli.dPercComissao = tTRVExcComisCli.dPercComissao
            
            objClienteTRV.colExcComissaoCli.Add objTRVExcComisCli
            
            'Busca Próximo
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196929
            
        Loop
    
        'Fecha Comando
        Call Comando_Fechar(lComando)
        
    End If

    TRVExcComissaoCli_Le = SUCESSO

    Exit Function

Erro_TRVExcComissaoCli_Le:

    TRVExcComissaoCli_Le = gErr

    Select Case gErr

        Case 196926
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 196927 To 196929
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVEXCCOMISSAOCLI", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196930)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function TRVTiposOcorrencia_Le(ByVal objTRVTiposOcorrencia As ClassTRVTiposOcorrencia) As Long

Dim lErro As Long
Dim lComando As Long
Dim tTRVTiposOcorrencia As typeTRVTiposOcorrencia

On Error GoTo Erro_TRVTiposOcorrencia_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 198081

    'Alocação de espaço no buffer
    tTRVTiposOcorrencia.sDescricao = String(STRING_TRV_TIPOOCR_DESCRICAO, 0)

    'Le a tabelaTRVTiposOcorrencia
    lErro = Comando_Executar(lComando, "SELECT Codigo, Descricao, EstornaAporteVou, ConsideraComisInt, AlteraComiVou, AlteraCMCC, " & _
                "AlteraCMC, AlteraCMR, AlteraOVER, AlteraCMA, AceitaVlrPositivo, AceitaVlrNegativo, IncideSobre FROM TRVTiposOcorrencia WHERE Codigo= ? ", _
                tTRVTiposOcorrencia.iCodigo, tTRVTiposOcorrencia.sDescricao, tTRVTiposOcorrencia.iEstornaAporteVou, tTRVTiposOcorrencia.iConsideraComisInt, _
                tTRVTiposOcorrencia.iAlteraComiVou, tTRVTiposOcorrencia.iAlteraCMCC, tTRVTiposOcorrencia.iAlteraCMC, tTRVTiposOcorrencia.iAlteraCMR, tTRVTiposOcorrencia.iAlteraOVER, _
                tTRVTiposOcorrencia.iAlteraCMA, tTRVTiposOcorrencia.iAceitaVlrPositivo, tTRVTiposOcorrencia.iAceitaVlrNegativo, tTRVTiposOcorrencia.iIncideSobre, _
                objTRVTiposOcorrencia.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 198082

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 198083

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objTRVTiposOcorrencia.iCodigo = tTRVTiposOcorrencia.iCodigo
    objTRVTiposOcorrencia.sDescricao = tTRVTiposOcorrencia.sDescricao
    objTRVTiposOcorrencia.iEstornaAporteVou = tTRVTiposOcorrencia.iEstornaAporteVou
    objTRVTiposOcorrencia.iConsideraComisInt = tTRVTiposOcorrencia.iConsideraComisInt
    objTRVTiposOcorrencia.iAlteraComiVou = tTRVTiposOcorrencia.iAlteraComiVou
    objTRVTiposOcorrencia.iAlteraCMCC = tTRVTiposOcorrencia.iAlteraCMCC
    objTRVTiposOcorrencia.iAlteraCMC = tTRVTiposOcorrencia.iAlteraCMC
    objTRVTiposOcorrencia.iAlteraCMR = tTRVTiposOcorrencia.iAlteraCMR
    objTRVTiposOcorrencia.iAlteraOVER = tTRVTiposOcorrencia.iAlteraOVER
    objTRVTiposOcorrencia.iAlteraCMA = tTRVTiposOcorrencia.iAlteraCMA
    objTRVTiposOcorrencia.iAceitaVlrPositivo = tTRVTiposOcorrencia.iAceitaVlrPositivo
    objTRVTiposOcorrencia.iAceitaVlrNegativo = tTRVTiposOcorrencia.iAceitaVlrNegativo
    objTRVTiposOcorrencia.iIncideSobre = tTRVTiposOcorrencia.iIncideSobre

    'Fecha Comando
    Call Comando_Fechar(lComando)

    TRVTiposOcorrencia_Le = SUCESSO

    Exit Function

Erro_TRVTiposOcorrencia_Le:

    TRVTiposOcorrencia_Le = gErr

    Select Case gErr

        Case 198081
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 198082, 198083
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVTIPOSOCORRENCIA", gErr)

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
Public Function TRVCarrega_TipoOrigemOCR(ByVal objCombo As Object) As Long

Dim lErro As Long
Dim lComando As Long
Dim tTRVTiposOcorrencia As typeTRVTiposOcorrencia

On Error GoTo Erro_TRVCarrega_TipoOrigemOCR

    objCombo.Clear

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 198001

    'Alocação de espaço no buffer
    tTRVTiposOcorrencia.sDescricao = String(STRING_TRV_TIPOOCR_DESCRICAO, 0)

    'Le a tabelaTRVTiposOcorrencia
    lErro = Comando_Executar(lComando, "SELECT Codigo, Descricao FROM TRVTiposOcorrencia ORDER BY Codigo ", _
                tTRVTiposOcorrencia.iCodigo, tTRVTiposOcorrencia.sDescricao)
    If lErro <> AD_SQL_SUCESSO Then gError 198002

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 198003
    
    Do While lErro <> AD_SQL_SEM_DADOS
    
        objCombo.AddItem tTRVTiposOcorrencia.iCodigo & SEPARADOR & tTRVTiposOcorrencia.sDescricao
        objCombo.ItemData(objCombo.NewIndex) = tTRVTiposOcorrencia.iCodigo
    
        'Busca Próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 198004

    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    TRVCarrega_TipoOrigemOCR = SUCESSO

    Exit Function

Erro_TRVCarrega_TipoOrigemOCR:

    TRVCarrega_TipoOrigemOCR = gErr

    Select Case gErr

        Case 198001
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 198002, 198003, 198004
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVTIPOSOCORRENCIA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198005)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function TRVGerComiInt_Le(ByVal objTRVGerComiInt As ClassTRVGerComiInt) As Long

Dim lErro As Long
Dim lComando As Long
Dim tTRVGerComiInt As typeTRVGerComiInt

On Error GoTo Erro_TRVGerComiInt_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 197892

    'Alocação de espaço no buffer
    tTRVGerComiInt.sUsuario = String(STRING_USUARIO_CODIGO, 0)

    'Le a tabelaTRVVoucherInfo
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, DataGeracao, HoraGeracao, Usuario, DataPagtoDe, DataPagtoAte, DataEmiDe, DataEmiAte FROM TRVGerComiInt WHERE Codigo = ? ", _
                tTRVGerComiInt.lNumIntDoc, tTRVGerComiInt.dtDataGeracao, tTRVGerComiInt.dHoraGeracao, tTRVGerComiInt.sUsuario, tTRVGerComiInt.dtDataPagtoDe, tTRVGerComiInt.dtDataPagtoAte, tTRVGerComiInt.dtDataEmiDe, tTRVGerComiInt.dtDataEmiAte, objTRVGerComiInt.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 197893

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197894

    If lErro <> AD_SQL_SUCESSO Then gError 197895

    objTRVGerComiInt.lNumIntDoc = tTRVGerComiInt.lNumIntDoc
    objTRVGerComiInt.dtDataGeracao = tTRVGerComiInt.dtDataGeracao
    objTRVGerComiInt.dHoraGeracao = tTRVGerComiInt.dHoraGeracao
    objTRVGerComiInt.sUsuario = tTRVGerComiInt.sUsuario
    objTRVGerComiInt.dtDataPagtoDe = tTRVGerComiInt.dtDataPagtoDe
    objTRVGerComiInt.dtDataPagtoAte = tTRVGerComiInt.dtDataPagtoAte
    objTRVGerComiInt.dtDataEmiDe = tTRVGerComiInt.dtDataEmiDe
    objTRVGerComiInt.dtDataEmiAte = tTRVGerComiInt.dtDataEmiAte

    'Fecha Comando
    Call Comando_Fechar(lComando)

    TRVGerComiInt_Le = SUCESSO

    Exit Function

Erro_TRVGerComiInt_Le:

    TRVGerComiInt_Le = gErr

    Select Case gErr

        Case 197892
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 197893, 197894
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVGERCOMIINT", gErr)

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
Public Function TRVProdTarifa_Le(ByVal colTarifas As Collection) As Long

Dim lErro As Long
Dim lComando As Long
Dim tTRVProdTarifa As typeTRVProdTarifa
Dim objTRVProdTarifa As ClassTRVProdTarifa
Dim sProdAnt As String
Dim objTRVProdTarifaCalc As ClassTRVProdTafCalc

On Error GoTo Erro_TRVProdTarifa_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 200240

    'Alocação de espaço no buffer
    tTRVProdTarifa.sProduto = String(STRING_PRODUTO, 0)

    'Le a tabelaTRVProdTarifa
    lErro = Comando_Executar(lComando, "SELECT Ano, Produto, DiasDe, DiasAte, Diario, Valor, " & _
                "ValorAdicional FROM TRVProdTarifa ORDER BY Produto, Ano, DiasDe ", _
                tTRVProdTarifa.iAno, tTRVProdTarifa.sProduto, tTRVProdTarifa.iDiasDe, tTRVProdTarifa.iDiasAte, _
                tTRVProdTarifa.iDiario, tTRVProdTarifa.dValor, tTRVProdTarifa.dValorAdicional)
    If lErro <> AD_SQL_SUCESSO Then gError 200241

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200242

    'Sem Dados
    Do While lErro <> AD_SQL_SEM_DADOS
    
        If UCase(sProdAnt) <> UCase(tTRVProdTarifa.sProduto) Then
            sProdAnt = tTRVProdTarifa.sProduto
            Set objTRVProdTarifaCalc = New ClassTRVProdTafCalc
            objTRVProdTarifaCalc.sProduto = tTRVProdTarifa.sProduto
            colTarifas.Add objTRVProdTarifaCalc
        End If
    
        Set objTRVProdTarifa = New ClassTRVProdTarifa

        objTRVProdTarifa.iAno = tTRVProdTarifa.iAno
        objTRVProdTarifa.sProduto = tTRVProdTarifa.sProduto
        objTRVProdTarifa.iDiasDe = tTRVProdTarifa.iDiasDe
        objTRVProdTarifa.iDiasAte = tTRVProdTarifa.iDiasAte
        objTRVProdTarifa.iDiario = tTRVProdTarifa.iDiario
        objTRVProdTarifa.dValor = tTRVProdTarifa.dValor
        objTRVProdTarifa.dValorAdicional = tTRVProdTarifa.dValorAdicional
        
        objTRVProdTarifaCalc.colTarifas.Add objTRVProdTarifa

        'Busca Próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200243

    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    TRVProdTarifa_Le = SUCESSO

    Exit Function

Erro_TRVProdTarifa_Le:

    TRVProdTarifa_Le = gErr

    Select Case gErr

        Case 200240
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 200241, 200242, 200243
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVPRODTARIFA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200244)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function TRVEmissor_Valida(ByVal lEmissor As Long, ByVal lCliente As Long) As Long

Dim lErro As Long
Dim lComando As Long
Dim lNumIntDoc As Long

On Error GoTo Erro_TRVEmissor_Valida

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 200240

    If lEmissor < FATOR_SOMA_COD_EMISSOR Then gError 200377 'Emissor fora faixa
    
    'Le a tabelaTRVProdTarifa
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc FROM TRVCliEmissores WHERE Cliente = ? AND FornEmissor = ? ", _
                lNumIntDoc, lCliente, lEmissor)
    If lErro <> AD_SQL_SUCESSO Then gError 200378

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200379

    'Sem Dados
    If lErro <> AD_SQL_SUCESSO Then gError 200380
    
    'Fecha Comando
    Call Comando_Fechar(lComando)

    TRVEmissor_Valida = SUCESSO

    Exit Function

Erro_TRVEmissor_Valida:

    TRVEmissor_Valida = gErr

    Select Case gErr
    
        Case 200377
            Call Rotina_Erro(vbOKOnly, "ERRO_FAIXA_EMISSOR_ERRADA", gErr, FATOR_SOMA_COD_EMISSOR)

        Case 200378, 200379
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVCLIEMISSORES", gErr)
            
        Case 200380
            Call Rotina_Erro(vbOKOnly, "ERRO_TRVCLIEMISSORES_NAO_CADASTRADO", gErr, lEmissor, lCliente)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200381)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function TRVRelGerComiInt_Le(ByVal objTRVRelGerComiInt As ClassTRVRelGerComiInt) As Long

Dim lErro As Long
Dim lComando As Long
Dim tTRVRelGerComiInt As typeTRVRelGerComiInt

On Error GoTo Erro_TRVRelGerComiInt_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 199494

    'Alocação de espaço no buffer
    tTRVRelGerComiInt.sUsuario = String(STRING_USUARIO_CODIGO, 0)

    'Le a tabelaTRVVoucherInfo
    lErro = Comando_Executar(lComando, "SELECT NumIntRel, DataGeracao, HoraGeracao, Usuario, DataPagtoDe, DataPagtoAte, DataEmiDe, DataEmiAte FROM TRVRelGerComiInt WHERE NumIntRel = ?", _
                tTRVRelGerComiInt.lNumIntRel, tTRVRelGerComiInt.dtDataGeracao, tTRVRelGerComiInt.dHoraGeracao, tTRVRelGerComiInt.sUsuario, tTRVRelGerComiInt.dtDataPagtoDe, tTRVRelGerComiInt.dtDataPagtoAte, tTRVRelGerComiInt.dtDataEmiDe, tTRVRelGerComiInt.dtDataEmiAte, objTRVRelGerComiInt.lNumIntRel)
    If lErro <> AD_SQL_SUCESSO Then gError 199495

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 199496

    If lErro <> AD_SQL_SUCESSO Then gError 199497

    objTRVRelGerComiInt.dtDataGeracao = tTRVRelGerComiInt.dtDataGeracao
    objTRVRelGerComiInt.dHoraGeracao = tTRVRelGerComiInt.dHoraGeracao
    objTRVRelGerComiInt.sUsuario = tTRVRelGerComiInt.sUsuario
    objTRVRelGerComiInt.dtDataPagtoDe = tTRVRelGerComiInt.dtDataPagtoDe
    objTRVRelGerComiInt.dtDataPagtoAte = tTRVRelGerComiInt.dtDataPagtoAte
    objTRVRelGerComiInt.dtDataEmiDe = tTRVRelGerComiInt.dtDataEmiDe
    objTRVRelGerComiInt.dtDataEmiAte = tTRVRelGerComiInt.dtDataEmiAte
        
    'Fecha Comando
    Call Comando_Fechar(lComando)

    TRVRelGerComiInt_Le = SUCESSO

    Exit Function

Erro_TRVRelGerComiInt_Le:

    TRVRelGerComiInt_Le = gErr

    Select Case gErr

        Case 199494
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 199495, 199496
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVRELGERCOMIINT", gErr)

        Case 199497

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 199498)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function TRVVouVendedores_Le(colVendedores As Collection, objTRVVoucher As ClassTRVVouchers, ByVal lComando As Long) As Long

Dim lErro As Long
Dim iVendedor As Integer
Dim objVendedor As ClassVendedor
Dim sNomeReduzido As String
Dim lComando1 As Long

On Error GoTo Erro_TRVVouVendedores_Le

    'Executa a abertura do Comando
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 199849

    'Le a tabelaTRVVoucherInfo
    lErro = Comando_Executar(lComando, "SELECT Vendedor FROM TRVVouVendedores WHERE TipVou = ? AND Serie = ? AND NumVou = ?", _
                iVendedor, objTRVVoucher.sTipVou, objTRVVoucher.sSerie, objTRVVoucher.lNumVou)
    If lErro <> AD_SQL_SUCESSO Then gError 199515

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 199516

    Do While lErro = AD_SQL_SUCESSO

        Set objVendedor = New ClassVendedor

        objVendedor.iCodigo = iVendedor

        colVendedores.Add objVendedor

        'Busca Primeiro
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 199517

    Loop
        
    For Each objVendedor In colVendedores
        
        'Alocação de espaço no buffer
        sNomeReduzido = String(STRING_VENDEDOR_NOME_REDUZIDO, 0)

        lErro = Comando_Executar(lComando1, "SELECT NomeReduzido FROM Vendedores WHERE Codigo = ?", _
                    sNomeReduzido, objVendedor.iCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 199847

        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 199848
        
        objVendedor.sNomeReduzido = sNomeReduzido
        
    Next
        
    'Fecha Comando
    Call Comando_Fechar(lComando1)
        
    TRVVouVendedores_Le = SUCESSO

    Exit Function

Erro_TRVVouVendedores_Le:

    TRVVouVendedores_Le = gErr

    Select Case gErr

        Case 199515, 199516, 199517
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVVOUVENDEDORES", gErr)

        Case 199847, 199848
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_VENDEDORES", gErr)

        Case 199849
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 199518)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Public Function TRVFatCartao_Le(ByVal colAuto As Collection) As Long
'Atualiza os Títulos de Cartão de Crédito a partir dos vouchers pagos com cartão

Dim lErro As Long
Dim lTransacao As Long
Dim alComando(0 To 0) As Long
Dim iIndice As Integer
Dim sNumAuto As String
Dim sNumAutoAnt As String
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
Dim objVouAutoCC As ClassTRVVouAutoCC
Dim objVou As ClassTRVFATInfoVoucher
Dim sSQL As String
Dim iFilialEmpresa As Integer

On Error GoTo Erro_TRVFatCartao_Le
    
    'Abre o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 200668
    Next
    
    sSQL = "SELECT V.CiaCart, V.NumCCred, V.NumAuto,V.ValidadeCC, V.NumVou, V.Serie, V.TipVou, V.Data, V.ClienteVou, C.NomeReduzido, " & _
        "V.Valor, V.ValorComissao, V.PassageiroNome, V.PassageiroSobrenome, V.Titular, V.TitularCPF, " & _
        "V.Cliente , A.NomeReduzido, V.QuantParc, P.Taxa, T.FilialEmpresa " & _
        "FROM TRVVouchers AS V, Clientes AS C, AdmMeioPagto AS M, Redes AS R, Clientes AS A, AdmMeioPagtoCondPagto AS P, FiliaisClientesTRV AS T " & _
        "WHERE V.Cartao = 1 AND V.Status <> 7 AND V.NumFatCoinfo = 0 AND V.TipoDoc = 'VOU' AND V.ExtraiInfoSigav = 0 AND V.CiaCart <> '' AND V.NumAuto <> '' AND " & _
        "V.ClienteVou = C.Codigo AND V.CiaCart = M.Nome AND M.FilialEmpresa = 1 AND M.Rede = R.Codigo AND A.Codigo = V.Cliente AND " & _
        "P.AdmMeioPagto = M.Codigo And P.NumParcelas = V.QuantParc And P.FilialEmpresa = 1 And T.CodCliente = C.Codigo " & _
        "ORDER BY V.CiaCart, V.Cliente, V.NumAuto, V.NumVou"
    
    sCia = String(STRING_TRV_VOU_CIACART, 0)
    sNumCCred = String(STRING_TRV_VOU_NUMCCRED, 0)
    sSerie = String(STRING_TRV_OCR_SERIE, 0)
    sTipo = String(STRING_TRV_OCR_TIPOVOU, 0)
    sClienteVou = String(STRING_CLIENTE_NOME_REDUZIDO, 0)
    sNomePax = String(STRING_TRV_VOU_PAXNOME, 0)
    sSobrenomePax = String(STRING_TRV_VOU_PAXNOME, 0)
    sTitular = String(STRING_TRV_VOU_TITULAR, 0)
    sTitularCPF = String(STRING_CGC, 0)
    sClienteFat = String(STRING_CLIENTE_NOME_REDUZIDO, 0)
    sNumAuto = String(STRING_TRV_VOU_NUMAUTO, 0)
    sValidadeCC = String(STRING_TRV_VOU_VALIDADECC, 0)

    dtDataAuto = gdtDataAtual 'Date
    
    'Lê vouchers pagos com cartão de crédito
    lErro = Comando_Executar(alComando(0), sSQL, sCia, sNumCCred, sNumAuto, sValidadeCC, lNumVou, sSerie, sTipo, dtData, lClienteVou, sClienteVou, _
    dValor, dValorComissao, sNomePax, sSobrenomePax, sTitular, sTitularCPF, lClienteFat, sClienteFat, iNumParc, dTaxa, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 200669
    
    'Busca o primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200670
        
    'Enquanto existirem arquivos
    Do While lErro = AD_SQL_SUCESSO
    
        'Se modificou o número do cartão de crédito, o administradora, a data ou o número de autorização deve gerar um novo título
        If sCiaAnt <> sCia Or sNumCCredAnt <> sNumCCred Or sNumAutoAnt <> sNumAuto Then
            
            sCiaAnt = sCia
            sNumCCredAnt = sNumCCred
            sNumAutoAnt = sNumAuto
            
            Set objVouAutoCC = New ClassTRVVouAutoCC
            
            objVouAutoCC.dTarifa = dTaxa
            objVouAutoCC.dtDataAutoCC = dtData 'dtDataAuto
            objVouAutoCC.iCodSeguranca = iCodSeguranca
            objVouAutoCC.iQuantParc = iNumParc
            objVouAutoCC.lClienteFat = lClienteFat
            objVouAutoCC.sNumAuto = sNumAuto
            objVouAutoCC.sBandeira = sCia
            objVouAutoCC.sClienteFat = sClienteFat
            objVouAutoCC.sNumCCred = sNumCCred
            objVouAutoCC.sTitular = sTitular
            objVouAutoCC.sTitularCPF = sTitularCPF
            objVouAutoCC.sValidadeCC = sValidadeCC
            objVouAutoCC.iFilialEmpresa = iFilialEmpresa
            
            colAuto.Add objVouAutoCC
            
        End If
    
        Set objVou = New ClassTRVFATInfoVoucher
        
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
        
    TRVFatCartao_Le = SUCESSO
     
    Exit Function
    
Erro_TRVFatCartao_Le:

    TRVFatCartao_Le = gErr
    
    Select Case gErr
          
        Case 200668
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 200669 To 200671
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVVOUCHERS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182722)
     
    End Select

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
        
    Exit Function

End Function

Public Function TP_Vendedor_LeTRV(ByVal objVendedorMaskEdBox As Object, ByVal objVendedor As Object, Optional ByVal iCria As Integer = 1) As Long
'Lê o Vendedor com Código ou NomeRed em objVendedorMaskEdBox.Text
'Devolve em objVendedor. Coloca código-NomeReduzido no .Text

Dim sVendedor As String
Dim iCodigo As Integer
Dim Vendedor As Object
Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim sDSN As String
Dim iLenDSN As Integer
Dim sParamOut As String
Dim iLenParamOut As Integer
Dim lConexaoTVA As Long
Dim objEmpresa As New ClassDicEmpresa
Dim alComandoX(1 To 1) As Long 'Tem que ler algumas informações da TVA, mesmo estando na TVI
Dim iIndice As Integer

On Error GoTo TP_Vendedor_LeTRV

    objEmpresa.lCodigo = 1 'TVA'

    lErro = Empresa_Le(objEmpresa)
    If lErro <> SUCESSO Then gError 189756

    sDSN = objEmpresa.sStringConexao
    iLenDSN = Len(sDSN)
    sParamOut = String(1024, 0)
    iLenParamOut = 1024

    lConexaoTVA = Conexao_AbrirExt(AD_SQL_DRIVER_ODBC, sDSN, iLenDSN, sParamOut, iLenParamOut)
    If lConexaoTVA = 0 Then gError 197890

    For iIndice = LBound(alComandoX) To UBound(alComandoX)
        alComandoX(iIndice) = Comando_AbrirExt(lConexaoTVA)
        If alComandoX(iIndice) = 0 Then gError 190857
    Next

    Set Vendedor = objVendedorMaskEdBox
    sVendedor = Trim(Vendedor.Text)
    
    'Tenta extrair código de sVendedor
    iCodigo = Codigo_Extrai(sVendedor)
    
    'Se é do tipo código
    If iCodigo > 0 Then
    
        objVendedor.iCodigo = iCodigo
        lErro = CF("Vendedor_Le", objVendedor, alComandoX(1))
        If lErro <> SUCESSO And lErro <> 12582 Then Error 25031
        If lErro <> SUCESSO Then Error 25032

        Vendedor.Text = CStr(objVendedor.iCodigo) & SEPARADOR & objVendedor.sNomeReduzido
        
    Else  'Se é do tipo String
            
         objVendedor.sNomeReduzido = sVendedor
         lErro = CF("Vendedor_Le_NomeReduzido", objVendedor, alComandoX(1))
         If lErro <> SUCESSO And lErro <> 25008 Then Error 25029
         If lErro <> SUCESSO Then Error 25030
        
         Vendedor.Text = CStr(objVendedor.iCodigo) & SEPARADOR & sVendedor
        
    End If

    'Fecha Comando
    For iIndice = LBound(alComandoX) To UBound(alComandoX)
        Call Comando_Fechar(alComandoX(iIndice))
    Next
    
    Call Conexao_FecharExt(lConexaoTVA)

    TP_Vendedor_LeTRV = SUCESSO

    Exit Function

TP_Vendedor_LeTRV:

    TP_Vendedor_LeTRV = Err

    Select Case Err
        
        Case 25029 To 25032 'Tratados nas rotinas chamadas

        Case 25030  'Vendedor com NomeReduzido não cadastrado

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153617)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComandoX) To UBound(alComandoX)
        Call Comando_Fechar(alComandoX(iIndice))
    Next
    
    Call Conexao_FecharExt(lConexaoTVA)

End Function

Public Function TRVEstVenda_Le(ByVal objEstVenda As ClassTRVEstVenda) As Long

Dim lErro As Long
Dim alComando(0 To 6) As Long
Dim iIndice As Integer
Dim sSQL1 As String
Dim sSQL2 As String
Dim sSQL3 As String
Dim sSQL4 As String
Dim sSQL5 As String
Dim sSQL6 As String
Dim sSQL7 As String
Dim dValorFat As Double
Dim dQtdeFat As Long
Dim dValorBrutoVouRS As Double
Dim dValorBrutoVouUSS As Double
Dim dValorFatVou As Double
Dim dQtdeVou As Long
Dim dQtdePax As Long

On Error GoTo Erro_TRVEstVenda_Le

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 200847
    Next
    
    'Faturas a receber - Todas
    sSQL1 = "SELECT SUM(Valor), COUNT(NumIntDoc) FROM TRVFaturas WHERE DataEmissao BETWEEN ? AND ? AND (Cliente >=? OR ?=0) AND (Cliente <=? OR ?=0) AND TipoDocDestino = 3"
    'Faturas a pagar - Todas
    sSQL2 = "SELECT SUM(Valor), COUNT(NumIntDoc) FROM TRVFaturas WHERE DataEmissao BETWEEN ? AND ? AND (Cliente >=? OR ?=0) AND (Cliente <=? OR ?=0) AND TipoDocDestino IN (4,5)"
    'Faturas a receber - Canceladas
    sSQL3 = "SELECT SUM(Valor), COUNT(NumIntDoc) FROM TRVFaturas WHERE DataEmissao BETWEEN ? AND ? AND (Cliente >=? OR ?=0) AND (Cliente <=? OR ?=0) AND TipoDocDestino = 3 AND Status = 5"
    'Faturas a pagar - Canceladas
    sSQL4 = "SELECT SUM(Valor), COUNT(NumIntDoc) FROM TRVFaturas WHERE DataEmissao BETWEEN ? AND ? AND (Cliente >=? OR ?=0) AND (Cliente <=? OR ?=0) AND TipoDocDestino IN (4,5) AND Status = 5"
    'Vouchers emitidos - Todos
    sSQL5 = "SELECT SUM(Pax), COUNT(V.NumVou) , SUM(V.ValorBruto), SUM(V.Valor), SUM(CASE WHEN V.Produto = 'NA        ' THEN V.ValorBruto/C.Cambio ELSE V.ValorCambio END) FROM TRVVouchers AS V, TRVCambio As C WHERE V.TipoDoc = 'VOU' AND V.Data=C.Data AND V.Data BETWEEN ? AND ? AND (V.ClienteVou >=? OR ?=0) AND (V.ClienteVou <=? OR ?=0)"
    'Vouchers emitidos - Cancelados
    sSQL6 = "SELECT SUM(Pax), COUNT(V.NumVou) , SUM(V.ValorBruto), SUM(V.Valor), SUM(CASE WHEN V.Produto = 'NA        ' THEN V.ValorBruto/C.Cambio ELSE V.ValorCambio END) FROM TRVVouchers AS V, TRVCambio As C WHERE V.TipoDoc = 'VOU' AND V.Data=C.Data AND V.Data BETWEEN ? AND ? AND (V.ClienteVou >=? OR ?=0) AND (V.ClienteVou <=? OR ?=0) AND V.Status = 7 AND V.NumFatCoinfo = 0"
    'Vouchers emitidos e faturados
    sSQL7 = "SELECT SUM(V.Pax), COUNT(V.NumVou) , SUM(V.ValorBruto), SUM(V.Valor), SUM(CASE WHEN V.Produto = 'NA        ' THEN V.ValorBruto/C.Cambio ELSE V.ValorCambio END) FROM TRVVouchers AS V, TitulosRecTodos as T, TRVCambio As C WHERE V.TipoDoc = 'VOU' AND V.Data=C.Data AND V.Data BETWEEN ? AND ? AND V.TipoDocDestino = 3 AND V.NumIntDocdestino = T.NumIntDoc AND T.DataEmissao BETWEEN ? AND ? AND (T.Cliente >=? OR ?=0) AND (T.Cliente <=? OR ?=0)"

    'Le Faturas a receber - Todas
    lErro = Comando_Executar(alComando(0), sSQL1, dValorFat, dQtdeFat, objEstVenda.dtDataFatDe, objEstVenda.dtDataFatAte, objEstVenda.lClienteDe, objEstVenda.lClienteDe, objEstVenda.lClienteAte, objEstVenda.lClienteAte)
    If lErro <> AD_SQL_SUCESSO Then gError 200848

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200849
    
    objEstVenda.dValorFatCR = dValorFat
    objEstVenda.dQtdeFatCR = dQtdeFat
    
    'Le Faturas a pagar - Todas
    lErro = Comando_Executar(alComando(1), sSQL2, dValorFat, dQtdeFat, objEstVenda.dtDataFatDe, objEstVenda.dtDataFatAte, objEstVenda.lClienteDe, objEstVenda.lClienteDe, objEstVenda.lClienteAte, objEstVenda.lClienteAte)
    If lErro <> AD_SQL_SUCESSO Then gError 200850

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200851
    
    objEstVenda.dValorFatCP = dValorFat
    objEstVenda.dQtdeFatCP = dQtdeFat
    
    'Le Faturas a receber - Canceladas
    lErro = Comando_Executar(alComando(2), sSQL3, dValorFat, dQtdeFat, objEstVenda.dtDataFatDe, objEstVenda.dtDataFatAte, objEstVenda.lClienteDe, objEstVenda.lClienteDe, objEstVenda.lClienteAte, objEstVenda.lClienteAte)
    If lErro <> AD_SQL_SUCESSO Then gError 200852

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200853
    
    objEstVenda.dValorFatCRCanc = dValorFat
    objEstVenda.dQtdeFatCRCanc = dQtdeFat
    
    'Le Faturas a pagar - Canceladas
    lErro = Comando_Executar(alComando(3), sSQL4, dValorFat, dQtdeFat, objEstVenda.dtDataFatDe, objEstVenda.dtDataFatAte, objEstVenda.lClienteDe, objEstVenda.lClienteDe, objEstVenda.lClienteAte, objEstVenda.lClienteAte)
    If lErro <> AD_SQL_SUCESSO Then gError 200854

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200855
    
    objEstVenda.dValorFatCPCanc = dValorFat
    objEstVenda.dQtdeFatCPCanc = dQtdeFat
    
    'Vouchers emitidos - Todos
    lErro = Comando_Executar(alComando(4), sSQL5, dQtdePax, dQtdeVou, dValorBrutoVouRS, dValorFatVou, dValorBrutoVouUSS, objEstVenda.dtDataEmiDe, objEstVenda.dtDataEmiAte, objEstVenda.lClienteDe, objEstVenda.lClienteDe, objEstVenda.lClienteAte, objEstVenda.lClienteAte)
    If lErro <> AD_SQL_SUCESSO Then gError 200856

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(4))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200857
    
    objEstVenda.dQtdeEmiPax = dQtdePax
    objEstVenda.dQtdeEmiVou = dQtdeVou
    objEstVenda.dValorBrutoVouRS = dValorBrutoVouRS
    objEstVenda.dValorFatVou = dValorFatVou
    objEstVenda.dValorBrutoVouUSS = dValorBrutoVouUSS
    
    'Vouchers emitidos - Cancelados
    lErro = Comando_Executar(alComando(5), sSQL6, dQtdePax, dQtdeVou, dValorBrutoVouRS, dValorFatVou, dValorBrutoVouUSS, objEstVenda.dtDataEmiDe, objEstVenda.dtDataEmiAte, objEstVenda.lClienteDe, objEstVenda.lClienteDe, objEstVenda.lClienteAte, objEstVenda.lClienteAte)
    If lErro <> AD_SQL_SUCESSO Then gError 200858

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(5))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200859
    
    objEstVenda.dQtdeEmiPaxCanc = dQtdePax
    objEstVenda.dQtdeEmiVouCanc = dQtdeVou
    objEstVenda.dValorBrutoVouRSCanc = dValorBrutoVouRS
    objEstVenda.dValorFatVouCanc = dValorFatVou
    objEstVenda.dValorBrutoVouUSSCanc = dValorBrutoVouUSS
    
    'Vouchers emitidos e faturados
    lErro = Comando_Executar(alComando(6), sSQL7, dQtdePax, dQtdeVou, dValorBrutoVouRS, dValorFatVou, dValorBrutoVouUSS, objEstVenda.dtDataEmiDe, objEstVenda.dtDataEmiAte, objEstVenda.dtDataFatDe, objEstVenda.dtDataFatAte, objEstVenda.lClienteDe, objEstVenda.lClienteDe, objEstVenda.lClienteAte, objEstVenda.lClienteAte)
    If lErro <> AD_SQL_SUCESSO Then gError 200860

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(6))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200861
    
    objEstVenda.dQtdeEmiPaxFat = dQtdePax
    objEstVenda.dQtdeEmiVouFat = dQtdeVou
    objEstVenda.dValorBrutoVouRSFat = dValorBrutoVouRS
    objEstVenda.dValorFatVouFat = dValorFatVou
    objEstVenda.dValorBrutoVouUSSFat = dValorBrutoVouUSS
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TRVEstVenda_Le = SUCESSO

    Exit Function

Erro_TRVEstVenda_Le:

    TRVEstVenda_Le = gErr

    Select Case gErr

        Case 200847
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 200848 To 200861
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVVOUCHERS", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200862)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function TRV_Analisa_Integracao(sMensagemArq As String, sMensagemDados As String, iRetorno As Integer) As Long

Dim lErro As Long
Dim alComando(0 To 6) As Long
Dim iIndice As Integer
Dim sArquivo As String
Dim sHora As String
Dim sData As String
Dim tVou As typeTRVVouchers
Dim sMensagemErros As String
Dim lAux As Long

On Error GoTo Erro_TRV_Analisa_Integracao

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 200892
    Next
    
    iRetorno = SUCESSO
    
    sArquivo = String(255, 0)

    lErro = Comando_Executar(alComando(0), "SELECT MAX(IdentificacaoArq) FROM CoinfoArqsLidos WHERE NomeArq = 'fpitemfatu' AND DataImportacao <> ?", _
                sArquivo, DATA_NULA)
    If lErro <> AD_SQL_SUCESSO Then gError 200893

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200894

    sData = Mid(sArquivo, 15, 2) & "/" & Mid(sArquivo, 13, 2) & "/" & Mid(sArquivo, 9, 4)
    sHora = Mid(sArquivo, 17, 2) & ":" & Mid(sArquivo, 19, 2)

    sMensagemArq = "O último arquivo importado pelo Corporator foi gerado pelo SisWeb no dia " & sData & " às " & sHora & "hs."

    'If StrParaInt(Mid(sArquivo, 15, 2)) <> Day(Date) Or Hour(Now) - 2 > StrParaInt(Mid(sArquivo, 17, 2)) Then
    If StrParaInt(Mid(sArquivo, 15, 2)) <> Day(Date) Then
        'Se o último arquivo foi gerado a mais de duas horas=> Erro=> Defasado => Coloca a mensagem em vermelho
        sMensagemDados = sMensagemDados & IIf(Len(Trim(sMensagemDados)) > 0, vbNewLine, "") & "O último arquivo importado está muito defasado."
        iRetorno = vbCritical
    End If
    
    tVou.sTipVou = String(STRING_TRV_OCR_TIPOVOU, 0)
    tVou.sSerie = String(STRING_TRV_OCR_SERIE, 0)

    lErro = Comando_Executar(alComando(1), "SELECT TipVou, Serie, NumVou FROM TRVVouchers WHERE NumFatCoinfo = 0 AND Status <> 7 AND TipoDoc = 'VOU' AND ExtraiInfoSigav = 1 ORDER BY Data, NumVou DESC", _
                tVou.sTipVou, tVou.sSerie, tVou.lNumVou)
    If lErro <> AD_SQL_SUCESSO Then gError 200895

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200896

    If lErro <> AD_SQL_SEM_DADOS Then
    
        If iRetorno <> vbCritical Then iRetorno = vbInformation
    
        tVou.sDestino = String(STRING_TRV_VOU_DESTINO, 0)
        tVou.sUsuarioWeb = String(STRING_TRV_VOU_USUARIOWEB, 0)
    
        lErro = Comando_Executar(alComando(2), "SELECT CodEmpVou, Data, Emissor, Destino, Usuario FROM CoinfoItemFatura WHERE TipVou = ? AND Serie = ? AND NumVou = ?  ORDER BY NumIntArq DESC, NumRegArqTexto DESC", _
                    tVou.lClienteVou, tVou.dtData, tVou.lEmissor, tVou.sDestino, tVou.sUsuarioWeb, tVou.sTipVou, tVou.sSerie, tVou.lNumVou)
        If lErro <> AD_SQL_SUCESSO Then gError 200897
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200898
    
        sMensagemDados = sMensagemDados & IIf(Len(Trim(sMensagemDados)) > 0, vbNewLine, "") & "Alguns vouchers não tiveram a importação concluída. Segue dados abaixo:"
        sMensagemDados = sMensagemDados & vbNewLine & "VOU: " & tVou.sTipVou & tVou.sSerie & CStr(tVou.lNumVou) & " CÓD.CLI: " & CStr(tVou.lClienteVou) & " DATA EMI.: " & Format(tVou.dtData, "dd/mm/yyyy") & " CÓD.EMI.CORP: " & CStr(tVou.lEmissor + FATOR_SOMA_COD_EMISSOR) & " CÓD.EMI.SW: " & CStr(IIf(tVou.lEmissor > 1000000, tVou.lEmissor - 1000000, tVou.lEmissor)) & " CÓD.DESTINO: " & tVou.sDestino & " USU: " & tVou.sUsuarioWeb
    
        'Le o código do destino
        lErro = Comando_Executar(alComando(3), "SELECT CodValor FROM CamposGenericosValores WHERE CodCampo = ? AND Valor = ? ", _
                    lAux, CAMPOSGENERICOS_DESTINO_VIAGEM, tVou.sDestino)
        If lErro <> AD_SQL_SUCESSO Then gError 200899
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200900
        
        If lErro <> AD_SQL_SUCESSO Then sMensagemErros = sMensagemErros & IIf(Len(Trim(sMensagemErros)) > 0, vbNewLine, "") & "Código de destino não cadastrado no Corporator."
            
        'Le o código do Cliente
        lErro = Comando_Executar(alComando(4), "SELECT Codigo FROM Clientes WHERE Codigo = ? ", _
                    lAux, tVou.lClienteVou)
        If lErro <> AD_SQL_SUCESSO Then gError 200901
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(4))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200902
        
        If lErro <> AD_SQL_SUCESSO Then sMensagemErros = sMensagemErros & IIf(Len(Trim(sMensagemErros)) > 0, vbNewLine, "") & "Código de cliente não cadastrado no Corporator."
            
        'Le o código do Emissor
        lErro = Comando_Executar(alComando(5), "SELECT Codigo FROM Fornecedores WHERE Codigo = ? ", _
                    lAux, tVou.lEmissor + FATOR_SOMA_COD_EMISSOR)
        If lErro <> AD_SQL_SUCESSO Then gError 200903
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(5))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200904
        
        If lErro <> AD_SQL_SUCESSO Then sMensagemErros = sMensagemErros & IIf(Len(Trim(sMensagemErros)) > 0, vbNewLine, "") & "Código de emissor não cadastrado no Corporator."
            
        'Le o código do Cliente\Emissor
        lErro = Comando_Executar(alComando(6), "SELECT NumIntDoc FROM TRVCliEmissores WHERE Cliente = ? AND FornEmissor = ? ", _
                    lAux, tVou.lClienteVou, tVou.lEmissor + FATOR_SOMA_COD_EMISSOR)
        If lErro <> AD_SQL_SUCESSO Then gError 200905
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(6))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200906
        
        If lErro <> AD_SQL_SUCESSO Then sMensagemErros = sMensagemErros & IIf(Len(Trim(sMensagemErros)) > 0, vbNewLine, "") & "O emissor não está associado ao cliente."
    
        If Len(Trim(sMensagemErros)) > 0 Then sMensagemDados = sMensagemDados & vbNewLine & vbNewLine & sMensagemErros
    
    End If

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TRV_Analisa_Integracao = SUCESSO

    Exit Function

Erro_TRV_Analisa_Integracao:

    TRV_Analisa_Integracao = gErr

    Select Case gErr

        Case 200892
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 200893 To 200906
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVVOUCHERS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200907)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function TRVCliDataCallCenter_Le(ByVal objCli As ClassCliente) As Long

Dim lErro As Long
Dim lComando As Long
Dim tTRVCliDataCallCenter As typeTRVCliDataCallCenter
Dim objTRVCliDataCallCenter As ClassTRVCliDataCallCenter
Dim objClienteTRV As ClassClienteTRV

On Error GoTo Erro_TRVCliDataCallCenter_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 202709
    
    Set objClienteTRV = objCli.objInfoUsu

    'Le a tabelaTRVCliDataCallCenter
    lErro = Comando_Executar(lComando, "SELECT Cliente, DataDe, DataAte FROM TRVCliDataCallCenter WHERE Cliente= ? ", _
                tTRVCliDataCallCenter.lCliente, tTRVCliDataCallCenter.dtDataDe, tTRVCliDataCallCenter.dtDataAte, _
                objCli.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 202710

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202711

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objTRVCliDataCallCenter = New ClassTRVCliDataCallCenter

        objTRVCliDataCallCenter.lCliente = tTRVCliDataCallCenter.lCliente
        objTRVCliDataCallCenter.dtDataDe = tTRVCliDataCallCenter.dtDataDe
        objTRVCliDataCallCenter.dtDataAte = tTRVCliDataCallCenter.dtDataAte
        
        objClienteTRV.colDataCallCenter.Add objTRVCliDataCallCenter

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202712

    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    TRVCliDataCallCenter_Le = SUCESSO

    Exit Function

Erro_TRVCliDataCallCenter_Le:

    TRVCliDataCallCenter_Le = gErr

    Select Case gErr

        Case 202709
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 202710 To 202712
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVCLIDATACALLCENTER", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202713)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function TRVVoucher_Verifica_CallCenter(ByVal lCliente As Long, ByVal dtData As Date, bCliEhCallCenter As Boolean) As Long

Dim lErro As Long
Dim lComando As Long
Dim lClienteAux As Long

On Error GoTo Erro_TRVVoucher_Verifica_CallCenter

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 202720
    
    bCliEhCallCenter = False

    'Le a tabelaTRVCliDataCallCenter
    lErro = Comando_Executar(lComando, "SELECT Cliente FROM TRVCliDataCallCenter WHERE Cliente= ? AND ((DataDe <= ? OR DataDe = ?) AND (DataAte >= ? OR DataAte = ?)) ", _
                lClienteAux, lCliente, dtData, DATA_NULA, dtData, DATA_NULA)
    If lErro <> AD_SQL_SUCESSO Then gError 202721

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202722

    If lErro = AD_SQL_SUCESSO Then bCliEhCallCenter = True
    
    'Fecha Comando
    Call Comando_Fechar(lComando)

    TRVVoucher_Verifica_CallCenter = SUCESSO

    Exit Function

Erro_TRVVoucher_Verifica_CallCenter:

    TRVVoucher_Verifica_CallCenter = gErr

    Select Case gErr

        Case 202720
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 202721 To 202722
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVCLIDATACALLCENTER", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202723)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function TRVVouEmiCI_Le(ByVal objVou As ClassTRVVouchers) As Long

Dim lErro As Long
Dim lComando As Long
Dim tTRVVouEmiCI As typeTRVVouEmiCI
Dim objTRVVouEmiCI As ClassTRVVouEmiCI

On Error GoTo Erro_TRVVouEmiCI_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 206570
    
    Set objVou.colEmiCI = New Collection
    
    'Le a tabelaTRVVouEmiCI
    lErro = Comando_Executar(lComando, "SELECT FornEmissor, PercCI FROM TRVVouEmiCI WHERE NumVou= ?  AND Serie= ?  AND TipVou= ? ", _
                tTRVVouEmiCI.lFornEmissor, tTRVVouEmiCI.dPercCI, objVou.lNumVou, objVou.sSerie, objVou.sTipVou)
    If lErro <> AD_SQL_SUCESSO Then gError 206571

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 206572

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objTRVVouEmiCI = New ClassTRVVouEmiCI

        objTRVVouEmiCI.lNumVou = objVou.lNumVou
        objTRVVouEmiCI.sSerie = objVou.sSerie
        objTRVVouEmiCI.sTipVou = objVou.sTipVou
        objTRVVouEmiCI.lFornEmissor = tTRVVouEmiCI.lFornEmissor
        objTRVVouEmiCI.dPercCI = tTRVVouEmiCI.dPercCI
        
        objVou.colEmiCI.Add objTRVVouEmiCI

        'Busca Próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 206573

    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    TRVVouEmiCI_Le = SUCESSO

    Exit Function

Erro_TRVVouEmiCI_Le:

    TRVVouEmiCI_Le = gErr

    Select Case gErr

        Case 206570
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 206571 To 206573
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVVOUEMICI", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206574)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function TRVCliEmissores_Le(ByVal objCliEmissoresTRV As ClassCliEmissoresTRV) As Long

Dim lErro As Long
Dim lComando As Long
Dim tCliEmissoresTRV As typeCliEmissoresTRV

On Error GoTo Erro_TRVCliEmissores_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 206590

    'Alocação de espaço no buffer
    tCliEmissoresTRV.sNumCartao = String(255, 0)

    'Le a tabelaTRVCliEmissores
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Cliente, Seq, FornEmissor, PercComissao, PercCI, " & _
                "Cargo, NumCartao FROM TRVCliEmissores WHERE FornEmissor= ? ", _
                tCliEmissoresTRV.lNumIntDoc, tCliEmissoresTRV.lCliente, tCliEmissoresTRV.iSeq, tCliEmissoresTRV.lFornEmissor, _
                tCliEmissoresTRV.dPercComissao, tCliEmissoresTRV.dPercCI, tCliEmissoresTRV.lCargo, tCliEmissoresTRV.sNumCartao, _
                objCliEmissoresTRV.lFornEmissor)
    If lErro <> AD_SQL_SUCESSO Then gError 206591

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 206592

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objCliEmissoresTRV.lNumIntDoc = tCliEmissoresTRV.lNumIntDoc
    objCliEmissoresTRV.lCliente = tCliEmissoresTRV.lCliente
    objCliEmissoresTRV.iSeq = tCliEmissoresTRV.iSeq
    objCliEmissoresTRV.lFornEmissor = tCliEmissoresTRV.lFornEmissor
    objCliEmissoresTRV.dPercComissao = tCliEmissoresTRV.dPercComissao
    objCliEmissoresTRV.dPercCI = tCliEmissoresTRV.dPercCI
    objCliEmissoresTRV.lCargo = tCliEmissoresTRV.lCargo
    objCliEmissoresTRV.sNumCartao = tCliEmissoresTRV.sNumCartao

    'Fecha Comando
    Call Comando_Fechar(lComando)

    TRVCliEmissores_Le = SUCESSO

    Exit Function

Erro_TRVCliEmissores_Le:

    TRVCliEmissores_Le = gErr

    Select Case gErr

        Case 206590
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 206591, 206592
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVCLIEMISSORES", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206593)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function TRVTitulosExp_Le(ByVal objTRVTitulosExp As ClassTRVTitulosExp) As Long

Dim lErro As Long
Dim lComando As Long
Dim tTRVTitulosExp As typeTRVTitulosExp

On Error GoTo Erro_TRVTitulosExp_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 206601

    'Alocação de espaço no buffer
    tTRVTitulosExp.sUsuario = String(50, 0)
    tTRVTitulosExp.sMotivo = String(255, 0)

    'Le a tabelaTRVTitulosExp
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Usuario, Data, Hora, TipoDocOrigem, NumIntDocOrigem, " & _
                "NumTitulo, Excluido, Exportado,  " & _
                "Motivo, TemQueContabilizar, ValorAporteCred, ValorAporte, ValorComissao, ValorCreditos, ValorDebitos, " & _
                "ValorTarifa, ValorDeducoes FROM TRVTitulosExp WHERE TipoDocOrigem= ?  AND NumIntDocOrigem= ?  AND Excluido= ? ", _
                tTRVTitulosExp.lNumIntDoc, tTRVTitulosExp.sUsuario, tTRVTitulosExp.dtData, tTRVTitulosExp.dHora, _
                tTRVTitulosExp.iTipoDocOrigem, tTRVTitulosExp.lNumIntDocOrigem, tTRVTitulosExp.lNumTitulo, tTRVTitulosExp.iExcluido, tTRVTitulosExp.iExportado, _
                tTRVTitulosExp.sMotivo, tTRVTitulosExp.iTemQueContabilizar, tTRVTitulosExp.dValorAporteCred, tTRVTitulosExp.dValorAporte, tTRVTitulosExp.dValorComissao, tTRVTitulosExp.dValorCreditos, _
                tTRVTitulosExp.dValorDebitos, tTRVTitulosExp.dValorTarifa, tTRVTitulosExp.dValorDeducoes, _
                objTRVTitulosExp.iTipoDocOrigem, objTRVTitulosExp.lNumIntDocOrigem, objTRVTitulosExp.iExcluido)
    If lErro <> AD_SQL_SUCESSO Then gError 206602

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 206603

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objTRVTitulosExp.lNumIntDoc = tTRVTitulosExp.lNumIntDoc
    objTRVTitulosExp.sUsuario = tTRVTitulosExp.sUsuario
    objTRVTitulosExp.dtData = tTRVTitulosExp.dtData
    objTRVTitulosExp.dHora = tTRVTitulosExp.dHora
    objTRVTitulosExp.iTipoDocOrigem = tTRVTitulosExp.iTipoDocOrigem
    objTRVTitulosExp.lNumIntDocOrigem = tTRVTitulosExp.lNumIntDocOrigem
    objTRVTitulosExp.lNumTitulo = tTRVTitulosExp.lNumTitulo
    objTRVTitulosExp.iExcluido = tTRVTitulosExp.iExcluido
    objTRVTitulosExp.iExportado = tTRVTitulosExp.iExportado
    objTRVTitulosExp.sMotivo = tTRVTitulosExp.sMotivo
    objTRVTitulosExp.iTemQueContabilizar = tTRVTitulosExp.iTemQueContabilizar
    objTRVTitulosExp.dValorAporteCred = tTRVTitulosExp.dValorAporteCred
    objTRVTitulosExp.dValorAporte = tTRVTitulosExp.dValorAporte
    objTRVTitulosExp.dValorComissao = tTRVTitulosExp.dValorComissao
    objTRVTitulosExp.dValorCreditos = tTRVTitulosExp.dValorCreditos
    objTRVTitulosExp.dValorDebitos = tTRVTitulosExp.dValorDebitos
    objTRVTitulosExp.dValorTarifa = tTRVTitulosExp.dValorTarifa
    objTRVTitulosExp.dValorDeducoes = tTRVTitulosExp.dValorDeducoes

    'Fecha Comando
    Call Comando_Fechar(lComando)

    TRVTitulosExp_Le = SUCESSO

    Exit Function

Erro_TRVTitulosExp_Le:

    TRVTitulosExp_Le = gErr

    Select Case gErr

        Case 206601
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 206602, 206603
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVTITULOSEXP", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206604)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function TRVOcrCasos_Le_NumVou(ByVal objOcrCaso As ClassTRVOcrCasos, ByVal colRetorno As Collection) As Long

Dim lErro As Long
Dim lComando As Long
Dim tOcrCasos As typeTRVOcrCasos

On Error GoTo Erro_TRVOcrCasos_Le_NumVou

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 208545

    'Alocação de espaço no buffer
    tOcrCasos.sCodigo = String(STRING_TRV_MAXIMO, 0)
    
    'Se tem o voucher preenchido
    If objOcrCaso.lNumVou <> 0 Or Len(Trim(objOcrCaso.sNumVouTexto)) > 0 Then
            
        If objOcrCaso.lNumVou <> 0 Then
            lErro = Comando_Executar(lComando, "SELECT Codigo FROM TrvOcrCasos WHERE NumVou = ? AND (TipVou = ? OR ? = ? OR TipVou = ?) AND (Serie = ? OR ? = ? OR Serie = ?) ", tOcrCasos.sCodigo, objOcrCaso.lNumVou, objOcrCaso.sTipVou, objOcrCaso.sTipVou, "", "", objOcrCaso.sSerie, objOcrCaso.sSerie, "", "")
        Else
            lErro = Comando_Executar(lComando, "SELECT Codigo FROM TrvOcrCasos WHERE NumVouTexto = ?  ", tOcrCasos.sCodigo, objOcrCaso.sNumVouTexto)
        End If
        If lErro <> AD_SQL_SUCESSO Then gError 208546
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208547
    
        Do While lErro <> AD_SQL_SEM_DADOS
        
            colRetorno.Add tOcrCasos.sCodigo
    
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208548
    
        Loop
        
    End If
    
    'Fecha Comando
    Call Comando_Fechar(lComando)

    TRVOcrCasos_Le_NumVou = SUCESSO

    Exit Function

Erro_TRVOcrCasos_Le_NumVou:

    TRVOcrCasos_Le_NumVou = gErr

    Select Case gErr

        Case 208545
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 208546 To 208548
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVOCRCASOS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208549)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function TRVOcrCasos_Le(ByVal objOcrCaso As ClassTRVOcrCasos) As Long

Dim lErro As Long
Dim lComando As Long
Dim tOcrCasos As typeTRVOcrCasos
Dim sTexto As String

On Error GoTo Erro_TRVOcrCasos_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 208506

    'Alocação de espaço no buffer
    tOcrCasos.sCodigo = String(STRING_TRV_MAXIMO, 0)
    tOcrCasos.sTipVou = String(STRING_TRV_OCR_TIPOVOU, 0)
    tOcrCasos.sSerie = String(STRING_TRV_OCR_SERIE, 0)
    tOcrCasos.spaxnome = String(STRING_TRV_MAXIMO, 0)
    tOcrCasos.sTitularNome = String(STRING_TRV_MAXIMO, 0)
    tOcrCasos.sProduto = String(STRING_PRODUTO, 0)
    tOcrCasos.sNumProcesso = String(STRING_TRV_MAXIMO, 0)
    tOcrCasos.sComarca = String(STRING_TRV_MAXIMO, 0)
    tOcrCasos.sAgencia = String(STRING_AGENCIA, 0)
    tOcrCasos.sContaCorrente = String(STRING_CONTA_CORRENTE, 0)
    tOcrCasos.sNomeFavorecido = String(STRING_TRV_MAXIMO, 0)
    tOcrCasos.sFavorecidoCGC = String(STRING_CGC, 0)
    tOcrCasos.sNumVouTexto = String(STRING_TRV_MAXIMO, 0)

    'Le a tabelaTRVOcrCasos
    If objOcrCaso.lNumIntDoc = 0 Then
        lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Codigo, TipVou, Serie, NumVou, PaxNome, " & _
                    "TitularNome, ClienteVou, DataEmissao, DataIda, DataVolta, Produto, QtdPax, " & _
                    "Endereco, DataAbertura, DataDocsRec, DataEnvioAnalise, CGAnalise, CGStatus, CGAutorizadoPor, " & _
                    "ValorAutorizadoTotalRS, ValorAutorizadoTotalUS, DataLimite, DataEnvioFinac, DataProgFinanc, DataPagtoPax, Judicial, " & _
                    "NumProcesso, Condenado, ValorCondenacao, Comarca, DataFimProcesso, DataPagtoCond, ValorAutorizadoSeguroRS, " & _
                    "ValorAutorizadoSeguroUS, ValorAutorizadoAssistRS, ValorAutorizadoAssistUS, Cambio, AnteciparPagtoSeguro, Banco, Agencia, " & _
                    "ContaCorrente, NomeFavorecido, CodFornFavorecido, FavorecidoCGC, ValorInvoicesTotal, ValorInvoicesTotalUS, ValorDespesasTotalRS, ValorDespesasTotalUS, " & _
                    "NumIntDocTitPagCobertura, NumIntDocTitPagProcesso, NumFatCobertura, NumFatProcesso, DataPriEvento, ProcessoDanoMaterial, ProcessoDanoMoral, PerdaTipo, Procon, DataIniProcesso, NumVouTexto, ValorAutoSegRespTrvRS FROM TRVOcrCasos WHERE Codigo= ? ", _
                    tOcrCasos.lNumIntDoc, tOcrCasos.sCodigo, tOcrCasos.sTipVou, tOcrCasos.sSerie, _
                    tOcrCasos.lNumVou, tOcrCasos.spaxnome, tOcrCasos.sTitularNome, tOcrCasos.lClienteVou, tOcrCasos.dtDataEmissao, _
                    tOcrCasos.dtDataIda, tOcrCasos.dtDataVolta, tOcrCasos.sProduto, tOcrCasos.iQtdPax, tOcrCasos.lEndereco, _
                    tOcrCasos.dtDataAbertura, tOcrCasos.dtDataDocsRec, tOcrCasos.dtDataEnvioAnalise, tOcrCasos.lCGAnalise, tOcrCasos.lCGStatus, _
                    tOcrCasos.lCGAutorizadoPor, tOcrCasos.dValorAutorizadoTotalRS, tOcrCasos.dValorAutorizadoTotalUS, tOcrCasos.dtDataLimite, tOcrCasos.dtDataEnvioFinac, _
                    tOcrCasos.dtDataProgFinanc, tOcrCasos.dtDataPagtoPax, tOcrCasos.iJudicial, tOcrCasos.sNumProcesso, tOcrCasos.iCondenado, _
                    tOcrCasos.dValorCondenacao, tOcrCasos.sComarca, tOcrCasos.dtDataFimProcesso, tOcrCasos.dtDataPagtoCond, tOcrCasos.dValorAutorizadoSeguroRS, _
                    tOcrCasos.dValorAutorizadoSeguroUS, tOcrCasos.dValorAutorizadoAssistRS, tOcrCasos.dValorAutorizadoAssistUS, tOcrCasos.dCambio, tOcrCasos.iAnteciparPagtoSeguro, _
                    tOcrCasos.iBanco, tOcrCasos.sAgencia, tOcrCasos.sContaCorrente, tOcrCasos.sNomeFavorecido, tOcrCasos.lCodFornFavorecido, _
                    tOcrCasos.sFavorecidoCGC, tOcrCasos.dValorInvoicesTotal, tOcrCasos.dValorInvoicesTotalUS, tOcrCasos.dValorDespesasTotalRS, tOcrCasos.dValorDespesasTotalUS, tOcrCasos.lNumIntDocTitPagCobertura, _
                    tOcrCasos.lNumIntDocTitPagProcesso, tOcrCasos.lNumFatCobertura, tOcrCasos.lNumFatProcesso, tOcrCasos.dtDataPriEvento, tOcrCasos.dProcessoDanoMaterial, tOcrCasos.dProcessoDanoMoral, tOcrCasos.iPerdaTipo, tOcrCasos.iProcon, tOcrCasos.dtDataIniProcesso, tOcrCasos.sNumVouTexto, tOcrCasos.dValorAutoSegRespTrvRS, _
                    objOcrCaso.sCodigo)
    Else
        lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Codigo, TipVou, Serie, NumVou, PaxNome, " & _
                    "TitularNome, ClienteVou, DataEmissao, DataIda, DataVolta, Produto, QtdPax, " & _
                    "Endereco, DataAbertura, DataDocsRec, DataEnvioAnalise, CGAnalise, CGStatus, CGAutorizadoPor, " & _
                    "ValorAutorizadoTotalRS, ValorAutorizadoTotalUS, DataLimite, DataEnvioFinac, DataProgFinanc, DataPagtoPax, Judicial, " & _
                    "NumProcesso, Condenado, ValorCondenacao, Comarca, DataFimProcesso, DataPagtoCond, ValorAutorizadoSeguroRS, " & _
                    "ValorAutorizadoSeguroUS, ValorAutorizadoAssistRS, ValorAutorizadoAssistUS, Cambio, AnteciparPagtoSeguro, Banco, Agencia, " & _
                    "ContaCorrente, NomeFavorecido, CodFornFavorecido, FavorecidoCGC, ValorInvoicesTotal, ValorInvoicesTotalUS, ValorDespesasTotalRS, ValorDespesasTotalUS, " & _
                    "NumIntDocTitPagCobertura, NumIntDocTitPagProcesso, NumFatCobertura, NumFatProcesso, DataPriEvento, ProcessoDanoMaterial, ProcessoDanoMoral, PerdaTipo, Procon, DataIniProcesso, NumVouTexto, ValorAutoSegRespTrvRS FROM TRVOcrCasos WHERE NumIntDoc= ? ", _
                    tOcrCasos.lNumIntDoc, tOcrCasos.sCodigo, tOcrCasos.sTipVou, tOcrCasos.sSerie, _
                    tOcrCasos.lNumVou, tOcrCasos.spaxnome, tOcrCasos.sTitularNome, tOcrCasos.lClienteVou, tOcrCasos.dtDataEmissao, _
                    tOcrCasos.dtDataIda, tOcrCasos.dtDataVolta, tOcrCasos.sProduto, tOcrCasos.iQtdPax, tOcrCasos.lEndereco, _
                    tOcrCasos.dtDataAbertura, tOcrCasos.dtDataDocsRec, tOcrCasos.dtDataEnvioAnalise, tOcrCasos.lCGAnalise, tOcrCasos.lCGStatus, _
                    tOcrCasos.lCGAutorizadoPor, tOcrCasos.dValorAutorizadoTotalRS, tOcrCasos.dValorAutorizadoTotalUS, tOcrCasos.dtDataLimite, tOcrCasos.dtDataEnvioFinac, _
                    tOcrCasos.dtDataProgFinanc, tOcrCasos.dtDataPagtoPax, tOcrCasos.iJudicial, tOcrCasos.sNumProcesso, tOcrCasos.iCondenado, _
                    tOcrCasos.dValorCondenacao, tOcrCasos.sComarca, tOcrCasos.dtDataFimProcesso, tOcrCasos.dtDataPagtoCond, tOcrCasos.dValorAutorizadoSeguroRS, _
                    tOcrCasos.dValorAutorizadoSeguroUS, tOcrCasos.dValorAutorizadoAssistRS, tOcrCasos.dValorAutorizadoAssistUS, tOcrCasos.dCambio, tOcrCasos.iAnteciparPagtoSeguro, _
                    tOcrCasos.iBanco, tOcrCasos.sAgencia, tOcrCasos.sContaCorrente, tOcrCasos.sNomeFavorecido, tOcrCasos.lCodFornFavorecido, _
                    tOcrCasos.sFavorecidoCGC, tOcrCasos.dValorInvoicesTotal, tOcrCasos.dValorInvoicesTotalUS, tOcrCasos.dValorDespesasTotalRS, tOcrCasos.dValorDespesasTotalUS, tOcrCasos.lNumIntDocTitPagCobertura, _
                    tOcrCasos.lNumIntDocTitPagProcesso, tOcrCasos.lNumFatCobertura, tOcrCasos.lNumFatProcesso, tOcrCasos.dtDataPriEvento, tOcrCasos.dProcessoDanoMaterial, tOcrCasos.dProcessoDanoMoral, tOcrCasos.iPerdaTipo, tOcrCasos.iProcon, tOcrCasos.dtDataIniProcesso, tOcrCasos.sNumVouTexto, tOcrCasos.dValorAutoSegRespTrvRS, _
                    objOcrCaso.lNumIntDoc)
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 208507

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208508

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objOcrCaso.lNumIntDoc = tOcrCasos.lNumIntDoc
    objOcrCaso.sCodigo = tOcrCasos.sCodigo
    objOcrCaso.sTipVou = tOcrCasos.sTipVou
    objOcrCaso.sSerie = tOcrCasos.sSerie
    objOcrCaso.lNumVou = tOcrCasos.lNumVou
    objOcrCaso.spaxnome = tOcrCasos.spaxnome
    objOcrCaso.sTitularNome = tOcrCasos.sTitularNome
    objOcrCaso.lClienteVou = tOcrCasos.lClienteVou
    objOcrCaso.dtDataEmissao = tOcrCasos.dtDataEmissao
    objOcrCaso.dtDataIda = tOcrCasos.dtDataIda
    objOcrCaso.dtDataVolta = tOcrCasos.dtDataVolta
    objOcrCaso.sProduto = tOcrCasos.sProduto
    objOcrCaso.iQtdPax = tOcrCasos.iQtdPax
    objOcrCaso.lEndereco = tOcrCasos.lEndereco
    objOcrCaso.dtDataAbertura = tOcrCasos.dtDataAbertura
    objOcrCaso.dtDataDocsRec = tOcrCasos.dtDataDocsRec
    objOcrCaso.dtDataEnvioAnalise = tOcrCasos.dtDataEnvioAnalise
    objOcrCaso.lCGAnalise = tOcrCasos.lCGAnalise
    objOcrCaso.lCGStatus = tOcrCasos.lCGStatus
    objOcrCaso.lCGAutorizadoPor = tOcrCasos.lCGAutorizadoPor
    objOcrCaso.dValorAutorizadoTotalRS = tOcrCasos.dValorAutorizadoTotalRS
    objOcrCaso.dValorAutorizadoTotalUS = tOcrCasos.dValorAutorizadoTotalUS
    objOcrCaso.dtDataLimite = tOcrCasos.dtDataLimite
    objOcrCaso.dtDataEnvioFinac = tOcrCasos.dtDataEnvioFinac
    objOcrCaso.dtDataProgFinanc = tOcrCasos.dtDataProgFinanc
    objOcrCaso.dtDataPagtoPax = tOcrCasos.dtDataPagtoPax
    objOcrCaso.iJudicial = tOcrCasos.iJudicial
    objOcrCaso.sNumProcesso = tOcrCasos.sNumProcesso
    objOcrCaso.iCondenado = tOcrCasos.iCondenado
    objOcrCaso.dValorCondenacao = tOcrCasos.dValorCondenacao
    objOcrCaso.sComarca = tOcrCasos.sComarca
    objOcrCaso.dtDataFimProcesso = tOcrCasos.dtDataFimProcesso
    objOcrCaso.dtDataPagtoCond = tOcrCasos.dtDataPagtoCond
    objOcrCaso.dValorAutorizadoSeguroRS = tOcrCasos.dValorAutorizadoSeguroRS
    objOcrCaso.dValorAutorizadoSeguroUS = tOcrCasos.dValorAutorizadoSeguroUS
    objOcrCaso.dValorAutorizadoAssistRS = tOcrCasos.dValorAutorizadoAssistRS
    objOcrCaso.dValorAutorizadoAssistUS = tOcrCasos.dValorAutorizadoAssistUS
    objOcrCaso.dCambio = tOcrCasos.dCambio
    objOcrCaso.iAnteciparPagtoSeguro = tOcrCasos.iAnteciparPagtoSeguro
    objOcrCaso.iBanco = tOcrCasos.iBanco
    objOcrCaso.sAgencia = tOcrCasos.sAgencia
    objOcrCaso.sContaCorrente = tOcrCasos.sContaCorrente
    objOcrCaso.sNomeFavorecido = tOcrCasos.sNomeFavorecido
    objOcrCaso.lCodFornFavorecido = tOcrCasos.lCodFornFavorecido
    objOcrCaso.sFavorecidoCGC = tOcrCasos.sFavorecidoCGC
    objOcrCaso.dValorInvoicesTotal = tOcrCasos.dValorInvoicesTotal
    objOcrCaso.dValorInvoicesTotalUS = tOcrCasos.dValorInvoicesTotalUS
    objOcrCaso.dValorDespesasTotalRS = tOcrCasos.dValorDespesasTotalRS
    objOcrCaso.dValorDespesasTotalUS = tOcrCasos.dValorDespesasTotalUS
    objOcrCaso.lNumIntDocTitPagCobertura = tOcrCasos.lNumIntDocTitPagCobertura
    objOcrCaso.lNumIntDocTitPagProcesso = tOcrCasos.lNumIntDocTitPagProcesso
    objOcrCaso.lNumFatCobertura = tOcrCasos.lNumFatCobertura
    objOcrCaso.lNumFatProcesso = tOcrCasos.lNumFatProcesso
    objOcrCaso.dtDataPriEvento = tOcrCasos.dtDataPriEvento
    objOcrCaso.dtDataIniProcesso = tOcrCasos.dtDataIniProcesso
    objOcrCaso.dProcessoDanoMaterial = tOcrCasos.dProcessoDanoMaterial
    objOcrCaso.dProcessoDanoMoral = tOcrCasos.dProcessoDanoMoral
    objOcrCaso.iProcon = tOcrCasos.iProcon
    objOcrCaso.iPerdaTipo = tOcrCasos.iPerdaTipo
    objOcrCaso.sNumVouTexto = tOcrCasos.sNumVouTexto
    objOcrCaso.dValorAutoSegRespTrvRS = tOcrCasos.dValorAutoSegRespTrvRS
    
    lErro = TRVOcrCasosSrv_Le(objOcrCaso)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = TRvOcrCasosInvoices_Le(objOcrCaso)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = TRVOcrCasosOutrasFat_Le(objOcrCaso)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
    lErro = TRVOcrCasosParcValorCond_Le(objOcrCaso)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = TRVOcrCasosTexto_Le(objOcrCaso.lNumIntDoc, TRV_OCRCASOTEXTO_TIPO_OBS, 0, sTexto)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = TRVOcrCasosDocs_Le(objOcrCaso)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = TRVOcrCasosGastosAdv_Le(objOcrCaso)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    objOcrCaso.sObs = sTexto

    lErro = CF("TRVOcrCasosHist_Le", objOcrCaso)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = CF("TRVOcrCasosAnotacoes_Le", objOcrCaso)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = CF("TRVOcrCasos_Le_Eventos", objOcrCaso)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
       
    objOcrCaso.objPreReceber.lNumIntDocOcrCaso = objOcrCaso.lNumIntDoc
    
    lErro = TRVOcrCasosPreRec_Le(objOcrCaso.objPreReceber)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
    
    Set objOcrCaso.objEndereco = New ClassEndereco
    
    objOcrCaso.objEndereco.lCodigo = objOcrCaso.lEndereco
    
    lErro = CF("Endereco_Le", objOcrCaso.objEndereco)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Fecha Comando
    Call Comando_Fechar(lComando)

    TRVOcrCasos_Le = SUCESSO

    Exit Function

Erro_TRVOcrCasos_Le:

    TRVOcrCasos_Le = gErr

    Select Case gErr

        Case 208506
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 208507, 208508
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVOCRCASOS", gErr)

        Case ERRO_LEITURA_SEM_DADOS
        
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208509)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function TRVOcrCasosSrv_Le(ByVal objOcrCasos As ClassTRVOcrCasos) As Long

Dim lErro As Long
Dim lComando As Long
Dim tOcrCasosSrv As typeTRVOcrCasosSrv
Dim objOcrCasosSrv As ClassTRVOcrCasosSrv

On Error GoTo Erro_TRVOcrCasosSrv_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 208510

    'Alocação de espaço no buffer
    tOcrCasosSrv.sDescricao = String(STRING_TRV_MAXIMO, 0)

    'Le a tabelaTRVOcrCasosSrv
    lErro = Comando_Executar(lComando, "SELECT NumIntDocOcrCaso, Seq, Solicitado, Autorizado, Descricao, ValorLimite, Moeda, " & _
                "Tipo, ValorSolicitadoRS, ValorSolicitadoUS, ValorAutorizadoRS, ValorAutorizadoUS, Solicitado, CodigoServ FROM TRVOcrCasosSrv WHERE NumIntDocOcrCaso= ? ORDER BY Seq", _
                tOcrCasosSrv.lNumIntDocOcrCaso, tOcrCasosSrv.iSeq, tOcrCasosSrv.iSolicitado, tOcrCasosSrv.iAutorizado, tOcrCasosSrv.sDescricao, _
                tOcrCasosSrv.dValorLimite, tOcrCasosSrv.iMoeda, tOcrCasosSrv.iTipo, tOcrCasosSrv.dValorSolicitadoRS, tOcrCasosSrv.dValorSolicitadoUS, _
                tOcrCasosSrv.dValorAutorizadoRS, tOcrCasosSrv.dValorAutorizadoUS, tOcrCasosSrv.iSolicitado, tOcrCasosSrv.lCodigoServ, _
                objOcrCasos.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 208511

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208512

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objOcrCasosSrv = New ClassTRVOcrCasosSrv

        objOcrCasosSrv.lNumIntDocOcrCaso = tOcrCasosSrv.lNumIntDocOcrCaso
        objOcrCasosSrv.iSeq = tOcrCasosSrv.iSeq
        objOcrCasosSrv.iAutorizado = tOcrCasosSrv.iAutorizado
        objOcrCasosSrv.iSolicitado = tOcrCasosSrv.iSolicitado
        objOcrCasosSrv.sDescricao = tOcrCasosSrv.sDescricao
        objOcrCasosSrv.dValorLimite = tOcrCasosSrv.dValorLimite
        objOcrCasosSrv.iMoeda = tOcrCasosSrv.iMoeda
        objOcrCasosSrv.iTipo = tOcrCasosSrv.iTipo
        objOcrCasosSrv.dValorSolicitadoRS = tOcrCasosSrv.dValorSolicitadoRS
        objOcrCasosSrv.dValorSolicitadoUS = tOcrCasosSrv.dValorSolicitadoUS
        objOcrCasosSrv.dValorAutorizadoRS = tOcrCasosSrv.dValorAutorizadoRS
        objOcrCasosSrv.dValorAutorizadoUS = tOcrCasosSrv.dValorAutorizadoUS
        objOcrCasosSrv.iSolicitado = tOcrCasosSrv.iSolicitado
        objOcrCasosSrv.lCodigoServ = tOcrCasosSrv.lCodigoServ
        
        objOcrCasos.colCoberturas.Add objOcrCasosSrv
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208513
        
    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    TRVOcrCasosSrv_Le = SUCESSO

    Exit Function

Erro_TRVOcrCasosSrv_Le:

    TRVOcrCasosSrv_Le = gErr

    Select Case gErr

        Case 208510
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 208511 To 208513
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVOCRCASOSSRV", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208514)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function TRvOcrCasosInvoices_Le(ByVal objOcrCasos As ClassTRVOcrCasos) As Long

Dim lErro As Long
Dim lComando As Long
Dim tOcrCasosInvoices As typeTRvOcrCasosInvoices
Dim objOcrCasosInvoices As ClassTRvOcrCasosInvoices

On Error GoTo Erro_TRvOcrCasosInvoices_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 208515

    'Alocação de espaço no buffer
    tOcrCasosInvoices.sNumero = String(STRING_TRV_MAXIMO, 0)
    tOcrCasosInvoices.sObs = String(STRING_TRV_MAXIMO, 0)

    'Le a tabelaTRvOcrCasosInvoices
    lErro = Comando_Executar(lComando, "SELECT NumIntDocOcrCaso, Seq, DataRecepcao, DataFatura, Numero, Moeda, " & _
                "ValorMoeda, ValorRS, Obs FROM TRvOcrCasosInvoices WHERE NumIntDocOcrCaso= ? ORDER BY Seq", _
                tOcrCasosInvoices.lNumIntDocOcrCaso, tOcrCasosInvoices.iSeq, tOcrCasosInvoices.dtDataRecepcao, tOcrCasosInvoices.dtDataFatura, _
                tOcrCasosInvoices.sNumero, tOcrCasosInvoices.iMoeda, tOcrCasosInvoices.dValorMoeda, tOcrCasosInvoices.dValorRS, tOcrCasosInvoices.sObs, _
                objOcrCasos.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 208516

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208517

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objOcrCasosInvoices = New ClassTRvOcrCasosInvoices

        objOcrCasosInvoices.lNumIntDocOcrCaso = tOcrCasosInvoices.lNumIntDocOcrCaso
        objOcrCasosInvoices.iSeq = tOcrCasosInvoices.iSeq
        objOcrCasosInvoices.dtDataRecepcao = tOcrCasosInvoices.dtDataRecepcao
        objOcrCasosInvoices.dtDataFatura = tOcrCasosInvoices.dtDataFatura
        objOcrCasosInvoices.sNumero = tOcrCasosInvoices.sNumero
        objOcrCasosInvoices.iMoeda = tOcrCasosInvoices.iMoeda
        objOcrCasosInvoices.dValorMoeda = tOcrCasosInvoices.dValorMoeda
        objOcrCasosInvoices.dValorRS = tOcrCasosInvoices.dValorRS
        objOcrCasosInvoices.sObs = tOcrCasosInvoices.sObs
    
        objOcrCasos.colInvoices.Add objOcrCasosInvoices
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208518
        
    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    TRvOcrCasosInvoices_Le = SUCESSO

    Exit Function

Erro_TRvOcrCasosInvoices_Le:

    TRvOcrCasosInvoices_Le = gErr

    Select Case gErr

        Case 208515
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 208516 To 208518
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVOCRCASOSINVOICES", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208519)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function TRVOcrCasosOutrasFat_Le(ByVal objOcrCasos As ClassTRVOcrCasos) As Long

Dim lErro As Long
Dim lComando As Long
Dim tOcrCasosOutrasFat As typeTRVOcrCasosOutrasFat
Dim objOcrCasosOutrasFat As ClassTRVOcrCasosOutrasFat

On Error GoTo Erro_TRVOcrCasosOutrasFat_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 208520

    'Alocação de espaço no buffer
    tOcrCasosOutrasFat.sNumero = String(STRING_TRV_MAXIMO, 0)
    tOcrCasosOutrasFat.sDescricao = String(STRING_TRV_MAXIMO, 0)

    'Le a tabelaTRVOcrCasosOutrasFat
    lErro = Comando_Executar(lComando, "SELECT NumIntDocOcrCaso, Seq, Considerar, DataRecepcao, DataFatura, Numero, " & _
                "ValorUS, ValorRS, Descricao FROM TRVOcrCasosOutrasFat WHERE NumIntDocOcrCaso= ? ORDER BY Seq", _
                tOcrCasosOutrasFat.lNumIntDocOcrCaso, tOcrCasosOutrasFat.iSeq, tOcrCasosOutrasFat.iConsiderar, tOcrCasosOutrasFat.dtDataRecepcao, _
                tOcrCasosOutrasFat.dtDataFatura, tOcrCasosOutrasFat.sNumero, tOcrCasosOutrasFat.dValorUS, tOcrCasosOutrasFat.dValorRS, tOcrCasosOutrasFat.sDescricao, _
                objOcrCasos.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 208521

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208522

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objOcrCasosOutrasFat = New ClassTRVOcrCasosOutrasFat

        objOcrCasosOutrasFat.lNumIntDocOcrCaso = tOcrCasosOutrasFat.lNumIntDocOcrCaso
        objOcrCasosOutrasFat.iSeq = tOcrCasosOutrasFat.iSeq
        objOcrCasosOutrasFat.iConsiderar = tOcrCasosOutrasFat.iConsiderar
        objOcrCasosOutrasFat.dtDataRecepcao = tOcrCasosOutrasFat.dtDataRecepcao
        objOcrCasosOutrasFat.dtDataFatura = tOcrCasosOutrasFat.dtDataFatura
        objOcrCasosOutrasFat.sNumero = tOcrCasosOutrasFat.sNumero
        objOcrCasosOutrasFat.dValorUS = tOcrCasosOutrasFat.dValorUS
        objOcrCasosOutrasFat.dValorRS = tOcrCasosOutrasFat.dValorRS
        objOcrCasosOutrasFat.sDescricao = tOcrCasosOutrasFat.sDescricao
    
        objOcrCasos.colOutrasFaturas.Add objOcrCasosOutrasFat
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208523
        
    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    TRVOcrCasosOutrasFat_Le = SUCESSO

    Exit Function

Erro_TRVOcrCasosOutrasFat_Le:

    TRVOcrCasosOutrasFat_Le = gErr

    Select Case gErr

        Case 208520
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 208521 To 208523
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVOCRCASOSOUTRASFAT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208524)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function TRVOcrCasosTexto_Le(ByVal lNumIntDoc As Long, ByVal iTipoPai As Integer, ByVal iSeqPai As Integer, sTexto As String) As Long

Dim lErro As Long
Dim lComando As Long
Dim tOcrCasosTexto As typeTRVOcrCasosTexto

On Error GoTo Erro_TRVOcrCasosTexto_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 208525
    
    sTexto = ""

    'Alocação de espaço no buffer
    tOcrCasosTexto.sTexto = String(STRING_TRV_MAXIMO, 0)

    'Le a tabelaTRVOcrCasosTexto
    lErro = Comando_Executar(lComando, "SELECT Texto FROM TRVOcrCasosTexto WHERE NumIntDocOcrCaso= ?  AND TipoPai= ?  AND SeqPai= ? ORDER BY Seq", _
                tOcrCasosTexto.sTexto, lNumIntDoc, iTipoPai, iSeqPai)
    If lErro <> AD_SQL_SUCESSO Then gError 208526

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208527

    Do While lErro <> AD_SQL_SEM_DADOS

        sTexto = sTexto & tOcrCasosTexto.sTexto

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208528
        
    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    TRVOcrCasosTexto_Le = SUCESSO

    Exit Function

Erro_TRVOcrCasosTexto_Le:

    TRVOcrCasosTexto_Le = gErr

    Select Case gErr

        Case 208525
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 208526 To 208528
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVOCRCASOSTEXTO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208529)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function TRVOcrCasosParcValorCond_Le(ByVal objOcrCasos As ClassTRVOcrCasos) As Long

Dim lErro As Long
Dim lComando As Long
Dim tOcrCasosParcCond As typeTRVOcrCasosParcCond
Dim objOcrCasosParcCond As ClassTRVOcrCasosParcCond

On Error GoTo Erro_TRVOcrCasosParcValorCond_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 208530

    'Le a tabelaTRVOcrCasosParcValorCond
    lErro = Comando_Executar(lComando, "SELECT NumIntDocOcrCaso, Seq, DataVencimento, Valor, DataPagto FROM TRVOcrCasosParcValorCond WHERE NumIntDocOcrCaso= ?  ORDER BY Seq", _
                tOcrCasosParcCond.lNumIntDocOcrCaso, tOcrCasosParcCond.iSeq, tOcrCasosParcCond.dtDataVencimento, tOcrCasosParcCond.dValor, _
                tOcrCasosParcCond.dtDataPagto, objOcrCasos.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 208531

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208532

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objOcrCasosParcCond = New ClassTRVOcrCasosParcCond

        objOcrCasosParcCond.lNumIntDocOcrCaso = tOcrCasosParcCond.lNumIntDocOcrCaso
        objOcrCasosParcCond.iSeq = tOcrCasosParcCond.iSeq
        objOcrCasosParcCond.dtDataVencimento = tOcrCasosParcCond.dtDataVencimento
        objOcrCasosParcCond.dValor = tOcrCasosParcCond.dValor
        objOcrCasosParcCond.dtDataPagto = tOcrCasosParcCond.dtDataPagto

        objOcrCasos.colParcProcesso.Add objOcrCasosParcCond
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208533
        
    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    TRVOcrCasosParcValorCond_Le = SUCESSO

    Exit Function

Erro_TRVOcrCasosParcValorCond_Le:

    TRVOcrCasosParcValorCond_Le = gErr

    Select Case gErr

        Case 208530
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 208531 To 208533
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVOCRCASOSPARCVALORCOND", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208534)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function TRVOcrCasosHist_Le(ByVal objOcrCasos As ClassTRVOcrCasos) As Long

Dim lErro As Long
Dim lComando As Long
Dim tOcrCasosHist As typeTRVOcrCasosHist
Dim objOcrCasosHist As ClassTRVOcrCasosHist
Dim sTexto As String

On Error GoTo Erro_TRVOcrCasosHist_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 208535

    'Alocação de espaço no buffer
    tOcrCasosHist.sUsuario = String(STRING_USUARIO_CODIGO, 0)

    'Le a tabelaTRVOcrCasosHist
    lErro = Comando_Executar(lComando, "SELECT NumIntDocOcrCaso, Seq, Data, Hora, DataReg, HoraReg, " & _
                "Usuario, Origem FROM TRVOcrCasosHist WHERE NumIntDocOcrCaso= ? ORDER BY Seq", _
                tOcrCasosHist.lNumIntDocOcrCaso, tOcrCasosHist.iSeq, tOcrCasosHist.dtData, tOcrCasosHist.dHora, _
                tOcrCasosHist.dtdatareg, tOcrCasosHist.dHoraReg, tOcrCasosHist.sUsuario, tOcrCasosHist.iOrigem, _
                objOcrCasos.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 208536

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208537

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objOcrCasosHist = New ClassTRVOcrCasosHist

        objOcrCasosHist.lNumIntDocOcrCaso = tOcrCasosHist.lNumIntDocOcrCaso
        objOcrCasosHist.iSeq = tOcrCasosHist.iSeq
        objOcrCasosHist.dtData = tOcrCasosHist.dtData
        objOcrCasosHist.dHora = tOcrCasosHist.dHora
        objOcrCasosHist.dtdatareg = tOcrCasosHist.dtdatareg
        objOcrCasosHist.dHoraReg = tOcrCasosHist.dHoraReg
        objOcrCasosHist.sUsuario = tOcrCasosHist.sUsuario
        objOcrCasosHist.iOrigem = tOcrCasosHist.iOrigem
        
        lErro = TRVOcrCasosTexto_Le(objOcrCasos.lNumIntDoc, TRV_OCRCASOTEXTO_TIPO_ASSUNTOHIST, objOcrCasosHist.iSeq, sTexto)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        objOcrCasosHist.sAssunto = sTexto
    
        objOcrCasos.colHistorico.Add objOcrCasosHist
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208538
        
    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    TRVOcrCasosHist_Le = SUCESSO

    Exit Function

Erro_TRVOcrCasosHist_Le:

    TRVOcrCasosHist_Le = gErr

    Select Case gErr

        Case 208535
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 208536 To 208538
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVOCRCASOSHIST", gErr)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208539)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function TRVOcrCasosAnotacoes_Le(ByVal objOcrCasos As ClassTRVOcrCasos) As Long

Dim lErro As Long
Dim lComando As Long
Dim tOcrCasosAnotacoes As typeTRVOcrCasosAnotacoes
Dim objOcrCasosAnotacoes As ClassTRVOcrCasosAnotacoes
Dim sTexto As String

On Error GoTo Erro_TRVOcrCasosAnotacoes_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 208540

    'Alocação de espaço no buffer
    tOcrCasosAnotacoes.sUsuario = String(STRING_USUARIO_CODIGO, 0)

    'Le a tabelaTRVOcrCasosAnotacoes
    lErro = Comando_Executar(lComando, "SELECT NumIntDocOcrCaso, Seq, Data, Hora, DataReg, HoraReg, " & _
                "Usuario FROM TRVOcrCasosAnotacoes WHERE NumIntDocOcrCaso= ? ORDER BY Seq", _
                tOcrCasosAnotacoes.lNumIntDocOcrCaso, tOcrCasosAnotacoes.iSeq, tOcrCasosAnotacoes.dtData, tOcrCasosAnotacoes.dHora, _
                tOcrCasosAnotacoes.dtdatareg, tOcrCasosAnotacoes.dHoraReg, tOcrCasosAnotacoes.sUsuario, _
                objOcrCasos.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 208541

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208542

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objOcrCasosAnotacoes = New ClassTRVOcrCasosAnotacoes

        objOcrCasosAnotacoes.lNumIntDocOcrCaso = tOcrCasosAnotacoes.lNumIntDocOcrCaso
        objOcrCasosAnotacoes.iSeq = tOcrCasosAnotacoes.iSeq
        objOcrCasosAnotacoes.dtData = tOcrCasosAnotacoes.dtData
        objOcrCasosAnotacoes.dHora = tOcrCasosAnotacoes.dHora
        objOcrCasosAnotacoes.dtdatareg = tOcrCasosAnotacoes.dtdatareg
        objOcrCasosAnotacoes.dHoraReg = tOcrCasosAnotacoes.dHoraReg
        objOcrCasosAnotacoes.sUsuario = tOcrCasosAnotacoes.sUsuario
    
        lErro = TRVOcrCasosTexto_Le(objOcrCasos.lNumIntDoc, TRV_OCRCASOTEXTO_TIPO_ANOTTEXTO, objOcrCasosAnotacoes.iSeq, sTexto)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        objOcrCasosAnotacoes.sTexto = sTexto
    
        objOcrCasos.colAnotacoes.Add objOcrCasosAnotacoes
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208543
        
    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    TRVOcrCasosAnotacoes_Le = SUCESSO

    Exit Function

Erro_TRVOcrCasosAnotacoes_Le:

    TRVOcrCasosAnotacoes_Le = gErr

    Select Case gErr

        Case 208540
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 208541 To 208543
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVOCRCASOSANOTACOES", gErr)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208544)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function TRVOcrCasos_Le_Eventos(ByVal objOcrCasos As ClassTRVOcrCasos) As Long

Dim lErro As Long
Dim lComando As Long
Dim objOcrCasoImport As ClassTRVOcrCasoImport
Dim tOcrCasoImport As typeTRVOcrCasoImport

On Error GoTo Erro_TRVOcrCasos_Le_Eventos

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 208550
    
    'Alocação de espaço no buffer
    tOcrCasoImport.sCidadeOCR = String(STRING_TRV_MAXIMO, 0)
    tOcrCasoImport.sEstadoOCR = String(STRING_TRV_MAXIMO, 0)
    tOcrCasoImport.sPaisOCR = String(STRING_TRV_MAXIMO, 0)
    tOcrCasoImport.sCarater = String(STRING_TRV_MAXIMO, 0)
    tOcrCasoImport.sTelefone = String(STRING_TRV_MAXIMO, 0)
    tOcrCasoImport.sCodigo = String(STRING_TRV_MAXIMO, 0)
    tOcrCasoImport.sTCliente = String(STRING_TRV_MAXIMO, 0)
    tOcrCasoImport.sNome = String(STRING_TRV_MAXIMO, 0)
    tOcrCasoImport.sSobrenome = String(STRING_TRV_MAXIMO, 0)
    tOcrCasoImport.sChaveVou = String(STRING_TRV_MAXIMO, 0)
    tOcrCasoImport.sTipVou = String(STRING_TRV_MAXIMO, 0)
    tOcrCasoImport.sSerie = String(STRING_TRV_MAXIMO, 0)
    tOcrCasoImport.sPrestador = String(STRING_TRV_MAXIMO, 0)
    tOcrCasoImport.sGrauSatisfacao = String(STRING_TRV_MAXIMO, 0)

    'Le a tabelaTRVOcrCasoImport
    lErro = Comando_Executar(lComando, "SELECT NumIntArq, Seq, Codigo, Data, TCliente, Nome, " & _
                "Sobrenome, ChaveVou, TipVou, Serie, NumVou, CidadeOCR, EstadoOCR, " & _
                "PaisOCR, Prestador, Carater, GrauSatisfacao, Telefone FROM TRVOcrCasoImport WHERE Codigo= ? ORDER BY NumIntArq, Seq", _
                tOcrCasoImport.lNumIntArq, tOcrCasoImport.lSeq, tOcrCasoImport.sCodigo, tOcrCasoImport.dtData, _
                tOcrCasoImport.sTCliente, tOcrCasoImport.sNome, tOcrCasoImport.sSobrenome, tOcrCasoImport.sChaveVou, tOcrCasoImport.sTipVou, _
                tOcrCasoImport.sSerie, tOcrCasoImport.lNumVou, tOcrCasoImport.sCidadeOCR, tOcrCasoImport.sEstadoOCR, tOcrCasoImport.sPaisOCR, _
                tOcrCasoImport.sPrestador, tOcrCasoImport.sCarater, tOcrCasoImport.sGrauSatisfacao, tOcrCasoImport.sTelefone, _
                objOcrCasos.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 208551

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208552

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objOcrCasoImport = New ClassTRVOcrCasoImport

        objOcrCasoImport.lNumIntArq = tOcrCasoImport.lNumIntArq
        objOcrCasoImport.lSeq = tOcrCasoImport.lSeq
        objOcrCasoImport.sCodigo = tOcrCasoImport.sCodigo
        objOcrCasoImport.dtData = tOcrCasoImport.dtData
        objOcrCasoImport.sTCliente = tOcrCasoImport.sTCliente
        objOcrCasoImport.sNome = tOcrCasoImport.sNome
        objOcrCasoImport.sSobrenome = tOcrCasoImport.sSobrenome
        objOcrCasoImport.sChaveVou = tOcrCasoImport.sChaveVou
        objOcrCasoImport.sTipVou = tOcrCasoImport.sTipVou
        objOcrCasoImport.sSerie = tOcrCasoImport.sSerie
        objOcrCasoImport.lNumVou = tOcrCasoImport.lNumVou
        objOcrCasoImport.sCidadeOCR = tOcrCasoImport.sCidadeOCR
        objOcrCasoImport.sEstadoOCR = tOcrCasoImport.sEstadoOCR
        objOcrCasoImport.sPaisOCR = tOcrCasoImport.sPaisOCR
        objOcrCasoImport.sPrestador = tOcrCasoImport.sPrestador
        objOcrCasoImport.sCarater = tOcrCasoImport.sCarater
        objOcrCasoImport.sGrauSatisfacao = tOcrCasoImport.sGrauSatisfacao
        objOcrCasoImport.sTelefone = tOcrCasoImport.sTelefone

        objOcrCasos.colEventos.Add objOcrCasoImport
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208553
        
    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    TRVOcrCasos_Le_Eventos = SUCESSO

    Exit Function

Erro_TRVOcrCasos_Le_Eventos:

    TRVOcrCasos_Le_Eventos = gErr

    Select Case gErr

        Case 208550
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 208551 To 208553
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVOCRCASOIMPORT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208554)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function TRVLibOcrCasos_Le(ByVal objLibOcr As ClassTRVLibOcrAssist) As Long
 
Dim lErro As Long
Dim iIndice As Integer
Dim sSQL As String
Dim sWhere As String
Dim vsCodigo As Variant
Dim alComando(1 To 1) As Long
Dim objOcorrencia As ClassTRVOcrCasos
Dim vsCodigoDe As Variant
Dim vsCodigoAte As Variant
Dim vdtProgDe As Variant
Dim vdtProgAte As Variant
Dim vdtDe As Variant
Dim vdtAte As Variant
Dim sData As String

On Error GoTo Erro_TRVLibOcrCasos_Le
    
    vsCodigoDe = objLibOcr.sCodigoDe
    vsCodigoAte = objLibOcr.sCodigoAte
    vdtDe = objLibOcr.dtDataDe
    vdtAte = objLibOcr.dtDataAte
    vdtProgDe = objLibOcr.dtDataProgDe
    vdtProgAte = objLibOcr.dtDataProgAte
    vsCodigo = String(STRING_TRV_MAXIMO, 0)
 
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 208742
    Next
    
    If objLibOcr.iTipo = TRV_OCRCASOS_LIB_COBERTURA Then
        sSQL = "SELECT O.Codigo FROM TRVOcrCasos AS O WHERE O.NumIntDocTitPagCobertura = 0 AND O.CGAutorizadoPor <> 0 AND ((O.ValorAutorizadoAssistRS > 0) OR (O.ValorAutorizadoSeguroRS > 0 AND O.AnteciparPagtoSeguro = 1) OR (O.ValorAutoSegRespTrvRS > 0))"
        sData = "AND O.DataLimite"
    Else
        sSQL = "SELECT O.Codigo FROM TRVOcrCasos AS O WHERE O.NumIntDocTitPagProcesso = 0 AND O.Condenado <> 0 AND O.ValorCondenacao >0 AND O.DataFimProcesso <> {d '1822-09-07'} "
        sData = "AND O.DataFimProcesso"
    End If
 
    If vsCodigoDe <> "" Then
        sWhere = sWhere & "AND O.Codigo >= ? "
    End If
    If vsCodigoAte <> "" Then
        sWhere = sWhere & "AND O.Codigo <= ? "
    End If

    If vdtDe <> DATA_NULA Then
        sWhere = sWhere & sData & " >= ? "
    End If
    If vdtAte <> DATA_NULA Then
        sWhere = sWhere & sData & " <= ? "
    End If
    
    If objLibOcr.iTipo = TRV_OCRCASOS_LIB_COBERTURA Then
        If vdtProgDe <> DATA_NULA Then
            sWhere = sWhere & "AND O.DataProgFinanc >= ? "
        End If
        If vdtProgAte <> DATA_NULA Then
            sWhere = sWhere & "AND O.DataProgFinanc <= ? "
        End If
    End If
    
    sSQL = sSQL & sWhere
    
    sSQL = sSQL & "ORDER BY O.Codigo "
 
    lErro = Comando_PrepararInt(alComando(1), sSQL)
    If (lErro <> AD_SQL_SUCESSO) Then gError 208743
 
    lErro = Comando_BindVarInt(alComando(1), vsCodigo)
    If (lErro <> AD_SQL_SUCESSO) Then gError 208744
 
    If vsCodigoDe <> "" Then
        lErro = Comando_BindVarInt(alComando(1), vsCodigoDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError 208745
    End If
    
    If vsCodigoAte <> "" Then
        lErro = Comando_BindVarInt(alComando(1), vsCodigoAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError 208746
    End If
 
    If vdtDe <> DATA_NULA Then
        lErro = Comando_BindVarInt(alComando(1), vdtDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError 208747
    End If
    
    If vdtAte <> DATA_NULA Then
        lErro = Comando_BindVarInt(alComando(1), vdtAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError 208748
    End If
    
    If objLibOcr.iTipo = TRV_OCRCASOS_LIB_COBERTURA Then
        If vdtProgDe <> DATA_NULA Then
            lErro = Comando_BindVarInt(alComando(1), vdtProgDe)
            If (lErro <> AD_SQL_SUCESSO) Then gError 208749
        End If
        
        If vdtProgAte <> DATA_NULA Then
            lErro = Comando_BindVarInt(alComando(1), vdtProgAte)
            If (lErro <> AD_SQL_SUCESSO) Then gError 208750
        End If
    End If
    
    lErro = Comando_ExecutarInt(alComando(1))
    If (lErro <> AD_SQL_SUCESSO) Then gError 208751
 
    'Busca o primeiro registro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208752
        
    'Para cada Produto
    Do While lErro = SUCESSO
    
        Set objOcorrencia = New ClassTRVOcrCasos

        objOcorrencia.sCodigo = vsCodigo
        
        'Lê o produto
        lErro = TRVOcrCasos_Le(objOcorrencia)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
        
        objLibOcr.colOcorrenciais.Add objOcorrencia
     
        'Busca o próximo registro
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208753
 
    Loop
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
 
    TRVLibOcrCasos_Le = SUCESSO
 
    Exit Function
 
Erro_TRVLibOcrCasos_Le:
 
    TRVLibOcrCasos_Le = gErr
 
    Select Case gErr
 
        Case 208742
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 208743 To 208753
            Call Rotina_Erro(vbOKOnly, "ERRO_PREPARACAO_COMANDO_SQL", gErr, sSQL)
 
        Case ERRO_SEM_MENSAGEM
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208754)
 
    End Select
 
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
 
    Exit Function
 
End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function TRVOcrCasosPreRec_Le(ByVal objOcrPreRec As ClassTRVOcrCasosPreRec) As Long

Dim lErro As Long
Dim lComando As Long
Dim tOcrPreRec As typeTRVOcrCasosPreRec

On Error GoTo Erro_TRVOcrCasosPreRec_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 208800

    'Alocação de espaço no buffer
    tOcrPreRec.sDescricao = String(STRING_TRV_MAXIMO, 0)

    'Le a tabelaTRVOcrCasosPreRec
    lErro = Comando_Executar(lComando, "SELECT NumIntDocOcrCaso, Valor, Data, NumIntDocTitRecReembolso, NumFatTitRecReembolso, Descricao, " & _
                "DataPagto FROM TRVOcrCasosPreRec WHERE NumIntDocOcrCaso= ? ", _
                tOcrPreRec.lNumIntDocOcrCaso, tOcrPreRec.dValor, tOcrPreRec.dtData, tOcrPreRec.lNumIntDocTitRecReembolso, _
                tOcrPreRec.lNumFatTitRecReembolso, tOcrPreRec.sDescricao, tOcrPreRec.dtDataPagto, _
                objOcrPreRec.lNumIntDocOcrCaso)
    If lErro <> AD_SQL_SUCESSO Then gError 208801

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208802

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objOcrPreRec.lNumIntDocOcrCaso = tOcrPreRec.lNumIntDocOcrCaso
    objOcrPreRec.dValor = tOcrPreRec.dValor
    objOcrPreRec.dtData = tOcrPreRec.dtData
    objOcrPreRec.lNumIntDocTitRecReembolso = tOcrPreRec.lNumIntDocTitRecReembolso
    objOcrPreRec.lNumFatTitRecReembolso = tOcrPreRec.lNumFatTitRecReembolso
    objOcrPreRec.sDescricao = tOcrPreRec.sDescricao
    objOcrPreRec.dtDataPagto = tOcrPreRec.dtDataPagto

    'Fecha Comando
    Call Comando_Fechar(lComando)

    TRVOcrCasosPreRec_Le = SUCESSO

    Exit Function

Erro_TRVOcrCasosPreRec_Le:

    TRVOcrCasosPreRec_Le = gErr

    Select Case gErr

        Case 208800
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 208801, 208802
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVOCRCASOSPREREC", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208803)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function TRVOcrCasosPreRec_Le_Vinculados(ByVal lNumIntDoc As Long, ByVal colPR As Collection) As Long

Dim lErro As Long
Dim lComando As Long
Dim tOcrPreRec As typeTRVOcrCasosPreRec
Dim objOcrPreRec As ClassTRVOcrCasosPreRec

On Error GoTo Erro_TRVOcrCasosPreRec_Le_Vinculados

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 208804

    'Alocação de espaço no buffer
    tOcrPreRec.sDescricao = String(STRING_TRV_MAXIMO, 0)

    'Le a tabelaTRVOcrCasosPreRec
    lErro = Comando_Executar(lComando, "SELECT NumIntDocOcrCaso, Valor, Data, NumIntDocTitRecReembolso, NumFatTitRecReembolso, Descricao, " & _
                "DataPagto FROM TRVOcrCasosPreRec WHERE NumIntDocTitRecReembolso= ? ", _
                tOcrPreRec.lNumIntDocOcrCaso, tOcrPreRec.dValor, tOcrPreRec.dtData, tOcrPreRec.lNumIntDocTitRecReembolso, _
                tOcrPreRec.lNumFatTitRecReembolso, tOcrPreRec.sDescricao, tOcrPreRec.dtDataPagto, lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 208805

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208806

    'Sem Dados
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objOcrPreRec = New ClassTRVOcrCasosPreRec

        objOcrPreRec.lNumIntDocOcrCaso = tOcrPreRec.lNumIntDocOcrCaso
        objOcrPreRec.dValor = tOcrPreRec.dValor
        objOcrPreRec.dtData = tOcrPreRec.dtData
        objOcrPreRec.lNumIntDocTitRecReembolso = tOcrPreRec.lNumIntDocTitRecReembolso
        objOcrPreRec.lNumFatTitRecReembolso = tOcrPreRec.lNumFatTitRecReembolso
        objOcrPreRec.sDescricao = tOcrPreRec.sDescricao
        objOcrPreRec.dtDataPagto = tOcrPreRec.dtDataPagto
        
        colPR.Add objOcrPreRec
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208807
        
    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    TRVOcrCasosPreRec_Le_Vinculados = SUCESSO

    Exit Function

Erro_TRVOcrCasosPreRec_Le_Vinculados:

    TRVOcrCasosPreRec_Le_Vinculados = gErr

    Select Case gErr

        Case 208804
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 208805 To 208807
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVOCRCASOSPREREC", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208808)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function TRVProdLimites_Le(ByVal sProduto As String, ByVal colProdLim As Collection) As Long

Dim lErro As Long
Dim lComando As Long
Dim tProdLim As typeTRVProdLimites
Dim objProdLim As ClassTRVProdLimites

On Error GoTo Erro_TRVProdLimites_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 209082

    'Alocação de espaço no buffer
    tProdLim.sProduto = String(STRING_PRODUTO, 0)
    tProdLim.sDescProd = String(STRING_TRV_MAXIMO, 0)
    tProdLim.sDescricao1 = String(STRING_TRV_MAXIMO, 0)
    tProdLim.sDescricao2 = String(STRING_TRV_MAXIMO, 0)

    'Le a tabelaTRVProdLimites
    lErro = Comando_Executar(lComando, "SELECT Produto, CodServ, DescProd, ValorLimite, Moeda, Descricao1, " & _
                "Descricao2, Tipo FROM TRVProdLimites WHERE Produto= ? ORDER BY CodServ", _
                tProdLim.sProduto, tProdLim.lCodServ, tProdLim.sDescProd, tProdLim.dValorLimite, _
                tProdLim.iMoeda, tProdLim.sDescricao1, tProdLim.sDescricao2, tProdLim.iTipo, _
                sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 209083

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209084

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objProdLim = New ClassTRVProdLimites
        colProdLim.Add objProdLim

        objProdLim.sProduto = tProdLim.sProduto
        objProdLim.lCodServ = tProdLim.lCodServ
        objProdLim.sDescProd = tProdLim.sDescProd
        objProdLim.dValorLimite = tProdLim.dValorLimite
        objProdLim.iMoeda = tProdLim.iMoeda
        objProdLim.sDescricao1 = tProdLim.sDescricao1
        objProdLim.sDescricao2 = tProdLim.sDescricao2
        objProdLim.iTipo = tProdLim.iTipo
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209085
        
    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    TRVProdLimites_Le = SUCESSO

    Exit Function

Erro_TRVProdLimites_Le:

    TRVProdLimites_Le = gErr

    Select Case gErr

        Case 209082
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 209083 To 209085
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVPRODLIMITES", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209086)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function TRVOcrCasosGastosAdv_Le(ByVal objOcrCasos As ClassTRVOcrCasos) As Long

Dim lErro As Long
Dim lComando As Long
Dim tOcrCasosGastosAdv As typeTRVOcrCasosGastosAdv
Dim objOcrCasosGastosAdv As ClassTRVOcrCasosGastosAdv

On Error GoTo Erro_TRVOcrCasosGastosAdv_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 209121
    
    tOcrCasosGastosAdv.sDescricao = String(STRING_TRV_MAXIMO, 0)

    'Le a tabelaTRVOcrCasosGastosAdv
    lErro = Comando_Executar(lComando, "SELECT NumIntDocOcrCaso, Seq, Data, Valor, Descricao FROM TRVOcrCasosGastosAdv WHERE NumIntDocOcrCaso= ?  ORDER BY Seq", _
                tOcrCasosGastosAdv.lNumIntDocOcrCaso, tOcrCasosGastosAdv.iSeq, tOcrCasosGastosAdv.dtData, tOcrCasosGastosAdv.dValor, _
                tOcrCasosGastosAdv.sDescricao, objOcrCasos.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 209122

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209123

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objOcrCasosGastosAdv = New ClassTRVOcrCasosGastosAdv

        objOcrCasosGastosAdv.lNumIntDocOcrCaso = tOcrCasosGastosAdv.lNumIntDocOcrCaso
        objOcrCasosGastosAdv.iSeq = tOcrCasosGastosAdv.iSeq
        objOcrCasosGastosAdv.dtData = tOcrCasosGastosAdv.dtData
        objOcrCasosGastosAdv.dValor = tOcrCasosGastosAdv.dValor
        objOcrCasosGastosAdv.sDescricao = tOcrCasosGastosAdv.sDescricao

        objOcrCasos.colGastosAdvs.Add objOcrCasosGastosAdv
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209124
        
    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    TRVOcrCasosGastosAdv_Le = SUCESSO

    Exit Function

Erro_TRVOcrCasosGastosAdv_Le:

    TRVOcrCasosGastosAdv_Le = gErr

    Select Case gErr

        Case 209121
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 209122 To 209124
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVOCRCASOSGASTOSADV", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209125)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function TRVOcrCasosDocs_Le(ByVal objOcrCasos As ClassTRVOcrCasos) As Long

Dim lErro As Long
Dim lComando As Long
Dim tOcrCasosDocs As typeTRVOcrCasosDocs
Dim objOcrCasosDocs As ClassTRVOcrCasosDocs

On Error GoTo Erro_TRVOcrCasosDocs_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 209126
    
    tOcrCasosDocs.sDescricao = String(STRING_TRV_MAXIMO, 0)
    tOcrCasosDocs.sObservacao = String(STRING_TRV_MAXIMO, 0)

    'Le a tabelaTRVOcrCasosDocs
    lErro = Comando_Executar(lComando, "SELECT N.NumIntDocOcrCaso, N.Seq, D.Descricao, N.Observacao, N.NecessSist, N.NecessUsu, N.Recebido, N.CodigoDoc  FROM TRVOcrCasosDocs AS N, TRVAssistDocs As D WHERE N.CodigoDoc = D.Codigo AND N.NumIntDocOcrCaso= ? ORDER BY Seq", _
                tOcrCasosDocs.lNumIntDocOcrCaso, tOcrCasosDocs.iSeq, tOcrCasosDocs.sDescricao, tOcrCasosDocs.sObservacao, tOcrCasosDocs.iNecessSist, tOcrCasosDocs.iNecessUsu, _
                tOcrCasosDocs.iRecebido, tOcrCasosDocs.lCodigoDoc, objOcrCasos.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 209127

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209128

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objOcrCasosDocs = New ClassTRVOcrCasosDocs

        objOcrCasosDocs.lNumIntDocOcrCaso = tOcrCasosDocs.lNumIntDocOcrCaso
        objOcrCasosDocs.iSeq = tOcrCasosDocs.iSeq
        objOcrCasosDocs.iNecessSist = tOcrCasosDocs.iNecessSist
        objOcrCasosDocs.iNecessUsu = tOcrCasosDocs.iNecessUsu
        objOcrCasosDocs.iRecebido = tOcrCasosDocs.iRecebido
        objOcrCasosDocs.lCodigoDoc = tOcrCasosDocs.lCodigoDoc
        objOcrCasosDocs.sDescricao = tOcrCasosDocs.sDescricao
        objOcrCasosDocs.sObservacao = tOcrCasosDocs.sObservacao

        objOcrCasos.colDocs.Add objOcrCasosDocs
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209129
        
    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    TRVOcrCasosDocs_Le = SUCESSO

    Exit Function

Erro_TRVOcrCasosDocs_Le:

    TRVOcrCasosDocs_Le = gErr

    Select Case gErr

        Case 209126
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 209127 To 209129
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVOCRCASOSDOCS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209130)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function TRVOcrCasosDocs_Carrega(ByVal objOcrCaso As ClassTRVOcrCasos) As Long

Dim lErro As Long, lComando As Long
Dim objOcrCasosSrv As ClassTRVOcrCasosSrv
Dim objOcrCasosDoc As ClassTRVOcrCasosDocs
Dim objOcrCasosDocAux As ClassTRVOcrCasosDocs
Dim objCodigoDescricao As AdmCodigoNome
Dim colCodigoDescricao As New AdmColCodigoNome
Dim iSeq As Integer, bAchou As Boolean, lCodigoServ As Long
Dim colSaida As New Collection, colCampos As New Collection, iIndice As Integer

On Error GoTo Erro_TRVOcrCasosDocs_Carrega

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 209131

    lErro = CF("Cod_Nomes_Le", "TRVAssistDocs", "Codigo", "Descricao", STRING_TRV_MAXIMO, colCodigoDescricao)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    iSeq = 0
    For Each objCodigoDescricao In colCodigoDescricao
    
        bAchou = False
        For Each objOcrCasosDocAux In objOcrCaso.colDocs
            If objOcrCasosDocAux.lCodigoDoc = objCodigoDescricao.iCodigo Then
                objOcrCasosDocAux.sDescricao = objCodigoDescricao.sNome
                bAchou = True
                Exit For
            End If
        Next
    
        If Not bAchou Then
            iSeq = iSeq + 1
    
            Set objOcrCasosDoc = New ClassTRVOcrCasosDocs
            
            objOcrCasosDoc.sDescricao = objCodigoDescricao.sNome
            objOcrCasosDoc.lCodigoDoc = objCodigoDescricao.iCodigo
            
            For Each objOcrCasosSrv In objOcrCaso.colCoberturas
                If objOcrCasosSrv.iSolicitado = MARCADO Then
                    
                    'Le os documentos necessários para cada serviço solicitado
                    lErro = Comando_Executar(lComando, "SELECT CodigoServ FROM TRVAssistDocsSrv WHERE CodigoDoc = ? AND CodigoServ In (?,?)", lCodigoServ, objOcrCasosDoc.lCodigoDoc, objOcrCasosSrv.lCodigoServ, 0)
                    If lErro <> AD_SQL_SUCESSO Then gError 209132
                
                    'Busca Primeiro
                    lErro = Comando_BuscarPrimeiro(lComando)
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209133
                    
                    If lErro = AD_SQL_SUCESSO Then
                        objOcrCasosDoc.iNecessSist = MARCADO
                        objOcrCasosDoc.iNecessUsu = MARCADO
                    End If
                    
                End If
            Next
    
            objOcrCaso.colDocs.Add objOcrCasosDoc
        End If
    
    Next
    
    colCampos.Add "lCodigoDoc"
    
    lErro = Ordena_Colecao(objOcrCaso.colDocs, colSaida, colCampos)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    For iIndice = objOcrCaso.colDocs.Count To 1 Step -1
        objOcrCaso.colDocs.Remove iIndice
    Next
    
    'Coloca primeiro os documentos necessário ordenados pelo código
    For Each objOcrCasosDoc In colSaida
        If objOcrCasosDoc.iNecessSist = MARCADO Or objOcrCasosDoc.iNecessUsu = MARCADO Then
            objOcrCaso.colDocs.Add objOcrCasosDoc
        End If
    Next
    'Depois os outros ordenados pelo código
    For Each objOcrCasosDoc In colSaida
        If objOcrCasosDoc.iNecessSist <> MARCADO And objOcrCasosDoc.iNecessUsu <> MARCADO Then
            objOcrCaso.colDocs.Add objOcrCasosDoc
        End If
    Next

    'Fecha Comando
    Call Comando_Fechar(lComando)

    TRVOcrCasosDocs_Carrega = SUCESSO

    Exit Function

Erro_TRVOcrCasosDocs_Carrega:

    TRVOcrCasosDocs_Carrega = gErr

    Select Case gErr

        Case 209131
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 209132, 209133
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVASSISTDOCSSRV", gErr)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209134)

    End Select
    
    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function TRVOcrCasosSrv_Carrega(ByVal objOcrCaso As ClassTRVOcrCasos) As Long

Dim lErro As Long
Dim objOcrCasosSrv As ClassTRVOcrCasosSrv
Dim objOcrCasosSrvAux As ClassTRVOcrCasosSrv
Dim objProdLin As ClassTRVProdLimites
Dim colProdLim As New Collection
Dim iSeq As Integer, bAchou As Boolean

On Error GoTo Erro_TRVOcrCasosSrv_Carrega

    If Len(Trim(objOcrCaso.sProduto)) > 0 Then

        lErro = CF("TRVProdLimites_Le", objOcrCaso.sProduto, colProdLim)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        iSeq = 0
        For Each objProdLin In colProdLim
        
            bAchou = False
            For Each objOcrCasosSrvAux In objOcrCaso.colCoberturas
                If objOcrCasosSrvAux.lCodigoServ = objProdLin.lCodServ Then
                    bAchou = True
                    Exit For
                End If
            Next
        
            If Not bAchou Then
                iSeq = iSeq + 1
        
                Set objOcrCasosSrv = New ClassTRVOcrCasosSrv
                
                objOcrCasosSrv.dValorLimite = objProdLin.dValorLimite
                objOcrCasosSrv.iMoeda = objProdLin.iMoeda
                objOcrCasosSrv.sDescricao = objProdLin.sDescricao1
                objOcrCasosSrv.iTipo = objProdLin.iTipo
                objOcrCasosSrv.iSeq = iSeq
                objOcrCasosSrv.lCodigoServ = objProdLin.lCodServ
        
                objOcrCaso.colCoberturas.Add objOcrCasosSrv
            End If
        
        Next
        
    End If

    TRVOcrCasosSrv_Carrega = SUCESSO

    Exit Function

Erro_TRVOcrCasosSrv_Carrega:

    TRVOcrCasosSrv_Carrega = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209087)

    End Select

    Exit Function

End Function

Public Function TRVAssistVlrAuto_Le(ByVal dtData As Date, dValor As Double, sTextoUtil As String) As Long

Dim lErro As Long, lComando As Long
Dim objFeriado As New ClassFeriado

On Error GoTo Erro_TRVAssistVlrAuto_Le

    dValor = 0
    sTextoUtil = ""

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 209142
                   
    'Le os documentos necessários para cada serviço solicitado
    lErro = Comando_Executar(lComando, "SELECT Valor FROM TRVAssistVlrAuto WHERE Data = ? ", dValor, dtData)
    If lErro <> AD_SQL_SUCESSO Then gError 209143

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209144
    
    If Weekday(dtData) = vbSunday Then
        sTextoUtil = "Domingo"
    ElseIf Weekday(dtData) = vbSaturday Then
        sTextoUtil = "Sábado"
    Else
        objFeriado.dtData = dtData
        objFeriado.iFilialEmpresa = giFilialEmpresa
    
        lErro = CF("Feriado_Le", objFeriado)
        If lErro <> SUCESSO And lErro <> 43379 Then gError ERRO_SEM_MENSAGEM
    
        If lErro = SUCESSO Then sTextoUtil = "Feriado"
    End If
    
    'Fecha Comando
    Call Comando_Fechar(lComando)

    TRVAssistVlrAuto_Le = SUCESSO

    Exit Function

Erro_TRVAssistVlrAuto_Le:

    TRVAssistVlrAuto_Le = gErr

    Select Case gErr

        Case 209142
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 209143, 209144
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVASSISTVLRAUTO", gErr)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209145)

    End Select
    
    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Private Function TRVFaturamento_Obtem_Cliente(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal objFATInfoVou As ClassTRVFATInfoVoucher) As Long

Dim lErro As Long
Dim objVoucherInfo As New ClassTRVVoucherInfo

On Error GoTo Erro_TRVFaturamento_Obtem_Cliente

    'O cliente está configurado para faturar para o passageiro
    If objFATInfoVou.iUmVouPorFat = 2 Then
    
        'Se for voucher, ocr ou nvl
        If (objFATInfoVou.sTipoDoc = TRV_TIPODOC_VOU_TEXTO Or objFATInfoVou.sTipoDoc = TRV_TIPODOC_OCR_TEXTO Or objFATInfoVou.sTipoDoc = TRV_TIPODOC_NVL_TEXTO) Then
                
            'Se o passageiro ainda não foi cadastrado como cliente -> Cadastra
            If objFATInfoVou.lCliPassageiro = 0 Then
                
                objVoucherInfo.lNumVou = objFATInfoVou.lNumVou
                objVoucherInfo.sSerie = objFATInfoVou.sSerie
                objVoucherInfo.sTipo = objFATInfoVou.sTipoVou
                
                lErro = CF("TRVVoucherInfoSigav_Le", objVoucherInfo, lComando2)
                If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
                                    
                'CADASTRAR O PASSAGEIRO E ACERTA A INFORMAÇÃO EM VOUCHERINFO
                lErro = CF("TRV_VoucherInfo_Cadastra_Cliente", objVoucherInfo, objFATInfoVou.iCondPagto)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                objFATInfoVou.lCliPassageiro = objVoucherInfo.lCliPassageiro
                
            End If
            
            objFATInfoVou.lCliente = objFATInfoVou.lCliPassageiro
            objFATInfoVou.sNome = objFATInfoVou.sPassageiroNome & " " & objFATInfoVou.sPassageiroSobreNome
        
        End If
        
    End If
 
    TRVFaturamento_Obtem_Cliente = SUCESSO

    Exit Function

Erro_TRVFaturamento_Obtem_Cliente:

    TRVFaturamento_Obtem_Cliente = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192739)

    End Select
    
    Exit Function

End Function

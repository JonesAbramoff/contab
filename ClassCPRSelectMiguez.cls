VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCPRSelectMiguez"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function TituloPagarRateio_Le(ByVal objTituloPagar As ClassTituloPagar, ByVal colTituloPagRateio As Collection) As Long
'Lê os Rateioos do Titulo cujo NumIntDoc foi passado em objTituloPagar

Dim lErro As Long
Dim lComando As Long
Dim tTituloPagRateio As typeTituloPagRateio
Dim objTituloPagRateio As ClassTituloPagRateio

On Error GoTo Erro_TituloPagarRateio_Le

    'Abre comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 129042
    
    tTituloPagRateio.sDossie = String(STRING_DOSSIERATEIO, 0)
    tTituloPagRateio.sReferencia = String(STRING_REFERENCIARATEIO, 0)
    
    'Seleciona Rateios vinculados ao Título
    lErro = Comando_Executar(lComando, "SELECT NumIntDocPag, Seq, NumIntDocRec, Valor, Cliente, Historico, Referencia, Dossie, Cobrar, ND, Data_Ger_ND FROM TitulosPagRateio WHERE NumIntDocPag = ? ORDER BY Seq ", tTituloPagRateio.lNumIntDocPag, tTituloPagRateio.iSeq, tTituloPagRateio.lNumIntDocRec, tTituloPagRateio.dValor, tTituloPagRateio.lCliente, tTituloPagRateio.lHistorico, tTituloPagRateio.sReferencia, tTituloPagRateio.sDossie, tTituloPagRateio.iCobrar, tTituloPagRateio.lND, tTituloPagRateio.dtDataGerND, objTituloPagar.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 129043

    'Tenta ler o primeiro item do Rateio
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 129044
        
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objTituloPagRateio = New ClassTituloPagRateio
    
        'Adiciona Rateio na coleção
        objTituloPagRateio.dValor = tTituloPagRateio.dValor
        objTituloPagRateio.iCobrar = tTituloPagRateio.iCobrar
        objTituloPagRateio.iSeq = tTituloPagRateio.iSeq
        objTituloPagRateio.lCliente = tTituloPagRateio.lCliente
        objTituloPagRateio.lHistorico = tTituloPagRateio.lHistorico
        objTituloPagRateio.lNumIntDocPag = tTituloPagRateio.lNumIntDocPag
        objTituloPagRateio.lNumIntDocRec = tTituloPagRateio.lNumIntDocRec
        objTituloPagRateio.sDossie = tTituloPagRateio.sDossie
        objTituloPagRateio.sReferencia = tTituloPagRateio.sReferencia
        objTituloPagRateio.dtDataGerND = tTituloPagRateio.dtDataGerND
        objTituloPagRateio.lND = tTituloPagRateio.lND
        
        colTituloPagRateio.Add objTituloPagRateio
        
        'Vai para o próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 129046
    
    Loop
    
    'Fecha comando
    Call Comando_Fechar(lComando)
    
    TituloPagarRateio_Le = SUCESSO
    
    Exit Function
    
Erro_TituloPagarRateio_Le:

    TituloPagarRateio_Le = gErr
    
    Select Case gErr
    
        Case 129042
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 129043, 129044, 129045, 129046
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_TITULOSPAG", gErr, objTituloPagar.lNumIntDoc)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
    
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function TituloRec_BuscaPrevisao(ByVal objTituloReceber As ClassTituloReceber, ByVal objTitRecPrevisao As ClassTituloReceber) As Long
'preenche objTitRecPrevisao com numintdoc de titulo com STATUS_PREVISAO de mesmo cliente-filial de objTituloReceber com mesmo mes de vencimento de parcela

Dim lErro As Long, lComando As Long, lNumIntDoc As Long, sSiglaDocumento As String

On Error GoTo Erro_TituloRec_BuscaPrevisao

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 106976
    
    sSiglaDocumento = String(STRING_SIGLA_DOCUMENTO, 0)
    
    lErro = Comando_Executar(lComando, "SELECT TitulosRec.SiglaDocumento, TitulosRec.NumIntDoc FROM TitulosRec, ParcelasRec WHERE TitulosRec.FilialEmpresa = ? AND TitulosRec.Status = ? AND TitulosRec.NumIntDoc = ParcelasRec.NumIntTitulo AND Cliente = ? AND Filial = ? AND {fn Month(ParcelasRec.DataVencimentoReal)} = ?  AND {fn Year(ParcelasRec.DataVencimentoReal)} = ?", _
        sSiglaDocumento, lNumIntDoc, objTituloReceber.iFilialEmpresa, STATUS_PREVISAO, objTituloReceber.lCliente, objTituloReceber.iFilial, Month(objTituloReceber.colParcelaReceber(1).dtDataVencimentoReal), Year(objTituloReceber.colParcelaReceber(1).dtDataVencimentoReal))
    If lErro <> AD_SQL_SUCESSO Then gError 106977

    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106798
    
    If lErro = AD_SQL_SEM_DADOS Then gError 106975
    
    objTitRecPrevisao.iFilialEmpresa = objTituloReceber.iFilialEmpresa
    objTitRecPrevisao.lNumIntDoc = lNumIntDoc
    objTitRecPrevisao.sSiglaDocumento = sSiglaDocumento
    
    Call Comando_Fechar(lComando)
    
    TituloRec_BuscaPrevisao = SUCESSO
     
    Exit Function
    
Erro_TituloRec_BuscaPrevisao:

    TituloRec_BuscaPrevisao = gErr
     
    Select Case gErr
          
        Case 106975
        
        Case 106976
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 106978
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PREVISAO_TITREC", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function NFiscal_BuscaPrevisaoRec(ByVal objNFiscal As ClassNFiscal, ByVal objTitRecPrevisao As ClassTituloReceber) As Long
'preenche objTitRecPrevisao com numintdoc de titulo com STATUS_PREVISAO de mesmo cliente-filial de objNfiscal com mesmo mes de vencimento de parcela

Dim lErro As Long, lComando As Long, lNumIntDoc As Long, sSiglaDocumento As String

On Error GoTo Erro_NFiscal_BuscaPrevisaoRec

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 106976
    
    sSiglaDocumento = String(STRING_SIGLA_DOCUMENTO, 0)
    
    lErro = Comando_Executar(lComando, "SELECT TitulosRec.SiglaDocumento, TitulosRec.NumIntDoc FROM TitulosRec, ParcelasRec WHERE TitulosRec.FilialEmpresa = ? AND TitulosRec.Status = ? AND TitulosRec.NumIntDoc = ParcelasRec.NumIntTitulo AND Cliente = ? AND Filial = ? AND {fn Month(ParcelasRec.DataVencimentoReal)} = ?  AND {fn Year(ParcelasRec.DataVencimentoReal)} = ?", _
        sSiglaDocumento, lNumIntDoc, objNFiscal.iFilialEmpresa, STATUS_PREVISAO, objNFiscal.lCliente, objNFiscal.iFilialCli, Month(objNFiscal.colParcelaReceber(1).dtDataVencimentoReal), Year(objNFiscal.colParcelaReceber(1).dtDataVencimentoReal))
    If lErro <> AD_SQL_SUCESSO Then gError 106977

    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106798
    
    If lErro = AD_SQL_SEM_DADOS Then gError 106975
    
    objTitRecPrevisao.iFilialEmpresa = objNFiscal.iFilialEmpresa
    objTitRecPrevisao.lNumIntDoc = lNumIntDoc
    objTitRecPrevisao.sSiglaDocumento = sSiglaDocumento
    
    Call Comando_Fechar(lComando)
    
    NFiscal_BuscaPrevisaoRec = SUCESSO
     
    Exit Function
    
Erro_NFiscal_BuscaPrevisaoRec:

    NFiscal_BuscaPrevisaoRec = gErr
     
    Select Case gErr
          
        Case 106975
        
        Case 106976
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 106978
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PREVISAO_TITREC", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassSRVSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function Comando_BindVarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_BindVar" (ByVal lComando As Long, lpVar As Variant) As Long
Private Declare Function Comando_PrepararInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Preparar" (ByVal lComando As Long, ByVal lpSQLStmt As String) As Long
Private Declare Function Comando_ExecutarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Executar" (ByVal lComando As Long) As Long

Public Function SolicitacaoSRV_Le(objSolicSRV As ClassSolicSRV) As Long

Dim lComando As Long
Dim lErro As Long
Dim tSolicSRV As typeSolicSRV

On Error GoTo Erro_SolicitacaoSRV_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 183255
    
    tSolicSRV.sObs = String(STRING_MAXIMO, 0)
        
    lErro = Comando_Executar(lComando, "SELECT Data, Hora, Cliente, FilialCliente, NumIntDoc, Vendedor, Atendente, ClienteBenef, FilialClienteBenef, Prazo, PrazoTipo, Obs, DataEntrega, Tipo, Fase FROM SolicitacaoSRV WHERE FilialEmpresa = ? AND Codigo = ?", _
    tSolicSRV.dtData, tSolicSRV.dHora, tSolicSRV.lCliente, tSolicSRV.iFilial, tSolicSRV.lNumIntDoc, tSolicSRV.iVendedor, tSolicSRV.iAtendente, tSolicSRV.lClienteBenef, tSolicSRV.iFilialClienteBenef, tSolicSRV.iPrazo, tSolicSRV.iPrazoTipo, tSolicSRV.sObs, tSolicSRV.dtDataEntrega, tSolicSRV.lTipo, tSolicSRV.lFase, objSolicSRV.iFilialEmpresa, objSolicSRV.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 183256
        
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 183257
        
    If lErro = AD_SQL_SEM_DADOS Then gError 183258

    objSolicSRV.dtData = tSolicSRV.dtData
    objSolicSRV.dtHora = tSolicSRV.dHora
    objSolicSRV.lCliente = tSolicSRV.lCliente
    objSolicSRV.iFilial = tSolicSRV.iFilial
    objSolicSRV.lNumIntDoc = tSolicSRV.lNumIntDoc
    objSolicSRV.iVendedor = tSolicSRV.iVendedor
    objSolicSRV.iAtendente = tSolicSRV.iAtendente
    objSolicSRV.lClienteBenef = tSolicSRV.lClienteBenef
    objSolicSRV.iFilialClienteBenef = tSolicSRV.iFilialClienteBenef
    objSolicSRV.iPrazo = tSolicSRV.iPrazo
    objSolicSRV.iPrazoTipo = tSolicSRV.iPrazoTipo
    objSolicSRV.sObs = tSolicSRV.sObs
    objSolicSRV.dtDataEntrega = tSolicSRV.dtDataEntrega
    objSolicSRV.lTipo = tSolicSRV.lTipo
    objSolicSRV.lFase = tSolicSRV.lFase

    lErro = CF("ItensSolicSRV_Le", objSolicSRV)
    If lErro <> SUCESSO Then gError 183265

    Call Comando_Fechar(lComando)
    
    SolicitacaoSRV_Le = SUCESSO
    
    Exit Function

Erro_SolicitacaoSRV_Le:
    
    SolicitacaoSRV_Le = gErr
        
    Select Case gErr
    
        Case 183255
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 183256, 183257
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SOLICITACAOSRV", gErr, objSolicSRV.iFilialEmpresa, objSolicSRV.lCodigo)
        
        Case 183258, 183265
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 183259)
            
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function ItensSolicSRV_Le(objSolicSRV As ClassSolicSRV) As Long

Dim lComando As Long
Dim lErro As Long
Dim tItensSolicSRV As typeItensSolicSRV
Dim objItensSolicSrv As ClassItensSolicSRV

On Error GoTo Erro_ItensSolicSRV_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 183260
        
    tItensSolicSRV.sProduto = String(STRING_PRODUTO, 0)
    tItensSolicSRV.sProdutoDesc = String(STRING_PRODUTO_DESCRICAO, 0)
    tItensSolicSRV.sServico = String(STRING_PRODUTO, 0)
    tItensSolicSRV.sServicoDesc = String(STRING_PRODUTO_DESCRICAO, 0)
    tItensSolicSRV.sLote = String(STRING_LOTE_RASTREAMENTO, 0)
    tItensSolicSRV.sSolicitacao = String(STRING_SOLICITACAO, 0)
    tItensSolicSRV.sContrato = String(STRING_CONTRATO_CODIGO, 0)
    tItensSolicSRV.sUM = String(STRING_UM_SIGLA, 0)
    tItensSolicSRV.sReparo = String(STRING_MAXIMO, 0)
        
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Produto, Produtos.Descricao, DataVenda, Servico, Servicos.Descricao, UM, Quantidade, Lote, FilialOP, Solicitacao, Garantia, Contrato, StatusItem, Reparo, DataBaixa FROM ItensSolicSRV, Produtos, Produtos As Servicos WHERE NumIntSolicSRV = ? AND ItensSolicSRV.Produto = Produtos.Codigo AND ItensSolicSRV.Servico = Servicos.Codigo", _
    tItensSolicSRV.lNumIntDoc, tItensSolicSRV.sProduto, tItensSolicSRV.sProdutoDesc, tItensSolicSRV.dtDataVenda, tItensSolicSRV.sServico, tItensSolicSRV.sServicoDesc, tItensSolicSRV.sUM, tItensSolicSRV.dQuantidade, tItensSolicSRV.sLote, tItensSolicSRV.iFilialOP, tItensSolicSRV.sSolicitacao, tItensSolicSRV.lGarantia, tItensSolicSRV.sContrato, tItensSolicSRV.iStatusItem, tItensSolicSRV.sReparo, tItensSolicSRV.dtDataBaixa, objSolicSRV.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 183261
        
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 183262
        
    Do While lErro = AD_SQL_SUCESSO
    
        Set objItensSolicSrv = New ClassItensSolicSRV
    
        objItensSolicSrv.lNumIntDoc = tItensSolicSRV.lNumIntDoc
        objItensSolicSrv.lNumIntSolicSRV = objSolicSRV.lNumIntDoc
        objItensSolicSrv.sProduto = tItensSolicSRV.sProduto
        objItensSolicSrv.sProdutoDesc = tItensSolicSRV.sProdutoDesc
        objItensSolicSrv.dtDataVenda = tItensSolicSRV.dtDataVenda
        objItensSolicSrv.sServico = tItensSolicSRV.sServico
        objItensSolicSrv.sServicoDesc = tItensSolicSRV.sServicoDesc
        objItensSolicSrv.sUM = tItensSolicSRV.sUM
        objItensSolicSrv.dQuantidade = tItensSolicSRV.dQuantidade
        objItensSolicSrv.sLote = tItensSolicSRV.sLote
        objItensSolicSrv.iFilialOP = tItensSolicSRV.iFilialOP
        objItensSolicSrv.sSolicitacao = tItensSolicSRV.sSolicitacao
        objItensSolicSrv.sReparo = tItensSolicSRV.sReparo
        objItensSolicSrv.lGarantia = tItensSolicSRV.lGarantia
        objItensSolicSrv.sContrato = tItensSolicSRV.sContrato
        objItensSolicSrv.iStatusItem = tItensSolicSRV.iStatusItem
        objItensSolicSrv.dtDataBaixa = tItensSolicSRV.dtDataBaixa

        objSolicSRV.colItens.Add objItensSolicSrv

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 183263

    Loop

    Call Comando_Fechar(lComando)
    
    ItensSolicSRV_Le = SUCESSO
    
    Exit Function

Erro_ItensSolicSRV_Le:
    
    ItensSolicSRV_Le = gErr
        
    Select Case gErr
    
        Case 183260
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 183261 To 183263
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSSOLICSRV", gErr, objSolicSRV.lNumIntDoc)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 183264)
            
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function SRVConfig_Le_Configs(ByVal colConfigs As Collection) As Long
'recupera da tabela SRVConfig uma colecao de registros

Dim lErro As Long
Dim objSRVConfig As ClassSRVConfig

On Error GoTo Erro_SRVConfig_Le_Configs
    
    'para cada elemento da colecao
    For Each objSRVConfig In colConfigs
    
        lErro = SRVConfig_Le_Batch(objSRVConfig)
        If lErro <> SUCESSO And lErro <> 183394 Then gError 183388
        
        If lErro <> SUCESSO Then gError 183389
        
    Next
    
    SRVConfig_Le_Configs = SUCESSO
    
    Exit Function
    
Erro_SRVConfig_Le_Configs:

    SRVConfig_Le_Configs = gErr
    
    Select Case gErr
    
        Case 183388
        
        Case 183389
            Call Rotina_Erro(vbOKOnly, "ERRO_SRVCONFIG_INEXISTENTE", Err, objSRVConfig.sCodigo, objSRVConfig.iFilialEmpresa)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 183390)
            
    End Select
    
    Exit Function
    
End Function

Private Function SRVConfig_Le_Batch(ByVal objSRVConfig As ClassSRVConfig) As Long
'le registro em SRVConfig.

Dim lErro As Long
Dim tSRVConfig As typeSRVConfig
Dim lComando As Long

On Error GoTo Erro_SRVConfig_Le_Batch

    'abrir comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 183391

    With tSRVConfig
        
        'inicializar buffers
        .sDescricao = String(STRING_SRVCONFIG_DESCRICAO, 0)
        .sConteudo = String(STRING_SRVCONFIG_CONTEUDO, 0)
        
        'ler registo
        lErro = Comando_Executar(lComando, "SELECT Descricao, Tipo, Conteudo FROM SRVConfig WHERE Codigo = ? AND FilialEmpresa = ?", .sDescricao, .iTipo, .sConteudo, objSRVConfig.sCodigo, objSRVConfig.iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 183392
        
    End With
    
    'Lê o primeiro registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 183393
    
    'Se não encontrou o registro
    If lErro = AD_SQL_SEM_DADOS Then gError 183394
    
    objSRVConfig.sDescricao = tSRVConfig.sDescricao
    objSRVConfig.iTipo = tSRVConfig.iTipo
    objSRVConfig.sConteudo = tSRVConfig.sConteudo
    
    Call Comando_Fechar(lComando)
    
    SRVConfig_Le_Batch = SUCESSO
     
    Exit Function
    
Erro_SRVConfig_Le_Batch:

    SRVConfig_Le_Batch = gErr
     
    Select Case gErr
     
        Case 183391
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
     
        Case 183392, 183393
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SRVCONFIG", gErr, objSRVConfig.sCodigo, objSRVConfig.iFilialEmpresa)
        
        Case 183394
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183395)
     
    End Select
     
    Call Comando_Fechar(lComando)
     
    Exit Function

End Function

Public Function Garantia_Le_Lote(objGarantia As ClassGarantia) As Long

Dim lComando As Long
Dim lErro As Long
Dim tGarantia As typeGarantia

On Error GoTo Erro_Garantia_Le_Lote

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 183442
        
    tGarantia.sProduto = String(STRING_PRODUTO, 0)
    tGarantia.sLote = String(STRING_LOTE_RASTREAMENTO, 0)
    tGarantia.sSerie = String(STRING_SERIE, 0)
        
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Codigo, FilialEmpresa, Garantia.Produto, DataVenda, Quantidade, Lote, FilialOP, Serie, NumNotaFiscal, TipoGarantia, GarantiaTotal, GarantiaTotalPrazo, Fornecedor, FilialFornecedor FROM Garantia, GarantiaProduto WHERE FilialEmpresa = ? AND Garantia.Produto = ? AND DataVenda <= ? AND Lote = ? AND Garantia.NumIntDoc = GarantiaProduto.NumIntGarantia AND DataVenda + Prazo >= ? AND GarantiaProduto.Produto = ?", _
    tGarantia.lNumIntDoc, tGarantia.lCodigo, tGarantia.iFilialEmpresa, tGarantia.sProduto, tGarantia.dtDataVenda, tGarantia.dQuantidade, tGarantia.sLote, tGarantia.iFilialOP, tGarantia.sSerie, tGarantia.lNumNotaFiscal, tGarantia.lTipoGarantia, tGarantia.iGarantiaTotal, tGarantia.iGarantiaTotalPrazo, tGarantia.lFornecedor, tGarantia.iFilialFornecedor, objGarantia.iFilialEmpresa, objGarantia.sProduto, gdtDataHoje, objGarantia.sLote, gdtDataHoje, objGarantia.sServico)
    If lErro <> AD_SQL_SUCESSO Then gError 183443
        
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 183444
        
    If lErro = AD_SQL_SEM_DADOS Then
    
        lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Codigo, FilialEmpresa, Garantia.Produto, DataVenda, Quantidade, Lote, FilialOP, Serie, NumNotaFiscal, TipoGarantia, GarantiaTotal, GarantiaTotalPrazo FROM Garantia WHERE FilialEmpresa = ? AND Garantia.Produto = ? AND DataVenda <= ? AND Lote = ? AND DataVenda + GarantiaTotalPrazo >= ? AND Garantia.GarantiaTotal = ? AND NOT EXISTS (SELECT * FROM GarantiaProduto WHERE GarantiaProduto.NumIntGarantia = Garantia.NumIntDoc AND GarantiaProduto.Produto = ?)", _
        tGarantia.lNumIntDoc, tGarantia.lCodigo, tGarantia.iFilialEmpresa, tGarantia.sProduto, tGarantia.dtDataVenda, tGarantia.dQuantidade, tGarantia.sLote, tGarantia.iFilialOP, tGarantia.sSerie, tGarantia.lNumNotaFiscal, tGarantia.lTipoGarantia, tGarantia.iGarantiaTotal, tGarantia.iGarantiaTotalPrazo, objGarantia.iFilialEmpresa, objGarantia.sProduto, gdtDataHoje, objGarantia.sLote, gdtDataHoje, GARANTIA_TOTAL, objGarantia.sServico)
        If lErro <> AD_SQL_SUCESSO Then gError 183486
    
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 183487
        
        If lErro = AD_SQL_SEM_DADOS Then gError 183445

    End If

    objGarantia.lNumIntDoc = tGarantia.lNumIntDoc
    objGarantia.lCodigo = tGarantia.lCodigo
    objGarantia.iFilialEmpresa = tGarantia.iFilialEmpresa
    objGarantia.sProduto = tGarantia.sProduto
    objGarantia.dtDataVenda = tGarantia.dtDataVenda
    objGarantia.dQuantidade = tGarantia.dQuantidade
    objGarantia.sLote = tGarantia.sLote
    objGarantia.iFilialOP = tGarantia.iFilialOP
    objGarantia.sSerie = tGarantia.sSerie
    objGarantia.lNumNotaFiscal = tGarantia.lNumNotaFiscal
    objGarantia.lTipoGarantia = tGarantia.lTipoGarantia
    objGarantia.iGarantiaTotal = tGarantia.iGarantiaTotal
    objGarantia.iGarantiaTotalPrazo = tGarantia.iGarantiaTotalPrazo
    objGarantia.lFornecedor = tGarantia.lFornecedor
    objGarantia.iFilialFornecedor = tGarantia.iFilialFornecedor

    lErro = CF("GarantiaProduto_Le", objGarantia)
    If lErro <> SUCESSO Then gError 183447

    Call Comando_Fechar(lComando)
    
    Garantia_Le_Lote = SUCESSO
    
    Exit Function

Erro_Garantia_Le_Lote:
    
    Garantia_Le_Lote = gErr
        
    Select Case gErr
    
        Case 183442
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 183443, 183444, 183486, 183687
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_GARANTIA_LOTE", gErr, objGarantia.iFilialEmpresa, objGarantia.sProduto, objGarantia.sLote)
        
        Case 183445, 183447
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 183446)
            
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function GarantiaNumSerie_Le(objGarantia As ClassGarantia) As Long

Dim lComando As Long
Dim lErro As Long
Dim tGarantiaNumSerie As typeGarantiaNumSerie
Dim objGarantiaNumSerie As ClassGarantiaNumSerie

On Error GoTo Erro_GarantiaNumSerie_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 183447
        
    tGarantiaNumSerie.sNumSerie = String(STRING_LOTE_RASTREAMENTO, 0)
        
    lErro = Comando_Executar(lComando, "SELECT NumIntGarantia, NumSerie FROM GarantiaNumSerie WHERE NumIntGarantia = ?", _
    tGarantiaNumSerie.lNumIntGarantia, tGarantiaNumSerie.sNumSerie, objGarantia.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 183448
        
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 183449
        
    Do While lErro = AD_SQL_SUCESSO

        Set objGarantiaNumSerie = New ClassGarantiaNumSerie

        objGarantiaNumSerie.lNumIntGarantia = tGarantiaNumSerie.lNumIntGarantia
        objGarantiaNumSerie.sNumSerie = tGarantiaNumSerie.sNumSerie

        objGarantia.colNumSerie.Add objGarantiaNumSerie
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 183450
        
    Loop

    Call Comando_Fechar(lComando)
    
    GarantiaNumSerie_Le = SUCESSO
    
    Exit Function

Erro_GarantiaNumSerie_Le:
    
    GarantiaNumSerie_Le = gErr
        
    Select Case gErr
    
        Case 183447
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 183448 To 183350
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_GARANTIANUMSERIE", gErr, objGarantia.lNumIntDoc)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 183451)
            
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function GarantiaProduto_Le(objGarantia As ClassGarantia) As Long

Dim lComando As Long
Dim lErro As Long
Dim tGarantiaProduto As typeGarantiaProduto
Dim objGarantiaProduto As ClassGarantiaProduto

On Error GoTo Erro_GarantiaProduto_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 183452
        
    tGarantiaProduto.sProduto = String(STRING_PRODUTO, 0)
        
    lErro = Comando_Executar(lComando, "SELECT NumIntGarantia, Produto, Prazo FROM GarantiaProduto WHERE NumIntGarantia = ?", _
    tGarantiaProduto.lNumIntGarantia, tGarantiaProduto.sProduto, tGarantiaProduto.iPrazo, objGarantia.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 183453
        
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 183454
        
    Do While lErro = AD_SQL_SUCESSO

        Set objGarantiaProduto = New ClassGarantiaProduto

        objGarantiaProduto.lNumIntGarantia = tGarantiaProduto.lNumIntGarantia
        objGarantiaProduto.sProduto = tGarantiaProduto.sProduto
        objGarantiaProduto.iPrazo = tGarantiaProduto.iPrazo

        objGarantia.colProduto.Add objGarantiaProduto
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 183455
        
    Loop

    Call Comando_Fechar(lComando)
    
    GarantiaProduto_Le = SUCESSO
    
    Exit Function

Erro_GarantiaProduto_Le:
    
    GarantiaProduto_Le = gErr
        
    Select Case gErr
    
        Case 183452
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 183453 To 183455
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_GARANTIAPRODUTO", gErr, objGarantia.lNumIntDoc)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 183456)
            
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function Garantia_Le_Lote_FilialOP(objGarantia As ClassGarantia) As Long

Dim lComando As Long
Dim lErro As Long
Dim tGarantia As typeGarantia

On Error GoTo Erro_Garantia_Le_Lote_FilialOP

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 183457
        
    tGarantia.sProduto = String(STRING_PRODUTO, 0)
    tGarantia.sLote = String(STRING_LOTE_RASTREAMENTO, 0)
    tGarantia.sSerie = String(STRING_SERIE, 0)
        
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Codigo, FilialEmpresa, Garantia.Produto, DataVenda, Quantidade, Lote, FilialOP, Serie, NumNotaFiscal, TipoGarantia, GarantiaTotal, GarantiaTotalPrazo, Fornecedor, FilialFornecedor FROM Garantia, GarantiaProduto WHERE FilialEmpresa = ? AND Garantia.Produto = ? AND DataVenda <= ? AND DataVenda+Prazo >= ? AND Lote = ? AND FilialOP = ? AND Garantia.NumIntDoc = GarantiaProduto.NumIntGarantia AND DataVenda + Prazo >= ? AND GarantiaProduto.Produto = ?", _
    tGarantia.lNumIntDoc, tGarantia.lCodigo, tGarantia.iFilialEmpresa, tGarantia.sProduto, tGarantia.dtDataVenda, tGarantia.dQuantidade, tGarantia.sLote, tGarantia.iFilialOP, tGarantia.sSerie, tGarantia.lNumNotaFiscal, tGarantia.lTipoGarantia, tGarantia.iGarantiaTotal, tGarantia.iGarantiaTotalPrazo, tGarantia.lFornecedor, tGarantia.iFilialFornecedor, objGarantia.iFilialEmpresa, objGarantia.sProduto, gdtDataHoje, gdtDataHoje, objGarantia.sLote, objGarantia.iFilialOP, gdtDataHoje, objGarantia.sServico)
    If lErro <> AD_SQL_SUCESSO Then gError 183458
        
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 183459
        
    If lErro = AD_SQL_SEM_DADOS Then
    
        lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Codigo, FilialEmpresa, Garantia.Produto, DataVenda, Quantidade, Lote, FilialOP, Serie, NumNotaFiscal, TipoGarantia, GarantiaTotal, GarantiaTotalPrazo FROM Garantia WHERE FilialEmpresa = ? AND Garantia.Produto = ? AND DataVenda <= ? AND DataVenda+GarantiaTotalPrazo >= ? AND Lote = ? AND FilialOP = ? AND GarantiaTotal = ? AND NOT EXISTS (SELECT * FROM GarantiaProduto WHERE GarantiaProduto.NumIntGarantia = Garantia.NumIntDoc AND GarantiaProduto.Produto = ?)", _
        tGarantia.lNumIntDoc, tGarantia.lCodigo, tGarantia.iFilialEmpresa, tGarantia.sProduto, tGarantia.dtDataVenda, tGarantia.dQuantidade, tGarantia.sLote, tGarantia.iFilialOP, tGarantia.sSerie, tGarantia.lNumNotaFiscal, tGarantia.lTipoGarantia, tGarantia.iGarantiaTotal, tGarantia.iGarantiaTotalPrazo, objGarantia.iFilialEmpresa, objGarantia.sProduto, gdtDataHoje, gdtDataHoje, objGarantia.sLote, objGarantia.iFilialOP, GARANTIA_TOTAL, objGarantia.sServico)
        If lErro <> AD_SQL_SUCESSO Then gError 183488
            
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 183489
    
        If lErro = AD_SQL_SEM_DADOS Then gError 183460

    End If

    objGarantia.lNumIntDoc = tGarantia.lNumIntDoc
    objGarantia.lCodigo = tGarantia.lCodigo
    objGarantia.iFilialEmpresa = tGarantia.iFilialEmpresa
    objGarantia.sProduto = tGarantia.sProduto
    objGarantia.dtDataVenda = tGarantia.dtDataVenda
    objGarantia.dQuantidade = tGarantia.dQuantidade
    objGarantia.sLote = tGarantia.sLote
    objGarantia.iFilialOP = tGarantia.iFilialOP
    objGarantia.sSerie = tGarantia.sSerie
    objGarantia.lNumNotaFiscal = tGarantia.lNumNotaFiscal
    objGarantia.lTipoGarantia = tGarantia.lTipoGarantia
    objGarantia.iGarantiaTotal = tGarantia.iGarantiaTotal
    objGarantia.iGarantiaTotalPrazo = tGarantia.iGarantiaTotalPrazo
    objGarantia.lFornecedor = tGarantia.lFornecedor
    objGarantia.iFilialFornecedor = tGarantia.iFilialFornecedor

    lErro = CF("GarantiaProduto_Le", objGarantia)
    If lErro <> SUCESSO Then gError 183461

    Call Comando_Fechar(lComando)
    
    Garantia_Le_Lote_FilialOP = SUCESSO
    
    Exit Function

Erro_Garantia_Le_Lote_FilialOP:
    
    Garantia_Le_Lote_FilialOP = gErr
        
    Select Case gErr
    
        Case 183457
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 183458, 183459, 183488, 183489
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_GARANTIA_FILIALOP", gErr, objGarantia.iFilialEmpresa, objGarantia.sProduto, objGarantia.sLote, objGarantia.iFilialOP)
        
        Case 183460, 183461
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 183462)
            
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function Garantia_Le_NumSerie(objGarantia As ClassGarantia) As Long

Dim lComando As Long
Dim lErro As Long
Dim tGarantia As typeGarantia

On Error GoTo Erro_Garantia_Le_NumSerie

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 183463
        
    tGarantia.sProduto = String(STRING_PRODUTO, 0)
    tGarantia.sLote = String(STRING_LOTE_RASTREAMENTO, 0)
    tGarantia.sSerie = String(STRING_SERIE, 0)
        
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Codigo, FilialEmpresa, Garantia.Produto, DataVenda, Quantidade, Lote, FilialOP, Serie, NumNotaFiscal, TipoGarantia, GarantiaTotal, GarantiaTotalPrazo, Fornecedor, FilialFornecedor FROM Garantia, GarantiaProduto, GarantiaNumSerie WHERE Garantia.NumINtDoc = GarantiaNumSerie.NumIntGarantia AND Garantia.NumIntDoc = GarantiaProduto.NumIntGarantia AND FilialEmpresa = ? AND Garantia.Produto = ? AND DataVenda <= ? AND GarantiaNumSerie.NumSerie = ? AND DataVenda + Prazo >= ? AND GarantiaProduto.Produto = ?", _
        tGarantia.lNumIntDoc, tGarantia.lCodigo, tGarantia.iFilialEmpresa, tGarantia.sProduto, tGarantia.dtDataVenda, tGarantia.dQuantidade, tGarantia.sLote, tGarantia.iFilialOP, tGarantia.sSerie, tGarantia.lNumNotaFiscal, tGarantia.lTipoGarantia, tGarantia.iGarantiaTotal, tGarantia.iGarantiaTotalPrazo, tGarantia.lFornecedor, tGarantia.iFilialFornecedor, objGarantia.iFilialEmpresa, objGarantia.sProduto, gdtDataHoje, objGarantia.sLote, gdtDataHoje, objGarantia.sServico)
    If lErro <> AD_SQL_SUCESSO Then gError 183464
        
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 183465
        
    If lErro = AD_SQL_SEM_DADOS Then
    
        lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Codigo, FilialEmpresa, Garantia.Produto, DataVenda, Quantidade, Lote, FilialOP, Serie, NumNotaFiscal, TipoGarantia, GarantiaTotal, GarantiaTotalPrazo FROM Garantia, GarantiaNumSerie WHERE Garantia.NumINtDoc = GarantiaNumSerie.NumIntGarantia AND FilialEmpresa = ? AND Garantia.Produto = ? AND DataVenda <= ? AND GarantiaNumSerie.NumSerie = ? AND DataVenda + GarantiaTotalPrazo >= ? AND GarantiaTotal = ? AND NOT EXISTS (SELECT * FROM GarantiaProduto WHERE GarantiaProduto.NumIntGarantia = Garantia.NumIntDoc AND GarantiaProduto.Produto = ?)", _
        tGarantia.lNumIntDoc, tGarantia.lCodigo, tGarantia.iFilialEmpresa, tGarantia.sProduto, tGarantia.dtDataVenda, tGarantia.dQuantidade, tGarantia.sLote, tGarantia.iFilialOP, tGarantia.sSerie, tGarantia.lNumNotaFiscal, tGarantia.lTipoGarantia, tGarantia.iGarantiaTotal, tGarantia.iGarantiaTotalPrazo, objGarantia.iFilialEmpresa, objGarantia.sProduto, gdtDataHoje, objGarantia.sLote, gdtDataHoje, GARANTIA_TOTAL, objGarantia.sServico)
        If lErro <> AD_SQL_SUCESSO Then gError 183483

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 183484
    
        If lErro = AD_SQL_SEM_DADOS Then gError 183466
        
    End If
    
    objGarantia.lNumIntDoc = tGarantia.lNumIntDoc
    objGarantia.lCodigo = tGarantia.lCodigo
    objGarantia.iFilialEmpresa = tGarantia.iFilialEmpresa
    objGarantia.sProduto = tGarantia.sProduto
    objGarantia.dtDataVenda = tGarantia.dtDataVenda
    objGarantia.dQuantidade = tGarantia.dQuantidade
    objGarantia.sLote = tGarantia.sLote
    objGarantia.iFilialOP = tGarantia.iFilialOP
    objGarantia.sSerie = tGarantia.sSerie
    objGarantia.lNumNotaFiscal = tGarantia.lNumNotaFiscal
    objGarantia.lTipoGarantia = tGarantia.lTipoGarantia
    objGarantia.iGarantiaTotal = tGarantia.iGarantiaTotal
    objGarantia.iGarantiaTotalPrazo = tGarantia.iGarantiaTotalPrazo
    objGarantia.lFornecedor = tGarantia.lFornecedor
    objGarantia.iFilialFornecedor = tGarantia.iFilialFornecedor

    lErro = CF("GarantiaNumSerie_Le", objGarantia)
    If lErro <> SUCESSO Then gError 183467

    lErro = CF("GarantiaProduto_Le", objGarantia)
    If lErro <> SUCESSO Then gError 183468

    Call Comando_Fechar(lComando)
    
    Garantia_Le_NumSerie = SUCESSO
    
    Exit Function

Erro_Garantia_Le_NumSerie:
    
    Garantia_Le_NumSerie = gErr
        
    Select Case gErr
    
        Case 183463
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 183464, 183465, 183483, 183484
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_GARANTIA_NUMSERIE", gErr, objGarantia.iFilialEmpresa, objGarantia.sProduto, objGarantia.sLote)
        
        Case 183466, 183467, 183468
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 183469)
            
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function ItensDeContratoSrv_Le_Lote(objItensContSrv As ClassItensDeContratoSrv) As Long

Dim lComando As Long
Dim lErro As Long
Dim tItensContSrv As typeItensDeContratoSrv

On Error GoTo Erro_ItensDeContratoSrv_Le_Lote

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 183477
        
    tItensContSrv.sProduto = String(STRING_PRODUTO, 0)
    tItensContSrv.sLote = String(STRING_LOTE_RASTREAMENTO, 0)
    tItensContSrv.sCodigoContrato = String(STRING_CONTRATO_CODIGO, 0)
        
    lErro = Comando_Executar(lComando, "SELECT ItensDeContratoSrv.NumIntDoc, ItensDeContratoSrv.Codigo, Contratos.Codigo, Contratos.FilialEmpresa, ItensDeContrato.Produto, Contratos.DataIniContrato, Contratos.DataFimContrato, ItensDeContratoSrv.Quantidade, ItensDeContratoSrv.Lote, ItensDeContratoSrv.FilialOP, ItensDeContratoSrv.TipoGarantia, ItensDeContratoSrv.GarantiaTotal FROM Contratos, ItensDeContrato, ItensDeContratoSrv, ItensDeContratoSrvProd WHERE " & _
    "Contratos.NumIntDoc = ItensDeContrato.NumIntContrato AND ItensDeContrato.NumIntDoc = ItensDeContratoSrv.NumIntItemContrato AND ItensDeContratoSrv.NumIntDoc = ItensDeContratoSrvProd.NumIntItemContratoSrv AND Contratos.FilialEmpresa = ? AND ItensDeContrato.Produto = ? AND Contratos.DataIniContrato <= ? And Contratos.DataFimContrato >= ? AND ItensDeContratoSrv.Lote = ? AND ItensDeContratoSrvProd.Produto = ? AND GarantiaTotal <> ?", _
    tItensContSrv.lNumIntDoc, tItensContSrv.lCodigo, tItensContSrv.sCodigoContrato, tItensContSrv.iFilialEmpresa, tItensContSrv.sProduto, tItensContSrv.dtDataContratoIni, tItensContSrv.dtDataContratoFim, tItensContSrv.dQuantidade, tItensContSrv.sLote, tItensContSrv.iFilialOP, tItensContSrv.lTipoGarantia, tItensContSrv.iGarantiaTotal, objItensContSrv.iFilialEmpresa, objItensContSrv.sProduto, gdtDataHoje, gdtDataHoje, objItensContSrv.sLote, objItensContSrv.sServico, GARANTIA_TOTAL)
    If lErro <> AD_SQL_SUCESSO Then gError 183478
        
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 183479
        
    If lErro = AD_SQL_SEM_DADOS Then
        
        lErro = Comando_Executar(lComando, "SELECT ItensDeContratoSrv.NumIntDoc, ItensDeContratoSrv.Codigo, Contratos.Codigo, Contratos.FilialEmpresa, ItensDeContrato.Produto, Contratos.DataIniContrato, Contratos.DataFimContrato, ItensDeContratoSrv.Quantidade, ItensDeContratoSrv.Lote, ItensDeContratoSrv.FilialOP, ItensDeContratoSrv.TipoGarantia, ItensDeContratoSrv.GarantiaTotal FROM Contratos, ItensDeContrato, ItensDeContratoSrv WHERE " & _
        "Contratos.NumIntDoc = ItensDeContrato.NumIntContrato AND ItensDeContrato.NumIntDoc = ItensDeContratoSrv.NumIntItemContrato AND Contratos.FilialEmpresa = ? AND ItensDeContrato.Produto = ? AND Contratos.DataIniContrato <= ? And Contratos.DataFimContrato >= ? AND ItensDeContratoSrv.Lote = ? AND ItensDeContratoSrv.GarantiaTotal = ? AND NOT EXISTS (SELECT * FROM ItensDeContratoSrvProd WHERE ItensDeContratoSrv.NumIntDoc = ItensDeContratoSrvProd.NumIntItemContratoSrv AND ItensDeContratoSrvProd.Produto = ?)", _
        tItensContSrv.lNumIntDoc, tItensContSrv.lCodigo, tItensContSrv.sCodigoContrato, tItensContSrv.iFilialEmpresa, tItensContSrv.sProduto, tItensContSrv.dtDataContratoIni, tItensContSrv.dtDataContratoFim, tItensContSrv.dQuantidade, tItensContSrv.sLote, tItensContSrv.iFilialOP, tItensContSrv.lTipoGarantia, tItensContSrv.iGarantiaTotal, objItensContSrv.iFilialEmpresa, objItensContSrv.sProduto, gdtDataHoje, gdtDataHoje, objItensContSrv.sLote, GARANTIA_TOTAL, objItensContSrv.sServico)
        If lErro <> AD_SQL_SUCESSO Then gError 183491
        
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 183492
        
        If lErro = AD_SQL_SEM_DADOS Then gError 183480

    End If

    objItensContSrv.lNumIntDoc = tItensContSrv.lNumIntDoc
    objItensContSrv.lCodigo = tItensContSrv.lCodigo
    objItensContSrv.sCodigoContrato = tItensContSrv.sCodigoContrato
    objItensContSrv.iFilialEmpresa = tItensContSrv.iFilialEmpresa
    objItensContSrv.sProduto = tItensContSrv.sProduto
    objItensContSrv.dtDataContratoIni = tItensContSrv.dtDataContratoIni
    objItensContSrv.dtDataContratoFim = tItensContSrv.dtDataContratoFim
    objItensContSrv.dQuantidade = tItensContSrv.dQuantidade
    objItensContSrv.sLote = tItensContSrv.sLote
    objItensContSrv.iFilialOP = tItensContSrv.iFilialOP
    objItensContSrv.lTipoGarantia = tItensContSrv.lTipoGarantia
    objItensContSrv.iGarantiaTotal = tItensContSrv.iGarantiaTotal

    lErro = CF("ItensDeContratoSrvProd_Le", objItensContSrv)
    If lErro <> SUCESSO Then gError 183481

    Call Comando_Fechar(lComando)
    
    ItensDeContratoSrv_Le_Lote = SUCESSO
    
    Exit Function

Erro_ItensDeContratoSrv_Le_Lote:
    
    ItensDeContratoSrv_Le_Lote = gErr
        
    Select Case gErr
    
        Case 183477
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 183478, 183479, 183491, 183492
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMCONTRATOSRV_LOTE", gErr, objItensContSrv.iFilialEmpresa, objItensContSrv.sProduto, objItensContSrv.sLote)
        
        Case 183480, 183481, 183490
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 183482)
            
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function ItensDeContratoSrvNumSerie_Le(objItensContSrv As ClassItensDeContratoSrv) As Long

Dim lComando As Long
Dim lErro As Long
Dim tItensContSrvNumSerie As typeItensDeContratoSrvNumSerie
Dim objItensContSrvNumSerie As ClassItensContSrvNumSerie

On Error GoTo Erro_ItensDeContratoSrvNumSerie_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 183493
        
    tItensContSrvNumSerie.sNumSerie = String(STRING_LOTE_RASTREAMENTO, 0)
        
    lErro = Comando_Executar(lComando, "SELECT NumIntItemContratoSrv, NumSerie FROM ItensDeContratoSrvNumSerie WHERE NumIntItemContratoSrv = ?", _
    tItensContSrvNumSerie.lNumIntItemContratoSrv, tItensContSrvNumSerie.sNumSerie, objItensContSrv.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 183494
        
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 183495
        
    Do While lErro = AD_SQL_SUCESSO

        Set objItensContSrvNumSerie = New ClassItensContSrvNumSerie

        objItensContSrvNumSerie.lNumIntItemContratoSrv = tItensContSrvNumSerie.lNumIntItemContratoSrv
        objItensContSrvNumSerie.sNumSerie = tItensContSrvNumSerie.sNumSerie

        objItensContSrv.colNumSerie.Add objItensContSrvNumSerie
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 183496
        
    Loop

    Call Comando_Fechar(lComando)
    
    ItensDeContratoSrvNumSerie_Le = SUCESSO
    
    Exit Function

Erro_ItensDeContratoSrvNumSerie_Le:
    
    ItensDeContratoSrvNumSerie_Le = gErr
        
    Select Case gErr
    
        Case 183493
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 183494, 183495, 183496
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSDECONTRATOSRVNUMSERIE", gErr, objItensContSrv.lNumIntDoc)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 183497)
            
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function ItensDeContratoSrvProd_Le(objItensContSrv As ClassItensDeContratoSrv) As Long

Dim lComando As Long
Dim lErro As Long
Dim tItensContSrvProd As typeItensDeContratoSrvProd
Dim objItensContSrvProd As ClassItensDeContratoSrvProd

On Error GoTo Erro_ItensDeContratoSrvProd_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 183498
        
    tItensContSrvProd.sProduto = String(STRING_PRODUTO, 0)
        
    lErro = Comando_Executar(lComando, "SELECT NumIntItemContratoSrv, Produto FROM ItensDeContratoSrvProd WHERE NumIntItemContratoSrv = ?", _
    tItensContSrvProd.lNumIntItemContratoSrv, tItensContSrvProd.sProduto, objItensContSrv.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 183499
        
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 183500
        
    Do While lErro = AD_SQL_SUCESSO

        Set objItensContSrvProd = New ClassItensDeContratoSrvProd

        objItensContSrvProd.lNumIntItemContratoSrv = tItensContSrvProd.lNumIntItemContratoSrv
        objItensContSrvProd.sProduto = tItensContSrvProd.sProduto

        objItensContSrv.colProduto.Add objItensContSrvProd
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 183501
        
    Loop

    Call Comando_Fechar(lComando)
    
    ItensDeContratoSrvProd_Le = SUCESSO
    
    Exit Function

Erro_ItensDeContratoSrvProd_Le:
    
    ItensDeContratoSrvProd_Le = gErr
        
    Select Case gErr
    
        Case 183498
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 183499 To 183501
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSDECONTRATOSRVPROD", gErr, objItensContSrv.lNumIntDoc)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 183502)
            
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function ItensDeContratoSrv_Le_Lote_FilialOP(objItensContSrv As ClassItensDeContratoSrv) As Long

Dim lComando As Long
Dim lErro As Long
Dim tItensContSrv As typeItensDeContratoSrv

On Error GoTo Erro_ItensDeContratoSrv_Le_Lote_FilialOP

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 183503
        
    tItensContSrv.sProduto = String(STRING_PRODUTO, 0)
    tItensContSrv.sLote = String(STRING_LOTE_RASTREAMENTO, 0)
    tItensContSrv.sCodigoContrato = String(STRING_CONTRATO_CODIGO, 0)
        
    lErro = Comando_Executar(lComando, "SELECT ItensDeContratoSrv.NumIntDoc, ItensDeContratoSrv.Codigo, Contratos.Codigo, Contratos.FilialEmpresa, ItensDeContrato.Produto, Contratos.DataIniContrato, Contratos.DataFimContrato, ItensDeContratoSrv.Quantidade, ItensDeContratoSrv.Lote, ItensDeContratoSrv.FilialOP, ItensDeContratoSrv.TipoGarantia, ItensDeContratoSrv.GarantiaTotal FROM Contratos, ItensDeContrato, ItensDeContratoSrv, ItensDeContratoSrvProd WHERE " & _
    "Contratos.NumIntDoc = ItensDeContrato.NumIntContrato AND ItensDeContrato.NumIntDoc = ItensDeContratoSrv.NumIntItemContrato AND ItensDeContratoSrv.NumIntDoc = ItensDeContratoSrvProd.NumIntItemContratoSrv AND Contratos.FilialEmpresa = ? AND ItensDeContrato.Produto = ? And Contratos.DataIniContrato <= ? And Contratos.DataFimContrato >= ? AND ItensDeContratoSrv.Lote = ? AND ItensDeContratoSrv.FilialOP = ? AND ItensDeContratoSrvProd.Produto = ? AND GarantiaTotal <> ?", _
    tItensContSrv.lNumIntDoc, tItensContSrv.lCodigo, tItensContSrv.sCodigoContrato, tItensContSrv.iFilialEmpresa, tItensContSrv.sProduto, tItensContSrv.dtDataContratoIni, tItensContSrv.dtDataContratoFim, tItensContSrv.dQuantidade, tItensContSrv.sLote, tItensContSrv.iFilialOP, tItensContSrv.lTipoGarantia, tItensContSrv.iGarantiaTotal, objItensContSrv.iFilialEmpresa, objItensContSrv.sProduto, gdtDataHoje, gdtDataHoje, objItensContSrv.sLote, objItensContSrv.iFilialOP, objItensContSrv.sServico, GARANTIA_TOTAL)
    If lErro <> AD_SQL_SUCESSO Then gError 183504
        
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 183505
        
    If lErro = AD_SQL_SEM_DADOS Then
    
        lErro = Comando_Executar(lComando, "SELECT ItensDeContratoSrv.NumIntDoc, ItensDeContratoSrv.Codigo, Contratos.Codigo, Contratos.FilialEmpresa, ItensDeContrato.Produto, Contratos.DataIniContrato, Contratos.DataFimContrato, ItensDeContratoSrv.Quantidade, ItensDeContratoSrv.Lote, ItensDeContratoSrv.FilialOP, ItensDeContratoSrv.TipoGarantia, ItensDeContratoSrv.GarantiaTotal FROM Contratos, ItensDeContrato, ItensDeContratoSrv WHERE " & _
        "Contratos.NumIntDoc = ItensDeContrato.NumIntContrato AND ItensDeContrato.NumIntDoc = ItensDeContratoSrv.NumIntItemContrato AND Contratos.FilialEmpresa = ? AND ItensDeContrato.Produto = ? AND Contratos.DataIniContrato <= ? And Contratos.DataFimContrato >= ? AND ItensDeContratoSrv.Lote = ? AND ItensDeContratoSrv.FilialOP = ? AND ItensDeContratoSrv.GarantiaTotal = ? AND NOT EXISTS (SELECT * FROM ItensDeContratoSrvProd WHERE ItensDeContratoSrv.NumIntDoc = ItensDeContratoSrvProd.NumIntItemContratoSrv AND ItensDeContratoSrvProd.Produto = ?)", _
        tItensContSrv.lNumIntDoc, tItensContSrv.lCodigo, tItensContSrv.sCodigoContrato, tItensContSrv.iFilialEmpresa, tItensContSrv.sProduto, tItensContSrv.dtDataContratoIni, tItensContSrv.dtDataContratoFim, tItensContSrv.dQuantidade, tItensContSrv.sLote, tItensContSrv.iFilialOP, tItensContSrv.lTipoGarantia, tItensContSrv.iGarantiaTotal, objItensContSrv.iFilialEmpresa, objItensContSrv.sProduto, gdtDataHoje, gdtDataHoje, objItensContSrv.sLote, objItensContSrv.iFilialOP, GARANTIA_TOTAL, objItensContSrv.sServico)
        If lErro <> AD_SQL_SUCESSO Then gError 183506
            
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 183507
    
        If lErro = AD_SQL_SEM_DADOS Then gError 183508

    End If

    objItensContSrv.lNumIntDoc = tItensContSrv.lNumIntDoc
    objItensContSrv.lCodigo = tItensContSrv.lCodigo
    objItensContSrv.sCodigoContrato = tItensContSrv.sCodigoContrato
    objItensContSrv.iFilialEmpresa = tItensContSrv.iFilialEmpresa
    objItensContSrv.sProduto = tItensContSrv.sProduto
    objItensContSrv.dtDataContratoIni = tItensContSrv.dtDataContratoIni
    objItensContSrv.dtDataContratoFim = tItensContSrv.dtDataContratoFim
    objItensContSrv.dQuantidade = tItensContSrv.dQuantidade
    objItensContSrv.sLote = tItensContSrv.sLote
    objItensContSrv.iFilialOP = tItensContSrv.iFilialOP
    objItensContSrv.lTipoGarantia = tItensContSrv.lTipoGarantia
    objItensContSrv.iGarantiaTotal = tItensContSrv.iGarantiaTotal

    lErro = CF("ItensDeContratoSrvProd_Le", objItensContSrv)
    If lErro <> SUCESSO Then gError 183509

    Call Comando_Fechar(lComando)
    
    ItensDeContratoSrv_Le_Lote_FilialOP = SUCESSO
    
    Exit Function

Erro_ItensDeContratoSrv_Le_Lote_FilialOP:
    
    ItensDeContratoSrv_Le_Lote_FilialOP = gErr
        
    Select Case gErr
    
        Case 183503
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 183504, 183505, 183506, 183507
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMCONTRATOSRV_FILIALOP", gErr, objItensContSrv.iFilialEmpresa, objItensContSrv.sProduto, objItensContSrv.sLote, objItensContSrv.iFilialOP)
        
        Case 183508, 183509
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 183510)
            
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function ItensDeContratoSrv_Le_NumSerie(objItensContSrv As ClassItensDeContratoSrv) As Long

Dim lComando As Long
Dim lErro As Long
Dim tItensContSrv As typeItensDeContratoSrv

On Error GoTo Erro_ItensDeContratoSrv_Le_NumSerie

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 183511
        
    tItensContSrv.sProduto = String(STRING_PRODUTO, 0)
    tItensContSrv.sLote = String(STRING_LOTE_RASTREAMENTO, 0)
    tItensContSrv.sCodigoContrato = String(STRING_CONTRATO_CODIGO, 0)
        
    lErro = Comando_Executar(lComando, "SELECT ItensDeContratoSrv.NumIntDoc, ItensDeContratoSrv.Codigo, Contratos.Codigo, Contratos.FilialEmpresa, ItensDeContrato.Produto, Contratos.DataIniContrato, Contratos.DataFimContrato, ItensDeContratoSrv.Quantidade, ItensDeContratoSrv.Lote, ItensDeContratoSrv.FilialOP, ItensDeContratoSrv.TipoGarantia, ItensDeContratoSrv.GarantiaTotal FROM Contratos, ItensDeContrato, ItensDeContratoSrv, ItensDeContratoSrvProd, ItensDeContratoSrvNumSerie WHERE " & _
    "Contratos.NumIntDoc = ItensDeContrato.NumIntContrato AND ItensDeContrato.NumIntDoc = ItensDeContratoSrv.NumIntItemContrato AND ItensDeContratoSrv.NumIntDoc = ItensDeContratoSrvProd.NumIntItemContratoSrv AND ItensDeContratoSrv.NumIntDoc = ItensDeContratoSrvNumSerie.NumIntItemContratoSrv AND Contratos.FilialEmpresa = ? AND ItensDeContrato.Produto = ? AND ItensDeContratoSrvProd.Produto = ? And Contratos.DataIniContrato <= ? And Contratos.DataFimContrato >= ? AND ItensDeContratoSrvProd.Produto = ? AND GarantiaTotal <> ? AND ItensDeContratoSrvNumSerie.NumSerie = ?", _
    tItensContSrv.lNumIntDoc, tItensContSrv.lCodigo, tItensContSrv.sCodigoContrato, tItensContSrv.iFilialEmpresa, tItensContSrv.sProduto, tItensContSrv.dtDataContratoIni, tItensContSrv.dtDataContratoFim, tItensContSrv.dQuantidade, tItensContSrv.sLote, tItensContSrv.iFilialOP, tItensContSrv.lTipoGarantia, tItensContSrv.iGarantiaTotal, objItensContSrv.iFilialEmpresa, objItensContSrv.sProduto, objItensContSrv.sProduto, gdtDataHoje, gdtDataHoje, objItensContSrv.sServico, GARANTIA_TOTAL, objItensContSrv.sLote)
    If lErro <> AD_SQL_SUCESSO Then gError 183512
        
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 183513
        
    If lErro = AD_SQL_SEM_DADOS Then
    
        lErro = Comando_Executar(lComando, "SELECT ItensDeContratoSrv.NumIntDoc, ItensDeContratoSrv.Codigo, Contratos.Codigo, Contratos.FilialEmpresa, ItensDeContrato.Produto, Contratos.DataIniContrato, Contratos.DataFimContrato, ItensDeContratoSrv.Quantidade, ItensDeContratoSrv.Lote, ItensDeContratoSrv.FilialOP, ItensDeContratoSrv.TipoGarantia, ItensDeContratoSrv.GarantiaTotal FROM Contratos, ItensDeContrato, ItensDeContratoSrv, ItensDeContratoSrvNumSerie WHERE " & _
        "Contratos.NumIntDoc = ItensDeContrato.NumIntContrato AND ItensDeContrato.NumIntDoc = ItensDeContratoSrv.NumIntItemContrato AND ItensDeContratoSrv.NumIntDoc = ItensDeContratoSrvNumSerie.NumIntItemContratoSrv AND Contratos.FilialEmpresa = ? AND ItensDeContrato.Produto = ? AND Contratos.DataIniContrato <= ? And Contratos.DataFimContrato >= ? AND ItensDeContratoSrv.GarantiaTotal = ? AND ItensDeContratoSrvNumSerie.NumSerie = ? AND NOT EXISTS (SELECT ItensDeContratoSrvProd.NumIntItemContratoSrv FROM ItensDeContratoSrvProd WHERE ItensDeContratoSrv.NumIntDoc = ItensDeContratoSrvProd.NumIntItemContratoSrv AND ItensDeContratoSrvProd.Produto = ?)", _
        tItensContSrv.lNumIntDoc, tItensContSrv.lCodigo, tItensContSrv.sCodigoContrato, tItensContSrv.iFilialEmpresa, tItensContSrv.sProduto, tItensContSrv.dtDataContratoIni, tItensContSrv.dtDataContratoFim, tItensContSrv.dQuantidade, tItensContSrv.sLote, tItensContSrv.iFilialOP, tItensContSrv.lTipoGarantia, tItensContSrv.iGarantiaTotal, objItensContSrv.iFilialEmpresa, objItensContSrv.sProduto, gdtDataHoje, gdtDataHoje, GARANTIA_TOTAL, objItensContSrv.sLote, objItensContSrv.sServico)
        If lErro <> AD_SQL_SUCESSO Then gError 183514
    
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 183515
    
        If lErro = AD_SQL_SEM_DADOS Then gError 183516
        
    End If
    
    objItensContSrv.lNumIntDoc = tItensContSrv.lNumIntDoc
    objItensContSrv.lCodigo = tItensContSrv.lCodigo
    objItensContSrv.sCodigoContrato = tItensContSrv.sCodigoContrato
    objItensContSrv.iFilialEmpresa = tItensContSrv.iFilialEmpresa
    objItensContSrv.sProduto = tItensContSrv.sProduto
    objItensContSrv.dtDataContratoIni = tItensContSrv.dtDataContratoIni
    objItensContSrv.dtDataContratoFim = tItensContSrv.dtDataContratoFim
    objItensContSrv.dQuantidade = tItensContSrv.dQuantidade
    objItensContSrv.sLote = tItensContSrv.sLote
    objItensContSrv.iFilialOP = tItensContSrv.iFilialOP
    objItensContSrv.lTipoGarantia = tItensContSrv.lTipoGarantia
    objItensContSrv.iGarantiaTotal = tItensContSrv.iGarantiaTotal

    lErro = CF("ItensDeContratoSrvNumSerie_Le", objItensContSrv)
    If lErro <> SUCESSO Then gError 183517

    lErro = CF("ItensDeContratoSrvProd_Le", objItensContSrv)
    If lErro <> SUCESSO Then gError 183518

    Call Comando_Fechar(lComando)
    
    ItensDeContratoSrv_Le_NumSerie = SUCESSO
    
    Exit Function

Erro_ItensDeContratoSrv_Le_NumSerie:
    
    ItensDeContratoSrv_Le_NumSerie = gErr
        
    Select Case gErr
    
        Case 183511
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 183512 To 183515
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMCONTRATOSRV_NUMSERIE", gErr, objItensContSrv.iFilialEmpresa, objItensContSrv.sProduto, objItensContSrv.sLote)
        
        Case 183516, 183517, 183518
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 183519)
            
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function Garantia_Le(objGarantia As ClassGarantia) As Long

Dim lComando As Long
Dim lErro As Long
Dim tGarantia As typeGarantia

On Error GoTo Erro_Garantia_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 183565
        
    tGarantia.sProduto = String(STRING_PRODUTO, 0)
    tGarantia.sLote = String(STRING_LOTE_RASTREAMENTO, 0)
    tGarantia.sSerie = String(STRING_SERIE, 0)
        
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Codigo, FilialEmpresa, Produto, DataVenda, Quantidade, Lote, FilialOP, Serie, NumNotaFiscal, TipoGarantia, GarantiaTotal, GarantiaTotalPrazo, Fornecedor, FilialFornecedor, CliFabr, FilialCliFabr FROM Garantia WHERE FilialEmpresa = ? AND Codigo = ? ", _
    tGarantia.lNumIntDoc, tGarantia.lCodigo, tGarantia.iFilialEmpresa, tGarantia.sProduto, tGarantia.dtDataVenda, tGarantia.dQuantidade, tGarantia.sLote, tGarantia.iFilialOP, tGarantia.sSerie, tGarantia.lNumNotaFiscal, tGarantia.lTipoGarantia, tGarantia.iGarantiaTotal, tGarantia.iGarantiaTotalPrazo, tGarantia.lFornecedor, tGarantia.iFilialFornecedor, tGarantia.lCliFabr, tGarantia.iFilialCliFabr, objGarantia.iFilialEmpresa, objGarantia.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 183566
        
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 183567
        
    If lErro = AD_SQL_SEM_DADOS Then gError 183568

    objGarantia.lNumIntDoc = tGarantia.lNumIntDoc
    objGarantia.lCodigo = tGarantia.lCodigo
    objGarantia.iFilialEmpresa = tGarantia.iFilialEmpresa
    objGarantia.sProduto = tGarantia.sProduto
    objGarantia.dtDataVenda = tGarantia.dtDataVenda
    objGarantia.dQuantidade = tGarantia.dQuantidade
    objGarantia.sLote = tGarantia.sLote
    objGarantia.iFilialOP = tGarantia.iFilialOP
    objGarantia.sSerie = tGarantia.sSerie
    objGarantia.lNumNotaFiscal = tGarantia.lNumNotaFiscal
    objGarantia.lTipoGarantia = tGarantia.lTipoGarantia
    objGarantia.iGarantiaTotal = tGarantia.iGarantiaTotal
    objGarantia.iGarantiaTotalPrazo = tGarantia.iGarantiaTotalPrazo
    objGarantia.lFornecedor = tGarantia.lFornecedor
    objGarantia.iFilialFornecedor = tGarantia.iFilialFornecedor
    objGarantia.lCliFabr = tGarantia.lCliFabr
    objGarantia.iFilialCliFabr = tGarantia.iFilialCliFabr
    
    lErro = CF("GarantiaProduto_Le", objGarantia)
    If lErro <> SUCESSO Then gError 183575

    lErro = CF("GarantiaNumSerie_Le", objGarantia)
    If lErro <> SUCESSO Then gError 183576

    Call Comando_Fechar(lComando)
    
    Garantia_Le = SUCESSO
    
    Exit Function

Erro_Garantia_Le:
    
    Garantia_Le = gErr
        
    Select Case gErr
    
        Case 183565
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 183566, 183567
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_GARANTIA", gErr, objGarantia.iFilialEmpresa, objGarantia.lCodigo)
        
        Case 183568, 183575, 183576
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 183569)
            
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function Testa_Garantia(objGarantia As ClassGarantia) As Long

Dim lErro As Long
Dim objGarantiaNumSerie As ClassGarantiaNumSerie
Dim objGarantiaProduto As ClassGarantiaProduto
Dim iProdutoPreenchido As Integer
Dim iServicoPreenchido As Integer
Dim sProdutoFormatado As String
Dim sServicoFormatado As String
Dim objGarantiaBD As New ClassGarantia
Dim objProduto As New ClassProduto
Dim iAchou As Integer

On Error GoTo Erro_Testa_Garantia


    If gobjSRV.iValidaGarantia = VALIDA_GARANTIA Then

        objGarantiaBD.iFilialEmpresa = objGarantia.iFilialEmpresa
        objGarantiaBD.lCodigo = objGarantia.lCodigo

        lErro = CF("Garantia_Le", objGarantiaBD)
        If lErro <> SUCESSO And lErro <> 183568 Then gError 183576

        If lErro <> SUCESSO Then gError 183577

        'Formata o Produto para o BD
        lErro = CF("Produto_Formata", objGarantia.sProduto, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 183578
            
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
        
            If objGarantiaBD.sProduto <> sProdutoFormatado Then gError 183579
            
            objProduto.sCodigo = sProdutoFormatado
        
            'Lê os demais atributos do Produto
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 183580
                
            If lErro = 28030 Then gError 183581
    
    
            lErro = CF("Produto_Formata", objGarantia.sServico, sServicoFormatado, iServicoPreenchido)
            If lErro <> SUCESSO Then gError 183582
    
            If iServicoPreenchido = PRODUTO_PREENCHIDO Then
                
                iAchou = 0
                'pesquisa o servico
                For Each objGarantiaProduto In objGarantiaBD.colProduto
                    If objGarantiaProduto.sProduto = sServicoFormatado Then
                        iAchou = 1
                        Exit For
                    End If
                Next
                    
                If objGarantiaBD.iGarantiaTotal = GARANTIA_TOTAL Then
                    
                    If iAchou = 1 Then
                        If objGarantiaBD.dtDataVenda + objGarantiaProduto.iPrazo < gdtDataHoje Or gdtDataHoje < objGarantiaBD.dtDataVenda Then gError 183583
                    Else
                        If objGarantiaBD.dtDataVenda + objGarantiaBD.iGarantiaTotalPrazo < gdtDataHoje Or gdtDataHoje < objGarantiaBD.dtDataVenda Then gError 183584
                    End If
                    
                Else
                
                    If iAchou = 0 Then gError 183585
                    
                    If iAchou = 1 Then
                        If objGarantiaBD.dtDataVenda + objGarantiaProduto.iPrazo < gdtDataHoje Or gdtDataHoje < objGarantiaBD.dtDataVenda Then gError 183586
                    End If
        
                End If
    
            End If
    
            If objProduto.iRastro = PRODUTO_RASTRO_LOTE Then
        
                If Len(Trim(objGarantia.sLote)) > 0 Then
                
                    If objGarantiaBD.sLote <> objGarantia.sLote Then gError 183587
        
                End If
                
            ElseIf objProduto.iRastro = PRODUTO_RASTRO_OP Then
                                
                If Len(Trim(objGarantia.sLote)) > 0 And Len(Trim(objGarantia.iFilialOP)) > 0 Then
        
                    If objGarantiaBD.sLote <> objGarantia.sLote Or _
                    objGarantiaBD.iFilialOP <> objGarantia.iFilialOP Then gError 183588
                    
                End If
                    
            ElseIf objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then
    
                If Len(Trim(objGarantia.sLote)) > 0 Then
                
                    iAchou = 0
                
                    For Each objGarantiaNumSerie In objGarantiaBD.colNumSerie
                
                        If objGarantiaNumSerie.sNumSerie = objGarantia.sLote Then
                            iAchou = 1
                            Exit For
                        End If
                    Next
                    
                    If iAchou = 0 Then gError 183589
                    
                End If
                    
            End If
        
        End If

    End If

    Testa_Garantia = SUCESSO

    Exit Function

Erro_Testa_Garantia:

    Testa_Garantia = gErr

    Select Case gErr

        Case 183576, 183578, 183580, 183582

        Case 183577
            Call Rotina_Erro(vbOKOnly, "ERRO_GARANTIA_INEXISTENTE", gErr, objGarantiaBD.iFilialEmpresa, objGarantiaBD.lCodigo)

        Case 183579
            Call Rotina_Erro(vbOKOnly, "ERRO_GARANTIA_PRODUTO_ERRADO", gErr, objGarantiaBD.sProduto, objGarantia.sProduto)

        Case 183581
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case 183583, 183584, 183586
            Call Rotina_Erro(vbOKOnly, "ERRO_SERVICO_FORA_PRAZO_GARANTIA", gErr, objProduto.sCodigo, objGarantia.sServico)

        Case 183585
            Call Rotina_Erro(vbOKOnly, "ERRO_SERVICO_SEM_GARANTIA", gErr, objProduto.sCodigo, objGarantia.sServico)

        Case 183587
            Call Rotina_Erro(vbOKOnly, "ERRO_LOTE_NAO_GARANTIA", gErr, objProduto.sCodigo, objGarantia.sLote, objGarantia.lCodigo)

        Case 183588
            Call Rotina_Erro(vbOKOnly, "ERRO_LOTE_FILIALOP_NAO_GARANTIA", gErr, objProduto.sCodigo, objGarantia.sLote, objGarantia.iFilialOP, objGarantia.lCodigo)

        Case 183589
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMSERIE_NAO_GARANTIA", gErr, objProduto.sCodigo, objGarantia.sLote, objGarantia.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183590)

    End Select

    Exit Function

End Function

Function Testa_Contrato(objItensDeContratoSRV As ClassItensDeContratoSrv) As Long

Dim lErro As Long
Dim objItensDeContratoSrvBD As New ClassItensDeContratoSrv
Dim colItensContSrv As New Collection
Dim objProduto As New ClassProduto
Dim objItensDeContratoSrvProd As ClassItensDeContratoSrvProd
Dim objItensDeContratoSrvNumSerie As ClassItensContSrvNumSerie
Dim iProdutoPreenchido As Integer
Dim iServicoPreenchido As Integer
Dim sProdutoFormatado As String
Dim sServicoFormatado As String
Dim iAchouProduto As Integer
Dim iAchouServico As Integer

On Error GoTo Erro_Testa_Contrato


    If gobjSRV.iValidaManutencao = VALIDA_MANUTENCAO Then

        objItensDeContratoSrvBD.iFilialEmpresa = objItensDeContratoSRV.iFilialEmpresa
        objItensDeContratoSrvBD.sCodigoContrato = objItensDeContratoSRV.sCodigoContrato

        lErro = CF("Contrato_Le_Srv", objItensDeContratoSrvBD, colItensContSrv)
        If lErro <> SUCESSO And lErro <> 183595 Then gError 183600

        If lErro <> SUCESSO Then gError 183601

        'Formata o Produto para o BD
        lErro = CF("Produto_Formata", objItensDeContratoSRV.sProduto, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 183602
            
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

            For Each objItensDeContratoSrvBD In colItensContSrv
        
                If objItensDeContratoSrvBD.sProduto = sProdutoFormatado Then
                
                    objProduto.sCodigo = sProdutoFormatado
                
                    'Lê os demais atributos do Produto
                    lErro = CF("Produto_Le", objProduto)
                    If lErro <> SUCESSO And lErro <> 28030 Then gError 183603
                
                    If lErro = 28030 Then gError 183604
    
                    If gdtDataHoje < objItensDeContratoSrvBD.dtDataContratoIni Or (gdtDataHoje > objItensDeContratoSrvBD.dtDataContratoFim And objItensDeContratoSrvBD.dtDataContratoFim <> DATA_NULA) Then gError 183605
    
                    iAchouProduto = 0
                    
                    If objProduto.iRastro = PRODUTO_RASTRO_LOTE Then
                
                        If Len(Trim(objItensDeContratoSRV.sLote)) > 0 Then
                        
                            If objItensDeContratoSrvBD.sLote = objItensDeContratoSRV.sLote Then iAchouProduto = 1
                
                        End If
                        
                    ElseIf objProduto.iRastro = PRODUTO_RASTRO_OP Then
                                        
                        If Len(Trim(objItensDeContratoSRV.sLote)) > 0 And Len(Trim(objItensDeContratoSRV.iFilialOP)) > 0 Then
                
                            If objItensDeContratoSrvBD.sLote = objItensDeContratoSRV.sLote And _
                            objItensDeContratoSrvBD.iFilialOP = objItensDeContratoSRV.iFilialOP Then iAchouProduto = 1
                            
                        End If
                            
                    ElseIf objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then
            
                        If Len(Trim(objItensDeContratoSRV.sLote)) > 0 Then
                        
                            For Each objItensDeContratoSrvNumSerie In objItensDeContratoSRV.colNumSerie
                        
                                If objItensDeContratoSrvNumSerie.sNumSerie = objItensDeContratoSRV.sLote Then
                                    iAchouProduto = 1
                                    Exit For
                                End If
                            Next
                            
                        End If
                        
                    Else
                        iAchouProduto = 1
                    End If
                    
                    If iAchouProduto = 1 Then
                    
                        lErro = CF("Produto_Formata", objItensDeContratoSRV.sServico, sServicoFormatado, iServicoPreenchido)
                        If lErro <> SUCESSO Then gError 183606
    
                        If iServicoPreenchido = PRODUTO_PREENCHIDO Then
                
                            iAchouServico = 0
                            'pesquisa o servico
                            For Each objItensDeContratoSrvProd In objItensDeContratoSrvBD.colProduto
                                If objItensDeContratoSrvProd.sProduto = sServicoFormatado Then
                                    iAchouServico = 1
                                    Exit For
                                End If
                            Next
                        
                            If objItensDeContratoSrvBD.iGarantiaTotal = GARANTIA_TOTAL Then
                        
                                If iAchouServico = 1 Then gError 183607
                                    
                            Else
                            
                                If iAchouServico = 0 Then gError 183608
                                
                            End If
                        
                        End If
                        
                    End If
    
                    Exit For
    
                End If
                
            Next
        
            If iAchouProduto = 0 Then gError 183609
        
        End If

    End If

    Testa_Contrato = SUCESSO

    Exit Function

Erro_Testa_Contrato:

    Testa_Contrato = gErr

    Select Case gErr

        Case 183600, 183602, 183603, 183606

        Case 183601
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTRATO_INEXISTENTE", gErr, objItensDeContratoSRV.iFilialEmpresa, objItensDeContratoSRV.sCodigoContrato)

        Case 183604
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case 183605
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_FORA_PRAZO_CONTRATO", gErr, objItensDeContratoSRV.iFilialEmpresa, objItensDeContratoSRV.sCodigoContrato)

        Case 183607, 183608
            Call Rotina_Erro(vbOKOnly, "ERRO_SERVICO_SEM_CONTRATO", gErr, objProduto.sCodigo, objItensDeContratoSRV.sServico)

        Case 183609
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_SEM_CONTRATO", gErr, objItensDeContratoSRV.iFilialEmpresa, objItensDeContratoSRV.sCodigoContrato, objProduto.sCodigo)


        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183610)

    End Select

    Exit Function

End Function

Public Function Contrato_Le_Srv(objItensContSrvAux As ClassItensDeContratoSrv, colItensContSrv As Collection) As Long

Dim lComando As Long
Dim lErro As Long
Dim tItensContSrv As typeItensDeContratoSrv
Dim objItensContSrv As ClassItensDeContratoSrv

On Error GoTo Erro_Contrato_Le_Srv

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 183592
        
    tItensContSrv.sProduto = String(STRING_PRODUTO, 0)
    tItensContSrv.sLote = String(STRING_LOTE_RASTREAMENTO, 0)
    tItensContSrv.sCodigoContrato = String(STRING_CONTRATO_CODIGO, 0)
    
        
    lErro = Comando_Executar(lComando, "SELECT Contratos.DataIniContrato, Contratos.DataFimContrato, ItensDeContrato.Produto, ItensDeContratoSrv.NumIntDoc, ItensDeContratoSrv.Codigo, ItensDeContratoSrv.NumIntItemContrato, ItensDeContratoSrv.Quantidade, ItensDeContratoSrv.Lote, ItensDeContratoSrv.FilialOP, ItensDeContratoSrv.TipoGarantia, ItensDeContratoSrv.GarantiaTotal FROM Contratos, ItensDeContrato, ItensDeContratoSrv WHERE " & _
    "Contratos.NumIntDoc = ItensDeContrato.NumIntContrato AND ItensDeContrato.NumIntDoc = ItensDeContratoSrv.NumIntItemContrato AND Contratos.FilialEmpresa = ? AND Contratos.Codigo = ? AND Contratos.Ativo = ?", _
    tItensContSrv.dtDataContratoIni, tItensContSrv.dtDataContratoFim, tItensContSrv.sProduto, tItensContSrv.lNumIntDoc, tItensContSrv.lCodigo, tItensContSrv.lNumIntItemContrato, tItensContSrv.dQuantidade, tItensContSrv.sLote, tItensContSrv.iFilialOP, tItensContSrv.lTipoGarantia, tItensContSrv.iGarantiaTotal, objItensContSrvAux.iFilialEmpresa, objItensContSrvAux.sCodigoContrato, CONTRATO_ATIVO)
    If lErro <> AD_SQL_SUCESSO Then gError 183593
        
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 183594
        
    If lErro = AD_SQL_SEM_DADOS Then gError 183595
    
    Do While lErro = AD_SQL_SUCESSO
        
        Set objItensContSrv = New ClassItensDeContratoSrv

        objItensContSrv.lNumIntDoc = tItensContSrv.lNumIntDoc
        objItensContSrv.lCodigo = tItensContSrv.lCodigo
        objItensContSrv.sCodigoContrato = objItensContSrvAux.sCodigoContrato
        objItensContSrv.iFilialEmpresa = objItensContSrvAux.iFilialEmpresa
        objItensContSrv.sProduto = tItensContSrv.sProduto
        objItensContSrv.dtDataContratoIni = tItensContSrv.dtDataContratoIni
        objItensContSrv.dtDataContratoFim = tItensContSrv.dtDataContratoFim
        objItensContSrv.dQuantidade = tItensContSrv.dQuantidade
        objItensContSrv.sLote = tItensContSrv.sLote
        objItensContSrv.iFilialOP = tItensContSrv.iFilialOP
        objItensContSrv.lTipoGarantia = tItensContSrv.lTipoGarantia
        objItensContSrv.iGarantiaTotal = tItensContSrv.iGarantiaTotal

        colItensContSrv.Add objItensContSrv
        
        lErro = CF("ItensDeContratoSrvProd_Le", objItensContSrv)
        If lErro <> SUCESSO Then gError 183596
        
        lErro = CF("ItensDeContratoSrvNumSerie_Le", objItensContSrv)
        If lErro <> SUCESSO Then gError 183597
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 183598
        
    Loop

    Call Comando_Fechar(lComando)
    
    Contrato_Le_Srv = SUCESSO
    
    Exit Function

Erro_Contrato_Le_Srv:
    
    Contrato_Le_Srv = gErr
        
    Select Case gErr
    
        Case 183592
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 183593, 183594, 183598
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMCONTRATOSRV_CONTRATO", gErr, objItensContSrvAux.iFilialEmpresa, objItensContSrvAux.sCodigoContrato)
        
        Case 183595, 183596, 183597
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 183599)
            
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function Garantia_LockShared(ByVal objGarantia As ClassGarantia) As Long
'Faz lock shared na garantia passada como parâmetro


Dim lErro As Long
Dim lComando As Long
Dim lCodAux As Long
Dim iAtivo As Integer

On Error GoTo Erro_Garantia_LockShared

    'Executa a abertura dos Comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 183643

    'Busca a garantia a ser lockado
    lErro = Comando_ExecutarLockado(lComando, "SELECT Codigo, Ativo FROM Garantia WHERE Codigo=? AND FilialEmpresa=?", lCodAux, iAtivo, objGarantia.lCodigo, objGarantia.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 183644
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 183645
    
    'Se não encontrou a garantia => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 183646
    
    'Efetua o lock da garantia
    lErro = Comando_LockShared(lComando)
    If lErro <> AD_SQL_SUCESSO Then gError 183647
    
    objGarantia.iAtivo = iAtivo
    
    'Executa o fechamento dos Comandos
    Call Comando_Fechar(lComando)

    Garantia_LockShared = SUCESSO

    Exit Function

Erro_Garantia_LockShared:

    Garantia_LockShared = gErr

    Select Case gErr

        Case 183643
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 183644, 183645
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_GARANTIA", gErr, objGarantia.iFilialEmpresa, objGarantia.lCodigo)

        Case 183646

        Case 183647
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_GARANTIA", gErr, objGarantia.iFilialEmpresa, objGarantia.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183648)

    End Select

    Call Comando_Fechar(lComando)

End Function

Public Function TipoGarantia_Le(objTipoGarantia As ClassTipoGarantia) As Long

Dim lComando As Long
Dim lErro As Long
Dim tTipoGarantia As typeTipoGarantia

On Error GoTo Erro_TipoGarantia_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 183846
        
    tTipoGarantia.sDescricao = String(SRING_TIPOGARANTIA_DESCRICAO, 0)
        
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Codigo, Descricao, PrazoPadrao, GarantiaTotal, GarantiaTotalPrazo FROM TipoGarantia WHERE Codigo = ? ", _
    tTipoGarantia.lNumIntDoc, tTipoGarantia.lCodigo, tTipoGarantia.sDescricao, tTipoGarantia.iPrazoPadrao, tTipoGarantia.iGarantiaTotal, tTipoGarantia.iGarantiaTotalPrazo, objTipoGarantia.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 183847
        
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 183848
        
    If lErro = AD_SQL_SEM_DADOS Then gError 183849

    objTipoGarantia.lNumIntDoc = tTipoGarantia.lNumIntDoc
    objTipoGarantia.lCodigo = tTipoGarantia.lCodigo
    objTipoGarantia.sDescricao = tTipoGarantia.sDescricao
    objTipoGarantia.iPrazoPadrao = tTipoGarantia.iPrazoPadrao
    objTipoGarantia.iGarantiaTotal = tTipoGarantia.iGarantiaTotal
    objTipoGarantia.iGarantiaTotalPrazo = tTipoGarantia.iGarantiaTotalPrazo

    lErro = CF("TipoGarantiaProduto_Le", objTipoGarantia)
    If lErro <> SUCESSO Then gError 183850

    Call Comando_Fechar(lComando)
    
    TipoGarantia_Le = SUCESSO
    
    Exit Function

Erro_TipoGarantia_Le:
    
    TipoGarantia_Le = gErr
        
    Select Case gErr
    
        Case 183846
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 183847, 183848
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOGARANTIA", gErr, objTipoGarantia.lCodigo)
        
        Case 183849, 183850
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 183851)
            
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function TipoGarantiaProduto_Le(objTipoGarantia As ClassTipoGarantia) As Long

Dim lComando As Long
Dim lErro As Long
Dim tTipoGarantiaProduto As typeTipoGarantiaProduto
Dim objTipoGarantiaProduto As ClassTipoGarantiaProduto

On Error GoTo Erro_TipoGarantiaProduto_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 183852
        
    tTipoGarantiaProduto.sProduto = String(STRING_PRODUTO, 0)
        
    lErro = Comando_Executar(lComando, "SELECT NumIntTipoGarantia, Produto, Prazo FROM TipoGarantiaProduto WHERE NumIntTipoGarantia = ?", _
    tTipoGarantiaProduto.lNumIntTipoGarantia, tTipoGarantiaProduto.sProduto, tTipoGarantiaProduto.iPrazo, objTipoGarantia.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 183853
        
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 183854
        
    Do While lErro = AD_SQL_SUCESSO

        Set objTipoGarantiaProduto = New ClassTipoGarantiaProduto

        objTipoGarantiaProduto.lNumIntTipoGarantia = tTipoGarantiaProduto.lNumIntTipoGarantia
        objTipoGarantiaProduto.sProduto = tTipoGarantiaProduto.sProduto
        objTipoGarantiaProduto.iPrazo = tTipoGarantiaProduto.iPrazo

        objTipoGarantia.colTipoGarantiaProduto.Add objTipoGarantiaProduto
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 183855
        
    Loop

    Call Comando_Fechar(lComando)
    
    TipoGarantiaProduto_Le = SUCESSO
    
    Exit Function

Erro_TipoGarantiaProduto_Le:
    
    TipoGarantiaProduto_Le = gErr
        
    Select Case gErr
    
        Case 183852
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 183853 To 183855
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOGARANTIAPRODUTO", gErr, objTipoGarantia.lNumIntDoc)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 183856)
            
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function ItensSolicSRV_Le_Garantia(objSolicSRV As ClassSolicSRV, ByVal lGarantia As Long, ByVal sProduto As String) As Long

Dim lComando As Long
Dim lErro As Long
Dim tItensSolicSRV As typeItensSolicSRV
Dim objItensSolicSrv As ClassItensSolicSRV

On Error GoTo Erro_ItensSolicSRV_Le_Garantia

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 183943
        
    tItensSolicSRV.sProduto = String(STRING_PRODUTO, 0)
    tItensSolicSRV.sProdutoDesc = String(STRING_PRODUTO_DESCRICAO, 0)
    tItensSolicSRV.sServico = String(STRING_PRODUTO, 0)
    tItensSolicSRV.sServicoDesc = String(STRING_PRODUTO_DESCRICAO, 0)
    tItensSolicSRV.sLote = String(STRING_LOTE_RASTREAMENTO, 0)
    tItensSolicSRV.sSolicitacao = String(STRING_SOLICITACAO, 0)
    tItensSolicSRV.sContrato = String(STRING_CONTRATO_CODIGO, 0)
    tItensSolicSRV.sUM = String(STRING_UM_SIGLA, 0)
    tItensSolicSRV.sReparo = String(STRING_MAXIMO, 0)
        
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Produto, Produtos.Descricao, DataVenda, Servico, Servicos.Descricao, UM, Quantidade, Lote, FilialOP, Solicitacao, Garantia, Contrato, StatusItem, Reparo, DataBaixa FROM ItensSolicSRV, Produtos, Produtos As Servicos WHERE Garantia = ? AND Produto = ? AND ItensSolicSRV.Produto = Produtos.Codigo AND ItensSolicSRV.Servico = Servicos.Codigo", _
    tItensSolicSRV.lNumIntDoc, tItensSolicSRV.sProduto, tItensSolicSRV.sProdutoDesc, tItensSolicSRV.dtDataVenda, tItensSolicSRV.sServico, tItensSolicSRV.sServicoDesc, tItensSolicSRV.sUM, tItensSolicSRV.dQuantidade, tItensSolicSRV.sLote, tItensSolicSRV.iFilialOP, tItensSolicSRV.sSolicitacao, tItensSolicSRV.lGarantia, tItensSolicSRV.sContrato, tItensSolicSRV.iStatusItem, tItensSolicSRV.sReparo, tItensSolicSRV.dtDataBaixa, lGarantia, sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 183944
        
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 183945
        
    If lErro = AD_SQL_SEM_DADOS Then gError 183946
        
    Do While lErro = AD_SQL_SUCESSO
    
        Set objItensSolicSrv = New ClassItensSolicSRV
    
        objItensSolicSrv.lNumIntDoc = tItensSolicSRV.lNumIntDoc
        objItensSolicSrv.lNumIntSolicSRV = objSolicSRV.lNumIntDoc
        objItensSolicSrv.sProduto = tItensSolicSRV.sProduto
        objItensSolicSrv.sProdutoDesc = tItensSolicSRV.sProdutoDesc
        objItensSolicSrv.dtDataVenda = tItensSolicSRV.dtDataVenda
        objItensSolicSrv.sServico = tItensSolicSRV.sServico
        objItensSolicSrv.sServicoDesc = tItensSolicSRV.sServicoDesc
        objItensSolicSrv.sUM = tItensSolicSRV.sUM
        objItensSolicSrv.dQuantidade = tItensSolicSRV.dQuantidade
        objItensSolicSrv.sLote = tItensSolicSRV.sLote
        objItensSolicSrv.iFilialOP = tItensSolicSRV.iFilialOP
        objItensSolicSrv.sSolicitacao = tItensSolicSRV.sSolicitacao
        objItensSolicSrv.sReparo = tItensSolicSRV.sReparo
        objItensSolicSrv.lGarantia = tItensSolicSRV.lGarantia
        objItensSolicSrv.sContrato = tItensSolicSRV.sContrato
        objItensSolicSrv.iStatusItem = tItensSolicSRV.iStatusItem
        objItensSolicSrv.dtDataBaixa = tItensSolicSRV.dtDataBaixa

        objSolicSRV.colItens.Add objItensSolicSrv

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 183947

    Loop

    Call Comando_Fechar(lComando)
    
    ItensSolicSRV_Le_Garantia = SUCESSO
    
    Exit Function

Erro_ItensSolicSRV_Le_Garantia:
    
    ItensSolicSRV_Le_Garantia = gErr
        
    Select Case gErr
    
        Case 183943
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 183944, 183945, 183947
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSSOLICSRV_GARANTIA", gErr, lGarantia, sProduto)
        
        Case 183946
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 183948)
            
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function OrcamentoSRV_Le(ByVal objOrcamentoVenda As ClassOrcamentoVenda) As Long
'Lê o Orcamento de Venda a partir do Código e Filial Empresa passados.

Dim lComando As Long
Dim lErro As Long
Dim tOrcamentoVenda As typeOrcamentoVenda
Dim iPrazoValidade As Integer

On Error GoTo Erro_OrcamentoSRV_Le
    
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 186531

    tOrcamentoVenda.sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
    tOrcamentoVenda.sNomeCli = String(STRING_CLIENTE_NOME_REDUZIDO, 0)
    tOrcamentoVenda.sNomeFilialCli = String(STRING_FILIAL_CLIENTE_NOME, 0)
    tOrcamentoVenda.sUsuario = String(STRING_USUARIO, 0)
    
    lErro = Comando_Executar(lComando, "SELECT NaturezaOp, NumIntPedSRV, NumIntNFiscal, NomeCli, NomeFilialCli, Vendedor, Cliente, Filial, CondicaoPagto, PercAcrescFinanceiro, DataEmissao,  ValorTotal, ValorFrete, ValorDesconto, ValorSeguro, TabelaPreco, DataReferencia, NumIntDoc,ValorOutrasDespesas,ValorProdutos, PrazoValidade, Usuario, CodigoBase, Status, MotivoPerda, NumIntSolicSRV, ValorDescontoTit, ValorItens, ValorDescontoItens FROM OrcamentoSRV WHERE FilialEmpresa = ? AND Codigo = ? ", _
        tOrcamentoVenda.sNaturezaOp, tOrcamentoVenda.lNumIntPedVenda, tOrcamentoVenda.lNumIntNFiscal, tOrcamentoVenda.sNomeCli, tOrcamentoVenda.sNomeFilialCli, tOrcamentoVenda.iVendedor, tOrcamentoVenda.lCliente, tOrcamentoVenda.iFilial, tOrcamentoVenda.iCondicaoPagto, tOrcamentoVenda.dPercAcrescFinanceiro, tOrcamentoVenda.dtDataEmissao, tOrcamentoVenda.dValorTotal, tOrcamentoVenda.dValorFrete, tOrcamentoVenda.dValorDesconto, tOrcamentoVenda.dValorSeguro, tOrcamentoVenda.iTabelaPreco, tOrcamentoVenda.dtDataReferencia, tOrcamentoVenda.lNumIntDoc, tOrcamentoVenda.dValorOutrasDespesas, tOrcamentoVenda.dValorProdutos, iPrazoValidade, tOrcamentoVenda.sUsuario, tOrcamentoVenda.lCodigoBase, tOrcamentoVenda.lStatus, tOrcamentoVenda.lMotivoPerda, tOrcamentoVenda.lNumIntSolicSRV, _
        tOrcamentoVenda.dValorDescontoTit, tOrcamentoVenda.dValorItens, tOrcamentoVenda.dValorDescontoItens, objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 186532

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 186533
    
    If lErro = AD_SQL_SEM_DADOS Then gError 186534

    objOrcamentoVenda.lNumIntPedVenda = tOrcamentoVenda.lNumIntPedVenda
    objOrcamentoVenda.lNumIntNFiscal = tOrcamentoVenda.lNumIntNFiscal
    objOrcamentoVenda.sNomeFilialCli = tOrcamentoVenda.sNomeFilialCli
    objOrcamentoVenda.sNomeCli = tOrcamentoVenda.sNomeCli
    objOrcamentoVenda.iPrazoValidade = iPrazoValidade
    objOrcamentoVenda.lCliente = tOrcamentoVenda.lCliente
    objOrcamentoVenda.iFilial = tOrcamentoVenda.iFilial
    objOrcamentoVenda.iCondicaoPagto = tOrcamentoVenda.iCondicaoPagto
    objOrcamentoVenda.dPercAcrescFinanceiro = tOrcamentoVenda.dPercAcrescFinanceiro
    objOrcamentoVenda.dtDataEmissao = tOrcamentoVenda.dtDataEmissao
    objOrcamentoVenda.sNaturezaOp = tOrcamentoVenda.sNaturezaOp
    objOrcamentoVenda.dValorTotal = tOrcamentoVenda.dValorTotal
    objOrcamentoVenda.dValorFrete = tOrcamentoVenda.dValorFrete
    objOrcamentoVenda.dValorDesconto = tOrcamentoVenda.dValorDesconto
    objOrcamentoVenda.dValorOutrasDespesas = tOrcamentoVenda.dValorOutrasDespesas
    objOrcamentoVenda.dValorSeguro = tOrcamentoVenda.dValorSeguro
    objOrcamentoVenda.iTabelaPreco = tOrcamentoVenda.iTabelaPreco
    objOrcamentoVenda.dtDataReferencia = tOrcamentoVenda.dtDataReferencia
    objOrcamentoVenda.lNumIntDoc = tOrcamentoVenda.lNumIntDoc
    objOrcamentoVenda.dValorProdutos = tOrcamentoVenda.dValorProdutos
    objOrcamentoVenda.iVendedor = tOrcamentoVenda.iVendedor
    objOrcamentoVenda.sUsuario = tOrcamentoVenda.sUsuario
    objOrcamentoVenda.lCodigoBase = tOrcamentoVenda.lCodigoBase
    objOrcamentoVenda.lStatus = tOrcamentoVenda.lStatus
    objOrcamentoVenda.lMotivoPerda = tOrcamentoVenda.lMotivoPerda
    objOrcamentoVenda.lNumIntSolicSRV = tOrcamentoVenda.lNumIntSolicSRV
    objOrcamentoVenda.dValorDescontoTit = tOrcamentoVenda.dValorDescontoTit
    objOrcamentoVenda.dValorDescontoItens = tOrcamentoVenda.dValorDescontoItens
    objOrcamentoVenda.dValorItens = tOrcamentoVenda.dValorItens
    
    Call Comando_Fechar(lComando)

    OrcamentoSRV_Le = SUCESSO

    Exit Function

Erro_OrcamentoSRV_Le:

    OrcamentoSRV_Le = gErr

    Select Case gErr

        Case 186531
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 186532, 186533
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORCAMENTOSRV", gErr, objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo)
        
        Case 186534

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 186535)
    
    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function OrcamentoSRV_Le_Itens_ComTributacao(ByVal objOrcamentoVenda As ClassOrcamentoVenda) As Long

Dim lComando As Long, tItemPedido As typeItemPedido, objItemOV As ClassItemOV
'Dim tTributacaoItemOV As typeTributacaoItemPV, objTributacaoItemOV As ClassTributacaoItemPV
Dim lErro As Long

Dim sVersaoKit As String
Dim sVersaoKitBase As String
Dim lStatus As Long
Dim lMotivoPerda As Long
Dim sObservacao As String
Dim iPeca As Integer
Dim lNumIntItensSolicSRV As Long
Dim lNumIntItemContratoSrv As Long
Dim lNumIntGarantia As Long
Dim objGarantia As New ClassGarantia
Dim objContrato As New ClassContrato
Dim objItemOrcamento As ClassItemOV
Dim objSolicSRV As New ClassSolicSRV
Dim objItensSolicSrv As New ClassItensSolicSRV

On Error GoTo Erro_OrcamentoSRV_Le_Itens_ComTributacao

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 186599

    tItemPedido.sDescricao = String(STRING_ITEM_PEDIDO_DESCRICAO, 0)
    tItemPedido.sLote = String(STRING_ITEM_PEDIDO_LOTE, 0)
    tItemPedido.sProduto = String(STRING_PRODUTO, 0)
    tItemPedido.sProdutoDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
    tItemPedido.sProdutoNomeReduzido = String(STRING_PRODUTO_NOME_REDUZIDO, 0)
    tItemPedido.sUMEstoque = String(STRING_PRODUTO_SIGLAUMESTOQUE, 0)
    tItemPedido.sUnidadeMed = String(STRING_UM_SIGLA, 0)

    sVersaoKit = String(STRING_KIT_VERSAO, 0)
    sVersaoKitBase = String(STRING_KIT_VERSAO, 0)

    sObservacao = String(STRING_OBSERVACAO_OBSERVACAO, 0)
'
'    With tTributacaoItemOV
'
'        .sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
'
'        lErro = Comando_Executar(lComando, "SELECT ItensOrcamentoSRV.NumIntDoc,ItensOrcamentoSRV.Produto, Quantidade, PrecoUnitario, PrecoTotal, ItensOrcamentoSRV.ClasseUM, UnidadeMed, ValorDesconto, DataEntrega, ItensOrcamentoSRV.Descricao, " & _
'            "ItensOrcamentoSRV.Descricao, NomeReduzido, SiglaUMEstoque, ControleEstoque, " & _
'            "NaturezaOp, NaturezaOpManual, TipoTributacao, TipoTributacaoManual, IPITipo, IPITipoManual, IPIBaseCalculo, IPIBaseManual, IPIPercRedBase, IPIPercRedBaseManual, TributacaoItemOrcSRV.IPIAliquota, IPIAliquotaManual, IPIValor, IPIValorManual, ICMSTipo, ICMSTipoManual, ICMSBase, ICMSBaseManual, ICMSPercRedBase, ICMSPercRedBaseManual, ICMSAliquota, ICMSAliquotaManual, ICMSValor, ICMSValorManual, " & _
'            "ICMSSubstBase, ICMSSubstBaseManual, ICMSSubstAliquota, ICMSSubstAliquotaManual, ICMSSubstValor, ICMSSubstValorManual, ItensOrcamentoSRV.VersaoKit, ItensOrcamentoSRV.VersaoKitBase, ItensOrcamentoSRV.Status, ItensOrcamentoSRV.MotivoPerda, ItensOrcamentoSRV.Observacao, Peca, PISCRedito, COFINSCredito, IPICredito, ICMSCredito " & _
'            "FROM ItensOrcamentoSRV, Produtos, TributacaoItemOrcSRV WHERE ItensOrcamentoSRV.FilialEmpresa = ? AND ItensOrcamentoSRV.CodOrcamento = ? AND ItensOrcamentoSRV.Produto = Produtos.Codigo AND ItensOrcamentoSRV.NumIntDoc = TributacaoItemOrcSRV.NumIntDoc ORDER BY ItensOrcamentoSRV.NumIntDoc", _
'            tItemPedido.lNumIntDoc, tItemPedido.sProduto, tItemPedido.dQuantidade, tItemPedido.dPrecoUnitario, tItemPedido.dPrecoTotal, tItemPedido.iClasseUM, tItemPedido.sUnidadeMed, tItemPedido.dValorDesconto, tItemPedido.dtDataEntrega, _
'            tItemPedido.sDescricao, tItemPedido.sProdutoDescricao, _
'            tItemPedido.sProdutoNomeReduzido, tItemPedido.sUMEstoque, tItemPedido.iControleEstoque, .sNaturezaOp, .iNaturezaOpManual, .iTipoTributacao, .iTipoTributacaoManual, .iIPITipo, .iIPITipoManual, .dIPIBaseCalculo, .iIPIBaseManual, .dIPIPercRedBase, .iIPIPercRedBaseManual, .dIPIAliquota, .iIPIAliquotaManual, .dIPIValor, _
'            .iIPIValorManual, .iICMSTipo, .iICMSTipoManual, .dICMSBase, .iICMSBaseManual, .dICMSPercRedBase, .iICMSPercRedBaseManual, .dICMSAliquota, .iICMSAliquotaManual, .dICMSValor, .iICMSValorManual, .dICMSSubstBase, .iICMSSubstBaseManual, .dICMSSubstAliquota, .iICMSSubstAliquotaManual, .dICMSSubstValor, .iICMSSubstValorManual, sVersaoKit, sVersaoKitBase, lStatus, lMotivoPerda, sObservacao, iPeca, .dPISCredito, .dCOFINSCredito, .dIPICredito, .dICMSCredito, objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo)
'        If lErro <> AD_SQL_SUCESSO Then gError 186600
'
'    End With
    
    lErro = Comando_Executar(lComando, "SELECT ItensOrcamentoSRV.NumIntDoc,ItensOrcamentoSRV.Produto, Quantidade, PrecoUnitario, PrecoTotal, ItensOrcamentoSRV.ClasseUM, UnidadeMed, ValorDesconto, DataEntrega, ItensOrcamentoSRV.Descricao, " & _
        "ItensOrcamentoSRV.Descricao, NomeReduzido, SiglaUMEstoque, ControleEstoque, " & _
        "ItensOrcamentoSRV.VersaoKit, ItensOrcamentoSRV.VersaoKitBase, ItensOrcamentoSRV.Status, ItensOrcamentoSRV.MotivoPerda, ItensOrcamentoSRV.Observacao, Peca " & _
        "FROM ItensOrcamentoSRV, Produtos WHERE ItensOrcamentoSRV.FilialEmpresa = ? AND ItensOrcamentoSRV.CodOrcamento = ? AND ItensOrcamentoSRV.Produto = Produtos.Codigo ORDER BY ItensOrcamentoSRV.NumIntDoc", _
        tItemPedido.lNumIntDoc, tItemPedido.sProduto, tItemPedido.dQuantidade, tItemPedido.dPrecoUnitario, tItemPedido.dPrecoTotal, tItemPedido.iClasseUM, tItemPedido.sUnidadeMed, tItemPedido.dValorDesconto, tItemPedido.dtDataEntrega, _
        tItemPedido.sDescricao, tItemPedido.sProdutoDescricao, _
        tItemPedido.sProdutoNomeReduzido, tItemPedido.sUMEstoque, tItemPedido.iControleEstoque, sVersaoKit, sVersaoKitBase, lStatus, lMotivoPerda, sObservacao, iPeca, objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 186600

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 186601

    If lErro = AD_SQL_SEM_DADOS Then gError 186602

    Do While lErro <> AD_SQL_SEM_DADOS
'
'        Set objTributacaoItemOV = New ClassTributacaoItemPV
'
'        With objTributacaoItemOV
'
'            .lNumIntDoc = tTributacaoItemOV.lNumIntDoc
'            .sNaturezaOp = tTributacaoItemOV.sNaturezaOp
'            .iNaturezaOpManual = tTributacaoItemOV.iNaturezaOpManual
'            .iTipoTributacao = tTributacaoItemOV.iTipoTributacao
'            .iTipoTributacaoManual = tTributacaoItemOV.iTipoTributacaoManual
'            .iIPITipo = tTributacaoItemOV.iIPITipo
'            .iIPITipoManual = tTributacaoItemOV.iIPITipoManual
'            .dIPIBaseCalculo = tTributacaoItemOV.dIPIBaseCalculo
'            .iIPIBaseManual = tTributacaoItemOV.iIPIBaseManual
'            .dIPIPercRedBase = tTributacaoItemOV.dIPIPercRedBase
'            .iIPIPercRedBaseManual = tTributacaoItemOV.iIPIPercRedBaseManual
'            .dIPIAliquota = tTributacaoItemOV.dIPIAliquota
'            .iIPIAliquotaManual = tTributacaoItemOV.iIPIAliquotaManual
'            .dIPIValor = tTributacaoItemOV.dIPIValor
'            .iIPIValorManual = tTributacaoItemOV.iIPIValorManual
'            .iICMSTipo = tTributacaoItemOV.iICMSTipo
'            .iICMSTipoManual = tTributacaoItemOV.iICMSTipoManual
'            .dICMSBase = tTributacaoItemOV.dICMSBase
'            .iICMSBaseManual = tTributacaoItemOV.iICMSBaseManual
'            .dICMSPercRedBase = tTributacaoItemOV.dICMSPercRedBase
'            .iICMSPercRedBaseManual = tTributacaoItemOV.iICMSPercRedBaseManual
'            .dICMSAliquota = tTributacaoItemOV.dICMSAliquota
'            .iICMSAliquotaManual = tTributacaoItemOV.iICMSAliquotaManual
'            .dICMSValor = tTributacaoItemOV.dICMSValor
'            .iICMSValorManual = tTributacaoItemOV.iICMSValorManual
'            .dICMSSubstBase = tTributacaoItemOV.dICMSSubstBase
'            .iICMSSubstBaseManual = tTributacaoItemOV.iICMSSubstBaseManual
'            .dICMSSubstAliquota = tTributacaoItemOV.dICMSSubstAliquota
'            .iICMSSubstAliquotaManual = tTributacaoItemOV.iICMSSubstAliquotaManual
'            .dICMSSubstValor = tTributacaoItemOV.dICMSSubstValor
'            .iICMSSubstValorManual = tTributacaoItemOV.iICMSSubstValorManual
'            .dICMSCredito = tTributacaoItemOV.dICMSCredito
'            .dIPICredito = tTributacaoItemOV.dIPICredito
'            .dCOFINSCredito = tTributacaoItemOV.dCOFINSCredito
'            .dPISCredito = tTributacaoItemOV.dPISCredito
'        End With
'
'
        Set objItemOV = New ClassItemOV

        objItemOV.sProduto = tItemPedido.sProduto
        objItemOV.dQuantidade = tItemPedido.dQuantidade
        objItemOV.dPrecoUnitario = tItemPedido.dPrecoUnitario
        objItemOV.dPrecoTotal = tItemPedido.dPrecoTotal
        objItemOV.dValorDesconto = tItemPedido.dValorDesconto
        objItemOV.dtDataEntrega = tItemPedido.dtDataEntrega
        objItemOV.sProdutoDescricao = tItemPedido.sProdutoDescricao
        objItemOV.sUnidadeMed = tItemPedido.sUnidadeMed
        objItemOV.iClasseUM = tItemPedido.iClasseUM
        objItemOV.sDescricao = tItemPedido.sDescricao
        objItemOV.iStatus = tItemPedido.iStatus
        objItemOV.lNumIntDoc = tItemPedido.lNumIntDoc
        objItemOV.lMotivoPerda = lMotivoPerda
        objItemOV.sObservacao = sObservacao
        objItemOV.sVersaoKit = sVersaoKit
        objItemOV.sVersaoKitBase = sVersaoKitBase

        objItemOV.objTributacao.iTipoDoc = TIPODOC_TRIB_OSRV
        objItemOV.objTributacao.lNumIntDocItem = objItemOV.lNumIntDoc
        lErro = CF("TributacaoDocItem_Le", objItemOV.objTributacao)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 186603
            
        If iPeca = 0 Then
            objOrcamentoVenda.colItensSRV.Add objItemOV
        Else
            objOrcamentoVenda.colItens.Add objItemOV
        End If

        lErro = CF("ProdSolicSRV_Le", objItemOV, objOrcamentoVenda.colProdSolicSRV)
        If lErro <> SUCESSO Then gError 186603

        lErro = CF("GarantiaContratoSRV_Le", objItemOV.lNumIntDoc, objOrcamentoVenda.colGarantiaContratoSRV)
        If lErro <> SUCESSO Then gError 186604

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 186606

    Loop

    Call Comando_Fechar(lComando)
    
    OrcamentoSRV_Le_Itens_ComTributacao = SUCESSO

    Exit Function

Erro_OrcamentoSRV_Le_Itens_ComTributacao:

    OrcamentoSRV_Le_Itens_ComTributacao = gErr

    Select Case gErr

        Case 186599
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 186600, 186601, 186606
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSORCAMENTOSRV", gErr, objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo)

        Case 186602, 186603, 186604
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 186607)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function Garantia_Le_NumIntDoc(objGarantia As ClassGarantia) As Long

Dim lComando As Long
Dim lErro As Long
Dim tGarantia As typeGarantia

On Error GoTo Erro_Garantia_Le_NumIntDoc

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 186536
        
    tGarantia.sProduto = String(STRING_PRODUTO, 0)
    tGarantia.sLote = String(STRING_LOTE_RASTREAMENTO, 0)
    tGarantia.sSerie = String(STRING_SERIE, 0)
        
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Codigo, FilialEmpresa, Produto, DataVenda, Quantidade, Lote, FilialOP, Serie, NumNotaFiscal, TipoGarantia, GarantiaTotal, GarantiaTotalPrazo, Fornecedor, FilialFornecedor FROM Garantia WHERE NumIntDoc = ?", _
    tGarantia.lNumIntDoc, tGarantia.lCodigo, tGarantia.iFilialEmpresa, tGarantia.sProduto, tGarantia.dtDataVenda, tGarantia.dQuantidade, tGarantia.sLote, tGarantia.iFilialOP, tGarantia.sSerie, tGarantia.lNumNotaFiscal, tGarantia.lTipoGarantia, tGarantia.iGarantiaTotal, tGarantia.iGarantiaTotalPrazo, tGarantia.lFornecedor, tGarantia.iFilialFornecedor, objGarantia.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 186537
        
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 186538
        
    If lErro = AD_SQL_SEM_DADOS Then gError 186539

    objGarantia.lNumIntDoc = tGarantia.lNumIntDoc
    objGarantia.lCodigo = tGarantia.lCodigo
    objGarantia.iFilialEmpresa = tGarantia.iFilialEmpresa
    objGarantia.sProduto = tGarantia.sProduto
    objGarantia.dtDataVenda = tGarantia.dtDataVenda
    objGarantia.dQuantidade = tGarantia.dQuantidade
    objGarantia.sLote = tGarantia.sLote
    objGarantia.iFilialOP = tGarantia.iFilialOP
    objGarantia.sSerie = tGarantia.sSerie
    objGarantia.lNumNotaFiscal = tGarantia.lNumNotaFiscal
    objGarantia.lTipoGarantia = tGarantia.lTipoGarantia
    objGarantia.iGarantiaTotal = tGarantia.iGarantiaTotal
    objGarantia.iGarantiaTotalPrazo = tGarantia.iGarantiaTotalPrazo
    objGarantia.lFornecedor = tGarantia.lFornecedor
    objGarantia.iFilialFornecedor = tGarantia.iFilialFornecedor

    lErro = CF("GarantiaProduto_Le", objGarantia)
    If lErro <> SUCESSO Then gError 186540

    lErro = CF("GarantiaNumSerie_Le", objGarantia)
    If lErro <> SUCESSO Then gError 186541

    Call Comando_Fechar(lComando)
    
    Garantia_Le_NumIntDoc = SUCESSO
    
    Exit Function

Erro_Garantia_Le_NumIntDoc:
    
    Garantia_Le_NumIntDoc = gErr
        
    Select Case gErr
    
        Case 186536
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 186537, 186538
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_GARANTIA1", gErr, objGarantia.lNumIntDoc)
        
        Case 186539 To 186541
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 186542)
            
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function Contrato_Le_NumIntItemContrato(lNumIntItemContrato As Long, objContrato As ClassContrato) As Long

Dim lComando As Long
Dim lErro As Long
Dim tContrato As typeContrato

On Error GoTo Erro_Contrato_Le_NumIntItemContrato

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 186574
        
    tContrato.sCodigo = String(STRING_CONTRATO_CODIGO, 0)
        
    lErro = Comando_Executar(lComando, "SELECT Contratos.FilialEmpresa, Contratos.Codigo, Contratos.DataIniContrato, Contratos.DataFimContrato FROM Contratos, ItensDeContrato, ItensDeContratoSrv WHERE " & _
    "Contratos.NumIntDoc = ItensDeContrato.NumIntContrato AND ItensDeContrato.NumIntDoc = ItensDeContratoSrv.NumIntItemContrato AND ItensDeContratoSrv.NumIntItemContrato = ? ", _
    tContrato.iFilialEmpresa, tContrato.sCodigo, tContrato.dtDataIniContrato, tContrato.dtDataFimContrato, lNumIntItemContrato)
    If lErro <> AD_SQL_SUCESSO Then gError 186575
        
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 186576
        
    If lErro = AD_SQL_SEM_DADOS Then gError 186577
    
    objContrato.iFilialEmpresa = tContrato.iFilialEmpresa
    objContrato.sCodigo = tContrato.sCodigo
    objContrato.dtDataIniContrato = tContrato.dtDataIniContrato
    objContrato.dtDataFimContrato = tContrato.dtDataFimContrato

    Call Comando_Fechar(lComando)
    
    Contrato_Le_NumIntItemContrato = SUCESSO
    
    Exit Function

Erro_Contrato_Le_NumIntItemContrato:
    
    Contrato_Le_NumIntItemContrato = gErr
        
    Select Case gErr
    
        Case 186574
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 186575, 186576
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTRATO_NUMINTITEMCONTRATO", gErr)
        
        Case 186577
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 186578)
            
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function SolicSRV_Le_NumIntSolicSRV(ByVal lNumIntSolicSRV As Long, objSolicSRV As ClassSolicSRV, objItensSolicSrv As ClassItensSolicSRV) As Long

Dim lComando As Long
Dim lErro As Long
Dim tSolicSRV As typeSolicSRV
Dim sProdutoDesc As String
Dim sProduto As String

On Error GoTo Erro_SolicSRV_Le_NumIntSolicSRV

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 186579
    
    sProduto = String(STRING_PRODUTO, 0)
    sProdutoDesc = String(STRING_PRODUTO_DESCRICAO, 0)
    tSolicSRV.sObs = String(STRING_MAXIMO, 0)
    
    lErro = Comando_Executar(lComando, "SELECT Produto, Produtos.Descricao, FilialEmpresa, Codigo, Data, Hora, Cliente, FilialCliente, NumIntDoc, Vendedor, Atendente, SolicitacaoSRV.Prazo, SolicitacaoSRV.PrazoTipo, SolicitacaoSRV.Obs, SolicitacaoSRV.DataEntrega, SolicitacaoSRV.Tipo, SolicitacaoSRV.Fase FROM ItensSolicSRV, SolicitacaoSRV, Produtos WHERE SolicitacaoSRV.NumIntDoc = ItensSolicSRV.NumIntSolicSRV AND ItensSolicSRV.Produto = Produtos.Codigo AND ItensSolicSRV.NumIntDoc = ?", _
    sProduto, sProdutoDesc, tSolicSRV.iFilialEmpresa, tSolicSRV.lCodigo, tSolicSRV.dtData, tSolicSRV.dHora, tSolicSRV.lCliente, tSolicSRV.iFilial, tSolicSRV.lNumIntDoc, tSolicSRV.iVendedor, tSolicSRV.iAtendente, tSolicSRV.iPrazo, tSolicSRV.iPrazoTipo, tSolicSRV.sObs, tSolicSRV.dtDataEntrega, tSolicSRV.lTipo, tSolicSRV.lFase, lNumIntSolicSRV)
    If lErro <> AD_SQL_SUCESSO Then gError 186580
        
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 186581
        
    If lErro = AD_SQL_SEM_DADOS Then gError 186582

    objSolicSRV.iFilialEmpresa = tSolicSRV.iFilialEmpresa
    objSolicSRV.lCodigo = tSolicSRV.lCodigo
    objSolicSRV.dtData = tSolicSRV.dtData
    objSolicSRV.dtHora = tSolicSRV.dHora
    objSolicSRV.lCliente = tSolicSRV.lCliente
    objSolicSRV.iFilial = tSolicSRV.iFilial
    objSolicSRV.lNumIntDoc = tSolicSRV.lNumIntDoc
    objSolicSRV.iVendedor = tSolicSRV.iVendedor
    objSolicSRV.iAtendente = tSolicSRV.iAtendente
    objSolicSRV.iPrazo = tSolicSRV.iPrazo
    objSolicSRV.iPrazoTipo = tSolicSRV.iPrazoTipo
    objSolicSRV.sObs = tSolicSRV.sObs
    objSolicSRV.dtDataEntrega = tSolicSRV.dtDataEntrega
    objSolicSRV.lTipo = tSolicSRV.lTipo
    objSolicSRV.lFase = tSolicSRV.lFase
    
    objItensSolicSrv.sProduto = sProduto
    objItensSolicSrv.sProdutoDesc = sProdutoDesc

    Call Comando_Fechar(lComando)
    
    SolicSRV_Le_NumIntSolicSRV = SUCESSO
    
    Exit Function

Erro_SolicSRV_Le_NumIntSolicSRV:
    
    SolicSRV_Le_NumIntSolicSRV = gErr
        
    Select Case gErr
    
        Case 186579
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 186580, 186581
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SOLICITACAOSRV1", gErr)
        
        Case 186582
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 186583)
            
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function ProdSolicSRV_Le(ByVal objItemOV As ClassItemOV, ByVal colProdSolicSRV As Collection) As Long

Dim lComando As Long
Dim lErro As Long
Dim tProdSolicSRV As typeProdSolicSRV
Dim sProdutoDesc As String
Dim sProduto As String
Dim objProdSolicSRV As ClassProdSolicSRV

On Error GoTo Erro_ProdSolicSRV_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 186584
    
    tProdSolicSRV.sProduto = String(STRING_PRODUTO, 0)
    tProdSolicSRV.sLote = String(STRING_LOTE_RASTREAMENTO, 0)
    tProdSolicSRV.sContrato = String(STRING_CONTRATO_CODIGO, 0)
    
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Produto, Quantidade, Lote, FilialOP, Garantia, Contrato  FROM ProdSolicSRV WHERE NumIntItensOrcSRV = ?", _
    tProdSolicSRV.lNumIntDoc, tProdSolicSRV.sProduto, tProdSolicSRV.dQuantidade, tProdSolicSRV.sLote, tProdSolicSRV.iFilialOP, tProdSolicSRV.lGarantia, tProdSolicSRV.sContrato, objItemOV.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 186585
        
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 186586
        
    Do While lErro = AD_SQL_SUCESSO

        Set objProdSolicSRV = New ClassProdSolicSRV

        objProdSolicSRV.lNumIntDoc = tProdSolicSRV.lNumIntDoc
        objProdSolicSRV.lNumIntItensOrcSRV = objItemOV.lNumIntDoc
        objProdSolicSRV.sServicoOrcSRV = objItemOV.sProduto
        objProdSolicSRV.sProduto = tProdSolicSRV.sProduto
        objProdSolicSRV.dQuantidade = tProdSolicSRV.dQuantidade
        objProdSolicSRV.sLote = tProdSolicSRV.sLote
        objProdSolicSRV.iFilialOP = tProdSolicSRV.iFilialOP
        objProdSolicSRV.lGarantia = tProdSolicSRV.lGarantia
        objProdSolicSRV.sContrato = tProdSolicSRV.sContrato
        
        colProdSolicSRV.Add objProdSolicSRV
        
        lErro = CF("ProdutoSRV_Le", objProdSolicSRV)
        If lErro <> SUCESSO Then gError 186605
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 186587

    Loop

    Call Comando_Fechar(lComando)
    
    ProdSolicSRV_Le = SUCESSO
    
    Exit Function

Erro_ProdSolicSRV_Le:
    
    ProdSolicSRV_Le = gErr
        
    Select Case gErr
    
        Case 186584
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 186585, 186586, 186587
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODSOLICSRV", objItemOV.lNumIntDoc)
        
        Case 186605
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 186588)
            
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function GarantiaContratoSRV_Le(ByVal lNumIntItensOrcSRV As Long, ByVal colGarantiaContratoSRV As Collection) As Long

Dim lComando As Long
Dim lErro As Long
Dim tGarantiaContratoSRV As typeGarantiaContratoSRV
Dim objGarantiaContratoSRV As ClassGarantiaContratoSRV

On Error GoTo Erro_GarantiaContratoSRV_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 186589
    
    tGarantiaContratoSRV.sContratoCod = String(STRING_CONTRATO_CODIGO, 0)
    
    lErro = Comando_Executar(lComando, "SELECT GarantiaContratoSRV.NumIntGarantia, GarantiaContratoSRV.NumIntItensContratoSRV, GarantiaContratoSRV.Quantidade, Contratos.Codigo, Garantia.Codigo FROM GarantiaContratoSRV LEFT OUTER JOIN Garantia " & _
                        "ON Garantia.NumIntDoc = GarantiaContratoSRV.NumIntGarantia LEFT OUTER JOIN " & _
                      "ItensDeContrato INNER JOIN ItensDeContratoSrv ON ItensDeContrato.NumIntDoc = ItensDeContratoSrv.NumIntItemContrato INNER JOIN " & _
                      "Contratos ON ItensDeContrato.NumIntContrato = Contratos.NumIntDoc ON GarantiaContratoSRV.NumIntItensContratoSRV = ItensDeContratoSrv.NumIntDoc WHERE NumIntItensOrcSRV = ?", _
                        tGarantiaContratoSRV.lNumIntGarantia, tGarantiaContratoSRV.lNumIntItensContratoSRV, tGarantiaContratoSRV.dQuantidade, tGarantiaContratoSRV.sContratoCod, tGarantiaContratoSRV.lGarantiaCod, lNumIntItensOrcSRV)
    If lErro <> AD_SQL_SUCESSO Then gError 186590
        
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 186591
        
    Do While lErro = AD_SQL_SUCESSO

        Set objGarantiaContratoSRV = New ClassGarantiaContratoSRV

        objGarantiaContratoSRV.lNumIntItensOrcSRV = lNumIntItensOrcSRV
        objGarantiaContratoSRV.lNumIntItensContratoSRV = tGarantiaContratoSRV.lNumIntItensContratoSRV
        objGarantiaContratoSRV.lNumIntGarantia = tGarantiaContratoSRV.lNumIntGarantia
        objGarantiaContratoSRV.dQuantidade = tGarantiaContratoSRV.dQuantidade
        objGarantiaContratoSRV.sContratoCod = tGarantiaContratoSRV.sContratoCod
        objGarantiaContratoSRV.lGarantiaCod = tGarantiaContratoSRV.lGarantiaCod
        
        colGarantiaContratoSRV.Add objGarantiaContratoSRV
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 186592

    Loop

    Call Comando_Fechar(lComando)
    
    GarantiaContratoSRV_Le = SUCESSO
    
    Exit Function

Erro_GarantiaContratoSRV_Le:
    
    GarantiaContratoSRV_Le = gErr
        
    Select Case gErr
    
        Case 186589
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 186590, 186591, 186592
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_GARANTIACONTRATOSRV", lNumIntItensOrcSRV)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 186593)
            
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function ProdutoSRV_Le(ByVal objProdSolicSRV As ClassProdSolicSRV) As Long

Dim lComando As Long
Dim lErro As Long
Dim tProdutoSRV As typeProdutoSRV
Dim sProduto As String
Dim objProdutoSRV As ClassProdutoSRV

On Error GoTo Erro_ProdutoSRV_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 186594
    
    tProdutoSRV.sProduto = String(STRING_PRODUTO, 0)
    tProdutoSRV.sContrato = String(STRING_CONTRATO_CODIGO, 0)
    
    lErro = Comando_Executar(lComando, "SELECT Produto, Quantidade, Garantia, Contrato FROM ProdutoSRV WHERE NumIntProdSolicSRV = ?", _
    tProdutoSRV.sProduto, tProdutoSRV.dQuantidade, tProdutoSRV.lGarantia, tProdutoSRV.sContrato, objProdSolicSRV.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 186595
        
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 186596
        
    Do While lErro = AD_SQL_SUCESSO

        Set objProdutoSRV = New ClassProdutoSRV

        objProdutoSRV.lNumIntProdSolicSRV = objProdSolicSRV.lNumIntDoc
        objProdutoSRV.sProduto = tProdutoSRV.sProduto
        objProdutoSRV.dQuantidade = tProdutoSRV.dQuantidade
        objProdutoSRV.lGarantia = tProdutoSRV.lGarantia
        objProdutoSRV.sContrato = tProdutoSRV.sContrato
        
        objProdSolicSRV.colProdutoSRV.Add objProdutoSRV
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 186597

    Loop

    Call Comando_Fechar(lComando)
    
    ProdutoSRV_Le = SUCESSO
    
    Exit Function

Erro_ProdutoSRV_Le:
    
    ProdutoSRV_Le = gErr
        
    Select Case gErr
    
        Case 186594
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 186595, 186596, 186597
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOSRV", objProdSolicSRV.lNumIntDoc)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 186598)
            
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function OrcamentoSRV_Le_Tributacao(ByVal objOrcamentoVenda As ClassOrcamentoVenda) As Long
'Le a Tributacao do Orcamento de Servico Passado como parametro.

Dim lErro As Long
'Dim lComando As Long
'Dim tTributacaoOV As typeTributacaoOV
'Dim objTributacaoOV As New ClassTributacaoOV
'Dim objTribComplOV As ClassTributacaoItemPV

On Error GoTo Erro_OrcamentoSRV_Le_Tributacao

'    lComando = Comando_Abrir()
'    If lComando = 0 Then gError 186620
'
'    With tTributacaoOV
'        lErro = Comando_Executar(lComando, "SELECT FilialEmpresa, CodOrcamento, TaxacaoAutomatica, TipoTributacao, TipoTributacaoManual, IPIBase, IPIBaseManual, IPIValor, IPIValorManual, ICMSBase, ICMSBaseManual, ICMSValor, ICMSValorManual, ICMSSubstBase, ICMSSubstBaseManual, ICMSSubstValor, ICMSSubstValorManual, ISSIncluso, ISSBase, ISSAliquota, ISSAliquotaManual, ISSValor, ISSValorManual, IRRFBase, IRRFAliquota, IRRFAliquotaManual, IRRFValor, IRRFValorManual, PISRetidoManual, COFINSRetidoManual, CSLLRetidoManual, PISRetido, COFINSRetido, CSLLRetido, " & _
'        "PISCredito,PISCreditoManual,COFINSCredito,COFINSCreditoManual,IPICredito,IPICreditoManual,ICMSCredito,ICMSCreditoManual,ValorINSS,ValorINSSManual,INSSRetido,INSSRetidoManual,INSSValorBase,INSSValorBaseManual,INSSValorDeducoes ,INSSValorDeducoesManual FROM TributacaoOrcSRV WHERE FilialEmpresa = ? AND CodOrcamento = ?", _
'            .iFilialEmpresa, .lCodOrcamento, .iTaxacaoAutomatica, .iTipoTributacao, .iTipoTributacaoManual, .dIPIBase, .iIPIBaseManual, .dIPIValor, .iIPIValorManual, .dICMSBase, .iICMSBaseManual, .dICMSValor, .iICMSValorManual, .dICMSSubstBase, .iICMSSubstBaseManual, .dICMSSubstValor, .iICMSSubstValorManual, .iISSIncluso, .dISSBase, .dISSAliquota, .iISSAliquotaManual, .dISSValor, .iISSValorManual, .dIRRFBase, .dIRRFAliquota, .iIRRFAliquotaManual, .dIRRFValor, .iIRRFValorManual, .iPISRetidoManual, .iCOFINSRetidoManual, .iCSLLRetidoManual, .dPISRetido, .dCOFINSRetido, .dCSLLRetido, _
'            .dPISCredito, .iPISCreditoManual, .dCOFINSCredito, .iCOFINSCreditoManual, .dIPICredito, .iIPICreditoManual, .dICMSCredito, .iICMSCreditoManual, .dValorINSS, .iINSSValorManual, .iINSSRetido, .iINSSRetidoManual, .dINSSBase, .iINSSBaseManual, .dINSSDeducoes, .iINSSDeducoesManual, _
'            objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo)
'    End With
'
'    If lErro <> AD_SQL_SUCESSO Then gError 186621
'
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 186622
'
'    If lErro = AD_SQL_SEM_DADOS Then gError 186623
'
'    With objTributacaoOV
'        .dICMSBase = tTributacaoOV.dICMSBase
'        .dICMSSubstBase = tTributacaoOV.dICMSSubstBase
'        .dICMSSubstValor = tTributacaoOV.dICMSSubstValor
'        .dICMSValor = tTributacaoOV.dICMSValor
'        .dIPIBase = tTributacaoOV.dIPIBase
'        .dIPIValor = tTributacaoOV.dIPIValor
'        .dIRRFAliquota = tTributacaoOV.dIRRFAliquota
'        .dIRRFBase = tTributacaoOV.dIRRFBase
'        .dIRRFValor = tTributacaoOV.dIRRFValor
'        .dISSAliquota = tTributacaoOV.dISSAliquota
'        .dISSBase = tTributacaoOV.dISSBase
'        .dISSValor = tTributacaoOV.dISSValor
'        .iFilialEmpresa = tTributacaoOV.iFilialEmpresa
'        .iICMSBaseManual = tTributacaoOV.iICMSBaseManual
'        .iICMSSubstBaseManual = tTributacaoOV.iICMSSubstBaseManual
'        .iICMSSubstValorManual = tTributacaoOV.iICMSSubstValorManual
'        .iICMSValorManual = tTributacaoOV.iICMSValorManual
'        .iIPIBaseManual = tTributacaoOV.iIPIBaseManual
'        .iIPIValorManual = tTributacaoOV.iIPIValorManual
'        .iIRRFAliquotaManual = tTributacaoOV.iIRRFAliquotaManual
'        .iIRRFValorManual = tTributacaoOV.iIRRFValorManual
'        .iISSIncluso = tTributacaoOV.iISSIncluso
'        .iISSInclusoManual = tTributacaoOV.iISSValorManual
'        .iISSValorManual = tTributacaoOV.iISSValorManual
'        .iTaxacaoAutomatica = tTributacaoOV.iTaxacaoAutomatica
'        .iTipoTributacao = tTributacaoOV.iTipoTributacao
'        .iTipoTributacaoManual = tTributacaoOV.iTipoTributacaoManual
'        .lCodOrcamento = tTributacaoOV.lCodOrcamento
'        .iPISRetidoManual = tTributacaoOV.iPISRetidoManual
'        .iCOFINSRetidoManual = tTributacaoOV.iCOFINSRetidoManual
'        .iCSLLRetidoManual = tTributacaoOV.iCSLLRetidoManual
'        .dPISRetido = tTributacaoOV.dPISRetido
'        .dCOFINSRetido = tTributacaoOV.dCOFINSRetido
'        .dCSLLRetido = tTributacaoOV.dCSLLRetido
'
'        .dPISCredito = tTributacaoOV.dPISCredito
'        .iPISCreditoManual = tTributacaoOV.iPISCreditoManual
'        .dCOFINSCredito = tTributacaoOV.dCOFINSCredito
'        .iCOFINSCreditoManual = tTributacaoOV.iCOFINSCreditoManual
'        .dIPICredito = tTributacaoOV.dIPICredito
'        .iIPICreditoManual = tTributacaoOV.iIPICreditoManual
'        .dICMSCredito = tTributacaoOV.dICMSCredito
'        .iICMSCreditoManual = tTributacaoOV.iICMSCreditoManual
'        .dValorINSS = tTributacaoOV.dValorINSS
'        .iINSSValorManual = tTributacaoOV.iINSSValorManual
'        .iINSSRetido = tTributacaoOV.iINSSRetido
'        .iINSSRetidoManual = tTributacaoOV.iINSSRetidoManual
'        .dINSSBase = tTributacaoOV.dINSSBase
'        .iINSSBaseManual = tTributacaoOV.iINSSBaseManual
'        .dINSSDeducoes = tTributacaoOV.dINSSDeducoes
'        .iINSSDeducoesManual = tTributacaoOV.iINSSDeducoesManual
'    End With
'
'    Set objOrcamentoVenda.objTributacaoOV = objTributacaoOV
'
'    lErro = CF("OrcamentoSRV_Le_Complementos", objOrcamentoVenda)
'    If lErro <> SUCESSO And lErro <> 186610 Then gError 186624
'
'    If lErro <> SUCESSO Then gError 186625
'
'    Call Comando_Fechar(lComando)

    objOrcamentoVenda.objTributacao.iTipoDoc = TIPODOC_TRIB_OSRV
    objOrcamentoVenda.objTributacao.lNumIntDoc = objOrcamentoVenda.lNumIntDoc

    lErro = CF("TributacaoDoc_Le", objOrcamentoVenda.objTributacao)
    If lErro <> SUCESSO Then gError 186623
    
    OrcamentoSRV_Le_Tributacao = SUCESSO

    Exit Function

Erro_OrcamentoSRV_Le_Tributacao:

    OrcamentoSRV_Le_Tributacao = gErr

    Select Case gErr
'
'        Case 186620
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 186621, 186622
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRIBUTACAOORCSRV", gErr, objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo)

        Case 186623 ', 186624
'
'        Case 186625
'            Call Rotina_Erro(vbOKOnly, "ERRO_COMPLEMENTO_OV_NAO_CADASTRADO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186626)

    End Select

'    Call Comando_Fechar(lComando)

    Exit Function

End Function

'Public Function OrcamentoSRV_Le_Complementos(ByVal objOrcamentoVenda As ClassOrcamentoVenda) As Long
'
'Dim lComando As Long, objTributacaoOV As ClassTributacaoOV
'Dim tTributacaoComplOV As typeTributacaoComplPV, objTributacaoItemOV As ClassTributacaoItemPV
'Dim lErro As Long
'
'On Error GoTo Erro_OrcamentoSRV_Le_Complementos
'
'    lComando = Comando_Abrir()
'    If lComando = 0 Then gError 186607
'
'    Set objTributacaoOV = objOrcamentoVenda.objTributacaoOV
'
'    tTributacaoComplOV.sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
'
'    'ler dados de frete, seguro, despesas, descontos da tabela TributacaoComplOV
'    lErro = Comando_Executar(lComando, "SELECT Item, NaturezaOp, NaturezaOpManual, TipoTributacao, TipoTributacaoManual, IPITipo, IPITipoManual, IPIBaseCalculo, IPIBaseManual, IPIPercRedBase, IPIPercRedBaseManual, IPIAliquota, " & _
'        "IPIAliquotaManual, IPIValor, IPIValorManual, ICMSTipo, ICMSTipoManual, ICMSBase, ICMSBaseManual, ICMSPercRedBase, ICMSPercRedBaseManual, ICMSAliquota, ICMSAliquotaManual, ICMSValor, ICMSValorManual, ICMSSubstBase, ICMSSubstBaseManual, " & _
'        "ICMSSubstAliquota, ICMSSubstAliquotaManual, ICMSSubstValor, ICMSSubstValorManual, PISCRedito, COFINSCredito, IPICredito, ICMSCredito FROM TributacaoComplOrcSRV WHERE FilialEmpresa = ? AND CodOrcamento = ?", _
'        tTributacaoComplOV.iItem, tTributacaoComplOV.sNaturezaOp, tTributacaoComplOV.iNaturezaOpManual, tTributacaoComplOV.iTipoTributacao, tTributacaoComplOV.iTipoTributacaoManual, tTributacaoComplOV.iIPITipo, tTributacaoComplOV.iIPITipoManual, tTributacaoComplOV.dIPIBaseCalculo, tTributacaoComplOV.iIPIBaseManual, tTributacaoComplOV.dIPIPercRedBase, _
'        tTributacaoComplOV.iIPIPercRedBaseManual, tTributacaoComplOV.dIPIAliquota, tTributacaoComplOV.iIPIAliquotaManual, tTributacaoComplOV.dIPIValor, tTributacaoComplOV.iIPIValorManual, tTributacaoComplOV.iICMSTipo, tTributacaoComplOV.iICMSTipoManual, tTributacaoComplOV.dICMSBase, tTributacaoComplOV.iICMSBaseManual, tTributacaoComplOV.dICMSPercRedBase, _
'        tTributacaoComplOV.iICMSPercRedBaseManual, tTributacaoComplOV.dICMSAliquota, tTributacaoComplOV.iICMSAliquotaManual, tTributacaoComplOV.dICMSValor, tTributacaoComplOV.iICMSValorManual, tTributacaoComplOV.dICMSSubstBase, tTributacaoComplOV.iICMSSubstBaseManual, _
'        tTributacaoComplOV.dICMSSubstAliquota, tTributacaoComplOV.iICMSSubstAliquotaManual, tTributacaoComplOV.dICMSSubstValor, tTributacaoComplOV.iICMSSubstValorManual, tTributacaoComplOV.dPISCredito, tTributacaoComplOV.dCOFINSCredito, tTributacaoComplOV.dIPICredito, tTributacaoComplOV.dICMSCredito, objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo)
'    If lErro <> AD_SQL_SUCESSO Then gError 186608
'
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 186609
'
'    If lErro = AD_SQL_SEM_DADOS Then gError 186610
'
'    Do While lErro <> AD_SQL_SEM_DADOS
'
'        Set objTributacaoItemOV = New ClassTributacaoItemPV
'
'        With objTributacaoItemOV
'
'            .lNumIntDoc = tTributacaoComplOV.iItem
'            .sNaturezaOp = tTributacaoComplOV.sNaturezaOp
'            .iNaturezaOpManual = tTributacaoComplOV.iNaturezaOpManual
'            .iTipoTributacao = tTributacaoComplOV.iTipoTributacao
'            .iTipoTributacaoManual = tTributacaoComplOV.iTipoTributacaoManual
'            .iIPITipo = tTributacaoComplOV.iIPITipo
'            .iIPITipoManual = tTributacaoComplOV.iIPITipoManual
'            .dIPIBaseCalculo = tTributacaoComplOV.dIPIBaseCalculo
'            .iIPIBaseManual = tTributacaoComplOV.iIPIBaseManual
'            .dIPIPercRedBase = tTributacaoComplOV.dIPIPercRedBase
'            .iIPIPercRedBaseManual = tTributacaoComplOV.iIPIPercRedBaseManual
'            .dIPIAliquota = tTributacaoComplOV.dIPIAliquota
'            .iIPIAliquotaManual = tTributacaoComplOV.iIPIAliquotaManual
'            .dIPIValor = tTributacaoComplOV.dIPIValor
'            .iIPIValorManual = tTributacaoComplOV.iIPIValorManual
'            .iICMSTipo = tTributacaoComplOV.iICMSTipo
'            .iICMSTipoManual = tTributacaoComplOV.iICMSTipoManual
'            .dICMSBase = tTributacaoComplOV.dICMSBase
'            .iICMSBaseManual = tTributacaoComplOV.iICMSBaseManual
'            .dICMSPercRedBase = tTributacaoComplOV.dICMSPercRedBase
'            .iICMSPercRedBaseManual = tTributacaoComplOV.iICMSPercRedBaseManual
'            .dICMSAliquota = tTributacaoComplOV.dICMSAliquota
'            .iICMSAliquotaManual = tTributacaoComplOV.iICMSAliquotaManual
'            .dICMSValor = tTributacaoComplOV.dICMSValor
'            .iICMSValorManual = tTributacaoComplOV.iICMSValorManual
'            .dICMSSubstBase = tTributacaoComplOV.dICMSSubstBase
'            .iICMSSubstBaseManual = tTributacaoComplOV.iICMSSubstBaseManual
'            .dICMSSubstAliquota = tTributacaoComplOV.dICMSSubstAliquota
'            .iICMSSubstAliquotaManual = tTributacaoComplOV.iICMSSubstAliquotaManual
'            .dICMSSubstValor = tTributacaoComplOV.dICMSSubstValor
'            .iICMSSubstValorManual = tTributacaoComplOV.iICMSSubstValorManual
'            .dICMSCredito = tTributacaoComplOV.dICMSCredito
'            .dIPICredito = tTributacaoComplOV.dIPICredito
'            .dCOFINSCredito = tTributacaoComplOV.dCOFINSCredito
'            .dPISCredito = tTributacaoComplOV.dPISCredito
'
'        End With
'
'        Select Case tTributacaoComplOV.iItem
'
'            Case ITEM_TIPO_FRETE
'                Set objTributacaoOV.objTributacaoFrete = objTributacaoItemOV
'
'            Case ITEM_TIPO_SEGURO
'                Set objTributacaoOV.objTributacaoSeguro = objTributacaoItemOV
'
'            Case ITEM_TIPO_DESCONTO
'                Set objTributacaoOV.objTributacaoDesconto = objTributacaoItemOV
'
'            Case ITEM_TIPO_OUTRAS_DESP
'                Set objTributacaoOV.objTributacaoOutras = objTributacaoItemOV
'
'            Case Else 'item invalido
'                gError 186611
'
'        End Select
'
'        lErro = Comando_BuscarProximo(lComando)
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 186612
'
'    Loop
'
'    Call Comando_Fechar(lComando)
'
'    OrcamentoSRV_Le_Complementos = SUCESSO
'
'    Exit Function
'
'Erro_OrcamentoSRV_Le_Complementos:
'
'    OrcamentoSRV_Le_Complementos = gErr
'
'    Select Case gErr
'
'        Case 186607
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 186608, 186609, 186612
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRIBUTACAOCOMPLORCSRV", gErr, objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo)
'
'        Case 186610
'
'        Case 186611
'            Call Rotina_Erro(vbOKOnly, "ERRO_COMPL_ORCSRV_TIPO_INVALIDO", gErr, tTributacaoComplOV.iItem)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 186613)
'
'    End Select
'
'    Call Comando_Fechar(lComando)
'
'    Exit Function
'
'End Function

Public Function ParcelasOrcSRV_Le(ByVal objOrcamentoVenda As ClassOrcamentoVenda) As Long
'Lê as parcelas associadas ao objOrcamentoVenda

Dim lErro As Long
Dim lComando As Long
Dim tParcelaOV As typeParcelaOV
Dim objParcelaOV As ClassParcelaOV

On Error GoTo Erro_ParcelasOrcSRV_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 186614

    'Seleciona Parcelas vinculadas ao Pedido de Venda
    lErro = Comando_Executar(lComando, "SELECT NumParcela, DataVencimento, Valor, Desconto1Codigo, Desconto1Ate, Desconto1Valor, Desconto2Codigo, Desconto2Ate, Desconto2Valor, Desconto3Codigo, Desconto3Ate, Desconto3Valor FROM ParcelasOrcSRV WHERE OrcamentoSRV = ? AND FilialEmpresa = ? ORDER BY DataVencimento", tParcelaOV.iNumParcela, tParcelaOV.dtDataVencimento, tParcelaOV.dValor, tParcelaOV.iDesconto1Codigo, tParcelaOV.dtDesconto1Ate, tParcelaOV.dDesconto1Valor, tParcelaOV.iDesconto2Codigo, tParcelaOV.dtDesconto2Ate, tParcelaOV.dDesconto2Valor, tParcelaOV.iDesconto3Codigo, tParcelaOV.dtDesconto3Ate, tParcelaOV.dDesconto3Valor, objOrcamentoVenda.lCodigo, objOrcamentoVenda.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 186615

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 186616
    
    If lErro = AD_SQL_SEM_DADOS Then gError 186617

    Do While lErro <> AD_SQL_SEM_DADOS

        With tParcelaOV
            Set objParcelaOV = New ClassParcelaOV

            objParcelaOV.dValor = .dValor
            objParcelaOV.dtDataVencimento = .dtDataVencimento
            objParcelaOV.iNumParcela = .iNumParcela
            objParcelaOV.iDesconto1Codigo = .iDesconto1Codigo
            objParcelaOV.dtDesconto1Ate = .dtDesconto1Ate
            objParcelaOV.dDesconto1Valor = .dDesconto1Valor
            objParcelaOV.iDesconto2Codigo = .iDesconto2Codigo
            objParcelaOV.dtDesconto2Ate = .dtDesconto2Ate
            objParcelaOV.dDesconto2Valor = .dDesconto2Valor
            objParcelaOV.dtDesconto3Ate = .dtDesconto3Ate
            objParcelaOV.dDesconto3Valor = .dDesconto3Valor
            objParcelaOV.iDesconto3Codigo = .iDesconto3Codigo

            objOrcamentoVenda.colParcela.Add objParcelaOV
        End With

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 186618

    Loop

    Call Comando_Fechar(lComando)

    ParcelasOrcSRV_Le = SUCESSO

    Exit Function

Erro_ParcelasOrcSRV_Le:

    ParcelasOrcSRV_Le = gErr

    Select Case gErr

        Case 186614
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 186615, 186616, 186618
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PARCELASORCSRV", gErr, objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo)

        Case 186617

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 186619)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function OrcSRVObjetosVinculados_Lock(ByVal objOrcamentoVenda As ClassOrcamentoVenda) As Long
'Faz "lock" nos objetos associados ao Orcamento de Venda

Dim lErro As Long
Dim alComando(1 To 1) As Long
Dim objProduto As New ClassProduto
Dim objItemOV As ClassItemOV
Dim iIndice As Integer

On Error GoTo Erro_OrcSRVObjetosVinculados_Lock

    'Abe os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 186797
    Next

    'Para cada item de Pedido Venda
    For Each objItemOV In objOrcamentoVenda.colItensSRV

        objProduto.sCodigo = objItemOV.sProduto
        'Locka  o produto
        lErro = CF("Produto_Lock_Shared", alComando(1), objProduto)
        If lErro <> SUCESSO Then gError 186798

        'Armazena os dados do produto para que possam ser usados depois
        objItemOV.iClasseUM = objProduto.iClasseUM

    Next

    'fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    OrcSRVObjetosVinculados_Lock = SUCESSO

    Exit Function

Erro_OrcSRVObjetosVinculados_Lock:

    OrcSRVObjetosVinculados_Lock = gErr

    Select Case gErr

        Case 186797
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 186798

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186799)

    End Select

    'fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function OrcamentoSRV_Lock(ByVal objOrcamentoVenda As ClassOrcamentoVenda, ByVal lComando As Long) As Long
'Lê e faz o Lock no Orcamento de Venda

Dim lErro As Long
Dim tOrcamentoVenda As typeOrcamentoVenda
Dim lCodigo As Long

On Error GoTo Erro_OrcamentoSRV_Lock

    'Lê o Orcamento com o Código e a filialempresa passados
    lErro = Comando_ExecutarPos(lComando, "SELECT NumIntDoc, ValorTotal FROM OrcamentoSRV WHERE FilialEmpresa = ? and Codigo = ?", 0, tOrcamentoVenda.lNumIntDoc, tOrcamentoVenda.dValorTotal, objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 186828

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 186829
    
    If lErro = AD_SQL_SEM_DADOS Then gError 186830

    'Faz o "Lock" no Orcamento de Venda
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then gError 186831

    objOrcamentoVenda.lNumIntDoc = tOrcamentoVenda.lNumIntDoc
    objOrcamentoVenda.dValorTotal = tOrcamentoVenda.dValorTotal

    OrcamentoSRV_Lock = SUCESSO

    Exit Function

Erro_OrcamentoSRV_Lock:

    OrcamentoSRV_Lock = gErr

    Select Case gErr

        Case 186828, 186829
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORCAMENTOSRV", gErr, objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo)

        Case 186830

        Case 186831
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_ORCAMENTOSRV", gErr, objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186832)

    End Select

    Exit Function

End Function

Public Function ItensOrcamentoSRV_Le_Lock(ByVal objOrcamentoVenda As ClassOrcamentoVenda) As Long
'Faz lock exclusive nos itens de OrcamentoVenda e lê os NumIntDoc desses itens

Dim lErro As Long
Dim lComando As Long
Dim tItensOV As typeItemOV
Dim objItensOV As ClassItemOV
Dim iPeca As Integer

On Error GoTo Erro_ItensOrcamentoSRV_Le_Lock

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 186833

    tItensOV.sDescricao = String(STRING_ITEM_PEDIDO_DESCRICAO, 0)
    tItensOV.sProduto = String(STRING_PRODUTO, 0)
    tItensOV.sUnidadeMed = String(STRING_UM_SIGLA, 0)

    '###########################
    'INSERIDO POR WAGNER
    tItensOV.sVersaoKit = String(STRING_KIT_VERSAO, 0)
    tItensOV.sVersaoKitBase = String(STRING_KIT_VERSAO, 0)
    '##########################

    '###########################
    'ALTERADO POR WAGNER
    'Lê os itens do orçamento passado
    lErro = Comando_ExecutarPos(lComando, "SELECT NumIntDoc,Produto,Quantidade,PrecoUnitario,PrecoTotal,ClasseUM,UnidadeMed,ValorDesconto,DataEntrega,Descricao, VersaoKit, VersaoKitBase, Peca FROM ItensOrcamentoSRV WHERE FilialEmpresa = ? AND CodOrcamento = ?", 0, tItensOV.lNumIntDoc, tItensOV.sProduto, tItensOV.dQuantidade, tItensOV.dPrecoUnitario, tItensOV.dPrecoTotal, tItensOV.iClasseUM, tItensOV.sUnidadeMed, tItensOV.dValorDesconto, tItensOV.dtDataEntrega, tItensOV.sDescricao, tItensOV.sVersaoKit, tItensOV.sVersaoKitBase, iPeca, objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 186834
    '###########################

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then gError 186835
    
    If lErro = AD_SQL_SEM_DADOS Then gError 186836

    Do While lErro = AD_SQL_SUCESSO
        'faz o lock do item
        lErro = Comando_LockExclusive(lComando)
        If lErro <> AD_SQL_SUCESSO Then gError 186837

        With tItensOV
            Set objItensOV = New ClassItemOV

            .iFilialEmpresa = objOrcamentoVenda.iFilialEmpresa
            .lCodPedidoOV = objOrcamentoVenda.lCodigo

            objItensOV.iFilialEmpresa = .iFilialEmpresa
            objItensOV.lCodOrcamento = .lCodPedidoOV
            objItensOV.sProduto = .sProduto
            objItensOV.dQuantidade = .dQuantidade
            objItensOV.dPrecoUnitario = .dPrecoUnitario
            objItensOV.dPrecoTotal = .dPrecoTotal
            objItensOV.dValorDesconto = .dValorDesconto
            objItensOV.dtDataEntrega = .dtDataEntrega
            objItensOV.sUnidadeMed = .sUnidadeMed
            objItensOV.iClasseUM = .iClasseUM
            objItensOV.sDescricao = .sDescricao
            objItensOV.iStatus = .iStatus
            objItensOV.lNumIntDoc = .lNumIntDoc

            '###########################
            'INSERIDO POR WAGNER
            objItensOV.sVersaoKit = .sVersaoKit
            objItensOV.sVersaoKitBase = .sVersaoKitBase
            '###########################

            If iPeca = 0 Then
                objOrcamentoVenda.colItensSRV.Add objItensOV
            Else
                'Adiciona o item a colecao de itens do orçamento de Venda
                objOrcamentoVenda.colItens.Add objItensOV
            End If
        End With

        'Busca o próximo item
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then gError 186838

    Loop

    'Fecha o comando
    Call Comando_Fechar(lComando)

    ItensOrcamentoSRV_Le_Lock = SUCESSO

    Exit Function

Erro_ItensOrcamentoSRV_Le_Lock:

    ItensOrcamentoSRV_Le_Lock = gErr

    Select Case gErr

        Case 186833
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 186834, 186835, 186838
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSORCAMENTOSRV", gErr, objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo)

        Case 186836
            Call Rotina_Erro(vbOKOnly, "ERRO_ORCAMENTOSRV_SEM_ITENS", gErr, objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo)

        Case 186837
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_ITENSORCAMENTOSRV", gErr, objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186839)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function SolicitacaoSRV_Le_NumIntDoc(objSolicSRV As ClassSolicSRV) As Long

Dim lComando As Long
Dim lErro As Long
Dim tSolicSRV As typeSolicSRV

On Error GoTo Erro_SolicitacaoSRV_Le_NumIntDoc

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 186985

    tSolicSRV.sObs = String(STRING_MAXIMO, 0)

    lErro = Comando_Executar(lComando, "SELECT FilialEmpresa, Codigo, Data, Hora, Cliente, FilialCliente, NumIntDoc, Vendedor, Atendente, Prazo, PrazoTipo, Obs, DataEntrega, Tipo, Fase FROM SolicitacaoSRV WHERE NumIntDoc = ?", _
    tSolicSRV.iFilialEmpresa, tSolicSRV.lCodigo, tSolicSRV.dtData, tSolicSRV.dHora, tSolicSRV.lCliente, tSolicSRV.iFilial, tSolicSRV.lNumIntDoc, tSolicSRV.iVendedor, tSolicSRV.iAtendente, tSolicSRV.iPrazo, tSolicSRV.iPrazoTipo, tSolicSRV.sObs, tSolicSRV.dtDataEntrega, tSolicSRV.lTipo, tSolicSRV.lFase, objSolicSRV.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 186986
        
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 186987
        
    If lErro = AD_SQL_SEM_DADOS Then gError 186988

    objSolicSRV.iFilialEmpresa = tSolicSRV.iFilialEmpresa
    objSolicSRV.lCodigo = tSolicSRV.lCodigo
    objSolicSRV.dtData = tSolicSRV.dtData
    objSolicSRV.dtHora = tSolicSRV.dHora
    objSolicSRV.lCliente = tSolicSRV.lCliente
    objSolicSRV.iFilial = tSolicSRV.iFilial
    objSolicSRV.lNumIntDoc = tSolicSRV.lNumIntDoc
    objSolicSRV.iVendedor = tSolicSRV.iVendedor
    objSolicSRV.iAtendente = tSolicSRV.iAtendente
    objSolicSRV.iPrazo = tSolicSRV.iPrazo
    objSolicSRV.iPrazoTipo = tSolicSRV.iPrazoTipo
    objSolicSRV.sObs = tSolicSRV.sObs
    objSolicSRV.dtDataEntrega = tSolicSRV.dtDataEntrega
    objSolicSRV.lTipo = tSolicSRV.lTipo
    objSolicSRV.lFase = tSolicSRV.lFase
    
    lErro = CF("ItensSolicSRV_Le", objSolicSRV)
    If lErro <> SUCESSO Then gError 186989

    Call Comando_Fechar(lComando)
    
    SolicitacaoSRV_Le_NumIntDoc = SUCESSO
    
    Exit Function

Erro_SolicitacaoSRV_Le_NumIntDoc:
    
    SolicitacaoSRV_Le_NumIntDoc = gErr
        
    Select Case gErr
    
        Case 186985
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 186986, 186987
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SOLICITACAOSRV", gErr, objSolicSRV.iFilialEmpresa, objSolicSRV.lCodigo)
        
        Case 186988, 186989
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 186990)
            
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function Pesquisa_Garantia(ByVal sProduto As String, ByVal sServico As String, ByVal sLote As String, ByVal iFilialOP As Integer, lGarantia As Long) As Long

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objGarantia As New ClassGarantia
Dim sServicoFormatado As String
Dim iServicoPreenchido As Integer

On Error GoTo Erro_Pesquisa_Garantia

    If gobjSRV.iGarantiaAutoSolic = GARANTIA_AUTOMATICA_SOLICITACAO Then
    
        'Formata o Produto para o BD
        lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 183470
            
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            
            If Len(Trim(sServico)) > 0 Then
                
                objProduto.sCodigo = sProdutoFormatado
                        
                'Lê os demais atributos do Produto
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 183471
                    
                If lErro = 28030 Then gError 183472
        
                objGarantia.iFilialEmpresa = giFilialEmpresa
                objGarantia.sProduto = objProduto.sCodigo
                
                'Formata o Produto para o BD
                lErro = CF("Produto_Formata", sServico, sServicoFormatado, iServicoPreenchido)
                If lErro <> SUCESSO Then gError 183733
                
                If iServicoPreenchido = PRODUTO_PREENCHIDO Then
                
                    objGarantia.sServico = sServicoFormatado
            
                    If objProduto.iRastro = PRODUTO_RASTRO_LOTE Then
                
                        If Len(sLote) > 0 Then
                        
                            objGarantia.sLote = sLote
                            
                            lErro = CF("Garantia_Le_Lote", objGarantia)
                            If lErro <> SUCESSO And lErro <> 183445 Then gError 183473
                            
                            If lErro = SUCESSO Then
                                lGarantia = objGarantia.lCodigo
                            End If
                            
                        End If
                
                    ElseIf objProduto.iRastro = PRODUTO_RASTRO_OP Then
                
                        If Len(sLote) > 0 And iFilialOP <> 0 Then
                
                            objGarantia.sLote = sLote
                            objGarantia.iFilialOP = iFilialOP
                
                            lErro = CF("Garantia_Le_Lote_FilialOP", objGarantia)
                            If lErro <> SUCESSO And lErro <> 183460 Then gError 183474
                
                            If lErro = SUCESSO Then
                                lGarantia = objGarantia.lCodigo
                            End If
                
                        End If
                
                    ElseIf objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then
            
                        If Len(sLote) > 0 Then
                        
                            objGarantia.sLote = sLote
                            
                            lErro = CF("Garantia_Le_NumSerie", objGarantia)
                            If lErro <> SUCESSO And lErro <> 183466 Then gError 183475
                            
                            If lErro = SUCESSO Then
                                lGarantia = objGarantia.lCodigo
                            End If
                        
                        End If
        
                    End If
        
                End If
    
            End If
    
        End If

    End If
    
    Pesquisa_Garantia = SUCESSO

    Exit Function

Erro_Pesquisa_Garantia:

    Pesquisa_Garantia = gErr

    Select Case gErr

        Case 183470, 183471, 183473 To 183475, 183733

        Case 183472
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183476)

    End Select

    Exit Function

End Function

Function Pesquisa_Contrato(ByVal sProduto As String, ByVal sServico As String, ByVal sLote As String, ByVal iFilialOP As Integer, sContrato As String) As Long

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objItensContratoSrv As New ClassItensDeContratoSrv
Dim sServicoFormatado As String
Dim iServicoPreenchido As Integer

On Error GoTo Erro_Pesquisa_Contrato

    If gobjSRV.iContratoAutoSolic = CONTRATO_AUTOMATICO_SOLICITACAO Then
    
        'Formata o Produto para o BD
        lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 183512
            
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            
            If Len(Trim(sServico)) > 0 Then
            
                objProduto.sCodigo = sProdutoFormatado
                        
                'Lê os demais atributos do Produto
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 183513
                    
                If lErro = 28030 Then gError 183514
        
                objItensContratoSrv.iFilialEmpresa = giFilialEmpresa
                objItensContratoSrv.sProduto = objProduto.sCodigo
                
                'Formata o Produto para o BD
                lErro = CF("Produto_Formata", sServico, sServicoFormatado, iServicoPreenchido)
                If lErro <> SUCESSO Then gError 183734
                
                If iServicoPreenchido = PRODUTO_PREENCHIDO Then
                
                    objItensContratoSrv.sServico = sServicoFormatado
        
                    If objProduto.iRastro = PRODUTO_RASTRO_LOTE Then
                
                        If Len(Trim(sLote)) > 0 Then
                        
                           objItensContratoSrv.sLote = sLote
                            
                            lErro = CF("ItensDeContratoSrv_Le_Lote", objItensContratoSrv)
                            If lErro <> SUCESSO And lErro <> 183480 Then gError 183515
                            
                            If lErro = SUCESSO Then
                                sContrato = objItensContratoSrv.sCodigoContrato
                            End If
                            
                        End If
                
                    ElseIf objProduto.iRastro = PRODUTO_RASTRO_OP Then
                
                        If Len(Trim(sLote)) > 0 And iFilialOP <> 0 Then
                
                            objItensContratoSrv.sLote = sLote
                            objItensContratoSrv.iFilialOP = iFilialOP
                
                            lErro = CF("ItensDeContratoSrv_Le_Lote_FilialOP", objItensContratoSrv)
                            If lErro <> SUCESSO And lErro <> 183508 Then gError 183516
                
                            If lErro = SUCESSO Then
                                sContrato = objItensContratoSrv.sCodigoContrato
                            End If
                
                        End If
                
                    ElseIf objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then
            
                        If Len(Trim(sLote)) > 0 Then
                        
                            objItensContratoSrv.sLote = sLote
                            
                            lErro = CF("ItensDeContratoSrv_Le_NumSerie", objItensContratoSrv)
                            If lErro <> SUCESSO And lErro <> 183516 Then gError 183517
                            
                            If lErro = SUCESSO Then
                                sContrato = objItensContratoSrv.sCodigoContrato
                            End If
                            
                        End If
            
                    End If
        
                End If
    
            End If
    
        End If

    End If
    
    Pesquisa_Contrato = SUCESSO

    Exit Function

Erro_Pesquisa_Contrato:

    Pesquisa_Contrato = gErr

    Select Case gErr

        Case 183512, 183513, 183515 To 183517, 183734

        Case 183514
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183518)

    End Select

    Exit Function

End Function

Function PedidoServico_Le_Todos_Completo(ByVal objPedidoDeVenda As ClassPedidoDeVenda, Optional iStatus As Integer) As Long
'Le o Pedido Inclusive se ele estiver Baixado

Dim lErro As Long

On Error GoTo Erro_PedidoServico_Le_Todos_Completo
    
    'Lê o Pedido de Venda
    lErro = CF("PedidoServico_Le", objPedidoDeVenda)
    If lErro <> SUCESSO And lErro <> 188828 Then gError 193252
        
    'Se não encontrou ---> ERRO
    If lErro <> SUCESSO Then gError 193253
    
    'Le a Tributação do Pedido
    lErro = CF("PedidoServico_Le_Tributacao", objPedidoDeVenda)
    If lErro <> SUCESSO And lErro <> 188837 Then gError 193255
                   
    'Le as Parcelas do Pedido
    lErro = CF("ParcelasPedidoSRV_Le", objPedidoDeVenda)
    If lErro <> SUCESSO Then gError 193256
    
    'Le as Comissoes do Pedido
    lErro = CF("ComissoesPedidoSRV_Le", objPedidoDeVenda)
    If lErro <> SUCESSO Then gError 193257
    
'    'Le os Bloqueios do Pedido
'    lErro = CF("BloqueiosPedidoSRV_Le", objPedidoDeVenda)
'    If lErro <> SUCESSO Then gError 193258
    
    'Le os Itens e a Tributacao
    lErro = CF("PedidoSRV_Le_Itens_ComTributacao", objPedidoDeVenda)
    If lErro <> SUCESSO Then gError 193259
    
    objPedidoDeVenda.objInfoAdic.iTipoDoc = TIPODOC_INFOADIC_PSRV
    objPedidoDeVenda.objInfoAdic.lNumIntDoc = objPedidoDeVenda.lNumIntDoc
    lErro = CF("InfoAdicionais_Le", objPedidoDeVenda.objInfoAdic)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 193259
                
    PedidoServico_Le_Todos_Completo = SUCESSO
    
    Exit Function
    
Erro_PedidoServico_Le_Todos_Completo:

    PedidoServico_Le_Todos_Completo = gErr
    
    Select Case gErr
        
        Case 193252, 293254 To 293259
        
        Case 193253
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_SERVICO_NAO_CADASTRADO", gErr, objPedidoDeVenda.lCodigo)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193260)
    
    End Select
        
End Function

Function PedidoServico_Le(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
'Lê o Pedido de Venda a partir do Código passado.

Dim lComando As Long
Dim lErro As Long

On Error GoTo Erro_PedidoServico_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 188826

    lErro = CF("PedidoServico_Le_Aux", objPedidoVenda, lComando)
    If lErro <> SUCESSO And lErro <> 188832 Then gError 188827
    
    'se nao encontrou o pedido
    If lErro <> SUCESSO Then gError 188828
    
    Call Comando_Fechar(lComando)

    PedidoServico_Le = SUCESSO

    Exit Function

Erro_PedidoServico_Le:

    PedidoServico_Le = gErr

    Select Case gErr

        Case 188826
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 188827, 188828

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 188829)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function PedidoServico_Le_Aux(ByVal objPedidoVenda As ClassPedidoDeVenda, ByVal lComando As Long) As Long
'Lê o Pedido de Servico a partir do Código passado.

Dim lErro As Long
Dim tPedidoVenda As typePedidoVenda

On Error GoTo Erro_PedidoServico_Le_Aux

    tPedidoVenda.sMensagemNota = String(STRING_PEDIDOVENDA_MENSAGEM_NOTA, 0)
    tPedidoVenda.sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
    tPedidoVenda.sPedidoCliente = String(STRING_PEDIDOVENDA_PEDIDO_CLIENTE, 0)
    tPedidoVenda.sPlaca = String(STRING_PEDIDOVENDA_PLACA, 0)
    tPedidoVenda.sPlacaUF = String(STRING_PEDIDOVENDA_PLACA_UF, 0)
    tPedidoVenda.sVolumeNumero = String(STRING_PEDIDOVENDA_VOLUME_NUMERO, 0)
    tPedidoVenda.sPedidoRepresentante = String(STRING_PEDIDOVENDA_PEDIDO_REPRESENTANTE, 0)
    
    lErro = Comando_Executar(lComando, "SELECT FilialEmpresaFaturamento, Cliente, Filial, FilialEntrega, CodTransportadora, CondicaoPagto, PercAcrescFinanceiro, DataEmissao, MensagemNota, NaturezaOp, ValorTotal, ValorFrete, ValorDesconto, ValorSeguro, PedidoCliente, CanalVenda, TabelaPreco, ProxSeqBloqueio, FaturaIntegral, ComissaoAutomatica, CobrancaAutomatica, DataReferencia, NumIntDoc,ValorOutrasDespesas,Placa,PlacaUF,VolumeQuant,VolumeEspecie,VolumeMarca,PesoLiq,PesoBruto, VolumeNumero,FreteRespons, ValorProdutos, CodTranspRedesp, DetPagFrete, DataEntrega, VolumeTotal, Moeda, TaxaMoeda,PedidoRepresentante, Status, NumIntSolicSRV,ValorDescontoTit,ValorDescontoItens,ValorItens FROM PedidoServico WHERE FilialEmpresa = ? AND Codigo = ? ", tPedidoVenda.iFilialEmpresaFaturamento, tPedidoVenda.lCliente, tPedidoVenda.iFilial, tPedidoVenda.iFilialEntrega, tPedidoVenda.iCodTransportadora, tPedidoVenda.iCondicaoPagto, _
        tPedidoVenda.dPercAcrescFinanceiro, tPedidoVenda.dtDataEmissao, tPedidoVenda.sMensagemNota, tPedidoVenda.sNaturezaOp, tPedidoVenda.dValorTotal, tPedidoVenda.dValorFrete, tPedidoVenda.dValorDesconto, tPedidoVenda.dValorSeguro, tPedidoVenda.sPedidoCliente, tPedidoVenda.iCanalVenda, tPedidoVenda.iTabelaPreco, tPedidoVenda.iProxSeqBloqueio, tPedidoVenda.iFaturaIntegral, tPedidoVenda.iComissaoAutomatica, tPedidoVenda.iCobrancaAutomatica, tPedidoVenda.dtDataReferencia, tPedidoVenda.lNumIntDoc, tPedidoVenda.dValorOutrasDespesas, tPedidoVenda.sPlaca, tPedidoVenda.sPlacaUF, tPedidoVenda.lVolumeQuant, tPedidoVenda.lVolumeEspecie, tPedidoVenda.lVolumeMarca, tPedidoVenda.dPesoLiq, tPedidoVenda.dPesoBruto, tPedidoVenda.sVolumeNumero, tPedidoVenda.iFreteRespons, tPedidoVenda.dValorProdutos, tPedidoVenda.iCodTranspRedesp, tPedidoVenda.iDetPagFrete, _
        tPedidoVenda.dtDataEntrega, tPedidoVenda.dVolumeTotal, tPedidoVenda.iMoeda, tPedidoVenda.dTaxaMoeda, tPedidoVenda.sPedidoRepresentante, tPedidoVenda.iStatus, tPedidoVenda.lNumIntSolicSRV, tPedidoVenda.dValorDescontoTit, tPedidoVenda.dValorDescontoItens, tPedidoVenda.dValorItens, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 188830

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 188831

    If lErro = AD_SQL_SEM_DADOS Then gError 188832

    objPedidoVenda.iFilialEmpresaFaturamento = tPedidoVenda.iFilialEmpresaFaturamento
    objPedidoVenda.lCliente = tPedidoVenda.lCliente
    objPedidoVenda.iFilial = tPedidoVenda.iFilial
    objPedidoVenda.iFilialEntrega = tPedidoVenda.iFilialEntrega
    objPedidoVenda.iCodTransportadora = tPedidoVenda.iCodTransportadora
    objPedidoVenda.iCodTranspRedesp = tPedidoVenda.iCodTranspRedesp
    objPedidoVenda.iDetPagFrete = tPedidoVenda.iDetPagFrete
    objPedidoVenda.dtDataEntrega = tPedidoVenda.dtDataEntrega
    objPedidoVenda.iCondicaoPagto = tPedidoVenda.iCondicaoPagto
    objPedidoVenda.dPercAcrescFinanceiro = tPedidoVenda.dPercAcrescFinanceiro
    objPedidoVenda.dtDataEmissao = tPedidoVenda.dtDataEmissao
    objPedidoVenda.sMensagemNota = tPedidoVenda.sMensagemNota
    objPedidoVenda.sNaturezaOp = tPedidoVenda.sNaturezaOp
    objPedidoVenda.dValorTotal = tPedidoVenda.dValorTotal
    objPedidoVenda.dValorFrete = tPedidoVenda.dValorFrete
    objPedidoVenda.dValorDesconto = tPedidoVenda.dValorDesconto
    objPedidoVenda.dValorOutrasDespesas = tPedidoVenda.dValorOutrasDespesas
    objPedidoVenda.dValorSeguro = tPedidoVenda.dValorSeguro
    objPedidoVenda.sPedidoCliente = tPedidoVenda.sPedidoCliente
    objPedidoVenda.iCanalVenda = tPedidoVenda.iCanalVenda
    objPedidoVenda.iTabelaPreco = tPedidoVenda.iTabelaPreco
    objPedidoVenda.iProxSeqBloqueio = tPedidoVenda.iProxSeqBloqueio
    objPedidoVenda.iFaturaIntegral = tPedidoVenda.iFaturaIntegral
    objPedidoVenda.iComissaoAutomatica = tPedidoVenda.iComissaoAutomatica
    objPedidoVenda.iCobrancaAutomatica = tPedidoVenda.iCobrancaAutomatica
    objPedidoVenda.dtDataReferencia = tPedidoVenda.dtDataReferencia
    objPedidoVenda.sPlaca = tPedidoVenda.sPlaca
    objPedidoVenda.sPlacaUF = tPedidoVenda.sPlacaUF
    objPedidoVenda.lVolumeQuant = tPedidoVenda.lVolumeQuant
    objPedidoVenda.lVolumeEspecie = tPedidoVenda.lVolumeEspecie
    objPedidoVenda.lVolumeMarca = tPedidoVenda.lVolumeMarca
    objPedidoVenda.dPesoLiq = tPedidoVenda.dPesoLiq
    objPedidoVenda.dPesoBruto = tPedidoVenda.dPesoBruto
    objPedidoVenda.lNumIntDoc = tPedidoVenda.lNumIntDoc
    objPedidoVenda.sVolumeNumero = tPedidoVenda.sVolumeNumero
    objPedidoVenda.iFreteRespons = tPedidoVenda.iFreteRespons
    objPedidoVenda.dValorProdutos = tPedidoVenda.dValorProdutos
    objPedidoVenda.dVolumeTotal = tPedidoVenda.dVolumeTotal
    objPedidoVenda.iMoeda = tPedidoVenda.iMoeda
    objPedidoVenda.dTaxaMoeda = tPedidoVenda.dTaxaMoeda
    objPedidoVenda.sPedidoRepresentante = tPedidoVenda.sPedidoRepresentante
    objPedidoVenda.iStatus = tPedidoVenda.iStatus
    objPedidoVenda.lNumIntSolicSRV = tPedidoVenda.lNumIntSolicSRV
    objPedidoVenda.dValorDescontoTit = tPedidoVenda.dValorDescontoTit
    objPedidoVenda.dValorDescontoItens = tPedidoVenda.dValorDescontoItens
    objPedidoVenda.dValorItens = tPedidoVenda.dValorItens
    
    PedidoServico_Le_Aux = SUCESSO

    Exit Function

Erro_PedidoServico_Le_Aux:

    PedidoServico_Le_Aux = gErr

    Select Case gErr

        Case 188830, 188831
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOSDEVENDA", gErr, objPedidoVenda.lCodigo)

        Case 188832

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 188833)

    End Select

    Exit Function

End Function

Function PedidoServico_Le_Tributacao(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long

'Dim lComando As Long, tTributacaoPV As typeTributacaoPV, objTributacaoPV As New ClassTributacaoPV
Dim lErro As Long

On Error GoTo Erro_PedidoServico_Le_Tributacao

'    lComando = Comando_Abrir()
'    If lComando = 0 Then gError 188834
'
'    'ler tributacaopv
'    lErro = Comando_Executar(lComando, "SELECT TaxacaoAutomatica, TipoTributacao, TipoTributacaoManual, IPIBase, IPIBaseManual, IPIValor, IPIValorManual, ICMSBase, ICMSBaseManual, ICMSValor, ICMSValorManual, ICMSSubstBase, ICMSSubstBaseManual, ICMSSubstValor, ICMSSubstValorManual, ISSIncluso, ISSBase, ISSAliquota, ISSAliquotaManual, ISSValor, ISSValorManual, IRRFBase, IRRFAliquota, IRRFAliquotaManual, IRRFValor, IRRFValorManual, PISRetidoManual, COFINSRetidoManual, CSLLRetidoManual, PISRetido, COFINSRetido, CSLLRetido, " & _
'        "PISCredito,PISCreditoManual,COFINSCredito,COFINSCreditoManual,IPICredito,IPICreditoManual,ICMSCredito,ICMSCreditoManual,ValorINSS,ValorINSSManual,INSSRetido,INSSRetidoManual,INSSValorBase,INSSValorBaseManual,INSSValorDeducoes ,INSSValorDeducoesManual FROM TributacaoPedSRV WHERE FilialEmpresa = ? AND CodPedido = ?", _
'        tTributacaoPV.iTaxacaoAutomatica, tTributacaoPV.iTipoTributacao, tTributacaoPV.iTipoTributacaoManual, tTributacaoPV.dIPIBase, tTributacaoPV.iIPIBaseManual, tTributacaoPV.dIPIValor, tTributacaoPV.iIPIValorManual, tTributacaoPV.dICMSBase, tTributacaoPV.iICMSBaseManual, tTributacaoPV.dICMSValor, tTributacaoPV.iICMSValorManual, tTributacaoPV.dICMSSubstBase, tTributacaoPV.iICMSSubstBaseManual, tTributacaoPV.dICMSSubstValor, tTributacaoPV.iICMSSubstValorManual, tTributacaoPV.iISSIncluso, _
'        tTributacaoPV.dISSBase, tTributacaoPV.dISSAliquota, tTributacaoPV.iISSAliquotaManual, tTributacaoPV.dISSValor, tTributacaoPV.iISSValorManual, tTributacaoPV.dIRRFBase, tTributacaoPV.dIRRFAliquota, tTributacaoPV.iIRRFAliquotaManual, tTributacaoPV.dIRRFValor, tTributacaoPV.iIRRFValorManual, tTributacaoPV.iPISRetidoManual, tTributacaoPV.iCOFINSRetidoManual, tTributacaoPV.iCSLLRetidoManual, tTributacaoPV.dPISRetido, tTributacaoPV.dCOFINSRetido, tTributacaoPV.dCSLLRetido, _
'        tTributacaoPV.dPISCredito, tTributacaoPV.iPISCreditoManual, tTributacaoPV.dCOFINSCredito, tTributacaoPV.iCOFINSCreditoManual, tTributacaoPV.dIPICredito, tTributacaoPV.iIPICreditoManual, tTributacaoPV.dICMSCredito, tTributacaoPV.iICMSCreditoManual, tTributacaoPV.dValorINSS, tTributacaoPV.iINSSValorManual, tTributacaoPV.iINSSRetido, tTributacaoPV.iINSSRetidoManual, tTributacaoPV.dINSSBase, tTributacaoPV.iINSSBaseManual, tTributacaoPV.dINSSDeducoes, tTributacaoPV.iINSSDeducoesManual, _
'        objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo)
'    If lErro <> AD_SQL_SUCESSO Then gError 188835
'
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 188836
'
'    If lErro = AD_SQL_SEM_DADOS Then gError 188837
'
'    objPedidoVenda.iNaturezaOpManual = VAR_PREENCH_MANUAL
'    objTributacaoPV.lCodPedido = objPedidoVenda.lCodigo
'    objTributacaoPV.iFilialEmpresa = giFilialEmpresa
'
'    'transferir dados lidos p/objeto
'    With objTributacaoPV
'
'        .iTaxacaoAutomatica = tTributacaoPV.iTaxacaoAutomatica
'        .iTipoTributacao = tTributacaoPV.iTipoTributacao
'        .iTipoTributacaoManual = tTributacaoPV.iTipoTributacaoManual
'        .dIPIBase = tTributacaoPV.dIPIBase
'        .iIPIBaseManual = tTributacaoPV.iIPIBaseManual
'        .dIPIValor = tTributacaoPV.dIPIValor
'        .iIPIValorManual = tTributacaoPV.iIPIValorManual
'        .dICMSBase = tTributacaoPV.dICMSBase
'        .iICMSBaseManual = tTributacaoPV.iICMSBaseManual
'        .dICMSValor = tTributacaoPV.dICMSValor
'        .iICMSValorManual = tTributacaoPV.iICMSValorManual
'        .dICMSSubstBase = tTributacaoPV.dICMSSubstBase
'        .iICMSSubstBaseManual = tTributacaoPV.iICMSSubstBaseManual
'        .dICMSSubstValor = tTributacaoPV.dICMSSubstValor
'        .iICMSSubstValorManual = tTributacaoPV.iICMSSubstValorManual
'        .iISSIncluso = tTributacaoPV.iISSIncluso
'        .dISSBase = tTributacaoPV.dISSBase
'        .dISSAliquota = tTributacaoPV.dISSAliquota
'        .iISSAliquotaManual = tTributacaoPV.iISSAliquotaManual
'        .dISSValor = tTributacaoPV.dISSValor
'        .iISSValorManual = tTributacaoPV.iISSValorManual
'        .dIRRFBase = tTributacaoPV.dIRRFBase
'        .dIRRFAliquota = tTributacaoPV.dIRRFAliquota
'        .iIRRFAliquotaManual = tTributacaoPV.iIRRFAliquotaManual
'        .dIRRFValor = tTributacaoPV.dIRRFValor
'        .iIRRFValorManual = tTributacaoPV.iIRRFValorManual
'        .dPISRetido = tTributacaoPV.dPISRetido
'        .iPISRetidoManual = tTributacaoPV.iPISRetidoManual
'        .dCOFINSRetido = tTributacaoPV.dCOFINSRetido
'        .iCOFINSRetidoManual = tTributacaoPV.iCOFINSRetidoManual
'        .dCSLLRetido = tTributacaoPV.dCSLLRetido
'        .iCSLLRetidoManual = tTributacaoPV.iCSLLRetidoManual
'    End With
'
'    Set objPedidoVenda.objTributacaoPV = objTributacaoPV
'
'    lErro = PedidoServico_Le_Complementos(objPedidoVenda)
'    If lErro <> SUCESSO Then gError 188838
'
'    Call Comando_Fechar(lComando)

    objPedidoVenda.objTributacao.iTipoDoc = TIPODOC_TRIB_PSRV
    objPedidoVenda.objTributacao.lNumIntDoc = objPedidoVenda.lNumIntDoc

    lErro = CF("TributacaoDoc_Le", objPedidoVenda.objTributacao)
    If lErro <> SUCESSO Then gError 188837

    PedidoServico_Le_Tributacao = SUCESSO

    Exit Function

Erro_PedidoServico_Le_Tributacao:

    PedidoServico_Le_Tributacao = gErr

    Select Case gErr

'        Case 188834
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 188835, 188836
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRIBUTACAOPEDSRV", gErr, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo)

        Case 188837 ', 188838

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 188839)

    End Select

'    Call Comando_Fechar(lComando)

    Exit Function

End Function
'
'Private Function PedidoServico_Le_Complementos(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
'
'Dim lComando As Long, objTributacaoPV As ClassTributacaoPV
'Dim tTributacaoComplPV As typeTributacaoComplPV, objTributacaoItemPV As ClassTributacaoItemPV
'Dim lErro As Long
'
'On Error GoTo Erro_PedidoServico_Le_Complementos
'
'    lComando = Comando_Abrir()
'    If lComando = 0 Then gError 188840
'
'    Set objTributacaoPV = objPedidoVenda.objTributacaoPV
'
'    tTributacaoComplPV.sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
'
'    'ler dados de frete, seguro, despesas, descontos da tabela TributacaoComplPV
'    lErro = Comando_Executar(lComando, "SELECT Item, NaturezaOp, NaturezaOpManual, TipoTributacao, TipoTributacaoManual, IPITipo, IPITipoManual, IPIBaseCalculo, IPIBaseManual, IPIPercRedBase, IPIPercRedBaseManual, IPIAliquota, " & _
'        "IPIAliquotaManual, IPIValor, IPIValorManual, ICMSTipo, ICMSTipoManual, ICMSBase, ICMSBaseManual, ICMSPercRedBase, ICMSPercRedBaseManual, ICMSAliquota, ICMSAliquotaManual, ICMSValor, ICMSValorManual, ICMSSubstBase, ICMSSubstBaseManual, " & _
'        "ICMSSubstAliquota, ICMSSubstAliquotaManual, ICMSSubstValor, ICMSSubstValorManual, PISCRedito, COFINSCredito, IPICredito, ICMSCredito FROM TributacaoComplPSRV WHERE FilialEmpresa = ? AND CodPedido = ?", _
'        tTributacaoComplPV.iItem, tTributacaoComplPV.sNaturezaOp, tTributacaoComplPV.iNaturezaOpManual, tTributacaoComplPV.iTipoTributacao, tTributacaoComplPV.iTipoTributacaoManual, tTributacaoComplPV.iIPITipo, tTributacaoComplPV.iIPITipoManual, tTributacaoComplPV.dIPIBaseCalculo, tTributacaoComplPV.iIPIBaseManual, tTributacaoComplPV.dIPIPercRedBase, _
'        tTributacaoComplPV.iIPIPercRedBaseManual, tTributacaoComplPV.dIPIAliquota, tTributacaoComplPV.iIPIAliquotaManual, tTributacaoComplPV.dIPIValor, tTributacaoComplPV.iIPIValorManual, tTributacaoComplPV.iICMSTipo, tTributacaoComplPV.iICMSTipoManual, tTributacaoComplPV.dICMSBase, tTributacaoComplPV.iICMSBaseManual, tTributacaoComplPV.dICMSPercRedBase, _
'        tTributacaoComplPV.iICMSPercRedBaseManual, tTributacaoComplPV.dICMSAliquota, tTributacaoComplPV.iICMSAliquotaManual, tTributacaoComplPV.dICMSValor, tTributacaoComplPV.iICMSValorManual, tTributacaoComplPV.dICMSSubstBase, tTributacaoComplPV.iICMSSubstBaseManual, _
'        tTributacaoComplPV.dICMSSubstAliquota, tTributacaoComplPV.iICMSSubstAliquotaManual, tTributacaoComplPV.dICMSSubstValor, tTributacaoComplPV.iICMSSubstValorManual, tTributacaoComplPV.dPISCredito, tTributacaoComplPV.dCOFINSCredito, tTributacaoComplPV.dIPICredito, tTributacaoComplPV.dICMSCredito, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo)
'    If lErro <> AD_SQL_SUCESSO Then gError 188841
'
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 188842
'
'    Do While lErro <> AD_SQL_SEM_DADOS
'
'        Set objTributacaoItemPV = New ClassTributacaoItemPV
'
'        With objTributacaoItemPV
'
'            .lNumIntDoc = tTributacaoComplPV.iItem
'            .sNaturezaOp = tTributacaoComplPV.sNaturezaOp
'            .iNaturezaOpManual = tTributacaoComplPV.iNaturezaOpManual
'            .iTipoTributacao = tTributacaoComplPV.iTipoTributacao
'            .iTipoTributacaoManual = tTributacaoComplPV.iTipoTributacaoManual
'            .iIPITipo = tTributacaoComplPV.iIPITipo
'            .iIPITipoManual = tTributacaoComplPV.iIPITipoManual
'            .dIPIBaseCalculo = tTributacaoComplPV.dIPIBaseCalculo
'            .iIPIBaseManual = tTributacaoComplPV.iIPIBaseManual
'            .dIPIPercRedBase = tTributacaoComplPV.dIPIPercRedBase
'            .iIPIPercRedBaseManual = tTributacaoComplPV.iIPIPercRedBaseManual
'            .dIPIAliquota = tTributacaoComplPV.dIPIAliquota
'            .iIPIAliquotaManual = tTributacaoComplPV.iIPIAliquotaManual
'            .dIPIValor = tTributacaoComplPV.dIPIValor
'            .iIPIValorManual = tTributacaoComplPV.iIPIValorManual
'            .iICMSTipo = tTributacaoComplPV.iICMSTipo
'            .iICMSTipoManual = tTributacaoComplPV.iICMSTipoManual
'            .dICMSBase = tTributacaoComplPV.dICMSBase
'            .iICMSBaseManual = tTributacaoComplPV.iICMSBaseManual
'            .dICMSPercRedBase = tTributacaoComplPV.dICMSPercRedBase
'            .iICMSPercRedBaseManual = tTributacaoComplPV.iICMSPercRedBaseManual
'            .dICMSAliquota = tTributacaoComplPV.dICMSAliquota
'            .iICMSAliquotaManual = tTributacaoComplPV.iICMSAliquotaManual
'            .dICMSValor = tTributacaoComplPV.dICMSValor
'            .iICMSValorManual = tTributacaoComplPV.iICMSValorManual
'            .dICMSSubstBase = tTributacaoComplPV.dICMSSubstBase
'            .iICMSSubstBaseManual = tTributacaoComplPV.iICMSSubstBaseManual
'            .dICMSSubstAliquota = tTributacaoComplPV.dICMSSubstAliquota
'            .iICMSSubstAliquotaManual = tTributacaoComplPV.iICMSSubstAliquotaManual
'            .dICMSSubstValor = tTributacaoComplPV.dICMSSubstValor
'            .iICMSSubstValorManual = tTributacaoComplPV.iICMSSubstValorManual
'            .dICMSCredito = tTributacaoComplPV.dICMSCredito
'            .dIPICredito = tTributacaoComplPV.dIPICredito
'            .dCOFINSCredito = tTributacaoComplPV.dCOFINSCredito
'            .dPISCredito = tTributacaoComplPV.dPISCredito
'        End With
'
'        Select Case tTributacaoComplPV.iItem
'
'            Case ITEM_TIPO_FRETE
'                Set objTributacaoPV.objTributacaoFrete = objTributacaoItemPV
'
'            Case ITEM_TIPO_SEGURO
'                Set objTributacaoPV.objTributacaoSeguro = objTributacaoItemPV
'
'            Case ITEM_TIPO_DESCONTO
'                Set objTributacaoPV.objTributacaoDesconto = objTributacaoItemPV
'
'            Case ITEM_TIPO_OUTRAS_DESP
'                Set objTributacaoPV.objTributacaoOutras = objTributacaoItemPV
'
'            Case Else 'item invalido
'                gError 188843
'
'        End Select
'
'        lErro = Comando_BuscarProximo(lComando)
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 188844
'
'    Loop
'
'    Call Comando_Fechar(lComando)
'
'    PedidoServico_Le_Complementos = SUCESSO
'
'    Exit Function
'
'Erro_PedidoServico_Le_Complementos:
'
'    PedidoServico_Le_Complementos = gErr
'
'    Select Case gErr
'
'        Case 188840
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 188841, 188842, 188844
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRIBUTACAOCOMPLPEDSRV", gErr, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo)
'
'        Case 188843
'            Call Rotina_Erro(vbOKOnly, "ERRO_COMPL_PEDSRV_TIPO_INVALIDO", gErr)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 188845)
'
'    End Select
'
'    Call Comando_Fechar(lComando)
'
'    Exit Function
'
'End Function

Function ParcelasPedidoSRV_Le(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
'Lê as parcelas associadas ao objPedidoVenda

Dim lErro As Long
Dim lComando As Long
Dim tParcelaPV As typeParcelaPedidoVenda
Dim objParcPV As ClassParcelaPedidoVenda

On Error GoTo Erro_ParcelasPedidoSRV_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 188846

    tParcelaPV.sAgenciaCheque = String(STRING_AGENCIA, 0)
    tParcelaPV.sAprovacaoCartao = String(STRING_APROVACAO_CARTAO, 0)
    tParcelaPV.sContaCorrenteCheque = String(STRING_CONTA_CORRENTE, 0)
    tParcelaPV.sNumeroCartao = String(STRING_NUMERO_CARTAO, 0)

    'Seleciona Parcelas vinculadas ao Pedido de Venda
    lErro = Comando_Executar(lComando, "SELECT NumParcela, DataVencimento, Valor, Desconto1Codigo, Desconto1Ate, Desconto1Valor, Desconto2Codigo, Desconto2Ate, Desconto2Valor, Desconto3Codigo, Desconto3Ate, Desconto3Valor, TipoPagto, CodConta, DataCredito, DataEmissaoCheque, BancoCheque, AgenciaCheque, ContaCorrenteCheque, NumeroCheque, DataDepositoCheque, AdmMeioPagto, Parcelamento, NumeroCartao, ValidadeCartao, AprovacaoCartao, DataTransacaoCartao FROM ParcelasPedidoSRV WHERE PedidoSRV = ? AND FilialEmpresa = ? ORDER BY DataVencimento", _
    tParcelaPV.iNumParcela, tParcelaPV.dtDataVencimento, tParcelaPV.dValor, tParcelaPV.iDesconto1Codigo, tParcelaPV.dtDesconto1Ate, tParcelaPV.dDesconto1Valor, tParcelaPV.iDesconto2Codigo, tParcelaPV.dtDesconto2Ate, tParcelaPV.dDesconto2Valor, tParcelaPV.iDesconto3Codigo, tParcelaPV.dtDesconto3Ate, tParcelaPV.dDesconto3Valor, _
    tParcelaPV.iTipoPagto, tParcelaPV.iCodConta, tParcelaPV.dtDataCredito, tParcelaPV.dtDataEmissaoCheque, tParcelaPV.iBancoCheque, tParcelaPV.sAgenciaCheque, tParcelaPV.sContaCorrenteCheque, tParcelaPV.lNumeroCheque, tParcelaPV.dtDataDepositoCheque, tParcelaPV.iAdmMeioPagto, tParcelaPV.iParcelamento, tParcelaPV.sNumeroCartao, tParcelaPV.dtValidadeCartao, tParcelaPV.sAprovacaoCartao, tParcelaPV.dtDataTransacaoCartao, objPedidoVenda.lCodigo, objPedidoVenda.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 188847

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 188848

    Do While lErro <> AD_SQL_SEM_DADOS

        With tParcelaPV
            Set objParcPV = objPedidoVenda.colParcelas.Add(.dValor, .dtDataVencimento, .iNumParcela, .iDesconto1Codigo, .dtDesconto1Ate, .dDesconto1Valor, .iDesconto2Codigo, .dtDesconto2Ate, .dDesconto2Valor, .dtDesconto3Ate, .dDesconto3Valor, .iDesconto3Codigo)
        
            objParcPV.iTipoPagto = .iTipoPagto
            objParcPV.iCodConta = .iCodConta
            objParcPV.dtDataCredito = .dtDataCredito
            objParcPV.dtDataEmissaoCheque = .dtDataEmissaoCheque
            objParcPV.iBancoCheque = .iBancoCheque
            objParcPV.sAgenciaCheque = .sAgenciaCheque
            objParcPV.sContaCorrenteCheque = .sContaCorrenteCheque
            objParcPV.lNumeroCheque = .lNumeroCheque
            objParcPV.dtDataDepositoCheque = .dtDataDepositoCheque
            objParcPV.iAdmMeioPagto = .iAdmMeioPagto
            objParcPV.iParcelamento = .iParcelamento
            objParcPV.sNumeroCartao = .sNumeroCartao
            objParcPV.dtValidadeCartao = .dtValidadeCartao
            objParcPV.sAprovacaoCartao = .sAprovacaoCartao
            objParcPV.dtDataTransacaoCartao = .dtDataTransacaoCartao
        
        End With

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 188849

    Loop

    Call Comando_Fechar(lComando)

    ParcelasPedidoSRV_Le = SUCESSO

    Exit Function

Erro_ParcelasPedidoSRV_Le:

    ParcelasPedidoSRV_Le = gErr

    Select Case gErr

        Case 188846
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 188847, 188848, 188849
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PARCELASPEDIDOSRV", gErr, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 188850)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ComissoesPedidoSRV_Le(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
'Lê as Comissões associadas ao Pedido de Venda passado e objPedidoVenda.

'*** Alterada em 05/03/02 por Luiz Nogueira ***
'*** ClassPedidoDeVenda teve o tipo de ColComissoes alterado, por isso foi necessário adicionar
'os dados lidos a um obj e depois adicionar o obj à coleção ***

'Alterada por Tulio em 26/03/02
'Incluindo campo DiretoIndireto

Dim lErro As Long
Dim lComando As Long
Dim tComissaoPV As typeComissaoPedVenda
Dim objComissaoPV As ClassComissaoPedVendas

On Error GoTo Erro_ComissoesPedidoSRV_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 188851
    
    lErro = Comando_Executar(lComando, "SELECT CodVendedor, ValorBase, Percentual, Valor, PercentualEmissao, ValorEmissao, Indireta, Seq FROM ComissoesPedidoSRV WHERE FilialEmpresa = ? AND PedidoSRV = ? ORDER BY Seq", tComissaoPV.iCodVendedor, tComissaoPV.dValorBase, tComissaoPV.dPercentual, tComissaoPV.dValor, tComissaoPV.dPercentualEmissao, tComissaoPV.dValorEmissao, tComissaoPV.iIndireta, tComissaoPV.iSeq, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 188852
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 188853
    
    Do While lErro <> AD_SQL_SEM_DADOS
    
        'Instancia um novo objComissaoPV
        Set objComissaoPV = New ClassComissaoPedVendas
        
        'Transfere os dados lidos para o obj
        With tComissaoPV
            
            objComissaoPV.iCodVendedor = .iCodVendedor
            objComissaoPV.dValorBase = .dValorBase
            objComissaoPV.dPercentual = .dPercentual
            objComissaoPV.dValor = .dValor
            objComissaoPV.dPercentualEmissao = .dPercentualEmissao
            objComissaoPV.dValorEmissao = .dValorEmissao
            objComissaoPV.iIndireta = .iIndireta
            objComissaoPV.iSeq = .iSeq

        End With
    
        'Adiciona o obj à coleção
        objPedidoVenda.colComissoes.Add objComissaoPV
            
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 188854
        
    Loop
    
    Call Comando_Fechar(lComando)
    
    ComissoesPedidoSRV_Le = SUCESSO
    
    Exit Function
    
Erro_ComissoesPedidoSRV_Le:

    ComissoesPedidoSRV_Le = gErr
    
    Select Case gErr
    
        Case 188851
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 188852 To 188854
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMISSOESPEDIDOSRV", gErr, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 188855)
            
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
 
End Function

Function PedidoSRV_Le_Itens_ComTributacao(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long

Dim lComando As Long, tItemPedido As typeItemPedido, objItemPedido As ClassItemPedido
'Dim tTributacaoItemPV As typeTributacaoItemPV, objTributacaoItemPV As ClassTributacaoItemPV
Dim lErro As Long

On Error GoTo Erro_PedidoSRV_Le_Itens_ComTributacao

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 188861
    
    tItemPedido.sDescricao = String(STRING_ITEM_PEDIDO_DESCRICAO, 0)
    tItemPedido.sLote = String(STRING_ITEM_PEDIDO_LOTE, 0)
    tItemPedido.sProduto = String(STRING_PRODUTO, 0)
    tItemPedido.sProdutoDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
    tItemPedido.sProdutoNomeReduzido = String(STRING_PRODUTO_NOME_REDUZIDO, 0)
    tItemPedido.sUMEstoque = String(STRING_PRODUTO_SIGLAUMESTOQUE, 0)
    tItemPedido.sUnidadeMed = String(STRING_UM_SIGLA, 0)

'    With tTributacaoItemPV
'
'        .sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
'
'        lErro = Comando_Executar(lComando, "SELECT ItensPedidoSRV.NumIntDoc,ItensPedidoSRV.Produto, Quantidade, PrecoUnitario, PrecoTotal, ItensPedidoSRV.ClasseUM, UnidadeMed, ValorDesconto, DataEntrega, ItensPedidoSRV.Descricao, " & _
'        "Lote, Status, ValorAbatComissao, QuantCancelada, QuantReservada, QuantFaturada, ItensPedidoSRV.Descricao, NomeReduzido, SiglaUMEstoque, ControleEstoque, " & _
'        "NaturezaOp, NaturezaOpManual, TipoTributacao, TipoTributacaoManual, IPITipo, IPITipoManual, IPIBaseCalculo, IPIBaseManual, IPIPercRedBase, IPIPercRedBaseManual, TributacaoItemPedidoSRV.IPIAliquota, IPIAliquotaManual, IPIValor, IPIValorManual, ICMSTipo, ICMSTipoManual, ICMSBase, ICMSBaseManual, ICMSPercRedBase, ICMSPercRedBaseManual, ICMSAliquota, ICMSAliquotaManual, ICMSValor, ICMSValorManual, " & _
'        "ICMSSubstBase, ICMSSubstBaseManual, ICMSSubstAliquota, ICMSSubstAliquotaManual, ICMSSubstValor, ICMSSubstValorManual, Peca, PISCRedito, COFINSCredito, IPICredito, ICMSCredito " & _
'        "FROM ItensPedidoSRV, Produtos, TributacaoItemPedidoSRV WHERE ItensPedidoSRV.FilialEmpresa = ? AND ItensPedidoSRV.CodPedido = ? AND ItensPedidoSRV.Produto = Produtos.Codigo AND ItensPedidoSRV.NumIntDoc = TributacaoItemPedidoSRV.NumIntDoc ORDER BY ItensPedidoSRV.NumIntDoc", _
'        tItemPedido.lNumIntDoc, tItemPedido.sProduto, tItemPedido.dQuantidade, tItemPedido.dPrecoUnitario, tItemPedido.dPrecoTotal, tItemPedido.iClasseUM, tItemPedido.sUnidadeMed, tItemPedido.dValorDesconto, tItemPedido.dtDataEntrega, _
'        tItemPedido.sDescricao, tItemPedido.sLote, tItemPedido.iStatus, tItemPedido.dValorAbatComissao, tItemPedido.dQuantCancelada, tItemPedido.dQuantReservada, tItemPedido.dQuantFaturada, tItemPedido.sProdutoDescricao, _
'        tItemPedido.sProdutoNomeReduzido, tItemPedido.sUMEstoque, tItemPedido.iControleEstoque, _
'        tTributacaoItemPV.sNaturezaOp, tTributacaoItemPV.iNaturezaOpManual, tTributacaoItemPV.iTipoTributacao, tTributacaoItemPV.iTipoTributacaoManual, tTributacaoItemPV.iIPITipo, tTributacaoItemPV.iIPITipoManual, _
'        tTributacaoItemPV.dIPIBaseCalculo, tTributacaoItemPV.iIPIBaseManual, _
'        tTributacaoItemPV.dIPIPercRedBase, tTributacaoItemPV.iIPIPercRedBaseManual, _
'        tTributacaoItemPV.dIPIAliquota, tTributacaoItemPV.iIPIAliquotaManual, tTributacaoItemPV.dIPIValor, tTributacaoItemPV.iIPIValorManual, _
'        tTributacaoItemPV.iICMSTipo, tTributacaoItemPV.iICMSTipoManual, tTributacaoItemPV.dICMSBase, tTributacaoItemPV.iICMSBaseManual, tTributacaoItemPV.dICMSPercRedBase, tTributacaoItemPV.iICMSPercRedBaseManual, tTributacaoItemPV.dICMSAliquota, tTributacaoItemPV.iICMSAliquotaManual, tTributacaoItemPV.dICMSValor, tTributacaoItemPV.iICMSValorManual, tTributacaoItemPV.dICMSSubstBase, tTributacaoItemPV.iICMSSubstBaseManual, tTributacaoItemPV.dICMSSubstAliquota, tTributacaoItemPV.iICMSSubstAliquotaManual, tTributacaoItemPV.dICMSSubstValor, tTributacaoItemPV.iICMSSubstValorManual, _
'        tItemPedido.iPeca, tTributacaoItemPV.dPISCredito, tTributacaoItemPV.dCOFINSCredito, tTributacaoItemPV.dIPICredito, tTributacaoItemPV.dICMSCredito, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo)
'        If lErro <> AD_SQL_SUCESSO Then gError 188862
'
'    End With

    lErro = Comando_Executar(lComando, "SELECT ItensPedidoSRV.NumIntDoc,ItensPedidoSRV.Produto, Quantidade, PrecoUnitario, PrecoTotal, ItensPedidoSRV.ClasseUM, UnidadeMed, ValorDesconto, DataEntrega, ItensPedidoSRV.Descricao, " & _
    "Lote, Status, ValorAbatComissao, QuantCancelada, QuantReservada, QuantFaturada, ItensPedidoSRV.Descricao, NomeReduzido, SiglaUMEstoque, ControleEstoque, Peca " & _
    "FROM ItensPedidoSRV, Produtos  WHERE ItensPedidoSRV.FilialEmpresa = ? AND ItensPedidoSRV.CodPedido = ? AND ItensPedidoSRV.Produto = Produtos.Codigo ORDER BY ItensPedidoSRV.NumIntDoc", _
    tItemPedido.lNumIntDoc, tItemPedido.sProduto, tItemPedido.dQuantidade, tItemPedido.dPrecoUnitario, tItemPedido.dPrecoTotal, tItemPedido.iClasseUM, tItemPedido.sUnidadeMed, tItemPedido.dValorDesconto, tItemPedido.dtDataEntrega, _
    tItemPedido.sDescricao, tItemPedido.sLote, tItemPedido.iStatus, tItemPedido.dValorAbatComissao, tItemPedido.dQuantCancelada, tItemPedido.dQuantReservada, tItemPedido.dQuantFaturada, tItemPedido.sProdutoDescricao, _
    tItemPedido.sProdutoNomeReduzido, tItemPedido.sUMEstoque, tItemPedido.iControleEstoque, tItemPedido.iPeca, _
    objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 188862

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 188863

    Do While lErro <> AD_SQL_SEM_DADOS
        
'        Set objTributacaoItemPV = New ClassTributacaoItemPV
'
'        With objTributacaoItemPV
'
'            .lNumIntDoc = tTributacaoItemPV.lNumIntDoc
'            .sNaturezaOp = tTributacaoItemPV.sNaturezaOp
'            .iNaturezaOpManual = tTributacaoItemPV.iNaturezaOpManual
'            .iTipoTributacao = tTributacaoItemPV.iTipoTributacao
'            .iTipoTributacaoManual = tTributacaoItemPV.iTipoTributacaoManual
'            .iIPITipo = tTributacaoItemPV.iIPITipo
'            .iIPITipoManual = tTributacaoItemPV.iIPITipoManual
'            .dIPIBaseCalculo = tTributacaoItemPV.dIPIBaseCalculo
'            .iIPIBaseManual = tTributacaoItemPV.iIPIBaseManual
'            .dIPIPercRedBase = tTributacaoItemPV.dIPIPercRedBase
'            .iIPIPercRedBaseManual = tTributacaoItemPV.iIPIPercRedBaseManual
'            .dIPIAliquota = tTributacaoItemPV.dIPIAliquota
'            .iIPIAliquotaManual = tTributacaoItemPV.iIPIAliquotaManual
'            .dIPIValor = tTributacaoItemPV.dIPIValor
'            .iIPIValorManual = tTributacaoItemPV.iIPIValorManual
'            .iICMSTipo = tTributacaoItemPV.iICMSTipo
'            .iICMSTipoManual = tTributacaoItemPV.iICMSTipoManual
'            .dICMSBase = tTributacaoItemPV.dICMSBase
'            .iICMSBaseManual = tTributacaoItemPV.iICMSBaseManual
'            .dICMSPercRedBase = tTributacaoItemPV.dICMSPercRedBase
'            .iICMSPercRedBaseManual = tTributacaoItemPV.iICMSPercRedBaseManual
'            .dICMSAliquota = tTributacaoItemPV.dICMSAliquota
'            .iICMSAliquotaManual = tTributacaoItemPV.iICMSAliquotaManual
'            .dICMSValor = tTributacaoItemPV.dICMSValor
'            .iICMSValorManual = tTributacaoItemPV.iICMSValorManual
'            .dICMSSubstBase = tTributacaoItemPV.dICMSSubstBase
'            .iICMSSubstBaseManual = tTributacaoItemPV.iICMSSubstBaseManual
'            .dICMSSubstAliquota = tTributacaoItemPV.dICMSSubstAliquota
'            .iICMSSubstAliquotaManual = tTributacaoItemPV.iICMSSubstAliquotaManual
'            .dICMSSubstValor = tTributacaoItemPV.dICMSSubstValor
'            .iICMSSubstValorManual = tTributacaoItemPV.iICMSSubstValorManual
'            .dICMSCredito = tTributacaoItemPV.dICMSCredito
'            .dIPICredito = tTributacaoItemPV.dIPICredito
'            .dCOFINSCredito = tTributacaoItemPV.dCOFINSCredito
'            .dPISCredito = tTributacaoItemPV.dPISCredito
'        End With
'
        Set objItemPedido = New ClassItemPedido
'
'        Set objItemPedido.objTributacaoItemPV = objTributacaoItemPV
        
        objItemPedido.iFilialEmpresa = objPedidoVenda.iFilialEmpresa
        objItemPedido.lCodPedido = objPedidoVenda.lCodigo
        objItemPedido.sProduto = tItemPedido.sProduto
        objItemPedido.dQuantidade = tItemPedido.dQuantidade
        objItemPedido.dPrecoUnitario = tItemPedido.dPrecoUnitario
        objItemPedido.dPrecoTotal = tItemPedido.dPrecoTotal
        objItemPedido.dValorDesconto = tItemPedido.dValorDesconto
        objItemPedido.dtDataEntrega = tItemPedido.dtDataEntrega
        objItemPedido.sProdutoDescricao = tItemPedido.sProdutoDescricao
        objItemPedido.dValorAbatComissao = tItemPedido.dValorAbatComissao
        objItemPedido.dQuantCancelada = tItemPedido.dQuantCancelada
        objItemPedido.dQuantReservada = tItemPedido.dQuantReservada
        Set objItemPedido.ColReserva = tItemPedido.colReservaItem
        objItemPedido.sProdutoNomeReduzido = tItemPedido.sProdutoNomeReduzido
        objItemPedido.sUnidadeMed = tItemPedido.sUnidadeMed
        objItemPedido.sLote = tItemPedido.sLote
        objItemPedido.sUMEstoque = tItemPedido.sUMEstoque
        objItemPedido.iClasseUM = tItemPedido.iClasseUM
        objItemPedido.dQuantFaturada = tItemPedido.dQuantFaturada
        objItemPedido.dQuantAFaturar = 0
        objItemPedido.sDescricao = tItemPedido.sDescricao
        objItemPedido.iStatus = tItemPedido.iStatus
        objItemPedido.iControleEstoque = tItemPedido.iControleEstoque
        objItemPedido.dQuantOP = tItemPedido.dQuantOP
        objItemPedido.dQuantSC = tItemPedido.dQuantSC
        objItemPedido.lNumIntDoc = tItemPedido.lNumIntDoc
        objItemPedido.iPeca = tItemPedido.iPeca
        
        objItemPedido.objTributacao.iTipoDoc = TIPODOC_TRIB_PSRV
        objItemPedido.objTributacao.lNumIntDocItem = objItemPedido.lNumIntDoc
        lErro = CF("TributacaoDocItem_Le", objItemPedido.objTributacao)
        If lErro <> SUCESSO Then gError 186603
                
        If tItemPedido.iPeca = 0 Then
            objPedidoVenda.colItensPedidoSRV.AddObj objItemPedido
        Else
            objPedidoVenda.colItensPedido.AddObj objItemPedido
        End If

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 188864
        
    Loop
    
    Call Comando_Fechar(lComando)
    
    PedidoSRV_Le_Itens_ComTributacao = SUCESSO

    Exit Function

Erro_PedidoSRV_Le_Itens_ComTributacao:

    PedidoSRV_Le_Itens_ComTributacao = gErr

    Select Case gErr
    
        Case 186603

        Case 188861
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 188862 To 188864
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDIDOSRV", gErr, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 188865)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function OrcamentoSRV_Le_NumIntPedSRV(objOrcamentoVenda As ClassOrcamentoVenda) As Long
'Lê o Orcamento de Venda a partir do NumIntPedVenda passado.

Dim lComando As Long
Dim lErro As Long
Dim tOrcamentoVenda As typeOrcamentoVenda
Dim iPrazoValidade As Integer

On Error GoTo Erro_OrcamentoSRV_Le_NumIntPedSRV

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 188866

    tOrcamentoVenda.sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
    tOrcamentoVenda.sNomeCli = String(STRING_CLIENTE_NOME_REDUZIDO, 0)
    tOrcamentoVenda.sNomeFilialCli = String(STRING_FILIAL_CLIENTE_NOME, 0)
    
    lErro = Comando_Executar(lComando, "SELECT NaturezaOp, NumIntNFiscal, Codigo, FilialEmpresa, NomeCli, NomeFilialCli, Vendedor, Cliente, Filial, CondicaoPagto, PercAcrescFinanceiro, DataEmissao,  ValorTotal, ValorFrete, ValorDesconto, ValorSeguro, TabelaPreco, DataReferencia, NumIntDoc,ValorOutrasDespesas,ValorProdutos, PrazoValidade FROM OrcamentoSRV WHERE NumIntPedSRV = ? ", _
    tOrcamentoVenda.sNaturezaOp, tOrcamentoVenda.lNumIntNFiscal, tOrcamentoVenda.lCodigo, tOrcamentoVenda.iFilialEmpresa, tOrcamentoVenda.sNomeCli, tOrcamentoVenda.sNomeFilialCli, tOrcamentoVenda.iVendedor, tOrcamentoVenda.lCliente, tOrcamentoVenda.iFilial, tOrcamentoVenda.iCondicaoPagto, tOrcamentoVenda.dPercAcrescFinanceiro, tOrcamentoVenda.dtDataEmissao, tOrcamentoVenda.dValorTotal, tOrcamentoVenda.dValorFrete, tOrcamentoVenda.dValorDesconto, tOrcamentoVenda.dValorSeguro, tOrcamentoVenda.iTabelaPreco, tOrcamentoVenda.dtDataReferencia, tOrcamentoVenda.lNumIntDoc, tOrcamentoVenda.dValorOutrasDespesas, tOrcamentoVenda.dValorProdutos, iPrazoValidade, objOrcamentoVenda.lNumIntPedVenda)
    If lErro <> AD_SQL_SUCESSO Then gError 188867

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 188868
    
    'Se não encontrou ---> Erro tratado na rotina Chamadora
    If lErro = AD_SQL_SEM_DADOS Then gError 188869

    objOrcamentoVenda.lCodigo = tOrcamentoVenda.lCodigo
    objOrcamentoVenda.iFilialEmpresa = tOrcamentoVenda.iFilialEmpresa
    objOrcamentoVenda.lNumIntNFiscal = tOrcamentoVenda.lNumIntNFiscal
    objOrcamentoVenda.sNomeFilialCli = tOrcamentoVenda.sNomeFilialCli
    objOrcamentoVenda.sNomeCli = tOrcamentoVenda.sNomeCli
    objOrcamentoVenda.iPrazoValidade = iPrazoValidade
    objOrcamentoVenda.lCliente = tOrcamentoVenda.lCliente
    objOrcamentoVenda.iFilial = tOrcamentoVenda.iFilial
    objOrcamentoVenda.iCondicaoPagto = tOrcamentoVenda.iCondicaoPagto
    objOrcamentoVenda.dPercAcrescFinanceiro = tOrcamentoVenda.dPercAcrescFinanceiro
    objOrcamentoVenda.dtDataEmissao = tOrcamentoVenda.dtDataEmissao
    objOrcamentoVenda.sNaturezaOp = tOrcamentoVenda.sNaturezaOp
    objOrcamentoVenda.dValorTotal = tOrcamentoVenda.dValorTotal
    objOrcamentoVenda.dValorFrete = tOrcamentoVenda.dValorFrete
    objOrcamentoVenda.dValorDesconto = tOrcamentoVenda.dValorDesconto
    objOrcamentoVenda.dValorOutrasDespesas = tOrcamentoVenda.dValorOutrasDespesas
    objOrcamentoVenda.dValorSeguro = tOrcamentoVenda.dValorSeguro
    objOrcamentoVenda.iTabelaPreco = tOrcamentoVenda.iTabelaPreco
    objOrcamentoVenda.dtDataReferencia = tOrcamentoVenda.dtDataReferencia
    objOrcamentoVenda.lNumIntDoc = tOrcamentoVenda.lNumIntDoc
    objOrcamentoVenda.dValorProdutos = tOrcamentoVenda.dValorProdutos
    objOrcamentoVenda.iVendedor = tOrcamentoVenda.iVendedor
    
    Call Comando_Fechar(lComando)

    OrcamentoSRV_Le_NumIntPedSRV = SUCESSO

    Exit Function

Erro_OrcamentoSRV_Le_NumIntPedSRV:

    OrcamentoSRV_Le_NumIntPedSRV = gErr

    Select Case gErr

        Case 188866
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 188867, 188868
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORCAMENTOSRV1", gErr, objOrcamentoVenda.lNumIntPedVenda)
        
        Case 188869

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 188870)
    
    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ReservasPedidoSRV_Le(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
'Lê as Reservas dos itens do Pedido passado

Dim lComando As Long
Dim lErro As Long
Dim tReserva As typeReserva
Dim objReserva As ClassReserva
Dim objItemPV As ClassItemPedido
Dim objItemRomaneio As ClassItemRomaneioGrade

On Error GoTo Erro_ReservasPedidoSRV_Le

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 188871

    'Para cada item do Pedido
    For Each objItemPV In objPedidoVenda.colItensPedido

        tReserva.sCodUsuario = String(STRING_RESERVA_USUARIO, 0)
        tReserva.sProduto = String(STRING_PRODUTO, 0)
        tReserva.sResponsavel = String(STRING_RESERVA_RESPONSAVEL, 0)

        'Seleciona reservas
        lErro = Comando_Executar(lComando, "SELECT Codigo,Produto,Almoxarifado,NumIntOrigem,Quantidade,DataReserva,DataValidade,CodUsuario,Responsavel,NumIntDoc FROM Reserva WHERE Filialempresa = ? AND TipoDoc = ? AND DocOrigem = ? AND NumIntOrigem = ?", tReserva.lCodigo, tReserva.sProduto, tReserva.iAlmoxarifado, tReserva.lNumIntOrigem, tReserva.dQuantidade, tReserva.dtDataReserva, tReserva.dtDataValidade, tReserva.sCodUsuario, tReserva.sResponsavel, tReserva.lNumIntDoc, objPedidoVenda.iFilialEmpresa, TIPO_PEDIDO_SRV_COD, objPedidoVenda.lCodigo, objItemPV.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 188872

        'Tenta ler a primeira
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 188873

        Do While lErro = AD_SQL_SUCESSO

            'Adiciona a reserva na coleção de reserva do item
            With tReserva
                objItemPV.ColReserva.Add objPedidoVenda.iFilialEmpresa, 0, .sProduto, .iAlmoxarifado, 0, 0, 0, .dQuantidade, DATA_NULA, .dtDataValidade, "", .sResponsavel, .lNumIntDoc, ""
            End With

            'Busca próxima reserva
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then gError 188874

        Loop
        

    Next

    'Fecha o comando
    Call Comando_Fechar(lComando)

    ReservasPedidoSRV_Le = SUCESSO

    Exit Function

Erro_ReservasPedidoSRV_Le:

    ReservasPedidoSRV_Le = gErr

    Select Case gErr

        Case 188871
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 188872 To 188874
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RESERVA1", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188875)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function
'
'Function BloqueiosPedidoSRV_Le(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
''Lê os Bloqueios associados ao objPedidoVenda
'
'Dim lErro As Long
'Dim lComando As Long
'Dim tBloqueioPV As typeBloqueioPV
'
'On Error GoTo Erro_BloqueiosPedidoSRV_Le
'
'    lComando = Comando_Abrir()
'    If lComando = 0 Then gError 191516
'
'    tBloqueioPV.sCodUsuario = String(STRING_BLOQUEIOSPV_COD_USUARIO, 0)
'    tBloqueioPV.sCodUsuarioLib = String(STRING_BLOQUEIOSPV_COD_USUARIO, 0)
'    tBloqueioPV.sResponsavel = String(STRING_BLOQUEIOSPV_RESPONSAVEL, 0)
'    tBloqueioPV.sResponsavelLib = String(STRING_BLOQUEIOSPV_RESPONSAVEL, 0)
'    tBloqueioPV.sObservacao = String(STRING_BLOQUEIOSPV_OBSERVACAO, 0)
'
'    lErro = Comando_Executar(lComando, "SELECT Sequencial, TipoDeBloqueio, CodUsuario, Responsavel, Data, CodUsuarioLib, ResponsavelLib, DataLib, Observacao FROM BloqueiosPedidoSRV WHERE FilialEmpresa = ? AND PedidoSRV = ? ORDER BY Sequencial", tBloqueioPV.iSequencial, tBloqueioPV.iTipoDeBloqueio, tBloqueioPV.sCodUsuario, tBloqueioPV.sResponsavel, tBloqueioPV.dtData, tBloqueioPV.sCodUsuarioLib, tBloqueioPV.sResponsavelLib, tBloqueioPV.dtDataLib, tBloqueioPV.sObservacao, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo)
'    If lErro <> AD_SQL_SUCESSO Then gError 191517
'
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 191518
'
'    Do While lErro <> AD_SQL_SEM_DADOS
'
'        With tBloqueioPV
'            objPedidoVenda.colBloqueiosPV.Add objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo, .iSequencial, .iTipoDeBloqueio, .sCodUsuario, .sResponsavel, .dtData, .sCodUsuarioLib, .sResponsavelLib, .dtDataLib, .sObservacao
'        End With
'
'        lErro = Comando_BuscarProximo(lComando)
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 191519
'
'    Loop
'
'    Call Comando_Fechar(lComando)
'
'    BloqueiosPedidoSRV_Le = SUCESSO
'
'    Exit Function
'
'Erro_BloqueiosPedidoSRV_Le:
'
'    BloqueiosPedidoSRV_Le = gErr
'
'    Select Case gErr
'
'        Case 191516
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 191517, 191518, 191519
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BLOQUEIOSPEDIDOSRV", Err, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 191520)
'
'    End Select
'
'    Call Comando_Fechar(lComando)
'
'    Exit Function
'
'End Function

Function PSRVObjetosVinculados_Lock(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
'Faz "lock" nos objetos associados ao Pedido de Venda

Dim lErro As Long
Dim alComando(1 To 1) As Long
Dim objProduto As New ClassProduto
Dim objItemPV As ClassItemPedido
Dim iIndice As Integer

On Error GoTo Erro_PSRVObjetosVinculados_Lock

    'Abe os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 191534
    Next

    'Para cada item de Pedido Venda
    For Each objItemPV In objPedidoVenda.colItensPedidoSRV

        objProduto.sCodigo = objItemPV.sProduto
        
        'Locka  o produto
        lErro = CF("Produto_Lock_Shared", alComando(1), objProduto)
        If lErro <> SUCESSO Then gError 191535

        'Armazena os dados do produto para que possam ser usados depois
        objItemPV.iControleEstoque = objProduto.iControleEstoque
        objItemPV.sUMEstoque = objProduto.sSiglaUMEstoque 'usa nas reservas
        objItemPV.iClasseUM = objProduto.iClasseUM
        objItemPV.iApropriacaoProd = objProduto.iApropriacaoCusto

    Next

    'fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

     PSRVObjetosVinculados_Lock = SUCESSO

    Exit Function

Erro_PSRVObjetosVinculados_Lock:

    PSRVObjetosVinculados_Lock = Err

    Select Case Err

        Case 191534
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 191535

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 191536)

    End Select

    'fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'Function ItemOP_Le_ItemPSRV(ByVal objItemOP As ClassItemOP, ByVal objItemPV As ClassItemPedido) As Long
''Lê o item de Ordem de Produção vinculado ao Item de Pedido de Venda passado
'
'Dim lErro As Long
'Dim lComando As Long
'Dim tItemOP As typeItemOP
'
'On Error GoTo Erro_ItemOP_Le_ItemPSRV
'
'    'Abre o comando
'    lComando = Comando_Abrir()
'    If lComando = 0 Then gError 191556
'
'    tItemOP.sCcl = String(STRING_BUFFER_MAX_TEXTO, 0)
'    tItemOP.sCodigo = String(STRING_BUFFER_MAX_TEXTO, 0)
'    tItemOP.sProduto = String(STRING_BUFFER_MAX_TEXTO, 0)
'    tItemOP.sSiglaUM = String(STRING_BUFFER_MAX_TEXTO, 0)
'
'    'Lê Item OP a partir do número interno do item de Pedido Venda (NumIntOrigem)
'    lErro = Comando_Executar(lComando, "SELECT NumIntDoc,FilialEmpresa,Codigo,Item,Produto,FilialPedido,CodPedido,Almoxarifado,Ccl,SiglaUM,Quantidade,DataInicioProd,DataFimProd,Prioridade,QuantidadeProd,Situacao,Destinacao FROM ItensOPTodosSRV WHERE NumIntOrigem = ?", tItemOP.lNumIntDoc, tItemOP.iFilialEmpresa, tItemOP.sCodigo, tItemOP.iItem, tItemOP.sProduto, tItemOP.iFilialPedido, tItemOP.lCodPedido, tItemOP.iAlmoxarifado, tItemOP.sCcl, tItemOP.sSiglaUM, tItemOP.dQuantidade, tItemOP.dtDataInicioProd, tItemOP.dtDataFimProd, tItemOP.iPrioridade, tItemOP.dQuantidadeProd, tItemOP.iSituacao, tItemOP.iDestinacao, objItemPV.lNumIntDoc)
'    If lErro <> AD_SQL_SUCESSO Then gError 191557
'
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then gError 191558
'
'    If lErro = AD_SQL_SEM_DADOS Then gError 191559
'
'    'Armazena os dados lidos
'    With tItemOP
'        objItemOP.dQuantidade = .dQuantidade
'        objItemOP.dQuantidadeProd = .dQuantidadeProd
'        objItemOP.dtDataFimProd = .dtDataFimProd
'        objItemOP.dtDataInicioProd = .dtDataInicioProd
'        objItemOP.iAlmoxarifado = .iAlmoxarifado
'        objItemOP.iDestinacao = .iDestinacao
'        objItemOP.iFilialEmpresa = .iFilialEmpresa
'        objItemOP.iFilialPedido = .iFilialPedido
'        objItemOP.iItem = .iItem
'        objItemOP.iPrioridade = .iPrioridade
'        objItemOP.iSituacao = .iSituacao
'        objItemOP.lCodPedido = .lCodPedido
'        objItemOP.lNumIntDoc = .lNumIntDoc
'        objItemOP.lNumIntOrigem = objItemPV.lNumIntDoc
'        objItemOP.sCcl = .sCcl
'        objItemOP.sCodigo = .sCodigo
'        objItemOP.sProduto = .sProduto
'        objItemOP.sSiglaUM = .sSiglaUM
'
'    End With
'
'    'Fecha o comando
'    Call Comando_Fechar(lComando)
'
'    ItemOP_Le_ItemPSRV = SUCESSO
'
'    Exit Function
'
'Erro_ItemOP_Le_ItemPSRV:
'
'    ItemOP_Le_ItemPSRV = gErr
'
'    Select Case gErr
'
'        Case 191556
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 191557, 191558
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSOP", gErr)
'
'        Case 191559
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 191560)
'
'    End Select
'
'    Call Comando_Fechar(lComando)
'
'    Exit Function
'
'End Function

Function PedidoServico_Lock(ByVal objPedidoVenda As ClassPedidoDeVenda, ByVal lComando As Long) As Long
'Lê e faz o Lock no Pedido de Servico

Dim lErro As Long
Dim tPedVenda As typePedidoVenda
Dim lCodigo As Long

On Error GoTo Erro_PedidoServico_Lock

    'Lê o Pedido com o Código e a filialempresa passados
    lErro = Comando_ExecutarPos(lComando, "SELECT NumIntDoc, ValorTotal, FilialEmpresaFaturamento FROM PedidoServico WHERE FilialEmpresa = ? and Codigo = ?", 0, tPedVenda.lNumIntDoc, tPedVenda.dValorTotal, tPedVenda.iFilialEmpresaFaturamento, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 191681

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 191682
    
    If lErro = AD_SQL_SEM_DADOS Then gError 191683

    'Faz o "Lock" no Pedido de Venda
    lErro = Comando_LockShared(lComando)
    If lErro <> AD_SQL_SUCESSO Then gError 191684

    objPedidoVenda.lNumIntDoc = tPedVenda.lNumIntDoc
    objPedidoVenda.dValorTotal = tPedVenda.dValorTotal
    objPedidoVenda.iFilialEmpresaFaturamento = tPedVenda.iFilialEmpresaFaturamento

    PedidoServico_Lock = SUCESSO

    Exit Function

Erro_PedidoServico_Lock:

    PedidoServico_Lock = gErr

    Select Case gErr

        Case 191681, 191682
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOSERVICO", gErr, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo)

        Case 191683

        Case 191684
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_PEDIDOSRV", gErr, objPedidoVenda.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191685)

    End Select

    Exit Function

End Function

Function ItensPSRV_Le_Lock(ByVal objPedVenda As ClassPedidoDeVenda) As Long
'Faz lock exclusive nos itens de PedVenda e lê os NumIntDoc desses itens

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim tItemPV As typeItemPedido
Dim tItemRomaneio As typeItemRomaneioGrade
Dim objItemRomaneio As ClassItemRomaneioGrade
Dim objItemPV As ClassItemPedido

On Error GoTo Erro_ItensPSRV_Le_Lock

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 193143

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 193144

    tItemPV.sDescricao = String(STRING_ITEM_PEDIDO_DESCRICAO, 0)
    tItemPV.sLote = String(STRING_ITEM_PEDIDO_LOTE, 0)
    tItemPV.sProduto = String(STRING_PRODUTO, 0)
    tItemPV.sUnidadeMed = String(STRING_UM_SIGLA, 0)

    'Lê os itens do pedido passado
    lErro = Comando_ExecutarPos(lComando, "SELECT NumIntDoc,Produto,Quantidade,PrecoUnitario,PrecoTotal,ClasseUM,UnidadeMed,ValorDesconto,DataEntrega,Descricao,Lote,Status,ValorAbatComissao,QuantCancelada,QuantReservada,QuantFaturada,QuantOP,QuantSC, Peca FROM ItensPedidoSRV WHERE FilialEmpresa = ? AND CodPedido = ?", 0, tItemPV.lNumIntDoc, tItemPV.sProduto, tItemPV.dQuantidade, tItemPV.dPrecoUnitario, tItemPV.dPrecoTotal, tItemPV.iClasseUM, tItemPV.sUnidadeMed, tItemPV.dValorDesconto, tItemPV.dtDataEntrega, tItemPV.sDescricao, tItemPV.sLote, tItemPV.iStatus, tItemPV.dValorAbatComissao, tItemPV.dQuantCancelada, tItemPV.dQuantReservada, tItemPV.dQuantFaturada, tItemPV.dQuantOP, tItemPV.dQuantSC, tItemPV.iPeca, objPedVenda.iFilialEmpresa, objPedVenda.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 193145

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then gError 193146
    
    If lErro = AD_SQL_SEM_DADOS Then gError 193147 'Não encontrou

    Do While lErro = AD_SQL_SUCESSO
        'faz o lock do item
        lErro = Comando_LockExclusive(lComando)
        If lErro <> AD_SQL_SUCESSO Then gError 193148

        With tItemPV
            Set tItemPV.colReservaItem = New ColReserva

            .iFilialEmpresa = objPedVenda.iFilialEmpresa
            .lCodPedido = objPedVenda.lCodigo

            If .iPeca = 1 Then

                'Adiciona o item a colecao de itens do pedido de Venda
                Set objItemPV = objPedVenda.colItensPedido.Add(.iFilialEmpresa, .lCodPedido, 0, .sProduto, .dQuantidade, .dPrecoUnitario, .dPrecoTotal, .dValorDesconto, .dtDataEntrega, "", .dValorAbatComissao, .dQuantCancelada, .dQuantReservada, .colReservaItem, "", .sUnidadeMed, .sLote, "", .iClasseUM, .dQuantFaturada, .dQuantAFaturar, .sDescricao, .iStatus, .iControleEstoque, .dQuantOP, .dQuantSC, 0, 0, 0, 0, 0, 0, .lNumIntDoc)
            
            Else
            
                'Adiciona o item a colecao de itens do pedido de Venda
                Set objItemPV = objPedVenda.colItensPedidoSRV.Add(.iFilialEmpresa, .lCodPedido, 0, .sProduto, .dQuantidade, .dPrecoUnitario, .dPrecoTotal, .dValorDesconto, .dtDataEntrega, "", .dValorAbatComissao, .dQuantCancelada, .dQuantReservada, .colReservaItem, "", .sUnidadeMed, .sLote, "", .iClasseUM, .dQuantFaturada, .dQuantAFaturar, .sDescricao, .iStatus, .iControleEstoque, .dQuantOP, .dQuantSC, 0, 0, 0, 0, 0, 0, .lNumIntDoc)
            
            End If
            
            objItemPV.iPeca = .iPeca
            
        End With
        
        'Busca o próximo item
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then gError 193149

    Loop

    'Fecha o comando
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    ItensPSRV_Le_Lock = SUCESSO

    Exit Function

Erro_ItensPSRV_Le_Lock:

    ItensPSRV_Le_Lock = gErr

    Select Case gErr

        Case 193143, 193144
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 193145, 193146, 193149
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDIDOSRV", gErr, objPedVenda.iFilialEmpresa, objPedVenda.lCodigo)

        Case 193147
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOSERVICO_SEM_ITENS", gErr, objPedVenda.iFilialEmpresa, objPedVenda.lCodigo)

        Case 193148
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_ITENSPEDIDOSRV", gErr, objPedVenda.iFilialEmpresa, objPedVenda.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 193150)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Public Function BloqueiosPedidoSRV_Atraso_Testa(ByVal objPedidoVenda As ClassPedidoDeVenda, iAtraso As Integer) As Long
'Verifica se tem bloqueio de atraso

Dim lErro As Long
Dim lComando As Long
Dim iTipoDeBloqueio As Integer

On Error GoTo Erro_BloqueiosPedidoSRV_Atraso_Testa

    'Abre os comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 191687

    'Busca Bloqueios de Crédito liberado
    lErro = Comando_Executar(lComando, "SELECT TipoDeBloqueio FROM BloqueiosPedidoSRV WHERE FilialEmpresa = ? AND PedidoSRV = ? AND TipoDeBloqueio = ? AND DataLib <> ?", iTipoDeBloqueio, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo, BLOQUEIO_DIAS_ATRASO, DATA_NULA)
    If lErro <> AD_SQL_SUCESSO Then gError 191688

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 191689

    iAtraso = 0

    'se o bloqueio de credito estiver aprovado ==> retona a flag indicativa
    If lErro = AD_SQL_SUCESSO Then iAtraso = BLOQUEIO_POR_ATRASO_LIBERADO

    'Fecha os comandos
    Call Comando_Fechar(lComando)

    BloqueiosPedidoSRV_Atraso_Testa = SUCESSO

    Exit Function

Erro_BloqueiosPedidoSRV_Atraso_Testa:

    BloqueiosPedidoSRV_Atraso_Testa = gErr

    Select Case gErr

        Case 191687
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 191688, 191689
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BLOQUEIOSPEDIDOSRV", gErr, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191690)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function BloqueiosPSRV_Aprovado_Testa(ByVal objPedidoVenda As ClassPedidoDeVenda, iStatus As Integer, ByVal iTipoBloqueio As Integer) As Long
'Verifica se existe o Bloqueio p/ o Tipo de Bloqueio recebido como parâmetro
'Função genérica

Dim lErro As Long
Dim lComando As Long
Dim iTipoBloqAux As Integer

On Error GoTo Erro_BloqueiosPSRV_Aprovado_Testa

    'Abre os comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 191697

    'Busca o tipo de Bloqueio recebido como parâmetro
    lErro = Comando_Executar(lComando, "SELECT TipoDeBloqueio FROM BloqueiosPedidoSRV WHERE FilialEmpresa = ? AND PedidoSRV = ? AND TipoDeBloqueio = ? AND DataLib <> ?", iTipoBloqAux, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo, iTipoBloqueio, DATA_NULA)
    If lErro <> AD_SQL_SUCESSO Then gError 191698

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 191699

    iStatus = 0

    'se o tipo de bloqueio estiver aprovado ==> retona a flag indicativa
    If lErro = AD_SQL_SUCESSO Then iStatus = BLOQUEIOPV_LIBERADO

    'Fecha os comandos
    Call Comando_Fechar(lComando)

    BloqueiosPSRV_Aprovado_Testa = SUCESSO

    Exit Function

Erro_BloqueiosPSRV_Aprovado_Testa:

    BloqueiosPSRV_Aprovado_Testa = gErr

    Select Case gErr

        Case 191697
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 191698, 191699
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BLOQUEIOSPEDIDOSRV", gErr, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191700)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function BloqueiosPSRV_Credito_Aprovado_Testa(ByVal objPedidoVenda As ClassPedidoDeVenda, iCreditoAprovado As Integer) As Long
'Verifica se tem bloqueio de crédito liberado

'Função alterada por Leo em 26/02/02. Agora só chama uma função mais genérica passando o tipo de bloqueio a ser verificado.
Dim lErro As Long

On Error GoTo Erro_BloqueiosPSRV_Credito_Aprovado_Testa

    'verifica se o pedido tem bloqueio por credito aprovado (função genérica, basta passar o tipo de Bloqueio)
    lErro = CF("BloqueiosPSRV_Aprovado_Testa", objPedidoVenda, iCreditoAprovado, BLOQUEIO_CREDITO)
    If lErro <> SUCESSO Then gError 191701

    BloqueiosPSRV_Credito_Aprovado_Testa = SUCESSO

    Exit Function

Erro_BloqueiosPSRV_Credito_Aprovado_Testa:

    BloqueiosPSRV_Credito_Aprovado_Testa = gErr

    Select Case gErr

        Case 94387
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191702)

    End Select

    Exit Function

End Function

Function BloqueioEstoquePSRV_Le(ByVal objBloqueioPV As ClassBloqueioPV) As Long
'Lê o Bloqueio de Estoque correspondente ao Pedido que não esteja liberado

Dim lErro As Long
Dim lComando As Long
Dim tBloqueioPV As typeBloqueioPV

On Error GoTo Erro_BloqueioEstoquePSRV_Le

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 191703

    tBloqueioPV.sCodUsuario = String(STRING_BLOQUEIOSPV_COD_USUARIO, 0)
    tBloqueioPV.sCodUsuarioLib = String(STRING_BLOQUEIOSPV_COD_USUARIO, 0)
    tBloqueioPV.sResponsavel = String(STRING_BLOQUEIOSPV_RESPONSAVEL, 0)
    tBloqueioPV.sResponsavelLib = String(STRING_BLOQUEIOSPV_RESPONSAVEL, 0)

    'Lê o bloqueio
    lErro = Comando_Executar(lComando, "SELECT Sequencial,TipoDeBloqueio,CodUsuario,Responsavel,Data,CodUsuarioLib,ResponsavelLib FROM BloqueiosPedidoSRV WHERE PedidoSRV = ? AND FilialEmpresa = ? AND (TipoDeBloqueio =? OR TipoDeBloqueio = ? OR TipoDeBloqueio = ?) AND DataLib =?", tBloqueioPV.iSequencial, tBloqueioPV.iTipoDeBloqueio, tBloqueioPV.sCodUsuario, tBloqueioPV.sResponsavel, tBloqueioPV.dtData, tBloqueioPV.sCodUsuarioLib, tBloqueioPV.sResponsavelLib, objBloqueioPV.lPedidoDeVendas, giFilialEmpresa, BLOQUEIO_TOTAL, BLOQUEIO_PARCIAL, BLOQUEIO_NAO_RESERVA, DATA_NULA)
    If lErro <> AD_SQL_SUCESSO Then gError 191704

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 191705
    
    If lErro <> AD_SQL_SUCESSO Then gError 191706

    'Armazena os dados do bloqueio
    objBloqueioPV.dtData = tBloqueioPV.dtData
    objBloqueioPV.dtDataLib = DATA_NULA
    objBloqueioPV.iFilialEmpresa = giFilialEmpresa
    objBloqueioPV.iSequencial = tBloqueioPV.iSequencial
    objBloqueioPV.iTipoDeBloqueio = tBloqueioPV.iTipoDeBloqueio
    objBloqueioPV.sCodUsuario = tBloqueioPV.sCodUsuario
    objBloqueioPV.sCodUsuarioLib = tBloqueioPV.sCodUsuarioLib
    objBloqueioPV.sResponsavel = tBloqueioPV.sResponsavel
    objBloqueioPV.sResponsavelLib = tBloqueioPV.sResponsavelLib

    'Fechar comando
    Call Comando_Fechar(lComando)

    BloqueioEstoquePSRV_Le = SUCESSO

    Exit Function

Erro_BloqueioEstoquePSRV_Le:

    BloqueioEstoquePSRV_Le = gErr

    Select Case gErr

        Case 191703
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 191704, 191705
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BLOQUEIOSPEDIDOSRV", gErr, giFilialEmpresa, objBloqueioPV.lPedidoDeVendas)

        Case 191706

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191707)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function PedidoServico_Le_Itens(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
'Lê no BD a coleção de ítens do Pedido de Venda e devolve em colItensPedido dentro de objPedidoVenda

Dim lErro As Long
Dim lComando As Long
Dim tItemPedido As typeItemPedido
Dim iItem As Integer
Dim objItemPV As ClassItemPedido

On Error GoTo Erro_PedidoServico_Le_Itens

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 191714

    tItemPedido.sDescricao = String(STRING_ITEM_PEDIDO_DESCRICAO, 0)
    tItemPedido.sLote = String(STRING_ITEM_PEDIDO_LOTE, 0)
    tItemPedido.sProduto = String(STRING_PRODUTO, 0)
    tItemPedido.sProdutoDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
    tItemPedido.sProdutoNomeReduzido = String(STRING_PRODUTO_NOME_REDUZIDO, 0)
    tItemPedido.sUMEstoque = String(STRING_PRODUTO_SIGLAUMESTOQUE, 0)
    tItemPedido.sUnidadeMed = String(STRING_UM_SIGLA, 0)

    lErro = Comando_Executar(lComando, "SELECT ItensPedidoSRV.NumIntDoc, ItensPedidoSRV.Produto, Quantidade, PrecoUnitario, PrecoTotal, ItensPedidoSRV.ClasseUM, UnidadeMed, ValorDesconto, DataEntrega, ItensPedidoSRV.Descricao, Lote, Status, ValorAbatComissao, QuantCancelada, QuantReservada, QuantFaturada, Produtos.Descricao, NomeReduzido, SiglaUMEstoque, ControleEstoque, ItensPedidoSRV.QuantSC, Peca FROM ItensPedidoSRV, Produtos WHERE ItensPedidoSRV.FilialEmpresa = ? AND CodPedido = ? AND ItensPedidoSRV.Produto = Produtos.Codigo ORDER BY ItensPedidoSRV.NumIntDoc", tItemPedido.lNumIntDoc, tItemPedido.sProduto, tItemPedido.dQuantidade, tItemPedido.dPrecoUnitario, tItemPedido.dPrecoTotal, tItemPedido.iClasseUM, tItemPedido.sUnidadeMed, tItemPedido.dValorDesconto, tItemPedido.dtDataEntrega, _
    tItemPedido.sDescricao, tItemPedido.sLote, tItemPedido.iStatus, tItemPedido.dValorAbatComissao, tItemPedido.dQuantCancelada, tItemPedido.dQuantReservada, tItemPedido.dQuantFaturada, tItemPedido.sProdutoDescricao, tItemPedido.sProdutoNomeReduzido, tItemPedido.sUMEstoque, tItemPedido.iControleEstoque, tItemPedido.dQuantSC, tItemPedido.iPeca, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 191715

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 191716
    
    iItem = 0

    Do While lErro <> AD_SQL_SEM_DADOS

        iItem = iItem + 1
        'Adiciona o ItemPedido na coleção de objPedidoVenda
        
        If tItemPedido.iPeca = 1 Then
        
            With tItemPedido
                Set objItemPV = objPedidoVenda.colItensPedido.Add(objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo, iItem, .sProduto, .dQuantidade, .dPrecoUnitario, .dPrecoTotal, .dValorDesconto, .dtDataEntrega, .sProdutoDescricao, .dValorAbatComissao, .dQuantCancelada, .dQuantReservada, .colReservaItem, .sProdutoNomeReduzido, .sUnidadeMed, .sLote, .sUMEstoque, .iClasseUM, .dQuantFaturada, .dQuantAFaturar, .sDescricao, .iStatus, .iControleEstoque, .dQuantOP, .dQuantSC, 0, 0, 0, 0, 0, 0, .lNumIntDoc)
                objItemPV.iPeca = tItemPedido.iPeca
            End With

        Else
        
            With tItemPedido
                Set objItemPV = objPedidoVenda.colItensPedidoSRV.Add(objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo, iItem, .sProduto, .dQuantidade, .dPrecoUnitario, .dPrecoTotal, .dValorDesconto, .dtDataEntrega, .sProdutoDescricao, .dValorAbatComissao, .dQuantCancelada, .dQuantReservada, .colReservaItem, .sProdutoNomeReduzido, .sUnidadeMed, .sLote, .sUMEstoque, .iClasseUM, .dQuantFaturada, .dQuantAFaturar, .sDescricao, .iStatus, .iControleEstoque, .dQuantOP, .dQuantSC, 0, 0, 0, 0, 0, 0, .lNumIntDoc)
                objItemPV.iPeca = tItemPedido.iPeca
            End With
        
        End If

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 191717
        
    Loop
    
    Call Comando_Fechar(lComando)

    PedidoServico_Le_Itens = SUCESSO

    Exit Function

Erro_PedidoServico_Le_Itens:

    PedidoServico_Le_Itens = gErr

    Select Case gErr

        Case 191714
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 191715, 191716, 191717
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDIDOSRV", gErr, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 191718)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ItensPSRV_ItensOP_Vinculados(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
'verifica se existe alguma Item de Ordem de Produção vincula do a algum Item do Pedido de Servico passado

Dim lErro As Long
Dim objItemPV As ClassItemPedido
Dim objItemOP As New ClassItemOP

On Error GoTo Erro_ItensPSRV_ItensOP_Vinculados

'    'Para casa item de Pedido de Venda
'    For Each objItemPV In objPedidoVenda.colItensPedido
'
'        'Lê o item de Ordem de Produção vinculado
'        lErro = CF("ItemOP_Le_ItemPSRV", objItemOP, objItemPV)
'        If lErro <> SUCESSO And lErro <> 191559 Then gError 191963
'
'        If lErro = SUCESSO Then gError 191964 'Se encontrou --> Erro
'
'    Next

    ItensPSRV_ItensOP_Vinculados = SUCESSO

    Exit Function

Erro_ItensPSRV_ItensOP_Vinculados:

    ItensPSRV_ItensOP_Vinculados = gErr

    Select Case gErr

        Case 191963

        Case 191964
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEM_OP_VINCULADO_ITEM_PSRV", gErr, objPedidoVenda.lCodigo, objItemPV.lNumIntDoc, objItemOP.lNumIntDoc)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191965)

    End Select

    Exit Function

End Function

Public Function PedidoServico_Busca_ReqCompra(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
'Verifica se existe requisição de compra, baixada ou não antes de excluir um pedido de venda

Dim lErro As Long
Dim lComando As Long
Dim lReqCompra As Long

On Error GoTo Erro_PedidoServico_Busca_ReqCompra

    'Abre Comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 193082
    
    'Faz leitura na pedido de venda
    lErro = Comando_Executar(lComando, "SELECT ReqCompra FROM PedidoServico WHERE FilialEmpresa=? AND Codigo=?", lReqCompra, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 193083
    
    'Posiciona o ponteiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 193084
    
    'Se o campo ReqCompra é <> 0
        '==> Existe Requisicao de Compras para este pedido - Erro
    If lReqCompra <> 0 Then gError 193085
    
    'Fecha comando
    Call Comando_Fechar(lComando)
    
    PedidoServico_Busca_ReqCompra = SUCESSO

    Exit Function
    
Erro_PedidoServico_Busca_ReqCompra:

    PedidoServico_Busca_ReqCompra = gErr
    
    Select Case gErr
    
        Case 193082
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 193083, 193084
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOSERVICO", gErr, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo)
           
        Case 193085
            Call Rotina_Erro(vbOKOnly, "ERRO_REQCOMPRA_EXISTENTE_SRV", gErr, lReqCompra)
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193086)
    
    End Select
    
    'Fecha Comandos - Saida por erro
    Call Comando_Fechar(lComando)

    Exit Function
    
End Function

Public Function PedidoServico_Busca_PedidoCompra(ByVal objPedidoVenda As ClassPedidoDeVenda, lCodigoPC As Long) As Long
'Verifica se existe pedido de compra, baixada ou não antes de excluir um pedido de venda

Dim lErro As Long
Dim lComando As Long

On Error GoTo Erro_PedidoServico_Busca_PedidoCompra

    'Abre Comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 193087
    
    'Faz leitura na pedido de venda
    lErro = Comando_Executar(lComando, "SELECT Codigo FROM PedidoCompraN WHERE FilialEmpresa=? AND CodigoPV=? AND TipoDocOrigem = ?", lCodigoPC, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo, PEDIDOCOMPRA_TIPODOCORIGEM_PSRV)
    If lErro <> AD_SQL_SUCESSO Then gError 193088
    
    'Posiciona o ponteiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 193089
    
    If lErro = AD_SQL_SEM_DADOS Then gError 193090
    
    'Fecha comando
    Call Comando_Fechar(lComando)
    
    PedidoServico_Busca_PedidoCompra = SUCESSO

    Exit Function
    
Erro_PedidoServico_Busca_PedidoCompra:

    PedidoServico_Busca_PedidoCompra = gErr
    
    Select Case gErr
    
        Case 193087
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 193088, 193089
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOCOMPRA", gErr)
           
        Case 193090
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193091)
    
    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function
    
End Function

Function ItemPSRVItemRC_Le_CodigoPV(ByVal objPedidoVenda As ClassPedidoDeVenda, ByVal colobjRC As Collection) As Long
'Lê os itens RC associados ao codigo do pedido de venda

Dim lErro As Long
Dim lComando As Long
Dim lCodigoRC As Long
Dim iFilialEmpresa As Integer
Dim objRC As ClassRequisicaoCompras

On Error GoTo Erro_ItemPSRVItemRC_Le_CodigoPV

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 193093


    'Lê Item RC a partir da Filial/Codigo do PSRV
    lErro = Comando_Executar(lComando, "SELECT RequisicaoCompraN.FilialEmpresa, RequisicaoCompraN.Codigo  FROM ItemPSRVItemRC, ItensPedidoDeVenda, ItensReqCompraN, RequisicaoCompraN WHERE ItemPV = ItensPedidoDeVenda.NumIntDoc AND ItemRC = ItensReqCompraN.NumINtDoc AND ItensReqCompraN.ReqCompra = RequisicaoCompraN.NumIntDoc AND ItensPedidoDeVenda.FilialEmpresa = ? AND CodPedido = ? ", iFilialEmpresa, lCodigoRC, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 193094

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then gError 193095

    Do While lErro = AD_SQL_SUCESSO
    
        Set objRC = New ClassRequisicaoCompras
    
        objRC.iFilialEmpresa = iFilialEmpresa
        objRC.lCodigo = lCodigoRC
    
        colobjRC.Add objRC
    
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then gError 193096
    
    Loop

    Call Comando_Fechar(lComando)

    ItemPSRVItemRC_Le_CodigoPV = SUCESSO

    Exit Function

Erro_ItemPSRVItemRC_Le_CodigoPV:

    ItemPSRVItemRC_Le_CodigoPV = gErr

    Select Case gErr

        Case 193093
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 193094, 193095, 193096
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMPSRVITEMRC", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 193097)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function PedidoServico_Le_Completo(ByVal objPedidoDeVenda As ClassPedidoDeVenda) As Long
'Lê o Pedido de Servico a partir do Código passado (incluindo tributacao, itens de pedido, parcelas, Comissoes, bloqueios)

Dim lErro As Long
Dim objOrcamentoVenda As New ClassOrcamentoVenda

On Error GoTo Erro_PedidoServico_Le_Completo

    lErro = CF("PedidoServico_Le", objPedidoDeVenda)
    If lErro <> SUCESSO And lErro <> 188828 Then gError 193244
    
    If lErro <> SUCESSO Then gError 193245

    lErro = CF("PedidoServico_Le_Tributacao", objPedidoDeVenda)
    If lErro <> SUCESSO And lErro <> 188837 Then gError 193246
    
    lErro = CF("ParcelasPedidoSRV_Le", objPedidoDeVenda)
    If lErro <> SUCESSO Then gError 193247
    
    'Le as Comissoes do Pedido
    lErro = CF("ComissoesPedidoSRV_Le", objPedidoDeVenda)
    If lErro <> SUCESSO Then gError 193248
    
'    'Le os Bloqueios do Pedido
'    lErro = CF("BloqueiosPedidoSRV_Le", objPedidoDeVenda)
'    If lErro <> SUCESSO Then gError 193249
    
    'Le os Itens e a Tributacao
    lErro = CF("PedidoSRV_Le_Itens_ComTributacao", objPedidoDeVenda)
    If lErro <> SUCESSO Then gError 193250
    
    objPedidoDeVenda.objInfoAdic.iTipoDoc = TIPODOC_INFOADIC_PSRV
    objPedidoDeVenda.objInfoAdic.lNumIntDoc = objPedidoDeVenda.lNumIntDoc
    lErro = CF("InfoAdicionais_Le", objPedidoDeVenda.objInfoAdic)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 193250

    PedidoServico_Le_Completo = SUCESSO

    Exit Function

Erro_PedidoServico_Le_Completo:

    PedidoServico_Le_Completo = gErr

    Select Case gErr

        Case 193244, 193246 To 193250
        
        Case 193245
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 193251)

    End Select

    Exit Function

End Function

Public Function OrcamentoSRV_Le_Completo(ByVal objOrcamentoVenda As ClassOrcamentoVenda) As Long
'Lê o Orcamento de Venda a partir do Código passado (incluindo tributacao, itens de orcamento, parcelas)

Dim lErro As Long

On Error GoTo Erro_OrcamentoSRV_Le_Completo

    lErro = CF("OrcamentoSRV_Le", objOrcamentoVenda)
    If lErro <> SUCESSO And lErro <> 186534 Then gError 193351
    
    If lErro <> SUCESSO Then gError 193343

    lErro = CF("OrcamentoSRV_Le_Tributacao", objOrcamentoVenda)
    If lErro <> SUCESSO And lErro <> 186623 Then gError 193344
    
    If lErro <> SUCESSO Then gError 193345

    'ler parcelas
    lErro = CF("ParcelasOrcSRV_Le", objOrcamentoVenda)
    If lErro <> SUCESSO And lErro <> 186617 Then gError 193346
    
    If lErro <> SUCESSO Then gError 193347

    'ler itens c/tributacao
    lErro = CF("OrcamentoSRV_Le_Itens_ComTributacao", objOrcamentoVenda)
    If lErro <> SUCESSO And lErro <> 186602 Then gError 193348
    
    If lErro <> SUCESSO Then gError 193349
    
    objOrcamentoVenda.objInfoAdic.iTipoDoc = TIPODOC_INFOADIC_OSRV
    objOrcamentoVenda.objInfoAdic.lNumIntDoc = objOrcamentoVenda.lNumIntDoc
    lErro = CF("InfoAdicionais_Le", objOrcamentoVenda.objInfoAdic)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 193348

    OrcamentoSRV_Le_Completo = SUCESSO

    Exit Function

Erro_OrcamentoSRV_Le_Completo:

    OrcamentoSRV_Le_Completo = gErr

    Select Case gErr

        Case 193343 To 193349, 193351

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 193350)

    End Select

    Exit Function

End Function

Function BloqueiosPSRV_Le(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
'Lê os Bloqueios associados ao objPedidoVenda

Dim lErro As Long
Dim lComando As Long
Dim tBloqueioPV As typeBloqueioPV

On Error GoTo Erro_BloqueiosPSRV_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 193471
    
    tBloqueioPV.sCodUsuario = String(STRING_BLOQUEIOSPV_COD_USUARIO, 0)
    tBloqueioPV.sCodUsuarioLib = String(STRING_BLOQUEIOSPV_COD_USUARIO, 0)
    tBloqueioPV.sResponsavel = String(STRING_BLOQUEIOSPV_RESPONSAVEL, 0)
    tBloqueioPV.sResponsavelLib = String(STRING_BLOQUEIOSPV_RESPONSAVEL, 0)
    tBloqueioPV.sObservacao = String(STRING_BLOQUEIOSPV_OBSERVACAO, 0)

    lErro = Comando_Executar(lComando, "SELECT Sequencial, TipoDeBloqueio, CodUsuario, Responsavel, Data, CodUsuarioLib, ResponsavelLib, DataLib, Observacao FROM BloqueiosPedidoSRV WHERE FilialEmpresa = ? AND PedidoSRV = ? ORDER BY Sequencial", tBloqueioPV.iSequencial, tBloqueioPV.iTipoDeBloqueio, tBloqueioPV.sCodUsuario, tBloqueioPV.sResponsavel, tBloqueioPV.dtData, tBloqueioPV.sCodUsuarioLib, tBloqueioPV.sResponsavelLib, tBloqueioPV.dtDataLib, tBloqueioPV.sObservacao, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 193472
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 193473
       
    Do While lErro <> AD_SQL_SEM_DADOS
    
        With tBloqueioPV
            objPedidoVenda.colBloqueiosPV.Add objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo, .iSequencial, .iTipoDeBloqueio, .sCodUsuario, .sResponsavel, .dtData, .sCodUsuarioLib, .sResponsavelLib, .dtDataLib, .sObservacao
        End With
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 193474
        
    Loop
    
    Call Comando_Fechar(lComando)
    
    BloqueiosPSRV_Le = SUCESSO
    
    Exit Function

Erro_BloqueiosPSRV_Le:

    BloqueiosPSRV_Le = gErr
    
    Select Case gErr
    
        Case 193471
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 193472, 193473, 193474
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BLOQUEIOSPEDIDOSRV", gErr, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo)
               
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 193475)
            
    End Select
    
    Call Comando_Fechar(lComando)
        
    Exit Function
    
End Function

Function OS_Le_ComItens(ByVal objOS As ClassOS) As Long
'carrega do bd a ordem de servico inclusive os seus itens

Dim lErro As Long
Dim iIndice As Integer
Dim tItemOS As typeItemOS
Dim objItemOS As ClassItemOS
Dim alComando(0 To 2) As Long
Dim tOS As typeOS

On Error GoTo Erro_OS_Le_ComItens

    For iIndice = LBound(alComando) To UBound(alComando)

        'Abertura comando
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 195057
        
    Next

    tOS.sProduto = String(STRING_PRODUTO, 0)
    tOS.sLote = String(STRING_LOTE_RASTREAMENTO, 0)

    'Leitura da tabela de Ordens de Servico
    lErro = Comando_Executar(alComando(0), "SELECT Tipo, DataEmissao, NumItens, NumItensBaixados, CodSolSRV, Produto, Lote, FilialOP, CodPedSRV, FilialPedSRV FROM OS WHERE FilialEmpresa = ? AND Codigo = ?", tOS.lTipo, tOS.dtDataEmissao, tOS.iNumItens, tOS.iNumItensBaixados, tOS.lCodSolSRV, tOS.sProduto, tOS.sLote, tOS.iFilialOP, tOS.lCodPedSRV, tOS.iFilialPedSRV, objOS.iFilialEmpresa, objOS.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 195058

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 195059

    If lErro <> AD_SQL_SUCESSO Then gError 195060

    objOS.dtDataEmissao = tOS.dtDataEmissao
    objOS.iNumItens = tOS.iNumItens
    objOS.iNumItensBaixados = tOS.iNumItensBaixados
    objOS.lCodSolSRV = tOS.lCodSolSRV
    objOS.sProduto = tOS.sProduto
    objOS.sLote = tOS.sLote
    objOS.iFilialOP = tOS.iFilialOP
    objOS.lCodPedSRV = tOS.lCodPedSRV
    objOS.iFilialPedSRV = tOS.iFilialPedSRV
    objOS.lTipo = tOS.lTipo

    tItemOS.sServico = String(STRING_PRODUTO, 0)
    tItemOS.sCcl = String(STRING_CCL, 0)
    tItemOS.sSiglaUM = String(STRING_UM_SIGLA, 0)
    tItemOS.sVersao = String(STRING_KIT_VERSAO, 0)
    tItemOS.sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
    tItemOS.sObservacao = String(STRING_ITEMOS_OBS, 0)

    'Leitura da tabela de Itens de Ordem de Servico
    lErro = Comando_Executar(alComando(1), "SELECT ItemOS.NumIntDoc, ItemOS.Item, ItemOS.Servico, Descricao,  " & _
    "ItemOS.CCL, ItemOS.SiglaUM, ItemOS.Quantidade, ItemOS.DataInicio, " & _
    "ItemOS.DataFim,  ItemOS.Prioridade, ItemOS.StatusItem, ItemOS.Versao, ItemOS.Observacao FROM ItemOS, Produtos WHERE ItemOS.FilialEmpresa = ? AND ItemOS.Codigo = ? AND ItemOS.Servico = Produtos.Codigo ORDER BY Item", _
    tItemOS.lNumIntDoc, tItemOS.iItem, tItemOS.sServico, tItemOS.sDescricao, _
    tItemOS.sCcl, tItemOS.sSiglaUM, tItemOS.dQuantidade, tItemOS.dtDataInicio, _
    tItemOS.dtDataFim, tItemOS.iPrioridade, tItemOS.iStatusItem, tItemOS.sVersao, tItemOS.sObservacao, objOS.iFilialEmpresa, objOS.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 195061

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 195062

    'Enquanto existirem itens para uma determinada ordem de produção
    Do While lErro <> AD_SQL_SEM_DADOS

        Set objItemOS = New ClassItemOS

        objItemOS.lNumIntDoc = tItemOS.lNumIntDoc
        objItemOS.iItem = tItemOS.iItem
        objItemOS.sServico = tItemOS.sServico
        objItemOS.sDescricao = tItemOS.sDescricao
        objItemOS.sCcl = tItemOS.sCcl
        objItemOS.sSiglaUM = tItemOS.sSiglaUM
        objItemOS.dQuantidade = tItemOS.dQuantidade
        objItemOS.dtDataInicio = tItemOS.dtDataInicio
        objItemOS.dtDataFim = tItemOS.dtDataFim
        objItemOS.iPrioridade = tItemOS.iPrioridade
        objItemOS.iStatusItem = tItemOS.iStatusItem
        objItemOS.iFilialEmpresa = objOS.iFilialEmpresa
        objItemOS.sCodigo = objOS.sCodigo
        objItemOS.sVersao = tItemOS.sVersao
        objItemOS.sObservacao = tItemOS.sObservacao

        lErro = ItemOSOperacoes_Le(objItemOS)
        If lErro <> SUCESSO And lErro <> 195070 Then gError 195063
    
        objOS.colItens.Add objItemOS

        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 195065

    Loop

   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    OS_Le_ComItens = SUCESSO

    Exit Function

Erro_OS_Le_ComItens:

    OS_Le_ComItens = gErr

    Select Case gErr

        Case 195057
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 195058, 195059
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_OS", gErr)

        Case 195060, 195063

        Case 195061, 195062, 195065
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMOS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195066)

    End Select

   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Function ItemOSOperacoes_Le(ByVal objItemOS As ClassItemOS) As Long

Dim lErro As Long
Dim lComando As Long
Dim sComandoSQL As String
Dim objItemOSOperacoes As ClassItemOSOperacoes
Dim tItemOSOperacoes As typeItemOSOperacoes

On Error GoTo Erro_ItemOSOperacoes_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 195067

    'Alocação de espaço no buffer
    tItemOSOperacoes.sServico = String(STRING_PRODUTO, 0)
    tItemOSOperacoes.sVersao = String(STRING_KIT_VERSAO, 0)
    tItemOSOperacoes.sObservacao = String(STRING_OBSERVACAO_OBSERVACAO, 0)
    
    sComandoSQL = "SELECT NumIntDoc, NumIntDocItemOS, Seq, NumIntDocCompet, NumIntDocCT, "
    sComandoSQL = sComandoSQL & "Observacao, SeqPai, SeqArvore, Nivel, "
    sComandoSQL = sComandoSQL & "SeqRoteiro, SeqRoteiroPai, NivelRoteiro, Servico, Versao, "
    sComandoSQL = sComandoSQL & "IgnoraTaxaProducao, ConsideraCarga, Origem, NumIntDocOperOrigem, NumMaxMaqPorOper, NumRepeticoes "
    sComandoSQL = sComandoSQL & "FROM ItemOSOperacoes WHERE NumIntDocItemOS = ? Order by Seq"

    With tItemOSOperacoes
        
        'Le a tabela Operacoes
        lErro = Comando_Executar(lComando, sComandoSQL, .lNumIntDoc, .lNumIntDocItemOS, _
                                .iSeq, .lNumIntDocCompet, .lNumIntDocCT, .sObservacao, _
                                .iSeqPai, .iSeqArvore, .iNivel, .iSeqRoteiro, _
                                .iSeqRoteiroPai, .iNivelRoteiro, .sServico, .sVersao, _
                                .iIgnoraTaxaProducao, .iConsideraCarga, .iOrigem, _
                                .lNumIntDocOPerOrigem, .iNumMaxMaqPorOper, .iNumRepeticoes, objItemOS.lNumIntDoc)
        
    End With
        
    If lErro <> AD_SQL_SUCESSO Then gError 195068

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 195069

    'Operacoes não cadastrada
    If lErro = AD_SQL_SEM_DADOS Then gError 195070

    'Passa todos para a collection
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objItemOSOperacoes = New ClassItemOSOperacoes
        
        objItemOSOperacoes.lNumIntDoc = tItemOSOperacoes.lNumIntDoc
        objItemOSOperacoes.lNumIntDocItemOS = tItemOSOperacoes.lNumIntDocItemOS
        objItemOSOperacoes.iSeq = tItemOSOperacoes.iSeq
        objItemOSOperacoes.lNumIntDocCompet = tItemOSOperacoes.lNumIntDocCompet
        objItemOSOperacoes.lNumIntDocCT = tItemOSOperacoes.lNumIntDocCT
        objItemOSOperacoes.sObservacao = tItemOSOperacoes.sObservacao
        objItemOSOperacoes.iSeqPai = tItemOSOperacoes.iSeqPai
        objItemOSOperacoes.iSeqArvore = tItemOSOperacoes.iSeqArvore
        objItemOSOperacoes.iNivel = tItemOSOperacoes.iNivel
        objItemOSOperacoes.iSeqRoteiro = tItemOSOperacoes.iSeqRoteiro
        objItemOSOperacoes.iSeqRoteiroPai = tItemOSOperacoes.iSeqRoteiroPai
        objItemOSOperacoes.iNivelRoteiro = tItemOSOperacoes.iNivelRoteiro
        objItemOSOperacoes.sServico = tItemOSOperacoes.sServico
        objItemOSOperacoes.sVersao = tItemOSOperacoes.sVersao
        objItemOSOperacoes.iIgnoraTaxaProducao = tItemOSOperacoes.iIgnoraTaxaProducao
        objItemOSOperacoes.iConsideraCarga = tItemOSOperacoes.iConsideraCarga
        objItemOSOperacoes.iOrigem = tItemOSOperacoes.iOrigem
        objItemOSOperacoes.lNumIntDocOPerOrigem = tItemOSOperacoes.lNumIntDocOPerOrigem
        objItemOSOperacoes.iNumMaxMaqPorOper = tItemOSOperacoes.iNumMaxMaqPorOper
        objItemOSOperacoes.iNumRepeticoes = tItemOSOperacoes.iNumRepeticoes
        objItemOSOperacoes.dQuantidade = objItemOS.dQuantidade
        objItemOSOperacoes.sUMedida = objItemOS.sSiglaUM
        
       
        'Lê as OperacaoInsumos associadas a Operação que está sendo lida
        lErro = ItemOSOperacoesPecas_Le(objItemOSOperacoes)
        If lErro <> SUCESSO Then gError 195071
        
        lErro = ItemOSOperacoesMaquinas_Le(objItemOSOperacoes)
        If lErro <> SUCESSO Then gError 195072
        
        lErro = ItemOSOperacoesMO_Le(objItemOSOperacoes)
        If lErro <> SUCESSO Then gError 195080
        
        objItemOS.colItemOSOperacoes.Add objItemOSOperacoes
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 195073
        
    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    ItemOSOperacoes_Le = SUCESSO
    
    Exit Function

Erro_ItemOSOperacoes_Le:
    
    ItemOSOperacoes_Le = gErr
    
    Select Case gErr
    
        Case 195067
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 195068, 195069, 195073
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMOSOPERACOES", gErr)

        Case 195070, 195071, 195072, 195080
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195074)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Public Function ItemOSOperacoesPecas_Le(ByVal objItemOSOperacoes As ClassItemOSOperacoes) As Long

Dim lErro As Long
Dim lComando As Long
Dim sComandoSQL As String
Dim objItemOSOperacoesPecas As ClassItemOSOperacoesPecas
Dim tItemOSOperacoesPecas As typeItemOSOperacoesPecas

On Error GoTo Erro_ItemOSOperacoesPecas_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 193807

    'Alocação de espaço no buffer
    tItemOSOperacoesPecas.sProduto = String(STRING_PRODUTO, 0)
    tItemOSOperacoesPecas.sUMProduto = String(STRING_UM_SIGLA, 0)
    tItemOSOperacoesPecas.sVersaoKitComp = String(STRING_KIT_VERSAO, 0)
    
    sComandoSQL = "SELECT NumIntDocOper, Produto, Quantidade, UMProduto, VersaoKitComp, "
    sComandoSQL = sComandoSQL & "Composicao, PercentualPerda, CustoStandard "
    sComandoSQL = sComandoSQL & "From ItemOSOperacoesPecas WHERE NumIntDocOper = ?"

    'Le a tabela Operacoes
    lErro = Comando_Executar(lComando, sComandoSQL, tItemOSOperacoesPecas.lNumIntDocOper, _
                            tItemOSOperacoesPecas.sProduto, tItemOSOperacoesPecas.dQuantidade, _
                            tItemOSOperacoesPecas.sUMProduto, tItemOSOperacoesPecas.sVersaoKitComp, _
                            tItemOSOperacoesPecas.iComposicao, tItemOSOperacoesPecas.dPercentualPerda, _
                            tItemOSOperacoesPecas.dCustoStandard, objItemOSOperacoes.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 193808

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 193809

    'Passa todos para a collection
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objItemOSOperacoesPecas = New ClassItemOSOperacoesPecas
        
        objItemOSOperacoesPecas.lNumIntDocOper = tItemOSOperacoesPecas.lNumIntDocOper
        objItemOSOperacoesPecas.sProduto = tItemOSOperacoesPecas.sProduto
        objItemOSOperacoesPecas.dQuantidade = tItemOSOperacoesPecas.dQuantidade
        objItemOSOperacoesPecas.sUMProduto = tItemOSOperacoesPecas.sUMProduto
        objItemOSOperacoesPecas.sVersaoKitComp = tItemOSOperacoesPecas.sVersaoKitComp
        objItemOSOperacoesPecas.iComposicao = tItemOSOperacoesPecas.iComposicao
        objItemOSOperacoesPecas.dPercentualPerda = tItemOSOperacoesPecas.dPercentualPerda
        objItemOSOperacoesPecas.dCustoStandard = tItemOSOperacoesPecas.dCustoStandard
                            
        objItemOSOperacoes.colPecas.Add objItemOSOperacoesPecas
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 193811
        
    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    ItemOSOperacoesPecas_Le = SUCESSO

    Exit Function

Erro_ItemOSOperacoesPecas_Le:

    ItemOSOperacoesPecas_Le = gErr

    Select Case gErr

        Case 193807
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 193808, 193809, 193811
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMOSOPERACOESPECAS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193812)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function ItemOSOperacoesMaquinas_Le(ByVal objItemOSOperacoes As ClassItemOSOperacoes) As Long

Dim lErro As Long
Dim lComando As Long
Dim tItemOSOperacoesMaquinas As typeItemOSOperacoesMaquinas
Dim objItemOSOperacoesMaquinas As ClassItemOSOperacoesMaquinas

On Error GoTo Erro_ItemOSOperacoesMaquinas_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 193813

    'Alocação de espaço no buffer

    'Le a tabelaItemOPOperacoesMaquinas
    lErro = Comando_Executar(lComando, "SELECT NumIntDocOper, NumIntDocMaq, " & _
                "Horas, Quantidade FROM ItemOSOperacoesMaquinas WHERE NumIntDocOper = ? ", _
                tItemOSOperacoesMaquinas.lNumIntDocOper, tItemOSOperacoesMaquinas.lNumIntDocMaq, _
                tItemOSOperacoesMaquinas.dHoras, tItemOSOperacoesMaquinas.iQuantidade, _
                objItemOSOperacoes.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 193814

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 193815

    'Sem Dados
    Do While lErro <> AD_SQL_SEM_DADOS

        Set objItemOSOperacoesMaquinas = New ClassItemOSOperacoesMaquinas

        objItemOSOperacoesMaquinas.lNumIntDocOper = tItemOSOperacoesMaquinas.lNumIntDocOper
        objItemOSOperacoesMaquinas.lNumIntDocMaq = tItemOSOperacoesMaquinas.lNumIntDocMaq
        objItemOSOperacoesMaquinas.dHoras = tItemOSOperacoesMaquinas.dHoras
        objItemOSOperacoesMaquinas.iQuantidade = tItemOSOperacoesMaquinas.iQuantidade
        
        objItemOSOperacoes.colMaquinas.Add objItemOSOperacoesMaquinas
        
        'Busca Primeiro
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 193817
        
    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    ItemOSOperacoesMaquinas_Le = SUCESSO

    Exit Function

Erro_ItemOSOperacoesMaquinas_Le:

    ItemOSOperacoesMaquinas_Le = gErr

    Select Case gErr

        Case 193813
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 193814, 193815, 193817
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMOSOPERACOESMAQUINAS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193818)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function MO_Le(ByVal objMO As ClassMO) As Long
'le a mao de obra a partir da chave(lCodigo) fornecida em objMO. Devolve os dados em objMO

Dim lComando As Long
Dim lComando2 As Long
Dim lErro As Long
Dim tMO As typeMO
Dim lCodigoCliente As Long

On Error GoTo Erro_MO_Le

    tMO.sNome = String(STRING_MO_NOME, 0)
    tMO.sNomeReduzido = String(STRING_MO_NOMERED, 0)
    tMO.sObservacao = String(STRING_MO_OBS, 0)
    
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 193814

    lErro = Comando_Executar(lComando, "SELECT Codigo, Nome, NomeReduzido, Observacao, Ativo, Tipo FROM MaoDeObra WHERE Codigo = ?", _
    tMO.lCodigo, tMO.sNome, tMO.sNomeReduzido, tMO.sObservacao, tMO.iAtivo, tMO.iTipo, objMO.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 193815

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 193816

    If lErro = AD_SQL_SEM_DADOS Then gError 193817

    objMO.lCodigo = tMO.lCodigo
    objMO.sNome = tMO.sNome
    objMO.sNomeReduzido = tMO.sNomeReduzido
    objMO.sObservacao = tMO.sObservacao
    objMO.iAtivo = tMO.iAtivo
    objMO.iTipo = tMO.iTipo
    
    lErro = Comando_Fechar(lComando)

    MO_Le = SUCESSO

    Exit Function

Erro_MO_Le:

    MO_Le = gErr

    Select Case gErr

        Case 193814
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 193815, 193816
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MAODEOBRA", gErr)

        Case 193817

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193818)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function MO_Le_Tipo(ByVal iTipo As Integer, ByVal colMO As Collection) As Long
'le a mao de obra a partir da chave(lCodigo) fornecida em objMO. Devolve os dados em objMO

Dim lComando As Long
Dim lComando2 As Long
Dim lErro As Long
Dim tMO As typeMO
Dim lCodigoCliente As Long
Dim objMO As ClassMO

On Error GoTo Erro_MO_Le_Tipo

    tMO.sNome = String(STRING_MO_NOME, 0)
    tMO.sNomeReduzido = String(STRING_MO_NOMERED, 0)
    tMO.sObservacao = String(STRING_MO_OBS, 0)
    
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 193814

    lErro = Comando_Executar(lComando, "SELECT Codigo, Nome, NomeReduzido, Observacao, Ativo, Tipo FROM MaoDeObra WHERE Tipo = ? ORDER BY Codigo", _
    tMO.lCodigo, tMO.sNome, tMO.sNomeReduzido, tMO.sObservacao, tMO.iAtivo, tMO.iTipo, iTipo)
    If lErro <> AD_SQL_SUCESSO Then gError 193815

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 193816

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objMO = New ClassMO
    
        objMO.lCodigo = tMO.lCodigo
        objMO.sNome = tMO.sNome
        objMO.sNomeReduzido = tMO.sNomeReduzido
        objMO.sObservacao = tMO.sObservacao
        objMO.iAtivo = tMO.iAtivo
        objMO.iTipo = tMO.iTipo
        
        colMO.Add objMO
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 193816
        
    Loop
    
    Call Comando_Fechar(lComando)

    MO_Le_Tipo = SUCESSO

    Exit Function

Erro_MO_Le_Tipo:

    MO_Le_Tipo = gErr

    Select Case gErr

        Case 193814
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 193815, 193816
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MAODEOBRA", gErr)

        Case 193817

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193818)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function MO_Le_NomeRed(ByVal objMO As ClassMO) As Long
'le a mao de obra a partir da chave(nome reduzido) fornecida em objMO. Devolve os dados em objMO

Dim lComando As Long
Dim lComando2 As Long
Dim lErro As Long
Dim tMO As typeMO
Dim lCodigoCliente As Long

On Error GoTo Erro_MO_Le_NomeRed

    tMO.sNome = String(STRING_MO_NOME, 0)
    tMO.sNomeReduzido = String(STRING_MO_NOMERED, 0)
    tMO.sObservacao = String(STRING_MO_OBS, 0)
    
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 193852

    lErro = Comando_Executar(lComando, "SELECT Codigo, Nome, NomeReduzido, Observacao, Ativo, Tipo FROM MaoDeObra WHERE NomeReduzido = ?", _
    tMO.lCodigo, tMO.sNome, tMO.sNomeReduzido, tMO.sObservacao, tMO.iAtivo, tMO.iTipo, objMO.sNomeReduzido)
    If lErro <> AD_SQL_SUCESSO Then gError 193853

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 193854

    If lErro = AD_SQL_SEM_DADOS Then gError 193855

    objMO.lCodigo = tMO.lCodigo
    objMO.sNome = tMO.sNome
    objMO.sNomeReduzido = tMO.sNomeReduzido
    objMO.sObservacao = tMO.sObservacao
    objMO.iAtivo = tMO.iAtivo
    objMO.iTipo = tMO.iTipo
    
    lErro = Comando_Fechar(lComando)

    MO_Le_NomeRed = SUCESSO

    Exit Function

Erro_MO_Le_NomeRed:

    MO_Le_NomeRed = gErr

    Select Case gErr

        Case 193852
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 193853, 193854
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MAODEOBRA", gErr)

        Case 193855

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193856)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function ItemOSOperacoesMO_Le(ByVal objItemOSOperacoes As ClassItemOSOperacoes) As Long

Dim lErro As Long
Dim lComando As Long
Dim sComandoSQL As String
Dim objItemOSOperacoesMO As ClassItemOSOperacoesMO
Dim tItemOSOperacoesMO As typeItemOSOperacoesMO

On Error GoTo Erro_ItemOSOperacoesMO_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 195075

    sComandoSQL = "SELECT NumIntDocOper, CodigoMO, Horas From ItemOSOperacoesMO WHERE NumIntDocOper = ?"

    'Le a tabela Operacoes
    lErro = Comando_Executar(lComando, sComandoSQL, tItemOSOperacoesMO.lNumIntDocOper, _
                            tItemOSOperacoesMO.lCodigoMO, tItemOSOperacoesMO.dHoras, objItemOSOperacoes.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 195076

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 195077

    'Passa todos para a collection
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objItemOSOperacoesMO = New ClassItemOSOperacoesMO
        
        objItemOSOperacoesMO.lNumIntDocOper = tItemOSOperacoesMO.lNumIntDocOper
        objItemOSOperacoesMO.lCodigoMO = tItemOSOperacoesMO.lCodigoMO
        objItemOSOperacoesMO.dHoras = tItemOSOperacoesMO.dHoras
                            
        objItemOSOperacoes.colMO.Add objItemOSOperacoesMO
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 195078
        
    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    ItemOSOperacoesMO_Le = SUCESSO

    Exit Function

Erro_ItemOSOperacoesMO_Le:

    ItemOSOperacoesMO_Le = gErr

    Select Case gErr

        Case 195075
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 195076 To 195078
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMOSOPERACOESMO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195079)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function OS_Le(ByVal objOS As ClassOS) As Long
'Lê a Ordem de Servico pelo Código e pela FilialEmpresa

Dim lErro As Long
Dim lComando As Long
Dim tOS As typeOS

On Error GoTo Erro_OS_Le

    If objOS.iFilialEmpresa = 0 Then objOS.iFilialEmpresa = giFilialEmpresa
    
    'Abertura de comando
    lComando = Comando_Abrir
    If lComando = 0 Then gError 195099

    With tOS

        .sProduto = String(STRING_PRODUTO, 0)
        .sLote = String(STRING_LOTE_RASTREAMENTO, 0)

        lErro = Comando_Executar(lComando, "SELECT Tipo, DataEmissao, NumItens, NumItensBaixados, Status, CodSolSRV, Produto, Lote, FilialOP, CodPedSRV, FilialPedSRV FROM OS WHERE FilialEmpresa = ? AND Codigo = ?", .lTipo, .dtDataEmissao, .iNumItens, .iNumItensBaixados, .iStatus, .lCodSolSRV, .sProduto, .sLote, .iFilialOP, .lCodPedSRV, .iFilialPedSRV, objOS.iFilialEmpresa, objOS.sCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 195100

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 195101

        If lErro = AD_SQL_SEM_DADOS Then gError 195102
        
        objOS.iNumItens = .iNumItens
        objOS.dtDataEmissao = .dtDataEmissao
        objOS.iNumItensBaixados = .iNumItensBaixados
        objOS.iStatus = .iStatus
        objOS.lCodSolSRV = tOS.lCodSolSRV
        objOS.sProduto = tOS.sProduto
        objOS.sLote = tOS.sLote
        objOS.iFilialOP = tOS.iFilialOP
        objOS.lCodPedSRV = tOS.lCodPedSRV
        objOS.iFilialPedSRV = tOS.iFilialPedSRV
        objOS.lTipo = tOS.lTipo
        
    End With

    Call Comando_Fechar(lComando)

    OS_Le = SUCESSO

    Exit Function

Erro_OS_Le:

    OS_Le = gErr

    Select Case gErr

        Case 195099
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 195100, 195101
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_OS", gErr)

        Case 195102

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195103)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function ItensDeContratoSrv_Le(objItensContSrv As ClassItensDeContratoSrv) As Long

Dim lComando As Long
Dim lErro As Long
Dim tItensContSrv As typeItensDeContratoSrv

On Error GoTo Erro_ItensDeContratoSrv_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 195581
        
    tItensContSrv.sProduto = String(STRING_PRODUTO, 0)
    tItensContSrv.sLote = String(STRING_LOTE_RASTREAMENTO, 0)
    tItensContSrv.sCodigoContrato = String(STRING_CONTRATO_CODIGO, 0)
        
    lErro = Comando_Executar(lComando, "SELECT ItensDeContratoSrv.NumIntDoc, ItensDeContratoSrv.Codigo, ItensDeContratoSrv.NumIntItemContrato, Contratos.Codigo, ItensDeContratoSRV.FilialEmpresa, Contratos.DataIniContrato, Contratos.DataFimContrato, ItensDeContratoSrv.Quantidade, ItensDeContratoSrv.Lote, ItensDeContratoSrv.FilialOP, ItensDeContratoSrv.TipoGarantia, ItensDeContratoSrv.GarantiaTotal, ItensDeContrato.Produto FROM Contratos, ItensDeContrato, ItensDeContratoSrv WHERE " & _
    "Contratos.NumIntDoc = ItensDeContrato.NumIntContrato AND ItensDeContrato.NumIntDoc = ItensDeContratoSrv.NumIntItemContrato AND ItensDeContratoSRV.FilialEmpresa = ? AND ItensDeContratoSRV.Codigo = ? ", _
    tItensContSrv.lNumIntDoc, tItensContSrv.lCodigo, tItensContSrv.lNumIntItemContrato, tItensContSrv.sCodigoContrato, tItensContSrv.iFilialEmpresa, tItensContSrv.dtDataContratoIni, tItensContSrv.dtDataContratoFim, tItensContSrv.dQuantidade, tItensContSrv.sLote, tItensContSrv.iFilialOP, tItensContSrv.lTipoGarantia, tItensContSrv.iGarantiaTotal, tItensContSrv.sProduto, objItensContSrv.iFilialEmpresa, objItensContSrv.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 195582
        
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 195583
        
    If lErro <> AD_SQL_SUCESSO Then gError 195584
        
    objItensContSrv.lNumIntDoc = tItensContSrv.lNumIntDoc
    objItensContSrv.lCodigo = tItensContSrv.lCodigo
    objItensContSrv.lNumIntItemContrato = tItensContSrv.lNumIntItemContrato
    objItensContSrv.sCodigoContrato = tItensContSrv.sCodigoContrato
    objItensContSrv.iFilialEmpresa = tItensContSrv.iFilialEmpresa
    objItensContSrv.dtDataContratoIni = tItensContSrv.dtDataContratoIni
    objItensContSrv.dtDataContratoFim = tItensContSrv.dtDataContratoFim
    objItensContSrv.dQuantidade = tItensContSrv.dQuantidade
    objItensContSrv.sLote = tItensContSrv.sLote
    objItensContSrv.iFilialOP = tItensContSrv.iFilialOP
    objItensContSrv.lTipoGarantia = tItensContSrv.lTipoGarantia
    objItensContSrv.iGarantiaTotal = tItensContSrv.iGarantiaTotal
    objItensContSrv.sProduto = tItensContSrv.sProduto

    lErro = CF("ItensDeContratoSrvProd_Le", objItensContSrv)
    If lErro <> SUCESSO Then gError 195608

    lErro = CF("ItensDeContratoSrvNumSerie_Le", objItensContSrv)
    If lErro <> SUCESSO Then gError 195609

    Call Comando_Fechar(lComando)
    
    ItensDeContratoSrv_Le = SUCESSO
    
    Exit Function

Erro_ItensDeContratoSrv_Le:
    
    ItensDeContratoSrv_Le = gErr
        
    Select Case gErr
    
        Case 195581
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 195582, 195583
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMCONTRATOSRV", gErr, objItensContSrv.iFilialEmpresa, objItensContSrv.lCodigo)
        
        Case 195584, 195608, 195609
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195585)
            
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function ItensSolicSRV_Le_Contrato(objSolicSRV As ClassSolicSRV, ByVal iFilialEmpresa As Integer, ByVal sContrato As String, ByVal sProduto As String) As Long
'pesquisa alguma solicitacao envolvendo o contrato e produto em questao

Dim lComando As Long
Dim lErro As Long
Dim tItensSolicSRV As typeItensSolicSRV
Dim objItensSolicSrv As ClassItensSolicSRV

On Error GoTo Erro_ItensSolicSRV_Le_Contrato

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 195593
        
    tItensSolicSRV.sProduto = String(STRING_PRODUTO, 0)
    tItensSolicSRV.sProdutoDesc = String(STRING_PRODUTO_DESCRICAO, 0)
    tItensSolicSRV.sServico = String(STRING_PRODUTO, 0)
    tItensSolicSRV.sServicoDesc = String(STRING_PRODUTO_DESCRICAO, 0)
    tItensSolicSRV.sLote = String(STRING_LOTE_RASTREAMENTO, 0)
    tItensSolicSRV.sSolicitacao = String(STRING_SOLICITACAO, 0)
    tItensSolicSRV.sContrato = String(STRING_CONTRATO_CODIGO, 0)
    tItensSolicSRV.sUM = String(STRING_UM_SIGLA, 0)
    tItensSolicSRV.sReparo = String(STRING_MAXIMO, 0)
        
    lErro = Comando_Executar(lComando, "SELECT ItensSolicSRV.NumIntDoc, Produto, Produtos.Descricao, DataVenda, Servico, Servicos.Descricao, UM, Quantidade, Lote, FilialOP, Solicitacao, Garantia, Contrato, StatusItem, Reparo, DataBaixa FROM SolicitacaoSRV, ItensSolicSRV, Produtos, Produtos As Servicos WHERE SolicitacaoSRV.FilialEmpresa = ? AND Contrato = ? AND Produto = ? AND SolicitacaoSRV.NumIntDoc = ItensSolicSRV.NumIntSolicSRV AND ItensSolicSRV.Produto = Produtos.Codigo AND ItensSolicSRV.Servico = Servicos.Codigo", _
    tItensSolicSRV.lNumIntDoc, tItensSolicSRV.sProduto, tItensSolicSRV.sProdutoDesc, tItensSolicSRV.dtDataVenda, tItensSolicSRV.sServico, tItensSolicSRV.sServicoDesc, tItensSolicSRV.sUM, tItensSolicSRV.dQuantidade, tItensSolicSRV.sLote, tItensSolicSRV.iFilialOP, tItensSolicSRV.sSolicitacao, tItensSolicSRV.lGarantia, tItensSolicSRV.sContrato, tItensSolicSRV.iStatusItem, tItensSolicSRV.sReparo, tItensSolicSRV.dtDataBaixa, iFilialEmpresa, sContrato, sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 195594
        
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 195595
        
    If lErro = AD_SQL_SEM_DADOS Then gError 195596
        
    Do While lErro = AD_SQL_SUCESSO
    
        Set objItensSolicSrv = New ClassItensSolicSRV
    
        objItensSolicSrv.lNumIntDoc = tItensSolicSRV.lNumIntDoc
        objItensSolicSrv.lNumIntSolicSRV = objSolicSRV.lNumIntDoc
        objItensSolicSrv.sProduto = tItensSolicSRV.sProduto
        objItensSolicSrv.sProdutoDesc = tItensSolicSRV.sProdutoDesc
        objItensSolicSrv.dtDataVenda = tItensSolicSRV.dtDataVenda
        objItensSolicSrv.sServico = tItensSolicSRV.sServico
        objItensSolicSrv.sServicoDesc = tItensSolicSRV.sServicoDesc
        objItensSolicSrv.sUM = tItensSolicSRV.sUM
        objItensSolicSrv.dQuantidade = tItensSolicSRV.dQuantidade
        objItensSolicSrv.sLote = tItensSolicSRV.sLote
        objItensSolicSrv.iFilialOP = tItensSolicSRV.iFilialOP
        objItensSolicSrv.sSolicitacao = tItensSolicSRV.sSolicitacao
        objItensSolicSrv.sReparo = tItensSolicSRV.sReparo
        objItensSolicSrv.lGarantia = tItensSolicSRV.lGarantia
        objItensSolicSrv.sContrato = tItensSolicSRV.sContrato
        objItensSolicSrv.iStatusItem = tItensSolicSRV.iStatusItem
        objItensSolicSrv.dtDataBaixa = tItensSolicSRV.dtDataBaixa

        objSolicSRV.colItens.Add objItensSolicSrv

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 195597

    Loop

    Call Comando_Fechar(lComando)
    
    ItensSolicSRV_Le_Contrato = SUCESSO
    
    Exit Function

Erro_ItensSolicSRV_Le_Contrato:
    
    ItensSolicSRV_Le_Contrato = gErr
        
    Select Case gErr
    
        Case 195593
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 195594, 195595, 195597
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSSOLICSRV_CONTRATO", gErr, sContrato, sProduto)
        
        Case 195596
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195598)
            
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function ItensDeContratoSrv_Le_NumIntItemContrato(objItensDeContratoSRV As ClassItensDeContratoSrv, colItensDeContratoSRV As Collection) As Long

Dim lComando As Long
Dim lErro As Long
Dim tItensContSrv As typeItensDeContratoSrv
Dim objItensContSrv As ClassItensDeContratoSrv

On Error GoTo Erro_ItensDeContratoSrv_Le_NumIntItemContrato

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 195690
        
    tItensContSrv.sProduto = String(STRING_PRODUTO, 0)
    tItensContSrv.sLote = String(STRING_LOTE_RASTREAMENTO, 0)
    tItensContSrv.sCodigoContrato = String(STRING_CONTRATO_CODIGO, 0)
        
    lErro = Comando_Executar(lComando, "SELECT ItensDeContratoSrv.NumIntDoc, ItensDeContratoSrv.Codigo, Contratos.Codigo, ItensDeContratoSRV.FilialEmpresa, Contratos.DataIniContrato, Contratos.DataFimContrato, ItensDeContratoSrv.Quantidade, ItensDeContratoSrv.Lote, ItensDeContratoSrv.FilialOP, ItensDeContratoSrv.TipoGarantia, ItensDeContratoSrv.GarantiaTotal, ItensDeContrato.Produto FROM Contratos, ItensDeContrato, ItensDeContratoSrv WHERE " & _
    "Contratos.NumIntDoc = ItensDeContrato.NumIntContrato AND ItensDeContrato.NumIntDoc = ItensDeContratoSrv.NumIntItemContrato AND ItensDeContratoSRV.FilialEmpresa = ? AND ItensDeContratoSRV.NumIntItemContrato = ? ", _
    tItensContSrv.lNumIntDoc, tItensContSrv.lCodigo, tItensContSrv.sCodigoContrato, tItensContSrv.iFilialEmpresa, tItensContSrv.dtDataContratoIni, tItensContSrv.dtDataContratoFim, tItensContSrv.dQuantidade, tItensContSrv.sLote, tItensContSrv.iFilialOP, tItensContSrv.lTipoGarantia, tItensContSrv.iGarantiaTotal, tItensContSrv.sProduto, objItensDeContratoSRV.iFilialEmpresa, objItensDeContratoSRV.lNumIntItemContrato)
    If lErro <> AD_SQL_SUCESSO Then gError 195691
        
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 195692
        
    If lErro <> AD_SQL_SUCESSO Then gError 195693
        
    Do While lErro = AD_SQL_SUCESSO
        
        Set objItensContSrv = New ClassItensDeContratoSrv
        
        objItensContSrv.lNumIntDoc = tItensContSrv.lNumIntDoc
        objItensContSrv.lCodigo = tItensContSrv.lCodigo
        objItensContSrv.sCodigoContrato = tItensContSrv.sCodigoContrato
        objItensContSrv.iFilialEmpresa = tItensContSrv.iFilialEmpresa
        objItensContSrv.dtDataContratoIni = tItensContSrv.dtDataContratoIni
        objItensContSrv.dtDataContratoFim = tItensContSrv.dtDataContratoFim
        objItensContSrv.dQuantidade = tItensContSrv.dQuantidade
        objItensContSrv.sLote = tItensContSrv.sLote
        objItensContSrv.iFilialOP = tItensContSrv.iFilialOP
        objItensContSrv.lTipoGarantia = tItensContSrv.lTipoGarantia
        objItensContSrv.iGarantiaTotal = tItensContSrv.iGarantiaTotal
        objItensContSrv.sProduto = tItensContSrv.sProduto
    
        lErro = CF("ItensDeContratoSrvProd_Le", objItensContSrv)
        If lErro <> SUCESSO Then gError 195694
    
        lErro = CF("ItensDeContratoSrvNumSerie_Le", objItensContSrv)
        If lErro <> SUCESSO Then gError 195695

        colItensDeContratoSRV.Add objItensContSrv
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 195705
        
    Loop

    Call Comando_Fechar(lComando)
    
    ItensDeContratoSrv_Le_NumIntItemContrato = SUCESSO
    
    Exit Function

Erro_ItensDeContratoSrv_Le_NumIntItemContrato:
    
    ItensDeContratoSrv_Le_NumIntItemContrato = gErr
        
    Select Case gErr
    
        Case 195690
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 195691, 195692
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMCONTRATOSRV1", gErr, objItensContSrv.iFilialEmpresa, objItensContSrv.lNumIntItemContrato)
        
        Case 195693 To 195695
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195696)
            
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function RoteiroSRV_Le(ByVal objRoteiroSRV As ClassRoteiroSRV, Optional ByVal bCompleto As Boolean = True) As Long

Dim lErro As Long
Dim lComando As Long
Dim tRoteiroSRV As typeRoteiroSRV

On Error GoTo Erro_RoteiroSRV_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 194859

    'Alocação de espaço no buffer
    tRoteiroSRV.sServico = String(STRING_PRODUTO, 0)
    tRoteiroSRV.sVersao = String(STRING_KIT_VERSAO, 0)
    tRoteiroSRV.sDescricao = String(STRING_OBSERVACAO_OBSERVACAO, 0)
    tRoteiroSRV.sUM = String(STRING_UM_SIGLA, 0)
    tRoteiroSRV.sAutor = String(STRING_USUARIO_CODIGO, 0)

    'Le a tabelaRoteiroSRV
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Servico, Versao, Descricao, DataCriacao, DataUltModificacao, " & _
                "Quantidade, UM, Autor, Duracao FROM RoteiroSRV WHERE Servico= ?  AND Versao= ? ", _
                tRoteiroSRV.lNumIntDoc, tRoteiroSRV.sServico, tRoteiroSRV.sVersao, tRoteiroSRV.sDescricao, _
                tRoteiroSRV.dtDataCriacao, tRoteiroSRV.dtDataUltModificacao, tRoteiroSRV.dQuantidade, tRoteiroSRV.sUM, tRoteiroSRV.sAutor, _
                tRoteiroSRV.iDuracao, objRoteiroSRV.sServico, objRoteiroSRV.sVersao)
    If lErro <> AD_SQL_SUCESSO Then gError 194860

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194861

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objRoteiroSRV.lNumIntDoc = tRoteiroSRV.lNumIntDoc
    objRoteiroSRV.sServico = tRoteiroSRV.sServico
    objRoteiroSRV.sVersao = tRoteiroSRV.sVersao
    objRoteiroSRV.sDescricao = tRoteiroSRV.sDescricao
    objRoteiroSRV.dtDataCriacao = tRoteiroSRV.dtDataCriacao
    objRoteiroSRV.dtDataUltModificacao = tRoteiroSRV.dtDataUltModificacao
    objRoteiroSRV.dQuantidade = tRoteiroSRV.dQuantidade
    objRoteiroSRV.sUM = tRoteiroSRV.sUM
    objRoteiroSRV.sAutor = tRoteiroSRV.sAutor
    objRoteiroSRV.iDuracao = tRoteiroSRV.iDuracao
    
    If bCompleto Then
    
        lErro = RoteiroSRVOper_Le(objRoteiroSRV)
        If lErro <> SUCESSO Then gError 194862
    
    End If

    'Fecha Comando
    Call Comando_Fechar(lComando)

    RoteiroSRV_Le = SUCESSO

    Exit Function

Erro_RoteiroSRV_Le:

    RoteiroSRV_Le = gErr

    Select Case gErr

        Case 194859
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 194860, 194861
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ROTEIROSRV", gErr)

        Case ERRO_LEITURA_SEM_DADOS
        
        Case 194862

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194863)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function RoteiroSRV_Le_Versoes(ByVal sServico As String, ByVal colRot As Collection, Optional ByVal bCompleto As Boolean = False) As Long

Dim lErro As Long
Dim lComando As Long
Dim tRoteiroSRV As typeRoteiroSRV
Dim objRoteiroSRV As ClassRoteiroSRV

On Error GoTo Erro_RoteiroSRV_Le_Versoes

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 194859

    'Alocação de espaço no buffer
    tRoteiroSRV.sServico = String(STRING_PRODUTO, 0)
    tRoteiroSRV.sVersao = String(STRING_KIT_VERSAO, 0)
    tRoteiroSRV.sDescricao = String(STRING_OBSERVACAO_OBSERVACAO, 0)
    tRoteiroSRV.sUM = String(STRING_UM_SIGLA, 0)
    tRoteiroSRV.sAutor = String(STRING_USUARIO_CODIGO, 0)

    'Le a tabelaRoteiroSRV
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Servico, Versao, Descricao, DataCriacao, DataUltModificacao, " & _
                "Quantidade, UM, Autor, Duracao FROM RoteiroSRV WHERE Servico= ?  ", _
                tRoteiroSRV.lNumIntDoc, tRoteiroSRV.sServico, tRoteiroSRV.sVersao, tRoteiroSRV.sDescricao, _
                tRoteiroSRV.dtDataCriacao, tRoteiroSRV.dtDataUltModificacao, tRoteiroSRV.dQuantidade, tRoteiroSRV.sUM, tRoteiroSRV.sAutor, _
                tRoteiroSRV.iDuracao, sServico)
    If lErro <> AD_SQL_SUCESSO Then gError 194860

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194861

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objRoteiroSRV = New ClassRoteiroSRV

        objRoteiroSRV.lNumIntDoc = tRoteiroSRV.lNumIntDoc
        objRoteiroSRV.sServico = tRoteiroSRV.sServico
        objRoteiroSRV.sVersao = tRoteiroSRV.sVersao
        objRoteiroSRV.sDescricao = tRoteiroSRV.sDescricao
        objRoteiroSRV.dtDataCriacao = tRoteiroSRV.dtDataCriacao
        objRoteiroSRV.dtDataUltModificacao = tRoteiroSRV.dtDataUltModificacao
        objRoteiroSRV.dQuantidade = tRoteiroSRV.dQuantidade
        objRoteiroSRV.sUM = tRoteiroSRV.sUM
        objRoteiroSRV.sAutor = tRoteiroSRV.sAutor
        objRoteiroSRV.iDuracao = tRoteiroSRV.iDuracao
        
        If bCompleto Then
        
            lErro = RoteiroSRVOper_Le(objRoteiroSRV)
            If lErro <> SUCESSO Then gError 194862
        
        End If
        
        colRot.Add objRoteiroSRV
        
        'Busca Próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194861
        
    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    RoteiroSRV_Le_Versoes = SUCESSO

    Exit Function

Erro_RoteiroSRV_Le_Versoes:

    RoteiroSRV_Le_Versoes = gErr

    Select Case gErr

        Case 194859
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 194860, 194861
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ROTEIROSRV", gErr)

        Case ERRO_LEITURA_SEM_DADOS
        
        Case 194862

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194863)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function RoteiroSRVOper_Le(ByVal objRoteiroSRV As ClassRoteiroSRV) As Long

Dim lErro As Long
Dim lComando As Long
Dim tRoteiroSRVOper As typeRoteiroSRVOper
Dim objRoteiroSRVOper As ClassRoteiroSRVOper

On Error GoTo Erro_RoteiroSRVOper_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 194887

    'Alocação de espaço no buffer
    tRoteiroSRVOper.sObservacao = String(STRING_OBSERVACAO_OBSERVACAO, 0)

    'Le a tabelaRoteiroSRVOper
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntDocRotSrv, Seq, NumIntDocCompet, NumIntDocCT, Observacao, " & _
                "SeqPai, SeqArvore, Nivel, PosicaoArvore FROM RoteiroSRVOper WHERE NumIntDocRotSrv= ? ORDER BY Seq ", _
                tRoteiroSRVOper.lNumIntDoc, tRoteiroSRVOper.lNumIntDocRotSRV, tRoteiroSRVOper.iSeq, tRoteiroSRVOper.lNumIntDocCompet, _
                tRoteiroSRVOper.lNumIntDocCT, tRoteiroSRVOper.sObservacao, tRoteiroSRVOper.iSeqPai, tRoteiroSRVOper.iSeqArvore, tRoteiroSRVOper.iNivel, _
                tRoteiroSRVOper.iPosicaoArvore, objRoteiroSRV.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 194888

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194889

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objRoteiroSRVOper = New ClassRoteiroSRVOper

        objRoteiroSRVOper.lNumIntDoc = tRoteiroSRVOper.lNumIntDoc
        objRoteiroSRVOper.lNumIntDocRotSRV = tRoteiroSRVOper.lNumIntDocRotSRV
        objRoteiroSRVOper.iSeq = tRoteiroSRVOper.iSeq
        objRoteiroSRVOper.lNumIntDocCompet = tRoteiroSRVOper.lNumIntDocCompet
        objRoteiroSRVOper.lNumIntDocCT = tRoteiroSRVOper.lNumIntDocCT
        objRoteiroSRVOper.sObservacao = tRoteiroSRVOper.sObservacao
        objRoteiroSRVOper.iSeqPai = tRoteiroSRVOper.iSeqPai
        objRoteiroSRVOper.iSeqArvore = tRoteiroSRVOper.iSeqArvore
        objRoteiroSRVOper.iNivel = tRoteiroSRVOper.iNivel
        objRoteiroSRVOper.iPosicaoArvore = tRoteiroSRVOper.iPosicaoArvore
        
        objRoteiroSRV.colOperacoes.Add objRoteiroSRVOper
        
        lErro = RoteiroSRVOperMaq_Le(objRoteiroSRVOper)
        If lErro <> SUCESSO Then gError 194903
        
        lErro = RoteiroSRVOperMO_Le(objRoteiroSRVOper)
        If lErro <> SUCESSO Then gError 194904
        
        lErro = RoteiroSRVOperMP_Le(objRoteiroSRVOper)
        If lErro <> SUCESSO Then gError 194905

        'Busca Próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194890

    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    RoteiroSRVOper_Le = SUCESSO

    Exit Function

Erro_RoteiroSRVOper_Le:

    RoteiroSRVOper_Le = gErr

    Select Case gErr

        Case 194887
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 194888, 194889, 194890
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ROTEIROSRVOPER", gErr)
            
        Case 194903 To 194905
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194891)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function RoteiroSRVOperMaq_Le(ByVal objRoteiroSRVOper As ClassRoteiroSRVOper) As Long

Dim lErro As Long
Dim lComando As Long
Dim tRoteiroSRVOperMaq As typeRoteiroSRVOperMaq
Dim objRoteiroSRVOperMaq As ClassRoteiroSRVOperMaq

On Error GoTo Erro_RoteiroSRVOperMaq_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 194915

    'Alocação de espaço no buffer
    tRoteiroSRVOperMaq.sObs = String(STRING_OBSERVACAO_OBSERVACAO, 0)

    'Le a tabelaRoteiroSRVOperMaq
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntDocOper, Seq, CodMaq, FilialEmpMaq, Horas, Qtd, " & _
                "Obs FROM RoteiroSRVOperMaq WHERE NumIntDocOper= ? ORDER BY Seq", _
                tRoteiroSRVOperMaq.lNumIntDoc, tRoteiroSRVOperMaq.lNumIntDocOper, tRoteiroSRVOperMaq.iSeq, tRoteiroSRVOperMaq.iCodMaq, tRoteiroSRVOperMaq.iFilialEmpMaq, _
                tRoteiroSRVOperMaq.dHoras, tRoteiroSRVOperMaq.iQtd, tRoteiroSRVOperMaq.sObs, _
                objRoteiroSRVOper.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 194916

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194917

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objRoteiroSRVOperMaq = New ClassRoteiroSRVOperMaq

        objRoteiroSRVOperMaq.lNumIntDoc = tRoteiroSRVOperMaq.lNumIntDoc
        objRoteiroSRVOperMaq.lNumIntDocOper = tRoteiroSRVOperMaq.lNumIntDocOper
        objRoteiroSRVOperMaq.iSeq = tRoteiroSRVOperMaq.iSeq
        objRoteiroSRVOperMaq.iCodMaq = tRoteiroSRVOperMaq.iCodMaq
        objRoteiroSRVOperMaq.dHoras = tRoteiroSRVOperMaq.dHoras
        objRoteiroSRVOperMaq.iQtd = tRoteiroSRVOperMaq.iQtd
        objRoteiroSRVOperMaq.sObs = tRoteiroSRVOperMaq.sObs
        objRoteiroSRVOperMaq.iFilialEmpMaq = tRoteiroSRVOperMaq.iFilialEmpMaq
        
        objRoteiroSRVOper.colMaq.Add objRoteiroSRVOperMaq
    
        'Busca Próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194918
        
    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    RoteiroSRVOperMaq_Le = SUCESSO

    Exit Function

Erro_RoteiroSRVOperMaq_Le:

    RoteiroSRVOperMaq_Le = gErr

    Select Case gErr

        Case 194915
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 194916, 194917, 194918
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ROTEIROSRVOPERMAQ", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194919)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function RoteiroSRVOperMO_Le(ByVal objRoteiroSRVOper As ClassRoteiroSRVOper) As Long

Dim lErro As Long
Dim lComando As Long
Dim tRoteiroSRVOperMO As typeRoteiroSRVOperMO
Dim objRoteiroSRVOperMO As ClassRoteiroSRVOperMO

On Error GoTo Erro_RoteiroSRVOperMO_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 194943

    'Alocação de espaço no buffer
    tRoteiroSRVOperMO.sObs = String(STRING_OBSERVACAO_OBSERVACAO, 0)

    'Le a tabelaRoteiroSRVOperMO
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntDocOper, Seq, CodMO, Horas, Qtd, " & _
                "Obs FROM RoteiroSRVOperMO WHERE NumIntDocOper= ? ORDER BY Seq", _
                tRoteiroSRVOperMO.lNumIntDoc, tRoteiroSRVOperMO.lNumIntDocOper, tRoteiroSRVOperMO.iSeq, tRoteiroSRVOperMO.iCodMO, _
                tRoteiroSRVOperMO.dHoras, tRoteiroSRVOperMO.iQtd, tRoteiroSRVOperMO.sObs, _
                objRoteiroSRVOper.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 194944

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194945

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objRoteiroSRVOperMO = New ClassRoteiroSRVOperMO

        objRoteiroSRVOperMO.lNumIntDoc = tRoteiroSRVOperMO.lNumIntDoc
        objRoteiroSRVOperMO.lNumIntDocOper = tRoteiroSRVOperMO.lNumIntDocOper
        objRoteiroSRVOperMO.iSeq = tRoteiroSRVOperMO.iSeq
        objRoteiroSRVOperMO.iCodMO = tRoteiroSRVOperMO.iCodMO
        objRoteiroSRVOperMO.dHoras = tRoteiroSRVOperMO.dHoras
        objRoteiroSRVOperMO.iQtd = tRoteiroSRVOperMO.iQtd
        objRoteiroSRVOperMO.sObs = tRoteiroSRVOperMO.sObs
        
        objRoteiroSRVOper.colMO.Add objRoteiroSRVOperMO
        
        'Busca Próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194946
        
    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    RoteiroSRVOperMO_Le = SUCESSO

    Exit Function

Erro_RoteiroSRVOperMO_Le:

    RoteiroSRVOperMO_Le = gErr

    Select Case gErr

        Case 194943
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 194944, 194945, 194946
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ROTEIROSRVOPERMO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194947)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function RoteiroSRVOperMP_Le(ByVal objRoteiroSRVOper As ClassRoteiroSRVOper) As Long

Dim lErro As Long
Dim lComando As Long
Dim tRoteiroSRVOperMP As typeRoteiroSRVOperMP
Dim objRoteiroSRVOperMP As ClassRoteiroSRVOperMP

On Error GoTo Erro_RoteiroSRVOperMP_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 194971

    'Alocação de espaço no buffer
    tRoteiroSRVOperMP.sProduto = String(STRING_PRODUTO, 0)
    tRoteiroSRVOperMP.sUM = String(STRING_UM_SIGLA, 0)
    tRoteiroSRVOperMP.sVersao = String(STRING_KIT_VERSAO, 0)
    tRoteiroSRVOperMP.sObs = String(STRING_OBSERVACAO_OBSERVACAO, 0)

    'Le a tabelaRoteiroSRVOperMP
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntDocOper, Seq, Produto, Quantidade, UM, " & _
                "Versao, Composicao, OBS FROM RoteiroSRVOperMP WHERE NumIntDocOper= ? ORDER BY Seq", _
                tRoteiroSRVOperMP.lNumIntDoc, tRoteiroSRVOperMP.lNumIntDocOper, tRoteiroSRVOperMP.iSeq, tRoteiroSRVOperMP.sProduto, _
                tRoteiroSRVOperMP.dQuantidade, tRoteiroSRVOperMP.sUM, tRoteiroSRVOperMP.sVersao, tRoteiroSRVOperMP.iComposicao, tRoteiroSRVOperMP.sObs, _
                objRoteiroSRVOper.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 194972

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194973

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objRoteiroSRVOperMP = New ClassRoteiroSRVOperMP

        objRoteiroSRVOperMP.lNumIntDoc = tRoteiroSRVOperMP.lNumIntDoc
        objRoteiroSRVOperMP.lNumIntDocOper = tRoteiroSRVOperMP.lNumIntDocOper
        objRoteiroSRVOperMP.iSeq = tRoteiroSRVOperMP.iSeq
        objRoteiroSRVOperMP.sProduto = tRoteiroSRVOperMP.sProduto
        objRoteiroSRVOperMP.dQuantidade = tRoteiroSRVOperMP.dQuantidade
        objRoteiroSRVOperMP.sUM = tRoteiroSRVOperMP.sUM
        objRoteiroSRVOperMP.sVersao = tRoteiroSRVOperMP.sVersao
        objRoteiroSRVOperMP.iComposicao = tRoteiroSRVOperMP.iComposicao
        objRoteiroSRVOperMP.sObs = tRoteiroSRVOperMP.sObs
        
        objRoteiroSRVOper.colMP.Add objRoteiroSRVOperMP
        
        'Busca Próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194974
        
    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    RoteiroSRVOperMP_Le = SUCESSO

    Exit Function

Erro_RoteiroSRVOperMP_Le:

    RoteiroSRVOperMP_Le = gErr

    Select Case gErr

        Case 194971
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 194972, 194973, 194974
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ROTEIROSRVOPERMP", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194975)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function MovimentoEstoqueAp_Le(ByVal objMovEstoque As ClassMovEstoque) As Long

Dim lErro As Long
Dim objOSAp As New ClassOSAp
Dim tOSAp As typeOSAp
Dim lCodigoMovEst As Long
Dim sSQL As String, lCodigo As Long
Dim lComando As Long, dHora As Double

On Error GoTo Erro_MovimentoEstoqueAp_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 205840
    
    If objMovEstoque.objApontamento Is Nothing Then
        Set objMovEstoque.objApontamento = objOSAp
    Else
        Set objOSAp = objMovEstoque.objApontamento
    End If
    
    If objMovEstoque.iFilialEmpresa = 0 Then objMovEstoque.iFilialEmpresa = giFilialEmpresa
    
    sSQL = "SELECT NumIntDoc, Codigo,CodigoMovEst,Data,Hora,Cliente FROM ApontamentoSRV WHERE FilialEmpresa = ? AND "
    If objOSAp.lCodigo = 0 Then
        sSQL = sSQL & "CodigoMovEst = ?"
        If objMovEstoque.lCodigo < 0 Then
            lCodigo = objMovEstoque.lCodigo
        Else
            lCodigo = objOSAp.lCodigoMovEst
        End If
    Else
        sSQL = sSQL & "Codigo = ?"
        lCodigo = objOSAp.lCodigo
    End If
    
    'Le a tabela OSAp
    lErro = Comando_Executar(lComando, sSQL, tOSAp.lNumIntDoc, tOSAp.lCodigo, tOSAp.lCodigoMovEst, tOSAp.dtData, dHora, tOSAp.lCliente, objMovEstoque.iFilialEmpresa, lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 205841

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 205841
    
    If lErro <> AD_SQL_SUCESSO Then gError ERRO_LEITURA_SEM_DADOS
       
    objOSAp.lNumIntDoc = tOSAp.lNumIntDoc
    objOSAp.iFilialEmpresa = objMovEstoque.iFilialEmpresa
    objOSAp.lCodigo = tOSAp.lCodigo
    objOSAp.lCodigoMovEst = tOSAp.lCodigoMovEst
    objOSAp.dtData = tOSAp.dtData
    objOSAp.dtHora = CDate(dHora)
    objOSAp.lCliente = tOSAp.lCliente
    
    objMovEstoque.lCodigo = objOSAp.lCodigoMovEst
    objMovEstoque.dtData = objOSAp.dtData
    objMovEstoque.dtHora = objOSAp.dtHora
    
    lErro = CF("MovEstoqueItens_Le1", objMovEstoque, MOV_EST_APONT_SRV)
    If lErro <> SUCESSO And lErro <> 55387 Then gError ERRO_SEM_MENSAGEM
        
    lErro = OSApMaq_Le(objOSAp)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = OSApMO_Le(objOSAp)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Fecha Comando
    Call Comando_Fechar(lComando)

    MovimentoEstoqueAp_Le = SUCESSO

    Exit Function

Erro_MovimentoEstoqueAp_Le:

    MovimentoEstoqueAp_Le = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM, ERRO_LEITURA_SEM_DADOS
    
        Case 205840
        
        Case 205841

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205842)

    End Select
    
    'Fecha Comando
    Call Comando_Fechar(lComando)
    
    Exit Function
End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function OSApMaq_Le(ByVal objOSAp As ClassOSAp) As Long

Dim lErro As Long
Dim lComando As Long
Dim tOSApMaq As typeOSApMaq
Dim objOSApMaq As ClassOSApMaq

On Error GoTo Erro_OSApMaq_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 205843

    'Alocação de espaço no buffer
    tOSApMaq.sOS = String(STRING_ORDEM_DE_PRODUCAO, 0)
    tOSApMaq.sProdutoOS = String(STRING_PRODUTO, 0)

    'Le a tabela OSApMaq
    lErro = Comando_Executar(lComando, "SELECT Seq, NumIntDocMaq, HorasGastas, Quantidade, OS, " & _
                "ProdutoOS FROM OSApMaq WHERE CodigoAp= ?  AND FilialEmpresa= ? ORDER BY Seq ", _
                tOSApMaq.iSeq, tOSApMaq.lNumIntDocMaq, tOSApMaq.dHorasGastas, _
                tOSApMaq.iQuantidade, tOSApMaq.sOS, tOSApMaq.sProdutoOS, _
                objOSAp.lCodigo, objOSAp.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 205844

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 205845

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objOSApMaq = New ClassOSApMaq

        objOSApMaq.lCodigoAp = objOSAp.lCodigo
        objOSApMaq.iFilialEmpresa = objOSAp.iFilialEmpresa
        objOSApMaq.iSeq = tOSApMaq.iSeq
        objOSApMaq.lNumIntDocMaq = tOSApMaq.lNumIntDocMaq
        objOSApMaq.dHorasGastas = tOSApMaq.dHorasGastas
        objOSApMaq.iQuantidade = tOSApMaq.iQuantidade
        objOSApMaq.sOS = tOSApMaq.sOS
        objOSApMaq.sProdutoOS = tOSApMaq.sProdutoOS
        
        objOSAp.colMaq.Add objOSApMaq

        'Busca Próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 205846

    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    OSApMaq_Le = SUCESSO

    Exit Function

Erro_OSApMaq_Le:

    OSApMaq_Le = gErr

    Select Case gErr

        Case 205843
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 205844 To 205846
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_OSAPMAQ", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205847)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function OSApMO_Le(ByVal objOSAp As ClassOSAp) As Long

Dim lErro As Long
Dim lComando As Long
Dim tOSApMO As typeOSApMO
Dim objOSApMO As ClassOSApMO
Dim iTipo As Integer

On Error GoTo Erro_OSApMO_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 205848

    'Alocação de espaço no buffer
    tOSApMO.sOS = String(STRING_ORDEM_DE_PRODUCAO, 0)
    tOSApMO.sProdutoOS = String(STRING_PRODUTO, 0)

    'Le a tabelaOSApMO
    lErro = Comando_Executar(lComando, "SELECT Seq, CodigoMO, HorasGastas, OS, ProdutoOS, Tipo FROM OSApMO, MaoDeObra WHERE CodigoMO = Codigo AND CodigoAp= ?  AND FilialEmpresa= ? ORDER BY Seq ", _
                tOSApMO.iSeq, tOSApMO.lCodigoMO, tOSApMO.dHorasGastas, _
                tOSApMO.sOS, tOSApMO.sProdutoOS, iTipo, _
                objOSAp.lCodigo, objOSAp.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 205849

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 205850

    'Sem Dados
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objOSApMO = New ClassOSApMO

        objOSApMO.lCodigoAp = objOSAp.lCodigo
        objOSApMO.iFilialEmpresa = objOSAp.iFilialEmpresa
        objOSApMO.iSeq = tOSApMO.iSeq
        objOSApMO.lCodigoMO = tOSApMO.lCodigoMO
        objOSApMO.dHorasGastas = tOSApMO.dHorasGastas
        objOSApMO.sOS = tOSApMO.sOS
        objOSApMO.sProdutoOS = tOSApMO.sProdutoOS
        objOSApMO.iTipo = iTipo
        
        objOSAp.colMO.Add objOSApMO
        
        'Busca Próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 205851

    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    OSApMO_Le = SUCESSO

    Exit Function

Erro_OSApMO_Le:

    OSApMO_Le = gErr

    Select Case gErr

        Case 205848
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 205849 To 205851
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_OSAPMO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205852)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function ItemOS_Le(ByVal objItemOS As ClassItemOS) As Long

Dim lErro As Long
Dim lComando As Long
Dim tItemOS As typeItemOS

On Error GoTo Erro_ItemOS_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 205890

    'Alocação de espaço no buffer
    tItemOS.sCodigo = String(STRING_ORDEM_DE_PRODUCAO, 0)
    tItemOS.sServico = String(STRING_PRODUTO, 0)
    tItemOS.sCcl = String(STRING_CCL, 0)
    tItemOS.sSiglaUM = String(STRING_UM_SIGLA, 0)
    tItemOS.sVersao = String(STRING_KIT_VERSAO, 0)
    tItemOS.sObservacao = String(STRING_OBSERVACAO_OBSERVACAO, 0)

    'Le a tabelaItemOS
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, FilialEmpresa, Codigo, Item, Servico, Ccl, " & _
                "SiglaUM, Quantidade, DataInicio, DataFim, Prioridade, StatusItem, Versao, " & _
                "Observacao FROM ItemOS WHERE FilialEmpresa= ?  AND Codigo= ?  AND Servico= ? ", _
                tItemOS.lNumIntDoc, tItemOS.iFilialEmpresa, tItemOS.sCodigo, tItemOS.iItem, _
                tItemOS.sServico, tItemOS.sCcl, tItemOS.sSiglaUM, tItemOS.dQuantidade, tItemOS.dtDataInicio, _
                tItemOS.dtDataFim, tItemOS.iPrioridade, tItemOS.iStatusItem, tItemOS.sVersao, tItemOS.sObservacao, _
                objItemOS.iFilialEmpresa, objItemOS.sCodigo, objItemOS.sServico)
    If lErro <> AD_SQL_SUCESSO Then gError 205891

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 205892

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objItemOS.lNumIntDoc = tItemOS.lNumIntDoc
    objItemOS.iFilialEmpresa = tItemOS.iFilialEmpresa
    objItemOS.sCodigo = tItemOS.sCodigo
    objItemOS.iItem = tItemOS.iItem
    objItemOS.sServico = tItemOS.sServico
    objItemOS.sCcl = tItemOS.sCcl
    objItemOS.sSiglaUM = tItemOS.sSiglaUM
    objItemOS.dQuantidade = tItemOS.dQuantidade
    objItemOS.dtDataInicio = tItemOS.dtDataInicio
    objItemOS.dtDataFim = tItemOS.dtDataFim
    objItemOS.iPrioridade = tItemOS.iPrioridade
    objItemOS.iStatusItem = tItemOS.iStatusItem
    objItemOS.sVersao = tItemOS.sVersao
    objItemOS.sObservacao = tItemOS.sObservacao
    
    lErro = ItemOSOperacoes_Le(objItemOS)
    If lErro <> SUCESSO And lErro <> 195070 Then gError 195063

    'Fecha Comando
    Call Comando_Fechar(lComando)

    ItemOS_Le = SUCESSO

    Exit Function

Erro_ItemOS_Le:

    ItemOS_Le = gErr

    Select Case gErr
    
        Case 195063

        Case 205890
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 205891, 205892
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMOS", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205893)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function CamposGenericos_ValidaExclusao_TipoOS() As Long
'Verifica se é válida a exclusão de valores para o campo Requisitante

Dim lComando As Long
Dim lErro As Long
Dim lTipo As Long

On Error GoTo Erro_CamposGenericos_ValidaExclusao_TipoOS
    
    'Abre comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 141377
    
    'Busca Tipos de Paradas em CTMaquinasParadas que não existam na tabela CamposGenericos Valores
    lErro = Comando_Executar(lComando, "SELECT Tipo FROM OS WHERE Tipo NOT IN (SELECT CodValor FROM CamposGenericosValores WHERE CodCampo = ?) AND Tipo <> 0", lTipo, CAMPOSGENERICOS_TIPOOS)
    If lErro <> AD_SQL_SUCESSO Then gError 141378
    
    'Posiciona no primeiro registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 141379
    
    'Se encontrou algum registro => erro de inconsistência
    If lErro = AD_SQL_SUCESSO Then gError 141380

    'Fecha comando
    Call Comando_Fechar(lComando)
    
    CamposGenericos_ValidaExclusao_TipoOS = SUCESSO
    
    Exit Function
    
Erro_CamposGenericos_ValidaExclusao_TipoOS:

    CamposGenericos_ValidaExclusao_TipoOS = gErr
    
    Select Case gErr
    
        Case 141377
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 141378, 141379
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_OS", gErr)
        
        Case 141380
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CAMPOSGENERICOSVALORES", gErr, lTipo, CAMPOSGENERICOS_TIPOOS)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150281)
    
    End Select

    'Fecha comando
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function Solicitacao_Le_Gastos(ByVal objSol As ClassSolicSRV, ByVal colItens As Collection, Optional ByVal bApont As Boolean = False) As Long
'Le os gastos ligados a OS ou apontamento para solicitação passada
'Desconsidera serviços em garantia ou em contrato
Dim lErro As Long
Dim alComando(1 To 3) As Long
Dim iIndice As Integer
Dim objSolBD As New ClassSolicSRV
Dim objItemSol As ClassItensSolicSRV
Dim sServico As String
Dim sSQL(1 To 3) As String
Dim sProduto As String, sUM   As String, dQuantidade As Double, iClasseUM As Integer, sDescricao As String, iNatureza As Integer
Dim objItem As ClassItensSolicSRV
Dim bAchou As Boolean, bCobra As Boolean, dFator As Double

On Error GoTo Erro_Solicitacao_Le_Gastos

    'abrir comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 206422
    Next

    If bApont Then
        'Le os dados do apontamento
        'Le os produtos
        sSQL(1) = "SELECT I.Servico, P.Codigo, M.SiglaUM, M.Quantidade , P.ClasseUM, P.Descricao, P.Natureza "
        sSQL(1) = sSQL(1) & "FROM OS, ItemOS AS I, MovimentoEstoque AS M, Produtos As P "
        sSQL(1) = sSQL(1) & "WHERE OS.Codigo = I.Codigo AND OS.FilialEmpresa = I.FilialEmpresa AND M.NumIntDocOrigem = I.NumIntDoc AND M.TipoNumIntDocOrigem = 8 AND M.TipoMov = 262  AND "
        sSQL(1) = sSQL(1) & "P.Codigo = M.Produto AND OS.CodSolSRV = ? AND OS.Produto = ? AND OS.Lote = ? AND OS.FilialOP = ? AND OS.FilialEmpresa = ? "
        'Le as máquinas
        sSQL(2) = "SELECT M.ProdutoOS, P.Codigo, 'h' as UMProduto, M.HorasGastas As Quantidade , P.ClasseUM, P.Descricao, P.Natureza "
        sSQL(2) = sSQL(2) & "FROM OS, ItemOS AS I, OSApMaq AS M, Maquinas AS Maq, Produtos As P "
        sSQL(2) = sSQL(2) & "WHERE OS.Codigo = i.Codigo And OS.FilialEmpresa = i.FilialEmpresa And M.NumIntDocMaq = Maq.NumIntDoc AND M.OS = OS.Codigo AND M.ProdutoOS = I.Servico AND "
        sSQL(2) = sSQL(2) & "P.Codigo = Maq.Produto AND OS.CodSolSRV = ? AND OS.Produto = ? AND OS.Lote = ? AND OS.FilialOP = ? AND OS.FilialEmpresa = ? "
        'Le a mão-de-obra
        sSQL(3) = "SELECT M.ProdutoOS, P.Codigo, 'h' as UMProduto, M.HorasGastas As Quantidade , P.ClasseUM, P.Descricao, P.Natureza "
        sSQL(3) = sSQL(3) & "FROM OS, ItemOS AS I, OSApMO AS M, MaoDeObra AS MO, Produtos As P, TiposDeMaoDeObra AS T "
        sSQL(3) = sSQL(3) & "WHERE OS.Codigo = I.Codigo AND OS.FilialEmpresa = I.FilialEmpresa AND M.CodigoMO = MO.Codigo AND M.OS = OS.Codigo AND M.ProdutoOS = I.Servico AND "
        sSQL(3) = sSQL(3) & "MO.Tipo = T.Codigo AND P.Codigo = T.Produto AND OS.CodSolSRV = ? AND OS.Produto = ? AND OS.Lote = ? AND OS.FilialOP = ? AND OS.FilialEmpresa = ? "
        
    Else
        'Le os dados da OS
        'Le os produtos
        sSQL(1) = "SELECT I.Servico, P.Produto, P.UMProduto, P.Quantidade, Prod.ClasseUM, Prod.Descricao, Prod.Natureza "
        sSQL(1) = sSQL(1) & "FROM OS, ItemOS AS I, ItemOSOperacoes AS O, ItemOSOperacoesPecas AS P, Produtos As Prod "
        sSQL(1) = sSQL(1) & "WHERE OS.Codigo = I.Codigo AND OS.FilialEmpresa = I.FilialEmpresa AND I.NumIntDoc = O.NumIntDocItemOS  AND "
        sSQL(1) = sSQL(1) & "O.NumIntDoc = P.NumIntDocOper AND Prod.Codigo = P.Produto AND OS.CodSolSRV = ? AND OS.Produto = ? AND OS.Lote = ? AND OS.FilialOP = ? AND OS.FilialEmpresa = ? "
        'Le as máquinas
        sSQL(2) = "SELECT I.Servico, P.Codigo, 'h' as UMProduto, M.Horas As Quantidade , P.ClasseUM, P.Descricao, P.Natureza "
        sSQL(2) = sSQL(2) & "FROM OS, ItemOS AS I, ItemOSOperacoes AS O, ItemOSOperacoesMaquinas AS M , Produtos As P, Maquinas AS Maq "
        sSQL(2) = sSQL(2) & "WHERE OS.Codigo = I.Codigo AND OS.FilialEmpresa = I.FilialEmpresa AND I.NumIntDoc = O.NumIntDocItemOS  AND M.NumIntDocMaq = Maq.NumIntDoc AND "
        sSQL(2) = sSQL(2) & "O.NumIntDoc = M.NumIntDocOper AND P.Codigo = Maq.Produto AND OS.CodSolSRV = ? AND OS.Produto = ? AND OS.Lote = ? AND OS.FilialOP = ? AND OS.FilialEmpresa = ? "
        'Le a mão-de-obra
        sSQL(3) = "SELECT I.Servico, P.Codigo, 'h' as UMProduto, M.Horas As Quantidade , P.ClasseUM, P.Descricao, P.Natureza "
        sSQL(3) = sSQL(3) & "FROM OS, ItemOS AS I, ItemOSOperacoes AS O, ItemOSOperacoesMO AS M , Produtos As P, TiposDeMaoDeObra AS T "
        sSQL(3) = sSQL(3) & "WHERE OS.Codigo = I.Codigo AND OS.FilialEmpresa = I.FilialEmpresa AND I.NumIntDoc = O.NumIntDocItemOS  AND M.CodigoMO = T.Codigo AND "
        sSQL(3) = sSQL(3) & "O.NumIntDoc = M.NumIntDocOper AND P.Codigo = T.Produto AND OS.CodSolSRV = ? AND OS.Produto = ? AND OS.Lote = ? AND OS.FilialOP = ? AND OS.FilialEmpresa = ? "
    End If
    
    objSolBD.lNumIntDoc = objSol.lNumIntDoc
    
    lErro = CF("ItensSolicSRV_Le", objSolBD)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    For Each objItemSol In objSolBD.colItens
    
        objItemSol.iFilialEmpresa = objSol.iFilialEmpresa
               
        For iIndice = 1 To 3
        
            sServico = String(STRING_PRODUTO, 0)
            sProduto = String(STRING_PRODUTO, 0)
            sUM = String(STRING_UM_SIGLA, 0)
            sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
        
            'Busca Tipos de Paradas em CTMaquinasParadas que não existam na tabela CamposGenericos Valores
            lErro = Comando_Executar(alComando(iIndice), sSQL(iIndice), sServico, sProduto, sUM, dQuantidade, iClasseUM, sDescricao, iNatureza, objSol.lCodigo, objItemSol.sProduto, objItemSol.sLote, objItemSol.iFilialOP, objItemSol.iFilialEmpresa)
            If lErro <> AD_SQL_SUCESSO Then gError 206423
            
            'Posiciona no primeiro registro
            lErro = Comando_BuscarPrimeiro(alComando(iIndice))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 206424
            
            Do While lErro <> AD_SQL_SEM_DADOS
            
                'Testa para ver se o serviço não está em garantia ou em contrato
                If iIndice = 1 Then
                    'Vê se a peça está coberta pela garantia ou contrato
                    lErro = ItemSolicSRV_Testa_Cobranca(objItemSol, objSol.dtData, sProduto, bCobra)
                Else
                    'Vê se o serviço da OS está coberto pela garantia ou contrato
                    lErro = ItemSolicSRV_Testa_Cobranca(objItemSol, objSol.dtData, sServico, bCobra)
                End If
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
                'Se vai cobrar pelo serviço
                If bCobra Then
                    
                    bAchou = False
                    For Each objItem In colItens
                        If objItem.sProduto = sProduto Then
                            bAchou = True
                            
                            lErro = CF("UM_Conversao_Trans", iClasseUM, sUM, objItem.sUM, dFator)
                            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                            
                            objItem.dQuantidade = objItem.dQuantidade + dQuantidade * dFator
                            
                            Exit For
                        End If
                    Next
                    
                    If Not bAchou Then
                        Set objItem = New ClassItensSolicSRV
                        colItens.Add objItem
                        objItem.sProduto = sProduto
                        objItem.sProdutoDesc = sDescricao
                        objItem.sUM = sUM
                        objItem.dQuantidade = dQuantidade
                        If iNatureza = NATUREZA_PROD_SERVICO Then
                            objItem.iTipo = ITEMNF_TIPO_SERVICO
                        Else
                            objItem.iTipo = ITEMNF_TIPO_PECA
                        End If
                    End If
                    
                End If
            
                lErro = Comando_BuscarProximo(alComando(iIndice))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 206425
            
            Loop
        
        Next
        
    Next
    
   'fechar comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Solicitacao_Le_Gastos = SUCESSO
    
    Exit Function
    
Erro_Solicitacao_Le_Gastos:

    Solicitacao_Le_Gastos = gErr
     
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
     
        Case 206422
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 206423 To 206425
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_OS", gErr)
     
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206426)
     
    End Select
     
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Public Function ItemSolicSRV_Testa_Cobranca(ByVal objItemSol As ClassItensSolicSRV, ByVal dtDataSol As Date, ByVal sServico As String, bCobra As Boolean, Optional bEmGarantia As Boolean = False) As Long

Dim lErro As Long
Dim objGarantiaProd As ClassGarantiaProduto
Dim objItensContSrv As New ClassItensDeContratoSrv
Dim objItensContProd As New ClassItensDeContratoSrvProd
Dim objGarantia As New ClassGarantia
Dim colItensContSrv As New Collection

On Error GoTo Erro_ItemSolicSRV_Testa_Cobranca

    bCobra = True
    bEmGarantia = False
    
    'Se tem garantia
    If objItemSol.lGarantia <> 0 Then
    
        'Le a Garantia
        objGarantia.lCodigo = objItemSol.lGarantia
        objGarantia.iFilialEmpresa = objItemSol.iFilialEmpresa
    
        lErro = CF("Garantia_Le", objGarantia)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        If objGarantia.iGarantiaTotal = MARCADO Then
        'Total com exceção dos serviços da coleção
            
            'PADRÃO
            'Está no prazo -> Não Cobra-> Está em garantia
            If DateAdd("d", objGarantia.dtDataVenda, objGarantia.iGarantiaTotalPrazo) >= gdtDataHoje Then
                bCobra = False
                bEmGarantia = True
            Else
            'Garantia expirada -> Cobra
                bCobra = True
                bEmGarantia = False
            End If
            
            'EXCEÇÕES
            For Each objGarantiaProd In objGarantia.colProduto
                If objGarantiaProd.sProduto = sServico Then
                'Se esse serviço\peça é uma exceção
                    If DateAdd("d", objGarantia.dtDataVenda, objGarantia.iGarantiaTotalPrazo) >= gdtDataHoje Then
                    'Está no prazo -> Não Cobra-> Está em garantia
                        bCobra = False
                        bEmGarantia = True
                    Else
                    'Garantia expirada -> Cobra
                        bCobra = True
                        bEmGarantia = False
                    End If
                    Exit For
                End If
            Next
        Else
        'Garantia somente para os serviços abaixo
            
            'PADRÃO
            'Sem Garantia -> Cobra
            bCobra = True
            bEmGarantia = False
            
            'EXCEÇÕES
            For Each objGarantiaProd In objGarantia.colProduto
                If objGarantiaProd.sProduto = sServico Then
                    If DateAdd("d", objGarantia.dtDataVenda, objGarantia.iGarantiaTotalPrazo) >= gdtDataHoje Then
                    'Está no prazo -> Não Cobra-> Está em garantia
                        bCobra = False
                        bEmGarantia = True
                    Else
                    'Garantia expirada -> Cobra
                        bCobra = True
                        bEmGarantia = False
                    End If
                    Exit For
                End If
            Next
        End If
    End If
    
    'Se tem contrato e já tem garantia para o serviço
    If Len(Trim(objItemSol.sContrato)) <> 0 And bCobra Then
    
        objItensContSrv.sCodigoContrato = objItemSol.sContrato
        objItensContSrv.iFilialEmpresa = objItemSol.iFilialEmpresa
    
        lErro = CF("Contrato_Le_Srv", objItensContSrv, colItensContSrv)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        For Each objItensContSrv In colItensContSrv
            'Se a solicitação está dentro da vigência do contrato
            If objItensContSrv.dtDataContratoIni >= dtDataSol And (objItensContSrv.dtDataContratoFim <= dtDataSol Or objItensContSrv.dtDataContratoFim = DATA_NULA) Then
                
                If objItensContSrv.iGarantiaTotal = MARCADO Then
                    'Total com exceção dos serviços abaixo
                    bCobra = False
                    For Each objItensContProd In objItensContSrv.colProduto
                        If objItensContProd.sProduto = sServico Then
                            bCobra = True
                            Exit For
                        End If
                    Next
                Else
                    'Garantia somente para os serviços abaixo
                    For Each objItensContProd In objItensContSrv.colProduto
                        If objItensContProd.sProduto = sServico Then
                            bCobra = False
                            Exit For
                        End If
                    Next
                End If
            End If
        Next
    End If

    ItemSolicSRV_Testa_Cobranca = SUCESSO
    
    Exit Function
    
Erro_ItemSolicSRV_Testa_Cobranca:

    ItemSolicSRV_Testa_Cobranca = gErr
     
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
     
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206427)
     
    End Select
    
    Exit Function
    
End Function

Public Function Garantia_Le_Gastos(ByVal lFabricante As Long, ByVal iFilialFabricante As Integer, ByVal colItens As Collection) As Long
'Le os gastos ligados a OS ou apontamento para solicitação passada
'Desconsidera serviços em garantia ou em contrato
Dim lErro As Long
Dim alComando(0 To 4) As Long
Dim iIndice As Integer
Dim lNumIntSolAnt As Long, lNumIntDoc As Long, lCodigoAp As Long, iSeq As Integer, sOS As String
Dim objSolBD As New ClassSolicSRV
Dim sServico As String
Dim sSQL(0 To 3) As String
Dim sTipo(1 To 3) As String
Dim sProduto As String, sUM   As String, dQuantidade As Double, iClasseUM As Integer, sDescricao As String, iNatureza As Integer
Dim objItem As ClassItensSolicSRV
Dim objSol As ClassSolicSRV
Dim bAchou As Boolean, bCobra As Boolean, dFator As Double, bEmGarantia As Boolean
Dim tSol As typeSolicSRV, tItemSol As typeItensSolicSRV
Dim objDet As ClassOSApDet, sDet As String, dQtdFaltaFaturarUM As Double, lNumIntDocAP As Long
Dim colItensSol As New Collection

On Error GoTo Erro_Garantia_Le_Gastos

    'abrir comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 206422
    Next
    
    tItemSol.sProduto = String(STRING_PRODUTO, 0)
    tItemSol.sProdutoDesc = String(STRING_PRODUTO_DESCRICAO, 0)
    tItemSol.sLote = String(STRING_LOTE_RASTREAMENTO, 0)
    
    'Le os itens de solicitação em garantia para esse fabricante
    sSQL(0) = "SELECT S.NumIntDoc, S.FilialEmpresa, S.Codigo, S.Data, I.NumIntDoc, I.Produto, P.Descricao, I.Garantia, I.Lote, I.FilialOP "
    sSQL(0) = sSQL(0) & "FROM SolicitacaoSRV AS S, ItensSolicSRV AS I, Garantia AS G, Produtos AS P "
    sSQL(0) = sSQL(0) & "WHERE I.NumIntSolicSRV = S.NumIntDoc AND I.Garantia = G.Codigo AND I.Produto = P.Codigo AND G.CliFabr = ? AND G.FilialCliFabr = ? "
    
    lErro = Comando_Executar(alComando(0), sSQL(0), tSol.lNumIntDoc, tSol.iFilialEmpresa, tSol.lCodigo, tSol.dtData, tItemSol.lNumIntDoc, tItemSol.sProduto, tItemSol.sProdutoDesc, tItemSol.lGarantia, tItemSol.sLote, tItemSol.iFilialOP, lFabricante, iFilialFabricante)
    If lErro <> AD_SQL_SUCESSO Then gError 206423
    
    'Posiciona no primeiro registro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 206424
    
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objSol = New ClassSolicSRV
        objSol.lNumIntDoc = tSol.lNumIntDoc
        objSol.iFilialEmpresa = tSol.iFilialEmpresa
        objSol.lCodigo = tSol.lCodigo
        objSol.dtData = tSol.dtData
        Set objItem = New ClassItensSolicSRV
        Set objItem.gobjInfoUsu = objSol
        colItensSol.Add objItem
        
        objItem.lNumIntDoc = tItemSol.lNumIntDoc
        objItem.lGarantia = tItemSol.lGarantia
        objItem.sProduto = tItemSol.sProduto
        objItem.sProdutoDesc = tItemSol.sProdutoDesc
        objItem.sLote = tItemSol.sLote
        objItem.iFilialOP = tItemSol.iFilialOP
        objItem.iFilialEmpresa = tSol.iFilialEmpresa
    
        'Posiciona no primeiro registro
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 206424
    
    Loop
    
    '*****************************************
    '1 - FALTA NÃO TRAZER O QUE JÁ FOI FATURADO (ItemNFItemAPSRV)
    '2 - IMPLEMENTAR CAMPO NUMINTDOC EM OSApMo e OSApMaq (Já foi acrescentado na tabela, falta gravação\leitura)
    '3 - Não deixar mexer em apontamento já utilizados em nfs (excluir, excluir itens ou alterá-los)

    'Le os dados do apontamento
    'Le os produtos
    sSQL(1) = "SELECT '' As Det, OS.Codigo, M.NumIntDoc, A.Codigo, M.Item, I.Servico, P.Codigo, M.SiglaUM, M.Quantidade , P.ClasseUM, P.Descricao, P.Natureza, A.NumIntDoc "
    sSQL(1) = sSQL(1) & "FROM OS, ItemOS AS I, MovimentoEstoque AS M, Produtos As P, ApontamentoSRV AS A "
    sSQL(1) = sSQL(1) & "WHERE OS.Codigo = I.Codigo AND OS.FilialEmpresa = I.FilialEmpresa AND M.NumIntDocOrigem = I.NumIntDoc AND M.TipoNumIntDocOrigem = 8 AND "
    sSQL(1) = sSQL(1) & "P.Codigo = M.Produto AND A.CodigoMovEst = M.Codigo AND OS.CodSolSRV = ? AND OS.Produto = ? AND OS.Lote = ? AND OS.FilialOP = ? AND OS.FilialEmpresa = ? "
    sTipo(1) = "Peças"
    'Le as máquinas
    sSQL(2) = "SELECT CONVERT(VARCHAR, Maq.Codigo) + '-' + Maq.NomeReduzido As Det, OS.Codigo, M.NumIntDoc, M.CodigoAP, M.Seq, M.ProdutoOS, P.Codigo, 'h' as UMProduto, M.HorasGastas As Quantidade , P.ClasseUM, P.Descricao, P.Natureza, A.NumIntDoc "
    sSQL(2) = sSQL(2) & "FROM OS, ItemOS AS I, OSApMaq AS M, Maquinas AS Maq, Produtos As P, ApontamentoSRV AS A "
    sSQL(2) = sSQL(2) & "WHERE OS.Codigo = i.Codigo And OS.FilialEmpresa = i.FilialEmpresa AND M.CodigoAP = A.Codigo AND M.NumIntDocMaq = Maq.NumIntDoc AND M.OS = OS.Codigo AND M.ProdutoOS = I.Servico AND "
    sSQL(2) = sSQL(2) & "P.Codigo = Maq.Produto AND OS.CodSolSRV = ? AND OS.Produto = ? AND OS.Lote = ? AND OS.FilialOP = ? AND OS.FilialEmpresa = ? "
    sTipo(2) = "Máquinas"
    'Le a mão-de-obra
    sSQL(3) = "SELECT CONVERT(VARCHAR, MO.Codigo) + '-' + MO.NomeReduzido As Det, OS.Codigo, M.NumIntDoc, M.CodigoAP, M.Seq, M.ProdutoOS, P.Codigo, 'h' as UMProduto, M.HorasGastas As Quantidade , P.ClasseUM, P.Descricao, P.Natureza, A.NumIntDoc "
    sSQL(3) = sSQL(3) & "FROM OS, ItemOS AS I, OSApMO AS M, MaoDeObra AS MO, Produtos As P, TiposDeMaoDeObra AS T, ApontamentoSRV AS A "
    sSQL(3) = sSQL(3) & "WHERE OS.Codigo = I.Codigo AND OS.FilialEmpresa = I.FilialEmpresa AND M.CodigoAP = A.Codigo AND M.CodigoMO = MO.Codigo AND M.OS = OS.Codigo AND M.ProdutoOS = I.Servico AND "
    sSQL(3) = sSQL(3) & "MO.Tipo = T.Codigo AND P.Codigo = T.Produto AND OS.CodSolSRV = ? AND OS.Produto = ? AND OS.Lote = ? AND OS.FilialOP = ? AND OS.FilialEmpresa = ? "
    sTipo(3) = "Mão-de-obra"

    'Para cada item com garantia
    For Each objItem In colItensSol
               
        For iIndice = 1 To 3
               
            sServico = String(STRING_PRODUTO, 0)
            sProduto = String(STRING_PRODUTO, 0)
            sUM = String(STRING_UM_SIGLA, 0)
            sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
            sOS = String(STRING_OS_CODIGO, 0)
            sDet = String(255, 0)
        
            'Busca Peças\Mão-de-obra\Máquinas apontadas
            lErro = Comando_Executar(alComando(iIndice), sSQL(iIndice), sDet, sOS, lNumIntDoc, lCodigoAp, iSeq, sServico, sProduto, sUM, dQuantidade, iClasseUM, sDescricao, iNatureza, lNumIntDocAP, objSol.lCodigo, objItem.sProduto, objItem.sLote, objItem.iFilialOP, objItem.iFilialEmpresa)
            If lErro <> AD_SQL_SUCESSO Then gError 206423
            
            'Posiciona no primeiro registro
            lErro = Comando_BuscarPrimeiro(alComando(iIndice))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 206424
            
            Do While lErro <> AD_SQL_SEM_DADOS
            
                'Testa para ver se o serviço está em garantia
                If iIndice = 1 Then
                    'Vê se a peça está coberta pela garantia
                    lErro = ItemSolicSRV_Testa_Cobranca(objItem, objSol.dtData, sProduto, bCobra, bEmGarantia)
                Else
                    'Vê se o serviço da OS está coberto pela garantia
                    lErro = ItemSolicSRV_Testa_Cobranca(objItem, objSol.dtData, sServico, bCobra, bEmGarantia)
                End If
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
                'Se está em Garantia
                If bEmGarantia Then
                
                    'Retorna só o que falta faturar
                    lErro = ItemAP_Testa_Qtd_Falta_Faturar(alComando(4), lNumIntDoc, iIndice, iClasseUM, sUM, dQuantidade, dQtdFaltaFaturarUM)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                                                       
                    'Guarda o detalhamento
                    Set objDet = New ClassOSApDet
                    objDet.iTipo = iIndice
                    objDet.dQuantidade = dQtdFaltaFaturarUM 'dQuantidade
                    objDet.iFilialEmpresa = objItem.iFilialEmpresa
                    objDet.iSeq = iSeq
                    objDet.iTipo = iIndice
                    objDet.lCodigoAp = lCodigoAp
                    objDet.lNumIntDoc = lNumIntDoc
                    objDet.lNumIntDocAP = lNumIntDocAP
                    objDet.sOS = sOS
                    objDet.sProduto = sProduto
                    objDet.sProdutoOS = sServico
                    objDet.sProdutoDesc = sDescricao
                    objDet.sUM = sUM
                    objDet.sTipo = sTipo(iIndice)
                    objDet.sDet = sDet
                    objDet.iClasseUM = iClasseUM
                    objDet.iMarcado = MARCADO
                    If iNatureza = NATUREZA_PROD_SERVICO Then
                        objDet.iPeca = DESMARCADO
                    Else
                        objDet.iPeca = MARCADO
                    End If
                    Set objDet.objItemSol = objItem
                    Set objDet.objSol = objItem.gobjInfoUsu
                    colItens.Add objDet
                    
                End If
            
                lErro = Comando_BuscarProximo(alComando(iIndice))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 206425
            
            Loop
        
        Next
        
    Next
    
   'fechar comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Garantia_Le_Gastos = SUCESSO
    
    Exit Function
    
Erro_Garantia_Le_Gastos:

    Garantia_Le_Gastos = gErr
     
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
     
        Case 206422
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 206423 To 206425
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_OS", gErr)
     
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206426)
     
    End Select
     
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Private Function ItemAP_Testa_Qtd_Falta_Faturar(ByVal lComando As Long, ByVal lNumIntDocAP As Long, ByVal iTipo As Integer, ByVal iClasseUMAP As Integer, ByVal sUMAP As String, ByVal dQuantidadeAp As Double, dQtdAFat As Double)

Dim lErro As Long
Dim sSQL As String
Dim sUM As String
Dim dQuantidade As Double
Dim dQtdFalta As Double
Dim dFator As Double

On Error GoTo Erro_ItemAP_Testa_Qtd_Falta_Faturar

    ''PARA TESTE
    'dQtdAFat = dQuantidadeAp
    
    'OFICIAL
    dQtdFalta = dQuantidadeAp
    sSQL = "SELECT Quantidade, UM FROM ItemNFItemAPSRV WHERE TipoNumIntDocAP = ? AND NumIntDocItemAP = ?"
    sUM = String(STRING_UM_SIGLA, 0)
    lErro = Comando_Executar(lComando, sSQL, dQuantidade, sUM, iTipo, lNumIntDocAP)
    If lErro <> AD_SQL_SUCESSO Then gError 206428
    
    'Posiciona no primeiro registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 206429
    
    Do While lErro <> AD_SQL_SEM_DADOS
    
        lErro = CF("UM_Conversao_Trans", iClasseUMAP, sUM, sUMAP, dFator)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        dQtdFalta = dQtdFalta - dQuantidade
    
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 206430
    
    Loop
    
    If dQtdFalta < 0 Then dQtdFalta = 0
    
    dQtdAFat = dQtdFalta
    
    ItemAP_Testa_Qtd_Falta_Faturar = SUCESSO
    
    Exit Function
    
Erro_ItemAP_Testa_Qtd_Falta_Faturar:

    ItemAP_Testa_Qtd_Falta_Faturar = gErr
     
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case 206428 To 206430
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMNFITEMAPSRV", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206431)
     
    End Select
    
    Exit Function
    
End Function

Public Function NFiscal_Le_Apontamentos(ByVal objNF As ClassNFiscal, ByVal colAPDet As Collection) As Long

Dim lErro As Long
Dim alComando(0 To 3) As Long
Dim iIndice As Integer, bAchou As Boolean
Dim objDet As ClassOSApDet
Dim objItemNF As ClassItemNF
Dim iTipoNumIntDocAP As Integer, lNumIntDocItemAP As Long, dQuantidade As Double
Dim sUM As String, lNumIntDocAP As Long, sDet As String, sOS As String
Dim lCodigoAp As Long, iSeq As Integer, sServico As String, sProduto As String, sDescricao As String
Dim sSQL(0 To 3) As String
Dim sTipo(1 To 3) As String
Dim colItens As New Collection
Dim objSolBD As New ClassSolicSRV
Dim objItemSol As ClassItensSolicSRV
Dim lCodigoSol As Long, sProdutoOS As String, sLoteOS As String, iFilialOP As Integer

On Error GoTo Erro_NFiscal_Le_Apontamentos
    
    'Abre comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 206438
    Next
    
    'Le os dados do apontamento
    'Le os produtos
    sSQL(1) = "SELECT OS.CodSolSRV,OS.Produto,OS.Lote,OS.FilialOP, '' As Det, OS.Codigo, A.Codigo, M.Item, I.Servico, P.Codigo, P.Descricao "
    sSQL(1) = sSQL(1) & "FROM OS, ItemOS AS I, MovimentoEstoque AS M, Produtos As P, ApontamentoSRV AS A "
    sSQL(1) = sSQL(1) & "WHERE OS.Codigo = I.Codigo AND OS.FilialEmpresa = I.FilialEmpresa AND M.NumIntDocOrigem = I.NumIntDoc AND M.TipoNumIntDocOrigem = 8 AND "
    sSQL(1) = sSQL(1) & "P.Codigo = M.Produto AND A.CodigoMovEst = M.Codigo AND M.NumIntDoc = ? "
    sTipo(1) = "Peças"
    'Le as máquinas
    sSQL(2) = "SELECT OS.CodSolSRV,OS.Produto,OS.Lote,OS.FilialOP, CONVERT(VARCHAR, Maq.Codigo) + '-' + Maq.NomeReduzido As Det, OS.Codigo, M.CodigoAP, M.Seq, M.ProdutoOS, P.Codigo,  P.Descricao "
    sSQL(2) = sSQL(2) & "FROM OS, ItemOS AS I, OSApMaq AS M, Maquinas AS Maq, Produtos As P, ApontamentoSRV AS A "
    sSQL(2) = sSQL(2) & "WHERE OS.Codigo = i.Codigo And OS.FilialEmpresa = i.FilialEmpresa AND M.CodigoAP = A.Codigo AND M.NumIntDocMaq = Maq.NumIntDoc AND M.OS = OS.Codigo AND M.ProdutoOS = I.Servico AND "
    sSQL(2) = sSQL(2) & "P.Codigo = Maq.Produto AND M.NumIntDoc = ? "
    sTipo(2) = "Máquinas"
    'Le a mão-de-obra
    sSQL(3) = "SELECT OS.CodSolSRV, OS.Produto, OS.Lote, OS.FilialOP, CONVERT(VARCHAR, MO.Codigo) + '-' + MO.NomeReduzido As Det, OS.Codigo, M.CodigoAP, M.Seq, M.ProdutoOS, P.Codigo,  P.Descricao "
    sSQL(3) = sSQL(3) & "FROM OS, ItemOS AS I, OSApMO AS M, MaoDeObra AS MO, Produtos As P, TiposDeMaoDeObra AS T, ApontamentoSRV AS A "
    sSQL(3) = sSQL(3) & "WHERE OS.Codigo = I.Codigo AND OS.FilialEmpresa = I.FilialEmpresa AND M.CodigoAP = A.Codigo AND M.CodigoMO = MO.Codigo AND M.OS = OS.Codigo AND M.ProdutoOS = I.Servico AND "
    sSQL(3) = sSQL(3) & "MO.Tipo = T.Codigo AND P.Codigo = T.Produto AND M.NumIntDoc = ? "
    sTipo(3) = "Mão-de-obra"
    
    For Each objItemNF In objNF.colItensNF
        colItens.Add objItemNF
    Next
    For Each objItemNF In objNF.colItensSRV
        colItens.Add objItemNF
    Next
    
    iIndice = 0
    For Each objItemNF In colItens
    
        iIndice = iIndice + 1
    
        sUM = String(STRING_UM_SIGLA, 0)
    
        lErro = Comando_Executar(alComando(0), "SELECT TipoNumIntDocAP,NumIntDocItemAP,Quantidade,UM, NumIntDocAP FROM ItemNFItemAPSRV WHERE NumIntItemNF = ?", iTipoNumIntDocAP, lNumIntDocItemAP, dQuantidade, sUM, lNumIntDocAP, objItemNF.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 206439
        
        'Posiciona no primeiro registro
        lErro = Comando_BuscarPrimeiro(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 206440
        
        Do While lErro = AD_SQL_SUCESSO
        
            Set objDet = New ClassOSApDet
            colAPDet.Add objDet
        
            sServico = String(STRING_PRODUTO, 0)
            sProduto = String(STRING_PRODUTO, 0)
            sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
            sOS = String(STRING_OS_CODIGO, 0)
            sDet = String(255, 0)
            sProdutoOS = String(STRING_PRODUTO, 0)
            sLoteOS = String(STRING_LOTE_RASTREAMENTO, 0)
                    
            'Busca Peças\Mão-de-obra\Máquinas apontadas
            lErro = Comando_Executar(alComando(iTipoNumIntDocAP), sSQL(iTipoNumIntDocAP), lCodigoSol, sProdutoOS, sLoteOS, iFilialOP, sDet, sOS, lCodigoAp, iSeq, sServico, sProduto, sDescricao, lNumIntDocItemAP)
            If lErro <> AD_SQL_SUCESSO Then gError 206441
        
            lErro = Comando_BuscarPrimeiro(alComando(iTipoNumIntDocAP))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 206442
        
            objDet.iTipo = iTipoNumIntDocAP
            objDet.dQuantidade = dQuantidade
            objDet.iFilialEmpresa = objNF.iFilialEmpresa
            objDet.lCodigoAp = lCodigoAp
            objDet.lNumIntDoc = iTipoNumIntDocAP
            objDet.lNumIntDocAP = lNumIntDocAP
            objDet.sOS = sOS
            objDet.sProduto = sProduto
            objDet.sProdutoOS = sServico
            objDet.sProdutoDesc = sDescricao
            objDet.sUM = sUM
            objDet.sTipo = sTipo(iTipoNumIntDocAP)
            objDet.sDet = sDet
            objDet.iMarcado = MARCADO
            If iIndice >= objNF.colItensNF.Count Then
                objDet.iPeca = DESMARCADO
            Else
                objDet.iPeca = MARCADO
            End If
            Set objSolBD = New ClassSolicSRV
            'Dim objItemSol As ClassItensSolicSRV
            objSolBD.lCodigo = lCodigoSol
            objSolBD.iFilialEmpresa = objNF.iFilialEmpresa
            lErro = CF("SolicitacaoSRV_Le", objSolBD)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            Set objDet.objSol = objSolBD
            bAchou = False
            For Each objItemSol In objSolBD.colItens
                If objItemSol.sProduto = sProdutoOS And objItemSol.sLote = sLoteOS And objItemSol.iFilialOP = iFilialOP Then
                    Set objDet.objItemSol = objItemSol
                    bAchou = True
                    Exit For
                End If
            Next
            If Not bAchou Then gError 206443
            
            'Posiciona no próximo registro
            lErro = Comando_BuscarProximo(alComando(0))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 206444

        Loop
        
    Next

    'Fecha comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    NFiscal_Le_Apontamentos = SUCESSO
    
    Exit Function
    
Erro_NFiscal_Le_Apontamentos:

    NFiscal_Le_Apontamentos = gErr
    
    Select Case gErr
    
        Case 206422
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 206439, 206440
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMNFITEMAPSRV", gErr)
        
        Case 206441, 206442, 206444
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_OS", gErr)
        
        Case 206443
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEMSOLSRV_NAO_CADASTRADO", gErr, lCodigoSol, sProdutoOS, sLoteOS, iFilialOP)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206445)
    
    End Select

    'Fecha comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Function BaixaPedidosSRV_ObterPedidos(ByVal objBaixaPedido As ClassBaixaPedidos) As Long
'Preenche a colecao colPVInfo em objBaixaPedido com os pedidos que podem ser faturados e atendem
'aos criterios passados em objBaixaPedido com restricoes de No. de pedido, data de emissao,...

Dim lErro As Long, sSelect As String
Dim lComando As Long

Dim vdtItensPedidoDeVendaDataEntrega As Variant, vsClientesNomeReduzido As Variant, vlPedidosDeVendaCodigo As Variant, vlPedidosDeVendaCliente As Variant
Dim viPedidosDeVendaFilial As Variant, vdtPedidosDeVendaDataEmissao As Variant, vdPedidosDeVendaValorTotal As Variant

On Error GoTo Erro_BaixaPedidosSRV_ObterPedidos

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 210650

    'Montar SELECT dos pedidos de venda dinamicamente
    Call BaixaPedidosSRV_ObterPedidos1(objBaixaPedido, sSelect)

    vsClientesNomeReduzido = String(255, 0)
    vdtItensPedidoDeVendaDataEntrega = DATA_NULA
    vlPedidosDeVendaCodigo = CLng(0)
    vlPedidosDeVendaCliente = CLng(0)
    viPedidosDeVendaFilial = CInt(0)
    vdtPedidosDeVendaDataEmissao = DATA_NULA
    vdPedidosDeVendaValorTotal = CDbl(0)
    
    'Preparar parte fixa do SELECT
    lErro = CF("BaixaPedidosSRV_ObterPedidos2", lComando, sSelect, vdtItensPedidoDeVendaDataEntrega, vsClientesNomeReduzido, vlPedidosDeVendaCodigo, vlPedidosDeVendaCliente, viPedidosDeVendaFilial, vdtPedidosDeVendaDataEmissao, vdPedidosDeVendaValorTotal)
    If lErro <> SUCESSO Then gError 210651

    'Complementa a passagem dos parametros que variam de acordo com a selecao do usuario
    lErro = CF("BaixaPedidosSRV_ObterPedidos3", lComando, objBaixaPedido)
    If lErro <> SUCESSO Then gError 210652

    'Processar os registros lidos
    lErro = CF("BaixaPedidosSRV_ObterPedidos4", lComando, objBaixaPedido, vdtItensPedidoDeVendaDataEntrega, vsClientesNomeReduzido, vlPedidosDeVendaCodigo, vlPedidosDeVendaCliente, viPedidosDeVendaFilial, vdtPedidosDeVendaDataEmissao, vdPedidosDeVendaValorTotal)
    If lErro <> SUCESSO And lErro <> 27163 Then gError 210653
    
    'Se não encontrou nenhum Pedido ---> Erro
    If lErro = 27163 Then gError 210654
    
    lErro = Comando_Fechar(lComando)

    BaixaPedidosSRV_ObterPedidos = SUCESSO

    Exit Function

Erro_BaixaPedidosSRV_ObterPedidos:

    BaixaPedidosSRV_ObterPedidos = gErr

    Select Case gErr

        Case 210650
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 210651 To 210654

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 210655)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Sub BaixaPedidosSRV_ObterPedidos1(ByVal objBaixaPedido As ClassBaixaPedidos, sSelect As String)
'Montar SELECT dinamicamente

Dim sFields As String, sFrom As String, sWhere As String, sOrderBy As String, sGroupBy As String, sHaving As String

    sFields = "MIN(ItensPedidoSRV.DataEntrega), MIN(Clientes.NomeReduzido), PedidoServico.Codigo, MIN(PedidoServico.Cliente), MIN(PedidoServico.Filial), MIN(PedidoServico.DataEmissao), MIN(PedidoServico.ValorTotal)"
    sFrom = " FROM PedidoServico, Clientes, ItensPedidoSRV, FiliaisClientes"
    sWhere = " WHERE PedidoServico.FilialEmpresa = ? AND PedidoServico.Cliente = Clientes.Codigo AND PedidoServico.Cliente = FiliaisClientes.CodCliente AND PedidoServico.Filial = FiliaisClientes.CodFilial AND PedidoServico.FilialEmpresa=ItensPedidoSRV.FilialEmpresa AND " & _
        "PedidoServico.Codigo=ItensPedidoSRV.CodPedido AND PedidoServico.Status <> 2 "
    sOrderBy = " ORDER BY " & objBaixaPedido.sOrdenacao
    sGroupBy = " GROUP BY " & "PedidoServico.Codigo"
    sHaving = ""

    If objBaixaPedido.lPedidosDe <> 0 Then sWhere = sWhere & " AND PedidoServico.Codigo >= ?"

    If objBaixaPedido.lPedidosAte <> 0 Then sWhere = sWhere & " AND PedidoServico.Codigo <= ?"

    If objBaixaPedido.lClientesDe <> 0 Then sWhere = sWhere & " AND PedidoServico.Cliente >= ?"

    If objBaixaPedido.lClientesAte <> 0 Then sWhere = sWhere & " AND PedidoServico.Cliente <= ?"

    If objBaixaPedido.dtEmissaoDe <> DATA_NULA Then sWhere = sWhere & " AND PedidoServico.DataEmissao >= ?"

    If objBaixaPedido.dtEmissaoAte <> DATA_NULA Then sWhere = sWhere & " AND PedidoServico.DataEmissao <= ?"

    If objBaixaPedido.dtEntregaDe <> DATA_NULA Then sHaving = " HAVING MIN(ItensPedidoSRV.DataEntrega) >= ?"

    If objBaixaPedido.dtEntregaAte <> DATA_NULA Then
        If sHaving = "" Then
            sHaving = " HAVING "
        Else
            sHaving = sHaving & " AND "
        End If

        sHaving = sHaving & "MIN(ItensPedidoSRV.DataEntrega) <= ?"
    End If

    sSelect = "SELECT " & sFields & sFrom & sWhere & sGroupBy & sHaving & sOrderBy

End Sub

Function BaixaPedidosSRV_ObterPedidos2(ByVal lComando As Long, ByVal sSelect As String, vItensPedidoDeVendaDataEntrega As Variant, vClientesNomeReduzido As Variant, vPedidosDeVendaCodigo As Variant, vPedidosDeVendaCliente As Variant, vPedidosDeVendaFilial As Variant, vPedidosDeVendaDataEmissao As Variant, vPedidosDeVendaValorTotal As Variant) As Long
'Preparar parte fixa do SELECT

Dim lErro As Long

On Error GoTo Erro_BaixaPedidosSRV_ObterPedidos2
    
    lErro = Comando_PrepararInt(lComando, sSelect)
    If (lErro <> AD_SQL_SUCESSO) Then gError 210656

    lErro = Comando_BindVarInt(lComando, vItensPedidoDeVendaDataEntrega)
    If (lErro <> AD_SQL_SUCESSO) Then gError 210657

    lErro = Comando_BindVarInt(lComando, vClientesNomeReduzido)
    If (lErro <> AD_SQL_SUCESSO) Then gError 210658

    lErro = Comando_BindVarInt(lComando, vPedidosDeVendaCodigo)
    If (lErro <> AD_SQL_SUCESSO) Then gError 210659

    lErro = Comando_BindVarInt(lComando, vPedidosDeVendaCliente)
    If (lErro <> AD_SQL_SUCESSO) Then gError 210660

    lErro = Comando_BindVarInt(lComando, vPedidosDeVendaFilial)
    If (lErro <> AD_SQL_SUCESSO) Then gError 210661

    lErro = Comando_BindVarInt(lComando, vPedidosDeVendaDataEmissao)
    If (lErro <> AD_SQL_SUCESSO) Then gError 210662
    
    lErro = Comando_BindVarInt(lComando, vPedidosDeVendaValorTotal)
    If (lErro <> AD_SQL_SUCESSO) Then gError 210663

    BaixaPedidosSRV_ObterPedidos2 = SUCESSO

    Exit Function

Erro_BaixaPedidosSRV_ObterPedidos2:

    BaixaPedidosSRV_ObterPedidos2 = gErr

    Select Case gErr

        Case 210656 To 210663
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOS_VENDA_BAIXA_PV", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 210664)

    End Select

    Exit Function

End Function

Function BaixaPedidosSRV_ObterPedidos3(ByVal lComando As Long, ByVal objBaixaPedido As ClassBaixaPedidos) As Long
'Filtra por giFilialEmpresa
'Complementa a passagem dos parametros que variam de acordo com a selecao do usuario

Dim lErro As Long
Dim vFilialEmpresa As Variant, vlPedidosDe As Variant, vlPedidosAte As Variant, vlClientesDe As Variant, vlClientesAte As Variant, viFilialPedidoDe As Variant, viFilialPedidoAte As Variant, vdtEmissaoDe As Variant, vdtEmissaoAte As Variant, vdtEntregaDe As Variant, vdtEntregaAte As Variant

On Error GoTo Erro_BaixaPedidosSRV_ObterPedidos3

    vFilialEmpresa = giFilialEmpresa
    lErro = Comando_BindVarInt(lComando, vFilialEmpresa)
    If (lErro <> AD_SQL_SUCESSO) Then gError 210665

    If objBaixaPedido.lPedidosDe <> 0 Then
        vlPedidosDe = objBaixaPedido.lPedidosDe
        lErro = Comando_BindVarInt(lComando, vlPedidosDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError 210666
    End If

    If objBaixaPedido.lPedidosAte <> 0 Then
        vlPedidosAte = objBaixaPedido.lPedidosAte
        lErro = Comando_BindVarInt(lComando, vlPedidosAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError 210667
    End If

    If objBaixaPedido.lClientesDe <> 0 Then
        vlClientesDe = objBaixaPedido.lClientesDe
        lErro = Comando_BindVarInt(lComando, vlClientesDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError 210668
    End If

    If objBaixaPedido.lClientesAte <> 0 Then
        vlClientesAte = objBaixaPedido.lClientesAte
        lErro = Comando_BindVarInt(lComando, vlClientesAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError 210669
    End If

    If objBaixaPedido.dtEmissaoDe <> DATA_NULA Then
        vdtEmissaoDe = objBaixaPedido.dtEmissaoDe
        lErro = Comando_BindVarInt(lComando, vdtEmissaoDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError 210670
    End If

    If objBaixaPedido.dtEmissaoAte <> DATA_NULA Then
        vdtEmissaoAte = objBaixaPedido.dtEmissaoAte
        lErro = Comando_BindVarInt(lComando, vdtEmissaoAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError 210671
    End If

    If objBaixaPedido.dtEntregaDe <> DATA_NULA Then
        vdtEntregaDe = objBaixaPedido.dtEntregaDe
        lErro = Comando_BindVarInt(lComando, vdtEntregaDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError 210672
    End If

    If objBaixaPedido.dtEntregaAte <> DATA_NULA Then
        vdtEntregaAte = objBaixaPedido.dtEntregaAte
        lErro = Comando_BindVarInt(lComando, vdtEntregaAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError 210673
    End If
    
    lErro = Comando_ExecutarInt(lComando)
    If (lErro <> AD_SQL_SUCESSO) Then gError 210674

    BaixaPedidosSRV_ObterPedidos3 = SUCESSO

    Exit Function

Erro_BaixaPedidosSRV_ObterPedidos3:

    BaixaPedidosSRV_ObterPedidos3 = gErr

    Select Case gErr

        Case 210665 To 210674
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOS_VENDA_BAIXA_PV", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 210675)

    End Select

    Exit Function

End Function

Function BaixaPedidosSRV_ObterPedidos4(ByVal lComando As Long, ByVal objBaixaPedido As ClassBaixaPedidos, vdtItensPedidoDeVendaDataEntrega As Variant, vsClientesNomeReduzido As Variant, vlPedidosDeVendaCodigo As Variant, vlPedidosDeVendaCliente As Variant, viPedidosDeVendaFilial As Variant, vdtPedidosDeVendaDataEmissao As Variant, vdPedidosDeVendaValorTotal As Variant) As Long
'Processar os registros lidos preenchendo a colecao

Dim lErro As Long, objPVInfo As ClassPVInfo

On Error GoTo Erro_BaixaPedidosSRV_ObterPedidos4

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 210676

    If lErro = AD_SQL_SEM_DADOS Then gError 210677

    Do While lErro = AD_SQL_SUCESSO

        Set objPVInfo = New ClassPVInfo

        objPVInfo.lCodPedido = vlPedidosDeVendaCodigo
        objPVInfo.lCliente = vlPedidosDeVendaCliente
        objPVInfo.sClienteNomeReduzido = vsClientesNomeReduzido
        objPVInfo.iFilialCliente = viPedidosDeVendaFilial
        objPVInfo.dtEmissao = vdtPedidosDeVendaDataEmissao
        objPVInfo.dtEntrega = vdtItensPedidoDeVendaDataEntrega
        objPVInfo.iMarcada = DESMARCADO
        objPVInfo.dValorTotal = vdPedidosDeVendaValorTotal
    
        objBaixaPedido.colPVInfo.Add objPVInfo

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 210678

    Loop

    BaixaPedidosSRV_ObterPedidos4 = SUCESSO

    Exit Function

Erro_BaixaPedidosSRV_ObterPedidos4:

    BaixaPedidosSRV_ObterPedidos4 = gErr

    Select Case gErr
        
        Case 210676, 210677
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOS_VENDA_GERACAO_NF", gErr)

        Case 210678
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SEM_PEDIDOS_VENDA_ENCONTRADOS", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 210679)

    End Select

    Exit Function

End Function

Public Function CamposGenericos_ValidaExclusao_TipoSS() As Long
'Verifica se é válida a exclusão de valores para o campo Requisitante

Dim lComando As Long
Dim lErro As Long
Dim lTipo As Long

On Error GoTo Erro_CamposGenericos_ValidaExclusao_TipoSS
    
    'Abre comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 141377
    
    'Busca Tipos de Paradas em CTMaquinasParadas que não existam na tabela CamposGenericos Valores
    lErro = Comando_Executar(lComando, "SELECT Tipo FROM SolicitacaoSRV WHERE Tipo NOT IN (SELECT CodValor FROM CamposGenericosValores WHERE CodCampo = ?) AND Tipo <> 0", lTipo, CAMPOSGENERICOS_TIPOSS)
    If lErro <> AD_SQL_SUCESSO Then gError 141378
    
    'Posiciona no primeiro registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 141379
    
    'Se encontrou algum registro => erro de inconsistência
    If lErro = AD_SQL_SUCESSO Then gError 141380

    'Fecha comando
    Call Comando_Fechar(lComando)
    
    CamposGenericos_ValidaExclusao_TipoSS = SUCESSO
    
    Exit Function
    
Erro_CamposGenericos_ValidaExclusao_TipoSS:

    CamposGenericos_ValidaExclusao_TipoSS = gErr
    
    Select Case gErr
    
        Case 141377
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 141378, 141379
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_OS", gErr)
        
        Case 141380
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CAMPOSGENERICOSVALORES", gErr, lTipo, CAMPOSGENERICOS_TIPOSS)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150281)
    
    End Select

    'Fecha comando
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function CamposGenericos_ValidaExclusao_FaseSS() As Long
'Verifica se é válida a exclusão de valores para o campo Requisitante

Dim lComando As Long
Dim lErro As Long
Dim lFase As Long

On Error GoTo Erro_CamposGenericos_ValidaExclusao_FaseSS
    
    'Abre comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 141377
    
    'Busca Tipos de Paradas em CTMaquinasParadas que não existam na tabela CamposGenericos Valores
    lErro = Comando_Executar(lComando, "SELECT Fase FROM SolicitacaoSRV WHERE Fase NOT IN (SELECT CodValor FROM CamposGenericosValores WHERE CodCampo = ?) AND Fase <> 0", lFase, CAMPOSGENERICOS_FASESS)
    If lErro <> AD_SQL_SUCESSO Then gError 141378
    
    'Posiciona no primeiro registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 141379
    
    'Se encontrou algum registro => erro de inconsistência
    If lErro = AD_SQL_SUCESSO Then gError 141380

    'Fecha comando
    Call Comando_Fechar(lComando)
    
    CamposGenericos_ValidaExclusao_FaseSS = SUCESSO
    
    Exit Function
    
Erro_CamposGenericos_ValidaExclusao_FaseSS:

    CamposGenericos_ValidaExclusao_FaseSS = gErr
    
    Select Case gErr
    
        Case 141377
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 141378, 141379
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_OS", gErr)
        
        Case 141380
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CAMPOSGENERICOSVALORES", gErr, lFase, CAMPOSGENERICOS_FASESS)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150281)
    
    End Select

    'Fecha comando
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

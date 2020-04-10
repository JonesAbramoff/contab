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

Private Declare Function Comando_BindVarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_BindVar" (ByVal lComando As Long, lpVar As Variant) As Long
Private Declare Function Comando_PrepararInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Preparar" (ByVal lComando As Long, ByVal lpSQLStmt As String) As Long
Private Declare Function Comando_ExecutarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Executar" (ByVal lComando As Long) As Long

Public Function Importa_Xml_NFe_Cust() As Long

Dim lErro As Long, iIndice As Integer
Dim lTransacao As Long
Dim alComando(0 To 10) As Long
Dim objNF As New ClassNFiscal
Dim objItemNF As New ClassItemNF
Dim objTipoDocInfo As ClassTipoDocInfo
Dim lForn As Long, iFilForn As Integer
Dim lNumNF As Long, sSerie As String, dtDataEmi As Date, sMsg(1 To 4) As String
Dim sMensagem As String, dTotal As Double, sChvNFe As String, sCFOP As String
Dim iItem As Integer, sProdXML As String, sProdDesc As String, sCFOPInterno As String
Dim sCFOPXML As String, sUMXML As String, dQtd As Double, dPreco As Double
Dim objProduto As ClassProduto, dTotalItens As Integer, sProduto As String
Dim objTribTab As New ClassTribTab, sNFs As String, sEANXML As String

Const CGC_EMITENTE = "20371515000105" '"07495653000219" ' "07095653000219" '"04713201000197"

On Error GoTo Erro_Importa_Xml_NFe_Cust

    'Abrir transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 209866

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 209867
    Next
    
    'Busca o código do fornecedor
    lErro = Comando_Executar(alComando(0), "SELECT CodFornecedor, CodFilial FROM FiliaisFornecedores WHERE CGC = ?", lForn, iFilForn, CGC_EMITENTE)
    If lErro <> AD_SQL_SUCESSO Then gError 209868
    
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209869
    
    If lErro <> AD_SQL_SUCESSO Then gError 209870 ' Fornecedor não cadastrado
       
    sSerie = String(STRING_MAXIMO, 0)
    sChvNFe = String(STRING_MAXIMO, 0)
    For iIndice = 1 To 4
        sMsg(iIndice) = String(STRING_MAXIMO, 0)
    Next
    
    'Só traz as NFe de um determinado fornecedor para filial empresa corrente que
    'não tenham sido importadas anteriormente e que não possuam
    'nem tributação e nem nenhum outro valor que afete o total da NF que não seja os itens
'    lErro = Comando_ExecutarPos(alComando(1), "SELECT ChvNFe, NumNF, Serie, DataEmissao, ValorTotalNF, SUBSTRING(Mensagem,1,250), SUBSTRING(Mensagem,251,250), SUBSTRING(Mensagem,501,250), SUBSTRING(Mensagem,751,250) FROM ImportNFeXml WHERE NumIntNF = 0 AND CGCEmitente = ? AND FilialEmpresa = ? AND ABS(ValorProdutos - ValorTotalNF) < 0.001 AND ICMSBase = 0 ", 0, _
'                                                sChvNFe, lNumNF, sSerie, dtDataEmi, dTotal, sMsg(1), sMsg(2), sMsg(3), sMsg(4), CGC_EMITENTE, giFilialEmpresa)
'    If lErro <> AD_SQL_SUCESSO Then gError 209873
    lErro = Comando_Executar(alComando(1), "SELECT ChvNFe, NumNF, Serie, DataEmissao, ValorTotalNF, SUBSTRING(Mensagem,1,250), SUBSTRING(Mensagem,251,250), SUBSTRING(Mensagem,501,250), SUBSTRING(Mensagem,751,250) FROM ImportNFeXml WHERE NumIntNF = 0 AND CGCEmitente = ? AND FilialEmpresa = ? AND ABS(ValorProdutos - ValorTotalNF) < 0.001 AND ICMSBase = 0 ", _
                                                sChvNFe, lNumNF, sSerie, dtDataEmi, dTotal, sMsg(1), sMsg(2), sMsg(3), sMsg(4), CGC_EMITENTE, giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 209873
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209874
    
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objNF = New ClassNFiscal
    
        'Inicializa a parte tributária
        Call objNF.Inicializa_Tributacao
        Call objNF.objTributacaoNF.Coloca_Auto
    
        objNF.lNumNotaFiscal = lNumNF
        objNF.sSerie = sSerie & "-e"
        objNF.dtDataEmissao = dtDataEmi
        objNF.iNFe = MARCADO
        For iIndice = 1 To 4
            objNF.sMensagemNota = objNF.sMensagemNota & sMsg(iIndice)
        Next
        objNF.iFilialEmpresa = giFilialEmpresa
        objNF.iTipoNFiscal = DOCINFO_NFERS 'NF externa de entrada remessa simbólica
        objNF.iTipoDocInfo = TIPODOCINFO_TIPO_NFEXT
        objNF.iNaoVerificaCredito = NAO_VERIFICA_CREDITO_CLIENTE
        objNF.iStatus = STATUS_BAIXADO
        objNF.dtDataReferencia = DATA_NULA
        objNF.iFreteRespons = FRETE_EMITENTE
        objNF.dtDataRegistro = Date
        objNF.dtDataSaida = DATA_NULA
        objNF.dtDataVencimento = DATA_NULA
        objNF.lFornecedor = lForn
        objNF.iFilialForn = iFilForn
        objNF.dtDataEntrada = objNF.dtDataEmissao 'Date
        
        objNF.dValorTotal = dTotal
        objNF.dValorProdutos = dTotal
        objNF.sChvNFe = sChvNFe
        
        sProdXML = String(STRING_MAXIMO, 0)
        sProdDesc = String(STRING_MAXIMO, 0)
        sCFOPXML = String(STRING_MAXIMO, 0)
        sUMXML = String(STRING_MAXIMO, 0)
        sEANXML = String(STRING_MAXIMO, 0)
    
        'Lê os itens da NF
        lErro = Comando_Executar(alComando(2), "SELECT Item, Produto, SUBSTRING(DescProd,1,250), CFOP, UM, Quantidade, PrecoUnitario, EAN FROM ImportNFeItensXml WHERE ChvNFE = ? ORDER BY Item", _
            iItem, sProdXML, sProdDesc, sCFOPXML, sUMXML, dQtd, dPreco, sEANXML, sChvNFe)
        If lErro <> AD_SQL_SUCESSO Then gError 209875
        
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209876
        
        sCFOP = ""
        sCFOPInterno = ""
        'Para cada item
        Do While lErro <> AD_SQL_SEM_DADOS
        
            Set objProduto = New ClassProduto
            Set objItemNF = New ClassItemNF
        
            'Inicializa parte tributária
            Call objItemNF.Inicializa_Tributacao
            Call objItemNF.objTributacaoItemNF.Coloca_Auto
            
            'Obtém dados fixos
            objItemNF.dtDataEntrega = DATA_NULA
        
            'Obtém dados do Registro
            objItemNF.iItem = iItem
            
            sCFOP = Replace(sCFOPXML, "Item", "")
            Select Case sCFOP
                Case "5902"
                    sCFOPInterno = "1902"
                Case "5906", "5907"
                    sCFOPInterno = "1907"
                Case "5905"
                    sCFOPInterno = "1905"
                Case Else
                    sCFOPInterno = "1949"
            End Select
            
            lErro = CF("Importa_Xml_NFe_Loc_Prod", giFilialEmpresa, lForn, iFilForn, sProdXML, sEANXML, objProduto)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
            objItemNF.sProduto = objProduto.sCodigo
            objItemNF.sDescricaoItem = sProdDesc
            objItemNF.dQuantidade = dQtd
            objItemNF.dPrecoUnitario = dPreco
            
            objItemNF.sUnidadeMed = objProduto.sSiglaUMCompra
            objItemNF.sUMEstoque = objProduto.sSiglaUMEstoque
            objItemNF.sUMVenda = objProduto.sSiglaUMVenda
            objItemNF.iClasseUM = objProduto.iClasseUM
            
            objNF.colItensNF.Add1 objItemNF
            
            objItemNF.objTributacaoItemNF.iTipoTributacao = 33
            objItemNF.objTributacaoItemNF.iNaturezaOpManual = VAR_PREENCH_MANUAL
            objItemNF.objTributacaoItemNF.sNaturezaOp = sCFOPInterno
            objItemNF.objTributacaoItemNF.iNaturezaOpManual = VAR_PREENCH_MANUAL
            
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
        
            lErro = Comando_BuscarProximo(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209877
        
        Loop
        
        objNF.sNaturezaOp = sCFOP
               
        objNF.objTributacaoNF.iTipoTributacao = 33
        objNF.objTributacaoNF.iNaturezaOpManual = VAR_PREENCH_MANUAL
        objNF.objTributacaoNF.sNaturezaOpInterna = sCFOPInterno
        objNF.objTributacaoNF.iTipoTributacaoManual = VAR_PREENCH_MANUAL
        
        'Inicializa o rastreamento
        Set objNF.objRastreamento = New ClassRastreamento
        
        'GRAVA A NF
        '****************************************************************************
        'Verifica se a existe nota fiscal está cadastrada
        lErro = CF("NFiscal_Le_1", objNF)
        If lErro <> SUCESSO And lErro <> 83971 Then gError ERRO_SEM_MENSAGEM
        
        'Se ela não está cadastrada
        If lErro <> SUCESSO Then

            Set objTipoDocInfo = New ClassTipoDocInfo
            
            objTipoDocInfo.iCodigo = objNF.iTipoNFiscal
            
            'Lê o Tipo de Documento
            lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
            If lErro <> SUCESSO And lErro <> 31415 Then gError ERRO_SEM_MENSAGEM
            
            'pesquisa a existencia da nota fiscal com os dados passados como parametro e não exibe mensagens
            lErro = CF("NFiscalEntrada_Verifica_Existencia2", objNF, objTipoDocInfo, True)
            If lErro <> SUCESSO And lErro <> 61414 And lErro <> 89723 Then gError ERRO_SEM_MENSAGEM
                          
            'Se for uma nota nova
            If lErro = SUCESSO Then
                
                Set objTribTab = New ClassTribTab
                
                lErro = objTribTab.TributacaoNF_Reset(objNF)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                'obtem dados do contrato
                Call objTribTab.TipoNFiscal_Definir(objTipoDocInfo.iCodigo, objTipoDocInfo.sSigla)
                Call objTribTab.DataEmissao_Alterada(objNF.dtDataEmissao)
                Call objTribTab.Serie_Alterada(objNF.sSerie)
                
                lErro = gobjTributacao.AtualizaImpostos(objTribTab.mvarobjNFTributacao, 0)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                lErro = objTribTab.PreencherNF(objNF)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                Set objTribTab = Nothing
                
                objNF.sNaturezaOp = objNF.objTributacaoNF.sNaturezaOpInterna
            
                'Grava a Nota em Transação
                lErro = CF("NFiscalEntradaRemessa_Grava_EmTrans", objNF, Nothing)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                If Len(Trim(sNFs)) <> 0 Then sNFs = sNFs & SEPARADOR
                sNFs = sNFs & CStr(objNF.lNumNotaFiscal)
                
            End If
        End If
        '****************************************************************************
        
'        lErro = Comando_ExecutarPos(alComando(3), "UPDATE ImportNFeXml SET NumIntNF = ?", alComando(1), objNF.lNumIntDoc)
'        If lErro <> AD_SQL_SUCESSO Then gError 209878
        lErro = Comando_Executar(alComando(3), "UPDATE ImportNFeXml SET NumIntNF = ? WHERE ChvNFe = ?", objNF.lNumIntDoc, sChvNFe)
        If lErro <> AD_SQL_SUCESSO Then gError 209878
    
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209879
    
    Loop
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Fechar transacao
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 209880
    
    Call Rotina_Aviso(vbOKOnly, "AVISO_NF_IMPORT_COM_SUCESSO", sNFs)

    Importa_Xml_NFe_Cust = SUCESSO

    Exit Function

Erro_Importa_Xml_NFe_Cust:

    Importa_Xml_NFe_Cust = gErr

    Select Case gErr

        Case 209866
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 209867
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 209868, 209869
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISFORNECEDORES", gErr)

        Case 209870
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORN_CGC_NAO_CADASTRADO", gErr, CGC_EMITENTE)
            
        Case 209871, 209872
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISEMPRESA", gErr)

        Case 209873, 209874, 209879
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_IMPORTNFEXML", gErr)

        Case 209875, 209876, 209877
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_IMPORTNFEXML", gErr)

        Case 209878
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_IMPORTNFEXML", gErr)

        Case 209880
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209881)

    End Select
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Call Transacao_Rollback

    Exit Function
    
End Function

Public Function RelOPGerEstTA_Prepara(ByVal dtDataDe As Date, ByVal dtDataAte As Date, ByVal sProdDe As String, ByVal sProdAte As String, ByVal iExibeProdSemMov As Integer, ByRef lNumIntRel As Long) As Long

Dim lErro As Long, iIndice As Integer
Dim lTransacao As Long, sSQL As String
Dim alComando(0 To 10) As Long, bConsidera As Boolean
Dim sTipoDoc As String, lNumIntDocItem As Long, lNumIntDoc As Long, lCodigo As Long
Dim iFilialEmpresa As Integer, lCliForn As Long, sNomeCliForn As String, dtDataPrev As Date
Dim sProduto As String, sDescricao As String, dQuantidade As Double, dPrecoUnitario As Double
Dim dValor As Double, sUM As String, dEstoqueInicial As Double, dQtdeAjuste As Double
Dim dQtdeAjusteAux As Double, bAchou As Boolean
Dim objEstoqueProduto As ClassEstoqueProduto, colEstoqueProduto As New Collection
Dim iSemana As Integer, dtDataIniSem As Date
Dim objEstoqueProduto1 As ClassEstoqueProduto, objEstoqueProduto2 As ClassEstoqueProduto

On Error GoTo Erro_RelOPGerEstTA_Prepara

    'Abrir transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 213214

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 213215
    Next
    
    sTipoDoc = String(STRING_MAXIMO, 0)
    sNomeCliForn = String(STRING_MAXIMO, 0)
    sProduto = String(STRING_MAXIMO, 0)
    sDescricao = String(STRING_MAXIMO, 0)
    sUM = String(STRING_MAXIMO, 0)
    
    'Obter NumIntRel
    lErro = CF("Config_ObterNumInt", "FATConfig", "NUMINTREL_RELGERESTTA", lNumIntRel)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = RelOPGerEstTA_PreparaSQL(sSQL, dtDataDe, dtDataAte, sProdDe, sProdAte)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = RelOPGerEstTA_PreparaInt(alComando(0), sSQL, giFilialEmpresa, dtDataDe, dtDataAte, sProdDe, sProdAte, sTipoDoc, lNumIntDocItem, lNumIntDoc, lCodigo, iFilialEmpresa, lCliForn, sNomeCliForn, dtDataPrev, sProduto, sDescricao, dQuantidade, dPrecoUnitario, dValor, sUM, dQtdeAjuste)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = Comando_BuscarProximo(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213216
    
    Do While lErro <> AD_SQL_SEM_DADOS
    
        bConsidera = True
        If dQtdeAjuste <> 0 And sTipoDoc = "PC" Then
            'Ajusta caso já tenha recebido parte do pedido de compra
            'Vai ignorar as primeiras entregas até ter saldo
            'se tiver saldo vai ajustar o valor e quantidade
            lErro = Comando_Executar(alComando(2), "SELECT SUM(Quantidade) FROM ItensPCEntrega WHERE NumIntItemPC = ? AND DataEntrega < ? ", dQtdeAjusteAux, dtDataPrev, lNumIntDocItem)
            If lErro <> AD_SQL_SUCESSO Then gError 213217
        
            lErro = Comando_BuscarPrimeiro(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213218
            
            If dQtdeAjuste > dQtdeAjusteAux Then
                dQtdeAjuste = dQtdeAjuste - dQtdeAjusteAux
            End If
            If dQtdeAjuste > dQuantidade Then
                bConsidera = False
            Else
                dValor = (dValor / dQuantidade) * dQtdeAjuste
                dQuantidade = dQtdeAjuste
            End If
        End If
        
        If bConsidera Then
        
            bAchou = False
            For Each objEstoqueProduto In colEstoqueProduto
                If objEstoqueProduto.sProduto = sProduto Then
                    bAchou = True
                    Exit For
                End If
            Next
            
            If Not bAchou Then
                Set objEstoqueProduto = New ClassEstoqueProduto
                Set objEstoqueProduto1 = New ClassEstoqueProduto
                Set objEstoqueProduto2 = New ClassEstoqueProduto
                objEstoqueProduto.sProduto = sProduto
            
                If giFilialEmpresa = EMPRESA_TODA Then
                    'Lê a soma de todas as quantidades para Produto Passado em todos os Almoxarifados da Empresa
                    lErro = CF("EstoqueProduto_Le_Todos_Almoxarifados", objEstoqueProduto)
                    If lErro <> SUCESSO And lErro <> 52979 Then gError ERRO_SEM_MENSAGEM
                Else
                    'Lê a soma de todas as quantidades para Produto Passado em todos os Almoxarifados da Filial
                    'lErro = CF("EstoqueProduto_Le_Todos_Almoxarifados_Filial", objEstoqueProduto, giFilialEmpresa)
                    'If lErro <> SUCESSO And lErro <> 64014 Then gError ERRO_SEM_MENSAGEM
                    
                    'A Márcia pediu para pegar só o saldo dos almoxarifados 1 e 2 e ignorar os outros
                    
                    objEstoqueProduto1.sProduto = sProduto
                    objEstoqueProduto1.iAlmoxarifado = 1
                    
                    lErro = CF("EstoqueProduto_Le", objEstoqueProduto1)
                    If lErro <> SUCESSO And lErro <> 21306 Then gError 33630

                    objEstoqueProduto2.sProduto = sProduto
                    objEstoqueProduto2.iAlmoxarifado = 2
                    
                    lErro = CF("EstoqueProduto_Le", objEstoqueProduto2)
                    If lErro <> SUCESSO And lErro <> 21306 Then gError 33630
                    
                    objEstoqueProduto.dQuantDispNossa = objEstoqueProduto1.dQuantDispNossa + objEstoqueProduto2.dQuantDispNossa
                    objEstoqueProduto.dQuantReservada = objEstoqueProduto1.dQuantReservada + objEstoqueProduto2.dQuantReservada
                    objEstoqueProduto.dQuantConsig3 = objEstoqueProduto1.dQuantConsig3 + objEstoqueProduto2.dQuantConsig3

                End If
                
                colEstoqueProduto.Add objEstoqueProduto
            End If
            
            Call CF("Rel12Semanas_Semana", dtDataPrev, iSemana, dtDataIniSem)
        
            lErro = Comando_Executar(alComando(1), "INSERT INTO RelOPGerEstTA (NumIntRel,Semana,DataIniSemana,DataFimSemana,TipoDoc,NumIntDocItem,NumIntDoc,Codigo,FilialEmpresa,CliForn,NomeCliForn,DataPrev,Produto,Descricao,Quantidade,PrecoUnitario,Valor,UM,EstoqueInicial ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) ", _
            lNumIntRel, iSemana, dtDataIniSem, DateAdd("d", 6, dtDataIniSem), sTipoDoc, lNumIntDocItem, lNumIntDoc, lCodigo, iFilialEmpresa, lCliForn, sNomeCliForn, dtDataPrev, sProduto, sDescricao, dQuantidade, dPrecoUnitario, dValor, sUM, objEstoqueProduto.dQuantDisponivel + objEstoqueProduto.dQuantReservada)
            If lErro <> AD_SQL_SUCESSO Then gError 213219
       
        End If
        
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213220
    
    Loop
    
    If iExibeProdSemMov = MARCADO Then
    
        sProduto = String(STRING_MAXIMO, 0)
        sDescricao = String(STRING_MAXIMO, 0)
        sUM = String(STRING_MAXIMO, 0)
    
        lErro = Comando_Executar(alComando(3), "SELECT Codigo, Descricao, SiglaUMVenda FROM Produtos WHERE Ativo = 0 AND Codigo NOT IN (SELECT Produto FROM RelOPGerEstTA WHERE NumIntRel = ?) ", _
                                                    sProduto, sDescricao, sUM, lNumIntRel)
        If lErro <> AD_SQL_SUCESSO Then gError 213221
        
        lErro = Comando_BuscarPrimeiro(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213222
        
        Do While lErro <> AD_SQL_SEM_DADOS
        
            dtDataPrev = Date
            
            Set objEstoqueProduto = New ClassEstoqueProduto
            objEstoqueProduto.sProduto = sProduto
        
            If giFilialEmpresa = EMPRESA_TODA Then
                'Lê a soma de todas as quantidades para Produto Passado em todos os Almoxarifados da Empresa
                lErro = CF("EstoqueProduto_Le_Todos_Almoxarifados", objEstoqueProduto)
                If lErro <> SUCESSO And lErro <> 52979 Then gError ERRO_SEM_MENSAGEM
            Else
                'Lê a soma de todas as quantidades para Produto Passado em todos os Almoxarifados da Filial
                lErro = CF("EstoqueProduto_Le_Todos_Almoxarifados_Filial", objEstoqueProduto, giFilialEmpresa)
                If lErro <> SUCESSO And lErro <> 64014 Then gError ERRO_SEM_MENSAGEM
            End If
            
            Call CF("Rel12Semanas_Semana", dtDataPrev, iSemana, dtDataIniSem)
        
            lErro = Comando_Executar(alComando(1), "INSERT INTO RelOPGerEstTA (NumIntRel,Semana,DataIniSemana,DataFimSemana,TipoDoc,NumIntDocItem,NumIntDoc,Codigo,FilialEmpresa,CliForn,NomeCliForn,DataPrev,Produto,Descricao,Quantidade,PrecoUnitario,Valor,UM,EstoqueInicial ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) ", _
            lNumIntRel, iSemana, dtDataIniSem, DateAdd("d", 6, dtDataIniSem), sProduto, 0, 0, 0, giFilialEmpresa, 0, "", dtDataPrev, sProduto, sDescricao, 0, 0, 0, sUM, objEstoqueProduto.dQuantDisponivel + objEstoqueProduto.dQuantReservada)
            If lErro <> AD_SQL_SUCESSO Then gError 213223
        
            lErro = Comando_BuscarProximo(alComando(3))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213224
        
        Loop
    
    End If
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Fechar transacao
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 213225
    
    RelOPGerEstTA_Prepara = SUCESSO

    Exit Function

Erro_RelOPGerEstTA_Prepara:

    RelOPGerEstTA_Prepara = gErr

    Select Case gErr

        Case 213214
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 213215
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 213216, 213217, 213218, 213220, 213221, 213222, 213224
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS", gErr)
            
        Case 213219, 213223
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERT_RELOPGERESTTA", gErr)
            
        Case 213225
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 213226)

    End Select
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Call Transacao_Rollback

    Exit Function
    
End Function

Public Function RelOPGerEstTA_PreparaSQL(sSQL As String, ByVal dtDataDe As Date, ByVal dtDataAte As Date, ByVal sProdDe As String, ByVal sProdAte As String) As Long

Dim lErro As Long
Dim sWhere As String
Dim sOrderBy As String

On Error GoTo Erro_RelOPGerEstTA_PreparaSQL

    sSQL = "SELECT TipoDoc,NumIntDocItem,NumIntDoc,Codigo,FilialEmpresa,CliForn,NomeCliForn,DataPrev,Produto, Descricao,Quantidade,PrecoUnitario,Valor, UM, QtdeAjuste FROM RelOPGerEstTAAux "
    sWhere = "WHERE DataPrev <> ?  "
    sOrderBy = "ORDER BY Produto, DataPrev "
    
    If giFilialEmpresa <> EMPRESA_TODA Then
        sWhere = sWhere & "AND FilialEmpresa = ? "
    End If
        
    If sProdDe <> "" Then
        sWhere = sWhere & "AND Produto >= ? "
    End If
    
    If sProdAte <> "" Then
        sWhere = sWhere & "AND Produto <= ? "
    End If
    
    If dtDataDe <> DATA_NULA Then
        sWhere = sWhere & "AND DataPrev >= ? "
    End If
    
    If dtDataAte <> DATA_NULA Then
        sWhere = sWhere & "AND DataPrev <= ? "
    End If
    
    sSQL = sSQL & sWhere & sOrderBy
    
    RelOPGerEstTA_PreparaSQL = SUCESSO
     
    Exit Function
    
Erro_RelOPGerEstTA_PreparaSQL:

    RelOPGerEstTA_PreparaSQL = gErr
     
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213227)
     
    End Select
     
    Exit Function

End Function

Private Function RelOPGerEstTA_PreparaInt(ByVal lComando As Long, ByVal sSQL As String, ByVal viFilEmp As Variant, ByVal vdtDataDe As Variant, ByVal vdtDataAte As Variant, ByVal vsProdDe As Variant, ByVal vsProdAte As Variant, vsTipoDoc As Variant, vlNumIntDocItem As Variant, vlNumIntDoc As Variant, vlCodigo As Variant, viFilialEmpresa As Variant, vlCliForn As Variant, vsNomeCliForn As Variant, vdtDataPrev As Variant, vsProduto As Variant, vsDescricao As Variant, vdQuantidade As Variant, vdPrecoUnitario As Variant, vdValor As Variant, vsUM As Variant, vdQtdeAjuste As Variant) As Long

Dim lErro As Long
Dim viStatus As Variant
Dim vdtDataNula As Variant

On Error GoTo Erro_RelOPGerEstTA_PreparaInt

    vdtDataNula = DATA_NULA

    lErro = Comando_PrepararInt(lComando, sSQL)
    If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM

    lErro = Comando_BindVarInt(lComando, vsTipoDoc)
    If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM

    lErro = Comando_BindVarInt(lComando, vlNumIntDocItem)
    If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    
    lErro = Comando_BindVarInt(lComando, vlNumIntDoc)
    If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    
    lErro = Comando_BindVarInt(lComando, vlCodigo)
    If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    
    lErro = Comando_BindVarInt(lComando, viFilialEmpresa)
    If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    
    lErro = Comando_BindVarInt(lComando, vlCliForn)
    If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    
    lErro = Comando_BindVarInt(lComando, vsNomeCliForn)
    If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    
    lErro = Comando_BindVarInt(lComando, vdtDataPrev)
    If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    
    lErro = Comando_BindVarInt(lComando, vsProduto)
    If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    
    lErro = Comando_BindVarInt(lComando, vsDescricao)
    If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    
    lErro = Comando_BindVarInt(lComando, vdQuantidade)
    If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    
    lErro = Comando_BindVarInt(lComando, vdPrecoUnitario)
    If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    
    lErro = Comando_BindVarInt(lComando, vdValor)
    If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    
    lErro = Comando_BindVarInt(lComando, vsUM)
    If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    
    lErro = Comando_BindVarInt(lComando, vdQtdeAjuste)
    If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    
    lErro = Comando_BindVarInt(lComando, vdtDataNula)
    If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    
    If viFilEmp <> 0 Then
        lErro = Comando_BindVarInt(lComando, viFilEmp)
        If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    End If
    
    If vsProdDe <> "" Then
        lErro = Comando_BindVarInt(lComando, vsProdDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    End If
    
    If vsProdAte <> "" Then
        lErro = Comando_BindVarInt(lComando, vsProdAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    End If
       
    If vdtDataDe <> DATA_NULA Then
        lErro = Comando_BindVarInt(lComando, vdtDataDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    End If
        
    If vdtDataAte <> DATA_NULA Then
        lErro = Comando_BindVarInt(lComando, vdtDataAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    End If
    
    lErro = Comando_ExecutarInt(lComando)
    If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    
    RelOPGerEstTA_PreparaInt = SUCESSO
     
    Exit Function
    
Erro_RelOPGerEstTA_PreparaInt:

    RelOPGerEstTA_PreparaInt = gErr
     
    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            Call Rotina_Erro(vbOKOnly, "ERRO_PREPARACAO_COMANDO_SQL", gErr, sSQL)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213228)
     
    End Select
     
    Exit Function

End Function

Public Function NFiscalEntCom_Gera_Reservas(ByVal objNF As ClassNFiscal, objME As ClassMovEstoque) As Long

Dim lErro As Long
Dim alComando(1 To 6) As Long, iIndice As Integer
Dim objReserva As New ClassReserva
Dim objItemME As ClassItemMovEstoque
Dim objMEReserva As ClassMovEstoque
Dim objItemMEReserva As ClassItemMovEstoque
Dim sSQL As String, objProduto As ClassProduto
Dim dQtdeFaltaReservar As Double, dQuantDisp As Double, dQuantReservadaPV As Double, dQuantReservadaME As Double
Dim dFator As Double, lNumIntDoc As Long, lCodigo As Long
Dim objPV As ClassPedidoDeVenda, objEstoqueProduto As ClassEstoqueProduto
Dim dQuantidade As Double, dQuantCancelada As Double, dQuantReservada As Double, dQuantFaturada As Double, sUM As String
Dim lNumIntDocItemPV As Long, lNumIntDocPV As Long, lCodPV As Long

On Error GoTo Erro_NFiscalEntCom_Gera_Reservas

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 213763
    Next
    
    For Each objItemME In objME.colItens
    
        Set objProduto = New ClassProduto
    
        objProduto.sCodigo = objItemME.sProduto
        
        'dQuantDisp = objItemME.dQuantidade
        
        Set objEstoqueProduto = New ClassEstoqueProduto
        objEstoqueProduto.sProduto = objItemME.sProduto
        objEstoqueProduto.iAlmoxarifado = objItemME.iAlmoxarifado
                    
        lErro = CF("EstoqueProduto_Le", objEstoqueProduto)
        If lErro <> SUCESSO And lErro <> 21306 Then gError ERRO_SEM_MENSAGEM
        
        'Considera o estoque disponível todo dos produtos que deram entrada
        dQuantDisp = objEstoqueProduto.dQuantDispNossa
        
        If dQuantDisp > QTDE_ESTOQUE_DELTA2 Then
        
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
            
            sUM = String(STRING_UM_SIGLA, 0)
        
            sSQL = "SELECT I.Quantidade, I.QuantCancelada, I.QuantReservada, I.QuantFaturada, I.UnidadeMed, I.NumIntDoc, P.Codigo, P.NumIntDoc"
            sSQL = sSQL & " FROM ItensPedidoDeVenda AS I, PedidosDeVenda AS P WHERE P.Codigo = I.CodPedido AND"
            sSQL = sSQL & " P.FilialEmpresa = I.FilialEmpresa AND P.FilialEmpresa = ? AND I.Produto = ? AND"
            sSQL = sSQL & " ABS (I.Quantidade - I.QuantCancelada - I.QuantReservada - I.QuantFaturada) > ? ORDER BY P.DataEmissao ASC"
    
            lErro = Comando_Executar(alComando(1), sSQL, dQuantidade, dQuantCancelada, dQuantReservada, dQuantFaturada, sUM, lNumIntDocItemPV, lCodPV, lNumIntDocPV, objItemME.iFilialEmpresa, objItemME.sProduto, QTDE_ESTOQUE_DELTA2)
            If lErro <> AD_SQL_SUCESSO Then gError 213764
            
            lErro = Comando_BuscarPrimeiro(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213765
                
            Do While lErro <> AD_SQL_SEM_DADOS
            
                Set objPV = New ClassPedidoDeVenda
               
                objPV.lNumIntDoc = lNumIntDocPV
                objPV.lCodigo = lCodPV
                objPV.iFilialEmpresa = objME.iFilialEmpresa
            
                'Quantidade que falta reservar
                dQtdeFaltaReservar = dQuantidade - dQuantCancelada - dQuantReservada - dQuantFaturada
                
                'Fator de conversão da UMPV para UM Estoque
                lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, sUM, objProduto.sSiglaUMEstoque, dFator)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                        
                'Quantidade do pedido convertida para a UM de Estoque
                dQtdeFaltaReservar = dQtdeFaltaReservar * dFator
            
                'Se a entrada for maior que a necessidade do pedido reserva tudo que o pedido precisa senão só o que estiver entrando no estoque
                If dQuantDisp > dQtdeFaltaReservar Then
                    dQuantReservadaME = dQtdeFaltaReservar
                    dQuantDisp = dQuantDisp - dQuantReservadaME
                Else
                    dQuantReservadaME = dQuantDisp
                    dQuantDisp = 0
                End If
                    
                'transforma a quantidade reservada para a UM de PV
                dQuantReservadaPV = dQuantReservadaME / dFator
                    
                lErro = Comando_Executar(alComando(2), "UPDATE ItensPedidoDeVenda SET QuantReservada = QuantReservada + ? WHERE NumIntDoc = ?", dQuantReservadaPV, lNumIntDocItemPV)
                If lErro <> AD_SQL_SUCESSO Then gError 213766
                    
                lErro = Comando_ExecutarPos(alComando(3), "SELECT Codigo FROM Reserva WHERE FilialEmpresa =? AND DocOrigem = ? AND Produto = ? AND Almoxarifado = ?", 0, lCodigo, objItemME.iFilialEmpresa, lCodPV, objItemME.sProduto, objItemME.iAlmoxarifado)
                If lErro <> AD_SQL_SUCESSO Then gError 213767
                
                lErro = Comando_BuscarPrimeiro(alComando(3))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213768
                
                If lErro = AD_SQL_SUCESSO Then
                
                    lErro = Comando_ExecutarPos(alComando(4), "UPDATE Reserva SET Quantidade = Quantidade + ?", alComando(3), dQuantReservadaME)
                    If lErro <> AD_SQL_SUCESSO Then gError 213769
                                        
                Else
                
                    'Zera a variável de número interno
                    lNumIntDoc = 0
            
                    'Gera o numero internos para as reservas
                    lErro = CF("Config_ObterNumInt", "MATConfig", "NUM_PROX_INT_RESERVA", lNumIntDoc)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
                    'Obtém o Código Automático
                    lErro = CF("Reserva_Automatico_Trans", lCodigo)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
                    'Insere a reserva no Banco de Dados
                    lErro = Comando_Executar(alComando(5), "INSERT INTO Reserva (FilialEmpresa,Codigo,Produto,Almoxarifado,TipoDoc,DocOrigem,NumIntOrigem,Quantidade,DataReserva,DataValidade,CodUsuario,Responsavel,NumIntDoc) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)", objItemME.iFilialEmpresa, lCodigo, objItemME.sProduto, objItemME.iAlmoxarifado, PEDIDO_DE_VENDA, lCodPV, lNumIntDocItemPV, dQuantReservadaPV, gdtDataHoje, DATA_NULA, gsUsuario, RESERVA_AUTO_RESP & " ENTRADA NF " & objNF.sSerie & " " & CStr(objNF.lNumNotaFiscal), lNumIntDoc)
                    If lErro <> AD_SQL_SUCESSO Then gError 213770
        
                End If
                
                Set objItemMEReserva = New ClassItemMovEstoque
            
                objItemMEReserva.iAlmoxarifado = objItemME.iAlmoxarifado
                objItemMEReserva.dQuantidade = dQuantReservadaME
                objItemMEReserva.sSiglaUM = objProduto.sSiglaUMEstoque
                objItemMEReserva.iFilialEmpresa = objItemME.iFilialEmpresa
                objItemMEReserva.iControleEstoque = objProduto.iControleEstoque
                objItemMEReserva.sSiglaUMEst = objProduto.sSiglaUMEstoque
                objItemMEReserva.sAlmoxarifadoNomeRed = objItemME.sAlmoxarifadoNomeRed
                objItemMEReserva.lNumIntDocOrigem = lNumIntDoc
                objItemMEReserva.iClasseUM = objProduto.iClasseUM
                objItemMEReserva.iApropriacaoProd = objItemME.iApropriacaoProd
                objItemMEReserva.sProduto = objItemME.sProduto
            
                'Gera o número interno do movimento
                lErro = CF("MovEstoque_Automatico_EmTransacao", objItemME.iFilialEmpresa, lCodigo)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
                'Preenche o objMovEstoque
                Set objMEReserva = New ClassMovEstoque
                objMEReserva.dtData = gdtDataHoje
                objMEReserva.iFilialEmpresa = objItemME.iFilialEmpresa
                objMEReserva.iTipoMov = MOV_EST_INCLUI_RESERVA
                objMEReserva.lCodigo = lCodigo
            
                'Adiciona ítem de MovEstoque
                Call objMEReserva.colItens.Add1(objItemMEReserva)
            
                'Grava o Movimento de Estoque
                lErro = CF("MovEstoque_Grava", objMEReserva)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
                lErro = CF("PV_Andamento_Auto", objPV, PV_STATUS_ANDAMENTO_TIPO_RESERVA_GRAVA)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
                'Se tudo que entrou já foi reservado não busca mais pedidos de venda
                If dQuantDisp < QTDE_ESTOQUE_DELTA2 Then Exit Do
                
                lErro = Comando_BuscarProximo(alComando(1))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213771
                
            Loop
        
        End If
    
    Next
    
    'Fechamento comando
     For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
     Next
             
    NFiscalEntCom_Gera_Reservas = SUCESSO

    Exit Function

Erro_NFiscalEntCom_Gera_Reservas:

    NFiscalEntCom_Gera_Reservas = gErr

    Select Case gErr

        Case 213763
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 213764, 213765, 213771
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDIDODEVENDA3", gErr, objItemME.iFilialEmpresa, lCodPV, objItemME.sProduto)
    
        Case 213766
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_ITENSPV", gErr, objItemME.iFilialEmpresa, lCodPV, objItemME.sProduto)

        Case 213767, 213768
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RESERVA2", gErr, objItemME.iFilialEmpresa, lCodPV, objItemME.sProduto, objItemME.iAlmoxarifado)

        Case 213769
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_RESERVA", gErr, objItemME.iFilialEmpresa, lCodPV, objItemME.sProduto, objItemME.iAlmoxarifado)

        Case 213770
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RESERVA", gErr, objItemME.iFilialEmpresa, lCodPV, objItemME.sProduto, objItemME.iAlmoxarifado)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213772)

    End Select

   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

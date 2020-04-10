VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFATGrava"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'colocar no regime normal até 24/02 ?

'212038
'lcomando 13

'cadastrar classesum 100,101 e 102 com siglas CJ, UP e PR. Incuir PC na classe 6.
'aumentar tamanho do nomereduzido do produto para 50
'cadastrar tipo de produto 9999
'preencher tabela ImportNFeTipoNF
'cadastrsr serie 1-e e 0-e
'cadastrar CJ e PR nas classe 2 e 6, UN nas classes 1 e 6, PC na classe 2...
'criar tipo de tributacao para cfops 5103 e 6103 em tributacao clientes
'aceitar estoque negativo
'criar almoxarifados 101, 102 e 103 nas filiaisempresa 1, 2 e 3
'update produtos set controleestoque =1
'criar meses no estoque a partir de 08/2010
'INSERT INTO EstoqueMes (filialEmpresa, Ano, Mes, Fechamento) VALUES (2, 2010, 8,1)
'GO
'INSERT INTO EstoqueMes (filialEmpresa, Ano, Mes, Fechamento) VALUES (2, 2010, 9,1)
'GO
'INSERT INTO EstoqueMes (filialEmpresa, Ano, Mes, Fechamento) VALUES (2, 2010, 10,1)
'GO
'INSERT INTO EstoqueMes (filialEmpresa, Ano, Mes, Fechamento) VALUES (2, 2010, 11,1)
'GO
'INSERT INTO EstoqueMes (filialEmpresa, Ano, Mes, Fechamento) VALUES (2, 2010, 12,1)
'GO
'excluir 35111112192120000117550000000023181000025497
'35111112192120000117550000000023181000025497
'35111112192120000117550010000023181000025493
'35111212192120000117550010000024241000026550
'35111212192120000117550000000024241000026554
'35111212192120000117550000000024891000027202
'35111212192120000117550010000024891000027209
'aumentar tamanho do nome reduzido do cliente (dic e dados) e tb razao social, ambos para 45
'será que tem nfs canceladas nos xmls enviados ?

'cadastrar fornecedor com cgc 64493273000106

Private Function TipoDocInfo_ObterFuncaoGravacao(ByVal objTipoDocInfo As ClassTipoDocInfo, sFuncaoGravacao As String) As Long

    Select Case objTipoDocInfo.sNomeTelaNFiscal
    
        Case "NFiscal"
            sFuncaoGravacao = "NFiscal_Grava_EmTrans"
        
        Case "NFiscalFatura"
            sFuncaoGravacao = "NFiscalFatura_Grava_EmTrans"
            
        Case "NFiscalDev"
            sFuncaoGravacao = "NFiscalDevolucao_Grava_EmTrans"
            
        Case "NFiscalEntDev"
            sFuncaoGravacao = "NFiscalEntradaDevolucao_Grava_EmTrans"
            
        Case "NFiscalRem"
            sFuncaoGravacao = "NFiscalRemessa_Grava_EmTrans"
            
        Case Else
            sFuncaoGravacao = ""
    
    End Select
    
    TipoDocInfo_ObterFuncaoGravacao = SUCESSO
    
End Function

Function ImportarXMLsProprios(ByVal iFilialEmpresa As Integer) As Long

Dim lErro As Long, iIndice As Integer
Dim lTransacao As Long, sSQL As String, sSQL2 As String, sCFOPPrincipal As String, sSQL3 As String, sSQL4 As String, sSQL5 As String
Dim alComando(0 To 30) As Long, lAux As Long, sFuncaoGravacao As String
Dim objNF As New ClassNFiscal, lCodCliente As Long, iCodFilial As Integer, lCodFornecedor As Long
Dim objItemNF As New ClassItemNF, objTribTab As ClassTribTab
Dim objTipoDocInfo As ClassTipoDocInfo
Dim objFilialEmpresa As New AdmFiliais, sNFs As String
Dim objProduto As ClassProduto
Dim sChvNFe As String, lNumNF As Long, sSerie As String, dtDataEmissao As Date, sCGCDestinatario As String, sNomeDestinatario As String, dNFICMSBase As Double, dNFICMSValor As Double, dNFICMSSubstBase As Double, dNFICMSSubstValor As Double, dValorProdutos As Double, dValorFrete As Double, dValorSeguro As Double, dValorDescontos As Double, dNFIPIValor As Double, dNFPISValor As Double, dNFCofinsValor As Double, dValorOutrasDespesas As Double
Dim dValorTotalNF As Double, sMsg(1 To 4) As String, sTipoES As String, sTipoCN As String, lNumIntNF As Long, sCGCTransportadora As String, dPesoLiq As Double, dPesoBruto As Double, sEspecie As String, lVolumeQuant As Long, sNomeEmitente As String, sNomeTransportadora As String
Dim lItem As Long, sProduto As String, sDescProd As String, sCFOP As String, sUM As String, dQuantidade As Double, dPrecoUnitario As Double, dPrecoTotal As Double, sEAN As String, sInfAdProd As String, dDesconto As Double, dFrete As Double, dSeguro As Double, dDespesas As Double, sISSCST As String, dISSBase As Double, dISSAliquota As Double, dISSValor As Double, iISSListServ As Integer, sISSCidadeIBGE As String, iOrigemMercadoria As Integer
Dim sICMSCST As String, sICMSCSOSN As String, iICMSBaseModalidade As Integer, dICMSBase As Double, dICMSPercRedBase As Double, dICMSAliquota As Double, dICMSValor As Double, iICMSSubstBaseModalidade As Integer, dICMSSubstBase As Double, dICMSSubstPercMVA As Double, dICMSSubstPercRedBase As Double, dICMSSubstAliquota As Double, dICMSSubstValor As Double
Dim dICMSSTCobrAntBase As Double, dICMSSTCobrAntValor As Double, dICMSPercBaseOperacaoPropria As Double, sICMSUFDevidoST As String, dICMSvBCSTRet As Double, dICMSvICMSSTRet As Double, dICMSvBCSTDest As Double, dICMSvICMSSTDest As Double, sICMSpCredSN As String, dICMSvCredSN As Double, sIPICST As String, sIPIEnquadramentoClasse As String
Dim sIPIEnquadramentoCodigo As String, sIPISeloCodigo As String, sIPICNPJProdutor As String, lIPISeloQtde As Long, dIPIBase As Double, dIPIAliquota As Double, dIPIValor As Double, dIPIUnidadePadraoQtde As Double, dIPIUnidadePadraoValor As Double, dIIValor As Double, dIIBase As Double, dIIDespAduaneira As Double, dIIIOF As Double, sPISCST As String, iPISTipoCalculo As Integer
Dim dPISBase As Double, dPISAliquota As Double, dPISValor As Double, dPISQtde As Double, dPISAliquotaValor As Double, iPISSTTipoCalculo As Integer, dPISSTBase As Double, dPISSTAliquota As Double, dPISSTValor As Double, dPISSTQtde As Double, dPISSTAliquotaValor As Double, sCOFINSCST As String, iCOFINSTipoCalculo As Integer, dCOFINSBase As Double, dCOFINSAliquota As Double
Dim dCOFINSValor As Double, dCOFINSQtde As Double, dCOFINSAliquotaValor As Double, iCOFINSSTTipoCalculo As Integer, dCOFINSSTBase As Double, dCOFINSSTAliquota As Double, dCOFINSSTValor As Double, dCOFINSSTQtde As Double, dCOFINSSTAliquotaValor As Double
Dim dtDataParc As Date, dValorParc As Double
Dim objParcela As ClassParcelaReceber, iTipoNFiscal As Integer
Dim sProdAux As String, sDescProdAux As String, sUMAux As String, sOrigemMercadoriaAux As String, sNomeRedAux As String, lContAux As Long, sNCMAux As String
Dim colTabelaPrecoItem As New Collection
Dim iICMSTipoSimples As Integer, iQtdNFsImportadas As Integer
Dim sProdEstProd As String, colRastreamento As New Collection, dFator As Double
Dim objEstoqueProduto As ClassEstoqueProduto, objItemAloc As ClassItemNFAlocacao
Dim objCliente As ClassCliente
Dim colEndereco As Collection, lProxCodCliente As Long
Dim objEndereco As New ClassEndereco, sMaxCFOP As String, sCGCAnt As String
Dim dtDataES As Date, sHoraES As String, dISSVlr As Double, dISSBC As Double, dPISVlr_ISS As Double, dCofinsVlr_ISS As Double, dVlrServico As Double, dPISRet As Double, dCofinsRet As Double, dCSLLRet As Double, dIRRFBC As Double, dIRRFVlr As Double, dINSSBCRet As Double, dINSSRet As Double, iExisteEndRetirada As Integer, iExisteEndEntrega As Integer, iExisteTransp As Integer, iExisteISSQNtot As Integer, iExisteRetTrib As Integer, iExisteCobr As Integer
Dim sNome As String, sNomeRed As String, sIE As String, sISUF As String, sIM As String, iCRT As Integer, sCodPais As String, sPais As String, sUF As String, sCodMunicipio As String, sMunicipio As String, sLogradouro As String, sNumero As String, sBairro As String, sComplemento As String, sCEP As String, sTelefone As String, sEmail As String
Dim sNCM As String, iIndTot As Integer, sEXTIPI As String, sEANTrib As String, dQtdTrib As Double, sUMTrib As String, dValorUnitTrib As Double
Dim sMunicipioAux As String, sCodMunicipioAux As String, lProxCodCidade As Long

On Error GoTo Erro_ImportarXMLsProprios

    'Abrir transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 212001

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 212002
    Next
        
    objFilialEmpresa.iCodFilial = iFilialEmpresa
    
    lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    sSQL = "SELECT DataES, HoraES, ISSVlr, ISSBC, PISVlr_ISS, CofinsVlr_ISS, VlrServico, PISRet, CofinsRet, CSLLRet, IRRFBC, IRRFVlr, INSSBCRet, INSSRet, ExisteEndRetirada, ExisteEndEntrega, ExisteTransp, ExisteISSQNtot, ExisteRetTrib, ExisteCobr, ChvNFe, NumNF, Serie, DataEmissao, CGCDestinatario, NomeDestinatario, NF.ICMSBase, NF.ICMSValor, NF.ICMSSubstBase, NF.ICMSSubstValor, ValorProdutos, ValorFrete, ValorSeguro, ValorDescontos, NF.IPIValor, NF.PISValor, NF.CofinsValor, ValorOutrasDespesas, "
    sSQL = sSQL & "ValorTotalNF, SUBSTRING(Mensagem,1,250), SUBSTRING(Mensagem,251,250), SUBSTRING(Mensagem,501,250), SUBSTRING(Mensagem,751,250), TipoES, TipoCN, CGCTransportadora, PesoLiq, PesoBruto, Especie, VolumeQuant, NomeEmitente, NomeTransportadora "
'    sSQL = sSQL & " FROM ImportNFEXML NF WHERE numnf not in (select numnf from importnfexml group by numnf having count(*) >1) AND LEFT(CGCDestinatario,8) NOT IN (SELECT left(cgcdestinatario, 8) from importnfexml group by left(cgcdestinatario, 8) having count(distinct cgcdestinatario)>1) AND CGCDestinatario <> '64493273000106' AND TipoES = 'S' AND ChvNFe NOT IN ('35110712192120000117550010000013821000016147', '35111112192120000117550000000023181000025497','35111212192120000117550000000024241000026554', '35111212192120000117550000000024891000027202') AND NumIntNF = 0 AND CGCEmitente = ? AND FilialEmpresa = ? AND NOT EXISTS (SELECT FilialEmpresa FROM NFiscal N WHERE N.FilialEmpresa = NF.FilialEmpresa AND N.DataEmissao = NF.DataEmissao AND (N.Serie+'-e') = NF.Serie AND NumNotaFiscal = NF.NumNF) ORDER BY ChvNFe"
    sSQL = sSQL & " FROM ImportNFEXML NF WHERE NumIntNF = 0 AND CGCEmitente = ? AND FilialEmpresa = ? AND NOT EXISTS (SELECT FilialEmpresa FROM NFiscal N WHERE N.FilialEmpresa = NF.FilialEmpresa AND N.DataEmissao = NF.DataEmissao AND (N.Serie+'-e') = NF.Serie AND NumNotaFiscal = NF.NumNF) ORDER BY ChvNFe"
    
    sSQL2 = "SELECT NCM, IndTot, EXTIPI, EANTrib, QtdTrib, UMTrib, ValorUnitTrib, Item, Produto, DescProd, CFOP, UM, Quantidade, PrecoUnitario, PrecoTotal, EAN, InfAdProd, Desconto, Frete, Seguro, Despesas, ISSCST, ISSBase, ISSAliquota, ISSValor, ISSListServ, ISSCidadeIBGE, OrigemMercadoria, "
    sSQL2 = sSQL2 & "ICMSCST, ICMSCSOSN, ICMSBaseModalidade, ICMSBase, ICMSPercRedBase, ICMSAliquota, ICMSValor, ICMSSubstBaseModalidade, ICMSSubstBase, ICMSSubstPercMVA, ICMSSubstPercRedBase, ICMSSubstAliquota, ICMSSubstValor, "
    sSQL2 = sSQL2 & "ICMSSTCobrAntBase, ICMSSTCobrAntValor, ICMSPercBaseOperacaoPropria, ICMSUFDevidoST, ICMSvBCSTRet, ICMSvICMSSTRet, ICMSvBCSTDest, ICMSvICMSSTDest, ICMSpCredSN, ICMSvCredSN, IPICST, IPIEnquadramentoClasse, "
    sSQL2 = sSQL2 & "IPIEnquadramentoCodigo, IPISeloCodigo, IPICNPJProdutor, IPISeloQtde, IPIBase, IPIAliquota, IPIValor, IPIUnidadePadraoQtde, IPIUnidadePadraoValor, IIValor, IIBase, IIDespAduaneira, IIIOF, PISCST, PISTipoCalculo, "
    sSQL2 = sSQL2 & "PISBase, PISAliquota, PISValor, PISQtde, PISAliquotaValor, PISSTTipoCalculo, PISSTBase, PISSTAliquota, PISSTValor, PISSTQtde, PISSTAliquotaValor, COFINSCST, COFINSTipoCalculo, COFINSBase, COFINSAliquota, "
    sSQL2 = sSQL2 & "CofinsValor , COFINSQtde, COFINSAliquotaValor, COFINSSTTipoCalculo, COFINSSTBase, COFINSSTAliquota, COFINSSTValor, COFINSSTQtde, COFINSSTAliquotaValor"
    sSQL2 = sSQL2 & " FROM ImportNFEItensXML WHERE ChvNFe = ? ORDER BY Item"
    
    sSQL3 = "SELECT Produto, MAX(DescProd), MAX(Um), MAX(OrigemMercadoria), MAX(NCM) FROM ImportNFEItensXML I1 WHERE I1.Chvnfe IN (SELECT ChvNFe FROM ImportNFEXml N, FiliaisEmpresa F WHERE F.CGC = N.CGCEmitente) AND Produto NOT IN (SELECT Codigo FROM Produtos) "
    sSQL3 = sSQL3 & "AND NOT EXISTS (SELECT Produto FROM ImportNFEItensXML I2 WHERE I2.Produto = I1.Produto AND LEN(I2.DescProd)> LEN(I1.DescProd)) GROUP BY Produto ORDER BY Produto"
    
    sSQL4 = "SELECT Codigo FROM Produtos P WHERE (ControleEstoque = 1 OR ControleEstoque = 2) AND NOT EXISTS (SELECT * FROM EstoqueProduto WHERE Produto = P.Codigo AND Almoxarifado = (100+?)) AND (EXISTS (SELECT * FROM ItensNFiscal I, NFiscal NF WHERE NF.NumIntDoc = I.NumIntNF AND Produto = P.Codigo AND FilialEmpresa = ?) OR EXISTS (SELECT * FROM ImportNFeItensXml I, ImportNFeXml NF WHERE NF.Chvnfe = I.chvnfe AND Produto = P.Codigo AND FilialEmpresa = ? AND CGCEmitente IN (SELECT CGC FROM FiliaisEmpresa))) ORDER BY Codigo"
    
'    sSQL5 = "SELECT DISTINCT CGCDestinatario FROM ImportNFeXml NF WHERE LEFT(CGCDestinatario,8) NOT IN (SELECT left(cgcdestinatario, 8) from importnfexml group by left(cgcdestinatario, 8) having count(distinct cgcdestinatario)>1) AND ChvNFe NOT IN ('35111112192120000117550000000023181000025497','35111212192120000117550000000024241000026554', '35111212192120000117550000000024891000027202') AND NumIntNF = 0 AND CGCEmitente = ? AND FilialEmpresa = ? AND NOT EXISTS (SELECT FilialEmpresa FROM NFiscal N WHERE N.FilialEmpresa = NF.FilialEmpresa AND N.DataEmissao = NF.DataEmissao AND (N.Serie+'-e') = NF.Serie AND NumNotaFiscal = NF.NumNF) AND (NOT EXISTS (SELECT * FROM FiliaisClientes WHERE CGC = CGCDestinatario) AND NOT EXISTS (SELECT * FROM FiliaisFornecedores WHERE CGC = CGCDestinatario)) ORDER BY CGCDestinatario"
    
    'cadastrar clientes que faltam
'    sCGCDestinatario = String(STRING_MAXIMO, 0)
'    lErro = Comando_Executar(alComando(10), sSQL5, sCGCDestinatario, objFilialEmpresa.sCgc, iFilialEmpresa)
'    If lErro <> AD_SQL_SUCESSO Then gError 212040
'
'    lErro = Comando_BuscarProximo(alComando(10))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 212041
'
'    Do While lErro <> AD_SQL_SEM_DADOS
'
'        If sCGCAnt <> sCGCDestinatario Then
'
'            If lProxCodCliente = 0 Then
'
'                lErro = Comando_Executar(alComando(11), "SELECT MAX(Codigo)+1 FROM Clientes", lProxCodCliente)
'                If lErro <> AD_SQL_SUCESSO Then gError 212043
'
'                lErro = Comando_BuscarProximo(alComando(11))
'                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 212044
'
'            End If
'
'            sNomeDestinatario = String(STRING_MAXIMO, 0)
'            sMaxCFOP = String(STRING_MAXIMO, 0)
'            lErro = Comando_Executar(alComando(12), "SELECT NomeDestinatario, (SELECT MAX(CFOP) FROM ImportNFEItensXml I WHERE I.ChvNFe = NF.ChvNFe) FROM ImportNFeXML NF WHERE CGCDestinatario = ? AND FilialEmpresa = ? ORDER BY ChvNFe", sNomeDestinatario, sMaxCFOP, sCGCDestinatario, iFilialEmpresa)
'            If lErro <> AD_SQL_SUCESSO Then gError 212045
'
'            lErro = Comando_BuscarProximo(alComando(12))
'            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 212046
'
'            Set objCliente = New ClassCliente
'            Set colEndereco = New Collection
'
'            With objCliente
'                .lCodigo = lProxCodCliente
'                lProxCodCliente = lProxCodCliente + 1
'
'                .sRazaoSocial = sNomeDestinatario
'                .sNomeReduzido = Trim(left(sNomeDestinatario, STRING_CLIENTE_NOME_REDUZIDO))
'                .iTipo = 1 '???
'                .iAtivo = 1 '???
'                .iProxCodFilial = 2 '???
'
'                .sCgc = sCGCDestinatario
'                .dtDataUltVisita = DATA_NULA
'                .iTipoFrete = TIPO_CIF '???
'
'                'Preenche colEndereco
'                Set objEndereco = New ClassEndereco
'
'                objEndereco.iCodigoPais = PAIS_BRASIL '???
'    '            objEndereco.sBairro = tImportCli.sBairro
'    '            objEndereco.sCEP = tImportCli.sCEP
'
'                'Verifica se o tamanho da string Cidade lida é maior que o permitido
'                objEndereco.sCidade = IIf(left(sMaxCFOP, 1) = "5", "São Paulo", "Rio de Janeiro")
'    '
'    '            objEndereco.sContato = tImportCli.sContato
'    '            objEndereco.sEmail = tImportCli.sEmail
'    '            objEndereco.sEndereco = tImportCli.sEndereco
'    '            objEndereco.sFax = tImportCli.sFax
'                objEndereco.sSiglaEstado = IIf(left(sMaxCFOP, 1) = "5", "SP", "RJ") '???
'
'                'Verifica se o tamanho da string Telefone lida é maior que o permitido
'    '            objEndereco.sTelefone1 = left(tImportCli.sTelefone1, STRING_TELEFONE)
'    '            objEndereco.sTelefone2 = left(tImportCli.sTelefone2, STRING_TELEFONE)
'
'                colEndereco.Add objEndereco
'
'                Set objEndereco = New ClassEndereco
'
'                objEndereco.iCodigoPais = PAIS_BRASIL '???
'    '            objEndereco.sBairro = tImportCli.sBairro1
'    '            objEndereco.sCEP = tImportCli.sCEP1
'    '
'                objEndereco.sCidade = IIf(left(sMaxCFOP, 1) = "5", "São Paulo", "Rio de Janeiro")
'    '
'    '            objEndereco.sContato = tImportCli.sContato1
'    '            objEndereco.sEmail = tImportCli.sEmail1
'    '            objEndereco.sEndereco = tImportCli.sEndereco1
'    '            objEndereco.sFax = tImportCli.sFax1
'                objEndereco.sSiglaEstado = IIf(left(sMaxCFOP, 1) = "5", "SP", "RJ") '???
'
'    '            objEndereco.sTelefone1 = tImportCli.sTelefone11
'    '            objEndereco.sTelefone2 = tImportCli.sTelefone21
'
'                colEndereco.Add objEndereco
'
'                Set objEndereco = New ClassEndereco
'
'                objEndereco.iCodigoPais = PAIS_BRASIL '???
'    '            objEndereco.sBairro = tImportCli.sBairro2
'    '            objEndereco.sCEP = tImportCli.sCEP2
'    '
'                objEndereco.sCidade = IIf(left(sMaxCFOP, 1) = "5", "São Paulo", "Rio de Janeiro")
'    '
'    '            objEndereco.sContato = tImportCli.sContato2
'    '            objEndereco.sEmail = tImportCli.sEmail2
'    '            objEndereco.sEndereco = tImportCli.sEndereco2
'    '            objEndereco.sFax = tImportCli.sFax2
'                objEndereco.sSiglaEstado = IIf(left(sMaxCFOP, 1) = "5", "SP", "RJ") '???
'
'    '            objEndereco.sTelefone1 = tImportCli.sTelefone12
'    '            objEndereco.sTelefone2 = tImportCli.sTelefone22
'
'                colEndereco.Add objEndereco
'
'
'            End With
'
'            'Grava o Cliente
'            lErro = CF("Cliente_Grava_EmTrans", objCliente, colEndereco)
'            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
'
'            sCGCAnt = sCGCDestinatario
'
'        End If
'
'        lErro = Comando_BuscarProximo(alComando(10))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 212042
'
'    Loop
    
    'cadastrar municipios
    sMunicipioAux = String(STRING_MAXIMO, 0)
    sCodMunicipioAux = String(STRING_MAXIMO, 0)
    lErro = Comando_Executar(alComando(14), "SELECT DISTINCT municipio, codmunicipio  from ImportNFeEndXml WHERE municipio <> '' and municipio not in (select descricao from cidades) ORDER BY Municipio", sMunicipioAux, sCodMunicipioAux)
    If lErro <> AD_SQL_SUCESSO Then gError 212022
    
    lErro = Comando_BuscarProximo(alComando(14))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 212022
    
    Do While lErro <> AD_SQL_SEM_DADOS
    
        If lProxCodCidade = 0 Then
        
            lErro = Comando_Executar(alComando(15), "SELECT MAX(Codigo)+1 FROM Cidades", lProxCodCidade)
            If lErro <> AD_SQL_SUCESSO Then gError 212022
            
            lErro = Comando_BuscarProximo(alComando(15))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 212022
            
        End If
        
        lErro = Comando_Executar(alComando(16), "INSERT INTO Cidades (Codigo, Descricao, CodIBGE) VALUES (?,?,?)", lProxCodCidade, sMunicipioAux, sCodMunicipioAux)
        If lErro <> AD_SQL_SUCESSO Then gError 212022
        
        lProxCodCidade = lProxCodCidade + 1
    
        lErro = Comando_BuscarProximo(alComando(14))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 212022
    
    Loop
    
''    'cadastrar produtos que faltam
''    sProdAux = String(STRING_PRODUTO, 0)
''    sDescProdAux = String(STRING_MAXIMO, 0)
''    sUMAux = String(STRING_MAXIMO, 0)
''    sOrigemMercadoriaAux = String(STRING_MAXIMO, 0)
''    sNCMAux = String(STRING_MAXIMO, 0)
''    lErro = Comando_Executar(alComando(7), sSQL3, sProdAux, sDescProdAux, sUMAux, sOrigemMercadoriaAux, sNCMAux)
''    If lErro <> AD_SQL_SUCESSO Then gError 212022
''
''    lErro = Comando_BuscarProximo(alComando(7))
''    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 212023
''
''    Do While lErro <> AD_SQL_SEM_DADOS
''
''        Set objProduto = New ClassProduto
''
''        With objProduto
''            .sCodigo = left(sProdAux + Space(20), 20)
''            .sDescricao = left(sDescProdAux, 50)
''            .sNomeReduzido = left(sDescProdAux, 50)
''            .iAtivo = 1
''            .iControleEstoque = PRODUTO_CONTROLE_RESERVA
''            .iNatureza = 4
''            .iFaturamento = 1
''            .iCompras = 1
''            .iApropriacaoCusto = APROPR_CUSTO_MEDIO
''            .iNivel = 1
''            .iTipo = 9999
''
''            Select Case sUMAux
''
''                Case "CJ", "CONJ", "JG"
''                    .iClasseUM = 100
''
''                Case "UN"
''                    .iClasseUM = 2
''
''                Case "PC", "PECAS"
''                    .iClasseUM = 6
''
''                Case "PÇ"
''                    .iClasseUM = 106
''
''                Case "UP", "UJ"
''                    .iClasseUM = 101
''
''                Case "PR", "PAR", "PARES"
''                    .iClasseUM = 102
''
''                Case "KG"
''                    .iClasseUM = 1
''
''                Case "M3"
''                    .iClasseUM = 103
''
''                Case "LITRO"
''                    .iClasseUM = 104
''
''                Case "M"
''                    .iClasseUM = 105
''
''                Case "01"
''                    .iClasseUM = 107
''
''                Case Else
''                    .iClasseUM = 12345
''
''            End Select
''
''            .sSiglaUMCompra = sUMAux
''            .sSiglaUMEstoque = sUMAux
''            .sSiglaUMVenda = sUMAux
''            .sSiglaUMTrib = sUMAux
''            .iOrigemMercadoria = StrParaInt(sOrigemMercadoriaAux)
''
''            'Informações referentes a Compras
''            .iConsideraQuantCotAnt = 1
''            .dPercentMenosQuantCotAnt = 0
''            .dPercentMaisQuantCotAnt = 0
''            .iTemFaixaReceb = 0
''            .dPercentMaisReceb = 0
''            .dPercentMenosReceb = 0
''            .iRecebForaFaixa = 1
''            .sIPICodigo = sNCMAux
''
''        End With
''
''        lErro = Comando_Executar(alComando(8), "SELECT COUNT(*) FROM Produtos WHERE NomeReduzido = ? AND Codigo <> ?", lContAux, objProduto.sNomeReduzido, objProduto.sCodigo)
''        If lErro <> AD_SQL_SUCESSO Then gError 212026
''
''        lErro = Comando_BuscarProximo(alComando(8))
''        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 212027
''
''        If lContAux <> 0 Then objProduto.sNomeReduzido = right(objProduto.sNomeReduzido + " - " + Trim(objProduto.sCodigo), 50)
''
''        'Grava o Produto
''        lErro = CF("Produto_Grava_Trans", objProduto, colTabelaPrecoItem)
''        If lErro <> SUCESSO Then gError 212024
''
''        lErro = Comando_BuscarProximo(alComando(7))
''        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 212025
''
''    Loop
''
''    'cadastrar estoque inicial para produtos que controlam estoque e ainda nao tem estoque inicial
''    sProdEstProd = String(STRING_PRODUTO, 0)
''
''    lErro = Comando_Executar(alComando(9), sSQL4, sProdEstProd, iFilialEmpresa, iFilialEmpresa, iFilialEmpresa)
''    If lErro <> AD_SQL_SUCESSO Then gError 212030
''
''    lErro = Comando_BuscarProximo(alComando(9))
''    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 212031
''
''    Do While lErro <> AD_SQL_SEM_DADOS
''
''        Set objEstoqueProduto = New ClassEstoqueProduto
''
''        objEstoqueProduto.sProduto = sProdEstProd
''        objEstoqueProduto.iAlmoxarifado = 100 + iFilialEmpresa
''        objEstoqueProduto.iFilialEmpresa = iFilialEmpresa
''        objEstoqueProduto.dQuantidadeInicial = 100000
''
''        Select Case iFilialEmpresa
''
''            Case 2
''                objEstoqueProduto.dtDataInicial = CDate("01/08/2010")
''
''            Case 1, 3
''                objEstoqueProduto.dtDataInicial = CDate("01/01/2007")
''
''            Case Else
''                gError 212033
''
''        End Select
''
''        lErro = CF("EstoqueInicial_Grava1", objEstoqueProduto, 100 + iFilialEmpresa, colRastreamento)
''        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
''
''        lErro = Comando_BuscarProximo(alComando(9))
''        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 212032
''
''    Loop
        
    sHoraES = String(STRING_MAXIMO, 0)
    sChvNFe = String(STRING_MAXIMO, 0)
    sSerie = String(STRING_MAXIMO, 0)
    sCGCDestinatario = String(STRING_MAXIMO, 0)
    sNomeDestinatario = String(STRING_MAXIMO, 0)
    For iIndice = 1 To 4
        sMsg(iIndice) = String(STRING_MAXIMO, 0)
    Next
    sTipoES = String(STRING_MAXIMO, 0)
    sTipoCN = String(STRING_MAXIMO, 0)
    sCGCTransportadora = String(STRING_MAXIMO, 0)
    sEspecie = String(STRING_MAXIMO, 0)
    sNomeEmitente = String(STRING_MAXIMO, 0)
    sNomeTransportadora = String(STRING_MAXIMO, 0)
    
    lErro = Comando_Executar(alComando(0), sSQL, _
        dtDataES, sHoraES, dISSVlr, dISSBC, dPISVlr_ISS, dCofinsVlr_ISS, dVlrServico, dPISRet, dCofinsRet, dCSLLRet, dIRRFBC, dIRRFVlr, dINSSBCRet, dINSSRet, iExisteEndRetirada, iExisteEndEntrega, iExisteTransp, iExisteISSQNtot, iExisteRetTrib, iExisteCobr, _
        sChvNFe, lNumNF, sSerie, dtDataEmissao, sCGCDestinatario, sNomeDestinatario, dNFICMSBase, dNFICMSValor, dNFICMSSubstBase, dNFICMSSubstValor, dValorProdutos, dValorFrete, dValorSeguro, dValorDescontos, dNFIPIValor, dNFPISValor, dNFCofinsValor, dValorOutrasDespesas, _
        dValorTotalNF, sMsg(1), sMsg(2), sMsg(3), sMsg(4), sTipoES, sTipoCN, sCGCTransportadora, dPesoLiq, dPesoBruto, sEspecie, lVolumeQuant, sNomeEmitente, sNomeTransportadora, _
        objFilialEmpresa.sCgc, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 212004
    
    lErro = Comando_BuscarProximo(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 212005
    
    Do While lErro <> AD_SQL_SEM_DADOS And iQtdNFsImportadas < 1
    
        DoEvents
        
        iQtdNFsImportadas = iQtdNFsImportadas + 1
    
        Set objNF = New ClassNFiscal
    
        'Inicializa a parte tributária
        Call objNF.Inicializa_Tributacao
        Call objNF.objTributacaoNF.Coloca_Auto
    
        sCFOPPrincipal = ""
    
        objNF.lNumNotaFiscal = lNumNF
        objNF.sSerie = sSerie & "-e"
        'If objNF.sSerie = "0-e" Then objNF.sSerie = "1-e"
        objNF.dtDataEmissao = dtDataEmissao
        objNF.iNFe = MARCADO
        For iIndice = 1 To 4
            objNF.sMensagemNota = objNF.sMensagemNota & sMsg(iIndice)
        Next
        objNF.iFilialEmpresa = iFilialEmpresa
        objNF.iNaoVerificaCredito = NAO_VERIFICA_CREDITO_CLIENTE
        objNF.iStatus = STATUS_BAIXADO
        objNF.dtDataReferencia = DATA_NULA
        objNF.iFreteRespons = FRETE_EMITENTE
        objNF.dtDataRegistro = Date
        objNF.dtDataVencimento = DATA_NULA
        
        objNF.dValorProdutos = dValorProdutos
        objNF.dValorFrete = dValorFrete
        objNF.dValorSeguro = dValorSeguro
        objNF.dValorOutrasDespesas = dValorOutrasDespesas
        objNF.dValorDescontoItens = dValorDescontos
        objNF.dValorTotal = dValorTotalNF
                
        objNF.dPesoLiq = dPesoLiq
        objNF.dPesoBruto = dPesoBruto
        'objNF. = sEspecie ,
        objNF.lVolumeQuant = lVolumeQuant
        
        If sTipoES = "S" Then
            objNF.objTributacaoNF.iTipoTributacao = 99
        Else
            objNF.objTributacaoNF.iTipoTributacao = 98
        End If
        objNF.objTributacaoNF.iTipoTributacaoManual = VAR_PREENCH_MANUAL
        objNF.objTributacaoNF.dICMSBase = dNFICMSBase
        objNF.objTributacaoNF.iICMSBaseManual = VAR_PREENCH_MANUAL
        objNF.objTributacaoNF.dICMSValor = dNFICMSValor
        objNF.objTributacaoNF.iICMSValorManual = VAR_PREENCH_MANUAL
        objNF.objTributacaoNF.dICMSSubstBase = dNFICMSSubstBase
        objNF.objTributacaoNF.iICMSSubstBaseManual = VAR_PREENCH_MANUAL
        objNF.objTributacaoNF.dICMSSubstValor = dNFICMSSubstValor
        objNF.objTributacaoNF.iICMSSubstValorManual = VAR_PREENCH_MANUAL
        objNF.objTributacaoNF.dIPIValor = dNFIPIValor
        objNF.objTributacaoNF.iIPIValorManual = VAR_PREENCH_MANUAL
        objNF.objTributacaoNF.dPISValor = dNFPISValor
        objNF.objTributacaoNF.iPISCreditoManual = VAR_PREENCH_MANUAL
        objNF.objTributacaoNF.dCOFINSValor = dNFCofinsValor
        objNF.objTributacaoNF.iCOFINSCreditoManual = VAR_PREENCH_MANUAL
        
        objNF.objTributacaoNF.dISSValor = dISSVlr
        objNF.objTributacaoNF.iISSValorManual = VAR_PREENCH_MANUAL
        objNF.objTributacaoNF.dISSBase = dISSBC
        objNF.objTributacaoNF.iISSBaseManual = VAR_PREENCH_MANUAL
        'dPISVlr_ISS
        'dCofinsVlr_ISS
        'dVlrServico
        objNF.objTributacaoNF.dPISRetido = dPISRet
        objNF.objTributacaoNF.iPISRetidoManual = VAR_PREENCH_MANUAL
        objNF.objTributacaoNF.dCOFINSRetido = dCofinsRet
        objNF.objTributacaoNF.iCOFINSRetidoManual = VAR_PREENCH_MANUAL
        objNF.objTributacaoNF.dCSLLRetido = dCSLLRet
        objNF.objTributacaoNF.iCSLLRetidoManual = VAR_PREENCH_MANUAL
        objNF.objTributacaoNF.dIRRFBase = dIRRFBC
        objNF.objTributacaoNF.iIRRFBaseManual = VAR_PREENCH_MANUAL
        objNF.objTributacaoNF.dIRRFValor = dIRRFVlr
        objNF.objTributacaoNF.iIRRFValorManual = VAR_PREENCH_MANUAL
        objNF.objTributacaoNF.dINSSBase = dINSSBCRet
        objNF.objTributacaoNF.iINSSBaseManual = VAR_PREENCH_MANUAL
        objNF.objTributacaoNF.dINSSValor = dINSSRet
        objNF.objTributacaoNF.iINSSValorManual = VAR_PREENCH_MANUAL
        If dINSSRet <> 0 Then objNF.objTributacaoNF.iINSSRetido = 1
        objNF.objTributacaoNF.iINSSRetidoManual = VAR_PREENCH_MANUAL
        
        sProduto = String(STRING_MAXIMO, 0)
        sDescProd = String(STRING_MAXIMO, 0)
        sCFOP = String(STRING_MAXIMO, 0)
        sUM = String(STRING_MAXIMO, 0)
        sEAN = String(STRING_MAXIMO, 0)
        sInfAdProd = String(STRING_MAXIMO, 0)
        sISSCST = String(STRING_MAXIMO, 0)
        sISSCidadeIBGE = String(STRING_MAXIMO, 0)
        sICMSCST = String(STRING_MAXIMO, 0)
        sICMSCSOSN = String(STRING_MAXIMO, 0)
        sICMSUFDevidoST = String(STRING_MAXIMO, 0)
        sICMSpCredSN = String(STRING_MAXIMO, 0)
        sIPICST = String(STRING_MAXIMO, 0)
        sIPIEnquadramentoClasse = String(STRING_MAXIMO, 0)
        sIPIEnquadramentoCodigo = String(STRING_MAXIMO, 0)
        sIPISeloCodigo = String(STRING_MAXIMO, 0)
        sIPICNPJProdutor = String(STRING_MAXIMO, 0)
        sPISCST = String(STRING_MAXIMO, 0)
        sCOFINSCST = String(STRING_MAXIMO, 0)
        sNCM = String(STRING_MAXIMO, 0)
        sEXTIPI = String(STRING_MAXIMO, 0)
        sEANTrib = String(STRING_MAXIMO, 0)
        sUMTrib = String(STRING_MAXIMO, 0)
                        
        lErro = Comando_Executar(alComando(1), sSQL2, _
            sNCM, iIndTot, sEXTIPI, sEANTrib, dQtdTrib, sUMTrib, dValorUnitTrib, _
            lItem, sProduto, sDescProd, sCFOP, sUM, dQuantidade, dPrecoUnitario, dPrecoTotal, sEAN, sInfAdProd, dDesconto, dFrete, dSeguro, dDespesas, sISSCST, dISSBase, dISSAliquota, dISSValor, iISSListServ, sISSCidadeIBGE, iOrigemMercadoria, _
            sICMSCST, sICMSCSOSN, iICMSBaseModalidade, dICMSBase, dICMSPercRedBase, dICMSAliquota, dICMSValor, iICMSSubstBaseModalidade, dICMSSubstBase, dICMSSubstPercMVA, dICMSSubstPercRedBase, dICMSSubstAliquota, dICMSSubstValor, _
            dICMSSTCobrAntBase, dICMSSTCobrAntValor, dICMSPercBaseOperacaoPropria, sICMSUFDevidoST, dICMSvBCSTRet, dICMSvICMSSTRet, dICMSvBCSTDest, dICMSvICMSSTDest, sICMSpCredSN, dICMSvCredSN, sIPICST, sIPIEnquadramentoClasse, _
            sIPIEnquadramentoCodigo, sIPISeloCodigo, sIPICNPJProdutor, lIPISeloQtde, dIPIBase, dIPIAliquota, dIPIValor, dIPIUnidadePadraoQtde, dIPIUnidadePadraoValor, dIIValor, dIIBase, dIIDespAduaneira, dIIIOF, sPISCST, iPISTipoCalculo, _
            dPISBase, dPISAliquota, dPISValor, dPISQtde, dPISAliquotaValor, iPISSTTipoCalculo, dPISSTBase, dPISSTAliquota, dPISSTValor, dPISSTQtde, dPISSTAliquotaValor, sCOFINSCST, iCOFINSTipoCalculo, dCOFINSBase, dCOFINSAliquota, _
            dCOFINSValor, dCOFINSQtde, dCOFINSAliquotaValor, iCOFINSSTTipoCalculo, dCOFINSSTBase, dCOFINSSTAliquota, dCOFINSSTValor, dCOFINSSTQtde, dCOFINSSTAliquotaValor, _
            sChvNFe)
        If lErro <> AD_SQL_SUCESSO Then gError 212006
        
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 212007
        
        Do While lErro <> AD_SQL_SEM_DADOS
    
            Set objItemNF = New ClassItemNF
        
            'Inicializa parte tributária
            Call objItemNF.Inicializa_Tributacao
            Call objItemNF.objTributacaoItemNF.Coloca_Auto
            
            'Obtém dados fixos
            objItemNF.dtDataEntrega = DATA_NULA
        
            'Obtém dados do Registro
            objItemNF.iItem = lItem
            
            Set objProduto = New ClassProduto
            
            objProduto.sCodigo = left(sProduto + Space(20), 20)
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 212008
            'Caso Produto não esteja cadastrado
            If lErro = 28030 Then gError 212009
            
            objItemNF.sProduto = objProduto.sCodigo
            objItemNF.sDescricaoItem = sDescProd
            objItemNF.dQuantidade = dQuantidade
            objItemNF.dPrecoUnitario = dPrecoUnitario
            objItemNF.dValorTotal = dPrecoTotal
            
            objItemNF.sUnidadeMed = sUM
            objItemNF.sUMEstoque = objProduto.sSiglaUMEstoque
            objItemNF.sUMVenda = objProduto.sSiglaUMVenda
            objItemNF.iClasseUM = objProduto.iClasseUM
            
            objNF.colItensNF.Add1 objItemNF
            
            objItemNF.objInfoAdicDocItem.iIncluiValorTotal = iIndTot
            
            objItemNF.objTributacaoItemNF.iTipoTributacao = objNF.objTributacaoNF.iTipoTributacao
            objItemNF.objTributacaoItemNF.iTipoTributacaoManual = VAR_PREENCH_MANUAL
            objItemNF.objTributacaoItemNF.iNaturezaOpManual = VAR_PREENCH_MANUAL
            objItemNF.objTributacaoItemNF.sNaturezaOp = sCFOP
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
            
            objItemNF.objTributacaoItemNF.dValorFreteItem = dFrete
            objItemNF.objTributacaoItemNF.dValorDescontoItem = dDesconto
            objItemNF.objTributacaoItemNF.dValorOutrasDespesasItem = dDespesas
            objItemNF.objTributacaoItemNF.dValorSeguroItem = dSeguro
            
            objItemNF.objTributacaoItemNF.sCST = sICMSCST
            objItemNF.objTributacaoItemNF.iCSTManual = VAR_PREENCH_MANUAL
            objItemNF.objTributacaoItemNF.sCSOSN = sICMSCSOSN
            objItemNF.objTributacaoItemNF.iICMSBaseModalidade = iICMSBaseModalidade
            objItemNF.objTributacaoItemNF.iICMSBaseModalidadeManual = VAR_PREENCH_MANUAL
            objItemNF.objTributacaoItemNF.dICMSBase = dICMSBase
            objItemNF.objTributacaoItemNF.iICMSBaseManual = VAR_PREENCH_MANUAL
            objItemNF.objTributacaoItemNF.dICMSPercRedBase = dICMSPercRedBase
            objItemNF.objTributacaoItemNF.iICMSPercRedBaseManual = VAR_PREENCH_MANUAL
            objItemNF.objTributacaoItemNF.dICMSAliquota = dICMSAliquota
            objItemNF.objTributacaoItemNF.iICMSAliquotaManual = VAR_PREENCH_MANUAL
            objItemNF.objTributacaoItemNF.dICMSValor = dICMSValor
            objItemNF.objTributacaoItemNF.iICMSValorManual = VAR_PREENCH_MANUAL
            
            objItemNF.objTributacaoItemNF.sUMTrib = sUMTrib
            objItemNF.objTributacaoItemNF.dQtdTrib = dQtdTrib
            objItemNF.objTributacaoItemNF.dValorUnitTrib = dValorUnitTrib
            
            If sICMSCSOSN <> "" Then
            
                lErro = Comando_Executar(alComando(8), "SELECT Tipo FROM TiposTribICMSSimples WHERE CSOSN = ?", iICMSTipoSimples, sICMSCSOSN)
                If lErro <> AD_SQL_SUCESSO Then gError 212028
        
                lErro = Comando_BuscarProximo(alComando(8))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 212029
                
                objItemNF.objTributacaoItemNF.iICMSSimplesTipo = iICMSTipoSimples
                objItemNF.objTributacaoItemNF.iICMSSimplesTipoManual = VAR_PREENCH_MANUAL
                
                objItemNF.objTributacaoItemNF.iRegimeTributario = REGIME_TRIBUTARIO_SIMPLES
            
            Else
            
                Select Case sICMSCST
                
                    Case "00"
                        objItemNF.objTributacaoItemNF.iICMSTipo = 1
                    
                    Case "10"
                        objItemNF.objTributacaoItemNF.iICMSTipo = 6
                    
                    Case "20"
                        objItemNF.objTributacaoItemNF.iICMSTipo = 7
                    
                    Case "30"
                        objItemNF.objTributacaoItemNF.iICMSTipo = 10
                    
                    Case "40"
                        objItemNF.objTributacaoItemNF.iICMSTipo = 2
                    
                    Case "41"
                        objItemNF.objTributacaoItemNF.iICMSTipo = 0
                    
                    Case "50"
                        objItemNF.objTributacaoItemNF.iICMSTipo = 3
                    
                    Case "51"
                        objItemNF.objTributacaoItemNF.iICMSTipo = 5
                    
                    Case "60"
                        objItemNF.objTributacaoItemNF.iICMSTipo = 8
                    
                    Case "70"
                        objItemNF.objTributacaoItemNF.iICMSTipo = 4
                    
                    Case "90"
                        objItemNF.objTributacaoItemNF.iICMSTipo = 99
                    
                    Case Else
                        gError 212038
                    
                End Select
                
                objItemNF.objTributacaoItemNF.iICMSTipoManual = VAR_PREENCH_MANUAL
                
                objItemNF.objTributacaoItemNF.iRegimeTributario = REGIME_TRIBUTARIO_NORMAL
                
            End If
            
            objItemNF.objTributacaoItemNF.iRegimeTributarioManual = VAR_PREENCH_MANUAL
                        
            If sCFOPPrincipal = "" Then sCFOPPrincipal = sCFOP
            
            '???? completar com ipi, pis, cofins, csll, iss...
            If sIPICST <> "" Then
            
                Select Case sIPICST
                
                    Case "50"
                        objItemNF.objTributacaoItemNF.iIPITipo = 1
                        
                    Case "51"
                        objItemNF.objTributacaoItemNF.iIPITipo = 4
                        
                    Case "52"
                        objItemNF.objTributacaoItemNF.iIPITipo = 3
                        
                    Case "53"
                        objItemNF.objTributacaoItemNF.iIPITipo = 0
                        
                    Case "54"
                        objItemNF.objTributacaoItemNF.iIPITipo = 7
                        
                    Case "55"
                        objItemNF.objTributacaoItemNF.iIPITipo = 5
                        
                    Case "99"
                        objItemNF.objTributacaoItemNF.iIPITipo = 6
                        
                    Case Else
                        gError 212039
                        
                End Select
                objItemNF.objTributacaoItemNF.iIPITipoManual = VAR_PREENCH_MANUAL
            
            Else
                objItemNF.objTributacaoItemNF.iIPITipoManual = VAR_PREENCH_AUTOMATICO
            End If
            
            Select Case sPISCST
            
                Case "01"
                    objItemNF.objTributacaoItemNF.iPISTipo = 1
                    
                Case Else
                    objItemNF.objTributacaoItemNF.iPISTipo = StrParaInt(sPISCST)
                
            End Select
            objItemNF.objTributacaoItemNF.iPISTipoManual = VAR_PREENCH_MANUAL
            
            Select Case sCOFINSCST
            
                Case "01"
                    objItemNF.objTributacaoItemNF.iCOFINSTipo = 1
                    
                Case Else
                    objItemNF.objTributacaoItemNF.iCOFINSTipo = StrParaInt(sCOFINSCST)
                
            End Select
            objItemNF.objTributacaoItemNF.iCOFINSTipoManual = VAR_PREENCH_MANUAL
            
            lErro = Comando_BuscarProximo(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 212010
        
        Loop
        
        If iExisteCobr <> 0 Then
        
            'carregar a cobrança
            lErro = Comando_Executar(alComando(2), "SELECT DataVencimento, Valor FROM ImportNFeCobrXml WHERE ChvNFe = ? ORDER BY Item", _
                dtDataParc, dValorParc, sChvNFe)
            If lErro <> AD_SQL_SUCESSO Then gError 212011
    
            lErro = Comando_BuscarProximo(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 212012
            
            iIndice = 0
            Do While lErro <> AD_SQL_SEM_DADOS
        
                iIndice = iIndice + 1
                
                Set objParcela = New ClassParcelaReceber
            
                With objParcela
                    .iNumParcela = iIndice
                    .dtDataVencimento = dtDataParc
                    .dValor = dValorParc
                    .dtDesconto1Ate = DATA_NULA
                    .dtDesconto2Ate = DATA_NULA
                    .dtDesconto3Ate = DATA_NULA
                    
                    objNF.ColParcelaReceber.Add 0, 0, iIndice, STATUS_ABERTO, .dtDataVencimento, .dtDataVencimento, .dValor, .dValor, 1, CARTEIRA_CARTEIRA, COBRADOR_PROPRIA_EMPRESA, "", 0, 0, 0, 0, 0, 0, .iDesconto1Codigo, .dtDesconto1Ate, .dDesconto1Valor, .iDesconto2Codigo, .dtDesconto2Ate, .dDesconto2Valor, .iDesconto3Codigo, .dtDesconto3Ate, .dDesconto3Valor, 0, 0, 0, 0, .iPrevisao, .sObservacao, .dValor
                End With
                
                lErro = Comando_BuscarProximo(alComando(2))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 212013
            
            Loop

        End If

        'ver pelos cfop qual o tipo de nf
        lErro = Comando_Executar(alComando(3), "SELECT TipoNFiscal FROM ImportNFeTipoNF WHERE Interna = 1 AND CFOP = ? AND Fatura = ?", _
            iTipoNFiscal, sCFOPPrincipal, IIf(objNF.ColParcelaReceber.Count <> 0, CInt(1), CInt(0)))
        If lErro <> AD_SQL_SUCESSO Then gError 212014

        lErro = Comando_BuscarProximo(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 212015
        
        'se pelo cfop sei qual o tipo de nf
        If lErro = AD_SQL_SUCESSO Then
        
            objNF.iTipoNFiscal = iTipoNFiscal
            
            Set objTipoDocInfo = New ClassTipoDocInfo
            
            objTipoDocInfo.iCodigo = objNF.iTipoNFiscal
            
            'Lê o Tipo de Documento
            lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
            If lErro <> SUCESSO And lErro <> 31415 Then gError ERRO_SEM_MENSAGEM
                
            If objTipoDocInfo.iDestinatario = DOCINFO_CLIENTE Or objTipoDocInfo.iOrigem = DOCINFO_CLIENTE Then
            
                'obter o cliente e filial
                lErro = Comando_Executar(alComando(5), "SELECT CodCliente, CodFilial FROM FiliaisClientes WHERE CGC = ? AND CGC <> ''", _
                    lCodCliente, iCodFilial, sCGCDestinatario)
                
                lErro = Comando_BuscarProximo(alComando(5))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 212016
                If lErro = AD_SQL_SUCESSO Then
                
                    objNF.lCliente = lCodCliente
                    objNF.iFilialCli = iCodFilial
                
                Else
                
                    'cadastrar o cliente
                    If lProxCodCliente = 0 Then
        
                        lErro = Comando_Executar(alComando(11), "SELECT MAX(Codigo)+1 FROM Clientes", lProxCodCliente)
                        If lErro <> AD_SQL_SUCESSO Then gError 212043
        
                        lErro = Comando_BuscarProximo(alComando(11))
                        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 212044
        
                    End If
                    
                    sNome = String(STRING_MAXIMO, 0)
                    sNomeRed = String(STRING_MAXIMO, 0)
                    sIE = String(STRING_MAXIMO, 0)
                    sISUF = String(STRING_MAXIMO, 0)
                    sIM = String(STRING_MAXIMO, 0)
                    sCodPais = String(STRING_MAXIMO, 0)
                    sPais = String(STRING_MAXIMO, 0)
                    sUF = String(STRING_MAXIMO, 0)
                    sCodMunicipio = String(STRING_MAXIMO, 0)
                    sMunicipio = String(STRING_MAXIMO, 0)
                    sLogradouro = String(STRING_MAXIMO, 0)
                    sNumero = String(STRING_MAXIMO, 0)
                    sBairro = String(STRING_MAXIMO, 0)
                    sComplemento = String(STRING_MAXIMO, 0)
                    sCEP = String(STRING_MAXIMO, 0)
                    sTelefone = String(STRING_MAXIMO, 0)
                    sEmail = String(STRING_MAXIMO, 0)
                    
                    lErro = Comando_Executar(alComando(13), "SELECT Nome, NomeRed, IE, ISUF, IM, CRT, CodPais, Pais, UF, CodMunicipio, Municipio, Logradouro, Numero, Bairro, Complemento, CEP, Telefone, Email FROM ImportNFePessoasXml P, ImportNFeEndXml E WHERE P.Tipo = E.Tipo AND P.ChvNFe = E.ChvNFe AND P.Tipo = 2 AND P.ChvNFe = ?", _
                        sNome, sNomeRed, sIE, sISUF, sIM, iCRT, sCodPais, sPais, sUF, sCodMunicipio, sMunicipio, sLogradouro, sNumero, sBairro, sComplemento, sCEP, sTelefone, sEmail, sChvNFe)
                    If lErro <> AD_SQL_SUCESSO Then gError 212047
    
                    lErro = Comando_BuscarProximo(alComando(13))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 212048
                    If lErro <> AD_SQL_SUCESSO Then gError 212049
                    
                    Set objCliente = New ClassCliente
                    Set colEndereco = New Collection
        
                    With objCliente
                        .lCodigo = lProxCodCliente
                        lProxCodCliente = lProxCodCliente + 1
        
                        .sRazaoSocial = left(Trim(sNome), STRING_CLIENTE_RAZAO_SOCIAL)
                        .sNomeReduzido = left(Trim(sNomeRed), STRING_CLIENTE_NOME_REDUZIDO)
                        If .sNomeReduzido = "" Then .sNomeReduzido = .sRazaoSocial
                        
                        .iTipo = 1 '???
                        .iAtivo = 1 '???
                        .iProxCodFilial = 2 '???
        
                        .sCgc = sCGCDestinatario
                        .sInscricaoEstadual = left(Trim(sIE), STRING_INSCR_EST)
                        .sInscricaoMunicipal = left(Trim(sIM), STRING_INSCR_MUN)
                        .sInscricaoSuframa = left(Trim(sISUF), STRING_INSCR_SUF)
                        .dtDataUltVisita = DATA_NULA
                        .iTipoFrete = TIPO_CIF '???
        
                        'Preenche colEndereco
                        Set objEndereco = New ClassEndereco
        
                        objEndereco.iCodigoPais = PAIS_BRASIL '???
                        objEndereco.sBairro = left(Trim(sBairro), STRING_BAIRRO)
                        objEndereco.sCEP = sCEP
        
                        'Verifica se o tamanho da string Cidade lida é maior que o permitido
                        objEndereco.sCidade = left(Trim(sMunicipio), STRING_CIDADE)
                        objEndereco.sEmail = left(Trim(sEmail), STRING_EMAIL)
                        
                        objEndereco.sEndereco = left(sLogradouro + IIf(Len(Trim(sNumero)) <> 0 And sNumero <> "SN", ", " + Trim(sNumero), "") + IIf(Len(Trim(sComplemento)) <> 0, " " + Trim(sComplemento), ""), STRING_ENDERECO)
                        
                        objEndereco.sSiglaEstado = sUF
        
                        
                        objEndereco.sTelefone1 = left(sTelefone, STRING_TELEFONE)
        
                        colEndereco.Add objEndereco
        
                        Set objEndereco = New ClassEndereco
        
                        If sCodPais <> "1058" Then gError 212050
                        
                        objEndereco.iCodigoPais = PAIS_BRASIL '???
                        objEndereco.sSiglaEstado = sUF
        
                        colEndereco.Add objEndereco
        
                        Set objEndereco = New ClassEndereco
        
                        objEndereco.iCodigoPais = PAIS_BRASIL '???
                        objEndereco.sSiglaEstado = sUF
        
                        colEndereco.Add objEndereco
        
                    End With
        
                    'Grava o Cliente
                    lErro = CF("Cliente_Grava_EmTrans", objCliente, colEndereco)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    
                    objNF.lCliente = objCliente.lCodigo
                    objNF.iFilialCli = FILIAL_MATRIZ

                End If
                
            Else
                
                'obter o fornecedor e filial
                lErro = Comando_Executar(alComando(6), "SELECT CodFornecedor, CodFilial FROM FiliaisFornecedores WHERE CGC = ? AND CGC <> ''", _
                    lCodFornecedor, iCodFilial, sCGCDestinatario)
                
                lErro = Comando_BuscarProximo(alComando(6))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 212018
                If lErro <> AD_SQL_SUCESSO Then gError 212019
                
                objNF.lFornecedor = lCodFornecedor
                objNF.iFilialForn = iCodFilial
                
            End If
            
            If objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIS Then
            
                objNF.dtDataEntrada = DATA_NULA
                objNF.dtHoraEntrada = 0
                
                If dtDataES <> DATA_NULA Then
                    objNF.dtDataSaida = dtDataES
                    objNF.iSemDataSaida = 0
                    objNF.dtHoraSaida = StrParaDate(sHoraES)
                Else
                    objNF.dtDataSaida = dtDataEmissao
                    objNF.iSemDataSaida = 1
                    objNF.dtHoraSaida = 0
                End If
                
            Else
            
                objNF.dtDataEntrada = IIf(dtDataES <> DATA_NULA, dtDataES, dtDataEmissao)
                objNF.dtHoraEntrada = StrParaDate(sHoraES)
                objNF.dtDataSaida = DATA_NULA
                objNF.dtHoraSaida = 0
            
            End If
            
            'de acordo com o tipo entao criar ou nao movimentacao
            If objTipoDocInfo.iTipoMovtoEstoque <> 0 Then
                
                For Each objItemNF In objNF.colItens
                
                    lErro = CF("UM_Conversao", objItemNF.iClasseUM, objItemNF.sUnidadeMed, objItemNF.sUMEstoque, dFator)
                    If lErro <> SUCESSO Then gError 212034
                    
                    Set objItemAloc = objItemNF.ColAlocacoes.Add(iFilialEmpresa + 100, "Outros " + CStr(iFilialEmpresa + 100), Arredonda_Estoque(objItemNF.dQuantidade * dFator))
                    objItemAloc.sUnidadeMed = objItemNF.sUMEstoque

                Next
            
            End If
    
            objNF.sNaturezaOp = sCFOPPrincipal
            objNF.objTributacaoNF.sNaturezaOp = sCFOPPrincipal
            objNF.objTributacaoNF.sNaturezaOpInterna = sCFOPPrincipal
            objNF.objTributacaoNF.iNaturezaOpManual = VAR_PREENCH_MANUAL
                    
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
            
            lErro = TipoDocInfo_ObterFuncaoGravacao(objTipoDocInfo, sFuncaoGravacao)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            objNF.iNumNFManual = 1 'para indicar que o numero da nf já é passado e nao deve ser buscado pela serie
            
            'Grava a Nota em Transação
            lErro = CF(sFuncaoGravacao, objNF, Nothing)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            If Len(Trim(sNFs)) <> 0 Then sNFs = sNFs & SEPARADOR
            sNFs = sNFs & CStr(objNF.lNumNotaFiscal)
            
            lErro = Comando_Executar(alComando(4), "UPDATE ImportNFeXml SET NumIntNF = ? WHERE ChvNFe = ? AND FilialEmpresa = ?", objNF.lNumIntDoc, sChvNFe, iFilialEmpresa)
            If lErro <> AD_SQL_SUCESSO Then gError 212020
        
        Else
'            MsgBox (sCFOPPrincipal)
        End If
        
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 212021
    
    Loop
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Fechar transacao
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 212003
    
    Call Rotina_Aviso(vbOKOnly, "AVISO_NF_IMPORT_COM_SUCESSO", sNFs)
    
    ImportarXMLsProprios = SUCESSO
    
    Exit Function
    
Erro_ImportarXMLsProprios:

    ImportarXMLsProprios = gErr
'Resume Next
    Select Case gErr

        Case 212001
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 212002
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 212003
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 212000)

    End Select
    
    Call Transacao_Rollback
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function


Function NF_Regera_MovEst() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim lTransacao As Long
Dim alComando(0 To 10) As Long
Dim objNF As ClassNFiscal, lNumIntDocNF As Long
Dim objMovEstoque As ClassMovEstoque
Dim lCodigoMovEst As Long, lNumIntDocMovEst As Long
Dim objItemNF As ClassItemNF, objItemAloc As ClassItemNFAlocacao
Dim dFator As Double, objProduto As ClassProduto
Dim alComando3(1 To 4) As Long
Dim alComando4(1 To NUM_MAX_LCOMANDO_MOVESTOQUE) As Long
Dim objTipoDocInfo As ClassTipoDocInfo

On Error GoTo Erro_NF_Regera_MovEst

    'Abrir transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 211545

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 211546
    Next
    
    'Abertura de Comando
    For iIndice = LBound(alComando4) To UBound(alComando4)
        alComando4(iIndice) = Comando_Abrir()
        If alComando4(iIndice) = 0 Then gError 211546
    Next
    
    'Abertura de Comando
    For iIndice = LBound(alComando3) To UBound(alComando3)
        alComando3(iIndice) = Comando_Abrir()
        If alComando3(iIndice) = 0 Then gError 211546
    Next
       
    lErro = Comando_Executar(alComando(1), "SELECT DISTINCT NumIntNF FROM ItensNFiscal WHERE NumIntDoc NOT IN (SELECT NumIntDocOrigem FROM MovimentoEstoque WHERE TipoNumIntDocOrigem = 1) AND NumIntNF IN (SELECT NumIntDoc FROM NFiscal WHERE Status <> 7 AND FilialEmpresa = 2 AND TipoNFiscal <> 119 ) ORDER BY NumIntNF", lNumIntDocNF)
    If lErro <> AD_SQL_SUCESSO Then gError 211547
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211548

    Do While lErro = AD_SQL_SUCESSO
           
        Set objNF = New ClassNFiscal
        Set objMovEstoque = New ClassMovEstoque
        
        objNF.lNumIntDoc = lNumIntDocNF
        
        lErro = CF("NFiscal_Le", objNF)
        If lErro <> SUCESSO And lErro <> 31442 Then gError ERRO_SEM_MENSAGEM
    
        'Lê os itens da nota fiscal
        lErro = CF("NFiscalItens_Le", objNF)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        Set objTipoDocInfo = New ClassTipoDocInfo
        
        objTipoDocInfo.iCodigo = objNF.iTipoNFiscal
        
        'Lê o Tipo de Documento
        lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
        If lErro <> SUCESSO And lErro <> 31415 Then gError ERRO_SEM_MENSAGEM
        
        objMovEstoque.iFilialEmpresa = objNF.iFilialEmpresa
        
        'Verifica se há movimento de estoque ligado a um dos itens
        lErro = Comando_Executar(alComando(2), "SELECT Codigo, NumIntDocEst FROM MovimentoEstoque WHERE TipoNumIntDocOrigem = ? AND NumIntDocOrigem IN (SELECT NumIntDoc FROM ItensNFiscal WHERE NumIntNF = ?) AND FilialEmpresa = ?", lCodigoMovEst, lNumIntDocMovEst, MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL, lNumIntDocNF, objMovEstoque.iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 83916
    
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 83917
        
        If lErro = AD_SQL_SUCESSO Then
        
            objMovEstoque.lCodigo = lCodigoMovEst
            
            'Lê os movimento dos itens de nota fiscal
            lErro = CF("MovEstoqueItens_Le_Desc1", objMovEstoque)
            If lErro <> SUCESSO And lErro <> 89050 Then gError ERRO_SEM_MENSAGEM
            
            If lErro = 89050 Then gError 89052
            
            'Redefine os movimento para fazer o cancelamento
            objMovEstoque.iOperacao = EXCLUSAO_MOVIMENTO
            
            'Grava o movimento de estoque
            lErro = CF("MovEstoque_Grava", objMovEstoque)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        End If
        
        'de acordo com o tipo entao criar ou nao movimentacao
        If objTipoDocInfo.iTipoMovtoEstoque <> 0 Then
            
            For Each objItemNF In objNF.colItens
            
                Set objProduto = New ClassProduto
                
                objProduto.sCodigo = objItemNF.sProduto
            
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
                
                objItemNF.iClasseUM = objProduto.iClasseUM
                objItemNF.sUMEstoque = objProduto.sSiglaUMEstoque
            
                lErro = CF("UM_Conversao", objItemNF.iClasseUM, objItemNF.sUnidadeMed, objItemNF.sUMEstoque, dFator)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                Set objItemAloc = objItemNF.ColAlocacoes.Add(objNF.iFilialEmpresa + 100, "Outros " + CStr(objNF.iFilialEmpresa + 100), Arredonda_Estoque(objItemNF.dQuantidade * dFator))
                objItemAloc.sUnidadeMed = objItemNF.sUMEstoque

            Next
        
        End If
        
        lErro = MovEstoque_Gera_Aux(alComando3(), alComando4(), objNF, objTipoDocInfo)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211554

    Loop
        
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    For iIndice = LBound(alComando3) To UBound(alComando3)
        Call Comando_Fechar(alComando3(iIndice))
    Next
    
    For iIndice = LBound(alComando4) To UBound(alComando4)
        Call Comando_Fechar(alComando4(iIndice))
    Next

    'Fechar transacao
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 211567
    
    Call Rotina_Aviso(vbOKOnly, "AVISO_ROTINA_EXECUTADA_COM_SUCESSO_REPROCESSE")

    NF_Regera_MovEst = SUCESSO

    Exit Function

Erro_NF_Regera_MovEst:

    NF_Regera_MovEst = gErr

    Select Case gErr

        Case 211545
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 211546
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)


        Case 211567
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
                        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211568)

    End Select
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    For iIndice = LBound(alComando3) To UBound(alComando3)
        Call Comando_Fechar(alComando3(iIndice))
    Next
    
    For iIndice = LBound(alComando4) To UBound(alComando4)
        Call Comando_Fechar(alComando4(iIndice))
    Next

    Call Transacao_Rollback

    Exit Function

End Function

Function MovEstoque_Gera_Aux(alComando() As Long, alComando1() As Long, ByVal objNFiscal As ClassNFiscal, ByVal objTipoDocInfo As ClassTipoDocInfo) As Long
'gera os movimentos de estoque da NFiscal, se necessários

Dim lErro As Long, iIndice As Integer
Dim iTipoMovtoEstoque As Integer
Dim objMovEstoque As ClassMovEstoque
Dim lCodigo As Long, dFator As Double
Dim objItemNF As ClassItemNF, dSomaAlocacao As Double
Dim objItemNFAloc As ClassItemNFAlocacao
Dim sDocOrigem As String
Dim objItemMovEstoque As ClassItemMovEstoque
Dim colInsumos As Collection
Dim colRastroMovto As Collection
Dim objProduto As New ClassProduto
Dim iTipoMov As Integer, bMovEst As Boolean
Dim dQuantItemNFUMEst As Double

On Error GoTo Erro_MovEstoque_Gera_Aux
   
    iTipoMovtoEstoque = objTipoDocInfo.iTipoMovtoEstoque
    sDocOrigem = objTipoDocInfo.sSigla

    If iTipoMovtoEstoque > 0 Then

        Set objMovEstoque = New ClassMovEstoque

        'Obtém o Código do Movimento
        lErro = CF("MovEstoque_Automatico_EmTransacao1", alComando(2), alComando(3), alComando(4), objNFiscal.iFilialEmpresa, lCodigo)
        If lErro <> SUCESSO Then gError 39001

        If objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIS Then
            objMovEstoque.dtData = objNFiscal.dtDataSaida
            objMovEstoque.dtHora = objNFiscal.dtHoraSaida
        Else
            objMovEstoque.dtData = objNFiscal.dtDataEntrada
            objMovEstoque.dtHora = objNFiscal.dtHoraEntrada
        End If
        objMovEstoque.iFilialEmpresa = objNFiscal.iFilialEmpresa
        objMovEstoque.iTipoMov = iTipoMovtoEstoque
        objMovEstoque.lCodigo = lCodigo
        objMovEstoque.lCliente = objNFiscal.lCliente
        objMovEstoque.lFornecedor = objNFiscal.lFornecedor
        objMovEstoque.sDocOrigem = sDocOrigem & " " & objNFiscal.sSerie & " " & CStr(objNFiscal.lNumNotaFiscal)

        If objTipoDocInfo.iTipo <> TIPODOCINFO_TIPO_NFIS Then
            'Se for para incluir o valor do frete e outras depesas no custo do ItemNF
            If gobjEST.iIncluiFreteNoCusto = INCLUI_FRETE_CUSTO Then
                'Para cada item da nota fiscal
                'Adiciona o valor rateado de Frete\Seguro e Outras Despesas (O Valor Creditado já está na tributação do Item)
                For Each objItemNF In objNFiscal.colItensNF
            
                    'Se é importação ignora as outras despesas
                    If objNFiscal.iTipoNFiscal = DOCINFO_NFIEIMP Or objNFiscal.iTipoNFiscal = DOCINFO_NFIEIMPSE Or objNFiscal.iTipoNFiscal = DOCINFO_NFIEIMPPC Then
                        objItemNF.dCusto = Arredonda_Moeda(objItemNF.dCusto + objItemNF.objTributacaoItemNF.dICMSValor + objItemNF.objTributacao.dValorFreteItem + objItemNF.objTributacao.dValorSeguroItem + objItemNF.dDespImpValorRateado)
                    Else
                        objItemNF.dCusto = Arredonda_Moeda(objItemNF.dCusto + objItemNF.objTributacao.dValorFreteItem + objItemNF.objTributacao.dValorOutrasDespesasItem + objItemNF.objTributacao.dValorSeguroItem)
                    End If
                    
                Next
            End If
        End If

        iIndice = 0
        
        'Adiciona itens ao Movimento
        For Each objItemNF In objNFiscal.colItensNF
                        
            iIndice = iIndice + 1
                                       
            If objItemNF.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then
                        
                If objItemNF.ColAlocacoes.Count = 0 And objItemNF.dQuantidade > 0 Then gError 42367

                lErro = CF("UM_Conversao", objItemNF.iClasseUM, objItemNF.sUnidadeMed, objItemNF.sUMEstoque, dFator)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

                dQuantItemNFUMEst = CDbl(Formata_Estoque(objItemNF.dQuantidade * dFator))

                If objTipoDocInfo.iTipo <> TIPODOCINFO_TIPO_NFIS Then
                    'calcula o custo do item da NF em função do credito do ICMS e do IPI
                    lErro = CF("NFiscalEntrada_Calcula_Custo", objItemNF)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                End If
                        
                dSomaAlocacao = 0
                                                                
                For Each objItemNFAloc In objItemNF.ColAlocacoes
                                        
                    dSomaAlocacao = dSomaAlocacao + Arredonda_Estoque(objItemNFAloc.dQuantidade)

                    Set colInsumos = New Collection
                    Set colRastroMovto = New Collection
                                        
                    Set objItemMovEstoque = objMovEstoque.colItens.Add(0, 0, objItemNF.dValorTotal, APROPR_CUSTO_INFORMADO, objItemNF.sProduto, objItemNF.sDescricaoItem, objItemNFAloc.sUnidadeMed, objItemNFAloc.dQuantidade, objItemNFAloc.iAlmoxarifado, objItemNFAloc.sAlmoxarifado, objItemNF.lNumIntDoc, "", 0, "", "", "", "", 0, colRastroMovto, colInsumos, DATA_NULA)
                    
                    If iTipoMov <> 0 Then objItemMovEstoque.iTipoMov = iTipoMov
                    
                    If objTipoDocInfo.iTipo <> TIPODOCINFO_TIPO_NFIS Then objItemMovEstoque.dCusto = objItemNF.dCusto

                    objItemMovEstoque.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL
                    
                    objItemMovEstoque.iClasseUM = objItemNF.iClasseUM
                    objItemMovEstoque.iControleEstoque = objItemNF.iControleEstoque
                    objItemMovEstoque.sSiglaUMEst = objItemNF.sUMEstoque
                    objItemMovEstoque.iApropriacaoProd = objItemNF.iApropriacaoProd
                    
                Next
                                
            End If
        Next

        'Grava os Movimentos de Estoque relacionados a Nota Fiscal
        lErro = CF("MovEstoque_Grava1", alComando1(), objMovEstoque)
        If lErro <> SUCESSO Then gError 39002
        
    End If

    MovEstoque_Gera_Aux = SUCESSO

    Exit Function

Erro_MovEstoque_Gera_Aux:

    MovEstoque_Gera_Aux = gErr

    Select Case gErr

        Case 35135, 35756
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPODOCINFO1", gErr, objNFiscal.iTipoNFiscal)

        Case 39000
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objNFiscal.iTipoNFiscal)

        Case 39001, 39002, 41525, 59261, 89461, 180064, 197727

        Case 42367
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCALIZACAO_ITEM_INEXISTENTE", gErr, iIndice)

        Case 42368
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCALIZACAO_ITEM_INCOMPLETA", gErr, iIndice, objItemNF.dQuantidade, dSomaAlocacao)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149807)

    End Select

    Exit Function

End Function


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPVGravaPhar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'#######################################################################
'Inserido por Wagner 02/02/2006
Public Function PedidoVenda_ArquivoLote_Gera(ByVal iOL As Integer, ByVal dtDataInicial As Date, ByVal dtDataFinal As Date, ByVal iFilialEmpresa As Integer, ByVal sDiretorio As String) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(1 To 2) As Long
Dim sSQL As String
Dim tPV As typePedidoVenda
Dim tNF As typeNFiscal
Dim tIPV As typeItemPedido
Dim tINF As typeItemNF
Dim tCliente As typeCliente
Dim tFilialCliente As typeFilialCliente
Dim tEndereco As typeEndereco
Dim sCabecalho As String
Dim sDetalhe As String
Dim lCodPVAntigo As Long
Dim lNumNotaAntigo As Long
Dim objItemMovEstoque As ClassItemMovEstoque
Dim objRastreamentoMovto As ClassRastreamentoMovto
Dim colMovEstoque As Collection
Dim colRastreamentoMovto As Collection
Dim bTemLote As Boolean
Dim sProduto As String
Dim dFator As Double
Dim sUM As String
Dim sTransportadora As String
Dim objProd As ClassProduto, lFornecedor As Long
Dim lNumIntDocAnot As Long, sTituloAnot As String, sTextoAnotParc As String
Dim sTituloAnotAnt As String, sTextoAnot As String

Const OL_KEEPERS = 1
Const OL_LMLOG = 2

Const OL_CODFORN_KEEPERS = 44
Const OL_CODFORN_LMLOG = 403

On Error GoTo Erro_PedidoVenda_ArquivoLote_Gera

    'Abertura comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 141635
    Next
    
    sUM = "CX"
    
    tCliente.sRazaoSocial = String(STRING_CLIENTE_RAZAO_SOCIAL, 0)
    tEndereco.sEndereco = String(STRING_ENDERECO, 0)
    tEndereco.sBairro = String(STRING_BAIRRO, 0)
    tEndereco.sCidade = String(STRING_CIDADE, 0)
    tEndereco.sSiglaEstado = String(STRING_ESTADO_SIGLA, 0)
    tEndereco.sCEP = String(STRING_CEP, 0)
    tFilialCliente.sCgc = String(STRING_CGC, 0)
    tINF.sProduto = String(STRING_PRODUTO, 0)
    tINF.sDescricaoItem = String(STRING_PRODUTO_DESCRICAO, 0)
    tINF.sUnidadeMed = String(STRING_UM_SIGLA, 0)
    tNF.sMensagemNota = String(STRING_NFISCAL_MENSAGEM, 0)
    sTransportadora = String(STRING_TRANSPORTADORA_NOME_REDUZIDO, 0)
    
    Open sDiretorio For Output As #1
    If iOL = OL_LMLOG Then
        lFornecedor = OL_CODFORN_LMLOG
        Open Replace(sDiretorio, "PEDIDO_0003_", "ITENS_PEDIDO_0003_") For Output As #2
        Open Replace(sDiretorio, "PEDIDO_0003_", "END_ENTREGA_0003_") For Output As #3
        Open Replace(sDiretorio, "PEDIDO_0003_", "OBS_PED_0003_") For Output As #4
    Else
        lFornecedor = OL_CODFORN_KEEPERS
    End If
 
    sSQL = "SELECT         PV.Codigo, C.Codigo, C.RazaoSocial, E.Endereco, E.Bairro, E.Cidade, E.SiglaEstado, E.CEP, FC.CGC, PV.DataEntrega, NF.DataEmissao, " & _
                          "INF.Item, INF.Produto, INF.DescricaoItem, INF.UnidadeMed, INF.Quantidade, INF.PrecoUnitario, NF.NumIntDoc, INF.NumIntdoc,  PV.NumIntDoc , IPV.NumIntDoc, NF.MensagemNota, IPV.ClasseUM, T.NomeReduzido, NF.NumNotaFiscal " & _
           "FROM           FiliaisClientes AS FC INNER JOIN " & _
                          "PedVenTodos AS PV ON FC.CodCliente = PV.Cliente AND FC.CodFilial = PV.Filial INNER JOIN " & _
                          "Clientes AS C ON PV.Cliente = C.Codigo INNER JOIN " & _
                          "Enderecos AS E ON FC.Endereco = E.Codigo INNER JOIN " & _
                          "NFiscal AS NF INNER JOIN " & _
                          "Phar_ItensNFiscalPack AS INF ON NF.NumIntDoc = INF.NumIntNF INNER JOIN " & _
                          "ItensPedidoDeVendaTodos AS IPV ON INF.NumIntItemPedVenda = IPV.NumIntDoc ON PV.Codigo = IPV.CodPedido AND PV.FilialEmpresa = IPV.FilialEmpresa LEFT OUTER JOIN " & _
                          "Transportadoras T ON NF.CodTransportadora = T.Codigo " & _
           "WHERE          NF.Status <> 7 AND NF.DataEmissao BETWEEN ? AND ? AND NF.FilialEmpresa = ? AND NF.FornEntTerc = ? ORDER BY NF.NumNotaFiscal, INF.Item"
    
    lErro = Comando_Executar(alComando(1), sSQL, tPV.lCodigo, tPV.lCliente, tCliente.sRazaoSocial, tEndereco.sEndereco, tEndereco.sBairro, tEndereco.sCidade, tEndereco.sSiglaEstado, tEndereco.sCEP, tFilialCliente.sCgc, tPV.dtDataEntrega, tNF.dtDataEmissao, _
    tINF.iItem, tINF.sProduto, tINF.sDescricaoItem, tINF.sUnidadeMed, tINF.dQuantidade, tINF.dPrecoUnitario, tNF.lNumIntDoc, tINF.lNumIntDoc, tPV.lNumIntDoc, tIPV.lNumIntDoc, tNF.sMensagemNota, tIPV.iClasseUM, sTransportadora, tNF.lNumNotaFiscal, dtDataInicial, dtDataFinal, iFilialEmpresa, lFornecedor)
    If lErro <> AD_SQL_SUCESSO Then gError 141636
        
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 141637

    Do While lErro = AD_SQL_SUCESSO
    
        lErro = Mascara_RetornaProdutoTela(tINF.sProduto, sProduto)
        If lErro <> SUCESSO Then gError 141720
        
        Set objProd = New ClassProduto
        
        objProd.sCodigo = tINF.sProduto
        
        lErro = CF("Produto_Le", objProd)
        If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
    
        If tNF.lNumNotaFiscal <> lNumNotaAntigo Then
        
            If iOL = OL_KEEPERS Then
            
                sCabecalho = "1" & FormataCpoTexto(CStr(tNF.lNumNotaFiscal), 10) & FormataCpoTexto(CStr(tPV.lCliente), 15) & FormataCpoTexto(tCliente.sRazaoSocial, 35) & FormataCpoTexto(tEndereco.sEndereco, 30) & FormataCpoTexto("", 20) & FormataCpoTexto(tEndereco.sBairro, 15) & _
                            FormataCpoTexto(tEndereco.sCidade, 20) & FormataCpoTexto(tEndereco.sSiglaEstado, 2) & FormataCpoTexto(tEndereco.sCEP, 8) & FormataCpoTexto(tFilialCliente.sCgc, 14) & FormataCpoTexto("", 8) & FormataCpoData(tNF.dtDataEmissao) & "M3S" & FormataCpoTexto(tNF.sMensagemNota, 200) & FormataCpoTexto(sTransportadora, 30)
        
                Print #1, sCabecalho
        
            Else
            
                sCabecalho = CStr(tNF.lNumNotaFiscal) & "|3|" & Format(tNF.dtDataEmissao, "DDMMYYYY") & "|" & CStr(tPV.lCodigo)
            
                Print #1, sCabecalho
            
                sCabecalho = CStr(tNF.lNumNotaFiscal) & "|3|" & tEndereco.sEndereco & "||" & tEndereco.sBairro & "|" & tEndereco.sCidade & "|" & tEndereco.sSiglaEstado & "|" & tEndereco.sCEP
                
                Print #3, sCabecalho
                
                sTituloAnot = String(STRING_MAXIMO, 0)
                sTextoAnotParc = String(STRING_MAXIMO, 0)
                
                lErro = Comando_Executar(alComando(2), "SELECT A.NumIntDoc, A.Titulo, L.Texto FROM Anotacoes AS A, AnotacoesLinha AS L WHERE A.NumIntDoc = L.NumIntDoc AND A.Origem = 3 AND A.ID = ? AND A.Titulo IN (?,?,?,?,?,?) ORDER BY A.Titulo, A.NumIntDoc, L.Seq", _
                lNumIntDocAnot, sTituloAnot, sTextoAnotParc, tNF.lNumIntDoc, "GERV", "CRED", "ADMV", "SEPA", "FATU", "EXPE")
                If lErro <> AD_SQL_SUCESSO Then gError 141636
                    
                lErro = Comando_BuscarPrimeiro(alComando(2))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 141637
            
                sTituloAnotAnt = UCase(sTituloAnot)
                
                If lErro = AD_SQL_SUCESSO Then
            
                    Do While lErro = AD_SQL_SUCESSO
                        
                        sTextoAnot = sTextoAnot & Replace(Replace(Replace(sTextoAnotParc, Chr(13), " "), Chr(10), ""), "|", "")
                    
                        lErro = Comando_BuscarProximo(alComando(2))
                        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 141637
                    
                        If sTituloAnotAnt <> UCase(sTituloAnot) Or lErro = AD_SQL_SEM_DADOS Then
                            
                            sCabecalho = CStr(tNF.lNumNotaFiscal) & "|3|" & sTituloAnotAnt & "|" & sTextoAnot
                            
                            sTituloAnotAnt = UCase(sTituloAnot)
                            
                            Print #4, sCabecalho
                            
                            sTextoAnot = ""
                        End If
                    
                    Loop
                
                Else
                    
                    sCabecalho = CStr(tNF.lNumNotaFiscal) & "|3|SEPA|"
                    
                    Print #4, sCabecalho
                    
                End If
                
            End If
       
            lCodPVAntigo = tPV.lCodigo
            lNumNotaAntigo = tNF.lNumNotaFiscal
            
        End If
        
        If iOL = OL_KEEPERS Then
            
            Set objItemMovEstoque = New ClassItemMovEstoque
    
            bTemLote = False
    
            'Lê item de movimento de estoque
            objItemMovEstoque.lNumIntDocOrigem = tINF.lNumIntDoc
            objItemMovEstoque.iTipoNumIntDocOrigem = TIPO_ORIGEM_ITEMNF
            objItemMovEstoque.iFilialEmpresa = iFilialEmpresa
            
            Set colMovEstoque = New Collection
            
            lErro = CF("MovEstoque_Le_ItemNF", objItemMovEstoque, colMovEstoque)
            If lErro <> SUCESSO Then gError 141638
            
            For Each objItemMovEstoque In colMovEstoque
            
                'se forem os movimentos de transferencia de material consignado ==> não trata-os, pois a venda de material consignado está sendo tratada
                If objItemMovEstoque.iTipoMov <> MOV_EST_SAIDA_TRANSF_CONSIG_TERC And _
                   objItemMovEstoque.iTipoMov <> MOV_EST_ENTRADA_TRANSF_DISP1 Then
            
                    'Lê movimentos de rastreamento vinculados ao itemNF passado ao ItemNF
                    Set colRastreamentoMovto = New Collection
                    
                    lErro = CF("RastreamentoMovto_Le_DocOrigem", objItemMovEstoque.lNumIntDoc, TIPO_RASTREAMENTO_MOVTO_MOVTO_ESTOQUE, colRastreamentoMovto)
                    If lErro <> SUCESSO Then gError 141639
            
                    'Guarda as quantidades alocadas dos lotes
                    For Each objRastreamentoMovto In colRastreamentoMovto
                    
                        bTemLote = True
                                                            
                        If tIPV.iClasseUM <> 6 Then
                        
                            'Calcula o fator de conversão de UM
                            lErro = CF("UM_Conversao", tIPV.iClasseUM, objRastreamentoMovto.sSiglaUM, tINF.sUnidadeMed, dFator)
                            If lErro <> SUCESSO Then gError 141770
                        
                        Else
                        
                            dFator = 1
                            
                        End If
                                                            
                        sDetalhe = "9" & FormataCpoTexto(CStr(tPV.lCodigo), 10) & FormataCpoTexto(CStr(tINF.iItem), 3) & FormataCpoTexto(sProduto, 25) & FormataCpoTexto(tINF.sDescricaoItem, 20) & FormataCpoTexto(sUM, 2) & FormataCpoNum(objRastreamentoMovto.dQuantidade * dFator, 11) & FormataCpoValor(tINF.dPrecoUnitario * 100 / dFator, 20) & FormataCpoTexto("", 15) & FormataCpoTexto(objRastreamentoMovto.sLote, 16)
                        Print #1, sDetalhe
                                
                    Next
                    
                End If
    
            Next
            
            dFator = 1
            
            If Not bTemLote Then
                sDetalhe = "9" & FormataCpoTexto(CStr(tPV.lCodigo), 10) & FormataCpoTexto(CStr(tINF.iItem), 3) & FormataCpoTexto(sProduto, 25) & FormataCpoTexto(tINF.sDescricaoItem, 20) & FormataCpoTexto(sUM, 2) & FormataCpoNum(tINF.dQuantidade * dFator, 11) & FormataCpoValor(tINF.dPrecoUnitario * 100 / dFator, 20) & FormataCpoTexto("", 15) & FormataCpoTexto("", 16)
                Print #1, sDetalhe
            End If
            
        Else
    
            sDetalhe = CStr(tNF.lNumNotaFiscal) & "|3|" & sProduto & "|" & objProd.sCodigoBarras & "|" & CStr(tINF.dQuantidade)
            Print #2, sDetalhe
        
        End If
        

        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 141640

    Loop
    
    Close #1
    If iOL = OL_LMLOG Then
        Close #2
        Close #3
        Close #4
    End If
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    PedidoVenda_ArquivoLote_Gera = SUCESSO
    
    Exit Function

Erro_PedidoVenda_ArquivoLote_Gera:

    PedidoVenda_ArquivoLote_Gera = gErr

    Select Case gErr
    
        Case 75
            Call Rotina_Erro(vbOKOnly, "ERRO_SEM_ACESSO_DIRETORIO", gErr)
        
        Case 76
            Call Rotina_Erro(vbOKOnly, "ERRO_DIRETORIO_INVALIDO", gErr, sDiretorio)

        Case 141635
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 141636, 141637, 141640
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDODEVENDA", gErr)
        
        Case 141638, 141639, 141720, 141770, 141771

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 141721)

    End Select
    
    Close #1
    If iOL = OL_LMLOG Then
        Close #2
        Close #3
        Close #4
    End If
    
   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function
    
End Function
'#######################################################################

Function PIS_Excecao_Pesquisar(ByVal colCateg As Collection, dAliquota As Double) As Long

Dim lErro As Long, objProdutoCategoria As ClassProdutoCategoria, iAchou As Integer

On Error GoTo Erro_PIS_Excecao_Pesquisar

    For Each objProdutoCategoria In colCateg
    
        If UCase(objProdutoCategoria.sCategoria) = "PISCOFINS" And UCase(objProdutoCategoria.sItem) = "DIFERENCIADA" Then
        
            dAliquota = 0.022
            iAchou = 1
            Exit For
            
        End If
        
    Next

    If iAchou = 0 Then gError 130761
    
    PIS_Excecao_Pesquisar = SUCESSO
     
    Exit Function
    
Erro_PIS_Excecao_Pesquisar:

    PIS_Excecao_Pesquisar = gErr
     
    Select Case gErr
          
        'nao encontrou excecao
        Case 130761
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 130764)
     
    End Select
     
    Exit Function

End Function

Function COFINS_Excecao_Pesquisar(ByVal colCateg As Collection, dAliquota As Double) As Long

Dim lErro As Long, objProdutoCategoria As ClassProdutoCategoria, iAchou As Integer

On Error GoTo Erro_COFINS_Excecao_Pesquisar

    For Each objProdutoCategoria In colCateg
    
        If UCase(objProdutoCategoria.sCategoria) = "PISCOFINS" And UCase(objProdutoCategoria.sItem) = "DIFERENCIADA" Then
        
            dAliquota = 0.103
            iAchou = 1
            Exit For
            
        End If
        
    Next

    If iAchou = 0 Then gError 130767
    
    COFINS_Excecao_Pesquisar = SUCESSO
     
    Exit Function
    
Erro_COFINS_Excecao_Pesquisar:

    COFINS_Excecao_Pesquisar = gErr
     
    Select Case gErr
          
        'nao encontrou excecao
        Case 130767
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 130766)
     
    End Select
     
    Exit Function

End Function

Public Function Integracao_PV_Obtem_NaturezaOP_Customizado(ByVal objPV As ClassPedidoDeVenda, ByVal objFilialCliente As ClassFilialCliente, ByVal objEndereco As ClassEndereco)
'Obtém a natureza de operação para a importação de pedidos de venda

Dim lErro As Long

On Error GoTo Erro_Integracao_PV_Obtem_NaturezaOP_Customizado

    objPV.iNaturezaOpManual = VAR_PREENCH_MANUAL
    
    If objPV.dValorTotal > DELTA_VALORMONETARIO Then
        If Len(objFilialCliente.sCgc) < 14 Then
            If objEndereco.sSiglaEstado = "SP" Then
                objPV.sNaturezaOp = "5102"
            Else
                objPV.sNaturezaOp = "6108"
            End If
        Else
            If objEndereco.sSiglaEstado = "SP" Then
                objPV.sNaturezaOp = "5102"
            Else
                If Len(Trim(objFilialCliente.sInscricaoEstadual)) > 0 Then
                    objPV.sNaturezaOp = "6102"
                Else
                    objPV.sNaturezaOp = "6108"
                End If
            End If
        End If
    Else
        If objEndereco.sSiglaEstado = "SP" Then
            objPV.sNaturezaOp = "5910"
        Else
            objPV.sNaturezaOp = "6910"
        End If
    End If
    
    Integracao_PV_Obtem_NaturezaOP_Customizado = SUCESSO
    
    Exit Function

Erro_Integracao_PV_Obtem_NaturezaOP_Customizado:

    Integracao_PV_Obtem_NaturezaOP_Customizado = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190088)

    End Select
       
    Exit Function

End Function

Public Function Integracao_PV_Obtem_CanalVenda_Customizado(ByVal objPV As ClassPedidoDeVenda, ByVal objTipoCliente As ClassTipoCliente)
'Obtém o canal de venda para a importação de pedidos de venda

Dim lErro As Long

On Error GoTo Erro_Integracao_PV_Obtem_CanalVenda_Customizado

    Select Case objTipoCliente.sDescricao
    
        Case "REDE"
             objPV.iCanalVenda = 1
             
        Case "DENTAL"
             objPV.iCanalVenda = 5
             
        Case "FARMACIA", "DISTRIBUIDOR"
             objPV.iCanalVenda = 3
             
        Case Else
             objPV.iCanalVenda = 6

    End Select
   
    Integracao_PV_Obtem_CanalVenda_Customizado = SUCESSO
    
    Exit Function

Erro_Integracao_PV_Obtem_CanalVenda_Customizado:

    Integracao_PV_Obtem_CanalVenda_Customizado = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190089)

    End Select
       
    Exit Function

End Function

Public Function Exporta_Dados_SldProd_AlmValido_Cust(ByVal iAlmoxarifado As Integer, bValido As Boolean)
'Verifica se o saldo do almoxarifado deve ser exportado

Dim lErro As Long

On Error GoTo Erro_Exporta_Dados_SldProd_AlmValido_Cust

    If iAlmoxarifado = 1 Then
        bValido = True
    Else
        bValido = False
    End If
   
    Exporta_Dados_SldProd_AlmValido_Cust = SUCESSO
    
    Exit Function

Erro_Exporta_Dados_SldProd_AlmValido_Cust:

    Exporta_Dados_SldProd_AlmValido_Cust = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190118)

    End Select
       
    Exit Function

End Function

Public Function Integracao_PV_Obtem_Frete_Customizado(ByVal objPV As ClassPedidoDeVenda, ByVal objTipoCliente As ClassTipoCliente)
'Obtém o canal de venda para a importação de pedidos de venda

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(1 To 3) As Long
Dim sSQL As String

Const FRETE_REDE = "FRETE_REDE"
Const FRETE_FARMACIA = "FRETE_FARMACIA"
Const FRETE_OUTROS = "FRETE_OUTROS"

Dim dFreteRede As Double
Dim dFreteFarmacia As Double
Dim dFreteOutros As Double
Dim sFrete As String
Dim sExpressao As String

On Error GoTo Erro_Integracao_PV_Obtem_Frete_Customizado

    'Abertura comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 192414
    Next
    
    sSQL = "SELECT Expressao FROM MnemonicoFPreco WHERE Mnemonico = ? AND FilialEmpresa = ?"
    
    sFrete = FRETE_REDE
    sExpressao = String(STRING_MNEMONICOFPRECO_EXPRESSAO, 0)
    lErro = Comando_Executar(alComando(1), sSQL, sExpressao, FRETE_REDE, giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 192415
        
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192416
    
    If lErro <> AD_SQL_SUCESSO Then gError 192417
    
    dFreteRede = StrParaDbl(sExpressao)

    sFrete = FRETE_FARMACIA
    sExpressao = String(STRING_MNEMONICOFPRECO_EXPRESSAO, 0)
    lErro = Comando_Executar(alComando(2), sSQL, sExpressao, FRETE_FARMACIA, giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 192418
        
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192419
    
    If lErro <> AD_SQL_SUCESSO Then gError 192420

    dFreteFarmacia = StrParaDbl(sExpressao)

    sFrete = FRETE_OUTROS
    sExpressao = String(STRING_MNEMONICOFPRECO_EXPRESSAO, 0)
    lErro = Comando_Executar(alComando(3), sSQL, sExpressao, FRETE_OUTROS, giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 192421
        
    lErro = Comando_BuscarPrimeiro(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192422
    
    If lErro <> AD_SQL_SUCESSO Then gError 192423

    dFreteOutros = StrParaDbl(sExpressao)
    
    If objPV.sNaturezaOp = "5910" Or objPV.sNaturezaOp = "6910" Then
        objPV.dValorFrete = 0
    Else

        If objPV.dValorFrete = 0 Then
    
            Select Case objTipoCliente.sDescricao
            
                Case "REDE", "DISTRIBUIDOR", "HOSPITAL"
                     objPV.dValorFrete = dFreteRede
    
                Case "FARMACIA"
                     objPV.dValorFrete = dFreteFarmacia
                     
                Case Else
                     objPV.dValorFrete = dFreteOutros
        
            End Select
            
        End If
        
    End If
    
   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
   
    Integracao_PV_Obtem_Frete_Customizado = SUCESSO
    
    Exit Function

Erro_Integracao_PV_Obtem_Frete_Customizado:

    Integracao_PV_Obtem_Frete_Customizado = gErr

    Select Case gErr
    
        Case 192414
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 192415, 192416, 192418, 192419, 192421, 192422
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MNEMONICOSFPPRECO", gErr)
            
        Case 192417, 192420, 192423
            Call Rotina_Erro(vbOKOnly, "ERRO_MNEMONICO_FRETE_NAO_CADASTRADO", gErr, sFrete, giFilialEmpresa)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192424)

    End Select
    
   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
       
    Exit Function

End Function

Public Function Integracao_PV_Obtem_UM_Customizado(ByVal objItemPV As ClassItemPedido, ByVal objProduto As ClassProduto)
'Obtém o canal de venda para a importação de pedidos de venda

Dim lErro As Long

On Error GoTo Erro_Integracao_PV_Obtem_UM_Customizado

    objItemPV.sUnidadeMed = objProduto.sSiglaUMEstoque
    objItemPV.sUMEstoque = objProduto.sSiglaUMEstoque
    objItemPV.iClasseUM = objProduto.iClasseUM
    
    Select Case Trim(objItemPV.sProduto)

        Case "5033001", "5033105"

            objItemPV.sUnidadeMed = "PCT12"

        Case "5033012", "5033015", "5033018", "5033031"

            objItemPV.sUnidadeMed = "PCT6"

    End Select
   
    Integracao_PV_Obtem_UM_Customizado = SUCESSO
    
    Exit Function

Erro_Integracao_PV_Obtem_UM_Customizado:

    Integracao_PV_Obtem_UM_Customizado = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190089)

    End Select
       
    Exit Function

End Function

Public Function Integracao_Cliente_Obtem_Tipo_Customizado(ByVal objCliente As ClassCliente, ByVal objFilialCliente As ClassFilialCliente, ByVal objEndereco As ClassEndereco, ByVal iTipo As Integer, ByVal sTipo As String)
'Obtém a natureza de operação para a importação de pedidos de venda

Dim lErro As Long
Dim bTipoCalculado As Boolean

On Error GoTo Erro_Integracao_Cliente_Obtem_Tipo_Customizado

    bTipoCalculado = True
    
    If Len(Trim(sTipo)) > 0 Then
    
        bTipoCalculado = False
    
        Select Case UCase(sTipo)
        
            Case "BUCO", "CLG", "IMPLA", "OPED", "ORTO", "PERIO"
                objCliente.iTipo = 1
                
            Case "FARMA"
                objCliente.iTipo = 2
            
            Case "REDE"
                objCliente.iTipo = 3
            
            Case "DENTAL"
                objCliente.iTipo = 4
        
            Case "HOSP"
                objCliente.iTipo = 7
        
            Case "DISTR"
                objCliente.iTipo = 8
        
            Case "CLINICA"
                objCliente.iTipo = 9
        
            Case "OUTROS"
                objCliente.iTipo = 14
        
            Case Else
                bTipoCalculado = True
        
        End Select
    
    End If
    
    If bTipoCalculado Then

        If objCliente.iTipo = 0 Then
        
            If Len(objFilialCliente.sCgc) < 14 Then
                'If objEndereco.sSiglaEstado = "SP" Then
                    objCliente.iTipo = 1
                'Else
                    'objCliente.iTipo = 1
                'End If
            Else
                'If objEndereco.sSiglaEstado = "SP" Then
                    'objCliente.iTipo = 2
                'Else
                    If Len(Trim(objFilialCliente.sInscricaoEstadual)) > 0 Then
                        objCliente.iTipo = 2
                    Else
                        objCliente.iTipo = 9
                    End If
                'End If
            End If
            
        End If
        
    End If
    
    Integracao_Cliente_Obtem_Tipo_Customizado = SUCESSO
    
    Exit Function

Erro_Integracao_Cliente_Obtem_Tipo_Customizado:

    Integracao_Cliente_Obtem_Tipo_Customizado = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190088)

    End Select
       
    Exit Function

End Function

Public Function Integracao_Cliente_Obtem_PadraoCobr_Cust(ByVal objCliente As ClassCliente, ByVal objFilialCliente As ClassFilialCliente, ByVal objEndereco As ClassEndereco)
    
    If Len(objFilialCliente.sCgc) < 14 Then
        objCliente.iCodPadraoCobranca = 2
    Else
        objCliente.iCodPadraoCobranca = 1
    End If
    
    Integracao_Cliente_Obtem_PadraoCobr_Cust = SUCESSO
    
End Function

Public Function Importa_Dados_Cliente_Arq_Funcao_Cust(sNomeFuncao As String) As Long
    sNomeFuncao = "Importa_Dados_Cliente_Arq_Xls_Phar"
    Importa_Dados_Cliente_Arq_Funcao_Cust = SUCESSO
End Function

Public Function Importa_Dados_PV_Arq_Funcao_Cust(sNomeFuncao As String) As Long
    sNomeFuncao = "Importa_Dados_PV_Arq_Xls_Phar"
    Importa_Dados_PV_Arq_Funcao_Cust = SUCESSO
End Function

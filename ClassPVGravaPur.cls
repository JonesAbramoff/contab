VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPVGrava"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function PedidoVenda_Exporta(ByVal iFilialEmpresa As Integer, ByVal lCodigo As Long) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(1 To 8) As Long
Dim sSQL As String
Dim tPV As typePedidoVenda
Dim tIPV As typeItemPedido
Dim tCliente As typeCliente
Dim tFilialCliente As typeFilialCliente
Dim tParcPV As typeParcelaPedidoVenda
Dim tEndereco As typeEndereco
Dim sRegistro As String
Dim sDiretorio As String
Dim sCodIBGE As String
Dim sArquivo As String
Dim sDir As String
Dim iItem As Integer
Dim sQuantidade As String
Dim sDesconto As String
Dim sEndereco As String
Dim sNumero As String
Dim sTelefone As String
Dim sLayout As String, iLayout As Integer, iPrazo As Integer, sTipoDoc As String
Dim iFormaPagto As Integer

On Error GoTo Erro_PedidoVenda_Exporta

    'Abertura comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 204879
    Next
    
    sDiretorio = String(STRING_CRFATCONFIG_CONTEUDO, 0)
    
    'Busca em FATConfig o código da filial de faturamento default do sistema
    lErro = Comando_Executar(alComando(1), "SELECT Conteudo FROM FATConfig WHERE Codigo = ? AND FilialEmpresa = ?", sDiretorio, "DIRETORIO_EXPORTA_PV", EMPRESA_TODA)
    If lErro <> AD_SQL_SUCESSO Then gError 204880
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 204881
        
    If lErro <> AD_SQL_SUCESSO Then gError 204882
    
    sArquivo = sDiretorio & IIf(right(sDiretorio, 1) = "\", "", "\") & "PV" & Format(lCodigo, "00000000") & ".txt"
    
    sDir = Dir(sArquivo)
    
    If Len(sDir) > 0 Then Kill sArquivo
    
    Open sArquivo For Output As #1
 
    sSQL = "SELECT         DataEmissao, Cliente, Filial, ValorTotal, TabelaPreco, CondicaoPagto FROM PedidosDeVenda " & _
           "WHERE          FilialEmpresa = ? AND Codigo = ? "
    
    lErro = Comando_Executar(alComando(2), sSQL, tPV.dtDataEmissao, tPV.lCliente, tPV.iFilial, tPV.dValorTotal, tPV.iTabelaPreco, tPV.iCondicaoPagto, iFilialEmpresa, lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 204883
        
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 204884

    If lErro = AD_SQL_SEM_DADOS Then


        sSQL = "SELECT         DataEmissao, Cliente, Filial, ValorTotal, TabelaPreco FROM PedidosDeVendaBaixados " & _
               "WHERE          FilialEmpresa = ? AND Codigo = ? "
        
        lErro = Comando_Executar(alComando(2), sSQL, tPV.dtDataEmissao, tPV.lCliente, tPV.iFilial, tPV.dValorTotal, tPV.iTabelaPreco, iFilialEmpresa, lCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 204885
            
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 204886
    
        If lErro = AD_SQL_SEM_DADOS Then gError 204887

    End If


    tCliente.sRazaoSocial = String(STRING_CLIENTE_RAZAO_SOCIAL, 0)
    tCliente.sNomeReduzido = String(STRING_CLIENTE_NOME_REDUZIDO, 0)
    tEndereco.sLogradouro = String(STRING_ENDERECO_LOGRADOURO, 0)
    tEndereco.sBairro = String(STRING_BAIRRO, 0)
    tEndereco.sCidade = String(STRING_CIDADE, 0)
    tEndereco.sSiglaEstado = String(STRING_ESTADO_SIGLA, 0)
    tEndereco.sCEP = String(STRING_CEP, 0)
    tEndereco.sTelNumero1 = String(STRING_ENDERECO_TELNUMERO1, 0)
    sCodIBGE = String(STRING_CIDADE_CODIBGE, 0)
    tFilialCliente.sCgc = String(STRING_CGC, 0)
    tFilialCliente.sInscricaoEstadual = String(STRING_INSCR_EST, 0)
    tFilialCliente.sRG = String(STRING_RG, 0)
    tEndereco.sEndereco = String(STRING_ENDERECO, 0)
    tEndereco.sTelefone1 = String(STRING_TELEFONE, 0)
    

    sSQL = "SELECT         RazaoSocial, NomeReduzido, Logradouro, Numero, Bairro, CEP, Cidade, SiglaEstado, CodIBGE, TelNumero1, CGC, InscricaoEstadual, RG, E.Endereco, Telefone1, TelDDD1  " & _
           "FROM Clientes AS C, FiliaisClientes AS F, Enderecos AS E, Cidades AS CID  WHERE C.Codigo = ? AND C.Codigo = CodCliente AND CodFilial = ? AND " & _
           "F.Endereco = E.Codigo AND Cidade = CID.Descricao"
           
    lErro = Comando_Executar(alComando(3), sSQL, tCliente.sRazaoSocial, tCliente.sNomeReduzido, tEndereco.sLogradouro, tEndereco.lNumero, tEndereco.sBairro, tEndereco.sCEP, tEndereco.sCidade, tEndereco.sSiglaEstado, sCodIBGE, tEndereco.sTelNumero1, tFilialCliente.sCgc, tFilialCliente.sInscricaoEstadual, tFilialCliente.sRG, tEndereco.sEndereco, tEndereco.sTelefone1, tEndereco.iTelDDD1, tPV.lCliente, tPV.iFilial)
    If lErro <> AD_SQL_SUCESSO Then gError 204888
        
    lErro = Comando_BuscarPrimeiro(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 204889

    If lErro = AD_SQL_SEM_DADOS Then gError 204890


    If Len(tEndereco.sLogradouro) > 0 Then
        sEndereco = tEndereco.sLogradouro
        sNumero = CStr(tEndereco.lNumero)
    Else
        sEndereco = tEndereco.sEndereco
        sNumero = "          "
    End If

    If Len(tEndereco.sTelNumero1) > 0 Then
        sTelefone = CStr(tEndereco.iTelDDD1) & tEndereco.sTelNumero1
    Else
        sTelefone = tEndereco.sTelefone1
    End If

    sRegistro = "A0|" & FormataCpoNum(tPV.lCliente, 8) & FormataCpoTexto(tCliente.sRazaoSocial, 50) & FormataCpoTexto(tCliente.sNomeReduzido, 50) & FormataCpoTexto(sEndereco, 60) & FormataCpoTexto(sNumero, 10) & FormataCpoTexto(tEndereco.sBairro, 25) & FormataCpoTexto(tEndereco.sCEP, 10) & _
                FormataCpoTexto(tEndereco.sCidade, 35) & FormataCpoTexto(tEndereco.sSiglaEstado, 2) & FormataCpoNum(sCodIBGE, 8) & FormataCpoTexto(sTelefone, 20) & FormataCpoTexto(tFilialCliente.sCgc, 18) & FormataCpoTexto(IIf(Len(tFilialCliente.sCgc) = 14, tFilialCliente.sInscricaoEstadual, tFilialCliente.sRG), 20) & IIf(Len(tFilialCliente.sCgc) = 14, "J", "F")

    Print #1, sRegistro
    
    sRegistro = "A1|" & FormataCpoNum(lCodigo, 8) & Format(tPV.dtDataEmissao, "ddmmyyyy") & FormataCpoNum(tPV.lCliente, 8) & FormataCpoValor(tPV.dValorTotal, 11) & "00000000000" & "00000000000" & FormataCpoValor(tPV.dValorTotal, 11) & FormataCpoNum(tPV.iTabelaPreco, 8)
    
    Print #1, sRegistro
    
    
    tIPV.sProduto = String(STRING_PRODUTO, 0)
    
    sSQL = "SELECT  Produto, Quantidade, PrecoUnitario, ValorDesconto, PrecoTotal  FROM ItensPedidoDeVenda " & _
           "WHERE FilialEmpresa = ? AND CodPedido = ? ORDER BY NumINtDoc"

    lErro = Comando_Executar(alComando(4), sSQL, tIPV.sProduto, tIPV.dQuantidade, tIPV.dPrecoUnitario, tIPV.dValorDesconto, tIPV.dPrecoTotal, iFilialEmpresa, lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 204891
        
    lErro = Comando_BuscarPrimeiro(alComando(4))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 204892

    If lErro = AD_SQL_SEM_DADOS Then

       tIPV.sProduto = String(STRING_PRODUTO, 0)

        sSQL = "SELECT  Produto, Quantidade, PrecoUnitario, ValorDesconto, PrecoTotal  FROM ItensPedidoDeVendaBaixados " & _
               "WHERE FilialEmpresa = ? AND CodPedido = ? ORDER BY NumIntDoc "
    
        lErro = Comando_Executar(alComando(4), sSQL, tIPV.sProduto, tIPV.dQuantidade, tIPV.dPrecoUnitario, tIPV.dValorDesconto, tIPV.dPrecoTotal, iFilialEmpresa, lCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 204893
            
        lErro = Comando_BuscarPrimeiro(alComando(4))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 204894
    
        If lErro = AD_SQL_SEM_DADOS Then gError 204895

    End If

    iItem = 0

    Do While lErro = AD_SQL_SUCESSO
    
        iItem = iItem + 1
    
        sQuantidade = Format(tIPV.dQuantidade, "00000.000")
        
        sQuantidade = left(sQuantidade, 5) & right(sQuantidade, 3)
        
        sDesconto = Format(tIPV.dValorDesconto / tIPV.dPrecoTotal * 100, "00000000.000")
        
        sDesconto = left(sDesconto, 8) & right(sDesconto, 3)
    
        sRegistro = "A2|" & FormataCpoNum(CStr(lCodigo), 8) & FormataCpoNum(CStr(iItem), 3) & FormataCpoNum(tIPV.sProduto, 8) & sQuantidade & FormataCpoValor(tIPV.dPrecoUnitario, 11) & sDesconto & "00000000000" & FormataCpoValor(tIPV.dPrecoTotal, 11)
        
        Print #1, sRegistro
        
        lErro = Comando_BuscarPrimeiro(alComando(4))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 204896
        
    Loop
    
    lErro = CF("Config_Le", "FATConfig", "PV_EXPORTA_PURAGUA_LAYOUT", 0, sLayout)
    If lErro <> SUCESSO And lErro <> 208279 Then gError ERRO_SEM_MENSAGEM
    
    iLayout = StrParaInt(sLayout)
    sTipoDoc = " "
    
    lErro = Comando_Executar(alComando(7), "SELECT FormaPagamento FROM CondicoesPagto WHERE Codigo = ?", iFormaPagto, tPV.iCondicaoPagto)
    If lErro <> AD_SQL_SUCESSO Then gError 204891
        
    lErro = Comando_BuscarPrimeiro(alComando(7))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 204892
    
    '2-Boleto
    If iFormaPagto = 2 Then sTipoDoc = "B"
    
    tIPV.sProduto = String(STRING_PRODUTO, 0)
    
    sSQL = "SELECT DataVencimento, NumParcela, Valor FROM ParcelasPedVendaTodas WHERE FilialEmpresa = ? AND PedidoDeVenda = ? ORDER BY NumParcela"

    lErro = Comando_Executar(alComando(6), sSQL, tParcPV.dtDataVencimento, tParcPV.iNumParcela, tParcPV.dValor, iFilialEmpresa, lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 204891
        
    lErro = Comando_BuscarPrimeiro(alComando(6))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 204892

    iItem = 0

    Do While lErro = AD_SQL_SUCESSO
    
        iItem = iItem + 1
       
        sRegistro = "A3|" & FormataCpoNum(CStr(lCodigo), 8) & Format(tParcPV.dtDataVencimento, "ddmmyyyy") & FormataCpoTexto(CStr(tParcPV.iNumParcela), 11) & FormataCpoValor(tParcPV.dValor, 8)
        
        If iLayout > 1 Then
        
            iPrazo = DateDiff("d", tPV.dtDataEmissao, tParcPV.dtDataVencimento)
        
            sRegistro = sRegistro & FormataCpoNum(iPrazo, 3) & sTipoDoc
        End If
        
        Print #1, sRegistro
        
        lErro = Comando_BuscarPrimeiro(alComando(6))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 204896
        
    Loop
    
    Close #1
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    PedidoVenda_Exporta = SUCESSO
    
    Exit Function

Erro_PedidoVenda_Exporta:

    PedidoVenda_Exporta = gErr

    Select Case gErr
    
        Case 75
            Call Rotina_Erro(vbOKOnly, "ERRO_SEM_ACESSO_DIRETORIO", gErr)
        
        Case 76
            Call Rotina_Erro(vbOKOnly, "ERRO_DIRETORIO_INVALIDO", gErr, sDiretorio)

        Case 204879
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 204880, 204881
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FATCONFIG2", Err, "DIRETORIO_EXPORTA_PV", EMPRESA_TODA)
            
        Case 204882
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FATCONFIG_INEXISTENTE", Err, "DIRETORIO_EXPORTA_PV", EMPRESA_TODA)
            
        Case 204883, 204884
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDODEVENDA", gErr)
        
        Case 204885, 204886
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDODEVENDABAIXADO", gErr)
        
        Case 204887
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOVENDA_NAO_CADASTRADA", gErr, lCodigo)
        
        Case 204888, 204889
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLIENTE", gErr, tPV.lCliente)
        
        Case 204890
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, tPV.lCliente)
        
        Case 204891, 204892, 204896
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMPEDIDODEVENDA", gErr)
        
        Case 204893, 204894
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMPEDIDODEVENDABAIXADO", gErr)
        
        Case 204895
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEMPEDIDOVENDA_NAO_CADASTRADO", gErr)
            
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 204897)

    End Select
    
    Close #1

   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function
    
End Function



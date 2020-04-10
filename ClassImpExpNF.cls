VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassImpExpNF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function NFiscais_Exportar(ByVal dtDataInicial As Date, ByVal dtDataFinal As Date, ByVal lNumNFDe As Long, ByVal lNumNFAte As Long, ByVal sDiretorio As String) As Long
'Gera o arquivo com as informações das NFs emitidas no intervalo acima

Dim colProdutos As New Collection
Dim lErro As Long
Dim colFornecedores As New Collection
Dim colcolFiliaisForn As New Collection
Dim colClientes As New Collection
Dim colcolFiliaisCli As New Collection
Dim colTransportadoras As New Collection
Dim colNaturezaOp As New Collection
Dim colNotasFiscais As New Collection
Dim colRegistros As New Collection
Dim colEnderecos As New Collection
Dim colcolEnderecos As New Collection
Dim iIndice As Integer
Dim colcolLancamentos As New Collection

On Error GoTo Erro_NFiscais_Exportar
    
    
    lErro = GeracaoArqNFs_Obtem_Transportadoras(colTransportadoras, colEnderecos, dtDataInicial, dtDataFinal)
    If lErro <> SUCESSO Then gError 86443
    
    lErro = Monta_Registros_Transportadoras(colRegistros, colTransportadoras, colEnderecos)
    If lErro <> SUCESSO Then gError 86444
    
    lErro = GeracaoArqNFs_Obtem_NaturezaOp(colNaturezaOp, dtDataInicial, dtDataFinal)
    If lErro <> SUCESSO Then gError 86452
    
    lErro = Monta_Registros_NaturezaOp(colRegistros, colNaturezaOp)
    If lErro <> SUCESSO Then gError 86453

    lErro = GeracaoArqNFs_Obtem_Produtos(colProdutos, dtDataInicial, dtDataFinal)
    If lErro <> SUCESSO Then gError 86454

    lErro = CF("Monta_Registros_Produtos", colRegistros, colProdutos)
    If lErro <> SUCESSO Then gError 86455
    
    lErro = GeracaoArqNFs_Obtem_Clientes_Filiais(colClientes, colcolFiliaisCli, colcolEnderecos, dtDataInicial, dtDataFinal)
    If lErro <> SUCESSO Then gError 86456
    
    lErro = CF("Monta_Registros_Clientes", colRegistros, colClientes, colcolFiliaisCli, colcolEnderecos)
    If lErro <> SUCESSO Then gError 86463

    lErro = GeracaoArqNFs_Obtem_Fornecedores_Filiais(colFornecedores, colcolFiliaisForn, colEnderecos, dtDataInicial, dtDataFinal)
    If lErro <> SUCESSO Then gError 86464

    lErro = Monta_Registros_Fornecedores(colRegistros, colFornecedores, colcolFiliaisForn, colEnderecos)
    If lErro <> SUCESSO Then gError 86471
    
    lErro = GeracaoArqNFs_Obtem_NotasFiscais(colNotasFiscais, colcolLancamentos, dtDataInicial, dtDataFinal)
    If lErro <> SUCESSO Then gError 86472

    lErro = CF("Monta_Registros_NotasFiscais", colRegistros, colNotasFiscais, colcolLancamentos)
    If lErro <> SUCESSO Then gError 86472

    lErro = GeracaoArqNFs_Cria_Arquivo(colRegistros, colNotasFiscais, sDiretorio)
    If lErro <> SUCESSO Then gError 86486
    
    NFiscais_Exportar = SUCESSO
    
    Exit Function
    
Erro_NFiscais_Exportar:

    NFiscais_Exportar = gErr
    
    Select Case gErr
    
        Case 86443, 86444, 86452, 86454, 86455, 86456, 86463, 86464, 86471, 86472, 86486
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150823)
            
    End Select
        
    Exit Function

End Function

Function GeracaoArqNFs_Cria_Arquivo(colRegistros As Collection, colNotasFiscais As Collection, ByVal sDiretorio As String) As Long

Dim lTransacao As Long, sLinha As String
Dim lErro As Long, lNumArq As Long
Dim iIndice As Integer
Dim lComando As Long, lComando2 As Long
Dim sNomeArq As String

On Error GoTo Erro_GeracaoArqNFs_Cria_Arquivo

    If colNotasFiscais.Count = 0 Then
        Call Rotina_Aviso(vbOKOnly, "AVISO_AUSENCIA_NF_TRANSFERIR")
    Else
    
        lTransacao = Transacao_Abrir()
        If lTransacao = 0 Then gError 86482
    
        lComando = Comando_Abrir()
        If lComando = 0 Then gError 86483
    
        lComando2 = Comando_Abrir()
        If lComando2 = 0 Then gError 86483
    
        sNomeArq = "TRANSFNFS_" & Format(Now, "YYYY_MM_DD_HHMMSS") & ".TXT"
        If Len(Trim(sDiretorio)) <> 0 Then
            sNomeArq = sDiretorio & IIf(right(sDiretorio, 1) = "\", "", "\") & sNomeArq
        End If
        
        'Cria o Arquivo de transferencia
        Open sNomeArq For Output As #1
        
        'Carrega as informações obtidas
        For iIndice = 1 To colRegistros.Count
            sLinha = colRegistros(iIndice)
            sLinha = Replace(sLinha, CStr(vbCr), " ")
            sLinha = Replace(sLinha, CStr(vbLf), " ")
            Print #1, sLinha
        Next
    
        'Fecha o Arquivo
        Close #1
        
        lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_PROX_ARQNFEXP", lNumArq)
        If lErro <> SUCESSO Then gError 130013
        
        lErro = Comando_Executar(lComando2, "INSERT INTO NFiscais_Transferidas_Arq ( NumArq, NomeArq ) (?,?)", lNumArq, sNomeArq)
        If lErro <> AD_SQL_SUCESSO Then gError 86484
        
        For iIndice = 1 To colNotasFiscais.Count
        
            lErro = Comando_Executar(lComando, "INSERT INTO NFiscais_Transferidas (NumIntNFiscal, NumArq) VALUES (?,?) ", colNotasFiscais(iIndice).lNumIntDoc, lNumArq)
            If lErro <> AD_SQL_SUCESSO Then gError 86484
            
        Next
        
        lErro = Transacao_Commit()
        If lErro <> SUCESSO Then gError 86485
        
        Call Comando_Fechar(lComando)
        Call Comando_Fechar(lComando2)
        
        Call Rotina_Aviso(vbOKOnly, "AVISO_ARQNFS_GERADO_SUCESSO", sNomeArq)
    
    End If
    
    GeracaoArqNFs_Cria_Arquivo = SUCESSO
    
    Exit Function

Erro_GeracaoArqNFs_Cria_Arquivo:

    GeracaoArqNFs_Cria_Arquivo = gErr
    
    Select Case gErr
    
        Case 130013
        
        Case 86482
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 86483
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 86484
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_NFISCAIS_TRANSFERIDAS", gErr)
        
        Case 86485
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT_TRANSACAO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150824)

    End Select
    
    Close #1
    
    Call Kill(sNomeArq)
    
    Call Transacao_Rollback
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    
    Exit Function

End Function

Function GeracaoArqNFs_Obtem_Transportadoras(colTransportadoras As Collection, colEnderecos As Collection, dtDataInicial As Date, dtDataFinal As Date) As Long

Dim lErro As Long
Dim lComando As Long
Dim objTransportadora As ClassTransportadora
Dim tTransportadora As typeTransportadora
Dim objEndereco As ClassEndereco
Dim tEndereco As typeEndereco
Dim sSelect As String

On Error GoTo Erro_GeracaoArqNFs_Obtem_Transportadoras

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 86437
    
    tTransportadora.sCgc = String(STRING_CGC, 0)
    tTransportadora.sGuia = String(STRING_TRANSPORTADORA_GUIA, 0)
    tTransportadora.sInscricaoEstadual = String(STRING_INSCR_EST, 0)
    tTransportadora.sInscricaoMunicipal = String(STRING_INSCR_MUN, 0)
    tTransportadora.sNome = String(STRING_TRANSPORTADORA_NOME, 0)
    tTransportadora.sNomeReduzido = String(STRING_TRANSPORTADORA_NOME_REDUZIDO, 0)
    tEndereco.sBairro = String(STRING_BAIRRO, 0)
    tEndereco.sCEP = String(STRING_CEP, 0)
    tEndereco.sCidade = String(STRING_CIDADE, 0)
    tEndereco.sContato = String(STRING_CONTATO, 0)
    tEndereco.sEmail = String(STRING_EMAIL, 0)
    tEndereco.sEndereco = String(STRING_ENDERECO, 0)
    tEndereco.sFax = String(STRING_FAX, 0)
    tEndereco.sSiglaEstado = String(STRING_ESTADO, 0)
    tEndereco.sTelefone1 = String(STRING_TELEFONE, 0)
    tEndereco.sTelefone2 = String(STRING_TELEFONE, 0)

    With tTransportadora
    
        sSelect = "SELECT DISTINCT CodTransp, Nome,NomeReduzido, T2.CGC, T2.InscricaoEstadual, ViaTransporte, T2.InscricaoMunicipal, Guia, PesoMinimo, Enderecos.Endereco, Enderecos.Bairro, Enderecos.Cidade, Enderecos.SiglaEstado, Enderecos.CodigoPais, Enderecos.CEP, Enderecos.Telefone1, Enderecos.Telefone2, Enderecos.Email, Enderecos.Fax, Enderecos.Contato FROM TransportadoraHistorico T1, Transportadoras T2, NFiscal, Enderecos WHERE T2.Endereco = Enderecos.Codigo AND NFiscal.CodTransportadora = T2.Codigo AND T2.Codigo = Codtransp AND T1.DataAtualizacao BETWEEN ? AND ? AND DataEmissao BETWEEN ? AND ? AND NFiscal.NumIntDoc NOT IN (SELECT NumIntNFiscal FROM NFiscais_Transferidas) UNION " & _
            "SELECT DISTINCT CodTransp, Nome,NomeReduzido, T2.CGC, T2.InscricaoEstadual, ViaTransporte, T2.InscricaoMunicipal, Guia, PesoMinimo, Enderecos.Endereco, Enderecos.Bairro, Enderecos.Cidade, Enderecos.SiglaEstado, Enderecos.CodigoPais, Enderecos.CEP, Enderecos.Telefone1, Enderecos.Telefone2, Enderecos.Email, Enderecos.Fax, Enderecos.Contato FROM TransportadoraHistorico T1, Transportadoras T2, NFiscal, Enderecos WHERE T2.Endereco = Enderecos.Codigo AND NFiscal.CodTranspRedesp = T2.Codigo AND T2.Codigo = Codtransp AND T1.DataAtualizacao BETWEEN ? AND ? AND DataEmissao BETWEEN ? AND ? AND NFiscal.NumIntDoc NOT IN (SELECT NumIntNFiscal FROM NFiscais_Transferidas)"
        lErro = Comando_Executar(lComando, sSelect, .iCodigo, .sNome, .sNomeReduzido, .sCgc, .sInscricaoEstadual, .iViaTransporte, .sInscricaoMunicipal, .sGuia, .dPesoMinimo, tEndereco.sEndereco, tEndereco.sBairro, tEndereco.sCidade, tEndereco.sSiglaEstado, tEndereco.iCodigoPais, tEndereco.sCEP, tEndereco.sTelefone1, tEndereco.sTelefone2, tEndereco.sEmail, tEndereco.sFax, tEndereco.sContato, dtDataInicial, dtDataFinal, dtDataInicial, dtDataFinal, dtDataInicial, dtDataFinal, dtDataInicial, dtDataFinal)
        If lErro <> AD_SQL_SUCESSO Then gError 86438
        
    End With
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 86439
    
    
    Set colEnderecos = New Collection
    
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objTransportadora = New ClassTransportadora
        
        objTransportadora.dPesoMinimo = tTransportadora.dPesoMinimo
        objTransportadora.iCodigo = tTransportadora.iCodigo
        objTransportadora.iViaTransporte = tTransportadora.iViaTransporte
        objTransportadora.lEndereco = tTransportadora.lEndereco
        objTransportadora.sCgc = tTransportadora.sCgc
        objTransportadora.sGuia = tTransportadora.sGuia
        objTransportadora.sInscricaoEstadual = tTransportadora.sInscricaoEstadual
        objTransportadora.sInscricaoMunicipal = tTransportadora.sInscricaoMunicipal
        objTransportadora.sNome = tTransportadora.sNome
        objTransportadora.sNomeReduzido = tTransportadora.sNomeReduzido
        
        Set objEndereco = New ClassEndereco
        
        objEndereco.iCodigoPais = tEndereco.iCodigoPais
        objEndereco.sBairro = tEndereco.sBairro
        objEndereco.sCEP = tEndereco.sCEP
        objEndereco.sCidade = tEndereco.sCidade
        objEndereco.sContato = tEndereco.sContato
        objEndereco.sEmail = tEndereco.sEmail
        objEndereco.sEndereco = tEndereco.sEndereco
        objEndereco.sFax = tEndereco.sFax
        objEndereco.sSiglaEstado = tEndereco.sSiglaEstado
        objEndereco.sTelefone1 = tEndereco.sTelefone1
        objEndereco.sTelefone2 = tEndereco.sTelefone2
        
        colTransportadoras.Add objTransportadora
        colEnderecos.Add objEndereco
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 86439
        
    Loop
        
    Call Comando_Fechar(lComando)
    
    Exit Function
    
    GeracaoArqNFs_Obtem_Transportadoras = SUCESSO

Erro_GeracaoArqNFs_Obtem_Transportadoras:

    GeracaoArqNFs_Obtem_Transportadoras = gErr
    
    Select Case gErr
    
        Case 86437
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 86438, 86439
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRANSPORTADORA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150825)
            
    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function GeracaoArqNFs_Obtem_NaturezaOp(colNaturezaOp As Collection, dtDataInicial As Date, dtDataFinal As Date) As Long

Dim lErro As Long
Dim lComando As Long
Dim objNaturezaOp As ClassNaturezaOp
Dim sCodigo As String
Dim sDescricao As String
Dim sDescricaoNF As String

On Error GoTo Erro_GeracaoArqNFs_Obtem_NaturezaOp

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 86440
    
    sCodigo = String(STRING_NATUREZAOP_CODIGO, 0)
    sDescricao = String(STRING_NATUREZAOP_DESCRICAO, 0)
    sDescricaoNF = String(STRING_NATUREZAOP_DESCRNF, 0)

    lErro = Comando_Executar(lComando, "SELECT DISTINCT Codigo, Descricao, DescrNF FROM NaturezaOpHistorico, NaturezaOp, NFiscal WHERE NaturezaOp = Codigo  AND NFiscal.NumIntDoc NOT IN (SELECT NumIntNFiscal FROM NFiscais_Transferidas) AND Codigo = CodNaturezaOp AND DataEmissao BETWEEN ? AND ? AND NatuRezaOpHistorico.DataAtualizacao BETWEEN ? AND ? ", sCodigo, sDescricao, sDescricaoNF, dtDataInicial, dtDataFinal, dtDataInicial, dtDataFinal)
    If lErro <> AD_SQL_SUCESSO Then gError 86441
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 86442
    
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objNaturezaOp = New ClassNaturezaOp
        
        objNaturezaOp.sCodigo = sCodigo
        objNaturezaOp.sDescricao = sDescricao
        objNaturezaOp.sDescrNF = sDescricaoNF
        
        colNaturezaOp.Add objNaturezaOp
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 86442
        
    Loop
        
    Call Comando_Fechar(lComando)
    
    Exit Function
    
    GeracaoArqNFs_Obtem_NaturezaOp = SUCESSO

Erro_GeracaoArqNFs_Obtem_NaturezaOp:

    GeracaoArqNFs_Obtem_NaturezaOp = gErr
    
    Select Case gErr
    
        Case 86440
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 86441, 86442
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NATUREZAOP1", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150826)
            
    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function GeracaoArqNFs_Obtem_Produtos(colProdutos As Collection, dtDataInicial As Date, dtDataFinal As Date) As Long

Dim lErro As Long
Dim objProduto As ClassProduto
Dim objProdutoFilial As ClassProdutoFilial
Dim tProduto As typeProduto
Dim sCodigo As String
Dim lComando As Long

On Error GoTo Erro_GeracaoArqNFs_Obtem_Produtos

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 86444
    
    sCodigo = String(STRING_PRODUTO, 0)

    lErro = Comando_Executar(lComando, "SELECT DISTINCT Codigo FROM ProdutoHistorico, Produtos WHERE CodProduto = Codigo AND DataAtualizacao BETWEEN ? AND ? AND (Codigo IN (SELECT Produto FROM ItensNFiscal, NFiscal WHERE ItensNFiscal.NumIntNF = NFiscal.NumintDoc AND DataEmissao BETWEEN ? AND ?  AND NFiscal.NumIntDoc NOT IN (SELECT NumIntNFiscal FROM NFiscais_Transferidas)) OR Codigo IN (SELECT ItensNFiscalGrade.Produto FROM ItensNFiscalGrade, ItensNFiscal, NFiscal WHERE ItensNFiscal.NumIntNF = NFiscal.NumintDoc AND ItensNFiscalGrade.NumIntItemNF = ItensNfiscal.NumIntDoc AND DataEmissao BETWEEN ? AND ?  AND NFiscal.NumIntDoc NOT IN (SELECT NumIntNFiscal FROM NFiscais_Transferidas) ))", sCodigo, dtDataInicial, dtDataFinal, dtDataInicial, dtDataFinal, dtDataInicial, dtDataFinal)
    If lErro <> AD_SQL_SUCESSO Then gError 86445
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 86446
    
    Do While lErro = AD_SQL_SUCESSO
    
        lErro = GeracaoArqNFs_Obtem_Produtos1(colProdutos, sCodigo)
        If lErro <> SUCESSO Then gError 130008
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 86451
        
    Loop
    
    Call Comando_Fechar(lComando)
    
    GeracaoArqNFs_Obtem_Produtos = SUCESSO
    
    Exit Function
    
Erro_GeracaoArqNFs_Obtem_Produtos:

    GeracaoArqNFs_Obtem_Produtos = gErr
    
    Select Case gErr
        
        Case 86444
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 86445, 86446, 86451
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS1", gErr)
            
        Case 86447, 86449, 86450, 130008
        
        Case 86448
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150827)
            
    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function GeracaoArqNFs_Obtem_Clientes_Filiais(colClientes As Collection, colcolFiliaisCli As Collection, colcolEnderecos As Collection, dtDataInicial As Date, dtDataFinal As Date) As Long

Dim lErro As Long
Dim lCodigo As Long
Dim objCliente As ClassCliente
Dim colFiliais As Collection
Dim tFilialCliente As typeFilialCliente
Dim objFilialCliente As ClassFilialCliente
Dim lComando As Long
Dim iFilial As Integer
Dim colEnderecos As Collection
Dim lEndereco As Long
Dim lEnderecoEntrega As Long
Dim lEnderecoCobranca As Long

On Error GoTo Erro_GeracaoArqNFs_Obtem_Clientes_Filiais

    Set colcolEnderecos = New Collection

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 86457
    
    lErro = Comando_Executar(lComando, "SELECT DISTINCT Codigo, Filial FROM ClienteHistorico, NFiscal WHERE ClienteHistorico.DataAtualizacao BETWEEN ? AND ? AND Codigo = Cliente  AND Filial = FilialCli AND DataEmissao BETWEEN ? AND ? AND NFiscal.NumIntDoc NOT IN (SELECT NumIntNFiscal FROM NFiscais_Transferidas) ORDER BY Codigo, Filial", lCodigo, iFilial, dtDataInicial, dtDataFinal, dtDataInicial, dtDataFinal)
    If lErro <> AD_SQL_SUCESSO Then gError 86458

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 86459
   
    Do While lErro = AD_SQL_SUCESSO
       
        Set objCliente = New ClassCliente
        
        objCliente.lCodigo = lCodigo
        
        lErro = CF("Cliente_Le", objCliente)
        If lErro <> SUCESSO And lErro <> 12293 Then gError 86460
        If lErro <> SUCESSO Then gError 86461
    
        lEndereco = objCliente.lEndereco
        lEnderecoCobranca = objCliente.lEnderecoCobranca
        lEnderecoEntrega = objCliente.lEnderecoEntrega
    
        Set colFiliais = New Collection
        Set colEnderecos = New Collection
        
        'mesmo que a matriz do cliente nao tenha tido nf ela tem que ser cadastrada antes das suas filiais
        If iFilial <> FILIAL_MATRIZ Then
        
            '??? testar se há clientehistorico no periodo para a matriz
            
            lErro = GeracaoArqNFs_FilCli_TrataEnderecos(lEndereco, lEnderecoCobranca, lEnderecoEntrega, colEnderecos)
            If lErro <> SUCESSO Then gError 130011
        
        End If
        
        Do While lCodigo = objCliente.lCodigo And lErro = AD_SQL_SUCESSO
        
            If iFilial <> FILIAL_MATRIZ Then
            
                Set objFilialCliente = New ClassFilialCliente
                
                objFilialCliente.lCodCliente = lCodigo
                objFilialCliente.iCodFilial = iFilial
                
                lErro = CF("FilialCliente_Le", objFilialCliente)
                If lErro <> SUCESSO Then Error 86462
                
                colFiliais.Add objFilialCliente
                
                lEndereco = objFilialCliente.lEndereco
                lEnderecoCobranca = objFilialCliente.lEnderecoCobranca
                lEnderecoEntrega = objFilialCliente.lEnderecoEntrega
                
                lErro = GeracaoArqNFs_FilCli_TrataEnderecos(lEndereco, lEnderecoCobranca, lEnderecoEntrega, colEnderecos)
                If lErro <> SUCESSO Then gError 130012
            
            End If
            
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 86463
        
        Loop
        
        colClientes.Add objCliente
        colcolFiliaisCli.Add colFiliais
        colcolEnderecos.Add colEnderecos
            
    Loop

    Call Comando_Fechar(lComando)

    GeracaoArqNFs_Obtem_Clientes_Filiais = SUCESSO
    
    Exit Function
    
Erro_GeracaoArqNFs_Obtem_Clientes_Filiais:

    GeracaoArqNFs_Obtem_Clientes_Filiais = gErr
    
    Select Case gErr
    
        Case 86457
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 86458, 86459, 86463
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLIENTES", gErr)
    
        Case 86460, 86462, 130011, 130012
        
        Case 86461
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, objCliente.lCodigo)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150828)
            
    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function GeracaoArqNFs_Obtem_Fornecedores_Filiais(colFornecedores As Collection, colcolFiliaisCli As Collection, colEnderecos As Collection, dtDataInicial As Date, dtDataFinal As Date) As Long

Dim lErro As Long
Dim lCodigo As Long
Dim objFornecedor As ClassFornecedor
Dim colFiliais As Collection
Dim objFilialFornecedor As ClassFilialFornecedor
Dim lComando As Long
Dim iFilial As Integer
Dim objEndereco As ClassEndereco
Dim lEndereco As Long

On Error GoTo Erro_GeracaoArqNFs_Obtem_Fornecedores_Filiais

    Set colEnderecos = New Collection

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 86465
    
    lErro = Comando_Executar(lComando, "SELECT DISTINCT Codigo, Filial FROM FornecedorHistorico, NFiscal WHERE FornecedorHistorico.DataAtualizacao BETWEEN ? AND ? AND Codigo = Fornecedor  AND Filial = FilialForn AND DataEmissao BETWEEN ? AND ? AND NFiscal.NumIntDoc NOT IN (SELECT NumIntNFiscal FROM NFiscais_Transferidas) ORDER BY Codigo, Filial", lCodigo, iFilial, dtDataInicial, dtDataFinal, dtDataInicial, dtDataFinal)
    If lErro <> AD_SQL_SUCESSO Then gError 86466

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 86467

    Set colFiliais = New Collection
    
    Do While lErro = AD_SQL_SUCESSO
       
        Set objFornecedor = New ClassFornecedor
        
        objFornecedor.lCodigo = lCodigo
        
        lErro = CF("Fornecedor_Le", objFornecedor)
        If lErro <> SUCESSO And lErro <> 12732 Then gError 86468
        If lErro <> SUCESSO Then gError 86469
        
        lEndereco = objFornecedor.lEndereco
        
        If iFilial <> FILIAL_MATRIZ Then
        
            Set objEndereco = New ClassEndereco
            
            objEndereco.lCodigo = lEndereco
            
            lErro = CF("Endereco_Le", objEndereco)
            If lErro <> SUCESSO And lErro <> 12309 Then gError 86489
            If lErro <> SUCESSO Then gError 86490
            
            colEnderecos.Add objEndereco
        
        End If
        
        Set colFiliais = New Collection
        
        Do While lCodigo = objFornecedor.lCodigo And lErro = AD_SQL_SUCESSO
               
            If iFilial <> FILIAL_MATRIZ Then
            
                Set objFilialFornecedor = New ClassFilialFornecedor
                
                objFilialFornecedor.lCodFornecedor = lCodigo
                objFilialFornecedor.iCodFilial = iFilial
                
                lErro = CF("FilialFornecedor_Le", objFilialFornecedor)
                If lErro <> SUCESSO Then gError 86470
                        
                colFiliais.Add objFilialFornecedor
                
                lEndereco = objFilialFornecedor.lEndereco
            
                Set objEndereco = New ClassEndereco
                
                objEndereco.lCodigo = lEndereco
                
                lErro = CF("Endereco_Le", objEndereco)
                If lErro <> SUCESSO And lErro <> 12309 Then gError 86489
                If lErro <> SUCESSO Then gError 86490
                
                colEnderecos.Add objEndereco
            
            End If
            
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 11114
        
        Loop
        
        colFornecedores.Add objFornecedor
        colcolFiliaisCli.Add colFiliais
            
    Loop

    Call Comando_Fechar(lComando)

    GeracaoArqNFs_Obtem_Fornecedores_Filiais = SUCESSO
    
    Exit Function
    
Erro_GeracaoArqNFs_Obtem_Fornecedores_Filiais:

    GeracaoArqNFs_Obtem_Fornecedores_Filiais = gErr
    
    Select Case gErr
        
        Case 86465
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 86466, 86467
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FORNECEDORES", gErr)
        
        Case 86468, 86470, 86489
        
        Case 86469
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO", gErr, objFornecedor.lCodigo)
            
        Case 86490
            Call Rotina_Erro(vbOKOnly, "ERRO_ENDERECO_NAO_CADASTRADO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150829)
            
    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function GeracaoArqNFs_Obtem_NotasFiscais(colNotasFiscais As Collection, colcolLancamentos As Collection, dtDataInicial As Date, dtDataFinal As Date) As Long

Dim lErro As Long
Dim lComando As Long
Dim tNFiscal As typeNFiscal
Dim objNFiscal As ClassNFiscal
Dim colLancamentos As Collection
Dim colLanPendente As Collection

On Error GoTo Erro_GeracaoArqNFs_Obtem_NotasFiscais

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 86473
    
    With tNFiscal
    
        .sPlaca = String(STRING_NFISCAL_PLACA, 0)
        .sPlacaUF = String(STRING_NFISCAL_PLACA_UF, 0)
        .sSerie = String(STRING_SERIE, 0)
        .sNumPedidoTerc = String(STRING_NUM_PEDIDO_TERC, 0)
        .sMensagemNota = String(STRING_NFISCAL_MENSAGEM, 0)
        .sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
        .sVolumeNumero = String(STRING_NFISCAL_VOLUME_NUMERO, 0)
        .sObservacao = String(STRING_NFISCAL_OBSERVACAO, 0)
        .sCodUsuarioCancel = String(STRING_USUARIO_CODIGO, 0)
        .sMotivoCancel = String(STRING_NFISCAL_MOTIVOCANCEL, 0)

        '??? obs: estou evitando recebimentos
        lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Status, FilialEmpresa, Serie, NumNotaFiscal, Cliente, FilialCli, FilialEntrega, Fornecedor, FilialForn, DataEmissao, DataEntrada, DataSaida, DataVencimento, DataReferencia, FilialPedido,  NumPedidoTerc, ClasseDocCPR, NumIntDocCPR, ValorTotal, ValorProdutos, ValorFrete, ValorSeguro, ValorOutrasDespesas, ValorDesconto, CodTransportadora, MensagemNota, TabelaPreco, TipoNFiscal, NaturezaOp, PesoLiq, PesoBruto, NumIntTrib, Placa, PlacaUF, VolumeQuant, VolumeEspecie, VolumeMarca, VolumeNumero, Canal, NumIntNotaOriginal, ClienteBenef, FilialCliBenef, FornecedorBenef, FilialFornBenef, FreteRespons, NumRecebimento, Observacao, CodUsuarioCancel, MotivoCancel, HoraEntrada, HoraSaida, CodTranspRedesp, DetPagFrete, SemDataSaida, NaoImpCobranca FROM NFiscal, TiposDocInfo WHERE NFiscal.TipoNFiscal = TiposDocInfo.Codigo AND DataEmissao BETWEEN ? AND ? AND TiposDocInfo.Tipo IN (1,2,3) AND " & _
            "NumIntDoc NOT IN (SELECT NumIntNFiscal FROM NFiscais_Transferidas) ORDER BY NumIntDoc", _
            .lNumIntDoc, .iStatus, .iFilialEmpresa, .sSerie, .lNumNotaFiscal, .lCliente, .iFilialCli, .iFilialEntrega, .lFornecedor, .iFilialForn, .dtDataEmissao, .dtDataEntrada, .dtDataSaida, .dtDataVencimento, .dtDataReferencia, .iFilialPedido, .sNumPedidoTerc, .iClasseDocCPR, .lNumIntDocCPR, .dValorTotal, .dValorProdutos, .dValorFrete, .dValorSeguro, .dValorOutrasDespesas, .dValorDesconto, .iCodTransportadora, .sMensagemNota, .iTabelaPreco, .iTipoNFiscal, .sNaturezaOp, .dPesoLiq, .dPesoBruto, .lNumIntTrib, .sPlaca, .sPlacaUF, .lVolumeQuant, .lVolumeEspecie, .lVolumeMarca, .sVolumeNumero, .iCanal, .lNumIntNotaOriginal, .lClienteBenef, .iFilialCliBenef, .lFornecedorBenef, .iFilialFornBenef, .iFreteRespons, .lNumRecebimento, .sObservacao, .sCodUsuarioCancel, .sMotivoCancel, .dHoraEntrada, .dHoraSaida, .iCodTranspRedesp, .iDetPagFrete, .iSemDataSaida, .iNaoImpCobranca, dtDataInicial, dtDataFinal)
        If lErro <> AD_SQL_SUCESSO Then gError 86474
        
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then gError 86475
        
        Do While lErro = AD_SQL_SUCESSO
            
            Set objNFiscal = New ClassNFiscal
            Set colLancamentos = New Collection
            Set colLanPendente = New Collection
            
            objNFiscal.lNumIntDoc = .lNumIntDoc
            objNFiscal.iStatus = .iStatus
            objNFiscal.iFilialEmpresa = .iFilialEmpresa
            objNFiscal.sSerie = .sSerie
            objNFiscal.lNumNotaFiscal = .lNumNotaFiscal
            objNFiscal.lCliente = .lCliente
            objNFiscal.iFilialCli = .iFilialCli
            objNFiscal.iFilialEntrega = .iFilialEntrega
            objNFiscal.lFornecedor = .lFornecedor
            objNFiscal.iFilialForn = .iFilialForn
            objNFiscal.dtDataEmissao = .dtDataEmissao
            objNFiscal.dtDataEntrada = .dtDataEntrada
            objNFiscal.dtDataSaida = .dtDataSaida
            objNFiscal.dtDataVencimento = .dtDataVencimento
            objNFiscal.dtDataReferencia = .dtDataReferencia
            objNFiscal.iFilialPedido = .iFilialPedido
            objNFiscal.lNumPedidoVenda = .lNumPedidoVenda
            objNFiscal.sNumPedidoTerc = .sNumPedidoTerc
            objNFiscal.iClasseDocCPR = .iClasseDocCPR
            objNFiscal.lNumIntDocCPR = .lNumIntDocCPR
            objNFiscal.dValorTotal = .dValorTotal
            objNFiscal.dValorProdutos = .dValorProdutos
            objNFiscal.dValorFrete = .dValorFrete
            objNFiscal.dValorSeguro = .dValorSeguro
            objNFiscal.dValorOutrasDespesas = .dValorOutrasDespesas
            objNFiscal.dValorDesconto = .dValorDesconto
            objNFiscal.iCodTransportadora = .iCodTransportadora
            objNFiscal.sMensagemNota = .sMensagemNota
            objNFiscal.iTabelaPreco = .iTabelaPreco
            objNFiscal.iTipoNFiscal = .iTipoNFiscal
            objNFiscal.sNaturezaOp = .sNaturezaOp
            objNFiscal.dPesoLiq = .dPesoLiq
            objNFiscal.dPesoBruto = .dPesoBruto
            objNFiscal.lNumIntTrib = .lNumIntTrib
            objNFiscal.sPlaca = .sPlaca
            objNFiscal.sPlacaUF = .sPlacaUF
            objNFiscal.lVolumeQuant = .lVolumeQuant
            objNFiscal.lVolumeEspecie = .lVolumeEspecie
            objNFiscal.lVolumeMarca = .lVolumeMarca
            objNFiscal.sVolumeNumero = .sVolumeNumero
            objNFiscal.iCanal = .iCanal
            objNFiscal.lNumIntNotaOriginal = .lNumIntNotaOriginal
            objNFiscal.lClienteBenef = .lClienteBenef
            objNFiscal.iFilialCliBenef = .iFilialCliBenef
            objNFiscal.lFornecedorBenef = .lFornecedorBenef
            objNFiscal.iFilialFornBenef = .iFilialFornBenef
            objNFiscal.iFreteRespons = .iFreteRespons
            objNFiscal.lNumRecebimento = .lNumRecebimento
            objNFiscal.sObservacao = .sObservacao
            objNFiscal.sCodUsuarioCancel = .sCodUsuarioCancel
            objNFiscal.sMotivoCancel = .sMotivoCancel
            objNFiscal.dtHoraEntrada = .dHoraEntrada
            objNFiscal.dtHoraSaida = .dHoraSaida
            objNFiscal.iCodTranspRedesp = .iCodTranspRedesp
            objNFiscal.iDetPagFrete = .iDetPagFrete
            objNFiscal.iSemDataSaida = .iSemDataSaida
            objNFiscal.iNaoImpCobranca = .iNaoImpCobranca
            
            colNotasFiscais.Add objNFiscal
            
            lErro = CF("NFiscalItens_Le", objNFiscal)
            If lErro <> SUCESSO Then gError 86476
            
            lErro = CF("NFiscal_Le_Tributacao", objNFiscal)
            If lErro <> SUCESSO Then gError 86477

            lErro = CF("NFiscalItens_Le_Trib", objNFiscal)
            If lErro <> SUCESSO Then gError 86478

            lErro = CF("AlocacoesNF_Le", objNFiscal)
            If lErro <> SUCESSO Then gError 86479
        
            lErro = CF("ParcelasRecNF_Le", objNFiscal)
            If lErro <> SUCESSO Then gError 86480
            
            lErro = CF("ParcelasPagNF_Le", objNFiscal)
            If lErro <> SUCESSO Then gError 86481
                        
            lErro = NFiscal_ObtemCondPagto(objNFiscal)
            If lErro <> SUCESSO Then gError 86481
            
            lErro = GeracaoArqNFs_Le_LancamentosCTBNF(objNFiscal, colLancamentos, colLanPendente)
            If lErro <> SUCESSO Then gError 86481
            
            colcolLancamentos.Add colLancamentos
            colcolLancamentos.Add colLanPendente
                        
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 11116
                        
        Loop
        
    End With
    
    Call Comando_Fechar(lComando)
    
    GeracaoArqNFs_Obtem_NotasFiscais = SUCESSO
    
    Exit Function

Erro_GeracaoArqNFs_Obtem_NotasFiscais:

    GeracaoArqNFs_Obtem_NotasFiscais = gErr
    
    Select Case gErr
    
        Case 86473
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 86474, 86475
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", gErr)
    
        Case 86476 To 86481
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150830)
            
    End Select
    
    Call Comando_Fechar(lComando)
        
    Exit Function
        
End Function

Function Monta_Registros_Transportadoras(colRegistros As Collection, colTransportadoras As Collection, colEnderecos As Collection) As Long

Dim objTransportadora As ClassTransportadora
Dim sRegistro As String
Dim iIndice As Integer
Dim objEndereco As ClassEndereco

On Error GoTo Erro_Monta_Registros_Transportadoras

    iIndice = 0
    For Each objTransportadora In colTransportadoras
        
        iIndice = iIndice + 1
        
        Set objEndereco = colEnderecos(iIndice)
        
        With objTransportadora
            sRegistro = TIPO_REGISTRO_TRANSPORTADORA & Chr(vbKeyEscape) & .iCodigo & Chr(vbKeyEscape) & .sNome & Chr(vbKeyEscape) & .sNomeReduzido & Chr(vbKeyEscape) & .sCgc & Chr(vbKeyEscape) & .sInscricaoEstadual & Chr(vbKeyEscape) & .iViaTransporte & Chr(vbKeyEscape) & .sInscricaoMunicipal & Chr(vbKeyEscape) & .sGuia & Chr(vbKeyEscape) & .dPesoMinimo & Chr(vbKeyEscape) & objEndereco.sEndereco & Chr(vbKeyEscape) & objEndereco.sBairro & Chr(vbKeyEscape) & objEndereco.sCidade & Chr(vbKeyEscape) & objEndereco.sSiglaEstado & Chr(vbKeyEscape) & objEndereco.iCodigoPais & Chr(vbKeyEscape) & objEndereco.sCEP & Chr(vbKeyEscape) & objEndereco.sTelefone1 & Chr(vbKeyEscape) & objEndereco.sTelefone2 & Chr(vbKeyEscape) & objEndereco.sEmail & Chr(vbKeyEscape) & objEndereco.sFax & Chr(vbKeyEscape) & objEndereco.sContato & Chr(vbKeyEnd)
        End With
        
        colRegistros.Add sRegistro
        
    Next

    Monta_Registros_Transportadoras = SUCESSO
    
    Exit Function

Erro_Monta_Registros_Transportadoras:

    Monta_Registros_Transportadoras = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150831)
    
    End Select
    
    Exit Function
    
End Function

Function Monta_Registros_NaturezaOp(colRegistros As Collection, colNaturezaOp As Collection) As Long

Dim objNaturezaOp As ClassNaturezaOp
Dim sRegistro As String

On Error GoTo Erro_Monta_Registros_NaturezaOp

    For Each objNaturezaOp In colNaturezaOp
        
        With objNaturezaOp
        
            sRegistro = TIPO_REGISTRO_NATUREZAOP & Chr(vbKeyEscape) & .sCodigo & Chr(vbKeyEscape) & .sDescricao & Chr(vbKeyEscape) & .sDescrNF & Chr(vbKeyEnd)
        
        End With
        
        colRegistros.Add sRegistro
        
    Next

    Monta_Registros_NaturezaOp = SUCESSO

    Exit Function

Erro_Monta_Registros_NaturezaOp:

    Monta_Registros_NaturezaOp = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150832)
    
    End Select
    
    Exit Function

End Function

Function Monta_Registros_Produtos(colRegistros As Collection, colProdutos As Collection) As Long

Dim objProduto As ClassProduto
Dim objProdutoCateg As ClassProdutoCategoria
Dim iIndice As Integer
Dim sRegistro As String

On Error GoTo Erro_Monta_Registros_Produtos

    iIndice = 0

    For Each objProduto In colProdutos
        
        iIndice = iIndice + 1
        With objProduto
        
            sRegistro = TIPO_REGISTRO_PRODUTO & Chr(vbKeyEscape) & .sCodigo & Chr(vbKeyEscape) & .sDescricao & Chr(vbKeyEscape) & .iGerencial & Chr(vbKeyEscape) & .iAtivo & Chr(vbKeyEscape) & .iFaturamento & Chr(vbKeyEscape) & .iClasseUM & Chr(vbKeyEscape) & .sSiglaUMEstoque & Chr(vbKeyEscape) & _
            .sSiglaUMCompra & Chr(vbKeyEscape) & .sSiglaUMVenda & Chr(vbKeyEscape) & .sSubstituto1 & Chr(vbKeyEscape) & .sSubstituto2 & Chr(vbKeyEscape) & .iNivel & Chr(vbKeyEscape) & .iTipo & Chr(vbKeyEscape) & .sNomeReduzido & Chr(vbKeyEscape) & .sModelo & Chr(vbKeyEscape) & .iPrazoValidade & Chr(vbKeyEscape) & .sCodigoBarras & Chr(vbKeyEscape) & _
            .iEtiquetasCodBarras & Chr(vbKeyEscape) & .dPesoLiq & Chr(vbKeyEscape) & .dPesoBruto & Chr(vbKeyEscape) & .dComprimento & Chr(vbKeyEscape) & .dEspessura & Chr(vbKeyEscape) & .dLargura & Chr(vbKeyEscape) & .sCor & Chr(vbKeyEscape) & .sObsFisica & Chr(vbKeyEscape) & .iCompras & Chr(vbKeyEscape) & .iPCP & Chr(vbKeyEscape) & .iKitBasico & Chr(vbKeyEscape) & _
            .iKitInt & Chr(vbKeyEscape) & .dIPIAliquota & Chr(vbKeyEscape) & .sIPICodigo & Chr(vbKeyEscape) & .sIPICodDIPI & Chr(vbKeyEscape) & .iControleEstoque & Chr(vbKeyEscape) & .iApropriacaoCusto & Chr(vbKeyEscape) & .sContaContabil & Chr(vbKeyEscape) & .sContaContabilProducao & Chr(vbKeyEscape) & .dResiduo & Chr(vbKeyEscape) & _
            .iNatureza & Chr(vbKeyEscape) & .dCustoReposicao & Chr(vbKeyEscape) & .iOrigemMercadoria & Chr(vbKeyEscape) & .iTabelaPreco & Chr(vbKeyEscape) & .dPercentMaisQuantCotAnt & Chr(vbKeyEscape) & .dPercentMenosQuantCotAnt & Chr(vbKeyEscape) & .iConsideraQuantCotAnt & Chr(vbKeyEscape) & _
            .iTemFaixaReceb & Chr(vbKeyEscape) & .dPercentMaisReceb & Chr(vbKeyEscape) & .dPercentMenosReceb & Chr(vbKeyEscape) & .iRecebForaFaixa & Chr(vbKeyEscape) & .iTempoProducao & Chr(vbKeyEscape) & .iRastro & Chr(vbKeyEscape) & .lHorasMaquina & Chr(vbKeyEscape) & .dPesoEspecifico & Chr(vbKeyEscape) & .iEmbalagem & Chr(vbKeyEscape) & _
            .sReferencia & Chr(vbKeyEscape) & .sFigura & Chr(vbKeyEscape) & .dINSSPercBase & Chr(vbKeyEscape) & .iKitVendaComp & Chr(vbKeyEscape) & .sGrade & Chr(vbKeyEscape) & .sSituacaoTribECF & Chr(vbKeyEscape) & .sICMSAliquota & Chr(vbKeyEnd)
        
        End With
        
        colRegistros.Add sRegistro
        
        For Each objProdutoCateg In objProduto.colCategoriaItem
        
            sRegistro = TIPO_REGISTRO_PRODUTOCATEGORIA & Chr(vbKeyEscape) & objProdutoCateg.sCategoria & Chr(vbKeyEscape) & objProdutoCateg.sItem & Chr(vbKeyEnd)
            colRegistros.Add sRegistro
                
        Next
            
    Next
    
    Monta_Registros_Produtos = SUCESSO
    
    Exit Function

Erro_Monta_Registros_Produtos:

    Monta_Registros_Produtos = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150833)
    
    End Select
    
    Exit Function
    
End Function

Function Monta_Registros_Clientes(colRegistros As Collection, colClientes As Collection, colcolFiliaisCli As Collection, colcolEnderecos As Collection) As Long

Dim objCliente As ClassCliente
Dim objFilialCli As ClassFilialCliente
Dim sRegistro As String
Dim iIndice As Integer
Dim iIndiceEnd As Integer
Dim objEndereco As ClassEndereco
Dim objEnderecoCobranca As ClassEndereco
Dim objEnderecoEntrega As ClassEndereco

On Error GoTo Erro_Monta_Registros_Clientes

    iIndice = 0
    iIndiceEnd = 0
    
    For Each objCliente In colClientes
        iIndiceEnd = iIndiceEnd + 1
        iIndice = iIndice + 1
        
        Set objEndereco = colcolEnderecos(iIndiceEnd)(1)
        Set objEnderecoCobranca = colcolEnderecos(iIndiceEnd)(2)
        Set objEnderecoEntrega = colcolEnderecos(iIndiceEnd)(3)
        
        With objCliente
            sRegistro = TIPO_REGISTRO_CLIENTE & Chr(vbKeyEscape) & .lCodigo & Chr(vbKeyEscape) & .sRazaoSocial & Chr(vbKeyEscape) & .sNomeReduzido & Chr(vbKeyEscape) & .iTipo & Chr(vbKeyEscape) & .sObservacao & Chr(vbKeyEscape) & .dLimiteCredito & Chr(vbKeyEscape) & .iCondicaoPagto & Chr(vbKeyEscape) & .dDesconto & Chr(vbKeyEscape) & .iCodMensagem & Chr(vbKeyEscape) & .iTabelaPreco & Chr(vbKeyEscape) & .iProxCodFilial & Chr(vbKeyEscape) & .iCodPadraoCobranca & Chr(vbKeyEscape) & .iAtivo & Chr(vbKeyEscape) & .iTransferido & Chr(vbKeyEscape) & _
            .sCgc & Chr(vbKeyEscape) & .sInscricaoEstadual & Chr(vbKeyEscape) & .sInscricaoMunicipal & Chr(vbKeyEscape) & .sInscricaoSuframa & Chr(vbKeyEscape) & .lEndereco & Chr(vbKeyEscape) & .lEnderecoEntrega & Chr(vbKeyEscape) & .lEnderecoCobranca & Chr(vbKeyEscape) & .iCodTransportadora & Chr(vbKeyEscape) & .sObservacao2 & Chr(vbKeyEscape) & .sContaContabil & Chr(vbKeyEscape) & .iVendedor & Chr(vbKeyEscape) & .dComissaoVendas & Chr(vbKeyEscape) & .iRegiao & Chr(vbKeyEscape) & .iFreqVisitas & Chr(vbKeyEscape) & .dtDataUltVisita & Chr(vbKeyEscape) & .iCodCobrador & Chr(vbKeyEscape) & .iTipoFrete & Chr(vbKeyEscape) & .sRG & Chr(vbKeyEscape) & .iCodTranspRedesp & Chr(vbKeyEscape) & .iDetPagFrete & Chr(vbKeyEscape) & .sGuia & _
            Chr(vbKeyEscape) & objEndereco.sEndereco & Chr(vbKeyEscape) & objEndereco.sBairro & Chr(vbKeyEscape) & objEndereco.sCidade & Chr(vbKeyEscape) & objEndereco.sSiglaEstado & Chr(vbKeyEscape) & objEndereco.iCodigoPais & Chr(vbKeyEscape) & objEndereco.sCEP & Chr(vbKeyEscape) & objEndereco.sTelefone1 & Chr(vbKeyEscape) & objEndereco.sTelefone2 & Chr(vbKeyEscape) & objEndereco.sEmail & Chr(vbKeyEscape) & objEndereco.sFax & Chr(vbKeyEscape) & objEndereco.sContato & _
            Chr(vbKeyEscape) & objEnderecoCobranca.sEndereco & Chr(vbKeyEscape) & objEnderecoCobranca.sBairro & Chr(vbKeyEscape) & objEnderecoCobranca.sCidade & Chr(vbKeyEscape) & objEnderecoCobranca.sSiglaEstado & Chr(vbKeyEscape) & objEnderecoCobranca.iCodigoPais & Chr(vbKeyEscape) & objEnderecoCobranca.sCEP & Chr(vbKeyEscape) & objEnderecoCobranca.sTelefone1 & Chr(vbKeyEscape) & objEnderecoCobranca.sTelefone2 & Chr(vbKeyEscape) & objEnderecoCobranca.sEmail & Chr(vbKeyEscape) & objEnderecoCobranca.sFax & Chr(vbKeyEscape) & objEnderecoCobranca.sContato & _
            Chr(vbKeyEscape) & objEnderecoEntrega.sEndereco & Chr(vbKeyEscape) & objEnderecoEntrega.sBairro & Chr(vbKeyEscape) & objEnderecoEntrega.sCidade & Chr(vbKeyEscape) & objEnderecoEntrega.sSiglaEstado & Chr(vbKeyEscape) & objEnderecoEntrega.iCodigoPais & Chr(vbKeyEscape) & objEnderecoEntrega.sCEP & Chr(vbKeyEscape) & objEnderecoEntrega.sTelefone1 & Chr(vbKeyEscape) & objEnderecoEntrega.sTelefone2 & Chr(vbKeyEscape) & objEnderecoEntrega.sEmail & Chr(vbKeyEscape) & objEnderecoEntrega.sFax & Chr(vbKeyEscape) & objEnderecoEntrega.sContato & Chr(vbKeyEnd)
        End With
        colRegistros.Add sRegistro
        
        For Each objFilialCli In colcolFiliaisCli(iIndice)
            With objFilialCli
                sRegistro = TIPO_REGISTRO_FILIALCLIENTE & Chr(vbKeyEscape) & .lCodCliente & Chr(vbKeyEscape) & .iCodFilial & Chr(vbKeyEscape) & .sNome & Chr(vbKeyEscape) & .sCgc & Chr(vbKeyEscape) & .sInscricaoEstadual & Chr(vbKeyEscape) & .sInscricaoMunicipal & Chr(vbKeyEscape) & .sInscricaoSuframa & Chr(vbKeyEscape) & .lEndereco & Chr(vbKeyEscape) & .lEnderecoEntrega & Chr(vbKeyEscape) & _
                .lEnderecoCobranca & Chr(vbKeyEscape) & .iCodTransportadora & Chr(vbKeyEscape) & .sObservacao & Chr(vbKeyEscape) & .sContaContabil & Chr(vbKeyEscape) & .iVendedor & Chr(vbKeyEscape) & .dComissaoVendas & Chr(vbKeyEscape) & .iRegiao & Chr(vbKeyEscape) & .iFreqVisitas & Chr(vbKeyEscape) & .dtDataUltVisita & Chr(vbKeyEscape) & .iCodCobrador & Chr(vbKeyEscape) & .iTipoFrete & Chr(vbKeyEscape) & .sRG & Chr(vbKeyEscape) & _
                .iAtivo & Chr(vbKeyEscape) & .iCodTranspRedesp & Chr(vbKeyEscape) & .iDetPagFrete & Chr(vbKeyEscape) & .sGuia & _
                Chr(vbKeyEscape) & objEndereco.sEndereco & Chr(vbKeyEscape) & objEndereco.sBairro & Chr(vbKeyEscape) & objEndereco.sCidade & Chr(vbKeyEscape) & objEndereco.sSiglaEstado & Chr(vbKeyEscape) & objEndereco.iCodigoPais & Chr(vbKeyEscape) & objEndereco.sCEP & Chr(vbKeyEscape) & objEndereco.sTelefone1 & Chr(vbKeyEscape) & objEndereco.sTelefone2 & Chr(vbKeyEscape) & objEndereco.sEmail & Chr(vbKeyEscape) & objEndereco.sFax & Chr(vbKeyEscape) & objEndereco.sContato & _
                Chr(vbKeyEscape) & objEnderecoCobranca.sEndereco & Chr(vbKeyEscape) & objEnderecoCobranca.sBairro & Chr(vbKeyEscape) & objEnderecoCobranca.sCidade & Chr(vbKeyEscape) & objEnderecoCobranca.sSiglaEstado & Chr(vbKeyEscape) & objEnderecoCobranca.iCodigoPais & Chr(vbKeyEscape) & objEnderecoCobranca.sCEP & Chr(vbKeyEscape) & objEnderecoCobranca.sTelefone1 & Chr(vbKeyEscape) & objEnderecoCobranca.sTelefone2 & Chr(vbKeyEscape) & objEnderecoCobranca.sEmail & Chr(vbKeyEscape) & objEnderecoCobranca.sFax & Chr(vbKeyEscape) & objEnderecoCobranca.sContato & _
                Chr(vbKeyEscape) & objEnderecoEntrega.sEndereco & Chr(vbKeyEscape) & objEnderecoEntrega.sBairro & Chr(vbKeyEscape) & objEnderecoEntrega.sCidade & Chr(vbKeyEscape) & objEnderecoEntrega.sSiglaEstado & Chr(vbKeyEscape) & objEnderecoEntrega.iCodigoPais & Chr(vbKeyEscape) & objEnderecoEntrega.sCEP & Chr(vbKeyEscape) & objEnderecoEntrega.sTelefone1 & Chr(vbKeyEscape) & objEnderecoEntrega.sTelefone2 & Chr(vbKeyEscape) & objEnderecoEntrega.sEmail & Chr(vbKeyEscape) & objEnderecoEntrega.sFax & Chr(vbKeyEscape) & objEnderecoEntrega.sContato & Chr(vbKeyEnd)
            End With
            colRegistros.Add sRegistro
        Next
        
    Next
    
    Monta_Registros_Clientes = SUCESSO

    Exit Function

Erro_Monta_Registros_Clientes:

    Monta_Registros_Clientes = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150834)
    
    End Select
    
    Exit Function
    
End Function

Function Monta_Registros_Fornecedores(colRegistros As Collection, colFornecedores As Collection, colcolFiliaisForn As Collection, colEnderecos As Collection) As Long

Dim objFornecedor As ClassFornecedor
Dim objFilialForn As ClassFilialFornecedor
Dim sRegistro As String
Dim iIndice As Integer
Dim objEndereco As ClassEndereco
Dim iIndiceEnd As Integer

On Error GoTo Erro_Monta_Registros_Fornecedores

    iIndice = 0
    iIndiceEnd = 0
    
    For Each objFornecedor In colFornecedores
        
        iIndice = iIndice + 1
        iIndiceEnd = iIndiceEnd + 1
        
        Set objEndereco = colEnderecos(iIndice)
        
        With objFornecedor
            sRegistro = TIPO_REGISTRO_FORNECEDOR & Chr(vbKeyEscape) & .sRazaoSocial & Chr(vbKeyEscape) & .sNomeReduzido & Chr(vbKeyEscape) & .iTipo & Chr(vbKeyEscape) & .sObservacao & Chr(vbKeyEscape) & .iCondicaoPagto & Chr(vbKeyEscape) & .dDesconto & Chr(vbKeyEscape) & .iProxCodFilial & Chr(vbKeyEscape) & .lCodigo & Chr(vbKeyEscape) & .lEndereco & Chr(vbKeyEscape) & .sCgc & Chr(vbKeyEscape) & .sInscricaoEstadual & Chr(vbKeyEscape) & .sInscricaoMunicipal & Chr(vbKeyEscape) & .sContaContabil & Chr(vbKeyEscape) & .sContaFornConsig & Chr(vbKeyEscape) & .iBanco & Chr(vbKeyEscape) & .sAgencia & Chr(vbKeyEscape) & .sContaCorrente & Chr(vbKeyEscape) & .sObservacao2 & Chr(vbKeyEscape) & .iTipoFrete & Chr(vbKeyEscape) & .sInscricaoINSS & Chr(vbKeyEscape) & .iTipoCobranca & Chr(vbKeyEscape) & .iGeraCredICMS & Chr(vbKeyEscape) & _
            .sContaDespesa & Chr(vbKeyEscape) & objEndereco.sEndereco & Chr(vbKeyEscape) & objEndereco.sBairro & Chr(vbKeyEscape) & objEndereco.sCidade & Chr(vbKeyEscape) & objEndereco.sSiglaEstado & Chr(vbKeyEscape) & objEndereco.iCodigoPais & Chr(vbKeyEscape) & objEndereco.sCEP & Chr(vbKeyEscape) & objEndereco.sTelefone1 & Chr(vbKeyEscape) & objEndereco.sTelefone2 & Chr(vbKeyEscape) & objEndereco.sEmail & Chr(vbKeyEscape) & objEndereco.sFax & Chr(vbKeyEscape) & objEndereco.sContato & Chr(vbKeyEnd)
            colRegistros.Add sRegistro
        End With
        
        For Each objFilialForn In colcolFiliaisForn(iIndice)
            Set objEndereco = colEnderecos(iIndice)
            With objFilialForn
                sRegistro = TIPO_REGISTRO_FILIALFORNECEDOR & Chr(vbKeyEscape) & .sNome & Chr(vbKeyEscape) & .lEndereco & Chr(vbKeyEscape) & .sCgc & Chr(vbKeyEscape) & .sInscricaoEstadual & Chr(vbKeyEscape) & .sInscricaoMunicipal & Chr(vbKeyEscape) & .sContaContabil & Chr(vbKeyEscape) & .sContaFornConsig & Chr(vbKeyEscape) & .iBanco & Chr(vbKeyEscape) & .sAgencia & Chr(vbKeyEscape) & .sContaCorrente & Chr(vbKeyEscape) & .sObservacao & Chr(vbKeyEscape) & .iTipoFrete & Chr(vbKeyEscape) & .sInscricaoINSS & Chr(vbKeyEscape) & .iTipoCobranca & Chr(vbKeyEscape) & .iGeraCredICMS & Chr(vbKeyEscape) & .sContaDespesa & Chr(vbKeyEscape) & .lCodFornecedor & Chr(vbKeyEscape) & .iCodFilial _
                & Chr(vbKeyEscape) & objEndereco.sEndereco & Chr(vbKeyEscape) & objEndereco.sBairro & Chr(vbKeyEscape) & objEndereco.sCidade & Chr(vbKeyEscape) & objEndereco.sSiglaEstado & Chr(vbKeyEscape) & objEndereco.iCodigoPais & Chr(vbKeyEscape) & objEndereco.sCEP & Chr(vbKeyEscape) & objEndereco.sTelefone1 & Chr(vbKeyEscape) & objEndereco.sTelefone2 & Chr(vbKeyEscape) & objEndereco.sEmail & Chr(vbKeyEscape) & objEndereco.sFax & Chr(vbKeyEscape) & objEndereco.sContato & Chr(vbKeyEnd)
            End With
            colRegistros.Add sRegistro
        Next
        
    Next
    
    Monta_Registros_Fornecedores = SUCESSO
    
    Exit Function

Erro_Monta_Registros_Fornecedores:

    Monta_Registros_Fornecedores = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150835)
    
    End Select
    
    Exit Function
    
End Function

Function Monta_Registros_NotasFiscais(colRegistros As Collection, colNotasFiscais As Collection, colcolLancamentos As Collection) As Long

Dim objNFiscal As ClassNFiscal
Dim objItemNF As ClassItemNF
Dim sRegistro As String
Dim objItemNFGrade As ClassItemRomaneioGrade
Dim objParcelaPag As ClassParcelaPagar
Dim objParcelaRec As ClassParcelaReceber
Dim objAlocacaoNF As ClassItemNFAlocacao
Dim objLocalizacaoItemGrade As ClassReservaItem
Dim colLancamentos As Collection
Dim colLanPendente As Collection
Dim iIndice As Integer
Dim objLancto As ClassLancamento_Detalhe

On Error GoTo Erro_Monta_Registros_NotasFiscais

    For Each objNFiscal In colNotasFiscais
        
        iIndice = iIndice + 1
        
        
        With objNFiscal
            sRegistro = TIPO_REGISTRO_NOTAFISCAL & Chr(vbKeyEscape) & .lNumIntDoc & Chr(vbKeyEscape) & .iStatus & Chr(vbKeyEscape) & .iFilialEmpresa & Chr(vbKeyEscape) & .sSerie & Chr(vbKeyEscape) & .lNumNotaFiscal & Chr(vbKeyEscape) & .lCliente & Chr(vbKeyEscape) & .iFilialCli & Chr(vbKeyEscape) & .iFilialEntrega & Chr(vbKeyEscape) & .lFornecedor & Chr(vbKeyEscape) & .iFilialForn & Chr(vbKeyEscape) & _
            .dtDataEmissao & Chr(vbKeyEscape) & .dtDataEntrada & Chr(vbKeyEscape) & .dtDataSaida & Chr(vbKeyEscape) & .dtDataVencimento & Chr(vbKeyEscape) & .dtDataReferencia & Chr(vbKeyEscape) & .iFilialPedido & Chr(vbKeyEscape) & .sNumPedidoTerc & Chr(vbKeyEscape) & .iClasseDocCPR & Chr(vbKeyEscape) & .lNumIntDocCPR & Chr(vbKeyEscape) & .dValorTotal & Chr(vbKeyEscape) & .dValorProdutos & Chr(vbKeyEscape) & _
            .dValorFrete & Chr(vbKeyEscape) & .dValorSeguro & Chr(vbKeyEscape) & .dValorOutrasDespesas & Chr(vbKeyEscape) & .dValorDesconto & Chr(vbKeyEscape) & .iCodTransportadora & Chr(vbKeyEscape) & .sMensagemNota & Chr(vbKeyEscape) & .iTabelaPreco & Chr(vbKeyEscape) & .iTipoNFiscal & Chr(vbKeyEscape) & .sNaturezaOp & Chr(vbKeyEscape) & .dPesoLiq & Chr(vbKeyEscape) & .dPesoBruto & Chr(vbKeyEscape) & _
            .lNumIntTrib & Chr(vbKeyEscape) & .sPlaca & Chr(vbKeyEscape) & .sPlacaUF & Chr(vbKeyEscape) & .lVolumeQuant & Chr(vbKeyEscape) & .lVolumeEspecie & Chr(vbKeyEscape) & .lVolumeMarca & Chr(vbKeyEscape) & .sVolumeNumero & Chr(vbKeyEscape) & .iCanal & Chr(vbKeyEscape) & .lNumIntNotaOriginal & Chr(vbKeyEscape) & .lClienteBenef & Chr(vbKeyEscape) & .iFilialCliBenef & Chr(vbKeyEscape) & .lFornecedorBenef & Chr(vbKeyEscape) & _
            .iFilialFornBenef & Chr(vbKeyEscape) & .iFreteRespons & Chr(vbKeyEscape) & .lNumRecebimento & Chr(vbKeyEscape) & .sObservacao & Chr(vbKeyEscape) & .sCodUsuarioCancel & Chr(vbKeyEscape) & .sMotivoCancel & Chr(vbKeyEscape) & .dtHoraEntrada & Chr(vbKeyEscape) & .dtHoraSaida & Chr(vbKeyEscape) & .iCodTranspRedesp & Chr(vbKeyEscape) & .iDetPagFrete & Chr(vbKeyEscape) & .iSemDataSaida & .iNaoImpCobranca
        End With
               
        With objNFiscal.objTributacaoNF
            sRegistro = sRegistro & Chr(vbKeyEscape) & .sNaturezaOpInterna & Chr(vbKeyEscape) & .iTipoTributacao & Chr(vbKeyEscape) & .dIPIBase & Chr(vbKeyEscape) & .dIPIValor & Chr(vbKeyEscape) & .dIPICredito & Chr(vbKeyEscape) & .dICMSBase & Chr(vbKeyEscape) & .dICMSValor & Chr(vbKeyEscape) & .dICMSSubstBase & Chr(vbKeyEscape) & .dICMSSubstValor & Chr(vbKeyEscape) & .dICMSCredito & Chr(vbKeyEscape) & .iISSIncluso & Chr(vbKeyEscape) & .dISSBase & Chr(vbKeyEscape) & .dISSAliquota & Chr(vbKeyEscape) & .dISSValor & Chr(vbKeyEscape) & .dIRRFBase & Chr(vbKeyEscape) & .dIRRFAliquota & Chr(vbKeyEscape) & .dIRRFValor & Chr(vbKeyEscape) & .dValorINSS & Chr(vbKeyEscape) & .iINSSRetido & Chr(vbKeyEscape) & .dINSSBase & Chr(vbKeyEscape) & .dINSSDeducoes & Chr(vbKeyEscape) & .dPISCredito & Chr(vbKeyEscape) & .dCOFINSCredito & Chr(vbKeyEscape) & .dPISRetido & Chr(vbKeyEscape) & .dCOFINSRetido & Chr(vbKeyEscape) & .dCSLLRetido
        End With
        'XXXX
        'VER COMO TRATAR
'        With objNFiscal.objTributacaoNF.objTributacaoDesconto
'             sRegistro = sRegistro & Chr(vbKeyEscape) & .dCOFINSCredito & Chr(vbKeyEscape) & .dICMSAliquota & Chr(vbKeyEscape) & .dICMSBase & Chr(vbKeyEscape) & .dICMSCredito & Chr(vbKeyEscape) & .dICMSPercRedBase & Chr(vbKeyEscape) & .dICMSSubstAliquota & Chr(vbKeyEscape) & .dICMSSubstBase & Chr(vbKeyEscape) & .dICMSSubstValor & Chr(vbKeyEscape) & .dICMSValor & Chr(vbKeyEscape) & .dIPIAliquota & Chr(vbKeyEscape) & .dIPIBaseCalculo & Chr(vbKeyEscape) & .dIPICredito & Chr(vbKeyEscape) & .dIPIPercRedBase & Chr(vbKeyEscape) & .dIPIValor & Chr(vbKeyEscape) & .dPISCredito & Chr(vbKeyEscape) & .iICMSAliquotaManual & Chr(vbKeyEscape) & .iICMSBaseManual & Chr(vbKeyEscape) & .iICMSPercRedBaseManual & Chr(vbKeyEscape) & _
'             .iICMSSubstAliquotaManual & Chr(vbKeyEscape) & .iICMSSubstBaseManual & Chr(vbKeyEscape) & .iICMSTipo & Chr(vbKeyEscape) & .iICMSTipoManual & Chr(vbKeyEscape) & .iICMSValorManual & Chr(vbKeyEscape) & .iIPIAliquotaManual & Chr(vbKeyEscape) & .iIPIBaseManual & Chr(vbKeyEscape) & .iIPIPercRedBaseManual & Chr(vbKeyEscape) & .iIPITipo & Chr(vbKeyEscape) & .iIPITipoManual & Chr(vbKeyEscape) & .iIPIValorManual & Chr(vbKeyEscape) & .iNaturezaOpManual & Chr(vbKeyEscape) & .iTipo & Chr(vbKeyEscape) & .iTipoTributacao & Chr(vbKeyEscape) & .iTipoTributacaoManual & Chr(vbKeyEscape) & .lNumIntDoc & Chr(vbKeyEscape) & .sNaturezaOp
'        End With
'        With objNFiscal.objTributacaoNF.objTributacaoFrete
'             sRegistro = sRegistro & Chr(vbKeyEscape) & .dCOFINSCredito & Chr(vbKeyEscape) & .dICMSAliquota & Chr(vbKeyEscape) & .dICMSBase & Chr(vbKeyEscape) & .dICMSCredito & Chr(vbKeyEscape) & .dICMSPercRedBase & Chr(vbKeyEscape) & .dICMSSubstAliquota & Chr(vbKeyEscape) & .dICMSSubstBase & Chr(vbKeyEscape) & .dICMSSubstValor & Chr(vbKeyEscape) & .dICMSValor & Chr(vbKeyEscape) & .dIPIAliquota & Chr(vbKeyEscape) & .dIPIBaseCalculo & Chr(vbKeyEscape) & .dIPICredito & Chr(vbKeyEscape) & .dIPIPercRedBase & Chr(vbKeyEscape) & .dIPIValor & Chr(vbKeyEscape) & .dPISCredito & Chr(vbKeyEscape) & .iICMSAliquotaManual & Chr(vbKeyEscape) & .iICMSBaseManual & Chr(vbKeyEscape) & .iICMSPercRedBaseManual & Chr(vbKeyEscape) & _
'             .iICMSSubstAliquotaManual & Chr(vbKeyEscape) & .iICMSSubstBaseManual & Chr(vbKeyEscape) & .iICMSTipo & Chr(vbKeyEscape) & .iICMSTipoManual & Chr(vbKeyEscape) & .iICMSValorManual & Chr(vbKeyEscape) & .iIPIAliquotaManual & Chr(vbKeyEscape) & .iIPIBaseManual & Chr(vbKeyEscape) & .iIPIPercRedBaseManual & Chr(vbKeyEscape) & .iIPITipo & Chr(vbKeyEscape) & .iIPITipoManual & Chr(vbKeyEscape) & .iIPIValorManual & Chr(vbKeyEscape) & .iNaturezaOpManual & Chr(vbKeyEscape) & .iTipo & Chr(vbKeyEscape) & .iTipoTributacao & Chr(vbKeyEscape) & .iTipoTributacaoManual & Chr(vbKeyEscape) & .lNumIntDoc & Chr(vbKeyEscape) & .sNaturezaOp
'        End With
'        With objNFiscal.objTributacaoNF.objTributacaoOutras
'             sRegistro = sRegistro & Chr(vbKeyEscape) & .dCOFINSCredito & Chr(vbKeyEscape) & .dICMSAliquota & Chr(vbKeyEscape) & .dICMSBase & Chr(vbKeyEscape) & .dICMSCredito & Chr(vbKeyEscape) & .dICMSPercRedBase & Chr(vbKeyEscape) & .dICMSSubstAliquota & Chr(vbKeyEscape) & .dICMSSubstBase & Chr(vbKeyEscape) & .dICMSSubstValor & Chr(vbKeyEscape) & .dICMSValor & Chr(vbKeyEscape) & .dIPIAliquota & Chr(vbKeyEscape) & .dIPIBaseCalculo & Chr(vbKeyEscape) & .dIPICredito & Chr(vbKeyEscape) & .dIPIPercRedBase & Chr(vbKeyEscape) & .dIPIValor & Chr(vbKeyEscape) & .dPISCredito & Chr(vbKeyEscape) & .iICMSAliquotaManual & Chr(vbKeyEscape) & .iICMSBaseManual & Chr(vbKeyEscape) & .iICMSPercRedBaseManual & Chr(vbKeyEscape) & _
'             .iICMSSubstAliquotaManual & Chr(vbKeyEscape) & .iICMSSubstBaseManual & Chr(vbKeyEscape) & .iICMSTipo & Chr(vbKeyEscape) & .iICMSTipoManual & Chr(vbKeyEscape) & .iICMSValorManual & Chr(vbKeyEscape) & .iIPIAliquotaManual & Chr(vbKeyEscape) & .iIPIBaseManual & Chr(vbKeyEscape) & .iIPIPercRedBaseManual & Chr(vbKeyEscape) & .iIPITipo & Chr(vbKeyEscape) & .iIPITipoManual & Chr(vbKeyEscape) & .iIPIValorManual & Chr(vbKeyEscape) & .iNaturezaOpManual & Chr(vbKeyEscape) & .iTipo & Chr(vbKeyEscape) & .iTipoTributacao & Chr(vbKeyEscape) & .iTipoTributacaoManual & Chr(vbKeyEscape) & .lNumIntDoc & Chr(vbKeyEscape) & .sNaturezaOp
'        End With
'        With objNFiscal.objTributacaoNF.objTributacaoSeguro
'             sRegistro = sRegistro & Chr(vbKeyEscape) & .dCOFINSCredito & Chr(vbKeyEscape) & .dICMSAliquota & Chr(vbKeyEscape) & .dICMSBase & Chr(vbKeyEscape) & .dICMSCredito & Chr(vbKeyEscape) & .dICMSPercRedBase & Chr(vbKeyEscape) & .dICMSSubstAliquota & Chr(vbKeyEscape) & .dICMSSubstBase & Chr(vbKeyEscape) & .dICMSSubstValor & Chr(vbKeyEscape) & .dICMSValor & Chr(vbKeyEscape) & .dIPIAliquota & Chr(vbKeyEscape) & .dIPIBaseCalculo & Chr(vbKeyEscape) & .dIPICredito & Chr(vbKeyEscape) & .dIPIPercRedBase & Chr(vbKeyEscape) & .dIPIValor & Chr(vbKeyEscape) & .dPISCredito & Chr(vbKeyEscape) & .iICMSAliquotaManual & Chr(vbKeyEscape) & .iICMSBaseManual & Chr(vbKeyEscape) & .iICMSPercRedBaseManual & Chr(vbKeyEscape) & _
'             .iICMSSubstAliquotaManual & Chr(vbKeyEscape) & .iICMSSubstBaseManual & Chr(vbKeyEscape) & .iICMSTipo & Chr(vbKeyEscape) & .iICMSTipoManual & Chr(vbKeyEscape) & .iICMSValorManual & Chr(vbKeyEscape) & .iIPIAliquotaManual & Chr(vbKeyEscape) & .iIPIBaseManual & Chr(vbKeyEscape) & .iIPIPercRedBaseManual & Chr(vbKeyEscape) & .iIPITipo & Chr(vbKeyEscape) & .iIPITipoManual & Chr(vbKeyEscape) & .iIPIValorManual & Chr(vbKeyEscape) & .iNaturezaOpManual & Chr(vbKeyEscape) & .iTipo & Chr(vbKeyEscape) & .iTipoTributacao & Chr(vbKeyEscape) & .iTipoTributacaoManual & Chr(vbKeyEscape) & .lNumIntDoc & Chr(vbKeyEscape) & .sNaturezaOp
'        End With
    
        sRegistro = sRegistro & Chr(vbKeyEnd)
    
        colRegistros.Add sRegistro
        
        For Each objParcelaPag In objNFiscal.ColParcelaPagar
            With objParcelaPag
                sRegistro = TIPO_REGISTRO_PARCELAPAG & Chr(vbKeyEscape) & .lNumIntDoc & Chr(vbKeyEscape) & .iNumParcela & Chr(vbKeyEscape) & .iStatus & Chr(vbKeyEscape) & .dtDataVencimento & Chr(vbKeyEscape) & .dtDataVencimentoReal & Chr(vbKeyEscape) & .dSaldo & Chr(vbKeyEscape) & .dValor & Chr(vbKeyEscape) & .iPortador & Chr(vbKeyEscape) & .iProxSeqBaixa & Chr(vbKeyEscape) & .iTipoCobranca & Chr(vbKeyEscape) & .iBancoCobrador & Chr(vbKeyEscape) & .sNossoNumero & Chr(vbKeyEscape) & .sCodigoDeBarras & Chr(vbKeyEnd)
            End With
            
            colRegistros.Add sRegistro
            
        Next
        
        For Each objParcelaRec In objNFiscal.ColParcelaReceber
            With objParcelaRec
                sRegistro = TIPO_REGISTRO_PARCELAREC & Chr(vbKeyEscape) & .lNumIntDoc & Chr(vbKeyEscape) & .iStatus & Chr(vbKeyEscape) & .dtDataVencimento & Chr(vbKeyEscape) & .dValor & Chr(vbKeyEscape) & .iCarteiraCobranca & Chr(vbKeyEscape) & .iCobrador & Chr(vbKeyEscape) & .iDesconto1Codigo & Chr(vbKeyEscape) & .dtDesconto1Ate & Chr(vbKeyEscape) & .dDesconto1Valor & Chr(vbKeyEscape) & .iDesconto2Codigo & Chr(vbKeyEscape) & .dtDesconto2Ate & Chr(vbKeyEscape) & .dDesconto2Valor & Chr(vbKeyEscape) & .iDesconto3Codigo & Chr(vbKeyEscape) & .dtDesconto3Ate & Chr(vbKeyEscape) & .dDesconto3Valor & Chr(vbKeyEnd)
            End With
            colRegistros.Add sRegistro
        Next
               
        Set colLancamentos = colcolLancamentos(2 * iIndice - 1)
        Set colLanPendente = colcolLancamentos(2 * iIndice)
            
        For Each objLancto In colLancamentos
            sRegistro = TIPO_REGISTRO_LANCAMENTO & Chr(vbKeyEscape)
           sRegistro = sRegistro & objLancto.iFilialEmpresa & Chr(vbKeyEscape) & objLancto.sOrigem & Chr(vbKeyEscape) & objLancto.iExercicio & Chr(vbKeyEscape) & objLancto.iPeriodoLan & Chr(vbKeyEscape) & objLancto.lDoc & Chr(vbKeyEscape) & objLancto.iSeq & Chr(vbKeyEscape) & objLancto.iLote & Chr(vbKeyEscape) & objLancto.iPeriodoLote & Chr(vbKeyEscape) & objLancto.dtData & Chr(vbKeyEscape) & objLancto.sConta & Chr(vbKeyEscape) & objLancto.sCcl & Chr(vbKeyEscape) & objLancto.sHistorico & Chr(vbKeyEscape) & objLancto.dValor & Chr(vbKeyEscape) & objLancto.lNumIntDoc & Chr(vbKeyEscape) & objLancto.iFilialCliForn & Chr(vbKeyEscape) & objLancto.lCliForn & Chr(vbKeyEscape) & objLancto.iTransacao & Chr(vbKeyEscape) & objLancto.lDoc & Chr(vbKeyEscape) & objLancto.iAglutina & Chr(vbKeyEscape) & objLancto.lContaSimples & Chr(vbKeyEscape) & objLancto.iSeqContraPartida & Chr(vbKeyEscape) & objLancto.iEscaninho_Custo & Chr(vbKeyEscape) & objLancto.dValorLivroAuxiliar & Chr(vbKeyEscape) & _
            objLancto.iClienteFornecedor & Chr(vbKeyEscape) & objLancto.sDocOrigem & Chr(vbKeyEscape) & objLancto.sProduto & Chr(vbKeyEscape) & objLancto.dQuantidade & Chr(vbKeyEscape) & objLancto.dtDataEstoque & Chr(vbKeyEscape) & objLancto.iApropriaCRProd & Chr(vbKeyEscape) & objLancto.iStatus & Chr(vbKeyEscape) & objLancto.sModelo & Chr(vbKeyEnd)
            colRegistros.Add sRegistro
        Next
               
        For Each objLancto In colLanPendente
            sRegistro = TIPO_REGISTRO_LANPENDENTE & Chr(vbKeyEscape)
            sRegistro = sRegistro & objLancto.iFilialEmpresa & Chr(vbKeyEscape) & objLancto.sOrigem & Chr(vbKeyEscape) & objLancto.iExercicio & Chr(vbKeyEscape) & objLancto.iPeriodoLan & Chr(vbKeyEscape) & objLancto.lDoc & Chr(vbKeyEscape) & objLancto.iSeq & Chr(vbKeyEscape) & objLancto.iLote & Chr(vbKeyEscape) & objLancto.iPeriodoLote & Chr(vbKeyEscape) & objLancto.dtData & Chr(vbKeyEscape) & objLancto.sConta & Chr(vbKeyEscape) & objLancto.sCcl & Chr(vbKeyEscape) & objLancto.sHistorico & Chr(vbKeyEscape) & objLancto.dValor & Chr(vbKeyEscape) & objLancto.lNumIntDoc & Chr(vbKeyEscape) & objLancto.iFilialCliForn & Chr(vbKeyEscape) & objLancto.lCliForn & Chr(vbKeyEscape) & objLancto.iTransacao & Chr(vbKeyEscape) & objLancto.lDoc & Chr(vbKeyEscape) & objLancto.iAglutina & Chr(vbKeyEscape) & objLancto.lContaSimples & Chr(vbKeyEscape) & objLancto.iSeqContraPartida & Chr(vbKeyEscape) & objLancto.iEscaninho_Custo & Chr(vbKeyEscape) & objLancto.dValorLivroAuxiliar & Chr(vbKeyEscape) & _
            objLancto.iClienteFornecedor & Chr(vbKeyEscape) & objLancto.sDocOrigem & Chr(vbKeyEscape) & objLancto.sProduto & Chr(vbKeyEscape) & objLancto.dQuantidade & Chr(vbKeyEscape) & objLancto.dtDataEstoque & Chr(vbKeyEscape) & objLancto.iApropriaCRProd & Chr(vbKeyEscape) & objLancto.iStatus & Chr(vbKeyEscape) & objLancto.sModelo & Chr(vbKeyEnd)
            colRegistros.Add sRegistro
        Next
        
        For Each objItemNF In objNFiscal.colItensNF
        
            With objItemNF
                sRegistro = TIPO_REGISTRO_ITEMNOTAFISCAL & Chr(vbKeyEscape) & .iItem & Chr(vbKeyEscape) & .sProduto & Chr(vbKeyEscape) & .sUnidadeMed & Chr(vbKeyEscape) & .dQuantidade & Chr(vbKeyEscape) & .dPrecoUnitario & Chr(vbKeyEscape) & .dPercDesc & Chr(vbKeyEscape) & .dValorDesconto & Chr(vbKeyEscape) & .dtDataEntrega & Chr(vbKeyEscape) & .sDescricaoItem & Chr(vbKeyEscape) & .dValorAbatComissao & Chr(vbKeyEscape) & .lNumIntPedVenda & Chr(vbKeyEscape) & .lNumIntItemPedVenda & Chr(vbKeyEscape) & .lNumIntDoc & Chr(vbKeyEscape) & .lNumIntTrib & Chr(vbKeyEscape) & .lNumIntDocOrig & Chr(vbKeyEscape) & .sCcl & Chr(vbKeyEscape) & .iAlmoxarifado
            End With
            
            With objItemNF.objTributacaoItemNF
                sRegistro = sRegistro & Chr(vbKeyEscape) & .iItem & Chr(vbKeyEscape) & .lNumIntDoc & Chr(vbKeyEscape) & .sNaturezaOp & Chr(vbKeyEscape) & .iTipoTributacao & Chr(vbKeyEscape) & .iIPITipo & Chr(vbKeyEscape) & .sIPICodProduto & Chr(vbKeyEscape) & .dIPIBaseCalculo & Chr(vbKeyEscape) & .dIPIPercRedBase & Chr(vbKeyEscape) & .dIPIAliquota & Chr(vbKeyEscape) & .dIPIValor & Chr(vbKeyEscape) & .dIPICredito & Chr(vbKeyEscape) & .iICMSTipo & Chr(vbKeyEscape) & .dICMSBase & Chr(vbKeyEscape) & .dICMSPercRedBase & Chr(vbKeyEscape) & .dICMSAliquota & Chr(vbKeyEscape) & .dICMSValor & Chr(vbKeyEscape) & .dICMSCredito & Chr(vbKeyEscape) & .dICMSSubstBase & Chr(vbKeyEscape) & .dICMSSubstAliquota & Chr(vbKeyEscape) & .dICMSSubstValor & Chr(vbKeyEscape) & .dPISCredito & Chr(vbKeyEscape) & .dCOFINSCredito & Chr(vbKeyEnd)
            End With
            colRegistros.Add sRegistro
            
            For Each objAlocacaoNF In objItemNF.colAlocacoes
                sRegistro = TIPO_REGISTRO_ALOCACAOITEMNF & Chr(vbKeyEscape) & objAlocacaoNF.dQuantidade & Chr(vbKeyEscape) & objAlocacaoNF.iAlmoxarifado & Chr(vbKeyEscape) & objAlocacaoNF.iTransferencia & Chr(vbKeyEscape) & objAlocacaoNF.sAlmoxarifado & Chr(vbKeyEscape) & objAlocacaoNF.sUnidadeMed & Chr(vbKeyEnd)
                colRegistros.Add sRegistro
            Next
            
            For Each objItemNFGrade In objItemNF.colItensRomaneioGrade
                With objItemNFGrade
                    sRegistro = TIPO_REGISTRO_ITEMNFGRADE & Chr(vbKeyEscape) & .iControleEstoque & Chr(vbKeyEscape) & .lNumIntDoc & Chr(vbKeyEscape) & .sProduto & Chr(vbKeyEscape) & .dQuantidade & Chr(vbKeyEscape) & .sDescricao & Chr(vbKeyEscape) & .sUMEstoque & Chr(vbKeyEscape) & .lNumIntItemPV & Chr(vbKeyEscape) & .iAlmoxarifado & Chr(vbKeyEnd)
                End With
                colRegistros.Add sRegistro
                
                For Each objLocalizacaoItemGrade In objItemNFGrade.colLocalizacao
                    sRegistro = TIPO_REGISTRO_LOCALIZACAOITEMNFGRADE & Chr(vbKeyEscape) & objLocalizacaoItemGrade.dQuantidade & Chr(vbKeyEscape) & objLocalizacaoItemGrade.dtDataValidade & Chr(vbKeyEscape) & objLocalizacaoItemGrade.iAlmoxarifado & Chr(vbKeyEscape) & objLocalizacaoItemGrade.iFilialEmpresa & Chr(vbKeyEscape) & objLocalizacaoItemGrade.sAlmoxarifado & Chr(vbKeyEscape) & objLocalizacaoItemGrade.sResponsavel & Chr(vbKeyEnd)
                    colRegistros.Add sRegistro
                Next
                            
            Next
        
        Next
        
    Next

    Monta_Registros_NotasFiscais = SUCESSO

    Exit Function

Erro_Monta_Registros_NotasFiscais:

    Monta_Registros_NotasFiscais = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150836)
    
    End Select
    
    Exit Function
    
End Function

Function NFiscais_Importar(ByVal sNomeArquivo As String) As Long

Dim lErro As Long
Dim iTipoRegistro As Integer
Dim sRegistro As String
Dim objProduto As ClassProduto
Dim objCliente As ClassCliente
Dim colFiliais As Collection
Dim objFornecedor As ClassFornecedor
Dim colcolEnderecos As Collection
Dim colEnderecos As Collection
Dim objNFiscal As ClassNFiscal
Dim colLancamentos As Collection
Dim colLanPendente As Collection
Dim lTransacao As Long

On Error GoTo Erro_NFiscais_Importar

    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 86491
    
    Open sNomeArquivo For Input As #1
    
    'Busca o próximo registro do arquivo
    Line Input #1, sRegistro

    Do While Not EOF(1)

        iTipoRegistro = StrParaInt(Mid(sRegistro, 1, InStr(sRegistro, Chr(vbKeyEscape)) - 1))
                
        Select Case iTipoRegistro
        
            Case TIPO_REGISTRO_TRANSPORTADORA
                
                lErro = ProcessaArqNFs_Grava_Transportadora(sRegistro)
                If lErro <> SUCESSO Then gError 86492
                Line Input #1, sRegistro
            
            Case TIPO_REGISTRO_NATUREZAOP
                
                lErro = ProcessaArqNFs_Grava_NaturezaOp(sRegistro)
                If lErro <> SUCESSO Then gError 86493
                Line Input #1, sRegistro

            Case TIPO_REGISTRO_PRODUTO
            
                Set objProduto = New ClassProduto
                
                lErro = ProcessaArqNFs_Obtem_Produto(sRegistro, objProduto)
                If lErro <> SUCESSO Then gError 86494
                
                Line Input #1, sRegistro
                iTipoRegistro = StrParaInt(Mid(sRegistro, 1, InStr(sRegistro, Chr(vbKeyEscape)) - 1))

                Do While iTipoRegistro = TIPO_REGISTRO_PRODUTOCATEGORIA
                     
                    lErro = ProcessaArqNFs_Obtem_ProdutoCategoria(sRegistro, objProduto)
                    If lErro <> SUCESSO Then gError 86495
                                      
                    Line Input #1, sRegistro
                    iTipoRegistro = StrParaInt(Mid(sRegistro, 1, InStr(sRegistro, Chr(vbKeyEscape)) - 1))
                     
                Loop
                            
                lErro = ProcessaArqNFs_Grava_Produto(objProduto)
                If lErro <> SUCESSO Then gError 86496
        
            Case TIPO_REGISTRO_CLIENTE
        
                Set objCliente = New ClassCliente
                Set colFiliais = New Collection
                Set colcolEnderecos = New Collection
                
                lErro = ProcessaArqNFs_Obtem_Cliente(sRegistro, objCliente, colcolEnderecos)
                If lErro <> SUCESSO Then gError 86497
        
                Line Input #1, sRegistro
                iTipoRegistro = StrParaInt(Mid(sRegistro, 1, InStr(sRegistro, Chr(vbKeyEscape)) - 1))
        
                Do While iTipoRegistro = TIPO_REGISTRO_FILIALCLIENTE
                    
                    lErro = ProcessaArqNFs_Obtem_FilialCliente(sRegistro, colFiliais, colcolEnderecos)
                    If lErro <> SUCESSO Then gError 86498
                                      
                    Line Input #1, sRegistro
                    iTipoRegistro = StrParaInt(Mid(sRegistro, 1, InStr(sRegistro, Chr(vbKeyEscape)) - 1))
                
                Loop
                
                lErro = ProcessaArqNFs_Grava_Cliente_Filiais(objCliente, colFiliais, colcolEnderecos)
                If lErro <> SUCESSO Then gError 86499
        
            Case TIPO_REGISTRO_FORNECEDOR
        
                Set objFornecedor = New ClassFornecedor
                Set colFiliais = New Collection
                Set colEnderecos = New Collection
                
                lErro = ProcessaArqNFs_Obtem_Fornecedor(sRegistro, objFornecedor, colEnderecos)
                If lErro <> SUCESSO Then gError 86500
        
                Line Input #1, sRegistro
                iTipoRegistro = StrParaInt(Mid(sRegistro, 1, InStr(sRegistro, Chr(vbKeyEscape)) - 1))
        
                Do While iTipoRegistro = TIPO_REGISTRO_FILIALFORNECEDOR
                    
                    lErro = ProcessaArqNFs_Obtem_FilialFornecedor(sRegistro, colFiliais, colEnderecos)
                    If lErro <> SUCESSO Then gError 86501
                                      
                    Line Input #1, sRegistro
                    iTipoRegistro = StrParaInt(Mid(sRegistro, 1, InStr(sRegistro, Chr(vbKeyEscape)) - 1))
                
                Loop
                
                lErro = ProcessaArqNFs_Grava_Fornecedor_Filiais(objFornecedor, colFiliais, colEnderecos)
                If lErro <> SUCESSO Then gError 86502
        
            Case TIPO_REGISTRO_NOTAFISCAL
        
                Set objNFiscal = New ClassNFiscal
                Set colLancamentos = New Collection
                Set colLanPendente = New Collection
                
                lErro = ProcessaArqNFs_Obtem_NFiscal(sRegistro, objNFiscal)
                If lErro <> SUCESSO Then gError 86503
                
                Line Input #1, sRegistro
                iTipoRegistro = StrParaInt(Mid(sRegistro, 1, InStr(sRegistro, Chr(vbKeyEscape)) - 1))
        
                Do While iTipoRegistro = TIPO_REGISTRO_PARCELAPAG
                    lErro = ProcessaArqNFs_Obtem_ParcelaPag(sRegistro, objNFiscal)
                    If lErro <> SUCESSO Then gError 86504
                                      
                    If EOF(1) Then Exit Do
                    Line Input #1, sRegistro
                    iTipoRegistro = StrParaInt(Mid(sRegistro, 1, InStr(sRegistro, Chr(vbKeyEscape)) - 1))
                Loop
                
                Do While iTipoRegistro = TIPO_REGISTRO_PARCELAREC
                    lErro = ProcessaArqNFs_Obtem_ParcelaRec(sRegistro, objNFiscal)
                    If lErro <> SUCESSO Then gError 86505
                                      
                    If EOF(1) Then Exit Do
                    Line Input #1, sRegistro
                    iTipoRegistro = StrParaInt(Mid(sRegistro, 1, InStr(sRegistro, Chr(vbKeyEscape)) - 1))
                Loop
                
                Do While iTipoRegistro = TIPO_REGISTRO_LANCAMENTO
                    lErro = ProcessaArqNFs_Obtem_Lancamentos(sRegistro, colLancamentos)
                    If lErro <> SUCESSO Then gError 86506
                                      
                    If EOF(1) Then Exit Do
                    Line Input #1, sRegistro
                    iTipoRegistro = StrParaInt(Mid(sRegistro, 1, InStr(sRegistro, Chr(vbKeyEscape)) - 1))
                    
                Loop

                
                Do While iTipoRegistro = TIPO_REGISTRO_LANPENDENTE
                    lErro = ProcessaArqNFs_Obtem_Lancamentos(sRegistro, colLanPendente)
                    If lErro <> SUCESSO Then gError 86507
                                      
                    If EOF(1) Then Exit Do
                    Line Input #1, sRegistro
                    iTipoRegistro = StrParaInt(Mid(sRegistro, 1, InStr(sRegistro, Chr(vbKeyEscape)) - 1))
                
                Loop
                
                Do While iTipoRegistro = TIPO_REGISTRO_ITEMNOTAFISCAL
                    lErro = ProcessaArqNFs_Obtem_ItemNFiscal(sRegistro, objNFiscal)
                    If lErro <> SUCESSO Then gError 86508
                                      
                    If EOF(1) Then Exit Do
                    Line Input #1, sRegistro
                    iTipoRegistro = StrParaInt(Mid(sRegistro, 1, InStr(sRegistro, Chr(vbKeyEscape)) - 1))
                
                    Do While iTipoRegistro = TIPO_REGISTRO_ALOCACAOITEMNF
                        lErro = ProcessaArqNFs_Obtem_AlocacaoItemNF(sRegistro, objNFiscal)
                        If lErro <> SUCESSO Then gError 86509
                                          
                        If EOF(1) Then Exit Do
                        Line Input #1, sRegistro
                        iTipoRegistro = StrParaInt(Mid(sRegistro, 1, InStr(sRegistro, Chr(vbKeyEscape)) - 1))
                    Loop
                
                    Do While iTipoRegistro = TIPO_REGISTRO_ITEMNFGRADE
                        lErro = ProcessaArqNFs_Obtem_ItemGradeItemNF(sRegistro, objNFiscal)
                        If lErro <> SUCESSO Then gError 86510
                                          
                        If EOF(1) Then Exit Do
                        Line Input #1, sRegistro
                        iTipoRegistro = StrParaInt(Mid(sRegistro, 1, InStr(sRegistro, Chr(vbKeyEscape)) - 1))
                        
                        Do While iTipoRegistro = TIPO_REGISTRO_LOCALIZACAOITEMNFGRADE
                            lErro = ProcessaArqNFs_Obtem_LocalizacaoItemNFGrade(sRegistro, objNFiscal)
                            If lErro <> SUCESSO Then gError 86511
                                              
                            If EOF(1) Then Exit Do
                            Line Input #1, sRegistro
                            iTipoRegistro = StrParaInt(Mid(sRegistro, 1, InStr(sRegistro, Chr(vbKeyEscape)) - 1))
                        Loop
                    
                    Loop
                Loop
                                
                lErro = ProcessaArqNFs_Grava_NFiscal(objNFiscal, colLancamentos, colLanPendente)
                If lErro <> SUCESSO Then gError 86512
                
        End Select
        
    Loop
    
    Close #1
    
    lErro = Transacao_Commit()
    If lErro <> SUCESSO Then Error 86513
    
    NFiscais_Importar = SUCESSO
    
    Exit Function
    
Erro_NFiscais_Importar:

    NFiscais_Importar = gErr

    Select Case gErr
    
        Case 86491
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 86492 To 86512
        
        Case 86513
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT_TRANSACAO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150837)
    
    End Select
    
    Call Transacao_Rollback
    
    Close #1
    
    Exit Function

End Function

Function ProcessaArqNFs_Grava_Transportadora(sRegistro As String) As Long

Dim objTransportadora As New ClassTransportadora
Dim objEndereco As New ClassEndereco
Dim iPosicaoIni As Integer
Dim iPosicaoFim As Integer
Dim lErro As Long

On Error GoTo Erro_ProcessaArqNFs_Grava_Transportadora

    iPosicaoIni = InStr(sRegistro, Chr(vbKeyEscape)) + 1
    iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    With objTransportadora
        .iCodigo = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sNome = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sNomeReduzido = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sCgc = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sInscricaoEstadual = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iViaTransporte = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sInscricaoMunicipal = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sGuia = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dPesoMinimo = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    End With
    
    objEndereco.sEndereco = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEndereco.sBairro = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEndereco.sCidade = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEndereco.sSiglaEstado = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEndereco.iCodigoPais = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEndereco.sCEP = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEndereco.sTelefone1 = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEndereco.sTelefone2 = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEndereco.sEmail = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEndereco.sFax = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEnd)) - 1
    objEndereco.sContato = Mid(sRegistro, iPosicaoIni, iPosicaoFim)
    
    lErro = Cidade_Importa(objEndereco.sCidade)
    If lErro <> SUCESSO Then gError 130007
    
    lErro = CF("Transportadora_Grava_EmTrans", objTransportadora, objEndereco)
    If lErro <> SUCESSO Then gError 86514

    ProcessaArqNFs_Grava_Transportadora = SUCESSO
        
    Exit Function
    
Erro_ProcessaArqNFs_Grava_Transportadora:

    ProcessaArqNFs_Grava_Transportadora = gErr
    
    Select Case gErr
            
        Case 86514, 130007
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150838)
    
    End Select

    Exit Function

End Function

Function ProcessaArqNFs_Grava_NaturezaOp(sRegistro As String) As Long

Dim objNaturezaOp As New ClassNaturezaOp
Dim lErro As Long
Dim iPosicaoIni As Integer
Dim iPosicaoFim As Integer

On Error GoTo Erro_ProcessaArqNFs_Grava_NaturezaOp
    
    iPosicaoIni = InStr(sRegistro, Chr(vbKeyEscape)) + 1
    iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    With objNaturezaOp
        .sCodigo = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sDescricao = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEnd)) - 1
        .sDescrNF = Mid(sRegistro, iPosicaoIni, iPosicaoFim)
    End With

    lErro = CF("NaturezaOperacao_Grava_EmTrans", objNaturezaOp)
    If lErro <> SUCESSO Then Error 86515
    
    ProcessaArqNFs_Grava_NaturezaOp = SUCESSO
    
    Exit Function
    
Erro_ProcessaArqNFs_Grava_NaturezaOp:

    ProcessaArqNFs_Grava_NaturezaOp = gErr

    Select Case gErr
    
        Case 86515
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150839)
    
    End Select
    
    Exit Function

End Function

Function ProcessaArqNFs_Obtem_Produto(sRegistro As String, objProduto As ClassProduto) As Long

Dim iPosicaoIni As Integer
Dim iPosicaoFim As Integer

On Error GoTo Erro_ProcessaArqNFs_Obtem_Produto

    iPosicaoIni = InStr(sRegistro, Chr(vbKeyEscape)) + 1
    iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1

    With objProduto
        .sCodigo = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sDescricao = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iGerencial = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iAtivo = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iFaturamento = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iClasseUM = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sSiglaUMEstoque = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sSiglaUMCompra = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sSiglaUMVenda = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sSubstituto1 = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sSubstituto2 = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iNivel = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iTipo = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sNomeReduzido = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sModelo = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iPrazoValidade = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sCodigoBarras = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iEtiquetasCodBarras = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dPesoLiq = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dPesoBruto = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dComprimento = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dEspessura = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dLargura = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sCor = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sObsFisica = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iCompras = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iPCP = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iKitBasico = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iKitInt = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dIPIAliquota = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sIPICodigo = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sIPICodDIPI = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iControleEstoque = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iApropriacaoCusto = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sContaContabil = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sContaContabilProducao = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dResiduo = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iNatureza = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dCustoReposicao = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iOrigemMercadoria = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iTabelaPreco = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dPercentMaisQuantCotAnt = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dPercentMenosQuantCotAnt = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iConsideraQuantCotAnt = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iTemFaixaReceb = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dPercentMaisReceb = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dPercentMenosReceb = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iRecebForaFaixa = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iTempoProducao = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iRastro = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .lHorasMaquina = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dPesoEspecifico = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iEmbalagem = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sReferencia = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sFigura = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dINSSPercBase = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iKitVendaComp = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sGrade = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sSituacaoTribECF = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1:  iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEnd)) - 1
        .sICMSAliquota = Mid(sRegistro, iPosicaoIni, iPosicaoFim)
        
    End With

    ProcessaArqNFs_Obtem_Produto = SUCESSO
    
    Exit Function

Erro_ProcessaArqNFs_Obtem_Produto:

    ProcessaArqNFs_Obtem_Produto = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150840)
    
    End Select
    
    Exit Function

End Function

Function ProcessaArqNFs_Obtem_ProdutoCategoria(sRegistro As String, objProduto As ClassProduto) As Long

Dim objProdutoCateg As New ClassProdutoCategoria
Dim iPosicaoIni As Integer
Dim iPosicaoFim As Integer

On Error GoTo Erro_ProcessaArqNFs_Obtem_ProdutoCategoria

    iPosicaoIni = InStr(sRegistro, Chr(vbKeyEscape)) + 1
    iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1

    objProdutoCateg.sCategoria = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1:  iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEnd)) - 1
    objProdutoCateg.sItem = Mid(sRegistro, iPosicaoIni, iPosicaoFim)

    objProduto.colCategoriaItem.Add objProdutoCateg

    ProcessaArqNFs_Obtem_ProdutoCategoria = SUCESSO
    
    Exit Function

Erro_ProcessaArqNFs_Obtem_ProdutoCategoria:

    ProcessaArqNFs_Obtem_ProdutoCategoria = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150841)
    
    End Select
    
    Exit Function

End Function

Function ProcessaArqNFs_Grava_Produto(objProduto As ClassProduto) As Long

Dim lErro As Long
Dim colTabelaPrecoItem As New Collection

On Error GoTo Erro_ProcessaArqNFs_Grava_Produto

    lErro = CF("Produto_Grava_Trans", objProduto, colTabelaPrecoItem)
    If lErro <> SUCESSO Then gError 86516

    ProcessaArqNFs_Grava_Produto = SUCESSO
    
    Exit Function
    
Erro_ProcessaArqNFs_Grava_Produto:

    ProcessaArqNFs_Grava_Produto = gErr

    Select Case gErr
    
        Case 86516
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150842)
    
    End Select
    
    Exit Function

End Function

Function ProcessaArqNFs_Obtem_Cliente(sRegistro As String, objCliente As ClassCliente, colcolEnderecos As Collection) As Long

Dim lErro As Long
Dim iPosicaoIni As Integer
Dim iPosicaoFim As Integer
Dim objEndereco As New ClassEndereco
Dim objEnderecoEntrega As New ClassEndereco
Dim objEnderecoCobranca As New ClassEndereco
Dim colEnderecos As New Collection

On Error GoTo Erro_ProcessaArqNFs_Obtem_Cliente

    iPosicaoIni = InStr(sRegistro, Chr(vbKeyEscape)) + 1
    iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1

    With objCliente
        .lCodigo = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sRazaoSocial = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sNomeReduzido = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iTipo = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sObservacao = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dLimiteCredito = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iCondicaoPagto = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dDesconto = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iCodMensagem = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iTabelaPreco = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iProxCodFilial = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iCodPadraoCobranca = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iAtivo = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iTransferido = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sCgc = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sInscricaoEstadual = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sInscricaoMunicipal = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sInscricaoSuframa = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .lEndereco = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .lEnderecoEntrega = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .lEnderecoCobranca = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iCodTransportadora = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sObservacao2 = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sContaContabil = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iVendedor = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dComissaoVendas = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iRegiao = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iFreqVisitas = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dtDataUltVisita = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iCodCobrador = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iTipoFrete = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sRG = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iCodTranspRedesp = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iDetPagFrete = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sGuia = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    End With

    objEndereco.sEndereco = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEndereco.sBairro = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEndereco.sCidade = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEndereco.sSiglaEstado = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEndereco.iCodigoPais = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEndereco.sCEP = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEndereco.sTelefone1 = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEndereco.sTelefone2 = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEndereco.sEmail = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEndereco.sFax = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEndereco.sContato = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEnderecoCobranca.sEndereco = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEnderecoCobranca.sBairro = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEnderecoCobranca.sCidade = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEnderecoCobranca.sSiglaEstado = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEnderecoCobranca.iCodigoPais = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEnderecoCobranca.sCEP = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEnderecoCobranca.sTelefone1 = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEnderecoCobranca.sTelefone2 = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEnderecoCobranca.sEmail = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEnderecoCobranca.sFax = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEnderecoCobranca.sContato = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEnderecoEntrega.sEndereco = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEnderecoEntrega.sBairro = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEnderecoEntrega.sCidade = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEnderecoEntrega.sSiglaEstado = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEnderecoEntrega.iCodigoPais = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEnderecoEntrega.sCEP = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEnderecoEntrega.sTelefone1 = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEnderecoEntrega.sTelefone2 = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEnderecoEntrega.sEmail = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEnderecoEntrega.sFax = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEnd)) - 1
    objEnderecoEntrega.sContato = Mid(sRegistro, iPosicaoIni, iPosicaoFim)
    
    lErro = Cidade_Importa(objEndereco.sCidade)
    If lErro <> SUCESSO Then gError 130010
    
    lErro = Cidade_Importa(objEnderecoCobranca.sCidade)
    If lErro <> SUCESSO Then gError 130011
    
    lErro = Cidade_Importa(objEnderecoEntrega.sCidade)
    If lErro <> SUCESSO Then gError 130012
    
    colEnderecos.Add objEndereco
    colEnderecos.Add objEnderecoCobranca
    colEnderecos.Add objEnderecoEntrega
    
    colcolEnderecos.Add colEnderecos
    
    ProcessaArqNFs_Obtem_Cliente = SUCESSO
    
    Exit Function

Erro_ProcessaArqNFs_Obtem_Cliente:

    ProcessaArqNFs_Obtem_Cliente = gErr

    Select Case gErr
    
        Case 130010, 130011, 130012
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150843)
    
    End Select
    
    Exit Function

End Function

Function ProcessaArqNFs_Obtem_FilialCliente(sRegistro As String, colFiliais As Collection, colcolEnderecos As Collection) As Long

Dim lErro As Long
Dim iPosicaoIni As Integer
Dim iPosicaoFim As Integer
Dim objFilialCli As New ClassFilialCliente
Dim objEndereco As New ClassEndereco
Dim objEnderecoEntrega As New ClassEndereco
Dim objEnderecoCobranca As New ClassEndereco
Dim colEnderecos As New Collection


On Error GoTo Erro_ProcessaArqNFs_Obtem_FilialCliente

    iPosicaoIni = InStr(sRegistro, Chr(vbKeyEscape)) + 1
    iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1

    With objFilialCli
        .lCodCliente = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iCodFilial = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sNome = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sCgc = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sInscricaoEstadual = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sInscricaoMunicipal = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sInscricaoSuframa = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .lEndereco = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .lEnderecoEntrega = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .lEnderecoCobranca = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iCodTransportadora = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sObservacao = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sContaContabil = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iVendedor = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dComissaoVendas = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iRegiao = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iFreqVisitas = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dtDataUltVisita = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iCodCobrador = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iTipoFrete = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sRG = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iAtivo = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iCodTranspRedesp = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iDetPagFrete = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sGuia = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        
    End With

    objEndereco.sEndereco = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEndereco.sBairro = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEndereco.sCidade = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEndereco.sSiglaEstado = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEndereco.iCodigoPais = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEndereco.sCEP = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEndereco.sTelefone1 = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEndereco.sTelefone2 = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEndereco.sEmail = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEndereco.sFax = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEndereco.sContato = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEnderecoCobranca.sEndereco = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEnderecoCobranca.sBairro = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEnderecoCobranca.sCidade = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEnderecoCobranca.sSiglaEstado = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEnderecoCobranca.iCodigoPais = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEnderecoCobranca.sCEP = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEnderecoCobranca.sTelefone1 = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEnderecoCobranca.sTelefone2 = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEnderecoCobranca.sEmail = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEnderecoCobranca.sFax = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEnderecoCobranca.sContato = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEnderecoEntrega.sEndereco = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEnderecoEntrega.sBairro = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEnderecoEntrega.sCidade = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEnderecoEntrega.sSiglaEstado = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEnderecoEntrega.iCodigoPais = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEnderecoEntrega.sCEP = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEnderecoEntrega.sTelefone1 = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEnderecoEntrega.sTelefone2 = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEnderecoEntrega.sEmail = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEnderecoEntrega.sFax = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEnd)) - 1
    objEnderecoEntrega.sContato = Mid(sRegistro, iPosicaoIni, iPosicaoFim)
    
    lErro = Cidade_Importa(objEndereco.sCidade)
    If lErro <> SUCESSO Then gError 130010
    
    lErro = Cidade_Importa(objEnderecoCobranca.sCidade)
    If lErro <> SUCESSO Then gError 130011
    
    lErro = Cidade_Importa(objEnderecoEntrega.sCidade)
    If lErro <> SUCESSO Then gError 130012
    
    colEnderecos.Add objEndereco
    colEnderecos.Add objEnderecoCobranca
    colEnderecos.Add objEnderecoEntrega
    
    colFiliais.Add objFilialCli
    colcolEnderecos.Add colEnderecos
    
    ProcessaArqNFs_Obtem_FilialCliente = SUCESSO
    
    Exit Function
    
Erro_ProcessaArqNFs_Obtem_FilialCliente:

    ProcessaArqNFs_Obtem_FilialCliente = gErr

    Select Case gErr
    
        Case 130010, 130011, 130012
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150844)
    
    End Select
    
    Exit Function

End Function

Function ProcessaArqNFs_Grava_Cliente_Filiais(objCliente As ClassCliente, colFiliais As Collection, colcolEnderecos As Collection) As Long

Dim objFilialCli As ClassFilialCliente
Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_ProcessaArqNFs_Grava_Cliente_Filiais

    iIndice = 1

    lErro = CF("Cliente_Grava_EmTrans", objCliente, colcolEnderecos(iIndice))
    If lErro <> SUCESSO Then gError 86517
    
    For Each objFilialCli In colFiliais
        iIndice = iIndice + 1
        lErro = CF("FiliaisClientes_Grava_EmTrans", objFilialCli, colcolEnderecos(iIndice))
        If lErro <> SUCESSO Then gError 86518
    
    Next

    ProcessaArqNFs_Grava_Cliente_Filiais = SUCESSO
    
    Exit Function
    
Erro_ProcessaArqNFs_Grava_Cliente_Filiais:

    ProcessaArqNFs_Grava_Cliente_Filiais = gErr

    Select Case gErr
    
        Case 86517, 86518
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150845)
    
    End Select
    
    Exit Function

End Function

Function ProcessaArqNFs_Obtem_Fornecedor(sRegistro As String, objFornecedor As ClassFornecedor, colEnderecos As Collection) As Long

Dim objEndereco As New ClassEndereco
Dim iPosicaoIni As Integer
Dim iPosicaoFim As Integer, lErro As Long

On Error GoTo Erro_ProcessaArqNFs_Obtem_Fornecedor

    iPosicaoIni = InStr(sRegistro, Chr(vbKeyEscape)) + 1
    iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1

    With objFornecedor
        .sRazaoSocial = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sNomeReduzido = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iTipo = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sObservacao = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iCondicaoPagto = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dDesconto = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iProxCodFilial = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .lCodigo = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .lEndereco = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sCgc = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sInscricaoEstadual = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sInscricaoMunicipal = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sContaContabil = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sContaFornConsig = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iBanco = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sAgencia = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sContaCorrente = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sObservacao2 = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iTipoFrete = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sInscricaoINSS = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iTipoCobranca = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iGeraCredICMS = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sContaDespesa = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    End With
    
    objEndereco.sEndereco = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEndereco.sBairro = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEndereco.sCidade = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEndereco.sSiglaEstado = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEndereco.iCodigoPais = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEndereco.sCEP = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEndereco.sTelefone1 = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEndereco.sTelefone2 = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEndereco.sEmail = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEndereco.sFax = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEnd)) - 1
    objEndereco.sContato = Mid(sRegistro, iPosicaoIni, iPosicaoFim)
    
    lErro = Cidade_Importa(objEndereco.sCidade)
    If lErro <> SUCESSO Then gError 130010
    
    colEnderecos.Add objEndereco
        
    ProcessaArqNFs_Obtem_Fornecedor = SUCESSO
    
    Exit Function
    
Erro_ProcessaArqNFs_Obtem_Fornecedor:

    ProcessaArqNFs_Obtem_Fornecedor = gErr
    
    Select Case gErr
    
        Case 130010
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150846)
    
    End Select
    
    Exit Function

End Function

Function ProcessaArqNFs_Obtem_FilialFornecedor(sRegistro As String, colFiliais As Collection, colEnderecos As Collection) As Long

Dim iPosicaoIni As Integer, lErro As Long
Dim iPosicaoFim As Integer
Dim objFilialForn As New ClassFilialFornecedor
Dim objEndereco As New ClassEndereco

On Error GoTo Erro_ProcessaArqNFs_Obtem_FilialFornecedor


    iPosicaoIni = InStr(sRegistro, Chr(vbKeyEscape)) + 1
    iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1

    With objFilialForn
        .sNome = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .lEndereco = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sCgc = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sInscricaoEstadual = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sInscricaoMunicipal = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sContaContabil = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sContaFornConsig = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iBanco = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sAgencia = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sContaCorrente = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sObservacao = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iTipoFrete = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sInscricaoINSS = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iTipoCobranca = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iGeraCredICMS = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sContaDespesa = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .lCodFornecedor = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iCodFilial = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        
    End With
    
    objEndereco.sEndereco = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEndereco.sBairro = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEndereco.sCidade = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEndereco.sSiglaEstado = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEndereco.iCodigoPais = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEndereco.sCEP = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEndereco.sTelefone1 = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEndereco.sTelefone2 = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEndereco.sEmail = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objEndereco.sFax = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEnd)) - 1
    objEndereco.sContato = Mid(sRegistro, iPosicaoIni, iPosicaoFim)
        
    lErro = Cidade_Importa(objEndereco.sCidade)
    If lErro <> SUCESSO Then gError 130010
    
    colEnderecos.Add objEndereco
    
    ProcessaArqNFs_Obtem_FilialFornecedor = SUCESSO
    
    Exit Function

Erro_ProcessaArqNFs_Obtem_FilialFornecedor:

    ProcessaArqNFs_Obtem_FilialFornecedor = gErr
    
    Select Case gErr
    
        Case 130010
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150847)
    
    End Select
    
    Exit Function

End Function

Function ProcessaArqNFs_Grava_Fornecedor_Filiais(objFornecedor As ClassFornecedor, colFiliais As Collection, colEnderecos As Collection) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objFilialForn As ClassFilialFornecedor

On Error GoTo Erro_ProcessaArqNFs_Grava_Fornecedor_Filiais

    iIndice = 1
    
    lErro = CF("Fornecedor_Grava_EmTrans", objFornecedor, colEnderecos(iIndice))
    If lErro <> SUCESSO Then gError 86519
    
    For Each objFilialForn In colFiliais
    
        iIndice = iIndice + 1
        
        lErro = CF("FiliaisFornecedores_Grava_EmTrans", objFilialForn, colEnderecos(iIndice))
        If lErro <> SUCESSO Then gError 86520

    Next

    ProcessaArqNFs_Grava_Fornecedor_Filiais = SUCESSO
    
    Exit Function
    
Erro_ProcessaArqNFs_Grava_Fornecedor_Filiais:

    ProcessaArqNFs_Grava_Fornecedor_Filiais = gErr

    Select Case gErr
    
        Case 86519, 86520
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150848)
    
    End Select
    
    Exit Function

End Function

Function ProcessaArqNFs_Obtem_NFiscal(sRegistro As String, objNFiscal As ClassNFiscal) As Long

Dim iPosicaoIni As Integer
Dim iPosicaoFim As Integer

On Error GoTo Erro_ProcessaArqNFs_Obtem_NFiscal

    iPosicaoIni = InStr(sRegistro, Chr(vbKeyEscape)) + 1
    iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1

    With objNFiscal
        .lNumIntDoc = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iStatus = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iFilialEmpresa = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sSerie = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .lNumNotaFiscal = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .lCliente = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iFilialCli = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iFilialEntrega = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .lFornecedor = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iFilialForn = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dtDataEmissao = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dtDataEntrada = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dtDataSaida = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dtDataVencimento = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dtDataReferencia = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iFilialPedido = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sNumPedidoTerc = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iClasseDocCPR = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .lNumIntDocCPR = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dValorTotal = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dValorProdutos = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dValorFrete = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dValorSeguro = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dValorOutrasDespesas = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dValorDesconto = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iCodTransportadora = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sMensagemNota = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iTabelaPreco = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iTipoNFiscal = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sNaturezaOp = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dPesoLiq = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dPesoBruto = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .lNumIntTrib = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sPlaca = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sPlacaUF = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .lVolumeQuant = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .lVolumeEspecie = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .lVolumeMarca = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sVolumeNumero = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iCanal = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .lNumIntNotaOriginal = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .lClienteBenef = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iFilialCliBenef = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .lFornecedorBenef = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iFilialFornBenef = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iFreteRespons = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .lNumRecebimento = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sObservacao = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sCodUsuarioCancel = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sMotivoCancel = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dtHoraEntrada = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dtHoraSaida = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iCodTranspRedesp = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iDetPagFrete = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iSemDataSaida = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iNaoImpCobranca = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        
    End With
           
'    Set objNFiscal.objTributacaoNF = New ClassTributacaoNF
'    Set objNFiscal.objTributacaoNF.objTributacaoDesconto = New ClassTribComplNF
'    Set objNFiscal.objTributacaoNF.objTributacaoFrete = New ClassTribComplNF
'    Set objNFiscal.objTributacaoNF.objTributacaoOutras = New ClassTribComplNF
'    Set objNFiscal.objTributacaoNF.objTributacaoSeguro = New ClassTribComplNF
'XXXX
'VER COMO TRATAR
    
    With objNFiscal.objTributacaoNF
        .sNaturezaOpInterna = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iTipoTributacao = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dIPIBase = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dIPIValor = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dIPICredito = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dICMSBase = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dICMSValor = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dICMSSubstBase = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dICMSSubstValor = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dICMSCredito = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iISSIncluso = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dISSBase = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dISSAliquota = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dISSValor = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dIRRFBase = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dIRRFAliquota = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dIRRFValor = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dValorINSS = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iINSSRetido = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dINSSBase = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dINSSDeducoes = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dPISCredito = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dCOFINSCredito = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dPISRetido = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dCOFINSRetido = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dCSLLRetido = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        
    End With
    
'    With objNFiscal.objTributacaoNF.objTributacaoDesconto
'        .dCOFINSCredito = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .dICMSAliquota = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .dICMSBase = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .dICMSCredito = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .dICMSPercRedBase = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .dICMSSubstAliquota = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .dICMSSubstBase = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .dICMSSubstValor = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .dICMSValor = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .dIPIAliquota = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .dIPIBaseCalculo = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .dIPICredito = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .dIPIPercRedBase = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .dIPIValor = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .dPISCredito = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .iICMSAliquotaManual = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .iICMSBaseManual = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .iICMSPercRedBaseManual = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .iICMSSubstAliquotaManual = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .iICMSSubstBaseManual = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .iICMSTipo = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .iICMSTipoManual = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .iICMSValorManual = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .iIPIAliquotaManual = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .iIPIBaseManual = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .iIPIPercRedBaseManual = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .iIPITipo = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .iIPITipoManual = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .iIPIValorManual = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .iNaturezaOpManual = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .iTipo = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .iTipoTributacao = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .iTipoTributacaoManual = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .lNumIntDoc = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .sNaturezaOp = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'
'    End With
'    With objNFiscal.objTributacaoNF.objTributacaoFrete
'        .dCOFINSCredito = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .dICMSAliquota = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .dICMSBase = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .dICMSCredito = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .dICMSPercRedBase = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .dICMSSubstAliquota = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .dICMSSubstBase = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .dICMSSubstValor = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .dICMSValor = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .dIPIAliquota = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .dIPIBaseCalculo = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .dIPICredito = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .dIPIPercRedBase = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .dIPIValor = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .dPISCredito = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .iICMSAliquotaManual = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .iICMSBaseManual = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .iICMSPercRedBaseManual = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .iICMSSubstAliquotaManual = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .iICMSSubstBaseManual = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .iICMSTipo = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .iICMSTipoManual = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .iICMSValorManual = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .iIPIAliquotaManual = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .iIPIBaseManual = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .iIPIPercRedBaseManual = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .iIPITipo = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .iIPITipoManual = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .iIPIValorManual = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .iNaturezaOpManual = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .iTipo = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .iTipoTributacao = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .iTipoTributacaoManual = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .lNumIntDoc = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .sNaturezaOp = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'
'    End With
'    With objNFiscal.objTributacaoNF.objTributacaoOutras
'        .dCOFINSCredito = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .dICMSAliquota = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .dICMSBase = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .dICMSCredito = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .dICMSPercRedBase = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .dICMSSubstAliquota = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .dICMSSubstBase = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .dICMSSubstValor = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .dICMSValor = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .dIPIAliquota = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .dIPIBaseCalculo = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .dIPICredito = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .dIPIPercRedBase = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .dIPIValor = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .dPISCredito = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .iICMSAliquotaManual = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .iICMSBaseManual = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .iICMSPercRedBaseManual = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .iICMSSubstAliquotaManual = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .iICMSSubstBaseManual = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .iICMSTipo = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .iICMSTipoManual = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .iICMSValorManual = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .iIPIAliquotaManual = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .iIPIBaseManual = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .iIPIPercRedBaseManual = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .iIPITipo = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .iIPITipoManual = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .iIPIValorManual = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .iNaturezaOpManual = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .iTipo = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .iTipoTributacao = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .iTipoTributacaoManual = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .lNumIntDoc = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .sNaturezaOp = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'
'    End With
'    With objNFiscal.objTributacaoNF.objTributacaoSeguro
'        .dCOFINSCredito = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .dICMSAliquota = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .dICMSBase = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .dICMSCredito = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .dICMSPercRedBase = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .dICMSSubstAliquota = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .dICMSSubstBase = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .dICMSSubstValor = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .dICMSValor = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .dIPIAliquota = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .dIPIBaseCalculo = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .dIPICredito = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .dIPIPercRedBase = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .dIPIValor = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .dPISCredito = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .iICMSAliquotaManual = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .iICMSBaseManual = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .iICMSPercRedBaseManual = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .iICMSSubstAliquotaManual = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .iICMSSubstBaseManual = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .iICMSTipo = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .iICMSTipoManual = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .iICMSValorManual = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .iIPIAliquotaManual = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .iIPIBaseManual = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .iIPIPercRedBaseManual = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .iIPITipo = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .iIPITipoManual = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .iIPIValorManual = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .iNaturezaOpManual = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .iTipo = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .iTipoTributacao = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .iTipoTributacaoManual = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
'        .lNumIntDoc = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEnd)) - 1
'        .sNaturezaOp = Mid(sRegistro, iPosicaoIni, iPosicaoFim)
'
'    End With

    ProcessaArqNFs_Obtem_NFiscal = SUCESSO

    Exit Function

Erro_ProcessaArqNFs_Obtem_NFiscal:

    ProcessaArqNFs_Obtem_NFiscal = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150849)
    
    End Select
    
    Exit Function

End Function

Function ProcessaArqNFs_Obtem_ParcelaPag(sRegistro As String, objNFiscal As ClassNFiscal) As Long

Dim iPosicaoIni As Integer
Dim iPosicaoFim As Integer
Dim objParcelaPag As New ClassParcelaPagar

On Error GoTo Erro_ProcessaArqNFs_Obtem_ParcelaPag

    iPosicaoIni = InStr(sRegistro, Chr(vbKeyEscape)) + 1
    iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1

    With objParcelaPag
        .lNumIntDoc = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iNumParcela = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iStatus = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dtDataVencimento = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dtDataVencimentoReal = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dSaldo = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dValor = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iPortador = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iProxSeqBaixa = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iTipoCobranca = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iBancoCobrador = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sNossoNumero = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEnd)) - 1
        .sCodigoDeBarras = Mid(sRegistro, iPosicaoIni, iPosicaoFim)
        
    End With

    ProcessaArqNFs_Obtem_ParcelaPag = SUCESSO

    Exit Function

Erro_ProcessaArqNFs_Obtem_ParcelaPag:

    ProcessaArqNFs_Obtem_ParcelaPag = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150850)
    
    End Select
    
    Exit Function

End Function

Function ProcessaArqNFs_Obtem_ParcelaRec(sRegistro As String, objNFiscal As ClassNFiscal) As Long

Dim iPosicaoIni As Integer
Dim iPosicaoFim As Integer
Dim objParcelaRec As New ClassParcelaReceber

On Error GoTo Erro_ProcessaArqNFs_Obtem_ParcelaRec

    iPosicaoIni = InStr(sRegistro, Chr(vbKeyEscape)) + 1
    iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1

    With objParcelaRec
        .lNumIntDoc = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iStatus = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dtDataVencimento = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dValor = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iCarteiraCobranca = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iCobrador = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iDesconto1Codigo = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dtDesconto1Ate = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dDesconto1Valor = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iDesconto2Codigo = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dtDesconto2Ate = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dDesconto2Valor = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iDesconto3Codigo = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dtDesconto3Ate = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEnd)) - 1
        .dDesconto3Valor = Mid(sRegistro, iPosicaoIni, iPosicaoFim)
        
    End With

    ProcessaArqNFs_Obtem_ParcelaRec = SUCESSO

    Exit Function

Erro_ProcessaArqNFs_Obtem_ParcelaRec:

    ProcessaArqNFs_Obtem_ParcelaRec = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150851)
    
    End Select
    
    Exit Function

End Function

Function ProcessaArqNFs_Obtem_ItemNFiscal(sRegistro As String, objNFiscal As ClassNFiscal) As Long

Dim iPosicaoIni As Integer
Dim iPosicaoFim As Integer
Dim objItemNF As New ClassItemNF

On Error GoTo Erro_ProcessaArqNFs_Obtem_ItemNFiscal

    iPosicaoIni = InStr(sRegistro, Chr(vbKeyEscape)) + 1
    iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    
    With objItemNF
        .iItem = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sProduto = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sUnidadeMed = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dQuantidade = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dPrecoUnitario = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dPercDesc = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dValorDesconto = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dtDataEntrega = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sDescricaoItem = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dValorAbatComissao = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .lNumIntPedVenda = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .lNumIntItemPedVenda = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .lNumIntDoc = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .lNumIntTrib = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .lNumIntDocOrig = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sCcl = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iAlmoxarifado = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        
    End With


    Set objItemNF.objTributacaoItemNF = New ClassTributacaoDocItem 'ClassTribItemNF

    With objItemNF.objTributacaoItemNF
        .iItem = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .lNumIntDoc = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sNaturezaOp = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iTipoTributacao = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iIPITipo = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sIPICodProduto = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dIPIBaseCalculo = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dIPIPercRedBase = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dIPIAliquota = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dIPIValor = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dIPICredito = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .iICMSTipo = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dICMSBase = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dICMSPercRedBase = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dICMSAliquota = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dICMSValor = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dICMSCredito = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dICMSSubstBase = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dICMSSubstAliquota = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dICMSSubstValor = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dPISCredito = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEnd)) - 1
        .dCOFINSCredito = Mid(sRegistro, iPosicaoIni, iPosicaoFim)
        
    End With

    objNFiscal.colItensNF.Add1 objItemNF
    
    ProcessaArqNFs_Obtem_ItemNFiscal = SUCESSO
    
    Exit Function

Erro_ProcessaArqNFs_Obtem_ItemNFiscal:

    ProcessaArqNFs_Obtem_ItemNFiscal = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150852)
        
    End Select
    
    Exit Function

End Function

Function ProcessaArqNFs_Obtem_AlocacaoItemNF(sRegistro As String, objNFiscal As ClassNFiscal) As Long

Dim iPosicaoIni As Integer
Dim iPosicaoFim As Integer
Dim objAlocacaoNF As New ClassItemNFAlocacao

On Error GoTo Erro_ProcessaArqNFs_Obtem_AlocacaoItemNF

    iPosicaoIni = InStr(sRegistro, Chr(vbKeyEscape)) + 1
    iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    
    objAlocacaoNF.dQuantidade = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objAlocacaoNF.iAlmoxarifado = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objAlocacaoNF.iTransferencia = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objAlocacaoNF.sAlmoxarifado = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEnd)) - 1
    objAlocacaoNF.sUnidadeMed = Mid(sRegistro, iPosicaoIni, iPosicaoFim)
    
    objNFiscal.colItensNF(objNFiscal.colItensNF.Count).sUMEstoque = objAlocacaoNF.sUnidadeMed
    objNFiscal.colItensNF(objNFiscal.colItensNF.Count).colAlocacoes.Add objAlocacaoNF.iAlmoxarifado, objAlocacaoNF.sAlmoxarifado, objAlocacaoNF.dQuantidade
    objNFiscal.colItensNF(objNFiscal.colItensNF.Count).colAlocacoes(objNFiscal.colItensNF(objNFiscal.colItensNF.Count).colAlocacoes.Count).iTransferencia = objAlocacaoNF.iTransferencia
    objNFiscal.colItensNF(objNFiscal.colItensNF.Count).colAlocacoes(objNFiscal.colItensNF(objNFiscal.colItensNF.Count).colAlocacoes.Count).sUnidadeMed = objAlocacaoNF.sUnidadeMed
    ProcessaArqNFs_Obtem_AlocacaoItemNF = SUCESSO
    
    Exit Function
    
Erro_ProcessaArqNFs_Obtem_AlocacaoItemNF:

    ProcessaArqNFs_Obtem_AlocacaoItemNF = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150853)
    
    End Select
    
    Exit Function

End Function

Function ProcessaArqNFs_Obtem_ItemGradeItemNF(sRegistro As String, objNFiscal As ClassNFiscal) As Long

Dim iPosicaoIni As Integer
Dim iPosicaoFim As Integer
Dim objItemNFGrade As New ClassItemRomaneioGrade

On Error GoTo Erro_ProcessaArqNFs_Obtem_ItemGradeItemNF

    iPosicaoIni = InStr(sRegistro, Chr(vbKeyEscape)) + 1
    iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1

    With objItemNFGrade
        .iControleEstoque = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .lNumIntDoc = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sProduto = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .dQuantidade = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sDescricao = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .sUMEstoque = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        .lNumIntItemPV = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEnd)) - 1
        .iAlmoxarifado = Mid(sRegistro, iPosicaoIni, iPosicaoFim)
        
    End With
    
    objNFiscal.colItensNF(objNFiscal.colItensNF.Count).colItensRomaneioGrade.Add objItemNFGrade
    
    ProcessaArqNFs_Obtem_ItemGradeItemNF = SUCESSO
    
    Exit Function
    
Erro_ProcessaArqNFs_Obtem_ItemGradeItemNF:

    ProcessaArqNFs_Obtem_ItemGradeItemNF = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150854)
    
    End Select
    
    Exit Function

End Function

Function ProcessaArqNFs_Obtem_LocalizacaoItemNFGrade(sRegistro As String, objNFiscal As ClassNFiscal) As Long

Dim iPosicaoIni As Integer
Dim iPosicaoFim As Integer
Dim objLocalizacaoItemGrade As New ClassReservaItem

On Error GoTo Erro_ProcessaArqNFs_Obtem_LocalizacaoItemNFGrade

    iPosicaoIni = InStr(sRegistro, Chr(vbKeyEscape)) + 1
    iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    
    objLocalizacaoItemGrade.dQuantidade = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objLocalizacaoItemGrade.dtDataValidade = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objLocalizacaoItemGrade.iAlmoxarifado = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objLocalizacaoItemGrade.iFilialEmpresa = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
    objLocalizacaoItemGrade.sAlmoxarifado = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEnd)) - 1
    objLocalizacaoItemGrade.sResponsavel = Mid(sRegistro, iPosicaoIni, iPosicaoFim)
    
    objNFiscal.colItensNF(objNFiscal.colItensNF.Count).colItensRomaneioGrade(objNFiscal.colItensNF(objNFiscal.colItensNF.Count).colItensRomaneioGrade.Count).colLocalizacao.Add objLocalizacaoItemGrade
    
    ProcessaArqNFs_Obtem_LocalizacaoItemNFGrade = SUCESSO
    
    Exit Function
    
Erro_ProcessaArqNFs_Obtem_LocalizacaoItemNFGrade:

    ProcessaArqNFs_Obtem_LocalizacaoItemNFGrade = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150855)
    
    End Select
    
    Exit Function

End Function

Function GeracaoArqNFs_Le_LancamentosCTBNF(objNFiscal As ClassNFiscal, colLancamentos As Collection, colLanPendente As Collection) As Long

Dim lErro As Long
Dim iTransacao As Integer
Dim lComando As Long
Dim tLancamento As typeLancamento
Dim objLancto As ClassLancamento_Detalhe

On Error GoTo Erro_GeracaoArqNFs_Le_LancamentosCTBNF

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 86521

    lErro = Comando_Executar(lComando, "SELECT TransacaoCTB.Codigo FROM NFiscal, TiposDocInfo, TransacaoCTB WHERE NFiscal.TipoNFiscal = TiposDocInfo.Codigo AND TransacaoCTB.Transacao = NomeTelaNFiscal AND TransacaoCTB.SubTipo = SubTipoContabil AND NumIntDoc = ? ", iTransacao, objNFiscal.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 86522

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then gError 86523
    
    tLancamento.sOrigem = String(STRING_ORIGEM, 0)
    tLancamento.sConta = String(STRING_CONTA, 0)
    tLancamento.sCcl = String(STRING_CCL, 0)
    tLancamento.sHistorico = String(STRING_HISTORICO, 0)
    tLancamento.sDocOrigem = String(STRING_DOCORIGEM, 0)
    tLancamento.sProduto = String(STRING_PRODUTO, 0)
    tLancamento.sModelo = String(STRING_PADRAOCONTAB_MODELO, 0)
 
    lErro = Comando_Executar(lComando, "SELECT FilialEmpresa, Origem, Exercicio, PeriodoLan, Doc, Seq, Lote, PeriodoLote, Data, Conta, Ccl, Historico, Valor, NumIntDoc, FilialCliForn, CliForn, Transacao, DocAglutinado, SeqAglutinado, Aglutinado, ContaSimples, SeqContraPartida, EscaninhoCusto, ValorLivroAuxiliar, ClienteFornecedor, DocOrigem, Produto, Quantidade, DataEstoque, ApropriaCRProd, Status, Modelo FROM Lancamentos WHERE FilialEmpresa = ? AND Transacao = ? AND NumIntDoc = ?", _
    tLancamento.iFilialEmpresa, tLancamento.sOrigem, tLancamento.iExercicio, tLancamento.iPeriodoLan, tLancamento.lDoc, tLancamento.iSeq, tLancamento.iLote, tLancamento.iPeriodoLote, tLancamento.dtData, tLancamento.sConta, tLancamento.sCcl, tLancamento.sHistorico, tLancamento.dValor, tLancamento.lNumIntDoc, tLancamento.iFilialCliForn, tLancamento.lCliForn, tLancamento.iTransacao, tLancamento.lDocAglutinado, tLancamento.iSeqAglutinado, tLancamento.iAglutina, tLancamento.lContaSimples, tLancamento.iSeqContraPartida, tLancamento.iEscaninho_Custo, tLancamento.dValorLivroAuxiliar, tLancamento.iClienteFornecedor, tLancamento.sDocOrigem, tLancamento.sProduto, tLancamento.dQuantidade, tLancamento.dtDataEstoque, tLancamento.iApropriaCRProd, tLancamento.iStatus, tLancamento.sModelo, objNFiscal.iFilialEmpresa, iTransacao, objNFiscal.lNumIntDoc)
    If lErro <> SUCESSO Then gError 86524

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 86524

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objLancto = New ClassLancamento_Detalhe
    
        objLancto.iFilialEmpresa = tLancamento.iFilialEmpresa
        objLancto.sOrigem = tLancamento.sOrigem
        objLancto.iExercicio = tLancamento.iExercicio
        objLancto.iPeriodoLan = tLancamento.iPeriodoLan
        objLancto.lDoc = tLancamento.lDoc
        objLancto.iSeq = tLancamento.iSeq
        objLancto.iLote = tLancamento.iLote
        objLancto.iPeriodoLote = tLancamento.iPeriodoLote
        objLancto.dtData = tLancamento.dtData
        objLancto.sConta = tLancamento.sConta
        objLancto.sCcl = tLancamento.sCcl
        objLancto.sHistorico = tLancamento.sHistorico
        objLancto.dValor = tLancamento.dValor
        objLancto.lNumIntDoc = tLancamento.lNumIntDoc
        objLancto.iFilialCliForn = tLancamento.iFilialCliForn
        objLancto.lCliForn = tLancamento.lCliForn
        objLancto.iTransacao = tLancamento.iTransacao
        objLancto.lDoc = tLancamento.lDoc
        If tLancamento.lDocAglutinado <> 0 Then objLancto.iAglutina = LANCAMENTO_AGLUTINADO
        objLancto.lContaSimples = tLancamento.lContaSimples
        objLancto.iSeqContraPartida = tLancamento.iSeqContraPartida
        objLancto.iEscaninho_Custo = tLancamento.iEscaninho_Custo
        objLancto.dValorLivroAuxiliar = tLancamento.dValorLivroAuxiliar
        objLancto.iClienteFornecedor = tLancamento.iClienteFornecedor
        objLancto.sDocOrigem = tLancamento.sDocOrigem
        objLancto.sProduto = tLancamento.sProduto
        objLancto.dQuantidade = tLancamento.dQuantidade
        objLancto.dtDataEstoque = tLancamento.dtDataEstoque
        objLancto.iApropriaCRProd = tLancamento.iApropriaCRProd
        objLancto.iStatus = tLancamento.iStatus
        objLancto.sModelo = tLancamento.sModelo
    
        colLancamentos.Add objLancto
    
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 86525
    
    Loop

    tLancamento.sOrigem = String(STRING_ORIGEM, 0)
    tLancamento.sConta = String(STRING_CONTA, 0)
    tLancamento.sCcl = String(STRING_CCL, 0)
    tLancamento.sHistorico = String(STRING_HISTORICO, 0)
    tLancamento.sDocOrigem = String(STRING_DOCORIGEM, 0)
    tLancamento.sProduto = String(STRING_PRODUTO, 0)
    tLancamento.sModelo = String(STRING_PADRAOCONTAB_MODELO, 0)
 
    lErro = Comando_Executar(lComando, "SELECT FilialEmpresa, Origem, Exercicio, PeriodoLan, Doc, Seq, Lote, PeriodoLote, Data, Conta, Ccl, Historico, Valor, NumIntDoc, FilialCliForn, CliForn, Transacao, ContaSimples, SeqContraPartida, EscaninhoCusto, ValorLivroAuxiliar, ClienteFornecedor, DocOrigem, Produto, Quantidade, DataEstoque, ApropriaCRProd, Status, Modelo FROM LanPendente WHERE FilialEmpresa = ? AND Transacao = ? AND NumIntDoc = ?", _
    tLancamento.iFilialEmpresa, tLancamento.sOrigem, tLancamento.iExercicio, tLancamento.iPeriodoLan, tLancamento.lDoc, tLancamento.iSeq, tLancamento.iLote, tLancamento.iPeriodoLote, tLancamento.dtData, tLancamento.sConta, tLancamento.sCcl, tLancamento.sHistorico, tLancamento.dValor, tLancamento.lNumIntDoc, tLancamento.iFilialCliForn, tLancamento.lCliForn, tLancamento.iTransacao, tLancamento.lContaSimples, tLancamento.iSeqContraPartida, tLancamento.iEscaninho_Custo, tLancamento.dValorLivroAuxiliar, tLancamento.iClienteFornecedor, tLancamento.sDocOrigem, tLancamento.sProduto, tLancamento.dQuantidade, tLancamento.dtDataEstoque, tLancamento.iApropriaCRProd, tLancamento.iStatus, tLancamento.sModelo, objNFiscal.iFilialEmpresa, iTransacao, objNFiscal.lNumIntDoc)
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 86524

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objLancto = New ClassLancamento_Detalhe
    
        objLancto.iFilialEmpresa = tLancamento.iFilialEmpresa
        objLancto.sOrigem = tLancamento.sOrigem
        objLancto.iExercicio = tLancamento.iExercicio
        objLancto.iPeriodoLan = tLancamento.iPeriodoLan
        objLancto.lDoc = tLancamento.lDoc
        objLancto.iSeq = tLancamento.iSeq
        objLancto.iLote = tLancamento.iLote
        objLancto.iPeriodoLote = tLancamento.iPeriodoLote
        objLancto.dtData = tLancamento.dtData
        objLancto.sConta = tLancamento.sConta
        objLancto.sCcl = tLancamento.sCcl
        objLancto.sHistorico = tLancamento.sHistorico
        objLancto.dValor = tLancamento.dValor
        objLancto.lNumIntDoc = tLancamento.lNumIntDoc
        objLancto.iFilialCliForn = tLancamento.iFilialCliForn
        objLancto.lCliForn = tLancamento.lCliForn
        objLancto.iTransacao = tLancamento.iTransacao
        objLancto.lDoc = tLancamento.lDoc
        If tLancamento.lDocAglutinado <> 0 Then objLancto.iAglutina = LANCAMENTO_AGLUTINADO
        objLancto.lContaSimples = tLancamento.lContaSimples
        objLancto.iSeqContraPartida = tLancamento.iSeqContraPartida
        objLancto.iEscaninho_Custo = tLancamento.iEscaninho_Custo
        objLancto.dValorLivroAuxiliar = tLancamento.dValorLivroAuxiliar
        objLancto.iClienteFornecedor = tLancamento.iClienteFornecedor
        objLancto.sDocOrigem = tLancamento.sDocOrigem
        objLancto.sProduto = tLancamento.sProduto
        objLancto.dQuantidade = tLancamento.dQuantidade
        objLancto.dtDataEstoque = tLancamento.dtDataEstoque
        objLancto.iApropriaCRProd = tLancamento.iApropriaCRProd
        objLancto.iStatus = tLancamento.iStatus
        objLancto.sModelo = tLancamento.sModelo
    
        colLanPendente.Add objLancto
    
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 86525
    
    Loop
    
    Call Comando_Fechar(lComando)

    GeracaoArqNFs_Le_LancamentosCTBNF = SUCESSO
    
    Exit Function
    
Erro_GeracaoArqNFs_Le_LancamentosCTBNF:

    GeracaoArqNFs_Le_LancamentosCTBNF = gErr
    
    Select Case gErr
        
        Case 86521
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 86522, 86523
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", gErr)
        
        Case 86524, 86525
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LANCAMENTOS3", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150856)
    
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function ProcessaArqNFs_Obtem_Lancamentos(sRegistro As String, colLancamento_Detalhe As Collection) As Long

Dim objLancto As New ClassLancamento_Detalhe
Dim iPosicaoIni As Integer
Dim iPosicaoFim As Integer

On Error GoTo Erro_ProcessaArqNFs_Obtem_Lancamentos

    iPosicaoIni = InStr(sRegistro, Chr(vbKeyEscape)) + 1
    iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1

    With objLancto
        objLancto.iFilialEmpresa = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        objLancto.sOrigem = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        objLancto.iExercicio = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        objLancto.iPeriodoLan = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        objLancto.lDoc = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        objLancto.iSeq = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        objLancto.iLote = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        objLancto.iPeriodoLote = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        objLancto.dtData = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        objLancto.sConta = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        objLancto.sCcl = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        objLancto.sHistorico = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        objLancto.dValor = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        objLancto.lNumIntDoc = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        objLancto.iFilialCliForn = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        objLancto.lCliForn = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        objLancto.iTransacao = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        objLancto.lDoc = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        objLancto.iAglutina = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        objLancto.lContaSimples = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        objLancto.iSeqContraPartida = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        objLancto.iEscaninho_Custo = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        objLancto.dValorLivroAuxiliar = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        objLancto.iClienteFornecedor = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        objLancto.sDocOrigem = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        objLancto.sProduto = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        objLancto.dQuantidade = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        objLancto.dtDataEstoque = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        objLancto.iApropriaCRProd = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEscape)) - 1
        objLancto.iStatus = Mid(sRegistro, iPosicaoIni, iPosicaoFim): iPosicaoIni = (iPosicaoIni + iPosicaoFim) + 1: iPosicaoFim = InStr(Mid(sRegistro, iPosicaoIni), Chr(vbKeyEnd)) - 1
        objLancto.sModelo = Mid(sRegistro, iPosicaoIni, iPosicaoFim)
        
    End With
    
    colLancamento_Detalhe.Add objLancto
    
    ProcessaArqNFs_Obtem_Lancamentos = SUCESSO
    
    Exit Function
    
Erro_ProcessaArqNFs_Obtem_Lancamentos:

    ProcessaArqNFs_Obtem_Lancamentos = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150857)
    
    End Select
    
    Exit Function

End Function

Function ProcessaArqNFs_Grava_NFiscal(objNFiscal As ClassNFiscal, colLancamentos As Collection, colLanPendente As Collection) As Long

Dim lErro As Long
Dim objContabil As ClassContabil
Dim lNum As Long
Dim objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Erro_ProcessaArqNFs_Grava_NFiscal
    
    objNFiscal.iTipoNFiscal = NFiscal_Converte_Tipo(objNFiscal)
    
    lNum = objNFiscal.lNumNotaFiscal
       
    objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
       
    lErro = CF("TipoDOcInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 86543
    If lErro <> SUCESSO Then gError 86544
    
    'Notas Internas
    If objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIE Or objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIS Then
    
        lErro = Prepara_Serie(objNFiscal)
        If lErro <> SUCESSO Then Error 86526

        objNFiscal.lNumNotaFiscal = 0
        
    End If
    
    Select Case objTipoDocInfo.sNomeTelaNFiscal
        'Tela NFiscal
        Case NOME_TELA_NFISCAL
            
            lErro = CF("NFiscal_Grava_EmTrans", objNFiscal, objContabil)
            If lErro <> SUCESSO Then gError 86527
        
        Case NOME_TELA_NFISCALDEV
            
            lErro = CF("NFiscalDevolucao_Grava_EmTrans", objNFiscal, objContabil)
            If lErro <> SUCESSO Then gError 86528
        
        Case NOME_TELA_NFISCALFATURA
                
            lErro = CF("NFiscalFatura_Grava_EmTrans", objNFiscal, objContabil)
            If lErro <> SUCESSO Then gError 86529
    
        Case NOME_TELA_NFISCALREM
            
            lErro = CF("NFiscalRemessa_Grava_EmTrans", objNFiscal, objContabil)
            If lErro <> SUCESSO Then gError 86530
    
        Case NOME_TELA_NFISCALENTRADA
            
            lErro = CF("NFiscalEntrada_Grava_EmTrans", objNFiscal, objContabil)
            If lErro <> SUCESSO Then gError 86531
    
        Case NOME_TELA_NFISCALENTDEV
        
            lErro = CF("NFiscalEntradaDevolucao_Grava_EmTrans", objNFiscal, objContabil)
            If lErro <> SUCESSO Then gError 86532
        
        Case NOME_TELA_NFISCALENTREM
        
            lErro = CF("NFiscalEntradaRemessa_Grava_EmTrans", objNFiscal, objContabil)
            If lErro <> SUCESSO Then gError 86533
        
        Case NOME_TELA_NFISCALFATENTRADA
        
            lErro = CF("NFiscalFatEntrada_Grava_EmTrans", objNFiscal, objContabil)
            If lErro <> SUCESSO Then gError 86534
        
    End Select
            
    lErro = NFiscal_Grava_LancamentosCTB(objNFiscal, colLancamentos, colLanPendente)
    If lErro <> SUCESSO Then gError 86535

    ProcessaArqNFs_Grava_NFiscal = SUCESSO

    Exit Function
    
Erro_ProcessaArqNFs_Grava_NFiscal:
    
    ProcessaArqNFs_Grava_NFiscal = gErr
    
    objNFiscal.lNumNotaFiscal = lNum
    
    Select Case gErr
    
        Case 86526 To 86535, 86543
        
        Case 86544
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPODOCINFO_NAO_CADASTRADO2", gErr, objTipoDocInfo.iCodigo)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150858)
    
    End Select
    
    Exit Function

End Function


Function NFiscal_Grava_LancamentosCTB(objNFiscal As ClassNFiscal, colLancamentos As Collection, colLanPendente As Collection) As Long

Dim lErro As Long
Dim objLancamento_Cabecalho As New ClassLancamento_Cabecalho
Dim objLancto As ClassLancamento_Detalhe
Dim colLancamentoDetalhe As New Collection
Dim lDoc As Long
Dim objLote As New ClassLote
Dim dDebito As Double
Dim dCredito As Double
Dim iNumLanctos As Integer

On Error GoTo Erro_NFiscal_Grava_LancamentosCTB

    For Each objLancto In colLancamentos
        colLancamentoDetalhe.Add objLancto
    Next

    For Each objLancto In colLanPendente
        colLancamentoDetalhe.Add objLancto
    Next
    
    If colLancamentoDetalhe.Count > 0 Then
        Set objLancto = colLancamentoDetalhe(1)
        
        objLancamento_Cabecalho.dtData = objLancto.dtData
        objLancamento_Cabecalho.iExercicio = objLancto.iExercicio
        objLancamento_Cabecalho.iFilialEmpresa = objLancto.iFilialEmpresa
        objLancamento_Cabecalho.iLote = objLancto.iLote
        objLancamento_Cabecalho.iPeriodoLan = objLancto.iPeriodoLan
        objLancamento_Cabecalho.iPeriodoLote = objLancto.iPeriodoLote
        objLancamento_Cabecalho.iTransacao = objLancto.iTransacao
        objLancamento_Cabecalho.sModelo = objLancto.sModelo
        objLancamento_Cabecalho.sOrigem = objLancto.sOrigem
        
        lErro = CF("Voucher_Automatico_Trans", objLancamento_Cabecalho.iFilialEmpresa, objLancamento_Cabecalho.iExercicio, objLancamento_Cabecalho.iPeriodoLan, objLancamento_Cabecalho.sOrigem, lDoc)
        If lErro <> SUCESSO Then gError 86536
        
        objLancamento_Cabecalho.lDoc = lDoc
        
        objLote.iExercicio = objLancamento_Cabecalho.iExercicio
        objLote.iFilialEmpresa = objLancamento_Cabecalho.iFilialEmpresa
        objLote.iPeriodo = objLancamento_Cabecalho.iPeriodoLote
        objLote.sOrigem = objLancamento_Cabecalho.sOrigem
        
        lErro = CF("Lote_Automatico1", objLote)
        If lErro <> SUCESSO Then gError 86537
        
        lErro = CF("LotePendente_Grava_Trans", objLote)
        If lErro <> SUCESSO Then Error 86538
        
        objLancamento_Cabecalho.iLote = objLote.iLote
        
        For Each objLancto In colLancamentoDetalhe
            objLancto.lNumIntDoc = objNFiscal.lNumIntDoc
            objLancto.lDoc = objLancamento_Cabecalho.lDoc
            objLancto.iLote = objLancamento_Cabecalho.iLote
            objLancto.iPeriodoLote = objLote.iPeriodo
            
        Next
        
        lErro = CF("Lancamento_Grava_BD", objLancamento_Cabecalho, colLancamentoDetalhe, dDebito, dCredito, iNumLanctos)
        If lErro <> SUCESSO Then gError 86539
    
    End If
    
    NFiscal_Grava_LancamentosCTB = SUCESSO
    
    Exit Function
    
Erro_NFiscal_Grava_LancamentosCTB:

    NFiscal_Grava_LancamentosCTB = gErr
    
    Select Case gErr
    
        Case 86536, 86537, 86538, 86539
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150859)
            
    End Select
            
    Exit Function
    
End Function

Function NFiscal_Converte_Tipo(ByVal objNFiscal As ClassNFiscal) As Integer
'Devolve o tipo de nfiscal analogo ao da nf passada mas que nao tenha vinculo com pedido de compras ou de vendas
'Bastará criar um select case das nfs vinculadas a pv e pc. No futuro posso colocar na tabela tiposdocinfo.

Dim iNovoTipo As Integer
    
    Select Case objNFiscal.iTipoNFiscal
    
        Case 156
            iNovoTipo = 149
        Case 103
            iNovoTipo = 75
        Case 102
            iNovoTipo = 74
        Case 101
            iNovoTipo = 73
        Case 100
            iNovoTipo = 72
        Case 81
            iNovoTipo = 67
        Case 80
            iNovoTipo = 66
        Case 92
            iNovoTipo = 63
        Case 93
            iNovoTipo = 62
        Case 94
            iNovoTipo = 61
        Case 97
            iNovoTipo = 60
        Case 95
            iNovoTipo = 59
        Case 96
            iNovoTipo = 58
        Case 85
            iNovoTipo = 57
        Case 84
            iNovoTipo = 56
        Case 99
            iNovoTipo = 52
        Case 98
            iNovoTipo = 48
        Case 153
            iNovoTipo = 47
        Case 155
            iNovoTipo = 46
        Case 154
            iNovoTipo = 43
        Case 150
            iNovoTipo = 36
        Case 151
            iNovoTipo = 34
        Case 152
            iNovoTipo = 33
        Case 83
            iNovoTipo = 27
        Case 86
            iNovoTipo = 26
        Case 88
            iNovoTipo = 25
        Case 87
            iNovoTipo = 19
        Case 82
            iNovoTipo = 11
        Case 91
            iNovoTipo = 10
        Case 90
            iNovoTipo = 9
        Case 89
            iNovoTipo = 3
        
        Case Else
            iNovoTipo = objNFiscal.iTipoNFiscal
    
    End Select
    
    NFiscal_Converte_Tipo = iNovoTipo
     
End Function

Function Prepara_Serie(objNFiscal As ClassNFiscal) As Long

Dim objSerie As New ClassSerie
Dim lErro As Long

On Error GoTo Erro_Prepara_Serie

    objSerie.sSerie = objNFiscal.sSerie
    
    lErro = CF("Serie_Le", objSerie)
    If lErro <> SUCESSO And lErro <> 22202 Then gError 86540
    If lErro <> AD_SQL_SUCESSO Then gError 86541
    
    objSerie.lProxNumNFiscal = objNFiscal.lNumNotaFiscal
    
    lErro = Serie_Grava_Trans_SemCriticar(objSerie)
    If lErro <> SUCESSO Then gError 86542
    
    Prepara_Serie = SUCESSO
    
    Exit Function

Erro_Prepara_Serie:

    Prepara_Serie = gErr
    
    Select Case gErr
    
        Case 86540, 86542
        
        Case 86541
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_CADASTRADA", gErr, objSerie.sSerie)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150860)
            
    End Select
    
    Exit Function
    
End Function


Private Function Serie_Grava_Trans_SemCriticar(ByVal objSerie As ClassSerie) As Long
'Nao executa críticas de existência de NF com número maior ou igual ao que está sendo alterado.
'Criada para permitir a gravação de NFs em bloco.(Transferência de NFs entre BDs).

Dim lErro As Long
Dim alComando(1) As Long
Dim lCodEndereco As Long
Dim iIndice As Integer
Dim tSerie As typeSerie
Dim iCodigo  As Integer
Dim objNFiscal As New ClassNFiscal
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Serie_Grava_Trans_SemCriticar

    'Abre os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 22203
    Next

    tSerie.sNomeTsk = String(STRING_NOME_TSK, 0)
    
    'Lê a Serie no BD
    lErro = Comando_ExecutarPos(alComando(0), "SELECT ProxNumNFiscal, TipoFormulario, NomeTsk FROM Serie WHERE FilialEmpresa = ? AND Serie = ?", 0, tSerie.lProxNumNFiscal, tSerie.iTipoFormulario, tSerie.sNomeTsk, giFilialEmpresa, objSerie.sSerie)
    If lErro <> AD_SQL_SUCESSO Then Error 22205

    'Tenta ler
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 22206

    'Serie não existe - CRIACAO
    If lErro = AD_SQL_SEM_DADOS Then

        'Preenche FilialEmpresa de objSerie
        objSerie.iFilialEmpresa = giFilialEmpresa

        'Insere a Serie na tabela Serie
        lErro = Comando_Executar(alComando(1), "INSERT INTO Serie (FilialEmpresa, Serie, ProxNumNFiscal, TipoFormulario, NomeTsk) VALUES (?,?,?,?,?)", objSerie.iFilialEmpresa, objSerie.sSerie, objSerie.lProxNumNFiscal, objSerie.iTipoFormulario, objSerie.sNomeTsk)
        If lErro <> AD_SQL_SUCESSO Then Error 22207

    Else 'Serie existe - ALTERACAO
        
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> AD_SQL_SUCESSO Then Error 60432
        
                
        'Atualiza ProxNumNFiscal na tabela de Serie
        lErro = Comando_ExecutarPos(alComando(1), "UPDATE Serie SET ProxNumNFiscal = ?, TipoFormulario = ?, NomeTsk = ?", alComando(0), objSerie.lProxNumNFiscal, objSerie.iTipoFormulario, objSerie.sNomeTsk)
        If lErro <> AD_SQL_SUCESSO Then Error 22208
        
    End If
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    Serie_Grava_Trans_SemCriticar = SUCESSO

    Exit Function

Erro_Serie_Grava_Trans_SemCriticar:

    Serie_Grava_Trans_SemCriticar = Err

    Select Case Err

        Case 22203
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 22205, 22206
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SERIE", Err)

        Case 22207
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SERIE", Err)

        Case 22208
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MODIFICACAO_SERIE", Err)
        
        Case 60432
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_SERIE", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150861)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function Cidade_Importa(ByVal sCidade As String) As Long
 
Dim lErro As Long
Dim iCodigo As Integer
Dim objCidades As New ClassCidades
Dim lComando As Long
 
On Error GoTo Erro_Cidade_Importa
 
    If Len(Trim(sCidade)) <> 0 Then
    
       lComando = Comando_Abrir()
       If lComando = 0 Then gError 129909
    
       lErro = Comando_Executar(lComando, "SELECT Codigo FROM Cidades WHERE Descricao = ? ", iCodigo, sCidade)
       If lErro <> SUCESSO Then gError 129910
    
       lErro = Comando_BuscarPrimeiro(lComando)
       If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 129911
    
       If lErro = AD_SQL_SEM_DADOS Then
    
          lErro = Comando_Executar(lComando, "SELECT MAX(Codigo) From Cidades ", iCodigo)
          If lErro <> SUCESSO Then gError 129910
       
          lErro = Comando_BuscarPrimeiro(lComando)
          If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 129911
       
          iCodigo = iCodigo + 1
       
          objCidades.iCodigo = iCodigo
          objCidades.sDescricao = sCidade
       
          lErro = CF("Cidade_Grava_EmTrans", objCidades)
          If lErro <> SUCESSO Then gError 129912
       
          lErro = Comando_Executar(lComando, "UPDATE FatConfig SET Conteudo = ? WHERE Codigo = ? ", iCodigo + 1, "NUM_PROX_CIDADECADASTRO")
          If lErro <> SUCESSO Then gError 129913
    
       End If
       
       Call Comando_Fechar(lComando)
    
    End If
    
    Cidade_Importa = SUCESSO
     
    Exit Function
    
Erro_Cidade_Importa:
 
    Cidade_Importa = gErr
     
    Select Case gErr
 
       Case 129909 To 129913
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150862)
     
    End Select
     
    Call Comando_Fechar(lComando)
    
    Exit Function
 
End Function

Function GeracaoArqNFs_Obtem_Produtos1(colProdutos As Collection, ByVal sCodigo As String) As Long

Dim lErro As Long
Dim objProduto As ClassProduto
Dim objProdutoFilial As ClassProdutoFilial
Dim objProdutoAux As ClassProduto
Dim sProdutoPai As String, bJaExiste As Boolean

On Error GoTo Erro_GeracaoArqNFs_Obtem_Produtos1

    bJaExiste = False
    For Each objProdutoAux In colProdutos
        If objProdutoAux.sCodigo = sCodigo Then
            bJaExiste = True
            Exit For
        End If
    Next
    
    If bJaExiste = False Then
    
        Set objProduto = New ClassProduto
        
        objProduto.sCodigo = sCodigo
        sProdutoPai = String(STRING_PRODUTO, 0)
        lErro = Mascara_RetornaProdutoPai(sCodigo, sProdutoPai)
        If lErro <> SUCESSO Then gError 129914
        
        If sProdutoPai <> "" Then
        
            lErro = GeracaoArqNFs_Obtem_Produtos1(colProdutos, sProdutoPai)
            If lErro <> SUCESSO Then gError 129915
        
        End If
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 129916
        If lErro <> AD_SQL_SUCESSO Then gError 129917
        
        lErro = CF("Produto_Le_Categorias", objProduto, objProduto.colCategoriaItem)
        If lErro <> SUCESSO Then gError 86449
        
        Set objProdutoFilial = New ClassProdutoFilial
        
        objProdutoFilial.sProduto = objProduto.sCodigo
        objProdutoFilial.iFilialEmpresa = giFilialEmpresa
        
        lErro = CF("ProdutoFilial_Le", objProdutoFilial)
        If lErro <> SUCESSO And lErro <> 28261 Then gError 86450
        
        objProduto.sICMSAliquota = objProdutoFilial.sICMSAliquota
        objProduto.sSituacaoTribECF = objProdutoFilial.sSituacaoTribECF
        
        colProdutos.Add objProduto
    
    End If
    
    GeracaoArqNFs_Obtem_Produtos1 = SUCESSO
    
    Exit Function
    
Erro_GeracaoArqNFs_Obtem_Produtos1:
    
    GeracaoArqNFs_Obtem_Produtos1 = gErr

    Select Case gErr

        Case 129914, 129915, 129916

        Case 129917
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150863)

    End Select

    Exit Function

End Function

Private Function GeracaoArqNFs_FilCli_TrataEnderecos(ByVal lEndereco As Long, ByVal lEnderecoCobranca As Long, ByVal lEnderecoEntrega As Long, ByVal colEnderecos As Collection) As Long

Dim lErro As Long
Dim objEndereco As ClassEndereco
Dim objEnderecoEntrega As ClassEndereco
Dim objEnderecoCobranca As ClassEndereco

On Error GoTo Erro_GeracaoArqNFs_FilCli_TrataEnderecos

    Set objEndereco = New ClassEndereco
    Set objEnderecoEntrega = New ClassEndereco
    Set objEnderecoCobranca = New ClassEndereco
    
    objEndereco.lCodigo = lEndereco
    objEnderecoCobranca.lCodigo = lEnderecoCobranca
    objEnderecoEntrega.lCodigo = lEnderecoEntrega
    
    lErro = CF("Endereco_Le", objEndereco)
    If lErro <> SUCESSO And lErro <> 12309 Then gError 86487
    If lErro <> SUCESSO Then gError 86488
    
    lErro = CF("Endereco_Le", objEnderecoCobranca)
    If lErro <> SUCESSO And lErro <> 12309 Then gError 86487
    If lErro <> SUCESSO Then gError 86488

    lErro = CF("Endereco_Le", objEnderecoEntrega)
    If lErro <> SUCESSO And lErro <> 12309 Then gError 86487
    If lErro <> SUCESSO Then gError 86488
    
    colEnderecos.Add objEndereco
    colEnderecos.Add objEnderecoCobranca
    colEnderecos.Add objEnderecoEntrega
            
    GeracaoArqNFs_FilCli_TrataEnderecos = SUCESSO
     
    Exit Function
    
Erro_GeracaoArqNFs_FilCli_TrataEnderecos:

    GeracaoArqNFs_FilCli_TrataEnderecos = gErr
     
    Select Case gErr
          
        Case 86487
        
        Case 86488
            Call Rotina_Erro(vbOKOnly, "ERRO_ENDERECO_NAO_CADASTRADO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150864)
     
    End Select
     
    Exit Function

End Function

Function NFiscal_ObtemCondPagto(ByVal objNFiscal As ClassNFiscal) As Long

Dim lErro As Long, objTituloRec As New ClassTituloReceber
Dim objTituloPag As New ClassTituloPagar

On Error GoTo Erro_NFiscal_ObtemCondPagto

    Select Case objNFiscal.iClasseDocCPR
    
        Case CPR_TITULO_PAGAR
                
            'carrega dados do titulo a pagar associado à nota fiscal para obter a condicao de pagto
            objTituloPag.lNumIntDoc = objNFiscal.lNumIntDocCPR
            
            lErro = CF("TituloPagar_Le", objTituloPag)
            If lErro <> SUCESSO And lErro <> 18372 Then gError 56662
            
            If lErro <> SUCESSO Then
            
                lErro = CF("TituloPagarBaixado_Le", objTituloPag)
                If lErro <> SUCESSO And lErro <> 56661 Then gError 56663
                
        '        If lErro <> SUCESSO Then gError 56664
            
            End If
    
            objNFiscal.iCondicaoPagto = objTituloPag.iCondicaoPagto
        
        Case CPR_TITULO_RECEBER
        
            'carrega dados do titulo a receber associado à nota fiscal para obter a condicao de pagto
            objTituloRec.lNumIntDoc = objNFiscal.lNumIntDocCPR
            lErro = CF("TituloReceber_Le", objTituloRec, 1)
            If lErro <> SUCESSO And lErro <> 26061 Then gError 56571
            If lErro <> SUCESSO Then
            
                lErro = CF("TituloReceberBaixado_Le", objTituloRec, 1)
                If lErro <> SUCESSO And lErro <> 56570 Then gError 56572
                If lErro <> SUCESSO Then gError 56573
            
            End If
            
            objNFiscal.iCondicaoPagto = objTituloRec.iCondicaoPagto
            
    End Select

    NFiscal_ObtemCondPagto = SUCESSO
     
    Exit Function
    
Erro_NFiscal_ObtemCondPagto:

    NFiscal_ObtemCondPagto = gErr
     
    Select Case gErr
          
        Case 56571, 56572, 56662, 56663
        
        Case 56664
            Call Rotina_Erro(vbOKOnly, "ERRO_TITULO_PAGAR_INEXISTENTE", gErr)
        
        Case 56573
            Call Rotina_Erro(vbOKOnly, "ERRO_TITULO_REC_INEXISTENTE", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150865)
     
    End Select
     
    Exit Function

End Function



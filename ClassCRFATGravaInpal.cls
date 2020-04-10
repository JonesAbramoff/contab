VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCRFATGrava"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'ALTERAÇÕES INPAL
'1 - Márcio 08/2000 - Incluindo os campos PadraoTaxaFin, PadraoJuros, TaxaFinanceira, Juros na gravação do Cliente para a INPAL
'2 - Márcio 08/2000 - Incluindo o tratamento de juros na inserção de Títulos à Receber para a INPAL
'3 - Márcio 08/2000 - Incluindo o tratamento de Valor Original e Motivo Diferença para as parcelas à pagar

Option Explicit

Function FilialFornecedor_Grava_Estatisticas(ByVal objNFiscal As ClassNFiscal, ByVal iCancelamento As Integer, Optional ByVal objFilFornFilEmp As ClassFilFornFilEmp) As Long
'Tem que ser chamado dentro de Transação
'Atualiza a gravação da parte de estatística do Fornecedor

Dim lErro As Long
Dim objFilFornFilEmp1 As New ClassFilFornFilEmp
Dim lComando As Long
Dim lComando2 As Long
Dim lComando3 As Long
Dim lMaiorAtraso As Long
Dim objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Erro_FilialFornecedor_Grava_Estatisticas

    If objFilFornFilEmp Is Nothing Then Set objFilFornFilEmp = New ClassFilFornFilEmp

    'Inicializa comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 52634

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then gError 52635

    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then gError 52636
    
    If Not objNFiscal Is Nothing Then
    
        objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
    
        lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
        If lErro <> SUCESSO And lErro <> 31415 Then gError 89129
    
        If lErro <> SUCESSO Then gError 89130
        
        'Trata-se de uma nota fiscal de venda
        If objTipoDocInfo.iCompras = TIPODOCINFO_COMPRAS Or objTipoDocInfo.iCompras = TIPODOCINFO_COMPRAS_DEVOLUCAO Then
    
            If objTipoDocInfo.iCompras = TIPODOCINFO_COMPRAS Then
    
                objFilFornFilEmp.lNumeroCompras = 1
                objFilFornFilEmp.dValorAcumuladoCompras = objNFiscal.dValorTotal
            
            Else
                
                objFilFornFilEmp.lNumTotalDevolucoes = 1
                objFilFornFilEmp.dValorAcumuladoDevolucoes = objNFiscal.dValorTotal
                
            End If
            
            objFilFornFilEmp1.lCodFornecedor = objNFiscal.lFornecedor
            objFilFornFilEmp1.iFilialEmpresa = objNFiscal.iFilialEmpresa
            objFilFornFilEmp1.iCodFilial = objNFiscal.iFilialForn
        
        Else
        
            objFilFornFilEmp1.lCodFornecedor = objFilFornFilEmp.lCodFornecedor
            objFilFornFilEmp1.iFilialEmpresa = objFilFornFilEmp.iFilialEmpresa
            objFilFornFilEmp1.iCodFilial = objFilFornFilEmp.iCodFilial
            
        End If

    Else
    
        objFilFornFilEmp1.lCodFornecedor = objFilFornFilEmp.lCodFornecedor
        objFilFornFilEmp1.iFilialEmpresa = objFilFornFilEmp.iFilialEmpresa
        objFilFornFilEmp1.iCodFilial = objFilFornFilEmp.iCodFilial

    End If

    If (Not objNFiscal Is Nothing And objTipoDocInfo.iOrigem = DOCINFO_ORIGEM_FORNECEDOR) Or objNFiscal Is Nothing Then
        
        lErro = CF("FilFornFilEmp_Le_Lock", lComando, objFilFornFilEmp1)
        If lErro <> SUCESSO And lErro <> 52642 Then gError 52637

        'se não encontrou
        If lErro = 52642 Then

            'se não é um cancelamento/exclusao de nota e não encontrou o registro na tabela FilFornFilEmp
            If iCancelamento <> CANCELAMENTO_FILIALFORNFILEMP Then

                'Insere na tabela de FilialFornFilEmp
                lErro = Comando_Executar(lComando3, "INSERT INTO FilialFornFilEmp (FilialEmpresa, CodFornecedor, CodFilial, NumeroCompras, ValorAcumuladoCompras, NumTotalDevolucoes, ValorAcumuladoDevolucoes, PedidosEmAberto, AtrasoAcumulado, PedidosRecebidos, ItensPedidosRecebidos) VALUES (?,?,?,?,?,?,?,?,?,?,?)", objFilFornFilEmp1.iFilialEmpresa, objFilFornFilEmp1.lCodFornecedor, objFilFornFilEmp1.iCodFilial, objFilFornFilEmp.lNumeroCompras, objFilFornFilEmp.dValorAcumuladoCompras, objFilFornFilEmp.lNumTotalDevolucoes, objFilFornFilEmp.dValorAcumuladoDevolucoes, objFilFornFilEmp.lPedidosEmAberto, objFilFornFilEmp.lAtrasoAcumulado, objFilFornFilEmp.lPedidosRecebidos, objFilFornFilEmp.lItensPedidosRecebidos)
                If lErro <> AD_SQL_SUCESSO Then gError 52638

            End If
            
        'se encontrou
        Else
            
            If iCancelamento = CANCELAMENTO_FILIALFORNFILEMP Then
            
                'Atualiza a tabela de FilialFornFilEmp subtraindo os valores
                lErro = Comando_ExecutarPos(lComando2, "UPDATE FilialFornFilEmp SET NumeroCompras=NumeroCompras - ?, ValorAcumuladoCompras=ValorAcumuladoCompras-?, NumTotalDevolucoes=NumTotalDevolucoes-?, ValorAcumuladoDevolucoes=ValorAcumuladoDevolucoes-?, PedidosEmAberto=PedidosEmAberto-?, AtrasoAcumulado=AtrasoAcumulado-?, PedidosRecebidos=PedidosRecebidos-?, ItensPedidosRecebidos=ItensPedidosRecebidos-?", lComando, objFilFornFilEmp.lNumeroCompras, objFilFornFilEmp.dValorAcumuladoCompras, objFilFornFilEmp.lNumTotalDevolucoes, objFilFornFilEmp.dValorAcumuladoDevolucoes, objFilFornFilEmp.lPedidosEmAberto, objFilFornFilEmp.lAtrasoAcumulado, objFilFornFilEmp.lPedidosRecebidos, objFilFornFilEmp.lItensPedidosRecebidos)
                If lErro <> AD_SQL_SUCESSO Then gError 89419
            
            Else
            
                'Atualiza a tabela de FilialFornFilEmp somando os valores
                lErro = Comando_ExecutarPos(lComando2, "UPDATE FilialFornFilEmp SET NumeroCompras=NumeroCompras + ?, ValorAcumuladoCompras=ValorAcumuladoCompras+?, NumTotalDevolucoes=NumTotalDevolucoes+?, ValorAcumuladoDevolucoes=ValorAcumuladoDevolucoes+?, PedidosEmAberto=PedidosEmAberto+?, AtrasoAcumulado=AtrasoAcumulado+?, PedidosRecebidos=PedidosRecebidos+?, ItensPedidosRecebidos=ItensPedidosRecebidos+?", lComando, objFilFornFilEmp.lNumeroCompras, objFilFornFilEmp.dValorAcumuladoCompras, objFilFornFilEmp.lNumTotalDevolucoes, objFilFornFilEmp.dValorAcumuladoDevolucoes, objFilFornFilEmp.lPedidosEmAberto, objFilFornFilEmp.lAtrasoAcumulado, objFilFornFilEmp.lPedidosRecebidos, objFilFornFilEmp.lItensPedidosRecebidos)
                If lErro <> AD_SQL_SUCESSO Then gError 52639

            End If

        End If
 
        If Not objNFiscal Is Nothing Then
        
            If objTipoDocInfo.iCompras = TIPODOCINFO_COMPRAS Then

                'Atualiza as tabelas SldMesForn e SldDiaForn
                lErro = CF("SldMesForn_Grava_Estatisticas", objNFiscal, iCancelamento)
                If lErro <> SUCESSO Then gError 89345
        
            ElseIf objTipoDocInfo.iCompras = TIPODOCINFO_COMPRAS_DEVOLUCAO Then

                If iCancelamento = CANCELAMENTO_FILIALFORNFILEMP Then

                    'Atualiza as tabelas SldMesForn e SldDiaForn
                    lErro = CF("SldMesForn_Grava_Estatisticas", objNFiscal, CADASTRAMENTO_FILIALFORNFILEMP)
                    If lErro <> SUCESSO Then gError 89793
                    
                Else

                    'Atualiza as tabelas SldMesForn e SldDiaForn
                    lErro = CF("SldMesForn_Grava_Estatisticas", objNFiscal, CANCELAMENTO_FILIALFORNFILEMP)
                    If lErro <> SUCESSO Then gError 89794

                End If

            End If

        End If

    End If

    'Libera comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    
    FilialFornecedor_Grava_Estatisticas = SUCESSO

    Exit Function

Erro_FilialFornecedor_Grava_Estatisticas:

    FilialFornecedor_Grava_Estatisticas = gErr

    Select Case gErr

        Case 52634, 52635, 52636
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 52637, 89129, 89345, 89793, 89794

        Case 52638
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERIR_FILIALFORNFILEMP", gErr)

        Case 52639, 89419
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_FILIALFORNFILEMP", gErr)

        Case 89130
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr)

        Case 89418
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNFILEMP_NAO_CADASTRADO", gErr, objFilFornFilEmp1.lCodFornecedor, objFilFornFilEmp1.iCodFilial, objFilFornFilEmp1.iFilialEmpresa)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    'Libera comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    
    Exit Function

End Function

Function Cliente_Grava_Customizado(ByVal objCliente As ClassCliente) As Long
'grava a customizacao de vendedor

Dim lErro As Long
Dim alComando(1 To 1) As Long
Dim iIndice As Integer
Dim objClienteInpal As ClassClienteInpal

On Error GoTo Erro_Cliente_Grava_Customizado

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 183137
    Next
    
    If Not (objCliente.objInfoUsu Is Nothing) Then
    
        Set objClienteInpal = objCliente.objInfoUsu
        
        'Altera registro na tabela Vendedores
        lErro = Comando_Executar(alComando(1), "UPDATE Clientes SET PadraoTaxaFin = ?, PadraoJuros = ?, TaxaFinanceira = ?, Juros = ? WHERE Codigo = ?", objClienteInpal.iPadraoTaxaFin, objClienteInpal.iPadraoJuros, objClienteInpal.dTaxaFinanceira, objClienteInpal.dJuros, objCliente.lCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 183138
        
    End If

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Cliente_Grava_Customizado = SUCESSO

    Exit Function

Erro_Cliente_Grava_Customizado:

    Cliente_Grava_Customizado = gErr

        Select Case gErr

        Case 183137
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 183138
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_CLIENTE", gErr, objCliente.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183139)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Function TituloReceber_Insere(ByVal objTituloReceber As ClassTituloReceber, ByVal colParcelaReceber As colParcelaReceber, ByVal colComissaoEmissao As ColComissao, ByVal colcolComissao As colcolComissao) As Long
'Insere o Título à Receber na Tabela de Titulos a Receber
    
Dim lErro As Long
Dim lNumIntDoc As Long
Dim lComando As Long
Dim lComando1 As Long
Dim objComissao As ClassComissao
Dim objParcelaReceber As ClassParcelaReceber
Dim lNumeroFatura As Long
Dim objCliente As New ClassCliente
Dim objClienteInpal As ClassClienteInpal

On Error GoTo Erro_TituloReceber_Insere

    'Inicializa comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 26393
        
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 26353
    
    'Se for Fatura a Receber
    If objTituloReceber.sSiglaDocumento = TIPODOC_FATURA_A_RECEBER Then
                
        'Gera o Número da Fatura
        lErro = CF("GeracaoFatura_Automatico_Trans", lNumeroFatura)
        If lErro <> SUCESSO Then gError 61464
        
        'Coloca no Número do Título
        objTituloReceber.lNumTitulo = lNumeroFatura
        
    End If
    
    'Verifica a existencia de um Título no BD com as mesmas características
    lErro = CF("TituloReceber_Testa_Repeticao", objTituloReceber)
    If lErro <> SUCESSO Then gError 26336
    
    'Gera um novo código interno para o Título
    lErro = CF("TituloReceber_Automatico", lNumIntDoc)
    If lErro <> SUCESSO Then gError 26347
    
    objTituloReceber.lNumIntDoc = lNumIntDoc
    
    'Verifica se Cliente e Filial estão cadastrados
    lErro = CF("FilialCliente_Lock", objTituloReceber.lCliente, objTituloReceber.iFilial)
    If lErro <> SUCESSO And lErro <> 43050 Then gError 26229
    
    'Não encontrou a Filial Cliente --> erro
    If lErro = 43050 Then gError 26230
    
    If objTituloReceber.iCondicaoPagto <> 0 Then
        'Verifica se Condição de Pagamento está cadastrada
        lErro = CF("CondicaoPagto_Lock", lComando, objTituloReceber.iCondicaoPagto)
        If lErro <> SUCESSO Then gError 26354
    End If
    
    '2 - Márcio 08/2000 - Incluindo o tratamento da juros na inserção de Títulos à Receber para a INPAL
                
    'Lê dados do cliente passado em objTituloReceber
    objCliente.lCodigo = objTituloReceber.lCliente
    lErro = CF("Cliente_ObterDadosBasicos", objCliente)
    If lErro <> SUCESSO And lErro <> 7616 Then gError 500004

    'Se não encontrou o cliente, erro
    If lErro = 7616 Then gError 500005

    lErro = CF("Cliente_Le_Customizado", objCliente)
    If lErro <> SUCESSO Then gError 500004
    
    Set objClienteInpal = objCliente.objInfoUsu
            
    'Se for Juros padrão
    If objClienteInpal.iPadraoJuros = VALOR_PADRAO Then

        'Guarda o juros padrão
        objTituloReceber.dPercJurosDiario = gobjCRFAT.dPercJurosDiario

    'Se o cliente possui Juros específico
    Else

        'Guarda o Juros do cliente
        objTituloReceber.dPercJurosDiario = objClienteInpal.dJuros / 30 '??? talvez deva arredondar. Jones 05/09/2000

    End If
    
    If objTituloReceber.iFilialEmpresa = 0 Then objTituloReceber.iFilialEmpresa = giFilialEmpresa
    
    'Insere o Título na tabela TitulosRec
    lErro = Comando_Executar(lComando1, "INSERT INTO TitulosRec (NumIntDoc, FilialEmpresa, Cliente, Filial, SiglaDocumento, NumTitulo, DataEmissao, Status, Saldo, NumParcelas, Valor, ValorIRRF, ValorISS, ISSRetido, ValorINSS, INSSRetido, PercJurosDiarios, PercMulta, Observacao, CondicaoPagto, DataRegistro, PISRetido, COFINSRetido, CSLLRetido, " & _
    "Moeda, Ccl, Natureza, CpoGenerico1, ReajustePeriodicidade, ReajusteBase, ReajustadoAte) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", lNumIntDoc, objTituloReceber.iFilialEmpresa, objTituloReceber.lCliente, objTituloReceber.iFilial, objTituloReceber.sSiglaDocumento, objTituloReceber.lNumTitulo, objTituloReceber.dtDataEmissao, STATUS_LANCADO, objTituloReceber.dSaldo, objTituloReceber.iNumParcelas, objTituloReceber.dValor, objTituloReceber.dValorIRRF, objTituloReceber.dValorISS, objTituloReceber.dISSRetido, objTituloReceber.dValorINSS, objTituloReceber.iINSSRetido, objTituloReceber.dPercJurosDiario, objTituloReceber.dPercMulta, objTituloReceber.sObservacao, objTituloReceber.iCondicaoPagto, _
        objTituloReceber.dtDataRegistro, objTituloReceber.dPISRetido, objTituloReceber.dCOFINSRetido, objTituloReceber.dCSLLRetido, objTituloReceber.iMoeda, objTituloReceber.sCcl, objTituloReceber.sNatureza, objTituloReceber.lCpoGenerico1, objTituloReceber.iReajustePeriodicidade, objTituloReceber.dtReajusteBase, objTituloReceber.dtReajustadoAte)
    If lErro <> AD_SQL_SUCESSO Then gError 26394

    For Each objComissao In colComissaoEmissao
        
        objComissao.iTipoTitulo = TIPO_TITULO_RECEBER
        objComissao.lNumIntDoc = lNumIntDoc
        objComissao.iStatus = STATUS_LIBERADO
        objComissao.dtDataGeracao = objTituloReceber.dtDataEmissao
        objComissao.dtDataBaixa = DATA_NULA

    Next

    'Grava as comissões do Título
    lErro = CF("Comissoes_Grava", colComissaoEmissao)
    If lErro <> SUCESSO Then gError 26348
    
    For Each objParcelaReceber In colParcelaReceber
        objParcelaReceber.lNumIntTitulo = lNumIntDoc
    Next

    'Grava as Parcelas a Receber
    lErro = CF("ParcelasReceber_Insere", colParcelaReceber, colcolComissao)
    If lErro <> SUCESSO Then gError 26349

    'Libera comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    TituloReceber_Insere = SUCESSO
    
    Exit Function
    
Erro_TituloReceber_Insere:

    TituloReceber_Insere = gErr
    
    Select Case gErr
    
        Case 26230
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_CADASTRADA", gErr, objTituloReceber.iFilial, objTituloReceber.lCliente)

        Case 26393, 26353
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 26229, 26336, 26347, 26348, 26349, 26354, 61464
        
        Case 26394
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TITULOS_REC", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
            
    End Select
    
    'Libera comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    Exit Function

End Function


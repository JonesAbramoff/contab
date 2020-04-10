VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFPrecoAuxCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function PrecoPraticado(ByVal objMnemonicoValor As ClassMnemonicoValor, ByVal objExeExp As ClassExeExp) As Long
'parametros: cliente , filial, produto
'retorna o preco do ultimo pedido de venda para o cliente , filial, produto
'se nao houver pedido retorna zero

Dim lErro As Long, iTabelaPreco As Integer
Dim lComando As Long, lCliente As Long, iFilial As Integer, sProduto As String
Dim iFilialFaturamento As Integer, dPreco As Double

On Error GoTo Erro_PrecoPraticado

    lCliente = StrParaLong(objMnemonicoValor.vParam(1))
    iFilial = StrParaInt(objMnemonicoValor.vParam(2))
    sProduto = objMnemonicoValor.vParam(3)
    iFilialFaturamento = objExeExp.objContexto.iFilialFaturamento

    If lCliente <> 0 And iFilial <> 0 And sProduto <> "" And iFilialFaturamento <> 0 Then

        lComando = Comando_Abrir()
        If lComando = 0 Then gError 106657

        'obtem tabela padrao do cliente
        lErro = Comando_Executar(lComando, "SELECT PrecoBase FROM a226 WHERE cliente = ? AND filial = ? AND produto = ? AND filialempresa = ?", dPreco, lCliente, iFilial, sProduto, iFilialFaturamento)
        If lErro <> AD_SQL_SUCESSO Then gError 106658

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106659

        Call Comando_Fechar(lComando)

    End If

    objMnemonicoValor.colValor.Add CDbl(dPreco)

    PrecoPraticado = SUCESSO

    Exit Function

Erro_PrecoPraticado:

    PrecoPraticado = gErr

    Select Case gErr

        Case 106657
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 106658, 106659
            Call Rotina_Erro(vbOKOnly, "ERRO_PRECO_PRATICADO", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150483)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ValPres(ByVal objMnemonicoValor As ClassMnemonicoValor, ByVal objExeExp As ClassExeExp) As Long
'parametros:Valor, Dias, TaxaMensal (juros compostos)
'retorna o valor presente (conceito financeiro)

Dim lErro As Long
Dim dValor As Double, dDias As Double, dTaxaMensal As Double

On Error GoTo Erro_ValPres

    dValor = StrParaDbl(objMnemonicoValor.vParam(1))
    dDias = StrParaDbl(objMnemonicoValor.vParam(2))
    dTaxaMensal = StrParaDbl(objMnemonicoValor.vParam(3))

    '??? talvez arredondar
    objMnemonicoValor.colValor.Add dValor / (1 + (dTaxaMensal / 100)) ^ (dDias / 30)

    ValPres = SUCESSO

    Exit Function

Erro_ValPres:

    ValPres = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150484)

    End Select

    Exit Function

End Function

Function ValFut(ByVal objMnemonicoValor As ClassMnemonicoValor, ByVal objExeExp As ClassExeExp) As Long
'parametros:Valor, Dias, TaxaMensal(juros compostos)
'retorna o valor futuro (conceito financeiro)

Dim lErro As Long

Dim dValor As Double, dDias As Double, dTaxaMensal As Double

On Error GoTo Erro_ValFut

    dValor = StrParaDbl(objMnemonicoValor.vParam(1))
    dDias = StrParaDbl(objMnemonicoValor.vParam(2))
    dTaxaMensal = StrParaDbl(objMnemonicoValor.vParam(3))

    '??? talvez arredondar
    objMnemonicoValor.colValor.Add dValor * (1 + (dTaxaMensal / 100)) ^ (dDias / 30)

    ValFut = SUCESSO

    Exit Function

Erro_ValFut:

    ValFut = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150485)

    End Select

    Exit Function

End Function

Function CustoMP(ByVal objMnemonicoValor As ClassMnemonicoValor, ByVal objExeExp As ClassExeExp) As Long
'parametros:Produto
'retorna o custo das materias primas necessarias para fabricar o produto,
'incluindo as materias primas para os subprodutos, recursivamente

Dim lErro As Long, sProduto As String
Dim colMP As New Collection, dCustoProd As Double
Dim dCustoMP As Double
Dim objCustoDirFabrPlanMP As ClassCustoDirFabrPlanMP
Dim lComando As Long

On Error GoTo Erro_CustoMP

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 106830
    
    sProduto = objExeExp.objContexto.sProduto
    If sProduto = "" Then sProduto = objMnemonicoValor.vParam(1)

    lErro = CF("CustoDireto_ObtemMP", objExeExp.objContexto.iFilialFaturamento, objExeExp.objContexto.iAno, sProduto, objExeExp.objContexto.dQuantidade, colMP, objExeExp.objContexto.iRotinaOrigem)
    If lErro <> SUCESSO Then gError 106833
    
    'para cada materia prima
    For Each objCustoDirFabrPlanMP In colMP
            
        'obter o seu custo em custembmp.
        lErro = CF("CustoDiretoProd_ObtemValor", objExeExp.objContexto.iFilialFaturamento, objCustoDirFabrPlanMP.sProdutoMP, objExeExp.objContexto.dTaxaDescPadrao, objExeExp.objContexto.dTaxaValFut, objExeExp.objContexto.dDiasValFut, dCustoMP, lComando)
        If lErro <> SUCESSO Then gError 106831
        
        'com a qtde e custo unitario, acumulo o custo do insumo no custo do produto.
        dCustoProd = dCustoProd + (dCustoMP * objCustoDirFabrPlanMP.dQtde)
    
    Next
    
    Call Comando_Fechar(lComando)
    
    objMnemonicoValor.colValor.Add dCustoProd
    
    CustoMP = SUCESSO

    Exit Function

Erro_CustoMP:

    CustoMP = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150486)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function CustoEmb(ByVal objMnemonicoValor As ClassMnemonicoValor, ByVal objExeExp As ClassExeExp) As Long
'parametros:Produto
'retorna o custo das embalagens e seus acessorios (embalagens de expedicao, rotulos, palletizacao,...) necessarios para fabricar o produto,
'incluindo as embalagens dos subprodutos, recursivamente

Dim lErro As Long, sProduto As String
Dim colEmb As New Collection, dCustoProd As Double
Dim dCustoEmb As Double
Dim objCustoDirFabrPlanEmb As ClassCustoDirFabrPlanEmb
Dim lComando As Long

On Error GoTo Erro_CustoEmb

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 106830
    
    sProduto = objExeExp.objContexto.sProduto
    If sProduto = "" Then sProduto = objMnemonicoValor.vParam(1)

    lErro = CF("CustoDireto_ObtemEmb", objExeExp.objContexto.iFilialFaturamento, objExeExp.objContexto.iAno, sProduto, objExeExp.objContexto.dQuantidade, colEmb, objExeExp.objContexto.iRotinaOrigem)
    If lErro <> SUCESSO Then gError 106833
    
    'para cada embalagem/acessorio
    For Each objCustoDirFabrPlanEmb In colEmb
            
        'obter o seu custo em custembmp.
        lErro = CF("CustoDiretoProd_ObtemValor", objExeExp.objContexto.iFilialFaturamento, objCustoDirFabrPlanEmb.sProdutoEmb, objExeExp.objContexto.dTaxaDescPadrao, objExeExp.objContexto.dTaxaValFut, objExeExp.objContexto.dDiasValFut, dCustoEmb, lComando)
        If lErro <> SUCESSO Then gError 106831
        
        'com a qtde e custo unitario, acumulo o custo do insumo no custo do produto.
        dCustoProd = dCustoProd + (dCustoEmb * objCustoDirFabrPlanEmb.dQtde)
    
    Next
    
    Call Comando_Fechar(lComando)
    
    objMnemonicoValor.colValor.Add dCustoProd
    
    CustoEmb = SUCESSO

    Exit Function

Erro_CustoEmb:

    CustoEmb = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150487)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function DVVTotal(ByVal objMnemonicoValor As ClassMnemonicoValor, ByVal objExeExp As ClassExeExp) As Long
'parametros:Coluna

Dim lErro As Long

On Error GoTo Erro_DVVTotal

    objMnemonicoValor.colValor.Add CDbl(0)
    
    DVVTotal = SUCESSO

    Exit Function

Erro_DVVTotal:

    DVVTotal = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150488)

    End Select

    Exit Function

End Function

Function ICMSAliqTabPreco(ByVal objMnemonicoValor As ClassMnemonicoValor, ByVal objExeExp As ClassExeExp) As Long
'parametros:TabelaPreco
'retorna a aliquota de ICMS correspondente à tabela de precos.
'Se a tabela de preços contiver uma aliquota interna entao o sistema buscará na tabela de estados qual a aliquota interna da filial de faturamento

Dim lErro As Long, lComando As Long, dAliquota As Double, iTabela As Integer
Dim objFilialEmpresa As New AdmFiliais, objEstado As New ClassEstado

On Error GoTo Erro_ICMSAliqTabPreco

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 106661

    iTabela = objExeExp.objContexto.iTabelaPreco
    If iTabela = 0 Then iTabela = StrParaInt(objMnemonicoValor.vParam(1))

    lErro = Comando_Executar(lComando, "SELECT AliquotaICMS FROM TabelasDePreco WHERE Codigo = ?", dAliquota, iTabela)
    If lErro <> AD_SQL_SUCESSO Then gError 106662

    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106663

    If lErro = AD_SQL_SUCESSO Then

        'se é uma aliquota interna
        If dAliquota > ICMS_MAIOR_ALIQ_INTERESTADUAL Then

            objFilialEmpresa.iCodFilial = objExeExp.objContexto.iFilialFaturamento
            lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
            If lErro <> SUCESSO Then gError 106664

            objEstado.sSigla = objFilialEmpresa.objEndereco.sSiglaEstado
            lErro = CF("Estado_Le", objEstado)
            If lErro <> SUCESSO Then gError 106665

            dAliquota = objEstado.dICMSAliquotaInterna

        End If

        'para ser tratada como %
        dAliquota = dAliquota * 100
        
    End If

    Call Comando_Fechar(lComando)

    objMnemonicoValor.colValor.Add dAliquota

    ICMSAliqTabPreco = SUCESSO

    Exit Function

Erro_ICMSAliqTabPreco:

    ICMSAliqTabPreco = gErr

    Select Case gErr

        Case 106661
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 106662, 106663
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELAPRECO", Err)

        Case 106664, 106665

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150489)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function TabelaPrecoPadrao(ByVal objMnemonicoValor As ClassMnemonicoValor, ByVal objExeExp As ClassExeExp) As Long
'parametros:Cliente, Filial

Dim lErro As Long

On Error GoTo Erro_TabelaPrecoPadrao

    objMnemonicoValor.colValor.Add CDbl(0)
    
    TabelaPrecoPadrao = SUCESSO

    Exit Function

Erro_TabelaPrecoPadrao:

    TabelaPrecoPadrao = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150490)

    End Select

    Exit Function

End Function

Function DevDuv(ByVal objMnemonicoValor As ClassMnemonicoValor, ByVal objExeExp As ClassExeExp) As Long
'parametros:Cliente, Filial, Produto
'retorna o percentual referente a 'devedores duvidosos' especifico da chave (cliente, filial, produto)
'se nao houver retorna zero.

Dim lErro As Long, lComando As Long, dPercDVV As Double
Dim lCliente As Long, iFilial As Integer, sProduto As String

On Error GoTo Erro_DevDuv

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 106666

    lCliente = StrParaLong(objMnemonicoValor.vParam(1))
    iFilial = StrParaInt(objMnemonicoValor.vParam(2))
    sProduto = objMnemonicoValor.vParam(3)

    lErro = Comando_Executar(lComando, "SELECT PercDVV FROM DVVClienteProd WHERE FilialEmpresa = ? AND CodCliente = ? AND CodFilial = ? AND Produto = ?", dPercDVV, objExeExp.objContexto.iFilialFaturamento, lCliente, iFilial, sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 106667

    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106668
    If lErro <> AD_SQL_SUCESSO And objExeExp.objContexto.iRotinaOrigem <> FORMACAO_PRECO_ANALISE_MARGCONTR_REL Then gError 106924

    Call objMnemonicoValor.colValor.Add(dPercDVV)

    Call Comando_Fechar(lComando)

    DevDuv = SUCESSO

    Exit Function

Erro_DevDuv:

    DevDuv = gErr

    Select Case gErr

        Case 106666
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 106667, 106668
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_DVVCLIENTEPROD", gErr)

        Case 106924
            Call Rotina_Erro(vbOKOnly, "ERRO_SEM_VALOR_DEVDUV_FILCLIPROD", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150491)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function CustoDireto(ByVal objMnemonicoValor As ClassMnemonicoValor, ByVal objExeExp As ClassExeExp) As Long
'parametros:Produto
'retorna o custo direto (por unidade de ...) associado ao produto pelo rateio dos custos diretos

Dim lErro As Long, lComando As Long, sProduto As String, dQtdePrevTotal As Double, lComando1 As Long
Dim adQuant(1 To KIT_NUM_FATORES) As Double, adCusto(0 To KIT_NUM_FATORES) As Double, adQuantProd(1 To KIT_NUM_FATORES) As Double
Dim dCustoDireto As Double, dQuantPrevista As Double, iIndice As Integer, dCustoOutros As Double

On Error GoTo Erro_CustoDireto

    sProduto = objExeExp.objContexto.sProduto
    If sProduto = "" Then sProduto = objMnemonicoValor.vParam(1)

    If Len(Trim(sProduto)) <> 0 Then

        lErro = CustoDirFabrProdInf_Obter(objExeExp.objContexto.iFilialFaturamento, objExeExp.objContexto.iAno, sProduto, dCustoDireto)
        If lErro <> SUCESSO And lErro <> 124288 Then gError 124289
        
        If lErro <> SUCESSO Then
        
            lComando = Comando_Abrir()
            If lComando = 0 Then gError 106669
    
            lErro = Comando_Executar(lComando, "SELECT QuantPrevista, CustoTotal, CustoDirFabr.QuantFator1, CustoFator1, CustoDirFabr.QuantFator2, CustoFator2, CustoDirFabr.QuantFator3, CustoFator3, CustoDirFabr.QuantFator4, CustoFator4, CustoDirFabr.QuantFator5, CustoFator5, CustoDirFabr.QuantFator6, CustoFator6, CustoDirFabrProd.QuantFator1, CustoDirFabrProd.QuantFator2, CustoDirFabrProd.QuantFator3, CustoDirFabrProd.QuantFator4, CustoDirFabrProd.QuantFator5, CustoDirFabrProd.QuantFator6 FROM CustoDirFabr, CustoDirFabrProd WHERE CustoDirFabr.FilialEmpresa = CustoDirFabrProd.FilialEmpresa AND CustoDirFabr.Ano = CustoDirFabrProd.Ano AND CustoDirFabr.FilialEmpresa = ? AND CustoDirFabr.Ano = ? AND CustoDirFabrProd.Produto = ?", _
                dQuantPrevista, adCusto(0), adQuant(1), adCusto(1), adQuant(2), adCusto(2), adQuant(3), adCusto(3), adQuant(4), adCusto(4), adQuant(5), adCusto(5), adQuant(6), adCusto(6), adQuantProd(1), adQuantProd(2), adQuantProd(3), adQuantProd(4), adQuantProd(5), adQuantProd(6), objExeExp.objContexto.iFilialFaturamento, objExeExp.objContexto.iAno, sProduto)
            If lErro <> AD_SQL_SUCESSO Then gError 106670
    
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106671
    
            If lErro = AD_SQL_SUCESSO And dQuantPrevista <> 0 Then
    
                dCustoOutros = adCusto(0)
                For iIndice = 1 To KIT_NUM_FATORES
                    If adQuant(iIndice) <> 0 Then
                        dCustoDireto = dCustoDireto + (adQuantProd(iIndice) * adCusto(iIndice) / adQuant(iIndice))
                        dCustoOutros = dCustoOutros - adCusto(iIndice)
                    End If
                Next
    
                If Abs(adCusto(0) - dCustoOutros) > DELTA_VALORMONETARIO Then
    
                    dCustoDireto = dCustoDireto + (dCustoDireto * dCustoOutros / (adCusto(0) - dCustoOutros))
                    dCustoDireto = dCustoDireto / dQuantPrevista
    
                Else
    
                    If objExeExp.objContexto.dQtdePrevTotal = -1 Then
                    
                        lComando1 = Comando_Abrir()
                        If lComando1 = 0 Then gError 106669
    
                        lErro = Comando_Executar(lComando1, "SELECT SUM(QuantPrevista) FROM CustoDirFabrProd WHERE FilialEmpresa = ? AND Ano = ?", _
                            dQtdePrevTotal, objExeExp.objContexto.iFilialFaturamento, objExeExp.objContexto.iAno)
                        If lErro <> AD_SQL_SUCESSO Then gError 106670
                
                        lErro = Comando_BuscarProximo(lComando1)
                        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106671
    
                        Call Comando_Fechar(lComando1)
                        
                        objExeExp.objContexto.dQtdePrevTotal = dQtdePrevTotal
                        
                    Else
                    
                        dQtdePrevTotal = objExeExp.objContexto.dQtdePrevTotal
                        
                    End If
    
                    dCustoDireto = dCustoOutros / dQtdePrevTotal
                    
                End If

            End If

            Call Comando_Fechar(lComando)

        End If

    End If

    objMnemonicoValor.colValor.Add dCustoDireto

    CustoDireto = SUCESSO

    Exit Function

Erro_CustoDireto:

    CustoDireto = gErr

    Select Case gErr

        Case 106669
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 106670, 106671
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CUSTODIRETO", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150492)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Function CustoFixo(ByVal objMnemonicoValor As ClassMnemonicoValor, ByVal objExeExp As ClassExeExp) As Long
'parametros:Produto
'retorna o custo fixo (por unidade de ...) associado ao produto.
'Pode ser o valor calculado pela rotina de rateio de custos ou o alterado manualmente pelo usuario.

Dim lErro As Long, lComando As Long, sProduto As String, dCustoFixo As Double
Dim iAutomatico As Integer, dCusto As Double, dCustoCalculado As Double

On Error GoTo Erro_CustoFixo

    sProduto = objExeExp.objContexto.sProduto
    If sProduto = "" Then sProduto = objMnemonicoValor.vParam(1)

    If Len(Trim(sProduto)) <> 0 Then

        lComando = Comando_Abrir()
        If lComando = 0 Then gError 106672

        lErro = Comando_Executar(lComando, "SELECT Automatico, Custo, CustoCalculado FROM CustoFixoProd WHERE FilialEmpresa = ? AND DataReferencia = ? AND Produto = ?", _
            iAutomatico, dCusto, dCustoCalculado, objExeExp.objContexto.iFilialFaturamento, objExeExp.objContexto.dtDataCustoFixo, sProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 106678

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106673

        If lErro = AD_SQL_SUCESSO Then

            dCustoFixo = IIf(iAutomatico = CUSTOFIXOPROD_AUTOMATICO, dCustoCalculado, dCusto)

        End If

        Call Comando_Fechar(lComando)

    End If

    objMnemonicoValor.colValor.Add dCustoFixo

    CustoFixo = SUCESSO

    Exit Function

Erro_CustoFixo:

    CustoFixo = gErr

    Select Case gErr

        Case 106672
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 106678, 106673
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CUSTOFIXOPROD", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150493)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function CalculaDias_CondPagamento(ByVal objCondPagto As ClassCondicaoPagto, dDias As Double) As Long
'calcula o numero medio de dias para uma condicao de pagto
'??? em ctpedidovendausu.cls fazer a chamada desta funcao por CF() e retirar a implementacao.

Dim iDiaParcelaAtual As Integer
Dim iTotalDias As Integer
Dim iNumeroParcelas As Integer
Dim iIntervalo As Integer

    'Se a condição de pagamento for à vista ou vendor
    If objCondPagto.iCodigo = COD_A_VISTA Or objCondPagto.iCodigo = 0 Or InStr(1, UCase(objCondPagto.sDescricao), "VENDOR") <> 0 Then

        dDias = 0

    'Se não for à vista
    Else

        'Se a condição de pagamento for Mensal
        If objCondPagto.iMensal = 1 Then
            iIntervalo = 30
        Else
            iIntervalo = objCondPagto.iIntervaloParcelas
        End If

        'Guarda o número de parcelas tirando a primeira
        iNumeroParcelas = objCondPagto.iNumeroParcelas - 1

        'Se possui mais de duas parcelas, calcula a média
        If iNumeroParcelas > 1 Then

            'Calcula total de dias da condição de pagamento
            Do While iNumeroParcelas >= 0

                'Calcula o número de dias que faltam para chegar a parcela em questão
                iDiaParcelaAtual = objCondPagto.iDiasParaPrimeiraParcela + (iIntervalo * iNumeroParcelas)

                'Acumula o número de dias de todas as parcelas
                iTotalDias = iTotalDias + iDiaParcelaAtual

                'Decrementa o número de parcelas
                iNumeroParcelas = iNumeroParcelas - 1

            Loop

            'Calcula a média ponderada de dias
            dDias = iTotalDias / objCondPagto.iNumeroParcelas

        'Se possui exatamente duas parcelas
        ElseIf iNumeroParcelas = 1 Then

            dDias = objCondPagto.iDiasParaPrimeiraParcela + (iIntervalo / 2)

        'Se só possui exatamente uma parcela
        Else

            dDias = objCondPagto.iDiasParaPrimeiraParcela

        End If

    End If

    CalculaDias_CondPagamento = SUCESSO

End Function

Function DiasCondPagto(ByVal objMnemonicoValor As ClassMnemonicoValor, ByVal objExeExp As ClassExeExp) As Long
'parametros:CondPagto
'retorna o prazo medio em dias correspondente à condicao de pagto

Dim lErro As Long, iCondPagto As Integer
Dim objCondicaoPagto As New ClassCondicaoPagto
Dim dDias As Double

On Error GoTo Erro_DiasCondPagto

    iCondPagto = StrParaInt(objMnemonicoValor.vParam(1))

    If iCondPagto <> 0 Then

        objCondicaoPagto.iCodigo = iCondPagto
        lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
        If lErro <> SUCESSO And lErro <> 19205 Then gError 106674

        'Se a condição de pagamento não está cadastrada, erro
        If lErro = 19205 Then gError 106675

        'Calcula a média ponderada de dias para a taxa financeira
        Call CF("CalculaDias_CondPagamento", objCondicaoPagto, dDias)

    End If

    objMnemonicoValor.colValor.Add dDias

    DiasCondPagto = SUCESSO

    Exit Function

Erro_DiasCondPagto:

    DiasCondPagto = gErr

    Select Case gErr

        Case 106674

        Case 106675 'nao cadastrada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150494)

    End Select

    Exit Function

End Function

Function CondPagtoCliente(ByVal objMnemonicoValor As ClassMnemonicoValor, ByVal objExeExp As ClassExeExp) As Long
'parametros:Cliente
'retorna a condicao de pagto default para o cliente

Dim lErro As Long, objCliente As New ClassCliente, lCliente As Long
Dim iCondPagto As Integer

On Error GoTo Erro_CondPagtoCliente

    lCliente = StrParaLong(objMnemonicoValor.vParam(1))

    If lCliente <> 0 Then

        objCliente.lCodigo = lCliente
        lErro = CF("Cliente_ObterDadosBasicos", objCliente)
        If lErro <> SUCESSO And lErro <> 7616 Then gError 106676

        iCondPagto = objCliente.iCondicaoPagto

    End If

    objMnemonicoValor.colValor.Add iCondPagto

    CondPagtoCliente = SUCESSO

    Exit Function

Erro_CondPagtoCliente:

    CondPagtoCliente = gErr

    Select Case gErr

        Case 106676

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150495)

    End Select

    Exit Function

End Function

Function PrecoUltPV(ByVal objMnemonicoValor As ClassMnemonicoValor, ByVal objExeExp As ClassExeExp) As Long
'parametros:Cliente, Filial, Produto
'retorna o último preco unitario utilizado para o produto num pedido de venda

Dim lErro As Long, lCliente As Long, iFilial As Integer, sProduto As String
Dim objItemPedido As New ClassItemPedido, dPrecoUnitario As Double

On Error GoTo Erro_PrecoUltPV

    lCliente = StrParaLong(objMnemonicoValor.vParam(1))
    iFilial = StrParaInt(objMnemonicoValor.vParam(2))
    sProduto = objMnemonicoValor.vParam(3)

    If lCliente <> 0 And iFilial <> 0 And sProduto <> "" And objExeExp.objContexto.iFilialFaturamento <> 0 Then

        objItemPedido.iFilialEmpresa = objExeExp.objContexto.iFilialFaturamento
        objItemPedido.sProduto = sProduto
        lErro = CF("ClienteFilial_Le_UltimoItemPedido", objItemPedido, lCliente, iFilial)
        If lErro <> SUCESSO And lErro <> 94412 Then gError 106677

        'Se encontrou o Item
        If lErro = SUCESSO Then
            lErro = CF("Produto_ConvPrecoUMAnalise", sProduto, objItemPedido.sUnidadeMed, objItemPedido.dPrecoUnitario, dPrecoUnitario)
            If lErro <> SUCESSO Then gError 130000
        End If

    End If

    objMnemonicoValor.colValor.Add dPrecoUnitario
    
    PrecoUltPV = SUCESSO

    Exit Function

Erro_PrecoUltPV:

    PrecoUltPV = gErr

    Select Case gErr

        Case 106677, 130000

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150496)

    End Select

    Exit Function

End Function

Function ComissoesPropCliProd(ByVal objMnemonicoValor As ClassMnemonicoValor, ByVal objExeExp As ClassExeExp) As Long
'parametros:Cliente, Filial, Produto

    ComissoesPropCliProd = ComissoesCliProd(objMnemonicoValor, objExeExp, True)

End Function

Function ComissoesTercCliProd(ByVal objMnemonicoValor As ClassMnemonicoValor, ByVal objExeExp As ClassExeExp) As Long
'parametros:Cliente, Filial, Produto

    ComissoesTercCliProd = ComissoesCliProd(objMnemonicoValor, objExeExp, False)
    
End Function

Function CustoDiretoProd_ObtemValor(ByVal iFilialEmpresa As Integer, ByVal sProduto As String, ByVal dTaxaDescPadrao As Double, ByVal dTaxaValFut As Double, ByVal dDiasValFut As Double, dCusto As Double, ByVal lComando As Long) As Long
'obtem o custo de um produto numa filialempresa baseado na tabela CustoEmbMP

Dim lErro As Long, tcustoemb As typeCustoEmbMP, dCustoAux As Double
Dim objCondicaoPagto As New ClassCondicaoPagto, dtDataUltFrete As Date, dCustoUltFrete As Double
Dim dDias As Double, dCustoICMS As Double, dCustoPIS As Double, dCustoCOFINS As Double, dCustoIPI As Double
Dim iCondPagto As Integer, dAliqICMS As Double, dFreteKG As Double, iCondPagtoUltCom As Integer
Dim objProduto As New ClassProduto, dCustoUltCom As Double, dtDataUltCom As Date, dAliquotaICMSUltCom As Double, sMoedaUltCom As String, dTaxaUltCom As Double, dFreteUltCom As Double
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_CustoDiretoProd_ObtemValor

    With tcustoemb
        lErro = Comando_Executar(lComando, "SELECT Custo, AliquotaICMS, CondicaoPagto, FretePorKG, AliquotaICMSInf, CondicaoPagtoInf, FretePorKGInf FROM CustoEmbMP WHERE FilialEmpresa = ? AND Produto = ?", _
            .dCusto, .dAliquotaICMS, .iCondicaoPagto, .dFretePorKg, .iAliquotaICMSInf, .iCondicaoPagtoInf, .iFretePorKGInf, iFilialEmpresa, sProduto)
    End With
    If lErro <> AD_SQL_SUCESSO Then gError 106801
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106802
    
    'se nao achou
    If lErro <> AD_SQL_SUCESSO Then
    
        dCusto = 0

    Else
    
        iCondPagto = tcustoemb.iCondicaoPagto
        dAliqICMS = tcustoemb.dAliquotaICMS
        dFreteKG = tcustoemb.dFretePorKg

        dCustoAux = tcustoemb.dCusto
        
        objProduto.sCodigo = sProduto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 106886
            
        'se nao for mercadoria importada
        If objProduto.iOrigemMercadoria <> 1 Then
        
            objFilialEmpresa.iCodFilial = iFilialEmpresa
            lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
            If lErro <> SUCESSO Then gError 130487
            
            'obter aliqicms e cond pagto da ultima compra
            If tcustoemb.iCondicaoPagtoInf = 0 Or tcustoemb.iAliquotaICMSInf = 0 Then
            
                lErro = CF("Produto_ObterDadosUltCompra", objProduto, iFilialEmpresa, dCustoUltCom, dtDataUltCom, dAliquotaICMSUltCom, sMoedaUltCom, dTaxaUltCom, dFreteUltCom, iCondPagtoUltCom)
                If lErro <> SUCESSO Then gError 106887
            
                'se achou alguma compra
                If dtDataUltCom <> DATA_NULA Then
                
                    If tcustoemb.iAliquotaICMSInf = 0 Then dAliqICMS = dAliquotaICMSUltCom
                    If tcustoemb.iCondicaoPagtoInf = 0 Then iCondPagto = iCondPagtoUltCom
                    
                End If
                
            End If
            
            'obter frete por kg da ultima entrada
            If tcustoemb.iFretePorKGInf = 0 Then
            
                lErro = CF("Produto_ObterFreteUltEnt", objProduto, iFilialEmpresa, dtDataUltFrete, dCustoUltFrete)
                If lErro <> SUCESSO Then gError 106888
                
                If dtDataUltFrete <> DATA_NULA Then dFreteKG = dCustoUltFrete
                
            End If
            
            If objFilialEmpresa.iICMSPorEstimativa = 0 Then
            
                'calcular o icms a valor presente
                dCustoICMS = (tcustoemb.dCusto * dAliqICMS) / ((1 + dTaxaDescPadrao / 100) ^ (20 / 30))
            
            End If
            
            'se está no simples
            If objFilialEmpresa.iSimplesFederal <> 0 Then
            
                'obter valor do IPI
                dCustoIPI = tcustoemb.dCusto * objProduto.dIPIAliquota
                
                'agregar o custo do IPI
                dCustoAux = dCustoAux + dCustoIPI
            
            Else
            
                'calcular o pis a valor presente
                If objFilialEmpresa.iPISNaoCumulativo = 1 Then dCustoPIS = (tcustoemb.dCusto * 0.0165) / ((1 + dTaxaDescPadrao / 100) ^ (25 / 30))
                
                'calcular o cofins a valor presente
                If objFilialEmpresa.iPISNaoCumulativo = 1 Then dCustoCOFINS = (tcustoemb.dCusto * 0.076) / ((1 + dTaxaDescPadrao / 100) ^ (25 / 30))
            
            End If
            
            If iCondPagto <> 0 Then
        
                'calcular a valor presente pela taxa gdTaxaDescPadrao com o prazo de pagto medio da condicao de pagto.
            
                objCondicaoPagto.iCodigo = iCondPagto
                lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
                If lErro <> SUCESSO And lErro <> 19205 Then gError 106804
        
                'Se a condição de pagamento não está cadastrada, erro
                If lErro = 19205 Then gError 106805
        
                'Calcula a média ponderada de dias para a taxa financeira
                Call CF("CalculaDias_CondPagamento", objCondicaoPagto, dDias)
        
                dCustoAux = dCustoAux / ((1 + dTaxaDescPadrao / 100) ^ (dDias / 30))
            
            End If
            
            'agregar o frete
            dCustoAux = dCustoAux + dFreteKG
            
            'aplicar fator referente à perda
            'já deve estar no calculo do custodirfabr
            
            'levar a valor futuro pela taxa gdTaxaValFut com numero de dias gdDiasValFut
            dCustoAux = dCustoAux * ((1 + dTaxaValFut / 100) ^ (dDiasValFut / 30)) - (dCustoICMS + dCustoPIS + dCustoCOFINS)
        
        End If
        
        dCusto = dCustoAux

    End If
    
    CustoDiretoProd_ObtemValor = SUCESSO
     
    Exit Function
    
Erro_CustoDiretoProd_ObtemValor:

    CustoDiretoProd_ObtemValor = gErr
     
    Select Case gErr
          
        Case 130487
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150497)
     
    End Select
     
    Exit Function

End Function

Function CustoDireto_ObtemMP(ByVal iFilialEmpresa As Integer, ByVal iAno As Integer, ByVal sProduto As String, ByVal dQuantidade As Double, ByVal colMP As Collection, ByVal iRotinaOrigem As Integer) As Long
'preenche a colecao de materias primas necessarias para produzir sProduto na Filialempresa

Dim lErro As Long, lComando As Long, dCusto As Double
Dim objCustoDirFabr As New ClassCustoDirFabr, objCustoDirFabrProd As ClassCustoDirFabrProd
Dim objCustoDirFabrPlan As ClassCustoDirFabrPlan
Dim objCustoDirFabrPlanMP As ClassCustoDirFabrPlanMP

On Error GoTo Erro_CustoDireto_ObtemMP

    objCustoDirFabr.iAno = iAno
    objCustoDirFabr.iFilialEmpresa = iFilialEmpresa
    objCustoDirFabr.iRotinaOrigem = iRotinaOrigem
    
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 124009
    
    'verifica se o produto está cadastrado como MP e portanto é transferido de outra filialempresa
    lErro = Comando_Executar(lComando, "SELECT Custo FROM CustoEmbMP WHERE FilialEmpresa = ? AND Produto = ?", dCusto, iFilialEmpresa, sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 124010
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 124011
            
    'se o produto nao é transferido
    If lErro = AD_SQL_SEM_DADOS Then
    
        lErro = CF("CustoDiretoFabricacao_Prod_Passo1", objCustoDirFabr, objCustoDirFabrProd, sProduto, dQuantidade)
        If lErro <> SUCESSO Then gError 106828
        
        For Each objCustoDirFabrPlan In objCustoDirFabrProd.colPlan
        
            For Each objCustoDirFabrPlanMP In objCustoDirFabrPlan.colMP
                       
                colMP.Add objCustoDirFabrPlanMP
        
            Next
            
        Next
    
    Else
    
        Set objCustoDirFabrPlanMP = New ClassCustoDirFabrPlanMP
        
        With objCustoDirFabrPlanMP
            .iFilialEmpresa = iFilialEmpresa
            .sProduto = sProduto
            .sProdutoMP = sProduto
            .iSeq = 1
            .iAno = iAno
            .dQtde = dQuantidade
            .dParticipacao = 1
        End With
        
        colMP.Add objCustoDirFabrPlanMP
        
    End If
    
    Call Comando_Fechar(lComando)
    
    CustoDireto_ObtemMP = SUCESSO
     
    Exit Function
    
Erro_CustoDireto_ObtemMP:

    CustoDireto_ObtemMP = gErr
     
    Select Case gErr
          
        Case 124009
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 124010, 124011
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CUSTOEMBMP", Err)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150498)
     
    End Select
     
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function CustoDirFabrProd_ObtemInfo(ByVal objCustoDirFabrProd As ClassCustoDirFabrProd, ByVal objCustoDirFabr As ClassCustoDirFabr, ByVal lComando As Long) As Long
'obtem a qtde prevista e o rateio de custo direto de fabricacao de um produto
'objCustoDirFabrProd terá que ter os campos produto, ano e filialempresa preenchidos

Dim lErro As Long, tcustodirprod As typeCustoDirFabrProd

On Error GoTo Erro_CustoDirFabrProd_ObtemInfo

    'procura registro em CustoDirFabrProd
    With tcustodirprod
        lErro = Comando_Executar(lComando, "SELECT QuantPrevista, QuantFator1, QuantFator2, QuantFator3, QuantFator4, QuantFator5, QuantFator6 FROM CustoDirFabrProd WHERE FilialEmpresa = ? AND Ano = ? AND Produto = ?", _
            .dQuantPrevista, .dQuantFator1, .dQuantFator2, .dQuantFator3, .dQuantFator4, .dQuantFator5, .dQuantFator6, objCustoDirFabrProd.iFilialEmpresa, objCustoDirFabrProd.iAno, objCustoDirFabrProd.sProduto)
    End With
    If lErro <> AD_SQL_SUCESSO Then gError 106791
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106792
    
    If lErro <> AD_SQL_SUCESSO Then
    
        With objCustoDirFabrProd
            .dQuantPrevista = 0
            .dQuantFator1 = 0
            .dQuantFator2 = 0
            .dQuantFator3 = 0
            .dQuantFator4 = 0
            .dQuantFator5 = 0
            .dQuantFator6 = 0
            .dCusto = 0
        End With
        
    Else
    
        'com os fatores de cada custo consigo obter o custo direto do produto
        With tcustodirprod
            objCustoDirFabrProd.dQuantPrevista = .dQuantPrevista
            objCustoDirFabrProd.dQuantFator1 = .dQuantFator1
            objCustoDirFabrProd.dQuantFator2 = .dQuantFator2
            objCustoDirFabrProd.dQuantFator3 = .dQuantFator3
            objCustoDirFabrProd.dQuantFator4 = .dQuantFator4
            objCustoDirFabrProd.dQuantFator5 = .dQuantFator5
            objCustoDirFabrProd.dQuantFator6 = .dQuantFator6
        End With
                
        'Calcula objCustoDirFabrProd.dCusto
        Call CustoDirFabrProd_Calcula(objCustoDirFabr, objCustoDirFabrProd)
    
    End If
    
    CustoDirFabrProd_ObtemInfo = SUCESSO
     
    Exit Function
    
Erro_CustoDirFabrProd_ObtemInfo:

    CustoDirFabrProd_ObtemInfo = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150499)
     
    End Select
     
    Exit Function

End Function

Function CustoDirFabrProd_Calcula(ByVal objCustoDirFabr As ClassCustoDirFabr, ByVal objCustoDirFabrProd As ClassCustoDirFabrProd) As Long
'Calcula objCustoDirFabrProd.dCusto

Dim lErro As Long, dCusto1 As Double, dCusto2 As Double, dCusto3 As Double, dCusto4 As Double, dCusto5 As Double, dCusto6 As Double, dCustoOutros As Double
Dim dCustoInformado As Double

On Error GoTo Erro_CustoDirFabrProd_Calcula
        
    lErro = CustoDirFabrProdInf_Obter(objCustoDirFabrProd.iFilialEmpresa, objCustoDirFabrProd.iAno, objCustoDirFabrProd.sProduto, dCustoInformado)
    If lErro <> SUCESSO And lErro <> 12288 Then gError 124289
    
    If lErro = SUCESSO Then
    
        objCustoDirFabrProd.dCusto = dCustoInformado
    
    Else
    
        If objCustoDirFabr.dQuantFator1 <> 0 Then
        
            dCusto1 = ((objCustoDirFabrProd.dQuantFator1 * objCustoDirFabr.dCustoFator1) / objCustoDirFabr.dQuantFator1)
        
        Else
        
            dCusto1 = 0
        
        End If
        
        If objCustoDirFabr.dQuantFator2 <> 0 Then
        
            dCusto2 = ((objCustoDirFabrProd.dQuantFator2 * objCustoDirFabr.dCustoFator2) / objCustoDirFabr.dQuantFator2)
        
        Else
        
            dCusto2 = 0
        
        End If
        
        If objCustoDirFabr.dQuantFator3 <> 0 Then
        
            dCusto3 = ((objCustoDirFabrProd.dQuantFator3 * objCustoDirFabr.dCustoFator3) / objCustoDirFabr.dQuantFator3)
        
        Else
        
            dCusto3 = 0
        
        End If
        
        If objCustoDirFabr.dQuantFator4 <> 0 Then
        
            dCusto4 = ((objCustoDirFabrProd.dQuantFator4 * objCustoDirFabr.dCustoFator4) / objCustoDirFabr.dQuantFator4)
        
        Else
        
            dCusto4 = 0
        
        End If
        
        If objCustoDirFabr.dQuantFator5 <> 0 Then
        
            dCusto5 = ((objCustoDirFabrProd.dQuantFator5 * objCustoDirFabr.dCustoFator5) / objCustoDirFabr.dQuantFator5)
        
        Else
        
            dCusto5 = 0
        
        End If
        
        If objCustoDirFabr.dQuantFator6 <> 0 Then
        
            dCusto6 = ((objCustoDirFabrProd.dQuantFator6 * objCustoDirFabr.dCustoFator6) / objCustoDirFabr.dQuantFator6)
        
        Else
        
            dCusto6 = 0
        
        End If
        
        dCustoOutros = (dCusto1 + dCusto2 + dCusto3 + dCusto4 + dCusto5 + dCusto6) / (objCustoDirFabr.dCustoTotal - objCustoDirFabr.dCustoOutros) * objCustoDirFabr.dCustoOutros
        
        objCustoDirFabrProd.dCusto = dCusto1 + dCusto2 + dCusto3 + dCusto4 + dCusto5 + dCusto6 + dCustoOutros
    
    End If
    
    CustoDirFabrProd_Calcula = SUCESSO
    
    Exit Function
     
Erro_CustoDirFabrProd_Calcula:

    CustoDirFabrProd_Calcula = gErr
    
    Select Case gErr
          
        Case 124289
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150500)
     
    End Select
     
    Exit Function

End Function

Function ContextoPlan_InicializaCusto(ByVal objContexto As ClassContextoPlan) As Long

Dim sMnemonico As String, sExpressao As String
Dim lErro As Long, lComando As Long

On Error GoTo Erro_ContextoPlan_InicializaCusto

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 106832
    
    'obter valores globais para calculo de valor presente e futuro de insumo
    sMnemonico = String(STRING_MNEMONICOFPRECO_MNEMONICO, 0)
    sExpressao = String(STRING_MNEMONICOFPRECO_EXPRESSAO, 0)
    lErro = Comando_Executar(lComando, "SELECT Mnemonico, Expressao FROM MnemonicoFPTipo WHERE TipoPlanilha = ? AND FilialEmpresa = ? AND Escopo = ? AND Mnemonico IN ('TaxaDescPadrao', 'TaxaValFut', 'DiasValFut' )", _
        sMnemonico, sExpressao, PLANILHA_TIPO_TODOS, objContexto.iFilialFaturamento, MNEMONICOFPRECO_ESCOPO_GERAL)
    If lErro <> AD_SQL_SUCESSO Then gError 106778
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106779
        
    Do While lErro = AD_SQL_SUCESSO
        
        Select Case sMnemonico
        
            Case "TaxaDescPadrao"
                objContexto.dTaxaDescPadrao = StrParaDbl(sExpressao)
                
            Case "TaxaValFut"
                objContexto.dTaxaValFut = StrParaDbl(sExpressao)
            
            Case "DiasValFut"
                objContexto.dDiasValFut = StrParaDbl(sExpressao)
        
        End Select
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106780
    
    Loop

    Call Comando_Fechar(lComando)
    
    ContextoPlan_InicializaCusto = SUCESSO
     
    Exit Function
    
Erro_ContextoPlan_InicializaCusto:

    ContextoPlan_InicializaCusto = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150501)
     
    End Select
     
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function CustoDireto_ObtemEmb(ByVal iFilialEmpresa As Integer, ByVal iAno As Integer, ByVal sProduto As String, ByVal dQuantidade As Double, ByVal colEmb As Collection, ByVal iRotinaOrigem As Integer) As Long
'preenche a colecao de embalagens  e seus acessorios (embalagens de expedicao, rotulos, palletizacao,...) necessarios para produzir sProduto na Filialempresa

Dim lErro As Long, lComando As Long, dCusto As Double
Dim objCustoDirFabr As New ClassCustoDirFabr, objCustoDirFabrProd As ClassCustoDirFabrProd
Dim objCustoDirFabrPlan As ClassCustoDirFabrPlan
Dim objCustoDirFabrPlanEmb As ClassCustoDirFabrPlanEmb

On Error GoTo Erro_CustoDireto_ObtemEmb

    objCustoDirFabr.iAno = iAno
    objCustoDirFabr.iFilialEmpresa = iFilialEmpresa
    objCustoDirFabr.iRotinaOrigem = iRotinaOrigem
    
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 124006
    
    'verifica se o produto está cadastrado como MP e portanto é transferido de outra filialempresa
    lErro = Comando_Executar(lComando, "SELECT Custo FROM CustoEmbMP WHERE FilialEmpresa = ? AND Produto = ?", dCusto, iFilialEmpresa, sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 124007
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 124008
            
    'se o produto nao é transferido
    If lErro = AD_SQL_SEM_DADOS Then
    
        lErro = CF("CustoDiretoFabricacao_Prod_Passo1", objCustoDirFabr, objCustoDirFabrProd, sProduto, dQuantidade)
        If lErro <> SUCESSO Then gError 106828
        
        For Each objCustoDirFabrPlan In objCustoDirFabrProd.colPlan
        
            For Each objCustoDirFabrPlanEmb In objCustoDirFabrPlan.colEmb
            
                colEmb.Add objCustoDirFabrPlanEmb
        
            Next
            
        Next
    
    End If
    
    Call Comando_Fechar(lComando)
    
    CustoDireto_ObtemEmb = SUCESSO
     
    Exit Function
    
Erro_CustoDireto_ObtemEmb:

    CustoDireto_ObtemEmb = gErr
     
    Select Case gErr
          
        Case 124006
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 124007, 124008
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CUSTOEMBMP", Err)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150502)
     
    End Select
     
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function CustoDireto_ObtemEmbMP(ByVal iFilialEmpresa As Integer, ByVal iAno As Integer, ByVal sProduto As String, ByVal dQuantidade As Double, ByVal colEmb As Collection, ByVal colMP As Collection, ByVal iRotinaOrigem As Integer) As Long
'preenche as colecoes de materia-prima e de embalagens e seus acessorios (embalagens de expedicao, rotulos, palletizacao,...) necessarios para produzir sProduto na Filialempresa

Dim lErro As Long
Dim objCustoDirFabr As New ClassCustoDirFabr, objCustoDirFabrProd As ClassCustoDirFabrProd
Dim objCustoDirFabrPlan As ClassCustoDirFabrPlan
Dim objCustoDirFabrPlanMP As ClassCustoDirFabrPlanMP
Dim objCustoDirFabrPlanEmb As ClassCustoDirFabrPlanEmb

On Error GoTo Erro_CustoDireto_ObtemEmbMP

    objCustoDirFabr.iAno = iAno
    objCustoDirFabr.iFilialEmpresa = iFilialEmpresa
    objCustoDirFabr.iRotinaOrigem = iRotinaOrigem
    
    lErro = CF("CustoDiretoFabricacao_Prod_Passo1", objCustoDirFabr, objCustoDirFabrProd, sProduto, dQuantidade)
    If lErro <> SUCESSO Then gError 106828
    
    For Each objCustoDirFabrPlan In objCustoDirFabrProd.colPlan
    
        For Each objCustoDirFabrPlanMP In objCustoDirFabrPlan.colMP
        
            colMP.Add objCustoDirFabrPlanMP
    
        Next
        
    Next
    
    For Each objCustoDirFabrPlan In objCustoDirFabrProd.colPlan
    
        For Each objCustoDirFabrPlanEmb In objCustoDirFabrPlan.colEmb
        
            colEmb.Add objCustoDirFabrPlanEmb
    
        Next
        
    Next
    
    CustoDireto_ObtemEmbMP = SUCESSO
     
    Exit Function
    
Erro_CustoDireto_ObtemEmbMP:

    CustoDireto_ObtemEmbMP = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150503)
     
    End Select
     
    Exit Function

End Function

Function CustoDiretoFabricacao_Prod_Passo1(ByVal objCustoDirFabr As ClassCustoDirFabr, objCustoDirFabrProd As ClassCustoDirFabrProd, ByVal sProduto As String, ByVal dQtde As Double) As Long
'Processar sProduto e seus componentes

Dim lErro As Long, objCustoDirFabrPlan As New ClassCustoDirFabrPlan, objCustoDirFabrPlanMP As ClassCustoDirFabrPlanMP
Dim iIndice As Integer, alComando(1 To 9) As Long, objCustoDirFabrPlanEmb As ClassCustoDirFabrPlanEmb
Dim iProdNaFilial As Integer, iCompras As Integer, objCustoDirFabrPlanMPAux As ClassCustoDirFabrPlanMP
Dim bProdNaFilial As Boolean, iNatureza As Integer, objCustoDirFabrPlanEmbAux As ClassCustoDirFabrPlanEmb

On Error GoTo Erro_CustoDiretoFabricacao_Prod_Passo1

    'Abertura dos comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir
        If alComando(iIndice) = 0 Then gError 106591
    Next
    
    If dQtde <> 0 Then
    
        Set objCustoDirFabrProd = New ClassCustoDirFabrProd
        objCustoDirFabrProd.iFilialEmpresa = objCustoDirFabr.iFilialEmpresa
        objCustoDirFabrProd.iAno = objCustoDirFabr.iAno
        objCustoDirFabrProd.dtData = gdtDataHoje
        objCustoDirFabrProd.dQuantPrevista = dQtde
        objCustoDirFabrProd.sProduto = sProduto
        
        'inicializar objCustoDirFabrPlan com dados do produto colocando-o como filho de si mesmo com participacao de 100%
        objCustoDirFabrPlan.iFilialEmpresa = objCustoDirFabrProd.iFilialEmpresa
        objCustoDirFabrPlan.iAno = objCustoDirFabrProd.iAno
        objCustoDirFabrPlan.dQuantPai = dQtde
        objCustoDirFabrPlan.dParticipacao = 1 '100%
        objCustoDirFabrPlan.sProduto = objCustoDirFabrProd.sProduto
        objCustoDirFabrPlan.sProdutoPai = objCustoDirFabrProd.sProduto
        objCustoDirFabrPlan.sProdutoFilho = objCustoDirFabrProd.sProduto
        
        lErro = Comando_Executar(alComando(9), "SELECT Natureza, Compras, ProdNaFilial FROM Produtos, ProdutosFilial WHERE ProdutosFilial.Produto = ? AND Produtos.Codigo = ProdutosFilial.Produto AND ProdutosFilial.FilialEmpresa = ?", _
            iNatureza, iCompras, iProdNaFilial, sProduto, objCustoDirFabr.iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 130342
        
        lErro = Comando_BuscarProximo(alComando(9))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130343
        If lErro <> AD_SQL_SUCESSO Then gError 130344
        
        'se o produto nao é produzido na filial
        If iCompras <> PRODUTO_PRODUZIVEL Or iProdNaFilial = 0 Then
        
            bProdNaFilial = False
            
            objCustoDirFabrPlan.iSeq = objCustoDirFabrProd.colPlan.Count + 1
            
            'incluir objCustoDirFabrPlan na colecao de objCustoDirFabrProd
            objCustoDirFabrProd.colPlan.Add objCustoDirFabrPlan
        
            If iNatureza <> NATUREZA_PROD_EMBALAGENS Then
            
                Set objCustoDirFabrPlanMPAux = New ClassCustoDirFabrPlanMP
                
                With objCustoDirFabrPlanMPAux
                    .iFilialEmpresa = objCustoDirFabrPlan.iFilialEmpresa
                    .iAno = objCustoDirFabrPlan.iAno
                    .sProduto = sProduto
                    .iSeq = objCustoDirFabrPlan.iSeq
                    .sProdutoMP = sProduto
                    .dQtde = dQtde
                End With
                
                objCustoDirFabrPlan.colMP.Add objCustoDirFabrPlanMPAux
        
            Else
            
                Set objCustoDirFabrPlanEmbAux = New ClassCustoDirFabrPlanEmb
                
                With objCustoDirFabrPlanEmbAux
                    
                    .iFilialEmpresa = objCustoDirFabrPlan.iFilialEmpresa
                    .iAno = objCustoDirFabrPlan.iAno
                    .sProduto = sProduto
                    .iSeq = objCustoDirFabrPlan.iSeq
                    .sProdutoEmb = sProduto
                    .dQtde = dQtde
                    
                End With
                
                objCustoDirFabrPlan.colEmb.Add objCustoDirFabrPlanEmbAux
            
            End If
            
        Else
        
            bProdNaFilial = True
            
            'incluir uma linha referente ao proprio produto e mandar obter dados de seus componentes fabricados recursivamente
            lErro = CustoDiretoFabricacao_Prod_Passo1A(objCustoDirFabr, objCustoDirFabrProd, objCustoDirFabrPlan, alComando, 0)
            If lErro <> SUCESSO Then gError 106613
        
        End If
        
        Select Case objCustoDirFabr.iRotinaOrigem
        
            Case FORMACAO_PRECO_ROTINA_CUSTOSDIRETOS
            
                If bProdNaFilial Then
                
                    'gravar registro em CustoDirFabrProd
                    With objCustoDirFabrProd
                        lErro = Comando_Executar(alComando(1), "INSERT INTO CustoDirFabrProd (FilialEmpresa, Ano, Produto, Data, QuantPrevista, QuantFator1, QuantFator2, QuantFator3, QuantFator4, QuantFator5, QuantFator6) VALUES (?,?,?,?,?,?,?,?,?,?,?)", .iFilialEmpresa, .iAno, .sProduto, .dtData, .dQuantPrevista, .dQuantFator1, .dQuantFator2, .dQuantFator3, .dQuantFator4, .dQuantFator5, .dQuantFator6)
                    End With
                    If lErro <> SUCESSO Then gError 106614
                    
                    For Each objCustoDirFabrPlan In objCustoDirFabrProd.colPlan
            
                        'gravar registros em CustoDirFabrPlan
                        With objCustoDirFabrPlan
                            lErro = Comando_Executar(alComando(2), "INSERT INTO CustoDirFabrPlan (FilialEmpresa, Ano, Produto, Seq,  ProdutoPai, ProdutoFilho, QuantPai, Participacao, PesoFator1, PesoFator2, PesoFator3, PesoFator4, PesoFator5, PesoFator6) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)", .iFilialEmpresa, .iAno, .sProduto, .iSeq, .sProdutoPai, .sProdutoFilho, .dQuantPai, .dParticipacao, .dPesoFator1, .dPesoFator2, .dPesoFator3, .dPesoFator4, .dPesoFator5, .dPesoFator6)
                        End With
                        If lErro <> SUCESSO Then gError 106615
                
                        For Each objCustoDirFabrPlanMP In objCustoDirFabrPlan.colMP
                            With objCustoDirFabrPlanMP
                                lErro = Comando_Executar(alComando(3), "INSERT INTO CustoDirFabrPlanMP (FilialEmpresa, Ano, Produto, Seq, ProdutoMP, Qtde) VALUES (?,?,?,?,?,?)", .iFilialEmpresa, .iAno, .sProduto, .iSeq, .sProdutoMP, .dQtde)
                            End With
                            If lErro <> SUCESSO Then gError 106617
                        Next
                        
                        For Each objCustoDirFabrPlanEmb In objCustoDirFabrPlan.colEmb
                            With objCustoDirFabrPlanEmb
                                lErro = Comando_Executar(alComando(8), "INSERT INTO CustoDirFabrPlanEmb (FilialEmpresa, Ano, Produto, Seq, ProdutoEmb, Qtde) VALUES (?,?,?,?,?,?)", .iFilialEmpresa, .iAno, .sProduto, .iSeq, .sProdutoEmb, .dQtde)
                            End With
                            If lErro <> SUCESSO Then gError 106617
                        Next
                        
                    Next
                    
                    'se nao estiver reprocessando apenas 1 produto entao atualizar acumuladores de quantidades multiplicadas pelos fatores em objCustoDirFabr
                    If objCustoDirFabr.sProduto = "" Then
                        objCustoDirFabr.dQuantFator1 = objCustoDirFabr.dQuantFator1 + objCustoDirFabrProd.dQuantFator1
                        objCustoDirFabr.dQuantFator2 = objCustoDirFabr.dQuantFator2 + objCustoDirFabrProd.dQuantFator2
                        objCustoDirFabr.dQuantFator3 = objCustoDirFabr.dQuantFator3 + objCustoDirFabrProd.dQuantFator3
                        objCustoDirFabr.dQuantFator4 = objCustoDirFabr.dQuantFator4 + objCustoDirFabrProd.dQuantFator4
                        objCustoDirFabr.dQuantFator5 = objCustoDirFabr.dQuantFator5 + objCustoDirFabrProd.dQuantFator5
                        objCustoDirFabr.dQuantFator6 = objCustoDirFabr.dQuantFator6 + objCustoDirFabrProd.dQuantFator6
                    End If

                End If
                
            Case FORMACAO_PRECO_ROTINA_CUSTOFIXO
            Case FORMACAO_PRECO_ROTINA_CALCPRECO
            Case FORMACAO_PRECO_ANALISE_MARGCONTR, FORMACAO_PRECO_ANALISE_MARGCONTR_REL
            
        End Select
        
    End If
    
    'Fechamento comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    CustoDiretoFabricacao_Prod_Passo1 = SUCESSO
     
    Exit Function
    
Erro_CustoDiretoFabricacao_Prod_Passo1:

    CustoDiretoFabricacao_Prod_Passo1 = gErr
     
    Select Case gErr
          
        Case 106613 To 106615, 106617
            '??? colocar erros adequados
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 106591
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 130342, 130343
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_INFO_PROD", gErr)
        
        Case 130344
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_INFO_PROD_FILIAL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150504)
     
    End Select
     
    'Fechamento comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Private Function CustoDiretoFabricacao_Prod_Passo1A(ByVal objCustoDirFabr As ClassCustoDirFabr, ByVal objCustoDirFabrProd As ClassCustoDirFabrProd, ByVal objCustoDirFabrPlan As ClassCustoDirFabrPlan, alComando() As Long, ByVal iNiveis As Integer) As Long
'processa o produto filho de objCustoDirFabrPlan, incorpora-o em objCustoDirFabrProd e se necessario processa os seus filhos

Dim lErro As Long, dPesoFator1 As Double, dPesoFator2 As Double, dPesoFator3 As Double, dPesoFator4 As Double, dPesoFator5 As Double, dPesoFator6 As Double, sVersao As String, iTipo As Integer
Dim colCustoDirFabrPlan As New Collection, dQuantFator1 As Double, dQuantFator2 As Double, dQuantFator3 As Double, dQuantFator4 As Double, dQuantFator5 As Double, dQuantFator6 As Double, iNatureza As Integer
Dim iNivel As Integer, sInsumo As String, dQuantidade As Double, dPercentualPerda As Double, dQuantidadeBase As Double, dQuantidadeProprio As Double, iCompras As Integer
Dim objCustoDirFabrPlanAux As ClassCustoDirFabrPlan, bAchou As Boolean, dQuantPai As Double, objCustoDirFabrPlanMPAux As ClassCustoDirFabrPlanMP
Dim objCustoDirFabrPlanEmbAux As ClassCustoDirFabrPlanEmb, dPercentualPerdaBase As Double

'??? deveria ler a UM em produtokit p/fazer conversoes levando em conta UMs diferentes

On Error GoTo Erro_CustoDiretoFabricacao_Prod_Passo1A

    If iNiveis >= 20 Then gError 130534
    
    'pesquisa na versao padrao na tabela kit os pesos dos fatores
    sVersao = String(STRING_KIT_VERSAO, 0)
    lErro = Comando_Executar(alComando(4), "SELECT Versao, PesoFator1, PesoFator2, PesoFator3, PesoFator4, PesoFator5, PesoFator6 FROM Kit WHERE ProdutoRaiz = ? AND VersaoFormPreco = ?", _
        sVersao, dPesoFator1, dPesoFator2, dPesoFator3, dPesoFator4, dPesoFator5, dPesoFator6, objCustoDirFabrPlan.sProdutoFilho, 1)
    If lErro <> AD_SQL_SUCESSO Then gError 106620
        
    lErro = Comando_BuscarProximo(alComando(4))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106621
    
    'completar objCustoDirFabrPlan
    objCustoDirFabrPlan.dPesoFator1 = dPesoFator1
    objCustoDirFabrPlan.dPesoFator2 = dPesoFator2
    objCustoDirFabrPlan.dPesoFator3 = dPesoFator3
    objCustoDirFabrPlan.dPesoFator4 = dPesoFator4
    objCustoDirFabrPlan.dPesoFator5 = dPesoFator5
    objCustoDirFabrPlan.dPesoFator6 = dPesoFator6
    objCustoDirFabrPlan.iSeq = objCustoDirFabrProd.colPlan.Count + 1
    
    'incluir objCustoDirFabrPlan na colecao de objCustoDirFabrProd
    objCustoDirFabrProd.colPlan.Add objCustoDirFabrPlan
    
    'acumular qtdes multiplicadas pelos fatores em objCustoDirFabrProd
    dQuantPai = objCustoDirFabrPlan.dQuantPai * objCustoDirFabrPlan.dParticipacao
    dQuantFator1 = ArredondaMod(objCustoDirFabrPlan.dPesoFator1 * dQuantPai, 0)
    dQuantFator2 = ArredondaMod(objCustoDirFabrPlan.dPesoFator2 * dQuantPai, 0)
    dQuantFator3 = ArredondaMod(objCustoDirFabrPlan.dPesoFator3 * dQuantPai, 0)
    dQuantFator4 = ArredondaMod(objCustoDirFabrPlan.dPesoFator4 * dQuantPai, 0)
    dQuantFator5 = ArredondaMod(objCustoDirFabrPlan.dPesoFator5 * dQuantPai, 0)
    dQuantFator6 = ArredondaMod(objCustoDirFabrPlan.dPesoFator6 * dQuantPai, 0)
    
    objCustoDirFabrProd.dQuantFator1 = objCustoDirFabrProd.dQuantFator1 + dQuantFator1
    objCustoDirFabrProd.dQuantFator2 = objCustoDirFabrProd.dQuantFator2 + dQuantFator2
    objCustoDirFabrProd.dQuantFator3 = objCustoDirFabrProd.dQuantFator3 + dQuantFator3
    objCustoDirFabrProd.dQuantFator4 = objCustoDirFabrProd.dQuantFator4 + dQuantFator4
    objCustoDirFabrProd.dQuantFator5 = objCustoDirFabrProd.dQuantFator5 + dQuantFator5
    objCustoDirFabrProd.dQuantFator6 = objCustoDirFabrProd.dQuantFator6 + dQuantFator6
    
    'para incluir as embalagens apenas do produto "principal", considerando que as embalagens intermediarias possam ser reaproveitadas e incorporadas a rateio do custo direto
    If objCustoDirFabrProd.colPlan.Count = 1 Then
    
        'incluir embalagem padrao p/o produto
        lErro = CustoDirFabrPlan_ObtemEmb(objCustoDirFabrPlan, alComando)
        If lErro <> SUCESSO Then gError 106837
    
    End If
    
    'pesquisa na versao padrao do produto filho de objCustoDirFabrPlan os componentes de nivel <= 1 ordenados por nivel, seq
    sInsumo = String(STRING_PRODUTO, 0)
    lErro = Comando_Executar(alComando(5), "SELECT Natureza, Compras, Produto, ProdutoKit.Nivel, Quantidade, PercentualPerda FROM ProdutoKit, Produtos WHERE Versao = ? AND ProdutoKit.ProdutoRaiz = ? AND ProdutoKit.Nivel <= 1 AND ProdutoKit.Produto = Produtos.Codigo ORDER BY ProdutoKit.Nivel, Seq", _
        iNatureza, iCompras, sInsumo, iNivel, dQuantidade, dPercentualPerda, sVersao, objCustoDirFabrPlan.sProdutoFilho)
    If lErro <> AD_SQL_SUCESSO Then gError 106622
        
    lErro = Comando_BuscarProximo(alComando(5))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106623
        
    Do While lErro = AD_SQL_SUCESSO
    
        'se o nivel for 0 entao
        If iNivel = 0 Then
            dQuantidadeBase = dQuantidade
            
            lErro = CF("CustoProd_AjustarPerda", objCustoDirFabrPlan.sProdutoFilho, sVersao, objCustoDirFabr.iFilialEmpresa, objCustoDirFabr.iRotinaOrigem, dPercentualPerda)
            If lErro <> SUCESSO Then gError 106836
            
            dPercentualPerdaBase = dPercentualPerda
                    
        Else
            
            'ajusta qtde pela perda
            If dPercentualPerda <> 0 And dPercentualPerda <> 1 And gobjEST.iUsaPerdaInsumosKit Then
                dQuantidade = dQuantidade / (1 - dPercentualPerda)
            End If
            
            If iCompras <> PRODUTO_PRODUZIVEL Then
            
                iTipo = CUSTO_DIRETO_COMPRA_LOCAL
                
            Else
            
                lErro = Produto_ObtemTipoCustoDir(sInsumo, objCustoDirFabrProd.iFilialEmpresa, iTipo, alComando(6))
                If lErro <> SUCESSO Then gError 106836
                
            End If
            
            'se o produto é produzido localmente
            If iTipo = CUSTO_DIRETO_PRODUCAO_LOCAL Then
            
                'para cada insumo encontrado que nao seja o proprio produto
                If sInsumo <> objCustoDirFabrPlan.sProdutoFilho Then
                
                    'inclui-lo em colCustoDirFabrPlan se ainda nao estiver lá. se estiver acumular a qtde
                    bAchou = False
                    For Each objCustoDirFabrPlanAux In colCustoDirFabrPlan
                        If objCustoDirFabrPlanAux.sProdutoFilho = sInsumo Then
                            
                            objCustoDirFabrPlanAux.dParticipacao = objCustoDirFabrPlanAux.dParticipacao + (dQuantidade / dQuantidadeBase)
                            
                            bAchou = True
                            Exit For
                        End If
                    Next
                    
                    If bAchou = False Then
                        Set objCustoDirFabrPlanAux = New ClassCustoDirFabrPlan
                        
                        objCustoDirFabrPlanAux.iFilialEmpresa = objCustoDirFabrProd.iFilialEmpresa
                        objCustoDirFabrPlanAux.iAno = objCustoDirFabrProd.iAno
                        If objCustoDirFabr.iRotinaOrigem = FORMACAO_PRECO_ROTINA_CUSTOSDIRETOS Or objCustoDirFabr.iRotinaOrigem = FORMACAO_PRECO_ROTINA_CUSTOFIXO Then
                            objCustoDirFabrPlanAux.dQuantPai = ArredondaMod(dQuantPai, 0)
                        Else
                            objCustoDirFabrPlanAux.dQuantPai = dQuantPai
                        End If
                        objCustoDirFabrPlanAux.dParticipacao = dQuantidade / dQuantidadeBase
                        objCustoDirFabrPlanAux.sProduto = objCustoDirFabrProd.sProduto
                        objCustoDirFabrPlanAux.sProdutoPai = objCustoDirFabrPlan.sProdutoFilho
                        
                        If iNiveis = 0 Then
                            objCustoDirFabrPlanAux.sProdutoPai1 = sInsumo
                        Else
                            objCustoDirFabrPlanAux.sProdutoPai1 = objCustoDirFabrPlan.sProdutoPai1
                        End If
                        
                        objCustoDirFabrPlanAux.sProdutoFilho = sInsumo
                        
                        colCustoDirFabrPlan.Add objCustoDirFabrPlanAux
                    End If
                    
                Else
                
                    'acumular a sua participacao
                    dQuantidadeProprio = dQuantidadeProprio + dQuantidade
                    
                End If
                
            Else
            
                'se nao for embalagem (para as embalagens há um calculo em separado
                If iNatureza <> NATUREZA_PROD_EMBALAGENS Or (objCustoDirFabrProd.colPlan.Count = 1 And objCustoDirFabrPlan.colEmb.Count = 0) Then
                
                    'se ainda nao estiver na colecao objCustoDirFabrPlan.colMP entao incluir, se estiver acumular a qtde
                    bAchou = False
                    For Each objCustoDirFabrPlanMPAux In objCustoDirFabrPlan.colMP
                        If objCustoDirFabrPlanMPAux.sProdutoMP = sInsumo Then
                        
                            If objCustoDirFabr.iRotinaOrigem = FORMACAO_PRECO_ROTINA_CUSTOSDIRETOS Or objCustoDirFabr.iRotinaOrigem = FORMACAO_PRECO_ROTINA_CUSTOFIXO Then
                                objCustoDirFabrPlanMPAux.dQtde = objCustoDirFabrPlanMPAux.dQtde + ArredondaMod(dQuantPai * dQuantidade / dQuantidadeBase, 0)
                            Else
                                objCustoDirFabrPlanMPAux.dQtde = objCustoDirFabrPlanMPAux.dQtde + dQuantPai * dQuantidade / dQuantidadeBase
                            End If
                            bAchou = True
                            Exit For
                        End If
                    Next
                    
                    If bAchou = False Then
                    
                        Set objCustoDirFabrPlanMPAux = New ClassCustoDirFabrPlanMP
                        
                        objCustoDirFabrPlanMPAux.iFilialEmpresa = objCustoDirFabrPlan.iFilialEmpresa
                        objCustoDirFabrPlanMPAux.iAno = objCustoDirFabrPlan.iAno
                        objCustoDirFabrPlanMPAux.sProduto = objCustoDirFabrPlan.sProduto
                        objCustoDirFabrPlanMPAux.sProdutoPai = objCustoDirFabrPlan.sProdutoFilho
                        objCustoDirFabrPlanMPAux.iSeq = objCustoDirFabrPlan.iSeq
                        objCustoDirFabrPlanMPAux.sProdutoMP = sInsumo
                        
                        If iNiveis = 0 Then
                            objCustoDirFabrPlanMPAux.sProdutoPai1 = sInsumo
                        Else
                            objCustoDirFabrPlanMPAux.sProdutoPai1 = objCustoDirFabrPlan.sProdutoPai1
                        End If
                        
                        If objCustoDirFabr.iRotinaOrigem = FORMACAO_PRECO_ROTINA_CUSTOSDIRETOS Or objCustoDirFabr.iRotinaOrigem = FORMACAO_PRECO_ROTINA_CUSTOFIXO Then
                            objCustoDirFabrPlanMPAux.dQtde = ArredondaMod(dQuantPai * dQuantidade / dQuantidadeBase, 0)
                        Else
                            objCustoDirFabrPlanMPAux.dQtde = dQuantPai * dQuantidade / dQuantidadeBase
                        End If
                        
                        objCustoDirFabrPlan.colMP.Add objCustoDirFabrPlanMPAux
                        
                    End If
                
                End If
            
            End If
            
        End If
        
        lErro = Comando_BuscarProximo(alComando(5))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106624
        
    Loop
    
    'se o produto pai for insumo dele mesmo
    If dQuantidadeProprio <> 0 Then
    
        If Abs(dQuantidadeProprio - dQuantidadeBase) < 0.001 Then
'?????
'            MsgBox (objCustoDirFabrPlan.sProdutoFilho)
'            gError 333333
            
        Else
        
            'percorrer colCustoDirFabrPlan dividindo pela participacao do proprio produto
            For Each objCustoDirFabrPlanAux In colCustoDirFabrPlan
            
                objCustoDirFabrPlanAux.dParticipacao = objCustoDirFabrPlanAux.dParticipacao * (dQuantidadeBase / (dQuantidadeBase - dQuantidadeProprio))
                
            Next
            
            'percorrer objCustoDirFabrPlan.colMP dividindo pela participacao do proprio produto
            For Each objCustoDirFabrPlanMPAux In objCustoDirFabrPlan.colMP
            
                objCustoDirFabrPlanMPAux.dQtde = objCustoDirFabrPlanMPAux.dQtde * (dQuantidadeBase / (dQuantidadeBase - dQuantidadeProprio))
                
            Next
            
            '1BJBLSDXE1:confirmar se deve ou nao considerar as embalagens para o proprio produto
            
            'percorrer objCustoDirFabrPlan.colEmb dividindo pela participacao do proprio produto
            For Each objCustoDirFabrPlanEmbAux In objCustoDirFabrPlan.colEmb
            
                objCustoDirFabrPlanEmbAux.dQtde = objCustoDirFabrPlanEmbAux.dQtde * (dQuantidadeBase / (dQuantidadeBase - dQuantidadeProprio))
                
            Next
    
        End If
    
    End If
        
    'se o produto pai tiver perda
    If dPercentualPerdaBase <> 0 Then
    
        'percorrer colCustoDirFabrPlan dividindo pelo rendimento
        For Each objCustoDirFabrPlanAux In colCustoDirFabrPlan
        
            objCustoDirFabrPlanAux.dParticipacao = objCustoDirFabrPlanAux.dParticipacao / (1 - dPercentualPerdaBase)
            
        Next
        
        'percorrer objCustoDirFabrPlan.colMP dividindo pelo rendimento
        For Each objCustoDirFabrPlanMPAux In objCustoDirFabrPlan.colMP
        
            objCustoDirFabrPlanMPAux.dQtde = objCustoDirFabrPlanMPAux.dQtde / (1 - dPercentualPerdaBase)
            
        Next
        
'        'percorrer objCustoDirFabrPlan.colEmb dividindo pelo rendimento
'        For Each objCustoDirFabrPlanEmbAux In objCustoDirFabrPlan.colEmb
'
'            objCustoDirFabrPlanEmbAux.dQtde = objCustoDirFabrPlanEmbAux.dQtde / (1 - dPercentualPerdaBase)
'
'        Next
    
    End If
    
    'forçar o processamento dos filhos
    For Each objCustoDirFabrPlanAux In colCustoDirFabrPlan
        
        lErro = CustoDiretoFabricacao_Prod_Passo1A(objCustoDirFabr, objCustoDirFabrProd, objCustoDirFabrPlanAux, alComando, iNiveis + 1)
        If lErro <> SUCESSO Then gError 106625
        
    Next
    
    CustoDiretoFabricacao_Prod_Passo1A = SUCESSO
     
    Exit Function
    
Erro_CustoDiretoFabricacao_Prod_Passo1A:

    CustoDiretoFabricacao_Prod_Passo1A = gErr
     
    Select Case gErr
          
        Case 106620 To 106625
            '??? colocar erros adequados
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 130534
            Call Rotina_Erro(vbOKOnly, "ERRO_CUSTOPROD_EM_LOOP", gErr, objCustoDirFabrProd.sProduto, objCustoDirFabrProd.iFilialEmpresa)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150505)
     
    End Select
     
    Exit Function

End Function

Private Function Produto_ObtemTipoCustoDir(ByVal sProduto As String, ByVal iFilialEmpresa As Integer, iTipo As Integer, ByVal lComando As Long) As Long
'obtem tipo de tratamento a ser dado a produto para efeito de custo

Dim lErro As Long, dCusto As Double

On Error GoTo Erro_Produto_ObtemTipoCustoDir

    lErro = Comando_Executar(lComando, "SELECT Custo FROM CustoEmbMP WHERE FilialEmpresa = ? AND Produto = ?", dCusto, iFilialEmpresa, sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 106834
        
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106835
        
    If lErro = AD_SQL_SUCESSO Then
    
        iTipo = CUSTO_DIRETO_PRODUCAO_TRANSFERIDA
    
    Else
    
        iTipo = CUSTO_DIRETO_PRODUCAO_LOCAL
        
    End If
    
    Produto_ObtemTipoCustoDir = SUCESSO
     
    Exit Function
    
Erro_Produto_ObtemTipoCustoDir:

    Produto_ObtemTipoCustoDir = gErr
     
    Select Case gErr
          
        '??? completar
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150506)
     
    End Select
     
    Exit Function

End Function

Private Function CustoDirFabrPlan_ObtemEmb(ByVal objCustoDirFabrPlan As ClassCustoDirFabrPlan, alComando() As Long) As Long
'obtem embalagens necessarias para o produto, incluindo a padrao e as de expedicao (embalagens adicionais, rotulos, paletizacao,...)

Dim dQuantAlocar As Double
Dim dQuantEmb As Double
Dim dQuantProduto As Double
Dim iAlteradoEmb As Integer
Dim iIndice1 As Integer
Dim lErro As Long, objCustoDirFabrPlanEmb As ClassCustoDirFabrPlanEmb
Dim dFator As Double, objCustoDirFabrPlanEmbAux As ClassCustoDirFabrPlanEmb
Dim dPesoLiq As Double
Dim dPesoBruto As Double, dCapacConv As Double
Dim dPesoEmbalagem As Double
Dim objProduto As New ClassProduto, colEmb As New Collection
Dim objProdutoEmbalagem As New ClassProdutoEmbalagem
Dim sEmbExp As String, dQtdeEmbExp As Double, colEmbExp As New Collection

'??? revisar se poderá haver problema de UM entre a do produto e a de produtoembalagem

On Error GoTo Erro_CustoDirFabrPlan_ObtemEmb

    dQuantAlocar = objCustoDirFabrPlan.dQuantPai * objCustoDirFabrPlan.dParticipacao

    objProdutoEmbalagem.sProduto = objCustoDirFabrPlan.sProdutoFilho
        
    'Seleciona embalagem padrao
    lErro = CF("ProdutoEmbalagem_Le_Padrao", objProdutoEmbalagem)
    If lErro <> SUCESSO And lErro <> 100000 Then gError 96115
            
    If lErro = SUCESSO Then
    
        objProduto.sCodigo = objProdutoEmbalagem.sProduto

        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 96116
        
        lErro = CustoDirFabrPlan_ObtemEmb1(objProdutoEmbalagem, dQuantAlocar, dQuantAlocar / objProdutoEmbalagem.dCapacidade, colEmb)
        If lErro <> SUCESSO Then gError 81716
        
'??? o codigo abaixo foi comentado pois fazia o calculo "exato" das embalagens, o que dava um resultado errado em se tratando de uma analise de custo
'        dQuantEmb = Fix(dQuantAlocar / objProdutoEmbalagem.dCapacidade)
'        dQuantProduto = Round(objProdutoEmbalagem.dCapacidade * dQuantEmb, 4)
'        dQuantAlocar = dQuantAlocar - dQuantProduto
'
'        If dQuantEmb <> 0 Then
'
'            lErro = CustoDirFabrPlan_ObtemEmb1(objProdutoEmbalagem, dQuantProduto, dQuantEmb, colEmb)
'            If lErro <> SUCESSO Then gError 81716
'
'        End If
'
'        Do While dQuantAlocar > QTDE_ESTOQUE_DELTA
'
'            'Le a capacidade maxima onde capacidade <= dQuantAlocar
'            lErro = CF("ProdutoEmbalagem_Le_CapacMaxima", objProdutoEmbalagem, dQuantAlocar, objProduto.sSiglaUMEstoque, objProduto, dFator)
'            If lErro <> SUCESSO And lErro <> 100008 Then gError 96119
'
'            If lErro <> 100008 Then
'
'                dQuantProduto = dQuantAlocar
'                dQuantAlocar = dQuantAlocar * dFator
'
'                'Resultado com casa decimais
'                dQuantEmb = dQuantAlocar / objProdutoEmbalagem.dCapacidade
'
'                'Arrendonda para cima o numero de embalagens
'                If dQuantEmb - Fix(dQuantEmb) <> 0 Then dQuantEmb = Fix(dQuantEmb) + 1
'
'                dQuantAlocar = dQuantAlocar - dQuantProduto
'                dQuantAlocar = dQuantAlocar / dFator
'
'                lErro = CustoDirFabrPlan_ObtemEmb1(objProdutoEmbalagem, dQuantProduto, dQuantEmb, colEmb)
'                If lErro <> SUCESSO Then gError 81717
'
'            Else
'
'                'Le a capacidade minima onde capacidade > dQuantAlocar
'                lErro = CF("ProdutoEmbalagem_Le_CapacMinima", objProdutoEmbalagem, dQuantAlocar, objProduto.sSiglaUMEstoque, objProduto, dFator)
'                If lErro <> SUCESSO And lErro <> 100009 Then gError 96120
'
'                dQuantProduto = dQuantAlocar * dFator
'                dQuantEmb = 1
'                dQuantAlocar = 0
'
'                lErro = CustoDirFabrPlan_ObtemEmb1(objProdutoEmbalagem, dQuantProduto, dQuantEmb, colEmb)
'                If lErro <> SUCESSO Then gError 81718
'
'            End If
'
'        Loop
        
        For Each objCustoDirFabrPlanEmb In colEmb
        
            objCustoDirFabrPlan.colEmb.Add objCustoDirFabrPlanEmb
            With objCustoDirFabrPlanEmb
                .iFilialEmpresa = objCustoDirFabrPlan.iFilialEmpresa
                .iAno = objCustoDirFabrPlan.iAno
                .sProduto = objCustoDirFabrPlan.sProduto
                .iSeq = objCustoDirFabrPlan.iSeq
            End With
            
            sEmbExp = String(STRING_PRODUTO, 0)
        
            lErro = Comando_Executar(alComando(7), "SELECT Produto, Quantidade FROM EmbalagensExpedicao WHERE Embalagem = ?", sEmbExp, dQtdeEmbExp, objCustoDirFabrPlanEmb.iEmbalagem)
            If lErro <> AD_SQL_SUCESSO Then gError 106838
            
            lErro = Comando_BuscarProximo(alComando(7))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106839
        
            Do While lErro = AD_SQL_SUCESSO
            
                Set objCustoDirFabrPlanEmbAux = New ClassCustoDirFabrPlanEmb
                
                With objCustoDirFabrPlanEmbAux
                    
                    .iFilialEmpresa = objCustoDirFabrPlan.iFilialEmpresa
                    .iAno = objCustoDirFabrPlan.iAno
                    .sProduto = objCustoDirFabrPlan.sProduto
                    .iSeq = objCustoDirFabrPlan.iSeq
                    .sProdutoEmb = sEmbExp
                    .dQtde = dQtdeEmbExp * objCustoDirFabrPlanEmb.dQtde
                    
                End With
                
                objCustoDirFabrPlan.colEmb.Add objCustoDirFabrPlanEmbAux
            
                lErro = Comando_BuscarProximo(alComando(7))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106840
            
            Loop
        
        Next
    
    End If
    
    CustoDirFabrPlan_ObtemEmb = SUCESSO

    Exit Function

Erro_CustoDirFabrPlan_ObtemEmb:

    CustoDirFabrPlan_ObtemEmb = gErr

    Select Case gErr
        'Janaina - Tratar mensagens de erro
        Case 96115, 96116, 96117, 96118, 96119, 96120, 81716, 81717, 81718

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150507)

    End Select

    Exit Function
'Janaina
End Function

Private Function CustoDirFabrPlan_ObtemEmb1(objProdutoEmbalagem As ClassProdutoEmbalagem, ByVal dQuantProduto As Double, ByVal dQuantEmb As Double, ByVal colEmb As Collection) As Long
'Calculo automatico das embalagens

Dim lErro As Long
Dim dFator As Double, bAchou As Boolean
Dim dPesoLiq As Double, objCustoDirFabrPlanEmb As ClassCustoDirFabrPlanEmb
Dim dPesoBruto As Double, iLinhaExistente As Integer
Dim dPesoEmbalagem As Double, iIndice As Integer
Dim objEmbalagem As New ClassEmbalagem

On Error GoTo Erro_CustoDirFabrPlan_ObtemEmb1

    'verifica se o par produto-embalagem já está na colecao
        
    'Le a sigla da embalagem
    objEmbalagem.iCodigo = objProdutoEmbalagem.iEmbalagem
    
    lErro = CF("Embalagem_Le", objEmbalagem)
    If lErro <> SUCESSO Then gError 96118
            
    bAchou = False
    
    For Each objCustoDirFabrPlanEmb In colEmb
    
        If objCustoDirFabrPlanEmb.sProdutoEmb = objEmbalagem.sProduto Then
        
            objCustoDirFabrPlanEmb.dQtde = objCustoDirFabrPlanEmb.dQtde + dQuantEmb
            
            bAchou = True
            Exit For
            
        End If
        
    Next
    
    If bAchou = False Then
    
        Set objCustoDirFabrPlanEmb = New ClassCustoDirFabrPlanEmb
        
        objCustoDirFabrPlanEmb.dQtde = dQuantEmb
        objCustoDirFabrPlanEmb.iEmbalagem = objProdutoEmbalagem.iEmbalagem
        If objEmbalagem.sProduto = "" Then gError 106973
        objCustoDirFabrPlanEmb.sProdutoEmb = objEmbalagem.sProduto
        
        Call colEmb.Add(objCustoDirFabrPlanEmb)
        
    End If
    
    CustoDirFabrPlan_ObtemEmb1 = SUCESSO

    Exit Function

Erro_CustoDirFabrPlan_ObtemEmb1:

    CustoDirFabrPlan_ObtemEmb1 = gErr

    Select Case gErr

        Case 96117, 96118
        
        Case 106973
            Call Rotina_Erro(vbOKOnly, "ERRO_EMABALAGEM_SEM_CODIGO_PRODUTO", gErr, objProdutoEmbalagem.iEmbalagem)
            Resume Next
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150508)

    End Select

    Exit Function

End Function

Function FPreco_ObtemValorCateg(ByVal objMnemonicoValor As ClassMnemonicoValor, ByVal objExeExp As ClassExeExp) As Long

Dim lErro As Long, lComando As Long, sProduto As String, dValor As Double, dvalor1 As Double, dvalor2 As Double, dvalor3 As Double, sCategoria As String, iValor As Integer, adValor(1 To 8) As Double

On Error GoTo Erro_FPreco_ObtemValorCateg

    sProduto = objExeExp.objContexto.sProduto
    sCategoria = objMnemonicoValor.vParam(1)
    iValor = StrParaInt(objMnemonicoValor.vParam(2))
    
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 106843
    
    lErro = Comando_Executar(lComando, "SELECT Valor1, Valor2, Valor3, Valor4, Valor5, Valor6, Valor7, Valor8 FROM produtocategoria, categoriaprodutoitem WHERE produtocategoria.categoria = categoriaprodutoitem.categoria AND produtocategoria.Item = categoriaprodutoitem.Item AND produtocategoria.Produto = ? AND produtocategoria.Categoria = ?", _
        adValor(1), adValor(2), adValor(3), adValor(4), adValor(5), adValor(6), adValor(7), adValor(8), sProduto, sCategoria)
    If lErro <> AD_SQL_SUCESSO Then gError 106844

    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106845

    If lErro = AD_SQL_SUCESSO Then
    
        dValor = adValor(iValor)
    
    End If
        
    objMnemonicoValor.colValor.Add dValor
    
    Call Comando_Fechar(lComando)
    
    FPreco_ObtemValorCateg = SUCESSO
     
    Exit Function
    
Erro_FPreco_ObtemValorCateg:

    FPreco_ObtemValorCateg = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150509)
     
    End Select
     
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function ICMSAliqFilCli(ByVal objMnemonicoValor As ClassMnemonicoValor, ByVal objExeExp As ClassExeExp) As Long
'retorna a aliquota de icms associada a filial do cliente

Dim lErro As Long, sUFCli As String, sInscricaoEstadual As String, iPais As Integer
Dim lComando As Long, lCliente As Long, iFilial As Integer, objEstado As New ClassEstado
Dim iFilialFaturamento As Integer, dAliquota As Double, objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_ICMSAliqFilCli

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 106847

    lCliente = StrParaLong(objMnemonicoValor.vParam(1))
    iFilial = StrParaInt(objMnemonicoValor.vParam(2))
    iFilialFaturamento = objExeExp.objContexto.iFilialFaturamento
    
    If lCliente <> 0 And iFilial <> 0 And iFilialFaturamento <> 0 Then
    
        objFilialEmpresa.iCodFilial = iFilialFaturamento
        lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
        If lErro <> SUCESSO Then gError 106848
        
        sUFCli = String(STRING_ESTADO_SIGLA, 0)
        sInscricaoEstadual = String(STRING_INSCR_EST, 0)
    
        lErro = Comando_Executar(lComando, "SELECT CodigoPais, SiglaEstado, InscricaoEstadual FROM FiliaisClientes, Enderecos WHERE FiliaisClientes.CodCliente = ? AND FiliaisClientes.CodFilial = ? AND FiliaisClientes.Endereco = Enderecos.Codigo", iPais, sUFCli, sInscricaoEstadual, lCliente, iFilial)
        If lErro <> AD_SQL_SUCESSO Then gError 106849

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106850

        If iPais <> PAIS_BRASIL Then
        
            dAliquota = 0
            
        Else
        
            If Len(Trim(sInscricaoEstadual)) = 0 Or objFilialEmpresa.objEndereco.sSiglaEstado = sUFCli Then
            
                objEstado.sSigla = objFilialEmpresa.objEndereco.sSiglaEstado
                lErro = CF("Estado_Le", objEstado)
                If lErro <> SUCESSO Then gError 106851
    
                dAliquota = objEstado.dICMSAliquotaInterna
            
            Else
            
                lErro = Comando_Executar(lComando, "SELECT Aliquota FROM ICMSAliquotasExternas WHERE SiglaOrig = ? AND SiglaDest = ?", dAliquota, objFilialEmpresa.objEndereco.sSiglaEstado, sUFCli)
                If lErro <> AD_SQL_SUCESSO Then gError 106852
        
                lErro = Comando_BuscarProximo(lComando)
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106853
                
            End If
        
        End If
    
    End If
    
    Call Comando_Fechar(lComando)
    
    objMnemonicoValor.colValor.Add CDbl(dAliquota)
    
    ICMSAliqFilCli = SUCESSO
     
    Exit Function
    
Erro_ICMSAliqFilCli:

    ICMSAliqFilCli = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150510)
     
    End Select
     
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function FilialEmpresa_ObtemDataCustoFixo(ByVal iFilialFaturamento As Integer, dtDataCustoFixo As Date) As Long

Dim lErro As Long
Dim lComando As Long

On Error GoTo Erro_FilialEmpresa_ObtemDataCustoFixo

    dtDataCustoFixo = DATA_NULA
    
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 106855

    lErro = Comando_Executar(lComando, "SELECT MAX(DataReferencia) FROM CustoFixo WHERE FilialEmpresa = ?", dtDataCustoFixo, iFilialFaturamento)
    If lErro <> AD_SQL_SUCESSO Then gError 106856

    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106857
    
    Call Comando_Fechar(lComando)
    
    FilialEmpresa_ObtemDataCustoFixo = SUCESSO
     
    Exit Function
    
Erro_FilialEmpresa_ObtemDataCustoFixo:

    FilialEmpresa_ObtemDataCustoFixo = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150511)
     
    End Select
     
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Private Function ComissoesCliProd(ByVal objMnemonicoValor As ClassMnemonicoValor, ByVal objExeExp As ClassExeExp, ByVal bDireta As Boolean) As Long
'parametros:Cliente, Filial, Produto

Dim lErro As Long, dPerc As Double, objMnemonicoValorAux As New ClassMnemonicoValor
Dim colComissoes As New Collection, dPrecoUnitario As Double
Dim objComissao As ClassComissaoPedVendas, dAcumValComiss As Double

On Error GoTo Erro_ComissoesCliProd

    If Not (objExeExp.objContexto.colComissoes Is Nothing) Then
    
        If objExeExp.objContexto.dQuantidade <> 0 Then
        
            Set colComissoes = objExeExp.objContexto.colComissoes
            dPrecoUnitario = objExeExp.objContexto.dPrecoPraticado
            
            'percorre a colecao de comissoes obtendo o % proprio
            For Each objComissao In colComissoes
            
                If (bDireta And objComissao.iIndireta = 0) Or (Not bDireta And objComissao.iIndireta <> 0) Then
                    dAcumValComiss = dAcumValComiss + objComissao.dValor
                End If
                
            Next
            
            If (dPrecoUnitario * objExeExp.objContexto.dQuantidade) <> 0 Then
                dPerc = dAcumValComiss / (dPrecoUnitario * objExeExp.objContexto.dQuantidade)
            Else
                dPerc = 0
            End If
        
        End If
    
    End If
    
    objMnemonicoValor.colValor.Add CDbl(dPerc)
    
    ComissoesCliProd = SUCESSO

    Exit Function

Erro_ComissoesCliProd:

    ComissoesCliProd = gErr

    Select Case gErr

        Case 106860, 106861
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150512)

    End Select

    Exit Function

End Function

Function PrecoTabela(ByVal objMnemonicoValor As ClassMnemonicoValor, ByVal objExeExp As ClassExeExp) As Long
'retorna o preco de tabela de um produto para a filial de faturamento

Dim lErro As Long
Dim objTabelaPrecoItem As New ClassTabelaPrecoItem
Dim dPrecoUnitario As Double

On Error GoTo Erro_PrecoTabela

    objTabelaPrecoItem.iCodTabela = objExeExp.objContexto.iTabelaPreco
    objTabelaPrecoItem.iFilialEmpresa = objExeExp.objContexto.iFilialFaturamento
    objTabelaPrecoItem.sCodProduto = objExeExp.objContexto.sProduto
    
    lErro = CF("TabelaPrecoItem_Le", objTabelaPrecoItem)
    If lErro <> SUCESSO And lErro <> 28014 Then gError 106861
    
    If lErro = 28014 Then
        objTabelaPrecoItem.iFilialEmpresa = EMPRESA_TODA
    
        lErro = CF("TabelaPrecoItem_Le", objTabelaPrecoItem)
        If lErro <> SUCESSO And lErro <> 28014 Then gError 106862
     End If

    If lErro <> 28014 Then dPrecoUnitario = objTabelaPrecoItem.dPreco
    
    objMnemonicoValor.colValor.Add CDbl(dPrecoUnitario)
    
    PrecoTabela = SUCESSO
     
    Exit Function
    
Erro_PrecoTabela:

    PrecoTabela = gErr
     
    Select Case gErr
          
        Case 106861, 106862
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150513)
     
    End Select
     
    Exit Function

End Function

Function EncComissoesPropCliProd(ByVal objMnemonicoValor As ClassMnemonicoValor, ByVal objExeExp As ClassExeExp) As Long
'parametros:Cliente, Filial, Produto

    EncComissoesPropCliProd = EncComissoesCliProd(objMnemonicoValor, objExeExp, True)

End Function

Function EncComissoesTercCliProd(ByVal objMnemonicoValor As ClassMnemonicoValor, ByVal objExeExp As ClassExeExp) As Long
'parametros:Cliente, Filial, Produto

    EncComissoesTercCliProd = EncComissoesCliProd(objMnemonicoValor, objExeExp, False)
    
End Function

Private Function EncComissoesCliProd(ByVal objMnemonicoValor As ClassMnemonicoValor, ByVal objExeExp As ClassExeExp, ByVal bDireta As Boolean) As Long
'parametros:Cliente, Filial, Produto

Dim lErro As Long, dPerc As Double, objMnemonicoValorAux As New ClassMnemonicoValor
Dim colComissoes As New Collection, dPrecoUnitario As Double
Dim objComissao As ClassComissaoPedVendas, dAcumValEnc As Double
Dim objVendedor As New ClassVendedor, dFator As Double

On Error GoTo Erro_EncComissoesCliProd

    If Not (objExeExp.objContexto.colComissoes Is Nothing) Then
    
        If objExeExp.objContexto.dQuantidade <> 0 Then
        
            Set colComissoes = objExeExp.objContexto.colComissoes
            dPrecoUnitario = objExeExp.objContexto.dPrecoPraticado
            
            'percorre a colecao de comissoes obtendo o % proprio
            For Each objComissao In colComissoes
            
                If (bDireta And objComissao.iIndireta = 0) Or (Not bDireta And objComissao.iIndireta <> 0) Then
                    
                    If objComissao.iCodVendedor <> 0 Then
                    
                        objVendedor.iCodigo = objComissao.iCodVendedor
                        lErro = CF("Vendedor_Le", objVendedor)
                        If lErro <> SUCESSO And lErro <> 12582 Then gError 106863
                        If lErro <> SUCESSO Then gError 106864
                        
                        Select Case objVendedor.iVinculo
                        
                            Case VENDEDOR_VINCULO_AUTONOMO, VENDEDOR_VINCULO_EMPRESA
                                dFator = 0.1111
                            
                            Case VENDEDOR_VINCULO_EMPREGADO
                                dFator = 0.59
                                Call CF("EncComissoesCliProd_Cust", objVendedor, dFator)
                            
                            Case Else
                                dFator = 0
                            
                        End Select
                    
                    Else
                    
                        dFator = 0
                        
                    End If
                    
                    dAcumValEnc = dAcumValEnc + (objComissao.dValor * dFator)
                
                End If
                
            Next
            
            If (dPrecoUnitario * objExeExp.objContexto.dQuantidade) <> 0 Then
                dPerc = dAcumValEnc / (dPrecoUnitario * objExeExp.objContexto.dQuantidade)
            Else
                dPerc = 0
            End If
        
        End If
    
    End If
    
    objMnemonicoValor.colValor.Add CDbl(dPerc)
    
    EncComissoesCliProd = SUCESSO

    Exit Function

Erro_EncComissoesCliProd:

    EncComissoesCliProd = gErr

    Select Case gErr

        Case 106863
        
        Case 106864 '???
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150514)

    End Select

    Exit Function

End Function

Function CalcMP_ObterValores(ByVal iFilialEmpresa As Integer, dTaxaDescPadrao As Double, dTaxaValFut As Double, dDiasValFut As Double) As Long
'obter valores globais para calculo de valor presente e futuro de insumo

Dim lErro As Long, lComando As Long
Dim sMnemonico As String, sExpressao As String

On Error GoTo Erro_CalcMP_ObterValores

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 106883
    
    sMnemonico = String(STRING_MNEMONICOFPRECO_MNEMONICO, 0)
    sExpressao = String(STRING_MNEMONICOFPRECO_EXPRESSAO, 0)
    lErro = Comando_Executar(lComando, "SELECT Mnemonico, Expressao FROM MnemonicoFPTipo WHERE TipoPlanilha = ? AND FilialEmpresa = ? AND Escopo = ? AND Mnemonico IN ('TaxaDescPadrao', 'TaxaValFut', 'DiasValFut' )", _
        sMnemonico, sExpressao, PLANILHA_TIPO_TODOS, iFilialEmpresa, MNEMONICOFPRECO_ESCOPO_GERAL)
    If lErro <> AD_SQL_SUCESSO Then gError 106778
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106779
        
    Do While lErro = AD_SQL_SUCESSO
        
        Select Case sMnemonico
        
            Case "TaxaDescPadrao"
                dTaxaDescPadrao = StrParaDbl(sExpressao)
                
            Case "TaxaValFut"
                dTaxaValFut = StrParaDbl(sExpressao)
            
            Case "DiasValFut"
                dDiasValFut = StrParaDbl(sExpressao)
        
        End Select
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106780
    
    Loop
    
    Call Comando_Fechar(lComando)
    
    CalcMP_ObterValores = SUCESSO
     
    Exit Function
    
Erro_CalcMP_ObterValores:

    CalcMP_ObterValores = gErr
     
    Select Case gErr
          
        '??? completar
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150515)
     
    End Select
     
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function Produto_ObterDadosUltCompra(ByVal objProduto As ClassProduto, ByVal iFilialEmpresa As Integer, dCusto As Double, dtData As Date, dAliquotaICMS As Double, sMoeda As String, dTaxa As Double, dFrete As Double, iCondPagto As Integer) As Long
'obtem dados para analise de custo da ultima cotacao de um produto

Dim lErro As Long, lComando As Long, dQuantidade As Double, objMoeda As New ClassMoedas
Dim dFator As Double, sUnidadeMed As String, iMoeda As Integer

On Error GoTo Erro_Produto_ObterDadosUltCompra

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 106729
    
    sUnidadeMed = String(STRING_UM_SIGLA, 0)
    
    lErro = Comando_Executar(lComando, "SELECT  Data, PrecoUnitario, Moeda, ValorFrete, AliquotaICMS, Taxa, Quantidade, UM, CondicaoPagto " & _
        " From PedidoCompraN, ItensPedCompraN WHERE PedidoCompraN.NumIntDoc = ItensPedCompraN.PedCompra AND FilialEmpresa = ? AND Produto = ? ORDER BY PedidoCompraN.DataEnvio DESC, PedidoCompraN.Codigo DESC", _
        dtData, dCusto, iMoeda, dFrete, dAliquotaICMS, dTaxa, dQuantidade, sUnidadeMed, iCondPagto, iFilialEmpresa, objProduto.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 106730
        
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106731
        
    If lErro <> AD_SQL_SUCESSO Then
    
        dtData = DATA_NULA
        dCusto = 0
        dAliquotaICMS = 0
        sMoeda = ""
        iCondPagto = 0
        dTaxa = 0
        dFrete = 0
        iCondPagto = 0
        
    Else
    
        If iMoeda = 0 Then
            sMoeda = ""
        Else
            sMoeda = String(STRING_SIMBOLO_MOEDA, 0)
            objMoeda.iCodigo = iMoeda
            lErro = CF("Moedas_le", objMoeda)
            If lErro <> SUCESSO And lErro <> 108821 Then gError 106732
            
            If lErro = SUCESSO Then
                sMoeda = objMoeda.sSimbolo
            Else
                sMoeda = ""
            End If
        End If
        
        'converter o custo p/a unidade de estoque se necessario
        If objProduto.sSiglaUMEstoque <> sUnidadeMed Then
        
            lErro = CF("UM_Conversao", objProduto.iClasseUM, objProduto.sSiglaUMEstoque, sUnidadeMed, dFator)
            If lErro <> SUCESSO Then gError 106733
            
            dCusto = Round(dCusto * dFator, 4)
        
        End If
        
    End If
        
    Call Comando_Fechar(lComando)
    
    Produto_ObterDadosUltCompra = SUCESSO
     
    Exit Function
    
Erro_Produto_ObterDadosUltCompra:

    Produto_ObterDadosUltCompra = gErr
     
    Select Case gErr
          
        Case 106729
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 106730, 106731
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_INFOULTCOMPRA", Err)
        
        Case 106732, 106733
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150516)
     
    End Select
     
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function Produto_ObterDadosUltEnt(ByVal objProduto As ClassProduto, ByVal iFilialEmpresa As Integer, dCusto As Double, dtData As Date, dAliquotaICMS As Double, dFrete As Double) As Long
'obtem dados para analise de custo da ultima entrada por compra de um produto

'??? poderia tentar incluir custo de nf tipo 165 (despesas com despacho aduaneiro e fretes)

Dim lErro As Long, lComando As Long, dQuantidade As Double, lComando2 As Long, lNumIntNF As Long
Dim dFator As Double, sUnidadeMed As String, dtDataUltFrete As Date, dCustoUltFrete As Double

On Error GoTo Erro_Produto_ObterDadosUltEnt

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 106716
    
    sUnidadeMed = String(STRING_UM_SIGLA, 0)
    lErro = Comando_Executar(lComando, "SELECT NFiscal.Numintdoc, Quantidade, UnidadeMed, DataEntrada, PrecoUnitario, ICMSAliquota FROM NFiscal, ItensNfiscal, Tiposdocinfo, TributacaoItemNF WHERE Tiposdocinfo.codigo = NFiscal.Tiponfiscal AND (Tiposdocinfo.Compras = 1 or TiposDocInfo.Codigo = ?) AND NFiscal.Numintdoc = Itensnfiscal.Numintnf AND Itensnfiscal.Produto = ? and NFiscal.status <> ? AND Filialempresa = ? AND Tiposdocinfo.Complementar = 0 AND Tiposdocinfo.codigo NOT IN (104,105,165,166, 168, 170) AND TributacaoItemNF.Numintnf = NFiscal.numintdoc and Tributacaoitemnf.item = Itensnfiscal.item ORDER BY DataEntrada DESC", _
        lNumIntNF, dQuantidade, sUnidadeMed, dtData, dCusto, dAliquotaICMS, DOCINFO_NFEEC, objProduto.sCodigo, STATUS_CANCELADO, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 106717
        
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106718
        
    If lErro <> AD_SQL_SUCESSO Then
    
        dtData = DATA_NULA
        dCusto = 0
        dAliquotaICMS = 0
        dFrete = 0
        
    Else
    
        lErro = CF("Produto_ObterFreteUltEnt", objProduto, iFilialEmpresa, dtDataUltFrete, dCustoUltFrete, lNumIntNF)
        If lErro <> SUCESSO Then gError 106889
        
        If dtDataUltFrete <> DATA_NULA Then
            dFrete = dCustoUltFrete
        Else
            dFrete = 0
        End If
        
        'converter os custos p/a unidade de estoque se necessario
        If objProduto.sSiglaUMEstoque <> sUnidadeMed Then
        
            lErro = CF("UM_Conversao", objProduto.iClasseUM, objProduto.sSiglaUMEstoque, sUnidadeMed, dFator)
            If lErro <> SUCESSO Then gError 106719
            
            dCusto = Round(dCusto * dFator, 4)
        
        End If
        
    End If
        
    Call Comando_Fechar(lComando)
    
    Produto_ObterDadosUltEnt = SUCESSO
     
    Exit Function
    
Erro_Produto_ObterDadosUltEnt:

    Produto_ObterDadosUltEnt = gErr
     
    Select Case gErr
          
        Case 106716
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 106717, 106718
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_INFOULTCOMPRA", Err)
        
        Case 106719
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150517)
     
    End Select
     
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function Produto_ObterDadosUltCot(ByVal objProduto As ClassProduto, ByVal iFilialEmpresa As Integer, dCusto As Double, dtData As Date, dAliquotaICMS As Double, sMoeda As String, dTaxa As Double, dFrete As Double) As Long
'obtem dados para analise de custo da ultima cotacao de um produto

Dim lErro As Long, lComando As Long, dQuantidade As Double, objMoeda As New ClassMoedas
Dim dFator As Double, sUnidadeMed As String, iMoeda As Integer

On Error GoTo Erro_Produto_ObterDadosUltCot

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 106726
    
    sUnidadeMed = String(STRING_UM_SIGLA, 0)
    
    lErro = Comando_Executar(lComando, "SELECT Data, PrecoUnitario, Moeda, (ValorFrete*Quantidade) As Frete, AliquotaICMS, Taxa, Quantidade, UM " & _
        " From PedidoCotacaoN, ItemPedCotacaoN, CotacaoProdutoN, ItensCotacaoN WHERE   PedidoCotacaoN.NumIntDoc = ItemPedCotacaoN.PedidoCotacao AND " & _
        " CotacaoProdutoN.NumIntDoc = ItemPedCotacaoN.CotacaoProduto AND FilialEmpresa = ? AND Produto = ? AND " & _
        " ItensCotacaoN.ItemPedCotacao = ItemPedCotacaoN.NumIntDoc ORDER BY PedidoCotacaoN.Data DESC, PedidoCotacaoN.Codigo DESC", _
        dtData, dCusto, iMoeda, dFrete, dAliquotaICMS, dTaxa, dQuantidade, sUnidadeMed, iFilialEmpresa, objProduto.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 106724
        
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106725
        
    If lErro <> AD_SQL_SUCESSO Then
    
        dtData = DATA_NULA
        dCusto = 0
        dAliquotaICMS = 0
        sMoeda = ""
        
    Else
    
        If iMoeda = 0 Then
            sMoeda = ""
        Else
            sMoeda = String(STRING_SIMBOLO_MOEDA, 0)
            objMoeda.iCodigo = iMoeda
            lErro = CF("Moedas_le", objMoeda)
            If lErro <> SUCESSO And lErro <> 108821 Then gError 106727
            
            If lErro = SUCESSO Then
                sMoeda = objMoeda.sSimbolo
            Else
                sMoeda = ""
            End If
        End If
        
        'converter o custo p/a unidade de estoque se necessario
        If objProduto.sSiglaUMEstoque <> sUnidadeMed Then
        
            lErro = CF("UM_Conversao", objProduto.iClasseUM, objProduto.sSiglaUMEstoque, sUnidadeMed, dFator)
            If lErro <> SUCESSO Then gError 106728
            
            dCusto = Round(dCusto * dFator, 4)
        
        End If
        
    
    End If
        
    Call Comando_Fechar(lComando)
    
    Produto_ObterDadosUltCot = SUCESSO
     
    Exit Function
    
Erro_Produto_ObterDadosUltCot:

    Produto_ObterDadosUltCot = gErr
     
    Select Case gErr
          
        Case 106726
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 106724, 106725
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_INFOULTCOTACAO", Err)
        
        Case 106727, 106728
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150518)
     
    End Select
     
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function Produto_ObterFreteUltEnt(ByVal objProduto As ClassProduto, ByVal iFilialEmpresa As Integer, dtData As Date, dFrete As Double, Optional ByVal lNumIntNF As Long = 0) As Long
'obtem custo unitario (na um de estoque) do frete da ultima entrada por compra de um produto

Dim lErro As Long, lComando As Long, lComando1 As Long, dQuantidade As Double
Dim dFator As Double, sUnidadeMed As String, dValorTotal As Double

On Error GoTo Erro_Produto_ObterFreteUltEnt

    dtData = DATA_NULA
    dFrete = 0
        
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 106716
    
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 106716
    
    If lNumIntNF = 0 Then
    
        lErro = Comando_Executar(lComando1, "SELECT NFiscal.Numintdoc FROM NFiscal, ItensNfiscal, Tiposdocinfo, TributacaoItemNF WHERE Tiposdocinfo.codigo = NFiscal.Tiponfiscal AND (Tiposdocinfo.Compras = 1 or TiposDocInfo.Codigo = ?) AND NFiscal.Numintdoc = Itensnfiscal.Numintnf AND Itensnfiscal.Produto = ? and NFiscal.status <> ? AND Filialempresa = ? AND Tiposdocinfo.Complementar = 0 AND Tiposdocinfo.codigo NOT IN (104,105,165,166, 168, 170) AND TributacaoItemNF.Numintnf = NFiscal.numintdoc and Tributacaoitemnf.item = Itensnfiscal.item ORDER BY DataEntrada DESC", _
            lNumIntNF, DOCINFO_NFEEC, objProduto.sCodigo, STATUS_CANCELADO, iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 130038
        
        lErro = Comando_BuscarProximo(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130039
        If lErro <> AD_SQL_SUCESSO Then lNumIntNF = 0
        
    End If
    
    If lNumIntNF <> 0 Then
    
        sUnidadeMed = String(STRING_UM_SIGLA, 0)
        lErro = Comando_Executar(lComando, "SELECT ItensNFEnt.Quantidade, ItensNFEnt.UnidadeMed, NFEnt.DataEntrada, MovimentoEstoque.Custo FROM MovimentoEstoque, NFiscal ConhecTransp, ItensNFiscal ItensNFFrete, NFiscal NFEnt, ItensNfiscal ItensNFEnt, Tiposdocinfo WHERE NFEnt.NumIntDoc = ConhecTransp.NumIntNotaOriginal AND (ConhecTransp.TipoNFiscal = 104 or ConhecTransp.TipoNFiscal = 170) AND Tiposdocinfo.codigo = NFEnt.Tiponfiscal AND (Tiposdocinfo.Compras = 1 or TiposDocInfo.Codigo = ?) AND NFEnt.Numintdoc = ItensNFEnt.Numintnf AND ItensNFEnt.Produto = ? and NFEnt.status <> ? AND NFEnt.Filialempresa = ? AND Tiposdocinfo.codigo NOT IN (104,105,165,166, 168, 170) AND MovimentoEstoque.NumIntDocOrigem = ItensNFFrete.NumIntDoc AND MovimentoEstoque.TipoNumIntDocOrigem = ? AND ItensNFFrete.NumIntNF = ConhecTransp.NumIntDoc AND ItensNFFrete.Produto = ItensNFEnt.Produto AND NFEnt.NumIntDoc = ?", _
            dQuantidade, sUnidadeMed, dtData, dValorTotal, DOCINFO_NFEEC, objProduto.sCodigo, STATUS_CANCELADO, iFilialEmpresa, MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL, lNumIntNF)
        If lErro <> AD_SQL_SUCESSO Then gError 106717
            
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106718
            
        If lErro = AD_SQL_SUCESSO Then
        
            If dQuantidade <> 0 Then
                
                dFrete = dValorTotal / dQuantidade
                
                'converter os custos p/a unidade de estoque se necessario
                If objProduto.sSiglaUMEstoque <> sUnidadeMed Then
                
                    lErro = CF("UM_Conversao", objProduto.iClasseUM, objProduto.sSiglaUMEstoque, sUnidadeMed, dFator)
                    If lErro <> SUCESSO Then gError 106719
                    
                    dFrete = Round(dFrete * dFator, 4)
                
                End If
                
            End If
        
        End If
            
    End If
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    Produto_ObterFreteUltEnt = SUCESSO
     
    Exit Function
    
Erro_Produto_ObterFreteUltEnt:

    Produto_ObterFreteUltEnt = gErr
     
    Select Case gErr
          
        Case 106716
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 106717, 106718, 130038, 130039
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_INFOULTCOMPRA", Err)
        
        Case 106719
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150519)
     
    End Select
     
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    Exit Function

End Function

Function Produto_ObterEmbMP(ByVal sProduto As String, ByVal iFilialEmpresa As Integer, colMP As Collection, colEmb As Collection) As Long
'obtem os componentes, separados em embalagens e mp, com as respectivas participacoes e custos calculados.
'a colecao de mps deverá incluir apenas o 1o nivel de filhos, com produtos intermediarios incorporando o custo das materias primas de seus "filhos"

Dim lErro As Long, iIndice As Integer, iIndice2 As Integer
Dim objCustoDirFabr As New ClassCustoDirFabr, objCustoDirFabrProd As ClassCustoDirFabrProd
Dim objCustoDirFabrPlan As ClassCustoDirFabrPlan, objCustoDirFabrPlanAux As ClassCustoDirFabrPlan
Dim objCustoDirFabrPlanMP As ClassCustoDirFabrPlanMP, objCustoDirFabrPlanMPInt As ClassCustoDirFabrPlanMP
Dim objCustoDirFabrPlanEmb As ClassCustoDirFabrPlanEmb, iAno As Integer, dCustoMP As Double
Dim dAcumCusto As Double, lComando As Long, dTaxaDescPadrao As Double, dTaxaValFut As Double, dDiasValFut As Double

On Error GoTo Erro_Produto_ObterEmbMP

    iAno = Year(gdtDataAtual)
    
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 106890
    
    lErro = CF("CalcMP_ObterValores", iFilialEmpresa, dTaxaDescPadrao, dTaxaValFut, dDiasValFut)
    If lErro <> SUCESSO Then gError 106891
    
    objCustoDirFabr.iAno = iAno
    objCustoDirFabr.iFilialEmpresa = iFilialEmpresa
    objCustoDirFabr.iRotinaOrigem = FORMACAO_PRECO_ANALISE_MARGCONTR
    
    lErro = CF("CustoDiretoFabricacao_Prod_Passo1", objCustoDirFabr, objCustoDirFabrProd, sProduto, FORMACAO_PRECO_QTDECALCPRECO)
    If lErro <> SUCESSO Then gError 106892
    
    iIndice = 0
    
    For Each objCustoDirFabrPlan In objCustoDirFabrProd.colPlan
    
        iIndice = iIndice + 1
        
        'somente considerar os "filhos diretos" do produto principal
        If objCustoDirFabrPlan.sProdutoPai = sProduto Then
        
            'se contém as materias primas do produto principal
            If objCustoDirFabrPlan.sProdutoFilho = sProduto Then
            
                For Each objCustoDirFabrPlanMP In objCustoDirFabrPlan.colMP
                    
                    objCustoDirFabrPlanMP.dParticipacao = objCustoDirFabrPlanMP.dQtde / objCustoDirFabrPlan.dQuantPai
                    
                    'obter o seu custo em custembmp.
                    lErro = CF("CustoDiretoProd_ObtemValor", iFilialEmpresa, objCustoDirFabrPlanMP.sProdutoMP, dTaxaDescPadrao, dTaxaValFut, dDiasValFut, dCustoMP, lComando)
                    If lErro <> SUCESSO Then gError 106893
                    
                    objCustoDirFabrPlanMP.dCustoUnitario = dCustoMP
                    
                    colMP.Add objCustoDirFabrPlanMP
                        
                Next
            
            Else
            
                'criar linha para o produto intermediario incorporando custo das materias primas dele e de seus filhos
                
                Set objCustoDirFabrPlanMPInt = New ClassCustoDirFabrPlanMP
                colMP.Add objCustoDirFabrPlanMPInt
                
                dAcumCusto = 0
                
                With objCustoDirFabrPlanMPInt
                    .dQtde = objCustoDirFabrPlan.dQuantPai * objCustoDirFabrPlan.dParticipacao
                    .iAno = iAno
                    .iFilialEmpresa = iFilialEmpresa
                    .iSeq = 0
                    .sProduto = sProduto
                    .sProdutoMP = objCustoDirFabrPlan.sProdutoFilho
                End With
                
                'obter dados da mps diretamente ligadas ao produto intermediario
                For Each objCustoDirFabrPlanMP In objCustoDirFabrPlan.colMP
                
                    'obter o seu custo em custembmp.
                    lErro = CF("CustoDiretoProd_ObtemValor", iFilialEmpresa, objCustoDirFabrPlanMP.sProdutoMP, dTaxaDescPadrao, dTaxaValFut, dDiasValFut, dCustoMP, lComando)
                    If lErro <> SUCESSO Then gError 106894
                
                    dAcumCusto = dAcumCusto + (dCustoMP * objCustoDirFabrPlanMP.dQtde)
                
                Next
                
                'obter dados das mps ligadas a filhos do produto intermediario
                For iIndice2 = iIndice + 1 To objCustoDirFabrProd.colPlan.Count
                
                    Set objCustoDirFabrPlanAux = objCustoDirFabrProd.colPlan.Item(iIndice2)
                    
                    'se achou uma linha que é filha direta do produto principal é pq já processou todos os filhos do produto intermediario
                    If objCustoDirFabrPlanAux.sProdutoPai = sProduto Then Exit For
                    
                    For Each objCustoDirFabrPlanMP In objCustoDirFabrPlanAux.colMP
                    
                        'obter o seu custo em custembmp.
                        lErro = CF("CustoDiretoProd_ObtemValor", iFilialEmpresa, objCustoDirFabrPlanMP.sProdutoMP, dTaxaDescPadrao, dTaxaValFut, dDiasValFut, dCustoMP, lComando)
                        If lErro <> SUCESSO Then gError 106895
                    
                        dAcumCusto = dAcumCusto + (dCustoMP * objCustoDirFabrPlanMP.dQtde)
                    
                    Next
                    
                Next
                
                objCustoDirFabrPlanMPInt.dParticipacao = objCustoDirFabrPlan.dParticipacao
                objCustoDirFabrPlanMPInt.dCustoUnitario = dAcumCusto / (objCustoDirFabrPlan.dParticipacao * objCustoDirFabrPlan.dQuantPai)
                
            End If
                    
        End If
        
    Next
    
    For Each objCustoDirFabrPlan In objCustoDirFabrProd.colPlan
    
        For Each objCustoDirFabrPlanEmb In objCustoDirFabrPlan.colEmb
        
            objCustoDirFabrPlanEmb.dQtde = objCustoDirFabrPlanEmb.dQtde / FORMACAO_PRECO_QTDECALCPRECO
            
            'obter o seu custo em custembmp.
            lErro = CF("CustoDiretoProd_ObtemValor", iFilialEmpresa, objCustoDirFabrPlanEmb.sProdutoEmb, dTaxaDescPadrao, dTaxaValFut, dDiasValFut, dCustoMP, lComando)
            If lErro <> SUCESSO Then gError 106898
            
            objCustoDirFabrPlanEmb.dCustoUnitario = dCustoMP
            
            colEmb.Add objCustoDirFabrPlanEmb
    
        Next
        
    Next
    
    Call Comando_Fechar(lComando)
    
    Produto_ObterEmbMP = SUCESSO
     
    Exit Function
    
Erro_Produto_ObterEmbMP:

    Produto_ObterEmbMP = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150520)
     
    End Select
     
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function FPreco_ObtemFreteFilCli(ByVal objMnemonicoValor As ClassMnemonicoValor, ByVal objExeExp As ClassExeExp) As Long
'retorna o valor unitario do frete associado à filial do cliente

Dim lErro As Long, lCliente As Long, iFilial As Integer, iFilialFaturamento As Integer
Dim lComando As Long, dPreco As Double

On Error GoTo Erro_FPreco_ObtemFreteFilCli

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 106916

    lCliente = StrParaLong(objMnemonicoValor.vParam(1))
    iFilial = StrParaInt(objMnemonicoValor.vParam(2))
    iFilialFaturamento = objExeExp.objContexto.iFilialFaturamento
    
    If lCliente <> 0 And iFilial <> 0 And iFilialFaturamento <> 0 Then
    
        lErro = Comando_Executar(lComando, "SELECT Preco FROM DVVCliente, TipoFreteFP WHERE DVVCliente.FilialEmpresa = ? AND DVVCliente.CodCliente = ? AND DVVCliente.CodFilial = ? AND DVVCliente.TipoFrete = TipoFreteFP.Codigo", _
            dPreco, iFilialFaturamento, lCliente, iFilial)
        If lErro <> AD_SQL_SUCESSO Then gError 106917
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106918
        If lErro <> AD_SQL_SUCESSO And objExeExp.objContexto.iRotinaOrigem <> FORMACAO_PRECO_ANALISE_MARGCONTR_REL Then gError 106919
            
    End If
    
    Call Comando_Fechar(lComando)
    
    objMnemonicoValor.colValor.Add CDbl(dPreco)
    
    FPreco_ObtemFreteFilCli = SUCESSO
     
    Exit Function
    
Erro_FPreco_ObtemFreteFilCli:

    FPreco_ObtemFreteFilCli = gErr
     
    Select Case gErr
          
        Case 106916
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 106917, 106918
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_VALOR_FRETE_FILCLI", gErr)
        
        Case 106919
            '??? Call Rotina_Aviso(vbOKOnly, "ERRO_SEM_VALOR_FRETE_FILCLI", gErr)
            Resume Next
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150521)
     
    End Select
     
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function FPreco_ConvQtdeOrcCusto(ByVal sProduto As String, dQtde As Double) As Long
'converte qtde orcada p/venda da unidade de medida de venda para a de calculo de custo p/formacao de precos

Dim lErro As Long, objProduto As New ClassProduto, dFator As Double

On Error GoTo Erro_FPreco_ConvQtdeOrcCusto

    objProduto.sCodigo = sProduto
    
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 124123
    If lErro <> SUCESSO Then gError 124124
    
    If objProduto.sSiglaUMVenda <> objProduto.sSiglaUMEstoque Then
    
        lErro = CF("UM_Conversao", objProduto.iClasseUM, objProduto.sSiglaUMEstoque, objProduto.sSiglaUMVenda, dFator)
        If lErro <> SUCESSO Then gError 124125
        
        dQtde = Round(dQtde * dFator, 4)
    
    End If
    
    FPreco_ConvQtdeOrcCusto = SUCESSO
     
    Exit Function
    
Erro_FPreco_ConvQtdeOrcCusto:

    FPreco_ConvQtdeOrcCusto = gErr
     
    Select Case gErr
          
        Case 124123, 124125
        
        Case 124124
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_ENCONTRADO", gErr, sProduto)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150522)
     
    End Select
     
    Exit Function

End Function

Function CustoProd_AjustarPerda(ByVal sProduto As String, ByVal sVersao As String, ByVal iFilialEmpresa As Integer, ByVal iRotinaOrigem As Integer, dPercentualPerda As Double) As Long
'criada para atender customizacao da INPAL
    CustoProd_AjustarPerda = SUCESSO
End Function

Private Function CustoDirFabrProdInf_Obter(ByVal iFilialEmpresa As Integer, ByVal iAno As Integer, ByVal sProduto As String, dCustoInformado As Double) As Long

Dim lErro As Long, lComando As Long

On Error GoTo Erro_CustoDirFabrProdInf_Obter

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 124285
    
    lErro = Comando_Executar(lComando, "SELECT Custo FROM CustoDirFabrProdInf WHERE FilialEmpresa = ? AND Ano = ? AND Produto = ?", dCustoInformado, iFilialEmpresa, iAno, sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 124286
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 124287
    
    If lErro <> AD_SQL_SUCESSO Then gError 124288
    
    Call Comando_Fechar(lComando)
    
    CustoDirFabrProdInf_Obter = SUCESSO
     
    Exit Function
    
Erro_CustoDirFabrProdInf_Obter:

    CustoDirFabrProdInf_Obter = gErr
     
    Select Case gErr
          
        Case 124285
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 124286, 124287
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CUSTODIRFABRPRODINF", Err)
        
        Case 124288
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150523)
     
    End Select
     
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function
        
Function PrecoUltNF(ByVal objMnemonicoValor As ClassMnemonicoValor, ByVal objExeExp As ClassExeExp) As Long
'parametros:Cliente, Filial, Produto
'retorna o último preco unitario utilizado para o produto numa nf de venda (que nao seja complemento)

Dim lErro As Long, lCliente As Long, iFilial As Integer, sProduto As String
Dim dPrecoUnitario As Double, objItemNF As New ClassItemNF

On Error GoTo Erro_PrecoUltNF

    lCliente = StrParaLong(objMnemonicoValor.vParam(1))
    iFilial = StrParaInt(objMnemonicoValor.vParam(2))
    sProduto = objMnemonicoValor.vParam(3)

    If lCliente <> 0 And iFilial <> 0 And sProduto <> "" And objExeExp.objContexto.iFilialFaturamento <> 0 Then

        objItemNF.sProduto = sProduto
        lErro = CF("ClienteFilial_Le_UltimoItemNFVenda", objItemNF, objExeExp.objContexto.iFilialFaturamento, lCliente, iFilial)
        If lErro = SUCESSO Then
        
            lErro = CF("Produto_ConvPrecoUMAnalise", sProduto, objItemNF.sUnidadeMed, objItemNF.dPrecoUnitario, dPrecoUnitario)
            If lErro <> SUCESSO Then gError 130001
            
        End If
        
    End If

    objMnemonicoValor.colValor.Add dPrecoUnitario
    
    PrecoUltNF = SUCESSO

    Exit Function

Erro_PrecoUltNF:

    PrecoUltNF = gErr

    Select Case gErr

        Case 130001

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150524)

    End Select

    Exit Function

End Function

Function Produto_ConvPrecoUMAnalise(ByVal sProduto As String, ByVal sUMOrig As String, ByVal dPrecoUnitarioOrig As Double, dPrecoUnitarioConv As Double) As Long
'converte um preco na unidade sUM no preco na unidade de analise de margem de conribuicao

Dim lErro As Long, sUMConv As String, objProduto As New ClassProduto, dFator As Double

On Error GoTo Erro_Produto_ConvPrecoUMAnalise

    objProduto.sCodigo = sProduto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 124304
    
    sUMConv = objProduto.sSiglaUMVenda
    
    'converter o custo p/a unidade necessaria
    If sUMConv <> sUMOrig Then
    
        lErro = CF("UM_Conversao", objProduto.iClasseUM, sUMConv, sUMOrig, dFator)
        If lErro <> SUCESSO Then gError 124305
        
        dPrecoUnitarioConv = dPrecoUnitarioOrig * dFator
        
    Else
        
        dPrecoUnitarioConv = dPrecoUnitarioOrig
    
    End If
    
    Produto_ConvPrecoUMAnalise = SUCESSO
     
    Exit Function
    
Erro_Produto_ConvPrecoUMAnalise:

    Produto_ConvPrecoUMAnalise = gErr
     
    Select Case gErr
          
        Case 124304, 124305
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150525)
     
    End Select
     
    Exit Function

End Function

Function ClienteFilial_Le_UltimoItemNFVenda(ByVal objItemNF As ClassItemNF, ByVal iFilialEmpresa As Integer, ByVal lCliente As Long, ByVal iFilialCli As Integer) As Long

Dim lErro As Long, lComando As Long, sUnidadeMed As String, dPrecoUnitario As Double

On Error GoTo Erro_ClienteFilial_Le_UltimoItemNFVenda

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 130003
    
    sUnidadeMed = String(STRING_UM_SIGLA, 0)

    lErro = Comando_Executar(lComando, "SELECT UnidadeMed, PrecoUnitario FROM NFiscal, ItensNFiscal, TiposDocInfo WHERE NFiscal.FilialEmpresa = ? AND NFiscal.Cliente = ? AND NFiscal.FilialCli = ? AND ItensNFiscal.Produto = ? AND NFiscal.NumIntDoc = ItensNFiscal.NumIntNF AND NFiscal.Status <> ? AND TiposDocInfo.Codigo = NFiscal.TipoNFiscal AND TiposDocInfo.Complementar = 0 AND TiposDocInfo.Tipo = 2 AND TiposDocInfo.Destinatario = 1 AND TiposDocInfo.NomeTelaNfiscal <> 'nfiscaldev' ORDER BY NFiscal.DataEmissao DESC, NFiscal.NumIntDoc DESC", _
        sUnidadeMed, dPrecoUnitario, iFilialEmpresa, lCliente, iFilialCli, objItemNF.sProduto, STATUS_CANCELADO)
    If lErro <> AD_SQL_SUCESSO Then gError 130004
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130005
    If lErro <> AD_SQL_SUCESSO Then gError 130006
    
    With objItemNF
        .sUnidadeMed = sUnidadeMed
        .dPrecoUnitario = dPrecoUnitario
    End With
    
    Call Comando_Fechar(lComando)

    ClienteFilial_Le_UltimoItemNFVenda = SUCESSO
     
    Exit Function
    
Erro_ClienteFilial_Le_UltimoItemNFVenda:

    ClienteFilial_Le_UltimoItemNFVenda = gErr
     
    Select Case gErr
          
        Case 130003
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 130004, 130005
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRECOULTNF", Err)
        
        Case 130006 'nao encontrado
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150526)
     
    End Select
     
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'######################################################################3
'Inserido Por Wagner 16/11/2004
Function CustoDiretoRateio(ByVal objCustoDirFabr As ClassCustoDirFabr, dCustoDireto As Double) As Long
'parametros:Produto
'retorna o custo direto (por unidade de ...) associado ao produto pelo rateio dos custos diretos

Dim lErro As Long, lComando As Long, sProduto As String, dQtdePrevTotal As Double, lComando1 As Long
Dim adQuant(1 To KIT_NUM_FATORES) As Double, adCusto(0 To KIT_NUM_FATORES) As Double, adQuantProd(1 To KIT_NUM_FATORES) As Double
Dim dQuantPrevista As Double, iIndice As Integer, dCustoOutros As Double

On Error GoTo Erro_CustoDiretoRateio

    sProduto = objCustoDirFabr.sProduto

    If Len(Trim(sProduto)) <> 0 Then
       
        lComando = Comando_Abrir()
        If lComando = 0 Then gError 131257
    
        lErro = Comando_Executar(lComando, "SELECT QuantPrevista, CustoTotal, CustoDirFabr.QuantFator1, CustoFator1, CustoDirFabr.QuantFator2, CustoFator2, CustoDirFabr.QuantFator3, CustoFator3, CustoDirFabr.QuantFator4, CustoFator4, CustoDirFabr.QuantFator5, CustoFator5, CustoDirFabr.QuantFator6, CustoFator6, CustoDirFabrProd.QuantFator1, CustoDirFabrProd.QuantFator2, CustoDirFabrProd.QuantFator3, CustoDirFabrProd.QuantFator4, CustoDirFabrProd.QuantFator5, CustoDirFabrProd.QuantFator6 FROM CustoDirFabr, CustoDirFabrProd WHERE CustoDirFabr.FilialEmpresa = CustoDirFabrProd.FilialEmpresa AND CustoDirFabr.Ano = CustoDirFabrProd.Ano AND CustoDirFabr.FilialEmpresa = ? AND CustoDirFabr.Ano = ? AND CustoDirFabrProd.Produto = ?", _
            dQuantPrevista, adCusto(0), adQuant(1), adCusto(1), adQuant(2), adCusto(2), adQuant(3), adCusto(3), adQuant(4), adCusto(4), adQuant(5), adCusto(5), adQuant(6), adCusto(6), adQuantProd(1), adQuantProd(2), adQuantProd(3), adQuantProd(4), adQuantProd(5), adQuantProd(6), objCustoDirFabr.iFilialEmpresa, objCustoDirFabr.iAno, sProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 131258
    
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 131259
    
        If lErro = AD_SQL_SUCESSO And dQuantPrevista <> 0 Then

            dCustoOutros = adCusto(0)
            For iIndice = 1 To KIT_NUM_FATORES
                If adQuant(iIndice) <> 0 Then
                    dCustoDireto = dCustoDireto + (adQuantProd(iIndice) * adCusto(iIndice) / adQuant(iIndice))
                    dCustoOutros = dCustoOutros - adCusto(iIndice)
                End If
            Next

            If Abs(adCusto(0) - dCustoOutros) > DELTA_VALORMONETARIO Then
    
                dCustoDireto = dCustoDireto + (dCustoDireto * dCustoOutros / (adCusto(0) - dCustoOutros))
                dCustoDireto = dCustoDireto / dQuantPrevista

            Else

                lComando1 = Comando_Abrir()
                If lComando1 = 0 Then gError 131260

                lErro = Comando_Executar(lComando1, "SELECT SUM(QuantPrevista) FROM CustoDirFabrProd WHERE FilialEmpresa = ? AND Ano = ?", _
                    dQtdePrevTotal, objCustoDirFabr.iFilialEmpresa, objCustoDirFabr.iAno)
                If lErro <> AD_SQL_SUCESSO Then gError 131261
        
                lErro = Comando_BuscarProximo(lComando1)
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 131262

                Call Comando_Fechar(lComando1)
                    
                dCustoDireto = dCustoOutros / dQtdePrevTotal
                
            End If

        End If

        Call Comando_Fechar(lComando)

    End If

    CustoDiretoRateio = SUCESSO

    Exit Function

Erro_CustoDiretoRateio:

    CustoDiretoRateio = gErr

    Select Case gErr

        Case 131257, 131260
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 131258, 131259, 131261, 131262
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CUSTODIRETO", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150527)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function
'######################################################################3
Function ContextoPlan_Inicializa(ByVal objContexto As ClassContextoPlan, ByVal iFilialEmpresa As Integer, ByVal sProduto As String, ByVal iTabelaPreco As Integer, ByVal dtDataReferencia As Date, ByVal dQuantidade As Double) As Long

Dim lErro As Long, dtDataCF As Date

On Error GoTo Erro_ContextoPlan_Inicializa

    objContexto.iFilialFaturamento = iFilialEmpresa
    objContexto.sProduto = sProduto
    objContexto.iTabelaPreco = iTabelaPreco
    objContexto.iRotinaOrigem = FORMACAO_PRECO_ROTINA_CALCPRECO
    
    If dQuantidade = 0 Then
        objContexto.dQuantidade = FORMACAO_PRECO_QTDECALCPRECO
    Else
        objContexto.dQuantidade = dQuantidade
    End If
    
    If dtDataReferencia = DATA_NULA Then
        objContexto.iAno = Year(gdtDataAtual)
    Else
        objContexto.iAno = Year(dtDataReferencia)
    End If
    
    lErro = CF("FilialEmpresa_ObtemDataCustoFixo", objContexto.iFilialFaturamento, dtDataCF)
    If lErro <> SUCESSO Then gError 106859
    
    objContexto.dtDataCustoFixo = dtDataCF
    
    lErro = CF("ContextoPlan_InicializaCusto", objContexto)
    If lErro <> SUCESSO Then gError 106829
    
    ContextoPlan_Inicializa = SUCESSO
     
    Exit Function
    
Erro_ContextoPlan_Inicializa:

    ContextoPlan_Inicializa = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150528)
     
    End Select
     
    Exit Function

End Function

Function FormPreco_IPIAliquota(ByVal objMnemonicoValor As ClassMnemonicoValor, ByVal objExeExp As ClassExeExp) As Long

Dim lErro As Long, lComando As Long, dAliquota As Double, sProduto As String

On Error GoTo Erro_FormPreco_IPIAliquota

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 130409

    sProduto = objExeExp.objContexto.sProduto
    If sProduto = "" Then sProduto = objMnemonicoValor.vParam(1)

    lErro = Comando_Executar(lComando, "SELECT IPIAliquota FROM Produtos WHERE Codigo = ?", dAliquota, sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 130410

    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130411

    If lErro <> AD_SQL_SUCESSO Then
        dAliquota = 0
    Else
        
        'para ser tratada como %
        dAliquota = dAliquota * 100
    
    End If
    
    Call Comando_Fechar(lComando)

    objMnemonicoValor.colValor.Add dAliquota

    FormPreco_IPIAliquota = SUCESSO
     
    Exit Function
    
Erro_FormPreco_IPIAliquota:

    FormPreco_IPIAliquota = gErr
     
    Select Case gErr
          
        Case 130409
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150529)
     
    End Select
     
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function PrecoTabela_Param(ByVal objMnemonicoValor As ClassMnemonicoValor, ByVal objExeExp As ClassExeExp) As Long
'retorna o preco de tabela de um produto para a filial de faturamento

Dim lErro As Long
Dim objTabelaPrecoItem As New ClassTabelaPrecoItem
Dim dPrecoUnitario As Double
Dim iPreenchido As Integer, sProduto As String

On Error GoTo Erro_PrecoTabela_Param

    lErro = CF("Produto_Formata", objMnemonicoValor.vParam(1), sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError 201036
    
    objTabelaPrecoItem.iCodTabela = objMnemonicoValor.vParam(2)
    objTabelaPrecoItem.iFilialEmpresa = objExeExp.objContexto.iFilialFaturamento
    
    objTabelaPrecoItem.sCodProduto = objExeExp.objContexto.sProduto
    
    lErro = CF("TabelaPrecoItem_Le", objTabelaPrecoItem)
    If lErro <> SUCESSO And lErro <> 28014 Then gError 106861
    
    If lErro = 28014 Then
        objTabelaPrecoItem.iFilialEmpresa = EMPRESA_TODA
    
        lErro = CF("TabelaPrecoItem_Le", objTabelaPrecoItem)
        If lErro <> SUCESSO And lErro <> 28014 Then gError 106862
     End If

    If lErro <> 28014 Then dPrecoUnitario = objTabelaPrecoItem.dPreco
    
    objMnemonicoValor.colValor.Add CDbl(dPrecoUnitario)
    
    PrecoTabela_Param = SUCESSO
     
    Exit Function
    
Erro_PrecoTabela_Param:

    PrecoTabela_Param = gErr
     
    Select Case gErr
          
        Case 106861, 106862, 201036
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150513)
     
    End Select
     
    Exit Function

End Function

Private Function ComissoesCliProdAux(dPerc As Double, ByVal objExeExp As ClassExeExp, ByVal bDireta As Boolean) As Long

Dim lErro As Long, objVendedor As New ClassVendedor, lMatricula As Long
Dim colComissoes As New Collection, dPrecoUnitario As Double
Dim objComissao As ClassComissaoPedVendas, dAcumValComiss As Double

On Error GoTo Erro_ComissoesCliProdAux

    Set colComissoes = objExeExp.objContexto.colComissoes
    dPrecoUnitario = objExeExp.objContexto.dPrecoPraticado
            
    'percorre a colecao de comissoes obtendo o % proprio
    For Each objComissao In colComissoes
    
        If (bDireta And objComissao.iIndireta = 0) Or (Not bDireta And objComissao.iIndireta <> 0) Then
            
            objVendedor.iCodigo = objComissao.iCodVendedor
            lErro = CF("Vendedor_Le", objVendedor)
            If lErro <> SUCESSO And lErro <> 12582 Then gError 130346
            If lErro <> SUCESSO Then gError 130347
            
            If IsNumeric(objVendedor.sMatricula) Then
                lMatricula = StrParaLong(objVendedor.sMatricula)
            End If
            
            If objVendedor.iVinculo = VENDEDOR_VINCULO_EMPREGADO And lMatricula >= 1261 Then
                dAcumValComiss = dAcumValComiss + (objComissao.dValor / 0.83)
            Else
                dAcumValComiss = dAcumValComiss + objComissao.dValor
            End If
            
        End If
        
    Next
    
    If (dPrecoUnitario * objExeExp.objContexto.dQuantidade) <> 0 Then
        dPerc = Arredonda_Moeda(dAcumValComiss / (dPrecoUnitario * objExeExp.objContexto.dQuantidade), 4)
    Else
        dPerc = 0
    End If

    ComissoesCliProdAux = SUCESSO
     
    Exit Function
    
Erro_ComissoesCliProdAux:

    ComissoesCliProdAux = gErr
     
    Select Case gErr
          
        Case 130346
        
        Case 130347
            Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_ENCONTRADO", gErr, objVendedor.iCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    Exit Function

End Function

Function ICMSAliqFilCliProd(ByVal objMnemonicoValor As ClassMnemonicoValor, ByVal objExeExp As ClassExeExp) As Long
'retorna a aliquota de icms associada a filial do cliente para o produto

Dim lErro As Long, sUFCli As String, sInscricaoEstadual As String, iPais As Integer
Dim lComando As Long, lCliente As Long, iFilial As Integer, objEstado As New ClassEstado, objNFiscal As New ClassNFiscal
Dim iFilialFaturamento As Integer, dAliquota As Double, objFilialEmpresa As New AdmFiliais, sProduto As String

On Error GoTo Erro_ICMSAliqFilCliProd

    dAliquota = 0
    
    lCliente = StrParaLong(objMnemonicoValor.vParam(1))
    If lCliente = 0 Then lCliente = objExeExp.objContexto.lCliente
    iFilial = StrParaInt(objMnemonicoValor.vParam(2))
    If iFilial = 0 Then iFilial = objExeExp.objContexto.iFilialCli
    iFilialFaturamento = objExeExp.objContexto.iFilialFaturamento
    sProduto = objExeExp.objContexto.sProduto
    
    If lCliente <> 0 And iFilial <> 0 And iFilialFaturamento <> 0 And sProduto <> "" Then
    
        lErro = ProdutoTribAutomatica(iFilialFaturamento, sProduto, lCliente, iFilial, objNFiscal)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        If objNFiscal.colItensNF(1).objTributacaoItemNF.dICMSInterestAliqUFDest <> 0 Then
            dAliquota = objNFiscal.colItensNF(1).objTributacaoItemNF.dICMSInterestAliqUFDest + objNFiscal.colItensNF(1).objTributacaoItemNF.dICMSInterestPercFCPUFDest
        Else
            dAliquota = objNFiscal.colItensNF(1).objTributacaoItemNF.dICMSAliquota
        End If
    
    End If
    
    objMnemonicoValor.colValor.Add CDbl(dAliquota)
    
    ICMSAliqFilCliProd = SUCESSO
     
    Exit Function
    
Erro_ICMSAliqFilCliProd:

    ICMSAliqFilCliProd = gErr
     
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150510)
     
    End Select
     
    Exit Function

End Function

Private Function ProdutoTribAutomatica(ByVal iFilialEmpresa As Integer, ByVal sProduto As String, ByVal lCliente As Long, iFilialCli As Integer, objNFiscal As ClassNFiscal) As Long

Dim lErro As Long
Dim iTotalItens As Integer
Dim iTotal As Integer
Dim objItemNF As ClassItemNF
Dim dValorTotal As Double, objTribTab As ClassTribTab
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim objProduto As ClassProduto

On Error GoTo Erro_ProdutoTribAutomatica

    Call objNFiscal.Inicializa_Tributacao
    Call objNFiscal.objTributacaoNF.Coloca_Auto
    
    objNFiscal.iCondicaoPagto = 0
    objNFiscal.iFilialCli = iFilialCli
    objNFiscal.lCliente = lCliente
    objNFiscal.dtDataReferencia = gdtDataAtual
    objNFiscal.dtDataEmissao = gdtDataAtual
    objNFiscal.dtDataSaida = gdtDataAtual
    objNFiscal.dtDataEntrada = DATA_NULA
    objNFiscal.dtDataVencimento = DATA_NULA
    objNFiscal.iFilialEmpresa = iFilialEmpresa
    objNFiscal.iFilialEntrega = iFilialCli
    
    objNFiscal.iTipoDocInfo = DOCINFO_NFISFV
    objNFiscal.iTipoNFiscal = DOCINFO_NFISFV
    objNFiscal.iStatus = STATUS_LANCADO
    objNFiscal.iRecibo = 0
    
    Set objItemNF = New ClassItemNF
    Call objItemNF.Inicializa_Tributacao
    Call objItemNF.objTributacaoItemNF.Coloca_Auto
    
    objItemNF.dCusto = 0
    objItemNF.dQuantidade = 10
    objItemNF.dPrecoUnitario = 100
    objItemNF.dValorTotal = 1000
    objItemNF.sProduto = sProduto
    objItemNF.iItem = 1

    objNFiscal.colItensNF.Add1 objItemNF
    
    Set objProduto = New ClassProduto
    objProduto.sCodigo = objItemNF.sProduto
    
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
    
    objItemNF.sUnidadeMed = objProduto.sSiglaUMVenda
    
    objItemNF.objTributacaoItemNF.sProduto = objProduto.sCodigo
    objItemNF.objTributacaoItemNF.iExTIPI = objProduto.iExTIPI
    objItemNF.objTributacaoItemNF.sGenero = objProduto.sGenero
    objItemNF.objTributacaoItemNF.iProdutoEspecifico = objProduto.iProdutoEspecifico
    objItemNF.objTributacaoItemNF.sUMTrib = objProduto.sSiglaUMTrib
    objItemNF.objTributacaoItemNF.sIPICodProduto = objProduto.sIPICodigo

    objItemNF.objTributacaoItemNF.dQtdTrib = objItemNF.dQuantidade
    objItemNF.objTributacaoItemNF.dValorUnitTrib = objItemNF.dPrecoUnitario
    objItemNF.objTributacaoItemNF.dPrecoTotal = objItemNF.dValorTotal
    objItemNF.objTributacaoItemNF.sProdutoDescricao = objProduto.sDescricao
    objItemNF.objTributacaoItemNF.sISSQN = objProduto.sISSQN
    
    objNFiscal.dValorProdutos = objItemNF.dValorTotal
    
    Set objTribTab = New ClassTribTab
    
    lErro = objTribTab.TributacaoNF_Reset(objNFiscal)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Call objTribTab.TipoNFiscal_Definir(DOCINFO_NFISFV, "NFISFV")
    Call objTribTab.DataEmissao_Alterada(gdtDataAtual)
    Call objTribTab.Serie_Alterada(objNFiscal.sSerie)
    
    lErro = gobjTributacao.AtualizaImpostos(objTribTab.mvarobjNFTributacao, 1)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = objTribTab.PreencherNF(objNFiscal)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Set objTribTab = Nothing
    
    objNFiscal.sNaturezaOp = objNFiscal.objTributacaoNF.sNaturezaOpInterna
    
    'acertar valor total e da parcela, considerando os tributos, retencóes,...
    objNFiscal.dValorTotal = dValorTotal + objNFiscal.objTributacaoNF.dICMSSubstValor + objNFiscal.objTributacaoNF.dIPIValor + objNFiscal.dValorFrete + objNFiscal.dValorSeguro + objNFiscal.dValorOutrasDespesas + IIf(objNFiscal.objTributacaoNF.iISSIncluso = 0, objNFiscal.objTributacaoNF.dISSValor, 0)
    
    ProdutoTribAutomatica = SUCESSO
    
    Exit Function
    
Erro_ProdutoTribAutomatica:

    ProdutoTribAutomatica = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201557)

    End Select
    
    Exit Function

End Function

Public Function PrecoUnitario_CalculaFormPreco(ByVal iFilialFaturamento As Integer, ByVal lCliente As Long, ByVal iFilialCli As Integer, ByVal sProduto As String, ByVal sUM As String, ByVal iTabelaPreco As Integer, dPrecoUnitario As Double) As Long
'calcula o preco unitario de venda segundo a planilha de formacao de preco

Dim lErro As Long
Dim objContexto As New ClassContextoPlan, dtDataCF As Date
Dim objFormacaoPreco As New ClassFormacaoPreco
Dim objFormacaoPrecoTemp As ClassFormacaoPreco
Dim colFormacaoPreco As Collection, dValor As Double, dFator As Double, objProduto As New ClassProduto

On Error GoTo Erro_PrecoUnitario_CalculaFormPreco

    objContexto.iFilialFaturamento = iFilialFaturamento
    objContexto.dQuantidade = FORMACAO_PRECO_QTDECALCPRECO
    objContexto.iAno = Year(gdtDataAtual)
    objContexto.lCliente = lCliente
    objContexto.iFilialCli = iFilialCli
    objContexto.sProduto = sProduto
    objContexto.iTabelaPreco = iTabelaPreco
    objContexto.iRotinaOrigem = FORMACAO_PRECO_ROTINA_CALCPRECO

    lErro = CF("FilialEmpresa_ObtemDataCustoFixo", objContexto.iFilialFaturamento, dtDataCF)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    objContexto.dtDataCustoFixo = dtDataCF
    
    lErro = CF("ContextoPlan_InicializaCusto", objContexto)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    objFormacaoPreco.iFilialEmpresa = iFilialFaturamento
    objFormacaoPreco.iTabelaPreco = iTabelaPreco
    objFormacaoPreco.sProduto = sProduto
    objFormacaoPreco.iEscopo = FORMACAO_PRECO_ESCOPO_TABPRECO

    Set colFormacaoPreco = New Collection

    'Lê a planilha de Formacao de Preço da Tabela/Produto
    lErro = CF("FormacaoPreco_Le1", objFormacaoPreco, colFormacaoPreco)
    If lErro <> SUCESSO And lErro <> 92434 And lErro <> 92432 Then gError 106812

    For Each objFormacaoPrecoTemp In colFormacaoPreco
    
        With objFormacaoPrecoTemp
            .iFilialEmpresa = iFilialFaturamento
            .iEscopo = FORMACAO_PRECO_ESCOPO_TABPRECO
            .iTabelaPreco = iTabelaPreco
            .sProduto = sProduto
        End With
    
    Next
    
    'Executa as formulas da planilha de preço. Retorna o valor da planilha em dValor (que é o valor da última linha da planilha) e o valor de cada linha em colFormacaoPreco.Item(?).dValor
    lErro = CF("Avalia_Expressao_FPreco1", colFormacaoPreco, dValor, sProduto, objContexto)
    If lErro <> SUCESSO Then gError 106818

    objProduto.sCodigo = sProduto
    'Lê o produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 84029 '26638
    If lErro = 28030 Then gError 84030 '26639
    
    If sUM <> objProduto.sSiglaUMVenda Then
    
        'Converte a quantidade para a UM de Venda
        lErro = CF("UM_Conversao", objProduto.iClasseUM, sUM, objProduto.sSiglaUMVenda, dFator)
        If lErro <> SUCESSO Then gError 84031 '26640
    
        dValor = dValor / dFator
        
    End If
    
    dPrecoUnitario = StrParaDbl(Format(dValor, gobjFAT.sFormatoPrecoUnitario))

    PrecoUnitario_CalculaFormPreco = SUCESSO
    
    Exit Function
    
Erro_PrecoUnitario_CalculaFormPreco:

    PrecoUnitario_CalculaFormPreco = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201578)

    End Select
    
    Exit Function

End Function


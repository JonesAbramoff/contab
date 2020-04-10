Attribute VB_Name = "GeraContratoCobranca"
Option Explicit

Private Function Rotina_GeraContrato_AtualizaTelaBatch()
'Atualiza tela de acompanhamento do Batch

Dim lErro As Long
Dim vbMsgBox As VbMsgBoxResult

On Error GoTo Erro_Rotina_GeraContrato_AtualizaTelaBatch

    'Atualiza tela de acompanhamento do Batch
    lErro = DoEvents()

    TelaAcompanhaBatchEST.dValorAtual = TelaAcompanhaBatchEST.dValorAtual + 1
    TelaAcompanhaBatchEST.TotReg.Caption = CStr(TelaAcompanhaBatchEST.dValorAtual)
    TelaAcompanhaBatchEST.ProgressBar1.Value = CInt((TelaAcompanhaBatchEST.dValorAtual / TelaAcompanhaBatchEST.dValorTotal) * 100)

    If TelaAcompanhaBatchEST.iCancelaBatch = CANCELA_BATCH Then

        vbMsgBox = Rotina_Aviso(vbYesNo, "AVISO_CANCELAR_GERACAOCONTRATO")

        If vbMsgBox = vbYes Then gError 129933

        TelaAcompanhaBatchEST.iCancelaBatch = 0

    End If
    
    Rotina_GeraContrato_AtualizaTelaBatch = SUCESSO
    
    Exit Function

Erro_Rotina_GeraContrato_AtualizaTelaBatch:

    Rotina_GeraContrato_AtualizaTelaBatch = gErr

    Select Case gErr

        Case 129933

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161553)

    End Select
       
    Exit Function

End Function

Public Function NFiscalContrato_Gera(objGeracaoFatContrato As ClassGeracaoFatContrato) As Long
'Gera as Notas Fiscais em cima dos itens de contratos/medição que vão ser faturados

Dim lErro As Long
Dim iTotalItens As Integer
Dim iTotal As Integer
Dim objNFiscal As ClassNFiscal, iTemp As Integer
Dim objItemNF As ClassItemNF
Dim objContabil As ClassContabil
Dim objContratoFat As ClassContratoFat
Dim objContratoFatItens As ClassContratoFatItens
Dim dValorTotal As Double, objTribTab As ClassTribTab
Dim objTipoDocInfo As New ClassTipoDocInfo, dValorLiquido As Double
Dim iItem As Integer
Dim objItemMedicao As ClassItensMedCtr
Dim colContFatItensAgrupado As New Collection
Dim objContFatItensAux1 As ClassContratoFatItens
Dim objContFatItensAux2 As ClassContratoFatItens
Dim iIndice1 As Integer
Dim iIndice2 As Integer
Dim dQuantidade As Double
Dim dCusto As Double
Dim dValor As Double
Dim bAchou As Boolean
Dim colMedicoes As Collection
Dim colcolMedicoes As New Collection
Dim colNumIntNF As New Collection
Dim objProduto As ClassProduto, dMultaPadrao As Double, dJurosPadrao As Double, bChederNaoDoador As Boolean

On Error GoTo Erro_NFiscalContrato_Gera

    'Lê o Tipo de Documento
    objTipoDocInfo.iCodigo = objGeracaoFatContrato.iTipoNFiscal
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 130023

    'se não estiver cadastrado ==> erro
    If lErro = 31415 Then gError 130024
    
    lErro = CF("NFiscalContrato_Le", objGeracaoFatContrato, iTotalItens, iTotal)
    If lErro <> AD_SQL_SUCESSO Then gError 129825

    'Inicializa barra de progressão
    TelaAcompanhaBatchEST.dValorTotal = iTotal
    TelaAcompanhaBatchEST.dValorAtual = 0
      
    lErro = CF("CobrancaContrato_Grava", objGeracaoFatContrato)
    If lErro <> SUCESSO Then gError 132974
   
    For Each objContratoFat In objGeracaoFatContrato.colItens
                      
        Set colcolMedicoes = New Collection
        Set colContFatItensAgrupado = New Collection
                      
        Set objNFiscal = New ClassNFiscal
                
        Call objNFiscal.Inicializa_Tributacao
        Call objNFiscal.objTributacaoNF.Coloca_Auto
        
        objNFiscal.iCondicaoPagto = objContratoFat.iCondPagto
        objNFiscal.iFilialCli = objContratoFat.iFilCli
        objNFiscal.lCliente = objContratoFat.lCliente
        objNFiscal.dtDataReferencia = objGeracaoFatContrato.dtDataRefVencimento
        objNFiscal.dtDataEmissao = objGeracaoFatContrato.dtDataEmissao
        objNFiscal.dtDataSaida = objGeracaoFatContrato.dtDataEmissao
        objNFiscal.dtDataEntrada = DATA_NULA
        objNFiscal.dtDataVencimento = DATA_NULA
        objNFiscal.iFilialEmpresa = objGeracaoFatContrato.iFilialEmpresa
        objNFiscal.iFilialEntrega = objContratoFat.iFilCli 'Alterado por Wagner
        objNFiscal.sSerie = objContratoFat.sSerie
        objNFiscal.sNaturezaOp = objContratoFat.sNaturezaOp
        objNFiscal.iTipoDocInfo = objGeracaoFatContrato.iTipoNFiscal
        objNFiscal.iTipoNFiscal = objGeracaoFatContrato.iTipoNFiscal
        objNFiscal.iStatus = STATUS_LANCADO
        objNFiscal.iRecibo = objContratoFat.iRecibo
        
        If Len(Trim(objContratoFat.sNaturezaOp)) <> 0 Then
            objNFiscal.sNaturezaOp = objContratoFat.sNaturezaOp
            objNFiscal.objTributacaoNF.sNaturezaOpInterna = objContratoFat.sNaturezaOp
            objNFiscal.objTributacaoNF.iNaturezaOpManual = VAR_PREENCH_MANUAL
        End If
        
        If objContratoFat.iTipoTributacao <> 0 Then
            objNFiscal.objTributacaoNF.iTipoTributacao = objContratoFat.iTipoTributacao
            objNFiscal.objTributacaoNF.iTipoTributacaoManual = VAR_PREENCH_MANUAL
        End If
                
        iItem = 0
        iIndice1 = 0
        
        'Agrupa itens
        For Each objContFatItensAux1 In objContratoFat.colItens
        
            iIndice1 = iIndice1 + 1
            iIndice2 = 0
            dQuantidade = 0
            dValor = 0
            dCusto = 0
            
            bAchou = False
            
            Set colMedicoes = New Collection

            For Each objContFatItensAux2 In colContFatItensAgrupado
                If objContFatItensAux1.sProduto = objContFatItensAux2.sProduto Then
                    bAchou = True
                    Exit For
                End If
            Next
            
            If Not bAchou Then
            
                'Soma os itens iguais
                For Each objContFatItensAux2 In objContratoFat.colItens
                    
                    iIndice2 = iIndice2 + 1
                
                    If iIndice1 <> iIndice2 Then
                        If objContFatItensAux1.sProduto = objContFatItensAux2.sProduto Then
                            dQuantidade = dQuantidade + objContFatItensAux2.dQuantidade
                            dValor = dValor + objContFatItensAux2.dVlrCobrar
                            dCusto = dCusto + objContFatItensAux2.dCusto
                            
                            If objContFatItensAux2.lMedicao <> 0 Then
                            
                                Set objItemMedicao = New ClassItensMedCtr
                            
                                objItemMedicao.lNumIntItensContrato = objContFatItensAux2.lNumIntItensContrato
                                objItemMedicao.iItem = objContFatItensAux2.iItem
                                objItemMedicao.lMedicao = objContFatItensAux2.lMedicao
                                objItemMedicao.dtDataCobranca = objContFatItensAux2.dtDataProxCobranca
                                objItemMedicao.dtDataRefIni = objContFatItensAux2.dtDataRefIni
                                objItemMedicao.dtDataRefFim = objContFatItensAux2.dtDataRefFim
                                
                                colMedicoes.Add objItemMedicao
                            
                            End If
                            
                        End If
                    End If
                
                Next
                
                If objContFatItensAux1.lMedicao <> 0 Then
                
                    Set objItemMedicao = New ClassItensMedCtr
                
                    objItemMedicao.lNumIntItensContrato = objContFatItensAux1.lNumIntItensContrato
                    objItemMedicao.iItem = objContFatItensAux1.iItem
                    objItemMedicao.lMedicao = objContFatItensAux1.lMedicao
                    objItemMedicao.dtDataCobranca = objContFatItensAux1.dtDataProxCobranca
                    objItemMedicao.dtDataRefIni = objContFatItensAux1.dtDataRefIni
                    objItemMedicao.dtDataRefFim = objContFatItensAux1.dtDataRefFim
                    
                    colMedicoes.Add objItemMedicao
                
                End If
                
                objContFatItensAux1.dCusto = objContFatItensAux1.dCusto + dCusto
                objContFatItensAux1.dQuantidade = objContFatItensAux1.dQuantidade + dQuantidade
                objContFatItensAux1.dVlrCobrar = objContFatItensAux1.dVlrCobrar + dValor
                
                colContFatItensAgrupado.Add objContFatItensAux1
                colcolMedicoes.Add colMedicoes
                
            End If
        
        Next
        
        For Each objContratoFatItens In colContFatItensAgrupado
                  
            iItem = iItem + 1
        
            Set objItemNF = New ClassItemNF
            Call objItemNF.Inicializa_Tributacao
            Call objItemNF.objTributacaoItemNF.Coloca_Auto
            
            objItemNF.dCusto = objContratoFatItens.dCusto
            objItemNF.dQuantidade = objContratoFatItens.dQuantidade
            objItemNF.dPrecoUnitario = objContratoFatItens.dVlrCobrar / objContratoFatItens.dQuantidade
            objItemNF.dValorTotal = objContratoFatItens.dVlrCobrar
            objItemNF.sCcl = objContratoFatItens.sCcl
            objItemNF.sDescricaoItem = objContratoFatItens.sDescProd
            objItemNF.sProduto = objContratoFatItens.sProduto
            objItemNF.iItem = iItem
            objItemNF.sUnidadeMed = objContratoFatItens.sUM
        
            objItemNF.objCobrItensContrato.lNumIntItensContrato = objContratoFatItens.lNumIntItensContrato
            objItemNF.objCobrItensContrato.dtDataUltCobranca = objContratoFatItens.dtDataProxCobranca
            objItemNF.objCobrItensContrato.lNumIntDocCobranca = objGeracaoFatContrato.lNumIntDoc
            objItemNF.objCobrItensContrato.dtDataRefIni = objContratoFatItens.dtDataRefIni
            objItemNF.objCobrItensContrato.dtDataRefFim = objContratoFatItens.dtDataRefFim
            
            Set objItemNF.objCobrItensContrato.colMedicoes = colcolMedicoes.Item(iItem)
        
            dValorTotal = dValorTotal + objItemNF.dValorTotal
        
            objNFiscal.ColItensNF.Add1 objItemNF
        
            If Len(Trim(objContratoFatItens.sNaturezaOp)) <> 0 Then
                objItemNF.objTributacaoItemNF.sNaturezaOp = objContratoFatItens.sNaturezaOp
                objItemNF.objTributacaoItemNF.iNaturezaOpManual = VAR_PREENCH_MANUAL
            Else
                If Len(Trim(objContratoFat.sNaturezaOp)) <> 0 Then
                    objItemNF.objTributacaoItemNF.sNaturezaOp = objContratoFat.sNaturezaOp
                    objItemNF.objTributacaoItemNF.iNaturezaOpManual = VAR_PREENCH_MANUAL
                End If
            End If
            
            If objContratoFatItens.iTipoTributacao <> 0 Then
                objItemNF.objTributacaoItemNF.iTipoTributacao = objContratoFatItens.iTipoTributacao
                objItemNF.objTributacaoItemNF.iTipoTributacaoManual = VAR_PREENCH_MANUAL
            Else
                If objContratoFat.iTipoTributacao <> 0 Then
                    objItemNF.objTributacaoItemNF.iTipoTributacao = objContratoFat.iTipoTributacao
                    objItemNF.objTributacaoItemNF.iTipoTributacaoManual = VAR_PREENCH_MANUAL
                End If
            End If
            
            Set objProduto = New ClassProduto
            objProduto.sCodigo = objItemNF.sProduto
            
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 136032
            
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
            
        Next
        
        objNFiscal.dValorProdutos = dValorTotal
        
        Set objTribTab = New ClassTribTab
        
        lErro = objTribTab.TributacaoNF_Reset(objNFiscal)
        If lErro <> SUCESSO Then gError 130020
        
        'obtem dados do contrato
        Call objTribTab.TipoNFiscal_Definir(objTipoDocInfo.iCodigo, objTipoDocInfo.sSigla)
        Call objTribTab.DataEmissao_Alterada(objGeracaoFatContrato.dtDataEmissao)
        Call objTribTab.Serie_Alterada(objNFiscal.sSerie)
        
        lErro = gobjTributacao.AtualizaImpostos(objTribTab.mvarobjNFTributacao, 0)
        If lErro <> SUCESSO Then gError 130021
        
        lErro = objTribTab.PreencherNF(objNFiscal)
        If lErro <> SUCESSO Then gError 130022
        
        Set objTribTab = Nothing
        
        objNFiscal.sNaturezaOp = objNFiscal.objTributacaoNF.sNaturezaOpInterna
        
        'acertar valor total e da parcela, considerando os tributos, retencóes,...
        objNFiscal.dValorTotal = dValorTotal + objNFiscal.objTributacaoNF.dICMSSubstValor + objNFiscal.objTributacaoNF.dIPIValor + objNFiscal.dValorFrete + objNFiscal.dValorSeguro + objNFiscal.dValorOutrasDespesas + IIf(objNFiscal.objTributacaoNF.iISSIncluso = 0, objNFiscal.objTributacaoNF.dISSValor, 0)
        'Adiciona a parcela na coleção de parcelas da Nota Fiscal
        dValorLiquido = objNFiscal.dValorTotal - (objNFiscal.objTributacaoNF.dCOFINSRetido + objNFiscal.objTributacaoNF.dCSLLRetido + objNFiscal.objTributacaoNF.dIRRFValor + objNFiscal.objTributacaoNF.dPISRetido + objNFiscal.objTributacaoNF.dISSRetido)
         
        lErro = NFiscalContrato_Gera_Parcelas(objNFiscal)
        If lErro <> SUCESSO Then gError 136029
        
        lErro = CF("NFiscalContrato_Gera_InfoBoletos", objNFiscal)
        If lErro <> SUCESSO Then gError 136029
        
        lErro = NFiscalContrato_Gera_Peso(objNFiscal)
        If lErro <> SUCESSO Then gError 136034
        
        iTemp = gobjCRFAT.iCreditoVerificaLimite
        gobjCRFAT.iCreditoVerificaLimite = 0
        
        lErro = CF("NFiscal_Valida_Diversos", objNFiscal)
        If lErro <> SUCESSO Then gError 136029
    
        lErro = NFiscalContrato_Trata_Msg(objNFiscal)
        If lErro <> SUCESSO Then gError 136029
        
        'especifico para nao doadores da escola do beit lubavitch
        bChederNaoDoador = False
        If gsNomeFilialEmpresa = "Cheder" And left(UCase(objContratoFat.sContrato), 1) <> "C" Then
        
            dMultaPadrao = gobjCRFAT.dPercMulta
            gobjCRFAT.dPercMulta = 0.02
            
            dJurosPadrao = gobjCRFAT.dPercJurosDiario
            gobjCRFAT.dPercJurosDiario = 0.0003
            
            bChederNaoDoador = True
            
        End If
        
        lErro = CF("NFiscalFatura_Grava", objNFiscal, objContabil)
        gobjCRFAT.iCreditoVerificaLimite = iTemp
        
        'especifico para nao doadores da escola do beit lubavitch
        If bChederNaoDoador Then
        
            gobjCRFAT.dPercMulta = dMultaPadrao
            gobjCRFAT.dPercJurosDiario = dJurosPadrao
        
        End If
        
        If lErro <> SUCESSO Then
        'NOTA COM PROBLEMA
            objContratoFat.iTipoErro = 1 '???
        Else
            objContratoFat.iTipoErro = 0
            objContratoFat.dValor = objNFiscal.dValorTotal
            objContratoFat.lNumNotaFiscal = objNFiscal.lNumNotaFiscal
            If ISSerieEletronica(objNFiscal.sSerie) Then
                colNumIntNF.Add objNFiscal.lNumIntDoc
            End If
        
        End If
        objContratoFat.iFilialEmpresa = objGeracaoFatContrato.iFilialEmpresa
        objContratoFat.lNumIntDocCobranca = objGeracaoFatContrato.lNumIntDoc
        
        lErro = CF("Contrato_Grava_RelErro", objContratoFat)
        If lErro <> SUCESSO Then gError 132850
                
        dValorTotal = 0
            
        'Atualiza Barra de Progressão
        lErro = Rotina_GeraContrato_AtualizaTelaBatch
        If lErro <> SUCESSO Then gError 129826
            
    Next
    
    If colNumIntNF.Count > 0 Then
        lErro = CF("NFE_Grava1", 0, objTipoDocInfo, colNumIntNF)
        If lErro <> SUCESSO Then gError 129826
    End If

    NFiscalContrato_Gera = SUCESSO

    Exit Function

Erro_NFiscalContrato_Gera:

    NFiscalContrato_Gera = gErr

    Select Case gErr

        Case 129825 To 129826, 130020, 130021, 130022, 130023, 132850, 132974, 136029, 136034
        
        Case 130024
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objTipoDocInfo.iCodigo)
            
        Case 136032
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161554)

    End Select

    Exit Function

End Function

Private Function NFiscalContrato_Gera_Parcelas(objNFiscal As ClassNFiscal) As Long

Dim lErro As Long
Dim objCondicaoPagto As New ClassCondicaoPagto
Dim dValorLiquido As Double
Dim iIndice As Integer

On Error GoTo Erro_NFiscalContrato_Gera_Parcelas

    dValorLiquido = objNFiscal.dValorTotal - (objNFiscal.objTributacaoNF.dCOFINSRetido + objNFiscal.objTributacaoNF.dCSLLRetido + objNFiscal.objTributacaoNF.dIRRFValor + objNFiscal.objTributacaoNF.dPISRetido + objNFiscal.objTributacaoNF.dISSRetido)

    objCondicaoPagto.iCodigo = objNFiscal.iCondicaoPagto
    
    'Lê a condição de pagamento
    lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
    If lErro <> SUCESSO And lErro <> 19205 Then gError 136030
    
    'Calcula os valores das Parcelas
    objCondicaoPagto.dValorTotal = dValorLiquido
    objCondicaoPagto.dtDataRef = objNFiscal.dtDataReferencia
    
    'Calcula os valores das Parcelas
    lErro = CF("CondicaoPagto_CalculaParcelas", objCondicaoPagto, True, True)
    If lErro <> SUCESSO Then gError 136031

    'Coloca os valores das Parcelas no Grid Parcelas
    For iIndice = 1 To objCondicaoPagto.colParcelas.Count
        objNFiscal.ColParcelaReceber.Add 0, 0, iIndice, STATUS_ABERTO, objCondicaoPagto.colParcelas(iIndice).dtVencimento, objCondicaoPagto.colParcelas(iIndice).dtVencimento, objCondicaoPagto.colParcelas(iIndice).dValor, objCondicaoPagto.colParcelas(iIndice).dValor, 1, CARTEIRA_CARTEIRA, COBRADOR_PROPRIA_EMPRESA, "", 0, 0, 0, 0, 0, 0, 0, DATA_NULA, 0, 0, DATA_NULA, 0, 0, DATA_NULA, 0, 0, 0, 0, 0, 0, "", objCondicaoPagto.colParcelas(iIndice).dValor
    Next
    
    NFiscalContrato_Gera_Parcelas = SUCESSO
    
    Exit Function

Erro_NFiscalContrato_Gera_Parcelas:

    NFiscalContrato_Gera_Parcelas = gErr

    Select Case gErr
    
        Case 136030, 136031

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161555)

    End Select
       
    Exit Function

End Function

Private Function NFiscalContrato_Gera_Peso(objNFiscal As ClassNFiscal) As Long

Dim lErro As Long
Dim objProduto As ClassProduto
Dim objItemNF As ClassItemNF
Dim dPesoL As Double
Dim dPesoB As Double
Dim dFator As Double

On Error GoTo Erro_NFiscalContrato_Gera_Peso

    For Each objItemNF In objNFiscal.ColItensNF
    
        Set objProduto = New ClassProduto
    
        objProduto.sCodigo = objItemNF.sProduto
    
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 136032
    
        'Converte a unidade de medida de UM de Venda para a UM da linha do Grid
        lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objItemNF.sUnidadeMed, objProduto.sSiglaUMVenda, dFator)
        If lErro <> SUCESSO Then gError 136033
    
        dPesoL = dPesoL + (objProduto.dPesoLiq * objItemNF.dQuantidade * dFator)
        dPesoB = dPesoB + (objProduto.dPesoBruto * objItemNF.dQuantidade * dFator)
    
    Next
    
    objNFiscal.dPesoBruto = dPesoB
    objNFiscal.dPesoLiq = dPesoL
    
    NFiscalContrato_Gera_Peso = SUCESSO
    
    Exit Function

Erro_NFiscalContrato_Gera_Peso:

    NFiscalContrato_Gera_Peso = gErr

    Select Case gErr
    
        Case 136032, 136033

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161556)

    End Select
       
    Exit Function

End Function

Private Function NFiscalContrato_Trata_Msg(ByVal objNFiscal As ClassNFiscal)

Dim lErro As Long
Dim objMsg As ClassMensagensRegra
Dim colMsg As New Collection
Dim sMsg(0 To 1) As String
Dim objTransacao As Object
Dim objGeracaoNFiscal As New ClassGeracaoNFiscal
Dim objPedidoVenda As New ClassPedidoDeVenda

On Error GoTo Erro_NFiscalContrato_Trata_Msg:

    Set objGeracaoNFiscal.objNFiscal = objNFiscal
    Set objGeracaoNFiscal.objPedidoVenda = objPedidoVenda

    Set objTransacao = objGeracaoNFiscal

    lErro = CF("RegrasMsg_Calcula_Regras", objTransacao, REGRAMSG_TIPODOC_NF, colMsg)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    For Each objMsg In colMsg
        sMsg(objMsg.iTipoMsg) = sMsg(objMsg.iTipoMsg) & IIf(Len(Trim(sMsg(objMsg.iTipoMsg))) = 0, "", vbNewLine) & objMsg.sMensagem
    Next
    
    objGeracaoNFiscal.objNFiscal.sMensagemNota = sMsg(REGRAMSG_TIPOMSG_NORMAL)
    objGeracaoNFiscal.objNFiscal.sMensagemCorpoNota = sMsg(REGRAMSG_TIPOMSG_CORPO)

    NFiscalContrato_Trata_Msg = SUCESSO

    Exit Function

Erro_NFiscalContrato_Trata_Msg:

    NFiscalContrato_Trata_Msg = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208263)

    End Select

    Exit Function
    
End Function

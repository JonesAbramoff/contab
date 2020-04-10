VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassECFSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'mario
Private Declare Function iCFImprimir_NFCe_Daruma Lib "DarumaFramework.dll" (ByVal sXML As String, ByVal sOFF As String, ByVal sQR As String, ByVal icol As Integer, ByVal iAux As Integer) As Integer
Private Declare Function iImprimirCFe_SAT_Daruma Lib "DarumaFramework.dll" (ByVal sXML As String, ByVal sTipo As String) As Integer

Public Sub MovimentoCaixa_Gera_Log(colRegistro As Collection, ByVal objMovimentoCaixa As ClassMovimentoCaixa)

Dim sTemp As String
Dim sLog As String

    'tulio incluiu o campo lTransferencia aqui em 13/09/02
    sTemp = TIPOREGISTROECF_MOVIMENTOCAIXA & Chr(vbKeyControl) & CStr(objMovimentoCaixa.iTipo) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.dHora) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.dtDataMovimento) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.iGerente) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.iCodOperador) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.lSequencial) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.iFilialEmpresa) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.iCaixa) & Chr(vbKeyEscape) & Chr(vbKeyEnd)
    
    If Len(Trim(sTemp)) + Len(Trim(sLog)) > MAX_TAM_STRING_MOVIMENTOCAIXA Then
        If Len(Trim(sLog)) <> 0 Then colRegistro.Add sLog
        sLog = ""
    End If
    sLog = sLog & sTemp
    sTemp = ""
    
    If Len(Trim(sLog)) > 0 Then colRegistro.Add sLog
    
End Sub

Public Sub FechamentoCaixa_Gera_Log(colRegistro As Collection, ByVal objMovimentoCaixa As ClassMovimentoCaixa, ByVal objLojaArqFisMestre As ClassLojaArqFisMestre, ByVal colLojaArqFisAnalitico As Collection, ByVal objUltimaReducao As ClassUltimaReducao)

Dim sTemp As String
Dim sLog As String
Dim objLojaArqFisAnalitico As ClassLojaArqFisAnalitico
Dim objAliquota As ClassAliquotaICMS
Dim lCOOIni As Long

    sTemp = TIPOREGISTROECF_FECHAMENTOCAIXA & Chr(vbKeyControl) & CStr(objMovimentoCaixa.iTipo) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.dHora) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.dtDataMovimento) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.iGerente) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.iCodOperador) & _
    Chr(vbKeyEscape) & CStr(objMovimentoCaixa.lSequencial) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.iFilialEmpresa) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.iCaixa) & Chr(vbKeyEscape) & _
    CStr(giCodECF) & Chr(vbKeyEscape) & objLojaArqFisMestre.sNumSerieECF & Chr(vbKeyEscape) & objLojaArqFisMestre.dGrandeTotal & Chr(vbKeyEscape) & objLojaArqFisMestre.lCOOIni & Chr(vbKeyEscape) & objLojaArqFisMestre.lCOOFim & Chr(vbKeyEscape) & objLojaArqFisMestre.lCRO & Chr(vbKeyEscape) & objLojaArqFisMestre.lCRZ & Chr(vbKeyEscape) & objLojaArqFisMestre.dVendaBruta & Chr(vbKeyEscape) & objLojaArqFisMestre.dtData & Chr(vbKeyEscape) & objLojaArqFisMestre.iNumEquip & Chr(vbKeyEscape)
    
    For Each objLojaArqFisAnalitico In colLojaArqFisAnalitico
    
        sTemp = sTemp & objLojaArqFisAnalitico.sSituacaoTrib & Chr(vbKeyEscape) & objLojaArqFisAnalitico.dTotalizador & Chr(vbKeyEscape)
    
        If Len(Trim(sTemp)) + Len(Trim(sLog)) + 1 > MAX_TAM_STRING_MOVIMENTOCAIXA Then
            If Len(Trim(sLog)) <> 0 Then colRegistro.Add sLog
            sLog = ""
        End If
        sLog = sLog & sTemp
        sTemp = ""
    
    Next
    
    sTemp = sTemp & objUltimaReducao.dAcrescimoICMS & Chr(vbKeyShift) & objUltimaReducao.dAcrescimoISS & Chr(vbKeyShift) & objUltimaReducao.dAcrescimoNaoFiscal & Chr(vbKeyShift) & _
    objUltimaReducao.dCancelamentoICMS & Chr(vbKeyShift) & objUltimaReducao.dCancelamentoISS & Chr(vbKeyShift) & objUltimaReducao.dCancelamentoNaoFiscal & Chr(vbKeyShift) & _
    objUltimaReducao.dDescontoICMS & Chr(vbKeyShift) & objUltimaReducao.dDescontoISS & Chr(vbKeyShift) & objUltimaReducao.dDescontoNaoFiscal & Chr(vbKeyShift) & _
    objUltimaReducao.dIsentoICMS & Chr(vbKeyShift) & objUltimaReducao.dIsentoISS & Chr(vbKeyShift) & objUltimaReducao.dNaoIncideICMS & Chr(vbKeyShift) & _
    objUltimaReducao.dNaoIncideISS & Chr(vbKeyShift) & objUltimaReducao.dSubstTribICMS & Chr(vbKeyShift) & objUltimaReducao.dSubstTribISS & Chr(vbKeyShift) & _
    objUltimaReducao.dTotalNaoFiscal & Chr(vbKeyShift) & objUltimaReducao.dtDataReducao & Chr(vbKeyShift) & objUltimaReducao.dVendaBruta & Chr(vbKeyShift) & _
    objUltimaReducao.sIncidenciaDescontoISS & Chr(vbKeyShift) & objUltimaReducao.sHoraReducao & Chr(vbKeyShift) & objUltimaReducao.dtDataMovimento & Chr(vbKeyShift) & _
    objUltimaReducao.lCOOReducaoZ & Chr(vbKeyShift)
    
    If Len(Trim(sTemp)) + Len(Trim(sLog)) + 1 > MAX_TAM_STRING_MOVIMENTOCAIXA Then
        If Len(Trim(sLog)) <> 0 Then colRegistro.Add sLog
        sLog = ""
    End If
    sLog = sLog & sTemp
    sTemp = ""
    
    For Each objAliquota In objUltimaReducao.colAliquotas
        
        sTemp = sTemp & objAliquota.dAliquota & Chr(vbKeyShift) & objAliquota.iISS & Chr(vbKeyShift) & objAliquota.dValorTotalizadoLoja & Chr(vbKeyShift)
    
        If Len(Trim(sTemp)) + Len(Trim(sLog)) + 1 > MAX_TAM_STRING_MOVIMENTOCAIXA Then
            If Len(Trim(sLog)) <> 0 Then colRegistro.Add sLog
            sLog = ""
        End If
        sLog = sLog & sTemp
        sTemp = ""
    
    Next
    
    sTemp = sTemp & Chr(vbKeyEnd)
    
    If Len(Trim(sTemp)) + Len(Trim(sLog)) + 1 > MAX_TAM_STRING_MOVIMENTOCAIXA Then
        If Len(Trim(sLog)) <> 0 Then colRegistro.Add sLog
        sLog = ""
    End If
    sLog = sLog & sTemp
    sTemp = ""
    
    If Len(Trim(sLog)) > 0 Then colRegistro.Add sLog
    
End Sub


Public Sub Venda_Gera_Log(colRegistro As Collection, ByVal objVenda As ClassVenda, Optional ByVal iOrcamento As Integer = 0)

Dim objCheque As ClassChequePre
Dim objMovCx As ClassMovimentoCaixa
Dim objCarneParc As ClassCarneParcelas
Dim objItens As ClassItemCupomFiscal
Dim objTroca As ClassTroca
Dim sLog As String
Dim sTemp As String

    'Joga o tipo go objVenda
    sTemp = TIPOREGISTROECF_VENDAS & Chr(vbKeyControl) & objVenda.iTipo
    
    sTemp = sTemp & Chr(vbKeySeparator)

    If iOrcamento = 0 Or (objVenda.iTipo = OPTION_CF Or ((objVenda.iTipo = OPTION_DAV Or objVenda.iTipo = OPTION_PREVENDA) And objVenda.objCupomFiscal.iStatus = STATUS_BAIXADO)) Then
    
        If Len(objVenda.objCarne.sCodBarrasCarne) > 0 Then
        
            'Joga o objCarne
            sTemp = sTemp & Chr(vbKeyControl) & objVenda.objCarne.dtDataReferencia & Chr(vbKeyEscape) & objVenda.objCarne.iFilialEmpresa & Chr(vbKeyEscape) & objVenda.objCarne.iStatus & Chr(vbKeyEscape) & objVenda.objCarne.lCliente & Chr(vbKeyEscape) & objVenda.objCarne.sAutorizacao & Chr(vbKeyEscape) & objVenda.objCarne.sCodBarrasCarne & Chr(vbKeyEscape)
            
            If Len(Trim(sTemp)) + Len(Trim(sLog)) + 10 > MAX_TAM_STRING_MOVIMENTOCAIXA Then
                If Len(Trim(sLog)) <> 0 Then colRegistro.Add sLog
                sLog = ""
            End If
            sLog = sLog & sTemp
            sTemp = ""
                    
            'Jogo todos os objcarne.colparcelas na string
            For Each objCarneParc In objVenda.objCarne.colParcelas
                sTemp = Chr(vbKeyShift) & objCarneParc.dtDataVencimento & Chr(vbKeyEscape) & objCarneParc.dValor & Chr(vbKeyEscape) & objCarneParc.iParcela & Chr(vbKeyEscape) & objCarneParc.iStatus & Chr(vbKeyEscape) & Chr(vbKeyEscape)
                
                If Len(Trim(sTemp)) + Len(Trim(sLog)) + 10 > MAX_TAM_STRING_MOVIMENTOCAIXA Then
                    If Len(Trim(sLog)) <> 0 Then colRegistro.Add sLog
                    sLog = ""
                End If
                sLog = sLog & sTemp
                sTemp = ""
            Next
        End If
    
    End If
    
    sTemp = sTemp & Chr(vbKeySeparator)
    
    If objVenda.objCupomFiscal.dHoraEmissao = 0 Then Call MsgBox("ERRO HORA ZERADA")
    
    'Joga o objCupomFiscal
    sTemp = sTemp & Chr(vbKeyControl) & objVenda.objCupomFiscal.dtDataEmissao & Chr(vbKeyEscape) & objVenda.objCupomFiscal.dHoraEmissao & Chr(vbKeyEscape) & objVenda.objCupomFiscal.dValorAcrescimo & Chr(vbKeyEscape) & objVenda.objCupomFiscal.dValorDesconto & Chr(vbKeyEscape) & objVenda.objCupomFiscal.dValorProdutos & Chr(vbKeyEscape) & objVenda.objCupomFiscal.dValorTotal & Chr(vbKeyEscape) & objVenda.objCupomFiscal.iECF & Chr(vbKeyEscape) & objVenda.objCupomFiscal.iFilialEmpresa & Chr(vbKeyEscape) & objVenda.objCupomFiscal.iStatus & Chr(vbKeyEscape) & objVenda.objCupomFiscal.iTabelaPreco
    sTemp = sTemp & Chr(vbKeyEscape) & objVenda.objCupomFiscal.lGerenteCancel & Chr(vbKeyEscape) & objVenda.objCupomFiscal.lNumero & Chr(vbKeyEscape) & objVenda.objCupomFiscal.lNumOrcamento & Chr(vbKeyEscape) & objVenda.objCupomFiscal.iVendedor & Chr(vbKeyEscape) & objVenda.objCupomFiscal.lCliente & Chr(vbKeyEscape) & objVenda.objCupomFiscal.sCPFCGC & Chr(vbKeyEscape) & objVenda.objCupomFiscal.sNomeCliente & Chr(vbKeyEscape) & objVenda.objCupomFiscal.sMotivoCancel & Chr(vbKeyEscape) & objVenda.objCupomFiscal.sNaturezaOp & Chr(vbKeyEscape) & objVenda.objCupomFiscal.lDuracao & Chr(vbKeyEscape)
    
    If Len(Trim(sTemp)) + Len(Trim(sLog)) + 10 > MAX_TAM_STRING_MOVIMENTOCAIXA Then
        If Len(Trim(sLog)) <> 0 Then colRegistro.Add sLog
        sLog = ""
    End If
    sLog = sLog & sTemp
    sTemp = ""
    
    'Joga o objCupomFiscal
    sTemp = sTemp & objVenda.objCupomFiscal.iCodCaixa & Chr(vbKeyEscape) & objVenda.objCupomFiscal.lCOOCupomOrigDAV & Chr(vbKeyEscape) & objVenda.objCupomFiscal.sTipoECF & Chr(vbKeyEscape) & objVenda.objCupomFiscal.sMarcaECF & Chr(vbKeyEscape) & objVenda.objCupomFiscal.sModeloECF & Chr(vbKeyEscape) & objVenda.objCupomFiscal.sNumSerieECF & Chr(vbKeyEscape) & objVenda.objCupomFiscal.lCCF & Chr(vbKeyEscape) & objVenda.objCupomFiscal.lNumeroDAV & Chr(vbKeyEscape) & objVenda.objCupomFiscal.sCPFCGC1 & Chr(vbKeyEscape) & objVenda.objCupomFiscal.sEndereco
    sTemp = sTemp & Chr(vbKeyEscape) & objVenda.objCupomFiscal.iSequencialECF & Chr(vbKeyEscape) & objVenda.objCupomFiscal.iSequencialECFOrigDAV & Chr(vbKeyEscape) & objVenda.objCupomFiscal.iDAVImpresso & Chr(vbKeyEscape) & objVenda.objCupomFiscal.dtDataReducao & Chr(vbKeyEscape)
    
    If Len(Trim(sTemp)) + Len(Trim(sLog)) + 10 > MAX_TAM_STRING_MOVIMENTOCAIXA Then
        If Len(Trim(sLog)) <> 0 Then colRegistro.Add sLog
        sLog = ""
    End If
    sLog = sLog & sTemp
    sTemp = ""
    
    If Len(Trim(objVenda.objCupomFiscal.sSATChaveAcesso)) <> 0 Then
        sTemp = sTemp & objVenda.objCupomFiscal.sSATChaveAcesso & Chr(vbKeyEscape) & objVenda.objCupomFiscal.sSATArqXml & Chr(vbKeyEscape)
    Else
        sTemp = sTemp & objVenda.objCupomFiscal.sNFeChaveAcesso & Chr(vbKeyEscape) & objVenda.objCupomFiscal.sNFeArqXml & Chr(vbKeyEscape)
    End If
    
    If Len(Trim(sTemp)) + Len(Trim(sLog)) + 10 > MAX_TAM_STRING_MOVIMENTOCAIXA Then
        If Len(Trim(sLog)) <> 0 Then colRegistro.Add sLog
        sLog = ""
    End If
    sLog = sLog & sTemp
    sTemp = ""
    
    sTemp = sTemp & objVenda.objCupomFiscal.objNF.sSerie & Chr(vbKeyEscape) & objVenda.objCupomFiscal.objNF.lNumNotaFiscal & Chr(vbKeyEscape) & objVenda.objCupomFiscal.objNF.dtDataEmissao & Chr(vbKeyEscape) & objVenda.objCupomFiscal.objNF.sDestino & Chr(vbKeyEscape)
    
    If Len(Trim(sTemp)) + Len(Trim(sLog)) + 10 > MAX_TAM_STRING_MOVIMENTOCAIXA Then
        If Len(Trim(sLog)) <> 0 Then colRegistro.Add sLog
        sLog = ""
    End If
    sLog = sLog & sTemp
    sTemp = ""
    
    
    'Jogo todos os objCupomFiscal.colItens na string
    For Each objItens In objVenda.objCupomFiscal.colItens
'        If objItens.icancel = ITEM_NORMAL Then
            sTemp = Chr(vbKeyShift) & objItens.dAliquotaICMS & Chr(vbKeyEscape) & objItens.dPercDesc & Chr(vbKeyEscape) & objItens.dPrecoUnitario & Chr(vbKeyEscape) & objItens.dQuantidade & Chr(vbKeyEscape)
            sTemp = sTemp & objItens.dValorDesconto & Chr(vbKeyEscape) & objItens.iStatus & Chr(vbKeyEscape) & objItens.sProduto & Chr(vbKeyEscape) & objItens.sUnidadeMed & Chr(vbKeyEscape)
            sTemp = sTemp & objItens.sSituacaoTrib & Chr(vbKeyEscape) & objVenda.objCupomFiscal.iCodCaixa & Chr(vbKeyEscape) & objItens.iItem & Chr(vbKeyEscape)
            
            If Len(Trim(sTemp)) + Len(Trim(sLog)) + 10 > MAX_TAM_STRING_MOVIMENTOCAIXA Then
                If Len(Trim(sLog)) <> 0 Then colRegistro.Add sLog
                sLog = ""
            End If
            sLog = sLog & sTemp
            sTemp = ""
            
            sTemp = sTemp & objItens.objTributacaoDocItem.dQuantidade & Chr(vbKeyEscape) & objItens.objTributacaoDocItem.dPrecoUnitario & Chr(vbKeyEscape) & objItens.objTributacaoDocItem.sNaturezaOp & Chr(vbKeyEscape)
            sTemp = sTemp & objItens.objTributacaoDocItem.iTipoTributacao & Chr(vbKeyEscape) & objItens.objTributacaoDocItem.iICMSTipo & Chr(vbKeyEscape) & objItens.objTributacaoDocItem.dICMSBase & Chr(vbKeyEscape)
            sTemp = sTemp & objItens.objTributacaoDocItem.dICMSPercRedBase & Chr(vbKeyEscape) & objItens.objTributacaoDocItem.dICMSAliquota & Chr(vbKeyEscape) & objItens.objTributacaoDocItem.dICMSValor & Chr(vbKeyEscape)
            sTemp = sTemp & objItens.objTributacaoDocItem.iPISTipo & Chr(vbKeyEscape) & objItens.objTributacaoDocItem.iPISTipoCalculo & Chr(vbKeyEscape) & objItens.objTributacaoDocItem.dPISBase & Chr(vbKeyEscape)
            sTemp = sTemp & objItens.objTributacaoDocItem.dPISAliquota & Chr(vbKeyEscape) & objItens.objTributacaoDocItem.dPISAliquotaValor & Chr(vbKeyEscape) & objItens.objTributacaoDocItem.dPISQtde & Chr(vbKeyEscape)
            sTemp = sTemp & objItens.objTributacaoDocItem.dPISValor & Chr(vbKeyEscape) & objItens.objTributacaoDocItem.iPISSTTipoCalculo & Chr(vbKeyEscape) & objItens.objTributacaoDocItem.dPISSTBase & Chr(vbKeyEscape)
            sTemp = sTemp & objItens.objTributacaoDocItem.dPISSTAliquota & Chr(vbKeyEscape) & objItens.objTributacaoDocItem.dPISSTAliquotaValor & Chr(vbKeyEscape) & objItens.objTributacaoDocItem.dPISSTQtde & Chr(vbKeyEscape)
            sTemp = sTemp & objItens.objTributacaoDocItem.dPISSTValor & Chr(vbKeyEscape) & objItens.objTributacaoDocItem.iCOFINSTipo & Chr(vbKeyEscape) & objItens.objTributacaoDocItem.iCOFINSTipoCalculo & Chr(vbKeyEscape)
            sTemp = sTemp & objItens.objTributacaoDocItem.dCOFINSBase & Chr(vbKeyEscape) & objItens.objTributacaoDocItem.dCOFINSAliquota & Chr(vbKeyEscape) & objItens.objTributacaoDocItem.dCOFINSAliquotaValor & Chr(vbKeyEscape)
            sTemp = sTemp & objItens.objTributacaoDocItem.dCOFINSQtde & Chr(vbKeyEscape) & objItens.objTributacaoDocItem.dCOFINSValor & Chr(vbKeyEscape) & objItens.objTributacaoDocItem.iCOFINSSTTipoCalculo & Chr(vbKeyEscape)
            
            If Len(Trim(sTemp)) + Len(Trim(sLog)) + 10 > MAX_TAM_STRING_MOVIMENTOCAIXA Then
                If Len(Trim(sLog)) <> 0 Then colRegistro.Add sLog
                sLog = ""
            End If
            sLog = sLog & sTemp
            sTemp = ""
            
            sTemp = sTemp & objItens.objTributacaoDocItem.dCOFINSSTBase & Chr(vbKeyEscape) & objItens.objTributacaoDocItem.dCOFINSSTAliquota & Chr(vbKeyEscape) & objItens.objTributacaoDocItem.dCOFINSSTAliquotaValor & Chr(vbKeyEscape)
            sTemp = sTemp & objItens.objTributacaoDocItem.dCOFINSSTQtde & Chr(vbKeyEscape) & objItens.objTributacaoDocItem.dCOFINSSTValor & Chr(vbKeyEscape) & objItens.objTributacaoDocItem.sCST & Chr(vbKeyEscape)
            sTemp = sTemp & objItens.objTributacaoDocItem.sISSQN & Chr(vbKeyEscape) & objItens.objTributacaoDocItem.dISSBase & Chr(vbKeyEscape) & objItens.objTributacaoDocItem.dISSAliquota & Chr(vbKeyEscape)
            sTemp = sTemp & objItens.objTributacaoDocItem.sISSCidadeIBGE & Chr(vbKeyEscape) & objItens.objTributacaoDocItem.dISSValor & Chr(vbKeyEscape) & objItens.objTributacaoDocItem.iICMSBaseModalidade & Chr(vbKeyEscape)
            sTemp = sTemp & objItens.objTributacaoDocItem.iOrigemMercadoria & Chr(vbKeyEscape) & objItens.objTributacaoDocItem.sGenero & Chr(vbKeyEscape) & objItens.objTributacaoDocItem.dQtdTrib & Chr(vbKeyEscape)
            sTemp = sTemp & objItens.objTributacaoDocItem.sUMTrib & Chr(vbKeyEscape) & objItens.objTributacaoDocItem.dValorUnitTrib & Chr(vbKeyEscape) & objItens.objTributacaoDocItem.sISSCST & Chr(vbKeyEscape)
            sTemp = sTemp & objItens.objTributacaoDocItem.dICMSValorIsento & Chr(vbKeyEscape) & objItens.objTributacaoDocItem.iICMSMotivo & Chr(vbKeyEscape) & objItens.objTributacaoDocItem.iICMSSimplesTipo & Chr(vbKeyEscape)
            sTemp = sTemp & objItens.objTributacaoDocItem.iISSTipo & Chr(vbKeyEscape) & objItens.objTributacaoDocItem.iRegimeTributario & Chr(vbKeyEscape) & objItens.objTributacaoDocItem.sCSOSN & Chr(vbKeyEscape)
            sTemp = sTemp & objItens.objTributacaoDocItem.dTotTrib & Chr(vbKeyEscape) & objItens.objTributacaoDocItem.iTotTribTipo & Chr(vbKeyEscape)
            
            If Len(Trim(sTemp)) + Len(Trim(sLog)) + 10 > MAX_TAM_STRING_MOVIMENTOCAIXA Then
                If Len(Trim(sLog)) <> 0 Then colRegistro.Add sLog
                sLog = ""
            End If
            sLog = sLog & sTemp
            sTemp = ""
'        End If
    Next
    
    If objVenda.objCupomFiscal.colItens.Count = 0 Then sTemp = sTemp & Chr(vbKeyShift)
    
    If iOrcamento = 0 Or (objVenda.iTipo = OPTION_CF Or ((objVenda.iTipo = OPTION_DAV Or objVenda.iTipo = OPTION_PREVENDA) And objVenda.objCupomFiscal.iStatus = STATUS_BAIXADO)) Then
    
        'Jogo todos os colMovCaixa na string
        For Each objMovCx In objVenda.colMovimentosCaixa
        
            objMovCx.dtDataMovimento = objVenda.objCupomFiscal.dtDataEmissao
            objMovCx.dHora = objVenda.objCupomFiscal.dHoraEmissao

            sTemp = sTemp & Chr(vbKeyControl) & objMovCx.dHora & Chr(vbKeyEscape) & objMovCx.dtDataMovimento & Chr(vbKeyEscape) & objMovCx.dValor & Chr(vbKeyEscape) & objMovCx.iAdmMeioPagto & Chr(vbKeyEscape) & objMovCx.iCaixa & Chr(vbKeyEscape) & objMovCx.iCodConta & Chr(vbKeyEscape) & objMovCx.iCodOperador & Chr(vbKeyEscape) & objMovCx.iFilialEmpresa & Chr(vbKeyEscape) & objMovCx.iGerente & Chr(vbKeyEscape) & objMovCx.iParcelamento & Chr(vbKeyEscape) & objMovCx.iTipo & Chr(vbKeyEscape) & objMovCx.iTipoCartao & Chr(vbKeyEscape) & objMovCx.lCupomFiscal & Chr(vbKeyEscape) & objMovCx.lNumero & Chr(vbKeyEscape) & objMovCx.lNumRefInterna & Chr(vbKeyEscape) & objMovCx.lSequencial & Chr(vbKeyEscape) & objMovCx.sFavorecido & Chr(vbKeyEscape) & objMovCx.sHistorico & Chr(vbKeyEscape) & objMovCx.dtDataPreDatado & Chr(vbKeyEscape)
            sTemp = sTemp & objMovCx.sAutorizacao & Chr(vbKeyEscape)
            
            If Len(Trim(sTemp)) + Len(Trim(sLog)) + 10 > MAX_TAM_STRING_MOVIMENTOCAIXA Then
                If Len(Trim(sLog)) <> 0 Then colRegistro.Add sLog
                sLog = ""
            End If
            sLog = sLog & sTemp
            sTemp = ""
        Next
    
    End If
    
    sTemp = sTemp & Chr(vbKeySeparator)
    
    If iOrcamento = 0 Or (objVenda.objCupomFiscal.iTipo = OPTION_CF Or ((objVenda.iTipo = OPTION_DAV Or objVenda.iTipo = OPTION_PREVENDA) And objVenda.objCupomFiscal.iStatus = STATUS_BAIXADO)) Then
    
        'Jogo todos os colcheques na string
        For Each objCheque In objVenda.colCheques
            sTemp = sTemp & Chr(vbKeyControl) & objCheque.dtDataDeposito & Chr(vbKeyEscape) & objCheque.dValor & Chr(vbKeyEscape) & objCheque.iAprovado & Chr(vbKeyEscape) & objCheque.iBanco & Chr(vbKeyEscape) & objCheque.iFilialEmpresaLoja & Chr(vbKeyEscape) & objCheque.iNaoEspecificado & Chr(vbKeyEscape) & objCheque.lNumero & Chr(vbKeyEscape) & objCheque.sAgencia & Chr(vbKeyEscape) & objCheque.sContaCorrente & Chr(vbKeyEscape) & objCheque.sCPFCGC & Chr(vbKeyEscape) & objCheque.lSequencialCaixa & Chr(vbKeyEscape)
            
            If Len(Trim(sTemp)) + Len(Trim(sLog)) + 10 > MAX_TAM_STRING_MOVIMENTOCAIXA Then
                If Len(Trim(sLog)) <> 0 Then colRegistro.Add sLog
                sLog = ""
            End If
            sLog = sLog & sTemp
            sTemp = ""
            
        Next
    
    End If
    
    sTemp = sTemp & Chr(vbKeySeparator)
    
    If iOrcamento = 0 Or (objVenda.objCupomFiscal.iTipo = OPTION_CF Or ((objVenda.iTipo = OPTION_DAV Or objVenda.iTipo = OPTION_PREVENDA) And objVenda.objCupomFiscal.iStatus = STATUS_BAIXADO)) Then
    
        'Jogo todos os colTroca na string
        For Each objTroca In objVenda.colTroca
            sTemp = sTemp & Chr(vbKeyControl) & objTroca.dQuantidade & Chr(vbKeyEscape) & objTroca.dValor & Chr(vbKeyEscape) & objTroca.sProduto & Chr(vbKeyEscape) & objTroca.sCodProduto & Chr(vbKeyEscape) & objTroca.sUnidadeMed & Chr(vbKeyEscape)
            
            If Len(Trim(sTemp)) + Len(Trim(sLog)) + 10 > MAX_TAM_STRING_MOVIMENTOCAIXA Then
                If Len(Trim(sLog)) <> 0 Then colRegistro.Add sLog
                sLog = ""
            End If
            sLog = sLog & sTemp
            sTemp = ""
            
        Next
    
    End If
    
    sLog = sLog & sTemp & Chr(vbKeyEnd)
    
    If Len(Trim(sLog)) > 0 Then colRegistro.Add sLog
    
End Sub

'Public Function Desmembrar_ECF(ByVal objVenda1 As ClassVenda) As Long
'
'Dim iPos As Integer
'Dim iIndice As Integer
'Dim iPosInicio As Integer
'Dim iPosFim As Integer
'Dim iPosMeio As Integer
'Dim iPosColInicio As Integer
'Dim iPosColFim As Integer
'Dim sRegistro As String
'Dim iPosStrFim As String
'Dim sNomeArq As String
'Dim objCheque As ClassChequePre
'Dim objMovCx As ClassMovimentoCaixa
'Dim objCarneParc As ClassCarneParcelas
'Dim objItens As ClassItemCupomFiscal
'Dim objTroca As ClassTroca
'Dim sTipo As String
'Dim objVenda As ClassVenda
'
'On Error GoTo Erro_Desmembrar_ECF
'
'    Set objVenda1 = New ClassVenda
'
'    sNomeArq = gsDirMVTEF & "MV" & CStr(Format(Date, "ddmmyy")) & (".txt")
'
'    Open sNomeArq For Input As #1
'
'    Do While Not EOF(1)
'
'        'Busca o próximo registro do arquivo
'        Line Input #1, sRegistro
'
'        'Primeira Posição
'        iPosInicio = 1
'
'        'Procura o Primeiro Control para saber o tipo do registro
'        iPos = InStr(iPosInicio, sRegistro, Chr(vbKeyControl))
'
'        sTipo = Mid(sRegistro, iPosInicio, iPos - iPosInicio)
'
'        If sTipo = TIPOREGISTROECF_VENDAS Then
'
'            'Procura o Primeiro Control para saber onde começa a string
'            iPosInicio = InStr(iPosInicio, sRegistro, Chr(vbKeyControl)) + 1
'
'            'Pega última posição e guarda
'            iPosFim = (InStr(iPosInicio, sRegistro, Chr(vbKeyEnd)))
'
'            'iPosicao1 Guarda a posição do Segundo Control(referente ao tipo)
'            iPosStrFim = InStr(iPosInicio + 1, sRegistro, Chr(vbKeyControl))
'
'            iPosInicio = iPosStrFim + 1
'
'            'Procura o Primeiro Escape dentro da String
'            iPosMeio = (InStr(iPosInicio, sRegistro, Chr(vbKeyEscape)))
'
'            'iPosicao1 Guarda a posição do terceiro Control(referente ao objcupom)
'            iPosColFim = InStr(iPosInicio, sRegistro, Chr(vbKeyControl))
'
'            'iPosicao1 Guarda a posição do inicio da colparcelas
'            iPosColInicio = InStr(iPosInicio, sRegistro, Chr(vbKeyShift))
'            If iPosColFim - iPosColInicio < 2 Then iPosColInicio = iPosColFim
'
'            iIndice = 0
'
'            Do While iPosInicio <= iPosMeio
'
'                iIndice = iIndice + 1
'                'Recolhe o objCarne
'                Select Case iIndice
'
'                    Case 1: objVenda1.objCarne.dtDataReferencia = StrParaDate(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                    Case 2: objVenda1.objCarne.iFilialEmpresa = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                    Case 3: objVenda1.objCarne.iStatus = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                    Case 4: objVenda1.objCarne.lCupomFiscal = StrParaLong(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                    Case 5: objVenda1.objCarne.lNumIntDoc = StrParaLong(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                    Case 6: objVenda1.objCarne.lCliente = StrParaLong(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                    Case 7: objVenda1.objCarne.sAutorizacao = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
'                    Case 8: objVenda1.objCarne.sCodBarrasCarne = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
'                    Case 9: Exit Do
'
'                End Select
'
'                'Atualiza as Posições
'                iPosInicio = iPosMeio + 1
'                iPosMeio = (InStr(iPosInicio, sRegistro, Chr(vbKeyEscape)))
'
'                If iPosMeio > iPosColInicio Then iPosMeio = iPosColInicio
'
'            Loop
'
'            Do While iPosColInicio < iPosColFim
'
'                iPosInicio = iPosColInicio + 1
'
'                'Procura o Primeiro Escape dentro da String
'                iPosMeio = (InStr(iPosInicio, sRegistro, Chr(vbKeyEscape)))
'
'                'Procura um shift dentro da String
'                iPosColInicio = (InStr(iPosInicio, sRegistro, Chr(vbKeyShift)))
'                If iPosColInicio = 0 Or iPosColInicio > iPosColFim Then iPosColInicio = iPosColFim
'
'                iIndice = 0
'                Set objCarneParc = New ClassCarneParcelas
'
'                Do While iPosInicio < iPosMeio
'
'                    iIndice = iIndice + 1
'                    'Recolhe as parcelas
'                    Select Case iIndice
'
'                        Case 1: objCarneParc.dtDataVencimento = StrParaDate(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                        Case 2: objCarneParc.dValor = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                        Case 3: objCarneParc.iFilialEmpresa = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                        Case 4: objCarneParc.iParcela = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                        Case 5: objCarneParc.iStatus = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                        Case 6: objCarneParc.lNumIntCarne = StrParaLong(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                        Case 7: objCarneParc.lNumIntDoc = StrParaLong(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                        Case 8: Exit Do
'                    End Select
'
'                    'Atualiza as Posições
'                    iPosInicio = iPosMeio + 1
'                    iPosMeio = (InStr(iPosInicio, sRegistro, Chr(vbKeyEscape)))
'
'                    If iPosMeio > iPosColInicio Then iPosMeio = iPosColInicio
'
'                Loop
'                objVenda1.objCarne.colParcelas.Add objCarneParc
'            Loop
'
'            iPosInicio = iPosColFim + 1
'
'            'Procura o Primeiro Escape dentro da String
'            iPosMeio = (InStr(iPosInicio, sRegistro, Chr(vbKeyEscape)))
'
'            'iPosicao1 Guarda a posição do quarto Control(referente ao colcheques)
'            iPosColFim = InStr(iPosInicio + 1, sRegistro, Chr(vbKeyControl))
'
'            'iPosicao1 Guarda a posição do inicio da colitens
'            iPosColInicio = InStr(iPosInicio + 1, sRegistro, Chr(vbKeyShift))
'            If iPosColInicio > iPosColFim Then iPosColInicio = iPosColFim
'
'            iIndice = 0
'
'            Do While iPosInicio <= iPosMeio
'
'                iIndice = iIndice + 1
'                'Recolhe o objCupom
'                Select Case iIndice
'
'                    Case 1: objVenda1.objCupomFiscal.dtDataEmissao = StrParaDate(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                    Case 2: objVenda1.objCupomFiscal.dHoraEmissao = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                    Case 3: objVenda1.objCupomFiscal.dValorAcrescimo = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                    Case 4: objVenda1.objCupomFiscal.dValorDesconto = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                    Case 5: objVenda1.objCupomFiscal.dValorProdutos = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                    Case 6: objVenda1.objCupomFiscal.dValorTotal = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                    Case 7: objVenda1.objCupomFiscal.dValorTroco = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                    Case 8: objVenda1.objCupomFiscal.iECF = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                    Case 9: objVenda1.objCupomFiscal.iFilialEmpresa = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                    Case 10: objVenda1.objCupomFiscal.iStatus = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                    Case 11: objVenda1.objCupomFiscal.iTabelaPreco = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                    Case 12: objVenda1.objCupomFiscal.lGerenteCancel = StrParaLong(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                    Case 13: objVenda1.objCupomFiscal.lNumero = StrParaLong(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                    Case 14: objVenda1.objCupomFiscal.lNumIntDoc = StrParaLong(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                    Case 15: objVenda1.objCupomFiscal.lNumOrcamento = StrParaLong(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                    Case 16: objVenda1.objCupomFiscal.iVendedor = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                    Case 17: objVenda1.objCupomFiscal.sCPFCGC = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
'                    Case 18: objVenda1.objCupomFiscal.sMotivoCancel = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
'                    Case 19: objVenda1.objCupomFiscal.sNaturezaOp = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
'                    Case 20: Exit Do
'
'                End Select
'
'                'Atualiza as Posições
'                iPosInicio = iPosMeio + 1
'                iPosMeio = (InStr(iPosInicio, sRegistro, Chr(vbKeyEscape)))
'
'                If iPosMeio > iPosColInicio Then iPosMeio = iPosColInicio
'
'            Loop
'
'            Do While iPosInicio < iPosColFim
'
'                iPosInicio = iPosColInicio + 1
'
'                'Procura o Primeiro Escape dentro da String
'                iPosMeio = (InStr(iPosInicio, sRegistro, Chr(vbKeyEscape)))
'
'                'Procura um shift dentro da String
'                iPosColInicio = (InStr(iPosInicio, sRegistro, Chr(vbKeyShift)))
'                If iPosColInicio = 0 Or iPosColInicio > iPosColFim Then iPosColInicio = iPosColFim
'
'                iIndice = 0
'                Set objItens = New ClassItemCupomFiscal
'
'                Do While iPosInicio <= iPosMeio
'
'                    iIndice = iIndice + 1
'                    'Recolhe os itens
'                    Select Case iIndice
'
'                        Case 1: objItens.dAliquotaICMS = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                        Case 2: objItens.dPercDesc = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                        Case 3: objItens.dPrecoUnitario = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                        Case 4: objItens.dQuantidade = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                        Case 5: objItens.dValorDesconto = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                        Case 6: objItens.icancel = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                        Case 7: objItens.iStatus = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                        Case 8: objItens.lNumIntCupom = StrParaLong(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                        Case 9: objItens.lNumIntDoc = StrParaLong(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                        Case 10: objItens.sProduto = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
'                        Case 11: objItens.sUnidadeMed = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
'                        Case 12: Exit Do
'                    End Select
'
'                    'Atualiza as Posições
'                    iPosInicio = iPosMeio + 1
'                    iPosMeio = (InStr(iPosInicio, sRegistro, Chr(vbKeyEscape)))
'
'                    If iPosMeio > iPosColInicio Then iPosMeio = iPosColInicio
'
'                Loop
'                objVenda1.objCupomFiscal.colItens.Add objItens
'            Loop
'
'            iPosInicio = iPosColFim + 1
'
'            'iPosicao1 Guarda a posição do quinto Control(referente ao colmovcaixa)
'            iPosColFim = InStr(iPosInicio, sRegistro, Chr(vbKeySeparator))
'
'            Do While iPosInicio < iPosColFim
'
'                'Procura o Primeiro Escape dentro da String
'                iPosMeio = (InStr(iPosInicio, sRegistro, Chr(vbKeyEscape)))
'
'                'Procura um control dentro da String
'                iPosColInicio = (InStr(iPosInicio, sRegistro, Chr(vbKeyControl)))
'                If iPosColInicio = 0 Or iPosColInicio > iPosColFim Then iPosColInicio = iPosColFim
'
'                iIndice = 0
'                Set objMovCx = New ClassMovimentoCaixa
'
'                Do While iPosInicio <= iPosMeio
'
'                    iIndice = iIndice + 1
'                    'Recolhe os Movimentos de caixa
'                    Select Case iIndice
'
'                        Case 1: objMovCx.dHora = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                        Case 2: objMovCx.dtDataMovimento = StrParaDate(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                        Case 3: objMovCx.dValor = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                        Case 4: objMovCx.iAdmMeioPagto = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                        Case 5: objMovCx.iCaixa = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                        Case 6: objMovCx.iCodConta = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                        Case 7: objMovCx.iCodOperador = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                        Case 8: objMovCx.iFilialEmpresa = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                        Case 9: objMovCx.iGerente = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                        Case 10: objMovCx.iParcelamento = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                        Case 11: objMovCx.iTipo = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                        Case 12: objMovCx.iTipoCartao = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                        Case 13: objMovCx.lCupomFiscal = StrParaLong(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                        Case 14: objMovCx.lMovtoEstorno = StrParaLong(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                        Case 15: objMovCx.lMovtoTransf = StrParaLong(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                        Case 16: objMovCx.lNumero = StrParaLong(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                        Case 17: objMovCx.lNumMovto = StrParaLong(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                        Case 18: objMovCx.lNumRefInterna = StrParaLong(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                        Case 19: objMovCx.lSequencial = StrParaLong(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                        Case 20: objMovCx.lSequencialConta = StrParaLong(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                        Case 21: objMovCx.sFavorecido = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
'                        Case 22: objMovCx.sHistorico = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
'                        Case 23: Exit Do
'
'                    End Select
'
'                    'Atualiza as Posições
'                    iPosInicio = iPosMeio + 1
'                    iPosMeio = (InStr(iPosInicio, sRegistro, Chr(vbKeyEscape)))
'
'                    If iPosMeio > iPosColInicio Then iPosMeio = iPosColInicio
'
'                Loop
'
'                iPosInicio = iPosColInicio + 1
'
'                objVenda1.colMovimentosCaixa.Add objMovCx
'            Loop
'
'            iPosInicio = iPosColFim + 1
'
'            'iPosicao1 Guarda a posição do quinto Control(referente ao colmovcaixa)
'            iPosColFim = InStr(iPosInicio, sRegistro, Chr(vbKeySeparator))
'            If iPosColFim = 0 Then iPosColFim = iPosFim
'
'            iPosInicio = iPosInicio + 1
'
'            Do While iPosInicio < iPosColFim
'
'                'Procura o Primeiro Escape dentro da String
'                iPosMeio = (InStr(iPosInicio, sRegistro, Chr(vbKeyEscape)))
'
'                'Procura um control dentro da String
'                iPosColInicio = (InStr(iPosInicio, sRegistro, Chr(vbKeyControl)))
'                If iPosColInicio = 0 Or iPosColInicio > iPosColFim Then iPosColInicio = iPosColFim
'
'                iIndice = 0
'                Set objCheque = New ClassChequePre
'
'                Do While iPosInicio < iPosMeio
'
'                    iIndice = iIndice + 1
'                    'Recolhe os Cheques
'                    Select Case iIndice
'
'                        Case 1: objCheque.dtDataDeposito = StrParaDate(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                        Case 2: objCheque.dValor = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                        Case 3: objCheque.iAprovado = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                        Case 4: objCheque.iBanco = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                        Case 5: objCheque.iFilialEmpresa = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                        Case 6: objCheque.iFilialEmpresaLoja = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                        Case 7: objCheque.iNaoEspecificado = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                        Case 8: objCheque.lNumero = StrParaLong(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                        Case 9: objCheque.lNumIntCheque = StrParaLong(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                        Case 10: objCheque.sAgencia = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
'                        Case 11: objCheque.sContaCorrente = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
'                        Case 12: objCheque.sCPFCGC = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
'                        Case 13: objCheque.lSequencialCaixa = StrParaLong(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                        Case 14: Exit Do
'
'                    End Select
'
'                    'Atualiza as Posições
'                    iPosInicio = iPosMeio + 1
'                    iPosMeio = (InStr(iPosInicio, sRegistro, Chr(vbKeyEscape)))
'
'                    If iPosMeio > iPosColInicio Or iPosMeio = 0 Then iPosMeio = iPosColInicio
'
'                Loop
'
'                iPosInicio = iPosColInicio + 1
'
'                objVenda1.colCheques.Add objCheque
'            Loop
'
'            iPosInicio = iPosColFim + 2
'
'            Do While iPosInicio < iPosFim
'
'                'Procura o Primeiro Escape dentro da String
'                iPosMeio = (InStr(iPosInicio, sRegistro, Chr(vbKeyEscape)))
'
'                'Procura um control dentro da String
'                iPosColInicio = (InStr(iPosInicio, sRegistro, Chr(vbKeyControl)))
'                If iPosColInicio = 0 Then iPosColInicio = iPosFim
'
'                iIndice = 0
'                Set objTroca = New ClassTroca
'
'                Do While iPosInicio < iPosMeio
'
'                    iIndice = iIndice + 1
'                    'Recolhe as trocas
'                    Select Case iIndice
'
'                        Case 1: objTroca.dQuantidade = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                        Case 2: objTroca.dValor = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                        Case 3: objTroca.iFilialEmpresa = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                        Case 4: objTroca.lNumIntDoc = StrParaLong(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                        Case 5: objTroca.sProduto = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
'                        Case 6: Exit Do
'
'                    End Select
'
'                    'Atualiza as Posições
'                    iPosInicio = iPosMeio + 1
'                    iPosMeio = (InStr(iPosInicio, sRegistro, Chr(vbKeyEscape)))
'
'                    If iPosMeio > iPosColInicio Or iPosMeio = 0 Then iPosMeio = iPosColInicio
'
'                Loop
'
'                iPosInicio = iPosColInicio + 1
'
'                objVenda1.colTroca.Add objTroca
'            Loop
'
'        End If
'
'    Loop
'
'    Close #1
'
'    Desmembrar_ECF = SUCESSO
'
'    Exit Function
'
'Erro_Desmembrar_ECF:
'
'    Desmembrar_ECF = gErr
'
'    Close #1
'
'    Select Case gErr
'
'        Case Else
'            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 149367)
'
'    End Select
'
'    Exit Function
'
'End Function

Public Function Retorna_MSGErro_AFRAC(ByVal lCodigo As Long, ByVal sMensagem As String) As Long

Dim lErro As Long
Dim iAcao As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim sCodigoErro As String
    
On Error GoTo Erro_Retorna_MSGErro_AFRAC
    
    If lCodigo <> 0 Then
    
        sCodigoErro = CStr(lCodigo)
        
        lErro = AFRAC_PegarCodigoErro(sCodigoErro, sMensagem, iAcao)
        If iAcao = 0 Then gError 99895
    
    End If
    
    Retorna_MSGErro_AFRAC = SUCESSO
    
    Exit Function

Erro_Retorna_MSGErro_AFRAC:
    
    Retorna_MSGErro_AFRAC = gErr
    
    Select Case gErr
                
        Case 99895
                            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149368)

    End Select
    
    Exit Function

End Function

Public Function Abre_Cupom(objVenda As ClassVenda) As Long

Dim lErro As Long
Dim sRetorno As String
Dim sCCF As String
Dim sData As String
Dim sHora As String
Dim dtDataSistema As Date
Dim dtHoraSistema As Date
Dim dtDataECF As Date
Dim dtHoraECF As Date

On Error GoTo Erro_Abre_Cupom

    
    dtHoraSistema = Time
    dtDataSistema = Date
    
    
    'verifica se ha diferenca de data, hora > 1 hora - PAFECF
    lErro = AFRAC_DataHoraImpressora(sData, sHora)
    If lErro <> SUCESSO Then
        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Abre Cupom")
        If lErro <> SUCESSO Then gError 204509
    End If

    dtDataECF = left(sData, 2) & "/" & Mid(sData, 3, 2) & "/" & right(sData, 2)

    If dtDataECF <> dtDataSistema Then gError 204510

    dtHoraECF = left(sHora, 2) & ":" & Mid(sHora, 3, 2) & ":" & right(sHora, 2)

    If Abs(DateDiff("n", dtHoraSistema, dtHoraECF)) > 60 Then gError 204511

    'Abre o cupom
    lErro = AFRAC_AbrirCupom(objVenda.objCupomFiscal.sCPFCGC1, objVenda.objCupomFiscal.sNomeCliente, objVenda.objCupomFiscal.sEndereco)
    If lErro <> SUCESSO And lErro <> 2 Then
        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Abre Cupom")
        If lErro <> SUCESSO Then gError 99812
    ElseIf lErro = 2 Then
        gError 214505
        
    End If
    
    'Atualiza o arquivo
    lErro = WritePrivateProfileString(APLICACAO_ECF, "CupomAberto", "1", NOME_ARQUIVO_CAIXA)
    If lErro = 0 Then gError 117531
            
    'funcao que cria/altera a chave e grava no caixaconfig.ini
    lErro = CF_ECF("CaixaConfig_Chave")
    If lErro <> SUCESSO Then gError 204293
            
    'cupom fiscal (número) no Afrac
    lErro = AFRAC_LerInformacaoImpressora("023", sRetorno)
    If lErro <> SUCESSO Then
        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Informação Impressora")
        If lErro <> SUCESSO Then gError 99816
    End If
           
    objVenda.objCupomFiscal.lNumero = StrParaLong(sRetorno)
           
    'le o CCF associado ao cupom fiscal
    lErro = AFRAC_CCF(sCCF)
    If lErro <> SUCESSO Then
        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "CCF")
        If lErro <> SUCESSO Then gError 204508
    End If
           
    objVenda.objCupomFiscal.lCCF = StrParaLong(sCCF)
    
    objVenda.objCupomFiscal.dtDataEmissao = Date
    objVenda.objCupomFiscal.dHoraEmissao = CDbl(Time)
    objVenda.objCupomFiscal.dtDataReducao = gdtDataAnterior

    
    Abre_Cupom = SUCESSO
    
    Exit Function

Erro_Abre_Cupom:
    
    Abre_Cupom = gErr
    
    Select Case gErr
                
        Case 99812, 99816, 204293, 204508, 204509, 214505
                                        
        Case 117531
            Call Rotina_ErroECF(vbOKOnly, ERRO_ARQUIVO_NAO_ENCONTRADO1, gErr, APLICACAO_ECF, "CupomAberto", NOME_ARQUIVO_CAIXA)
                                        
        Case 204510
            Call Rotina_ErroECF(vbOKOnly, ERRO_DATAS_DIFEREM, gErr, dtDataECF, dtDataSistema)
            
        Case 204511
            Call Rotina_ErroECF(vbOKOnly, ERRO_HORAS_DIFEREM, gErr, dtHoraECF, dtHoraSistema)
                                        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149369)

    End Select
    
    Exit Function
    
End Function

Function MovimentoCaixa_Move_Dados_Memoria(ByVal iTipoMovCaixa As Integer, ByVal objMovimentoCaixa As ClassMovimentoCaixa) As Long
'Função que move os Dados para o objMovimentosCaixa , Dados que Serão Usados no Movimento de Caixa

Dim lErro As Long

On Error GoTo Erro_MovimentoCaixa_Move_Dados_Memoria

    'Guarda as Informações Globais no objMovimentosCaixa
    
    objMovimentoCaixa.iFilialEmpresa = giFilialEmpresa

    objMovimentoCaixa.iCaixa = giCodCaixa

    objMovimentoCaixa.iCodOperador = giCodOperador

    objMovimentoCaixa.iTipo = iTipoMovCaixa

    objMovimentoCaixa.dtDataMovimento = Date
    
    objMovimentoCaixa.dHora = CDbl(Time)

    MovimentoCaixa_Move_Dados_Memoria = SUCESSO

    Exit Function

Erro_MovimentoCaixa_Move_Dados_Memoria:

    MovimentoCaixa_Move_Dados_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149370)

    End Select

    Exit Function

End Function

Function Caixa_Executa_LeituraX_EmTrans(lSequencial As Long, Optional iCodGerente As Integer) As Long
'Função que Executa a Função de LeituraX

Dim lErro As Long
Dim objMovimentosCaixa As New ClassMovimentoCaixa
Dim sLog As String
Dim colRegistro As New Collection

On Error GoTo Erro_Caixa_Executa_LeituraX_EmTrans

    If Not AFRAC_ImpressoraCFe(giCodModeloECF) Then
        
        If giOrcamentoECF <> CAIXA_SO_ORCAMENTO Then
    
            'Função que move os dados para a memoria
            lErro = CF_ECF("MovimentoCaixa_Move_Dados_Memoria", MOVIMENTO_CAIXA_LEITURA_X, objMovimentosCaixa)
            If lErro <> SUCESSO Then gError 99838
        
            'Guarda o lSequencial no objMovimentosCaixa
            objMovimentosCaixa.lSequencial = lSequencial
        
            'Preencher o campo Gerente
            objMovimentosCaixa.iGerente = iCodGerente
            
            Call CF_ECF("MovimentoCaixa_Gera_Log", colRegistro, objMovimentosCaixa)
            
            'Função que Grava as Informações no Arquivo
            lErro = CF_ECF("MovimentoCaixaECF_Grava", colRegistro)
            If lErro <> SUCESSO Then gError 99839
            
            Set colRegistro = New Collection
            
            'Função que Executa a LeituraX
            lErro = AFRAC_LeituraX()
            lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Leitura X")
            If lErro <> SUCESSO Then If lErro <> SUCESSO Then gError 99840
                
            'Adcionar o objMovimentoCaixa na Coleção Global
            gcolMovimentosCaixa.Add objMovimentosCaixa
    
            lErro = CF_ECF("Requisito_VII_Item8_E3")
            If lErro <> SUCESSO Then gError 214335
    
    
    
        
        End If
    
    End If
    
    Caixa_Executa_LeituraX_EmTrans = SUCESSO

    Exit Function

Erro_Caixa_Executa_LeituraX_EmTrans:

    Caixa_Executa_LeituraX_EmTrans = gErr

        Select Case gErr
        
            Case 99838 To 99840, 214335
                'Erros Tratdados Dentro da Função Chamadora

            Case Else
                Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149371)

    End Select

    Exit Function

End Function



Function Caixa_Executa_LeituraX(Optional dtData As Date) As Long
'Função Que Executa a LeituraX

Dim lErro As Long
Dim lSequencial As Long
Dim objOperador As New ClassOperador
Dim iCodGerente As Integer

On Error GoTo Erro_Caixa_Executa_LeituraX
    
    If Not AFRAC_ImpressoraCFe(giCodModeloECF) Then
        
        If giOrcamentoECF <> CAIXA_SO_ORCAMENTO Then
        
    '        If dtData = "00:00:00" Then dtData = gdtDataHoje
            
            'Se for Necessário a autorização do Gerente para abertura do Caixa
            If gobjLojaECF.iGerenteAutoriza = AUTORIZACAO_GERENTE Then
        
                'Chama a Tela de Senha
                Call Chama_TelaECF_Modal("OperadorLogin", objOperador, LOGIN_APENAS_GERENTE)
        
                'Sai de Função se a Tela de Login não Retornar ok
                If giRetornoTela <> vbOK Then gError 107628
        
                'Se o Operador for Gerente
                iCodGerente = objOperador.iCodigo
        
            End If
        
            'Função que Abre a Transação
            lErro = CF_ECF("Caixa_Transacao_Abrir", lSequencial)
            If lErro <> SUCESSO Then gError 107564

            'Função que Realiza a Leitura X
            lErro = CF_ECF("Caixa_Executa_LeituraX_EmTrans", lSequencial, iCodGerente)
            If lErro <> SUCESSO Then gError 107565
    
            'Função que Encerra a Transação
            lErro = CF_ECF("Caixa_Transacao_Fechar", lSequencial)
            If lErro <> SUCESSO Then gError 107566
        
        Else
            
            Call CF_ECF("Trata_Caixa_So_Orcamento")
            
        End If

    End If

    Caixa_Executa_LeituraX = SUCESSO

    Exit Function

Erro_Caixa_Executa_LeituraX:

    Caixa_Executa_LeituraX = gErr

    Select Case gErr

        Case 107564, 107565, 107566, 107628
            'Error Tratados Dentro da Função

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149372)

    End Select

'    lSequencial = glSeqTransacaoAberta
            
    'Rollback na transação
    Call CF_ECF("Caixa_Transacao_Rollback", glTransacaoPAFECF)

    Exit Function

End Function

Function MemoriaFiscal_Executa_Leitura(ByVal iTipoLeitura As Integer, ByVal sDe As String, ByVal sAte As String, ByVal iTipo As Integer, ByVal iArquivo As Integer) As Long
'Função que Vai Chamar a Função da Afrac que vai Executar a Leitura de Memória Fiscal

Dim lErro As Long
Dim lSequencial As Long
Dim iCodGerente As Integer
Dim objOperador As New ClassOperador

On Error GoTo Erro_MemoriaFiscal_Executa_Leitura

    If Not AFRAC_ImpressoraCFe(giCodModeloECF) Then
        
        'Se for Necessário a autorização do Gerente para abertura do Caixa
        If gobjLojaECF.iGerenteAutoriza = AUTORIZACAO_GERENTE Then
    
            'Chama a Tela de Senha
            Call Chama_TelaECF_Modal("OperadorLogin", objOperador, LOGIN_APENAS_GERENTE)
    
            'Sai de Função se a Tela de Login não Retornar ok
            If giRetornoTela <> vbOK Then gError 107621
    
            'Se o Operador For Gerente
            iCodGerente = objOperador.iCodigo
    
        End If
    
        'Função que Abre a Transação
        lErro = CF_ECF("Caixa_Transacao_Abrir", lSequencial)
        If lErro <> SUCESSO Then gError 107616
    
        'Função que Referenciar a Função da Afrac que Vai Executar a Leitura
        lErro = CF_ECF("MemoriaFiscal_Executa_LeituraEmTrans", lSequencial, iCodGerente, iTipoLeitura, sDe, sAte, iTipo, iArquivo)
        If lErro <> SUCESSO Then gError 107617
    
        'Função que Encerra a Transação com o Caixa
        lErro = CF_ECF("Caixa_Transacao_Fechar", lSequencial)
        If lErro <> SUCESSO Then gError 107618

    End If

    MemoriaFiscal_Executa_Leitura = SUCESSO

    Exit Function

Erro_MemoriaFiscal_Executa_Leitura:

    MemoriaFiscal_Executa_Leitura = gErr

     Select Case gErr

        Case 107616, 107617, 107618, 107621
            'Erro Tratado Dentro da Função que Foi Chamada

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149373)

    End Select
    
'    lSequencial = glSeqTransacaoAberta
            
    'Rollback na transação
    Call CF_ECF("Caixa_Transacao_Rollback", glTransacaoPAFECF)

    Exit Function

End Function

Function MemoriaFiscal_Executa_LeituraEmTrans(lSequencial As Long, iCodGerente As Integer, ByVal iTipoLeitura As Integer, ByVal sDe As String, ByVal sAte As String, ByVal iTipo As Integer, ByVal iArquivo As Integer) As Long
'Função que Chama a Função da Afrac que Executa a Leitura da Memória Fiscal

Dim lErro As Long
Dim objMovimentoCaixa As New ClassMovimentoCaixa
Dim sCampoReducao1 As String
Dim sCampoReducao2 As String
Dim sLog As String
Dim colRegistro As New Collection

On Error GoTo Erro_MemoriaFiscal_Executa_LeituraEmTrans

    If Not AFRAC_ImpressoraCFe(giCodModeloECF) Then
        
        If giOrcamentoECF <> CAIXA_SO_ORCAMENTO Then
    
    
            'Função que Move os Dados para a Memória
            lErro = CF_ECF("MovimentoCaixa_Move_Dados_Memoria", MOVIMENTO_CAIXA_LEITURA_MEMORIA_FISCAL, objMovimentoCaixa)
            If lErro <> SUCESSO Then gError 107619
            
            'Guarda o Sequencial no objMovimentos Caixa
            objMovimentoCaixa.lSequencial = lSequencial
            
            ' Codigo Gerente
            objMovimentoCaixa.iGerente = iCodGerente
            
            Call CF_ECF("MovimentoCaixa_Gera_Log", colRegistro, objMovimentoCaixa)
            
            'Guarda as Informações no Arquivo
            lErro = CF_ECF("MovimentoCaixaECF_Grava", colRegistro)
            If lErro <> SUCESSO Then gError 107620
            
            Set colRegistro = New Collection
        
            'Verifica qual das CheckBox estão Marcadas
            If iTipoLeitura = LEITURA_COMPLETA Then
        
                'Considera-se como ambos
                lErro = AFRAC_EmitirLeituraMemoriaFiscal(iTipoLeitura, "", "", iTipo, iArquivo)
                lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Memória Fiscal")
                If lErro <> SUCESSO Then gError 107597
                
            ElseIf iTipoLeitura = LEITURA_DATAS Then
                          
                'Verifica se a Data Inicial é Maior que a segunda
                If StrParaDate(sDe) > StrParaDate(sAte) Then gError 107598
        
               lErro = AFRAC_EmitirLeituraMemoriaFiscal(iTipoLeitura, Format(sDe, "DDMMYYYY"), Format(sAte, "DDMMYYYY"), iTipo, iArquivo)
               lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Memória Fiscal")
               If lErro <> SUCESSO Then gError 107589
               
            ElseIf iTipoLeitura = LEITURA_REDUCOES Then
        
                'Verificar se as Reduções Estão Preenchidas se Erro
                If Len(Trim(sDe)) = 0 Or Len(Trim(sAte)) = 0 Then gError 107615
        
               If StrParaLong(sDe) > StrParaLong(sAte) Then gError 107593
        
                'Função Que Coloca Zero a Esqueda
                sCampoReducao1 = Formata_Campo(ALINHAMENTO_ESQUERDA, 4, "0", StrParaLong(sDe))
        
                'Função Que Coloca Zero a Esqueda
                sCampoReducao2 = Formata_Campo(ALINHAMENTO_ESQUERDA, 4, "0", StrParaLong(sAte))
        
                'Se for por Intervalo de Redução
                lErro = AFRAC_EmitirLeituraMemoriaFiscal(iTipoLeitura, sCampoReducao1, sCampoReducao2, iTipo, iArquivo)
                lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Memória Fiscal")
                If lErro <> SUCESSO Then gError 107590
                
        '    ElseIf iTipoLeitura = LEITURA_DATAS_SIMPLES Then
        '
        '        'Verifica se a Data Inicial é Maior que a segunda
        '        If StrParaDate(sDe) > StrParaDate(sAte) Then gError 204395
        '
        '       lErro = AFRAC_EmitirLeituraMemoriaFiscal(LEITURA_MEMORIAFISCAL_TIPO_DATAS, Format(sDe, "DDMMYYYY"), Format(sAte, "DDMMYYYY"), LEITURA_SIMPLES, iArquivo)
        '       lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Memória Fiscal")
        '       If lErro <> SUCESSO Then gError 204396
        '
        '    ElseIf iTipoLeitura = LEITURA_REDUCOES_SIMPLES Then
        '
        '        'Verificar se as Reduções Estão Preenchidas se Erro
        '        If Len(Trim(sDe)) = 0 Or Len(Trim(sAte)) = 0 Then gError 204397
        '
        '       If StrParaLong(sDe) > StrParaLong(sAte) Then gError 204398
        '
        '        'Função Que Coloca Zero a Esqueda
        '        sCampoReducao1 = Formata_Campo(ALINHAMENTO_ESQUERDA, 4, "0", StrParaLong(sDe))
        '
        '        'Função Que Coloca Zero a Esqueda
        '        sCampoReducao2 = Formata_Campo(ALINHAMENTO_ESQUERDA, 4, "0", StrParaLong(sAte))
        '
        '        'Se for por Intervalo de Redução
        '        lErro = AFRAC_EmitirLeituraMemoriaFiscal(LEITURA_MEMORIAFISCAL_TIPO_REDUCAO, sCampoReducao1, sCampoReducao2, LEITURA_SIMPLES, iArquivo)
        '        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Memória Fiscal")
        '        If lErro <> SUCESSO Then gError 204399
                
                
                
            End If
    
        
            lErro = CF_ECF("Requisito_VII_Item8_E3")
            If lErro <> SUCESSO Then gError 214343
    
        
        
        Else
            
            Call CF_ECF("Trata_Caixa_So_Orcamento")
            
        End If

    End If

    MemoriaFiscal_Executa_LeituraEmTrans = SUCESSO
    
    Exit Function

Erro_MemoriaFiscal_Executa_LeituraEmTrans:

    MemoriaFiscal_Executa_LeituraEmTrans = gErr

        Select Case gErr
        
        Case 107589, 107590, 204396, 204399, 214343
            'Erro Tratados Dentro da Função Chamada

        Case 107593, 204398
            Call Rotina_ErroECF(vbOKOnly, ERRO_NUMERO_INTERVALO_REDUCAO_INICIAL_MAIOR, gErr)

        Case 107597
            'Erro Tratado Dentro da Função

        Case 107598, 204395
            Call Rotina_ErroECF(vbOKOnly, ERRO_DATA_INICIAL_MAIOR1, gErr)

        Case 107614
            Call Rotina_ErroECF(vbOKOnly, ERRO_DATAS_MEMORIAFISCAL_NAO_PREENCHIDA, gErr)

        Case 107615, 204397
            Call Rotina_ErroECF(vbOKOnly, ERRO_REDUCAO_NAO_PREENCHIDA, gErr)
        
        Case 107619, 107620
            'Erro Tratado Dentro da Função

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149374)

    End Select

    Exit Function

End Function

Function Caixa_MovimentoCaixa_Le_NumMovto(colMovimentosCaixa As Collection, ByVal lNumMovto As Long) As Long
'Função que carrega a Coleção de Movimentos de Caixa, com informação de MOVTO's Caixa Global ao Sistema

Dim lErro As Long
Dim objMovimentoCaixa As ClassMovimentoCaixa
Dim bAchou As Boolean
Dim objMovCx As ClassMovimentoCaixa

On Error GoTo Erro_Caixa_MovimentoCaixa_Le_NumMovto
    
    bAchou = False
    For Each objMovimentoCaixa In gcolMovimentosCaixa

        'Seleciona o Movimento de Caixa com numero do Movto passado por parâmetro
        If lNumMovto = objMovimentoCaixa.lNumMovto Then

            Set objMovCx = New ClassMovimentoCaixa

            lErro = CF_ECF("MovimentoCaixa_Copia", objMovimentoCaixa, objMovCx)
            If lErro <> SUCESSO Then gError 105702

            'Guarda  na Coleção passada por parâmetro
            colMovimentosCaixa.Add objMovCx

            bAchou = True
            
        End If

    Next
    'Se não existir o movimento com o Codigo passado como parâmetre
    If bAchou = False Then gError 107850

    Caixa_MovimentoCaixa_Le_NumMovto = SUCESSO

    Exit Function

Erro_Caixa_MovimentoCaixa_Le_NumMovto:

    Caixa_MovimentoCaixa_Le_NumMovto = gErr

    Select Case gErr

        Case 105702, 107850
    
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149375)

    End Select

    Exit Function

End Function

Public Function Caixa_Carne_Imprime_ECF(ByVal colCarneParcelas As Collection) As Long

    Caixa_Carne_Imprime_ECF = SUCESSO
    
    Exit Function
    
End Function

Function Caixa_Cliente_Le_NomeReduzido(objCliente As ClassCliente) As Long

Dim objCliente2 As ClassCliente
Dim bAchou As Boolean
Dim lErro As Long

On Error GoTo Erro_Caixa_Cliente_Le_NomeReduzido

'    For Each objCliente2 In gcolCliente
'        If objCliente2.sNomeReduzido = objCliente.sNomeReduzido Then
'            Set objCliente = objCliente2
'            bAchou = True
'            Exit For
'        End If
'    Next
'
'    'Cliente não encontrado
'    If Not (bAchou) Then gError 109540
    
    lErro = CF_ECF("Clientes_Le_NomeReduzido", objCliente.sNomeReduzido, objCliente)
    If lErro <> SUCESSO And lErro <> 214878 Then gError 214901
    
    If lErro <> SUCESSO Then gError 109540
    
    
    Exit Function
    
Erro_Caixa_Cliente_Le_NomeReduzido:

    Select Case gErr
                
        Case 109540
            lErro = Rotina_ErroECF(vbOKOnly, ERRO_CLIENTE_NAO_EXISTENTE, gErr)
            
        Case 214901
            
        Case Else
            lErro = Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 149376)
            
    End Select
    
    Exit Function
    
End Function

Function Caixa_Cliente_Le_Codigo(objCliente As ClassCliente)

Dim objCliente2 As ClassCliente
Dim bAchou As Boolean
Dim lErro As Long

On Error GoTo Erro_Caixa_Cliente_Le_Codigo

'    For Each objCliente2 In gcolCliente
'        If objCliente2.lCodigo = objCliente.lCodigo Then
'            Set objCliente = objCliente2
'            bAchou = True
'            Exit For
'        End If
'    Next
'
'    'Cliente não encontrado
'    If Not (bAchou) Then gError 109550
    
    
    lErro = CF_ECF("Clientes_Le_Codigo", objCliente.lCodigo, objCliente)
    If lErro <> SUCESSO And lErro <> 214899 Then gError 214903
    
    If lErro <> SUCESSO Then gError 109550
    
    Exit Function
    
Erro_Caixa_Cliente_Le_Codigo:

    Select Case gErr
                
        Case 109550
            lErro = Rotina_ErroECF(vbOKOnly, ERRO_CLIENTE_NAO_EXISTENTE, gErr)
            
        Case 214903
            
        Case Else
            lErro = Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 149377)
            
    End Select
    
    Exit Function
    
End Function


Function Caixa_Cliente_Le_CPFCGC(objCliente As ClassCliente) As Long

Dim objCliente2 As ClassCliente
Dim bAchou As Boolean
Dim lErro As Long

On Error GoTo Erro_Caixa_Cliente_Le_CPFCGC

'    For Each objCliente2 In gcolCliente
'        If objCliente2.sCGC = objCliente.sCGC Then
'            Set objCliente = objCliente2
'            bAchou = True
'            Exit For
'        End If
'    Next
'
'    'Cliente não encontrado
'    If Not (bAchou) Then gError 109551
    
    
    lErro = CF_ECF("Clientes_Le_CPFCNPJ", objCliente.sCgc, objCliente)
    If lErro <> SUCESSO And lErro <> 214887 Then gError 214902
    
    If lErro <> SUCESSO Then gError 109551
    
    
    Exit Function
    
Erro_Caixa_Cliente_Le_CPFCGC:

    Select Case gErr
                
        Case 109551
            lErro = Rotina_ErroECF(vbOKOnly, ERRO_CLIENTE_NAO_EXISTENTE, gErr)
            
        Case 214902
            
        Case Else
            lErro = Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 149378)
            
    End Select
    
    Exit Function
    
End Function

Sub MovimentoBoleto_Gera_Log(colRegistro As Collection, objMovimentoCaixa As ClassMovimentoCaixa)

Dim lErro As Long
Dim iIndice As Integer
Dim cColInformação As New Collection
Dim iCont As Integer
Dim sTemp As String
Dim sLog As String

On Error GoTo Erro_MovimentoBoleto_Gera_Log

    'Monta o Log
    sTemp = TIPOREGISTROECF_MOVIMENTOCAIXA_BOLETO & Chr(vbKeyControl) & CStr(objMovimentoCaixa.iTipo) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.dHora) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.dtDataMovimento) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.iGerente) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.iCodOperador) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.lSequencial) & Chr(vbKeyEscape) & _
    CStr(objMovimentoCaixa.iFilialEmpresa) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.dValor) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.iAdmMeioPagto) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.iParcelamento) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.iCaixa) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.iTipoCartao) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.lNumMovto) & Chr(vbKeyEscape) & _
    CStr(objMovimentoCaixa.lNumRefInterna) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.dtDataPreDatado) & Chr(vbKeyEscape)
    
    If Len(Trim(sTemp)) + Len(Trim(sLog)) > MAX_TAM_STRING_MOVIMENTOCAIXA Then
        If Len(Trim(sLog)) <> 0 Then colRegistro.Add sLog
        sLog = ""
    End If
    sLog = sLog & sTemp
    sTemp = ""
    
    sTemp = sTemp & Chr(vbKeyEnd)
    
    If Len(Trim(sTemp)) + Len(Trim(sLog)) + 1 > MAX_TAM_STRING_MOVIMENTOCAIXA Then
        If Len(Trim(sLog)) <> 0 Then colRegistro.Add sLog
        sLog = ""
    End If
    sLog = sLog & sTemp
    sTemp = ""

    If Len(Trim(sLog)) <> 0 Then colRegistro.Add sLog
    
    Exit Sub

Erro_MovimentoBoleto_Gera_Log:

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149379)

    End Select

    Exit Sub

End Sub

Sub MovimentoOutros_Gera_Log(colRegistro As Collection, objMovimentoCaixa As ClassMovimentoCaixa)

Dim lErro As Long
Dim iIndice As Integer
Dim cColInformação As New Collection
Dim iCont As Integer
Dim sLog As String
Dim sTemp As String

On Error GoTo Erro_MovimentoOutros_Gera_Log

    'Monta o Log
    sTemp = TIPOREGISTROECF_MOVIMENTOCAIXA_OUTROS & Chr(vbKeyControl) & CStr(objMovimentoCaixa.iTipo) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.dHora) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.dtDataMovimento) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.iGerente) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.iCodOperador) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.lSequencial) & Chr(vbKeyEscape) & _
    CStr(objMovimentoCaixa.iFilialEmpresa) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.dValor) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.iAdmMeioPagto) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.iParcelamento) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.iCaixa) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.lNumMovto) & Chr(vbKeyEscape)
        
    If Len(Trim(sTemp)) + Len(Trim(sLog)) > MAX_TAM_STRING_MOVIMENTOCAIXA Then
        If Len(Trim(sLog)) <> 0 Then colRegistro.Add sLog
        sLog = ""
    End If
    sLog = sLog & sTemp
    sTemp = ""
    
    sTemp = sTemp & Chr(vbKeyEnd)
        
    If Len(Trim(sTemp)) + Len(Trim(sLog)) > MAX_TAM_STRING_MOVIMENTOCAIXA Then
        If Len(Trim(sLog)) <> 0 Then colRegistro.Add sLog
        sLog = ""
    End If
    sLog = sLog & sTemp
    sTemp = ""
    
    If Len(Trim(sLog)) <> 0 Then colRegistro.Add sLog
    
    Exit Sub

Erro_MovimentoOutros_Gera_Log:

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149380)

    End Select

    Exit Sub

End Sub

Sub MovimentoDinheiro_Gera_Log(colRegistro As Collection, objMovimentoCaixa As ClassMovimentoCaixa)

Dim lErro As Long
Dim iIndice As Integer
Dim iCont As Integer
Dim cColInformação As New Collection
Dim sLog As String
Dim sTemp As String

On Error GoTo Erro_MovimentoDinheiro_Gera_Log

    'Monta o Log
    sTemp = TIPOREGISTROECF_MOVIMENTOCAIXA_DINHEIRO & Chr(vbKeyControl) & CStr(objMovimentoCaixa.iTipo) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.dHora) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.dtDataMovimento) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.iGerente) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.iCodOperador) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.lSequencial) & Chr(vbKeyEscape) & _
    CStr(objMovimentoCaixa.iFilialEmpresa) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.dValor) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.iAdmMeioPagto) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.iParcelamento) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.iCaixa) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.lNumMovto) & Chr(vbKeyEscape)
    
    sTemp = sTemp & Chr(vbKeyEnd)
    
    colRegistro.Add sTemp
    
    Exit Sub

Erro_MovimentoDinheiro_Gera_Log:

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149381)

    End Select

    Exit Sub

End Sub

Sub MovimentoCheque_Gera_Log(colRegistro As Collection, objMovimentoCaixa As ClassMovimentoCaixa, colInfoComplementar As Collection)

Dim lErro As Long
Dim iIndice As Integer
Dim colInfoComplementarAux As Collection
Dim iCont As Integer
Dim sLog As String
Dim sTemp As String

On Error GoTo Erro_MovimentoCheque_Gera_Log

    'Monta o Log
    sTemp = TIPOREGISTROECF_MOVIMENTOCAIXA_CHEQUES & Chr(vbKeyControl) & CStr(objMovimentoCaixa.iTipo) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.dHora) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.dtDataMovimento) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.iGerente) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.iCodOperador) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.lSequencial) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.iFilialEmpresa) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.dValor) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.iAdmMeioPagto) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.iParcelamento) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.iCaixa) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.lNumMovto) & Chr(vbKeyEscape)
    
    'Varre a Coleção
    For iIndice = 1 To colInfoComplementar.Count
                
        sTemp = sTemp & CStr(colInfoComplementar.Item(iIndice)) & Chr(vbKeyShift)
        
        If Len(Trim(sTemp)) + Len(Trim(sLog)) + 1 > MAX_TAM_STRING_MOVIMENTOCAIXA Then
            If Len(Trim(sLog)) <> 0 Then colRegistro.Add sLog
            sLog = ""
        End If
        sLog = sLog & sTemp
        sTemp = ""
    Next
        
    sTemp = sTemp & Chr(vbKeyEnd)
                    
    If Len(Trim(sTemp)) + Len(Trim(sLog)) > MAX_TAM_STRING_MOVIMENTOCAIXA Then
        If Len(Trim(sLog)) <> 0 Then colRegistro.Add sLog
        sLog = ""
    End If
    sLog = sLog & sTemp
    sTemp = ""
                
    If Len(Trim(sLog)) > 0 Then colRegistro.Add sLog

    Exit Sub

Erro_MovimentoCheque_Gera_Log:

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149382)

    End Select

    Exit Sub

End Sub

Sub MovimentoCheque_Gera_Log_1(colRegistro As Collection, objMovimentoCaixa As ClassMovimentoCaixa, colCheque As Collection)

Dim lErro As Long
Dim iIndice As Integer
Dim colInfoComplementarAux As Collection
Dim iCont As Integer
Dim sLog As String
Dim sTemp As String
Dim objCheque As ClassChequePre

On Error GoTo Erro_MovimentoCheque_Gera_Log_1

    'Monta o Log
    sTemp = TIPOREGISTROECF_MOVIMENTOCAIXA_CHEQUES & Chr(vbKeyControl) & CStr(objMovimentoCaixa.iTipo) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.dHora) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.dtDataMovimento) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.iGerente) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.iCodOperador) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.lSequencial) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.iFilialEmpresa) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.dValor) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.iAdmMeioPagto) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.iParcelamento) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.iCaixa) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.lNumMovto) & Chr(vbKeyEscape)
    
    'Jogo todos os colcheques na string
    For Each objCheque In colCheque
        sTemp = sTemp & Chr(vbKeyControl) & objCheque.dtDataDeposito & Chr(vbKeyEscape) & objCheque.dValor & Chr(vbKeyEscape) & objCheque.iAprovado & Chr(vbKeyEscape) & objCheque.iBanco & Chr(vbKeyEscape) & objCheque.iFilialEmpresaLoja & Chr(vbKeyEscape) & objCheque.iNaoEspecificado & Chr(vbKeyEscape) & objCheque.lNumero & Chr(vbKeyEscape) & objCheque.sAgencia & Chr(vbKeyEscape) & objCheque.sContaCorrente & Chr(vbKeyEscape) & objCheque.sCPFCGC & Chr(vbKeyEscape) & objCheque.lSequencialCaixa & Chr(vbKeyEscape)
        
        If Len(Trim(sTemp)) + Len(Trim(sLog)) + 1 > MAX_TAM_STRING_MOVIMENTOCAIXA Then
            If Len(Trim(sLog)) <> 0 Then colRegistro.Add sLog
            sLog = ""
        End If
        sLog = sLog & sTemp
        sTemp = ""
        
    Next

    sTemp = sTemp & Chr(vbKeyEnd)
                    
    If Len(Trim(sTemp)) + Len(Trim(sLog)) > MAX_TAM_STRING_MOVIMENTOCAIXA Then
        If Len(Trim(sLog)) <> 0 Then colRegistro.Add sLog
        sLog = ""
    End If
    sLog = sLog & sTemp
    sTemp = ""
                
    If Len(Trim(sLog)) > 0 Then colRegistro.Add sLog

    Exit Sub

Erro_MovimentoCheque_Gera_Log_1:

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149383)

    End Select

    Exit Sub

End Sub

Sub MovimentoTicket_Gera_Log(colRegistro As Collection, objMovimentoCaixa As ClassMovimentoCaixa)

Dim lErro As Long
Dim iIndice As Integer
Dim cColInformação As New Collection
Dim iCont As Integer
Dim sLog As String
Dim sTemp As String

On Error GoTo Erro_MovimentoTicket_Gera_Log

    'Monta o Log
    sTemp = TIPOREGISTROECF_MOVIMENTOCAIXA_TICKETS & Chr(vbKeyControl) & CStr(objMovimentoCaixa.iTipo) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.dHora) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.dtDataMovimento) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.iGerente) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.iCodOperador) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.lSequencial) & Chr(vbKeyEscape) & _
    CStr(objMovimentoCaixa.iFilialEmpresa) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.dValor) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.iAdmMeioPagto) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.iParcelamento) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.iCaixa) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.lNumMovto) & Chr(vbKeyEscape)
    
    If Len(Trim(sTemp)) + Len(Trim(sLog)) > MAX_TAM_STRING_MOVIMENTOCAIXA Then
        If Len(Trim(sLog)) <> 0 Then colRegistro.Add sLog
        sLog = ""
    End If
    sLog = sLog & sTemp
    sTemp = ""

    sTemp = sTemp & Chr(vbKeyEnd)
    
    If Len(Trim(sTemp)) + Len(Trim(sLog)) > MAX_TAM_STRING_MOVIMENTOCAIXA Then
        If Len(Trim(sLog)) <> 0 Then colRegistro.Add sLog
        sLog = ""
    End If
    sLog = sLog & sTemp
    sTemp = ""
    
    If Len(Trim(sLog)) <> 0 Then colRegistro.Add sLog
    
    Exit Sub

Erro_MovimentoTicket_Gera_Log:

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149384)

    End Select

    Exit Sub

End Sub

Function Desmembra_MovimentosCaixa(ByVal colMovimentosCaixa As Collection, ByVal colImfCompl As Collection, sTipoCaixa As String) As Long
'Função que Desmembra o Movimentos de Caixa e Carrega em gcolMovimentosCaixa

Dim lErro As Long
Dim objMovimentosCaixa As ClassMovimentoCaixa
Dim iPosInicio As Integer
Dim sRegistro As String
Dim iPos As Integer
Dim iPosFinal As Integer
Dim iPosicao3  As Integer
Dim iPosShift As Integer
Dim iIndice As Integer
Dim sNomeArq As String
Dim sTipo As String
Dim iPosEnd As Integer
Dim iCont As Integer
Dim iPosInicioShift As Integer
Dim lConteudo As Long
Dim iInicial As Integer

On Error GoTo Erro_Desmembra_MovimentosCaixa

    'Desmembra o Arquivo do dia Corrente
    sNomeArq = gsDirMVTEF & "MV" & CStr(Format(Date, "ddmmyy")) & (".txt")
    
    Open sNomeArq For Input As #1
    
    Do While Not EOF(1)
         
        iInicial = 1
             
        'Primeira Posição
        iPosInicio = 1
        
        'Inicializa a variavel
        iIndice = 0
        
        'Instancia o Obj da ClassMovimentoCaixa
        Set objMovimentosCaixa = New ClassMovimentoCaixa
    
        'Busca o próximo registro do arquivo
        Line Input #1, sRegistro
        
        'Posição Final
        iPosEnd = InStr(iPosInicio, sRegistro, Chr(vbKeyEnd))
        
        'Procura o Primeiro Control para saber o tipo do registro
        iPos = InStr(iPosInicio, sRegistro, Chr(vbKeyControl))
        
        'String para verificar se o tipo de registro é do tipo MovimentoCaixa
        sTipo = Mid(sRegistro, iPosInicio, iPos - iPosInicio)
        
        'Verifica a Posição do Shifth
        iPosShift = InStr(iInicial, sRegistro, Chr(vbKeyShift))
        
        'Se for então
        If sTipo = sTipoCaixa Then
        
            Do While iPosInicio < (iPosEnd - 1)
        
                iIndice = iIndice + 1
                
                'acerta os Ponteiros
                If iIndice = 1 Then
                
                    iPosInicio = iPos + 1
                    iPos = InStr(iPos + 1, sRegistro, Chr(vbKeyControl))
                    
                End If
                    
                'Recolhe os Dados do arquivo de movimentosCaixa e Coloca no obj
                
                Select Case iIndice
                    
                    Case 1: objMovimentosCaixa.iTipo = StrParaInt(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
                    Case 2: objMovimentosCaixa.dHora = StrParaDbl(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
                    Case 3: objMovimentosCaixa.dtDataMovimento = StrParaDate(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
                    Case 4: objMovimentosCaixa.iGerente = StrParaInt(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
                    Case 5: objMovimentosCaixa.iCodOperador = StrParaInt(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
                    Case 6: objMovimentosCaixa.lSequencial = StrParaLong(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
                    Case 7: objMovimentosCaixa.iFilialEmpresa = StrParaInt(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
                    Case 8: objMovimentosCaixa.lTransferencia = StrParaLong(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
                    Case 9: objMovimentosCaixa.dValor = StrParaDbl(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
                    Case 10: objMovimentosCaixa.lNumMovto = StrParaLong(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
                    Case 11: objMovimentosCaixa.iAdmMeioPagto = StrParaInt(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
                    Case 12: objMovimentosCaixa.iParcelamento = StrParaInt(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
                    Case 13: objMovimentosCaixa.iExcluiu = StrParaInt(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
                    Case 14: objMovimentosCaixa.iCaixa = StrParaInt(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
                    Case 15: objMovimentosCaixa.iTipoCartao = StrParaInt(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
                    Case 16: objMovimentosCaixa.lCupomFiscal = StrParaLong(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
                    Case 17: objMovimentosCaixa.lMovtoEstorno = StrParaLong(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
                    Case 18: objMovimentosCaixa.lMovtoTransf = StrParaLong(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
                    Case 19: objMovimentosCaixa.lSequencialConta = StrParaLong(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
                    Case 20: objMovimentosCaixa.sFavorecido = Mid(sRegistro, iPosInicio, iPos - iPosInicio)
                    Case 21: objMovimentosCaixa.sHistorico = Mid(sRegistro, iPosInicio, iPos - iPosInicio)
                    Case 22: Exit Do
            
                End Select
                
                'Atualiza as Posições
                iPosInicio = iPos + 1
                iPos = (InStr(iPosInicio, sRegistro, Chr(vbKeyControl)))
        
            Loop
            
            'Adciona na Coleção Global
            colMovimentosCaixa.Add objMovimentosCaixa
    
            
            'atualiza as posições
            iPosInicioShift = iPosShift + 1
            iPosShift = InStr(iPosShift + 1, sRegistro, Chr(vbKeyShift))
            
            'Verifica se Existe na String o caracterShift se exitir executa o LOOP
            If iPosShift <> 0 Then
            
                'Enquanto for menor que a penultima posição
                Do While iPosInicioShift < (iPosEnd - 1)
                
                    'Adciona em lConteudo
                    lConteudo = StrParaLong(Mid(sRegistro, iPosInicioShift, iPosShift - iPosInicioShift))
                    
                    'Adciona na Coleção
                    colImfCompl.Add lConteudo
                    
                    'Atualiza as posições
                    iPosInicioShift = iPosShift + 1
                    iPosShift = InStr(iPosShift + 1, sRegistro, Chr(vbKeyShift))
                    
                    'Verifica se Chegou na Penultima posição se Chegou então
                    If iPosInicioShift = (iPosEnd - 1) Then
                    
                        'Procura a marcação do flag end
                        iPosShift = InStr(iPosShift + 1, sRegistro, Chr(vbKeyEnd))
                        'adciona na variável
                        lConteudo = StrParaLong(Mid(sRegistro, iPosInicioShift, iPosShift - iPosInicioShift))
                         
                        'Adciona na Coleção
                        colImfCompl.Add lConteudo
                   
                   End If
                
                Loop
            
            End If
            
        End If
        
    Loop
    
    'Fecha o Arquivo
    Close #1
    
    Desmembra_MovimentosCaixa = SUCESSO
       
    Exit Function

Erro_Desmembra_MovimentosCaixa:
    
    Desmembra_MovimentosCaixa = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149385)

    End Select
    
    'Fecha o Arquivo em caso de Erro
    Close #1

    Exit Function

End Function

Function Sangria_AFRAC(ByVal dValor As Double, ByVal sMensagem As String, ByVal iTipoSangria As Integer, Optional ByVal objTela As Object) As Long

Dim lErro As Long, bRelAberto As Boolean
Dim bSangria As Boolean, sTexto As String, iIndice As Integer, iCopias As Integer
Dim objOperador As New ClassOperador, sNomeOperador As String, objRel As Object

On Error GoTo Erro_Sangria_AFRAC
    
    If iTipoSangria = MOVIMENTOCAIXA_SANGRIA_DINHEIRO Or iTipoSangria = MOVIMENTOCAIXA_SUPRIMENTO_DINHEIRO Then
    
        If iTipoSangria = MOVIMENTOCAIXA_SANGRIA_DINHEIRO Then bSangria = True
        If iTipoSangria = MOVIMENTOCAIXA_SUPRIMENTO_DINHEIRO Then bSangria = False
        
        If giCodModeloECF = 1 Or giCodModeloECF = 2 Or giCodModeloECF = 3 Or giCodModeloECF = 5 Or giCodModeloECF = IMPRESSORA_DARUMA_FS600 Or giCodModeloECF = IMPRESSORA_BEMATECH_MP4200_TH_FI Or giCodModeloECF = IMPRESSORA_BEMATECH_MP2100_TH_FI Then
            If dValor > 0 Then
                lErro = AFRAC_SangriaSuprimento(bSangria, dValor)
                lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Sangria - Suprimento")
                If lErro <> SUCESSO Then gError 109695
            End If
        ElseIf giCodModeloECF = 4 Or AFRAC_ImpressoraCFe(giCodModeloECF) Then
        
            lErro = AFRAC_AbrirRelatorioGerencial2(RELGER_SANGRIASUPRIMENTO_DINHEIRO, objTela, objRel)
            bRelAberto = True
            
            For iCopias = 1 To 2
            
                sTexto = Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "Data: " & Format(Date, "dd/mm/yy") & " Hora: " & Format(Time, "hh:mm:ss") & " Caixa: " & giCodCaixa)
                lErro = AFRAC_ImprimirRelatorioGerencial2(sTexto, objTela, objRel)
                
                'Seleiona o nome do operador
                For Each objOperador In gcolOperadores
                    If objOperador.iCodigo = giCodOperador Then sNomeOperador = objOperador.sNome
                Next
        
                sTexto = Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "Operador: " & sNomeOperador)
                lErro = AFRAC_ImprimirRelatorioGerencial2(sTexto, objTela, objRel)
        
                sTexto = Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "")
                lErro = AFRAC_ImprimirRelatorioGerencial2(sTexto, objTela, objRel)
        
                If iTipoSangria = MOVIMENTOCAIXA_SANGRIA_DINHEIRO Then
                    sTexto = Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "Sangria Dinheiro - Valor: " & Format(dValor, "standard"))
                ElseIf iTipoSangria = MOVIMENTOCAIXA_SUPRIMENTO_DINHEIRO Then
                    sTexto = Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "Suprimento Dinheiro - Valor: " & Format(dValor, "standard"))
                ElseIf iTipoSangria = MOVIMENTOCAIXA_EXCLUSAO_SANGRIA_DINHEIRO Then
                    sTexto = Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "Exclusão Sangria Dinheiro - Valor: " & Format(dValor, "standard"))
                ElseIf iTipoSangria = MOVIMENTOCAIXA_EXCLUSAO_SUPRIMENTO_DINHEIRO Then
                    sTexto = Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "Exclusão Suprimento Dinheiro - Valor: " & Format(dValor, "standard"))
                End If
                lErro = AFRAC_ImprimirRelatorioGerencial2(sTexto, objTela, objRel)
                
                sTexto = Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "")
                lErro = AFRAC_ImprimirRelatorioGerencial2(sTexto, objTela, objRel)
                lErro = AFRAC_ImprimirRelatorioGerencial2(sTexto, objTela, objRel)
                
                sTexto = Formata_Campo(ALINHAMENTO_DIREITA, 48, "_", "")
                lErro = AFRAC_ImprimirRelatorioGerencial2(sTexto, objTela, objRel)
                
                sTexto = Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "")
                For iIndice = 1 To 10
                    lErro = AFRAC_ImprimirRelatorioGerencial2(sTexto, objTela, objRel)
                Next
                
            Next
        
            lErro = AFRAC_FecharRelatorioGerencial2(objTela, objRel)
            bRelAberto = False
            
            lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Sangria - Boleto")
            If lErro <> SUCESSO Then gError 133093
            
        End If
    
    ElseIf iTipoSangria = MOVIMENTOCAIXA_SANGRIA_CHEQUE Then
    
        lErro = AFRAC_AbrirRelatorioGerencial(RELGER_SANGRIA_CHEQUE, objTela)

        lErro = AFRAC_ImprimirRelatorioGerencial("Sangria Cheque - Valor: " & Format(dValor, "standard"), objTela)
        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Sangria - Cheque")
        If lErro <> SUCESSO Then gError 214000

        lErro = AFRAC_FecharRelatorioGerencial(objTela)
        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Sangria - Cheque")
        If lErro <> SUCESSO Then gError 214001
    
    
    ElseIf iTipoSangria = MOVIMENTOCAIXA_SANGRIA_BOLETO_CC Then
    
        lErro = AFRAC_AbrirRelatorioGerencial(RELGER_SANGRIA_CARTAO_CREDITO, objTela)

        lErro = AFRAC_ImprimirRelatorioGerencial("Sangria Boleto - Valor: " & Format(dValor, "standard"), objTela)
        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Sangria - Boleto")
        If lErro <> SUCESSO Then gError 214002

        lErro = AFRAC_FecharRelatorioGerencial(objTela)
        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Sangria - Boleto")
        If lErro <> SUCESSO Then gError 214003
    
    
    ElseIf iTipoSangria = MOVIMENTOCAIXA_SANGRIA_TICKET Then
    
        lErro = AFRAC_AbrirRelatorioGerencial(RELGER_SANGRIA_TICKET, objTela)

        lErro = AFRAC_ImprimirRelatorioGerencial("Sangria Ticket - Valor: " & Format(dValor, "standard"), objTela)
        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Sangria - Ticket")
        If lErro <> SUCESSO Then gError 214004

        lErro = AFRAC_FecharRelatorioGerencial(objTela)
        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Sangria - Ticket")
        If lErro <> SUCESSO Then gError 214005
    
    
    ElseIf iTipoSangria = MOVIMENTOCAIXA_SANGRIA_OUTROS Then
    
        lErro = AFRAC_AbrirRelatorioGerencial(RELGER_SANGRIA_OUTROS, objTela)

        lErro = AFRAC_ImprimirRelatorioGerencial("Sangria Outros - Valor: " & Format(dValor, "standard"), objTela)
        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Sangria - Outros")
        If lErro <> SUCESSO Then gError 214006

        lErro = AFRAC_FecharRelatorioGerencial(objTela)
        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Sangria - Outros")
        If lErro <> SUCESSO Then gError 214007
    
    
'        lErro = AFRAC_AbrirNaoFiscalNaoVinculado
'        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Abrir Não Fiscal Não Vinculado")
'        If lErro <> SUCESSO Then gError 109696
'
'        If bSangria Then
'            lErro = AFRAC_RegistrarNaoFiscal(TOTALIZADOR_SANGRIA, dValor, sMensagem)
'            lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Registrar Não Fiscal")
'            If lErro <> SUCESSO Then gError 109697
'        Else
'            lErro = AFRAC_RegistrarNaoFiscal(TOTALIZADOR_SUPRIMENTO, dValor, sMensagem)
'            lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Registrar Não fiscal")
'            If lErro <> SUCESSO Then gError 109804
'
'        End If
        
'        Call CF_ECF("Grava_R06", "CN")

'        lErro = AFRAC_FecharVinculado
'        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Fechar Vinculado")
'        If lErro <> SUCESSO Then gError 109698
    
        
    End If
    
    Sangria_AFRAC = SUCESSO
       
    Exit Function

Erro_Sangria_AFRAC:
    
    Sangria_AFRAC = gErr
    
    Select Case gErr
        
        Case 109695 To 109698, 109804, 133089 To 133093, 214000 To 214007
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149386)

    End Select
    
    If bRelAberto Then Call AFRAC_FecharRelatorioGerencial2(objTela, objRel)

    Exit Function
    
End Function
    
Function Desmembra_MovimentosCheque(ByVal colMovimentosCaixa As Collection, ByVal colImfCompl As Collection, sTipoCheque As String) As Long
'Função que Desmembra o Movimentos de Cheque e Carrega em gcolMovimentosCheque

Dim lErro As Long
Dim objMovimentosCaixa As ClassMovimentoCaixa
Dim iPosInicio As Integer
Dim sRegistro As String
Dim sTipo As String
Dim iPos As Integer
Dim iPosFinal As Integer
Dim iPosicao3  As Integer
Dim iPosShift As Integer
Dim iIndice As Integer
Dim sNomeArq As String
Dim iPosEnd As Integer
Dim iCont As Integer
Dim iPosInicioShift As Integer
Dim lConteudo As Long
Dim iInicial As Integer

On Error GoTo Erro_Desmembra_MovimentosCheque

    'Desmembra o Arquivo do dia Corrente
    sNomeArq = gsDirMVTEF & "MV" & CStr(Format(Date, "ddmmyy")) & (".txt")
    
    Open sNomeArq For Input As #1
    
    Do While Not EOF(1)
         
        iInicial = 1
             
        'Primeira Posição
        iPosInicio = 1
        
        'Inicializa a variavel
        iIndice = 0
        
        'Instancia o Obj da ClassMovimentoCheque
        Set objMovimentosCaixa = New ClassMovimentoCaixa
    
        'Busca o próximo registro do arquivo
        Line Input #1, sRegistro
        
        'Posição Final
        iPosEnd = InStr(iPosInicio, sRegistro, Chr(vbKeyEnd))
        
        'Procura o Primeiro Control para saber o tipo do registro
        iPos = InStr(iPosInicio, sRegistro, Chr(vbKeyControl))
        
        'String para verificar se o tipo de registro é do tipo MovimentoCheque
        sTipo = Mid(sRegistro, iPosInicio, iPos - iPosInicio)
        
        'Verifica a Posição do Shifth
        iPosShift = InStr(iInicial, sRegistro, Chr(vbKeyShift))
        
        'Se for então
        If sTipo = sTipoCheque Then
        
            Do While iPosInicio < (iPosEnd - 1)
        
                iIndice = iIndice + 1
                
                'acerta os Ponteiros
                If iIndice = 1 Then
                
                    iPosInicio = iPos + 1
                    iPos = InStr(iPos + 1, sRegistro, Chr(vbKeyControl))
                    
                End If
                    
                'Recolhe os Dados do arquivo de movimentosCheque e Coloca no obj
                
                Select Case iIndice
                    
                    Case 1: objMovimentosCaixa.iTipo = StrParaInt(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
                    Case 2: objMovimentosCaixa.dHora = StrParaDbl(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
                    Case 3: objMovimentosCaixa.dtDataMovimento = StrParaDate(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
                    Case 4: objMovimentosCaixa.iGerente = StrParaInt(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
                    Case 5: objMovimentosCaixa.iCodOperador = StrParaInt(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
                    Case 6: objMovimentosCaixa.lSequencial = StrParaLong(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
                    Case 7: objMovimentosCaixa.iFilialEmpresa = StrParaInt(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
                    Case 8: objMovimentosCaixa.dValor = StrParaDbl(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
                    Case 9: objMovimentosCaixa.iAdmMeioPagto = StrParaInt(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
                    Case 10: objMovimentosCaixa.iParcelamento = StrParaInt(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
                    Case 11: objMovimentosCaixa.iCaixa = StrParaInt(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
                    Case 12: Exit Do
            
                End Select
                
                'Atualiza as Posições
                iPosInicio = iPos + 1
                iPos = (InStr(iPosInicio, sRegistro, Chr(vbKeyControl)))
        
            Loop
            
            'Adciona na Coleção Global
            colMovimentosCaixa.Add objMovimentosCaixa
    
            
            'atualiza as posições
            iPosInicioShift = iPosShift + 1
            iPosShift = InStr(iPosShift + 1, sRegistro, Chr(vbKeyShift))
            
            'Verifica se Existe na String o caracterShift se exitir executa o LOOP
            If iPosShift <> 0 Then
            
                'Enquanto for menor que a penultima posição
                Do While iPosInicioShift < (iPosEnd - 1)
                
                    'Adciona em lConteudo
                    lConteudo = StrParaLong(Mid(sRegistro, iPosInicioShift, iPosShift - iPosInicioShift))
                    
                    'Adciona na Coleção
                    colImfCompl.Add lConteudo
                    
                    'Atualiza as posições
                    iPosInicioShift = iPosShift + 1
                    iPosShift = InStr(iPosShift + 1, sRegistro, Chr(vbKeyShift))
                    
                    'Verifica se Chegou na Penultima posição se Chegou então
                    If iPosInicioShift = (iPosEnd - 1) Then
                    
                        'Procura a marcação do flag end
                        iPosShift = InStr(iPosShift + 1, sRegistro, Chr(vbKeyEnd))
                        'adciona na variável
                        lConteudo = StrParaLong(Mid(sRegistro, iPosInicioShift, iPosShift - iPosInicioShift))
                         
                        'Adciona na Coleção
                        colImfCompl.Add lConteudo
                   
                   End If
                
                Loop
            
            End If
            
        End If
        
    Loop
    
    'Fecha o Arquivo
    Close #1
    
    Desmembra_MovimentosCheque = SUCESSO
       
    Exit Function

Erro_Desmembra_MovimentosCheque:
    
    Desmembra_MovimentosCheque = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149387)

    End Select
    
    'Fecha o Arquivo em caso de Erro
    Close #1

    Exit Function

End Function

Function Desmembra_AberturaCaixa(sNumSerie As String, dGT As Double, iCOO As Integer, iNumReducoesz As Integer) As Long
'Função que Desmembra a abertura do caixa

Dim lErro As Long
Dim iPosInicio As Integer
Dim sRegistro As String
Dim iPos As Integer
Dim iPosFinal As Integer
Dim iIndice As Integer
Dim sNomeArq As String
Dim sTipo As String
Dim iPosEnd As Integer

On Error GoTo Erro_Desmembra_AberturaCaixa

    'Desmembra o Arquivo do dia Corrente
    sNomeArq = gsDirMVTEF & "MV" & CStr(Format(Date, "ddmmyy")) & (".txt")
    
    Open sNomeArq For Input As #1
    
    Do While Not EOF(1)
         
        'Primeira Posição
        iPosInicio = 1
        
        'Inicializa a variavel
        iIndice = 0
        
        'Busca o próximo registro do arquivo
        Line Input #1, sRegistro
        
        'Posição Final
        iPosEnd = InStr(iPosInicio, sRegistro, Chr(vbKeyEnd))
        
        'Procura o Primeiro Control para saber o tipo do registro
        iPos = InStr(iPosInicio, sRegistro, Chr(vbKeyControl))
        
        'String para verificar se o tipo de registro é do tipo MovimentoCaixa
        sTipo = Mid(sRegistro, iPosInicio, iPos - iPosInicio)
        
        'Se for então
        If sTipo = TIPOREGISTROECF_ABERTURACAIXA Then
        
            Do While iPosInicio < (iPosEnd - 1)
        
                iIndice = iIndice + 1
                
                'acerta os Ponteiros
                If iIndice = 1 Then
                
                    iPosInicio = iPos + 1
                    iPos = InStr(iPos + 1, sRegistro, Chr(vbKeyControl))
                    
                End If
                    
                Select Case iIndice
                    
                    Case 1: sNumSerie = Mid(sRegistro, iPosInicio, iPos - iPosInicio)
                    Case 2: dGT = StrParaDbl(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
                    Case 3: iCOO = StrParaInt(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
                    Case 4: iNumReducoesz = StrParaInt(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
                    Case 5: Exit Do
            
                End Select
                
                'Atualiza as Posições
                iPosInicio = iPos + 1
                iPos = (InStr(iPosInicio, sRegistro, Chr(vbKeyControl)))
        
            Loop
        End If
    Loop
    
    'Fecha o Arquivo
    Close #1
    
    Desmembra_AberturaCaixa = SUCESSO
       
    Exit Function

Erro_Desmembra_AberturaCaixa:
    
    Desmembra_AberturaCaixa = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149388)

    End Select
    
    'Fecha o Arquivo em caso de Erro
    Close #1

    Exit Function

End Function

Function Desmembra_FechamentoCaixa(sNumSerie As String, dGT As Double, iCOO As Integer) As Long
'Função que Desmembra a Fechamento do caixa

Dim lErro As Long
Dim iPosInicio As Integer
Dim sRegistro As String
Dim iPos As Integer
Dim iPosFinal As Integer
Dim iIndice As Integer
Dim sNomeArq As String
Dim sTipo As String
Dim iPosEnd As Integer

On Error GoTo Erro_Desmembra_FechamentoCaixa

    'Desmembra o Arquivo do dia Corrente
    sNomeArq = gsDirMVTEF & "MV" & CStr(Format(Date, "ddmmyy")) & (".txt")
    
    Open sNomeArq For Input As #1
    
    Do While Not EOF(1)
         
        'Primeira Posição
        iPosInicio = 1
        
        'Inicializa a variavel
        iIndice = 0
        
        'Busca o próximo registro do arquivo
        Line Input #1, sRegistro
        
        'Posição Final
        iPosEnd = InStr(iPosInicio, sRegistro, Chr(vbKeyEnd))
        
        'Procura o Primeiro Control para saber o tipo do registro
        iPos = InStr(iPosInicio, sRegistro, Chr(vbKeyControl))
        
        'String para verificar se o tipo de registro é do tipo MovimentoCaixa
        sTipo = Mid(sRegistro, iPosInicio, iPos - iPosInicio)
        
        'Se for então
        If sTipo = TIPOREGISTROECF_FECHAMENTOCAIXA Then
        
            Do While iPosInicio < (iPosEnd - 1)
        
                iIndice = iIndice + 1
                
                'acerta os Ponteiros
                If iIndice = 1 Then
                
                    iPosInicio = iPos + 1
                    iPos = InStr(iPos + 1, sRegistro, Chr(vbKeyControl))
                    
                End If
                    
                Select Case iIndice
                    
                    Case 1: sNumSerie = Mid(sRegistro, iPosInicio, iPos - iPosInicio)
                    Case 2: dGT = StrParaDbl(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
                    Case 3: iCOO = StrParaInt(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
                    Case 4: Exit Do
            
                End Select
                
                'Atualiza as Posições
                iPosInicio = iPos + 1
                iPos = (InStr(iPosInicio, sRegistro, Chr(vbKeyControl)))
        
            Loop
        End If
    Loop
    
    'Fecha o Arquivo
    Close #1
    
    Desmembra_FechamentoCaixa = SUCESSO
       
    Exit Function

Erro_Desmembra_FechamentoCaixa:
    
    Desmembra_FechamentoCaixa = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149389)

    End Select
    
    'Fecha o Arquivo em caso de Erro
    Close #1

    Exit Function

End Function

Function Desmembra_Aliquotas(sNumSerie As String, sAliquota As String, dValor As Double) As Long
'Função que Desmembra a Fechamento do caixa

Dim lErro As Long
Dim iPosInicio As Integer
Dim sRegistro As String
Dim iPos As Integer
Dim iPosFinal As Integer
Dim iIndice As Integer
Dim sNomeArq As String
Dim sTipo As String
Dim iPosEnd As Integer

On Error GoTo Erro_Desmembra_Aliquotas

    'Desmembra o Arquivo do dia Corrente
    sNomeArq = gsDirMVTEF & "MV" & CStr(Format(Date, "ddmmyy")) & (".txt")
    
    Open sNomeArq For Input As #1
    
    Do While Not EOF(1)
         
        'Primeira Posição
        iPosInicio = 1
        
        'Inicializa a variavel
        iIndice = 0
        
        'Busca o próximo registro do arquivo
        Line Input #1, sRegistro
        
        'Posição Final
        iPosEnd = InStr(iPosInicio, sRegistro, Chr(vbKeyEnd))
        
        'Procura o Primeiro Control para saber o tipo do registro
        iPos = InStr(iPosInicio, sRegistro, Chr(vbKeyControl))
        
        'String para verificar se o tipo de registro é do tipo MovimentoCaixa
        sTipo = Mid(sRegistro, iPosInicio, iPos - iPosInicio)
        
        'Se for então
        If sTipo = TIPOREGISTROECF_ALIQUOTAS Then
        
            Do While iPosInicio < (iPosEnd - 1)
        
                iIndice = iIndice + 1
                
                'acerta os Ponteiros
                If iIndice = 1 Then
                
                    iPosInicio = iPos + 1
                    iPos = InStr(iPos + 1, sRegistro, Chr(vbKeyControl))
                    
                End If
                    
                Select Case iIndice
                    
                    Case 1: sNumSerie = Mid(sRegistro, iPosInicio, iPos - iPosInicio)
                    Case 2: sAliquota = Mid(sRegistro, iPosInicio, iPos - iPosInicio)
                    Case 3: dValor = StrParaDbl(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
                    Case 4: Exit Do
            
                End Select
                
                'Atualiza as Posições
                iPosInicio = iPos + 1
                iPos = (InStr(iPosInicio, sRegistro, Chr(vbKeyControl)))
        
            Loop
        End If
    Loop
    
    'Fecha o Arquivo
    Close #1
    
    Desmembra_Aliquotas = SUCESSO
       
    Exit Function

Erro_Desmembra_Aliquotas:
    
    Desmembra_Aliquotas = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149390)

    End Select
    
    'Fecha o Arquivo em caso de Erro
    Close #1

    Exit Function

End Function

Public Function DeTipoMovtoParaTMP(ByVal iMovto As Integer) As Integer
'essa funcao recebe o tipo do movimento e retorna o tipo de meio de pagamento a ele associado...
'iMovto eh parametro de input

   Select Case iMovto

      Case MOVIMENTOCAIXA_ENTRADA_TRANSF_DINHEIRO, MOVIMENTOCAIXA_SAIDA_TRANSF_DINHEIRO, MOVIMENTOCAIXA_EXCLUSAO_ENTRADA_TRANSF_DINHEIRO, MOVIMENTOCAIXA_EXCLUSAO_SAIDA_TRANSF_DINHEIRO
         DeTipoMovtoParaTMP = TIPOMEIOPAGTOLOJA_DINHEIRO

      Case MOVIMENTOCAIXA_ENTRADA_TRANSF_CHEQUE, MOVIMENTOCAIXA_SAIDA_TRANSF_CHEQUE, MOVIMENTOCAIXA_EXCLUSAO_ENTRADA_TRANSF_CHEQUE, MOVIMENTOCAIXA_EXCLUSAO_SAIDA_TRANSF_CHEQUE
         DeTipoMovtoParaTMP = TIPOMEIOPAGTOLOJA_CHEQUE

      Case MOVIMENTOCAIXA_ENTRADA_TRANSF_VALETICKET, MOVIMENTOCAIXA_SAIDA_TRANSF_VALETICKET, MOVIMENTOCAIXA_EXCLUSAO_ENTRADA_TRANSF_VALETICKET, MOVIMENTOCAIXA_EXCLUSAO_SAIDA_TRANSF_VALETICKET
         DeTipoMovtoParaTMP = TIPOMEIOPAGTOLOJA_VALE_TICKET

      Case MOVIMENTOCAIXA_ENTRADA_TRANSF_CARTAO_CREDITO, MOVIMENTOCAIXA_SAIDA_TRANSF_CARTAO_CREDITO, MOVIMENTOCAIXA_EXCLUSAO_ENTRADA_TRANSF_CARTAO_CREDITO, MOVIMENTOCAIXA_EXCLUSAO_SAIDA_TRANSF_CARTAO_CREDITO
         DeTipoMovtoParaTMP = TIPOMEIOPAGTOLOJA_CARTAO_CREDITO

      Case MOVIMENTOCAIXA_ENTRADA_TRANSF_CARTAO_DEBITO, MOVIMENTOCAIXA_SAIDA_TRANSF_CARTAO_DEBITO, MOVIMENTOCAIXA_EXCLUSAO_ENTRADA_TRANSF_CARTAO_DEBITO, MOVIMENTOCAIXA_EXCLUSAO_SAIDA_TRANSF_CARTAO_DEBITO
         DeTipoMovtoParaTMP = TIPOMEIOPAGTOLOJA_CARTAO_DEBITO

      Case MOVIMENTOCAIXA_ENTRADA_TRANSF_OUTROS, MOVIMENTOCAIXA_SAIDA_TRANSF_OUTROS, MOVIMENTOCAIXA_EXCLUSAO_SAIDA_TRANSF_OUTROS, MOVIMENTOCAIXA_EXCLUSAO_ENTRADA_TRANSF_OUTROS
         DeTipoMovtoParaTMP = TIPOMEIOPAGTOLOJA_OUTROS

      Case Else
         DeTipoMovtoParaTMP = -1

   End Select

   Exit Function

End Function

Public Function Obtem_Nome_AdmMeioPagto(ByVal objMovimentoCaixa As ClassMovimentoCaixa, sAdmMeioPagto As String) As Long

Dim objAdmMeioPagto As ClassAdmMeioPagto

    If objMovimentoCaixa.iAdmMeioPagto = 0 Then
        
        Select Case CF_ECF("DeTipoMovtoParaTMP", objMovimentoCaixa.iTipo)
        
            Case TIPOMEIOPAGTOLOJA_CARTAO_CREDITO
                sAdmMeioPagto = STRING_NAO_DETALHADO_CCREDITO
                
            Case TIPOMEIOPAGTOLOJA_CARTAO_DEBITO
                sAdmMeioPagto = STRING_NAO_DETALHADO_CDEBITO
            
            Case Else
                sAdmMeioPagto = STRING_NAO_DETALHADO
        
        End Select
        
    Else
    
        For Each objAdmMeioPagto In gcolAdmMeioPagto
        
            If objAdmMeioPagto.iCodigo = objMovimentoCaixa.iAdmMeioPagto Then
                
                sAdmMeioPagto = objAdmMeioPagto.sNome
                Exit For
                
            End If
        
        Next
    
    End If
    
    Obtem_Nome_AdmMeioPagto = SUCESSO

End Function

Public Function TransfCaixa_Le(objTransfCaixa As ClassTransfCaixa) As Long
'busca transferencia retornando no objTranfCaixa (q eh parametro de input
'pois traz o numero da transferencia (propriedade lTransferencia) e tbm eh
'parametro de output (pois retorna a transferencia caso a mesma tenha sido encontrada
'na colecao de movimentos...
Dim objMovCaixaAux As ClassMovimentoCaixa
Dim bAchou As Boolean

On Error GoTo Erro_TransfCaixa_Le

    bAchou = False
    
    For Each objMovCaixaAux In gcolMovimentosCaixa
    
        'se achou o movimento q participa da transferencia
        If objMovCaixaAux.lTransferencia = objTransfCaixa.objMovCaixaDe.lTransferencia Then
        
            'se o movimento é de saída
            If Movimento_Eh_De_Saida(objMovCaixaAux) = True Then
                
                With objTransfCaixa.objMovCaixaDe
                    
                    .dHora = objMovCaixaAux.dHora
                    .dtDataMovimento = objMovCaixaAux.dtDataMovimento
                    .dValor = objMovCaixaAux.dValor
                    .iAdmMeioPagto = objMovCaixaAux.iAdmMeioPagto
                    .iCaixa = objMovCaixaAux.iCaixa
                    .iCodConta = objMovCaixaAux.iCodConta
                    .iCodOperador = objMovCaixaAux.iCodOperador
                    .iExcluiu = objMovCaixaAux.iExcluiu
                    .iFilialEmpresa = objMovCaixaAux.iFilialEmpresa
                    .iGerente = objMovCaixaAux.iGerente
                    .iParcelamento = objMovCaixaAux.iParcelamento
                    .iTipo = objMovCaixaAux.iTipo
                    .lCupomFiscal = objMovCaixaAux.lCupomFiscal
                    .lMovtoEstorno = objMovCaixaAux.lMovtoEstorno
                    .lMovtoTransf = objMovCaixaAux.lMovtoTransf
                    .lNumero = objMovCaixaAux.lNumero
                    .lNumIntExt = objMovCaixaAux.lNumIntExt
                    .lNumMovto = objMovCaixaAux.lNumMovto
                    .lNumRefInterna = objMovCaixaAux.lNumRefInterna
                    .lSequencial = objMovCaixaAux.lSequencial
                    .lSequencialConta = objMovCaixaAux.lSequencialConta
                    .lTransferencia = objMovCaixaAux.lTransferencia
                    .sFavorecido = objMovCaixaAux.sFavorecido
                    .sHistorico = objMovCaixaAux.sHistorico
                    .iTipoCartao = objMovCaixaAux.iTipoCartao
                    .lNumIntDocLog = objMovCaixaAux.lNumIntDocLog
                
                End With
                
            'caso contrário
            Else
                
                With objTransfCaixa.objMovCaixaPara

                    .dHora = objMovCaixaAux.dHora
                    .dtDataMovimento = objMovCaixaAux.dtDataMovimento
                    .dValor = objMovCaixaAux.dValor
                    .iAdmMeioPagto = objMovCaixaAux.iAdmMeioPagto
                    .iCaixa = objMovCaixaAux.iCaixa
                    .iCodConta = objMovCaixaAux.iCodConta
                    .iCodOperador = objMovCaixaAux.iCodOperador
                    .iExcluiu = objMovCaixaAux.iExcluiu
                    .iFilialEmpresa = objMovCaixaAux.iFilialEmpresa
                    .iGerente = objMovCaixaAux.iGerente
                    .iParcelamento = objMovCaixaAux.iParcelamento
                    .iTipo = objMovCaixaAux.iTipo
                    .lCupomFiscal = objMovCaixaAux.lCupomFiscal
                    .lMovtoEstorno = objMovCaixaAux.lMovtoEstorno
                    .lMovtoTransf = objMovCaixaAux.lMovtoTransf
                    .lNumero = objMovCaixaAux.lNumero
                    .lNumIntExt = objMovCaixaAux.lNumIntExt
                    .lNumMovto = objMovCaixaAux.lNumMovto
                    .lNumRefInterna = objMovCaixaAux.lNumRefInterna
                    .lSequencial = objMovCaixaAux.lSequencial
                    .lSequencialConta = objMovCaixaAux.lSequencialConta
                    .lTransferencia = objMovCaixaAux.lTransferencia
                    .sFavorecido = objMovCaixaAux.sFavorecido
                    .sHistorico = objMovCaixaAux.sHistorico
                    .iTipoCartao = objMovCaixaAux.iTipoCartao
                    .lNumIntDocLog = objMovCaixaAux.lNumIntDocLog
                    
                End With
            
            End If
        
            bAchou = True
        
        End If
        
    Next
    
    If bAchou = False Then gError 109404
    
    TransfCaixa_Le = SUCESSO
    
    Exit Function
    
Erro_TransfCaixa_Le:
    
    TransfCaixa_Le = gErr
    
    Select Case gErr
    
        Case 109404
            'será tratado na rotina chamadora
    
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 149391)
    
    End Select
    
    Exit Function
    
End Function

Public Function Cheque_Le_SequencialCaixa(objCheque As ClassChequePre) As Long

Dim bAchou As Boolean
Dim lErro As Long
Dim objChequeAux As ClassChequePre

On Error GoTo Erro_Cheque_Le_SequencialCaixa

    bAchou = False
    
    For Each objChequeAux In gcolCheque
    
        If objCheque.lSequencialCaixa = objChequeAux.lSequencialCaixa Then
        
'            objCheque.dtDataDeposito = objChequeAux.dtDataDeposito
'            objCheque.dValor = objChequeAux.dValor
'            objCheque.iAprovado = objChequeAux.iAprovado
'            objCheque.iBanco = objChequeAux.iBanco
'            objCheque.iCaixa = objChequeAux.iCaixa
'            objCheque.iChequeSel = objChequeAux.iChequeSel
'            objCheque.iECF = objChequeAux.iECF
'            objCheque.iFilial = objChequeAux.iFilial
'            objCheque.iFilialEmpresa = objChequeAux.iFilialEmpresa
'            objCheque.iFilialEmpresaLoja = objChequeAux.iFilialEmpresaLoja
'            objCheque.iNaoEspecificado = objChequeAux.iNaoEspecificado
'            objCheque.iStatus = objChequeAux.iStatus
'            objCheque.iTipoBordero = objChequeAux.iTipoBordero
'            objCheque.lCliente = objChequeAux.lCliente
'            objCheque.lNumBordero = objChequeAux.lNumBordero
'            objCheque.lNumBorderoLoja = objChequeAux.lNumBorderoLoja
'            objCheque.lNumero = objChequeAux.lNumero
'            objCheque.lNumIntCheque = objChequeAux.lNumIntCheque
'            objCheque.lNumIntChequeBord = objChequeAux.lNumIntChequeBord
'            objCheque.lNumIntExt = objChequeAux.lNumIntExt
'            objCheque.lNumMovtoCaixa = objChequeAux.lNumMovtoCaixa
'            objCheque.lNumMovtoSangria = objChequeAux.lNumMovtoSangria
'            objCheque.lSequencial = objChequeAux.lSequencial
'            objCheque.lSequencialBack = objChequeAux.lSequencialBack
'            objCheque.lSequencialCaixa = objChequeAux.lSequencialCaixa
'            objCheque.lSequencialLoja = objChequeAux.lSequencialLoja
'            objCheque.sAgencia = objChequeAux.sAgencia
'            objCheque.sCarne = objChequeAux.sCarne
'            objCheque.sContaCorrente = objChequeAux.sContaCorrente
'            objCheque.sCPFCGC = objChequeAux.sCPFCGC
'            objCheque.lCupomFiscal = objChequeAux.lCupomFiscal
        
            Set objCheque = objChequeAux
        
            bAchou = True
            
            Exit For
        
        End If
    
    Next
    
    If bAchou = False Then gError 109462
    
    Cheque_Le_SequencialCaixa = SUCESSO
    
    Exit Function
    
Erro_Cheque_Le_SequencialCaixa:

    Cheque_Le_SequencialCaixa = gErr
    
    Select Case gErr
    
        Case 109462
        'será tratado na rotina chamadora
    
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 149392)
    
    End Select
    
    Exit Function

End Function

Public Function Movimento_Eh_De_Saida(objMovCaixaAux As ClassMovimentoCaixa) As Boolean
'verifica se o movimento eh de saida ou de entrada...
'objmovcaixaaux eh parametro de input

    If objMovCaixaAux.iTipo = MOVIMENTOCAIXA_SAIDA_TRANSF_CARTAO_CREDITO Or _
            objMovCaixaAux.iTipo = MOVIMENTOCAIXA_SAIDA_TRANSF_CARTAO_DEBITO Or _
            objMovCaixaAux.iTipo = MOVIMENTOCAIXA_SAIDA_TRANSF_CHEQUE Or _
            objMovCaixaAux.iTipo = MOVIMENTOCAIXA_SAIDA_TRANSF_DINHEIRO Or _
            objMovCaixaAux.iTipo = MOVIMENTOCAIXA_SAIDA_TRANSF_OUTROS Or _
            objMovCaixaAux.iTipo = MOVIMENTOCAIXA_SAIDA_TRANSF_VALETICKET Or _
            objMovCaixaAux.iTipo = MOVIMENTOCAIXA_EXCLUSAO_ENTRADA_TRANSF_CARTAO_CREDITO Or _
            objMovCaixaAux.iTipo = MOVIMENTOCAIXA_EXCLUSAO_ENTRADA_TRANSF_CARTAO_DEBITO Or _
            objMovCaixaAux.iTipo = MOVIMENTOCAIXA_EXCLUSAO_ENTRADA_TRANSF_CHEQUE Or _
            objMovCaixaAux.iTipo = MOVIMENTOCAIXA_EXCLUSAO_ENTRADA_TRANSF_DINHEIRO Or _
            objMovCaixaAux.iTipo = MOVIMENTOCAIXA_EXCLUSAO_ENTRADA_TRANSF_OUTROS Or _
            objMovCaixaAux.iTipo = MOVIMENTOCAIXA_EXCLUSAO_ENTRADA_TRANSF_VALETICKET Then

            Movimento_Eh_De_Saida = True

   Else
            Movimento_Eh_De_Saida = False

   End If

End Function

Function Caixa_Obtem_NumAutomatico(lNumero As Long) As Long
'Função que Serve para Gerar um novo Numero de código

Dim lErro As Long
Dim sProxNumero As String
Dim dtData As Date

On Error GoTo Erro_Caixa_Obtem_NumAutomatico

    lErro = CF_ECF("MovimentoCaixa_Le", lNumero, dtData)
    If lErro <> SUCESSO Then gError 210435

'    lErro = 105352
'
'    Do While lErro <> SUCESSO
'
'        'Função que Gera o Codigo Atraves do Código do Ultimo Movto
'        lErro = Caixa_Obtem_NumAutomatico_Aux(lNumero)
'        If lErro <> SUCESSO And lErro <> 105352 Then gError 105388
'
'    Loop

    glNumProxMovto = lNumero + 1

'    'Transforma o long em String
'    sProxNumero = CStr(glNumProxMovto)
'
'    'Grava o novo numero no Arquivo de Configuração de Caixa
'    lErro = WritePrivateProfileString(CAIXA_CAIXA_CONFIG, "NumProxMovto", sProxNumero, NOME_ARQUIVO_CAIXA)
'    If lErro = 0 Then gError 117532

    'funcao que cria/altera a chave e grava no caixaconfig.ini
    lErro = CF_ECF("CaixaConfig_Chave")
    If lErro <> SUCESSO Then gError 204294

    Caixa_Obtem_NumAutomatico = SUCESSO

    Exit Function

Erro_Caixa_Obtem_NumAutomatico:

    Caixa_Obtem_NumAutomatico = gErr

    Select Case gErr

        Case 105388, 204294, 210435

        Case 117532
            Call Rotina_ErroECF(vbOKOnly, ERRO_ARQUIVO_NAO_ENCONTRADO1, gErr, CAIXA_CAIXA_CONFIG, "NumProxMovto", NOME_ARQUIVO_CAIXA)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149393)

    End Select

    Exit Function

End Function

Private Function Caixa_Obtem_NumAutomatico_Aux(lNumero As Long) As Long
'Função que Atualiza o Numero do Movimento, para se gravado no Arquivo e configuração

Dim lErro As Long
Dim objMovimentoCaixa As New ClassMovimentoCaixa

On Error GoTo Erro_Caixa_Obtem_NumAutomatico_Aux

    'Verificar na Coleção Global
    For Each objMovimentoCaixa In gcolMovimentosCaixa

        'Verifica e o Numero é Igual ao Numero do Movimento na Coleção global de movimentos se for chama a função
        If lNumero = objMovimentoCaixa.lNumMovto Then

            'Incremento o Numero do Movimento
            lNumero = lNumero + 1

            gError 105352

        End If

    Next

    Caixa_Obtem_NumAutomatico_Aux = SUCESSO

    Exit Function

Erro_Caixa_Obtem_NumAutomatico_Aux:

    Caixa_Obtem_NumAutomatico_Aux = gErr

    Select Case gErr

        Case 105352

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149394)

    End Select

    Exit Function

End Function

Function Move_DadosGlobais_Memoria(objMovimentoCaixa As ClassMovimentoCaixa) As Long
'Função que Move os Dados para a Memória

Dim lErro As Long

On Error GoTo Erro_Move_DadosGlobais_Memoria

    'Guarda as Informações Globais no objMovimentosCaixa

    objMovimentoCaixa.iFilialEmpresa = giFilialEmpresa

    objMovimentoCaixa.iCaixa = giCodCaixa

    objMovimentoCaixa.iCodOperador = giCodOperador

    objMovimentoCaixa.dtDataMovimento = Date

    objMovimentoCaixa.dHora = CDbl(Time)

    Move_DadosGlobais_Memoria = SUCESSO

    Exit Function

Erro_Move_DadosGlobais_Memoria:

    Move_DadosGlobais_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149395)

    End Select

    Exit Function

End Function

Function MovimentoCaixa_Copia(objMovimentoCaixa As ClassMovimentoCaixa, objMovCx As ClassMovimentoCaixa) As Long
'Copia os dados de objMovimentoCaixa para objMovCx

Dim lErro As Long

On Error GoTo Erro_MovimentoCaixa_Copia

    objMovCx.dHora = objMovimentoCaixa.dHora
    objMovCx.dtDataMovimento = objMovimentoCaixa.dtDataMovimento
    objMovCx.dValor = objMovimentoCaixa.dValor
    objMovCx.iAdmMeioPagto = objMovimentoCaixa.iAdmMeioPagto
    objMovCx.iCaixa = objMovimentoCaixa.iCaixa
    objMovCx.iCodConta = objMovimentoCaixa.iCodConta
    objMovCx.iCodOperador = objMovimentoCaixa.iCodOperador
    objMovCx.iExcluiu = objMovimentoCaixa.iExcluiu
    objMovCx.iFilialEmpresa = objMovimentoCaixa.iFilialEmpresa
    objMovCx.iGerente = objMovimentoCaixa.iGerente
    objMovCx.iParcelamento = objMovimentoCaixa.iParcelamento
    objMovCx.iQuantLog = objMovimentoCaixa.iQuantLog
    objMovCx.iTipo = objMovimentoCaixa.iTipo
    objMovCx.iTipoCartao = objMovimentoCaixa.iTipoCartao
    objMovCx.lCupomFiscal = objMovimentoCaixa.lCupomFiscal
    objMovCx.lMovtoEstorno = objMovimentoCaixa.lMovtoEstorno
    objMovCx.lMovtoTransf = objMovimentoCaixa.lMovtoTransf
    objMovCx.lNumero = objMovimentoCaixa.lNumero
    objMovCx.lNumIntDocLog = objMovimentoCaixa.lNumIntDocLog
    objMovCx.lNumIntExt = objMovimentoCaixa.lNumIntExt
    objMovCx.lNumMovto = objMovimentoCaixa.lNumMovto
    objMovCx.lNumRefInterna = objMovimentoCaixa.lNumRefInterna
    objMovCx.lSequencial = objMovimentoCaixa.lSequencial
    objMovCx.lSequencialConta = objMovimentoCaixa.lSequencialConta
    objMovCx.lTransferencia = objMovimentoCaixa.lTransferencia
    objMovCx.sFavorecido = objMovimentoCaixa.sFavorecido
    objMovCx.sHistorico = objMovimentoCaixa.sHistorico
    objMovCx.dtDataPreDatado = objMovimentoCaixa.dtDataPreDatado
    objMovCx.sAutorizacao = objMovimentoCaixa.sAutorizacao

    MovimentoCaixa_Copia = SUCESSO

    Exit Function

Erro_MovimentoCaixa_Copia:

    MovimentoCaixa_Copia = gErr

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149396)

    End Select

    Exit Function

End Function

Function MovimentoBoleto_Atualiza_Memoria1(objMovimentoCaixa As ClassMovimentoCaixa) As Long
'Função que Limpa a Coleção Global a Tela apos a Função de Gravação

Dim lErro As Long
Dim objAdmMeioPagtoCondPagto As New ClassAdmMeioPagtoCondPagto

On Error GoTo Erro_MovimentoBoleto_Atualiza_Memoria1

    For Each objAdmMeioPagtoCondPagto In gcolCartao

        'Verifica se é a Mensma administardora ,parcelamento , Cartão
        If objAdmMeioPagtoCondPagto.iAdmMeioPagto = objMovimentoCaixa.iAdmMeioPagto And objAdmMeioPagtoCondPagto.iParcelamento = objMovimentoCaixa.iParcelamento And objAdmMeioPagtoCondPagto.iTipoCartao = objMovimentoCaixa.iTipoCartao And objAdmMeioPagtoCondPagto.dtDataPreDatado = objMovimentoCaixa.dtDataPreDatado Then
            
            If objMovimentoCaixa.iAdmMeioPagto = 0 Then

                'Soma ao Saldo da administardora ,parcelamento , Cartão o Movimento de Caixa Correspondente
                If objAdmMeioPagtoCondPagto.sNomeAdmMeioPagto = STRING_NAO_DETALHADO_CCREDITO Then
                
                    If objMovimentoCaixa.iTipo = MOVIMENTOCAIXA_EXCLUSAO_SANGRIA_B0LETO Then
                        objAdmMeioPagtoCondPagto.dSaldo = objAdmMeioPagtoCondPagto.dSaldo + objMovimentoCaixa.dValor
                    Else
                        objAdmMeioPagtoCondPagto.dSaldo = objAdmMeioPagtoCondPagto.dSaldo - objMovimentoCaixa.dValor
                    End If
                    
                    Exit For
                    
                End If
                    
            Else
            
                If objMovimentoCaixa.iTipo = MOVIMENTOCAIXA_EXCLUSAO_SANGRIA_B0LETO Then
                    objAdmMeioPagtoCondPagto.dSaldo = objAdmMeioPagtoCondPagto.dSaldo + objMovimentoCaixa.dValor
                Else
                    objAdmMeioPagtoCondPagto.dSaldo = objAdmMeioPagtoCondPagto.dSaldo - objMovimentoCaixa.dValor
                End If
                
                Exit For
                    
            End If

        End If

    Next

    MovimentoBoleto_Atualiza_Memoria1 = SUCESSO

    Exit Function

Erro_MovimentoBoleto_Atualiza_Memoria1:

    MovimentoBoleto_Atualiza_Memoria1 = gErr

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 149397)

    End Select

    Exit Function

End Function

Function MovimentoTicket_Atualiza_Memoria1(objMovimentoCaixa As ClassMovimentoCaixa) As Long
'Função que Limpa a Coleção Global a Tela apos a Função de Gravação

Dim lErro As Long
Dim objAdmMeioPagtoCondPagto As New ClassAdmMeioPagtoCondPagto

On Error GoTo Erro_MovimentoTicket_Atualiza_Memoria1

    For Each objAdmMeioPagtoCondPagto In gcolTicket

        'Verifica se é a Mensma administardora
        If objAdmMeioPagtoCondPagto.iAdmMeioPagto = objMovimentoCaixa.iAdmMeioPagto And objAdmMeioPagtoCondPagto.iParcelamento = objMovimentoCaixa.iParcelamento Then

                If objMovimentoCaixa.iTipo = MOVIMENTOCAIXA_EXCLUSAO_SANGRIA_TICKET Then
    
                    'Soma ao Saldo da administardora  com o do Movto
                    objAdmMeioPagtoCondPagto.dSaldo = objAdmMeioPagtoCondPagto.dSaldo + objMovimentoCaixa.dValor
                    
                Else
                
                    'Soma ao Saldo da administardora  com o do Movto
                    objAdmMeioPagtoCondPagto.dSaldo = objAdmMeioPagtoCondPagto.dSaldo - objMovimentoCaixa.dValor
                            
                End If
                
                Exit For

        End If

    Next

    MovimentoTicket_Atualiza_Memoria1 = SUCESSO

    Exit Function

Erro_MovimentoTicket_Atualiza_Memoria1:

    MovimentoTicket_Atualiza_Memoria1 = gErr

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 149398)

    End Select

    Exit Function

End Function

Function MovimentoOutros_Atualiza_Memoria1(objMovimentoCaixa As ClassMovimentoCaixa) As Long
'Função que Limpa a Coleção Global a Tela apos a Função de Gravação

Dim lErro As Long
Dim objAdmMeioPagtoCondPagto As New ClassAdmMeioPagtoCondPagto

On Error GoTo Erro_MovimentoOutros_Atualiza_Memoria1

    For Each objAdmMeioPagtoCondPagto In gcolOutros

        'Verifica se é o mesmo meio de Pagto
        If objAdmMeioPagtoCondPagto.iAdmMeioPagto = objMovimentoCaixa.iAdmMeioPagto And objAdmMeioPagtoCondPagto.iParcelamento = objMovimentoCaixa.iParcelamento Then

            If objMovimentoCaixa.iTipo = MOVIMENTOCAIXA_EXCLUSAO_SANGRIA_OUTROS Then

                'Soma ao Saldo do meio Pagto com o do Movto
                objAdmMeioPagtoCondPagto.dSaldo = objAdmMeioPagtoCondPagto.dSaldo + objMovimentoCaixa.dValor
                
            Else
            
                'Subtrai ao Saldo do meio Pagto com o do Movto
                objAdmMeioPagtoCondPagto.dSaldo = objAdmMeioPagtoCondPagto.dSaldo - objMovimentoCaixa.dValor
            
            End If
            
            Exit For

        End If

    Next

    MovimentoOutros_Atualiza_Memoria1 = SUCESSO

    Exit Function

Erro_MovimentoOutros_Atualiza_Memoria1:

    MovimentoOutros_Atualiza_Memoria1 = gErr

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 149399)

    End Select

    Exit Function

End Function

Function Registro_ECF_CC(ByVal colRegistro As Collection, ByVal sArquivo As String, ByVal iRegistro As Integer, ByVal iIndiceReg As Integer, ByVal colVendas As Collection, ByVal iCodModeloECF As Integer) As Long
'Descobre o tipo de registro e chama o tratador adequado. sArquivo contem o nome do arquivo em que sRegistro foi lido.
'iRegistro é o numero do registro.

Dim objVenda As New ClassVenda
Dim objMovimentosCaixa As New ClassMovimentoCaixa
Dim objTransfCaixa As New ClassTransfCaixa
Dim colImfCompl As New Collection
Dim objChequePara As New ClassChequePre
Dim lErro As Long
Dim iPosKeyControl As Integer
Dim iTipoRegistro As Integer
Dim sRegistro As String
Dim sRegistro1 As String
Dim iIndice As Integer
Dim iPosInicio1 As Integer
Dim sSeq1 As String
Dim iPosEscape As Integer, iPosKeyEnd As Integer
Dim lSeq As Long
Dim sSeqAbert As String
Dim sSeqFech As String
Dim sTipo As String
Dim iPos As Integer, iECF As Integer
Dim objLojaArqFisMestre As New ClassLojaArqFisMestre
Dim colLojaArqFisAnalitico As New Collection
Dim objUltimaReducao As New ClassUltimaReducao

On Error GoTo Erro_Registro_ECF_CC

    sRegistro = colRegistro.Item(1)
    iPosInicio1 = 1
    iRegistro = 1
        
    iPosKeyControl = InStr(sRegistro, Chr(vbKeyControl))

    'se nao encontrar o tipo do registro no arquivo
    If iPosKeyControl = 0 Then gError 110003

    iTipoRegistro = CInt(left(sRegistro, iPosKeyControl - 1))

    Select Case iTipoRegistro

        Case TIPOREGISTROECF_VENDAS
            objVenda.iCodModeloECF = iCodModeloECF
            lErro = Registro_ECF_Vendas(colRegistro, sArquivo, iRegistro, objVenda)
            If lErro <> SUCESSO Then gError 110004
            
'            'se estiver tratando a carga primeira vez no dia do orcamento ==> se o orcamento está vencido ==> nao carrega-o
'            If Not colRegistroParam Is Nothing Then
'
'                If objVenda.objCupomFiscal.dtDataEmissao + objVenda.objCupomFiscal.lDuracao >= (Date - 7) Then colVendas.Add objVenda
'
'            Else
'
'                colVendas.Add objVenda
'
'            End If

            colVendas.Add objVenda

        
        
        Case TIPOREGISTROECF_EXCLUSAOCUPOM
            
            iPosEscape = InStr(iPosInicio1, sRegistro, Chr(vbKeyEscape))
            sSeq1 = Mid(sRegistro, iPosKeyControl + 1, iPosEscape - (iPosKeyControl + 1))
            iPosKeyEnd = InStr(sRegistro, Chr(vbKeyEnd))
            iECF = CInt(Mid(sRegistro, iPosEscape + 1, iPosKeyEnd - (iPosEscape + 1)))
            
            If iECF = giCodECF Then
            
                For iIndice = 1 To colVendas.Count
                    Set objVenda = colVendas.Item(iIndice)
                    If StrParaLong(sSeq1) = objVenda.objCupomFiscal.lNumero Then
                        colVendas.Remove (iIndice)
                        Exit For
                    End If
                Next
                
            End If
            
        Case TIPOREGISTROECF_EXCLUSAOORCAMENTO
            
            iPosEscape = InStr(iPosInicio1, sRegistro, Chr(vbKeyEscape))
            sSeq1 = Mid(sRegistro, iPosKeyControl + 1, iPosEscape - (iPosKeyControl + 1))
            iPosKeyEnd = InStr(sRegistro, Chr(vbKeyEnd))
            iECF = CInt(Mid(sRegistro, iPosEscape + 1, iPosKeyEnd - (iPosEscape + 1)))
            
            If iECF = giCodECF Then
            
                For iIndice = 1 To colVendas.Count
                    Set objVenda = colVendas.Item(iIndice)
                    If objVenda.iTipo <> OPTION_CF And StrParaLong(sSeq1) = objVenda.objCupomFiscal.lNumOrcamento Then
                        colVendas.Remove (iIndice)
                        Exit For
                    End If
                Next
                
            End If
            
'            'se colRegistroParam foi passado como parametro ==> está tratando orcamento ==> tenta retirar o orcamento que está sendo excluido
'            If Not colRegistroParam Is Nothing Then
'
'                sSeqFech = ""
'                sSeqAbert = ""
'
'                'pesquisa em colRegistroParam pelo numero do orcamento, quais os registros da colecao que devem ser excluidos
'                For iIndice = colRegistroParam.Count To 1 Step -1
'
'                    sRegistro1 = colRegistroParam.Item(iIndice)
'
'                    'Procura o Primeiro Control para saber o tipo do registro
'                    iPos = InStr(1, sRegistro1, Chr(vbKeyControl))
'
'                    sTipo = Mid(sRegistro1, 1, iPos - 1)
'
'                    Select Case sTipo
'
'                        'verifica se é de abertura
'                        Case TIPOREGISTROECF_ABRESEQ
'                            'Recolhe o sequencial
'                            sSeqAbert = Mid(sRegistro1, iPos + 1, Len(sRegistro1) - (iPos + 1))
'
'                            If sSeqAbert = sSeq1 Then colRegistroParam.Remove (iIndice)
'
'                        'verifica se é de fechamento
'                        Case TIPOREGISTROECF_ENCERRASEQ
'                            'Recolhe o sequencial
'                            sSeqFech = Mid(sRegistro1, iPos + 1, Len(sRegistro1) - (iPos + 1))
'
'                            If sSeqFech = sSeq1 Then colRegistroParam.Remove (iIndice)
'
'                        Case Else
''                            If sSeqFech = sSeq1 And sSeqFech <> sSeqAbert Then
'                                colRegistroParam.Remove (iIndice)
''                            End If
'
'                    End Select
'
'                Next
'
'            End If
            
        Case TIPOREGISTROECF_MOVIMENTOCAIXA_DINHEIRO
            lErro = Registro_ECF_MovcxDin(colRegistro, sArquivo, iRegistro, objMovimentosCaixa)
            If lErro <> SUCESSO Then gError 110041
            'Atualiza os Dados da Memoria
            'Chama a Função que Atauliza a Memoria
            lErro = MovimentoDinheiro_Atualiza_Memoria(objMovimentosCaixa)
            If lErro <> SUCESSO Then gError 109791
            
        Case TIPOREGISTROECF_MOVIMENTOCAIXA_CHEQUES
            lErro = Registro_ECF_MovcxChq(colRegistro, sArquivo, iRegistro, objMovimentosCaixa, colImfCompl)
            If lErro <> SUCESSO Then gError 110042
            'Atualiza os Dados da Memoria
            lErro = MovimentoCheque_Atualiza_Memoria(objMovimentosCaixa, colImfCompl)
            If lErro <> SUCESSO Then gError 109789
            
        Case TIPOREGISTROECF_MOVIMENTOCAIXA_BOLETO
            lErro = Registro_ECF_MovcxBol(colRegistro, sArquivo, iRegistro, objMovimentosCaixa)
            If lErro <> SUCESSO Then gError 110046
            'Atualiza os Dados da Memoria
            lErro = MovimentoBoleto_Atualiza_Memoria(objMovimentosCaixa)
            If lErro <> SUCESSO Then gError 109786
            
        Case TIPOREGISTROECF_MOVIMENTOCAIXA_TICKETS
            lErro = Registro_ECF_MovcxTkt(colRegistro, sArquivo, iRegistro, objMovimentosCaixa)
            If lErro <> SUCESSO Then gError 110047
            
            'Atualiza os Dados da Memoria
            lErro = MovimentoTicket_Atualiza_Memoria(objMovimentosCaixa)
            If lErro <> SUCESSO Then gError 109797
            
        Case TIPOREGISTROECF_MOVIMENTOCAIXA_OUTROS
            lErro = Registro_ECF_MovcxOut(colRegistro, sArquivo, iRegistro, objMovimentosCaixa)
            If lErro <> SUCESSO Then gError 112091
            'Atualiza os Dados da Memoria
            lErro = MovimentoOutros_Atualiza_Memoria(objMovimentosCaixa)
            If lErro <> SUCESSO Then gError 109795

        Case TIPOREGISTROECF_MOVIMENTOCAIXA
            lErro = Registro_ECF_Movcx(colRegistro, sArquivo, iRegistro, objMovimentosCaixa)
            If lErro <> SUCESSO Then gError 112092
            
            'Atualiza os Movimentos nas coleções globais
            lErro = MovimentoCaixa_GravaMemoria(objMovimentosCaixa)
            If lErro <> SUCESSO Then gError 109805
                        
        Case TIPOREGISTROECF_FECHAMENTOCAIXA
        
            Set objLojaArqFisMestre = New ClassLojaArqFisMestre
            Set colLojaArqFisAnalitico = New Collection
            Set objUltimaReducao = New ClassUltimaReducao
        
            lErro = Registro_ECF_MovcxFechamento(colRegistro, sArquivo, iRegistro, objMovimentosCaixa, objLojaArqFisMestre, colLojaArqFisAnalitico, objUltimaReducao)
            If lErro <> SUCESSO Then gError 204512
            
            'Atualiza os Movimentos nas coleções globais
            lErro = MovimentoCaixa_GravaMemoria(objMovimentosCaixa)
            If lErro <> SUCESSO Then gError 204513
                        
                        
        Case TIPOREGISTROECF_TRANSFCAIXA_INCLUSAO
            lErro = Registro_ECF_TransfCaixa(colRegistro, sArquivo, objTransfCaixa, objChequePara)
            If lErro <> SUCESSO Then gError 109802
                
            lErro = CF_ECF("TransfCaixa_Grava_Memoria", objTransfCaixa, objChequePara)
            If lErro <> SUCESSO Then gError 105340

        Case TIPOREGISTROECF_TRANSFCAIXA_EXCLUSAO
            lErro = Registro_ECF_TransfCaixa_Exclusao(colRegistro, sArquivo, objTransfCaixa)
            If lErro <> SUCESSO Then gError 105342
                
            'busca se já existe uma transfcaixa com o código da tela
            'i.e. busca por um movimento de caixa casado (de/Para) com lTransferencia preenchidos
            lErro = CF_ECF("TransfCaixa_Le", objTransfCaixa)
            If lErro <> SUCESSO And lErro <> 109404 Then gError 105343
                
            If lErro = 109404 Then gError 105344
            
            'chama funcao q ira cuidar de adicionar/subtrair os totais consolidados.
            'apesar do nome induzir, ela nao gravara na memoria, mas sim retirara...
            lErro = CF_ECF("TransfCaixa_Exclui_Memoria", objTransfCaixa)
            If lErro <> SUCESSO Then gError 105345
            
        Case TIPOREGISTROECF_ARQUIVO
                    
            iPosEscape = InStr(iPosInicio1, sRegistro, Chr(vbKeyEnd))
                    
            lSeq = StrParaLong(Mid(sRegistro, iPosKeyControl + 1, iPosEscape - (iPosKeyControl + 1)))
                    
            gcolMarcas.Add lSeq
            
            For Each objMovimentosCaixa In gcolMovimentosCaixa
                objMovimentosCaixa.lNumIntDocLog = MARCADO
            Next
                        
    End Select

    Registro_ECF_CC = SUCESSO

    Exit Function

Erro_Registro_ECF_CC:

    Registro_ECF_CC = gErr

    Select Case gErr
        
        Case 105340, 105342, 105343, 105345, 109785 To 109791, 109794 To 109797, 109802, 109803, 109817, 204512, 204513
        
        Case 105344
            Call Rotina_ErroECF(vbOKOnly, "ERRO_TRANSFCAIXA_NAO_ENCONTRADO", gErr, objTransfCaixa.objMovCaixaDe.lTransferencia)
        
        Case 110003
            Call Rotina_ErroECF(vbOKOnly, "ERRO_AUSENCIA_TIPO_REGISTRO", gErr, iRegistro, sArquivo)

        Case 110004, 110041, 110042, 110046, 110047, 112091, 112092

        Case 110005
            Call Rotina_ErroECF(vbOKOnly, "ERRO_TIPOREGISTRO_NAO_TRATAVEL", gErr, iTipoRegistro, iRegistro, sArquivo)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149400)

    End Select

    Exit Function

End Function

Function Registro_ECF_Vendas(ByVal colRegistro As Collection, ByVal sArquivo As String, ByVal iRegistro As Integer, objVenda As ClassVenda) As Long
'Trata o registro de vendas e coloca os dados no bd do caixa central. sArquivo contem o nome do arquivo em que sRegistro foi lido.
'iRegistro é o numero do registro.

Dim iPosAtual As Integer
Dim iPosFimAtual As Integer
Dim sRegistro As String
Dim iRegistroCol As Integer
Dim lErro As Long
On Error GoTo Erro_Registro_ECF_Vendas

    sRegistro = colRegistro.Item(1)

    iPosAtual = InStr(sRegistro, Chr(vbKeyControl))
    iPosAtual = iPosAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeySeparator))

    objVenda.iTipo = StrParaInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
    objVenda.objCupomFiscal.iTipo = objVenda.iTipo
    iRegistroCol = 1
   
    'guarda as infos de carne se houverem
    lErro = Vendas_Carne(iPosAtual, iPosFimAtual, colRegistro, iRegistroCol, objVenda.objCarne)
    If lErro <> SUCESSO Then gError 110006

    'guarda as infos de cupom
    lErro = Vendas_Cupom(iPosAtual, iPosFimAtual, colRegistro, iRegistroCol, objVenda.objCupomFiscal)
    If lErro <> SUCESSO Then gError 110007
    
    If objVenda.objCupomFiscal.sSATChaveAcesso <> "" Then objVenda.iCodModeloECF = IMPRESSORA_SAT_2_5_15
    If objVenda.objCupomFiscal.sNFeChaveAcesso <> "" And Mid(objVenda.objCupomFiscal.sNFeChaveAcesso, 21, 2) = "65" Then objVenda.iCodModeloECF = IMPRESSORA_NFCE

    'guarda as infos de movimento de caixa
    lErro = Vendas_Movcx(iPosAtual, iPosFimAtual, colRegistro, iRegistroCol, objVenda.colMovimentosCaixa)
    If lErro <> SUCESSO Then gError 110008
    
    'guarda as infos de cheque se houverem
    lErro = Vendas_Cheque(iPosAtual, iPosFimAtual, colRegistro, iRegistroCol, objVenda.colCheques, objVenda)
    If lErro <> SUCESSO Then gError 110009

    'guarda as infos de troca se houverem
    lErro = Vendas_Troca(iPosAtual, iPosFimAtual, colRegistro, iRegistroCol, objVenda.colTroca)
    If lErro <> SUCESSO Then gError 110010

    Registro_ECF_Vendas = SUCESSO

    Exit Function

Erro_Registro_ECF_Vendas:

    Registro_ECF_Vendas = gErr

    Select Case gErr

        Case 110006, 110007, 110008, 110009, 110010, 110011

        Case Else
            lErro = Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149401)

    End Select

    Exit Function

End Function

Function Registro_ECF_MovcxDin(ByVal colRegistro As Collection, ByVal sArquivo As String, ByVal iRegistro As Integer, objMovCx As ClassMovimentoCaixa) As Long
'Trata o registro de movimentacao de dinheiro e coloca os dados no bd do caixa central. sArquivo contem o nome do arquivo em que sRegistro foi lido.
'iRegistro é o numero do registro.

Dim iPosAtual As Integer
Dim iPosFimAtual As Integer
Dim lErro As Long

On Error GoTo Erro_Registro_ECF_MovcxDin

    'guarda as infos de movimento de caixa
    lErro = MovcxDin_Movcx(colRegistro.Item(1), objMovCx)
    If lErro <> SUCESSO Then gError 112158

    Registro_ECF_MovcxDin = SUCESSO

    Exit Function

Erro_Registro_ECF_MovcxDin:

    Registro_ECF_MovcxDin = gErr

    Select Case gErr

        Case 112158

        Case Else
            lErro = Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149402)

    End Select

    Exit Function

End Function


Function Registro_ECF_MovcxBol(ByVal colRegistro As Collection, ByVal sArquivo As String, ByVal iRegistro As Integer, objMovCx As ClassMovimentoCaixa) As Long
'Trata o registro de movimentacao de boletos e coloca os dados no bd do caixa central. sArquivo contem o nome do arquivo em que sRegistro foi lido.
'iRegistro é o numero do registro.

Dim lErro As Long

On Error GoTo Erro_Registro_ECF_MovcxBol

    'guarda as infos de movimento de caixa
    lErro = MovcxBol_Movcx(colRegistro, objMovCx)
    If lErro <> SUCESSO Then gError 112240

    Registro_ECF_MovcxBol = SUCESSO

    Exit Function

Erro_Registro_ECF_MovcxBol:

    Registro_ECF_MovcxBol = gErr

    Select Case gErr

        Case 112240

        Case Else
            lErro = Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149403)

    End Select

    Exit Function

End Function

Function Registro_ECF_MovcxChq(ByVal colRegistro As Collection, ByVal sArquivo As String, ByVal iRegistro As Integer, objMovCx As ClassMovimentoCaixa, colCheque As Collection) As Long
'Trata o registro de movimentacao de cheques e coloca os dados no bd do caixa central. sArquivo contem o nome do arquivo em que sRegistro foi lido.
'iRegistro é o numero do registro.

Dim iPosAtual As Integer
Dim iPosFimAtual As Integer
Dim sRegistro As String
Dim lErro As Long
Dim iRegistroCol As Integer

On Error GoTo Erro_Registro_ECF_MovcxChq

    sRegistro = colRegistro(1)

    iPosFimAtual = InStr(sRegistro, Chr(vbKeyControl))

    'guarda as infos de movimento de caixa
    lErro = MovcxChq_Movcx(iPosAtual, iPosFimAtual, sRegistro, objMovCx)
    If lErro <> SUCESSO Then gError 112192

    If objMovCx.iTipo = MOVIMENTOCAIXA_TRANSF_SUP_CHEQUE Then
    
        iRegistroCol = 1
    
        iPosAtual = iPosFimAtual + 1
        iPosFimAtual = iPosAtual
    
        'guarda as infos de cheque se houverem
        lErro = Vendas_Cheque(iPosAtual, iPosFimAtual, colRegistro, iRegistroCol, colCheque)
        If lErro <> SUCESSO Then gError 105422

    Else

        'guarda as infos de cheque
        lErro = MovcxChq_Cheque(iPosAtual, iPosFimAtual, colRegistro, colCheque)
        If lErro <> SUCESSO Then gError 112193

    End If

    Registro_ECF_MovcxChq = SUCESSO

    Exit Function

Erro_Registro_ECF_MovcxChq:

    Registro_ECF_MovcxChq = gErr

    Select Case gErr

        Case 105422, 112192, 112193

        Case Else
            lErro = Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149404)

    End Select

    Exit Function

End Function

Function Registro_ECF_Movcx(ByVal colRegistro As Collection, ByVal sArquivo As String, ByVal iRegistro As Integer, objMovCx As ClassMovimentoCaixa) As Long
'Trata o registro de movimentacao de caixa e coloca os dados no bd do caixa central. sArquivo contem o nome do arquivo em que sRegistro foi lido.
'iRegistro é o numero do registro.

Dim lErro As Long

On Error GoTo Erro_Registro_ECF_Movcx

    'guarda as infos de movimento de caixa
    lErro = Movcx_Movcx(colRegistro, objMovCx)
    If lErro <> SUCESSO Then gError 112342

    Registro_ECF_Movcx = SUCESSO

    Exit Function

Erro_Registro_ECF_Movcx:

    Registro_ECF_Movcx = gErr

    Select Case gErr

        Case 112342

        Case Else
            lErro = Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149405)

    End Select

    Exit Function

End Function

Function Registro_ECF_MovcxOut(ByVal colRegistro As Collection, ByVal sArquivo As String, ByVal iRegistro As Integer, objMovCx As ClassMovimentoCaixa) As Long
'Trata o registro de movimentacao de outros e coloca os dados no bd do caixa central. sArquivo contem o nome do arquivo em que sRegistro foi lido.
'iRegistro é o numero do registro.

Dim lErro As Long

On Error GoTo Erro_Registro_ECF_MovcxOut

    'guarda as infos de movimento de caixa
    lErro = MovcxOut_Movcx(colRegistro, objMovCx)
    If lErro <> SUCESSO Then gError 112308
    
    Registro_ECF_MovcxOut = SUCESSO

    Exit Function

Erro_Registro_ECF_MovcxOut:

    Registro_ECF_MovcxOut = gErr

    Select Case gErr

        Case 112308

        Case Else
            lErro = Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149406)

    End Select

    Exit Function

End Function

Function Registro_ECF_MovcxTkt(ByVal colRegistro As Collection, ByVal sArquivo As String, ByVal iRegistro As Integer, objMovCx As ClassMovimentoCaixa) As Long
'Trata o registro de movimentacao de ticket e coloca os dados no bd do caixa central. sArquivo contem o nome do arquivo em que sRegistro foi lido.
'iRegistro é o numero do registro.

Dim lErro As Long

On Error GoTo Erro_Registro_ECF_MovcxTkt

    'guarda as infos de movimento de caixa
    lErro = MovcxTkt_Movcx(colRegistro, objMovCx)
    If lErro <> SUCESSO Then gError 112274

    Registro_ECF_MovcxTkt = SUCESSO

    Exit Function

Erro_Registro_ECF_MovcxTkt:

    Registro_ECF_MovcxTkt = gErr

    Select Case gErr

        Case 112274

        Case Else
            lErro = Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149407)

    End Select

    Exit Function

End Function

Function Vendas_Carne(iPosAtual As Integer, iPosFimAtual As Integer, ByVal colRegistro As Collection, iRegistroCol As Integer, ByVal objCarne As ClassCarne) As Long
'guarda as infos de carne se houverem

Dim objCarneParc As ClassCarneParcelas
Dim lErro As Long
Dim sRegistro As String
Dim iIndice As Integer

On Error GoTo Erro_Vendas_Carne
        
    sRegistro = colRegistro.Item(iRegistroCol)
'
    'pula o primeiro separator
    iPosAtual = iPosFimAtual + 1
    
    iPosFimAtual = iPosFimAtual + 1
    
    'se o proximo caracter nao for um separator ==> gravar carne
    If Mid(sRegistro, iPosAtual, 1) <> Chr(vbKeySeparator) Then

        iPosAtual = iPosFimAtual + 1
        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
        objCarne.dtDataReferencia = CDate(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
    
        iPosAtual = iPosFimAtual + 1
        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
        objCarne.iFilialEmpresa = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
    
        iPosAtual = iPosFimAtual + 1
        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
        objCarne.iStatus = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
    
        iPosAtual = iPosFimAtual + 1
        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
        objCarne.lCliente = CLng(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
    
        iPosAtual = iPosFimAtual + 1
        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
        objCarne.sAutorizacao = Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual)
    
        iPosAtual = iPosFimAtual + 1
        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
        objCarne.sCodBarrasCarne = Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual)
    
        iPosAtual = iPosFimAtual + 1
        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyShift))
        iPosAtual = iPosFimAtual
    
        'Se o proximo caracter apos o vbKeyShift nao for vbKeyControl ou vbKeySeparator ==> tem pagamento de carne envolvido
        If Mid(sRegistro, iPosAtual + 1, 1) <> Chr(vbKeyControl) Then
    
            For iIndice = iRegistroCol To colRegistro.Count
    
                sRegistro = colRegistro.Item(iIndice)
    
                Do While Mid(sRegistro, iPosAtual, 1) = Chr(vbKeyShift)
    
                    Set objCarneParc = New ClassCarneParcelas
    
                    objCarne.colParcelas.Add objCarneParc
    
                    iPosAtual = iPosFimAtual + 1
                    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                    objCarneParc.dtDataVencimento = CDate(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
    
                    iPosAtual = iPosFimAtual + 1
                    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                    objCarneParc.dValor = CDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
    
                    iPosAtual = iPosFimAtual + 1
                    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                    objCarneParc.iParcela = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
    
                    iPosAtual = iPosFimAtual + 1
                    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                    objCarneParc.iStatus = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
    
                    iPosAtual = iPosFimAtual + 2
                    iPosFimAtual = iPosFimAtual + 2
                    
                Loop
    
                'se saiu antes de terminar o registro ==> as infos do carne terminaram
                If iPosAtual <= Len(sRegistro) Then Exit For
    
                iPosAtual = 1
                iPosFimAtual = 1
    
            Next
    
            iRegistroCol = iIndice
    
        End If

    End If

    'pula o key separator
    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = iPosFimAtual + 1

    Vendas_Carne = SUCESSO

    Exit Function

Erro_Vendas_Carne:

    Vendas_Carne = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149408)

    End Select

    Exit Function
     
End Function

Function Vendas_Cupom(iPosAtual As Integer, iPosFimAtual As Integer, ByVal colRegistro As Collection, iRegistroCol As Integer, ByVal objCupomFiscal As ClassCupomFiscal) As Long
'guarda as infos de cupom

Dim sRegistro As String
Dim lErro As Long
Dim iIndiceFinal As Integer
Dim iPosKeyShift As Integer
On Error GoTo Erro_Vendas_Cupom
    
    sRegistro = colRegistro.Item(iRegistroCol)
        
    iIndiceFinal = colRegistro.Count
        
    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objCupomFiscal.dtDataEmissao = CDate(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objCupomFiscal.dHoraEmissao = CDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objCupomFiscal.dValorAcrescimo = CDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objCupomFiscal.dValorDesconto = CDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objCupomFiscal.dValorProdutos = CDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objCupomFiscal.dValorTotal = CDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objCupomFiscal.iECF = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objCupomFiscal.iFilialEmpresa = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objCupomFiscal.iStatus = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objCupomFiscal.iTabelaPreco = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
    
    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objCupomFiscal.lGerenteCancel = CLng(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objCupomFiscal.lNumero = CLng(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objCupomFiscal.lNumOrcamento = CLng(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objCupomFiscal.iVendedor = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
    
    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objCupomFiscal.lCliente = StrParaLong(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objCupomFiscal.sCPFCGC = Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual)

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objCupomFiscal.sNomeCliente = Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual)

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objCupomFiscal.sMotivoCancel = Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual)

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objCupomFiscal.sNaturezaOp = Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual)
    
    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objCupomFiscal.lDuracao = CLng(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
    
    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    
    If iPosFimAtual = 0 Then
        
        If iRegistroCol + 1 <= iIndiceFinal Then
                    
            iRegistroCol = iRegistroCol + 1
            sRegistro = colRegistro.Item(iRegistroCol)
            iPosAtual = 1
            iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                    
        End If
            
    End If
    
    
    If iPosFimAtual <> 0 Then
    
        objCupomFiscal.iCodCaixa = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
    
        iPosAtual = iPosFimAtual + 1
        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
        objCupomFiscal.lCOOCupomOrigDAV = CLng(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
    
        iPosAtual = iPosFimAtual + 1
        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
        objCupomFiscal.sTipoECF = Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual)
    
        iPosAtual = iPosFimAtual + 1
        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
        objCupomFiscal.sMarcaECF = Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual)
    
        iPosAtual = iPosFimAtual + 1
        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
        objCupomFiscal.sModeloECF = Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual)
    
        iPosAtual = iPosFimAtual + 1
        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
        objCupomFiscal.sNumSerieECF = Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual)
    
        iPosAtual = iPosFimAtual + 1
        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
        objCupomFiscal.lCCF = CLng(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
    
        iPosAtual = iPosFimAtual + 1
        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
        objCupomFiscal.lNumeroDAV = CLng(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
    
        iPosAtual = iPosFimAtual + 1
        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
        objCupomFiscal.sCPFCGC1 = Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual)
    
        iPosAtual = iPosFimAtual + 1
        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
        objCupomFiscal.sEndereco = Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual)
    
        iPosAtual = iPosFimAtual + 1
        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
        objCupomFiscal.iSequencialECF = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
    
        iPosAtual = iPosFimAtual + 1
        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
        objCupomFiscal.iSequencialECFOrigDAV = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
    
        iPosAtual = iPosFimAtual + 1
        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
        objCupomFiscal.iDAVImpresso = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
    
        iPosAtual = iPosFimAtual + 1
        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
        objCupomFiscal.dtDataReducao = StrParaDate(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
        
        iPosAtual = iPosFimAtual + 1
        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    
        If iPosFimAtual = 0 Then
        
            If iRegistroCol + 1 <= iIndiceFinal Then
                    
                    iRegistroCol = iRegistroCol + 1
                    sRegistro = colRegistro.Item(iRegistroCol)
                    iPosAtual = 1
                    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                    
            End If
            
        End If
    
    
        iPosKeyShift = InStr(iPosAtual, sRegistro, Chr(vbKeyShift))
        If iPosKeyShift = 0 Then iPosKeyShift = 501
                    
    
        If iPosFimAtual <> 0 And iPosFimAtual < iPosKeyShift Then
        
            objCupomFiscal.sNFeChaveAcesso = Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual)
    
            If Len(Trim(objCupomFiscal.sNFeChaveAcesso)) <> 0 Then
            
                If Len(Trim(objCupomFiscal.sNFeChaveAcesso)) = 44 And Mid(objCupomFiscal.sNFeChaveAcesso, 21, 2) = "65" Then
                
                    'pegar dados de nfceinfo como o nprot, para poder cancelar, se necessário
                    lErro = NFCE_ObterInfo(objCupomFiscal)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                Else
                
                    objCupomFiscal.sSATChaveAcesso = objCupomFiscal.sNFeChaveAcesso
                    objCupomFiscal.sNFeChaveAcesso = ""
                    
                    lErro = SAT_ObterInfo(objCupomFiscal)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    
                End If
                
            End If
    
            iPosAtual = iPosFimAtual + 1
            iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
            objCupomFiscal.sNFeArqXml = Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual)
        
            iPosAtual = iPosFimAtual + 1
            iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
        
            If iPosFimAtual = 0 Then
            
                If iRegistroCol + 1 <= iIndiceFinal Then
                        
                        iRegistroCol = iRegistroCol + 1
                        sRegistro = colRegistro.Item(iRegistroCol)
                        iPosAtual = 1
                        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                        
                End If

            End If

            iPosKeyShift = InStr(iPosAtual, sRegistro, Chr(vbKeyShift))
            If iPosKeyShift = 0 Then iPosKeyShift = 501
        
            If iPosFimAtual <> 0 And iPosFimAtual < iPosKeyShift Then
            
                objCupomFiscal.objNF.sSerie = Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual)
        
                iPosAtual = iPosFimAtual + 1
                iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                objCupomFiscal.objNF.lNumNotaFiscal = StrParaLong(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
        
                iPosAtual = iPosFimAtual + 1
                iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                objCupomFiscal.objNF.dtDataEmissao = StrParaDate(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
        
                iPosAtual = iPosFimAtual + 1
                iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                objCupomFiscal.objNF.sDestino = Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual)
        
            End If
        
        End If

        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyShift))
        iPosAtual = iPosFimAtual
        

    End If

    'guarda as infos dos itens do cupom
    lErro = Vendas_ItensCF(iPosAtual, iPosFimAtual, colRegistro, iRegistroCol, objCupomFiscal)
    If lErro <> SUCESSO Then gError 112157

    Vendas_Cupom = SUCESSO

    Exit Function

Erro_Vendas_Cupom:

    Vendas_Cupom = gErr

    Select Case gErr
        
        Case 112157, ERRO_SEM_MENSAGEM
        
        Case Else
            lErro = Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149409)

    End Select

    Exit Function

End Function

Function Vendas_ItensCF(iPosAtual As Integer, iPosFimAtual As Integer, ByVal colRegistro As Collection, iRegistroCol As Integer, ByVal objCupomFiscal As ClassCupomFiscal) As Long
'guarda as infos dos itens do cupom

Dim objItens As ClassItemCupomFiscal
Dim lErro As Long
Dim iIndiceFinal As Integer
Dim sRegistro As String
Dim iPosKeyControl As Integer
Dim iPosKeyShift As Integer
Dim iPosKeySeparator As Integer
Dim iPosFimAtual1 As Integer
Dim objProduto As New ClassProduto
Dim sProduto As String

On Error GoTo Erro_Vendas_ItensCF
  
    iIndiceFinal = colRegistro.Count
  
    sRegistro = colRegistro.Item(iRegistroCol)

    If iPosFimAtual = 0 Then
    
        If iRegistroCol + 1 <= iIndiceFinal Then
                
                iRegistroCol = iRegistroCol + 1
                sRegistro = colRegistro.Item(iRegistroCol)
                iPosAtual = 1
                iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyShift))
                iPosAtual = iPosFimAtual
                
        End If
        
    End If



    'Se o proximo caracter apos o vbKeyShift nao for vbKeyControl ou vbKeySeparator ==> tem itens de cupom envolvidos
    If Mid(sRegistro, iPosAtual + 1, 1) <> Chr(vbKeyControl) And Mid(sRegistro, iPosAtual + 1, 1) <> Chr(vbKeySeparator) Then


'        For iIndice = iRegistroCol To colRegistro.Count
        
        Do While iRegistroCol <= iIndiceFinal

            Do While Mid(sRegistro, iPosAtual, 1) = Chr(vbKeyShift)

                Set objItens = New ClassItemCupomFiscal

                objCupomFiscal.colItens.Add objItens

                objItens.iFilialEmpresa = objCupomFiscal.iFilialEmpresa

                iPosAtual = iPosFimAtual + 1
                iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                objItens.dAliquotaICMS = CDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

                iPosAtual = iPosFimAtual + 1
                iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                objItens.dPercDesc = CDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

                iPosAtual = iPosFimAtual + 1
                iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                objItens.dPrecoUnitario = CDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

                iPosAtual = iPosFimAtual + 1
                iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                objItens.dQuantidade = CDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

                iPosAtual = iPosFimAtual + 1
                iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                objItens.dValorDesconto = CDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

                iPosAtual = iPosFimAtual + 1
                iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                objItens.iStatus = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

                iPosAtual = iPosFimAtual + 1
                iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                objItens.sProduto = Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual)

                sProduto = objItens.sProduto
                
                '??? Jones 28/08/16
                '??? Call TP_Produto_Le_Col2(gaobjProdutosCodigo, sProduto, objProduto)
                Set objProduto = gaobjProdutosCodigo.Busca(sProduto)
    
                'Não encontrou o Produto
                If objProduto Is Nothing Then Set objProduto = New ClassProduto
                '??? Jones 28/08/16

                objItens.sProdutoNomeRed = objProduto.sNomeReduzido
                objItens.sProdutoDescricao = objProduto.sDescricao

                iPosAtual = iPosFimAtual + 1
                iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                objItens.sUnidadeMed = Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual)
                
                iPosAtual = iPosFimAtual + 1
                iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                objItens.sSituacaoTrib = Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual)
                
                iPosAtual = iPosFimAtual + 1
                iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                objItens.iCodCaixa = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
                
                iPosAtual = iPosFimAtual + 1
                iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                objItens.iItem = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
                
                iPosFimAtual1 = InStr(iPosFimAtual + 1, sRegistro, Chr(vbKeyEscape))
                
                If iPosFimAtual1 = 0 Then
                    
                        If iRegistroCol + 1 <= iIndiceFinal Then
                                
                                iRegistroCol = iRegistroCol + 1
                                sRegistro = colRegistro.Item(iRegistroCol)
                                iPosAtual = 1
                                iPosFimAtual = 0
                                iPosFimAtual1 = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                                
                        End If
                        
                End If
                
                
                iPosKeyShift = InStr(iPosAtual, sRegistro, Chr(vbKeyShift))
                If iPosKeyShift = 0 Then iPosKeyShift = 501
                
                iPosKeyControl = InStr(iPosAtual, sRegistro, Chr(vbKeyControl))
                If iPosKeyControl = 0 Then iPosKeyControl = 501
                
                iPosKeySeparator = InStr(iPosAtual, sRegistro, Chr(vbKeySeparator))
                If iPosKeySeparator = 0 Then iPosKeySeparator = 501
                
                If iPosFimAtual1 <> 0 And iPosFimAtual1 < iPosKeyControl And iPosFimAtual1 < iPosKeySeparator And iPosFimAtual1 < iPosKeyShift Then
                
                    iPosAtual = iPosFimAtual + 1
                    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                    objItens.objTributacaoDocItem.dQuantidade = StrParaDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
                    
                    iPosAtual = iPosFimAtual + 1
                    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                    objItens.objTributacaoDocItem.dPrecoUnitario = StrParaDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
                    
                    iPosAtual = iPosFimAtual + 1
                    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                    objItens.objTributacaoDocItem.sNaturezaOp = Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual)
                    
                    iPosAtual = iPosFimAtual + 1
                    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                    objItens.objTributacaoDocItem.iTipoTributacao = StrParaInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
                    
                    iPosAtual = iPosFimAtual + 1
                    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                    objItens.objTributacaoDocItem.iICMSTipo = StrParaInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
                    
                    iPosAtual = iPosFimAtual + 1
                    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                    objItens.objTributacaoDocItem.dICMSBase = StrParaDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
                    
                    iPosAtual = iPosFimAtual + 1
                    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                    objItens.objTributacaoDocItem.dICMSPercRedBase = StrParaDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
                    
                    iPosAtual = iPosFimAtual + 1
                    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                    objItens.objTributacaoDocItem.dICMSAliquota = StrParaDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
                    
                    iPosAtual = iPosFimAtual + 1
                    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                    objItens.objTributacaoDocItem.dICMSValor = StrParaDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
                    
                    iPosAtual = iPosFimAtual + 1
                    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                    objItens.objTributacaoDocItem.iPISTipo = StrParaInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
                    
                    iPosAtual = iPosFimAtual + 1
                    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                    objItens.objTributacaoDocItem.iPISTipoCalculo = StrParaInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
                    
                    iPosAtual = iPosFimAtual + 1
                    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                    objItens.objTributacaoDocItem.dPISBase = StrParaDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
                    
                    iPosAtual = iPosFimAtual + 1
                    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                    objItens.objTributacaoDocItem.dPISAliquota = StrParaDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
                
                    iPosAtual = iPosFimAtual + 1
                    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                    objItens.objTributacaoDocItem.dPISAliquotaValor = StrParaDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
                
                    iPosAtual = iPosFimAtual + 1
                    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                    objItens.objTributacaoDocItem.dPISQtde = StrParaDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
                
                    iPosAtual = iPosFimAtual + 1
                    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                    objItens.objTributacaoDocItem.dPISValor = StrParaDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
                
                    iPosAtual = iPosFimAtual + 1
                    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                    objItens.objTributacaoDocItem.iPISSTTipoCalculo = StrParaInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
                    
                    iPosAtual = iPosFimAtual + 1
                    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                    objItens.objTributacaoDocItem.dPISSTBase = StrParaDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
                    
                    iPosAtual = iPosFimAtual + 1
                    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                    objItens.objTributacaoDocItem.dPISSTAliquota = StrParaDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
                
                    iPosAtual = iPosFimAtual + 1
                    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                    objItens.objTributacaoDocItem.dPISSTAliquotaValor = StrParaDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
                
                    iPosAtual = iPosFimAtual + 1
                    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                    objItens.objTributacaoDocItem.dPISSTQtde = StrParaDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
                
                    iPosAtual = iPosFimAtual + 1
                    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                    objItens.objTributacaoDocItem.dPISSTValor = StrParaDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
                
                
                    iPosAtual = iPosFimAtual + 1
                    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                    objItens.objTributacaoDocItem.iCOFINSTipo = StrParaInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
                    
                    iPosAtual = iPosFimAtual + 1
                    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                    objItens.objTributacaoDocItem.iCOFINSTipoCalculo = StrParaInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
                    
                    iPosAtual = iPosFimAtual + 1
                    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                    objItens.objTributacaoDocItem.dCOFINSBase = StrParaDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
                    
                    iPosAtual = iPosFimAtual + 1
                    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                    objItens.objTributacaoDocItem.dCOFINSAliquota = StrParaDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
                
                    iPosAtual = iPosFimAtual + 1
                    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                    objItens.objTributacaoDocItem.dCOFINSAliquotaValor = StrParaDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
                
                    iPosAtual = iPosFimAtual + 1
                    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                    objItens.objTributacaoDocItem.dCOFINSQtde = StrParaDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
                
                    iPosAtual = iPosFimAtual + 1
                    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                    objItens.objTributacaoDocItem.dCOFINSValor = StrParaDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
                
                    iPosAtual = iPosFimAtual + 1
                    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                    objItens.objTributacaoDocItem.iCOFINSSTTipoCalculo = StrParaInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
                    
                    iPosFimAtual1 = InStr(iPosFimAtual + 1, sRegistro, Chr(vbKeyEscape))
                    
                    If iPosFimAtual1 = 0 Then
                        
                            If iRegistroCol + 1 <= iIndiceFinal Then
                                    
                                    iRegistroCol = iRegistroCol + 1
                                    sRegistro = colRegistro.Item(iRegistroCol)
                                    iPosAtual = 1
                                    iPosFimAtual = 0
                                    iPosFimAtual1 = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                                    
                            End If
                            
                    End If
                    
                    
                    If iPosFimAtual1 <> 0 Then
                    
                    
                        iPosAtual = iPosFimAtual + 1
                        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                        objItens.objTributacaoDocItem.dCOFINSSTBase = StrParaDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
                        
                        iPosAtual = iPosFimAtual + 1
                        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                        objItens.objTributacaoDocItem.dCOFINSSTAliquota = StrParaDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
                    
                        iPosAtual = iPosFimAtual + 1
                        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                        objItens.objTributacaoDocItem.dCOFINSSTAliquotaValor = StrParaDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
                    
                        iPosAtual = iPosFimAtual + 1
                        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                        objItens.objTributacaoDocItem.dCOFINSSTQtde = StrParaDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
                    
                        iPosAtual = iPosFimAtual + 1
                        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                        objItens.objTributacaoDocItem.dCOFINSSTValor = StrParaDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
                    
                    
                        iPosAtual = iPosFimAtual + 1
                        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                        objItens.objTributacaoDocItem.sCST = Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual)
                    
                        iPosAtual = iPosFimAtual + 1
                        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                        objItens.objTributacaoDocItem.sISSQN = Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual)
                    
                        iPosAtual = iPosFimAtual + 1
                        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                        objItens.objTributacaoDocItem.dISSBase = StrParaDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
                    
                        iPosAtual = iPosFimAtual + 1
                        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                        objItens.objTributacaoDocItem.dISSAliquota = StrParaDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
                    
                        iPosAtual = iPosFimAtual + 1
                        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                        objItens.objTributacaoDocItem.sISSCidadeIBGE = Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual)
                    
                        iPosAtual = iPosFimAtual + 1
                        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                        objItens.objTributacaoDocItem.dISSValor = StrParaDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
                    
                        iPosAtual = iPosFimAtual + 1
                        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                        objItens.objTributacaoDocItem.iICMSBaseModalidade = StrParaInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
                    
                        iPosAtual = iPosFimAtual + 1
                        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                        objItens.objTributacaoDocItem.iOrigemMercadoria = StrParaInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
                    
                        iPosAtual = iPosFimAtual + 1
                        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                        objItens.objTributacaoDocItem.sGenero = Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual)
                    
                        iPosAtual = iPosFimAtual + 1
                        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                        objItens.objTributacaoDocItem.dQtdTrib = StrParaDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
                    
                        iPosAtual = iPosFimAtual + 1
                        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                        objItens.objTributacaoDocItem.sUMTrib = Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual)
                    
                        iPosAtual = iPosFimAtual + 1
                        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                        objItens.objTributacaoDocItem.dValorUnitTrib = StrParaDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
                    
                        iPosAtual = iPosFimAtual + 1
                        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                        objItens.objTributacaoDocItem.sISSCST = Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual)
                    
                        iPosAtual = iPosFimAtual + 1
                        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                        objItens.objTributacaoDocItem.dICMSValorIsento = StrParaInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
                    
                        iPosAtual = iPosFimAtual + 1
                        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                        objItens.objTributacaoDocItem.iICMSMotivo = StrParaInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
                    
                        iPosAtual = iPosFimAtual + 1
                        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                        objItens.objTributacaoDocItem.iICMSSimplesTipo = StrParaInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
                    
                        iPosAtual = iPosFimAtual + 1
                        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                        objItens.objTributacaoDocItem.iISSTipo = StrParaInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
                    
                        iPosAtual = iPosFimAtual + 1
                        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                        objItens.objTributacaoDocItem.iRegimeTributario = StrParaInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
                    
                        iPosAtual = iPosFimAtual + 1
                        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                        objItens.objTributacaoDocItem.sCSOSN = Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual)
                        
                        iPosAtual = iPosFimAtual + 1
                        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                        objItens.objTributacaoDocItem.dTotTrib = StrParaDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
                    
                        iPosAtual = iPosFimAtual + 1
                        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                        objItens.objTributacaoDocItem.iTotTribTipo = StrParaInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
                        
                    Else
                    
                '        iPosFimAtual = iPosAtual
                    
                    End If
                
                Else
            
                 '   iPosFimAtual = iPosAtual
                    
                End If
                
                
                iPosAtual = iPosFimAtual + 1
                iPosFimAtual = iPosFimAtual + 1
                
                
            Loop

            'se nao esta no final do registro significa que nao encontrou mais itens e pode sair do loop de itens e procurar os movimentos de caixa
            If iPosAtual <= Len(sRegistro) Then Exit Do

            'se esta no final do registro ==> pega o proximo e pula os caracteres iniciais

            iRegistroCol = iRegistroCol + 1
            sRegistro = colRegistro.Item(iRegistroCol)

            iPosAtual = InStr(1, sRegistro, Chr(vbKeyLButton)) + 1
            iPosFimAtual = InStr(1, sRegistro, Chr(vbKeyLButton)) + 1

'        Next
        Loop

    Else
    
        iPosAtual = iPosAtual + 1


    End If

    Vendas_ItensCF = SUCESSO

    Exit Function

Erro_Vendas_ItensCF:
    
    Vendas_ItensCF = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149410)

    End Select

    Exit Function

End Function

Function Vendas_Movcx(iPosAtual As Integer, iPosFimAtual As Integer, ByVal colRegistro As Collection, iRegistroCol As Integer, ByVal colMovCx As Collection) As Long
'guarda as infos de movimento de caixa

Dim lErro As Long
Dim objMovCx As ClassMovimentoCaixa
Dim sRegistro As String
Dim iPosAux As Integer
Dim iPosAux1 As Integer

On Error GoTo Erro_Vendas_Movcx

    sRegistro = colRegistro.Item(iRegistroCol)

    'Se o proximo caracter for o vbKeyControl ==> tem movimentos de caixa envolvidos
    If Mid(sRegistro, iPosAtual, 1) = Chr(vbKeyControl) Then

 '       For iIndice = iRegistroCol To colRegistro.Count
        Do While iRegistroCol <= colRegistro.Count

            Do While Mid(sRegistro, iPosAtual, 1) = Chr(vbKeyControl)
                
                Set objMovCx = New ClassMovimentoCaixa
                
                colMovCx.Add objMovCx
                
                iPosAtual = iPosFimAtual + 1
                iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                objMovCx.dHora = CDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

                iPosAtual = iPosFimAtual + 1
                iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                objMovCx.dtDataMovimento = StrParaDate(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

                iPosAtual = iPosFimAtual + 1
                iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                objMovCx.dValor = CDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

                iPosAtual = iPosFimAtual + 1
                iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                objMovCx.iAdmMeioPagto = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

                iPosAtual = iPosFimAtual + 1
                iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                objMovCx.iCaixa = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

                iPosAtual = iPosFimAtual + 1
                iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                objMovCx.iCodConta = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

                iPosAtual = iPosFimAtual + 1
                iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                objMovCx.iCodOperador = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

                iPosAtual = iPosFimAtual + 1
                iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                objMovCx.iFilialEmpresa = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

                iPosAtual = iPosFimAtual + 1
                iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                objMovCx.iGerente = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

                iPosAtual = iPosFimAtual + 1
                iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                objMovCx.iParcelamento = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

                iPosAtual = iPosFimAtual + 1
                iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                objMovCx.iTipo = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

                iPosAtual = iPosFimAtual + 1
                iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                objMovCx.iTipoCartao = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

                iPosAtual = iPosFimAtual + 1
                iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                objMovCx.lCupomFiscal = CLng(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

                iPosAtual = iPosFimAtual + 1
                iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                objMovCx.lNumero = CLng(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

                iPosAtual = iPosFimAtual + 1
                iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                objMovCx.lNumRefInterna = CLng(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

                iPosAtual = iPosFimAtual + 1
                iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                objMovCx.lSequencial = CLng(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

                iPosAtual = iPosFimAtual + 1
                iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                objMovCx.sFavorecido = Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual)

                iPosAtual = iPosFimAtual + 1
                iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                objMovCx.sHistorico = Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual)

                
                iPosAux = InStr(iPosFimAtual + 1, sRegistro, Chr(vbKeySeparator))
                iPosAux1 = InStr(iPosFimAtual + 1, sRegistro, Chr(vbKeyEscape))
                
                If (iPosAux1 <> 0 And iPosAux1 < iPosAux) Or (iPosAux1 <> 0 And iPosAux = 0) Then
                
                    iPosAtual = iPosFimAtual + 1
                    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                    objMovCx.dtDataPreDatado = StrParaDate(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
    
                    iPosAux = InStr(iPosFimAtual + 1, sRegistro, Chr(vbKeySeparator))
                    iPosAux1 = InStr(iPosFimAtual + 1, sRegistro, Chr(vbKeyEscape))
                    
                    If (iPosAux1 <> 0 And iPosAux1 < iPosAux) Or (iPosAux1 <> 0 And iPosAux = 0) Then
        
                        iPosAtual = iPosFimAtual + 1
                        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                        objMovCx.sAutorizacao = Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual)


                    End If

                End If


                iPosAtual = iPosFimAtual + 1
                iPosFimAtual = iPosFimAtual + 1
                
            Loop

            If iPosAtual <= Len(sRegistro) Then Exit Do

'            iPosAtual = 1
'            iPosFimAtual = 1

            iRegistroCol = iRegistroCol + 1
            sRegistro = colRegistro.Item(iRegistroCol)

            iPosAtual = InStr(1, sRegistro, Chr(vbKeyLButton)) + 1
            iPosFimAtual = InStr(1, sRegistro, Chr(vbKeyLButton)) + 1


        Loop

    End If

    'pula o vbKeySeparator
    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = iPosFimAtual + 1

    Vendas_Movcx = SUCESSO

    Exit Function

Erro_Vendas_Movcx:

    Vendas_Movcx = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149411)

    End Select

    Exit Function

End Function

Function Vendas_Cheque(iPosAtual As Integer, iPosFimAtual As Integer, ByVal colRegistro As Collection, iRegistroCol As Integer, ByVal colCheque As Collection, Optional objVenda As ClassVenda) As Long
'guarda as infos de movimento de cheque

Dim lErro As Long
Dim objCheque As ClassChequePre
Dim sRegistro As String

On Error GoTo Erro_Vendas_Cheque

    sRegistro = colRegistro.Item(iRegistroCol)

    'Se o proximo caracter for o vbKeyControl ==> tem movimentos de caixa envolvidos
    If Mid(sRegistro, iPosAtual, 1) = Chr(vbKeyControl) Then

        Do While iRegistroCol <= colRegistro.Count
    
            Do While Mid(sRegistro, iPosAtual, 1) = Chr(vbKeyControl)
    
                Set objCheque = New ClassChequePre
    
                colCheque.Add objCheque
                
                If objVenda Is Nothing Then
                    gcolCheque.Add objCheque
                Else
                    If (objVenda.iTipo <> OPTION_DAV And objVenda.iTipo <> OPTION_PREVENDA) Or objVenda.objCupomFiscal.iStatus = STATUS_BAIXADO Then
                        gcolCheque.Add objCheque
                    End If
                End If
                
                iPosAtual = iPosFimAtual + 1
                iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                objCheque.dtDataDeposito = CDate(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
    
                iPosAtual = iPosFimAtual + 1
                iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                objCheque.dValor = CDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
                gdSaldocheques = gdSaldocheques + objCheque.dValor
            
                iPosAtual = iPosFimAtual + 1
                iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                objCheque.iAprovado = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
    
                iPosAtual = iPosFimAtual + 1
                iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                objCheque.iBanco = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
    
                iPosAtual = iPosFimAtual + 1
                iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                objCheque.iFilialEmpresaLoja = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
    
                iPosAtual = iPosFimAtual + 1
                iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                objCheque.iNaoEspecificado = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
    
                iPosAtual = iPosFimAtual + 1
                iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                objCheque.lNumero = CLng(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
    
                iPosAtual = iPosFimAtual + 1
                iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                objCheque.sAgencia = Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual)
    
                iPosAtual = iPosFimAtual + 1
                iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                objCheque.sContaCorrente = Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual)
    
                iPosAtual = iPosFimAtual + 1
                iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                objCheque.sCPFCGC = Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual)
    
                iPosAtual = iPosFimAtual + 1
                iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                objCheque.lSequencialCaixa = CLng(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
    
                iPosAtual = iPosFimAtual + 1
                iPosFimAtual = iPosFimAtual + 1
    
            Loop
    
            If iPosAtual <= Len(sRegistro) Then Exit Do
    
    
            iRegistroCol = iRegistroCol + 1
            sRegistro = colRegistro.Item(iRegistroCol)
    
            iPosAtual = InStr(1, sRegistro, Chr(vbKeyLButton)) + 1
            iPosFimAtual = InStr(1, sRegistro, Chr(vbKeyLButton)) + 1
    
    
        Loop

    End If
    
    'pula o vbKeySeparator
    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = iPosFimAtual + 1

    Vendas_Cheque = SUCESSO

    Exit Function

Erro_Vendas_Cheque:

    Vendas_Cheque = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149412)

    End Select

    Exit Function

End Function

Function Vendas_Troca(iPosAtual As Integer, iPosFimAtual As Integer, ByVal colRegistro As Collection, iRegistroCol As Integer, ByVal colTroca As Collection) As Long
'guarda as infos de troca se houverem

Dim lErro As Long
Dim objTroca As ClassTroca
Dim iIndice As Integer
Dim sRegistro As String

On Error GoTo Erro_Vendas_Troca
    
        
    sRegistro = colRegistro.Item(iRegistroCol)

    'Se o proximo caracter for o vbKeyControl ==> tem movimentos de caixa envolvidos
    If Mid(sRegistro, iPosAtual, 1) = Chr(vbKeyControl) Then

        Do While iRegistroCol <= colRegistro.Count
        
            Do While Mid(sRegistro, iPosAtual, 1) = Chr(vbKeyControl)
    
                Set objTroca = New ClassTroca
    
                colTroca.Add objTroca
                            
                iPosAtual = iPosFimAtual + 1
                iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                objTroca.dQuantidade = CDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
    
                iPosAtual = iPosFimAtual + 1
                iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                objTroca.dValor = CDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
    
                iPosAtual = iPosFimAtual + 1
                iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                objTroca.sProduto = CStr(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
    
                iPosAtual = iPosFimAtual + 1
                iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                objTroca.sCodProduto = CStr(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
    
                iPosAtual = iPosFimAtual + 1
                iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
                objTroca.sUnidadeMed = CStr(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
    
                iPosAtual = iPosFimAtual + 1
                iPosFimAtual = iPosFimAtual + 1
            Loop

            If iPosAtual <= Len(sRegistro) Then Exit Do
    
            iRegistroCol = iRegistroCol + 1
            sRegistro = colRegistro.Item(iRegistroCol)
        
            iPosAtual = InStr(1, sRegistro, Chr(vbKeyLButton)) + 1
            iPosFimAtual = InStr(1, sRegistro, Chr(vbKeyLButton)) + 1

        Loop

    End If

    Vendas_Troca = SUCESSO

    Exit Function

Erro_Vendas_Troca:

    Vendas_Troca = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149413)

    End Select

    Exit Function

End Function

Function MovimentoDinheiro_Atualiza_Memoria(objMovimentosCaixa As ClassMovimentoCaixa) As Long
'Função que Atualiza a Memoria

Dim lErro As Long
Dim iIndice As Integer
Dim dValor As Double

On Error GoTo Erro_MovimentoDinheiro_Atualiza_Memoria
    
    'para não ficar 3 movimentos com o mesmo Código(Numero de Movto) na Coleção gcolMovto
    If objMovimentosCaixa.iTipo <> MOVIMENTOCAIXA_EXCLUSAO_SANGRIA_DINHEIRO And objMovimentosCaixa.iTipo <> MOVIMENTOCAIXA_EXCLUSAO_SUPRIMENTO_DINHEIRO Then

        'Adcionar a Coleção Global o objMovimento Caixa
        gcolMovimentosCaixa.Add objMovimentosCaixa
        
        'Atualiza a Variável Global
        
        If objMovimentosCaixa.iTipo = MOVIMENTOCAIXA_SANGRIA_DINHEIRO Then

            '??? 24/08/2016 gdSaldoDinheiro = gdSaldoDinheiro - objMovimentosCaixa.dValor
            
        ElseIf objMovimentosCaixa.iTipo = MOVIMENTOCAIXA_SUPRIMENTO_DINHEIRO Or objMovimentosCaixa.iTipo = MOVIMENTOCAIXA_TRANSF_SUP_DINHEIRO Then

            '??? 24/08/2016 gdSaldoDinheiro = gdSaldoDinheiro + objMovimentosCaixa.dValor

        End If

    Else
    
        'Função que Retira de memória os Movimentos Excluidos
        lErro = MovimentoCaixa_Exclui_Memoria(objMovimentosCaixa)
        If lErro <> SUCESSO Then gError 109792
        
        If objMovimentosCaixa.iTipo = MOVIMENTOCAIXA_EXCLUSAO_SANGRIA_DINHEIRO Then
            dValor = objMovimentosCaixa.dValor
        Else
            dValor = (-1) * objMovimentosCaixa.dValor
        End If
        
        'Atualiza o Saldo Global
        '??? 24/08/2016 gdSaldoDinheiro = gdSaldoDinheiro + dValor
        
    End If

    MovimentoDinheiro_Atualiza_Memoria = SUCESSO

    Exit Function

Erro_MovimentoDinheiro_Atualiza_Memoria:

    MovimentoDinheiro_Atualiza_Memoria = gErr

    Select Case gErr

        Case 109792

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 149414)

    End Select

    Exit Function

End Function

Function MovimentoOutros_Atualiza_Memoria(objMovimentosCaixa As ClassMovimentoCaixa) As Long
'Função que Limpa a Coleção Global a Tela apos a Função de Gravação

Dim lErro As Long
Dim iIndice As Integer
Dim objAdmMeioPagtoCondPagto As New ClassAdmMeioPagtoCondPagto

On Error GoTo Erro_MovimentoOutros_Atualiza_Memoria

    If objMovimentosCaixa.iTipo = MOVIMENTOCAIXA_TRANSF_SUP_OUTROS Then

        'Adcionar a Coleção Global o objMovimento Caixa
        gcolMovimentosCaixa.Add objMovimentosCaixa

        'Atualiza o Saldo na Coleção Global após a Sangria(Movimento SangriaBoleto)
        For Each objAdmMeioPagtoCondPagto In gcolOutros
            'Verifica o Nome do meio Pagto
            If objMovimentosCaixa.iAdmMeioPagto = objAdmMeioPagtoCondPagto.iAdmMeioPagto Then
                objAdmMeioPagtoCondPagto.dSaldo = objAdmMeioPagtoCondPagto.dSaldo + objMovimentosCaixa.dValor
                Exit For
            End If
        Next

    ElseIf objMovimentosCaixa.iTipo = MOVIMENTOCAIXA_EXCLUSAO_SANGRIA_OUTROS Then

        'Função que Retira de memória os Movimentos Excluidos
        lErro = MovimentoCaixa_Exclui_Memoria(objMovimentosCaixa)
        If lErro <> SUCESSO Then gError 109798

        'Função que Atualiza os meios de Pagto Excluidos
        lErro = CF_ECF("MovimentoOutros_Atualiza_Memoria1", objMovimentosCaixa)
        If lErro <> SUCESSO Then gError 105720

    Else
    
        'Adcionar a Coleção Global o objMovimento Caixa
        gcolMovimentosCaixa.Add objMovimentosCaixa

        'Função que Atualiza os meios de Pagto Excluidos
        lErro = CF_ECF("MovimentoOutros_Atualiza_Memoria1", objMovimentosCaixa)
        If lErro <> SUCESSO Then gError 105721
        
    End If
    
    MovimentoOutros_Atualiza_Memoria = SUCESSO

    Exit Function

Erro_MovimentoOutros_Atualiza_Memoria:

    MovimentoOutros_Atualiza_Memoria = gErr

    Select Case gErr

        Case 105720, 105721, 109798

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 149415)

    End Select

    Exit Function

End Function

Function MovimentoTicket_Atualiza_Memoria(objMovimentosCaixa As ClassMovimentoCaixa) As Long
'Função que Limpa a Coleção Global a Tela apos a Função de Gravação

Dim lErro As Long
Dim iIndice As Integer
Dim objAdmMeioPagtoCondPagto As New ClassAdmMeioPagtoCondPagto

On Error GoTo Erro_MovimentoTicket_Atualiza_Memoria

    If objMovimentosCaixa.iTipo = MOVIMENTOCAIXA_TRANSF_SUP_TICKET Then

        'Adcionar a Coleção Global o objMovimento Caixa
        gcolMovimentosCaixa.Add objMovimentosCaixa

        For Each objAdmMeioPagtoCondPagto In gcolTicket
            'Verifica o Nome da Admnistradora
            If objMovimentosCaixa.iAdmMeioPagto = objAdmMeioPagtoCondPagto.iAdmMeioPagto Then
                objAdmMeioPagtoCondPagto.dSaldo = objAdmMeioPagtoCondPagto.dSaldo + objMovimentosCaixa.dValor
                Exit For
            End If
        Next

    ElseIf objMovimentosCaixa.iTipo = MOVIMENTOCAIXA_EXCLUSAO_SANGRIA_TICKET Then

        'Função que Retira de memória os Movimentos Excluidos
        lErro = MovimentoCaixa_Exclui_Memoria(objMovimentosCaixa)
        If lErro <> SUCESSO Then gError 109800

        'Função que Atualiza os Boletos Excluidos
        lErro = CF_ECF("MovimentoTicket_Atualiza_Memoria1", objMovimentosCaixa)
        If lErro <> SUCESSO Then gError 105718

    Else
        'Adcionar a Coleção Global o objMovimento Caixa
        gcolMovimentosCaixa.Add objMovimentosCaixa

        'Função que Atualiza os Boletos Excluidos
        lErro = CF_ECF("MovimentoTicket_Atualiza_Memoria1", objMovimentosCaixa)
        If lErro <> SUCESSO Then gError 105719

    End If
    
    MovimentoTicket_Atualiza_Memoria = SUCESSO

    Exit Function

Erro_MovimentoTicket_Atualiza_Memoria:

    MovimentoTicket_Atualiza_Memoria = gErr

    Select Case gErr

        Case 105718, 105719, 109800

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 149416)

    End Select

    Exit Function

End Function

Function MovimentoBoleto_Atualiza_Memoria(objMovimentoCaixa As ClassMovimentoCaixa) As Long
'Função que Limpa a Coleção Global a Tela apos a Função de Gravação

Dim lErro As Long
Dim iIndice As Integer
Dim objAdmMeioPagtoCondPagto As ClassAdmMeioPagtoCondPagto
Dim sNomeTerminal As String

On Error GoTo Erro_MovimentoBoleto_Atualiza_Memoria
    
    'para não ficar 3 movimentos com o mesmo Código(Numero de Movto) na Coleção gcolMovto
    If objMovimentoCaixa.iTipo = MOVIMENTOCAIXA_TRANSF_SUP_BOLETO_CC Then

        'Adcionar a Coleção Global o objMovimento Caixa
        gcolMovimentosCaixa.Add objMovimentoCaixa

        'Atualiza o Saldo na Coleção Global após a Sangria(Movimento SangriaBoleto)
        For Each objAdmMeioPagtoCondPagto In gcolCartao
            'Verifica o Nome da Admnistradora
            If objMovimentoCaixa.iAdmMeioPagto = objAdmMeioPagtoCondPagto.iAdmMeioPagto And ((objMovimentoCaixa.iParcelamento = objAdmMeioPagtoCondPagto.iParcelamento And objAdmMeioPagtoCondPagto.iTipoCartao = objMovimentoCaixa.iTipoCartao And objAdmMeioPagtoCondPagto.iAdmMeioPagto <> 0 And objAdmMeioPagtoCondPagto.dtDataPreDatado = objMovimentoCaixa.dtDataPreDatado) Or (objAdmMeioPagtoCondPagto.iAdmMeioPagto = 0 And objAdmMeioPagtoCondPagto.sNomeAdmMeioPagto = STRING_NAO_DETALHADO_CCREDITO)) Then
                objAdmMeioPagtoCondPagto.dSaldo = objAdmMeioPagtoCondPagto.dSaldo + StrParaDbl(objMovimentoCaixa.dValor)
                Exit For
            End If
        Next
    
    'para não ficar 3 movimentos com o mesmo Código(Numero de Movto) na Coleção gcolMovto
    ElseIf objMovimentoCaixa.iTipo <> MOVIMENTOCAIXA_EXCLUSAO_SANGRIA_B0LETO Then

        'Adcionar a Coleção Global o objMovimento Caixa
        gcolMovimentosCaixa.Add objMovimentoCaixa

        'Função que Atualiza os Boletos Excluidos
        lErro = CF_ECF("MovimentoBoleto_Atualiza_Memoria1", objMovimentoCaixa)
        If lErro <> SUCESSO Then gError 105716
        
    Else

        'Função que Retira de memória os Movimentos Excluidos
        lErro = MovimentoCaixa_Exclui_Memoria(objMovimentoCaixa)
        If lErro <> SUCESSO Then gError 109783

        'Função que Atualiza os Boletos Excluidos
        lErro = CF_ECF("MovimentoBoleto_Atualiza_Memoria1", objMovimentoCaixa)
        If lErro <> SUCESSO Then gError 105717
     
    End If
    
    MovimentoBoleto_Atualiza_Memoria = SUCESSO

    Exit Function

Erro_MovimentoBoleto_Atualiza_Memoria:

    MovimentoBoleto_Atualiza_Memoria = gErr

    Select Case gErr

        Case 105716, 105717, 109783

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 149417)

    End Select

    Exit Function

End Function

Function MovimentoCheque_Atualiza_Memoria(objMovimentosCaixa As ClassMovimentoCaixa, colImfCompl As Collection) As Long
'Função que Limpa a Coleção Global a Tela apos a Função de Gravação

Dim lErro As Long
Dim iIndice As Integer
Dim objCheque As New ClassChequePre
Dim iCont As Integer
Dim objChequeAux As New ClassChequePre
Dim dValorMovimento As Double
Dim lSeq As Long

On Error GoTo Erro_MovimentoCheque_Atualiza_Memoria
    
    If objMovimentosCaixa.iTipo = MOVIMENTOCAIXA_SANGRIA_CHEQUE Then
    
        'Adcionar a Coleção Global o objMovimento Caixa
        gcolMovimentosCaixa.Add objMovimentosCaixa
        
        gdSaldocheques = gdSaldocheques - objMovimentosCaixa.dValor
                
        For iIndice = 1 To colImfCompl.Count
        
            lSeq = colImfCompl.Item(iIndice)
            'Atualiza a Coleção de Forma a Voltar para aColeção global de Cheques
            For Each objCheque In gcolCheque
                'coloca o numero do movimento de sangria no cheque
                If objCheque.lSequencialCaixa = lSeq Then objCheque.lNumMovtoSangria = objMovimentosCaixa.lNumMovto
            Next
        Next
        
    ElseIf objMovimentosCaixa.iTipo = MOVIMENTOCAIXA_TRANSF_SUP_CHEQUE Then
    
        'Adcionar a Coleção Global o objMovimento Caixa
        gcolMovimentosCaixa.Add objMovimentosCaixa
        
    Else
    
        'Função que Retira de memória os Movimentos Excluidos
        lErro = MovimentoCaixa_Exclui_Memoria(objMovimentosCaixa)
        If lErro <> SUCESSO Then gError 107886

        gdSaldocheques = gdSaldocheques + objMovimentosCaixa.dValor

        For iIndice = 1 To colImfCompl.Count
        
            lSeq = colImfCompl.Item(iIndice)
            'Atualiza a Coleção de Forma a Voltar para aColeção global de Cheques
            For Each objCheque In gcolCheque
                'Verifica se o numero do movimento de sangria é igual ao numero do movimento dacoleção global de Cheque --> 'Zera o Numero da Sangria
                If objCheque.lSequencialCaixa = lSeq Then objCheque.lNumMovtoSangria = 0
            Next
        Next
    
    End If
    
    MovimentoCheque_Atualiza_Memoria = SUCESSO

    Exit Function

Erro_MovimentoCheque_Atualiza_Memoria:

    MovimentoCheque_Atualiza_Memoria = gErr

    Select Case gErr

        Case 107886, 107887

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 149418)

    End Select

    Exit Function

End Function

Function MovimentoCaixa_Exclui_Memoria(objMovimentoCaixa As ClassMovimentoCaixa) As Long
'Função que Exclui da Memória os Movimentos que Foram Alterados

Dim lErro As Long
Dim objMovimentoCaixaAux As New ClassMovimentoCaixa
Dim iIndice As Integer

On Error GoTo Erro_MovimentoCaixa_Exclui_Memoria

    For iIndice = gcolMovimentosCaixa.Count To 1 Step -1

        Set objMovimentoCaixaAux = gcolMovimentosCaixa.Item(iIndice)

        'Verifica se o movimento é o mesmo
        If objMovimentoCaixa.lNumMovto = objMovimentoCaixaAux.lNumMovto Then

            'Exclui o movimento da Coleção Global de MovimentosCaixa
            gcolMovimentosCaixa.Remove (iIndice)

            'Sai do Loop
            Exit For

        End If

    Next

    MovimentoCaixa_Exclui_Memoria = SUCESSO

    Exit Function

Erro_MovimentoCaixa_Exclui_Memoria:

    MovimentoCaixa_Exclui_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 149419)

    End Select

    Exit Function

End Function

Function MovimentoCaixa_GravaMemoria(objMovimentosCaixa As ClassMovimentoCaixa) As Long
        
Dim sStatusSessao As String
Dim sStatusCaixa As String

On Error GoTo Erro_MovimentoCaixa_GravaMemoria

'    Select Case objMovimentosCaixa.iTipo
'
'        Case MOVIMENTO_CAIXA_ABERTURA
'
'            'Altera o Status do Caixa, Instancia a Varialvel Global
'            giStatusCaixa = STATUS_CAIXA_ABERTO
'
'        Case MOVIMENTO_CAIXA_FECHAMENTO
'
'            'Caixa Muda o Status
'            giStatusCaixa = STATUS_CAIXA_FECHADO
'
'        Case MOVIMENTO_CAIXA_REDUCAO_Z
'            'Atualizar na Variável Global a Data da Ultima Redução Z
'            gdtUltimaReducao = objMovimentosCaixa.dtDataMovimento
'
'        Case MOVIMENTO_CAIXA_SESSAO_ABERTURA
'
'            giStatusSessao = SESSAO_ABERTA
'
'        Case MOVIMENTO_CAIXA_SESSAO_FECHAMENTO
'
'            'Indica o Novo status da Sessão Suspensa
'            giStatusSessao = SESSAO_ENCERRADA
'
'        Case MOVIMENTO_CAIXA_SESSAO_SUSPENSAO
'
'            'Indica o Novo status da Sessão Suspensa
'            giStatusSessao = SESSAO_SUSPENSA
'
'    End Select
        
    gcolMovimentosCaixa.Add objMovimentosCaixa
    
    MovimentoCaixa_GravaMemoria = SUCESSO
        
    Exit Function

Erro_MovimentoCaixa_GravaMemoria:
    
    MovimentoCaixa_GravaMemoria = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149420)

    End Select
    
    Exit Function
    
End Function

Function Registro_ECF_TransfCaixa(ByVal colRegistro As Collection, ByVal sArquivo As String, objTransfCaixa As ClassTransfCaixa, objChequePara As ClassChequePre) As Long
'guarda as infos de movimento de caixa

Dim lErro As Long
Dim sRegistro As String
Dim iPosAtual As Integer
Dim iPosFimAtual As Integer

On Error GoTo Erro_Registro_ECF_TransfCaixa
 
    sRegistro = colRegistro.Item(1)

    iPosFimAtual = InStr(sRegistro, Chr(vbKeyControl))
 
    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objTransfCaixa.objMovCaixaDe.iTipo = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objTransfCaixa.objMovCaixaDe.lSequencial = CLng(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objTransfCaixa.objMovCaixaDe.iAdmMeioPagto = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objTransfCaixa.objMovCaixaDe.iParcelamento = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objTransfCaixa.objMovCaixaDe.dHora = CDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objTransfCaixa.objMovCaixaDe.iTipoCartao = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objTransfCaixa.objMovCaixaDe.lCupomFiscal = CLng(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objTransfCaixa.objMovCaixaDe.lNumRefInterna = CLng(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objTransfCaixa.objMovCaixaPara.iTipo = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objTransfCaixa.objMovCaixaPara.lSequencial = CLng(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objTransfCaixa.objMovCaixaPara.iAdmMeioPagto = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objTransfCaixa.objMovCaixaPara.iParcelamento = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objTransfCaixa.objMovCaixaPara.dHora = CDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objTransfCaixa.objMovCaixaPara.iTipoCartao = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objTransfCaixa.objMovCaixaPara.lCupomFiscal = CLng(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objTransfCaixa.objMovCaixaPara.lNumRefInterna = CLng(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objTransfCaixa.objMovCaixaDe.iCaixa = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objTransfCaixa.objMovCaixaDe.iFilialEmpresa = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objTransfCaixa.objMovCaixaDe.iCodOperador = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objTransfCaixa.objMovCaixaDe.dtDataMovimento = CDate(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objTransfCaixa.objMovCaixaDe.dValor = CDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objTransfCaixa.objMovCaixaDe.lTransferencia = CLng(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    objTransfCaixa.objMovCaixaPara.iCaixa = objTransfCaixa.objMovCaixaDe.iCaixa
    objTransfCaixa.objMovCaixaPara.iFilialEmpresa = objTransfCaixa.objMovCaixaDe.iFilialEmpresa
    objTransfCaixa.objMovCaixaPara.iCodOperador = objTransfCaixa.objMovCaixaDe.iCodOperador
    objTransfCaixa.objMovCaixaPara.dtDataMovimento = objTransfCaixa.objMovCaixaDe.dtDataMovimento
    objTransfCaixa.objMovCaixaPara.dValor = objTransfCaixa.objMovCaixaDe.dValor
    objTransfCaixa.objMovCaixaPara.lTransferencia = objTransfCaixa.objMovCaixaDe.lTransferencia

    'se entrou um cheque
    If objTransfCaixa.objMovCaixaPara.iTipo = MOVIMENTOCAIXA_ENTRADA_TRANSF_CHEQUE Then

        iPosAtual = iPosFimAtual + 1
        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
        objChequePara.dtDataDeposito = CDate(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
    
        iPosAtual = iPosFimAtual + 1
        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
        objChequePara.iBanco = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
    
        iPosAtual = iPosFimAtual + 1
        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
        objChequePara.iNaoEspecificado = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
    
        iPosAtual = iPosFimAtual + 1
        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
        objChequePara.iStatus = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
    
        iPosAtual = iPosFimAtual + 1
        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
        objChequePara.lNumero = CLng(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
    
        iPosAtual = iPosFimAtual + 1
        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
        objChequePara.lSequencialCaixa = CLng(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

        iPosAtual = iPosFimAtual + 1
        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
        objChequePara.sAgencia = Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual)
    
        iPosAtual = iPosFimAtual + 1
        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
        objChequePara.sContaCorrente = Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual)
    
        iPosAtual = iPosFimAtual + 1
        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
        objChequePara.sCPFCGC = Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual)
        
        iPosAtual = iPosFimAtual + 1
        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
        objChequePara.iLocalizacao = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

        iPosAtual = iPosFimAtual + 1
        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
        objChequePara.iECF = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
        
        objChequePara.dValor = objTransfCaixa.objMovCaixaPara.dValor
        objChequePara.lCupomFiscal = objTransfCaixa.objMovCaixaPara.lCupomFiscal

    End If
    
    Registro_ECF_TransfCaixa = SUCESSO

    Exit Function

Erro_Registro_ECF_TransfCaixa:

    Registro_ECF_TransfCaixa = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149421)

    End Select

    Exit Function

End Function

Function Registro_ECF_TransfCaixa_Exclusao(ByVal colRegistro As Collection, ByVal sArquivo As String, objTransfCaixa As ClassTransfCaixa) As Long
'guarda as infos de movimento de caixa

Dim lErro As Long
Dim sRegistro As String
Dim iPosAtual As Integer
Dim iPosFimAtual As Integer

On Error GoTo Erro_Registro_ECF_TransfCaixa_Exclusao
 
    sRegistro = colRegistro.Item(1)

    iPosFimAtual = InStr(sRegistro, Chr(vbKeyControl))
 
    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objTransfCaixa.iFilialEmpresa = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objTransfCaixa.iCaixa = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
 
    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objTransfCaixa.objMovCaixaDe.lTransferencia = CLng(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
    
    objTransfCaixa.objMovCaixaPara.lTransferencia = objTransfCaixa.objMovCaixaDe.lTransferencia

    Registro_ECF_TransfCaixa_Exclusao = SUCESSO

    Exit Function

Erro_Registro_ECF_TransfCaixa_Exclusao:

    Registro_ECF_TransfCaixa_Exclusao = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149422)

    End Select

    Exit Function

End Function

Function Movcx_Movcx(ByVal colRegistro As Collection, ByVal objMovCx As ClassMovimentoCaixa) As Long
'guarda as infos de movimento de caixa relativos a operacoes de caixa (abertura/fechamento/etc)

Dim lErro As Long
Dim iPosAtual As Integer
Dim iPosFimAtual As Integer
Dim sRegistro As String


On Error GoTo Erro_Movcx_Movcx

    sRegistro = colRegistro(1)

    iPosFimAtual = InStr(sRegistro, Chr(vbKeyControl))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.iTipo = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.dHora = CDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.dtDataMovimento = CDate(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.iGerente = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.iCodOperador = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.lSequencial = CLng(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.iFilialEmpresa = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.iCaixa = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    Movcx_Movcx = SUCESSO

    Exit Function

Erro_Movcx_Movcx:

    Movcx_Movcx = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149423)

    End Select

    Exit Function

End Function

Function MovcxBol_Movcx(ByVal colRegistro As Collection, ByVal objMovCx As ClassMovimentoCaixa) As Long
'guarda as infos de movimento de caixa relativos a sangria de boleto

Dim lErro As Long
Dim iPosAtual As Integer
Dim iPosFimAtual As Integer
Dim sRegistro As String

On Error GoTo Erro_MovcxBol_Movcx

    sRegistro = colRegistro(1)
   
    iPosFimAtual = InStr(sRegistro, Chr(vbKeyControl))
   
    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.iTipo = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.dHora = CDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.dtDataMovimento = CDate(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.iGerente = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.iCodOperador = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.lSequencial = CLng(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.iFilialEmpresa = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.dValor = CDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.iAdmMeioPagto = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.iParcelamento = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.iCaixa = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.iTipoCartao = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
    
    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.lNumMovto = CLng(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.lNumRefInterna = CLng(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    If iPosFimAtual <> 0 Then
        objMovCx.dtDataMovimento = CDate(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
    End If
    
    MovcxBol_Movcx = SUCESSO

    Exit Function

Erro_MovcxBol_Movcx:

    MovcxBol_Movcx = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149424)

    End Select

    Exit Function

End Function

Function MovcxChq_Cheque(iPosAtual As Integer, iPosFimAtual As Integer, ByVal colRegistro As Collection, ByVal colCheque As Collection) As Long
'guarda as infos de troca se houverem

Dim lErro As Long
Dim vRegistro As Variant

On Error GoTo Erro_MovcxChq_Cheque

    For Each vRegistro In colRegistro

        Do While Mid(vRegistro, iPosFimAtual + 1, 1) <> Chr(vbKeyEnd) And iPosFimAtual + 1 <= Len(vRegistro)

            iPosAtual = iPosFimAtual + 1
            iPosFimAtual = InStr(iPosAtual, vRegistro, Chr(vbKeyShift))
            colCheque.Add CLng(Mid(vRegistro, iPosAtual, iPosFimAtual - iPosAtual))

        Loop

        iPosFimAtual = 0

    Next

    MovcxChq_Cheque = SUCESSO

    Exit Function

Erro_MovcxChq_Cheque:

    MovcxChq_Cheque = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149425)

    End Select

    Exit Function

End Function

Function MovcxChq_Movcx(iPosAtual As Integer, iPosFimAtual As Integer, ByVal sRegistro As String, ByVal objMovCx As ClassMovimentoCaixa) As Long
'guarda as infos de movimento de caixa relativos a sangria de cheque

Dim lErro As Long

On Error GoTo Erro_MovcxChq_Movcx

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.iTipo = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.dHora = CDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.dtDataMovimento = CDate(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.iGerente = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.iCodOperador = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.lSequencial = CLng(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.iFilialEmpresa = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.dValor = CDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.iAdmMeioPagto = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.iParcelamento = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.iCaixa = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.lNumMovto = CLng(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    MovcxChq_Movcx = SUCESSO

    Exit Function

Erro_MovcxChq_Movcx:

    MovcxChq_Movcx = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149426)

    End Select

    Exit Function

End Function

Function MovcxDin_Movcx(ByVal sRegistro As String, ByVal objMovCx As ClassMovimentoCaixa) As Long
'guarda as infos de movimento de caixa

Dim lErro As Long
Dim iPosAtual As Integer
Dim iPosFimAtual As Integer

On Error GoTo Erro_MovcxDin_Movcx
 
    iPosFimAtual = InStr(sRegistro, Chr(vbKeyControl))
 
    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.iTipo = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.dHora = CDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.dtDataMovimento = CDate(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.iGerente = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.iCodOperador = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.lSequencial = CLng(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.iFilialEmpresa = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.dValor = CDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.iAdmMeioPagto = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.iParcelamento = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.iCaixa = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
        
    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.lNumMovto = CLng(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
    
    MovcxDin_Movcx = SUCESSO

    Exit Function

Erro_MovcxDin_Movcx:

    MovcxDin_Movcx = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149427)

    End Select

    Exit Function

End Function

Function MovcxOut_Movcx(ByVal colRegistro As Collection, ByVal objMovCx As ClassMovimentoCaixa) As Long
'guarda as infos de movimento de caixa relativos a sangria outros

Dim lErro As Long
Dim sRegistro As String
Dim iPosAtual As Integer
Dim iPosFimAtual As Integer

On Error GoTo Erro_MovcxOut_Movcx

    sRegistro = colRegistro(1)

    iPosFimAtual = InStr(sRegistro, Chr(vbKeyControl))


    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.iTipo = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.dHora = CDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.dtDataMovimento = CDate(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.iGerente = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.iCodOperador = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.lSequencial = CLng(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.iFilialEmpresa = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.dValor = CDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.iAdmMeioPagto = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.iParcelamento = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.iCaixa = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.lNumMovto = CLng(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    MovcxOut_Movcx = SUCESSO

    Exit Function

Erro_MovcxOut_Movcx:

    MovcxOut_Movcx = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149428)

    End Select

    Exit Function

End Function

Function MovcxTkt_Movcx(ByVal colRegistro As Collection, ByVal objMovCx As ClassMovimentoCaixa) As Long
'guarda as infos de movimento de caixa relativos a sangria de ticket

Dim lErro As Long
Dim iPosAtual As Integer
Dim iPosFimAtual As Integer
Dim sRegistro As String

On Error GoTo Erro_MovcxTkt_Movcx

    sRegistro = colRegistro(1)

    iPosFimAtual = InStr(sRegistro, Chr(vbKeyControl))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.iTipo = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.dHora = CDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.dtDataMovimento = CDate(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.iGerente = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.iCodOperador = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.lSequencial = CLng(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.iFilialEmpresa = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.dValor = CDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.iAdmMeioPagto = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.iParcelamento = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.iCaixa = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.lNumMovto = CLng(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    MovcxTkt_Movcx = SUCESSO

    Exit Function

Erro_MovcxTkt_Movcx:

    MovcxTkt_Movcx = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149429)

    End Select

    Exit Function

End Function

'Function ArqSeq_Le(objArqSeq As ClassArqSeq) As Long
''le o arquivo ArqSeqOr
'
'Dim sArqSeq As String
'Dim sRetorno As String
'Dim lTamanho As Long
'Dim tArqSeq As typeArqSeq
'Dim sDir As String
'Dim lTentaAcessoArquivo As Long
'Dim lErro As Long
'
'On Error GoTo Erro_ArqSeq_Le
'
'    lTamanho = 255
'    sRetorno = String(lTamanho, 0)
'
'    'funcao que verifica se o caixaconfig.ini foi alterado
'    lErro = CF_ECF("CaixaConfig_Verifica_Chave")
'    If lErro <> SUCESSO Then gError 204240
'
'    Call GetPrivateProfileString(APLICACAO_ORCAMENTO, "Dir", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
'    If sRetorno = String(lTamanho, 0) Then gError 105847
'
'    sDir = StringZ(sRetorno)
'
'    'se o diretorio nao for terminado por \  ===> acrescentar
'    If right(sDir, 1) <> "\" Then sDir = sDir & "\"
'
'    sArqSeq = sDir & "ArqSeqOr.txt"
'
'    sRetorno = Dir(sArqSeq)
'
'    'se nao encontrou o arquivo que contem o sequencial de orcamento ==> criar
'    If Len(sRetorno) = 0 Then
'
'        lTentaAcessoArquivo = 0
'        Open sArqSeq For Random Lock Read Write As 1 Len = Len(tArqSeq)
'        tArqSeq.sSequencial = "0"
'        tArqSeq.sData = Format(DATA_NULA, "dd/mm/yyyy")
'        Put #1, 1, tArqSeq
'        Close #1
'
'    End If
'
'    lTentaAcessoArquivo = 0
'    Open sArqSeq For Random As 1 Len = Len(tArqSeq)
'    Get #1, 1, tArqSeq
'    Close #1
'
'    objArqSeq.dtData = StrParaDate(tArqSeq.sData)
'    objArqSeq.lSequencial = StrParaLong(tArqSeq.sSequencial)
'
'    ArqSeq_Le = SUCESSO
'
'    Exit Function
'
'Erro_ArqSeq_Le:
'
'    ArqSeq_Le = gErr
'
'    Close #1
'
'    Select Case gErr
'
'        Case 70
'            lTentaAcessoArquivo = lTentaAcessoArquivo + 1
'            If lTentaAcessoArquivo < NUM_MAX_LOCK Then Resume
'            Call Rotina_ErroECF(vbOKOnly, ERRO_ARQUIVO_LOCADO, gErr, sArqSeq)
'
'        Case 105847
'            Call Rotina_ErroECF(vbOKOnly, ERRO_PREENCHIMENTO_ARQUIVO_CONFIG, gErr, "Dir", APLICACAO_ORCAMENTO, NOME_ARQUIVO_CAIXA)
'
'        Case 204240
'
'        Case Else
'            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149430)
'
'    End Select
'
'    Exit Function
'
'End Function

'Function OrcamentoECF_Le_Primeiro(objArqSeq As ClassArqSeq, sRegistro As String) As Long
''Le o primeiro registro do arquivo de orcamento
'
'Dim sNomeArq As String
'Dim sNomeArq0 As String
'Dim sNomeArq1 As String
'Dim sRet As String
'Dim sDir As String
'Dim sRetorno As String
'Dim lTamanho As Long
'Dim lTentaAcessoArquivo As Long
'Dim lErro As Long
'
'On Error GoTo Erro_OrcamentoECF_Le_Primeiro
'
'    lTamanho = 255
'    sRetorno = String(lTamanho, 0)
'
'    'funcao que verifica se o caixaconfig.ini foi alterado
'    lErro = CF_ECF("CaixaConfig_Verifica_Chave")
'    If lErro <> SUCESSO Then gError 204241
'
'    Call GetPrivateProfileString(APLICACAO_ORCAMENTO, "Dir", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
'    If sRetorno = String(lTamanho, 0) Then gError 105848
'
'    sDir = StringZ(sRetorno)
'
'    'se o diretorio nao for terminado por \  ===> acrescentar
'    If right(sDir, 1) <> "\" Then sDir = sDir & "\"
'
'    sNomeArq1 = "OR" & Format(objArqSeq.dtData, "ddmmyy") & (".txt")
'
'    sNomeArq = sDir & sNomeArq1
'
'    'esta atribuicao é feita pois caso haja algum erro será mostrado o nome contido em sNomeArq0
'    sNomeArq0 = sNomeArq
'
'    'Verifica se tem orcamento para a data de hoje
'    sRet = Dir(sNomeArq)
'
'    If Len(sRet) = 0 Then
'
'        lTentaAcessoArquivo = 0
'        Open sNomeArq For Append Lock Read Write As #2
'
'        'Inserido no Arquivo o nome do arquivo
'        Print #2, sNomeArq1
'
'        Close #2
'
'    End If
'
'    lTentaAcessoArquivo = 0
'    Open sNomeArq For Input As #2
'
'    'Busca o próximo registro do arquivo
'    Line Input #2, sRegistro
'
'    'se nao tiver o nome do arquivo no primeiro registro ==> nao é o arquivo de orcamento que deveria ser lido
'    If sRegistro <> sNomeArq1 Then gError 105840
'
'    OrcamentoECF_Le_Primeiro = SUCESSO
'
'    Exit Function
'
'Erro_OrcamentoECF_Le_Primeiro:
'
'    OrcamentoECF_Le_Primeiro = gErr
'
'    Close #2
'
'    Select Case gErr
'
'        Case 70
'            lTentaAcessoArquivo = lTentaAcessoArquivo + 1
'            If lTentaAcessoArquivo < NUM_MAX_LOCK Then Resume
'            Call Rotina_ErroECF(vbOKOnly, ERRO_ARQUIVO_LOCADO, gErr, sNomeArq)
'
'        Case 105840
'            Call Rotina_ErroECF(vbOKOnly, ERRO_ARQUIVO_ORCAMENTO_INVALIDO, gErr, sNomeArq0)
'
'        Case 105848
'            Call Rotina_ErroECF(vbOKOnly, ERRO_PREENCHIMENTO_ARQUIVO_CONFIG, gErr, "Dir", APLICACAO_ORCAMENTO, NOME_ARQUIVO_CAIXA)
'
'        Case 204241
'
'        Case Else
'            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149431)
'
'    End Select
'
'    Exit Function
'
'End Function

'Function OrcamentoECF_Le_Lock_Primeiro(objArqSeq As ClassArqSeq, sRegistro As String) As Long
''Le o primeiro registro do arquivo de orcamento
'
'Dim sNomeArq As String
'Dim sNomeArq0 As String
'Dim sNomeArq1 As String
'Dim sRet As String
'Dim sDir As String
'Dim sRetorno As String
'Dim lTamanho As Long
'Dim lTentaAcessoArquivo As Long
'Dim lErro As Long
'
'On Error GoTo Erro_OrcamentoECF_Le_Lock_Primeiro
'
'    lTamanho = 255
'    sRetorno = String(lTamanho, 0)
'
'    'funcao que verifica se o caixaconfig.ini foi alterado
'    lErro = CF_ECF("CaixaConfig_Verifica_Chave")
'    If lErro <> SUCESSO Then gError 204242
'
'    Call GetPrivateProfileString(APLICACAO_ORCAMENTO, "Dir", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
'    If sRetorno = String(lTamanho, 0) Then gError 105848
'
'    sDir = StringZ(sRetorno)
'
'    'se o diretorio nao for terminado por \  ===> acrescentar
'    If right(sDir, 1) <> "\" Then sDir = sDir & "\"
'
'    sNomeArq1 = "OR" & Format(objArqSeq.dtData, "ddmmyy") & (".txt")
'
'    sNomeArq = sDir & sNomeArq1
'
'    'esta atribuicao é feita pois caso haja algum erro será mostrado o nome contido em sNomeArq0
'    sNomeArq0 = sNomeArq
'
'    'Verifica se tem orcamento para a data de hoje
'    sRet = Dir(sNomeArq)
'
'    If Len(sRet) = 0 Then gError 105875
'
'    lTentaAcessoArquivo = 0
'    Open sNomeArq For Input Lock Read Write As #2
'
'    'Busca o próximo registro do arquivo
'    Line Input #2, sRegistro
'
'    'se nao tiver o nome do arquivo no primeiro registro ==> nao é o arquivo de orcamento que deveria ser lido
'    If sRegistro <> sNomeArq1 Then gError 105840
'
'    OrcamentoECF_Le_Lock_Primeiro = SUCESSO
'
'    Exit Function
'
'Erro_OrcamentoECF_Le_Lock_Primeiro:
'
'    OrcamentoECF_Le_Lock_Primeiro = gErr
'
'    Close #2
'
'    Select Case gErr
'
'        Case 70
'            lTentaAcessoArquivo = lTentaAcessoArquivo + 1
'            If lTentaAcessoArquivo < NUM_MAX_LOCK Then Resume
'            Call Rotina_ErroECF(vbOKOnly, ERRO_ARQUIVO_LOCADO, gErr, sNomeArq)
'
'        Case 105840
'            Call Rotina_ErroECF(vbOKOnly, ERRO_ARQUIVO_ORCAMENTO_INVALIDO, gErr, sNomeArq0)
'
'        Case 105848
'            Call Rotina_ErroECF(vbOKOnly, ERRO_PREENCHIMENTO_ARQUIVO_CONFIG, gErr, "Dir", APLICACAO_ORCAMENTO, NOME_ARQUIVO_CAIXA)
'
'        Case 105875
'            Call Rotina_ErroECF(vbOKOnly, ERRO_ORCAMENTO_NAO_CADASTRADO1, gErr, sNomeArq0)
'
'        Case 204242
'
'        Case Else
'            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149432)
'
'    End Select
'
'    Exit Function
'
'End Function


'Function OrcamentoECF_Le_Proximo(sRegistro As String) As Long
''Le o proximo registro do arquivo de orcamento
'
'On Error GoTo Erro_OrcamentoECF_Le_Proximo
'
'    If EOF(2) Then gError 105843
'
'    'Busca o próximo registro do arquivo
'    Line Input #2, sRegistro
'
'    OrcamentoECF_Le_Proximo = SUCESSO
'
'    Exit Function
'
'Erro_OrcamentoECF_Le_Proximo:
'
'    OrcamentoECF_Le_Proximo = gErr
'
'    Select Case gErr
'
'        Case 105843
'
'        Case Else
'            Close #2
'            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149433)
'
'    End Select
'
'    Exit Function
'
'End Function

'Public Function OrcamentoECF_Le_Lock(ByVal colRegistroParam As Collection, ByVal colOrcamento As Collection) As Long
''Função Que le os orcamentos
'
'Dim lErro As Long
'Dim sTipo As String
'Dim iPos As Integer
'Dim iPosEsc As Integer
'Dim iIndice As Integer
'Dim iIndice1 As Integer
'Dim iIndice2 As Integer
'Dim sNomeArq As String
'Dim iPosInicio As Integer
'Dim bAchou As Boolean
'Dim bAchouAbert As Boolean
'Dim colRegistroTemp As New Collection
'Dim colRegistroTemp1 As New Collection
'Dim sSeqAbert As String
'Dim sSeqFech As String
'Dim sRegistro As String
'Dim colRegistro As New Collection
'Dim iRegistro As Integer
'Dim iNum As Integer
'Dim iSeqArq As Integer
'Dim lTamanho As Long
'Dim sRetorno As String
'Dim sNomeArq1 As String
'Dim sNomeArq2 As String
'Dim sRet As String
'Dim sArquivo As String
'Dim objVenda As New ClassVenda
'Dim sNomeArq0 As String
'Dim iStatus As Integer
'Dim vbMsg As VbMsgBoxResult
'Dim objArqSeq As New ClassArqSeq
'Dim lCountOrcamento As Long
'
'On Error GoTo Erro_OrcamentoECF_Le_Lock
'
'    bAchouAbert = False
'    bAchou = False
'
'    iPosInicio = 1
'
'    'le o arquivo ArqSeq para descobrir a ultima data que foi gravado o arquivo de orcamento
'    lErro = CF_ECF("ArqSeq_Le", objArqSeq)
'    If lErro <> SUCESSO Then gError 105841
'
'    'se o arquivo de orcamento antigo existir  ==>
'    'transfere os dados do arquivo antigo para o atual
'    If objArqSeq.dtData <> DATA_NULA Then
'
'        'le o primeiro registro do arquivo Orcamento antigo
'        lErro = CF_ECF("OrcamentoECF_Le_Lock_Primeiro", objArqSeq, sRegistro)
'        If lErro <> SUCESSO Then gError 105842
'
'        iIndice = 1
'
'        Do While True
'
'            'le o proximo registro do arquivo Orcamento antigo
'            lErro = CF_ECF("OrcamentoECF_Le_Proximo", sRegistro)
'            If lErro <> SUCESSO And lErro <> 105843 Then gError 105844
'
'            'se terminou o arquivo de orcamento
'            If lErro = 105843 Then Exit Do
'
'            If sRegistro <> "" Then
'
'                iIndice = iIndice + 1
'
'                'Procura o Primeiro Control para saber o tipo do registro
'                iPos = InStr(iPosInicio, sRegistro, Chr(vbKeyControl))
'
'                sTipo = Mid(sRegistro, iPosInicio, iPos - iPosInicio)
'
'                Select Case sTipo
'
'                    'verifica se é de abertura
'                    Case TIPOREGISTROECF_ABRESEQ
'                        'se encontrei um sequencial aberto antes do encerramento do anterior-->desprezo o anterior
'                        If bAchouAbert Then Set colRegistroTemp = New Collection
'                        'Recolhe o sequencial
'                        sSeqAbert = Mid(sRegistro, iPos + 1, Len(sRegistro) - (iPos + 1))
'                        bAchouAbert = True
'                            'Procura o end
'                        iPosEsc = InStr(iPos, sRegistro, Chr(vbKeyEnd))
'
'                        Set colRegistroTemp1 = New Collection
'
'                        colRegistroTemp1.Add sRegistro
'
'                    'verifica se é de fechamento
'                    Case TIPOREGISTROECF_ENCERRASEQ
'                        'Recolhe o sequencial
'                        sSeqFech = Mid(sRegistro, iPos + 1, Len(sRegistro) - (iPos + 1))
'                        If sSeqAbert = sSeqFech Then bAchou = True
'
'                        colRegistroTemp1.Add sRegistro
'
'                    Case TIPOREGISTROECF_VENDAS
'                        colRegistroTemp1.Add sRegistro
'                        colRegistroTemp.Add sRegistro
'
'                    Case Else
'                        colRegistroTemp.Add sRegistro
'
'                End Select
'
'                If bAchou Then
'
'                    For iIndice1 = 1 To colRegistroTemp.Count
'
'                        sRegistro = colRegistroTemp.Item(iIndice1)
'
'                        iRegistro = iRegistro + 1
'
'                        colRegistro.Add sRegistro
'
'                        If Mid(sRegistro, Len(sRegistro), 1) = Chr(vbKeyEnd) Then
'
'                            lCountOrcamento = colOrcamento.Count
'
'                            lErro = CF_ECF("Registro_ECF_CC", colRegistro, sArquivo, iRegistro, iIndice, colOrcamento, colRegistroParam)
'                            If lErro <> SUCESSO Then gError 105856
'
'                            'se acrescentou registro a colOrcamento  ==> leu um novo orcamento que deve ser tambem colocado em colRegistroParam
'                            If colOrcamento.Count > lCountOrcamento Then
'                                For iIndice2 = 1 To colRegistroTemp1.Count
'                                    sRegistro = colRegistroTemp1.Item(iIndice2)
'                                    colRegistroParam.Add sRegistro
'                                Next
'                            End If
'
'                            Set colRegistro = New Collection
'                            Set colRegistroTemp1 = New Collection
'                            iRegistro = 0
'                        End If
'                    Next
'                    Set colRegistroTemp = New Collection
'                    bAchouAbert = False
'                    bAchou = False
'                End If
'            End If
'        Loop
'
'    End If
'
'    OrcamentoECF_Le_Lock = SUCESSO
'
'    Exit Function
'
'Erro_OrcamentoECF_Le_Lock:
'
'    OrcamentoECF_Le_Lock = gErr
'    lErro = gErr
'
'    Select Case lErro
'
'        Case 105841, 105842, 105844, 105856
'
'        Case Else
'            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, lErro, Error$, 149435)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'
'Function OrcamentoECF_Lock(Optional dtData As Date) As Long
'
'Dim sNomeArq As String
'Dim sNomeArq1 As String
'Dim sRet As String
'Dim sRetorno As String
'Dim lTamanho As Long
'Dim sDir As String
'Dim objArqSeq As New ClassArqSeq
'Dim lErro As Long
'Dim lTentaAcessoArquivo As Long
'
'On Error GoTo Erro_OrcamentoECF_Lock
'
'    If dtData = "00:00:00" Then dtData = gdtDataHoje
'
'    lTamanho = 255
'    sRetorno = String(lTamanho, 0)
'
'    'funcao que verifica se o caixaconfig.ini foi alterado
'    lErro = CF_ECF("CaixaConfig_Verifica_Chave")
'    If lErro <> SUCESSO Then gError 204243
'
'    Call GetPrivateProfileString(APLICACAO_ORCAMENTO, "Dir", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
'    If sRetorno = String(lTamanho, 0) Then gError 105865
'
'    sDir = StringZ(sRetorno)
'
'    'se o diretorio nao for terminado por \  ===> acrescentar
'    If right(sDir, 1) <> "\" Then sDir = sDir & "\"
'
'    sNomeArq1 = "OR" & Format(dtData, "ddmmyy") & (".txt")
'
'    sNomeArq = sDir & sNomeArq1
'
'    'Verifica se tem orcamento para a data de hoje
'    sRet = Dir(sNomeArq)
'
'    If Len(sRet) = 0 Then
'
'        lTentaAcessoArquivo = 0
'
'        Open sNomeArq For Append Lock Read Write As #3
'
'        'Inserido no Arquivo o nome do arquivo
'        Print #3, sNomeArq1
'
'        Close #3
'
'        objArqSeq.dtData = dtData
'
'        'grava a data do novo arquivo de orcamentos
'        lErro = CF_ECF("ArqSeq_Grava_Data", objArqSeq)
'        If lErro <> SUCESSO Then gError 105874
'
'    End If
'
'    lTentaAcessoArquivo = 0
'    Open sNomeArq For Append Lock Read Write As #3
'
'    OrcamentoECF_Lock = SUCESSO
'
'    Exit Function
'
'Erro_OrcamentoECF_Lock:
'
'    OrcamentoECF_Lock = gErr
'
'    Select Case gErr
'
'        Case 70
'            lTentaAcessoArquivo = lTentaAcessoArquivo + 1
'            If lTentaAcessoArquivo < NUM_MAX_LOCK Then Resume
'            Call Rotina_ErroECF(vbOKOnly, ERRO_ARQUIVO_LOCADO, gErr, sNomeArq)
'
'        Case 105865, 105874, 204243
'
'        Case Else
'            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149437)
'
'    End Select
'
'    Exit Function
'
'End Function


Function Imprime_OrcamentoECF(dtDataFinal As Date, ByVal lNumOrcamento As Long, ByVal objTela As Object, ByVal objVenda As ClassVenda) As Long
'imprime o orcamento

Dim lErro As Long
Dim colMensagem As New Collection
Dim objItens As ClassItemCupomFiscal
Dim objProduto As ClassProduto
Dim iIndice As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim bCont As Boolean
Dim sMensagem As String
Dim dTotal As Double
Dim dPreco As Double
Dim bAchou As Boolean
Dim colOrcamento As New Collection
Dim iIndice1 As Integer
Dim objCliente As ClassCliente
Dim lIndice As Long
Dim sCliente As String
Dim sCPFCGC As String
Dim sVendedor As String
Dim objVendedor As ClassVendedor
Dim sProduto As String
Dim sCNPJ As String
Dim iPos As Integer
Dim lNumDAV As Long
Dim sRetorno As String
Dim lSequencialOrc As Long
Dim lSequencial As Long

On Error GoTo Erro_Imprime_OrcamentoECF

    'Abre a Transação
    lErro = CF_ECF("Caixa_Transacao_Abrir", lSequencial)
    If lErro <> SUCESSO Then gError 210261

    'Abre a Transação
    lErro = CF_ECF("Caixa_Transacao_Abrir_Orc", lSequencialOrc)
    If lErro <> SUCESSO Then gError 204685


    If objTela.Name = "Pagamento" Or objTela.Name = "FechaOrcamento" Then
        'Se o valor é insuficiente para pagar
        If StrParaDbl(objTela.objFalta.Caption) > 0 Then gError 210889
    
        If StrParaDbl(objTela.objPago.Caption) = 0 Then gError 210890

    End If

    lErro = CF_ECF("Imprime_OrcamentoECF_EmTrans", dtDataFinal, lNumOrcamento, objTela, objVenda)
    If lErro <> SUCESSO Then gError 204739

    
    'Fecha a Transação
    lErro = CF_ECF("Caixa_Transacao_Fechar_Orc", lSequencialOrc)
    If lErro <> SUCESSO Then gError 204686
    
    'Fecha a Transação
    lErro = CF_ECF("Caixa_Transacao_Fechar", lSequencial)
    If lErro <> SUCESSO Then gError 210262
    
    Imprime_OrcamentoECF = SUCESSO
    
    Exit Function
    
Erro_Imprime_OrcamentoECF:

    Imprime_OrcamentoECF = gErr

    Select Case gErr

        Case 99803, 99894 To 99898, 105868, 133088, 204319, 204324, 204329, 204739, 210261, 210262

        Case 99893
            Call AFRAC_FecharRelatorioGerencial(objTela)

        Case 210889
            Call Rotina_ErroECF(vbOKOnly, ERRO_VALOR_INSUFICIENTE, gErr)
        
        Case 210890
            Call Rotina_ErroECF(vbOKOnly, ERRO_VALOR_JA_PAGO1, gErr)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149439)

    End Select

    Call CF_ECF("Caixa_Transacao_Rollback", glTransacaoPAFECF)
    Call CF_ECF("Caixa_Transacao_Rollback", glTransacaoOrcPAFECF)

    Exit Function

End Function

Private Function Orcamento_ImpressaoCFE1(ByVal objVenda As ClassVenda, colMensagem As Collection) As Long

Dim lErro As Long, sCNPJ As String, iItem As Integer, sCPFCGC As String, iIndice As Integer
Dim objItens As ClassItemCupomFiscal, dTotal As Double, objMovCaixa As ClassMovimentoCaixa
Dim objProduto As ClassProduto, sTexto1 As String, sTexto2 As String, iAux As Integer, dPreco As Double

On Error GoTo Erro_Orcamento_ImpressaoCFE1

    Set colMensagem = New Collection
    
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", gsNomeEmpresa)
    Call Formata_CNPJ_CPF(sCNPJ, gsCNPJ)
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "CNPJ:" & sCNPJ & " IE:" & gsInscricaoEstadual & " IM:" & gsInscricaoMunicipal)
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", gsEndereco)
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", "")
    colMensagem.Add Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", IIf(objVenda.objCupomFiscal.iStatus = STATUS_BAIXADO, "Relatório Gerencial", "Documento Auxiliar de Venda"))
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", "")
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "Codigo Descricao         Qtde Un  VlrUnit VlrTot")
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", "")

    iItem = 0
    
    For Each objItens In objVenda.objCupomFiscal.colItens
    
        If objItens.iStatus <> STATUS_CANCELADO Then
    
            iItem = iItem + 1
    
            Set objProduto = New ClassProduto
    
            lErro = CF_ECF("Produtos_Le", objItens.sProduto, objProduto)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 214855
            
            If lErro = SUCESSO Then
                objItens.sProdutoNomeRed = objProduto.sNomeReduzido
            Else
                'caso o produto não seja encontrado
                gError 99894
            End If
                
            sTexto1 = objProduto.sCodigo & " " & objProduto.sDescricao
            sTexto1 = left(sTexto1, 48)
                
            sTexto2 = Format(objItens.dQuantidade, "0.0###")
            If right(sTexto2, 2) = ",0" Then sTexto2 = left(sTexto2, Len(sTexto2) - 2)
            sTexto2 = sTexto2 & " " & objItens.sUnidadeMed
            sTexto2 = sTexto2 & " X " & Format(objItens.dPrecoUnitario, "standard")
            sTexto2 = sTexto2 & " " & Format(objItens.dPrecoUnitario * objItens.dQuantidade, "standard")
                
            iAux = Len(sTexto1 & " " & sTexto2)
            If Len(sTexto1 & " " & sTexto2) <= 48 Then
                colMensagem.Add sTexto1 & String(48 - iAux, " ") & " " & sTexto2
            Else
                colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", sTexto1)
                colMensagem.Add Formata_Campo(ALINHAMENTO_ESQUERDA, 48, " ", sTexto2)
            End If
            
            dPreco = objItens.dPrecoUnitario * objItens.dQuantidade
                        
            'se existir desconto sobre o item...
            If objItens.dPercDesc > 0 Then
                colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 15, " ", "DESC. ITEM:") & Formata_Campo(ALINHAMENTO_DIREITA, 10, " ", Format(objItens.dPercDesc, "standard") & "%") & "     =  " & Formata_Campo(ALINHAMENTO_ESQUERDA, 15, " ", Format(objItens.dPrecoUnitario * objItens.dQuantidade * (1 - (objItens.dPercDesc / 100)), "standard"))
                dPreco = dPreco - (objItens.dPrecoUnitario * objItens.dQuantidade * objItens.dPercDesc / 100)
            ElseIf objItens.dValorDesconto > 0 Then
                'colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 15, " ", "DESC. ITEM:") & Formata_Campo(ALINHAMENTO_DIREITA, 10, " ", Format(-objItens.dValorDesconto, "standard")) & "     =  " & Formata_Campo(ALINHAMENTO_ESQUERDA, 15, " ", Format((objItens.dPrecoUnitario - objItens.dValorDesconto) * objItens.dQuantidade, "standard"))
                'dPreco = dPreco - (objItens.dQuantidade * objItens.dValorDesconto)
                colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 15, " ", "DESC. ITEM:") & Formata_Campo(ALINHAMENTO_DIREITA, 10, " ", Format(-objItens.dValorDesconto, "standard")) & "     =  " & Formata_Campo(ALINHAMENTO_ESQUERDA, 15, " ", Format((objItens.dPrecoUnitario * objItens.dQuantidade - objItens.dValorDesconto), "standard"))
                dPreco = dPreco - (objItens.dValorDesconto)
            End If
            
            dTotal = dTotal + dPreco
        
        End If
        
    Next
                      
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", "")
    colMensagem.Add FormataTextoEspacosNoMeio("QTD. TOTAL DE ITENS", Format(iItem, "000"), 48)
    colMensagem.Add FormataTextoEspacosNoMeio("VALOR TOTAL R$", Format(dTotal, "standard"), 48)
    
    If (objVenda.objCupomFiscal.dValorDesconto + objVenda.objCupomFiscal.dValorDesconto1) <> 0 Then
        colMensagem.Add Formata_Campo(ALINHAMENTO_ESQUERDA, 48, " ", "DESCONTO  :  -" & Format(objVenda.objCupomFiscal.dValorDesconto + objVenda.objCupomFiscal.dValorDesconto1, "standard"))
        dTotal = dTotal - (objVenda.objCupomFiscal.dValorDesconto + objVenda.objCupomFiscal.dValorDesconto1)
    End If
    
    If objVenda.objCupomFiscal.dValorAcrescimo <> 0 Then
        colMensagem.Add Formata_Campo(ALINHAMENTO_ESQUERDA, 48, " ", "ACRESCIMO :  +" & Format(objVenda.objCupomFiscal.dValorAcrescimo, "standard"))
        dTotal = dTotal + objVenda.objCupomFiscal.dValorAcrescimo
    End If
    
    If (objVenda.objCupomFiscal.dValorDesconto + objVenda.objCupomFiscal.dValorDesconto1) <> 0 Or objVenda.objCupomFiscal.dValorAcrescimo <> 0 Then
        colMensagem.Add Formata_Campo(ALINHAMENTO_ESQUERDA, 48, " ", "TOTAL LIQ.:   " & Format(dTotal, "standard"))
    End If
    
    colMensagem.Add FormataTextoEspacosNoMeio("FORMA DE PAGAMENTO", "Valor", 48)
    
    'colocar o CodigoCFe devido nos meios de pagto
    lErro = CF_ECF("Venda_CalculaCodigoCFe", objVenda)
    If lErro <> SUCESSO And lErro <> 201999 Then gError ERRO_SEM_MENSAGEM
    
    For Each objMovCaixa In objVenda.colMovimentosCaixa
        If objMovCaixa.iTipo <> MOVIMENTOCAIXA_TROCO_DINHEIRO And objMovCaixa.iTipo <> MOVIMENTOCAIXA_TROCO_CONTRAVALE And objMovCaixa.iTipo <> MOVIMENTOCAIXA_TROCO_VALE Then
            colMensagem.Add FormataTextoEspacosNoMeio(MeioPagtoCfe_Descricao(objMovCaixa.iCodigoCFe), Format(objMovCaixa.dValor, "standard"), 48)
        Else
            colMensagem.Add FormataTextoEspacosNoMeio("Troco R$", Format(objMovCaixa.dValor, "standard"), 48)
        End If
    Next
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", "")
    
    sTexto1 = objVenda.objCupomFiscal.sNFCeMensagem
    If Len(sTexto1) <> 0 Then
        Do While Len(sTexto1) <> 0
            sTexto2 = left(sTexto1, 48)
            colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", sTexto2)
            sTexto1 = Mid(sTexto1, 49)
        Loop
        colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", "")
    End If
    
    If objVenda.objCupomFiscal.lNumeroDAV <> 0 Then
        colMensagem.Add Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "No.:" & CStr(objVenda.objCupomFiscal.lNumeroDAV) & " Emissao:" & Format(objVenda.objCupomFiscal.dtDataEmissao, "dd/mm/yyyy") & " " & Format(objVenda.objCupomFiscal.dHoraEmissao, "hh:mm:ss"))
    Else
        colMensagem.Add Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "No.Orc.:" & CStr(objVenda.objCupomFiscal.lNumOrcamento) & " Emissao:" & Format(objVenda.objCupomFiscal.dtDataEmissao, "dd/mm/yyyy") & " " & Format(objVenda.objCupomFiscal.dHoraEmissao, "hh:mm:ss"))
    End If
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", "")
    
    If Len(objVenda.objCupomFiscal.sCPFCGC) <> 0 Then
        Call Formata_CNPJ_CPF(sCPFCGC, objVenda.objCupomFiscal.sCPFCGC)
        colMensagem.Add Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "CNPJ/CPF/Id.Estrangeiro: " & sCPFCGC)
    Else
        colMensagem.Add Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "CONSUMIDOR NAO IDENTIFICADO")
    End If

    If Len(Trim(objVenda.objCupomFiscal.sNomeCliente)) <> 0 Then colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "  Nome: " & objVenda.objCupomFiscal.sNomeCliente)
    
    If Len(Trim(objVenda.objCupomFiscal.sEndEntLogradouro)) <> 0 Then
    
        colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "  End.: " & objVenda.objCupomFiscal.sEndEntLogradouro & ", " & objVenda.objCupomFiscal.sEndEntNúmero)
        colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "Compl.: " & objVenda.objCupomFiscal.sEndEntComplemento)
        colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "Bairro: " & objVenda.objCupomFiscal.sEndEntBairro)
        
    End If
    
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", "")
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "")
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "")
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "")
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "")
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "")
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "")
    
    Orcamento_ImpressaoCFE1 = SUCESSO
    
    Exit Function
    
Erro_Orcamento_ImpressaoCFE1:

    Orcamento_ImpressaoCFE1 = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149439)

    End Select

    Exit Function

End Function

Private Function Orcamento_ImpressaoCFE(ByVal objTela As Object, ByVal objVenda As ClassVenda) As Long

Dim lErro As Long
Dim colMensagem As New Collection
Dim sMensagem As String, iIndice As Integer
Dim objBemaNaoFiscal As New ClassBematechNaoFiscalImp
Dim objRel As Object, bRelAberto As Boolean

On Error GoTo Erro_Orcamento_ImpressaoCFE

    lErro = AFRAC_AbrirRelatorioGerencial2(RELGER_ORCAMENTO, objTela, objRel)
    bRelAberto = True
    
    lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Abrir Orcamento")
    If lErro <> SUCESSO Then gError 133088

    lErro = Orcamento_ImpressaoCFE1(objVenda, colMensagem)
    If lErro <> SUCESSO Then gError 133088
    
    For iIndice = 1 To colMensagem.Count
            
        sMensagem = colMensagem.Item(iIndice)
        If sMensagem = "" Then sMensagem = "    "
        lErro = AFRAC_ImprimirRelatorioGerencial2(sMensagem, objTela, objRel)
        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Orcamento")
        If lErro <> SUCESSO Then gError 99893
            
    Next
        
    'guilhotina
    lErro = AFRAC_AcionarGuilhotina2("P", objRel)
    lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Acionar Guilhotina")
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    lErro = AFRAC_FecharRelatorioGerencial2(objTela, objRel)
    bRelAberto = False
    
    lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Fecha Orcamento")
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Orcamento_ImpressaoCFE = SUCESSO
    
    Exit Function
    
Erro_Orcamento_ImpressaoCFE:

    Orcamento_ImpressaoCFE = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 99893
            Call AFRAC_FecharRelatorioGerencial(objTela)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149439)

    End Select

    If bRelAberto Then Call AFRAC_FecharRelatorioGerencial2(objTela, objRel)

    Exit Function

End Function

Function Imprime_OrcamentoECF_EmTrans(dtDataFinal As Date, ByVal lNumOrcamento As Long, ByVal objTela As Object, ByVal objVenda As ClassVenda) As Long
'imprime o DAV

Dim lErro As Long
Dim colMensagem As New Collection
Dim objItens As ClassItemCupomFiscal
Dim objProduto As ClassProduto
Dim iIndice As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim bCont As Boolean
Dim sMensagem As String
Dim dTotal As Double
Dim dPreco As Double
Dim bAchou As Boolean
Dim colOrcamento As New Collection
Dim iIndice1 As Integer
Dim objCliente As ClassCliente
Dim lIndice As Long
Dim sCliente As String
Dim sCPFCGC As String
Dim sVendedor As String
Dim objVendedor As ClassVendedor
Dim sProduto As String
Dim sCNPJ As String
Dim iPos As Integer
Dim lNumDAV As Long
Dim sRetorno As String
Dim objProdutoNomeRed As TextBox
Dim sTexto As String

On Error GoTo Erro_Imprime_OrcamentoECF_EmTrans

'    'le os registros do orcamento e loca o arquivo
'    lErro = CF_ECF("OrcamentoECF_Le", colOrcamento)
'    If lErro <> SUCESSO Then gError 105878
'
'    For iIndice = 1 To colOrcamento.Count
'       Set objVenda = colOrcamento.Item(iIndice)
'       'Se achou o Orçamento com o Número
'       If objVenda.iTipo = OPTION_ORCAMENTO And objVenda.objCupomFiscal.lNumOrcamento = lNumOrcamento Then
'
'            'descobre o nome reduzido dos produtos
'            For Each objItens In objVenda.objCupomFiscal.colItens
'                For iIndice1 = 1 To gaobjProdutosNome.Count
'                    Set objProduto = gaobjProdutosNome.Item(iIndice1)
'                    If objItens.sProduto = objProduto.sCodigo Then
'                        objItens.sProdutoNomeRed = objProduto.sNomeReduzido
'                        Exit For
'                    End If
'                Next
'            Next
'
'           bAchou = True
'           Exit For
'       End If
'    Next
    
'    If Not (bAchou) Then
'        For iIndice = 1 To gcolVendas.Count
'           Set objVenda = gcolVendas.Item(iIndice)
'           'Se achou o Orçamento com o Número
'           If objVenda.iTipo = OPTION_ORCAMENTO And objVenda.objCupomFiscal.lNumOrcamento = lNumOrcamento Then
'               'Apaga o orçamento
'               bAchou = True
'               Exit For
'           End If
'        Next
'    End If
    
'    If Not (bAchou) Then
'        vbMsgRes = Rotina_AvisoECF(vbOK, AVISO_ORCAMENTO_INEXISTENTE)
'        Exit Function
'    End If
    
    If objVenda.objCupomFiscal.lNumOrcamento = 0 Then gError 210471
    
    'descobre o nome reduzido dos produtos
'    For Each objItens In objVenda.objCupomFiscal.colItens
'        For iIndice1 = 1 To gaobjProdutosNome.Count
'            Set objProduto = gaobjProdutosNome.Item(iIndice1)
'            If objItens.sProduto = objProduto.sCodigo Then
'                objItens.sProdutoNomeRed = objProduto.sNomeReduzido
'                Exit For
'            End If
'        Next
'    Next
    
    For Each objItens In objVenda.objCupomFiscal.colItens
        
        Set objProduto = New ClassProduto

        lErro = CF_ECF("Produtos_Le", objItens.sProduto, objProduto)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 214855
        
        If lErro = SUCESSO Then
            objItens.sProdutoNomeRed = objProduto.sNomeReduzido
        End If
        
        
    Next
    
    If objVenda.objCupomFiscal.lNumeroDAV = 0 And objVenda.objCupomFiscal.iStatus <> STATUS_BAIXADO Then
    
        'Função que obtém o próximo número de DAV
        lErro = CF_ECF("DAV_Obtem_Num_Automatico_EmTrans", lNumDAV)
        If lErro <> SUCESSO Then gError 210470
        
        objVenda.objCupomFiscal.lNumeroDAV = lNumDAV

    End If
    
    objVenda.objCupomFiscal.iDAVImpresso = 1
    
    Set objCliente = gobjClienteCPF.Busca(objVenda.objCupomFiscal.sCPFCGC, lIndice)
    If Not objCliente Is Nothing Then sCliente = objCliente.sNomeReduzido & " - "
    Call Formata_CNPJ_CPF(sCPFCGC, objVenda.objCupomFiscal.sCPFCGC)
    sCliente = sCliente & sCPFCGC
    
'    For Each objVendedor In gcolVendedores
'        'verifica se existe o vendedor na col
'        If objVendedor.iCodigo = objVenda.objCupomFiscal.iVendedor Then
'            sVendedor = objVendedor.sNomeReduzido & " - "
'            Exit For
'        End If
'    Next
    
    Call Formata_CNPJ_CPF(sCNPJ, gsCNPJ)
    
'    sVendedor = sVendedor & CStr(objVenda.objCupomFiscal.iVendedor)
    
    If AFRAC_ImpressoraCFe(giCodModeloECF) Then
    
        lErro = Orcamento_ImpressaoCFE(objTela, objVenda)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Else
    
        'imprime o DAV em impressoras nao fiscais usando o formato do Anexo 2 do ato COTEPE 06/08
        If giUsaImpressoraFiscal = 0 Then
            colMensagem.Add Chr(124) & Formata_Campo(ALINHAMENTO_DIREITA, 84, Chr(151), "") & Chr(124)
            colMensagem.Add Chr(124) & Formata_Campo(ALINHAMENTO_ESQUERDA, 63, " ", "DOCUMENTO AUXILIAR DE VENDA - ORÇAMENTO") & Formata_Campo(ALINHAMENTO_ESQUERDA, 22, " ", Chr(124))
            colMensagem.Add Chr(124) & Formata_Campo(ALINHAMENTO_ESQUERDA, 85, Chr(151), Chr(124))
            colMensagem.Add Chr(124) & " NÃO É DOCUMENTO FISCAL - NÃO É VÁLIDO COMO RECIBO E COMO GARANTIA DE MERCADORIA -  " & Chr(124)
            colMensagem.Add Chr(124) & Formata_Campo(ALINHAMENTO_ESQUERDA, 52, " ", "NÃO COMPROVA PAGAMENTO") & Formata_Campo(ALINHAMENTO_ESQUERDA, 33, " ", Chr(124))
            colMensagem.Add Chr(124) & Formata_Campo(ALINHAMENTO_ESQUERDA, 85, Chr(151), Chr(124))
            colMensagem.Add Chr(124) & Formata_Campo(ALINHAMENTO_ESQUERDA, 61, " ", "Identificação do Estabelecimento Emitente") & Formata_Campo(ALINHAMENTO_ESQUERDA, 24, " ", Chr(124))
            colMensagem.Add Chr(124) & Formata_Campo(ALINHAMENTO_ESQUERDA, 85, Chr(151), Chr(124))
            colMensagem.Add Chr(124) & Formata_Campo(ALINHAMENTO_DIREITA, 55, " ", " Denominação:" & gsNomeEmpresa) & Formata_Campo(ALINHAMENTO_DIREITA, 29, " ", Chr(124) & " CNPJ: " & sCNPJ) & Chr(124)
            colMensagem.Add Chr(124) & Formata_Campo(ALINHAMENTO_ESQUERDA, 85, Chr(151), Chr(124))
            colMensagem.Add Chr(124) & Formata_Campo(ALINHAMENTO_DIREITA, 84, " ", " Identificação do Destinatário") & Chr(124)
            colMensagem.Add Chr(124) & Formata_Campo(ALINHAMENTO_ESQUERDA, 85, Chr(151), Chr(124))
            colMensagem.Add Chr(124) & Formata_Campo(ALINHAMENTO_DIREITA, 55, " ", " Nome:" & objVenda.objCupomFiscal.sNomeCliente) & Formata_Campo(ALINHAMENTO_DIREITA, 29, " ", Chr(124) & " CNPJ/CPF: " & sCPFCGC) & Chr(124)
            colMensagem.Add Chr(124) & Formata_Campo(ALINHAMENTO_ESQUERDA, 85, Chr(151), Chr(124))
            colMensagem.Add Chr(124) & Formata_Campo(ALINHAMENTO_DIREITA, 40, " ", " N" & Chr(176) & " do Documento: " & Format(objVenda.objCupomFiscal.lNumeroDAV, "0000000000")) & Formata_Campo(ALINHAMENTO_DIREITA, 44, " ", Chr(124) & " N" & Chr(176) & " do Documento Fiscal: ______________") & Chr(124)
            colMensagem.Add Chr(124) & Formata_Campo(ALINHAMENTO_ESQUERDA, 85, Chr(151), Chr(124))
        
            For Each objItens In objVenda.objCupomFiscal.colItens
            
                    sProduto = objItens.sProdutoNomeRed
                    
                    Call TP_Produto_Le_Col(gaobjProdutosReferencia, gaobjProdutosCodBarras, gaobjProdutosNome, sProduto, objProduto)
                            
                    'caso o produto não seja encontrado
                    If objProduto Is Nothing Then gError 99894
                        
                    'Joga no cupom o item
                    colMensagem.Add Chr(124) & " " & Formata_Campo(ALINHAMENTO_ESQUERDA, 3, 0, objItens.iItem) & " " & Formata_Campo(ALINHAMENTO_DIREITA, 79, " ", objProduto.sCodigo & " " & objProduto.sDescricao) & Chr(124)
                    
                    sTexto = Chr(124) & Formata_Campo(ALINHAMENTO_ESQUERDA, 10, " ", objItens.dQuantidade) & Formata_Campo(ALINHAMENTO_CENTRALIZADO, 8, " ", "x") & Formata_Campo(ALINHAMENTO_DIREITA, 10, " ", Format(objItens.dPrecoUnitario, "standard")) & "  =  " & Formata_Campo(ALINHAMENTO_ESQUERDA, 15, " ", Format(objItens.dPrecoUnitario * objItens.dQuantidade, "standard"))
                    
                    If objItens.iStatus = STATUS_CANCELADO Then
                        sTexto = sTexto & "     ***** ITEM CANCELADO *****     " & Chr(124)
                    Else
                        sTexto = sTexto & Space(36) & Chr(124)
                    End If
                    
                    colMensagem.Add sTexto
                    
    '                colMensagem.Add Chr(124) & Formata_Campo(ALINHAMENTO_ESQUERDA, 10, " ", objItens.dQuantidade) & Formata_Campo(ALINHAMENTO_CENTRALIZADO, 8, " ", "x") & Formata_Campo(ALINHAMENTO_DIREITA, 10, " ", Format(objItens.dPrecoUnitario, "standard")) & "  =  " & Formata_Campo(ALINHAMENTO_ESQUERDA, 15, " ", Format(objItens.dPrecoUnitario * objItens.dQuantidade, "standard")) & Space(36) & Chr(124)
                    If objItens.iStatus = STATUS_ATIVO Then
                        dPreco = objItens.dPrecoUnitario * objItens.dQuantidade
                    End If
                                
                    'se existir desconto sobre o item...
                    If objItens.dPercDesc > 0 Then
                        sTexto = Chr(124) & Formata_Campo(ALINHAMENTO_DIREITA, 15, " ", "DESC. ITEM:") & Formata_Campo(ALINHAMENTO_DIREITA, 10, " ", Format(objItens.dPercDesc, "standard") & "%") & "     =  " & Formata_Campo(ALINHAMENTO_ESQUERDA, 15, " ", Format(objItens.dPrecoUnitario * objItens.dQuantidade * (1 - (objItens.dPercDesc / 100)), "standard")) & Space(36) & Chr(124)
                        colMensagem.Add sTexto
                        If objItens.iStatus = STATUS_ATIVO Then
                            dPreco = dPreco - (objItens.dPrecoUnitario * objItens.dQuantidade * objItens.dPercDesc / 100)
                        End If
                    ElseIf objItens.dValorDesconto > 0 Then
                        'sTexto = Chr(124) & Formata_Campo(ALINHAMENTO_DIREITA, 15, " ", "DESC. ITEM:") & Formata_Campo(ALINHAMENTO_DIREITA, 10, " ", Format(-objItens.dValorDesconto, "standard")) & "     =  " & Formata_Campo(ALINHAMENTO_ESQUERDA, 15, " ", Format((objItens.dPrecoUnitario - objItens.dValorDesconto) * objItens.dQuantidade, "standard")) & Space(36) & Chr(124)
                        sTexto = Chr(124) & Formata_Campo(ALINHAMENTO_DIREITA, 15, " ", "DESC. ITEM:") & Formata_Campo(ALINHAMENTO_DIREITA, 10, " ", Format(-objItens.dValorDesconto, "standard")) & "     =  " & Formata_Campo(ALINHAMENTO_ESQUERDA, 15, " ", Format((objItens.dPrecoUnitario * objItens.dQuantidade - objItens.dValorDesconto), "standard")) & Space(36) & Chr(124)
                        colMensagem.Add sTexto
                        If objItens.iStatus = STATUS_ATIVO Then
                            'dPreco = dPreco - (objItens.dQuantidade * objItens.dValorDesconto)
                            dPreco = dPreco - objItens.dValorDesconto
                        End If
                    End If
                    If objItens.iStatus = STATUS_ATIVO Then
                        dTotal = dTotal + dPreco
                    End If
            Next
                              
            colMensagem.Add Chr(124) & Formata_Campo(ALINHAMENTO_ESQUERDA, 48, " ", "TOTAL :   " & Format(dTotal, "standard")) & Space(36) & Chr(124)
            
            If (objVenda.objCupomFiscal.dValorDesconto + objVenda.objCupomFiscal.dValorDesconto1) <> 0 Then
                colMensagem.Add Chr(124) & Formata_Campo(ALINHAMENTO_ESQUERDA, 48, " ", "DESCONTO  :  -" & Format(objVenda.objCupomFiscal.dValorDesconto + objVenda.objCupomFiscal.dValorDesconto1, "standard")) & Space(36) & Chr(124)
                dTotal = dTotal - (objVenda.objCupomFiscal.dValorDesconto + objVenda.objCupomFiscal.dValorDesconto1)
            End If
            
            If objVenda.objCupomFiscal.dValorAcrescimo <> 0 Then
                colMensagem.Add Chr(124) & Formata_Campo(ALINHAMENTO_ESQUERDA, 48, " ", "ACRESCIMO :  +" & Format(objVenda.objCupomFiscal.dValorAcrescimo, "standard")) & Space(36) & Chr(124)
                dTotal = dTotal + objVenda.objCupomFiscal.dValorAcrescimo
            End If
            
            If (objVenda.objCupomFiscal.dValorDesconto + objVenda.objCupomFiscal.dValorDesconto1) <> 0 Or objVenda.objCupomFiscal.dValorAcrescimo <> 0 Then
                colMensagem.Add Chr(124) & Formata_Campo(ALINHAMENTO_ESQUERDA, 48, " ", "TOTAL LIQ.:   " & Format(dTotal, "standard")) & Space(36) & Chr(124)
            End If
            
            colMensagem.Add Chr(124) & Formata_Campo(ALINHAMENTO_ESQUERDA, 85, Chr(151), Chr(124))
            colMensagem.Add Chr(124) & Formata_Campo(ALINHAMENTO_CENTRALIZADO, 84, " ", "É vedada a autenticação deste documento") & Chr(124)
            colMensagem.Add Chr(124) & Formata_Campo(ALINHAMENTO_ESQUERDA, 85, Chr(151), Chr(124))
        
        'imprime o DAV em impressoras fiscais
        Else
        
            colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "Emitente: " & gsNomeEmpresa)
            colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "CNPJ Emitente: " & sCNPJ)
            colMensagem.Add TRACO_CAB
            colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "Destinatário: " & objVenda.objCupomFiscal.sNomeCliente)
            colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "CPF/CNPJ Destinatário: " & sCPFCGC)
            colMensagem.Add TRACO_CAB
            colMensagem.Add " N" & Chr(176) & " do Documento: " & Format(objVenda.objCupomFiscal.lNumeroDAV, "0000000000")
            colMensagem.Add TRACO_CAB
    
        
            For Each objItens In objVenda.objCupomFiscal.colItens
            
                sProduto = objItens.sProdutoNomeRed
                
                Call TP_Produto_Le_Col(gaobjProdutosReferencia, gaobjProdutosCodBarras, gaobjProdutosNome, sProduto, objProduto)
                        
                'caso o produto não seja encontrado
                If objProduto Is Nothing Then gError 99894
                    
                'Joga no cupom o item
                colMensagem.Add Formata_Campo(ALINHAMENTO_ESQUERDA, 3, 0, objItens.iItem) & " " & Formata_Campo(ALINHAMENTO_DIREITA, 45, " ", objProduto.sCodigo & " " & objProduto.sDescricao)
                colMensagem.Add Formata_Campo(ALINHAMENTO_ESQUERDA, 10, " ", objItens.dQuantidade) & Formata_Campo(ALINHAMENTO_CENTRALIZADO, 8, " ", "x") & Formata_Campo(ALINHAMENTO_DIREITA, 10, " ", Format(objItens.dPrecoUnitario, "standard")) & "  =  " & Formata_Campo(ALINHAMENTO_ESQUERDA, 15, " ", Format(objItens.dPrecoUnitario * objItens.dQuantidade, "standard"))
                dPreco = objItens.dPrecoUnitario * objItens.dQuantidade
                            
                'se existir desconto sobre o item...
                If objItens.dPercDesc > 0 Then
                    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 15, " ", "DESC. ITEM:") & Formata_Campo(ALINHAMENTO_DIREITA, 10, " ", Format(objItens.dPercDesc, "standard") & "%") & "     =  " & Formata_Campo(ALINHAMENTO_ESQUERDA, 15, " ", Format(objItens.dPrecoUnitario * objItens.dQuantidade * (1 - (objItens.dPercDesc / 100)), "standard"))
                    dPreco = dPreco - (objItens.dPrecoUnitario * objItens.dQuantidade * objItens.dPercDesc / 100)
                ElseIf objItens.dValorDesconto > 0 Then
'                    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 15, " ", "DESC. ITEM:") & Formata_Campo(ALINHAMENTO_DIREITA, 10, " ", Format(-objItens.dValorDesconto, "standard")) & "     =  " & Formata_Campo(ALINHAMENTO_ESQUERDA, 15, " ", Format((objItens.dPrecoUnitario - objItens.dValorDesconto) * objItens.dQuantidade, "standard"))
'                    dPreco = dPreco - (objItens.dQuantidade * objItens.dValorDesconto)
                    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 15, " ", "DESC. ITEM:") & Formata_Campo(ALINHAMENTO_DIREITA, 10, " ", Format(-objItens.dValorDesconto, "standard")) & "     =  " & Formata_Campo(ALINHAMENTO_ESQUERDA, 15, " ", Format((objItens.dPrecoUnitario * objItens.dQuantidade - objItens.dValorDesconto), "standard"))
                    dPreco = dPreco - objItens.dValorDesconto
                End If
                
                If objItens.iStatus = STATUS_ATIVO Then
                    dTotal = dTotal + dPreco
                ElseIf objItens.iStatus = STATUS_CANCELADO Then
                    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "************* ITEM " & Format(objItens.iItem, "000") & " CANCELADO *************")
                End If
                
            Next
                              
            colMensagem.Add Formata_Campo(ALINHAMENTO_ESQUERDA, 48, " ", "TOTAL :   " & Format(dTotal, "standard"))
            
            If (objVenda.objCupomFiscal.dValorDesconto + objVenda.objCupomFiscal.dValorDesconto1) <> 0 Then
                colMensagem.Add Formata_Campo(ALINHAMENTO_ESQUERDA, 48, " ", "DESCONTO  :  -" & Format(objVenda.objCupomFiscal.dValorDesconto + objVenda.objCupomFiscal.dValorDesconto1, "standard"))
                dTotal = dTotal - (objVenda.objCupomFiscal.dValorDesconto + objVenda.objCupomFiscal.dValorDesconto1)
            End If
            
            If objVenda.objCupomFiscal.dValorAcrescimo <> 0 Then
                colMensagem.Add Formata_Campo(ALINHAMENTO_ESQUERDA, 48, " ", "ACRESCIMO :  +" & Format(objVenda.objCupomFiscal.dValorAcrescimo, "standard"))
                dTotal = dTotal + objVenda.objCupomFiscal.dValorAcrescimo
            End If
            
            If (objVenda.objCupomFiscal.dValorDesconto + objVenda.objCupomFiscal.dValorDesconto1) <> 0 Or objVenda.objCupomFiscal.dValorAcrescimo <> 0 Then
                colMensagem.Add Formata_Campo(ALINHAMENTO_ESQUERDA, 48, " ", "TOTAL LIQ.:   " & Format(dTotal, "standard"))
            End If
        
    
    
        End If
    
        lErro = AFRAC_AbrirRelatorioGerencial(RELGER_ORCAMENTO, objTela)
        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Abrir Orcamento")
        If lErro <> SUCESSO Then gError 133088
    
        For iIndice = 1 To colMensagem.Count
                
                
            sMensagem = colMensagem.Item(iIndice)
            If sMensagem = "" Then sMensagem = "    "
            lErro = AFRAC_ImprimirRelatorioGerencial(sMensagem, objTela)
            lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Orcamento")
            If lErro <> SUCESSO Then gError 99893
                
                
    '            lErro = AFRAC_RegistrarNaoFiscal("", 0, sMensagem)
    '            lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Registrar não Fiscal")
    '            If lErro <> SUCESSO Then
    '                'Impressora não responde
    '                vbMsgRes = Rotina_AvisoECF(vbYesNo, AVISO_IMPRESSORA_NAO_RESPONDE)
    '                If vbMsgRes = vbNo Then
    '                    gError 99803
    '                Else
    '                    bCont = False
    '                End If
    '            End If
        Next
            
    
        iPos = InStr(1, objTela.Controls("RT1").Text, "NÃO É DOCUMENTO FISCAL")
        objTela.Controls("RT1").SelStart = iPos
        objTela.Controls("RT1").SelLength = 81
        objTela.Controls("RT1").SelBold = True
        objTela.Controls("RT1").SelFontSize = 14
        
        iPos = InStr(1, objTela.Controls("RT1").Text, "NÃO COMPROVA PAGAMENTO")
        objTela.Controls("RT1").SelStart = iPos
        objTela.Controls("RT1").SelLength = 22
        objTela.Controls("RT1").SelBold = True
        objTela.Controls("RT1").SelFontSize = 14
        
        'se for impressora fiscal
        If giCodModeloECF <> 4 Then
        
            'le o numero do cupom
            lErro = AFRAC_LerInformacaoImpressora("023", sRetorno)
            lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Informação Impressora")
            If lErro <> SUCESSO Then gError 204319
            
            'PAFECF - se for impressora fiscal - guarda o numero do COO associado ao DAV pois vai ser necessario posteriormente
            objVenda.objCupomFiscal.lNumero = StrParaLong(sRetorno)
        
        Else
        
            'senao zera o ECF
            objVenda.objCupomFiscal.iECF = 0
            
        End If
    
        lErro = AFRAC_FecharRelatorioGerencial(objTela)
        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Fecha Orcamento")
        If lErro <> SUCESSO Then gError 99898
        
        
        objVenda.objCupomFiscal.sMarcaECF = gsMarcaECF
        objVenda.objCupomFiscal.sTipoECF = gsTipoECF
        objVenda.objCupomFiscal.sModeloECF = gsModeloECF
        objVenda.objCupomFiscal.sNumSerieECF = gsNumSerie
        
        
        'grava os dados do DAV em arquivo para permitir posterior recuperação
        lErro = CF_ECF("Grava_Orcamento_ECF1_EmTrans", objVenda)
        If lErro <> SUCESSO Then gError 204324
    
    End If
    
'    lErro = CF_ECF("DAVECF_Grava_EmTrans", objVenda)
'    If lErro <> SUCESSO Then gError 204329
    
    Imprime_OrcamentoECF_EmTrans = SUCESSO
    
    Exit Function
    
Erro_Imprime_OrcamentoECF_EmTrans:

    Imprime_OrcamentoECF_EmTrans = gErr

    Select Case gErr

        Case 99803, 99894 To 99898, 105868, 133088, 204319, 204324, 204329, 210470, 214855, ERRO_SEM_MENSAGEM

        Case 99893
            Call AFRAC_FecharRelatorioGerencial(objTela)

        Case 210471
            Call Rotina_ErroECF(vbOKOnly, ERRO_ORCAMENTO_NAO_PREENCHIDO, lErro)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149439)

    End Select

    Exit Function

End Function

Public Function MovimentoDinheiro_Form_Load(ByVal objSangria As Object, ByVal objSuprimento As Object) As Long

Dim lErro As Long
Dim iAcao As Integer
Dim vbMsgRes As VbMsgBoxResult
    
On Error GoTo Erro_MovimentoDinheiro_Form_Load
    
    MovimentoDinheiro_Form_Load = SUCESSO
    
    Exit Function

Erro_MovimentoDinheiro_Form_Load:
    
    MovimentoDinheiro_Form_Load = gErr
    
    Select Case gErr
                
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149440)

    End Select
    
    Exit Function

End Function

Public Function Customiza_Menu_Principal(ByVal objTela As Object) As Long

Dim lErro As Long
Dim iAcao As Integer
Dim vbMsgRes As VbMsgBoxResult
    
On Error GoTo Erro_Customiza_Menu_Principal
    
    Customiza_Menu_Principal = SUCESSO
    
    Exit Function

Erro_Customiza_Menu_Principal:
    
    Customiza_Menu_Principal = gErr
    
    Select Case gErr
                
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149441)

    End Select
    
    Exit Function

End Function

Public Function Testa_Limite_Desconto(ByVal objVenda As ClassVenda) As Long

Dim lErro As Long
Dim dDescontoTotal As Double
Dim objItem As ClassItemCupomFiscal
Dim objProduto As ClassProduto
Dim dDescontoMaximo As Double
    
On Error GoTo Erro_Testa_Limite_Desconto
    
    dDescontoTotal = objVenda.objCupomFiscal.dValorDesconto + objVenda.objCupomFiscal.dValorDesconto1

    For Each objItem In objVenda.objCupomFiscal.colItens
    
        Set objProduto = gaobjProdutosNome.Busca(objItem.sProdutoNomeRed)
    
        If Not objProduto Is Nothing Then
        
            If objProduto.dLimiteDesconto > 0 Then
        
                dDescontoMaximo = dDescontoMaximo + (objProduto.dLimiteDesconto * (objItem.dPrecoUnitario * objItem.dQuantidade))
        
            Else
            
                dDescontoMaximo = dDescontoMaximo + (objItem.dPrecoUnitario * objItem.dQuantidade)
            
            End If
        Else
        
            dDescontoMaximo = dDescontoMaximo + (objItem.dPrecoUnitario * objItem.dQuantidade)
        
        End If
        
    Next
    
    If dDescontoTotal > dDescontoMaximo Then gError 126778
    
    Testa_Limite_Desconto = SUCESSO
    
    Exit Function

Erro_Testa_Limite_Desconto:
    
    Testa_Limite_Desconto = gErr
    
    Select Case gErr
                
        Case 126778
            Call Rotina_ErroECF(vbOKOnly, ERRO_LIMITE_DESCONTO_ULTRAPASSADO, gErr, dDescontoMaximo, dDescontoTotal)
            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149442)

    End Select
    
    Exit Function

End Function

Function Carrega_Caixa_Config() As Long

Dim sRetorno As String
Dim lTamanho As Long
Dim objECF As New ClassECFConfig
Dim sArquivo As String
Dim sArquivo1 As String
Dim sDir As String
Dim lErro As Long

On Error GoTo Erro_Carrega_Caixa_Config

    If giDebug = 1 Then MsgBox ("10")
    
    sDir = String(STRING_MAX_NOME_ARQUIVO, 0)
    
    lErro = GetWindowsDirectory(sDir, STRING_MAX_NOME_ARQUIVO)
    If lErro = 0 Then gError 126369
    
    If giDebug = 1 Then MsgBox ("10.01")
    
    sDir = StringZ(sDir)
    
    If right(sDir, 1) <> "\" Then sDir = sDir & "\"
    
    sArquivo = sDir & NOME_ARQUIVO_CAIXA
        
    sArquivo1 = Dir(sArquivo)
    
    If Len(sArquivo1) = 0 Then gError 126370
        
    lTamanho = 150
    sRetorno = String(lTamanho, 0)
    
    If giDebug = 1 Then MsgBox ("10.02")
    
    'funcao que verifica se o caixaconfig.ini foi alterado
    lErro = CF_ECF("CaixaConfig_Verifica_Chave")
    If lErro <> SUCESSO Then gError 204244
    
    If giDebug = 1 Then MsgBox ("10.1")
    
    'verifica se o caixa possui ECF
'    Call GetPrivateProfileString(APLICACAO_CAIXA, "RemoveOrcamento", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
'    If sRetorno = String(lTamanho, 0) Then gError 105760
'
'    giRemoveOrc = StrParaInt(sRetorno)

    Call GetPrivateProfileString(APLICACAO_CAIXA, "CodCaixa", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
    If sRetorno = String(lTamanho, 0) Then gError 105686
    
    giCodCaixa = StrParaInt(sRetorno)
                    
    sRetorno = String(lTamanho, 0)
    Call GetPrivateProfileString(APLICACAO_CAIXA, "StatusCaixa", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
    If sRetorno = String(lTamanho, 0) Then gError 105687
    
    giStatusCaixa = StrParaInt(sRetorno)
        
    If giDebug = 1 Then MsgBox ("10.2")
        
'    sRetorno = String(lTamanho, 0)
'    Call GetPrivateProfileString(APLICACAO_CAIXA, "NumProxOrcamento", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
'    If sRetorno <> String(lTamanho, 0) Then glNumProxOrcamento = StrParaLong(sRetorno)
'
'    If glNumProxOrcamento = 0 Then glNumProxOrcamento = 1
    
'    sRetorno = String(lTamanho, 0)
'    Call GetPrivateProfileString(APLICACAO_CAIXA, "NumProxTransfCaixa", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
'    If sRetorno <> String(lTamanho, 0) Then glNumProxTransfCaixa = StrParaLong(sRetorno)
'
'    If glNumProxTransfCaixa = 0 Then glNumProxTransfCaixa = 1
    
    
'    sRetorno = String(lTamanho, 0)
'    Call GetPrivateProfileString(APLICACAO_CAIXA, "TipoTerminal", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
'    If sRetorno <> String(lTamanho, 0) Then giTEF = StrParaInt(sRetorno)
'
    sRetorno = String(lTamanho, 0)
    Call GetPrivateProfileString(APLICACAO_CAIXA, "Teclado", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
    If sRetorno <> String(lTamanho, 0) Then giTeclado = StrParaInt(sRetorno)
    
'    sRetorno = String(lTamanho, 0)
'    Call GetPrivateProfileString(APLICACAO_CAIXA, "POS", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
'    gsPOS = sRetorno
'
    sRetorno = String(lTamanho, 0)
    Call GetPrivateProfileString(APLICACAO_EMPRESA, "FilialEmpresa", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
    If sRetorno = String(lTamanho, 0) Then gError 105688
    
    giFilialEmpresa = StrParaInt(sRetorno)
    
    If giDebug = 1 Then MsgBox ("10.3")
    
    sRetorno = String(lTamanho, 0)
    Call GetPrivateProfileString(APLICACAO_SESSAO, "StatusSessao", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
    If sRetorno <> String(lTamanho, 0) Then giStatusSessao = StrParaInt(sRetorno)
    
    sRetorno = String(lTamanho, 0)
    Call GetPrivateProfileString(APLICACAO_SESSAO, "CodOperador", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
    If sRetorno <> String(lTamanho, 0) Then giCodOperador = StrParaInt(sRetorno)

'    sRetorno = String(lTamanho, 0)
'    Call GetPrivateProfileString(APLICACAO_TRANSACAO, "SeqUltimaTransacaoEncerrada", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
'    If sRetorno <> String(lTamanho, 0) Then glSeqTransacaoEncerrada = StrParaInt(sRetorno)
'
'    sRetorno = String(lTamanho, 0)
'    Call GetPrivateProfileString(APLICACAO_TRANSACAO, "SeqUltimaTransacaoAberta", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
'    If sRetorno <> String(lTamanho, 0) Then glSeqTransacaoAberta = StrParaInt(sRetorno)
    
    'agora esta sendo preenchido com dados de dadoscc.ccc
'    sRetorno = String(lTamanho, 0)
'    Call GetPrivateProfileString(APLICACAO_ECF, "CodECF", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
'    If sRetorno <> String(lTamanho, 0) Then giCodECF = StrParaInt(sRetorno)
    
    sRetorno = String(lTamanho, 0)
    Call GetPrivateProfileString(APLICACAO_ECF, "DataUltimaReducaoZ", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
    If sRetorno <> String(lTamanho, 0) Then
        gdtUltimaReducao = StrParaDate(sRetorno)
    Else
        gdtUltimaReducao = DATA_NULA
    End If
    
    If giDebug = 1 Then MsgBox ("10.4")
    
'    sRetorno = String(lTamanho, 0)
'    Call GetPrivateProfileString(APLICACAO_ECF, "ImpressoraCheque", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
'    If sRetorno <> String(lTamanho, 0) Then giImpressoraCheque = StrParaInt(sRetorno)
'
'    sRetorno = String(lTamanho, 0)
'    Call GetPrivateProfileString(APLICACAO_ECF, "ImpressoraECF", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
'    If sRetorno <> String(lTamanho, 0) Then giImpressoraECF = StrParaInt(sRetorno)
    
    sRetorno = String(lTamanho, 0)
    Call GetPrivateProfileString(APLICACAO_ECF, "NumProxIdent", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
    If sRetorno <> String(lTamanho, 0) Then glNumProxIdentificacao = StrParaLong(sRetorno)
    
    sRetorno = String(lTamanho, 0)
    Call GetPrivateProfileString(APLICACAO_ECF, "TipoTEF", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
    If sRetorno <> String(lTamanho, 0) Then giTipoTEF = StrParaInt(sRetorno)
    
    sRetorno = String(lTamanho, 0)
    Call GetPrivateProfileString(APLICACAO_ECF, "TEFImpressoraFolhaDupla", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
    If sRetorno <> String(lTamanho, 0) Then giTEFImpressoraFolhaDupla = StrParaInt(sRetorno)
    
    
    If giDebug = 1 Then MsgBox ("10.5")

    sRetorno = String(lTamanho, 0)
    Call GetPrivateProfileString(APLICACAO_EMPRESA, "Nome", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
    If sRetorno = String(lTamanho, 0) Then gError 105689
    
    If InStr(sRetorno, Chr(0)) > 0 Then
        gsNomeEmpresa = left(sRetorno, InStr(sRetorno, Chr(0)) - 1)
    Else
        gsNomeEmpresa = sRetorno
    End If
    
    Call GetPrivateProfileString(APLICACAO_EMPRESA, "CNPJ", "", sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
    
    If sRetorno <> String(lTamanho, 0) Then
    
        If InStr(sRetorno, Chr(0)) > 0 Then
            gsCNPJ = left(sRetorno, InStr(sRetorno, Chr(0)) - 1)
        Else
            gsCNPJ = sRetorno
        End If
    
    End If
    
    If giDebug = 1 Then MsgBox ("10.6")
    
    sRetorno = String(lTamanho, 0)
    Call GetPrivateProfileString(APLICACAO_EMPRESA, "IE", "", sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
    
    If sRetorno <> String(lTamanho, 0) Then
    
        If InStr(sRetorno, Chr(0)) > 0 Then
            gsInscricaoEstadual = left(sRetorno, InStr(sRetorno, Chr(0)) - 1)
        Else
            gsInscricaoEstadual = sRetorno
        End If
    
    End If
    
    sRetorno = String(lTamanho, 0)
    Call GetPrivateProfileString(APLICACAO_EMPRESA, "IM", "", sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
    
    If sRetorno <> String(lTamanho, 0) Then
    
        If InStr(sRetorno, Chr(0)) > 0 Then
            gsInscricaoMunicipal = left(sRetorno, InStr(sRetorno, Chr(0)) - 1)
        Else
            gsInscricaoMunicipal = sRetorno
        End If
    
    End If
    
    If giDebug = 1 Then MsgBox ("10.7")
    
    sRetorno = String(lTamanho, 0)
    Call GetPrivateProfileString(APLICACAO_EMPRESA, "Endereco", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
    
    If sRetorno <> String(lTamanho, 0) Then
    
        If InStr(sRetorno, Chr(0)) > 0 Then
            gsEndereco = left(sRetorno, InStr(sRetorno, Chr(0)) - 1)
        Else
            gsEndereco = sRetorno
        End If
        
    End If
    
    
    sRetorno = String(lTamanho, 0)
    Call GetPrivateProfileString(APLICACAO_EMPRESA, "Bairro", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
    
    If sRetorno <> String(lTamanho, 0) Then
    
        If InStr(sRetorno, Chr(0)) > 0 Then
            gsEnderecoComplemento = left(sRetorno, InStr(sRetorno, Chr(0)) - 1)
        Else
            gsEnderecoComplemento = sRetorno
        End If
        gsBairro = gsEnderecoComplemento
        
    End If
    
    sRetorno = String(lTamanho, 0)
    Call GetPrivateProfileString(APLICACAO_EMPRESA, "Cidade", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
    
    If sRetorno <> String(lTamanho, 0) Then
    
        If InStr(sRetorno, Chr(0)) > 0 Then
            gsCidade = left(sRetorno, InStr(sRetorno, Chr(0)) - 1)
        Else
            gsCidade = sRetorno
        End If
        
    End If
    
    sRetorno = String(lTamanho, 0)
    Call GetPrivateProfileString(APLICACAO_EMPRESA, "UF", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
    
    If sRetorno <> String(lTamanho, 0) Then
    
        If InStr(sRetorno, Chr(0)) > 0 Then
            gsUF = left(sRetorno, InStr(sRetorno, Chr(0)) - 1)
        Else
            gsUF = sRetorno
        End If
        
    End If
    
    sRetorno = String(lTamanho, 0)
    Call GetPrivateProfileString(APLICACAO_EMPRESA, "CEP", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
    
    If sRetorno <> String(lTamanho, 0) Then
    
        If InStr(sRetorno, Chr(0)) > 0 Then
            gsCEP = left(sRetorno, InStr(sRetorno, Chr(0)) - 1)
        Else
            gsCEP = sRetorno
        End If
        
    End If
    
    sRetorno = String(lTamanho, 0)
    Call GetPrivateProfileString(APLICACAO_EMPRESA, "CNAE", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
    
    If sRetorno <> String(lTamanho, 0) Then
    
        If InStr(sRetorno, Chr(0)) > 0 Then
            gsCNAE = left(sRetorno, InStr(sRetorno, Chr(0)) - 1)
        Else
            gsCNAE = sRetorno
        End If
        
    End If
    
    sRetorno = String(lTamanho, 0)
    Call GetPrivateProfileString(APLICACAO_EMPRESA, "NomeReduzido", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
    
    If sRetorno <> String(lTamanho, 0) Then
    
        If InStr(sRetorno, Chr(0)) > 0 Then
            gsNomeReduzido = left(sRetorno, InStr(sRetorno, Chr(0)) - 1)
        Else
            gsNomeReduzido = sRetorno
        End If
        
    End If
    
    sRetorno = String(lTamanho, 0)
    Call GetPrivateProfileString(APLICACAO_EMPRESA, "Logradouro", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
    
    If sRetorno <> String(lTamanho, 0) Then
    
        If InStr(sRetorno, Chr(0)) > 0 Then
            gsEndLogradouro = left(sRetorno, InStr(sRetorno, Chr(0)) - 1)
        Else
            gsEndLogradouro = sRetorno
        End If
        
    End If
    
    sRetorno = String(lTamanho, 0)
    Call GetPrivateProfileString(APLICACAO_EMPRESA, "Numero", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
    
    If sRetorno <> String(lTamanho, 0) Then
    
        If InStr(sRetorno, Chr(0)) > 0 Then
            gsEndNumero = left(sRetorno, InStr(sRetorno, Chr(0)) - 1)
        Else
            gsEndNumero = sRetorno
        End If
        
    End If
    
    sRetorno = String(lTamanho, 0)
    Call GetPrivateProfileString(APLICACAO_EMPRESA, "Complemento", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
    
    If sRetorno <> String(lTamanho, 0) Then
    
        If InStr(sRetorno, Chr(0)) > 0 Then
            gsEndComplemento = left(sRetorno, InStr(sRetorno, Chr(0)) - 1)
        Else
            gsEndComplemento = sRetorno
        End If
        
    End If
    
    If giDebug = 1 Then MsgBox ("10.8")
    
'    sRetorno = String(lTamanho, 0)
'    Call GetPrivateProfileString(APLICACAO_CAIXA, "NumProxMovto", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
'    If sRetorno <> String(lTamanho, 0) Then glNumProxMovto = StrParaLong(sRetorno)
    
    sRetorno = String(lTamanho, 0)
    Call GetPrivateProfileString(APLICACAO_CODBARRAS, "Porta", "", sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
    
    If sRetorno <> String(lTamanho, 0) Then
    
        If InStr(sRetorno, Chr(0)) > 0 Then
            gsCodBarrasPorta = Trim(left(sRetorno, InStr(sRetorno, Chr(0)) - 1))
        Else
            gsCodBarrasPorta = Trim(sRetorno)
        End If
    
    End If
    
    lTamanho = 255
    sRetorno = String(lTamanho, 0)
    
    Call GetPrivateProfileString(APLICACAO_DADOS, "DirMVTEF", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
    If sRetorno <> String(lTamanho, 0) Then gsDirMVTEF = StringZ(sRetorno)
    
    If right(gsDirMVTEF, 1) <> "\" Then gsDirMVTEF = gsDirMVTEF & "\"
    
    If giDebug = 1 Then MsgBox ("10.9")
    
    lTamanho = 15
    sRetorno = String(lTamanho, 0)
    
    'pega o código da empresa
    Call GetPrivateProfileString(APLICACAO_EMPRESA, "CodEmpresa", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
    If sRetorno <> String(lTamanho, 0) Then giCodEmpresa = StrParaInt(sRetorno)
    
    sRetorno = String(lTamanho, 0)
    Call GetPrivateProfileString(APLICACAO_CAIXA, "BalancaPorta", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
    If sRetorno <> String(lTamanho, 0) And StringZ(sRetorno) <> Str(CONSTANTE_ERRO) Then gsBalancaPorta = StringZ(sRetorno)
    
    sRetorno = String(lTamanho, 0)
    Call GetPrivateProfileString(APLICACAO_CAIXA, "BalancaModelo", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
    If sRetorno <> String(lTamanho, 0) And StringZ(sRetorno) <> Str(CONSTANTE_ERRO) Then giBalancaModelo = StrParaInt(sRetorno)
    
    sRetorno = String(lTamanho, 0)
    Call GetPrivateProfileString(APLICACAO_CAIXA, "BalancaNome", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
    If sRetorno <> String(lTamanho, 0) And StringZ(sRetorno) <> Str(CONSTANTE_ERRO) Then gsBalancaNome = StringZ(sRetorno)
    
    
    If giDebug = 1 Then MsgBox ("10.10")
    
    'funcao que verifica se o orcamento.ini foi alterado
    lErro = CF_ECF("OrcamentoIni_Verifica_Chave")
    If lErro <> SUCESSO Then gError 204303
        
    If giDebug = 1 Then MsgBox ("10.11")
    
    sRetorno = String(lTamanho, 0)
    Call GetPrivateProfileString(APLICACAO_ORCAMENTO, "PreVenda", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_ORCAMENTO)
    If sRetorno <> String(lTamanho, 0) Then giPreVenda = StrParaInt(sRetorno)
    
    sRetorno = String(lTamanho, 0)
    Call GetPrivateProfileString(APLICACAO_ORCAMENTO, "UsaImpressoraFiscal", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_ORCAMENTO)
    If sRetorno <> String(lTamanho, 0) Then giUsaImpressoraFiscal = StrParaInt(sRetorno)
    
    sRetorno = String(lTamanho, 0)
    Call GetPrivateProfileString(APLICACAO_ORCAMENTO, "DAV", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_ORCAMENTO)
    If sRetorno <> String(lTamanho, 0) Then giDAV = StrParaInt(sRetorno)
    
    If giPreVenda <> 0 And giDAV <> 0 Then gError 210783
    
'    If giPreVenda <> 0 And giUsaImpressoraFiscal <> 0 Then gError 210992
    
    If giDebug = 1 Then MsgBox ("10.12")
    
    lTamanho = 257
    sRetorno = String(lTamanho, 0)

    sRetorno = String(lTamanho, 0)
    Call GetPrivateProfileString(APLICACAO_RSA, "ChavePublica", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
    If sRetorno <> String(lTamanho, 0) Then gsChavePublica = StringZ(sRetorno)
    
    lTamanho = 257
    sRetorno = String(lTamanho, 0)
    
    Call GetPrivateProfileString(APLICACAO_RSA, "ChavePrivada", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
    If sRetorno <> String(lTamanho, 0) Then gsChavePrivada = StringZ(sRetorno)
    
    lTamanho = 15
    sRetorno = String(lTamanho, 0)
    
    'pega o código da empresa
    Call GetPrivateProfileString(APLICACAO_EMPRESA, "PapelAcabando", "10", sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
    If sRetorno <> String(lTamanho, 0) Then giIntervaloPapelAcabando = StrParaLong(sRetorno)
    
    
'    Call GetPrivateProfileString(APLICACAO_PAFECF, "Versao", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
'    If sRetorno <> String(lTamanho, 0) Then gsVersaoEspecPAFECF = StringZ(sRetorno)
    
    
'    lErro = CF_ECF("EcfCorporator_Descriptografa")
'    If lErro <> SUCESSO Then gError 204507
    
    
    If giDebug = 1 Then MsgBox ("11")
    
    Carrega_Caixa_Config = SUCESSO
    
    Exit Function

Erro_Carrega_Caixa_Config:

    Carrega_Caixa_Config = gErr
    
    Select Case gErr
        
        Case 105686
            Call Rotina_ErroECF(vbOKOnly, ERRO_PREENCHIMENTO_ARQUIVO_CONFIG, gErr, "CodCaixa", APLICACAO_CAIXA, NOME_ARQUIVO_CAIXA)
        
        Case 105687
            Call Rotina_ErroECF(vbOKOnly, ERRO_PREENCHIMENTO_ARQUIVO_CONFIG, gErr, "StatusCaixa", APLICACAO_CAIXA, NOME_ARQUIVO_CAIXA)
        
        Case 105688
            Call Rotina_ErroECF(vbOKOnly, ERRO_PREENCHIMENTO_ARQUIVO_CONFIG, gErr, "FilialEmpresa", APLICACAO_EMPRESA, NOME_ARQUIVO_CAIXA)
        
        Case 105689
            Call Rotina_ErroECF(vbOKOnly, ERRO_PREENCHIMENTO_ARQUIVO_CONFIG, gErr, "Nome", APLICACAO_EMPRESA, NOME_ARQUIVO_CAIXA)
        
        Case 105760
            Call Rotina_ErroECF(vbOKOnly, ERRO_PREENCHIMENTO_ARQUIVO_CONFIG, gErr, "RemoveOrcamento", APLICACAO_CAIXA, NOME_ARQUIVO_CAIXA)
        
        Case 126369
            Call Rotina_ErroECF(vbOKOnly, ERRO_GETWINDOWSDIRECTORY, gErr)
        
        Case 126370
            Call Rotina_ErroECF(vbOKOnly, ERRO_CAIXACONFIGINI_NAO_ENCONTRADO, gErr, sArquivo)
        
        Case 204244, 204303, 204507
        
        Case 204314
            Call Rotina_ErroECF(vbOKOnly, ERRO_PREENCHIMENTO_ARQUIVO_CONFIG, gErr, "CNPJ", APLICACAO_EMPRESA, NOME_ARQUIVO_CAIXA)
        
        Case 210783
            Call Rotina_ErroECF(vbOKOnly, ERRO_PREVENDA_DAV_SIMULTANEOS, gErr)
        
        Case 210992
            Call Rotina_ErroECF(vbOKOnly, ERRO_PREVENDA_CUPOM_SIMULTANEOS, gErr)
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_CARREGAMENTO_CAIXA_CONFIG, gErr, Error)
    
    End Select
    
    Exit Function
    
End Function

Function TabelaPrecoItens_Exclui_EmTrans(ByVal lComando As Long) As Long

Dim lErro As Long

On Error GoTo Erro_TabelaPrecoItens_Exclui_EmTrans

    lErro = Comando_Executar(lComando, "DELETE FROM TabelaPrecoItens")
    If lErro <> AD_SQL_SUCESSO Then gError 216210
        
    TabelaPrecoItens_Exclui_EmTrans = SUCESSO

    Exit Function
    
Erro_TabelaPrecoItens_Exclui_EmTrans:
    
    TabelaPrecoItens_Exclui_EmTrans = gErr
    
    Select Case gErr

        Case 216210
            Call Rotina_ErroECF(vbOKOnly, ERRO_EXCLUSAO_TABELAPRECOITENS, gErr)
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error, 216211)

    End Select

    Exit Function

End Function

Function ProdutoDesconto_Exclui_EmTrans(ByVal lComando As Long) As Long

Dim lErro As Long

On Error GoTo Erro_ProdutoDesconto_Exclui_EmTrans

    lErro = Comando_Executar(lComando, "DELETE FROM ProdutoDesconto")
    If lErro <> AD_SQL_SUCESSO Then gError 216212

    ProdutoDesconto_Exclui_EmTrans = SUCESSO

    Exit Function
    
Erro_ProdutoDesconto_Exclui_EmTrans:
    
    ProdutoDesconto_Exclui_EmTrans = gErr
    
    Select Case gErr
      
        Case 216212
            Call Rotina_ErroECF(vbOKOnly, ERRO_EXCLUSAO_PRODUTODESCONTO, gErr)
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error, 216213)

    End Select

    Exit Function

End Function

Function Carrega_Arquivo_FonteDados(Optional ByVal iDownload As Integer = 0) As Long

Dim sRegistro As String
Dim iPos As Integer
Dim sTipo As String
Dim iPosInicio As Integer
Dim lErro As Long
Dim objProduto As ClassProduto
Dim sArquivo As String
'Dim colTabPreco As New Collection
'Dim objTabPreco As New ClassTabelaPrecoItem
'Dim colProduto As New Collection
'Dim colProdutoDesconto As New Collection
Dim lTamanho As Long
Dim sRetorno As String
Dim objObject As Object
Dim lVersao As Long
Dim sCodProduto As String
Dim alComando(1 To 11) As Long
Dim iIndice As Integer
Dim lSequencial As Long
Dim lErroProduto As Long
Dim lErroCliente As Long
Dim lCodCliente As Long, dtUltimoMovto As Date
Dim lCodVendedor As Long
Dim lErroVendedor As Long

On Error GoTo Erro_Carrega_Arquivo_FonteDados
    
    
    If giDebug = 1 Then MsgBox ("12")
    
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_AbrirExt(glConexaoPAFECF)
        If alComando(iIndice) = 0 Then gError 214834
    Next
    
    sCodProduto = String(STRING_PRODUTO, 0)
    
    'se a tabela de produtos estiver vazia tenta carrega-la
    'normalmente ela sera carrega usando os botoes de carregar tabelas
    lErro = Comando_Executar(alComando(1), "SELECT Codigo FROM Produtos ", sCodProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 214835
    
    lErroProduto = Comando_BuscarPrimeiro(alComando(1))
    If lErroProduto <> AD_SQL_SUCESSO And lErroProduto <> AD_SQL_SEM_DADOS Then gError 214836
        
    'se a tabela de clientes estiver vazia tenta carrega-la
    'normalmente ela sera carrega usando os botoes de carregar tabelas
    lErro = Comando_Executar(alComando(2), "SELECT Codigo FROM Clientes ", lCodCliente)
    If lErro <> AD_SQL_SUCESSO Then gError 214835
    
    lErroCliente = Comando_BuscarPrimeiro(alComando(2))
    If lErroCliente <> AD_SQL_SUCESSO And lErroCliente <> AD_SQL_SEM_DADOS Then gError 214836
        
    'se a tabela de clientes estiver vazia tenta carrega-la
    'normalmente ela sera carrega usando os botoes de carregar tabelas
'    lErro = Comando_Executar(alComando(3), "SELECT Codigo FROM Vendedores ", lCodVendedor)
'    If lErro <> AD_SQL_SUCESSO Then gError 214950
'
'    lErroVendedor = Comando_BuscarPrimeiro(alComando(3))
'    If lErroVendedor <> AD_SQL_SUCESSO And lErroVendedor <> AD_SQL_SEM_DADOS Then gError 214951
        
'    If iDownload = 1 Or lErroProduto = AD_SQL_SEM_DADOS Or lErroCliente = AD_SQL_SEM_DADOS Or lErroVendedor = AD_SQL_SEM_DADOS Then
    
    lErro = CF_ECF("Caixa_Transacao_Abrir", lSequencial)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
'    End If
    
    If iDownload = 1 Then
        
        lErro = CF_ECF("Produto_Exclui_EmTrans")
        If lErro <> SUCESSO Then gError 214817
        
        lErro = CF_ECF("TributacaoDocItem_Exclui_EmTrans")
        If lErro <> SUCESSO Then gError 214833
    
        lErro = CF_ECF("ProdutoCodBarras_Exclui_EmTrans")
        If lErro <> SUCESSO Then gError 214866
    
        lErro = CF_ECF("Clientes_Exclui_EmTrans")
        If lErro <> SUCESSO Then gError 214904
    
        lErro = TabelaPrecoItens_Exclui_EmTrans(alComando(3))
        If lErro <> SUCESSO Then gError 214817
    
        lErro = ProdutoDesconto_Exclui_EmTrans(alComando(4))
        If lErro <> SUCESSO Then gError 214817
    
    End If
    
    lErro = CF_ECF("Vendedores_Exclui_EmTrans")
    If lErro <> SUCESSO Then gError 214904
    
     If Len(gobjLojaECF.sFTPURL) > 0 And iDownload = 1 Then
            
        Set objObject = gobjLojaECF
                
        lErro = CF_ECF("Rotina_FTP_Envio_Caixa", objObject, iDownload)
        If lErro <> SUCESSO Then gError 133578
            
    End If
    
    If giDebug = 1 Then MsgBox ("13")

    Set gcolAdmMeioPagto = New Collection
    Set gcolOperadores = New Collection
    Set gobjLojaECF = New ClassLojaECF
    gsNomeCaixa = ""
    giPos = 0
    giTEF = 0
    giBoletoManual = 0
    giTeclado = 0
    giOrcamentoECF = 0
    giCodECF = 0
    giImpressoraECF = 0
    giImpressoraCheque = 0
    giCodModeloECF = 0
    Set gcolECF = New Collection
    Set gcolRedes = New Collection
    Set gcolTeclados = New Collection
    Set gcolTiposMeiosPagtos = New Collection
'    Set gcolCliente = New Collection
'    Set gobjClienteNome = New ClassClienteNome
'    Set gobjClienteCPF = New ClassClienteCPF
'    Set gcolVendedores = New Collection
    Set gaobjProdutosNome = New ClassProdNome
    Set gaobjProdutosCodBarras = New ClassProdCodBarra
    Set gaobjProdutosReferencia = New ClassProdReferencia
    
    lTamanho = 255
    sRetorno = String(lTamanho, 0)
    
    'funcao que verifica se o caixaconfig.ini foi alterado
    lErro = CF_ECF("CaixaConfig_Verifica_Chave")
    If lErro <> SUCESSO Then gError 204245
    
    'Obtém o diretório onde deve ser armazenado o arquivo com dados do backoffice
    Call GetPrivateProfileString(APLICACAO_DADOS, "DirDadosCC", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
    
    'Retira os espaços no final da string
    sRetorno = StringZ(sRetorno)
    
    'Se não encontrou
    If Len(Trim(sRetorno)) = 0 Or sRetorno = CStr(CONSTANTE_ERRO) Then gError 127097
    
    If right(sRetorno, 1) <> "\" Then sRetorno = sRetorno & "\"
    
    sArquivo = sRetorno & giCodEmpresa & "_" & giFilialEmpresa & "_" & NOME_ARQUIVOCC
    
    'Abre o arquivo de retorno
    Open sArquivo For Input As #1
    
    'Busca o próximo registro do arquivo
    Line Input #1, sRegistro
    
    lVersao = StrParaLong(sRegistro)
    
    If lVersao > VERSAO_ECF Then gError 133826
       
    If giDebug = 1 Then MsgBox ("14")
       
    'Até chegar ao fim do arquivo
    Do While Not EOF(1)
        
        'Busca o próximo registro do arquivo
        Line Input #1, sRegistro
        
        If sRegistro <> "" Then
        
            iPosInicio = 1
            
            'Procura o Primeiro Control para saber o tipo do registro
            iPos = InStr(iPosInicio, sRegistro, Chr(vbKeyControl))
            
            sTipo = Mid(sRegistro, iPosInicio, iPos - iPosInicio)
            
            'Seleciona de acordo com o tipo
            Select Case StrParaLong(sTipo)
                    
                Case TIPOREGISTROECF_MEIOSPAGAMENTOS
                    lErro = Carrega_Dados_MeiosPagamentos(sRegistro)
                    If lErro <> SUCESSO Then gError 86246
                    
                Case TIPOREGISTROECF_OPERADORES
                    lErro = Carrega_Dados_Operadores(sRegistro)
                    If lErro <> SUCESSO Then gError 86247
                    
                Case TIPOREGISTROECF_CONFIGURACAOLOJA
                    lErro = Carrega_Dados_ConfiguracaoLoja(sRegistro)
                    If lErro <> SUCESSO Then gError 86248
                    
                Case TIPOREGISTROECF_CAIXAS
                    lErro = Carrega_Dados_Caixas(sRegistro)
                    If lErro <> SUCESSO And lErro <> 112556 Then gError 86249
                    
                Case TIPOREGISTROECF_ECF
                    lErro = Carrega_Dados_ECF(sRegistro)
                    If lErro <> SUCESSO And lErro <> 112560 Then gError 86250
                    
                Case TIPOREGISTROECF_REDES
                    lErro = Carrega_Dados_Redes(sRegistro)
                    If lErro <> SUCESSO Then gError 86251
                    
                Case TIPOREGISTROECF_TECLADOS
                    lErro = Carrega_Dados_Teclados(sRegistro)
                    If lErro <> SUCESSO And lErro <> 112561 Then gError 86252
                             
                Case TIPOREGISTROECF_TIPOSMEIOSPAGAMENTOS
                    lErro = Carrega_Dados_TiposMeioPagtos(sRegistro)
                    If lErro <> SUCESSO Then gError 86252
                    
                Case TIPOREGISTROECF_CLIENTE
                    If iDownload = 1 Or lErroCliente = AD_SQL_SEM_DADOS Then
                        lErro = Carrega_Dados_Cliente(sRegistro)
                        If lErro <> SUCESSO Then gError 86253
                    Else
                        Exit Do
                    End If
                    
                Case TIPOREGISTROECF_PRODUTODESCONTO
                    If iDownload = 1 Or lErroProduto = AD_SQL_SEM_DADOS Then
                        'lErro = Carrega_Dados_ProdutosDesconto(sRegistro, colProdutoDesconto)
                        lErro = Carrega_Dados_ProdutosDesconto(sRegistro, alComando(5))
                        If lErro <> SUCESSO Then gError 86254
                    Else
                        Exit Do
                    End If
                
                Case TIPOREGISTROECF_VENDEDORES
'                    If iDownload = 1 Or lErroVendedor = AD_SQL_SEM_DADOS Then
                    lErro = Carrega_Dados_Vendedores(sRegistro)
                    If lErro <> SUCESSO Then gError 112552
'                    Else
'                        Exit Do
'                    End If
                
                Case TIPOREGISTROECF_PRODUTOSNOME
                    If iDownload = 1 Or lErroProduto = AD_SQL_SEM_DADOS Then
                        'lErro = Carrega_Dados_ProdutosNome1(sRegistro, colTabPreco, colProdutoDesconto, lVersao)
                        lErro = Carrega_Dados_ProdutosNome1(sRegistro, lVersao, alComando(6), alComando(7), alComando(8), alComando(9), alComando(10))
                        If lErro <> SUCESSO Then gError 112554
                    Else
                        Exit Do
                    End If

                Case TIPOREGISTROECF_TABELAPRECOITENS
                    If iDownload = 1 Or lErroProduto = AD_SQL_SEM_DADOS Then
                        'lErro = Carrega_Dados_TabelaPreco(sRegistro, colTabPreco)
                        lErro = Carrega_Dados_TabelaPreco(sRegistro, alComando(11))
                        If lErro <> SUCESSO And lErro <> 112562 Then gError 112555
                    Else
                        Exit Do
                    End If
                
            End Select
        End If
    Loop
    
    If giDebug = 1 Then MsgBox ("15")
    
    Close #1
        
    ' Se o caixa padrão não foi encontrado
    If gsNomeCaixa = "" Then gError 118020
        
        
'    If iDownload = 1 Or lErroProduto = AD_SQL_SEM_DADOS Or lErroCliente = AD_SQL_SEM_DADOS Or lErroVendedor = AD_SQL_SEM_DADOS Then
        'Fechar a Transação
        lErro = CF_ECF("Caixa_Transacao_Fechar", lSequencial)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
'    End If
        
    lErro = CF_ECF("Inicializa_Dados")
    If lErro <> SUCESSO Then gError 99776
   
    
    If giOrcamentoECF <> CAIXA_SO_ORCAMENTO Then

        lErro = CF_ECF("Verifica_Arquivo_Loja", dtUltimoMovto)
        If lErro <> SUCESSO And lErro <> 53 Then gError 109782

    End If
    
    
    
'    '*** Trecho alterado por Luiz Nogueira em 16/04/04 ***
'
'    lTamanho = 255
'    sRetorno = String(lTamanho, 0)
'
'    'Obtém o diretório onde deve ser armazenado o arquivo com dados do backoffice
'    Call GetPrivateProfileString(APLICACAO_DADOS, "DirDadosBack", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
'
'    'Retira os espaços no final da string
'    sRetorno = StringZ(sRetorno)
'
'    'Se não encontrou
'    If Len(Trim(sRetorno)) = 0 Or sRetorno = CStr(CONSTANTE_ERRO) Then gError 127096
'
'    sArquivo = sRetorno & "\" & NOME_ARQUIVOBACK
'    'Abre o arquivo de retorno
'    Open sArquivo For Input As #1
'    '*************************************************************
'
'    'Até chegar ao fim do arquivo
'    Do While Not EOF(1)
'
'        'Busca o próximo registro do arquivo
'        Line Input #1, sRegistro
'
'        If sRegistro <> "" Then
'
'            iPosInicio = 1
'
'            'Procura o Primeiro Control para saber o tipo do registro
'            iPos = InStr(iPosInicio, sRegistro, Chr(vbKeyControl))
'
'            sTipo = Mid(sRegistro, iPosInicio, iPos - iPosInicio)
'
'            'Seleciona de acordo com o tipo
'            Select Case StrParaLong(sTipo)
'
''''                Case TIPOREGISTROECF_PRODUTOSCODBARRAS
''''                    lErro = Carrega_Dados_ProdutosCodBarras(sRegistro, colTabPreco, colProdutoDesconto)
''''                    If lErro <> SUCESSO And lErro <> 112613 Then gError 112551
''''
''''                Case TIPOREGISTROECF_PRODUTOSREFERENCIA
''''                    lErro = Carrega_Dados_ProdutosReferencia(sRegistro, colTabPreco, colProdutoDesconto)
''''                    If lErro <> SUCESSO And lErro <> 112615 Then gError 112553
'
'                Case TIPOREGISTROECF_PRODUTOSNOME
'                    lErro = Carrega_Dados_ProdutosNome(sRegistro, colTabPreco, colProdutoDesconto)
'                    If lErro <> SUCESSO And lErro <> 112614 Then gError 112554
'
'                Case TIPOREGISTROECF_TABELAPRECOITENS
'                    lErro = Carrega_Dados_TabelaPreco(sRegistro, colTabPreco)
'                    If lErro <> SUCESSO And lErro <> 112562 Then gError 112555
'
'            End Select
'        End If
'    Loop
'
'    Close #1
  
    Carrega_Arquivo_FonteDados = SUCESSO
    
    Exit Function
    
Erro_Carrega_Arquivo_FonteDados:

    Carrega_Arquivo_FonteDados = gErr

    Select Case gErr
    
        Case 53, 76
            Call Rotina_ErroECF(vbOKOnly, ARQUIVO_INICIALIZACAO_NAO_EXISTENTE, gErr, sArquivo)
            
        Case 86244 To 86256, 112551 To 112555, 133578, 204245, 214817, 214833, 214866, 214904
        
        Case 118020
            Call Rotina_ErroECF(vbOKOnly, ERRO_CAIXA_NAO_EXISTENTE, gErr)
            
        Case 127096
            Call Rotina_ErroECF(vbOKOnly, ERRO_LOCAL_ARQUIVO_NAO_CONFIGURADO, gErr, NOME_ARQUIVOBACK)
            
        Case 127097
            Call Rotina_ErroECF(vbOKOnly, ERRO_LOCAL_ARQUIVO_NAO_CONFIGURADO, gErr, giCodEmpresa & "_" & giFilialEmpresa & "_" & NOME_ARQUIVOCC)
            
        Case 133826
            Call Rotina_ErroECF(vbOKOnly, "ERRO_VERSAO_ARQUIVO_MAIOR_PROGRAMA", gErr, lVersao, VERSAO_ECF)
            
        Case 214834
            Call Rotina_ErroECF(vbOKOnly, ERRO_ABERTURA_COMANDO, gErr)
            
        Case 214835, 214836
            Call Rotina_ErroECF(vbOKOnly, ERRO_LEITURA_PRODUTOS_0, gErr, sCodProduto)
            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error, 149443)
        
    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
 '   If iDownload = 1 Or lErroProduto = AD_SQL_SEM_DADOS Or lErroCliente = AD_SQL_SEM_DADOS Then
    Call CF_ECF("Caixa_Transacao_Rollback", glTransacaoPAFECF)
'    End If
    
    Exit Function
    
End Function

Private Function Carrega_Dados_MeiosPagamentos(sRegistro As String) As Long

Dim iPos As Integer
Dim objAdmMeioPagto As ClassAdmMeioPagto
Dim objAdmMeioPagtoCondPagto As ClassAdmMeioPagtoCondPagto
Dim objAdmMeioPagtoParc As ClassAdmMeioPagtoParcelas
Dim iIndice As Integer
Dim iIndice1 As Integer
Dim iPosInicio As Integer
Dim iPosFim As Integer
Dim iPosMeio As Integer
Dim iPosCondPagtoInicio As Integer
Dim sAdm As String
    
    Set objAdmMeioPagto = New ClassAdmMeioPagto
            
    'Primeira Posição
    iPosInicio = 1
    
    'Procura o Primeiro Control para saber onde começa a string
    iPosInicio = InStr(iPosInicio, sRegistro, Chr(vbKeyControl)) + 1
    
    'Procura o Primeiro Escape dentro da String
    iPosMeio = (InStr(iPosInicio, sRegistro, Chr(vbKeyEscape)))
    
    'Pega última posição e guarda
    iPosFim = (InStr(iPosInicio, sRegistro, Chr(vbKeyEnd)))
    
    'iPosicao1 Guarda a posição do Segundo Control(referente ao inicio da segunda cond pagto, se houver)
    iPosCondPagtoInicio = InStr(iPosInicio, sRegistro, Chr(vbKeyControl))
    If iPosCondPagtoInicio = 0 Then iPosCondPagtoInicio = iPosFim
    
    iIndice = 0
            
    Do While iPosMeio <> 0
        
       iIndice = iIndice + 1
        
        Select Case iIndice
            
            Case 1: objAdmMeioPagto.iCodigo = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 2: objAdmMeioPagto.sNome = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
            Case 3: objAdmMeioPagto.iRede = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 4: objAdmMeioPagto.iTipoMeioPagto = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 5: objAdmMeioPagto.iAtivo = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
                    Exit Do
        End Select
        
        'Atualiza as Posições
        iPosInicio = iPosMeio + 1
        iPosMeio = (InStr(iPosInicio, sRegistro, Chr(vbKeyEscape)))
        
    Loop
    
    'pula o vbEscape e o vbControl
    iPosInicio = iPosMeio + 2
    
    Do While iPosInicio < iPosFim
        
        'Procura o Primeiro Escape dentro da String
        iPosMeio = (InStr(iPosInicio, sRegistro, Chr(vbKeyEscape)))
        
        'Procura um control dentro da String
        iPosCondPagtoInicio = (InStr(iPosInicio, sRegistro, Chr(vbKeyControl)))
        If iPosCondPagtoInicio = 0 Then iPosCondPagtoInicio = iPosFim
        
        iIndice = 0
        
        Set objAdmMeioPagtoCondPagto = New ClassAdmMeioPagtoCondPagto
        
        objAdmMeioPagtoCondPagto.iAdmMeioPagto = objAdmMeioPagto.iCodigo
        objAdmMeioPagtoCondPagto.sNomeAdmMeioPagto = objAdmMeioPagto.sNome
        
        Do While iPosMeio <> 0
                        
            iIndice = iIndice + 1
        
            Select Case iIndice
            
                Case 1: objAdmMeioPagtoCondPagto.iParcelamento = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
                Case 2: objAdmMeioPagtoCondPagto.sNomeParcelamento = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
                Case 3: objAdmMeioPagtoCondPagto.iNumParcelas = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
                Case 4: objAdmMeioPagtoCondPagto.iParcelasRecebto = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
                Case 5: objAdmMeioPagtoCondPagto.dValorMinimo = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
                Case 6: objAdmMeioPagtoCondPagto.dDesconto = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
                Case 7: objAdmMeioPagtoCondPagto.iJurosParcelamento = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
                Case 8: objAdmMeioPagtoCondPagto.dJuros = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
                Case 9: objAdmMeioPagtoCondPagto.iAtivo = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
                Case 10: objAdmMeioPagtoCondPagto.iPreDatado = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
                        Exit Do
            End Select
       
            'Atualiza as Posições
            iPosInicio = iPosMeio + 1
            iPosMeio = (InStr(iPosInicio, sRegistro, Chr(vbKeyEscape)))
            
        Loop
        
        'pula o vbescape e o vbshift
        iPosInicio = iPosMeio + 2
            
        
        'enquanto nao chegar ao proximo cond pagto ou ao fim
        Do While iPosInicio < iPosCondPagtoInicio
                        
            'Procura o Primeiro Escape dentro da String
            iPosMeio = (InStr(iPosInicio, sRegistro, Chr(vbKeyEscape)))
            
            iIndice1 = 0
            
            Set objAdmMeioPagtoParc = New ClassAdmMeioPagtoParcelas
            
            Do While iPosMeio <> 0
                            
                iIndice1 = iIndice1 + 1
            
                Select Case iIndice1
                
                    Case 1: objAdmMeioPagtoParc.iParcela = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
                    Case 2: objAdmMeioPagtoParc.dPercRecebimento = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
                    Case 3: objAdmMeioPagtoParc.iIntervaloRecebimento = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
                            Exit Do
                    
                End Select
           
                'Atualiza as Posições
                iPosInicio = iPosMeio + 1
                iPosMeio = (InStr(iPosInicio, sRegistro, Chr(vbKeyEscape)))
                
            Loop
            
            objAdmMeioPagtoCondPagto.colParcelas.Add objAdmMeioPagtoParc
            
            'pula o vbescape e o vbshift
            iPosInicio = iPosMeio + 2
            
        Loop
        objAdmMeioPagto.colCondPagtoLoja.Add objAdmMeioPagtoCondPagto
    Loop
    
    'Joga na coleção
    gcolAdmMeioPagto.Add objAdmMeioPagto
    
    Carrega_Dados_MeiosPagamentos = SUCESSO
    
End Function

Private Function Carrega_Dados_Operadores(sRegistro As String) As Long

Dim iPos As Integer
Dim objOperador As ClassOperador
Dim iIndice As Integer
Dim iPosInicio As Integer
Dim iPosFim As Integer
Dim iPosMeio As Integer
        
    Set objOperador = New ClassOperador
            
    'Primeira Posição
    iPosInicio = 1
    
    'Procura o Primeiro Control para saber onde começa a string
    iPosInicio = InStr(iPosInicio, sRegistro, Chr(vbKeyControl)) + 1
    
    'Procura o Primeiro Escape dentro da String
    iPosMeio = (InStr(iPosInicio, sRegistro, Chr(vbKeyEscape)))
    
    'Pega última posição e guarda
    iPosFim = (InStr(iPosInicio, sRegistro, Chr(vbKeyEnd)))
    
    iIndice = 0
            
    Do While iPosMeio <> 0
        
       iIndice = iIndice + 1
        
        Select Case iIndice
            
            Case 1: objOperador.iCodigo = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 2: objOperador.sNome = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
            Case 3: objOperador.sSenha = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
            Case 4: objOperador.iDesconto = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 5: objOperador.iLimiteDesconto = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 6: objOperador.iCodigoVendedor = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 7: objOperador.iGerente = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            
        End Select
        
        'Atualiza as Posições
        iPosInicio = iPosMeio + 1
        iPosMeio = (InStr(iPosInicio, sRegistro, Chr(vbKeyEscape)))
        
        If iPosInicio < iPosFim And iPosMeio = 0 Then iPosMeio = iPosFim
                    
    Loop
    
    'Joga na coleção
    gcolOperadores.Add objOperador
    
    Carrega_Dados_Operadores = SUCESSO
    
End Function

Private Function Carrega_Dados_ConfiguracaoLoja(sRegistro As String) As Long

Dim iPos As Integer
Dim iIndice As Integer
Dim iPosInicio As Integer
Dim iPosFim As Integer
Dim iPosMeio As Integer
Dim iPosColInicio As Integer
Dim iPosColFim As Integer
Dim objAliquotaICMS As New ClassAliquotaICMS
    
    'Primeira Posição
    iPosInicio = 1
    
    'Procura o Primeiro Control para saber onde começa a string
    iPosInicio = InStr(iPosInicio, sRegistro, Chr(vbKeyControl)) + 1
    
    'Procura o Primeiro Escape dentro da String
    iPosMeio = (InStr(iPosInicio, sRegistro, Chr(vbKeyEscape)))
    
    'Pega última posição e guarda
    iPosFim = (InStr(iPosInicio, sRegistro, Chr(vbKeyEnd)))
    
    'iPosicao1 Guarda a posição do Segundo Control(referente a col de Aliquotas)
    iPosColInicio = InStr(iPosInicio + 1, sRegistro, Chr(vbKeyControl))

    iIndice = 0
            
    Do While iPosMeio <> 0
        
       iIndice = iIndice + 1
        
        Select Case iIndice
            
            Case 1: gdtDataAtualizacaoDadosCCC = StrParaDate(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 2: gobjLojaECF.iAtualizacaoECF = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 3: gobjLojaECF.iCupomDescreveFormaPagto = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 4: gobjLojaECF.iGerenteAutoriza = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 5: gobjLojaECF.iHorarioVerao = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 6: gobjLojaECF.iImprimeItemAItem = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 7: gobjLojaECF.iLinhasEntreCupons = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 8: gobjLojaECF.iOperadorIgualVendedor = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 9: gobjLojaECF.iTabelaPreco = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 10: gobjLojaECF.iTelaVendaMP = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 11: gobjLojaECF.lEspacoEntreLinhas = StrParaLong(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 12: gsMensagemCupom = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
            Case 13: gobjLojaECF.sNatOpPadrao = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
            Case 14: gobjLojaECF.sNumLimRO = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
            Case 15: gobjLojaECF.sSimboloMoeda = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
            Case 16: gobjLojaECF.sTruncamentoArredondamento = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
            Case 17: gobjLojaECF.iAbreAposFechamento = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 18: gobjLojaECF.iVendedorObrigatorio = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 19: gobjLojaECF.sFTPURL = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
            Case 20: gobjLojaECF.sFTPUserName = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
            Case 21: gobjLojaECF.sFTPPassword = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
            Case 22: gobjLojaECF.sFTPDiretorio = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
            Case 23: gobjLojaECF.lIntervaloTrans = StrParaLong(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 24: Exit Do
            
        End Select
        
        'Atualiza as Posições
        iPosInicio = iPosMeio + 1
        iPosMeio = (InStr(iPosInicio, sRegistro, Chr(vbKeyEscape)))
        
        If iPosMeio > iPosColInicio Then iPosMeio = iPosColInicio
                    
    Loop
    
    'Procura um control dentro da String
    iPosColInicio = (InStr(iPosInicio, sRegistro, Chr(vbKeyControl)))
    If iPosColInicio = 0 Then iPosColInicio = iPosFim
    
    iPosInicio = iPosColInicio + 2
        
    Do While iPosInicio < iPosFim
        
        'Procura o Primeiro Escape dentro da String
        iPosMeio = (InStr(iPosInicio, sRegistro, Chr(vbKeyEscape)))
                
        'Procura um control dentro da String
        iPosColFim = (InStr(iPosInicio, sRegistro, Chr(vbKeyControl)))
        If iPosColFim = 0 Then iPosColFim = iPosFim
        
        iIndice = 0
        Set objAliquotaICMS = New ClassAliquotaICMS
        
        Do While iPosMeio <> 0
                        
            iIndice = iIndice + 1
        
            Select Case iIndice
            
                Case 1: objAliquotaICMS.dAliquota = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
                Case 2: objAliquotaICMS.iISS = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
                Case 3: objAliquotaICMS.sSigla = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
                Case 4: Exit Do
                
            End Select
       
            'Atualiza as Posições
            iPosInicio = iPosMeio + 1
            iPosMeio = (InStr(iPosInicio, sRegistro, Chr(vbKeyEscape)))
            
            If iPosMeio > iPosColFim Or iPosMeio = 0 Then iPosMeio = iPosColFim
            
        Loop
        gobjLojaECF.colAliquotaICMS.Add objAliquotaICMS
    Loop
    
    Carrega_Dados_ConfiguracaoLoja = SUCESSO
    
End Function

Private Function Carrega_Dados_Caixas(sRegistro As String) As Long

Dim iPos As Integer
Dim objCaixa As ClassCaixa
Dim iIndice As Integer
Dim iPosInicio As Integer
Dim iPosFim As Integer
Dim iPosMeio As Integer
Dim iCod As Integer
Dim bIgual As Boolean
Dim lErro As Long

On Error GoTo Erro_Carrega_dados_Caixa

    Set objCaixa = New ClassCaixa
            
    'Primeira Posição
    iPosInicio = 1
    
    'Procura o Primeiro Control para saber onde começa a string
    iPosInicio = InStr(iPosInicio, sRegistro, Chr(vbKeyControl)) + 1
    
    'Procura o Primeiro Escape dentro da String
    iPosMeio = (InStr(iPosInicio, sRegistro, Chr(vbKeyEscape)))
    
    'Pega última posição e guarda
    iPosFim = (InStr(iPosInicio, sRegistro, Chr(vbKeyEnd)))
    
    iIndice = 0
    bIgual = True
    
    Do While iPosMeio <> 0
        
       iIndice = iIndice + 1
        
        Select Case iIndice
            
            Case 1
                'Verifica se é o código desejado(gicodcaixa)
                iCod = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
                If iCod <> giCodCaixa Then gError 112556
            Case 2: gsNomeCaixa = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
            Case 3: giPos = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 4: giTEF = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 5: giBoletoManual = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 6: giTeclado = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 7: giOrcamentoECF = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            
            If giUsaImpressoraFiscal = 0 Then
                giOrcamentoECF = CAIXA_SO_ORCAMENTO
            End If
            
        End Select
        
                
        
        'Atualiza as Posições
        iPosInicio = iPosMeio + 1
        iPosMeio = (InStr(iPosInicio, sRegistro, Chr(vbKeyEscape)))
        
        If iPosInicio < iPosFim And iPosMeio = 0 Then iPosMeio = iPosFim
                    
    Loop
    
'    If objCaixa.iTipoTEF <> giTipoTEF Then
'        giTipoTEF = objCaixa.iTipoTEF
'        'Atualiza o arquivo
'        Call WritePrivateProfileString(APLICACAO_CAIXA, "TipoTerminal", CStr(giTipoTEF), NOME_ARQUIVO_CAIXA)
'    End If
'    If objCaixa.iTeclado <> giTeclado Then
'        giTeclado = objCaixa.iTeclado
'        'Atualiza o arquivo
'        Call WritePrivateProfileString(APLICACAO_CAIXA, "Teclado", CStr(giTeclado), NOME_ARQUIVO_CAIXA)
'    End If
'    If objCaixa.sPOS <> gsPOS Then
'        gsPOS = objCaixa.sPOS
'        'Atualiza o arquivo
'        Call WritePrivateProfileString(APLICACAO_CAIXA, "POS", gsPOS, NOME_ARQUIVO_CAIXA)
'    End If
'
    If giUsaImpressoraFiscal = 0 Then
        giOrcamentoECF = CAIXA_SO_ORCAMENTO
    End If
    
    Carrega_Dados_Caixas = SUCESSO
    
    Exit Function
    
Erro_Carrega_dados_Caixa:
    
    Carrega_Dados_Caixas = gErr
    
    Select Case gErr
    
        Case 112556
            
            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, 149444)

    End Select
    
    Exit Function
    
End Function

Private Function Carrega_Dados_ECF(sRegistro As String) As Long

Dim iPos As Integer
Dim objECF As ClassECF
Dim iIndice As Integer
Dim iPosInicio As Integer
Dim iPosFim As Integer
Dim iPosMeio As Integer
Dim bIgual As Boolean
Dim iCod As Integer
Dim lErro As Long
Dim iCaixa As Integer

On Error GoTo Erro_Carrega_Dados_ECF
        
    Set objECF = New ClassECF
            
    'Primeira Posição
    iPosInicio = 1
    
    'Procura o Primeiro Control para saber onde começa a string
    iPosInicio = InStr(iPosInicio, sRegistro, Chr(vbKeyControl)) + 1
    
    'Procura o Primeiro Escape dentro da String
    iPosMeio = (InStr(iPosInicio, sRegistro, Chr(vbKeyEscape)))
    
    'Pega última posição e guarda
    iPosFim = (InStr(iPosInicio, sRegistro, Chr(vbKeyEnd)))
    
    iIndice = 0
    
    Do While iPosMeio <> 0
        
       iIndice = iIndice + 1
        
        Select Case iIndice
            
            Case 1: iCaixa = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 2: objECF.iCodigo = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 3: objECF.iImpressoraECF = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 4: objECF.iImpressoraCheque = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 5: objECF.iCodModeloECF = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            
        End Select
        
        'Atualiza as Posições
        iPosInicio = iPosMeio + 1
        iPosMeio = (InStr(iPosInicio, sRegistro, Chr(vbKeyEscape)))
        
        If iPosInicio < iPosFim And iPosMeio = 0 Then iPosMeio = iPosFim
                    
    Loop
    
'    If objECF.iCodigo <> giCodECF Then
'        giCodECF = objECF.iCodigo
'        'Atualiza o arquivo
'        lErro = WritePrivateProfileString(APLICACAO_ECF, "CodECF", CStr(giCodECF), NOME_ARQUIVO_CAIXA)
'        If lErro = 0 Then gError 117533
'
'    End If
'    If objECF.iImpressoraECF <> giImpressoraECF Then
'        giImpressoraECF = objECF.iImpressoraECF
'        'Atualiza o arquivo
'        lErro = WritePrivateProfileString(APLICACAO_ECF, "ImpressoraECF", CStr(giImpressoraECF), NOME_ARQUIVO_CAIXA)
'        If lErro = 0 Then gError 117534
'    End If
'    If objECF.iImpressoraCheque <> giImpressoraCheque Then
'        giImpressoraCheque = objECF.iImpressoraCheque
'        'Atualiza o arquivo
'        lErro = WritePrivateProfileString(APLICACAO_CAIXA, "ImpressoraCheque", CStr(giImpressoraCheque), NOME_ARQUIVO_CAIXA)
'        If lErro = 0 Then gError 117535
'    End If
        
        
    gcolECF.Add objECF
        
    If iCaixa = giCodCaixa Then
        giCodECF = objECF.iCodigo
        giImpressoraECF = objECF.iImpressoraECF
        giImpressoraCheque = objECF.iImpressoraCheque
        giCodModeloECF = objECF.iCodModeloECF
    End If
        
    Carrega_Dados_ECF = SUCESSO
    
    Exit Function
    
Erro_Carrega_Dados_ECF:
    
    Carrega_Dados_ECF = gErr
    
    Select Case gErr
    
        Case 112560
        
        Case 117533
            Call Rotina_ErroECF(vbOKOnly, ERRO_ARQUIVO_NAO_ENCONTRADO1, gErr, APLICACAO_ECF, "CodECF", NOME_ARQUIVO_CAIXA)

        Case 117534
            Call Rotina_ErroECF(vbOKOnly, ERRO_ARQUIVO_NAO_ENCONTRADO1, gErr, APLICACAO_ECF, "ImpressoraECF", NOME_ARQUIVO_CAIXA)

        Case 117535
            Call Rotina_ErroECF(vbOKOnly, ERRO_ARQUIVO_NAO_ENCONTRADO1, gErr, APLICACAO_CAIXA, "ImpressoraCheque", NOME_ARQUIVO_CAIXA)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, 149445)

    End Select
    
    Exit Function
    
End Function

Private Function Carrega_Dados_TabelaPreco(sRegistro As String, ByVal lComando As Long) As Long

Dim iPos As Integer
Dim iIndice As Integer
Dim iPosInicio As Integer
Dim iPosFim As Integer
Dim iPosMeio As Integer
Dim objTabPreco As New ClassTabelaPrecoItem
Dim objTabPreco1 As New ClassTabelaPrecoItem
Dim bAchou As Boolean
Dim lErro As Long

On Error GoTo Erro_Carrega_Dados_TabelaPreco

    Set objTabPreco = New ClassTabelaPrecoItem
            
    'Primeira Posição
    iPosInicio = 1
    
    'Procura o Primeiro Control para saber onde começa a string
    iPosInicio = InStr(iPosInicio, sRegistro, Chr(vbKeyControl)) + 1
    
    'Procura o Primeiro Escape dentro da String
    iPosMeio = (InStr(iPosInicio, sRegistro, Chr(vbKeyEscape)))
    
    'Pega última posição e guarda
    iPosFim = (InStr(iPosInicio, sRegistro, Chr(vbKeyEnd)))
    
    iIndice = 0
    
    Do While iPosMeio <> 0
        
       iIndice = iIndice + 1
        
        Select Case iIndice
            
            Case 1
                objTabPreco.iCodTabela = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
                If gobjLojaECF.iTabelaPreco <> objTabPreco.iCodTabela Then gError 112562
                
            Case 2: objTabPreco.sCodProduto = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
                
            Case 3: objTabPreco.dtDataVigencia = StrParaDate(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 4: objTabPreco.dPreco = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 5: Exit Do
            
        End Select
        
        'Atualiza as Posições
        iPosInicio = iPosMeio + 1
        iPosMeio = (InStr(iPosInicio, sRegistro, Chr(vbKeyEscape)))
        
        If iPosInicio < iPosFim And iPosMeio = 0 Then iPosMeio = iPosFim
                    
    Loop
         
'    bAchou = False
'    If objTabPreco.dtDataVigencia <= Date Then
'
'        For iIndice = 1 To colTabPreco.Count
'            Set objTabPreco1 = colTabPreco.Item(iIndice)
'            If UCase(objTabPreco1.sCodProduto) = UCase(objTabPreco.sCodProduto) Then
'                If objTabPreco1.dtDataVigencia < objTabPreco.dtDataVigencia Then
'                    colTabPreco.Remove (iIndice)
'                    colTabPreco.Add objTabPreco
'                    bAchou = True
'                End If
'            End If
'        Next
'        If Not (bAchou) Then colTabPreco.Add objTabPreco
'
'    End If

    lErro = Comando_Executar(lComando, "INSERT INTO TabelaPrecoItens (CodProduto, DataVigencia, Preco) VALUES (?,?,?)", objTabPreco.sCodProduto, objTabPreco.dtDataVigencia, objTabPreco.dPreco)
    If lErro <> AD_SQL_SUCESSO Then gError 216209
    
    Carrega_Dados_TabelaPreco = SUCESSO
    
    Exit Function
    
Erro_Carrega_Dados_TabelaPreco:

    Carrega_Dados_TabelaPreco = gErr

    Select Case gErr
    
        Case 112562
            
        Case 216209
            Call Rotina_ErroECF(vbOKOnly, ERRO_INCLUSAO_TABELAPRECOITENS, gErr)
            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error, 149446)
        
    End Select
    
End Function


Private Function Carrega_Dados_Redes(sRegistro As String) As Long

Dim iPos As Integer
Dim objRede As ClassRede
Dim iIndice As Integer
Dim iPosInicio As Integer
Dim iPosFim As Integer
Dim iPosMeio As Integer
    
    Set objRede = New ClassRede
            
    'Primeira Posição
    iPosInicio = 1
    
    'Procura o Primeiro Control para saber onde começa a string
    iPosInicio = InStr(iPosInicio, sRegistro, Chr(vbKeyControl)) + 1
    
    'Procura o Primeiro Escape dentro da String
    iPosMeio = (InStr(iPosInicio, sRegistro, Chr(vbKeyEscape)))
    
    'Pega última posição e guarda
    iPosFim = (InStr(iPosInicio, sRegistro, Chr(vbKeyEnd)))
    
    iIndice = 0
            
    Do While iPosMeio <> 0
        
       iIndice = iIndice + 1
        
        Select Case iIndice
            
            Case 1: objRede.iCodigo = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 2: objRede.sNome = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
            
            'só na VERSAO_ECF_303 em diante
            Case 3: objRede.sCgc = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
                        
        End Select
        
        'Atualiza as Posições
        iPosInicio = iPosMeio + 1
        iPosMeio = (InStr(iPosInicio, sRegistro, Chr(vbKeyEscape)))
        
        If iPosInicio < iPosFim And iPosMeio = 0 Then iPosMeio = iPosFim
                    
    Loop
    
    'Joga na coleção
    gcolRedes.Add objRede
    
    Carrega_Dados_Redes = SUCESSO
    
End Function

Private Function Carrega_Dados_Teclados(sRegistro As String) As Long

Dim iPos As Integer
Dim iIndice As Integer
Dim iPosInicio As Integer
Dim iPosFim As Integer
Dim iPosMeio As Integer
Dim iPosColInicio As Integer
Dim iPosColFim As Integer
Dim objTecladoProduto As ClassTecladoProduto
Dim objTecladoProdutoItem As ClassTecladoProdutoItem
Dim lErro As Long

On Error GoTo Erro_Carrega_Dados_Teclados

    Set objTecladoProduto = New ClassTecladoProduto
    
    'Primeira Posição
    iPosInicio = 1
    
    'Procura o Primeiro Control para saber onde começa a string
    iPosInicio = InStr(iPosInicio, sRegistro, Chr(vbKeyControl)) + 1
    
    'Procura o Primeiro Escape dentro da String
    iPosMeio = (InStr(iPosInicio, sRegistro, Chr(vbKeyEscape)))
    
    'Pega última posição e guarda
    iPosFim = (InStr(iPosInicio, sRegistro, Chr(vbKeyEnd)))
    
    'Guarda a posição do Segundo Control(referente a col de codBarras)
    iPosColInicio = InStr(iPosInicio + 1, sRegistro, Chr(vbKeyControl))

    iIndice = 0
            
    Do While iPosMeio <> 0
        
       iIndice = iIndice + 1
        
        Select Case iIndice
            
            Case 1
                objTecladoProduto.iTeclado = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
                If objTecladoProduto.iTeclado <> giTeclado Then gError 112561
            Case 2: objTecladoProduto.sDescricao = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
            Case 3: objTecladoProduto.iPadrao = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 4: objTecladoProduto.iCodigo = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 5: Exit Do
            
        End Select
        
        'Atualiza as Posições
        iPosInicio = iPosMeio + 1
        iPosMeio = (InStr(iPosInicio, sRegistro, Chr(vbKeyEscape)))
        
    Loop
    
    Do While iPosInicio < iPosFim
    
        iPosInicio = iPosColInicio + 1
        
        'Procura o Primeiro Escape dentro da String
        iPosMeio = (InStr(iPosInicio, sRegistro, Chr(vbKeyEscape)))
        
        'Procura um control dentro da String
        iPosColInicio = (InStr(iPosInicio, sRegistro, Chr(vbKeyControl)))
        If iPosColInicio = 0 Then iPosColInicio = iPosFim
        
        iIndice = 0
        
        Set objTecladoProdutoItem = New ClassTecladoProdutoItem
        
        Do While iPosMeio <> 0
                        
            iIndice = iIndice + 1
        
            Select Case iIndice
            
                Case 1: objTecladoProdutoItem.sProduto = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
                Case 2: objTecladoProdutoItem.sTitulo = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
                Case 3: objTecladoProdutoItem.lColor = StrParaLong(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
                Case 4: objTecladoProdutoItem.iTecla = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
                Case 5: objTecladoProdutoItem.iIndice = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
                Case 6: objTecladoProdutoItem.sReferencia = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
                Case 7: objTecladoProdutoItem.sCodigoBarras = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
                Case 8: objTecladoProdutoItem.sArvoreKey = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
                Case 9: Exit Do
                
            End Select
       
            'Atualiza as Posições
            iPosInicio = iPosMeio + 1
            iPosMeio = (InStr(iPosInicio, sRegistro, Chr(vbKeyEscape)))
            
            If iPosMeio > iPosColInicio Or iPosMeio = 0 Then iPosMeio = iPosColInicio
            
        Loop
        If objTecladoProdutoItem.sReferencia <> "" Or objTecladoProdutoItem.sCodigoBarras <> "" Or objTecladoProdutoItem.sProduto = "" Then
            objTecladoProduto.colTecladoProdutoItem.Add objTecladoProdutoItem, objTecladoProdutoItem.sArvoreKey
        End If
    Loop
    
    gcolTeclados.Add objTecladoProduto
    
    Carrega_Dados_Teclados = SUCESSO
    
    Exit Function
    
Erro_Carrega_Dados_Teclados:
    
    Carrega_Dados_Teclados = gErr
    
    Select Case gErr
    
        Case 112561
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, 149447)

    End Select
    
    Exit Function
    
End Function

Private Function Carrega_Dados_TiposMeioPagtos(sRegistro As String) As Long

Dim iPos As Integer
Dim objTipoMeioPagto As ClassTMPLoja
Dim iIndice As Integer
Dim iPosInicio As Integer
Dim iPosFim As Integer
Dim iPosMeio As Integer

On Error GoTo Erro_Carrega_Dados_TiposMeioPagtos
    
    Set objTipoMeioPagto = New ClassTMPLoja
            
    'Primeira Posição
    iPosInicio = 1
    
    'Procura o Primeiro Control para saber onde começa a string
    iPosInicio = InStr(iPosInicio, sRegistro, Chr(vbKeyControl)) + 1
    
    'Procura o Primeiro Escape dentro da String
    iPosMeio = (InStr(iPosInicio, sRegistro, Chr(vbKeyEscape)))
    
    'Pega última posição e guarda
    iPosFim = (InStr(iPosInicio, sRegistro, Chr(vbKeyEnd)))
    
    iIndice = 0
            
    Do While iPosMeio <> 0
        
       iIndice = iIndice + 1
        
        Select Case iIndice
            
            Case 1: objTipoMeioPagto.iTipo = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 2: objTipoMeioPagto.sDescricao = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
            Case 3: objTipoMeioPagto.iTransferencia = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 4: objTipoMeioPagto.iIndice = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 5: objTipoMeioPagto.iCodigoCFe = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 6: Exit Do
            
        End Select
        
        
        
        'Atualiza as Posições
        iPosInicio = iPosMeio + 1
        iPosMeio = (InStr(iPosInicio, sRegistro, Chr(vbKeyEscape)))
        
        If iPosInicio >= iPosFim Then Exit Do
        
        If iPosMeio = 0 Then iPosMeio = iPosFim
                    
    Loop
    
    'Joga na coleção
    gcolTiposMeiosPagtos.Add objTipoMeioPagto
    
    Carrega_Dados_TiposMeioPagtos = SUCESSO
    
    Exit Function

Erro_Carrega_Dados_TiposMeioPagtos:

    Carrega_Dados_TiposMeioPagtos = gErr
    
    Select Case Err
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_CARREGA_DADOS_TIPOSMEIOPAGTO, gErr)
    
    End Select
    
    Exit Function
    
End Function

'Private Function Carrega_Dados_ProdutosNome(sRegistro As String, colTabPreco As Collection, colProdutoDesconto As Collection) As Long
'
'Dim iPOS As Integer
'Dim objProduto As ClassProduto
'Dim objProduto2 As ClassProduto
'Dim iIndice As Integer
'Dim iPosInicio As Integer
'Dim iPosFim As Integer
'Dim iPosMeio As Integer
'Dim iPosColInicio As Integer
'Dim sProduto As String
'Dim objTabPreco As New ClassTabelaPrecoItem
'Dim colTab As New Collection
'Dim dValor As Double
'Dim lErro As Long
'Dim bAchou As Boolean
'Dim sCodBarras As String
'
'On Error GoTo Erro_Carrega_Dados_ProdutosNome
'
'    Set objProduto = New ClassProduto
'
'    'Primeira Posição
'    iPosInicio = 1
'
'    'Procura o Primeiro Control para saber onde começa a string
'    iPosInicio = InStr(iPosInicio, sRegistro, Chr(vbKeyControl)) + 1
'
'    'Procura o Primeiro Escape dentro da String
'    iPosMeio = (InStr(iPosInicio, sRegistro, Chr(vbKeyEscape)))
'
'    'Pega última posição e guarda
'    iPosFim = (InStr(iPosInicio, sRegistro, Chr(vbKeyEnd)))
'
'    iIndice = 0
'
'    Do While iPosMeio <> 0
'
'       iIndice = iIndice + 1
'
'        Select Case iIndice
'
'
'            Case 1: objProduto.sCodigo = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
'            Case 2: objProduto.sNomeReduzido = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
'            Case 3: objProduto.sSiglaUMVenda = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
'            Case 4: objProduto.sReferencia = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
'            Case 5: objProduto.sFigura = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
'            Case 6: objProduto.sSituacaoTribECF = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
'            Case 7: objProduto.sICMSAliquota = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
'            Case 8: objProduto.sDescricao = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
'            Case 9: objProduto.iClasseUM = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'            Case 10: objProduto.sSiglaUMEstoque = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
'            Case 11: objProduto.sSiglaUMCompra = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
'            Case 12: objProduto.iAtivo = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'            Case 13: objProduto.dLimiteDesconto = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'            Case 14: objProduto.iUsaBalanca = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'            Case 15: objProduto.iCompras = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                     objProduto.sTruncamento = gobjLojaECF.sTruncamentoArredondamento
'            Case 16: objProduto.dQuantEstLoja = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'
'
'
'            Case 17: objProduto.objTributacaoDocItem.dQuantidade = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'            Case 18: objProduto.objTributacaoDocItem.dPrecoUnitario = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'            Case 19: objProduto.objTributacaoDocItem.sNaturezaOp = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
'            Case 20: objProduto.objTributacaoDocItem.iTipoTributacao = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'            Case 21: objProduto.objTributacaoDocItem.iICMSTipo = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'            Case 22: objProduto.objTributacaoDocItem.dICMSBase = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'            Case 23: objProduto.objTributacaoDocItem.dICMSPercRedBase = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'            Case 24: objProduto.objTributacaoDocItem.dICMSAliquota = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'            Case 25: objProduto.objTributacaoDocItem.dICMSValor = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'
'            Case 26: objProduto.objTributacaoDocItem.iPISTipo = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'            Case 27: objProduto.objTributacaoDocItem.iPISTipoCalculo = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'            Case 28: objProduto.objTributacaoDocItem.dPISBase = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'            Case 29: objProduto.objTributacaoDocItem.dPISAliquota = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'            Case 30: objProduto.objTributacaoDocItem.dPISAliquotaValor = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'            Case 31: objProduto.objTributacaoDocItem.dPISQtde = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'            Case 32: objProduto.objTributacaoDocItem.dPISValor = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'
'            Case 33: objProduto.objTributacaoDocItem.iPISSTTipoCalculo = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'            Case 34: objProduto.objTributacaoDocItem.dPISSTBase = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'            Case 35: objProduto.objTributacaoDocItem.dPISSTAliquota = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'            Case 36: objProduto.objTributacaoDocItem.dPISSTAliquotaValor = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'            Case 37: objProduto.objTributacaoDocItem.dPISSTQtde = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'            Case 38: objProduto.objTributacaoDocItem.dPISSTValor = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'
'            Case 39: objProduto.objTributacaoDocItem.iCOFINSTipo = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'            Case 40: objProduto.objTributacaoDocItem.iCOFINSTipoCalculo = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'            Case 41: objProduto.objTributacaoDocItem.dCOFINSBase = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'            Case 42: objProduto.objTributacaoDocItem.dCOFINSAliquota = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'            Case 43: objProduto.objTributacaoDocItem.dCOFINSAliquotaValor = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'            Case 44: objProduto.objTributacaoDocItem.dCOFINSQtde = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'            Case 45: objProduto.objTributacaoDocItem.dCOFINSValor = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'
'            Case 46: objProduto.objTributacaoDocItem.iCOFINSSTTipoCalculo = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'            Case 47: objProduto.objTributacaoDocItem.dCOFINSSTBase = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'            Case 48: objProduto.objTributacaoDocItem.dCOFINSSTAliquota = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'            Case 49: objProduto.objTributacaoDocItem.dCOFINSSTAliquotaValor = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'            Case 50: objProduto.objTributacaoDocItem.dCOFINSSTQtde = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'            Case 51: objProduto.objTributacaoDocItem.dCOFINSSTValor = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'
'            Case 52: objProduto.objTributacaoDocItem.sCST = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
'            Case 53: objProduto.objTributacaoDocItem.sISSQN = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
'            Case 54: objProduto.objTributacaoDocItem.dISSBase = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'            Case 55: objProduto.objTributacaoDocItem.dISSAliquota = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'            Case 56: objProduto.objTributacaoDocItem.sISSCidadeIBGE = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
'            Case 57: objProduto.objTributacaoDocItem.dISSValor = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'
'            Case 58: objProduto.objTributacaoDocItem.iICMSBaseModalidade = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'            Case 59: objProduto.objTributacaoDocItem.iOrigemMercadoria = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'            Case 60: objProduto.objTributacaoDocItem.sGenero = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
'            Case 61: objProduto.objTributacaoDocItem.dQtdTrib = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'            Case 62: objProduto.objTributacaoDocItem.sUMTrib = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
'            Case 63: objProduto.objTributacaoDocItem.dValorUnitTrib = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'
'            Case 64: objProduto.objTributacaoDocItem.sISSCST = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
'            Case 65: objProduto.objTributacaoDocItem.dICMSValorIsento = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'            Case 66: objProduto.objTributacaoDocItem.iICMSMotivo = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'            Case 67: objProduto.objTributacaoDocItem.iICMSSimplesTipo = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'            Case 68: objProduto.objTributacaoDocItem.iISSTipo = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'
'            Case 69: objProduto.objTributacaoDocItem.iRegimeTributario = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'            Case 70: objProduto.objTributacaoDocItem.sCSOSN = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
'            Case 71: objProduto.objTributacaoDocItem.dTotTrib = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'            Case 72: objProduto.objTributacaoDocItem.iTotTribTipo = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'            Case 73: objProduto.objTributacaoDocItem.sIPICodProduto = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
'
'            Case 74 To (74 + NUM_MAX_CODBARRAS_PRODUTO)
'                If objProduto.sCodigoBarras = "" Then
'                    objProduto.sCodigoBarras = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
'                    objProduto.objTributacaoDocItem.sEAN = objProduto.sCodigoBarras
'                    objProduto.objTributacaoDocItem.sEANTrib = objProduto.sCodigoBarras
'                Else
'                    objProduto.colCodBarras.Add objProduto.sCodigoBarras
'                End If
'
'        End Select
'
'        'Atualiza as Posições
'        iPosInicio = iPosMeio + 1
'        iPosMeio = (InStr(iPosInicio, sRegistro, Chr(vbKeyEscape)))
'
'        If iPosInicio <= iPosFim And iPosMeio = 0 Then iPosMeio = iPosFim
'
'    Loop
'
'    bAchou = False
'
'    For Each objTabPreco In colTabPreco
'        If UCase(objTabPreco.sCodProduto) = UCase(objProduto.sCodigo) Then
'            objProduto.dPrecoLoja = objTabPreco.dPreco
'            bAchou = True
'            Exit For
'        End If
'    Next
'
'    If bAchou Then
'        For Each objProduto2 In colProdutoDesconto
'            If objProduto2.sCodigo = objProduto.sCodigo Then
'                objProduto.dPercentMenosReceb = objProduto2.dPercentMenosReceb
'                objProduto.dDescontoValor = objProduto2.dDescontoValor
'                Exit For
'            End If
'        Next
'
'        gaobjProdutosNome.Add objProduto
'        gaobjProdutosCodigo.Add objProduto
'
'        For iIndice = 1 To objProduto.colCodBarras.Count
'            objProduto.sCodigoBarras = objProduto.colCodBarras(iIndice)
'            gaobjProdutosCodBarras.Add objProduto
'        Next
'
'        If Len(Trim(objProduto.sReferencia)) > 0 Then
'            gaobjProdutosReferencia.Add objProduto
'        End If
'
'
'    End If
'
'    Carrega_Dados_ProdutosNome = SUCESSO
'
'    Exit Function
'
'Erro_Carrega_Dados_ProdutosNome:
'
'    Carrega_Dados_ProdutosNome = gErr
'
'    Select Case gErr
'
'        Case Else
'            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, 149448)
'
'    End Select
'
'    Exit Function
'
'End Function

Private Function Carrega_Dados_ProdutosDesconto(sRegistro As String, ByVal lComando As Long) As Long

Dim lErro As Long
Dim iPos As Integer
Dim objProduto As ClassProduto
Dim objProduto2 As ClassProduto
Dim iIndice As Integer
Dim iPosInicio As Integer
Dim iPosFim As Integer
Dim iPosMeio As Integer
Dim iPosColInicio As Integer
Dim sProduto As String
Dim objTabPreco As New ClassTabelaPrecoItem

On Error GoTo Erro_Carrega_Dados_ProdutosDesconto

    Set objProduto = New ClassProduto
            
    'Primeira Posição
    iPosInicio = 1
            
    'Procura o Primeiro Control para saber onde começa a string
    iPosInicio = InStr(iPosInicio, sRegistro, Chr(vbKeyControl)) + 1
    
    'Procura o Primeiro Escape dentro da String
    iPosMeio = (InStr(iPosInicio, sRegistro, Chr(vbKeyEscape)))
    
    'Pega última posição e guarda
    iPosFim = (InStr(iPosInicio, sRegistro, Chr(vbKeyEnd)))
    
    iIndice = 0
    
    Do While iPosMeio <> 0
        
       iIndice = iIndice + 1
        
        Select Case iIndice
            
            Case 1: objProduto.sCodigo = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
            Case 2: objProduto.sNomeReduzido = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
            Case 3: objProduto.dPercentMenosReceb = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 4: objProduto.dDescontoValor = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            
        End Select
        
        'Atualiza as Posições
        iPosInicio = iPosMeio + 1
        iPosMeio = (InStr(iPosInicio, sRegistro, Chr(vbKeyEscape)))
        
        If iPosInicio <= iPosFim And iPosMeio = 0 Then iPosMeio = iPosFim
                    
    Loop
    
    'colProdutosDesconto.Add objProduto
    
    lErro = Comando_Executar(lComando, "INSERT INTO ProdutoDesconto (CodProduto, PercentMenosReceb, DescontoValor) VALUES (?,?,?)", objProduto.sCodigo, objProduto.dPercentMenosReceb, objProduto.dDescontoValor)
    If lErro <> AD_SQL_SUCESSO Then gError 216215
        
    Carrega_Dados_ProdutosDesconto = SUCESSO
    
    Exit Function
    
Erro_Carrega_Dados_ProdutosDesconto:

    Carrega_Dados_ProdutosDesconto = gErr

    Select Case gErr
             
        Case 216215
            Call Rotina_ErroECF(vbOKOnly, ERRO_INCLUSAO_TABELAPRECOITENS, gErr)
            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error, 216216)
        
    End Select
    
End Function

Private Function Carrega_Dados_Vendedores(sRegistro As String) As Long

Dim iPos As Integer
Dim objVendedor As ClassVendedor
Dim iIndice As Integer
Dim iPosInicio As Integer
Dim iPosFim As Integer
Dim iPosMeio As Integer
Dim lErro As Long
    
    Set objVendedor = New ClassVendedor
            
    'Primeira Posição
    iPosInicio = 1
    
    'Procura o Primeiro Control para saber onde começa a string
    iPosInicio = InStr(iPosInicio, sRegistro, Chr(vbKeyControl)) + 1
    
    'Procura o Primeiro Escape dentro da String
    iPosMeio = (InStr(iPosInicio, sRegistro, Chr(vbKeyEscape)))
    
    'Pega última posição e guarda
    iPosFim = (InStr(iPosInicio, sRegistro, Chr(vbKeyEnd)))
    
    iIndice = 0
            
    Do While iPosMeio <> 0
        
       iIndice = iIndice + 1
        
        Select Case iIndice
            
            Case 1: objVendedor.iCodigo = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 2: objVendedor.sNomeReduzido = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
            Case 3: Exit Do
            
        End Select
        
        'Atualiza as Posições
        iPosInicio = iPosMeio + 1
        iPosMeio = (InStr(iPosInicio, sRegistro, Chr(vbKeyEscape)))
        
        If iPosInicio < iPosFim And iPosMeio = 0 Then iPosMeio = iPosFim
                    
    Loop
    
    lErro = CF_ECF("Vendedores_Grava_EmTrans", objVendedor)
    If lErro <> SUCESSO Then gError 214939
    
    
    'Joga na coleção
'    gcolVendedores.Add objVendedor
    
    Carrega_Dados_Vendedores = SUCESSO
    
End Function

Private Function Carrega_Dados_Cliente(sRegistro As String) As Long

Dim iPos As Integer
Dim objCliente As ClassCliente
Dim iIndice As Integer
Dim iPosInicio As Integer
Dim iPosFim As Integer
Dim iPosMeio As Integer
Dim lErro As Long
    
    Set objCliente = New ClassCliente
            
    'Primeira Posição
    iPosInicio = 1
    
    'Procura o Primeiro Control para saber onde começa a string
    iPosInicio = InStr(iPosInicio, sRegistro, Chr(vbKeyControl)) + 1
    
    'Procura o Primeiro Escape dentro da String
    iPosMeio = (InStr(iPosInicio, sRegistro, Chr(vbKeyEscape)))
    
    'Pega última posição e guarda
    iPosFim = (InStr(iPosInicio, sRegistro, Chr(vbKeyEnd)))
    
    iIndice = 0
            
    Do While iPosMeio <> 0
        
       iIndice = iIndice + 1
        
        Select Case iIndice
            
            Case 1: objCliente.lCodigo = StrParaLong(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 2: objCliente.sNomeReduzido = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
            Case 3: objCliente.sCgc = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
            Case 4: Exit Do
            
        End Select
        
        'Atualiza as Posições
        iPosInicio = iPosMeio + 1
        iPosMeio = (InStr(iPosInicio, sRegistro, Chr(vbKeyEscape)))
        
        If iPosInicio < iPosFim And iPosMeio = 0 Then iPosMeio = iPosFim
                    
    Loop
    
    'Joga na coleção
'    gcolCliente.Add objCliente
'
'    gobjClienteNome.Add objCliente
'
'    gobjClienteCPF.Add objCliente
    
    lErro = CF_ECF("Clientes_Grava_EmTrans", objCliente)
    If lErro <> SUCESSO Then gError 214816
    
    
    Carrega_Dados_Cliente = SUCESSO
    
End Function

Public Function Trata_MovCaixa(objMovCaixa As ClassMovimentoCaixa, colMeiosPag As Collection) As Long

Dim objMovCaixa1 As ClassMovimentoCaixa
Dim iAchou As Integer

On Error GoTo Erro_Trata_MovCaixa

    If objMovCaixa.iTipoCartao = TIPO_TEF Then
        
        For Each objMovCaixa1 In colMeiosPag
            If objMovCaixa1.iTipo = TIPOMEIOPAGTOLOJA_TEF Then
                objMovCaixa1.dValor = objMovCaixa1.dValor + objMovCaixa.dValor
                iAchou = 1
                Exit For
            End If
        Next
                
        If iAchou = 0 Then
            Set objMovCaixa1 = New ClassMovimentoCaixa
            objMovCaixa1.iTipo = TIPOMEIOPAGTOLOJA_TEF
            objMovCaixa1.dValor = objMovCaixa.dValor
            colMeiosPag.Add objMovCaixa1
        End If
        
                
    ElseIf objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_CHEQUE And objMovCaixa.iTipoCartao = TIPOMEIOPAGTOLOJA_TEFCHQ Then
    
            Set objMovCaixa1 = New ClassMovimentoCaixa
            objMovCaixa1.iTipo = TIPOMEIOPAGTOLOJA_TEFCHQ
            objMovCaixa1.dValor = objMovCaixa.dValor
            objMovCaixa1.iIndiceImpChq = objMovCaixa.iIndiceImpChq
            colMeiosPag.Add objMovCaixa1

    Else

        'Seleciona o tipo
        Select Case objMovCaixa.iTipo
        
            Case MOVIMENTOCAIXA_RECEB_CARNE:
                    
                    For Each objMovCaixa1 In colMeiosPag
                        If objMovCaixa1.iTipo = TIPOMEIOPAGTOLOJA_CARNE Then
                            objMovCaixa1.dValor = objMovCaixa1.dValor + objMovCaixa.dValor
                            iAchou = 1
                            Exit For
                        End If
                    Next
                    
                    If iAchou = 0 Then
                        Set objMovCaixa1 = New ClassMovimentoCaixa
                        objMovCaixa1.iTipo = TIPOMEIOPAGTOLOJA_CARNE
                        objMovCaixa1.dValor = objMovCaixa.dValor
                        colMeiosPag.Add objMovCaixa1
                    End If
                    
            Case MOVIMENTOCAIXA_RECEB_CARTAOCREDITO, MOVIMENTOCAIXA_CANC_RECEB_CARTAOCREDITO:
                    
                    For Each objMovCaixa1 In colMeiosPag
                        If objMovCaixa1.iTipo = TIPOMEIOPAGTOLOJA_CARTAO_CREDITO Then
                            objMovCaixa1.dValor = objMovCaixa1.dValor + objMovCaixa.dValor
                            iAchou = 1
                            Exit For
                        End If
                    Next
                    
                    If iAchou = 0 Then
                        Set objMovCaixa1 = New ClassMovimentoCaixa
                        objMovCaixa1.iTipo = TIPOMEIOPAGTOLOJA_CARTAO_CREDITO
                        objMovCaixa1.dValor = objMovCaixa.dValor
                        colMeiosPag.Add objMovCaixa1
                    End If
                
            Case MOVIMENTOCAIXA_RECEB_CARTAODEBITO, MOVIMENTOCAIXA_CANC_RECEB_CARTAODEBITO:
                
                    For Each objMovCaixa1 In colMeiosPag
                        If objMovCaixa1.iTipo = TIPOMEIOPAGTOLOJA_CARTAO_DEBITO Then
                            objMovCaixa1.dValor = objMovCaixa1.dValor + objMovCaixa.dValor
                            iAchou = 1
                            Exit For
                        End If
                    Next
                    
                    If iAchou = 0 Then
                        Set objMovCaixa1 = New ClassMovimentoCaixa
                        objMovCaixa1.iTipo = TIPOMEIOPAGTOLOJA_CARTAO_DEBITO
                        objMovCaixa1.dValor = objMovCaixa.dValor
                        colMeiosPag.Add objMovCaixa1
                    End If
                
                
                
            Case MOVIMENTOCAIXA_RECEB_CHEQUE:
            
                    For Each objMovCaixa1 In colMeiosPag
                        If objMovCaixa1.iTipo = TIPOMEIOPAGTOLOJA_CHEQUE Then
                            objMovCaixa1.dValor = objMovCaixa1.dValor + objMovCaixa.dValor
                            iAchou = 1
                            Exit For
                        End If
                    Next
                    
                    If iAchou = 0 Then
                        Set objMovCaixa1 = New ClassMovimentoCaixa
                        objMovCaixa1.iTipo = TIPOMEIOPAGTOLOJA_CHEQUE
                        objMovCaixa1.dValor = objMovCaixa.dValor
                        colMeiosPag.Add objMovCaixa1
                    End If
            
            Case MOVIMENTOCAIXA_RECEB_DINHEIRO:
                
                    For Each objMovCaixa1 In colMeiosPag
                        If objMovCaixa1.iTipo = TIPOMEIOPAGTOLOJA_DINHEIRO Then
                            objMovCaixa1.dValor = objMovCaixa1.dValor + objMovCaixa.dValor
                            iAchou = 1
                            Exit For
                        End If
                    Next
                    
                    If iAchou = 0 Then
                        Set objMovCaixa1 = New ClassMovimentoCaixa
                        objMovCaixa1.iTipo = TIPOMEIOPAGTOLOJA_DINHEIRO
                        objMovCaixa1.dValor = objMovCaixa.dValor
                        colMeiosPag.Add objMovCaixa1
                    End If
                
                
            Case MOVIMENTOCAIXA_RECEB_OUTROS:
                
                    For Each objMovCaixa1 In colMeiosPag
                        If objMovCaixa1.iTipo = TIPOMEIOPAGTOLOJA_OUTROS Then
                            objMovCaixa1.dValor = objMovCaixa1.dValor + objMovCaixa.dValor
                            iAchou = 1
                            Exit For
                        End If
                    Next
                    
                    If iAchou = 0 Then
                        Set objMovCaixa1 = New ClassMovimentoCaixa
                        objMovCaixa1.iTipo = TIPOMEIOPAGTOLOJA_OUTROS
                        objMovCaixa1.dValor = objMovCaixa.dValor
                        colMeiosPag.Add objMovCaixa1
                    End If
                
                
            Case MOVIMENTOCAIXA_RECEB_TROCA:
                
                    For Each objMovCaixa1 In colMeiosPag
                        If objMovCaixa1.iTipo = TIPOMEIOPAGTOLOJA_TROCA Then
                            objMovCaixa1.dValor = objMovCaixa1.dValor + objMovCaixa.dValor
                            iAchou = 1
                            Exit For
                        End If
                    Next
                    
                    If iAchou = 0 Then
                        Set objMovCaixa1 = New ClassMovimentoCaixa
                        objMovCaixa1.iTipo = TIPOMEIOPAGTOLOJA_TROCA
                        objMovCaixa1.dValor = objMovCaixa.dValor
                        colMeiosPag.Add objMovCaixa1
                    End If
                
            Case MOVIMENTOCAIXA_RECEB_VALETICKET:
                
                    For Each objMovCaixa1 In colMeiosPag
                        If objMovCaixa1.iTipo = TIPOMEIOPAGTOLOJA_VALE_TICKET Then
                            objMovCaixa1.dValor = objMovCaixa1.dValor + objMovCaixa.dValor
                            iAchou = 1
                            Exit For
                        End If
                    Next
                    
                    If iAchou = 0 Then
                        Set objMovCaixa1 = New ClassMovimentoCaixa
                        objMovCaixa1.iTipo = TIPOMEIOPAGTOLOJA_VALE_TICKET
                        objMovCaixa1.dValor = objMovCaixa.dValor
                        colMeiosPag.Add objMovCaixa1
                    End If
                
        End Select

    End If

    Trata_MovCaixa = SUCESSO
    
    Exit Function

Erro_Trata_MovCaixa:

    Trata_MovCaixa = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 149449)

    End Select
        
    Exit Function

End Function

Function ECF_AcrescimoDescontoCupom(ByVal sValorAcrescimo As String, ByVal sValorDesconto As String, ByVal sMsgDesconto As String, ByVal sMsgAcrescimo As String) As Long

Dim vbMsg As VbMsgBoxResult
Dim lErro As Long
Dim lErro1 As Long

On Error GoTo Erro_ECF_AcrescimoDescontoCupom

    DoEvents

    lErro = AFRAC_AcrescimoDescontoCupom(sValorAcrescimo, sValorDesconto, sMsgDesconto, sMsgAcrescimo)
    lErro1 = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Acrescimo - Desconto")
'    If lErro <> SUCESSO And lErro <> 1 Then gError 133741
    'If lErro1 <> SUCESSO Then gError 133741
    
    vbMsg = vbYes
    
    'se nao conseguiu imprimir os comprovantes e quer continuar tentando
    Do While vbMsg = vbYes And lErro1 <> SUCESSO
    
        vbMsg = Rotina_AvisoECF(vbYesNo, "Impressora não responde, tentar novamente?")
        
        If vbMsg = vbYes Then
        
            lErro = AFRAC_AcrescimoDescontoCupom(sValorAcrescimo, sValorDesconto, sMsgDesconto, sMsgAcrescimo)
            lErro1 = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Acrescimo - Desconto")
'            If lErro <> SUCESSO And lErro <> 1 Then gError 133742
            'If lErro1 <> SUCESSO Then gError 133742
            
        Else
            gError 133740
        End If
    
    Loop

    ECF_AcrescimoDescontoCupom = SUCESSO
    
    Exit Function
    
Erro_ECF_AcrescimoDescontoCupom:

    ECF_AcrescimoDescontoCupom = gErr
    
    Select Case gErr
    
        Case 133740 To 133742
    
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 149450)

    End Select
        
    Exit Function

End Function
    
Function ECF_FormaPagamento(ByVal sDescForma As String, ByVal sIndice As String, ByVal sValor As String, ByVal sMsg As String) As Long

Dim vbMsg As VbMsgBoxResult
Dim lErro As Long
Dim lErro1 As Long

On Error GoTo Erro_ECF_FormaPagamento

    'informa a forma de pagamento passando os seus dados
    lErro = AFRAC_FormaPagamento(sDescForma, sIndice, sValor, sMsg)
    lErro1 = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Forma Pagamento")
'    If lErro1 <> SUCESSO Then gError 133743

    vbMsg = vbYes
    
    'se nao conseguiu imprimir os comprovantes e quer continuar tentando
    Do While vbMsg = vbYes And lErro1 <> SUCESSO
    
        vbMsg = Rotina_AvisoECF(vbYesNo, "Impressora não responde, tentar novamente?")
        
        If vbMsg = vbYes Then
        
            'informa a forma de pagamento passando os seus dados
            lErro = AFRAC_FormaPagamento(sDescForma, sIndice, sValor, sMsg)
'            lErro1 = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Forma Pagamento")
'            If lErro <> SUCESSO And lErro <> 1 Then gError 133744
'            If lErro1 <> SUCESSO Then gError 133744
        
        Else
            gError 133745
        End If
    
    Loop

    ECF_FormaPagamento = SUCESSO
    
    Exit Function
    
Erro_ECF_FormaPagamento:

    ECF_FormaPagamento = gErr
    
    Select Case gErr
    
        Case 133743 To 133745
    
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 149451)

    End Select
        
    Exit Function

End Function
    
Function ECF_AbrirGaveta() As Long

Dim vbMsg As VbMsgBoxResult
Dim lErro As Long
Dim lErro1 As Long

On Error GoTo Erro_ECF_AbrirGaveta

    'informa a forma de pagamento passando os seus dados
    lErro = AFRAC_AbrirGaveta()
    lErro1 = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Abrir Gaveta")
'    If lErro <> SUCESSO And lErro <> 1 Then gError 133746
'    If lErro1 <> SUCESSO Then gError 133746

    vbMsg = vbYes
    
    'se nao conseguiu imprimir os comprovantes e quer continuar tentando
    Do While vbMsg = vbYes And lErro1 <> SUCESSO
    
        vbMsg = Rotina_AvisoECF(vbYesNo, "Impressora não responde, tentar novamente?")
        
        If vbMsg = vbYes Then
        
            'informa a forma de pagamento passando os seus dados
            lErro = AFRAC_AbrirGaveta()
            lErro1 = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Abrir Gaveta")
'            If lErro <> SUCESSO And lErro <> 1 Then gError 133747
        
        Else
            gError 133748
        End If
    
    Loop

    ECF_AbrirGaveta = SUCESSO
    
    Exit Function
    
Erro_ECF_AbrirGaveta:

    ECF_AbrirGaveta = gErr
    
    Select Case gErr
    
        Case 133746 To 133748
    
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 149452)

    End Select
        
    Exit Function

End Function

Function ECF_FecharCupom(ByVal objTela As Object, ByVal objVenda As ClassVenda, ByVal bVinculado As Boolean, ByVal sCPF As String, ByVal sNomeCliente As String, ByVal sEndereco As String, ByVal bCupomAdicional As Boolean, ByVal sMsg As String, ByVal sVendedor As String) As Long

Dim vbMsg As VbMsgBoxResult
Dim lErro As Long
Dim lErro1 As Long

On Error GoTo Erro_ECF_FecharCupom

    'informa a forma de pagamento passando os seus dados
    lErro = AFRAC_FecharCupom(objTela, objVenda, bVinculado, sCPF, sNomeCliente, sEndereco, bCupomAdicional, sMsg, sVendedor)
    lErro1 = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Fechar Cupom")
'    If lErro <> SUCESSO And lErro <> 1 Then gError 133749
'    If lErro1 <> SUCESSO Then gError 133749
    
    vbMsg = vbYes
    
    'se nao conseguiu imprimir os comprovantes e quer continuar tentando
    Do While vbMsg = vbYes And lErro1 <> SUCESSO
    
        vbMsg = Rotina_AvisoECF(vbYesNo, "Impressora não responde, tentar novamente?")
        
        If vbMsg = vbYes Then
        
            'informa a forma de pagamento passando os seus dados
            lErro = AFRAC_FecharCupom(objTela, objVenda, bVinculado, sCPF, sNomeCliente, sEndereco, bCupomAdicional, sMsg, sVendedor)
            lErro1 = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Fechar Cupom")
'            If lErro <> SUCESSO And lErro <> 1 Then gError 133750
        
        Else
            gError 133751
        End If
    
    Loop

    ECF_FecharCupom = SUCESSO
    
    Exit Function
    
Erro_ECF_FecharCupom:

    ECF_FecharCupom = gErr
    
    Select Case gErr
    
        Case 133749 To 133751
    
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 149453)

    End Select
        
    Exit Function

End Function

Function CaixasAtivos_Lock_Old() As Long

Dim sNomeArq As String
Dim sNomeArq1 As String
Dim sRet As String
Dim sRetorno As String
Dim lTamanho As Long
Dim sDir As String
Dim objArqSeq As New ClassArqSeq
Dim lErro As Long
Dim lTentaAcessoArquivo As Long
Dim dtData As Date

On Error GoTo Erro_CaixasAtivos_Lock_Old

    dtData = gdtDataHoje
        
    lTamanho = 255
    sRetorno = String(lTamanho, 0)
        
    'funcao que verifica se o caixaconfig.ini foi alterado
    lErro = CF_ECF("CaixaConfig_Verifica_Chave")
    If lErro <> SUCESSO Then gError 204246
        
    Call GetPrivateProfileString(APLICACAO_ORCAMENTO, "Dir", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
    If sRetorno = String(lTamanho, 0) Then gError 204222
        
    sDir = StringZ(sRetorno)
    
    'se o diretorio nao for terminado por \  ===> acrescentar
    If right(sDir, 1) <> "\" Then sDir = sDir & "\"

    sNomeArq1 = "CaixasAtivos_" & Format(dtData, "ddmmyy") & ".txt"

    sNomeArq = sDir & sNomeArq1

    'Verifica se tem orcamento para a data de hoje
    sRet = Dir(sNomeArq)
    
    If Len(sRet) = 0 Then
        
        sRet = Dir("CaixasAtivos_*.txt")
        
        If Len(sRet) > 0 Then
        
            'exclui os arquivos antigos
            Kill "CaixasAtivos_*.txt"
            
        End If
        
    End If
        
    lTentaAcessoArquivo = 0
    Open sNomeArq For Append Lock Read Write As #4

    CaixasAtivos_Lock_Old = SUCESSO

    Exit Function

Erro_CaixasAtivos_Lock_Old:

    CaixasAtivos_Lock_Old = gErr
    
    Select Case gErr

        Case 70
            lTentaAcessoArquivo = lTentaAcessoArquivo + 1
            If lTentaAcessoArquivo < NUM_MAX_LOCK Then Resume
            Call Rotina_ErroECF(vbOKOnly, ERRO_ARQUIVO_LOCADO, gErr, sNomeArq)

        Case 204222, 204246

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 204223)

    End Select
    
    Exit Function

End Function

Public Function CaixasAtivos_Le_Old(ByVal colCaixa As Collection) As Long
'Função Que le os caixas ativos

Dim lErro As Long
Dim sRegistro As String


On Error GoTo Erro_CaixasAtivos_Le_Old
    
    lErro = CF_ECF("CaixasAtivos_Lock")
    If lErro <> SUCESSO Then gError 204224
    
    Do While True
                
        If EOF(4) Then Exit Do
    
        'Busca o próximo registro do arquivo
        Line Input #4, sRegistro
                
        colCaixa.Add sRegistro
                
    Loop
        
    Close #4
    
    CaixasAtivos_Le_Old = SUCESSO
    
    Exit Function
    
Erro_CaixasAtivos_Le_Old:
    
    CaixasAtivos_Le_Old = gErr
    lErro = gErr
    
    Close #4
    
    Select Case lErro
            
        Case 204224
    
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, lErro, Error$, 204225)

    End Select
    
    Exit Function
    
End Function

Function Configura_Impressora_ECF(objTela As Object) As Long

Dim xPageLog As Long
Dim yPageLog As Long
Dim tPoint As POINTAPI
Dim lIdDC As Long


On Error GoTo Erro_Configura_Impressora_ECF

    If Not (objTela Is Nothing) Then
'        objTela.Controls("CD1").Flags = cdlPDReturnDC
        objTela.Controls("CD1").Flags = &H100
        objTela.Controls("CD1").CancelError = True
        objTela.Controls("CD1").ShowPrinter
        objTela.Controls("RT1").Text = ""
        
        lIdDC = SaveDC(objTela.Controls("CD1").hdc)
        
        xPageLog = GetDeviceCaps(objTela.Controls("CD1").hdc, HORZRES)
        yPageLog = GetDeviceCaps(objTela.Controls("CD1").hdc, VERTRES)
    
        
        If Escape(objTela.Controls("CD1").hdc, GETPHYSPAGESIZE, 0, "", tPoint) <= 0 Then
            tPoint.X = xPageLog
            tPoint.Y = yPageLog
        End If
    
    
        Call SetMapMode(objTela.Controls("CD1").hdc, MM_LOMETRIC)
        Call DPtoLP(objTela.Controls("CD1").hdc, tPoint, 2)
        
        
        If (Abs(tPoint.X) < 2099 Or Abs(tPoint.Y) < 1480) And (Abs(tPoint.X) < 2399 Or Abs(tPoint.Y) < 1400) Then gError 204313
        
        Call RestoreDC(objTela.Controls("CD1").hdc, lIdDC)
    
    End If
    
    Configura_Impressora_ECF = SUCESSO
    
    Exit Function

Erro_Configura_Impressora_ECF:
    
    Configura_Impressora_ECF = gErr
    
    Select Case gErr
            
        Case 32755 'usuario cancelou no dialogo de selecao da impressora
        
        Case 204313
            Call Rotina_ErroECF(vbOKOnly, ERRO_TAMANHO_DO_PAPEL_MENOR_QUE_MINIMO, gErr)
            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 204312)

    End Select
    
    Exit Function

End Function

Function Inicializa_LeitoraCodBarras(objTela As Object) As Long

Dim objAliquota As ClassAliquotaICMS
Dim sTributacao As String
Dim iIndice As Integer
Dim lErro As Long
Dim lErro1 As Long
Dim objTiposMeiosPagtos As ClassTMPLoja
Dim iMetodo As Integer
Dim sRetorno As String
Dim sRetorno1 As String
Dim iPos As Integer
Dim sAliquota As String
Dim bAchou As Boolean
Dim sTrib As String
Dim iInicio As Integer
Dim iFim As Integer
Dim objLeitoraCodBarras As Object
Dim objOptionCF As Object
Dim objOptionOrcamento As Object
Dim objBotaoProxNum As Object

On Error GoTo Erro_Inicializa_LeitoraCodBarras

'    If giUsaImpressoraFiscal = 0 Then
'        Set objOptionCF = objTela.Controls("OptionCF")
'        Set objOptionOrcamento = objTela.Controls("OptionOrcamento")
'
'        objOptionCF.Enabled = False
'        objOptionOrcamento.Value = True
'
'    End If
    
'    'se imprime DAV ==> nao usa numero de orcamento
'    If giDAV = 1 Then
''        Set objBotaoProxNum = objTela.Controls("BotaoProxNum")
'        objBotaoProxNum.Visible = False
'    End If

    Set objLeitoraCodBarras = objTela.Controls("LeitoraCodBarras")

    If Len(Trim(gsCodBarrasPorta)) > 0 Then

        objLeitoraCodBarras.CommPort = gsCodBarrasPorta
        
        ' 9600 baud, no parity, 8 data, and 1 stop bit.
        objLeitoraCodBarras.Settings = "9600,N,8,1"
        
        ' Tell the control to read entire buffer when Input
        ' is used.
        objLeitoraCodBarras.InputLen = 0
        
        ' Open the port.
        objLeitoraCodBarras.PortOpen = True
        
        objLeitoraCodBarras.RThreshold = 1
        
    
    End If
    
    Inicializa_LeitoraCodBarras = SUCESSO
    
    Exit Function
    
Erro_Inicializa_LeitoraCodBarras:
    
    Inicializa_LeitoraCodBarras = gErr
    
    Select Case gErr
                
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error$, 175694)

    End Select
    
    Exit Function
    
End Function

Function Valida_Orcamento(objTela As Object) As Long

Dim lErro As Long
Dim lNumero As Long

On Error GoTo Erro_Valida_Orcamento

'    'se nao imprime o orcamento, ou seja, é um PV ==> tem que colocar numero do orcamento
'    If giImprimeOrc = 0 Then

        'If Len(Trim(objTela.Controls("Orcamento").Text)) = 0 Then gError 204331
        
    If objTela.gobjVenda.objCupomFiscal.lNumOrcamento = 0 Then
    
        lErro = CF_ECF("Venda_Obtem_Num_Automatico", lNumero)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        objTela.gobjVenda.objCupomFiscal.lNumOrcamento = lNumero
    
    End If

'    End If

    Valida_Orcamento = SUCESSO
    
    Exit Function
    
Erro_Valida_Orcamento:
    
    Valida_Orcamento = gErr
    
    Select Case gErr
                
        Case 204331
            Call Rotina_ErroECF(vbOKOnly, ERRO_ORCAMENTO_NAO_PREENCHIDO, gErr, Error$)
                
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error$, 204332)

    End Select
    
    Exit Function
    
End Function

'Public Function DAV_Le(ByVal dtData As Date, ByVal colOrcamento As Collection) As Long
''Função Que le os DAVs
'
'Dim lErro As Long
'Dim sNomeArq As String
'Dim sRegistro As String
'Dim colRegistro As New Collection
'Dim lTamanho As Long
'Dim sRetorno As String
'Dim sNomeArq1 As String
'Dim sRet As String
'Dim sNomeArq0 As String
'Dim lTentaAcessoArquivo As Long
'Dim sDir As String
'
'On Error GoTo Erro_DAV_Le
'
'
'    lTamanho = 255
'    sRetorno = String(lTamanho, 0)
'
'    Call GetPrivateProfileString(APLICACAO_ORCAMENTO, "Dir", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
'    If sRetorno = String(lTamanho, 0) Then gError 204340
'
'    sDir = StringZ(sRetorno)
'
'    'se o diretorio nao for terminado por \  ===> acrescentar
'    If right(sDir, 1) <> "\" Then sDir = sDir & "\"
'
'    sNomeArq1 = "DAV" & Format(dtData, "ddmmyy") & (".txt")
'
'    sNomeArq = sDir & sNomeArq1
'
'    'esta atribuicao é feita pois caso haja algum erro será mostrado o nome contido em sNomeArq0
'    sNomeArq0 = sNomeArq
'
'    'Verifica se tem orcamento para a data de hoje
'    sRet = Dir(sNomeArq)
'
'    If Len(sRet) = 0 Then
'
'        lTentaAcessoArquivo = 0
'        Open sNomeArq For Append Lock Read Write As #2
'
'        'Inserido no Arquivo o nome do arquivo
'        Print #2, sNomeArq1
'
'        Close #2
'
'    End If
'
'    lTentaAcessoArquivo = 0
'    Open sNomeArq For Input Lock Write As #2
'
'    'Busca o próximo registro do arquivo
'    Line Input #2, sRegistro
'
'    'se nao tiver o nome do arquivo no primeiro registro ==> nao é o arquivo de orcamento que deveria ser lido
'    If sRegistro <> sNomeArq1 Then gError 204341
'
'    Do While Not EOF(2)
'
'        'Busca o próximo registro do arquivo
'        Line Input #2, sRegistro
'
'        If sRegistro <> "" Then
'
'            Set colRegistro = New Collection
'
'            colRegistro.Add sRegistro
'
'            lErro = CF_ECF("Registro_ECF_CC", colRegistro, "", 0, 0, colOrcamento)
'            If lErro <> SUCESSO Then gError 204342
'
'        End If
'
'    Loop
'
'    Close #2
'
'    DAV_Le = SUCESSO
'
'    Exit Function
'
'Erro_DAV_Le:
'
'    DAV_Le = gErr
'
'    Close #2
'
'    Select Case gErr
'
'        Case 70
'            lTentaAcessoArquivo = lTentaAcessoArquivo + 1
'            If lTentaAcessoArquivo < NUM_MAX_LOCK Then Resume
'            Call Rotina_ErroECF(vbOKOnly, ERRO_ARQUIVO_LOCADO, gErr, sNomeArq)
'
'        Case 204340
'            Call Rotina_ErroECF(vbOKOnly, ERRO_PREENCHIMENTO_ARQUIVO_CONFIG, gErr, "Dir", APLICACAO_ORCAMENTO, NOME_ARQUIVO_CAIXA)
'
'        Case 204341
'            Call Rotina_ErroECF(vbOKOnly, ERRO_ARQUIVO_DAV_INVALIDO, gErr, sNomeArq0)
'
'        Case 204342
'
'        Case Else
'            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, lErro, Error$, 204343)
'
'    End Select
'
'    Exit Function
'
'End Function

Sub Inicializa_FechaOrcamento(objTela As Object)

'    If giImprimeOrc = 0 Then 'PAFECF
'        objTela.Controls("BotaoImprimir").Visible = False
'        objTela.Controls("F12").Visible = False
'    Else
'        objTela.Controls("BotaoGravar").Visible = False
'        objTela.Controls("F3").Visible = False
'    End If

'    objTela.Controls("BotaoExcluir").Visible = False
'    objTela.Controls("Escape").Visible = False

End Sub


Function DAV_Executa_RelGer(ByVal dtDataDe As Date, ByVal dtDataAte As Date, objTela As Object) As Long

Dim lErro As Long
Dim lSequencial As Long

On Error GoTo Erro_DAV_Executa_RelGer

    'Função que Abre a Transação
    lErro = CF_ECF("Caixa_Transacao_Abrir", lSequencial)
    If lErro <> SUCESSO Then gError 204351

    'Função que Referenciar a Função da Afrac que Vai Executar a Leitura
    lErro = CF_ECF("DAV_Executa_RelGerEmTrans", lSequencial, dtDataDe, dtDataAte, objTela)
    If lErro <> SUCESSO Then gError 204352

    'Função que Encerra a Transação com o Caixa
    lErro = CF_ECF("Caixa_Transacao_Fechar", lSequencial)
    If lErro <> SUCESSO Then gError 204353

    DAV_Executa_RelGer = SUCESSO

    Exit Function

Erro_DAV_Executa_RelGer:

    DAV_Executa_RelGer = gErr

     Select Case gErr

        Case 204351, 204352, 204353
            'Erro Tratado Dentro da Função que Foi Chamada

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 204354)

    End Select
    
'    lSequencial = glSeqTransacaoAberta
            
    'Rollback na transação
    Call CF_ECF("Caixa_Transacao_Rollback", glTransacaoPAFECF)

    Exit Function

End Function

Function DAV_Executa_RelGerEmTrans(lSequencial As Long, ByVal dtDataDe As Date, ByVal dtDataAte As Date, objTela As Object) As Long

Dim lErro As Long
Dim objMovimentoCaixa As New ClassMovimentoCaixa
Dim sLog As String
Dim colRegistro As New Collection
Dim sMensagem As String
Dim colOrcamento As Collection
Dim objVenda As ClassVenda
Dim iIndice As Integer
Dim colMensagem As New Collection
Dim iImprimiu As Integer
Dim dtData As Date
Dim alComando(1 To 1) As Long
Dim sMsg As String
Dim sArquivo As String
Dim iRegistro As Integer
Dim tOrcamento As typeOrcamentoECF
Dim lCodigo As Long
Dim lCOO As Long
Dim lCOOCupom As Long


On Error GoTo Erro_DAV_Executa_RelGerEmTrans

    If giOrcamentoECF <> CAIXA_SO_ORCAMENTO Then


        For iIndice = LBound(alComando) To UBound(alComando)
            alComando(iIndice) = Comando_AbrirExt(glConexaoOrcPAF)
            If alComando(iIndice) = 0 Then gError 204745
        Next
    
    
        'Função que Move os Dados para a Memória
        lErro = CF_ECF("MovimentoCaixa_Move_Dados_Memoria", MOVIMENTOCAIXA_DAV_RELGER, objMovimentoCaixa)
        If lErro <> SUCESSO Then gError 204359
        
        'Guarda o Sequencial no objMovimentos Caixa
        objMovimentoCaixa.lSequencial = lSequencial
        
        Call CF_ECF("MovimentoCaixa_Gera_Log", colRegistro, objMovimentoCaixa)
        
        'Guarda as Informações no Arquivo
        lErro = CF_ECF("MovimentoCaixaECF_Grava", colRegistro)
        If lErro <> SUCESSO Then gError 204360
    
        lErro = AFRAC_AbrirRelatorioGerencial(RELGER_DAV, objTela)
        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Abrir DAV Emitidos")
        If lErro <> SUCESSO Then gError 204361
    
        colMensagem.Add Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "DAV EMITIDOS")
        colMensagem.Add Space(48)
        colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "Periodo De: " & Format(dtDataDe, "dd/mm/yyyy") & " a " & Format(dtDataAte, "dd/mm/yyyy"))
        colMensagem.Add Space(48)
        colMensagem.Add "   Número      Data     Título       Valor Total"
        
        For iIndice = 1 To colMensagem.Count
                
            sMensagem = colMensagem.Item(iIndice)
    
            lErro = AFRAC_ImprimirRelatorioGerencial(sMensagem, objTela)
            lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "DAV Emitidos")
            If lErro <> SUCESSO Then gError 204362
            
        Next
    
        tOrcamento.sMsg = String(STRING_ORCAMENTO_MSG, 0)
        tOrcamento.sNumeroDAV = String(STRING_ORCAMENTO_NUMERODAV, 0)
        
        lErro = Comando_Executar(alComando(1), "SELECT Codigo, Seq, Msg, COO, COOCupom, NumeroDAV FROM Orcamento WHERE DataEmissao >= ? AND DataEmissao <= ? AND NumeroDAV <> '0000000000' AND NumeroDAV <> '0000000000' ORDER BY Codigo, Seq", tOrcamento.lCodigo, tOrcamento.iSeq, tOrcamento.sMsg, tOrcamento.lCOO, tOrcamento.lCOOCupom, tOrcamento.sNumeroDAV, dtDataDe, dtDataAte)
        If lErro <> AD_SQL_SUCESSO Then gError 204770
        
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 204771
        
        lCodigo = tOrcamento.lCodigo
        
        lCOO = tOrcamento.lCOO
        lCOOCupom = tOrcamento.lCOOCupom
        
        Set colRegistro = New Collection
        
        Do While lErro = AD_SQL_SUCESSO
        
            If tOrcamento.lCodigo <> lCodigo Then
                
                Set objVenda = New ClassVenda
            
                iRegistro = 1
        
                lErro = CF_ECF("Registro_ECF_Vendas", colRegistro, sArquivo, iRegistro, objVenda)
                If lErro <> SUCESSO Then gError 204772
        
                If objVenda.objCupomFiscal.iTipo = OPTION_DAV Then
                    
                    sMensagem = Formata_Campo(ALINHAMENTO_DIREITA, 13, " ", Format(objVenda.objCupomFiscal.lNumeroDAV, "0000000000")) & Formata_Campo(ALINHAMENTO_DIREITA, 9, " ", Format(objVenda.objCupomFiscal.dtDataEmissao, "dd/mm/yy")) & " Orçamento" & Formata_Campo(ALINHAMENTO_ESQUERDA, 16, " ", Format(objVenda.objCupomFiscal.dValorTotal, "standard"))
                
                    lErro = AFRAC_ImprimirRelatorioGerencial(sMensagem, objTela)
                    lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "DAV Emitidos")
                    If lErro <> SUCESSO Then gError 204364
                    
                    sMensagem = "COO: " & Format(lCOO, "000000") & "  COO Vinculado: " & Format(lCOOCupom, "000000")
                
                    lErro = AFRAC_ImprimirRelatorioGerencial(sMensagem, objTela)
                    lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "DAV Emitidos")
                    If lErro <> SUCESSO Then gError 210812
                    
                    iImprimiu = 1
                    
                    lCOO = tOrcamento.lCOO
                    lCOOCupom = tOrcamento.lCOOCupom
                
                End If
    
                Set colRegistro = New Collection
    
                lCodigo = tOrcamento.lCodigo
        
            End If
    
            colRegistro.Add tOrcamento.sMsg
    
            lErro = Comando_BuscarProximo(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 204773
    
        Loop
        
        If colRegistro.Count > 0 Then
        
            Set objVenda = New ClassVenda
        
            iRegistro = 1
    
            lErro = CF_ECF("Registro_ECF_Vendas", colRegistro, sArquivo, iRegistro, objVenda)
            If lErro <> SUCESSO Then gError 204772
    
            If objVenda.objCupomFiscal.iTipo = OPTION_DAV Then
                
                sMensagem = Formata_Campo(ALINHAMENTO_DIREITA, 13, " ", Format(objVenda.objCupomFiscal.lNumeroDAV, "0000000000")) & Formata_Campo(ALINHAMENTO_DIREITA, 9, " ", Format(objVenda.objCupomFiscal.dtDataEmissao, "dd/mm/yy")) & " Orçamento" & Formata_Campo(ALINHAMENTO_ESQUERDA, 16, " ", Format(objVenda.objCupomFiscal.dValorTotal, "standard"))
            
                lErro = AFRAC_ImprimirRelatorioGerencial(sMensagem, objTela)
                lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "DAV Emitidos")
                If lErro <> SUCESSO Then gError 204364
                
                sMensagem = "COO: " & Format(lCOO, "000000") & "  COO Vinculado: " & Format(lCOOCupom, "000000")
            
                lErro = AFRAC_ImprimirRelatorioGerencial(sMensagem, objTela)
                lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "DAV Emitidos")
                If lErro <> SUCESSO Then gError 210812
                
                iImprimiu = 1
            
            End If
    
            Set colRegistro = New Collection
        
        End If
        
        If iImprimiu = 0 Then
        
            sMensagem = "******* NÃO HÁ DAV NO PERIODO SOLICITADO *******"
    
            lErro = AFRAC_ImprimirRelatorioGerencial(sMensagem, objTela)
            lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "DAV Emitidos")
            If lErro <> SUCESSO Then gError 204365
    
        End If
        
        lErro = AFRAC_FecharRelatorioGerencial(objTela)
        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Fecha Orcamento")
        If lErro <> SUCESSO Then gError 204366
    
        
    
        For iIndice = LBound(alComando) To UBound(alComando)
            Call Comando_Fechar(alComando(iIndice))
        Next

    
    Else
        
        Call CF_ECF("Trata_Caixa_So_Orcamento")
        
    End If

    DAV_Executa_RelGerEmTrans = SUCESSO
    
    Exit Function

Erro_DAV_Executa_RelGerEmTrans:

    DAV_Executa_RelGerEmTrans = gErr

        Select Case gErr
        
            Case 204359 To 204365, 210812
                Call AFRAC_FecharRelatorioGerencial(objTela)
        
            Case 204366, 204772
        
            Case 204745
                Call Rotina_ErroECF(vbOKOnly, ERRO_ABERTURA_COMANDO, gErr)
        
            Case 204770, 204771, 204773
                Call Rotina_ErroECF(vbOKOnly, ERRO_LEITURA_ORCAMENTO, gErr)
                
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 204367)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'Public Function DAV_Le1(ByVal dtData As Date, ByVal colOrcamento As Collection) As Long
''Função Que le os DAVs
'
'Dim lErro As Long
'Dim sNomeArq As String
'Dim sRegistro As String
'Dim colRegistro As New Collection
'Dim lTamanho As Long
'Dim sRetorno As String
'Dim sNomeArq1 As String
'Dim sRet As String
'Dim sNomeArq0 As String
'Dim lTentaAcessoArquivo As Long
'Dim sDir As String
'
'On Error GoTo Erro_DAV_Le1
'
'
'    lTamanho = 255
'    sRetorno = String(lTamanho, 0)
'
'    Call GetPrivateProfileString(APLICACAO_ORCAMENTO, "Dir", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
'    If sRetorno = String(lTamanho, 0) Then gError 204355
'
'    sDir = StringZ(sRetorno)
'
'    'se o diretorio nao for terminado por \  ===> acrescentar
'    If right(sDir, 1) <> "\" Then sDir = sDir & "\"
'
'    sNomeArq1 = "DAV" & Format(dtData, "ddmmyy") & (".txt")
'
'    sNomeArq = sDir & sNomeArq1
'
'    'esta atribuicao é feita pois caso haja algum erro será mostrado o nome contido em sNomeArq0
'    sNomeArq0 = sNomeArq
'
'    'Verifica se tem orcamento para a data de hoje
'    sRet = Dir(sNomeArq)
'
'    'se o arquivo existe
'    If Len(sRet) <> 0 Then
'
'        lTentaAcessoArquivo = 0
'        Open sNomeArq For Input Lock Write As #2
'
'        'Busca o próximo registro do arquivo
'        Line Input #2, sRegistro
'
'        'se nao tiver o nome do arquivo no primeiro registro ==> nao é o arquivo de orcamento que deveria ser lido
'        If sRegistro <> sNomeArq1 Then gError 204356
'
'        Do While Not EOF(2)
'
'            'Busca o próximo registro do arquivo
'            Line Input #2, sRegistro
'
'            If sRegistro <> "" Then
'
'                Set colRegistro = New Collection
'
'                colRegistro.Add sRegistro
'
'                lErro = CF_ECF("Registro_ECF_CC", colRegistro, "", 0, 0, colOrcamento)
'                If lErro <> SUCESSO Then gError 204357
'
'            End If
'
'        Loop
'
'        Close #2
'
'    End If
'
'    DAV_Le1 = SUCESSO
'
'    Exit Function
'
'Erro_DAV_Le1:
'
'    DAV_Le1 = gErr
'
'    Close #2
'
'    Select Case gErr
'
'        Case 70
'            lTentaAcessoArquivo = lTentaAcessoArquivo + 1
'            If lTentaAcessoArquivo < NUM_MAX_LOCK Then Resume
'            Call Rotina_ErroECF(vbOKOnly, ERRO_ARQUIVO_LOCADO, gErr, sNomeArq)
'
'        Case 204355
'            Call Rotina_ErroECF(vbOKOnly, ERRO_PREENCHIMENTO_ARQUIVO_CONFIG, gErr, "Dir", APLICACAO_ORCAMENTO, NOME_ARQUIVO_CAIXA)
'
'        Case 204356
'            Call Rotina_ErroECF(vbOKOnly, ERRO_ARQUIVO_DAV_INVALIDO, gErr, sNomeArq0)
'
'        Case 204357
'
'        Case Else
'            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, lErro, Error$, 204358)
'
'    End Select
'
'    Exit Function
'
'End Function

Sub Executa_LeituraX()

Dim lErro As Long

On Error GoTo Erro_BotaoCaixaLeituraX_Click

    If giOrcamentoECF = CAIXA_SO_ORCAMENTO Then gError 105910

    If giUsaImpressoraFiscal = 0 Then gError 204384

    'Se o Caixa Estiver Fechada Emite menssagem de Erro
'    If giStatusCaixa = STATUS_CAIXA_FECHADO Then gError 107562

    'Chama a Função que faz LeituraX
    lErro = CF_ECF("Caixa_Executa_LeituraX")
    If lErro <> SUCESSO Then gError 107563

    Exit Sub

Erro_BotaoCaixaLeituraX_Click:

    Select Case gErr

        Case 105910
            Call Rotina_ErroECF(vbOKOnly, ERRO_CAIXA_SO_ORCAMENTO, gErr)

        Case 107562
            Call Rotina_ErroECF(vbOKOnly, ERRO_CAIXA_FECHADO, gErr, giCodCaixa)

        Case 107563

                
        Case 204384
            Call Rotina_ErroECF(vbOKOnly, ERRO_FUNCAO_NAO_DISPONIVEL_NAOFISCAL, gErr)
                
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163648)

    End Select

    Exit Sub

End Sub

Function EspelhoMFD_Executa(ByVal iTipoLeitura As Integer, ByVal sDe As String, ByVal sAte As String) As Long
'Função que Vai Chamar a Função da Afrac que vai Executar a Leitura de Memória Fiscal

Dim lErro As Long
Dim lSequencial As Long
Dim iCodGerente As Integer
Dim objOperador As New ClassOperador

On Error GoTo Erro_EspelhoMFD_Executa

    'Se for Necessário a autorização do Gerente para abertura do Caixa
    If gobjLojaECF.iGerenteAutoriza = AUTORIZACAO_GERENTE Then

        'Chama a Tela de Senha
        Call Chama_TelaECF_Modal("OperadorLogin", objOperador, LOGIN_APENAS_GERENTE)

        'Sai de Função se a Tela de Login não Retornar ok
        If giRetornoTela <> vbOK Then gError 204438

        'Se o Operador For Gerente
        iCodGerente = objOperador.iCodigo

    End If

    'Função que Abre a Transação
    lErro = CF_ECF("Caixa_Transacao_Abrir", lSequencial)
    If lErro <> SUCESSO Then gError 204439

    'Função que Referenciar a Função da Afrac que Vai Executar a Leitura
    lErro = CF_ECF("EspelhoMFD_ExecutaEmTrans", lSequencial, iCodGerente, iTipoLeitura, sDe, sAte)
    If lErro <> SUCESSO Then gError 204440

    'Função que Encerra a Transação com o Caixa
    lErro = CF_ECF("Caixa_Transacao_Fechar", lSequencial)
    If lErro <> SUCESSO Then gError 204441

    EspelhoMFD_Executa = SUCESSO

    Exit Function

Erro_EspelhoMFD_Executa:

    EspelhoMFD_Executa = gErr

     Select Case gErr

        Case 204438 To 204441
            'Erro Tratado Dentro da Função que Foi Chamada

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 204442)

    End Select
    
'    lSequencial = glSeqTransacaoAberta
            
    'Rollback na transação
    Call CF_ECF("Caixa_Transacao_Rollback", glTransacaoPAFECF)

    Exit Function

End Function

Function EspelhoMFD_ExecutaEmTrans(lSequencial As Long, iCodGerente As Integer, ByVal iTipoLeitura As Integer, ByVal sDe As String, ByVal sAte As String) As Long

Dim lErro As Long
Dim objMovimentoCaixa As New ClassMovimentoCaixa
Dim sCampoReducao1 As String
Dim sCampoReducao2 As String
Dim sLog As String
Dim colRegistro As New Collection

On Error GoTo Erro_EspelhoMFD_ExecutaEmTrans


    If giOrcamentoECF <> CAIXA_SO_ORCAMENTO Then

        'Função que Move os Dados para a Memória
        lErro = CF_ECF("MovimentoCaixa_Move_Dados_Memoria", MOVIMENTOCAIXA_ESPELHO_MFD, objMovimentoCaixa)
        If lErro <> SUCESSO Then gError 204443
        
        'Guarda o Sequencial no objMovimentos Caixa
        objMovimentoCaixa.lSequencial = lSequencial
        
        ' Codigo Gerente
        objMovimentoCaixa.iGerente = iCodGerente
        
        Call CF_ECF("MovimentoCaixa_Gera_Log", colRegistro, objMovimentoCaixa)
        
        'Guarda as Informações no Arquivo
        lErro = CF_ECF("MovimentoCaixaECF_Grava", colRegistro)
        If lErro <> SUCESSO Then gError 204444
        
        Set colRegistro = New Collection
    
        If iTipoLeitura = LEITURA_DATAS Then
                      
            'Verifica se a Data Inicial é Maior que a segunda
            If StrParaDate(sDe) > StrParaDate(sAte) Then gError 204445
    
           lErro = AFRAC_EspelhoMFD(LEITURA_DATAS, Format(sDe, "dd/mm/yy"), Format(sAte, "dd/mm/yy"))
           lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Espelho MFD")
           If lErro <> SUCESSO Then gError 204446
           
        ElseIf iTipoLeitura = LEITURA_COO Then
    
            'Verificar se as Reduções Estão Preenchidas se Erro
            If Len(Trim(sDe)) = 0 Or Len(Trim(sAte)) = 0 Then gError 204447
    
           If StrParaLong(sDe) > StrParaLong(sAte) Then gError 204448
    
            'Função Que Coloca Zero a Esqueda
            sDe = Formata_Campo(ALINHAMENTO_ESQUERDA, 6, "0", StrParaLong(sDe))
    
            'Função Que Coloca Zero a Esqueda
            sAte = Formata_Campo(ALINHAMENTO_ESQUERDA, 6, "0", StrParaLong(sAte))
    
            'Se for por Intervalo de Redução
           lErro = AFRAC_EspelhoMFD(LEITURA_COO, sDe, sAte)
           lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Espelho MFD")
           If lErro <> SUCESSO Then gError 204449
            
        End If
    
        EspelhoMFD_ExecutaEmTrans = SUCESSO
    
    Else
        
        Call CF_ECF("Trata_Caixa_So_Orcamento")
        
    End If
    
    Exit Function

Erro_EspelhoMFD_ExecutaEmTrans:

    EspelhoMFD_ExecutaEmTrans = gErr

        Select Case gErr
        
        Case 204443 To 204444, 204446, 204449

        Case 204445
            Call Rotina_ErroECF(vbOKOnly, ERRO_DATA_INICIAL_MAIOR1, gErr)

        Case 204447
            Call Rotina_ErroECF(vbOKOnly, ERRO_COO_NAO_PREENCHIDO, gErr)
        
        Case 204448
            Call Rotina_ErroECF(vbOKOnly, ERRO_COO_INICIAL_MAIOR_FINAL, gErr)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 204450)

    End Select

    Exit Function

End Function

Function ArqMFD_Executa(ByVal iTipoLeitura As Integer, ByVal sDe As String, ByVal sAte As String) As Long
'Função que Vai Chamar a Função da Afrac que vai Executar a Leitura de Memória Fiscal

Dim lErro As Long
Dim lSequencial As Long
Dim iCodGerente As Integer
Dim objOperador As New ClassOperador

On Error GoTo Erro_ArqMFD_Executa

    'Se for Necessário a autorização do Gerente para abertura do Caixa
    If gobjLojaECF.iGerenteAutoriza = AUTORIZACAO_GERENTE Then

        'Chama a Tela de Senha
        Call Chama_TelaECF_Modal("OperadorLogin", objOperador, LOGIN_APENAS_GERENTE)

        'Sai de Função se a Tela de Login não Retornar ok
        If giRetornoTela <> vbOK Then gError 204498

        'Se o Operador For Gerente
        iCodGerente = objOperador.iCodigo

    End If

    'Função que Abre a Transação
    lErro = CF_ECF("Caixa_Transacao_Abrir", lSequencial)
    If lErro <> SUCESSO Then gError 204499

    'Função que Referenciar a Função da Afrac que Vai Executar a Leitura
    lErro = CF_ECF("ArqMFD_ExecutaEmTrans", lSequencial, iCodGerente, iTipoLeitura, sDe, sAte)
    If lErro <> SUCESSO Then gError 204500

    'Função que Encerra a Transação com o Caixa
    lErro = CF_ECF("Caixa_Transacao_Fechar", lSequencial)
    If lErro <> SUCESSO Then gError 204501

    ArqMFD_Executa = SUCESSO

    Exit Function

Erro_ArqMFD_Executa:

    ArqMFD_Executa = gErr

     Select Case gErr

        Case 204498 To 204501
            'Erro Tratado Dentro da Função que Foi Chamada

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 204502)

    End Select
    
'    lSequencial = glSeqTransacaoAberta
            
    'Rollback na transação
    Call CF_ECF("Caixa_Transacao_Rollback", glTransacaoPAFECF)

    Exit Function

End Function

Function ArqMFD_ExecutaEmTrans(lSequencial As Long, iCodGerente As Integer, ByVal iTipoLeitura As Integer, ByVal sDe As String, ByVal sAte As String) As Long

Dim lErro As Long
Dim objMovimentoCaixa As New ClassMovimentoCaixa
Dim sCampoReducao1 As String
Dim sCampoReducao2 As String
Dim sLog As String
Dim colRegistro As New Collection

On Error GoTo Erro_ArqMFD_ExecutaEmTrans


    If giOrcamentoECF <> CAIXA_SO_ORCAMENTO Then

        'Função que Move os Dados para a Memória
        lErro = CF_ECF("MovimentoCaixa_Move_Dados_Memoria", MOVIMENTOCAIXA_ARQUIVO_MFD, objMovimentoCaixa)
        If lErro <> SUCESSO Then gError 204503
        
        'Guarda o Sequencial no objMovimentos Caixa
        objMovimentoCaixa.lSequencial = lSequencial
        
        ' Codigo Gerente
        objMovimentoCaixa.iGerente = iCodGerente
        
        Call CF_ECF("MovimentoCaixa_Gera_Log", colRegistro, objMovimentoCaixa)
        
        'Guarda as Informações no Arquivo
        lErro = CF_ECF("MovimentoCaixaECF_Grava", colRegistro)
        If lErro <> SUCESSO Then gError 204504
        
        Set colRegistro = New Collection
    
        If iTipoLeitura = LEITURA_DATAS Then
                      
            'Verifica se a Data Inicial é Maior que a segunda
            If StrParaDate(sDe) > StrParaDate(sAte) Then gError 204505
    
           lErro = AFRAC_ArqMFD(LEITURA_DATAS, Format(sDe, "dd/mm/yy"), Format(sAte, "dd/mm/yy"))
           lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "ArqMFD")
           If lErro <> SUCESSO Then gError 204506
           
        ElseIf iTipoLeitura = LEITURA_COO Then
    
            'Verificar se as Reduções Estão Preenchidas se Erro
            If Len(Trim(sDe)) = 0 Or Len(Trim(sAte)) = 0 Then gError 204507
    
           If StrParaLong(sDe) > StrParaLong(sAte) Then gError 204508
    
            'Função Que Coloca Zero a Esqueda
            sDe = Formata_Campo(ALINHAMENTO_ESQUERDA, 6, "0", StrParaLong(sDe))
    
            'Função Que Coloca Zero a Esqueda
            sAte = Formata_Campo(ALINHAMENTO_ESQUERDA, 6, "0", StrParaLong(sAte))
    
            'Se for por Intervalo de Redução
           lErro = AFRAC_ArqMFD(LEITURA_COO, sDe, sAte)
           lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Espelho MFD")
           If lErro <> SUCESSO Then gError 204509
            
        End If
    
    Else
        
        Call CF_ECF("Trata_Caixa_So_Orcamento")

    End If

    ArqMFD_ExecutaEmTrans = SUCESSO
    
    Exit Function

Erro_ArqMFD_ExecutaEmTrans:

    ArqMFD_ExecutaEmTrans = gErr

        Select Case gErr
        
        Case 204504, 204504, 204506, 204509

        Case 204505
            Call Rotina_ErroECF(vbOKOnly, ERRO_DATA_INICIAL_MAIOR1, gErr)

        Case 204507
            Call Rotina_ErroECF(vbOKOnly, ERRO_COO_NAO_PREENCHIDO, gErr)
        
        Case 204508
            Call Rotina_ErroECF(vbOKOnly, ERRO_COO_INICIAL_MAIOR_FINAL, gErr)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 204510)

    End Select

    Exit Function

End Function

Function Registro_ECF_MovcxFechamento(ByVal colRegistro As Collection, ByVal sArquivo As String, ByVal iRegistro As Integer, ByVal objMovCx As ClassMovimentoCaixa, ByVal objLojaArqFisMestre As ClassLojaArqFisMestre, ByVal colLojaArqFisAnalitico As Collection, ByVal objUltimaReducao As ClassUltimaReducao) As Long
'Trata o registro de fechamento de caixa e coloca os dados no bd do caixa central. sArquivo contem o nome do arquivo em que sRegistro foi lido.
'iRegistro é o numero do registro.

Dim lErro As Long

On Error GoTo Erro_Registro_ECF_MovcxFechamento

    'guarda as infos de movimento de caixa
    lErro = Movcx_FechamentoCaixa(colRegistro, objMovCx, objLojaArqFisMestre, colLojaArqFisAnalitico, objUltimaReducao)
    If lErro <> SUCESSO Then gError 204514
    
    Registro_ECF_MovcxFechamento = SUCESSO

    Exit Function

Erro_Registro_ECF_MovcxFechamento:

    Registro_ECF_MovcxFechamento = gErr

    Select Case gErr

        Case 204514

        Case Else
            lErro = Rotina_ErroECF(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 204515)

    End Select

    Exit Function

End Function

Private Function Movcx_FechamentoCaixa(ByVal colRegistro As Collection, ByVal objMovCx As ClassMovimentoCaixa, ByVal objLojaArqFisMestre As ClassLojaArqFisMestre, ByVal colLojaArqFisAnalitico As Collection, ByVal objUltimaReducao As ClassUltimaReducao) As Long
'guarda as infos do fechamento de caixa

Dim lErro As Long
Dim objLojaArqFisAnalitico As ClassLojaArqFisAnalitico
Dim objAliquota As ClassAliquotaICMS
Dim iPosAtual As Integer
Dim iPosFimAtual As Integer
Dim sRegistro As String
Dim iRegistroCol As Integer
Dim iIndice As Integer

On Error GoTo Erro_Movcx_FechamentoCaixa

    sRegistro = colRegistro(1)
    iRegistroCol = 1

    iPosFimAtual = InStr(sRegistro, Chr(vbKeyControl))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.iTipo = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.dHora = CDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.dtDataMovimento = CDate(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.iGerente = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.iCodOperador = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.lSequencial = CLng(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.iFilialEmpresa = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objMovCx.iCaixa = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objLojaArqFisMestre.iCodECF = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objLojaArqFisMestre.sNumSerieECF = Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual)

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objLojaArqFisMestre.dGrandeTotal = CDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
    
    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objLojaArqFisMestre.lCOOIni = CLng(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
    
    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objLojaArqFisMestre.lCOOFim = CLng(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objLojaArqFisMestre.lCRO = CLng(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objLojaArqFisMestre.lCRZ = CLng(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))

    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objLojaArqFisMestre.dVendaBruta = CDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
    
    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objLojaArqFisMestre.dtData = CDate(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
    
    iPosAtual = iPosFimAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    objLojaArqFisMestre.iNumEquip = CInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
    
    objLojaArqFisMestre.iFilialEmpresa = objMovCx.iFilialEmpresa
    
    iPosAtual = iPosFimAtual + 1
    
    If iPosAtual > Len(sRegistro) Then
        iRegistroCol = iRegistroCol + 1
        iPosAtual = 1
    End If
    
    For iIndice = iRegistroCol To colRegistro.Count

        sRegistro = colRegistro.Item(iIndice)
    
        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
    
        Do While iPosAtual <= Len(sRegistro) And iPosFimAtual <> 0
    
            Set objLojaArqFisAnalitico = New ClassLojaArqFisAnalitico
        
            objLojaArqFisAnalitico.sSituacaoTrib = Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual)
            
            iPosAtual = iPosFimAtual + 1
            iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
            
            objLojaArqFisAnalitico.dTotalizador = CDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
        
            iPosAtual = iPosFimAtual + 1
            iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyEscape))
        
            colLojaArqFisAnalitico.Add objLojaArqFisAnalitico
        
        Loop
        
        If iPosAtual <= Len(sRegistro) Then
            iRegistroCol = iIndice
            Exit For
        Else
            iPosAtual = 1
        End If
            
    
    Next
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyShift))
        
    If iPosFimAtual <> 0 Then
        objUltimaReducao.dAcrescimoICMS = CDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
        
        iPosAtual = iPosFimAtual + 1
        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyShift))
        objUltimaReducao.dAcrescimoISS = CDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
        
        iPosAtual = iPosFimAtual + 1
        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyShift))
        objUltimaReducao.dAcrescimoNaoFiscal = CDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
        
        iPosAtual = iPosFimAtual + 1
        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyShift))
        objUltimaReducao.dCancelamentoICMS = CDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
        
        iPosAtual = iPosFimAtual + 1
        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyShift))
        objUltimaReducao.dCancelamentoISS = CDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
        
        iPosAtual = iPosFimAtual + 1
        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyShift))
        objUltimaReducao.dCancelamentoNaoFiscal = CDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
        
        iPosAtual = iPosFimAtual + 1
        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyShift))
        objUltimaReducao.dDescontoICMS = CDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
        
        iPosAtual = iPosFimAtual + 1
        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyShift))
        objUltimaReducao.dDescontoISS = CDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
        
        iPosAtual = iPosFimAtual + 1
        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyShift))
        objUltimaReducao.dDescontoNaoFiscal = CDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
        
        iPosAtual = iPosFimAtual + 1
        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyShift))
        objUltimaReducao.dIsentoICMS = CDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
        
        iPosAtual = iPosFimAtual + 1
        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyShift))
        objUltimaReducao.dIsentoISS = CDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
        
        iPosAtual = iPosFimAtual + 1
        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyShift))
        objUltimaReducao.dNaoIncideICMS = CDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
        
        iPosAtual = iPosFimAtual + 1
        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyShift))
        objUltimaReducao.dNaoIncideISS = CDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
        
        iPosAtual = iPosFimAtual + 1
        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyShift))
        objUltimaReducao.dSubstTribICMS = CDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
        
        iPosAtual = iPosFimAtual + 1
        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyShift))
        objUltimaReducao.dSubstTribISS = CDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
        
        iPosAtual = iPosFimAtual + 1
        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyShift))
        objUltimaReducao.dTotalNaoFiscal = CDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
        
        iPosAtual = iPosFimAtual + 1
        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyShift))
        objUltimaReducao.dtDataReducao = CDate(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
        
        iPosAtual = iPosFimAtual + 1
        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyShift))
        objUltimaReducao.dVendaBruta = CDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
        
        iPosAtual = iPosFimAtual + 1
        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyShift))
        objUltimaReducao.sIncidenciaDescontoISS = Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual)
        
        iPosAtual = iPosFimAtual + 1
        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyShift))
        objUltimaReducao.sHoraReducao = Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual)
        
        iPosAtual = iPosFimAtual + 1
        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyShift))
        objUltimaReducao.dtDataMovimento = CDate(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
        
        iPosAtual = iPosFimAtual + 1
        iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyShift))
        objUltimaReducao.lCOOReducaoZ = StrParaLong(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
        
        
        iPosAtual = iPosFimAtual + 1
        
        If iPosAtual > Len(sRegistro) Then
            iRegistroCol = iRegistroCol + 1
            iPosAtual = 1
        End If
        
        For iIndice = iRegistroCol To colRegistro.Count

            sRegistro = colRegistro.Item(iIndice)
    
            iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyShift))
    
            Do While iPosAtual <= Len(sRegistro) And iPosFimAtual <> 0
        
        
                Set objAliquota = New ClassAliquotaICMS
            
                'formato 18,00 ou 12,00
                objAliquota.dAliquota = CDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
                
                iPosAtual = iPosFimAtual + 1
                iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyShift))
                
                objAliquota.iISS = CDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
                
                iPosAtual = iPosFimAtual + 1
                iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyShift))
            
                objAliquota.dValorTotalizadoLoja = CDbl(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
                
                iPosAtual = iPosFimAtual + 1
                iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeyShift))
            
                objUltimaReducao.colAliquotas.Add objAliquota
            
            Loop
    
            If iPosAtual <= Len(sRegistro) Then
                iRegistroCol = iIndice
                Exit For
            Else
                iPosAtual = 1
            End If
                
        
        Next
    
    End If
    
    Movcx_FechamentoCaixa = SUCESSO

    Exit Function

Erro_Movcx_FechamentoCaixa:

    Movcx_FechamentoCaixa = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_ErroECF(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 204516)

    End Select

    Exit Function

End Function



Public Function OrcamentoECF_Le(objVenda As ClassVenda) As Long
'Função Que le um orcamento especifico passado como parametro, objVenda.objCupomFiscal.lNumOrcamento.

Dim lErro As Long
Dim iIndice As Integer
Dim colRegistro As New Collection
Dim sArquivo As String
Dim alComando(1 To 1) As Long
Dim tOrcamento As typeOrcamentoECF
Dim IdExterno As String, NumCaixa As Integer, StatusExterno As Integer, sEndEntEmail As String
Dim sEndEntLogradouro As String, sEndEntNumero As String, sEndEntComplemento As String, sEndEntBairro As String, sEndEntCidade As String, sEndEntUF As String, lEndEntIBGECidade As Long

On Error GoTo Erro_OrcamentoECF_Le
    
    
   'Abre os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_AbrirExt(glConexaoOrcPAF)
        If alComando(iIndice) = 0 Then gError 204687
    Next
    
    tOrcamento.sMsg = String(STRING_ORCAMENTO_MSG, 0)
    IdExterno = String(255, 0)
    sEndEntLogradouro = String(255, 0)
    sEndEntNumero = String(255, 0)
    sEndEntComplemento = String(255, 0)
    sEndEntBairro = String(255, 0)
    sEndEntCidade = String(255, 0)
    sEndEntUF = String(255, 0)
    sEndEntEmail = String(255, 0)
    
    lErro = Comando_Executar(alComando(1), "SELECT Seq, Msg, Status, DAVImpresso, IdExterno, NumCaixa, StatusExterno, EndEntLogradouro, EndEntNumero, EndEntComplemento, EndEntBairro, EndEntCidade, EndEntUF, EndEntIBGECidade, EndEntEmail FROM Orcamento WHERE Codigo = ? ORDER BY Seq", tOrcamento.iSeq, tOrcamento.sMsg, tOrcamento.iStatus, tOrcamento.iDAVImpresso, IdExterno, NumCaixa, StatusExterno, _
        sEndEntLogradouro, sEndEntNumero, sEndEntComplemento, sEndEntBairro, sEndEntCidade, sEndEntUF, lEndEntIBGECidade, sEndEntEmail, objVenda.objCupomFiscal.lNumOrcamento)
    If lErro <> AD_SQL_SUCESSO Then gError 204688
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 204689
    
    If lErro <> AD_SQL_SUCESSO Then gError 204690
    
    Do While lErro = AD_SQL_SUCESSO
    
        colRegistro.Add tOrcamento.sMsg

        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 210527

    Loop
    
    Set objVenda = New ClassVenda

    With objVenda.objCupomFiscal
        .IdExterno = IdExterno
        .NumCaixa = NumCaixa
        .StatusExterno = StatusExterno
        .sEndEntLogradouro = sEndEntLogradouro
        .sEndEntNúmero = sEndEntNumero
        .sEndEntComplemento = sEndEntComplemento
        .sEndEntBairro = sEndEntBairro
        .sEndEntCidade = sEndEntCidade
        .sEndEntUF = sEndEntUF
        .lEndEntIBGECidade = lEndEntIBGECidade
        .sEndEntEmail = sEndEntEmail
    End With

    'passa os dados de sMsg para objVenda
    lErro = CF_ECF("Registro_ECF_Vendas", colRegistro, sArquivo, 1, objVenda)
    If lErro <> SUCESSO Then gError 204691
    
    If tOrcamento.iStatus = STATUS_BAIXADO Then gError 210447
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    
    OrcamentoECF_Le = SUCESSO
    
    Exit Function
    
Erro_OrcamentoECF_Le:
    
    OrcamentoECF_Le = gErr
    lErro = gErr
    
    Select Case lErro
            
        Case 204687
            Call Rotina_ErroECF(vbOKOnly, ERRO_ABERTURA_COMANDO, gErr)
    
        Case 204688, 204689, 210527
            Call Rotina_ErroECF(vbOKOnly, ERRO_LEITURA_ORCAMENTO, gErr)
            
        Case 204690, 204691, 210447
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, lErro, Error$, 204692)

    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Function ArquivoSeq_Le(colArqSeq As Collection) As Long
'retorna todos os registro de ArquivoSeq

Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim lErro As Long
Dim objArqSeq As ClassArqSeq
Dim lSequencial As Long
Dim lUltimoNumIntDoc As Long

On Error GoTo Erro_ArquivoSeq_Le


    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_AbrirExt(glConexaoPAFECF)
        If alComando(iIndice) = 0 Then gError 204758
    Next

    lErro = Comando_Executar(alComando(1), "SELECT TOP 1000 Sequencial, UltimoNumMovto FROM ArquivoSeq ORDER BY Sequencial DESC", lSequencial, lUltimoNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 204759
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 204760

    Do While lErro = AD_SQL_SUCESSO
        
        Set objArqSeq = New ClassArqSeq
        
        objArqSeq.lSequencial = lSequencial
        objArqSeq.lUltimoNumIntDoc = lUltimoNumIntDoc
        
        colArqSeq.Add objArqSeq
        
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 204761
        
    Loop
    
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next
    
    ArquivoSeq_Le = SUCESSO

    Exit Function

Erro_ArquivoSeq_Le:

    ArquivoSeq_Le = gErr

    Select Case gErr

        Case 204758
            Call Rotina_ErroECF(vbOKOnly, ERRO_ABERTURA_COMANDO, gErr)

        Case 204759, 204760, 204761
            Call Rotina_ErroECF(vbOKOnly, ERRO_LEITURA_ARQUIVOSEQ, gErr)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error, 204762)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function
    
    
Function MeioPagamento_Executa_RelGer(ByVal dtDataDe As Date, ByVal dtDataAte As Date, objTela As Object) As Long

Dim lErro As Long
Dim lSequencial As Long

On Error GoTo Erro_MeioPagamento_Executa_RelGer

    'Função que Abre a Transação
    lErro = CF_ECF("Caixa_Transacao_Abrir", lSequencial)
    If lErro <> SUCESSO Then gError 207280

    'Função que Referenciar a Função da Afrac que Vai Executar a Leitura
    lErro = CF_ECF("MeioPagamento_Executa_RelGerEmTrans", lSequencial, dtDataDe, dtDataAte, objTela)
    If lErro <> SUCESSO Then gError 207281

    'Função que Encerra a Transação com o Caixa
    lErro = CF_ECF("Caixa_Transacao_Fechar", lSequencial)
    If lErro <> SUCESSO Then gError 207282

    MeioPagamento_Executa_RelGer = SUCESSO

    Exit Function

Erro_MeioPagamento_Executa_RelGer:

    MeioPagamento_Executa_RelGer = gErr

     Select Case gErr

        Case 27280, 207281, 207282

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 207283)

    End Select
    
'    lSequencial = glSeqTransacaoAberta
            
    'Rollback na transação
    Call CF_ECF("Caixa_Transacao_Rollback", glTransacaoPAFECF)

    Exit Function

End Function

Function MeioPagamento_Executa_RelGerEmTrans(lSequencial As Long, ByVal dtDataDe As Date, ByVal dtDataAte As Date, objTela As Object) As Long

Dim lErro As Long
Dim objMovimentoCaixa As New ClassMovimentoCaixa
Dim sLog As String
Dim colRegistro As New Collection
Dim sMensagem As String
Dim colOrcamento As Collection
Dim objVenda As ClassVenda
Dim iIndice As Integer
Dim colMensagem As New Collection
Dim iImprimiu As Integer
Dim dtData As Date
Dim alComando(1 To 1) As Long
Dim tMeioPagamento As typeMeioPagamento
Dim dTotalMeioPagamento As Double
Dim sMeioPagamento As String
Dim objAdmMeioPagto As ClassAdmMeioPagto
Dim sMeioPagamento1 As String

On Error GoTo Erro_MeioPagamento_Executa_RelGerEmTrans

'    If giOrcamentoECF <> CAIXA_SO_ORCAMENTO Then
'
'        For iIndice = LBound(alComando) To UBound(alComando)
'            alComando(iIndice) = Comando_AbrirExt(glConexaoPAFECF)
'            If alComando(iIndice) = 0 Then gError 207284
'        Next
'
'        'Função que Move os Dados para a Memória
'        lErro = CF_ECF("MovimentoCaixa_Move_Dados_Memoria", MOVIMENTOCAIXA_MEIOPAGTO_RELGER, objMovimentoCaixa)
'        If lErro <> SUCESSO Then gError 207285
'
'        'Guarda o Sequencial no objMovimentos Caixa
'        objMovimentoCaixa.lSequencial = lSequencial
'
'        Call CF_ECF("MovimentoCaixa_Gera_Log", colRegistro, objMovimentoCaixa)
'
'        'Guarda as Informações no Arquivo
'        lErro = CF_ECF("MovimentoCaixaECF_Grava", colRegistro)
'        If lErro <> SUCESSO Then gError 207286
'
'
'        lErro = AFRAC_AbrirRelatorioGerencial(RELGER_MEIOPAGAMENTO, objTela)
'        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Rel Ger Meios de Pagamento")
'        If lErro <> SUCESSO Then gError 207287
'
'        colMensagem.Add Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "MEIOS DE PAGAMENTO")
'        colMensagem.Add Space(48)
'        colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "Periodo De: " & Format(dtDataDe, "dd/mm/yyyy") & " a " & Format(dtDataAte, "dd/mm/yyyy"))
'        colMensagem.Add Space(48)
'        colMensagem.Add "       Meio            Tipo   Valor       Data"
'
'        For iIndice = 1 To colMensagem.Count
'
'            sMensagem = colMensagem.Item(iIndice)
'
'            lErro = AFRAC_ImprimirRelatorioGerencial(sMensagem, objTela)
'            lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Meios de Pagamento")
'            If lErro <> SUCESSO Then gError 207288
'
'        Next
'
'        tMeioPagamento.sMeioPagamento = String(STRING_MEIOPAGAMENTO_MEIOPAGAMENTO, 0)
'        tMeioPagamento.sTipoDocumento = String(STRING_MEIOPAGAMENTO_TIPODOCUMENTO, 0)
'
'        lErro = Comando_Executar(alComando(1), "SELECT MeioPagamento, TipoDocumento, Valor, DataMovimento, AdmMeioPagto, TipoMeioPagto FROM MeioPagamento WHERE DataMovimento >= ? AND DataMovimento <= ? ORDER BY MeioPagamento, AdmMeioPagto, DataMovimento", tMeioPagamento.sMeioPagamento, tMeioPagamento.sTipoDocumento, tMeioPagamento.dValor, tMeioPagamento.dtDataMovimento, tMeioPagamento.iAdmMeioPagto, tMeioPagamento.iTipoMeioPagto, dtDataDe, dtDataAte)
'        If lErro <> AD_SQL_SUCESSO Then gError 207289
'
'        lErro = Comando_BuscarPrimeiro(alComando(1))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 207290
'
'        sMeioPagamento = tMeioPagamento.sMeioPagamento
'
'        sMeioPagamento1 = sMeioPagamento
'
'        Do While lErro = AD_SQL_SUCESSO
'
'            If sMeioPagamento <> sMeioPagamento1 Then
'
'                Set colMensagem = New Collection
'                colMensagem.Add Space(48)
'                colMensagem.Add "Total " & sMeioPagamento & " = " & Format(dTotalMeioPagamento, "standard")
'                colMensagem.Add Space(48)
'                colMensagem.Add "       Meio            Tipo   Valor       Data"
'
'                For iIndice = 1 To colMensagem.Count
'
'                    sMensagem = colMensagem.Item(iIndice)
'
'                    lErro = AFRAC_ImprimirRelatorioGerencial(sMensagem, objTela)
'                    lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Meios de Pagamento")
'                    If lErro <> SUCESSO Then gError 207291
'
'                Next
'
'                dTotalMeioPagamento = 0
'                sMeioPagamento = sMeioPagamento1
'
'            End If
'
'
'            sMensagem = Formata_Campo(ALINHAMENTO_DIREITA, 24, " ", right(sMeioPagamento1, 24)) & tMeioPagamento.sTipoDocumento & Formata_Campo(ALINHAMENTO_ESQUERDA, 12, " ", Format(tMeioPagamento.dValor, "standard")) & " " & Format(tMeioPagamento.dtDataMovimento, "dd/mm/yy")
'
'            lErro = AFRAC_ImprimirRelatorioGerencial(sMensagem, objTela)
'            lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "DAV Emitidos")
'            If lErro <> SUCESSO Then gError 207292
'
'            iImprimiu = 1
'
'            dTotalMeioPagamento = dTotalMeioPagamento + tMeioPagamento.dValor
'
'            lErro = Comando_BuscarProximo(alComando(1))
'            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 207293
'
'
'            sMeioPagamento1 = tMeioPagamento.sMeioPagamento
'
'
'        Loop
'
'        If iImprimiu = 0 Then
'
'            sMensagem = "******* NÃO HÁ RECEBIMENTOS NO PERIODO SOLICITADO *******"
'
'            lErro = AFRAC_ImprimirRelatorioGerencial(sMensagem, objTela)
'            lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Meios de Pagamento")
'            If lErro <> SUCESSO Then gError 207294
'
'        Else
'
'            Set colMensagem = New Collection
'            colMensagem.Add Space(48)
'            colMensagem.Add "Total " & sMeioPagamento1 & " = " & Format(dTotalMeioPagamento, "standard")
'            colMensagem.Add Space(48)
'            colMensagem.Add String(48, "-")
'
'            For iIndice = 1 To colMensagem.Count
'
'                sMensagem = colMensagem.Item(iIndice)
'
'                lErro = AFRAC_ImprimirRelatorioGerencial(sMensagem, objTela)
'                lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "DAV Emitidos")
'                If lErro <> SUCESSO Then gError 207296
'
'            Next
'
'        End If
'
'        lErro = AFRAC_FecharRelatorioGerencial(objTela)
'        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Fecha Orcamento")
'        If lErro <> SUCESSO Then gError 207295
'
'    Else
'
'        Call CF_ECF("Trata_Caixa_So_Orcamento")
'
'    End If
'
'
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next


    MeioPagamento_Executa_RelGerEmTrans = SUCESSO
    
    Exit Function

Erro_MeioPagamento_Executa_RelGerEmTrans:

    MeioPagamento_Executa_RelGerEmTrans = gErr

        Select Case gErr
        
            Case 207284
                Call Rotina_ErroECF(vbOKOnly, ERRO_ABERTURA_COMANDO, gErr)
        
            Case 207285, 206286, 207287, 207295, 207286
        
            Case 207288, 207291, 207292, 207294, 207296
                Call AFRAC_FecharRelatorioGerencial(objTela)
        
            Case 207289, 207290, 207293
                Call Rotina_ErroECF(vbOKOnly, ERRO_LEITURA_MEIOPAGAMENTO, gErr)
                
            Case Else
                Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 207297)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function


Function Requisito_XLIII(objTela As Object) As Long

Dim lErro As Long
Dim lSequencial As Long

On Error GoTo Erro_Requisito_XLIII

    'Função que Abre a Transação
    lErro = CF_ECF("Caixa_Transacao_Abrir", lSequencial)
    If lErro <> SUCESSO Then gError 207316

    'Função que Referenciar a Função da Afrac que Vai Executar a Leitura
    lErro = CF_ECF("Requisito_XLIII_EmTrans", lSequencial, objTela)
    If lErro <> SUCESSO Then gError 207317

    'Função que Encerra a Transação com o Caixa
    lErro = CF_ECF("Caixa_Transacao_Fechar", lSequencial)
    If lErro <> SUCESSO Then gError 207318

    Requisito_XLIII = SUCESSO

    Exit Function

Erro_Requisito_XLIII:

    Requisito_XLIII = gErr

     Select Case gErr

        Case 207316 To 207318

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 207319)

    End Select
    
'    lSequencial = glSeqTransacaoAberta
            
    'Rollback na transação
    Call CF_ECF("Caixa_Transacao_Rollback", glTransacaoPAFECF)

    Exit Function

End Function

Function Requisito_XLIII_EmTrans(lSequencial As Long, objTela As Object) As Long
'Requisito XLIII

Dim lErro As Long
Dim objMovimentoCaixa As New ClassMovimentoCaixa
Dim sLog As String
Dim colRegistro As New Collection
Dim sMensagem As String
Dim colOrcamento As Collection
Dim objVenda As ClassVenda
Dim iIndice As Integer
Dim colMensagem As New Collection
Dim iImprimiu As Integer
Dim dtData As Date
Dim alComando(1 To 1) As Long
Dim dTotalMeioPagamento As Double
Dim sMeioPagamento As String
Dim objConfiguracaoECF As New ClassConfiguracaoECF
Dim sMD5PrincExec As String
Dim sNomeCompleto As String
Dim sMD5 As String
Dim objECFCorporator As New ClassECFCorporator
Dim objECFAutorizado As ClassECFAutorizado
Dim sNome As String

On Error GoTo Erro_Requisito_XLIII_EmTrans

    If giOrcamentoECF <> CAIXA_SO_ORCAMENTO Then

        For iIndice = LBound(alComando) To UBound(alComando)
            alComando(iIndice) = Comando_AbrirExt(glConexaoPAFECF)
            If alComando(iIndice) = 0 Then gError 207320
        Next
    
        'Função que Move os Dados para a Memória
        lErro = CF_ECF("MovimentoCaixa_Move_Dados_Memoria", MOVIMENTOCAIXA_IDENTPAFECF_RELGER, objMovimentoCaixa)
        If lErro <> SUCESSO Then gError 207321
        
        'Guarda o Sequencial no objMovimentos Caixa
        objMovimentoCaixa.lSequencial = lSequencial
        
        Call CF_ECF("MovimentoCaixa_Gera_Log", colRegistro, objMovimentoCaixa)
        
        'Guarda as Informações no Arquivo
        lErro = CF_ECF("MovimentoCaixaECF_Grava", colRegistro)
        If lErro <> SUCESSO Then gError 207322
        
        lErro = CF_ECF("ConfiguracaoECF_Le", objConfiguracaoECF)
        If lErro <> SUCESSO Then gError 207934
        
        sNomeCompleto = objConfiguracaoECF.sDirPAFECF & "\" & objConfiguracaoECF.sPrincExec
        
        lErro = AFRAC_MD5(sNomeCompleto, sMD5PrincExec)
        If lErro <> SUCESSO Then gError 207935
    
        colMensagem.Add String(48, "-")
        colMensagem.Add Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "IDENTIFICAÇÃO DO PAF-ECF")
        colMensagem.Add Space(48)
        colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "N. do Laudo: " & objConfiguracaoECF.sNumLaudo)
        colMensagem.Add Space(48)
        colMensagem.Add Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "Empresa Desenvolvedora")
        colMensagem.Add Space(48)
        colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "CNPJ: " & objConfiguracaoECF.sCNPJDesenv)
        colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "Razão Social: " & objConfiguracaoECF.sRazaoSocialDesenv)
        colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "Endereço: " & left(objConfiguracaoECF.sEndDesenv, 38))
        If Len(Mid(objConfiguracaoECF.sEndDesenv, 39, 48)) > 0 Then colMensagem.Add Mid(objConfiguracaoECF.sEndDesenv, 39, 48)
        colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "Telefone: " & objConfiguracaoECF.sTelDesenv)
        colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "Contato: " & objConfiguracaoECF.sContatoDesenv)
        colMensagem.Add Space(48)
        colMensagem.Add Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "PAF-ECF")
        colMensagem.Add Space(48)
        colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "Nome Comercial: " & objConfiguracaoECF.sNomePAFECF)
        colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "Versão: " & objConfiguracaoECF.sVersaoPAFECF)
        colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "Princ. Arq. Exec.: " & objConfiguracaoECF.sPrincExec)
        colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "MD5 Princ Arq: " & sMD5PrincExec)
        colMensagem.Add Space(48)
        colMensagem.Add Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "Demais Arquivos")
        colMensagem.Add Space(48)
        
        
        sNome = String(ARQUIVOECF_NOME, 0)
        
        lErro = Comando_Executar(alComando(1), "SELECT Nome FROM ArquivoECF ORDER BY Nome", sNome)
        If lErro <> AD_SQL_SUCESSO Then gError 207936
    
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 207937
    
        Do While lErro = AD_SQL_SUCESSO
        
            sNomeCompleto = objConfiguracaoECF.sDirPAFECF & "\" & sNome
        
            lErro = AFRAC_MD5(sNomeCompleto, sMD5)
            If lErro <> SUCESSO Then gError 207938
        
            colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "Nome: " & sNome)
            colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "MD5: " & sMD5)
        
            lErro = Comando_BuscarProximo(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 207939
        
        Loop
        
        
        
    
        colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "Arqs. Autent.: REQUISITO_XI.txt")
    
        lErro = CF_ECF("EcfCorporator_Descriptografa", objECFCorporator)
        If lErro <> SUCESSO And lErro <> 204505 Then gError 207940
        
    
        colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "MD5: " & objECFCorporator.sMD5PAFECF)
        colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "Versão da ER PAF-ECF: " & objConfiguracaoECF.sVersaoEspReqPAFECF)
    
        colMensagem.Add Space(48)
        colMensagem.Add Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "ECFs Autorizados")
        colMensagem.Add Space(48)
        
        For Each objECFAutorizado In objECFCorporator.colECFAutorizado
        
            colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "Num. Fab.: " & objECFAutorizado.sNumFabricacao)
        
        Next
    
        colMensagem.Add Space(48)
        colMensagem.Add String(48, "-")
        colMensagem.Add Space(48)
        
    
        
        lErro = AFRAC_AbrirRelatorioGerencial(RELGER_IDENTPAFECF, objTela)
        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Identificaçao PAFECF")
        If lErro <> SUCESSO Then gError 207323
    
        For iIndice = 1 To colMensagem.Count
                
            sMensagem = colMensagem.Item(iIndice)
    
            lErro = AFRAC_ImprimirRelatorioGerencial(sMensagem, objTela)
            lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "DAV Emitidos")
            If lErro <> SUCESSO Then gError 207324
            
        Next
    
        lErro = AFRAC_FecharRelatorioGerencial(objTela)
        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Fecha Orcamento")
        If lErro <> SUCESSO Then gError 207325
    
        
    
        For iIndice = LBound(alComando) To UBound(alComando)
            Call Comando_Fechar(alComando(iIndice))
        Next
    
    Else
        
        Call CF_ECF("Trata_Caixa_So_Orcamento")
        
    End If

    Requisito_XLIII_EmTrans = SUCESSO
    
    Exit Function

Erro_Requisito_XLIII_EmTrans:

    Requisito_XLIII_EmTrans = gErr

        Select Case gErr
        
            Case 207320
                Call Rotina_ErroECF(vbOKOnly, ERRO_ABERTURA_COMANDO, gErr)
        
            Case 207321, 207322, 207323, 207325, 207921, 207934, 207935, 207938, 207940
        
            Case 207936, 207937, 207939
                Call Rotina_ErroECF(vbOKOnly, ERRO_LEITURA_ARQUIVOECF, gErr)
        
            Case 207325
                Call AFRAC_FecharRelatorioGerencial(objTela)
        
            Case Else
                Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 207326)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function ConfiguracaoECF_Le(ByVal objConfiguracaoECF As ClassConfiguracaoECF) As Long
'carrega as configuracoes q estao armazenadas no BD

Dim alComando(1 To 1) As Long
Dim iIndice As Integer
Dim lErro As Long
Dim sConteudo As String
Dim tConfiguracaoECF As typeConfiguracaoECF

On Error GoTo Erro_ConfiguracaoECF_Le

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_AbrirExt(glConexaoPAFECF)
        If alComando(iIndice) = 0 Then gError 207916
    Next

    tConfiguracaoECF.sCNPJDesenv = String(CONFECF_CNPJDESENV, 0)
    tConfiguracaoECF.sIEDesenv = String(CONFECF_IEDESENV, 0)
    tConfiguracaoECF.sIMDesenv = String(CONFECF_IMDESENV, 0)
    tConfiguracaoECF.sRazaoSocialDesenv = String(CONFECF_RAZAOSOCIALDESENV, 0)
    tConfiguracaoECF.sNomePAFECF = String(CONFECF_NOMEPAFECF, 0)
    tConfiguracaoECF.sVersaoPAFECF = String(CONFECF_VERSAOPAFECF, 0)
    tConfiguracaoECF.sEndDesenv = String(CONFECF_ENDDESENV, 0)
    tConfiguracaoECF.sTelDesenv = String(CONFECF_TELDESENV, 0)
    tConfiguracaoECF.sContatoDesenv = String(CONFECF_CONTATODESENV, 0)
    tConfiguracaoECF.sNumLaudo = String(CONFECF_NUMLAUDO, 0)
    tConfiguracaoECF.sVersaoEspReqPAFECF = String(CONFECF_VERSAOESPREQPAFECF, 0)
    tConfiguracaoECF.sPrincExec = String(CONFECF_PRINCEXEC, 0)
    tConfiguracaoECF.sDirPAFECF = String(CONFECF_DIRPAFECF, 0)
    
    
    lErro = Comando_Executar(alComando(1), "SELECT CNPJDesenv, IEDesenv, IMDesenv, RazaoSocialDesenv, NomePAFECF, VersaoPAFECF, EndDesenv, TelDesenv, ContatoDesenv, NumLaudo, VersaoEspReqPAFECF, PrincExec, DirPAFECF FROM ConfiguracaoECF ", _
    tConfiguracaoECF.sCNPJDesenv, tConfiguracaoECF.sIEDesenv, tConfiguracaoECF.sIMDesenv, tConfiguracaoECF.sRazaoSocialDesenv, tConfiguracaoECF.sNomePAFECF, tConfiguracaoECF.sVersaoPAFECF, tConfiguracaoECF.sEndDesenv, _
    tConfiguracaoECF.sTelDesenv, tConfiguracaoECF.sContatoDesenv, tConfiguracaoECF.sNumLaudo, tConfiguracaoECF.sVersaoEspReqPAFECF, tConfiguracaoECF.sPrincExec, tConfiguracaoECF.sDirPAFECF)
    If lErro <> AD_SQL_SUCESSO Then gError 214770

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 207918

    If lErro <> AD_SQL_SUCESSO Then gError 207919
    
    objConfiguracaoECF.sCNPJDesenv = tConfiguracaoECF.sCNPJDesenv
    objConfiguracaoECF.sIEDesenv = tConfiguracaoECF.sIEDesenv
    objConfiguracaoECF.sIMDesenv = tConfiguracaoECF.sIMDesenv
    objConfiguracaoECF.sRazaoSocialDesenv = tConfiguracaoECF.sRazaoSocialDesenv
    objConfiguracaoECF.sNomePAFECF = tConfiguracaoECF.sNomePAFECF
    objConfiguracaoECF.sVersaoPAFECF = tConfiguracaoECF.sVersaoPAFECF
    objConfiguracaoECF.sEndDesenv = tConfiguracaoECF.sEndDesenv
    objConfiguracaoECF.sTelDesenv = tConfiguracaoECF.sTelDesenv
    objConfiguracaoECF.sContatoDesenv = tConfiguracaoECF.sContatoDesenv
    objConfiguracaoECF.sNumLaudo = tConfiguracaoECF.sNumLaudo
    objConfiguracaoECF.sVersaoEspReqPAFECF = tConfiguracaoECF.sVersaoEspReqPAFECF
    objConfiguracaoECF.sPrincExec = tConfiguracaoECF.sPrincExec
    objConfiguracaoECF.sDirPAFECF = tConfiguracaoECF.sDirPAFECF
    
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next
    
    ConfiguracaoECF_Le = SUCESSO

    Exit Function

Erro_ConfiguracaoECF_Le:

    ConfiguracaoECF_Le = gErr

    Select Case gErr

        Case 207916
            Call Rotina_ErroECF(vbOKOnly, ERRO_ABERTURA_COMANDO, gErr)

        Case 207918, 214770
            Call Rotina_ErroECF(vbOKOnly, ERRO_LEITURA_CONFIGURACAOECF, gErr)

        Case 207919
            Call Rotina_ErroECF(vbOKOnly, ERRO_CONFIGURACAOECF_NAO_CADASTRADO, gErr)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error, 207920)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function EcfCorporator_Descriptografa(objECFCorporator As ClassECFCorporator) As Long

Dim lErro As Long
Dim sRegistro As String
Dim lTotal As Long
Dim icol As Integer
Dim iLinha As Integer
Dim sArquivo As String
Dim sArquivo1 As String
Dim sDir As String
Dim lTentaAcessoArquivo As Long
Dim lChave As Long
Dim iPos As Integer
Dim iAdd As Integer
Dim iIncremento As Integer
Dim sMensagem As String
Dim objECFAutorizado As ClassECFAutorizado

On Error GoTo Erro_EcfCorporator_Descriptografa

    Set objECFCorporator.colECFAutorizado = New Collection

    sArquivo = gsDirMVTEF & "ecfcorporator.crp"

    sArquivo1 = Dir(sArquivo)

    If Len(sArquivo1) = 0 Then gError 204504

    iLinha = 0
    
    Open sArquivo For Input Lock Read Write As #1

    Do While Not EOF(1)

        iAdd = 10
        iIncremento = -1
        iPos = 1
        sMensagem = ""

        'Busca o próximo registro do arquivo
        Line Input #1, sRegistro

        iLinha = iLinha + 1

        If Mid(sRegistro, 1, 6) = "chave=" Then Exit Do

        Do While iPos <= Len(sRegistro)
            
            sMensagem = sMensagem & Chr(Asc(Mid(sRegistro, iPos, 1)) - iAdd)
    
            iPos = iPos + 1
    
            If iAdd = 1 Then
                iIncremento = 1
            ElseIf iAdd = 10 Then
                iIncremento = -1
            End If
            
            iAdd = iAdd + iIncremento
        
        Loop

        If iLinha = 1 Then objECFCorporator.sMD5PAFECF = sMensagem
        
        If iLinha > 1 Then
            
            If iLinha Mod 2 = 0 Then
                
                Set objECFAutorizado = New ClassECFAutorizado
                objECFAutorizado.sNumFabricacao = sMensagem
            Else
            
                objECFAutorizado.sGT = sMensagem
                objECFCorporator.colECFAutorizado.Add objECFAutorizado
            
            End If
            
        End If

        For icol = 1 To Len(sMensagem)
            lTotal = lTotal + Asc(Mid(sMensagem, icol, 1)) * icol
        Next

    Loop

    If lTotal = 0 Then gError 214868

    lChave = StrParaLong(Mid(sRegistro, 7))

    If lChave <> lTotal Then gError 204505
    
    Close #1
        
    EcfCorporator_Descriptografa = SUCESSO

    Exit Function

Erro_EcfCorporator_Descriptografa:

    EcfCorporator_Descriptografa = gErr
    
    Close #1
    
    Select Case gErr

        Case 70
            lTentaAcessoArquivo = lTentaAcessoArquivo + 1
            If lTentaAcessoArquivo < NUM_MAX_LOCK Then Resume
            Call Rotina_ErroECF(vbOKOnly, ERRO_ARQUIVO_LOCADO, gErr, sArquivo)
            
            
        Case 204503
        
        Case 204504
            Call Rotina_ErroECF(vbOKOnly, ERRO_ARQ_CRIPTOGRAFADO_INEXISTENTE, gErr, sArquivo)

        Case 204505
'            Call Rotina_ErroECF(vbOKOnly, ERRO_ARQUIVO_ECFCORPORATOR_ALTERADO, gErr)

        Case 214868
            Call Rotina_ErroECF(vbOKOnly, ERRO_ARQ_CRIPTOGRAFADO_VAZIO, gErr, sArquivo)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 204506)

    End Select
    
    
    
    Exit Function

End Function

Function EcfCorporator_Criptografa(objECFCorporator As ClassECFCorporator) As Long

Dim lErro As Long
Dim sRegistro As String
Dim lTotal As Long
Dim icol As Integer
Dim iLinha As Integer
Dim sArquivo As String
Dim sArquivo1 As String
Dim sDir As String
Dim lTentaAcessoArquivo As Long
Dim lChave As Long
Dim iPos As Integer
Dim iAdd As Integer
Dim iIncremento As Integer
Dim sMensagem As String
Dim objECFAutorizado As ClassECFAutorizado
Dim sArqTmp As String
Dim sArqTmp1 As String
Dim iIndice As Integer
Dim sArqErro As String

On Error GoTo Erro_EcfCorporator_Criptografa

    sArqErro = sArqTmp

    sArqTmp = gsDirMVTEF & "ecfcorporator.tmp"

    sArqTmp1 = Dir(sArqTmp)

    If Len(sArqTmp1) <> 0 Then Kill sArqTmp

    Open sArqTmp For Append Lock Read Write As #1

    sRegistro = objECFCorporator.sMD5PAFECF
    
    iAdd = 10
    iIncremento = -1
    iPos = 1
    sMensagem = ""

    Do While iPos <= Len(sRegistro)
        
        sMensagem = sMensagem & Chr(Asc(Mid(sRegistro, iPos, 1)) + iAdd)

        iPos = iPos + 1

        If iAdd = 1 Then
            iIncremento = 1
        ElseIf iAdd = 10 Then
            iIncremento = -1
        End If
        
        iAdd = iAdd + iIncremento
    
    Loop
    
    For icol = 1 To Len(sRegistro)
        lTotal = lTotal + Asc(Mid(sRegistro, icol, 1)) * icol
    Next
    
    Print #1, sMensagem


    For Each objECFAutorizado In objECFCorporator.colECFAutorizado


        For iIndice = 1 To 2

            If iIndice = 1 Then
                sRegistro = objECFAutorizado.sNumFabricacao
            Else
                sRegistro = objECFAutorizado.sGT
            End If
            
            iAdd = 10
            iIncremento = -1
            iPos = 1
            sMensagem = ""
            
            
            Do While iPos <= Len(sRegistro)
                
                sMensagem = sMensagem & Chr(Asc(Mid(sRegistro, iPos, 1)) + iAdd)
        
                iPos = iPos + 1
        
                If iAdd = 1 Then
                    iIncremento = 1
                ElseIf iAdd = 10 Then
                    iIncremento = -1
                End If
                
                iAdd = iAdd + iIncremento
            
            Loop
            
            Print #1, sMensagem
            
            For icol = 1 To Len(sRegistro)
                lTotal = lTotal + Asc(Mid(sRegistro, icol, 1)) * icol
            Next
        
        Next
        
    Next

    
    Print #1, "chave=" & CStr(lTotal)

    Close #1

    sArquivo = gsDirMVTEF & "ecfcorporator.crp"

    sArqErro = sArquivo

    sArquivo1 = Dir(sArquivo)

    If Len(sArquivo1) <> 0 Then Kill sArquivo

    Name sArqTmp As sArquivo
    
    EcfCorporator_Criptografa = SUCESSO

    Exit Function

Erro_EcfCorporator_Criptografa:

    EcfCorporator_Criptografa = gErr

    Close #1

    Select Case gErr

        Case 70
            lTentaAcessoArquivo = lTentaAcessoArquivo + 1
            If lTentaAcessoArquivo < 10 Then Resume
            Call Rotina_ErroECF(vbOKOnly, ERRO_ARQUIVO_LOCADO, gErr, sArqErro)
            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 207922)

    End Select
    
    Exit Function

End Function


Sub MovimentoSPED_Gera_Log(colRegistro As Collection, objMovimentoCaixa As ClassMovimentoCaixa)

Dim lErro As Long
Dim iIndice As Integer
Dim iCont As Integer
Dim cColInformação As New Collection
Dim sLog As String
Dim sTemp As String

On Error GoTo Erro_MovimentoSPED_Gera_Log

    'Monta o Log
    sTemp = TIPOREGISTROECF_PAFECF_SPEDFISCAL & Chr(vbKeyControl) & CStr(objMovimentoCaixa.iTipo) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.dHora) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.dtDataMovimento) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.iGerente) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.iCodOperador) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.lSequencial) & Chr(vbKeyEscape) & _
    CStr(objMovimentoCaixa.iFilialEmpresa) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.dtDataInicio) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.dtDataFim) & Chr(vbKeyEscape)
    
    If Len(Trim(sTemp)) + Len(Trim(sLog)) > MAX_TAM_STRING_MOVIMENTOCAIXA Then
        If Len(Trim(sLog)) <> 0 Then colRegistro.Add sLog
        sLog = ""
    End If
    sLog = sLog & sTemp
    sTemp = ""

    sTemp = sTemp & Chr(vbKeyEnd)
    
    If Len(Trim(sTemp)) + Len(Trim(sLog)) + 1 > MAX_TAM_STRING_MOVIMENTOCAIXA Then
        If Len(Trim(sLog)) <> 0 Then colRegistro.Add sLog
        sLog = ""
    End If
    sLog = sLog & sTemp
    sTemp = ""

    If Len(Trim(sLog)) <> 0 Then colRegistro.Add sLog
    
    Exit Sub

Erro_MovimentoSPED_Gera_Log:

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 210140)

    End Select

    Exit Sub

End Sub

Sub MovimentoSINTEGRA_Gera_Log(colRegistro As Collection, objMovimentoCaixa As ClassMovimentoCaixa)

Dim lErro As Long
Dim iIndice As Integer
Dim iCont As Integer
Dim cColInformação As New Collection
Dim sLog As String
Dim sTemp As String

On Error GoTo Erro_MovimentoSINTEGRA_Gera_Log

    'Monta o Log
    sTemp = TIPOREGISTROECF_PAFECF_SINTEGRA & Chr(vbKeyControl) & CStr(objMovimentoCaixa.iTipo) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.dHora) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.dtDataMovimento) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.iGerente) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.iCodOperador) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.lSequencial) & Chr(vbKeyEscape) & _
    CStr(objMovimentoCaixa.iFilialEmpresa) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.dtDataInicio) & Chr(vbKeyEscape) & CStr(objMovimentoCaixa.dtDataFim) & Chr(vbKeyEscape)
    
    If Len(Trim(sTemp)) + Len(Trim(sLog)) > MAX_TAM_STRING_MOVIMENTOCAIXA Then
        If Len(Trim(sLog)) <> 0 Then colRegistro.Add sLog
        sLog = ""
    End If
    sLog = sLog & sTemp
    sTemp = ""

    sTemp = sTemp & Chr(vbKeyEnd)
    
    If Len(Trim(sTemp)) + Len(Trim(sLog)) + 1 > MAX_TAM_STRING_MOVIMENTOCAIXA Then
        If Len(Trim(sLog)) <> 0 Then colRegistro.Add sLog
        sLog = ""
    End If
    sLog = sLog & sTemp
    sTemp = ""

    If Len(Trim(sLog)) <> 0 Then colRegistro.Add sLog
    
    Exit Sub

Erro_MovimentoSINTEGRA_Gera_Log:

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 210139)

    End Select

    Exit Sub

End Sub


Function R01_Le_ECF(colECF As Collection) As Long
'le todos os numeros de serie diferentes contidos na tabela R1

Dim alComando(1 To 1) As Long
Dim iIndice As Integer
Dim lErro As Long
Dim sConteudo As String
Dim sNumSerie As String

On Error GoTo Erro_R01_Le_ECF

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_AbrirExt(glConexaoPAFECF)
        If alComando(iIndice) = 0 Then gError 210142
    Next

    sNumSerie = String(R01_NUMEROSERIE, 0)
    
    lErro = Comando_Executar(alComando(1), "SELECT DISTINCT NumeroSerie FROM R01 ", sNumSerie)
    If lErro <> AD_SQL_SUCESSO Then gError 210143

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 210144

    Do While lErro = AD_SQL_SUCESSO
        
        colECF.Add sNumSerie
    
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 210145
    
    Loop
    
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next
    
    R01_Le_ECF = SUCESSO

    Exit Function

Erro_R01_Le_ECF:

    R01_Le_ECF = gErr

    Select Case gErr

        Case 210142
            Call Rotina_ErroECF(vbOKOnly, ERRO_ABERTURA_COMANDO, gErr)

        Case 210143 To 210145
            Call Rotina_ErroECF(vbOKOnly, ERRO_LEITURA_R01, gErr)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error, 210146)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function MovimentoCaixa_Le(lNumMovto As Long, dtData As Date) As Long
'devolve o proximo numero de movimento disponivel a partir do q esta cadastrado na tabela MovimentoCaixa

Dim alComando(1 To 1) As Long
Dim iIndice As Integer
Dim lErro As Long


On Error GoTo Erro_MovimentoCaixa_Le

   'Abre os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_AbrirExt(glConexaoPAFECF)
        If alComando(iIndice) = 0 Then gError 210431
    Next

    lErro = Comando_Executar(alComando(1), "SELECT Data, NumMovto FROM MovimentoCaixa  ORDER BY NumMovto DESC", dtData, lNumMovto)
    If lErro <> AD_SQL_SUCESSO Then gError 210432

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 210433

    If lErro = AD_SQL_SEM_DADOS Then
        lNumMovto = 1
    Else
        lNumMovto = lNumMovto + 1
    End If

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    MovimentoCaixa_Le = SUCESSO

    Exit Function
    
Erro_MovimentoCaixa_Le:
    
    MovimentoCaixa_Le = gErr

    Select Case gErr
    
        Case 210431
            Call Rotina_ErroECF(vbOKOnly, ERRO_ABERTURA_COMANDO, gErr)
        
        Case 210432, 210433
            Call Rotina_ErroECF(vbOKOnly, ERRO_LEITURA_MOVIMENTOCAIXA, gErr)
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 210434)
        
    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Public Function OrcamentoECF_Le1(ByVal colOrcamento As Collection) As Long
'Função Que le todos os orcamentos validos

Dim lErro As Long
Dim iIndice As Integer
Dim colRegistro As New Collection
Dim sArquivo As String
Dim alComando(1 To 1) As Long
Dim objVenda As ClassVenda
Dim tOrcamento As typeOrcamentoECF
Dim lCodigo As Long

On Error GoTo Erro_OrcamentoECF_Le1
       
   'Abre os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_AbrirExt(glConexaoOrcPAF)
        If alComando(iIndice) = 0 Then gError 210533
    Next
    
    tOrcamento.sMsg = String(STRING_ORCAMENTO_MSG, 0)
    
    lErro = Comando_Executar(alComando(1), "SELECT Codigo, Seq, Msg FROM Orcamento WHERE Status <> ?", tOrcamento.lCodigo, tOrcamento.iSeq, tOrcamento.sMsg, STATUS_BAIXADO)
    If lErro <> AD_SQL_SUCESSO Then gError 210530
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 210531
    
    lCodigo = tOrcamento.lCodigo
    
    Set colRegistro = New Collection
    
    Do While lErro = AD_SQL_SUCESSO

        If lCodigo <> tOrcamento.lCodigo Then

            Set objVenda = New ClassVenda
    
            lErro = CF_ECF("Registro_ECF_Vendas", colRegistro, sArquivo, 1, objVenda)
            If lErro <> SUCESSO Then gError 210529
    
            colOrcamento.Add objVenda
    
            Set colRegistro = New Collection
    
            lCodigo = tOrcamento.lCodigo
    
    
    '        If objVenda.iTipo = OPTION_PREVENDA Then
    '
    '            lErro = CF_ECF("Imprime_Orcamento", objVenda)
    '            If lErro <> SUCESSO Then gError 204678
    '
    '        End If

        End If

        colRegistro.Add tOrcamento.sMsg

        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 210532
        
    Loop
    
    If colRegistro.Count > 0 Then
    
        Set objVenda = New ClassVenda
    
        lErro = CF_ECF("Registro_ECF_Vendas", colRegistro, sArquivo, 1, objVenda)
        If lErro <> SUCESSO Then gError 210528
    
        colOrcamento.Add objVenda
    
    End If
        
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    OrcamentoECF_Le1 = SUCESSO
    
    Exit Function
    
Erro_OrcamentoECF_Le1:
    
    OrcamentoECF_Le1 = gErr
    lErro = gErr
    
    Select Case lErro
            
        Case 210533
            Call Rotina_ErroECF(vbOKOnly, ERRO_ABERTURA_COMANDO, gErr)
    
        Case 210530 To 210532
            Call Rotina_ErroECF(vbOKOnly, ERRO_LEITURA_ORCAMENTO, gErr)
            
        Case 210528, 210529
        
        Case ERRO_SEM_MENSAGEM
            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, lErro, Error$, 210534)

    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Function Rel_Parametros_Configuracao(objTela As Object) As Long

Dim lErro As Long
Dim lSequencial As Long

On Error GoTo Erro_Rel_Parametros_Configuracao

    'Função que Abre a Transação
    lErro = CF_ECF("Caixa_Transacao_Abrir", lSequencial)
    If lErro <> SUCESSO Then gError 210812

    'Função que Referenciar a Função da Afrac que Vai Executar a Leitura
    lErro = CF_ECF("Rel_Parametros_Configuracao_EmTrans", lSequencial, objTela)
    If lErro <> SUCESSO Then gError 210813

    'Função que Encerra a Transação com o Caixa
    lErro = CF_ECF("Caixa_Transacao_Fechar", lSequencial)
    If lErro <> SUCESSO Then gError 210814

    Rel_Parametros_Configuracao = SUCESSO

    Exit Function

Erro_Rel_Parametros_Configuracao:

    Rel_Parametros_Configuracao = gErr

     Select Case gErr

        Case 210812 To 210814

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 210815)

    End Select
    
'    lSequencial = glSeqTransacaoAberta
            
    'Rollback na transação
    Call CF_ECF("Caixa_Transacao_Rollback", glTransacaoPAFECF)

    Exit Function

End Function

Function Rel_Parametros_Configuracao_EmTrans(lSequencial As Long, objTela As Object) As Long
'Requisito XLIII

Dim lErro As Long
Dim objMovimentoCaixa As New ClassMovimentoCaixa
Dim sLog As String
Dim colRegistro As New Collection
Dim sMensagem As String
Dim colOrcamento As Collection
Dim objVenda As ClassVenda
Dim iIndice As Integer
Dim colMensagem As New Collection
Dim iImprimiu As Integer
Dim dtData As Date
Dim alComando(1 To 1) As Long
Dim dTotalMeioPagamento As Double
Dim sMeioPagamento As String
Dim objConfiguracaoECF As New ClassConfiguracaoECF
Dim sMD5PrincExec As String
Dim sNomeCompleto As String
Dim sMD5 As String
Dim objECFCorporator As New ClassECFCorporator
Dim objECFAutorizado As ClassECFAutorizado
Dim sNome As String
Dim sPerfil As String

On Error GoTo Erro_Rel_Parametros_Configuracao_EmTrans



    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_AbrirExt(glConexaoPAFECF)
        If alComando(iIndice) = 0 Then gError 210816
    Next

    'Função que Move os Dados para a Memória
    lErro = CF_ECF("MovimentoCaixa_Move_Dados_Memoria", MOVIMENTOCAIXA_PARAMCONFIG_RELGER, objMovimentoCaixa)
    If lErro <> SUCESSO Then gError 210817
    
    'Guarda o Sequencial no objMovimentos Caixa
    objMovimentoCaixa.lSequencial = lSequencial
    
    Call CF_ECF("MovimentoCaixa_Gera_Log", colRegistro, objMovimentoCaixa)
    
    'Guarda as Informações no Arquivo
    lErro = CF_ECF("MovimentoCaixaECF_Grava", colRegistro)
    If lErro <> SUCESSO Then gError 210818
    
    Call Afrac_UF_ObtemPerfil(sPerfil)
 
    colMensagem.Add Space(48)
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "Perfil de Requisitos: " & sPerfil)
    colMensagem.Add Space(48)
    colMensagem.Add String(48, "-")
    colMensagem.Add Space(48)
    

    
    lErro = AFRAC_AbrirRelatorioGerencial(RELGER_PARAMCONFIG, objTela)
    lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Parametros de Configuração")
    If lErro <> SUCESSO Then gError 210819

    For iIndice = 1 To colMensagem.Count
            
        sMensagem = colMensagem.Item(iIndice)

        lErro = AFRAC_ImprimirRelatorioGerencial(sMensagem, objTela)
        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Parametros de Configuração")
        If lErro <> SUCESSO Then gError 210820
        
    Next

    lErro = AFRAC_FecharRelatorioGerencial(objTela)
    lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Parametros de Configuração")
    If lErro <> SUCESSO Then gError 210821

    

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Rel_Parametros_Configuracao_EmTrans = SUCESSO
    
    Exit Function

Erro_Rel_Parametros_Configuracao_EmTrans:

    Rel_Parametros_Configuracao_EmTrans = gErr

        Select Case gErr
        
            Case 210816
                Call Rotina_ErroECF(vbOKOnly, ERRO_ABERTURA_COMANDO, gErr)
        
            Case 210817 To 210820
        
            Case 210821
                Call AFRAC_FecharRelatorioGerencial(objTela)
        
            Case Else
                Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 210822)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function SAT_Imprime(ByVal objTela As Object, ByVal objVenda As ClassVenda) As Long
'imprime o SAT

Dim lErro As Long
Dim colMensagem As New Collection
Dim iIndice As Integer
Dim sMensagem As String

On Error GoTo Erro_SAT_Imprime
    
    lErro = AFRAC_AbrirRelatorioGerencial(RELGER_SAT, objTela)
    lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Abrir SAT")
    If lErro <> SUCESSO Then gError 133088

    Select Case gobjSATInfo.iModeloImpressora
        
        Case IMPRESSORA_DARUMA_DR700
        
            Call iImprimirCFe_SAT_Daruma(objVenda.objCupomFiscal.sSATArqXml, "1")

            'lErro = SAT_Imprime_Daruma(objVenda)
            'If lErro <> SUCESSO Then gError 133088
    
        Case IMPRESSORA_ELGIN_VOX, IMPRESSORA_BEMATECH_ESCBEMA, IMPRESSORA_EPSON_ESCPOS, IMPRESSORA_DARUMA_ESCPOS
        
            lErro = ImpressoraESC_ImprimirSAT(objVenda)
        
        Case Else
        
            lErro = SAT_Imprime1(objVenda, colMensagem)
            If lErro <> SUCESSO Then gError 133088
    
            For iIndice = 1 To colMensagem.Count
                    
                sMensagem = colMensagem.Item(iIndice)
                If sMensagem = "" Then sMensagem = "    "
                lErro = AFRAC_ImprimirRelatorioGerencial(sMensagem, objTela)
                lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "SAT")
                If lErro <> SUCESSO Then gError 99893
                    
            Next

    End Select
    
    lErro = AFRAC_FecharRelatorioGerencial(objTela)
    lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Fecha SAT")
    If lErro <> SUCESSO Then gError 99898
    
    SAT_Imprime = SUCESSO
    
    Exit Function
    
Erro_SAT_Imprime:

    SAT_Imprime = gErr

    Select Case gErr

        Case 99803, 99894 To 99898, 105868, 133088, 204319, 204324, 204329, 210470

        Case 99893
            Call AFRAC_FecharRelatorioGerencial(objTela)

        Case 210471
            Call Rotina_ErroECF(vbOKOnly, ERRO_ORCAMENTO_NAO_PREENCHIDO, lErro)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149439)

    End Select

    Exit Function

End Function

Private Function SAT_Imprime1(ByVal objVenda As ClassVenda, colMensagem As Collection) As Long
'http://www.desenvolvedoresdaruma.com.br/home/downloads/Site_2011/Help/DarumaFrameworkHelpOnline/DarumaFramework/Mini_Impressora/Trabalhando_com_C%C3%B3digo_de_Barras_-_Impressora_N%C3%A3o-Fiscal.htm

Dim lErro As Long
Dim objItem As ClassItemCupomFiscal
Dim objProduto As ClassProduto
Dim iIndice As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim bCont As Boolean
Dim sMensagem As String
Dim dTotal As Double
Dim dPreco As Double
Dim bAchou As Boolean
Dim colOrcamento As New Collection
Dim iIndice1 As Integer
Dim objCliente As ClassCliente
Dim lIndice As Long
Dim sCliente As String
Dim sCPFCGC As String
Dim sVendedor As String
Dim objVendedor As ClassVendedor
Dim sProduto As String
Dim sCNPJ As String
Dim iPos As Integer
Dim lNumDAV As Long, objMovCaixa As ClassMovimentoCaixa
Dim sRetorno As String, sSATChaveAcesso As String
Dim objProdutoNomeRed As TextBox, sSATtimeStamp As String
Dim sTexto As String, sTexto2 As String, iTam As Integer

On Error GoTo Erro_SAT_Imprime1

    Set colMensagem = New Collection
    
    Set objCliente = gobjClienteCPF.Busca(objVenda.objCupomFiscal.sCPFCGC, lIndice)
    If Not objCliente Is Nothing Then sCliente = objCliente.sNomeReduzido & " - "
    Call Formata_CNPJ_CPF(sCPFCGC, objVenda.objCupomFiscal.sCPFCGC)
    sCliente = sCliente & sCPFCGC
    
    Call Formata_CNPJ_CPF(sCNPJ, gsCNPJ)
    
    colMensagem.Add Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", gsNomeReduzido)
    colMensagem.Add Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", gsNomeEmpresa)
    colMensagem.Add Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", gsEndereco)
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "")
    colMensagem.Add Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "CNPJ:" & sCNPJ)
    colMensagem.Add Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "IE:" & gsInscricaoEstadual)
    colMensagem.Add Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "IM:" & gsInscricaoMunicipal)
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", "")
    
    If objVenda.objCupomFiscal.iSATtpAmb <> 1 Then
    
        colMensagem.Add Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "Extrato No. 000000")
        colMensagem.Add Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "CUPOM FISCAL ELETRÔNICO - SAT")
        colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "")
        colMensagem.Add Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "= T E S T E =")
        colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "")
        colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, ">", "")
        colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, ">", "")
        colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, ">", "")
    
    Else
    
        colMensagem.Add Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "Extrato No. " & Format(objVenda.objCupomFiscal.lNumero, "000000"))
        colMensagem.Add Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "CUPOM FISCAL ELETRÔNICO - SAT")
    
    End If
    
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", "")
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "CPF/CNPJ do Consumidor: " & sCPFCGC)
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", "")
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "#|COD|DESC|QTD|UN|VL UN R$|VL ITEM R$")
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", "")
    
    iIndice = 0
    
    For Each objItem In objVenda.objCupomFiscal.colItens
    
        If objItem.iStatus = STATUS_ATIVO Then
        
            iIndice = iIndice + 1
            
            sTexto = Format(iIndice, "000") & " " & objItem.sProduto & " " & objItem.sProdutoNomeRed & " " & IIf(right(Format(objItem.dQuantidade, "0.0###"), 2) = ",0", Format(objItem.dQuantidade, "###0"), Format(objItem.dQuantidade, "0.0###")) & " " & objItem.sUnidadeMed & " X " & Format(objItem.dPrecoUnitario, "#,###,##0.00")
            sTexto2 = Format(Arredonda_Moeda(objItem.dQuantidade * objItem.dPrecoUnitario, 4), "#,###,##0.00")
            iTam = Len(sTexto) + Len(sTexto2)
            
            If iTam < 48 Then
                colMensagem.Add sTexto & String(48 - iTam, " ") & sTexto2
            Else
                colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", sTexto)
                colMensagem.Add Formata_Campo(ALINHAMENTO_ESQUERDA, 48, " ", sTexto2)
            End If
            
            dPreco = Arredonda_Moeda(objItem.dPrecoUnitario * objItem.dQuantidade)
            
            'se existir desconto sobre o item...
            If objItem.dPercDesc > 0 Then
                sTexto = Formata_Campo(ALINHAMENTO_DIREITA, 15, " ", "DESC. ITEM:") & Formata_Campo(ALINHAMENTO_DIREITA, 10, " ", Format(objItem.dPercDesc, "standard") & "%") & " = " & Formata_Campo(ALINHAMENTO_ESQUERDA, 15, " ", Format(objItem.dPrecoUnitario * objItem.dQuantidade * (1 - (objItem.dPercDesc / 100)), "standard"))
                colMensagem.Add sTexto
                dPreco = dPreco - (objItem.dPrecoUnitario * objItem.dQuantidade * objItem.dPercDesc / 100)
            ElseIf objItem.dValorDesconto > 0 Then
                sTexto = Formata_Campo(ALINHAMENTO_DIREITA, 15, " ", "DESC. ITEM:") & Formata_Campo(ALINHAMENTO_DIREITA, 10, " ", Format(-objItem.dValorDesconto, "standard")) & " = " & Formata_Campo(ALINHAMENTO_ESQUERDA, 15, " ", Format((objItem.dPrecoUnitario - objItem.dValorDesconto) * objItem.dQuantidade, "standard"))
                colMensagem.Add sTexto
                dPreco = dPreco - objItem.dValorDesconto
            End If
            
            dTotal = dTotal + dPreco
    
        End If
    
    Next
                      
    If (objVenda.objCupomFiscal.dValorDesconto + objVenda.objCupomFiscal.dValorDesconto1) <> 0 Or objVenda.objCupomFiscal.dValorAcrescimo <> 0 Then
        colMensagem.Add Formata_Campo(ALINHAMENTO_ESQUERDA, 48, " ", "Subtotal :   " & Format(dTotal, "standard"))
    
        If (objVenda.objCupomFiscal.dValorDesconto + objVenda.objCupomFiscal.dValorDesconto1) <> 0 Then
            colMensagem.Add Formata_Campo(ALINHAMENTO_ESQUERDA, 48, " ", "Descontos  :  -" & Format(objVenda.objCupomFiscal.dValorDesconto + objVenda.objCupomFiscal.dValorDesconto1, "standard"))
            dTotal = dTotal - (objVenda.objCupomFiscal.dValorDesconto + objVenda.objCupomFiscal.dValorDesconto1)
        End If
    
        If objVenda.objCupomFiscal.dValorAcrescimo <> 0 Then
            colMensagem.Add Formata_Campo(ALINHAMENTO_ESQUERDA, 48, " ", "Acréscimos :  +" & Format(objVenda.objCupomFiscal.dValorAcrescimo, "standard"))
            dTotal = dTotal + objVenda.objCupomFiscal.dValorAcrescimo
        End If

    End If

    colMensagem.Add Formata_Campo(ALINHAMENTO_ESQUERDA, 48, " ", "TOTAL R$:   " & Format(dTotal, "standard"))
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "")
    For Each objMovCaixa In objVenda.colMovimentosCaixa
        If objMovCaixa.iTipo <> MOVIMENTOCAIXA_TROCO_DINHEIRO And objMovCaixa.iTipo <> MOVIMENTOCAIXA_TROCO_CONTRAVALE And objMovCaixa.iTipo <> MOVIMENTOCAIXA_TROCO_VALE Then
            colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", MeioPagtoCfe_Descricao(objMovCaixa.iCodigoCFe) & ":" & Format(objMovCaixa.dValor, "standard"))
        Else
            colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "Troco R$:" & Format(objMovCaixa.dValor, "standard"))
        End If
    Next

    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", "")
    colMensagem.Add Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "SAT No. " & objVenda.objCupomFiscal.sNumSerieECF)
    
    sSATtimeStamp = objVenda.objCupomFiscal.sSATtimeStamp
    colMensagem.Add Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", Mid(sSATtimeStamp, 7, 2) & "/" & Mid(sSATtimeStamp, 5, 2) & "/" & left(sSATtimeStamp, 4) & " - " & Mid(sSATtimeStamp, 9, 2) & ":" & Mid(sSATtimeStamp, 11, 2) & ":" & Mid(sSATtimeStamp, 13, 2))
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "")
    
    sSATChaveAcesso = Mid(objVenda.objCupomFiscal.sSATChaveAcesso, 4)
    
    colMensagem.Add Mid(sSATChaveAcesso, 1, 11) & " " & Mid(sSATChaveAcesso, 12, 11) & " " & Mid(sSATChaveAcesso, 23, 11) & " " & Mid(sSATChaveAcesso, 34, 11)
    
    SAT_Imprime1 = SUCESSO
    
    Exit Function
    
Erro_SAT_Imprime1:

    SAT_Imprime1 = gErr

    Select Case gErr

        Case 99803, 99894 To 99898, 105868, 133088, 204319, 204324, 204329, 210470

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149439)

    End Select

    Exit Function

End Function

Function ConfiguracaoSAT_Le(ByVal objConfiguracaoSAT As ClassConfiguracaoSAT) As Long
'carrega as configuracoes q estao armazenadas no BD

Dim alComando(1 To 1) As Long
Dim iIndice As Integer
Dim lErro As Long
Dim sConteudo As String
Dim tConfiguracaoSAT As typeConfiguracaoSAT

On Error GoTo Erro_ConfiguracaoSAT_Le

    If giLocalOperacao = LOCALOPERACAO_ECF Then

        For iIndice = LBound(alComando) To UBound(alComando)
            alComando(iIndice) = Comando_AbrirExt(glConexaoPAFECF)
            If alComando(iIndice) = 0 Then gError 207916
        Next
    
        With tConfiguracaoSAT
            .sCodigoDeAtivacao = String(STRING_SAT_CODIGO_ATIVACAO, 0)
            .sDirArqXml = String(255, 0)
            .sNomeArqDLL = String(255, 0)
            .sNomeArqEmulador = String(255, 0)
            .sNomeArqLogo = String(255, 0)
            .sPortaImpressora = String(STRING_PORTA_IMPRESSORA, 0)
            .ssignAC = String(STRING_SAT_SIGNAC, 0)
        End With
        
        With tConfiguracaoSAT
            lErro = Comando_Executar(alComando(1), "SELECT CodigoDeAtivacao, EmuladorSefaz, DirArqXml, NomeArqDLL, NomeArqEmulador,  ModeloImpressora, LayoutImpressao, EmTeste, NomeArqLogo, PortaImpressora, signAC FROM ConfiguracaoSAT ", _
                .sCodigoDeAtivacao, .iEmuladorSefaz, .sDirArqXml, .sNomeArqDLL, .sNomeArqEmulador, .iModeloImpressora, .iLayoutImpressao, .iEmTeste, .sNomeArqLogo, .sPortaImpressora, .ssignAC)
        End With
        If lErro <> AD_SQL_SUCESSO Then gError 214771
    
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 214775
    
        If lErro <> AD_SQL_SUCESSO Then gError 214776
        
        With objConfiguracaoSAT
            .sCodigoDeAtivacao = tConfiguracaoSAT.sCodigoDeAtivacao
            .iEmuladorSefaz = tConfiguracaoSAT.iEmuladorSefaz
            .sDirArqXml = tConfiguracaoSAT.sDirArqXml
            .sNomeArqDLL = tConfiguracaoSAT.sNomeArqDLL
            .sNomeArqEmulador = tConfiguracaoSAT.sNomeArqEmulador
            .iModeloImpressora = tConfiguracaoSAT.iModeloImpressora
            .iLayoutImpressao = tConfiguracaoSAT.iLayoutImpressao
            .iEmTeste = tConfiguracaoSAT.iEmTeste
            .sNomeArqLogo = tConfiguracaoSAT.sNomeArqLogo
            .sPortaImpressora = tConfiguracaoSAT.sPortaImpressora
            .ssignAC = tConfiguracaoSAT.ssignAC
        End With
        
        For iIndice = LBound(alComando) To UBound(alComando)
            lErro = Comando_Fechar(alComando(iIndice))
        Next
    
    End If
    
    ConfiguracaoSAT_Le = SUCESSO

    Exit Function

Erro_ConfiguracaoSAT_Le:

    ConfiguracaoSAT_Le = gErr

    Select Case gErr

        Case 207916
            Call Rotina_ErroECF(vbOKOnly, ERRO_ABERTURA_COMANDO, gErr)

        Case 214771, 214775
            Call Rotina_ErroECF(vbOKOnly, ERRO_LEITURA_CONFIGURACAOECF, gErr)

        Case 214776
            Call Rotina_ErroECF(vbOKOnly, ERRO_CONFIGURACAOSAT_NAO_CADASTRADO, gErr)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error, 207920)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function


Function MeioPagamento_Relatorio_Registros(ByVal dtDataDe As Date, ByVal dtDataAte As Date) As Long

Dim lErro As Long
Dim objMovimentoCaixa As New ClassMovimentoCaixa
Dim sLog As String
Dim colRegistro As New Collection
Dim sMensagem As String
Dim colOrcamento As Collection
Dim objVenda As ClassVenda
Dim iIndice As Integer
Dim colMensagem As New Collection
Dim iImprimiu As Integer
Dim dtData As Date
Dim alComandoOrc(1 To 1) As Long
Dim tMeioPagamento As typeMeioPagamento
Dim dTotalMeioPagamento As Double
Dim sMeioPagamento As String
Dim objAdmMeioPagto As ClassAdmMeioPagto
Dim sMeioPagamento1 As String
Dim sTexto As String
Dim sEAD As String

On Error GoTo Erro_MeioPagamento_Relatorio_Registros

    For iIndice = LBound(alComandoOrc) To UBound(alComandoOrc)
        alComandoOrc(iIndice) = Comando_AbrirExt(glConexaoOrcPAF)
        If alComandoOrc(iIndice) = 0 Then gError 214062
    Next

    tMeioPagamento.sMeioPagamento = String(STRING_MEIOPAGAMENTO_MEIOPAGAMENTO, 0)
    tMeioPagamento.sEAD = String(STRING_MEIOPAGAMENTO_EAD, 0)
    
    lErro = Comando_Executar(alComandoOrc(1), "SELECT MeioPagamento, CodTipoDocumento, Valor, DataMovimento, EAD FROM MeioPagamento WHERE DataMovimento >= ? AND DataMovimento <= ? ORDER BY DataMovimento, MeioPagamento, CodTipoDocumento", tMeioPagamento.sMeioPagamento, tMeioPagamento.iCodTipoDocumento, tMeioPagamento.dValor, tMeioPagamento.dtDataMovimento, tMeioPagamento.sEAD, dtDataDe, dtDataAte)
    If lErro <> AD_SQL_SUCESSO Then gError 214063
    
    lErro = Comando_BuscarPrimeiro(alComandoOrc(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 214064
    
    Do While lErro = AD_SQL_SUCESSO
    
        sTexto = CStr(tMeioPagamento.dtDataMovimento) & CStr(tMeioPagamento.iCodTipoDocumento) & tMeioPagamento.sMeioPagamento & CStr(tMeioPagamento.dValor)
        
        lErro = CF_ECF("Gerar_EAD", sTexto, sEAD)
        If lErro <> SUCESSO Then gError 214065
            
        sMensagem = "A2"
        sMensagem = sMensagem & Format(tMeioPagamento.dtDataMovimento, "yyyymmdd")
        sMensagem = sMensagem & IIf(sEAD <> tMeioPagamento.sEAD, Replace(Formata_Campo(ALINHAMENTO_DIREITA, 25, " ", tMeioPagamento.sMeioPagamento), " ", "?"), Formata_Campo(ALINHAMENTO_DIREITA, 25, " ", tMeioPagamento.sMeioPagamento))
        sMensagem = sMensagem & CStr(tMeioPagamento.iCodTipoDocumento)
        sMensagem = sMensagem & Formata_Campo(ALINHAMENTO_ESQUERDA, 12, "0", CDbl(Format(tMeioPagamento.dValor, "0.00")) * 100)
        
        Print #7, sMensagem
    
    
        lErro = Comando_BuscarProximo(alComandoOrc(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 214066

    Loop
        
    For iIndice = LBound(alComandoOrc) To UBound(alComandoOrc)
        Call Comando_Fechar(alComandoOrc(iIndice))
    Next


    MeioPagamento_Relatorio_Registros = SUCESSO
    
    Exit Function

Erro_MeioPagamento_Relatorio_Registros:

    MeioPagamento_Relatorio_Registros = gErr

        Select Case gErr
        
            Case 214062
                Call Rotina_ErroECF(vbOKOnly, ERRO_ABERTURA_COMANDO, gErr)
        
            Case 214065
        
            Case 214063, 214064, 214066
                Call Rotina_ErroECF(vbOKOnly, ERRO_LEITURA_MEIOPAGAMENTO, gErr)
                
            Case Else
                Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 214067)

    End Select

    For iIndice = LBound(alComandoOrc) To UBound(alComandoOrc)
        Call Comando_Fechar(alComandoOrc(iIndice))
    Next

    Exit Function

End Function

Function Relatorio_Registros_PAFECF(ByVal dtDataDe As Date, ByVal dtDataAte As Date, ByVal colProdutos As Collection) As Long

Dim lErro As Long
Dim objMovimentoCaixa As New ClassMovimentoCaixa
Dim sLog As String
Dim colRegistro As New Collection
Dim sMensagem As String
Dim colOrcamento As Collection
Dim objVenda As ClassVenda
Dim iIndice As Integer
Dim colMensagem As New Collection
Dim iImprimiu As Integer
Dim dtData As Date
Dim alComando(1 To 15) As Long
Dim alComandoOrc(1 To 12) As Long
Dim tMeioPagamento As typeMeioPagamento
Dim dTotalMeioPagamento As Double
Dim sMeioPagamento As String
Dim objAdmMeioPagto As ClassAdmMeioPagto
Dim sMeioPagamento1 As String
Dim sTexto As String
Dim sEAD As String
Dim sNomeArq As String
Dim sRet As String
Dim sRegistroEAD As String
Dim lSequencial As Long
Dim lTotRegMeioPagamento As Long
Dim lTotBDRegMeioPagamento As Long
Dim iErro As Integer
Dim lTotRegP2 As Long
Dim lTotBDRegP2 As Long
Dim lTotRegE2 As Long
Dim lTotBDRegE2 As Long
Dim lTotRegE3 As Long
Dim lTotBDRegE3 As Long
Dim lTotRegD2 As Long
Dim lTotBDRegD2 As Long
Dim lTotRegD3 As Long
Dim lTotBDRegD3 As Long
Dim lTotRegD4 As Long
Dim lTotBDRegD4 As Long
Dim lTotRegR02 As Long
Dim lTotBDRegR02 As Long
Dim lTotRegR03 As Long
Dim lTotBDRegR03 As Long
Dim lTotRegR04 As Long
Dim lTotBDRegR04 As Long
Dim lTotRegR05 As Long
Dim lTotBDRegR05 As Long
Dim lTotRegR06 As Long
Dim lTotBDRegR06 As Long
Dim lTotRegR07 As Long
Dim lTotBDRegR07 As Long
Dim sConteudo As String

On Error GoTo Erro_Relatorio_Registros_PAFECF

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_AbrirExt(glConexaoPAFECF)
        If alComando(iIndice) = 0 Then gError 214081
    Next

    For iIndice = LBound(alComandoOrc) To UBound(alComandoOrc)
        alComandoOrc(iIndice) = Comando_AbrirExt(glConexaoOrcPAF)
        If alComandoOrc(iIndice) = 0 Then gError 214082
    Next

    sNomeArq = CurDir() & IIf(right(CurDir, 1) <> "\", "\", "") & "REGISTROS_PAFECF.TXT"

    sRet = Dir(sNomeArq)

    If Len(sRet) <> 0 Then Kill sNomeArq


    sConteudo = String(STRING_BUFFER_MAX_TEXTO, 0)

    'Pesquisa o número automático que está no BD para giFilialEmpresa
    lErro = Comando_ExecutarPos(alComandoOrc(11), "SELECT Conteudo FROM OrcamentoConfig WHERE Codigo = 'TOT_REG_MEIOPAGAMENTO' AND FilialEmpresa = 0", 0, sConteudo)
    If lErro <> AD_SQL_SUCESSO Then gError 214148

    'Le o registro com o número automático
    lErro = Comando_BuscarPrimeiro(alComandoOrc(11))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 214149
    If lErro = AD_SQL_SEM_DADOS And giFilialEmpresa = EMPRESA_TODA Then gError 214150

    lTotRegMeioPagamento = StrParaLong(sConteudo)


    lErro = Comando_Executar(alComando(1), "SELECT TotRegE3 FROM ConfiguracaoECF ", lTotRegE3)
    If lErro <> AD_SQL_SUCESSO Then gError 214083

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 214084
        
        
    lErro = Comando_Executar(alComandoOrc(12), "SELECT Count(*) FROM MeioPagamento ", lTotBDRegMeioPagamento)
    If lErro <> AD_SQL_SUCESSO Then gError 214085

    lErro = Comando_BuscarPrimeiro(alComandoOrc(12))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 214086
    
    
    lErro = Comando_Executar(alComando(3), "SELECT TotRegP2 FROM P1 ", lTotRegP2)
    If lErro <> AD_SQL_SUCESSO Then gError 214087

    lErro = Comando_BuscarPrimeiro(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 214088
        
        
    lErro = Comando_Executar(alComando(4), "SELECT Count(*) FROM P2 ", lTotBDRegP2)
    If lErro <> AD_SQL_SUCESSO Then gError 214089

    lErro = Comando_BuscarPrimeiro(alComando(4))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 214090
    
    lErro = Comando_Executar(alComando(5), "SELECT SUM(TotRegE2) FROM E3 WHERE DataEstoque >= ? AND DataEstoque <= ? ", lTotRegE2, dtDataDe, dtDataAte)
    If lErro <> AD_SQL_SUCESSO Then gError 214091

    lErro = Comando_BuscarPrimeiro(alComando(5))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 214092
        
        
    lErro = Comando_Executar(alComando(6), "SELECT Count(*) FROM E2 WHERE DataEstoque >= ? AND DataEstoque <= ? ", lTotBDRegE2, dtDataDe, dtDataAte)
    If lErro <> AD_SQL_SUCESSO Then gError 214093

    lErro = Comando_BuscarPrimeiro(alComando(6))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 214094
    
    lErro = Comando_Executar(alComando(8), "SELECT Count(*) FROM E3 ", lTotBDRegE3)
    If lErro <> AD_SQL_SUCESSO Then gError 214095

    lErro = Comando_BuscarPrimeiro(alComando(8))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 214096
    
    lErro = Comando_Executar(alComandoOrc(1), "SELECT SUM(TotRegD2), SUM(TotRegD3), SUM(TotRegD4) FROM D1 WHERE DataDAV >= ? AND DataDAV <= ? ", lTotRegD2, lTotRegD3, lTotRegD4, dtDataDe, dtDataAte)
    If lErro <> AD_SQL_SUCESSO Then gError 214097

    lErro = Comando_BuscarPrimeiro(alComandoOrc(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 214098
    
    lErro = Comando_Executar(alComandoOrc(2), "SELECT Count(*) FROM D2 WHERE DataDAV >= ? AND DataDAV <= ? ", lTotBDRegD2, dtDataDe, dtDataAte)
    If lErro <> AD_SQL_SUCESSO Then gError 214099

    lErro = Comando_BuscarPrimeiro(alComandoOrc(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 214100
    
    lErro = Comando_Executar(alComandoOrc(3), "SELECT Count(*) FROM D3 WHERE DataDAV >= ? AND DataDAV <= ? ", lTotBDRegD3, dtDataDe, dtDataAte)
    If lErro <> AD_SQL_SUCESSO Then gError 214101

    lErro = Comando_BuscarPrimeiro(alComandoOrc(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 214102
    
    lErro = Comando_Executar(alComandoOrc(4), "SELECT Count(*) FROM D4 WHERE DataAlteracao >= ? AND DataAlteracao <= ? ", lTotBDRegD4, dtDataDe, dtDataAte)
    If lErro <> AD_SQL_SUCESSO Then gError 214103

    lErro = Comando_BuscarPrimeiro(alComandoOrc(4))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 214104
    
    lErro = Comando_Executar(alComando(9), "SELECT SUM(TotRegR02), SUM(TotRegR03), SUM(TotRegR04), SUM(TotRegR05), SUM(TotRegR06), SUM(TotRegR07) FROM R01 WHERE Data >= ? AND Data <= ? ", _
    lTotRegR02, lTotRegR03, lTotRegR04, lTotRegR05, lTotRegR06, lTotRegR07, dtDataDe, dtDataAte)
    If lErro <> AD_SQL_SUCESSO Then gError 214105

    lErro = Comando_BuscarPrimeiro(alComando(9))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 214106
    
    
    lErro = Comando_Executar(alComando(10), "SELECT Count(*) FROM R02 WHERE DataMovimento >= ? AND DataMovimento <= ? ", lTotBDRegR02, dtDataDe, dtDataAte)
    If lErro <> AD_SQL_SUCESSO Then gError 214107

    lErro = Comando_BuscarPrimeiro(alComando(10))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 214108
    
    lErro = Comando_Executar(alComando(11), "SELECT Count(*) FROM R03 WHERE DataMovimento >= ? AND DataMovimento <= ? ", lTotBDRegR03, dtDataDe, dtDataAte)
    If lErro <> AD_SQL_SUCESSO Then gError 214109

    lErro = Comando_BuscarPrimeiro(alComando(11))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 214110
    
    lErro = Comando_Executar(alComando(12), "SELECT Count(*) FROM R04 WHERE DataMovimento >= ? AND DataMovimento <= ? ", lTotBDRegR04, dtDataDe, dtDataAte)
    If lErro <> AD_SQL_SUCESSO Then gError 214111

    lErro = Comando_BuscarPrimeiro(alComando(12))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 214112
    
    lErro = Comando_Executar(alComando(13), "SELECT Count(*) FROM R05 WHERE DataMovimento >= ? AND DataMovimento <= ? ", lTotBDRegR05, dtDataDe, dtDataAte)
    If lErro <> AD_SQL_SUCESSO Then gError 214113

    lErro = Comando_BuscarPrimeiro(alComando(13))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 214114
    
    lErro = Comando_Executar(alComando(14), "SELECT Count(*) FROM R06 WHERE DataMovimento >= ? AND DataMovimento <= ? ", lTotBDRegR06, dtDataDe, dtDataAte)
    If lErro <> AD_SQL_SUCESSO Then gError 214115

    lErro = Comando_BuscarPrimeiro(alComando(14))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 214116
    
    lErro = Comando_Executar(alComando(15), "SELECT Count(*) FROM R07 WHERE DataMovimento >= ? AND DataMovimento <= ? ", lTotBDRegR07, dtDataDe, dtDataAte)
    If lErro <> AD_SQL_SUCESSO Then gError 214117

    lErro = Comando_BuscarPrimeiro(alComando(15))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 214118
    
'    If lTotRegMeioPagamento <> lTotBDRegMeioPagamento Or lTotRegP2 <> lTotBDRegP2 _
'    Or lTotRegE2 <> lTotBDRegE2 Or lTotRegE3 <> lTotBDRegE3 Or lTotRegD2 <> lTotBDRegD2 _
'    Or lTotRegD3 <> lTotBDRegD3 Or lTotRegD4 <> lTotBDRegD4 Or lTotRegR02 <> lTotBDRegR02 _
'    Or lTotRegR03 <> lTotBDRegR03 Or lTotRegR04 <> lTotBDRegR04 Or lTotRegR05 <> lTotBDRegR05 _
'    Or lTotRegR06 <> lTotBDRegR06 Or lTotRegR07 <> lTotBDRegR07 Then
'        iErro = 1
'    End If
    


    Open sNomeArq For Append Lock Read Write As #7

    sMensagem = "U1"
    sMensagem = sMensagem & Formata_Campo(ALINHAMENTO_DIREITA, 14, " ", gsCNPJ)
    sMensagem = sMensagem & Formata_Campo(ALINHAMENTO_DIREITA, 14, " ", gsInscricaoEstadual)
    sMensagem = sMensagem & Formata_Campo(ALINHAMENTO_DIREITA, 14, " ", gsInscricaoMunicipal)
    sMensagem = sMensagem & IIf(iErro = 1, Replace(Formata_Campo(ALINHAMENTO_DIREITA, 50, " ", gsNomeEmpresa), " ", "?"), Formata_Campo(ALINHAMENTO_DIREITA, 50, " ", gsNomeEmpresa))
    
    Print #7, sMensagem
    
    lErro = CF_ECF("MeioPagamento_Relatorio_Registros", dtDataDe, dtDataAte)
    If lErro <> SUCESSO Then gError 214119
        
    lErro = CF_ECF("TabProd_GravaEmTrans", lSequencial, sNomeArq)
    If lErro <> SUCESSO Then gError 214120
        
    'Função que Referenciar a Função da Afrac que Vai Executar a Leitura
    If colProdutos Is Nothing Then
    
        lErro = CF_ECF("EstProd_GravaEmTrans", lSequencial, sNomeArq)
        If lErro <> SUCESSO Then gError 214121
    
    Else
    
        lErro = CF_ECF("EstProdParcial_GravaEmTrans", lSequencial, sNomeArq, colProdutos)
        If lErro <> SUCESSO Then gError 214122
        
    End If
    
    lErro = CF_ECF("DAVEmitidos_GravaEmTrans", lSequencial, dtDataDe, dtDataAte, sNomeArq)
    If lErro <> SUCESSO Then gError 214123
    
    
    lErro = CF_ECF("MovimentoPorECF_Grava_Data", dtDataDe, dtDataAte)
    If lErro <> SUCESSO Then gError 214124
        
    Close #7

    lErro = AFRAC_CriaAssinaturaRSA(sNomeArq, sRegistroEAD)
    If lErro <> SUCESSO Then gError 214125
        
    Call Rotina_AvisoECF(vbOKOnly, "Arquivo de Registros PAF-ECF " & sNomeArq & " foi gerado com sucesso.")
        
        
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    For iIndice = LBound(alComandoOrc) To UBound(alComandoOrc)
        Call Comando_Fechar(alComandoOrc(iIndice))
    Next

    Relatorio_Registros_PAFECF = SUCESSO
    
    Exit Function

Erro_Relatorio_Registros_PAFECF:

    Relatorio_Registros_PAFECF = gErr

        Select Case gErr
        
            Case 214081, 214082
                Call Rotina_ErroECF(vbOKOnly, ERRO_ABERTURA_COMANDO, gErr)
        
            Case 214083, 214084
                Call Rotina_ErroECF(vbOKOnly, ERRO_LEITURA_CONFIGURACAOECF, gErr)
                
            Case 214085, 214086
                Call Rotina_ErroECF(vbOKOnly, ERRO_LEITURA_MEIOPAGAMENTO, gErr)
                
            Case 214087, 214088
                Call Rotina_ErroECF(vbOKOnly, ERRO_LEITURA_P1, gErr)
                
            Case 214089, 214090
                Call Rotina_ErroECF(vbOKOnly, ERRO_LEITURA_P2, gErr)
                
            Case 214091, 214092, 214095, 214096
                Call Rotina_ErroECF(vbOKOnly, ERRO_LEITURA_E3, gErr)
                
            Case 214093, 214094
                Call Rotina_ErroECF(vbOKOnly, ERRO_LEITURA_E2, gErr)
                
            Case 214097, 214098
                Call Rotina_ErroECF(vbOKOnly, ERRO_LEITURA_D1, gErr)
                
            Case 214099, 214100
                Call Rotina_ErroECF(vbOKOnly, ERRO_LEITURA_D2, gErr)
                
            Case 214101, 214102
                Call Rotina_ErroECF(vbOKOnly, ERRO_LEITURA_D3, gErr)
                
            Case 214103, 214104
                Call Rotina_ErroECF(vbOKOnly, ERRO_LEITURA_D4, gErr)
                
            Case 214105, 214106
                Call Rotina_ErroECF(vbOKOnly, ERRO_LEITURA_R01, gErr)
                
            Case 214107, 214108
                Call Rotina_ErroECF(vbOKOnly, ERRO_LEITURA_R02, gErr)
                
            Case 214109, 214110
                Call Rotina_ErroECF(vbOKOnly, ERRO_LEITURA_R03, gErr)
                
            Case 214111, 214112
                Call Rotina_ErroECF(vbOKOnly, ERRO_LEITURA_R04, gErr)
                
            Case 214113, 214114
                Call Rotina_ErroECF(vbOKOnly, ERRO_LEITURA_R05, gErr)
                
            Case 214115, 214116
                Call Rotina_ErroECF(vbOKOnly, ERRO_LEITURA_R06, gErr)
                
            Case 214117, 214118
                Call Rotina_ErroECF(vbOKOnly, ERRO_LEITURA_R07, gErr)
                
            Case 214148 To 214150
                Call Rotina_ErroECF(vbOKOnly, ERRO_ECF_LEITURA_TABELA_CONFIG, gErr, "OrcamentoConfig", "TOT_REG_MEIOPAGAMENTO")

            Case Else
                Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 214126)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    For iIndice = LBound(alComandoOrc) To UBound(alComandoOrc)
        Call Comando_Fechar(alComandoOrc(iIndice))
    Next

    Close #7

    Exit Function

End Function

Function ArqMF_Executa() As Long
'Função que Vai Chamar a Função da Afrac que vai Executar a Leitura de Memória Fiscal

Dim lErro As Long
Dim lSequencial As Long
Dim iCodGerente As Integer
Dim objOperador As New ClassOperador

On Error GoTo Erro_ArqMF_Executa

    'Se for Necessário a autorização do Gerente para abertura do Caixa
    If gobjLojaECF.iGerenteAutoriza = AUTORIZACAO_GERENTE Then

        'Chama a Tela de Senha
        Call Chama_TelaECF_Modal("OperadorLogin", objOperador, LOGIN_APENAS_GERENTE)

        'Sai de Função se a Tela de Login não Retornar ok
        If giRetornoTela <> vbOK Then gError 214070

        'Se o Operador For Gerente
        iCodGerente = objOperador.iCodigo

    End If

    'Função que Abre a Transação
    lErro = CF_ECF("Caixa_Transacao_Abrir", lSequencial)
    If lErro <> SUCESSO Then gError 214071

    'Função que Referenciar a Função da Afrac que Vai Executar a Leitura
    lErro = CF_ECF("ArqMF_ExecutaEmTrans", lSequencial, iCodGerente)
    If lErro <> SUCESSO Then gError 214072

    'Função que Encerra a Transação com o Caixa
    lErro = CF_ECF("Caixa_Transacao_Fechar", lSequencial)
    If lErro <> SUCESSO Then gError 214073

    ArqMF_Executa = SUCESSO

    Exit Function

Erro_ArqMF_Executa:

    ArqMF_Executa = gErr

     Select Case gErr

        Case 214070 To 214073
            'Erro Tratado Dentro da Função que Foi Chamada

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 214074)

    End Select
    
'    lSequencial = glSeqTransacaoAberta
            
    'Rollback na transação
    Call CF_ECF("Caixa_Transacao_Rollback", glTransacaoPAFECF)

    Exit Function

End Function

Function ArqMF_ExecutaEmTrans(lSequencial As Long, iCodGerente As Integer) As Long

Dim lErro As Long
Dim objMovimentoCaixa As New ClassMovimentoCaixa
Dim sCampoReducao1 As String
Dim sCampoReducao2 As String
Dim sLog As String
Dim colRegistro As New Collection

On Error GoTo Erro_ArqMF_ExecutaEmTrans


    If giOrcamentoECF <> CAIXA_SO_ORCAMENTO Then

        'Função que Move os Dados para a Memória
        lErro = CF_ECF("MovimentoCaixa_Move_Dados_Memoria", MOVIMENTOCAIXA_ARQUIVO_MF, objMovimentoCaixa)
        If lErro <> SUCESSO Then gError 214075
        
        'Guarda o Sequencial no objMovimentos Caixa
        objMovimentoCaixa.lSequencial = lSequencial
        
        ' Codigo Gerente
        objMovimentoCaixa.iGerente = iCodGerente
        
        Call CF_ECF("MovimentoCaixa_Gera_Log", colRegistro, objMovimentoCaixa)
        
        'Guarda as Informações no Arquivo
        lErro = CF_ECF("MovimentoCaixaECF_Grava", colRegistro)
        If lErro <> SUCESSO Then gError 214076
        
        Set colRegistro = New Collection
        
        lErro = AFRAC_ArqMF()
        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "ArqMF")
        If lErro <> SUCESSO Then gError 214077
        
    Else
        
        Call CF_ECF("Trata_Caixa_So_Orcamento")

    End If

    ArqMF_ExecutaEmTrans = SUCESSO
    
    Exit Function

Erro_ArqMF_ExecutaEmTrans:

    ArqMF_ExecutaEmTrans = gErr

        Select Case gErr
        
        Case 214075 To 214077

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 214078)

    End Select

    Exit Function

End Function

Function ConfiguracaoNFe_Le(ByVal objConfiguracaoNFe As ClassConfiguracaoNFe) As Long
'carrega as configuracoes q estao armazenadas no BD

Dim alComando(1 To 1) As Long
Dim iIndice As Integer
Dim lErro As Long
Dim sConteudo As String
Dim tConfiguracaoNFe As typeConfiguracaoNFe

On Error GoTo Erro_ConfiguracaoNFe_Le

    If giLocalOperacao = LOCALOPERACAO_ECF Then

        For iIndice = LBound(alComando) To UBound(alComando)
            alComando(iIndice) = Comando_AbrirExt(glConexaoPAFECF)
            If alComando(iIndice) = 0 Then gError 207916
        Next
    
        With tConfiguracaoNFe
            .sToken = String(STRING_NFCE_TOKEN, 0)
            .sDirArqXml = String(255, 0)
            .sNFCeSerie = String(STRING_SERIE, 0)
            .sCertificadoA1A3 = String(STRING_CERTIFICADOA1A3, 0)
            .sNFeSerie = String(STRING_SERIE, 0)
            .sDirXsd = String(255, 0)
            .sNFCECSC = String(STRING_NFCE_CSC, 0)
            .sidNFCECSC = String(STRING_NFCE_ID_CSC, 0)
            .sContigenciaxJust = String(255, 0)
            .sPortaImpressora = String(STRING_PORTA_IMPRESSORA, 0)
            .sSMTP = String(255, 0)
            .sSMTPUsu = String(255, 0)
            .sSMTPSenha = String(255, 0)
        
            lErro = Comando_Executar(alComando(1), "SELECT Token, NFCeSerie, NFCeProximoNum, DirArqXml, CertificadoA1A3, NFeAmbiente, NFeSerie, NFeProximoNum, NFeProximoLote, DirXsd, NFCeProximoLote, NFCeAmbiente, NFCECSC, idNFCECSC, EmContingencia, ContingenciaDataEntrada, ContingenciaHoraEntrada, ContigenciaxJust, ModeloImpressora, PortaImpressora, SMTP, SMTPUsu, SMTPSenha, SMTPPorta, NFCeImprimir, NFCeEnviarEmail, NFDescricaoProd, VersaoNFe, FocaTipoVenda FROM ConfiguracaoNFe ", _
                .sToken, .sNFCeSerie, .lNFCeProximoNum, .sDirArqXml, .sCertificadoA1A3, .iNFeAmbiente, .sNFeSerie, .lNFeProximoNum, .lNFeProximoLote, .sDirXsd, .lNFCeProximoLote, .iNFCeAmbiente, .sNFCECSC, .sidNFCECSC, .iEmContingencia, .dtContingenciaDataEntrada, .dContingenciaHoraEntrada, .sContigenciaxJust, .iModeloImpressora, .sPortaImpressora, .sSMTP, .sSMTPUsu, .sSMTPSenha, .lSMTPPorta, .iNFCeImprimir, .iNFCeEnviarEmail, .iNFDescricaoProd, .iVersaoNFe, .iFocaTipoVenda)
        End With
        If lErro <> AD_SQL_SUCESSO Then gError 207917
    
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 207918
    
        If lErro <> AD_SQL_SUCESSO Then gError 207919
        
        With objConfiguracaoNFe
            .sToken = tConfiguracaoNFe.sToken
            .sNFCeSerie = tConfiguracaoNFe.sNFCeSerie
            .lNFCeProximoNum = tConfiguracaoNFe.lNFCeProximoNum
            .sDirArqXml = tConfiguracaoNFe.sDirArqXml
            .sCertificadoA1A3 = tConfiguracaoNFe.sCertificadoA1A3
            .iNFeAmbiente = tConfiguracaoNFe.iNFeAmbiente
            .sNFeSerie = tConfiguracaoNFe.sNFeSerie
            .lNFeProximoNum = tConfiguracaoNFe.lNFeProximoNum
            .lNFeProximoLote = tConfiguracaoNFe.lNFeProximoLote
            .sDirXsd = tConfiguracaoNFe.sDirXsd
            .lNFCeProximoLote = tConfiguracaoNFe.lNFCeProximoLote
            .iNFCeAmbiente = tConfiguracaoNFe.iNFCeAmbiente
            .sNFCECSC = tConfiguracaoNFe.sNFCECSC
            .sidNFCECSC = tConfiguracaoNFe.sidNFCECSC
            .iEmContingencia = tConfiguracaoNFe.iEmContingencia
            .dtContingenciaDataEntrada = tConfiguracaoNFe.dtContingenciaDataEntrada
            .dContingenciaHoraEntrada = tConfiguracaoNFe.dContingenciaHoraEntrada
            .sContigenciaxJust = tConfiguracaoNFe.sContigenciaxJust
            .iModeloImpressora = tConfiguracaoNFe.iModeloImpressora
            .sPortaImpressora = tConfiguracaoNFe.sPortaImpressora
            .sSMTP = tConfiguracaoNFe.sSMTP
            .sSMTPUsu = tConfiguracaoNFe.sSMTPUsu
            .sSMTPSenha = tConfiguracaoNFe.sSMTPSenha
            .lSMTPPorta = tConfiguracaoNFe.lSMTPPorta
            .iNFCeImprimir = tConfiguracaoNFe.iNFCeImprimir
            .iNFCeEnviarEmail = tConfiguracaoNFe.iNFCeEnviarEmail
            .iNFDescricaoProd = tConfiguracaoNFe.iNFDescricaoProd
            .iVersaoNFe = tConfiguracaoNFe.iVersaoNFe
            .iFocaTipoVenda = tConfiguracaoNFe.iFocaTipoVenda
        End With
        
        For iIndice = LBound(alComando) To UBound(alComando)
            lErro = Comando_Fechar(alComando(iIndice))
        Next
    
    End If
    
    ConfiguracaoNFe_Le = SUCESSO

    Exit Function

Erro_ConfiguracaoNFe_Le:

    ConfiguracaoNFe_Le = gErr

    Select Case gErr

        Case 207916
            Call Rotina_ErroECF(vbOKOnly, ERRO_ABERTURA_COMANDO, gErr)

        Case 207917, 207918
            Call Rotina_ErroECF(vbOKOnly, ERRO_LEITURA_CONFIGURACAONFE, gErr)

        Case 207919
            Call Rotina_ErroECF(vbOKOnly, ERRO_CONFIGURACAOECF_NAO_CADASTRADO, gErr)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error, 207920)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function


Function NFD2_Le(ByVal objNF As ClassNFiscal) As Long

Dim alComando(1 To 2) As Long, lErro As Long
Dim iIndice As Integer, objItem As ClassItemNF
Dim lNumIntDoc As Long, lNumero As Long
Dim sSerie As String, dtDataEmissao As Date, sDestinatario As String
Dim dQtde As Double, sDiscriminacao As String, dPrecoUnitario As Double
Dim sProdutoNomeRed As String, sProduto As String

On Error GoTo Erro_NFD2_Le

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_AbrirExt(glConexaoOrcPAF)
        If alComando(iIndice) = 0 Then gError 213423
    Next
    
    sSerie = String(STRING_MAXIMO, 0)
    sDestinatario = String(STRING_MAXIMO, 0)

    If objNF.lNumIntDoc = 0 Then
        lErro = Comando_Executar(alComando(1), "SELECT NumIntDoc, Numero, Serie, DataEmissao, Destinatario FROM NFD2 WHERE Serie = ? AND Numero = ? AND DataEmissao BETWEEN ? AND ?", lNumIntDoc, lNumero, sSerie, dtDataEmissao, sDestinatario, objNF.sSerie, objNF.lNumNotaFiscal, objNF.dtDataEmissao - PERIODO_EMISSAO, objNF.dtDataEmissao + PERIODO_EMISSAO)
    Else
        lErro = Comando_Executar(alComando(1), "SELECT NumIntDoc, Numero, Serie, DataEmissao, Destinatario FROM NFD2 WHERE NumIntDoc = ?", lNumIntDoc, lNumero, sSerie, dtDataEmissao, sDestinatario, objNF.lNumIntDoc)
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 213424
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213425
    
    If lErro <> AD_SQL_SUCESSO Then gError ERRO_LEITURA_SEM_DADOS
    
    objNF.lNumNotaFiscal = lNumero
    objNF.sSerie = sSerie
    objNF.sDestino = sDestinatario
    objNF.lNumIntDoc = lNumIntDoc
    objNF.dtDataEmissao = dtDataEmissao
    
    sDiscriminacao = String(STRING_MAXIMO, 0)
    
    sProduto = String(STRING_MAXIMO, 0)
    sProdutoNomeRed = String(STRING_MAXIMO, 0)
    
    lErro = Comando_Executar(alComando(2), "SELECT Qtde, Produto, ProdutoNomeRed, Discriminacao, PrecoUnitario FROM NFD2Itens WHERE NumIntNF = ? ORDER BY Item", dQtde, sProduto, sProdutoNomeRed, sDiscriminacao, dPrecoUnitario, objNF.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 213426
    
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213427
    
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objItem = New ClassItemNF
    
        objItem.dQuantidade = dQtde
        objItem.sDescricaoItem = sDiscriminacao
        objItem.dPrecoUnitario = dPrecoUnitario
        objItem.sProduto = sProduto
        objItem.sProdutoXml = sProdutoNomeRed
        
        objNF.ColItensNF.Add1 objItem
    
        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213428
    
    Loop
    
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    NFD2_Le = SUCESSO

    Exit Function

Erro_NFD2_Le:

    NFD2_Le = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case ERRO_LEITURA_SEM_DADOS

        Case 213423
            Call Rotina_ErroECF(vbOKOnly, ERRO_ABERTURA_COMANDO, gErr)

        Case 213424 To 213428
            Call Rotina_ErroECF(vbOKOnly, ERRO_LEITURA_NFD2, gErr)
            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error, 213429)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function NFD2_Le_Todas(ByVal colNF As Collection) As Long

Dim alComando(1 To 2) As Long, lErro As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long, lNumero As Long
Dim sSerie As String, dtDataEmissao As Date, sDestinatario As String
Dim objNF As ClassNFiscal

On Error GoTo Erro_NFD2_Le_Todas

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_AbrirExt(glConexaoOrcPAF)
        If alComando(iIndice) = 0 Then gError 213466
    Next
    
    sSerie = String(STRING_MAXIMO, 0)
    sDestinatario = String(STRING_MAXIMO, 0)

    lErro = Comando_Executar(alComando(1), "SELECT NumIntDoc, Numero, Serie, DataEmissao, Destinatario FROM NFD2 ORDER BY DataEmissao, NumIntDoc", lNumIntDoc, lNumero, sSerie, dtDataEmissao, sDestinatario)
    If lErro <> AD_SQL_SUCESSO Then gError 213467
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213468
    
    Do While lErro = AD_SQL_SUCESSO
    
        Set objNF = New ClassNFiscal
    
        objNF.lNumNotaFiscal = lNumero
        objNF.sSerie = sSerie
        objNF.sDestino = sDestinatario
        objNF.lNumIntDoc = lNumIntDoc
        objNF.dtDataEmissao = dtDataEmissao
    
        colNF.Add objNF
    
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213469
        
    Loop
        
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    NFD2_Le_Todas = SUCESSO

    Exit Function

Erro_NFD2_Le_Todas:

    NFD2_Le_Todas = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case 213466
            Call Rotina_ErroECF(vbOKOnly, ERRO_ABERTURA_COMANDO, gErr)

        Case 213467 To 213469
            Call Rotina_ErroECF(vbOKOnly, ERRO_LEITURA_NFD2, gErr)
            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error, 213470)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function NFe_Imprime(ByVal objTela As Object, ByVal objVenda As ClassVenda) As Long
'imprime o NFe

Dim lErro As Long
Dim colMensagem As New Collection
Dim iIndice As Integer
Dim sMensagem As String

On Error GoTo Erro_NFe_Imprime
    
    lErro = NFe_Imprime1(objVenda, colMensagem)
    If lErro <> SUCESSO Then gError 133088
    
    lErro = AFRAC_AbrirRelatorioGerencial(RELGER_NFE, objTela)
    lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Abrir NFe")
    If lErro <> SUCESSO Then gError 133088

    For iIndice = 1 To colMensagem.Count
            
        sMensagem = colMensagem.Item(iIndice)
        If sMensagem = "" Then sMensagem = "    "
        lErro = AFRAC_ImprimirRelatorioGerencial(sMensagem, objTela)
        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "NFe")
        If lErro <> SUCESSO Then gError 99893
            
    Next

    lErro = AFRAC_FecharRelatorioGerencial(objTela)
    lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Fecha NFe")
    If lErro <> SUCESSO Then gError 99898
    
    NFe_Imprime = SUCESSO
    
    Exit Function
    
Erro_NFe_Imprime:

    NFe_Imprime = gErr

    Select Case gErr

        Case 99803, 99894 To 99898, 105868, 133088, 204319, 204324, 204329, 210470

        Case 99893
            Call AFRAC_FecharRelatorioGerencial(objTela)

        Case 210471
            Call Rotina_ErroECF(vbOKOnly, ERRO_ORCAMENTO_NAO_PREENCHIDO, lErro)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149439)

    End Select

    Exit Function

End Function

Private Function NFe_Imprime1(ByVal objVenda As ClassVenda, colMensagem As Collection) As Long
'imprime a NFe

Dim lErro As Long
Dim objItens As ClassItemCupomFiscal
Dim objProduto As ClassProduto
Dim iIndice As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim bCont As Boolean
Dim sMensagem As String
Dim dTotal As Double
Dim dPreco As Double
Dim bAchou As Boolean
Dim colOrcamento As New Collection
Dim iIndice1 As Integer
Dim objCliente As ClassCliente
Dim lIndice As Long
Dim sCliente As String
Dim sCPFCGC As String
Dim sVendedor As String
Dim objVendedor As ClassVendedor
Dim sProduto As String
Dim sCNPJ As String
Dim iPos As Integer
Dim lNumDAV As Long
Dim sRetorno As String
Dim objProdutoNomeRed As TextBox
Dim sTexto As String

On Error GoTo Erro_NFe_Imprime1

    Set colMensagem = New Collection
    
    Set objCliente = gobjClienteCPF.Busca(objVenda.objCupomFiscal.sCPFCGC, lIndice)
    If Not objCliente Is Nothing Then sCliente = objCliente.sNomeReduzido & " - "
    Call Formata_CNPJ_CPF(sCPFCGC, objVenda.objCupomFiscal.sCPFCGC)
    sCliente = sCliente & sCPFCGC
    
    Call Formata_CNPJ_CPF(sCNPJ, gsCNPJ)
    
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "Emitente: " & gsNomeEmpresa)
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "CNPJ Emitente: " & sCNPJ)
    colMensagem.Add TRACO_CAB
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "Destinatário: " & objVenda.objCupomFiscal.sNomeCliente)
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "CPF/CNPJ Destinatário: " & sCPFCGC)
    colMensagem.Add TRACO_CAB
    colMensagem.Add " N" & Chr(176) & " do Documento: " & Format(objVenda.objCupomFiscal.lNumero, "0000000000")
    colMensagem.Add TRACO_CAB

    For Each objItens In objVenda.objCupomFiscal.colItens
    
        Set objProduto = New ClassProduto
            
        lErro = CF_ECF("Produtos_Le", objItens.sProduto, objProduto)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 214855
        
        If lErro = SUCESSO Then
            objItens.sProdutoNomeRed = objProduto.sNomeReduzido
        Else
            'caso o produto não seja encontrado
            gError 99894
        End If
            
        'Joga no cupom o item
        colMensagem.Add Formata_Campo(ALINHAMENTO_ESQUERDA, 3, 0, objItens.iItem) & " " & Formata_Campo(ALINHAMENTO_DIREITA, 45, " ", objProduto.sCodigo & " " & objProduto.sDescricao)
        colMensagem.Add Formata_Campo(ALINHAMENTO_ESQUERDA, 10, " ", objItens.dQuantidade) & Formata_Campo(ALINHAMENTO_CENTRALIZADO, 8, " ", "x") & Formata_Campo(ALINHAMENTO_DIREITA, 10, " ", Format(objItens.dPrecoUnitario, "standard")) & "  =  " & Formata_Campo(ALINHAMENTO_ESQUERDA, 15, " ", Format(objItens.dPrecoUnitario * objItens.dQuantidade, "standard"))
        dPreco = objItens.dPrecoUnitario * objItens.dQuantidade
                    
        'se existir desconto sobre o item...
        If objItens.dPercDesc > 0 Then
            colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 15, " ", "DESC. ITEM:") & Formata_Campo(ALINHAMENTO_DIREITA, 10, " ", Format(objItens.dPercDesc, "standard") & "%") & "     =  " & Formata_Campo(ALINHAMENTO_ESQUERDA, 15, " ", Format(objItens.dPrecoUnitario * objItens.dQuantidade * (1 - (objItens.dPercDesc / 100)), "standard"))
            dPreco = dPreco - (objItens.dPrecoUnitario * objItens.dQuantidade * objItens.dPercDesc / 100)
        ElseIf objItens.dValorDesconto > 0 Then
            'colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 15, " ", "DESC. ITEM:") & Formata_Campo(ALINHAMENTO_DIREITA, 10, " ", Format(-objItens.dValorDesconto, "standard")) & "     =  " & Formata_Campo(ALINHAMENTO_ESQUERDA, 15, " ", Format((objItens.dPrecoUnitario - objItens.dValorDesconto) * objItens.dQuantidade, "standard"))
            'dPreco = dPreco - (objItens.dQuantidade * objItens.dValorDesconto)
            colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 15, " ", "DESC. ITEM:") & Formata_Campo(ALINHAMENTO_DIREITA, 10, " ", Format(-objItens.dValorDesconto, "standard")) & "     =  " & Formata_Campo(ALINHAMENTO_ESQUERDA, 15, " ", Format((objItens.dPrecoUnitario * objItens.dQuantidade - objItens.dValorDesconto), "standard"))
            dPreco = dPreco - objItens.dValorDesconto
        End If
        
        If objItens.iStatus = STATUS_ATIVO Then
            dTotal = dTotal + dPreco
        ElseIf objItens.iStatus = STATUS_CANCELADO Then
            colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "************* ITEM " & Format(objItens.iItem, "000") & " CANCELADO *************")
        End If
        
    Next
                      
    colMensagem.Add Formata_Campo(ALINHAMENTO_ESQUERDA, 48, " ", "TOTAL :   " & Format(dTotal, "standard"))
    
    If (objVenda.objCupomFiscal.dValorDesconto + objVenda.objCupomFiscal.dValorDesconto1) <> 0 Then
        colMensagem.Add Formata_Campo(ALINHAMENTO_ESQUERDA, 48, " ", "DESCONTO  :  -" & Format(objVenda.objCupomFiscal.dValorDesconto + objVenda.objCupomFiscal.dValorDesconto1, "standard"))
        dTotal = dTotal - (objVenda.objCupomFiscal.dValorDesconto + objVenda.objCupomFiscal.dValorDesconto1)
    End If
    
    If objVenda.objCupomFiscal.dValorAcrescimo <> 0 Then
        colMensagem.Add Formata_Campo(ALINHAMENTO_ESQUERDA, 48, " ", "ACRESCIMO :  +" & Format(objVenda.objCupomFiscal.dValorAcrescimo, "standard"))
        dTotal = dTotal + objVenda.objCupomFiscal.dValorAcrescimo
    End If
    
    If (objVenda.objCupomFiscal.dValorDesconto + objVenda.objCupomFiscal.dValorDesconto1) <> 0 Or objVenda.objCupomFiscal.dValorAcrescimo <> 0 Then
        colMensagem.Add Formata_Campo(ALINHAMENTO_ESQUERDA, 48, " ", "TOTAL LIQ.:   " & Format(dTotal, "standard"))
    End If
    
    NFe_Imprime1 = SUCESSO
    
    Exit Function
    
Erro_NFe_Imprime1:

    NFe_Imprime1 = gErr

    Select Case gErr

        Case 99803, 99894 To 99898, 105868, 133088, 204319, 204324, 204329, 210470

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149439)

    End Select

    Exit Function

End Function

Function CodNacId_Le(ByVal sMarca As String, ByVal sModelo As String, ByVal sVersaoSB As String, sCodNacId As String) As Long
'retorna o codigo nacional de identificacao do ecf cuja marca, modelo e versaoSB foram passados como parametros

Dim alComando(1 To 1) As Long
Dim iIndice As Integer
Dim lErro As Long
Dim objArqSeq As ClassArqSeq
Dim lSequencial As Long
Dim lUltimoNumIntDoc As Long

On Error GoTo Erro_CodNacId_Le


    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_AbrirExt(glConexaoPAFECF)
        If alComando(iIndice) = 0 Then gError 214506
    Next

    sCodNacId = String(STRING_CODNACID_IDENTIFICACAOECF, 0)

    lErro = Comando_Executar(alComando(1), "SELECT Identificacao FROM CodNacId WHERE Marca = ? AND Modelo = ? AND VersaoSB = ?", sCodNacId, sMarca, sModelo, sVersaoSB)
    If lErro <> AD_SQL_SUCESSO Then gError 214507
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 214508

    If lErro <> AD_SQL_SUCESSO Then gError 214509

    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next
    
    CodNacId_Le = SUCESSO

    Exit Function

Erro_CodNacId_Le:

    CodNacId_Le = gErr

    Select Case gErr

        Case 214506
            Call Rotina_ErroECF(vbOKOnly, ERRO_ABERTURA_COMANDO, gErr)

        Case 214507, 214508
            Call Rotina_ErroECF(vbOKOnly, ERRO_LEITURA_CODNACID, gErr)

        Case 214509
            Call Rotina_ErroECF(vbOKOnly, ERRO_CODNACID_NAO_CADASTRADO, gErr, sMarca, sModelo, sVersaoSB)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error, 214510)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Private Function Venda_ObterTotTrib(ByVal objVenda As ClassVenda) As Double

Dim lErro As Long, dTotTrib As Double
Dim objItem As ClassItemCupomFiscal

    If objVenda.objCupomFiscal.objTributacaoDoc.dTotTrib <> 0 Then
    
        dTotTrib = objVenda.objCupomFiscal.objTributacaoDoc.dTotTrib
        
    Else
    
        For Each objItem In objVenda.objCupomFiscal.colItens
        
            If objItem.iStatus <> STATUS_CANCELADO Then
            
                dTotTrib = Arredonda_Moeda(dTotTrib + objItem.objTributacaoDocItem.dTotTrib)
                
            End If
        
        Next
    
    End If
    
    Venda_ObterTotTrib = dTotTrib

End Function

Function Venda_MsgValAproxTrib(ByVal objVenda As ClassVenda, sMsg As String) As Long
'Monta mensagem de valor aproximado de tributos para sair no cupom fiscal (lei 12741)

Dim lErro As Long, dTotTrib As Double, sMsgAux As String
Dim objItem As ClassItemCupomFiscal, iTotTribTipo As Integer

On Error GoTo Erro_Venda_MsgValAproxTrib
    
    iTotTribTipo = -1
    
    sMsg = "Val.Aprox.Tributos R$"
    
    For Each objItem In objVenda.objCupomFiscal.colItens
    
        If objItem.iStatus <> STATUS_CANCELADO Then
        
            If iTotTribTipo = -1 Then
                iTotTribTipo = objItem.objTributacaoDocItem.iTotTribTipo
            Else
                If iTotTribTipo <> objItem.objTributacaoDocItem.iTotTribTipo Then iTotTribTipo = -2
            End If
            dTotTrib = Arredonda_Moeda(dTotTrib + objItem.objTributacaoDocItem.dTotTrib)
            
        End If
    
    Next
    
    If objVenda.objCupomFiscal.objTributacaoDoc.dTotTrib <> 0 Then dTotTrib = objVenda.objCupomFiscal.objTributacaoDoc.dTotTrib
    
    sMsg = sMsg & Format(dTotTrib, "Fixed") & "(" & Format(IIf(dTotTrib = 0 Or objVenda.objCupomFiscal.dValorTotal = 0, 0, dTotTrib / objVenda.objCupomFiscal.dValorTotal), "Percent") & ")"
    
    Select Case iTotTribTipo
    
        Case 1
           sMsgAux = " Fonte:IBPT"
    
        Case -2 'mais de um tipo de calculo
           sMsgAux = ""
        
        Case Else
           sMsgAux = ""
        
    End Select
    
    sMsg = sMsg & sMsgAux
    
    Venda_MsgValAproxTrib = SUCESSO
    
    Exit Function
    
Erro_Venda_MsgValAproxTrib:

    Venda_MsgValAproxTrib = gErr

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149439)

    End Select

    Exit Function

End Function

Private Function MeioPagtoCfe_Descricao(ByVal iCodigoCFe As Integer) As String
Dim sTexto As String

    Select Case iCodigoCFe
    
        Case 1
            sTexto = "DINHEIRO"
        Case 2
            sTexto = "Cheque"
        Case 3
            sTexto = "Cartão de Crédito"
        Case 4
            sTexto = "Cartão de Débito"
        Case 5
            sTexto = "Crédito Loja"
        Case 10
            sTexto = "Vale Alimentação"
        Case 11
            sTexto = "Vale Refeição"
        Case 12
            sTexto = "Vale Presente"
        Case 13
            sTexto = "Vale Combustível"
        Case Else
            sTexto = "Outros"
            
    End Select

    MeioPagtoCfe_Descricao = sTexto

End Function

Function SAT_Imprime_Canc(ByVal objTela As Object, ByVal objVenda As ClassVenda) As Long
'imprime o SAT

Dim lErro As Long
Dim colMensagem As New Collection
Dim iIndice As Integer
Dim sMensagem As String

On Error GoTo Erro_SAT_Imprime_Canc
    
    lErro = AFRAC_AbrirRelatorioGerencial(RELGER_SAT_CANC, objTela)
    lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Abrir SAT Canc")
    If lErro <> SUCESSO Then gError 133088

    Select Case gobjSATInfo.iModeloImpressora
        
        Case IMPRESSORA_DARUMA_DR700
        
            Call iImprimirCFe_SAT_Daruma(objVenda.objCupomFiscal.sSATCancArqXml, "2")

            'lErro = SAT_Imprime_Canc_Daruma(objVenda)
            'If lErro <> SUCESSO Then gError 133088
    
        Case Else
        
            lErro = SAT_Imprime_Canc1(objVenda, colMensagem)
            If lErro <> SUCESSO Then gError 133088
    
            For iIndice = 1 To colMensagem.Count
                    
                sMensagem = colMensagem.Item(iIndice)
                If sMensagem = "" Then sMensagem = "    "
                lErro = AFRAC_ImprimirRelatorioGerencial(sMensagem, objTela)
                lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "SAT Canc")
                If lErro <> SUCESSO Then gError 99893
                    
            Next

    End Select

    lErro = AFRAC_FecharRelatorioGerencial(objTela)
    lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Fecha SAT Canc")
    If lErro <> SUCESSO Then gError 99898
    
    SAT_Imprime_Canc = SUCESSO
    
    Exit Function
    
Erro_SAT_Imprime_Canc:

    SAT_Imprime_Canc = gErr

    Select Case gErr

        Case 99803, 99894 To 99898, 105868, 133088, 204319, 204324, 204329, 210470

        Case 99893
            Call AFRAC_FecharRelatorioGerencial(objTela)

        Case 210471
            Call Rotina_ErroECF(vbOKOnly, ERRO_ORCAMENTO_NAO_PREENCHIDO, lErro)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149439)

    End Select

    Exit Function

End Function

Private Function SAT_Imprime_Canc1(ByVal objVenda As ClassVenda, colMensagem As Collection) As Long
'preenche colMensagem com as linhas a serem impressas no cancelamento do SAT

Dim lErro As Long
Dim objItem As ClassItemCupomFiscal
Dim objProduto As ClassProduto
Dim iIndice As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim bCont As Boolean
Dim sMensagem As String
Dim dTotal As Double
Dim dPreco As Double
Dim bAchou As Boolean
Dim colOrcamento As New Collection
Dim iIndice1 As Integer
Dim objCliente As ClassCliente
Dim lIndice As Long
Dim sCliente As String
Dim sCPFCGC As String
Dim sVendedor As String
Dim objVendedor As ClassVendedor
Dim sProduto As String
Dim sCNPJ As String
Dim iPos As Integer
Dim lNumDAV As Long, objMovCaixa As ClassMovimentoCaixa
Dim sRetorno As String, sSATChaveAcesso As String
Dim objProdutoNomeRed As TextBox, sSATtimeStamp As String
Dim sTexto As String, sTexto2 As String, iTam As Integer

On Error GoTo Erro_SAT_Imprime_Canc1

    Set colMensagem = New Collection
    
    Set objCliente = gobjClienteCPF.Busca(objVenda.objCupomFiscal.sCPFCGC, lIndice)
    If Not objCliente Is Nothing Then sCliente = objCliente.sNomeReduzido & " - "
    Call Formata_CNPJ_CPF(sCPFCGC, objVenda.objCupomFiscal.sCPFCGC)
    sCliente = sCliente & sCPFCGC
    
    Call Formata_CNPJ_CPF(sCNPJ, gsCNPJ)
    
    colMensagem.Add Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", gsNomeReduzido)
    colMensagem.Add Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", gsNomeEmpresa)
    colMensagem.Add Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", gsEndereco)
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "")
    colMensagem.Add Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "CNPJ:" & sCNPJ)
    colMensagem.Add Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "IE:" & gsInscricaoEstadual)
    colMensagem.Add Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "IM:" & gsInscricaoMunicipal)
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", "")
    
    If objVenda.objCupomFiscal.iSATtpAmb <> 1 Then
    
        colMensagem.Add Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "Extrato No. 000000")
        colMensagem.Add Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "CUPOM FISCAL ELETRÔNICO - SAT")
        colMensagem.Add Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "CANCELAMENTO")
        colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "")
        colMensagem.Add Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "= T E S T E =")
        colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "")
        colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, ">", "")
        colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, ">", "")
        colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, ">", "")
    
    Else
    
        colMensagem.Add Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "Extrato No. " & Format(objVenda.objCupomFiscal.lNumero, "000000"))
        colMensagem.Add Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "CUPOM FISCAL ELETRÔNICO - SAT")
        colMensagem.Add Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "CANCELAMENTO")
    
    End If
    
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", "")
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "CPF/CNPJ do Consumidor: " & sCPFCGC)
    
    colMensagem.Add Formata_Campo(ALINHAMENTO_ESQUERDA, 48, " ", "TOTAL R$:   " & Format(objVenda.objCupomFiscal.dValorTotal, "standard"))

    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", "")
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "DADOS DO CUPOM FISCAL ELETRÔNICO CANCELADO")
    colMensagem.Add Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "SAT No. " & objVenda.objCupomFiscal.sNumSerieECF)
    
    sSATtimeStamp = objVenda.objCupomFiscal.sSATtimeStamp
    colMensagem.Add Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", Mid(sSATtimeStamp, 7, 2) & "/" & Mid(sSATtimeStamp, 5, 2) & "/" & left(sSATtimeStamp, 4) & " - " & Mid(sSATtimeStamp, 9, 2) & ":" & Mid(sSATtimeStamp, 11, 2) & ":" & Mid(sSATtimeStamp, 13, 2))
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "")
    
    sSATChaveAcesso = Mid(objVenda.objCupomFiscal.sSATChaveAcesso, 4)
    
    colMensagem.Add Mid(sSATChaveAcesso, 1, 11) & " " & Mid(sSATChaveAcesso, 12, 11) & " " & Mid(sSATChaveAcesso, 23, 11) & " " & Mid(sSATChaveAcesso, 34, 11)
    
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "")
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "")
    
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", "")
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "DADOS DO CUPOM FISCAL ELETRÔNICO DE CANCELAMENTO")
    colMensagem.Add Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "SAT No. " & objVenda.objCupomFiscal.sNumSerieECF)
    
    sSATtimeStamp = objVenda.objCupomFiscal.sSATCanctimeStamp
    colMensagem.Add Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", Mid(sSATtimeStamp, 7, 2) & "/" & Mid(sSATtimeStamp, 5, 2) & "/" & left(sSATtimeStamp, 4) & " - " & Mid(sSATtimeStamp, 9, 2) & ":" & Mid(sSATtimeStamp, 11, 2) & ":" & Mid(sSATtimeStamp, 13, 2))
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "")
    
    sSATChaveAcesso = Mid(objVenda.objCupomFiscal.sSATCancChaveAcesso, 4)
    
    colMensagem.Add Mid(sSATChaveAcesso, 1, 11) & " " & Mid(sSATChaveAcesso, 12, 11) & " " & Mid(sSATChaveAcesso, 23, 11) & " " & Mid(sSATChaveAcesso, 34, 11)
    
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "")
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "")
    
    SAT_Imprime_Canc1 = SUCESSO
    
    Exit Function
    
Erro_SAT_Imprime_Canc1:

    SAT_Imprime_Canc1 = gErr

    Select Case gErr

        Case 99803, 99894 To 99898, 105868, 133088, 204319, 204324, 204329, 210470

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149439)

    End Select

    Exit Function

End Function

Private Function FormataTextoEspacosNoMeio(ByVal sTexto1 As String, ByVal sTexto2 As String, ByVal iTam As Integer) As String
'completa com espaços entre sTexto1 e sTexto2 para chegar ao iTam
    
    Dim iAux As Integer, sTexto As String
    
    iAux = Len(sTexto1 & " " & sTexto2)
    If iAux < iTam Then
        sTexto = sTexto1 & " " & String(iTam - iAux, " ") & sTexto2
    Else
        sTexto = left(sTexto1 & " " & sTexto2, iTam)
    End If
    
    FormataTextoEspacosNoMeio = sTexto

End Function

Private Function BematechNaoFiscal_ImprimirNFCE(ByVal objVenda As ClassVenda) As Long

Dim lErro As Long
Dim objItens As ClassItemCupomFiscal
Dim objProduto As ClassProduto
Dim iIndice As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim bCont As Boolean
Dim sMensagem As String
Dim dTotal As Double
Dim dPreco As Double
Dim bAchou As Boolean
Dim colOrcamento As New Collection
Dim iIndice1 As Integer
Dim objCliente As ClassCliente
Dim lIndice As Long
Dim sCliente As String
Dim sCPFCGC As String
Dim sVendedor As String
Dim objVendedor As ClassVendedor
Dim sProduto As String
Dim sCNPJ As String
Dim iPos As Integer
Dim lNumDAV As Long
Dim sRetorno As String
Dim objProdutoNomeRed As TextBox, objMovCaixa As ClassMovimentoCaixa
Dim sTexto As String, sChaveAcesso As String, sTexto1 As String, sTexto2 As String, iAux As Integer, iItem As Integer
Dim objBemaNaoFiscal As New ClassBematechNaoFiscalImp

On Error GoTo Erro_BematechNaoFiscal_ImprimirNFCE

    Call objBemaNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", gsNomeEmpresa))
    Call Formata_CNPJ_CPF(sCNPJ, gsCNPJ)
    Call objBemaNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "CNPJ:" & sCNPJ & " IE:" & gsInscricaoEstadual & " IM:" & gsInscricaoMunicipal))
    Call objBemaNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", gsEndereco))
    Call objBemaNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", ""))
    Call objBemaNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "DANFE NFC-e - Documento Auxiliar"))
    Call objBemaNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "da Nota Fiscal de Consumidor Eletronica"))
    Call objBemaNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "Não permite aproveitamento de crédito de ICMS"))
    Call objBemaNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", ""))
    Call objBemaNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "Codigo Descricao         Qtde Un  VlrUnit VlrTot"))
    Call objBemaNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", ""))

    iItem = 0
    
    For Each objItens In objVenda.objCupomFiscal.colItens
    
        If objItens.iStatus <> STATUS_CANCELADO Then
    
            iItem = iItem + 1
    
            Set objProduto = New ClassProduto
            
            lErro = CF_ECF("Produtos_Le", objItens.sProduto, objProduto)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 214855
            
            If lErro = SUCESSO Then
                objItens.sProdutoNomeRed = objProduto.sNomeReduzido
            Else
                'caso o produto não seja encontrado
                gError 99894
            End If
                
            sTexto1 = objProduto.sCodigo & " " & objProduto.sDescricao
            sTexto1 = left(sTexto1, 48)
                
            sTexto2 = Format(objItens.dQuantidade, "0.0###")
            If right(sTexto2, 2) = ",0" Then sTexto2 = left(sTexto2, Len(sTexto2) - 2)
            sTexto2 = sTexto2 & " " & objItens.sUnidadeMed
            sTexto2 = sTexto2 & " X " & Format(objItens.dPrecoUnitario, "standard")
            sTexto2 = sTexto2 & " " & Format(objItens.dPrecoUnitario * objItens.dQuantidade, "standard")
                
            iAux = Len(sTexto1 & " " & sTexto2)
            If Len(sTexto1 & " " & sTexto2) <= 48 Then
                Call objBemaNaoFiscal.ImprimeNormal(sTexto1 & String(48 - iAux, " ") & " " & sTexto2)
            Else
                If Len(sTexto1 & " " & sTexto2) <= 67 Then
                    Call objBemaNaoFiscal.ImprimeFormatado(sTexto1 & String(67 - iAux, " ") & " " & sTexto2, 1, 0, 0, 0, 0)
                Else
                    Call objBemaNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", sTexto1))
                    Call objBemaNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_ESQUERDA, 48, " ", sTexto2))
                End If
            End If
            
            dPreco = objItens.dPrecoUnitario * objItens.dQuantidade
                        
            'se existir desconto sobre o item...
            If objItens.dPercDesc > 0 Then
                Call objBemaNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 15, " ", "DESC. ITEM:") & Formata_Campo(ALINHAMENTO_DIREITA, 10, " ", Format(objItens.dPercDesc, "standard") & "%") & "     =  " & Formata_Campo(ALINHAMENTO_ESQUERDA, 15, " ", Format(objItens.dPrecoUnitario * objItens.dQuantidade * (1 - (objItens.dPercDesc / 100)), "standard")))
                dPreco = dPreco - (objItens.dPrecoUnitario * objItens.dQuantidade * objItens.dPercDesc / 100)
            ElseIf objItens.dValorDesconto > 0 Then
                'Call objBemaNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 15, " ", "DESC. ITEM:") & Formata_Campo(ALINHAMENTO_DIREITA, 10, " ", Format(-objItens.dValorDesconto, "standard")) & "     =  " & Formata_Campo(ALINHAMENTO_ESQUERDA, 15, " ", Format((objItens.dPrecoUnitario - objItens.dValorDesconto) * objItens.dQuantidade, "standard")))
                'dPreco = dPreco - (objItens.dQuantidade * objItens.dValorDesconto)
                Call objBemaNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 15, " ", "DESC. ITEM:") & Formata_Campo(ALINHAMENTO_DIREITA, 10, " ", Format(-objItens.dValorDesconto, "standard")) & "     =  " & Formata_Campo(ALINHAMENTO_ESQUERDA, 15, " ", Format((objItens.dPrecoUnitario * objItens.dQuantidade - objItens.dValorDesconto), "standard")))
                dPreco = dPreco - objItens.dValorDesconto
            End If
            
            dTotal = dTotal + dPreco
        
        End If
        
    Next
                      
    Call objBemaNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", ""))
    Call objBemaNaoFiscal.ImprimeNormal(FormataTextoEspacosNoMeio("QTD. TOTAL DE ITENS", Format(iItem, "000"), 48))
    Call objBemaNaoFiscal.ImprimeNormal(FormataTextoEspacosNoMeio("VALOR TOTAL R$", Format(dTotal, "standard"), 48))
    
    If (objVenda.objCupomFiscal.dValorDesconto + objVenda.objCupomFiscal.dValorDesconto1) <> 0 Then
        Call objBemaNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_ESQUERDA, 48, " ", "DESCONTO  :  -" & Format(objVenda.objCupomFiscal.dValorDesconto + objVenda.objCupomFiscal.dValorDesconto1, "standard")))
        dTotal = dTotal - (objVenda.objCupomFiscal.dValorDesconto + objVenda.objCupomFiscal.dValorDesconto1)
    End If
    
    If objVenda.objCupomFiscal.dValorAcrescimo <> 0 Then
        Call objBemaNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_ESQUERDA, 48, " ", "ACRESCIMO :  +" & Format(objVenda.objCupomFiscal.dValorAcrescimo, "standard")))
        dTotal = dTotal + objVenda.objCupomFiscal.dValorAcrescimo
    End If
    
    If (objVenda.objCupomFiscal.dValorDesconto + objVenda.objCupomFiscal.dValorDesconto1) <> 0 Or objVenda.objCupomFiscal.dValorAcrescimo <> 0 Then
        Call objBemaNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_ESQUERDA, 48, " ", "TOTAL LIQ.:   " & Format(dTotal, "standard")))
    End If
    
    Call objBemaNaoFiscal.ImprimeNormal(FormataTextoEspacosNoMeio("FORMA DE PAGAMENTO", "Valor Pago", 48))
    For Each objMovCaixa In objVenda.colMovimentosCaixa
        If objMovCaixa.iTipo <> MOVIMENTOCAIXA_TROCO_DINHEIRO And objMovCaixa.iTipo <> MOVIMENTOCAIXA_TROCO_CONTRAVALE And objMovCaixa.iTipo <> MOVIMENTOCAIXA_TROCO_VALE Then
            Call objBemaNaoFiscal.ImprimeNormal(FormataTextoEspacosNoMeio(MeioPagtoCfe_Descricao(objMovCaixa.iCodigoCFe), Format(objMovCaixa.dValor, "standard"), 48))
        Else
            Call objBemaNaoFiscal.ImprimeNormal(FormataTextoEspacosNoMeio("Troco R$", Format(objMovCaixa.dValor, "standard"), 48))
        End If
    Next
    Call objBemaNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", ""))
    Call objBemaNaoFiscal.ImprimeNormal("Informação dos Tributos Totais Incidentes (Lei F")
    Call objBemaNaoFiscal.ImprimeNormal(FormataTextoEspacosNoMeio("ederal 12.741 /2012) R$", Format(Venda_ObterTotTrib(objVenda), "standard"), 48))
    Call objBemaNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", ""))
    
    sTexto1 = objVenda.objCupomFiscal.sNFCeMensagem
    If Len(sTexto1) <> 0 Then
        Do While Len(sTexto1) <> 0
            sTexto2 = left(sTexto1, 48)
            Call objBemaNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", sTexto2))
            sTexto1 = Mid(sTexto1, 49)
        Loop
        Call objBemaNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", ""))
    End If
    
    If objVenda.objCupomFiscal.iNFetpAmb = 2 Then
        objBemaNaoFiscal.ImprimeNormal (Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "EMITIDA EM AMBIENTE DE HOMOLOGACAO"))
        objBemaNaoFiscal.ImprimeNormal (Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "SEM VALOR FISCAL"))
    End If
    If objVenda.objCupomFiscal.sNFenProt = "" Then objBemaNaoFiscal.ImprimeNormal (Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "EMITIDA EM CONTINGENCIA"))
    
    Call objBemaNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "No.:" & CStr(objVenda.objCupomFiscal.lNumero) & " Serie:" & objVenda.objCupomFiscal.sNumSerieECF & " Emissao:" & Format(objVenda.objCupomFiscal.dtDataEmissao, "dd/mm/yyyy") & " " & Format(objVenda.objCupomFiscal.dHoraEmissao, "hh:mm:ss")))
    Call objBemaNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "Via Consumidor"))
    
    sChaveAcesso = objVenda.objCupomFiscal.sNFeChaveAcesso
    Call objBemaNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "Consulte pela chave de acesso em:"))
    Call objBemaNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", URLConsultaNFCE(left(sChaveAcesso, 2))))
    Call objBemaNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "CHAVE DE ACESSO"))
    Call objBemaNaoFiscal.ImprimeFormatado(Formata_Campo(ALINHAMENTO_CENTRALIZADO, 67, " ", Mid(sChaveAcesso, 1, 4) & " " & Mid(sChaveAcesso, 5, 4) & " " & Mid(sChaveAcesso, 9, 4) & " " & Mid(sChaveAcesso, 13, 4) & " " & Mid(sChaveAcesso, 17, 4) & " " & _
        Mid(sChaveAcesso, 21, 4) & " " & Mid(sChaveAcesso, 25, 4) & " " & Mid(sChaveAcesso, 29, 4) & " " & Mid(sChaveAcesso, 33, 4) & " " & Mid(sChaveAcesso, 37, 4) & " " & Mid(sChaveAcesso, 41, 4)), 1, 0, 0, 0, 0)
    
    Call objBemaNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", ""))
    
    If Len(objVenda.objCupomFiscal.sCPFCGC) <> 0 Then
        Call Formata_CNPJ_CPF(sCPFCGC, objVenda.objCupomFiscal.sCPFCGC)
        Call objBemaNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "CNPJ/CPF/Id.Estrangeiro: " & sCPFCGC))
        
        Call objBemaNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "  Nome: " & objVenda.objCupomFiscal.sNomeCliente))
        
        If Len(Trim(objVenda.objCupomFiscal.sEndEntLogradouro)) <> 0 Then
            Call objBemaNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "  End.: " & objVenda.objCupomFiscal.sEndEntLogradouro & ", " & objVenda.objCupomFiscal.sEndEntNúmero))
            Call objBemaNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "Compl.: " & objVenda.objCupomFiscal.sEndEntComplemento))
            Call objBemaNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "Bairro: " & objVenda.objCupomFiscal.sEndEntBairro))
        End If
        
    Else
        Call objBemaNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "CONSUMIDOR NAO IDENTIFICADO"))
    End If
    
    Call objBemaNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", ""))
    
    Call objBemaNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "Consulte via leitor de QR Code"))
    Call objBemaNaoFiscal.ImprimeQRCode(objVenda.objCupomFiscal.sNFCeQRCode, objVenda.objCupomFiscal.sNFeChaveAcesso)
    
    If objVenda.objCupomFiscal.sNFenProt <> "" Then
        Call objBemaNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "Protocolo de Autorização:" & " " & objVenda.objCupomFiscal.sNFenProt))
        Call objBemaNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", Format(objVenda.objCupomFiscal.dtNFeData, "dd/mm/yyyy") & " " & Format(objVenda.objCupomFiscal.dNFeHora, "hh:mm:ss")))
    End If
    
    Call objBemaNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", ""))
    Call objBemaNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", ""))
    Call objBemaNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", ""))
    Call objBemaNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", ""))
    Call objBemaNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", ""))
    Call objBemaNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", ""))
    Call objBemaNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", ""))
    Call objBemaNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", ""))
        
    Call objBemaNaoFiscal.TerminaImpressao
    
    Call BematechNaoFiscal_AcionaGuilhotina(0) 'disponível somente para interface Paralela
    
    BematechNaoFiscal_ImprimirNFCE = SUCESSO
    
    Exit Function
    
Erro_BematechNaoFiscal_ImprimirNFCE:

    BematechNaoFiscal_ImprimirNFCE = gErr

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 201571)

    End Select

    Exit Function

End Function

Function NFCE_Verifica_NFCeInfo(ByVal objVenda As ClassVenda, Optional ByVal bCancelada As Boolean = False) As Long
'verifica pela chave de acesso se nfce está gravada na tabela NFCeInfo

Dim lErro As Long
Dim lComando As Long, dtCancData As Date, iCancPendente As Integer

On Error GoTo Erro_NFCE_Verifica_NFCeInfo

    lComando = Comando_AbrirExt(glConexaoPAFECF)
    If lComando = 0 Then gError 201588

    lErro = Comando_Executar(lComando, "SELECT CancData, CancPendente FROM NFCeInfo WHERE chNFe = ?", dtCancData, iCancPendente, objVenda.objCupomFiscal.sNFeChaveAcesso)
    If lErro <> AD_SQL_SUCESSO Then gError 201589

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201590

    'nfce já deveria estar gravada
    If lErro = AD_SQL_SEM_DADOS Then gError 201561
    
    If bCancelada And iCancPendente = 0 And dtCancData = DATA_NULA Then gError 201563
    
    Call Comando_Fechar(lComando)
    
    NFCE_Verifica_NFCeInfo = SUCESSO

    Exit Function

Erro_NFCE_Verifica_NFCeInfo:

    NFCE_Verifica_NFCeInfo = gErr

    Select Case gErr

        Case 201588
            Call Rotina_ErroECF(vbOKOnly, ERRO_ABERTURA_COMANDO, gErr)

        Case 201589, 201560
            Call Rotina_ErroECF(vbOKOnly, ERRO_LEITURA_NFCEINFO, gErr)

        Case 201561
            Call Rotina_ErroECF(vbOKOnly, ERRO_NFCEINFO_NAO_CADASTRADO, gErr, objVenda.objCupomFiscal.sNFeChaveAcesso)

        Case 201563
            Call Rotina_ErroECF(vbOKOnly, ERRO_NFCEINFO_CANC_NAO_CADASTRADO, gErr, objVenda.objCupomFiscal.sNFeChaveAcesso)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 201562)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function NFCE_Imprime(ByVal objTela As Object, ByVal objVenda As ClassVenda) As Long
'imprime o NFCE

Dim lErro As Long
Dim colMensagem As New Collection
Dim iIndice As Integer
Dim sMensagem As String, sNFCeMensagem As String, sMsgTotTrib As String

On Error GoTo Erro_NFCE_Imprime
    
    lErro = NFCE_Verifica_NFCeInfo(objVenda)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If Len(Trim(objVenda.objCupomFiscal.sNFCeMensagem)) = 0 Then
    
        If gsUF = "RJ" Then sNFCeMensagem = "PROCON - R. da Ajuda 5 - RJ - (21) 151. ALERJ - Rua 1o de Março - RJ - (21) 25881418."
        
        'Mensagem referente ao total aproximado dos tributos lei 12741
        Call CF_ECF("Venda_MsgValAproxTrib", objVenda, sMsgTotTrib)
        
        objVenda.objCupomFiscal.sNFCeMensagem = sNFCeMensagem

    End If
    
    lErro = AFRAC_AbrirRelatorioGerencial(RELGER_NFCE, objTela)
    lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Abrir NFCE")
    If lErro <> SUCESSO Then gError 133088
    
    'Estava imprimindo tudo "Outros" na forma de pagamento porque estava indo zerado
    If objVenda.colMovimentosCaixa.Count <> 0 Then
        If objVenda.colMovimentosCaixa.Item(1).iCodigoCFe = 0 Then
            'colocar o CodigoCFe devido nos meios de pagto
            Call CF_ECF("Venda_CalculaCodigoCFe", objVenda)
        End If
    End If

    Select Case gobjNFeInfo.iModeloImpressora
        
        Case IMPRESSORA_EPSON_TM_T20
        
            lErro = EpsonNaoFiscal_ImprimirNFCE(objVenda)
        
        Case IMPRESSORA_ELGIN_VOX, IMPRESSORA_BEMATECH_ESCBEMA, IMPRESSORA_EPSON_ESCPOS, IMPRESSORA_DARUMA_ESCPOS
        
            lErro = ImpressoraESC_ImprimirNFCE(objVenda)
        
        Case IMPRESSORA_BEMATECH_MP4000_TH
        
            lErro = BematechNaoFiscal_ImprimirNFCE(objVenda)
        
        Case IMPRESSORA_DARUMA_DR700
        
            If Len(Trim(objVenda.objCupomFiscal.sNFenProt)) <> 0 Then
                lErro = iCFImprimir_NFCe_Daruma(objVenda.objCupomFiscal.sNFeArqXml, "", objVenda.objCupomFiscal.sNFCeQRCode, 48, 1)
            Else
                lErro = iCFImprimir_NFCe_Daruma(objVenda.objCupomFiscal.sNFeArqXmlPre, "OFFLINE", objVenda.objCupomFiscal.sNFCeQRCode, 48, 1)
            End If
        
        Case Else
        
            lErro = NFCE_Imprime1(objVenda, colMensagem)
            If lErro <> SUCESSO Then gError 133088
    
            For iIndice = 1 To colMensagem.Count
                    
                sMensagem = colMensagem.Item(iIndice)
                If sMensagem = "" Then sMensagem = "    "
                lErro = AFRAC_ImprimirRelatorioGerencial(sMensagem, objTela)
                lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "NFCE")
                If lErro <> SUCESSO Then gError 99893
                    
            Next
        
    End Select

    lErro = AFRAC_FecharRelatorioGerencial(objTela)
    lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Fecha NFCE")
    If lErro <> SUCESSO Then gError 99898
    
    lErro = CF_ECF("TEF_Imprime_Gerencial2_PAYGO", objVenda, objTela)
    If lErro <> SUCESSO Then gError 99898
    
    NFCE_Imprime = SUCESSO
    
    Exit Function
    
Erro_NFCE_Imprime:

    NFCE_Imprime = gErr

    Select Case gErr

        Case 99803, 99894 To 99898, 105868, 133088, 204319, 204324, 204329, 210470, ERRO_SEM_MENSAGEM

        Case 99893
            Call AFRAC_FecharRelatorioGerencial(objTela)

        Case 210471
            Call Rotina_ErroECF(vbOKOnly, ERRO_ORCAMENTO_NAO_PREENCHIDO, lErro)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149439)

    End Select

    Exit Function

End Function

Private Function NFCE_Imprime1(ByVal objVenda As ClassVenda, colMensagem As Collection) As Long
'imprime o NFCE

Dim lErro As Long
Dim objItens As ClassItemCupomFiscal
Dim objProduto As ClassProduto
Dim iIndice As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim bCont As Boolean
Dim sMensagem As String
Dim dTotal As Double
Dim dPreco As Double
Dim bAchou As Boolean
Dim colOrcamento As New Collection
Dim iIndice1 As Integer
Dim objCliente As ClassCliente
Dim lIndice As Long
Dim sCliente As String
Dim sCPFCGC As String
Dim sVendedor As String
Dim objVendedor As ClassVendedor
Dim sProduto As String
Dim sCNPJ As String
Dim iPos As Integer
Dim lNumDAV As Long
Dim sRetorno As String
Dim objProdutoNomeRed As TextBox, objMovCaixa As ClassMovimentoCaixa
Dim sTexto As String, sChaveAcesso As String, sSATtimeStamp As String, sTexto1 As String, sTexto2 As String, iAux As Integer, iItem As Integer

On Error GoTo Erro_NFCE_Imprime1

    Set colMensagem = New Collection
    
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", gsNomeEmpresa)
    Call Formata_CNPJ_CPF(sCNPJ, gsCNPJ)
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "CNPJ:" & sCNPJ & " IE:" & gsInscricaoEstadual & " IM:" & gsInscricaoMunicipal)
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", gsEndereco)
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", "")
    colMensagem.Add Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "DANFE NFC-e - Documento Auxiliar")
    colMensagem.Add Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "da Nota Fiscal de Consumidor Eletronica")
    colMensagem.Add Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "Não permite aproveitamento de crédito de ICMS")
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", "")
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "Codigo Descricao         Qtde Un  VlrUnit VlrTot")
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", "")

    iItem = 0
    
    For Each objItens In objVenda.objCupomFiscal.colItens
    
        If objItens.iStatus <> STATUS_CANCELADO Then
    
            iItem = iItem + 1
    
            Set objProduto = New ClassProduto
    
            lErro = CF_ECF("Produtos_Le", objItens.sProduto, objProduto)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 214855
            
            If lErro = SUCESSO Then
                objItens.sProdutoNomeRed = objProduto.sNomeReduzido
            Else
                'caso o produto não seja encontrado
                gError 99894
            End If
                
            sTexto1 = objProduto.sCodigo & " " & objProduto.sDescricao
            sTexto1 = left(sTexto1, 48)
                
            sTexto2 = Format(objItens.dQuantidade, "0.0###")
            If right(sTexto2, 2) = ",0" Then sTexto2 = left(sTexto2, Len(sTexto2) - 2)
            sTexto2 = sTexto2 & " " & objItens.sUnidadeMed
            sTexto2 = sTexto2 & " X " & Format(objItens.dPrecoUnitario, "standard")
            sTexto2 = sTexto2 & " " & Format(objItens.dPrecoUnitario * objItens.dQuantidade, "standard")
                
            iAux = Len(sTexto1 & " " & sTexto2)
            If Len(sTexto1 & " " & sTexto2) <= 48 Then
                colMensagem.Add sTexto1 & String(48 - iAux, " ") & " " & sTexto2
            Else
                colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", sTexto1)
                colMensagem.Add Formata_Campo(ALINHAMENTO_ESQUERDA, 48, " ", sTexto2)
            End If
            
            dPreco = objItens.dPrecoUnitario * objItens.dQuantidade
                        
            'se existir desconto sobre o item...
            If objItens.dPercDesc > 0 Then
                colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 15, " ", "DESC. ITEM:") & Formata_Campo(ALINHAMENTO_DIREITA, 10, " ", Format(objItens.dPercDesc, "standard") & "%") & "     =  " & Formata_Campo(ALINHAMENTO_ESQUERDA, 15, " ", Format(objItens.dPrecoUnitario * objItens.dQuantidade * (1 - (objItens.dPercDesc / 100)), "standard"))
                dPreco = dPreco - (objItens.dPrecoUnitario * objItens.dQuantidade * objItens.dPercDesc / 100)
            ElseIf objItens.dValorDesconto > 0 Then
                'colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 15, " ", "DESC. ITEM:") & Formata_Campo(ALINHAMENTO_DIREITA, 10, " ", Format(-objItens.dValorDesconto, "standard")) & "     =  " & Formata_Campo(ALINHAMENTO_ESQUERDA, 15, " ", Format((objItens.dPrecoUnitario - objItens.dValorDesconto) * objItens.dQuantidade, "standard"))
                'dPreco = dPreco - (objItens.dQuantidade * objItens.dValorDesconto)
                colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 15, " ", "DESC. ITEM:") & Formata_Campo(ALINHAMENTO_DIREITA, 10, " ", Format(-objItens.dValorDesconto, "standard")) & "     =  " & Formata_Campo(ALINHAMENTO_ESQUERDA, 15, " ", Format((objItens.dPrecoUnitario * objItens.dQuantidade - objItens.dValorDesconto), "standard"))
                dPreco = dPreco - objItens.dValorDesconto
            End If
            
            dTotal = dTotal + dPreco
        
        End If
        
    Next
                      
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", "")
    colMensagem.Add FormataTextoEspacosNoMeio("QTD. TOTAL DE ITENS", Format(iItem, "000"), 48)
    colMensagem.Add FormataTextoEspacosNoMeio("VALOR TOTAL R$", Format(dTotal, "standard"), 48)
    
    If (objVenda.objCupomFiscal.dValorDesconto + objVenda.objCupomFiscal.dValorDesconto1) <> 0 Then
        colMensagem.Add Formata_Campo(ALINHAMENTO_ESQUERDA, 48, " ", "DESCONTO  :  -" & Format(objVenda.objCupomFiscal.dValorDesconto + objVenda.objCupomFiscal.dValorDesconto1, "standard"))
        dTotal = dTotal - (objVenda.objCupomFiscal.dValorDesconto + objVenda.objCupomFiscal.dValorDesconto1)
    End If
    
    If objVenda.objCupomFiscal.dValorAcrescimo <> 0 Then
        colMensagem.Add Formata_Campo(ALINHAMENTO_ESQUERDA, 48, " ", "ACRESCIMO :  +" & Format(objVenda.objCupomFiscal.dValorAcrescimo, "standard"))
        dTotal = dTotal + objVenda.objCupomFiscal.dValorAcrescimo
    End If
    
    If (objVenda.objCupomFiscal.dValorDesconto + objVenda.objCupomFiscal.dValorDesconto1) <> 0 Or objVenda.objCupomFiscal.dValorAcrescimo <> 0 Then
        colMensagem.Add Formata_Campo(ALINHAMENTO_ESQUERDA, 48, " ", "TOTAL LIQ.:   " & Format(dTotal, "standard"))
    End If
    
    colMensagem.Add FormataTextoEspacosNoMeio("FORMA DE PAGAMENTO", "Valor Pago", 48)
    For Each objMovCaixa In objVenda.colMovimentosCaixa
        If objMovCaixa.iTipo <> MOVIMENTOCAIXA_TROCO_DINHEIRO And objMovCaixa.iTipo <> MOVIMENTOCAIXA_TROCO_CONTRAVALE And objMovCaixa.iTipo <> MOVIMENTOCAIXA_TROCO_VALE Then
            colMensagem.Add FormataTextoEspacosNoMeio(MeioPagtoCfe_Descricao(objMovCaixa.iCodigoCFe), Format(objMovCaixa.dValor, "standard"), 48)
        Else
            colMensagem.Add FormataTextoEspacosNoMeio("Troco R$", Format(objMovCaixa.dValor, "standard"), 48)
        End If
    Next
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", "")
    colMensagem.Add "Informação dos Tributos Totais Incidentes (Lei F"
    colMensagem.Add FormataTextoEspacosNoMeio("ederal 12.741 /2012) R$", Format(Venda_ObterTotTrib(objVenda), "standard"), 48)
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", "")
    
    sTexto1 = objVenda.objCupomFiscal.sNFCeMensagem
    If Len(sTexto1) <> 0 Then
        Do While Len(sTexto1) <> 0
            sTexto2 = left(sTexto1, 48)
            colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", sTexto2)
            sTexto1 = Mid(sTexto1, 49)
        Loop
        colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", "")
    End If
    
    If objVenda.objCupomFiscal.iNFetpAmb = 2 Then
        colMensagem.Add Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "EMITIDA EM AMBIENTE DE HOMOLOGACAO")
        colMensagem.Add Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "SEM VALOR FISCAL")
    End If
    If objVenda.objCupomFiscal.sNFenProt = "" Then colMensagem.Add Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "EMITIDA EM CONTINGENCIA")
    
    colMensagem.Add Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "No.:" & CStr(objVenda.objCupomFiscal.lNumero) & " Serie:" & objVenda.objCupomFiscal.sNumSerieECF & " Emissao:" & Format(objVenda.objCupomFiscal.dtDataEmissao, "dd/mm/yyyy") & " " & Format(objVenda.objCupomFiscal.dHoraEmissao, "hh:mm:ss"))
    colMensagem.Add Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "Via Consumidor")
    
    sChaveAcesso = objVenda.objCupomFiscal.sNFeChaveAcesso
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "Consulte pela chave de acesso em:")
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", URLConsultaNFCE(left(sChaveAcesso, 2)))
    colMensagem.Add Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "CHAVE DE ACESSO")
    colMensagem.Add Mid(sChaveAcesso, 1, 11) & " " & Mid(sChaveAcesso, 12, 11) & " " & Mid(sChaveAcesso, 23, 11) & " " & Mid(sChaveAcesso, 34, 11)
    
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", "")
    
    If Len(objVenda.objCupomFiscal.sCPFCGC) <> 0 Then
        Call Formata_CNPJ_CPF(sCPFCGC, objVenda.objCupomFiscal.sCPFCGC)
        colMensagem.Add Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "CNPJ/CPF/Id.Estrangeiro: " & sCPFCGC)
        
        colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "  Nome: " & objVenda.objCupomFiscal.sNomeCliente)
        
        If Len(Trim(objVenda.objCupomFiscal.sEndEntLogradouro)) <> 0 Then
            colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "  End.: " & objVenda.objCupomFiscal.sEndEntLogradouro & ", " & objVenda.objCupomFiscal.sEndEntNúmero)
            colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "Compl.: " & objVenda.objCupomFiscal.sEndEntComplemento)
            colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "Bairro: " & objVenda.objCupomFiscal.sEndEntBairro)
        End If
        
        
    Else
        colMensagem.Add Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "CONSUMIDOR NAO IDENTIFICADO")
    End If
    
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", "")
    
'''    colMensagem.Add Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "Consulte via leitor de QR Code")
'''    colMensagem.Add objVenda.objCupomFiscal.sNFCeQRCode
'''    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "")
'''    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "")
    If objVenda.objCupomFiscal.sNFenProt <> "" Then
        colMensagem.Add Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "Protocolo de Autorização:" & objVenda.objCupomFiscal.sNFenProt)
        colMensagem.Add Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", Format(objVenda.objCupomFiscal.dtNFeData, "dd/mm/yyyy") & " " & Format(objVenda.objCupomFiscal.dNFeHora, "hh:mm:ss"))
    End If
    
    NFCE_Imprime1 = SUCESSO
    
    Exit Function
    
Erro_NFCE_Imprime1:

    NFCE_Imprime1 = gErr

    Select Case gErr

        Case 99803, 99894 To 99898, 105868, 133088, 204319, 204324, 204329, 210470

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149439)

    End Select

    Exit Function

End Function

Public Function NFCeOffline_QtdePendente(lQtdeArqs As Long) As Long
'retorna a qtde de arquivos xml de nfces emitidas em contingencia offline que ainda precisam ser enviadas para a receita
    
Dim lErro As Long, lQtd As Long
Dim iIndice As Integer
Dim alComando(1 To 2) As Long

On Error GoTo Erro_NFCeOffline_QtdePendente
    
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_AbrirExt(glConexaoPAFECF)
        If alComando(iIndice) = 0 Then gError 201548
    Next
    
    lQtdeArqs = 0
    
    lErro = Comando_Executar(alComando(1), "SELECT COUNT(*) FROM NFCeInfo WHERE nProt = ''", lQtd)
    If lErro <> AD_SQL_SUCESSO Then gError 201549
    
    lErro = Comando_BuscarProximo(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201550
    
    lQtdeArqs = lQtd
    
    If lQtdeArqs <> 0 And Len(Trim(gobjLojaECF.sFTPURL)) > 0 Then

        lErro = Comando_Executar(alComando(2), "SELECT COUNT(*) FROM NFCeInfo WHERE nProt = '' AND XmlTransmitido = 0", lQtd)
        If lErro <> AD_SQL_SUCESSO Then gError 201549
        
        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201550
        
        If lQtd <> 0 Then gError 201586
    
    End If
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    NFCeOffline_QtdePendente = SUCESSO
    
    Exit Function
    
Erro_NFCeOffline_QtdePendente:

    NFCeOffline_QtdePendente = gErr

    Select Case gErr

        Case 201586

        Case 201548
            Call Rotina_ErroECF(vbOKOnly, ERRO_ABERTURA_COMANDO, gErr)
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 201551)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Private Function NFCE_ObterInfo(ByVal objCupomFiscal As ClassCupomFiscal) As Long
'pegar dados de nfceinfo como o nprot, para poder cancelar e/ou reimprimir, se necessário

Dim lErro As Long
Dim iIndice As Integer, snProt As String
Dim alComando(1 To 1) As Long, sQRCode As String, itpAmb As Integer, dHora As Double, dtData As Date

On Error GoTo Erro_NFCE_ObterInfo
    
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_AbrirExt(glConexaoPAFECF)
        If alComando(iIndice) = 0 Then gError 201548
    Next
    
    snProt = String(STRING_NFE_NPROT, 0)
    sQRCode = String(512, 0)
    lErro = Comando_Executar(alComando(1), "SELECT nProt, QRCode, tpAmb, data, hora FROM NFCeInfo WHERE chNFe = ?", snProt, sQRCode, itpAmb, dtData, dHora, objCupomFiscal.sNFeChaveAcesso)
    If lErro <> AD_SQL_SUCESSO Then gError 201549
    
    lErro = Comando_BuscarProximo(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201550
    
    If lErro = AD_SQL_SUCESSO Then
    
        objCupomFiscal.sNFenProt = snProt
        objCupomFiscal.sNFCeQRCode = sQRCode
        objCupomFiscal.iNFetpAmb = itpAmb
        objCupomFiscal.dtNFeData = dtData
        objCupomFiscal.dNFeHora = dHora
    
    End If
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    NFCE_ObterInfo = SUCESSO
    
    Exit Function
    
Erro_NFCE_ObterInfo:

    NFCE_ObterInfo = gErr

    Select Case gErr

        Case 201548
            Call Rotina_ErroECF(vbOKOnly, ERRO_ABERTURA_COMANDO, gErr)
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 201551)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Function NFCE_Imprime_Canc(ByVal objTela As Object, ByVal objVenda As ClassVenda) As Long
'imprime o NFCE

Dim lErro As Long
Dim colMensagem As New Collection
Dim iIndice As Integer
Dim sMensagem As String
Dim objBemaNaoFiscal As New ClassBematechNaoFiscalImp, objEpsonNaoFiscal As New ClassEpsonNaoFiscalImp
Dim objEscPos As New ClassImpressoraESC

On Error GoTo Erro_NFCE_Imprime_Canc
    
    lErro = NFCE_Verifica_NFCeInfo(objVenda, True)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = NFCE_Imprime_Canc1(objVenda, colMensagem)
    If lErro <> SUCESSO Then gError 133088
    
    lErro = AFRAC_AbrirRelatorioGerencial(RELGER_NFCE_CANC, objTela)
    lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Abrir NFCE Canc")
    If lErro <> SUCESSO Then gError 133088

    Select Case gobjNFeInfo.iModeloImpressora
    
        Case IMPRESSORA_DARUMA_DR700
            If Len(Trim(objVenda.objCupomFiscal.sNFenProt)) <> 0 Then
                lErro = iCFImprimir_NFCe_Daruma(Replace(objVenda.objCupomFiscal.sNFeArqXml, "-nfe.xml", "-ret-evento-canc.xml"), "", objVenda.objCupomFiscal.sNFCeQRCode, 48, 5)
            Else
                For iIndice = 1 To colMensagem.Count
                        
                    sMensagem = colMensagem.Item(iIndice)
                    If sMensagem = "" Then sMensagem = "    "
                    lErro = AFRAC_ImprimirRelatorioGerencial(sMensagem, objTela)
                    lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "NFCE")
                    If lErro <> SUCESSO Then gError 99893
                        
                Next
            End If
        
        Case IMPRESSORA_ELGIN_VOX, IMPRESSORA_BEMATECH_ESCBEMA, IMPRESSORA_EPSON_ESCPOS, IMPRESSORA_DARUMA_ESCPOS
            
            Call objEscPos.Define_Impressora(gobjNFeInfo.sPortaImpressora, gobjNFeInfo.iModeloImpressora)
            
            For iIndice = 1 To colMensagem.Count
                    
                sMensagem = colMensagem.Item(iIndice)
                If sMensagem = "" Then sMensagem = "    "
                Call objEscPos.ImprimeNormal(sMensagem)
                    
            Next
            
            Call objEscPos.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", ""))
            Call objEscPos.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", ""))
            
            Call objEscPos.AcionarGuilhotina("P")
            
            Call objEscPos.TerminaImpressao
        
        Case IMPRESSORA_BEMATECH_MP4000_TH
        
            For iIndice = 1 To colMensagem.Count
                    
                sMensagem = colMensagem.Item(iIndice)
                If sMensagem = "" Then sMensagem = "    "
                Call objBemaNaoFiscal.ImprimeNormal(sMensagem)
                    
            Next
            
            Call objBemaNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", ""))
            Call objBemaNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", ""))
            Call objBemaNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", ""))
            Call objBemaNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", ""))
            Call objBemaNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", ""))
            Call objBemaNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", ""))
            Call objBemaNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", ""))
            Call objBemaNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", ""))
            
            Call BematechNaoFiscal_AcionaGuilhotina(0) 'disponível somente para interface Paralela
            
            Call objBemaNaoFiscal.TerminaImpressao
            
        
        Case IMPRESSORA_EPSON_TM_T20
        
            For iIndice = 1 To colMensagem.Count
                    
                sMensagem = colMensagem.Item(iIndice)
                If sMensagem = "" Then sMensagem = "    "
                Call objEpsonNaoFiscal.ImprimeNormal(sMensagem)
                    
            Next
            
            Call objEpsonNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", ""))
            Call objEpsonNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", ""))
            Call objEpsonNaoFiscal.ImprimeNormal("<gui></gui>")
            
            Call objEpsonNaoFiscal.TerminaImpressao
            
        Case Else
        
            For iIndice = 1 To colMensagem.Count
                    
                sMensagem = colMensagem.Item(iIndice)
                If sMensagem = "" Then sMensagem = "    "
                lErro = AFRAC_ImprimirRelatorioGerencial(sMensagem, objTela)
                lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "NFCE")
                If lErro <> SUCESSO Then gError 99893
                    
            Next
    
    End Select

    lErro = AFRAC_FecharRelatorioGerencial(objTela)
    lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Fecha NFCE Canc")
    If lErro <> SUCESSO Then gError 99898
    
    NFCE_Imprime_Canc = SUCESSO
    
    Exit Function
    
Erro_NFCE_Imprime_Canc:

    NFCE_Imprime_Canc = gErr

    Select Case gErr

        Case 99803, 99894 To 99898, 105868, 133088, 204319, 204324, 204329, 210470

        Case 99893
            Call AFRAC_FecharRelatorioGerencial(objTela)

        Case 210471
            Call Rotina_ErroECF(vbOKOnly, ERRO_ORCAMENTO_NAO_PREENCHIDO, lErro)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149439)

    End Select

    Exit Function

End Function

Private Function NFCE_Imprime_Canc1(ByVal objVenda As ClassVenda, colMensagem As Collection) As Long
'imprime o DAV
'http://www.desenvolvedoresdaruma.com.br/home/downloads/Site_2011/Help/DarumaFrameworkHelpOnline/DarumaFramework/Mini_Impressora/Trabalhando_com_C%C3%B3digo_de_Barras_-_Impressora_N%C3%A3o-Fiscal.htm

Dim lErro As Long
Dim objItem As ClassItemCupomFiscal
Dim objProduto As ClassProduto
Dim iIndice As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim bCont As Boolean
Dim sMensagem As String
Dim dTotal As Double
Dim dPreco As Double
Dim bAchou As Boolean
Dim colOrcamento As New Collection
Dim iIndice1 As Integer
Dim objCliente As ClassCliente
Dim lIndice As Long
Dim sCliente As String
Dim sCPFCGC As String
Dim sVendedor As String
Dim objVendedor As ClassVendedor
Dim sProduto As String
Dim sCNPJ As String
Dim iPos As Integer
Dim lNumDAV As Long, objMovCaixa As ClassMovimentoCaixa
Dim sRetorno As String, sNFCEChaveAcesso As String
Dim objProdutoNomeRed As TextBox, sNFCEtimeStamp As String
Dim sTexto As String, sTexto2 As String, iTam As Integer

On Error GoTo Erro_NFCE_Imprime_Canc1

    Set colMensagem = New Collection
    
    Set objCliente = gobjClienteCPF.Busca(objVenda.objCupomFiscal.sCPFCGC, lIndice)
    If Not objCliente Is Nothing Then sCliente = objCliente.sNomeReduzido & " - "
    Call Formata_CNPJ_CPF(sCPFCGC, objVenda.objCupomFiscal.sCPFCGC)
    sCliente = sCliente & sCPFCGC

    Call Formata_CNPJ_CPF(sCNPJ, gsCNPJ)

    colMensagem.Add Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", gsNomeReduzido)
    colMensagem.Add Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", gsNomeEmpresa)
    colMensagem.Add Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", gsEndereco)
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "")
    colMensagem.Add Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "CNPJ:" & sCNPJ)
    colMensagem.Add Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "IE:" & gsInscricaoEstadual)
    colMensagem.Add Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "IM:" & gsInscricaoMunicipal)
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", "")

    colMensagem.Add Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "NFCE No. " & Format(objVenda.objCupomFiscal.lNumero, "000000"))
    colMensagem.Add Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "CANCELAMENTO")

    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", "")
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "CPF/CNPJ do Consumidor: " & sCPFCGC)

    colMensagem.Add Formata_Campo(ALINHAMENTO_ESQUERDA, 48, " ", "TOTAL R$:   " & Format(objVenda.objCupomFiscal.dValorTotal, "standard"))

    sNFCEChaveAcesso = objVenda.objCupomFiscal.sNFeChaveAcesso

    colMensagem.Add Mid(sNFCEChaveAcesso, 1, 11) & " " & Mid(sNFCEChaveAcesso, 12, 11) & " " & Mid(sNFCEChaveAcesso, 23, 11) & " " & Mid(sNFCEChaveAcesso, 34, 11)

    If objVenda.objCupomFiscal.sNFeCancnProt <> "" Then
        colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", "")
        colMensagem.Add Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "Protocolo de Cancelamento:" & objVenda.objCupomFiscal.sNFeCancnProt)
        colMensagem.Add Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", Format(objVenda.objCupomFiscal.dtNFeCancData, "dd/mm/yyyy") & " " & Format(objVenda.objCupomFiscal.dNFeCancHora, "hh:mm:ss"))
    End If
    
    NFCE_Imprime_Canc1 = SUCESSO
    
    Exit Function
    
Erro_NFCE_Imprime_Canc1:

    NFCE_Imprime_Canc1 = gErr

    Select Case gErr

        Case 99803, 99894 To 99898, 105868, 133088, 204319, 204324, 204329, 210470

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149439)

    End Select

    Exit Function

End Function

Private Function URLConsultaNFCE(ByVal sUF As String) As String

    Select Case sUF
    
        Case "33"
            URLConsultaNFCE = "http://nfce.fazenda.rj.gov.br/consulta"

        Case "12" 'Acre
            URLConsultaNFCE = "http://sefaznet.ac.gov.br/nfce/"
            
        Case "13" 'Amazonas
            URLConsultaNFCE = "http://sistemas.sefaz.am.gov.br/nfceweb/formConsulta.do"
            
        Case "28" 'Sergipe
            URLConsultaNFCE = "http://www.nfe.se.gov.br/portal/portalNoticias.jsp"
            
        Case "21" 'Maranhão
            URLConsultaNFCE = "http://www.nfce.sefaz.ma.gov.br/portal/consultaNFe.do?method=preFilterCupom&"
            
        Case "24" 'Rio Grande do Norte
            URLConsultaNFCE = "http://nfce.set.rn.gov.br/portalDFE/NFCe/ConsultaNFCe.aspx"
            
        Case "51" 'Mato Grosso
            URLConsultaNFCE = "http://www.sefaz.mt.gov.br/nfce/consultanfce"
            
        Case "43" 'Rio Grande do Sul
            URLConsultaNFCE = "https://www.sefaz.rs.gov.br/NFE/NFE-COM.aspx ou https://www.sefaz.rs.gov.br/NFE/NFE-NFC.aspx"
            
        Case "15" 'Pará
            URLConsultaNFCE = "https://appnfc.sefa.pa.gov.br/portal/view/consultas/nfce/consultanfce.seam"
            
        Case "25" 'Paraíba
            URLConsultaNFCE = "https://www3.receita.pb.gov.br/servirtual/cidadao/documentos-fiscais/consultar-nfc-e"
            
        Case "11" 'Rondônia
            URLConsultaNFCE = "http://www.nfce.sefin.ro.gov.br/home.jsp"
            
        Case "14" 'Roraima
            URLConsultaNFCE = "https://www.sefaz.rr.gov.br/nfce/servlet/wp_consulta_nfce"
            
        Case "22" 'Piauí
            URLConsultaNFCE = "http://webas.sefaz.pi.gov.br/nfceweb/consultarNFCe.jsf"
            
        Case "41" 'Paraná
            URLConsultaNFCE = "http://www.fazenda.pr.gov.br/"
            
        Case "29" 'Bahia
            URLConsultaNFCE = "http://nfe.sefaz.ba.gov.br/servicos/nfce/Modulos/Geral/NFCEC_consulta_chave_acesso.aspx"
            
        Case "35" 'SP
            URLConsultaNFCE = "http://www.nfce.fazenda.sp.gov.br/NFCePortal/Default.aspx"
        
        Case Else
            URLConsultaNFCE = ""


    End Select
    
End Function

Public Function Produtos_Le_CodBarras(ByVal sCodBarras As String, ByVal objProduto As ClassProduto) As Long

Dim lErro As Long
Dim tProduto As typeProduto
Dim alComando(1 To 1) As Long
Dim iIndice As Integer

On Error GoTo Erro_Produtos_Le_CodBarras

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_AbrirExt(glConexaoPAFECF)
        If alComando(iIndice) = 0 Then gError 214777
    Next

    tProduto.sCodigo = String(STRING_PRODUTO, 0)
    tProduto.sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
    tProduto.sNomeReduzido = String(STRING_PRODUTO_NOME_REDUZIDO, 0)
    tProduto.sSiglaUMVenda = String(STRING_PRODUTO_SIGLAUMVENDA, 0)
    tProduto.sReferencia = String(STRING_PRODUTO_REFERENCIA, 0)
    tProduto.sFigura = String(STRING_NOME_ARQ_COMPLETO, 0)
    tProduto.sSituacaoTribECF = String(STRING_PRODUTOFILIAL_SITUACAOTRIBECF, 0)
    tProduto.sICMSAliquota = String(STRING_PRODUTO_ICMSALIQUOTA, 0)
    tProduto.sTruncamento = String(STRING_PRODUTO_TRUNCAMENTO, 0)
        
    'obtem o conjunto de codigos de barra relativos ao produto em questao
    lErro = Comando_Executar(alComando(1), "SELECT Codigo, NomeReduzido, SiglaUMVenda, Referencia, Figura, PrecoLoja, SituacaoTribECF, " & _
    "ICMSAliquota, PercentMenosReceb, Descricao, DescontoValor, UsaBalanca, Compras, Truncamento, QuantEstLoja " & _
    " FROM Produtos As P, ProdutoCodBarras As B WHERE B.CodBarras = ? AND B.CodProduto = P.Codigo", tProduto.sCodigo, _
    tProduto.sNomeReduzido, tProduto.sSiglaUMVenda, tProduto.sReferencia, tProduto.sFigura, tProduto.dPrecoLoja, _
    tProduto.sSituacaoTribECF, tProduto.sICMSAliquota, tProduto.dPercentMenosReceb, tProduto.sDescricao, tProduto.dDesconto, _
    tProduto.iUsaBalanca, tProduto.iCompras, tProduto.sTruncamento, tProduto.dQuantEstLoja, sCodBarras)
    If lErro <> AD_SQL_SUCESSO Then gError 214778
    
    'se posiciona no 1o registro do conjunto obtido anteriormente
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 214779
    
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS
    
    objProduto.sCodigo = tProduto.sCodigo
    objProduto.sNomeReduzido = tProduto.sNomeReduzido
    objProduto.sSiglaUMVenda = tProduto.sSiglaUMVenda
    objProduto.sReferencia = tProduto.sReferencia
    objProduto.sFigura = tProduto.sFigura
    objProduto.dPrecoLoja = tProduto.dPrecoLoja
    objProduto.sSituacaoTribECF = tProduto.sSituacaoTribECF
    objProduto.sICMSAliquota = tProduto.sICMSAliquota
    objProduto.dPercentMenosReceb = tProduto.dPercentMenosReceb
    objProduto.sDescricao = tProduto.sDescricao
    objProduto.dDescontoValor = tProduto.dDesconto
    objProduto.iUsaBalanca = tProduto.iUsaBalanca
    objProduto.iCompras = tProduto.iCompras
    objProduto.sTruncamento = tProduto.sTruncamento
    objProduto.dQuantEstLoja = tProduto.dQuantEstLoja
    
    lErro = ProdutoCodBarras_Le(tProduto.sCodigo, objProduto)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 214799
    
    lErro = TributacaoDocItem_Le(tProduto.sCodigo, objProduto)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 214800
    
    objProduto.sCodigoBarras = sCodBarras
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Produtos_Le_CodBarras = SUCESSO
       
    Exit Function

Erro_Produtos_Le_CodBarras:

    Produtos_Le_CodBarras = gErr

    Select Case gErr
        
        Case 214777
            Call Rotina_ErroECF(vbOKOnly, ERRO_ABERTURA_COMANDO, gErr)
            
        Case 214778, 214779
            Call Rotina_ErroECF(vbOKOnly, ERRO_LEITURA_PRODUTOCODBARRAS, gErr, sCodBarras)
        
        Case 214799, 214800, ERRO_LEITURA_SEM_DADOS
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 214781)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Public Function Produtos_Le(ByVal sCodigo As String, ByVal objProduto As ClassProduto) As Long

Dim lErro As Long
Dim sCodBarras As String
Dim tProduto As typeProduto
Dim alComando(1 To 1) As Long
Dim iIndice As Integer

On Error GoTo Erro_Produtos_Le

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_AbrirExt(glConexaoPAFECF)
        If alComando(iIndice) = 0 Then gError 214792
    Next

    tProduto.sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
    tProduto.sNomeReduzido = String(STRING_PRODUTO_NOME_REDUZIDO, 0)
    tProduto.sSiglaUMVenda = String(STRING_PRODUTO_SIGLAUMVENDA, 0)
    tProduto.sReferencia = String(STRING_PRODUTO_REFERENCIA, 0)
    tProduto.sFigura = String(STRING_NOME_ARQ_COMPLETO, 0)
    tProduto.sSituacaoTribECF = String(STRING_PRODUTOFILIAL_SITUACAOTRIBECF, 0)
    tProduto.sICMSAliquota = String(STRING_PRODUTO_ICMSALIQUOTA, 0)
    tProduto.sTruncamento = String(STRING_PRODUTO_TRUNCAMENTO, 0)
        
    'obtem o conjunto de codigos de barra relativos ao produto em questao
    lErro = Comando_Executar(alComando(1), "SELECT NomeReduzido, SiglaUMVenda, Referencia, Figura, PrecoLoja, SituacaoTribECF, " & _
    "ICMSAliquota, PercentMenosReceb, Descricao, DescontoValor, UsaBalanca, Compras, Truncamento, QuantEstLoja " & _
    " FROM Produtos WHERE Codigo = ?", _
    tProduto.sNomeReduzido, tProduto.sSiglaUMVenda, tProduto.sReferencia, tProduto.sFigura, tProduto.dPrecoLoja, _
    tProduto.sSituacaoTribECF, tProduto.sICMSAliquota, tProduto.dPercentMenosReceb, tProduto.sDescricao, tProduto.dDesconto, _
    tProduto.iUsaBalanca, tProduto.iCompras, tProduto.sTruncamento, tProduto.dQuantEstLoja, sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 214793
    
    'se posiciona no 1o registro do conjunto obtido anteriormente
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 214794
    
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS
    
    objProduto.sCodigo = sCodigo
    objProduto.sNomeReduzido = tProduto.sNomeReduzido
    objProduto.sSiglaUMVenda = tProduto.sSiglaUMVenda
    objProduto.sReferencia = tProduto.sReferencia
    objProduto.sFigura = tProduto.sFigura
    objProduto.dPrecoLoja = tProduto.dPrecoLoja
    objProduto.sSituacaoTribECF = tProduto.sSituacaoTribECF
    objProduto.sICMSAliquota = tProduto.sICMSAliquota
    objProduto.dPercentMenosReceb = tProduto.dPercentMenosReceb
    objProduto.sDescricao = tProduto.sDescricao
    objProduto.dDescontoValor = tProduto.dDesconto
    objProduto.iUsaBalanca = tProduto.iUsaBalanca
    objProduto.iCompras = tProduto.iCompras
    objProduto.sTruncamento = tProduto.sTruncamento
    objProduto.dQuantEstLoja = tProduto.dQuantEstLoja
    
    lErro = ProdutoCodBarras_Le(sCodigo, objProduto)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 214796
    
    lErro = TributacaoDocItem_Le(sCodigo, objProduto)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 214797
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Produtos_Le = SUCESSO
       
    Exit Function

Erro_Produtos_Le:

    Produtos_Le = gErr

    Select Case gErr
        
        Case 214792
            Call Rotina_ErroECF(vbOKOnly, ERRO_ABERTURA_COMANDO, gErr)
            
        Case 214793, 214794
            Call Rotina_ErroECF(vbOKOnly, ERRO_LEITURA_PRODUTOS_0, gErr, sCodigo)
            
        Case 214796, 214797, ERRO_LEITURA_SEM_DADOS
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 214798)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Private Function TributacaoDocItem_Le(ByVal sCodProduto As String, ByVal objProduto As ClassProduto) As Long

Dim lErro As Long
Dim alComando(1 To 1) As Long
Dim tTributacaoDocItem As typeTributacaoDocItem
Dim iIndice As Integer
Dim objTributacaoDocItem As New ClassTributacaoDocItem

On Error GoTo Erro_TributacaoDocItem_Le

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_AbrirExt(glConexaoPAFECF)
        If alComando(iIndice) = 0 Then gError 214782
    Next

    With tTributacaoDocItem
    
        'Alocação de espaço no buffer
        .sProduto = String(STRING_PRODUTO, 0)
        .sProdutoDescricao = String(STRING_ITEMNF_DESCRICAO, 0)
        .sUnidadeMed = String(STRING_UM_SIGLA, 0)
        .sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
        .sIPICodProduto = String(STRING_PRODUTO_IPI_CODIGO, 0)
        .sCST = String(STRING_CST, 0)
        .sISSQN = String(STRING_ISSQN_CODIGO, 0)
        .sISSCidadeIBGE = String(STRING_CIDADE_CODIBGE, 0)
        .sIPIEnquadramentoClasse = String(STRING_IPI_CLASSE_ENQUADRAMENTO, 0)
        .sIPIEnquadramentoCodigo = String(STRING_IPI_CODIGO_ENQUADRAMENTO, 0)
        .sIPICNPJProdutor = String(STRING_CGC, 0)
        .sIPISeloCodigo = String(STRING_IPI_SELO_CODIGO, 0)
        .sGenero = String(STRING_PRODUTOGENERO_CODIGO, 0)
        .sEAN = String(STRING_PRODUTO_CODIGO_BARRAS, 0)
        .sEANTrib = String(STRING_PRODUTO_CODIGO_BARRAS, 0)
        .sUMTrib = String(STRING_UM_SIGLA, 0)
        .sICMSUFDevidoST = String(STRING_ESTADO, 0)
        
        .sISSCST = String(STRING_CST, 0)
        .sCSOSN = String(STRING_CSOSN, 0)
        .sNatBCCred = String(STRING_NATBCCRED, 0)
        .sFCI = String(STRING_FCI, 0)
    
        .sISSCodServ = String(STRING_PRODUTO_ISS_CODSERV, 0)
        .sISSMunicIncidImp = String(STRING_CIDADE_CODIBGE, 0)
        .sISSNumProcesso = String(STRING_ISS_NUMPROCESSO, 0)
        .sCEST = String(STRING_MAXIMO, 0)
        .scBenef = String(STRING_MAXIMO, 0)
        .sindEscala = String(STRING_MAXIMO, 0)
        .sCNPJFab = String(STRING_MAXIMO, 0)
    
        'Le a tabelaTributacaoDocItem
        lErro = Comando_Executar(alComando(1), "SELECT Quantidade, PrecoUnitario, NaturezaOp, TipoTributacao, ICMSTipo, ICMSBase, ICMSPercRedBase, ICMSAliquota, ICMSValor, " & _
                    "PISTipo, PISTipoCalculo, PISBase, PISAliquota, PISAliquotaValor, PISQtde, " & _
                    "PISValor, PISSTTipoCalculo, PISSTBase, PISSTAliquota, PISSTAliquotaValor, PISSTQtde, PISSTValor, COFINSTipo, COFINSTipoCalculo, COFINSBase, COFINSAliquota, " & _
                    "COFINSAliquotaValor, COFINSQtde, COFINSValor, COFINSSTTipoCalculo, COFINSSTBase, COFINSSTAliquota, COFINSSTAliquotaValor, COFINSSTQtde, COFINSSTValor, CST, " & _
                    "ISSQN, ISSBase, ISSAliquota, ISSCidadeIBGE, ISSValor, " & _
                    "ICMSBaseModalidade, OrigemMercadoria, EAN, EANTrib, QtdTrib, UMTrib, ValorUnitTrib, ISSCST,ICMSValorIsento,ICMSMotivo, ICMSSimplesTipo, " & _
                    "ISSTipo,  RegimeTributario, CSOSN, TotTrib, TotTribTipo, IPICodProduto, CEST, cBenef, indEscala, CNPJFab, ICMSvBCFCP, ICMSpFCP, ICMSvFCP " & _
                    "FROM TributacaoDocItem WHERE Produto = ? ", _
                    .dQuantidade, .dPrecoUnitario, .sNaturezaOp, .iTipoTributacao, _
                    .iICMSTipo, .dICMSBase, .dICMSPercRedBase, .dICMSAliquota, .dICMSValor, .iPISTipo, .iPISTipoCalculo, .dPISBase, .dPISAliquota, .dPISAliquotaValor, .dPISQtde, _
                    .dPISValor, .iPISSTTipoCalculo, .dPISSTBase, .dPISSTAliquota, .dPISSTAliquotaValor, .dPISSTQtde, .dPISSTValor, .iCOFINSTipo, .iCOFINSTipoCalculo, .dCOFINSBase, .dCOFINSAliquota, _
                    .dCOFINSAliquotaValor, .dCOFINSQtde, .dCOFINSValor, .iCOFINSSTTipoCalculo, .dCOFINSSTBase, .dCOFINSSTAliquota, .dCOFINSSTAliquotaValor, .dCOFINSSTQtde, .dCOFINSSTValor, .sCST, _
                    .sISSQN, .dISSBase, .dISSAliquota, .sISSCidadeIBGE, .dISSValor, .iICMSBaseModalidade, .iOrigemMercadoria, .sEAN, .sEANTrib, .dQtdTrib, .sUMTrib, .dValorUnitTrib, _
                    .sISSCST, .dICMSValorIsento, .iICMSMotivo, .iICMSSimplesTipo, .iISSTipo, .iRegimeTributario, .sCSOSN, .dTotTrib, .iTotTribTipo, .sIPICodProduto, .sCEST, .scBenef, .sindEscala, .sCNPJFab, .dICMSvBCFCP, .dICMSpFCP, .dICMSvFCP, sCodProduto)
    End With
    If lErro <> AD_SQL_SUCESSO Then gError 214783

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 214784

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS
    
    Call objTributacaoDocItem.Copia2(tTributacaoDocItem)

    Set objProduto.objTributacaoDocItem = objTributacaoDocItem
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TributacaoDocItem_Le = SUCESSO

    Exit Function

Erro_TributacaoDocItem_Le:

    TributacaoDocItem_Le = gErr

    Select Case gErr

        Case 214782
            Call Rotina_ErroECF(vbOKOnly, ERRO_ABERTURA_COMANDO, gErr)

        Case 214783, 214784
            Call Rotina_ErroECF(vbOKOnly, ERRO_LEITURA_TRIBUTACAODOCITEM, gErr, sCodProduto)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 214785)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Private Function ProdutoCodBarras_Le(ByVal sCodigo As String, ByVal objProduto As ClassProduto) As Long

Dim lErro As Long
Dim sCodBarras As String
Dim tProduto As typeProduto
Dim alComando(1 To 1) As Long
Dim iIndice As Integer

On Error GoTo Erro_ProdutoCodBarras_Le

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_AbrirExt(glConexaoPAFECF)
        If alComando(iIndice) = 0 Then gError 214786
    Next

    sCodBarras = String(STRING_PRODUTO_CODIGO_BARRAS, 0)
    
    'obtem o conjunto de codigos de barra relativos ao produto em questao
    lErro = Comando_Executar(alComando(1), "SELECT CodBarras FROM ProdutoCodBarras WHERE CodProduto = ?", sCodBarras, sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 214787
    
    'se posiciona no 1o registro do conjunto obtido anteriormente
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 214788
    
    If lErro = AD_SQL_SUCESSO Then
    
        objProduto.sCodigoBarras = sCodBarras
    
    End If
    
    Do While lErro = AD_SQL_SUCESSO
    
        objProduto.colCodBarras.Add sCodBarras
    
        'se posiciona no 1o registro do conjunto obtido anteriormente
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 214789
    
    Loop

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    ProdutoCodBarras_Le = SUCESSO
       
    Exit Function

Erro_ProdutoCodBarras_Le:

    ProdutoCodBarras_Le = gErr

    Select Case gErr
        
        Case 214786
            Call Rotina_ErroECF(vbOKOnly, ERRO_ABERTURA_COMANDO, gErr)
            
        Case 214787, 214788, 214789
            Call Rotina_ErroECF(vbOKOnly, ERRO_LEITURA_PRODUTOCODBARRAS, gErr, sCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 214790)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
            
End Function

Public Function Produtos_Le_NomeReduzido(ByVal sNomeReduzido As String, ByVal objProduto As ClassProduto) As Long

Dim lErro As Long
Dim sCodBarras As String
Dim tProduto As typeProduto
Dim alComando(1 To 1) As Long
Dim iIndice As Integer

On Error GoTo Erro_Produtos_Le_NomeReduzido

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_AbrirExt(glConexaoPAFECF)
        If alComando(iIndice) = 0 Then gError 214801
    Next

    tProduto.sCodigo = String(STRING_PRODUTO, 0)
    tProduto.sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
    tProduto.sNomeReduzido = String(STRING_PRODUTO_NOME_REDUZIDO, 0)
    tProduto.sSiglaUMVenda = String(STRING_PRODUTO_SIGLAUMVENDA, 0)
    tProduto.sReferencia = String(STRING_PRODUTO_REFERENCIA, 0)
    tProduto.sFigura = String(STRING_NOME_ARQ_COMPLETO, 0)
    tProduto.sSituacaoTribECF = String(STRING_PRODUTOFILIAL_SITUACAOTRIBECF, 0)
    tProduto.sICMSAliquota = String(STRING_PRODUTO_ICMSALIQUOTA, 0)
    tProduto.sTruncamento = String(STRING_PRODUTO_TRUNCAMENTO, 0)
        
    'obtem o conjunto de codigos de barra relativos ao produto em questao
    lErro = Comando_Executar(alComando(1), "SELECT Codigo, SiglaUMVenda, Referencia, Figura, PrecoLoja, SituacaoTribECF, " & _
    "ICMSAliquota, PercentMenosReceb, Descricao, DescontoValor, UsaBalanca, Compras, Truncamento, QuantEstLoja " & _
    " FROM Produtos WHERE NomeReduzido = ?", _
    tProduto.sCodigo, tProduto.sSiglaUMVenda, tProduto.sReferencia, tProduto.sFigura, tProduto.dPrecoLoja, _
    tProduto.sSituacaoTribECF, tProduto.sICMSAliquota, tProduto.dPercentMenosReceb, tProduto.sDescricao, tProduto.dDesconto, _
    tProduto.iUsaBalanca, tProduto.iCompras, tProduto.sTruncamento, tProduto.dQuantEstLoja, sNomeReduzido)
    If lErro <> AD_SQL_SUCESSO Then gError 214802
    
    'se posiciona no 1o registro do conjunto obtido anteriormente
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 214803
    
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS
    
    objProduto.sCodigo = tProduto.sCodigo
    objProduto.sNomeReduzido = sNomeReduzido
    objProduto.sSiglaUMVenda = tProduto.sSiglaUMVenda
    objProduto.sReferencia = tProduto.sReferencia
    objProduto.sFigura = tProduto.sFigura
    objProduto.dPrecoLoja = tProduto.dPrecoLoja
    objProduto.sSituacaoTribECF = tProduto.sSituacaoTribECF
    objProduto.sICMSAliquota = tProduto.sICMSAliquota
    objProduto.dPercentMenosReceb = tProduto.dPercentMenosReceb
    objProduto.sDescricao = tProduto.sDescricao
    objProduto.dDescontoValor = tProduto.dDesconto
    objProduto.iUsaBalanca = tProduto.iUsaBalanca
    objProduto.iCompras = tProduto.iCompras
    objProduto.sTruncamento = tProduto.sTruncamento
    objProduto.dQuantEstLoja = tProduto.dQuantEstLoja
    
    lErro = ProdutoCodBarras_Le(tProduto.sCodigo, objProduto)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 214804
    
    lErro = TributacaoDocItem_Le(tProduto.sCodigo, objProduto)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 214805
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Produtos_Le_NomeReduzido = SUCESSO
       
    Exit Function

Erro_Produtos_Le_NomeReduzido:

    Produtos_Le_NomeReduzido = gErr

    Select Case gErr
        
        Case 214801
            Call Rotina_ErroECF(vbOKOnly, ERRO_ABERTURA_COMANDO, gErr)
            
        Case 214802, 214803
            Call Rotina_ErroECF(vbOKOnly, ERRO_LEITURA_PRODUTOS_NOME, gErr, sNomeReduzido)
            
        Case 214804, 214805, ERRO_LEITURA_SEM_DADOS
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 214806)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function
            
Public Function Produtos_Le_Referencia(ByVal sReferencia As String, ByVal objProduto As ClassProduto) As Long

Dim lErro As Long
Dim sCodBarras As String
Dim tProduto As typeProduto
Dim alComando(1 To 1) As Long
Dim iIndice As Integer

On Error GoTo Erro_Produtos_Le_Referencia

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_AbrirExt(glConexaoPAFECF)
        If alComando(iIndice) = 0 Then gError 214807
    Next

    If giDebug = 1 Then MsgBox ("Produtos_Le_Referencia 1")

    tProduto.sCodigo = String(STRING_PRODUTO, 0)
    tProduto.sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
    tProduto.sNomeReduzido = String(STRING_PRODUTO_NOME_REDUZIDO, 0)
    tProduto.sSiglaUMVenda = String(STRING_PRODUTO_SIGLAUMVENDA, 0)
    tProduto.sReferencia = String(STRING_PRODUTO_REFERENCIA, 0)
    tProduto.sFigura = String(STRING_NOME_ARQ_COMPLETO, 0)
    tProduto.sSituacaoTribECF = String(STRING_PRODUTOFILIAL_SITUACAOTRIBECF, 0)
    tProduto.sICMSAliquota = String(STRING_PRODUTO_ICMSALIQUOTA, 0)
    tProduto.sTruncamento = String(STRING_PRODUTO_TRUNCAMENTO, 0)
        
    'obtem o conjunto de codigos de barra relativos ao produto em questao
    lErro = Comando_Executar(alComando(1), "SELECT Codigo, SiglaUMVenda, NomeReduzido, Figura, PrecoLoja, SituacaoTribECF, " & _
    "ICMSAliquota, PercentMenosReceb, Descricao, DescontoValor, UsaBalanca, Compras, Truncamento, QuantEstLoja " & _
    " FROM Produtos WHERE Referencia = ?", _
    tProduto.sCodigo, tProduto.sSiglaUMVenda, tProduto.sNomeReduzido, tProduto.sFigura, tProduto.dPrecoLoja, _
    tProduto.sSituacaoTribECF, tProduto.sICMSAliquota, tProduto.dPercentMenosReceb, tProduto.sDescricao, tProduto.dDesconto, _
    tProduto.iUsaBalanca, tProduto.iCompras, tProduto.sTruncamento, tProduto.dQuantEstLoja, sReferencia)
    If lErro <> AD_SQL_SUCESSO Then gError 214808
    
    'se posiciona no 1o registro do conjunto obtido anteriormente
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 214809
    
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS
    
    objProduto.sCodigo = tProduto.sCodigo
    objProduto.sNomeReduzido = tProduto.sNomeReduzido
    objProduto.sSiglaUMVenda = tProduto.sSiglaUMVenda
    objProduto.sReferencia = sReferencia
    objProduto.sFigura = tProduto.sFigura
    objProduto.dPrecoLoja = tProduto.dPrecoLoja
    objProduto.sSituacaoTribECF = tProduto.sSituacaoTribECF
    objProduto.sICMSAliquota = tProduto.sICMSAliquota
    objProduto.dPercentMenosReceb = tProduto.dPercentMenosReceb
    objProduto.sDescricao = tProduto.sDescricao
    objProduto.dDescontoValor = tProduto.dDesconto
    objProduto.iUsaBalanca = tProduto.iUsaBalanca
    objProduto.iCompras = tProduto.iCompras
    objProduto.sTruncamento = tProduto.sTruncamento
    objProduto.dQuantEstLoja = tProduto.dQuantEstLoja
    
    If giDebug = 1 Then MsgBox ("Produtos_Le_Referencia 2")
    
    lErro = ProdutoCodBarras_Le(tProduto.sCodigo, objProduto)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 214810
    
    If giDebug = 1 Then MsgBox ("Produtos_Le_Referencia 3")
    
    lErro = TributacaoDocItem_Le(tProduto.sCodigo, objProduto)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 214811
    
    If giDebug = 1 Then MsgBox ("Produtos_Le_Referencia 4")
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Produtos_Le_Referencia = SUCESSO
       
    Exit Function

Erro_Produtos_Le_Referencia:

    Produtos_Le_Referencia = gErr

    Select Case gErr
        
        Case 214807
            Call Rotina_ErroECF(vbOKOnly, ERRO_ABERTURA_COMANDO, gErr)
            
        Case 214808, 214809
            Call Rotina_ErroECF(vbOKOnly, ERRO_LEITURA_PRODUTOS_REFERENCIA, gErr, sReferencia)
            
        Case 214810, 214811, ERRO_LEITURA_SEM_DADOS
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 214812)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

'Function Carrega_Arquivo_FonteDados1() As Long
'
'Dim sRegistro As String
'Dim iPOS As Integer
'Dim sTipo As String
'Dim iPosInicio As Integer
'Dim lErro As Long
'Dim objProduto As ClassProduto
'Dim sArquivo As String
'Dim colTabPreco As New Collection
'Dim objTabPreco As New ClassTabelaPrecoItem
'Dim colProduto As New Collection
'Dim colProdutoDesconto As New Collection
'Dim lTamanho As Long
'Dim sRetorno As String
'Dim objObject As Object
'Dim lVersao As Long
'Dim lSequencial As Long
'
'On Error GoTo Erro_Carrega_Arquivo_FonteDados1
'
'
'    If giDebug = 1 Then MsgBox ("Carrega_Arquivo_FonteDados1 1")
'
'
'    lErro = CF_ECF("Caixa_Transacao_Abrir", lSequencial)
'    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
'
'    lErro = CF_ECF("Produto_Exclui_EmTrans")
'    If lErro <> SUCESSO Then gError 214817
'
'    lErro = CF_ECF("TributacaoDocItem_Exclui_EmTrans")
'    If lErro <> SUCESSO Then gError 214833
'
'
'    If giDebug = 1 Then MsgBox ("Carrega_Arquivo_FonteDados1 2")
'
''    Set gcolAdmMeioPagto = New Collection
''    Set gcolOperadores = New Collection
''    Set gobjLojaECF = New ClassLojaECF
''    gsNomeCaixa = ""
''    giPos = 0
''    giTEF = 0
''    giBoletoManual = 0
''    giTeclado = 0
''    giOrcamentoECF = 0
''    giCodECF = 0
''    giImpressoraECF = 0
''    giImpressoraCheque = 0
''    giCodModeloECF = 0
''    Set gcolECF = New Collection
''    Set gcolRedes = New Collection
''    Set gcolTeclados = New Collection
''    Set gcolTiposMeiosPagtos = New Collection
''    Set gcolCliente = New Collection
''    Set gobjClienteNome = New ClassClienteNome
''    Set gobjClienteCPF = New ClassClienteCPF
''    Set gcolVendedores = New Collection
''    Set gaobjProdutosNome = New ClassProdNome
''    Set gaobjProdutosCodBarras = New ClassProdCodBarra
''    Set gaobjProdutosReferencia = New ClassProdReferencia
'
'    lTamanho = 255
'    sRetorno = String(lTamanho, 0)
'
'    'funcao que verifica se o caixaconfig.ini foi alterado
'    lErro = CF_ECF("CaixaConfig_Verifica_Chave")
'    If lErro <> SUCESSO Then gError 214819
'
'    'Obtém o diretório onde deve ser armazenado o arquivo com dados do backoffice
'    Call GetPrivateProfileString(APLICACAO_DADOS, "DirDadosCC", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
'
'    'Retira os espaços no final da string
'    sRetorno = StringZ(sRetorno)
'
'    'Se não encontrou
'    If Len(Trim(sRetorno)) = 0 Or sRetorno = CStr(CONSTANTE_ERRO) Then gError 214820
'
'    If right(sRetorno, 1) <> "\" Then sRetorno = sRetorno & "\"
'
'    sArquivo = sRetorno & giCodEmpresa & "_" & giFilialEmpresa & "_" & NOME_ARQUIVOCC
'
'    'Abre o arquivo de retorno
'    Open sArquivo For Input As #1
'
'    'Busca o próximo registro do arquivo
'    Line Input #1, sRegistro
'
'    lVersao = StrParaLong(sRegistro)
'
'    If lVersao > VERSAO_ECF Then gError 214821
'
'    If giDebug = 1 Then MsgBox ("Carrega_Arquivo_FonteDados1 3")
'
'    'Até chegar ao fim do arquivo
'    Do While Not EOF(1)
'
'        'Busca o próximo registro do arquivo
'        Line Input #1, sRegistro
'
'        If sRegistro <> "" Then
'
'            iPosInicio = 1
'
'            'Procura o Primeiro Control para saber o tipo do registro
'            iPOS = InStr(iPosInicio, sRegistro, Chr(vbKeyControl))
'
'            sTipo = Mid(sRegistro, iPosInicio, iPOS - iPosInicio)
'
'            'Seleciona de acordo com o tipo
'            Select Case StrParaLong(sTipo)
'
'
''                Case TIPOREGISTROECF_CLIENTE
''                    lErro = Carrega_Dados_Cliente(sRegistro)
''                    If lErro <> SUCESSO Then gError 86253
'
'                Case TIPOREGISTROECF_PRODUTODESCONTO
'                    lErro = Carrega_Dados_ProdutosDesconto(sRegistro, colProdutoDesconto)
'                    If lErro <> SUCESSO Then gError 214822
'
''                Case TIPOREGISTROECF_VENDEDORES
''                    lErro = Carrega_Dados_Vendedores(sRegistro)
''                    If lErro <> SUCESSO Then gError 112552
'
'                Case TIPOREGISTROECF_PRODUTOSNOME
'                    lErro = Carrega_Dados_ProdutosNome1(sRegistro, colTabPreco, colProdutoDesconto)
'                    If lErro <> SUCESSO Then gError 214823
'
'                Case TIPOREGISTROECF_TABELAPRECOITENS
'                    lErro = Carrega_Dados_TabelaPreco(sRegistro, colTabPreco)
'                    If lErro <> SUCESSO And lErro <> 112562 Then gError 214824
'
'            End Select
'        End If
'    Loop
'
'    'Fechar a Transação
'    lErro = CF_ECF("Caixa_Transacao_Fechar", lSequencial)
'    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
'
'    If giDebug = 1 Then MsgBox ("Carrega_Arquivo_FonteDados1 4")
'
'    Close #1
'
'
'    Carrega_Arquivo_FonteDados1 = SUCESSO
'
'    Exit Function
'
'Erro_Carrega_Arquivo_FonteDados1:
'
'    Carrega_Arquivo_FonteDados1 = gErr
'
'    Select Case gErr
'
'        Case 53, 76
'            Call Rotina_ErroECF(vbOKOnly, ARQUIVO_INICIALIZACAO_NAO_EXISTENTE, gErr, sArquivo)
'
'        Case 214817 To 214819, 214822 To 214824, 214833, ERRO_SEM_MENSAGEM
'
'        Case 214820
'            Call Rotina_ErroECF(vbOKOnly, ERRO_LOCAL_ARQUIVO_NAO_CONFIGURADO, gErr, giCodEmpresa & "_" & giFilialEmpresa & "_" & NOME_ARQUIVOCC)
'
'        Case 214821
'            Call Rotina_ErroECF(vbOKOnly, "ERRO_VERSAO_ARQUIVO_MAIOR_PROGRAMA", gErr, lVersao, VERSAO_ECF)
'
'        Case Else
'            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error, 214825)
'
'    End Select
'
'    Call CF_ECF("Caixa_Transacao_Rollback", glTransacaoPAFECF)
'
'    Exit Function
'
'End Function

Private Function Carrega_Dados_ProdutosNome1(sRegistro As String, ByVal lVersao As Long, ByVal lCom1 As Long, ByVal lCom2 As Long, ByVal lCom3 As Long, ByVal lCom4 As Long, ByVal lCom5 As Long) As Long

Dim iPos As Integer
Dim objProduto As ClassProduto
Dim objProduto2 As ClassProduto
Dim iIndice As Integer
Dim iPosInicio As Integer
Dim iPosFim As Integer
Dim iPosMeio As Integer
Dim iPosColInicio As Integer
Dim sProduto As String
Dim objTabPreco As New ClassTabelaPrecoItem
Dim colTab As New Collection
Dim dValor As Double
Dim lErro As Long
Dim bAchou As Boolean
Dim sCodBarras As String, iIndiceCodBarras As Integer
Dim dPreco As Double, dPercentMenosReceb As Double, dDescontoValor As Double

On Error GoTo Erro_Carrega_Dados_ProdutosNome1

    Select Case lVersao
    
        Case VERSAO_ECF_304
            iIndiceCodBarras = 80
    
        Case VERSAO_ECF_303
            iIndiceCodBarras = 75
    
        Case VERSAO_ECF_INICIAL
            iIndiceCodBarras = 74
            
    End Select
    
    Set objProduto = New ClassProduto
            
    'Primeira Posição
    iPosInicio = 1
            
    'Procura o Primeiro Control para saber onde começa a string
    iPosInicio = InStr(iPosInicio, sRegistro, Chr(vbKeyControl)) + 1
    
    'Procura o Primeiro Escape dentro da String
    iPosMeio = (InStr(iPosInicio, sRegistro, Chr(vbKeyEscape)))
    
    'Pega última posição e guarda
    iPosFim = (InStr(iPosInicio, sRegistro, Chr(vbKeyEnd)))
    
    iIndice = 0
    
    Do While iPosMeio <> 0
        
       iIndice = iIndice + 1
        
        Select Case iIndice
            
            Case 1: objProduto.sCodigo = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
            Case 2: objProduto.sNomeReduzido = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
            Case 3: objProduto.sSiglaUMVenda = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
            Case 4: objProduto.sReferencia = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
            Case 5: objProduto.sFigura = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
            Case 6: objProduto.sSituacaoTribECF = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
            Case 7: objProduto.sICMSAliquota = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
            Case 8: objProduto.sDescricao = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
            Case 9: objProduto.iClasseUM = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 10: objProduto.sSiglaUMEstoque = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
            Case 11: objProduto.sSiglaUMCompra = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
            Case 12: objProduto.iAtivo = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 13: objProduto.dLimiteDesconto = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 14: objProduto.iUsaBalanca = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 15: objProduto.iCompras = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
                     objProduto.sTruncamento = gobjLojaECF.sTruncamentoArredondamento
            Case 16: objProduto.dQuantEstLoja = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            
            Case 17: objProduto.objTributacaoDocItem.dQuantidade = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 18: objProduto.objTributacaoDocItem.dPrecoUnitario = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 19: objProduto.objTributacaoDocItem.sNaturezaOp = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
            Case 20: objProduto.objTributacaoDocItem.iTipoTributacao = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 21: objProduto.objTributacaoDocItem.iICMSTipo = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 22: objProduto.objTributacaoDocItem.dICMSBase = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 23: objProduto.objTributacaoDocItem.dICMSPercRedBase = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 24: objProduto.objTributacaoDocItem.dICMSAliquota = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 25: objProduto.objTributacaoDocItem.dICMSValor = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            
            Case 26: objProduto.objTributacaoDocItem.iPISTipo = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 27: objProduto.objTributacaoDocItem.iPISTipoCalculo = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 28: objProduto.objTributacaoDocItem.dPISBase = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 29: objProduto.objTributacaoDocItem.dPISAliquota = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 30: objProduto.objTributacaoDocItem.dPISAliquotaValor = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 31: objProduto.objTributacaoDocItem.dPISQtde = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 32: objProduto.objTributacaoDocItem.dPISValor = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            
            Case 33: objProduto.objTributacaoDocItem.iPISSTTipoCalculo = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 34: objProduto.objTributacaoDocItem.dPISSTBase = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 35: objProduto.objTributacaoDocItem.dPISSTAliquota = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 36: objProduto.objTributacaoDocItem.dPISSTAliquotaValor = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 37: objProduto.objTributacaoDocItem.dPISSTQtde = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 38: objProduto.objTributacaoDocItem.dPISSTValor = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            
            Case 39: objProduto.objTributacaoDocItem.iCOFINSTipo = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 40: objProduto.objTributacaoDocItem.iCOFINSTipoCalculo = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 41: objProduto.objTributacaoDocItem.dCOFINSBase = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 42: objProduto.objTributacaoDocItem.dCOFINSAliquota = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 43: objProduto.objTributacaoDocItem.dCOFINSAliquotaValor = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 44: objProduto.objTributacaoDocItem.dCOFINSQtde = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 45: objProduto.objTributacaoDocItem.dCOFINSValor = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            
            Case 46: objProduto.objTributacaoDocItem.iCOFINSSTTipoCalculo = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 47: objProduto.objTributacaoDocItem.dCOFINSSTBase = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 48: objProduto.objTributacaoDocItem.dCOFINSSTAliquota = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 49: objProduto.objTributacaoDocItem.dCOFINSSTAliquotaValor = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 50: objProduto.objTributacaoDocItem.dCOFINSSTQtde = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 51: objProduto.objTributacaoDocItem.dCOFINSSTValor = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            
            Case 52: objProduto.objTributacaoDocItem.sCST = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
            Case 53: objProduto.objTributacaoDocItem.sISSQN = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
            Case 54: objProduto.objTributacaoDocItem.dISSBase = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 55: objProduto.objTributacaoDocItem.dISSAliquota = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 56: objProduto.objTributacaoDocItem.sISSCidadeIBGE = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
            Case 57: objProduto.objTributacaoDocItem.dISSValor = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            
            Case 58: objProduto.objTributacaoDocItem.iICMSBaseModalidade = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 59: objProduto.objTributacaoDocItem.iOrigemMercadoria = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 60: objProduto.objTributacaoDocItem.sGenero = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
            Case 61: objProduto.objTributacaoDocItem.dQtdTrib = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 62: objProduto.objTributacaoDocItem.sUMTrib = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
            Case 63: objProduto.objTributacaoDocItem.dValorUnitTrib = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            
            Case 64: objProduto.objTributacaoDocItem.sISSCST = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
            Case 65: objProduto.objTributacaoDocItem.dICMSValorIsento = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 66: objProduto.objTributacaoDocItem.iICMSMotivo = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 67: objProduto.objTributacaoDocItem.iICMSSimplesTipo = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 68: objProduto.objTributacaoDocItem.iISSTipo = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            
            Case 69: objProduto.objTributacaoDocItem.iRegimeTributario = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 70: objProduto.objTributacaoDocItem.sCSOSN = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
            Case 71: objProduto.objTributacaoDocItem.dTotTrib = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 72: objProduto.objTributacaoDocItem.iTotTribTipo = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 73: objProduto.objTributacaoDocItem.sIPICodProduto = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
                     
            Case Else
            
                If lVersao >= VERSAO_ECF_303 Then
                
                    Select Case iIndice
                    
                        Case 74: objProduto.objTributacaoDocItem.sCEST = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
                    
                    End Select
                    
                End If
                     
                If lVersao >= VERSAO_ECF_304 Then
                
                    Select Case iIndice
                    
                        Case 75: objProduto.objTributacaoDocItem.dICMSpFCP = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
                        Case 76: objProduto.objTributacaoDocItem.dICMSvFCP = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
                        Case 77: objProduto.objTributacaoDocItem.scBenef = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
                        Case 78: objProduto.objTributacaoDocItem.sindEscala = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
                        Case 79: objProduto.objTributacaoDocItem.sCNPJFab = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
                    
                    End Select
                    
                End If
                
                If iIndice >= iIndiceCodBarras Then
                
                    sCodBarras = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
                    If objProduto.sCodigoBarras = "" Then
                        objProduto.sCodigoBarras = sCodBarras
                        objProduto.objTributacaoDocItem.sEAN = objProduto.sCodigoBarras
                        objProduto.objTributacaoDocItem.sEANTrib = objProduto.sCodigoBarras
                    End If
                    
                    If sCodBarras <> "" Then
                        objProduto.colCodBarras.Add sCodBarras
                    End If
                
                End If
                
        End Select
        
        objProduto.objTributacaoDocItem.sProduto = objProduto.sCodigo
        
        'Atualiza as Posições
        iPosInicio = iPosMeio + 1
        iPosMeio = (InStr(iPosInicio, sRegistro, Chr(vbKeyEscape)))
        
'        If iPosInicio <= iPosFim And iPosMeio = 0 Then iPosMeio = iPosFim
                    
    Loop
    
    bAchou = False
    
'    For Each objTabPreco In colTabPreco
'        If UCase(objTabPreco.sCodProduto) = UCase(objProduto.sCodigo) Then
'            objProduto.dPrecoLoja = objTabPreco.dPreco
'            bAchou = True
'            Exit For
'        End If
'    Next
'
'    If bAchou Then
'        For Each objProduto2 In colProdutoDesconto
'            If objProduto2.sCodigo = objProduto.sCodigo Then
'                objProduto.dPercentMenosReceb = objProduto2.dPercentMenosReceb
'                objProduto.dDescontoValor = objProduto2.dDescontoValor
'                Exit For
'            End If
'        Next
'
'    End If

    'Busca o preço
    lErro = Comando_Executar(lCom4, "SELECT Preco FROM TabelaPrecoItens WHERE CodProduto = ? AND DataVigencia <= ? ORDER BY DataVigencia DESC", dPreco, objProduto.sCodigo, Date)
     If lErro <> AD_SQL_SUCESSO Then gError 216217
    
    lErro = Comando_BuscarPrimeiro(lCom4)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 216218

    If lErro = AD_SQL_SUCESSO Then bAchou = True
    
    If bAchou Then
    
        objProduto.dPrecoLoja = dPreco
    
        lErro = Comando_Executar(lCom5, "SELECT PercentMenosReceb, DescontoValor FROM ProdutoDesconto WHERE CodProduto = ? ", dPercentMenosReceb, dDescontoValor, objProduto.sCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 216219
        
        lErro = Comando_BuscarPrimeiro(lCom5)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 216220
    
        If lErro = AD_SQL_SUCESSO Then
            objProduto.dPercentMenosReceb = dPercentMenosReceb
            objProduto.dDescontoValor = dDescontoValor
        End If
    
        lErro = CF_ECF("Produto_Grava_EmTrans", objProduto, lCom1, lCom2, lCom3)
        If lErro <> SUCESSO Then gError 214816
    
    End If
    
    Carrega_Dados_ProdutosNome1 = SUCESSO
    
    Exit Function
    
Erro_Carrega_Dados_ProdutosNome1:
    
    Carrega_Dados_ProdutosNome1 = gErr
    
    Select Case gErr
    
        Case 216217, 216218
            Call Rotina_ErroECF(vbOKOnly, ERRO_LEITURA_TABELAPRECOITENS, gErr)
        
        Case 216219, 216220
            Call Rotina_ErroECF(vbOKOnly, ERRO_LEITURA_PRODUTODESCONTO, gErr)
    
        Case 214816
    
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, 149448)

    End Select
    
    Exit Function
    
End Function

Public Function Produtos_Le_NomeReduzido1(ByVal colProdutos As Collection) As Long

Dim lErro As Long
Dim sCodBarras As String
Dim tProduto As typeProduto
Dim alComando(1 To 1) As Long
Dim iIndice As Integer
Dim objProduto As New ClassProduto

On Error GoTo Erro_Produtos_Le_NomeReduzido1

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_AbrirExt(glConexaoPAFECF)
        If alComando(iIndice) = 0 Then gError 214837
    Next

    tProduto.sCodigo = String(STRING_PRODUTO, 0)
    tProduto.sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
    tProduto.sNomeReduzido = String(STRING_PRODUTO_NOME_REDUZIDO, 0)
    tProduto.sSiglaUMVenda = String(STRING_PRODUTO_SIGLAUMVENDA, 0)
    tProduto.sReferencia = String(STRING_PRODUTO_REFERENCIA, 0)
    tProduto.sFigura = String(STRING_NOME_ARQ_COMPLETO, 0)
    tProduto.sSituacaoTribECF = String(STRING_PRODUTOFILIAL_SITUACAOTRIBECF, 0)
    tProduto.sICMSAliquota = String(STRING_PRODUTO_ICMSALIQUOTA, 0)
    tProduto.sTruncamento = String(STRING_PRODUTO_TRUNCAMENTO, 0)
        
    'obtem o conjunto de codigos de barra relativos ao produto em questao
    lErro = Comando_Executar(alComando(1), "SELECT Codigo, NomeReduzido, SiglaUMVenda, Referencia, Figura, PrecoLoja, SituacaoTribECF, " & _
    "ICMSAliquota, PercentMenosReceb, Descricao, DescontoValor, UsaBalanca, Compras, Truncamento, QuantEstLoja " & _
    " FROM Produtos ORDER BY NomeReduzido", _
    tProduto.sCodigo, tProduto.sNomeReduzido, tProduto.sSiglaUMVenda, tProduto.sReferencia, tProduto.sFigura, tProduto.dPrecoLoja, _
    tProduto.sSituacaoTribECF, tProduto.sICMSAliquota, tProduto.dPercentMenosReceb, tProduto.sDescricao, tProduto.dDesconto, _
    tProduto.iUsaBalanca, tProduto.iCompras, tProduto.sTruncamento, tProduto.dQuantEstLoja)
    If lErro <> AD_SQL_SUCESSO Then gError 214838
    
    'se posiciona no 1o registro do conjunto obtido anteriormente
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 214839
    
    Do While lErro = AD_SQL_SUCESSO
    
        Set objProduto = New ClassProduto
    
        objProduto.sCodigo = tProduto.sCodigo
        objProduto.sNomeReduzido = tProduto.sNomeReduzido
        objProduto.sSiglaUMVenda = tProduto.sSiglaUMVenda
        objProduto.sReferencia = tProduto.sReferencia
        objProduto.sFigura = tProduto.sFigura
        objProduto.dPrecoLoja = tProduto.dPrecoLoja
        objProduto.sSituacaoTribECF = tProduto.sSituacaoTribECF
        objProduto.sICMSAliquota = tProduto.sICMSAliquota
        objProduto.dPercentMenosReceb = tProduto.dPercentMenosReceb
        objProduto.sDescricao = tProduto.sDescricao
        objProduto.dDescontoValor = tProduto.dDesconto
        objProduto.iUsaBalanca = tProduto.iUsaBalanca
        objProduto.iCompras = tProduto.iCompras
        objProduto.sTruncamento = tProduto.sTruncamento
        objProduto.dQuantEstLoja = tProduto.dQuantEstLoja
        
        lErro = ProdutoCodBarras_Le(tProduto.sCodigo, objProduto)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 214840
        
        colProdutos.Add objProduto
        
        'se posiciona no 1o registro do conjunto obtido anteriormente
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 214841
    
    Loop
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Produtos_Le_NomeReduzido1 = SUCESSO
       
    Exit Function

Erro_Produtos_Le_NomeReduzido1:

    Produtos_Le_NomeReduzido1 = gErr

    Select Case gErr
        
        Case 214837
            Call Rotina_ErroECF(vbOKOnly, ERRO_ABERTURA_COMANDO, gErr)
            
        Case 214838, 214839, 214841
            Call Rotina_ErroECF(vbOKOnly, ERRO_LEITURA_PRODUTOS_0, gErr)
            
        Case 214840
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 214842)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function


Public Function Produtos_Le_Count(lRows As Long) As Long

Dim lErro As Long
Dim sCodBarras As String
Dim tProduto As typeProduto
Dim alComando(1 To 1) As Long
Dim iIndice As Integer
Dim objProduto As ClassProduto

On Error GoTo Erro_Produtos_Le_Count

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_AbrirExt(glConexaoPAFECF)
        If alComando(iIndice) = 0 Then gError 214844
    Next

        
    'obtem o conjunto de codigos de barra relativos ao produto em questao
    lErro = Comando_Executar(alComando(1), "SELECT Count(*) FROM Produtos", lRows)
    If lErro <> AD_SQL_SUCESSO Then gError 214845
    
    'se posiciona no 1o registro do conjunto obtido anteriormente
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 214846
    
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Produtos_Le_Count = SUCESSO
       
    Exit Function

Erro_Produtos_Le_Count:

    Produtos_Le_Count = gErr

    Select Case gErr
        
        Case 214844
            Call Rotina_ErroECF(vbOKOnly, ERRO_ABERTURA_COMANDO, gErr)
            
        Case 214845, 214846
            Call Rotina_ErroECF(vbOKOnly, ERRO_LEITURA_PRODUTOS_0, gErr)
            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 214847)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function


Function Inicializa_Dados() As Long
    
Dim objAdmMeioPagtoCondPagto As ClassAdmMeioPagtoCondPagto
Dim objAdmMeioPagtoCondPagto1 As ClassAdmMeioPagtoCondPagto
Dim objAdmMeioPagto As ClassAdmMeioPagto
Dim iTicketAtivo As Integer
Dim iCartaoCreditoAtivo As Integer
Dim iCartaoDebitoAtivo As Integer
Dim iOutrosAtivo As Integer
Dim iAchou As Integer

Set gcolCartao = New Collection
Set gcolOutros = New Collection
Set gcolTicket = New Collection

    If giDebug = 1 Then MsgBox ("16")
    
    
    For Each objAdmMeioPagto In gcolAdmMeioPagto
    
        Select Case objAdmMeioPagto.iTipoMeioPagto
            
                    
            Case TIPOMEIOPAGTOLOJA_DINHEIRO
                Set objAdmMeioPagtoCondPagto = objAdmMeioPagto.colCondPagtoLoja(1)
                giDinheiroAtivo = objAdmMeioPagtoCondPagto.iAtivo
            
            Case TIPOMEIOPAGTOLOJA_CHEQUE
                Set objAdmMeioPagtoCondPagto = objAdmMeioPagto.colCondPagtoLoja(1)
                giChequeAtivo = objAdmMeioPagtoCondPagto.iAtivo
            
            Case TIPOMEIOPAGTOLOJA_TROCA
                Set objAdmMeioPagtoCondPagto = objAdmMeioPagto.colCondPagtoLoja(1)
                giTrocaAtivo = objAdmMeioPagtoCondPagto.iAtivo
            
            Case TIPOMEIOPAGTOLOJA_CARNE
                For Each objAdmMeioPagtoCondPagto In objAdmMeioPagto.colCondPagtoLoja
                    If objAdmMeioPagtoCondPagto.iAtivo = ADMMEIOPAGTOCONDPAGTO_ATIVO Then
                        giCarneAtivo = objAdmMeioPagtoCondPagto.iAtivo
                        Exit For
                    End If
                Next
            
            Case TIPOMEIOPAGTOLOJA_CARTAO_CREDITO, TIPOMEIOPAGTOLOJA_CARTAO_DEBITO
            
                For Each objAdmMeioPagtoCondPagto In objAdmMeioPagto.colCondPagtoLoja
        
                    If objAdmMeioPagtoCondPagto.iAtivo = ADMMEIOPAGTOCONDPAGTO_ATIVO Then
                        
                        If objAdmMeioPagto.iTipoMeioPagto = TIPOMEIOPAGTOLOJA_CARTAO_CREDITO Then giCartaoCreditoAtivo = ADMMEIOPAGTOCONDPAGTO_ATIVO
                        
                        If objAdmMeioPagto.iTipoMeioPagto = TIPOMEIOPAGTOLOJA_CARTAO_DEBITO Then giCartaoDebitoAtivo = ADMMEIOPAGTOCONDPAGTO_ATIVO
        
                    End If
        
                    If giPos = CAIXA_ACEITA_POS Then
        
                        Call Cria_AdmMeioPagtoCondPagto(objAdmMeioPagtoCondPagto, objAdmMeioPagtoCondPagto1)
                        objAdmMeioPagtoCondPagto1.iTipoCartao = TIPO_POS
                        
                        gcolCartao.Add objAdmMeioPagtoCondPagto1
                                                        
                    End If
                    
                    If giBoletoManual = CAIXA_ACEITA_BOLETO_MANUAL And objAdmMeioPagto.iTipoMeioPagto = TIPOMEIOPAGTOLOJA_CARTAO_CREDITO Then
        
                        Call Cria_AdmMeioPagtoCondPagto(objAdmMeioPagtoCondPagto, objAdmMeioPagtoCondPagto1)
                        objAdmMeioPagtoCondPagto1.iTipoCartao = TIPO_MANUAL
                                                        
                        gcolCartao.Add objAdmMeioPagtoCondPagto1
                                                        
                    End If
                    
                    If giTEF = CAIXA_ACEITA_TEF Then
        
                        Call Cria_AdmMeioPagtoCondPagto(objAdmMeioPagtoCondPagto, objAdmMeioPagtoCondPagto1)
                        objAdmMeioPagtoCondPagto1.iTipoCartao = TIPO_TEF
                                                        
                        gcolCartao.Add objAdmMeioPagtoCondPagto1
                                                        
                    End If
        
                Next
    
            Case TIPOMEIOPAGTOLOJA_VALE_TICKET, TIPOMEIOPAGTOLOJA_VALE_REFEICAO, TIPOMEIOPAGTOLOJA_VALE_PRESENTE, TIPOMEIOPAGTOLOJA_VALE_COMBUSTIVEL
            
                For Each objAdmMeioPagtoCondPagto In objAdmMeioPagto.colCondPagtoLoja
            
                    Call Cria_AdmMeioPagtoCondPagto(objAdmMeioPagtoCondPagto, objAdmMeioPagtoCondPagto1)
                    objAdmMeioPagtoCondPagto1.iTipoCartao = TIPO_MANUAL
                    gcolTicket.Add objAdmMeioPagtoCondPagto1
                    If objAdmMeioPagtoCondPagto.iAtivo = ADMMEIOPAGTOCONDPAGTO_ATIVO Then giTicketAtivo = ADMMEIOPAGTOCONDPAGTO_ATIVO
            
            
                    If giTEF = CAIXA_ACEITA_TEF Then
        
                        objAdmMeioPagtoCondPagto1.iTipoCartao = TIPO_TEF
                        gcolCartao.Add objAdmMeioPagtoCondPagto1
                                                        
                    End If
            
                    If giPos = CAIXA_ACEITA_POS Then
        
                        objAdmMeioPagtoCondPagto1.iTipoCartao = TIPO_POS
                        gcolCartao.Add objAdmMeioPagtoCondPagto1
                                                        
                    End If
            
            
            
                Next
            
            
            
            Case TIPOMEIOPAGTOLOJA_OUTROS
    
                For Each objAdmMeioPagtoCondPagto In objAdmMeioPagto.colCondPagtoLoja
            
                    Call Cria_AdmMeioPagtoCondPagto(objAdmMeioPagtoCondPagto, objAdmMeioPagtoCondPagto1)
                    gcolOutros.Add objAdmMeioPagtoCondPagto1
                    If objAdmMeioPagtoCondPagto.iAtivo = ADMMEIOPAGTOCONDPAGTO_ATIVO Then giOutrosAtivo = ADMMEIOPAGTOCONDPAGTO_ATIVO
            
                Next
    
        End Select
    
    Next
    
    If giDebug = 1 Then MsgBox ("17")
    
    Set objAdmMeioPagtoCondPagto = New ClassAdmMeioPagtoCondPagto
    objAdmMeioPagtoCondPagto.sNomeAdmMeioPagto = STRING_NAO_DETALHADO
    objAdmMeioPagtoCondPagto.iTipoCartao = TIPO_POS
    objAdmMeioPagtoCondPagto.iParcelamento = PARCELAMENTO_AVISTA
    objAdmMeioPagtoCondPagto.sNomeParcelamento = NOME_A_VISTA
    objAdmMeioPagtoCondPagto.iOrigem = OPTION_CF
    objAdmMeioPagtoCondPagto.iAtivo = giTicketAtivo
    gcolTicket.Add objAdmMeioPagtoCondPagto
    
    Set objAdmMeioPagtoCondPagto = New ClassAdmMeioPagtoCondPagto
    objAdmMeioPagtoCondPagto.sNomeAdmMeioPagto = STRING_NAO_DETALHADO_CCREDITO
    objAdmMeioPagtoCondPagto.iTipoCartao = TIPO_POS
    objAdmMeioPagtoCondPagto.iParcelamento = PARCELAMENTO_AVISTA
    objAdmMeioPagtoCondPagto.sNomeParcelamento = NOME_A_VISTA
    objAdmMeioPagtoCondPagto.iOrigem = OPTION_CF
    objAdmMeioPagtoCondPagto.iAtivo = giCartaoCreditoAtivo
    gcolCartao.Add objAdmMeioPagtoCondPagto
    
    Set objAdmMeioPagtoCondPagto = New ClassAdmMeioPagtoCondPagto
    objAdmMeioPagtoCondPagto.sNomeAdmMeioPagto = STRING_NAO_DETALHADO_CDEBITO
    objAdmMeioPagtoCondPagto.iTipoCartao = TIPO_POS
    objAdmMeioPagtoCondPagto.iParcelamento = PARCELAMENTO_AVISTA
    objAdmMeioPagtoCondPagto.sNomeParcelamento = NOME_A_VISTA
    objAdmMeioPagtoCondPagto.iOrigem = OPTION_CF
    objAdmMeioPagtoCondPagto.iAtivo = giCartaoDebitoAtivo
    gcolCartao.Add objAdmMeioPagtoCondPagto
    
    Set objAdmMeioPagtoCondPagto = New ClassAdmMeioPagtoCondPagto
    objAdmMeioPagtoCondPagto.sNomeAdmMeioPagto = STRING_NAO_DETALHADO
    objAdmMeioPagtoCondPagto.iTipoCartao = TIPO_POS
    objAdmMeioPagtoCondPagto.iParcelamento = PARCELAMENTO_AVISTA
    objAdmMeioPagtoCondPagto.sNomeParcelamento = NOME_A_VISTA
    objAdmMeioPagtoCondPagto.iOrigem = OPTION_CF
    objAdmMeioPagtoCondPagto.iAtivo = giOutrosAtivo
    gcolOutros.Add objAdmMeioPagtoCondPagto
    
    If giDebug = 1 Then MsgBox ("18")
    
    
    Inicializa_Dados = SUCESSO
    
End Function

Private Sub Cria_AdmMeioPagtoCondPagto(ByVal objAdmMeioPagtoCondPagto As ClassAdmMeioPagtoCondPagto, objAdmMeioPagtoCondPagto1 As ClassAdmMeioPagtoCondPagto)

    Set objAdmMeioPagtoCondPagto1 = New ClassAdmMeioPagtoCondPagto

    objAdmMeioPagtoCondPagto1.sNomeAdmMeioPagto = objAdmMeioPagtoCondPagto.sNomeAdmMeioPagto
    objAdmMeioPagtoCondPagto1.sNomeParcelamento = objAdmMeioPagtoCondPagto.sNomeParcelamento
    objAdmMeioPagtoCondPagto1.iAdmMeioPagto = objAdmMeioPagtoCondPagto.iAdmMeioPagto
    objAdmMeioPagtoCondPagto1.iParcelamento = objAdmMeioPagtoCondPagto.iParcelamento
    objAdmMeioPagtoCondPagto1.iOrigem = OPTION_CF
    objAdmMeioPagtoCondPagto1.iAtivo = objAdmMeioPagtoCondPagto.iAtivo

End Sub

Function Verifica_Arquivo_Loja(dtUltimoMovto As Date) As Long
'faz a leitura do ultimo arquivo, mesmo que o dia tenha mudado para cuidar do caixa (suprimento) que ainda sobrou

Dim lErro As Long
Dim sTipo As String
Dim iPos As Integer
Dim iPosEsc As Integer
Dim iIndice As Integer
Dim iIndice1 As Integer
Dim sNomeArq As String
Dim iPosInicio As Integer
Dim bAchou As Boolean
Dim bAchou1 As Boolean
Dim iPosInicio1 As Integer
Dim bAchouAbert As Boolean
Dim colRegistroTemp As New Collection
Dim sSeqAbert As String
Dim sSeqFech As String
Dim sNum As String
Dim sRegistro As String
Dim colRegistro As New Collection
Dim iRegistro As Integer
Dim iNum As Integer
Dim iSeqArq As Integer
Dim lTamanho As Long
Dim sRetorno As String
Dim sNomeArq1 As String
Dim sNomeArq2 As String
Dim sRet As String
Dim iCodEmpresa As Integer
Dim sArquivo As String
Dim colVendas As New Collection
Dim objVenda As New ClassVenda
Dim objVendedor As ClassVendedor
Dim objProduto As ClassProduto
Dim objItens As ClassItemCupomFiscal
Dim sNomeArq0 As String
Dim iStatus As Integer
Dim bNaoPulaCodigo As Boolean
Dim vbMsg As VbMsgBoxResult
Dim dtData As Date
Dim dtData1 As Date
Dim alComando(1 To 3) As Long
Dim lNumMovto As Long
Dim iTipo As Integer
Dim dHora As Double
Dim iSeq As Integer
Dim sMsg As String
Dim lErro1 As Long
Dim objMovCaixa As ClassMovimentoCaixa
Dim iCodModeloECF As Integer
Dim objAliquota As ClassAliquotaICMS

On Error GoTo Erro_Verifica_Arquivo_Loja
    
    '??? 24/08/2016 gdSaldoDinheiro = 0
    gdSaldocheques = 0
    Set gcolCarne = New Collection
    
    For Each objAliquota In gcolAliquotasTotal
        objAliquota.dValorTotalizadoLoja = 0
    Next
    
    dtData1 = DATA_NULA
    bNaoPulaCodigo = True
    bAchouAbert = True
    bAchou = False
      
    iPosInicio = 1
    
    If giDebug = 1 Then MsgBox ("46")
    
   'Abre os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_AbrirExt(glConexaoPAFECF)
        If alComando(iIndice) = 0 Then gError 210425
    Next
                    
    sRegistro = String(STRING_MOVIMENTOCAIXA_MSG, 0)
    
    lErro = Comando_Executar(alComando(1), "SELECT Data  FROM MovimentoCaixa  ORDER BY Data DESC ", dtData1)
    If lErro <> AD_SQL_SUCESSO Then gError 210426

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 210427
        
    If lErro = AD_SQL_SUCESSO Then
    
        sMsg = String(STRING_MOVIMENTOCAIXA_MSG, 0)
    
        lErro = Comando_Executar(alComando(2), "SELECT NumMovto, Data, Tipo, Msg, Hora, Seq, CodModeloECF FROM MovimentoCaixa  WHERE Data = ? ORDER BY NumMovto, Seq  ", lNumMovto, dtData, iTipo, sMsg, dHora, iSeq, iCodModeloECF, dtData1)
        If lErro <> AD_SQL_SUCESSO Then gError 210428
    
        lErro1 = Comando_BuscarPrimeiro(alComando(2))
        If lErro1 <> AD_SQL_SUCESSO And lErro1 <> AD_SQL_SEM_DADOS Then gError 210429
        
        Do While lErro1 = AD_SQL_SUCESSO
                    
            colRegistro.Add sMsg
                    
            lErro1 = Comando_BuscarProximo(alComando(2))
            If lErro1 <> AD_SQL_SUCESSO And lErro1 <> AD_SQL_SEM_DADOS Then gError 210430
            
            If iSeq = 1 Or lErro1 = AD_SQL_SEM_DADOS Then
                    
                lErro = CF_ECF("Registro_ECF_CC", colRegistro, sArquivo, iRegistro, iIndice, colVendas, iCodModeloECF)
                If lErro <> SUCESSO Then gError 110002
                    
                Set colRegistro = New Collection
                    
            End If
            
        Loop
        
        
        
        If giDebug = 1 Then MsgBox ("48")
        
        'Primeira Posição
        iPosInicio = 1
        
        
        For Each objVenda In colVendas
        
            bAchou = False
            'se for um orçamneto
            If objVenda.objCupomFiscal.iTipo = OPTION_DAV Or objVenda.objCupomFiscal.iTipo = OPTION_PREVENDA Then
                'verifica se o vendedor ainda existe-> se naão existir não impacta na memória
                If objVenda.objCupomFiscal.iVendedor <> 0 Then

                    Set objVendedor = New ClassVendedor

                    lErro = CF_ECF("Vendedores_Le_Codigo", objVenda.objCupomFiscal.iVendedor, objVendedor)
                    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 214944
                    
                    If lErro = SUCESSO Then bAchou = True
                    
                Else
                
                    bAchou = True
                End If
                    
                If bAchou Then
                    'verifica se os produtos ainda existe-> se naão existir não impacta na memória
                    
                    For Each objItens In objVenda.objCupomFiscal.colItens
                        bAchou = False
                        
                        Set objProduto = New ClassProduto
                        
                        lErro = CF_ECF("Produtos_Le", objItens.sProduto, objProduto)
                        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 214854
                        
                        If lErro = SUCESSO Then
                            objItens.sProdutoNomeRed = objProduto.sNomeReduzido
                            bAchou = True
                        Else
                            Exit For
                            
                        End If
                        
                        
                    Next
                    
                End If
            Else
                bAchou = True
            End If
            
            If bAchou Then
            
                For Each objMovCaixa In objVenda.colMovimentosCaixa
                
                    If objMovCaixa.iTipo = MOVIMENTOCAIXA_TROCO_DINHEIRO Then
                        objMovCaixa.dValor = -objMovCaixa.dValor
                    End If
                    
                Next
            
                If objVenda.objCupomFiscal.lNumero <> 0 Then
                    'Atualiza os Movimentos nas coleções globais
                    Call CF_ECF("Atualiza_Movimentos_Memoria", objVenda)
                    Call Atualiza_Movimentos(objVenda)
                ElseIf objVenda.objCupomFiscal.iStatus = STATUS_BAIXADO Then
                    'Atualiza os Movimentos nas coleções globais
                    Call CF_ECF("Atualiza_Movimentos_Memoria1", objVenda)
                    Call Atualiza_Movimentos(objVenda)
                End If
                
                'Atribui para a coleção global o objvenda
                gcolVendas.Add objVenda
            End If
        Next
    
    End If
    
    If giDebug = 1 Then MsgBox ("49")
    
    'mario
    'se foi lido um arquivo de outra data ==> limpa as colecoes
    If dtData1 <> Date Then
        Set gcolMovimentosCaixa = New Collection
        Set gcolVendas = New Collection
    End If
    
    dtUltimoMovto = dtData1

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    If giDebug = 1 Then MsgBox ("50")
    
    
    Verifica_Arquivo_Loja = SUCESSO
    
    Exit Function
    
Erro_Verifica_Arquivo_Loja:
    
    Verifica_Arquivo_Loja = gErr
    
    Close 1
    
    Select Case gErr
            
        Case 53
            Call Rotina_ErroECF(vbOKOnly, ERRO_ARQUIVO_NAO_ENCONTRADO, gErr, sNomeArq0)

        Case 58
            Call Rotina_ErroECF(vbOKOnly, ERRO_ARQUIVO_EXISTENTE, gErr, sNomeArq0)
            
        Case 105693, 110002, 214854, 214944
        
        Case 210425
            Call Rotina_ErroECF(vbOKOnly, ERRO_ABERTURA_COMANDO, gErr)
        
        Case 210426 To 210430
            Call Rotina_ErroECF(vbOKOnly, ERRO_LEITURA_MOVIMENTOCAIXA, gErr)
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error$, 165191)

    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
End Function

Private Sub Atualiza_Movimentos(objVenda As ClassVenda)
    
Dim objMovCaixa As ClassMovimentoCaixa
Dim objCheque As ClassChequePre
Dim lErro As Long

On Error GoTo Erro_Atualiza_Movimentos

   'Jogo todos os cheques na col global
    For Each objCheque In objVenda.colCheques
        'Atualiza o saldos de cheques
        gdSaldocheques = gdSaldocheques + objCheque.dValor
    Next
        
    'Para cada movimento da venda
    For Each objMovCaixa In objVenda.colMovimentosCaixa
        'Se for de cartao de crédito ou débito especificado
        '??? 24/08/2016 If objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_DINHEIRO Then gdSaldoDinheiro = gdSaldoDinheiro + objMovCaixa.dValor
    Next
    
    Exit Sub
    
Erro_Atualiza_Movimentos:

    Select Case gErr
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error, 165207)

    End Select
        
    Exit Sub
    
End Sub

Function Clientes_Le_NomeReduzido(ByVal sNomeReduzido As String, ByVal objCliente As ClassCliente) As Long

Dim alComando(1 To 1) As Long
Dim lErro As Long
Dim sCgc As String
Dim lCodigo As Long
Dim iIndice As Integer


On Error GoTo Erro_Clientes_Le_NomeReduzido

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_AbrirExt(glConexaoPAFECF)
        If alComando(iIndice) = 0 Then gError 214875
    Next

    sCgc = String(STRING_CLIENTE_CGC, 0)

    'Le o NomeReduzido da tabela Clientes
    lErro = Comando_Executar(alComando(1), "SELECT Codigo, CGC FROM Clientes WHERE NomeReduzido = ?", lCodigo, sCgc, sNomeReduzido)
    If lErro <> AD_SQL_SUCESSO Then gError 214876

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 214877

    If lErro = AD_SQL_SEM_DADOS Then gError 214878

    objCliente.sNomeReduzido = sNomeReduzido
    objCliente.lCodigo = lCodigo
    objCliente.sCgc = sCgc

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Clientes_Le_NomeReduzido = SUCESSO

Exit Function

Erro_Clientes_Le_NomeReduzido:

    Clientes_Le_NomeReduzido = gErr

    Select Case gErr
        
        Case 214875
            Call Rotina_ErroECF(vbOKOnly, ERRO_ABERTURA_COMANDO, gErr)
            
        Case 214876, 214877
            Call Rotina_ErroECF(vbOKOnly, ERRO_LEITURA_CLIENTE_NOMEREDUZIDO, gErr, sNomeReduzido)
            
        Case 214878
            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 214879)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function Clientes_Le_Count(lRows As Long) As Long

Dim lErro As Long
Dim sCodBarras As String
Dim tProduto As typeProduto
Dim alComando(1 To 1) As Long
Dim iIndice As Integer
Dim objProduto As ClassProduto

On Error GoTo Erro_Clientes_Le_Count

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_AbrirExt(glConexaoPAFECF)
        If alComando(iIndice) = 0 Then gError 214880
    Next

        
    'obtem o conjunto de codigos de barra relativos ao produto em questao
    lErro = Comando_Executar(alComando(1), "SELECT Count(*) FROM Clientes", lRows)
    If lErro <> AD_SQL_SUCESSO Then gError 214881
    
    'se posiciona no 1o registro do conjunto obtido anteriormente
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 214882
    
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Clientes_Le_Count = SUCESSO
       
    Exit Function

Erro_Clientes_Le_Count:

    Clientes_Le_Count = gErr

    Select Case gErr
        
        Case 214880
            Call Rotina_ErroECF(vbOKOnly, ERRO_ABERTURA_COMANDO, gErr)
            
        Case 214881, 214882
            Call Rotina_ErroECF(vbOKOnly, ERRO_LEITURA_CLIENTES_0, gErr)
            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 214883)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Function Clientes_Le_CPFCNPJ(ByVal sCPFCNPJ As String, ByVal objCliente As ClassCliente) As Long

Dim alComando(1 To 1) As Long
Dim lErro As Long
Dim sNomeReduzido As String
Dim lCodigo As Long
Dim iIndice As Integer


On Error GoTo Erro_Clientes_Le_CPFCNPJ

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_AbrirExt(glConexaoPAFECF)
        If alComando(iIndice) = 0 Then gError 214884
    Next

    sNomeReduzido = String(STRING_CLIENTE_NOME_REDUZIDO, 0)

    'Le o NomeReduzido da tabela Clientes
    lErro = Comando_Executar(alComando(1), "SELECT Codigo, NomeReduzido FROM Clientes WHERE CGC = ?", lCodigo, sNomeReduzido, sCPFCNPJ)
    If lErro <> AD_SQL_SUCESSO Then gError 214885

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 214886

    If lErro = AD_SQL_SEM_DADOS Then gError 214887

    objCliente.sNomeReduzido = sNomeReduzido
    objCliente.lCodigo = lCodigo
    objCliente.sCgc = sCPFCNPJ

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Clientes_Le_CPFCNPJ = SUCESSO

Exit Function

Erro_Clientes_Le_CPFCNPJ:

    Clientes_Le_CPFCNPJ = gErr

    Select Case gErr
        
        Case 214884
            Call Rotina_ErroECF(vbOKOnly, ERRO_ABERTURA_COMANDO, gErr)
            
        Case 214885, 214886
            Call Rotina_ErroECF(vbOKOnly, ERRO_LEITURA_CLIENTE_NOMEREDUZIDO, gErr, sNomeReduzido)
            
        Case 214887
            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 214888)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function Clientes_Le_NomeReduzido1(ByVal colClientes As Collection) As Long

Dim alComando(1 To 1) As Long
Dim lErro As Long
Dim sCgc As String
Dim lCodigo As Long
Dim iIndice As Integer
Dim sNomeReduzido As String
Dim objCliente As ClassCliente

On Error GoTo Erro_Clientes_Le_NomeReduzido1

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_AbrirExt(glConexaoPAFECF)
        If alComando(iIndice) = 0 Then gError 214889
    Next

    sCgc = String(STRING_CLIENTE_CGC, 0)
    sNomeReduzido = String(STRING_CLIENTE_NOME_REDUZIDO, 0)

    'Le o NomeReduzido da tabela Clientes
    lErro = Comando_Executar(alComando(1), "SELECT Codigo, CGC, NomeReduzido FROM Clientes ORDER BY NomeReduzido", lCodigo, sCgc, sNomeReduzido)
    If lErro <> AD_SQL_SUCESSO Then gError 214890

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 214891

    Do While lErro = AD_SQL_SUCESSO

        Set objCliente = New ClassCliente

        objCliente.sNomeReduzido = sNomeReduzido
        objCliente.lCodigo = lCodigo
        objCliente.sCgc = sCgc
        
        colClientes.Add objCliente

        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 214892

    Loop

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Clientes_Le_NomeReduzido1 = SUCESSO

Exit Function

Erro_Clientes_Le_NomeReduzido1:

    Clientes_Le_NomeReduzido1 = gErr

    Select Case gErr
        
        Case 214889
            Call Rotina_ErroECF(vbOKOnly, ERRO_ABERTURA_COMANDO, gErr)
            
        Case 214890, 214891, 214892
            Call Rotina_ErroECF(vbOKOnly, ERRO_LEITURA_CLIENTES_0, gErr)
            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 214893)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function Clientes_Le_Codigo(ByVal lCodigo As Long, ByVal objCliente As ClassCliente) As Long

Dim alComando(1 To 1) As Long
Dim lErro As Long
Dim sNomeReduzido As String
Dim iIndice As Integer
Dim sCgc As String

On Error GoTo Erro_Clientes_Le_Codigo

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_AbrirExt(glConexaoPAFECF)
        If alComando(iIndice) = 0 Then gError 214896
    Next

    sNomeReduzido = String(STRING_CLIENTE_NOME_REDUZIDO, 0)
    sCgc = String(STRING_CLIENTE_CGC, 0)

    'Le o NomeReduzido da tabela Clientes
    lErro = Comando_Executar(alComando(1), "SELECT CGC, NomeReduzido FROM Clientes WHERE Codigo = ?", sCgc, sNomeReduzido, lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 214897

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 214898

    If lErro = AD_SQL_SEM_DADOS Then gError 214899

    objCliente.sNomeReduzido = sNomeReduzido
    objCliente.lCodigo = lCodigo
    objCliente.sCgc = sCgc

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Clientes_Le_Codigo = SUCESSO

Exit Function

Erro_Clientes_Le_Codigo:

    Clientes_Le_Codigo = gErr

    Select Case gErr
        
        Case 214896
            Call Rotina_ErroECF(vbOKOnly, ERRO_ABERTURA_COMANDO, gErr)
            
        Case 214897, 214898
            Call Rotina_ErroECF(vbOKOnly, ERRO_LEITURA_CLIENTE_CODIGO, gErr, lCodigo)
            
        Case 214899
            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 214900)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function Vendedores_Le_Codigo(ByVal iCodigo As Integer, ByVal objVendedor As ClassVendedor) As Long

Dim alComando(1 To 1) As Long
Dim lErro As Long
Dim sNomeReduzido As String
Dim iIndice As Integer
Dim sCgc As String

On Error GoTo Erro_Vendedores_Le_Codigo

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_AbrirExt(glConexaoPAFECF)
        If alComando(iIndice) = 0 Then gError 214940
    Next

    sNomeReduzido = String(STRING_VENDEDOR_NOME_REDUZIDO, 0)

    'Le o NomeReduzido da tabela Vendedores
    lErro = Comando_Executar(alComando(1), "SELECT NomeReduzido FROM Vendedores WHERE Codigo = ?", sNomeReduzido, iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 214941

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 214942

    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objVendedor.sNomeReduzido = sNomeReduzido

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Vendedores_Le_Codigo = SUCESSO

Exit Function

Erro_Vendedores_Le_Codigo:

    Vendedores_Le_Codigo = gErr

    Select Case gErr
        
        Case 214940
            Call Rotina_ErroECF(vbOKOnly, ERRO_ABERTURA_COMANDO, gErr)
            
        Case 214941, 214942
            Call Rotina_ErroECF(vbOKOnly, ERRO_LEITURA_VENDEDOR_CODIGO, gErr, iCodigo)
            
        Case ERRO_LEITURA_SEM_DADOS
            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 214943)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Private Function SAT_Imprime_Daruma(ByVal objVenda As ClassVenda) As Long
'http://www.desenvolvedoresdaruma.com.br/home/downloads/Site_2011/Help/DarumaFrameworkHelpOnline/DarumaFramework/Mini_Impressora/Trabalhando_com_C%C3%B3digo_de_Barras_-_Impressora_N%C3%A3o-Fiscal.htm

Dim lErro As Long
Dim objItem As ClassItemCupomFiscal
Dim objProduto As ClassProduto
Dim iIndice As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim bCont As Boolean
Dim sMensagem As String
Dim dTotal As Double
Dim dPreco As Double
Dim bAchou As Boolean
Dim colOrcamento As New Collection
Dim iIndice1 As Integer
Dim objCliente As ClassCliente
Dim lIndice As Long
Dim sCliente As String
Dim sCPFCGC As String
Dim sVendedor As String
Dim objVendedor As ClassVendedor
Dim sProduto As String
Dim sCNPJ As String
Dim iPos As Integer
Dim lNumDAV As Long, objMovCaixa As ClassMovimentoCaixa
Dim sRetorno As String, sSATChaveAcesso As String
Dim objProdutoNomeRed As TextBox, sSATtimeStamp As String
Dim sTexto As String, sTexto2 As String, iTam As Integer
Dim objDarumaNaoFiscalImp As New ClassDarumaNaoFiscalImp

On Error GoTo Erro_SAT_Imprime_Daruma

    Set objCliente = gobjClienteCPF.Busca(objVenda.objCupomFiscal.sCPFCGC, lIndice)
    If Not objCliente Is Nothing Then sCliente = objCliente.sNomeReduzido & " - "
    Call Formata_CNPJ_CPF(sCPFCGC, objVenda.objCupomFiscal.sCPFCGC)
    sCliente = sCliente & sCPFCGC
    
    Call Formata_CNPJ_CPF(sCNPJ, gsCNPJ)
    
    Call objDarumaNaoFiscalImp.ImprimeNormal(Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", gsNomeReduzido))
    Call objDarumaNaoFiscalImp.ImprimeNormal(Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", gsNomeEmpresa))
    Call objDarumaNaoFiscalImp.ImprimeNormal(Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", gsEndereco))
    Call objDarumaNaoFiscalImp.ImprimeNormal("<l>")
    Call objDarumaNaoFiscalImp.ImprimeNormal(Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "CNPJ:" & sCNPJ))
    Call objDarumaNaoFiscalImp.ImprimeNormal(Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "IE:" & gsInscricaoEstadual))
    Call objDarumaNaoFiscalImp.ImprimeNormal(Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "IM:" & gsInscricaoMunicipal))
    Call objDarumaNaoFiscalImp.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", ""))
    
    If objVenda.objCupomFiscal.iSATtpAmb <> 1 Then
    
        Call objDarumaNaoFiscalImp.ImprimeNormal("<b>" & Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "Extrato No. 000000") & "</b>")
        Call objDarumaNaoFiscalImp.ImprimeNormal("<b>" & Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "CUPOM FISCAL ELETRÔNICO - SAT") & "</b>")
        Call objDarumaNaoFiscalImp.ImprimeNormal("<l>")
        Call objDarumaNaoFiscalImp.ImprimeNormal(Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "= T E S T E ="))
        Call objDarumaNaoFiscalImp.ImprimeNormal("<l>")
        Call objDarumaNaoFiscalImp.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, ">", ""))
        Call objDarumaNaoFiscalImp.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, ">", ""))
        Call objDarumaNaoFiscalImp.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, ">", ""))
    
    Else
    
        Call objDarumaNaoFiscalImp.ImprimeNormal("<b>" & Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "Extrato No. " & Format(objVenda.objCupomFiscal.lNumero, "000000")) & "</b>")
        Call objDarumaNaoFiscalImp.ImprimeNormal("<b>" & Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "CUPOM FISCAL ELETRÔNICO - SAT") & "</b>")
    
    End If
    
    Call objDarumaNaoFiscalImp.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", ""))
    Call objDarumaNaoFiscalImp.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "CPF/CNPJ do Consumidor: " & sCPFCGC))
    Call objDarumaNaoFiscalImp.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", ""))
    Call objDarumaNaoFiscalImp.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "#|COD|DESC|QTD|UN|VL UN R$|VL ITEM R$"))
    Call objDarumaNaoFiscalImp.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", ""))
    
    iIndice = 0
    
    For Each objItem In objVenda.objCupomFiscal.colItens
    
        If objItem.iStatus = STATUS_ATIVO Then
        
            iIndice = iIndice + 1
            
            sTexto = Format(iIndice, "000") & " " & objItem.sProduto & " " & objItem.sProdutoNomeRed & " " & IIf(right(Format(objItem.dQuantidade, "0.0###"), 2) = ",0", Format(objItem.dQuantidade, "###0"), Format(objItem.dQuantidade, "0.0###")) & " " & objItem.sUnidadeMed & " X " & Format(objItem.dPrecoUnitario, "#,###,##0.00")
            sTexto2 = Format(Arredonda_Moeda(objItem.dQuantidade * objItem.dPrecoUnitario, 4), "#,###,##0.00")
            iTam = Len(sTexto) + Len(sTexto2)
            
            If iTam < 48 Then
                Call objDarumaNaoFiscalImp.ImprimeNormal(sTexto & String(48 - iTam, " ") & sTexto2)
            Else
                Call objDarumaNaoFiscalImp.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", sTexto))
                Call objDarumaNaoFiscalImp.ImprimeNormal(Formata_Campo(ALINHAMENTO_ESQUERDA, 48, " ", sTexto2))
            End If
            
            dPreco = Arredonda_Moeda(objItem.dPrecoUnitario * objItem.dQuantidade)
            
            'se existir desconto sobre o item...
            If objItem.dPercDesc > 0 Then
                sTexto = Formata_Campo(ALINHAMENTO_DIREITA, 15, " ", "DESC. ITEM:") & Formata_Campo(ALINHAMENTO_DIREITA, 10, " ", Format(objItem.dPercDesc, "standard") & "%") & " = " & Formata_Campo(ALINHAMENTO_ESQUERDA, 15, " ", Format(objItem.dPrecoUnitario * objItem.dQuantidade * (1 - (objItem.dPercDesc / 100)), "standard"))
                Call objDarumaNaoFiscalImp.ImprimeNormal(sTexto)
                dPreco = dPreco - (objItem.dPrecoUnitario * objItem.dQuantidade * objItem.dPercDesc / 100)
            ElseIf objItem.dValorDesconto > 0 Then
                sTexto = Formata_Campo(ALINHAMENTO_DIREITA, 15, " ", "DESC. ITEM:") & Formata_Campo(ALINHAMENTO_DIREITA, 10, " ", Format(-objItem.dValorDesconto, "standard")) & " = " & Formata_Campo(ALINHAMENTO_ESQUERDA, 15, " ", Format((objItem.dPrecoUnitario - objItem.dValorDesconto) * objItem.dQuantidade, "standard"))
                Call objDarumaNaoFiscalImp.ImprimeNormal(sTexto)
                dPreco = dPreco - objItem.dValorDesconto
            End If
            
            dTotal = dTotal + dPreco
    
        End If
    
    Next
                      
    If (objVenda.objCupomFiscal.dValorDesconto + objVenda.objCupomFiscal.dValorDesconto1) <> 0 Or objVenda.objCupomFiscal.dValorAcrescimo <> 0 Then
        Call objDarumaNaoFiscalImp.ImprimeNormal("<c>" & Formata_Campo(ALINHAMENTO_ESQUERDA, 48, " ", "Subtotal :   " & Format(dTotal, "standard")) & "</c>")
    
        If (objVenda.objCupomFiscal.dValorDesconto + objVenda.objCupomFiscal.dValorDesconto1) <> 0 Then
            Call objDarumaNaoFiscalImp.ImprimeNormal("<c>" & Formata_Campo(ALINHAMENTO_ESQUERDA, 48, " ", "Descontos  :  -" & Format(objVenda.objCupomFiscal.dValorDesconto + objVenda.objCupomFiscal.dValorDesconto1, "standard")) & "</c>")
            dTotal = dTotal - (objVenda.objCupomFiscal.dValorDesconto + objVenda.objCupomFiscal.dValorDesconto1)
        End If
    
        If objVenda.objCupomFiscal.dValorAcrescimo <> 0 Then
            Call objDarumaNaoFiscalImp.ImprimeNormal("<c>" & Formata_Campo(ALINHAMENTO_ESQUERDA, 48, " ", "Acréscimos :  +" & Format(objVenda.objCupomFiscal.dValorAcrescimo, "standard")) & "</c>")
            dTotal = dTotal + objVenda.objCupomFiscal.dValorAcrescimo
        End If

    End If

    Call objDarumaNaoFiscalImp.ImprimeNormal("<e>" & Formata_Campo(ALINHAMENTO_ESQUERDA, 48, " ", "TOTAL R$:   " & Format(dTotal, "standard")) & "</e>")
    Call objDarumaNaoFiscalImp.ImprimeNormal("<l>")
    For Each objMovCaixa In objVenda.colMovimentosCaixa
        If objMovCaixa.iTipo <> MOVIMENTOCAIXA_TROCO_DINHEIRO And objMovCaixa.iTipo <> MOVIMENTOCAIXA_TROCO_CONTRAVALE And objMovCaixa.iTipo <> MOVIMENTOCAIXA_TROCO_VALE Then
            Call objDarumaNaoFiscalImp.ImprimeNormal("<c>" & Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", MeioPagtoCfe_Descricao(objMovCaixa.iCodigoCFe) & ":" & Format(objMovCaixa.dValor, "standard")) & "</c>")
        Else
            Call objDarumaNaoFiscalImp.ImprimeNormal("<c>" & Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "Troco R$:" & Format(objMovCaixa.dValor, "standard")) & "</c>")
        End If
    Next

    Call objDarumaNaoFiscalImp.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", ""))
    Call objDarumaNaoFiscalImp.ImprimeNormal("<b>" & Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "SAT No. " & objVenda.objCupomFiscal.sNumSerieECF) & "</b>")
    
    sSATtimeStamp = objVenda.objCupomFiscal.sSATtimeStamp
    Call objDarumaNaoFiscalImp.ImprimeNormal(Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", Mid(sSATtimeStamp, 7, 2) & "/" & Mid(sSATtimeStamp, 5, 2) & "/" & left(sSATtimeStamp, 4) & " - " & Mid(sSATtimeStamp, 9, 2) & ":" & Mid(sSATtimeStamp, 11, 2) & ":" & Mid(sSATtimeStamp, 13, 2)))
    
    sSATChaveAcesso = Mid(objVenda.objCupomFiscal.sSATChaveAcesso, 4)
    
    Call objDarumaNaoFiscalImp.ImprimeNormal(Mid(sSATChaveAcesso, 1, 11) & " " & Mid(sSATChaveAcesso, 12, 11) & " " & Mid(sSATChaveAcesso, 23, 11) & " " & Mid(sSATChaveAcesso, 34, 11))
    
    Call objDarumaNaoFiscalImp.ImprimeNormal("<ce><code128>" & left(sSATChaveAcesso, 22) & "</code128></ce>")
    Call objDarumaNaoFiscalImp.ImprimeNormal("<ce><code128>" & right(sSATChaveAcesso, 22) & "</code128></ce>")
    
    Call objDarumaNaoFiscalImp.ImprimeNormal("<ce><qrcode>" & sSATChaveAcesso & "|" & sSATtimeStamp & "|" & objVenda.objCupomFiscal.sSATvalorTotalCFe & "|" & objVenda.objCupomFiscal.sSATCPFCNPJValue & "|" & objVenda.objCupomFiscal.sSATQRCode & "<lmodulo>4</lmodulo></qrcode></ce>")
    
    Call objDarumaNaoFiscalImp.ImprimeNormal("<sl>4</sl><confgui>P</confgui><gui></gui>")
    
    Call objDarumaNaoFiscalImp.TerminaImpressao
        
    SAT_Imprime_Daruma = SUCESSO
    
    Exit Function
    
Erro_SAT_Imprime_Daruma:

    SAT_Imprime_Daruma = gErr

    Select Case gErr

        Case 99803, 99894 To 99898, 105868, 133088, 204319, 204324, 204329, 210470

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149439)

    End Select

    Exit Function

End Function

Private Function SAT_Imprime_Canc_Daruma(ByVal objVenda As ClassVenda) As Long
'http://www.desenvolvedoresdaruma.com.br/home/downloads/Site_2011/Help/DarumaFrameworkHelpOnline/DarumaFramework/Mini_Impressora/Trabalhando_com_C%C3%B3digo_de_Barras_-_Impressora_N%C3%A3o-Fiscal.htm

Dim lErro As Long
Dim objItem As ClassItemCupomFiscal
Dim objProduto As ClassProduto
Dim iIndice As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim bCont As Boolean
Dim sMensagem As String
Dim dTotal As Double
Dim dPreco As Double
Dim bAchou As Boolean
Dim colOrcamento As New Collection
Dim iIndice1 As Integer
Dim objCliente As ClassCliente
Dim lIndice As Long
Dim sCliente As String
Dim sCPFCGC As String
Dim sVendedor As String
Dim objVendedor As ClassVendedor
Dim sProduto As String
Dim sCNPJ As String
Dim iPos As Integer
Dim lNumDAV As Long, objMovCaixa As ClassMovimentoCaixa
Dim sRetorno As String, sSATChaveAcesso As String
Dim objProdutoNomeRed As TextBox, sSATtimeStamp As String
Dim sTexto As String, sTexto2 As String, iTam As Integer
Dim objDarumaNaoFiscalImp As New ClassDarumaNaoFiscalImp

On Error GoTo Erro_SAT_Imprime_Canc_Daruma

    Set objCliente = gobjClienteCPF.Busca(objVenda.objCupomFiscal.sCPFCGC, lIndice)
    If Not objCliente Is Nothing Then sCliente = objCliente.sNomeReduzido & " - "
    Call Formata_CNPJ_CPF(sCPFCGC, objVenda.objCupomFiscal.sCPFCGC)
    sCliente = sCliente & sCPFCGC
    
    Call Formata_CNPJ_CPF(sCNPJ, gsCNPJ)
    
    Call objDarumaNaoFiscalImp.ImprimeNormal(Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", gsNomeReduzido))
    Call objDarumaNaoFiscalImp.ImprimeNormal(Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", gsNomeEmpresa))
    Call objDarumaNaoFiscalImp.ImprimeNormal(Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", gsEndereco))
    Call objDarumaNaoFiscalImp.ImprimeNormal("<l>")
    Call objDarumaNaoFiscalImp.ImprimeNormal(Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "CNPJ:" & sCNPJ))
    Call objDarumaNaoFiscalImp.ImprimeNormal(Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "IE:" & gsInscricaoEstadual))
    Call objDarumaNaoFiscalImp.ImprimeNormal(Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "IM:" & gsInscricaoMunicipal))
    Call objDarumaNaoFiscalImp.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", ""))
    
    If objVenda.objCupomFiscal.iSATtpAmb <> 1 Then
    
        Call objDarumaNaoFiscalImp.ImprimeNormal("<b>" & Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "Extrato No. 000000") & "</b>")
        Call objDarumaNaoFiscalImp.ImprimeNormal("<b>" & Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "CUPOM FISCAL ELETRÔNICO - SAT") & "</b>")
        Call objDarumaNaoFiscalImp.ImprimeNormal("<b>" & Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "CANCELAMENTO") & "</b>")
        Call objDarumaNaoFiscalImp.ImprimeNormal("<l>")
        Call objDarumaNaoFiscalImp.ImprimeNormal(Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "= T E S T E ="))
        Call objDarumaNaoFiscalImp.ImprimeNormal("<l>")
        Call objDarumaNaoFiscalImp.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, ">", ""))
        Call objDarumaNaoFiscalImp.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, ">", ""))
        Call objDarumaNaoFiscalImp.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, ">", ""))
    
    Else
    
        Call objDarumaNaoFiscalImp.ImprimeNormal("<b>" & Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "Extrato No. " & Format(objVenda.objCupomFiscal.lNumero, "000000")) & "</b>")
        Call objDarumaNaoFiscalImp.ImprimeNormal("<b>" & Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "CUPOM FISCAL ELETRÔNICO - SAT") & "</b>")
        Call objDarumaNaoFiscalImp.ImprimeNormal("<b>" & Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "CANCELAMENTO") & "</b>")
    
    End If
    
    Call objDarumaNaoFiscalImp.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", ""))
    Call objDarumaNaoFiscalImp.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "CPF/CNPJ do Consumidor: " & sCPFCGC))
    
    Call objDarumaNaoFiscalImp.ImprimeNormal("<e>" & Formata_Campo(ALINHAMENTO_ESQUERDA, 48, " ", "TOTAL R$:   " & Format(objVenda.objCupomFiscal.dValorTotal, "standard")) & "</e>")

    Call objDarumaNaoFiscalImp.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", ""))
    Call objDarumaNaoFiscalImp.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "DADOS DO CUPOM FISCAL ELETRÔNICO CANCELADO"))
    Call objDarumaNaoFiscalImp.ImprimeNormal("<b>" & Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "SAT No. " & objVenda.objCupomFiscal.sNumSerieECF) & "</b>")
    
    sSATtimeStamp = objVenda.objCupomFiscal.sSATtimeStamp
    Call objDarumaNaoFiscalImp.ImprimeNormal(Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", Mid(sSATtimeStamp, 7, 2) & "/" & Mid(sSATtimeStamp, 5, 2) & "/" & left(sSATtimeStamp, 4) & " - " & Mid(sSATtimeStamp, 9, 2) & ":" & Mid(sSATtimeStamp, 11, 2) & ":" & Mid(sSATtimeStamp, 13, 2)))
    
    sSATChaveAcesso = Mid(objVenda.objCupomFiscal.sSATChaveAcesso, 4)
    
    Call objDarumaNaoFiscalImp.ImprimeNormal(Mid(sSATChaveAcesso, 1, 11) & " " & Mid(sSATChaveAcesso, 12, 11) & " " & Mid(sSATChaveAcesso, 23, 11) & " " & Mid(sSATChaveAcesso, 34, 11))
    
    Call objDarumaNaoFiscalImp.ImprimeNormal("<ce><code128>" & left(sSATChaveAcesso, 22) & "</code128></ce>")
    Call objDarumaNaoFiscalImp.ImprimeNormal("<ce><code128>" & right(sSATChaveAcesso, 22) & "</code128></ce>")
    
    Call objDarumaNaoFiscalImp.ImprimeNormal("<ce><qrcode>" & sSATChaveAcesso & "|" & sSATtimeStamp & "|" & objVenda.objCupomFiscal.sSATvalorTotalCFe & "|" & objVenda.objCupomFiscal.sSATCPFCNPJValue & "|" & objVenda.objCupomFiscal.sSATQRCode & "<lmodulo>4</lmodulo></qrcode></ce>")
    
    '???????????????
    
    Call objDarumaNaoFiscalImp.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", ""))
    Call objDarumaNaoFiscalImp.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "DADOS DO CUPOM FISCAL ELETRÔNICO DE CANCELAMENTO"))
    Call objDarumaNaoFiscalImp.ImprimeNormal("<b>" & Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "SAT No. " & objVenda.objCupomFiscal.sNumSerieECF) & "</b>")
    
    sSATtimeStamp = objVenda.objCupomFiscal.sSATCanctimeStamp
    Call objDarumaNaoFiscalImp.ImprimeNormal(Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", Mid(sSATtimeStamp, 7, 2) & "/" & Mid(sSATtimeStamp, 5, 2) & "/" & left(sSATtimeStamp, 4) & " - " & Mid(sSATtimeStamp, 9, 2) & ":" & Mid(sSATtimeStamp, 11, 2) & ":" & Mid(sSATtimeStamp, 13, 2)))
    
    sSATChaveAcesso = Mid(objVenda.objCupomFiscal.sSATCancChaveAcesso, 4)
    
    Call objDarumaNaoFiscalImp.ImprimeNormal(Mid(sSATChaveAcesso, 1, 11) & " " & Mid(sSATChaveAcesso, 12, 11) & " " & Mid(sSATChaveAcesso, 23, 11) & " " & Mid(sSATChaveAcesso, 34, 11))
    
    Call objDarumaNaoFiscalImp.ImprimeNormal("<ce><code128>" & left(sSATChaveAcesso, 22) & "</code128></ce>")
    Call objDarumaNaoFiscalImp.ImprimeNormal("<ce><code128>" & right(sSATChaveAcesso, 22) & "</code128></ce>")
    
    Call objDarumaNaoFiscalImp.ImprimeNormal("<ce><qrcode>" & sSATChaveAcesso & "|" & sSATtimeStamp & "|" & objVenda.objCupomFiscal.sSATvalorTotalCFe & "|" & objVenda.objCupomFiscal.sSATCPFCNPJValue & "|" & objVenda.objCupomFiscal.sSATQRCode & "<lmodulo>4</lmodulo></qrcode></ce>")
    
    Call objDarumaNaoFiscalImp.ImprimeNormal("<sl>4</sl><confgui>P</confgui><gui></gui>")
    
    Call objDarumaNaoFiscalImp.TerminaImpressao
        
    SAT_Imprime_Canc_Daruma = SUCESSO
    
    Exit Function
    
Erro_SAT_Imprime_Canc_Daruma:

    SAT_Imprime_Canc_Daruma = gErr

    Select Case gErr

        Case 99803, 99894 To 99898, 105868, 133088, 204319, 204324, 204329, 210470

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149439)

    End Select

    Exit Function

End Function

Private Function SAT_ObterInfo(ByVal objCupomFiscal As ClassCupomFiscal) As Long
'pegar dados de SATinfo, para poder cancelar, se necessário

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(1 To 1) As Long, dtDataEmissao As Date, dHoraEmissao As Double, sQRCode As String, itpAmb As Integer
Dim sValorTotalCFE As String, sCPFCNPJValue As String, sArqXml As String, sCancArqXml As String

On Error GoTo Erro_SAT_ObterInfo
    
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_AbrirExt(glConexaoPAFECF)
        If alComando(iIndice) = 0 Then gError 201548
    Next
    
    sQRCode = String(512, 0)
    sArqXml = String(512, 0)
    sCancArqXml = String(512, 0)
    sValorTotalCFE = String(30, 0)
    sCPFCNPJValue = String(20, 0)
    
    lErro = Comando_Executar(alComando(1), "SELECT DataEmissao, HoraEmissao, QRCode, tpAmb, ValorTotalCFE, CPFCNPJValue, ArqXml, CancArqXml FROM SATInfo WHERE ChaveAcesso = ?", dtDataEmissao, dHoraEmissao, sQRCode, itpAmb, sValorTotalCFE, sCPFCNPJValue, sArqXml, sCancArqXml, objCupomFiscal.sSATChaveAcesso)
    If lErro <> AD_SQL_SUCESSO Then gError 201549
    
    lErro = Comando_BuscarProximo(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201550
    
    If lErro = AD_SQL_SUCESSO Then
        objCupomFiscal.sSATtimeStamp = Format(dtDataEmissao, "yyyymmdd") & Format(dHoraEmissao, "hhmmss")
        objCupomFiscal.iSATtpAmb = itpAmb
        objCupomFiscal.sSATQRCode = sQRCode
        objCupomFiscal.sSATvalorTotalCFe = sValorTotalCFE
        objCupomFiscal.sSATCPFCNPJValue = sCPFCNPJValue
        objCupomFiscal.sSATArqXml = sArqXml
        objCupomFiscal.sSATCancArqXml = sCancArqXml
    End If
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    SAT_ObterInfo = SUCESSO
    
    Exit Function
    
Erro_SAT_ObterInfo:

    SAT_ObterInfo = gErr

    Select Case gErr

        Case 201548
            Call Rotina_ErroECF(vbOKOnly, ERRO_ABERTURA_COMANDO, gErr)
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 201551)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Private Function EpsonNaoFiscal_ImprimirNFCE(ByVal objVenda As ClassVenda) As Long

Dim lErro As Long
Dim objItens As ClassItemCupomFiscal
Dim objProduto As ClassProduto
Dim iIndice As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim bCont As Boolean
Dim sMensagem As String
Dim dTotal As Double
Dim dPreco As Double
Dim bAchou As Boolean
Dim colOrcamento As New Collection
Dim iIndice1 As Integer
Dim objCliente As ClassCliente
Dim lIndice As Long
Dim sCliente As String
Dim sCPFCGC As String
Dim sVendedor As String
Dim objVendedor As ClassVendedor
Dim sProduto As String
Dim sCNPJ As String
Dim iPos As Integer
Dim lNumDAV As Long
Dim sRetorno As String
Dim objProdutoNomeRed As TextBox, objMovCaixa As ClassMovimentoCaixa
Dim sTexto As String, sChaveAcesso As String, sTexto1 As String, sTexto2 As String, iAux As Integer, iItem As Integer
Dim objEpsonNaoFiscal As New ClassEpsonNaoFiscalImp

On Error GoTo Erro_EpsonNaoFiscal_ImprimirNFCE

    Call objEpsonNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", gsNomeEmpresa))
    Call Formata_CNPJ_CPF(sCNPJ, gsCNPJ)
    Call objEpsonNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "CNPJ:" & sCNPJ & " IE:" & gsInscricaoEstadual & " IM:" & gsInscricaoMunicipal))
    Call objEpsonNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", gsEndereco))
    Call objEpsonNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", ""))
    Call objEpsonNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "DANFE NFC-e - Documento Auxiliar"))
    Call objEpsonNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "da Nota Fiscal de Consumidor Eletronica"))
    Call objEpsonNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "Não permite aproveitamento de crédito de ICMS"))
    Call objEpsonNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", ""))
    Call objEpsonNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "Codigo Descricao         Qtde Un  VlrUnit VlrTot"))
    Call objEpsonNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", ""))

    iItem = 0
    
    For Each objItens In objVenda.objCupomFiscal.colItens
    
        If objItens.iStatus <> STATUS_CANCELADO Then
    
            iItem = iItem + 1
    
            Set objProduto = New ClassProduto
            
            lErro = CF_ECF("Produtos_Le", objItens.sProduto, objProduto)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 214855
            
            If lErro = SUCESSO Then
                objItens.sProdutoNomeRed = objProduto.sNomeReduzido
            Else
                'caso o produto não seja encontrado
                gError 99894
            End If
                
            sTexto1 = objProduto.sCodigo & " " & objProduto.sDescricao
            sTexto1 = left(sTexto1, 48)
                
            sTexto2 = Format(objItens.dQuantidade, "0.0###")
            If right(sTexto2, 2) = ",0" Then sTexto2 = left(sTexto2, Len(sTexto2) - 2)
            sTexto2 = sTexto2 & " " & objItens.sUnidadeMed
            sTexto2 = sTexto2 & " X " & Format(objItens.dPrecoUnitario, "standard")
            sTexto2 = sTexto2 & " " & Format(objItens.dPrecoUnitario * objItens.dQuantidade, "standard")
                
            iAux = Len(sTexto1 & " " & sTexto2)
            If Len(sTexto1 & " " & sTexto2) <= 48 Then
                Call objEpsonNaoFiscal.ImprimeNormal(sTexto1 & String(48 - iAux, " ") & " " & sTexto2)
            Else
                If Len(sTexto1 & " " & sTexto2) <= 67 Then
                    Call objEpsonNaoFiscal.ImprimeFormatado(sTexto1 & String(67 - iAux, " ") & " " & sTexto2, 1, 0, 0, 0, 0)
                Else
                    Call objEpsonNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", sTexto1))
                    Call objEpsonNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_ESQUERDA, 48, " ", sTexto2))
                End If
            End If
            
            dPreco = objItens.dPrecoUnitario * objItens.dQuantidade
                        
            'se existir desconto sobre o item...
            If objItens.dPercDesc > 0 Then
                Call objEpsonNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 15, " ", "DESC. ITEM:") & Formata_Campo(ALINHAMENTO_DIREITA, 10, " ", Format(objItens.dPercDesc, "standard") & "%") & "     =  " & Formata_Campo(ALINHAMENTO_ESQUERDA, 15, " ", Format(objItens.dPrecoUnitario * objItens.dQuantidade * (1 - (objItens.dPercDesc / 100)), "standard")))
                dPreco = dPreco - (objItens.dPrecoUnitario * objItens.dQuantidade * objItens.dPercDesc / 100)
            ElseIf objItens.dValorDesconto > 0 Then
                'Call objEpsonNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 15, " ", "DESC. ITEM:") & Formata_Campo(ALINHAMENTO_DIREITA, 10, " ", Format(-objItens.dValorDesconto, "standard")) & "     =  " & Formata_Campo(ALINHAMENTO_ESQUERDA, 15, " ", Format((objItens.dPrecoUnitario - objItens.dValorDesconto) * objItens.dQuantidade, "standard")))
                'dPreco = dPreco - (objItens.dQuantidade * objItens.dValorDesconto)
                Call objEpsonNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 15, " ", "DESC. ITEM:") & Formata_Campo(ALINHAMENTO_DIREITA, 10, " ", Format(-objItens.dValorDesconto, "standard")) & "     =  " & Formata_Campo(ALINHAMENTO_ESQUERDA, 15, " ", Format((objItens.dPrecoUnitario * objItens.dQuantidade - objItens.dValorDesconto), "standard")))
                dPreco = dPreco - objItens.dValorDesconto
            End If
            
            dTotal = dTotal + dPreco
        
        End If
        
    Next
                      
    Call objEpsonNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", ""))
    Call objEpsonNaoFiscal.ImprimeNormal(FormataTextoEspacosNoMeio("QTD. TOTAL DE ITENS", Format(iItem, "000"), 48))
    Call objEpsonNaoFiscal.ImprimeNormal(FormataTextoEspacosNoMeio("VALOR TOTAL R$", Format(dTotal, "standard"), 48))
    
    If (objVenda.objCupomFiscal.dValorDesconto + objVenda.objCupomFiscal.dValorDesconto1) <> 0 Then
        Call objEpsonNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_ESQUERDA, 48, " ", "DESCONTO  :  -" & Format(objVenda.objCupomFiscal.dValorDesconto + objVenda.objCupomFiscal.dValorDesconto1, "standard")))
        dTotal = dTotal - (objVenda.objCupomFiscal.dValorDesconto + objVenda.objCupomFiscal.dValorDesconto1)
    End If
    
    If objVenda.objCupomFiscal.dValorAcrescimo <> 0 Then
        Call objEpsonNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_ESQUERDA, 48, " ", "ACRESCIMO :  +" & Format(objVenda.objCupomFiscal.dValorAcrescimo, "standard")))
        dTotal = dTotal + objVenda.objCupomFiscal.dValorAcrescimo
    End If
    
    If (objVenda.objCupomFiscal.dValorDesconto + objVenda.objCupomFiscal.dValorDesconto1) <> 0 Or objVenda.objCupomFiscal.dValorAcrescimo <> 0 Then
        Call objEpsonNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_ESQUERDA, 48, " ", "TOTAL LIQ.:   " & Format(dTotal, "standard")))
    End If
    
    Call objEpsonNaoFiscal.ImprimeNormal(FormataTextoEspacosNoMeio("FORMA DE PAGAMENTO", "Valor Pago", 48))
    For Each objMovCaixa In objVenda.colMovimentosCaixa
        If objMovCaixa.iTipo <> MOVIMENTOCAIXA_TROCO_DINHEIRO And objMovCaixa.iTipo <> MOVIMENTOCAIXA_TROCO_CONTRAVALE And objMovCaixa.iTipo <> MOVIMENTOCAIXA_TROCO_VALE Then
            Call objEpsonNaoFiscal.ImprimeNormal(FormataTextoEspacosNoMeio(MeioPagtoCfe_Descricao(objMovCaixa.iCodigoCFe), Format(objMovCaixa.dValor, "standard"), 48))
        Else
            Call objEpsonNaoFiscal.ImprimeNormal(FormataTextoEspacosNoMeio("Troco R$", Format(objMovCaixa.dValor, "standard"), 48))
        End If
    Next
    Call objEpsonNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", ""))
    Call objEpsonNaoFiscal.ImprimeNormal("Informação dos Tributos Totais Incidentes (Lei F")
    Call objEpsonNaoFiscal.ImprimeNormal(FormataTextoEspacosNoMeio("ederal 12.741 /2012) R$", Format(Venda_ObterTotTrib(objVenda), "standard"), 48))
    Call objEpsonNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", ""))
    
    sTexto1 = objVenda.objCupomFiscal.sNFCeMensagem
    If Len(sTexto1) <> 0 Then
        Do While Len(sTexto1) <> 0
            sTexto2 = left(sTexto1, 48)
            Call objEpsonNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", sTexto2))
            sTexto1 = Mid(sTexto1, 49)
        Loop
        Call objEpsonNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", ""))
    End If
    
    If objVenda.objCupomFiscal.iNFetpAmb = 2 Then
        objEpsonNaoFiscal.ImprimeNormal (Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "EMITIDA EM AMBIENTE DE HOMOLOGACAO"))
        objEpsonNaoFiscal.ImprimeNormal (Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "SEM VALOR FISCAL"))
    End If
    If objVenda.objCupomFiscal.sNFenProt = "" Then objEpsonNaoFiscal.ImprimeNormal (Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "EMITIDA EM CONTINGENCIA"))
    
    Call objEpsonNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "No.:" & CStr(objVenda.objCupomFiscal.lNumero) & " Serie:" & objVenda.objCupomFiscal.sNumSerieECF & " Emissao:" & Format(objVenda.objCupomFiscal.dtDataEmissao, "dd/mm/yyyy") & " " & Format(objVenda.objCupomFiscal.dHoraEmissao, "hh:mm:ss")))
    Call objEpsonNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "Via Consumidor"))
    
    sChaveAcesso = objVenda.objCupomFiscal.sNFeChaveAcesso
    Call objEpsonNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "Consulte pela chave de acesso em:"))
    Call objEpsonNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", URLConsultaNFCE(left(sChaveAcesso, 2))))
    Call objEpsonNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "CHAVE DE ACESSO"))
    Call objEpsonNaoFiscal.ImprimeFormatado(Formata_Campo(ALINHAMENTO_CENTRALIZADO, 67, " ", Mid(sChaveAcesso, 1, 4) & " " & Mid(sChaveAcesso, 5, 4) & " " & Mid(sChaveAcesso, 9, 4) & " " & Mid(sChaveAcesso, 13, 4) & " " & Mid(sChaveAcesso, 17, 4) & " " & _
        Mid(sChaveAcesso, 21, 4) & " " & Mid(sChaveAcesso, 25, 4) & " " & Mid(sChaveAcesso, 29, 4) & " " & Mid(sChaveAcesso, 33, 4) & " " & Mid(sChaveAcesso, 37, 4) & " " & Mid(sChaveAcesso, 41, 4)), 1, 0, 0, 0, 0)
    
    Call objEpsonNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", ""))
    
    If Len(objVenda.objCupomFiscal.sCPFCGC) <> 0 Then
        Call Formata_CNPJ_CPF(sCPFCGC, objVenda.objCupomFiscal.sCPFCGC)
        Call objEpsonNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "CNPJ/CPF/Id.Estrangeiro: " & sCPFCGC))
        
        Call objEpsonNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "  Nome: " & objVenda.objCupomFiscal.sNomeCliente))
        
        If Len(Trim(objVenda.objCupomFiscal.sEndEntLogradouro)) <> 0 Then
            Call objEpsonNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "  End.: " & objVenda.objCupomFiscal.sEndEntLogradouro & ", " & objVenda.objCupomFiscal.sEndEntNúmero))
            Call objEpsonNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "Compl.: " & objVenda.objCupomFiscal.sEndEntComplemento))
            Call objEpsonNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "Bairro: " & objVenda.objCupomFiscal.sEndEntBairro))
        End If
        
    Else
        Call objEpsonNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "CONSUMIDOR NAO IDENTIFICADO"))
    End If
    
    Call objEpsonNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", ""))
    
    Call objEpsonNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "Consulte via leitor de QR Code"))
    Call objEpsonNaoFiscal.ImprimeQRCode(objVenda.objCupomFiscal.sNFCeQRCode)
    
    If objVenda.objCupomFiscal.sNFenProt <> "" Then
        Call objEpsonNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "Protocolo de Autorização:" & " " & objVenda.objCupomFiscal.sNFenProt))
        Call objEpsonNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", Format(objVenda.objCupomFiscal.dtNFeData, "dd/mm/yyyy") & " " & Format(objVenda.objCupomFiscal.dNFeHora, "hh:mm:ss")))
    End If
    
    Call objEpsonNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", ""))
    Call objEpsonNaoFiscal.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", ""))
        
    Call objEpsonNaoFiscal.ImprimeNormal("<gui></gui>")
    
    Call objEpsonNaoFiscal.TerminaImpressao
    
    EpsonNaoFiscal_ImprimirNFCE = SUCESSO
    
    Exit Function
    
Erro_EpsonNaoFiscal_ImprimirNFCE:

    EpsonNaoFiscal_ImprimirNFCE = gErr

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 201571)

    End Select

    Exit Function

End Function


'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function BackupConfig_Le(ByVal objBackupConfig As ClassBackupConfig) As Long

Dim lErro As Long
Dim lComando As Long
Dim tBackupConfig As typeBackupConfig

On Error GoTo Erro_BackupConfig_Le

    If giLocalOperacao = LOCALOPERACAO_ECF Then
    
        'Executa a abertura do Comando
        lComando = Comando_AbrirExt(glConexaoPAFECF)
        If lComando = 0 Then gError 208987
    
        'Alocação de espaço no buffer
        tBackupConfig.sDescricao = String(STRING_MAXIMO, 0)
        tBackupConfig.sDiretorio = String(STRING_MAXIMO, 0)
        tBackupConfig.sFTPURL = String(STRING_MAXIMO, 0)
        tBackupConfig.sFTPUsu = String(STRING_MAXIMO, 0)
        tBackupConfig.sFTPSenha = String(STRING_MAXIMO, 0)
        tBackupConfig.sFTPDir = String(STRING_MAXIMO, 0)
        tBackupConfig.sDirDownload = String(STRING_MAXIMO, 0)
    
        'Le a tabelaBackupConfig
        lErro = Comando_Executar(lComando, "SELECT Codigo, Descricao, Habilitado, DataInicio, Hora, RepetirDias, " & _
                    "Diretorio, IncluirDataNomeArq, DataUltBkp, DataProxBkp, Compactar, TransfFTP, FTPURL, FTPUsu, FTPSenha, FTPDir, DirDownload FROM BackupConfig WHERE Codigo= ? ", _
                    tBackupConfig.lCodigo, tBackupConfig.sDescricao, tBackupConfig.iHabilitado, tBackupConfig.dtDataInicio, _
                    tBackupConfig.dHora, tBackupConfig.iRepetirDias, tBackupConfig.sDiretorio, tBackupConfig.iIncluirDataNomeArq, tBackupConfig.dtDataUltBkp, _
                    tBackupConfig.dtDataProxBkp, tBackupConfig.iCompactar, tBackupConfig.iTransfFTP, tBackupConfig.sFTPURL, tBackupConfig.sFTPUsu, tBackupConfig.sFTPSenha, tBackupConfig.sFTPDir, tBackupConfig.sDirDownload, _
                    objBackupConfig.lCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 208988
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208989
    
        'Sem Dados
        If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS
    
        objBackupConfig.lCodigo = tBackupConfig.lCodigo
        objBackupConfig.sDescricao = tBackupConfig.sDescricao
        objBackupConfig.iHabilitado = tBackupConfig.iHabilitado
        objBackupConfig.dtDataInicio = tBackupConfig.dtDataInicio
        objBackupConfig.dHora = tBackupConfig.dHora
        objBackupConfig.iRepetirDias = tBackupConfig.iRepetirDias
        objBackupConfig.sDiretorio = tBackupConfig.sDiretorio
        objBackupConfig.iIncluirDataNomeArq = tBackupConfig.iIncluirDataNomeArq
        objBackupConfig.dtDataUltBkp = tBackupConfig.dtDataUltBkp
        objBackupConfig.dtDataProxBkp = tBackupConfig.dtDataProxBkp
    
        objBackupConfig.iTransfFTP = tBackupConfig.iTransfFTP
        objBackupConfig.iCompactar = tBackupConfig.iCompactar
        objBackupConfig.sFTPURL = tBackupConfig.sFTPURL
        objBackupConfig.sFTPUsu = tBackupConfig.sFTPUsu
        objBackupConfig.sFTPSenha = tBackupConfig.sFTPSenha
        objBackupConfig.sFTPDir = tBackupConfig.sFTPDir
        objBackupConfig.sDirDownload = tBackupConfig.sDirDownload
    
        'Fecha Comando
        Call Comando_Fechar(lComando)
        
    End If

    BackupConfig_Le = SUCESSO

    Exit Function

Erro_BackupConfig_Le:

    BackupConfig_Le = gErr

    Select Case gErr

        Case 208987
            Call Rotina_ErroECF(vbOKOnly, ERRO_ABERTURA_COMANDO, gErr)

        Case 208988, 208989
            Call Rotina_ErroECF(vbOKOnly, ERRO_LEITURA_BACKUPCONFIG, gErr)

        Case ERRO_LEITURA_SEM_DADOS

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 208990)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function Backup_Verifica_Habilitado(iBKPAtivo As Integer) As Long

Dim lErro As Long
Dim lComando As Long
Dim iAux As Integer

On Error GoTo Erro_Backup_Verifica_Habilitado

    If giLocalOperacao = LOCALOPERACAO_ECF Then
           
        iBKPAtivo = DESMARCADO
    
        'Executa a abertura do Comando
        lComando = Comando_AbrirExt(glConexaoPAFECF)
        If lComando = 0 Then gError 209025
    
        'Le a tabelaBackupConfig
        lErro = Comando_Executar(lComando, "SELECT Habilitado FROM BackupConfig WHERE DataProxBkp <> ? AND Habilitado <> 0 ", _
                    iAux, DATA_NULA)
        If lErro <> AD_SQL_SUCESSO Then gError 209026
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209027
    
        'Sem Dados
        If lErro <> AD_SQL_SEM_DADOS Then iBKPAtivo = MARCADO
        
        'Fecha Comando
        Call Comando_Fechar(lComando)
        
    End If

    Backup_Verifica_Habilitado = SUCESSO

    Exit Function

Erro_Backup_Verifica_Habilitado:

    Backup_Verifica_Habilitado = gErr

    Select Case gErr

        Case 209025
            Call Rotina_ErroECF(vbOKOnly, ERRO_ABERTURA_COMANDO, gErr)

        Case 209026, 209027
            Call Rotina_ErroECF(vbOKOnly, ERRO_LEITURA_BACKUPCONFIG, gErr)

        Case ERRO_LEITURA_SEM_DADOS

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 209028)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Carrega_Movimento_Atual(ByVal objVenda As ClassVenda) As Long

Dim lErro As Long
Dim alComando(1 To 3) As Long
Dim sMsg As String, iEditavel As Integer
Dim colRegistro As New Collection
Dim iIndice As Integer

On Error GoTo Erro_Carrega_Movimento_Atual
    
   'Abre os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_AbrirExt(glConexaoPAFECF)
        If alComando(iIndice) = 0 Then gError 210425
    Next
                    
    sMsg = String(STRING_MOVIMENTOCAIXA_MSG, 0)

    lErro = Comando_Executar(alComando(1), "SELECT Msg, Editavel FROM MovimentoCaixaAtual ORDER BY Seq", sMsg, iEditavel)
    If lErro <> AD_SQL_SUCESSO Then gError 210428

    lErro = Comando_BuscarProximo(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 210429
    
    Do While lErro = AD_SQL_SUCESSO
                
        colRegistro.Add sMsg
                
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 210430
        
    Loop
        
    If colRegistro.Count <> 0 Then
    
        lErro = Registro_ECF_Vendas(colRegistro, "", 0, objVenda)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        objVenda.objCupomFiscal.bEditavel = (iEditavel <> 0)
                
    End If
                
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Carrega_Movimento_Atual = SUCESSO
    
    Exit Function
    
Erro_Carrega_Movimento_Atual:
    
    Carrega_Movimento_Atual = gErr
    
    Select Case gErr
            
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error$, 165191)

    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
End Function

Public Function NTK_ImprimirPedido(ByVal objTela As Object, ByVal objPedido As ClassNTKOrder) As Long

Dim lErro As Long, bRelAberto As Boolean
Dim colMensagem As New Collection, sProdQtde As String, sProdValor As String
Dim sMensagem As String, iIndice As Integer, sTexto1 As String, sTexto2 As String
Dim objRel As Object, objItem As ClassNTKItem, objItem2 As ClassNTKItem

On Error GoTo Erro_NTK_ImprimirPedido

    lErro = AFRAC_AbrirRelatorioGerencial2(RELGER_PEDIDO_NTK, objTela, objRel)
    bRelAberto = True
    
    lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Abrir Pedido NTK")
    If lErro <> SUCESSO Then gError 133088
    
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", gsNomeEmpresa)
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", "")
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "NUMERO DO PEDIDO: " & objPedido.order_number)
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "  DATA DO PEDIDO: " & Format(objPedido.order_date, "dd/mm/yy") & " " & Format(objPedido.order_time, "hh:mm"))
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", "")
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "ITENS DO PEDIDO")
    
    For Each objItem In objPedido.items
    
        colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "")
        
        sProdQtde = CStr(objItem.qtd) & " X " & objItem.Name
        sProdValor = " R$ " & Format(objItem.subtotal_cents / 100, "standard")
        
        If Len(sProdQtde & sProdValor) <= 48 Then
            colMensagem.Add FormataTextoEspacosNoMeio(sProdQtde, sProdValor, 48)
        Else
            colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", sProdQtde)
            colMensagem.Add Formata_Campo(ALINHAMENTO_ESQUERDA, 48, " ", sProdValor)
        End If
        
        If Len(Trim(objItem.obs)) <> 0 Then
    
            sTexto1 = "      Obs: " & objItem.obs
            
            Do While Len(sTexto1) <> 0
                sTexto2 = left(sTexto1, 48)
                colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", sTexto2)
                sTexto1 = Mid(sTexto1, 49)
            Loop
            
        End If
    
        For Each objItem2 In objItem.complements
        
            colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "")
            colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "         + " & objItem2.Name & IIf(objItem2.obs <> "", " " & objItem2.obs, ""))
        
        Next
    
    Next
    
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", "")
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "TOTAL PRODUTOS: R$ " & Format((objPedido.total - objPedido.delivery_fee) / 100, "standard"))
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "TAXA ENTREGA  : R$ " & Format(objPedido.delivery_fee / 100, "standard"))
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "TOTAL         : R$ " & Format(objPedido.total / 100, "standard"))
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", "")
    
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "DADOS DO CLIENTE")
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "")
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "  Nome: " & objPedido.customer_name)
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "  Tel.: " & objPedido.customer_phone)
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "  End.: " & objPedido.address.street) & ", " & objPedido.address.Number
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "Compl.: " & objPedido.address.complement)
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "Bairro: " & objPedido.address.neighborhood)
    
    If Len(Trim(objPedido.address.reference)) <> 0 Then

        sTexto1 = "  Ref.: " & objPedido.address.reference
        
        Do While Len(sTexto1) <> 0
            sTexto2 = left(sTexto1, 48)
            colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", sTexto2)
            sTexto1 = Mid(sTexto1, 49)
        Loop
        
    End If
    
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", "")
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "FORMA DE PAGAMENTO: " & objPedido.payment_type & IIf(objPedido.troco <> 0, ", troco p/R$ " & Format(objPedido.troco, "standard"), ""))
    
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "")
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "")
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "")
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "")
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "")
    colMensagem.Add Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "")
    
    For iIndice = 1 To colMensagem.Count
            
        sMensagem = colMensagem.Item(iIndice)
        If sMensagem = "" Then sMensagem = "    "
        lErro = AFRAC_ImprimirRelatorioGerencial2(sMensagem, objTela, objRel)
        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Pedido NTK")
        If lErro <> SUCESSO Then gError 99893
            
    Next
        
    'guilhotina
    lErro = AFRAC_AcionarGuilhotina2("P", objRel)
    lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Acionar Guilhotina")
    'If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    lErro = AFRAC_FecharRelatorioGerencial2(objTela, objRel)
    bRelAberto = False
    
    lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Fecha Pedido NTK")
    'If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    NTK_ImprimirPedido = SUCESSO
    
    Exit Function
    
Erro_NTK_ImprimirPedido:

    NTK_ImprimirPedido = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 99893
            Call AFRAC_FecharRelatorioGerencial(objTela)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149439)

    End Select

    If bRelAberto Then Call AFRAC_FecharRelatorioGerencial2(objTela, objRel)

    Exit Function

End Function

Private Function ImpressoraESC_ImprimirNFCE(ByVal objVenda As ClassVenda) As Long

Dim lErro As Long
Dim objItens As ClassItemCupomFiscal
Dim objProduto As ClassProduto
Dim iIndice As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim bCont As Boolean
Dim sMensagem As String
Dim dTotal As Double
Dim dPreco As Double
Dim bAchou As Boolean
Dim colOrcamento As New Collection
Dim iIndice1 As Integer
Dim objCliente As ClassCliente
Dim lIndice As Long
Dim sCliente As String
Dim sCPFCGC As String
Dim sVendedor As String
Dim objVendedor As ClassVendedor
Dim sProduto As String
Dim sCNPJ As String
Dim iPos As Integer
Dim lNumDAV As Long
Dim sRetorno As String
Dim objProdutoNomeRed As TextBox, objMovCaixa As ClassMovimentoCaixa
Dim sTexto As String, sChaveAcesso As String, sTexto1 As String, sTexto2 As String, iAux As Integer, iItem As Integer
Dim objImpressoraESC As New ClassImpressoraESC

On Error GoTo Erro_ImpressoraESC_ImprimirNFCE

    Call objImpressoraESC.Define_Impressora(gobjNFeInfo.sPortaImpressora, gobjNFeInfo.iModeloImpressora)

    Call objImpressoraESC.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", gsNomeEmpresa))
    Call Formata_CNPJ_CPF(sCNPJ, gsCNPJ)
    Call objImpressoraESC.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "CNPJ:" & sCNPJ & " IE:" & gsInscricaoEstadual & " IM:" & gsInscricaoMunicipal))
    Call objImpressoraESC.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", gsEndereco))
    Call objImpressoraESC.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", ""))
    Call objImpressoraESC.ImprimeNormal(Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "DANFE NFC-e - Documento Auxiliar"))
    Call objImpressoraESC.ImprimeNormal(Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "da Nota Fiscal de Consumidor Eletronica"))
    Call objImpressoraESC.ImprimeNormal(Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "Não permite aproveitamento de crédito de ICMS"))
    Call objImpressoraESC.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", ""))
    Call objImpressoraESC.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "Codigo Descricao         Qtde Un  VlrUnit VlrTot"))
    Call objImpressoraESC.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", ""))

    iItem = 0
    
    For Each objItens In objVenda.objCupomFiscal.colItens
    
        If objItens.iStatus <> STATUS_CANCELADO Then
    
            iItem = iItem + 1
    
            Set objProduto = New ClassProduto
            
            lErro = CF_ECF("Produtos_Le", objItens.sProduto, objProduto)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 214855
            
            If lErro = SUCESSO Then
                objItens.sProdutoNomeRed = objProduto.sNomeReduzido
            Else
                'caso o produto não seja encontrado
                gError 99894
            End If
                
            sTexto1 = objProduto.sCodigo & " " & objProduto.sDescricao
            sTexto1 = left(sTexto1, 48)
                
            sTexto2 = Format(objItens.dQuantidade, "0.0###")
            If right(sTexto2, 2) = ",0" Then sTexto2 = left(sTexto2, Len(sTexto2) - 2)
            sTexto2 = sTexto2 & " " & objItens.sUnidadeMed
            sTexto2 = sTexto2 & " X " & Format(objItens.dPrecoUnitario, "standard")
            sTexto2 = sTexto2 & " " & Format(objItens.dPrecoUnitario * objItens.dQuantidade, "standard")
                
            iAux = Len(sTexto1 & " " & sTexto2)
            If Len(sTexto1 & " " & sTexto2) <= 48 Then
                Call objImpressoraESC.ImprimeNormal(sTexto1 & String(48 - iAux, " ") & " " & sTexto2)
            Else
                If Len(sTexto1 & " " & sTexto2) <= 67 Then
                    Call objImpressoraESC.ImprimeFormatado(sTexto1 & String(67 - iAux, " ") & " " & sTexto2, 1, 0, 0, 0, 0)
                Else
                    Call objImpressoraESC.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", sTexto1))
                    Call objImpressoraESC.ImprimeNormal(Formata_Campo(ALINHAMENTO_ESQUERDA, 48, " ", sTexto2))
                End If
            End If
            
            dPreco = objItens.dPrecoUnitario * objItens.dQuantidade
                        
            'se existir desconto sobre o item...
            If objItens.dPercDesc > 0 Then
                Call objImpressoraESC.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 15, " ", "DESC. ITEM:") & Formata_Campo(ALINHAMENTO_DIREITA, 10, " ", Format(objItens.dPercDesc, "standard") & "%") & "     =  " & Formata_Campo(ALINHAMENTO_ESQUERDA, 15, " ", Format(objItens.dPrecoUnitario * objItens.dQuantidade * (1 - (objItens.dPercDesc / 100)), "standard")))
                dPreco = dPreco - (objItens.dPrecoUnitario * objItens.dQuantidade * objItens.dPercDesc / 100)
            ElseIf objItens.dValorDesconto > 0 Then
                'Call objImpressoraESC.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 15, " ", "DESC. ITEM:") & Formata_Campo(ALINHAMENTO_DIREITA, 10, " ", Format(-objItens.dValorDesconto, "standard")) & "     =  " & Formata_Campo(ALINHAMENTO_ESQUERDA, 15, " ", Format((objItens.dPrecoUnitario - objItens.dValorDesconto) * objItens.dQuantidade, "standard")))
                'dPreco = dPreco - (objItens.dQuantidade * objItens.dValorDesconto)
                Call objImpressoraESC.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 15, " ", "DESC. ITEM:") & Formata_Campo(ALINHAMENTO_DIREITA, 10, " ", Format(-objItens.dValorDesconto, "standard")) & "     =  " & Formata_Campo(ALINHAMENTO_ESQUERDA, 15, " ", Format((objItens.dPrecoUnitario * objItens.dQuantidade - objItens.dValorDesconto), "standard")))
                dPreco = dPreco - objItens.dValorDesconto
            End If
            
            dTotal = dTotal + dPreco
        
        End If
        
    Next
                      
    Call objImpressoraESC.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", ""))
    Call objImpressoraESC.ImprimeNormal(FormataTextoEspacosNoMeio("QTD. TOTAL DE ITENS", Format(iItem, "000"), 48))
    Call objImpressoraESC.ImprimeNormal(FormataTextoEspacosNoMeio("VALOR TOTAL R$", Format(dTotal, "standard"), 48))
    
    If (objVenda.objCupomFiscal.dValorDesconto + objVenda.objCupomFiscal.dValorDesconto1) <> 0 Then
        Call objImpressoraESC.ImprimeNormal(Formata_Campo(ALINHAMENTO_ESQUERDA, 48, " ", "DESCONTO  :  -" & Format(objVenda.objCupomFiscal.dValorDesconto + objVenda.objCupomFiscal.dValorDesconto1, "standard")))
        dTotal = dTotal - (objVenda.objCupomFiscal.dValorDesconto + objVenda.objCupomFiscal.dValorDesconto1)
    End If
    
    If objVenda.objCupomFiscal.dValorAcrescimo <> 0 Then
        Call objImpressoraESC.ImprimeNormal(Formata_Campo(ALINHAMENTO_ESQUERDA, 48, " ", "ACRESCIMO :  +" & Format(objVenda.objCupomFiscal.dValorAcrescimo, "standard")))
        dTotal = dTotal + objVenda.objCupomFiscal.dValorAcrescimo
    End If
    
    If (objVenda.objCupomFiscal.dValorDesconto + objVenda.objCupomFiscal.dValorDesconto1) <> 0 Or objVenda.objCupomFiscal.dValorAcrescimo <> 0 Then
        Call objImpressoraESC.ImprimeNormal(Formata_Campo(ALINHAMENTO_ESQUERDA, 48, " ", "TOTAL LIQ.:   " & Format(dTotal, "standard")))
    End If
    
    Call objImpressoraESC.ImprimeNormal(FormataTextoEspacosNoMeio("FORMA DE PAGAMENTO", "Valor Pago", 48))
    For Each objMovCaixa In objVenda.colMovimentosCaixa
        If objMovCaixa.iTipo <> MOVIMENTOCAIXA_TROCO_DINHEIRO And objMovCaixa.iTipo <> MOVIMENTOCAIXA_TROCO_CONTRAVALE And objMovCaixa.iTipo <> MOVIMENTOCAIXA_TROCO_VALE Then
            Call objImpressoraESC.ImprimeNormal(FormataTextoEspacosNoMeio(MeioPagtoCfe_Descricao(objMovCaixa.iCodigoCFe), Format(objMovCaixa.dValor, "standard"), 48))
        Else
            Call objImpressoraESC.ImprimeNormal(FormataTextoEspacosNoMeio("Troco R$", Format(objMovCaixa.dValor, "standard"), 48))
        End If
    Next
    Call objImpressoraESC.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", ""))
    Call objImpressoraESC.ImprimeNormal("Informação dos Tributos Totais Incidentes (Lei F")
    Call objImpressoraESC.ImprimeNormal(FormataTextoEspacosNoMeio("ederal 12.741 /2012) R$", Format(Venda_ObterTotTrib(objVenda), "standard"), 48))
    Call objImpressoraESC.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", ""))
    
    sTexto1 = objVenda.objCupomFiscal.sNFCeMensagem
    If Len(sTexto1) <> 0 Then
        Do While Len(sTexto1) <> 0
            sTexto2 = left(sTexto1, 48)
            Call objImpressoraESC.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", sTexto2))
            sTexto1 = Mid(sTexto1, 49)
        Loop
        Call objImpressoraESC.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", ""))
    End If
    
    If objVenda.objCupomFiscal.iNFetpAmb = 2 Then
        objImpressoraESC.ImprimeNormal (Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "EMITIDA EM AMBIENTE DE HOMOLOGACAO"))
        objImpressoraESC.ImprimeNormal (Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "SEM VALOR FISCAL"))
    End If
    If objVenda.objCupomFiscal.sNFenProt = "" Then objImpressoraESC.ImprimeNormal (Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "EMITIDA EM CONTINGENCIA"))
    
    Call objImpressoraESC.ImprimeNormal(Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "No.:" & CStr(objVenda.objCupomFiscal.lNumero) & " Serie:" & objVenda.objCupomFiscal.sNumSerieECF & " Emissao:" & Format(objVenda.objCupomFiscal.dtDataEmissao, "dd/mm/yyyy") & " " & Format(objVenda.objCupomFiscal.dHoraEmissao, "hh:mm:ss")))
    Call objImpressoraESC.ImprimeNormal(Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "Via Consumidor"))
    
    sChaveAcesso = objVenda.objCupomFiscal.sNFeChaveAcesso
    Call objImpressoraESC.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "Consulte pela chave de acesso em:"))
    Call objImpressoraESC.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", URLConsultaNFCE(left(sChaveAcesso, 2))))
    Call objImpressoraESC.ImprimeNormal(Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "CHAVE DE ACESSO"))
    Call objImpressoraESC.ImprimeFormatado(Formata_Campo(ALINHAMENTO_DIREITA, 62, " ", Mid(sChaveAcesso, 1, 4) & " " & Mid(sChaveAcesso, 5, 4) & " " & Mid(sChaveAcesso, 9, 4) & " " & Mid(sChaveAcesso, 13, 4) & " " & Mid(sChaveAcesso, 17, 4) & " " & _
        Mid(sChaveAcesso, 21, 4) & " " & Mid(sChaveAcesso, 25, 4) & " " & Mid(sChaveAcesso, 29, 4) & " " & Mid(sChaveAcesso, 33, 4) & " " & Mid(sChaveAcesso, 37, 4) & " " & Mid(sChaveAcesso, 41, 4)), 1, 0, 0, 0, 0)
    
    Call objImpressoraESC.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", ""))
    
    If Len(objVenda.objCupomFiscal.sCPFCGC) <> 0 Then
        
        Call Formata_CNPJ_CPF(sCPFCGC, objVenda.objCupomFiscal.sCPFCGC)
        Call objImpressoraESC.ImprimeNormal(Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "CNPJ/CPF/Id.Estrangeiro: " & sCPFCGC))
        
        Call objImpressoraESC.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "  Nome: " & objVenda.objCupomFiscal.sNomeCliente))
        
        If Len(Trim(objVenda.objCupomFiscal.sEndEntLogradouro)) <> 0 Then
            Call objImpressoraESC.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "  End.: " & objVenda.objCupomFiscal.sEndEntLogradouro & ", " & objVenda.objCupomFiscal.sEndEntNúmero))
            Call objImpressoraESC.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "Compl.: " & objVenda.objCupomFiscal.sEndEntComplemento))
            Call objImpressoraESC.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "Bairro: " & objVenda.objCupomFiscal.sEndEntBairro))
        End If
        
    Else
        Call objImpressoraESC.ImprimeNormal(Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "CONSUMIDOR NAO IDENTIFICADO"))
    End If
    
    Call objImpressoraESC.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", ""))
    
    Call objImpressoraESC.ImprimeNormal(Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "Consulte via leitor de QR Code"))
    Call objImpressoraESC.ImprimeQRCode(objVenda.objCupomFiscal.sNFCeQRCode)
    
    If objVenda.objCupomFiscal.sNFenProt <> "" Then
        Call objImpressoraESC.ImprimeNormal(Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "Protocolo de Autorização:" & " " & objVenda.objCupomFiscal.sNFenProt))
        Call objImpressoraESC.ImprimeNormal(Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", Format(objVenda.objCupomFiscal.dtNFeData, "dd/mm/yyyy") & " " & Format(objVenda.objCupomFiscal.dNFeHora, "hh:mm:ss")))
    End If
    
    Call objImpressoraESC.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", ""))
    Call objImpressoraESC.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", ""))
        
    Call objImpressoraESC.AcionarGuilhotina("P")
    
    Call objImpressoraESC.TerminaImpressao
    
    ImpressoraESC_ImprimirNFCE = SUCESSO
    
    Exit Function
    
Erro_ImpressoraESC_ImprimirNFCE:

    ImpressoraESC_ImprimirNFCE = gErr

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 201571)

    End Select

    Exit Function

End Function

Public Function NTK_BaixarPedidos(ByVal objTela As Object) As Long

Dim lErro As Long, objIntegracao As Object, colPedidos As New Collection, lNum As Long
Dim objPedido As ClassNTKOrder, objItem As ClassNTKItem
Dim objVenda As New ClassVenda, sNumOrcamento As String, objPedidoAux As ClassNTKOrder

Dim bAchou As Boolean, iItem As Integer
Dim objProduto As ClassProduto
Dim objItemCupom As New ClassItemCupomFiscal
Dim sProduto As String
Dim lNumero As Long
Dim objAliquota As ClassAliquotaICMS
Dim objCliente As ClassCliente
Dim sCPF As String
Dim sRet As String
Dim lErro1 As Long
Dim sPrecoItem As String
Dim sAliquota As String
Dim sPeso As String
Dim sPrecoKilo As String
Dim sPrecoTotal As String
Dim sProduto1 As String
Dim objAliquota1 As ClassAliquotaICMS
Dim iTotalizador As Integer
Dim sSubtotal As String
Dim sPrecoItem1 As String, bDescontoValor As Boolean
Dim objMovCaixa As ClassMovimentoCaixa, dSaldoFrete As Double, dValorFreteItem As Double, dValorLiquido As Double, lIBGECidade As Long
Dim sEndEntCidade As String
Dim iImprimePV As Integer, sFiltro As String, iCopias As Integer

On Error GoTo Erro_NTK_BaixarPedidos

    iImprimePV = GetPrivateProfileInt(APLICACAO_ECF, "NTKImprimePV", 1, NOME_ARQUIVO_CAIXA)

    Set objIntegracao = CreateObject(gobjNFeInfo.PGM_SGENFE & ".NTKIntegracao")
    
    sFiltro = Replace("&date=" & DateAdd("h", -2, Now), " ", "%20")
    
    lErro = objIntegracao.ObterPedidos(gsNTKToken, gsNTKMerchantId, gsNTKURLServidor, sFiltro, colPedidos)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    For Each objPedidoAux In colPedidos
    
        sNumOrcamento = "0001" + right("000000" + objPedidoAux.order_number, 6)
    
        Set objVenda = New ClassVenda
        
        'verificar se o pedido já é um orcamento
        objVenda.objCupomFiscal.lNumOrcamento = CLng(sNumOrcamento)

        lErro = CF_ECF("OrcamentoECF_Le", objVenda)
        If lErro <> SUCESSO And lErro <> 204690 And lErro <> 210447 Then gError ERRO_SEM_MENSAGEM

        'se nao for
        If lErro = 204690 Then
        
            'pular pedidos cancelados
            
            If objPedidoAux.Status <> 3 Then
            
                If objPedidoAux.items.Count = 0 Then
                
                    lErro = objIntegracao.ObterPedido(gsNTKToken, gsNTKURLServidor, objPedidoAux.id, objPedido)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                Else
                
                    Set objPedido = objPedidoAux
                    
                End If
                
                Set objVenda = New ClassVenda
        
                objVenda.iTipo = OPTION_DAV
        
                'copiar dados do pedido externo para obj
                With objVenda.objCupomFiscal
                    .iTipo = OPTION_DAV
                    .dtDataEmissao = objPedido.order_date
                    .dHoraEmissao = objPedido.order_time
                    .dValorTroco = 0
                    .lDuracao = 0
                    .iFilialEmpresa = giFilialEmpresa
                    .iCodCaixa = giCodCaixa
                    .iECF = giCodECF
                    .iTabelaPreco = gobjLojaECF.iTabelaPreco
                    .dValorTotal = Arredonda_Moeda(objPedido.total / 100#)
                    .dValorProdutos = Arredonda_Moeda((objPedido.total - objPedido.delivery_fee) / 100#)
                    .dtDataReducao = gdtDataAnterior
                    .lNumOrcamento = CLng(sNumOrcamento)
                    .lNumeroDAV = .lNumOrcamento
                    .dValorAcrescimo = Arredonda_Moeda(objPedido.delivery_fee / 100#)
                    .IdExterno = objPedido.id
                    .NumCaixa = 0
                    .sEndEntBairro = left(objPedido.address.neighborhood, 60)
                    .sEndEntCidade = left(objPedido.address.city, 60)
                    .sEndEntComplemento = left(objPedido.address.complement, 60)
                    .sEndEntLogradouro = left(objPedido.address.street, 60)
                    .sEndEntNúmero = left(objPedido.address.Number, 60)
                    .sEndEntUF = objPedido.address.uf
                    .StatusExterno = objPedido.Status
                    
                    lIBGECidade = 0
                    If .sEndEntCidade <> "" And .sEndEntUF <> "" Then
                    
                        sEndEntCidade = .sEndEntCidade
                        lErro = CF_ECF("Cidade_ObtemIBGE", .sEndEntUF, sEndEntCidade, lIBGECidade)
                        If lErro = SUCESSO Then .sEndEntCidade = sEndEntCidade
                        
                    End If
                    .lEndEntIBGECidade = lIBGECidade
                    
                    .sCPFCGC = objPedido.cpf
                    .sCPFCGC1 = objPedido.cpf
                    .sEndEntEmail = ""
                    .sNomeCliente = objPedido.customer_name
                    
                End With
                
                'resseta estas variaveis
                Set objVenda.colMovimentosCaixa = New Collection
                Set objVenda.colCheques = Nothing
                Set objVenda.colTroca = Nothing
                Set objVenda.objCarne = Nothing
                
                Set objMovCaixa = New ClassMovimentoCaixa
                
                objMovCaixa.dtDataMovimento = objPedido.order_date
                objMovCaixa.dValor = objVenda.objCupomFiscal.dValorTotal
                objMovCaixa.iFilialEmpresa = giFilialEmpresa
                objMovCaixa.iCaixa = giCodCaixa
                objMovCaixa.iCodOperador = giCodOperador
                objMovCaixa.dHora = objPedido.order_time
                
                Select Case objPedido.payment_type
                    
                    Case "Cartão de Crédito", "ONLINE", "Cartão - Crédito"
                        objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_CARTAOCREDITO
                        objMovCaixa.iTipoCartao = TIPO_POS
                        
                    Case "Cartão de Débito", "Cartão - Débito"
                        objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_CARTAODEBITO
                        objMovCaixa.iTipoCartao = TIPO_POS
                        
                    Case Else
                        objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_DINHEIRO
                        objMovCaixa.iAdmMeioPagto = MEIO_PAGAMENTO_DINHEIRO
                    
                End Select
                
                objVenda.colMovimentosCaixa.Add objMovCaixa
                
                iItem = 0
                
                For Each objItem In objPedido.items
                
                    iItem = iItem + 1
                    
                    Set objItemCupom = New ClassItemCupomFiscal
                                
                    If objItem.Codigo <> "" Then
                    
                        sProduto1 = objItem.Codigo
                    
                    Else
                    
                        sProduto1 = objItem.Name
                        
                    End If
                    
                    Call TP_Produto_Le_Col(gaobjProdutosReferencia, gaobjProdutosCodBarras, gaobjProdutosNome, sProduto1, objProduto)
                            
                    'caso o produto não seja encontrado
                    If objProduto Is Nothing Then
                        
                        'se tentou pelo codigo e nao conseguiu entao tentar pelo nome
                        If objItem.Codigo <> "" Then
                        
                            sProduto1 = objItem.Name
                            Call TP_Produto_Le_Col(gaobjProdutosReferencia, gaobjProdutosCodBarras, gaobjProdutosNome, sProduto1, objProduto)
                            
                        End If
                        
                        'caso o produto não seja encontrado
                        If objProduto Is Nothing Then gError 99606
                    
                    End If
                    
                    For Each objAliquota In gcolAliquotasTotal
                        If objAliquota.sSigla = objProduto.sICMSAliquota Then
                            objItemCupom.dAliquotaICMS = objAliquota.dAliquota
                            Exit For
                        End If
                    Next
                    
                    objItemCupom.dQuantidade = StrParaDbl(objItem.qtd)
                    objItemCupom.dPrecoUnitario = Arredonda_Moeda(objItem.subtotal_cents / (objItem.qtd * 100))
                    objItemCupom.sUnidadeMed = objProduto.sSiglaUMVenda
                    objItemCupom.sSituacaoTrib = objProduto.sSituacaoTribECF
                    objItemCupom.sProduto = objProduto.sCodigo
                    objItemCupom.sProdutoNomeRed = objProduto.sNomeReduzido
                    objItemCupom.sProdutoDescricao = objProduto.sDescricao
                    
                    objItemCupom.iItem = iItem
                    objItemCupom.iStatus = STATUS_ATIVO
                    
                    If objItemCupom.dAliquotaICMS > 0 Then
                        
                        If objProduto.sSituacaoTribECF = TIPOTRIBISS_SITUACAOTRIBECF_INTEGRAL Then
                            
                            sAliquota = TIPOTRIBISS_SITUACAOTRIBECF_INTEGRAL & Format(objItemCupom.dAliquotaICMS * 10000, "0000")
                            
                            objItemCupom.sSituacaoTrib = sAliquota
                            
                        Else
                            
                            sAliquota = Format(objItemCupom.dAliquotaICMS * 10000, "0000")
                            objItemCupom.sSituacaoTrib = TIPOTRIBICMS_SITUACAOTRIBECF_INTEGRAL & sAliquota
                            
                        End If
                    Else
                        sAliquota = left(objProduto.sSituacaoTribECF, 1)
                        objItemCupom.sSituacaoTrib = sAliquota & "1"
                    End If
                
                    'Joga na col
                    objVenda.objCupomFiscal.colItens.Add objItemCupom
                    
                    Set objItemCupom.objTributacaoDocItem = New ClassTributacaoDocItem
                    Call objItemCupom.objTributacaoDocItem.Copia(objProduto.objTributacaoDocItem)
                    
                Next
                
                'ratear frete, se houver
                If objVenda.objCupomFiscal.dValorAcrescimo <> 0 Then
                
                    dSaldoFrete = objVenda.objCupomFiscal.dValorAcrescimo
                    iItem = 0
                    
                    For Each objItemCupom In objVenda.objCupomFiscal.colItens
                    
                        iItem = iItem + 1
                        
                        'se for o ultimo item
                        If iItem = objVenda.objCupomFiscal.colItens.Count Then
                        
                            dValorFreteItem = dSaldoFrete
                            
                        Else
    
                            dValorLiquido = Arredonda_Moeda(Arredonda_Moeda(objItemCupom.dQuantidade * objItemCupom.dPrecoUnitario) - Arredonda_Moeda(objItemCupom.dValorDesconto))
                            dValorFreteItem = Arredonda_Moeda(objVenda.objCupomFiscal.dValorAcrescimo * dValorLiquido / objVenda.objCupomFiscal.dValorTotal)
                            dSaldoFrete = Arredonda_Moeda(dSaldoFrete - dValorFreteItem)
                        
                        End If
                    
                        objItemCupom.objTributacaoDocItem.dValorFreteItem = dValorFreteItem
                        
                    Next
                    
                    dValorFreteItem = 0
                    For Each objItemCupom In objVenda.objCupomFiscal.colItens
                        If objItemCupom.iStatus <> STATUS_CANCELADO Then
                            If objItemCupom.objTributacaoDocItem.dValorFreteItem < 0 Then objItemCupom.objTributacaoDocItem.dValorFreteItem = 0
                            dValorFreteItem = dValorFreteItem + objItemCupom.objTributacaoDocItem.dValorFreteItem
                        End If
                    Next
                    dSaldoFrete = Arredonda_Moeda(objVenda.objCupomFiscal.dValorAcrescimo - dValorFreteItem)
                    Do While Abs(dSaldoFrete) > DELTA_VALORMONETARIO
                        For Each objItemCupom In objVenda.objCupomFiscal.colItens
                            If objItemCupom.iStatus <> STATUS_CANCELADO Then
                                If dSaldoFrete > 0 Then
                                    objItemCupom.objTributacaoDocItem.dValorFreteItem = objItemCupom.objTributacaoDocItem.dValorFreteItem + 0.01
                                    dSaldoFrete = dSaldoFrete - 0.01
                                Else
                                    If objItemCupom.objTributacaoDocItem.dValorFreteItem - 0.01 > -DELTA_VALORMONETARIO2 Then
                                        objItemCupom.objTributacaoDocItem.dValorFreteItem = objItemCupom.objTributacaoDocItem.dValorFreteItem - 0.01
                                    End If
                                    dSaldoFrete = dSaldoFrete + 0.01
                                End If
                                If Abs(dSaldoFrete) < DELTA_VALORMONETARIO Then Exit For
                            End If
                        Next
                    Loop
                
                End If
                
                For Each objItemCupom In objVenda.objCupomFiscal.colItens
                
                    Call ItemCupom_AjustaTrib(objItemCupom)
                
                Next
                    
                'gravar pedido como orcamento
                lErro = CF_ECF("Grava_Orcamento_ECF1", objVenda)
                If lErro <> SUCESSO Then gError 105915
                
                For iCopias = 1 To iImprimePV
                
                    'imprimir o pedido externo
                    lErro = CF_ECF("NTK_ImprimirPedido", objTela, objPedido)
                    If lErro <> SUCESSO Then gError 105915
                
                Next
            
            End If
            
        End If
    
    Next
    
    NTK_BaixarPedidos = SUCESSO
    
    Exit Function
    
Erro_NTK_BaixarPedidos:

    NTK_BaixarPedidos = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 201601)

    End Select

    Exit Function

End Function

Private Sub ItemCupom_AjustaTrib(objItem As ClassItemCupomFiscal)
'Ajusta a tributacao de acordo com o que foi efetivamente vendido
'??? falta tratar ST

Dim objTributacaoDocItem As ClassTributacaoDocItem
Dim dValorLiquido As Double

    Set objTributacaoDocItem = objItem.objTributacaoDocItem

    dValorLiquido = Arredonda_Moeda(Arredonda_Moeda(objItem.dQuantidade * objItem.dPrecoUnitario) - Arredonda_Moeda(objItem.dValorDesconto) + objTributacaoDocItem.dValorFreteItem)
    
    'dados gerais
    objTributacaoDocItem.dQuantidade = objItem.dQuantidade
    objTributacaoDocItem.dQtdTrib = objItem.dQuantidade
    objTributacaoDocItem.dPrecoUnitario = Arredonda_Moeda(objItem.dPrecoUnitario)
    objTributacaoDocItem.dValorUnitTrib = Arredonda_Moeda(objItem.dPrecoUnitario)
    objTributacaoDocItem.dDescontoGrid = Arredonda_Moeda(objItem.dValorDesconto)
    
    If objTributacaoDocItem.dTotTrib <> 0 Then
        objTributacaoDocItem.dTotTrib = Arredonda_Moeda(objTributacaoDocItem.dTotTrib * dValorLiquido / 1000)
    End If
    
    'icms
    If objTributacaoDocItem.dICMSAliquota <> 0 Then
        objTributacaoDocItem.dICMSBase = dValorLiquido
        objTributacaoDocItem.dICMSValor = Arredonda_Moeda(objTributacaoDocItem.dICMSBase * objTributacaoDocItem.dICMSAliquota)
    Else
        objTributacaoDocItem.dICMSBase = 0
        objTributacaoDocItem.dICMSValor = 0
    End If
    
    'FCP
    If objTributacaoDocItem.dICMSpFCP <> 0 Then
        objTributacaoDocItem.dICMSvBCFCP = dValorLiquido
        objTributacaoDocItem.dICMSvFCP = Arredonda_Moeda(objTributacaoDocItem.dICMSvBCFCP * objTributacaoDocItem.dICMSpFCP)
    Else
        objTributacaoDocItem.dICMSvBCFCP = 0
        objTributacaoDocItem.dICMSvFCP = 0
    End If
    
    'pis
    If objTributacaoDocItem.dPISAliquota <> 0 Then
        objTributacaoDocItem.dPISBase = dValorLiquido
        objTributacaoDocItem.dPISValor = Arredonda_Moeda(objTributacaoDocItem.dPISBase * objTributacaoDocItem.dPISAliquota)
    Else
        If objTributacaoDocItem.dPISAliquotaValor <> 0 Then
            objTributacaoDocItem.dPISBase = 0 '??? rever
            objTributacaoDocItem.dPISQtde = objItem.dQuantidade  '??? rever pq pode ter que converter unidade de medida
            objTributacaoDocItem.dPISValor = Arredonda_Moeda(objTributacaoDocItem.dPISQtde * objTributacaoDocItem.dPISAliquota)
        Else
            objTributacaoDocItem.dPISBase = 0
            objTributacaoDocItem.dPISValor = 0
    
        End If
    End If
    
    'cofins
    If objTributacaoDocItem.dCOFINSAliquota <> 0 Then
        objTributacaoDocItem.dCOFINSBase = dValorLiquido
        objTributacaoDocItem.dCOFINSValor = Arredonda_Moeda(objTributacaoDocItem.dCOFINSBase * objTributacaoDocItem.dCOFINSAliquota)
    Else
        If objTributacaoDocItem.dCOFINSAliquotaValor <> 0 Then
            objTributacaoDocItem.dCOFINSQtde = objItem.dQuantidade  '??? rever pq pode ter que converter unidade de medida
            objTributacaoDocItem.dCOFINSValor = Arredonda_Moeda(objTributacaoDocItem.dCOFINSQtde * objTributacaoDocItem.dCOFINSAliquota)
        Else
            objTributacaoDocItem.dCOFINSBase = 0
            objTributacaoDocItem.dCOFINSValor = 0
    
        End If
    End If
    
    If objTributacaoDocItem.dISSAliquota <> 0 Then
        objTributacaoDocItem.dISSBase = dValorLiquido
        objTributacaoDocItem.dISSValor = Arredonda_Moeda(objTributacaoDocItem.dISSBase * objTributacaoDocItem.dISSAliquota)
    Else
        objTributacaoDocItem.dISSBase = 0
        objTributacaoDocItem.dISSValor = 0
    End If
    
End Sub

Public Function Importar_Pedidos(ByVal objTela As Object) As Long

Dim lErro As Long

On Error GoTo Erro_Importar_Pedidos

    If InStr(gsIntegracaoTipo, CStr(INTEGRACAO_TIPO_NTK)) <> 0 Then
    
        lErro = NTK_BaixarPedidos(objTela)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    End If
    
    If InStr(gsIntegracaoTipo, CStr(INTEGRACAO_TIPO_PP)) <> 0 Then
    
        lErro = PP_Importar_Pedidos(objTela)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    End If
    
    If InStr(gsIntegracaoTipo, CStr(INTEGRACAO_TIPO_IFOOD)) <> 0 Then
    
        lErro = IFOOD_Importar_Pedidos(objTela)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    End If
    
    Importar_Pedidos = SUCESSO
    
    Exit Function
    
Erro_Importar_Pedidos:

    Importar_Pedidos = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 201601)

    End Select

    Exit Function

End Function

Private Function PP_Importar_Pedido(ByVal objTela As Object, ByVal objPedido As ClassPPPedido, ByVal iImprimePV As Integer) As Long

Dim lErro As Long, objIntegracao As Object, lNum As Long
Dim objItem As ClassPPItemPedido
Dim objVenda As New ClassVenda, sNumOrcamento As String
Dim bAchou As Boolean, iItem As Integer
Dim objProduto As ClassProduto
Dim objItemCupom As New ClassItemCupomFiscal
Dim sProduto As String
Dim lNumero As Long
Dim objAliquota As ClassAliquotaICMS
Dim objCliente As ClassCliente
Dim sCPF As String
Dim sRet As String
Dim lErro1 As Long
Dim sPrecoItem As String
Dim sAliquota As String
Dim sPeso As String
Dim sPrecoKilo As String
Dim sPrecoTotal As String
Dim sProduto1 As String
Dim objAliquota1 As ClassAliquotaICMS
Dim iTotalizador As Integer
Dim sSubtotal As String
Dim sPrecoItem1 As String, bDescontoValor As Boolean
Dim objMovCaixa As ClassMovimentoCaixa, dSaldoFrete As Double, dValorFreteItem As Double, dValorLiquido As Double, lIBGECidade As Long
Dim sEndEntCidade As String

On Error GoTo Erro_PP_Importar_Pedido

    sNumOrcamento = "0002" + right("000000" + CStr(objPedido.IdCarrinho), 6)

    Set objVenda = New ClassVenda
    
    'verificar se o pedido já é um orcamento
    objVenda.objCupomFiscal.lNumOrcamento = CLng(sNumOrcamento)

    lErro = CF_ECF("OrcamentoECF_Le", objVenda)
    If lErro <> SUCESSO And lErro <> 204690 And lErro <> 210447 Then gError ERRO_SEM_MENSAGEM

    'se nao for
    If lErro = 204690 Then
    
        Set objVenda = New ClassVenda
    
        objVenda.iTipo = OPTION_DAV
    
        'copiar dados do pedido externo para obj
        With objVenda.objCupomFiscal
            .iTipo = OPTION_DAV
            .dtDataEmissao = objPedido.DataRegistro
            .dHoraEmissao = objPedido.HoraRegistro
            .dValorTroco = 0
            .lDuracao = 0
            .iFilialEmpresa = giFilialEmpresa
            .iCodCaixa = giCodCaixa
            .iECF = giCodECF
            .iTabelaPreco = gobjLojaECF.iTabelaPreco
            .dValorTotal = Arredonda_Moeda(objPedido.ValorTotal)
            .dValorProdutos = Arredonda_Moeda(objPedido.ValorTotal - objPedido.TaxaEntrega)
            .dtDataReducao = gdtDataAnterior
            .lNumOrcamento = CLng(sNumOrcamento)
            .lNumeroDAV = .lNumOrcamento
            .dValorAcrescimo = Arredonda_Moeda(objPedido.TaxaEntrega)
            .IdExterno = objPedido.IdCarrinho
            .NumCaixa = 0
            .sEndEntBairro = left(objPedido.EnderecoBairro, 60)
            .sEndEntCidade = left(objPedido.EnderecoCidade, 60)
            .sEndEntComplemento = left(objPedido.EnderecoComplemento, 60)
            .sEndEntLogradouro = left(objPedido.EnderecoLogradouro, 60)
            .sEndEntNúmero = left(objPedido.EnderecoNumero, 60)
            .sEndEntUF = objPedido.EnderecoUF
            .StatusExterno = objPedido.Status
            
            lIBGECidade = 0
            If .sEndEntCidade <> "" And .sEndEntUF <> "" Then
            
                sEndEntCidade = .sEndEntCidade
                lErro = CF_ECF("Cidade_ObtemIBGE", .sEndEntUF, sEndEntCidade, lIBGECidade)
                If lErro = SUCESSO Then .sEndEntCidade = sEndEntCidade
                
            End If
            .lEndEntIBGECidade = lIBGECidade
            
            .sCPFCGC = ""
            .sCPFCGC1 = ""
            .sEndEntEmail = left(objPedido.email, 60)
            .sNomeCliente = objPedido.Nome
            
        End With
        
        'resseta estas variaveis
        Set objVenda.colMovimentosCaixa = New Collection
        Set objVenda.colCheques = Nothing
        Set objVenda.colTroca = Nothing
        Set objVenda.objCarne = Nothing
        
        Set objMovCaixa = New ClassMovimentoCaixa
        
        objMovCaixa.dtDataMovimento = objPedido.DataRegistro
        objMovCaixa.dValor = objVenda.objCupomFiscal.dValorTotal
        objMovCaixa.iFilialEmpresa = giFilialEmpresa
        objMovCaixa.iCaixa = giCodCaixa
        objMovCaixa.iCodOperador = giCodOperador
        objMovCaixa.dHora = objPedido.HoraRegistro
        
        Select Case objPedido.FormaPagto
            
            Case 1
                objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_CARTAOCREDITO
                objMovCaixa.iTipoCartao = TIPO_POS
                
            Case 2
                objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_CARTAODEBITO
                objMovCaixa.iTipoCartao = TIPO_POS
                
            Case Else
                objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_DINHEIRO
                objMovCaixa.iAdmMeioPagto = MEIO_PAGAMENTO_DINHEIRO
            
        End Select
        
        objVenda.colMovimentosCaixa.Add objMovCaixa
        
        iItem = 0
        
        For Each objItem In objPedido.colItens
        
            iItem = iItem + 1
            
            Set objItemCupom = New ClassItemCupomFiscal
                        
            If objItem.ProdutoCodigo <> "" Then
            
                sProduto1 = objItem.ProdutoCodigo
                Call TP_Produto_Le_Col2(gaobjProdutosCodigo, sProduto1, objProduto)
                
            Else
            
                sProduto1 = objItem.ProdutoEan
                Call TP_Produto_Le_Col(gaobjProdutosReferencia, gaobjProdutosCodBarras, gaobjProdutosNome, sProduto1, objProduto)
                
            End If
            
            'caso o produto não seja encontrado
            If objProduto Is Nothing Then
                
                '???????????? apenas para teste
                If objProduto Is Nothing Then
                    
                    sProduto1 = "000090"
                    Call TP_Produto_Le_Col(gaobjProdutosReferencia, gaobjProdutosCodBarras, gaobjProdutosNome, sProduto1, objProduto)
                
                End If
                
                'caso o produto não seja encontrado
                If objProduto Is Nothing Then gError 99606
            
            End If
            
            For Each objAliquota In gcolAliquotasTotal
                If objAliquota.sSigla = objProduto.sICMSAliquota Then
                    objItemCupom.dAliquotaICMS = objAliquota.dAliquota
                    Exit For
                End If
            Next
            
            objItemCupom.dQuantidade = StrParaDbl(objItem.quantidade)
            objItemCupom.dPrecoUnitario = Arredonda_Moeda(objItem.PrecoUnitario)
            objItemCupom.sUnidadeMed = objProduto.sSiglaUMVenda
            objItemCupom.sSituacaoTrib = objProduto.sSituacaoTribECF
            objItemCupom.sProduto = objProduto.sCodigo
            objItemCupom.sProdutoNomeRed = objProduto.sNomeReduzido
            objItemCupom.sProdutoDescricao = objItem.Descricao
            
            objItemCupom.iItem = iItem
            objItemCupom.iStatus = STATUS_ATIVO
            
            If objItemCupom.dAliquotaICMS > 0 Then
                
                If objProduto.sSituacaoTribECF = TIPOTRIBISS_SITUACAOTRIBECF_INTEGRAL Then
                    
                    sAliquota = TIPOTRIBISS_SITUACAOTRIBECF_INTEGRAL & Format(objItemCupom.dAliquotaICMS * 10000, "0000")
                    
                    objItemCupom.sSituacaoTrib = sAliquota
                    
                Else
                    
                    sAliquota = Format(objItemCupom.dAliquotaICMS * 10000, "0000")
                    objItemCupom.sSituacaoTrib = TIPOTRIBICMS_SITUACAOTRIBECF_INTEGRAL & sAliquota
                    
                End If
            Else
                sAliquota = left(objProduto.sSituacaoTribECF, 1)
                objItemCupom.sSituacaoTrib = sAliquota & "1"
            End If
        
            'Joga na col
            objVenda.objCupomFiscal.colItens.Add objItemCupom
            
            Set objItemCupom.objTributacaoDocItem = New ClassTributacaoDocItem
            Call objItemCupom.objTributacaoDocItem.Copia(objProduto.objTributacaoDocItem)
            
        Next
        
        'ratear frete, se houver
        If objVenda.objCupomFiscal.dValorAcrescimo <> 0 Then
        
            dSaldoFrete = objVenda.objCupomFiscal.dValorAcrescimo
            iItem = 0
            
            For Each objItemCupom In objVenda.objCupomFiscal.colItens
            
                iItem = iItem + 1
                
                'se for o ultimo item
                If iItem = objVenda.objCupomFiscal.colItens.Count Then
                
                    dValorFreteItem = dSaldoFrete
                    
                Else
    
                    dValorLiquido = Arredonda_Moeda(Arredonda_Moeda(objItemCupom.dQuantidade * objItemCupom.dPrecoUnitario) - Arredonda_Moeda(objItemCupom.dValorDesconto))
                    dValorFreteItem = Arredonda_Moeda(objVenda.objCupomFiscal.dValorAcrescimo * dValorLiquido / objVenda.objCupomFiscal.dValorTotal)
                    dSaldoFrete = Arredonda_Moeda(dSaldoFrete - dValorFreteItem)
                
                End If
            
                objItemCupom.objTributacaoDocItem.dValorFreteItem = dValorFreteItem
                
            Next
        
            dValorFreteItem = 0
            For Each objItemCupom In objVenda.objCupomFiscal.colItens
                If objItemCupom.iStatus <> STATUS_CANCELADO Then
                    If objItemCupom.objTributacaoDocItem.dValorFreteItem < 0 Then objItemCupom.objTributacaoDocItem.dValorFreteItem = 0
                    dValorFreteItem = dValorFreteItem + objItemCupom.objTributacaoDocItem.dValorFreteItem
                End If
            Next
            dSaldoFrete = Arredonda_Moeda(objVenda.objCupomFiscal.dValorAcrescimo - dValorFreteItem)
            Do While Abs(dSaldoFrete) > DELTA_VALORMONETARIO
                For Each objItemCupom In objVenda.objCupomFiscal.colItens
                    If objItemCupom.iStatus <> STATUS_CANCELADO Then
                        If dSaldoFrete > 0 Then
                            objItemCupom.objTributacaoDocItem.dValorFreteItem = objItemCupom.objTributacaoDocItem.dValorFreteItem + 0.01
                            dSaldoFrete = dSaldoFrete - 0.01
                        Else
                            If objItemCupom.objTributacaoDocItem.dValorFreteItem - 0.01 > -DELTA_VALORMONETARIO2 Then
                                objItemCupom.objTributacaoDocItem.dValorFreteItem = objItemCupom.objTributacaoDocItem.dValorFreteItem - 0.01
                            End If
                            dSaldoFrete = dSaldoFrete + 0.01
                        End If
                        If Abs(dSaldoFrete) < DELTA_VALORMONETARIO Then Exit For
                    End If
                Next
            Loop
                    
        End If
        
        For Each objItemCupom In objVenda.objCupomFiscal.colItens
        
            Call ItemCupom_AjustaTrib(objItemCupom)
        
        Next
            
        'gravar pedido como orcamento
        lErro = CF_ECF("Grava_Orcamento_ECF1", objVenda)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        If iImprimePV <> 0 Then
        
            'imprimir o pedido externo
            lErro = Orcamento_ImpressaoCFE(objTela, objVenda)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        End If
        
    End If
        
    PP_Importar_Pedido = SUCESSO
    
    Exit Function
    
Erro_PP_Importar_Pedido:

    PP_Importar_Pedido = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 201601)

    End Select

    Exit Function

End Function

Public Function PP_Importar_Pedidos(ByVal objTela As Object) As Long

Dim lErro As Long, objIntegracao As Object, lNum As Long
Dim objPedido As ClassPPPedido
Dim iImprimePV As Integer

On Error GoTo Erro_PP_Importar_Pedidos

    iImprimePV = GetPrivateProfileInt(APLICACAO_ECF, "PPImprimePV", 1, NOME_ARQUIVO_CAIXA)

    Set objIntegracao = CreateObject(gobjNFeInfo.PGM_SGENFE & ".PPIntegracao")
    
    Do While True
    
        Set objPedido = New ClassPPPedido
    
        lErro = objIntegracao.Pedido_Importa(gsPPAmbiente, gsPPVersao, gsPPEmailLojista, gsPPSenhaLojista, glPPCodLoja, objPedido)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        lErro = PP_Importar_Pedido(objTela, objPedido, iImprimePV)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        lErro = objIntegracao.Pedido_Confirma_Importacao(gsPPAmbiente, gsPPVersao, gsPPEmailLojista, gsPPSenhaLojista, glPPCodLoja, objPedido.IdCarrinho)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Loop
        
    PP_Importar_Pedidos = SUCESSO
    
    Exit Function
    
Erro_PP_Importar_Pedidos:

    PP_Importar_Pedidos = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 201601)

    End Select

    Exit Function

End Function

Public Function IFOOD_Importar_Pedidos(ByVal objTela As Object) As Long

Dim lErro As Long, objIntegracao As Object
Dim objPedido As ClassiFoodPedido, sArquivoXML As String
Dim lTamanho As Long, iRet As Integer, sRetorno As String
Dim sDiriFood As String, sDiriFoodProc As String, sDiriFoodErr As String

On Error GoTo Erro_IFOOD_Importar_Pedidos

    Set objIntegracao = CreateObject(gobjNFeInfo.PGM_SGENFE & ".IFoodIntegracao")
    
    'Obtem as pastas do iFood
    lTamanho = 255
    sRetorno = String(lTamanho, 0)
    iRet = GetPrivateProfileString(APLICACAO_CAIXA, "DiriFood", "C:\TEMP\", sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
    sDiriFood = left(sRetorno, iRet)

    lTamanho = 255
    sRetorno = String(lTamanho, 0)
    iRet = GetPrivateProfileString(APLICACAO_CAIXA, "DiriFoodProc", "C:\TEMP\Processed\", sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
    sDiriFoodProc = left(sRetorno, iRet)
    
    lTamanho = 255
    sRetorno = String(lTamanho, 0)
    iRet = GetPrivateProfileString(APLICACAO_CAIXA, "DiriFoodErr", "C:\TEMP\Error\", sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
    sDiriFoodErr = left(sRetorno, iRet)
    
    If right(sDiriFood, 1) <> "\" And right(sDiriFood, 1) <> "/" Then sDiriFood = sDiriFood & "\"
    If right(sDiriFoodProc, 1) <> "\" And right(sDiriFoodProc, 1) <> "/" Then sDiriFoodProc = sDiriFoodProc & "\"
    If right(sDiriFoodErr, 1) <> "\" And right(sDiriFoodErr, 1) <> "/" Then sDiriFoodErr = sDiriFoodErr & "\"
    
    sArquivoXML = Dir(sDiriFood & "*.xml")

    Do While Len(sArquivoXML) > 0
    
        Set objPedido = New ClassiFoodPedido
    
        lErro = objIntegracao.Pedido_Importa(sDiriFood & sArquivoXML, objPedido)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        lErro = IFOOD_Importar_Pedido(objTela, objPedido)
        If lErro = SUCESSO Then
            Name sDiriFood & sArquivoXML As sDiriFoodProc & sArquivoXML
        Else
            Name sDiriFood & sArquivoXML As sDiriFoodErr & sArquivoXML
        End If

        sArquivoXML = Dir

    Loop
        
    IFOOD_Importar_Pedidos = SUCESSO
    
    Exit Function
    
Erro_IFOOD_Importar_Pedidos:

    IFOOD_Importar_Pedidos = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 201601)

    End Select

    Exit Function

End Function

Private Function IFOOD_Importar_Pedido(ByVal objTela As Object, ByVal objPedido As ClassiFoodPedido) As Long

Dim lErro As Long, objIntegracao As Object, lNum As Long
Dim objItem As ClassiFoodItem
Dim objPgto As ClassiFoodPagto
Dim objVenda As New ClassVenda, sNumOrcamento As String
Dim bAchou As Boolean, iItem As Integer
Dim objProduto As ClassProduto
Dim objItemCupom As New ClassItemCupomFiscal
Dim sProduto As String
Dim lNumero As Long
Dim objAliquota As ClassAliquotaICMS
Dim objCliente As ClassCliente
Dim sCPF As String
Dim sRet As String
Dim lErro1 As Long
Dim sPrecoItem As String
Dim sAliquota As String
Dim sPeso As String
Dim sPrecoKilo As String
Dim sPrecoTotal As String
Dim sProduto1 As String
Dim objAliquota1 As ClassAliquotaICMS
Dim iTotalizador As Integer
Dim sSubtotal As String
Dim sPrecoItem1 As String, bDescontoValor As Boolean
Dim objMovCaixa As ClassMovimentoCaixa, dSaldoFrete As Double, dValorFreteItem As Double, dValorLiquido As Double, lIBGECidade As Long
Dim sEndEntCidade As String
Dim dValorAux As Double

On Error GoTo Erro_IFOOD_Importar_Pedido

    sNumOrcamento = "0003" + right("000000" + CStr(objPedido.idPedidoCurto), 6)

    Set objVenda = New ClassVenda
    
    'verificar se o pedido já é um orcamento
    objVenda.objCupomFiscal.lNumOrcamento = CLng(sNumOrcamento)

    lErro = CF_ECF("OrcamentoECF_Le", objVenda)
    If lErro <> SUCESSO And lErro <> 204690 And lErro <> 210447 Then gError ERRO_SEM_MENSAGEM

    'se nao for
    If lErro = 204690 Then
    
        Set objVenda = New ClassVenda
    
        objVenda.iTipo = OPTION_DAV
    
        'copiar dados do pedido externo para obj
        With objVenda.objCupomFiscal
            .iTipo = OPTION_DAV
            .dtDataEmissao = StrParaDate(Format(objPedido.dataPedidoComanda, "dd/mm/yyyy"))
            .dHoraEmissao = CDate(Format(objPedido.dataPedidoComanda, "HH:MM:SS"))
            .dValorTroco = objPedido.vlrTroco
            .lDuracao = 0
            .iFilialEmpresa = giFilialEmpresa
            .iCodCaixa = giCodCaixa
            .iECF = giCodECF
            .iTabelaPreco = gobjLojaECF.iTabelaPreco
            .dValorTotal = Arredonda_Moeda(objPedido.vlrTotal)
            .dValorProdutos = Arredonda_Moeda(objPedido.vlrPratos)
            .dtDataReducao = gdtDataAnterior
            .lNumOrcamento = CLng(sNumOrcamento)
            .lNumeroDAV = .lNumOrcamento
            .dValorAcrescimo = Arredonda_Moeda(objPedido.vlrTaxa)
            .IdExterno = objPedido.idPedidoCurto
            .NumCaixa = 0
            .sEndEntBairro = left(objPedido.bairro, 60)
            .sEndEntCidade = left(objPedido.cidade, 60)
            .sEndEntComplemento = left(objPedido.complemento, 60)
            .sEndEntLogradouro = left(objPedido.logradouro, 60)
            .sEndEntNúmero = left(objPedido.logradouroNum, 60)
            .sEndEntUF = objPedido.estado
            .StatusExterno = 0  'objPedido.Status
            
            lIBGECidade = 0
            If .sEndEntCidade <> "" And .sEndEntUF <> "" Then
            
                sEndEntCidade = .sEndEntCidade
                lErro = CF_ECF("Cidade_ObtemIBGE", .sEndEntUF, sEndEntCidade, lIBGECidade)
                If lErro = SUCESSO Then .sEndEntCidade = sEndEntCidade
                
            End If
            .lEndEntIBGECidade = lIBGECidade
            
            .sCPFCGC = ""
            .sCPFCGC1 = ""
            .sEndEntEmail = left(objPedido.email, 60)
            .sNomeCliente = objPedido.Nome
            
        End With
        
        'resseta estas variaveis
        Set objVenda.colMovimentosCaixa = New Collection
        Set objVenda.colCheques = Nothing
        Set objVenda.colTroca = Nothing
        Set objVenda.objCarne = Nothing
        
        For Each objPgto In objPedido.pagamentos
        
            Set objMovCaixa = New ClassMovimentoCaixa
            
            objMovCaixa.dtDataMovimento = objVenda.objCupomFiscal.dtDataEmissao
            objMovCaixa.dValor = objPgto.valor
            objMovCaixa.iFilialEmpresa = giFilialEmpresa
            objMovCaixa.iCaixa = giCodCaixa
            objMovCaixa.iCodOperador = giCodOperador
            objMovCaixa.dHora = objVenda.objCupomFiscal.dHoraEmissao
            
            Select Case objPgto.codTipoCondPagto
                
                '1-Cartão Online, 2-Cheque, 3-Dinheiro, 4-Vale Refeição, C-Crédito, D-Desconto, O-Outros, 5-Cartão Offline , E-Conta Empresa, A-App Online
                Case "C"
                    objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_CARTAOCREDITO
                    objMovCaixa.iTipoCartao = TIPO_POS
                    
                Case "1", "5"
                    objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_CARTAODEBITO
                    objMovCaixa.iTipoCartao = TIPO_POS
                    
                Case "2"
                    objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_CHEQUE
                    objMovCaixa.iAdmMeioPagto = MEIO_PAGAMENTO_CHEQUE
                    
                Case "4"
                    objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_VALETICKET
                    objMovCaixa.iAdmMeioPagto = MEIO_PAGAMENTO_CONTRAVALE
                    
                Case "3"
                    objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_DINHEIRO
                    objMovCaixa.iAdmMeioPagto = MEIO_PAGAMENTO_DINHEIRO
                    
                Case Else
                    objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_OUTROS
                    objMovCaixa.iAdmMeioPagto = MEIO_PAGAMENTO_DINHEIRO
                
            End Select
            
            objMovCaixa.sHistorico = objPgto.bundleKey
            
            objVenda.colMovimentosCaixa.Add objMovCaixa
            
        Next
        
        iItem = 0
        
        For Each objItem In objPedido.colItens
        
            iItem = iItem + 1
            
            'Se é o pai
            If objItem.codPai = "" Then
                
                Set objItemCupom = New ClassItemCupomFiscal
                
                Set objProduto = Nothing
                            
                If objItem.codProdutoPdv <> "" Then
                
                    sProduto1 = objItem.codProdutoPdv
                    Call TP_Produto_Le_Col2(gaobjProdutosCodigo, sProduto1, objProduto)
                        
                End If
                               
                'caso o produto não seja encontrado
                If objProduto Is Nothing Then gError 216227
                           
                For Each objAliquota In gcolAliquotasTotal
                    If objAliquota.sSigla = objProduto.sICMSAliquota Then
                        objItemCupom.dAliquotaICMS = objAliquota.dAliquota
                        Exit For
                    End If
                Next
                
                objItemCupom.dQuantidade = StrParaDbl(objItem.quantidade)
                objItemCupom.dPrecoUnitario = Arredonda_Moeda(objItem.vlrUnitLiq)
                objItemCupom.sUnidadeMed = objProduto.sSiglaUMVenda
                objItemCupom.sSituacaoTrib = objProduto.sSituacaoTribECF
                objItemCupom.sProduto = objProduto.sCodigo
                objItemCupom.sProdutoNomeRed = objProduto.sNomeReduzido
                objItemCupom.sProdutoDescricao = objItem.Descricao
                
                objItemCupom.iItem = iItem
                objItemCupom.iStatus = STATUS_ATIVO
                
                If objItemCupom.dAliquotaICMS > 0 Then
                    
                    If objProduto.sSituacaoTribECF = TIPOTRIBISS_SITUACAOTRIBECF_INTEGRAL Then
                        
                        sAliquota = TIPOTRIBISS_SITUACAOTRIBECF_INTEGRAL & Format(objItemCupom.dAliquotaICMS * 10000, "0000")
                        
                        objItemCupom.sSituacaoTrib = sAliquota
                        
                    Else
                        
                        sAliquota = Format(objItemCupom.dAliquotaICMS * 10000, "0000")
                        objItemCupom.sSituacaoTrib = TIPOTRIBICMS_SITUACAOTRIBECF_INTEGRAL & sAliquota
                        
                    End If
                Else
                    sAliquota = left(objProduto.sSituacaoTribECF, 1)
                    objItemCupom.sSituacaoTrib = sAliquota & "1"
                End If
            
                'Joga na col
                objVenda.objCupomFiscal.colItens.Add objItemCupom
                
                Set objItemCupom.objTributacaoDocItem = New ClassTributacaoDocItem
                Call objItemCupom.objTributacaoDocItem.Copia(objProduto.objTributacaoDocItem)
                
            Else
            
                'Ajusta os preço do produto pai do cardápio para englobar o preço dos itens opcionais
                 objItemCupom.dPrecoUnitario = Arredonda_Moeda(objItemCupom.dPrecoUnitario + ((objItem.vlrUnitLiq) * StrParaDbl(objItem.quantidade)) / objItemCupom.dQuantidade)
            
            End If
                
        Next
        
        'ratear frete, se houver
        If objVenda.objCupomFiscal.dValorAcrescimo <> 0 Then
        
            dSaldoFrete = objVenda.objCupomFiscal.dValorAcrescimo
            iItem = 0
            
            For Each objItemCupom In objVenda.objCupomFiscal.colItens
            
                iItem = iItem + 1
                
                'se for o ultimo item
                If iItem = objVenda.objCupomFiscal.colItens.Count Then
                
                    dValorFreteItem = dSaldoFrete
                    
                Else
    
                    dValorLiquido = Arredonda_Moeda(Arredonda_Moeda(objItemCupom.dQuantidade * objItemCupom.dPrecoUnitario) - Arredonda_Moeda(objItemCupom.dValorDesconto))
                    dValorFreteItem = Arredonda_Moeda(objVenda.objCupomFiscal.dValorAcrescimo * dValorLiquido / objVenda.objCupomFiscal.dValorTotal)
                    dSaldoFrete = Arredonda_Moeda(dSaldoFrete - dValorFreteItem)
                
                End If
            
                objItemCupom.objTributacaoDocItem.dValorFreteItem = dValorFreteItem
                
            Next
        
            dValorFreteItem = 0
            For Each objItemCupom In objVenda.objCupomFiscal.colItens
                If objItemCupom.iStatus <> STATUS_CANCELADO Then
                    If objItemCupom.objTributacaoDocItem.dValorFreteItem < 0 Then objItemCupom.objTributacaoDocItem.dValorFreteItem = 0
                    dValorFreteItem = dValorFreteItem + objItemCupom.objTributacaoDocItem.dValorFreteItem
                End If
            Next
            dSaldoFrete = Arredonda_Moeda(objVenda.objCupomFiscal.dValorAcrescimo - dValorFreteItem)
            Do While Abs(dSaldoFrete) > DELTA_VALORMONETARIO
                For Each objItemCupom In objVenda.objCupomFiscal.colItens
                    If objItemCupom.iStatus <> STATUS_CANCELADO Then
                        If dSaldoFrete > 0 Then
                            objItemCupom.objTributacaoDocItem.dValorFreteItem = objItemCupom.objTributacaoDocItem.dValorFreteItem + 0.01
                            dSaldoFrete = dSaldoFrete - 0.01
                        Else
                            If objItemCupom.objTributacaoDocItem.dValorFreteItem - 0.01 > -DELTA_VALORMONETARIO2 Then
                                objItemCupom.objTributacaoDocItem.dValorFreteItem = objItemCupom.objTributacaoDocItem.dValorFreteItem - 0.01
                            End If
                            dSaldoFrete = dSaldoFrete + 0.01
                        End If
                        If Abs(dSaldoFrete) < DELTA_VALORMONETARIO Then Exit For
                    End If
                Next
            Loop
                    
        End If
        
        For Each objItemCupom In objVenda.objCupomFiscal.colItens
        
            Call ItemCupom_AjustaTrib(objItemCupom)
        
        Next
            
        'gravar pedido como orcamento
        lErro = CF_ECF("Grava_Orcamento_ECF1", objVenda)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
'        If iImprimePV <> 0 Then
'
'            'imprimir o pedido externo
'            lErro = Orcamento_ImpressaoCFE(objTela, objVenda)
'            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
'
'        End If
        
    End If
        
    IFOOD_Importar_Pedido = SUCESSO
    
    Exit Function
    
Erro_IFOOD_Importar_Pedido:

    IFOOD_Importar_Pedido = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case 216227
            Call Rotina_ErroECF(vbOKOnly, "Ocorreu um erro no item " & CStr(objItem.sequencia) & " do pedido " & CStr(objPedido.idPedidoCurto) & " referece ao produto " & objItem.codProdutoPdv & " que não foi localizado no sistema. Dados Adicionais: Cód.Cardapio=" & objItem.codCardapio & " e Descrição=" & objItem.Descricao, gErr)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 201601)

    End Select

    Exit Function

End Function
Public Function SaldoEmDinheiro_Le(dSaldoEmDinheiro As Double, Optional ByVal bRetornaErroSemDados As Boolean = False) As Long

Dim lErro As Long
Dim dValorAux As Double
Dim lComando As Long

On Error GoTo Erro_SaldoEmDinheiro_Le

    dSaldoEmDinheiro = 0
    
    lComando = Comando_AbrirExt(glConexaoPAFECF)
    If lComando = 0 Then gError 216176

    lErro = Comando_Executar(lComando, "SELECT Valor FROM SaldoEmDinheiro ORDER BY NumUltMovto DESC ", dValorAux)
    If lErro <> AD_SQL_SUCESSO Then gError 216177
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 216178
    
    If lErro <> AD_SQL_SUCESSO And bRetornaErroSemDados Then gError ERRO_LEITURA_SEM_DADOS
    
    If lErro = AD_SQL_SUCESSO Then dSaldoEmDinheiro = dValorAux
    
    Call Comando_Fechar(lComando)
    
    SaldoEmDinheiro_Le = SUCESSO
    
    Exit Function
    
Erro_SaldoEmDinheiro_Le:

    SaldoEmDinheiro_Le = gErr

    Select Case gErr

        Case 216176
            Call Rotina_ErroECF(vbOKOnly, ERRO_ABERTURA_COMANDO, gErr)

        Case 216177 To 216178
            Call Rotina_ErroECF(vbOKOnly, ERRO_LEITURA_SALDOEMDINHEIRO, gErr)
            
        Case ERRO_LEITURA_SEM_DADOS

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error$, 216179)

    End Select

    Call Comando_Fechar(lComando)

End Function

Private Function ImpressoraESC_ImprimirSAT(ByVal objVenda As ClassVenda) As Long

Dim lErro As Long
Dim objItem As ClassItemCupomFiscal
Dim objProduto As ClassProduto
Dim iIndice As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim bCont As Boolean
Dim sMensagem As String
Dim dTotal As Double
Dim dPreco As Double
Dim bAchou As Boolean
Dim colOrcamento As New Collection
Dim iIndice1 As Integer
Dim objCliente As ClassCliente
Dim lIndice As Long
Dim sCliente As String
Dim sCPFCGC As String
Dim sVendedor As String
Dim objVendedor As ClassVendedor
Dim sProduto As String
Dim sCNPJ As String
Dim iPos As Integer
Dim lNumDAV As Long, objMovCaixa As ClassMovimentoCaixa
Dim sRetorno As String, sSATChaveAcesso As String
Dim objProdutoNomeRed As TextBox, sSATtimeStamp As String
Dim sTexto As String, sTexto2 As String, iTam As Integer
Dim objImpressoraESC As New ClassImpressoraESC

On Error GoTo Erro_ImpressoraESC_ImprimirSAT

    Call objImpressoraESC.Define_Impressora(gobjSATInfo.sPortaImpressora, gobjSATInfo.iModeloImpressora)

    Set objCliente = gobjClienteCPF.Busca(objVenda.objCupomFiscal.sCPFCGC, lIndice)
    If Not objCliente Is Nothing Then sCliente = objCliente.sNomeReduzido & " - "
    Call Formata_CNPJ_CPF(sCPFCGC, objVenda.objCupomFiscal.sCPFCGC)
    sCliente = sCliente & sCPFCGC
    
    Call Formata_CNPJ_CPF(sCNPJ, gsCNPJ)
    
    Call objImpressoraESC.ImprimeNormal(Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", gsNomeReduzido))
    Call objImpressoraESC.ImprimeNormal(Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", gsNomeEmpresa))
    Call objImpressoraESC.ImprimeNormal(Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", gsEndereco))
    Call objImpressoraESC.ImprimeNormal("")
    Call objImpressoraESC.ImprimeNormal(Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "CNPJ:" & sCNPJ))
    Call objImpressoraESC.ImprimeNormal(Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "IE:" & gsInscricaoEstadual))
    Call objImpressoraESC.ImprimeNormal(Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "IM:" & gsInscricaoMunicipal))
    Call objImpressoraESC.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", ""))
    
    If objVenda.objCupomFiscal.iSATtpAmb <> 1 Then
    
        Call objImpressoraESC.ImprimeFormatado(Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "Extrato No. 000000"), 0, 0, 0, 0, 1)
        Call objImpressoraESC.ImprimeFormatado(Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "CUPOM FISCAL ELETRÔNICO - SAT"), 0, 0, 0, 0, 1)
        Call objImpressoraESC.ImprimeNormal("")
        Call objImpressoraESC.ImprimeNormal(Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "= T E S T E ="))
        Call objImpressoraESC.ImprimeNormal("")
        Call objImpressoraESC.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, ">", ""))
        Call objImpressoraESC.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, ">", ""))
        Call objImpressoraESC.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, ">", ""))
    
    Else
    
        Call objImpressoraESC.ImprimeFormatado(Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "Extrato No. " & Format(objVenda.objCupomFiscal.lNumero, "000000")), 0, 0, 0, 0, 1)
        Call objImpressoraESC.ImprimeFormatado(Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "CUPOM FISCAL ELETRÔNICO - SAT"), 0, 0, 0, 0, 1)
    
    End If
    
    Call objImpressoraESC.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", ""))
    Call objImpressoraESC.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "CPF/CNPJ do Consumidor: " & sCPFCGC))
    Call objImpressoraESC.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", ""))
    Call objImpressoraESC.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "#|COD|DESC|QTD|UN|VL UN R$|VL ITEM R$"))
    Call objImpressoraESC.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", ""))
    
    iIndice = 0
    
    For Each objItem In objVenda.objCupomFiscal.colItens
    
        If objItem.iStatus = STATUS_ATIVO Then
        
            iIndice = iIndice + 1
            
            sTexto = Format(iIndice, "000") & " " & objItem.sProduto & " " & objItem.sProdutoNomeRed & " " & IIf(right(Format(objItem.dQuantidade, "0.0###"), 2) = ",0", Format(objItem.dQuantidade, "###0"), Format(objItem.dQuantidade, "0.0###")) & " " & objItem.sUnidadeMed & " X " & Format(objItem.dPrecoUnitario, "#,###,##0.00")
            sTexto2 = Format(Arredonda_Moeda(objItem.dQuantidade * objItem.dPrecoUnitario, 4), "#,###,##0.00")
            iTam = Len(sTexto) + Len(sTexto2)
            
            If iTam < 48 Then
                Call objImpressoraESC.ImprimeNormal(sTexto & String(48 - iTam, " ") & sTexto2)
            Else
                Call objImpressoraESC.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", sTexto))
                Call objImpressoraESC.ImprimeNormal(Formata_Campo(ALINHAMENTO_ESQUERDA, 48, " ", sTexto2))
            End If
            
            dPreco = Arredonda_Moeda(objItem.dPrecoUnitario * objItem.dQuantidade)
            
            'se existir desconto sobre o item...
            If objItem.dPercDesc > 0 Then
                sTexto = Formata_Campo(ALINHAMENTO_DIREITA, 15, " ", "DESC. ITEM:") & Formata_Campo(ALINHAMENTO_DIREITA, 10, " ", Format(objItem.dPercDesc, "standard") & "%") & " = " & Formata_Campo(ALINHAMENTO_ESQUERDA, 15, " ", Format(objItem.dPrecoUnitario * objItem.dQuantidade * (1 - (objItem.dPercDesc / 100)), "standard"))
                Call objImpressoraESC.ImprimeNormal(sTexto)
                dPreco = dPreco - (objItem.dPrecoUnitario * objItem.dQuantidade * objItem.dPercDesc / 100)
            ElseIf objItem.dValorDesconto > 0 Then
                sTexto = Formata_Campo(ALINHAMENTO_DIREITA, 15, " ", "DESC. ITEM:") & Formata_Campo(ALINHAMENTO_DIREITA, 10, " ", Format(-objItem.dValorDesconto, "standard")) & " = " & Formata_Campo(ALINHAMENTO_ESQUERDA, 15, " ", Format((objItem.dPrecoUnitario - objItem.dValorDesconto) * objItem.dQuantidade, "standard"))
                Call objImpressoraESC.ImprimeNormal(sTexto)
                dPreco = dPreco - objItem.dValorDesconto
            End If
            
            dTotal = dTotal + dPreco
    
        End If
    
    Next
                      
    If (objVenda.objCupomFiscal.dValorDesconto + objVenda.objCupomFiscal.dValorDesconto1) <> 0 Or objVenda.objCupomFiscal.dValorAcrescimo <> 0 Then
        Call objImpressoraESC.ImprimeNormal(Formata_Campo(ALINHAMENTO_ESQUERDA, 48, " ", "Subtotal :   " & Format(dTotal, "standard")))
    
        If (objVenda.objCupomFiscal.dValorDesconto + objVenda.objCupomFiscal.dValorDesconto1) <> 0 Then
            Call objImpressoraESC.ImprimeNormal(Formata_Campo(ALINHAMENTO_ESQUERDA, 48, " ", "Descontos  :  -" & Format(objVenda.objCupomFiscal.dValorDesconto + objVenda.objCupomFiscal.dValorDesconto1, "standard")))
            dTotal = dTotal - (objVenda.objCupomFiscal.dValorDesconto + objVenda.objCupomFiscal.dValorDesconto1)
        End If
    
        If objVenda.objCupomFiscal.dValorAcrescimo <> 0 Then
            Call objImpressoraESC.ImprimeNormal(Formata_Campo(ALINHAMENTO_ESQUERDA, 48, " ", "Acréscimos :  +" & Format(objVenda.objCupomFiscal.dValorAcrescimo, "standard")))
            dTotal = dTotal + objVenda.objCupomFiscal.dValorAcrescimo
        End If

    End If

    Call objImpressoraESC.ImprimeFormatado(Formata_Campo(ALINHAMENTO_ESQUERDA, 48, " ", "TOTAL R$:   " & Format(dTotal, "standard")), 0, 0, 0, 1, 0)
    Call objImpressoraESC.ImprimeNormal("")
    
    If objVenda.colMovimentosCaixa.Count <> 0 Then
    
        If objVenda.colMovimentosCaixa.Item(1).iCodigoCFe = 0 Then
    
            'colocar o CodigoCFe devido nos meios de pagto
            Call CF_ECF("Venda_CalculaCodigoCFe", objVenda)
            
        End If
    
    End If
    
    For Each objMovCaixa In objVenda.colMovimentosCaixa
        If objMovCaixa.iTipo <> MOVIMENTOCAIXA_TROCO_DINHEIRO And objMovCaixa.iTipo <> MOVIMENTOCAIXA_TROCO_CONTRAVALE And objMovCaixa.iTipo <> MOVIMENTOCAIXA_TROCO_VALE Then
            Call objImpressoraESC.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", MeioPagtoCfe_Descricao(objMovCaixa.iCodigoCFe) & ":" & Format(objMovCaixa.dValor, "standard")))
        Else
            Call objImpressoraESC.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "Troco R$:" & Format(objMovCaixa.dValor, "standard")))
        End If
    Next

    Call objImpressoraESC.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, "-", ""))
    Call objImpressoraESC.ImprimeFormatado(Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", "SAT No. " & objVenda.objCupomFiscal.sNumSerieECF), 0, 0, 0, 0, 1)
    
    sSATtimeStamp = objVenda.objCupomFiscal.sSATtimeStamp
    Call objImpressoraESC.ImprimeNormal(Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", Mid(sSATtimeStamp, 7, 2) & "/" & Mid(sSATtimeStamp, 5, 2) & "/" & left(sSATtimeStamp, 4) & " - " & Mid(sSATtimeStamp, 9, 2) & ":" & Mid(sSATtimeStamp, 11, 2) & ":" & Mid(sSATtimeStamp, 13, 2)))
    
    sSATChaveAcesso = Mid(objVenda.objCupomFiscal.sSATChaveAcesso, 4)
    
    Call objImpressoraESC.ImprimeNormal(Mid(sSATChaveAcesso, 1, 11) & " " & Mid(sSATChaveAcesso, 12, 11) & " " & Mid(sSATChaveAcesso, 23, 11) & " " & Mid(sSATChaveAcesso, 34, 11))
    
    Call objImpressoraESC.ImprimeCodBarras(sSATChaveAcesso, "CODE128")
    
    Call objImpressoraESC.ImprimeQRCode(sSATChaveAcesso & "|" & sSATtimeStamp & "|" & objVenda.objCupomFiscal.sSATvalorTotalCFe & "|" & objVenda.objCupomFiscal.sSATCPFCNPJValue & "|" & objVenda.objCupomFiscal.sSATQRCode)
    
    Call objImpressoraESC.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", ""))
    Call objImpressoraESC.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", ""))
    Call objImpressoraESC.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", ""))
    Call objImpressoraESC.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", ""))
    Call objImpressoraESC.ImprimeNormal(Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", ""))
    
    Call objImpressoraESC.AcionarGuilhotina("P")
    
    Call objImpressoraESC.TerminaImpressao
        
    ImpressoraESC_ImprimirSAT = SUCESSO
    
    Exit Function
    
Erro_ImpressoraESC_ImprimirSAT:

    ImpressoraESC_ImprimirSAT = gErr

    Select Case gErr

        Case 99803, 99894 To 99898, 105868, 133088, 204319, 204324, 204329, 210470

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149439)

    End Select

    Exit Function

End Function



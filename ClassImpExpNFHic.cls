VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassImpExpNFHic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

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
            .dValorFrete & Chr(vbKeyEscape) & .dValorSeguro & Chr(vbKeyEscape) & .dValorOutrasDespesas & Chr(vbKeyEscape) & .dValorDesconto & Chr(vbKeyEscape) & .iCodTransportadora & Chr(vbKeyEscape) & .sMensagemNota & Chr(vbKeyEscape) & CStr(0) & Chr(vbKeyEscape) & .iTipoNFiscal & Chr(vbKeyEscape) & .sNaturezaOp & Chr(vbKeyEscape) & .dPesoLiq & Chr(vbKeyEscape) & .dPesoBruto & Chr(vbKeyEscape) & _
            .lNumIntTrib & Chr(vbKeyEscape) & .sPlaca & Chr(vbKeyEscape) & .sPlacaUF & Chr(vbKeyEscape) & .lVolumeQuant & Chr(vbKeyEscape) & .lVolumeEspecie & Chr(vbKeyEscape) & .lVolumeMarca & Chr(vbKeyEscape) & .sVolumeNumero & Chr(vbKeyEscape) & .iCanal & Chr(vbKeyEscape) & CStr(0) & Chr(vbKeyEscape) & .lClienteBenef & Chr(vbKeyEscape) & .iFilialCliBenef & Chr(vbKeyEscape) & .lFornecedorBenef & Chr(vbKeyEscape) & _
            .iFilialFornBenef & Chr(vbKeyEscape) & .iFreteRespons & Chr(vbKeyEscape) & CStr(0) & Chr(vbKeyEscape) & .sObservacao & Chr(vbKeyEscape) & .sCodUsuarioCancel & Chr(vbKeyEscape) & .sMotivoCancel & Chr(vbKeyEscape) & .dtHoraEntrada & Chr(vbKeyEscape) & .dtHoraSaida & Chr(vbKeyEscape) & .iCodTranspRedesp & Chr(vbKeyEscape) & .iDetPagFrete & Chr(vbKeyEscape) & .iSemDataSaida
        End With
               
        With objNFiscal.objTributacaoNF
            sRegistro = sRegistro & Chr(vbKeyEscape) & .sNaturezaOpInterna & Chr(vbKeyEscape) & .iTipoTributacao & Chr(vbKeyEscape) & .dIPIBase & Chr(vbKeyEscape) & .dIPIValor & Chr(vbKeyEscape) & .dIPICredito & Chr(vbKeyEscape) & .dICMSBase & Chr(vbKeyEscape) & .dICMSValor & Chr(vbKeyEscape) & .dICMSSubstBase & Chr(vbKeyEscape) & .dICMSSubstValor & Chr(vbKeyEscape) & .dICMSCredito & Chr(vbKeyEscape) & .iISSIncluso & Chr(vbKeyEscape) & .dISSBase & Chr(vbKeyEscape) & .dISSAliquota & Chr(vbKeyEscape) & .dISSValor & Chr(vbKeyEscape) & .dIRRFBase & Chr(vbKeyEscape) & .dIRRFAliquota & Chr(vbKeyEscape) & .dIRRFValor & Chr(vbKeyEscape) & .dValorINSS & Chr(vbKeyEscape) & .iINSSRetido & Chr(vbKeyEscape) & .dINSSBase & Chr(vbKeyEscape) & .dINSSDeducoes & Chr(vbKeyEscape) & .dPISCredito & Chr(vbKeyEscape) & .dCOFINSCredito & Chr(vbKeyEscape) & .dPISRetido & Chr(vbKeyEscape) & .dCOFINSRetido & Chr(vbKeyEscape) & .dCSLLRetido
        End With
        
        'Mantém os campos para manter a compatibilidade, porém retorna tudo zerado
        With objNFiscal.objTributacaoNF '.objTributacaoDesconto
             sRegistro = sRegistro & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & _
             0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & .lNumIntDoc & Chr(vbKeyEscape) & ""
        End With
        With objNFiscal.objTributacaoNF '.objTributacaoFrete
             sRegistro = sRegistro & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & _
             0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & .lNumIntDoc & Chr(vbKeyEscape) & ""
        End With
        With objNFiscal.objTributacaoNF '.objTributacaoOutras
             sRegistro = sRegistro & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & _
             0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & .lNumIntDoc & Chr(vbKeyEscape) & ""
        End With
        With objNFiscal.objTributacaoNF '.objTributacaoSeguro
             sRegistro = sRegistro & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & _
             0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & 0 & Chr(vbKeyEscape) & .lNumIntDoc & Chr(vbKeyEscape) & ""
        End With
        
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
               
        '??? hicare sem lctos
'        Set colLancamentos = colcolLancamentos(2 * iIndice - 1)
'        Set colLanPendente = colcolLancamentos(2 * iIndice)
'
'        For Each objLancto In colLancamentos
'            sRegistro = TIPO_REGISTRO_LANCAMENTO & Chr(vbKeyEscape)
'           sRegistro = sRegistro & objLancto.iFilialEmpresa & Chr(vbKeyEscape) & objLancto.sOrigem & Chr(vbKeyEscape) & objLancto.iExercicio & Chr(vbKeyEscape) & objLancto.iPeriodoLan & Chr(vbKeyEscape) & objLancto.lDoc & Chr(vbKeyEscape) & objLancto.iSeq & Chr(vbKeyEscape) & objLancto.iLote & Chr(vbKeyEscape) & objLancto.iPeriodoLote & Chr(vbKeyEscape) & objLancto.dtData & Chr(vbKeyEscape) & objLancto.sConta & Chr(vbKeyEscape) & objLancto.sCcl & Chr(vbKeyEscape) & objLancto.sHistorico & Chr(vbKeyEscape) & objLancto.dValor & Chr(vbKeyEscape) & objLancto.lNumIntDoc & Chr(vbKeyEscape) & objLancto.iFilialCliForn & Chr(vbKeyEscape) & objLancto.lCliForn & Chr(vbKeyEscape) & objLancto.iTransacao & Chr(vbKeyEscape) & objLancto.lDoc & Chr(vbKeyEscape) & objLancto.iAglutina & Chr(vbKeyEscape) & objLancto.lContaSimples & Chr(vbKeyEscape) & objLancto.iSeqContraPartida & Chr(vbKeyEscape) & objLancto.iEscaninho_Custo & Chr(vbKeyEscape) & objLancto.dValorLivroAuxiliar & Chr(vbKeyEscape) & _
'            objLancto.iClienteFornecedor & Chr(vbKeyEscape) & objLancto.sDocOrigem & Chr(vbKeyEscape) & objLancto.sProduto & Chr(vbKeyEscape) & objLancto.dQuantidade & Chr(vbKeyEscape) & objLancto.dtDataEstoque & Chr(vbKeyEscape) & objLancto.iApropriaCRProd & Chr(vbKeyEscape) & objLancto.iStatus & Chr(vbKeyEscape) & objLancto.sModelo & Chr(vbKeyEnd)
'            colRegistros.Add sRegistro
'        Next
'
'        For Each objLancto In colLanPendente
'            sRegistro = TIPO_REGISTRO_LANPENDENTE & Chr(vbKeyEscape)
'            sRegistro = sRegistro & objLancto.iFilialEmpresa & Chr(vbKeyEscape) & objLancto.sOrigem & Chr(vbKeyEscape) & objLancto.iExercicio & Chr(vbKeyEscape) & objLancto.iPeriodoLan & Chr(vbKeyEscape) & objLancto.lDoc & Chr(vbKeyEscape) & objLancto.iSeq & Chr(vbKeyEscape) & objLancto.iLote & Chr(vbKeyEscape) & objLancto.iPeriodoLote & Chr(vbKeyEscape) & objLancto.dtData & Chr(vbKeyEscape) & objLancto.sConta & Chr(vbKeyEscape) & objLancto.sCcl & Chr(vbKeyEscape) & objLancto.sHistorico & Chr(vbKeyEscape) & objLancto.dValor & Chr(vbKeyEscape) & objLancto.lNumIntDoc & Chr(vbKeyEscape) & objLancto.iFilialCliForn & Chr(vbKeyEscape) & objLancto.lCliForn & Chr(vbKeyEscape) & objLancto.iTransacao & Chr(vbKeyEscape) & objLancto.lDoc & Chr(vbKeyEscape) & objLancto.iAglutina & Chr(vbKeyEscape) & objLancto.lContaSimples & Chr(vbKeyEscape) & objLancto.iSeqContraPartida & Chr(vbKeyEscape) & objLancto.iEscaninho_Custo & Chr(vbKeyEscape) & objLancto.dValorLivroAuxiliar & Chr(vbKeyEscape) & _
'            objLancto.iClienteFornecedor & Chr(vbKeyEscape) & objLancto.sDocOrigem & Chr(vbKeyEscape) & objLancto.sProduto & Chr(vbKeyEscape) & objLancto.dQuantidade & Chr(vbKeyEscape) & objLancto.dtDataEstoque & Chr(vbKeyEscape) & objLancto.iApropriaCRProd & Chr(vbKeyEscape) & objLancto.iStatus & Chr(vbKeyEscape) & objLancto.sModelo & Chr(vbKeyEnd)
'            colRegistros.Add sRegistro
'        Next
        
        For Each objItemNF In objNFiscal.ColItensNF
        
            With objItemNF
                sRegistro = TIPO_REGISTRO_ITEMNOTAFISCAL & Chr(vbKeyEscape) & .iItem & Chr(vbKeyEscape) & .sProduto & Chr(vbKeyEscape) & .sUnidadeMed & Chr(vbKeyEscape) & .dQuantidade & Chr(vbKeyEscape) & .dPrecoUnitario & Chr(vbKeyEscape) & .dPercDesc & Chr(vbKeyEscape) & .dValorDesconto & Chr(vbKeyEscape) & .dtDataEntrega & Chr(vbKeyEscape) & .sDescricaoItem & Chr(vbKeyEscape) & .dValorAbatComissao & Chr(vbKeyEscape) & CStr(0) & Chr(vbKeyEscape) & CStr(0) & Chr(vbKeyEscape) & .lNumIntDoc & Chr(vbKeyEscape) & .lNumIntTrib & Chr(vbKeyEscape) & CStr(0) & Chr(vbKeyEscape) & .sCcl & Chr(vbKeyEscape) & CStr(0)
            End With
            
            With objItemNF.objTributacaoItemNF
                sRegistro = sRegistro & Chr(vbKeyEscape) & .iItem & Chr(vbKeyEscape) & .lNumIntDoc & Chr(vbKeyEscape) & .sNaturezaOp & Chr(vbKeyEscape) & .iTipoTributacao & Chr(vbKeyEscape) & .iIPITipo & Chr(vbKeyEscape) & .sIPICodProduto & Chr(vbKeyEscape) & .dIPIBaseCalculo & Chr(vbKeyEscape) & .dIPIPercRedBase & Chr(vbKeyEscape) & .dIPIAliquota & Chr(vbKeyEscape) & .dIPIValor & Chr(vbKeyEscape) & .dIPICredito & Chr(vbKeyEscape) & .iICMSTipo & Chr(vbKeyEscape) & .dICMSBase & Chr(vbKeyEscape) & .dICMSPercRedBase & Chr(vbKeyEscape) & .dICMSAliquota & Chr(vbKeyEscape) & .dICMSValor & Chr(vbKeyEscape) & .dICMSCredito & Chr(vbKeyEscape) & .dICMSSubstBase & Chr(vbKeyEscape) & .dICMSSubstAliquota & Chr(vbKeyEscape) & .dICMSSubstValor & Chr(vbKeyEscape) & .dPISCredito & Chr(vbKeyEscape) & .dCOFINSCredito & Chr(vbKeyEnd)
            End With
            colRegistros.Add sRegistro
            
            '??? hicare sem controle de estoque
'            For Each objAlocacaoNF In objItemNF.ColAlocacoes
'                sRegistro = TIPO_REGISTRO_ALOCACAOITEMNF & Chr(vbKeyEscape) & objAlocacaoNF.dQuantidade & Chr(vbKeyEscape) & objAlocacaoNF.iAlmoxarifado & Chr(vbKeyEscape) & objAlocacaoNF.iTransferencia & Chr(vbKeyEscape) & objAlocacaoNF.sAlmoxarifado & Chr(vbKeyEscape) & objAlocacaoNF.sUnidadeMed & Chr(vbKeyEnd)
'                colRegistros.Add sRegistro
'            Next
            
            For Each objItemNFGrade In objItemNF.colItensRomaneioGrade
                With objItemNFGrade
                    sRegistro = TIPO_REGISTRO_ITEMNFGRADE & Chr(vbKeyEscape) & .iControleEstoque & Chr(vbKeyEscape) & .lNumIntDoc & Chr(vbKeyEscape) & .sProduto & Chr(vbKeyEscape) & .dQuantidade & Chr(vbKeyEscape) & .sDescricao & Chr(vbKeyEscape) & .sUMEstoque & Chr(vbKeyEscape) & CStr(0) & Chr(vbKeyEscape) & CStr(0) & Chr(vbKeyEnd)
                End With
                colRegistros.Add sRegistro
                
                '??? hicare sem controle de estoque
'                For Each objLocalizacaoItemGrade In objItemNFGrade.colLocalizacao
'                    sRegistro = TIPO_REGISTRO_LOCALIZACAOITEMNFGRADE & Chr(vbKeyEscape) & objLocalizacaoItemGrade.dQuantidade & Chr(vbKeyEscape) & objLocalizacaoItemGrade.dtDataValidade & Chr(vbKeyEscape) & objLocalizacaoItemGrade.iAlmoxarifado & Chr(vbKeyEscape) & objLocalizacaoItemGrade.iFilialEmpresa & Chr(vbKeyEscape) & objLocalizacaoItemGrade.sAlmoxarifado & Chr(vbKeyEscape) & objLocalizacaoItemGrade.sResponsavel & Chr(vbKeyEnd)
'                    colRegistros.Add sRegistro
'                Next
                            
            Next
        
        Next
        
    Next

    Monta_Registros_NotasFiscais = SUCESSO

    Exit Function

Erro_Monta_Registros_NotasFiscais:

    Monta_Registros_NotasFiscais = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179196)
    
    End Select
    
    Exit Function
    
End Function

Function Monta_Registros_Produtos(colRegistros As Collection, colProdutos As Collection) As Long

Dim objProduto As ClassProduto
Dim objProdutoCateg As ClassProdutoCategoria
Dim iIndice As Integer
Dim sRegistro As String
'???hicare: sem controle de estoque

On Error GoTo Erro_Monta_Registros_Produtos

    iIndice = 0

    For Each objProduto In colProdutos
        
        iIndice = iIndice + 1
        With objProduto
        
            sRegistro = TIPO_REGISTRO_PRODUTO & Chr(vbKeyEscape) & .sCodigo & Chr(vbKeyEscape) & .sDescricao & Chr(vbKeyEscape) & .iGerencial & Chr(vbKeyEscape) & .iAtivo & Chr(vbKeyEscape) & .iFaturamento & Chr(vbKeyEscape) & .iClasseUM & Chr(vbKeyEscape) & .sSiglaUMEstoque & Chr(vbKeyEscape) & _
            .sSiglaUMCompra & Chr(vbKeyEscape) & .sSiglaUMVenda & Chr(vbKeyEscape) & .sSubstituto1 & Chr(vbKeyEscape) & .sSubstituto2 & Chr(vbKeyEscape) & .iNivel & Chr(vbKeyEscape) & .iTipo & Chr(vbKeyEscape) & .sNomeReduzido & Chr(vbKeyEscape) & .sModelo & Chr(vbKeyEscape) & .iPrazoValidade & Chr(vbKeyEscape) & .sCodigoBarras & Chr(vbKeyEscape) & _
            .iEtiquetasCodBarras & Chr(vbKeyEscape) & .dPesoLiq & Chr(vbKeyEscape) & .dPesoBruto & Chr(vbKeyEscape) & .dComprimento & Chr(vbKeyEscape) & .dEspessura & Chr(vbKeyEscape) & .dLargura & Chr(vbKeyEscape) & .sCor & Chr(vbKeyEscape) & .sObsFisica & Chr(vbKeyEscape) & .iCompras & Chr(vbKeyEscape) & .iPCP & Chr(vbKeyEscape) & .iKitBasico & Chr(vbKeyEscape) & _
            .iKitInt & Chr(vbKeyEscape) & .dIPIAliquota & Chr(vbKeyEscape) & .sIPICodigo & Chr(vbKeyEscape) & .sIPICodDIPI & Chr(vbKeyEscape) & PRODUTO_CONTROLE_SEM_ESTOQUE & Chr(vbKeyEscape) & .iApropriacaoCusto & Chr(vbKeyEscape) & .sContaContabil & Chr(vbKeyEscape) & .sContaContabilProducao & Chr(vbKeyEscape) & .dResiduo & Chr(vbKeyEscape) & _
            .iNatureza & Chr(vbKeyEscape) & .dCustoReposicao & Chr(vbKeyEscape) & .iOrigemMercadoria & Chr(vbKeyEscape) & CStr(0) & Chr(vbKeyEscape) & .dPercentMaisQuantCotAnt & Chr(vbKeyEscape) & .dPercentMenosQuantCotAnt & Chr(vbKeyEscape) & .iConsideraQuantCotAnt & Chr(vbKeyEscape) & _
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179197)
    
    End Select
    
    Exit Function
    
End Function

Function Monta_Registros_Clientes(colRegistros As Collection, colClientes As Collection, colcolFiliaisCli As Collection, colcolEnderecos As Collection) As Long
'hicare: elimina informacao do vendedor, regiao de venda

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
            sRegistro = TIPO_REGISTRO_CLIENTE & Chr(vbKeyEscape) & .lCodigo & Chr(vbKeyEscape) & .sRazaoSocial & Chr(vbKeyEscape) & .sNomeReduzido & Chr(vbKeyEscape) & .iTipo & Chr(vbKeyEscape) & .sObservacao & Chr(vbKeyEscape) & .dLimiteCredito & Chr(vbKeyEscape) & .iCondicaoPagto & Chr(vbKeyEscape) & .dDesconto & Chr(vbKeyEscape) & CStr(0) & Chr(vbKeyEscape) & CStr(0) & Chr(vbKeyEscape) & .iProxCodFilial & Chr(vbKeyEscape) & CStr(0) & Chr(vbKeyEscape) & .iAtivo & Chr(vbKeyEscape) & .iTransferido & Chr(vbKeyEscape) & _
            .sCgc & Chr(vbKeyEscape) & .sInscricaoEstadual & Chr(vbKeyEscape) & .sInscricaoMunicipal & Chr(vbKeyEscape) & .sInscricaoSuframa & Chr(vbKeyEscape) & .lEndereco & Chr(vbKeyEscape) & .lEnderecoEntrega & Chr(vbKeyEscape) & .lEnderecoCobranca & Chr(vbKeyEscape) & .iCodTransportadora & Chr(vbKeyEscape) & .sObservacao2 & Chr(vbKeyEscape) & .sContaContabil & Chr(vbKeyEscape) & CStr(0) & Chr(vbKeyEscape) & .dComissaoVendas & Chr(vbKeyEscape) & CStr(0) & Chr(vbKeyEscape) & .iFreqVisitas & Chr(vbKeyEscape) & .dtDataUltVisita & Chr(vbKeyEscape) & CStr(0) & Chr(vbKeyEscape) & .iTipoFrete & Chr(vbKeyEscape) & .sRG & Chr(vbKeyEscape) & .iCodTranspRedesp & Chr(vbKeyEscape) & .iDetPagFrete & Chr(vbKeyEscape) & .sGuia & _
            Chr(vbKeyEscape) & objEndereco.sEndereco & Chr(vbKeyEscape) & objEndereco.sBairro & Chr(vbKeyEscape) & objEndereco.sCidade & Chr(vbKeyEscape) & objEndereco.sSiglaEstado & Chr(vbKeyEscape) & objEndereco.iCodigoPais & Chr(vbKeyEscape) & objEndereco.sCEP & Chr(vbKeyEscape) & objEndereco.sTelefone1 & Chr(vbKeyEscape) & objEndereco.sTelefone2 & Chr(vbKeyEscape) & objEndereco.sEmail & Chr(vbKeyEscape) & objEndereco.sFax & Chr(vbKeyEscape) & objEndereco.sContato & _
            Chr(vbKeyEscape) & objEnderecoCobranca.sEndereco & Chr(vbKeyEscape) & objEnderecoCobranca.sBairro & Chr(vbKeyEscape) & objEnderecoCobranca.sCidade & Chr(vbKeyEscape) & objEnderecoCobranca.sSiglaEstado & Chr(vbKeyEscape) & objEnderecoCobranca.iCodigoPais & Chr(vbKeyEscape) & objEnderecoCobranca.sCEP & Chr(vbKeyEscape) & objEnderecoCobranca.sTelefone1 & Chr(vbKeyEscape) & objEnderecoCobranca.sTelefone2 & Chr(vbKeyEscape) & objEnderecoCobranca.sEmail & Chr(vbKeyEscape) & objEnderecoCobranca.sFax & Chr(vbKeyEscape) & objEnderecoCobranca.sContato & _
            Chr(vbKeyEscape) & objEnderecoEntrega.sEndereco & Chr(vbKeyEscape) & objEnderecoEntrega.sBairro & Chr(vbKeyEscape) & objEnderecoEntrega.sCidade & Chr(vbKeyEscape) & objEnderecoEntrega.sSiglaEstado & Chr(vbKeyEscape) & objEnderecoEntrega.iCodigoPais & Chr(vbKeyEscape) & objEnderecoEntrega.sCEP & Chr(vbKeyEscape) & objEnderecoEntrega.sTelefone1 & Chr(vbKeyEscape) & objEnderecoEntrega.sTelefone2 & Chr(vbKeyEscape) & objEnderecoEntrega.sEmail & Chr(vbKeyEscape) & objEnderecoEntrega.sFax & Chr(vbKeyEscape) & objEnderecoEntrega.sContato & Chr(vbKeyEnd)
        End With
        colRegistros.Add sRegistro
        
        For Each objFilialCli In colcolFiliaisCli(iIndice)
            With objFilialCli
                sRegistro = TIPO_REGISTRO_FILIALCLIENTE & Chr(vbKeyEscape) & .lCodCliente & Chr(vbKeyEscape) & .iCodFilial & Chr(vbKeyEscape) & .sNome & Chr(vbKeyEscape) & .sCgc & Chr(vbKeyEscape) & .sInscricaoEstadual & Chr(vbKeyEscape) & .sInscricaoMunicipal & Chr(vbKeyEscape) & .sInscricaoSuframa & Chr(vbKeyEscape) & .lEndereco & Chr(vbKeyEscape) & .lEnderecoEntrega & Chr(vbKeyEscape) & _
                .lEnderecoCobranca & Chr(vbKeyEscape) & .iCodTransportadora & Chr(vbKeyEscape) & .sObservacao & Chr(vbKeyEscape) & .sContaContabil & Chr(vbKeyEscape) & CStr(0) & Chr(vbKeyEscape) & .dComissaoVendas & Chr(vbKeyEscape) & CStr(0) & Chr(vbKeyEscape) & .iFreqVisitas & Chr(vbKeyEscape) & .dtDataUltVisita & Chr(vbKeyEscape) & CStr(0) & Chr(vbKeyEscape) & .iTipoFrete & Chr(vbKeyEscape) & .sRG & Chr(vbKeyEscape) & _
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179198)
    
    End Select
    
    Exit Function
    
End Function



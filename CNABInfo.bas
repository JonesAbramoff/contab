Attribute VB_Name = "CNABInfo"
Option Explicit

'====================== ITAÚ =========================================

'------------ Remessa de Títulos a Receber ----------------------------

'Registro de Header do arquivo de remessa titulos a receber
Public Type typeRemHeaderItau
    sIdentificacaoReg As String
    sIdentificacaoArq As String
    sIdentificacaoMov As String
    sCodTipoServico As String
    sIdentificacaoTipoServ As String
    sNomeEmpresa As String
    sDataEmissaoArq As String
    sSequencialRegistro As String
    sCodEmpresa As String
    sComplemento1 As String
    sNumeroBanco As String
    sNomeBanco As String
    sComplemento2 As String
End Type

'Registro de Detalhe do arquivo de remessa titulos a receber
Public Type typeRemDetalheItau
    sIdentificacaoReg As String
    sTipoInscricaoEmpSac As String
    sNumInscricaoEmpSac As String
    sIdentificacaoEmp As String
    sComplemento1 As String
    sUsoDaEmpresa As String
    sNossoNumero As String
    sQuantMoeda As String
    sNumCarteiraBanco As String
    sUsoDoBanco As String
    sCodCarteira As String
    sCodOcorrencia As String
    sSeuNumero As String
    sVencimento As String
    sValorTitulo As String
    sNumeroBanco As String
    sAgenciaCobradora As String
    sEspecie As String
    sAceite As String
    sDataEmissao As String
    sInstrucao1 As String
    sInstrucao2 As String
    sJurosDiarios As String
    sDescontoAte As String
    sValorDesconto As String
    sValorIOF As String
    sAbatimento As String
    sCodigoInscSacado As String
    sNumInscricaoSacado As String
    sNome As String
    sComplemento2 As String
    sLogradouro As String
    sBairro As String
    sCEP As String
    sCidade As String
    sEstado As String
    sSacadorAvalista As String
    sComplemento3 As String
    sDataDeMora As String
    sPrazo As String
    sComplemento4 As String
    sNumSequencial As String
End Type

'Registro de Trailer do arquivo de remessa titulos a receber
Type typeRemTrailerItau
    sIdentificacaoReg As String
    sComplemento1 As String
    sSequencialRegistro As String
End Type

'------------------ Retorno de Títulos a Receber -------------------------

'Registro de Header do arquivo de retorno titulos a receber
Public Type typeRetHeaderItau
    sIdentificacaoReg As String
    sIdentificacaoArq As String
    sIdentificacaoMov As String
    sCodTipoServico As String
    sIdentificacaoTipoServ As String
    sNomeEmpresa As String
    sDataEmissaoArq As String
    sSequencialRegistro As String
    sCodEmpresa As String
    sComplemento1 As String
    sNumeroBanco As String
    sNomeBanco As String
    sComplemento2 As String
    sDensidade As String
    sUnidadeDensid As String
    sNumSequencialArqRet As String
    sDataCredito As String
End Type

'Registro de detalhe do arquivo de retorno de titulos a receber
Public Type typeRetDetalheItau
    sIdentificacaoReg As String
    sTipoInscricaoEmpresa As String
    sNumInscricaoEmpresa As String
    sCodEmpresa As String
    sComplemento1 As String
    sUsoDaEmpresa As String
    sNossoNumero1 As String
    sComplemento2 As String
    sNumCarteira As String
    sNossoNumero2 As String
    sDACNossoNumero2 As String
    sCodCarteira As String
    sCodOcorrencia As String
    sDataOcorrencia As String
    sSeuNumero As String
    sNossoNumero3 As String
    sComplemento3 As String
    sVencimento As String
    sValorTitulo As String
    sNumeroBanco As String
    sAgenciaCobradora As String
    sDACAgCobradora As String
    sEspecie As String
    sTarifaCobranca As String
    sComplemento4 As String
    sValorIOF As String
    sAbatimento As String
    sDescontos As String
    sValorPrincipal As String
    sJuros As String
    sLogradouro As String
    sComplemento5 As String
    sDataCredito As String
    sComplemento6 As String
    sNomeSacado As String
    sComplemento7 As String
    sErros As String
    sComplemento8 As String
    sComplemento9 As String
    sNumSequencial As String
    sValorEntregue As String
End Type

'Registro de trailer do arquivo de retorno titulos a receber
Public Type typeRetTrailerItau
    sIdentificacaoReg As String
    sCodigoRetorno As String
    sCodigoServico As String
    sNumeroBanco As String
    sComplemento1 As String
    sQuantTitulosCobrSimp As String
    sValorTotalSimp As String
    sAvisoBancario As String
    sComplemento2 As String
    sQuantTitulosCobrVinc As String
    sValorTotalVinc As String
    sAvisoBancario2 As String
    sComplemento3 As String
    sQuantTitulosCobrEscr As String
    sValorTotalEscr As String
    sAvisoBancario3 As String
    sControleArquivo As String
    sQuantDetalhes As String
    sValorTotalInformado As String
    sComplemento4 As String
    sSequencialRegistro As String
    
End Type

'-------------------- Remessa de Contas a Pagar -----------------------------

'Registro de Header de Arquivo de Remessa/Retorno de títulos a pagar
Public Type typePagtoHeaderArqItau
    sNumeroBanco As String
    sCodigoLote As String
    sTipoRegistro As String
    sComplemento1 As String
    sLayoutArq As String
    sTipoInscEmpresa As String
    sCGCEmpresaDebitada As String
    sComplemento2 As String
    sAgencia As String
    sComplemento3 As String
    sConta As String
    sComplemento4 As String
    sDACAgDebitada As String
    sNomeEmpresa As String
    sNomeBanco As String
    sComplemento5 As String
    sArquivoCodigo As String
    sDataGeracao As String
    sHoraGeracao As String
    sZeros As String
    sUnidadeDensidade As String
    sComplemento6 As String
       
End Type

Public Type typePagtoHeaderLote
    sNumeroBanco As String
    sCodigoLote As String
    sTipoRegistro As String
    sTipoOperacao As String
    sTipoPagamento As String
    sFormaPagamento As String
    sLayoutLote As String
    sComplemento1 As String
    sTipoInscEmpresa As String
    sCGCEmpresaDebitada As String
    sComplemento2 As String
    sAgencia As String
    sComplemento3 As String
    sContaDebitada As String
    sComplemento4 As String
    sDACAgDebitada As String
    sNomeEmpresa As String
    sFinalidadeLote As String
    sHistorico As String
    sEnderecoEmpresa As String
    sNumeroLocal As String
    sComplementoEndereco As String
    sCidade As String
    sCEP As String
    sEstado As String
    sComplemento5 As String
    sOcorrenciasRetorno As String
End Type

Public Type typePagtoSegmentoA
    sNumeroBanco As String
    sCodigoLote As String
    sTipoRegistro As String
    sNumeroRegistro As String
    sSegmento As String
    sTipoMovimento As String
    sZeros1 As String
    sBancoFavorecido As String
    sAgenciaConta As String
    sNomeFavorecido As String
    sSeuNumero As String
    sDataPagamento As String
    sTipoMoeda As String
    sZeros2 As String
    sValorPagamento As String
    sNossoNumero As String
    sComplemento1 As String
    sDataEfetiva As String
    sValorEfetivo As String
    sFinalidadeDetalhe As String
    sComplemento2 As String
    sAviso As String
    sOcorrenciasRetorno As String
End Type

Public Type typePagtoSegmentoB
    sNumeroBanco As String
    sCodigoLote As String
    sTipoRegistro As String
    sNumeroRegistro As String
    sSegmento As String
    sComplemento1 As String
    sTipoInscEmpresa As String
    sNumInscFavorecido As String
    sEndereco As String
    sNumeroLocal As String
    sComplementoEndereco As String
    sBairro As String
    sCidade As String
    sCEP As String
    sEstado As String
    sComplemento2 As String
    
End Type
    
Public Type typePagtoTrailerLoteItau
    sNumeroBanco As String
    sCodigoLote As String
    sTipoRegistro As String
    sComplemento1 As String
    sTotalQtdRegistros As String
    sTotalValorPagtos As String
    sZeros As String
    sComplemento2 As String
    sOcorrenciasRetorno As String
End Type

Public Type typePagtoSegmentoJ
    sNumeroBanco As String
    sCodigoLote As String
    sTipoRegistro As String
    sNumeroRegistro As String
    sSegmento As String
    sTipoMovimento As String
    sBancoFavorecido As String
    sMoeda As String
    sDV As String
    sValor As String
    sCampoLivre As String
    sNomeFavorecido As String
    sDataVencto As String
    sValorTitulo As String
    sDescontos As String
    sAcrescimos As String
    sDataPagamento As String
    sValorPagamento As String
    sZeros As String
    sSeuNumero As String
    sComplemento As String
    sNossoNumero As String
    sOcorrenciasRetorno As String
    
End Type
    

Public Type typePagtoTrailerArqItau
    sNumeroBanco As String
    sCodigoLote As String
    sTipoRegistro As String
    sComplemento1 As String
    sTotalQtdLotes As String
    sTotalQtdRegistros As String
    sComplemento2 As String
End Type

'============================ BRADESCO ===================================

'------------------ Remessa de Titulos a Receber --------------------

'Registro de Header do arquivo de remessa titulos a receber
Public Type typeRemHeaderBradesco
    sIdentificacaoReg As String
    sIdentificacaoArq As String
    sLiteralRemessa As String
    sCodTipoServico As String
    sLiteralServico As String
    sCodigoEmpresa As String
    sNomeEmpresa As String
    sNumeroBanco As String
    sNomeBanco As String
    sDataEmissaoArq As String
    sComplemento1 As String
    sIdentificacaoSistema As String
    sSequencialArq As String
    sComplemento2 As String
    sSequencialRegistro As String
End Type

'Registro de Detalhe do arquivo de remessa titulos a receber
Public Type typeRemDetalheBradesco
    sIdentificacaoReg As String
    sAgenciaDebito As String
    sDigitoAgDebito As String
    sRazaoContaCorrente As String
    sContaCorrente As String
    sDigitoContaCorrente As String
    sIdentifEmpresa As String
    sNumControle As String
    sCodigoBancoDebitado As String
    sZeros As String
    sNossoNumero As String
    sDesconto As String
    sCondEmissaoPapel As String
    sIdentEmitePapel As String
    sIdentOpBanco As String
    sIndicadorRateio As String
    sEnderecamento As String
    sBranco As String
    sIdentificacaoOcorr As String
    sNumDocto As String
    sDataVencto As String
    sValorTitulo As String
    sBancoCobranca As String
    sAgenciaDepositaria As String
    sEspecie As String
    sAceite As String
    sDataEmissao As String
    sInstrucao1 As String
    sInstrucao2 As String
    sValorJuros As String
    sDataDesconto As String
    sValorDesconto As String
    sValorIOF As String
    sAbatimento As String
    sIdentfInscSacado As String
    sNumInscricaoSacado As String
    sNomeSacado As String
    sEnderecoSacado As String
    sMensagem1 As String
    sCEP As String
    sCEPSufixo As String
    sEstado As String
    sSacadorAvalista As String
    sNumSequencialRegistro As String

End Type

'Registro de Trailer do arquivo de remessa titulos a receber
Type typeRemTrailerBradesco
    sIdentificacaoReg As String
    sComplemento1 As String
    sSequencialRegistro As String
End Type

'------------------ Retorno de Titulos a Receber --------------------

'Registro de Header do arquivo de retorno titulos a receber
Public Type typeRetHeaderBradesco
    sIdentificacaoReg As String
    sIdentificacaoArq As String
    sLiteralRetorno As String
    sCodTipoServico As String
    sLiteralServico As String
    sCodEmpresa As String
    sNomeEmpresa As String
    sNumeroBanco As String
    sNomeBanco As String
    sDataEmissaoArq As String
    sZeros As String
    sNumAviso As String
    sComplemento1 As String
    sDataCredito As String
    sComplemento2 As String
    sSequencialRegistro As String
End Type

'Registro de detalhe do arquivo de retorno de titulos a receber
Public Type typeRetDetalheBradesco
    sIdentificacaoReg As String
    sTipoInscricaoEmpresa As String
    sNumInscricaoEmpresa As String
    sZeros1 As String
    sIdentfEmpresa As String
    sNumControle As String
    sZeros2 As String
    sNossoNumero1 As String
    sUsoDoBanco1 As String
    sUsoDoBanco2 As String
    sIndicadorRateio As String
    sZeros As String
    sCodCarteira As String
    sCodOcorrencia As String
    sDataOcorrencia As String
    sNumDocto As String
    sNossoNumero2 As String
    sVencimento As String
    sValorTitulo As String
    sBancoCobrador As String
    sAgenciaCobradora As String
    sEspecie As String
    sTarifaCobranca As String
    sOutrasDespesas As String
    sJurosOp As String
    sValorIOF As String
    sAbatimento As String
    sDescontos As String
    sValorPrincipal As String
    sJuros As String
    sOutrosCreditos As String
    sComplemento1 As String
    sMotivoOcorr1 As String
    sDataCredito As String
    sComplemento2 As String
    sMotivoOcorr2 As String
    sComplemento3 As String
    sNumSequencialRegistro As String
End Type

'Registro de trailer do arquivo de retorno titulos a receber
Public Type typeRetTrailerBradesco
    sIdentificacaoReg As String
    sCodigoRetorno As String
    sCodigoTipoRegistro As String
    sNumeroBanco As String
    sComplemento1 As String
    sQuantTitulosCobr As String
    sValorTotalCobr As String
    sAvisoBancario As String
    sComplemento2 As String
    sQuantRegOcorr_02 As String
    sValorRegOcorr_02 As String
    sValorRegOcorr_06 As String
    sQuantRegOcorr_06 As String
    sValorRegOcorr_06_09_10 As String
    sQuantRegOcorr_09_10 As String
    sValorRegOcorr_09_10 As String
    sQuantRegOcorr_13 As String
    sValorRegOcorr_13 As String
    sQuantRegOcorr_14 As String
    sValorRegOcorr_14 As String
    sQuantRegOcorr_12 As String
    sValorRegOcorr_12 As String
    sQuantRegOcorr_19 As String
    sValorRegOcorr_19 As String
    sComplemento3 As String
    sValorTotalRateios As String
    sQuantTotalRateios As String
    sSequencialRegistro As String
    
End Type

'-------------------- Remessa de Contas a Pagar -----------------------------

'Registro de Header de Arquivo de Remessa/Retorno de títulos a pagar
Public Type typePagtoHeaderArqBradesco
    sTipoRegistro As String
    sCodComunicacao As String
    sTipoInscEmpresa As String
    sCGCEmpresaDebitada As String
    sNomeEmpresa As String
    sTipoServico As String
    sCodigoOrigem As String
    sNumRemessa As String
    sNumRetorno As String
    sDataGeracao As String
    sHoraGeracao As String
    sDensidade As String
    sUnidadeDensidade As String
    sIdentifModulo As String
    sTipoProcessamento As String
    sReservadoEmpresa As String
    sReservadoBanco As String
    sReservadoExpansao As String
    sNumSequencialRegistro As String
End Type

Public Type typePagtoDetalheBradesco
    sIdentificacao As String
    sTipoInscrForn As String
    sInscricaoForn As String
    sNomeFornecedor As String
    sEnderecoForn As String
    sCepForn As String
    sCepComplementoForn As String
    sBancoForn As String
    sAgenciaForn As String
    sDVAgenciaForn As String
    sContaCorrenteForn As String
    sDVContaCorrenteForn As String
    sNumeroPagto As String
    sCarteira As String
    sAnoNossoNumero As String
    sNossoNumero As String
    sSeuNumero As String
    sDataVencimento As String
    sDataEmissao As String
    sDataDesconto As String
    sValorDocto As String
    sValorPagto As String
    sValorDesconto As String
    sValorAcrescimo As String
    sTipoDocumento As String
    sNumeroTitulo As String
    sSerieDocumento As String
    sModalidadePagto As String
    sDataPagto As String
    sMoeda As String
    sSituacaoAgendamento As String
    sInformacaoRetorno As String
    sTipoMovimento As String
    sCodigoMovimento As String
    sEnderecoSacado As String
    sSacadorAvalista As String
    sReserva1 As String
    sNivelInformacao As String
    sInformacoesCompl As String
    sCodigoAreaEmpresa As String
    sUsoDaEmpresa As String
    sReserva2 As String
    sCodLancamento As String
    sReserva3 As String
    sTipoContaForn As String
    sContaComplementar As String
    sReserva4 As String
    sNumSequencialRegistro As String
End Type

Public Type typePagtoTrailerArqBradesco
    sIdentificacao As String
    sTotalQtdRegistros As String
    sTotalValorPagtos As String
    sReserva As String
    sNumSequencialRegistro As String
End Type

'============================ UNIBANCO ===================================

'------------------ Remessa de Titulos a Receber --------------------

'Registro de Header do arquivo de remessa titulos a receber
Public Type typeRemHeaderUnibanco
    sIdentificacaoReg As String
    sIdentificacaoArq As String
    sLiteralRemessa As String
    sCodTipoServico As String
    sLiteralServico As String
    sBrancos1 As String
    sAgCredito As String
    sContaCredito As String
    sDVContaCredito As String
    sZeros1 As String
    sGrupoEmpresarial1 As String
    sNomeEmpresa As String
    sTipoFormulario As String
    sTipoCritica As String
    sTipoPostagem As String
    sGrupoEmpresarial2 As String
    sNumeroBanco As String
    sNomeBanco As String
    sBrancos2 As String
    sDataEmissaoArq As String
    sDensidade As String
    sLiteralDensidade As String
    sZeros2 As String
    sVersaoArquivo As String
    sSequencialRegistro As String
End Type

'Registro de Detalhe do arquivo de remessa titulos a receber
Public Type typeRemDetalheUnibanco
    sIdentificacaoReg As String
    sNumReferenciaCli As String
    sDVReferencia As String
    sAgenciaDepositaria As String
    sCodigoInscricao As String
    sNumeroInscricao As String
    sAgenciaDebito As String
    sContaCorrente As String
    sDigitoContaCorrente As String
    sZeros1 As String
    sUsoEmpresa As String
    sNossoNumero As String
    sDVNossoNumero As String
    sMensagem As String
    sMoeda As String
    sCarteira As String
    sTipoTransacao As String
    sSeuNumero As String
    sDataVencto As String
    sValorTitulo As String
    sBancoCobranca As String
    sZeros2 As String
    sAgenciaCobradora As String
    sEspecie As String
    sAceite As String
    sDataEmissao As String
    sInstrucao1 As String
    sInstrucao2 As String
    sValorJuros As String
    sDataDesconto As String
    sValorDesconto As String
    sZeros3 As String
    sAbatimento As String
    sIdentfInscSacado As String
    sNumInscricaoSacado As String
    sNomeSacado As String
    sZeros4 As String
    sEnderecoSacado As String
    sEnderecoCompl As String
    sBairro As String
    sCEP As String
    sCidade As String
    sEstado As String
    sSacadorAvalista As String
    sZeros5 As String
    sPrazoProtesto As String
    sZeros6 As String
    sNumSequencialRegistro As String
    sBrancos1 As String
    sQuantMoedas As String
    sNumParcela As String
    sBrancos2 As String
    sNumTitulo As String
    sIndicadorMsg As String
    sDataMulta As String
    sPrazoMora As String
    sBrancos3 As String
    sEndereco2 As String
    sBrancos4 As String
    sValorMulta As String
    sDataProcessamento As String
End Type

'Registro de Trailer do arquivo de remessa titulos a receber
Type typeRemTrailerUnibanco
    sIdentificacaoReg As String
    sZeros As String
    sSequencialArq As String
    sSequencialRegistro As String
    sBrancos1 As String
    sAgCtaDVCedente As String
    sBrancos2 As String
    sQuantRegistros As String
    sTotalTitulos As String
End Type


'------------------ Retorno de Titulos a Receber --------------------

'Registro de Header do arquivo de retorno titulos a receber
Public Type typeRetHeaderUnibanco
    sIdentificacaoReg As String
    sIdentificacaoArq As String
    sLiteralRetorno As String
    sCodTipoServico As String
    sLiteralServico As String
    sBrancos1 As String
    sCodEmpresa As String
    sZeros1 As String
    sNomeEmpresa As String
    sNumeroBanco As String
    sNomeBanco As String
    sBrancos2 As String
    sDataEmissaoArq As String
    sDensidade As String
    sLiteralDensidade As String
    sBrancos3 As String
    sDataMovto As String
    sComplemento1 As String
    sZeros As String
    sMoeda As String
    sComplemento2 As String
    sBrancos4 As String
    sNumGeracaoArq As String
    sSequencialRegistro As String
    sCodOperacao As String
End Type

'Registro de detalhe do arquivo de retorno de titulos a receber
Public Type typeRetDetalheUnibanco
    sIdentificacaoReg As String
    sTipoInscricaoEmpresa As String
    sNumInscricaoEmpresa As String
    sCodEmpresa As String
    sZeros1 As String
    sUsoDaEmpresa As String
    sNossoNumero1 As String
    sZeros2 As String
    sIdenfOpBanco As String
    sCarteira As String
    sCodOcorrencia As String
    sDataOcorrencia As String
    sSeuNumero As String
    sNossoNumero2 As String
    sZeros3 As String
    sVencimento As String
    sValorTitulo As String
    sBancoCobrador As String
    sAgenciaCobradora As String
    sEspecie As String
    sTarifaCobranca As String
    sOutrasDespesas As String
    sZeros4 As String
    sAbatimento As String
    sDescontos As String
    sValorPrincipal As String
    sJuros As String
    sZeros5 As String
    sValorOriginalTitulo As String
    sNomeSacado As String
    sMoeda As String
    sZeros6 As String
    sDataEmissaoArq As String
    sNumGeracaoArq As String
    sNumSequencialRegistro As String
End Type

'Registro de trailer do arquivo de retorno titulos a receber
Public Type typeRetTrailerUnibanco
    sIdentificacaoReg As String
    sQuantTitulos As String
    sValorTotal As String
    sAvisoBancario As String
    sZeros As String
    sNumGeracaoArq As String
    sSequencialRegistro As String
    
End Type

'-------------------- Remessa de Contas a Pagar -----------------------------

'Registro de Header de Arquivo de Remessa/Retorno de títulos a pagar
Public Type typePagtoHeaderArqUnibanco
    sIdentificacaoReg As String
    sIdentificacaoArq As String
    sLiteralRemessa As String
    sCodTipoServico As String
    sLiteralServico As String
    sAgenciaCliente As String
    sContaCliente As String
    sDVCliente As String
    sNomeCliente As String
    sNumeroBanco As String
    sNomeBanco As String
    sDataEmissaoArq As String
    sBrancos As String
    sSequencialRegistro As String
End Type

Public Type typePagtoDetalheUnibanco
    sIdentificacaoReg As String
    sAgenciaCliente As String
    sContaCliente As String
    sDVCliente As String
    sRervado01 As String
    sBrancos1 As String
    sComplCepForn As String
    sBancoForn As String
    sAgenciaForn As String
    sContaCorrenteForn As String
    sCodigoEmpresa As String
    sNumFavorecido As String
    sCodSolicitacaoServ As String
    sDVNumFavorecido As String
    sMeioRepasse As String
    sEmissaoAviso As String
    sTipoOperacao As String
    sNumeroPagto As String
    sTipoServico As String
    sCarteira As String
    sValorCredito As String
    sIdentfFavorecido As String
    sDataGravacao As String
    sDataCredito As String
    sCodigoBarras As String
    sUsoBanco As String
    sNossoNumOutroBanco As String
    sSeuNumOutroBanco As String
    sNomeFavorecido As String
    sEndereco As String
    sCidade As String
    sEstado As String
    sCEP As String
    sBrancos2 As String
    sNumeroOTC As String
    sNomeAgencia As String
    sPracaCompensacao As String
    sHistorico As String
    sIdentfTitulo As String
    sMoeda As String
    sCodigoCVT As String
    sDataVencimento As String
    sBrancos3 As String
    sNossoNumero As String
    sDataPagto As String
    sValorTitulo As String
    sValorMora As String
    sValorAbatimento As String
    sValorDesconto As String
    sValorLiquido As String
    sValorMulta As String
    sCodOcorrencia As String
    sMesmaTitularidade As String
    sTipoDocumento As String
    sBrancos4 As String
    sCheckHorizontal As String
    sReferenciaCliente As String
    sBrancos5 As String
    sCepForn As String
    sCodHistorico As String
    sBrancos6 As String
    sDVAgenciaForn As String
    sDVCCFavorecido As String
    sAgCodCedente As String
    sNumSequencialRegistro As String
    sBrancos7 As String
End Type

Public Type typePagtoTrailerArqUnibanco
    sIdentificacaoReg As String
    sTotalQtdRegistros As String
    sTotalValorPagtos As String
    sQuantRegistrosArq As String
    sBrancos As String
    sNumSequencialRegistro As String
End Type

'====================== BANCO DO BRASIL ===============================
'400 bytes
'------------ Remessa de Títulos a Receber ----------------------------

'Registro de Header do arquivo de remessa titulos a receber
Public Type typeRemHeaderBcoBrasil
    sIdentificacaoReg As String
    sIdentificacaoArq As String
    sLiteralRemessa As String
    sCodTipoServico As String
    sLiteralServico As String
    sBrancos1 As String
    sAgencia As String
    sDVAgencia As String
    sCodigoCedente As String
    sDVCodCedente As String
    sNumCovenente As String
    sNomeCedente As String
    sNumeroBanco As String
    sNomeBanco As String
    sDataGravacao As String
    sSeqRemessa As String
    sBrancos2 As String
    sSequencialReg As String
End Type

'Registro de Detalhe do arquivo de remessa titulos a receber
Public Type typeRemDetalheBcoBrasil
    sIdentificacaoReg As String
    sTipoInscricaoCedente As String
    sNumInscricaoCedente As String
    sAgencia As String
    sDVAgencia As String
    sCodigoCedente As String
    sDVCodCedente As String
    sNumConvenio As String
    sNumControleCli As String
    sNossoNumero As String
    sDVNossoNumero As String
    sNumPrestacao As String
    sIndGrupoValor As String
    sBrancos1 As String
    sPrefixoTitulo As String
    sVariacao As String
    sContaCaucao As String
    sCodRespons As String
    sDVCodRespons As String
    sNumBordero As String
    sBrancos2 As String
    sNumCarteiraBanco As String
    sComando As String
    sNumTitulo As String
    sVencimento As String
    sValorTitulo As String
    sNumeroBanco As String
    sAgenciaCobradora As String
    sDVAgenciaCobradora As String
    sEspecie As String
    sAceite As String
    sDataEmissao As String
    sInstrucao1 As String
    sInstrucao2 As String
    sJurosDiarios As String
    sDescontoAte As String
    sValorDesconto As String
    sValorIOF As String
    sQuantUnidVar As String
    sEspecieValor As String
    sAbatimento As String
    sTipoInscSacado As String
    sNumInscricaoSacado As String
    sNomeSacado As String
    sBrancos3 As String
    sEndereco As String
    sBrancos4 As String
    sCEP As String
    sCidade As String
    sEstado As String
    sObservacoes As String
    sBrancos5 As String
    sNumSequencialReg As String
End Type

'Registro de Trailer do arquivo de remessa titulos a receber
Type typeRemTrailerBcoBrasil
    sIdentificacaoReg As String
    sBrancos As String
    sSequencialRegistro As String
End Type


'------------------ Retorno de Títulos a Receber -------------------------

'Registro de Header do arquivo de retorno titulos a receber
Public Type typeRetHeaderBcoBrasil
    sIdentificacaoReg As String
    sIdentificacaoArq As String
    sLiteralRetorno As String
    sCodTipoServico As String
    sLiteralServico As String
    sBrancos1 As String
    sAgencia As String
    sDVAgencia As String
    sCodigoCedente As String
    sDVCodCedente As String
    sNumCovenente As String
    sNomeCedente As String
    sNumeroBanco As String
    sNomeBanco As String
    sDataGravacao As String
    sSeqRetorno As String
    sBrancos2 As String
    sSequencialReg As String
End Type

'Registro de Detalhe do arquivo de retorno titulos a receber
Public Type typeRetDetalheBcoBrasil
    sIdentificacaoReg As String
    sTipoInscricaoCedente As String
    sNumInscricaoCedente As String
    sAgencia As String
    sDVAgencia As String
    sCodigoCedente As String
    sDVCodCedente As String
    sNumConvenio As String
    sNumControleCli As String
    sNossoNumero As String
    sDVNossoNumero As String
    sNumPrestacao As String
    sDiasCalculo As String
    sIndGrupoValor As String
    sNaturezaRecebto As String
    sPrefixoTitulo As String
    sVariacao As String
    sContaCaucao As String
    sCodRespons As String
    sDVCodRespons As String
    sTaxaDesconto As String
    sTaxaIOF As String
    sBrancos1 As String
    sNumCarteiraBanco As String
    sComando As String
    sDataLiquidacao As String
    sNumTitulo As String
    sNossoNumero2 As String
    sVencimento As String
    sValorTitulo As String
    sNumeroBanco As String
    sAgenciaCobradora As String
    sDVAgenciaCobradora As String
    sEspecie As String
    sDataCredito As String
    sValorTarifa As String
    sOutrasDespesas As String
    sJurosDesconto As String
    sIOFDesconto As String
    sAbatimento As String
    sDescontoConcedido As String
    sValorRecebido As String
    sJurosMora As String
    sOutrosRecebtos As String
    sAbatNaoAproveitado As String
    sValorLancto As String
    sIndicativoDebCred As String
    sIndicativoValor As String
    sValorAjuste As String
    sBrancos2 As String
    sDataProcessamento As String
    sTituloRazao As String
    sDVTituloRazao As String
    sOrigem As String
    sValorlanctosDep As String
    sHistorico As String
    sNumDocto As String
    sDataValorizacao As String
    sNumSequencialReg As String
End Type

'Registro de trailer do arquivo de retorno titulos a receber
Public Type typeRetTrailerBcoBrasil
    sIdentificacaoReg As String
    sCodigoRetorno As String
    sCodigoServico As String
    sNumeroBanco As String
    sBrancos1 As String
    sQuantTitulosCobrSimp As String
    sValorTotalSimp As String
    sAvisoBancario As String
    sBrancos2 As String
    sQuantTitulosCobrVinc As String
    sValorTotalVinc As String
    sAvisoBancario2 As String
    sBrancos3 As String
    sQuantTitulosCobrCauc As String
    sValorTotalCauc As String
    sAvisoBancario3 As String
    sBrancos4 As String
    sQuantTitulosCobrDesc As String
    sValorTotalDesc As String
    sAvisoBancario4 As String
    sBrancos5 As String
    sSequencialRegistro As String
    
End Type

'============================ BANDEIRANTES ===================================

'------------------ Remessa de Titulos a Receber --------------------

'Registro de Header do arquivo de remessa titulos a receber
Public Type typeRemHeaderBandeirantes
    sIdentificacaoReg As String
    sIdentificacaoArq As String
    sLiteralRemessa As String
    sCodTipoServico As String
    sLiteralServico As String
    sCodigoEmpresa As String
    sBrancos1 As String
    sNomeEmpresa As String
    sNumeroBanco As String
    sNomeBanco As String
    sDataGravacao As String
    sDensidade As String
    sLiteralDensidade As String
    sBrancos2 As String
    sSequencialRegistro As String
End Type

'Registro de Detalhe do arquivo de remessa titulos a receber
Public Type typeRemDetalheBandeirantes
    sIdentificacaoReg As String
    sTipoInscricaoEmpSac As String
    sNumInscricaoEmpSac As String
    sIdentificacaoEmp As String
    sBrancos1 As String
    sUsoDaEmpresa As String
    sNossoNumero As String
    sBrancos2 As String
    sDiasDevolucao As String
    sUsoDoBanco As String
    sCodCarteira As String
    sCodOcorrencia As String
    sSeuNumero As String
    sVencimento As String
    sValorTitulo As String
    sNumeroBanco As String
    sAgenciaCobradora As String
    sEspecie As String
    sAceite As String
    sDataEmissao As String
    sInstrucao1 As String
    sInstrucao2 As String
    sJurosDiarios As String
    sDescontoAte As String
    sValorDesconto As String
    sValorIOF As String
    sAbatimento As String
    sCodigoInscSacado As String
    sNumInscricaoSacado As String
    sNome As String
    sEndereco As String
    sBairro As String
    sCEP As String
    sCidade As String
    sEstado As String
    sMensagem As String
    sBrancos3 As String
    sCodigoMoeda As String
    sNumSequencial As String
End Type

'Registro de Trailer do arquivo de remessa titulos a receber
Type typeRemTrailerBandeirantes
    sIdentificacaoReg As String
    sComplemento1 As String
    sSequencialRegistro As String
End Type

'------------------ Retorno de Títulos a Receber -------------------------

'Registro de Header do arquivo de retorno titulos a receber
Public Type typeRetHeaderBandeirantes
    sIdentificacaoReg As String
    sIdentificacaoArq As String
    sLiteralRetorno As String
    sCodTipoServico As String
    sLiteralServico As String
    sCodEmpresa As String
    sNomeEmpresa As String
    sNumeroBanco As String
    sNomeBanco As String
    sDataEmissaoArq As String
    sDensidade As String
    sLiteralDensidade As String
    sNumSequencialArq As String
    sBrancos As String
    sSequencialRegistro As String
End Type

'Registro de detalhe do arquivo de retorno de titulos a receber
Public Type typeRetDetalheBandeirantes
    sIdentificacaoReg As String
    sTipoInscricaoEmpresa As String
    sNumInscricaoEmpresa As String
    sCodEmpresa As String
    sBrancos1 As String
    sUsoDaEmpresa As String
    sNossoNumero1 As String
    sBrancos2 As String
    sUsoDoBanco As String
    sCodCarteira As String
    sCodOcorrencia As String
    sDataOcorrencia As String
    sSeuNumero As String
    sNossoNumero2 As String
    sBrancos3 As String
    sVencimento As String
    sValorTitulo As String
    sNumeroBanco As String
    sAgenciaCobradora As String
    sEspecie As String
    sTarifaCobranca As String
    sBrancos4 As String
    sValorIOF As String
    sOutrasDeducoes As String
    sDescontos As String
    sValorPrincipal As String
    sJuros As String
    sCodigoMoeda As String
    sAgenciaRec As String
    sBrancos5 As String
    sTipoCobranca As String
    sDataCredito As String
    sBrancos6 As String
    sNumSequencial As String
End Type

'Registro de trailer do arquivo de retorno titulos a receber
Public Type typeRetTrailerBandeirantes
    sIdentificacaoReg As String
    sCodigoRetorno As String
    sCodigoServico As String
    sNumeroBanco As String
    sBrancos1 As String
    sQuantTitulos As String
    sValorTotal As String
    sAvisoBancario As String
    sBrancos2 As String
    sSequencialRegistro As String
End Type

'====================== BANCO DO BRASIL ===============================
'240 BYTES versao 3.0

'Registro de Header do arquivo
Public Type typeHeaderBcoBrasil030 '030 é a versao
    sNumeroBanco As String
    sLoteServico As String
    sIdentificacaoReg As String
    sReservadoCNAB1 As String
    sTipoInscricaoEmpresa As String
    sNumInscricaoEmpresa As String
    sConvenio As String
    sAgencia As String
    sDVAgencia As String
    sContaCorrente As String
    sDVContaCorrente As String
    sDVAgenciaConta As String
    sNomeEmpresa As String
    sNomeBanco As String
    sReservadoCNAB2 As String
    sArqCodigo As String
    sArqDataGeracao As String
    sArqHoraGeracao As String
    sArqSequencial As String
    sArqVersaoLayout As String
    sArqDensidadeGrav As String
    sReservadoBanco As String
    sReservadoEmpresa As String
    sReservadoCNAB3 As String
    sIdentifVANS As String
    sControleVANS As String
    sServico As String
    sOcorrencias As String
End Type

'Registro de trailer do arquivo
Public Type typeTrailerBcoBrasil030 '030 é a versao
    sNumeroBanco As String
    sLoteServico As String
    sIdentificacaoReg As String
    sReservadoCNAB1 As String
    sQtdeLotes As String
    sQtdeRegistros As String
    sQtdeContasConciliacao As String
    sReservadoCNAB2 As String
End Type

''------------ Remessa de Títulos a Receber ----------------------------
'
''Registro de Header do arquivo de remessa titulos a receber
'Public Type typeRemHeaderBcoBrasil
'    sIdentificacaoReg As String
'    sIdentificacaoArq As String
'    sLiteralRemessa As String
'    sCodTipoServico As String
'    sLiteralServico As String
'    sBrancos1 As String
'    sAgencia As String
'    sDVAgencia As String
'    sCodigoCedente As String
'    sDVCodCedente As String
'    sNumCovenente As String
'    sNomeCedente As String
'    sNumeroBanco As String
'    sNomeBanco As String
'    sDataGravacao As String
'    sSeqRemessa As String
'    sBrancos2 As String
'    sSequencialReg As String
'End Type
'
''Registro de Detalhe do arquivo de remessa titulos a receber
'Public Type typeRemDetalheBcoBrasil
'    sIdentificacaoReg As String
'    sTipoInscricaoCedente As String
'    sNumInscricaoCedente As String
'    sAgencia As String
'    sDVAgencia As String
'    sCodigoCedente As String
'    sDVCodCedente As String
'    sNumConvenio As String
'    sNumControleCli As String
'    sNossoNumero As String
'    sDVNossoNumero As String
'    sNumPrestacao As String
'    sIndGrupoValor As String
'    sBrancos1 As String
'    sPrefixoTitulo As String
'    sVariacao As String
'    sContaCaucao As String
'    sCodRespons As String
'    sDVCodRespons As String
'    sNumBordero As String
'    sBrancos2 As String
'    sNumCarteiraBanco As String
'    sComando As String
'    sNumTitulo As String
'    sVencimento As String
'    sValorTitulo As String
'    sNumeroBanco As String
'    sAgenciaCobradora As String
'    sDVAgenciaCobradora As String
'    sEspecie As String
'    sAceite As String
'    sDataEmissao As String
'    sInstrucao1 As String
'    sInstrucao2 As String
'    sJurosDiarios As String
'    sDescontoAte As String
'    sValorDesconto As String
'    sValorIOF As String
'    sQuantUnidVar As String
'    sEspecieValor As String
'    sAbatimento As String
'    sTipoInscSacado As String
'    sNumInscricaoSacado As String
'    sNomeSacado As String
'    sBrancos3 As String
'    sEndereco As String
'    sBrancos4 As String
'    sCEP As String
'    sCidade As String
'    sEstado As String
'    sObservacoes As String
'    sBrancos5 As String
'    sNumSequencialReg As String
'End Type
'
''Registro de Trailer do arquivo de remessa titulos a receber
'Type typeRemTrailerBcoBrasil
'    sIdentificacaoReg As String
'    sBrancos As String
'    sSequencialRegistro As String
'End Type


'------------------ Retorno de Títulos a Receber -------------------------

'===============================FIM CNAB==================================
'===============================INICIO REAL==================================

Public Type typeRetCobrHeaderLoteCNAB240_030 '030 é a versao
    sNumeroBanco As String
    sLoteServico As String
    sIdentificacaoReg As String
    sTipoOperacao As String
    sTipoServico As String
    sFormaLcto As String
    sLayoutLote As String
    sReservadoCNAB1 As String
    sTipoInscricaoEmpresa As String
    sNumInscricaoEmpresa As String
    sConvenio As String
    sAgencia As String
    sDVAgencia As String
    sContaCorrente As String
    sDVContaCorrente As String
    sDVAgenciaConta As String
    sNomeEmpresa As String
    sMensagem1 As String
    sMensagem2 As String
    sNumRemRet As String
    sDataRemRet As String
    sDataCredito As String
    sReservadoCNAB2 As String
End Type

Public Type typeRetCobrTrailerLoteCNAB240_030 '030 é a versao
    sNumeroBanco As String
    sLoteServico As String
    sIdentificacaoReg As String
    sReservadoCNAB1 As String
    sQtdeRegs As String
    sQtdeTitCobrSimples As String
    sValTitCobrSimples As String
    sQtdeTitCobrVinculada As String
    sValTitCobrVinculada As String
    sQtdeTitCobrCaucionada As String
    sValTitCobrCaucionada As String
    sQtdeTitCobrDescontada As String
    sValTitCobrDescontada As String
    sNumAviso As String
    sReservadoCNAB2 As String
End Type

'Registro de Detalhe do arquivo de retorno titulos a receber segmento T versao 3.0
Public Type typeRetCobrDetalheCNAB240_T030
    sNumeroBanco As String
    sLoteServico As String
    sIdentificacaoReg As String
    sNumSeqRegLote As String
    sCodSegmento As String
    sReservadoCNAB1 As String
    sCodMovto As String
    sAgencia As String
    sDVAgencia As String
    sContaCorrente As String
    sDVContaCorrente As String
    sDVAgenciaConta As String
    sNossoNumero As String
    sCodCarteira As String
    sNumDoc As String
    sVencimento As String
    sValorTitulo As String
    sBcoCobrRec As String
    sAgCobrRec As String
    sDVAgCobrRec As String
    sUsoEmpresa As String
    sCodMoeda As String
    sSacadoTipoInscr As String
    sSacadoNumInscr As String
    sSacadoNome As String
    sNumContrato As String
    sValorTarifa As String
    sMotivoOcorrencia As String
    sReservadoCNAB2 As String
End Type

'Registro de Detalhe do arquivo de retorno titulos a receber segmento U versao 3.0
Public Type typeRetCobrDetalheCNAB240_U030
    sNumeroBanco As String
    sLoteServico As String
    sIdentificacaoReg As String
    sNumSeqRegLote As String
    sCodSegmento As String
    sReservadoCNAB1 As String
    sCodMovto As String
    sJurosMulta As String
    sDesconto As String
    sAbatimento As String
    sIOFRecolhido As String
    sValorPagoSacado As String
    sValorCreditado As String
    sOutrasDespesas As String
    sOutrosCreditos As String
    sDataOcorrencia As String
    sDataCredito As String
    sOcorrSacCodigo As String
    sOcorrSacData As String
    sOcorrSacValor As String
    sOcorrSacCompl As String
    sCodBcoCorresp As String
    sNossoNumBcoCorresp As String
    sReservadoCNAB2 As String
End Type

'Registro de Detalhe do arquivo de retorno titulos a receber global (incorporando segmentos T e U)
Public Type typeRetCobrDetalheCNAB240
    tT030 As typeRetCobrDetalheCNAB240_T030
    tU030 As typeRetCobrDetalheCNAB240_U030
End Type

'Registro de Header do arquivo de retorno titulos a receber
Public Type typeRetHeaderCNAB240
    sIdentificacaoReg As String
    sIdentificacaoArq As String
    sLiteralRetorno As String
    sCodTipoServico As String
    sLiteralServico As String
    sBrancos1 As String
    sAgencia As String
    sDVAgencia As String
    sCodigoCedente As String
    sDVCodCedente As String
    sNumCovenente As String
    sNomeCedente As String
    sNumeroBanco As String
    sNomeBanco As String
    sDataGravacao As String
    sSeqRetorno As String
    sBrancos2 As String
    sSequencialReg As String
End Type

'Registro de Detalhe do arquivo de retorno titulos a receber
Public Type typeRetDetalheCNAB240
    sIdentificacaoReg As String
    sTipoInscricaoCedente As String
    sNumInscricaoCedente As String
    sAgencia As String
    sDVAgencia As String
    sCodigoCedente As String
    sDVCodCedente As String
    sNumConvenio As String
    sNumControleCli As String
    sNossoNumero As String
    sDVNossoNumero As String
    sNumPrestacao As String
    sDiasCalculo As String
    sIndGrupoValor As String
    sNaturezaRecebto As String
    sPrefixoTitulo As String
    sVariacao As String
    sContaCaucao As String
    sCodRespons As String
    sDVCodRespons As String
    sTaxaDesconto As String
    sTaxaIOF As String
    sBrancos1 As String
    sNumCarteiraBanco As String
    sComando As String
    sDataLiquidacao As String
    sNumTitulo As String
    sNossoNumero2 As String
    sVencimento As String
    sValorTitulo As String
    sNumeroBanco As String
    sAgenciaCobradora As String
    sDVAgenciaCobradora As String
    sEspecie As String
    sDataCredito As String
    sValorTarifa As String
    sOutrasDespesas As String
    sJurosDesconto As String
    sIOFDesconto As String
    sAbatimento As String
    sDescontoConcedido As String
    sValorRecebido As String
    sJurosMora As String
    sOutrosRecebtos As String
    sAbatNaoAproveitado As String
    sValorLancto As String
    sIndicativoDebCred As String
    sIndicativoValor As String
    sValorAjuste As String
    sBrancos2 As String
    sDataProcessamento As String
    sTituloRazao As String
    sDVTituloRazao As String
    sOrigem As String
    sValorlanctosDep As String
    sHistorico As String
    sNumDocto As String
    sDataValorizacao As String
    sNumSequencialReg As String
End Type

'Registro de trailer do arquivo de retorno titulos a receber
Public Type typeRetTrailerCNAB240
    sIdentificacaoReg As String
    sCodigoRetorno As String
    sCodigoServico As String
    sNumeroBanco As String
    sBrancos1 As String
    sQuantTitulosCobrSimp As String
    sValorTotalSimp As String
    sAvisoBancario As String
    sBrancos2 As String
    sQuantTitulosCobrVinc As String
    sValorTotalVinc As String
    sAvisoBancario2 As String
    sBrancos3 As String
    sQuantTitulosCobrCauc As String
    sValorTotalCauc As String
    sAvisoBancario3 As String
    sBrancos4 As String
    sQuantTitulosCobrDesc As String
    sValorTotalDesc As String
    sAvisoBancario4 As String
    sBrancos5 As String
    sSequencialRegistro As String
    
End Type

'Registro de Header de Arquivo de Remessa/Retorno de títulos a pagar
Public Type typePagtoHeaderArqCNAB240
    sIdentificacaoReg As String
    sCodigoRemessa As String
    sBrancos1 As String
    sCodTipoServico As String
    sindicadorCGC As String
    sValorTarifa As String
    sBrancos2 As String
    sCGCouPrefixoAg As String
    sDVAgencia As String
    sContaCorrente As String
    sDVContaCorrente As String
    sBrancos3 As String
    sNomeEmpresa As String
    sNumeroBanco As String
    sCodigoConvenente As String
    sTipoRetorno As String
    sUsoDaEmpresa As String
    sMeioFisicoRet As String
    sContadorRemessa As String
    sBrancos4 As String
    sOpcaoLayOut As String
    sBrancos5 As String
    sTipoRetornoEnv As String
    sSequencialRegistro As String
End Type

Public Type typePagtoDetalheCNAB240
    sIdentificacaoReg As String
    sBrancos1 As String
    sIndicadorConf As String
    sCGCouCPFConferir As String
    sDVCGCouCPF As String
    sCGCouPrefixoEmpresa As String
    sDVPrefixo As String
    sNumContaEmpresa As String
    sDVNumConta As String
    sUsoDaEmpresa As String
    sBrancos2 As String
    sCodCamara As String
    sCodBancoDest As String
    sAgenciaFavorecido As String
    sDVAgenciaFavorecido As String
    sNumContaFavorecido As String
    sDVContaFavorecido As String
    sBrancos3 As String
    sNomeFavorecido As String
    sDataProcessamento As String
    sValor As String
    sCodServico As String
    sMensagem As String
    sOcorrenciasRetorno As String
    sNumSequencialRegistro As String
End Type

Public Type typePagtoTrailerArqCNAB240
    sIdentificacaoReg As String
    sBrancos As String
    sNumSequencialRegistro As String
End Type

'Registro de Header do arquivo de remessa titulos a receber
Public Type typeRemHeaderCNAB240
    sCodBancoComp As String
    sLoteServico As String
    sRegHeaderLote As String
    sTipoOperacao As String
    sTipoServico As String
    sFormaLancamento As String
    sNVersaoLayLote As String
    sBrancosCNAB1 As String
    sTipoInscEmpresa As String
    sCodInsEmpresa As String
    sCodConvBanco As String
    sAgMantConta As String
    sDVAgencia As String
    sNumContaCorrente As String
    sDVConta As String
    sDVAgenciaConta As String
    sNomeEmpresa As String
    sMensagem1 As String
    sMensagem2 As String
    sNumRemessa As String
    sDataGravRem As String
    sDataCredito As String
    sBrancosCNAB2 As String
    sReservadoEmpresa As String
End Type


Public Type typeRemDetalheCNAB240SegP
    sCodBancoComp As String
    sLoteServico As String
    sRegDetalhe As String
    sNumSequencialReg As String
    sCodSegRegDetalhe As String
    sBrancosCNAB1 As String
    sCodMovimento As String
    sAgMantConta As String
    sDVAgMantConta As String
    sNumContaCorrente As String
    sDVConta As String
    sDVAgenciaConta As String
    sDirecionamentoCobranca As String
    sModalidadeCorrespondentes As String
    sBrancosCNAB3 As String
    sModalidadeBcoCobranca As String
    sIdentTituloBanco As String
    sCodCarteira As String
    sFormaCadTituloBanco As String
    sTipoDocumento As String
    sIdentEmissaoBloq As String
    sIdentDistrib As String
    sNumDocCobranca As String
    sDataVencTitulo As String
    sValorNominalTitulo As String
    sAgEncCobranca As String
    sDVAgencia As String
    sEspecieTitulo As String
    sIdentTituloAceite As String
    sDataEmissaoTitulo As String
    sCodJurosMora As String
    sDataJurosMora As String
    sJurosMoraDiaTaxa As String
    sCodDesconto1 As String
    sDataDesconto1 As String
    sValorPerConcedido As String
    sValorIOF As String
    sValorAbatimento As String
    sIdentTituloEmp As String
    sCodProtesto As String
    sNumDiasProtesto As String
    sCodBaixaDevolucao As String
    sNumDiasBaixaDevol As String
    sCodMoeda As String
    sNumContratoOpCred As String
    sBrancosCNAB2 As String
    sContaCobranca As String
    sDVContaCobranca As String
End Type

Public Type typeRemDetalheCNAB240SegQ
    sCodBancoComp As String
    sLoteServico As String
    sRegDetalhe As String
    sNumSequencialReg As String
    sCodSegRegDetalhe As String
    sBrancosCNAB1 As String
    sCodMovimento As String
    sTipoInscricao As String
    sNumInscricao As String
    sNome As String
    sEndereco As String
    sBairro As String
    sCEP  As String
    sSufixoCEP As String
    sCidade As String
    sUnidFederacao As String
    sTipoInscricaoSacAval As String
    sNumInscricaoSacAval As String
    sNomeSacadorAvalista As String
    sBancoCompensacao As String
    sNossoNumBancoCorresp As String
    sBrancosCNAB2 As String
  
End Type

'Registro de Detalhe do arquivo de remessa titulos a receber
Public Type typeRemDetalheCNAB240SegR
    sCodBancoComp As String
    sLoteServico As String
    sRegDetalhe As String
    sNumSequencialReg As String
    sCodSegRegDetalhe As String
    sBrancosCNAB1 As String
    sCodMovimento As String
    sCodDesconto2 As String
    sDataDesconto2 As String
    sValorPerConcedido2 As String
    sCodDesconto3 As String
    sDataDesconto3 As String
    sValorPerConcedido3 As String
    sCodMulta As String
    sDataMulta As String
    sValorPerAplicado As String
    sInfoBancoSacado As String
    sMensagem3 As String
    sMensagem4 As String
    sBrancosCNAB3 As String
    sCodOcorrenciaSac As String
    sCodBancoContaDeb As String
    sCodAgContaDeb As String
    sCodAgDVDeb As String
    sCodContaDVDeb As String
    sCodAgContaDVDeb As String
    sBrancosCNAB2 As String
    
End Type

'Registro de Trailer do arquivo de remessa titulos a receber
Type typeRemTrailerCNAB240
    
    sCodBancoComp As String
    sLoteServico As String
    sRegTrailer As String
    sBrancosCNAB1 As String
    sQuantRegLote As String
    sQuantTitCobS As String
    sValTotalTitCartS As String
    sQuantTitCobV As String
    sValTotalTitCartV As String
    sQuantTitCobC As String
    sValTotalTitCartC As String
    sQuantTitCobD As String
    sValTotalTitCartD As String
    sNumAvisoLancamento As String
    sBrancosCNAB2 As String

End Type

Public Type typeTrailerArqCNAB240
    sNumeroBanco As String
    sLoteServico As String
    sIdentificacaoReg As String
    sReservadoCNAB1 As String
    sQtdeLotes As String
    sQtdeRegistros As String
    sQtdeContasConciliacao As String
    sReservadoCNAB2 As String
End Type


'============================ PRIMUS ===================================

'------------------ Remessa de Titulos a Receber --------------------

'Registro de Header do arquivo de remessa titulos a receber
Public Type typeRemHeaderPrimus
    sIdentificacaoReg As String
    sBrancosCNAB1 As String
    sCodigoEmpresa As String
    sBrancosCNAB2 As String
    sNomeEmpresa As String
    sBrancosCNAB3 As String
    sDataEmissaoArq As String
    sBrancosCNAB4 As String
    sContaCobrancaDireta As String
    sBrancosCNAB5 As String
    sNumSequencial As String
End Type

'Registro de Detalhe do arquivo de remessa titulos a receber
Public Type typeRemDetalhePrimus
    sIdentificacaoReg As String
    sBrancosCNAB1 As String
    sCGCEmpresa As String
    sCodigoEmpresa As String
    sBrancosCNAB2 As String
    sUsoEmpresa As String
    sNossoNumero As String
    sBancoCobrancaDireta As String
    sCodigoCarteira As String
    sIdentificacaoOcorr As String
    sSeuNumero As String
    sDataVencto As String
    sValorTitulo As String
    sBrancosCNAB3 As String
    sEspecie As String
    sAceite As String
    sDataEmissao As String
    sInstrucao1 As String
    sInstrucao2 As String
    sValorJuros As String
    sDataDesconto As String
    sValorDesconto As String
    sAbatimento As String
    sIdentfInscSacado As String
    sNumInscricaoSacado As String
    sNomeSacado As String
    sEnderecoSacado As String
    sBairroSacado As String
    sCEPSacado As String
    sCidadeSacado As String
    sUFSacado As String
    sMensagem1 As String
    sDataJurosMora As String
    sPrazo As String
    sNumSequencialRegistro As String
    sTipoEmitente As String
    sNomeEmitente As String
    sEnderecoEmitente As String
    sCidadeEmitente As String
    sUFEmitente As String
    sBrancosCNAB4 As String
    sBrancosCNAB5 As String
    sBrancosCNAB6 As String
    sBrancosCNAB7 As String
    sBrancosCNAB8 As String
    sBrancosCNAB9 As String

End Type

'Registro de Trailer do arquivo de remessa titulos a receber
Type typeRemTrailerPrimus
    sIdentificacaoReg As String
    sComplemento1 As String
    sSequencialRegistro As String
End Type


'------------------ Retorno de Titulos a Receber --------------------

'Registro de Header do arquivo de retorno titulos a receber
Public Type typeRetHeaderPrimus
    sIdentificacaoReg As String
    sIdentificacaoArq As String
    sCodEmpresa As String
    sNumeroBanco As String
    sContaCobranca As String
    sSequencialRegistro As String
End Type

'Registro de detalhe do arquivo de retorno de titulos a receber
Public Type typeRetDetalhePrimus
    sIdentificacaoReg As String
    sIdentfEmpresa As String
    sNumControle As String
    sNossoNumero1 As String
    sCodCarteira As String
    sCodOcorrencia As String
    sDataOcorrencia As String
    sNumDocto As String
    sNossoNumero2 As String
    sVencimento As String
    sValorTitulo As String
    sBancoCobrador As String
    sTarifaCobranca As String
    sAbatimento As String
    sDescontos As String
    sValorPrincipal As String
    sJuros As String
    sOutrosCreditos As String
    sCodigoErro As String
    sNumSequencialRegistro As String
End Type

'Registro de trailer do arquivo de retorno titulos a receber
Public Type typeRetTrailerPrimus
    sIdentificacaoReg As String
    sSequencialRegistro As String
    
End Type


'============================ RURAL ===================================

'------------------ Remessa de Titulos a Receber --------------------

'Registro de Header do arquivo de remessa titulos a receber
Public Type typeRemHeaderRural
    sIdentificacaoReg As String
    sIdentificacaoArq As String
    sLiteralRemessa As String
    sCodTipoServico As String
    sLiteralServico As String
    sCodigoEmpresa As String
    sBrancosCNAB1 As String
    sNomeEmpresa As String
    sNumeroBanco As String
    sNomeBanco As String
    sBrancosCNAB2 As String
    sCodigoVersao As String
    sBrancosCNAB3 As String
    sSequencialRegistro As String
End Type

'Registro de Detalhe do arquivo de remessa titulos a receber
Public Type typeRemDetalheRural
    sIdentificacaoReg As String
    sBrancosCNAB1 As String
    sCodigoEmpresa As String
    sPrazoProtesto As String
    sMoeda As String
    sTipoMora As String
    sIdentifMensagem As String
    sNumControle As String
    sNossoNumero As String
    sDVNossoNumero As String
    sNumContrato As String
    sBrancosCNAB2 As String
    sSacadorAvalista As String
    sBrancosCNAB3 As String
    sCodCarteira As String
    sIdentificacaoOcorr As String
    sNumDocto As String
    sDataVencto As String
    sValorTitulo As String
    sBancoCobranca As String
    sBrancosCNAB4 As String
    sEspecie As String
    sAceite As String
    sDataEmissao As String
    sInstrucao1 As String
    sValorJuros As String
    sDataDesconto As String
    sValorDesconto As String
    sValorIOF As String
    sAbatimento As String
    sIdentfInscSacado As String
    sNumInscricaoSacado As String
    sNomeSacado As String
    sEnderecoSacado As String
    sComplemento As String
    sCEP As String
    sCidadeSacado As String
    sEstado As String
    sSacadorAvalista2 As String
    sNumSequencialRegistro As String

End Type

'Registro de Trailer do arquivo de remessa titulos a receber
Type typeRemTrailerRural
    sIdentificacaoReg As String
    sComplemento1 As String
    sSequencialRegistro As String
End Type

'------------------ Retorno de Titulos a Receber --------------------

'Registro de Header do arquivo de retorno titulos a receber
Public Type typeRetHeaderRural
    sIdentificacaoReg As String
    sIdentificacaoArq As String
    sLiteralRetorno As String
    sCodTipoServico As String
    sLiteralServico As String
    sCodEmpresa As String
    sNomeEmpresa As String
    sNumeroBanco As String
    sNomeBanco As String
    sDataEmissaoArq As String
    sZeros As String
    sNumAviso As String
    sComplemento1 As String
    sDataCredito As String
    sComplemento2 As String
    sSequencialRegistro As String
End Type

'Registro de detalhe do arquivo de retorno de titulos a receber
Public Type typeRetDetalheRural
    sIdentificacaoReg As String
    sTipoInscricaoEmpresa As String
    sNumInscricaoEmpresa As String
    sIdentfEmpresa As String
    sDiasProtesto As String
    sMoeda As String
    sNumControle As String
    sNossoNumero1 As String
    sDVNossoNumero1 As String
    sNumContrato As String
    sUsoDoBanco2 As String
    sIndicadorRateio As String
    sZeros2 As String
    sCodCarteira As String
    sCodOcorrencia As String
    sDataOcorrencia As String
    sNumDocto As String
    sVencimento As String
    sValorTitulo As String
    sBancoCobrador As String
    sEspecie As String
    sTarifaCobranca As String
    sValorIOF As String
    sAbatimento As String
    sDescontos As String
    sValorPrincipal As String
    sJuros As String
    sOutrosCreditos As String
    sDataCredito As String
    sTabelaErros As String
    sNumSequencialRegistro As String
End Type

'Registro de trailer do arquivo de retorno titulos a receber
Public Type typeRetTrailerRural
    sIdentificacaoReg As String
    sCodigoRetorno As String
    sCodigoTipoRegistro As String
    sNumeroBanco As String
    sComplemento1 As String
    sQuantTitulosCobr As String
    sValorTotalCobr As String
    sAvisoBancario As String
    sQuantTitulosCobrDireta As String
    sValorTotalCobrDireta As String
    sQuantTitulosNaoCad As String
    sValorTotalTitNaoCad As String
    sSequencialRegistro As String
    
End Type


'============================ BICBANCO ===================================

'------------------ Remessa de Titulos a Receber --------------------

'Registro de Header do arquivo de remessa titulos a receber
Public Type typeRemHeaderBicBanco
    sIdentificacaoReg As String
    sIdentificacaoArq As String
    sLiteralRemessa As String
    sCodTipoServico As String
    sLiteralServico As String
    sCodigoEmpresa As String
    sBrancos1 As String
    sAgencia As String
    sNomeEmpresa As String
    sNumeroBanco As String
    sNomeBanco As String
    sDataGravacao As String
    sDensidade As String
    sLiteralDensidade As String
    sSequencialArq As String
    sBrancos2 As String
    sSequencialRegistro As String
End Type

'Registro de Detalhe do arquivo de remessa titulos a receber
Public Type typeRemDetalheBicBanco
    sIdentificacaoReg As String
    sTipoInscricaoEmpSac As String
    sNumInscricaoEmpSac As String
    sIdentificacaoEmp As String
    sBrancos1 As String
    sUsoDaEmpresa As String
    sNossoNumero As String
    sBrancos2 As String
    sDiasDevolucao As String
    sUsoDoBanco As String
    sCodCarteira As String
    sCodOcorrencia As String
    sSeuNumero As String
    sVencimento As String
    sValorTitulo As String
    sNumeroBanco As String
    sAgenciaCobradora As String
    sEspecie As String
    sAceite As String
    sDataEmissao As String
    sInstrucao1 As String
    sInstrucao2 As String
    sJurosDiarios As String
    sDescontoAte As String
    sValorDesconto As String
    sValorIOF As String
    sAbatimento As String
    sCodigoInscSacado As String
    sNumInscricaoSacado As String
    sNome As String
    sEndereco As String
    sBairro As String
    sBrancosCNAB2 As String
    sCEP As String
    sCidade As String
    sEstado As String
    sMensagem As String
    sBrancos3 As String
    sCodigoMoeda As String
    sNumSequencial As String
End Type

'Registro de Trailer do arquivo de remessa titulos a receber
Type typeRemTrailerBicBanco
    sIdentificacaoReg As String
    sComplemento1 As String
    sSequencialRegistro As String
End Type

'------------------ Retorno de Títulos a Receber -------------------------

'Registro de Header do arquivo de retorno titulos a receber
Public Type typeRetHeaderBicBanco
    sIdentificacaoReg As String
    sIdentificacaoArq As String
    sLiteralRetorno As String
    sCodTipoServico As String
    sLiteralServico As String
    sCodEmpresa As String
    sNomeEmpresa As String
    sNumeroBanco As String
    sNomeBanco As String
    sDataEmissaoArq As String
    sDensidade As String
    sLiteralDensidade As String
    sNumSequencialArq As String
    sBrancos As String
    sSequencialRegistro As String
End Type

'Registro de detalhe do arquivo de retorno de titulos a receber
Public Type typeRetDetalheBicBanco
    sIdentificacaoReg As String
    sTipoInscricaoEmpresa As String
    sNumInscricaoEmpresa As String
    sCodEmpresa As String
    sBrancos1 As String
    sUsoDaEmpresa As String
    sNossoNumero1 As String
    sBrancos2 As String
    sCodCarteira As String
    sCodOcorrencia As String
    sDataOcorrencia As String
    sSeuNumero As String
    sNossoNumero2 As String
    sCodRejeicao As String
    sBrancos3 As String
    sVencimento As String
    sValorTitulo As String
    sNumeroBanco As String
    sAgenciaCobradora As String
    sEspecie As String
    sTarifaCobranca As String
    sOutrasDespesas As String
    sDescontos As String
    sValorIOF As String
    sOutrasDeducoes As String
    sValorPrincipal As String
    sJuros As String
    sOutrosCreditos As String
    sDataCredito As String
    sValorLiquido  As String
    sNomeSacado As String
    sMoeda As String
    sNumSequencial As String
End Type

'Registro de trailer do arquivo de retorno titulos a receber
Public Type typeRetTrailerBicBanco
    sIdentificacaoReg As String
    sCodigoRetorno As String
    sCodigoServico As String
    sNumeroBanco As String
    sBrancos1 As String
    sQuantTitulos As String
    sValorTotal As String
    sAvisoBancario As String
    sDataCredito As String
    sBrancos2 As String
    sSequencialRegistro As String
End Type


'============================ MERCANTIL ===================================

'------------------ Remessa de Titulos a Receber --------------------

'Registro de Header do arquivo de remessa titulos a receber
Public Type typeRemHeaderMercantil
    sIdentificacaoReg As String
    sIdentificacaoArq As String
    sLiteralRemessa As String
    sCodTipoServico As String
    sLiteralServico As String
    sBrancosCNAB1 As String
    sAgencia As String
    sCliente As String
    sBrancosCNAB2 As String
    sNomeEmpresa As String
    sNumeroBanco As String
    sNomeBanco As String
    sDataGravacao As String
    sBrancosCNAB3 As String
    sDensidade As String
    sLiteralDensidade As String
    sSequencialArquivo As String
    sSequencialRegistro As String
End Type

'Registro de Detalhe do arquivo de remessa titulos a receber
Public Type typeRemDetalheMercantil
    sIdentificacaoReg As String
    
    'campos do layout antigo
    sTipoInscricaoEmpSac As String
    sNumInscricaoEmpSac As String
    sAgencia As String
    sContaCorrente As String
    'fim
    
    sIndicadorMulta As String
    sCodMulta As String
    sValorPerAplicado As String
    sDataMulta As String
    sFiller1 As String
    
    sNumeroContrato As String
    sUsoDaEmpresa As String
    sAgenciaOrigem As String
    sNossoNumero As String
    sBrancosCNAB1 As String
    sCGCCedente As String
    sQtdMoeda As String
    sCodOperacao As String
    sCodOcorrencia As String
    sSeuNumero As String
    sVencimento As String
    sValorTitulo As String
    sNumeroBanco As String
    sAgenciaCobradora As String
    sEspecie As String
    sAceite As String
    sDataEmissao As String
    sInstrucao1 As String
    sInstrucao2 As String
    sJurosDiarios As String
    sDescontoAte As String
    sValorDesconto As String
    sValorIOF As String
    sAbatimento As String
    sCodigoInscSacado As String
    sNumInscricaoSacado As String
    sNome As String
    sEndereco As String
    sBairro As String
    sCEP As String
    sCidade As String
    sEstado As String
    sMensagem As String
    sBrancos3 As String
    sCodigoMoeda As String
    sNumSequencial As String
End Type

'Registro de Trailer do arquivo de remessa titulos a receber
Type typeRemTrailerMercantil
    sIdentificacaoReg As String
    sComplemento1 As String
    sSequencialRegistro As String
End Type

'------------------ Retorno de Títulos a Receber -------------------------

'Registro de Header do arquivo de retorno titulos a receber
Public Type typeRetHeaderMercantil
    sIdentificacaoReg As String
    sIdentificacaoArq As String
    sLiteralRetorno As String
    sCodTipoServico As String
    sLiteralServico As String
    sAgenciaOrigem As String
    sCGCEmpresa As String
    sNomeEmpresa As String
    sNumeroBanco As String
    sNomeBanco As String
    sDataEmissaoArq As String
    sAgenciaOrigem2 As String
    sNumeroContrato As String
    sNumSequencialArq As String
    sSequencialRegistro As String
End Type

'Registro de detalhe do arquivo de retorno de titulos a receber
Public Type typeRetDetalheMercantil
    sIdentificacaoReg As String
    sTipoInscricaoEmpresa As String
    sNumInscricaoEmpresa As String
    sAgenciaOrigem As String
    sContaCorrente As String
    sNumeroContrato As String
    sUsoDaEmpresa As String
    sNossoNumero1 As String
    sCodCarteira As String
    sCodOcorrencia As String
    sDataOcorrencia As String
    sSeuNumero As String
    sVencimento As String
    sValorTitulo As String
    sNumeroBanco As String
    sAgenciaCobradora As String
    sEspecie As String
    sTarifaCobranca As String
    sOutrasDespesas As String
    sJuros As String
    sValorIOF As String
    sAbatimento As String
    sDescontos As String
    sValorPrincipal As String
    sJurosMora As String
    sOutrosCreditos As String
    sDataCredito As String
    sIndicadorMora  As String
    sNomeSacado As String
    sTaxaPermanencia As String
    sDataLimite As String
    sDescontoLimite As String
    sInstrucao1 As String
    sInstrucao2 As String
    sQtdMoeda As String
    sCodRejeicao As String
    sProtesto As String
    sMoeda As String
    sNumSequencial As String
End Type

'Registro de trailer do arquivo de retorno titulos a receber
Public Type typeRetTrailerMercantil
    sIdentificacaoReg As String
    sCodigoRetorno As String
    sCodigoServico As String
    sNumeroBanco As String
    sBrancos1 As String
    sQuantTitulos As String
    sValorTotal As String
    sAvisoBancario As String
    sSequencialRegistro As String
End Type



'============================ SAFRA 400 ===================================

'------------------ Remessa de Titulos a Receber --------------------

'Registro de Header do arquivo de remessa titulos a receber
Public Type typeRemHeaderSafra
    sIdentificacaoReg As String
    sIdentificacaoArq As String
    sLiteralRemessa As String
    sCodTipoServico As String
    sLiteralServico As String
    sBrancosCNAB1 As String
    sCodigoEmpresa As String
    sBrancos1 As String
    sAgencia As String
    sNomeEmpresa As String
    sNumeroBanco As String
    sNomeBanco As String
    sBrancosCNAB2 As String
    sDataGravacao As String
    sDensidade As String
    sBrancos2 As String
    sSequencialArquivo As String
    sSequencialRegistro As String
End Type

'Registro de Detalhe do arquivo de remessa titulos a receber
Public Type typeRemDetalheSafra
    sIdentificacaoReg As String
    sTipoInscricaoEmpSac As String
    sNumInscricaoEmpSac As String
    sIdentificacaoEmp As String
    sBrancos1 As String
    sUsoDaEmpresa As String
    sNossoNumero As String
    sBrancos2 As String
    sCodigoIOF As String
    sDiasDevolucao As String
    sUsoDoBanco As String
    sCodCarteira As String
    sCodOcorrencia As String
    sSeuNumero As String
    sVencimento As String
    sValorTitulo As String
    sNumeroBanco As String
    sAgenciaCobradora As String
    sEspecie As String
    sAceite As String
    sDataEmissao As String
    sInstrucao1 As String
    sInstrucao2 As String
    sJurosDiarios As String
    sDescontoAte As String
    sValorDesconto As String
    sValorIOF As String
    sAbatimento As String
    sCodigoInscSacado As String
    sNumInscricaoSacado As String
    sNome As String
    sEndereco As String
    sBairro As String
    sBrancosCNAB2 As String
    sCEP As String
    sCidade As String
    sEstado As String
    sMensagem As String
    sBrancos3 As String
    sBancoEmitente As String
    sCodigoMoeda As String
    sNumSeqArquivo As String
    sNumSequencial As String
End Type

'Registro de Trailer do arquivo de remessa titulos a receber
Type typeRemTrailerSafra
    sIdentificacaoReg As String
    sComplemento1 As String
    sSequencialRegistro As String
End Type

'------------------ Retorno de Títulos a Receber -------------------------

'Registro de Header do arquivo de retorno titulos a receber
Public Type typeRetHeaderSafra
    sIdentificacaoReg As String
    sIdentificacaoArq As String
    sLiteralRetorno As String
    sCodTipoServico As String
    sLiteralServico As String
    sCodEmpresa As String
    sNomeEmpresa As String
    sNumeroBanco As String
    sNomeBanco As String
    sDataEmissaoArq As String
    sDensidade As String
    sLiteralDensidade As String
    sNumSequencialArq As String
    sBrancos As String
    sSequencialRegistro As String
End Type

'Registro de detalhe do arquivo de retorno de titulos a receber
Public Type typeRetDetalheSafra
    sIdentificacaoReg As String
    sTipoInscricaoEmpresa As String
    sNumInscricaoEmpresa As String
    sCodEmpresa As String
    sBrancos1 As String
    sUsoDaEmpresa As String
    sNossoNumero1 As String
    sBrancos2 As String
    sCodCarteira As String
    sCodOcorrencia As String
    sDataOcorrencia As String
    sSeuNumero As String
    sNossoNumero2 As String
    sCodRejeicao As String
    sBrancos3 As String
    sVencimento As String
    sValorTitulo As String
    sNumeroBanco As String
    sAgenciaCobradora As String
    sEspecie As String
    sTarifaCobranca As String
    sOutrasDespesas As String
    sDescontos As String
    sValorIOF As String
    sOutrasDeducoes As String
    sValorPrincipal As String
    sJuros As String
    sOutrosCreditos As String
    sDataCredito As String
    sValorLiquido  As String
    sNomeSacado As String
    sMoeda As String
    sNumSequencial As String
End Type

'Registro de trailer do arquivo de retorno titulos a receber
Public Type typeRetTrailerSafra
    sIdentificacaoReg As String
    sCodigoRetorno As String
    sCodigoServico As String
    sNumeroBanco As String
    sBrancos1 As String
    sQuantTitulos As String
    sValorTotal As String
    sAvisoBancario As String
    sDataCredito As String
    sBrancos2 As String
    sSequencialRegistro As String
End Type

'=============================== CEDULA ===================================

'------------------ Remessa de Titulos a Receber --------------------

'Registro de Header do arquivo de remessa titulos a receber
Public Type typeRemHeaderCedula
    sIdentificacaoReg As String
    sIdentificacaoArq As String
    sLiteralRemessa As String
    sCodTipoServico As String
    sLiteralServico As String
    sCodigoEmpresa As String
    sNomeEmpresa As String
    sNumeroBanco As String
    sNomeBanco As String
    sDataEmissaoArq As String
    sComplemento1 As String
    sIdentificacaoSistema As String
    sSequencialArq As String
    sComplemento2 As String
    sSequencialRegistro As String
End Type

'Registro de Detalhe do arquivo de remessa titulos a receber
Public Type typeRemDetalheCedula
    sIdentificacaoReg As String
    sAgenciaDebito As String
    sDigitoAgDebito As String
    sRazaoContaCorrente As String
    sContaCorrente As String
    sDigitoContaCorrente As String
    sIdentifEmpresa As String
    sNumControle As String
    sCodigoBancoDebitado As String
    sZeros As String
    sNossoNumero As String
    sDesconto As String
    sCondEmissaoPapel As String
    sIdentEmitePapel As String
    sIdentOpBanco As String
    sIndicadorRateio As String
    sEnderecamento As String
    sBranco As String
    sIdentificacaoOcorr As String
    sNumDocto As String
    sDataVencto As String
    sValorTitulo As String
    sBancoCobranca As String
    sAgenciaDepositaria As String
    sEspecie As String
    sAceite As String
    sDataEmissao As String
    sInstrucao1 As String
    sInstrucao2 As String
    sValorJuros As String
    sDataDesconto As String
    sValorDesconto As String
    sValorIOF As String
    sAbatimento As String
    sIdentfInscSacado As String
    sNumInscricaoSacado As String
    sNomeSacado As String
    sEnderecoSacado As String
    sMensagem1 As String
    sCEP As String
    sCEPSufixo As String
    sEstado As String
    sSacadorAvalista As String
    sNumSequencialRegistro As String

End Type

'Registro de Trailer do arquivo de remessa titulos a receber
Type typeRemTrailerCedula
    sIdentificacaoReg As String
    sComplemento1 As String
    sSequencialRegistro As String
End Type

'------------------ Retorno de Titulos a Receber --------------------

'Registro de Header do arquivo de retorno titulos a receber
Public Type typeRetHeaderCedula
    sIdentificacaoReg As String
    sIdentificacaoArq As String
    sLiteralRetorno As String
    sCodTipoServico As String
    sLiteralServico As String
    sCodEmpresa As String
    sNomeEmpresa As String
    sNumeroBanco As String
    sNomeBanco As String
    sDataEmissaoArq As String
    sZeros As String
    sNumAviso As String
    sComplemento1 As String
    sDataCredito As String
    sComplemento2 As String
    sSequencialRegistro As String
End Type

'Registro de detalhe do arquivo de retorno de titulos a receber
Public Type typeRetDetalheCedula
    sIdentificacaoReg As String
    sTipoInscricaoEmpresa As String
    sNumInscricaoEmpresa As String
    sZeros1 As String
    sIdentfEmpresa As String
    sNumControle As String
    sZeros2 As String
    sNossoNumero1 As String
    sUsoDoBanco1 As String
    sUsoDoBanco2 As String
    sIndicadorRateio As String
    sZeros As String
    sCodCarteira As String
    sCodOcorrencia As String
    sDataOcorrencia As String
    sNumDocto As String
    sNossoNumero2 As String
    sVencimento As String
    sValorTitulo As String
    sBancoCobrador As String
    sAgenciaCobradora As String
    sEspecie As String
    sTarifaCobranca As String
    sOutrasDespesas As String
    sJurosOp As String
    sValorIOF As String
    sAbatimento As String
    sDescontos As String
    sValorPrincipal As String
    sJuros As String
    sOutrosCreditos As String
    sComplemento1 As String
    sMotivoOcorr1 As String
    sDataCredito As String
    sComplemento2 As String
    sMotivoOcorr2 As String
    sComplemento3 As String
    sNumSequencialRegistro As String
End Type

'Registro de trailer do arquivo de retorno titulos a receber
Public Type typeRetTrailerCedula
    sIdentificacaoReg As String
    sCodigoRetorno As String
    sCodigoTipoRegistro As String
    sNumeroBanco As String
    sComplemento1 As String
    sQuantTitulosCobr As String
    sValorTotalCobr As String
    sAvisoBancario As String
    sComplemento2 As String
    sQuantRegOcorr_02 As String
    sValorRegOcorr_02 As String
    sValorRegOcorr_06 As String
    sQuantRegOcorr_06 As String
    sValorRegOcorr_06_09_10 As String
    sQuantRegOcorr_09_10 As String
    sValorRegOcorr_09_10 As String
    sQuantRegOcorr_13 As String
    sValorRegOcorr_13 As String
    sQuantRegOcorr_14 As String
    sValorRegOcorr_14 As String
    sQuantRegOcorr_12 As String
    sValorRegOcorr_12 As String
    sQuantRegOcorr_19 As String
    sValorRegOcorr_19 As String
    sComplemento3 As String
    sValorTotalRateios As String
    sQuantTotalRateios As String
    sSequencialRegistro As String
    
End Type

'-------------------- Remessa de Contas a Pagar CNAB 240 -----------------------------

'versao 040
'Registro de Header de Arquivo de Remessa/Retorno de títulos a pagar
Public Type typePagto240HeaderArq
    sNumeroBanco As String
    sCodigoLote As String
    sTipoRegistro As String
    sComplemento1 As String
    sLayoutArq As String
    sTipoInscEmpresa As String
    sCGCEmpresaDebitada As String
    sConvenio As String
    sAgencia As String
    sDVAgencia As String
    sConta As String
    sDVNumConta As String
    sDVAgConta As String
    sNomeEmpresa As String
    sNomeBanco As String
    sComplemento2 As String
    sArquivoCodigo As String
    sDataGeracao As String
    sHoraGeracao As String
    sUnidadeDensidade As String
    sComplemento3 As String
    sSequencialArquivo As String
End Type

'versao 030
Public Type typePagto240HeaderLote
    sNumeroBanco As String
    sCodigoLote As String
    sTipoRegistro As String
    sTipoOperacao As String
    sTipoPagamento As String
    sFormaPagamento As String
    sLayoutLote As String
    sComplemento1 As String
    sTipoInscEmpresa As String
    sCGCEmpresaDebitada As String
    sConvenio As String
    sAgencia As String
    sDVAgencia As String
    sConta As String
    sDVNumConta As String
    sDVAgConta As String
    sNomeEmpresa As String
    sMensagem As String
    sEnderecoEmpresa As String
    sNumeroLocal As String
    sComplementoEndereco As String
    sCidade As String
    sCEP As String
    sEstado As String
    sComplemento2 As String
    sOcorrenciasRetorno As String
End Type

Public Type typePagto240SegmentoA
    sNumeroBanco As String
    sCodigoLote As String
    sTipoRegistro As String
    sNumeroRegistro As String
    sSegmento As String
    sTipoMovimento As String
    sCodigoInstrucaoMovto As String
    sCodCamaraCentralizadora As String
    sBancoFavorecido As String
    sAgenciaConta As String
    sNomeFavorecido As String
    sSeuNumero As String
    sDataPagamento As String
    sTipoMoeda As String
    sQtdeMoeda As String
    sValorPagamento As String
    sNossoNumero As String
    sDataEfetiva As String
    sValorEfetivo As String
    sMensagem2 As String
    sFinalidadeDetalhe As String
    sComplemento1 As String
    sAviso As String
    sOcorrenciasRetorno As String
End Type

Public Type typePagto240SegmentoB
    sNumeroBanco As String
    sCodigoLote As String
    sTipoRegistro As String
    sNumeroRegistro As String
    sSegmento As String
    sComplemento1 As String
    sTipoInscEmpresa As String
    sNumInscFavorecido As String
    sEndereco As String
    sNumeroLocal As String
    sComplementoEndereco As String
    sBairro As String
    sCidade As String
    sCEP As String
    sEstado As String
    sDataVencimento As String
    sValorDocumento As String
    sValorAbatimento As String
    sValorDesconto As String
    sValorMora As String
    sValorMulta As String
    sCodDocFavorecido As String
    sAvisoFavorecido As String
    sComplemento2 As String
End Type
    
Public Type typePagto240TrailerLote
    sNumeroBanco As String
    sCodigoLote As String
    sTipoRegistro As String
    sComplemento1 As String
    sTotalQtdRegistros As String
    sTotalValorPagtos As String
    sZeros As String
    sComplemento2 As String
    sOcorrenciasRetorno As String
End Type

Public Type typePagto240SegmentoJ
    sNumeroBanco As String
    sCodigoLote As String
    sTipoRegistro As String
    sNumeroRegistro As String
    sSegmento As String
    sTipoMovimento As String
    sCodigoInstrucaoMovto As String
    sBancoFavorecido As String
    sMoeda As String
    sDV As String
    sValor As String
    sCampoLivre As String
    sNomeFavorecido As String
    sDataVencto As String
    sValorTitulo As String
    sDescontos As String
    sAcrescimos As String
    sDataPagamento As String
    sValorPagamento As String
    sZeros As String
    sSeuNumero As String
    sComplemento As String
    sNossoNumero As String
    sOcorrenciasRetorno As String
    
End Type
    
Public Type typePagto240SegmentoJ52
    sNumeroBanco As String
    sCodigoLote As String
    sTipoRegistro As String
    sNumeroRegistro As String
    sSegmento As String
    sFiller As String
    sTipoMovimento As String
    sIdentificacaoRegistroOpcional As String
    sSacadoTipoInscricao As String
    sSacadoCNPJ As String
    sSacadoNome As String
    sCedenteTipoInscricao As String
    sCedenteCNPJ As String
    sCedenteNome As String
    sSacadorTipoInscricao As String
    sSacadorCNPJ As String
    sSacadorNome As String
    sFiller2 As String
End Type
    

Public Type typePagto240TrailerArq
    sNumeroBanco As String
    sCodigoLote As String
    sTipoRegistro As String
    sComplemento1 As String
    sTotalQtdLotes As String
    sTotalQtdRegistros As String
    sTotalContasConc As String
    sComplemento2 As String
End Type

Public Type typePagto240SegmentoO
    sNumeroBanco As String
    sCodigoLote As String
    sTipoRegistro As String
    sNumeroRegistro As String
    sSegmento As String
    sTipoMovimento As String
    sCodigoInstrucaoMovto As String
    sCodigoBarras As String
    sNomeConcessionaria As String
    sDataVencto As String
    sDataPagamento As String
    sValorPagamento As String
    sNumDoctoEmpresa As String
    sNumDoctoBanco As String
    sCNAB As String
    sOcorrenciasRetorno As String
    
End Type

'-------------------- FIM da Remessa de Contas a Pagar CNAB 240 -----------------------------


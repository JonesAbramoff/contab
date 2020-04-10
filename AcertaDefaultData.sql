CREATE DEFAULT [FORPRINT_DATA_NULA] AS {d '1822-09-07'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[AdmExtFinAmexArq].[DataCriacao]' 
GO

UPDATE AdmExtFinAmexArq SET DataCriacao = {d '1822-09-07'} WHERE DataCriacao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[AdmExtFinAmexSOC].[DataPagto]' 
GO

UPDATE AdmExtFinAmexSOC SET DataPagto = {d '1822-09-07'} WHERE DataPagto = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[AdmExtFinAmexSOC].[DataVenda]' 
GO

UPDATE AdmExtFinAmexSOC SET DataVenda = {d '1822-09-07'} WHERE DataVenda = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[AdmExtFinAmexSOC].[DataPagto24hs]' 
GO

UPDATE AdmExtFinAmexSOC SET DataPagto24hs = {d '1822-09-07'} WHERE DataPagto24hs = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[AdmExtFinArqsLidos].[DataImportacao]' 
GO

UPDATE AdmExtFinArqsLidos SET DataImportacao = {d '1822-09-07'} WHERE DataImportacao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[AdmExtFinArqsLidos].[DataAtualizado]' 
GO

UPDATE AdmExtFinArqsLidos SET DataAtualizado = {d '1822-09-07'} WHERE DataAtualizado = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[AdmExtFinCieloArq].[DataProcessamento]' 
GO

UPDATE AdmExtFinCieloArq SET DataProcessamento = {d '1822-09-07'} WHERE DataProcessamento = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[AdmExtFinCieloRO].[DtDeposito]' 
GO

UPDATE AdmExtFinCieloRO SET DtDeposito = {d '1822-09-07'} WHERE DtDeposito = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[AdmExtFinCieloRO].[DtPrevPag]' 
GO

UPDATE AdmExtFinCieloRO SET DtPrevPag = {d '1822-09-07'} WHERE DtPrevPag = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[AdmExtFinCieloRO].[DtEnvBco]' 
GO

UPDATE AdmExtFinCieloRO SET DtEnvBco = {d '1822-09-07'} WHERE DtEnvBco = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[AdmExtFinMov].[Data]' 
GO

UPDATE AdmExtFinMov SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[AdmExtFinMovDet].[DataCompra]' 
GO

UPDATE AdmExtFinMovDet SET DataCompra = {d '1822-09-07'} WHERE DataCompra = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[AdmExtFinVisanetArq].[DataProcessamento]' 
GO

UPDATE AdmExtFinVisanetArq SET DataProcessamento = {d '1822-09-07'} WHERE DataProcessamento = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[AdmExtfinVisanetRO].[DtDeposito]' 
GO

UPDATE AdmExtfinVisanetRO SET DtDeposito = {d '1822-09-07'} WHERE DtDeposito = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[AdmExtfinVisanetRO].[DtPrevPag]' 
GO

UPDATE AdmExtfinVisanetRO SET DtPrevPag = {d '1822-09-07'} WHERE DtPrevPag = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[AdmExtfinVisanetRO].[DtEnvBco]' 
GO

UPDATE AdmExtfinVisanetRO SET DtEnvBco = {d '1822-09-07'} WHERE DtEnvBco = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[AdmMeioPagto].[DataLog]' 
GO

UPDATE AdmMeioPagto SET DataLog = {d '1822-09-07'} WHERE DataLog = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[AdmMeioPagtoCondPagto].[DataLog]' 
GO

UPDATE AdmMeioPagtoCondPagto SET DataLog = {d '1822-09-07'} WHERE DataLog = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Anotacoes].[DataAlteracao]' 
GO

UPDATE Anotacoes SET DataAlteracao = {d '1822-09-07'} WHERE DataAlteracao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Aplicacoes].[DataAplicacao]' 
GO

UPDATE Aplicacoes SET DataAplicacao = {d '1822-09-07'} WHERE DataAplicacao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Aplicacoes].[DataBaixa]' 
GO

UPDATE Aplicacoes SET DataBaixa = {d '1822-09-07'} WHERE DataBaixa = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Aplicacoes].[DataResgatePrevista]' 
GO

UPDATE Aplicacoes SET DataResgatePrevista = {d '1822-09-07'} WHERE DataResgatePrevista = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ApontamentoProducao].[Data]' 
GO

UPDATE ApontamentoProducao SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ApontamentoSRV].[Data]' 
GO

UPDATE ApontamentoSRV SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ApontPRJ].[Data]' 
GO

UPDATE ApontPRJ SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ApontPRJ].[DataRegistro]' 
GO

UPDATE ApontPRJ SET DataRegistro = {d '1822-09-07'} WHERE DataRegistro = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ApontPRJ].[DataUltAlt]' 
GO

UPDATE ApontPRJ SET DataUltAlt = {d '1822-09-07'} WHERE DataUltAlt = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ArqExportacao].[DataExportacao]' 
GO

UPDATE ArqExportacao SET DataExportacao = {d '1822-09-07'} WHERE DataExportacao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ArqExportacaoAux].[DataGeracao]' 
GO

UPDATE ArqExportacaoAux SET DataGeracao = {d '1822-09-07'} WHERE DataGeracao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ArqExportacaoAux].[ExpDataDe]' 
GO

UPDATE ArqExportacaoAux SET ExpDataDe = {d '1822-09-07'} WHERE ExpDataDe = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ArqExportacaoAux].[ExpDataAte]' 
GO

UPDATE ArqExportacaoAux SET ExpDataAte = {d '1822-09-07'} WHERE ExpDataAte = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ArqImportacao].[DataImportacao]' 
GO

UPDATE ArqImportacao SET DataImportacao = {d '1822-09-07'} WHERE DataImportacao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ArqImportacao].[DataAtualizacao]' 
GO

UPDATE ArqImportacao SET DataAtualizacao = {d '1822-09-07'} WHERE DataAtualizacao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[AvisoWFW].[Data]' 
GO

UPDATE AvisoWFW SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[AvisoWFW].[DataUltAviso]' 
GO

UPDATE AvisoWFW SET DataUltAviso = {d '1822-09-07'} WHERE DataUltAviso = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[BaixaBoletoParcela].[DataCancelamento]' 
GO

UPDATE BaixaBoletoParcela SET DataCancelamento = {d '1822-09-07'} WHERE DataCancelamento = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[BaixaBoletosParcelas].[Data]' 
GO

UPDATE BaixaBoletosParcelas SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[BaixaBoletosParcelas].[DataContabil]' 
GO

UPDATE BaixaBoletosParcelas SET DataContabil = {d '1822-09-07'} WHERE DataContabil = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[BaixaBoletosParcelas].[DataRegistro]' 
GO

UPDATE BaixaBoletosParcelas SET DataRegistro = {d '1822-09-07'} WHERE DataRegistro = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[BaixaCreditosPagForn].[Data]' 
GO

UPDATE BaixaCreditosPagForn SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[BaixaDebitosRecCli].[Data]' 
GO

UPDATE BaixaDebitosRecCli SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[BaixaPagAntecipados].[Data]' 
GO

UPDATE BaixaPagAntecipados SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[BaixaRecebAntecipados].[Data]' 
GO

UPDATE BaixaRecebAntecipados SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[BaixasAgrupadas].[DataBaixa]' 
GO

UPDATE BaixasAgrupadas SET DataBaixa = {d '1822-09-07'} WHERE DataBaixa = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[BaixasCarne].[DataBaixa]' 
GO

UPDATE BaixasCarne SET DataBaixa = {d '1822-09-07'} WHERE DataBaixa = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[BaixasPag].[Data]' 
GO

UPDATE BaixasPag SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[BaixasPag].[DataContabil]' 
GO

UPDATE BaixasPag SET DataContabil = {d '1822-09-07'} WHERE DataContabil = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[BaixasPag].[DataRegistro]' 
GO

UPDATE BaixasPag SET DataRegistro = {d '1822-09-07'} WHERE DataRegistro = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[BaixasParcPag].[DataCancelamento]' 
GO

UPDATE BaixasParcPag SET DataCancelamento = {d '1822-09-07'} WHERE DataCancelamento = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[BaixasParcRec].[DataCancelamento]' 
GO

UPDATE BaixasParcRec SET DataCancelamento = {d '1822-09-07'} WHERE DataCancelamento = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[BaixasParcRec].[DataRegCancelamento]' 
GO

UPDATE BaixasParcRec SET DataRegCancelamento = {d '1822-09-07'} WHERE DataRegCancelamento = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[BaixasParcRecCanc].[DataCancelamento]' 
GO

UPDATE BaixasParcRecCanc SET DataCancelamento = {d '1822-09-07'} WHERE DataCancelamento = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[BaixasRec].[Data]' 
GO

UPDATE BaixasRec SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[BaixasRec].[DataContabil]' 
GO

UPDATE BaixasRec SET DataContabil = {d '1822-09-07'} WHERE DataContabil = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[BaixasRec].[DataRegistro]' 
GO

UPDATE BaixasRec SET DataRegistro = {d '1822-09-07'} WHERE DataRegistro = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[BaixaValesTickets].[Data]' 
GO

UPDATE BaixaValesTickets SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[BaixaValesTickets].[DataContabil]' 
GO

UPDATE BaixaValesTickets SET DataContabil = {d '1822-09-07'} WHERE DataContabil = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[BaixaValesTickets].[DataRegistro]' 
GO

UPDATE BaixaValesTickets SET DataRegistro = {d '1822-09-07'} WHERE DataRegistro = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[BaixaValeTicket].[DataCancelamento]' 
GO

UPDATE BaixaValeTicket SET DataCancelamento = {d '1822-09-07'} WHERE DataCancelamento = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Bancos].[DataLog]' 
GO

UPDATE Bancos SET DataLog = {d '1822-09-07'} WHERE DataLog = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[BloqueiosGen].[Data]' 
GO

UPDATE BloqueiosGen SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[BloqueiosGen].[DataLib]' 
GO

UPDATE BloqueiosGen SET DataLib = {d '1822-09-07'} WHERE DataLib = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[BloqueiosPC].[Data]' 
GO

UPDATE BloqueiosPC SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[BloqueiosPC].[DataLib]' 
GO

UPDATE BloqueiosPC SET DataLib = {d '1822-09-07'} WHERE DataLib = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[BloqueiosPedidoSRV].[Data]' 
GO

UPDATE BloqueiosPedidoSRV SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[BloqueiosPedidoSRV].[DataLib]' 
GO

UPDATE BloqueiosPedidoSRV SET DataLib = {d '1822-09-07'} WHERE DataLib = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[BloqueiosPV].[Data]' 
GO

UPDATE BloqueiosPV SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[BloqueiosPV].[DataLib]' 
GO

UPDATE BloqueiosPV SET DataLib = {d '1822-09-07'} WHERE DataLib = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[BloqueiosPVBaixados].[Data]' 
GO

UPDATE BloqueiosPVBaixados SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[BloqueiosPVBaixados].[DataLib]' 
GO

UPDATE BloqueiosPVBaixados SET DataLib = {d '1822-09-07'} WHERE DataLib = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Boleto].[DataTransacao]' 
GO

UPDATE Boleto SET DataTransacao = {d '1822-09-07'} WHERE DataTransacao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[BoletoParcela].[DataVencimento]' 
GO

UPDATE BoletoParcela SET DataVencimento = {d '1822-09-07'} WHERE DataVencimento = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[BorderoBoleto].[DataImpressao]' 
GO

UPDATE BorderoBoleto SET DataImpressao = {d '1822-09-07'} WHERE DataImpressao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[BorderoBoleto].[DataEnvio]' 
GO

UPDATE BorderoBoleto SET DataEnvio = {d '1822-09-07'} WHERE DataEnvio = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[BorderoBoleto].[DataBackoffice]' 
GO

UPDATE BorderoBoleto SET DataBackoffice = {d '1822-09-07'} WHERE DataBackoffice = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[BorderoBoletoItem].[DataPreDatado]' 
GO

UPDATE BorderoBoletoItem SET DataPreDatado = {d '1822-09-07'} WHERE DataPreDatado = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[BorderoCheque].[DataImpressao]' 
GO

UPDATE BorderoCheque SET DataImpressao = {d '1822-09-07'} WHERE DataImpressao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[BorderoCheque].[DataEnvio]' 
GO

UPDATE BorderoCheque SET DataEnvio = {d '1822-09-07'} WHERE DataEnvio = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[BorderoCheque].[DataBackoffice]' 
GO

UPDATE BorderoCheque SET DataBackoffice = {d '1822-09-07'} WHERE DataBackoffice = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[BorderoDescChq].[DataEmissao]' 
GO

UPDATE BorderoDescChq SET DataEmissao = {d '1822-09-07'} WHERE DataEmissao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[BorderoDescChq].[DataContabil]' 
GO

UPDATE BorderoDescChq SET DataContabil = {d '1822-09-07'} WHERE DataContabil = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[BorderoDescChq].[DataDeposito]' 
GO

UPDATE BorderoDescChq SET DataDeposito = {d '1822-09-07'} WHERE DataDeposito = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[BorderoOutros].[DataEnvio]' 
GO

UPDATE BorderoOutros SET DataEnvio = {d '1822-09-07'} WHERE DataEnvio = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[BorderoOutros].[DataImpressao]' 
GO

UPDATE BorderoOutros SET DataImpressao = {d '1822-09-07'} WHERE DataImpressao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[BorderoOutros].[DataBackoffice]' 
GO

UPDATE BorderoOutros SET DataBackoffice = {d '1822-09-07'} WHERE DataBackoffice = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[BorderosChequesPre].[DataEmissao]' 
GO

UPDATE BorderosChequesPre SET DataEmissao = {d '1822-09-07'} WHERE DataEmissao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[BorderosChequesPre].[DataDeposito]' 
GO

UPDATE BorderosChequesPre SET DataDeposito = {d '1822-09-07'} WHERE DataDeposito = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[BorderosChequesPre].[DataContabil]' 
GO

UPDATE BorderosChequesPre SET DataContabil = {d '1822-09-07'} WHERE DataContabil = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[BorderosCobranca].[DataEmissao]' 
GO

UPDATE BorderosCobranca SET DataEmissao = {d '1822-09-07'} WHERE DataEmissao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[BorderosCobranca].[DataCancelamento]' 
GO

UPDATE BorderosCobranca SET DataCancelamento = {d '1822-09-07'} WHERE DataCancelamento = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[BorderosCobranca].[DataContabilCancelamento]' 
GO

UPDATE BorderosCobranca SET DataContabilCancelamento = {d '1822-09-07'} WHERE DataContabilCancelamento = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[BorderosPagto].[DataEmissao]' 
GO

UPDATE BorderosPagto SET DataEmissao = {d '1822-09-07'} WHERE DataEmissao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[BorderosPagto].[DataEnvio]' 
GO

UPDATE BorderosPagto SET DataEnvio = {d '1822-09-07'} WHERE DataEnvio = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[BorderosPagto].[DataVencimento]' 
GO

UPDATE BorderosPagto SET DataVencimento = {d '1822-09-07'} WHERE DataVencimento = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[BorderosRetCobr].[DataRecepcao]' 
GO

UPDATE BorderosRetCobr SET DataRecepcao = {d '1822-09-07'} WHERE DataRecepcao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[BorderoValeTicket].[DataEnvio]' 
GO

UPDATE BorderoValeTicket SET DataEnvio = {d '1822-09-07'} WHERE DataEnvio = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[BorderoValeTicket].[DataImpressao]' 
GO

UPDATE BorderoValeTicket SET DataImpressao = {d '1822-09-07'} WHERE DataImpressao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[BorderoValeTicket].[DataBackoffice]' 
GO

UPDATE BorderoValeTicket SET DataBackoffice = {d '1822-09-07'} WHERE DataBackoffice = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Caixa].[DataInicial]' 
GO

UPDATE Caixa SET DataInicial = {d '1822-09-07'} WHERE DataInicial = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Caixa].[DataLog]' 
GO

UPDATE Caixa SET DataLog = {d '1822-09-07'} WHERE DataLog = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[CamposCustomizados].[Data1]' 
GO

UPDATE CamposCustomizados SET Data1 = {d '1822-09-07'} WHERE Data1 = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[CamposCustomizados].[Data2]' 
GO

UPDATE CamposCustomizados SET Data2 = {d '1822-09-07'} WHERE Data2 = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[CamposCustomizados].[Data3]' 
GO

UPDATE CamposCustomizados SET Data3 = {d '1822-09-07'} WHERE Data3 = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[CamposCustomizados].[Data4]' 
GO

UPDATE CamposCustomizados SET Data4 = {d '1822-09-07'} WHERE Data4 = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[CamposCustomizados].[Data5]' 
GO

UPDATE CamposCustomizados SET Data5 = {d '1822-09-07'} WHERE Data5 = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Carne].[DataReferencia]' 
GO

UPDATE Carne SET DataReferencia = {d '1822-09-07'} WHERE DataReferencia = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[CarneParcelas].[DataVencimento]' 
GO

UPDATE CarneParcelas SET DataVencimento = {d '1822-09-07'} WHERE DataVencimento = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[CCIMovDia].[Data]' 
GO

UPDATE CCIMovDia SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Ccl].[DataRegistro]' 
GO

UPDATE Ccl SET DataRegistro = {d '1822-09-07'} WHERE DataRegistro = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[CclHistorico].[DataAtualizacao]' 
GO

UPDATE CclHistorico SET DataAtualizacao = {d '1822-09-07'} WHERE DataAtualizacao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[CCMovDia].[Data]' 
GO

UPDATE CCMovDia SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ChequePreN].[DataDeposito]' 
GO

UPDATE ChequePreN SET DataDeposito = {d '1822-09-07'} WHERE DataDeposito = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ChequePreN].[DataEmissao]' 
GO

UPDATE ChequePreN SET DataEmissao = {d '1822-09-07'} WHERE DataEmissao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ChequePrePag].[DataEmissao]' 
GO

UPDATE ChequePrePag SET DataEmissao = {d '1822-09-07'} WHERE DataEmissao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ChequePrePag].[DataBomPara]' 
GO

UPDATE ChequePrePag SET DataBomPara = {d '1822-09-07'} WHERE DataBomPara = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ChequePrePag].[DataDeposito]' 
GO

UPDATE ChequePrePag SET DataDeposito = {d '1822-09-07'} WHERE DataDeposito = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ClassificacaoABC].[Data]' 
GO

UPDATE ClassificacaoABC SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ClienteContatos].[DataNasc]' 
GO

UPDATE ClienteContatos SET DataNasc = {d '1822-09-07'} WHERE DataNasc = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ClienteHistorico].[DataAtualizacao]' 
GO

UPDATE ClienteHistorico SET DataAtualizacao = {d '1822-09-07'} WHERE DataAtualizacao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ClienteHistorico].[DataReg]' 
GO

UPDATE ClienteHistorico SET DataReg = {d '1822-09-07'} WHERE DataReg = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[CobrancaContrato].[DataCobrIni]' 
GO

UPDATE CobrancaContrato SET DataCobrIni = {d '1822-09-07'} WHERE DataCobrIni = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[CobrancaContrato].[DataCobrFim]' 
GO

UPDATE CobrancaContrato SET DataCobrFim = {d '1822-09-07'} WHERE DataCobrFim = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[CobrancaContrato].[DataGeracao]' 
GO

UPDATE CobrancaContrato SET DataGeracao = {d '1822-09-07'} WHERE DataGeracao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[CobrancaContrato].[DataEmissao]' 
GO

UPDATE CobrancaContrato SET DataEmissao = {d '1822-09-07'} WHERE DataEmissao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[CobrancaContrato].[DataRefVencimento]' 
GO

UPDATE CobrancaContrato SET DataRefVencimento = {d '1822-09-07'} WHERE DataRefVencimento = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[CodAtivContrPrev].[DataIniEscrituracao]' 
GO

UPDATE CodAtivContrPrev SET DataIniEscrituracao = {d '1822-09-07'} WHERE DataIniEscrituracao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[CodAtivContrPrev].[DataFimEscrituracao]' 
GO

UPDATE CodAtivContrPrev SET DataFimEscrituracao = {d '1822-09-07'} WHERE DataFimEscrituracao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Comissoes].[DataGeracao]' 
GO

UPDATE Comissoes SET DataGeracao = {d '1822-09-07'} WHERE DataGeracao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Comissoes].[DataBaixa]' 
GO

UPDATE Comissoes SET DataBaixa = {d '1822-09-07'} WHERE DataBaixa = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Comissoes_BKP_20100809].[DataGeracao]' 
GO

UPDATE Comissoes_BKP_20100809 SET DataGeracao = {d '1822-09-07'} WHERE DataGeracao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Comissoes_BKP_20100809].[DataBaixa]' 
GO

UPDATE Comissoes_BKP_20100809 SET DataBaixa = {d '1822-09-07'} WHERE DataBaixa = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ComissoesAvulsas].[Data]' 
GO

UPDATE ComissoesAvulsas SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ConcorrenciaN].[Data]' 
GO

UPDATE ConcorrenciaN SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Configuracao].[DataImplantacao]' 
GO

UPDATE Configuracao SET DataImplantacao = {d '1822-09-07'} WHERE DataImplantacao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ContasCorrentesInternas].[DataSaldoInicial]' 
GO

UPDATE ContasCorrentesInternas SET DataSaldoInicial = {d '1822-09-07'} WHERE DataSaldoInicial = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ContasCorrentesInternas].[DataLog]' 
GO

UPDATE ContasCorrentesInternas SET DataLog = {d '1822-09-07'} WHERE DataLog = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ContratoFornecimento].[Data]' 
GO

UPDATE ContratoFornecimento SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ContratoFornecimento].[DataFim]' 
GO

UPDATE ContratoFornecimento SET DataFim = {d '1822-09-07'} WHERE DataFim = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ContratoFornecimento].[DataUltimoPC]' 
GO

UPDATE ContratoFornecimento SET DataUltimoPC = {d '1822-09-07'} WHERE DataUltimoPC = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Contratos].[DataIniContrato]' 
GO

UPDATE Contratos SET DataIniContrato = {d '1822-09-07'} WHERE DataIniContrato = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Contratos].[DataFimContrato]' 
GO

UPDATE Contratos SET DataFimContrato = {d '1822-09-07'} WHERE DataFimContrato = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Contratos].[DataRenovContrato]' 
GO

UPDATE Contratos SET DataRenovContrato = {d '1822-09-07'} WHERE DataRenovContrato = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Contratos].[DataIniCobrancaPadrao]' 
GO

UPDATE Contratos SET DataIniCobrancaPadrao = {d '1822-09-07'} WHERE DataIniCobrancaPadrao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ControleLogBack].[Data]' 
GO

UPDATE ControleLogBack SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ControleLogCaixaCC].[DataInicio]' 
GO

UPDATE ControleLogCaixaCC SET DataInicio = {d '1822-09-07'} WHERE DataInicio = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ControleLogCaixaCC].[DataFim]' 
GO

UPDATE ControleLogCaixaCC SET DataFim = {d '1822-09-07'} WHERE DataFim = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ControleLogCCBack].[Data]' 
GO

UPDATE ControleLogCCBack SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[CotacaoItemConcorrenciaN].[DataEntrega]' 
GO

UPDATE CotacaoItemConcorrenciaN SET DataEntrega = {d '1822-09-07'} WHERE DataEntrega = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[CotacaoN].[Data]' 
GO

UPDATE CotacaoN SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[CotacoesMoeda].[Data]' 
GO

UPDATE CotacoesMoeda SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[CreditosPagForn].[DataEmissao]' 
GO

UPDATE CreditosPagForn SET DataEmissao = {d '1822-09-07'} WHERE DataEmissao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[CTMaquinaProgDisponibilidade].[Data]' 
GO

UPDATE CTMaquinaProgDisponibilidade SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[CTMaquinaProgTurno].[Data]' 
GO

UPDATE CTMaquinaProgTurno SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[CTMaquinasParadas].[Data]' 
GO

UPDATE CTMaquinasParadas SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[CupomFiscal].[DataEmissao]' 
GO

UPDATE CupomFiscal SET DataEmissao = {d '1822-09-07'} WHERE DataEmissao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[CupomFiscal].[DataReducao]' 
GO

UPDATE CupomFiscal SET DataReducao = {d '1822-09-07'} WHERE DataReducao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[CupomFiscal].[DataEmissaoNF]' 
GO

UPDATE CupomFiscal SET DataEmissaoNF = {d '1822-09-07'} WHERE DataEmissaoNF = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Cursos].[DataInicio]' 
GO

UPDATE Cursos SET DataInicio = {d '1822-09-07'} WHERE DataInicio = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Cursos].[DataConclusao]' 
GO

UPDATE Cursos SET DataConclusao = {d '1822-09-07'} WHERE DataConclusao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[CusteioRoteiroFabricacao].[DataCusteio]' 
GO

UPDATE CusteioRoteiroFabricacao SET DataCusteio = {d '1822-09-07'} WHERE DataCusteio = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[CusteioRoteiroFabricacao].[DataValidade]' 
GO

UPDATE CusteioRoteiroFabricacao SET DataValidade = {d '1822-09-07'} WHERE DataValidade = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[CustoDirFabr].[Data]' 
GO

UPDATE CustoDirFabr SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[CustoDirFabrProd].[Data]' 
GO

UPDATE CustoDirFabrProd SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[CustoEmbMP].[DataAtualizacao]' 
GO

UPDATE CustoEmbMP SET DataAtualizacao = {d '1822-09-07'} WHERE DataAtualizacao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[CustoFixo].[DataReferencia]' 
GO

UPDATE CustoFixo SET DataReferencia = {d '1822-09-07'} WHERE DataReferencia = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[CustoFixo].[DataAtualizacao]' 
GO

UPDATE CustoFixo SET DataAtualizacao = {d '1822-09-07'} WHERE DataAtualizacao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[CustoFixoProd].[DataReferencia]' 
GO

UPDATE CustoFixoProd SET DataReferencia = {d '1822-09-07'} WHERE DataReferencia = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[DanfeCobranca].[DataVencimento]' 
GO

UPDATE DanfeCobranca SET DataVencimento = {d '1822-09-07'} WHERE DataVencimento = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[DebitosRecCli].[DataEmissao]' 
GO

UPDATE DebitosRecCli SET DataEmissao = {d '1822-09-07'} WHERE DataEmissao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[DEInfo].[Data]' 
GO

UPDATE DEInfo SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[DEInfo].[DataConhEmbarque]' 
GO

UPDATE DEInfo SET DataConhEmbarque = {d '1822-09-07'} WHERE DataConhEmbarque = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[DEInfo].[DataAverbacao]' 
GO

UPDATE DEInfo SET DataAverbacao = {d '1822-09-07'} WHERE DataAverbacao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[DERegistro].[DataRegistro]' 
GO

UPDATE DERegistro SET DataRegistro = {d '1822-09-07'} WHERE DataRegistro = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[DevolucaoCheque].[Data]' 
GO

UPDATE DevolucaoCheque SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[DevolucaoCheque].[DataVencimento]' 
GO

UPDATE DevolucaoCheque SET DataVencimento = {d '1822-09-07'} WHERE DataVencimento = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[DIInfo].[Data]' 
GO

UPDATE DIInfo SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[DIInfo].[DataDesembaraco]' 
GO

UPDATE DIInfo SET DataDesembaraco = {d '1822-09-07'} WHERE DataDesembaraco = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ECF].[DataLog]' 
GO

UPDATE ECF SET DataLog = {d '1822-09-07'} WHERE DataLog = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[EFDTabelas].[DataAtualizacao]' 
GO

UPDATE EFDTabelas SET DataAtualizacao = {d '1822-09-07'} WHERE DataAtualizacao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[EmailsEnviados].[Data]' 
GO

UPDATE EmailsEnviados SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Empenho].[Data]' 
GO

UPDATE Empenho SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Estados].[ICMSProt21Ini]' 
GO

UPDATE Estados SET ICMSProt21Ini = {d '1822-09-07'} WHERE ICMSProt21Ini = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Estados].[ICMSProt21Fim]' 
GO

UPDATE Estados SET ICMSProt21Fim = {d '1822-09-07'} WHERE ICMSProt21Fim = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Estados].[DataIniAliqInternaAtual]' 
GO

UPDATE Estados SET DataIniAliqInternaAtual = {d '1822-09-07'} WHERE DataIniAliqInternaAtual = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Estados].[DataIniAliqImportacaoAtual]' 
GO

UPDATE Estados SET DataIniAliqImportacaoAtual = {d '1822-09-07'} WHERE DataIniAliqImportacaoAtual = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Estados].[DataIniAliqFCPAtual]' 
GO

UPDATE Estados SET DataIniAliqFCPAtual = {d '1822-09-07'} WHERE DataIniAliqFCPAtual = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[EstoqueProduto].[DataInventario]' 
GO

UPDATE EstoqueProduto SET DataInventario = {d '1822-09-07'} WHERE DataInventario = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[EstoqueProduto].[DataInicial]' 
GO

UPDATE EstoqueProduto SET DataInicial = {d '1822-09-07'} WHERE DataInicial = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[EstoqueProduto].[DataUltAtualizacao]' 
GO

UPDATE EstoqueProduto SET DataUltAtualizacao = {d '1822-09-07'} WHERE DataUltAtualizacao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[EstoqueProdutoTerc].[DataInventario]' 
GO

UPDATE EstoqueProdutoTerc SET DataInventario = {d '1822-09-07'} WHERE DataInventario = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[EstoqueProdutoTerc].[DataInicial]' 
GO

UPDATE EstoqueProdutoTerc SET DataInicial = {d '1822-09-07'} WHERE DataInicial = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[EstoqueProdutoTerc].[DataUltAtualizacao]' 
GO

UPDATE EstoqueProdutoTerc SET DataUltAtualizacao = {d '1822-09-07'} WHERE DataUltAtualizacao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Exercicios].[DataInicio]' 
GO

UPDATE Exercicios SET DataInicio = {d '1822-09-07'} WHERE DataInicio = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Exercicios].[DataFim]' 
GO

UPDATE Exercicios SET DataFim = {d '1822-09-07'} WHERE DataFim = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ExerciciosFilial].[DataApuracao]' 
GO

UPDATE ExerciciosFilial SET DataApuracao = {d '1822-09-07'} WHERE DataApuracao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ExtratosBancarios].[DataGravacao]' 
GO

UPDATE ExtratosBancarios SET DataGravacao = {d '1822-09-07'} WHERE DataGravacao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ExtratosBancarios].[DataSaldoInicial]' 
GO

UPDATE ExtratosBancarios SET DataSaldoInicial = {d '1822-09-07'} WHERE DataSaldoInicial = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ExtratosBancarios].[DataSaldoFinal]' 
GO

UPDATE ExtratosBancarios SET DataSaldoFinal = {d '1822-09-07'} WHERE DataSaldoFinal = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ExtratosBancarios].[DataImport]' 
GO

UPDATE ExtratosBancarios SET DataImport = {d '1822-09-07'} WHERE DataImport = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[FaturamentoContratosRelErros].[DataReferencia]' 
GO

UPDATE FaturamentoContratosRelErros SET DataReferencia = {d '1822-09-07'} WHERE DataReferencia = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[FechamentoBoletos].[DataFechamento]' 
GO

UPDATE FechamentoBoletos SET DataFechamento = {d '1822-09-07'} WHERE DataFechamento = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Feriados].[Data]' 
GO

UPDATE Feriados SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[FiliaisClientes].[DataUltVisita]' 
GO

UPDATE FiliaisClientes SET DataUltVisita = {d '1822-09-07'} WHERE DataUltVisita = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[FiliaisClientes].[DataUltAtualizacao]' 
GO

UPDATE FiliaisClientes SET DataUltAtualizacao = {d '1822-09-07'} WHERE DataUltAtualizacao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[FiliaisClientes].[DataRegistro]' 
GO

UPDATE FiliaisClientes SET DataRegistro = {d '1822-09-07'} WHERE DataRegistro = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[FiliaisContatos].[DataUltVisita]' 
GO

UPDATE FiliaisContatos SET DataUltVisita = {d '1822-09-07'} WHERE DataUltVisita = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[FiliaisEmpresa].[DataJucerja]' 
GO

UPDATE FiliaisEmpresa SET DataJucerja = {d '1822-09-07'} WHERE DataJucerja = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[FiliaisFornecedores].[DataRegistro]' 
GO

UPDATE FiliaisFornecedores SET DataRegistro = {d '1822-09-07'} WHERE DataRegistro = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[FiliaisFornecedores].[DataAlteracao]' 
GO

UPDATE FiliaisFornecedores SET DataAlteracao = {d '1822-09-07'} WHERE DataAlteracao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[FilialClienteFilEmp].[DataPrimeiraCompra]' 
GO

UPDATE FilialClienteFilEmp SET DataPrimeiraCompra = {d '1822-09-07'} WHERE DataPrimeiraCompra = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[FilialClienteFilEmp].[DataUltimaCompra]' 
GO

UPDATE FilialClienteFilEmp SET DataUltimaCompra = {d '1822-09-07'} WHERE DataUltimaCompra = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[FilialClienteFilEmp].[DataUltimoProtesto]' 
GO

UPDATE FilialClienteFilEmp SET DataUltimoProtesto = {d '1822-09-07'} WHERE DataUltimoProtesto = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[FilialContatoData].[Data]' 
GO

UPDATE FilialContatoData SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[FilialFornFilEmp].[DataPrimeiraCompra]' 
GO

UPDATE FilialFornFilEmp SET DataPrimeiraCompra = {d '1822-09-07'} WHERE DataPrimeiraCompra = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[FilialFornFilEmp].[DataUltimaCompra]' 
GO

UPDATE FilialFornFilEmp SET DataUltimaCompra = {d '1822-09-07'} WHERE DataUltimaCompra = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[FilialFornFilEmp].[DataUltDevolucao]' 
GO

UPDATE FilialFornFilEmp SET DataUltDevolucao = {d '1822-09-07'} WHERE DataUltDevolucao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Fluxo].[DataInicial]' 
GO

UPDATE Fluxo SET DataInicial = {d '1822-09-07'} WHERE DataInicial = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Fluxo].[DataFinal]' 
GO

UPDATE Fluxo SET DataFinal = {d '1822-09-07'} WHERE DataFinal = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Fluxo].[DataDadosReais]' 
GO

UPDATE Fluxo SET DataDadosReais = {d '1822-09-07'} WHERE DataDadosReais = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[FluxoAnalitico].[Data]' 
GO

UPDATE FluxoAnalitico SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[FluxoAnalitico].[DataReferencia]' 
GO

UPDATE FluxoAnalitico SET DataReferencia = {d '1822-09-07'} WHERE DataReferencia = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[FluxoAplic].[DataResgatePrevista]' 
GO

UPDATE FluxoAplic SET DataResgatePrevista = {d '1822-09-07'} WHERE DataResgatePrevista = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[FluxoContratoItemNF].[Data]' 
GO

UPDATE FluxoContratoItemNF SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[FluxoContratoPag].[DataReferencia]' 
GO

UPDATE FluxoContratoPag SET DataReferencia = {d '1822-09-07'} WHERE DataReferencia = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[FluxoContratoPag].[Data]' 
GO

UPDATE FluxoContratoPag SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[FluxoContratoRec].[DataReferencia]' 
GO

UPDATE FluxoContratoRec SET DataReferencia = {d '1822-09-07'} WHERE DataReferencia = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[FluxoContratoRec].[Data]' 
GO

UPDATE FluxoContratoRec SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[FluxoForn].[Data]' 
GO

UPDATE FluxoForn SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[FluxoPC].[Data]' 
GO

UPDATE FluxoPC SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[FluxoPV].[Data]' 
GO

UPDATE FluxoPV SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[FluxoSintetico].[Data]' 
GO

UPDATE FluxoSintetico SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[FluxoTipoAplic].[Data]' 
GO

UPDATE FluxoTipoAplic SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[FluxoTipoForn].[Data]' 
GO

UPDATE FluxoTipoForn SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[FormacaoPrecoCalc].[Data]' 
GO

UPDATE FormacaoPrecoCalc SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[FormacaoPrecoCalcLin].[Data]' 
GO

UPDATE FormacaoPrecoCalcLin SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[FornecedorHistorico].[DataAtualizacao]' 
GO

UPDATE FornecedorHistorico SET DataAtualizacao = {d '1822-09-07'} WHERE DataAtualizacao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[FornecedorProduto].[DataPedido]' 
GO

UPDATE FornecedorProduto SET DataPedido = {d '1822-09-07'} WHERE DataPedido = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[FornecedorProduto].[DataReceb]' 
GO

UPDATE FornecedorProduto SET DataReceb = {d '1822-09-07'} WHERE DataReceb = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[FornecedorProdutoFF].[DataUltimaCompra]' 
GO

UPDATE FornecedorProdutoFF SET DataUltimaCompra = {d '1822-09-07'} WHERE DataUltimaCompra = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[FornecedorProdutoFF].[DataPedido]' 
GO

UPDATE FornecedorProdutoFF SET DataPedido = {d '1822-09-07'} WHERE DataPedido = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[FornecedorProdutoFF].[DataReceb]' 
GO

UPDATE FornecedorProdutoFF SET DataReceb = {d '1822-09-07'} WHERE DataReceb = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[FornecedorProdutoFF].[DataUltimaCotacao]' 
GO

UPDATE FornecedorProdutoFF SET DataUltimaCotacao = {d '1822-09-07'} WHERE DataUltimaCotacao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[FrequenciaDeTarefas].[DataUltExecucao]' 
GO

UPDATE FrequenciaDeTarefas SET DataUltExecucao = {d '1822-09-07'} WHERE DataUltExecucao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Garantia].[DataVenda]' 
GO

UPDATE Garantia SET DataVenda = {d '1822-09-07'} WHERE DataVenda = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[GNRICMS].[DataPagto]' 
GO

UPDATE GNRICMS SET DataPagto = {d '1822-09-07'} WHERE DataPagto = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[GNRICMS].[Vencimento]' 
GO

UPDATE GNRICMS SET Vencimento = {d '1822-09-07'} WHERE Vencimento = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[GNRICMS].[DataRef]' 
GO

UPDATE GNRICMS SET DataRef = {d '1822-09-07'} WHERE DataRef = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[GuiasICMS].[Data]' 
GO

UPDATE GuiasICMS SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[GuiasICMS].[DataEntrega]' 
GO

UPDATE GuiasICMS SET DataEntrega = {d '1822-09-07'} WHERE DataEntrega = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[GuiasICMS].[ApuracaoDe]' 
GO

UPDATE GuiasICMS SET ApuracaoDe = {d '1822-09-07'} WHERE ApuracaoDe = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[GuiasICMS].[ApuracaoAte]' 
GO

UPDATE GuiasICMS SET ApuracaoAte = {d '1822-09-07'} WHERE ApuracaoAte = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[GuiasICMS].[Vencimento]' 
GO

UPDATE GuiasICMS SET Vencimento = {d '1822-09-07'} WHERE Vencimento = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[GuiasICMSST].[Data]' 
GO

UPDATE GuiasICMSST SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[GuiasICMSST].[DataEntrega]' 
GO

UPDATE GuiasICMSST SET DataEntrega = {d '1822-09-07'} WHERE DataEntrega = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[GuiasICMSST].[ApuracaoDe]' 
GO

UPDATE GuiasICMSST SET ApuracaoDe = {d '1822-09-07'} WHERE ApuracaoDe = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[GuiasICMSST].[ApuracaoAte]' 
GO

UPDATE GuiasICMSST SET ApuracaoAte = {d '1822-09-07'} WHERE ApuracaoAte = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[GuiasICMSST].[Vencimento]' 
GO

UPDATE GuiasICMSST SET Vencimento = {d '1822-09-07'} WHERE Vencimento = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[IBPTax].[Validade]' 
GO

UPDATE IBPTax SET Validade = {d '1822-09-07'} WHERE Validade = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[IBPTaxUF].[VigenciaDe]' 
GO

UPDATE IBPTaxUF SET VigenciaDe = {d '1822-09-07'} WHERE VigenciaDe = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[IBPTaxUF].[VigenciaAte]' 
GO

UPDATE IBPTaxUF SET VigenciaAte = {d '1822-09-07'} WHERE VigenciaAte = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ICMSExcecoes].[ICMSSTBaseDuplaIni]' 
GO

UPDATE ICMSExcecoes SET ICMSSTBaseDuplaIni = {d '1822-09-07'} WHERE ICMSSTBaseDuplaIni = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportacaoCtb].[Data]' 
GO

UPDATE ImportacaoCtb SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportacaoInv].[Data]' 
GO

UPDATE ImportacaoInv SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportCli].[FilialDataUltVisita]' 
GO

UPDATE ImportCli SET FilialDataUltVisita = {d '1822-09-07'} WHERE FilialDataUltVisita = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportCtbArq].[DataArquivo]' 
GO

UPDATE ImportCtbArq SET DataArquivo = {d '1822-09-07'} WHERE DataArquivo = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportCtbArq].[DataImportacao]' 
GO

UPDATE ImportCtbArq SET DataImportacao = {d '1822-09-07'} WHERE DataImportacao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportCtbLctos].[Data]' 
GO

UPDATE ImportCtbLctos SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportDIXml].[DataDesembaraco]' 
GO

UPDATE ImportDIXml SET DataDesembaraco = {d '1822-09-07'} WHERE DataDesembaraco = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportDIXml].[DataRegistro]' 
GO

UPDATE ImportDIXml SET DataRegistro = {d '1822-09-07'} WHERE DataRegistro = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportDIXml].[DataChegada]' 
GO

UPDATE ImportDIXml SET DataChegada = {d '1822-09-07'} WHERE DataChegada = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportDIXml].[DataEmbarque]' 
GO

UPDATE ImportDIXml SET DataEmbarque = {d '1822-09-07'} WHERE DataEmbarque = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportEstoqueInicial].[DataInicial]' 
GO

UPDATE ImportEstoqueInicial SET DataInicial = {d '1822-09-07'} WHERE DataInicial = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportIBPTaxAC].[vigenciainicio]' 
GO

UPDATE ImportIBPTaxAC SET vigenciainicio = {d '1822-09-07'} WHERE vigenciainicio = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportIBPTaxAC].[vigenciafim]' 
GO

UPDATE ImportIBPTaxAC SET vigenciafim = {d '1822-09-07'} WHERE vigenciafim = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportIBPTaxAL].[vigenciainicio]' 
GO

UPDATE ImportIBPTaxAL SET vigenciainicio = {d '1822-09-07'} WHERE vigenciainicio = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportIBPTaxAL].[vigenciafim]' 
GO

UPDATE ImportIBPTaxAL SET vigenciafim = {d '1822-09-07'} WHERE vigenciafim = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportIBPTaxAM].[vigenciainicio]' 
GO

UPDATE ImportIBPTaxAM SET vigenciainicio = {d '1822-09-07'} WHERE vigenciainicio = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportIBPTaxAM].[vigenciafim]' 
GO

UPDATE ImportIBPTaxAM SET vigenciafim = {d '1822-09-07'} WHERE vigenciafim = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportIBPTaxAP].[vigenciainicio]' 
GO

UPDATE ImportIBPTaxAP SET vigenciainicio = {d '1822-09-07'} WHERE vigenciainicio = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportIBPTaxAP].[vigenciafim]' 
GO

UPDATE ImportIBPTaxAP SET vigenciafim = {d '1822-09-07'} WHERE vigenciafim = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportIBPTaxBA].[vigenciainicio]' 
GO

UPDATE ImportIBPTaxBA SET vigenciainicio = {d '1822-09-07'} WHERE vigenciainicio = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportIBPTaxBA].[vigenciafim]' 
GO

UPDATE ImportIBPTaxBA SET vigenciafim = {d '1822-09-07'} WHERE vigenciafim = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportIBPTaxCE].[vigenciainicio]' 
GO

UPDATE ImportIBPTaxCE SET vigenciainicio = {d '1822-09-07'} WHERE vigenciainicio = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportIBPTaxCE].[vigenciafim]' 
GO

UPDATE ImportIBPTaxCE SET vigenciafim = {d '1822-09-07'} WHERE vigenciafim = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportIBPTaxDF].[vigenciainicio]' 
GO

UPDATE ImportIBPTaxDF SET vigenciainicio = {d '1822-09-07'} WHERE vigenciainicio = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportIBPTaxDF].[vigenciafim]' 
GO

UPDATE ImportIBPTaxDF SET vigenciafim = {d '1822-09-07'} WHERE vigenciafim = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportIBPTaxES].[vigenciainicio]' 
GO

UPDATE ImportIBPTaxES SET vigenciainicio = {d '1822-09-07'} WHERE vigenciainicio = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportIBPTaxES].[vigenciafim]' 
GO

UPDATE ImportIBPTaxES SET vigenciafim = {d '1822-09-07'} WHERE vigenciafim = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportIBPTaxGO].[vigenciainicio]' 
GO

UPDATE ImportIBPTaxGO SET vigenciainicio = {d '1822-09-07'} WHERE vigenciainicio = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportIBPTaxGO].[vigenciafim]' 
GO

UPDATE ImportIBPTaxGO SET vigenciafim = {d '1822-09-07'} WHERE vigenciafim = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportIBPTaxMA].[vigenciainicio]' 
GO

UPDATE ImportIBPTaxMA SET vigenciainicio = {d '1822-09-07'} WHERE vigenciainicio = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportIBPTaxMA].[vigenciafim]' 
GO

UPDATE ImportIBPTaxMA SET vigenciafim = {d '1822-09-07'} WHERE vigenciafim = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportIBPTaxMG].[vigenciainicio]' 
GO

UPDATE ImportIBPTaxMG SET vigenciainicio = {d '1822-09-07'} WHERE vigenciainicio = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportIBPTaxMG].[vigenciafim]' 
GO

UPDATE ImportIBPTaxMG SET vigenciafim = {d '1822-09-07'} WHERE vigenciafim = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportIBPTaxMS].[vigenciainicio]' 
GO

UPDATE ImportIBPTaxMS SET vigenciainicio = {d '1822-09-07'} WHERE vigenciainicio = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportIBPTaxMS].[vigenciafim]' 
GO

UPDATE ImportIBPTaxMS SET vigenciafim = {d '1822-09-07'} WHERE vigenciafim = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportIBPTaxMT].[vigenciainicio]' 
GO

UPDATE ImportIBPTaxMT SET vigenciainicio = {d '1822-09-07'} WHERE vigenciainicio = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportIBPTaxMT].[vigenciafim]' 
GO

UPDATE ImportIBPTaxMT SET vigenciafim = {d '1822-09-07'} WHERE vigenciafim = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportIBPTaxPA].[vigenciainicio]' 
GO

UPDATE ImportIBPTaxPA SET vigenciainicio = {d '1822-09-07'} WHERE vigenciainicio = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportIBPTaxPA].[vigenciafim]' 
GO

UPDATE ImportIBPTaxPA SET vigenciafim = {d '1822-09-07'} WHERE vigenciafim = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportIBPTaxPB].[vigenciainicio]' 
GO

UPDATE ImportIBPTaxPB SET vigenciainicio = {d '1822-09-07'} WHERE vigenciainicio = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportIBPTaxPB].[vigenciafim]' 
GO

UPDATE ImportIBPTaxPB SET vigenciafim = {d '1822-09-07'} WHERE vigenciafim = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportIBPTaxPE].[vigenciainicio]' 
GO

UPDATE ImportIBPTaxPE SET vigenciainicio = {d '1822-09-07'} WHERE vigenciainicio = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportIBPTaxPE].[vigenciafim]' 
GO

UPDATE ImportIBPTaxPE SET vigenciafim = {d '1822-09-07'} WHERE vigenciafim = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportIBPTaxPI].[vigenciainicio]' 
GO

UPDATE ImportIBPTaxPI SET vigenciainicio = {d '1822-09-07'} WHERE vigenciainicio = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportIBPTaxPI].[vigenciafim]' 
GO

UPDATE ImportIBPTaxPI SET vigenciafim = {d '1822-09-07'} WHERE vigenciafim = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportIBPTaxPR].[vigenciainicio]' 
GO

UPDATE ImportIBPTaxPR SET vigenciainicio = {d '1822-09-07'} WHERE vigenciainicio = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportIBPTaxPR].[vigenciafim]' 
GO

UPDATE ImportIBPTaxPR SET vigenciafim = {d '1822-09-07'} WHERE vigenciafim = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportIBPTaxRJ].[vigenciainicio]' 
GO

UPDATE ImportIBPTaxRJ SET vigenciainicio = {d '1822-09-07'} WHERE vigenciainicio = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportIBPTaxRJ].[vigenciafim]' 
GO

UPDATE ImportIBPTaxRJ SET vigenciafim = {d '1822-09-07'} WHERE vigenciafim = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportIBPTaxRN].[vigenciainicio]' 
GO

UPDATE ImportIBPTaxRN SET vigenciainicio = {d '1822-09-07'} WHERE vigenciainicio = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportIBPTaxRN].[vigenciafim]' 
GO

UPDATE ImportIBPTaxRN SET vigenciafim = {d '1822-09-07'} WHERE vigenciafim = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportIBPTaxRO].[vigenciainicio]' 
GO

UPDATE ImportIBPTaxRO SET vigenciainicio = {d '1822-09-07'} WHERE vigenciainicio = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportIBPTaxRO].[vigenciafim]' 
GO

UPDATE ImportIBPTaxRO SET vigenciafim = {d '1822-09-07'} WHERE vigenciafim = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportIBPTaxRR].[vigenciainicio]' 
GO

UPDATE ImportIBPTaxRR SET vigenciainicio = {d '1822-09-07'} WHERE vigenciainicio = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportIBPTaxRR].[vigenciafim]' 
GO

UPDATE ImportIBPTaxRR SET vigenciafim = {d '1822-09-07'} WHERE vigenciafim = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportIBPTaxRS].[vigenciainicio]' 
GO

UPDATE ImportIBPTaxRS SET vigenciainicio = {d '1822-09-07'} WHERE vigenciainicio = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportIBPTaxRS].[vigenciafim]' 
GO

UPDATE ImportIBPTaxRS SET vigenciafim = {d '1822-09-07'} WHERE vigenciafim = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportIBPTaxSC].[vigenciainicio]' 
GO

UPDATE ImportIBPTaxSC SET vigenciainicio = {d '1822-09-07'} WHERE vigenciainicio = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportIBPTaxSC].[vigenciafim]' 
GO

UPDATE ImportIBPTaxSC SET vigenciafim = {d '1822-09-07'} WHERE vigenciafim = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportIBPTaxSE].[vigenciainicio]' 
GO

UPDATE ImportIBPTaxSE SET vigenciainicio = {d '1822-09-07'} WHERE vigenciainicio = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportIBPTaxSE].[vigenciafim]' 
GO

UPDATE ImportIBPTaxSE SET vigenciafim = {d '1822-09-07'} WHERE vigenciafim = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportIBPTaxSP].[vigenciainicio]' 
GO

UPDATE ImportIBPTaxSP SET vigenciainicio = {d '1822-09-07'} WHERE vigenciainicio = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportIBPTaxSP].[vigenciafim]' 
GO

UPDATE ImportIBPTaxSP SET vigenciafim = {d '1822-09-07'} WHERE vigenciafim = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportIBPTaxTO].[vigenciainicio]' 
GO

UPDATE ImportIBPTaxTO SET vigenciainicio = {d '1822-09-07'} WHERE vigenciainicio = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportIBPTaxTO].[vigenciafim]' 
GO

UPDATE ImportIBPTaxTO SET vigenciafim = {d '1822-09-07'} WHERE vigenciafim = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportKit].[Data]' 
GO

UPDATE ImportKit SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportMovCaixa].[Data]' 
GO

UPDATE ImportMovCaixa SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportNFeCobrXml].[DataVencimento]' 
GO

UPDATE ImportNFeCobrXml SET DataVencimento = {d '1822-09-07'} WHERE DataVencimento = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportNFeXml].[DataEmissao]' 
GO

UPDATE ImportNFeXml SET DataEmissao = {d '1822-09-07'} WHERE DataEmissao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportNFeXml].[DataES]' 
GO

UPDATE ImportNFeXml SET DataES = {d '1822-09-07'} WHERE DataES = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportNFeXml].[DataPrestServico]' 
GO

UPDATE ImportNFeXml SET DataPrestServico = {d '1822-09-07'} WHERE DataPrestServico = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportOVAcomp].[DataProxCobr]' 
GO

UPDATE ImportOVAcomp SET DataProxCobr = {d '1822-09-07'} WHERE DataProxCobr = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportOVAcomp].[DataPrevReceb]' 
GO

UPDATE ImportOVAcomp SET DataPrevReceb = {d '1822-09-07'} WHERE DataPrevReceb = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportOVAcomp01].[DataProxCobr]' 
GO

UPDATE ImportOVAcomp01 SET DataProxCobr = {d '1822-09-07'} WHERE DataProxCobr = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportOVAcomp01].[DataPrevReceb]' 
GO

UPDATE ImportOVAcomp01 SET DataPrevReceb = {d '1822-09-07'} WHERE DataPrevReceb = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportPrevVenda].[DataAtualizacao]' 
GO

UPDATE ImportPrevVenda SET DataAtualizacao = {d '1822-09-07'} WHERE DataAtualizacao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportRastroInicial].[DataEntrada]' 
GO

UPDATE ImportRastroInicial SET DataEntrada = {d '1822-09-07'} WHERE DataEntrada = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportTabelaPrecoItem].[DataVigencia]' 
GO

UPDATE ImportTabelaPrecoItem SET DataVigencia = {d '1822-09-07'} WHERE DataVigencia = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportVend].[DataUltVenda]' 
GO

UPDATE ImportVend SET DataUltVenda = {d '1822-09-07'} WHERE DataUltVenda = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ImportVend].[DataLog]' 
GO

UPDATE ImportVend SET DataLog = {d '1822-09-07'} WHERE DataLog = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[IN86Modelos].[DataInicio]' 
GO

UPDATE IN86Modelos SET DataInicio = {d '1822-09-07'} WHERE DataInicio = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[IN86Modelos].[DataFim]' 
GO

UPDATE IN86Modelos SET DataFim = {d '1822-09-07'} WHERE DataFim = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[InfoAdicDocItem].[DataLimiteFaturamento]' 
GO

UPDATE InfoAdicDocItem SET DataLimiteFaturamento = {d '1822-09-07'} WHERE DataLimiteFaturamento = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[InfoArqICMS].[DataInicial]' 
GO

UPDATE InfoArqICMS SET DataInicial = {d '1822-09-07'} WHERE DataInicial = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[InfoArqICMS].[DataFinal]' 
GO

UPDATE InfoArqICMS SET DataFinal = {d '1822-09-07'} WHERE DataFinal = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[IntegracaoCliente].[DataAtualizacao]' 
GO

UPDATE IntegracaoCliente SET DataAtualizacao = {d '1822-09-07'} WHERE DataAtualizacao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[IntegracaoLog].[Data]' 
GO

UPDATE IntegracaoLog SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[IntegracaoPV].[DataEmissao]' 
GO

UPDATE IntegracaoPV SET DataEmissao = {d '1822-09-07'} WHERE DataEmissao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[IntegracaoPV].[DataEntrega]' 
GO

UPDATE IntegracaoPV SET DataEntrega = {d '1822-09-07'} WHERE DataEntrega = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[IntegracaoPV].[DataAtualizacao]' 
GO

UPDATE IntegracaoPV SET DataAtualizacao = {d '1822-09-07'} WHERE DataAtualizacao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[IntegracaoPVParc].[Vencimento]' 
GO

UPDATE IntegracaoPVParc SET Vencimento = {d '1822-09-07'} WHERE Vencimento = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[IntegracaoRelCli].[Data]' 
GO

UPDATE IntegracaoRelCli SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[IntegracaoRelCli].[DataFim]' 
GO

UPDATE IntegracaoRelCli SET DataFim = {d '1822-09-07'} WHERE DataFim = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[IntegracaoSldProd].[DataAtualizacao]' 
GO

UPDATE IntegracaoSldProd SET DataAtualizacao = {d '1822-09-07'} WHERE DataAtualizacao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[InvCliForn].[Data]' 
GO

UPDATE InvCliForn SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[InvCliForn].[DataGravacao]' 
GO

UPDATE InvCliForn SET DataGravacao = {d '1822-09-07'} WHERE DataGravacao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Inventario].[Data]' 
GO

UPDATE Inventario SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[InventarioPendente].[Data]' 
GO

UPDATE InventarioPendente SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[InventarioTerc].[Data]' 
GO

UPDATE InventarioTerc SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[InventarioTercProd].[Data]' 
GO

UPDATE InventarioTercProd SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ItemOPOperacoesMaquinas].[Data]' 
GO

UPDATE ItemOPOperacoesMaquinas SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ItemOS].[DataInicio]' 
GO

UPDATE ItemOS SET DataInicio = {d '1822-09-07'} WHERE DataInicio = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ItemOS].[DataFim]' 
GO

UPDATE ItemOS SET DataFim = {d '1822-09-07'} WHERE DataFim = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ItensConcorrenciaN].[DataNecessidade]' 
GO

UPDATE ItensConcorrenciaN SET DataNecessidade = {d '1822-09-07'} WHERE DataNecessidade = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ItensCotacaoN].[DataReferencia]' 
GO

UPDATE ItensCotacaoN SET DataReferencia = {d '1822-09-07'} WHERE DataReferencia = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ItensDeContrato].[DataIniCobranca]' 
GO

UPDATE ItensDeContrato SET DataIniCobranca = {d '1822-09-07'} WHERE DataIniCobranca = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ItensDeContrato].[DataProxCobranca]' 
GO

UPDATE ItensDeContrato SET DataProxCobranca = {d '1822-09-07'} WHERE DataProxCobranca = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ItensDeContrato].[DataRefIni]' 
GO

UPDATE ItensDeContrato SET DataRefIni = {d '1822-09-07'} WHERE DataRefIni = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ItensDeContrato].[DataRefFim]' 
GO

UPDATE ItensDeContrato SET DataRefFim = {d '1822-09-07'} WHERE DataRefFim = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ItensDeContratoCobranca].[DataUltCobranca]' 
GO

UPDATE ItensDeContratoCobranca SET DataUltCobranca = {d '1822-09-07'} WHERE DataUltCobranca = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ItensDeContratoCobranca].[DataRefIni]' 
GO

UPDATE ItensDeContratoCobranca SET DataRefIni = {d '1822-09-07'} WHERE DataRefIni = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ItensDeContratoCobranca].[DataRefFim]' 
GO

UPDATE ItensDeContratoCobranca SET DataRefFim = {d '1822-09-07'} WHERE DataRefFim = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ItensDeMedicaoCobranca].[DataUltCobranca]' 
GO

UPDATE ItensDeMedicaoCobranca SET DataUltCobranca = {d '1822-09-07'} WHERE DataUltCobranca = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ItensDeMedicaoCobranca].[DataRefIni]' 
GO

UPDATE ItensDeMedicaoCobranca SET DataRefIni = {d '1822-09-07'} WHERE DataRefIni = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ItensDeMedicaoCobranca].[DataRefFim]' 
GO

UPDATE ItensDeMedicaoCobranca SET DataRefFim = {d '1822-09-07'} WHERE DataRefFim = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ItensDeMedicaoContrato].[DataRefIni]' 
GO

UPDATE ItensDeMedicaoContrato SET DataRefIni = {d '1822-09-07'} WHERE DataRefIni = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ItensDeMedicaoContrato].[DataRefFim]' 
GO

UPDATE ItensDeMedicaoContrato SET DataRefFim = {d '1822-09-07'} WHERE DataRefFim = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ItensDeMedicaoContrato].[DataCobranca]' 
GO

UPDATE ItensDeMedicaoContrato SET DataCobranca = {d '1822-09-07'} WHERE DataCobranca = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ItensNFEntrega].[DataEntrega]' 
GO

UPDATE ItensNFEntrega SET DataEntrega = {d '1822-09-07'} WHERE DataEntrega = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ItensNFiscal].[DataEntrega]' 
GO

UPDATE ItensNFiscal SET DataEntrega = {d '1822-09-07'} WHERE DataEntrega = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ItensOPGrade].[DataInicioProd]' 
GO

UPDATE ItensOPGrade SET DataInicioProd = {d '1822-09-07'} WHERE DataInicioProd = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ItensOPGrade].[DataFimProd]' 
GO

UPDATE ItensOPGrade SET DataFimProd = {d '1822-09-07'} WHERE DataFimProd = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ItensOrcamentoSRV].[DataEntrega]' 
GO

UPDATE ItensOrcamentoSRV SET DataEntrega = {d '1822-09-07'} WHERE DataEntrega = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ItensOrcamentoVenda].[DataEntrega]' 
GO

UPDATE ItensOrcamentoVenda SET DataEntrega = {d '1822-09-07'} WHERE DataEntrega = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ItensOrcamentoVendaHist].[DataEntrega]' 
GO

UPDATE ItensOrcamentoVendaHist SET DataEntrega = {d '1822-09-07'} WHERE DataEntrega = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ItensOrdemProducaoBaixadasN].[DataInicioProd]' 
GO

UPDATE ItensOrdemProducaoBaixadasN SET DataInicioProd = {d '1822-09-07'} WHERE DataInicioProd = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ItensOrdemProducaoBaixadasN].[DataFimProd]' 
GO

UPDATE ItensOrdemProducaoBaixadasN SET DataFimProd = {d '1822-09-07'} WHERE DataFimProd = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ItensOrdemProducaoBaixadasN].[DataRealFim]' 
GO

UPDATE ItensOrdemProducaoBaixadasN SET DataRealFim = {d '1822-09-07'} WHERE DataRealFim = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ItensOrdemProducaoN].[DataInicioProd]' 
GO

UPDATE ItensOrdemProducaoN SET DataInicioProd = {d '1822-09-07'} WHERE DataInicioProd = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ItensOrdemProducaoN].[DataFimProd]' 
GO

UPDATE ItensOrdemProducaoN SET DataFimProd = {d '1822-09-07'} WHERE DataFimProd = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ItensOrdemServicoProdCons].[Data]' 
GO

UPDATE ItensOrdemServicoProdCons SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ItensPCDI].[DataPC]' 
GO

UPDATE ItensPCDI SET DataPC = {d '1822-09-07'} WHERE DataPC = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ItensPCEntrega].[DataEntrega]' 
GO

UPDATE ItensPCEntrega SET DataEntrega = {d '1822-09-07'} WHERE DataEntrega = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ItensPedCompraN].[DataLimite]' 
GO

UPDATE ItensPedCompraN SET DataLimite = {d '1822-09-07'} WHERE DataLimite = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ItensPedCompraN].[DeliveryDate]' 
GO

UPDATE ItensPedCompraN SET DeliveryDate = {d '1822-09-07'} WHERE DeliveryDate = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ItensPedidoDeVenda].[DataEntrega]' 
GO

UPDATE ItensPedidoDeVenda SET DataEntrega = {d '1822-09-07'} WHERE DataEntrega = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ItensPedidoDeVendaBaixados].[DataEntrega]' 
GO

UPDATE ItensPedidoDeVendaBaixados SET DataEntrega = {d '1822-09-07'} WHERE DataEntrega = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ItensPedidoSRV].[DataEntrega]' 
GO

UPDATE ItensPedidoSRV SET DataEntrega = {d '1822-09-07'} WHERE DataEntrega = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ItensPVEntrega].[DataEntrega]' 
GO

UPDATE ItensPVEntrega SET DataEntrega = {d '1822-09-07'} WHERE DataEntrega = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ItensSolicSRV].[DataVenda]' 
GO

UPDATE ItensSolicSRV SET DataVenda = {d '1822-09-07'} WHERE DataVenda = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ItensSolicSRV].[DataBaixa]' 
GO

UPDATE ItensSolicSRV SET DataBaixa = {d '1822-09-07'} WHERE DataBaixa = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Kit].[Data]' 
GO

UPDATE Kit SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[KitVenda].[Data]' 
GO

UPDATE KitVenda SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Lancamentos].[Data]' 
GO

UPDATE Lancamentos SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Lancamentos].[DataEstoque]' 
GO

UPDATE Lancamentos SET DataEstoque = {d '1822-09-07'} WHERE DataEstoque = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Lancamentos].[DataRegistro]' 
GO

UPDATE Lancamentos SET DataRegistro = {d '1822-09-07'} WHERE DataRegistro = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[LanPendente].[Data]' 
GO

UPDATE LanPendente SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[LanPendente].[DataEstoque]' 
GO

UPDATE LanPendente SET DataEstoque = {d '1822-09-07'} WHERE DataEstoque = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[LanPendente].[DataRegistro]' 
GO

UPDATE LanPendente SET DataRegistro = {d '1822-09-07'} WHERE DataRegistro = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[LctosExtratoBancario].[Data]' 
GO

UPDATE LctosExtratoBancario SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[LivRegES].[DataEmissao]' 
GO

UPDATE LivRegES SET DataEmissao = {d '1822-09-07'} WHERE DataEmissao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[LivRegES].[Data]' 
GO

UPDATE LivRegES SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[LivRegESCadProd].[DataInicial]' 
GO

UPDATE LivRegESCadProd SET DataInicial = {d '1822-09-07'} WHERE DataInicial = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[LivRegESCadProd].[DataFinal]' 
GO

UPDATE LivRegESCadProd SET DataFinal = {d '1822-09-07'} WHERE DataFinal = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[LivrosFechados].[DataInicial]' 
GO

UPDATE LivrosFechados SET DataInicial = {d '1822-09-07'} WHERE DataInicial = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[LivrosFechados].[DataFinal]' 
GO

UPDATE LivrosFechados SET DataFinal = {d '1822-09-07'} WHERE DataFinal = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[LivrosFechados].[DataImpressao]' 
GO

UPDATE LivrosFechados SET DataImpressao = {d '1822-09-07'} WHERE DataImpressao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[LivrosFilial].[DataInicial]' 
GO

UPDATE LivrosFilial SET DataInicial = {d '1822-09-07'} WHERE DataInicial = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[LivrosFilial].[DataFinal]' 
GO

UPDATE LivrosFilial SET DataFinal = {d '1822-09-07'} WHERE DataFinal = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[LivrosFilial].[ImpressoEm]' 
GO

UPDATE LivrosFilial SET ImpressoEm = {d '1822-09-07'} WHERE ImpressoEm = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Log].[Data]' 
GO

UPDATE Log SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[LogInterno].[Data]' 
GO

UPDATE LogInterno SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[LogMovEstoque].[DataLog]' 
GO

UPDATE LogMovEstoque SET DataLog = {d '1822-09-07'} WHERE DataLog = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[LogMovEstoque].[Data]' 
GO

UPDATE LogMovEstoque SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[LogMovEstoque].[DataInicioProducao]' 
GO

UPDATE LogMovEstoque SET DataInicioProducao = {d '1822-09-07'} WHERE DataInicioProducao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[LogMovEstoque].[DataRegistro]' 
GO

UPDATE LogMovEstoque SET DataRegistro = {d '1822-09-07'} WHERE DataRegistro = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[LogWFW].[Data]' 
GO

UPDATE LogWFW SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[LojaArqFisAnalitico].[Data]' 
GO

UPDATE LojaArqFisAnalitico SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[LojaArqFisMestre].[Data]' 
GO

UPDATE LojaArqFisMestre SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[LojaReducao].[DataMovimento]' 
GO

UPDATE LojaReducao SET DataMovimento = {d '1822-09-07'} WHERE DataMovimento = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[LojaReducao].[DataReducao]' 
GO

UPDATE LojaReducao SET DataReducao = {d '1822-09-07'} WHERE DataReducao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[LojaReducaoAliquota].[Data]' 
GO

UPDATE LojaReducaoAliquota SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Lote].[DataRegistro]' 
GO

UPDATE Lote SET DataRegistro = {d '1822-09-07'} WHERE DataRegistro = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[LotesCPRPend].[DataRecebto]' 
GO

UPDATE LotesCPRPend SET DataRecebto = {d '1822-09-07'} WHERE DataRecebto = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[MapaCotacao].[Data]' 
GO

UPDATE MapaCotacao SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[MapaDeEntrega].[Data]' 
GO

UPDATE MapaDeEntrega SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[MedicaoContrato].[Data]' 
GO

UPDATE MedicaoContrato SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[MedicaoContrato].[DataDeReferencia]' 
GO

UPDATE MedicaoContrato SET DataDeReferencia = {d '1822-09-07'} WHERE DataDeReferencia = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[MovimentoEstoque].[Data]' 
GO

UPDATE MovimentoEstoque SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[MovimentoEstoque].[DataInicioProducao]' 
GO

UPDATE MovimentoEstoque SET DataInicioProducao = {d '1822-09-07'} WHERE DataInicioProducao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[MovimentoEstoque].[DataRegistro]' 
GO

UPDATE MovimentoEstoque SET DataRegistro = {d '1822-09-07'} WHERE DataRegistro = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[MovimentoEstoqueGrade].[Data]' 
GO

UPDATE MovimentoEstoqueGrade SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[MovimentoEstoqueGrade].[DataInicioProducao]' 
GO

UPDATE MovimentoEstoqueGrade SET DataInicioProducao = {d '1822-09-07'} WHERE DataInicioProducao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[MovimentoEstoqueGrade].[DataRegistro]' 
GO

UPDATE MovimentoEstoqueGrade SET DataRegistro = {d '1822-09-07'} WHERE DataRegistro = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[MovimentosCaixa].[DataMovimento]' 
GO

UPDATE MovimentosCaixa SET DataMovimento = {d '1822-09-07'} WHERE DataMovimento = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[MovimentosCaixa].[DataPreDatado]' 
GO

UPDATE MovimentosCaixa SET DataPreDatado = {d '1822-09-07'} WHERE DataPreDatado = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[MovimentosContaCorrente].[DataMovimento]' 
GO

UPDATE MovimentosContaCorrente SET DataMovimento = {d '1822-09-07'} WHERE DataMovimento = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[MvDiaCcl].[Data]' 
GO

UPDATE MvDiaCcl SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[MvDiaCli].[Data]' 
GO

UPDATE MvDiaCli SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[MvDiaCta].[Data]' 
GO

UPDATE MvDiaCta SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[MvDiaForn].[Data]' 
GO

UPDATE MvDiaForn SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[NaturezaOPHistorico].[DataAtualizacao]' 
GO

UPDATE NaturezaOPHistorico SET DataAtualizacao = {d '1822-09-07'} WHERE DataAtualizacao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[NCM].[DataValidadeDe]' 
GO

UPDATE NCM SET DataValidadeDe = {d '1822-09-07'} WHERE DataValidadeDe = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[NCM].[DataValidadeAte]' 
GO

UPDATE NCM SET DataValidadeAte = {d '1822-09-07'} WHERE DataValidadeAte = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[NFe].[Data]' 
GO

UPDATE NFe SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[NFe].[DataEmissaoRPS]' 
GO

UPDATE NFe SET DataEmissaoRPS = {d '1822-09-07'} WHERE DataEmissaoRPS = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[NFe].[DataCancelamento]' 
GO

UPDATE NFe SET DataCancelamento = {d '1822-09-07'} WHERE DataCancelamento = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[NFe].[DataQuitacaoGuia]' 
GO

UPDATE NFe SET DataQuitacaoGuia = {d '1822-09-07'} WHERE DataQuitacaoGuia = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[NFeCab].[DataImportacao]' 
GO

UPDATE NFeCab SET DataImportacao = {d '1822-09-07'} WHERE DataImportacao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[NFeCab].[DataInicio]' 
GO

UPDATE NFeCab SET DataInicio = {d '1822-09-07'} WHERE DataInicio = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[NFeCab].[DataFim]' 
GO

UPDATE NFeCab SET DataFim = {d '1822-09-07'} WHERE DataFim = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[NFeFedLote].[Data]' 
GO

UPDATE NFeFedLote SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[NFeFedLoteLog].[Data]' 
GO

UPDATE NFeFedLoteLog SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[NFeFedProtNFe].[data]' 
GO

UPDATE NFeFedProtNFe SET data = {d '1822-09-07'} WHERE data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[NFeFedProtNFe].[DataRegistro]' 
GO

UPDATE NFeFedProtNFe SET DataRegistro = {d '1822-09-07'} WHERE DataRegistro = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[NFeFedRetCancNFe].[data]' 
GO

UPDATE NFeFedRetCancNFe SET data = {d '1822-09-07'} WHERE data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[NFeFedRetConsReci].[Data]' 
GO

UPDATE NFeFedRetConsReci SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[NFeFedRetEnvCCe].[dataRegEvento]' 
GO

UPDATE NFeFedRetEnvCCe SET dataRegEvento = {d '1822-09-07'} WHERE dataRegEvento = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[NFeFedRetEnvEventoCanc].[dataRegEvento]' 
GO

UPDATE NFeFedRetEnvEventoCanc SET dataRegEvento = {d '1822-09-07'} WHERE dataRegEvento = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[NFeFedRetEnvi].[Data]' 
GO

UPDATE NFeFedRetEnvi SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[NFeFedRetInutNFe].[data]' 
GO

UPDATE NFeFedRetInutNFe SET data = {d '1822-09-07'} WHERE data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[NFeFedScan].[DataEntrada]' 
GO

UPDATE NFeFedScan SET DataEntrada = {d '1822-09-07'} WHERE DataEntrada = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[NFeFedScan].[DataSaida]' 
GO

UPDATE NFeFedScan SET DataSaida = {d '1822-09-07'} WHERE DataSaida = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[NFeRecebidas].[Data]' 
GO

UPDATE NFeRecebidas SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[NFeRecebidas].[DataEmissaoRPS]' 
GO

UPDATE NFeRecebidas SET DataEmissaoRPS = {d '1822-09-07'} WHERE DataEmissaoRPS = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[NFeRecebidas].[DataCancelamento]' 
GO

UPDATE NFeRecebidas SET DataCancelamento = {d '1822-09-07'} WHERE DataCancelamento = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[NFeRecebidas].[DataQuitacaoGuia]' 
GO

UPDATE NFeRecebidas SET DataQuitacaoGuia = {d '1822-09-07'} WHERE DataQuitacaoGuia = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[NFiscal].[DataEmissao]' 
GO

UPDATE NFiscal SET DataEmissao = {d '1822-09-07'} WHERE DataEmissao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[NFiscal].[DataEntrada]' 
GO

UPDATE NFiscal SET DataEntrada = {d '1822-09-07'} WHERE DataEntrada = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[NFiscal].[DataSaida]' 
GO

UPDATE NFiscal SET DataSaida = {d '1822-09-07'} WHERE DataSaida = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[NFiscal].[DataVencimento]' 
GO

UPDATE NFiscal SET DataVencimento = {d '1822-09-07'} WHERE DataVencimento = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[NFiscal].[DataReferencia]' 
GO

UPDATE NFiscal SET DataReferencia = {d '1822-09-07'} WHERE DataReferencia = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[NFiscal].[DataAtualizacao]' 
GO

UPDATE NFiscal SET DataAtualizacao = {d '1822-09-07'} WHERE DataAtualizacao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[NFiscal].[DataCancel]' 
GO

UPDATE NFiscal SET DataCancel = {d '1822-09-07'} WHERE DataCancel = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[NFiscal].[DataRegCancel]' 
GO

UPDATE NFiscal SET DataRegCancel = {d '1822-09-07'} WHERE DataRegCancel = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[NFiscal].[DataCadastro]' 
GO

UPDATE NFiscal SET DataCadastro = {d '1822-09-07'} WHERE DataCadastro = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[NFsPag].[DataEmissao]' 
GO

UPDATE NFsPag SET DataEmissao = {d '1822-09-07'} WHERE DataEmissao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[NFsPag].[DataVencimento]' 
GO

UPDATE NFsPag SET DataVencimento = {d '1822-09-07'} WHERE DataVencimento = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[NFsPagBaixadas].[DataEmissao]' 
GO

UPDATE NFsPagBaixadas SET DataEmissao = {d '1822-09-07'} WHERE DataEmissao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[NFsPagBaixadas].[DataVencimento]' 
GO

UPDATE NFsPagBaixadas SET DataVencimento = {d '1822-09-07'} WHERE DataVencimento = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[NFsPagDtCtb].[DataContabil]' 
GO

UPDATE NFsPagDtCtb SET DataContabil = {d '1822-09-07'} WHERE DataContabil = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[OcorrenciasRemParcRec].[DataRegistro]' 
GO

UPDATE OcorrenciasRemParcRec SET DataRegistro = {d '1822-09-07'} WHERE DataRegistro = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[OcorrenciasRemParcRec].[Data]' 
GO

UPDATE OcorrenciasRemParcRec SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[OcorrenciasRemParcRec].[NovaDataVcto]' 
GO

UPDATE OcorrenciasRemParcRec SET NovaDataVcto = {d '1822-09-07'} WHERE NovaDataVcto = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[OcorrenciasRetParcRec].[DataComplementar]' 
GO

UPDATE OcorrenciasRetParcRec SET DataComplementar = {d '1822-09-07'} WHERE DataComplementar = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[OLAP_Datas].[Data]' 
GO

UPDATE OLAP_Datas SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Operador].[DataLog]' 
GO

UPDATE Operador SET DataLog = {d '1822-09-07'} WHERE DataLog = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[OrcamentoSRV].[DataEmissao]' 
GO

UPDATE OrcamentoSRV SET DataEmissao = {d '1822-09-07'} WHERE DataEmissao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[OrcamentoSRV].[DataReferencia]' 
GO

UPDATE OrcamentoSRV SET DataReferencia = {d '1822-09-07'} WHERE DataReferencia = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[OrcamentoVenda].[DataEmissao]' 
GO

UPDATE OrcamentoVenda SET DataEmissao = {d '1822-09-07'} WHERE DataEmissao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[OrcamentoVenda].[DataReferencia]' 
GO

UPDATE OrcamentoVenda SET DataReferencia = {d '1822-09-07'} WHERE DataReferencia = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[OrcamentoVenda].[DataUltAlt]' 
GO

UPDATE OrcamentoVenda SET DataUltAlt = {d '1822-09-07'} WHERE DataUltAlt = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[OrcamentoVenda].[DataPrevReceb]' 
GO

UPDATE OrcamentoVenda SET DataPrevReceb = {d '1822-09-07'} WHERE DataPrevReceb = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[OrcamentoVenda].[DataProxCobr]' 
GO

UPDATE OrcamentoVenda SET DataProxCobr = {d '1822-09-07'} WHERE DataProxCobr = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[OrcamentoVenda].[DataEntrega]' 
GO

UPDATE OrcamentoVenda SET DataEntrega = {d '1822-09-07'} WHERE DataEntrega = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[OrcamentoVenda].[DataEnvio]' 
GO

UPDATE OrcamentoVenda SET DataEnvio = {d '1822-09-07'} WHERE DataEnvio = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[OrcamentoVenda].[DataPerda]' 
GO

UPDATE OrcamentoVenda SET DataPerda = {d '1822-09-07'} WHERE DataPerda = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[OrcamentoVendaHist].[DataEmissao]' 
GO

UPDATE OrcamentoVendaHist SET DataEmissao = {d '1822-09-07'} WHERE DataEmissao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[OrcamentoVendaHist].[DataReferencia]' 
GO

UPDATE OrcamentoVendaHist SET DataReferencia = {d '1822-09-07'} WHERE DataReferencia = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[OrcamentoVendaHist].[DataUltAlt]' 
GO

UPDATE OrcamentoVendaHist SET DataUltAlt = {d '1822-09-07'} WHERE DataUltAlt = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[OrcamentoVendaHist].[DataEntrega]' 
GO

UPDATE OrcamentoVendaHist SET DataEntrega = {d '1822-09-07'} WHERE DataEntrega = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[OrcamentoVendaHist].[DataEnvio]' 
GO

UPDATE OrcamentoVendaHist SET DataEnvio = {d '1822-09-07'} WHERE DataEnvio = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[OrcamentoVendaHist].[DataPerda]' 
GO

UPDATE OrcamentoVendaHist SET DataPerda = {d '1822-09-07'} WHERE DataPerda = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[OrcamentoVendaHist].[DataPrevReceb]' 
GO

UPDATE OrcamentoVendaHist SET DataPrevReceb = {d '1822-09-07'} WHERE DataPrevReceb = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[OrcamentoVendaHist].[DataProxCobr]' 
GO

UPDATE OrcamentoVendaHist SET DataProxCobr = {d '1822-09-07'} WHERE DataProxCobr = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[OrdemServicoProd].[DataEmissao]' 
GO

UPDATE OrdemServicoProd SET DataEmissao = {d '1822-09-07'} WHERE DataEmissao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[OrdensDeProducao].[DataEmissao]' 
GO

UPDATE OrdensDeProducao SET DataEmissao = {d '1822-09-07'} WHERE DataEmissao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[OrdensDeProducaoBaixadas].[DataEmissao]' 
GO

UPDATE OrdensDeProducaoBaixadas SET DataEmissao = {d '1822-09-07'} WHERE DataEmissao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[OS].[DataEmissao]' 
GO

UPDATE OS SET DataEmissao = {d '1822-09-07'} WHERE DataEmissao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PagtosPeriodicos].[Inicio]' 
GO

UPDATE PagtosPeriodicos SET Inicio = {d '1822-09-07'} WHERE Inicio = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PagtosPeriodicos].[Termino]' 
GO

UPDATE PagtosPeriodicos SET Termino = {d '1822-09-07'} WHERE Termino = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PagtosPeriodicos].[Proximo]' 
GO

UPDATE PagtosPeriodicos SET Proximo = {d '1822-09-07'} WHERE Proximo = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PalmLog].[Data]' 
GO

UPDATE PalmLog SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ParcelasOrcSRV].[DataVencimento]' 
GO

UPDATE ParcelasOrcSRV SET DataVencimento = {d '1822-09-07'} WHERE DataVencimento = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ParcelasOrcSRV].[Desconto1Ate]' 
GO

UPDATE ParcelasOrcSRV SET Desconto1Ate = {d '1822-09-07'} WHERE Desconto1Ate = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ParcelasOrcSRV].[Desconto2Ate]' 
GO

UPDATE ParcelasOrcSRV SET Desconto2Ate = {d '1822-09-07'} WHERE Desconto2Ate = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ParcelasOrcSRV].[Desconto3Ate]' 
GO

UPDATE ParcelasOrcSRV SET Desconto3Ate = {d '1822-09-07'} WHERE Desconto3Ate = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ParcelasOV].[DataVencimento]' 
GO

UPDATE ParcelasOV SET DataVencimento = {d '1822-09-07'} WHERE DataVencimento = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ParcelasOV].[Desconto1Ate]' 
GO

UPDATE ParcelasOV SET Desconto1Ate = {d '1822-09-07'} WHERE Desconto1Ate = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ParcelasOV].[Desconto2Ate]' 
GO

UPDATE ParcelasOV SET Desconto2Ate = {d '1822-09-07'} WHERE Desconto2Ate = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ParcelasOV].[Desconto3Ate]' 
GO

UPDATE ParcelasOV SET Desconto3Ate = {d '1822-09-07'} WHERE Desconto3Ate = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ParcelasOVHist].[DataVencimento]' 
GO

UPDATE ParcelasOVHist SET DataVencimento = {d '1822-09-07'} WHERE DataVencimento = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ParcelasOVHist].[Desconto1Ate]' 
GO

UPDATE ParcelasOVHist SET Desconto1Ate = {d '1822-09-07'} WHERE Desconto1Ate = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ParcelasOVHist].[Desconto2Ate]' 
GO

UPDATE ParcelasOVHist SET Desconto2Ate = {d '1822-09-07'} WHERE Desconto2Ate = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ParcelasOVHist].[Desconto3Ate]' 
GO

UPDATE ParcelasOVHist SET Desconto3Ate = {d '1822-09-07'} WHERE Desconto3Ate = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ParcelasPag].[DataVencimento]' 
GO

UPDATE ParcelasPag SET DataVencimento = {d '1822-09-07'} WHERE DataVencimento = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ParcelasPag].[DataVencimentoReal]' 
GO

UPDATE ParcelasPag SET DataVencimentoReal = {d '1822-09-07'} WHERE DataVencimentoReal = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ParcelasPag].[DataLib]' 
GO

UPDATE ParcelasPag SET DataLib = {d '1822-09-07'} WHERE DataLib = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ParcelasPagBaixadas].[DataVencimento]' 
GO

UPDATE ParcelasPagBaixadas SET DataVencimento = {d '1822-09-07'} WHERE DataVencimento = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ParcelasPagBaixadas].[DataVencimentoReal]' 
GO

UPDATE ParcelasPagBaixadas SET DataVencimentoReal = {d '1822-09-07'} WHERE DataVencimentoReal = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ParcelasPagBaixadas].[DataLib]' 
GO

UPDATE ParcelasPagBaixadas SET DataLib = {d '1822-09-07'} WHERE DataLib = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ParcelasPedidoDeVenda].[DataVencimento]' 
GO

UPDATE ParcelasPedidoDeVenda SET DataVencimento = {d '1822-09-07'} WHERE DataVencimento = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ParcelasPedidoDeVenda].[Desconto1Ate]' 
GO

UPDATE ParcelasPedidoDeVenda SET Desconto1Ate = {d '1822-09-07'} WHERE Desconto1Ate = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ParcelasPedidoDeVenda].[Desconto2Ate]' 
GO

UPDATE ParcelasPedidoDeVenda SET Desconto2Ate = {d '1822-09-07'} WHERE Desconto2Ate = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ParcelasPedidoDeVenda].[Desconto3Ate]' 
GO

UPDATE ParcelasPedidoDeVenda SET Desconto3Ate = {d '1822-09-07'} WHERE Desconto3Ate = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ParcelasPedidoDeVenda].[DataCredito]' 
GO

UPDATE ParcelasPedidoDeVenda SET DataCredito = {d '1822-09-07'} WHERE DataCredito = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ParcelasPedidoDeVenda].[DataEmissaoCheque]' 
GO

UPDATE ParcelasPedidoDeVenda SET DataEmissaoCheque = {d '1822-09-07'} WHERE DataEmissaoCheque = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ParcelasPedidoDeVenda].[DataDepositoCheque]' 
GO

UPDATE ParcelasPedidoDeVenda SET DataDepositoCheque = {d '1822-09-07'} WHERE DataDepositoCheque = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ParcelasPedidoDeVenda].[ValidadeCartao]' 
GO

UPDATE ParcelasPedidoDeVenda SET ValidadeCartao = {d '1822-09-07'} WHERE ValidadeCartao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ParcelasPedidoDeVenda].[DataTransacaoCartao]' 
GO

UPDATE ParcelasPedidoDeVenda SET DataTransacaoCartao = {d '1822-09-07'} WHERE DataTransacaoCartao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ParcelasPedidoDeVendaBaixado].[DataVencimento]' 
GO

UPDATE ParcelasPedidoDeVendaBaixado SET DataVencimento = {d '1822-09-07'} WHERE DataVencimento = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ParcelasPedidoDeVendaBaixado].[Desconto1Ate]' 
GO

UPDATE ParcelasPedidoDeVendaBaixado SET Desconto1Ate = {d '1822-09-07'} WHERE Desconto1Ate = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ParcelasPedidoDeVendaBaixado].[Desconto2Ate]' 
GO

UPDATE ParcelasPedidoDeVendaBaixado SET Desconto2Ate = {d '1822-09-07'} WHERE Desconto2Ate = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ParcelasPedidoDeVendaBaixado].[Desconto3Ate]' 
GO

UPDATE ParcelasPedidoDeVendaBaixado SET Desconto3Ate = {d '1822-09-07'} WHERE Desconto3Ate = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ParcelasPedidoDeVendaBaixado].[DataCredito]' 
GO

UPDATE ParcelasPedidoDeVendaBaixado SET DataCredito = {d '1822-09-07'} WHERE DataCredito = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ParcelasPedidoDeVendaBaixado].[DataEmissaoCheque]' 
GO

UPDATE ParcelasPedidoDeVendaBaixado SET DataEmissaoCheque = {d '1822-09-07'} WHERE DataEmissaoCheque = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ParcelasPedidoDeVendaBaixado].[DataDepositoCheque]' 
GO

UPDATE ParcelasPedidoDeVendaBaixado SET DataDepositoCheque = {d '1822-09-07'} WHERE DataDepositoCheque = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ParcelasPedidoDeVendaBaixado].[ValidadeCartao]' 
GO

UPDATE ParcelasPedidoDeVendaBaixado SET ValidadeCartao = {d '1822-09-07'} WHERE ValidadeCartao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ParcelasPedidoDeVendaBaixado].[DataTransacaoCartao]' 
GO

UPDATE ParcelasPedidoDeVendaBaixado SET DataTransacaoCartao = {d '1822-09-07'} WHERE DataTransacaoCartao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ParcelasPedidoSRV].[DataVencimento]' 
GO

UPDATE ParcelasPedidoSRV SET DataVencimento = {d '1822-09-07'} WHERE DataVencimento = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ParcelasPedidoSRV].[Desconto1Ate]' 
GO

UPDATE ParcelasPedidoSRV SET Desconto1Ate = {d '1822-09-07'} WHERE Desconto1Ate = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ParcelasPedidoSRV].[Desconto2Ate]' 
GO

UPDATE ParcelasPedidoSRV SET Desconto2Ate = {d '1822-09-07'} WHERE Desconto2Ate = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ParcelasPedidoSRV].[Desconto3Ate]' 
GO

UPDATE ParcelasPedidoSRV SET Desconto3Ate = {d '1822-09-07'} WHERE Desconto3Ate = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ParcelasPedidoSRV].[DataEmissao]' 
GO

UPDATE ParcelasPedidoSRV SET DataEmissao = {d '1822-09-07'} WHERE DataEmissao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ParcelasPedidoSRV].[DataDepositoCheque]' 
GO

UPDATE ParcelasPedidoSRV SET DataDepositoCheque = {d '1822-09-07'} WHERE DataDepositoCheque = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ParcelasPedidoSRV].[ValidadeCartao]' 
GO

UPDATE ParcelasPedidoSRV SET ValidadeCartao = {d '1822-09-07'} WHERE ValidadeCartao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ParcelasPedidoSRV].[DataCredito]' 
GO

UPDATE ParcelasPedidoSRV SET DataCredito = {d '1822-09-07'} WHERE DataCredito = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ParcelasPedidoSRV].[DataEmissaoCheque]' 
GO

UPDATE ParcelasPedidoSRV SET DataEmissaoCheque = {d '1822-09-07'} WHERE DataEmissaoCheque = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ParcelasPedidoSRV].[DataTransacaoCartao]' 
GO

UPDATE ParcelasPedidoSRV SET DataTransacaoCartao = {d '1822-09-07'} WHERE DataTransacaoCartao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ParcelasRec].[DataVencimento]' 
GO

UPDATE ParcelasRec SET DataVencimento = {d '1822-09-07'} WHERE DataVencimento = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ParcelasRec].[DataVencimentoReal]' 
GO

UPDATE ParcelasRec SET DataVencimentoReal = {d '1822-09-07'} WHERE DataVencimentoReal = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ParcelasRec].[Desconto1Ate]' 
GO

UPDATE ParcelasRec SET Desconto1Ate = {d '1822-09-07'} WHERE Desconto1Ate = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ParcelasRec].[Desconto2Ate]' 
GO

UPDATE ParcelasRec SET Desconto2Ate = {d '1822-09-07'} WHERE Desconto2Ate = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ParcelasRec].[Desconto3Ate]' 
GO

UPDATE ParcelasRec SET Desconto3Ate = {d '1822-09-07'} WHERE Desconto3Ate = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ParcelasRec].[DataImpressaoBoleto]' 
GO

UPDATE ParcelasRec SET DataImpressaoBoleto = {d '1822-09-07'} WHERE DataImpressaoBoleto = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ParcelasRec].[DataPrevisao]' 
GO

UPDATE ParcelasRec SET DataPrevisao = {d '1822-09-07'} WHERE DataPrevisao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ParcelasRec].[DataCredito]' 
GO

UPDATE ParcelasRec SET DataCredito = {d '1822-09-07'} WHERE DataCredito = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ParcelasRec].[DataEmissaoCheque]' 
GO

UPDATE ParcelasRec SET DataEmissaoCheque = {d '1822-09-07'} WHERE DataEmissaoCheque = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ParcelasRec].[DataDepositoCheque]' 
GO

UPDATE ParcelasRec SET DataDepositoCheque = {d '1822-09-07'} WHERE DataDepositoCheque = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ParcelasRec].[ValidadeCartao]' 
GO

UPDATE ParcelasRec SET ValidadeCartao = {d '1822-09-07'} WHERE ValidadeCartao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ParcelasRec].[DataTransacaoCartao]' 
GO

UPDATE ParcelasRec SET DataTransacaoCartao = {d '1822-09-07'} WHERE DataTransacaoCartao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ParcelasRecBaixadas].[DataVencimento]' 
GO

UPDATE ParcelasRecBaixadas SET DataVencimento = {d '1822-09-07'} WHERE DataVencimento = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ParcelasRecBaixadas].[DataVencimentoReal]' 
GO

UPDATE ParcelasRecBaixadas SET DataVencimentoReal = {d '1822-09-07'} WHERE DataVencimentoReal = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ParcelasRecBaixadas].[Desconto1Ate]' 
GO

UPDATE ParcelasRecBaixadas SET Desconto1Ate = {d '1822-09-07'} WHERE Desconto1Ate = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ParcelasRecBaixadas].[Desconto2Ate]' 
GO

UPDATE ParcelasRecBaixadas SET Desconto2Ate = {d '1822-09-07'} WHERE Desconto2Ate = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ParcelasRecBaixadas].[Desconto3Ate]' 
GO

UPDATE ParcelasRecBaixadas SET Desconto3Ate = {d '1822-09-07'} WHERE Desconto3Ate = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ParcelasRecBaixadas].[DataCredito]' 
GO

UPDATE ParcelasRecBaixadas SET DataCredito = {d '1822-09-07'} WHERE DataCredito = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ParcelasRecBaixadas].[DataEmissaoCheque]' 
GO

UPDATE ParcelasRecBaixadas SET DataEmissaoCheque = {d '1822-09-07'} WHERE DataEmissaoCheque = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ParcelasRecBaixadas].[DataDepositoCheque]' 
GO

UPDATE ParcelasRecBaixadas SET DataDepositoCheque = {d '1822-09-07'} WHERE DataDepositoCheque = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ParcelasRecBaixadas].[ValidadeCartao]' 
GO

UPDATE ParcelasRecBaixadas SET ValidadeCartao = {d '1822-09-07'} WHERE ValidadeCartao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ParcelasRecBaixadas].[DataTransacaoCartao]' 
GO

UPDATE ParcelasRecBaixadas SET DataTransacaoCartao = {d '1822-09-07'} WHERE DataTransacaoCartao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ParcelasRecDif].[DataRegistro]' 
GO

UPDATE ParcelasRecDif SET DataRegistro = {d '1822-09-07'} WHERE DataRegistro = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PedidoCompraN].[Data]' 
GO

UPDATE PedidoCompraN SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PedidoCompraN].[DataEmissao]' 
GO

UPDATE PedidoCompraN SET DataEmissao = {d '1822-09-07'} WHERE DataEmissao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PedidoCompraN].[DataEnvio]' 
GO

UPDATE PedidoCompraN SET DataEnvio = {d '1822-09-07'} WHERE DataEnvio = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PedidoCompraN].[DataAlteracao]' 
GO

UPDATE PedidoCompraN SET DataAlteracao = {d '1822-09-07'} WHERE DataAlteracao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PedidoCompraN].[DataBaixa]' 
GO

UPDATE PedidoCompraN SET DataBaixa = {d '1822-09-07'} WHERE DataBaixa = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PedidoCompraN].[DataFixa]' 
GO

UPDATE PedidoCompraN SET DataFixa = {d '1822-09-07'} WHERE DataFixa = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PedidoCompraN].[DataRefFluxo]' 
GO

UPDATE PedidoCompraN SET DataRefFluxo = {d '1822-09-07'} WHERE DataRefFluxo = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PedidoCompraN].[DataReg]' 
GO

UPDATE PedidoCompraN SET DataReg = {d '1822-09-07'} WHERE DataReg = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PedidoCompraN].[DataRegEnvio]' 
GO

UPDATE PedidoCompraN SET DataRegEnvio = {d '1822-09-07'} WHERE DataRegEnvio = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PedidoCompraN].[DataRegAprov]' 
GO

UPDATE PedidoCompraN SET DataRegAprov = {d '1822-09-07'} WHERE DataRegAprov = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PedidoCotacaoN].[DataEmissao]' 
GO

UPDATE PedidoCotacaoN SET DataEmissao = {d '1822-09-07'} WHERE DataEmissao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PedidoCotacaoN].[Data]' 
GO

UPDATE PedidoCotacaoN SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PedidoCotacaoN].[DataValidade]' 
GO

UPDATE PedidoCotacaoN SET DataValidade = {d '1822-09-07'} WHERE DataValidade = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PedidoCotacaoN].[DataBaixa]' 
GO

UPDATE PedidoCotacaoN SET DataBaixa = {d '1822-09-07'} WHERE DataBaixa = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PedidoCotacaoN].[DataFixa]' 
GO

UPDATE PedidoCotacaoN SET DataFixa = {d '1822-09-07'} WHERE DataFixa = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PedidosDeVenda].[DataEmissao]' 
GO

UPDATE PedidosDeVenda SET DataEmissao = {d '1822-09-07'} WHERE DataEmissao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PedidosDeVenda].[DataReferencia]' 
GO

UPDATE PedidosDeVenda SET DataReferencia = {d '1822-09-07'} WHERE DataReferencia = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PedidosDeVenda].[DataEntrega]' 
GO

UPDATE PedidosDeVenda SET DataEntrega = {d '1822-09-07'} WHERE DataEntrega = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PedidosDeVenda].[DataRefFluxo]' 
GO

UPDATE PedidosDeVenda SET DataRefFluxo = {d '1822-09-07'} WHERE DataRefFluxo = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PedidosDeVenda].[DataInclusao]' 
GO

UPDATE PedidosDeVenda SET DataInclusao = {d '1822-09-07'} WHERE DataInclusao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PedidosDeVenda].[DataAlteracao]' 
GO

UPDATE PedidosDeVenda SET DataAlteracao = {d '1822-09-07'} WHERE DataAlteracao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PedidosDeVendaBaixados].[DataEmissao]' 
GO

UPDATE PedidosDeVendaBaixados SET DataEmissao = {d '1822-09-07'} WHERE DataEmissao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PedidosDeVendaBaixados].[DataReferencia]' 
GO

UPDATE PedidosDeVendaBaixados SET DataReferencia = {d '1822-09-07'} WHERE DataReferencia = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PedidosDeVendaBaixados].[DataEntrega]' 
GO

UPDATE PedidosDeVendaBaixados SET DataEntrega = {d '1822-09-07'} WHERE DataEntrega = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PedidosDeVendaBaixados].[DataInclusao]' 
GO

UPDATE PedidosDeVendaBaixados SET DataInclusao = {d '1822-09-07'} WHERE DataInclusao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PedidosDeVendaBaixados].[DataAlteracao]' 
GO

UPDATE PedidosDeVendaBaixados SET DataAlteracao = {d '1822-09-07'} WHERE DataAlteracao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PedidoServico].[DataEmissao]' 
GO

UPDATE PedidoServico SET DataEmissao = {d '1822-09-07'} WHERE DataEmissao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PedidoServico].[DataReferencia]' 
GO

UPDATE PedidoServico SET DataReferencia = {d '1822-09-07'} WHERE DataReferencia = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PedidoServico].[DataEntrega]' 
GO

UPDATE PedidoServico SET DataEntrega = {d '1822-09-07'} WHERE DataEntrega = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PedidoServico].[DataRefFluxo]' 
GO

UPDATE PedidoServico SET DataRefFluxo = {d '1822-09-07'} WHERE DataRefFluxo = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Periodo].[DataInicio]' 
GO

UPDATE Periodo SET DataInicio = {d '1822-09-07'} WHERE DataInicio = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Periodo].[DataFim]' 
GO

UPDATE Periodo SET DataFim = {d '1822-09-07'} WHERE DataFim = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PeriodosFilial].[DataApuracao]' 
GO

UPDATE PeriodosFilial SET DataApuracao = {d '1822-09-07'} WHERE DataApuracao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PlanoConta].[DataRegistro]' 
GO

UPDATE PlanoConta SET DataRegistro = {d '1822-09-07'} WHERE DataRegistro = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PlanoContaHistorico].[DataAtualizacao]' 
GO

UPDATE PlanoContaHistorico SET DataAtualizacao = {d '1822-09-07'} WHERE DataAtualizacao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PlanoContaRef].[ValidadeDe]' 
GO

UPDATE PlanoContaRef SET ValidadeDe = {d '1822-09-07'} WHERE ValidadeDe = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PlanoContaRef].[ValidadeAte]' 
GO

UPDATE PlanoContaRef SET ValidadeAte = {d '1822-09-07'} WHERE ValidadeAte = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PlanoContaRef].[DataInclusaoSist]' 
GO

UPDATE PlanoContaRef SET DataInclusaoSist = {d '1822-09-07'} WHERE DataInclusaoSist = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PlanoContaRef].[DataAlteracaoSist]' 
GO

UPDATE PlanoContaRef SET DataAlteracaoSist = {d '1822-09-07'} WHERE DataAlteracaoSist = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PlanoContaRefModelo].[DataCad]' 
GO

UPDATE PlanoContaRefModelo SET DataCad = {d '1822-09-07'} WHERE DataCad = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PlanoContaRefModelo].[DataAlt]' 
GO

UPDATE PlanoContaRefModelo SET DataAlt = {d '1822-09-07'} WHERE DataAlt = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PlanoMestreProducao].[DataGeracao]' 
GO

UPDATE PlanoMestreProducao SET DataGeracao = {d '1822-09-07'} WHERE DataGeracao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PlanoMestreProducaoItens].[DataNecessidade]' 
GO

UPDATE PlanoMestreProducaoItens SET DataNecessidade = {d '1822-09-07'} WHERE DataNecessidade = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PlanoOperacional].[DataInicio]' 
GO

UPDATE PlanoOperacional SET DataInicio = {d '1822-09-07'} WHERE DataInicio = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PlanoOperacional].[DataFim]' 
GO

UPDATE PlanoOperacional SET DataFim = {d '1822-09-07'} WHERE DataFim = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PlanoOperacionalMaquinas].[Data]' 
GO

UPDATE PlanoOperacionalMaquinas SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PrecoCalculado].[DataReferencia]' 
GO

UPDATE PrecoCalculado SET DataReferencia = {d '1822-09-07'} WHERE DataReferencia = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PrecoCalculado].[DataVigencia]' 
GO

UPDATE PrecoCalculado SET DataVigencia = {d '1822-09-07'} WHERE DataVigencia = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PrevVenda].[DataPrevisao]' 
GO

UPDATE PrevVenda SET DataPrevisao = {d '1822-09-07'} WHERE DataPrevisao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PrevVenda].[DataInicio]' 
GO

UPDATE PrevVenda SET DataInicio = {d '1822-09-07'} WHERE DataInicio = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PrevVenda].[DataFim]' 
GO

UPDATE PrevVenda SET DataFim = {d '1822-09-07'} WHERE DataFim = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PrevVendaMensal].[DataAtualizacao1]' 
GO

UPDATE PrevVendaMensal SET DataAtualizacao1 = {d '1822-09-07'} WHERE DataAtualizacao1 = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PrevVendaMensal].[DataAtualizacao2]' 
GO

UPDATE PrevVendaMensal SET DataAtualizacao2 = {d '1822-09-07'} WHERE DataAtualizacao2 = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PrevVendaMensal].[DataAtualizacao3]' 
GO

UPDATE PrevVendaMensal SET DataAtualizacao3 = {d '1822-09-07'} WHERE DataAtualizacao3 = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PrevVendaMensal].[DataAtualizacao4]' 
GO

UPDATE PrevVendaMensal SET DataAtualizacao4 = {d '1822-09-07'} WHERE DataAtualizacao4 = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PrevVendaMensal].[DataAtualizacao5]' 
GO

UPDATE PrevVendaMensal SET DataAtualizacao5 = {d '1822-09-07'} WHERE DataAtualizacao5 = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PrevVendaMensal].[DataAtualizacao6]' 
GO

UPDATE PrevVendaMensal SET DataAtualizacao6 = {d '1822-09-07'} WHERE DataAtualizacao6 = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PrevVendaMensal].[DataAtualizacao7]' 
GO

UPDATE PrevVendaMensal SET DataAtualizacao7 = {d '1822-09-07'} WHERE DataAtualizacao7 = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PrevVendaMensal].[DataAtualizacao8]' 
GO

UPDATE PrevVendaMensal SET DataAtualizacao8 = {d '1822-09-07'} WHERE DataAtualizacao8 = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PrevVendaMensal].[DataAtualizacao9]' 
GO

UPDATE PrevVendaMensal SET DataAtualizacao9 = {d '1822-09-07'} WHERE DataAtualizacao9 = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PrevVendaMensal].[DataAtualizacao10]' 
GO

UPDATE PrevVendaMensal SET DataAtualizacao10 = {d '1822-09-07'} WHERE DataAtualizacao10 = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PrevVendaMensal].[DataAtualizacao11]' 
GO

UPDATE PrevVendaMensal SET DataAtualizacao11 = {d '1822-09-07'} WHERE DataAtualizacao11 = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PrevVendaMensal].[DataAtualizacao12]' 
GO

UPDATE PrevVendaMensal SET DataAtualizacao12 = {d '1822-09-07'} WHERE DataAtualizacao12 = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PrevVendaPrevConsumo].[Data]' 
GO

UPDATE PrevVendaPrevConsumo SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PRJContratoItens].[DataEntrega]' 
GO

UPDATE PRJContratoItens SET DataEntrega = {d '1822-09-07'} WHERE DataEntrega = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PRJContratos].[Data]' 
GO

UPDATE PRJContratos SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PRJEtapaMaquinas].[Data]' 
GO

UPDATE PRJEtapaMaquinas SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PRJEtapaMateriais].[Data]' 
GO

UPDATE PRJEtapaMateriais SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PRJEtapaMO].[Data]' 
GO

UPDATE PRJEtapaMO SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PRJEtapas].[DataInicio]' 
GO

UPDATE PRJEtapas SET DataInicio = {d '1822-09-07'} WHERE DataInicio = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PRJEtapas].[DataFim]' 
GO

UPDATE PRJEtapas SET DataFim = {d '1822-09-07'} WHERE DataFim = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PRJEtapas].[DataInicioReal]' 
GO

UPDATE PRJEtapas SET DataInicioReal = {d '1822-09-07'} WHERE DataInicioReal = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PRJEtapas].[DataFimReal]' 
GO

UPDATE PRJEtapas SET DataFimReal = {d '1822-09-07'} WHERE DataFimReal = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PRJEtapas].[DataVistoria]' 
GO

UPDATE PRJEtapas SET DataVistoria = {d '1822-09-07'} WHERE DataVistoria = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PRJEtapas].[ValidadeVistoria]' 
GO

UPDATE PRJEtapas SET ValidadeVistoria = {d '1822-09-07'} WHERE ValidadeVistoria = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PRJEtapaVistorias].[Data]' 
GO

UPDATE PRJEtapaVistorias SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PRJEtapaVistorias].[DataValidade]' 
GO

UPDATE PRJEtapaVistorias SET DataValidade = {d '1822-09-07'} WHERE DataValidade = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PRJPropostaItens].[DataEntrega]' 
GO

UPDATE PRJPropostaItens SET DataEntrega = {d '1822-09-07'} WHERE DataEntrega = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PRJPropostas].[Data]' 
GO

UPDATE PRJPropostas SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ProcReajusteTitRec].[DataProc]' 
GO

UPDATE ProcReajusteTitRec SET DataProc = {d '1822-09-07'} WHERE DataProc = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ProcReajusteTitRec].[AtualizadoAte]' 
GO

UPDATE ProcReajusteTitRec SET AtualizadoAte = {d '1822-09-07'} WHERE AtualizadoAte = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ProdutoHistorico].[DataAtualizacao]' 
GO

UPDATE ProdutoHistorico SET DataAtualizacao = {d '1822-09-07'} WHERE DataAtualizacao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Produtos].[DataLog]' 
GO

UPDATE Produtos SET DataLog = {d '1822-09-07'} WHERE DataLog = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Projetos].[DataCriacao]' 
GO

UPDATE Projetos SET DataCriacao = {d '1822-09-07'} WHERE DataCriacao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Projetos].[DataInicio]' 
GO

UPDATE Projetos SET DataInicio = {d '1822-09-07'} WHERE DataInicio = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Projetos].[DataFim]' 
GO

UPDATE Projetos SET DataFim = {d '1822-09-07'} WHERE DataFim = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Projetos].[DataInicioReal]' 
GO

UPDATE Projetos SET DataInicioReal = {d '1822-09-07'} WHERE DataInicioReal = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Projetos].[DataFimReal]' 
GO

UPDATE Projetos SET DataFimReal = {d '1822-09-07'} WHERE DataFimReal = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Projetos_BKP_20130222].[DataCriacao]' 
GO

UPDATE Projetos_BKP_20130222 SET DataCriacao = {d '1822-09-07'} WHERE DataCriacao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Projetos_BKP_20130222].[DataInicio]' 
GO

UPDATE Projetos_BKP_20130222 SET DataInicio = {d '1822-09-07'} WHERE DataInicio = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Projetos_BKP_20130222].[DataFim]' 
GO

UPDATE Projetos_BKP_20130222 SET DataFim = {d '1822-09-07'} WHERE DataFim = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Projetos_BKP_20130222].[DataInicioReal]' 
GO

UPDATE Projetos_BKP_20130222 SET DataInicioReal = {d '1822-09-07'} WHERE DataInicioReal = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Projetos_BKP_20130222].[DataFimReal]' 
GO

UPDATE Projetos_BKP_20130222 SET DataFimReal = {d '1822-09-07'} WHERE DataFimReal = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PVHistAndamento].[Data]' 
GO

UPDATE PVHistAndamento SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[PVHistAndamento].[DataEntrega]' 
GO

UPDATE PVHistAndamento SET DataEntrega = {d '1822-09-07'} WHERE DataEntrega = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RastreamentoLote].[DataValidade]' 
GO

UPDATE RastreamentoLote SET DataValidade = {d '1822-09-07'} WHERE DataValidade = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RastreamentoLote].[DataEntrada]' 
GO

UPDATE RastreamentoLote SET DataEntrada = {d '1822-09-07'} WHERE DataEntrada = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RastreamentoLote].[DataFabricacao]' 
GO

UPDATE RastreamentoLote SET DataFabricacao = {d '1822-09-07'} WHERE DataFabricacao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RastreamentoLoteTeste].[RegistroAnaliseData]' 
GO

UPDATE RastreamentoLoteTeste SET RegistroAnaliseData = {d '1822-09-07'} WHERE RegistroAnaliseData = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RecebPeriodicos].[Inicio]' 
GO

UPDATE RecebPeriodicos SET Inicio = {d '1822-09-07'} WHERE Inicio = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RecebPeriodicos].[Termino]' 
GO

UPDATE RecebPeriodicos SET Termino = {d '1822-09-07'} WHERE Termino = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RecebPeriodicos].[Proximo]' 
GO

UPDATE RecebPeriodicos SET Proximo = {d '1822-09-07'} WHERE Proximo = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Redes].[DataLog]' 
GO

UPDATE Redes SET DataLog = {d '1822-09-07'} WHERE DataLog = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RegApuracaoICMS].[DataInicial]' 
GO

UPDATE RegApuracaoICMS SET DataInicial = {d '1822-09-07'} WHERE DataInicial = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RegApuracaoICMS].[DataFinal]' 
GO

UPDATE RegApuracaoICMS SET DataFinal = {d '1822-09-07'} WHERE DataFinal = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RegApuracaoICMS].[DataEntregaGIA]' 
GO

UPDATE RegApuracaoICMS SET DataEntregaGIA = {d '1822-09-07'} WHERE DataEntregaGIA = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RegApuracaoICMS].[DataImpressao]' 
GO

UPDATE RegApuracaoICMS SET DataImpressao = {d '1822-09-07'} WHERE DataImpressao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RegApuracaoICMSItem].[Data]' 
GO

UPDATE RegApuracaoICMSItem SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RegApuracaoICMSST].[DataInicial]' 
GO

UPDATE RegApuracaoICMSST SET DataInicial = {d '1822-09-07'} WHERE DataInicial = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RegApuracaoICMSST].[DataFinal]' 
GO

UPDATE RegApuracaoICMSST SET DataFinal = {d '1822-09-07'} WHERE DataFinal = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RegApuracaoICMSST].[DataEntregaGIA]' 
GO

UPDATE RegApuracaoICMSST SET DataEntregaGIA = {d '1822-09-07'} WHERE DataEntregaGIA = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RegApuracaoICMSST].[DataImpressao]' 
GO

UPDATE RegApuracaoICMSST SET DataImpressao = {d '1822-09-07'} WHERE DataImpressao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RegApuracaoICMSSTItem].[Data]' 
GO

UPDATE RegApuracaoICMSSTItem SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RegApuracaoIPI].[DataInicial]' 
GO

UPDATE RegApuracaoIPI SET DataInicial = {d '1822-09-07'} WHERE DataInicial = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RegApuracaoIPI].[DataFinal]' 
GO

UPDATE RegApuracaoIPI SET DataFinal = {d '1822-09-07'} WHERE DataFinal = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RegApuracaoIPI].[DataEntregaGIA]' 
GO

UPDATE RegApuracaoIPI SET DataEntregaGIA = {d '1822-09-07'} WHERE DataEntregaGIA = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RegApuracaoIPI].[DataImpressao]' 
GO

UPDATE RegApuracaoIPI SET DataImpressao = {d '1822-09-07'} WHERE DataImpressao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RegApuracaoIPIItem].[Data]' 
GO

UPDATE RegApuracaoIPIItem SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RegInventario].[Data]' 
GO

UPDATE RegInventario SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RegInventarioAlmox].[Data]' 
GO

UPDATE RegInventarioAlmox SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RegInventarioReq].[Data]' 
GO

UPDATE RegInventarioReq SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RegInventarioReq].[DataReq]' 
GO

UPDATE RegInventarioReq SET DataReq = {d '1822-09-07'} WHERE DataReq = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RegraWFW].[DataUltExec]' 
GO

UPDATE RegraWFW SET DataUltExec = {d '1822-09-07'} WHERE DataUltExec = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Rel12Semanas].[D1]' 
GO

UPDATE Rel12Semanas SET D1 = {d '1822-09-07'} WHERE D1 = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Rel12Semanas].[D2]' 
GO

UPDATE Rel12Semanas SET D2 = {d '1822-09-07'} WHERE D2 = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Rel12Semanas].[D3]' 
GO

UPDATE Rel12Semanas SET D3 = {d '1822-09-07'} WHERE D3 = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Rel12Semanas].[D4]' 
GO

UPDATE Rel12Semanas SET D4 = {d '1822-09-07'} WHERE D4 = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Rel12Semanas].[D5]' 
GO

UPDATE Rel12Semanas SET D5 = {d '1822-09-07'} WHERE D5 = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Rel12Semanas].[D6]' 
GO

UPDATE Rel12Semanas SET D6 = {d '1822-09-07'} WHERE D6 = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Rel12Semanas].[D7]' 
GO

UPDATE Rel12Semanas SET D7 = {d '1822-09-07'} WHERE D7 = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Rel12Semanas].[D8]' 
GO

UPDATE Rel12Semanas SET D8 = {d '1822-09-07'} WHERE D8 = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Rel12Semanas].[D9]' 
GO

UPDATE Rel12Semanas SET D9 = {d '1822-09-07'} WHERE D9 = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Rel12Semanas].[D10]' 
GO

UPDATE Rel12Semanas SET D10 = {d '1822-09-07'} WHERE D10 = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Rel12Semanas].[D11]' 
GO

UPDATE Rel12Semanas SET D11 = {d '1822-09-07'} WHERE D11 = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Rel12Semanas].[D12]' 
GO

UPDATE Rel12Semanas SET D12 = {d '1822-09-07'} WHERE D12 = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Rel12Semanas].[DataPrevReceb]' 
GO

UPDATE Rel12Semanas SET DataPrevReceb = {d '1822-09-07'} WHERE DataPrevReceb = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Rel12SemanasProdTemp].[DataIniSemana]' 
GO

UPDATE Rel12SemanasProdTemp SET DataIniSemana = {d '1822-09-07'} WHERE DataIniSemana = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Rel12SemanasProdTemp].[DataFimSemana]' 
GO

UPDATE Rel12SemanasProdTemp SET DataFimSemana = {d '1822-09-07'} WHERE DataFimSemana = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Rel12SemanasProdTemp].[DataPrev]' 
GO

UPDATE Rel12SemanasProdTemp SET DataPrev = {d '1822-09-07'} WHERE DataPrev = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelacionamentoClientes].[Data]' 
GO

UPDATE RelacionamentoClientes SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelacionamentoClientes].[DataProxCobr]' 
GO

UPDATE RelacionamentoClientes SET DataProxCobr = {d '1822-09-07'} WHERE DataProxCobr = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelacionamentoClientes].[DataPrevReceb]' 
GO

UPDATE RelacionamentoClientes SET DataPrevReceb = {d '1822-09-07'} WHERE DataPrevReceb = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelacionamentoClientes].[DataFim]' 
GO

UPDATE RelacionamentoClientes SET DataFim = {d '1822-09-07'} WHERE DataFim = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelacionamentoContatos].[Data]' 
GO

UPDATE RelacionamentoContatos SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelAcompPRJ].[DataIniPrevEtapa]' 
GO

UPDATE RelAcompPRJ SET DataIniPrevEtapa = {d '1822-09-07'} WHERE DataIniPrevEtapa = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelAcompPRJ].[DataFimPrevEtapa]' 
GO

UPDATE RelAcompPRJ SET DataFimPrevEtapa = {d '1822-09-07'} WHERE DataFimPrevEtapa = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelAcompPRJ].[DataIniRealEtapa]' 
GO

UPDATE RelAcompPRJ SET DataIniRealEtapa = {d '1822-09-07'} WHERE DataIniRealEtapa = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelAcompPRJ].[DataFimRealEtapa]' 
GO

UPDATE RelAcompPRJ SET DataFimRealEtapa = {d '1822-09-07'} WHERE DataFimRealEtapa = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelAcompPRJ].[DataIniPrevPRJ]' 
GO

UPDATE RelAcompPRJ SET DataIniPrevPRJ = {d '1822-09-07'} WHERE DataIniPrevPRJ = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelAcompPRJ].[DataFimPrevPRJ]' 
GO

UPDATE RelAcompPRJ SET DataFimPrevPRJ = {d '1822-09-07'} WHERE DataFimPrevPRJ = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelAcompPRJ].[DataIniRealPRJ]' 
GO

UPDATE RelAcompPRJ SET DataIniRealPRJ = {d '1822-09-07'} WHERE DataIniRealPRJ = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelAcompPRJ].[DataFimRealPRJ]' 
GO

UPDATE RelAcompPRJ SET DataFimRealPRJ = {d '1822-09-07'} WHERE DataFimRealPRJ = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelApurICMSSubst].[Data]' 
GO

UPDATE RelApurICMSSubst SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelApurICMSSubst].[DataEmissao]' 
GO

UPDATE RelApurICMSSubst SET DataEmissao = {d '1822-09-07'} WHERE DataEmissao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelBoleto].[Vencimento]' 
GO

UPDATE RelBoleto SET Vencimento = {d '1822-09-07'} WHERE Vencimento = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelBoleto].[DocData]' 
GO

UPDATE RelBoleto SET DocData = {d '1822-09-07'} WHERE DocData = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelBoleto].[DataProc]' 
GO

UPDATE RelBoleto SET DataProc = {d '1822-09-07'} WHERE DataProc = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelBxPorProdDet].[DataEmissaoNF]' 
GO

UPDATE RelBxPorProdDet SET DataEmissaoNF = {d '1822-09-07'} WHERE DataEmissaoNF = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelBxPorProdDet].[DataBaixa]' 
GO

UPDATE RelBxPorProdDet SET DataBaixa = {d '1822-09-07'} WHERE DataBaixa = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelClienteVendedor].[DataPrimeiraCompra]' 
GO

UPDATE RelClienteVendedor SET DataPrimeiraCompra = {d '1822-09-07'} WHERE DataPrimeiraCompra = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelClienteVendedor].[DataUltimaCompra]' 
GO

UPDATE RelClienteVendedor SET DataUltimaCompra = {d '1822-09-07'} WHERE DataUltimaCompra = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelComisVend].[DataRef]' 
GO

UPDATE RelComisVend SET DataRef = {d '1822-09-07'} WHERE DataRef = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelContratoNFiscal].[DataReferencia]' 
GO

UPDATE RelContratoNFiscal SET DataReferencia = {d '1822-09-07'} WHERE DataReferencia = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelContratoNFiscalItens].[DataRefIni]' 
GO

UPDATE RelContratoNFiscalItens SET DataRefIni = {d '1822-09-07'} WHERE DataRefIni = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelContratoNFiscalItens].[DataRefFim]' 
GO

UPDATE RelContratoNFiscalItens SET DataRefFim = {d '1822-09-07'} WHERE DataRefFim = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelContratoNFiscalItens].[DataCobranca]' 
GO

UPDATE RelContratoNFiscalItens SET DataCobranca = {d '1822-09-07'} WHERE DataCobranca = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelCTMaqDispItens].[Data1]' 
GO

UPDATE RelCTMaqDispItens SET Data1 = {d '1822-09-07'} WHERE Data1 = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelCTMaqDispItens].[Data2]' 
GO

UPDATE RelCTMaqDispItens SET Data2 = {d '1822-09-07'} WHERE Data2 = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelCTMaqDispItens].[Data3]' 
GO

UPDATE RelCTMaqDispItens SET Data3 = {d '1822-09-07'} WHERE Data3 = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelCTMaqDispItens].[Data4]' 
GO

UPDATE RelCTMaqDispItens SET Data4 = {d '1822-09-07'} WHERE Data4 = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelCTMaqDispItens].[Data5]' 
GO

UPDATE RelCTMaqDispItens SET Data5 = {d '1822-09-07'} WHERE Data5 = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelCTMaqDispItens].[Data6]' 
GO

UPDATE RelCTMaqDispItens SET Data6 = {d '1822-09-07'} WHERE Data6 = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelCTMaqDispItens].[Data7]' 
GO

UPDATE RelCTMaqDispItens SET Data7 = {d '1822-09-07'} WHERE Data7 = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelDemonstrativo].[Data]' 
GO

UPDATE RelDemonstrativo SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelDiarioCP].[DataContabil]' 
GO

UPDATE RelDiarioCP SET DataContabil = {d '1822-09-07'} WHERE DataContabil = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelDiarioCR].[DataContabil]' 
GO

UPDATE RelDiarioCR SET DataContabil = {d '1822-09-07'} WHERE DataContabil = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelEstatisticaCliente].[DataPrimeiraCompra]' 
GO

UPDATE RelEstatisticaCliente SET DataPrimeiraCompra = {d '1822-09-07'} WHERE DataPrimeiraCompra = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelEstatisticaCliente].[DataUltChequeDevolvido]' 
GO

UPDATE RelEstatisticaCliente SET DataUltChequeDevolvido = {d '1822-09-07'} WHERE DataUltChequeDevolvido = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelEstatisticaCliente].[DataUltimaCompra]' 
GO

UPDATE RelEstatisticaCliente SET DataUltimaCompra = {d '1822-09-07'} WHERE DataUltimaCompra = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelFlCxAn].[DataVenctoReal]' 
GO

UPDATE RelFlCxAn SET DataVenctoReal = {d '1822-09-07'} WHERE DataVenctoReal = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelFlCxAn].[DataVencto]' 
GO

UPDATE RelFlCxAn SET DataVencto = {d '1822-09-07'} WHERE DataVencto = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelFlCxCtb].[Data]' 
GO

UPDATE RelFlCxCtb SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelFluxoCaixa].[Data]' 
GO

UPDATE RelFluxoCaixa SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelFluxoPRJ].[Data]' 
GO

UPDATE RelFluxoPRJ SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelForSaldoIni].[Data]' 
GO

UPDATE RelForSaldoIni SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelGarantia].[DataVenda]' 
GO

UPDATE RelGarantia SET DataVenda = {d '1822-09-07'} WHERE DataVenda = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelIdadeEstoque].[DataEntrada]' 
GO

UPDATE RelIdadeEstoque SET DataEntrada = {d '1822-09-07'} WHERE DataEntrada = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelLcto].[Data]' 
GO

UPDATE RelLcto SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelLcto].[DataEstoque]' 
GO

UPDATE RelLcto SET DataEstoque = {d '1822-09-07'} WHERE DataEstoque = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelLctosCPAux].[Data]' 
GO

UPDATE RelLctosCPAux SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelLogCliente].[Data]' 
GO

UPDATE RelLogCliente SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelMatUtiPerPRJ].[Data]' 
GO

UPDATE RelMatUtiPerPRJ SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelMOUtiPerPRJ].[Data]' 
GO

UPDATE RelMOUtiPerPRJ SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelMovEstTec].[Data]' 
GO

UPDATE RelMovEstTec SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelMvDiaCtaRef].[Data]' 
GO

UPDATE RelMvDiaCtaRef SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelMvLctoCtaRef].[DataLcto]' 
GO

UPDATE RelMvLctoCtaRef SET DataLcto = {d '1822-09-07'} WHERE DataLcto = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelOPGerEstTA].[DataPrev]' 
GO

UPDATE RelOPGerEstTA SET DataPrev = {d '1822-09-07'} WHERE DataPrev = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelOPGerEstTA].[DataIniSemana]' 
GO

UPDATE RelOPGerEstTA SET DataIniSemana = {d '1822-09-07'} WHERE DataIniSemana = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelOPGerEstTA].[DataFimSemana]' 
GO

UPDATE RelOPGerEstTA SET DataFimSemana = {d '1822-09-07'} WHERE DataFimSemana = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelOPUsoMaquina].[Data]' 
GO

UPDATE RelOPUsoMaquina SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelPCPCompConsumo].[DataFinalProducao]' 
GO

UPDATE RelPCPCompConsumo SET DataFinalProducao = {d '1822-09-07'} WHERE DataFinalProducao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelPreviaCargaCT].[Data]' 
GO

UPDATE RelPreviaCargaCT SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelPreviaReqCompra].[DataNecessidade]' 
GO

UPDATE RelPreviaReqCompra SET DataNecessidade = {d '1822-09-07'} WHERE DataNecessidade = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelPreviaReqCompra].[DataLimite]' 
GO

UPDATE RelPreviaReqCompra SET DataLimite = {d '1822-09-07'} WHERE DataLimite = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelProdEst].[DataInicioOP]' 
GO

UPDATE RelProdEst SET DataInicioOP = {d '1822-09-07'} WHERE DataInicioOP = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelRankingProdCab].[DataDe]' 
GO

UPDATE RelRankingProdCab SET DataDe = {d '1822-09-07'} WHERE DataDe = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelRankingProdCab].[DataAte]' 
GO

UPDATE RelRankingProdCab SET DataAte = {d '1822-09-07'} WHERE DataAte = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelRankingProdCab].[DataGer]' 
GO

UPDATE RelRankingProdCab SET DataGer = {d '1822-09-07'} WHERE DataGer = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelRCPEMod3Cab].[DataDe]' 
GO

UPDATE RelRCPEMod3Cab SET DataDe = {d '1822-09-07'} WHERE DataDe = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelRCPEMod3Cab].[DataAte]' 
GO

UPDATE RelRCPEMod3Cab SET DataAte = {d '1822-09-07'} WHERE DataAte = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelRCPEMod3Det].[DataMov]' 
GO

UPDATE RelRCPEMod3Det SET DataMov = {d '1822-09-07'} WHERE DataMov = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelRCPEMod3Det].[DataEmissaoNF]' 
GO

UPDATE RelRCPEMod3Det SET DataEmissaoNF = {d '1822-09-07'} WHERE DataEmissaoNF = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelRecibo].[DataBaixa]' 
GO

UPDATE RelRecibo SET DataBaixa = {d '1822-09-07'} WHERE DataBaixa = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelRecibo].[DataRegistro]' 
GO

UPDATE RelRecibo SET DataRegistro = {d '1822-09-07'} WHERE DataRegistro = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelRecPorProdDet].[DataEmissaoNF]' 
GO

UPDATE RelRecPorProdDet SET DataEmissaoNF = {d '1822-09-07'} WHERE DataEmissaoNF = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelRecPorProdDet].[DataVencimento]' 
GO

UPDATE RelRecPorProdDet SET DataVencimento = {d '1822-09-07'} WHERE DataVencimento = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelRegApurIPI].[DataDe]' 
GO

UPDATE RelRegApurIPI SET DataDe = {d '1822-09-07'} WHERE DataDe = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelRegApurIPI].[DataAte]' 
GO

UPDATE RelRegApurIPI SET DataAte = {d '1822-09-07'} WHERE DataAte = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelRoteiro].[Data]' 
GO

UPDATE RelRoteiro SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelRotuloProducao].[DataValidade]' 
GO

UPDATE RelRotuloProducao SET DataValidade = {d '1822-09-07'} WHERE DataValidade = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelRotuloProducao].[DataFabricacao]' 
GO

UPDATE RelRotuloProducao SET DataFabricacao = {d '1822-09-07'} WHERE DataFabricacao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelSimulacaoESTItens].[Data]' 
GO

UPDATE RelSimulacaoESTItens SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelTAPrevReserva].[DataReservaPC1]' 
GO

UPDATE RelTAPrevReserva SET DataReservaPC1 = {d '1822-09-07'} WHERE DataReservaPC1 = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelTAPrevReserva].[DataReservaPC2]' 
GO

UPDATE RelTAPrevReserva SET DataReservaPC2 = {d '1822-09-07'} WHERE DataReservaPC2 = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelTAPrevReserva].[DataReservaPC3]' 
GO

UPDATE RelTAPrevReserva SET DataReservaPC3 = {d '1822-09-07'} WHERE DataReservaPC3 = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelVendEstProdCab].[DataDe]' 
GO

UPDATE RelVendEstProdCab SET DataDe = {d '1822-09-07'} WHERE DataDe = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelVendEstProdCab].[DataAte]' 
GO

UPDATE RelVendEstProdCab SET DataAte = {d '1822-09-07'} WHERE DataAte = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RelVendEstProdCab].[DataGer]' 
GO

UPDATE RelVendEstProdCab SET DataGer = {d '1822-09-07'} WHERE DataGer = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RequisicaoCompraN].[Data]' 
GO

UPDATE RequisicaoCompraN SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RequisicaoCompraN].[DataEnvio]' 
GO

UPDATE RequisicaoCompraN SET DataEnvio = {d '1822-09-07'} WHERE DataEnvio = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RequisicaoCompraN].[DataLimite]' 
GO

UPDATE RequisicaoCompraN SET DataLimite = {d '1822-09-07'} WHERE DataLimite = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RequisicaoCompraN].[DataBaixa]' 
GO

UPDATE RequisicaoCompraN SET DataBaixa = {d '1822-09-07'} WHERE DataBaixa = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RequisicaoCompraN].[DataReg]' 
GO

UPDATE RequisicaoCompraN SET DataReg = {d '1822-09-07'} WHERE DataReg = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RequisicaoCompraN].[DataRegEnvio]' 
GO

UPDATE RequisicaoCompraN SET DataRegEnvio = {d '1822-09-07'} WHERE DataRegEnvio = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RequisicaoCompraN].[DataRegAprov]' 
GO

UPDATE RequisicaoCompraN SET DataRegAprov = {d '1822-09-07'} WHERE DataRegAprov = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RequisicaoCompraN].[DataRegBaixa]' 
GO

UPDATE RequisicaoCompraN SET DataRegBaixa = {d '1822-09-07'} WHERE DataRegBaixa = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Reserva].[DataReserva]' 
GO

UPDATE Reserva SET DataReserva = {d '1822-09-07'} WHERE DataReserva = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Reserva].[DataValidade]' 
GO

UPDATE Reserva SET DataValidade = {d '1822-09-07'} WHERE DataValidade = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RetCobrErros].[DataArq]' 
GO

UPDATE RetCobrErros SET DataArq = {d '1822-09-07'} WHERE DataArq = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RetCobrErros].[DataVencimento]' 
GO

UPDATE RetCobrErros SET DataVencimento = {d '1822-09-07'} WHERE DataVencimento = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RetCobrErros].[DataCredito]' 
GO

UPDATE RetCobrErros SET DataCredito = {d '1822-09-07'} WHERE DataCredito = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RetPagto].[DataImport]' 
GO

UPDATE RetPagto SET DataImport = {d '1822-09-07'} WHERE DataImport = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RetPagto].[DataGeracao]' 
GO

UPDATE RetPagto SET DataGeracao = {d '1822-09-07'} WHERE DataGeracao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RetPagtoDet].[DataPagto]' 
GO

UPDATE RetPagtoDet SET DataPagto = {d '1822-09-07'} WHERE DataPagto = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RetPagtoDet].[DataReal]' 
GO

UPDATE RetPagtoDet SET DataReal = {d '1822-09-07'} WHERE DataReal = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RoteirosDeFabricacao].[DataCriacao]' 
GO

UPDATE RoteirosDeFabricacao SET DataCriacao = {d '1822-09-07'} WHERE DataCriacao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RoteirosDeFabricacao].[DataUltModificacao]' 
GO

UPDATE RoteirosDeFabricacao SET DataUltModificacao = {d '1822-09-07'} WHERE DataUltModificacao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RoteiroSRV].[DataCriacao]' 
GO

UPDATE RoteiroSRV SET DataCriacao = {d '1822-09-07'} WHERE DataCriacao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RoteiroSRV].[DataUltModificacao]' 
GO

UPDATE RoteiroSRV SET DataUltModificacao = {d '1822-09-07'} WHERE DataUltModificacao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RPS].[DataEmissao]' 
GO

UPDATE RPS SET DataEmissao = {d '1822-09-07'} WHERE DataEmissao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RPS].[DataUltAlteracao]' 
GO

UPDATE RPS SET DataUltAlteracao = {d '1822-09-07'} WHERE DataUltAlteracao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RPSCab].[DataGeracao]' 
GO

UPDATE RPSCab SET DataGeracao = {d '1822-09-07'} WHERE DataGeracao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RPSCab].[DataInicio]' 
GO

UPDATE RPSCab SET DataInicio = {d '1822-09-07'} WHERE DataInicio = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RPSCab].[DataFim]' 
GO

UPDATE RPSCab SET DataFim = {d '1822-09-07'} WHERE DataFim = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RPSEnviados].[DataEmissao]' 
GO

UPDATE RPSEnviados SET DataEmissao = {d '1822-09-07'} WHERE DataEmissao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RPSWEBConsLote].[Data]' 
GO

UPDATE RPSWEBConsLote SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RPSWEBConsSitLote].[Data]' 
GO

UPDATE RPSWEBConsSitLote SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RPSWEBLote].[Data]' 
GO

UPDATE RPSWEBLote SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RPSWEBLoteLog].[Data]' 
GO

UPDATE RPSWEBLoteLog SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RPSWEBProt].[Data]' 
GO

UPDATE RPSWEBProt SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RPSWEBProt].[DataEmissao]' 
GO

UPDATE RPSWEBProt SET DataEmissao = {d '1822-09-07'} WHERE DataEmissao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RPSWEBProt].[DataEmissaoRPS]' 
GO

UPDATE RPSWEBProt SET DataEmissaoRPS = {d '1822-09-07'} WHERE DataEmissaoRPS = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RPSWEBRetCanc].[Data]' 
GO

UPDATE RPSWEBRetCanc SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RPSWEBRetEnvi].[Data]' 
GO

UPDATE RPSWEBRetEnvi SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[RPSWEBRetEnvi].[DataRecebimento]' 
GO

UPDATE RPSWEBRetEnvi SET DataRecebimento = {d '1822-09-07'} WHERE DataRecebimento = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Sessao].[DataInicio]' 
GO

UPDATE Sessao SET DataInicio = {d '1822-09-07'} WHERE DataInicio = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Sessao].[DataFim]' 
GO

UPDATE Sessao SET DataFim = {d '1822-09-07'} WHERE DataFim = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[SldDiaEst].[Data]' 
GO

UPDATE SldDiaEst SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[SldDiaEstAlm].[Data]' 
GO

UPDATE SldDiaEstAlm SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[SldDiaEstTerc].[Data]' 
GO

UPDATE SldDiaEstTerc SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[SldDiaFat].[Data]' 
GO

UPDATE SldDiaFat SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[SldDiaFatCx].[Data]' 
GO

UPDATE SldDiaFatCx SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[SldDiaForn].[Data]' 
GO

UPDATE SldDiaForn SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[SldDiaMeioPagtoCx].[Data]' 
GO

UPDATE SldDiaMeioPagtoCx SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[SolicitacaoSRV].[Data]' 
GO

UPDATE SolicitacaoSRV SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[SolicitacaoSRV].[DataEntrega]' 
GO

UPDATE SolicitacaoSRV SET DataEntrega = {d '1822-09-07'} WHERE DataEntrega = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[SpedCtbVersaoLeiaute].[DataInicio]' 
GO

UPDATE SpedCtbVersaoLeiaute SET DataInicio = {d '1822-09-07'} WHERE DataInicio = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[SpedDocFiscais].[DataEmissao]' 
GO

UPDATE SpedDocFiscais SET DataEmissao = {d '1822-09-07'} WHERE DataEmissao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[SpedDocFiscais].[DataES]' 
GO

UPDATE SpedDocFiscais SET DataES = {d '1822-09-07'} WHERE DataES = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[SpedECFVersaoLeiaute].[DataInicio]' 
GO

UPDATE SpedECFVersaoLeiaute SET DataInicio = {d '1822-09-07'} WHERE DataInicio = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[SpedFisArrecadacaoRef].[DataVcto]' 
GO

UPDATE SpedFisArrecadacaoRef SET DataVcto = {d '1822-09-07'} WHERE DataVcto = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[SpedFisArrecadacaoRef].[DataPagto]' 
GO

UPDATE SpedFisArrecadacaoRef SET DataPagto = {d '1822-09-07'} WHERE DataPagto = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[SpedFisCupomRef].[DataDoc]' 
GO

UPDATE SpedFisCupomRef SET DataDoc = {d '1822-09-07'} WHERE DataDoc = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[SpedFisDocRef].[DataDoc]' 
GO

UPDATE SpedFisDocRef SET DataDoc = {d '1822-09-07'} WHERE DataDoc = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[SpedFisPisVersaoLeiaute].[DataInicio]' 
GO

UPDATE SpedFisPisVersaoLeiaute SET DataInicio = {d '1822-09-07'} WHERE DataInicio = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[SpedFisRG110].[DataInicial]' 
GO

UPDATE SpedFisRG110 SET DataInicial = {d '1822-09-07'} WHERE DataInicial = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[SpedFisRG110].[DataFinal]' 
GO

UPDATE SpedFisRG110 SET DataFinal = {d '1822-09-07'} WHERE DataFinal = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[SpedFisRG125].[DataMovto]' 
GO

UPDATE SpedFisRG125 SET DataMovto = {d '1822-09-07'} WHERE DataMovto = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[SpedFisRG126].[DataInicial]' 
GO

UPDATE SpedFisRG126 SET DataInicial = {d '1822-09-07'} WHERE DataInicial = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[SpedFisRG126].[DataFinal]' 
GO

UPDATE SpedFisRG126 SET DataFinal = {d '1822-09-07'} WHERE DataFinal = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[SpedFisRG130G140].[DataEmissaoDoc]' 
GO

UPDATE SpedFisRG130G140 SET DataEmissaoDoc = {d '1822-09-07'} WHERE DataEmissaoDoc = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[SpedFisVersaoLeiaute].[DataInicio]' 
GO

UPDATE SpedFisVersaoLeiaute SET DataInicio = {d '1822-09-07'} WHERE DataInicio = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[TabeladeMoedas].[DataExcPtax]' 
GO

UPDATE TabeladeMoedas SET DataExcPtax = {d '1822-09-07'} WHERE DataExcPtax = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[TabelasDePreco].[DataLog]' 
GO

UPDATE TabelasDePreco SET DataLog = {d '1822-09-07'} WHERE DataLog = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[TabelasDePrecoItens].[DataVigencia]' 
GO

UPDATE TabelasDePrecoItens SET DataVigencia = {d '1822-09-07'} WHERE DataVigencia = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[TabelasDePrecoItens].[DataLog]' 
GO

UPDATE TabelasDePrecoItens SET DataLog = {d '1822-09-07'} WHERE DataLog = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[TaxaDeProducao].[Data]' 
GO

UPDATE TaxaDeProducao SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[TaxaDeProducao].[DataDesativacao]' 
GO

UPDATE TaxaDeProducao SET DataDesativacao = {d '1822-09-07'} WHERE DataDesativacao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[TipoFreteFP].[DataAtualizacao]' 
GO

UPDATE TipoFreteFP SET DataAtualizacao = {d '1822-09-07'} WHERE DataAtualizacao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[TitPagDtCtb].[DataContabil]' 
GO

UPDATE TitPagDtCtb SET DataContabil = {d '1822-09-07'} WHERE DataContabil = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[TitulosPag].[DataEmissao]' 
GO

UPDATE TitulosPag SET DataEmissao = {d '1822-09-07'} WHERE DataEmissao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[TitulosPag].[DataRegistro]' 
GO

UPDATE TitulosPag SET DataRegistro = {d '1822-09-07'} WHERE DataRegistro = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[TitulosPagBaixados].[DataEmissao]' 
GO

UPDATE TitulosPagBaixados SET DataEmissao = {d '1822-09-07'} WHERE DataEmissao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[TitulosPagBaixados].[DataRegistro]' 
GO

UPDATE TitulosPagBaixados SET DataRegistro = {d '1822-09-07'} WHERE DataRegistro = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[TitulosRec].[DataEmissao]' 
GO

UPDATE TitulosRec SET DataEmissao = {d '1822-09-07'} WHERE DataEmissao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[TitulosRec].[DataRegistro]' 
GO

UPDATE TitulosRec SET DataRegistro = {d '1822-09-07'} WHERE DataRegistro = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[TitulosRec].[ReajusteBase]' 
GO

UPDATE TitulosRec SET ReajusteBase = {d '1822-09-07'} WHERE ReajusteBase = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[TitulosRec].[ReajustadoAte]' 
GO

UPDATE TitulosRec SET ReajustadoAte = {d '1822-09-07'} WHERE ReajustadoAte = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[TitulosRecBaixados].[DataEmissao]' 
GO

UPDATE TitulosRecBaixados SET DataEmissao = {d '1822-09-07'} WHERE DataEmissao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[TitulosRecBaixados].[DataRegistro]' 
GO

UPDATE TitulosRecBaixados SET DataRegistro = {d '1822-09-07'} WHERE DataRegistro = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[TitulosRecBaixados].[ReajusteBase]' 
GO

UPDATE TitulosRecBaixados SET ReajusteBase = {d '1822-09-07'} WHERE ReajusteBase = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[TitulosRecBaixados].[ReajustadoAte]' 
GO

UPDATE TitulosRecBaixados SET ReajustadoAte = {d '1822-09-07'} WHERE ReajustadoAte = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[TransfCartCobr].[Data]' 
GO

UPDATE TransfCartCobr SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[TransfCartCobr].[DataRegistro]' 
GO

UPDATE TransfCartCobr SET DataRegistro = {d '1822-09-07'} WHERE DataRegistro = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[TransferenciaCaixa].[DataBackoffice]' 
GO

UPDATE TransferenciaCaixa SET DataBackoffice = {d '1822-09-07'} WHERE DataBackoffice = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[TransferenciaLoja].[DataBackoffice]' 
GO

UPDATE TransferenciaLoja SET DataBackoffice = {d '1822-09-07'} WHERE DataBackoffice = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[TransportadoraHistorico].[DataAtualizacao]' 
GO

UPDATE TransportadoraHistorico SET DataAtualizacao = {d '1822-09-07'} WHERE DataAtualizacao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[TribBaseDupla].[DataIni]' 
GO

UPDATE TribBaseDupla SET DataIni = {d '1822-09-07'} WHERE DataIni = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[TribResumoDiaCupom].[Data]' 
GO

UPDATE TribResumoDiaCupom SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[TributacaoDoc].[DataPrestServico]' 
GO

UPDATE TributacaoDoc SET DataPrestServico = {d '1822-09-07'} WHERE DataPrestServico = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[UnidadesDeMedida].[DataLog]' 
GO

UPDATE UnidadesDeMedida SET DataLog = {d '1822-09-07'} WHERE DataLog = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Vendedores].[DataUltVenda]' 
GO

UPDATE Vendedores SET DataUltVenda = {d '1822-09-07'} WHERE DataUltVenda = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[Vendedores].[DataLog]' 
GO

UPDATE Vendedores SET DataLog = {d '1822-09-07'} WHERE DataLog = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[VersaoBD].[DataAtualizacao]' 
GO

UPDATE VersaoBD SET DataAtualizacao = {d '1822-09-07'} WHERE DataAtualizacao = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[XXXItensSolicitacaoDeCompra].[DataEntrega]' 
GO

UPDATE XXXItensSolicitacaoDeCompra SET DataEntrega = {d '1822-09-07'} WHERE DataEntrega = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ZZLanPendenteSalva].[Data]' 
GO

UPDATE ZZLanPendenteSalva SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ZZLanPrePendente].[Data]' 
GO

UPDATE ZZLanPrePendente SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[ZZLanPrePendenteBaixado].[Data]' 
GO

UPDATE ZZLanPrePendenteBaixado SET Data = {d '1822-09-07'} WHERE Data = {d '1822-07-09'}
GO

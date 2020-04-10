VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCRFATSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'??? por que as funcoes ???_Lock estao fazendo lock exclusive ??

Option Explicit

'Incluído por Luiz Nogueira em 27/10/03
Private Declare Function Comando_BindVarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_BindVar" (ByVal lComando As Long, lpVar As Variant) As Long
Private Declare Function Comando_PrepararInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Preparar" (ByVal lComando As Long, ByVal lpSQLStmt As String) As Long
Private Declare Function Comando_ExecutarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Executar" (ByVal lComando As Long) As Long

Function DebitoReceber_Le(ByVal objDebitoRecCli As ClassDebitoRecCli, Optional ByVal iExclusao As Integer = 0) As Long
'Le o debito a receber a partir da chave fornecida em objDebitoRecCli. Retorna os dados em objDebitosRecCli

Dim lComando As Long
Dim lErro As Long
Dim tDebitoRecCli As typeDebitosRecCli

On Error GoTo Erro_DebitoReceber_Le

    tDebitoRecCli.sSiglaDocumento = String(STRING_DEB_REC_SIGLA, 0)
    tDebitoRecCli.sObservacao = String(STRING_DEB_REC_OBS, 0)

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 17832

    If iExclusao = 1 Then
    
        'Pega os dados da tabela DebitosRecCli
        lErro = Comando_Executar(lComando, "SELECT FilialEmpresa, Cliente, Filial, SiglaDocumento, NumTitulo, Status, DataEmissao, ValorTotal, Saldo, ValorSeguro, ValorFrete, OutrasDespesas, ValorProdutos, ValorICMS, ValorICMSSubst, ValorIPI, ValorIRRF, Observacao, PISRetido, COFINSRetido, CSLLRetido FROM DebitosRecCli WHERE NumIntDoc = ?", _
            tDebitoRecCli.iFilialEmpresa, tDebitoRecCli.lCliente, tDebitoRecCli.iFilial, tDebitoRecCli.sSiglaDocumento, tDebitoRecCli.lNumTitulo, tDebitoRecCli.iStatus, tDebitoRecCli.dtDataEmissao, tDebitoRecCli.dValorTotal, tDebitoRecCli.dSaldo, tDebitoRecCli.dValorSeguro, tDebitoRecCli.dValorFrete, tDebitoRecCli.dOutrasDespesas, tDebitoRecCli.dValorProdutos, tDebitoRecCli.dValorICMS, tDebitoRecCli.dValorICMSSubst, tDebitoRecCli.dValorIPI, tDebitoRecCli.dValorIRRF, tDebitoRecCli.sObservacao, tDebitoRecCli.dPISRetido, tDebitoRecCli.dCOFINSRetido, tDebitoRecCli.dCSLLRetido, objDebitoRecCli.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then Error 17833
        
    Else

        'Pega os dados da tabela DebitosRecCli
        lErro = Comando_Executar(lComando, "SELECT FilialEmpresa, Cliente, Filial, SiglaDocumento, NumTitulo, Status, DataEmissao, ValorTotal, Saldo, ValorSeguro, ValorFrete, OutrasDespesas, ValorProdutos, ValorICMS, ValorICMSSubst, ValorIPI, ValorIRRF, Observacao, PISRetido, COFINSRetido, CSLLRetido FROM DebitosRecCli WHERE NumIntDoc = ? AND Status <> ?", _
            tDebitoRecCli.iFilialEmpresa, tDebitoRecCli.lCliente, tDebitoRecCli.iFilial, tDebitoRecCli.sSiglaDocumento, tDebitoRecCli.lNumTitulo, tDebitoRecCli.iStatus, tDebitoRecCli.dtDataEmissao, tDebitoRecCli.dValorTotal, tDebitoRecCli.dSaldo, tDebitoRecCli.dValorSeguro, tDebitoRecCli.dValorFrete, tDebitoRecCli.dOutrasDespesas, tDebitoRecCli.dValorProdutos, tDebitoRecCli.dValorICMS, tDebitoRecCli.dValorICMSSubst, tDebitoRecCli.dValorIPI, tDebitoRecCli.dValorIRRF, tDebitoRecCli.sObservacao, tDebitoRecCli.dPISRetido, tDebitoRecCli.dCOFINSRetido, tDebitoRecCli.dCSLLRetido, objDebitoRecCli.lNumIntDoc, STATUS_EXCLUIDO)
        If lErro <> AD_SQL_SUCESSO Then Error 17833

    End If
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 17834

    If lErro = AD_SQL_SEM_DADOS Then Error 17835

    With objDebitoRecCli
        .iFilialEmpresa = tDebitoRecCli.iFilialEmpresa
        .lCliente = tDebitoRecCli.lCliente
        .iFilial = tDebitoRecCli.iFilial
        .sSiglaDocumento = tDebitoRecCli.sSiglaDocumento
        .lNumTitulo = tDebitoRecCli.lNumTitulo
        .iStatus = tDebitoRecCli.iStatus
        .dtDataEmissao = tDebitoRecCli.dtDataEmissao
        .dValorTotal = tDebitoRecCli.dValorTotal
        .dSaldo = tDebitoRecCli.dSaldo
        .dValorSeguro = tDebitoRecCli.dValorSeguro
        .dValorFrete = tDebitoRecCli.dValorFrete
        .dOutrasDespesas = tDebitoRecCli.dOutrasDespesas
        .dValorProdutos = tDebitoRecCli.dValorProdutos
        .dValorICMS = tDebitoRecCli.dValorICMS
        .dValorICMSSubst = tDebitoRecCli.dValorICMSSubst
        .dValorIPI = tDebitoRecCli.dValorIPI
        .dValorIRRF = tDebitoRecCli.dValorIRRF
        .sObservacao = tDebitoRecCli.sObservacao
        .dPISRetido = tDebitoRecCli.dPISRetido
        .dCOFINSRetido = tDebitoRecCli.dCOFINSRetido
        .dCSLLRetido = tDebitoRecCli.dCSLLRetido
    End With
    
    lErro = Comando_Fechar(lComando)

    DebitoReceber_Le = SUCESSO

    Exit Function

Erro_DebitoReceber_Le:

    DebitoReceber_Le = Err

    Select Case Err

        Case 17832
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 17833, 17834
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_DEBITOSRECCLI", Err, objDebitoRecCli.lNumIntDoc)

        Case 17835

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149050)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ParcelaReceber_Le(ByVal objParcelaReceber As ClassParcelaReceber) As Long
'le dados de uma parcela de um titulo a receber identificada por objParcelaReceber.NumIntDoc

Dim lErro As Long
Dim lComando As Long
Dim tParcelaReceber As typeParcelaReceber

On Error GoTo Erro_ParcelaReceber_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 19144
    
    tParcelaReceber.sNumTitCobrador = String(STRING_NOSSO_NUMERO, 0)
    
    '###############################################
    'INSERIDO POR WAGNER
    tParcelaReceber.sObservacao = String(STRING_MAX_NOME_ARQUIVO, 0)
    '###############################################
    
    tParcelaReceber.sAgenciaCheque = String(STRING_AGENCIA, 0)
    tParcelaReceber.sAprovacaoCartao = String(STRING_APROVACAO_CARTAO, 0)
    tParcelaReceber.sContaCorrenteCheque = String(STRING_CONTA_CORRENTE, 0)
    tParcelaReceber.sNumeroCartao = String(STRING_NUMERO_CARTAO, 0)
    
    'Janaina
    '#################################################
    'ALTERADO POR WAGNER
    lErro = Comando_Executar(lComando, "SELECT NumIntTitulo,NumParcela,Status,DataVencimento,DataVencimentoReal,Saldo,Valor,Cobrador,CarteiraCobranca,NumTitCobrador,ProxSeqBaixa,ProxSeqOcorr,Desconto1Codigo,Desconto1Ate,Desconto1ValoR,Desconto2Codigo,Desconto2Ate,Desconto2Valor,Desconto3Codigo,Desconto3Ate,Desconto3Valor,Aceite,Descontada, Previsao, Observacao, TipoPagto, CodConta, DataCredito, DataEmissaoCheque, BancoCheque, AgenciaCheque, ContaCorrenteCheque, NumeroCheque, DataDepositoCheque, AdmMeioPagto, Parcelamento, NumeroCartao, ValidadeCartao, AprovacaoCartao, DataTransacaoCartao FROM ParcelasRec WHERE NumIntDoc = ?", _
    tParcelaReceber.lNumIntTitulo, tParcelaReceber.iNumParcela, tParcelaReceber.iStatus, tParcelaReceber.dtDataVencimento, tParcelaReceber.dtDataVencimentoReal, tParcelaReceber.dSaldo, tParcelaReceber.dValor, tParcelaReceber.iCobrador, tParcelaReceber.iCarteiraCobranca, tParcelaReceber.sNumTitCobrador, tParcelaReceber.iProxSeqBaixa, tParcelaReceber.iProxSeqOcorr, tParcelaReceber.iDesconto1Codigo, _
    tParcelaReceber.dtDesconto1Ate, tParcelaReceber.dDesconto1Valor, tParcelaReceber.iDesconto2Codigo, tParcelaReceber.dtDesconto2Ate, tParcelaReceber.dDesconto2Valor, tParcelaReceber.iDesconto3Codigo, tParcelaReceber.dtDesconto3Ate, tParcelaReceber.dDesconto3Valor, tParcelaReceber.iAceite, tParcelaReceber.iDescontada, tParcelaReceber.iPrevisao, tParcelaReceber.sObservacao, _
    tParcelaReceber.iTipoPagto, tParcelaReceber.iCodConta, tParcelaReceber.dtDataCredito, tParcelaReceber.dtDataEmissaoCheque, tParcelaReceber.iBancoCheque, tParcelaReceber.sAgenciaCheque, tParcelaReceber.sContaCorrenteCheque, tParcelaReceber.lNumeroCheque, tParcelaReceber.dtDataDepositoCheque, tParcelaReceber.iAdmMeioPagto, tParcelaReceber.iParcelamento, tParcelaReceber.sNumeroCartao, tParcelaReceber.dtValidadeCartao, tParcelaReceber.sAprovacaoCartao, tParcelaReceber.dtDataTransacaoCartao, objParcelaReceber.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then Error 19145
    '#################################################
    'Janaina
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 19146
    
    If lErro = AD_SQL_SEM_DADOS Then Error 19147
    
    With tParcelaReceber

        objParcelaReceber.dDesconto1Valor = .dDesconto1Valor
        objParcelaReceber.dDesconto2Valor = .dDesconto2Valor
        objParcelaReceber.dDesconto3Valor = .dDesconto3Valor
        objParcelaReceber.dSaldo = .dSaldo
        objParcelaReceber.dtDataVencimento = .dtDataVencimento
        objParcelaReceber.dtDataVencimentoReal = .dtDataVencimentoReal
        objParcelaReceber.dtDesconto1Ate = .dtDesconto1Ate
        objParcelaReceber.dtDesconto2Ate = .dtDesconto2Ate
        objParcelaReceber.dtDesconto3Ate = .dtDesconto3Ate
        objParcelaReceber.dValor = .dValor
        objParcelaReceber.iAceite = .iAceite
        objParcelaReceber.iCarteiraCobranca = .iCarteiraCobranca
        objParcelaReceber.iCobrador = .iCobrador
        objParcelaReceber.iDescontada = .iDescontada
        objParcelaReceber.iDesconto1Codigo = .iDesconto1Codigo
        objParcelaReceber.iDesconto2Codigo = .iDesconto2Codigo
        objParcelaReceber.iDesconto3Codigo = .iDesconto3Codigo
        objParcelaReceber.iNumParcela = .iNumParcela
        objParcelaReceber.iProxSeqBaixa = .iProxSeqBaixa
        objParcelaReceber.iProxSeqOcorr = .iProxSeqOcorr
        objParcelaReceber.iStatus = .iStatus
        objParcelaReceber.lNumIntTitulo = .lNumIntTitulo
        objParcelaReceber.sNumTitCobrador = .sNumTitCobrador
        objParcelaReceber.iTipoPagto = .iTipoPagto
        objParcelaReceber.iCodConta = .iCodConta
        objParcelaReceber.dtDataCredito = .dtDataCredito
        objParcelaReceber.dtDataEmissaoCheque = .dtDataEmissaoCheque
        objParcelaReceber.iBancoCheque = .iBancoCheque
        objParcelaReceber.sAgenciaCheque = .sAgenciaCheque
        objParcelaReceber.sContaCorrenteCheque = .sContaCorrenteCheque
        objParcelaReceber.lNumeroCheque = .lNumeroCheque
        objParcelaReceber.dtDataDepositoCheque = .dtDataDepositoCheque
        objParcelaReceber.iAdmMeioPagto = .iAdmMeioPagto
        objParcelaReceber.iParcelamento = .iParcelamento
        objParcelaReceber.sNumeroCartao = .sNumeroCartao
        objParcelaReceber.dtValidadeCartao = .dtValidadeCartao
        objParcelaReceber.sAprovacaoCartao = .sAprovacaoCartao
        objParcelaReceber.dtDataTransacaoCartao = .dtDataTransacaoCartao
        
    
        '#################################################
        'ALTERADO POR WAGNER
        objParcelaReceber.iPrevisao = .iPrevisao
        objParcelaReceber.sObservacao = .sObservacao
        '#################################################
    
    End With
    
    Call Comando_Fechar(lComando)
    
    ParcelaReceber_Le = SUCESSO
    
    Exit Function
    
Erro_ParcelaReceber_Le:

    ParcelaReceber_Le = Err
    
    Select Case Err
    
        Case 19144
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 19145, 19146
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PARCELAS_REC", Err)
            
        Case 19147
            'Tratado pela rotina chamadora
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149051)
    
    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ParcelaReceber_Baixada_Le(ByVal objParcelaReceber As ClassParcelaReceber) As Long
'le dados de uma parcela Baixada de um titulo a receber identificada por objParcelaReceber.NumIntDoc

Dim lErro As Long
Dim lComando As Long
Dim tParcelaReceber As typeParcelaReceber

On Error GoTo Erro_ParcelaReceber_Baixada_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 58556
    
    tParcelaReceber.sNumTitCobrador = String(STRING_NOSSO_NUMERO, 0)
    
    tParcelaReceber.sAgenciaCheque = String(STRING_AGENCIA, 0)
    tParcelaReceber.sAprovacaoCartao = String(STRING_APROVACAO_CARTAO, 0)
    tParcelaReceber.sContaCorrenteCheque = String(STRING_CONTA_CORRENTE, 0)
    tParcelaReceber.sNumeroCartao = String(STRING_NUMERO_CARTAO, 0)
    
    
    lErro = Comando_Executar(lComando, "SELECT NumIntTitulo,NumParcela,Status,DataVencimento,DataVencimentoReal,Valor,Cobrador,CarteiraCobranca,NumTitCobrador,ProxSeqBaixa,ProxSeqOcorr,Desconto1Codigo,Desconto1Ate,Desconto1ValoR,Desconto2Codigo,Desconto2Ate,Desconto2Valor,Desconto3Codigo,Desconto3Ate,Desconto3Valor,Aceite, TipoPagto, CodConta, DataCredito, DataEmissaoCheque, BancoCheque, AgenciaCheque, ContaCorrenteCheque, NumeroCheque, DataDepositoCheque, AdmMeioPagto, Parcelamento, NumeroCartao, ValidadeCartao, AprovacaoCartao, DataTransacaoCartao FROM ParcelasRecBaixadas WHERE NumIntDoc = ?", _
    tParcelaReceber.lNumIntTitulo, tParcelaReceber.iNumParcela, tParcelaReceber.iStatus, tParcelaReceber.dtDataVencimento, tParcelaReceber.dtDataVencimentoReal, tParcelaReceber.dValor, tParcelaReceber.iCobrador, tParcelaReceber.iCarteiraCobranca, tParcelaReceber.sNumTitCobrador, tParcelaReceber.iProxSeqBaixa, tParcelaReceber.iProxSeqOcorr, tParcelaReceber.iDesconto1Codigo, _
    tParcelaReceber.dtDesconto1Ate, tParcelaReceber.dDesconto1Valor, tParcelaReceber.iDesconto2Codigo, tParcelaReceber.dtDesconto2Ate, tParcelaReceber.dDesconto2Valor, tParcelaReceber.iDesconto3Codigo, tParcelaReceber.dtDesconto3Ate, tParcelaReceber.dDesconto3Valor, tParcelaReceber.iAceite, _
    tParcelaReceber.iTipoPagto, tParcelaReceber.iCodConta, tParcelaReceber.dtDataCredito, tParcelaReceber.dtDataEmissaoCheque, tParcelaReceber.iBancoCheque, tParcelaReceber.sAgenciaCheque, tParcelaReceber.sContaCorrenteCheque, tParcelaReceber.lNumeroCheque, tParcelaReceber.dtDataDepositoCheque, tParcelaReceber.iAdmMeioPagto, tParcelaReceber.iParcelamento, tParcelaReceber.sNumeroCartao, tParcelaReceber.dtValidadeCartao, tParcelaReceber.sAprovacaoCartao, tParcelaReceber.dtDataTransacaoCartao, objParcelaReceber.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then Error 58557
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 58558
    
    If lErro = AD_SQL_SEM_DADOS Then Error 58559
    
    With tParcelaReceber

        objParcelaReceber.dDesconto1Valor = .dDesconto1Valor
        objParcelaReceber.dDesconto2Valor = .dDesconto2Valor
        objParcelaReceber.dDesconto3Valor = .dDesconto3Valor
        objParcelaReceber.dSaldo = .dSaldo
        objParcelaReceber.dtDataVencimento = .dtDataVencimento
        objParcelaReceber.dtDataVencimentoReal = .dtDataVencimentoReal
        objParcelaReceber.dtDesconto1Ate = .dtDesconto1Ate
        objParcelaReceber.dtDesconto2Ate = .dtDesconto2Ate
        objParcelaReceber.dtDesconto3Ate = .dtDesconto3Ate
        objParcelaReceber.dValor = .dValor
        objParcelaReceber.iAceite = .iAceite
        objParcelaReceber.iCarteiraCobranca = .iCarteiraCobranca
        objParcelaReceber.iCobrador = .iCobrador
        objParcelaReceber.iDescontada = .iDescontada
        objParcelaReceber.iDesconto1Codigo = .iDesconto1Codigo
        objParcelaReceber.iDesconto2Codigo = .iDesconto2Codigo
        objParcelaReceber.iDesconto3Codigo = .iDesconto3Codigo
        objParcelaReceber.iNumParcela = .iNumParcela
        objParcelaReceber.iProxSeqBaixa = .iProxSeqBaixa
        objParcelaReceber.iProxSeqOcorr = .iProxSeqOcorr
        objParcelaReceber.iStatus = .iStatus
        objParcelaReceber.lNumIntTitulo = .lNumIntTitulo
        objParcelaReceber.sNumTitCobrador = .sNumTitCobrador
        objParcelaReceber.iTipoPagto = .iTipoPagto
        objParcelaReceber.iCodConta = .iCodConta
        objParcelaReceber.dtDataCredito = .dtDataCredito
        objParcelaReceber.dtDataEmissaoCheque = .dtDataEmissaoCheque
        objParcelaReceber.iBancoCheque = .iBancoCheque
        objParcelaReceber.sAgenciaCheque = .sAgenciaCheque
        objParcelaReceber.sContaCorrenteCheque = .sContaCorrenteCheque
        objParcelaReceber.lNumeroCheque = .lNumeroCheque
        objParcelaReceber.dtDataDepositoCheque = .dtDataDepositoCheque
        objParcelaReceber.iAdmMeioPagto = .iAdmMeioPagto
        objParcelaReceber.iParcelamento = .iParcelamento
        objParcelaReceber.sNumeroCartao = .sNumeroCartao
        objParcelaReceber.dtValidadeCartao = .dtValidadeCartao
        objParcelaReceber.sAprovacaoCartao = .sAprovacaoCartao
        objParcelaReceber.dtDataTransacaoCartao = .dtDataTransacaoCartao
    
    End With
    
    Call Comando_Fechar(lComando)
    
    ParcelaReceber_Baixada_Le = SUCESSO
    
    Exit Function
    
Erro_ParcelaReceber_Baixada_Le:

    ParcelaReceber_Baixada_Le = Err
    
    Select Case Err
    
        Case 58556
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 58557, 58558
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PARCELAS_REC_BAIXADAS", Err)
            
        Case 58559
            'Tratado pela rotina chamadora
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149052)
    
    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function NFiscal_Le_Comissoes(ByVal objNFiscal As ClassNFiscal) As Long
'preenche a colecao de comissoes associadas a uma NF.

'*** Alterada em 05/03/02 por Luiz Nogueira ***
'*** ClassNFiscal teve o tipo de ColComissoesNF alterado, por isso foi necessário adicionar
'os dados lidos a um obj e depois adicionar o obj à coleção ***

'Alterado por Tulio em 28/03: Incluido campo direto/indireto

Dim lErro As Long
Dim iIndice As Integer
Dim lComando As Long
Dim lComando1 As Long
Dim tComissaoNF As typeComissaoNF
Dim objComissaoNF As ClassComissaoNF
Dim tNFiscal As typeNFiscal

On Error GoTo Erro_NFiscal_Le_Comissoes

    'Abertura comando
    lComando = 0
    lComando1 = 0

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 21379

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 21380

    lErro = Comando_Executar(lComando1, "SELECT Cliente, FilialCli, DataEmissao FROM NFiscal WHERE NumIntDoc = ?", _
    tNFiscal.lCliente, tNFiscal.iFilialCli, tNFiscal.dtDataEmissao, objNFiscal.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then Error 21381

    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 21382

    If lErro = AD_SQL_SEM_DADOS Then Error 21383

    objNFiscal.lCliente = tNFiscal.lCliente
    objNFiscal.iFilialCli = tNFiscal.iFilialCli
    objNFiscal.dtDataEmissao = tNFiscal.dtDataEmissao

    'Comentado por Luiz Nogueira, pois essa coleção não é mais "tipada"
    'Set objNFiscal.ColComissoesNF = New ColComissoesNF

    lErro = Comando_Executar(lComando, "SELECT CodVendedor, ValorBase, Percentual, Valor, PercentualEmissao, ValorEmissao, Indireta, Seq " & _
    "FROM ComissoesNF WHERE NumIntDoc = ? ORDER BY Seq", tComissaoNF.iCodVendedor, tComissaoNF.dValorBase, tComissaoNF.dPercentual, tComissaoNF.dValor, _
    tComissaoNF.dPercentualEmissao, tComissaoNF.dValorEmissao, tComissaoNF.iIndireta, tComissaoNF.iSeq, objNFiscal.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 21384

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 21385

    If lErro = AD_SQL_SEM_DADOS Then Error 21386

    Do While lErro <> AD_SQL_SEM_DADOS

        'Instancia um novo objComissaoNF
        Set objComissaoNF = New ClassComissaoNF
        
        'Transfere os dados do type para o obj
        With objComissaoNF
        
            .lNumIntDoc = tComissaoNF.lNumIntDoc
            .iCodVendedor = tComissaoNF.iCodVendedor
            .dValorBase = tComissaoNF.dValorBase
            .dPercentual = tComissaoNF.dPercentual
            .dValor = tComissaoNF.dValor
            .dPercentualEmissao = tComissaoNF.dPercentualEmissao
            .dValorEmissao = tComissaoNF.dValorEmissao
            .iIndireta = tComissaoNF.iIndireta
            .iSeq = tComissaoNF.iSeq
                    
        End With
        
        'Adiciona o obj à coleção
        objNFiscal.ColComissoesNF.Add objComissaoNF
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 21387

    Loop

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    NFiscal_Le_Comissoes = SUCESSO

    Exit Function

Erro_NFiscal_Le_Comissoes:

    NFiscal_Le_Comissoes = Err

    Select Case Err

        Case 21379, 21380
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err, Error$)

        Case 21381, 21382
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NOTA_FISCAL_NUM_SERIE", Err, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.iFilialEmpresa)

        Case 21383
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NFISCAL_NUM_SERIE_NAO_CADASTRADA", Err, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.iFilialEmpresa)

        Case 21384, 21385, 21387
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMISSOESNF", Err)

        Case 21386

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149053)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Function TipoDocInfo_Le(ByVal objTipoDocInfo As ClassTipoDocInfo) As Long
'objTipoDocInfo.sSigla tem que estar preenchido
'Preenche objTipoDocInfo a partir do bd p/registro identificado por objTipoDocInfo.sSigla
'Retorna 27263 se nao encontrar

Dim lErro As Long, lComando As Long
Dim tTipoDocInfo As typeTipoDocInfo

On Error GoTo Erro_TipoDocInfo_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 27260
    
    tTipoDocInfo.sSigla = String(STRING_TIPODOCINFO_SIGLA, 0)
    tTipoDocInfo.sDescricao = String(STRING_TIPODOCINFO_DESCRICAO, 0)
    tTipoDocInfo.sNomeReduzido = String(STRING_TIPODOCINFO_NOMEREDUZIDO, 0)
    tTipoDocInfo.sNaturezaOperacaoPadrao = String(STRING_NATUREZAOP_CODIGO, 0)
    tTipoDocInfo.sNomeTelaNFiscal = String(STRING_NOME_TELA, 0)
    tTipoDocInfo.sTipoDocCPR = String(STRING_TIPO_DOC_SIGLA, 0)
    tTipoDocInfo.sTituloTelaNFiscal = String(STRING_TIPODOCINFO_TITULO_TELA, 0)
    tTipoDocInfo.sNatOpExtPadrao = String(STRING_NATUREZAOP_CODIGO, 0)
    
    lErro = Comando_Executar(lComando, "SELECT Descricao, TipoMovtoEstoque, TipoMovtoEstoque2, TipoMovtoEstoqueBenef, TipoDocCPR, NaturezaOperacaoPadrao, InfoContabilizacao, Codigo, NomeReduzido, NomeTelaNFiscal, TituloTelaNFiscal, Faturavel, Comissao, Emitente, Destinatario, Origem, Faturamento, TipoOperacaoTrib, NatOpExtPadrao, Rastreavel, ModDocFis, ModDocFisE, FinalidadeNFe, IndConsumidorFinal, IndPresenca FROM TiposDocInfo WHERE Sigla = ?", _
        tTipoDocInfo.sDescricao, tTipoDocInfo.iTipoMovtoEstoque, tTipoDocInfo.iTipoMovtoEstoque2, tTipoDocInfo.iTipoMovtoEstoqueBenef, tTipoDocInfo.sTipoDocCPR, tTipoDocInfo.sNaturezaOperacaoPadrao, tTipoDocInfo.iInfoContabilizacao, tTipoDocInfo.iCodigo, tTipoDocInfo.sNomeReduzido, tTipoDocInfo.sNomeTelaNFiscal, tTipoDocInfo.sTituloTelaNFiscal, tTipoDocInfo.iFaturavel, tTipoDocInfo.iComissao, tTipoDocInfo.iEmitente, tTipoDocInfo.iDestinatario, tTipoDocInfo.iOrigem, tTipoDocInfo.iFaturamento, tTipoDocInfo.iTipoOperacaoTrib, tTipoDocInfo.sNatOpExtPadrao, tTipoDocInfo.iRastreavel, tTipoDocInfo.iModDocFis, tTipoDocInfo.iModDocFisE, tTipoDocInfo.iFinalidadeNFe, tTipoDocInfo.iIndConsumidorFinal, tTipoDocInfo.iIndPresenca, objTipoDocInfo.sSigla)
    If lErro <> AD_SQL_SUCESSO Then Error 27261
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 27262
        
    If lErro <> AD_SQL_SUCESSO Then Error 27263
    
    objTipoDocInfo.iCodigo = tTipoDocInfo.iCodigo
    objTipoDocInfo.iComissao = tTipoDocInfo.iComissao
    objTipoDocInfo.iFaturavel = tTipoDocInfo.iFaturavel
    objTipoDocInfo.sNomeReduzido = tTipoDocInfo.sNomeReduzido
    objTipoDocInfo.sNomeTelaNFiscal = tTipoDocInfo.sNomeTelaNFiscal
    objTipoDocInfo.sTipoDocCPR = tTipoDocInfo.sTipoDocCPR
    objTipoDocInfo.sTituloTelaNFiscal = tTipoDocInfo.sTituloTelaNFiscal
    objTipoDocInfo.iInfoContabilizacao = tTipoDocInfo.iInfoContabilizacao
    objTipoDocInfo.iTipoMovtoEstoque = tTipoDocInfo.iTipoMovtoEstoque
    objTipoDocInfo.iTipoMovtoEstoque2 = tTipoDocInfo.iTipoMovtoEstoque2
    objTipoDocInfo.iTipoMovtoEstoqueBenef = tTipoDocInfo.iTipoMovtoEstoqueBenef
    objTipoDocInfo.sDescricao = tTipoDocInfo.sDescricao
    objTipoDocInfo.sNaturezaOperacaoPadrao = tTipoDocInfo.sNaturezaOperacaoPadrao
    objTipoDocInfo.iEmitente = tTipoDocInfo.iEmitente
    objTipoDocInfo.iDestinatario = tTipoDocInfo.iDestinatario
    objTipoDocInfo.iOrigem = tTipoDocInfo.iOrigem
    objTipoDocInfo.iFaturamento = tTipoDocInfo.iFaturamento
    objTipoDocInfo.iTipoOperacaoTrib = tTipoDocInfo.iTipoOperacaoTrib
    objTipoDocInfo.sNatOpExtPadrao = tTipoDocInfo.sNatOpExtPadrao
    objTipoDocInfo.iRastreavel = tTipoDocInfo.iRastreavel
    
    'nfe 3.10
    objTipoDocInfo.iModDocFis = tTipoDocInfo.iModDocFis
    objTipoDocInfo.iModDocFisE = tTipoDocInfo.iModDocFisE
    objTipoDocInfo.iFinalidadeNFe = tTipoDocInfo.iFinalidadeNFe
    objTipoDocInfo.iIndConsumidorFinal = tTipoDocInfo.iIndConsumidorFinal
    objTipoDocInfo.iIndPresenca = tTipoDocInfo.iIndPresenca
    
    lErro = Comando_Fechar(lComando)
    
    TipoDocInfo_Le = SUCESSO

    Exit Function

Erro_TipoDocInfo_Le:

    TipoDocInfo_Le = Err

    Select Case Err

        Case 27260
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 27261, 27262
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPODOCINFO", Err, objTipoDocInfo.sSigla)
        
        Case 27263 'tipo nao encontrado
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149054)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function Estados_Le_Todos(ByVal colEstados As Collection) As Long
'Preenche colEstados com objetos da classe ClassEstado percorrendo a tabela de Estados

Dim lErro As Long, lComando As Long
Dim objEstado As ClassEstado, tEstado As typeEstado

On Error GoTo Erro_Estados_Le_Todos

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 27287
    
    tEstado.sSigla = String(STRING_ESTADO, 0)
    tEstado.sNome = String(STRING_ESTADO_NOME, 0)
    
    lErro = Comando_Executar(lComando, "SELECT Sigla, Nome, ICMSAliquotaInterna, ICMSAliquotaExportacao, ICMSAliquotaImportacao, ICMSPercFCP,ICMSAliquotaInternaAnt,ICMSAliquotaImportacaoAnt,ICMSPercFCPAnt,DataIniAliqInternaAtual,DataIniAliqImportacaoAtual,DataIniAliqFCPAtual FROM Estados", _
        tEstado.sSigla, tEstado.sNome, tEstado.dICMSAliquotaInterna, tEstado.dICMSAliquotaExportacao, tEstado.dICMSAliquotaImportacao, tEstado.dICMSPercFCP, tEstado.dICMSAliquotaInternaAnt, tEstado.dICMSAliquotaImportacaoAnt, tEstado.dICMSPercFCPAnt, tEstado.dtDataIniAliqInternaAtual, tEstado.dtDataIniAliqImportacaoAtual, tEstado.dtDataIniAliqFCPAtual)
    If lErro <> AD_SQL_SUCESSO Then Error 27288
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 27289
    
    Do While lErro <> AD_SQL_SEM_DADOS
        
        Set objEstado = New ClassEstado
        
        objEstado.dICMSAliquotaExportacao = tEstado.dICMSAliquotaExportacao
        objEstado.dICMSAliquotaImportacao = tEstado.dICMSAliquotaImportacao
        objEstado.dICMSAliquotaInterna = tEstado.dICMSAliquotaInterna
        objEstado.sNome = tEstado.sNome
        objEstado.sSigla = tEstado.sSigla
        objEstado.dICMSPercFCP = tEstado.dICMSPercFCP
        
        objEstado.dICMSAliquotaInternaAnt = tEstado.dICMSAliquotaInternaAnt
        objEstado.dICMSAliquotaImportacaoAnt = tEstado.dICMSAliquotaImportacaoAnt
        objEstado.dICMSPercFCPAnt = tEstado.dICMSPercFCPAnt
        
        objEstado.dtDataIniAliqInternaAtual = tEstado.dtDataIniAliqInternaAtual
        objEstado.dtDataIniAliqImportacaoAtual = tEstado.dtDataIniAliqImportacaoAtual
        objEstado.dtDataIniAliqFCPAtual = tEstado.dtDataIniAliqFCPAtual
        
        colEstados.Add objEstado, tEstado.sSigla
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 27290
    
    Loop
    
    lErro = Comando_Fechar(lComando)
    
    Estados_Le_Todos = SUCESSO

    Exit Function

Erro_Estados_Le_Todos:

    Estados_Le_Todos = Err

    Select Case Err

        Case 27287
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 27288, 27289, 27290
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ESTADOS", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149055)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function Estado_Le(ByVal objEstado As ClassEstado) As Long
'Lê o Estado a partir do código(sigla) em objEstado.
'Devolve os dados em objEstado.

Dim lErro As Long
Dim lComando As Long
Dim tEstado As typeEstado

On Error GoTo Erro_Estado_Le

    tEstado.sNome = String(STRING_ESTADO_NOME, 0)

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 28482

    'Pesquisa no BD, o Estado passado como parâmetro
    lErro = Comando_Executar(lComando, "SELECT Nome, ICMSAliquotaInterna, ICMSAliquotaExportacao, ICMSAliquotaImportacao, CodIBGE, ICMSPercFCP, ICMSAliquotaInternaAnt,ICMSAliquotaImportacaoAnt,ICMSPercFCPAnt,DataIniAliqInternaAtual,DataIniAliqImportacaoAtual,DataIniAliqFCPAtual FROM Estados WHERE Sigla = ?", tEstado.sNome, tEstado.dICMSAliquotaInterna, tEstado.dICMSAliquotaExportacao, tEstado.dICMSAliquotaImportacao, tEstado.iCodIBGE, tEstado.dICMSPercFCP, tEstado.dICMSAliquotaInternaAnt, tEstado.dICMSAliquotaImportacaoAnt, tEstado.dICMSPercFCPAnt, tEstado.dtDataIniAliqInternaAtual, tEstado.dtDataIniAliqImportacaoAtual, tEstado.dtDataIniAliqFCPAtual, objEstado.sSigla)
    If lErro <> AD_SQL_SUCESSO Then Error 28483

    'Verifica se o Estado existe
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 28484

    'Estado não está cadastrado
    If lErro = AD_SQL_SEM_DADOS Then Error 28485

    'Carrega os dados lidos no objEstado
    objEstado.sNome = tEstado.sNome
    objEstado.dICMSAliquotaInterna = tEstado.dICMSAliquotaInterna
    objEstado.dICMSAliquotaExportacao = tEstado.dICMSAliquotaExportacao
    objEstado.dICMSAliquotaImportacao = tEstado.dICMSAliquotaImportacao
    objEstado.iCodIBGE = tEstado.iCodIBGE
    objEstado.dICMSPercFCP = tEstado.dICMSPercFCP
    
    objEstado.dICMSAliquotaInternaAnt = tEstado.dICMSAliquotaInternaAnt
    objEstado.dICMSAliquotaImportacaoAnt = tEstado.dICMSAliquotaImportacaoAnt
    objEstado.dICMSPercFCPAnt = tEstado.dICMSPercFCPAnt
    
    objEstado.dtDataIniAliqInternaAtual = tEstado.dtDataIniAliqInternaAtual
    objEstado.dtDataIniAliqImportacaoAtual = tEstado.dtDataIniAliqImportacaoAtual
    objEstado.dtDataIniAliqFCPAtual = tEstado.dtDataIniAliqFCPAtual

    Call Comando_Fechar(lComando)

    Estado_Le = SUCESSO

Exit Function

Erro_Estado_Le:

    Estado_Le = Err

    Select Case Err

        Case 28482
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 28483, 28484
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ESTADOS1", Err, objEstado.sSigla)

        Case 28485 'Estado não cadastrado

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149056)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function CarteirasDeCobranca_Le_Todas(ByVal colCarteirasCobranca As Collection) As Long
'Lê todos as Carteiras contidas na tabela CarteirasCobranca por ordem de codigo e as coloca na coleçao colCarteirasCobranca

Dim lComando As Long
Dim lErro As Long
Dim objCarteiraCobranca As ClassCarteiraCobranca
Dim iCodigo As Integer, iValidaPara As Integer
Dim sDescricao As String

On Error GoTo Erro_CarteirasDeCobranca_Le_Todas

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 23403

    sDescricao = String(STRING_DESCRICAO_CARTCOBR, 0)

    'Lê a tabela CarteirasCobranca
    lErro = Comando_Executar(lComando, "SELECT Codigo, Descricao,ValidaPara FROM CarteirasCobranca ORDER BY Codigo", iCodigo, sDescricao, iValidaPara)
    If lErro <> AD_SQL_SUCESSO Then Error 23404

    'Lê o primeira Carteira de Cobrança
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 23405

    'Enquanto a tabela tiver dados
    Do While lErro <> AD_SQL_SEM_DADOS

        Set objCarteiraCobranca = New ClassCarteiraCobranca

        objCarteiraCobranca.iCodigo = iCodigo
        objCarteiraCobranca.sDescricao = sDescricao
        objCarteiraCobranca.iValidaPara = iValidaPara

        'Coloca a carteira lida na coleção
        colCarteirasCobranca.Add objCarteiraCobranca

        'Lê a próxima carteira
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 23406

    Loop

    Call Comando_Fechar(lComando)

    CarteirasDeCobranca_Le_Todas = SUCESSO

    Exit Function

Erro_CarteirasDeCobranca_Le_Todas:

    CarteirasDeCobranca_Le_Todas = Err

    Select Case Err

        Case 23403
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 23404, 23405, 23406
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA_CARTEIRASCOBRANCA", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149057)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function CarteiraDeCobranca_Le(ByVal objCarteiraCobranca As ClassCarteiraCobranca) As Long
'Lê a Carteira a partir da chave fornecida em objCarteiraCobranca (Código) e retorna os dados em objCarteiraCobranca

Dim lComando As Long
Dim lErro As Long
Dim sDescricao As String
Dim iValidaPara As Integer

On Error GoTo Erro_CarteiraDeCobranca_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 23410

    sDescricao = String(STRING_DESCRICAO_CARTCOBR, 0)

    'Seleciona as carteiras cujo Código está no Obj e coloca a descrição no sDescricao e ValidaPara no iValidaPara
    lErro = Comando_Executar(lComando, "SELECT Descricao,ValidaPara FROM CarteirasCobranca WHERE Codigo = ?", sDescricao, iValidaPara, objCarteiraCobranca.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 23411

    'Posiciona na primeira Carteira das que foram lidas
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 23412

    'Se não encontrou nenhuma Carteira com o Código em questão
    If lErro = AD_SQL_SEM_DADOS Then Error 23413

    'Passa a descrição e Validação da Carteira lida para sDescricao e iValidaPara
    objCarteiraCobranca.sDescricao = sDescricao
    objCarteiraCobranca.iValidaPara = iValidaPara

    Call Comando_Fechar(lComando)

    CarteiraDeCobranca_Le = SUCESSO

    Exit Function

Erro_CarteiraDeCobranca_Le:

    CarteiraDeCobranca_Le = Err

    Select Case Err

        Case 23410
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 23411, 23412
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA_CARTEIRASCOBRANCA1", Err, objCarteiraCobranca.iCodigo)

        Case 23413

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149058)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Cobrador_Le_Carteiras(ByVal objCobrador As ClassCobrador, colCarteirasCobrador As Collection) As Long
'Lê as Carteiras a partir da chave fornecida em iCobrador
'Retorna as Carteiras do Cobrador lidas na coleção (com o codigo como chave dos elementos da colecao)

Dim objCarteiraCobrador As ClassCarteiraCobrador
Dim lComando As Long
Dim lErro As Long
Dim tCarteira As typeCarteiraCobrador

On Error GoTo Erro_Cobrador_Le_Carteiras

    Set colCarteirasCobrador = New Collection

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 23497

    tCarteira.sContaContabil = String(STRING_CONTA, 0)
    tCarteira.sContaDuplDescontadas = String(STRING_CONTA, 0)
    tCarteira.sFaixaNossoNumeroInicial = String(STRING_FAIXA_NUM_COB, 0)
    tCarteira.sFaixaNossoNumeroFinal = String(STRING_FAIXA_NUM_COB, 0)
    tCarteira.sFaixaNossoNumeroProx = String(STRING_FAIXA_NUM_COB, 0)
    tCarteira.sNomeNoBanco = String(STRING_NOMENOBANCO, 0)
    tCarteira.sCodCarteiraNoBanco = String(STRING_CODCARTNOBANCO, 0)

    'Seleciona as carteiras cujo Código está no Obj e coloca os dados no type
    lErro = Comando_Executar(lComando, "SELECT CodCarteiraCobranca, Desativada, ContaContabil, ContaDuplDescontadas, DiasDeRetencao, TaxaCobranca , TaxaDesconto, FaixaNossoNumeroInicial, FaixaNossoNumeroFinal, FaixaNossoNumeroProx, NomeNoBanco, CodCarteiraNoBanco, NumCarteiraNoBanco, QuantidadeAtual , QuantidadeAtualBanco, SaldoAtual, SaldoAtualBanco,ImprimeBoleta,ComRegistro,GeraNossoNumero, FormPreImp FROM CarteirasCobrador WHERE Cobrador = ?", _
    tCarteira.iCodCarteiraCobranca, tCarteira.iDesativada, tCarteira.sContaContabil, tCarteira.sContaDuplDescontadas, tCarteira.iDiasDeRetencao, tCarteira.dTaxaCobranca, tCarteira.dTaxaDesconto, tCarteira.sFaixaNossoNumeroInicial, tCarteira.sFaixaNossoNumeroFinal, tCarteira.sFaixaNossoNumeroProx, tCarteira.sNomeNoBanco, tCarteira.sCodCarteiraNoBanco, tCarteira.iNumCarteiraNoBanco, tCarteira.lQuantidadeAtual, tCarteira.lQuantidadeAtualBanco, tCarteira.dSaldoAtual, tCarteira.dSaldoAtualBanco, tCarteira.iImprimeBoleta, tCarteira.iComRegistro, tCarteira.iGeraNossoNumero, tCarteira.iFormPreImp, objCobrador.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 23498

    'Posiciona na primeira Carteira das que foram lidas
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 23499

    'Se não encontrou nenhuma Carteira com o Código em questão
    If lErro = AD_SQL_SEM_DADOS Then Error 23500

    Do While lErro <> AD_SQL_SEM_DADOS
        
        Set objCarteiraCobrador = New ClassCarteiraCobrador
        
        'Move carteira Cobrador para o objeto
        Call CarteiraCobrador_MoverCamposObj(objCobrador.iCodigo, objCarteiraCobrador, tCarteira)
               
        'Coloca objCarteiraCobrador na Coleção
        colCarteirasCobrador.Add objCarteiraCobrador, CStr(objCarteiraCobrador.iCodCarteiraCobranca)

        'Buscar próxima carteira
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 23499

    Loop

    Call Comando_Fechar(lComando)

    Cobrador_Le_Carteiras = SUCESSO

    Exit Function

Erro_Cobrador_Le_Carteiras:

    Cobrador_Le_Carteiras = Err

    Select Case Err

        Case 23497
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 23498, 23499
            'lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA_CARTEIRASCOBRADOR", Err, objCarteiraCobrador.iCobrador)

        Case 23500

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149059)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Private Sub CarteiraCobrador_MoverCamposObj(ByVal iCodigo As Integer, objCarteiraCobrador As ClassCarteiraCobrador, tCarteira As typeCarteiraCobrador)

    'Coloca os dados da Carteira em objCarteiraCobrador
    With objCarteiraCobrador
        .iCobrador = iCodigo
        .iCodCarteiraCobranca = tCarteira.iCodCarteiraCobranca
        .iDesativada = tCarteira.iDesativada
        .sContaContabil = tCarteira.sContaContabil
        .sContaDuplDescontadas = tCarteira.sContaDuplDescontadas
        .iDiasDeRetencao = tCarteira.iDiasDeRetencao
        .dTaxaCobranca = tCarteira.dTaxaCobranca
        .dTaxaDesconto = tCarteira.dTaxaDesconto
        .sFaixaNossoNumeroInicial = tCarteira.sFaixaNossoNumeroInicial
        .sFaixaNossoNumeroFinal = tCarteira.sFaixaNossoNumeroFinal
        .sFaixaNossoNumeroProx = tCarteira.sFaixaNossoNumeroProx
        .sNomeNoBanco = tCarteira.sNomeNoBanco
        .sCodCarteiraNoBanco = tCarteira.sCodCarteiraNoBanco
        .iNumCarteiraNoBanco = tCarteira.iNumCarteiraNoBanco
        .lQuantidadeAtual = tCarteira.lQuantidadeAtual
        .lQuantidadeAtualBanco = tCarteira.lQuantidadeAtualBanco
        .dSaldoAtual = tCarteira.dSaldoAtual
        .dSaldoAtualBanco = tCarteira.dSaldoAtualBanco
        .iImprimeBoleta = tCarteira.iImprimeBoleta
        .iComRegistro = tCarteira.iComRegistro
        .iGeraNossoNumero = tCarteira.iGeraNossoNumero
        .iFormPreImp = tCarteira.iFormPreImp
    End With

End Sub

Function FilialCliente_Le_Categorias(ByVal objFilialCliente As ClassFilialCliente, ByVal colCategoriaItem As Collection) As Long
'Lê na tabela de FilialClienteCategorias, todas as categorias da Filial de Cliente (com o respectivo valor em cada categoria), guardando na coleção colCategoriaItem

Dim lErro As Long
Dim iIndice As Integer
Dim lComando As Long
Dim sCategoria As String
Dim sItem As String
Dim objFilialCliCategoria As ClassFilialCliCategoria

On Error GoTo Erro_FilialCliente_Le_Categorias

    'Inicializar comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 33001

    sCategoria = String(STRING_CATEGORIACLIENTE_CATEGORIA, 0)
    sItem = String(STRING_CATEGORIACLIENTEITEM_ITEM, 0)

    'Pesquisa no BD as categoria da filial do cliente
    lErro = Comando_Executar(lComando, "SELECT Categoria, Item FROM FilialClienteCategorias WHERE Cliente = ? AND Filial = ?", sCategoria, sItem, objFilialCliente.lCodCliente, objFilialCliente.iCodFilial)
    If lErro <> AD_SQL_SUCESSO Then Error 33002

    'Verifica se alguma categoria existe
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 33003

    'Passa para a collection
    Do While lErro <> AD_SQL_SEM_DADOS

        Set objFilialCliCategoria = New ClassFilialCliCategoria

        objFilialCliCategoria.lCliente = objFilialCliente.lCodCliente
        objFilialCliCategoria.iFilial = objFilialCliente.iCodFilial
        objFilialCliCategoria.sCategoria = sCategoria
        objFilialCliCategoria.sItem = sItem

        colCategoriaItem.Add objFilialCliCategoria

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 33004

    Loop

    'Fechamento comando
    Call Comando_Fechar(lComando)

    FilialCliente_Le_Categorias = SUCESSO

    Exit Function

Erro_FilialCliente_Le_Categorias:

    FilialCliente_Le_Categorias = Err

    Select Case Err

        Case 33001
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 33002, 33003, 33004
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIALCLIENTECATEGORIAS2", Err, objFilialCliCategoria.lCliente)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149060)

    End Select

    'Fechamento comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Private Sub PadroesCobranca_MoverCamposObj(ByVal objPadraoCobranca As ClassPadraoCobranca, tPadraoCobranca As typePadraoCobranca)

    objPadraoCobranca.iCodigo = tPadraoCobranca.iCodigo
    objPadraoCobranca.iInativo = tPadraoCobranca.iInativo
    objPadraoCobranca.sDescricao = tPadraoCobranca.sDescricao
    objPadraoCobranca.iInstrucao1 = tPadraoCobranca.iInstrucao1
    objPadraoCobranca.iInstrucao2 = tPadraoCobranca.iInstrucao2
    objPadraoCobranca.dJuros = tPadraoCobranca.dJuros
    objPadraoCobranca.iDiasDeProtesto1 = tPadraoCobranca.iDiasDeProtesto1
    objPadraoCobranca.iDiasDeProtesto2 = tPadraoCobranca.iDiasDeProtesto2

End Sub

Function PadroesCobranca_Le_Todos(ByVal colPadroesCobranca As Collection) As Long
'carrega uma colecao com elementos objPadraoCobranca correspondentes
'aos registros ATIVOS E INATIVOS da tabela PadroesCobranca

Dim lComando As Long
Dim lErro As Long, objPadraoCobranca As ClassPadraoCobranca
Dim tPadraoCobranca As typePadraoCobranca

On Error GoTo Erro_PadroesCobranca_Le_Todos

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 7558

    tPadraoCobranca.sDescricao = String(STRING_PADRAO_COBRANCA_DESCRICAO, 0)

    lErro = Comando_Executar(lComando, "SELECT Codigo, Descricao, Instrucao1, Instrucao2, Juros, DiasDeProtesto1, DiasDeProtesto2, Inativo FROM PadroesCobranca", tPadraoCobranca.iCodigo, tPadraoCobranca.sDescricao, tPadraoCobranca.iInstrucao1, tPadraoCobranca.iInstrucao2, tPadraoCobranca.dJuros, tPadraoCobranca.iDiasDeProtesto1, tPadraoCobranca.iDiasDeProtesto2, tPadraoCobranca.iInativo)
    If lErro <> AD_SQL_SUCESSO Then Error 7559

    'Lê o primeiro registro
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 19319
    
    'Enquanto a tabela tiver dados
    Do While lErro <> AD_SQL_SEM_DADOS

        Set objPadraoCobranca = New ClassPadraoCobranca

        Call PadroesCobranca_MoverCamposObj(objPadraoCobranca, tPadraoCobranca)

        colPadroesCobranca.Add objPadraoCobranca

        'Lê o próximo registro
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 7559

    Loop

    lErro = Comando_Fechar(lComando)

    PadroesCobranca_Le_Todos = SUCESSO

Exit Function

Erro_PadroesCobranca_Le_Todos:

    PadroesCobranca_Le_Todos = Err

    Select Case Err

        Case 7558
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 7559, 19319
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PADRAO_COBRANCA2", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149061)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function


End Function

Function PadraoCobranca_Le(ByVal objPadraoCobranca As ClassPadraoCobranca) As Long
'le o PadraoCobranca a partir da chave fornecida em objPadraoCobranca. Devolve os dados em objPadraoCobranca

Dim lComando As Long
Dim lErro As Long
Dim tPadraoCobranca As typePadraoCobranca

On Error GoTo Erro_PadraoCobranca_Le

    tPadraoCobranca.sDescricao = String(STRING_PADRAO_COBRANCA_DESCRICAO, 0)

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 19295

     'Verifica se o PadraoCobranca existe, e se existir carrega seus dados em objPadraoCobranca
    lErro = Comando_Executar(lComando, "SELECT  Descricao, Instrucao1, Instrucao2, Juros, DiasDeProtesto1, DiasDeProtesto2, Inativo FROM PadroesCobranca WHERE Codigo = ?", tPadraoCobranca.sDescricao, tPadraoCobranca.iInstrucao1, tPadraoCobranca.iInstrucao2, tPadraoCobranca.dJuros, tPadraoCobranca.iDiasDeProtesto1, tPadraoCobranca.iDiasDeProtesto2, tPadraoCobranca.iInativo, objPadraoCobranca.iCodigo)

    If lErro <> AD_SQL_SUCESSO Then Error 19296

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 19297

    'PadraoCobranca não está cadastrado
    If lErro = AD_SQL_SEM_DADOS Then Error 19298

    objPadraoCobranca.sDescricao = tPadraoCobranca.sDescricao
    objPadraoCobranca.iInstrucao1 = tPadraoCobranca.iInstrucao1
    objPadraoCobranca.iInstrucao2 = tPadraoCobranca.iInstrucao2
    objPadraoCobranca.dJuros = tPadraoCobranca.dJuros
    objPadraoCobranca.iDiasDeProtesto1 = tPadraoCobranca.iDiasDeProtesto1
    objPadraoCobranca.iDiasDeProtesto2 = tPadraoCobranca.iDiasDeProtesto2
    objPadraoCobranca.iInativo = tPadraoCobranca.iInativo
    
    lErro = Comando_Fechar(lComando)

    PadraoCobranca_Le = SUCESSO

Exit Function

Erro_PadraoCobranca_Le:

    PadraoCobranca_Le = Err

    Select Case Err

        Case 19295
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 19296, 19297
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PADRAO_COBRANCA", Err, objPadraoCobranca.iCodigo)

        Case 19298 'PadraoCobranca nao cadastrado

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149062)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function OcorrRemParcRec_Le(ByVal objOcorrRemParcRec As ClassOcorrRemParcRec) As Long
'Traz do bd as informações para preencher um objeto da classe ClassOcorrRemParcRec à partir
'da chave fornecida em objOcorrRemParcRec.lNumIntParc, objOcorrRemParcRec.iNumSeqOcorr e objOcorrRemParcRec.iFilialEmpresa que devem estar preenchidos

Dim lErro As Long
Dim lComando As Long
Dim tTipoOcorrRemParcRec As typeTipoOcorrRemParcRec

On Error GoTo Erro_OcorrRemParcRec_Le

    'Inicializa comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 28532

    tTipoOcorrRemParcRec.sNumTitCobrador = String(STRING_NUMTITCOBRADOR, 0)
    
    lErro = Comando_Executar(lComando, "SELECT FilialEmpresa, Cobrador, CodOcorrencia, DataRegistro, Data, TituloVoltaCarteira, NovaDataVcto, Juros, Instrucao1, DiasDeProtesto1, Instrucao2, DiasDeProtesto2, ValorCobrado, NumBordero, NumTitCobrador FROM OcorrenciasRemParcRec WHERE NumIntParc = ? AND NumSeqOcorr = ? AND FilialEmpresa = ?", _
         tTipoOcorrRemParcRec.iFilialEmpresa, tTipoOcorrRemParcRec.iCobrador, tTipoOcorrRemParcRec.iCodOcorrencia, tTipoOcorrRemParcRec.dtDataRegistro, tTipoOcorrRemParcRec.dtData, tTipoOcorrRemParcRec.iTituloVoltaCarteira, tTipoOcorrRemParcRec.dtNovaDataVcto, tTipoOcorrRemParcRec.dJuros, tTipoOcorrRemParcRec.iInstrucao1, tTipoOcorrRemParcRec.iDiasDeProtesto1, tTipoOcorrRemParcRec.iInstrucao2, tTipoOcorrRemParcRec.iDiasDeProtesto2, tTipoOcorrRemParcRec.dValorCobrado, tTipoOcorrRemParcRec.lNumBordero, tTipoOcorrRemParcRec.sNumTitCobrador, _
        objOcorrRemParcRec.lNumIntParc, objOcorrRemParcRec.iNumSeqOcorr, objOcorrRemParcRec.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then Error 28533

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 28534

    If lErro = AD_SQL_SEM_DADOS Then Error 28535

    objOcorrRemParcRec.iFilialEmpresa = tTipoOcorrRemParcRec.iFilialEmpresa
    objOcorrRemParcRec.iCobrador = tTipoOcorrRemParcRec.iCobrador
    objOcorrRemParcRec.iCodOcorrencia = tTipoOcorrRemParcRec.iCodOcorrencia
    objOcorrRemParcRec.dtDataRegistro = tTipoOcorrRemParcRec.dtDataRegistro
    objOcorrRemParcRec.dtData = tTipoOcorrRemParcRec.dtData
    objOcorrRemParcRec.iTituloVoltaCarteira = tTipoOcorrRemParcRec.iTituloVoltaCarteira
    objOcorrRemParcRec.dtNovaDataVcto = tTipoOcorrRemParcRec.dtNovaDataVcto
    objOcorrRemParcRec.dJuros = tTipoOcorrRemParcRec.dJuros
    objOcorrRemParcRec.iInstrucao1 = tTipoOcorrRemParcRec.iInstrucao1
    objOcorrRemParcRec.iDiasDeProtesto1 = tTipoOcorrRemParcRec.iDiasDeProtesto1
    objOcorrRemParcRec.iInstrucao2 = tTipoOcorrRemParcRec.iInstrucao2
    objOcorrRemParcRec.iDiasDeProtesto2 = tTipoOcorrRemParcRec.iDiasDeProtesto2
    objOcorrRemParcRec.dValorCobrado = tTipoOcorrRemParcRec.dValorCobrado
    objOcorrRemParcRec.lNumBordero = tTipoOcorrRemParcRec.lNumBordero
    objOcorrRemParcRec.sNumTitCobrador = tTipoOcorrRemParcRec.sNumTitCobrador
    
    'Libera comando
    Call Comando_Fechar(lComando)

    OcorrRemParcRec_Le = SUCESSO

    Exit Function

Erro_OcorrRemParcRec_Le:

    OcorrRemParcRec_Le = Err

    Select Case Err

        Case 28532
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 28533, 28534
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_OCORRENCIASREMPARCREC", Err, objOcorrRemParcRec.lNumIntParc, objOcorrRemParcRec.iNumSeqOcorr)

        Case 28535
            'Não encontrou. Erro tratado na rotina chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149063)

    End Select

    'Libera comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Private Sub ParcelaRec_MoveCamposObj(ByVal objParcelaReceber As ClassParcelaReceber, tParcelaReceber As typeParcelaReceber)
'transfere campos do type para o obj

    'Preenche objParcelaReceber com os dados lidos
    objParcelaReceber.lNumIntTitulo = tParcelaReceber.lNumIntTitulo
    objParcelaReceber.iNumParcela = tParcelaReceber.iNumParcela
    objParcelaReceber.lNumIntDoc = tParcelaReceber.lNumIntDoc
    objParcelaReceber.iStatus = tParcelaReceber.iStatus
    objParcelaReceber.dtDataVencimento = tParcelaReceber.dtDataVencimento
    objParcelaReceber.dtDataVencimentoReal = tParcelaReceber.dtDataVencimentoReal
    objParcelaReceber.dValor = tParcelaReceber.dValor
    objParcelaReceber.iCobrador = tParcelaReceber.iCobrador
    objParcelaReceber.iCarteiraCobranca = tParcelaReceber.iCarteiraCobranca
    objParcelaReceber.sNumTitCobrador = tParcelaReceber.sNumTitCobrador
    objParcelaReceber.iProxSeqBaixa = tParcelaReceber.iProxSeqBaixa
    objParcelaReceber.iProxSeqOcorr = tParcelaReceber.iProxSeqOcorr
    objParcelaReceber.iDesconto1Codigo = tParcelaReceber.iDesconto1Codigo
    objParcelaReceber.dtDesconto1Ate = tParcelaReceber.dtDesconto1Ate
    objParcelaReceber.dDesconto1Valor = tParcelaReceber.dDesconto1Valor
    objParcelaReceber.iDesconto2Codigo = tParcelaReceber.iDesconto2Codigo
    objParcelaReceber.dtDesconto2Ate = tParcelaReceber.dtDesconto2Ate
    objParcelaReceber.dDesconto2Valor = tParcelaReceber.dDesconto2Valor
    objParcelaReceber.iDesconto3Codigo = tParcelaReceber.iDesconto3Codigo
    objParcelaReceber.dtDesconto3Ate = tParcelaReceber.dtDesconto3Ate
    objParcelaReceber.dDesconto3Valor = tParcelaReceber.dDesconto3Valor
    objParcelaReceber.iAceite = tParcelaReceber.iAceite
    objParcelaReceber.iTipoPagto = tParcelaReceber.iTipoPagto
    objParcelaReceber.iCodConta = tParcelaReceber.iCodConta
    objParcelaReceber.dtDataCredito = tParcelaReceber.dtDataCredito
    objParcelaReceber.dtDataEmissaoCheque = tParcelaReceber.dtDataEmissaoCheque
    objParcelaReceber.iBancoCheque = tParcelaReceber.iBancoCheque
    objParcelaReceber.sAgenciaCheque = tParcelaReceber.sAgenciaCheque
    objParcelaReceber.sContaCorrenteCheque = tParcelaReceber.sContaCorrenteCheque
    objParcelaReceber.lNumeroCheque = tParcelaReceber.lNumeroCheque
    objParcelaReceber.dtDataDepositoCheque = tParcelaReceber.dtDataDepositoCheque
    objParcelaReceber.iAdmMeioPagto = tParcelaReceber.iAdmMeioPagto
    objParcelaReceber.iParcelamento = tParcelaReceber.iParcelamento
    objParcelaReceber.sNumeroCartao = tParcelaReceber.sNumeroCartao
    objParcelaReceber.dtValidadeCartao = tParcelaReceber.dtValidadeCartao
    objParcelaReceber.sAprovacaoCartao = tParcelaReceber.sAprovacaoCartao
    objParcelaReceber.dtDataTransacaoCartao = tParcelaReceber.dtDataTransacaoCartao

End Sub

Function ParcelaReceberBaixada_Le_SemNumIntDoc(ByVal objParcelaReceber As ClassParcelaReceber) As Long
'Lê Número Interno da Parcela a partir de NumIntTitulo e NumParcela

Dim lErro As Long
Dim lComando As Long
Dim tParcelaReceber As typeParcelaReceber

On Error GoTo Erro_ParcelaReceberBaixada_Le_SemNumIntDoc

    'Inicializa comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 28562

    tParcelaReceber.sNumTitCobrador = String(STRING_NOSSO_NUMERO, 0)

    tParcelaReceber.sAgenciaCheque = String(STRING_AGENCIA, 0)
    tParcelaReceber.sAprovacaoCartao = String(STRING_APROVACAO_CARTAO, 0)
    tParcelaReceber.sContaCorrenteCheque = String(STRING_CONTA_CORRENTE, 0)
    tParcelaReceber.sNumeroCartao = String(STRING_NUMERO_CARTAO, 0)

    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Status, DataVencimento, DataVencimentoReal, Valor, Cobrador, CarteiraCobranca, NumTitCobrador, ProxSeqBaixa, ProxSeqOcorr, Desconto1Codigo, Desconto1Ate, Desconto1Valor, Desconto2Codigo, Desconto2Ate, Desconto2Valor, Desconto3Codigo, Desconto3Ate, Desconto3Valor,  Aceite, TipoPagto, CodConta, DataCredito, DataEmissaoCheque, BancoCheque, AgenciaCheque, ContaCorrenteCheque, NumeroCheque, DataDepositoCheque, AdmMeioPagto, Parcelamento, NumeroCartao, ValidadeCartao, AprovacaoCartao, DataTransacaoCartao FROM ParcelasRecBaixadas WHERE NumIntTitulo = ? AND NumParcela = ? AND Status <> ?", _
        tParcelaReceber.lNumIntDoc, tParcelaReceber.iStatus, tParcelaReceber.dtDataVencimento, tParcelaReceber.dtDataVencimentoReal, tParcelaReceber.dValor, tParcelaReceber.iCobrador, tParcelaReceber.iCarteiraCobranca, tParcelaReceber.sNumTitCobrador, tParcelaReceber.iProxSeqBaixa, tParcelaReceber.iProxSeqOcorr, tParcelaReceber.iDesconto1Codigo, tParcelaReceber.dtDesconto1Ate, tParcelaReceber.dDesconto1Valor, tParcelaReceber.iDesconto2Codigo, tParcelaReceber.dtDesconto2Ate, tParcelaReceber.dDesconto2Valor, tParcelaReceber.iDesconto3Codigo, tParcelaReceber.dtDesconto3Ate, tParcelaReceber.dDesconto3Valor, tParcelaReceber.iAceite, _
        tParcelaReceber.iTipoPagto, tParcelaReceber.iCodConta, tParcelaReceber.dtDataCredito, tParcelaReceber.dtDataEmissaoCheque, tParcelaReceber.iBancoCheque, tParcelaReceber.sAgenciaCheque, tParcelaReceber.sContaCorrenteCheque, tParcelaReceber.lNumeroCheque, tParcelaReceber.dtDataDepositoCheque, tParcelaReceber.iAdmMeioPagto, tParcelaReceber.iParcelamento, tParcelaReceber.sNumeroCartao, tParcelaReceber.dtValidadeCartao, tParcelaReceber.sAprovacaoCartao, tParcelaReceber.dtDataTransacaoCartao, objParcelaReceber.lNumIntTitulo, objParcelaReceber.iNumParcela, STATUS_EXCLUIDO)
    If lErro <> AD_SQL_SUCESSO Then Error 28564

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 28565

    'Se não encontrou a Parcela Receber Baixada --> Erro
    If lErro = AD_SQL_SEM_DADOS Then Error 28567

    tParcelaReceber.lNumIntTitulo = objParcelaReceber.lNumIntTitulo
    tParcelaReceber.iNumParcela = objParcelaReceber.iNumParcela
    Call ParcelaRec_MoveCamposObj(objParcelaReceber, tParcelaReceber)
    
    'Libera comando
    Call Comando_Fechar(lComando)

    ParcelaReceberBaixada_Le_SemNumIntDoc = SUCESSO

    Exit Function

Erro_ParcelaReceberBaixada_Le_SemNumIntDoc:

    ParcelaReceberBaixada_Le_SemNumIntDoc = Err

    Select Case Err

        Case 28562
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 28564, 28565
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PARCELASREC_NUMINT", Err, objParcelaReceber.lNumIntTitulo, objParcelaReceber.iNumParcela)

        Case 28567
            'Não encontrou. Erro tratado na rotina chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149064)

    End Select

    'Libera comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ParcelaReceber_Le_SemNumIntDoc(ByVal objParcelaReceber As ClassParcelaReceber) As Long
'Lê Número Interno da Parcela a partir de NumIntTitulo e NumParcela

Dim lErro As Long
Dim lComando As Long
Dim tParcelaReceber As typeParcelaReceber

On Error GoTo Erro_ParcelaReceber_Le_SemNumIntDoc

    'Inicializa comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 28587

    tParcelaReceber.sNumTitCobrador = String(STRING_NOSSO_NUMERO, 0)
    '###############################################
    'INSERIDO POR WAGNER
    tParcelaReceber.sObservacao = String(STRING_MAX_NOME_ARQUIVO, 0)
    '###############################################

    tParcelaReceber.sAgenciaCheque = String(STRING_AGENCIA, 0)
    tParcelaReceber.sAprovacaoCartao = String(STRING_APROVACAO_CARTAO, 0)
    tParcelaReceber.sContaCorrenteCheque = String(STRING_CONTA_CORRENTE, 0)
    tParcelaReceber.sNumeroCartao = String(STRING_NUMERO_CARTAO, 0)

    '#################################################
    'ALTERADO POR WAGNER
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Status, DataVencimento, DataVencimentoReal, Saldo, Valor, Cobrador, CarteiraCobranca, NumTitCobrador, ProxSeqBaixa, ProxSeqOcorr, Desconto1Codigo, Desconto1Ate, Desconto1Valor, Desconto2Codigo, Desconto2Ate, Desconto2Valor, Desconto3Codigo, Desconto3Ate, Desconto3Valor,  Aceite, Descontada, Previsao, Observacao, TipoPagto, CodConta, DataCredito, DataEmissaoCheque, BancoCheque, AgenciaCheque, ContaCorrenteCheque, NumeroCheque, DataDepositoCheque, AdmMeioPagto, Parcelamento, NumeroCartao, ValidadeCartao, AprovacaoCartao, DataTransacaoCartao FROM ParcelasRec WHERE NumIntTitulo = ? AND NumParcela = ? AND Status <> ?", _
        tParcelaReceber.lNumIntDoc, tParcelaReceber.iStatus, tParcelaReceber.dtDataVencimento, tParcelaReceber.dtDataVencimentoReal, tParcelaReceber.dSaldo, tParcelaReceber.dValor, tParcelaReceber.iCobrador, tParcelaReceber.iCarteiraCobranca, tParcelaReceber.sNumTitCobrador, tParcelaReceber.iProxSeqBaixa, tParcelaReceber.iProxSeqOcorr, tParcelaReceber.iDesconto1Codigo, tParcelaReceber.dtDesconto1Ate, tParcelaReceber.dDesconto1Valor, tParcelaReceber.iDesconto2Codigo, tParcelaReceber.dtDesconto2Ate, tParcelaReceber.dDesconto2Valor, tParcelaReceber.iDesconto3Codigo, tParcelaReceber.dtDesconto3Ate, tParcelaReceber.dDesconto3Valor, tParcelaReceber.iAceite, tParcelaReceber.iDescontada, tParcelaReceber.iPrevisao, tParcelaReceber.sObservacao, _
        tParcelaReceber.iTipoPagto, tParcelaReceber.iCodConta, tParcelaReceber.dtDataCredito, tParcelaReceber.dtDataEmissaoCheque, tParcelaReceber.iBancoCheque, tParcelaReceber.sAgenciaCheque, tParcelaReceber.sContaCorrenteCheque, tParcelaReceber.lNumeroCheque, tParcelaReceber.dtDataDepositoCheque, tParcelaReceber.iAdmMeioPagto, tParcelaReceber.iParcelamento, tParcelaReceber.sNumeroCartao, tParcelaReceber.dtValidadeCartao, tParcelaReceber.sAprovacaoCartao, tParcelaReceber.dtDataTransacaoCartao, objParcelaReceber.lNumIntTitulo, objParcelaReceber.iNumParcela, STATUS_EXCLUIDO)
    If lErro <> AD_SQL_SUCESSO Then Error 28588
    '#################################################

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 28589

    If lErro = AD_SQL_SEM_DADOS Then Error 28590

    'Preenche objParcelaReceber com os dados lidos
    objParcelaReceber.lNumIntDoc = tParcelaReceber.lNumIntDoc
    objParcelaReceber.iStatus = tParcelaReceber.iStatus
    objParcelaReceber.dtDataVencimento = tParcelaReceber.dtDataVencimento
    objParcelaReceber.dtDataVencimentoReal = tParcelaReceber.dtDataVencimentoReal
    objParcelaReceber.dSaldo = tParcelaReceber.dSaldo
    objParcelaReceber.dValor = tParcelaReceber.dValor
    objParcelaReceber.iCobrador = tParcelaReceber.iCobrador
    objParcelaReceber.iCarteiraCobranca = tParcelaReceber.iCarteiraCobranca
    objParcelaReceber.sNumTitCobrador = tParcelaReceber.sNumTitCobrador
    objParcelaReceber.iProxSeqBaixa = tParcelaReceber.iProxSeqBaixa
    objParcelaReceber.iProxSeqOcorr = tParcelaReceber.iProxSeqOcorr
    objParcelaReceber.iDesconto1Codigo = tParcelaReceber.iDesconto1Codigo
    objParcelaReceber.dtDesconto1Ate = tParcelaReceber.dtDesconto1Ate
    objParcelaReceber.dDesconto1Valor = tParcelaReceber.dDesconto1Valor
    objParcelaReceber.iDesconto2Codigo = tParcelaReceber.iDesconto2Codigo
    objParcelaReceber.dtDesconto2Ate = tParcelaReceber.dtDesconto2Ate
    objParcelaReceber.dDesconto2Valor = tParcelaReceber.dDesconto2Valor
    objParcelaReceber.iDesconto3Codigo = tParcelaReceber.iDesconto3Codigo
    objParcelaReceber.dtDesconto3Ate = tParcelaReceber.dtDesconto3Ate
    objParcelaReceber.dDesconto3Valor = tParcelaReceber.dDesconto3Valor
    objParcelaReceber.iAceite = tParcelaReceber.iAceite
    objParcelaReceber.iDescontada = tParcelaReceber.iDescontada

    '#################################################
    'ALTERADO POR WAGNER
    objParcelaReceber.iPrevisao = tParcelaReceber.iPrevisao
    objParcelaReceber.sObservacao = tParcelaReceber.sObservacao
    '#################################################

    objParcelaReceber.iTipoPagto = tParcelaReceber.iTipoPagto
    objParcelaReceber.iCodConta = tParcelaReceber.iCodConta
    objParcelaReceber.dtDataCredito = tParcelaReceber.dtDataCredito
    objParcelaReceber.dtDataEmissaoCheque = tParcelaReceber.dtDataEmissaoCheque
    objParcelaReceber.iBancoCheque = tParcelaReceber.iBancoCheque
    objParcelaReceber.sAgenciaCheque = tParcelaReceber.sAgenciaCheque
    objParcelaReceber.sContaCorrenteCheque = tParcelaReceber.sContaCorrenteCheque
    objParcelaReceber.lNumeroCheque = tParcelaReceber.lNumeroCheque
    objParcelaReceber.dtDataDepositoCheque = tParcelaReceber.dtDataDepositoCheque
    objParcelaReceber.iAdmMeioPagto = tParcelaReceber.iAdmMeioPagto
    objParcelaReceber.iParcelamento = tParcelaReceber.iParcelamento
    objParcelaReceber.sNumeroCartao = tParcelaReceber.sNumeroCartao
    objParcelaReceber.dtValidadeCartao = tParcelaReceber.dtValidadeCartao
    objParcelaReceber.sAprovacaoCartao = tParcelaReceber.sAprovacaoCartao
    objParcelaReceber.dtDataTransacaoCartao = tParcelaReceber.dtDataTransacaoCartao

    'Libera comando
    Call Comando_Fechar(lComando)

    ParcelaReceber_Le_SemNumIntDoc = SUCESSO

    Exit Function

Erro_ParcelaReceber_Le_SemNumIntDoc:

    ParcelaReceber_Le_SemNumIntDoc = Err

    Select Case Err

        Case 28587
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 28588, 28589
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PARCELASREC_NUMINT", Err, objParcelaReceber.lNumIntTitulo, objParcelaReceber.iNumParcela)

        Case 28590
            'Não encontrou. Erro tratado na rotina chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149065)

    End Select

    'Libera comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function TiposInstrCobranca_Le_Todos(ByVal colTiposInstrCobranca As Collection) As Long
'Carrega uma coleção com objetos ClassTipoInstrCobr para todos os tipos da tabela TiposInstrCobranca

Dim lErro As Long
Dim lComando As Long
Dim objTipoInstrCobr As ClassTipoInstrCobr
Dim tTipoInstrCobr As typeTipoInstrCobr

On Error GoTo Erro_TiposInstrCobranca_Le_Todos

    'Inicializa comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 28528

    tTipoInstrCobr.sDescricao = String(STRING_TIPO_INSTR_COBRANCA_DESCRICAO, 0)

    lErro = Comando_Executar(lComando, "SELECT Codigo, RequerDias, Descricao FROM TiposInstrCobranca", tTipoInstrCobr.iCodigo, tTipoInstrCobr.iRequerDias, tTipoInstrCobr.sDescricao)
    If lErro <> AD_SQL_SUCESSO Then Error 28529

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 28530

    Do While (lErro <> AD_SQL_SEM_DADOS)

        Set objTipoInstrCobr = New ClassTipoInstrCobr

        objTipoInstrCobr.iCodigo = tTipoInstrCobr.iCodigo
        objTipoInstrCobr.iRequerDias = tTipoInstrCobr.iRequerDias
        objTipoInstrCobr.sDescricao = tTipoInstrCobr.sDescricao

        colTiposInstrCobranca.Add objTipoInstrCobr

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 28531

    Loop

    'Libera comando
    Call Comando_Fechar(lComando)

    TiposInstrCobranca_Le_Todos = SUCESSO

    Exit Function

Erro_TiposInstrCobranca_Le_Todos:

    TiposInstrCobranca_Le_Todos = Err

    Select Case Err

        Case 28528
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 28529, 28530, 28531
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOSINSTRCOBRANCA", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149066)

    End Select

    'Libera comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function TipoInstrCobranca_Le(ByVal objTipoInstrCobr As ClassTipoInstrCobr) As Long
'le os dados do Tipo Instrucao Cobranca, cujo codigo foi passado como parametro em objTipoInstrCobr

Dim lErro As Long
Dim lComando As Long
Dim tTipoInstrCobr As typeTipoInstrCobr

On Error GoTo Erro_TipoInstrCobranca_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 16546

    tTipoInstrCobr.sDescricao = String(STRING_TIPO_INSTR_COBRANCA_DESCRICAO, 0)
    
    'pesquisa Tipo Instrucao Cobranca no BD
    lErro = Comando_Executar(lComando, "SELECT Descricao, RequerDias FROM TiposInstrCobranca WHERE Codigo = ? ", tTipoInstrCobr.sDescricao, tTipoInstrCobr.iRequerDias, objTipoInstrCobr.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 16547

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 16548

    If lErro = AD_SQL_SEM_DADOS Then Error 16549

    'Preenche objeto
    objTipoInstrCobr.sDescricao = tTipoInstrCobr.sDescricao
    objTipoInstrCobr.iRequerDias = tTipoInstrCobr.iRequerDias

    lErro = Comando_Fechar(lComando)

    TipoInstrCobranca_Le = SUCESSO

    Exit Function

Erro_TipoInstrCobranca_Le:

    TipoInstrCobranca_Le = Err

    Select Case Err

        Case 16546
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 16547, 16548
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPO_INSTRUCAO_COBRANCA", Err, objTipoInstrCobr.iCodigo)

        Case 16549
            'Erro será tratado na rotina chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149067)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function TiposCliente_Le_PadraoCobranca(ByVal iCodigo As Integer) As Long
'verifica se existe tipo de cliente associado a um padrao de cobranca
'retorna 16604 se nao houver tipo de cliente associado

Dim lErro As Long
Dim lComando As Long
Dim iCodigo2 As Integer

On Error GoTo Erro_TiposCliente_Le_PadraoCobranca

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 16601

    'pesquisa Tipos de Cliente no BD
    lErro = Comando_Executar(lComando, "SELECT PadraoCobranca FROM TiposDeCliente WHERE PadraoCobranca = ? ", iCodigo2, iCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 16602

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 16603

    If lErro = AD_SQL_SEM_DADOS Then Error 16604 'Nao existe PadraoCobranca relacionado co Tipos de Cliente

    lErro = Comando_Fechar(lComando)

    TiposCliente_Le_PadraoCobranca = SUCESSO

    Exit Function

Erro_TiposCliente_Le_PadraoCobranca:

    TiposCliente_Le_PadraoCobranca = Err

    Select Case Err

        Case 16601
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 16602, 16603
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOCLIENTE", Err)

        Case 16604

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149068)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Cliente_Le(ByVal objCliente As ClassCliente, Optional ByVal lComandoP1 As Long = 0, Optional ByVal lComandoP2 As Long = 0) As Long
'le o cliente ( e sua matriz ) a partir da chave(lCodigo) fornecida em objCliente. Devolve os dados em objCliente

Dim lComando As Long
Dim lComando2 As Long
Dim lErro As Long
Dim tCliente As typeCliente
Dim lCodigoCliente As Long

On Error GoTo Erro_Cliente_Le

    tCliente.sRazaoSocial = String(STRING_CLIENTE_RAZAO_SOCIAL, 0)
    tCliente.sNomeReduzido = String(STRING_CLIENTE_NOME_REDUZIDO, 0)
    tCliente.sObservacao = String(STRING_CLIENTE_OBSERVACAO, 0)
    tCliente.sNome = String(STRING_FILIAL_CLIENTE_NOME, 0)
    tCliente.sCgc = String(STRING_CGC, 0)
    tCliente.sRG = String(STRING_RG, 0)
    tCliente.sInscricaoEstadual = String(STRING_INSCR_EST, 0)
    tCliente.sInscricaoMunicipal = String(STRING_INSCR_MUN, 0)
    tCliente.sInscricaoSuframa = String(STRING_INSCR_SUF, 0)
    tCliente.sObservacao2 = String(STRING_CLIENTE_OBSERVACAO, 0)
    tCliente.sContaContabil = String(STRING_CONTA, 0)
    tCliente.sGuia = String(STRING_CLIENTE_GUIA, 0)
    tCliente.sUsuarioCobrador = String(STRING_USUARIO_CODIGO, 0)
    tCliente.sUsuRespCallCenter = String(STRING_USUARIO_CODIGO, 0)
    
    'guarda o código correto do cliente para msg de erro
    If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL And objCliente.lCodigo = 0 Then
        lCodigoCliente = objCliente.lCodigoLoja
    Else
        lCodigoCliente = objCliente.lCodigo
    End If

    If lComandoP1 <> 0 Then
        lComando = lComandoP1
    Else
        lComando = Comando_Abrir()
        If lComando = 0 Then Error 12290
    End If

    If lComandoP2 <> 0 Then
        lComando2 = lComandoP2
    Else
        lComando2 = Comando_Abrir()
        If lComando2 = 0 Then Error 12377
    End If

    'Pega os dados da tabela Clientes
    If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL And objCliente.lCodigo = 0 Then
        
        '######################################
        'Alterado por Wagner
        'faço o select com codigoloja
        lErro = Comando_Executar(lComando, "SELECT Codigo, RazaoSocial, NomeReduzido, Tipo, Observacao, LimiteCredito, CondicaoPagto, Desconto, CodMensagem, TabelaPreco,ProxCodFilial, CodPadraoCobranca, Ativo, Transferido, Bloqueado, UsuarioCobrador, UsuRespCallCenter, IgnoraRecebPadrao, TemFaixaReceb,PercentMaisReceb,PercentMenosReceb,RecebForaFaixa  FROM Clientes WHERE FilialEmpresaLoja = ? AND CodigoLoja = ? ", _
            tCliente.lCodigo, tCliente.sRazaoSocial, tCliente.sNomeReduzido, tCliente.iTipo, tCliente.sObservacao, tCliente.dLimiteCredito, tCliente.iCondicaoPagto, tCliente.dDesconto, tCliente.iCodMensagem, tCliente.iTabelaPreco, tCliente.iProxCodFilial, tCliente.iCodPadraoCobranca, tCliente.iAtivo, tCliente.iTransferido, tCliente.iBloqueado, tCliente.sUsuarioCobrador, tCliente.sUsuRespCallCenter, tCliente.iIgnoraRecebPadrao, tCliente.iTemFaixaReceb, tCliente.dPercentMaisReceb, tCliente.dPercentMenosReceb, tCliente.iRecebForaFaixa, objCliente.iFilialEmpresaLoja, objCliente.lCodigoLoja)
        '######################################
    
    Else
        
        '######################################
        'Alterado por Wagner
        'faço o select com codigo
        lErro = Comando_Executar(lComando, "SELECT Codigo, RazaoSocial, NomeReduzido, Tipo, Observacao, LimiteCredito, CondicaoPagto, Desconto, CodMensagem, TabelaPreco,ProxCodFilial, CodPadraoCobranca, Ativo, Transferido, Bloqueado, UsuarioCobrador, UsuRespCallCenter, IgnoraRecebPadrao, TemFaixaReceb,PercentMaisReceb,PercentMenosReceb,RecebForaFaixa  FROM Clientes WHERE Codigo = ?", _
            tCliente.lCodigo, tCliente.sRazaoSocial, tCliente.sNomeReduzido, tCliente.iTipo, tCliente.sObservacao, tCliente.dLimiteCredito, tCliente.iCondicaoPagto, tCliente.dDesconto, tCliente.iCodMensagem, tCliente.iTabelaPreco, tCliente.iProxCodFilial, tCliente.iCodPadraoCobranca, tCliente.iAtivo, tCliente.iTransferido, tCliente.iBloqueado, tCliente.sUsuarioCobrador, tCliente.sUsuRespCallCenter, tCliente.iIgnoraRecebPadrao, tCliente.iTemFaixaReceb, tCliente.dPercentMaisReceb, tCliente.dPercentMenosReceb, tCliente.iRecebForaFaixa, objCliente.lCodigo)
        '######################################
    
    End If
    
    If lErro <> AD_SQL_SUCESSO Then Error 12291

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 12292

    If lErro = AD_SQL_SEM_DADOS Then Error 12293

    objCliente.lCodigo = tCliente.lCodigo
    objCliente.sRazaoSocial = tCliente.sRazaoSocial
    objCliente.sNomeReduzido = tCliente.sNomeReduzido
    objCliente.iTipo = tCliente.iTipo
    objCliente.sObservacao = tCliente.sObservacao
    objCliente.dLimiteCredito = tCliente.dLimiteCredito
    objCliente.iCondicaoPagto = tCliente.iCondicaoPagto
    objCliente.dDesconto = tCliente.dDesconto
    objCliente.iCodMensagem = tCliente.iCodMensagem
    objCliente.iTabelaPreco = tCliente.iTabelaPreco
    objCliente.iProxCodFilial = tCliente.iProxCodFilial
    objCliente.iCodPadraoCobranca = tCliente.iCodPadraoCobranca
    objCliente.iAtivo = tCliente.iAtivo
    objCliente.iTransferido = tCliente.iTransferido
    objCliente.iBloqueado = tCliente.iBloqueado 'Alterado por Wagner
    objCliente.sUsuarioCobrador = tCliente.sUsuarioCobrador 'Alterado por Wagner
    objCliente.sUsuRespCallCenter = tCliente.sUsuRespCallCenter
    
    objCliente.iIgnoraRecebPadrao = tCliente.iIgnoraRecebPadrao 'Alterado por Wagner
    objCliente.iTemFaixaReceb = tCliente.iTemFaixaReceb 'Alterado por Wagner
    objCliente.dPercentMaisReceb = tCliente.dPercentMaisReceb 'Alterado por Wagner
    objCliente.dPercentMenosReceb = tCliente.dPercentMenosReceb 'Alterado por Wagner
    objCliente.iRecebForaFaixa = tCliente.iRecebForaFaixa 'Alterado por Wagner
    
    'Pega os dados da tabela Filiais Clientes
    If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL And objCliente.lCodigo = 0 Then
        lErro = Comando_Executar(lComando2, "SELECT CGC, InscricaoEstadual, InscricaoMunicipal,InscricaoSuframa, Endereco, EnderecoEntrega, EnderecoCobranca, CodTransportadora, Observacao, ContaContabil, Vendedor, ComissaoVendas, Regiao, FreqVisitas, DataUltVisita, CodCobrador, TipoFrete, RG, CodTranspRedesp, DetPagFrete, Guia, RegimeTributario, IEIsento, IENaoContrib FROM FiliaisClientes WHERE FilialEmpresaLoja = ? AND CodClienteLoja = ? AND CodFilialLoja = ?", _
        tCliente.sCgc, tCliente.sInscricaoEstadual, tCliente.sInscricaoMunicipal, tCliente.sInscricaoSuframa, tCliente.lEndereco, tCliente.lEnderecoEntrega, tCliente.lEnderecoCobranca, tCliente.iCodTransportadora, tCliente.sObservacao2, tCliente.sContaContabil, tCliente.iVendedor, tCliente.dComissaoVendas, tCliente.iRegiao, tCliente.iFreqVisitas, tCliente.dtDataUltVisita, tCliente.iCodCobrador, tCliente.iTipoFrete, tCliente.sRG, tCliente.iCodTranspRedesp, tCliente.iDetPagFrete, tCliente.sGuia, tCliente.iRegimeTributario, tCliente.iIEIsento, tCliente.iIENaoContrib, objCliente.iFilialEmpresaLoja, objCliente.lCodigoLoja, FILIAL_MATRIZ)
    Else
        lErro = Comando_Executar(lComando2, "SELECT CGC, InscricaoEstadual, InscricaoMunicipal,InscricaoSuframa, Endereco, EnderecoEntrega, EnderecoCobranca, CodTransportadora, Observacao, ContaContabil, Vendedor, ComissaoVendas, Regiao, FreqVisitas, DataUltVisita, CodCobrador, TipoFrete, RG, CodTranspRedesp, DetPagFrete, Guia, RegimeTributario, IEIsento, IENaoContrib FROM FiliaisClientes WHERE CodCliente = ? AND CodFilial = ?", _
        tCliente.sCgc, tCliente.sInscricaoEstadual, tCliente.sInscricaoMunicipal, tCliente.sInscricaoSuframa, tCliente.lEndereco, tCliente.lEnderecoEntrega, tCliente.lEnderecoCobranca, tCliente.iCodTransportadora, tCliente.sObservacao2, tCliente.sContaContabil, tCliente.iVendedor, tCliente.dComissaoVendas, tCliente.iRegiao, tCliente.iFreqVisitas, tCliente.dtDataUltVisita, tCliente.iCodCobrador, tCliente.iTipoFrete, tCliente.sRG, tCliente.iCodTranspRedesp, tCliente.iDetPagFrete, tCliente.sGuia, tCliente.iRegimeTributario, tCliente.iIEIsento, tCliente.iIENaoContrib, objCliente.lCodigo, FILIAL_MATRIZ)
    End If

    If lErro <> AD_SQL_SUCESSO Then Error 12378

    lErro = Comando_BuscarPrimeiro(lComando2)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 12379

    If lErro = AD_SQL_SEM_DADOS Then Error 12380
    
    objCliente.sRG = tCliente.sRG
    objCliente.sCgc = tCliente.sCgc
    objCliente.sInscricaoEstadual = tCliente.sInscricaoEstadual
    objCliente.sInscricaoMunicipal = tCliente.sInscricaoMunicipal
    objCliente.sInscricaoSuframa = tCliente.sInscricaoSuframa
    objCliente.lEndereco = tCliente.lEndereco
    objCliente.lEnderecoEntrega = tCliente.lEnderecoEntrega
    objCliente.lEnderecoCobranca = tCliente.lEnderecoCobranca
    objCliente.iCodTransportadora = tCliente.iCodTransportadora
    objCliente.sObservacao2 = tCliente.sObservacao2
    objCliente.sContaContabil = tCliente.sContaContabil
    objCliente.iVendedor = tCliente.iVendedor
    objCliente.dComissaoVendas = tCliente.dComissaoVendas
    objCliente.iRegiao = tCliente.iRegiao
    objCliente.iFreqVisitas = tCliente.iFreqVisitas
    objCliente.dtDataUltVisita = tCliente.dtDataUltVisita
    objCliente.iCodCobrador = tCliente.iCodCobrador
    objCliente.iTipoFrete = tCliente.iTipoFrete
    objCliente.iCodTranspRedesp = tCliente.iCodTranspRedesp
    objCliente.iDetPagFrete = tCliente.iDetPagFrete
    objCliente.sGuia = tCliente.sGuia
    objCliente.iRegimeTributario = tCliente.iRegimeTributario
    objCliente.iIEIsento = tCliente.iIEIsento
    objCliente.iIENaoContrib = tCliente.iIENaoContrib
    
    If lComandoP1 = 0 Then Call Comando_Fechar(lComando)
    If lComandoP2 = 0 Then Call Comando_Fechar(lComando2)

    Cliente_Le = SUCESSO

Exit Function

Erro_Cliente_Le:

    Cliente_Le = Err

    Select Case Err

        Case 12290, 12377
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 12378, 12379
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISCLIENTES", Err)

        Case 12291, 12292
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLIENTES", Err)

        Case 12293 'cliente nao cadastrado em Clientes

        Case 12380
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_SEM_FILIAL", Err, lCodigoCliente)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149069)

    End Select

    If lComandoP1 = 0 Then Call Comando_Fechar(lComando)
    If lComandoP2 = 0 Then Call Comando_Fechar(lComando2)

    Exit Function

End Function

Function Cliente_Le_Cgc(ByVal objCliente As ClassCliente, iCodFilial As Integer) As Long
'Le o Cliente a partir do Cgc fornecido em objCliente. Devolve os dados em objCliente
'Devolve combinação de Cliente + FilialCliente com este CGC
'Devolve o código da filial em iCodFilial

Dim lComando As Long
Dim lErro As Long
Dim tCliente As typeCliente

On Error GoTo Erro_Cliente_Le_Cgc

    tCliente.sRazaoSocial = String(STRING_CLIENTE_RAZAO_SOCIAL, 0)
    tCliente.sNomeReduzido = String(STRING_CLIENTE_NOME_REDUZIDO, 0)
    tCliente.sObservacao = String(STRING_CLIENTE_OBSERVACAO, 0)
    tCliente.sCgc = String(STRING_CGC, 0)
    tCliente.sRG = String(STRING_RG, 0)
    tCliente.sInscricaoEstadual = String(STRING_INSCR_EST, 0)
    tCliente.sInscricaoMunicipal = String(STRING_INSCR_MUN, 0)
    tCliente.sInscricaoSuframa = String(STRING_INSCR_SUF, 0)
    tCliente.sObservacao2 = String(STRING_CLIENTE_OBSERVACAO, 0)
    tCliente.sContaContabil = String(STRING_CONTA, 0)
    tCliente.sGuia = String(STRING_CLIENTE_GUIA, 0)
    tCliente.sUsuarioCobrador = String(STRING_USUARIO_CODIGO, 0)
    tCliente.sUsuRespCallCenter = String(STRING_USUARIO_CODIGO, 0)
    
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 6714

    lErro = Comando_Executar(lComando, "SELECT IENaoContrib, IEIsento, CodCliente, CodClienteLoja, CodFilial, CodFilialLoja, FilialEmpresaLoja, InscricaoEstadual, InscricaoMunicipal,InscricaoSuframa, Endereco, EnderecoEntrega, EnderecoCobranca, CodTransportadora, Observacao, ContaContabil, Vendedor, ComissaoVendas, Regiao, FreqVisitas, DataUltVisita, CodCobrador, TipoFrete, RG, CodTranspRedesp, DetPagFrete, Guia FROM FiliaisClientes WHERE CGC=?", _
        tCliente.iIENaoContrib, tCliente.iIEIsento, tCliente.lCodigo, tCliente.lCodigoLoja, tCliente.iCodFilial, tCliente.iCodFilialLoja, tCliente.iFilialEmpresaLoja, tCliente.sInscricaoEstadual, tCliente.sInscricaoMunicipal, tCliente.sInscricaoSuframa, tCliente.lEndereco, tCliente.lEnderecoEntrega, tCliente.lEnderecoCobranca, tCliente.iCodTransportadora, tCliente.sObservacao2, tCliente.sContaContabil, tCliente.iVendedor, tCliente.dComissaoVendas, tCliente.iRegiao, tCliente.iFreqVisitas, tCliente.dtDataUltVisita, tCliente.iCodCobrador, tCliente.iTipoFrete, tCliente.sRG, tCliente.iCodTranspRedesp, tCliente.iDetPagFrete, tCliente.sGuia, objCliente.sCgc)
    
    If lErro <> AD_SQL_SUCESSO Then Error 6708

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 6709
    If lErro = AD_SQL_SEM_DADOS Then Error 6710
    
    If tCliente.iCodFilial <> 0 Then
        iCodFilial = tCliente.iCodFilial
    Else
        iCodFilial = tCliente.iCodFilialLoja
    End If
    
    objCliente.lCodigo = tCliente.lCodigo
    objCliente.lCodigoLoja = tCliente.lCodigoLoja
    objCliente.iFilialEmpresaLoja = tCliente.iFilialEmpresaLoja
    objCliente.sInscricaoEstadual = tCliente.sInscricaoEstadual
    objCliente.sInscricaoMunicipal = tCliente.sInscricaoMunicipal
    objCliente.sInscricaoSuframa = tCliente.sInscricaoSuframa
    objCliente.lEndereco = tCliente.lEndereco
    objCliente.lEnderecoEntrega = tCliente.lEnderecoEntrega
    objCliente.lEnderecoCobranca = tCliente.lEnderecoCobranca
    objCliente.iCodTransportadora = tCliente.iCodTransportadora
    objCliente.sObservacao2 = tCliente.sObservacao2
    objCliente.sContaContabil = tCliente.sContaContabil
    objCliente.iVendedor = tCliente.iVendedor
    objCliente.dComissaoVendas = tCliente.dComissaoVendas
    objCliente.iRegiao = tCliente.iRegiao
    objCliente.iFreqVisitas = tCliente.iFreqVisitas
    objCliente.dtDataUltVisita = tCliente.dtDataUltVisita
    objCliente.iCodCobrador = tCliente.iCodCobrador
    objCliente.iTipoFrete = tCliente.iTipoFrete
    objCliente.sRG = tCliente.sRG
    objCliente.iCodTranspRedesp = tCliente.iCodTranspRedesp
    objCliente.iDetPagFrete = tCliente.iDetPagFrete
    objCliente.sGuia = tCliente.sGuia
    objCliente.iIEIsento = tCliente.iIEIsento
    objCliente.iIENaoContrib = tCliente.iIENaoContrib
    
    If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL Then
        
        'Pega os dados da tabela Clientes
        'Alterado por Wagner - Inserido Ativo/ Inserido Bloqueado 02/03/05
        lErro = Comando_Executar(lComando, "SELECT RazaoSocial, NomeReduzido, Tipo, Observacao, LimiteCredito, CondicaoPagto, Desconto, CodPadraoCobranca, CodMensagem, TabelaPreco, ProxCodFilial, Ativo, Bloqueado, UsuarioCobrador, UsuRespCallCenter, IgnoraRecebPadrao, TemFaixaReceb,PercentMaisReceb,PercentMenosReceb,RecebForaFaixa  FROM Clientes WHERE CodigoLoja = ?", _
            tCliente.sRazaoSocial, tCliente.sNomeReduzido, tCliente.iTipo, tCliente.sObservacao, tCliente.dLimiteCredito, tCliente.iCondicaoPagto, tCliente.dDesconto, tCliente.iCodPadraoCobranca, tCliente.iCodMensagem, tCliente.iTabelaPreco, tCliente.iProxCodFilial, tCliente.iAtivo, tCliente.iBloqueado, tCliente.sUsuarioCobrador, tCliente.sUsuRespCallCenter, tCliente.iIgnoraRecebPadrao, tCliente.iTemFaixaReceb, tCliente.dPercentMaisReceb, tCliente.dPercentMenosReceb, tCliente.iRecebForaFaixa, objCliente.lCodigoLoja)
    
    Else
        
        'Pega os dados da tabela Clientes
        'Alterado por Wagner - Inserido Ativo/ Inserido Bloqueado 02/03/05
        lErro = Comando_Executar(lComando, "SELECT RazaoSocial, NomeReduzido, Tipo, Observacao, LimiteCredito, CondicaoPagto, Desconto, CodPadraoCobranca, CodMensagem, TabelaPreco, ProxCodFilial, Ativo, Bloqueado, UsuarioCobrador, UsuRespCallCenter, IgnoraRecebPadrao, TemFaixaReceb,PercentMaisReceb,PercentMenosReceb,RecebForaFaixa  FROM Clientes WHERE Codigo = ?", _
            tCliente.sRazaoSocial, tCliente.sNomeReduzido, tCliente.iTipo, tCliente.sObservacao, tCliente.dLimiteCredito, tCliente.iCondicaoPagto, tCliente.dDesconto, tCliente.iCodPadraoCobranca, tCliente.iCodMensagem, tCliente.iTabelaPreco, tCliente.iProxCodFilial, tCliente.iAtivo, tCliente.iBloqueado, tCliente.sUsuarioCobrador, tCliente.sUsuRespCallCenter, tCliente.iIgnoraRecebPadrao, tCliente.iTemFaixaReceb, tCliente.dPercentMaisReceb, tCliente.dPercentMenosReceb, tCliente.iRecebForaFaixa, objCliente.lCodigo)
    
    End If

    If lErro <> AD_SQL_SUCESSO Then Error 6711

    lErro = Comando_BuscarPrimeiro(lComando)
    
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 6712
    If lErro = AD_SQL_SEM_DADOS Then Error 6713

    objCliente.sRazaoSocial = tCliente.sRazaoSocial
    objCliente.sNomeReduzido = tCliente.sNomeReduzido
    objCliente.iTipo = tCliente.iTipo
    objCliente.sObservacao = tCliente.sObservacao
    objCliente.dLimiteCredito = tCliente.dLimiteCredito
    objCliente.iCondicaoPagto = tCliente.iCondicaoPagto
    objCliente.dDesconto = tCliente.dDesconto
    objCliente.iCodPadraoCobranca = tCliente.iCodPadraoCobranca
    objCliente.iCodMensagem = tCliente.iCodMensagem
    objCliente.iTabelaPreco = tCliente.iTabelaPreco
    objCliente.iProxCodFilial = tCliente.iProxCodFilial
    
    'Inserido por Wagner
    '#############
    objCliente.iAtivo = tCliente.iAtivo
    objCliente.iBloqueado = tCliente.iBloqueado
    objCliente.sUsuarioCobrador = tCliente.sUsuarioCobrador
    objCliente.sUsuRespCallCenter = tCliente.sUsuRespCallCenter
    '#############
    
    objCliente.iIgnoraRecebPadrao = tCliente.iIgnoraRecebPadrao 'Alterado por Wagner
    objCliente.iTemFaixaReceb = tCliente.iTemFaixaReceb 'Alterado por Wagner
    objCliente.dPercentMaisReceb = tCliente.dPercentMaisReceb 'Alterado por Wagner
    objCliente.dPercentMenosReceb = tCliente.dPercentMenosReceb 'Alterado por Wagner
    objCliente.iRecebForaFaixa = tCliente.iRecebForaFaixa 'Alterado por Wagner

    lErro = Comando_Fechar(lComando)

    Cliente_Le_Cgc = SUCESSO

Exit Function

Erro_Cliente_Le_Cgc:

    Cliente_Le_Cgc = Err

    Select Case Err

        Case 6714
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 6708, 6709
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISCLIENTES", Err)

        Case 6710  'Não existe FilialCliente com este CGC

        Case 6711, 6712
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLIENTES", Err)

        Case 6713
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_DESASSOCIADA_CLIENTE", Err, objCliente.sCgc)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149070)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Cliente_Le_NomeRed(ByVal objCodigoNome As AdmlCodigoNome) As Long
'Le o NomeReduzido do cliente a partir da chave(lCodigo) fornecida em objCodigoNome.

Dim lComando As Long
Dim lErro As Long
Dim sNomeReduzido As String

On Error GoTo Erro_Cliente_Le_NomeRed

    sNomeReduzido = String(STRING_CLIENTE_NOME_REDUZIDO, 0)

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 12550

    If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL Then
        
        'Le o NomeReduzido da tabela Clientes
        lErro = Comando_Executar(lComando, "SELECT NomeReduzido FROM Clientes WHERE CodigoLoja = ?", sNomeReduzido, objCodigoNome.lCodigo)
    
    Else
        
        'Le o NomeReduzido da tabela Clientes
        lErro = Comando_Executar(lComando, "SELECT NomeReduzido FROM Clientes WHERE Codigo = ?", sNomeReduzido, objCodigoNome.lCodigo)
    
    End If

    If lErro <> AD_SQL_SUCESSO Then Error 12551

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 12552

    If lErro = AD_SQL_SEM_DADOS Then Error 12553

    objCodigoNome.sNome = sNomeReduzido

    lErro = Comando_Fechar(lComando)

    Cliente_Le_NomeRed = SUCESSO

Exit Function

Erro_Cliente_Le_NomeRed:

    Cliente_Le_NomeRed = Err

    Select Case Err

        Case 12550
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 12551, 12552
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLIENTES", Err)

        Case 12553 'cliente nao cadastrado em Clientes

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149071)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Cliente_Le_NomeReduzido(ByVal objCliente As ClassCliente) As Long
'le o cliente ( e dados da matriz ) a partir da chave(sNomeReduzido) fornecida em objCliente. Devolve os dados em objCliente

Dim lComando As Long
Dim lErro As Long
Dim tCliente As typeCliente
Dim lCodigoCliente As Long

On Error GoTo Erro_Cliente_Le_NomeReduzido

    tCliente.sRazaoSocial = String(STRING_CLIENTE_RAZAO_SOCIAL, 0)
    tCliente.sObservacao = String(STRING_CLIENTE_OBSERVACAO, 0)
    tCliente.sCgc = String(STRING_CGC, 0)
    tCliente.sRG = String(STRING_RG, 0)
    tCliente.sInscricaoEstadual = String(STRING_INSCR_EST, 0)
    tCliente.sInscricaoMunicipal = String(STRING_INSCR_MUN, 0)
    tCliente.sInscricaoSuframa = String(STRING_INSCR_SUF, 0)
    tCliente.sObservacao2 = String(STRING_CLIENTE_OBSERVACAO, 0)
    tCliente.sContaContabil = String(STRING_CONTA, 0)
    tCliente.sGuia = String(STRING_CLIENTE_GUIA, 0)
    tCliente.sUsuarioCobrador = String(STRING_USUARIO_CODIGO, 0)
    tCliente.sUsuRespCallCenter = String(STRING_USUARIO_CODIGO, 0)
    
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 12345
    
    If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL Then

        'Seleciona os dados da tabela Clientes com o CodigoLoja
        'Alterado por Wagner - Inserido Ativo/ Inserido Bloqueado 02/03/05
        lErro = Comando_Executar(lComando, "SELECT Codigo, CodigoLoja, FilialEmpresaLoja, RazaoSocial, Tipo, Observacao, LimiteCredito, CondicaoPagto, Desconto, CodPadraoCobranca, CodMensagem, TabelaPreco, ProxCodFilial, Ativo, Bloqueado, UsuarioCobrador, UsuRespCallCenter, IgnoraRecebPadrao, TemFaixaReceb,PercentMaisReceb,PercentMenosReceb,RecebForaFaixa  FROM Clientes WHERE NomeReduzido = ?", _
            tCliente.lCodigo, tCliente.lCodigoLoja, tCliente.iFilialEmpresaLoja, tCliente.sRazaoSocial, tCliente.iTipo, tCliente.sObservacao, tCliente.dLimiteCredito, tCliente.iCondicaoPagto, tCliente.dDesconto, tCliente.iCodPadraoCobranca, tCliente.iCodMensagem, tCliente.iTabelaPreco, tCliente.iProxCodFilial, tCliente.iAtivo, tCliente.iBloqueado, tCliente.sUsuarioCobrador, tCliente.sUsuRespCallCenter, tCliente.iIgnoraRecebPadrao, tCliente.iTemFaixaReceb, tCliente.dPercentMaisReceb, tCliente.dPercentMenosReceb, tCliente.iRecebForaFaixa, objCliente.sNomeReduzido)
        
    Else
        
        'Seleciona os dados da tabela Clientes com o Codigo
        'Alterado por Wagner - Inserido Ativo/ Inserido Bloqueado 02/03/05
        lErro = Comando_Executar(lComando, "SELECT Codigo,CodigoLoja, FilialEmpresaLoja, RazaoSocial, Tipo, Observacao, LimiteCredito, CondicaoPagto, Desconto, CodPadraoCobranca, CodMensagem, TabelaPreco, ProxCodFilial, Ativo, Bloqueado, UsuarioCobrador, UsuRespCallCenter, IgnoraRecebPadrao, TemFaixaReceb,PercentMaisReceb,PercentMenosReceb,RecebForaFaixa  FROM Clientes WHERE NomeReduzido = ?", _
            tCliente.lCodigo, tCliente.lCodigoLoja, tCliente.iFilialEmpresaLoja, tCliente.sRazaoSocial, tCliente.iTipo, tCliente.sObservacao, tCliente.dLimiteCredito, tCliente.iCondicaoPagto, tCliente.dDesconto, tCliente.iCodPadraoCobranca, tCliente.iCodMensagem, tCliente.iTabelaPreco, tCliente.iProxCodFilial, tCliente.iAtivo, tCliente.iBloqueado, tCliente.sUsuarioCobrador, tCliente.sUsuRespCallCenter, tCliente.iIgnoraRecebPadrao, tCliente.iTemFaixaReceb, tCliente.dPercentMaisReceb, tCliente.dPercentMenosReceb, tCliente.iRecebForaFaixa, objCliente.sNomeReduzido)
        
    End If

    If lErro <> AD_SQL_SUCESSO Then Error 12346

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 12347

    If lErro = AD_SQL_SEM_DADOS Then Error 12348

    objCliente.lCodigo = tCliente.lCodigo
    objCliente.lCodigoLoja = tCliente.lCodigoLoja
    objCliente.iFilialEmpresaLoja = tCliente.iFilialEmpresaLoja
    objCliente.sRazaoSocial = tCliente.sRazaoSocial
    objCliente.iTipo = tCliente.iTipo
    objCliente.sObservacao = tCliente.sObservacao
    objCliente.dLimiteCredito = tCliente.dLimiteCredito
    objCliente.iCondicaoPagto = tCliente.iCondicaoPagto
    objCliente.dDesconto = tCliente.dDesconto
    objCliente.iCodPadraoCobranca = tCliente.iCodPadraoCobranca
    objCliente.iCodMensagem = tCliente.iCodMensagem
    objCliente.iTabelaPreco = tCliente.iTabelaPreco
    objCliente.iProxCodFilial = tCliente.iProxCodFilial
    
    'Inserido Por Wagner
    '######
    objCliente.iAtivo = tCliente.iAtivo
    objCliente.iBloqueado = tCliente.iBloqueado
    objCliente.sUsuarioCobrador = tCliente.sUsuarioCobrador
    objCliente.sUsuRespCallCenter = tCliente.sUsuRespCallCenter
    '######

    objCliente.iIgnoraRecebPadrao = tCliente.iIgnoraRecebPadrao 'Alterado por Wagner
    objCliente.iTemFaixaReceb = tCliente.iTemFaixaReceb 'Alterado por Wagner
    objCliente.dPercentMaisReceb = tCliente.dPercentMaisReceb 'Alterado por Wagner
    objCliente.dPercentMenosReceb = tCliente.dPercentMenosReceb 'Alterado por Wagner
    objCliente.iRecebForaFaixa = tCliente.iRecebForaFaixa 'Alterado por Wagner

    'Seleciona os dados da tabela Filiais Clientes
    If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL And objCliente.lCodigo = 0 Then
        lErro = Comando_Executar(lComando, "SELECT IENaoContrib, IEIsento, CGC, InscricaoEstadual, InscricaoMunicipal, InscricaoSuframa, Endereco, EnderecoEntrega, EnderecoCobranca, CodTransportadora, Observacao, ContaContabil, Vendedor, ComissaoVendas, Regiao, FreqVisitas, DataUltVisita, CodCobrador, TipoFrete, RG, CodTranspRedesp, DetPagFrete, Guia, RegimeTributario FROM FiliaisClientes WHERE CodCliente = ? AND CodFilial = ?", _
                tCliente.iIENaoContrib, tCliente.iIEIsento, tCliente.sCgc, tCliente.sInscricaoEstadual, tCliente.sInscricaoMunicipal, tCliente.sInscricaoSuframa, tCliente.lEndereco, tCliente.lEnderecoEntrega, tCliente.lEnderecoCobranca, tCliente.iCodTransportadora, tCliente.sObservacao2, tCliente.sContaContabil, tCliente.iVendedor, tCliente.dComissaoVendas, tCliente.iRegiao, tCliente.iFreqVisitas, tCliente.dtDataUltVisita, tCliente.iCodCobrador, tCliente.iTipoFrete, tCliente.sRG, tCliente.iCodTranspRedesp, tCliente.iDetPagFrete, tCliente.sGuia, tCliente.iRegimeTributario, objCliente.lCodigoLoja, FILIAL_MATRIZ)
    Else
        lErro = Comando_Executar(lComando, "SELECT IENaoContrib, IEIsento, CGC, InscricaoEstadual, InscricaoMunicipal, InscricaoSuframa, Endereco, EnderecoEntrega, EnderecoCobranca, CodTransportadora, Observacao, ContaContabil, Vendedor, ComissaoVendas, Regiao, FreqVisitas, DataUltVisita, CodCobrador, TipoFrete, RG, CodTranspRedesp, DetPagFrete, Guia, RegimeTributario FROM FiliaisClientes WHERE CodCliente = ? AND CodFilial = ?", _
                tCliente.iIENaoContrib, tCliente.iIEIsento, tCliente.sCgc, tCliente.sInscricaoEstadual, tCliente.sInscricaoMunicipal, tCliente.sInscricaoSuframa, tCliente.lEndereco, tCliente.lEnderecoEntrega, tCliente.lEnderecoCobranca, tCliente.iCodTransportadora, tCliente.sObservacao2, tCliente.sContaContabil, tCliente.iVendedor, tCliente.dComissaoVendas, tCliente.iRegiao, tCliente.iFreqVisitas, tCliente.dtDataUltVisita, tCliente.iCodCobrador, tCliente.iTipoFrete, tCliente.sRG, tCliente.iCodTranspRedesp, tCliente.iDetPagFrete, tCliente.sGuia, tCliente.iRegimeTributario, objCliente.lCodigo, FILIAL_MATRIZ)
    End If
    
    If lErro <> AD_SQL_SUCESSO Then Error 12488

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 12489
    If lErro = AD_SQL_SEM_DADOS Then Error 12490
    
    objCliente.sRG = tCliente.sRG
    objCliente.sCgc = tCliente.sCgc
    objCliente.sInscricaoEstadual = tCliente.sInscricaoEstadual
    objCliente.sInscricaoMunicipal = tCliente.sInscricaoMunicipal
    objCliente.sInscricaoSuframa = tCliente.sInscricaoSuframa
    objCliente.lEndereco = tCliente.lEndereco
    objCliente.lEnderecoEntrega = tCliente.lEnderecoEntrega
    objCliente.lEnderecoCobranca = tCliente.lEnderecoCobranca
    objCliente.iCodTransportadora = tCliente.iCodTransportadora
    objCliente.sObservacao2 = tCliente.sObservacao2
    objCliente.sContaContabil = tCliente.sContaContabil
    objCliente.iVendedor = tCliente.iVendedor
    objCliente.dComissaoVendas = tCliente.dComissaoVendas
    objCliente.iRegiao = tCliente.iRegiao
    objCliente.iFreqVisitas = tCliente.iFreqVisitas
    objCliente.dtDataUltVisita = tCliente.dtDataUltVisita
    objCliente.iCodCobrador = tCliente.iCodCobrador
    objCliente.iTipoFrete = tCliente.iTipoFrete
    objCliente.iCodTranspRedesp = tCliente.iCodTranspRedesp
    objCliente.iDetPagFrete = tCliente.iDetPagFrete
    objCliente.sGuia = tCliente.sGuia
    objCliente.iRegimeTributario = tCliente.iRegimeTributario
    objCliente.iIEIsento = tCliente.iIEIsento
    objCliente.iIENaoContrib = tCliente.iIENaoContrib
    
    Call Comando_Fechar(lComando)

    Cliente_Le_NomeReduzido = SUCESSO

Exit Function

Erro_Cliente_Le_NomeReduzido:

    Cliente_Le_NomeReduzido = Err

    Select Case Err

        Case 12345
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 12346, 12347
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLIENTES", Err)

        Case 12348 'nao existe NomeReduzido (cliente nao cadastrado)

        Case 12488, 12489
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISCLIENTES", Err)

        Case 12490
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_SEM_FILIAL", Err, lCodigoCliente)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149072)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Cliente_ObterDadosBasicos(ByVal objCliente As ClassCliente) As Long
'le o cliente a partir da chave(lCodigo) fornecida em objCliente.
'Devolve os dados em objCliente
'nao acessa dados da matriz do cliente

Dim lComando As Long
Dim lErro As Long
Dim tCliente As typeCliente
Dim lCodigoCliente As Long

On Error GoTo Erro_Cliente_ObterDadosBasicos

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 7613

    tCliente.sRazaoSocial = String(STRING_CLIENTE_RAZAO_SOCIAL, 0)
    tCliente.sNomeReduzido = String(STRING_CLIENTE_NOME_REDUZIDO, 0)
    tCliente.sObservacao = String(STRING_CLIENTE_OBSERVACAO, 0)
    tCliente.sNome = String(STRING_FILIAL_CLIENTE_NOME, 0)
    tCliente.sCgc = String(STRING_CGC, 0)
    tCliente.sRG = String(STRING_RG, 0)
    tCliente.sInscricaoEstadual = String(STRING_INSCR_EST, 0)
    tCliente.sInscricaoMunicipal = String(STRING_INSCR_MUN, 0)
    tCliente.sObservacao2 = String(STRING_CLIENTE_OBSERVACAO, 0)
    tCliente.sContaContabil = String(STRING_CONTA, 0)

    If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL Then
        
        lCodigoCliente = objCliente.lCodigoLoja
        
        'Pega os dados da tabela Clientes
        lErro = Comando_Executar(lComando, "SELECT RazaoSocial, NomeReduzido, Tipo, Observacao, LimiteCredito, CondicaoPagto, Desconto, CodMensagem, TabelaPreco, ProxCodFilial FROM Clientes WHERE CodigoLoja = ?", _
        tCliente.sRazaoSocial, tCliente.sNomeReduzido, tCliente.iTipo, tCliente.sObservacao, tCliente.dLimiteCredito, tCliente.iCondicaoPagto, tCliente.dDesconto, tCliente.iCodMensagem, tCliente.iTabelaPreco, tCliente.iProxCodFilial, objCliente.lCodigoLoja)
        
    
    Else
        
        lCodigoCliente = objCliente.lCodigo
        
        'Pega os dados da tabela Clientes
        lErro = Comando_Executar(lComando, "SELECT RazaoSocial, NomeReduzido, Tipo, Observacao, LimiteCredito, CondicaoPagto, Desconto, CodMensagem, TabelaPreco, ProxCodFilial FROM Clientes WHERE Codigo = ?", _
        tCliente.sRazaoSocial, tCliente.sNomeReduzido, tCliente.iTipo, tCliente.sObservacao, tCliente.dLimiteCredito, tCliente.iCondicaoPagto, tCliente.dDesconto, tCliente.iCodMensagem, tCliente.iTabelaPreco, tCliente.iProxCodFilial, objCliente.lCodigo)
    
    End If
    
    If lErro <> AD_SQL_SUCESSO Then Error 7614

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 7615

    If lErro = AD_SQL_SEM_DADOS Then Error 7616

    objCliente.sRazaoSocial = tCliente.sRazaoSocial
    objCliente.sNomeReduzido = tCliente.sNomeReduzido
    objCliente.iTipo = tCliente.iTipo
    objCliente.sObservacao = tCliente.sObservacao
    objCliente.dLimiteCredito = tCliente.dLimiteCredito
    objCliente.iCondicaoPagto = tCliente.iCondicaoPagto
    objCliente.dDesconto = tCliente.dDesconto
    objCliente.iCodMensagem = tCliente.iCodMensagem
    objCliente.iTabelaPreco = tCliente.iTabelaPreco
    objCliente.iProxCodFilial = tCliente.iProxCodFilial

    lErro = Comando_Fechar(lComando)

    Cliente_ObterDadosBasicos = SUCESSO

    Exit Function

Erro_Cliente_ObterDadosBasicos:

    Cliente_ObterDadosBasicos = Err

    Select Case Err

        Case 7613
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 7614, 7615
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLIENTES", Err)

        Case 7616
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", Err, lCodigoCliente)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149073)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Clientes_Le_Tabelas_Lock(ByVal objCliente As ClassCliente) As Long
'Faz o lock das tabelas relacionadas a tabela Cliente para assegurar a gravação de um novo Cliente
'Verifica se Condição de Pagto pode ser usada em Contas Receber
'Chamada DENTRO DE TRANSAÇÃO por Cliente_Grava

Dim alComando(10) As Long
Dim iCodigo As Integer
Dim iTipoConta As Integer
Dim iAtivo As Integer
Dim iEmRecebimento As Integer
Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Clientes_Le_Tabelas_Lock

    For iIndice = 0 To 9
    
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 12390

    Next
    
    If giLocalOperacao <> LOCALOPERACAO_CAIXA_CENTRAL Then

        'Tipos de Cliente
        If objCliente.iTipo > 0 Then
        
            lErro = Comando_ExecutarLockado(alComando(0), "SELECT Codigo FROM TiposDeCliente WHERE Codigo=?", iCodigo, objCliente.iTipo)
            If lErro <> AD_SQL_SUCESSO Then Error 12399
        
            lErro = Comando_BuscarPrimeiro(alComando(0))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 12400
            If lErro = AD_SQL_SEM_DADOS Then Error 6928
        
            lErro = Comando_LockShared(alComando(0))
            If lErro <> AD_SQL_SUCESSO Then Error 12401
            
        End If
    End If
    
    'Tabelas de Preco
    If objCliente.iTabelaPreco > 0 Then
    
        lErro = CF("TabelaDePreco_Lock", objCliente.iTabelaPreco)
        If lErro <> SUCESSO Then Error 12402
        
    End If

    'CondicoesPagto
    If objCliente.iCondicaoPagto > 0 Then
    
        lErro = Comando_ExecutarLockado(alComando(2), "SELECT EmRecebimento FROM CondicoesPagto WHERE Codigo=?", iEmRecebimento, objCliente.iCondicaoPagto)
        If lErro <> AD_SQL_SUCESSO Then Error 12405
    
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 12406
        If lErro = AD_SQL_SEM_DADOS Then Error 6930
    
        lErro = Comando_LockShared(alComando(2))
        If lErro <> AD_SQL_SUCESSO Then Error 12407
        
        'Verifica se CondicaoPagto pode ser usada em Contas a Receber
        If iEmRecebimento = 0 Then Error 25000
        
    End If

    'Mensagem
    If objCliente.iCodMensagem > 0 Then
    
        lErro = CF("Mensagem_Lock", alComando(3), objCliente.iCodMensagem)
        If lErro <> SUCESSO Then Error 12408
    
    End If

    'PlanoConta
    If Len(Trim(objCliente.sContaContabil)) > 0 Then
    
        lErro = CF("Conta_Analitica_Lock", objCliente.sContaContabil, MODULO_CONTABILIDADE)
        If lErro <> SUCESSO Then Error 12411
        
    End If

    'Cobrador
    If objCliente.iCodCobrador > 0 Then
        
        lErro = CF("Cobrador_Lock", alComando(5), objCliente.iCodCobrador)
        If lErro <> SUCESSO Then Error 12414
        
    End If

    'Transportadora
    If objCliente.iCodTransportadora > 0 Then
        
        lErro = CF("Transportadora_Lock", alComando(6), objCliente.iCodTransportadora)
        If lErro <> SUCESSO Then Error 12417
        
    End If

    'Transportadora de Redespacho
    If objCliente.iCodTranspRedesp > 0 Then
        
        lErro = CF("Transportadora_Lock", alComando(6), objCliente.iCodTranspRedesp)
        If lErro <> SUCESSO Then Error 12417
        
    End If

    'Vendedor
    If objCliente.iVendedor > 0 Then
    
        lErro = CF("Vendedor_Lock", alComando(7), objCliente.iVendedor)
        If lErro <> SUCESSO Then Error 12420
        
    End If
    
    'RegioesVendas
    If objCliente.iRegiao > 0 Then
    
        lErro = CF("RegiaoVenda_Lock", alComando(8), objCliente.iRegiao)
        If lErro <> SUCESSO Then Error 12423
        
    End If
    
    'Padrão de Cobrança
    If objCliente.iCodPadraoCobranca > 0 Then
    
        lErro = CF("PadraoCobranca_Lock", alComando(9), objCliente.iCodPadraoCobranca)
        If lErro <> SUCESSO Then Error 58257
    
    End If
        
    For iIndice = 0 To 9
        Call Comando_Fechar(alComando(iIndice))
    Next

    Clientes_Le_Tabelas_Lock = SUCESSO

    Exit Function

Erro_Clientes_Le_Tabelas_Lock:

    Clientes_Le_Tabelas_Lock = Err

    Select Case Err

        Case 6928
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_CLIENTE_NAO_CADASTRADO", Err, objCliente.iTipo)
        
        Case 6930
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_CADASTRADA", Err, objCliente.iCondicaoPagto)
        
        Case 12390
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 12399, 12400
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOSCLIENTE", Err, objCliente.iTipo)

        Case 12401
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_TIPOSCLIENTE", Err)

        Case 12405, 12406
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONDICAOPAGTO", Err, objCliente.iCondicaoPagto)

        Case 12407
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_CONDICAOPAGTO", Err)

        Case 12402, 12408, 12411, 12414, 12417, 12420, 12423, 58257
        
        Case 25000
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_RECEBIMENTO", Err, objCliente.iCondicaoPagto)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149074)

    End Select

    For iIndice = 0 To 9
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function NFs_Cliente_Nao_Relacionado(ByVal lCodCliente As Long) As Long
'Verifica se Cliente não está relacionado com Notas Fiscais (lançadas ou baixadas)
'Retorna SUCESSO se não estiver relacionado.
Dim lErro As Long, lComando As Long
Dim lNumIntDoc As Long
On Error GoTo Erro_NFs_Cliente_Nao_Relacionado

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 27491

    lErro = Comando_Executar(lComando, "SELECT NumIntDoc FROM NFiscal WHERE Cliente = ?", lNumIntDoc, lCodCliente)
    If lErro <> AD_SQL_SUCESSO Then Error 27492

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 27493
    
    'se achou alguma nota relacionada
    If lErro = AD_SQL_SUCESSO Then Error 27494
    
    lErro = Comando_Fechar(lComando)
    
    NFs_Cliente_Nao_Relacionado = SUCESSO

    Exit Function

Erro_NFs_Cliente_Nao_Relacionado:

    NFs_Cliente_Nao_Relacionado = Err

    Select Case Err

        Case 27491
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
                    
        Case 27492, 27493
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NOTAS_FISCAIS_CLI", Err)
        
        Case 27494 'achou alguma nota relacionada
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXISTEM_NOTAS_FISCAIS_CLI", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149075)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function TitulosReceber_Cliente_Nao_Relacionado(ByVal lCodCliente As Long) As Long
'Verifica se Cliente não está relacionado com Títulos a Receber (lançados ou baixados)
'Retorna SUCESSO se não estiver relacionado.

Dim colCodTitulosReceber As New Collection
Dim lErro As Long
    
On Error GoTo Erro_TitulosReceber_Cliente_Nao_Relacionado
    
    'Preenche a coleção de códigos de Títulos a Receber relacionados ao Cliente
    lErro = CF("TitulosReceber_Le_Cliente", lCodCliente, colCodTitulosReceber)
    If lErro <> SUCESSO And lErro <> 12443 Then Error 12445

    'Verifica se Cliente está relacionado com Títulos a Receber
    If lErro = SUCESSO Then Error 12446
   
    'Preenche a coleção de códigos de Títulos a Receber Baixados relacionados ao Cliente
    lErro = CF("TitulosReceberBaixados_Le_Cliente", lCodCliente, colCodTitulosReceber)
    If lErro <> SUCESSO And lErro <> 6871 Then Error 6874

    'Verifica se Cliente está relacionado com Títulos a Receber Baixado
    If lErro = SUCESSO Then Error 6876

    TitulosReceber_Cliente_Nao_Relacionado = SUCESSO

    Exit Function

Erro_TitulosReceber_Cliente_Nao_Relacionado:

    TitulosReceber_Cliente_Nao_Relacionado = Err

    Select Case Err
        
        Case 6874, 12445
        
        Case 6876
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_REL_TITULOS_REC_BAIXADOS", Err, lCodCliente)

        Case 12446
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_REL_TITULOS_REC", Err, lCodCliente)
      
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149076)

    End Select

    Exit Function

End Function

Function DebitosRecCli_Le_Cliente(ByVal lCodigo As Long) As Long
'Le todos os códigos internos de Débitos a Receber relacionados à chave de Cliente (lcodigo) passada

Dim lComando As Long
Dim lNumIntDoc As Long
Dim lErro As Long

On Error GoTo Erro_DebitosRecCli_Le_Cliente

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 12449

    lErro = Comando_Executar(lComando, "SELECT NumIntDoc FROM DebitosRecCli WHERE Cliente = ? ORDER BY NumIntDoc", lNumIntDoc, lCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 12450

    'le o primeiro debito
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 12451

    If lErro = AD_SQL_SEM_DADOS Then Error 12452

    lErro = Comando_Fechar(lComando)

    DebitosRecCli_Le_Cliente = SUCESSO

    Exit Function

Erro_DebitosRecCli_Le_Cliente:

    DebitosRecCli_Le_Cliente = Err

    Select Case Err

        Case 12449
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 12450, 12451, 12453
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_DEBITOSRECCLI", Err)

        Case 12452 'sem dados

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149077)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function RecebAntecipados_Le_Cliente(ByVal lCodigo As Long) As Long
'Le todos os códigos internos de Recebimentos Antecipados relacionados à chave de Cliente (lcodigo) passada

Dim lComando As Long
Dim lNumIntRec As Long
Dim lErro As Long

On Error GoTo Erro_RecebAntecipados_Le_Cliente

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 12456

    lErro = Comando_Executar(lComando, "SELECT NumIntRec FROM RecebAntecipados WHERE Cliente = ? ORDER BY NumIntRec", lNumIntRec, lCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 12457

    'le o primeiro movimento
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 12458

    If lErro = AD_SQL_SEM_DADOS Then Error 12459

    lErro = Comando_Fechar(lComando)

    RecebAntecipados_Le_Cliente = SUCESSO

    Exit Function

Erro_RecebAntecipados_Le_Cliente:

    RecebAntecipados_Le_Cliente = Err

    Select Case Err

        Case 12456
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 12457, 12458, 12460
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RECEB_ANTEC", Err)

        Case 12459 'sem dados

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149078)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function TituloReceber_Le_SemNumIntDoc(ByVal objTituloReceber As ClassTituloReceber) As Long
'Lê dados do Título a partir de FilialEmpresa, Cliente, Filial, SiglaDocumento, NumTitulo

Dim lErro As Long
Dim lComando As Long
Dim tTituloReceber As typeTituloReceber

On Error GoTo Erro_TituloReceber_Le_SemNumIntDoc

    'Inicializa comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 28571

    tTituloReceber.sObservacao = String(STRING_TITULO_OBSERVACAO, 0)
    tTituloReceber.sCcl = String(STRING_CCL, 0)
    tTituloReceber.sNatureza = String(STRING_NATMOVCTA_CODIGO, 0)

    If objTituloReceber.dtDataEmissao <> DATA_NULA Then
        lErro = Comando_Executar(lComando, "SELECT NumIntDoc, DataEmissao, Status, Saldo, NumParcelas, Valor, ValorIRRF, ValorISS, ISSRetido, ValorINSS, INSSRetido, PercJurosDiarios, PercMulta, Observacao, CondicaoPagto, DataRegistro, PISRetido, COFINSRetido, CSLLRetido, Moeda, Ccl, Natureza, CpoGenerico1, ReajustePeriodicidade, ReajusteBase, ReajustadoAte FROM TitulosRec WHERE FilialEmpresa = ? AND Cliente = ? AND Filial = ? AND SiglaDocumento = ? AND NumTitulo = ? AND Status <> ? AND DataEmissao = ?", _
            tTituloReceber.lNumIntDoc, tTituloReceber.dtDataEmissao, tTituloReceber.iStatus, tTituloReceber.dSaldo, tTituloReceber.iNumParcelas, tTituloReceber.dValor, tTituloReceber.dValorIRRF, tTituloReceber.dValorISS, tTituloReceber.dISSRetido, tTituloReceber.dValorINSS, tTituloReceber.iINSSRetido, tTituloReceber.dPercJurosDiario, tTituloReceber.dPercMulta, tTituloReceber.sObservacao, tTituloReceber.iCondicaoPagto, tTituloReceber.dtDataRegistro, tTituloReceber.dPISRetido, tTituloReceber.dCOFINSRetido, tTituloReceber.dCSLLRetido, tTituloReceber.iMoeda, tTituloReceber.sCcl, tTituloReceber.sNatureza, tTituloReceber.lCpoGenerico1, tTituloReceber.iReajustePeriodicidade, tTituloReceber.dtReajusteBase, tTituloReceber.dtReajustadoAte, objTituloReceber.iFilialEmpresa, objTituloReceber.lCliente, objTituloReceber.iFilial, objTituloReceber.sSiglaDocumento, objTituloReceber.lNumTitulo, STATUS_EXCLUIDO, objTituloReceber.dtDataEmissao)
    Else
        lErro = Comando_Executar(lComando, "SELECT NumIntDoc, DataEmissao, Status, Saldo, NumParcelas, Valor, ValorIRRF, ValorISS, ISSRetido, ValorINSS, INSSRetido, PercJurosDiarios, PercMulta, Observacao, CondicaoPagto, DataRegistro, PISRetido, COFINSRetido, CSLLRetido, Moeda, Ccl, Natureza, CpoGenerico1, ReajustePeriodicidade, ReajusteBase, ReajustadoAte FROM TitulosRec WHERE FilialEmpresa = ? AND Cliente = ? AND Filial = ? AND SiglaDocumento = ? AND NumTitulo = ? AND Status <> ? ORDER BY DataEmissao", _
            tTituloReceber.lNumIntDoc, tTituloReceber.dtDataEmissao, tTituloReceber.iStatus, tTituloReceber.dSaldo, tTituloReceber.iNumParcelas, tTituloReceber.dValor, tTituloReceber.dValorIRRF, tTituloReceber.dValorISS, tTituloReceber.dISSRetido, tTituloReceber.dValorINSS, tTituloReceber.iINSSRetido, tTituloReceber.dPercJurosDiario, tTituloReceber.dPercMulta, tTituloReceber.sObservacao, tTituloReceber.iCondicaoPagto, tTituloReceber.dtDataRegistro, tTituloReceber.dPISRetido, tTituloReceber.dCOFINSRetido, tTituloReceber.dCSLLRetido, tTituloReceber.iMoeda, tTituloReceber.sCcl, tTituloReceber.sNatureza, tTituloReceber.lCpoGenerico1, tTituloReceber.iReajustePeriodicidade, tTituloReceber.dtReajusteBase, tTituloReceber.dtReajustadoAte, objTituloReceber.iFilialEmpresa, objTituloReceber.lCliente, objTituloReceber.iFilial, objTituloReceber.sSiglaDocumento, objTituloReceber.lNumTitulo, STATUS_EXCLUIDO)
    End If
    If lErro <> AD_SQL_SUCESSO Then Error 28572

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 28573

    If lErro = AD_SQL_SEM_DADOS Then Error 28574

    tTituloReceber.iFilialEmpresa = objTituloReceber.iFilialEmpresa
    tTituloReceber.lNumTitulo = objTituloReceber.lNumTitulo
    tTituloReceber.sSiglaDocumento = objTituloReceber.sSiglaDocumento
    tTituloReceber.lCliente = objTituloReceber.lCliente
    tTituloReceber.iFilial = objTituloReceber.iFilial
    If objTituloReceber.dtDataEmissao <> DATA_NULA Then tTituloReceber.dtDataEmissao = objTituloReceber.dtDataEmissao

    Call TituloReceber_MoverCamposObj(objTituloReceber, tTituloReceber)
    
    'Libera comando
    Call Comando_Fechar(lComando)

    TituloReceber_Le_SemNumIntDoc = SUCESSO

    Exit Function

Erro_TituloReceber_Le_SemNumIntDoc:

    TituloReceber_Le_SemNumIntDoc = Err

    Select Case Err

        Case 28571
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 28572, 28573
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOSREC2", Err, objTituloReceber.iFilialEmpresa, objTituloReceber.lCliente, objTituloReceber.iFilial, objTituloReceber.sSiglaDocumento, objTituloReceber.lNumTitulo)

        Case 28574
            'Não encontrou. Erro tratado na rotina chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149079)

    End Select

    'Libera comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function TitulosReceber_Le_Cliente(ByVal lCodigo As Long, ByVal colCodTitulosReceber As Collection) As Long
'Lê todos códigos de Títulos a Receber relacionados a chave do Cliente (lcodigo) passada
'Devolve em colCodTitulosReceber

Dim lComando As Long
Dim lNumIntDoc As Long
Dim lErro As Long

On Error GoTo Erro_TitulosReceber_Le_Cliente

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 12440

    'Seleciona os códigos de Títulos a Receber relacionados ao Cliente
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc FROM TitulosRec WHERE Cliente = ? ORDER BY NumIntDoc", lNumIntDoc, lCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 12441

    'Lê o primeiro código de Título a Receber
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 12442
    If lErro = AD_SQL_SEM_DADOS Then Error 12443

    Do While lErro <> AD_SQL_SEM_DADOS

        'Adiciona o código do Título a Receber lido na coleção
        colCodTitulosReceber.Add lNumIntDoc

        'Lê o próximo Título
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 12444

    Loop

    lErro = Comando_Fechar(lComando)

    TitulosReceber_Le_Cliente = SUCESSO

    Exit Function

Erro_TitulosReceber_Le_Cliente:

    TitulosReceber_Le_Cliente = Err

    Select Case Err

        Case 12440
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 12441, 12442, 12444
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOS_REC", Err)

        Case 12443 'Sem dados. A ser tratado na rotina chamadora.

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149080)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function TitulosReceberBaixados_Le_Cliente(ByVal lCodigo As Long, ByVal colCodTitulosReceber As Collection) As Long
'Lê todos códigos de Títulos a Receber Baixados relacionados a chave do Cliente (lcodigo) passada
'Devolve em colCodTitulosReceber

Dim lComando As Long
Dim lNumIntDoc As Long
Dim lErro As Long

On Error GoTo Erro_TitulosReceberBaixados_Le_Cliente

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 6868

    'Seleciona os códigos de Títulos a Receber Baixados relacionados ao Cliente
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc FROM TitulosRecBaixados WHERE Cliente = ? ORDER BY NumIntDoc", lNumIntDoc, lCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 6869

    'Lê o primeiro código de Título a Receber
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 6870
    If lErro = AD_SQL_SEM_DADOS Then Error 6871

    Do While lErro <> AD_SQL_SEM_DADOS

        'Adiciona o código do Título a Receber lido na coleção
        colCodTitulosReceber.Add lNumIntDoc

        'Lê o próximo Título
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 6872

    Loop

    lErro = Comando_Fechar(lComando)

    TitulosReceberBaixados_Le_Cliente = SUCESSO

    Exit Function

Erro_TitulosReceberBaixados_Le_Cliente:

    TitulosReceberBaixados_Le_Cliente = Err

    Select Case Err

        Case 6868
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 6869, 6870, 6872
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOS_REC_BAIXADOS", Err)

        Case 6871 'Sem dados. A ser tratado na rotina chamadora.

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149081)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function CategoriaCliente_Le_Todos(ByVal colCategorias As Collection) As Long
'Lê na tabela CategoriaCliente todas as Categorias, passando para a coleção colCategorias

Dim lErro As Long
Dim lComando As Long
Dim objCategoriaCliente As ClassCategoriaCliente
Dim sCategoria As String
Dim sDescricao As String

On Error GoTo Erro_CategoriaCliente_Le_Todos

    'Abertura comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 28840

    sCategoria = String(STRING_CATEGORIACLIENTE_CATEGORIA, 0)
    sDescricao = String(STRING_CATEGORIACLIENTE_DESCRICAO, 0)

    'Pesquisa no BD as Categorias
    lErro = Comando_Executar(lComando, "SELECT Categoria, Descricao FROM CategoriaCliente", sCategoria, sDescricao)
    If lErro <> AD_SQL_SUCESSO Then Error 28841

    'Verifica se existe
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 28842

    'Passa para a collection
    Do While lErro = AD_SQL_SUCESSO

        Set objCategoriaCliente = New ClassCategoriaCliente

        objCategoriaCliente.sCategoria = sCategoria
        objCategoriaCliente.sDescricao = sDescricao

        colCategorias.Add objCategoriaCliente

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 28844

    Loop

    'Fechamento comando
    Call Comando_Fechar(lComando)

    CategoriaCliente_Le_Todos = SUCESSO

    Exit Function

Erro_CategoriaCliente_Le_Todos:

    CategoriaCliente_Le_Todos = Err

    Select Case Err

        Case 28840
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 28841, 28842, 28844
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CATEGORIACLIENTE", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149082)

    End Select

    'Fechamento comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function CategoriaCliente_Le(ByVal objCategoriaCliente As ClassCategoriaCliente) As Long
'Lê na tabela CategoriaCliente a descrição da categoria

Dim lErro As Long
Dim lComando As Long
Dim sDescricao As String

On Error GoTo Erro_CategoriaCliente_Le

    'Abertura comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 22364

    sDescricao = String(STRING_CATEGORIACLIENTE_DESCRICAO, 0)

    'Pesquisa no BD a categoria passada como parâmetro
    lErro = Comando_Executar(lComando, "SELECT Descricao FROM CategoriaCliente WHERE Categoria = ?", sDescricao, objCategoriaCliente.sCategoria)
    If lErro <> AD_SQL_SUCESSO Then Error 28845

    'Verifica se existe
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 28846

    'Se não encontrou
    If lErro = AD_SQL_SEM_DADOS Then Error 28847

    objCategoriaCliente.sDescricao = sDescricao

    'Fechamento comando
    Call Comando_Fechar(lComando)

    CategoriaCliente_Le = SUCESSO

    Exit Function

Erro_CategoriaCliente_Le:

    CategoriaCliente_Le = Err

    Select Case Err

        Case 22364
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 28845, 28846
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CATEGORIACLIENTE1", Err, objCategoriaCliente.sCategoria)

        Case 28847
            'Não encontrou. Erro tratado na rotina chamadora.

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149083)

    End Select

    'Fechamento comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function CategoriaCliente_Le_Itens(ByVal objCategoriaCliente As ClassCategoriaCliente, ByVal colItensCategoria As Collection) As Long
'Lê na tabela de CategoriaClienteItem todos os itens de uma Categoria e os retorna na coleção colItensCategoria

Dim lErro As Long
Dim lComando As Long
Dim sItem As String
Dim sDescricaoItem As String
Dim objCategoriaClienteItem As ClassCategoriaClienteItem

On Error GoTo Erro_CategoriaCliente_Le_Itens

    'Abertura comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 28848

    sItem = String(STRING_CATEGORIACLIENTEITEM_ITEM, 0)
    sDescricaoItem = String(STRING_CATEGORIACLIENTEITEM_DESCRICAO, 0)

    'Pesquisa no BD os itens da categoria
    lErro = Comando_Executar(lComando, "SELECT Item, Descricao FROM CategoriaClienteItem WHERE Categoria = ? ", sItem, sDescricaoItem, objCategoriaCliente.sCategoria)
    If lErro <> AD_SQL_SUCESSO Then Error 28849

    'Verifica se itens da categoria existe
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 28850

    'Se não encontrou
    If lErro = AD_SQL_SEM_DADOS Then Error 28851

    'Passa todos para a collection
    Do While lErro <> AD_SQL_SEM_DADOS

        Set objCategoriaClienteItem = New ClassCategoriaClienteItem
        
        objCategoriaClienteItem.sCategoria = objCategoriaCliente.sCategoria
        objCategoriaClienteItem.sItem = sItem
        objCategoriaClienteItem.sDescricao = sDescricaoItem

        colItensCategoria.Add objCategoriaClienteItem

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 28852

    Loop

    'Fechamento comando
    Call Comando_Fechar(lComando)

    CategoriaCliente_Le_Itens = SUCESSO

    Exit Function

Erro_CategoriaCliente_Le_Itens:

    CategoriaCliente_Le_Itens = Err

    Select Case Err

        Case 28848
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 28849, 28850, 28852
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CATEGORIACLIENTEITEM", Err, objCategoriaClienteItem.sCategoria)

        Case 28851
            'Não encontrou. Erro tratado na rotina chamadora.

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149084)

    End Select

    'Fechamento comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function TipoDeCliente_Le(ByVal objTipoCliente As ClassTipoCliente) As Long
'Lê a tabela TiposDeCliente à partir do Tipo De Cliente passado como parâmetro

Dim lErro As Long
Dim lComando As Long
Dim tTipoCliente As typeTipoCliente
Dim objTipoClienteCategoria As New ClassTipoClienteCategoria

On Error GoTo Erro_TipoDeCliente_Le

    'Abertura comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 28935

    tTipoCliente.sDescricao = String(STRING_TIPO_CLIENTE_DESCRICAO, 0)
    tTipoCliente.sObservacao = String(STRING_TIPO_CLIENTE_OBS, 0)
    tTipoCliente.sContaContabil = String(STRING_CONTA, 0)

    'Carrega iTipo em objTipoClienteCategoria
    objTipoClienteCategoria.iTipoDeCliente = objTipoCliente.iCodigo

    'Lê a tabela TiposDeCliente a partir do Tipo de Cliente
    lErro = Comando_Executar(lComando, "SELECT Descricao, LimiteCredito, CondicaoPagto, Desconto, CodMensagem, TabelaPreco, Observacao, ContaContabil, Vendedor, ComissaoVendas, Regiao, FreqVisitas, CodTransportadora, CodCobrador, PadraoCobranca FROM TiposDeCliente WHERE Codigo = ?", _
        tTipoCliente.sDescricao, tTipoCliente.dLimiteCredito, tTipoCliente.iCondicaoPagto, tTipoCliente.dDesconto, tTipoCliente.iCodMensagem, tTipoCliente.iTabelaPreco, tTipoCliente.sObservacao, tTipoCliente.sContaContabil, tTipoCliente.iVendedor, tTipoCliente.dComissaoVendas, tTipoCliente.iRegiao, tTipoCliente.iFreqVisitas, tTipoCliente.iCodTransportadora, tTipoCliente.iCodCobrador, tTipoCliente.iPadraoCobranca, objTipoCliente.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 28941

    'Verifica se Tipo de Cliente existe
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 28942

    'Se não encontrou ==> erro
    If lErro = AD_SQL_SEM_DADOS Then Error 28943

    'Transfere os dados para objTipoCliente
    objTipoCliente.sDescricao = tTipoCliente.sDescricao
    objTipoCliente.dLimiteCredito = tTipoCliente.dLimiteCredito
    objTipoCliente.iCondicaoPagto = tTipoCliente.iCondicaoPagto
    objTipoCliente.dDesconto = tTipoCliente.dDesconto
    objTipoCliente.iCodMensagem = tTipoCliente.iCodMensagem
    objTipoCliente.iTabelaPreco = tTipoCliente.iTabelaPreco
    objTipoCliente.sObservacao = tTipoCliente.sObservacao
    objTipoCliente.sContaContabil = tTipoCliente.sContaContabil
    objTipoCliente.iVendedor = tTipoCliente.iVendedor
    objTipoCliente.dComissaoVendas = tTipoCliente.dComissaoVendas
    objTipoCliente.iRegiao = tTipoCliente.iRegiao
    objTipoCliente.iFreqVisitas = tTipoCliente.iFreqVisitas
    objTipoCliente.iCodTransportadora = tTipoCliente.iCodTransportadora
    objTipoCliente.iCodCobrador = tTipoCliente.iCodCobrador
    objTipoCliente.iPadraoCobranca = tTipoCliente.iPadraoCobranca

    'Fechamento comando
    Call Comando_Fechar(lComando)

    TipoDeCliente_Le = SUCESSO

    Exit Function

Erro_TipoDeCliente_Le:

    TipoDeCliente_Le = Err

    Select Case Err

        Case 28935
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 28941, 28942
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOSDECLIENTE1", Err, objTipoCliente.iCodigo)

        Case 28943
            'Não encontrou o Tipo de Cliente. Erro tratado na rotina chamadora.

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149085)

    End Select

    'Fechamento comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function TipoDeClienteCategorias_Le(ByVal objTipoCliente As ClassTipoCliente, ByVal colCategoriaItem As Collection) As Long
'Lê na tabela de TiposDeClienteCategoria, todas as categorias de um Tipo de Cliente, guardando na coleção colCategoriaItem

Dim lErro As Long
Dim iIndice As Integer
Dim lComando As Long
Dim sCategoria As String
Dim sItem As String
Dim objTipoClienteCategoria As ClassTipoClienteCategoria
On Error GoTo Erro_TipoDeClienteCategorias_Le

    'Inicializar comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 28937

    sCategoria = String(STRING_CATEGORIACLIENTE_CATEGORIA, 0)
    sItem = String(STRING_CATEGORIACLIENTEITEM_ITEM, 0)

    'Pesquisa no BD as categoria de um tipo de cliente
    lErro = Comando_Executar(lComando, "SELECT Categoria, Item FROM TipoDeClienteCategorias WHERE TipoDeCliente = ?", sCategoria, sItem, objTipoCliente.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 28938

    'Verifica se alguma categoria existe
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 28939

    'Passa para a collection
    Do While lErro <> AD_SQL_SEM_DADOS

        Set objTipoClienteCategoria = New ClassTipoClienteCategoria

        objTipoClienteCategoria.sCategoria = sCategoria
        objTipoClienteCategoria.sItem = sItem

        colCategoriaItem.Add objTipoClienteCategoria

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 28940

    Loop

    'Fechamento comando
    Call Comando_Fechar(lComando)

    TipoDeClienteCategorias_Le = SUCESSO

    Exit Function

Erro_TipoDeClienteCategorias_Le:

    TipoDeClienteCategorias_Le = Err

    Select Case Err

        Case 28937
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 28938, 28939, 28940
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPODECLIENTECATEGORIAS", Err, objTipoClienteCategoria.iTipoDeCliente)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149086)

    End Select

    'Fechamento comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function CategoriaClienteItem_Le(ByVal objCategoriaClienteItem As ClassCategoriaClienteItem) As Long
'Lê na tabela de CategoriaClienteItem a Categoria e o Item de uma deteminada Categoria de Cliente

Dim lErro As Long
Dim iIndice As Integer
Dim lComando As Long
Dim sCategoria As String
Dim sItem As String
Dim sDescricaoItem As String

On Error GoTo Erro_CategoriaClienteItem_Le

    'Inicializar comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 28988

    sCategoria = String(STRING_CATEGORIACLIENTE_CATEGORIA, 0)
    sItem = String(STRING_CATEGORIACLIENTEITEM_ITEM, 0)
    sDescricaoItem = String(STRING_CATEGORIACLIENTEITEM_DESCRICAO, 0)
    
    'Executar comando SQL
    lErro = Comando_Executar(lComando, "SELECT Descricao FROM CategoriaClienteItem WHERE Categoria = ? AND Item = ?", sDescricaoItem, objCategoriaClienteItem.sCategoria, objCategoriaClienteItem.sItem)
    If lErro <> AD_SQL_SUCESSO Then Error 28989

    'Verifica se o item da categoria existe
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 28990

    'Se não encontrou
    If lErro = AD_SQL_SEM_DADOS Then Error 28991

    objCategoriaClienteItem.sDescricao = sDescricaoItem
    
    'Libera comando
    Call Comando_Fechar(lComando)

    CategoriaClienteItem_Le = SUCESSO

    Exit Function

Erro_CategoriaClienteItem_Le:

    CategoriaClienteItem_Le = Err

    Select Case Err

        Case 28988
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 28989, 28990
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CATEGORIACLIENTEITEM1", Err, objCategoriaClienteItem.sItem, objCategoriaClienteItem.sCategoria)

        Case 28991
            'Não encontrou o item da categoria. Erro tratado na rotina chamadora.

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149087)

    End Select

    'Libera comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function TipoDocInfo_Le_Codigo(ByVal objTipoDocInfo As ClassTipoDocInfo) As Long
'objTipoDocInfo.iCodigo tem que estar preenchido
'Preenche objTipoDocInfo a partir do bd p/registro identificado por objTipoDocInfo.iCodigo
'Retorna 31415 se nao encontrar
Dim lErro As Long, lComando As Long
Dim tTipoDocInfo As typeTipoDocInfo

On Error GoTo Erro_TipoDocInfo_Le_Codigo

    Call objTipoDocInfo.Copiar(gobjCRFAT.colTiposDocInfo.Item(CStr(objTipoDocInfo.iCodigo)))
    ''Set objTipoDocInfo = gobjCRFAT.colTiposDocInfo.Item(CStr(objTipoDocInfo.iCodigo))

    TipoDocInfo_Le_Codigo = SUCESSO

    Exit Function

Erro_TipoDocInfo_Le_Codigo:

    TipoDocInfo_Le_Codigo = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149088)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function TiposDocInfo_Le(ByVal colTipoDocInfo As colTipoDocInfo) As Long
'Lê todos os registros de TiposDocInfo

Dim lErro As Long
Dim tTipoDocInfo As typeTipoDocInfo
Dim lComando As Long
Dim objTipoDocInfo As ClassTipoDocInfo

On Error GoTo Erro_TiposDocInfo_Le

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 31397

    'Inicializa as variáveis do tipo Texto
    With tTipoDocInfo
        .sSigla = String(STRING_TIPO_DOC_INFO_SIGLA, 0)
        .sDescricao = String(STRING_TIPODOCINFO_DESCRICAO, 0)
        .sNaturezaOperacaoPadrao = String(STRING_NATUREZAOP_CODIGO, 0)
        .sNomeReduzido = String(STRING_TIPO_DOC_INFO_NOMEREDUZIDO, 0)
        .sNomeTelaNFiscal = String(STRING_NOME_TELA, 0)
        .sTipoDocCPR = String(STRING_TIPO_DOC_INFO_TIPODOCCPR, 0)
        .sTituloTelaNFiscal = String(STRING_TIPO_DOC_INFO_TITULOTELANFISCAL, 0)
        .sSiglaNFOriginal = String(STRING_TIPO_DOC_INFO_SIGLANFORIGINAL, 0)
        .sNatOpExtPadrao = String(STRING_NATUREZAOP_CODIGO, 0)
        
        'Lê todos os Tipos de Documentos do BD
        lErro = Comando_Executar(lComando, "SELECT Sigla, Descricao, TipoMovtoEstoque, TipoMovtoEstoque2,TipoMovtoEstoqueConsig,TipoMovtoEstoqueConsig2, TipoMovtoEstoqueBenef,TipoMovtoEstoqueBenef2, TipoDocCPR, NaturezaOperacaoPadrao, InfoContabilizacao, Codigo, NomeReduzido, NomeTelaNFiscal, TituloTelaNFiscal, Faturavel, Comissao, Emitente, Destinatario, Origem, Complementar, Tipo, SiglaNFOriginal, Faturamento, ModeloArqICMS, Padrao, NatOpExtPadrao, NFFatura, TipoOperacaoTrib, SubTipoContabil, Rastreavel, Compras, EscaninhoRastro, ModDocFis, ModDocFisE, FinalidadeNFe, IndConsumidorFinal, IndPresenca FROM TiposDocInfo ORDER BY NomeReduzido", _
            .sSigla, .sDescricao, .iTipoMovtoEstoque, .iTipoMovtoEstoque2, .iTipoMovtoEstoqueConsig, .iTipoMovtoEstoqueConsig2, .iTipoMovtoEstoqueBenef, .iTipoMovtoEstoqueBenef2, .sTipoDocCPR, .sNaturezaOperacaoPadrao, .iInfoContabilizacao, .iCodigo, .sNomeReduzido, .sNomeTelaNFiscal, .sTituloTelaNFiscal, .iFaturavel, .iComissao, .iEmitente, .iDestinatario, .iOrigem, .iComplementar, .iTipo, .sSiglaNFOriginal, .iFaturamento, .iModeloArqICMS, .iPadrao, .sNatOpExtPadrao, .iNFFatura, .iTipoOperacaoTrib, .iSubTipoContabil, .iRastreavel, .iCompras, .iEscaninhoRastro, .iModDocFis, .iModDocFisE, .iFinalidadeNFe, .iIndConsumidorFinal, .iIndPresenca)
    End With
    If lErro <> AD_SQL_SUCESSO Then Error 31398

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 31399

    Do While lErro <> AD_SQL_SEM_DADOS

        With tTipoDocInfo
            
            'Guarda em colTipoDocInfo os Tipos de Documento lidos
            Set objTipoDocInfo = colTipoDocInfo.Add(.sSigla, .sDescricao, .iTipoMovtoEstoque, .iTipoMovtoEstoque2, .iTipoMovtoEstoqueBenef, .sNaturezaOperacaoPadrao, .iInfoContabilizacao, .sTipoDocCPR, .iCodigo, .sNomeReduzido, .sNomeTelaNFiscal, .sTituloTelaNFiscal, .iFaturavel, .iComissao, .iEmitente, .iDestinatario, .iOrigem, .iComplementar, .iTipo, .sSiglaNFOriginal, .iFaturamento, .iModeloArqICMS, .iPadrao, .sNatOpExtPadrao, .iNFFatura, .iTipoOperacaoTrib, CStr(.iCodigo))
            
            objTipoDocInfo.iTipoMovtoEstoqueBenef2 = .iTipoMovtoEstoqueBenef2
            objTipoDocInfo.iTipoMovtoEstoqueConsig = .iTipoMovtoEstoqueConsig
            objTipoDocInfo.iTipoMovtoEstoqueConsig2 = .iTipoMovtoEstoqueConsig2
            objTipoDocInfo.iSubTipoContabil = .iSubTipoContabil
            objTipoDocInfo.iRastreavel = .iRastreavel
            objTipoDocInfo.iCompras = .iCompras
            objTipoDocInfo.iEscaninhoRastro = .iEscaninhoRastro
            
            'nfe 3.10
            objTipoDocInfo.iModDocFis = .iModDocFis
            objTipoDocInfo.iModDocFisE = .iModDocFisE
            objTipoDocInfo.iFinalidadeNFe = .iFinalidadeNFe
            objTipoDocInfo.iIndConsumidorFinal = .iIndConsumidorFinal
            objTipoDocInfo.iIndPresenca = .iIndPresenca
                
        End With
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 31400

    Loop

    'Fecha o comando
    Call Comando_Fechar(lComando)

    TiposDocInfo_Le = SUCESSO

    Exit Function

Erro_TiposDocInfo_Le:

    TiposDocInfo_Le = Err

    Select Case Err

        Case 31397
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149089)

        Case 31398, 31399, 31400
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPODOCINFO1", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149090)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function TiposDocInfo_Le_NomeTelaComplemetarTipo(ByVal objTipoDocInfo As ClassTipoDocInfo) As Long
'??? ver c/a Shirley a descricao: usada em NFiscalEntrada. Nem 'Complemetar' foi escrito direito

Dim lErro As Long
Dim lComando As Long
Dim iCodigo As Integer

On Error GoTo Erro_TiposDocInfo_Le_NomeTelaComplemetarTipo

    'Abrir comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 35100

    'Lê a Nota Fiscal Original
    lErro = Comando_Executar(lComando, "SELECT Codigo FROM TiposDocInfo WHERE NomeTelaNFiscal =? AND Complementar = ? AND Tipo = ? ", iCodigo, objTipoDocInfo.sNomeTelaNFiscal, objTipoDocInfo.iComplementar, objTipoDocInfo.iTipo)
    If lErro <> AD_SQL_SUCESSO Then Error 35101

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 35102

    'Se não encontrar --> Erro
    If lErro = AD_SQL_SEM_DADOS Then Error 35103

    'Retorna o Número Interno da Nota Fiscal
    objTipoDocInfo.iCodigo = iCodigo

    'Fecha o comando
    Call Comando_Fechar(lComando)

    TiposDocInfo_Le_NomeTelaComplemetarTipo = SUCESSO

    Exit Function

Erro_TiposDocInfo_Le_NomeTelaComplemetarTipo:

    TiposDocInfo_Le_NomeTelaComplemetarTipo = Err

    Select Case Err

        Case 35100
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 35101, 35102
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPODOCINFO1", Err)

        Case 35103

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149091)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function NaturezaOperacao_Le(ByVal objNaturezaOperacao As ClassNaturezaOp) As Long
'Lê a natureza da operação a partir da chave fornecida em objNaturezaOperacao. Retorna os dados em objNaturezaOperacao

Dim lErro As Long
Dim lComando As Long
Dim sDescricao As String
Dim sDescrNF As String
Dim sNatBCCredPadrao As String, iFinalidadeEspec As Integer

On Error GoTo Erro_NaturezaOperacao_Le

    sDescricao = String(STRING_NATUREZAOP_DESCRICAO, 0)
    sDescrNF = String(STRING_NATUREZAOP_DESCRNF, 0)
    sNatBCCredPadrao = String(STRING_NATBCCRED, 0)

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 17955

    lErro = Comando_Executar(lComando, "SELECT Descricao, DescrNF, NatBCCredPadrao, FinalidadeEspec FROM NaturezaOp WHERE Codigo = ?", sDescricao, sDescrNF, sNatBCCredPadrao, iFinalidadeEspec, objNaturezaOperacao.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 17956

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 17957

    If lErro = AD_SQL_SEM_DADOS Then Error 17958

    objNaturezaOperacao.sDescricao = sDescricao
    objNaturezaOperacao.sDescrNF = sDescrNF
    objNaturezaOperacao.sNatBCCredPadrao = sNatBCCredPadrao
    objNaturezaOperacao.iFinalidadeEspec = iFinalidadeEspec

    lErro = Comando_Fechar(lComando)

    NaturezaOperacao_Le = SUCESSO

    Exit Function

Erro_NaturezaOperacao_Le:

    NaturezaOperacao_Le = Err

    Select Case Err

        Case 17955
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 17956, 17957
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NATUREZAOP", Err, objNaturezaOperacao.sCodigo)

        Case 17958

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149092)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function NaturezaOperacao_Le_Todos(ByVal colNaturezaOperacao As Collection)
'Lê todos as naturezas de operação contidas na tabela de NaturezaOp e coloca-as em colNaturezaOperacao

Dim lErro As Long
Dim lComando As Long
Dim objNaturezaOperacao As ClassNaturezaOp
Dim sCodigo As String
Dim sDescricao As String
Dim sNatBCCredPadrao As String, iFinalidadeEspec As Integer

On Error GoTo Erro_NaturezaOperacao_Le_Todos

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 17959

    sCodigo = String(STRING_NATUREZAOP_CODIGO, 0)
    sDescricao = String(STRING_NATUREZAOP_DESCRICAO, 0)
    sNatBCCredPadrao = String(STRING_NATBCCRED, 0)

    'Lê a tabela de NaturezaOp
    lErro = Comando_Executar(lComando, "SELECT Codigo, Descricao, NatBCCredPadrao, FinalidadeEspec FROM NaturezaOp ORDER BY Codigo", sCodigo, sDescricao, sNatBCCredPadrao, iFinalidadeEspec)
    If lErro <> AD_SQL_SUCESSO Then Error 17960

    'Lê a primeira natureza de operação
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 17961

    Do While lErro <> AD_SQL_SEM_DADOS

        Set objNaturezaOperacao = New ClassNaturezaOp

        objNaturezaOperacao.sCodigo = sCodigo
        objNaturezaOperacao.sDescricao = sDescricao
        objNaturezaOperacao.sNatBCCredPadrao = sNatBCCredPadrao
        objNaturezaOperacao.iFinalidadeEspec = iFinalidadeEspec

        'Coloca a natureza de operação lida na coleção
        colNaturezaOperacao.Add objNaturezaOperacao

        'Lê a próxima natureza de operação
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 17962

    Loop

    lErro = Comando_Fechar(lComando)

    NaturezaOperacao_Le_Todos = SUCESSO

    Exit Function

Erro_NaturezaOperacao_Le_Todos:

    NaturezaOperacao_Le_Todos = Err

    Select Case Err

        Case 17959
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 17960, 17961, 17962
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NATUREZAOP1", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149093)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function PedidoVenda_Le_Cliente(ByVal objCliente As ClassCliente, lPedido As Long, iFilial As Integer) As Long
'Lê nas tabelas de PedidosVenda e PedidosVendaBaixados o Cliente à partir da chave(lCodigo) fornecida em objCliente.
'Retorna sucesso se achar algum Cliente relacionado com Pedido de Venda.
'Retorna o código do Pedido em lPedido e o da FilialEmpresa em iFilial

Dim lErro As Long
Dim lComando As Long
Dim lCodigoCliente As Long

On Error GoTo Erro_PedidoVenda_Le_Cliente

    'Inicializa comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 33932

    'Lê o Cliente na tabela PedidosDeVenda e PedidosDeVendaBaixados
    If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL Then
        lCodigoCliente = objCliente.lCodigoLoja
        lErro = Comando_Executar(lComando, "SELECT FilialEmpresa, Codigo FROM PedidosDeVenda WHERE Cliente = ? UNION ALL SELECT FilialEmpresa, Codigo FROM PedidosDeVendaBaixados WHERE Cliente = ?", iFilial, lPedido, objCliente.lCodigoLoja, objCliente.lCodigoLoja)
    Else
        lCodigoCliente = objCliente.lCodigo
        lErro = Comando_Executar(lComando, "SELECT FilialEmpresa, Codigo FROM PedidosDeVenda WHERE Cliente = ? UNION ALL SELECT FilialEmpresa, Codigo FROM PedidosDeVendaBaixados WHERE Cliente = ?", iFilial, lPedido, objCliente.lCodigo, objCliente.lCodigo)
    End If
    If lErro <> AD_SQL_SUCESSO Then Error 33933

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 33934
    
    'Não achou o Cliente
    If lErro = AD_SQL_SEM_DADOS Then Error 33935

    Call Comando_Fechar(lComando)

    PedidoVenda_Le_Cliente = SUCESSO

Exit Function

Erro_PedidoVenda_Le_Cliente:

    PedidoVenda_Le_Cliente = Err

    Select Case Err

        Case 33932
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 33933, 33934
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLIENTE_PEDIDOS_DE_VENDA", Err, lCodigoCliente)

        Case 33935
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149094)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function ChequePre_Le_Cliente(ByVal objCliente As ClassCliente, lNumIntCheque As Long) As Long
'Lê na tabela de ChequePre, Cliente à partir da chave(lCodigo) fornecida em objCliente.
'Retorna sucesso se achar algum Cliente relacionado à Cheque Pré.
'Em lNumIntCheque retorna o número interno do cheque

Dim lErro As Long
Dim lComando As Long
Dim lCodigoCliente As Long

On Error GoTo Erro_ChequePre_Le_Cliente

    'Inicializa comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 33938

    'Lê o Cliente na tabela ChequePre
    If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL Then
        lCodigoCliente = objCliente.lCodigoLoja
        lErro = Comando_Executar(lComando, "SELECT NumIntCheque FROM ChequePre WHERE Cliente = ?", lNumIntCheque, objCliente.lCodigoLoja)
    Else
        lCodigoCliente = objCliente.lCodigo
        lErro = Comando_Executar(lComando, "SELECT NumIntCheque FROM ChequePre WHERE Cliente = ?", lNumIntCheque, objCliente.lCodigo)
    End If
    If lErro <> AD_SQL_SUCESSO Then Error 33939

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 33940
    
    'Não achou o Cliente
    If lErro = AD_SQL_SEM_DADOS Then Error 33941

    Call Comando_Fechar(lComando)

    ChequePre_Le_Cliente = SUCESSO

Exit Function

Erro_ChequePre_Le_Cliente:

    ChequePre_Le_Cliente = Err

    Select Case Err

        Case 33938
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 33939, 33940
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CHEQUEPRE2", Err, lCodigoCliente)

        Case 33941
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149095)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function PedidosDeVenda_Le_CondicaoPagto(ByVal objCondicaoPagto As ClassCondicaoPagto, lPedido As Long, iFilial As Integer) As Long
'Verifica se CondicaoPagto está vinculada à Pedidos de Venda ou à Pedidos de Venda Baixados
'retorna SUCESSO se estiver
Dim lErro As Long
Dim lComando As Long

On Error GoTo Erro_PedidosDeVenda_Le_CondicaoPagto

    'Inicializa comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 33986

    'Lê a CondicaoPagto na tabela PedidosDeVenda e PedidosDeVendaBaixados
    lErro = Comando_Executar(lComando, "SELECT FilialEmpresa, Codigo FROM PedidosDeVenda WHERE CondicaoPagto = ? UNION ALL SELECT FilialEmpresa, Codigo FROM PedidosDeVendaBaixados WHERE CondicaoPagto = ?", iFilial, lPedido, objCondicaoPagto.iCodigo, objCondicaoPagto.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 33987

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 33988
    
    'Não achou a CondicaoPagto
    If lErro = AD_SQL_SEM_DADOS Then Error 33989

    Call Comando_Fechar(lComando)

    PedidosDeVenda_Le_CondicaoPagto = SUCESSO

Exit Function

Erro_PedidosDeVenda_Le_CondicaoPagto:

    PedidosDeVenda_Le_CondicaoPagto = Err

    Select Case Err

        Case 33986
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 33987, 33988
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONDICAOPAGTO_PEDIDOS_DE_VENDA", Err, objCondicaoPagto.iCodigo)

        Case 33989
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149096)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function PedidoVenda_Le_FilialCliente(ByVal objFilialCliente As ClassFilialCliente, lPedido As Long, iFilial As Integer) As Long
'Retorna sucesso se achar algum pedido de venda (baixado ou nao) associado a filial do cliente

Dim lErro As Long
Dim lComando As Long

On Error GoTo Erro_PedidoVenda_Le_FilialCliente

    'Inicializa comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 33944

    'Lê a Filial do Cliente na tabela PedidosDeVenda e PedidosDeVendaBaixados
    lErro = Comando_Executar(lComando, "SELECT FilialEmpresa, Codigo FROM PedidosDeVenda WHERE Cliente = ? AND Filial = ? UNION ALL SELECT FilialEmpresa, Codigo FROM PedidosDeVendaBaixados WHERE Cliente = ? AND Filial = ?", iFilial, lPedido, objFilialCliente.lCodCliente, objFilialCliente.iCodFilial, objFilialCliente.lCodCliente, objFilialCliente.iCodFilial)
    If lErro <> AD_SQL_SUCESSO Then Error 33945

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 33946
    
    'Não achou a Filial do Cliente
    If lErro = AD_SQL_SEM_DADOS Then Error 33947

    Call Comando_Fechar(lComando)

    PedidoVenda_Le_FilialCliente = SUCESSO

Exit Function

Erro_PedidoVenda_Le_FilialCliente:

    PedidoVenda_Le_FilialCliente = Err

    Select Case Err

        Case 33944
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 33945, 33946
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIALCLIENTE_PEDIDOS_DE_VENDA", Err, objFilialCliente.lCodCliente, objFilialCliente.iCodFilial)

        Case 33947
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149097)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function TipoVendedor_Le(ByVal objTipoVendedor As ClassTipoVendedor) As Long
'Lê os dados dos Tipos de Vendedor, cujo código foi passado como parâmetro

Dim lErro As Long
Dim lComando As Long
Dim tTipoVendedor As typeTipoVendedor

On Error GoTo Erro_TipoVendedor_Le

    'Inicializa comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 16213

    tTipoVendedor.sDescricao = String(STRING_TIPO_DE_VENDEDOR_DESCRICAO, 0)

    'Pesquisa tipos de vendedor no BD
    lErro = Comando_Executar(lComando, "SELECT  Descricao, PercComissao, PercComissaoBaixa, PercComissaoEmissao, ComissaoSobreTotal, ComissaoFrete, ComissaoSeguro, ComissaoICM, ComissaoIPI FROM TiposDeVendedor WHERE Codigo = ? ", tTipoVendedor.sDescricao, tTipoVendedor.dPercComissao, tTipoVendedor.dPercComissaoBaixa, tTipoVendedor.dPercComissaoEmissao, tTipoVendedor.iComissaoSobreTotal, tTipoVendedor.iComissaoFrete, tTipoVendedor.iComissaoSeguro, tTipoVendedor.iComissaoICM, tTipoVendedor.iComissaoIPI, objTipoVendedor.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 16214

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 16215

    'Não encontrou o Tipo de Vendedor
    If lErro = AD_SQL_SEM_DADOS Then Error 16216

    objTipoVendedor.sDescricao = tTipoVendedor.sDescricao
    objTipoVendedor.dPercComissao = tTipoVendedor.dPercComissao
    objTipoVendedor.dPercComissaoBaixa = tTipoVendedor.dPercComissaoBaixa
    objTipoVendedor.dPercComissaoEmissao = tTipoVendedor.dPercComissaoEmissao
    objTipoVendedor.iComissaoSobreTotal = tTipoVendedor.iComissaoSobreTotal
    objTipoVendedor.iComissaoFrete = tTipoVendedor.iComissaoFrete
    objTipoVendedor.iComissaoSeguro = tTipoVendedor.iComissaoSeguro
    objTipoVendedor.iComissaoICM = tTipoVendedor.iComissaoICM
    objTipoVendedor.iComissaoIPI = tTipoVendedor.iComissaoIPI

    'Libera comando
    Call Comando_Fechar(lComando)

    TipoVendedor_Le = SUCESSO

    Exit Function

Erro_TipoVendedor_Le:

    TipoVendedor_Le = Err

    Select Case Err

        Case 16213
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 16214, 16215
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOSDEVENDEDOR", Err, objTipoVendedor.iCodigo)

        Case 16216
            'Erro será tratado na rotina chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149098)

    End Select

    'Libera comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Vendedor_Le_Tipo(ByVal objTipoVendedor As ClassTipoVendedor, iCodigo As Integer) As Long
'Tenta ler vendedor relacionado a Tipo de Vendedor passado como parâmetro
'Se existir devolve código do vendedor em iCodigo

Dim lErro As Long
Dim lComando As Long

On Error GoTo Erro_Vendedor_Le_Tipo

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 16247

    'Pesquisa vendedores no BD
    lErro = Comando_Executar(lComando, "SELECT Codigo FROM Vendedores WHERE Tipo = ? ", iCodigo, objTipoVendedor.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 16248

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 16249

    If lErro = AD_SQL_SEM_DADOS Then Error 16250

    Call Comando_Fechar(lComando)

    Vendedor_Le_Tipo = SUCESSO

    Exit Function

Erro_Vendedor_Le_Tipo:

    Vendedor_Le_Tipo = Err

    Select Case Err

        Case 16247
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 16248, 16249
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_VENDEDORES", Err)

        Case 16250

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149099)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Vendedor_Objetos_Relacionados_Testa(ByVal objVendedor As ClassVendedor) As Long
'Lê os objetos relacionados com Vendedor

Dim lErro As Long
Dim lCliente As Long
Dim iFilial As Integer
Dim lPedido As Long
Dim lNumInt As Long
Dim iTipoCliente As Integer
Dim iOperador As Integer

On Error GoTo Erro_Vendedor_Objetos_Relacionados_Testa

    'Verifica se Vendedor esta relacionado com Comissões
    lErro = CF("Comissao_Le_Vendedor", objVendedor, lNumInt)
    If lErro <> SUCESSO And lErro <> 16714 Then gError 16708

    If lErro = SUCESSO Then gError 16707
    
    lErro = FilialCliente_Le_Vendedor_Um(objVendedor, lCliente, iFilial)
    If lErro <> SUCESSO And lErro <> 33962 Then gError 33963
    
    'Vendedor está relacionado com Filial Cliente --> erro
    If lErro = SUCESSO Then gError 33964
    
    lErro = ComissaoPedVenda_Le_Vendedor_Um(objVendedor, lPedido)
    If lErro <> SUCESSO And lErro <> 33969 Then gError 33970
    
    'Vendedor está relacionado com ComissaoPedVenda --> erro
    If lErro = SUCESSO Then gError 33971
    
    lErro = ComissaoPedVendaBaixado_Le_Vendedor_Um(objVendedor, lPedido)
    If lErro <> SUCESSO And lErro <> 33975 Then gError 33976
    
    'Vendedor está relacionado com ComissaoPedVendaBaixado --> erro
    If lErro = SUCESSO Then gError 33977
    
    lErro = ComissaoNF_Le_Vendedor_Um(objVendedor, lNumInt)
    If lErro <> SUCESSO And lErro <> 33981 Then gError 33982
    
    'Vendedor está relacionado com ComissaoNF --> erro
    If lErro = SUCESSO Then gError 33983
    
    lErro = TipoCliente_Le_Vendedor_Um(objVendedor, iTipoCliente)
    If lErro <> SUCESSO And lErro <> 43504 Then gError 43505
    
    'Vendedor está relacionado com Tipo de Cliente --> erro
    If lErro = SUCESSO Then gError 43506
    
    'Verifica se Vendedor esta relacionado com Operador
    lErro = Operador_Le_Vendedor(objVendedor, iOperador)
    If lErro <> SUCESSO And lErro <> 126465 Then gError 126466
    
    'Vendedor está relacionado com Operador --> erro
    If lErro = SUCESSO Then gError 126467
    
    Vendedor_Objetos_Relacionados_Testa = SUCESSO

    Exit Function

Erro_Vendedor_Objetos_Relacionados_Testa:

    Vendedor_Objetos_Relacionados_Testa = gErr

    Select Case gErr

        Case 33963, 33970, 33976, 33982, 43505, 16708, 126466
            
        Case 16707
            Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_RELACIONADO_COMISSAO", gErr, objVendedor.iCodigo, lNumInt)
            
        Case 33964
            Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_REL_FILIALCLIENTE", gErr, objVendedor.iCodigo, lCliente, iFilial)

        Case 33971
            Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_REL_COMISSAOPEDVENDA", gErr, objVendedor.iCodigo, lPedido)

        Case 33977
            Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_REL_COMISSAOPEDVENDABAIXADA", gErr, objVendedor.iCodigo, lPedido)

        Case 33983
            Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_REL_COMISSAONF", gErr, objVendedor.iCodigo, lNumInt)

        Case 43506
            Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_REL_TIPOCLIENTE", gErr, objVendedor.iCodigo, iTipoCliente)

        Case 126467
            Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_REL_OPERADOR", gErr, objVendedor.iCodigo, iOperador)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149100)

    End Select

    Exit Function

End Function

Function FilialCliente_Le_Vendedor_Um(ByVal objVendedor As ClassVendedor, lCliente As Long, iFilial As Integer) As Long
'Retorna sucesso se achar alguma filial de cliente relacionada com o Vendedor

Dim lErro As Long
Dim lComando As Long

On Error GoTo Erro_FilialCliente_Le_Vendedor_Um

    'Inicializa comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 33959

    'Lê o Vendedor na tabela FilialCliente
    lErro = Comando_Executar(lComando, "SELECT CodCliente, CodFilial FROM FiliaisClientes WHERE Vendedor = ?", lCliente, iFilial, objVendedor.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 33960

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 33961

    'Não achou o Vendedor
    If lErro = AD_SQL_SEM_DADOS Then Error 33962

    Call Comando_Fechar(lComando)

    FilialCliente_Le_Vendedor_Um = SUCESSO

    Exit Function

Erro_FilialCliente_Le_Vendedor_Um:

    FilialCliente_Le_Vendedor_Um = Err

    Select Case Err

        Case 33959
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 33960, 33961
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISCLIENTES1", Err, objVendedor.iCodigo)

        Case 33962
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149101)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function ComissaoPedVenda_Le_Vendedor_Um(ByVal objVendedor As ClassVendedor, lPedido As Long) As Long
'Retorna sucesso se achar alguma Comissao de Pedido de Venda relacionada com o Vendedor.

Dim lErro As Long
Dim lComando As Long

On Error GoTo Erro_ComissaoPedVenda_Le_Vendedor_Um

    'Inicializa comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 33966

    'Lê o Vendedor na tabela ComissoesPedVendas
    lErro = Comando_Executar(lComando, "SELECT PedidoDeVendas FROM ComissoesPedVendas WHERE CodVendedor = ?", lPedido, objVendedor.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 33967

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 33968

    'Não achou o Vendedor
    If lErro = AD_SQL_SEM_DADOS Then Error 33969

    Call Comando_Fechar(lComando)

    ComissaoPedVenda_Le_Vendedor_Um = SUCESSO

    Exit Function

Erro_ComissaoPedVenda_Le_Vendedor_Um:

    ComissaoPedVenda_Le_Vendedor_Um = Err

    Select Case Err

        Case 33966
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 33967, 33968
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMISSOESPEDVENDAS1", Err, objVendedor.iCodigo)

        Case 33969
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149102)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function ComissaoPedVendaBaixado_Le_Vendedor_Um(ByVal objVendedor As ClassVendedor, lPedido As Long) As Long
'Retorna sucesso se achar alguma Comissao de Pedido de Venda Baixado relacionada com o Vendedor.

Dim lErro As Long
Dim lComando As Long

On Error GoTo Erro_ComissaoPedVendaBaixado_Le_Vendedor_Um

    'Inicializa comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 33972

    'Lê o Vendedor na tabela ComissoesPedVendasBaixados
    lErro = Comando_Executar(lComando, "SELECT PedidoDeVendas FROM ComissoesPedVendasBaixados WHERE CodVendedor = ?", lPedido, objVendedor.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 33973

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 33974

    'Não achou o Vendedor
    If lErro = AD_SQL_SEM_DADOS Then Error 33975

    Call Comando_Fechar(lComando)

    ComissaoPedVendaBaixado_Le_Vendedor_Um = SUCESSO

    Exit Function

Erro_ComissaoPedVendaBaixado_Le_Vendedor_Um:

    ComissaoPedVendaBaixado_Le_Vendedor_Um = Err

    Select Case Err

        Case 33972
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 33973, 33974
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMISSOESPEDVENDASBAIXADOS1", Err, objVendedor.iCodigo)

        Case 33975
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149103)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function ComissaoNF_Le_Vendedor_Um(ByVal objVendedor As ClassVendedor, lNumInt As Long) As Long
'Retorna sucesso se achar alguma Comissao de NF relacionada com o Vendedor.

Dim lErro As Long
Dim lComando As Long

On Error GoTo Erro_ComissaoNF_Le_Vendedor_Um

    'Inicializa comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 33978

    'Lê o Vendedor na tabela ComissoesNF
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc FROM ComissoesNF WHERE CodVendedor = ?", lNumInt, objVendedor.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 33979

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 33980

    'Não achou o Vendedor
    If lErro = AD_SQL_SEM_DADOS Then Error 33981

    Call Comando_Fechar(lComando)

    ComissaoNF_Le_Vendedor_Um = SUCESSO

    Exit Function

Erro_ComissaoNF_Le_Vendedor_Um:

    ComissaoNF_Le_Vendedor_Um = Err

    Select Case Err

        Case 33978
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 33979, 33980
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMISSOESNF1", Err, objVendedor.iCodigo)

        Case 33981
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149104)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function Comissao_Le_Vendedor(ByVal objVendedor As ClassVendedor, lNumIntCom As Long) As Long
'Verifica se existe comissao relacionada ao Vendedor passado como parametro
'se existir devolve código interno da comissao em lNumIntCom

Dim lErro As Long
Dim lComando As Long

On Error GoTo Erro_Comissao_Le_Vendedor

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 16711

    'pesquisa comissoes no BD
    lErro = Comando_Executar(lComando, "SELECT NumIntCom FROM Comissoes WHERE CodVendedor = ? ", lNumIntCom, objVendedor.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 16712

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 16713

    If lErro = AD_SQL_SEM_DADOS Then Error 16714

    lErro = Comando_Fechar(lComando)

    Comissao_Le_Vendedor = SUCESSO

    Exit Function

Erro_Comissao_Le_Vendedor:

    Comissao_Le_Vendedor = Err

    Select Case Err

        Case 16711
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 16712, 16713
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_VENDEDORES", Err)

        Case 16714

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149105)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function TipoCliente_Le_Vendedor_Um(ByVal objVendedor As ClassVendedor, iTipoCliente As Integer) As Long
'Retorna sucesso se achar algum Tipo de Cliente relacionado com Vendedor.

Dim lErro As Long
Dim lComando As Long

On Error GoTo Erro_TipoCliente_Le_Vendedor_Um

    'Inicializa comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 43501

    'Lê o Vendedor na tabela TiposDeCliente
    lErro = Comando_Executar(lComando, "SELECT Codigo FROM TiposDeCliente WHERE Vendedor = ?", iTipoCliente, objVendedor.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 43502

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 43503

    'Não achou o Vendedor
    If lErro = AD_SQL_SEM_DADOS Then Error 43504

    Call Comando_Fechar(lComando)

    TipoCliente_Le_Vendedor_Um = SUCESSO

    Exit Function

Erro_TipoCliente_Le_Vendedor_Um:

    TipoCliente_Le_Vendedor_Um = Err

    Select Case Err

        Case 43501
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 43502, 43503
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOSDECLIENTE2", Err, objVendedor.iCodigo)

        Case 43504
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149106)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function


Function TituloReceber_Le_Numero(ByVal objTituloReceber As ClassTituloReceber) As Long
'Lê o Título nao excluido, da filial corrente, com o SiglaDocumento, Numtitulo e Data de Emissao passados em objtituloReceber

Dim lErro As Long
Dim lComando As Long
Dim tTituloReceber As typeTituloReceber

On Error GoTo Erro_TituloReceber_Le_Numero
    
    'Abre comando
    lComando = Comando_Abrir
    If lComando = 0 Then Error 26431
    
    tTituloReceber.sObservacao = String(STRING_TITULO_OBSERVACAO, 0)
    tTituloReceber.sCcl = String(STRING_CCL, 0)
    tTituloReceber.sNatureza = String(STRING_NATMOVCTA_CODIGO, 0)
    
    'Seleciona TituloReceber a partir dos dados fornecidos
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Cliente, Filial, Status, Saldo, NumParcelas, Valor, ValorIRRF, ValorISS, ISSRetido, ValorINSS, INSSRetido, PercJurosDiarios, PercMulta, CondicaoPagto, Observacao, DataRegistro, PISRetido, COFINSRetido, CSLLRetido, Moeda, Ccl, Natureza, CpoGenerico1, ReajustePeriodicidade, ReajusteBase, ReajustadoAte FROM TitulosRec WHERE NumTitulo = ? AND DataEmissao = ? AND SiglaDocumento = ? AND FilialEmpresa = ? AND Status <> ? ", tTituloReceber.lNumIntDoc, tTituloReceber.lCliente, tTituloReceber.iFilial, tTituloReceber.iStatus, tTituloReceber.dSaldo, tTituloReceber.iNumParcelas, tTituloReceber.dValor, tTituloReceber.dValorIRRF, tTituloReceber.dValorISS, tTituloReceber.dISSRetido, tTituloReceber.dValorINSS, tTituloReceber.iINSSRetido, tTituloReceber.dPercJurosDiario, tTituloReceber.dPercMulta, tTituloReceber.iCondicaoPagto, tTituloReceber.sObservacao, _
        tTituloReceber.dtDataRegistro, tTituloReceber.dPISRetido, tTituloReceber.dCOFINSRetido, tTituloReceber.dCSLLRetido, tTituloReceber.iMoeda, tTituloReceber.sCcl, tTituloReceber.sNatureza, tTituloReceber.lCpoGenerico1, tTituloReceber.iReajustePeriodicidade, tTituloReceber.dtReajusteBase, tTituloReceber.dtReajustadoAte, objTituloReceber.lNumTitulo, objTituloReceber.dtDataEmissao, objTituloReceber.sSiglaDocumento, giFilialEmpresa, STATUS_EXCLUIDO)
    If lErro <> AD_SQL_SUCESSO Then Error 26432
    
    'Tenta ler
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 26433
    
    'Se não encontrou o Título à Receber --> Erro
    If lErro = AD_SQL_SEM_DADOS Then Error 26434
    
    tTituloReceber.iFilialEmpresa = giFilialEmpresa
    tTituloReceber.lNumTitulo = objTituloReceber.lNumTitulo
    tTituloReceber.dtDataEmissao = objTituloReceber.dtDataEmissao
    tTituloReceber.sSiglaDocumento = objTituloReceber.sSiglaDocumento
    
    'Preenche objTituloReceber com os dados lidos
    Call TituloReceber_MoverCamposObj(objTituloReceber, tTituloReceber)
    
    'Fecha comando
    Call Comando_Fechar(lComando)
    
    TituloReceber_Le_Numero = SUCESSO
    
    Exit Function
    
Erro_TituloReceber_Le_Numero:

    TituloReceber_Le_Numero = Err

    Select Case Err
    
        Case 26431
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 26432, 26433
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOSREC", Err, objTituloReceber.lNumTitulo, objTituloReceber.sSiglaDocumento)
        
        Case 26434
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 149107)
            
    End Select

    'fecha o Comando
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function TituloReceberBaixado_Le_Numero(ByVal objTituloReceber As ClassTituloReceber) As Long
'Lê na Tabela de Titulos Baixados o Titulo nao excluido, da filial corrente, com a SiglaDocumento, NumTitulo e Data de Emissao passados em objTituloReceber

Dim lErro As Long
Dim lComando As Long
Dim tTituloReceber As typeTituloReceber

On Error GoTo Erro_TituloReceberBaixado_Le_Numero

    'Abre comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 26435
    
    tTituloReceber.sObservacao = String(STRING_TITULO_OBSERVACAO, 0)
    
    'Seleciona TítuloReceberBaixado com os dados passados
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Cliente, Filial, Status, Saldo, NumParcelas, Valor, ValorIRRF, ValorISS, ISSRetido, ValorINSS, INSSRetido, PercJurosDiarios, PercMulta, Observacao, CondicaoPagto, DataRegistro, PISRetido, COFINSRetido, CSLLRetido FROM TitulosRecBaixados WHERE NumTitulo = ? AND DataEmissao = ? AND SiglaDocumento = ? AND FilialEmpresa = ? AND Status <> ? ", _
        tTituloReceber.lNumIntDoc, tTituloReceber.lCliente, tTituloReceber.iFilial, tTituloReceber.iStatus, tTituloReceber.dSaldo, tTituloReceber.iNumParcelas, tTituloReceber.dValor, tTituloReceber.dValorIRRF, tTituloReceber.dValorISS, tTituloReceber.dISSRetido, tTituloReceber.dValorINSS, tTituloReceber.iINSSRetido, tTituloReceber.dPercJurosDiario, tTituloReceber.dPercMulta, tTituloReceber.sObservacao, tTituloReceber.iCondicaoPagto, tTituloReceber.dtDataRegistro, tTituloReceber.dPISRetido, tTituloReceber.dCOFINSRetido, tTituloReceber.dCSLLRetido, objTituloReceber.lNumTitulo, objTituloReceber.dtDataEmissao, objTituloReceber.sSiglaDocumento, giFilialEmpresa, STATUS_EXCLUIDO)
    If lErro <> AD_SQL_SUCESSO Then Error 26436
    
    'Tenta ler
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 26437
    
    'Se não encontrou --> Erro
    If lErro = AD_SQL_SEM_DADOS Then Error 26438
    
    'Preenche objTituloReceber com os dados lidos no BD
    
    tTituloReceber.iFilialEmpresa = giFilialEmpresa
    tTituloReceber.lNumTitulo = objTituloReceber.lNumTitulo
    tTituloReceber.dtDataEmissao = objTituloReceber.dtDataEmissao
    tTituloReceber.sSiglaDocumento = objTituloReceber.sSiglaDocumento
    
    Call TituloReceber_MoverCamposObj(objTituloReceber, tTituloReceber)
    
    'Fecha comando
    Call Comando_Fechar(lComando)
    
    TituloReceberBaixado_Le_Numero = SUCESSO
    
    Exit Function

Erro_TituloReceberBaixado_Le_Numero:

    TituloReceberBaixado_Le_Numero = Err
    
    Select Case Err
    
        Case 26435
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 26436, 26437
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOSRECBAIXADOS", Err, objTituloReceber.lNumTitulo, objTituloReceber.sSiglaDocumento)
        
        Case 26438 'Título não encontrado
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149108)
            
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function ChequePre_Le(ByVal objChequePre As ClassChequePre) As Long
'Le o ChequePre a partir da chave(lNumIntCheque) fornecida em objChequePre. Devolve os dados em objChequePre

Dim lComando1 As Long
Dim lErro As Long
Dim tChequePre As typeChequePre

On Error GoTo Erro_ChequePre_Le

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 17639

    tChequePre.sAgencia = String(STRING_AGENCIA, 0)
    tChequePre.sContaCorrente = String(STRING_CONTA_CORRENTE, 0)

    'Pega os dados da tabela ChequePre
    lErro = Comando_Executar(lComando1, "SELECT FilialEmpresa, Cliente, Filial, Banco, Agencia, ContaCorrente, Numero, DataDeposito, Valor, NumBordero, SequencialBack, DataEmissao, Localizacao, DataDeposito FROM ChequePre WHERE NumIntCheque = ? ", _
        tChequePre.iFilialEmpresa, tChequePre.lCliente, tChequePre.iFilial, tChequePre.iBanco, tChequePre.sAgencia, tChequePre.sContaCorrente, tChequePre.lNumero, tChequePre.dtDataDeposito, tChequePre.dValor, tChequePre.lNumBordero, tChequePre.lSequencialBack, tChequePre.dtDataEmissao, tChequePre.iLocalizacao, tChequePre.dtDataDeposito, objChequePre.lNumIntCheque)
    If lErro <> AD_SQL_SUCESSO Then Error 17640

    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 17641

    If lErro = AD_SQL_SEM_DADOS Then Error 17642

    objChequePre.iFilialEmpresa = tChequePre.iFilialEmpresa
    objChequePre.lCliente = tChequePre.lCliente
    objChequePre.iFilial = tChequePre.iFilial
    objChequePre.iBanco = tChequePre.iBanco
    objChequePre.sAgencia = tChequePre.sAgencia
    objChequePre.sContaCorrente = tChequePre.sContaCorrente
    objChequePre.lNumero = tChequePre.lNumero
    objChequePre.dtDataDeposito = tChequePre.dtDataDeposito
    objChequePre.dValor = tChequePre.dValor
    objChequePre.lNumBordero = tChequePre.lNumBordero
    objChequePre.lSequencialBack = tChequePre.lSequencialBack
    objChequePre.dtDataEmissao = tChequePre.dtDataEmissao
    objChequePre.dtDataDeposito = tChequePre.dtDataDeposito
    objChequePre.iLocalizacao = tChequePre.iLocalizacao
    
    Call Comando_Fechar(lComando1)

    ChequePre_Le = SUCESSO

Exit Function

Erro_ChequePre_Le:

    ChequePre_Le = Err

    Select Case Err

        Case 17639
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 17640, 17641
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CHEQUEPRE", Err, objChequePre.lNumIntCheque)

        Case 17642

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149109)

    End Select

    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Function TiposDocumento_Le_TituloRec(ByVal colTipoDocumento As Collection) As Long
'Le Tipos de Documento que tem campo EmTituloRec = 1

Dim lErro As Long
Dim lComando As Long
Dim objTipoDocumento As ClassTipoDocumento
Dim tTipoDocumento As typeTipoDocumento

On Error GoTo Erro_TiposDocumento_Le_TituloRec

    'Abre o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 26034

    tTipoDocumento.sSigla = String(STRING_SIGLA_DOCUMENTO, 0)
    tTipoDocumento.sDescricao = String(STRING_DESCRICAO_CAMPO, 0)
    tTipoDocumento.sDescricaoReduzida = String(STRING_DESCRICAO_CAMPO, 0)
    
    'Seleciona os Tipos de Documentos com EmTituloRec = 1
    lErro = Comando_Executar(lComando, "SELECT Sigla, Descricao, DescricaoReduzida FROM TiposDeDocumento WHERE EmTituloRec = 1", tTipoDocumento.sSigla, tTipoDocumento.sDescricao, tTipoDocumento.sDescricaoReduzida)
    If lErro <> AD_SQL_SUCESSO Then Error 26035

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 26036

    'Guar em colTipoDocumento os Tipo de Documentos encontradas
    Do While lErro <> AD_SQL_SEM_DADOS
        
        Set objTipoDocumento = New ClassTipoDocumento
            
        objTipoDocumento.sSigla = tTipoDocumento.sSigla
        objTipoDocumento.sDescricao = tTipoDocumento.sDescricao
        objTipoDocumento.sDescricaoReduzida = tTipoDocumento.sDescricaoReduzida
        
        colTipoDocumento.Add objTipoDocumento

        'Busca a próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 26037

    Loop

    'Fecha o comando
    Call Comando_Fechar(lComando)

    TiposDocumento_Le_TituloRec = SUCESSO

    Exit Function

Erro_TiposDocumento_Le_TituloRec:

    TiposDocumento_Le_TituloRec = Err

    Select Case Err

        Case 26034
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 26035, 26036, 26037
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPODOCUMENTO1", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149110)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function TituloReceber_Le(ByVal objTituloReceber As ClassTituloReceber, Optional ByVal iExcluido As Integer = 0) As Long
'Lê Titulo a Receber com o Número Interno passado como parâmetro
'No Parametro Optional: Se iExcluido = 1 então não entra os Titulos Excluidos

Dim lErro As Long
Dim lComando As Long
Dim tTituloReceber As typeTituloReceber

On Error GoTo Erro_TituloReceber_Le

    'Abre comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 26058

    tTituloReceber.sSiglaDocumento = String(STRING_SIGLA_DOCUMENTO, 0)
    tTituloReceber.sObservacao = String(STRING_TITULO_OBSERVACAO, 0)
    tTituloReceber.sCcl = String(STRING_CCL, 0)
    tTituloReceber.sNatureza = String(STRING_NATMOVCTA_CODIGO, 0)
    
    If iExcluido = 0 Then
        'Seleciona o TituloReceber a partir do NumIntDoc sem Excluidos
        lErro = Comando_Executar(lComando, "SELECT FilialEmpresa, Cliente, Filial, SiglaDocumento, NumTitulo, DataEmissao, Status, Saldo, NumParcelas, Valor, ValorIRRF, ValorISS, ISSRetido, ValorINSS, INSSRetido, PercJurosDiarios, PercMulta, Observacao, CondicaoPagto, DataRegistro, PISRetido, COFINSRetido, CSLLRetido, Moeda, Ccl, Natureza, CpoGenerico1, ReajustePeriodicidade, ReajusteBase, ReajustadoAte FROM TitulosRec WHERE NumIntDoc = ? AND Status <> ?", tTituloReceber.iFilialEmpresa, tTituloReceber.lCliente, tTituloReceber.iFilial, tTituloReceber.sSiglaDocumento, tTituloReceber.lNumTitulo, tTituloReceber.dtDataEmissao, tTituloReceber.iStatus, tTituloReceber.dSaldo, tTituloReceber.iNumParcelas, tTituloReceber.dValor, tTituloReceber.dValorIRRF, tTituloReceber.dValorISS, tTituloReceber.dISSRetido, tTituloReceber.dValorINSS, tTituloReceber.iINSSRetido, tTituloReceber.dPercJurosDiario, tTituloReceber.dPercMulta, tTituloReceber.sObservacao, tTituloReceber.iCondicaoPagto, _
            tTituloReceber.dtDataRegistro, tTituloReceber.dPISRetido, tTituloReceber.dCOFINSRetido, tTituloReceber.dCSLLRetido, tTituloReceber.iMoeda, tTituloReceber.sCcl, tTituloReceber.sNatureza, tTituloReceber.lCpoGenerico1, tTituloReceber.iReajustePeriodicidade, tTituloReceber.dtReajusteBase, tTituloReceber.dtReajustadoAte, objTituloReceber.lNumIntDoc, STATUS_EXCLUIDO)
    ElseIf iExcluido = 1 Then
        'Seleciona o TituloReceber a partir do NumIntDoc com Excluidos
        lErro = Comando_Executar(lComando, "SELECT FilialEmpresa, Cliente, Filial, SiglaDocumento, NumTitulo, DataEmissao, Status, Saldo, NumParcelas, Valor, ValorIRRF, ValorISS, ISSRetido, ValorINSS, INSSRetido, PercJurosDiarios, PercMulta, Observacao, CondicaoPagto, DataRegistro, PISRetido, COFINSRetido, CSLLRetido, Moeda, Ccl, Natureza, CpoGenerico1, ReajustePeriodicidade, ReajusteBase, ReajustadoAte FROM TitulosRec WHERE NumIntDoc = ? ", tTituloReceber.iFilialEmpresa, tTituloReceber.lCliente, tTituloReceber.iFilial, tTituloReceber.sSiglaDocumento, tTituloReceber.lNumTitulo, tTituloReceber.dtDataEmissao, tTituloReceber.iStatus, tTituloReceber.dSaldo, tTituloReceber.iNumParcelas, tTituloReceber.dValor, tTituloReceber.dValorIRRF, tTituloReceber.dValorISS, tTituloReceber.dISSRetido, tTituloReceber.dValorINSS, tTituloReceber.iINSSRetido, tTituloReceber.dPercJurosDiario, tTituloReceber.dPercMulta, tTituloReceber.sObservacao, tTituloReceber.iCondicaoPagto, _
            tTituloReceber.dtDataRegistro, tTituloReceber.dPISRetido, tTituloReceber.dCOFINSRetido, tTituloReceber.dCSLLRetido, tTituloReceber.iMoeda, tTituloReceber.sCcl, tTituloReceber.sNatureza, tTituloReceber.lCpoGenerico1, tTituloReceber.iReajustePeriodicidade, tTituloReceber.dtReajusteBase, tTituloReceber.dtReajustadoAte, objTituloReceber.lNumIntDoc)
    End If
    
    If lErro <> AD_SQL_SUCESSO Then Error 26059

    'Tenta ler
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 26060

    'Se não encontrou --> Erro
    If lErro = AD_SQL_SEM_DADOS Then Error 26061

    tTituloReceber.lNumIntDoc = objTituloReceber.lNumIntDoc
    
    'Preenche objTituloReceber com os dados do Título
    Call TituloReceber_MoverCamposObj(objTituloReceber, tTituloReceber)
    
    'Fecha o comando
    Call Comando_Fechar(lComando)

    TituloReceber_Le = SUCESSO

    Exit Function

Erro_TituloReceber_Le:

    TituloReceber_Le = Err

    Select Case Err

        Case 26058
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 26059, 26060
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOSREC1", Err, objTituloReceber.lNumIntDoc)

        Case 26061  'Título não encontrado

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149111)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ParcelasReceber_Le(ByVal objTituloReceber As ClassTituloReceber, ByVal colParcelasRec As colParcelaReceber) As Long
'Lê as Parcelas a Receber do Titulo cujo NumIntDoc foi passado em objTituloReceber

Dim lErro As Long
Dim lComando As Long
Dim tParcelaReceber As typeParcelaReceber
Dim objParcelaReceber As ClassParcelaReceber

On Error GoTo Erro_ParcelasReceber_Le

    'Abre comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 26457

    tParcelaReceber.sNumTitCobrador = String(STRING_NOSSO_NUMERO, 0)

    '###############################################
    'INSERIDO POR WAGNER
    tParcelaReceber.sObservacao = String(STRING_MAX_NOME_ARQUIVO, 0)
    '###############################################

    tParcelaReceber.sAgenciaCheque = String(STRING_AGENCIA, 0)
    tParcelaReceber.sAprovacaoCartao = String(STRING_APROVACAO_CARTAO, 0)
    tParcelaReceber.sContaCorrenteCheque = String(STRING_CONTA_CORRENTE, 0)
    tParcelaReceber.sNumeroCartao = String(STRING_NUMERO_CARTAO, 0)

    'Seleciona Parcelas vinculadas ao Título
    '#################################################
    'ALTERADO POR WAGNER
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumParcela, Status, DataVencimento, DataVencimentoReal, Saldo, Valor, Cobrador, CarteiraCobranca, NumTitCobrador,ProxSeqBaixa, ProxSeqOcorr, Desconto1Codigo, Desconto1Ate, Desconto1Valor, Desconto2Codigo, Desconto2Ate, Desconto2Valor, Desconto3Codigo, Desconto3Ate, Desconto3Valor, Aceite, Descontada, Previsao, Observacao, ValorOriginal, TipoPagto, CodConta, DataCredito, DataEmissaoCheque, BancoCheque, AgenciaCheque, ContaCorrenteCheque, NumeroCheque, DataDepositoCheque, AdmMeioPagto, Parcelamento, NumeroCartao, ValidadeCartao, AprovacaoCartao, DataTransacaoCartao FROM ParcelasRec WHERE NumIntTitulo = ? AND Status <> ? ORDER BY NumParcela", _
    tParcelaReceber.lNumIntDoc, tParcelaReceber.iNumParcela, tParcelaReceber.iStatus, tParcelaReceber.dtDataVencimento, tParcelaReceber.dtDataVencimentoReal, tParcelaReceber.dSaldo, tParcelaReceber.dValor, tParcelaReceber.iCobrador, tParcelaReceber.iCarteiraCobranca, tParcelaReceber.sNumTitCobrador, tParcelaReceber.iProxSeqBaixa, tParcelaReceber.iProxSeqOcorr, tParcelaReceber.iDesconto1Codigo, tParcelaReceber.dtDesconto1Ate, tParcelaReceber.dDesconto1Valor, _
    tParcelaReceber.iDesconto2Codigo, tParcelaReceber.dtDesconto2Ate, tParcelaReceber.dDesconto2Valor, tParcelaReceber.iDesconto3Codigo, tParcelaReceber.dtDesconto3Ate, tParcelaReceber.dDesconto3Valor, tParcelaReceber.iAceite, tParcelaReceber.iDescontada, tParcelaReceber.iPrevisao, tParcelaReceber.sObservacao, tParcelaReceber.dValorOriginal, _
    tParcelaReceber.iTipoPagto, tParcelaReceber.iCodConta, tParcelaReceber.dtDataCredito, tParcelaReceber.dtDataEmissaoCheque, tParcelaReceber.iBancoCheque, tParcelaReceber.sAgenciaCheque, tParcelaReceber.sContaCorrenteCheque, tParcelaReceber.lNumeroCheque, tParcelaReceber.dtDataDepositoCheque, tParcelaReceber.iAdmMeioPagto, tParcelaReceber.iParcelamento, tParcelaReceber.sNumeroCartao, tParcelaReceber.dtValidadeCartao, tParcelaReceber.sAprovacaoCartao, tParcelaReceber.dtDataTransacaoCartao, objTituloReceber.lNumIntDoc, STATUS_EXCLUIDO)
    If lErro <> AD_SQL_SUCESSO Then Error 26458
    '#################################################

    'Tenta ler a primeira Parcela
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 26459
    
    'Se não encontrou Parcela --> erro
    If lErro = AD_SQL_SEM_DADOS Then Error 26460

    'Loop de adição de Parcelas à coleção
    Do While lErro <> AD_SQL_SEM_DADOS

        'Adiciona Parcela na coleção
        With tParcelaReceber
            '#################################################
            'ALTERADO POR WAGNER
            Set objParcelaReceber = colParcelasRec.Add(.lNumIntDoc, objTituloReceber.lNumIntDoc, .iNumParcela, .iStatus, .dtDataVencimento, .dtDataVencimentoReal, .dSaldo, .dValor, .iProxSeqBaixa, .iCarteiraCobranca, .iCobrador, .sNumTitCobrador, 0, 0, 0, 0, 0, 0, .iDesconto1Codigo, .dtDesconto1Ate, .dDesconto1Valor, .iDesconto2Codigo, .dtDesconto2Ate, .dDesconto2Valor, .iDesconto3Codigo, .dtDesconto3Ate, .dDesconto3Valor, 0, 0, .iDescontada, .iProxSeqOcorr, .iPrevisao, .sObservacao, .dValorOriginal)
            '#################################################
            
            objParcelaReceber.iTipoPagto = .iTipoPagto
            objParcelaReceber.iCodConta = .iCodConta
            objParcelaReceber.dtDataCredito = .dtDataCredito
            objParcelaReceber.dtDataEmissaoCheque = .dtDataEmissaoCheque
            objParcelaReceber.iBancoCheque = .iBancoCheque
            objParcelaReceber.sAgenciaCheque = .sAgenciaCheque
            objParcelaReceber.sContaCorrenteCheque = .sContaCorrenteCheque
            objParcelaReceber.lNumeroCheque = .lNumeroCheque
            objParcelaReceber.dtDataDepositoCheque = .dtDataDepositoCheque
            objParcelaReceber.iAdmMeioPagto = .iAdmMeioPagto
            objParcelaReceber.iParcelamento = .iParcelamento
            objParcelaReceber.sNumeroCartao = .sNumeroCartao
            objParcelaReceber.dtValidadeCartao = .dtValidadeCartao
            objParcelaReceber.sAprovacaoCartao = .sAprovacaoCartao
            objParcelaReceber.dtDataTransacaoCartao = .dtDataTransacaoCartao
            
        End With

        'Vai para a próxima Parcela
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 26461

    Loop

    'Fecha comando
    Call Comando_Fechar(lComando)

    ParcelasReceber_Le = SUCESSO

    Exit Function

Erro_ParcelasReceber_Le:

    ParcelasReceber_Le = Err

    Select Case Err

        Case 26457
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 26458, 26459, 26461
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PARCELASREC", Err, objTituloReceber.lNumIntDoc)

        Case 26460
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TITULORECEBER_SEM_PARCELAS", Err, objTituloReceber.lNumIntDoc)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149112)

    End Select

    'Fecha o comanodo
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Comissoes_Le(ByVal lNumIntDoc As Long, ByVal colComissao As colComissao, ByVal iTipoTitulo As Integer) As Long
'Le as Comissões do Documento com o Tipo e o Número Interno passados

Dim lErro As Long
Dim lComando As Long
Dim tInfoComissao As typeInfoComissao

On Error GoTo Erro_Comissoes_Le

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 26157

    'Le as comissoes com os dados passados
    lErro = Comando_Executar(lComando, "SELECT NumIntCom, Status, CodVendedor, DataBaixa, Percentual, ValorBase, Valor, DataGeracao FROM Comissoes WHERE TipoTitulo = ? AND NumIntDoc = ? AND FilialEmpresa = ? ORDER BY CodVendedor", tInfoComissao.lNumIntCom, tInfoComissao.iStatus, tInfoComissao.iCodVendedor, tInfoComissao.dtDataBaixa, tInfoComissao.dPercentual, tInfoComissao.dValorBase, tInfoComissao.dValor, tInfoComissao.dtDataGeracao, iTipoTitulo, lNumIntDoc, giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then Error 26158

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 26159

    'Guarda em colComissao as Comissoes encontradas
    Do While lErro = AD_SQL_SUCESSO

        With tInfoComissao
            colComissao.Add .lNumIntCom, giFilialEmpresa, .iStatus, iTipoTitulo, lNumIntDoc, .iCodVendedor, .dtDataBaixa, .dPercentual, .dValorBase, .dValor, .dtDataGeracao
        End With

        'Busca a próxima
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 26160

    Loop

    'Fecha o Comando
    Call Comando_Fechar(lComando)

    Comissoes_Le = SUCESSO

    Exit Function

Erro_Comissoes_Le:

    Comissoes_Le = Err

    Select Case Err

        Case 26157
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 26158, 26459, 26160
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMISSOES", Err, lNumIntDoc)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149113)

    End Select

    'Fecha o Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function DescontoPadrao_Le(ByVal colDescontoPadrao As Collection) As Long
'Le os atributos do Desconto Padrão em CPRConfig

Dim lErro As Long
Dim lComando As Long
Dim sConteudo As String
Dim iIndice As Integer
Dim objDescontoPadrao As ClassDescontoPadrao

On Error GoTo Erro_DescontoPadrao_Le

    'Abre o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 26165
    
    'Lê em CPRConfig os atributos do Desconto Padrão
    
    For iIndice = 1 To 3
        
        Set objDescontoPadrao = New ClassDescontoPadrao
        
        sConteudo = String(STRING_CPRCONFIG_CONTEUDO, 0)
        
        'Lê o Código
        lErro = Comando_Executar(lComando, "SELECT Conteudo FROM CRFATConfig WHERE Codigo = ? ", sConteudo, "DESCONTO_PADRAO_CODIGO" & CStr(iIndice))
        If lErro <> AD_SQL_SUCESSO Then Error 26166
        
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 26167
        
        'Guarda o código em objDescontoPadrao
        objDescontoPadrao.iCodigo = CInt(sConteudo)
        
        sConteudo = String(STRING_CPRCONFIG_CONTEUDO, 0)
        
        'Lê os dias
        lErro = Comando_Executar(lComando, "SELECT Conteudo FROM CRFATConfig WHERE Codigo = ? ", sConteudo, "DESCONTO_PADRAO_DIAS" & CStr(iIndice))
        If lErro <> AD_SQL_SUCESSO Then Error 26168
        
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 26169
        
        'Guarda em objDescontoPadrão
        objDescontoPadrao.iDias = CInt(sConteudo)
        
        sConteudo = String(STRING_CPRCONFIG_CONTEUDO, 0)
        
        'Lê o Percentual
        lErro = Comando_Executar(lComando, "SELECT Conteudo FROM CRFATConfig WHERE Codigo = ? ", sConteudo, "DESCONTO_PADRAO_PERCENTUAL" & CStr(iIndice))
        If lErro <> AD_SQL_SUCESSO Then Error 26170
        
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 26384
        
        'Guarda em objDescontoPadrao
        objDescontoPadrao.dPercentual = CDbl(sConteudo)
    
        colDescontoPadrao.Add objDescontoPadrao
    
    Next
    
    'Fecha o Comando
    Call Comando_Fechar(lComando)
    
    DescontoPadrao_Le = SUCESSO
    
    Exit Function
    
Erro_DescontoPadrao_Le:

    DescontoPadrao_Le = Err
    
    Select Case Err
    
        Case 26165
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 26166, 26167
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CPRCONFIG", Err, "DESCONTO_PADRAO_CODIGO")
    
        Case 26168, 26169
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CPRCONFIG", Err, "DESCONTO_PADRAO_DIAS")
    
        Case 26170, 26384
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CPRCONFIG", Err, "DESCONTO_PADRAO_PERCENTUAL")
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149114)
    
    End Select
    
    'Fecha o Comando
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function CarteiraCobrador_Le(ByVal objCarteiraCobrador As ClassCarteiraCobrador) As Long
'Lê a Carteira a partir da chave fornecida em objCarteiraCobrador
'e devolve dados em objCarteiraCobrador

Dim lComando As Long
Dim lErro As Long
Dim tCarteira As typeCarteiraCobrador

On Error GoTo Erro_CarteiraCobrador_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 23548

    tCarteira.sContaContabil = String(STRING_CONTA, 0)
    tCarteira.sContaDuplDescontadas = String(STRING_CONTA, 0)
    tCarteira.sFaixaNossoNumeroInicial = String(STRING_FAIXA_NUM_COB, 0)
    tCarteira.sFaixaNossoNumeroFinal = String(STRING_FAIXA_NUM_COB, 0)
    tCarteira.sFaixaNossoNumeroProx = String(STRING_FAIXA_NUM_COB, 0)
    tCarteira.sNomeNoBanco = String(STRING_NOMENOBANCO, 0)
    tCarteira.sCodCarteiraNoBanco = String(STRING_CODCARTNOBANCO, 0)

    'Seleciona a carteira cujo Código está no Obj e coloca os dados no type
    lErro = Comando_Executar(lComando, "SELECT CodCarteiraCobranca, Desativada, ContaContabil, ContaDuplDescontadas, DiasDeRetencao,TaxaCobranca , TaxaDesconto, FaixaNossoNumeroInicial, FaixaNossoNumeroFinal, FaixaNossoNumeroProx, NomeNoBanco,CodCarteiraNoBanco, NumCarteiraNoBanco, QuantidadeAtual , QuantidadeAtualBanco, SaldoAtual, SaldoAtualBanco,ImprimeBoleta,ComRegistro,GeraNossoNumero FROM CarteirasCobrador WHERE Cobrador = ? AND CodCarteiraCobranca = ?", _
    tCarteira.iCodCarteiraCobranca, tCarteira.iDesativada, tCarteira.sContaContabil, tCarteira.sContaDuplDescontadas, tCarteira.iDiasDeRetencao, tCarteira.dTaxaCobranca, tCarteira.dTaxaDesconto, tCarteira.sFaixaNossoNumeroInicial, tCarteira.sFaixaNossoNumeroFinal, tCarteira.sFaixaNossoNumeroProx, tCarteira.sNomeNoBanco, tCarteira.sCodCarteiraNoBanco, tCarteira.iNumCarteiraNoBanco, tCarteira.lQuantidadeAtual, tCarteira.lQuantidadeAtualBanco, tCarteira.dSaldoAtual, tCarteira.dSaldoAtualBanco, tCarteira.iImprimeBoleta, tCarteira.iComRegistro, tCarteira.iGeraNossoNumero, objCarteiraCobrador.iCobrador, objCarteiraCobrador.iCodCarteiraCobranca)
    If lErro <> AD_SQL_SUCESSO Then Error 23549

    'Posiciona na primeira Carteira das que foram lidas
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 23550

    'Se não encontrou nenhuma Carteira com o Código em questão
    If lErro = AD_SQL_SEM_DADOS Then Error 23551
    
    'Coloca os dados da Carteira em objCarteiraCobrador
    Call CarteiraCobrador_MoverCamposObj(objCarteiraCobrador.iCobrador, objCarteiraCobrador, tCarteira)
        
    Call Comando_Fechar(lComando)

    CarteiraCobrador_Le = SUCESSO

    Exit Function

Erro_CarteiraCobrador_Le:

    CarteiraCobrador_Le = Err

    Select Case Err

        Case 23548
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 23549, 23550
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA_CARTEIRASCOBRADOR", Err, objCarteiraCobrador.iCobrador)

        Case 23551 'carteira nao encontrada
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149115)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Clientes_Le_PadraoCobranca(ByVal iCodigo As Integer) As Long
'verifica se existem clientes relacionados ao codigo do Padrao de Cobranca passado como parametro
'retorna SUCESSO se existir e 40616 se nao houver

Dim lErro As Long
Dim lComando As Long
Dim iCodigo2 As Integer

On Error GoTo Erro_Clientes_Le_PadraoCobranca

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 40613

    'pesquisa  Clientes no BD
    lErro = Comando_Executar(lComando, "SELECT CodPadraoCobranca FROM Clientes WHERE CodPadraoCobranca = ? ", iCodigo2, iCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 40614

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 40615

    If lErro = AD_SQL_SEM_DADOS Then Error 40616 'Nao existe PadraoCobranca relacionado com  Clientes

    lErro = Comando_Fechar(lComando)

    Clientes_Le_PadraoCobranca = SUCESSO

    Exit Function

Erro_Clientes_Le_PadraoCobranca:

    Clientes_Le_PadraoCobranca = Err

    Select Case Err

        Case 40613
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 40614, 40615
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLIENTES", Err)

        Case 40616

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149116)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Cliente_ObterDadosBasicos1(ByVal lComando As Long, ByVal objCliente As ClassCliente) As Long
'le o cliente a partir da chave(lCodigo) fornecida em objCliente.
'Devolve os dados em objCliente
'nao acessa dados da matriz do cliente
'É CHAMADA EM TRANSACAO PARA ALTERACAO DE DADOS DO CLIENTE

Dim lErro As Long
Dim tCliente As typeCliente
Dim lCodigoCliente As Long

On Error GoTo Erro_Cliente_ObterDadosBasicos1

    tCliente.sRazaoSocial = String(STRING_CLIENTE_RAZAO_SOCIAL, 0)
    tCliente.sNomeReduzido = String(STRING_CLIENTE_NOME_REDUZIDO, 0)
    tCliente.sObservacao = String(STRING_CLIENTE_OBSERVACAO, 0)
    tCliente.sNome = String(STRING_FILIAL_CLIENTE_NOME, 0)
    tCliente.sCgc = String(STRING_CGC, 0)
    tCliente.sInscricaoEstadual = String(STRING_INSCR_EST, 0)
    tCliente.sInscricaoMunicipal = String(STRING_INSCR_MUN, 0)
    tCliente.sObservacao2 = String(STRING_CLIENTE_OBSERVACAO, 0)
    tCliente.sContaContabil = String(STRING_CONTA, 0)

    If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL Then
        
        lCodigoCliente = objCliente.lCodigoLoja
        
        'Pega os dados da tabela Clientes
        lErro = Comando_ExecutarPos(lComando, "SELECT RazaoSocial, NomeReduzido, Tipo, Observacao, LimiteCredito, Bloqueado, CondicaoPagto, Desconto, CodMensagem, TabelaPreco, ProxCodFilial FROM Clientes WHERE CodigoLoja = ?", 0, _
        tCliente.sRazaoSocial, tCliente.sNomeReduzido, tCliente.iTipo, tCliente.sObservacao, tCliente.dLimiteCredito, tCliente.iBloqueado, tCliente.iCondicaoPagto, tCliente.dDesconto, tCliente.iCodMensagem, tCliente.iTabelaPreco, tCliente.iProxCodFilial, objCliente.lCodigoLoja)
    
    Else
    
        lCodigoCliente = objCliente.lCodigo
        
        'Pega os dados da tabela Clientes
        lErro = Comando_ExecutarPos(lComando, "SELECT RazaoSocial, NomeReduzido, Tipo, Observacao, LimiteCredito, Bloqueado, CondicaoPagto, Desconto, CodMensagem, TabelaPreco, ProxCodFilial FROM Clientes WHERE Codigo = ?", 0, _
        tCliente.sRazaoSocial, tCliente.sNomeReduzido, tCliente.iTipo, tCliente.sObservacao, tCliente.dLimiteCredito, tCliente.iBloqueado, tCliente.iCondicaoPagto, tCliente.dDesconto, tCliente.iCodMensagem, tCliente.iTabelaPreco, tCliente.iProxCodFilial, objCliente.lCodigo)
        
    End If
    
    If lErro <> AD_SQL_SUCESSO Then Error 36981

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 36982

    If lErro = AD_SQL_SEM_DADOS Then Error 36983

    objCliente.sRazaoSocial = tCliente.sRazaoSocial
    objCliente.sNomeReduzido = tCliente.sNomeReduzido
    objCliente.iTipo = tCliente.iTipo
    objCliente.sObservacao = tCliente.sObservacao
    objCliente.dLimiteCredito = tCliente.dLimiteCredito
    objCliente.iCondicaoPagto = tCliente.iCondicaoPagto
    objCliente.dDesconto = tCliente.dDesconto
    objCliente.iCodMensagem = tCliente.iCodMensagem
    objCliente.iTabelaPreco = tCliente.iTabelaPreco
    objCliente.iProxCodFilial = tCliente.iProxCodFilial
    objCliente.iBloqueado = tCliente.iBloqueado

    Cliente_ObterDadosBasicos1 = SUCESSO

    Exit Function

Erro_Cliente_ObterDadosBasicos1:

    Cliente_ObterDadosBasicos1 = Err

    Select Case Err

        Case 36981, 36982
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLIENTES", Err)

        Case 36983
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", Err, lCodigoCliente)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149117)

    End Select

    Exit Function

End Function

Function NFiscal_Le_Num_Forn(ByVal objNFiscal As ClassNFiscal, ByVal iTipo As Integer) As Long
'Lê a Nota Fiscal a partir de TipoNFiscal, Fornecedor, Filial, Serie e NFiscal.

Dim lErro As Long
Dim lComando As Long
Dim tNFiscal As typeNFiscal

On Error GoTo Erro_NFiscal_Le_Num_Forn

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 31464

    With tNFiscal

        .sMensagemNota = String(STRING_NFISCAL_MENSAGEM, 0)
        .sMensagemCorpoNota = String(STRING_NFISCAL_MENSAGEM, 0)
        .sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
        .sPlaca = String(STRING_NFISCAL_PLACA, 0)
        .sPlacaUF = String(STRING_NFISCAL_PLACA_UF, 0)
        .sSerie = String(STRING_SERIE, 0)
        .sNumPedidoTerc = String(STRING_NUM_PEDIDO_TERC, 0)
        .sVolumeNumero = String(STRING_NFISCAL_VOLUME_NUMERO, 0)
        
        'Removido por Luiz Nogueira em 21/08/03
        '.lVolumeEspecie = String(STRING_NFISCAL_VOLUME_ESPECIE, 0)
        '.lVolumeMarca = String(STRING_NFISCAL_VOLUME_MARCA, 0)
        
        'Pesquisa a Nota Fiscal com os dados passados
        lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Status, Cliente, FilialCli, FilialEntrega, DataEmissao, DataEntrada, DataSaida, DataVencimento, NumPedidoVenda, NumPedidoTerc, ClasseDocCPR, NumIntDocCPR, ValorTotal, ValorProdutos, ValorFrete, ValorSeguro, ValorOutrasDespesas, ValorDesconto, CodTransportadora, MensagemNota, TabelaPreco , NaturezaOp, PesoLiq, PesoBruto, NumIntTrib, Placa, PlacaUF, VolumeQuant, VolumeEspecie, VolumeMarca, Canal, NumIntNotaOriginal, VolumeNumero, FreteRespons, CodTranspRedesp, DetPagFrete, Moeda, TaxaMoeda, VolumeTotal, SemDataSaida, NaoImpCobranca, MensagemCorpoNota FROM TipoDocInfoNFiscal WHERE FilialEmpresa = ? AND NumNotaFiscal = ? AND Tipo = ? AND Complementar <> ? AND Fornecedor = ? AND FilialForn = ? AND Serie = ? AND Status <> ? ORDER BY DataEntrada DESC", .lNumIntDoc, .iStatus, .lCliente, .iFilialCli, .iFilialEntrega, .dtDataEmissao, .dtDataEntrada, .dtDataSaida, .dtDataVencimento, _
            .lNumPedidoVenda, .sNumPedidoTerc, .iClasseDocCPR, .lNumIntDocCPR, .dValorTotal, .dValorProdutos, .dValorFrete, .dValorSeguro, .dValorOutrasDespesas, .dValorDesconto, .iCodTransportadora, .sMensagemNota, .iTabelaPreco, .sNaturezaOp, .dPesoLiq, .dPesoBruto, .lNumIntTrib, .sPlaca, .sPlacaUF, .lVolumeQuant, .lVolumeEspecie, .lVolumeMarca, .iCanal, .lNumIntNotaOriginal, .sVolumeNumero, .iFreteRespons, .iCodTranspRedesp, .iDetPagFrete, .iMoeda, .dTaxaMoeda, .dVolumeTotal, .iSemDataSaida, .iNaoImpCobranca, .sMensagemCorpoNota, giFilialEmpresa, objNFiscal.lNumNotaFiscal, iTipo, DOCINFO_COMPLEMENTO, objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.sSerie, STATUS_EXCLUIDO)
        
    End With
        
    If lErro <> AD_SQL_SUCESSO Then Error 31465

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 31466
    
    'Se não encontrar --> Erro
    If lErro = AD_SQL_SEM_DADOS Then Error 31467

    With tNFiscal
    
        'Preenche objNFiscal com os dados lidos do BD
        objNFiscal.lNumIntDoc = .lNumIntDoc
        objNFiscal.dPesoBruto = .dPesoBruto
        objNFiscal.dPesoLiq = .dPesoLiq
        objNFiscal.dtDataEmissao = .dtDataEmissao
        objNFiscal.dtDataEntrada = .dtDataEntrada
        objNFiscal.dtDataSaida = .dtDataSaida
        objNFiscal.dtDataVencimento = .dtDataVencimento
        objNFiscal.dValorDesconto = .dValorDesconto
        objNFiscal.dValorFrete = .dValorFrete
        objNFiscal.dValorOutrasDespesas = .dValorOutrasDespesas
        objNFiscal.dValorProdutos = .dValorProdutos
        objNFiscal.dValorSeguro = .dValorSeguro
        objNFiscal.dValorTotal = .dValorTotal
        objNFiscal.iCanal = .iCanal
        objNFiscal.iClasseDocCPR = .iClasseDocCPR
        objNFiscal.iCodTransportadora = .iCodTransportadora
        objNFiscal.iCodTranspRedesp = .iCodTranspRedesp
        objNFiscal.iDetPagFrete = .iDetPagFrete
        objNFiscal.iFilialCli = .iFilialCli
        objNFiscal.iFilialEmpresa = giFilialEmpresa
        objNFiscal.iFilialEntrega = .iFilialEntrega
        objNFiscal.iStatus = .iStatus
        objNFiscal.iTabelaPreco = .iTabelaPreco
        objNFiscal.lVolumeQuant = .lVolumeQuant
        objNFiscal.lCliente = .lCliente
        objNFiscal.lNumIntDocCPR = .lNumIntDocCPR
        objNFiscal.lNumIntNotaOriginal = .lNumIntNotaOriginal
        objNFiscal.lNumIntTrib = .lNumIntTrib
        objNFiscal.sNumPedidoTerc = .sNumPedidoTerc
        objNFiscal.lNumPedidoVenda = .lNumPedidoVenda
        objNFiscal.sMensagemNota = .sMensagemNota
        objNFiscal.sNaturezaOp = .sNaturezaOp
        objNFiscal.sPlaca = .sPlaca
        objNFiscal.sPlacaUF = .sPlacaUF
        objNFiscal.lVolumeEspecie = .lVolumeEspecie
        objNFiscal.lVolumeMarca = .lVolumeMarca
        objNFiscal.iFreteRespons = .iFreteRespons
        
        objNFiscal.iMoeda = .iMoeda
        objNFiscal.dTaxaMoeda = .dTaxaMoeda
        objNFiscal.dVolumeTotal = .dVolumeTotal
        objNFiscal.iSemDataSaida = .iSemDataSaida
        objNFiscal.iNaoImpCobranca = .iNaoImpCobranca
        objNFiscal.sMensagemCorpoNota = .sMensagemCorpoNota
        
    End With

    'Fechar o Comando
    Call Comando_Fechar(lComando)

    NFiscal_Le_Num_Forn = SUCESSO

    Exit Function

Erro_NFiscal_Le_Num_Forn:

    NFiscal_Le_Num_Forn = Err

    Select Case Err

        Case 31464
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 31465, 31466
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL1", Err, objNFiscal.iTipoNFiscal, objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal)

        Case 31467

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149118)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function NFiscal_Le_Num_Forn1(ByVal objNFiscal As ClassNFiscal) As Long
'Lê a Nota Fiscal a partir de TipoNFiscal, Fornecedor, Filial, Serie e NFiscal.

Dim lErro As Long
Dim lComando As Long
Dim tNFiscal As typeNFiscal

On Error GoTo Erro_NFiscal_Le_Num_Forn1

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 55882

    With tNFiscal

        .sMensagemNota = String(STRING_NFISCAL_MENSAGEM, 0)
        .sMensagemCorpoNota = String(STRING_NFISCAL_MENSAGEM, 0)
        .sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
        .sPlaca = String(STRING_NFISCAL_PLACA, 0)
        .sPlacaUF = String(STRING_NFISCAL_PLACA_UF, 0)
        .sSerie = String(STRING_SERIE, 0)
        .sNumPedidoTerc = String(STRING_NUM_PEDIDO_TERC, 0)
        .sVolumeNumero = String(STRING_NFISCAL_VOLUME_NUMERO, 0)
        
        'Removido por Luiz Nogueira em 21/08/03
        '.lVolumeEspecie = String(STRING_NFISCAL_VOLUME_ESPECIE, 0)
        '.lVolumeMarca = String(STRING_NFISCAL_VOLUME_MARCA, 0)
        
        'Pesquisa a Nota Fiscal com os dados passados
        lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Status, Cliente, FilialCli, FilialEntrega, DataEmissao, DataEntrada, DataSaida, DataVencimento, NumPedidoVenda, NumPedidoTerc, ClasseDocCPR, NumIntDocCPR, ValorTotal, ValorProdutos, ValorFrete, ValorSeguro, ValorOutrasDespesas, ValorDesconto, CodTransportadora, MensagemNota, TabelaPreco , NaturezaOp, PesoLiq, PesoBruto, NumIntTrib, Placa, PlacaUF, VolumeQuant, VolumeEspecie, VolumeMarca, Canal, NumIntNotaOriginal,VolumeNumero, FreteRespons, NumRecebimento, CodTranspRedesp, DetPagFrete, Moeda, TaxaMoeda, VolumeTotal, SemDataSaida, NaoImpCobranca, MensagemCorpoNota FROM NFiscal WHERE FilialEmpresa = ? AND NumNotaFiscal = ? AND TipoNFiscal = ? AND Fornecedor = ? AND FilialForn = ? AND Serie = ? AND Status <> ? ORDER BY DataEntrada DESC", .lNumIntDoc, .iStatus, .lCliente, .iFilialCli, .iFilialEntrega, .dtDataEmissao, .dtDataEntrada, .dtDataSaida, .dtDataVencimento, _
            .lNumPedidoVenda, .sNumPedidoTerc, .iClasseDocCPR, .lNumIntDocCPR, .dValorTotal, .dValorProdutos, .dValorFrete, .dValorSeguro, .dValorOutrasDespesas, .dValorDesconto, .iCodTransportadora, .sMensagemNota, .iTabelaPreco, .sNaturezaOp, .dPesoLiq, .dPesoBruto, .lNumIntTrib, .sPlaca, .sPlacaUF, .lVolumeQuant, .lVolumeEspecie, .lVolumeMarca, .iCanal, .lNumIntNotaOriginal, .sVolumeNumero, .iFreteRespons, .lNumRecebimento, .iCodTranspRedesp, .iDetPagFrete, .iMoeda, .dTaxaMoeda, .dVolumeTotal, .iSemDataSaida, .iNaoImpCobranca, .sMensagemCorpoNota, giFilialEmpresa, objNFiscal.lNumNotaFiscal, objNFiscal.iTipoNFiscal, objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.sSerie, STATUS_EXCLUIDO)
        
    End With
        
    If lErro <> AD_SQL_SUCESSO Then Error 55883

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 55884

    'Se não encontrar --> Erro
    If lErro = AD_SQL_SEM_DADOS Then Error 55885

    With tNFiscal
    
        'Preenche objNFiscal com os dados lidos do BD
        objNFiscal.lNumIntDoc = .lNumIntDoc
        objNFiscal.dPesoBruto = .dPesoBruto
        objNFiscal.dPesoLiq = .dPesoLiq
        objNFiscal.dtDataEmissao = .dtDataEmissao
        objNFiscal.dtDataEntrada = .dtDataEntrada
        objNFiscal.dtDataSaida = .dtDataSaida
        objNFiscal.dtDataVencimento = .dtDataVencimento
        objNFiscal.dValorDesconto = .dValorDesconto
        objNFiscal.dValorFrete = .dValorFrete
        objNFiscal.dValorOutrasDespesas = .dValorOutrasDespesas
        objNFiscal.dValorProdutos = .dValorProdutos
        objNFiscal.dValorSeguro = .dValorSeguro
        objNFiscal.dValorTotal = .dValorTotal
        objNFiscal.iCanal = .iCanal
        objNFiscal.iClasseDocCPR = .iClasseDocCPR
        objNFiscal.iCodTransportadora = .iCodTransportadora
        objNFiscal.iCodTranspRedesp = .iCodTranspRedesp
        objNFiscal.iDetPagFrete = .iDetPagFrete
        objNFiscal.iFilialCli = .iFilialCli
        objNFiscal.iFilialEmpresa = giFilialEmpresa
        objNFiscal.iFilialEntrega = .iFilialEntrega
        objNFiscal.iStatus = .iStatus
        objNFiscal.iTabelaPreco = .iTabelaPreco
        objNFiscal.lVolumeQuant = .lVolumeQuant
        objNFiscal.lCliente = .lCliente
        objNFiscal.lNumIntDocCPR = .lNumIntDocCPR
        objNFiscal.lNumIntNotaOriginal = .lNumIntNotaOriginal
        objNFiscal.lNumIntTrib = .lNumIntTrib
        objNFiscal.sNumPedidoTerc = .sNumPedidoTerc
        objNFiscal.lNumPedidoVenda = .lNumPedidoVenda
        objNFiscal.sMensagemNota = .sMensagemNota
        objNFiscal.sNaturezaOp = .sNaturezaOp
        objNFiscal.sPlaca = .sPlaca
        objNFiscal.sPlacaUF = .sPlacaUF
        objNFiscal.lVolumeEspecie = .lVolumeEspecie
        objNFiscal.lVolumeMarca = .lVolumeMarca
        objNFiscal.iFreteRespons = .iFreteRespons
        objNFiscal.lNumRecebimento = .lNumRecebimento
        objNFiscal.iMoeda = .iMoeda
        objNFiscal.dTaxaMoeda = .dTaxaMoeda
        objNFiscal.dVolumeTotal = .dVolumeTotal
        objNFiscal.iSemDataSaida = .iSemDataSaida
        objNFiscal.iNaoImpCobranca = .iNaoImpCobranca
        objNFiscal.sMensagemCorpoNota = .sMensagemCorpoNota
        
    End With

    'Fechar o Comando
    Call Comando_Fechar(lComando)

    NFiscal_Le_Num_Forn1 = SUCESSO

    Exit Function

Erro_NFiscal_Le_Num_Forn1:

    NFiscal_Le_Num_Forn1 = Err

    Select Case Err

        Case 55882
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 55883, 55884
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL1", Err, objNFiscal.iTipoNFiscal, objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal)

        Case 55885

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149119)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function NFiscal_Le_Num_Forn2(ByVal objNFiscal As ClassNFiscal) As Long
'Lê a Nota Fiscal a partir de Fornecedor, Filial, Serie e NFiscal e que não seja nota de complemento

Dim lErro As Long
Dim lComando As Long
Dim tNFiscal As typeNFiscal

On Error GoTo Erro_NFiscal_Le_Num_Forn2

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 92083

    With tNFiscal

        .sMensagemNota = String(STRING_NFISCAL_MENSAGEM, 0)
        .sMensagemCorpoNota = String(STRING_NFISCAL_MENSAGEM, 0)
        .sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
        .sPlaca = String(STRING_NFISCAL_PLACA, 0)
        .sPlacaUF = String(STRING_NFISCAL_PLACA_UF, 0)
        .sSerie = String(STRING_SERIE, 0)
        .sNumPedidoTerc = String(STRING_NUM_PEDIDO_TERC, 0)
        .sVolumeNumero = String(STRING_NFISCAL_VOLUME_NUMERO, 0)
        
        'Removido por Luiz Nogueira em 21/08/03
        '.lVolumeEspecie = String(STRING_NFISCAL_VOLUME_ESPECIE, 0)
        '.lVolumeMarca = String(STRING_NFISCAL_VOLUME_MARCA, 0)
        
        'Pesquisa a Nota Fiscal com os dados passados
        lErro = Comando_Executar(lComando, "SELECT TipoNFiscal, NumIntDoc, Status, Fornecedor, FilialForn, Cliente, FilialCli, FilialEntrega, DataEmissao, DataEntrada, DataSaida, DataVencimento, NumPedidoVenda, NumPedidoTerc, ClasseDocCPR, NumIntDocCPR, ValorTotal, ValorProdutos, ValorFrete, ValorSeguro, ValorOutrasDespesas, ValorDesconto, CodTransportadora, MensagemNota, TabelaPreco , NaturezaOp, PesoLiq, PesoBruto, NumIntTrib, Placa, PlacaUF, VolumeQuant, VolumeEspecie, VolumeMarca, Canal, NumIntNotaOriginal, VolumeNumero, FreteRespons, CodTranspRedesp, DetPagFrete, Moeda, TaxaMoeda, VolumeTotal, SemDataSaida, NaoImpCobranca, MensagemCorpoNota FROM TipoDocInfoNFiscal WHERE FilialEmpresa = ? AND NumNotaFiscal = ? AND Complementar <> ? AND Fornecedor = ? AND FilialForn = ? AND Serie = ? AND Status <> ? ORDER BY DataEntrada DESC", .iTipoNFiscal, .lNumIntDoc, .iStatus, .lFornecedor, .iFilialForn, .lCliente, .iFilialCli, .iFilialEntrega, .dtDataEmissao, .dtDataEntrada, .dtDataSaida, .dtDataVencimento, _
            .lNumPedidoVenda, .sNumPedidoTerc, .iClasseDocCPR, .lNumIntDocCPR, .dValorTotal, .dValorProdutos, .dValorFrete, .dValorSeguro, .dValorOutrasDespesas, .dValorDesconto, .iCodTransportadora, .sMensagemNota, .iTabelaPreco, .sNaturezaOp, .dPesoLiq, .dPesoBruto, .lNumIntTrib, .sPlaca, .sPlacaUF, .lVolumeQuant, .lVolumeEspecie, .lVolumeMarca, .iCanal, .lNumIntNotaOriginal, .sVolumeNumero, .iFreteRespons, .iCodTranspRedesp, .iDetPagFrete, .iMoeda, .dTaxaMoeda, .dVolumeTotal, .iSemDataSaida, .iNaoImpCobranca, .sMensagemCorpoNota, giFilialEmpresa, objNFiscal.lNumNotaFiscal, DOCINFO_COMPLEMENTO, objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.sSerie, STATUS_EXCLUIDO)
        
    End With
        
    If lErro <> AD_SQL_SUCESSO Then gError 92084

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92085
    
    'Se não encontrar --> Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 92086

    With tNFiscal
    
        'Preenche objNFiscal com os dados lidos do BD
        objNFiscal.iTipoNFiscal = .iTipoNFiscal
        objNFiscal.lNumIntDoc = .lNumIntDoc
        objNFiscal.dPesoBruto = .dPesoBruto
        objNFiscal.dPesoLiq = .dPesoLiq
        objNFiscal.dtDataEmissao = .dtDataEmissao
        objNFiscal.dtDataEntrada = .dtDataEntrada
        objNFiscal.dtDataSaida = .dtDataSaida
        objNFiscal.dtDataVencimento = .dtDataVencimento
        objNFiscal.dValorDesconto = .dValorDesconto
        objNFiscal.dValorFrete = .dValorFrete
        objNFiscal.dValorOutrasDespesas = .dValorOutrasDespesas
        objNFiscal.dValorProdutos = .dValorProdutos
        objNFiscal.dValorSeguro = .dValorSeguro
        objNFiscal.dValorTotal = .dValorTotal
        objNFiscal.iCanal = .iCanal
        objNFiscal.iClasseDocCPR = .iClasseDocCPR
        objNFiscal.iCodTransportadora = .iCodTransportadora
        objNFiscal.iCodTranspRedesp = .iCodTranspRedesp
        objNFiscal.iDetPagFrete = .iDetPagFrete
        objNFiscal.iFilialCli = .iFilialCli
        objNFiscal.iFilialEmpresa = giFilialEmpresa
        objNFiscal.iFilialEntrega = .iFilialEntrega
        objNFiscal.iStatus = .iStatus
        objNFiscal.iTabelaPreco = .iTabelaPreco
        objNFiscal.lVolumeQuant = .lVolumeQuant
        objNFiscal.lCliente = .lCliente
        objNFiscal.lFornecedor = .lFornecedor
        objNFiscal.iFilialForn = .iFilialForn
        objNFiscal.lNumIntDocCPR = .lNumIntDocCPR
        objNFiscal.lNumIntNotaOriginal = .lNumIntNotaOriginal
        objNFiscal.lNumIntTrib = .lNumIntTrib
        objNFiscal.sNumPedidoTerc = .sNumPedidoTerc
        objNFiscal.lNumPedidoVenda = .lNumPedidoVenda
        objNFiscal.sMensagemNota = .sMensagemNota
        objNFiscal.sNaturezaOp = .sNaturezaOp
        objNFiscal.sPlaca = .sPlaca
        objNFiscal.sPlacaUF = .sPlacaUF
        objNFiscal.lVolumeEspecie = .lVolumeEspecie
        objNFiscal.lVolumeMarca = .lVolumeMarca
        objNFiscal.iFreteRespons = .iFreteRespons
        objNFiscal.iMoeda = .iMoeda
        objNFiscal.dTaxaMoeda = .dTaxaMoeda
        objNFiscal.dVolumeTotal = .dVolumeTotal
        objNFiscal.iSemDataSaida = .iSemDataSaida
        objNFiscal.iNaoImpCobranca = .iNaoImpCobranca
        objNFiscal.sMensagemCorpoNota = .sMensagemCorpoNota
        
    End With

    'Fechar o Comando
    Call Comando_Fechar(lComando)

    NFiscal_Le_Num_Forn2 = SUCESSO

    Exit Function

Erro_NFiscal_Le_Num_Forn2:

    NFiscal_Le_Num_Forn2 = gErr

    Select Case gErr

        Case 92083
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 92084, 92085
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL1", gErr, objNFiscal.iTipoNFiscal, objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal)

        Case 92086

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149120)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function NFiscal_Le_Num_Cli2(ByVal objNFiscal As ClassNFiscal) As Long
'Lê a Nota Fiscal a partir de Fornecedor, Filial, Serie e NFiscal e que não seja nota de complemento

Dim lErro As Long
Dim lComando As Long
Dim tNFiscal As typeNFiscal

On Error GoTo Erro_NFiscal_Le_Num_Cli2

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 92087

    With tNFiscal

        .sMensagemNota = String(STRING_NFISCAL_MENSAGEM, 0)
        .sMensagemCorpoNota = String(STRING_NFISCAL_MENSAGEM, 0)
        .sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
        .sPlaca = String(STRING_NFISCAL_PLACA, 0)
        .sPlacaUF = String(STRING_NFISCAL_PLACA_UF, 0)
        .sSerie = String(STRING_SERIE, 0)
        .sNumPedidoTerc = String(STRING_NUM_PEDIDO_TERC, 0)
        .sVolumeNumero = String(STRING_NFISCAL_VOLUME_NUMERO, 0)
        
        'Removido por Luiz Nogueira em 21/08/03
        '.lVolumeEspecie = String(STRING_NFISCAL_VOLUME_ESPECIE, 0)
        '.lVolumeMarca = String(STRING_NFISCAL_VOLUME_MARCA, 0)
        
        'Pesquisa a Nota Fiscal com os dados passados
        lErro = Comando_Executar(lComando, "SELECT TipoNFiscal, NumIntDoc, Status, Fornecedor, FilialForn, Cliente, FilialCli, FilialEntrega, DataEmissao, DataEntrada, DataSaida, DataVencimento, NumPedidoVenda, NumPedidoTerc, ClasseDocCPR, NumIntDocCPR, ValorTotal, ValorProdutos, ValorFrete, ValorSeguro, ValorOutrasDespesas, ValorDesconto, CodTransportadora, MensagemNota, TabelaPreco , NaturezaOp, PesoLiq, PesoBruto, NumIntTrib, Placa, PlacaUF, VolumeQuant, VolumeEspecie, VolumeMarca, Canal, NumIntNotaOriginal, VolumeNumero, FreteRespons, CodTranspRedesp, DetPagFrete, Moeda, TaxaMoeda, VolumeTotal, SemDataSaida, NaoImpCobranca, MensagemCorpoNota FROM TipoDocInfoNFiscal WHERE FilialEmpresa = ? AND NumNotaFiscal = ? AND Complementar <> ? AND Cliente = ? AND FilialCli = ? AND Serie = ? AND Status <> ? ORDER BY DataEntrada DESC", .iTipoNFiscal, .lNumIntDoc, .iStatus, .lFornecedor, .iFilialForn, .lCliente, .iFilialCli, .iFilialEntrega, .dtDataEmissao, .dtDataEntrada, .dtDataSaida, .dtDataVencimento, _
            .lNumPedidoVenda, .sNumPedidoTerc, .iClasseDocCPR, .lNumIntDocCPR, .dValorTotal, .dValorProdutos, .dValorFrete, .dValorSeguro, .dValorOutrasDespesas, .dValorDesconto, .iCodTransportadora, .sMensagemNota, .iTabelaPreco, .sNaturezaOp, .dPesoLiq, .dPesoBruto, .lNumIntTrib, .sPlaca, .sPlacaUF, .lVolumeQuant, .lVolumeEspecie, .lVolumeMarca, .iCanal, .lNumIntNotaOriginal, .sVolumeNumero, .iFreteRespons, .iCodTranspRedesp, .iDetPagFrete, .iMoeda, .dTaxaMoeda, .dVolumeTotal, .iSemDataSaida, .iNaoImpCobranca, .sMensagemCorpoNota, giFilialEmpresa, objNFiscal.lNumNotaFiscal, DOCINFO_COMPLEMENTO, objNFiscal.lCliente, objNFiscal.iFilialCli, objNFiscal.sSerie, STATUS_EXCLUIDO)
        
    End With
        
    If lErro <> AD_SQL_SUCESSO Then gError 92088

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92089
    
    'Se não encontrar --> Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 92090

    With tNFiscal
    
        'Preenche objNFiscal com os dados lidos do BD
        objNFiscal.iTipoNFiscal = .iTipoNFiscal
        objNFiscal.lNumIntDoc = .lNumIntDoc
        objNFiscal.dPesoBruto = .dPesoBruto
        objNFiscal.dPesoLiq = .dPesoLiq
        objNFiscal.dtDataEmissao = .dtDataEmissao
        objNFiscal.dtDataEntrada = .dtDataEntrada
        objNFiscal.dtDataSaida = .dtDataSaida
        objNFiscal.dtDataVencimento = .dtDataVencimento
        objNFiscal.dValorDesconto = .dValorDesconto
        objNFiscal.dValorFrete = .dValorFrete
        objNFiscal.dValorOutrasDespesas = .dValorOutrasDespesas
        objNFiscal.dValorProdutos = .dValorProdutos
        objNFiscal.dValorSeguro = .dValorSeguro
        objNFiscal.dValorTotal = .dValorTotal
        objNFiscal.iCanal = .iCanal
        objNFiscal.iClasseDocCPR = .iClasseDocCPR
        objNFiscal.iCodTransportadora = .iCodTransportadora
        objNFiscal.iCodTranspRedesp = .iCodTranspRedesp
        objNFiscal.iDetPagFrete = .iDetPagFrete
        objNFiscal.iFilialCli = .iFilialCli
        objNFiscal.iFilialEmpresa = giFilialEmpresa
        objNFiscal.iFilialEntrega = .iFilialEntrega
        objNFiscal.iStatus = .iStatus
        objNFiscal.iTabelaPreco = .iTabelaPreco
        objNFiscal.lVolumeQuant = .lVolumeQuant
        objNFiscal.lCliente = .lCliente
        objNFiscal.lFornecedor = .lFornecedor
        objNFiscal.iFilialForn = .iFilialForn
        objNFiscal.lNumIntDocCPR = .lNumIntDocCPR
        objNFiscal.lNumIntNotaOriginal = .lNumIntNotaOriginal
        objNFiscal.lNumIntTrib = .lNumIntTrib
        objNFiscal.sNumPedidoTerc = .sNumPedidoTerc
        objNFiscal.lNumPedidoVenda = .lNumPedidoVenda
        objNFiscal.sMensagemNota = .sMensagemNota
        objNFiscal.sNaturezaOp = .sNaturezaOp
        objNFiscal.sPlaca = .sPlaca
        objNFiscal.sPlacaUF = .sPlacaUF
        objNFiscal.lVolumeEspecie = .lVolumeEspecie
        objNFiscal.lVolumeMarca = .lVolumeMarca
        objNFiscal.iFreteRespons = .iFreteRespons
        objNFiscal.iMoeda = .iMoeda
        objNFiscal.dTaxaMoeda = .dTaxaMoeda
        objNFiscal.dVolumeTotal = .dVolumeTotal
        objNFiscal.iSemDataSaida = .iSemDataSaida
        objNFiscal.iNaoImpCobranca = .iNaoImpCobranca
        objNFiscal.sMensagemCorpoNota = .sMensagemCorpoNota
        
    End With

    'Fechar o Comando
    Call Comando_Fechar(lComando)

    NFiscal_Le_Num_Cli2 = SUCESSO

    Exit Function

Erro_NFiscal_Le_Num_Cli2:

    NFiscal_Le_Num_Cli2 = gErr

    Select Case gErr

        Case 92087
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 92088, 92089
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL1", gErr, objNFiscal.iTipoNFiscal, objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal)

        Case 92090

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149121)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function DataVencto_Real(ByVal dtDataVencimento As Date, dtDataVenctoReal As Date) As Long
'Gera a Data de Vencimento Real em cima da Data de Vencimento passada por parâmetro
'os feriados considerados sao da filialempesa corrente
'AVISO: ESSA ROTINA TEM QUE SER CHAMADA DENTRO DE UMA TRANSACAO

Dim vbDiaDaSemana As VbDayOfWeek
Dim dtFeriado As Date
Dim lErro As Long
Dim lComando As Long

On Error GoTo Erro_DataVencto_Real

    'Abre comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 18444
    
    dtDataVenctoReal = dtDataVencimento
    
    vbDiaDaSemana = Weekday(dtDataVenctoReal)
    
    'Se for a data for sábado ou domingo transferimos para a segunda
    If vbDiaDaSemana = vbSunday Then dtDataVenctoReal = dtDataVenctoReal + 1
    If vbDiaDaSemana = vbSaturday Then dtDataVenctoReal = dtDataVenctoReal + 2

    'Pesquisa-se as na Tabela de Feriados as datas maiores ou iguais a Data de Vencimento
    lErro = Comando_Executar(lComando, "SELECT Data FROM Feriados WHERE FilialEmpresa = ? AND Data >= ? ORDER BY Data", dtFeriado, giFilialEmpresa, dtDataVenctoReal)
    If lErro <> AD_SQL_SUCESSO Then Error 18445
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 18446
    
    Do While lErro = AD_SQL_SUCESSO
    
        'Se for maior é porque a data de vencimento não é feriado
        If (dtFeriado > dtDataVenctoReal) Then Exit Do

        'Incrementa-se a data
        dtDataVenctoReal = dtDataVenctoReal + 1
        
        vbDiaDaSemana = Weekday(dtDataVenctoReal)
        
        'Verifica se é sábado ou domingo
        If vbDiaDaSemana = vbSunday Then dtDataVenctoReal = dtDataVenctoReal + 1
        If vbDiaDaSemana = vbSaturday Then dtDataVenctoReal = dtDataVenctoReal + 2
        
        'Loop para adiantar DataFeriado se for menor que DataVenctoReal
        Do While dtFeriado < dtDataVenctoReal And lErro <> AD_SQL_SEM_DADOS
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 18447
        Loop
    Loop
    
    Call Comando_Fechar(lComando)
    
    DataVencto_Real = SUCESSO
    
    Exit Function
    
Erro_DataVencto_Real:

    DataVencto_Real = Err
    
    Select Case Err
    
        Case 18444
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 18445, 18446, 18447
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FERIADOS", Err)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149122)
            
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
        
End Function

'mario
Function FilialCliente_Lock1(ByVal lComando As Long, ByVal lCodCliente As Long, ByVal iCodFilial As Integer) As Long
'Faz lock da filial do cliente passado por parâmetro
'Chamada DENTRO DE TRANSAÇÃO

Dim lErro As Long
Dim lCliente As Long
Dim iFilial As Integer

On Error GoTo Erro_FilialCliente_Lock1

    'Lê a Filial do Cliente que será lockada
    lErro = Comando_ExecutarLockado(lComando, "SELECT CodCliente, CodFilial FROM FiliaisClientes WHERE CodCliente =? AND CodFilial = ?", lCliente, iFilial, lCodCliente, iCodFilial)
    If lErro <> AD_SQL_SUCESSO Then Error 43048

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 43049

    'Não achou a Filial do Cliente
    If lErro = AD_SQL_SEM_DADOS Then Error 43050

    'Lock da Filial do Cliente
    lErro = Comando_LockShared(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 43051

    FilialCliente_Lock1 = SUCESSO

    Exit Function

Erro_FilialCliente_Lock1:

    FilialCliente_Lock1 = Err

    Select Case Err
        
        Case 43048, 43049
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISCLIENTES", Err, lCodCliente, iCodFilial)

        Case 43050
            'Erro tratado na rotina chamada

        Case 43051
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_FILIAISCLIENTES", Err, lCodCliente, iCodFilial)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149123)

    End Select

    Exit Function

End Function

'mario
Function CanalVenda_Lock(ByVal lComando As Long, ByVal iCanalVenda As Integer) As Long
'Faz lock do Canal de Venda passado como parâmetro
'Chamada DENTRO DE TRANSAÇÃO

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_CanalVenda_Lock
    
    'Seleciona o Canal de Venda à ser lockado
    lErro = Comando_ExecutarLockado(lComando, "SELECT Codigo FROM CanalVenda WHERE Codigo =?", iCodigo, iCanalVenda)
    If lErro <> AD_SQL_SUCESSO Then Error 43024

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 43025

    'Não encontrou o Canal de Venda
    If lErro = AD_SQL_SEM_DADOS Then Error 43026

    lErro = Comando_LockShared(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 43027

    CanalVenda_Lock = SUCESSO

    Exit Function

Erro_CanalVenda_Lock:

    CanalVenda_Lock = Err

    Select Case Err
        
        Case 43024, 43025
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CANALVENDA1", Err, iCanalVenda)
            
        Case 43026
            'Erro tratado na rotina chamada

        Case 43027
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_CANALVENDA", Err, iCanalVenda)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149124)

    End Select

    Exit Function

End Function

'mario
Function Serie_Lock(ByVal lComando As Long, ByVal sSerie As String) As Long
'Faz lock da Série passada como parâmetro
'Chamada por DENTRO DE TRANSAÇÃO

Dim lErro As Long
Dim sSerie1 As String

On Error GoTo Erro_Serie_Lock

    sSerie1 = String(STRING_SERIE, 0)
    
    'Seleciona a Série à ser lockada
    lErro = Comando_ExecutarLockado(lComando, "SELECT Serie FROM Serie WHERE Serie =?", sSerie1, sSerie)
    If lErro <> AD_SQL_SUCESSO Then Error 43030

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 43031

    'Não encontrou a Série
    If lErro = AD_SQL_SEM_DADOS Then Error 43032

    lErro = Comando_LockShared(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 43033

    Serie_Lock = SUCESSO

    Exit Function

Erro_Serie_Lock:

    Serie_Lock = Err

    Select Case Err
        
        Case 43030, 43031
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SERIE1", Err, sSerie)
            
        Case 43032
            'Erro tratado na rotina chamada

        Case 43033
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_SERIE1", Err, sSerie)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149125)

    End Select

    Exit Function

End Function

'mario
Function Transportadora_Lock(ByVal lComando As Long, ByVal iTransportadora As Integer) As Long
'Faz lock da Transportadora passada como parâmetro
'Chamada por DENTRO DE TRANSAÇÃO

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_Transportadora_Lock

    'Seleciona a Transportadora à ser lockada
    lErro = Comando_ExecutarLockado(lComando, "SELECT Codigo FROM Transportadoras WHERE Codigo =?", iCodigo, iTransportadora)
    If lErro <> AD_SQL_SUCESSO Then Error 43036

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 43037

    'Não encontrou a Transportadora
    If lErro = AD_SQL_SEM_DADOS Then Error 43038

    lErro = Comando_LockShared(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 43039

    Transportadora_Lock = SUCESSO

    Exit Function

Erro_Transportadora_Lock:

    Transportadora_Lock = Err

    Select Case Err
        
        Case 43036, 43037
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRANSPORTADORA", Err, iTransportadora)
            
        Case 43038
            'Erro tratado na rotina chamada

        Case 43039
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_TRANSPORTADORA1", Err, iTransportadora)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149126)

    End Select

    Exit Function

End Function

'mario
Function Vendedor_Lock(ByVal lComando As Long, ByVal iVendedor As Integer) As Long
'Faz lock do Vendedor passado como parâmetro
'Chamada por DENTRO DE TRANSAÇÃO

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_Vendedor_Lock

    'Seleciona o Vendedor à ser lockado
    lErro = Comando_ExecutarLockado(lComando, "SELECT Codigo FROM Vendedores WHERE Codigo =?", iCodigo, iVendedor)
    If lErro <> AD_SQL_SUCESSO Then Error 43042

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 43043

    'Não encontrou o Vendedor
    If lErro = AD_SQL_SEM_DADOS Then Error 43044

    lErro = Comando_LockShared(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 43045

    Vendedor_Lock = SUCESSO

    Exit Function

Erro_Vendedor_Lock:

    Vendedor_Lock = Err

    Select Case Err
        
        Case 43041
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
                
        Case 43042, 43043
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_VENDEDOR", Err, iVendedor)
            
        Case 43044
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO", Err, iVendedor)

        Case 43045
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_VENDEDOR1", Err, iVendedor)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149127)

    End Select

    Exit Function

End Function

Function ComissoesNF_Le(ByVal objNFiscal As ClassNFiscal) As Long
'Le as Comissões da Nota Fiscal com o Número Interno passado para objNFiscal.colComissoesNF

'*** Alterada em 05/03/02 por Luiz Nogueira ***
'*** ClassNFiscal teve o tipo de ColComissoesNF alterado, por isso foi necessário adicionar
'os dados lidos a um obj e depois adicionar o obj à coleção ***

'Alterada por Tulio em 28/03: Inclusao do campo direto/indireto

Dim lErro As Long
Dim lComando As Long
Dim tComissaoNF As typeComissaoNF
Dim objComissaoNF As ClassComissaoNF

On Error GoTo Erro_ComissoesNF_Le

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 39031

    With tComissaoNF

        'Le as ComissoesNF com os dados passados
        lErro = Comando_Executar(lComando, "SELECT ComissoesNF.CodVendedor, ComissoesNF.ValorBase, ComissoesNF.Percentual, ComissoesNF.Valor, ComissoesNF.PercentualEmissao, ComissoesNF.ValorEmissao, ComissoesNF.Indireta, ComissoesNF.Seq FROM ComissoesNF, Vendedores WHERE ComissoesNF.NumIntDoc = ? AND ComissoesNF.CodVendedor = Vendedores.Codigo ORDER BY Seq", .iCodVendedor, .dValorBase, .dPercentual, .dValor, .dPercentualEmissao, .dValorEmissao, .iIndireta, .iSeq, objNFiscal.lNumIntDoc)
        
    End With
    
    If lErro <> AD_SQL_SUCESSO Then Error 39032

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 39033

    'Guar em colComissao as ComissoesNF encontradas
    Do While lErro = AD_SQL_SUCESSO

        'Instancia um novo objComissaoNF
        Set objComissaoNF = New ClassComissaoNF
        
        'Transfere os dados do type para o obj
        With objComissaoNF
        
            .lNumIntDoc = tComissaoNF.lNumIntDoc
            .iCodVendedor = tComissaoNF.iCodVendedor
            .dValorBase = tComissaoNF.dValorBase
            .dPercentual = tComissaoNF.dPercentual
            .dValor = tComissaoNF.dValor
            .dPercentualEmissao = tComissaoNF.dPercentualEmissao
            .dValorEmissao = tComissaoNF.dValorEmissao
            .iIndireta = tComissaoNF.iIndireta
            .iSeq = tComissaoNF.iSeq
                    
        End With
        
        'Adiciona o obj à coleção
        objNFiscal.ColComissoesNF.Add objComissaoNF

        'Busca a próxima
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 39034

    Loop

    'Fecha o Comando
    Call Comando_Fechar(lComando)

    ComissoesNF_Le = SUCESSO

    Exit Function

Erro_ComissoesNF_Le:

    ComissoesNF_Le = Err

    Select Case Err

        Case 39031
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 39032, 39033, 39034
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMISSOESNF1", Err, objNFiscal.lNumIntDoc)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149128)

    End Select

    'Fecha o Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'mario
Function ParcelasRecNF_Le(ByVal objNFiscal As ClassNFiscal) As Long
'Lê as parcelas a receber vinculadas à Nota Fiscal

Dim lErro As Long
Dim lComando As Long
Dim tParcelaReceber As typeParcelaReceber
Dim dtDataReferencia As Date
Dim objParcelaReceber As ClassParcelaReceber

On Error GoTo Erro_ParcelasRecNF_Le

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 42461
    
    '###############################################
    'INSERIDO POR WAGNER
    tParcelaReceber.sObservacao = String(STRING_MAX_NOME_ARQUIVO, 0)
    '###############################################
    
    tParcelaReceber.sAgenciaCheque = String(STRING_AGENCIA, 0)
    tParcelaReceber.sAprovacaoCartao = String(STRING_APROVACAO_CARTAO, 0)
    tParcelaReceber.sContaCorrenteCheque = String(STRING_CONTA_CORRENTE, 0)
    tParcelaReceber.sNumeroCartao = String(STRING_NUMERO_CARTAO, 0)
    
    'Busca as Parcelas a Receber vinculadas a Nfiscal
    '#################################################
    'ALTERADO POR WAGNER
    lErro = Comando_Executar(lComando, "SELECT NFiscal.DataReferencia, ParcelasRec.NumIntDoc, ParcelasRec.DataVencimento, ParcelasRec.Valor, ParcelasRec.Cobrador, ParcelasRec.CarteiraCobranca, ParcelasRec.Desconto1Codigo, ParcelasRec.Desconto1Ate, ParcelasRec.Desconto1Valor, ParcelasRec.Desconto2Codigo, ParcelasRec.Desconto2Ate, ParcelasRec.Desconto2Valor, ParcelasRec.Desconto3Codigo, ParcelasRec.Desconto3Ate, ParcelasRec.Desconto3Valor, ParcelasRec.Status, ParcelasRec.Previsao, ParcelasRec.Observacao, ParcelasRec.ValorOriginal, " & _
    "ParcelasRec.TipoPagto, ParcelasRec.CodConta, ParcelasRec.DataCredito, ParcelasRec.DataEmissaoCheque, ParcelasRec.BancoCheque, ParcelasRec.AgenciaCheque, ParcelasRec.ContaCorrenteCheque, ParcelasRec.NumeroCheque, ParcelasRec.DataDepositoCheque, ParcelasRec.AdmMeioPagto, ParcelasRec.Parcelamento, ParcelasRec.NumeroCartao, ParcelasRec.ValidadeCartao, ParcelasRec.AprovacaoCartao, ParcelasRec.DataTransacaoCartao FROM NFiscal, ParcelasRec WHERE NFiscal.NumIntDocCPR = ParcelasRec.NumIntTitulo AND NFiscal.NumIntDoc = ? ORDER BY ParcelasRec.NumParcela", _
        dtDataReferencia, tParcelaReceber.lNumIntDoc, tParcelaReceber.dtDataVencimento, tParcelaReceber.dValor, tParcelaReceber.iCobrador, tParcelaReceber.iCarteiraCobranca, tParcelaReceber.iDesconto1Codigo, tParcelaReceber.dtDesconto1Ate, tParcelaReceber.dDesconto1Valor, tParcelaReceber.iDesconto2Codigo, tParcelaReceber.dtDesconto2Ate, tParcelaReceber.dDesconto2Valor, tParcelaReceber.iDesconto3Codigo, tParcelaReceber.dtDesconto3Ate, tParcelaReceber.dDesconto3Valor, tParcelaReceber.iStatus, tParcelaReceber.iPrevisao, tParcelaReceber.sObservacao, tParcelaReceber.dValorOriginal, _
        tParcelaReceber.iTipoPagto, tParcelaReceber.iCodConta, tParcelaReceber.dtDataCredito, tParcelaReceber.dtDataEmissaoCheque, tParcelaReceber.iBancoCheque, tParcelaReceber.sAgenciaCheque, tParcelaReceber.sContaCorrenteCheque, tParcelaReceber.lNumeroCheque, tParcelaReceber.dtDataDepositoCheque, tParcelaReceber.iAdmMeioPagto, tParcelaReceber.iParcelamento, tParcelaReceber.sNumeroCartao, tParcelaReceber.dtValidadeCartao, tParcelaReceber.sAprovacaoCartao, tParcelaReceber.dtDataTransacaoCartao, objNFiscal.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then Error 42462
    '#################################################
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then Error 42463
    
    'Se não encontrou
    If lErro = AD_SQL_SEM_DADOS Then
        
        'Busca as Parcelas a receber Baixadas vinculadas a Nota Fiscal
        lErro = Comando_Executar(lComando, "SELECT NFiscal.DataReferencia, ParcelasRecBaixadas.NumIntDoc, ParcelasRecBaixadas.DataVencimento, ParcelasRecBaixadas.Valor, ParcelasRecBaixadas.Cobrador, ParcelasRecBaixadas.CarteiraCobranca, ParcelasRecBaixadas.Desconto1Codigo, ParcelasRecBaixadas.Desconto1Ate, ParcelasRecBaixadas.Desconto1Valor, ParcelasRecBaixadas.Desconto2Codigo, ParcelasRecBaixadas.Desconto2Ate, ParcelasRecBaixadas.Desconto2Valor, ParcelasRecBaixadas.Desconto3Codigo, ParcelasRecBaixadas.Desconto3Ate, ParcelasRecBaixadas.Desconto3Valor, ParcelasRecBaixadas.Status, ParcelasRecBaixadas.Observacao, ParcelasRecBaixadas.ValorOriginal, " & _
        "ParcelasRecBaixadas.TipoPagto, ParcelasRecBaixadas.CodConta, ParcelasRecBaixadas.DataCredito, ParcelasRecBaixadas.DataEmissaoCheque, ParcelasRecBaixadas.BancoCheque, ParcelasRecBaixadas.AgenciaCheque, ParcelasRecBaixadas.ContaCorrenteCheque, ParcelasRecBaixadas.NumeroCheque, ParcelasRecBaixadas.DataDepositoCheque, ParcelasRecBaixadas.AdmMeioPagto, ParcelasRecBaixadas.Parcelamento, ParcelasRecBaixadas.NumeroCartao, ParcelasRecBaixadas.ValidadeCartao, ParcelasRecBaixadas.AprovacaoCartao, ParcelasRecBaixadas.DataTransacaoCartao FROM NFiscal, ParcelasRecBaixadas WHERE NFiscal.NumIntDocCPR = ParcelasRecBaixadas.NumIntTitulo AND NFiscal.NumIntDoc = ? ORDER BY ParcelasRecBaixadas.NumParcela", _
            dtDataReferencia, tParcelaReceber.lNumIntDoc, tParcelaReceber.dtDataVencimento, tParcelaReceber.dValor, tParcelaReceber.iCobrador, tParcelaReceber.iCarteiraCobranca, tParcelaReceber.iDesconto1Codigo, tParcelaReceber.dtDesconto1Ate, tParcelaReceber.dDesconto1Valor, tParcelaReceber.iDesconto2Codigo, tParcelaReceber.dtDesconto2Ate, tParcelaReceber.dDesconto2Valor, tParcelaReceber.iDesconto3Codigo, tParcelaReceber.dtDesconto3Ate, tParcelaReceber.dDesconto3Valor, tParcelaReceber.iStatus, tParcelaReceber.sObservacao, tParcelaReceber.dValorOriginal, _
            tParcelaReceber.iTipoPagto, tParcelaReceber.iCodConta, tParcelaReceber.dtDataCredito, tParcelaReceber.dtDataEmissaoCheque, tParcelaReceber.iBancoCheque, tParcelaReceber.sAgenciaCheque, tParcelaReceber.sContaCorrenteCheque, tParcelaReceber.lNumeroCheque, tParcelaReceber.dtDataDepositoCheque, tParcelaReceber.iAdmMeioPagto, tParcelaReceber.iParcelamento, tParcelaReceber.sNumeroCartao, tParcelaReceber.dtValidadeCartao, tParcelaReceber.sAprovacaoCartao, tParcelaReceber.dtDataTransacaoCartao, objNFiscal.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then Error 42465
        
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 42466
                
    End If
    
    'Para cada parcela encontrada
    Do While lErro = AD_SQL_SUCESSO
        
        'Armazena em objNFiscal
        With tParcelaReceber
            '#################################################
            'ALTERADO POR WAGNER
            Set objParcelaReceber = objNFiscal.colParcelaReceber.Add(.lNumIntDoc, 0, 0, tParcelaReceber.iStatus, .dtDataVencimento, DATA_NULA, 0, .dValor, 0, .iCarteiraCobranca, .iCobrador, "", 0, 0, 0, 0, 0, 0, .iDesconto1Codigo, .dtDesconto1Ate, .dDesconto1Valor, .iDesconto2Codigo, .dtDesconto2Ate, .dDesconto2Valor, .iDesconto3Codigo, .dtDesconto3Ate, .dDesconto3Valor, 0, 0, 0, 0, .iPrevisao, .sObservacao, .dValorOriginal)
            '#################################################
            
            objParcelaReceber.iTipoPagto = .iTipoPagto
            objParcelaReceber.iCodConta = .iCodConta
            objParcelaReceber.dtDataCredito = .dtDataCredito
            objParcelaReceber.dtDataEmissaoCheque = .dtDataEmissaoCheque
            objParcelaReceber.iBancoCheque = .iBancoCheque
            objParcelaReceber.sAgenciaCheque = .sAgenciaCheque
            objParcelaReceber.sContaCorrenteCheque = .sContaCorrenteCheque
            objParcelaReceber.lNumeroCheque = .lNumeroCheque
            objParcelaReceber.dtDataDepositoCheque = .dtDataDepositoCheque
            objParcelaReceber.iAdmMeioPagto = .iAdmMeioPagto
            objParcelaReceber.iParcelamento = .iParcelamento
            objParcelaReceber.sNumeroCartao = .sNumeroCartao
            objParcelaReceber.dtValidadeCartao = .dtValidadeCartao
            objParcelaReceber.sAprovacaoCartao = .sAprovacaoCartao
            objParcelaReceber.dtDataTransacaoCartao = .dtDataTransacaoCartao
            
        End With

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 42468
    
    Loop
    
    objNFiscal.dtDataReferencia = dtDataReferencia
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    ParcelasRecNF_Le = SUCESSO
    
    Exit Function
    
Erro_ParcelasRecNF_Le:

    Select Case Err
    
        Case 42461
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 42462, 42463, 42468
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PARCELAS_REC_NF", Err, objNFiscal.lNumNotaFiscal)
        
        Case 42465, 42466
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PARCELAS_REC_BAIXADAS_NF", Err, objNFiscal.lNumNotaFiscal)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149129)
     
    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function CarteiraCobrador_LeMod(ByVal objCarteiraCobrador As ClassCarteiraCobrador, ByVal lComando As Long) As Long
'faz a leitura de um registro da tabela CarteirasCobrador p/objCarteiraCobrador c/ExecutarPos
'objCarteiraCobrador.iCobrador e objCarteiraCobrador.iCodCarteiraCobranca entram preenchidos
'TEM QUE SER CHAMADO EM TRANSACAO

Dim lErro As Long, tCarteiraCobrador As typeCarteiraCobrador

On Error GoTo Erro_CarteiraCobrador_LeMod

    With tCarteiraCobrador
        .sContaContabil = String(STRING_CONTA, 0)
        .sContaDuplDescontadas = String(STRING_CONTA, 0)
        .sFaixaNossoNumeroInicial = String(STRING_NOSSO_NUMERO, 0)
        .sFaixaNossoNumeroFinal = String(STRING_NOSSO_NUMERO, 0)
        .sFaixaNossoNumeroProx = String(STRING_NOSSO_NUMERO, 0)
        .sCodCarteiraNoBanco = String(STRING_CODCARTNOBANCO, 0)
    End With
    
    lErro = Comando_ExecutarPos(lComando, "SELECT ContaContabil, Desativada, DiasDeRetencao, TaxaCobranca, TaxaDesconto, ContaDuplDescontadas, QuantidadeAtual, QuantidadeAtualBanco, SaldoAtual, SaldoAtualBanco, FaixaNossoNumeroInicial, FaixaNossoNumeroFinal, FaixaNossoNumeroProx,ImprimeBoleta,ComRegistro,GeraNossoNumero, CodCarteiraNoBanco, NumCarteiraNoBanco FROM CarteirasCobrador WHERE Cobrador = ? AND CodCarteiraCobranca = ?", 0, tCarteiraCobrador.sContaContabil, tCarteiraCobrador.iDesativada, tCarteiraCobrador.iDiasDeRetencao, tCarteiraCobrador.dTaxaCobranca, tCarteiraCobrador.dTaxaDesconto, tCarteiraCobrador.sContaDuplDescontadas, tCarteiraCobrador.lQuantidadeAtual, tCarteiraCobrador.lQuantidadeAtualBanco, tCarteiraCobrador.dSaldoAtual, tCarteiraCobrador.dSaldoAtualBanco, _
        tCarteiraCobrador.sFaixaNossoNumeroInicial, tCarteiraCobrador.sFaixaNossoNumeroFinal, tCarteiraCobrador.sFaixaNossoNumeroProx, tCarteiraCobrador.iImprimeBoleta, tCarteiraCobrador.iComRegistro, tCarteiraCobrador.iGeraNossoNumero, tCarteiraCobrador.sCodCarteiraNoBanco, tCarteiraCobrador.iNumCarteiraNoBanco, objCarteiraCobrador.iCobrador, objCarteiraCobrador.iCodCarteiraCobranca)
    If lErro <> AD_SQL_SUCESSO Then Error 7357

    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 7556

    If lErro <> AD_SQL_SUCESSO Then Error 7557

    tCarteiraCobrador.iCobrador = objCarteiraCobrador.iCobrador
    tCarteiraCobrador.iCodCarteiraCobranca = objCarteiraCobrador.iCodCarteiraCobranca
    Call CarteiraCobrador_MoverCamposObj(objCarteiraCobrador.iCobrador, objCarteiraCobrador, tCarteiraCobrador)

    CarteiraCobrador_LeMod = SUCESSO

    Exit Function

Erro_CarteiraCobrador_LeMod:

    CarteiraCobrador_LeMod = Err

    Select Case Err

        Case 7537, 7556
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CARTEIRAS_COBRADOR", Err)

        Case 7557
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CARTEIRA_COBRADOR_INEXISTENTE", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149130)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function PedidosDeVenda_Le_NaturezaOp(ByVal sCodigo As String, iFilialEmpresa As Integer, lCodPedVenda As Long) As Long
'funcao pesquisa na tabela de Pedido de Venda a natureza passada se encontrou ---> ERRO

Dim lComando As Long
Dim lErro As Long
    
On Error GoTo Erro_PedidosDeVenda_Le_NaturezaOp

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 49904
    
    'verifica se a NaturezaOp está presente em algum registro de PedidosDeVenda
    lErro = Comando_Executar(lComando, "SELECT FilialEmpresa, Codigo FROM PedidosDeVenda WHERE NaturezaOp = ?", iFilialEmpresa, lCodPedVenda, sCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 49868
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 49869
    
    If lErro = AD_SQL_SUCESSO Then Error 49870 'Encontrou
            
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
    PedidosDeVenda_Le_NaturezaOp = SUCESSO
    
Erro_PedidosDeVenda_Le_NaturezaOp:
    
    PedidosDeVenda_Le_NaturezaOp = Err
    
    Select Case Err
        
        Case 49868, 49869
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDODEVENDA", Err)
            
        Case 49870 'Tratado na Rotina chamadora
        
        Case 49904
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case Else
                lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149131)
    End Select
            
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function PedidosDeVendaBaixado_Le_NaturezaOp(ByVal sCodigo As String, iFilialEmpresa As Integer, lCodPedVenda As Long) As Long
'funcao pesquisa na tabela de Pedido de Venda Baixado a natureza passada se encontrou ---> ERRO

Dim lComando As Long
Dim lErro As Long
    
On Error GoTo Erro_PedidosDeVendaBaixado_Le_NaturezaOp

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 49906
    
    'verifica se a NaturezaOp está presente em algum registro de PedidoDeVendaBaivados
    lErro = Comando_Executar(lComando, "SELECT FilialEmpresa, Codigo FROM PedidosDeVendaBaixados WHERE NaturezaOp = ?", iFilialEmpresa, lCodPedVenda, sCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 49871
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 49872
    
    If lErro = AD_SQL_SUCESSO Then Error 49873 'Encontrou
    
'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
    PedidosDeVendaBaixado_Le_NaturezaOp = SUCESSO
    
Erro_PedidosDeVendaBaixado_Le_NaturezaOp:
    
    PedidosDeVendaBaixado_Le_NaturezaOp = Err
    
    Select Case Err
        
        Case 49871, 49872
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDODEVENDABAIXADOS", Err)
        
        Case 49873
        
        Case 49906
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case Else
                lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149132)
    End Select
            
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function NFiscal_Le_NaturezaOp(ByVal sCodigo As String, iFilialEmpresa As Integer, lNumNotaFiscal As Long, sSerie As String) As Long
'funcao pesquisa na tabela de NFiscal a natureza passada se encontrou ---> ERRO

Dim lComando As Long
Dim lErro As Long
    
On Error GoTo Erro_NFiscal_Le_NaturezaOp

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 49907
    
    sSerie = String(STRING_SERIE, 0)
    
    'verifica se a NaturezaOp está presente em algum registro de NFiscal
    lErro = Comando_Executar(lComando, "SELECT FilialEmpresa, NumNotaFiscal, Serie FROM NFiscal WHERE NaturezaOp = ?", iFilialEmpresa, lNumNotaFiscal, sSerie, sCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 49874
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 49875
    
    If lErro = AD_SQL_SUCESSO Then Error 49876 'Encontrou
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
    NFiscal_Le_NaturezaOp = SUCESSO
    
Erro_NFiscal_Le_NaturezaOp:
    
    NFiscal_Le_NaturezaOp = Err
    
    Select Case Err
        
        Case 49874, 49875
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL3", Err)
        
        Case 49876 'Tratado na Rotina chamadora
                
        Case 49907
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case Else
                lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149133)
    End Select
            
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function PadraoTribSaida_Le_NaturezaOp(ByVal sCodigo As String) As Long
'funcao pesquisa na tabela de PadraoTribSaida a natureza passada se encontrou ---> ERRO

Dim lComando As Long
Dim lErro As Long
Dim sNaturezaOp As String
    
On Error GoTo Erro_PadraoTribSaida_Le_NaturezaOp

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 49909
    
    sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
    
    'verifica se a NaturezaOp está presente em algum registro de PadroesTribSaida
    lErro = Comando_Executar(lComando, "SELECT NaturezaOperacao FROM PadroesTribSaida WHERE NaturezaOperacao = ?", sNaturezaOp, sCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 49879
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 49880
    If lErro = AD_SQL_SUCESSO Then Error 49881 'Encontrou
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
        
    Exit Function
    
    PadraoTribSaida_Le_NaturezaOp = SUCESSO
    
Erro_PadraoTribSaida_Le_NaturezaOp:
    
    PadraoTribSaida_Le_NaturezaOp = Err
    
    Select Case Err
                
        Case 49879, 49880
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PADRAOTRIBSAIDA", Err)
        
        Case 49881
                
        Case 49909
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case Else
                lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149134)
    End Select
            
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function PadraoTribEntrada_Le_NaturezaOp(ByVal sCodigo As String) As Long
'funcao pesquisa na tabela de  PadraoTribEntrada a natureza passada se encontrou ---> ERRO

Dim lComando As Long
Dim lErro As Long
Dim sNaturezaOp As String
    
On Error GoTo Erro_PadraoTribEntrada_Le_NaturezaOp

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 49910
    
    sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
    
    'verifica se a NaturezaOp está presente em algum registro de PadroesTribEntrada
    lErro = Comando_Executar(lComando, "SELECT NaturezaOperacao FROM PadroesTribEntrada WHERE NaturezaOperacao = ?", sNaturezaOp, sCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 49882
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 49883
    If lErro = AD_SQL_SUCESSO Then Error 49884 'Encontrou
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
        
    Exit Function
    
    PadraoTribEntrada_Le_NaturezaOp = SUCESSO
    
Erro_PadraoTribEntrada_Le_NaturezaOp:
    
    PadraoTribEntrada_Le_NaturezaOp = Err
    
    Select Case Err
        
        Case 49882, 49883
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PADRAOTRIBENTRADA", Err)
        
        Case 49884
                
        Case 49910
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case Else
                lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149135)
    End Select
            
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function TipoDocInfo_Le_NaturezaOp(ByVal sCodigo As String, sSigla As String) As Long
'funcao pesquisa na tabela de TipoDocInfo a natureza passada se encontrou ---> ERRO

Dim lComando As Long
Dim lErro As Long
    
On Error GoTo Erro_TipoDocInfo_Le_NaturezaOp

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 49911
    
    sSigla = String(STRING_TIPO_DOC_INFO_SIGLA, 0)
    
    'verifica se a NaturezaOp está presente em algum registro de TipoDocInfo
    lErro = Comando_Executar(lComando, "SELECT Sigla FROM TiposDocInfo WHERE NaturezaOperacaoPadrao = ?", sSigla, sCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 49885
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 49886
    
    If lErro = AD_SQL_SUCESSO Then Error 49887 'Encontrou
            
    'Fecha o comando
    Call Comando_Fechar(lComando)
        
    Exit Function
    
    TipoDocInfo_Le_NaturezaOp = SUCESSO
    
Erro_TipoDocInfo_Le_NaturezaOp:
    
    TipoDocInfo_Le_NaturezaOp = Err
    
    Select Case Err
        
        Case 49885, 49886
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPODOCINFO1", Err)
        
        Case 49887
            
        Case 49911
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case Else
                lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149136)
    End Select
            
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function TributacaoNF_Le_NaturezaOp(ByVal sCodigo As String) As Long
'funcao pesquisa na tabela de TributacaoNF a natureza passada se encontrou ---> ERRO

Dim lComando As Long
Dim lErro As Long
Dim sNaturezaOp As String
    
On Error GoTo Erro_TributacaoNF_Le_NaturezaOp

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 49912
        
    sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
        
    'verifica se a NaturezaOp está presente em algum registro de TributacaoNF
    lErro = Comando_Executar(lComando, "SELECT NaturezaOpInterna FROM TributacaoNF WHERE NaturezaOpInterna = ?", sNaturezaOp, sCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 49888
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 49889
    If lErro = AD_SQL_SUCESSO Then Error 49890 'Encontrou
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
        
    Exit Function
    
    TributacaoNF_Le_NaturezaOp = SUCESSO
    
Erro_TributacaoNF_Le_NaturezaOp:
    
    TributacaoNF_Le_NaturezaOp = Err
    
    Select Case Err
        
        Case 49888, 49889
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRIBUTACAONF", Err)
        
        Case 49890
            
        Case 49912
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case Else
                lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149137)
    End Select
            
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function TributacaoItemPV_Le_NaturezaOp(ByVal sCodigo As String) As Long
'funcao pesquisa na tabela de TributacaoItemPV a natureza passada se encontrou ---> ERRO

Dim lComando As Long
Dim lErro As Long
Dim sNaturezaOp As String
    
On Error GoTo Erro_TributacaoItemPV_Le_NaturezaOp

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 49913
            
    sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
            
    'verifica se a NaturezaOp está presente em algum registro de TributacaoItemPV
    lErro = Comando_Executar(lComando, "SELECT NaturezaOp FROM TributacaoItemPV WHERE NaturezaOp = ?", sNaturezaOp, sCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 49891
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 49892
    If lErro = AD_SQL_SUCESSO Then Error 49903 'Encontrou
            
    'Fecha o comando
    Call Comando_Fechar(lComando)
        
    Exit Function
    
    TributacaoItemPV_Le_NaturezaOp = SUCESSO
    
Erro_TributacaoItemPV_Le_NaturezaOp:
    
    TributacaoItemPV_Le_NaturezaOp = Err
    
    Select Case Err
        
        Case 49891, 49892
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRIBUTACAOITEMPV", Err)
        
        Case 49903
            
        Case 49913
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case Else
                lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149138)
    End Select
            
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function TributacaoComplNF_Le_NaturezaOp(ByVal sCodigo As String) As Long
'funcao pesquisa na tabela de TributacaoComplNF a natureza passada se encontrou ---> ERRO

Dim lComando As Long
Dim lErro As Long
Dim sNaturezaOp As String
    
On Error GoTo Erro_TributacaoComplNF_Le_NaturezaOp

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 49914
            
    sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
            
    'verifica se a NaturezaOp está presente em algum registro de TributacaoComplNF
    lErro = Comando_Executar(lComando, "SELECT NaturezaOp FROM TributacaoComplNF WHERE NaturezaOp = ?", sNaturezaOp, sCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 49893
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 49894
    
    If lErro = AD_SQL_SUCESSO Then Error 49895 'Encontrou
            
    'Fecha o comando
    Call Comando_Fechar(lComando)
        
    Exit Function
    
    TributacaoComplNF_Le_NaturezaOp = SUCESSO
    
Erro_TributacaoComplNF_Le_NaturezaOp:
    
    TributacaoComplNF_Le_NaturezaOp = Err
    
    Select Case Err
        
        Case 49893, 49894
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRIBUTACAOCOMPLNF", Err)
        
        Case 49895
            
        Case 49914
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case Else
                lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149139)
    End Select
            
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function TributacaoComplPV_Le_NaturezaOp(ByVal sCodigo As String) As Long
'funcao pesquisa na tabela de TributacaoComplPV a natureza passada se encontrou ---> ERRO

Dim lComando As Long
Dim lErro As Long
Dim sNaturezaOp As String
    
On Error GoTo Erro_TributacaoComplPV_Le_NaturezaOp

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 49915
    
    sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
    
    'verifica se a NaturezaOp está presente em algum registro de TributacaoComplPV
    lErro = Comando_Executar(lComando, "SELECT NaturezaOp FROM TributacaoComplPV WHERE NaturezaOp = ?", sNaturezaOp, sCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 49896
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 49897
    If lErro = AD_SQL_SUCESSO Then Error 49898 'Encontrou
            
    'Fecha o comando
    Call Comando_Fechar(lComando)
        
    Exit Function
    
    TributacaoComplPV_Le_NaturezaOp = SUCESSO
    
Erro_TributacaoComplPV_Le_NaturezaOp:
    
    TributacaoComplPV_Le_NaturezaOp = Err
    
    Select Case Err
        
        Case 49896, 49897
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRIBUTACAOCOMPLPV", Err)
        
        Case 49898
        
        Case 49915
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case Else
                lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149140)
    End Select
            
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function TributacaoItemNF_Le_NaturezaOp(ByVal sCodigo As String) As Long
'funcao pesquisa na tabela de TributacaoItemNF a natureza passada se encontrou ---> ERRO

Dim lComando As Long
Dim lErro As Long
Dim sNaturezaOp As String
    
On Error GoTo Erro_TributacaoItemNF_Le_NaturezaOp

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 49916
                
    sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
                
    'verifica se a NaturezaOp está presente em algum registro de TributacaoItemNF
    lErro = Comando_Executar(lComando, "SELECT NaturezaOp FROM TributacaoItemNF WHERE NaturezaOp = ?", sNaturezaOp, sCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 49899
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 49900
    If lErro = AD_SQL_SUCESSO Then Error 49901 'Encontrou
                    
    'Fecha o comando
    Call Comando_Fechar(lComando)
        
    Exit Function
    
    TributacaoItemNF_Le_NaturezaOp = SUCESSO
    
Erro_TributacaoItemNF_Le_NaturezaOp:
    
    TributacaoItemNF_Le_NaturezaOp = Err
    
    Select Case Err
        
        Case 49899, 49900
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRIBUTACAOITEMNF", Err)
                
        Case 49901
            
        Case 49916
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case Else
                lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149141)
    End Select
            
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function Serie_Lock1(ByVal sSerie As String) As Long
'Faz lock da Série passada como parâmetro
'Chamada por DENTRO DE TRANSAÇÃO

Dim lErro As Long
Dim lComando As Long
Dim sSerie1 As String

On Error GoTo Erro_Serie_Lock1

    sSerie1 = String(STRING_SERIE, 0)
    
    'Inicializa comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 43029
    
    'Seleciona a Série à ser lockada
    lErro = Comando_ExecutarLockado(lComando, "SELECT Serie FROM Serie WHERE Serie =?", sSerie1, sSerie)
    If lErro <> AD_SQL_SUCESSO Then Error 43030

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 43031

    'Não encontrou a Série
    If lErro = AD_SQL_SEM_DADOS Then Error 43032

    lErro = Comando_LockShared(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 43033

    'Libera comando
    Call Comando_Fechar(lComando)

    Serie_Lock1 = SUCESSO

    Exit Function

Erro_Serie_Lock1:

    Serie_Lock1 = Err

    Select Case Err
        
        Case 43029
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
                
        Case 43030, 43031
            lErro = Rotina_Erro(vbOKOnly, "0", Err, sSerie)
            
        Case 43032
            'Erro tratado na rotina chamada

        Case 43033
            lErro = Rotina_Erro(vbOKOnly, "0", Err, sSerie)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149142)

    End Select

    'Libera comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Transportadora_Lock1(ByVal iTransportadora As Integer) As Long
'Faz lock da Transportadora passada como parâmetro
'Chamada por DENTRO DE TRANSAÇÃO

Dim lErro As Long
Dim lComando As Long
Dim iCodigo As Integer

On Error GoTo Erro_Transportadora_Lock1

    'Inicializa comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 43035
    
    'Seleciona a Transportadora à ser lockada
    lErro = Comando_ExecutarLockado(lComando, "SELECT Codigo FROM Transportadoras WHERE Codigo =?", iCodigo, iTransportadora)
    If lErro <> AD_SQL_SUCESSO Then Error 43036

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 43037

    'Não encontrou a Transportadora
    If lErro = AD_SQL_SEM_DADOS Then Error 43038

    lErro = Comando_LockShared(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 43039

    'Libera comando
    Call Comando_Fechar(lComando)

    Transportadora_Lock1 = SUCESSO

    Exit Function

Erro_Transportadora_Lock1:

    Transportadora_Lock1 = Err

    Select Case Err
        
        Case 43035
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
                
        Case 43036, 43037
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRANSPORTADORA", Err, iTransportadora)
            
        Case 43038
            'Erro tratado na rotina chamada

        Case 43039
            lErro = Rotina_Erro(vbOKOnly, "0", Err, iTransportadora)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149143)

    End Select

    'Libera comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function FilialCliente_Lock(ByVal lCodCliente As Long, ByVal iCodFilial As Integer, Optional ByVal objFilialCliente As ClassFilialCliente = Nothing, Optional ByVal lComandoP As Long = 0) As Long
'Faz lock da filial do cliente passado por parâmetro
'Chamada DENTRO DE TRANSAÇÃO

Dim lErro As Long
Dim lComando As Long
Dim lCliente As Long
Dim iFilial As Integer
Dim iAtivo As Integer

On Error GoTo Erro_FilialCliente_Lock

    If lComandoP <> 0 Then
        lComando = lComandoP
    Else
        'Inicializa comandos
        lComando = Comando_Abrir()
        If lComando = 0 Then Error 43047
    End If
    
    'Lê a Filial do Cliente que será lockada
    lErro = Comando_ExecutarLockado(lComando, "SELECT CodCliente, CodFilial, Ativo FROM FiliaisClientes WHERE CodCliente =? AND CodFilial = ?", lCliente, iFilial, iAtivo, lCodCliente, iCodFilial)
    If lErro <> AD_SQL_SUCESSO Then Error 43048

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 43049

    'Não achou a Filial do Cliente
    If lErro = AD_SQL_SEM_DADOS Then Error 43050

    'Lock da Filial do Cliente
    lErro = Comando_LockShared(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 43051

    If Not objFilialCliente Is Nothing Then
        objFilialCliente.lCodCliente = lCliente
        objFilialCliente.iCodFilial = iFilial
        objFilialCliente.iAtivo = iAtivo
    End If
    
    'Libera comandos
    If lComandoP = 0 Then Call Comando_Fechar(lComando)

    FilialCliente_Lock = SUCESSO

    Exit Function

Erro_FilialCliente_Lock:

    FilialCliente_Lock = Err

    Select Case Err
        
        Case 43047
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
                
        Case 43048, 43049
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISCLIENTES", Err, lCodCliente, iCodFilial)

        Case 43050
            'Erro tratado na rotina chamada

        Case 43051
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_FILIAISCLIENTES", Err, lCodCliente, iCodFilial)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149144)

    End Select

    'Libera comandos
    If lComandoP = 0 Then Call Comando_Fechar(lComando)

    Exit Function

End Function

Function FilialFornecedor_Lock(ByVal lCodFornecedor As Long, ByVal iCodFilial As Integer) As Long
'Faz lock da filial do Fornecedor passado por parâmetro
'Chamada DENTRO DE TRANSAÇÃO

Dim lErro As Long
Dim lComando As Long
Dim lFornecedor As Long
Dim iFilial As Integer

On Error GoTo Erro_FilialFornecedor_Lock

    'Inicializa comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 42867
    
    'Lê a Filial do Fornecedor que será lockada
    lErro = Comando_ExecutarLockado(lComando, "SELECT CodFornecedor, CodFilial FROM FiliaisFornecedores WHERE CodFornecedor =? AND CodFilial = ?", lFornecedor, iFilial, lCodFornecedor, iCodFilial)
    If lErro <> AD_SQL_SUCESSO Then Error 42868

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 42869

    'Não achou a Filial do Fornecedor
    If lErro = AD_SQL_SEM_DADOS Then Error 42870

    'Lock da Filial do Fornecedor
    lErro = Comando_LockShared(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 42871

    'Libera comandos
    Call Comando_Fechar(lComando)

    FilialFornecedor_Lock = SUCESSO

    Exit Function

Erro_FilialFornecedor_Lock:

    FilialFornecedor_Lock = Err

    Select Case Err
        
        Case 42867
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
                
        Case 42868, 42869
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISFORNECEDORES", Err, lCodFornecedor, iCodFilial)

        Case 42870
            'Erro tratado na rotina chamada

        Case 42871
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_FILIAISFORNECEDORES", Err, lCodFornecedor, iCodFilial)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149145)

    End Select

    'Libera comandos
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ItemNFiscal_Le(ByVal objItemNFiscal As ClassItemNF) As Long
'Lê o Ítem de Nota Fiscal com o Número Interno passado em objItemNFiscal
'obs.: o campo almoxarifado (sAlmoxarifadoNomeRed) nao está sendo preenchido

Dim lErro As Long
Dim lComando As Long
Dim tItemNFiscal As typeItemNF

On Error GoTo Erro_ItemNFiscal_Le

    'Abre o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 35222

    With tItemNFiscal

        .sDescricaoItem = String(STRING_ITEMNF_DESCRICAO, 0)
        .sProduto = String(STRING_PRODUTO, 0)
        .sUnidadeMed = String(STRING_UM_SIGLA, 0)

        'Busca o Ítem o Número Interno passado
        lErro = Comando_Executar(lComando, "SELECT Item, Produto, UnidadeMed, Quantidade, PrecoUnitario, PercDesc, ValorDesconto, DataEntrega, DescricaoItem, ValorAbatComissao, NumIntPedVenda, NumIntItemPedVenda, NumIntNF, NumIntTrib, ItensNFiscal.Almoxarifado, PrecoUnitarioMoeda, PercentMaisReceb, PercentMenosReceb, RecebForaFaixa, Comissao, TabelaPreco FROM ItensNFiscal WHERE NumIntDoc = ?", .iItem, .sProduto, .sUnidadeMed, .dQuantidade, .dPrecoUnitario, .dPercDesc, .dValorDesconto, .dtDataEntrega, .sDescricaoItem, .dValorAbatComissao, .lNumIntPedVenda, .lNumIntItemPedVenda, .lNumIntNF, .lNumIntTrib, .iAlmoxarifado, .dPrecoUnitarioMoeda, .dPercentMaisReceb, .dPercentMenosReceb, .iRecebForaFaixa, .dComissao, .iTabelaPreco, objItemNFiscal.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then Error 35223

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 35224
        If lErro = AD_SQL_SEM_DADOS Then Error 35225 'Não encontrou

        'Preenche o objItemNFiscal com os dados lidos
        objItemNFiscal.dPercDesc = .dPercDesc
        objItemNFiscal.dPrecoUnitario = .dPrecoUnitario
        objItemNFiscal.dQuantidade = .dQuantidade
        objItemNFiscal.dtDataEntrega = .dtDataEntrega
        objItemNFiscal.dValorAbatComissao = .dValorAbatComissao
        objItemNFiscal.dValorDesconto = .dValorDesconto
        objItemNFiscal.iAlmoxarifado = .iAlmoxarifado
        objItemNFiscal.iItem = .iItem
        objItemNFiscal.iStatus = .iStatus
        objItemNFiscal.lNumIntDocOrig = .lNumIntDocOrig
        objItemNFiscal.lNumIntItemPedVenda = .lNumIntItemPedVenda
        objItemNFiscal.lNumIntNF = .lNumIntNF
        objItemNFiscal.lNumIntPedVenda = .lNumIntPedVenda
        objItemNFiscal.lNumIntTrib = .lNumIntTrib
        objItemNFiscal.sAlmoxarifadoNomeRed = ""
        objItemNFiscal.sDescricaoItem = .sDescricaoItem
        objItemNFiscal.sProduto = .sProduto
        objItemNFiscal.sUnidadeMed = .sUnidadeMed
        objItemNFiscal.dPrecoUnitarioMoeda = .dPrecoUnitarioMoeda
        objItemNFiscal.dPercentMaisReceb = .dPercentMaisReceb
        objItemNFiscal.dPercentMenosReceb = .dPercentMenosReceb
        objItemNFiscal.iRecebForaFaixa = .iRecebForaFaixa
        objItemNFiscal.dComissao = .dComissao
        objItemNFiscal.iTabelaPreco = .iTabelaPreco
        
    End With

    'Fecha o comando
    Call Comando_Fechar(lComando)

    ItemNFiscal_Le = SUCESSO

    Exit Function

Erro_ItemNFiscal_Le:

    ItemNFiscal_Le = Err

    Select Case Err

        Case 35222
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 35223, 35224
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSNFISCAL1", Err)

        Case 35225

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149146)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

'mario
Function NFiscalItens_Le(ByVal objNFiscal As ClassNFiscal) As Long
'Lê os ítens da Nota Fiscal passada em objNFiscal

Dim iIndice As Integer
Dim lErro As Long
Dim lComando As Long
Dim tItemNF As typeItemNF
Dim colAlocacoes As ColAlocacoesItemNF
Dim objItemNFiscal As ClassItemNF
Dim iTipo As Integer
Dim objNFiscalAux As New ClassNFiscal

On Error GoTo Erro_NFiscalItens_Le

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 31426
    
    With tItemNF

        .sAlmoxarifadoNomeRed = String(STRING_ALMOXARIFADO_NOME_REDUZIDO, 0)
        .sDescricaoItem = String(STRING_ITEMNF_DESCRICAO, 0)
        .sProduto = String(STRING_PRODUTO, 0)
        .sUnidadeMed = String(STRING_UM_SIGLA, 0)
        .sCcl = String(STRING_CCL, 0)
        
        'Pesquisa os ítens vinculados à Nota Fiscal
        lErro = Comando_Executar(lComando, "SELECT Tipo, Item, Produto, UnidadeMed, Quantidade, PrecoUnitario, PercDesc, ValorDesconto, DataEntrega, DescricaoItem, ValorAbatComissao, NumIntPedVenda, NumIntItemPedVenda, NumIntDoc, NumIntTrib, NumIntDocOrig, Ccl, ItensNFiscal.Almoxarifado, PrecoUnitarioMoeda, PercentMaisReceb, PercentMenosReceb, RecebForaFaixa, Comissao, TabelaPreco FROM ItensNFiscal WHERE NumIntNF = ?  ORDER BY Item", iTipo, .iItem, .sProduto, .sUnidadeMed, .dQuantidade, .dPrecoUnitario, .dPercDesc, .dValorDesconto, .dtDataEntrega, .sDescricaoItem, .dValorAbatComissao, .lNumIntPedVenda, .lNumIntItemPedVenda, .lNumIntDoc, .lNumIntTrib, .lNumIntDocOrig, .sCcl, .iAlmoxarifado, .dPrecoUnitarioMoeda, .dPercentMaisReceb, .dPercentMenosReceb, .iRecebForaFaixa, .dComissao, .iTabelaPreco, objNFiscal.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 31427

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 31428

        Do While lErro <> AD_SQL_SEM_DADOS
            
            If iTipo = ITEMNF_TIPO_PECA Then
                
                'Guarda na coleção de ítens de objNFiscal
                Set objItemNFiscal = objNFiscal.ColItensNF.Add(.lNumIntNF, .iItem, .sProduto, .sUnidadeMed, .dQuantidade, .dPrecoUnitario, .dPercDesc, .dValorDesconto, .dtDataEntrega, .sDescricaoItem, .dValorAbatComissao, .lNumIntPedVenda, .lNumIntDoc, .lNumIntTrib, .lNumIntItemPedVenda, colAlocacoes, .iAlmoxarifado, "", .sCcl, .iStatus, .lNumIntDocOrig, "", 0, 0, 0, "", 0, .dPrecoUnitarioMoeda)
                
                objItemNFiscal.dPercentMaisReceb = .dPercentMaisReceb
                objItemNFiscal.dPercentMenosReceb = .dPercentMenosReceb
                objItemNFiscal.iRecebForaFaixa = .iRecebForaFaixa
                objItemNFiscal.dComissao = .dComissao
                objItemNFiscal.iTabelaPreco = .iTabelaPreco
               
            Else
            
                'Guarda na coleção de ítensSRV de objNFiscal
                'colItensSRV é uma coleção normal -> para manter o mesmo código foi
                'trocado o objNF só para gerar o item que é adicionado na coleção certa depois
                Set objItemNFiscal = objNFiscalAux.ColItensNF.Add(.lNumIntNF, .iItem, .sProduto, .sUnidadeMed, .dQuantidade, .dPrecoUnitario, .dPercDesc, .dValorDesconto, .dtDataEntrega, .sDescricaoItem, .dValorAbatComissao, .lNumIntPedVenda, .lNumIntDoc, .lNumIntTrib, .lNumIntItemPedVenda, colAlocacoes, .iAlmoxarifado, "", .sCcl, .iStatus, .lNumIntDocOrig, "", 0, 0, 0, "", 0, .dPrecoUnitarioMoeda)
                
                objItemNFiscal.dPercentMaisReceb = .dPercentMaisReceb
                objItemNFiscal.dPercentMenosReceb = .dPercentMenosReceb
                objItemNFiscal.iRecebForaFaixa = .iRecebForaFaixa
                objItemNFiscal.dComissao = .dComissao
                objItemNFiscal.iTabelaPreco = .iTabelaPreco
                
                objNFiscal.colItensSRV.Add objItemNFiscal
            
            End If
            
            lErro = CF("ItensNFGrade_Le", objItemNFiscal)
            If lErro <> SUCESSO Then gError 86368
            
            objItemNFiscal.objInfoAdicDocItem.iTipoDoc = TIPODOC_INFOADIC_ITEMNF
            objItemNFiscal.objInfoAdicDocItem.lNumIntDocItem = objItemNFiscal.lNumIntDoc
            lErro = CF("InfoAdicDocItem_Le", objItemNFiscal.objInfoAdicDocItem)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 86368
           
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 31429

        Loop

    End With

    lErro = CF("ItensNFEntrega_Le", objNFiscal)
    If lErro <> SUCESSO Then gError 183234

    lErro = CF("NFiscalItens_Le_Customizado", objNFiscal)
    If lErro <> SUCESSO Then gError 177045

    'Fecha os comandos
    Call Comando_Fechar(lComando)

    NFiscalItens_Le = SUCESSO

    Exit Function

Erro_NFiscalItens_Le:

    NFiscalItens_Le = gErr

    Select Case gErr

        Case 31426
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 31427, 31428, 31429
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSNFISCAL", gErr)

        Case 86368, 177045, 183234

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149147)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function NFiscalInternaSaida_Le_Numero(ByVal objNFiscalOriginal As ClassNFiscal) As Long
'Pesquisa nas Tabelas [NFiscal e TiposDocInfo] UNION [NotaFiscalBaixadas e TiposDocInfo]
'Busca pela Serie e Número passados
'Filtros : TiposDocInfo.Tipo=DOCINFO_NF_INT_SAIDA
'          NFiscal.Status<>STATUS_EXCLUIDO
'          NomeTelaNFiscal<>"NFiscalDev"
'Ordena por DataEmissao, DECENDING
'Lê todos os Dados da Tabela NFiscal

Dim lErro As Long
Dim lComando As Long
Dim sComandoSQL As String
Dim tNFiscal As typeNFiscal

On Error GoTo Erro_NFiscalInternaSaida_Le_Numero

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 30762

    If objNFiscalOriginal.iFilialEmpresa = 0 Then objNFiscalOriginal.iFilialEmpresa = giFilialEmpresa

    With tNFiscal

        .sMensagemNota = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sMensagemCorpoNota = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sNaturezaOp = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sPlaca = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sPlacaUF = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sSerie = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sNumPedidoTerc = String(STRING_BUFFER_MAX_TEXTO, 0)
        
        'Removido por Luiz Nogueira em 21/08/03
        '.lVolumeEspecie = String(STRING_NFISCAL_VOLUME_ESPECIE, 0)
        '.lVolumeMarca = String(STRING_NFISCAL_VOLUME_MARCA, 0)
        
        sComandoSQL = "SELECT NumIntDoc, Status, FilialEmpresa, Serie, NumNotaFiscal, Cliente, FilialCli, FilialEntrega, Fornecedor, FilialForn, DataEmissao, DataEntrada, DataSaida, DataVencimento, NumPedidoVenda, NumPedidoTerc, ClasseDocCPR, NumIntDocCPR, ValorTotal, ValorProdutos, ValorFrete, ValorSeguro, ValorOutrasDespesas, ValorDesconto, CodTransportadora, MensagemNota, TabelaPreco, TipoNFiscal, NaturezaOp, PesoLiq, PesoBruto, NumIntTrib, Placa, PlacaUF, VolumeQuant, VolumeEspecie, VolumeMarca, Canal, NumIntNotaOriginal, CodTranspRedesp, DetPagFrete, Moeda, TaxaMoeda, VolumeTotal, SemDataSaida, NaoImpCobranca, MensagemCorpoNota FROM NFiscal, TiposDocInfo WHERE NFiscal.TipoNFiscal = TiposDocInfo.Codigo AND NFiscal.NumNotaFiscal = ? AND NFiscal.Serie = ? AND TiposDocInfo.Tipo = ? AND NFiscal.Status <> ? AND TiposDocInfo.NomeTelaNFiscal <> ? AND NFiscal.FilialEmpresa = ? ORDER BY DataEmissao DESC "

        lErro = Comando_Executar(lComando, sComandoSQL, .lNumIntDoc, .iStatus, .iFilialEmpresa, .sSerie, .lNumNotaFiscal, .lCliente, .iFilialCli, .iFilialEntrega, .lFornecedor, .iFilialForn, .dtDataEmissao, .dtDataEntrada, .dtDataSaida, .dtDataVencimento, .lNumPedidoVenda, .sNumPedidoTerc, .iClasseDocCPR, .lNumIntDocCPR, .dValorTotal, .dValorProdutos, .dValorFrete, .dValorSeguro, .dValorOutrasDespesas, .dValorDesconto, .iCodTransportadora, .sMensagemNota, .iTabelaPreco, .iTipoNFiscal, .sNaturezaOp, .dPesoLiq, .dPesoBruto, .lNumIntTrib, .sPlaca, .sPlacaUF, .lVolumeQuant, .lVolumeEspecie, .lVolumeMarca, .iCanal, .lNumIntNotaOriginal, .iCodTranspRedesp, .iDetPagFrete, .iMoeda, .dTaxaMoeda, .dVolumeTotal, .iSemDataSaida, .iNaoImpCobranca, .sMensagemCorpoNota, _
            objNFiscalOriginal.lNumNotaFiscal, objNFiscalOriginal.sSerie, DOCINFO_NF_INT_SAIDA, STATUS_EXCLUIDO, "NFiscalDev", objNFiscalOriginal.iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then Error 30763

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 30764

        If lErro = AD_SQL_SEM_DADOS Then Error 30765

        objNFiscalOriginal.dPesoBruto = .dPesoBruto
        objNFiscalOriginal.dPesoLiq = .dPesoLiq
        objNFiscalOriginal.dtDataEmissao = .dtDataEmissao
        objNFiscalOriginal.dtDataEntrada = .dtDataEntrada
        objNFiscalOriginal.dtDataSaida = .dtDataSaida
        objNFiscalOriginal.dtDataVencimento = .dtDataVencimento
        objNFiscalOriginal.dValorDesconto = .dValorDesconto
        objNFiscalOriginal.dValorFrete = .dValorFrete
        objNFiscalOriginal.dValorOutrasDespesas = .dValorOutrasDespesas
        objNFiscalOriginal.dValorProdutos = .dValorProdutos
        objNFiscalOriginal.dValorSeguro = .dValorSeguro
        objNFiscalOriginal.dValorTotal = .dValorTotal
        objNFiscalOriginal.iCanal = .iCanal
        objNFiscalOriginal.iClasseDocCPR = .iClasseDocCPR
        objNFiscalOriginal.iCodTransportadora = .iCodTransportadora
        objNFiscalOriginal.iCodTranspRedesp = .iCodTranspRedesp
        objNFiscalOriginal.iDetPagFrete = .iDetPagFrete
        objNFiscalOriginal.iFilialCli = .iFilialCli
        objNFiscalOriginal.iFilialEntrega = .iFilialEntrega
        objNFiscalOriginal.iFilialForn = .iFilialForn
        objNFiscalOriginal.iStatus = .iStatus
        objNFiscalOriginal.iTabelaPreco = .iTabelaPreco
        objNFiscalOriginal.iTipoNFiscal = .iTipoNFiscal
        objNFiscalOriginal.lVolumeQuant = .lVolumeQuant
        objNFiscalOriginal.lCliente = .lCliente
        objNFiscalOriginal.lFornecedor = .lFornecedor
        objNFiscalOriginal.lNumIntDocCPR = .lNumIntDocCPR
        objNFiscalOriginal.lNumIntNotaOriginal = .lNumIntNotaOriginal
        objNFiscalOriginal.lNumIntTrib = .lNumIntTrib
        objNFiscalOriginal.lNumNotaFiscal = .lNumNotaFiscal
        objNFiscalOriginal.sNumPedidoTerc = .sNumPedidoTerc
        objNFiscalOriginal.lNumPedidoVenda = .lNumPedidoVenda
        objNFiscalOriginal.sMensagemNota = .sMensagemNota
        objNFiscalOriginal.sNaturezaOp = .sNaturezaOp
        objNFiscalOriginal.sPlaca = .sPlaca
        objNFiscalOriginal.sPlacaUF = .sPlacaUF
        objNFiscalOriginal.sSerie = .sSerie
        objNFiscalOriginal.lVolumeEspecie = .lVolumeEspecie
        objNFiscalOriginal.lVolumeMarca = .lVolumeMarca
        objNFiscalOriginal.lNumIntDoc = .lNumIntDoc
        objNFiscalOriginal.iFilialEmpresa = .iFilialEmpresa
        objNFiscalOriginal.iMoeda = .iMoeda
        objNFiscalOriginal.dTaxaMoeda = .dTaxaMoeda
        objNFiscalOriginal.dVolumeTotal = .dVolumeTotal
        objNFiscalOriginal.iSemDataSaida = .iSemDataSaida
        objNFiscalOriginal.iNaoImpCobranca = .iNaoImpCobranca
        objNFiscalOriginal.sMensagemCorpoNota = .sMensagemCorpoNota

    End With

    Call Comando_Fechar(lComando)

    NFiscalInternaSaida_Le_Numero = SUCESSO

    Exit Function

Erro_NFiscalInternaSaida_Le_Numero:

    NFiscalInternaSaida_Le_Numero = Err

    Select Case Err

        Case 30762
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 30763, 30764
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", Err)

        Case 30765 'Tratado pela rotina chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149148)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function TabelasPreco_Le_Codigos(ByVal colCodigo As Collection) As Long
'Lê todos os códigos contidos na Tabela De Preços e coloca-os em colCodigo

Dim lErro As Long
Dim lComando As Long
Dim iCodigo As Integer

On Error GoTo Erro_TabelasPreco_Le_Codigos

    'Abertura comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 17994

    'Lê a tabela de TabelasDePreco
    lErro = Comando_Executar(lComando, "SELECT Codigo FROM TabelasDePreco ORDER BY Codigo", iCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 17995

    'Lê o primeiro código da tabela de preços
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 17996

    Do While lErro <> AD_SQL_SEM_DADOS

        'Coloca o código lido na coleção
         colCodigo.Add iCodigo

        'Lê o próximo código
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 17997

    Loop

    'Fechamento comando
    Call Comando_Fechar(lComando)

    TabelasPreco_Le_Codigos = SUCESSO

    Exit Function

Erro_TabelasPreco_Le_Codigos:

    TabelasPreco_Le_Codigos = Err

    Select Case Err

        Case 17994
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 17995, 17996, 17997
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELASDEPRECO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149149)

    End Select

    'Fechamento comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function TabelaPreco_Le(ByVal objTabelaPreco As ClassTabelaPreco) As Long
'Lê a tabela de preço a partir da chave fornecida em objTabelaPreco. Retorna os dados em objTabelaPreco

Dim lErro As Long
Dim lComando As Long
Dim sDescricao As String
Dim iAjustaTabVlrCompCoefMaior As Integer
Dim dVlrCompCoef As Double
Dim dMarkUp As Double, dComissao As Double, iAtivo As Integer
Dim iCargoMinimo As Integer, iMoeda As Integer
Dim iTipo As Integer, iDescontoLimitado As Integer, dDescontoMaximo As Double

On Error GoTo Erro_TabelaPreco_Le

    'Abertura comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 28001

    sDescricao = String(STRING_TABELAPRECO_DESCRICAO, 0)

    'Pesquisa no BD a Tabela de Preço
    lErro = Comando_Executar(lComando, "SELECT Descricao, AjustaTabVlrCompCoefMaior, VlrCompCoef, MarkUp, CargoMinimo, Moeda, Comissao, Ativo, Tipo, DescontoLimitado, DescontoMaximo FROM TabelasDePreco WHERE Codigo = ?", sDescricao, iAjustaTabVlrCompCoefMaior, dVlrCompCoef, dMarkUp, iCargoMinimo, iMoeda, dComissao, iAtivo, iTipo, iDescontoLimitado, dDescontoMaximo, objTabelaPreco.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 28002

    'Verifica se Tabela de Preço existe
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 28003

    'Não encontrou a Tabela de Preço
    If lErro = AD_SQL_SEM_DADOS Then Error 28004

    'Preenche objTabelaPreco
    objTabelaPreco.sDescricao = sDescricao
    objTabelaPreco.iAjustaTabVlrCompCoefMaior = iAjustaTabVlrCompCoefMaior
    objTabelaPreco.dVlrCompCoef = dVlrCompCoef
    objTabelaPreco.dMarkUp = dMarkUp
    objTabelaPreco.iCargoMinimo = iCargoMinimo
    objTabelaPreco.iMoeda = iMoeda
    objTabelaPreco.dComissao = dComissao
    objTabelaPreco.iAtivo = iAtivo
    objTabelaPreco.iTipo = iTipo
    objTabelaPreco.iDescontoLimitado = iDescontoLimitado
    objTabelaPreco.dDescontoMaximo = dDescontoMaximo

    'Fechamento comando
    Call Comando_Fechar(lComando)

    TabelaPreco_Le = SUCESSO

    Exit Function

Erro_TabelaPreco_Le:

    TabelaPreco_Le = Err

    Select Case Err

        Case 28001
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 28002, 28003
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELASDEPRECO1", Err, objTabelaPreco.iCodigo)

        Case 28004
            'Não encontrou a Tabela de Preço. Erro tratado na rotina chamadora.

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149150)

    End Select

    'Fechamento comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function


Function FilialEmpresa_Le_Codigos(ByVal colCodigo As Collection) As Long
'Lê todos os códigos de Filiais da Empresa corrente na tabela FiliaisEmpresa.
'Devolve em colCodFilial

Dim lErro As Long
Dim lComando As Long
Dim iCodFilial As Integer

On Error GoTo Erro_FilialEmpresa_Le_Codigos

    'Abertura comando
    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 28152

    'Lê a tabela de FiliaisEmpresa
    lErro = Comando_Executar(lComando, "SELECT CodFilial FROM FiliaisEmpresas WHERE CodEmpresa = ? ORDER BY CodFilial", iCodFilial, glEmpresa)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 28153

    'Lê o primeiro código da tabela de FiliaisEmpresas
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 28154

    Do While lErro <> AD_SQL_SEM_DADOS

        'Coloca o código lido na coleção
         colCodigo.Add iCodFilial

        'Lê o próximo código
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 28155

    Loop

    'Fechamento do comando
    Call Comando_Fechar(lComando)

    FilialEmpresa_Le_Codigos = SUCESSO

    Exit Function

Erro_FilialEmpresa_Le_Codigos:

    FilialEmpresa_Le_Codigos = Err

    Select Case Err

        Case 28152
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 28153, 28154, 28155
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISEMPRESAS", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149151)

    End Select

    'Fechamento do comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Clientes_Le_TabelaPreco(ByVal objTabelaPreco As ClassTabelaPreco, ByVal colCodCliente As Collection) As Long
'Lê códigos de Clientes que se relacionam com objTabelaPreco e coloca-os em colCodClientes

Dim lErro As Long
Dim lComando As Long
Dim lCodigo As Long

On Error GoTo Erro_Clientes_Le_TabelaPreco

    'Abertura comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 28047

    'Verifica se o existe algum cliente associado a tabela de preço
    lErro = Comando_Executar(lComando, "SELECT  Codigo FROM Clientes WHERE TabelaPreco = ?", lCodigo, objTabelaPreco.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 28048

    'Lê o cliente, se estiver no BD
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 28049

    Do While lErro = AD_SQL_SUCESSO

        colCodCliente.Add lCodigo

        'Busca o próximo Cliente
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 28050

    Loop

    'Fechamento comando
    Call Comando_Fechar(lComando)

    Clientes_Le_TabelaPreco = SUCESSO

    Exit Function

Erro_Clientes_Le_TabelaPreco:

    Clientes_Le_TabelaPreco = Err

    Select Case Err

        Case 28047
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 28048, 28049, 28050
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLIENTES1", Err, objTabelaPreco.iCodigo)

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149152)

    End Select

    'Fechamento comando
    Call Comando_Fechar(lComando)

End Function

Function CRFATConfig_Le_Configs(ByVal colConfigs As ColCRFATConfig) As Long
'recupera da tabela FATConfig uma colecao de registros

Dim lErro As Long, lComando As Long
Dim objCRFATConfig As ClassCRFATConfig

On Error GoTo Erro_CRFATConfig_Le_Configs
    
    'abrir comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 55000
    
    'para cada elemento da colecao
    For Each objCRFATConfig In colConfigs
    
        lErro = CRFATConfig_Le_Batch(objCRFATConfig, lComando)
        If lErro <> SUCESSO And lErro <> 55003 Then Error 55004
        
        If lErro = 55003 Then Error 55005
        
    Next

    'fechar comando
    Call Comando_Fechar(lComando)
    
    CRFATConfig_Le_Configs = SUCESSO
    
    Exit Function
    
Erro_CRFATConfig_Le_Configs:

    CRFATConfig_Le_Configs = Err
    
    Select Case Err
    
        Case 55000
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 55004
        
        Case 55005
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CRFATCONFIG_INEXISTENTE", Err, objCRFATConfig.sCodigo, objCRFATConfig.iFilialEmpresa)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149153)
            
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Private Function CRFATConfig_Le_Batch(ByVal objCRFATConfig As ClassCRFATConfig, ByVal lComando As Long) As Long
'le registro em fatconfig.

Dim lErro As Long, tCRFATConfig As typeCRFATConfig

On Error GoTo Erro_CRFATConfig_Le_Batch

    With tCRFATConfig
        'inicializar buffers
        .sDescricao = String(STRING_CRFATCONFIG_DESCRICAO, 0)
        .sConteudo = String(STRING_CRFATCONFIG_CONTEUDO, 0)
        'ler registo
        lErro = Comando_Executar(lComando, "SELECT Descricao, Tipo, Conteudo FROM CRFATConfig WHERE Codigo = ? AND FilialEmpresa = ?", .sDescricao, .iTipo, .sConteudo, objCRFATConfig.sCodigo, objCRFATConfig.iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then Error 55001
        
    End With
    
    'Lê o primeiro registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 55002
    
    'Se não encontrou o registro
    If lErro = AD_SQL_SEM_DADOS Then Error 55003
    
    objCRFATConfig.sDescricao = tCRFATConfig.sDescricao
    objCRFATConfig.iTipo = tCRFATConfig.iTipo
    objCRFATConfig.sConteudo = tCRFATConfig.sConteudo
    
    CRFATConfig_Le_Batch = SUCESSO
     
    Exit Function
    
Erro_CRFATConfig_Le_Batch:

    CRFATConfig_Le_Batch = Err
     
    Select Case Err
     
        Case 55001, 55002
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CRFATCONFIG1", Err, objCRFATConfig.sCodigo, objCRFATConfig.iFilialEmpresa)
        
        Case 55003 'chave nao encontrada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149154)
     
    End Select
     
    Exit Function

End Function

Private Sub TituloReceber_MoverCamposObj(ByVal objTituloReceber As ClassTituloReceber, tTituloReceber As typeTituloReceber)

    objTituloReceber.lNumIntDoc = tTituloReceber.lNumIntDoc
    objTituloReceber.iFilialEmpresa = tTituloReceber.iFilialEmpresa
    objTituloReceber.lCliente = tTituloReceber.lCliente
    objTituloReceber.iFilial = tTituloReceber.iFilial
    objTituloReceber.sSiglaDocumento = tTituloReceber.sSiglaDocumento
    objTituloReceber.lNumTitulo = tTituloReceber.lNumTitulo
    objTituloReceber.dtDataEmissao = tTituloReceber.dtDataEmissao
    objTituloReceber.iStatus = tTituloReceber.iStatus
    objTituloReceber.dSaldo = tTituloReceber.dSaldo
    objTituloReceber.iNumParcelas = tTituloReceber.iNumParcelas
    objTituloReceber.dValor = tTituloReceber.dValor
    objTituloReceber.dValorIRRF = tTituloReceber.dValorIRRF
    objTituloReceber.dValorISS = tTituloReceber.dValorISS
    objTituloReceber.dISSRetido = tTituloReceber.dISSRetido
    objTituloReceber.dValorINSS = tTituloReceber.dValorINSS
    objTituloReceber.iINSSRetido = tTituloReceber.iINSSRetido
    objTituloReceber.dPercJurosDiario = tTituloReceber.dPercJurosDiario
    objTituloReceber.dPercMulta = tTituloReceber.dPercMulta
    objTituloReceber.sObservacao = tTituloReceber.sObservacao
    objTituloReceber.iCondicaoPagto = tTituloReceber.iCondicaoPagto
    objTituloReceber.dtDataRegistro = tTituloReceber.dtDataRegistro
    objTituloReceber.dPISRetido = tTituloReceber.dPISRetido
    objTituloReceber.dCOFINSRetido = tTituloReceber.dCOFINSRetido
    objTituloReceber.dCSLLRetido = tTituloReceber.dCSLLRetido
    objTituloReceber.iMoeda = tTituloReceber.iMoeda
    objTituloReceber.sCcl = tTituloReceber.sCcl
    objTituloReceber.sNatureza = tTituloReceber.sNatureza
    objTituloReceber.lCpoGenerico1 = tTituloReceber.lCpoGenerico1
    objTituloReceber.iReajustePeriodicidade = tTituloReceber.iReajustePeriodicidade
    objTituloReceber.dtReajusteBase = tTituloReceber.dtReajusteBase
    objTituloReceber.dtReajustadoAte = tTituloReceber.dtReajustadoAte

End Sub


Function TipoDocumento_Le(ByVal objTipoDocumento As ClassTipoDocumento) As Long
'Lê o tipo de documento a partir da chave fornecida em objTipoDocumento. Retorna os dados em objTipoDocumento

Dim lErro As Long
Dim lComando As Long
Dim tTipoDocumento As typeTipoDocumento

On Error GoTo Erro_TipoDocumento_Le

    tTipoDocumento.sDescricao = String(STRING_TIPO_DOC_DESCRICAO, 0)

    'Abertura comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 17075

    'Lê os tipos de documentos cadastrados no BD
    lErro = Comando_Executar(lComando, "SELECT Descricao, Contabiliza, AumentaValorPagto,  EmNFFatPag, EmCreditopagForn, EmTituloRec, EmDebitosRecCli, ClasseDocCPR  FROM TiposDeDocumento WHERE Sigla = ?", _
        tTipoDocumento.sDescricao, tTipoDocumento.iContabiliza, tTipoDocumento.iAumentaValorPagto, tTipoDocumento.iEmNFFatPag, tTipoDocumento.iEmCreditoPagForn, tTipoDocumento.iEmTituloRec, tTipoDocumento.iEmDebitosRecCli, tTipoDocumento.iClasseDocCPR, objTipoDocumento.sSigla)
    If lErro <> AD_SQL_SUCESSO Then Error 17076

    'Tenta ler
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 17077

    'Não encontrou ==> erro
    If lErro = AD_SQL_SEM_DADOS Then Error 17078

    'Preenche objTipoDocumento
    objTipoDocumento.sDescricao = tTipoDocumento.sDescricao
    objTipoDocumento.iContabiliza = tTipoDocumento.iContabiliza
    objTipoDocumento.iAumentaValorPagto = tTipoDocumento.iAumentaValorPagto
    objTipoDocumento.iEmNFFatPag = tTipoDocumento.iEmNFFatPag
    objTipoDocumento.iEmCreditoPagForn = tTipoDocumento.iEmCreditoPagForn
    objTipoDocumento.iEmTituloRec = tTipoDocumento.iEmTituloRec
    objTipoDocumento.iEmDebitosRecCli = tTipoDocumento.iEmDebitosRecCli
    objTipoDocumento.iClasseDocCPR = tTipoDocumento.iClasseDocCPR

    'Fechamento comando
    Call Comando_Fechar(lComando)

    TipoDocumento_Le = SUCESSO

    Exit Function

Erro_TipoDocumento_Le:

    TipoDocumento_Le = Err

    Select Case Err

        Case 17075
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 17076, 17077
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPODOCUMENTO", Err, objTipoDocumento.sSigla)

        Case 17078
            'Não encontrou o Tipo de Documento. Erro tratado na rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149155)

    End Select

    'Fechamento comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function TituloReceberBaixado_Le(ByVal objTituloReceber As ClassTituloReceber, Optional ByVal iExcluido As Integer = 0) As Long
'Lê Titulo a Receber Baixado com o Número Interno passado como parâmetro

Dim lErro As Long
Dim lComando As Long
Dim tTituloReceber As typeTituloReceber

On Error GoTo Erro_TituloReceberBaixado_Le

    'Abre comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 56567

    tTituloReceber.sSiglaDocumento = String(STRING_SIGLA_DOCUMENTO, 0)
    tTituloReceber.sObservacao = String(STRING_TITULO_OBSERVACAO, 0)
    tTituloReceber.sCcl = String(STRING_CCL, 0)
    tTituloReceber.sNatureza = String(STRING_NATMOVCTA_CODIGO, 0)
    
    If iExcluido = 0 Then
        'Seleciona o TituloReceber a partir do NumIntDoc
        lErro = Comando_Executar(lComando, "SELECT FilialEmpresa, Cliente, Filial, SiglaDocumento, NumTitulo, DataEmissao, Status, Saldo, NumParcelas, Valor, ValorIRRF, ValorISS, ISSRetido, PercJurosDiarios, PercMulta, Observacao, CondicaoPagto, PISRetido, COFINSRetido, CSLLRetido, Moeda, Ccl, Natureza, CpoGenerico1, ReajustePeriodicidade, ReajusteBase, ReajustadoAte FROM TitulosRecBaixados WHERE NumIntDoc = ? AND Status <> ? ", tTituloReceber.iFilialEmpresa, tTituloReceber.lCliente, tTituloReceber.iFilial, tTituloReceber.sSiglaDocumento, tTituloReceber.lNumTitulo, tTituloReceber.dtDataEmissao, tTituloReceber.iStatus, tTituloReceber.dSaldo, tTituloReceber.iNumParcelas, tTituloReceber.dValor, tTituloReceber.dValorIRRF, tTituloReceber.dValorISS, tTituloReceber.dISSRetido, tTituloReceber.dPercJurosDiario, tTituloReceber.dPercMulta, tTituloReceber.sObservacao, _
            tTituloReceber.iCondicaoPagto, tTituloReceber.dPISRetido, tTituloReceber.dCOFINSRetido, tTituloReceber.dCSLLRetido, tTituloReceber.iMoeda, tTituloReceber.sCcl, tTituloReceber.sNatureza, tTituloReceber.lCpoGenerico1, tTituloReceber.iReajustePeriodicidade, tTituloReceber.dtReajusteBase, tTituloReceber.dtReajustadoAte, objTituloReceber.lNumIntDoc, STATUS_EXCLUIDO)
    ElseIf iExcluido = 1 Then
        'Seleciona o TituloReceber a partir do NumIntDoc
        lErro = Comando_Executar(lComando, "SELECT FilialEmpresa, Cliente, Filial, SiglaDocumento, NumTitulo, DataEmissao, Status, Saldo, NumParcelas, Valor, ValorIRRF, ValorISS, ISSRetido, PercJurosDiarios, PercMulta, Observacao, CondicaoPagto, PISRetido, COFINSRetido, CSLLRetido, Moeda, Ccl, Natureza, CpoGenerico1, ReajustePeriodicidade, ReajusteBase, ReajustadoAte FROM TitulosRecBaixados WHERE NumIntDoc = ? ", tTituloReceber.iFilialEmpresa, tTituloReceber.lCliente, tTituloReceber.iFilial, tTituloReceber.sSiglaDocumento, tTituloReceber.lNumTitulo, tTituloReceber.dtDataEmissao, tTituloReceber.iStatus, tTituloReceber.dSaldo, tTituloReceber.iNumParcelas, tTituloReceber.dValor, tTituloReceber.dValorIRRF, tTituloReceber.dValorISS, tTituloReceber.dISSRetido, tTituloReceber.dPercJurosDiario, tTituloReceber.dPercMulta, tTituloReceber.sObservacao, _
            tTituloReceber.iCondicaoPagto, tTituloReceber.dPISRetido, tTituloReceber.dCOFINSRetido, tTituloReceber.dCSLLRetido, tTituloReceber.iMoeda, tTituloReceber.sCcl, tTituloReceber.sNatureza, tTituloReceber.lCpoGenerico1, tTituloReceber.iReajustePeriodicidade, tTituloReceber.dtReajusteBase, tTituloReceber.dtReajustadoAte, objTituloReceber.lNumIntDoc)
    End If
    If lErro <> AD_SQL_SUCESSO Then Error 56568

    'Tenta ler
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 56569

    'Se não encontrou --> Erro
    If lErro = AD_SQL_SEM_DADOS Then Error 56570

'    tTituloReceber.iFilialEmpresa = giFilialEmpresa
    tTituloReceber.lNumIntDoc = objTituloReceber.lNumIntDoc
    
    'Preenche objTituloReceber com os dados do Título
    Call TituloReceber_MoverCamposObj(objTituloReceber, tTituloReceber)
        
    'Fecha o comando
    Call Comando_Fechar(lComando)

    TituloReceberBaixado_Le = SUCESSO

    Exit Function

Erro_TituloReceberBaixado_Le:

    TituloReceberBaixado_Le = Err

    Select Case Err

        Case 56567
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 56568, 56569
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOSREC1", Err, objTituloReceber.lNumIntDoc)

        Case 56570  'Título não encontrado

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149156)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function FilFornFilEmp_Le_Lock(ByVal lComando As Long, ByVal objFilFornFilEmp As ClassFilFornFilEmp) As Long
'Lê e lock a tabela FilFornFilEmp com as chaves iFilialEmpresa, lCodFornecedor, iCodFilial vinda do objFilFornFilEmp

Dim lErro As Long
Dim tFilFornFilEmp As typeFilFornFilEmp

On Error GoTo Erro_FilFornFilEmp_Le_Lock

    'Lê a tabela FilFornFilEmp para trazer os dados e lockar
    lErro = Comando_ExecutarPos(lComando, "SELECT NumeroCompras, ValorAcumuladoCompras,ValorAcumuladoDevolucoes, NumTotalDevolucoes, PedidosEmAberto, AtrasoAcumulado, PedidosRecebidos, ItensPedidosRecebidos FROM FilialFornFilEmp WHERE FilialEmpresa = ? AND CodFornecedor =? AND CodFilial = ?", 0, tFilFornFilEmp.lNumeroCompras, tFilFornFilEmp.dValorAcumuladoCompras, tFilFornFilEmp.dValorAcumuladoDevolucoes, tFilFornFilEmp.lNumTotalDevolucoes, tFilFornFilEmp.lPedidosEmAberto, tFilFornFilEmp.lAtrasoAcumulado, tFilFornFilEmp.lPedidosRecebidos, tFilFornFilEmp.lItensPedidosRecebidos, objFilFornFilEmp.iFilialEmpresa, objFilFornFilEmp.lCodFornecedor, objFilFornFilEmp.iCodFilial)
    If lErro <> AD_SQL_SUCESSO Then Error 52640

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 52641

    'se não achou  - - -> Erro
    If lErro = AD_SQL_SEM_DADOS Then Error 52642

    'Lock da FilFornFilEmp
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 52643

    objFilFornFilEmp.lNumeroCompras = tFilFornFilEmp.lNumeroCompras
    objFilFornFilEmp.dValorAcumuladoCompras = tFilFornFilEmp.dValorAcumuladoCompras
    objFilFornFilEmp.dValorAcumuladoDevolucoes = tFilFornFilEmp.dValorAcumuladoDevolucoes
    objFilFornFilEmp.lNumTotalDevolucoes = tFilFornFilEmp.lNumTotalDevolucoes
    objFilFornFilEmp.lPedidosEmAberto = tFilFornFilEmp.lPedidosEmAberto
    objFilFornFilEmp.lAtrasoAcumulado = tFilFornFilEmp.lAtrasoAcumulado
    objFilFornFilEmp.lPedidosRecebidos = tFilFornFilEmp.lPedidosRecebidos
    objFilFornFilEmp.lItensPedidosRecebidos = tFilFornFilEmp.lItensPedidosRecebidos
       
    FilFornFilEmp_Le_Lock = SUCESSO
    
    Exit Function
    
Erro_FilFornFilEmp_Le_Lock:

    FilFornFilEmp_Le_Lock = Err

    Select Case Err

        Case 52640, 52641
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIALFORNFILEMP", Err)
        
        Case 52642 'Tratado na rotina chamadora
        
        Case 52643
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_FILIALFORNFILEMP", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149157)

    End Select

    Exit Function

End Function

Function TipoCliente_Transportadora_Nao_Relacionada(ByVal iCodigo As Integer) As Long
'Verifica se a Transportadora identificada por iCodigo está relacionada a algum Tipo de Cliente

Dim lErro As Long
Dim lComando As Long
Dim iCodigoTipo As Integer

On Error GoTo Erro_TipoCliente_Transportadora_Nao_Relacionada
    
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 58104

    lErro = Comando_Executar(lComando, "SELECT Codigo FROM TiposDeCliente WHERE CodTransportadora = ?", iCodigoTipo, iCodigo)
    If lErro <> SUCESSO Then Error 58105
        
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 58106
    
    'se está sendo utilizada ==>Erro.
    If lErro = AD_SQL_SUCESSO Then Error 58107
    
    Call Comando_Fechar(lComando)
    
    TipoCliente_Transportadora_Nao_Relacionada = SUCESSO
    
    Exit Function
    
Erro_TipoCliente_Transportadora_Nao_Relacionada:

    TipoCliente_Transportadora_Nao_Relacionada = Err
    
    Select Case Err
        
        Case 58104
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 58105, 58106
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOCLIENTE", Err)
                
        Case 58107
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_TIPOCLIENTE", Err, iCodigoTipo)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149158)
        
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function TitulosPag_Le_CondicaoPagto(ByVal iCodigo As Integer) As Long
'Testa a Integridade da Titulos a Pagar e Titulos a Pagar Baixados
'com relação a Condição de Pagamento

Dim lErro As Long
Dim lComando As Long
Dim lComando2 As Long
Dim lNumTitulo As Long
Dim iFilialEmpresa As Integer
Dim lNumTituloBaixados As Long
Dim iFilialEmpresaBaixados As Integer

On Error GoTo Erro_TitulosPag_Le_CondicaoPagto
    
    'Abre os Comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 58110
    
    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 58111
    
    'Seleciona Títulos a Pagar com a condição de Pagamento passada como parametro
    lErro = Comando_Executar(lComando, "SELECT NumTitulo, FilialEmpresa FROM TitulosPag WHERE CondicaoPagto = ?", lNumTitulo, iFilialEmpresa, iCodigo)
    If lErro <> SUCESSO Then Error 58112
        
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 58113
    
    'se está sendo utilizada ==>Erro.
    If lErro = AD_SQL_SUCESSO Then Error 58114
    
    'Seleciona Títulos a Pagar Baixados com a condição de Pagamento passada como parametro
    lErro = Comando_Executar(lComando, "SELECT NumTitulo, FilialEmpresa FROM TitulosPagBaixados WHERE CondicaoPagto = ?", lNumTituloBaixados, iFilialEmpresaBaixados, iCodigo)
    If lErro <> SUCESSO Then Error 58115
        
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 58116
    
    'se está sendo utilizada ==>Erro.
    If lErro = AD_SQL_SUCESSO Then Error 58117
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    
    TitulosPag_Le_CondicaoPagto = SUCESSO
    
    Exit Function
    
Erro_TitulosPag_Le_CondicaoPagto:
    
    TitulosPag_Le_CondicaoPagto = Err
    
    Select Case Err
        
        Case 58110, 58111
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 58112, 58113
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOS_PAGAR", Err)
                
        Case 58114
            Call Rotina_Erro(vbOKOnly, "ERRO_COND_PAGAMENTO_TITULOSPAG", Err, iCodigo)
        
        Case 58115, 58116
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOS_PAGAR_BAIXADO", Err)
        
        Case 58117
            Call Rotina_Erro(vbOKOnly, "ERRO_COND_PAGAMENTO_TITULOSPAG_BAIXADOS", Err, iCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149159)
        
    End Select
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    
    Exit Function
    
End Function

Function ChequePre_Le_FilialCliente(ByVal objFilialCliente As ClassFilialCliente) As Long
'Lê na tabela de ChequePre, FilialCliente à partir da chave(lCodigo) fornecida em objFilialCliente.

Dim lErro As Long
Dim lComando As Long
Dim lNumIntCheque As Long

On Error GoTo Erro_ChequePre_Le_FilialCliente

    'Inicializa comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 58135

    'Lê  Filial,Cliente na tabela ChequePre
    lErro = Comando_Executar(lComando, "SELECT NumIntCheque FROM ChequePre WHERE Cliente = ? AND Filial = ? ", lNumIntCheque, objFilialCliente.lCodCliente, objFilialCliente.iCodFilial)
    If lErro <> AD_SQL_SUCESSO Then Error 58136

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 58137
    
    'Não achou a Filial Cliente
    If lErro = AD_SQL_SUCESSO Then Error 58138

    Call Comando_Fechar(lComando)

    ChequePre_Le_FilialCliente = SUCESSO

Exit Function

Erro_ChequePre_Le_FilialCliente:

    ChequePre_Le_FilialCliente = Err

    Select Case Err

        Case 58135
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 58136, 58137
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CHEQUEPRE2", Err, objFilialCliente.lCodCliente)

        Case 58138
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_CHEQUEPRE", Err, objFilialCliente.iCodFilial, objFilialCliente.lCodCliente)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149160)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function TitulosRec_Le_CondicaoPagto(ByVal iCodigo As Integer) As Long
'Testa a Integridade da Titulos a Receber e Titulos a Receber Baixados
'com relação a Condição de Pagamento

Dim lErro As Long
Dim lComando As Long
Dim lComando2 As Long
Dim lNumTitulo As Long
Dim iFilialEmpresa As Integer
Dim lNumTituloBaixados As Long
Dim iFilialEmpresaBaixados As Integer

On Error GoTo Erro_TitulosRec_Le_CondicaoPagto
    
    'Abre os Comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 58150
    
    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 58151
    
    'Seleciona Títulos a Receber com a condição de Pagamento passada como parametro
    lErro = Comando_Executar(lComando, "SELECT NumTitulo, FilialEmpresa FROM TitulosRec WHERE CondicaoPagto = ?", lNumTitulo, iFilialEmpresa, iCodigo)
    If lErro <> SUCESSO Then Error 58152
        
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 58153
    
    'se está sendo utilizada ==>Erro.
    If lErro = AD_SQL_SUCESSO Then Error 58154
    
    'Seleciona Títulos a Receber Baixados com a condição de Pagamento passada como parametro
    lErro = Comando_Executar(lComando, "SELECT NumTitulo, FilialEmpresa FROM TitulosRecBaixados WHERE CondicaoPagto = ?", lNumTituloBaixados, iFilialEmpresaBaixados, iCodigo)
    If lErro <> SUCESSO Then Error 58155
        
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 58156
    
    'se está sendo utilizada ==>Erro.
    If lErro = AD_SQL_SUCESSO Then Error 58157
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    
    TitulosRec_Le_CondicaoPagto = SUCESSO
    
    Exit Function
    
Erro_TitulosRec_Le_CondicaoPagto:
    
    TitulosRec_Le_CondicaoPagto = Err
    
    Select Case Err
        
        Case 58150, 58151
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 58152, 58153
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOS_REC", Err)
                
        Case 58154
            Call Rotina_Erro(vbOKOnly, "ERRO_COND_PAGAMENTO_TITULOSREC", Err, iCodigo)
        
        Case 58155, 58156
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOS_REC_BAIXADOS", Err)
        
        Case 58157
            Call Rotina_Erro(vbOKOnly, "ERRO_COND_PAGAMENTO_TITULOSREC_BAIXADOS", Err, iCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149161)
        
    End Select
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    
    Exit Function
    
End Function

Public Function ParcelasPag_Le_DataValor(ByVal lNumIntTitulo As Long, ByVal colParcPag As ColParcelaPagar, lComando As Long) As Long
'Lê as datas de vencimento e os valores das ParcelasPag associadas ao título com NumIntTitulo passado.
'IMPORTANTE --- Esta função é usada para relatório.
'IMPORTANTE --- Caso o lComando venha zerado será aberto e permanecerá aberto.

Dim lErro As Long
Dim tParcelaPagar As typeParcelaPagar

On Error GoTo Erro_ParcelasPag_Le_DataValor
    
    'Abre o Comando passado
    If lComando = 0 Then
        lComando = Comando_AbrirExt(GL_lConexaoRel)
        If lComando = 0 Then Error 58231
    End If
    
    'Le a Data e o Valor da Parcela a Pagar
    lErro = Comando_Executar(lComando, "SELECT NumParcela, DataVencimento, Valor FROM ParcelasPag WHERE NumIntTitulo = ? ORDER BY NumParcela", tParcelaPagar.iNumParcela, tParcelaPagar.dtDataVencimento, tParcelaPagar.dValor, lNumIntTitulo)
    If lErro <> AD_SQL_SUCESSO Then Error 58232
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 58233
    
    Do While lErro = SUCESSO
    
        With tParcelaPagar
            
            'Preenche a colecao
            Call colParcPag.Add(.lNumIntDoc, .lNumIntTitulo, .iNumParcela, 0, .dtDataVencimento, DATA_NULA, 0, .dValor, 0, 0, 0, 0, "", "", CStr(tParcelaPagar.iNumParcela))
    
        End With
    
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 58234
        
    Loop
    
    ParcelasPag_Le_DataValor = SUCESSO
    
    Exit Function
    
Erro_ParcelasPag_Le_DataValor:

    ParcelasPag_Le_DataValor = Err
    
    Select Case Err
    
        Case 58231
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 58232, 58233, 58234
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PARCELAS_PAG1", Err)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149162)
    
    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function ParcelasRec_Le_DataValor(ByVal lNumIntTitulo As Long, ByVal colParcRec As colParcelaReceber, lComando As Long) As Long
'Lê as datas de vencimento e os valores das ParcelasRec associadas ao título com NumIntTitulo passado.
'IMPORTANTE --- Esta função é usada para relatório.
'IMPORTANTE --- Caso o lComando venha zerado será aberto e permanecerá aberto.

Dim lErro As Long
Dim tParcelaReceber As typeParcelaReceber

On Error GoTo Erro_ParcelasRec_Le_DataValor
    
    'Abre o Comando passado
    If lComando = 0 Then
        lComando = Comando_AbrirExt(GL_lConexaoRel)
        If lComando = 0 Then Error 58225
    End If
    
    'Le a Data e o Valor da Parcela a Receber
    lErro = Comando_Executar(lComando, "SELECT NumParcela, DataVencimento, Valor, Saldo, ValorOriginal FROM ParcelasRecTodas WHERE NumIntTitulo = ? ORDER BY NumParcela", tParcelaReceber.iNumParcela, tParcelaReceber.dtDataVencimento, tParcelaReceber.dValor, tParcelaReceber.dSaldo, tParcelaReceber.dValorOriginal, lNumIntTitulo)
    If lErro <> AD_SQL_SUCESSO Then Error 58226
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 58227
    
    Do While lErro = SUCESSO
    
        With tParcelaReceber
            
            '################################################
            'ALTERADO POR WAGNER
            colParcRec.Add .lNumIntDoc, .lNumIntTitulo, .iNumParcela, 0, .dtDataVencimento, DATA_NULA, .dSaldo, .dValor, 0, 0, 0, "", 0, 0, 0, 0, 0, 0, 0, DATA_NULA, 0, 0, 0, 0, 0, DATA_NULA, 0, 0, 0, 0, 0, .iPrevisao, .sObservacao, .dValorOriginal, CStr(.iNumParcela)
            '################################################
    
        End With
    
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 58228
        
    Loop
    
    ParcelasRec_Le_DataValor = SUCESSO
    
    Exit Function
    
Erro_ParcelasRec_Le_DataValor:

    ParcelasRec_Le_DataValor = Err
    
    Select Case Err
    
        Case 58225
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 58226, 58227, 58228
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PARCELAS_REC", Err)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149163)
    
    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ClienteCategorias_Le_FilialCliente(ByVal objFilialCliente As ClassFilialCliente, ByVal colCateg As Collection) As Long
'preenche uma colecao com objClienteCategoria(s) associados a objFilialCliente

Dim lErro As Long, objClienteCategoria As ClassFilialCliCategoria
Dim lComando As Long, sCategoria As String, sItem As String

On Error GoTo Erro_ClienteCategorias_Le_FilialCliente

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 27397
    
    sCategoria = String(STRING_CATEGORIACLIENTE_CATEGORIA, 0)
    sItem = String(STRING_CATEGORIACLIENTEITEM_ITEM, 0)
    
    lErro = Comando_Executar(lComando, "SELECT Categoria, Item FROM FilialClienteCategorias WHERE Cliente = ? AND Filial = ?", sCategoria, sItem, objFilialCliente.lCodCliente, objFilialCliente.iCodFilial)
    If lErro <> AD_SQL_SUCESSO Then Error 27398
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 27399
        
    Do While lErro <> AD_SQL_SEM_DADOS
        
        Set objClienteCategoria = New ClassFilialCliCategoria
        
        objClienteCategoria.lCliente = objFilialCliente.lCodCliente
        objClienteCategoria.iFilial = objFilialCliente.iCodFilial
        objClienteCategoria.sCategoria = sCategoria
        objClienteCategoria.sItem = sItem
        
        colCateg.Add objClienteCategoria
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 27400
    
    Loop
    
    lErro = Comando_Fechar(lComando)
    
    ClienteCategorias_Le_FilialCliente = SUCESSO

    Exit Function

Erro_ClienteCategorias_Le_FilialCliente:

    ClienteCategorias_Le_FilialCliente = Err

    Select Case Err

        Case 27397
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 27398, 27399, 27400
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CATEGS_CLI", Err, objFilialCliente.lCodCliente, objFilialCliente.iCodFilial)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149164)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function FilialCliente_Testa_CGC(ByVal lCodCliente As Long, ByVal iCodFilial As Integer, ByVal sCgc As String, Optional iLoja As Integer = 0) As Long
'Verifica se tem outro cliente com o mesmo CGC e dá aviso.
'Não está levando em conta Filial, será melhorado futuramente
'iLoja=1 indica que está sendo passado como parametro em lCodCliente , CodClienteLoja.

Dim lErro As Long
Dim lCliente As Long
Dim vbMsgRes As VbMsgBoxResult
Dim lComando As Long

On Error GoTo Erro_FilialCliente_Testa_CGC
     
    'Inicializa comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 58613
          
    If iLoja = 1 Then
          
        'Lê em Filial Cliente algum Cliente com o Mesmo CGC ou CPF
        lErro = Comando_Executar(lComando, "SELECT CodCliente FROM FiliaisClientes WHERE CGC = ? AND CodClienteLoja <> ? ", lCliente, sCgc, lCodCliente)
        If lErro <> AD_SQL_SUCESSO Then Error 58614
          
    Else
          
        'Lê em Filial Cliente algum Cliente com o Mesmo CGC ou CPF
        lErro = Comando_Executar(lComando, "SELECT CodCliente FROM FiliaisClientes WHERE CGC = ? AND CodCliente <> ?", lCliente, sCgc, lCodCliente)
        If lErro <> AD_SQL_SUCESSO Then Error 58614

    End If

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 58615
    
    'Se encontrou aviso:
    If lErro = AD_SQL_SUCESSO Then
    
        'Envia aviso se deve se gravar o Cliente com o mesmo CGC de um Cliente já Gravado
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CLIENTE_CGC_IGUAL", sCgc)

        If vbMsgRes = vbNo Then Error 58616
    
    End If
    
    Call Comando_Fechar(lComando)

    FilialCliente_Testa_CGC = SUCESSO
        
    Exit Function
    
Erro_FilialCliente_Testa_CGC:

    FilialCliente_Testa_CGC = Err
    
    Select Case Err
        
        Case 58613
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 58614, 58615
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLIENTES", Err)
        
        Case 58616 'CASO NAO QUEIRA GRAVAR
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149165)

    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Cobrador_Le(ByVal objCobrador As ClassCobrador) As Long
'le o Cobrador a partir da chave fornecida em objCobrador. Devolve os dados em objCobrador

Dim lComando As Long
Dim lErro As Long
Dim tCobrador As typeCobrador

On Error GoTo Erro_Cobrador_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 19291

    tCobrador.sNomeReduzido = String(STRING_COBRADOR_NOME_REDUZIDO, 0)
    tCobrador.sNome = String(STRING_COBRADORES_NOME, 0)

    'Verifica se o Cobrador existe, e se existir carrega seus dados em objCobrador
    lErro = Comando_Executar(lComando, "SELECT  FilialEmpresa, Inativo, NomeReduzido, Nome, Endereco, CodBanco, CobrancaEletronica, CodCCI, CNABProxSeqArqCobr, Fornecedor, Filial FROM Cobradores WHERE Codigo = ?", tCobrador.iFilialEmpresa, tCobrador.iInativo, tCobrador.sNomeReduzido, tCobrador.sNome, tCobrador.lEndereco, tCobrador.iCodBanco, tCobrador.iCobrancaEletronica, tCobrador.iCodCCI, tCobrador.lCNABProxSeqArqCobr, tCobrador.lFornecedor, tCobrador.iFilial, objCobrador.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 19292

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 19293

    'Cobrador não está cadastrado
    If lErro = AD_SQL_SEM_DADOS Then Error 19294
    
    objCobrador.iFilialEmpresa = tCobrador.iFilialEmpresa
    objCobrador.iInativo = tCobrador.iInativo
    objCobrador.sNomeReduzido = tCobrador.sNomeReduzido
    objCobrador.sNome = tCobrador.sNome
    objCobrador.lEndereco = tCobrador.lEndereco
    objCobrador.iCodBanco = tCobrador.iCodBanco
    objCobrador.iCobrancaEletronica = tCobrador.iCobrancaEletronica
    objCobrador.iCodCCI = tCobrador.iCodCCI
    objCobrador.lCNABProxSeqArqCobr = tCobrador.lCNABProxSeqArqCobr
    objCobrador.lFornecedor = tCobrador.lFornecedor
    objCobrador.iFilial = tCobrador.iFilial
    
    lErro = Comando_Fechar(lComando)

    Cobrador_Le = SUCESSO

    Exit Function

Erro_Cobrador_Le:

    Cobrador_Le = Err

    Select Case Err

        Case 19291
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 19292, 19293
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COBRADOR", Err, objCobrador.iCodigo)

        Case 19294 'Cobrador nao cadastrado

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149166)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Cobrador_Le_NomeReduzido(ByVal objCobrador As ClassCobrador) As Long
'Lê os dados do Cobrador à partir do Nome Reduzido passado como parâmetro

Dim lErro As Long
Dim lComando As Long
Dim tCobrador As typeCobrador

On Error GoTo Erro_Cobrador_Le_NomeReduzido

    'Inicializa comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 43554

    tCobrador.sNome = String(STRING_COBRADORES_NOME, 0)

    'Verifica se o Cobrador existe, e se existir carrega seus dados em objCobrador
    lErro = Comando_Executar(lComando, "SELECT  Codigo, FilialEmpresa, Inativo, Nome, Endereco, CodBanco, CobrancaEletronica, CodCCI, CNABProxSeqArqCobr, Fornecedor, Filial FROM Cobradores WHERE NomeReduzido = ?", tCobrador.iCodigo, tCobrador.iFilialEmpresa, tCobrador.iInativo, tCobrador.sNome, tCobrador.lEndereco, tCobrador.iCodBanco, tCobrador.iCobrancaEletronica, tCobrador.iCodCCI, tCobrador.lCNABProxSeqArqCobr, tCobrador.lFornecedor, tCobrador.iFilial, objCobrador.sNomeReduzido)
    If lErro <> AD_SQL_SUCESSO Then Error 43555

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 43556

    'Se não encontrou o Cobrador --> Erro
    If lErro = AD_SQL_SEM_DADOS Then Error 43557

    objCobrador.iCodigo = tCobrador.iCodigo
    objCobrador.iFilialEmpresa = tCobrador.iFilialEmpresa
    objCobrador.iInativo = tCobrador.iInativo
    objCobrador.sNome = tCobrador.sNome
    objCobrador.lEndereco = tCobrador.lEndereco
    objCobrador.iCodBanco = tCobrador.iCodBanco
    objCobrador.iCobrancaEletronica = tCobrador.iCobrancaEletronica
    objCobrador.iCodCCI = tCobrador.iCodCCI
    objCobrador.lCNABProxSeqArqCobr = tCobrador.lCNABProxSeqArqCobr
    objCobrador.lFornecedor = tCobrador.lFornecedor
    objCobrador.iFilial = tCobrador.iFilial
    
    'Libera comando
    Call Comando_Fechar(lComando)

    Cobrador_Le_NomeReduzido = SUCESSO

    Exit Function

Erro_Cobrador_Le_NomeReduzido:

    Cobrador_Le_NomeReduzido = Err

    Select Case Err

        Case 43554
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 43555, 43556
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COBRADOR1", Err, objCobrador.sNomeReduzido)

        Case 43557 'Cobrador não Cadastrado

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149167)

    End Select

    'Libera comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function


Function NFiscalInternaSaida_Le_Numero2(ByVal objNFiscalOriginal As ClassNFiscal) As Long
Attribute NFiscalInternaSaida_Le_Numero2.VB_Description = " Lê as Notas Fiscais Internas de Saída, incluindo as cadastradas na tela NFiscalDev, com a série e o número passados."
'Pesquisa nas Tabelas [NFiscal e TiposDocInfo] UNION [NotaFiscalBaixadas e TiposDocInfo]
'Busca pela Serie e Número passados
'Filtros : TiposDocInfo.Tipo=DOCINFO_NF_INT_SAIDA
'          NFiscal.Status<>STATUS_EXCLUIDO
'Ordena por DataEmissao, DESCENDING
'Lê todos os Dados da Tabela NFiscal

'**** DIFERENÇA ENTRE AS FUNÇÕES...
'... NFiscalInternaSaida_Le_Numero e NFiscalInternaSaida_Le_Numero2
'- Aqui não excluimos da seleção as NFs cadastradas na tela "NFiscalDev"

Dim lErro As Long
Dim lComando As Long
Dim sComandoSQL As String
Dim tNFiscal As typeNFiscal

On Error GoTo Erro_NFiscalInternaSaida_Le_Numero2

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 62141

    With tNFiscal

        .sMensagemNota = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sMensagemCorpoNota = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sNaturezaOp = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sPlaca = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sPlacaUF = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sSerie = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sNumPedidoTerc = String(STRING_BUFFER_MAX_TEXTO, 0)
        
        'Removido por Luiz Nogueira em 21/08/03
        '.lVolumeEspecie = String(STRING_NFISCAL_VOLUME_ESPECIE, 0)
        '.lVolumeMarca = String(STRING_NFISCAL_VOLUME_MARCA, 0)
        
        sComandoSQL = "SELECT NumIntDoc, Status, FilialEmpresa, Serie, NumNotaFiscal, Cliente, FilialCli, FilialEntrega, Fornecedor, FilialForn, DataEmissao, DataEntrada, DataSaida, DataVencimento, FilialPedido, NumPedidoVenda, NumPedidoTerc, ClasseDocCPR, NumIntDocCPR, ValorTotal, ValorProdutos, ValorFrete, ValorSeguro, ValorOutrasDespesas, ValorDesconto, CodTransportadora, MensagemNota, TabelaPreco, TipoNFiscal, NaturezaOp, PesoLiq, PesoBruto, NumIntTrib, Placa, PlacaUF, VolumeQuant, VolumeEspecie, VolumeMarca, Canal, NumIntNotaOriginal, CodTranspRedesp, DetPagFrete, Moeda, TaxaMoeda, VolumeTotal, SemDataSaida, NaoImpCobranca, MensagemCorpoNota, RPS FROM NFiscal, TiposDocInfo WHERE FilialEmpresa = ? AND NFiscal.TipoNFiscal = TiposDocInfo.Codigo AND NFiscal.NumNotaFiscal = ? AND NFiscal.Serie = ? AND TiposDocInfo.Tipo = ? AND NFiscal.Status <> ? ORDER BY DataEmissao DESC "

        lErro = Comando_Executar(lComando, sComandoSQL, .lNumIntDoc, .iStatus, .iFilialEmpresa, .sSerie, .lNumNotaFiscal, .lCliente, .iFilialCli, .iFilialEntrega, .lFornecedor, .iFilialForn, .dtDataEmissao, .dtDataEntrada, .dtDataSaida, .dtDataVencimento, .iFilialPedido, .lNumPedidoVenda, .sNumPedidoTerc, .iClasseDocCPR, .lNumIntDocCPR, .dValorTotal, .dValorProdutos, .dValorFrete, .dValorSeguro, .dValorOutrasDespesas, .dValorDesconto, .iCodTransportadora, .sMensagemNota, .iTabelaPreco, .iTipoNFiscal, .sNaturezaOp, .dPesoLiq, .dPesoBruto, .lNumIntTrib, .sPlaca, .sPlacaUF, .lVolumeQuant, .lVolumeEspecie, .lVolumeMarca, .iCanal, .lNumIntNotaOriginal, .iCodTranspRedesp, .iDetPagFrete, .iMoeda, .dTaxaMoeda, .dVolumeTotal, .iSemDataSaida, .iNaoImpCobranca, .sMensagemCorpoNota, .iRecibo, _
            objNFiscalOriginal.iFilialEmpresa, objNFiscalOriginal.lNumNotaFiscal, objNFiscalOriginal.sSerie, DOCINFO_NF_INT_SAIDA, STATUS_EXCLUIDO)
        If lErro <> AD_SQL_SUCESSO Then Error 62142

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 62143

        If lErro = AD_SQL_SEM_DADOS Then Error 62144

        objNFiscalOriginal.dPesoBruto = .dPesoBruto
        objNFiscalOriginal.dPesoLiq = .dPesoLiq
        objNFiscalOriginal.dtDataEmissao = .dtDataEmissao
        objNFiscalOriginal.dtDataEntrada = .dtDataEntrada
        objNFiscalOriginal.dtDataSaida = .dtDataSaida
        objNFiscalOriginal.dtDataVencimento = .dtDataVencimento
        objNFiscalOriginal.dValorDesconto = .dValorDesconto
        objNFiscalOriginal.dValorFrete = .dValorFrete
        objNFiscalOriginal.dValorOutrasDespesas = .dValorOutrasDespesas
        objNFiscalOriginal.dValorProdutos = .dValorProdutos
        objNFiscalOriginal.dValorSeguro = .dValorSeguro
        objNFiscalOriginal.dValorTotal = .dValorTotal
        objNFiscalOriginal.iCanal = .iCanal
        objNFiscalOriginal.iClasseDocCPR = .iClasseDocCPR
        objNFiscalOriginal.iCodTransportadora = .iCodTransportadora
        objNFiscalOriginal.iCodTranspRedesp = .iCodTranspRedesp
        objNFiscalOriginal.iDetPagFrete = .iDetPagFrete
        objNFiscalOriginal.iFilialCli = .iFilialCli
        objNFiscalOriginal.iFilialEntrega = .iFilialEntrega
        objNFiscalOriginal.iFilialForn = .iFilialForn
        objNFiscalOriginal.iFilialPedido = .iFilialPedido
        objNFiscalOriginal.iStatus = .iStatus
        objNFiscalOriginal.iTabelaPreco = .iTabelaPreco
        objNFiscalOriginal.iTipoNFiscal = .iTipoNFiscal
        objNFiscalOriginal.lVolumeQuant = .lVolumeQuant
        objNFiscalOriginal.lCliente = .lCliente
        objNFiscalOriginal.lFornecedor = .lFornecedor
        objNFiscalOriginal.lNumIntDocCPR = .lNumIntDocCPR
        objNFiscalOriginal.lNumIntNotaOriginal = .lNumIntNotaOriginal
        objNFiscalOriginal.lNumIntTrib = .lNumIntTrib
        objNFiscalOriginal.lNumNotaFiscal = .lNumNotaFiscal
        objNFiscalOriginal.sNumPedidoTerc = .sNumPedidoTerc
        objNFiscalOriginal.lNumPedidoVenda = .lNumPedidoVenda
        objNFiscalOriginal.sMensagemNota = .sMensagemNota
        objNFiscalOriginal.sNaturezaOp = .sNaturezaOp
        objNFiscalOriginal.sPlaca = .sPlaca
        objNFiscalOriginal.sPlacaUF = .sPlacaUF
        objNFiscalOriginal.sSerie = .sSerie
        objNFiscalOriginal.lVolumeEspecie = .lVolumeEspecie
        objNFiscalOriginal.lVolumeMarca = .lVolumeMarca
        objNFiscalOriginal.lNumIntDoc = .lNumIntDoc
        objNFiscalOriginal.iFilialEmpresa = .iFilialEmpresa
        objNFiscalOriginal.iMoeda = .iMoeda
        objNFiscalOriginal.dTaxaMoeda = .dTaxaMoeda
        objNFiscalOriginal.dVolumeTotal = .dVolumeTotal
        objNFiscalOriginal.iSemDataSaida = .iSemDataSaida
        objNFiscalOriginal.sMensagemCorpoNota = .sMensagemCorpoNota
        objNFiscalOriginal.iRecibo = .iRecibo

    End With

    Call Comando_Fechar(lComando)

    NFiscalInternaSaida_Le_Numero2 = SUCESSO

    Exit Function

Erro_NFiscalInternaSaida_Le_Numero2:

    NFiscalInternaSaida_Le_Numero2 = Err

    Select Case Err

        Case 62141
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 62142, 62143
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", Err)

        Case 62144 'Tratado pela rotina chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149168)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function CRFatConfig_Le(ByVal sCodigo As String, ByVal iFilialEmpresa As Integer, sConteudo As String) As Long
'Lê o Conteudo da Tabela de CRFatConfig de acordo com o código passado

Dim lErro As Long
Dim objCRFATConfig As New ClassCRFATConfig
Dim lComando As Long

On Error GoTo Erro_CRFatConfig_Le
    
    'Abrir comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 61453

    objCRFATConfig.sCodigo = sCodigo
    objCRFATConfig.iFilialEmpresa = iFilialEmpresa
    
    lErro = CRFATConfig_Le_Batch(objCRFATConfig, lComando)
    If lErro <> SUCESSO And lErro <> 33707 Then Error 61454

    'Não foi encontrado registro em CRConfig
    If lErro = 33707 Then Error 61455
    
    sConteudo = objCRFATConfig.sConteudo
    
    'Fechar comando
    Call Comando_Fechar(lComando)

    CRFatConfig_Le = SUCESSO
    
    Exit Function

Erro_CRFatConfig_Le:
        
    CRFatConfig_Le = Err
    
    Select Case Err
    
        Case 61453
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 61454 'Tratado na Rotina chamada

        Case 61455
            Call Rotina_Erro(vbOKOnly, "ERRO_CRFATCONFIG_INEXISTENTE", Err, objCRFATConfig.sCodigo, objCRFATConfig.iFilialEmpresa)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149169)
     
    End Select
     
    'Fechar comando
    Call Comando_Fechar(lComando)
 
    Exit Function
    
End Function

Public Function Cliente_Le_Codigo_NomeReduzido(ByVal objCliente As ClassCliente) As Long
'obtem o codigo do cliente à partir de seu nome reduzido

Dim lErro As Long
Dim lComando As Long, lCodigo As Long

On Error GoTo Erro_Cliente_Le_Codigo_NomeReduzido

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 59270

    If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL Then
    
        lErro = Comando_Executar(lComando, "SELECT CodigoLoja FROM Clientes WHERE NomeReduzido = ?", _
            lCodigo, objCliente.sNomeReduzido)
        
    Else
    
        lErro = Comando_Executar(lComando, "SELECT Codigo FROM Clientes WHERE NomeReduzido = ?", _
            lCodigo, objCliente.sNomeReduzido)
    
    End If
    If lErro <> AD_SQL_SUCESSO Then Error 59271

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 59272

    If lErro = AD_SQL_SEM_DADOS Then Error 59273

    If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL Then
        objCliente.lCodigoLoja = lCodigo
    Else
        objCliente.lCodigo = lCodigo
    End If
    
    
    Call Comando_Fechar(lComando)

    Cliente_Le_Codigo_NomeReduzido = SUCESSO
     
    Exit Function
    
Erro_Cliente_Le_Codigo_NomeReduzido:

    Cliente_Le_Codigo_NomeReduzido = Err
     
    Select Case Err
          
        Case 59270
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 59271, 59272
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLIENTES", Err)

        Case 59273
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO_2", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149170)
     
    End Select
     
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function ComprasConfig_Le(ByVal objConfiguraCOM As ClassConfiguraCOM) As Long

Dim lErro As Long
Dim sConteudo As String
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim objComprasConfig As ClassComprasConfig
Dim colComprasConfig As New colComprasConfig

On Error GoTo Erro_ComprasConfig_Le

    'Abre os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 49356
    Next

    lErro = CF("ConfiguraCOM_MoverCampos_ColComprasConfig", objConfiguraCOM, colComprasConfig)
    If lErro <> SUCESSO Then Error 49359

    For Each objComprasConfig In colComprasConfig

        sConteudo = String(STRING_CONTEUDO, 0)

        'Seleciona tipo e conteudo na tabela ComprasConfig de acordo com o codigo passado como parametro
        lErro = Comando_Executar(alComando(0), "SELECT Conteudo FROM ComprasConfig WHERE Codigo = ? And FilialEmpresa = ?", sConteudo, objComprasConfig.sCodigo, objComprasConfig.iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then Error 49354

        lErro = Comando_BuscarPrimeiro(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 49355

        'Nao encontrou registro em ComprasConfig
        If lErro = AD_SQL_SEM_DADOS Then
        
            'trata valores que podem só estar disponiveis apos a inicializacao do modulo
            
            Select Case objComprasConfig.sCodigo
            
                Case "NFISCAL_DIFERENTE_PED_COMPRA"
                    objComprasConfig.iTipo = 1
                    objComprasConfig.sConteudo = "0"
            
                Case Else
                    Error 56200
                
            End Select
            
        End If
        
        'Preenche objComprasConfig com os dados obtidos no BD
        If lErro = AD_SQL_SUCESSO Then

            objComprasConfig.sConteudo = sConteudo

        End If

    Next

    'Preenche objeto objConfiguraCOM com os dados lidos do BD
    Call CF("ColComprasConfig_MoverCampos_ConfiguraCOM", objConfiguraCOM, colComprasConfig)

    'Fecha os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next

    ComprasConfig_Le = SUCESSO

    Exit Function

Erro_ComprasConfig_Le:

    ComprasConfig_Le = Err

    Select Case Err

        Case 49354, 49355
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMPRASCONFIG", Err, objComprasConfig.sCodigo)

        Case 49356
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 49357
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 49358
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case 49359
            'Erro tratado na rotina chamada
        Case 56200
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REGISTRO_COMPRAS_CONFIG_NAO_ENCONTRADO", Err, objComprasConfig.sCodigo, objComprasConfig.iFilialEmpresa)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149171)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function ConfiguraCOM_MoverCampos_ColComprasConfig(ByVal objConfiguraCOM As ClassConfiguraCOM, ByVal colComprasConfig As colComprasConfig) As Long
'Transmite os dados do objConfiguraCOM para ColConfiguraCOM

Dim lErro As Long

On Error GoTo Erro_Configura_COM_MoverCampos_ColComprasConfig

    colComprasConfig.Add "COMPRADOR_AUMENTA_QUANT", EMPRESA_TODA, "", 0, CStr(objConfiguraCOM.iCompradorAumentaQuant)
    colComprasConfig.Add "CONSIDERA_QUANT_COTACAO_ANTERIOR", EMPRESA_TODA, "", 0, CStr(objConfiguraCOM.iConsideraQuantCotacaoAnterior)
    colComprasConfig.Add "CONSUMO_MEDIO_MAX", EMPRESA_TODA, "", 0, CStr(objConfiguraCOM.dConsumoMedioMax)
    colComprasConfig.Add "CONTROLE_ALCADA", EMPRESA_TODA, "", 0, CStr(objConfiguraCOM.iControleAlcada)
    colComprasConfig.Add "FILIAL_COMPRA_DEFAULT", EMPRESA_TODA, "", 0, CStr(objConfiguraCOM.iFilialCompra)
    colComprasConfig.Add "MESES_CONSUMO_MEDIO", EMPRESA_TODA, "", 0, CStr(objConfiguraCOM.iMesesConsumoMedio)
    colComprasConfig.Add "MESES_MEDIA_TEMPO_RESSUP", EMPRESA_TODA, "", 0, CStr(objConfiguraCOM.iMesesMediaTempoRessup)
    colComprasConfig.Add "PERCENT_MAIS_QUANT_COTACAO_ANTERIOR", EMPRESA_TODA, "", 0, CStr(objConfiguraCOM.dPercentMaisQuantCotacaoAnterior)
    colComprasConfig.Add "PERCENT_MAIS_RECEB", EMPRESA_TODA, "", 0, CStr(objConfiguraCOM.dPercentMaisReceb)
    colComprasConfig.Add "PERCENT_MENOS_QUANT_COTACAO_ANTERIOR", EMPRESA_TODA, "", 0, CStr(objConfiguraCOM.dPercentMenosQuantCotacaoAnterior)
    colComprasConfig.Add "PERCENT_MENOS_RECEB", EMPRESA_TODA, "", 0, CStr(objConfiguraCOM.dPercentMenosReceb)
    colComprasConfig.Add "RECEB_FORA_FAIXA", EMPRESA_TODA, "", 0, CStr(objConfiguraCOM.iRecebForaFaixa)
    colComprasConfig.Add "RESIDUO", EMPRESA_TODA, "", 0, CStr(objConfiguraCOM.dResiduo)
    colComprasConfig.Add "TAXA_FINANCEIRA_EMPRESA", EMPRESA_TODA, "", 0, CStr(objConfiguraCOM.dTaxaFinanceiraEmpresa)
    colComprasConfig.Add "TEM_FAIXA_RECEB", EMPRESA_TODA, "", 0, CStr(objConfiguraCOM.iTemFaixaReceb)
    colComprasConfig.Add "TEMPO_RESSUP_MAX", EMPRESA_TODA, "", 0, CStr(objConfiguraCOM.dTempoRessupMax)
    colComprasConfig.Add "NFISCAL_DIFERENTE_PED_COMPRA", giFilialEmpresa, "", 0, CStr(objConfiguraCOM.iNFDiferentePC)
    colComprasConfig.Add "NUM_COMPRAS_MEDIA_ATRASO", EMPRESA_TODA, "", 0, CStr(objConfiguraCOM.iNumComprasMediaAtraso)
    colComprasConfig.Add "NUM_COMPRAS_TEMPO_RESSUP", EMPRESA_TODA, "", 0, CStr(objConfiguraCOM.iNumComprasTempoRessup)

    colComprasConfig.Add "FORMATO_PRECO_UNITARIO", EMPRESA_TODA, "", 0, CStr(objConfiguraCOM.sFormatoPrecoUnitario)
    colComprasConfig.Add "EMAIL_ALTERA_DATA_ENVIO", EMPRESA_TODA, "", 0, CStr(objConfiguraCOM.iEmailAlteraDataEnvioPC)
    colComprasConfig.Add "REQCOMPRA_BLOQUEIA_DESCPROD", EMPRESA_TODA, "", 0, CStr(objConfiguraCOM.iReqCompraDescProdBloq)
    colComprasConfig.Add "PEDCOMPRA_BLOQUEIA_ENVIO_SEM_APROV", EMPRESA_TODA, "", 0, CStr(objConfiguraCOM.iPedCompraBloqEnvioSemAprov)
    
    colComprasConfig.Add "PEDCOMPRA_EXIBE_DELIVERYDATE", EMPRESA_TODA, "", 0, CStr(objConfiguraCOM.iPCExibeDeliveryDate)
    
    ConfiguraCOM_MoverCampos_ColComprasConfig = SUCESSO
    
    Exit Function

Erro_Configura_COM_MoverCampos_ColComprasConfig:

    ConfiguraCOM_MoverCampos_ColComprasConfig = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149172)

    End Select

    Exit Function

End Function

Function ColComprasConfig_MoverCampos_ConfiguraCOM(ByVal objConfiguraCOM As ClassConfiguraCOM, ByVal colComprasConfig As colComprasConfig) As Long
'Preenche objConfiguraCOM com os dados de objComprasConfig

Dim objComprasConfig As ClassComprasConfig
Dim lErro As Long

On Error GoTo Erro_ColComprasConfig_MoverCampos_ConfiguraCOM

    For Each objComprasConfig In colComprasConfig

        Select Case objComprasConfig.sCodigo

            Case "COMPRADOR_AUMENTA_QUANT"
                objConfiguraCOM.iCompradorAumentaQuant = StrParaInt(objComprasConfig.sConteudo)

            Case "CONSIDERA_QUANT_COTACAO_ANTERIOR"
                objConfiguraCOM.iConsideraQuantCotacaoAnterior = StrParaInt(objComprasConfig.sConteudo)

            Case "CONSUMO_MEDIO_MAX"
                objConfiguraCOM.dConsumoMedioMax = StrParaDbl(objComprasConfig.sConteudo)

            Case "CONTROLE_ALCADA"
                objConfiguraCOM.iControleAlcada = StrParaInt(objComprasConfig.sConteudo)

            Case "FILIAL_COMPRA_DEFAULT"
                objConfiguraCOM.iFilialCompra = StrParaInt(objComprasConfig.sConteudo)

            Case "MESES_CONSUMO_MEDIO"
                objConfiguraCOM.iMesesConsumoMedio = StrParaInt(objComprasConfig.sConteudo)

            Case "MESES_MEDIA_TEMPO_RESSUP"
                objConfiguraCOM.iMesesMediaTempoRessup = StrParaInt(objComprasConfig.sConteudo)

            Case "PERCENT_MAIS_QUANT_COTACAO_ANTERIOR"
                objConfiguraCOM.dPercentMaisQuantCotacaoAnterior = StrParaDbl(objComprasConfig.sConteudo)

            Case "PERCENT_MAIS_RECEB"
                objConfiguraCOM.dPercentMaisReceb = StrParaDbl(objComprasConfig.sConteudo)

            Case "PERCENT_MENOS_QUANT_COTACAO_ANTERIOR"
                objConfiguraCOM.dPercentMenosQuantCotacaoAnterior = StrParaDbl(objComprasConfig.sConteudo)

            Case "PERCENT_MENOS_RECEB"
                objConfiguraCOM.dPercentMenosReceb = StrParaDbl(objComprasConfig.sConteudo)

            Case "RECEB_FORA_FAIXA"
                objConfiguraCOM.iRecebForaFaixa = StrParaInt(objComprasConfig.sConteudo)

            Case "RESIDUO"
                objConfiguraCOM.dResiduo = StrParaDbl(objComprasConfig.sConteudo)

            Case "TAXA_FINANCEIRA_EMPRESA"
                objConfiguraCOM.dTaxaFinanceiraEmpresa = StrParaDbl(objComprasConfig.sConteudo)

            Case "TEM_FAIXA_RECEB"
                objConfiguraCOM.iTemFaixaReceb = StrParaInt(objComprasConfig.sConteudo)

            Case "TEMPO_RESSUP_MAX"
                objConfiguraCOM.dTempoRessupMax = StrParaDbl(objComprasConfig.sConteudo)
            
            Case "NFISCAL_DIFERENTE_PED_COMPRA"
                objConfiguraCOM.iNFDiferentePC = StrParaInt(objComprasConfig.sConteudo)

            Case "NUM_COMPRAS_TEMPO_RESSUP"
                objConfiguraCOM.iNumComprasTempoRessup = StrParaInt(objComprasConfig.sConteudo)
                
            Case "NUM_COMPRAS_MEDIA_ATRASO"
                objConfiguraCOM.iNumComprasMediaAtraso = StrParaInt(objComprasConfig.sConteudo)

            Case "FORMATO_PRECO_UNITARIO"
                objConfiguraCOM.sFormatoPrecoUnitario = objComprasConfig.sConteudo
                
            Case "EMAIL_ALTERA_DATA_ENVIO"
                objConfiguraCOM.iEmailAlteraDataEnvioPC = StrParaInt(objComprasConfig.sConteudo)

            Case "REQCOMPRA_BLOQUEIA_DESCPROD"
                objConfiguraCOM.iReqCompraDescProdBloq = StrParaInt(objComprasConfig.sConteudo)
            
            Case "PEDCOMPRA_BLOQUEIA_ENVIO_SEM_APROV"
                objConfiguraCOM.iPedCompraBloqEnvioSemAprov = StrParaInt(objComprasConfig.sConteudo)
            
            Case "PEDCOMPRA_EXIBE_DELIVERYDATE"
                objConfiguraCOM.iPCExibeDeliveryDate = StrParaInt(objComprasConfig.sConteudo)
            
        End Select

    Next

    ColComprasConfig_MoverCampos_ConfiguraCOM = SUCESSO
    
    Exit Function

Erro_ColComprasConfig_MoverCampos_ConfiguraCOM:

    ColComprasConfig_MoverCampos_ConfiguraCOM = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149173)

    End Select

    Exit Function

End Function

Function NFiscal_Le_1(ByVal objNFiscal As ClassNFiscal) As Long
'Verifica se foi cadastrada uma Nota Fiscal no BD a partir de FilialEmpresa, Serie, Número, Tipo, Data Emissão

Dim lErro As Long
Dim lComando As Long
Dim lNumIntDoc As Long
Dim tNFiscal As typeNFiscal

On Error GoTo Erro_NFiscal_Le_1

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 83968

    With tNFiscal

        .sMensagemNota = String(STRING_NFISCAL_MENSAGEM, 0)
        .sMensagemCorpoNota = String(STRING_NFISCAL_MENSAGEM, 0)
        .sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
        .sPlaca = String(STRING_NFISCAL_PLACA, 0)
        .sPlacaUF = String(STRING_NFISCAL_PLACA_UF, 0)
        .sSerie = String(STRING_SERIE, 0)
        .sNumPedidoTerc = String(STRING_NUM_PEDIDO_TERC, 0)
        .sVolumeNumero = String(STRING_NFISCAL_VOLUME_NUMERO, 0)
        .sObservacao = String(STRING_NFISCAL_OBSERVACAO, 0)

        'Removido por Luiz Nogueira em 21/08/03
        '.lVolumeEspecie = String(STRING_NFISCAL_VOLUME_ESPECIE, 0)
        '.lVolumeMarca = String(STRING_NFISCAL_VOLUME_MARCA, 0)

        'Pesquisa a Nota Fiscal com os dados passados
        lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Status, Serie, NumNotaFiscal, Cliente, FilialCli, FilialEntrega, Fornecedor, FilialForn, DataEmissao, DataEntrada, DataSaida, DataVencimento,DataReferencia, NumPedidoVenda, NumPedidoTerc, ClasseDocCPR, NumIntDocCPR, ValorTotal, ValorProdutos, ValorFrete, ValorSeguro, ValorOutrasDespesas, ValorDesconto, CodTransportadora, MensagemNota, TabelaPreco, TipoNFiscal, NaturezaOp, PesoLiq, PesoBruto, NumIntTrib, Placa, PlacaUF, VolumeQuant, VolumeEspecie, VolumeMarca, Canal, NumIntNotaOriginal, VolumeNumero, FreteRespons, FilialPedido, Observacao, HoraEntrada, HoraSaida, CodTranspRedesp, DetPagFrete, Moeda, TaxaMoeda, VolumeTotal, SemDataSaida, NaoImpCobranca, MensagemCorpoNota, NFe FROM NFiscal WHERE FilialEmpresa = ? AND Serie = ? AND NumNotaFiscal = ? AND TipoNFiscal = ? AND DataEmissao = ? AND Cliente = ? AND FilialCli = ? AND Fornecedor=? AND FilialForn=?", _
            .lNumIntDoc, .iStatus, .sSerie, .lNumNotaFiscal, .lCliente, .iFilialCli, .iFilialEntrega, .lFornecedor, .iFilialForn, .dtDataEmissao, .dtDataEntrada, .dtDataSaida, .dtDataVencimento, .dtDataReferencia, .lNumPedidoVenda, .sNumPedidoTerc, .iClasseDocCPR, .lNumIntDocCPR, .dValorTotal, .dValorProdutos, .dValorFrete, .dValorSeguro, .dValorOutrasDespesas, .dValorDesconto, .iCodTransportadora, .sMensagemNota, .iTabelaPreco, .iTipoNFiscal, .sNaturezaOp, .dPesoLiq, .dPesoBruto, .lNumIntTrib, .sPlaca, .sPlacaUF, .lVolumeQuant, .lVolumeEspecie, .lVolumeMarca, .iCanal, .lNumIntNotaOriginal, .sVolumeNumero, .iFreteRespons, .iFilialPedido, .sObservacao, .dHoraEntrada, .dHoraSaida, .iCodTranspRedesp, .iDetPagFrete, .iMoeda, .dTaxaMoeda, .dVolumeTotal, .iSemDataSaida, .iNaoImpCobranca, .sMensagemCorpoNota, .iNFe, _
            objNFiscal.iFilialEmpresa, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.iTipoNFiscal, objNFiscal.dtDataEmissao, objNFiscal.lCliente, objNFiscal.iFilialCli, objNFiscal.lFornecedor, objNFiscal.iFilialForn)
        If lErro <> AD_SQL_SUCESSO Then gError 83969
    
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 83970
        
        'Se não encontrar --> Erro
        If lErro = AD_SQL_SEM_DADOS Then gError 83971

        'Preenche objNFiscal com os dados lidos
        objNFiscal.dPesoBruto = .dPesoBruto
        objNFiscal.dPesoLiq = .dPesoLiq
        objNFiscal.dtDataEmissao = .dtDataEmissao
        objNFiscal.dtDataEntrada = .dtDataEntrada
        objNFiscal.dtDataSaida = .dtDataSaida
        objNFiscal.dtDataVencimento = .dtDataVencimento
        objNFiscal.dValorDesconto = .dValorDesconto
        objNFiscal.dValorFrete = .dValorFrete
        objNFiscal.dValorOutrasDespesas = .dValorOutrasDespesas
        objNFiscal.dValorProdutos = .dValorProdutos
        objNFiscal.dValorSeguro = .dValorSeguro
        objNFiscal.dValorTotal = .dValorTotal
        objNFiscal.iCanal = .iCanal
        objNFiscal.iClasseDocCPR = .iClasseDocCPR
        objNFiscal.iCodTransportadora = .iCodTransportadora
        objNFiscal.iCodTranspRedesp = .iCodTranspRedesp
        objNFiscal.iDetPagFrete = .iDetPagFrete
        objNFiscal.iFilialCli = .iFilialCli
'        objNFiscal.iFilialEmpresa = giFilialEmpresa
        objNFiscal.iFilialEntrega = .iFilialEntrega
        objNFiscal.iFilialForn = .iFilialForn
        objNFiscal.iStatus = .iStatus
        objNFiscal.iTabelaPreco = .iTabelaPreco
        objNFiscal.iTipoNFiscal = .iTipoNFiscal
        objNFiscal.lVolumeQuant = .lVolumeQuant
        objNFiscal.lCliente = .lCliente
        objNFiscal.lFornecedor = .lFornecedor
        objNFiscal.lNumIntDoc = .lNumIntDoc
        objNFiscal.lNumIntDocCPR = .lNumIntDocCPR
        objNFiscal.lNumIntNotaOriginal = .lNumIntNotaOriginal
        objNFiscal.lNumIntTrib = .lNumIntTrib
        objNFiscal.lNumNotaFiscal = .lNumNotaFiscal
        objNFiscal.sNumPedidoTerc = .sNumPedidoTerc
        objNFiscal.lNumPedidoVenda = .lNumPedidoVenda
        objNFiscal.sMensagemNota = .sMensagemNota
        objNFiscal.sNaturezaOp = .sNaturezaOp
        objNFiscal.sPlaca = .sPlaca
        objNFiscal.sPlacaUF = .sPlacaUF
        objNFiscal.sSerie = .sSerie
        objNFiscal.lVolumeEspecie = .lVolumeEspecie
        objNFiscal.lVolumeMarca = .lVolumeMarca
        objNFiscal.sVolumeNumero = .sVolumeNumero
        objNFiscal.iFreteRespons = .iFreteRespons
        objNFiscal.dtDataReferencia = .dtDataReferencia
        objNFiscal.iFilialPedido = .iFilialPedido
        objNFiscal.sObservacao = .sObservacao
        objNFiscal.dtHoraEntrada = .dHoraEntrada
        objNFiscal.dtHoraSaida = .dHoraSaida
        objNFiscal.iMoeda = .iMoeda
        objNFiscal.dTaxaMoeda = .dTaxaMoeda
        objNFiscal.dVolumeTotal = .dVolumeTotal
        objNFiscal.iSemDataSaida = .iSemDataSaida
        objNFiscal.iNaoImpCobranca = .iNaoImpCobranca
        objNFiscal.sMensagemCorpoNota = .sMensagemCorpoNota
        objNFiscal.iNFe = .iNFe

    End With

    'Fechar o Comando
    Call Comando_Fechar(lComando)

    NFiscal_Le_1 = SUCESSO

    Exit Function

Erro_NFiscal_Le_1:

    NFiscal_Le_1 = gErr

    Select Case gErr

        Case 83968
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 83969, 83970
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL2", gErr, objNFiscal.iTipoNFiscal, objNFiscal.lCliente, objNFiscal.iFilialForn, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal)

        Case 83971 'Nota Fiscal não cadastrada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149174)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Pesquisa_Ultima_Devolucao_Forn(ByVal objFilialFornecedorEst As ClassFilialFornecedorEst) As Long
'pesquisa a ultima nota fiscal de Devolução de Devolução da FilialEmpresa (data de emissao mais recente) para o Fornecedor em questão.

Dim lErro As Long
Dim alComando(1 To 1) As Long
Dim dtDataEmissao As Date
Dim iIndice As Integer
Dim sSQL As String

On Error GoTo Erro_Pesquisa_Ultima_Devolucao_Forn

    'Abre os Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 89406
    Next
    
    sSQL = "SELECT NFiscal.DataEmissao FROM NFiscal, TiposDocInfo WHERE NFiscal.TipoNFiscal = TiposDocInfo.Codigo AND TiposDocInfo.Codigo = ? AND NFiscal.Fornecedor = ? And NFiscal.FilialEmpresa = ? AND NFiscal.Status <> ? ORDER BY NFiscal.DataEmissao DESC"
    
    lErro = Comando_Executar(alComando(1), sSQL, dtDataEmissao, DOCINFO_NFISDCM, objFilialFornecedorEst.lCodFornecedor, objFilialFornecedorEst.iFilialEmpresa, STATUS_CANCELADO)
    If lErro <> AD_SQL_SUCESSO Then gError 89407
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89408
    
    If lErro = AD_SQL_SUCESSO Then
        objFilialFornecedorEst.dtDataUltDevolucao = dtDataEmissao
    Else
        objFilialFornecedorEst.dtDataUltDevolucao = DATA_NULA
    End If
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Pesquisa_Ultima_Devolucao_Forn = SUCESSO
    
    Exit Function
    
Erro_Pesquisa_Ultima_Devolucao_Forn:

    Pesquisa_Ultima_Devolucao_Forn = gErr
    
    Select Case gErr
        
        Case 89406
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 89407, 89408
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PESQUISA_ULT_DEVOLUCAO_FORN", gErr, objFilialFornecedorEst.lCodFornecedor)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149175)
        
    End Select
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Function Pesquisa_Ultima_Devolucao_FornFil(ByVal objFilialFornecedorEst As ClassFilialFornecedorEst) As Long
'pesquisa a ultima nota fiscal de Devolução de Compra da FilialEmpresa (data de emissao mais recente) para o Fornecedor/FilialForn em questão.

Dim lErro As Long
Dim alComando(1 To 1) As Long
Dim dtDataEmissao As Date
Dim iIndice As Integer
Dim sSQL As String

On Error GoTo Erro_Pesquisa_Ultima_Devolucao_FornFil

    'Abre os Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 89409
    Next
    
    sSQL = "SELECT NFiscal.DataEmissao FROM NFiscal, TiposDocInfo WHERE NFiscal.TipoNFiscal = TiposDocInfo.Codigo AND TiposDocInfo.Codigo = ? AND NFiscal.Fornecedor = ? And NFiscal.FilialForn = ? AND NFiscal.FilialEmpresa = ? AND NFiscal.Status <> ? ORDER BY NFiscal.DataEmissao DESC"
    
    lErro = Comando_Executar(alComando(1), sSQL, dtDataEmissao, DOCINFO_NFISDCM, objFilialFornecedorEst.lCodFornecedor, objFilialFornecedorEst.iCodFilial, objFilialFornecedorEst.iFilialEmpresa, STATUS_CANCELADO)
    If lErro <> AD_SQL_SUCESSO Then gError 89410
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89411
    
    If lErro = AD_SQL_SUCESSO Then
        objFilialFornecedorEst.dtDataUltDevolucao = dtDataEmissao
    Else
        objFilialFornecedorEst.dtDataUltDevolucao = DATA_NULA
    End If
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Pesquisa_Ultima_Devolucao_FornFil = SUCESSO
    
    Exit Function
    
Erro_Pesquisa_Ultima_Devolucao_FornFil:

    Pesquisa_Ultima_Devolucao_FornFil = gErr
    
    Select Case gErr
        
        Case 89409
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 89410, 89411
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PESQUISA_ULT_DEVOLUCAO_FORNFIL", gErr, objFilialFornecedorEst.lCodFornecedor, objFilialFornecedorEst.iCodFilial)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149176)
        
    End Select
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Public Function NFiscal_Le_SemNumIntDoc(ByVal objNFiscal As ClassNFiscal) As Long
'Lê os dados da nota fiscal a partir dos dados Numero, Serie, FilialEmpresa, Tipo, Fornecedor ou Cliente, FilialForn ou FilialCli, DataEmissao, DataEntrada

Dim lErro As Long
Dim lComando As Long
Dim sSelecaoSQL As String
Dim tNFiscal As typeNFiscal

On Error GoTo Erro_NFiscal_Le_SemNumIntDoc

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 79711
        
    With tNFiscal
        
        'Inicializa a string que receberá a Série
        .sNumPedidoTerc = String(STRING_NUM_PEDIDO_TERC, 0)
        .sMensagemNota = String(STRING_NFISCAL_MENSAGEM, 0)
        .sMensagemCorpoNota = String(STRING_NFISCAL_MENSAGEM, 0)
        .sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
        .sPlaca = String(STRING_NFISCAL_PLACA, 0)
        .sPlacaUF = String(STRING_NFISCAL_PLACA_UF, 0)
        .sVolumeNumero = String(STRING_NFISCAL_VOLUME_NUMERO, 0)
        .sObservacao = String(STRING_NFISCAL_OBSERVACAO, 0)
        .sCodUsuarioCancel = String(STRING_USUARIO_CODIGO, 0)
        .sMotivoCancel = String(STRING_NFISCAL_MOTIVOCANCEL, 0)
        
        If objNFiscal.iFilialEmpresa = 0 Then objNFiscal.iFilialEmpresa = giFilialEmpresa
        
        'Removido por Luiz Nogueira em 21/08/03
        '.lVolumeEspecie = String(STRING_NFISCAL_VOLUME_ESPECIE, 0)
        '.lVolumeMarca = String(STRING_NFISCAL_VOLUME_MARCA, 0)
        
        'Define o comando SQL que será passado no select
        sSelecaoSQL = "SELECT NumIntDoc, Status, FilialEmpresa, FilialEntrega, DataVencimento, DataReferencia, FilialPedido, NumPedidoVenda, NumPedidoTerc, ClasseDocCPR, NumIntDocCPR, ValorTotal, ValorProdutos, ValorFrete, ValorSeguro, ValorOutrasDespesas, ValorDesconto, CodTransportadora, MensagemNota, TabelaPreco, NaturezaOp, PesoLiq, PesoBruto, NumIntTrib, Placa, PlacaUF, VolumeQuant, VolumeEspecie, VolumeMarca, VolumeNumero, Canal, NumIntNotaOriginal, ClienteBenef, FilialCliBenef, FornecedorBenef, FilialFornBenef, FreteRespons, NumRecebimento, Observacao, CodUsuarioCancel, MotivoCancel, CodTranspRedesp, DetPagFrete, Moeda, TaxaMoeda, VolumeTotal, SemDataSaida, NaoImpCobranca, MensagemCorpoNota, NFe FROM NFiscal WHERE NumNotaFiscal = ? AND FilialEmpresa = ? AND Serie = ? AND Fornecedor = ? AND Cliente = ? AND FilialForn = ? AND FilialCli = ? AND DataEmissao = ? AND DataEntrada = ? AND TipoNFiscal = ? AND Status <> ?"
        
        'Busca no BD os campos necessários para se definir a tela e o doc que será exibido
        lErro = Comando_Executar(lComando, sSelecaoSQL, .lNumIntDoc, .iStatus, .iFilialEmpresa, .iFilialEntrega, .dtDataVencimento, .dtDataReferencia, .iFilialPedido, .lNumPedidoVenda, .sNumPedidoTerc, .iClasseDocCPR, .lNumIntDocCPR, .dValorTotal, .dValorProdutos, .dValorFrete, .dValorSeguro, .dValorOutrasDespesas, .dValorDesconto, .iCodTransportadora, .sMensagemNota, .iTabelaPreco, .sNaturezaOp, .dPesoLiq, .dPesoBruto, .lNumIntTrib, .sPlaca, .sPlacaUF, .lVolumeQuant, .lVolumeEspecie, .lVolumeMarca, .sVolumeNumero, .iCanal, .lNumIntNotaOriginal, .lClienteBenef, .iFilialCliBenef, .lFornecedorBenef, .iFilialFornBenef, .iFreteRespons, .lNumRecebimento, .sObservacao, .sCodUsuarioCancel, .sMotivoCancel, .iCodTranspRedesp, .iDetPagFrete, .iMoeda, .dTaxaMoeda, .dVolumeTotal, .iSemDataSaida, .iNaoImpCobranca, .sMensagemCorpoNota, .iNFe, _
            objNFiscal.lNumNotaFiscal, objNFiscal.iFilialEmpresa, objNFiscal.sSerie, objNFiscal.lFornecedor, objNFiscal.lCliente, objNFiscal.iFilialForn, objNFiscal.iFilialCli, objNFiscal.dtDataEmissao, objNFiscal.dtDataEntrada, objNFiscal.iTipoDocInfo, STATUS_CANCELADO)
        If lErro <> AD_SQL_SUCESSO Then gError 79712
        
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 79713
        
        'Se não encontrou a NFiscal = > erro
        If lErro = AD_SQL_SEM_DADOS Then gError 79714
    
    End With
        
    'Guarda no objNFiscal os dados obtidos no select
    Call Move_NFiscal_Obj(objNFiscal, tNFiscal)
        
    Call Comando_Fechar(lComando)
    
    NFiscal_Le_SemNumIntDoc = SUCESSO
    
    Exit Function

Erro_NFiscal_Le_SemNumIntDoc:

    NFiscal_Le_SemNumIntDoc = gErr
    
    Select Case gErr
    
        Case 79711
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 79712, 79713
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", gErr)
        
        Case 79714
        'Sem dados
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149177)
    
    End Select
        
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Private Sub Move_NFiscal_Obj(ByVal objNFiscal As ClassNFiscal, tNFiscal As typeNFiscal)

    With tNFiscal
    
        objNFiscal.lNumIntDoc = .lNumIntDoc
        objNFiscal.iStatus = .iStatus
        objNFiscal.iFilialEmpresa = .iFilialEmpresa
        objNFiscal.iFilialEntrega = .iFilialEntrega
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
        objNFiscal.dValorOutrasDespesas = .dValorOutrasDespesas
        objNFiscal.dValorDesconto = .dValorDesconto
        objNFiscal.iCodTransportadora = .iCodTransportadora
        objNFiscal.iCodTranspRedesp = .iCodTranspRedesp
        objNFiscal.iDetPagFrete = .iDetPagFrete
        objNFiscal.sMensagemNota = .sMensagemNota
        objNFiscal.iTabelaPreco = .iTabelaPreco
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
        objNFiscal.iMoeda = .iMoeda
        objNFiscal.dTaxaMoeda = .dTaxaMoeda
        objNFiscal.dVolumeTotal = .dVolumeTotal
        objNFiscal.iSemDataSaida = .iSemDataSaida
        objNFiscal.iNaoImpCobranca = .iNaoImpCobranca
        objNFiscal.sMensagemCorpoNota = .sMensagemCorpoNota
        objNFiscal.iNFe = .iNFe

    End With

End Sub

Function NFiscal_Le(ByVal objNFiscal As ClassNFiscal) As Long
'Lê a Nota Fiscal com o Número Interno passado em objNFiscal.

Dim lErro As Long
Dim lComando As Long
Dim tNFiscal As typeNFiscal
Dim dHoraEntrada As Double
Dim dHoraSaida As Double
Dim sMotivo As String
Dim sCodVerificacao As String, lNumNFe As Long

On Error GoTo Erro_NFiscal_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 31439

    With tNFiscal

        .sMensagemNota = String(STRING_NFISCAL_MENSAGEM, 0)
        .sMensagemCorpoNota = String(STRING_NFISCAL_MENSAGEM, 0)
        .sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
        .sPlaca = String(STRING_NFISCAL_PLACA, 0)
        .sPlacaUF = String(STRING_NFISCAL_PLACA_UF, 0)
        .sSerie = String(STRING_SERIE, 0)
        .sNumPedidoTerc = String(STRING_NUM_PEDIDO_TERC, 0)
        .sVolumeNumero = String(STRING_NFISCAL_VOLUME_NUMERO, 0)
        .sObservacao = String(STRING_NFISCAL_OBSERVACAO, 0)
        .sCodVerificacaoNFe = String(STRING_MAXIMO, 0)
        '.sChvNFe = String(STRING_NFE_CHNFE, 0)
        .sMotivoCancel = String(STRING_NFISCAL_MOTIVOCANCEL, 0)
        .sSerieNFPOrig = String(STRING_SERIE, 0)

        'Removido por Luiz Nogueira em 21/08/03
        '.lVolumeEspecie = String(STRING_NFISCAL_VOLUME_ESPECIE, 0)
        '.lVolumeMarca = String(STRING_NFISCAL_VOLUME_MARCA, 0)

        'Pesquisa a Nota Fiscal com o Número Interno passado
'        lErro = Comando_Executar(lComando, "SELECT Status, Serie, NumNotaFiscal,FilialEmpresa, Cliente, FilialCli, FilialEntrega, Fornecedor, FilialForn, DataEmissao, DataEntrada, DataSaida, DataVencimento,DataReferencia, NumPedidoVenda, NumPedidoTerc, ClasseDocCPR, NumIntDocCPR, ValorTotal, ValorProdutos, ValorFrete, ValorSeguro, ValorOutrasDespesas, ValorDesconto, CodTransportadora, MensagemNota, TabelaPreco, TipoNFiscal, NaturezaOp, PesoLiq, PesoBruto, NumIntTrib, Placa, PlacaUF, VolumeQuant, VolumeEspecie, VolumeMarca, Canal, NumIntNotaOriginal, VolumeNumero, FreteRespons, FilialPedido, Observacao, HoraEntrada, HoraSaida, CodTranspRedesp, DetPagFrete, Moeda, TaxaMoeda, VolumeTotal, SemDataSaida, NaoImpCobranca, MensagemCorpoNota, RPS, NumNFe, CodVerificacaoNFe, NFe, ValorDescontoTit, ValorDescontoItens, ValorItens, FornEntTerc, FilialFornEntTerc, ChvNFe FROM NFiscal WHERE NumIntDoc = ? AND Status <> ?", .iStatus, .sSerie, .lNumNotaFiscal, _
'           .iFilialEmpresa, .lCliente, .iFilialCli, .iFilialEntrega, .lFornecedor, .iFilialForn, .dtDataEmissao, .dtDataEntrada, .dtDataSaida, .dtDataVencimento, .dtDataReferencia, .lNumPedidoVenda, .sNumPedidoTerc, .iClasseDocCPR, .lNumIntDocCPR, .dValorTotal, .dValorProdutos, .dValorFrete, .dValorSeguro, .dValorOutrasDespesas, .dValorDesconto, .iCodTransportadora, .sMensagemNota, .iTabelaPreco, .iTipoNFiscal, .sNaturezaOp, .dPesoLiq, .dPesoBruto, .lNumIntTrib, .sPlaca, .sPlacaUF, .lVolumeQuant, .lVolumeEspecie, .lVolumeMarca, .iCanal, .lNumIntNotaOriginal, .sVolumeNumero, .iFreteRespons, .iFilialPedido, .sObservacao, dHoraEntrada, dHoraSaida, .iCodTranspRedesp, .iDetPagFrete, .iMoeda, .dTaxaMoeda, .dVolumeTotal, .iSemDataSaida, .iNaoImpCobranca, .sMensagemCorpoNota, .iRecibo, .lNumNFe, .sCodVerificacaoNFe, .iNFe, .dValorDescontoTit, .dValorDescontoItens, .dValorItens, .lFornEntTerc, .iFilialFornEntTerc, .sChvNFe, objNFiscal.lNumIntDoc, STATUS_EXCLUIDO)
        lErro = Comando_Executar(lComando, "SELECT Status, Serie, NumNotaFiscal,FilialEmpresa, Cliente, FilialCli, FilialEntrega, Fornecedor, FilialForn, DataEmissao, DataEntrada, DataSaida, DataVencimento,DataReferencia, NumPedidoVenda, NumPedidoTerc, ClasseDocCPR, NumIntDocCPR, ValorTotal, ValorProdutos, ValorFrete, ValorSeguro, ValorOutrasDespesas, ValorDesconto, CodTransportadora, MensagemNota, TabelaPreco, TipoNFiscal, NaturezaOp, PesoLiq, PesoBruto, NumIntTrib, Placa, PlacaUF, VolumeQuant, VolumeEspecie, VolumeMarca, Canal, NumIntNotaOriginal, VolumeNumero, FreteRespons, FilialPedido, Observacao, HoraEntrada, HoraSaida, CodTranspRedesp, DetPagFrete, Moeda, TaxaMoeda, VolumeTotal, SemDataSaida, NaoImpCobranca, MensagemCorpoNota, RPS, NumNFe, CodVerificacaoNFe, NFe, ValorDescontoTit, ValorDescontoItens, ValorItens, FornEntTerc, FilialFornEntTerc,CliIntermediario,FilialCliIntermediario, MotivoCancel, NumRecebimento, " & _
                                            "SerieNFPOrig, NumNFPOrig FROM NFiscal WHERE NumIntDoc = ? AND Status <> ?", .iStatus, .sSerie, .lNumNotaFiscal, _
           .iFilialEmpresa, .lCliente, .iFilialCli, .iFilialEntrega, .lFornecedor, .iFilialForn, .dtDataEmissao, .dtDataEntrada, .dtDataSaida, .dtDataVencimento, .dtDataReferencia, .lNumPedidoVenda, .sNumPedidoTerc, .iClasseDocCPR, .lNumIntDocCPR, .dValorTotal, .dValorProdutos, .dValorFrete, .dValorSeguro, .dValorOutrasDespesas, .dValorDesconto, .iCodTransportadora, .sMensagemNota, .iTabelaPreco, .iTipoNFiscal, .sNaturezaOp, .dPesoLiq, .dPesoBruto, .lNumIntTrib, .sPlaca, .sPlacaUF, .lVolumeQuant, .lVolumeEspecie, .lVolumeMarca, .iCanal, .lNumIntNotaOriginal, .sVolumeNumero, .iFreteRespons, .iFilialPedido, .sObservacao, dHoraEntrada, dHoraSaida, .iCodTranspRedesp, .iDetPagFrete, .iMoeda, .dTaxaMoeda, .dVolumeTotal, .iSemDataSaida, .iNaoImpCobranca, .sMensagemCorpoNota, .iRecibo, .lNumNFe, .sCodVerificacaoNFe, .iNFe, .dValorDescontoTit, .dValorDescontoItens, .dValorItens, .lFornEntTerc, .iFilialFornEntTerc, .lCliIntermediario, .iFilialCliIntermediario, .sMotivoCancel, .lNumRecebimento, _
            .sSerieNFPOrig, .lNumNFPOrig, objNFiscal.lNumIntDoc, STATUS_EXCLUIDO)
        If lErro <> AD_SQL_SUCESSO Then Error 31440

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 31441
        'Se não encontrar
        If lErro = AD_SQL_SEM_DADOS Then Error 31442

        'Preenche objNFiscal com os dados lidos
        objNFiscal.dPesoBruto = .dPesoBruto
        objNFiscal.dPesoLiq = .dPesoLiq
        objNFiscal.dtDataEmissao = .dtDataEmissao
        objNFiscal.dtDataEntrada = .dtDataEntrada
        objNFiscal.dtDataSaida = .dtDataSaida
        objNFiscal.dtDataVencimento = .dtDataVencimento
        objNFiscal.dValorDesconto = .dValorDesconto
        objNFiscal.dValorFrete = .dValorFrete
        objNFiscal.dValorOutrasDespesas = .dValorOutrasDespesas
        objNFiscal.dValorProdutos = .dValorProdutos
        objNFiscal.dValorSeguro = .dValorSeguro
        objNFiscal.dValorTotal = .dValorTotal
        objNFiscal.iCanal = .iCanal
        objNFiscal.iClasseDocCPR = .iClasseDocCPR
        objNFiscal.iCodTransportadora = .iCodTransportadora
        objNFiscal.iCodTranspRedesp = .iCodTranspRedesp
        objNFiscal.iDetPagFrete = .iDetPagFrete
        objNFiscal.iFilialCli = .iFilialCli
        objNFiscal.iFilialEmpresa = .iFilialEmpresa
        objNFiscal.iFilialEntrega = .iFilialEntrega
        objNFiscal.iFilialForn = .iFilialForn
        objNFiscal.iStatus = .iStatus
        objNFiscal.iTabelaPreco = .iTabelaPreco
        objNFiscal.iTipoNFiscal = .iTipoNFiscal
        objNFiscal.lVolumeQuant = .lVolumeQuant
        objNFiscal.lCliente = .lCliente
        objNFiscal.lFornecedor = .lFornecedor
        objNFiscal.lNumIntDocCPR = .lNumIntDocCPR
        objNFiscal.lNumIntNotaOriginal = .lNumIntNotaOriginal
        objNFiscal.lNumIntTrib = .lNumIntTrib
        objNFiscal.lNumNotaFiscal = .lNumNotaFiscal
        objNFiscal.sNumPedidoTerc = .sNumPedidoTerc
        objNFiscal.lNumPedidoVenda = .lNumPedidoVenda
        objNFiscal.sMensagemNota = .sMensagemNota
        objNFiscal.sNaturezaOp = .sNaturezaOp
        objNFiscal.sPlaca = .sPlaca
        objNFiscal.sPlacaUF = .sPlacaUF
        objNFiscal.sSerie = .sSerie
        objNFiscal.lVolumeEspecie = .lVolumeEspecie
        objNFiscal.lVolumeMarca = .lVolumeMarca
        objNFiscal.sVolumeNumero = .sVolumeNumero
        objNFiscal.iFreteRespons = .iFreteRespons
        objNFiscal.dtDataReferencia = .dtDataReferencia
        objNFiscal.iFilialPedido = .iFilialPedido
        objNFiscal.sObservacao = .sObservacao
        objNFiscal.dtHoraEntrada = CDate(dHoraEntrada)
        objNFiscal.dtHoraSaida = CDate(dHoraSaida)
        objNFiscal.iMoeda = .iMoeda
        objNFiscal.dTaxaMoeda = .dTaxaMoeda
        objNFiscal.dVolumeTotal = .dVolumeTotal
        objNFiscal.iSemDataSaida = .iSemDataSaida
        objNFiscal.iNaoImpCobranca = .iNaoImpCobranca
        objNFiscal.sMensagemCorpoNota = .sMensagemCorpoNota
        objNFiscal.iRecibo = .iRecibo
        objNFiscal.lNumNFe = .lNumNFe
        objNFiscal.sCodVerificacaoNFe = .sCodVerificacaoNFe
        objNFiscal.iNFe = .iNFe
        objNFiscal.dValorDescontoTit = .dValorDescontoTit
        objNFiscal.dValorItens = .dValorItens
        objNFiscal.dValorDescontoItens = .dValorDescontoItens
        objNFiscal.lFornEntTerc = .lFornEntTerc
        objNFiscal.iFilialFornEntTerc = .iFilialFornEntTerc
        'objNFiscal.sChvNFe = .sChvNFe
        
        objNFiscal.lCliIntermediario = .lCliIntermediario
        objNFiscal.iFilialCliIntermediario = .iFilialCliIntermediario
        
        objNFiscal.sMotivoCancel = .sMotivoCancel
        
        objNFiscal.lNumRecebimento = .lNumRecebimento
        
        objNFiscal.sSerieNFPOrig = .sSerieNFPOrig
        objNFiscal.lNumNFPOrig = .lNumNFPOrig
        
    End With

    sMotivo = String(STRING_NFE_XMOTIVO, 0)

    If objNFiscal.iRecibo = 0 Then
        lErro = CF("NFeFedProtNFE_Le", objNFiscal.lNumIntDoc, sMotivo)
        If lErro <> SUCESSO Then gError 203055
    Else
        lErro = RPSWEBProt_Le(objNFiscal.lNumIntDoc, sMotivo, sCodVerificacao, lNumNFe)
        If lErro <> SUCESSO Then gError 207069
        If sCodVerificacao <> "" Then
            objNFiscal.sCodVerificacaoNFe = sCodVerificacao
            objNFiscal.lNumNFe = lNumNFe
        End If
    End If
    
    objNFiscal.objInfoAdic.iTipoDoc = TIPODOC_INFOADIC_NF
    objNFiscal.objInfoAdic.lNumIntDoc = objNFiscal.lNumIntDoc
    lErro = InfoAdicionais_Le(objNFiscal.objInfoAdic)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 203055

    objNFiscal.sStatusNFeFed = sMotivo

    'Fecha o comando
    Call Comando_Fechar(lComando)

    NFiscal_Le = SUCESSO

    Exit Function

Erro_NFiscal_Le:

    NFiscal_Le = Err

    Select Case Err

        Case 31439
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 31440, 31441
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", Err)

        Case 31442, 203055, 207069

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149178)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Banco_Lock(ByVal lComando As Long, ByVal iCodBanco As Integer) As Long
'le o banco e faz lock shared

Dim lErro As Long, iBco As Integer

On Error GoTo Erro_Banco_Lock

    lErro = Comando_ExecutarLockado(lComando, "SELECT CodBanco FROM Bancos WHERE CodBanco = ? ", iBco, iCodBanco)
    If lErro <> AD_SQL_SUCESSO Then Error 19382

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 19383

    If lErro = AD_SQL_SEM_DADOS Then Error 19384
    
    lErro = Comando_LockShared(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 19385
    
    Banco_Lock = SUCESSO
    
    Exit Function
    
Erro_Banco_Lock:

    Banco_Lock = Err
    
    Select Case Err
        
        Case 19382, 19383
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BANCOS", Err, iCodBanco)
          
        Case 19384
            'Erro será tratado na rotina chamadora
            
        Case 19385
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_BANCOS", Err, iCodBanco)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149179)
    
    End Select
    
    Exit Function

End Function

Function MnemonicoFPreco_Le_Todos(ByVal objMnemonicoFPreco As ClassMnemonicoFPreco, ByVal colMnemonicoFPreco As Collection) As Long
'retorna todos os mnemonicos da tabela MnemonicoFPreco.

Dim lComando As Long
Dim sMnemonico As String
Dim iTipo As Integer
Dim sMnemonicoDesc As String
Dim lErro As Long
Dim objMnemonicoFPreco1 As ClassMnemonicoFPreco
Dim tMnemonicoFPreco As typeMnemonicoFPreco

On Error GoTo Erro_MnemonicoFPreco_Le_Todos

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 92299

    tMnemonicoFPreco.sMnemonico = String(STRING_MNEMONICOFPRECO_MNEMONICO, 0)
    tMnemonicoFPreco.sMnemonicoDesc = String(STRING_MNEMONICOFPRECO_MNEMONICODESC, 0)
    tMnemonicoFPreco.sExpressao = String(STRING_MNEMONICOFPRECO_EXPRESSAO, 0)

    'seleciona todos os mnemonicos
    lErro = Comando_Executar(lComando, "SELECT Mnemonico, Tipo, NumParam, Param1, Param2, Param3, MnemonicoDesc, Expressao FROM MnemonicoFPreco WHERE FilialEmpresa = ? AND Escopo = ? AND ItemCategoria = ? AND Produto = ? AND TabelaPreco = ? ORDER BY Mnemonico", tMnemonicoFPreco.sMnemonico, tMnemonicoFPreco.iTipo, tMnemonicoFPreco.iNumParam, tMnemonicoFPreco.iParam1, tMnemonicoFPreco.iParam2, tMnemonicoFPreco.iParam3, tMnemonicoFPreco.sMnemonicoDesc, tMnemonicoFPreco.sExpressao, objMnemonicoFPreco.iFilialEmpresa, objMnemonicoFPreco.iEscopo, objMnemonicoFPreco.sItemCategoria, objMnemonicoFPreco.sProduto, objMnemonicoFPreco.iTabelaPreco)
    If lErro <> AD_SQL_SUCESSO Then gError 92300

    'le o primeiro mnemonico
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92301

    Do While lErro = AD_SQL_SUCESSO

        Set objMnemonicoFPreco1 = New ClassMnemonicoFPreco


        objMnemonicoFPreco1.iFilialEmpresa = objMnemonicoFPreco.iFilialEmpresa
        objMnemonicoFPreco1.iEscopo = objMnemonicoFPreco.iEscopo
        objMnemonicoFPreco1.sItemCategoria = objMnemonicoFPreco.sItemCategoria
        objMnemonicoFPreco1.sProduto = objMnemonicoFPreco.sProduto
        objMnemonicoFPreco1.iTabelaPreco = objMnemonicoFPreco.iTabelaPreco
        objMnemonicoFPreco1.sMnemonico = tMnemonicoFPreco.sMnemonico
        objMnemonicoFPreco1.iTipo = tMnemonicoFPreco.iTipo
        objMnemonicoFPreco1.iNumParam = tMnemonicoFPreco.iNumParam
        objMnemonicoFPreco1.iParam1 = tMnemonicoFPreco.iParam1
        objMnemonicoFPreco1.iParam2 = tMnemonicoFPreco.iParam2
        objMnemonicoFPreco1.iParam3 = tMnemonicoFPreco.iParam3
        objMnemonicoFPreco1.sMnemonicoDesc = tMnemonicoFPreco.sMnemonicoDesc
        objMnemonicoFPreco1.sExpressao = tMnemonicoFPreco.sExpressao

        colMnemonicoFPreco.Add objMnemonicoFPreco1

        'le o proximo mnemonico
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92302

    Loop

    Call Comando_Fechar(lComando)

    MnemonicoFPreco_Le_Todos = SUCESSO

    Exit Function

Erro_MnemonicoFPreco_Le_Todos:

    MnemonicoFPreco_Le_Todos = gErr

    Select Case gErr

        Case 92299
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 92300, 92301, 92302
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MNEMONICOFPRECO", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149180)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function MnemonicoFPreco_Le_Todos1(ByVal objMnemonicoFPreco As ClassMnemonicoFPreco, ByVal colMnemonicoFPreco As Collection) As Long
'retorna todos os mnemonicos da tabela MnemonicoFPreco.

Dim lComando As Long
Dim sMnemonico As String
Dim iTipo As Integer
Dim sMnemonicoDesc As String
Dim lErro As Long
Dim objMnemonicoFPreco1 As ClassMnemonicoFPreco
Dim tMnemonicoFPreco As typeMnemonicoFPreco
Dim sSelecao As String
Dim sCategoria As String
Dim sProduto As String
Dim sMnemonicoAnterior As String
Dim sItem As String

On Error GoTo Erro_MnemonicoFPreco_Le_Todos1

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 92299

    tMnemonicoFPreco.sMnemonico = String(STRING_MNEMONICOFPRECO_MNEMONICO, 0)
    tMnemonicoFPreco.sMnemonicoDesc = String(STRING_MNEMONICOFPRECO_MNEMONICODESC, 0)
    tMnemonicoFPreco.sExpressao = String(STRING_MNEMONICOFPRECO_EXPRESSAO, 0)

    If objMnemonicoFPreco.iEscopo = MNEMONICOFPRECO_ESCOPO_GERAL Then

        'seleciona todos os mnemonicos
        lErro = Comando_Executar(lComando, "SELECT Mnemonico, Tipo, NumParam, Param1, Param2, Param3, MnemonicoDesc, Expressao, Funcao FROM MnemonicoFPreco WHERE (FilialEmpresa = 0 OR FilialEmpresa = ?) AND Escopo = ? ORDER BY Mnemonico, FilialEmpresa DESC", tMnemonicoFPreco.sMnemonico, tMnemonicoFPreco.iTipo, tMnemonicoFPreco.iNumParam, tMnemonicoFPreco.iParam1, tMnemonicoFPreco.iParam2, tMnemonicoFPreco.iParam3, tMnemonicoFPreco.sMnemonicoDesc, tMnemonicoFPreco.sExpressao, tMnemonicoFPreco.iFuncao, objMnemonicoFPreco.iFilialEmpresa, objMnemonicoFPreco.iEscopo)
        If lErro <> AD_SQL_SUCESSO Then gError 92407

    ElseIf objMnemonicoFPreco.iEscopo = MNEMONICOFPRECO_ESCOPO_CATEGORIA Then

        'seleciona todos os mnemonicos
        lErro = Comando_Executar(lComando, "SELECT Mnemonico, Tipo, NumParam, Param1, Param2, Param3, MnemonicoDesc, Expressao, Funcao FROM MnemonicoFPreco WHERE (FilialEmpresa = 0 OR FilialEmpresa = ?) AND (Escopo = ? OR (Escopo = ? AND ItemCategoria =?)) ORDER BY Mnemonico, Escopo DESC, FilialEmpresa DESC", tMnemonicoFPreco.sMnemonico, tMnemonicoFPreco.iTipo, tMnemonicoFPreco.iNumParam, tMnemonicoFPreco.iParam1, tMnemonicoFPreco.iParam2, tMnemonicoFPreco.iParam3, tMnemonicoFPreco.sMnemonicoDesc, tMnemonicoFPreco.sExpressao, tMnemonicoFPreco.iFuncao, objMnemonicoFPreco.iFilialEmpresa, MNEMONICOFPRECO_ESCOPO_GERAL, objMnemonicoFPreco.iEscopo, objMnemonicoFPreco.sItemCategoria)
        If lErro <> AD_SQL_SUCESSO Then gError 92408

    ElseIf objMnemonicoFPreco.iEscopo = MNEMONICOFPRECO_ESCOPO_PRODUTO Then

        sCategoria = CATEGORIAPRODUTO_PRECO
        sProduto = objMnemonicoFPreco.sProduto

        'retorna o item do produto/categoria
        lErro = CF("ProdutoCategoria_Le", sProduto, sCategoria, sItem)
        If lErro <> SUCESSO And lErro <> 92405 Then gError 92406

        'seleciona todos os mnemonicos
        lErro = Comando_Executar(lComando, "SELECT Mnemonico, Tipo, NumParam, Param1, Param2, Param3, MnemonicoDesc, Expressao, Funcao FROM MnemonicoFPreco WHERE (FilialEmpresa = 0 OR FilialEmpresa = ?) AND (Escopo = ? OR (Escopo = ? AND ItemCategoria =?) OR (Escopo =? AND Produto =?)) ORDER BY Mnemonico, Escopo DESC, FilialEmpresa DESC ", _
            tMnemonicoFPreco.sMnemonico, tMnemonicoFPreco.iTipo, tMnemonicoFPreco.iNumParam, tMnemonicoFPreco.iParam1, tMnemonicoFPreco.iParam2, tMnemonicoFPreco.iParam3, tMnemonicoFPreco.sMnemonicoDesc, tMnemonicoFPreco.sExpressao, tMnemonicoFPreco.iFuncao, objMnemonicoFPreco.iFilialEmpresa, MNEMONICOFPRECO_ESCOPO_GERAL, MNEMONICOFPRECO_ESCOPO_CATEGORIA, sItem, objMnemonicoFPreco.iEscopo, objMnemonicoFPreco.sProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 92409

    ElseIf objMnemonicoFPreco.iEscopo = MNEMONICOFPRECO_ESCOPO_TABPRECO Then

        sCategoria = CATEGORIAPRODUTO_PRECO
        sProduto = objMnemonicoFPreco.sProduto

        'retorna o item do produto/categoria
        lErro = CF("ProdutoCategoria_Le", sProduto, sCategoria, sItem)
        If lErro <> SUCESSO And lErro <> 92405 Then gError 92410

        'seleciona todos os mnemonicos
        lErro = Comando_Executar(lComando, "SELECT Mnemonico, Tipo, NumParam, Param1, Param2, Param3, MnemonicoDesc, Expressao, Funcao FROM MnemonicoFPreco WHERE (FilialEmpresa = 0 OR FilialEmpresa = ?) AND (Escopo = ? OR (Escopo = ? AND ItemCategoria =?) OR (Escopo =? AND Produto =?) OR (Escopo =? AND Produto =? AND TabelaPreco = ?)) ORDER BY Mnemonico, Escopo DESC, FilialEmpresa DESC", _
            tMnemonicoFPreco.sMnemonico, tMnemonicoFPreco.iTipo, tMnemonicoFPreco.iNumParam, tMnemonicoFPreco.iParam1, tMnemonicoFPreco.iParam2, tMnemonicoFPreco.iParam3, tMnemonicoFPreco.sMnemonicoDesc, tMnemonicoFPreco.sExpressao, tMnemonicoFPreco.iFuncao, objMnemonicoFPreco.iFilialEmpresa, MNEMONICOFPRECO_ESCOPO_GERAL, MNEMONICOFPRECO_ESCOPO_CATEGORIA, sItem, MNEMONICOFPRECO_ESCOPO_PRODUTO, objMnemonicoFPreco.sProduto, objMnemonicoFPreco.iEscopo, objMnemonicoFPreco.sProduto, objMnemonicoFPreco.iTabelaPreco)
        If lErro <> AD_SQL_SUCESSO Then gError 92411

    End If

    'le o primeiro mnemonico
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92301

    Do While lErro = AD_SQL_SUCESSO

        If sMnemonicoAnterior <> tMnemonicoFPreco.sMnemonico Then

            sMnemonicoAnterior = tMnemonicoFPreco.sMnemonico

            Set objMnemonicoFPreco1 = New ClassMnemonicoFPreco

            objMnemonicoFPreco1.iFilialEmpresa = objMnemonicoFPreco.iFilialEmpresa
            objMnemonicoFPreco1.iEscopo = objMnemonicoFPreco.iEscopo
            objMnemonicoFPreco1.sItemCategoria = objMnemonicoFPreco.sItemCategoria
            objMnemonicoFPreco1.sProduto = objMnemonicoFPreco.sProduto
            objMnemonicoFPreco1.iTabelaPreco = objMnemonicoFPreco.iTabelaPreco
            objMnemonicoFPreco1.sMnemonico = tMnemonicoFPreco.sMnemonico
            objMnemonicoFPreco1.iTipo = tMnemonicoFPreco.iTipo
            objMnemonicoFPreco1.iNumParam = tMnemonicoFPreco.iNumParam
            objMnemonicoFPreco1.iParam1 = tMnemonicoFPreco.iParam1
            objMnemonicoFPreco1.iParam2 = tMnemonicoFPreco.iParam2
            objMnemonicoFPreco1.iParam3 = tMnemonicoFPreco.iParam3
            objMnemonicoFPreco1.sMnemonicoDesc = tMnemonicoFPreco.sMnemonicoDesc
            objMnemonicoFPreco1.sExpressao = tMnemonicoFPreco.sExpressao
            objMnemonicoFPreco1.iFuncao = tMnemonicoFPreco.iFuncao

            colMnemonicoFPreco.Add objMnemonicoFPreco1

        End If

        'le o proximo mnemonico
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92302

    Loop

    Call Comando_Fechar(lComando)

    MnemonicoFPreco_Le_Todos1 = SUCESSO

    Exit Function

Erro_MnemonicoFPreco_Le_Todos1:

    MnemonicoFPreco_Le_Todos1 = gErr

    Select Case gErr

        Case 92299
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 92406, 92410

        Case 92300, 92301, 92302, 92406, 92407, 92408, 92409, 92411
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MNEMONICOFPRECO", gErr)

        Case 92412, 92413
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTOCATEGORIA_NAO_CADASTRADO", gErr, sCategoria, sProduto)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149181)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function MnemonicoFPreco_Le(ByVal objMnemonicoFPreco As ClassMnemonicoFPreco) As Long
'le e retorna os dados do mnemonico passado como parametro

Dim lComando As Long
Dim lErro As Long
Dim tMnemonicoFPreco As typeMnemonicoFPreco

On Error GoTo Erro_MnemonicoFPreco_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 92233

    tMnemonicoFPreco.sMnemonicoDesc = String(STRING_MNEMONICOFPRECO_MNEMONICODESC, 0)
    tMnemonicoFPreco.sExpressao = String(STRING_MNEMONICOFPRECO_EXPRESSAO, 0)

    'seleciona o mnemonico
    lErro = Comando_Executar(lComando, "SELECT Tipo, NumParam, Param1, Param2, Param3, MnemonicoDesc, Expressao FROM MnemonicoFPreco WHERE FilialEmpresa = ? AND Escopo = ? AND ItemCategoria = ? AND Produto = ? AND TabelaPreco = ? AND Mnemonico = ?", tMnemonicoFPreco.iTipo, tMnemonicoFPreco.iNumParam, tMnemonicoFPreco.iParam1, tMnemonicoFPreco.iParam2, tMnemonicoFPreco.iParam3, tMnemonicoFPreco.sMnemonicoDesc, tMnemonicoFPreco.sExpressao, objMnemonicoFPreco.iFilialEmpresa, objMnemonicoFPreco.iEscopo, objMnemonicoFPreco.sItemCategoria, objMnemonicoFPreco.sProduto, objMnemonicoFPreco.iTabelaPreco, objMnemonicoFPreco.sMnemonico)
    If lErro <> AD_SQL_SUCESSO Then gError 92234

    'le o primeiro mnemonico
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92235

    'se o registro não foi encontrado
    If lErro = AD_SQL_SEM_DADOS Then gError 92236

    objMnemonicoFPreco.iTipo = tMnemonicoFPreco.iTipo
    objMnemonicoFPreco.iNumParam = tMnemonicoFPreco.iNumParam
    objMnemonicoFPreco.iParam1 = tMnemonicoFPreco.iParam1
    objMnemonicoFPreco.iParam2 = tMnemonicoFPreco.iParam2
    objMnemonicoFPreco.iParam3 = tMnemonicoFPreco.iParam3
    objMnemonicoFPreco.sMnemonicoDesc = tMnemonicoFPreco.sMnemonicoDesc
    objMnemonicoFPreco.sExpressao = tMnemonicoFPreco.sExpressao

    Call Comando_Fechar(lComando)

    MnemonicoFPreco_Le = SUCESSO

    Exit Function

Erro_MnemonicoFPreco_Le:

    MnemonicoFPreco_Le = gErr

    Select Case gErr

        Case 92233
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 92234, 92235
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MNEMONICOFPRECO1", gErr, objMnemonicoFPreco.sMnemonico)

        Case 92236

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149182)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function FormacaoPreco_Le(ByVal objFormacaoPreco As ClassFormacaoPreco, ByVal colFormacaoPreco As Collection) As Long
'le a formula de formação de preço em questão

Dim lErro As Long
Dim alComando(1 To 1) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim objFormacaoPreco1 As ClassFormacaoPreco
Dim sTitulo As String
Dim sExpressao As String
Dim iLinha As Integer

On Error GoTo Erro_FormacaoPreco_Le

    'abrir comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 92220
    Next

    sExpressao = String(STRING_FORMACAOPRECO_EXPRESSAO, 0)
    sTitulo = String(STRING_FORMACAOPRECO_TITULO, 0)

    'seleciona a formula segundo os parametros passados
    lErro = Comando_Executar(alComando(1), "SELECT NumIntDoc, Expressao, Linha, Titulo FROM FormacaoPreco WHERE FilialEmpresa=? AND Escopo=? AND ItemCategoria=? AND Produto=? AND TabelaPreco=? ORDER BY Linha", _
    lNumIntDoc, sExpressao, iLinha, sTitulo, objFormacaoPreco.iFilialEmpresa, objFormacaoPreco.iEscopo, objFormacaoPreco.sItemCategoria, objFormacaoPreco.sProduto, objFormacaoPreco.iTabelaPreco)
    If lErro <> AD_SQL_SUCESSO Then gError 92221

    'le a formula em questão
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92222

    'se a expressão não estiver cadastrada ==> erro
    If lErro = AD_SQL_SEM_DADOS Then gError 92223

    Do While lErro = AD_SQL_SUCESSO

        Set objFormacaoPreco1 = New ClassFormacaoPreco

        objFormacaoPreco1.iFilialEmpresa = objFormacaoPreco.iFilialEmpresa
        objFormacaoPreco1.iEscopo = objFormacaoPreco.iEscopo
        objFormacaoPreco1.sItemCategoria = objFormacaoPreco.sItemCategoria
        objFormacaoPreco1.sProduto = objFormacaoPreco.sProduto
        objFormacaoPreco1.iTabelaPreco = objFormacaoPreco.iTabelaPreco
        objFormacaoPreco1.lNumIntDoc = lNumIntDoc
        objFormacaoPreco1.sExpressao = sExpressao
        objFormacaoPreco1.iLinha = iLinha
        objFormacaoPreco1.sTitulo = sTitulo

        colFormacaoPreco.Add objFormacaoPreco1

        'le a formula em questão
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92268

    Loop

   'fechar comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    FormacaoPreco_Le = SUCESSO

    Exit Function

Erro_FormacaoPreco_Le:

    FormacaoPreco_Le = gErr

    Select Case gErr

        Case 92220
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 92221, 92222, 92268
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FORMACAOPRECO", gErr, objFormacaoPreco.iFilialEmpresa, objFormacaoPreco.iEscopo, objFormacaoPreco.sItemCategoria, objFormacaoPreco.sProduto, objFormacaoPreco.iTabelaPreco)

        Case 92223

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149183)

    End Select

   'fechar comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function FormacaoPreco_Le1(ByVal objFormacaoPreco As ClassFormacaoPreco, ByVal colFormacaoPreco As Collection) As Long
'le a formula de formação de preço em questão

Dim lErro As Long
Dim sProduto As String
Dim sCategoria As String
Dim sItem As String

On Error GoTo Erro_FormacaoPreco_Le1

    lErro = CF("FormacaoPreco_Le", objFormacaoPreco, colFormacaoPreco)
    If lErro <> SUCESSO And lErro <> 92223 Then gError 92430

    If lErro = 92223 And objFormacaoPreco.iEscopo = FORMACAO_PRECO_ESCOPO_GERAL Then gError 92434

    If lErro = 92223 Then

        If objFormacaoPreco.iEscopo = FORMACAO_PRECO_ESCOPO_TABPRECO Then

            objFormacaoPreco.iEscopo = FORMACAO_PRECO_ESCOPO_PRODUTO
            objFormacaoPreco.iTabelaPreco = 0

        ElseIf objFormacaoPreco.iEscopo = FORMACAO_PRECO_ESCOPO_PRODUTO Then

            sCategoria = CATEGORIAPRODUTO_PRECO
            sProduto = objFormacaoPreco.sProduto

            'retorna o item do produto/categoria
            lErro = CF("ProdutoCategoria_Le", sProduto, sCategoria, sItem)
            If lErro <> SUCESSO And lErro <> 92405 Then gError 92431

            objFormacaoPreco.iEscopo = FORMACAO_PRECO_ESCOPO_CATEGORIA
            objFormacaoPreco.sItemCategoria = sItem
            objFormacaoPreco.sProduto = ""

        ElseIf objFormacaoPreco.iEscopo = FORMACAO_PRECO_ESCOPO_CATEGORIA Then

            objFormacaoPreco.iEscopo = FORMACAO_PRECO_ESCOPO_GERAL
            objFormacaoPreco.sItemCategoria = ""

        End If

        lErro = CF("FormacaoPreco_Le1", objFormacaoPreco, colFormacaoPreco)
        If lErro <> SUCESSO Then gError 92432

    End If

    FormacaoPreco_Le1 = SUCESSO

    Exit Function

Erro_FormacaoPreco_Le1:

    FormacaoPreco_Le1 = gErr

    Select Case gErr

        Case 92430, 92431, 92432, 92434

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149184)

    End Select

    Exit Function

End Function

'Rotina Criada por Daniel em 18/02/02
Function CondPagto_Critica_Usu(objCondicaoPagto As ClassCondicaoPagto) As Long
'Como é Versao2 => apenas retorna SUCESSO
'Em versões USU ela Critica a Condicao de Pagamento a
'partir do Cliente (Ex. GrandeRio, INPAL...)

    CondPagto_Critica_Usu = SUCESSO

End Function

'Rotina Criada por Daniel em 18/02/02
Function Cliente_Critica_Usu(objCliente As ClassCliente) As Long
'Como é Versao2 => apenas retorna SUCESSO
'Em versões USU ela Critica o uso do Cliente em SolicitacaoServico e TabPreco

    Cliente_Critica_Usu = SUCESSO

End Function

'Revisado - OK Luiz 19/02/02
Public Function ClienteHistorico_Le_Todos(ByVal colClienteHistorico As Collection) As Long
'Lê todos os registros da tabela ClienteHistorico e devolve-os na coleção
'colClienteHistorico RETORNA(Output) os registros lidos

Dim lErro As Long
Dim lComando As Long
Dim sSelect As String
Dim tClienteHistorico As typeClienteHistorico
Dim objClienteHistorico As ClassClienteHistorico

On Error GoTo Erro_ClienteHistorico_Le_Todos

    'Abre o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 97806
    
    'Carrega sSelect com a Expressão SQL
    sSelect = "SELECT ClienteHistorico.DataAtualizacao, ClienteHistorico.Codigo, ClienteHistorico.Filial , ClienteHistorico.CPF_CGC, ClienteHistorico.InscricaoEstadual,ClienteHistorico.InscricaoSuframa, ClienteHistorico.InscricaoMunicipal, ClienteHistorico.RazaoSocial, ClienteHistorico.Endereco, ClienteHistorico.Bairro, ClienteHistorico.Cidade, ClienteHistorico.UF, ClienteHistorico.Pais, ClienteHistorico.CEP FROM ClienteHistorico"
    
    With tClienteHistorico
        
        'Inicializa as Strings do Type
        .sBairro = String(STRING_BAIRRO, 0)
        .sCEP = String(STRING_CEP, 0)
        .sCgc = String(STRING_CGC, 0)
        .sCidade = String(STRING_CIDADE, 0)
        .sEndereco = String(STRING_ENDERECO, 0)
        .sInscricaoEstadual = String(STRING_INSCR_EST, 0)
        .sInscricaoMunicipal = String(STRING_INSCR_MUN, 0)
        .sInscricaoSuframa = String(STRING_INSCR_SUF, 0)
        .sPais = String(STRING_PAISES_NOME, 0)
        .sRazaoSocial = String(STRING_CLIENTE_RAZAO_SOCIAL, 0)
        .sUF = String(STRING_ESTADO_SIGLA, 0)
    
        'Executa o Comando SQL
        lErro = Comando_Executar(lComando, sSelect, .dtDataAtualizacao, .lCodigo, .iFilialCliente, .sCgc, .sInscricaoEstadual, .sInscricaoSuframa, .sInscricaoMunicipal, .sRazaoSocial, .sEndereco, .sBairro, .sCidade, .sUF, .sPais, .sCEP)
        If lErro <> AD_SQL_SUCESSO Then gError 97807
    
    End With
        
    'Busca o Primeiro Registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97808

    'Se nao encontrou => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 97805
    
    'Enquanto tem registro => adiciona na coleção
    Do While lErro <> AD_SQL_SEM_DADOS
                                    
        'Inicializa o obj
        Set objClienteHistorico = New ClassClienteHistorico
        
        With objClienteHistorico
    
            .dtDataAtualizacao = tClienteHistorico.dtDataAtualizacao
            .iFilialCliente = tClienteHistorico.iFilialCliente
            .lCodigo = tClienteHistorico.lCodigo
            .sBairro = tClienteHistorico.sBairro
            .sCEP = tClienteHistorico.sCEP
            .sCgc = tClienteHistorico.sCgc
            .sCidade = tClienteHistorico.sCidade
            .sEndereco = tClienteHistorico.sEndereco
            .sInscricaoEstadual = tClienteHistorico.sInscricaoEstadual
            .sInscricaoMunicipal = tClienteHistorico.sInscricaoMunicipal
            .sInscricaoSuframa = tClienteHistorico.sInscricaoSuframa
            .sUF = tClienteHistorico.sUF
            .sRazaoSocial = tClienteHistorico.sRazaoSocial
            .sPais = tClienteHistorico.sPais
            
        End With
    
        'Adiciona na Colecao
        colClienteHistorico.Add objClienteHistorico
    
        'Busca o Próximo Registro
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97809
        
    Loop
    
    'Fecha o Comando
    Call Comando_Fechar(lComando)
        
    ClienteHistorico_Le_Todos = SUCESSO
    
    Exit Function
    
Erro_ClienteHistorico_Le_Todos:

    ClienteHistorico_Le_Todos = gErr

    Select Case gErr
        
        Case 97806
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 97807, 97808, 97809
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLIENTEHISTORICO", gErr)
            
        Case 97805 'Sem Dados
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149185)
    
    End Select

    'Fecha o Comando
    Call Comando_Fechar(lComando)
    
End Function

'Revisado - OK Luiz 19/02/02
Public Function TranspHistorico_Le_Todos(ByVal colTranspHistorico As Collection) As Long
'Lê todos os registros da tabela TransportadoraHistorico e devolve-os na coleção
'colTranspHistorico RETORNA(Output) os registros lidos

Dim lErro As Long
Dim lComando As Long
Dim sSelect As String
Dim tTranspHistorico As typeTranspHistorico
Dim objTranspHistorico As ClassTranspHistorico

On Error GoTo Erro_TranspHistorico_Le_Todos

    'Abre o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 97795
    
    'Carrega sSelect com a Expressão SQL
    sSelect = "SELECT TransportadoraHistorico.DataAtualizacao, TransportadoraHistorico.CodTransp, TransportadoraHistorico.CGC, TransportadoraHistorico.InscricaoEstadual, TransportadoraHistorico.InscricaoMunicipal, TransportadoraHistorico.RazaoSocial, TransportadoraHistorico.Endereco, TransportadoraHistorico.Bairro, TransportadoraHistorico.Cidade, TransportadoraHistorico.UF, TransportadoraHistorico.Pais, TransportadoraHistorico.CEP FROM TransportadoraHistorico"
    
    With tTranspHistorico
        
        'Inicializa as Strings do Type
        .sBairro = String(STRING_BAIRRO, 0)
        .sCEP = String(STRING_CEP, 0)
        .sCgc = String(STRING_CGC, 0)
        .sCidade = String(STRING_CIDADE, 0)
        .sEndereco = String(STRING_ENDERECO, 0)
        .sInscricaoEstadual = String(STRING_INSCR_EST, 0)
        .sInscricaoMunicipal = String(STRING_INSCR_MUN, 0)
        .sPais = String(STRING_PAISES_NOME, 0)
        .sRazaoSocial = String(STRING_TRANSPORTADORA_NOME, 0)
        .sUF = String(STRING_ESTADO_SIGLA, 0)
    
        'Executa o Comando SQL
        lErro = Comando_Executar(lComando, sSelect, .dtDataAtualizacao, .iCodTransp, .sCgc, .sInscricaoEstadual, .sInscricaoMunicipal, .sRazaoSocial, .sEndereco, .sBairro, .sCidade, .sUF, .sPais, .sCEP)
        If lErro <> AD_SQL_SUCESSO Then gError 97796
    
    End With
        
    'Busca o Primeiro Registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97797

    If lErro = AD_SQL_SEM_DADOS Then gError 97798
    
    'Enquanto tem Transportadora => adiciona na colecao
    Do While lErro <> AD_SQL_SEM_DADOS
        
        'Inicializa o obj
        Set objTranspHistorico = New ClassTranspHistorico
        
        With objTranspHistorico
    
            .dtDataAtualizacao = tTranspHistorico.dtDataAtualizacao
            .iCodTransp = tTranspHistorico.iCodTransp
            .sBairro = tTranspHistorico.sBairro
            .sCEP = tTranspHistorico.sCEP
            .sCgc = tTranspHistorico.sCgc
            .sCidade = tTranspHistorico.sCidade
            .sEndereco = tTranspHistorico.sEndereco
            .sInscricaoEstadual = tTranspHistorico.sInscricaoEstadual
            .sInscricaoMunicipal = tTranspHistorico.sInscricaoMunicipal
            .sUF = tTranspHistorico.sUF
            .sRazaoSocial = tTranspHistorico.sRazaoSocial
            .sPais = tTranspHistorico.sPais
            
        End With
    
        'Adiciona na colecao
        colTranspHistorico.Add objTranspHistorico
    
        'Busca o Próximo Registro
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97799
            
    Loop
    
    'Fecha o Comando
    Call Comando_Fechar(lComando)
    
    TranspHistorico_Le_Todos = SUCESSO
    
    Exit Function
    
Erro_TranspHistorico_Le_Todos:

    TranspHistorico_Le_Todos = gErr

    Select Case gErr
        
        Case 97795
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 97796, 97797, 97799
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRANSPHISTORICO", gErr)
            
        Case 97798 'Sem dados
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149186)
    
    End Select

    'Fecha o Comando
    Call Comando_Fechar(lComando)

End Function

'Revisado - OK Luiz 19/02/02
Public Function NaturezaOPHistorico_Le_Todos(ByVal colNaturezaOPHistorico As Collection) As Long
'Lê todos os registros da tabela NaturezaOPHistorico e devolve-os na coleção
'colNaturezaOPHistorico RETORNA(Output) os registros lidos

Dim lErro As Long
Dim lComando As Long
Dim sSelect As String
Dim tNaturezaOPHistorico As typeNaturezaOPHistorico
Dim objNaturezaOPHistorico As ClassNaturezaOPHistorico

On Error GoTo Erro_NaturezaOPHistorico_Le_Todos

    'Abre o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 97801
        
    'Carrega sSelect com a Expressão SQL
    sSelect = "SELECT NaturezaOPHistorico.DataAtualizacao, NaturezaOPHistorico.CodNaturezaOP, NaturezaOPHistorico.DescNaturezaOP FROM NaturezaOPHistorico"
    
    With tNaturezaOPHistorico
        
        'Inicializa as Strings do Type
        .sCodigo = String(STRING_NATUREZAOP_CODIGO, 0)
        .sDescricao = String(STRING_NATUREZAOP_DESCRNF, 0)
        
        'Executa o Comando SQL
        lErro = Comando_Executar(lComando, sSelect, .dtDataAtualizacao, .sCodigo, .sDescricao)
        If lErro <> AD_SQL_SUCESSO Then gError 97802
    
    End With
        
    'Busca o Primeiro Registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97803

    'Se nao encontrou => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 97800
    
    'Enquanto tem NaturezaOP => adiciona na colecao
    Do While lErro <> AD_SQL_SEM_DADOS
                                    
        'Inicializa o Obj
        Set objNaturezaOPHistorico = New ClassNaturezaOPHistorico
        
        With objNaturezaOPHistorico
    
            .dtDataAtualizacao = tNaturezaOPHistorico.dtDataAtualizacao
            .sCodigo = tNaturezaOPHistorico.sCodigo
            .sDescricao = tNaturezaOPHistorico.sDescricao
            
        End With
    
        'Adiciona na Colecao
        colNaturezaOPHistorico.Add objNaturezaOPHistorico
    
        'Busca o Próximo Registro
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97804
            
    Loop
    
    'Fecha o Comando
    Call Comando_Fechar(lComando)
        
    NaturezaOPHistorico_Le_Todos = SUCESSO
    
    Exit Function
    
Erro_NaturezaOPHistorico_Le_Todos:

    NaturezaOPHistorico_Le_Todos = gErr

    Select Case gErr
        
        Case 97801
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 97802, 97803, 97804
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NATUREZAOPHISTORICO", gErr)
            
        Case 97800 'Sem dados
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149187)
    
    End Select

    'Fecha o Comando
    Call Comando_Fechar(lComando)
        
End Function

'Incluida por Tulio em 20/03 na hora de subir
'as funcoes da tela PlanComissoesInpal

Public Function Cliente_Lock(ByVal lCliente As Long) As Long
'Faz o lock shared do Cliente a partir do código passado em lCliente
'lCliente RECEBE (INPUT) o codigo do cliente que sera utilizado
'para a realizacao do lock

Dim lErro As Long
Dim lComando As Long
Dim lCodCli As Long

On Error GoTo Erro_Cliente_lock

    'abre comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 98891
    
    If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL Then
        
        'obtem o conjunto unitario de clientes que se deseja lockar
        lErro = Comando_ExecutarLockado(lComando, "SELECT CodigoLoja FROM Clientes WHERE CodigoLoja = ?", lCodCli, lCliente)
    
    Else
        
        'obtem o conjunto unitario de clientes que se deseja lockar
        lErro = Comando_ExecutarLockado(lComando, "SELECT Codigo FROM Clientes WHERE Codigo = ?", lCodCli, lCliente)
    
    End If
    
    If lErro <> AD_SQL_SUCESSO Then gError 98892

    'se posiciona no unico elemento do conjunto obtido anteriormente
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98893

    'se o conjunto obtido nao tinha elementos...
    If lErro = AD_SQL_SEM_DADOS Then Error 98895
    
    '"LockShared" do cliente
    lErro = Comando_LockShared(lComando)
    If lErro <> SUCESSO Then Error 98894
    
    'fecha comando
    Call Comando_Fechar(lComando)
    
    Cliente_Lock = SUCESSO
    
    Exit Function

Erro_Cliente_lock:

    Cliente_Lock = gErr
    
    Select Case gErr
        
        Case 98891
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 98892, 98893
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLIENTES1", gErr)
        
        Case 98894
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_CLIENTES", gErr)
                
        Case 98895
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, lCliente)
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 149188)
            
    End Select
    
    'fecha comando
    Call Comando_Fechar(lComando)
    
    Exit Function
        
End Function

'*************
'Sergio Tela de VendedorFilial dia 3/07/02(Inicio)

Function VendedorLoja_Le_FilalEmpresa(objVendedor As ClassVendedor) As Long
'Lê todos os Vendedores Ativos da Empresa

Dim lErro As Long
Dim lComando As Long
Dim tVendedorFilial As typeVendedorFilial
Dim objFilialEmpresa As AdmFiliais

On Error GoTo Erro_VendedorLoja_Le_FilalEmpresa

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 107505
    
    'Seleciona todas as Filiais que o Vendedor esta Cadastrado
    lErro = Comando_Executar(lComando, "SELECT FilialEmpresa FROM VendedorLoja WHERE Vendedor = ?", tVendedorFilial.iFilialEmpresa, objVendedor.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 107506

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 107507

    If lErro = AD_SQL_SEM_DADOS Then gError 107503
    
    Do While lErro = AD_SQL_SUCESSO
    
        'Inicializar o objFilialEmpresa
        Set objFilialEmpresa = New AdmFiliais
        
        'Guarda na coleção os códigos das Filiais Empresa que o Vendedor está cadastrada
        objFilialEmpresa.iCodFilial = tVendedorFilial.iFilialEmpresa
        
        'Adcionar a Filial do Vendedor na Coleção de Filiais dentro do objVendedor
        objVendedor.colFiliaisLoja.Add objFilialEmpresa

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 107509

    Loop

    Call Comando_Fechar(lComando)

    VendedorLoja_Le_FilalEmpresa = SUCESSO

    Exit Function

Erro_VendedorLoja_Le_FilalEmpresa:

    VendedorLoja_Le_FilalEmpresa = gErr

    Select Case gErr
        
        Case 107503
            'Não emite msg de Erro
        
        Case 107505
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 107506, 107507, 107509
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_VENDEDORLOJA", gErr, objVendedor.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149189)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

'criado por Mauricio Maciel em 31/03/2003
Function Planilhas_Le(ByVal objPlanilhas As ClassPlanilhas, ByVal colPlanilhas As Collection) As Long
'le a formula de formação de preço em questão

Dim lErro As Long
Dim alComando(1 To 1) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim objPlanilhas1 As ClassPlanilhas
Dim sTitulo As String
Dim sExpressao As String
Dim iLinha As Integer

On Error GoTo Erro_Planilhas_Le

    'abrir comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 92220
    Next

    sExpressao = String(STRING_FORMACAOPRECO_EXPRESSAO, 0)
    sTitulo = String(STRING_FORMACAOPRECO_TITULO, 0)

    'seleciona a formula segundo os parametros passados
    lErro = Comando_Executar(alComando(1), "SELECT NumIntDoc, Expressao, Linha, Titulo, TipoPlanilha FROM Planilhas WHERE FilialEmpresa=? AND Escopo=? AND ItemCategoria=? AND Produto=? AND TabelaPreco=? AND TipoPlanilha = ? ORDER BY Linha", _
    lNumIntDoc, sExpressao, iLinha, sTitulo, objPlanilhas.iTipoPlanilha, objPlanilhas.iFilialEmpresa, objPlanilhas.iEscopo, objPlanilhas.sItemCategoria, objPlanilhas.sProduto, objPlanilhas.iTabelaPreco, objPlanilhas.iTipoPlanilha)
    If lErro <> AD_SQL_SUCESSO Then gError 92221

    'le a formula em questão
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92222

    'se a expressão não estiver cadastrada ==> erro
    If lErro = AD_SQL_SEM_DADOS Then gError 92223

    Do While lErro = AD_SQL_SUCESSO

        Set objPlanilhas1 = New ClassPlanilhas

        objPlanilhas1.iFilialEmpresa = objPlanilhas.iFilialEmpresa
        objPlanilhas1.iEscopo = objPlanilhas.iEscopo
        objPlanilhas1.sItemCategoria = objPlanilhas.sItemCategoria
        objPlanilhas1.sProduto = objPlanilhas.sProduto
        objPlanilhas1.iTabelaPreco = objPlanilhas.iTabelaPreco
        objPlanilhas1.iTipoPlanilha = objPlanilhas.iTipoPlanilha
        objPlanilhas1.lNumIntDoc = lNumIntDoc
        objPlanilhas1.sExpressao = sExpressao
        objPlanilhas1.iLinha = iLinha
        objPlanilhas1.sTitulo = sTitulo

        colPlanilhas.Add objPlanilhas1

        'le a formula em questão
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92268

    Loop

   'fechar comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Planilhas_Le = SUCESSO

    Exit Function

Erro_Planilhas_Le:

    Planilhas_Le = gErr

    Select Case gErr

        Case 92220
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 92221, 92222, 92268
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANILHAS", gErr, objPlanilhas.iFilialEmpresa, objPlanilhas.iEscopo, objPlanilhas.sItemCategoria, objPlanilhas.sProduto, objPlanilhas.iTabelaPreco)

        Case 92223

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149190)

    End Select

   'fechar comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'criado por Mauricio Maciel em 31/03/2003
Function Planilhas_Le1(ByVal objPlanilhas As ClassPlanilhas, ByVal colPlanilhas As Collection) As Long
'le a formula de formação de preço em questão

Dim lErro As Long
Dim sProduto As String
Dim sCategoria As String
Dim sItem As String

On Error GoTo Erro_Planilhas_Le1

    lErro = CF("Planilhas_Le", objPlanilhas, colPlanilhas)
    If lErro <> SUCESSO And lErro <> 92223 Then gError 92430

    If lErro = 92223 And objPlanilhas.iEscopo = FORMACAO_PRECO_ESCOPO_GERAL Then gError 92434

    If lErro = 92223 Then

        If objPlanilhas.iEscopo = FORMACAO_PRECO_ESCOPO_TABPRECO Then

            objPlanilhas.iEscopo = FORMACAO_PRECO_ESCOPO_PRODUTO
            objPlanilhas.iTabelaPreco = 0

        ElseIf objPlanilhas.iEscopo = FORMACAO_PRECO_ESCOPO_PRODUTO Then

            sCategoria = CATEGORIAPRODUTO_PRECO
            sProduto = objPlanilhas.sProduto

            'retorna o item do produto/categoria
            lErro = CF("ProdutoCategoria_Le", sProduto, sCategoria, sItem)
            If lErro <> SUCESSO And lErro <> 92405 Then gError 92431

            objPlanilhas.iEscopo = FORMACAO_PRECO_ESCOPO_CATEGORIA
            objPlanilhas.sItemCategoria = sItem
            objPlanilhas.sProduto = ""

        ElseIf objPlanilhas.iEscopo = FORMACAO_PRECO_ESCOPO_CATEGORIA Then

            objPlanilhas.iEscopo = FORMACAO_PRECO_ESCOPO_GERAL
            objPlanilhas.sItemCategoria = ""

        End If

        lErro = CF("Planilhas_Le1", objPlanilhas, colPlanilhas)
        If lErro <> SUCESSO Then gError 92432

    End If

    Planilhas_Le1 = SUCESSO

    Exit Function

Erro_Planilhas_Le1:

    Planilhas_Le1 = gErr

    Select Case gErr

        Case 92430, 92431, 92432, 92434

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149191)

    End Select

    Exit Function

End Function

'função criada por Mauricio Maciel em 31/03/2003
Function TipoPlanilha_Le_Todos(ByVal colTipoPlanilha As Collection) As Long
'retorna todos os Tipo de planilha da tabela TipoPLanilha.

Dim lComando As Long
Dim sDescricao As String
Dim iGrupo As Integer
Dim iCodigo As Integer
Dim lErro As Long
Dim objTipoPlanilha As ClassTipoPlanilha

On Error GoTo Erro_TipoPlanilha_Le_Todos

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 116796

    sDescricao = String(STRING_TIPOPLANILHA_DESCRICAO, 0)
    
    'seleciona todos os Tipos de planilha
    lErro = Comando_Executar(lComando, "SELECT Codigo, Descricao, Grupo From TipoPlanilha", iCodigo, sDescricao, iGrupo)
    If lErro <> AD_SQL_SUCESSO Then gError 116797

    'le o primeiro tipo de planilha
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 116798
    
    'Se não encontrou dados => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 116816

    Do While lErro = AD_SQL_SUCESSO

        Set objTipoPlanilha = New ClassTipoPlanilha

        objTipoPlanilha.iCodigo = iCodigo
        objTipoPlanilha.sDescricao = sDescricao
        objTipoPlanilha.iGrupo = iGrupo
        
        colTipoPlanilha.Add objTipoPlanilha

        'le o proximo mnemonico
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 116799

    Loop

    Call Comando_Fechar(lComando)

    TipoPlanilha_Le_Todos = SUCESSO

    Exit Function

Erro_TipoPlanilha_Le_Todos:

    TipoPlanilha_Le_Todos = gErr

    Select Case gErr

        Case 116796
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 116797, 116798, 116799
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOPLANILHA", gErr)

        Case 116816 'Sem dados
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149192)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

'função criada por Mauricio Maciel em 02/04/2003
Function TipoPlanilha_Le(ByVal objTipoPlanilha As ClassTipoPlanilha) As Long
'retorna o Tipo de planilha da tabela TipoPLanilha a partir do parametro codigo.

Dim lComando As Long
Dim sDescricao As String
Dim iGrupo As Integer
Dim lErro As Long

On Error GoTo Erro_TipoPlanilha_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 116801

    sDescricao = String(STRING_TIPOPLANILHA_DESCRICAO, 0)
    
    'seleciona todos os Tipos de planilha
    lErro = Comando_Executar(lComando, "SELECT Descricao, Grupo From TipoPlanilha WHERE Codigo = ?", sDescricao, iGrupo, objTipoPlanilha.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 116802

    'le o primeiro tipo de planilha
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 116803

    'se não encontrou erros
    If lErro = AD_SQL_SEM_DADOS Then gError 116804
        
    objTipoPlanilha.sDescricao = sDescricao
    objTipoPlanilha.iGrupo = iGrupo
    
    Call Comando_Fechar(lComando)

    TipoPlanilha_Le = SUCESSO

    Exit Function

Erro_TipoPlanilha_Le:

    TipoPlanilha_Le = gErr

    Select Case gErr

        Case 116801
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 116802, 116803
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOPLANILHA", gErr)

        Case 116804
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149193)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

'função criada por Mauricio Maciel em 04/04/2003
Function MnemonicoFPTipo_Le_Todos(ByVal objMnemonicoFPTipo As ClassMnemonicoFPTipo, ByVal colMnemonicoFPTipo As Collection) As Long
'retorna todos os mnemonicos da tabela MnemonicoFPreco.

Dim lComando As Long
Dim sMnemonico As String
Dim iTipo As Integer
Dim sMnemonicoDesc As String
Dim lErro As Long
Dim objMnemonicoFPTipo1 As ClassMnemonicoFPTipo
Dim tMnemonicoFPTipo As typeMnemonicoFPTipo

On Error GoTo Erro_MnemonicoFPTipo_Le_Todos

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 92299

    tMnemonicoFPTipo.sMnemonico = String(STRING_MNEMONICOFPRECO_MNEMONICO, 0)
    tMnemonicoFPTipo.sMnemonicoDesc = String(STRING_MNEMONICOFPRECO_MNEMONICODESC, 0)
    tMnemonicoFPTipo.sExpressao = String(STRING_MNEMONICOFPRECO_EXPRESSAO, 0)

    'seleciona todos os mnemonicos
    lErro = Comando_Executar(lComando, "SELECT Funcao, Mnemonico, Tipo, NumParam, Param1, Param2, Param3, MnemonicoDesc, Expressao, TipoPlanilha FROM MnemonicoFPTipo WHERE FilialEmpresa = ? AND Escopo = ? AND ItemCategoria = ? AND Produto = ? AND TabelaPreco = ? AND TipoPlanilha = ? ORDER BY Mnemonico", tMnemonicoFPTipo.iFuncao, tMnemonicoFPTipo.sMnemonico, tMnemonicoFPTipo.iTipo, tMnemonicoFPTipo.iNumParam, tMnemonicoFPTipo.iParam1, tMnemonicoFPTipo.iParam2, tMnemonicoFPTipo.iParam3, tMnemonicoFPTipo.sMnemonicoDesc, tMnemonicoFPTipo.sExpressao, tMnemonicoFPTipo.iTipoPlanilha, objMnemonicoFPTipo.iFilialEmpresa, objMnemonicoFPTipo.iEscopo, objMnemonicoFPTipo.sItemCategoria, objMnemonicoFPTipo.sProduto, objMnemonicoFPTipo.iTabelaPreco, objMnemonicoFPTipo.iTipoPlanilha)
    If lErro <> AD_SQL_SUCESSO Then gError 92300

    'le o primeiro mnemonico
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92301

    'se retornou sem dados
    If lErro = AD_SQL_SEM_DADOS Then gError 116812
    
    Do While lErro = AD_SQL_SUCESSO

        Set objMnemonicoFPTipo1 = New ClassMnemonicoFPTipo

        objMnemonicoFPTipo1.iFilialEmpresa = objMnemonicoFPTipo.iFilialEmpresa
        objMnemonicoFPTipo1.iEscopo = objMnemonicoFPTipo.iEscopo
        objMnemonicoFPTipo1.sItemCategoria = objMnemonicoFPTipo.sItemCategoria
        objMnemonicoFPTipo1.sProduto = objMnemonicoFPTipo.sProduto
        objMnemonicoFPTipo1.iTabelaPreco = objMnemonicoFPTipo.iTabelaPreco
        objMnemonicoFPTipo1.sMnemonico = tMnemonicoFPTipo.sMnemonico
        objMnemonicoFPTipo1.iTipo = tMnemonicoFPTipo.iTipo
        objMnemonicoFPTipo1.iNumParam = tMnemonicoFPTipo.iNumParam
        objMnemonicoFPTipo1.iParam1 = tMnemonicoFPTipo.iParam1
        objMnemonicoFPTipo1.iParam2 = tMnemonicoFPTipo.iParam2
        objMnemonicoFPTipo1.iParam3 = tMnemonicoFPTipo.iParam3
        objMnemonicoFPTipo1.sMnemonicoDesc = tMnemonicoFPTipo.sMnemonicoDesc
        objMnemonicoFPTipo1.sExpressao = tMnemonicoFPTipo.sExpressao
        objMnemonicoFPTipo1.iTipoPlanilha = tMnemonicoFPTipo.iTipoPlanilha
        objMnemonicoFPTipo1.iFuncao = tMnemonicoFPTipo.iFuncao
        
        colMnemonicoFPTipo.Add objMnemonicoFPTipo1

        'le o proximo mnemonico
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92302

    Loop

    Call Comando_Fechar(lComando)

    MnemonicoFPTipo_Le_Todos = SUCESSO

    Exit Function

Erro_MnemonicoFPTipo_Le_Todos:

    MnemonicoFPTipo_Le_Todos = gErr

    Select Case gErr

        Case 116812
        
        Case 92299
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 92300, 92301, 92302
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MNEMONICOFPTIPO", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149194)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function MnemonicoFPlanilha_Le_Tipo(ByVal objMnemonicoFPTipo As ClassMnemonicoFPTipo, ByVal colMnemonicoFPTipo As Collection) As Long
'retorna todos os mnemonicos da tabela MnemonicoFPTipo.

Dim lComando As Long
Dim sMnemonico As String
Dim iTipo As Integer
Dim sMnemonicoDesc As String
Dim lErro As Long
Dim objMnemonicoFPTipo1 As ClassMnemonicoFPTipo
Dim tMnemonicoFPTipo As typeMnemonicoFPTipo
Dim sSelecao As String
Dim sCategoria As String
Dim sProduto As String
Dim sMnemonicoAnterior As String
Dim sItem As String

On Error GoTo Erro_MnemonicoFPlanilha_Le_Tipo

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 92299

    tMnemonicoFPTipo.sMnemonico = String(STRING_MNEMONICOFPRECO_MNEMONICO, 0)
    tMnemonicoFPTipo.sMnemonicoDesc = String(STRING_MNEMONICOFPRECO_MNEMONICODESC, 0)
    tMnemonicoFPTipo.sExpressao = String(STRING_MNEMONICOFPRECO_EXPRESSAO, 0)

    If objMnemonicoFPTipo.iEscopo = MNEMONICOFPRECO_ESCOPO_GERAL Then

        'seleciona todos os mnemonicos
        lErro = Comando_Executar(lComando, "SELECT Mnemonico, Tipo, NumParam, Param1, Param2, Param3, MnemonicoDesc, Expressao, Funcao, TipoPlanilha FROM MnemonicoFPTipo WHERE (FilialEmpresa = 0 OR FilialEmpresa = ?) AND Escopo = ? AND (TipoPlanilha = ? OR TipoPlanilha = ?) ORDER BY Mnemonico, FilialEmpresa DESC", tMnemonicoFPTipo.sMnemonico, tMnemonicoFPTipo.iTipo, tMnemonicoFPTipo.iNumParam, tMnemonicoFPTipo.iParam1, tMnemonicoFPTipo.iParam2, tMnemonicoFPTipo.iParam3, tMnemonicoFPTipo.sMnemonicoDesc, tMnemonicoFPTipo.sExpressao, tMnemonicoFPTipo.iFuncao, tMnemonicoFPTipo.iTipoPlanilha, objMnemonicoFPTipo.iFilialEmpresa, objMnemonicoFPTipo.iEscopo, objMnemonicoFPTipo.iTipoPlanilha, PLANILHA_TIPO_TODOS)
        If lErro <> AD_SQL_SUCESSO Then gError 92407

    ElseIf objMnemonicoFPTipo.iEscopo = MNEMONICOFPRECO_ESCOPO_CATEGORIA Then

        'seleciona todos os mnemonicos
        lErro = Comando_Executar(lComando, "SELECT Mnemonico, Tipo, NumParam, Param1, Param2, Param3, MnemonicoDesc, Expressao, Funcao, TipoPlanilha FROM MnemonicoFPTipo WHERE (FilialEmpresa = 0 OR FilialEmpresa = ?) AND (Escopo = ? OR (Escopo = ? AND ItemCategoria =?)) AND (TipoPlanilha = ? OR TipoPlanilha = ?) ORDER BY Mnemonico, Escopo DESC, FilialEmpresa DESC", tMnemonicoFPTipo.sMnemonico, tMnemonicoFPTipo.iTipo, tMnemonicoFPTipo.iNumParam, tMnemonicoFPTipo.iParam1, tMnemonicoFPTipo.iParam2, tMnemonicoFPTipo.iParam3, tMnemonicoFPTipo.sMnemonicoDesc, tMnemonicoFPTipo.sExpressao, tMnemonicoFPTipo.iFuncao, tMnemonicoFPTipo.iTipoPlanilha, objMnemonicoFPTipo.iFilialEmpresa, MNEMONICOFPRECO_ESCOPO_GERAL, objMnemonicoFPTipo.iEscopo, objMnemonicoFPTipo.sItemCategoria, objMnemonicoFPTipo.iTipoPlanilha, PLANILHA_TIPO_TODOS)
        If lErro <> AD_SQL_SUCESSO Then gError 92408

    ElseIf objMnemonicoFPTipo.iEscopo = MNEMONICOFPRECO_ESCOPO_PRODUTO Then

        sCategoria = CATEGORIAPRODUTO_PRECO
        sProduto = objMnemonicoFPTipo.sProduto

        'retorna o item do produto/categoria
        lErro = CF("ProdutoCategoria_Le", sProduto, sCategoria, sItem)
        If lErro <> SUCESSO And lErro <> 92405 Then gError 92406

        'seleciona todos os mnemonicos
        lErro = Comando_Executar(lComando, "SELECT Mnemonico, Tipo, NumParam, Param1, Param2, Param3, MnemonicoDesc, Expressao, Funcao, TipoPlanilha FROM MnemonicoFPTipo WHERE (FilialEmpresa = 0 OR FilialEmpresa = ?) AND (Escopo = ? OR (Escopo = ? AND ItemCategoria =?) OR (Escopo =? AND Produto =?)) AND (TipoPlanilha = ? OR TipoPlanilha = ?) ORDER BY Mnemonico, Escopo DESC, FilialEmpresa DESC", tMnemonicoFPTipo.sMnemonico, tMnemonicoFPTipo.iTipo, tMnemonicoFPTipo.iNumParam, tMnemonicoFPTipo.iParam1, tMnemonicoFPTipo.iParam2, tMnemonicoFPTipo.iParam3, tMnemonicoFPTipo.sMnemonicoDesc, tMnemonicoFPTipo.sExpressao, tMnemonicoFPTipo.iFuncao, tMnemonicoFPTipo.iTipoPlanilha, objMnemonicoFPTipo.iFilialEmpresa, MNEMONICOFPRECO_ESCOPO_GERAL, MNEMONICOFPRECO_ESCOPO_CATEGORIA, sItem, objMnemonicoFPTipo.iEscopo, objMnemonicoFPTipo.sProduto, objMnemonicoFPTipo.iTipoPlanilha, PLANILHA_TIPO_TODOS)
        If lErro <> AD_SQL_SUCESSO Then gError 92409

    ElseIf objMnemonicoFPTipo.iEscopo = MNEMONICOFPRECO_ESCOPO_TABPRECO Then

        sCategoria = CATEGORIAPRODUTO_PRECO
        sProduto = objMnemonicoFPTipo.sProduto

        'retorna o item do produto/categoria
        lErro = CF("ProdutoCategoria_Le", sProduto, sCategoria, sItem)
        If lErro <> SUCESSO And lErro <> 92405 Then gError 92410

        'seleciona todos os mnemonicos
        lErro = Comando_Executar(lComando, "SELECT Mnemonico, Tipo, NumParam, Param1, Param2, Param3, MnemonicoDesc, Expressao, Funcao, TipoPlanilha FROM MnemonicoFPTipo WHERE (FilialEmpresa = 0 OR FilialEmpresa = ?) AND (Escopo = ? OR (Escopo = ? AND ItemCategoria =?) OR (Escopo =? AND Produto =?) OR (Escopo =? AND Produto =? AND TabelaPreco = ?)) AND (TipoPlanilha = ? OR TipoPlanilha = ?) ORDER BY Mnemonico, Escopo DESC, FilialEmpresa DESC", _
            tMnemonicoFPTipo.sMnemonico, tMnemonicoFPTipo.iTipo, tMnemonicoFPTipo.iNumParam, tMnemonicoFPTipo.iParam1, tMnemonicoFPTipo.iParam2, tMnemonicoFPTipo.iParam3, tMnemonicoFPTipo.sMnemonicoDesc, tMnemonicoFPTipo.sExpressao, tMnemonicoFPTipo.iFuncao, tMnemonicoFPTipo.iTipoPlanilha, objMnemonicoFPTipo.iFilialEmpresa, MNEMONICOFPRECO_ESCOPO_GERAL, MNEMONICOFPRECO_ESCOPO_CATEGORIA, sItem, MNEMONICOFPRECO_ESCOPO_PRODUTO, objMnemonicoFPTipo.sProduto, objMnemonicoFPTipo.iEscopo, objMnemonicoFPTipo.sProduto, objMnemonicoFPTipo.iTabelaPreco, objMnemonicoFPTipo.iTipoPlanilha, PLANILHA_TIPO_TODOS)
        If lErro <> AD_SQL_SUCESSO Then gError 92411

    End If

    'le o primeiro mnemonico
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92301

    Do While lErro = AD_SQL_SUCESSO

        If sMnemonicoAnterior <> tMnemonicoFPTipo.sMnemonico Then

            sMnemonicoAnterior = tMnemonicoFPTipo.sMnemonico

            Set objMnemonicoFPTipo1 = New ClassMnemonicoFPTipo

            objMnemonicoFPTipo1.iFilialEmpresa = objMnemonicoFPTipo.iFilialEmpresa
            objMnemonicoFPTipo1.iEscopo = objMnemonicoFPTipo.iEscopo
            objMnemonicoFPTipo1.sItemCategoria = objMnemonicoFPTipo.sItemCategoria
            objMnemonicoFPTipo1.sProduto = objMnemonicoFPTipo.sProduto
            objMnemonicoFPTipo1.iTabelaPreco = objMnemonicoFPTipo.iTabelaPreco
            objMnemonicoFPTipo1.sMnemonico = tMnemonicoFPTipo.sMnemonico
            objMnemonicoFPTipo1.iTipo = tMnemonicoFPTipo.iTipo
            objMnemonicoFPTipo1.iNumParam = tMnemonicoFPTipo.iNumParam
            objMnemonicoFPTipo1.iParam1 = tMnemonicoFPTipo.iParam1
            objMnemonicoFPTipo1.iParam2 = tMnemonicoFPTipo.iParam2
            objMnemonicoFPTipo1.iParam3 = tMnemonicoFPTipo.iParam3
            objMnemonicoFPTipo1.sMnemonicoDesc = tMnemonicoFPTipo.sMnemonicoDesc
            objMnemonicoFPTipo1.sExpressao = tMnemonicoFPTipo.sExpressao
            objMnemonicoFPTipo1.iFuncao = tMnemonicoFPTipo.iFuncao
            objMnemonicoFPTipo1.iTipoPlanilha = tMnemonicoFPTipo.iTipoPlanilha
            
            colMnemonicoFPTipo.Add objMnemonicoFPTipo1

        End If

        'le o proximo mnemonico
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92302

    Loop

    Call Comando_Fechar(lComando)

    MnemonicoFPlanilha_Le_Tipo = SUCESSO

    Exit Function

Erro_MnemonicoFPlanilha_Le_Tipo:

    MnemonicoFPlanilha_Le_Tipo = gErr

    Select Case gErr

        Case 92299
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 92406, 92410

        Case 92300, 92301, 92302, 92406, 92407, 92408, 92409, 92411
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MNEMONICOFPTIPO", gErr)

        Case 92412, 92413
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTOCATEGORIA_NAO_CADASTRADO", gErr, sCategoria, sProduto)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149195)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function MnemonicoFPTipo_Le(ByVal objMnemonicoFPTipo As ClassMnemonicoFPTipo) As Long

Dim lErro As Long, lComando As Long
Dim tMnemonicoFPTipo As typeMnemonicoFPTipo

On Error GoTo Erro_MnemonicoFPTipo_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 106909
    
    With tMnemonicoFPTipo
        .sMnemonicoDesc = String(STRING_MNEMONICOFPRECO_MNEMONICODESC, 0)
        .sExpressao = String(STRING_MNEMONICOFPRECO_EXPRESSAO, 0)
            
        lErro = Comando_Executar(lComando, "SELECT Tipo, NumParam, Param1, Param2, Param3, MnemonicoDesc, Expressao, Funcao FROM MnemonicoFPTipo WHERE FilialEmpresa = ? AND Escopo = ? AND TipoPlanilha = ? AND ItemCategoria = ? AND Produto = ? AND TabelaPreco = ? AND Mnemonico = ?", _
            .iTipo, .iNumParam, .iParam1, .iParam2, .iParam3, .sMnemonicoDesc, .sExpressao, .iFuncao, _
            objMnemonicoFPTipo.iFilialEmpresa, objMnemonicoFPTipo.iEscopo, objMnemonicoFPTipo.iTipoPlanilha, objMnemonicoFPTipo.sItemCategoria, objMnemonicoFPTipo.sProduto, objMnemonicoFPTipo.iTabelaPreco, objMnemonicoFPTipo.sMnemonico)
    End With
    If lErro <> AD_SQL_SUCESSO Then gError 106910

    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106911
    If lErro <> AD_SQL_SUCESSO Then gError 106912
    
    With objMnemonicoFPTipo
        .iTipo = tMnemonicoFPTipo.iTipo
        .iNumParam = tMnemonicoFPTipo.iNumParam
        .iParam1 = tMnemonicoFPTipo.iParam1
        .iParam2 = tMnemonicoFPTipo.iParam2
        .iParam3 = tMnemonicoFPTipo.iParam3
        .sMnemonicoDesc = tMnemonicoFPTipo.sMnemonicoDesc
        .sExpressao = tMnemonicoFPTipo.sExpressao
        .iFuncao = tMnemonicoFPTipo.iFuncao
    End With

    Call Comando_Fechar(lComando)
    
    MnemonicoFPTipo_Le = SUCESSO
     
    Exit Function
    
Erro_MnemonicoFPTipo_Le:

    MnemonicoFPTipo_Le = gErr
     
    Select Case gErr
          
        Case 106909
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 106910, 106911
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MNEMONICOFPTIPO", gErr)
        
        Case 106912 'nao encontrou
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149196)
     
    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function ClassificacaoFiscal_Le(objClassificacaoFiscal As ClassClassificacaoFiscal) As Long
'Le as informações sobre o código passado

Dim lErro As Long
Dim lComando As Long
Dim tClassificacaoFiscal As typeClassificacaoFiscal

On Error GoTo Erro_ClassificacaoFiscal_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 123491
    
    'Inicializa as strings
    tClassificacaoFiscal.sDescricao = String(STRING_CLASSIFICACAOFISCAL_DESCRICAO, 0)
    
    'Realiza a seleção atrvés do Código que foi passado
    lErro = Comando_Executar(lComando, "SELECT Descricao, IIAliquota, IPIAliquota, PISAliquota, COFINSAliquota, ICMSAliquota FROM ClassificacaoFiscal WHERE Codigo = ?", tClassificacaoFiscal.sDescricao, tClassificacaoFiscal.dIIAliquota, tClassificacaoFiscal.dIPIAliquota, tClassificacaoFiscal.dPISAliquota, tClassificacaoFiscal.dCOFINSAliquota, tClassificacaoFiscal.dICMSAliquota, objClassificacaoFiscal.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 123492
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 123493
    
    If lErro <> AD_SQL_SUCESSO Then gError 123494
    
    objClassificacaoFiscal.sDescricao = tClassificacaoFiscal.sDescricao
    objClassificacaoFiscal.dIIAliquota = tClassificacaoFiscal.dIIAliquota
    objClassificacaoFiscal.dPISAliquota = tClassificacaoFiscal.dPISAliquota
    objClassificacaoFiscal.dCOFINSAliquota = tClassificacaoFiscal.dCOFINSAliquota
    objClassificacaoFiscal.dICMSAliquota = tClassificacaoFiscal.dICMSAliquota
    objClassificacaoFiscal.dIPIAliquota = tClassificacaoFiscal.dIPIAliquota
        
    'Executa o fechamento do Comando
    Call Comando_Fechar(lComando)
    
    ClassificacaoFiscal_Le = SUCESSO
    
    Exit Function
    
Erro_ClassificacaoFiscal_Le:

    ClassificacaoFiscal_Le = gErr
    
    Select Case gErr

        Case 123491
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 123492, 123493
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLASSIFICACAOFISCAL", gErr, objClassificacaoFiscal.sCodigo)
        
        Case 123494 'Código não encontrado no Banco de Dados
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149197)


    End Select
    
    'Executa o fechamento do Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ClassificacaoFiscal_Lock(ByVal sCodigo As String) As Long
'Faz lock shared de uma Código para garantir a sua existência

Dim lErro As Long
Dim sCodigoClasFis As String
Dim lComando As Long

On Error GoTo Erro_ClassificacaoFiscal_Lock

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 9999

    sCodigoClasFis = String(STRING_PRODUTO_IPI_CODIGO, 0)

    'Pesquisa o Código da Clasificação Fiscal
    lErro = Comando_ExecutarLockado(lComando, "SELECT Codigo FROM ClassificacaoFiscal WHERE Codigo = ?", sCodigoClasFis, sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 125060

    'Tenta ler
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 125061

    'Se não encontrou --> ERRO
    If lErro = AD_SQL_SEM_DADOS Then gError 125062

    '"Lock" da ClassificacaoFiscal
    lErro = Comando_LockShared(lComando)
    If lErro <> SUCESSO Then gError 125063
    
    Call Comando_Fechar(lComando)
    
    ClassificacaoFiscal_Lock = SUCESSO
    
    Exit Function
    
Erro_ClassificacaoFiscal_Lock:

    ClassificacaoFiscal_Lock = gErr
    
    Select Case gErr
    
        Case 125060, 125061
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLASSIFICACAOFISCAL", gErr, sCodigo)
        
        Case 125062 'Código não existe no BD
        
        Case 125063
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_CLASSIFICACAOFISCAL", gErr, sCodigo)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149198)

    End Select
    
    Call Comando_Fechar(lComando)

    Call Transacao_Rollback
    
    Exit Function
    
End Function

'Incluída por Luiz Nogueira em 27/10/03
Public Function RelacionamentoClientes_Le(ByVal objRelacionamentoClientes As ClassRelacClientes) As Long
'objRelacionamentoClientes RECEBE(Input) código/filialempresa do relacionamento que deve ser lido
'objRelacionamentoClientes RETORNA(Output) os dados do relacionamento lido

Dim lErro As Long
Dim lComando As Long
Dim tRelacionamentoClientes As typeRelacionamentoClientes

On Error GoTo Erro_RelacionamentoClientes_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 102505
    
    'Inicializa as strings para leitura de dados
    tRelacionamentoClientes.sAssunto1 = String(STRING_BUFFER_MAX_TEXTO, 0)
    tRelacionamentoClientes.sAssunto2 = String(STRING_BUFFER_MAX_TEXTO, 0)
    
    'Lê no BD o relacionamento com código e filialempresa passados como parâmetro
    lErro = Comando_Executar(lComando, "SELECT Codigo, FilialEmpresa, Origem, Data, Hora, Tipo, Cliente, FilialCliente, Contato, Atendente, RelacionamentoAnt, Assunto1, Assunto2, Status, NumIntParcRec, DataProxCobr, DataPrevReceb, StatusCG, DataFim, HoraFim, Motivo, TipoDoc, NumIntDocOrigem, Satisfacao FROM RelacionamentoClientes WHERE Codigo = ? AND FilialEmpresa = ?", tRelacionamentoClientes.lCodigo, _
    tRelacionamentoClientes.iFilialEmpresa, tRelacionamentoClientes.iOrigem, tRelacionamentoClientes.dtData, tRelacionamentoClientes.dHora, tRelacionamentoClientes.lTipo, tRelacionamentoClientes.lCliente, tRelacionamentoClientes.iFilialCliente, tRelacionamentoClientes.iContato, tRelacionamentoClientes.iAtendente, _
    tRelacionamentoClientes.lRelacionamentoAnt, tRelacionamentoClientes.sAssunto1, tRelacionamentoClientes.sAssunto2, tRelacionamentoClientes.iStatus, tRelacionamentoClientes.lNumIntParcRec, tRelacionamentoClientes.dtDataProxCobr, tRelacionamentoClientes.dtDataPrevReceb, tRelacionamentoClientes.iStatusCG, tRelacionamentoClientes.dtDataFim, tRelacionamentoClientes.dHoraFim, tRelacionamentoClientes.lMotivo, tRelacionamentoClientes.iTipoDoc, tRelacionamentoClientes.lNumIntDocOrigem, tRelacionamentoClientes.lSatisfacao, _
    objRelacionamentoClientes.lCodigo, objRelacionamentoClientes.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 102506
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 102507
    
    'Se não encontrou o relacionamento => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 102508
    
    'Transfere os dados lidos para o obj
    With objRelacionamentoClientes
    
        .lCodigo = tRelacionamentoClientes.lCodigo
        .iFilialEmpresa = tRelacionamentoClientes.iFilialEmpresa
        .iOrigem = tRelacionamentoClientes.iOrigem
        .dtData = tRelacionamentoClientes.dtData
        .dtHora = CDate(tRelacionamentoClientes.dHora)
        .lTipo = tRelacionamentoClientes.lTipo
        .lCliente = tRelacionamentoClientes.lCliente
        .iFilialCliente = tRelacionamentoClientes.iFilialCliente
        .iContato = tRelacionamentoClientes.iContato
        .iAtendente = tRelacionamentoClientes.iAtendente
        .lRelacionamentoAnt = tRelacionamentoClientes.lRelacionamentoAnt
        .sAssunto1 = tRelacionamentoClientes.sAssunto1
        .sAssunto2 = tRelacionamentoClientes.sAssunto2
        .iStatus = tRelacionamentoClientes.iStatus
        .lNumIntParcRec = tRelacionamentoClientes.lNumIntParcRec
        .dtDataPrevReceb = tRelacionamentoClientes.dtDataPrevReceb
        .dtDataProxCobr = tRelacionamentoClientes.dtDataProxCobr
        .iStatusCG = tRelacionamentoClientes.iStatusCG
        .dtDataFim = tRelacionamentoClientes.dtDataFim
        .dtHoraFim = CDate(tRelacionamentoClientes.dHoraFim)
        .lMotivo = tRelacionamentoClientes.lMotivo
        .iTipoDoc = tRelacionamentoClientes.iTipoDoc
        .lNumIntDocOrigem = tRelacionamentoClientes.lNumIntDocOrigem
        .lSatisfacao = tRelacionamentoClientes.lSatisfacao
        
        If .iTipoDoc = RELACCLI_TIPODOC_PARCELARECEBER And .lNumIntParcRec <> 0 Then
            .lNumIntDocOrigem = .lNumIntParcRec
        End If
    
    End With
    
    lErro = RelacCliTexto_Le(objRelacionamentoClientes)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Executa o fechamento do Comando
    Call Comando_Fechar(lComando)

    RelacionamentoClientes_Le = SUCESSO

    Exit Function

Erro_RelacionamentoClientes_Le:

    RelacionamentoClientes_Le = gErr

    Select Case gErr

        Case 102505
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 102506, 102507
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELACIONAMENTOCLIENTES", gErr)
            
        Case 102508 'Registro não encontrado
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149199)

    End Select

    'Executa o fechamento do Comando
    Call Comando_Fechar(lComando)

End Function

'Incluída por Luiz Nogueira em 27/10/03
Public Function RelacionamentoClientes_Locks_Relacionados(ByVal objRelacionamentoClientes As ClassRelacClientes) As Long
'Efetua locks nas tabelas relacionadas à gravação de relacionamentos

Dim lErro As Long

On Error GoTo Erro_RelacionamentoClientes_Locks_Relacionados

    'Verifica se Cliente e Filial estão cadastrados
    lErro = CF("FilialCliente_Lock", objRelacionamentoClientes.lCliente, objRelacionamentoClientes.iFilialCliente)
    If lErro <> SUCESSO And lErro <> 43050 Then gError 102554
    
    'Não encontrou a Filial Cliente --> erro
    If lErro = 43050 Then gError 102555
    
    'Efetua lock do atendente
    lErro = CF("Atendentes_LockShared", objRelacionamentoClientes.iAtendente, objRelacionamentoClientes.iFilialEmpresa)
    If lErro <> SUCESSO And lErro <> 102552 Then gError 102556
    
    'Se não encontrou o campo
    If lErro = 102552 Then gError 102557
    
    'Efetua lock do Relacionamento Anterior
    lErro = CF("RelacionamentoClientes_LockShared", objRelacionamentoClientes.lRelacionamentoAnt, objRelacionamentoClientes.iFilialEmpresa)
    If lErro <> SUCESSO And lErro <> 102562 Then gError 102564
    
    RelacionamentoClientes_Locks_Relacionados = SUCESSO

    Exit Function

Erro_RelacionamentoClientes_Locks_Relacionados:

    RelacionamentoClientes_Locks_Relacionados = gErr

    Select Case gErr
    
        Case 102554, 102556
        
        Case 102555
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_CADASTRADA", gErr, objRelacionamentoClientes.iFilialCliente, objRelacionamentoClientes.lCliente)

        Case 102557
            Call Rotina_Erro(vbOKOnly, "ERRO_ATENDENTE_NAO_ENCONTRADO", gErr, objRelacionamentoClientes.iAtendente, objRelacionamentoClientes.iFilialEmpresa)
        
        Case 102564
            Call Rotina_Erro(vbOKOnly, "ERRO_RELACIONAMENTO_NAO_ENCONTRADO", gErr, objRelacionamentoClientes.lRelacionamentoAnt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149200)

    End Select

End Function

'Incluído por Luiz Nogueira em 27/10/03
Public Function RelacionamentoClientes_LockShared(ByVal lCodigo As Long, Optional ByVal iFilialEmpresa As Integer) As Long
'Faz lock shared no relacionamento passado como parâmetro
'lCodigo recebe o código do relacionamento a ser "lockado"

Dim lErro As Long
Dim lComando As Long
Dim iIndice As Integer
Dim lCodAux As Long

On Error GoTo Erro_RelacionamentoClientes_LockShared

    'Executa a abertura dos Comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 102559

    'Se não indicou a filial do atendente => usa a filial empresa ativa
    If iFilialEmpresa = 0 Then iFilialEmpresa = giFilialEmpresa
    
    'Busca o relacionamento a ser lockado
    lErro = Comando_ExecutarLockado(lComando, "SELECT Codigo FROM RelacionamentoClientes WHERE Codigo=? AND FilialEmpresa=?", lCodAux, lCodigo, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 102560
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 102561
    
    'Se não encontrou o relacionamento => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 102562
    
    'Efetua o lock do relacionamento
    lErro = Comando_LockShared(lComando)
    If lErro <> AD_SQL_SUCESSO Then gError 102563
    
    'Executa o fechamento dos Comandos
    Call Comando_Fechar(lComando)

    RelacionamentoClientes_LockShared = SUCESSO

    Exit Function

Erro_RelacionamentoClientes_LockShared:

    RelacionamentoClientes_LockShared = gErr

    Select Case gErr

        Case 102559
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 102560, 102561
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELACIONAMENTOCLIENTES", gErr)

        Case 102562

        Case 102563
            Call Rotina_Erro(vbOKOnly, "ERRO_CAMPOGENERICO_LOCKSHARED", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149201)

    End Select

    Call Comando_Fechar(lComando)

End Function

'Incluída por Luiz Nogueira em 27/10/03
Public Function RelacionamentoClientes_Verifica_Cliente(ByVal lCliente As Long, lCodRelacionamento As Long, Optional ByVal iFilialCliente As Integer) As Long
'Verifica se existe algum relacionamento para o cliente ou filial do cliente em questão
'lCliente recebe o código do cliente para o qual deve ser verificado se existe relacionamento
'iFilialCliente recebe o código da filial para a qual deve ser feita a mesma verificação
'Se iFilialCliente for 0, a verificação será feita para todas as filiais do cliente

Dim lErro As Long
Dim lComando As Long

On Error GoTo Erro_RelacionamentoClientes_Verifica_Cliente

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 102695
    
    'Se foi passada uma filial de Cliente
    If iFilialCliente <> 0 Then
    
        'Verifica se existe relacionamento para a filial do cliente com código passado como parâemtro
        lErro = Comando_Executar(lComando, "SELECT Codigo FROM RelacionamentoClientes WHERE Cliente=? AND FilialCliente=?", lCodRelacionamento, lCliente, iFilialCliente)
        If lErro <> AD_SQL_SUCESSO Then gError 102713
    
    'Senão, ou seja, se não foi passada uma filial específica
    Else
    
        'Verifica se existe relacionamento para qualquer filial do cliente com código passado como parâemtro
        lErro = Comando_Executar(lComando, "SELECT Codigo FROM RelacionamentoClientes WHERE Cliente=?", lCodRelacionamento, lCliente)
        If lErro <> AD_SQL_SUCESSO Then gError 102692
    
    End If
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 102693
    
    'Se não encontrou => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 102694

    'Executa o fechamento do Comando
    Call Comando_Fechar(lComando)

    RelacionamentoClientes_Verifica_Cliente = SUCESSO

    Exit Function

Erro_RelacionamentoClientes_Verifica_Cliente:

    RelacionamentoClientes_Verifica_Cliente = gErr

    Select Case gErr

        Case 102695
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 102692, 102693, 102713
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELACIONAMENTOCLIENTES", gErr)
        
        Case 102694 'Não encontrado

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149202)

    End Select

    'Executa o fechamento do Comando
    Call Comando_Fechar(lComando)

End Function

'Incluída por Luiz Nogueira em 27/10/03
Public Function RelacionamentoClientes_Valida_Alteracao(ByVal objRelacionamentoClientes As ClassRelacClientes, bAlteracaoValida As Boolean) As Long
'Verifica se o relacionamento passado como parâmetro está sendo apontado por outro relacionamento como relacionamento anterior
'e com a data anterior a este relacionamento
'Isso é para evitar que a data de um relacionamento seja alterada e torne inconsistente os demais relacionamentos que o utilizam
'como relacionamento anterior

Dim lErro As Long
Dim lComando As Long
Dim lCodigo As Long

On Error GoTo Erro_RelacionamentoClientes_Valida_Alteracao

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 102707

    'Verifica se o relacionamento que está sendo alterado é apontado por outro relacionamento
    'como relacionamento anterior
    lErro = Comando_Executar(lComando, "SELECT Codigo FROM RelacionamentoClientes WHERE RelacionamentoAnt=? AND FilialEmpresa=? AND (Cliente<>? OR Data<?)", lCodigo, objRelacionamentoClientes.lCodigo, objRelacionamentoClientes.iFilialEmpresa, objRelacionamentoClientes.lCliente, objRelacionamentoClientes.dtData)
    If lErro <> AD_SQL_SUCESSO Then gError 102708
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 102709
    
    'Se encontrou => significa que a data do relacionamento não pode ser alterada
    If lErro = AD_SQL_SUCESSO Then bAlteracaoValida = False
    
    'Executa o fechamento do Comando
    Call Comando_Fechar(lComando)

    RelacionamentoClientes_Valida_Alteracao = SUCESSO

    Exit Function

Erro_RelacionamentoClientes_Valida_Alteracao:

    RelacionamentoClientes_Valida_Alteracao = gErr

    Select Case gErr

        Case 102707
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 102708, 102709
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELACIONAMENTOCLIENTES", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149203)

    End Select

    'Executa o fechamento do Comando
    Call Comando_Fechar(lComando)

End Function

'Incluída por Luiz Nogueira em 27/10/03
Public Function Atendentes_Le(ByVal objAtendente As ClassAtendentes) As Long
'Le o atendente e os dados do usuário relacionados ao atendente conforme código e filial passados como parâmetros

Dim lErro As Long
Dim lComando As Long
Dim tUsuario As typeUsuario

On Error GoTo Erro_Atendentes_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 102749

    'Inicializa string para leitura no bd
    tUsuario.sCodUsuario = String(STRING_USUARIO_CODIGO, 0)
    tUsuario.sNome = String(STRING_USUARIO_NOME, 0)
    tUsuario.sNomeReduzido = String(STRING_USUARIO_NOMEREDUZIDO, 0)

    'Lê os dados do atendente
    lErro = Comando_Executar(lComando, "SELECT Atendentes.CodUsuario,Usuario.Nome, Usuario.NomeReduzido FROM Atendentes, Usuario WHERE Atendentes.CodUsuario = Usuario.CodUsuario AND Atendentes.Codigo=? AND Atendentes.FilialEmpresa=?", tUsuario.sCodUsuario, tUsuario.sNome, tUsuario.sNomeReduzido, objAtendente.iCodigo, objAtendente.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 102750

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 102751

    'Se não encontrou => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 102752

    'Guarda no obj os dados lidos
    With objAtendente
        .sCodUsuario = tUsuario.sCodUsuario
        .sNome = tUsuario.sNome
        .sNomeReduzido = tUsuario.sNomeReduzido
    End With

    Call Comando_Fechar(lComando)

    Atendentes_Le = SUCESSO

    Exit Function

Erro_Atendentes_Le:

    Atendentes_Le = gErr

    Select Case gErr

        Case 102749
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 102750, 102751
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ATENDENTES", gErr)

        Case 102752 'Sem dados

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149204)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

'Incluída por Luiz Nogueira em 27/10/03
Public Function Atendentes_Le_Usuario(ByVal objAtendente As ClassAtendentes) As Long
'Le os dados de um atendente a partir de do código do usuário

Dim lErro As Long
Dim lComando As Long
Dim iCodigo As Integer

On Error GoTo Erro_Atendentes_Le_Usuario

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 102753

    'Lê os dados do atendente
    lErro = Comando_Executar(lComando, "SELECT Codigo FROM Atendentes WHERE CodUsuario=? AND FilialEmpresa=?", iCodigo, objAtendente.sCodUsuario, objAtendente.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 102754

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 102755

    'Se não encontrou o atendente => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 102756

    'Guarda os dados lidos no obj
    objAtendente.iCodigo = iCodigo

    Call Comando_Fechar(lComando)

    Atendentes_Le_Usuario = SUCESSO

    Exit Function

Erro_Atendentes_Le_Usuario:

    Atendentes_Le_Usuario = gErr

    Select Case gErr

        Case 102753
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 102754, 102755
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ATENDENTES", gErr, objAtendente.sCodUsuario)

        Case 102756 'Sem dados

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149205)

    End Select

    Call Comando_Fechar(lComando)

End Function

'Incluído por Luiz Nogueira em 27/10/03
Function Atendentes_Le_Todos(ByVal colAtendentes As Collection, Optional iFilialEmpresa As Integer) As Long
'Lê todos os atendentes da filial empresa passada como parâmetro
'Se não for passada uma filial, lê os atendentes de giFilialEmpresa

Dim lErro As Long
Dim lComando As Long
Dim objAtendente As ClassAtendentes
Dim iCodigo As Integer
Dim sCodUsuario As String
Dim tUsuario As typeUsuario

On Error GoTo Erro_Atendentes_Le_Todos

    'Abre comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 102757

    'Inicializa string para leitura no bd
    tUsuario.sCodUsuario = String(STRING_USUARIO_CODIGO, 0)
    tUsuario.sNome = String(STRING_USUARIO_NOME, 0)
    tUsuario.sNomeReduzido = String(STRING_USUARIO_NOMEREDUZIDO, 0)
    
    'Se não foi passada uma filial empresa => lê os atendentes de giFilialEmpresa
    If iFilialEmpresa = 0 Then iFilialEmpresa = giFilialEmpresa

    'Lê os atendentes
    lErro = Comando_Executar(lComando, "SELECT Atendentes.Codigo, Atendentes.CodUsuario, Usuario.Nome, Usuario.NomeReduzido FROM Atendentes, Usuario WHERE Atendentes.CodUsuario = Usuario.CodUsuario AND Atendentes.FilialEmpresa=?", iCodigo, tUsuario.sCodUsuario, tUsuario.sNome, tUsuario.sNomeReduzido, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 102758

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 102759
    
    'Se não encontrou => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 102761

    'Enquanto houverem atendentes
    Do While lErro = AD_SQL_SUCESSO

        'Instancia um novo obj
        Set objAtendente = New ClassAtendentes

        'Guarda no obj os dados lidos
        With objAtendente
            .iCodigo = iCodigo
            .sCodUsuario = tUsuario.sCodUsuario
            .sNome = tUsuario.sNome
            .sNomeReduzido = tUsuario.sNomeReduzido
        End With

        'Guarda o obj na coleção
        colAtendentes.Add objAtendente

        'Busca o próximo atendente lido
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 102760

    Loop

    Call Comando_Fechar(lComando)

    Atendentes_Le_Todos = SUCESSO

    Exit Function

Erro_Atendentes_Le_Todos:

    Atendentes_Le_Todos = gErr

    Select Case gErr

        Case 102757
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 102758 To 102760
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ATENDENTES", gErr)

        Case 102761 'Sem dados

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149206)

    End Select

    Call Comando_Fechar(lComando)

End Function

'Incluída por Luiz Nogueira em 27/10/03
Public Function Atendentes_Valida_Exclusao(ByVal objAtendente As ClassAtendentes) As Long
'Verifica se um atendente não está sendo utilizado em outras tabelas do sistema
'objAtendente recebe os dados do atendente que deve ser checado

Dim lErro As Long
Dim lComando As Long
Dim lCodRelacionamento As Long

On Error GoTo Erro_Atendentes_Valida_Exclusao

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 102773
    
    '*** VERIFICAÇÃO NA TABELA RELACIONAMENTOCLIENTES ***
    'Busca em relacionamentoclientes um atendente da filial empresa passada como parâmetro
    'e com o código também passado como parâmetro
    lErro = Comando_Executar(lComando, "SELECT Codigo FROM RelacionamentoClientes WHERE Atendente=? AND FilialEmpresa=?", lCodRelacionamento, objAtendente.iCodigo, objAtendente.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 102774
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 102775
    
    'Se encontrou relacionamento => erro de inconsistência
    If lErro = AD_SQL_SUCESSO Then gError 102776
    
    '*** VERIFICAÇÃO NA TABELA RELACIONAMENTOCLIENTES ***
    'Busca em relacionamentoclientes um atendente da filial empresa passada como parâmetro
    'e com o código também passado como parâmetro
    lErro = Comando_Executar(lComando, "SELECT Codigo FROM RelacionamentoContatos WHERE Atendente=? AND FilialEmpresa=?", lCodRelacionamento, objAtendente.iCodigo, objAtendente.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 102774
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 102775
    
    'Se encontrou relacionamento => erro de inconsistência
    If lErro = AD_SQL_SUCESSO Then gError 102776

    'Executa o fechamento do Comando
    Call Comando_Fechar(lComando)

    Atendentes_Valida_Exclusao = SUCESSO

    Exit Function

Erro_Atendentes_Valida_Exclusao:

    Atendentes_Valida_Exclusao = gErr

    Select Case gErr

        Case 102773
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 102774, 102775
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELACIONAMENTOCLIENTES", gErr)
        
        Case 102776
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ATENDENTE_VINCULADO", gErr, objAtendente.iCodigo, lCodRelacionamento)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149207)

    End Select

    'Executa o fechamento do Comando
    Call Comando_Fechar(lComando)

End Function

'Incluída por Luiz Nogueira em 27/10/03
Public Function Carrega_Atendentes(ByVal objControle As Object) As Long
'Carrega a combo ou list atendentes com os atendentes da filial empresa ativa

Dim lErro As Long
Dim objAtendente As New ClassAtendentes
Dim colUsuarios As New Collection
Dim objUsuarios As New ClassUsuarios
Dim colAtendentes As New Collection
Dim sTipoControle As String

On Error GoTo Erro_Carrega_Atendentes

    'Obtém o tipo do controle recebido como parâmetro
    sTipoControle = TypeName(objControle)
    
    'Lê todos os usuários do sistema
    lErro = CF("Usuarios_Le_Todos", colUsuarios)
    If lErro <> SUCESSO Then gError 102747

    'Le todos os Atendentes da Filial Empresa
    lErro = CF("Atendentes_Le_Todos", colAtendentes)
    If lErro <> SUCESSO And lErro <> 102761 Then gError 102748

    'Para cada atendente encontrado
    For Each objAtendente In colAtendentes
    
        'Para cada usuário encontrado
        For Each objUsuarios In colUsuarios
            
            'Se o código do atendente for igual ao código do usuário
            If objAtendente.sCodUsuario = objUsuarios.sCodUsuario Then
                
                'Se o controle for tipo combobox
                If sTipoControle = "ComboBox" Then
                
                    'Adiciona o atendente à lista de atendentes, exibindo código e nomereduzido
                    objControle.AddItem objAtendente.iCodigo & SEPARADOR & objUsuarios.sNomeReduzido
                    
                    'Se o atendente adicionado for o atendente do usuário ativo => 'Seleciona o atendente na combo
                    If gsUsuario = objUsuarios.sCodUsuario Then objControle.ListIndex = objControle.NewIndex
                    
                Else
                
                    'Adiciona o atendente à lista de atendentes, exibindo código e nomereduzido
                    objControle.AddItem objUsuarios.sNomeReduzido
                
                End If
                
                'Adiciona o código do atendente ao item data
                objControle.ItemData(objControle.NewIndex) = objAtendente.iCodigo
            
            End If
        
        Next
    
    Next

    Carrega_Atendentes = SUCESSO

    Exit Function

Erro_Carrega_Atendentes:

    Carrega_Atendentes = gErr

    Select Case gErr

        Case 102747, 102748

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149208)

    End Select

End Function

'Incluída por Luiz Nogueira em 27/10/03
Public Function Atendente_Validate(ByVal objComboBox As Object, Optional ByVal bCria As Boolean = True) As Long
'Faz a validação da combo atendente passado como parâmetro

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objAtendente As New ClassAtendentes
Dim iCodigo As Integer

On Error GoTo Erro_Atendente_Validate

    'Se o atendente não foi preenchido => sai da função
    If Len(Trim(objComboBox.Text)) = 0 Then Exit Function
    
    'Se o atendente foi selecionado na própria combo => sai da função
    If objComboBox.Text = objComboBox.List(objComboBox.ListIndex) Then Exit Function
    
    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(objComboBox, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 102816

    'Se não encontrou o atendente na combo, mas retornou um código
    If lErro = 6730 Then
    
        'Guarda os dados necessários para tentar ler o atendente
        objAtendente.iCodigo = iCodigo
        objAtendente.iFilialEmpresa = giFilialEmpresa
        
        'Lê o atendente a partir dos dados passados
        lErro = CF("Atendentes_Le", objAtendente)
        If lErro <> SUCESSO And lErro <> 102752 Then gError 102817
        
        'Se não encontrou o atendente
        If lErro = 102752 Then gError 102818
        
        'Exibe o atendente na tela
        objComboBox.Text = objAtendente.iCodigo & SEPARADOR & objAtendente.sNomeReduzido
        
    End If
    
    'Se foi digitado o nome do atendente
    'e esse nome não foi encontrado na combo => erro
    If lErro = 6731 Then
    
        'Guarda os dados necessário para tentar ler o atendente a partir do nome reduzido
        objAtendente.sNomeReduzido = Trim(objComboBox.Text)
        objAtendente.iFilialEmpresa = giFilialEmpresa
        
        'Lê o atendente a partir dos dados passados
        lErro = CF("Atendentes_Le_NomeReduzido", objAtendente)
        If lErro <> SUCESSO And lErro <> 102826 Then gError 102817
        
        'Se não encontrou o atendente
        If lErro = 102826 Then gError 102819
        
        'Exibe o atendente na tela
        objComboBox.Text = objAtendente.iCodigo & SEPARADOR & objAtendente.sNomeReduzido
    
    End If

    Atendente_Validate = SUCESSO
    
    Exit Function

Erro_Atendente_Validate:

    Atendente_Validate = gErr

    Select Case gErr

        Case 102816, 102817
        
        Case 102818, 102819
            
            'Se é para perguntar ao usuário se deve criar o atendente
            If bCria Then
            
                'Verifica se o usuário deseja criar um novo contato
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_ATENDENTE", Trim(objComboBox.Text), giFilialEmpresa)
    
                'Se o usuário respondeu sim
                If vbMsgRes = vbYes Then
                
                    'Guarda o nome reduzido no objatendente
                    objAtendente.sNomeReduzido = Trim(objComboBox.Text)
                    
                    'Chama a tela para cadastro de contatos
                    Call Chama_Tela("Atendentes", objAtendente)
                End If

            'Senão
            Else
            
                'Msg de erro
                Call Rotina_Erro(vbOKOnly, "ERRO_ATENDENTE_NAO_ENCONTRADO", gErr, Trim(objComboBox.Text), giFilialEmpresa)
            
            End If
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149209)

    End Select

End Function

'Incluída por Luiz Nogueira em 27/10/03
Public Function Atendentes_Le_NomeReduzido(ByVal objAtendente As ClassAtendentes) As Long
'Lê o atendente usando o nome reduzido do usuário vinculado a ele e a filial empresa do atendente

Dim lErro As Long
Dim lComando As Long
Dim tUsuario As typeUsuario
Dim iCodigo As Integer

On Error GoTo Erro_Atendentes_Le_NomeReduzido

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 102823

    'Inicializa string para leitura no bd
    tUsuario.sCodUsuario = String(STRING_USUARIO_CODIGO, 0)
    tUsuario.sNome = String(STRING_USUARIO_NOME, 0)

    'Lê os dados do atendente
    lErro = Comando_Executar(lComando, "SELECT Atendentes.Codigo, Atendentes.CodUsuario, Usuario.Nome FROM Atendentes, Usuario WHERE Atendentes.CodUsuario = Usuario.CodUsuario AND Usuario.NomeReduzido=? AND Atendentes.FilialEmpresa=?", iCodigo, tUsuario.sCodUsuario, tUsuario.sNome, objAtendente.sNomeReduzido, objAtendente.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 102824

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 102825

    'Se não encontrou => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 102826

    'Guarda no obj os dados lidos
    With objAtendente
        .iCodigo = iCodigo
        .sCodUsuario = tUsuario.sCodUsuario
        .sNome = tUsuario.sNome
    End With

    Call Comando_Fechar(lComando)

    Atendentes_Le_NomeReduzido = SUCESSO

    Exit Function

Erro_Atendentes_Le_NomeReduzido:

    Atendentes_Le_NomeReduzido = gErr

    Select Case gErr

        Case 102823
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 102824, 102825
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ATENDENTES", gErr)

        Case 102826 'Sem dados

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149210)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

'Incluído por Luiz Nogueira em 27/10/03
Public Function RelacionamentoClientesCons_Le(ByVal objRelacionamentoClientesCons As ClassRelacClientesCons) As Long

Dim lErro As Long
Dim lComando As Long
Dim tRelacionamentoClientesCons As typeRelacionamentoClientesCons
Dim sSelect As String

On Error GoTo Erro_RelacionamentoClientesCons_Le

    'Abre o comando para leitura
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 102867

    'Monta o Select que será executado
    Call RelacionamentoClientesCons_Le1(sSelect, objRelacionamentoClientesCons)

    'Faz o Bind dos campos que serão selecionados
    lErro = RelacionamentoClientesCons_Le2(lComando, sSelect, tRelacionamentoClientesCons)
    If lErro <> SUCESSO Then gError 102868

    'Faz o Bind dos campos que estão sendo passados para a cláusula WHERE
    lErro = RelacionamentoClientesCons_Le3(lComando, tRelacionamentoClientesCons, objRelacionamentoClientesCons)
    If lErro <> SUCESSO Then gError 102869

    'Executa a leitura do BD e guarda os dados lidos no obj
    lErro = RelacionamentoClientesCons_Le4(lComando, tRelacionamentoClientesCons, objRelacionamentoClientesCons)
    If lErro <> SUCESSO And lErro <> 102866 Then gError 102870

    'Se não encontrou => erro. O erro deve ser tratado na rotina chamadora
    If lErro = 102866 Then gError 102871

    Call Comando_Fechar(lComando)

    RelacionamentoClientesCons_Le = SUCESSO

    Exit Function

Erro_RelacionamentoClientesCons_Le:

    RelacionamentoClientesCons_Le = gErr

    Select Case gErr

        Case 102867
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr, Error$)

        Case 102868 To 102870

        Case 102871 'Sem Dados

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149211)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

'Incluído por Luiz Nogueira em 27/10/03
Private Sub RelacionamentoClientesCons_Le1(sSelect As String, ByVal objRelacionamentoClientesCons As ClassRelacClientesCons)

Dim sFrom As String
Dim sWhere As String
Dim sOrderBy As String

    'Monta a parte fixa do select
    sSelect = "SELECT Codigo, FilialEmpresa, Origem, Data, Hora, Tipo, Cliente, FilialCliente, Contato, Atendente, RelacionamentoAnt, Assunto1, Assunto2, Status "
    sFrom = "FROM RelacionamentoClientes R1 "
    sWhere = "WHERE FilialEmpresa = ? AND NOT EXISTS (SELECT * FROM RelacionamentoClientes R2 WHERE R1.FilialEmpresa = R2.FilialEmpresa AND R1.Codigo = R2.RelacionamentoAnt) "
    sOrderBy = "ORDER BY Data, Hora, Atendente, Cliente, FilialCliente"

    'Monta a parte dinâmica da cláusula WHERE
    'Se o código inicial do relacionamento foi informado
    If objRelacionamentoClientesCons.lCodigoDe > 0 Then
    
        'Acrescenta à cláusula WHERE o filtro de Código maior igual ao código inicial informado
        sWhere = sWhere & "AND Codigo>=? "
    
    End If
    
    'Se o código final do relacionamento foi informado
    If objRelacionamentoClientesCons.lCodigoAte > 0 Then
    
        'Acrescenta à cláusula WHERE o filtro de Código menor igual ao código final informado
        sWhere = sWhere & "AND Codigo<=? "
    
    End If
    
    'Se a data inicial está preenchida
    If objRelacionamentoClientesCons.dtDataDe <> DATA_NULA Then

        'Acrescenta à cláusula WHERE o filtro de Data maior igual à Data Inicial informada
        sWhere = sWhere & "AND Data>=? "

    End If

    'Se a data final está preenchida
    If objRelacionamentoClientesCons.dtDataAte <> DATA_NULA Then

        'Acrescenta à cláusula WHERE o filtro de Data menor igual à Data Final informada
        sWhere = sWhere & "AND Data<=? "

    End If

    'Se o cliente está preenchido
    If objRelacionamentoClientesCons.lCliente > 0 Then

        'Acrescenta à cláusula WHERE o filtro de cliente
        sWhere = sWhere & "AND Cliente=? "

    End If

    'Se a filial do cliente está preenchida
    If objRelacionamentoClientesCons.iFilialCliente > 0 Then

        'Acrescenta à cláusula WHERE o filtro de filial de cliente
        sWhere = sWhere & "AND FilialCliente=? "

    End If

    'Se o atendente inicial do relacionamento foi informado
    If objRelacionamentoClientesCons.iAtendenteDe > 0 Then
    
        'Acrescenta à cláusula WHERE o filtro de Atendente maior igual ao atendente inicial informado
        sWhere = sWhere & "AND Atendente>=? "
    
    End If

    'Se o atendente final do relacionamento foi informado
    If objRelacionamentoClientesCons.iAtendenteAte > 0 Then
    
        'Acrescenta à cláusula WHERE o filtro de Atendente menor igual ao atendente final informado
        sWhere = sWhere & "AND Atendente<=? "
    
    End If
    
    'Se o tipo de relacionamento está preenchido
    If objRelacionamentoClientesCons.lTipo > 0 Then

        'Acrescenta à cláusula WHERE o filtro de tipo
        sWhere = sWhere & "AND Tipo=? "

    End If
    
    'Se a origem do relacionamento está preenchida
    If objRelacionamentoClientesCons.iOrigem > 0 Then

        'Acrescenta à cláusula WHERE o filtro de origem
        sWhere = sWhere & "AND Origem=? "

    End If
    
    'Se o status do relacionamento está preenchido
    If objRelacionamentoClientesCons.iStatus <> -1 Then

        'Acrescenta à cláusula WHERE o filtro de tipo
        sWhere = sWhere & "AND Status=? "

    End If
    
    sSelect = sSelect & sFrom & sWhere & sOrderBy

End Sub

'Incluído por Luiz Nogueira em 27/10/03
Private Function RelacionamentoClientesCons_Le2(ByVal lComando As Long, ByVal sSelect As String, tRelacionamentoClientesCons As typeRelacionamentoClientesCons) As Long

Dim iRetorno As Integer

On Error GoTo Erro_RelacionamentoClientesCons_Le2

    With tRelacionamentoClientesCons

        iRetorno = Comando_PrepararInt(lComando, sSelect)
        If (iRetorno <> AD_SQL_SUCESSO) Then gError 102836

        .vlCodigo = CLng(.vlCodigo)
        iRetorno = Comando_BindVarInt(lComando, .vlCodigo)
        If (iRetorno <> AD_SQL_SUCESSO) Then gError 102837

        .viFilialEmpresa = CInt(.viFilialEmpresa)
        iRetorno = Comando_BindVarInt(lComando, .viFilialEmpresa)
        If (iRetorno <> AD_SQL_SUCESSO) Then gError 102838

        .viOrigem = CInt(.viOrigem)
        iRetorno = Comando_BindVarInt(lComando, .viOrigem)
        If (iRetorno <> AD_SQL_SUCESSO) Then gError 102839

        .vdtData = CDate(.vdtData)
        iRetorno = Comando_BindVarInt(lComando, .vdtData)
        If (iRetorno <> AD_SQL_SUCESSO) Then gError 102840

        .vdHora = CDbl(.vdHora)
        iRetorno = Comando_BindVarInt(lComando, .vdHora)
        If (iRetorno <> AD_SQL_SUCESSO) Then gError 102841

        .vlTipo = CLng(.vlTipo)
        iRetorno = Comando_BindVarInt(lComando, .vlTipo)
        If (iRetorno <> AD_SQL_SUCESSO) Then gError 102842

        .vlCliente = CLng(.vlCliente)
        iRetorno = Comando_BindVarInt(lComando, .vlCliente)
        If (iRetorno <> AD_SQL_SUCESSO) Then gError 102843

        .viFilialCliente = CInt(.viFilialCliente)
        iRetorno = Comando_BindVarInt(lComando, .viFilialCliente)
        If (iRetorno <> AD_SQL_SUCESSO) Then gError 102844

        .viContato = CInt(.viContato)
        iRetorno = Comando_BindVarInt(lComando, .viContato)
        If (iRetorno <> AD_SQL_SUCESSO) Then gError 102845

         .viAtendente = CInt(.viAtendente)
        iRetorno = Comando_BindVarInt(lComando, .viAtendente)
        If (iRetorno <> AD_SQL_SUCESSO) Then gError 102846

        .vlRelacionamentoAnt = CLng(.vlRelacionamentoAnt)
        iRetorno = Comando_BindVarInt(lComando, .vlRelacionamentoAnt)
        If (iRetorno <> AD_SQL_SUCESSO) Then gError 102847

        .vsAssunto1 = CStr(.vsAssunto1)
        iRetorno = Comando_BindVarInt(lComando, .vsAssunto1)
        If (iRetorno <> AD_SQL_SUCESSO) Then gError 102848

        .vsAssunto2 = CStr(.vsAssunto2)
        iRetorno = Comando_BindVarInt(lComando, .vsAssunto2)
        If (iRetorno <> AD_SQL_SUCESSO) Then gError 102849

        .viStatus = CInt(.viStatus)
        iRetorno = Comando_BindVarInt(lComando, .viStatus)
        If (iRetorno <> AD_SQL_SUCESSO) Then gError 102850
    
    End With

    RelacionamentoClientesCons_Le2 = SUCESSO

    Exit Function

Erro_RelacionamentoClientesCons_Le2:

    RelacionamentoClientesCons_Le2 = gErr

    Select Case gErr

        Case 102836, 102850
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELACIONAMENTOCLIENTES", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149212)

    End Select

    Exit Function

End Function

'Incluído por Luiz Nogueira em 27/10/03
Private Function RelacionamentoClientesCons_Le3(ByVal lComando As Long, tRelacionamentoClientesCons As typeRelacionamentoClientesCons, ByVal objRelacionamentoClientesCons As ClassRelacClientesCons) As Long

Dim iRetorno As Integer

On Error GoTo Erro_RelacionamentoClientesCons_Le3

    With tRelacionamentoClientesCons

        .viFilialEmpresa = objRelacionamentoClientesCons.iFilialEmpresa
        iRetorno = Comando_BindVarInt(lComando, .viFilialEmpresa)
        If (iRetorno <> AD_SQL_SUCESSO) Then gError 102851

        'Faz o bind da parte dinâmica da cláusula WHERE
        'Se o código inicial do relacionamento foi informado
        If objRelacionamentoClientesCons.lCodigoDe > 0 Then
            .vlCodigoDe = objRelacionamentoClientesCons.lCodigoDe
            iRetorno = Comando_BindVarInt(lComando, .vlCodigoDe)
            If (iRetorno <> AD_SQL_SUCESSO) Then gError 102852
        End If

        'Se o código final do relacionamento foi informado
        If objRelacionamentoClientesCons.lCodigoAte > 0 Then
            .vlCodigoAte = objRelacionamentoClientesCons.lCodigoAte
            iRetorno = Comando_BindVarInt(lComando, .vlCodigoAte)
            If (iRetorno <> AD_SQL_SUCESSO) Then gError 102853
        End If

        'Se a data inicial está preenchida
        If objRelacionamentoClientesCons.dtDataDe <> DATA_NULA Then
            .vdtDataDe = objRelacionamentoClientesCons.dtDataDe
            iRetorno = Comando_BindVarInt(lComando, .vdtDataDe)
            If (iRetorno <> AD_SQL_SUCESSO) Then gError 102854
        End If

        'Se a data final está preenchida
        If objRelacionamentoClientesCons.dtDataAte <> DATA_NULA Then
            .vdtDataAte = objRelacionamentoClientesCons.dtDataAte
            iRetorno = Comando_BindVarInt(lComando, .vdtDataAte)
            If (iRetorno <> AD_SQL_SUCESSO) Then gError 102855
        End If

        'Se o cliente está preenchido
        If objRelacionamentoClientesCons.lCliente > 0 Then
            .vlCliente = objRelacionamentoClientesCons.lCliente
            iRetorno = Comando_BindVarInt(lComando, .vlCliente)
            If (iRetorno <> AD_SQL_SUCESSO) Then gError 102856
        End If

        'Se a filial do cliente está preenchida
        If objRelacionamentoClientesCons.iFilialCliente > 0 Then
            .viFilialCliente = objRelacionamentoClientesCons.iFilialCliente
            iRetorno = Comando_BindVarInt(lComando, .viFilialCliente)
            If (iRetorno <> AD_SQL_SUCESSO) Then gError 102857
        End If

        'Se o atendente inicial do relacionamento foi informado
        If objRelacionamentoClientesCons.iAtendenteDe > 0 Then
            .viAtendenteDe = objRelacionamentoClientesCons.iAtendenteDe
            iRetorno = Comando_BindVarInt(lComando, .viAtendenteDe)
            If (iRetorno <> AD_SQL_SUCESSO) Then gError 102858
        End If

        'Se o atendente final do relacionamento foi informado
        If objRelacionamentoClientesCons.iAtendenteAte > 0 Then
            .viAtendenteAte = objRelacionamentoClientesCons.iAtendenteAte
            iRetorno = Comando_BindVarInt(lComando, .viAtendenteAte)
            If (iRetorno <> AD_SQL_SUCESSO) Then gError 102859
        End If

        'Se o tipo de relacionamento está preenchido
        If objRelacionamentoClientesCons.lTipo > 0 Then
            .vlTipo = objRelacionamentoClientesCons.lTipo
            iRetorno = Comando_BindVarInt(lComando, .vlTipo)
            If (iRetorno <> AD_SQL_SUCESSO) Then gError 102860
        End If

        'Se a origem do relacionamento está preenchida
        If objRelacionamentoClientesCons.iOrigem > 0 Then
            .viOrigem = objRelacionamentoClientesCons.iOrigem
            iRetorno = Comando_BindVarInt(lComando, .viOrigem)
            If (iRetorno <> AD_SQL_SUCESSO) Then gError 102861
        End If

        'Se o status do relacionamento está preenchido
        If objRelacionamentoClientesCons.iStatus <> -1 Then
            .viStatus = objRelacionamentoClientesCons.iStatus
            iRetorno = Comando_BindVarInt(lComando, .viStatus)
            If (iRetorno <> AD_SQL_SUCESSO) Then gError 102862
        End If
        
    End With

    RelacionamentoClientesCons_Le3 = SUCESSO

    Exit Function

Erro_RelacionamentoClientesCons_Le3:

    RelacionamentoClientesCons_Le3 = gErr

    Select Case gErr

        Case 102850 To 102862
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELACIONAMENTOCLIENTES", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149213)

    End Select

    Exit Function

End Function

'Incluído por Luiz Nogueira em 27/10/03
Private Function RelacionamentoClientesCons_Le4(ByVal lComando As Long, tRelacionamentoClientesCons As typeRelacionamentoClientesCons, ByVal objRelacionamentoClientesCons As ClassRelacClientesCons) As Long

Dim iRetorno As Integer
Dim objRelacionamentoClientes As ClassRelacClientes
Dim lErro As Long

On Error GoTo Erro_RelacionamentoClientesCons_Le4

    'Inicializa as strings para leitura de dados
    tRelacionamentoClientesCons.vsAssunto1 = String(STRING_BUFFER_MAX_TEXTO, 0)
    tRelacionamentoClientesCons.vsAssunto2 = String(STRING_BUFFER_MAX_TEXTO, 0)

    'Executa a leitura
    iRetorno = Comando_ExecutarInt(lComando)
    If (iRetorno <> AD_SQL_SUCESSO) Then gError 102863

    'Busca o primeiro registro encontrado
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 102864

    'Se não encontrou nenhum registro => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 102866

    'Enquanto houverem registros
    Do While lErro = SUCESSO

        'Instancia um novo obj
        Set objRelacionamentoClientes = New ClassRelacClientes

        'Transfere os dados lidos para o obj
        With objRelacionamentoClientes
            .lCodigo = tRelacionamentoClientesCons.vlCodigo
            .iFilialEmpresa = tRelacionamentoClientesCons.viFilialEmpresa
            .iOrigem = tRelacionamentoClientesCons.viOrigem
            .dtData = tRelacionamentoClientesCons.vdtData
            .dtHora = CDate(tRelacionamentoClientesCons.vdHora)
            .lTipo = tRelacionamentoClientesCons.vlTipo
            .lCliente = tRelacionamentoClientesCons.vlCliente
            .iFilialCliente = tRelacionamentoClientesCons.viFilialCliente
            .iContato = tRelacionamentoClientesCons.viContato
            .iAtendente = tRelacionamentoClientesCons.viAtendente
            .lRelacionamentoAnt = tRelacionamentoClientesCons.vlRelacionamentoAnt
            .sAssunto1 = tRelacionamentoClientesCons.vsAssunto1
            .sAssunto2 = tRelacionamentoClientesCons.vsAssunto2
            .iStatus = tRelacionamentoClientesCons.viStatus
        End With
        
        lErro = RelacCliTexto_Le(objRelacionamentoClientes)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        'Adiciona o registro à coleção
        objRelacionamentoClientesCons.colRelacionamentoClientes.Add objRelacionamentoClientes

        'Busca o próximo registro encontrado
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 102865

    Loop

    RelacionamentoClientesCons_Le4 = SUCESSO

    Exit Function

Erro_RelacionamentoClientesCons_Le4:

    RelacionamentoClientesCons_Le4 = gErr

    Select Case gErr

        Case 102863 To 102865
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELACIONAMENTOCLIENTES", gErr)

        Case 102866
        
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149214)

    End Select

    Exit Function

End Function

'Incluído por Luiz Nogueira em 27/10/03
Public Function Atendentes_LockShared(ByVal iCodigo As Integer, Optional ByVal iFilialEmpresa As Integer) As Long
'Faz lock shared no atendente passado como parâmetro
'iCodigo recebe o código do atendente a ser "lockado"

Dim lErro As Long
Dim lComando As Long
Dim iIndice As Integer
Dim iCodAux As Long

On Error GoTo Erro_Atendentes_LockShared

    'Executa a abertura dos Comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 102980

    'Se não indicou a filial do atendente => usa a filial empresa ativa
    If iFilialEmpresa = 0 Then iFilialEmpresa = giFilialEmpresa
    
    'Busca o relacionamento a ser lockado
    lErro = Comando_ExecutarLockado(lComando, "SELECT Codigo FROM Atendentes WHERE Codigo=? AND FilialEmpresa=?", iCodAux, iCodigo, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 102981
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 102982
    
    'Se não encontrou o relacionamento => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 102983
    
    'Efetua o lock do relacionamento
    lErro = Comando_LockShared(lComando)
    If lErro <> AD_SQL_SUCESSO Then gError 102984
    
    'Executa o fechamento dos Comandos
    Call Comando_Fechar(lComando)

    Atendentes_LockShared = SUCESSO

    Exit Function

Erro_Atendentes_LockShared:

    Atendentes_LockShared = gErr

    Select Case gErr

        Case 102980
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 102981, 102982
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ATENDENTES", gErr)

        Case 102983

        Case 102984
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_ATENDENTES", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149215)

    End Select

    Call Comando_Fechar(lComando)

End Function

Private Function Operador_Le_Vendedor(ByVal objVendedor As ClassVendedor, iOperador As Integer) As Long
'Verifica se existe operador relacionada ao Vendedor passado como parametro

Dim lErro As Long
Dim lComando As Long

On Error GoTo Erro_Operador_Le_Vendedor

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 126462

    'pesquisa comissoes no BD
    lErro = Comando_Executar(lComando, "SELECT Codigo FROM Operador WHERE CodVendedor = ? ", iOperador, objVendedor.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 126463

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 126464

    If lErro = AD_SQL_SEM_DADOS Then gError 126465

    Call Comando_Fechar(lComando)

    Operador_Le_Vendedor = SUCESSO

    Exit Function

Erro_Operador_Le_Vendedor:

    Operador_Le_Vendedor = gErr

    Select Case gErr

        Case 126462
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 126463, 126464
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_OPERADOR3", gErr, objVendedor.iCodigo)

        Case 126465

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149216)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Cliente_Le_NomeReduzido_Parcial(sNomeReduzidoParte As String, sNomeReduzido As String) As Long
'Le o NomeReduzido do cliente a partir de parte do nome reduzido. Se a parte não individualiza o cliente ==> retorna codigo de erro

Dim lComando As Long
Dim lErro As Long
Dim iTamanho As Integer
Dim sNomeReduzidoAux As String

On Error GoTo Erro_Cliente_Le_NomeReduzido_Parcial

    sNomeReduzido = String(STRING_CLIENTE_NOME_REDUZIDO, 0)

    iTamanho = Len(sNomeReduzidoParte)

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 71645

    'Le o NomeReduzido da tabela Clientes
    lErro = Comando_Executar(lComando, "SELECT NomeReduzido FROM Clientes WHERE NomeReduzido >= ?", sNomeReduzido, sNomeReduzidoParte)
    If lErro <> AD_SQL_SUCESSO Then gError 71646

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 71647

    'se não encontrou nenhum cliente com este prefixo ==> erro
    If lErro = AD_SQL_SEM_DADOS Then gError 71648

    'se encontrou um cliente, mas o prefixo difere ==> erro
    If UCase(sNomeReduzidoParte) <> UCase(left(sNomeReduzido, iTamanho)) Then gError 71649

    sNomeReduzidoAux = sNomeReduzido
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 71650

    If lErro = AD_SQL_SUCESSO Then
        
        'se encontrou um segundo registro com o mesmo prefixo ==> ainda não pode trazer o nome pois ainda não conseguiu individualizar
        If UCase(sNomeReduzidoParte) = UCase(left(sNomeReduzido, iTamanho)) Then gError 71651
        
    End If
    
    sNomeReduzido = sNomeReduzidoAux

    lErro = Comando_Fechar(lComando)

    Cliente_Le_NomeReduzido_Parcial = SUCESSO

    Exit Function

Erro_Cliente_Le_NomeReduzido_Parcial:

    Cliente_Le_NomeReduzido_Parcial = gErr

    Select Case gErr

        Case 71645
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 71646, 71647, 71650
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLIENTES", gErr)

        Case 71648, 71649, 71651

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149217)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Cliente_Pesquisa_NomeReduzido(objCliente As Object, sNomeReduzidoParte As String) As Long
'rotina para trazer o nome do cliente com uma parte dos caracteres digitados

Dim iCursorPos As Integer
Dim sNomeReduzido As String
Dim lErro As Long

On Error GoTo Erro_Cliente_Pesquisa_NomeReduzido

        If sNomeReduzidoParte <> objCliente.Text And Len(sNomeReduzidoParte) <= Len(objCliente.Text) And Not IsNumeric(objCliente.Text) Then
            
            sNomeReduzidoParte = objCliente.Text
            
            lErro = Cliente_Le_NomeReduzido_Parcial(sNomeReduzidoParte, sNomeReduzido)
            If lErro <> SUCESSO Then gError 71652
            
            If sNomeReduzido <> objCliente.Text Then
            
                sNomeReduzidoParte = sNomeReduzido
                iCursorPos = objCliente.SelStart
                objCliente.Text = sNomeReduzido
                objCliente.SelStart = iCursorPos
                objCliente.SelLength = Len(objCliente.Text) - (iCursorPos)
        
            End If
        
        Else
        
            sNomeReduzidoParte = objCliente.Text
        
        End If

    Cliente_Pesquisa_NomeReduzido = SUCESSO

    Exit Function

Erro_Cliente_Pesquisa_NomeReduzido:

    Cliente_Pesquisa_NomeReduzido = gErr

    Select Case gErr

        Case 71652

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149218)

    End Select

End Function

Function TiposDocInfo_Le_Sigla(ByVal objTipoDocInfo As ClassTipoDocInfo) As Long

Dim lErro As Long
Dim lComando As Long
Dim iCodigo As Integer
Dim sNomeTelaNFiscal As String

On Error GoTo Erro_TiposDocInfo_Le_Sigla

    'Abrir comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 133358

    sNomeTelaNFiscal = String(STRING_NOME, 0)

    'Lê a Nota Fiscal Original
    lErro = Comando_Executar(lComando, "SELECT NomeTelaNFiscal FROM TiposDocInfo WHERE Sigla =?", sNomeTelaNFiscal, objTipoDocInfo.sSigla)
    If lErro <> AD_SQL_SUCESSO Then gError 133359

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 133360

    'Se não encontrar --> Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 133361

    'Retorna o Número Interno da Nota Fiscal
    objTipoDocInfo.sNomeTelaNFiscal = sNomeTelaNFiscal

    'Fecha o comando
    Call Comando_Fechar(lComando)

    TiposDocInfo_Le_Sigla = SUCESSO

    Exit Function

Erro_TiposDocInfo_Le_Sigla:

    TiposDocInfo_Le_Sigla = gErr

    Select Case gErr

        Case 133358
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 133359, 133360
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPODOCINFO1", gErr)

        Case 133361

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149219)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function ItensNFGrade_Le1(lNumIntNF As Long, ByVal lNumIntGrade As Long) As Long

Dim lComando As Long
Dim lErro As Long

On Error GoTo Erro_ItensNFGrade_Le1

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 133851
        
    lErro = Comando_Executar(lComando, "SELECT NumIntNF FROM ItensNFiscalGrade, ItensNFiscal WHERE ItensNFiscalGrade.NumIntDoc =? AND ItensNFiscalGrade.NumIntItemNF = ItensNFiscal.NumIntDoc", _
    lNumIntNF, lNumIntGrade)
    If lErro <> AD_SQL_SUCESSO Then gError 133852
        
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 133853
        
    If lErro = AD_SQL_SEM_DADOS Then gError 133854

    Call Comando_Fechar(lComando)
    
    ItensNFGrade_Le1 = SUCESSO
    
    Exit Function

Erro_ItensNFGrade_Le1:
    
    ItensNFGrade_Le1 = gErr
        
    Select Case gErr
    
        Case 133851
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 133852, 133853
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSNFGRADE", gErr)
        
        Case 133854
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149220)
            
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function TituloReceberBaixado_Le_SemNumIntDoc(ByVal objTituloReceber As ClassTituloReceber) As Long
'Lê dados do Título a partir de FilialEmpresa, Cliente, Filial, SiglaDocumento, NumTitulo

Dim lErro As Long
Dim lComando As Long
Dim tTituloReceber As typeTituloReceber

On Error GoTo Erro_TituloReceberBaixado_Le_SemNumIntDoc

    'Inicializa comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 28571

    tTituloReceber.sObservacao = String(STRING_TITULO_OBSERVACAO, 0)
    tTituloReceber.sCcl = String(STRING_CCL, 0)
    tTituloReceber.sNatureza = String(STRING_NATMOVCTA_CODIGO, 0)

    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, DataEmissao, Status, Saldo, NumParcelas, Valor, ValorIRRF, ValorISS, ISSRetido, ValorINSS, INSSRetido, PercJurosDiarios, PercMulta, Observacao, CondicaoPagto, DataRegistro, PISRetido, COFINSRetido, CSLLRetido, Moeda, Ccl, Natureza, CpoGenerico1, ReajustePeriodicidade, ReajusteBase, ReajustadoAte FROM TitulosRecBaixados WHERE FilialEmpresa = ? AND Cliente = ? AND Filial = ? AND SiglaDocumento = ? AND NumTitulo = ? AND Status <> ?", _
        tTituloReceber.lNumIntDoc, tTituloReceber.dtDataEmissao, tTituloReceber.iStatus, tTituloReceber.dSaldo, tTituloReceber.iNumParcelas, tTituloReceber.dValor, tTituloReceber.dValorIRRF, tTituloReceber.dValorISS, tTituloReceber.dISSRetido, tTituloReceber.dValorINSS, tTituloReceber.iINSSRetido, tTituloReceber.dPercJurosDiario, tTituloReceber.dPercMulta, tTituloReceber.sObservacao, tTituloReceber.iCondicaoPagto, tTituloReceber.dtDataRegistro, tTituloReceber.dPISRetido, tTituloReceber.dCOFINSRetido, tTituloReceber.dCSLLRetido, tTituloReceber.iMoeda, tTituloReceber.sCcl, tTituloReceber.sNatureza, tTituloReceber.lCpoGenerico1, tTituloReceber.iReajustePeriodicidade, tTituloReceber.dtReajusteBase, tTituloReceber.dtReajustadoAte, objTituloReceber.iFilialEmpresa, objTituloReceber.lCliente, objTituloReceber.iFilial, objTituloReceber.sSiglaDocumento, objTituloReceber.lNumTitulo, STATUS_EXCLUIDO)
    If lErro <> AD_SQL_SUCESSO Then Error 28572

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 28573

    If lErro = AD_SQL_SEM_DADOS Then Error 28574

    tTituloReceber.iFilialEmpresa = objTituloReceber.iFilialEmpresa
    tTituloReceber.lNumTitulo = objTituloReceber.lNumTitulo
    tTituloReceber.sSiglaDocumento = objTituloReceber.sSiglaDocumento
    tTituloReceber.lCliente = objTituloReceber.lCliente
    tTituloReceber.iFilial = objTituloReceber.iFilial
    tTituloReceber.dtDataEmissao = objTituloReceber.dtDataEmissao

    Call TituloReceber_MoverCamposObj(objTituloReceber, tTituloReceber)
    
    'Libera comando
    Call Comando_Fechar(lComando)

    TituloReceberBaixado_Le_SemNumIntDoc = SUCESSO

    Exit Function

Erro_TituloReceberBaixado_Le_SemNumIntDoc:

    TituloReceberBaixado_Le_SemNumIntDoc = Err

    Select Case Err

        Case 28571
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 28572, 28573
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOSREC2", Err, objTituloReceber.iFilialEmpresa, objTituloReceber.lCliente, objTituloReceber.iFilial, objTituloReceber.sSiglaDocumento, objTituloReceber.lNumTitulo)

        Case 28574
            'Não encontrou. Erro tratado na rotina chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149079)

    End Select

    'Libera comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'##############################################################
'Inserido por Wagner 06/07/2006
Function Contato_Le(ByVal objContato As ClassContatos) As Long
'le o Contato ( e sua matriz ) a partir da chave(lCodigo) fornecida em objContato. Devolve os dados em objContato

Dim lComando As Long
Dim lComando2 As Long
Dim lErro As Long
Dim tContato As typeContato
Dim lCodigoContato As Long

On Error GoTo Erro_Contato_Le

    tContato.sRazaoSocial = String(STRING_CLIENTE_RAZAO_SOCIAL, 0)
    tContato.sNomeReduzido = String(STRING_CLIENTE_NOME_REDUZIDO, 0)
    tContato.sObservacao = String(STRING_CLIENTE_OBSERVACAO, 0)
    tContato.sNome = String(STRING_FILIAL_CLIENTE_NOME, 0)
    tContato.sCgc = String(STRING_CGC, 0)
    tContato.sRG = String(STRING_RG, 0)
    tContato.sInscricaoEstadual = String(STRING_INSCR_EST, 0)
    tContato.sInscricaoMunicipal = String(STRING_INSCR_MUN, 0)
    tContato.sInscricaoSuframa = String(STRING_INSCR_SUF, 0)
    tContato.sObservacao2 = String(STRING_CLIENTE_OBSERVACAO, 0)
    tContato.sContaContabil = String(STRING_CONTA, 0)
    tContato.sGuia = String(STRING_CLIENTE_GUIA, 0)
    
    'guarda o código correto do Contato para msg de erro
    If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL And objContato.lCodigo = 0 Then
        lCodigoContato = objContato.lCodigoLoja
    Else
        lCodigoContato = objContato.lCodigo
    End If

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 180544

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then gError 180545

    'Pega os dados da tabela Contatos
    If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL And objContato.lCodigo = 0 Then
        'faço o select com codigoloja
        lErro = Comando_Executar(lComando, "SELECT Codigo, RazaoSocial, NomeReduzido, Tipo, Observacao, LimiteCredito, CondicaoPagto, Desconto, CodMensagem, TabelaPreco,ProxCodFilial, CodPadraoCobranca, Ativo, Transferido, Bloqueado FROM Contatos WHERE FilialEmpresaLoja = ? AND CodigoLoja = ? ", _
        tContato.lCodigo, tContato.sRazaoSocial, tContato.sNomeReduzido, tContato.iTipo, tContato.sObservacao, tContato.dLimiteCredito, tContato.iCondicaoPagto, tContato.dDesconto, tContato.iCodMensagem, tContato.iTabelaPreco, tContato.iProxCodFilial, tContato.iCodPadraoCobranca, tContato.iAtivo, tContato.iTransferido, tContato.iBloqueado, objContato.iFilialEmpresaLoja, objContato.lCodigoLoja)
    Else
        'faço o select com codigo
        lErro = Comando_Executar(lComando, "SELECT Codigo, RazaoSocial, NomeReduzido, Tipo, Observacao, LimiteCredito, CondicaoPagto, Desconto, CodMensagem, TabelaPreco,ProxCodFilial, CodPadraoCobranca, Ativo, Transferido, Bloqueado FROM Contatos WHERE Codigo = ?", _
        tContato.lCodigo, tContato.sRazaoSocial, tContato.sNomeReduzido, tContato.iTipo, tContato.sObservacao, tContato.dLimiteCredito, tContato.iCondicaoPagto, tContato.dDesconto, tContato.iCodMensagem, tContato.iTabelaPreco, tContato.iProxCodFilial, tContato.iCodPadraoCobranca, tContato.iAtivo, tContato.iTransferido, tContato.iBloqueado, objContato.lCodigo)
    End If
    
    If lErro <> AD_SQL_SUCESSO Then gError 180546

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 180547

    If lErro = AD_SQL_SEM_DADOS Then gError 180548

    objContato.lCodigo = tContato.lCodigo
    objContato.sRazaoSocial = tContato.sRazaoSocial
    objContato.sNomeReduzido = tContato.sNomeReduzido
    objContato.iTipo = tContato.iTipo
    objContato.sObservacao = tContato.sObservacao
    objContato.dLimiteCredito = tContato.dLimiteCredito
    objContato.iCondicaoPagto = tContato.iCondicaoPagto
    objContato.dDesconto = tContato.dDesconto
    objContato.iCodMensagem = tContato.iCodMensagem
    objContato.iTabelaPreco = tContato.iTabelaPreco
    objContato.iProxCodFilial = tContato.iProxCodFilial
    objContato.iCodPadraoCobranca = tContato.iCodPadraoCobranca
    objContato.iAtivo = tContato.iAtivo
    objContato.iTransferido = tContato.iTransferido
    objContato.iBloqueado = tContato.iBloqueado
    
    'Pega os dados da tabela Filiais Contatos
    If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL And objContato.lCodigo = 0 Then
        lErro = Comando_Executar(lComando2, "SELECT CGC, InscricaoEstadual, InscricaoMunicipal,InscricaoSuframa, Endereco, EnderecoEntrega, EnderecoCobranca, CodTransportadora, Observacao, ContaContabil, Vendedor, ComissaoVendas, Regiao, FreqVisitas, DataUltVisita, CodCobrador, TipoFrete, RG, CodTranspRedesp, DetPagFrete, Guia FROM FiliaisContatos WHERE FilialEmpresaLoja = ? AND CodContatoLoja = ? AND CodFilialLoja = ?", _
        tContato.sCgc, tContato.sInscricaoEstadual, tContato.sInscricaoMunicipal, tContato.sInscricaoSuframa, tContato.lEndereco, tContato.lEnderecoEntrega, tContato.lEnderecoCobranca, tContato.iCodTransportadora, tContato.sObservacao2, tContato.sContaContabil, tContato.iVendedor, tContato.dComissaoVendas, tContato.iRegiao, tContato.iFreqVisitas, tContato.dtDataUltVisita, tContato.iCodCobrador, tContato.iTipoFrete, tContato.sRG, tContato.iCodTranspRedesp, tContato.iDetPagFrete, tContato.sGuia, objContato.iFilialEmpresaLoja, objContato.lCodigoLoja, FILIAL_MATRIZ)
    Else
        lErro = Comando_Executar(lComando2, "SELECT CGC, InscricaoEstadual, InscricaoMunicipal,InscricaoSuframa, Endereco, EnderecoEntrega, EnderecoCobranca, CodTransportadora, Observacao, ContaContabil, Vendedor, ComissaoVendas, Regiao, FreqVisitas, DataUltVisita, CodCobrador, TipoFrete, RG, CodTranspRedesp, DetPagFrete, Guia FROM FiliaisContatos WHERE CodContato = ? AND CodFilial = ?", _
        tContato.sCgc, tContato.sInscricaoEstadual, tContato.sInscricaoMunicipal, tContato.sInscricaoSuframa, tContato.lEndereco, tContato.lEnderecoEntrega, tContato.lEnderecoCobranca, tContato.iCodTransportadora, tContato.sObservacao2, tContato.sContaContabil, tContato.iVendedor, tContato.dComissaoVendas, tContato.iRegiao, tContato.iFreqVisitas, tContato.dtDataUltVisita, tContato.iCodCobrador, tContato.iTipoFrete, tContato.sRG, tContato.iCodTranspRedesp, tContato.iDetPagFrete, tContato.sGuia, objContato.lCodigo, FILIAL_MATRIZ)
    End If

    If lErro <> AD_SQL_SUCESSO Then gError 180549

    lErro = Comando_BuscarPrimeiro(lComando2)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 180550

    If lErro = AD_SQL_SEM_DADOS Then gError 180551
    
    objContato.sRG = tContato.sRG
    objContato.sCgc = tContato.sCgc
    objContato.sInscricaoEstadual = tContato.sInscricaoEstadual
    objContato.sInscricaoMunicipal = tContato.sInscricaoMunicipal
    objContato.sInscricaoSuframa = tContato.sInscricaoSuframa
    objContato.lEndereco = tContato.lEndereco
    objContato.lEnderecoEntrega = tContato.lEnderecoEntrega
    objContato.lEnderecoCobranca = tContato.lEnderecoCobranca
    objContato.iCodTransportadora = tContato.iCodTransportadora
    objContato.sObservacao2 = tContato.sObservacao2
    objContato.sContaContabil = tContato.sContaContabil
    objContato.iVendedor = tContato.iVendedor
    objContato.dComissaoVendas = tContato.dComissaoVendas
    objContato.iRegiao = tContato.iRegiao
    objContato.iFreqVisitas = tContato.iFreqVisitas
    objContato.dtDataUltVisita = tContato.dtDataUltVisita
    objContato.iCodCobrador = tContato.iCodCobrador
    objContato.iTipoFrete = tContato.iTipoFrete
    objContato.iCodTranspRedesp = tContato.iCodTranspRedesp
    objContato.iDetPagFrete = tContato.iDetPagFrete
    objContato.sGuia = tContato.sGuia
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)

    Contato_Le = SUCESSO

Exit Function

Erro_Contato_Le:

    Contato_Le = gErr

    Select Case gErr

        Case 180544, 180545
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 180546, 180547
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTATOS", gErr)

        Case 180548 'Contato nao cadastrado em Contatos

        Case 180549, 180550
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISCONTATOS", gErr)

        Case 180551
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTATO_SEM_FILIAL", gErr, lCodigoContato)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180552)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)

    Exit Function

End Function

Function Contato_Le_NomeReduzido(ByVal objContato As ClassContatos) As Long
'le o Contato ( e sua matriz ) a partir da nomereduzido fornecida em objContato. Devolve os dados em objContato

Dim lComando As Long
Dim lComando2 As Long
Dim lErro As Long
Dim tContato As typeContato
Dim lCodigoContato As Long

On Error GoTo Erro_Contato_Le_NomeReduzido

    tContato.sRazaoSocial = String(STRING_CLIENTE_RAZAO_SOCIAL, 0)
    tContato.sNomeReduzido = String(STRING_CLIENTE_NOME_REDUZIDO, 0)
    tContato.sObservacao = String(STRING_CLIENTE_OBSERVACAO, 0)
    tContato.sNome = String(STRING_FILIAL_CLIENTE_NOME, 0)
    tContato.sCgc = String(STRING_CGC, 0)
    tContato.sRG = String(STRING_RG, 0)
    tContato.sInscricaoEstadual = String(STRING_INSCR_EST, 0)
    tContato.sInscricaoMunicipal = String(STRING_INSCR_MUN, 0)
    tContato.sInscricaoSuframa = String(STRING_INSCR_SUF, 0)
    tContato.sObservacao2 = String(STRING_CLIENTE_OBSERVACAO, 0)
    tContato.sContaContabil = String(STRING_CONTA, 0)
    tContato.sGuia = String(STRING_CLIENTE_GUIA, 0)
    
    'guarda o código correto do Contato para msg de erro
    If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL And objContato.lCodigo = 0 Then
        lCodigoContato = objContato.lCodigoLoja
    Else
        lCodigoContato = objContato.lCodigo
    End If

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 180741

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then gError 180742

    lErro = Comando_Executar(lComando, "SELECT Codigo, RazaoSocial, NomeReduzido, Tipo, Observacao, LimiteCredito, CondicaoPagto, Desconto, CodMensagem, TabelaPreco,ProxCodFilial, CodPadraoCobranca, Ativo, Transferido, Bloqueado FROM Contatos WHERE NomeReduzido = ? ", _
    tContato.lCodigo, tContato.sRazaoSocial, tContato.sNomeReduzido, tContato.iTipo, tContato.sObservacao, tContato.dLimiteCredito, tContato.iCondicaoPagto, tContato.dDesconto, tContato.iCodMensagem, tContato.iTabelaPreco, tContato.iProxCodFilial, tContato.iCodPadraoCobranca, tContato.iAtivo, tContato.iTransferido, tContato.iBloqueado, objContato.sNomeReduzido)
    
    If lErro <> AD_SQL_SUCESSO Then gError 180743

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 180744

    If lErro = AD_SQL_SEM_DADOS Then gError 180745

    objContato.lCodigo = tContato.lCodigo
    objContato.sRazaoSocial = tContato.sRazaoSocial
    objContato.sNomeReduzido = tContato.sNomeReduzido
    objContato.iTipo = tContato.iTipo
    objContato.sObservacao = tContato.sObservacao
    objContato.dLimiteCredito = tContato.dLimiteCredito
    objContato.iCondicaoPagto = tContato.iCondicaoPagto
    objContato.dDesconto = tContato.dDesconto
    objContato.iCodMensagem = tContato.iCodMensagem
    objContato.iTabelaPreco = tContato.iTabelaPreco
    objContato.iProxCodFilial = tContato.iProxCodFilial
    objContato.iCodPadraoCobranca = tContato.iCodPadraoCobranca
    objContato.iAtivo = tContato.iAtivo
    objContato.iTransferido = tContato.iTransferido
    objContato.iBloqueado = tContato.iBloqueado
    
    'Pega os dados da tabela Filiais Contatos
    If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL And objContato.lCodigo = 0 Then
        lErro = Comando_Executar(lComando2, "SELECT CGC, InscricaoEstadual, InscricaoMunicipal,InscricaoSuframa, Endereco, EnderecoEntrega, EnderecoCobranca, CodTransportadora, Observacao, ContaContabil, Vendedor, ComissaoVendas, Regiao, FreqVisitas, DataUltVisita, CodCobrador, TipoFrete, RG, CodTranspRedesp, DetPagFrete, Guia FROM FiliaisContatos WHERE FilialEmpresaLoja = ? AND CodContatoLoja = ? AND CodFilialLoja = ?", _
        tContato.sCgc, tContato.sInscricaoEstadual, tContato.sInscricaoMunicipal, tContato.sInscricaoSuframa, tContato.lEndereco, tContato.lEnderecoEntrega, tContato.lEnderecoCobranca, tContato.iCodTransportadora, tContato.sObservacao2, tContato.sContaContabil, tContato.iVendedor, tContato.dComissaoVendas, tContato.iRegiao, tContato.iFreqVisitas, tContato.dtDataUltVisita, tContato.iCodCobrador, tContato.iTipoFrete, tContato.sRG, tContato.iCodTranspRedesp, tContato.iDetPagFrete, tContato.sGuia, tContato.iFilialEmpresaLoja, tContato.lCodigoLoja, FILIAL_MATRIZ)
    Else
        lErro = Comando_Executar(lComando2, "SELECT CGC, InscricaoEstadual, InscricaoMunicipal,InscricaoSuframa, Endereco, EnderecoEntrega, EnderecoCobranca, CodTransportadora, Observacao, ContaContabil, Vendedor, ComissaoVendas, Regiao, FreqVisitas, DataUltVisita, CodCobrador, TipoFrete, RG, CodTranspRedesp, DetPagFrete, Guia FROM FiliaisContatos WHERE CodContato = ? AND CodFilial = ?", _
        tContato.sCgc, tContato.sInscricaoEstadual, tContato.sInscricaoMunicipal, tContato.sInscricaoSuframa, tContato.lEndereco, tContato.lEnderecoEntrega, tContato.lEnderecoCobranca, tContato.iCodTransportadora, tContato.sObservacao2, tContato.sContaContabil, tContato.iVendedor, tContato.dComissaoVendas, tContato.iRegiao, tContato.iFreqVisitas, tContato.dtDataUltVisita, tContato.iCodCobrador, tContato.iTipoFrete, tContato.sRG, tContato.iCodTranspRedesp, tContato.iDetPagFrete, tContato.sGuia, tContato.lCodigo, FILIAL_MATRIZ)
    End If

    If lErro <> AD_SQL_SUCESSO Then gError 180746

    lErro = Comando_BuscarPrimeiro(lComando2)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 180747

    If lErro = AD_SQL_SEM_DADOS Then gError 180748
    
    objContato.sRG = tContato.sRG
    objContato.sCgc = tContato.sCgc
    objContato.sInscricaoEstadual = tContato.sInscricaoEstadual
    objContato.sInscricaoMunicipal = tContato.sInscricaoMunicipal
    objContato.sInscricaoSuframa = tContato.sInscricaoSuframa
    objContato.lEndereco = tContato.lEndereco
    objContato.lEnderecoEntrega = tContato.lEnderecoEntrega
    objContato.lEnderecoCobranca = tContato.lEnderecoCobranca
    objContato.iCodTransportadora = tContato.iCodTransportadora
    objContato.sObservacao2 = tContato.sObservacao2
    objContato.sContaContabil = tContato.sContaContabil
    objContato.iVendedor = tContato.iVendedor
    objContato.dComissaoVendas = tContato.dComissaoVendas
    objContato.iRegiao = tContato.iRegiao
    objContato.iFreqVisitas = tContato.iFreqVisitas
    objContato.dtDataUltVisita = tContato.dtDataUltVisita
    objContato.iCodCobrador = tContato.iCodCobrador
    objContato.iTipoFrete = tContato.iTipoFrete
    objContato.iCodTranspRedesp = tContato.iCodTranspRedesp
    objContato.iDetPagFrete = tContato.iDetPagFrete
    objContato.sGuia = tContato.sGuia
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)

    Contato_Le_NomeReduzido = SUCESSO

Exit Function

Erro_Contato_Le_NomeReduzido:

    Contato_Le_NomeReduzido = gErr

    Select Case gErr

        Case 180741, 180742
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 180743, 180744
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTATOS", gErr)

        Case 180745 'Contato nao cadastrado em Contatos

        Case 180746, 180747
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISCONTATOS", gErr)

        Case 180748
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTATO_SEM_FILIAL", gErr, lCodigoContato)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180749)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)

    Exit Function

End Function

Function FilialContato_Le_Categorias(ByVal objFilialContato As ClassFilialContato, ByVal colCategoriaItem As Collection) As Long
'Lê na tabela de FilialContatoCategorias, todas as categorias da Filial de Contato (com o respectivo valor em cada categoria), guardando na coleção colCategoriaItem

Dim lErro As Long
Dim iIndice As Integer
Dim lComando As Long
Dim sCategoria As String
Dim sItem As String
Dim objFilialConCategoria As ClassFilialConCategoria

On Error GoTo Erro_FilialContato_Le_Categorias

    'Inicializar comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 180762

    sCategoria = String(STRING_CATEGORIACLIENTE_CATEGORIA, 0)
    sItem = String(STRING_CATEGORIACLIENTEITEM_ITEM, 0)

    'Pesquisa no BD as categoria da filial do Contato
    lErro = Comando_Executar(lComando, "SELECT Categoria, Item FROM FilialContatoCategorias WHERE Contato = ? AND Filial = ?", sCategoria, sItem, objFilialContato.lCodContato, objFilialContato.iCodFilial)
    If lErro <> AD_SQL_SUCESSO Then gError 180763

    'Verifica se alguma categoria existe
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 180764

    'Passa para a collection
    Do While lErro <> AD_SQL_SEM_DADOS

        Set objFilialConCategoria = New ClassFilialConCategoria

        objFilialConCategoria.lContato = objFilialContato.lCodContato
        objFilialConCategoria.iFilial = objFilialContato.iCodFilial
        objFilialConCategoria.sCategoria = sCategoria
        objFilialConCategoria.sItem = sItem

        colCategoriaItem.Add objFilialConCategoria

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 180765

    Loop

    'Fechamento comando
    Call Comando_Fechar(lComando)

    FilialContato_Le_Categorias = SUCESSO

    Exit Function

Erro_FilialContato_Le_Categorias:

    FilialContato_Le_Categorias = gErr

    Select Case gErr

        Case 180762
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 180763, 180764, 180765
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIALContatoCATEGORIAS2", gErr, objFilialConCategoria.lContato)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 180766)

    End Select

    'Fechamento comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function FilialContato_Le(ByVal objFilialContato As ClassFilialContato) As Long
'Le todos os dados da tabela FiliaisContatos, correspondente a chave passada por parametro

Dim lComando As Long
Dim lErro As Long
Dim tFilialContato As typeFilialContato
Dim iAtivo As Integer
Dim lCodContato As Long
Dim iCodFilial As Integer

On Error GoTo Erro_FilialContato_Le

    tFilialContato.sNome = String(STRING_FILIAL_CLIENTE_NOME, 0)
    tFilialContato.sRG = String(STRING_RG, 0)
    tFilialContato.sCgc = String(STRING_CGC, 0)
    tFilialContato.sInscricaoEstadual = String(STRING_INSCR_EST, 0)
    tFilialContato.sInscricaoMunicipal = String(STRING_INSCR_MUN, 0)
    tFilialContato.sInscricaoSuframa = String(STRING_INSCR_SUF, 0)
    tFilialContato.sObservacao = String(STRING_CLIENTE_OBSERVACAO, 0)
    tFilialContato.sContaContabil = String(STRING_CONTA, 0)
    tFilialContato.sGuia = String(STRING_CLIENTE_GUIA, 0)

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 180798

    If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL And objFilialContato.lCodContato = 0 And objFilialContato.iCodFilialLoja <> 0 Then

        lCodContato = objFilialContato.lCodContatoLoja
        iCodFilial = objFilialContato.iCodFilialLoja

        'Pega os dados da tabela Filiais Contatos
        lErro = Comando_Executar(lComando, "SELECT Nome, CGC, InscricaoEstadual, InscricaoMunicipal,InscricaoSuframa, Endereco, EnderecoEntrega, EnderecoCobranca, CodTransportadora, Observacao, ContaContabil, Vendedor, ComissaoVendas, Regiao, FreqVisitas, DataUltVisita, CodCobrador, TipoFrete, RG, Ativo, CodTranspRedesp, DetPagFrete, Guia, CodMensagem FROM FiliaisContatos WHERE FilialEmpresaLoja = ? And CodContatoLoja = ? AND CodFilialLoja = ?", _
            tFilialContato.sNome, tFilialContato.sCgc, tFilialContato.sInscricaoEstadual, tFilialContato.sInscricaoMunicipal, tFilialContato.sInscricaoSuframa, tFilialContato.lEndereco, tFilialContato.lEnderecoEntrega, tFilialContato.lEnderecoCobranca, tFilialContato.iCodTransportadora, tFilialContato.sObservacao, tFilialContato.sContaContabil, tFilialContato.iVendedor, tFilialContato.dComissaoVendas, tFilialContato.iRegiao, tFilialContato.iFreqVisitas, tFilialContato.dtDataUltVisita, tFilialContato.iCodCobrador, tFilialContato.iTipoFrete, tFilialContato.sRG, iAtivo, tFilialContato.iCodTranspRedesp, tFilialContato.iDetPagFrete, tFilialContato.sGuia, tFilialContato.iCodMensagem, objFilialContato.iFilialEmpresaLoja, objFilialContato.lCodContatoLoja, objFilialContato.iCodFilialLoja)

    ElseIf giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL And objFilialContato.lCodContato <> 0 And objFilialContato.iCodFilialLoja <> 0 Then

        lCodContato = objFilialContato.lCodContato
        iCodFilial = objFilialContato.iCodFilialLoja

        'Pega os dados da tabela Filiais Contatos
        lErro = Comando_Executar(lComando, "SELECT Nome, CGC, InscricaoEstadual, InscricaoMunicipal,InscricaoSuframa, Endereco, EnderecoEntrega, EnderecoCobranca, CodTransportadora, Observacao, ContaContabil, Vendedor, ComissaoVendas, Regiao, FreqVisitas, DataUltVisita, CodCobrador, TipoFrete, RG, Ativo, CodTranspRedesp, DetPagFrete, Guia, CodMensagem FROM FiliaisContatos WHERE FilialEmpresaLoja = ? AND CodContato = ? AND CodFilialLoja = ?", _
            tFilialContato.sNome, tFilialContato.sCgc, tFilialContato.sInscricaoEstadual, tFilialContato.sInscricaoMunicipal, tFilialContato.sInscricaoSuframa, tFilialContato.lEndereco, tFilialContato.lEnderecoEntrega, tFilialContato.lEnderecoCobranca, tFilialContato.iCodTransportadora, tFilialContato.sObservacao, tFilialContato.sContaContabil, tFilialContato.iVendedor, tFilialContato.dComissaoVendas, tFilialContato.iRegiao, tFilialContato.iFreqVisitas, tFilialContato.dtDataUltVisita, tFilialContato.iCodCobrador, tFilialContato.iTipoFrete, tFilialContato.sRG, iAtivo, tFilialContato.iCodTranspRedesp, tFilialContato.iDetPagFrete, tFilialContato.sGuia, tFilialContato.iCodMensagem, objFilialContato.iFilialEmpresaLoja, objFilialContato.lCodContato, objFilialContato.iCodFilialLoja)

    Else
        
        lCodContato = objFilialContato.lCodContato
        iCodFilial = objFilialContato.iCodFilial
        
        'Pega os dados da tabela Filiais Contatos
        lErro = Comando_Executar(lComando, "SELECT Nome, CGC, InscricaoEstadual, InscricaoMunicipal,InscricaoSuframa, Endereco, EnderecoEntrega, EnderecoCobranca, CodTransportadora, Observacao, ContaContabil, Vendedor, ComissaoVendas, Regiao, FreqVisitas, DataUltVisita, CodCobrador, TipoFrete, RG, Ativo, CodTranspRedesp, DetPagFrete, Guia, CodMensagem FROM FiliaisContatos WHERE CodContato = ? AND CodFilial = ?", _
            tFilialContato.sNome, tFilialContato.sCgc, tFilialContato.sInscricaoEstadual, tFilialContato.sInscricaoMunicipal, tFilialContato.sInscricaoSuframa, tFilialContato.lEndereco, tFilialContato.lEnderecoEntrega, tFilialContato.lEnderecoCobranca, tFilialContato.iCodTransportadora, tFilialContato.sObservacao, tFilialContato.sContaContabil, tFilialContato.iVendedor, tFilialContato.dComissaoVendas, tFilialContato.iRegiao, tFilialContato.iFreqVisitas, tFilialContato.dtDataUltVisita, tFilialContato.iCodCobrador, tFilialContato.iTipoFrete, tFilialContato.sRG, iAtivo, tFilialContato.iCodTranspRedesp, tFilialContato.iDetPagFrete, tFilialContato.sGuia, tFilialContato.iCodMensagem, objFilialContato.lCodContato, objFilialContato.iCodFilial)

    End If
    
    If lErro <> AD_SQL_SUCESSO Then gError 180799

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 180800

    If lErro = AD_SQL_SEM_DADOS Then gError 180801

    objFilialContato.sNome = tFilialContato.sNome
    objFilialContato.sCgc = tFilialContato.sCgc
    objFilialContato.sRG = tFilialContato.sRG
    objFilialContato.sInscricaoEstadual = tFilialContato.sInscricaoEstadual
    objFilialContato.sInscricaoMunicipal = tFilialContato.sInscricaoMunicipal
    objFilialContato.sInscricaoSuframa = tFilialContato.sInscricaoSuframa
    objFilialContato.lEndereco = tFilialContato.lEndereco
    objFilialContato.lEnderecoEntrega = tFilialContato.lEnderecoEntrega
    objFilialContato.lEnderecoCobranca = tFilialContato.lEnderecoCobranca
    objFilialContato.iCodTransportadora = tFilialContato.iCodTransportadora
    objFilialContato.sObservacao = tFilialContato.sObservacao
    objFilialContato.sContaContabil = tFilialContato.sContaContabil
    objFilialContato.iVendedor = tFilialContato.iVendedor
    objFilialContato.dComissaoVendas = tFilialContato.dComissaoVendas
    objFilialContato.iRegiao = tFilialContato.iRegiao
    objFilialContato.iFreqVisitas = tFilialContato.iFreqVisitas
    objFilialContato.dtDataUltVisita = tFilialContato.dtDataUltVisita
    objFilialContato.iCodCobrador = tFilialContato.iCodCobrador
    objFilialContato.iTipoFrete = tFilialContato.iTipoFrete
    objFilialContato.iAtivo = iAtivo
    objFilialContato.iCodTranspRedesp = tFilialContato.iCodTranspRedesp
    objFilialContato.iDetPagFrete = tFilialContato.iDetPagFrete
    objFilialContato.sGuia = tFilialContato.sGuia
    objFilialContato.iCodMensagem = tFilialContato.iCodMensagem
    
    Call Comando_Fechar(lComando)

    FilialContato_Le = SUCESSO

    Exit Function

Erro_FilialContato_Le:

    FilialContato_Le = gErr

    Select Case gErr

        Case 180798
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 180799, 180800
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISContatoS", gErr, lCodContato, iCodFilial)

        Case 180801 'nao encontrou a filial (sem dados)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180802)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function


Public Function RelacionamentoContatos_Le(ByVal objRelacionamentoClientes As ClassRelacClientes) As Long
'objRelacionamentoClientes RECEBE(Input) código/filialempresa do relacionamento que deve ser lido
'objRelacionamentoClientes RETORNA(Output) os dados do relacionamento lido

Dim lErro As Long
Dim lComando As Long
Dim tRelacionamentoClientes As typeRelacionamentoClientes

On Error GoTo Erro_RelacionamentoContatos_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 181028
    
    'Inicializa as strings para leitura de dados
    tRelacionamentoClientes.sAssunto1 = String(STRING_BUFFER_MAX_TEXTO, 0)
    tRelacionamentoClientes.sAssunto2 = String(STRING_BUFFER_MAX_TEXTO, 0)
    
    'Lê no BD o relacionamento com código e filialempresa passados como parâmetro
    lErro = Comando_Executar(lComando, "SELECT Codigo, FilialEmpresa, Origem, Data, Hora, Tipo, Cliente, FilialCliente, Contato, Atendente, RelacionamentoAnt, Assunto1, Assunto2, Status FROM RelacionamentoContatos WHERE Codigo = ? AND FilialEmpresa = ?", tRelacionamentoClientes.lCodigo, tRelacionamentoClientes.iFilialEmpresa, tRelacionamentoClientes.iOrigem, tRelacionamentoClientes.dtData, tRelacionamentoClientes.dHora, tRelacionamentoClientes.lTipo, tRelacionamentoClientes.lCliente, tRelacionamentoClientes.iFilialCliente, tRelacionamentoClientes.iContato, tRelacionamentoClientes.iAtendente, tRelacionamentoClientes.lRelacionamentoAnt, tRelacionamentoClientes.sAssunto1, tRelacionamentoClientes.sAssunto2, tRelacionamentoClientes.iStatus, objRelacionamentoClientes.lCodigo, objRelacionamentoClientes.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 181029
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 181030
    
    'Se não encontrou o relacionamento => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 181031
    
    'Transfere os dados lidos para o obj
    With objRelacionamentoClientes
    
        .lCodigo = tRelacionamentoClientes.lCodigo
        .iFilialEmpresa = tRelacionamentoClientes.iFilialEmpresa
        .iOrigem = tRelacionamentoClientes.iOrigem
        .dtData = tRelacionamentoClientes.dtData
        .dtHora = CDate(tRelacionamentoClientes.dHora)
        .lTipo = tRelacionamentoClientes.lTipo
        .lCliente = tRelacionamentoClientes.lCliente
        .iFilialCliente = tRelacionamentoClientes.iFilialCliente
        .iContato = tRelacionamentoClientes.iContato
        .iAtendente = tRelacionamentoClientes.iAtendente
        .lRelacionamentoAnt = tRelacionamentoClientes.lRelacionamentoAnt
        .sAssunto1 = tRelacionamentoClientes.sAssunto1
        .sAssunto2 = tRelacionamentoClientes.sAssunto2
        .iStatus = tRelacionamentoClientes.iStatus
    
    End With

    'Executa o fechamento do Comando
    Call Comando_Fechar(lComando)

    RelacionamentoContatos_Le = SUCESSO

    Exit Function

Erro_RelacionamentoContatos_Le:

    RelacionamentoContatos_Le = gErr

    Select Case gErr

        Case 181028
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 181029, 181030
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELACIONAMENTOCLIENTES", gErr)
            
        Case 181031 'Registro não encontrado
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181032)

    End Select

    'Executa o fechamento do Comando
    Call Comando_Fechar(lComando)

End Function

'Incluída por Luiz Nogueira em 27/10/03
Public Function RelacionamentoContatos_Locks_Relacionados(ByVal objRelacionamentoClientes As ClassRelacClientes) As Long
'Efetua locks nas tabelas relacionadas à gravação de relacionamentos

Dim lErro As Long

On Error GoTo Erro_RelacionamentoContatos_Locks_Relacionados

    'Verifica se Cliente e Filial estão cadastrados
    lErro = FilialContato_Lock(objRelacionamentoClientes.lCliente, objRelacionamentoClientes.iFilialCliente)
    If lErro <> SUCESSO And lErro <> 181036 Then gError 181039
    
    'Não encontrou a Filial Cliente --> erro
    If lErro = 43050 Then gError 181040
    
    'Efetua lock do atendente
    lErro = CF("Atendentes_LockShared", objRelacionamentoClientes.iAtendente, objRelacionamentoClientes.iFilialEmpresa)
    If lErro <> SUCESSO And lErro <> 102552 Then gError 181041
    
    'Se não encontrou o campo
    If lErro = 102552 Then gError 181042
    
    'Efetua lock do Relacionamento Anterior
    lErro = RelacionamentoContatos_LockShared(objRelacionamentoClientes.lRelacionamentoAnt, objRelacionamentoClientes.iFilialEmpresa)
    If lErro <> SUCESSO And lErro <> 181048 Then gError 181043
    
    RelacionamentoContatos_Locks_Relacionados = SUCESSO

    Exit Function

Erro_RelacionamentoContatos_Locks_Relacionados:

    RelacionamentoContatos_Locks_Relacionados = gErr

    Select Case gErr
    
        Case 181039, 181041
        
        Case 181040
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_CADASTRADA", gErr, objRelacionamentoClientes.iFilialCliente, objRelacionamentoClientes.lCliente)

        Case 181042
            Call Rotina_Erro(vbOKOnly, "ERRO_ATENDENTE_NAO_ENCONTRADO", gErr, objRelacionamentoClientes.iAtendente, objRelacionamentoClientes.iFilialEmpresa)
        
        Case 181043
            Call Rotina_Erro(vbOKOnly, "ERRO_RELACIONAMENTO_NAO_ENCONTRADO", gErr, objRelacionamentoClientes.lRelacionamentoAnt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181044)

    End Select

End Function

'Incluído por Luiz Nogueira em 27/10/03
Public Function RelacionamentoContatos_LockShared(ByVal lCodigo As Long, Optional ByVal iFilialEmpresa As Integer) As Long
'Faz lock shared no relacionamento passado como parâmetro
'lCodigo recebe o código do relacionamento a ser "lockado"

Dim lErro As Long
Dim lComando As Long
Dim iIndice As Integer
Dim lCodAux As Long

On Error GoTo Erro_RelacionamentoContatos_LockShared

    'Executa a abertura dos Comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 181045

    'Se não indicou a filial do atendente => usa a filial empresa ativa
    If iFilialEmpresa = 0 Then iFilialEmpresa = giFilialEmpresa
    
    'Busca o relacionamento a ser lockado
    lErro = Comando_ExecutarLockado(lComando, "SELECT Codigo FROM RelacionamentoContatos WHERE Codigo=? AND FilialEmpresa=?", lCodAux, lCodigo, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 181046
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 181047
    
    'Se não encontrou o relacionamento => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 181048
    
    'Efetua o lock do relacionamento
    lErro = Comando_LockShared(lComando)
    If lErro <> AD_SQL_SUCESSO Then gError 181049
    
    'Executa o fechamento dos Comandos
    Call Comando_Fechar(lComando)

    RelacionamentoContatos_LockShared = SUCESSO

    Exit Function

Erro_RelacionamentoContatos_LockShared:

    RelacionamentoContatos_LockShared = gErr

    Select Case gErr

        Case 181045
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 181046, 181047
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELACIONAMENTOCLIENTES", gErr)

        Case 181048

        Case 181049
            Call Rotina_Erro(vbOKOnly, "ERRO_CAMPOGENERICO_LOCKSHARED", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181050)

    End Select

    Call Comando_Fechar(lComando)

End Function

'Incluída por Luiz Nogueira em 27/10/03
Public Function RelacionamentoContatos_Verifica_Cliente(ByVal lCliente As Long, lCodRelacionamento As Long, Optional ByVal iFilialCliente As Integer) As Long
'Verifica se existe algum relacionamento para o cliente ou filial do cliente em questão
'lCliente recebe o código do cliente para o qual deve ser verificado se existe relacionamento
'iFilialCliente recebe o código da filial para a qual deve ser feita a mesma verificação
'Se iFilialCliente for 0, a verificação será feita para todas as filiais do cliente

Dim lErro As Long
Dim lComando As Long

On Error GoTo Erro_RelacionamentoContatos_Verifica_Cliente

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 181051
    
    'Se foi passada uma filial de Cliente
    If iFilialCliente <> 0 Then
    
        'Verifica se existe relacionamento para a filial do cliente com código passado como parâemtro
        lErro = Comando_Executar(lComando, "SELECT Codigo FROM RelacionamentoContatos WHERE Cliente=? AND FilialCliente=?", lCodRelacionamento, lCliente, iFilialCliente)
        If lErro <> AD_SQL_SUCESSO Then gError 181052
    
    'Senão, ou seja, se não foi passada uma filial específica
    Else
    
        'Verifica se existe relacionamento para qualquer filial do cliente com código passado como parâemtro
        lErro = Comando_Executar(lComando, "SELECT Codigo FROM RelacionamentoContatos WHERE Cliente=?", lCodRelacionamento, lCliente)
        If lErro <> AD_SQL_SUCESSO Then gError 181053
    
    End If
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 181054
    
    'Se não encontrou => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 181055

    'Executa o fechamento do Comando
    Call Comando_Fechar(lComando)

    RelacionamentoContatos_Verifica_Cliente = SUCESSO

    Exit Function

Erro_RelacionamentoContatos_Verifica_Cliente:

    RelacionamentoContatos_Verifica_Cliente = gErr

    Select Case gErr

        Case 181051
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 181053, 181054, 181052
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELACIONAMENTOCLIENTES", gErr)
        
        Case 181055 'Não encontrado

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181056)

    End Select

    'Executa o fechamento do Comando
    Call Comando_Fechar(lComando)

End Function

'Incluída por Luiz Nogueira em 27/10/03
Public Function RelacionamentoContatos_Valida_Alteracao(ByVal objRelacionamentoClientes As ClassRelacClientes, bAlteracaoValida As Boolean) As Long
'Verifica se o relacionamento passado como parâmetro está sendo apontado por outro relacionamento como relacionamento anterior
'e com a data anterior a este relacionamento
'Isso é para evitar que a data de um relacionamento seja alterada e torne inconsistente os demais relacionamentos que o utilizam
'como relacionamento anterior

Dim lErro As Long
Dim lComando As Long
Dim lCodigo As Long

On Error GoTo Erro_RelacionamentoContatos_Valida_Alteracao

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 181057

    'Verifica se o relacionamento que está sendo alterado é apontado por outro relacionamento
    'como relacionamento anterior
    lErro = Comando_ExecutarPos(lComando, "SELECT Codigo FROM RelacionamentoContatos WHERE RelacionamentoAnt=? AND FilialEmpresa=? AND (Cliente<>? OR Data<?)", 0, lCodigo, objRelacionamentoClientes.lCodigo, objRelacionamentoClientes.iFilialEmpresa, objRelacionamentoClientes.lCliente, objRelacionamentoClientes.dtData)
    If lErro <> AD_SQL_SUCESSO Then gError 181058
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 181059
    
    'Se encontrou => significa que a data do relacionamento não pode ser alterada
    If lErro = AD_SQL_SUCESSO Then bAlteracaoValida = False
    
    'Executa o fechamento do Comando
    Call Comando_Fechar(lComando)

    RelacionamentoContatos_Valida_Alteracao = SUCESSO

    Exit Function

Erro_RelacionamentoContatos_Valida_Alteracao:

    RelacionamentoContatos_Valida_Alteracao = gErr

    Select Case gErr

        Case 181057
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 181058, 181059
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELACIONAMENTOCLIENTES", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181060)

    End Select

    'Executa o fechamento do Comando
    Call Comando_Fechar(lComando)

End Function

Function FilialContato_Lock(ByVal lCodContato As Long, ByVal iCodFilial As Integer) As Long
'Faz lock da filial do Contato passado por parâmetro
'Chamada DENTRO DE TRANSAÇÃO

Dim lErro As Long
Dim lComando As Long
Dim lContato As Long
Dim iFilial As Integer

On Error GoTo Erro_FilialContato_Lock

    'Inicializa comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 181033
    
    'Lê a Filial do Contato que será lockada
    lErro = Comando_ExecutarLockado(lComando, "SELECT CodContato, CodFilial FROM FiliaisContatos WHERE CodContato =? AND CodFilial = ?", lContato, iFilial, lCodContato, iCodFilial)
    If lErro <> AD_SQL_SUCESSO Then gError 181034

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 181035

    'Não achou a Filial do Contato
    If lErro = AD_SQL_SEM_DADOS Then gError 181036

    'Lock da Filial do Contato
    lErro = Comando_LockShared(lComando)
    If lErro <> AD_SQL_SUCESSO Then gError 181037

    'Libera comandos
    Call Comando_Fechar(lComando)

    FilialContato_Lock = SUCESSO

    Exit Function

Erro_FilialContato_Lock:

    FilialContato_Lock = gErr

    Select Case gErr
        
        Case 181033
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
                
        Case 181034, 181035
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISCONTATOS", gErr, lCodContato, iCodFilial)

        Case 181036
            'Erro tratado na rotina chamada

        Case 181037
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_FILIAISCONTATOS", gErr, lCodContato, iCodFilial)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181038)

    End Select

    'Libera comandos
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function Carrega_ClienteFContatos(ByVal objControle As Object, objClienteFContatos As ClassClienteContatos) As Long
'Carrega uma combo com os contatos de um determinado cliente/filial
'objControle é a combo que será carregada
'objClienteFContatos recebe o código do cliente e da filial que terão seus contatos adicionados à combo

Dim lErro As Long
Dim colClienteFContatos As New Collection

On Error GoTo Erro_Carrega_ClienteFContatos

    'Limpa a combo
    objControle.Clear
    
    'Le os dados da tabela ContatoGeral e Preenche a colContatoGeral
    lErro = CF("ClienteFContatos_Le_Todos", colClienteFContatos, objClienteFContatos)
    If lErro <> SUCESSO And lErro <> 181086 Then gError 181069
    
    'Se não encontrou contatos => erro
    If lErro = 181086 Then gError 181070
    
    'Para cada contato encontrado
    For Each objClienteFContatos In colClienteFContatos
    
        'Adiciona o contato a combo
        objControle.AddItem objClienteFContatos.iCodigo & SEPARADOR & objClienteFContatos.sContato
        objControle.ItemData(objControle.NewIndex) = objClienteFContatos.iCodigo
        
        'Se o contato adicionado está marcado como padrão => seleciona-o
        If objClienteFContatos.iPadrao = MARCADO Then objControle.ListIndex = objControle.NewIndex
    
    Next

    Carrega_ClienteFContatos = SUCESSO

    Exit Function

Erro_Carrega_ClienteFContatos:

    Carrega_ClienteFContatos = gErr

    Select Case gErr

        Case 181069
        
        Case 181070 'Sem dados... deve ser tratado na rotina chamadora
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181071)

    End Select

End Function

'Incluída por Luiz Nogueira em 27/10/03
Public Function ClienteFContatos_Le(ByVal objClienteFContatos As ClassClienteContatos) As Long
'Lê um contato a partir de cliente/filial + código do contato
'objContatoClientes recebe o código do cliente e da filial
'objContatoClientes retorna os dados lidos

Dim lErro As Long
Dim lComando As Long
Dim tClienteFContatos As typeClienteContatos

On Error GoTo Erro_ClienteFContatos_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 181072
    
    'Inicializa as string que serão lidas
    tClienteFContatos.sContato = String(STRING_CONTATOGERAL_CONTATO, 0)
    tClienteFContatos.sSetor = String(STRING_CONTATOGERAL_SETOR, 0)
    tClienteFContatos.sCargo = String(STRING_CONTATOGERAL_CARGO, 0)
    tClienteFContatos.sTelefone = String(STRING_TELEFONE, 0)
    tClienteFContatos.sFax = String(STRING_FAX, 0)
    tClienteFContatos.sEmail = String(STRING_EMAIL, 0)
    
    'Lê o contato no BD
    lErro = Comando_Executar(lComando, "SELECT Codigo, Contato, Setor, Cargo, Telefone, Fax, Email, Padrao FROM ClienteFContatos WHERE Cliente=? AND FilialCliente=? AND Codigo=?", tClienteFContatos.iCodigo, tClienteFContatos.sContato, tClienteFContatos.sSetor, tClienteFContatos.sCargo, tClienteFContatos.sTelefone, tClienteFContatos.sFax, tClienteFContatos.sEmail, tClienteFContatos.iPadrao, objClienteFContatos.lCliente, objClienteFContatos.iFilialCliente, objClienteFContatos.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 181073
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 181074
    
    'Se não encontrou o contato => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 181075
    
    'Move os dados lidos para o obj
    With objClienteFContatos
        .iCodigo = tClienteFContatos.iCodigo
        .sContato = tClienteFContatos.sContato
        .sSetor = tClienteFContatos.sSetor
        .sCargo = tClienteFContatos.sCargo
        .sTelefone = tClienteFContatos.sTelefone
        .sFax = tClienteFContatos.sFax
        .sEmail = tClienteFContatos.sEmail
        .iPadrao = tClienteFContatos.iPadrao
    End With
        
    'Executa o fechamento do Comando
    Call Comando_Fechar(lComando)

    ClienteFContatos_Le = SUCESSO

    Exit Function

Erro_ClienteFContatos_Le:

    ClienteFContatos_Le = gErr

    Select Case gErr

        Case 181072
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 181073, 181074
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ClienteFContatoS", gErr)
        
        Case 181075 'Não encontrado
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181076)

    End Select

    'Executa o fechamento do Comando
    Call Comando_Fechar(lComando)

End Function

Public Function ClienteFContatos_Le_Telefone(ByVal sTelefone As String, ByVal colClienteFContatos As Collection) As Long
'Lê no bd os contatos de clientes que possuem o telefone passado como parâmetro
'sTelefone recebe o telefone que será usado para pesquisa
'colClienteFContatos retorna os contatos encontrados com o telefone em questão

Dim lErro As Long
Dim lComando As Long
Dim tClienteFContatos As typeClienteContatos
Dim objClienteFContatos As ClassClienteContatos

On Error GoTo Erro_ClienteFContatos_Le_Telefone

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 181077
    
    'Inicializa as strings que serão usadas na leitura
    tClienteFContatos.sContato = String(STRING_CONTATOGERAL_CONTATO, 0)
    tClienteFContatos.sSetor = String(STRING_CONTATOGERAL_SETOR, 0)
    tClienteFContatos.sCargo = String(STRING_CONTATOGERAL_CARGO, 0)
    tClienteFContatos.sTelefone = String(STRING_TELEFONE, 0)
    tClienteFContatos.sFax = String(STRING_FAX, 0)
    tClienteFContatos.sEmail = String(STRING_EMAIL, 0)
    
    'Busca no BD um contato de cliente com o telefone passado como parâmetro
    lErro = Comando_Executar(lComando, "SELECT Cliente, FilialCliente, Codigo, Contato, Setor, Cargo, Telefone, Fax, Email, Padrao FROM ClienteFContatos WHERE Telefone=?", tClienteFContatos.lCliente, tClienteFContatos.iFilialCliente, tClienteFContatos.iCodigo, tClienteFContatos.sContato, tClienteFContatos.sSetor, tClienteFContatos.sCargo, tClienteFContatos.sTelefone, tClienteFContatos.sFax, tClienteFContatos.sEmail, tClienteFContatos.iPadrao, sTelefone)
    If lErro <> AD_SQL_SUCESSO Then gError 181078
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 181079
    
    'Se não encontrou o contato => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 181080

    'Para cada contato encontrado
    Do While lErro = AD_SQL_SUCESSO
    
        'Instancia um novo obj
        Set objClienteFContatos = New ClassClienteContatos
        
        'Transfere os dados lidos para obj
        With objClienteFContatos
            .lCliente = tClienteFContatos.lCliente
            .iFilialCliente = tClienteFContatos.iFilialCliente
            .iCodigo = tClienteFContatos.iCodigo
            .sContato = tClienteFContatos.sContato
            .sSetor = tClienteFContatos.sSetor
            .sCargo = tClienteFContatos.sCargo
            .sTelefone = tClienteFContatos.sTelefone
            .sFax = tClienteFContatos.sFax
            .sEmail = tClienteFContatos.sEmail
            .iPadrao = tClienteFContatos.iPadrao
        End With
        
        'Adiciona o contato à coleção
        colClienteFContatos.Add objClienteFContatos
        
        'Busca o próximo contato encontrado
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 181081
    
    Loop

    'Executa o fechamento do Comando
    Call Comando_Fechar(lComando)

    ClienteFContatos_Le_Telefone = SUCESSO

    Exit Function

Erro_ClienteFContatos_Le_Telefone:

    ClienteFContatos_Le_Telefone = gErr

    Select Case gErr

        Case 181077
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 181078, 181079, 181081
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLIENTECONTATOS", gErr)
            
        Case 181080 'Não encontrado

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181082)

    End Select

    'Executa o fechamento do Comando
    Call Comando_Fechar(lComando)

End Function

Public Function ClienteFContatos_Le_Todos(ByVal colClienteFContatos As Collection, objClienteFContatos As ClassClienteContatos) As Long
'Lê todos os contatos para o cliente/filial passados como parâmetros
'objClienteFContatos recebe o código do cliente e da filial
'colClienteFContatos retorna os contatos encontrados

Dim lErro As Long

Dim lComando As Long
Dim tClienteFContatos As typeClienteContatos
Dim objClienteFContatosAux As ClassClienteContatos

On Error GoTo Erro_ClienteFContatos_Le_Todos

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 181083

    'Inicializa as string que serão lidas
    tClienteFContatos.sContato = String(STRING_CONTATOGERAL_CONTATO, 0)
    tClienteFContatos.sSetor = String(STRING_CONTATOGERAL_SETOR, 0)
    tClienteFContatos.sCargo = String(STRING_CONTATOGERAL_CARGO, 0)
    tClienteFContatos.sTelefone = String(STRING_TELEFONE, 0)
    tClienteFContatos.sFax = String(STRING_FAX, 0)
    tClienteFContatos.sEmail = String(STRING_EMAIL, 0)
    
    'Lê os contatos no BD
    lErro = Comando_Executar(lComando, "SELECT Codigo, Contato, Setor, Cargo, Telefone, Fax, Email, Padrao FROM ClienteFContatos WHERE Cliente=? AND FilialCliente=?", tClienteFContatos.iCodigo, tClienteFContatos.sContato, tClienteFContatos.sSetor, tClienteFContatos.sCargo, tClienteFContatos.sTelefone, tClienteFContatos.sFax, tClienteFContatos.sEmail, tClienteFContatos.iPadrao, objClienteFContatos.lCliente, objClienteFContatos.iFilialCliente)
    If lErro <> AD_SQL_SUCESSO Then gError 181084
    
    'Busca o primeiro contato encontrado
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 181085
    
    'Se não encontrou contatos cadastrados para o cliente em questão
    If lErro = AD_SQL_SEM_DADOS Then gError 181086
    
    'Enquanto houverem contatos
    Do While lErro <> AD_SQL_SEM_DADOS
    
        'Instancia um novo objClienteFContatos
        Set objClienteFContatosAux = New ClassClienteContatos
        
        'Move os dados lidos para o obj
        With objClienteFContatosAux
            .iCodigo = tClienteFContatos.iCodigo
            .sContato = tClienteFContatos.sContato
            .sSetor = tClienteFContatos.sSetor
            .sCargo = tClienteFContatos.sCargo
            .sTelefone = tClienteFContatos.sTelefone
            .sFax = tClienteFContatos.sFax
            .sEmail = tClienteFContatos.sEmail
            .iPadrao = tClienteFContatos.iPadrao
            .lCliente = objClienteFContatos.lCliente
            .iFilialCliente = objClienteFContatos.iFilialCliente
        End With
        
        'Guarda na coleção, os dados do contato
        colClienteFContatos.Add objClienteFContatosAux
        
        'Busca o próximo contato encontrado
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 181087
            
    Loop
    
    'Executa o fechamento do Comando
    Call Comando_Fechar(lComando)

    ClienteFContatos_Le_Todos = SUCESSO

    Exit Function

Erro_ClienteFContatos_Le_Todos:

    ClienteFContatos_Le_Todos = gErr

    Select Case gErr

        Case 181083
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 181084, 181085, 181087
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ClienteFContatos", gErr)
        
        Case 181086 'Não encontrado
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181088)

    End Select

    'Executa o fechamento do Comando
    Call Comando_Fechar(lComando)

End Function

Public Function ClienteFContatos_Valida_Exclusao(ByVal objClienteFContatos As ClassClienteContatos) As Long
'Verifica se sobrou alguma inconsistência no BD após a exclusão de contatos de cliente
'A tabela verifica se foram excluídos contatos que estão sendo referenciados em outras tabelas do sistema
'Caso tenha encontrado alguma inconsistência, sai por erro e a exclusão não é efetuada
'objClienteFContatos recebe código do cliente/filial para os quais deve ser checada a consistência

Dim lErro As Long
Dim lComando As Long
Dim iContato As Integer
Dim lCliente As Long
Dim iFilialCliente As Integer
Dim lRelacionamentoCliente As Long

On Error GoTo Erro_ClienteFContatos_Valida_Exclusao

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 181089
    
    'Busca os registro em RelacionamentoContatos que utilizam um contato que não
    'está mais na tabela ClienteFContatos
    lErro = Comando_Executar(lComando, "SELECT RelacionamentoContatos.Codigo, RelacionamentoContatos.Cliente, RelacionamentoContatos.FilialCliente, RelacionamentoContatos.Contato FROM RelacionamentoContatos WHERE RelacionamentoContatos.Contato NOT IN (SELECT Codigo FROM ClienteFContatos WHERE ClienteFContatos.Cliente=? AND ClienteFContatos.FilialCliente=?) AND Contato<>0 AND RelacionamentoContatos.Cliente=? AND RelacionamentoContatos.FilialCliente=?", lRelacionamentoCliente, lCliente, iFilialCliente, iContato, objClienteFContatos.lCliente, objClienteFContatos.iFilialCliente, objClienteFContatos.lCliente, objClienteFContatos.iFilialCliente)
    If lErro <> AD_SQL_SUCESSO Then gError 181090
    
    'Posiciona no primeiro registro encontrado
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 181091
    
    'Se encontrou algum registro => erro, pois significa que foi excluído um contato utilizado em outra tabela
    If lErro = AD_SQL_SUCESSO Then gError 181092

    'Executa o fechamento do Comando
    Call Comando_Fechar(lComando)

    ClienteFContatos_Valida_Exclusao = SUCESSO

    Exit Function

Erro_ClienteFContatos_Valida_Exclusao:

    ClienteFContatos_Valida_Exclusao = gErr

    Select Case gErr

        Case 181089
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 181090, 181091
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLIENTECONTATOS", gErr)
        
        Case 181092
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CLIENTECONTATOS_INVALIDA", gErr, iContato, lCliente, iFilialCliente, lRelacionamentoCliente)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181093)

    End Select

    'Executa o fechamento do Comando
    Call Comando_Fechar(lComando)

End Function

Public Function RelacionamentoContatosCons_Le(ByVal objRelacionamentoContatosCons As ClassRelacClientesCons) As Long

Dim lErro As Long
Dim lComando As Long
Dim tRelacionamentoContatosCons As typeRelacionamentoClientesCons
Dim sSelect As String

On Error GoTo Erro_RelacionamentoContatosCons_Le

    'Abre o comando para leitura
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 181099

    'Monta o Select que será executado
    Call RelacionamentoContatosCons_Le1(sSelect, objRelacionamentoContatosCons)

    'Faz o Bind dos campos que serão selecionados
    lErro = RelacionamentoContatosCons_Le2(lComando, sSelect, tRelacionamentoContatosCons)
    If lErro <> SUCESSO Then gError 181100

    'Faz o Bind dos campos que estão sendo passados para a cláusula WHERE
    lErro = RelacionamentoContatosCons_Le3(lComando, tRelacionamentoContatosCons, objRelacionamentoContatosCons)
    If lErro <> SUCESSO Then gError 181101

    'Executa a leitura do BD e guarda os dados lidos no obj
    lErro = RelacionamentoContatosCons_Le4(lComando, tRelacionamentoContatosCons, objRelacionamentoContatosCons)
    If lErro <> SUCESSO And lErro <> 181096 Then gError 181102

    'Se não encontrou => erro. O erro deve ser tratado na rotina chamadora
    If lErro = 181096 Then gError 181103

    Call Comando_Fechar(lComando)

    RelacionamentoContatosCons_Le = SUCESSO

    Exit Function

Erro_RelacionamentoContatosCons_Le:

    RelacionamentoContatosCons_Le = gErr

    Select Case gErr

        Case 181099
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr, Error$)

        Case 181100 To 181102

        Case 181103 'Sem Dados

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 181104)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

'Incluído por Luiz Nogueira em 27/10/03
Private Sub RelacionamentoContatosCons_Le1(sSelect As String, ByVal objRelacionamentoContatosCons As ClassRelacClientesCons)

Dim sFrom As String
Dim sWhere As String
Dim sOrderBy As String

    'Monta a parte fixa do select
    sSelect = "SELECT Codigo, FilialEmpresa, Origem, Data, Hora, Tipo, Cliente, FilialCliente, Contato, Atendente, RelacionamentoAnt, Assunto1, Assunto2, Status "
    sFrom = "FROM RelacionamentoContatos "
    sWhere = "WHERE FilialEmpresa = ? "
    sOrderBy = "ORDER BY Data, Hora, Atendente, Cliente, FilialCliente"

    'Monta a parte dinâmica da cláusula WHERE
    'Se o código inicial do relacionamento foi informado
    If objRelacionamentoContatosCons.lCodigoDe > 0 Then
    
        'Acrescenta à cláusula WHERE o filtro de Código maior igual ao código inicial informado
        sWhere = sWhere & "AND Codigo>=? "
    
    End If
    
    'Se o código final do relacionamento foi informado
    If objRelacionamentoContatosCons.lCodigoAte > 0 Then
    
        'Acrescenta à cláusula WHERE o filtro de Código menor igual ao código final informado
        sWhere = sWhere & "AND Codigo<=? "
    
    End If
    
    'Se a data inicial está preenchida
    If objRelacionamentoContatosCons.dtDataDe <> DATA_NULA Then

        'Acrescenta à cláusula WHERE o filtro de Data maior igual à Data Inicial informada
        sWhere = sWhere & "AND Data>=? "

    End If

    'Se a data final está preenchida
    If objRelacionamentoContatosCons.dtDataAte <> DATA_NULA Then

        'Acrescenta à cláusula WHERE o filtro de Data menor igual à Data Final informada
        sWhere = sWhere & "AND Data<=? "

    End If

    'Se o cliente está preenchido
    If objRelacionamentoContatosCons.lCliente > 0 Then

        'Acrescenta à cláusula WHERE o filtro de cliente
        sWhere = sWhere & "AND Cliente=? "

    End If

    'Se a filial do cliente está preenchida
    If objRelacionamentoContatosCons.iFilialCliente > 0 Then

        'Acrescenta à cláusula WHERE o filtro de filial de cliente
        sWhere = sWhere & "AND FilialCliente=? "

    End If

    'Se o atendente inicial do relacionamento foi informado
    If objRelacionamentoContatosCons.iAtendenteDe > 0 Then
    
        'Acrescenta à cláusula WHERE o filtro de Atendente maior igual ao atendente inicial informado
        sWhere = sWhere & "AND Atendente>=? "
    
    End If

    'Se o atendente final do relacionamento foi informado
    If objRelacionamentoContatosCons.iAtendenteAte > 0 Then
    
        'Acrescenta à cláusula WHERE o filtro de Atendente menor igual ao atendente final informado
        sWhere = sWhere & "AND Atendente<=? "
    
    End If
    
    'Se o tipo de relacionamento está preenchido
    If objRelacionamentoContatosCons.lTipo > 0 Then

        'Acrescenta à cláusula WHERE o filtro de tipo
        sWhere = sWhere & "AND Tipo=? "

    End If
    
    'Se a origem do relacionamento está preenchida
    If objRelacionamentoContatosCons.iOrigem > 0 Then

        'Acrescenta à cláusula WHERE o filtro de origem
        sWhere = sWhere & "AND Origem=? "

    End If
    
    'Se o status do relacionamento está preenchido
    If objRelacionamentoContatosCons.iStatus <> -1 Then

        'Acrescenta à cláusula WHERE o filtro de tipo
        sWhere = sWhere & "AND Status=? "

    End If
    
    sSelect = sSelect & sFrom & sWhere & sOrderBy

End Sub

'Incluído por Luiz Nogueira em 27/10/03
Private Function RelacionamentoContatosCons_Le2(ByVal lComando As Long, ByVal sSelect As String, tRelacionamentoContatosCons As typeRelacionamentoClientesCons) As Long

Dim iRetorno As Integer

On Error GoTo Erro_RelacionamentoContatosCons_Le2

    With tRelacionamentoContatosCons

        iRetorno = Comando_PrepararInt(lComando, sSelect)
        If (iRetorno <> AD_SQL_SUCESSO) Then gError 181105

        .vlCodigo = CLng(.vlCodigo)
        iRetorno = Comando_BindVarInt(lComando, .vlCodigo)
        If (iRetorno <> AD_SQL_SUCESSO) Then gError 181106

        .viFilialEmpresa = CInt(.viFilialEmpresa)
        iRetorno = Comando_BindVarInt(lComando, .viFilialEmpresa)
        If (iRetorno <> AD_SQL_SUCESSO) Then gError 181107

        .viOrigem = CInt(.viOrigem)
        iRetorno = Comando_BindVarInt(lComando, .viOrigem)
        If (iRetorno <> AD_SQL_SUCESSO) Then gError 181108

        .vdtData = CDate(.vdtData)
        iRetorno = Comando_BindVarInt(lComando, .vdtData)
        If (iRetorno <> AD_SQL_SUCESSO) Then gError 181109

        .vdHora = CDbl(.vdHora)
        iRetorno = Comando_BindVarInt(lComando, .vdHora)
        If (iRetorno <> AD_SQL_SUCESSO) Then gError 181110

        .vlTipo = CLng(.vlTipo)
        iRetorno = Comando_BindVarInt(lComando, .vlTipo)
        If (iRetorno <> AD_SQL_SUCESSO) Then gError 181111

        .vlCliente = CLng(.vlCliente)
        iRetorno = Comando_BindVarInt(lComando, .vlCliente)
        If (iRetorno <> AD_SQL_SUCESSO) Then gError 181112

        .viFilialCliente = CInt(.viFilialCliente)
        iRetorno = Comando_BindVarInt(lComando, .viFilialCliente)
        If (iRetorno <> AD_SQL_SUCESSO) Then gError 181113

        .viContato = CInt(.viContato)
        iRetorno = Comando_BindVarInt(lComando, .viContato)
        If (iRetorno <> AD_SQL_SUCESSO) Then gError 181114

         .viAtendente = CInt(.viAtendente)
        iRetorno = Comando_BindVarInt(lComando, .viAtendente)
        If (iRetorno <> AD_SQL_SUCESSO) Then gError 181115

        .vlRelacionamentoAnt = CLng(.vlRelacionamentoAnt)
        iRetorno = Comando_BindVarInt(lComando, .vlRelacionamentoAnt)
        If (iRetorno <> AD_SQL_SUCESSO) Then gError 181116

        .vsAssunto1 = CStr(.vsAssunto1)
        iRetorno = Comando_BindVarInt(lComando, .vsAssunto1)
        If (iRetorno <> AD_SQL_SUCESSO) Then gError 181117

        .vsAssunto2 = CStr(.vsAssunto2)
        iRetorno = Comando_BindVarInt(lComando, .vsAssunto2)
        If (iRetorno <> AD_SQL_SUCESSO) Then gError 181118

        .viStatus = CInt(.viStatus)
        iRetorno = Comando_BindVarInt(lComando, .viStatus)
        If (iRetorno <> AD_SQL_SUCESSO) Then gError 181119
    
    End With

    RelacionamentoContatosCons_Le2 = SUCESSO

    Exit Function

Erro_RelacionamentoContatosCons_Le2:

    RelacionamentoContatosCons_Le2 = gErr

    Select Case gErr

        Case 181105 To 181119
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELACIONAMENTOCLIENTES", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 181120)

    End Select

    Exit Function

End Function

Private Function RelacionamentoContatosCons_Le3(ByVal lComando As Long, tRelacionamentoContatosCons As typeRelacionamentoClientesCons, ByVal objRelacionamentoContatosCons As ClassRelacClientesCons) As Long

Dim iRetorno As Integer

On Error GoTo Erro_RelacionamentoContatosCons_Le3

    With tRelacionamentoContatosCons

        .viFilialEmpresa = objRelacionamentoContatosCons.iFilialEmpresa
        iRetorno = Comando_BindVarInt(lComando, .viFilialEmpresa)
        If (iRetorno <> AD_SQL_SUCESSO) Then gError 181121

        'Faz o bind da parte dinâmica da cláusula WHERE
        'Se o código inicial do relacionamento foi informado
        If objRelacionamentoContatosCons.lCodigoDe > 0 Then
            .vlCodigoDe = objRelacionamentoContatosCons.lCodigoDe
            iRetorno = Comando_BindVarInt(lComando, .vlCodigoDe)
            If (iRetorno <> AD_SQL_SUCESSO) Then gError 181122
        End If

        'Se o código final do relacionamento foi informado
        If objRelacionamentoContatosCons.lCodigoAte > 0 Then
            .vlCodigoAte = objRelacionamentoContatosCons.lCodigoAte
            iRetorno = Comando_BindVarInt(lComando, .vlCodigoAte)
            If (iRetorno <> AD_SQL_SUCESSO) Then gError 181123
        End If

        'Se a data inicial está preenchida
        If objRelacionamentoContatosCons.dtDataDe <> DATA_NULA Then
            .vdtDataDe = objRelacionamentoContatosCons.dtDataDe
            iRetorno = Comando_BindVarInt(lComando, .vdtDataDe)
            If (iRetorno <> AD_SQL_SUCESSO) Then gError 181124
        End If

        'Se a data final está preenchida
        If objRelacionamentoContatosCons.dtDataAte <> DATA_NULA Then
            .vdtDataAte = objRelacionamentoContatosCons.dtDataAte
            iRetorno = Comando_BindVarInt(lComando, .vdtDataAte)
            If (iRetorno <> AD_SQL_SUCESSO) Then gError 181125
        End If

        'Se o cliente está preenchido
        If objRelacionamentoContatosCons.lCliente > 0 Then
            .vlCliente = objRelacionamentoContatosCons.lCliente
            iRetorno = Comando_BindVarInt(lComando, .vlCliente)
            If (iRetorno <> AD_SQL_SUCESSO) Then gError 181126
        End If

        'Se a filial do cliente está preenchida
        If objRelacionamentoContatosCons.iFilialCliente > 0 Then
            .viFilialCliente = objRelacionamentoContatosCons.iFilialCliente
            iRetorno = Comando_BindVarInt(lComando, .viFilialCliente)
            If (iRetorno <> AD_SQL_SUCESSO) Then gError 181127
        End If

        'Se o atendente inicial do relacionamento foi informado
        If objRelacionamentoContatosCons.iAtendenteDe > 0 Then
            .viAtendenteDe = objRelacionamentoContatosCons.iAtendenteDe
            iRetorno = Comando_BindVarInt(lComando, .viAtendenteDe)
            If (iRetorno <> AD_SQL_SUCESSO) Then gError 181128
        End If

        'Se o atendente final do relacionamento foi informado
        If objRelacionamentoContatosCons.iAtendenteAte > 0 Then
            .viAtendenteAte = objRelacionamentoContatosCons.iAtendenteAte
            iRetorno = Comando_BindVarInt(lComando, .viAtendenteAte)
            If (iRetorno <> AD_SQL_SUCESSO) Then gError 181129
        End If

        'Se o tipo de relacionamento está preenchido
        If objRelacionamentoContatosCons.lTipo > 0 Then
            .vlTipo = objRelacionamentoContatosCons.lTipo
            iRetorno = Comando_BindVarInt(lComando, .vlTipo)
            If (iRetorno <> AD_SQL_SUCESSO) Then gError 181130
        End If

        'Se a origem do relacionamento está preenchida
        If objRelacionamentoContatosCons.iOrigem > 0 Then
            .viOrigem = objRelacionamentoContatosCons.iOrigem
            iRetorno = Comando_BindVarInt(lComando, .viOrigem)
            If (iRetorno <> AD_SQL_SUCESSO) Then gError 181131
        End If

        'Se o status do relacionamento está preenchido
        If objRelacionamentoContatosCons.iStatus <> -1 Then
            .viStatus = objRelacionamentoContatosCons.iStatus
            iRetorno = Comando_BindVarInt(lComando, .viStatus)
            If (iRetorno <> AD_SQL_SUCESSO) Then gError 181132
        End If
        
    End With

    RelacionamentoContatosCons_Le3 = SUCESSO

    Exit Function

Erro_RelacionamentoContatosCons_Le3:

    RelacionamentoContatosCons_Le3 = gErr

    Select Case gErr

        Case 181121 To 181132
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELACIONAMENTOCLIENTES", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 181133)

    End Select

    Exit Function

End Function

'Incluído por Luiz Nogueira em 27/10/03
Private Function RelacionamentoContatosCons_Le4(ByVal lComando As Long, tRelacionamentoContatosCons As typeRelacionamentoClientesCons, ByVal objRelacionamentoContatosCons As ClassRelacClientesCons) As Long

Dim lErro As Long
Dim objRelacionamentoContatos As ClassRelacClientes

On Error GoTo Erro_RelacionamentoContatosCons_Le4

    'Inicializa as strings para leitura de dados
    tRelacionamentoContatosCons.vsAssunto1 = String(STRING_BUFFER_MAX_TEXTO, 0)
    tRelacionamentoContatosCons.vsAssunto2 = String(STRING_BUFFER_MAX_TEXTO, 0)

    'Executa a leitura
    lErro = Comando_ExecutarInt(lComando)
    If (lErro <> AD_SQL_SUCESSO) Then gError 181094

    'Busca o primeiro registro encontrado
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 181095

    'Se não encontrou nenhum registro => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 181096

    'Enquanto houverem registros
    Do While lErro = SUCESSO

        'Instancia um novo obj
        Set objRelacionamentoContatos = New ClassRelacClientes

        'Transfere os dados lidos para o obj
        With objRelacionamentoContatos
            .lCodigo = tRelacionamentoContatosCons.vlCodigo
            .iFilialEmpresa = tRelacionamentoContatosCons.viFilialEmpresa
            .iOrigem = tRelacionamentoContatosCons.viOrigem
            .dtData = tRelacionamentoContatosCons.vdtData
            .dtHora = CDate(tRelacionamentoContatosCons.vdHora)
            .lTipo = tRelacionamentoContatosCons.vlTipo
            .lCliente = tRelacionamentoContatosCons.vlCliente
            .iFilialCliente = tRelacionamentoContatosCons.viFilialCliente
            .iContato = tRelacionamentoContatosCons.viContato
            .iAtendente = tRelacionamentoContatosCons.viAtendente
            .lRelacionamentoAnt = tRelacionamentoContatosCons.vlRelacionamentoAnt
            .sAssunto1 = tRelacionamentoContatosCons.vsAssunto1
            .sAssunto2 = tRelacionamentoContatosCons.vsAssunto2
            .iStatus = tRelacionamentoContatosCons.viStatus
        End With

        'Adiciona o registro à coleção
        objRelacionamentoContatosCons.colRelacionamentoClientes.Add objRelacionamentoContatos

        'Busca o próximo registro encontrado
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 181097

    Loop

    RelacionamentoContatosCons_Le4 = SUCESSO

    Exit Function

Erro_RelacionamentoContatosCons_Le4:

    RelacionamentoContatosCons_Le4 = gErr

    Select Case gErr

        Case 181094, 181095, 181097
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELACIONAMENTOCLIENTES", gErr)

        Case 181096

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 181098)

    End Select

    Exit Function

End Function

Function FiliaisContatos_Le_Contato(ByVal objContato As ClassContatos, ByVal colCodigoNome As AdmColCodigoNome) As Long
'Le na tabela FiliaisContatos todos os Codigos e Nomes de Filiais
'relacionadas ao objContato. Retorna na colecao colCodigoNome

Dim lComando As Long
Dim iCodFilial As Integer
Dim sNome As String
Dim lErro As Long
Dim lCodigoContato As Long

On Error GoTo Erro_FiliaisContatos_Le_Contato

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 181134

    sNome = String(STRING_FILIAL_CLIENTE_NOME, 0)
    
    lCodigoContato = objContato.lCodigo
    lErro = Comando_Executar(lComando, "SELECT CodFilial, Nome FROM FiliaisContatos WHERE CodContato=? AND CodFilial <> 0 ORDER BY CodFilial", iCodFilial, sNome, objContato.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 181134

    'le a primeira filial
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 181134
    If lErro = AD_SQL_SEM_DADOS Then gError 181135

    Do While lErro <> AD_SQL_SEM_DADOS

        'coloca a filial lida na coleção
        colCodigoNome.Add iCodFilial, sNome

        'le a proxima filial
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 181136

    Loop

    Call Comando_Fechar(lComando)

    FiliaisContatos_Le_Contato = SUCESSO

    Exit Function

Erro_FiliaisContatos_Le_Contato:

    FiliaisContatos_Le_Contato = gErr

    Select Case gErr

        Case 181134
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 181134, 181134, 181136
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISCONTATOS", gErr)

        Case 181135
            Call Rotina_Erro(vbOKOnly, "ERRO_Contato_SEM_FILIAL", gErr, lCodigoContato)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181137)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function RelacionamentoContatos_Le_Contato(ByVal objContato As ClassContatos, ByVal colRelacCli As Collection) As Long
'objRelacionamentoClientes RECEBE(Input) código/filialempresa do relacionamento que deve ser lido
'objRelacionamentoClientes RETORNA(Output) os dados do relacionamento lido

Dim lErro As Long
Dim lComando As Long
Dim tRelacionamentoClientes As typeRelacionamentoClientes
Dim objRelacionamentoClientes As ClassRelacClientes

On Error GoTo Erro_RelacionamentoContatos_Le_Contato

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 181028
    
    'Inicializa as strings para leitura de dados
    tRelacionamentoClientes.sAssunto1 = String(STRING_BUFFER_MAX_TEXTO, 0)
    tRelacionamentoClientes.sAssunto2 = String(STRING_BUFFER_MAX_TEXTO, 0)
    
    'Lê no BD o relacionamento com código e filialempresa passados como parâmetro
    lErro = Comando_Executar(lComando, "SELECT Codigo, FilialEmpresa, Origem, Data, Hora, Tipo, Cliente, FilialCliente, Contato, Atendente, RelacionamentoAnt, Assunto1, Assunto2, Status FROM RelacionamentoContatos WHERE Cliente = ? ", tRelacionamentoClientes.lCodigo, tRelacionamentoClientes.iFilialEmpresa, tRelacionamentoClientes.iOrigem, tRelacionamentoClientes.dtData, tRelacionamentoClientes.dHora, tRelacionamentoClientes.lTipo, tRelacionamentoClientes.lCliente, tRelacionamentoClientes.iFilialCliente, tRelacionamentoClientes.iContato, tRelacionamentoClientes.iAtendente, tRelacionamentoClientes.lRelacionamentoAnt, tRelacionamentoClientes.sAssunto1, tRelacionamentoClientes.sAssunto2, tRelacionamentoClientes.iStatus, objContato.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 181029
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 181030
    
    'Se não encontrou o relacionamento => erro
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objRelacionamentoClientes = New ClassRelacClientes
    
        'Transfere os dados lidos para o obj
        With objRelacionamentoClientes
        
            .lCodigo = tRelacionamentoClientes.lCodigo
            .iFilialEmpresa = tRelacionamentoClientes.iFilialEmpresa
            .iOrigem = tRelacionamentoClientes.iOrigem
            .dtData = tRelacionamentoClientes.dtData
            .dtHora = CDate(tRelacionamentoClientes.dHora)
            .lTipo = tRelacionamentoClientes.lTipo
            .lCliente = tRelacionamentoClientes.lCliente
            .iFilialCliente = tRelacionamentoClientes.iFilialCliente
            .iContato = tRelacionamentoClientes.iContato
            .iAtendente = tRelacionamentoClientes.iAtendente
            .lRelacionamentoAnt = tRelacionamentoClientes.lRelacionamentoAnt
            .sAssunto1 = tRelacionamentoClientes.sAssunto1
            .sAssunto2 = tRelacionamentoClientes.sAssunto2
            .iStatus = tRelacionamentoClientes.iStatus
        
        End With
        
        colRelacCli.Add objRelacionamentoClientes
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 181030
    
    Loop
    
    'Executa o fechamento do Comando
    Call Comando_Fechar(lComando)

    RelacionamentoContatos_Le_Contato = SUCESSO

    Exit Function

Erro_RelacionamentoContatos_Le_Contato:

    RelacionamentoContatos_Le_Contato = gErr

    Select Case gErr

        Case 181028
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 181029, 181030
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELACIONAMENTOCLIENTES", gErr)
            
        Case 181031 'Registro não encontrado
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181032)

    End Select

    'Executa o fechamento do Comando
    Call Comando_Fechar(lComando)

End Function

Public Function ClienteFContatos_Le_Nome(ByVal objClienteFContatos As ClassClienteContatos) As Long
'Lê um contato a partir de cliente/filial + nome do contato
'objContatoClientes recebe o código do cliente e da filial
'objContatoClientes retorna os dados lidos

Dim lErro As Long
Dim lComando As Long
Dim tClienteFContatos As typeClienteContatos

On Error GoTo Erro_ClienteFContatos_Le_Nome

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 178437
    
    'Inicializa as string que serão lidas
    tClienteFContatos.sContato = String(STRING_CONTATOGERAL_CONTATO, 0)
    tClienteFContatos.sSetor = String(STRING_CONTATOGERAL_SETOR, 0)
    tClienteFContatos.sCargo = String(STRING_CONTATOGERAL_CARGO, 0)
    tClienteFContatos.sTelefone = String(STRING_TELEFONE, 0)
    tClienteFContatos.sFax = String(STRING_FAX, 0)
    tClienteFContatos.sEmail = String(STRING_EMAIL, 0)
    
    'Lê o contato no BD
    lErro = Comando_Executar(lComando, "SELECT Codigo, Contato, Setor, Cargo, Telefone, Fax, Email, Padrao FROM ClienteFContatos WHERE Cliente=? AND FilialCliente=? AND Contato=?", tClienteFContatos.iCodigo, tClienteFContatos.sContato, tClienteFContatos.sSetor, tClienteFContatos.sCargo, tClienteFContatos.sTelefone, tClienteFContatos.sFax, tClienteFContatos.sEmail, tClienteFContatos.iPadrao, objClienteFContatos.lCliente, objClienteFContatos.iFilialCliente, objClienteFContatos.sContato)
    If lErro <> AD_SQL_SUCESSO Then gError 178438
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 178439
    
    'Se não encontrou o contato => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 178440
    
    'Move os dados lidos para o obj
    With objClienteFContatos
        .iCodigo = tClienteFContatos.iCodigo
        .sContato = tClienteFContatos.sContato
        .sSetor = tClienteFContatos.sSetor
        .sCargo = tClienteFContatos.sCargo
        .sTelefone = tClienteFContatos.sTelefone
        .sFax = tClienteFContatos.sFax
        .sEmail = tClienteFContatos.sEmail
        .iPadrao = tClienteFContatos.iPadrao
    End With
        
    'Executa o fechamento do Comando
    Call Comando_Fechar(lComando)

    ClienteFContatos_Le_Nome = SUCESSO

    Exit Function

Erro_ClienteFContatos_Le_Nome:

    ClienteFContatos_Le_Nome = gErr

    Select Case gErr

        Case 178437
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 178438, 178439
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ClienteFContatos", gErr)
        
        Case 178440
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178441)

    End Select

    'Executa o fechamento do Comando
    Call Comando_Fechar(lComando)

End Function
'######################################################################

Function ParcelasRecNF_Le_PrazoMedio(ByVal lNumIntNF As Long, dPrazoMedio As Double) As Long
'Lê as parcelas a receber vinculadas à Nota Fiscal

Dim lErro As Long
Dim lComando As Long
Dim tParcelaReceber As typeParcelaReceber
Dim dtDataReferencia As Date
Dim objNFiscal As New ClassNFiscal
Dim lNumParc As Long

On Error GoTo Erro_ParcelasRecNF_Le_PrazoMedio

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 124146
    
    'Busca as Parcelas a Receber vinculadas a Nfiscal
    lErro = Comando_Executar(lComando, "SELECT SUM(DATEDIFF(day, NFiscal.DataEmissao,ParcelasRec.DataVencimento))/Count(*), Count(*) FROM NFiscal, ParcelasRec WHERE NFiscal.NumIntDocCPR = ParcelasRec.NumIntTitulo AND NFiscal.NumIntDoc = ?", _
        dPrazoMedio, lNumParc, lNumIntNF)
    If lErro <> AD_SQL_SUCESSO Then gError 126147
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then gError 126148
    
    'Se não encontrou
    If lErro = AD_SQL_SEM_DADOS Or lNumParc = 0 Then
        
        'Busca as Parcelas a receber Baixadas vinculadas a Nota Fiscal
        lErro = Comando_Executar(lComando, "SELECT SUM(DATEDIFF(day, NFiscal.DataEmissao,ParcelasRecBaixadas.DataVencimento))/Count(*) FROM NFiscal, ParcelasRecBaixadas WHERE NFiscal.NumIntDocCPR = ParcelasRecBaixadas.NumIntTitulo AND NFiscal.NumIntDoc = ? ", _
            dPrazoMedio, lNumIntNF)
        If lErro <> AD_SQL_SUCESSO Then gError 126149
        
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 126150
                
    End If
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    ParcelasRecNF_Le_PrazoMedio = SUCESSO
    
    Exit Function
    
Erro_ParcelasRecNF_Le_PrazoMedio:

    ParcelasRecNF_Le_PrazoMedio = gErr

    Select Case gErr
    
        Case 126146
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 126147, 126148
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PARCELAS_REC_NF", gErr)
        
        Case 126149, 126150
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PARCELAS_REC_BAIXADAS_NF", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
     
    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function NFiscal_BuscaPrevisaoRec(ByVal objNFiscal As ClassNFiscal, ByVal objTitRecPrevisao As ClassTituloReceber) As Long
'preenche objTitRecPrevisao com numintdoc de titulo com STATUS_PREVISAO de mesmo cliente-filial de objNfiscal com mesmo mes de vencimento de parcela

Dim lErro As Long, lComando As Long, lNumIntDoc As Long, sSiglaDocumento As String

On Error GoTo Erro_NFiscal_BuscaPrevisaoRec

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 106976
    
    sSiglaDocumento = String(STRING_SIGLA_DOCUMENTO, 0)
    
    lErro = Comando_Executar(lComando, "SELECT TitulosRec.SiglaDocumento, TitulosRec.NumIntDoc FROM TitulosRec, ParcelasRec WHERE TitulosRec.FilialEmpresa = ? AND TitulosRec.Status = ? AND TitulosRec.NumIntDoc = ParcelasRec.NumIntTitulo AND Cliente = ? AND Filial = ? AND {fn Month(ParcelasRec.DataVencimentoReal)} = ?  AND {fn Year(ParcelasRec.DataVencimentoReal)} = ?", _
        sSiglaDocumento, lNumIntDoc, objNFiscal.iFilialEmpresa, STATUS_PREVISAO, objNFiscal.lCliente, objNFiscal.iFilialCli, Month(objNFiscal.colParcelaReceber(1).dtDataVencimentoReal), Year(objNFiscal.colParcelaReceber(1).dtDataVencimentoReal))
    If lErro <> AD_SQL_SUCESSO Then gError 106977

    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106798
    
    If lErro = AD_SQL_SEM_DADOS Then gError 106975
    
    objTitRecPrevisao.iFilialEmpresa = objNFiscal.iFilialEmpresa
    objTitRecPrevisao.lNumIntDoc = lNumIntDoc
    objTitRecPrevisao.sSiglaDocumento = sSiglaDocumento
    
    Call Comando_Fechar(lComando)
    
    NFiscal_BuscaPrevisaoRec = SUCESSO
     
    Exit Function
    
Erro_NFiscal_BuscaPrevisaoRec:

    NFiscal_BuscaPrevisaoRec = gErr
     
    Select Case gErr
          
        Case 106975
        
        Case 106976
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 106978
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PREVISAO_TITREC", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function ChequeParcelaRec_Le(ByVal objChequePre As ClassChequePre, ByVal lNumIntParcRec As Long, dValor As Double) As Long
'retorna o valor dos cheque em questao associado a parcela passada como parametro e que ainda nao foi depositado se houver

Dim lComando1 As Long
Dim lErro As Long

On Error GoTo Erro_ChequeParcelaRec_Le

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 188344

    lErro = Comando_Executar(lComando1, "SELECT CP.ValorRecebto-CP.Juros-CP.Multa+CP.Desconto FROM ChequeParcelaRec AS CP, ChequePre AS C WHERE C.Localizacao <> 2 AND CP.NumIntParcela = ? AND CP.NumIntCheque = C.NumIntCheque AND C.Banco = ? AND C.Agencia = ? AND C.ContaCorrente=? AND C.Numero=?", dValor, lNumIntParcRec, objChequePre.iBanco, objChequePre.sAgencia, objChequePre.sContaCorrente, objChequePre.lNumero)
    If lErro <> AD_SQL_SUCESSO Then gError 188345

    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 188346

    Call Comando_Fechar(lComando1)

    ChequeParcelaRec_Le = SUCESSO

Exit Function

Erro_ChequeParcelaRec_Le:

    ChequeParcelaRec_Le = gErr

    Select Case gErr

        Case 188344
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 188345, 188346
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CHEQUEPRE", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188343)

    End Select

    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Function Vendedor_Le_Customizado(ByVal objVendedor As ClassVendedor) As Long
    Vendedor_Le_Customizado = SUCESSO
End Function

Function FilialClienteCategorias_Le(objFilialCliCategoria As ClassFilialCliCategoria) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim lComando As Long
Dim sCategoria As String
Dim sItem As String

On Error GoTo Erro_FilialClienteCategorias_Le

    'Inicializar comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 199398

    sItem = String(STRING_CATEGORIACLIENTEITEM_ITEM, 0)

    'Pesquisa no BD as categoria da filial do cliente
    lErro = Comando_Executar(lComando, "SELECT Item FROM FilialClienteCategorias WHERE Cliente = ? AND Filial = ? AND Categoria = ? AND Item = ?", sItem, objFilialCliCategoria.lCliente, objFilialCliCategoria.iFilial, objFilialCliCategoria.sCategoria, objFilialCliCategoria.sItem)
    If lErro <> AD_SQL_SUCESSO Then gError 199399

    'Verifica se alguma categoria existe
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 199400

    If lErro = AD_SQL_SEM_DADOS Then gError 199401

    'Fechamento comando
    Call Comando_Fechar(lComando)

    FilialClienteCategorias_Le = SUCESSO

    Exit Function

Erro_FilialClienteCategorias_Le:

    FilialClienteCategorias_Le = gErr

    Select Case gErr

        Case 199398
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 199399, 199400
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIALCLIENTECATEGORIAS2", gErr, objFilialCliCategoria.lCliente)

        Case 199401

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 199402)

    End Select

    'Fechamento comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function DASAliquotas_Le(ByVal objDASAliquotas As ClassDASAliquotas) As Long

Dim lErro As Long
Dim lComando As Long
Dim tDASAliquotas As typeDASAliquotas, iFilialEmpresa As Integer

On Error GoTo Erro_DASAliquotas_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 200908

    If objDASAliquotas.iFilialEmpresa = 0 Then
        iFilialEmpresa = giFilialEmpresa
    Else
        iFilialEmpresa = objDASAliquotas.iFilialEmpresa
    End If

    'Le a tabelaDASAliquotas
    lErro = Comando_Executar(lComando, "SELECT  AliquotaICMSServ, AliquotaICMS, AliquotaTotal FROM DASAliquotas WHERE Ano= ?  AND Mes= ? AND FilialEmpresa = ? ", _
                tDASAliquotas.dAliquotaICMSServ, tDASAliquotas.dAliquotaICMS, tDASAliquotas.dAliquotaTotal, _
                objDASAliquotas.iAno, objDASAliquotas.iMes, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 200909

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200910

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError 200911

    objDASAliquotas.dAliquotaICMS = tDASAliquotas.dAliquotaICMS
    objDASAliquotas.dAliquotaICMSServ = tDASAliquotas.dAliquotaICMSServ
    objDASAliquotas.dAliquotaTotal = tDASAliquotas.dAliquotaTotal

    'Fecha Comando
    Call Comando_Fechar(lComando)

    DASAliquotas_Le = SUCESSO

    Exit Function

Erro_DASAliquotas_Le:

    DASAliquotas_Le = gErr

    Select Case gErr

        Case 200908
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 200909, 200910
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_DASALIQUOTAS", gErr)

        Case 200911 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200962)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function NFeFedProtNFE_Le(ByVal lNumIntNF As Long, sMotivo As String) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(1 To 4) As Long
Dim iEnviado As Integer
Dim iAmbiente As Integer, iStatus As Integer

On Error GoTo Erro_NFeFedProtNFE_Le

    'Executa a abertura do Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 203007
    Next
    
    'busca como cancelada com homologação na receita
    sMotivo = String(STRING_NFE_XMOTIVO, 0)
    lErro = Comando_Executar(alComando(1), "SELECT tpAmb, xMotivo FROM NFEFedRetCancNFe WHERE NumIntNF = ? AND cStat IN ('101','151','135','155') ORDER BY tpAmb", iAmbiente, sMotivo, lNumIntNF)
    If lErro <> AD_SQL_SUCESSO Then gError 203008
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203009

    If lErro = AD_SQL_SUCESSO Then
        If iAmbiente = 1 Then
            sMotivo = "H - " & sMotivo
        Else
            sMotivo = "P - " & sMotivo
        End If
    Else
        
        'buscca como denegada
        sMotivo = String(STRING_NFE_XMOTIVO, 0)
        lErro = Comando_Executar(alComando(2), "SELECT tpAmb, xMotivo FROM NFEFedProtNFe WHERE NumIntNF = ? AND (cStat = '205' OR cStat = '110' OR cStat = '301' OR CStat = '302') ORDER BY tpAmb", iAmbiente, sMotivo, lNumIntNF)
        If lErro <> AD_SQL_SUCESSO Then gError 203008
    
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203009
    
        If lErro = AD_SQL_SUCESSO Then
        
            If iAmbiente = 1 Then
                sMotivo = "H - " & sMotivo
            Else
                sMotivo = "P - " & sMotivo
            End If
        
        Else
        
            'pega status para ver se está cancelada sem homolgação na receita
            lErro = Comando_Executar(alComando(4), "SELECT Status FROM NFiscal WHERE NumIntDoc = ?", iStatus, lNumIntNF)
            If lErro <> AD_SQL_SUCESSO Then gError 203008
            
            lErro = Comando_BuscarPrimeiro(alComando(4))
            If lErro <> AD_SQL_SUCESSO Then gError 203009
            
            If iStatus = STATUS_CANCELADO Then
            
                sMotivo = "CANCELAMENTO NÃO HOMOLOGADO NA RECEITA"
            
            Else
            
                'busca como autorizada
                sMotivo = String(STRING_NFE_XMOTIVO, 0)
                lErro = Comando_Executar(alComando(3), "SELECT tpAmb, xMotivo FROM NFEFedProtNFe WHERE NumIntNF = ? AND cStat = '100' ORDER BY tpAmb", iAmbiente, sMotivo, lNumIntNF)
                If lErro <> AD_SQL_SUCESSO Then gError 203008
            
                lErro = Comando_BuscarPrimeiro(alComando(3))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203009
            
                If lErro = AD_SQL_SUCESSO Then
                    If iAmbiente = 1 Then
                        sMotivo = "H - " & sMotivo
                    Else
                        sMotivo = "P - " & sMotivo
                    End If
                Else
            
                    lErro = CF("NFeFedLote_Le", lNumIntNF, iEnviado)
                    If lErro <> SUCESSO Then gError 203054
                    
                    If iEnviado = 1 Then
                        sMotivo = "Enviado"
                    Else
                        sMotivo = ""
                    End If
                    
                End If
                
            End If
        
        End If
    
    End If

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    NFeFedProtNFE_Le = SUCESSO

    Exit Function

Erro_NFeFedProtNFE_Le:

    NFeFedProtNFE_Le = gErr

    Select Case gErr

        Case 203007
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 203008, 203009
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFEFEDPROTNFE", gErr)

        Case 203054

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 203011)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function NFeFedLote_Le(ByVal lNumIntNF As Long, iEnviado As Integer) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim lComando As Long
Dim lLote As Long

On Error GoTo Erro_NFeFedLote_Le

    'Inicializar comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 203050


    lErro = Comando_Executar(lComando, "SELECT Lote FROM NFEFedLote WHERE NumIntNF = ? ORDER BY DATA Desc, Hora DESC", lLote, lNumIntNF)
    If lErro <> AD_SQL_SUCESSO Then gError 203051

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203052

    If lErro = AD_SQL_SUCESSO Then
        iEnviado = 1
    Else
        iEnviado = 0
    End If


    'Fechamento comando
    Call Comando_Fechar(lComando)

    NFeFedLote_Le = SUCESSO

    Exit Function

Erro_NFeFedLote_Le:

    NFeFedLote_Le = gErr

    Select Case gErr

        Case 203050
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 203051, 203052
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFEFEDLOTE", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 203053)

    End Select

    'Fechamento comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function CamposGenericos_ValidaExclusao_CargoVendedor() As Long

Dim lComando As Long
Dim lErro As Long
Dim iVendedor As Integer
Dim iCargo As Integer

On Error GoTo Erro_CamposGenericos_ValidaExclusao_CargoVendedor
    
    'Abre comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 195987
    
    'Busca em relacionamentos com clientes, tipos de relacionamentos que não existam na tabela CamposGenericos Valores
    lErro = Comando_Executar(lComando, "SELECT Cargo, Codigo FROM Vendedores WHERE Cargo NOT IN (SELECT CodValor FROM CamposGenericosValores WHERE CodCampo = ?) AND Cargo <> 0", iCargo, iVendedor, CAMPOSGENERICOS_CARGO_VENDEDOR)
    If lErro <> AD_SQL_SUCESSO Then gError 195988
    
    'Posiciona no primeiro registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 195989
    
    'Se encontrou algum registro => erro de inconsistência
    If lErro = AD_SQL_SUCESSO Then gError 195990
    
    'Fecha comando
    Call Comando_Fechar(lComando)
    
    CamposGenericos_ValidaExclusao_CargoVendedor = SUCESSO
    
    Exit Function
    
Erro_CamposGenericos_ValidaExclusao_CargoVendedor:

    CamposGenericos_ValidaExclusao_CargoVendedor = gErr
    
    Select Case gErr
    
        Case 195987
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 195988, 195989
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_VENDEDORTRV_CAMPOSGENERICOS", gErr)
        
        Case 195990
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CARGO", gErr, iCargo, iVendedor)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195991)
    
    End Select

    'Fecha comando
    Call Comando_Fechar(lComando)

End Function

Public Function RPSWEBProt_Le(ByVal lNumIntNF As Long, sMotivo As String, sCodVerificacao As String, lNumNFe As Long) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim lComando As Long
Dim iEnviado As Integer
Dim iAmbiente As Integer
Dim sNumNFE As String

On Error GoTo Erro_RPSWEBProt_Le

    sCodVerificacao = ""
    lNumNFe = 0
    
    'Inicializar comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 207062
    
    lErro = Comando_Executar(lComando, "SELECT tpAmb FROM RPSWEBRetCanc WHERE NumIntNF = ? AND CodigoCancelamento <> 0 ", iAmbiente, lNumIntNF)
    If lErro <> AD_SQL_SUCESSO Then gError 207063
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 207064

    If lErro = AD_SQL_SUCESSO Then
        If iAmbiente = 2 Then
            sMotivo = "H - Cancelada"
        Else
            sMotivo = "P - Cancelada"
        End If
    Else
        
        sCodVerificacao = String(20, 0)
        sNumNFE = String(20, 0)
        lErro = Comando_Executar(lComando, "SELECT Ambiente, CodigoVerificacao, Numero FROM RPSWEBProt WHERE NumIntNF = ? ORDER BY Ambiente", iAmbiente, sCodVerificacao, sNumNFE, lNumIntNF)
        If lErro <> AD_SQL_SUCESSO Then gError 207065
    
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 207066
    
        If lErro = AD_SQL_SUCESSO Then
            If iAmbiente = 2 Then
                sMotivo = "H - Autorizado"
            Else
                sMotivo = "P - Autorizado"
            End If
            If IsNumeric(sNumNFE) Then
                If Len(Trim(sNumNFE)) > 9 Then
                    lNumNFe = StrParaLong(right(sNumNFE, 9))
                Else
                    lNumNFe = StrParaLong(sNumNFE)
                End If
            End If
            
        Else
            sCodVerificacao = ""
            lNumNFe = 0
            lErro = CF("RPSWEBLote_Le", lNumIntNF, iEnviado)
            If lErro <> SUCESSO Then gError 207067
            
            If iEnviado = 1 Then
                sMotivo = "Enviado"
            Else
                sMotivo = ""
            End If
            
        End If
        
    End If

    'Fechamento comando
    Call Comando_Fechar(lComando)

    RPSWEBProt_Le = SUCESSO

    Exit Function

Erro_RPSWEBProt_Le:

    RPSWEBProt_Le = gErr

    Select Case gErr

        Case 207062
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 207063, 207064
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RPSWEBRETCANC", gErr)

        Case 207065, 207066
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RPSWEBPROT", gErr)

        Case 207067

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207068)

    End Select

    'Fechamento comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function RPSWEBLote_Le(ByVal lNumIntNF As Long, iEnviado As Integer) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim lComando As Long
Dim lLote As Long

On Error GoTo Erro_RPSWEBLote_Le

    'Inicializar comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 207070

    lErro = Comando_Executar(lComando, "SELECT Lote FROM RPSWEBLote WHERE NumIntNF = ? ORDER BY DATA Desc, Hora DESC", lLote, lNumIntNF)
    If lErro <> AD_SQL_SUCESSO Then gError 207071

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 207072

    If lErro = AD_SQL_SUCESSO Then
        iEnviado = 1
    Else
        iEnviado = 0
    End If

    'Fechamento comando
    Call Comando_Fechar(lComando)

    RPSWEBLote_Le = SUCESSO

    Exit Function

Erro_RPSWEBLote_Le:

    RPSWEBLote_Le = gErr

    Select Case gErr

        Case 207070
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 207071, 207072
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RPSWEBLOTE", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207073)

    End Select

    'Fechamento comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function TipoFrete_Le(colTipoFrete As Collection) As Long
'Lê as parcelas a receber vinculadas à Nota Fiscal

Dim lErro As Long
Dim lComando As Long
Dim tTipoFrete As typeTipoFrete
Dim objTipoFrete As ClassTipoFrete

On Error GoTo Erro_TipoFrete_Le

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 207408
    
    Set colTipoFrete = New Collection
    
    tTipoFrete.sDescricao = String(STRING_TIPOFRETE_DESCRICAO, 0)
    
    'Busca as Parcelas a Receber vinculadas a Nfiscal
    lErro = Comando_Executar(lComando, "SELECT Codigo, Descricao, CodigoNFE FROM TipoFrete ", _
        tTipoFrete.iCodigo, tTipoFrete.sDescricao, tTipoFrete.iCodigoNFE)
    If lErro <> AD_SQL_SUCESSO Then gError 207409
    
     lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then gError 207410
    
    'Se não encontrou
    Do While lErro = AD_SQL_SUCESSO
        
        Set objTipoFrete = New ClassTipoFrete
        
        objTipoFrete.iCodigo = tTipoFrete.iCodigo
        objTipoFrete.sDescricao = tTipoFrete.sDescricao
        objTipoFrete.iCodigoNFE = tTipoFrete.iCodigoNFE
        
        colTipoFrete.Add objTipoFrete
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then gError 207411
        
    Loop
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    TipoFrete_Le = SUCESSO
    
    Exit Function
    
Erro_TipoFrete_Le:

    TipoFrete_Le = gErr

    Select Case gErr
    
        Case 207408
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 207409, 207410, 207411
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOFRETE", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207412)
     
    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function VersaoNFe_Le(ByVal objVersaoNFe As ClassVersaoNFe) As Long

Dim lErro As Long
Dim lComando As Long
Dim tVersaoNFe As typeVersaoNFe


On Error GoTo Erro_VersaoNFe_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 209173
    
    If objVersaoNFe.iCodigo = 0 Then objVersaoNFe.iCodigo = 1

    'Alocação de espaço no buffer
    tVersaoNFe.sDescricao = String(STRING_MAXIMO, 0)
    tVersaoNFe.sColunaCFOP = String(STRING_MAXIMO, 0)
    tVersaoNFe.sProgramaEnvio = String(STRING_MAXIMO, 0)

    'Le a tabelaVersaoNFe
    lErro = Comando_Executar(lComando, "SELECT Codigo, Descricao, ColunaCFOP, ProgramaEnvio, Ativo FROM VersaoNFe WHERE Codigo= ? ", _
                tVersaoNFe.iCodigo, tVersaoNFe.sDescricao, tVersaoNFe.sColunaCFOP, tVersaoNFe.sProgramaEnvio, tVersaoNFe.iAtivo, _
                objVersaoNFe.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 209174

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209175

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objVersaoNFe.iCodigo = tVersaoNFe.iCodigo
    objVersaoNFe.sDescricao = tVersaoNFe.sDescricao
    objVersaoNFe.sColunaCFOP = tVersaoNFe.sColunaCFOP
    objVersaoNFe.sProgramaEnvio = tVersaoNFe.sProgramaEnvio
    objVersaoNFe.iAtivo = tVersaoNFe.iAtivo

    'Fecha Comando
    Call Comando_Fechar(lComando)

    VersaoNFe_Le = SUCESSO

    Exit Function

Erro_VersaoNFe_Le:

    VersaoNFe_Le = gErr

    Select Case gErr

        Case 209173
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 209174, 209175
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_VERSAONFE", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209176)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function CFOP_Testa_Versao(ByVal sCFOP As String, ByVal sColuna As String, ByVal sVersao As String) As Long

Dim lErro As Long
Dim lComando As Long
Dim sDescricao As String

On Error GoTo Erro_CFOP_Testa_Versao


    sDescricao = String(STRING_NATUREZAOP_DESCRICAO, 0)

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 209184


    lErro = Comando_Executar(lComando, "SELECT Descricao FROM NaturezaOp WHERE Codigo = ? AND " & sColuna & " >= 0 ", sDescricao, sCFOP)
    If lErro <> AD_SQL_SUCESSO Then gError 209185

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209186

    If lErro = AD_SQL_SEM_DADOS Then gError 209187


    Call Comando_Fechar(lComando)

    CFOP_Testa_Versao = SUCESSO

    Exit Function

Erro_CFOP_Testa_Versao:

    CFOP_Testa_Versao = gErr

    Select Case gErr

        Case 209184
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 209185, 209186
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NATUREZAOP", gErr)

        Case 209187
            Call Rotina_Erro(vbOKOnly, "ERRO_TRIB_CFOP_INVALIDO_VERSAO", gErr, sCFOP, sVersao)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149092)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function RetiradaEntrega_Le(ByVal objRetiradaEntrega As ClassRetiradaEntrega) As Long

Dim lErro As Long
Dim lComando As Long
Dim tRetiradaEntrega As typeRetiradaEntrega
Dim objEnderecoRet As New ClassEndereco
Dim objEnderecoEnt As New ClassEndereco

On Error GoTo Erro_RetiradaEntrega_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 207722

    'Alocação de espaço no buffer
    tRetiradaEntrega.sCNPJCPFRet = String(STRING_CLIENTE_CGC, 0)
    tRetiradaEntrega.sCNPJCPFEnt = String(STRING_CLIENTE_CGC, 0)

    'Le a tabelaRetiradaEntrega
    lErro = Comando_Executar(lComando, "SELECT EnderecoRet, EnderecoEnt, CNPJCPFRet, CNPJCPFEnt, ClienteRet, " & _
                "FornecedorRet, FilialCliRet, FilialFornRet, ClienteEnt, FornecedorEnt, FilialCliEnt, FilialFornEnt FROM RetiradaEntrega WHERE TipoDoc = ? AND NumIntDoc = ?", _
                tRetiradaEntrega.lEnderecoRet, tRetiradaEntrega.lEnderecoEnt, tRetiradaEntrega.sCNPJCPFRet, _
                tRetiradaEntrega.sCNPJCPFEnt, tRetiradaEntrega.lClienteRet, tRetiradaEntrega.lFornecedorRet, tRetiradaEntrega.iFilialCliRet, tRetiradaEntrega.iFilialFornRet, _
                tRetiradaEntrega.lClienteEnt, tRetiradaEntrega.lFornecedorEnt, tRetiradaEntrega.iFilialCliEnt, tRetiradaEntrega.iFilialFornEnt, objRetiradaEntrega.iTipoDoc, objRetiradaEntrega.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 207723

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 207724

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objRetiradaEntrega.lEnderecoRet = tRetiradaEntrega.lEnderecoRet
    objRetiradaEntrega.lEnderecoEnt = tRetiradaEntrega.lEnderecoEnt
    objRetiradaEntrega.sCNPJCPFRet = tRetiradaEntrega.sCNPJCPFRet
    objRetiradaEntrega.sCNPJCPFEnt = tRetiradaEntrega.sCNPJCPFEnt
    objRetiradaEntrega.lClienteRet = tRetiradaEntrega.lClienteRet
    objRetiradaEntrega.lFornecedorRet = tRetiradaEntrega.lFornecedorRet
    objRetiradaEntrega.iFilialCliRet = tRetiradaEntrega.iFilialCliRet
    objRetiradaEntrega.iFilialFornRet = tRetiradaEntrega.iFilialFornRet
    objRetiradaEntrega.lClienteEnt = tRetiradaEntrega.lClienteEnt
    objRetiradaEntrega.lFornecedorEnt = tRetiradaEntrega.lFornecedorEnt
    objRetiradaEntrega.iFilialCliEnt = tRetiradaEntrega.iFilialCliEnt
    objRetiradaEntrega.iFilialFornEnt = tRetiradaEntrega.iFilialFornEnt

    objEnderecoRet.lCodigo = objRetiradaEntrega.lEnderecoRet

    'Endereco de Retirada
    lErro = CF("Endereco_Le", objEnderecoRet)
    If lErro <> SUCESSO And lErro <> 12309 Then gError 207749

    Set objRetiradaEntrega.objEnderecoRet = objEnderecoRet

    objEnderecoEnt.lCodigo = objRetiradaEntrega.lEnderecoEnt

    'Endereco de Retirada
    lErro = CF("Endereco_Le", objEnderecoEnt)
    If lErro <> SUCESSO And lErro <> 12309 Then gError 207750

    Set objRetiradaEntrega.objEnderecoEnt = objEnderecoEnt

    'Fecha Comando
    Call Comando_Fechar(lComando)

    RetiradaEntrega_Le = SUCESSO

    Exit Function

Erro_RetiradaEntrega_Le:

    RetiradaEntrega_Le = gErr

    Select Case gErr

        Case 207722
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 207723, 207724
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RETIRADAENTREGA", gErr)

        Case ERRO_LEITURA_SEM_DADOS
        
        Case 207749, 207750

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 207725)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function InfoAdicionais_Le(ByVal objInfoAdic As ClassInfoAdic) As Long

Dim lErro As Long
Dim lComando As Long, sNatureza As String, sCcl As String

On Error GoTo Erro_InfoAdicionais_Le

    If Not (objInfoAdic Is Nothing) Then
        
        'Executa a abertura do Comando
        lComando = Comando_Abrir()
        If lComando = 0 Then gError 209190

        sNatureza = String(STRING_NATMOVCTA_CODIGO, 0)
        sCcl = String(STRING_CCL, 0)
        
        'Le a tabela InfoAdicDoc
        lErro = Comando_Executar(lComando, "SELECT Natureza, Ccl FROM InfoAdicDoc WHERE TipoDoc= ?  AND NumIntDoc= ? ", _
                    sNatureza, sCcl, objInfoAdic.iTipoDoc, objInfoAdic.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 209191
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209192
    
        'Sem Dados
        If lErro = AD_SQL_SUCESSO Then
            objInfoAdic.sNatureza = sNatureza
            objInfoAdic.sCcl = sCcl
        Else
            objInfoAdic.sNatureza = ""
            objInfoAdic.sCcl = ""
        End If

        Set objInfoAdic.objRetEnt = New ClassRetiradaEntrega
        objInfoAdic.objRetEnt.iTipoDoc = objInfoAdic.iTipoDoc
        objInfoAdic.objRetEnt.lNumIntDoc = objInfoAdic.lNumIntDoc
        lErro = CF("RetiradaEntrega_Le", objInfoAdic.objRetEnt)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
        If lErro <> SUCESSO Then Set objInfoAdic.objRetEnt = Nothing
        
        Set objInfoAdic.objCompra = New ClassInfoAdicCompra
        objInfoAdic.objCompra.iTipoDoc = objInfoAdic.iTipoDoc
        objInfoAdic.objCompra.lNumIntDoc = objInfoAdic.lNumIntDoc
        lErro = InfoAdicCompra_Le(objInfoAdic.objCompra)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
        If lErro <> SUCESSO Then Set objInfoAdic.objCompra = Nothing
        
        Set objInfoAdic.objExportacao = New ClassInfoAdicExportacao
        objInfoAdic.objExportacao.iTipoDoc = objInfoAdic.iTipoDoc
        objInfoAdic.objExportacao.lNumIntDoc = objInfoAdic.lNumIntDoc
        lErro = InfoAdicExportacao_Le(objInfoAdic.objExportacao)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
        If lErro <> SUCESSO Then Set objInfoAdic.objExportacao = Nothing
        
        Set objInfoAdic.objAnexos = New ClassAnexos
        Call CF("InfoAdic_AjustaAnexo", objInfoAdic)
        lErro = CF("Anexos_Le", objInfoAdic.objAnexos)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        'Fecha Comando
        Call Comando_Fechar(lComando)

    End If
    
    InfoAdicionais_Le = SUCESSO

    Exit Function

Erro_InfoAdicionais_Le:

    InfoAdicionais_Le = gErr

    Select Case gErr
        
        Case 209190
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 209191, 209192
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_INFOADICDOC", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora
        
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209185)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Private Function InfoAdicExportacao_Le(ByVal objInfoAdicExportacao As ClassInfoAdicExportacao) As Long

Dim lErro As Long
Dim lComando As Long
Dim tInfoAdicExportacao As typeInfoAdicExportacao

On Error GoTo Erro_InfoAdicExportacao_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 209186

    'Alocação de espaço no buffer
    tInfoAdicExportacao.sUFEmbarque = String(STRING_ESTADO_SIGLA, 0)
    tInfoAdicExportacao.sLocalEmbarque = String(STRING_MAXIMO, 0)
    tInfoAdicExportacao.sNumRE = String(STRING_MAXIMO, 0)

    'Le a tabelaInfoAdicExportacao
    lErro = Comando_Executar(lComando, "SELECT UFEmbarque, LocalEmbarque, NumIntDE, NumRE FROM InfoAdicExportacao WHERE TipoDoc= ? AND NumIntDoc= ? ", _
                tInfoAdicExportacao.sUFEmbarque, tInfoAdicExportacao.sLocalEmbarque, tInfoAdicExportacao.lNumIntDE, tInfoAdicExportacao.sNumRE, _
                objInfoAdicExportacao.iTipoDoc, objInfoAdicExportacao.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 209187

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209188

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objInfoAdicExportacao.sUFEmbarque = tInfoAdicExportacao.sUFEmbarque
    objInfoAdicExportacao.sLocalEmbarque = tInfoAdicExportacao.sLocalEmbarque
    objInfoAdicExportacao.lNumIntDE = tInfoAdicExportacao.lNumIntDE
    objInfoAdicExportacao.sNumRE = tInfoAdicExportacao.sNumRE

    'Fecha Comando
    Call Comando_Fechar(lComando)

    InfoAdicExportacao_Le = SUCESSO

    Exit Function

Erro_InfoAdicExportacao_Le:

    InfoAdicExportacao_Le = gErr

    Select Case gErr

        Case 209186
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 209187, 209188
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_INFOADICEXPORTACAO", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209189)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function InfoAdicCompra_Le(ByVal objInfoAdicCompra As ClassInfoAdicCompra) As Long

Dim lErro As Long
Dim lComando As Long
Dim tInfoAdicCompra As typeInfoAdicCompra

On Error GoTo Erro_InfoAdicCompra_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 209190

    'Alocação de espaço no buffer
    tInfoAdicCompra.sNotaEmpenho = String(STRING_MAXIMO, 0)
    tInfoAdicCompra.sPedido = String(STRING_MAXIMO, 0)
    tInfoAdicCompra.sContrato = String(STRING_MAXIMO, 0)

    'Le a tabelaInfoAdicCompra
    lErro = Comando_Executar(lComando, "SELECT NotaEmpenho, Pedido, Contrato FROM InfoAdicCompra WHERE TipoDoc= ?  AND NumIntDoc= ? ", _
                tInfoAdicCompra.sNotaEmpenho, tInfoAdicCompra.sPedido, tInfoAdicCompra.sContrato, _
                objInfoAdicCompra.iTipoDoc, objInfoAdicCompra.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 209191

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209192

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objInfoAdicCompra.sNotaEmpenho = tInfoAdicCompra.sNotaEmpenho
    objInfoAdicCompra.sPedido = tInfoAdicCompra.sPedido
    objInfoAdicCompra.sContrato = tInfoAdicCompra.sContrato

    'Fecha Comando
    Call Comando_Fechar(lComando)

    InfoAdicCompra_Le = SUCESSO

    Exit Function

Erro_InfoAdicCompra_Le:

    InfoAdicCompra_Le = gErr

    Select Case gErr

        Case 209190
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 209191, 209192
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_INFOADICCOMPRA", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209193)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function NFeFedProtNFe_Le1(ByVal lNumIntNF As Long, sStat As String, dtData As Date) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim lComando As Long
Dim iEnviado As Integer
Dim iAmbiente As Integer
Dim sMotivo As String

On Error GoTo Erro_NFeFedProtNFe_Le1

    'Inicializar comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 207556
    
    sMotivo = String(STRING_NFE_XMOTIVO, 0)

    lErro = Comando_Executar(lComando, "SELECT tpAmb, xMotivo, data FROM NFEFedRetCancNFe WHERE NumIntNF = ? AND cStat IN ('101','151','135','155') ", iAmbiente, sMotivo, dtData, lNumIntNF)
    If lErro <> AD_SQL_SUCESSO Then gError 207557
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 207558

    If lErro = AD_SQL_SUCESSO Then
        sStat = "101"
        If iAmbiente = 1 Then
            sMotivo = "H - " & sMotivo
        Else
            sMotivo = "P - " & sMotivo
        End If
    Else
        
        sMotivo = String(STRING_NFE_XMOTIVO, 0)
        sStat = String(STRING_NFE_XMOTIVO, 0)
    
        lErro = Comando_Executar(lComando, "SELECT tpAmb, xMotivo, cStat, Data FROM NFEFedProtNFe WHERE NumIntNF = ? AND (cStat = '100' OR cStat = '205' OR cStat = '110' OR cStat = '301' OR CStat = '302') ", iAmbiente, sMotivo, sStat, dtData, lNumIntNF)
        If lErro <> AD_SQL_SUCESSO Then gError 207559
    
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 207560
    
        If lErro = AD_SQL_SUCESSO Then
            If iAmbiente = 1 Then
                sMotivo = "H - " & sMotivo
            Else
                sMotivo = "P - " & sMotivo
            End If
        Else
    
            lErro = CF("NFeFedLote_Le", lNumIntNF, iEnviado)
            If lErro <> SUCESSO Then gError 207561
            
            If iEnviado = 1 Then
                sStat = "Enviado"
                sMotivo = "Enviado"
            Else
                sMotivo = ""
            End If
            
        End If
        
    End If

    'Fechamento comando
    Call Comando_Fechar(lComando)

    NFeFedProtNFe_Le1 = SUCESSO

    Exit Function

Erro_NFeFedProtNFe_Le1:

    NFeFedProtNFe_Le1 = gErr

    Select Case gErr

        Case 207556
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 207557, 207558
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFEFEDRETCANCNFE", gErr)

        Case 207559, 207560
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFEFEDPROTNFE", gErr)

        Case 207561

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207562)

    End Select

    'Fechamento comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function NFiscal_Le_NumNotaFiscal(ByVal objNFiscal As ClassNFiscal) As Long
'verifical se a nota fiscal já está cadastrada ou se já existe uma nota com os mesmos dados em um dado periodo

Dim lErro As Long
Dim lNumIntDoc As Long
Dim dtDataEmissao As Date
Dim vbMsgRes As VbMsgBoxResult
Dim sStat As String
Dim alComando(1 To 2) As Long
Dim iIndice As Integer


On Error GoTo Erro_NFiscal_Le_NumNotaFiscal

    'Executa a abertura do Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 207672
    Next

    'Verifica se já existe nota com esses dados
    lErro = Comando_Executar(alComando(1), "SELECT NumIntDoc FROM TipoDocInfoNFiscal WHERE FilialEmpresa = ? AND Serie = ? AND NumNotaFiscal = ? AND DataEmissao = ? AND Status <> ? AND (Tipo = ? OR Tipo = ?) ", lNumIntDoc, objNFiscal.iFilialEmpresa, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.dtDataEmissao, STATUS_EXCLUIDO, DOCINFO_NF_INT_ENTRADA, DOCINFO_NF_INT_SAIDA)
    If lErro <> AD_SQL_SUCESSO Then gError 207673
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 207674
    
    'Existe
    If lErro = AD_SQL_SUCESSO Then
    
        objNFiscal.lNumIntDoc = lNumIntDoc
        
    Else
    
        'Verifica em um período de emissão se exite alguma nota com os mesmos dados
        lErro = Comando_Executar(alComando(2), "SELECT DataEmissao, NumIntDoc FROM TipoDocInfoNFiscal WHERE FilialEmpresa = ? AND Serie = ? AND NumNotaFiscal = ? AND DataEmissao > ? AND DataEmissao < ? AND Status <> ? AND (Tipo = ? OR Tipo = ?)", dtDataEmissao, lNumIntDoc, objNFiscal.iFilialEmpresa, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.dtDataEmissao - PERIODO_EMISSAO, objNFiscal.dtDataEmissao + PERIODO_EMISSAO, STATUS_EXCLUIDO, DOCINFO_NF_INT_ENTRADA, DOCINFO_NF_INT_SAIDA)
        If lErro <> AD_SQL_SUCESSO Then gError 207675
        
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 207676
        
        If lErro <> AD_SQL_SUCESSO Then gError 207677
            
        objNFiscal.lNumIntDoc = lNumIntDoc
            
    End If

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    NFiscal_Le_NumNotaFiscal = SUCESSO

    Exit Function
    
Erro_NFiscal_Le_NumNotaFiscal:

    NFiscal_Le_NumNotaFiscal = gErr
    
    Select Case gErr
    
        Case 207672
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 207673 To 207676
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NOTA_FISCAL", gErr, objNFiscal.iTipoNFiscal, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal)
        
        Case 207677
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207678)
            
    End Select
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Public Function RelacCliTexto_Le(ByVal objRelacCli As ClassRelacClientes) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer, sTexto As String
Dim sAssuntoCompleto As String

On Error GoTo Erro_RelacCliTexto_Le

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 209822
    Next
    
    sTexto = String(STRING_MAXIMO, 0)
    
    lErro = Comando_Executar(alComando(0), "SELECT Texto FROM RelacCliTexto WHERE Codigo = ? AND FilialEmpresa = ? ORDER BY Seq", sTexto, objRelacCli.lCodigo, objRelacCli.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 209823
    
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209824
   
    sAssuntoCompleto = ""
    Do While lErro <> AD_SQL_SEM_DADOS
    
        sAssuntoCompleto = sAssuntoCompleto & sTexto
    
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209825
    
    Loop
    
    If Len(Trim(sAssuntoCompleto)) > 0 Then
        objRelacCli.sAssunto1 = left(sAssuntoCompleto, STRING_BUFFER_MAX_TEXTO - 1)
        objRelacCli.sAssunto2 = Mid(sAssuntoCompleto, STRING_BUFFER_MAX_TEXTO)
    End If

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    RelacCliTexto_Le = SUCESSO

    Exit Function

Erro_RelacCliTexto_Le:

    RelacCliTexto_Le = gErr

    Select Case gErr

        Case 209822
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 209823 To 209825
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELACIONAMENTOCLIENTES", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209826)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function
    
End Function

Public Function RelacCli_Le_TipoDoc(ByVal objRelacCli As ClassRelacClientes) As Long

Dim lErro As Long
Dim lComando As Long
Dim tRelaCli As typeRelacionamentoClientes

On Error GoTo Erro_RelacCli_Le_TipoDoc

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 211226
    
    'Lê no BD o relacionamento com código e filialempresa passados como parâmetro
    lErro = Comando_Executar(lComando, "SELECT Codigo, FilialEmpresa FROM RelacionamentoClientes WHERE TipoDoc = ? AND NumIntDocOrigem = ? ORDER BY Data, Codigo", _
    tRelaCli.lCodigo, tRelaCli.iFilialEmpresa, objRelacCli.iTipoDoc, objRelacCli.lNumIntDocOrigem)
    If lErro <> AD_SQL_SUCESSO Then gError 211227
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211228
    
    'Se não encontrou o relacionamento => erro
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS
       
    objRelacCli.lCodigo = tRelaCli.lCodigo
    objRelacCli.iFilialEmpresa = tRelaCli.iFilialEmpresa
    
    lErro = CF("RelacionamentoClientes_Le", objRelacCli)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Executa o fechamento do Comando
    Call Comando_Fechar(lComando)

    RelacCli_Le_TipoDoc = SUCESSO

    Exit Function

Erro_RelacCli_Le_TipoDoc:

    RelacCli_Le_TipoDoc = gErr

    Select Case gErr

        Case 211226
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 211227, 211228
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELACIONAMENTOCLIENTES", gErr)
            
        Case ERRO_LEITURA_SEM_DADOS 'Registro não encontrado
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211229)

    End Select

    'Executa o fechamento do Comando
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function RelacCli_Le_TipoDoc2(ByVal iTipoDoc As Integer, ByVal lNumIntDocOrigem As Long, ByVal colRelacCli As Collection) As Long

Dim lErro As Long
Dim lComando As Long
Dim tRelaCli As typeRelacionamentoClientes
Dim objRelacCli As ClassRelacClientes

On Error GoTo Erro_RelacCli_Le_TipoDoc2

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 211226
    
    tRelaCli.sAssunto1 = String(STRING_MAXIMO, 0)
    tRelaCli.sAssunto2 = String(STRING_MAXIMO, 0)
    
    'Lê no BD o relacionamento com código e filialempresa passados como parâmetro
    lErro = Comando_Executar(lComando, "SELECT Codigo, FilialEmpresa, Data, Hora, Assunto1, Assunto2, DataProxCobr, DataPrevReceb FROM RelacionamentoClientes WHERE TipoDoc = ? AND NumIntDocOrigem = ? ORDER BY Data DESC, Codigo DESC", _
    tRelaCli.lCodigo, tRelaCli.iFilialEmpresa, tRelaCli.dtData, tRelaCli.dHora, tRelaCli.sAssunto1, tRelaCli.sAssunto2, tRelaCli.dtDataProxCobr, tRelaCli.dtDataPrevReceb, iTipoDoc, lNumIntDocOrigem)
    If lErro <> AD_SQL_SUCESSO Then gError 211227
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211228
    
    'Se não encontrou o relacionamento => erro
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objRelacCli = New ClassRelacClientes
       
        objRelacCli.lCodigo = tRelaCli.lCodigo
        objRelacCli.iFilialEmpresa = tRelaCli.iFilialEmpresa
        objRelacCli.dtData = tRelaCli.dtData
        objRelacCli.dtHora = tRelaCli.dHora
        objRelacCli.sAssunto1 = tRelaCli.sAssunto1
        objRelacCli.sAssunto2 = tRelaCli.sAssunto2
        objRelacCli.dtDataProxCobr = tRelaCli.dtDataProxCobr
        objRelacCli.dtDataPrevReceb = tRelaCli.dtDataPrevReceb
        
        colRelacCli.Add objRelacCli
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211228
        
    Loop

    'Executa o fechamento do Comando
    Call Comando_Fechar(lComando)

    RelacCli_Le_TipoDoc2 = SUCESSO

    Exit Function

Erro_RelacCli_Le_TipoDoc2:

    RelacCli_Le_TipoDoc2 = gErr

    Select Case gErr

        Case 211226
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 211227, 211228
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELACIONAMENTOCLIENTES", gErr)
            
        Case ERRO_LEITURA_SEM_DADOS 'Registro não encontrado
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211229)

    End Select

    'Executa o fechamento do Comando
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function NFiscal_Le_ChvNFe(ByVal objNF As ClassNFiscal) As Long

Dim lErro As Long, iIndice As Integer
Dim alComando(1 To 2) As Long
Dim sChvNFe1 As String, sChvNFe2 As String, sQRCode As String

On Error GoTo Erro_NFiscal_Le_ChvNFe

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 211254
    Next
    
    objNF.sChvNFe = ""
    
    sChvNFe1 = String(STRING_NFE_CHNFE, 0)
    sQRCode = String(400, 0)
    
    lErro = Comando_Executar(alComando(1), "SELECT chNFe, QRCode FROM NFEFedProtNFe WHERE NumIntNF = ? AND cStat = '100' ", sChvNFe1, sQRCode, objNF.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 211255

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211256

    If lErro = AD_SQL_SUCESSO Then
        objNF.sChvNFe = sChvNFe1
        objNF.sQRCode = sQRCode
    Else
    
        sChvNFe2 = String(STRING_NFE_CHNFE, 0)
    
        lErro = Comando_Executar(alComando(2), "SELECT chvNFe FROM NFiscal WHERE NumIntDoc = ? ", sChvNFe2, objNF.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 211257
    
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211258
    
        If lErro = AD_SQL_SUCESSO Then objNF.sChvNFe = sChvNFe2
    
    End If
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    NFiscal_Le_ChvNFe = SUCESSO

    Exit Function

Erro_NFiscal_Le_ChvNFe:

    NFiscal_Le_ChvNFe = gErr

    Select Case gErr

        Case 211254
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 211255 To 211258
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", gErr)
                   
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211259)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Public Function NFiscal_Le_Pela_ChvNFe(ByVal objNF As ClassNFiscal, Optional ByVal lNumIntNFDesconsiderar As Long = 0) As Long

Dim lErro As Long, iIndice As Integer
Dim alComando(1 To 2) As Long
Dim lNumIntDoc As Long, iFilialEmpresa As Integer

On Error GoTo Erro_NFiscal_Le_Pela_ChvNFe

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 211260
    Next
    
    objNF.lNumIntDoc = 0
    
    iFilialEmpresa = objNF.iFilialEmpresa
    If iFilialEmpresa = 0 Then iFilialEmpresa = giFilialEmpresa
    
    lErro = Comando_Executar(alComando(1), "SELECT NumIntNF FROM NFEFedProtNFe WHERE chNFe = ? AND cStat = '100' AND NumIntNF <> ? AND FilialEmpresa = ? ", lNumIntDoc, objNF.sChvNFe, lNumIntNFDesconsiderar, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 211261

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211262

    If lErro = AD_SQL_SUCESSO Then
        objNF.lNumIntDoc = lNumIntDoc
    Else
    
        lErro = Comando_Executar(alComando(2), "SELECT NumIntDoc FROM NFiscal WHERE chvNFe = ? AND NumIntDoc <> ? AND FilialEmpresa = ? ", lNumIntDoc, objNF.sChvNFe, lNumIntNFDesconsiderar, iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 211263
    
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211264
    
        If lErro <> AD_SQL_SUCESSO Then gError ERRO_LEITURA_SEM_DADOS
        
        objNF.lNumIntDoc = lNumIntDoc
    
    End If
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    NFiscal_Le_Pela_ChvNFe = SUCESSO

    Exit Function

Erro_NFiscal_Le_Pela_ChvNFe:

    NFiscal_Le_Pela_ChvNFe = gErr

    Select Case gErr

        Case 211260
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 211261 To 211264
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", gErr)
            
        Case ERRO_LEITURA_SEM_DADOS
                   
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211265)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Public Function ImportNFeXml_Le_Pela_ChvNFe(ByVal objNF As ClassNFiscal) As Long

Dim lErro As Long, iIndice As Integer
Dim alComando(1 To 1) As Long, iFilialEmpresa As Integer
Dim lNumNF As Long, sSerie As String

On Error GoTo Erro_ImportNFeXml_Le_Pela_ChvNFe

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 211268
    Next
    
    sSerie = String(STRING_MAXIMO, 0)
    
    iFilialEmpresa = objNF.iFilialEmpresa
    If iFilialEmpresa = 0 Then iFilialEmpresa = giFilialEmpresa
    
    lErro = Comando_Executar(alComando(1), "SELECT Serie, NumNF FROM ImportNFeXml WHERE chvNFe = ? AND FilialEmpresa = ? ", sSerie, lNumNF, objNF.sChvNFe, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 211269

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211270
    
    If lErro <> AD_SQL_SUCESSO Then gError ERRO_LEITURA_SEM_DADOS
    
    objNF.lNumNotaFiscal = lNumNF
    objNF.sSerie = sSerie
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    ImportNFeXml_Le_Pela_ChvNFe = SUCESSO

    Exit Function

Erro_ImportNFeXml_Le_Pela_ChvNFe:

    ImportNFeXml_Le_Pela_ChvNFe = gErr

    Select Case gErr

        Case 211268
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 211269 To 211270
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", gErr)
            
        Case ERRO_LEITURA_SEM_DADOS
                   
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211271)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Public Function Vendedor_Le_Hierarquia(ByVal iVendedor As Integer, ByVal colVendedores As Collection, ByVal iNivel As Integer) As Long

Dim lErro As Long
Dim lComando As Long
Dim iCodigo As Integer, bAchou As Boolean
Dim objVendedor As ClassVendedor

On Error GoTo Erro_Vendedor_Le_Hierarquia

    iNivel = iNivel + 1
    
    If iNivel > 20 Then gError 211580
    
    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 211226

    'Seleciona todos os Vendedores Ativos
    lErro = Comando_Executar(lComando, "SELECT Codigo FROM Vendedores WHERE Superior = ? ", iCodigo, iVendedor)
    If lErro <> AD_SQL_SUCESSO Then gError 211576

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211577
    
    Do While lErro = AD_SQL_SUCESSO
    
        bAchou = False
        For Each objVendedor In colVendedores
            If objVendedor.iCodigo = iVendedor Then
                bAchou = True
                Exit For
            End If
        Next
        If Not bAchou Then
            Set objVendedor = New ClassVendedor
            objVendedor.iCodigo = iCodigo
            colVendedores.Add objVendedor
        End If
        
        lErro = Vendedor_Le_Hierarquia(iCodigo, colVendedores, iNivel)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211578
    
    Loop
    
    'Executa o fechamento do Comando
    Call Comando_Fechar(lComando)

    Vendedor_Le_Hierarquia = SUCESSO

    Exit Function

Erro_Vendedor_Le_Hierarquia:

    Vendedor_Le_Hierarquia = gErr

    Select Case gErr
    
        Case 211576 To 211578
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_VENDEDORES", gErr)
            
        Case 211580 'ERRO_LIMITE_RECURSIVIDADE
            Call Rotina_Erro(vbOKOnly, "ERRO_LIMITE_RECURSIVIDADE", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211579)

    End Select
    
    'Executa o fechamento do Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function Bordero_Le_NumNF_Serie(ByVal lNumBordero As Integer, ByVal sSerie As String, lNumNFDe As Long, lNumNFAte As Long) As Long

Dim lErro As Long
Dim lComando As Long
Dim sSQL As String

On Error GoTo Erro_Bordero_Le_NumNF_Serie
    
    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 211692

    sSQL = "SELECT MIN(N.NumNotaFiscal), MAX(N.NumNotaFiscal) FROM ParcelasRecTodas AS P, TitulosRecTodos AS T, NFiscal As N, OcorrenciasRemParcRec AS O"
    sSQL = sSQL & " " & "WHERE N.ClasseDocCPR = 2 AND T.NumIntDoc = N.NumIntDocCPR AND T.NumIntDoc = P.NumIntTitulo AND"
    sSQL = sSQL & " " & "T.Status <> 5 AND O.CodOcorrencia = 1 AND O.NumIntParc = P.NumIntDoc AND O.NumBordero = ? AND N.Serie = ?"
        
    lErro = Comando_Executar(lComando, sSQL, lNumNFDe, lNumNFAte, lNumBordero, sSerie)
    If lErro <> AD_SQL_SUCESSO Then gError 211693

    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211694
    
    'Executa o fechamento do Comando
    Call Comando_Fechar(lComando)

    Bordero_Le_NumNF_Serie = SUCESSO

    Exit Function

Erro_Bordero_Le_NumNF_Serie:

    Bordero_Le_NumNF_Serie = gErr

    Select Case gErr
    
        Case 211692
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 211693 To 211694
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", gErr)
  
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211695)

    End Select
    
    'Executa o fechamento do Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function ParcelasRec_Le_DataValor_NFe(ByVal schNFe As String, ByVal lNumIntTitulo As Long, ByVal colParcRec As colParcelaReceber, lComando1 As Long, lComando2 As Long) As Long
'Lê as datas de vencimento e os valores das ParcelasRec associadas ao título com a chave passada. Se nao achar pega pelo lNumIntTitulo passado.
'IMPORTANTE --- Esta função é usada para relatório.
'IMPORTANTE --- Caso lComando1 ou 2 venham zerados seráo abertos e permaneceráo abertos.

Dim lErro As Long
Dim tParcelaReceber As typeParcelaReceber

On Error GoTo Erro_ParcelasRec_Le_DataValor_NFe
    
    'Abre o Comando passado
    If lComando1 = 0 Then
        lComando1 = Comando_AbrirExt(GL_lConexaoRel)
        If lComando1 = 0 Then gError 58225
    End If
    
    'Abre o Comando passado
    If lComando2 = 0 Then
        lComando2 = Comando_AbrirExt(GL_lConexaoRel)
        If lComando2 = 0 Then gError 58225
    End If
    
    'Le a Data e o Valor da Parcela a Receber
    lErro = Comando_Executar(lComando1, "SELECT NumParcela, DataVencimento, Valor FROM DanfeCobranca WHERE chNFe = ? ORDER BY NumParcela", tParcelaReceber.iNumParcela, tParcelaReceber.dtDataVencimento, tParcelaReceber.dValor, schNFe)
    If lErro <> AD_SQL_SUCESSO Then gError 58226
    
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 58227
    
    If lErro = AD_SQL_SUCESSO Then
    
        Do While lErro = AD_SQL_SUCESSO
        
            With tParcelaReceber
                colParcRec.Add .lNumIntDoc, .lNumIntTitulo, .iNumParcela, 0, .dtDataVencimento, DATA_NULA, .dValor, .dValor, 0, 0, 0, "", 0, 0, 0, 0, 0, 0, 0, DATA_NULA, 0, 0, 0, 0, 0, DATA_NULA, 0, 0, 0, 0, 0, .iPrevisao, .sObservacao, .dValor, CStr(.iNumParcela)
            End With
        
            lErro = Comando_BuscarProximo(lComando1)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 58228
            
        Loop
    
    Else
    
        'Le a Data e o Valor da Parcela a Receber
        lErro = Comando_Executar(lComando2, "SELECT NumParcela, DataVencimento, Valor FROM ParcelasRecTodas WHERE NumIntTitulo = ? ORDER BY NumParcela", tParcelaReceber.iNumParcela, tParcelaReceber.dtDataVencimento, tParcelaReceber.dValor, lNumIntTitulo)
        If lErro <> AD_SQL_SUCESSO Then gError 58226
        
        lErro = Comando_BuscarPrimeiro(lComando2)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 58227
        
        Do While lErro = AD_SQL_SUCESSO
        
            With tParcelaReceber
                colParcRec.Add .lNumIntDoc, .lNumIntTitulo, .iNumParcela, 0, .dtDataVencimento, DATA_NULA, .dValor, .dValor, 0, 0, 0, "", 0, 0, 0, 0, 0, 0, 0, DATA_NULA, 0, 0, 0, 0, 0, DATA_NULA, 0, 0, 0, 0, 0, .iPrevisao, .sObservacao, .dValor, CStr(.iNumParcela)
            End With
        
            lErro = Comando_BuscarProximo(lComando2)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 58228
            
        Loop
    
    End If
    
    ParcelasRec_Le_DataValor_NFe = SUCESSO
    
    Exit Function
    
Erro_ParcelasRec_Le_DataValor_NFe:

    ParcelasRec_Le_DataValor_NFe = gErr
    
    Select Case gErr
    
        Case 58225
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 58226, 58227, 58228
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PARCELAS_REC", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149163)
    
    End Select
    
    Exit Function

End Function

Public Function NFe_Le_Faixa(ByVal iFilialEmpresa As Integer, ByVal sSerie As String, ByVal lNumNFDe As Long, ByVal lNumNFAte As Long, ByVal colNF As Collection) As Long

Dim lErro As Long
Dim alComando(1 To 1) As Long
Dim iIndice As Integer
Dim objNF As ClassNFiscal
Dim lNumNF As Long, lNumIntNF As Long
Dim dtDataEmissao As Date

On Error GoTo Erro_NFe_Le_Faixa
    
    'abrir comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 211722
    Next
    
    'Verifica se está vinculada a uma NFe autorizada
    lErro = Comando_Executar(alComando(1), "SELECT DISTINCT NFNumIntDoc, NFNumNotaFiscal, NFDataEmissao FROM Danfe WHERE NFFilialEmpresa = ? AND NFSerie = ? AND NFNumNotaFiscal BETWEEN ? AND ? ORDER BY NFNumNotaFiscal", lNumIntNF, lNumNF, dtDataEmissao, iFilialEmpresa, sSerie, lNumNFDe, lNumNFAte)
    If lErro <> AD_SQL_SUCESSO Then gError 211723
    
    'Lê o primeiro registro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211724
    
    'Se encontrou o registro é porque a NFe já foi autorizada, então grava o histórico se não existir
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objNF = New ClassNFiscal
        colNF.Add objNF
        
        objNF.lNumIntDoc = lNumIntNF
        objNF.lNumNotaFiscal = lNumNF
        objNF.sSerie = sSerie
        objNF.dtDataEmissao = dtDataEmissao
        objNF.iFilialEmpresa = iFilialEmpresa
        
        'Lê o primeiro registro
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211724
        
    Loop
    
   'fechar comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
          
    NFe_Le_Faixa = SUCESSO
    
    Exit Function
    
Erro_NFe_Le_Faixa:
    
    NFe_Le_Faixa = gErr
    
    Select Case gErr
        
        Case 211722
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 211723, 211724
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", gErr)
        
        Case 211725, 211726
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PARCELAS_REC", gErr)
     
        Case 211727
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PARCELAS_REC", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211728)
            
    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Public Function NFe_Obtem_Nome_Danfe(ByVal objNF As ClassNFiscal, sDanfe As String) As Long

Dim lErro As Long
Dim sDir As String

On Error GoTo Erro_NFe_Obtem_Nome_Danfe

    lErro = CF("NFe_Obtem_Dir_Xml", sDir)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'sDir = "C:\SGE\Log\"

    sDanfe = sDir & "DANFE_" & Format(glEmpresa, "00") & Format(objNF.iFilialEmpresa, "00") & "_" & objNF.sSerie & "_" & Format(objNF.lNumNotaFiscal, "000000000") & "_" & Format(objNF.dtDataEmissao, "yyyymmdd") & gsExtensaoGerRelExp

    NFe_Obtem_Nome_Danfe = SUCESSO
    
    Exit Function
    
Erro_NFe_Obtem_Nome_Danfe:
    
    NFe_Obtem_Nome_Danfe = gErr
    
    Select Case gErr
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211729)
            
    End Select
    
    Exit Function
    
End Function

Sub InfoAdic_AjustaAnexo(ByVal objInfoAdic As ClassInfoAdic)

Dim objAnexosArq As ClassAnexosArq

    Select Case objInfoAdic.iTipoDoc

            Case TIPODOC_INFOADIC_NF
                objInfoAdic.objAnexos.iTipoDoc = ANEXO_TIPO_NF
                
            Case TIPODOC_INFOADIC_PV
                objInfoAdic.objAnexos.iTipoDoc = ANEXO_TIPO_PV
            
            Case TIPODOC_INFOADIC_OV
                objInfoAdic.objAnexos.iTipoDoc = ANEXO_TIPO_OV
            
            Case TIPODOC_INFOADIC_OVHIST
                objInfoAdic.objAnexos.iTipoDoc = ANEXO_TIPO_OVHIST
                
            Case TIPODOC_INFOADIC_PC
                objInfoAdic.objAnexos.iTipoDoc = ANEXO_TIPO_PC

    End Select
    
    objInfoAdic.objAnexos.lNumIntDoc = objInfoAdic.lNumIntDoc
    For Each objAnexosArq In objInfoAdic.objAnexos.colArq
    
        objAnexosArq.iTipoDoc = objInfoAdic.objAnexos.iTipoDoc
        objAnexosArq.lNumIntDoc = objInfoAdic.lNumIntDoc
    
    Next

End Sub

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function NFSeLinks_Le(ByVal objNFSeLinks As ClassNFSeLinks) As Long

Dim lErro As Long
Dim lComando As Long
Dim tNFSeLinks As typeNFSeLinks
Dim sCodIBGE As String
Dim objFilialEmpresa As New AdmFiliais
Dim objCidade As ClassCidades

On Error GoTo Erro_NFSeLinks_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 213226
    
    If Len(Trim(objNFSeLinks.sCodIBGE)) <> 0 Then
        sCodIBGE = objNFSeLinks.sCodIBGE
    Else
        objFilialEmpresa.iCodFilial = objNFSeLinks.iFilialEmpresa
        If objFilialEmpresa.iCodFilial = 0 Then objFilialEmpresa.iCodFilial = giFilialEmpresa
        
        lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        Set objCidade = New ClassCidades

        objCidade.sDescricao = objFilialEmpresa.objEndereco.sCidade
        
        lErro = CF("Cidade_Le_Nome", objCidade)
        If lErro <> SUCESSO And lErro <> ERRO_OBJETO_NAO_CADASTRADO Then gError ERRO_SEM_MENSAGEM
        
        sCodIBGE = objCidade.sCodIBGE
    End If

    'Alocação de espaço no buffer
    tNFSeLinks.sCodIBGE = String(STRING_MAXIMO, 0)
    tNFSeLinks.sLinkConsulta = String(STRING_MAXIMO, 0)
    tNFSeLinks.sParamConsulta = String(STRING_MAXIMO, 0)
    tNFSeLinks.sLinkVerificacao = String(STRING_MAXIMO, 0)
    tNFSeLinks.sLinkSite = String(STRING_MAXIMO, 0)
    tNFSeLinks.sEmail = String(STRING_MAXIMO, 0)

    'Le a tabelaNFSeLinks
    lErro = Comando_Executar(lComando, "SELECT CodIBGE, LinkConsulta, ParamConsulta, LinkVerificacao, LinkSite, Email FROM NFSeLinks WHERE CodIBGE = ?", _
                tNFSeLinks.sCodIBGE, tNFSeLinks.sLinkConsulta, tNFSeLinks.sParamConsulta, tNFSeLinks.sLinkVerificacao, _
                tNFSeLinks.sLinkSite, tNFSeLinks.sEmail, sCodIBGE)
    If lErro <> AD_SQL_SUCESSO Then gError 213227

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213228
    
    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objNFSeLinks.sCodIBGE = tNFSeLinks.sCodIBGE
    objNFSeLinks.sLinkConsulta = tNFSeLinks.sLinkConsulta
    objNFSeLinks.sParamConsulta = tNFSeLinks.sParamConsulta
    objNFSeLinks.sLinkVerificacao = tNFSeLinks.sLinkVerificacao
    objNFSeLinks.sLinkSite = tNFSeLinks.sLinkSite
    objNFSeLinks.sEmail = tNFSeLinks.sEmail

    'Fecha Comando
    Call Comando_Fechar(lComando)

    NFSeLinks_Le = SUCESSO

    Exit Function

Erro_NFSeLinks_Le:

    NFSeLinks_Le = gErr

    Select Case gErr

        Case 213226
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 213227, 213228
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFSELINKS", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213229)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function Carrega_Combo_ModDocFis(ByVal objCombo As Object, Optional ByVal sFiltro As String) As Long

Dim lErro As Long
Dim lComando As Long
Dim sModelo As String, sDescr As String, iCod As Integer

On Error GoTo Erro_Carrega_Combo_ModDocFis

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 213589
    
    sModelo = String(STRING_MAXIMO, 0)
    sDescr = String(STRING_MAXIMO, 0)

    objCombo.Clear

    lErro = Comando_Executar(lComando, "SELECT Codigo, Modelo, Descricao FROM ModelosDocFiscais " & IIf(Len(Trim(sFiltro)) > 0, " WHERE ", " ") & sFiltro & " ORDER BY Codigo", iCod, sModelo, sDescr)
    If lErro <> AD_SQL_SUCESSO Then gError 213590

    'le o primeiro codigo de sTabela
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213591

    Do While lErro <> AD_SQL_SEM_DADOS

        objCombo.AddItem sModelo & SEPARADOR & sDescr
        objCombo.ItemData(objCombo.NewIndex) = iCod

        'le o proximo registro da tabela
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213592

    Loop

    Call Comando_Fechar(lComando)

    Carrega_Combo_ModDocFis = SUCESSO

    Exit Function

Erro_Carrega_Combo_ModDocFis:

    Carrega_Combo_ModDocFis = gErr

    Select Case gErr

        Case 213589
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 213590 To 213592
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA", gErr, "ModelosDocFiscais")

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213593)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function ModelosDocFiscais_Le(ByVal objModelosDocFiscais As ClassModelosDocFiscais) As Long

Dim lErro As Long
Dim lComando As Long
Dim tModelosDocFiscais As typeModelosDocFiscais

On Error GoTo Erro_ModelosDocFiscais_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 213594

    'Alocação de espaço no buffer
    tModelosDocFiscais.sModelo = String(STRING_MAXIMO, 0)
    tModelosDocFiscais.sDescricao = String(STRING_MAXIMO, 0)

    'Le a tabelaModelosDocFiscais
    lErro = Comando_Executar(lComando, "SELECT Codigo, Tipo, Eletronica, Modelo, Descricao FROM ModelosDocFiscais WHERE Codigo = ? ", _
                tModelosDocFiscais.iCodigo, tModelosDocFiscais.iTipo, tModelosDocFiscais.iEletronica, tModelosDocFiscais.sModelo, _
                tModelosDocFiscais.sDescricao, objModelosDocFiscais.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 213595

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213596

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objModelosDocFiscais.iCodigo = tModelosDocFiscais.iCodigo
    objModelosDocFiscais.iTipo = tModelosDocFiscais.iTipo
    objModelosDocFiscais.iEletronica = tModelosDocFiscais.iEletronica
    objModelosDocFiscais.sModelo = tModelosDocFiscais.sModelo
    objModelosDocFiscais.sDescricao = tModelosDocFiscais.sDescricao

    'Fecha Comando
    Call Comando_Fechar(lComando)

    ModelosDocFiscais_Le = SUCESSO

    Exit Function

Erro_ModelosDocFiscais_Le:

    ModelosDocFiscais_Le = gErr

    Select Case gErr

        Case 213594
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 213595, 213596
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA", gErr, "ModelosDocFiscais")

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213597)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function ClienteHistorico_Le(ByVal objClienteHistorico As ClassClienteHistorico) As Long

Dim lErro As Long
Dim lComando As Long
Dim tClienteHistorico As typeClienteHistorico

On Error GoTo Erro_ClienteHistorico_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 97806

    'Alocação de espaço no buffer
    tClienteHistorico.sCgc = String(STRING_MAXIMO, 0)
    tClienteHistorico.sInscricaoEstadual = String(STRING_MAXIMO, 0)
    tClienteHistorico.sInscricaoMunicipal = String(STRING_MAXIMO, 0)
    tClienteHistorico.sRazaoSocial = String(STRING_MAXIMO, 0)
    tClienteHistorico.sEndereco = String(STRING_MAXIMO, 0)
    tClienteHistorico.sBairro = String(STRING_MAXIMO, 0)
    tClienteHistorico.sCidade = String(STRING_MAXIMO, 0)
    tClienteHistorico.sUF = String(STRING_MAXIMO, 0)
    tClienteHistorico.sPais = String(STRING_MAXIMO, 0)
    tClienteHistorico.sCEP = String(STRING_MAXIMO, 0)
    tClienteHistorico.sInscricaoSuframa = String(STRING_MAXIMO, 0)
    tClienteHistorico.sUsuario = String(STRING_MAXIMO, 0)
    tClienteHistorico.sNomeReduzido = String(STRING_MAXIMO, 0)
    tClienteHistorico.sObservacao = String(STRING_MAXIMO, 0)
    tClienteHistorico.sUsuarioCobrador = String(STRING_MAXIMO, 0)
    tClienteHistorico.sUsuRespCallCenter = String(STRING_MAXIMO, 0)
    tClienteHistorico.sRG = String(STRING_MAXIMO, 0)
    tClienteHistorico.sTelefone1 = String(STRING_MAXIMO, 0)
    tClienteHistorico.sTelefone2 = String(STRING_MAXIMO, 0)
    tClienteHistorico.sEmail = String(STRING_MAXIMO, 0)
    tClienteHistorico.sEmail2 = String(STRING_MAXIMO, 0)
    tClienteHistorico.sMensagemNF = String(STRING_MAXIMO, 0)

    'Le a tabelaClienteHistorico
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, DataAtualizacao, Codigo, Filial, CPF_CGC, InscricaoEstadual, " & _
                "InscricaoMunicipal, RazaoSocial, Endereco, Bairro, Cidade, UF, Pais, " & _
                "CEP, Ativo, CodigoLoja, FilialEmpresaLoja, CodFilialLoja, InscricaoSuframa, Usuario, " & _
                "DataReg, HoraReg, NomeReduzido, Tipo, LimiteCredito, Observacao, CondicaoPagto, " & _
                "TabelaPreco, UsuarioCobrador, UsuRespCallCenter, CodTransportadora, Vendedor, RG, RegimeTributario, " & _
                "Telefone1, Telefone2, Email, Email2, MensagemNF, Desconto, " & _
                "ComissaoVendas, Regiao FROM ClienteHistorico WHERE NumIntDoc= ? ", _
                tClienteHistorico.lNumIntDoc, tClienteHistorico.dtDataAtualizacao, tClienteHistorico.lCodigo, tClienteHistorico.iFilialCliente, _
                tClienteHistorico.sCgc, tClienteHistorico.sInscricaoEstadual, tClienteHistorico.sInscricaoMunicipal, tClienteHistorico.sRazaoSocial, tClienteHistorico.sEndereco, _
                tClienteHistorico.sBairro, tClienteHistorico.sCidade, tClienteHistorico.sUF, tClienteHistorico.sPais, tClienteHistorico.sCEP, _
                tClienteHistorico.iAtivo, tClienteHistorico.lCodigoLoja, tClienteHistorico.iFilialEmpresaLoja, tClienteHistorico.iCodFilialLoja, tClienteHistorico.sInscricaoSuframa, _
                tClienteHistorico.sUsuario, tClienteHistorico.dtDataReg, tClienteHistorico.dHoraReg, tClienteHistorico.sNomeReduzido, tClienteHistorico.iTipo, _
                tClienteHistorico.dLimiteCredito, tClienteHistorico.sObservacao, tClienteHistorico.iCondicaoPagto, tClienteHistorico.iTabelaPreco, tClienteHistorico.sUsuarioCobrador, _
                tClienteHistorico.sUsuRespCallCenter, tClienteHistorico.iCodTransportadora, tClienteHistorico.iVendedor, tClienteHistorico.sRG, tClienteHistorico.iRegimeTributario, _
                tClienteHistorico.sTelefone1, tClienteHistorico.sTelefone2, tClienteHistorico.sEmail, tClienteHistorico.sEmail2, tClienteHistorico.sMensagemNF, _
                tClienteHistorico.dDesconto, tClienteHistorico.dComissaoVendas, tClienteHistorico.iRegiao, _
                objClienteHistorico.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 97807

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97808

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objClienteHistorico.dtDataAtualizacao = tClienteHistorico.dtDataAtualizacao
    objClienteHistorico.lCodigo = tClienteHistorico.lCodigo
    objClienteHistorico.iFilialCliente = tClienteHistorico.iFilialCliente
    objClienteHistorico.sCgc = tClienteHistorico.sCgc
    objClienteHistorico.sInscricaoEstadual = tClienteHistorico.sInscricaoEstadual
    objClienteHistorico.sInscricaoMunicipal = tClienteHistorico.sInscricaoMunicipal
    objClienteHistorico.sRazaoSocial = tClienteHistorico.sRazaoSocial
    objClienteHistorico.sEndereco = tClienteHistorico.sEndereco
    objClienteHistorico.sBairro = tClienteHistorico.sBairro
    objClienteHistorico.sCidade = tClienteHistorico.sCidade
    objClienteHistorico.sUF = tClienteHistorico.sUF
    objClienteHistorico.sPais = tClienteHistorico.sPais
    objClienteHistorico.sCEP = tClienteHistorico.sCEP
    objClienteHistorico.iAtivo = tClienteHistorico.iAtivo
    objClienteHistorico.lCodigoLoja = tClienteHistorico.lCodigoLoja
    objClienteHistorico.iFilialEmpresaLoja = tClienteHistorico.iFilialEmpresaLoja
    objClienteHistorico.iCodFilialLoja = tClienteHistorico.iCodFilialLoja
    objClienteHistorico.sInscricaoSuframa = tClienteHistorico.sInscricaoSuframa
    objClienteHistorico.sUsuario = tClienteHistorico.sUsuario
    objClienteHistorico.dtDataReg = tClienteHistorico.dtDataReg
    objClienteHistorico.dHoraReg = tClienteHistorico.dHoraReg
    objClienteHistorico.sNomeReduzido = tClienteHistorico.sNomeReduzido
    objClienteHistorico.iTipo = tClienteHistorico.iTipo
    objClienteHistorico.dLimiteCredito = tClienteHistorico.dLimiteCredito
    objClienteHistorico.sObservacao = tClienteHistorico.sObservacao
    objClienteHistorico.iCondicaoPagto = tClienteHistorico.iCondicaoPagto
    objClienteHistorico.iTabelaPreco = tClienteHistorico.iTabelaPreco
    objClienteHistorico.sUsuarioCobrador = tClienteHistorico.sUsuarioCobrador
    objClienteHistorico.sUsuRespCallCenter = tClienteHistorico.sUsuRespCallCenter
    objClienteHistorico.iCodTransportadora = tClienteHistorico.iCodTransportadora
    objClienteHistorico.iVendedor = tClienteHistorico.iVendedor
    objClienteHistorico.sRG = tClienteHistorico.sRG
    objClienteHistorico.iRegimeTributario = tClienteHistorico.iRegimeTributario
    objClienteHistorico.sTelefone1 = tClienteHistorico.sTelefone1
    objClienteHistorico.sTelefone2 = tClienteHistorico.sTelefone2
    objClienteHistorico.sEmail = tClienteHistorico.sEmail
    objClienteHistorico.sEmail2 = tClienteHistorico.sEmail2
    objClienteHistorico.sMensagemNF = tClienteHistorico.sMensagemNF
    objClienteHistorico.dDesconto = tClienteHistorico.dDesconto
    objClienteHistorico.dComissaoVendas = tClienteHistorico.dComissaoVendas
    objClienteHistorico.iRegiao = tClienteHistorico.iRegiao

    'Fecha Comando
    Call Comando_Fechar(lComando)

    ClienteHistorico_Le = SUCESSO

    Exit Function

Erro_ClienteHistorico_Le:

    ClienteHistorico_Le = gErr

    Select Case gErr

        Case 97806
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 97807, 97808
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLIENTEHISTORICO", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 97809)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

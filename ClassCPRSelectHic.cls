VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCPRSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function NFPag_Le(ByVal objNFPag As ClassNFsPag, Optional ByVal iExclusao As Integer = 0) As Long
'Retorna os dados da Nota Fiscal com nº interno passado como parametro

Dim lErro As Long
Dim lComando As Long
Dim tNFsPag As typeNFsPag
Dim objNFPagTitPag As New ClassNfPagTitPagHic 'Inserido por Wagner
Dim lSituacao As Long
Dim lHistorico As Long
Dim lSubConta As Long
Dim sObservacao As String

On Error GoTo Erro_NFPag_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 18254

    '##############################
    'Inserido por Wagner
    sObservacao = String(STRING_NFPAG_OBSERVACAO, 0)
    '##############################

    If iExclusao = 1 Then

        '##########################
        'Alterado por Wagner
        'Procura a Nota com o NumIntDoc passado como parametro
        lErro = Comando_Executar(lComando, "SELECT FilialEmpresa, Fornecedor, Filial, NumNotaFiscal, DataEmissao, Status, NumIntTitPag, DataVencimento, ValorTotal, ValorSeguro, ValorFrete, OutrasDespesas, ValorProdutos, ValorIRRF, ValorICMS, ValorICMSSubst, CreditoICMS, ValorIPI, CreditoIPI, PISRetido, COFINSRetido, CSLLRetido, TaxaMoeda, Situacao, Historico, SubConta, Observacao FROM NFsPag WHERE NumIntDoc = ?", tNFsPag.iFilialEmpresa, tNFsPag.lFornecedor, tNFsPag.iFilial, tNFsPag.lNumNotaFiscal, tNFsPag.dtDataEmissao, tNFsPag.iStatus, tNFsPag.lNumIntTitPag, tNFsPag.dtDataVencimento, tNFsPag.dValorTotal, tNFsPag.dValorSeguro, tNFsPag.dValorFrete, tNFsPag.dOutrasDespesas, tNFsPag.dValorProdutos, tNFsPag.dValorIRRF, tNFsPag.dValorICMS, tNFsPag.dValorICMSSubst, tNFsPag.iCreditoICMS, tNFsPag.dValorIPI, tNFsPag.iCreditoIPI, tNFsPag.dPISRetido, tNFsPag.dCOFINSRetido, tNFsPag.dCSLLRetido, tNFsPag.dTaxaMoeda, lSituacao, lHistorico, lSubConta, sObservacao, objNFPag.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then Error 18255
        '##########################

    Else
    
        '##########################
        'Alterado por Wagner
        'Procura a Nota com o NumIntDoc passado como parametro
        lErro = Comando_Executar(lComando, "SELECT FilialEmpresa, Fornecedor, Filial, NumNotaFiscal, DataEmissao, Status, NumIntTitPag, DataVencimento, ValorTotal, ValorSeguro, ValorFrete, OutrasDespesas, ValorProdutos, ValorIRRF, ValorICMS, ValorICMSSubst, CreditoICMS, ValorIPI, CreditoIPI, PISRetido, COFINSRetido, CSLLRetido, TaxaMoeda, Situacao, Historico, SubConta, Observacao FROM NFsPag WHERE NumIntDoc = ? AND Status <> ?", tNFsPag.iFilialEmpresa, tNFsPag.lFornecedor, tNFsPag.iFilial, tNFsPag.lNumNotaFiscal, tNFsPag.dtDataEmissao, tNFsPag.iStatus, tNFsPag.lNumIntTitPag, tNFsPag.dtDataVencimento, tNFsPag.dValorTotal, tNFsPag.dValorSeguro, tNFsPag.dValorFrete, tNFsPag.dOutrasDespesas, tNFsPag.dValorProdutos, tNFsPag.dValorIRRF, tNFsPag.dValorICMS, tNFsPag.dValorICMSSubst, tNFsPag.iCreditoICMS, tNFsPag.dValorIPI, tNFsPag.iCreditoIPI, tNFsPag.dPISRetido, tNFsPag.dCOFINSRetido, tNFsPag.dCSLLRetido, tNFsPag.dTaxaMoeda, lSituacao, lHistorico, lSubConta, _
                sObservacao, objNFPag.lNumIntDoc, STATUS_EXCLUIDO)
        If lErro <> AD_SQL_SUCESSO Then Error 18255
        '##########################
    
    End If
    
    'Le a nota
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 18256

    'Se não existe erro
    If lErro = AD_SQL_SEM_DADOS Then Error 18257

    'Passa para objNFPag os dados Lidos
    objNFPag.iFilialEmpresa = tNFsPag.iFilialEmpresa
    objNFPag.lFornecedor = tNFsPag.lFornecedor
    objNFPag.iFilial = tNFsPag.iFilial
    objNFPag.lNumNotaFiscal = tNFsPag.lNumNotaFiscal
    objNFPag.dtDataEmissao = tNFsPag.dtDataEmissao
    objNFPag.iStatus = tNFsPag.iStatus
    objNFPag.lNumIntTitPag = tNFsPag.lNumIntTitPag
    objNFPag.dtDataVencimento = tNFsPag.dtDataVencimento
    objNFPag.dValorTotal = tNFsPag.dValorTotal
    objNFPag.dValorSeguro = tNFsPag.dValorSeguro
    objNFPag.dValorFrete = tNFsPag.dValorFrete
    objNFPag.dOutrasDespesas = tNFsPag.dOutrasDespesas
    objNFPag.dValorProdutos = tNFsPag.dValorProdutos
    objNFPag.dValorIRRF = tNFsPag.dValorIRRF
    objNFPag.dValorICMS = tNFsPag.dValorICMS
    objNFPag.dValorICMSSubst = tNFsPag.dValorICMSSubst
    objNFPag.iCreditoICMS = tNFsPag.iCreditoICMS
    objNFPag.dValorIPI = tNFsPag.dValorIPI
    objNFPag.iCreditoIPI = tNFsPag.iCreditoIPI
    objNFPag.dPISRetido = tNFsPag.dPISRetido
    objNFPag.dCOFINSRetido = tNFsPag.dCOFINSRetido
    objNFPag.dCSLLRetido = tNFsPag.dCSLLRetido
    objNFPag.dTaxaMoeda = tNFsPag.dTaxaMoeda
    
    '###############################################
    'Inserido por Wagner 18/02/2005
    objNFPagTitPag.lSituacao = lSituacao
    objNFPagTitPag.lHistorico = lHistorico
    objNFPagTitPag.lSubConta = lSubConta
    objNFPagTitPag.sObservacao = sObservacao
    
    Set objNFPag.objInfoUsu = objNFPagTitPag
    '###############################################

    Call Comando_Fechar(lComando)

    NFPag_Le = SUCESSO

    Exit Function

Erro_NFPag_Le:

    NFPag_Le = Err

    Select Case Err

        Case 18254
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 18255, 18256
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFSPAG1", Err, objNFPag.lNumNotaFiscal)

        Case 18257  'Será tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 179163)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function NFPag_Le_Numero(ByVal objNFsPag As ClassNFsPag) As Long
'Le os dados da Nota Fiscal com o Fornecedor, Filial, Numero Nota Fiscal e Data de Emissão passados por parametro

Dim lErro As Long
Dim lComando As Long
Dim tNFsPag As typeNFsPag
Dim objNFPagTitPag As New ClassNfPagTitPagHic 'Inserido por Wagner
Dim lSituacao As Long
Dim lHistorico As Long
Dim lSubConta As Long
Dim sObservacao As String

On Error GoTo Erro_NFPag_Le_Numero

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 18335
    
    '##############################
    'Inserido por Wagner
    sObservacao = String(STRING_NFPAG_OBSERVACAO, 0)
    '##############################
    
    '##########################
    'Alterado por Wagner
    'Seleciona a nota com os dados passados
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Status, NumIntTitPag, DataVencimento, ValorTotal, ValorSeguro, ValorFrete, OutrasDespesas, ValorProdutos, ValorIRRF, ValorICMS, ValorICMSSubst, CreditoICMS, ValorIPI, CreditoIPI, PISRetido, COFINSRetido, CSLLRetido, TaxaMoeda, Situacao, Historico, SubConta, Observacao FROM NFsPag WHERE Fornecedor = ? AND Filial = ? AND NumNotaFiscal = ? AND DataEmissao = ? AND FilialEmpresa = ? AND Status <> ? ", tNFsPag.lNumIntDoc, tNFsPag.iStatus, tNFsPag.lNumIntTitPag, tNFsPag.dtDataVencimento, tNFsPag.dValorTotal, tNFsPag.dValorSeguro, tNFsPag.dValorFrete, tNFsPag.dOutrasDespesas, tNFsPag.dValorProdutos, tNFsPag.dValorIRRF, tNFsPag.dValorICMS, tNFsPag.dValorICMSSubst, tNFsPag.iCreditoICMS, tNFsPag.dValorIPI, tNFsPag.iCreditoIPI, tNFsPag.dPISRetido, tNFsPag.dCOFINSRetido, tNFsPag.dCSLLRetido, tNFsPag.dTaxaMoeda _
                                        , lSituacao, lHistorico, lSubConta, sObservacao, objNFsPag.lFornecedor, objNFsPag.iFilial, objNFsPag.lNumNotaFiscal, objNFsPag.dtDataEmissao, giFilialEmpresa, STATUS_EXCLUIDO)
    If lErro <> AD_SQL_SUCESSO Then Error 18336
    '##########################
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 18337
    
    'Se nao encontra erro
    If lErro = AD_SQL_SEM_DADOS Then Error 18338
    
    'Preenche o objNFsPag
    objNFsPag.lNumIntDoc = tNFsPag.lNumIntDoc
    objNFsPag.iFilialEmpresa = giFilialEmpresa
    objNFsPag.iStatus = tNFsPag.iStatus
    objNFsPag.lNumIntTitPag = tNFsPag.lNumIntTitPag
    objNFsPag.dtDataVencimento = tNFsPag.dtDataVencimento
    objNFsPag.dValorTotal = tNFsPag.dValorTotal
    objNFsPag.dValorSeguro = tNFsPag.dValorSeguro
    objNFsPag.dValorFrete = tNFsPag.dValorFrete
    objNFsPag.dOutrasDespesas = tNFsPag.dOutrasDespesas
    objNFsPag.dValorProdutos = tNFsPag.dValorProdutos
    objNFsPag.dValorIRRF = tNFsPag.dValorIRRF
    objNFsPag.dValorICMS = tNFsPag.dValorICMS
    objNFsPag.dValorICMSSubst = tNFsPag.dValorICMSSubst
    objNFsPag.iCreditoICMS = tNFsPag.iCreditoICMS
    objNFsPag.dValorIPI = tNFsPag.dValorIPI
    objNFsPag.iCreditoIPI = tNFsPag.iCreditoIPI
    objNFsPag.dPISRetido = tNFsPag.dPISRetido
    objNFsPag.dCOFINSRetido = tNFsPag.dCOFINSRetido
    objNFsPag.dCSLLRetido = tNFsPag.dCSLLRetido
    objNFsPag.dTaxaMoeda = tNFsPag.dTaxaMoeda
    
    '###############################################
    'Inserido por Wagner 18/02/2005
    objNFPagTitPag.lSituacao = lSituacao
    objNFPagTitPag.lHistorico = lHistorico
    objNFPagTitPag.lSubConta = lSubConta
    objNFPagTitPag.sObservacao = sObservacao
    
    Set objNFsPag.objInfoUsu = objNFPagTitPag
    '###############################################
    
    Call Comando_Fechar(lComando)
    
    NFPag_Le_Numero = SUCESSO
    
    Exit Function
    
Erro_NFPag_Le_Numero:

    NFPag_Le_Numero = Err
    
    Select Case Err
    
        Case 18335
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 18336, 18337
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFSPAG1", Err, objNFsPag.lNumNotaFiscal)
        
        Case 18338  'Erro tratado na rotina chamadora
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 179164)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function NFPagBaixada_Le_Numero(ByVal objNFsPag As ClassNFsPag) As Long
'Retorna os dados da nota fiscal com Fornecedor, Filial, NumNotaFiscal e DataEmissao passados passado por parametro

Dim lErro As Long
Dim lComando As Long
Dim tNFsPag As typeNFsPag
Dim objNFPagTitPag As New ClassNfPagTitPagHic 'Inserido por Wagner
Dim lSituacao As Long
Dim lHistorico As Long
Dim lSubConta As Long
Dim sObservacao As String

On Error GoTo Erro_NFPagBaixada_Le_Numero

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 18341
    
    '##############################
    'Inserido por Wagner
    sObservacao = String(STRING_NFPAG_OBSERVACAO, 0)
    '##############################
    
    '##############################
    'Alterado por Wagner 18/02/2005
    'Pesquisa a nota com os dados dos parametros
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Status, NumIntTitPag, DataVencimento, ValorTotal, ValorSeguro, ValorFrete, OutrasDespesas, ValorProdutos, ValorIRRF, ValorICMS, ValorICMSSubst, CreditoICMS, ValorIPI, CreditoIPI, PISRetido, COFINSRetido, CSLLRetido, TaxaMoeda, Situacao, Historico, SubConta, Observacao FROM NFsPagBaixadas WHERE Fornecedor = ? AND Filial = ? AND NumNotaFiscal = ? AND DataEmissao = ? AND FilialEmpresa = ? AND Status <> ? ", tNFsPag.lNumIntDoc, tNFsPag.iStatus, tNFsPag.lNumIntTitPag, tNFsPag.dtDataVencimento, tNFsPag.dValorTotal, tNFsPag.dValorSeguro, tNFsPag.dValorFrete, tNFsPag.dOutrasDespesas, tNFsPag.dValorProdutos, tNFsPag.dValorIRRF, tNFsPag.dValorICMS, tNFsPag.dValorICMSSubst, tNFsPag.iCreditoICMS, tNFsPag.dValorIPI, tNFsPag.iCreditoIPI, tNFsPag.dPISRetido, tNFsPag.dCOFINSRetido, tNFsPag.dCSLLRetido, tNFsPag.dTaxaMoeda _
                                        , lSituacao, lHistorico, lSubConta, sObservacao, objNFsPag.lFornecedor, objNFsPag.iFilial, objNFsPag.lNumNotaFiscal, objNFsPag.dtDataEmissao, giFilialEmpresa, STATUS_EXCLUIDO)
    If lErro <> AD_SQL_SUCESSO Then Error 18342
    '##############################
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 18343
    
    'NÃo encontrou
    If lErro = AD_SQL_SEM_DADOS Then Error 18344
    
    'Preenche o objNFsPag
    objNFsPag.lNumIntDoc = tNFsPag.lNumIntDoc
    objNFsPag.iFilialEmpresa = giFilialEmpresa
    objNFsPag.iStatus = tNFsPag.iStatus
    objNFsPag.lNumIntTitPag = tNFsPag.lNumIntTitPag
    objNFsPag.dtDataVencimento = tNFsPag.dtDataVencimento
    objNFsPag.dValorTotal = tNFsPag.dValorTotal
    objNFsPag.dValorSeguro = tNFsPag.dValorSeguro
    objNFsPag.dValorFrete = tNFsPag.dValorFrete
    objNFsPag.dOutrasDespesas = tNFsPag.dOutrasDespesas
    objNFsPag.dValorProdutos = tNFsPag.dValorProdutos
    objNFsPag.dValorIRRF = tNFsPag.dValorIRRF
    objNFsPag.dValorICMS = tNFsPag.dValorICMS
    objNFsPag.dValorICMSSubst = tNFsPag.dValorICMSSubst
    objNFsPag.iCreditoICMS = tNFsPag.iCreditoICMS
    objNFsPag.dValorIPI = tNFsPag.dValorIPI
    objNFsPag.iCreditoIPI = tNFsPag.iCreditoIPI
    objNFsPag.dPISRetido = tNFsPag.dPISRetido
    objNFsPag.dCOFINSRetido = tNFsPag.dCOFINSRetido
    objNFsPag.dCSLLRetido = tNFsPag.dCSLLRetido
    objNFsPag.dTaxaMoeda = tNFsPag.dTaxaMoeda
    
    '###############################################
    'Inserido por Wagner 18/02/2005
    objNFPagTitPag.lSituacao = lSituacao
    objNFPagTitPag.lHistorico = lHistorico
    objNFPagTitPag.lSubConta = lSubConta
    objNFPagTitPag.sObservacao = sObservacao
    
    Set objNFsPag.objInfoUsu = objNFPagTitPag
    '###############################################
    
    Call Comando_Fechar(lComando)
    
    NFPagBaixada_Le_Numero = SUCESSO
    
    Exit Function
    
Erro_NFPagBaixada_Le_Numero:

    NFPagBaixada_Le_Numero = Err
    
    Select Case Err
    
        Case 18341
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 18342, 18343
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFSPAGBAIXADAS", Err, objNFsPag.lNumNotaFiscal)
        
        Case 18344 ' Erro tratado na rotina chamadora
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 179165)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function NFPagBaixada_Le(ByVal objNFPag As ClassNFsPag, Optional ByVal iExclusao As Integer = 0) As Long
'Retorna os dados da Nota Fiscal com nº interno passado como parametro

Dim lErro As Long
Dim lComando As Long
Dim tNFsPag As typeNFsPag
Dim objNFPagTitPag As New ClassNfPagTitPagHic 'Inserido por Wagner
Dim lSituacao As Long
Dim lHistorico As Long
Dim lSubConta As Long
Dim sObservacao As String

On Error GoTo Erro_NFPagBaixada_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 18254

    '##############################
    'Inserido por Wagner
    sObservacao = String(STRING_NFPAG_OBSERVACAO, 0)
    '##############################

    If iExclusao = 1 Then

        '##########################
        'Alterado por Wagner
        'Procura a Nota com o NumIntDoc passado como parametro
        lErro = Comando_Executar(lComando, "SELECT FilialEmpresa, Fornecedor, Filial, NumNotaFiscal, DataEmissao, Status, NumIntTitPag, DataVencimento, ValorTotal, ValorSeguro, ValorFrete, OutrasDespesas, ValorProdutos, ValorIRRF, ValorICMS, ValorICMSSubst, CreditoICMS, ValorIPI, CreditoIPI, PISRetido, COFINSRetido, CSLLRetido, TaxaMoeda, Situacao, Historico, SubConta, Observacao FROM NFsPagBaixadas WHERE NumIntDoc = ?", tNFsPag.iFilialEmpresa, tNFsPag.lFornecedor, tNFsPag.iFilial, tNFsPag.lNumNotaFiscal, tNFsPag.dtDataEmissao, tNFsPag.iStatus, tNFsPag.lNumIntTitPag, tNFsPag.dtDataVencimento, tNFsPag.dValorTotal, tNFsPag.dValorSeguro, tNFsPag.dValorFrete, tNFsPag.dOutrasDespesas, tNFsPag.dValorProdutos, tNFsPag.dValorIRRF, tNFsPag.dValorICMS, tNFsPag.dValorICMSSubst, tNFsPag.iCreditoICMS, tNFsPag.dValorIPI, tNFsPag.iCreditoIPI, tNFsPag.dPISRetido, tNFsPag.dCOFINSRetido, tNFsPag.dCSLLRetido, tNFsPag.dTaxaMoeda _
                                        , lSituacao, lHistorico, lSubConta, sObservacao, objNFPag.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then Error 18255
        '##########################

    Else
    
        '##########################
        'Alterado por Wagner
        'Procura a Nota com o NumIntDoc passado como parametro
        lErro = Comando_Executar(lComando, "SELECT FilialEmpresa, Fornecedor, Filial, NumNotaFiscal, DataEmissao, Status, NumIntTitPag, DataVencimento, ValorTotal, ValorSeguro, ValorFrete, OutrasDespesas, ValorProdutos, ValorIRRF, ValorICMS, ValorICMSSubst, CreditoICMS, ValorIPI, CreditoIPI, PISRetido, COFINSRetido, CSLLRetido, TaxaMoeda, Situacao, Historico, SubConta, Observacao FROM NFsPagBaixadas WHERE NumIntDoc = ? AND Status <> ?", tNFsPag.iFilialEmpresa, tNFsPag.lFornecedor, tNFsPag.iFilial, tNFsPag.lNumNotaFiscal, tNFsPag.dtDataEmissao, tNFsPag.iStatus, tNFsPag.lNumIntTitPag, tNFsPag.dtDataVencimento, tNFsPag.dValorTotal, tNFsPag.dValorSeguro, tNFsPag.dValorFrete, tNFsPag.dOutrasDespesas, tNFsPag.dValorProdutos, tNFsPag.dValorIRRF, tNFsPag.dValorICMS, tNFsPag.dValorICMSSubst, tNFsPag.iCreditoICMS, tNFsPag.dValorIPI, tNFsPag.iCreditoIPI, tNFsPag.dPISRetido, tNFsPag.dCOFINSRetido, tNFsPag.dCSLLRetido, tNFsPag.dTaxaMoeda _
                                            , lSituacao, lHistorico, lSubConta, sObservacao, objNFPag.lNumIntDoc, STATUS_EXCLUIDO)
        If lErro <> AD_SQL_SUCESSO Then Error 18255
        '##########################
    
    End If
    
    'Le a nota
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 18256

    'Se não existe erro
    If lErro = AD_SQL_SEM_DADOS Then Error 18257

    'Passa para objNFPag os dados Lidos
    objNFPag.iFilialEmpresa = tNFsPag.iFilialEmpresa
    objNFPag.lFornecedor = tNFsPag.lFornecedor
    objNFPag.iFilial = tNFsPag.iFilial
    objNFPag.lNumNotaFiscal = tNFsPag.lNumNotaFiscal
    objNFPag.dtDataEmissao = tNFsPag.dtDataEmissao
    objNFPag.iStatus = tNFsPag.iStatus
    objNFPag.lNumIntTitPag = tNFsPag.lNumIntTitPag
    objNFPag.dtDataVencimento = tNFsPag.dtDataVencimento
    objNFPag.dValorTotal = tNFsPag.dValorTotal
    objNFPag.dValorSeguro = tNFsPag.dValorSeguro
    objNFPag.dValorFrete = tNFsPag.dValorFrete
    objNFPag.dOutrasDespesas = tNFsPag.dOutrasDespesas
    objNFPag.dValorProdutos = tNFsPag.dValorProdutos
    objNFPag.dValorIRRF = tNFsPag.dValorIRRF
    objNFPag.dValorICMS = tNFsPag.dValorICMS
    objNFPag.dValorICMSSubst = tNFsPag.dValorICMSSubst
    objNFPag.iCreditoICMS = tNFsPag.iCreditoICMS
    objNFPag.dValorIPI = tNFsPag.dValorIPI
    objNFPag.iCreditoIPI = tNFsPag.iCreditoIPI
    objNFPag.dPISRetido = tNFsPag.dPISRetido
    objNFPag.dCOFINSRetido = tNFsPag.dCOFINSRetido
    objNFPag.dCSLLRetido = tNFsPag.dCSLLRetido
    objNFPag.dTaxaMoeda = tNFsPag.dTaxaMoeda

    '###############################################
    'Inserido por Wagner 18/02/2005
    objNFPagTitPag.lSituacao = lSituacao
    objNFPagTitPag.lHistorico = lHistorico
    objNFPagTitPag.lSubConta = lSubConta
    objNFPagTitPag.sObservacao = sObservacao
    
    Set objNFPag.objInfoUsu = objNFPagTitPag
    '###############################################
    
    Call Comando_Fechar(lComando)

    NFPagBaixada_Le = SUCESSO

    Exit Function

Erro_NFPagBaixada_Le:

    NFPagBaixada_Le = Err

    Select Case Err

        Case 18254
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 18255, 18256
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFSPAG1", Err, objNFPag.lNumNotaFiscal)

        Case 18257  'Será tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 179166)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function TituloPagar_Le(ByVal objTituloPagar As ClassTituloPagar, Optional ByVal bLeExcluido As Boolean = False) As Long
'Lê Titulo a Pagar com o Número Interno passado como parâmetro

Dim lErro As Long
Dim lComando As Long
Dim tTituloPagar As typeTituloPagar
Dim objNFPagTitPag As New ClassNfPagTitPagHic 'Inserido por Wagner
Dim iChequeRetido As Integer 'Inserido por Wagner
Dim lSituacao As Long
Dim lHistorico As Long
Dim lSubConta As Long

On Error GoTo Erro_TituloPagar_Le

    'Abre comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 18369
    
    tTituloPagar.sSiglaDocumento = String(STRING_SIGLA_DOCUMENTO, 0)
    tTituloPagar.sObservacao = String(STRING_TITULO_OBSERVACAO, 0)
    
    '##############################
    'Alterado por Wagner 18/02/2005
    'Seleciona o TituloPagar a partir do NumIntDoc
    If bLeExcluido Then
        lErro = Comando_Executar(lComando, "SELECT FilialEmpresa, Fornecedor, Filial, SiglaDocumento, NumTitulo, DataEmissao, Saldo, NumParcelas, ValorTotal, ValorSeguro, ValorFrete, OutrasDespesas, ValorProdutos, ValorICMS, ValorICMSSubst, CreditoICMS, ValorIPI, CreditoIPI, ValorIRRF, ValorINSS, INSSRetido, DataRegistro, " & _
            "CondicaoPagto, Observacao, Status, FilialPedCompra, NumPedCompra, PISRetido, ISSRetido, COFINSRetido, CSLLRetido, TaxaMoeda, Situacao, Historico, SubConta, ChequeRetido FROM TitulosPag WHERE NumIntDoc = ? ", tTituloPagar.iFilialEmpresa, tTituloPagar.lFornecedor, tTituloPagar.iFilial, tTituloPagar.sSiglaDocumento, tTituloPagar.lNumTitulo, tTituloPagar.dtDataEmissao, tTituloPagar.dSaldo, tTituloPagar.iNumParcelas, tTituloPagar.dValorTotal, tTituloPagar.dValorSeguro, tTituloPagar.dValorFrete, tTituloPagar.dOutrasDespesas, tTituloPagar.dValorProdutos, tTituloPagar.dValorICMS, tTituloPagar.dValorICMSSubst, tTituloPagar.iCreditoICMS, tTituloPagar.dValorIPI, tTituloPagar.iCreditoIPI, tTituloPagar.dValorIRRF, tTituloPagar.dValorINSS, tTituloPagar.iINSSRetido, tTituloPagar.dtDataRegistro, tTituloPagar.iCondicaoPagto, _
        tTituloPagar.sObservacao, tTituloPagar.iStatus, tTituloPagar.iFilialPedCompra, tTituloPagar.lNumPedCompra, tTituloPagar.dPISRetido, tTituloPagar.dISSRetido, tTituloPagar.dCOFINSRetido, tTituloPagar.dCSLLRetido, tTituloPagar.dTaxaMoeda, lSituacao, lHistorico, lSubConta, iChequeRetido, objTituloPagar.lNumIntDoc)
    
    Else
        lErro = Comando_Executar(lComando, "SELECT FilialEmpresa, Fornecedor, Filial, SiglaDocumento, NumTitulo, DataEmissao, Saldo, NumParcelas, ValorTotal, ValorSeguro, ValorFrete, OutrasDespesas, ValorProdutos, ValorICMS, ValorICMSSubst, CreditoICMS, ValorIPI, CreditoIPI, ValorIRRF, ValorINSS, INSSRetido, DataRegistro, " & _
            "CondicaoPagto, Observacao, Status, FilialPedCompra, NumPedCompra, PISRetido, ISSRetido, COFINSRetido, CSLLRetido, TaxaMoeda, Situacao, Historico, SubConta, ChequeRetido FROM TitulosPag WHERE NumIntDoc = ? AND Status <> ?", tTituloPagar.iFilialEmpresa, tTituloPagar.lFornecedor, tTituloPagar.iFilial, tTituloPagar.sSiglaDocumento, tTituloPagar.lNumTitulo, tTituloPagar.dtDataEmissao, tTituloPagar.dSaldo, tTituloPagar.iNumParcelas, tTituloPagar.dValorTotal, tTituloPagar.dValorSeguro, tTituloPagar.dValorFrete, tTituloPagar.dOutrasDespesas, tTituloPagar.dValorProdutos, tTituloPagar.dValorICMS, tTituloPagar.dValorICMSSubst, tTituloPagar.iCreditoICMS, tTituloPagar.dValorIPI, tTituloPagar.iCreditoIPI, tTituloPagar.dValorIRRF, tTituloPagar.dValorINSS, tTituloPagar.iINSSRetido, tTituloPagar.dtDataRegistro, tTituloPagar.iCondicaoPagto, _
        tTituloPagar.sObservacao, tTituloPagar.iStatus, tTituloPagar.iFilialPedCompra, tTituloPagar.lNumPedCompra, tTituloPagar.dPISRetido, tTituloPagar.dISSRetido, tTituloPagar.dCOFINSRetido, tTituloPagar.dCSLLRetido, tTituloPagar.dTaxaMoeda, lSituacao, lHistorico, lSubConta, iChequeRetido, objTituloPagar.lNumIntDoc, STATUS_EXCLUIDO)
    End If
    If lErro <> AD_SQL_SUCESSO Then Error 18370
    '##############################
    
    'Tenta ler
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 18371
    
    'Se não encontrou --> Erro
    If lErro = AD_SQL_SEM_DADOS Then Error 18372
    
    tTituloPagar.lNumIntDoc = objTituloPagar.lNumIntDoc
    
    'Preenche objTituloPagar com os dados do Título
    Call TituloPagar_MoverCamposObj(objTituloPagar, tTituloPagar)

    '###############################################
    'Inserido por Wagner 18/02/2005
    objNFPagTitPag.lSituacao = lSituacao
    objNFPagTitPag.lHistorico = lHistorico
    objNFPagTitPag.lSubConta = lSubConta
    objNFPagTitPag.sObservacao = tTituloPagar.sObservacao
    objNFPagTitPag.iChequeRetido = iChequeRetido '19/04/2005
    
    Set objTituloPagar.objInfoUsu = objNFPagTitPag
    '###############################################
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    TituloPagar_Le = SUCESSO
    
    Exit Function
    
Erro_TituloPagar_Le:

    TituloPagar_Le = Err
    
    Select Case Err
    
        Case 18369
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 18370, 18371
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOSPAG", Err, objTituloPagar.lNumIntDoc)
        
        Case 18372  'Título não encontrado
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 179167)
    
    End Select
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function TituloPagar_Le_Numero(ByVal objTituloPagar As ClassTituloPagar) As Long
'Lê o Titulo da filial corrente com o Fornecedor, Filial, Numtitulo, Data de Emissao e Sigla do Documento passados em objtitulopagar

Dim lErro As Long
Dim lComando As Long
Dim tTituloPagar As typeTituloPagar
Dim objNFPagTitPag As New ClassNfPagTitPagHic 'Inserido por Wagner
Dim lSituacao As Long
Dim lHistorico As Long
Dim lSubConta As Long

On Error GoTo Erro_TituloPagar_Le_Numero
    
    'Abre comando
    lComando = Comando_Abrir
    If lComando = 0 Then Error 18548
    
    tTituloPagar.sObservacao = String(STRING_TITULO_OBSERVACAO, 0)
    
    '##############################
    'Alterado por Wagner 18/02/2005
    'Seleciona TituloPagar a partir dos dados fornecidos
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, ValorTotal, ValorSeguro, ValorFrete, OutrasDespesas, ValorProdutos, ValorICMS, ValorICMSSubst, CreditoICMS, ValorIPI, CreditoIPI, ValorIRRF, ValorINSS,INSSRetido,DataRegistro, Observacao, Saldo, NumParcelas, Status, CondicaoPagto, FilialPedCompra, NumPedCompra, PISRetido, ISSRetido, COFINSRetido, CSLLRetido, TaxaMoeda, Situacao, Historico, SubConta FROM TitulosPag WHERE Fornecedor = ? AND Filial = ? AND NumTitulo = ? AND DataEmissao = ? AND SiglaDocumento = ? AND FilialEmpresa = ? AND Status <> ? ", _
        tTituloPagar.lNumIntDoc, tTituloPagar.dValorTotal, tTituloPagar.dValorSeguro, _
        tTituloPagar.dValorFrete, tTituloPagar.dOutrasDespesas, tTituloPagar.dValorProdutos, tTituloPagar.dValorICMS, tTituloPagar.dValorICMSSubst, tTituloPagar.iCreditoICMS, tTituloPagar.dValorIPI, tTituloPagar.iCreditoIPI, tTituloPagar.dValorIRRF, tTituloPagar.dValorINSS, tTituloPagar.iINSSRetido, tTituloPagar.dtDataRegistro, tTituloPagar.sObservacao, tTituloPagar.dSaldo, tTituloPagar.iNumParcelas, tTituloPagar.iStatus, tTituloPagar.iCondicaoPagto, tTituloPagar.iFilialPedCompra, tTituloPagar.lNumPedCompra, tTituloPagar.dPISRetido, tTituloPagar.dISSRetido, tTituloPagar.dCOFINSRetido, tTituloPagar.dCSLLRetido, tTituloPagar.dTaxaMoeda, lSituacao, lHistorico, lSubConta, _
        objTituloPagar.lFornecedor, objTituloPagar.iFilial, objTituloPagar.lNumTitulo, objTituloPagar.dtDataEmissao, objTituloPagar.sSiglaDocumento, giFilialEmpresa, STATUS_EXCLUIDO)
    If lErro <> AD_SQL_SUCESSO Then Error 18549
    '##############################
    
    'Tenta ler
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 18550
    'Se não encontrou --> Erro
    If lErro = AD_SQL_SEM_DADOS Then Error 18551
    
    'Preenche objTituloPagar com os dados lidos
    tTituloPagar.iFilialEmpresa = giFilialEmpresa
    tTituloPagar.lFornecedor = objTituloPagar.lFornecedor
    tTituloPagar.iFilial = objTituloPagar.iFilial
    tTituloPagar.lNumTitulo = objTituloPagar.lNumTitulo
    tTituloPagar.dtDataEmissao = objTituloPagar.dtDataEmissao
    tTituloPagar.sSiglaDocumento = objTituloPagar.sSiglaDocumento
    Call TituloPagar_MoverCamposObj(objTituloPagar, tTituloPagar)
    
    '###############################################
    'Inserido por Wagner 18/02/2005
    objNFPagTitPag.lSituacao = lSituacao
    objNFPagTitPag.lHistorico = lHistorico
    objNFPagTitPag.lSubConta = lSubConta
    objNFPagTitPag.sObservacao = tTituloPagar.sObservacao
    
    Set objTituloPagar.objInfoUsu = objNFPagTitPag
    '###############################################
    
    'Fecha comando
    Call Comando_Fechar(lComando)
    
    TituloPagar_Le_Numero = SUCESSO
    
    Exit Function
    
Erro_TituloPagar_Le_Numero:

    TituloPagar_Le_Numero = Err

    Select Case Err
    
        Case 18548
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 18549, 18550
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFFATURA", Err, objTituloPagar.lNumTitulo)
        
        Case 18551
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 179168)
            
    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function TituloPagarBaixado_Le_Numero(ByVal objTituloPagar As ClassTituloPagar) As Long
'Lê na Tabela de Titulos Baixados o Titulo da filial corrente com o Fornecedor, Filial, NumTitulo e Data de Emissao passados em objTituloPagar

Dim lErro As Long
Dim lComando As Long
Dim tTituloPagar As typeTituloPagar
Dim objNFPagTitPag As New ClassNfPagTitPagHic 'Inserido por Wagner
Dim lSituacao As Long
Dim lHistorico As Long
Dim lSubConta As Long

On Error GoTo Erro_TituloPagarBaixado_Le_Numero

    'Abre comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 18553
    
    tTituloPagar.sObservacao = String(STRING_TITULO_OBSERVACAO, 0)
    
    '##############################
    'Alterado por Wagner 18/02/2005
    'Seleciona TítuloPagarBaixado com os dados passados
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, ValorTotal, ValorSeguro, ValorFrete, OutrasDespesas, ValorProdutos, ValorICMS, ValorICMSSubst, CreditoICMS, ValorIPI, CreditoIPI, ValorIRRF, ValorINSS, INSSRetido, DataRegistro, Observacao, Saldo, NumParcelas, Status, CondicaoPagto, FilialPedCompra, NumPedCompra, PISRetido, ISSRetido, COFINSRetido, CSLLRetido, TaxaMoeda, Situacao, Historico, SubConta FROM TitulosPagBaixados WHERE Fornecedor = ? AND Filial = ? AND NumTitulo = ? AND DataEmissao = ? AND SiglaDocumento = ? AND FilialEmpresa = ? AND Status <> ? ", _
        tTituloPagar.lNumIntDoc, tTituloPagar.dValorTotal, tTituloPagar.dValorSeguro, tTituloPagar.dValorFrete, tTituloPagar.dOutrasDespesas, tTituloPagar.dValorProdutos, tTituloPagar.dValorICMS, tTituloPagar.dValorICMSSubst, tTituloPagar.iCreditoICMS, tTituloPagar.dValorIPI, tTituloPagar.iCreditoIPI, tTituloPagar.dValorIRRF, tTituloPagar.dValorINSS, tTituloPagar.iINSSRetido, tTituloPagar.dtDataRegistro, tTituloPagar.sObservacao, tTituloPagar.dSaldo, tTituloPagar.iNumParcelas, tTituloPagar.iStatus, tTituloPagar.iCondicaoPagto, _
        tTituloPagar.iFilialPedCompra, tTituloPagar.lNumPedCompra, tTituloPagar.dPISRetido, tTituloPagar.dISSRetido, tTituloPagar.dCOFINSRetido, tTituloPagar.dCSLLRetido, tTituloPagar.dTaxaMoeda, lSituacao, lHistorico, lSubConta, objTituloPagar.lFornecedor, objTituloPagar.iFilial, objTituloPagar.lNumTitulo, objTituloPagar.dtDataEmissao, objTituloPagar.sSiglaDocumento, giFilialEmpresa, STATUS_EXCLUIDO)
    If lErro <> AD_SQL_SUCESSO Then Error 18554
    '##############################
    
    'Tenta ler
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 18555
    
    'Se não encontrou --> Erro
    If lErro = AD_SQL_SEM_DADOS Then Error 18556
    
    'Preenche objTituloPagar com os dados do Título
    tTituloPagar.lFornecedor = objTituloPagar.lFornecedor
    tTituloPagar.iFilial = objTituloPagar.iFilial
    tTituloPagar.lNumTitulo = objTituloPagar.lNumTitulo
    tTituloPagar.dtDataEmissao = objTituloPagar.dtDataEmissao
    tTituloPagar.sSiglaDocumento = objTituloPagar.sSiglaDocumento
    tTituloPagar.iFilialEmpresa = giFilialEmpresa
    Call TituloPagar_MoverCamposObj(objTituloPagar, tTituloPagar)
    
    '###############################################
    'Inserido por Wagner 18/02/2005
    objNFPagTitPag.lSituacao = lSituacao
    objNFPagTitPag.lHistorico = lHistorico
    objNFPagTitPag.lSubConta = lSubConta
    objNFPagTitPag.sObservacao = objNFPagTitPag.sObservacao
    
    Set objTituloPagar.objInfoUsu = objNFPagTitPag
    '###############################################
    
    'Fecha comando
    Call Comando_Fechar(lComando)
    
    TituloPagarBaixado_Le_Numero = SUCESSO
    
    Exit Function

Erro_TituloPagarBaixado_Le_Numero:

    TituloPagarBaixado_Le_Numero = Err
    
    Select Case Err
    
        Case 18553
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 18554, 18555
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFFATBAIXADA", Err, objTituloPagar.lNumTitulo)
        
        Case 18556 'Título não encontrado
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 179169)
            
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function TituloPagarBaixado_Le(ByVal objTituloPagar As ClassTituloPagar) As Long
'Lê Titulo a Pagar Baixado com o Número Interno passado como parâmetro

Dim lErro As Long
Dim lComando As Long
Dim tTituloPagar As typeTituloPagar
Dim objNFPagTitPag As New ClassNfPagTitPagHic 'Inserido por Wagner
Dim lSituacao As Long
Dim lHistorico As Long
Dim lSubConta As Long

On Error GoTo Erro_TituloPagarBaixado_Le

    'Abre comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 56658
    
    tTituloPagar.sSiglaDocumento = String(STRING_SIGLA_DOCUMENTO, 0)
    tTituloPagar.sObservacao = String(STRING_TITULO_OBSERVACAO, 0)
    
    '##############################
    'Alterado por Wagner 18/02/2005
    'Seleciona o TituloPagar a partir do NumIntDoc
    lErro = Comando_Executar(lComando, "SELECT Fornecedor, Filial, SiglaDocumento, NumTitulo, DataEmissao, Saldo, NumParcelas, ValorTotal, ValorSeguro, ValorFrete, OutrasDespesas, ValorProdutos, ValorICMS, ValorICMSSubst, CreditoICMS, ValorIPI, CreditoIPI, ValorIRRF, ValorINSS, INSSRetido, DataRegistro, " & _
        "CondicaoPagto, Observacao, Status, FilialPedCompra, NumPedCompra, PISRetido, ISSRetido, COFINSRetido, CSLLRetido, TaxaMoeda, Situacao, Historico, SubConta FROM TitulosPagBaixados WHERE NumIntDoc = ? AND Status <> ?", tTituloPagar.lFornecedor, tTituloPagar.iFilial, tTituloPagar.sSiglaDocumento, tTituloPagar.lNumTitulo, tTituloPagar.dtDataEmissao, tTituloPagar.dSaldo, tTituloPagar.iNumParcelas, tTituloPagar.dValorTotal, tTituloPagar.dValorSeguro, tTituloPagar.dValorFrete, tTituloPagar.dOutrasDespesas, tTituloPagar.dValorProdutos, tTituloPagar.dValorICMS, tTituloPagar.dValorICMSSubst, tTituloPagar.iCreditoICMS, tTituloPagar.dValorIPI, tTituloPagar.iCreditoIPI, tTituloPagar.dValorIRRF, tTituloPagar.dValorINSS, tTituloPagar.iINSSRetido, tTituloPagar.dtDataRegistro, tTituloPagar.iCondicaoPagto, _
        tTituloPagar.sObservacao, tTituloPagar.iStatus, tTituloPagar.iFilialPedCompra, tTituloPagar.lNumPedCompra, tTituloPagar.dPISRetido, tTituloPagar.dISSRetido, tTituloPagar.dCOFINSRetido, tTituloPagar.dCSLLRetido, tTituloPagar.dTaxaMoeda, lSituacao, lHistorico, lSubConta, objTituloPagar.lNumIntDoc, STATUS_EXCLUIDO)
    If lErro <> AD_SQL_SUCESSO Then Error 56659
    '##############################
    
    'Tenta ler
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 56660
    
    'Se não encontrou --> Erro
    If lErro = AD_SQL_SEM_DADOS Then Error 56661
    
    'Preenche objTituloPagar com os dados do Título
    objTituloPagar.iFilialEmpresa = giFilialEmpresa
    objTituloPagar.lFornecedor = tTituloPagar.lFornecedor
    objTituloPagar.iFilial = tTituloPagar.iFilial
    objTituloPagar.sSiglaDocumento = tTituloPagar.sSiglaDocumento
    objTituloPagar.lNumTitulo = tTituloPagar.lNumTitulo
    objTituloPagar.dtDataEmissao = tTituloPagar.dtDataEmissao
    objTituloPagar.dSaldo = tTituloPagar.dSaldo
    objTituloPagar.iNumParcelas = tTituloPagar.iNumParcelas
    objTituloPagar.dValorTotal = tTituloPagar.dValorTotal
    objTituloPagar.dValorSeguro = tTituloPagar.dValorSeguro
    objTituloPagar.dValorFrete = tTituloPagar.dValorFrete
    objTituloPagar.dOutrasDespesas = tTituloPagar.dOutrasDespesas
    objTituloPagar.dValorProdutos = tTituloPagar.dValorProdutos
    objTituloPagar.dValorICMS = tTituloPagar.dValorICMS
    objTituloPagar.dValorICMSSubst = tTituloPagar.dValorICMSSubst
    objTituloPagar.iCreditoICMS = tTituloPagar.iCreditoICMS
    objTituloPagar.iCreditoIPI = tTituloPagar.iCreditoIPI
    objTituloPagar.dValorIPI = tTituloPagar.dValorIPI
    objTituloPagar.dValorIRRF = tTituloPagar.dValorIRRF
    objTituloPagar.dValorINSS = tTituloPagar.dValorINSS
    objTituloPagar.sObservacao = tTituloPagar.sObservacao
    objTituloPagar.iStatus = tTituloPagar.iStatus
    objTituloPagar.iFilialPedCompra = tTituloPagar.iFilialPedCompra
    objTituloPagar.lNumPedCompra = tTituloPagar.lNumPedCompra
    objTituloPagar.iINSSRetido = tTituloPagar.iINSSRetido
    objTituloPagar.dtDataRegistro = tTituloPagar.dtDataRegistro
    objTituloPagar.iCondicaoPagto = tTituloPagar.iCondicaoPagto
    objTituloPagar.dPISRetido = tTituloPagar.dPISRetido
    objTituloPagar.dISSRetido = tTituloPagar.dISSRetido
    objTituloPagar.dCOFINSRetido = tTituloPagar.dCOFINSRetido
    objTituloPagar.dCSLLRetido = tTituloPagar.dCSLLRetido
    objTituloPagar.dTaxaMoeda = tTituloPagar.dTaxaMoeda
    
    '###############################################
    'Inserido por Wagner 18/02/2005
    objNFPagTitPag.lSituacao = lSituacao
    objNFPagTitPag.lHistorico = lHistorico
    objNFPagTitPag.lSubConta = lSubConta
    objNFPagTitPag.sObservacao = tTituloPagar.sObservacao
    
    Set objTituloPagar.objInfoUsu = objNFPagTitPag
    '###############################################
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    TituloPagarBaixado_Le = SUCESSO
    
    Exit Function
    
Erro_TituloPagarBaixado_Le:

    TituloPagarBaixado_Le = Err
    
    Select Case Err
    
        Case 56658
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 56659, 56660
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOSPAGBAIXADOS", Err, objTituloPagar.lNumIntDoc)
        
        Case 56661  'Título não encontrado
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 179170)
    
    End Select
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Private Sub TituloPagar_MoverCamposObj(ByVal objTituloPagar As ClassTituloPagar, tTituloPagar As typeTituloPagar)
'copia informacoes de tTituloPagar para objTituloPagar

    objTituloPagar.lNumIntDoc = tTituloPagar.lNumIntDoc
    objTituloPagar.iFilialEmpresa = tTituloPagar.iFilialEmpresa
    objTituloPagar.lFornecedor = tTituloPagar.lFornecedor
    objTituloPagar.iFilial = tTituloPagar.iFilial
    objTituloPagar.sSiglaDocumento = tTituloPagar.sSiglaDocumento
    objTituloPagar.lNumTitulo = tTituloPagar.lNumTitulo
    objTituloPagar.dtDataEmissao = tTituloPagar.dtDataEmissao
    objTituloPagar.dValorTotal = tTituloPagar.dValorTotal
    objTituloPagar.dValorSeguro = tTituloPagar.dValorSeguro
    objTituloPagar.dValorFrete = tTituloPagar.dValorFrete
    objTituloPagar.dOutrasDespesas = tTituloPagar.dOutrasDespesas
    objTituloPagar.dValorProdutos = tTituloPagar.dValorProdutos
    objTituloPagar.dValorICMS = tTituloPagar.dValorICMS
    objTituloPagar.dValorICMSSubst = tTituloPagar.dValorICMSSubst
    objTituloPagar.iCreditoICMS = tTituloPagar.iCreditoICMS
    objTituloPagar.dValorIPI = tTituloPagar.dValorIPI
    objTituloPagar.iCreditoIPI = tTituloPagar.iCreditoIPI
    objTituloPagar.dValorIRRF = tTituloPagar.dValorIRRF
    objTituloPagar.dValorINSS = tTituloPagar.dValorINSS
    objTituloPagar.sObservacao = tTituloPagar.sObservacao
    objTituloPagar.dSaldo = tTituloPagar.dSaldo
    objTituloPagar.iNumParcelas = tTituloPagar.iNumParcelas
    objTituloPagar.iStatus = tTituloPagar.iStatus
    objTituloPagar.iCondicaoPagto = tTituloPagar.iCondicaoPagto
    objTituloPagar.iFilialPedCompra = tTituloPagar.iFilialPedCompra
    objTituloPagar.lNumPedCompra = tTituloPagar.lNumPedCompra
    objTituloPagar.iINSSRetido = tTituloPagar.iINSSRetido
    objTituloPagar.dtDataRegistro = tTituloPagar.dtDataRegistro
    objTituloPagar.dPISRetido = tTituloPagar.dPISRetido
    objTituloPagar.dISSRetido = tTituloPagar.dISSRetido
    objTituloPagar.dCOFINSRetido = tTituloPagar.dCOFINSRetido
    objTituloPagar.dCSLLRetido = tTituloPagar.dCSLLRetido
    objTituloPagar.dTaxaMoeda = tTituloPagar.dTaxaMoeda
    
End Sub

Sub ParcelasRec_Le_BorderoCobranca1(ByVal objBorderoCobrancaEmissao As ClassBorderoCobrancaEmissao, sSelect As String, iParamFilialEmpresa As Integer, iParamTipoCliente As Integer)
'auxiliar a ParcelasRec_Le_BorderoCobranca
'monta o SELECT dinamicamente
Dim sFrom As String, sWhere As String, sFields As String, sOrderBy As String

    sFields = " Clientes.NomeReduzido, Clientes.RazaoSocial, Clientes.CodPadraoCobranca"
    sFields = sFields & ", TitulosRec.NumTitulo, TitulosRec.FilialEmpresa, TitulosRec.Cliente, TitulosRec.Filial, TitulosRec.SiglaDocumento"
    sFields = sFields & ", ParcelasRec.NumIntDoc, ParcelasRec.NumParcela, ParcelasRec.DataVencimento, ParcelasRec.Saldo"

    sFrom = " FROM ParcelasRec, TitulosRec, Clientes, FiliaisClientes"
    sWhere = " WHERE ParcelasRec.DataVencimento <= ? AND ParcelasRec.Status = ? AND ParcelasRec.NumIntTitulo = TitulosRec.NumIntDoc AND TitulosRec.Cliente = Clientes.Codigo AND ParcelasRec.Cobrador = ? AND ParcelasRec.CarteiraCobranca = ? AND Clientes.Codigo = FiliaisClientes.CodCliente AND TitulosRec.Filial = FiliaisClientes.CodFilial AND ( FiliaisClientes.CodCobrador = ? OR FiliaisClientes.CodCobrador = 0 )"

    'Se entrar como filial centralizadora da cobranca pode pegar titulos de todas as filiais,
    If giFilialEmpresa = gobjCR.iFilialCentralizadora Then
        iParamFilialEmpresa = 0
    Else 'Senao pegar titulos somente da filial corrente
        sWhere = sWhere & " AND TitulosRec.FilialEmpresa = ?"
        iParamFilialEmpresa = 1
    End If

    If objBorderoCobrancaEmissao.sSiglaEstado <> "" Then
        sWhere = sWhere & " AND FiliaisClientes.EnderecoCobranca = Enderecos.Codigo AND Enderecos.SiglaEstado = '" & objBorderoCobrancaEmissao.sSiglaEstado & "'"
        sFrom = sFrom & ", Enderecos"
    End If
    
    If objBorderoCobrancaEmissao.iTipoCliente <> 0 Then
        sWhere = sWhere & " AND Clientes.Tipo = ?"
        iParamTipoCliente = 1
    Else
        iParamTipoCliente = 0
    End If
    
    sOrderBy = " ORDER BY TitulosRec.NumTitulo, ParcelasRec.NumParcela, TitulosRec.Cliente, TitulosRec.Filial" 'Alterado por Wagner
    
    sSelect = "SELECT " & sFields & sFrom & sWhere & sOrderBy

End Sub

'#############################################################
'Inserido por Wagner
Function ChequesPagAvulso_ChequesSelecionados(ByVal objChequesPagAvulso As ClassChequesPagAvulso) As Long
'utilizada na emissao de cheque manual (ou avulso)
'preenche informacoes sobre o cheque
'deve ser chamada apos a selecao das parcelas que efetivamente serao pagas em cheque
Dim objInfoParcPag As ClassInfoParcPag, objInfoChequePag As ClassInfoChequePag
Dim iPrimeiro As Integer, dValor As Double, sFavorecido As String, lNumRealCheque As Long, iPortador As Integer
Dim lComandoSelPort As Long, lComandoSelForn As Long, lErro As Long

'########################################
'Inserido por Wagner
Dim objTituloPag As New ClassTituloPagar
Dim objParcelaPagar As New ClassParcelaPagar
Dim objParcelaPagHic As New ClassParcelaPagHic
Dim objNFPagTitPag As New ClassNfPagTitPagHic
Dim lNumCheque As Long
Dim lComando As Long
'########################################

On Error GoTo Erro_ChequesPagAvulso_ChequesSelecionados

    lComandoSelPort = Comando_Abrir()
    If lComandoSelPort = 0 Then gError 7397
    
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 7397
    
    lComandoSelForn = Comando_Abrir()
    If lComandoSelForn = 0 Then gError 7398

    dValor = 0
    iPrimeiro = 1

    For Each objInfoParcPag In objChequesPagAvulso.colInfoParcPag

        'pula parcelas que nao serao pagas
        If objInfoParcPag.iSeqCheque <> 0 Then

            If iPrimeiro = 1 Then
                iPrimeiro = 0
                iPortador = objInfoParcPag.iPortador
                
                lErro = CF("FavorecidoChequeTitulo_Obter", sFavorecido, objInfoParcPag.lFornecedor, objInfoParcPag.iTipoCobranca, objInfoParcPag.sNomeRedPortador, lComandoSelPort, lComandoSelForn)
                If lErro <> SUCESSO Then gError 7399
                
                '#######################################
                'Inserido por Wagner
                Set objTituloPag = New ClassTituloPagar
                
                objParcelaPagar.lNumIntDoc = objInfoParcPag.lNumIntParc
                
                'Le o NumInterno do Titulo para passar no objTituloPag
                lErro = CF("ParcelaPagar_Le", objParcelaPagar)
                If lErro <> SUCESSO And lErro <> 60479 Then gError 132571
                
                'Se não encontrou a Parcela --> ERRO
                If lErro = 60479 Then gError 132572
                
                objTituloPag.lNumIntDoc = objParcelaPagar.lNumIntTitulo
                
                'Le os Dados do Titulo
                lErro = CF("TituloPagar_Le", objTituloPag)
                If lErro <> SUCESSO And lErro <> 18372 Then gError 132573
                
                If lErro = 18372 Then gError 132574
                
                If Not (objTituloPag.objInfoUsu Is Nothing) Then
                    Set objNFPagTitPag = objTituloPag.objInfoUsu
                End If
                
                'Se foi retido o favorecido na parcela, considera este
                If Not (objParcelaPagar.objInfoUsu Is Nothing) Then
                
                    Set objParcelaPagHic = objParcelaPagar.objInfoUsu
                    
                    If Len(Trim(objParcelaPagHic.sFavorecido)) <> 0 Then
                        sFavorecido = objParcelaPagHic.sFavorecido
                    End If
                
                End If
                
                lErro = Comando_Executar(lComando, "SELECT Cheque FROM ChequeHic WHERE NumIntParcela= ? AND ContaCorrente = ? ", lNumCheque, objParcelaPagar.lNumIntDoc, objChequesPagAvulso.iCta)
                If lErro <> AD_SQL_SUCESSO Then gError 132871
                
                lErro = Comando_BuscarPrimeiro(lComando)
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 132872
                '########################################
                
            End If

            dValor = dValor + objInfoParcPag.dValor + objInfoParcPag.dValorJuros + objInfoParcPag.dValorMulta - objInfoParcPag.dValorDesconto

        End If

    Next

    'incluir o cheque, se existir
    If iPrimeiro <> 1 Then
        objChequesPagAvulso.lNumCheque = lNumCheque 'Inserido por Wagner
        objChequesPagAvulso.objInfoChequePag.dValor = dValor
        objChequesPagAvulso.objInfoChequePag.iImprimir = 1
        objChequesPagAvulso.objInfoChequePag.iSeqCheque = 1
        objChequesPagAvulso.objInfoChequePag.lNumRealCheque = objChequesPagAvulso.lNumCheque
        objChequesPagAvulso.objInfoChequePag.sFavorecido = sFavorecido
        objChequesPagAvulso.objInfoChequePag.iPortador = iPortador
        objChequesPagAvulso.objInfoChequePag.sObservacao = objTituloPag.sObservacao 'Inserido por Wagner
    End If

    lErro = Comando_Fechar(lComandoSelForn)
    lErro = Comando_Fechar(lComandoSelPort)

    ChequesPagAvulso_ChequesSelecionados = SUCESSO

    Exit Function

Erro_ChequesPagAvulso_ChequesSelecionados:

    ChequesPagAvulso_ChequesSelecionados = gErr

    Select Case gErr

        Case 7397, 7398
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 7399
        
        '###############################
        'Inserido por Wagner
        Case 132571, 132573, 132871, 132872 'Tratado na rotina chamada
        
        Case 132572
            Call Rotina_Erro(vbOKOnly, "ERRO_PARCELA_PAGAR_INEXISTENTE", gErr)
        
        Case 132574
            Call Rotina_Erro(vbOKOnly, "ERRO_TITULO_PAGAR_INEXISTENTE", gErr)
        '###############################
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179171)

    End Select

    Call Comando_Fechar(lComandoSelForn)
    Call Comando_Fechar(lComandoSelPort)
    
    Exit Function

End Function

Function ChequesPag_ChequesSelecionados(ByVal objChequesPag As ClassChequesPag) As Long
'utilizada na emissao de cheques automaticos
'preenche colecao (em objChequesPag) com informacoes sobre os cheques
'deve ser chamada apos a selecao das parcelas que efetivamente serao pagas com os agrupamentos definitivos de parcelas em cheques
Dim objInfoParcPag As ClassInfoParcPag, objInfoChequePag As ClassInfoChequePag
Dim iSeqChequeAnterior As Integer, dValor As Double, sFavorecido As String, lNumRealCheque As Long, iPortadorAnterior As Integer
Dim lComandoSelPort As Long, lComandoSelForn As Long, lErro As Long

'#########################################
'Inserido por Wagner
Dim objTituloPag As ClassTituloPagar
Dim objParcelaPagar As New ClassParcelaPagar
Dim objParcelaPagHic As New ClassParcelaPagHic
Dim objNFPagTitPag As New ClassNfPagTitPagHic
'#########################################

On Error GoTo Erro_ChequesPag_ChequesSelecionados

    'p/usuario poder voltar do passo 3 p/o 2 alterar dados
    'e prosseguir, esvazio a colecao
    Do While objChequesPag.ColInfoChequePag.Count <> 0
        objChequesPag.ColInfoChequePag.Remove (1)
    Loop
    
    lComandoSelPort = Comando_Abrir()
    If lComandoSelPort = 0 Then gError 7390

    lComandoSelForn = Comando_Abrir()
    If lComandoSelForn = 0 Then gError 7391

    'flag para tratar o primeiro cheque
    iSeqChequeAnterior = -1

    For Each objInfoParcPag In objChequesPag.colInfoParcPag

        '#######################################
        'Inserido por Wagner
        'Le o Titulo e a parcela a Pagar relacionada ao objInfoParcPag
        Set objTituloPag = New ClassTituloPagar
        
        objParcelaPagar.lNumIntDoc = objInfoParcPag.lNumIntParc
        
        'Le o NumInterno do Titulo para passar no objTituloPag
        lErro = CF("ParcelaPagar_Le", objParcelaPagar)
        If lErro <> SUCESSO And lErro <> 60479 Then gError 132567
        
        'Se não encontrou a Parcela --> ERRO
        If lErro = 60479 Then gError 132568
        
        If Not (objParcelaPagar.objInfoUsu Is Nothing) Then
            Set objParcelaPagHic = objParcelaPagar.objInfoUsu
        End If
        
        objTituloPag.lNumIntDoc = objParcelaPagar.lNumIntTitulo
        
        'Le os Dados do Titulo
        lErro = CF("TituloPagar_Le", objTituloPag)
        If lErro <> SUCESSO And lErro <> 18372 Then gError 132569
        
        If lErro = 18372 Then gError 132570
        
        If Not (objTituloPag.objInfoUsu Is Nothing) Then
            Set objNFPagTitPag = objTituloPag.objInfoUsu
        End If
        '#######################################
        
        'pula parcelas que nao serao pagas
        If objInfoParcPag.iSeqCheque <> 0 Then

            If iSeqChequeAnterior = -1 Then
                'prepara para o primeiro cheque
                iSeqChequeAnterior = objInfoParcPag.iSeqCheque
                dValor = 0
                iPortadorAnterior = objInfoParcPag.iPortador
                
                lErro = CF("FavorecidoChequeTitulo_Obter", sFavorecido, objInfoParcPag.lFornecedor, objInfoParcPag.iTipoCobranca, objInfoParcPag.sNomeRedPortador, lComandoSelPort, lComandoSelForn)
                If lErro <> SUCESSO Then gError 7392

                lNumRealCheque = objChequesPag.lNumCheque
            End If

            'terminou de incluir todas as parcelas do cheque
            If iSeqChequeAnterior <> objInfoParcPag.iSeqCheque Then
                Call objChequesPag.ColInfoChequePag.Add(iSeqChequeAnterior, sFavorecido, dValor, lNumRealCheque, 1, iPortadorAnterior)

                '####################################
                'Inserido por Wagner
                'Obtém a observação do Título a pagar
                objChequesPag.ColInfoChequePag.Item(objChequesPag.ColInfoChequePag.Count).sObservacao = objTituloPag.sObservacao
                '####################################
                
                'prepara para o proximo cheque
                iSeqChequeAnterior = objInfoParcPag.iSeqCheque
                dValor = 0
                iPortadorAnterior = objInfoParcPag.iPortador

                lErro = CF("FavorecidoChequeTitulo_Obter", sFavorecido, objInfoParcPag.lFornecedor, objInfoParcPag.iTipoCobranca, objInfoParcPag.sNomeRedPortador, lComandoSelPort, lComandoSelForn)
                If lErro <> SUCESSO Then gError 7393
                
                '########################################
                'Inserido por Wagner
                'Se foi retido o favorecido na parcela, considera este
                If Len(Trim(objParcelaPagHic.sFavorecido)) <> 0 Then
                    sFavorecido = objParcelaPagHic.sFavorecido
                End If
                '########################################

                lNumRealCheque = lNumRealCheque + 1
            End If

            dValor = dValor + objInfoParcPag.dValor + objInfoParcPag.dValorJuros + objInfoParcPag.dValorMulta - objInfoParcPag.dValorDesconto

        End If

    Next

    'incluir o ultimo cheque, se existir
    If iSeqChequeAnterior <> -1 Then
        Call objChequesPag.ColInfoChequePag.Add(iSeqChequeAnterior, sFavorecido, dValor, lNumRealCheque, 1, iPortadorAnterior)
    
        '####################################
        'Inserido por Wagner
        objChequesPag.ColInfoChequePag.Item(objChequesPag.ColInfoChequePag.Count).sObservacao = objTituloPag.sObservacao
        '####################################
    
    End If

    lErro = Comando_Fechar(lComandoSelPort)
    lErro = Comando_Fechar(lComandoSelForn)

    ChequesPag_ChequesSelecionados = SUCESSO

    Exit Function

Erro_ChequesPag_ChequesSelecionados:

    ChequesPag_ChequesSelecionados = gErr

    Select Case gErr

        Case 7390, 7391
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 7392, 7393, 132567 'Inserido por Wagner

        '###############################
        'Inserido por Wagner
        Case 132567, 132569 'Tratado na rotina chamada
        
        Case 132568
            Call Rotina_Erro(vbOKOnly, "ERRO_PARCELA_PAGAR_INEXISTENTE", gErr)
        
        Case 132570
            Call Rotina_Erro(vbOKOnly, "ERRO_TITULO_PAGAR_INEXISTENTE", gErr)
        '###############################

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179172)

    End Select

    Call Comando_Fechar(lComandoSelPort)
    Call Comando_Fechar(lComandoSelForn)
    
    Exit Function

End Function
'#######################################################################

Function ParcelaPagar_Le(ByVal objParcelaPagar As ClassParcelaPagar) As Long
'Le os dados da Parcela com o NumIntDoc da Parcela

Dim lErro As Long
Dim lComando As Long
Dim tParcelaPagar As typeParcelaPagar
Dim sFavorecido 'Inserido por Wagner
Dim objParcelaPagHic As New ClassParcelaPagHic

On Error GoTo Erro_ParcelaPagar_Le

    'Abre o comando
    lComando = Comando_Abrir()
    If lErro <> SUCESSO Then Error 60476
    
    tParcelaPagar.sNossoNumero = String(STRING_NOSSO_NUMERO, 0)
    
    sFavorecido = String(STRING_FAVORECIDO, 0) 'Inserido por Wagner
    
    '##################################
    'Alterado por Wagner
    'Seleciona a Parcela com os dados passados
    lErro = Comando_Executar(lComando, "SELECT NumIntTitulo, Status, DataVencimento, DataVencimentoReal, Saldo, Valor, TipoCobranca, Portador, BancoCobrador, ProxSeqBaixa, NossoNumero, Favorecido FROM ParcelasPag WHERE NumIntDoc = ? ", tParcelaPagar.lNumIntTitulo, tParcelaPagar.iStatus, tParcelaPagar.dtDataVencimento, tParcelaPagar.dtDataVencimentoReal, tParcelaPagar.dSaldo, tParcelaPagar.dValor, tParcelaPagar.iTipoCobranca, tParcelaPagar.iPortador, tParcelaPagar.iBancoCobrador, tParcelaPagar.iProxSeqBaixa, tParcelaPagar.sNossoNumero, sFavorecido, objParcelaPagar.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then Error 60477
    '##################################
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 60478
    
    'Se não encontrar --> erro
    If lErro = AD_SQL_SEM_DADOS Then Error 60479
        
    'Carrega em objParcelaPagar os dados trazidos
    objParcelaPagar.lNumIntTitulo = tParcelaPagar.lNumIntTitulo
    objParcelaPagar.iStatus = tParcelaPagar.iStatus
    objParcelaPagar.dtDataVencimento = tParcelaPagar.dtDataVencimento
    objParcelaPagar.dtDataVencimentoReal = tParcelaPagar.dtDataVencimentoReal
    objParcelaPagar.dSaldo = tParcelaPagar.dSaldo
    objParcelaPagar.dValor = tParcelaPagar.dValor
    objParcelaPagar.iTipoCobranca = tParcelaPagar.iTipoCobranca
    objParcelaPagar.iPortador = tParcelaPagar.iPortador
    objParcelaPagar.iBancoCobrador = tParcelaPagar.iBancoCobrador
    objParcelaPagar.iProxSeqBaixa = tParcelaPagar.iProxSeqBaixa
    objParcelaPagar.sNossoNumero = tParcelaPagar.sNossoNumero
    
    '####################################
    'Inserido por Wagner
    objParcelaPagHic.sFavorecido = sFavorecido
    
    Set objParcelaPagar.objInfoUsu = objParcelaPagHic
    '####################################
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    ParcelaPagar_Le = SUCESSO
    
    Exit Function
    
Erro_ParcelaPagar_Le:

    ParcelaPagar_Le = Err
    
    Select Case Err
    
        Case 60476
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 60477, 60478
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PARCELASPAG2", Err)

        Case 60479 'Não encontrou a Parcela, Tratar na Rotina chamadora
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 179173)
        
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function TituloPagar_Le1(ByVal objTituloPagar As ClassTituloPagar) As Long
'Lê Titulo a Pagar com o Número Interno passado como parâmetro inclusive os excluidos logicamente (Status=STATUS_EXCLUIDO)

Dim lErro As Long
Dim lComando As Long
Dim tTituloPagar As typeTituloPagar
Dim objNFPagTitPag As New ClassNfPagTitPagHic 'Inserido por Wagner
Dim iChequeRetido As Integer 'Inserido por Wagner
Dim lSituacao As Long
Dim lHistorico As Long
Dim lSubConta As Long

On Error GoTo Erro_TituloPagar_Le1

    'Abre comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 89119
    
    tTituloPagar.sSiglaDocumento = String(STRING_SIGLA_DOCUMENTO, 0)
    tTituloPagar.sObservacao = String(STRING_TITULO_OBSERVACAO, 0)
    
    'Seleciona o TituloPagar a partir do NumIntDoc
    lErro = Comando_Executar(lComando, "SELECT FilialEmpresa, Fornecedor, Filial, SiglaDocumento, NumTitulo, DataEmissao, Saldo, NumParcelas, ValorTotal, ValorSeguro, ValorFrete, OutrasDespesas, ValorProdutos, ValorICMS, ValorICMSSubst, CreditoICMS, ValorIPI, CreditoIPI, ValorIRRF, ValorINSS, INSSRetido, DataRegistro, " & _
        "CondicaoPagto, Observacao, Status, FilialPedCompra, NumPedCompra, PISRetido, ISSRetido, COFINSRetido, CSLLRetido, TaxaMoeda, Situacao, Historico, SubConta, ChequeRetido FROM TitulosPag WHERE NumIntDoc = ?", tTituloPagar.iFilialEmpresa, tTituloPagar.lFornecedor, tTituloPagar.iFilial, tTituloPagar.sSiglaDocumento, tTituloPagar.lNumTitulo, tTituloPagar.dtDataEmissao, tTituloPagar.dSaldo, tTituloPagar.iNumParcelas, tTituloPagar.dValorTotal, tTituloPagar.dValorSeguro, tTituloPagar.dValorFrete, tTituloPagar.dOutrasDespesas, tTituloPagar.dValorProdutos, tTituloPagar.dValorICMS, tTituloPagar.dValorICMSSubst, tTituloPagar.iCreditoICMS, tTituloPagar.dValorIPI, tTituloPagar.iCreditoIPI, tTituloPagar.dValorIRRF, tTituloPagar.dValorINSS, tTituloPagar.iINSSRetido, tTituloPagar.dtDataRegistro, tTituloPagar.iCondicaoPagto, _
        tTituloPagar.sObservacao, tTituloPagar.iStatus, tTituloPagar.iFilialPedCompra, tTituloPagar.lNumPedCompra, tTituloPagar.dPISRetido, tTituloPagar.dISSRetido, tTituloPagar.dCOFINSRetido, tTituloPagar.dCSLLRetido, tTituloPagar.dTaxaMoeda, lSituacao, lHistorico, lSubConta, iChequeRetido, objTituloPagar.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 89120
    
    'Tenta ler
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89121
    
    'Se não encontrou --> Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 89122
    
    tTituloPagar.lNumIntDoc = objTituloPagar.lNumIntDoc
    
    'Preenche objTituloPagar com os dados do Título
    Call TituloPagar_MoverCamposObj(objTituloPagar, tTituloPagar)

    '###############################################
    'Inserido por Wagner 18/02/2005
    objNFPagTitPag.lSituacao = lSituacao
    objNFPagTitPag.lHistorico = lHistorico
    objNFPagTitPag.lSubConta = lSubConta
    objNFPagTitPag.sObservacao = tTituloPagar.sObservacao
    objNFPagTitPag.iChequeRetido = iChequeRetido '19/04/2005
    
    Set objTituloPagar.objInfoUsu = objNFPagTitPag
    '###############################################
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    TituloPagar_Le1 = SUCESSO
    
    Exit Function
    
Erro_TituloPagar_Le1:

    TituloPagar_Le1 = gErr
    
    Select Case gErr
    
        Case 89119
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 89120, 89121
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOSPAG", gErr, objTituloPagar.lNumIntDoc)
        
        Case 89122  'Título não encontrado
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 148732)
    
    End Select
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function


Function TituloPagar_Le_Titulo(ByVal objTituloPagar As ClassTituloPagar) As Long
'Lê o Título com o Fornecedor, Filial e Número do Título passados em objTituloPagar
'como pode haver mais de um titulo vou pegar o que foi registrado mais recentemente

Dim lErro As Long
Dim lComando As Long
Dim tTituloPagar As typeTituloPagar
Dim objNFPagTitPag As New ClassNfPagTitPagHic 'Inserido por Wagner
Dim iChequeRetido As Integer 'Inserido por Wagner
Dim lSituacao As Long
Dim lHistorico As Long
Dim lSubConta As Long

On Error GoTo Erro_TituloPagar_Le_Titulo
    
    'Abre comando
    lComando = Comando_Abrir
    If lComando = 0 Then Error 43448
    
    tTituloPagar.sObservacao = String(STRING_TITULO_OBSERVACAO, 0)
    tTituloPagar.sSiglaDocumento = String(STRING_SIGLA_DOCUMENTO, 0)
    
    'Seleciona TituloPagar a partir dos dados fornecidos
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, SiglaDocumento, DataEmissao, Status, Saldo, NumParcelas, ValorTotal, ValorSeguro, ValorFrete, OutrasDespesas, ValorProdutos, ValorICMS, ValorICMSSubst, CreditoICMS, ValorIPI, CreditoIPI, ValorIRRF, ValorINSS, INSSRetido, Observacao, FilialPedCompra, NumPedCompra, DataRegistro, CondicaoPagto, PISRetido, ISSRetido, COFINSRetido, CSLLRetido, TaxaMoeda, Situacao, Historico, SubConta, ChequeRetido FROM TitulosPag WHERE Fornecedor = ? AND Filial = ? AND NumTitulo = ? AND FilialEmpresa = ? AND Status <> ? ORDER BY DataRegistro DESC", tTituloPagar.lNumIntDoc, tTituloPagar.sSiglaDocumento, tTituloPagar.dtDataEmissao, _
        tTituloPagar.iStatus, tTituloPagar.dSaldo, tTituloPagar.iNumParcelas, tTituloPagar.dValorTotal, tTituloPagar.dValorSeguro, tTituloPagar.dValorFrete, tTituloPagar.dOutrasDespesas, tTituloPagar.dValorProdutos, tTituloPagar.dValorICMS, tTituloPagar.dValorICMSSubst, tTituloPagar.iCreditoICMS, tTituloPagar.dValorIPI, tTituloPagar.iCreditoIPI, tTituloPagar.dValorIRRF, tTituloPagar.dValorINSS, tTituloPagar.iINSSRetido, tTituloPagar.sObservacao, tTituloPagar.iFilialPedCompra, tTituloPagar.lNumPedCompra, tTituloPagar.dtDataRegistro, tTituloPagar.iCondicaoPagto, tTituloPagar.dPISRetido, tTituloPagar.dISSRetido, tTituloPagar.dCOFINSRetido, tTituloPagar.dCSLLRetido, tTituloPagar.dTaxaMoeda, lSituacao, lHistorico, lSubConta, iChequeRetido, _
        objTituloPagar.lFornecedor, objTituloPagar.iFilial, objTituloPagar.lNumTitulo, giFilialEmpresa, STATUS_EXCLUIDO)
    If lErro <> AD_SQL_SUCESSO Then Error 43449
    
    'Tenta ler
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 18550
    
    'Se não encontrou --> Erro
    If lErro = AD_SQL_SEM_DADOS Then Error 43451
    
    tTituloPagar.iFilialEmpresa = giFilialEmpresa
    tTituloPagar.lFornecedor = objTituloPagar.lFornecedor
    tTituloPagar.iFilial = objTituloPagar.iFilial
    tTituloPagar.lNumTitulo = objTituloPagar.lNumTitulo
    
    'Preenche objTituloPagar com os dados lidos
    Call TituloPagar_MoverCamposObj(objTituloPagar, tTituloPagar)
    
    '###############################################
    'Inserido por Wagner 18/02/2005
    objNFPagTitPag.lSituacao = lSituacao
    objNFPagTitPag.lHistorico = lHistorico
    objNFPagTitPag.lSubConta = lSubConta
    objNFPagTitPag.sObservacao = tTituloPagar.sObservacao
    objNFPagTitPag.iChequeRetido = iChequeRetido '19/04/2005
    
    Set objTituloPagar.objInfoUsu = objNFPagTitPag
    '###############################################
    
    'Fecha comando
    Call Comando_Fechar(lComando)
    
    TituloPagar_Le_Titulo = SUCESSO
    
    Exit Function
    
Erro_TituloPagar_Le_Titulo:

    TituloPagar_Le_Titulo = Err

    Select Case Err
    
        Case 43448
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 43449, 43450
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOSPAG2", Err, objTituloPagar.lFornecedor, objTituloPagar.iFilial)
        
        Case 43451
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 148639)
            
    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function


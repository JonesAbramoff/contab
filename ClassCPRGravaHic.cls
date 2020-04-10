VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCPRGrava"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'##############################################
'Inserido por Wagner
Public Function Geracao_NomeArquivo_CobrancaEletronica(ByVal objCobrador As ClassCobrador, ByVal sDiretorio As String, sNomeArquivo As String, lSequencialArquivo As Long)

Dim lErro As Long
Dim lSequencialArq As Long
Dim sNomeAux As String
Dim bGerou As Boolean

On Error GoTo Erro_Geracao_NomeArquivo_CobrancaEletronica

    bGerou = False
    lSequencialArq = 0

    Do While Not bGerou

        lSequencialArq = lSequencialArq + 1
        
        Select Case objCobrador.iCodBanco
        
            Case BANCO_BB
            
                sNomeAux = "CBR" & Format(gdtDataHoje, "DDMM") & "." & FormataCpoNum(lSequencialArq, 3)
            
            Case BANCO_ITAU
            
                sNomeAux = "D" & Format(gdtDataHoje, "DDMMYY") & "A.TXT"
                bGerou = True 'Não tem sequencial
            
            Case BANCO_BICBANCO
            
                sNomeAux = "E1493721." & Format(gdtDataHoje, "DD") & FormataCpoNum(lSequencialArq, 1)
            
            Case BANCO_RURAL
                
                sNomeAux = "HICA" & Format(gdtDataHoje, "DDMM") & ".REM"
                bGerou = True 'Não tem sequencial
            
            Case Else
        
                sNomeAux = "CB" & Format(gdtDataHoje, "DDMM") & FormataCpoNum(lSequencialArq, 2) & ".REM"
    
        End Select
    
        If Len(Trim(Dir(sDiretorio & "\" & sNomeAux))) = 0 Then bGerou = True

    Loop

    lSequencialArquivo = lSequencialArq

    sNomeArquivo = sDiretorio & "\" & sNomeAux

    Geracao_NomeArquivo_CobrancaEletronica = SUCESSO

    Exit Function

Erro_Geracao_NomeArquivo_CobrancaEletronica:

    Geracao_NomeArquivo_CobrancaEletronica = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179151)

    End Select

    Exit Function
    
End Function
'##############################################


'#########################################################################
'Inserido por Wagner 18/02/2005
'Nessa rotinas foram inseridos a gravação dos campos Situação, Historico, SubConta e Observação.
Function NFPag_Grava_BD(ByVal objNFsPag As ClassNFsPag, Optional ByVal iOrigemNFsPag As Integer = NFPAG_CONTASAPAGAR) As Long
'Grava Nota Fiscal a Pagar no BD
'Chamada por NFPag_Grava DENTRO DE TANSAÇÃO
'O Parâmetro iOrigemNFsPag serve p/ identificar o documento que esta gerando a gravação.Incluido por Leo em 20/02/02
' 0 = Oriundo de Contas a Pagar, 1 = Inserção NotaFiscal, 2 = Alteração NotaFiscal
'Const NFPAG_CONTASAPAGAR = 0
'Const NFPAG_NFINSERCAO = 1
'Const NFPAG_NFALTERACAO = 2

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim lComando2 As Long
Dim iFilialEmpresa As Integer
Dim lNumIntDoc As Long
Dim tNFsPag As typeNFsPag
Dim vbMsgRes As VbMsgBoxResult
Dim dtDataEmissao As Date
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim objNFPagTitPagHic As New ClassNfPagTitPagHic 'Inserido por Wagner

On Error GoTo Erro_NFPag_Grava_BD

    lComando = 0
    lComando1 = 0
    
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 18430
    
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 18431
    
    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then gError 94358
    
    '###############################################
    'Inserido por Wagner
    If Not (objNFsPag.objInfoUsu Is Nothing) Then
        Set objNFPagTitPagHic = objNFsPag.objInfoUsu
    End If
    '###############################################
        
    'Trecho incluido por leo em 25/02/02
    
    'Procura em NFsPagBaixadas se existe alguma Nota Fiscal com o índice fornecido pela Tela - Fornecedor, Filial, Numero, DataEmissao
    lErro = Comando_ExecutarPos(lComando2, "SELECT NumIntDoc,Fornecedor,Filial,ValorTotal,ValorIRRF, PISRetido, COFINSRetido, CSLLRetido, TaxaMoeda FROM NfsPagBaixadas WHERE Fornecedor = ? AND Filial = ? AND NumNotaFiscal = ? AND DataEmissao = ?", 0, tNFsPag.lNumIntDoc, tNFsPag.lFornecedor, tNFsPag.iFilial, tNFsPag.dValorTotal, tNFsPag.dValorIRRF, tNFsPag.dPISRetido, tNFsPag.dCOFINSRetido, tNFsPag.dCSLLRetido, tNFsPag.dTaxaMoeda, objNFsPag.lFornecedor, objNFsPag.iFilial, objNFsPag.lNumNotaFiscal, objNFsPag.dtDataEmissao)
    If lErro <> AD_SQL_SUCESSO Then gError 94360
    
    lErro = Comando_BuscarPrimeiro(lComando2)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 94361
    
    'Se a nota em questão esta baixada ...
    If lErro = AD_SQL_SUCESSO Then
        
        '... e veio do cp
        If iOrigemNFsPag = NFPAG_CONTASAPAGAR Then
        
            'poe dados necessarios para gravar ctb no obj
            objNFsPag.lNumIntDoc = tNFsPag.lNumIntDoc
            objNFsPag.lFornecedor = tNFsPag.lFornecedor
            objNFsPag.iFilial = tNFsPag.iFilial
            objNFsPag.dValorTotal = tNFsPag.dValorTotal
            objNFsPag.dValorIRRF = tNFsPag.dValorIRRF
            objNFsPag.dPISRetido = tNFsPag.dPISRetido
            objNFsPag.dCOFINSRetido = tNFsPag.dCOFINSRetido
            objNFsPag.dCSLLRetido = tNFsPag.dCSLLRetido
            objNFsPag.dTaxaMoeda = tNFsPag.dTaxaMoeda
            
            '############################################################
            'Alterado por Wagner
            'exibe aviso de alteracao somente da ctb...
            'If Rotina_Aviso(vbYesNo, "AVISO_TITBAIXADO_GRAVA_CONTAB_SOMENTE") = vbNo Then gError 120100
            
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CAMPOS_ALTERAVEIS_NFPAGBAIXADAS") 'Mensagem alterada por Wagner
            If vbMsgRes = vbNo Then gError 18315
        
            'Altera no BD
            lErro = Comando_ExecutarPos(lComando1, "UPDATE NfsPagBaixadas SET Situacao = ?, Historico = ?, SubConta = ?, Observacao = ?", lComando2, objNFPagTitPagHic.lSituacao, objNFPagTitPagHic.lHistorico, objNFPagTitPagHic.lSubConta, objNFPagTitPagHic.sObservacao)
            If lErro <> AD_SQL_SUCESSO Then gError 18316
            '############################################################
            
            'fecha os comandos
            Call Comando_Fechar(lComando)
            Call Comando_Fechar(lComando1)
            Call Comando_Fechar(lComando2)
            
            'retorna sucesso
            NFPag_Grava_BD = SUCESSO
            
            Exit Function
            
        'senao, mantem o tratamento que era usado anteriormente (msg de erro...)
        Else
        
            gError 94362
            
        End If
    
    
    End If
        
    tNFsPag.lNumIntDoc = 0
            
    'Leo até aqui
            
    'Procura em NFsPag se existe alguma Nota Fiscal com o índice fornecido pela Tela - Fornecedor, Filial, Numero, DataEmissao
    lErro = Comando_ExecutarPos(lComando, "SELECT NumIntDoc, FilialEmpresa, DataVencimento, ValorTotal, ValorSeguro, ValorFrete, OutrasDespesas, ValorProdutos, ValorIRRF, ValorICMS, ValorICMSSubst, CreditoICMS, ValorIPI, CreditoIPI, PISRetido, COFINSRetido, CSLLRetido, TaxaMoeda FROM NfsPag WHERE Fornecedor = ? AND Filial = ? AND NumNotaFiscal = ? AND DataEmissao = ? AND Status <> ? ", 0, tNFsPag.lNumIntDoc, tNFsPag.iFilialEmpresa, tNFsPag.dtDataVencimento, tNFsPag.dValorTotal, tNFsPag.dValorSeguro, tNFsPag.dValorFrete, tNFsPag.dOutrasDespesas, tNFsPag.dValorProdutos, tNFsPag.dValorIRRF, tNFsPag.dValorICMS, tNFsPag.dValorICMSSubst, tNFsPag.iCreditoICMS, tNFsPag.dValorIPI, tNFsPag.iCreditoIPI, tNFsPag.dPISRetido, tNFsPag.dCOFINSRetido, tNFsPag.dCSLLRetido, tNFsPag.dTaxaMoeda, objNFsPag.lFornecedor, objNFsPag.iFilial, objNFsPag.lNumNotaFiscal, objNFsPag.dtDataEmissao, STATUS_EXCLUIDO)
    If lErro <> AD_SQL_SUCESSO Then gError 18302
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 18303
    
    'Procura em NFsPag se existe alguma Nota Fiscal com o índice fornecido pela Tela - Fornecedor, Filial, Numero, DataEmissao
    lErro = Comando_ExecutarPos(lComando, "SELECT NumIntDoc, FilialEmpresa, DataVencimento, ValorTotal, ValorSeguro, ValorFrete, OutrasDespesas, ValorProdutos, ValorIRRF, ValorICMS, ValorICMSSubst, CreditoICMS, ValorIPI, CreditoIPI, PISRetido, COFINSRetido, CSLLRetido, TaxaMoeda FROM NfsPag WHERE Fornecedor = ? AND Filial = ? AND NumNotaFiscal = ? AND DataEmissao = ? AND Status <> ? ", 0, tNFsPag.lNumIntDoc, tNFsPag.iFilialEmpresa, tNFsPag.dtDataVencimento, tNFsPag.dValorTotal, tNFsPag.dValorSeguro, tNFsPag.dValorFrete, tNFsPag.dOutrasDespesas, tNFsPag.dValorProdutos, tNFsPag.dValorIRRF, tNFsPag.dValorICMS, tNFsPag.dValorICMSSubst, tNFsPag.iCreditoICMS, tNFsPag.dValorIPI, tNFsPag.iCreditoIPI, tNFsPag.dPISRetido, tNFsPag.dCOFINSRetido, tNFsPag.dCSLLRetido, tNFsPag.dTaxaMoeda, objNFsPag.lFornecedor, objNFsPag.iFilial, objNFsPag.lNumNotaFiscal, objNFsPag.dtDataEmissao, STATUS_EXCLUIDO)
    If lErro <> AD_SQL_SUCESSO Then gError 18302
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 18303
    
    'Se existir
    If lErro = AD_SQL_SUCESSO Then
        
        'Se é uma Inserção de Nota Fiscal, Erro. Nota Fiscal já existe.
        If iOrigemNFsPag = NFPAG_NFINSERCAO Then gError 94330 'Incluido por Leo em 20/02/02
        
        'Verifica se algum campo não alterável foi modificado
        If tNFsPag.dValorTotal <> objNFsPag.dValorTotal Or tNFsPag.dValorSeguro <> objNFsPag.dValorSeguro Or tNFsPag.dValorFrete <> objNFsPag.dValorFrete Or tNFsPag.dOutrasDespesas <> objNFsPag.dOutrasDespesas Or _
           tNFsPag.dValorProdutos <> objNFsPag.dValorProdutos Or tNFsPag.dValorIRRF <> objNFsPag.dValorIRRF Or tNFsPag.dValorICMS <> objNFsPag.dValorICMS Or tNFsPag.dValorICMSSubst <> objNFsPag.dValorICMSSubst Or _
           tNFsPag.iCreditoICMS <> objNFsPag.iCreditoICMS Or tNFsPag.dValorIPI <> tNFsPag.dValorIPI Or tNFsPag.iCreditoIPI <> objNFsPag.iCreditoIPI Or _
           tNFsPag.dPISRetido <> objNFsPag.dPISRetido Or tNFsPag.dCOFINSRetido <> objNFsPag.dCOFINSRetido Or tNFsPag.dCSLLRetido <> objNFsPag.dCSLLRetido _
           Then
           
           'Avisar que só a data de vencimento será alterada
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CAMPOS_ALTERAVEIS_NFPAG") 'Mensagem alterada por Wagner
            If vbMsgRes = vbNo Then gError 18315
           
        End If
   
        '#######################################
        'Alterado por Wagner 18/02/2005
        'Altera no BD
        lErro = Comando_ExecutarPos(lComando1, "UPDATE NfsPag SET DataVencimento =?, Situacao = ?, Historico = ?, SubConta = ?, Observacao = ?", lComando, objNFsPag.dtDataVencimento, objNFPagTitPagHic.lSituacao, objNFPagTitPagHic.lHistorico, objNFPagTitPagHic.lSubConta, objNFPagTitPagHic.sObservacao)
        If lErro <> AD_SQL_SUCESSO Then gError 18316
        '#######################################
    
        objNFsPag.lNumIntDoc = tNFsPag.lNumIntDoc

    'Se não encontrar
    ElseIf lErro = AD_SQL_SEM_DADOS Then
    
        'Verifica se a Nota Fiscal está baixada
        lErro = Comando_Executar(lComando, "SELECT FilialEmpresa FROM NfsPagBaixadas WHERE Fornecedor = ? AND Filial = ? AND NumNotaFiscal = ? AND DataEmissao = ? AND Status <> ? ", iFilialEmpresa, objNFsPag.lFornecedor, objNFsPag.iFilial, objNFsPag.lNumNotaFiscal, objNFsPag.dtDataEmissao, STATUS_EXCLUIDO)
        If lErro <> AD_SQL_SUCESSO Then gError 18305
        
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 18306
   
        'Se estiver erro
        If lErro = AD_SQL_SUCESSO Then gError 18307
        
        'faz verificacoes p/evitar entrada em duplicidade
        lErro = CF("NFPag_Testa_Repeticao", objNFsPag)
        If lErro <> SUCESSO Then gError 18530
    
        objFilialFornecedor.lCodFornecedor = objNFsPag.lFornecedor
        objFilialFornecedor.iCodFilial = objNFsPag.iFilial
        
        'Lê a Filial Fornecedor
        lErro = CF("FilialFornecedor_Lock1", objFilialFornecedor)
        If lErro <> SUCESSO And lErro <> 18900 Then gError 58633
    
        'Se não encontrar Filial Fornecedor --> Erro
        If lErro = 18900 Then gError 58634

        'Gera um novo número interno
        lErro = CF("NFPag_Automatico", lNumIntDoc)
        If lErro <> SUCESSO Then gError 18308

        objNFsPag.lNumIntDoc = lNumIntDoc
        If objNFsPag.iFilialEmpresa = EMPRESA_TODA Then objNFsPag.iFilialEmpresa = giFilialEmpresa
        objNFsPag.iStatus = STATUS_LANCADO
        
        '#######################################
        'Alterado por Wagner 18/02/2005
        'Inserir o novo registro em NfsPag
        lErro = Comando_Executar(lComando, "INSERT INTO NfsPag (NumIntDoc,FilialEmpresa,Fornecedor,Filial,NumNotaFiscal,DataEmissao,Status,DataVencimento,ValorTotal,ValorSeguro, ValorFrete, OutrasDespesas, ValorProdutos, ValorIRRF, ValorICMS, ValorICMSSUbst, CreditoICMS, ValorIPI, CreditoIPI,NumIntTitPag, PISRetido, COFINSRetido, CSLLRetido, TaxaMoeda, Situacao, Historico, SubConta, Observacao) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
            objNFsPag.lNumIntDoc, objNFsPag.iFilialEmpresa, objNFsPag.lFornecedor, objNFsPag.iFilial, objNFsPag.lNumNotaFiscal, objNFsPag.dtDataEmissao, objNFsPag.iStatus, objNFsPag.dtDataVencimento, objNFsPag.dValorTotal, objNFsPag.dValorSeguro, objNFsPag.dValorFrete, objNFsPag.dOutrasDespesas, objNFsPag.dValorProdutos, objNFsPag.dValorIRRF, objNFsPag.dValorICMS, objNFsPag.dValorICMSSubst, objNFsPag.iCreditoICMS, objNFsPag.dValorIPI, objNFsPag.iCreditoIPI, objNFsPag.lNumIntTitPag, objNFsPag.dPISRetido, objNFsPag.dCOFINSRetido, objNFsPag.dCSLLRetido, objNFsPag.dTaxaMoeda, objNFPagTitPagHic.lSituacao, objNFPagTitPagHic.lHistorico, objNFPagTitPagHic.lSubConta, objNFPagTitPagHic.sObservacao)
        If lErro <> AD_SQL_SUCESSO Then gError 18317
        '#######################################

    End If
   
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    
    NFPag_Grava_BD = SUCESSO
    
    Exit Function

Erro_NFPag_Grava_BD:

    NFPag_Grava_BD = gErr
    
    Select Case gErr
        
        Case 18302, 18303
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFSPAG1", gErr, objNFsPag.lNumNotaFiscal)
            
        Case 18305, 18306, 94360, 94361
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFSPAGBAIXADAS", gErr, objNFsPag.lNumNotaFiscal)
            
        Case 18307
            Call Rotina_Erro(vbOKOnly, "ERRO_NF_BAIXADA_MODIFICACAO", gErr, objNFsPag.lNumNotaFiscal)
            
        Case 18315, 18308, 18530, 58633, 120100 'Tratados pelas rotinas chamadas
            
        Case 18316
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_NFSPAG", gErr, objNFsPag.lNumNotaFiscal)
            
        Case 18317
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_NFSPAG", gErr, objNFsPag.lNumNotaFiscal)
       
        Case 18430, 18431, 94358, 94359
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 58634
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_CADASTRADA", gErr, objNFsPag.lFornecedor, objNFsPag.iFilial)
        
        Case 94330
            Call Rotina_Erro(vbOKOnly, "ERRO_NFSPAG_EXISTENTE", gErr, objNFsPag.lNumNotaFiscal)
                       
        Case 94362
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_BAIXADA_NFSPAGBAIXADAS", gErr, objNFsPag.lNumNotaFiscal)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179152)
                
    End Select
        
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    
    Exit Function
        
End Function

Function NFFatPag_Grava_BD(ByVal objTituloPagar As ClassTituloPagar, ByVal colParcelaPagar As colParcelaPagar, Optional ByVal bPodeAlterar As Boolean = True) As Long
'Grava NotaFiscalFatura (e outros titulos, menos fatura que cobra nfs simples ) no BD (nível + baixo)
'CHAMADA EM TRANSAÇÃO por NFFatPag_Grava

Dim lErro As Long
Dim lComando As Long, lComando1 As Long
Dim lComando2 As Long, lComando3 As Long
Dim tTituloPagar As typeTituloPagar
Dim iIndice As Integer
Dim colParcelaPagarBD As New colParcelaPagar
Dim objParcelaPagar As ClassParcelaPagar
Dim objParcelaPagar1 As ClassParcelaPagar
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim vbMsgRes As VbMsgBoxResult
Dim iCamposAlterados As Integer
Dim objNFPagTitPagHic As New ClassNfPagTitPagHic 'Inserido por Wagner
Dim lSituacao As Long
Dim lHistorico As Long
Dim lSubConta As Long

On Error GoTo Erro_NFFatPag_Grava_BD
    
    'Abre os comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 18471

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 18472
    
    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then gError 48919
    
    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then gError 59066
    
    '###############################################
    'Inserido por Wagner
    If Not (objTituloPagar.objInfoUsu Is Nothing) Then
        Set objNFPagTitPagHic = objTituloPagar.objInfoUsu
    End If
    '###############################################
    
    'Verifica se a NotaFiscalFatura já existe em TitulosPag
    lErro = Comando_ExecutarPos(lComando, "SELECT FilialEmpresa, NumIntDoc, ValorTotal, ValorICMS, CreditoICMS, ValorICMSSubst, ValorFrete, ValorINSS, ValorIPI, CreditoIPI, OutrasDespesas , ValorProdutos, ValorIRRF, ValorSeguro, ValorINSS, INSSRetido, FilialPedCompra, NumPedCompra, PISRetido, ISSRetido, COFINSRetido, CSLLRetido, TaxaMoeda, Situacao, Historico, SubConta FROM TitulosPag WHERE Fornecedor = ? AND Filial = ? AND NumTitulo = ? AND DataEmissao = ? AND SiglaDocumento = ? AND Status <> ?", 0, _
        tTituloPagar.iFilialEmpresa, tTituloPagar.lNumIntDoc, tTituloPagar.dValorTotal, _
        tTituloPagar.dValorICMS, tTituloPagar.iCreditoICMS, tTituloPagar.dValorICMSSubst, _
        tTituloPagar.dValorFrete, tTituloPagar.dValorINSS, tTituloPagar.dValorIPI, tTituloPagar.iCreditoIPI, tTituloPagar.dOutrasDespesas, _
        tTituloPagar.dValorProdutos, tTituloPagar.dValorIRRF, tTituloPagar.dValorSeguro, tTituloPagar.dValorINSS, tTituloPagar.iINSSRetido, _
        tTituloPagar.iFilialPedCompra, tTituloPagar.lNumPedCompra, tTituloPagar.dPISRetido, tTituloPagar.dISSRetido, tTituloPagar.dCOFINSRetido, tTituloPagar.dCSLLRetido, tTituloPagar.dTaxaMoeda, lSituacao, lHistorico, lSubConta, objTituloPagar.lFornecedor, objTituloPagar.iFilial, objTituloPagar.lNumTitulo, objTituloPagar.dtDataEmissao, objTituloPagar.sSiglaDocumento, STATUS_EXCLUIDO)
    If lErro <> AD_SQL_SUCESSO Then gError 111785
    '###############################################

    'Tenta ler
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 111786
    
    'Se for uma alteracao ...
    If lErro <> AD_SQL_SEM_DADOS Then

        If Not bPodeAlterar Then gError 180335
        
        'Faz lock no TítuloPagar
        lErro = Comando_LockExclusive(lComando)
        If lErro <> AD_SQL_SUCESSO Then gError 25028
        
        'Se não pertence a essa FilialEmpresa, erro
        If tTituloPagar.iFilialEmpresa <> objTituloPagar.iFilialEmpresa Then Error 18475
        
        '###########################################
        'Inserido por Wagner 18/02/2005
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CAMPOS_ALTERAVEIS_NFFATPAG") 'Mensagem alterada por Wagner
        If vbMsgRes = vbNo Then gError 132224
        '###########################################
        
        '###########################################
        'Alterado por Wagner 18/02/2005
        'verifica a alteracao dos campos alteraveis , se houve  - - - > altera no BD
        If Abs(tTituloPagar.dTaxaMoeda - objTituloPagar.dTaxaMoeda) > DELTA_VALORMONETARIO2 Or tTituloPagar.iFilialPedCompra <> objTituloPagar.iFilialPedCompra Or tTituloPagar.lNumPedCompra <> objTituloPagar.lNumPedCompra _
            Or lHistorico <> objNFPagTitPagHic.lHistorico Or lSituacao <> objNFPagTitPagHic.lSituacao Or lSubConta <> objNFPagTitPagHic.lSubConta Or tTituloPagar.sObservacao <> objNFPagTitPagHic.sObservacao Then
            
            lErro = Comando_ExecutarPos(lComando2, "UPDATE TitulosPag SET TaxaMoeda = ?, FilialPedCompra = ?, NumPedCompra =?, Situacao = ?, Historico = ?, SubConta = ?, Observacao = ? ", lComando, objTituloPagar.dTaxaMoeda, objTituloPagar.iFilialPedCompra, objTituloPagar.lNumPedCompra, objNFPagTitPagHic.lSituacao, objNFPagTitPagHic.lHistorico, objNFPagTitPagHic.lSubConta, objNFPagTitPagHic.sObservacao)
            If lErro <> AD_SQL_SUCESSO Then gError 48918
        
        End If
        '###########################################
        
        'Coloca o NumIntTitulo lido na coleção de parcelas
        For iIndice = 1 To colParcelaPagar.Count
            colParcelaPagar(iIndice).lNumIntTitulo = tTituloPagar.lNumIntDoc
        Next
        
        'Coloca o NumIntDoc lido no objTituloPagar
        objTituloPagar.lNumIntDoc = tTituloPagar.lNumIntDoc
        
        'Lê no BD as Parcelas referentes ao Título
        lErro = CF("ParcelasPagar_Le", objTituloPagar, colParcelaPagarBD)
        If lErro <> SUCESSO Then gError 18502

        'Verifica se o número de Parcelas da tela é igual ao do BD
        If colParcelaPagarBD.Count <> colParcelaPagar.Count Then gError 18503
                  
        'Flag para alterações
        iCamposAlterados = 0
        
        'Verifica se os campos da Tela não alteráveis foram alterados
        If tTituloPagar.dValorTotal <> objTituloPagar.dValorTotal Or tTituloPagar.dValorICMS <> objTituloPagar.dValorICMS Or tTituloPagar.iCreditoICMS <> objTituloPagar.iCreditoICMS Or tTituloPagar.dValorICMSSubst <> objTituloPagar.dValorICMSSubst Or _
           tTituloPagar.dValorFrete <> objTituloPagar.dValorFrete Or tTituloPagar.dValorINSS <> objTituloPagar.dValorINSS Or tTituloPagar.dValorIPI <> objTituloPagar.dValorIPI Or tTituloPagar.iCreditoIPI <> objTituloPagar.iCreditoIPI Or _
           tTituloPagar.dOutrasDespesas <> objTituloPagar.dOutrasDespesas Or tTituloPagar.dValorProdutos <> objTituloPagar.dValorProdutos Or tTituloPagar.dValorIRRF <> objTituloPagar.dValorIRRF Or tTituloPagar.dValorSeguro <> objTituloPagar.dValorSeguro Or _
           tTituloPagar.dValorINSS <> objTituloPagar.dValorINSS Or tTituloPagar.iINSSRetido <> objTituloPagar.iINSSRetido Or _
           tTituloPagar.dPISRetido <> objTituloPagar.dPISRetido Or tTituloPagar.dISSRetido <> objTituloPagar.dISSRetido Or tTituloPagar.dCOFINSRetido <> objTituloPagar.dCOFINSRetido Or tTituloPagar.dCSLLRetido <> objTituloPagar.dCSLLRetido _
           Then iCamposAlterados = 1
        
        'Verifica se campos do Grid Parcelas não alteráveis foram alterados
        If iCamposAlterados = 0 Then
            
            For iIndice = 1 To colParcelaPagar.Count
                If colParcelaPagar(iIndice).dValor <> colParcelaPagarBD(iIndice).dValor Then iCamposAlterados = 1
            Next
            
        End If
        
        'Se houve alteração de campos não alteráveis
        If iCamposAlterados = 1 Then
            
            'Pergunta se deseja prosseguir na alteração para os alteráveis
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_TITULOS_PAGAR_CAMPOS_ALTERAVEIS")
            If vbMsgRes = vbNo Then gError 18504
                        
        End If
                    
        'Altera os campos alteráveis
        lErro = CF("ParcelasPagar_Altera_CamposAlteraveis", colParcelaPagar)
        If lErro <> SUCESSO Then gError 18505

    Else  'Nota Fiscal Fatura Pagar não existe na tabela TitulosPag

        If Not bPodeAlterar Then
        
            'se existir o titulo como baixado entao dar erro
    
            'Verifica se a NotaFiscalFatura já existe em TitulosPagBaixados
            lErro = Comando_ExecutarPos(lComando, "SELECT NumIntDoc FROM TitulosPagBaixados WHERE Fornecedor = ? AND Filial = ? AND NumTitulo = ? AND DataEmissao = ? AND SiglaDocumento = ? AND Status <> ?", 0, _
                tTituloPagar.lNumIntDoc, objTituloPagar.lFornecedor, objTituloPagar.iFilial, objTituloPagar.lNumTitulo, objTituloPagar.dtDataEmissao, objTituloPagar.sSiglaDocumento, STATUS_EXCLUIDO)
            If lErro <> AD_SQL_SUCESSO Then gError 130781
        
            'Tenta ler
            lErro = Comando_BuscarPrimeiro(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130782
            
            'Se for uma alteracao ...
            If lErro <> AD_SQL_SEM_DADOS Then gError 180335
            
        End If
                
        'Testa se há Titulo com mesmo número no Sistema
        lErro = CF("NFFatPag_Testa_Repeticao", objTituloPagar)
        If lErro <> SUCESSO Then gError 18529
            
        objFilialFornecedor.lCodFornecedor = objTituloPagar.lFornecedor
        objFilialFornecedor.iCodFilial = objTituloPagar.iFilial
        
        'Lê a Filial Fornecedor
        lErro = CF("FilialFornecedor_Lock1", objFilialFornecedor)
        If lErro <> SUCESSO And lErro <> 18900 Then gError 18464
    
        'Se não encontrar Filial Fornecedor --> Erro
        If lErro = 18900 Then gError 18465

        If objTituloPagar.iCondicaoPagto <> 0 Then
        
            lErro = CF("CondicaoPagto_Lock", lComando3, objTituloPagar.iCondicaoPagto)
            If lErro <> SUCESSO Then gError 59067
            
        End If
        
        'Gera número interno automático
        lErro = CF("TituloPagar_Automatico", tTituloPagar.lNumIntDoc)
        If lErro <> SUCESSO Then gError 18479
        
        objTituloPagar.lNumIntDoc = tTituloPagar.lNumIntDoc
        objTituloPagar.dtDataRegistro = gdtDataHoje

        'Preenche Saldo = ValorTotal - retencoes
        objTituloPagar.dSaldo = objTituloPagar.dValorTotal - (objTituloPagar.dValorIRRF + objTituloPagar.dPISRetido + objTituloPagar.dISSRetido + objTituloPagar.dCOFINSRetido + objTituloPagar.dCSLLRetido) - IIf(objTituloPagar.iINSSRetido = 1, objTituloPagar.dValorINSS, 0)
        
        If objTituloPagar.iFilialEmpresa = 0 Then objTituloPagar.iFilialEmpresa = giFilialEmpresa
        
        '#############################################
        'Alterado por Wagner 18/02/2005
        'Insere o Título na tabela TitulosPag
        lErro = Comando_Executar(lComando1, "INSERT INTO TitulosPag (NumIntDoc,FilialEmpresa,Fornecedor,Filial,SiglaDocumento,NumTitulo,DataEmissao,Status,Saldo,NumParcelas,ValorTotal,ValorSeguro,ValorFrete,OutrasDespesas,ValorProdutos,ValorICMS,ValorICMSSubst,CreditoICMS,ValorIPI,CreditoIPI,ValorIRRF,ValorINSS, INSSRetido, DataRegistro, Observacao, FilialPedCompra, NumPedCompra, CondicaoPagto, PISRetido, ISSRetido, COFINSRetido, CSLLRetido, TaxaMoeda, Situacao, Historico, SubConta) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", tTituloPagar.lNumIntDoc, objTituloPagar.iFilialEmpresa, objTituloPagar.lFornecedor, objTituloPagar.iFilial, objTituloPagar.sSiglaDocumento, objTituloPagar.lNumTitulo, objTituloPagar.dtDataEmissao, STATUS_LANCADO, objTituloPagar.dSaldo, objTituloPagar.iNumParcelas, _
            objTituloPagar.dValorTotal, objTituloPagar.dValorSeguro, objTituloPagar.dValorFrete, objTituloPagar.dOutrasDespesas, objTituloPagar.dValorProdutos, objTituloPagar.dValorICMS, objTituloPagar.dValorICMSSubst, objTituloPagar.iCreditoICMS, objTituloPagar.dValorIPI, objTituloPagar.iCreditoIPI, objTituloPagar.dValorIRRF, objTituloPagar.dValorINSS, objTituloPagar.iINSSRetido, objTituloPagar.dtDataRegistro, objNFPagTitPagHic.sObservacao, objTituloPagar.iFilialPedCompra, objTituloPagar.lNumPedCompra, objTituloPagar.iCondicaoPagto, objTituloPagar.dPISRetido, objTituloPagar.dISSRetido, objTituloPagar.dCOFINSRetido, objTituloPagar.dCSLLRetido, objTituloPagar.dTaxaMoeda, objNFPagTitPagHic.lSituacao, objNFPagTitPagHic.lHistorico, objNFPagTitPagHic.lSubConta)
        If lErro <> AD_SQL_SUCESSO Then gError 18480
        '#############################################

        'Armazena em colParcelaPagar o lNumIntTitulo para cada parcela
        For iIndice = 1 To colParcelaPagar.Count
            colParcelaPagar(iIndice).lNumIntTitulo = tTituloPagar.lNumIntDoc
        Next
        
        'Grava ParcelasPagar no BD
        lErro = CF("ParcelasPagar_Grava", colParcelaPagar)
        If lErro <> SUCESSO Then gError 18481
    
    End If
        
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    
    NFFatPag_Grava_BD = SUCESSO
    
    Exit Function

Erro_NFFatPag_Grava_BD:

    NFFatPag_Grava_BD = gErr
    
    Select Case gErr
    
        Case 18465
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_CADASTRADA", gErr, objTituloPagar.lFornecedor, objTituloPagar.iFilial)
    
        Case 18471, 18472, 48919, 59066
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 18473, 18474, 130781, 130782
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFFATURA", gErr, objTituloPagar.lNumTitulo)
            
        Case 18475
            Call Rotina_Erro(vbOKOnly, "ERRO_NFFAT_FILIALEMPRESA_DIFERENTE", gErr, objTituloPagar.lNumTitulo)
        
        Case 18464, 18479, 18481, 18502, 18504, 18505, 18529, 59067, 132224, 180335
        
        Case 18480
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_NFFATURA", gErr, objTituloPagar.lNumTitulo)
                    
        Case 18503
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_PARCELAS_ALTERADO", gErr, colParcelaPagarBD.Count, colParcelaPagar.Count)
        
        Case 25028
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_TITULOSPAG", gErr, objTituloPagar.lNumIntDoc)
            
        Case 48918
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_TITULOSPAG", gErr, tTituloPagar.lNumIntDoc)
            
        Case 108791
            Call Rotina_Erro(vbOKOnly, "ERRO_TITULO_PAGAR_INEXISTENTE", gErr)
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179153)

    End Select
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    
End Function

Function NotasFiscaisPag_Baixar(ByVal lNumIntTitulo As Long, ByVal lComandoSel As Long, ByVal lComandoDel As Long, ByVal lComandoIns As Long) As Long
'transfere as notas fiscais vinculadas a um titulo a pagar de NFsPag para NFsPagBaixadas
'TEM QUE SER CHAMADA DENTRO DE TRANSACAO
Dim lErro As Long
Dim tNFsPag As typeNFsPag
Dim iIndice As Integer
Dim lSituacao As Long
Dim lHistorico As Long
Dim lSubConta As Long
Dim sObservacao As String

On Error GoTo Erro_NotasFiscaisPag_Baixar

    '#######################################
    'Inserido por Wagner
    sObservacao = String(STRING_NFPAG_OBSERVACAO, 0)
    '#######################################

    '#######################################
    'Alterado por Wagner
    lErro = Comando_ExecutarPos(lComandoSel, "SELECT NumIntDoc, FilialEmpresa, Fornecedor, Filial, NumNotaFiscal, DataEmissao, Status, DataVencimento, ValorTotal, ValorSeguro, ValorFrete, OutrasDespesas, ValorProdutos, ValorIRRF, ValorICMS, ValorICMSSubst, CreditoICMS, ValorIPI, CreditoIPI, PISRetido, COFINSRetido, CSLLRetido, TaxaMoeda, Situacao, Historico, SubConta, Observacao FROM NFsPag WHERE NumIntTitPag = ?", 0, tNFsPag.lNumIntDoc, tNFsPag.iFilialEmpresa, tNFsPag.lFornecedor, tNFsPag.iFilial, tNFsPag.lNumNotaFiscal, tNFsPag.dtDataEmissao, tNFsPag.iStatus, tNFsPag.dtDataVencimento, tNFsPag.dValorTotal, tNFsPag.dValorSeguro, tNFsPag.dValorFrete, tNFsPag.dOutrasDespesas, tNFsPag.dValorProdutos, tNFsPag.dValorIRRF, tNFsPag.dValorICMS, tNFsPag.dValorICMSSubst, tNFsPag.iCreditoICMS, tNFsPag.dValorIPI, tNFsPag.iCreditoIPI, tNFsPag.dPISRetido, tNFsPag.dCOFINSRetido, tNFsPag.dCSLLRetido, tNFsPag.dTaxaMoeda, lSituacao, lHistorico, lSubConta, sObservacao, lNumIntTitulo)
    If lErro <> AD_SQL_SUCESSO Then Error 7321
    '#######################################

    lErro = Comando_BuscarProximo(lComandoSel)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 7322

    If lErro = AD_SQL_SEM_DADOS Then Error 19377
        
    Do While lErro = AD_SQL_SUCESSO

        'Excluir a NF da tabela NFsPag
        lErro = Comando_ExecutarPos(lComandoDel, "DELETE FROM NFsPag", lComandoSel)
        If lErro <> AD_SQL_SUCESSO Then Error 7323

        '#######################################
        'Alterado por Wagner - Inseridos campos Situacao, Historico, SubConta e Observacao
        lErro = Comando_Executar(lComandoIns, "INSERT INTO NFsPagBaixadas(NumIntDoc,FilialEmpresa,Fornecedor,Filial,NumNotaFiscal,DataEmissao,Status, NumIntTitPag, DataVencimento,ValorTotal,ValorSeguro, ValorFrete, OutrasDespesas, ValorProdutos, ValorIRRF, ValorICMS, ValorICMSSUbst, CreditoICMS, ValorIPI, CreditoIPI, PISRetido, COFINSRetido, CSLLRetido, TaxaMoeda, Situacao, Historico, SubConta, Observacao) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
            tNFsPag.lNumIntDoc, tNFsPag.iFilialEmpresa, tNFsPag.lFornecedor, tNFsPag.iFilial, tNFsPag.lNumNotaFiscal, tNFsPag.dtDataEmissao, tNFsPag.iStatus, lNumIntTitulo, tNFsPag.dtDataVencimento, tNFsPag.dValorTotal, tNFsPag.dValorSeguro, tNFsPag.dValorFrete, tNFsPag.dOutrasDespesas, tNFsPag.dValorProdutos, tNFsPag.dValorIRRF, tNFsPag.dValorICMS, tNFsPag.dValorICMSSubst, tNFsPag.iCreditoICMS, tNFsPag.dValorIPI, tNFsPag.iCreditoIPI, tNFsPag.dPISRetido, tNFsPag.dCOFINSRetido, tNFsPag.dCSLLRetido, tNFsPag.dTaxaMoeda, lSituacao, lHistorico, lSubConta, sObservacao)
        If lErro <> AD_SQL_SUCESSO Then Error 7325
        '#######################################
        
        '#######################################
        'Alterado por Wagner - Chama via CF
        'Altera o Status da Nota para Baixado
        lErro = CF("NFiscal_Altera_Status", tNFsPag.lNumIntDoc, CPR_NF_PAGAR, STATUS_BAIXADO)
        If lErro <> SUCESSO Then Error 61291
        '#######################################
        
        lErro = Comando_BuscarProximo(lComandoSel)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 7324

    Loop

    NotasFiscaisPag_Baixar = SUCESSO

    Exit Function

Erro_NotasFiscaisPag_Baixar:

    NotasFiscaisPag_Baixar = Err

    Select Case Err

        Case 7321, 7322, 7324, 19377
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NOTAS_FISCAIS_EXT", Err)

        Case 7323
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_NOTAS_FISCAIS_EXT", Err)

        Case 7325
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_NOTAS_FISCAIS_EXT_BAIXADAS", Err)
        
        Case 61291 'Tratado na rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 179154)

    End Select

    Exit Function

End Function

Function TitulosPagBaixados_Inserir(ByVal objTituloPagar As ClassTituloPagar, ByVal lComandoTitIns As Long) As Long
'auxiliar a baixa de parcelas ( e portanto de titulos )

Dim lErro As Long
Dim objNFPagTitPagHic As New ClassNfPagTitPagHic 'Inserido por Wagner

On Error GoTo Erro_TitulosPagBaixados_Inserir

    '###############################################
    'Inserido por Wagner
    If Not (objTituloPagar.objInfoUsu Is Nothing) Then
        Set objNFPagTitPagHic = objTituloPagar.objInfoUsu
    End If
    '###############################################

    '#######################################
    'Alterado por Wagner
    lErro = Comando_Executar(lComandoTitIns, "INSERT INTO TitulosPagBaixados " & _
            "( NumIntDoc, FilialEmpresa, Fornecedor, Filial, SiglaDocumento, NumTitulo, DataEmissao, Status, Saldo, NumParcelas, ValorTotal, ValorSeguro, ValorFrete, OutrasDespesas, ValorProdutos, ValorICMS, ValorICMSSubst, CreditoICMS, ValorIPI, CreditoIPI, ValorIRRF, ValorINSS, INSSRetido, DataRegistro, Observacao, CondicaoPagto, FilialPedCompra, NumPedCompra, PISRetido, ISSRetido, COFINSRetido, CSLLRetido, TaxaMoeda, Situacao, Historico, SubConta) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
            objTituloPagar.lNumIntDoc, objTituloPagar.iFilialEmpresa, objTituloPagar.lFornecedor, objTituloPagar.iFilial, objTituloPagar.sSiglaDocumento, objTituloPagar.lNumTitulo, objTituloPagar.dtDataEmissao, objTituloPagar.iStatus, objTituloPagar.dSaldo, objTituloPagar.iNumParcelas, objTituloPagar.dValorTotal, objTituloPagar.dValorSeguro, objTituloPagar.dValorFrete, objTituloPagar.dOutrasDespesas, objTituloPagar.dValorProdutos, objTituloPagar.dValorICMS, objTituloPagar.dValorICMSSubst, objTituloPagar.iCreditoICMS, objTituloPagar.dValorIPI, objTituloPagar.iCreditoIPI, objTituloPagar.dValorIRRF, objTituloPagar.dValorINSS, objTituloPagar.iINSSRetido, objTituloPagar.dtDataRegistro, objTituloPagar.sObservacao, objTituloPagar.iCondicaoPagto, objTituloPagar.iFilialPedCompra, objTituloPagar.lNumPedCompra, objTituloPagar.dPISRetido, objTituloPagar.dISSRetido, objTituloPagar.dCOFINSRetido, objTituloPagar.dCSLLRetido, objTituloPagar.dTaxaMoeda, _
            objNFPagTitPagHic.lSituacao, objNFPagTitPagHic.lHistorico, objNFPagTitPagHic.lSubConta)
    If lErro <> AD_SQL_SUCESSO Then Error 7313
    '#######################################

    'Altera o status da Nfiscal para Baixado caso tenha nota vinculada
    lErro = CF("NFiscal_Altera_Status", objTituloPagar.lNumIntDoc, CPR_TITULO_PAGAR, STATUS_BAIXADO)
    If lErro <> SUCESSO Then Error 61289
    
    TitulosPagBaixados_Inserir = SUCESSO

    Exit Function

Erro_TitulosPagBaixados_Inserir:

    TitulosPagBaixados_Inserir = Err

    Select Case Err

        Case 7313
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TITULOS_PAG", Err)
        
        Case 61289 'Tratado na rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 179155)

    End Select

    Exit Function

End Function

Function TitulosPagar_LeMod(ByVal objTituloPagar As ClassTituloPagar, ByVal lComando As Long) As Long
'faz a leitura de um registro da tabela TitulosPagar p/objTituloPagar c/ExecutarPos
'objTituloPagar.lNumIntDoc entra preenchido
'TEM QUE SER CHAMADO EM TRANSACAO

Dim lErro As Long
Dim tTituloPagar As typeTituloPagar
Dim objNFPagTitPag As New ClassNfPagTitPagHic 'Inserido por Wagner
Dim lSituacao As Long
Dim lHistorico As Long
Dim lSubConta As Long

On Error GoTo Erro_TitulosPagar_LeMod

    tTituloPagar.sObservacao = String(STRING_TITULO_OBSERVACAO, 0)
    tTituloPagar.sSiglaDocumento = String(STRING_SIGLA_DOCUMENTO, 0)

    '#######################################
    'Alterado por Wagner
    lErro = Comando_ExecutarPos(lComando, "SELECT OutrasDespesas , Saldo, DataEmissao, ValorFrete, ValorICMS, ValorICMSSubst, ValorINSS,INSSRetido, DataRegistro, " & _
        "ValorIPI , ValorIRRF, ValorProdutos, ValorSeguro, ValorTotal, CreditoICMS, CreditoIPI, " & _
        "Filial, FilialEmpresa, NumParcelas, Status, Fornecedor, " & _
        "NumTitulo, Observacao, SiglaDocumento, CondicaoPagto, FilialPedCompra, NumPedCompra, PISRetido, ISSRetido, COFINSRetido, CSLLRetido, TaxaMoeda, Situacao, Historico, SubConta FROM TitulosPag WHERE TitulosPag.NumIntDoc = ?", 0, _
        tTituloPagar.dOutrasDespesas, tTituloPagar.dSaldo, tTituloPagar.dtDataEmissao, tTituloPagar.dValorFrete, tTituloPagar.dValorICMS, tTituloPagar.dValorICMSSubst, tTituloPagar.dValorINSS, tTituloPagar.iINSSRetido, tTituloPagar.dtDataRegistro, _
        tTituloPagar.dValorIPI, tTituloPagar.dValorIRRF, tTituloPagar.dValorProdutos, tTituloPagar.dValorSeguro, tTituloPagar.dValorTotal, tTituloPagar.iCreditoICMS, tTituloPagar.iCreditoIPI, _
        tTituloPagar.iFilial, tTituloPagar.iFilialEmpresa, tTituloPagar.iNumParcelas, tTituloPagar.iStatus, tTituloPagar.lFornecedor, _
        tTituloPagar.lNumTitulo, tTituloPagar.sObservacao, tTituloPagar.sSiglaDocumento, tTituloPagar.iCondicaoPagto, tTituloPagar.iFilialPedCompra, tTituloPagar.lNumPedCompra, tTituloPagar.dPISRetido, tTituloPagar.dISSRetido, tTituloPagar.dCOFINSRetido, tTituloPagar.dCSLLRetido, tTituloPagar.dTaxaMoeda, lSituacao, lHistorico, lSubConta, _
        objTituloPagar.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then Error 7311
    '#######################################

    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 7407
    
    If lErro <> AD_SQL_SUCESSO Then Error 7408
    
    tTituloPagar.lNumIntDoc = objTituloPagar.lNumIntDoc
    Call TituloPag_MoverCamposObj(objTituloPagar, tTituloPagar)

    '###############################################
    'Inserido por Wagner 18/02/2005
    objNFPagTitPag.lSituacao = lSituacao
    objNFPagTitPag.lHistorico = lHistorico
    objNFPagTitPag.lSubConta = lSubConta
    objNFPagTitPag.sObservacao = tTituloPagar.sObservacao
    
    Set objTituloPagar.objInfoUsu = objNFPagTitPag
    '###############################################

    TitulosPagar_LeMod = SUCESSO

    Exit Function

Erro_TitulosPagar_LeMod:

    TitulosPagar_LeMod = Err

    Select Case Err

        Case 7311, 7407
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOS_PAGAR", Err)

        Case 7408
            Call Rotina_Erro(vbOKOnly, "ERRO_TITULO_PAGAR_INEXISTENTE", Err)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 179156)

    End Select

    Exit Function

End Function

Private Sub TituloPag_MoverCamposObj(ByVal objTituloPagar As ClassTituloPagar, tTituloPagar As typeTituloPagar)

    objTituloPagar.dOutrasDespesas = tTituloPagar.dOutrasDespesas
    objTituloPagar.dSaldo = tTituloPagar.dSaldo
    objTituloPagar.dtDataEmissao = tTituloPagar.dtDataEmissao
    objTituloPagar.dtDataRegistro = tTituloPagar.dtDataRegistro
    objTituloPagar.dValorFrete = tTituloPagar.dValorFrete
    objTituloPagar.dValorICMS = tTituloPagar.dValorICMS
    objTituloPagar.dValorICMSSubst = tTituloPagar.dValorICMSSubst
    objTituloPagar.dValorINSS = tTituloPagar.dValorINSS
    objTituloPagar.dValorIPI = tTituloPagar.dValorIPI
    objTituloPagar.dValorIRRF = tTituloPagar.dValorIRRF
    objTituloPagar.dValorProdutos = tTituloPagar.dValorProdutos
    objTituloPagar.dValorSeguro = tTituloPagar.dValorSeguro
    objTituloPagar.dValorTotal = tTituloPagar.dValorTotal
    objTituloPagar.iCreditoICMS = tTituloPagar.iCreditoICMS
    objTituloPagar.iCreditoIPI = tTituloPagar.iCreditoIPI
    objTituloPagar.iFilial = tTituloPagar.iFilial
    objTituloPagar.iFilialEmpresa = tTituloPagar.iFilialEmpresa
    objTituloPagar.iINSSRetido = tTituloPagar.iINSSRetido
    objTituloPagar.iNumParcelas = tTituloPagar.iNumParcelas
    objTituloPagar.iStatus = tTituloPagar.iStatus
    objTituloPagar.lFornecedor = tTituloPagar.lFornecedor
    objTituloPagar.lNumTitulo = tTituloPagar.lNumTitulo
    objTituloPagar.sObservacao = tTituloPagar.sObservacao
    objTituloPagar.sSiglaDocumento = tTituloPagar.sSiglaDocumento
    objTituloPagar.iCondicaoPagto = tTituloPagar.iCondicaoPagto
    objTituloPagar.iFilialPedCompra = tTituloPagar.iFilialPedCompra
    objTituloPagar.lNumPedCompra = tTituloPagar.lNumPedCompra
    objTituloPagar.dPISRetido = tTituloPagar.dPISRetido
    objTituloPagar.dISSRetido = tTituloPagar.dISSRetido
    objTituloPagar.dCOFINSRetido = tTituloPagar.dCOFINSRetido
    objTituloPagar.dCSLLRetido = tTituloPagar.dCSLLRetido
    objTituloPagar.dTaxaMoeda = tTituloPagar.dTaxaMoeda
    
End Sub

Function NFFatPag_Grava(ByVal objTituloPagar As ClassTituloPagar, ByVal colParcelas As colParcelaPagar, ByVal objContabil As ClassContabil) As Long
'Grava Nota Fiscal Fatura a Pagar no BD (nível + alto)

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim lTransacao As Long, lNumIntNF As Long
Dim iFilialEmpresa As Integer
Dim objParcelaPagar As ClassParcelaPagar, dValorLivroAux As Double
Dim lNumIntDoc As Long
Dim vbMsgRes As VbMsgBoxResult 'Inserido por Wagner
Dim objNFPagTitPagHic As New ClassNfPagTitPagHic 'Inserido por Wagner

On Error GoTo Erro_NFFatPag_Grava

    'Abre a transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 18461
    
    'Abre os comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 18462
    
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 18463
    
    '###############################################
    'Inserido por Wagner
    If Not (objTituloPagar.objInfoUsu Is Nothing) Then
        Set objNFPagTitPagHic = objTituloPagar.objInfoUsu
    End If
    '###############################################
    
    '##########################################################
    'Alterado por Wagner
    'Verifica se o Título está cadastrado como um Título Baixado
    lErro = Comando_ExecutarPos(lComando, "SELECT NumIntDoc FROM TitulosPagBaixados WHERE Fornecedor = ? AND Filial = ? AND NumTitulo = ? AND DataEmissao = ? AND SiglaDocumento = ? AND Status <> ?", 0, lNumIntDoc, objTituloPagar.lFornecedor, objTituloPagar.iFilial, objTituloPagar.lNumTitulo, objTituloPagar.dtDataEmissao, objTituloPagar.sSiglaDocumento, STATUS_EXCLUIDO)
    If lErro <> AD_SQL_SUCESSO Then gError 18476
    '##########################################################
    
    'Tenta ler
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 18477
    
    'Se nao encontrou
    If lErro <> AD_SQL_SUCESSO Then
    
        'Verifica se o Título a pagar possui um Adiantamento vinculado ao Fornecedor
        lErro = CF("TituloPagar_Verifica_Adiantamento", objTituloPagar.lFornecedor, objTituloPagar.iFilial)
        If lErro <> SUCESSO Then gError 75478
            
        'Grava no BD
        lErro = CF("NFFatPag_Grava_BD", objTituloPagar, colParcelas)
        If lErro <> SUCESSO Then gError 18470
        
    Else
        
        '##############################################
        'Inserido por Wagner 21/02/2005
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CAMPOS_ALTERAVEIS_NFFATPAGBAIXADO") 'Mensagem alterada por Wagner
        If vbMsgRes = vbNo Then gError 132242
            
        lErro = Comando_ExecutarPos(lComando1, "UPDATE TitulosPagBaixados SET Situacao = ?, Historico = ?, SubConta = ?, Observacao = ? ", lComando, objNFPagTitPagHic.lSituacao, objNFPagTitPagHic.lHistorico, objNFPagTitPagHic.lSubConta, objNFPagTitPagHic.sObservacao)
        If lErro <> AD_SQL_SUCESSO Then gError 132243
        '##############################################

    End If
        
    If (gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO) Then
    
        'se o documento foi originado fora de CPR nao pode ter a contabilidade alterada em CPR
        lErro = CF("DocumentoCPR_OrigemNFiscal", objTituloPagar.sSiglaDocumento, objTituloPagar.lNumIntDoc, lNumIntNF)
        If lErro <> SUCESSO And lErro <> 41542 Then gError 59224
        
        'se nao é um docto gerado fora do cpr
        If lErro <> SUCESSO Then
            
            dValorLivroAux = 0
                        
            For Each objParcelaPagar In colParcelas
            
                dValorLivroAux = dValorLivroAux + Round(objParcelaPagar.dValor, 2)
                
            Next
            
            'Grava os dados contábeis (contabilidade)
            lErro = objContabil.Contabil_Gravar_Registro(objTituloPagar.lNumIntDoc, objTituloPagar.lFornecedor, objTituloPagar.iFilial, DATA_NULA, , 0, , dValorLivroAux)
            If lErro <> SUCESSO Then gError 20473
    
        End If
        
    End If
    
    'Encerra transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 18471
    
    'Fecha os comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    NFFatPag_Grava = SUCESSO
    
    Exit Function
    
Erro_NFFatPag_Grava:

    NFFatPag_Grava = gErr
    
    Select Case gErr
    
        Case 18461
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 18462, 18463
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 18470, 20473, 59224, 75478
        
        Case 18471
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case 18476, 18477
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFFATBAIXADA", gErr, objTituloPagar.lNumTitulo)
        
        Case 18478
        
        '################################
        'Inserido por Wagner
        Case 132242
        
        Case 132243
            Call Rotina_Erro(vbOKOnly, "ERRO_TITULOSPAGBAIXADO_NAO_EXISTENTE", gErr, objTituloPagar.lNumTitulo)
        '################################
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179157)

    End Select
    
    Call Transacao_Rollback
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    Exit Function

End Function

Public Function TitulosPag_Inserir(ByVal objTituloPagar As ClassTituloPagar, ByVal lComandoTitIns As Long) As Long
'auxiliar ao cancelamento de uma baixa

Dim lErro As Long
Dim objNFPagTitPagHic As New ClassNfPagTitPagHic 'Inserido por Wagner

On Error GoTo Erro_TitulosPag_Inserir
    
    '###############################################
    'Inserido por Wagner
    If Not (objTituloPagar.objInfoUsu Is Nothing) Then
        Set objNFPagTitPagHic = objTituloPagar.objInfoUsu
    End If
    '###############################################
    
    '###############################################
    'Alterado por Wagner
    lErro = Comando_Executar(lComandoTitIns, "INSERT INTO TitulosPag " & _
            "( NumIntDoc, FilialEmpresa, Fornecedor, Filial, SiglaDocumento, NumTitulo, DataEmissao, Status, Saldo, NumParcelas, ValorTotal, ValorSeguro, ValorFrete, OutrasDespesas, ValorProdutos, ValorICMS, ValorICMSSubst, CreditoICMS, ValorIPI, CreditoIPI, ValorIRRF, ValorINSS, INSSRetido, DataRegistro, Observacao, CondicaoPagto, FilialPedCompra, NumPedCompra, PISRetido, ISSRetido, COFINSRetido, CSLLRetido, TaxaMoeda, Situacao, Historico, SubConta) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
            objTituloPagar.lNumIntDoc, objTituloPagar.iFilialEmpresa, objTituloPagar.lFornecedor, objTituloPagar.iFilial, objTituloPagar.sSiglaDocumento, objTituloPagar.lNumTitulo, objTituloPagar.dtDataEmissao, objTituloPagar.iStatus, objTituloPagar.dSaldo, objTituloPagar.iNumParcelas, objTituloPagar.dValorTotal, objTituloPagar.dValorSeguro, objTituloPagar.dValorFrete, objTituloPagar.dOutrasDespesas, objTituloPagar.dValorProdutos, objTituloPagar.dValorICMS, objTituloPagar.dValorICMSSubst, objTituloPagar.iCreditoICMS, objTituloPagar.dValorIPI, objTituloPagar.iCreditoIPI, objTituloPagar.dValorIRRF, objTituloPagar.dValorINSS, objTituloPagar.iINSSRetido, objTituloPagar.dtDataRegistro, objNFPagTitPagHic.sObservacao, objTituloPagar.iCondicaoPagto, objTituloPagar.iFilialPedCompra, objTituloPagar.lNumPedCompra, objTituloPagar.dPISRetido, objTituloPagar.dISSRetido, objTituloPagar.dCOFINSRetido, objTituloPagar.dCSLLRetido, objTituloPagar.dTaxaMoeda, _
            objNFPagTitPagHic.lSituacao, objNFPagTitPagHic.lHistorico, objNFPagTitPagHic.lSubConta)
    If lErro <> AD_SQL_SUCESSO Then Error 7874
    '###############################################

    'Atualiza o Status da NFiscal caso tenha vinculo com Nota Fiscal
    lErro = CF("NFiscal_Altera_Status", objTituloPagar.lNumIntDoc, CPR_TITULO_PAGAR, STATUS_LANCADO)
    If lErro <> SUCESSO Then Error 61287

    TitulosPag_Inserir = SUCESSO

    Exit Function

Erro_TitulosPag_Inserir:

    TitulosPag_Inserir = Err

    Select Case Err

        Case 7874
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TITULOS_PAG", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 179158)

    End Select

    Exit Function

End Function

Function TituloPagarBaixado_LeMod(ByVal objTituloPagar As ClassTituloPagar, ByVal lComando As Long) As Long
'faz a leitura de um registro da tabela TitulosPagarBaixados p/objTituloPagar c/ExecutarPos
'objTituloPagar.lNumIntDoc entra preenchido
'TEM QUE SER CHAMADO EM TRANSACAO
Dim lErro As Long
Dim tTituloPagar As typeTituloPagar
Dim objNFPagTitPag As New ClassNfPagTitPagHic 'Inserido por Wagner
Dim lSituacao As Long
Dim lHistorico As Long
Dim lSubConta As Long

On Error GoTo Erro_TituloPagarBaixado_LeMod

    tTituloPagar.sObservacao = String(STRING_TITULO_OBSERVACAO, 0)
    tTituloPagar.sSiglaDocumento = String(STRING_SIGLA_DOCUMENTO, 0)

    '#######################################
    'Alterado por Wagner
    lErro = Comando_ExecutarPos(lComando, "SELECT OutrasDespesas, Saldo, DataEmissao, ValorFrete, ValorICMS, ValorICMSSubst, ValorINSS, INSSRetido, DataRegistro, " & _
        "ValorIPI , ValorIRRF, ValorProdutos, ValorSeguro, ValorTotal, CreditoICMS, CreditoIPI, " & _
        "Filial, FilialEmpresa, NumParcelas, Status, Fornecedor, " & _
        "NumTitulo, Observacao, SiglaDocumento, CondicaoPagto, FilialPedCompra, NumPedCompra, PISRetido, ISSRetido, COFINSRetido, CSLLRetido, TaxaMoeda, Situacao, Historico, SubConta FROM TitulosPagBaixados WHERE NumIntDoc = ?", 0, _
        tTituloPagar.dOutrasDespesas, tTituloPagar.dSaldo, tTituloPagar.dtDataEmissao, tTituloPagar.dValorFrete, tTituloPagar.dValorICMS, tTituloPagar.dValorICMSSubst, tTituloPagar.dValorINSS, tTituloPagar.iINSSRetido, tTituloPagar.dtDataRegistro, _
        tTituloPagar.dValorIPI, tTituloPagar.dValorIRRF, tTituloPagar.dValorProdutos, tTituloPagar.dValorSeguro, tTituloPagar.dValorTotal, tTituloPagar.iCreditoICMS, tTituloPagar.iCreditoIPI, _
        tTituloPagar.iFilial, tTituloPagar.iFilialEmpresa, tTituloPagar.iNumParcelas, tTituloPagar.iStatus, tTituloPagar.lFornecedor, _
        tTituloPagar.lNumTitulo, tTituloPagar.sObservacao, tTituloPagar.sSiglaDocumento, tTituloPagar.iCondicaoPagto, tTituloPagar.iFilialPedCompra, tTituloPagar.lNumPedCompra, tTituloPagar.dPISRetido, tTituloPagar.dISSRetido, tTituloPagar.dCOFINSRetido, tTituloPagar.dCSLLRetido, tTituloPagar.dTaxaMoeda, lSituacao, lHistorico, lSubConta, objTituloPagar.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then Error 7871
    '#######################################

    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 7872

    If lErro <> AD_SQL_SUCESSO Then Error 7873

    tTituloPagar.lNumIntDoc = objTituloPagar.lNumIntDoc
    Call TituloPag_MoverCamposObj(objTituloPagar, tTituloPagar)

    '###############################################
    'Inserido por Wagner 18/02/2005
    objNFPagTitPag.lSituacao = lSituacao
    objNFPagTitPag.lHistorico = lHistorico
    objNFPagTitPag.lSubConta = lSubConta
    objNFPagTitPag.sObservacao = objNFPagTitPag.sObservacao
    
    Set objTituloPagar.objInfoUsu = objNFPagTitPag
    '###############################################

    TituloPagarBaixado_LeMod = SUCESSO

    Exit Function

Erro_TituloPagarBaixado_LeMod:

    TituloPagarBaixado_LeMod = Err

    Select Case Err

        Case 7871, 7872
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOS_PAGAR_BAIXADO", Err)

        Case 7873
            Call Rotina_Erro(vbOKOnly, "ERRO_TITULO_PAGAR_BAIXADO_INEXISTENTE", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 179159)

    End Select

    Exit Function

End Function
'#########################################################################

'#########################################################################
'Inserido por Wagner 19/04/2005
'Faz a Gravação do Favorecido na Parcela a Pagar e atualiza a Observação do Título
Public Function ParcelaPag_Grava_ChequeAvulso(ByVal objChequesPagAvulso As ClassChequesPagAvulso) As Long

Dim lErro As Long
Dim lTransacao  As Long
Dim alComando(1 To 4) As Long
Dim iIndice As Integer
Dim objInfoParcPag As ClassInfoParcPag
Dim lNumIntTitulo As Long
Dim lNumIntTituloAnt As Long
Dim lNumCheque As Long

On Error GoTo Erro_ParcelaPag_Grava_ChequeAvulso

    'Abre a transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 132575
    
    'abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 132576
    Next
    
    lNumIntTituloAnt = 0
    
    For Each objInfoParcPag In objChequesPagAvulso.colInfoParcPag
    
        If objInfoParcPag.iSeqCheque = objChequesPagAvulso.objInfoChequePag.iSeqCheque Then
        
            'Le a Parcela a Pagar
            lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntTitulo FROM ParcelasPag WHERE NumIntDoc = ? ", 0, lNumIntTitulo, objInfoParcPag.lNumIntParc)
            If lErro <> AD_SQL_SUCESSO Then gError 132577
            
            lErro = Comando_BuscarPrimeiro(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 132578
            
            'Se não encontrou=> Erro
            If lErro = AD_SQL_SEM_DADOS Then gError 132579
            
            'Atualiza o Favorecido
            lErro = Comando_ExecutarPos(alComando(2), "UPDATE ParcelasPag SET Favorecido = ?", alComando(1), objChequesPagAvulso.objInfoChequePag.sFavorecido)
            If lErro <> AD_SQL_SUCESSO Then gError 132580
            
            'Se a parcela é de um título diferente da anterior
            If lNumIntTituloAnt <> lNumIntTitulo Then
            
                'Le o Titulo a Pagar
                lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntDoc FROM TitulosPag WHERE NumIntDoc = ? ", 0, lNumIntTituloAnt, lNumIntTitulo)
                If lErro <> AD_SQL_SUCESSO Then gError 132590
            
                'Traz o primeiro registro
                lErro = Comando_BuscarPrimeiro(alComando(1))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 132591
                
                'Se não encontrou=> Erro
                If lErro = AD_SQL_SEM_DADOS Then gError 132592
            
                'Atualiza a Observação
                lErro = Comando_ExecutarPos(alComando(2), "UPDATE TitulosPag SET Observacao = ?, ChequeRetido = ?", alComando(1), objChequesPagAvulso.objInfoChequePag.sObservacao, MARCADO)
                If lErro <> AD_SQL_SUCESSO Then gError 132593
                
            End If
            
            '#######################################################
            'Inserido por Wagner
            'Insere a Conta Corrente, Cheque e NumIntDocParcela
            lErro = Comando_ExecutarPos(alComando(3), "SELECT Cheque FROM ChequeHic WHERE NumIntParcela = ? ", 0, lNumCheque, objInfoParcPag.lNumIntParc)
            If lErro <> AD_SQL_SUCESSO Then gError 132873

            lErro = Comando_BuscarPrimeiro(alComando(3))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 132874

            If lErro = AD_SQL_SEM_DADOS Then

                lErro = Comando_Executar(alComando(4), "INSERT INTO ChequeHic (ContaCorrente, NumIntParcela, Cheque) VALUES (?,?,?)", objChequesPagAvulso.iCta, objInfoParcPag.lNumIntParc, objChequesPagAvulso.lNumCheque)
                If lErro <> AD_SQL_SUCESSO Then gError 132875

            Else

                lErro = Comando_ExecutarPos(alComando(4), "UPDATE ChequeHic  SET ContaCorrente = ?, Cheque = ?", alComando(3), objChequesPagAvulso.iCta, objChequesPagAvulso.lNumCheque)
                If lErro <> AD_SQL_SUCESSO Then gError 132876

            End If
            '#######################################################
        
        End If
        
    Next
    
    'Encerra transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 132581
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    ParcelaPag_Grava_ChequeAvulso = SUCESSO

    Exit Function

Erro_ParcelaPag_Grava_ChequeAvulso:

    ParcelaPag_Grava_ChequeAvulso = gErr

    Select Case gErr

        Case 132575
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 132576
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 132577, 132578
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PARCELASPAG2", gErr)
        
        Case 132579
            Call Rotina_Erro(vbOKOnly, "ERRO_PARCELA_PAGAR_INEXISTENTE", gErr)
        
        Case 132580
            Call Rotina_Erro(vbOKOnly, "ERRO_MODIFICACAO_PARCELAS_PAGAR", gErr)
           
        Case 132581
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case 132590, 132591
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOS_PAGAR", gErr)
        
        Case 132592
            Call Rotina_Erro(vbOKOnly, "ERRO_TITULOSPAG_NAO_EXISTENTE", gErr, lNumIntTitulo)
        
        Case 132593
            Call Rotina_Erro(vbOKOnly, "ERRO_MODIFICACAO_TITULOS_PAGAR", gErr)
           
        Case 132873 To 132876
            Call Rotina_Erro(vbOKOnly, "ERRO_MODIFICACAO_PARCELAS_PAGAR", gErr)
           
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179160)

    End Select
    
    'Desfaz alterações
    Call Transacao_Rollback

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function ParcelaPag_Grava_Cheque(ByVal objChequesPag As ClassChequesPag) As Long

Dim lErro As Long
Dim lTransacao  As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim objInfoParcPag As ClassInfoParcPag
Dim objInfoChequePag As ClassInfoChequePag
Dim lNumIntTitulo As Long
Dim lNumIntTituloAnt As Long

On Error GoTo Erro_ParcelaPag_Grava_Cheque

    'Abre a transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 132583
    
    'abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 132584
    Next
    
    For Each objInfoChequePag In objChequesPag.ColInfoChequePag
        
        For Each objInfoParcPag In objChequesPag.colInfoParcPag
            If objInfoParcPag.iSeqCheque = objInfoChequePag.iSeqCheque Then Exit For
        Next
    
        lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntTitulo FROM ParcelasPag WHERE NumIntDoc = ? ", 0, lNumIntTitulo, objInfoParcPag.lNumIntParc)
        If lErro <> AD_SQL_SUCESSO Then gError 132585
        
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 132586
        
        If lErro = AD_SQL_SEM_DADOS Then gError 132587
        
        lErro = Comando_ExecutarPos(alComando(2), "UPDATE ParcelasPag SET Favorecido = ?", alComando(1), objInfoChequePag.sFavorecido)
        If lErro <> AD_SQL_SUCESSO Then gError 132588
        
        'Se a parcela é de um título diferente da anterior
        If lNumIntTituloAnt <> lNumIntTitulo Then
        
            'Le o Titulo a Pagar
            lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntDoc FROM TitulosPag WHERE NumIntDoc = ? ", 0, lNumIntTituloAnt, lNumIntTitulo)
            If lErro <> AD_SQL_SUCESSO Then gError 132594
        
            'Traz o primeiro registro
            lErro = Comando_BuscarPrimeiro(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 132595
            
            'Se não encontrou=> Erro
            If lErro = AD_SQL_SEM_DADOS Then gError 132596
        
            'Atualiza a Observação
            lErro = Comando_ExecutarPos(alComando(2), "UPDATE TitulosPag SET Observacao = ?, ChequeRetido = ?", alComando(1), objInfoChequePag.sObservacao, MARCADO)
            If lErro <> AD_SQL_SUCESSO Then gError 132597
            
        End If
    Next
    
    'Encerra transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 132589
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    ParcelaPag_Grava_Cheque = SUCESSO

    Exit Function

Erro_ParcelaPag_Grava_Cheque:

    ParcelaPag_Grava_Cheque = gErr

    Select Case gErr

        Case 132583
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 132584
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 132585, 132586
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PARCELASPAG2", gErr)
        
        Case 132587
            Call Rotina_Erro(vbOKOnly, "ERRO_PARCELA_PAGAR_INEXISTENTE", gErr)
        
        Case 132588
            Call Rotina_Erro(vbOKOnly, "ERRO_MODIFICACAO_PARCELAS_PAGAR", gErr)
           
        Case 132589
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case 132594, 132595
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOS_PAGAR", gErr)
        
        Case 132596
            Call Rotina_Erro(vbOKOnly, "ERRO_TITULOSPAG_NAO_EXISTENTE", gErr, lNumIntTitulo)
        
        Case 132597
            Call Rotina_Erro(vbOKOnly, "ERRO_MODIFICACAO_TITULOS_PAGAR", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179161)

    End Select
    
    'Desfaz alterações
    Call Transacao_Rollback

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function
'#########################################################################

'###########################################################################
'Inserido por Wagner 02/05/2005
Function ChequesPagAvulso_AtualizarBD_Customizado(ByVal objChequesPagAvulso As ClassChequesPagAvulso) As Long
'faz as atualizacoes no bd correspondentes a criacao de um cheque avulso (manual)

Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim objInfoParcPag As ClassInfoParcPag
Dim iIndice As Integer
Dim lNumIntParcela As Long

On Error GoTo Erro_ChequesPagAvulso_AtualizarBD_Customizado
    
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 132881
    Next
    
    For Each objInfoParcPag In objChequesPagAvulso.colInfoParcPag
    
        'pular parcelas com iseqcheque = 0
        If objInfoParcPag.iSeqCheque <> 0 Then
    
            'Buscar as informações dos cobradores no BD
            lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntParcela FROM ChequeHic WHERE NumIntParcela = ?", 0, lNumIntParcela, objInfoParcPag.lNumIntParc)
            If lErro <> AD_SQL_SUCESSO Then gError 132882
        
            lErro = Comando_BuscarPrimeiro(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 132883
        
            'Enquanto encontrar
            If lErro = AD_SQL_SUCESSO Then
    
                'Apagar a inforação do BD
                lErro = Comando_ExecutarPos(alComando(2), "DELETE FROM ChequeHic", alComando(1))
                If lErro <> AD_SQL_SUCESSO Then gError 132884
    
            End If
    
        End If
    
    Next
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    ChequesPagAvulso_AtualizarBD_Customizado = SUCESSO
    
    Exit Function
    
Erro_ChequesPagAvulso_AtualizarBD_Customizado:

    ChequesPagAvulso_AtualizarBD_Customizado = gErr

    Select Case gErr

        Case 132881
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case 132882, 132883
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PARCELASPAG2", gErr)

        Case 132884
            Call Rotina_Erro(vbOKOnly, "ERRO_MODIFICACAO_TITULOS_PAGAR", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179162)

    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function
'#####################################################################


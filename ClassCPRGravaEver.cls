VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCPRGravaEver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

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
Dim sSQLAux As String

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
    
    sSQLAux = "SELECT FilialEmpresa, NumIntDoc, ValorTotal, ValorICMS, CreditoICMS, ValorICMSSubst, ValorFrete, ValorINSS, ValorIPI, CreditoIPI, OutrasDespesas , ValorProdutos, ValorIRRF, ValorSeguro, ValorINSS, INSSRetido, FilialPedCompra, NumPedCompra, PISRetido, ISSRetido, COFINSRetido, CSLLRetido, TaxaMoeda FROM TitulosPag WHERE "
    
    If objTituloPagar.iFilialEmpresa <= DELTA_FILIALREAL_OFICIAL Then
        sSQLAux = sSQLAux & " FilialEmpresa <= ? AND "
    Else
        sSQLAux = sSQLAux & " FilialEmpresa > ? AND "
    End If
        
    sSQLAux = sSQLAux & "Fornecedor = ? AND Filial = ? AND NumTitulo = ? AND DataEmissao = ? AND SiglaDocumento = ? AND Status <> ?"
    
    'Verifica se a NotaFiscalFatura já existe em TitulosPag
    lErro = Comando_ExecutarPos(lComando, sSQLAux, 0, _
        tTituloPagar.iFilialEmpresa, tTituloPagar.lNumIntDoc, tTituloPagar.dValorTotal, _
        tTituloPagar.dValorICMS, tTituloPagar.iCreditoICMS, tTituloPagar.dValorICMSSubst, _
        tTituloPagar.dValorFrete, tTituloPagar.dValorINSS, tTituloPagar.dValorIPI, tTituloPagar.iCreditoIPI, tTituloPagar.dOutrasDespesas, _
        tTituloPagar.dValorProdutos, tTituloPagar.dValorIRRF, tTituloPagar.dValorSeguro, tTituloPagar.dValorINSS, tTituloPagar.iINSSRetido, _
        tTituloPagar.iFilialPedCompra, tTituloPagar.lNumPedCompra, tTituloPagar.dPISRetido, tTituloPagar.dISSRetido, tTituloPagar.dCOFINSRetido, tTituloPagar.dCSLLRetido, tTituloPagar.dTaxaMoeda, _
        DELTA_FILIALREAL_OFICIAL, objTituloPagar.lFornecedor, objTituloPagar.iFilial, objTituloPagar.lNumTitulo, objTituloPagar.dtDataEmissao, objTituloPagar.sSiglaDocumento, STATUS_EXCLUIDO)
    If lErro <> AD_SQL_SUCESSO Then gError 111785

    'Tenta ler
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 111786
    
    'Se for uma alteracao ...
    If lErro <> AD_SQL_SEM_DADOS Then

        'Verifica se a NotaFiscalFatura já existe em TitulosPag
        'lErro = Comando_ExecutarPos(lComando, "SELECT NumIntDoc FROM TitulosPag WHERE NumIntDoc = ?", 0, tTituloPagar.lNumIntDoc, objTituloPagar.lNumIntDoc)
        'If lErro <> AD_SQL_SUCESSO Then gError 18473

        'Tenta ler
'        lErro = Comando_BuscarPrimeiro(lComando)
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 18474
        
        'Se nao encontrou => Erro
 '       If lErro = AD_SQL_SEM_DADOS Then gError 108791
 
        If Not bPodeAlterar Then gError 180335

        'Faz lock no TítuloPagar
        lErro = Comando_LockExclusive(lComando)
        If lErro <> AD_SQL_SUCESSO Then gError 25028
        
        'Se não pertence a essa FilialEmpresa, erro
        If tTituloPagar.iFilialEmpresa <> objTituloPagar.iFilialEmpresa Then Error 18475
        
        'verifica a alteracao dos campos alteraveis , se houve  - - - > altera no BD
        If Abs(tTituloPagar.dTaxaMoeda - objTituloPagar.dTaxaMoeda) > DELTA_VALORMONETARIO2 Or tTituloPagar.iFilialPedCompra <> objTituloPagar.iFilialPedCompra Or tTituloPagar.lNumPedCompra <> objTituloPagar.lNumPedCompra Then
            
            lErro = Comando_ExecutarPos(lComando2, "UPDATE TitulosPag SET TaxaMoeda = ?, FilialPedCompra = ?, NumPedCompra =?, Historico=? ", lComando, objTituloPagar.dTaxaMoeda, objTituloPagar.iFilialPedCompra, objTituloPagar.lNumPedCompra, objTituloPagar.sHistorico)
            If lErro <> AD_SQL_SUCESSO Then gError 48918
        
        End If
        
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
            
            sSQLAux = "SELECT NumIntDoc FROM TitulosPagBaixados WHERE Fornecedor = ? AND Filial = ? AND NumTitulo = ? AND DataEmissao = ? AND SiglaDocumento = ? AND Status <> ? "
            
            If objTituloPagar.iFilialEmpresa <= DELTA_FILIALREAL_OFICIAL Then
                sSQLAux = sSQLAux & " AND FilialEmpresa <= ? "
            Else
                sSQLAux = sSQLAux & " AND FilialEmpresa > ? "
            End If
    
            'Verifica se a NotaFiscalFatura já existe em TitulosPagBaixados
            lErro = Comando_ExecutarPos(lComando, sSQLAux, 0, _
                tTituloPagar.lNumIntDoc, objTituloPagar.lFornecedor, objTituloPagar.iFilial, objTituloPagar.lNumTitulo, objTituloPagar.dtDataEmissao, objTituloPagar.sSiglaDocumento, STATUS_EXCLUIDO, DELTA_FILIALREAL_OFICIAL)
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
        
        'Insere o Título na tabela TitulosPag
        lErro = Comando_Executar(lComando1, "INSERT INTO TitulosPag (NumIntDoc,FilialEmpresa,Fornecedor,Filial,SiglaDocumento,NumTitulo,DataEmissao,Status,Saldo,NumParcelas,ValorTotal,ValorSeguro,ValorFrete,OutrasDespesas,ValorProdutos,ValorICMS,ValorICMSSubst,CreditoICMS,ValorIPI,CreditoIPI,ValorIRRF,ValorINSS, INSSRetido, DataRegistro, Observacao, FilialPedCompra, NumPedCompra, CondicaoPagto, PISRetido, ISSRetido, COFINSRetido, CSLLRetido, TaxaMoeda, Historico) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", tTituloPagar.lNumIntDoc, objTituloPagar.iFilialEmpresa, objTituloPagar.lFornecedor, objTituloPagar.iFilial, objTituloPagar.sSiglaDocumento, objTituloPagar.lNumTitulo, objTituloPagar.dtDataEmissao, STATUS_LANCADO, objTituloPagar.dSaldo, objTituloPagar.iNumParcelas, _
            objTituloPagar.dValorTotal, objTituloPagar.dValorSeguro, objTituloPagar.dValorFrete, objTituloPagar.dOutrasDespesas, objTituloPagar.dValorProdutos, objTituloPagar.dValorICMS, objTituloPagar.dValorICMSSubst, objTituloPagar.iCreditoICMS, objTituloPagar.dValorIPI, objTituloPagar.iCreditoIPI, objTituloPagar.dValorIRRF, objTituloPagar.dValorINSS, objTituloPagar.iINSSRetido, objTituloPagar.dtDataRegistro, objTituloPagar.sObservacao, objTituloPagar.iFilialPedCompra, objTituloPagar.lNumPedCompra, objTituloPagar.iCondicaoPagto, objTituloPagar.dPISRetido, objTituloPagar.dISSRetido, objTituloPagar.dCOFINSRetido, objTituloPagar.dCSLLRetido, objTituloPagar.dTaxaMoeda, objTituloPagar.sHistorico)
        If lErro <> AD_SQL_SUCESSO Then gError 18480

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
        
        Case 18464, 18479, 18481, 18502, 18504, 18505, 18529, 59067, 180335
        
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179000)

    End Select
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    
End Function

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
Dim sSQLAux As String

On Error GoTo Erro_NFPag_Grava_BD

    lComando = 0
    lComando1 = 0
    
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 18430
    
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 18431
    
    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then gError 94358
        
    'Trecho incluido por leo em 25/02/02
    
    sSQLAux = "SELECT NumIntDoc,Fornecedor,Filial,ValorTotal,ValorIRRF, PISRetido, COFINSRetido, CSLLRetido, TaxaMoeda FROM NfsPagBaixadas WHERE "
    
    If objNFsPag.iFilialEmpresa <= DELTA_FILIALREAL_OFICIAL Then
        sSQLAux = sSQLAux & " FilialEmpresa <= ? AND "
    Else
        sSQLAux = sSQLAux & " FilialEmpresa > ? AND "
    End If
            
    sSQLAux = sSQLAux & "Fornecedor = ? AND Filial = ? AND NumNotaFiscal = ? AND DataEmissao = ?"
    
    'Procura em NFsPagBaixadas se existe alguma Nota Fiscal com o índice fornecido pela Tela - Fornecedor, Filial, Numero, DataEmissao
    lErro = Comando_Executar(lComando2, sSQLAux, tNFsPag.lNumIntDoc, tNFsPag.lFornecedor, tNFsPag.iFilial, tNFsPag.dValorTotal, tNFsPag.dValorIRRF, tNFsPag.dPISRetido, tNFsPag.dCOFINSRetido, tNFsPag.dCSLLRetido, tNFsPag.dTaxaMoeda, DELTA_FILIALREAL_OFICIAL, objNFsPag.lFornecedor, objNFsPag.iFilial, objNFsPag.lNumNotaFiscal, objNFsPag.dtDataEmissao)
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
            
            'exibe aviso de alteracao somente da ctb...
            If Rotina_Aviso(vbYesNo, "AVISO_TITBAIXADO_GRAVA_CONTAB_SOMENTE") = vbNo Then gError 120100
            
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
            
    sSQLAux = "SELECT NumIntDoc, FilialEmpresa, DataVencimento, ValorTotal, ValorSeguro, ValorFrete, OutrasDespesas, ValorProdutos, ValorIRRF, ValorICMS, ValorICMSSubst, CreditoICMS, ValorIPI, CreditoIPI, PISRetido, COFINSRetido, CSLLRetido, TaxaMoeda FROM NfsPag WHERE "
    
    If objNFsPag.iFilialEmpresa <= DELTA_FILIALREAL_OFICIAL Then
        sSQLAux = sSQLAux & " FilialEmpresa <= ? AND "
    Else
        sSQLAux = sSQLAux & " FilialEmpresa > ? AND "
    End If
            
    sSQLAux = sSQLAux & "Fornecedor = ? AND Filial = ? AND NumNotaFiscal = ? AND DataEmissao = ? AND Status <> ?"
    
    'Procura em NFsPag se existe alguma Nota Fiscal com o índice fornecido pela Tela - Fornecedor, Filial, Numero, DataEmissao
    lErro = Comando_ExecutarPos(lComando, sSQLAux, 0, tNFsPag.lNumIntDoc, tNFsPag.iFilialEmpresa, tNFsPag.dtDataVencimento, tNFsPag.dValorTotal, tNFsPag.dValorSeguro, tNFsPag.dValorFrete, tNFsPag.dOutrasDespesas, tNFsPag.dValorProdutos, tNFsPag.dValorIRRF, tNFsPag.dValorICMS, tNFsPag.dValorICMSSubst, tNFsPag.iCreditoICMS, tNFsPag.dValorIPI, tNFsPag.iCreditoIPI, tNFsPag.dPISRetido, tNFsPag.dCOFINSRetido, tNFsPag.dCSLLRetido, tNFsPag.dTaxaMoeda, DELTA_FILIALREAL_OFICIAL, objNFsPag.lFornecedor, objNFsPag.iFilial, objNFsPag.lNumNotaFiscal, objNFsPag.dtDataEmissao, STATUS_EXCLUIDO)
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
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_DATAVENCIMENTO_ALTERAVEL")
            If vbMsgRes = vbNo Then gError 18315
           
        End If
   
        'Altera no BD
        lErro = Comando_ExecutarPos(lComando1, "UPDATE NfsPag SET DataVencimento =?, Historico=?", lComando, objNFsPag.dtDataVencimento, objNFsPag.sHistorico)
        If lErro <> AD_SQL_SUCESSO Then gError 18316
    
        objNFsPag.lNumIntDoc = tNFsPag.lNumIntDoc

    'Se não encontrar
    ElseIf lErro = AD_SQL_SEM_DADOS Then
    
        sSQLAux = "SELECT FilialEmpresa FROM NfsPagBaixadas WHERE "
        
        If objNFsPag.iFilialEmpresa <= DELTA_FILIALREAL_OFICIAL Then
            sSQLAux = sSQLAux & " FilialEmpresa <= ? AND "
        Else
            sSQLAux = sSQLAux & " FilialEmpresa > ? AND "
        End If
            
        sSQLAux = sSQLAux & "Fornecedor = ? AND Filial = ? AND NumNotaFiscal = ? AND DataEmissao = ? AND Status <> ?"
        
        'Verifica se a Nota Fiscal está baixada
        lErro = Comando_Executar(lComando, sSQLAux, iFilialEmpresa, DELTA_FILIALREAL_OFICIAL, objNFsPag.lFornecedor, objNFsPag.iFilial, objNFsPag.lNumNotaFiscal, objNFsPag.dtDataEmissao, STATUS_EXCLUIDO)
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
        If objNFsPag.iFilialEmpresa = EMPRESA_TODA Then gError 106554
        objNFsPag.iStatus = STATUS_LANCADO
        
        'Inserir o novo registro em NfsPag
        lErro = Comando_Executar(lComando, "INSERT INTO NfsPag (NumIntDoc,FilialEmpresa,Fornecedor,Filial,NumNotaFiscal,DataEmissao,Status,DataVencimento,ValorTotal,ValorSeguro, ValorFrete, OutrasDespesas, ValorProdutos, ValorIRRF, ValorICMS, ValorICMSSUbst, CreditoICMS, ValorIPI, CreditoIPI,NumIntTitPag, PISRetido, COFINSRetido, CSLLRetido, TaxaMoeda, Historico) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
            objNFsPag.lNumIntDoc, objNFsPag.iFilialEmpresa, objNFsPag.lFornecedor, objNFsPag.iFilial, objNFsPag.lNumNotaFiscal, objNFsPag.dtDataEmissao, objNFsPag.iStatus, objNFsPag.dtDataVencimento, objNFsPag.dValorTotal, objNFsPag.dValorSeguro, objNFsPag.dValorFrete, objNFsPag.dOutrasDespesas, objNFsPag.dValorProdutos, objNFsPag.dValorIRRF, objNFsPag.dValorICMS, objNFsPag.dValorICMSSubst, objNFsPag.iCreditoICMS, objNFsPag.dValorIPI, objNFsPag.iCreditoIPI, objNFsPag.lNumIntTitPag, objNFsPag.dPISRetido, objNFsPag.dCOFINSRetido, objNFsPag.dCSLLRetido, objNFsPag.dTaxaMoeda, objNFsPag.sHistorico)
        If lErro <> AD_SQL_SUCESSO Then gError 18317

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179001)
                
    End Select
        
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    
    Exit Function
        
End Function

Function TituloPagar_Testa_Repeticao(ByVal objTituloPagar As ClassTituloPagar) As Long
'Pesquisa se já existe no BD em um certo Periodo de Emissão uma
'Fatura a Pagar ou Baixada com o mesmo Número
'Obs.: estou desconsiderando a sigla do docto e a filial da empresa

Dim lErro As Long
Dim lComando As Long
Dim dtDataEmissao As Date, sSQLAux As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_TituloPagar_Testa_Repeticao
    
    'Abre o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 18724
    
    sSQLAux = "SELECT DataEmissao FROM TitulosPagBaixados WHERE "
        
    If objTituloPagar.iFilialEmpresa <= DELTA_FILIALREAL_OFICIAL Then
        sSQLAux = sSQLAux & " FilialEmpresa <= ? AND "
    Else
        sSQLAux = sSQLAux & " FilialEmpresa > ? AND "
    End If
        
    If objTituloPagar.dtDataEmissao <> DATA_NULA Then
    
        'Verifica se em um certo Período de Emissão foi cadastrada em TitulosPagBaixados um Título de mesmo número, independente da filialempresa
        sSQLAux = sSQLAux & " Fornecedor = ? AND Filial = ? AND SiglaDocumento = ? AND NumTitulo = ? AND ((DataEmissao > ? AND DataEmissao < ?) OR DataEmissao = ?) AND Status <> ? "
        
        lErro = Comando_Executar(lComando, sSQLAux, dtDataEmissao, DELTA_FILIALREAL_OFICIAL, objTituloPagar.lFornecedor, objTituloPagar.iFilial, objTituloPagar.sSiglaDocumento, objTituloPagar.lNumTitulo, objTituloPagar.dtDataEmissao - PERIODO_EMISSAO, objTituloPagar.dtDataEmissao + PERIODO_EMISSAO, DATA_NULA, STATUS_EXCLUIDO)
        
    Else
    
        'Verifica se foi cadastrada em TitulosPagBaixados um Título de mesmo número independente de data de emissao
        sSQLAux = sSQLAux & " Fornecedor = ? AND Filial = ? AND SiglaDocumento = ? AND NumTitulo = ? AND Status <> ? "
        
        lErro = Comando_Executar(lComando, sSQLAux, dtDataEmissao, DELTA_FILIALREAL_OFICIAL, objTituloPagar.lFornecedor, objTituloPagar.iFilial, objTituloPagar.sSiglaDocumento, objTituloPagar.lNumTitulo, STATUS_EXCLUIDO)
    
    End If
    
    If lErro <> AD_SQL_SUCESSO Then Error 18731
    
    'Tenta ler
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 18732
    
    'Se encontrou
    If lErro = AD_SQL_SUCESSO Then
        
        'Pergunta se deseja cadastrar uma Fatura com o mesmo número
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_FATURAPAG_BAIXADA_MESMO_NUMERO", objTituloPagar.lFornecedor, objTituloPagar.iFilial, objTituloPagar.lNumTitulo, dtDataEmissao)
        
        'Se não, erro
        If vbMsgRes = vbNo Then Error 18733
    
    End If
    
    sSQLAux = "SELECT DataEmissao FROM TitulosPag WHERE "
        
    If objTituloPagar.iFilialEmpresa <= DELTA_FILIALREAL_OFICIAL Then
        sSQLAux = sSQLAux & " FilialEmpresa <= ? AND "
    Else
        sSQLAux = sSQLAux & " FilialEmpresa > ? AND "
    End If
        
    If objTituloPagar.dtDataEmissao <> DATA_NULA Then

        'Verifica se em um certo Período de Emissão foi cadastrada em TitulosPag um Título de mesmo número
        sSQLAux = sSQLAux & " Fornecedor = ? AND Filial = ? AND NumTitulo = ? AND ((DataEmissao > ? AND DataEmissao < ?) OR DataEmissao = ?) AND Status <> ? "
        
        lErro = Comando_Executar(lComando, sSQLAux, dtDataEmissao, DELTA_FILIALREAL_OFICIAL, objTituloPagar.lFornecedor, objTituloPagar.iFilial, objTituloPagar.lNumTitulo, objTituloPagar.dtDataEmissao - PERIODO_EMISSAO, objTituloPagar.dtDataEmissao + PERIODO_EMISSAO, DATA_NULA, STATUS_EXCLUIDO)
        
    Else
    
        'Verifica se foi cadastrada em TitulosPag um Título de mesmo número independente de data de emissao
        sSQLAux = sSQLAux & " Fornecedor = ? AND Filial = ? AND NumTitulo = ? AND Status <> ? "
        
        lErro = Comando_Executar(lComando, sSQLAux, dtDataEmissao, DELTA_FILIALREAL_OFICIAL, objTituloPagar.lFornecedor, objTituloPagar.iFilial, objTituloPagar.lNumTitulo, STATUS_EXCLUIDO)
        
    End If
    
    If lErro <> AD_SQL_SUCESSO Then Error 18728

    'Tenta ler
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 18729
    
    'Se encontrou
    If lErro = AD_SQL_SUCESSO Then
    
        'Pergunta se deseja cadastrar uma Fatura com o mesmo número
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_FATURAPAG_MESMO_NUMERO", objTituloPagar.lFornecedor, objTituloPagar.iFilial, objTituloPagar.lNumTitulo, dtDataEmissao)
        
        'Se não, erro
        If vbMsgRes = vbNo Then Error 18730
    
    End If
    
    'Fecha comando
    Call Comando_Fechar(lComando)
    
    TituloPagar_Testa_Repeticao = SUCESSO
    
    Exit Function

Erro_TituloPagar_Testa_Repeticao:

    TituloPagar_Testa_Repeticao = Err
    
    Select Case Err
    
        Case 18730, 18733
        
        Case 18728, 18729
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FATURA1", Err, objTituloPagar.lNumTitulo)
            
        Case 18731, 18732
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FATURASPAGBAIXADAS", Err, objTituloPagar.lNumTitulo)
    
        Case 18724
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 179002)
            
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function


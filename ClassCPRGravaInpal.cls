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

Const MOTIVO_AJUDA_CUSTO = 4

'1 - Márcio 09/2000 - Incluido os campos ValorOriginal e MotivoDiferenca na gravação das Parcelas de TitulosPag

'colocar no dicinpal
Const ERRO_SALDO_PARCPAG_MENOR_DIF = "O saldo da parcela está ficando menor que a diferenca entre o valor original e o da parcela"

Private Declare Function Comando_PrepararPosInt Lib "ADSQLMN.DLL" Alias "AD_Comando_PrepararPos" (ByVal lComando As Long, ByVal lpSQLStmt As String, ByVal lSelect As Long) As Long
Private Declare Function Comando_PrepararInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Preparar" (ByVal lComando As Long, ByVal lpSQLStmt As String) As Long
Private Declare Function Comando_BindVarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_BindVar" (ByVal lComando As Long, lpVar As Variant) As Long
Private Declare Function Comando_ExecutarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Executar" (ByVal lComando As Long) As Long

Private Sub ParcelasPag_ObtemSaldo(ByVal colParcelaPagar As colParcelaPagar, dSaldo As Double)

Dim lErro As Long, iIndice As Integer

On Error GoTo Erro_ParcelasPag_ObtemSaldo

    dSaldo = 0
    
    'obtem saldo e valor total do titulo
    For iIndice = 1 To colParcelaPagar.Count
    
        dSaldo = dSaldo + Round(colParcelaPagar(iIndice).dSaldo, 2)
        
    Next
        
    Exit Sub
    
Erro_ParcelasPag_ObtemSaldo:

    Select Case Err
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error)
     
    End Select
     
    Exit Sub

End Sub

Function NFFatPag_Grava_BD(ByVal objTituloPagar As ClassTituloPagar, ByVal colParcelaPagar As colParcelaPagar, Optional ByVal bPodeAlterar As Boolean = True) As Long

'Grava NotaFiscalFatura (e outros titulos, menos fatura que cobra nfs simples ) no BD (nível + baixo)
'CHAMADA EM TRANSAÇÃO por NFFatPag_Grava
'???? ATENCAO:tem o parametro adicional iOrigemNFiscal em relacao a versao padrao

Dim lErro As Long
Dim lComando As Long, lComando1 As Long
Dim lComando2 As Long, lComando3 As Long
Dim tTituloPagar As typeTituloPagar
Dim iIndice As Integer
Dim colParcelaPagarBD As New colParcelaPagar
Dim objParcelaPagar As ClassParcelaPagar
Dim objParcelaPagar1 As ClassParcelaPagar
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim vbMsgRes As VbMsgBoxResult, lNumIntNF As Long
Dim iCamposAlterados As Integer, dSaldo As Double
Dim iOrigemNFiscal As Integer

On Error GoTo Erro_NFFatPag_Grava_BD
    
    'Abre os comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 18471

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 18472
    
    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 48919
    
    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then Error 59066
    
    'Verifica se a NotaFiscalFatura já existe em TitulosPag
    lErro = Comando_ExecutarPos(lComando, "SELECT FilialEmpresa, NumIntDoc, ValorTotal, ValorICMS, CreditoICMS, ValorICMSSubst, ValorFrete, ValorINSS, ValorIPI, CreditoIPI, OutrasDespesas , ValorProdutos, ValorIRRF, ValorSeguro, ValorINSS, INSSRetido, FilialPedCompra, NumPedCompra, PISRetido, COFINSRetido, CSLLRetido FROM TitulosPag WHERE Fornecedor = ? AND Filial = ? AND NumTitulo = ? AND DataEmissao = ? AND SiglaDocumento = ? AND Status <> ?", 0, _
        tTituloPagar.iFilialEmpresa, tTituloPagar.lNumIntDoc, tTituloPagar.dValorTotal, _
        tTituloPagar.dValorICMS, tTituloPagar.iCreditoICMS, tTituloPagar.dValorICMSSubst, _
        tTituloPagar.dValorFrete, tTituloPagar.dValorINSS, tTituloPagar.dValorIPI, tTituloPagar.iCreditoIPI, tTituloPagar.dOutrasDespesas, _
        tTituloPagar.dValorProdutos, tTituloPagar.dValorIRRF, tTituloPagar.dValorSeguro, tTituloPagar.dValorINSS, tTituloPagar.iINSSRetido, _
        tTituloPagar.iFilialPedCompra, tTituloPagar.lNumPedCompra, tTituloPagar.dPISRetido, tTituloPagar.dCOFINSRetido, tTituloPagar.dCSLLRetido, objTituloPagar.lFornecedor, objTituloPagar.iFilial, objTituloPagar.lNumTitulo, objTituloPagar.dtDataEmissao, objTituloPagar.sSiglaDocumento, STATUS_EXCLUIDO)
    If lErro <> AD_SQL_SUCESSO Then Error 18473
    
    'Tenta ler
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 18474
    
    'Se existir
    If lErro = AD_SQL_SUCESSO Then

        'Faz lock no TítuloPagar
        lErro = Comando_LockExclusive(lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 25028
        
        'Se não pertence a essa FilialEmpresa, erro
        If tTituloPagar.iFilialEmpresa <> objTituloPagar.iFilialEmpresa Then Error 18475
        
        If iOrigemNFiscal = 0 Then
        
            'testa se o documento foi originado fora de CPR
            lErro = CF("DocumentoCPR_OrigemNFiscal", objTituloPagar.sSiglaDocumento, objTituloPagar.lNumIntDoc, lNumIntNF)
            If lErro <> SUCESSO And lErro <> 41542 Then Error 59224
            
            If lErro <> SUCESSO Then
                iOrigemNFiscal = 0
            Else
                iOrigemNFiscal = 1
            End If
        
        End If
        
        'Coloca o NumIntTitulo lido na coleção de parcelas
        For iIndice = 1 To colParcelaPagar.Count
            colParcelaPagar(iIndice).lNumIntTitulo = tTituloPagar.lNumIntDoc
        Next
        
        'Coloca o NumIntDoc lido no objTituloPagar
        objTituloPagar.lNumIntDoc = tTituloPagar.lNumIntDoc
        
        'Lê no BD as Parcelas referentes ao Título
        lErro = CF("ParcelasPagar_Le", objTituloPagar, colParcelaPagarBD)
        If lErro <> SUCESSO Then Error 18502

        'Verifica se o número de Parcelas da tela é igual ao do BD
        If colParcelaPagarBD.Count <> colParcelaPagar.Count Then Error 18503
                  
        'Flag para alterações
        iCamposAlterados = 0
        
        'Verifica se os campos da Tela não alteráveis foram alterados
        If tTituloPagar.dValorICMS <> objTituloPagar.dValorICMS Or tTituloPagar.iCreditoICMS <> objTituloPagar.iCreditoICMS Or tTituloPagar.dValorICMSSubst <> objTituloPagar.dValorICMSSubst Or _
           tTituloPagar.dValorFrete <> objTituloPagar.dValorFrete Or tTituloPagar.dValorINSS <> objTituloPagar.dValorINSS Or tTituloPagar.dValorIPI <> objTituloPagar.dValorIPI Or tTituloPagar.iCreditoIPI <> objTituloPagar.iCreditoIPI Or _
           tTituloPagar.dOutrasDespesas <> objTituloPagar.dOutrasDespesas Or tTituloPagar.dValorProdutos <> objTituloPagar.dValorProdutos Or tTituloPagar.dValorIRRF <> objTituloPagar.dValorIRRF Or tTituloPagar.dValorSeguro <> objTituloPagar.dValorSeguro Or _
           tTituloPagar.dValorINSS <> objTituloPagar.dValorINSS Or tTituloPagar.iINSSRetido <> objTituloPagar.iINSSRetido Or _
           tTituloPagar.dPISRetido <> objTituloPagar.dPISRetido Or tTituloPagar.dCOFINSRetido <> objTituloPagar.dCOFINSRetido Or tTituloPagar.dCSLLRetido <> objTituloPagar.dCSLLRetido _
           Then iCamposAlterados = 1
                
        If iOrigemNFiscal = 1 And iCamposAlterados = 1 Then Error 59476
        
''        'Verifica se campos do Grid Parcelas não alteráveis foram alterados
''        If iCamposAlterados = 0 Then
''
''            For iIndice = 1 To colParcelaPagar.Count
''                If colParcelaPagar(iIndice).dValor <> colParcelaPagarBD(iIndice).dValor Then iCamposAlterados = 1
''            Next
''
''        End If
        
        'Se houve alteração de campos não alteráveis
        If iCamposAlterados = 1 Then
            
            'Pergunta se deseja prosseguir na alteração para os alteráveis
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_TITULOS_PAGAR_CAMPOS_ALTERAVEIS")
            If vbMsgRes = vbNo Then Error 18504
                        
        End If
                    
        'Altera os campos alteráveis
        lErro = CF("ParcelasPagar_Altera_CamposAlteraveis", colParcelaPagar)
        If lErro <> SUCESSO Then Error 18505

        Call ParcelasPag_ObtemSaldo(colParcelaPagar, dSaldo)
        objTituloPagar.dSaldo = dSaldo
        
'        'verifica a alteracao dos campo que relacionam com Compras ou diferenca, se houve  - - - > altera no BD
'        If tTituloPagar.iFilialPedCompra <> objTituloPagar.iFilialPedCompra Or tTituloPagar.lNumPedCompra <> objTituloPagar.lNumPedCompra Then
            
            lErro = Comando_ExecutarPos(lComando2, "UPDATE TitulosPag SET ValorTotal = ?, Saldo = ?, ValorSeguro = ?, ValorFrete = ?, OutrasDespesas = ?, ValorProdutos = ?, ValorICMS = ?, ValorICMSSubst = ?, CreditoICMS = ?, ValorIPI = ?, CreditoIPI = ?, ValorIRRF = ?, ValorINSS = ?, INSSRetido = ?, FilialPedCompra = ?, NumPedCompra =?, PISRetido = ?, COFINSRetido = ?, CSLLRetido = ?, TaxaMoeda = ?, Historico=?", lComando, objTituloPagar.dValorTotal, objTituloPagar.dSaldo, objTituloPagar.dValorSeguro, objTituloPagar.dValorFrete, objTituloPagar.dOutrasDespesas, objTituloPagar.dValorProdutos, objTituloPagar.dValorICMS, objTituloPagar.dValorICMSSubst, objTituloPagar.iCreditoICMS, objTituloPagar.dValorIPI, objTituloPagar.iCreditoIPI, objTituloPagar.dValorIRRF, objTituloPagar.dValorINSS, objTituloPagar.iINSSRetido, objTituloPagar.iFilialPedCompra, objTituloPagar.lNumPedCompra, objTituloPagar.dPISRetido, objTituloPagar.dCOFINSRetido, objTituloPagar.dCSLLRetido, objTituloPagar.dTaxaMoeda, objTituloPagar.sHistorico)
            If lErro <> AD_SQL_SUCESSO Then Error 48918
        
'        End If
        
    Else  'Nota Fiscal Fatura Pagar não existe na tabela TitulosPag
        
        If iOrigemNFiscal <> 1 Then iOrigemNFiscal = 0
    
        'Testa se há Titulo com mesmo número no Sistema
        lErro = CF("NFFatPag_Testa_Repeticao", objTituloPagar)
        If lErro <> SUCESSO Then Error 18529
            
        objFilialFornecedor.lCodFornecedor = objTituloPagar.lFornecedor
        objFilialFornecedor.iCodFilial = objTituloPagar.iFilial
        
        'Lê a Filial Fornecedor
        lErro = CF("FilialFornecedor_Lock1", objFilialFornecedor)
        If lErro <> SUCESSO And lErro <> 18900 Then Error 18464
    
        'Se não encontrar Filial Fornecedor --> Erro
        If lErro = 18900 Then Error 18465

        If objTituloPagar.iCondicaoPagto <> 0 Then
        
            lErro = CF("CondicaoPagto_Lock", lComando3, objTituloPagar.iCondicaoPagto)
            If lErro <> SUCESSO Then Error 59067
            
        End If
        
        'Gera número interno automático
        lErro = CF("TituloPagar_Automatico", tTituloPagar.lNumIntDoc)
        If lErro <> SUCESSO Then Error 18479
        
        objTituloPagar.lNumIntDoc = tTituloPagar.lNumIntDoc
        objTituloPagar.dtDataRegistro = gdtDataHoje

        'Preenche Saldo = ValorTotal - retencoes
        objTituloPagar.dSaldo = objTituloPagar.dValorTotal - (objTituloPagar.dValorIRRF + objTituloPagar.dPISRetido + objTituloPagar.dCOFINSRetido + objTituloPagar.dCSLLRetido) - IIf(objTituloPagar.iINSSRetido = 1, objTituloPagar.dValorINSS, 0)
        
        'Insere o Título na tabela TitulosPag
        lErro = Comando_Executar(lComando1, "INSERT INTO TitulosPag (NumIntDoc,FilialEmpresa,Fornecedor,Filial,SiglaDocumento,NumTitulo,DataEmissao,Status,Saldo,NumParcelas,ValorTotal,ValorSeguro,ValorFrete,OutrasDespesas,ValorProdutos,ValorICMS,ValorICMSSubst,CreditoICMS,ValorIPI,CreditoIPI,ValorIRRF,ValorINSS, INSSRetido, DataRegistro, Observacao, FilialPedCompra, NumPedCompra, CondicaoPagto, PISRetido, COFINSRetido, CSLLRetido, TaxaMoeda, Historico,Natureza) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", tTituloPagar.lNumIntDoc, objTituloPagar.iFilialEmpresa, objTituloPagar.lFornecedor, objTituloPagar.iFilial, objTituloPagar.sSiglaDocumento, objTituloPagar.lNumTitulo, objTituloPagar.dtDataEmissao, STATUS_LANCADO, objTituloPagar.dSaldo, objTituloPagar.iNumParcelas, _
            objTituloPagar.dValorTotal, objTituloPagar.dValorSeguro, objTituloPagar.dValorFrete, objTituloPagar.dOutrasDespesas, objTituloPagar.dValorProdutos, objTituloPagar.dValorICMS, objTituloPagar.dValorICMSSubst, objTituloPagar.iCreditoICMS, objTituloPagar.dValorIPI, objTituloPagar.iCreditoIPI, objTituloPagar.dValorIRRF, objTituloPagar.dValorINSS, objTituloPagar.iINSSRetido, objTituloPagar.dtDataRegistro, objTituloPagar.sObservacao, objTituloPagar.iFilialPedCompra, objTituloPagar.lNumPedCompra, objTituloPagar.iCondicaoPagto, objTituloPagar.dPISRetido, objTituloPagar.dCOFINSRetido, objTituloPagar.dCSLLRetido, objTituloPagar.dTaxaMoeda, objTituloPagar.sHistorico, objTituloPagar.sNatureza)
        If lErro <> AD_SQL_SUCESSO Then Error 18480

        'Armazena em colParcelaPagar o lNumIntTitulo para cada parcela
        For iIndice = 1 To colParcelaPagar.Count
            colParcelaPagar(iIndice).lNumIntTitulo = tTituloPagar.lNumIntDoc
        Next
        
        'Grava ParcelasPagar no BD
        lErro = CF("ParcelasPagar_Grava", colParcelaPagar)
        If lErro <> SUCESSO Then Error 18481
    
    End If
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    
    NFFatPag_Grava_BD = SUCESSO
    
    Exit Function

Erro_NFFatPag_Grava_BD:

    NFFatPag_Grava_BD = Err
    
    Select Case Err
    
        Case 18465
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_CADASTRADA", Err, objTituloPagar.lFornecedor, objTituloPagar.iFilial)
    
        Case 18471, 18472, 48919, 59066
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 18473, 18474
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFFATURA", Err, objTituloPagar.lNumTitulo)
            
        Case 18475
            Call Rotina_Erro(vbOKOnly, "ERRO_NFFAT_FILIALEMPRESA_DIFERENTE", Err, objTituloPagar.lNumTitulo)
        
        Case 18464, 18479, 18481, 18502, 18504, 18505, 18529, 59067, 59224
        
        Case 18480
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_NFFATURA", Err, objTituloPagar.lNumTitulo)
                    
        Case 18503
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_PARCELAS_ALTERADO", Err, colParcelaPagarBD.Count, colParcelaPagar.Count)
        
        Case 25028
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_TITULOSPAG", Err, objTituloPagar.lNumIntDoc)
            
        Case 48918
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_TITULOSPAG", Err, tTituloPagar.lNumIntDoc)
        
        Case 59476 '??? colocar msg de erro : nào pode alterar valores de nf criada no est
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_TITULOSPAG", Err, tTituloPagar.lNumIntDoc)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error)

    End Select
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    
    Exit Function
    
End Function

'1 - Márcio 09/2000 - Incluido os campos ValorOriginal e MotivoDiferenca na gravação das Parcelas de TitulosPag
'****************************** Inicio ************************************************************
Function ParcelasPagar_Altera_CamposAlteraveis(ByVal colParcelaPagar As colParcelaPagar) As Long
'Altera no BD campos alteráveis das Parcelas de Titulo a Pagar
'cujo lNumIntTitulo é passado dentro de colParcelaPagar

Dim lErro As Long
Dim lComando As Long, dSaldoBD As Double, dNovoSaldo As Double
Dim lComando1 As Long, objParcelaPagar As ClassParcelaPagar
Dim iNumParcela As Integer
Dim lNumIntDoc As Long, dValorBD As Double, dValorOriginalBD As Double

On Error GoTo Erro_ParcelasPagar_Altera_CamposAlteraveis

    'Abre os comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 18507
    
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 18515
        
    'Seleciona as Parcelas a alterar (nao inclui as baixadas)
    lErro = Comando_ExecutarPos(lComando, "SELECT Saldo, NumIntDoc, NumParcela, Valor, ValorOriginal FROM ParcelasPag WHERE NumIntTitulo = ? AND Status <> ? ORDER BY DataVencimento", 0, dSaldoBD, lNumIntDoc, iNumParcela, dValorBD, dValorOriginalBD, colParcelaPagar.Item(1).lNumIntTitulo, STATUS_BAIXADO)
    If lErro <> AD_SQL_SUCESSO Then Error 18508
    
    'Tenta ler a Parcela
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then Error 18509
    
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objParcelaPagar = colParcelaPagar.Item(iNumParcela)
        
        'saldo novo = valor novo - valor ja' baixado
        dNovoSaldo = Round(Round(objParcelaPagar.dValor, 2) - Round(dValorBD - dSaldoBD, 2), 2)
        If dNovoSaldo < 0.01 Then Error 59475
        
        objParcelaPagar.dSaldo = dNovoSaldo
        
        'Altera os campos alteráveis - 1 - 09/2000 Marcio - Incluido o campo MotivoDiferenca
        lErro = Comando_ExecutarPos(lComando1, "UPDATE ParcelasPag SET Valor = ?, ValorOriginal = ?, Saldo = ?, DataVencimento = ?, DataVencimentoReal =?, Status = ?, TipoCobranca = ?, Portador = ? , BancoCobrador=?, MotivoDiferenca=? ", lComando, objParcelaPagar.dValor, objParcelaPagar.dValorOriginal, dNovoSaldo, objParcelaPagar.dtDataVencimento, objParcelaPagar.dtDataVencimentoReal, objParcelaPagar.iStatus, objParcelaPagar.iTipoCobranca, objParcelaPagar.iPortador, objParcelaPagar.iBancoCobrador, objParcelaPagar.iMotivoDiferenca)
        If lErro <> AD_SQL_SUCESSO Then Error 18511
        
        'Tenta ler a Parcela
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then Error 18510
    
    Loop

    'Fecha os comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    ParcelasPagar_Altera_CamposAlteraveis = SUCESSO
    
    Exit Function
    
Erro_ParcelasPagar_Altera_CamposAlteraveis:
    
    ParcelasPagar_Altera_CamposAlteraveis = Err
    
    Select Case Err
    
        Case 18507, 18515
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 18508
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PARCELASPAG", Err, colParcelaPagar(iNumParcela).lNumIntTitulo)
       
        Case 18509, 18510
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PARCELASPAG1", Err, colParcelaPagar(iNumParcela).lNumIntTitulo, iNumParcela)
    
        Case 18511
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_PARCELASPAG", Err, colParcelaPagar(iNumParcela).lNumIntTitulo, iNumParcela)
        
        Case 59475 '???? trocar msg Jones 14/09/00
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_PARCELASPAG", Err, colParcelaPagar(iNumParcela).lNumIntTitulo, iNumParcela)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Private Sub ParcelaPag_MoverCamposObj(ByVal objParcelaPagar As ClassParcelaPagar, tParcelaPagar As typeParcelaPagar)

    objParcelaPagar.lNumIntDoc = tParcelaPagar.lNumIntDoc
    objParcelaPagar.dSaldo = tParcelaPagar.dSaldo
    objParcelaPagar.dtDataVencimento = tParcelaPagar.dtDataVencimento
    objParcelaPagar.dtDataVencimentoReal = tParcelaPagar.dtDataVencimentoReal
    objParcelaPagar.dValor = tParcelaPagar.dValor
    objParcelaPagar.dValorOriginal = tParcelaPagar.dValorOriginal
    objParcelaPagar.iMotivoDiferenca = tParcelaPagar.iMotivoDiferenca
    objParcelaPagar.iBancoCobrador = tParcelaPagar.iBancoCobrador
    objParcelaPagar.iNumParcela = tParcelaPagar.iNumParcela
    objParcelaPagar.iPortador = tParcelaPagar.iPortador
    objParcelaPagar.iProxSeqBaixa = tParcelaPagar.iProxSeqBaixa
    objParcelaPagar.iStatus = tParcelaPagar.iStatus
    objParcelaPagar.iTipoCobranca = tParcelaPagar.iTipoCobranca
    objParcelaPagar.lNumIntTitulo = tParcelaPagar.lNumIntTitulo
    objParcelaPagar.sNossoNumero = tParcelaPagar.sNossoNumero
    objParcelaPagar.sCodigoDeBarras = tParcelaPagar.sCodigoDeBarras

End Sub

' *** CUSTOMIZAÇÃO DE BAIXA DE COMISSÕES P/INPAL - INÍCIO ***
' *** MOTIVO: INCLUIR A BAIXA DA AJUDA DE CUSTO ***
' *** AUTOR: TULIO JORGE SUPERVISOR: LUIZ NOGUEIRA ***
' *** DATA: 17/05/02 A 24/05/02 ***
Public Function ComissoesVendasBaixar(ByVal objComissoesPag As ClassComissoesPag) As Long
'Atualiza as comissões de acordo com os parâmetros passados
'objComissoesPag RECEBE(Input) os parâmetros indicando se é para baixar ou cancelar baixa das comissões,
'e os critérios que devem ser utilizados para selecionar as comissões que serão alteradas

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_ComissoesVendasBaixar

    'Abre transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 102066
    
    'seta a filial como empresa toda... somente para inpal
    objComissoesPag.iFilialEmpresa = EMPRESA_TODA
    
    'Atualiza os registros referentes a ajuda de custo
    lErro = AjudaCustoInpal_Atualiza(objComissoesPag)
    If lErro <> SUCESSO Then gError 101654
    
    'Executa a atualização na tabela de comissões
    lErro = CF("ComissoesVendasBaixar_EmTrans", objComissoesPag)
    If lErro <> SUCESSO Then gError 102073
    
    'confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 102072

    ComissoesVendasBaixar = SUCESSO
    
    Exit Function
    
Erro_ComissoesVendasBaixar:

    ComissoesVendasBaixar = gErr
    
    Select Case gErr
    
        Case 101654, 102073
    
        Case 102066
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 102072
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
            
    End Select
    
    Call Transacao_Rollback
    
    Exit Function

End Function

Private Function AjudaCustoInpal_Atualiza(ByVal objComissoesPag As ClassComissoesPag) As Long
'Atualiza os registros referentes a ajuda de custo
'objComissoesPag RECEBE(Input) os parâmetros indicando se é para baixar ou cancelar baixa de ajuda de custo,
'e os critérios que devem ser utilizados para selecionar as comissões / ajuda de custos que serão alteradas
'ATENÇÃO: ESSA FUNÇÃO SÓ DEVE SER CHAMADA EM TRANSAÇÃO!!!

Dim asSelect(0 To 2) As String
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim tComissoesPagVariant As typeComissoesPagVariantInpal
Dim sNomeTabela As String
Dim objComissoesAvulsas As ClassComissoesAvulsas
Dim lErro As Long

On Error GoTo Erro_AjudaCustoInpal_Atualiza

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 101647
    Next

    'Monta os selects necessarios para atualizar a ajuda de custos
    lErro = AjudaCustoInpal_Monta_Select(objComissoesPag, asSelect())
    If lErro <> SUCESSO Then gError 101665
    
    'Se a data da baixa for nula, significa que esta cancelando o pagamento da ajudas de custos
    If objComissoesPag.dtDataBaixa = DATA_NULA Then
        
        'Guarda no type o filtro que irá garantir a leitura apenas das ajudas de custos baixadas
        tComissoesPagVariant.viStatusAnterior = STATUS_BAIXADO
        
        'Exclui as ajudas de custo de acordo com o filtro passado
        lErro = AjudaCustoInpal_Exclui(objComissoesPag, asSelect(2))
        If lErro <> SUCESSO Then gError 101646
    
    'Senão, significa que esta efetuando o pagamento das ajudas de custos
    Else
    
        'Guarda no type o filtro que irá garantir a leitura apenas das comissões liberadas para servirem
        'como base de cálculo para o valor da ajuda de custo
        tComissoesPagVariant.viStatusAnterior = STATUS_LIBERADO
        
        'Guarda no type o status que será gravado para as ajudas de custos que serão inseridas no BD
        tComissoesPagVariant.viStatusNovo = STATUS_BAIXADO
        
        'Transfere os dados do objeto para o type correspondente com as variaveis variant
        With tComissoesPagVariant
            .vdtComisGeradasAte = objComissoesPag.dtComisGeradasAte
            .vdtComisGeradasDe = objComissoesPag.dtComisGeradasDe
            .viCodVendedorFim = objComissoesPag.iCodVendedorFim
            .viCodVendedorIni = objComissoesPag.iCodVendedorIni
            .viTipo = objComissoesPag.iTipo
            .vlNumIntComissao = objComissoesPag.lNumIntComissao
            .vdtDataBaixa = objComissoesPag.dtDataBaixa
        End With
        
        'Le na tabela ComissoesInpalPlan os vendedores que possuem ajuda de custos
        lErro = AjudaCusto_Executa_Select1(alComando(0), asSelect(0), tComissoesPagVariant)
        If lErro <> SUCESSO Then gError 101648
    
        'Seta o nome da tabela que ira ser mostrada em caso de erro
        sNomeTabela = "ComissoesInpalPlan"
    
        'se posiciona no 1o registro do conjunto obtido anteriormente
        lErro = Comando_BuscarPrimeiro(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 101649
        
        'Enquanto houverem vendedores com ajuda de custo
        Do While lErro <> AD_SQL_SEM_DADOS
            
            'se a ajuda de custo for minima, significa que o vendedor só tem direito a receber ajuda de custo,
            'caso o valor desta seja maior do que o valor de comissões
            'Portanto, é necessário calcular o valor exato da ajuda de custo, que é igual à diferença do valor
            'mínimo de ajuda de custo, menos o valor de comissões que o vendedor tem direito a receber
            If CInt(tComissoesPagVariant.viTipoAjudaCusto) = AJUDACUSTO_MINIMA Then
        
                'Seta o nome da tabela que ira ser mostrada em caso de erro
                sNomeTabela = "Comissoes"
                
                'le o valor total de comissoes a ser recebido pelo vendedor atual do loop, de acordo
                'com os parametros passados pelo usuario
                lErro = AjudaCusto_Executa_Select2(alComando(1), asSelect(1), tComissoesPagVariant)
                If lErro <> SUCESSO Then gError 101650
                
                'zera o valor total de comissoes
                tComissoesPagVariant.vdValorTotalComissoes = 0
                
                'Se posiciona no primeiro registro do conjunto obtido anteriormente
                lErro = Comando_BuscarPrimeiro(alComando(1))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 101651
                
                'se o valor total das comissoes retornado anteriormente for menor que a ajuda de custo
                'obs. se o buscar_primeiro de alcomando(1) retornar sem dados, o valor do total
                'das comissoes ficara como 0, logo, sem dados nao foi tratado
                If CDbl(tComissoesPagVariant.vdAjudaCusto) - CDbl(tComissoesPagVariant.vdValorTotalComissoes) > DELTA_VALORMONETARIO Then
        
                    'Calcula o complemento referente a ajuda custo
                    tComissoesPagVariant.vdValorTotalComissoes = CDbl(tComissoesPagVariant.vdAjudaCusto) - CDbl(tComissoesPagVariant.vdValorTotalComissoes)
                
                End If
                    
            End If
                
            'instancia objComissoesAvulsas
            Set objComissoesAvulsas = New ClassComissoesAvulsas
                
            'Guarda os dados que serao gravados em objComissoesAvulsas
            With tComissoesPagVariant
                objComissoesAvulsas.iVendedor = CInt(.viVendedor)
                objComissoesAvulsas.dtData = CDate(.vdtDataBaixa)
                objComissoesAvulsas.iCodigoMotivo = MOTIVO_AJUDA_CUSTO
                objComissoesAvulsas.dBaseCalculo = CDbl(.vdValorTotalComissoes)
                objComissoesAvulsas.dAliquota = 1
                objComissoesAvulsas.dValorComissao = CDbl(.vdValorTotalComissoes)
            End With
        
            'Grava no BD um registro referente a ajuda de custo
            lErro = CF("ComissoesAvulsas_Grava_EmTrans", objComissoesAvulsas, objComissoesPag.dtDataBaixa)
            If lErro <> SUCESSO Then gError 101652
        
            'Seta o nome da tabela que vai ser exibido em caso de erro
            sNomeTabela = "ComissoesInpalPlan"
        
            'se posiciona no proximo registro do conjunto de planilhas
            lErro = Comando_BuscarProximo(alComando(0))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 101653
                    
        Loop
        
    End If
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    AjudaCustoInpal_Atualiza = SUCESSO
    
    Exit Function
    
Erro_AjudaCustoInpal_Atualiza:
    
    AjudaCustoInpal_Atualiza = gErr
    
    Select Case gErr

        Case 101646, 101648, 101650, 101652, 101665

        Case 101647
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 101649, 101651, 101653
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA", gErr, sNomeTabela)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
            
    End Select

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Private Function AjudaCustoInpal_Monta_Select(ByVal objComissoesPag As ClassComissoesPag, asSelect() As String) As Long
'Monta os selects necessarios para atualizar a ajuda de custos
'objComissoesPag eh parametro de INPUT
'asSelect eh parametro de OUTPUT

On Error GoTo Erro_AjudaCustoInpal_Monta_Select

    'Monta o select que le os vendedores que terao ajuda de custo atualizada
    Call AjudaCustoInpal_Monta_Select1(objComissoesPag, asSelect(0))
    
    'Monta o select que le o valor total das comissoes para os vendedores que terao
    'ajuda de custo atualizada
    Call AjudaCustoInpal_Monta_Select2(objComissoesPag, asSelect(1))
    
    'Monta o select que le as comissoes referentes a ajuda de custo que serao
    'excluidas caso seja uma operacao de cancelamento...
    Call AjudaCustoInpal_Monta_Select3(objComissoesPag, asSelect(2))
        
    AjudaCustoInpal_Monta_Select = SUCESSO
    
    Exit Function
    
Erro_AjudaCustoInpal_Monta_Select:
    
    AjudaCustoInpal_Monta_Select = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
            
    End Select

    Exit Function
    
End Function

Private Sub AjudaCustoInpal_Monta_Select1(ByVal objComissoesPag As ClassComissoesPag, sSelect As String)
'Monta o select que ira ler os vendedores que terao a ajuda de custo atualizada
'objComissoesPag eh parametro de INPUT
'sSelect eh parametro de OUTPUT

Dim sSelectBuffer As String

On Error GoTo Erro_AjudaCustoInpal_Monta_Select1

    'Monta um select para ler em comissoesinpalplan os campos: Vendedor, TipoAjudaCusto e AjudaCusto
    sSelectBuffer = "SELECT DISTINCT Vendedor, AjudaCusto, TipoAjudaCusto FROM ComissoesInpalPlan WHERE AjudaCusto > 0"

    'se o vendedor inicial for preenchido
    If objComissoesPag.iCodVendedorIni > 0 Then
        
        'Inclui um filtro por vendedor inicial na clausula WHERE
        sSelectBuffer = sSelectBuffer & " AND Vendedor >= ?"
        
    End If
    
    'se o vendedor final for preenchido
    If objComissoesPag.iCodVendedorFim > 0 Then
    
        'Adiciona o Vendedor final depois do Vendedor inicial
        sSelectBuffer = sSelectBuffer & " AND Vendedor <= ?"
            
    End If
    
    'retorna em sSelect o conteudo da string montada durante essa rotina
    sSelect = sSelectBuffer

    Exit Sub

Erro_AjudaCustoInpal_Monta_Select1:
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
            
    End Select

    Exit Sub
    
End Sub

Private Sub AjudaCustoInpal_Monta_Select2(ByVal objComissoesPag As ClassComissoesPag, sSelect As String)
'Monta o select que le o valor total das comissoes para
'os vendedores que terao ajuda de custo atualizadas
'objComissoesPag eh parametro de INPUT
'sSelect eh parametro de OUTPUT

Dim sSelectBuffer As String

On Error GoTo Erro_AjudaCustoInpal_Monta_Select2

    'Monta um select para ler em Comissoes, o valor do grupo de
    'comissoes que servira de base para calcular a ajuda de custo
    sSelectBuffer = "SELECT CodVendedor, SUM(Valor) as ValorTotal FROM Comissoes WHERE Status = ? AND DataGeracao >= ? AND DataGeracao <= ? "
    
    'Inclui um filtro por vendedor na clausula WHERE
    sSelectBuffer = sSelectBuffer & " AND CodVendedor = ?"
    
    'se esta baixando pela emissao, inclui os 4 filtros para os tipos de titulos
    If objComissoesPag.iTipo = COMISSAO_EMISSAO Then sSelectBuffer = sSelectBuffer & " AND ( TipoTitulo = ? OR TipoTitulo = ? OR TipoTitulo = ? OR TipoTitulo = ?)"
    
    'se esta baixando pela baixa, inclui 1 filtro para o tipo de titulo
    If objComissoesPag.iTipo = COMISSAO_BAIXA Then sSelectBuffer = sSelectBuffer & " AND TipoTitulo = ?"
    
    'Retorna em sSelect o select montado nessa rotina adicionando a clausula group by ao final...
    sSelect = sSelectBuffer & " Group By CodVendedor"
        
    Exit Sub
        
Erro_AjudaCustoInpal_Monta_Select2:
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
            
    End Select

    Exit Sub

End Sub

Private Sub AjudaCustoInpal_Monta_Select3(ByVal objComissoesPag As ClassComissoesPag, sSelect As String)
'Monta o select que le le as comissoes referentes a ajuda de custo que
'serao excluidas, caso seja uma operacao de cancelamento
'os vendedores que terao ajuda de custo atualizadas
'objComissoesPag eh parametro de INPUT
'sSelect eh parametro de OUTPUT

Dim sSelectBuffer As String

On Error GoTo Erro_AjudaCustoInpal_Monta_Select3

    'Monta o select que le as comissoes referentes a ajuda de custo que serao
    'excluidas, caso seja uma operacao de cancelamento
    sSelectBuffer = "SELECT ComissoesAvulsas.NumIntDoc FROM ComissoesAvulsas, Comissoes WHERE ComissoesAvulsas.NumIntDoc = Comissoes.NumIntDoc AND Comissoes.DataGeracao >= ? AND Comissoes.DataGeracao <= ? AND ComissoesAvulsas.CodigoMotivo = ? AND Comissoes.Status = ? "
    
    'Inclui um filtro por vendedor na clausula WHERE
    'se foi definido um vendedor inicial
    If objComissoesPag.iCodVendedorIni <> 0 Then sSelectBuffer = sSelectBuffer & " AND CodVendedor >= ?"
    
    'Inclui um filtro por vendedor na clausula WHERE
    'se foi definido um vendedor final
    If objComissoesPag.iCodVendedorFim <> 0 Then sSelectBuffer = sSelectBuffer & " AND CodVendedor <= ?"
    
    'Retorna em sSelect o select montado nessa rotina
    sSelect = sSelectBuffer
    
    Exit Sub
    
Erro_AjudaCustoInpal_Monta_Select3:
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
            
    End Select

    Exit Sub

End Sub

Private Function AjudaCustoInpal_Exclui(ByVal objComissoesPag As ClassComissoesPag, ByVal sSelect As String) As Long
'Exclui as ajudas de custo de acordo com o filtro passado
'objComissoesPag eh parametro de INPUT
'sSelect eh parametro de INPUT
'ATENÇÃO: ESSA FUNÇÃO DEVE SER CHAMADA EM TRANSAÇÃO!!!

Dim lErro As Long
Dim lComando As Long
Dim tComissoesPagVariant As typeComissoesPagVariantInpal
Dim objComissoesAvulsas As ClassComissoesAvulsas

On Error GoTo Erro_AjudaCustoInpal_Exclui

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 101641

    'copia os filtros necessarios para o type
    tComissoesPagVariant.vdtComisGeradasDe = objComissoesPag.dtComisGeradasDe
    tComissoesPagVariant.vdtComisGeradasAte = objComissoesPag.dtComisGeradasAte
    tComissoesPagVariant.viCodigoMotivo = MOTIVO_AJUDA_CUSTO
    tComissoesPagVariant.viStatusAnterior = STATUS_BAIXADO
    tComissoesPagVariant.viCodVendedorIni = objComissoesPag.iCodVendedorIni
    tComissoesPagVariant.viCodVendedorFim = objComissoesPag.iCodVendedorFim
        
    'Executa o select que le os campos NumIntDoc e CodigoMotivo das tabelas Comissoes e
    'ComissoesAvulsas, selecionando os registros referentes a ajuda de custos que
    'se enquadram detro do periodo determinado pelo usuario
    lErro = AjudaCusto_Executa_Select3(lComando, sSelect, tComissoesPagVariant)
    If lErro <> SUCESSO Then gError 101642
    
    'Busca o primeiro elemento do conjunto
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 101643
    
    'Enquanto exisitrem comissoes no conjunto
    Do While lErro <> AD_SQL_SEM_DADOS
    
        'cria uma nova instancia de objComissoesAvulsas
        Set objComissoesAvulsas = New ClassComissoesAvulsas
    
        'Guarda o NumIntDoc do registro no obj
        objComissoesAvulsas.lNumIntDoc = tComissoesPagVariant.vlNumIntComissao
    
        'Exclui a comissao de ComissoesAvulsas e Comissoes
        lErro = CF("ComissoesAvulsas_Exclui_NumIntDoc", objComissoesAvulsas)
        If lErro <> SUCESSO Then gError 101645
    
        'Busca a proxima do conjunto de comissoes
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 101644
    
    Loop
    
    'fecha o comando
    Call Comando_Fechar(lComando)
    
    AjudaCustoInpal_Exclui = SUCESSO

    Exit Function

Erro_AjudaCustoInpal_Exclui:

    AjudaCustoInpal_Exclui = gErr
    
    Select Case gErr
    
        Case 101641
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 101642, 101645
        
        Case 101643, 101644
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA", gErr, "Comissoes/ComissoesAvulsas")
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
            
    End Select
    
    'fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Private Function AjudaCusto_Executa_Select3(lComando As Long, sSelect As String, tComissoesPag As typeComissoesPagVariantInpal) As Long
'Executa o select que le os campos numintdoc e o codigomotivo nas tabelas comissoes e
'comissoes avulsas, selecionando registros referentes a ajuda de custos e se enquadram
'no periodo determinado pelo usuario...
'lComando eh parametro de INPUT
'tComissoesPag eh parametro de INPUT

Dim lErro As Long

On Error GoTo Erro_AjudaCusto_Executa_Select3

    'executa a preparacao do comando select
    lErro = Comando_PrepararInt(lComando, sSelect)
    If lErro <> AD_SQL_SUCESSO Then gError 101656
    
    'faz o bind das variaveis que serao lidas...
    tComissoesPag.vlNumIntComissao = CLng(tComissoesPag.vlNumIntComissao)
    lErro = Comando_BindVarInt(lComando, tComissoesPag.vlNumIntComissao)
    If lErro <> AD_SQL_SUCESSO Then gError 101657
    
    lErro = Comando_BindVarInt(lComando, tComissoesPag.vdtComisGeradasDe)
    If lErro <> AD_SQL_SUCESSO Then gError 101667
    
    lErro = Comando_BindVarInt(lComando, tComissoesPag.vdtComisGeradasAte)
    If lErro <> AD_SQL_SUCESSO Then gError 101668
    
    'binda cod motivo
    lErro = Comando_BindVarInt(lComando, tComissoesPag.viCodigoMotivo)
    If lErro <> AD_SQL_SUCESSO Then gError 101658
    
    'binda o status
    lErro = Comando_BindVarInt(lComando, tComissoesPag.viStatusAnterior)
    If lErro <> AD_SQL_SUCESSO Then gError 101687
    
    'faz o bind da variavel que passa o filtro vendedor
    If tComissoesPag.viCodVendedorIni <> 0 Then
        lErro = Comando_BindVarInt(lComando, tComissoesPag.viCodVendedorIni)
        If lErro <> AD_SQL_SUCESSO Then gError 101659
    End If
    
    If tComissoesPag.viCodVendedorFim <> 0 Then
        lErro = Comando_BindVarInt(lComando, tComissoesPag.viCodVendedorFim)
        If lErro <> AD_SQL_SUCESSO Then gError 101660
    End If
    
    'executa finalmente o select
    lErro = Comando_ExecutarInt(lComando)
    If lErro <> AD_SQL_SUCESSO Then gError 101662
    
    AjudaCusto_Executa_Select3 = SUCESSO
    
    Exit Function
    
Erro_AjudaCusto_Executa_Select3:

    AjudaCusto_Executa_Select3 = gErr
    
    Select Case gErr

        Case 101656 To 101662, 101667, 101668, 101687
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA", gErr, "ComissoesAvulsas/Comissoes")
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Function AjudaCusto_Executa_Select1(lComando As Long, sSelect As String, tComissoesPag As typeComissoesPagVariantInpal) As Long
'Executa o select que ira ler os vendedores que terao a ajuda de custo atualizada
'lComando eh parametro de INPUT
'tComissoesPag eh parametro de INPUT

Dim lErro As Long

On Error GoTo Erro_AjudaCusto_Executa_Select1

    'executa a preparacao do comando select
    lErro = Comando_PrepararPosInt(lComando, sSelect, 0)
    If lErro <> AD_SQL_SUCESSO Then gError 101672
    
    'faz o bind das variaveis que serao lidas...
    tComissoesPag.viVendedor = CInt(tComissoesPag.viVendedor)
    lErro = Comando_BindVarInt(lComando, tComissoesPag.viVendedor)
    If lErro <> AD_SQL_SUCESSO Then gError 101669
    
    tComissoesPag.vdAjudaCusto = CDbl(tComissoesPag.vdAjudaCusto)
    lErro = Comando_BindVarInt(lComando, tComissoesPag.vdAjudaCusto)
    If lErro <> AD_SQL_SUCESSO Then gError 101670
    
    tComissoesPag.viTipoAjudaCusto = CInt(tComissoesPag.viTipoAjudaCusto)
    lErro = Comando_BindVarInt(lComando, tComissoesPag.viTipoAjudaCusto)
    If lErro <> AD_SQL_SUCESSO Then gError 101671
        
    'faz o bind da variavel que passa o filtro vendedor
    If tComissoesPag.viCodVendedorIni <> 0 Then
        lErro = Comando_BindVarInt(lComando, tComissoesPag.viCodVendedorIni)
        If lErro <> AD_SQL_SUCESSO Then gError 101672
    End If
    
    If tComissoesPag.viCodVendedorFim <> 0 Then
        lErro = Comando_BindVarInt(lComando, tComissoesPag.viCodVendedorFim)
        If lErro <> AD_SQL_SUCESSO Then gError 101673
    End If
    
    'executa finalmente o select
    lErro = Comando_ExecutarInt(lComando)
    If lErro <> AD_SQL_SUCESSO Then gError 101674
    
    AjudaCusto_Executa_Select1 = SUCESSO
    
    Exit Function
    
Erro_AjudaCusto_Executa_Select1:

    AjudaCusto_Executa_Select1 = gErr
    
    Select Case gErr

        Case 101669 To 101674
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA", gErr, "ComissoesInpalPlan")
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Function AjudaCusto_Executa_Select2(lComando As Long, sSelect As String, tComissoesPag As typeComissoesPagVariantInpal) As Long
'Executa o select que le o valor total das comissoes para os vendedores que terao ajuda de custo atualizadas
'no periodo determinado pelo usuario...
'lComando eh parametro de INPUT
'tComissoesPag eh parametro de INPUT

Dim lErro As Long
Dim vVendedor As Variant
Dim avTipoEmissao(1 To 4) As Variant
Dim vTipoBaixa As Variant

 
On Error GoTo Erro_AjudaCusto_Executa_Select2

    'executa a preparacao do comando select
    lErro = Comando_PrepararInt(lComando, sSelect)
    If lErro <> AD_SQL_SUCESSO Then gError 101675
    
    'copia o conteudo do campo vendedor do type para a variavel vendedor...
    'motivo: usar vvendedor  como filtro..
    vVendedor = CInt(tComissoesPag.viVendedor)
    
    'faz o bind das variaveis que serao lidas...
    tComissoesPag.viVendedor = CInt(tComissoesPag.viVendedor)
    lErro = Comando_BindVarInt(lComando, tComissoesPag.viVendedor)
    If lErro <> AD_SQL_SUCESSO Then gError 101676
    
    tComissoesPag.vdValorTotalComissoes = CDbl(tComissoesPag.vdValorTotalComissoes)
    lErro = Comando_BindVarInt(lComando, tComissoesPag.vdValorTotalComissoes)
    If lErro <> AD_SQL_SUCESSO Then gError 101677
        
    'binda o status
    lErro = Comando_BindVarInt(lComando, tComissoesPag.viStatusAnterior)
    If lErro <> AD_SQL_SUCESSO Then gError 101678
        
    'binda as datas
    lErro = Comando_BindVarInt(lComando, tComissoesPag.vdtComisGeradasDe)
    If lErro <> AD_SQL_SUCESSO Then gError 101679
    
    lErro = Comando_BindVarInt(lComando, tComissoesPag.vdtComisGeradasAte)
    If lErro <> AD_SQL_SUCESSO Then gError 101680
        
    'faz o bind da variavel que passa o filtro vendedor
    lErro = Comando_BindVarInt(lComando, vVendedor)
    If lErro <> AD_SQL_SUCESSO Then gError 101681
    
    'faz os binds em caso da comissao ser na emissao
    If tComissoesPag.viTipo = COMISSAO_EMISSAO Then
        
        avTipoEmissao(1) = TIPO_NF
        avTipoEmissao(2) = TIPO_TITULO_RECEBER
        avTipoEmissao(3) = TIPO_DEBITO
        avTipoEmissao(4) = TIPO_COMISSAO_AVULSA
                
        lErro = Comando_BindVarInt(lComando, avTipoEmissao(1))
        If lErro <> AD_SQL_SUCESSO Then gError 101682
        
        lErro = Comando_BindVarInt(lComando, avTipoEmissao(2))
        If lErro <> AD_SQL_SUCESSO Then gError 101683
        
        lErro = Comando_BindVarInt(lComando, avTipoEmissao(3))
        If lErro <> AD_SQL_SUCESSO Then gError 101684
    
        lErro = Comando_BindVarInt(lComando, avTipoEmissao(4))
        If lErro <> AD_SQL_SUCESSO Then gError 101685
    
    End If
    
    'faz os binds em caso da comissao ser na baixa
    If tComissoesPag.viTipo = COMISSAO_BAIXA Then
        
        vTipoBaixa = TIPO_PARCELA
        
        lErro = Comando_BindVarInt(lComando, vTipoBaixa)
        If lErro <> AD_SQL_SUCESSO Then gError 101686
        
    End If
    
    'executa finalmente o select
    lErro = Comando_ExecutarInt(lComando)
    If lErro <> AD_SQL_SUCESSO Then gError 101662
    
    AjudaCusto_Executa_Select2 = SUCESSO
    
    Exit Function
    
Erro_AjudaCusto_Executa_Select2:

    AjudaCusto_Executa_Select2 = gErr
    
    Select Case gErr

        Case 101675 To 101686
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA", gErr, "Comissoes")
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Function TituloPagar_Testa_Repeticao(ByVal objTituloPagar As ClassTituloPagar) As Long
'Pesquisa se já existe no BD em um certo Periodo de Emissão uma
'Fatura a Pagar ou Baixada com o mesmo Número
'Obs.: estou desconsiderando a sigla do docto e a filial da empresa

Dim lErro As Long
Dim lComando As Long
Dim dtDataEmissao As Date
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_TituloPagar_Testa_Repeticao
    
    'Abre o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 18724
    
    If objTituloPagar.dtDataEmissao <> DATA_NULA Then
    
        'Verifica se em um certo Período de Emissão foi cadastrada em TitulosPagBaixados um Título de mesmo número
        lErro = Comando_Executar(lComando, "SELECT DataEmissao FROM TitulosPagBaixados WHERE Fornecedor = ? AND Filial = ? AND SiglaDocumento = ? AND NumTitulo = ? AND ((DataEmissao > ? AND DataEmissao < ?) OR DataEmissao = ?) AND Status <> ? ", dtDataEmissao, objTituloPagar.lFornecedor, objTituloPagar.iFilial, objTituloPagar.sSiglaDocumento, objTituloPagar.lNumTitulo, objTituloPagar.dtDataEmissao - PERIODO_EMISSAO, objTituloPagar.dtDataEmissao + PERIODO_EMISSAO, DATA_NULA, STATUS_EXCLUIDO)
        
    Else
    
        'Verifica se foi cadastrada em TitulosPagBaixados um Título de mesmo número independente de data de emissao
        lErro = Comando_Executar(lComando, "SELECT DataEmissao FROM TitulosPagBaixados WHERE Fornecedor = ? AND Filial = ? AND SiglaDocumento = ? AND NumTitulo = ? AND Status <> ? ", dtDataEmissao, objTituloPagar.lFornecedor, objTituloPagar.iFilial, objTituloPagar.sSiglaDocumento, objTituloPagar.lNumTitulo, STATUS_EXCLUIDO)
    
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
    
    If objTituloPagar.dtDataEmissao <> DATA_NULA Then

        'Verifica se em um certo Período de Emissão foi cadastrada em TitulosPag um Título de mesmo número
        lErro = Comando_Executar(lComando, "SELECT DataEmissao FROM TitulosPag WHERE Fornecedor = ? AND Filial = ? AND NumTitulo = ? AND ((DataEmissao > ? AND DataEmissao < ?) OR DataEmissao = ?) AND Status <> ? ", dtDataEmissao, objTituloPagar.lFornecedor, objTituloPagar.iFilial, objTituloPagar.lNumTitulo, objTituloPagar.dtDataEmissao - PERIODO_EMISSAO, objTituloPagar.dtDataEmissao + PERIODO_EMISSAO, DATA_NULA, STATUS_EXCLUIDO)
        
    Else
    
        'Verifica se foi cadastrada em TitulosPag um Título de mesmo número independente de data de emissao
        lErro = Comando_Executar(lComando, "SELECT DataEmissao FROM TitulosPag WHERE Fornecedor = ? AND Filial = ? AND NumTitulo = ? AND Status <> ? ", dtDataEmissao, objTituloPagar.lFornecedor, objTituloPagar.iFilial, objTituloPagar.lNumTitulo, STATUS_EXCLUIDO)
        
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error)
            
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function RelForSaldo_Prepara(ByVal iFilialEmpresa As Integer, lNumIntRel As Long, ByVal lFornInic As Long, ByVal lFornFinal As Long, ByVal dtDataSaldo As Date) As Long
'dtDataSaldo eh a data inicial do relatorio (ex. 01/01/2002)

Dim lErro As Long, dtData As Date, dtDataJaCalculada As Date
Dim lTransacao As Long, alComando(0 To 8) As Long, iIndice As Integer, lFornAnterior As Long, dSaldo As Double
Dim iTipoLcto As Integer, lFornecedor As Long, dValorTotal As Double, dValorIRRF As Double, dValorINSS As Double, iINSSRetido As Integer, dBaixasParcPag_ValorBaixado As Double, iBaixasParcPag_Status As Integer, dtBaixasParcPag_DataCancelamento As Date
Dim dtDataSaldoIni As Date, lNumIntRelAux As Long
Dim dSaldoIni As Double
Dim lForn As Long
Dim iTipoReg As Integer
Dim iStatus As Integer
Dim dValorTitulo As Double
Dim dValorBaixa As Double, bAchou As Boolean
Dim objSaldo As AdmlCodigoNome, iFilialEmp As Integer
Dim colSaldos As New Collection

On Error GoTo Erro_RelForSaldo_Prepara

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 81792
    Next

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 81793

    'obtem numintrel
    lErro = CF("Config_ObterNumInt", "CPConfig", "NUM_PROX_REL_FOR_SALDO", lNumIntRel)
    If lErro <> SUCESSO Then gError 81794
        
    If (glEmpresa = 1 Or glEmpresa = 3 Or glEmpresa = 4 Or glEmpresa = 5) And dtDataSaldo = #1/1/2002# Then
        
        lErro = Comando_Executar(alComando(8), "INSERT INTO RelForSaldo (NumIntRel, Fornecedor, Saldo) SELECT ?, Fornecedor, Saldo FROM RelForSaldoIni WHERE FilialEmpresa = 0 AND Data = '01-01-2002'", lNumIntRel)
        If lErro <> SUCESSO Then gError 111739
        
    Else
    
        If (glEmpresa = 1 Or glEmpresa = 3 Or glEmpresa = 4 Or glEmpresa = 5) And dtDataSaldo < #1/1/2002# Then gError 106616
                
        dtData = dtDataSaldo - 1
        dtDataJaCalculada = DATA_NULA
    
        iFilialEmp = IIf(iFilialEmpresa = EMPRESA_TODA Or gobjCTB.giContabCentralizada <> 0, EMPRESA_TODA, iFilialEmpresa)
        
        'deleta o saldo de datas posteriores
        lErro = Comando_Executar(alComando(3), "DELETE FROM RelForSaldoIni WHERE FilialEmpresa = ? AND Data >= ?", iFilialEmp, dtDataSaldo)
        If lErro <> AD_SQL_SUCESSO Then gError 111741
        
        'pesquisar em relforsaldoini registros da filial em questao
        lErro = Comando_Executar(alComando(2), "SELECT Data, Saldo, Fornecedor FROM RelForSaldoIni WHERE FilialEmpresa = ? AND Data < ? ORDER BY Data DESC", _
            dtDataSaldoIni, dSaldoIni, lForn, iFilialEmp, dtDataSaldo)
        If lErro <> SUCESSO Then gError 111739
        
        'busca o primeiro saldo
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 111740
    
        'enquanto existirem saldos
        Do While lErro <> AD_SQL_SEM_DADOS
            
            If dtDataJaCalculada = DATA_NULA Then
                dtDataJaCalculada = dtDataSaldoIni
            End If
                
            If dtDataSaldoIni < dtDataJaCalculada Then
                Exit Do
            End If
            
            'instancia o objeto
            Set objSaldo = New AdmlCodigoNome
            
            'guarda o saldo no obj...
            objSaldo.sNome = CStr(dSaldoIni)
            
            'guarda o fonecedor
            objSaldo.lCodigo = lForn
            
            'adiciona o obj na colecao onde a chave eh "X" concatenado com o codigo do fornecedor
            colSaldos.Add objSaldo, "X" & lForn
            
            'busca o proximo
            lErro = Comando_BuscarProximo(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 111742
        
        Loop
            
        lErro = CF("RelLctosCPAux_Prepara", lNumIntRelAux, dtDataJaCalculada, dtData)
        If lErro <> SUCESSO Then gError 111741
        
        'se achou registro
        If dtDataJaCalculada <> DATA_NULA Then
            
            lErro = Comando_Executar(alComando(4), "SELECT Fornecedor, TipoReg, StatusLcto, SUM(BaixasParcPag_ValorDiferenca -  BaixasParcPag_ValorBaixado) as ValorBaixa, SUM(ValorTotal - ValorDiferenca - (ValorIRRF+PISRetido+COFINSRetido+CSLLRetido) - ( INSSRetido * ValorINSS )) as ValorTitulo FROM RelLctosCP WHERE NumIntRel = ? AND Data >= ? AND Data < ? GROUP BY Fornecedor, TipoReg, StatusLcto", _
                                    lForn, iTipoReg, iStatus, dValorBaixa, dValorTitulo, lNumIntRelAux, dtDataJaCalculada, dtDataSaldo)
            
        'senao
        Else
        
            'select analogo ao anterior, mas sem data >= ?
            lErro = Comando_Executar(alComando(4), "SELECT Fornecedor, TipoReg, StatusLcto, SUM(BaixasParcPag_ValorDiferenca -  BaixasParcPag_ValorBaixado) as ValorBaixa, SUM(ValorTotal - ValorDiferenca - (ValorIRRF+PISRetido+COFINSRetido+CSLLRetido) - ( INSSRetido * ValorINSS )) as ValorTitulo FROM RelLctosCP WHERE NumIntRel = ? AND Data < ? GROUP BY Fornecedor, TipoReg, StatusLcto", _
                                    lForn, iTipoReg, iStatus, dValorBaixa, dValorTitulo, lNumIntRelAux, dtDataSaldo)
        
        End If
            
        If lErro <> AD_SQL_SUCESSO Then gError 111743
            
        'busca o registro do select obtido anteriormente
        lErro = Comando_BuscarPrimeiro(alComando(4))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 111744
        
        Do While lErro <> AD_SQL_SEM_DADOS
        
            bAchou = False
            For Each objSaldo In colSaldos
                If objSaldo.lCodigo = lForn Then
                    bAchou = True
                    Exit For
                End If
            Next
            
            If bAchou = False Then
            
                'instancia o objeto
                Set objSaldo = New AdmlCodigoNome
                
                'guarda o saldo no obj...
                objSaldo.sNome = CStr(0)
                
                'guarda o fonecedor
                objSaldo.lCodigo = lForn
                
                'adiciona o obj na colecao onde a chave eh "X" concatenado com o codigo do fornecedor
                colSaldos.Add objSaldo, "X" & CStr(lForn)
                
            End If
            
            'atualiza os saldos de acordo com o tiporeg e o statuslcto
            If iTipoReg < 9 And iStatus <> 2 Then
                objSaldo.sNome = CStr(StrParaDbl(objSaldo.sNome) + dValorTitulo)
            ElseIf iTipoReg >= 9 And iStatus = 2 Then
                objSaldo.sNome = CStr(StrParaDbl(objSaldo.sNome) - dValorBaixa)
            ElseIf iTipoReg >= 9 And iStatus <> 2 Then
                objSaldo.sNome = CStr(StrParaDbl(objSaldo.sNome) + dValorBaixa)
            ElseIf iTipoReg < 9 And iStatus = 2 Then
                objSaldo.sNome = CStr(StrParaDbl(objSaldo.sNome) - dValorTitulo)
            End If
            
            lErro = Comando_BuscarProximo(alComando(4))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 111744
        
        Loop
    
        lErro = Comando_Executar(alComando(7), "DELETE FROM RelLCtosCPAux WHERE NumIntRel = ?", lNumIntRelAux)
        If lErro <> AD_SQL_SUCESSO Then gError 111741
        
        For Each objSaldo In colSaldos
        
            'percorrer a colecao inserindo valores em relforsaldoini
            lErro = Comando_Executar(alComando(5), "INSERT INTO RelForSaldoIni (FilialEmpresa, Data, Fornecedor, Saldo) VALUES (?,?,?,?)", iFilialEmp, dtDataSaldo, objSaldo.lCodigo, StrParaDbl(objSaldo.sNome))
            If lErro <> AD_SQL_SUCESSO Then gError 111745
            
            'percorrer a colecao inserindo valores em relforsaldo
            If StrParaDbl(objSaldo.sNome) <> 0 Then
                lErro = Comando_Executar(alComando(6), "INSERT INTO RelForSaldo (NumIntRel, Fornecedor, Saldo) VALUES (?,?,?)", lNumIntRel, objSaldo.lCodigo, StrParaDbl(objSaldo.sNome))
                If lErro <> AD_SQL_SUCESSO Then gError 111746
            End If
            
        Next
    
    End If
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 81800
    
    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next
   
    RelForSaldo_Prepara = SUCESSO
     
    Exit Function
    
Erro_RelForSaldo_Prepara:

    RelForSaldo_Prepara = gErr
     
    Select Case gErr
          
        Case 106616
            Call Rotina_Erro(vbOKOnly, "ERRO_INPAL_DATA_INICIAL_INVALIDA", gErr)
        
        Case 81794
        
        Case 81792
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
 
        Case 81793
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 81795, 81796, 81797
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_POSCPDATALCTOS", gErr)
        
        Case 81798, 81799
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RELFORSALDO", gErr)
        
        Case 81800
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case 111739, 111740, 111742
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELFORSALDOINI", gErr)
            
        Case 111741
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_RELFORSALDOINI", gErr)

        Case 111743, 111744
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELLCTOSCP", gErr)

        Case 111745
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RELFORSALDOINI", gErr)

        Case 111746
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RELFORSALDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    Call Transacao_Rollback
    
    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next
   
    Exit Function

End Function

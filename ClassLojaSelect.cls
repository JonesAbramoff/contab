VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassLojaSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Private Declare Function Comando_BindVarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_BindVar" (ByVal lComando As Long, lpVar As Variant) As Long
Private Declare Function Comando_PrepararInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Preparar" (ByVal lComando As Long, ByVal lpSQLStmt As String) As Long
Private Declare Function Comando_ExecutarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Executar" (ByVal lComando As Long) As Long

Public Function Cheque_NumIntCheque_Le(ByVal objChequePre As ClassChequePre, ByVal bExcluido As Boolean) As Long
'essa função lê um cheque pelo seu nuumintcheque na tabela chequepre

Dim lErro As Long
Dim lComando As Long
Dim tChequePre As typeChequePre

On Error GoTo Erro_Cheque_NumIntCheque_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 113607
    
    'inicializa as strings
    tChequePre.sAgencia = String(STRING_AGENCIA, 0)
    tChequePre.sContaCorrente = String(STRING_CONTA_CORRENTE, 0)
    tChequePre.sCPFCGC = String(STRING_CGC, 0)
    
    'executo o select
    lErro = Comando_Executar(lComando, "SELECT Cliente, Filial, Banco, Agencia, ContaCorrente, Numero, DataDeposito, Valor, NumBordero, CPFCGC, NumMovtoCaixa, NumMovtoSangria, Aprovado, NaoEspecificado, NumBorderoLoja, SequencialLoja, SequencialBack, FilialEmpresaLoja, Status, TipoBordero, Caixa, SequencialCaixa, NumIntExt, DataEmissao " & _
                                       "FROM ChequePre " & _
                                       "WHERE NumIntCheque=? " & _
                                       "AND FilialEmpresa=?", _
                                       tChequePre.lCliente, tChequePre.iFilial, tChequePre.iBanco, tChequePre.sAgencia, tChequePre.sContaCorrente, tChequePre.lNumero, tChequePre.dtDataDeposito, tChequePre.dValor, tChequePre.lNumBordero, tChequePre.sCPFCGC, tChequePre.lNumMovtoCaixa, tChequePre.lNumMovtoSangria, tChequePre.iAprovado, tChequePre.iNaoEspecificado, tChequePre.lNumBorderoLoja, tChequePre.lSequencialLoja, tChequePre.lSequencialBack, tChequePre.iFilialEmpresaLoja, tChequePre.iStatus, tChequePre.iTipoBordero, tChequePre.iCaixa, tChequePre.lSequencialCaixa, tChequePre.lNumIntExt, tChequePre.dtDataEmissao, _
                                       objChequePre.lNumIntCheque, _
                                       objChequePre.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 113608
    
    'busco o primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 113609
    
    'se não encontrar erro
    If lErro = AD_SQL_SEM_DADOS Then gError 113610
    
    'preencho o objChequePre com os dados do BD
    objChequePre.lCliente = tChequePre.lCliente
    objChequePre.iFilial = tChequePre.iFilial
    objChequePre.iBanco = tChequePre.iBanco
    objChequePre.sAgencia = tChequePre.sAgencia
    objChequePre.sContaCorrente = tChequePre.sContaCorrente
    objChequePre.lNumero = tChequePre.lNumero
    objChequePre.dtDataDeposito = tChequePre.dtDataDeposito
    objChequePre.dValor = tChequePre.dValor
    objChequePre.lNumBordero = tChequePre.lNumBordero
    objChequePre.sCPFCGC = tChequePre.sCPFCGC
    objChequePre.lNumMovtoCaixa = tChequePre.lNumMovtoCaixa
    objChequePre.lNumMovtoSangria = tChequePre.lNumMovtoSangria
    objChequePre.iAprovado = tChequePre.iAprovado
    objChequePre.iNaoEspecificado = tChequePre.iNaoEspecificado
    objChequePre.lNumBorderoLoja = tChequePre.lNumBorderoLoja
    objChequePre.lSequencialLoja = tChequePre.lSequencialLoja
    objChequePre.lSequencialBack = tChequePre.lSequencialBack
    objChequePre.iFilialEmpresaLoja = tChequePre.iFilialEmpresaLoja
    objChequePre.iStatus = tChequePre.iStatus
    objChequePre.iTipoBordero = tChequePre.iTipoBordero
    objChequePre.iCaixa = tChequePre.iCaixa
    objChequePre.lSequencialCaixa = tChequePre.lSequencialCaixa
    objChequePre.lNumIntExt = tChequePre.lNumIntExt
    objChequePre.dtDataEmissao = tChequePre.dtDataEmissao

    Cheque_NumIntCheque_Le = SUCESSO
    
    Exit Function
    
Erro_Cheque_NumIntCheque_Le:

    Cheque_NumIntCheque_Le = gErr
    
    Select Case gErr
    
        Case 113607
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 113608, 113609
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CHEQUEPRE", gErr)
        
        Case 113610
            'será tratado na rotina chamadora
    
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function TiposMovtoCaixa_Le_Codigo(ByVal objTipoMovtoCaixa As ClassTipoMovtoCaixa) As Long

On Error GoTo Erro_TiposMovtoCaixa_Le_Codigo

    Call objTipoMovtoCaixa.Copiar(gobjLoja.colTiposMovtoCaixa.Item(CStr(objTipoMovtoCaixa.iCodigo)))

    TiposMovtoCaixa_Le_Codigo = SUCESSO
    
    Exit Function
    
Erro_TiposMovtoCaixa_Le_Codigo:
    
    TiposMovtoCaixa_Le_Codigo = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151403)
    
    End Select
    
    Exit Function

End Function

Public Function TiposMovtoCaixa_Le_Todos(colTiposMovtoCaixa As ColTipoMovtoCaixa) As Long

Dim lErro As Long
Dim lComando As Long
Dim tTipoMovtoCaixa As typeTipoMovtoCaixa

On Error GoTo Erro_TiposMovtoCaixa_Le_Todos
    
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 113573
    
    'inicializa as strings
    tTipoMovtoCaixa.sDescricao = String(STRING_TIPOMOVTOCAIXA_DESCRICAO, 0)
    tTipoMovtoCaixa.sNomeReduzido = String(STRING_TIPOMOVTOCAIXA_NOMERED, 0)
    tTipoMovtoCaixa.sSigla = String(STRING_TIPOMOVTOCAIXA_SIGLA, 0)
    tTipoMovtoCaixa.sSignificadoRefInterna = String(STRING_TIPOMOVTOCAIXA_SIGN_REF_INT, 0)
    
    'executa o select
    lErro = Comando_Executar(lComando, "SELECT Codigo, Descricao, NomeReduzido, Sigla, SignificadoRefInterna, Credito, PodeCaixaECF, PodeCaixaCentral, PodeCaixaGeral, RelMovCaixa, TipoMeioPagto " & _
                                       "FROM TiposMovtoCaixa ", _
                                       tTipoMovtoCaixa.iCodigo, tTipoMovtoCaixa.sDescricao, tTipoMovtoCaixa.sNomeReduzido, tTipoMovtoCaixa.sSigla, tTipoMovtoCaixa.sSignificadoRefInterna, tTipoMovtoCaixa.iCredito, tTipoMovtoCaixa.iPodeCaixaECF, tTipoMovtoCaixa.iPodeCaixaCentral, tTipoMovtoCaixa.iPodeCaixaGeral, tTipoMovtoCaixa.iRelMovCaixa, tTipoMovtoCaixa.iTipoMeioPagto)
    If lErro <> AD_SQL_SUCESSO Then gError 113578
    
    'busca o primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 113579
    
    'se nào encontrar nenhum-> erro
    If lErro = AD_SQL_SEM_DADOS Then gError 113575
    
    Set colTiposMovtoCaixa = New ColTipoMovtoCaixa
    
    Do While lErro <> AD_SQL_SEM_DADOS
    
        'adiciona o mesmo à coleção de forma indexada
        Call colTiposMovtoCaixa.Add(tTipoMovtoCaixa.iCodigo, _
                                    tTipoMovtoCaixa.sDescricao, _
                                    tTipoMovtoCaixa.sNomeReduzido, _
                                    tTipoMovtoCaixa.sSigla, _
                                    tTipoMovtoCaixa.sSignificadoRefInterna, _
                                    tTipoMovtoCaixa.iCredito, _
                                    tTipoMovtoCaixa.iPodeCaixaECF, _
                                    tTipoMovtoCaixa.iPodeCaixaCentral, _
                                    tTipoMovtoCaixa.iPodeCaixaGeral, _
                                    tTipoMovtoCaixa.iRelMovCaixa, _
                                    tTipoMovtoCaixa.iTipoMeioPagto, _
                                    CStr(tTipoMovtoCaixa.iCodigo))
    
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 113580
    
    Loop
    
    Call Comando_Fechar(lComando)
    
    TiposMovtoCaixa_Le_Todos = SUCESSO
    
    Exit Function
    
Erro_TiposMovtoCaixa_Le_Todos:
    
    TiposMovtoCaixa_Le_Todos = gErr
    
    Select Case gErr
    
        Case 113578, 113579, 113580
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOSMOVTOCAIXA", gErr)
            
        Case 113575
            'será tratado na rotina chamadora
    
        Case 113573
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151404)
    
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function VendedorFilial_Le(ByVal objVendedor As ClassVendedor, iFilialEmpresa As Integer) As Long
'Le todos os dados da tabela Vendedores, correspondente ao codigo do vendedor passado por parametro
'e de uma filial específica

Dim lComando As Long
Dim lErro As Long
Dim tVendedor As typeVendedor

On Error GoTo Erro_VendedorFilial_Le

    'Abre comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 109495

    'Prepara variáveis string para receber conteúdo do BD
    tVendedor.sNome = String(STRING_VENDEDOR_NOME, 0)
    tVendedor.sNomeReduzido = String(STRING_VENDEDOR_NOME_REDUZIDO, 0)
    tVendedor.sMatricula = String(STRING_VENDEDOR_MATRICULA, 0)
    tVendedor.sAgencia = String(STRING_VENDEDOR_AGENCIA, 0)
    tVendedor.sContaCorrente = String(STRING_VENDEDOR_CONTA_CORRENTE, 0)

    'Seleciona os dados da tabela Vendedores
    lErro = Comando_Executar(lComando, "SELECT Nome, NomeReduzido, Endereco, Tipo, Matricula, CodRegiao, PercComissao, PercComissaoBaixa, PercComissaoEmissao, ComissaoSobreTotal, ComissaoFrete, ComissaoICM, ComissaoIPI, ComissaoSeguro, Banco, Agencia, ContaCorrente " & _
                                       "FROM Vendedores, VendedorLoja " & _
                                       "WHERE Vendedores.Codigo = VendedorLoja.Vendedor " & _
                                       "AND Codigo = ? " & _
                                       "AND FilialEmpresa = ?", _
                                       tVendedor.sNome, tVendedor.sNomeReduzido, tVendedor.lEndereco, tVendedor.iTipo, tVendedor.sMatricula, tVendedor.iCodRegiao, tVendedor.dPercComissao, tVendedor.dPercComissaoBaixa, tVendedor.dPercComissaoEmissao, tVendedor.iComissaoSobreTotal, tVendedor.iComissaoFrete, tVendedor.iComissaoICM, tVendedor.iComissaoIPI, tVendedor.iComissaoSeguro, tVendedor.iBanco, tVendedor.sAgencia, tVendedor.sContaCorrente, _
                                       objVendedor.iCodigo, _
                                       iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 109496

    'Tenta ler
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 109497

    'Vendedor não existe ---> erro
    If lErro = AD_SQL_SEM_DADOS Then gError 109498

    'Passa dados lidas para objVendedor
    objVendedor.sNome = tVendedor.sNome
    objVendedor.sNomeReduzido = tVendedor.sNomeReduzido
    objVendedor.lEndereco = tVendedor.lEndereco
    objVendedor.iTipo = tVendedor.iTipo
    objVendedor.sMatricula = tVendedor.sMatricula
    objVendedor.iCodRegiao = tVendedor.iCodRegiao
    objVendedor.dPercComissao = tVendedor.dPercComissao
    objVendedor.dPercComissaoBaixa = tVendedor.dPercComissaoBaixa
    objVendedor.dPercComissaoEmissao = tVendedor.dPercComissaoEmissao
    objVendedor.iComissaoSobreTotal = tVendedor.iComissaoSobreTotal
    objVendedor.iComissaoFrete = tVendedor.iComissaoFrete
    objVendedor.iComissaoICM = tVendedor.iComissaoICM
    objVendedor.iComissaoIPI = tVendedor.iComissaoIPI
    objVendedor.iComissaoSeguro = tVendedor.iComissaoSeguro
    objVendedor.iBanco = tVendedor.iBanco
    objVendedor.sAgencia = tVendedor.sAgencia
    objVendedor.sContaCorrente = tVendedor.sContaCorrente

    'Fecha comando
    Call Comando_Fechar(lComando)

    VendedorFilial_Le = SUCESSO

    Exit Function

Erro_VendedorFilial_Le:

    VendedorFilial_Le = gErr

    Select Case gErr

        Case 109495
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 109496, 109497
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_VENDEDOR", gErr, objVendedor.iCodigo)

        Case 109498 'Não existe o vendedor (sem dados)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151405)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function VendedorFilial_Le_NomeReduzido(ByVal objVendedor As ClassVendedor, iFilialEmpresa As Integer) As Long
'Lê os dados do Vendedor, correspondente ao NomeReduzido passado por parâmetro
'de uma dada filial

Dim lComando As Long
Dim lErro As Long
Dim tVendedor As typeVendedor

On Error GoTo Erro_VendedorFilial_Le_NomeReduzido

    'Abre comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 109499

    'Prepara variáveis tipo String para receber do BD
    tVendedor.sNome = String(STRING_VENDEDOR_NOME, 0)
    tVendedor.sMatricula = String(STRING_VENDEDOR_MATRICULA, 0)
    tVendedor.sAgencia = String(STRING_VENDEDOR_AGENCIA, 0)
    tVendedor.sContaCorrente = String(STRING_VENDEDOR_CONTA_CORRENTE, 0)

    'Seleciona Vendedor correspondente ao NomeReduzido
    lErro = Comando_Executar(lComando, "SELECT Codigo, Nome, Endereco, Tipo, Matricula, CodRegiao, PercComissao, PercComissaoBaixa, PercComissaoEmissao, ComissaoSobreTotal, ComissaoFrete, ComissaoICM, ComissaoIPI, Banco, Agencia, ContaCorrente " & _
                                       "FROM Vendedores, VendedorLoja " & _
                                       "WHERE Vendedores.Codigo = VendedorLoja.Vendedor " & _
                                       "AND NomeReduzido = ? " & _
                                       "AND FilialEmpresa = ?", _
                                       tVendedor.iCodigo, tVendedor.sNome, tVendedor.lEndereco, tVendedor.iTipo, tVendedor.sMatricula, tVendedor.iCodRegiao, tVendedor.dPercComissao, tVendedor.dPercComissaoBaixa, tVendedor.dPercComissaoEmissao, tVendedor.iComissaoSobreTotal, tVendedor.iComissaoFrete, tVendedor.iComissaoICM, tVendedor.iComissaoIPI, tVendedor.iBanco, tVendedor.sAgencia, tVendedor.sContaCorrente, _
                                       objVendedor.sNomeReduzido, _
                                       iFilialEmpresa)

    If lErro <> AD_SQL_SUCESSO Then gError 113501

    'Tenta ler
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 113502

    'Se não existe ---> erro
    If lErro = AD_SQL_SEM_DADOS Then gError 113503

    'Passa dados lidos para objVendedor
    objVendedor.iCodigo = tVendedor.iCodigo
    objVendedor.sNome = tVendedor.sNome
    objVendedor.lEndereco = tVendedor.lEndereco
    objVendedor.iTipo = tVendedor.iTipo
    objVendedor.sMatricula = tVendedor.sMatricula
    objVendedor.iCodRegiao = tVendedor.iCodRegiao
    objVendedor.dPercComissao = tVendedor.dPercComissao
    objVendedor.dPercComissaoBaixa = tVendedor.dPercComissaoBaixa
    objVendedor.dPercComissaoEmissao = tVendedor.dPercComissaoEmissao
    objVendedor.iComissaoSobreTotal = tVendedor.iComissaoFrete
    objVendedor.iComissaoFrete = tVendedor.iComissaoFrete
    objVendedor.iComissaoICM = tVendedor.iComissaoICM
    objVendedor.iComissaoIPI = tVendedor.iComissaoIPI
    objVendedor.iBanco = tVendedor.iBanco
    objVendedor.sAgencia = tVendedor.sAgencia
    objVendedor.sContaCorrente = tVendedor.sContaCorrente

    'Fecha comando
    lErro = Comando_Fechar(lComando)

    VendedorFilial_Le_NomeReduzido = SUCESSO

    Exit Function

Erro_VendedorFilial_Le_NomeReduzido:

    VendedorFilial_Le_NomeReduzido = gErr

    Select Case gErr

        Case 109499
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 113501, 113502
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_VENDEDOR1", gErr, objVendedor.sNomeReduzido)

        Case 113503 'Não existe o vendedor (sem dados)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151406)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

'Rafael Menezes( Início ) - Tela: BorderoOUTROS
Public Function BorderoOutros_Le(ByVal objBorderoOutros As ClassBorderoOutros) As Long
'Função que recebe um objBorderoOutros com filialempresa e numbordero preenchidos e retorna objeto totalmente preenchido

Dim lErro As Long
Dim lComando As Long
Dim tBorderoOutros As typeBorderoOutros

On Error GoTo Erro_BorderoOutros_Le

    'abre o comando
    lComando = Comando_Abrir
    If lComando = 0 Then gError 108056

    'seleciona o bordero
    lErro = Comando_Executar(lComando, "SELECT AdmMeioPagto,Parcelamento,DataEnvio,DataImpressao,DataBackoffice,Valor,NumIntDocCPR " & _
                                       "FROM BorderoOutros " & _
                                       "WHERE FilialEmpresa=? AND NumBordero=?", _
                                       tBorderoOutros.iAdmMeioPagto, tBorderoOutros.iParcelamento, tBorderoOutros.dtDataEnvio, tBorderoOutros.dtDataImpressao, tBorderoOutros.dtDataBackoffice, tBorderoOutros.dValor, tBorderoOutros.lNumIntDocCPR, _
                                       objBorderoOutros.iFilialEmpresa, objBorderoOutros.lNumBordero)
    If lErro <> AD_SQL_SUCESSO Then gError 108057

    'busca o primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 108058

    'se nào encontrar-> erro
    If lErro = AD_SQL_SEM_DADOS Then gError 108059

    'preenche os dados do bordero
    objBorderoOutros.iAdmMeioPagto = tBorderoOutros.iAdmMeioPagto
    objBorderoOutros.iParcelamento = tBorderoOutros.iParcelamento
    objBorderoOutros.dtDataEnvio = tBorderoOutros.dtDataEnvio
    objBorderoOutros.dtDataImpressao = tBorderoOutros.dtDataImpressao
    objBorderoOutros.dtDataBackoffice = tBorderoOutros.dtDataBackoffice
    objBorderoOutros.dValor = tBorderoOutros.dValor
    objBorderoOutros.lNumIntDocCPR = tBorderoOutros.lNumIntDocCPR

    'fecha o comando
    Call Comando_Fechar(lComando)

    BorderoOutros_Le = SUCESSO

    Exit Function

Erro_BorderoOutros_Le:

    BorderoOutros_Le = gErr

    Select Case gErr

        Case 108056
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 108057, 108058
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BORDEROOUTROS", objBorderoOutros.iFilialEmpresa, objBorderoOutros.lNumBordero)

        Case 108059
            'será tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151407)

    End Select

    'fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function
'Rafael Menezes( FIM ) - Tela: BorderoOUTROS

'Rafael Menezes( Início ) - Tela: BorderoValeTicket
Public Function AdmMeioPagto_Le_TipoMeioPagto(ByVal iTipoMeioPagto As Integer, ByVal colAdmMeioPagto As Collection) As Long
'Função que lê as admmeiopagto de um determinado tipomeiopagto

Dim lErro As Long
Dim lComando As Long
Dim tAdmMeioPagto As typeAdmMeioPagto
Dim objAdmMeioPagto As ClassAdmMeioPagto

On Error GoTo Erro_AdmMeioPagto_Le_TipoMeioPagto

    'abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 107362

    tAdmMeioPagto.sAgencia = String(STRING_ADMMEIOPAGTO_AGENCIA, 0)
    tAdmMeioPagto.sNome = String(STRING_ADMMEIOPAGTO_NOME, 0)

    'seleciona as admmeiopagto com o tipo especificado
    lErro = Comando_Executar(lComando, "SELECT Codigo,FilialEmpresa,Nome,Rede,TaxaVista,TaxaParcelado,TipoMeioPagto,DefasagemPagtoVista,ContaCorrenteInterna,GeraTituloRec,Endereco,CodBanco,Agencia " & _
                                       "FROM AdmMeioPagto " & _
                                       "WHERE TipoMeioPagto=? AND FilialEmpresa=? AND Ativo = ? " & _
                                       "ORDER BY Codigo", _
                                       tAdmMeioPagto.iCodigo, tAdmMeioPagto.iFilialEmpresa, tAdmMeioPagto.sNome, tAdmMeioPagto.iRede, tAdmMeioPagto.dTaxaVista, tAdmMeioPagto.dTaxaParcelado, tAdmMeioPagto.iTipoMeioPagto, tAdmMeioPagto.iDefasagemPagtoVista, tAdmMeioPagto.iContaCorrenteInterna, tAdmMeioPagto.iGeraTituloRec, tAdmMeioPagto.lEndereco, tAdmMeioPagto.iCodBanco, tAdmMeioPagto.sAgencia, _
                                       iTipoMeioPagto, giFilialEmpresa, ADMMEIOPAGTO_ATIVO)
    If lErro <> AD_SQL_SUCESSO Then gError 107358

    'seleciona o primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 107359

    'se não encontrar nenhum-> erro
    If lErro = AD_SQL_SEM_DADOS Then gError 107360

    'varre a seleção
    Do While lErro <> AD_SQL_SEM_DADOS

        'instancia um novo admmeiopagto
        Set objAdmMeioPagto = New ClassAdmMeioPagto

        'preenche seus atributos
        objAdmMeioPagto.iCodigo = tAdmMeioPagto.iCodigo
        objAdmMeioPagto.iFilialEmpresa = tAdmMeioPagto.iFilialEmpresa
        objAdmMeioPagto.sNome = tAdmMeioPagto.sNome
        objAdmMeioPagto.iRede = tAdmMeioPagto.iRede
        objAdmMeioPagto.dTaxaVista = tAdmMeioPagto.dTaxaVista
        objAdmMeioPagto.dTaxaParcelado = tAdmMeioPagto.dTaxaParcelado
        objAdmMeioPagto.iTipoMeioPagto = tAdmMeioPagto.iTipoMeioPagto
        objAdmMeioPagto.iDefasagemPagtoVista = tAdmMeioPagto.iDefasagemPagtoVista
        objAdmMeioPagto.iContaCorrenteInterna = tAdmMeioPagto.iContaCorrenteInterna
        objAdmMeioPagto.iGeraTituloRec = tAdmMeioPagto.iGeraTituloRec
        objAdmMeioPagto.lEndereco = tAdmMeioPagto.lEndereco
        objAdmMeioPagto.iCodBanco = tAdmMeioPagto.iCodBanco
        objAdmMeioPagto.sAgencia = tAdmMeioPagto.sAgencia

        'adiciona à coleção
        colAdmMeioPagto.Add objAdmMeioPagto

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 107361

    Loop

    'fecha o comando
    Call Comando_Fechar(lComando)

    AdmMeioPagto_Le_TipoMeioPagto = SUCESSO

    Exit Function

Erro_AdmMeioPagto_Le_TipoMeioPagto:

    AdmMeioPagto_Le_TipoMeioPagto = gErr

    Select Case gErr

        Case 107362
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 107358, 107359, 107361
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ADMMEIOPAGTO_TIPOMEIOPAGTO", gErr, iTipoMeioPagto)

        Case 107360
            'será tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151408)

    End Select

    'fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function BorderoValeTicket_Le(ByVal objBorderoValeTicket As ClassBorderoValeTicket) As Long
'Função que recebe um objBorderoValeTicket com filialempresa e numbordero preenchidos e retorna objeto totalmente preenchido

Dim lErro As Long
Dim lComando As Long
Dim tBorderoValeTicket As typeBorderoValeTicket

On Error GoTo Erro_BorderoValeTicket_Le

    'abre o comando
    lComando = Comando_Abrir
    If lComando = 0 Then gError 107367

    'seleciona o bordero
    lErro = Comando_Executar(lComando, "SELECT AdmMeioPagto,Parcelamento,DataEnvio,DataImpressao,DataBackoffice,Valor,NumIntDocCPR " & _
                                       "FROM BorderoValeTicket " & _
                                       "WHERE FilialEmpresa=? AND NumBordero=?", _
                                       tBorderoValeTicket.iAdmMeioPagto, tBorderoValeTicket.iParcelamento, tBorderoValeTicket.dtDataEnvio, tBorderoValeTicket.dtDataImpressao, tBorderoValeTicket.dtDataBackoffice, tBorderoValeTicket.dValor, tBorderoValeTicket.lNumIntDocCPR, _
                                       objBorderoValeTicket.iFilialEmpresa, objBorderoValeTicket.lNumBordero)
    If lErro <> AD_SQL_SUCESSO Then gError 107368

    'busca o primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 107369

    'se nào encontrar-> erro
    If lErro = AD_SQL_SEM_DADOS Then gError 107370

    'preenche os dados do bordero
    objBorderoValeTicket.iAdmMeioPagto = tBorderoValeTicket.iAdmMeioPagto
    objBorderoValeTicket.iParcelamento = tBorderoValeTicket.iParcelamento
    objBorderoValeTicket.dtDataEnvio = tBorderoValeTicket.dtDataEnvio
    objBorderoValeTicket.dtDataImpressao = tBorderoValeTicket.dtDataImpressao
    objBorderoValeTicket.dtDataBackoffice = tBorderoValeTicket.dtDataBackoffice
    objBorderoValeTicket.dValor = tBorderoValeTicket.dValor
    objBorderoValeTicket.lNumIntDocCPR = tBorderoValeTicket.lNumIntDocCPR

    'fecha o comando
    Call Comando_Fechar(lComando)

    BorderoValeTicket_Le = SUCESSO

    Exit Function

Erro_BorderoValeTicket_Le:

    BorderoValeTicket_Le = gErr

    Select Case gErr

        Case 107367
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 107368, 107369
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BORDEROVALETICKET", objBorderoValeTicket.iFilialEmpresa, objBorderoValeTicket.lNumBordero)

        Case 107370
            'será tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151409)

    End Select

    'fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function
'Rafael Menezes( Fim ) - Tela: BorderoValeTicket

'Rafael Menezes( Início ) - Tela: BorderoBoleto
Public Function AdmMeioPagtoCondPagto_Le_Parcelamento(ByVal objAdmMeioPagtoCondPagto As ClassAdmMeioPagtoCondPagto) As Long
'Funçào que seleciona uma admmeiopagtocondpagto específica da tabela admmeiopagtocondpagto

Dim lComando As Long
Dim lErro As Long
Dim tAdmMeioPagtoCondPagto As typeAdmMeioPagtoCondPagto

On Error GoTo Erro_AdmMeioPagtoCondPagto_Le_Parcelamento

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 107294

    lErro = Comando_Executar(lComando, "SELECT AdmMeioPagto,NumParcelas,ParcelasRecebto,Taxa,ValorMinimo,Desconto,JurosParcelamento,Juros,Saldo, Ativo, DataLog, PreDatado " & _
                                       "FROM AdmMeioPagtoCondPagto " & _
                                       "WHERE AdmMeioPagto=? " & _
                                       "AND FilialEmpresa=? " & _
                                       "AND Parcelamento=? ORDER BY AdmMeioPagto", _
                                       tAdmMeioPagtoCondPagto.iAdmMeioPagto, tAdmMeioPagtoCondPagto.iNumParcelas, tAdmMeioPagtoCondPagto.iParcelasRecebto, tAdmMeioPagtoCondPagto.dTaxa, _
                                       tAdmMeioPagtoCondPagto.dValorMinimo, tAdmMeioPagtoCondPagto.dDesconto, tAdmMeioPagtoCondPagto.iJurosParcelamento, tAdmMeioPagtoCondPagto.dJuros, _
                                       tAdmMeioPagtoCondPagto.dSaldo, tAdmMeioPagtoCondPagto.iAtivo, tAdmMeioPagtoCondPagto.dtDataLog, tAdmMeioPagtoCondPagto.iPreDatado, _
                                       objAdmMeioPagtoCondPagto.iAdmMeioPagto, _
                                       objAdmMeioPagtoCondPagto.iFilialEmpresa, _
                                       objAdmMeioPagtoCondPagto.iParcelamento)
    If lErro <> AD_SQL_SUCESSO Then gError 107295

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 107296

    If lErro = AD_SQL_SEM_DADOS Then gError 107297

    objAdmMeioPagtoCondPagto.sNomeParcelamento = tAdmMeioPagtoCondPagto.sNomeParcelamento
    objAdmMeioPagtoCondPagto.iNumParcelas = tAdmMeioPagtoCondPagto.iNumParcelas
    objAdmMeioPagtoCondPagto.iParcelasRecebto = tAdmMeioPagtoCondPagto.iParcelasRecebto
    objAdmMeioPagtoCondPagto.dTaxa = tAdmMeioPagtoCondPagto.dTaxa
    objAdmMeioPagtoCondPagto.dValorMinimo = tAdmMeioPagtoCondPagto.dValorMinimo
    objAdmMeioPagtoCondPagto.dDesconto = tAdmMeioPagtoCondPagto.dDesconto
    objAdmMeioPagtoCondPagto.iJurosParcelamento = tAdmMeioPagtoCondPagto.iJurosParcelamento
    objAdmMeioPagtoCondPagto.dJuros = tAdmMeioPagtoCondPagto.dJuros
    objAdmMeioPagtoCondPagto.dSaldo = tAdmMeioPagtoCondPagto.dSaldo
    objAdmMeioPagtoCondPagto.iAtivo = tAdmMeioPagtoCondPagto.iAtivo
    objAdmMeioPagtoCondPagto.dtDataLog = tAdmMeioPagtoCondPagto.dtDataLog
    objAdmMeioPagtoCondPagto.iPreDatado = tAdmMeioPagtoCondPagto.iPreDatado

    Call Comando_Fechar(lComando)

    AdmMeioPagtoCondPagto_Le_Parcelamento = SUCESSO

    Exit Function

Erro_AdmMeioPagtoCondPagto_Le_Parcelamento:

    AdmMeioPagtoCondPagto_Le_Parcelamento = gErr

    Select Case gErr

        Case 107294
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 107295, 107296
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ADMMEIOPAGTOCONDPAGTO", gErr)

        Case 107297
            'será tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151410)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function BorderoBoletoItem_Le_Bordero(ByVal objBorderoBoleto As ClassBorderoBoleto) As Long
'FUnção que recebe um borderô boleto com filialempresa e número preenchidos e preenche a coleção
'de itens do objborderoboleto

Dim objBorderoBoletoItem As ClassBorderoBoletoItem
Dim lErro As Long
Dim lComando As Long
Dim tBorderoBoletoItem As typeBorderoBoletoItem
Dim iRede As Integer

On Error GoTo Erro_BorderoBoletoItem_Le_Bordero

    'abre o Comando
    lComando = Comando_Abrir
    If lComando = 0 Then gError 107155

    tBorderoBoletoItem.sNomeParcelamento = String(STRING_NOME_PARCELAMENTO, 0)
    tBorderoBoletoItem.sNomeAdmMeioPagto = String(STRING_ADMMEIOPAGTO_NOME, 0)

    'busca os itens de um certo bordero com seus respectivos saldos na tabela admmeiopagtocondpagto
    lErro = Comando_Executar(lComando, "SELECT NomeParcelamento,NumIntDoc,BorderoBoletoItem.AdmMeioPagto,Valor,BorderoBoletoItem.Parcelamento,NumIntDocCPR,Saldo,Nome,AdmMeioPagto.Rede " & _
                                       "FROM AdmMeioPagto,AdmMeioPagtoCondPagto,BorderoBoletoItem " & _
                                       "WHERE BorderoBoletoItem.FilialEmpresa=AdmMeioPagtoCondPagto.FilialEmpresa " & _
                                       "AND BorderoBoletoItem.AdmMeioPagto=AdmMeioPagtoCondPagto.AdmMeioPagto " & _
                                       "AND BorderoBoletoItem.Parcelamento=AdmMeioPagtoCondPagto.Parcelamento " & _
                                       "AND AdmMeioPagto.FilialEmpresa = BorderoBoletoItem.FilialEmpresa " & _
                                       "AND AdmMeioPagto.Codigo=BorderoBoletoItem.AdmMeioPagto " & _
                                       "AND BorderoBoletoItem.FilialEmpresa=? " & _
                                       "AND BorderoBoletoItem.NumBordero=? ORDER BY BorderoBoletoItem.AdmMeioPagto", _
                                       tBorderoBoletoItem.sNomeParcelamento, tBorderoBoletoItem.lNumIntDoc, tBorderoBoletoItem.iAdmMeioPagto, tBorderoBoletoItem.dValor, tBorderoBoletoItem.iParcelamento, tBorderoBoletoItem.lNumIntDocCPR, tBorderoBoletoItem.dSaldo, tBorderoBoletoItem.sNomeAdmMeioPagto, iRede, _
                                       objBorderoBoleto.iFilialEmpresa, _
                                       objBorderoBoleto.lNumBordero)
    If lErro <> AD_SQL_SUCESSO Then gError 107156

    'busca o primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 107157

    'se a seleção estiver vazia -> erro
    If lErro = AD_SQL_SEM_DADOS Then gError 107158

    'enquanto houver dados para leitura
    Do While lErro <> AD_SQL_SEM_DADOS

        'cria um objborderoboletoitem novo
        Set objBorderoBoletoItem = New ClassBorderoBoletoItem

        'o preenche
        objBorderoBoletoItem.iFilialEmpresa = objBorderoBoleto.iFilialEmpresa
        objBorderoBoletoItem.lNumBordero = objBorderoBoleto.lNumBordero
        objBorderoBoletoItem.dSaldo = tBorderoBoletoItem.dSaldo
        objBorderoBoletoItem.dValor = tBorderoBoletoItem.dValor
        objBorderoBoletoItem.iParcelamento = tBorderoBoletoItem.iParcelamento
        objBorderoBoletoItem.lNumIntDoc = tBorderoBoletoItem.lNumIntDoc
        objBorderoBoletoItem.iAdmMeioPagto = tBorderoBoletoItem.iAdmMeioPagto
        objBorderoBoletoItem.lNumIntDocCPR = tBorderoBoletoItem.lNumIntDocCPR
        objBorderoBoletoItem.sNomeParcelamento = tBorderoBoletoItem.sNomeParcelamento
        objBorderoBoletoItem.sNomeAdmMeioPagto = tBorderoBoletoItem.sNomeAdmMeioPagto
        objBorderoBoleto.iCodigoRede = iRede

        'adiciona o novo item à coleção de itens do borderô
        objBorderoBoleto.colBorderoBoletoItem.Add objBorderoBoletoItem

        'busca o próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 107159

    Loop


    'fecha o comando
    Call Comando_Fechar(lComando)

    BorderoBoletoItem_Le_Bordero = SUCESSO

    Exit Function

Erro_BorderoBoletoItem_Le_Bordero:

    BorderoBoletoItem_Le_Bordero = gErr

    Select Case gErr

        Case 107155
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 107156, 107157, 107159
            Call Rotina_Erro(vbOKOnly, "ERRO_BORDEROBOLETOITEM_BORDEROBOLETO_LEITURA", gErr, objBorderoBoleto.iFilialEmpresa, objBorderoBoleto.lNumBordero)

        Case 107158
            'será tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151411)

    End Select

    'fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function BorderoBoleto_Le(ByVal objBorderoBoleto As ClassBorderoBoleto) As Long
'Função que recebe um bordero com filialempresa numbordero preenchidos e retorna o mesmo com os dados preenchidos
'e sua coleção, idem

Dim lErro As Long

On Error GoTo Erro_BorderoBoleto_Le

    'le os dados do bordero somente
    lErro = CF("BorderoBoleto_Le1", objBorderoBoleto)
    If lErro <> SUCESSO And lErro <> 107154 Then gError 107160

    'se não encontrá-lo-> erro
    If lErro = 107154 Then gError 107161

    'lê os itens da coleção do borderô
    lErro = CF("BorderoBoletoItem_Le_Bordero", objBorderoBoleto)
    If lErro <> SUCESSO And lErro <> 107158 Then gError 107162

    'se não encontrar nenhum->erro
    If lErro = 107158 Then gError 107163

    BorderoBoleto_Le = SUCESSO

    Exit Function

Erro_BorderoBoleto_Le:

    BorderoBoleto_Le = gErr

    Select Case gErr

        Case 107160 To 107162

        Case 107163
            Call Rotina_Erro(vbOKOnly, "ERRO_BORDEROBOLETOITEM_BORDEROBOLETO_NAOENCONTRADO", gErr, objBorderoBoleto.iFilialEmpresa, objBorderoBoleto.lNumBordero)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151412)

    End Select

    Exit Function

End Function

Public Function BorderoBoleto_Le1(ByVal objBorderoBoleto As ClassBorderoBoleto) As Long
'função que recebe um borderoboleto preenchido com numbordero e filialempresa e retorna o mesmo com os dados preenchidos

Dim lComando As Long
Dim lErro As Long
Dim tBorderoBoleto As typeBorderoBoleto

On Error GoTo Erro_BorderoBoleto_Le1

    'abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 107151

    'inicializa as strings
    tBorderoBoleto.sNumero = String(STRING_BORDEROBOLETO_NUMERO, 0)

    'seleciona o bordero boleto
    lErro = Comando_Executar(lComando, "SELECT Numero,DataImpressao,DataEnvio,DataBackoffice,ExibeTela " & _
                                       "FROM BorderoBoleto " & _
                                       "WHERE FilialEmpresa=? " & _
                                       "AND NumBordero=? ORDER BY Numero", _
                                       tBorderoBoleto.sNumero, tBorderoBoleto.dtDataImpressao, tBorderoBoleto.dtDataEnvio, tBorderoBoleto.dtDataBackoffice, tBorderoBoleto.iExibeTela, _
                                       objBorderoBoleto.iFilialEmpresa, _
                                       objBorderoBoleto.lNumBordero)
    If lErro <> AD_SQL_SUCESSO Then gError 107152

    'busca o primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then gError 107153

    'se não encontrou-> erro
    If lErro = AD_SQL_SEM_DADOS Then gError 107154

    'preenche os dados do borderoboleto
    objBorderoBoleto.dtDataBackoffice = tBorderoBoleto.dtDataBackoffice
    objBorderoBoleto.dtDataEnvio = tBorderoBoleto.dtDataEnvio
    objBorderoBoleto.dtDataImpressao = tBorderoBoleto.dtDataImpressao
    objBorderoBoleto.iExibeTela = tBorderoBoleto.iExibeTela
    objBorderoBoleto.sNumero = tBorderoBoleto.sNumero

    'fecha o comando
    Call Comando_Fechar(lComando)

    BorderoBoleto_Le1 = SUCESSO

    Exit Function

Erro_BorderoBoleto_Le1:

    BorderoBoleto_Le1 = gErr

    Select Case gErr

        Case 107151
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 107152, 107153
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BORDEROBOLETO", gErr, objBorderoBoleto.iFilialEmpresa, objBorderoBoleto.lNumBordero)

        Case 107154

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151413)

    End Select

    'fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function AdmMeioPagtoCondPagto_Le1(iRede As Integer, ByVal objAdmMeioPagtoCondPagto As ClassAdmMeioPagtoCondPagto, ByVal colAdmMeioPagtoCondPagto As Collection) As Long
'Função que abre o comando e chama as demais para a montagem e execução do comando SQL

Dim lErro As Long
Dim lComando As Long
Dim sSQL As String
Dim iAdmMeioPagto As Integer
Dim sNomeAdmMeioPagto As String
Dim iParcelamento As Integer
Dim sNomeParcelamento As String
Dim dSaldo As Double
Dim sSelect As String

On Error GoTo Erro_AdmMeioPagtoCondPagto_Le1

    'abre o comando
    lComando = Comando_Abrir
    If lComando = 0 Then gError 107220

    'monta o esqueleto do select
    Call AdmMeioPagtoCondPagto_Le2(objAdmMeioPagtoCondPagto, sSelect)

    ' inicializar as strings
    sNomeAdmMeioPagto = String(STRING_ADMMEIOPAGTO_NOME, 0)
    sNomeParcelamento = String(STRING_NOME_PARCELAMENTO, 0)

    'monta a parte do select que contém os campos de retorno
    lErro = AdmMeioPagtoCondPagto_Le3(lComando, sSelect, iAdmMeioPagto, sNomeAdmMeioPagto, iParcelamento, sNomeParcelamento, dSaldo)
    If lErro <> SUCESSO Then gError 107222

    'monta a parte do select que contém os campos de filtro
    lErro = AdmMeioPagtoCondPagto_Le4(lComando, objAdmMeioPagtoCondPagto, iRede)
    If lErro <> SUCESSO Then gError 107223

    'executa o select
    lErro = AdmMeioPagtoCondPagto_Le5(lComando, colAdmMeioPagtoCondPagto, iAdmMeioPagto, sNomeAdmMeioPagto, iParcelamento, sNomeParcelamento, dSaldo)
    If lErro <> SUCESSO And lErro <> 107218 Then gError 107224

    If lErro = 107218 Then gError 107225

    Call Comando_Fechar(lComando)

    AdmMeioPagtoCondPagto_Le1 = SUCESSO

    Exit Function

Erro_AdmMeioPagtoCondPagto_Le1:

    AdmMeioPagtoCondPagto_Le1 = gErr

    Select Case gErr
    
        Case 107220
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 107222 To 107224

        Case 107225
            'será tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151414)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Private Sub AdmMeioPagtoCondPagto_Le2(ByVal objAdmMeioPagtoCondPagto As ClassAdmMeioPagtoCondPagto, sSelect As String)
'Função que monta o comando SQL

Dim sFields As String
Dim sFrom As String
Dim sWhere As String
Dim sOrderBy As String

On Error GoTo Erro_AdmMeioPagtoCondPagto_Le2

    'select fixo
    sSelect = "SELECT AdmMeioPagtoCondPagto.AdmMeioPagto, AdmMeiopagto.Nome,AdmMeioPagtoCondPagto.Parcelamento,AdmMeioPagtoCondPagto.NomeParcelamento,AdmMeioPagtoCondPagto.Saldo"

    'cláusula FROM
    sFrom = " FROM AdmMeioPagtoCondPagto,AdmMeioPagto"

    'cláusula WHERE fixa
    sWhere = " WHERE AdmMeioPagtoCondPagto.AdmMeioPagto=AdmMeioPagto.Codigo" & _
             " AND AdmMeioPagtoCondPagto.FilialEmpresa=AdmMeioPagto.FilialEmpresa" & _
             " AND AdmMeioPagto.Rede=?" & _
             " AND AdmMeioPagto.FilialEmpresa=?" & _
             " AND AdmMeioPagtoCondPagto.Saldo>0"

    'se foi especificado um cartão, incluir na cláusula WHERE
    If objAdmMeioPagtoCondPagto.iAdmMeioPagto <> 0 Then sWhere = sWhere & " AND AdmMeioPagtoCondPagto.AdmMeioPagto=?"

    'verifica qual o tipo de parcelamento
    Select Case objAdmMeioPagtoCondPagto.iParcelamento

        'se for à vista
        Case PARCELAMENTO_AVISTA
            '...AND Parcelamento = PARCELAMENTO_AVISTA
            sWhere = sWhere & " AND (Parcelamento=?)"

        'se for parcelado
        Case PARCELAMENTO_PARCELADO

            'verifica qual o tipo de juros
            Select Case objAdmMeioPagtoCondPagto.iJurosParcelamento

                'no caso de juros da loja ou(exclusivo) da administradora
                Case JUROS_LOJA, JUROS_ADM
                    '...AND Parcelamento > PARCELAMENTO_AVISTA AND JurosParcelamento = iJurosParcelamento-1
                    sWhere = sWhere & " AND Parcelamento>? AND JurosParcelamento=?"

                'no caso de ambas
                Case JUROS_LOJA + JUROS_ADM
                    '...AND Parcelamento > PARCELAMENTO_AVISTA
                    sWhere = sWhere & " AND Parcelamento>?"

            End Select

        'se for parcelado ou à vista
        Case PARCELAMENTO_AMBOS

            'verifica qual o tipo de juros
            Select Case objAdmMeioPagtoCondPagto.iJurosParcelamento

                'se estiver marcado um dos dois...
                Case JUROS_LOJA, JUROS_ADM

                    'haverá dois jurosparcelamento. um valendo 0 para o caso à vista e outro com o tipo específico de parcelamento
                    '...AND(JurosParcelamento=0 OR JurosParcelamento=iJurosParcelamento)
                    sWhere = sWhere & " AND(JurosParcelamento=? OR JurosParcelamento=?)"

                'se ambos estiverem selecionados, pode ser qq valor em jurosparcelamento e parcelamento
                Case JUROS_LOJA + JUROS_ADM
                    'sem cláusulas where

            End Select

    End Select

    'cláusula order by
    sOrderBy = " ORDER BY AdmMeioPagto.Nome, AdmMeioPagtoCondPagto.NomeParcelamento"

    'monta o comando sql completo
    sSelect = sSelect & sFrom & sWhere & sOrderBy
    
    Exit Sub

Erro_AdmMeioPagtoCondPagto_Le2:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151415)

    End Select

    Exit Sub

End Sub

Private Function AdmMeioPagtoCondPagto_Le3(ByVal lComando As Long, ByVal sSelect As String, viAdmMeioPagto As Variant, vsNomeAdmMeioPagto As Variant, viParcelamento As Variant, vsNomeParcelamento As Variant, vdSaldo As Variant) As Long
'Função que faz o bind das colunas do comando SQL

Dim lErro As Long

On Error GoTo Erro_AdmMeioPagtoCondPagto_Le3

    'prepara o comando sql
    lErro = Comando_PrepararInt(lComando, sSelect)
    If lErro <> AD_SQL_SUCESSO Then gError 107206

    'faz o bind das colunas
    lErro = Comando_BindVarInt(lComando, viAdmMeioPagto)
    If lErro <> AD_SQL_SUCESSO Then gError 107207

    lErro = Comando_BindVarInt(lComando, vsNomeAdmMeioPagto)
    If lErro <> AD_SQL_SUCESSO Then gError 107208

    lErro = Comando_BindVarInt(lComando, viParcelamento)
    If lErro <> AD_SQL_SUCESSO Then gError 107209

    lErro = Comando_BindVarInt(lComando, vsNomeParcelamento)
    If lErro <> AD_SQL_SUCESSO Then gError 107210

    lErro = Comando_BindVarInt(lComando, vdSaldo)
    If lErro <> AD_SQL_SUCESSO Then gError 107211

    AdmMeioPagtoCondPagto_Le3 = SUCESSO

    Exit Function

Erro_AdmMeioPagtoCondPagto_Le3:

AdmMeioPagtoCondPagto_Le3 = gErr

    Select Case gErr

        Case 107206 To 107211
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ADMMEIOPAGTOCONDPAGTO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151416)

    End Select

    Exit Function

End Function

Private Function AdmMeioPagtoCondPagto_Le4(ByVal lComando As Long, ByVal objAdmMeioPagtoCondPagto As ClassAdmMeioPagtoCondPagto, ByVal iRede As Integer) As Long
'Função que faz o bind do filtro do SQL

Dim lErro As Long
Dim viAdmMeioPagto As Variant
Dim viParcelamento As Variant
Dim viJurosParcelamento(0 To 1) As Variant
Dim viRede As Variant
Dim viFilialEmpresa As Variant

On Error GoTo Erro_AdmMeioPagtoCondPagto_Le4

    'faz o bind dos filtros
    viRede = iRede
    lErro = Comando_BindVarInt(lComando, viRede)
    If lErro <> AD_SQL_SUCESSO Then gError 107212

    viFilialEmpresa = giFilialEmpresa
    lErro = Comando_BindVarInt(lComando, viFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 107211

    If objAdmMeioPagtoCondPagto.iAdmMeioPagto <> 0 Then

        viAdmMeioPagto = objAdmMeioPagtoCondPagto.iAdmMeioPagto
        lErro = Comando_BindVarInt(lComando, viAdmMeioPagto)
        If lErro <> AD_SQL_SUCESSO Then gError 107213

    End If

    Select Case objAdmMeioPagtoCondPagto.iParcelamento

        Case PARCELAMENTO_AVISTA
            viParcelamento = PARCELAMENTO_AVISTA
            lErro = Comando_BindVarInt(lComando, viParcelamento)
            If lErro <> AD_SQL_SUCESSO Then gError 107233

        Case PARCELAMENTO_PARCELADO

            Select Case objAdmMeioPagtoCondPagto.iJurosParcelamento

                Case JUROS_LOJA, JUROS_ADM
                    viParcelamento = PARCELAMENTO_AVISTA
                    lErro = Comando_BindVarInt(lComando, viParcelamento)
                    If lErro <> AD_SQL_SUCESSO Then gError 107234

                    viJurosParcelamento(0) = objAdmMeioPagtoCondPagto.iJurosParcelamento
                    lErro = Comando_BindVarInt(lComando, viJurosParcelamento(0))
                    If lErro <> AD_SQL_SUCESSO Then gError 107235

                Case JUROS_LOJA + JUROS_ADM
                    viParcelamento = PARCELAMENTO_AVISTA
                    lErro = Comando_BindVarInt(lComando, viParcelamento)
                    If lErro <> AD_SQL_SUCESSO Then gError 107236

            End Select

        Case PARCELAMENTO_AMBOS

            Select Case objAdmMeioPagtoCondPagto.iJurosParcelamento

                Case JUROS_LOJA, JUROS_ADM
                    viJurosParcelamento(0) = JUROS_LOJA
                    lErro = Comando_BindVarInt(lComando, viJurosParcelamento(0))
                    If lErro <> AD_SQL_SUCESSO Then gError 107237

                    viJurosParcelamento(1) = objAdmMeioPagtoCondPagto.iJurosParcelamento
                    lErro = Comando_BindVarInt(lComando, viJurosParcelamento(1))
                    If lErro <> AD_SQL_SUCESSO Then gError 107238

                Case JUROS_LOJA + JUROS_ADM
                    'sem cláusulas where

            End Select

    End Select

    lErro = Comando_ExecutarInt(lComando)
    If lErro <> AD_SQL_SUCESSO Then gError 107216

    AdmMeioPagtoCondPagto_Le4 = SUCESSO

    Exit Function

Erro_AdmMeioPagtoCondPagto_Le4:

    AdmMeioPagtoCondPagto_Le4 = gErr

    Select Case gErr

        Case 107212 To 107213, 107216, 107233 To 107238
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ADMMEIOPAGTOCONDPAGTO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151417)

    End Select

    Exit Function

End Function

Private Function AdmMeioPagtoCondPagto_Le5(ByVal lComando As Long, ByVal colAdmMeioPagtoCondPagto As Collection, iAdmMeioPagto As Integer, sNomeAdmMeioPagto As String, iParcelamento As Integer, sNomeParcelamento As String, dSaldo As Double) As Long
'Função que executa o comando SQL e preenche a coleção de admmeiopagtocondpagto

Dim objAdmMeioPagtoCondPagto As ClassAdmMeioPagtoCondPagto
Dim lErro As Long

On Error GoTo Erro_AdmMeioPagtoCondPagto_Le5

    'busca o primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 107217

    'se não encontrar-> erro
    If lErro = AD_SQL_SEM_DADOS Then gError 107218

    'varre a seleção feita
    Do While lErro <> AD_SQL_SEM_DADOS

        'instancia uma admmeiopagtocondpagto
        Set objAdmMeioPagtoCondPagto = New ClassAdmMeioPagtoCondPagto

        'preenche seus atributos
        objAdmMeioPagtoCondPagto.iAdmMeioPagto = iAdmMeioPagto
        objAdmMeioPagtoCondPagto.sNomeAdmMeioPagto = sNomeAdmMeioPagto
        objAdmMeioPagtoCondPagto.iParcelamento = iParcelamento
        objAdmMeioPagtoCondPagto.sNomeParcelamento = sNomeParcelamento
        objAdmMeioPagtoCondPagto.dSaldo = dSaldo

        'adiciona à coleção
        colAdmMeioPagtoCondPagto.Add objAdmMeioPagtoCondPagto

        'busca o próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then gError 107219

    Loop

    AdmMeioPagtoCondPagto_Le5 = SUCESSO

    Exit Function

Erro_AdmMeioPagtoCondPagto_Le5:

    AdmMeioPagtoCondPagto_Le5 = gErr

    Select Case gErr

        Case 107217, 107219
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ADMMEIOPAGTOCONDPAGTO", gErr)

        Case 107218
            'será tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151418)

    End Select

    Exit Function

End Function
'Rafael Menezes( Fim ) - Tela: BorderoBoleto

Public Function Cheque_Le1(colCheque As Collection, colChequeN As Collection, dtDataInicio As Date, dtDataFim As Date) As Long
'Função que recebe uma colecao de cheques especificados, uma colecao de cheques não especificados e um intervalo de datas para selecioná-los
'na tabela de cheque-pre, desde que eles não pertençam a nenhum bordero.
 
Dim lErro As Long
Dim lComando As Long
Dim sSQL As String
Dim tChequePre As typeChequePre
Dim objChequePre As ClassChequePre
 
On Error GoTo Erro_Cheque_Le1
 
    'abre o comando
    lComando = Comando_Abrir
    If lComando = 0 Then gError 107004
 
    'inicializa as strings
    tChequePre.sAgencia = String(STRING_AGENCIA, 0)
    tChequePre.sContaCorrente = String(STRING_CONTA_CORRENTE, 0)
    tChequePre.sCPFCGC = String(STRING_CGC, 0)
 
    'se a data de inicio e data de fim forem diferentes de zero
    If dtDataInicio <> 0 And dtDataFim <> 0 Then
 
        'monta o comando sql
        sSQL = "SELECT NumIntCheque, Cliente, Filial, Banco, Agencia, ContaCorrente, Numero, DataDeposito, Valor, NumBordero, CPFCGC, NumMovtoCaixa, Aprovado, NaoEspecificado, NumBorderoLoja, FilialEmpresa, SequencialLoja, SequencialBack, FilialEmpresaLoja, DataEmissao FROM ChequePre WHERE DataDeposito>= ? AND DataDeposito <= ? AND Localizacao=? AND Status <> ?"
 
        'selecionar os cheques da filiempresaloja que não pertencem a nenhum borderô e com data de deposito compreendida no intervalo recebido por parâmetro
        lErro = Comando_Executar(lComando, sSQL, tChequePre.lNumIntCheque, tChequePre.lCliente, tChequePre.iFilial, tChequePre.iBanco, tChequePre.sAgencia, tChequePre.sContaCorrente, tChequePre.lNumero, tChequePre.dtDataDeposito, tChequePre.dValor, tChequePre.lNumBordero, tChequePre.sCPFCGC, tChequePre.lNumMovtoCaixa, tChequePre.iAprovado, tChequePre.iNaoEspecificado, tChequePre.lNumBorderoLoja, tChequePre.iFilialEmpresa, tChequePre.lSequencialLoja, tChequePre.lSequencialBack, tChequePre.iFilialEmpresaLoja, tChequePre.dtDataEmissao, dtDataInicio, dtDataFim, CHEQUEPRE_LOCALIZACAO_LOJA, STATUS_EXCLUIDO)
        If lErro <> AD_SQL_SUCESSO Then gError 107005
 
    'se menos do que as duas datas estiver preenchida
    Else
 
        'se somente a data de início estiver preenchida
        If dtDataInicio <> 0 Then
 
            'monta o comando sql
            sSQL = "SELECT NumIntCheque, Cliente, Filial, Banco, Agencia, ContaCorrente, Numero, DataDeposito, Valor, NumBordero, CPFCGC, NumMovtoCaixa, Aprovado, NaoEspecificado, NumBorderoLoja, FilialEmpresa, SequencialLoja, SequencialBack, FilialEmpresaLoja, DataEmissao FROM ChequePre WHERE DataDeposito>= ? AND Localizacao=? AND Status <> ?"
 
            'selecionar os cheques da filiempresaloja que não pertencem a nenhum borderô e com data de deposito compreendida maior que a data de inicio fornecida
            lErro = Comando_Executar(lComando, sSQL, tChequePre.lNumIntCheque, tChequePre.lCliente, tChequePre.iFilial, tChequePre.iBanco, tChequePre.sAgencia, tChequePre.sContaCorrente, tChequePre.lNumero, tChequePre.dtDataDeposito, tChequePre.dValor, tChequePre.lNumBordero, tChequePre.sCPFCGC, tChequePre.lNumMovtoCaixa, tChequePre.iAprovado, tChequePre.iNaoEspecificado, tChequePre.lNumBorderoLoja, tChequePre.iFilialEmpresa, tChequePre.lSequencialLoja, tChequePre.lSequencialBack, tChequePre.iFilialEmpresaLoja, tChequePre.dtDataEmissao, dtDataInicio, CHEQUEPRE_LOCALIZACAO_LOJA, STATUS_EXCLUIDO)
            If lErro <> AD_SQL_SUCESSO Then gError 107006
 
        Else
 
            'se somente a data de fim estiver preenchida
            If dtDataFim <> 0 Then
 
                'monta o comando sql
                sSQL = "SELECT NumIntCheque, Cliente, Filial, Banco, Agencia, ContaCorrente, Numero, DataDeposito, Valor, NumBordero, CPFCGC, NumMovtoCaixa, Aprovado, NaoEspecificado, NumBorderoLoja, FilialEmpresa, SequencialLoja, SequencialBack, FilialEmpresaLoja, DataEmissao FROM ChequePre WHERE DataDeposito<= ? AND Localizacao=? AND Status <> ?"
 
                'selecionar os cheques da filiempresaloja que não pertencem a nenhum borderô e com data de deposito compreendida menor que a data de fim fornecida
                lErro = Comando_Executar(lComando, sSQL, tChequePre.lNumIntCheque, tChequePre.lCliente, tChequePre.iFilial, tChequePre.iBanco, tChequePre.sAgencia, tChequePre.sContaCorrente, tChequePre.lNumero, tChequePre.dtDataDeposito, tChequePre.dValor, tChequePre.lNumBordero, tChequePre.sCPFCGC, tChequePre.lNumMovtoCaixa, tChequePre.iAprovado, tChequePre.iNaoEspecificado, tChequePre.lNumBorderoLoja, tChequePre.iFilialEmpresa, tChequePre.lSequencialLoja, tChequePre.lSequencialBack, tChequePre.iFilialEmpresaLoja, tChequePre.dtDataEmissao, dtDataFim, CHEQUEPRE_LOCALIZACAO_LOJA, STATUS_EXCLUIDO)
                If lErro <> AD_SQL_SUCESSO Then gError 107007
 
            'se nenhuma das datas estiver preenchida
            Else
 
                'monta o comando sql
                sSQL = "SELECT NumIntCheque, Cliente, Filial, Banco, Agencia, ContaCorrente, Numero, DataDeposito, Valor, NumBordero, CPFCGC, NumMovtoCaixa, Aprovado, NaoEspecificado, NumBorderoLoja, FilialEmpresa, SequencialLoja, SequencialBack, FilialEmpresaLoja, DataEmissao FROM ChequePre WHERE Localizacao=? AND Status <> ?"
 
                'selecionar os cheques da filiempresaloja que não pertencem a nenhum borderô
                lErro = Comando_Executar(lComando, sSQL, tChequePre.lNumIntCheque, tChequePre.lCliente, tChequePre.iFilial, tChequePre.iBanco, tChequePre.sAgencia, tChequePre.sContaCorrente, tChequePre.lNumero, tChequePre.dtDataDeposito, tChequePre.dValor, tChequePre.lNumBordero, tChequePre.sCPFCGC, tChequePre.lNumMovtoCaixa, tChequePre.iAprovado, tChequePre.iNaoEspecificado, tChequePre.lNumBorderoLoja, tChequePre.iFilialEmpresa, tChequePre.lSequencialLoja, tChequePre.lSequencialBack, tChequePre.iFilialEmpresaLoja, tChequePre.dtDataEmissao, CHEQUEPRE_LOCALIZACAO_LOJA, STATUS_EXCLUIDO)
                If lErro <> AD_SQL_SUCESSO Then gError 107008
 
            End If
 
        End If
 
    End If
 
    'busca o primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 107009
 
    'se a seleção retornar vazia-> erro
    If lErro = AD_SQL_SEM_DADOS Then gError 107010
 
    'enquanto houver dados na seleção feita
    Do While lErro <> AD_SQL_SEM_DADOS
 
        'cria um novo objeto
        Set objChequePre = New ClassChequePre
 
        'o preenche
        With objChequePre
 
            .lNumIntCheque = tChequePre.lNumIntCheque
            .lCliente = tChequePre.lCliente
            .iFilial = tChequePre.iFilial
            .iBanco = tChequePre.iBanco
            .sAgencia = tChequePre.sAgencia
            .sContaCorrente = tChequePre.sContaCorrente
            .lNumero = tChequePre.lNumero
            .dtDataDeposito = tChequePre.dtDataDeposito
            .dValor = tChequePre.dValor
            .lNumBordero = tChequePre.lNumBordero
            .sCPFCGC = tChequePre.sCPFCGC
            .lNumMovtoCaixa = tChequePre.lNumMovtoCaixa
            .iAprovado = tChequePre.iAprovado
            .iNaoEspecificado = tChequePre.iNaoEspecificado
            .lNumBorderoLoja = tChequePre.lNumBorderoLoja
            .iFilialEmpresa = tChequePre.iFilialEmpresa
            .lSequencialLoja = tChequePre.lSequencialLoja
            .lSequencialBack = tChequePre.lSequencialBack
            .iFilialEmpresaLoja = tChequePre.iFilialEmpresaLoja
            .dtDataEmissao = tChequePre.dtDataEmissao
            
 
        End With
 
        'se for especificado o coloca na coleção de cheques especificados
        If objChequePre.iNaoEspecificado = CHEQUE_ESPECIFICADO Then colCheque.Add objChequePre
 
        'se for especificado o coloca na coleção de cheques não especificados
        If objChequePre.iNaoEspecificado = CHEQUE_NAO_ESPECIFICADO Then colChequeN.Add objChequePre
 
        'busca o próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 107011
 
    Loop
 
    'fecha os comandos
    Call Comando_Fechar(lComando)
 
    Cheque_Le1 = SUCESSO
 
    Exit Function
 
Erro_Cheque_Le1:
 
    Cheque_Le1 = gErr
 
    Select Case gErr
 
        Case 107004
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
 
        Case 107005 To 107009, 107011
            Call Rotina_Erro(vbOKOnly, "ERRO_CHEQUEPRE_LEITURA_INTERVALO", gErr)
 
        Case 107010
            'será tratado na rotina chamadora
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 151419)
 
    End Select
 
    Call Comando_Fechar(lComando)
 
    Exit Function
 
End Function

Public Function Cheque_Le_Bordero(ByVal objBorderoCheque As ClassBorderoCheque) As Long
'Função que recebe um código de borderô, uma coleção de cheques detalhados e uma de não detalhados e retorna ambas preenchidas

Dim objChequePre As ClassChequePre
Dim lErro As Long
Dim lComando As Long
Dim tChequePre As typeChequePre

On Error GoTo Erro_Cheque_Le_Bordero

    'abre o comando
    lComando = Comando_Abrir
    If lComando = 0 Then gError 103960

    'inicializa as strings
    tChequePre.sAgencia = String(STRING_AGENCIA, 0)
    tChequePre.sContaCorrente = String(STRING_CONTA_CORRENTE, 0)
    tChequePre.sCPFCGC = String(STRING_CGC, 0)

    If objBorderoCheque.iCodNossaConta <> 0 Then

        'tenta selecionar todos os cheques referentes ao bordero
        lErro = Comando_Executar(lComando, "SELECT NumIntCheque,Cliente,Filial,Banco,Agencia,ContaCorrente,Numero,DataDeposito,Valor,NumBordero,CPFCGC,NumMovtoCaixa,Aprovado,NaoEspecificado,FilialEmpresa,SequencialLoja,SequencialBack,FilialEmpresaLoja, DataEmissao FROM ChequePre WHERE NumBorderoLojaBanco=?", _
        tChequePre.lNumIntCheque, tChequePre.lCliente, tChequePre.iFilial, tChequePre.iBanco, tChequePre.sAgencia, tChequePre.sContaCorrente, tChequePre.lNumero, tChequePre.dtDataDeposito, tChequePre.dValor, tChequePre.lNumBordero, tChequePre.sCPFCGC, tChequePre.lNumMovtoCaixa, tChequePre.iAprovado, tChequePre.iNaoEspecificado, tChequePre.iFilialEmpresa, tChequePre.lSequencialLoja, tChequePre.lSequencialBack, tChequePre.iFilialEmpresaLoja, tChequePre.dtDataEmissao, objBorderoCheque.lNumBordero)
        If lErro <> AD_SQL_SUCESSO Then gError 103961

    Else

        'tenta selecionar todos os cheques referentes ao bordero
        lErro = Comando_Executar(lComando, "SELECT NumIntCheque,Cliente,Filial,Banco,Agencia,ContaCorrente,Numero,DataDeposito,Valor,NumBordero,CPFCGC,NumMovtoCaixa,Aprovado,NaoEspecificado,FilialEmpresa,SequencialLoja,SequencialBack,FilialEmpresaLoja, DataEmissao FROM ChequePre WHERE NumBorderoLoja=?", _
        tChequePre.lNumIntCheque, tChequePre.lCliente, tChequePre.iFilial, tChequePre.iBanco, tChequePre.sAgencia, tChequePre.sContaCorrente, tChequePre.lNumero, tChequePre.dtDataDeposito, tChequePre.dValor, tChequePre.lNumBordero, tChequePre.sCPFCGC, tChequePre.lNumMovtoCaixa, tChequePre.iAprovado, tChequePre.iNaoEspecificado, tChequePre.iFilialEmpresa, tChequePre.lSequencialLoja, tChequePre.lSequencialBack, tChequePre.iFilialEmpresaLoja, tChequePre.dtDataEmissao, objBorderoCheque.lNumBordero)
        If lErro <> AD_SQL_SUCESSO Then gError 103961

    End If

    'Busca o primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 103962

    'se não encontrar->erro
    If lErro = AD_SQL_SEM_DADOS Then gError 103963

    Do While lErro <> AD_SQL_SEM_DADOS

        'cria um novo objchequepre
        Set objChequePre = New ClassChequePre

        'e o preenche
        With objChequePre

            .lNumIntCheque = tChequePre.lNumIntCheque
            .lCliente = tChequePre.lCliente
            .iFilial = tChequePre.iFilial
            .iBanco = tChequePre.iBanco
            .sAgencia = tChequePre.sAgencia
            .sContaCorrente = tChequePre.sContaCorrente
            .lNumero = tChequePre.lNumero
            .dtDataDeposito = tChequePre.dtDataDeposito
            .dValor = tChequePre.dValor
            .lNumBordero = tChequePre.lNumBordero
            .sCPFCGC = tChequePre.sCPFCGC
            .lNumMovtoCaixa = tChequePre.lNumMovtoCaixa
            .iAprovado = tChequePre.iAprovado
            .iNaoEspecificado = tChequePre.iNaoEspecificado
            .iFilialEmpresa = tChequePre.iFilialEmpresa
            .lSequencialLoja = tChequePre.lSequencialLoja
            .lSequencialBack = tChequePre.lSequencialBack
            .iFilialEmpresaLoja = tChequePre.iFilialEmpresaLoja
            If objBorderoCheque.iCodNossaConta <> 0 Then
                .lNumBorderoLojaBanco = objBorderoCheque.lNumBordero
            Else
                .lNumBorderoLoja = objBorderoCheque.lNumBordero
            End If
            .dtDataEmissao = tChequePre.dtDataEmissao
            
        End With

        'se for especificado o coloca na colecao de cheques especificados
        If objChequePre.iNaoEspecificado = CHEQUE_ESPECIFICADO Then objBorderoCheque.colCheque.Add objChequePre

        'se for não especificado o coloca na colecao de cheques não especificados
        If objChequePre.iNaoEspecificado = CHEQUE_NAO_ESPECIFICADO Then objBorderoCheque.colChequeN.Add objChequePre

        'busca o próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 103964

    Loop

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Cheque_Le_Bordero = SUCESSO

    Exit Function

Erro_Cheque_Le_Bordero:

    Cheque_Le_Bordero = gErr

    Select Case gErr

        Case 103960
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 103961, 103962, 103964
            Call Rotina_Erro(vbOKOnly, "ERRO_CHEQUEPRE_BORDEROCHEQUE_LEITURA", gErr, giFilialEmpresa, objBorderoCheque.lNumBordero)

        Case 103963
            'será tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151420)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function BorderoCheque_Le1(ByVal objBorderoCheque As ClassBorderoCheque) As Long
'Função que recebe um objBorderoCheque com os campos filialempresa e codigo preenchidos
'e o retorna preenchido sem a coleção de cheques, caso encontre

Dim lErro As Long
Dim lComando As Long
Dim tBorderoCheque As typeBorderoCheque

On Error GoTo Erro_BorderoCheque_Le1

    'abre o comando
    lComando = Comando_Abrir
    If lComando = 0 Then gError 103956

    'tenta selecionar o bordero recebido por parâmetro
    lErro = Comando_Executar(lComando, "SELECT CodNossaConta,DataImpressao,DataEnvio,DataBackoffice FROM BorderoCheque WHERE FilialEmpresa=? AND NumBordero=?", _
    tBorderoCheque.iCodNossaConta, tBorderoCheque.dtDataImpressao, tBorderoCheque.dtDataEnvio, tBorderoCheque.dtDataBackoffice, objBorderoCheque.iFilialEmpresa, objBorderoCheque.lNumBordero)
    If lErro <> AD_SQL_SUCESSO Then gError 103957

    'busca o primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then gError 103958

    'se não encontrar->erro
    If lErro = AD_SQL_SEM_DADOS Then gError 103959

    'preenche os dados do borderô
    With objBorderoCheque

        .dtDataBackoffice = tBorderoCheque.dtDataBackoffice
        .dtDataEnvio = tBorderoCheque.dtDataEnvio
        .dtDataImpressao = tBorderoCheque.dtDataImpressao
        .iCodNossaConta = tBorderoCheque.iCodNossaConta

    End With

    'fecha o comando
    Call Comando_Fechar(lComando)

    BorderoCheque_Le1 = SUCESSO

    Exit Function

Erro_BorderoCheque_Le1:

    BorderoCheque_Le1 = gErr

    Select Case gErr

        Case 103956
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 103957, 103958
            Call Rotina_Erro(vbOKOnly, "ERRO_BORDEROCHEQUE_LEITURA", gErr, objBorderoCheque.iFilialEmpresa, objBorderoCheque.lNumBordero)

        Case 103959
        'será tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151421)

    End Select

    'fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function
'Rafael Menezes( Fim ) - Tela: BorderoCheque

Public Function TipoMeioPagtoLojaFilial_Le(ByVal objTMPLojaFilial As ClassTMPLojaFilial) As Long

Dim lComando As Long
Dim lErro As Long
Dim tTMPLojaFilial As typeTMPLojaFilial
Dim bAbriuTransacao As Boolean
Dim lTransacao As Long

On Error GoTo Erro_TipoMeioPagtoLojaFilial_Le

    lComando = Comando_Abrir
    If lComando = 0 Then gError 107043
    
    bAbriuTransacao = False
    
    lErro = Comando_Executar(lComando, "SELECT Saldo FROM TipoMeioPagtoLojaFilial WHERE FilialEmpresa=? AND Tipo=?", tTMPLojaFilial.dSaldo, objTMPLojaFilial.iFilialEmpresa, objTMPLojaFilial.iTipo)
    If lErro <> AD_SQL_SUCESSO Then gError 107044
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 107045
    
    If lErro = AD_SQL_SEM_DADOS Then
    
        If GL_lTransacao = 0 Then
            bAbriuTransacao = True
            'abre a transacao
            lTransacao = Transacao_Abrir
            If lTransacao = 0 Then gError 103907
        End If
    
        lErro = Comando_Executar(lComando, "INSERT INTO TipoMeioPagtoLojaFilial (FilialEmpresa, Tipo, Saldo) VALUES (?, ?, ?)", objTMPLojaFilial.iFilialEmpresa, objTMPLojaFilial.iTipo, 0)
        If lErro <> AD_SQL_SUCESSO Then gError 107046
        
        If bAbriuTransacao Then
            lErro = Transacao_Commit()
            If lErro <> SUCESSO Then gError 103925
        End If
        
        tTMPLojaFilial.dSaldo = 0
    End If
    
    objTMPLojaFilial.dSaldo = tTMPLojaFilial.dSaldo

    TipoMeioPagtoLojaFilial_Le = SUCESSO
    
    Call Comando_Fechar(lComando)

    Exit Function

Erro_TipoMeioPagtoLojaFilial_Le:

    TipoMeioPagtoLojaFilial_Le = gErr

    Select Case gErr
        
        Case 103907
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 103925
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case 107043
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 107044, 107045
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOMEIOPAGTOLOJAFILIAL", gErr, objTMPLojaFilial.iFilialEmpresa, objTMPLojaFilial.iTipo)
        
        Case 107046
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TIPOMEIOPAGTOLOJAFILIAL", gErr, objTMPLojaFilial.iTipo, objTMPLojaFilial.iFilialEmpresa)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151422)
    
    End Select
    
    If bAbriuTransacao Then Call Transacao_Rollback

    Call Comando_Fechar(lComando)

    Exit Function

End Function

'Rafael Menezes( Início ) - Tela: SaqueCaixa
Public Function ContasCorrentesInternas_Le_Todas(ByVal colContasCorrentesInternas As Collection) As Long
'Função que recebe uma coleção de contas correntes internas vazia e retorna com ela preenchida com as contas correntes
'da tabela de contas correntes internas ou erro caso a tabela esteja vazia.

Dim lComando As Long
Dim lErro As Long
Dim tContaCorrenteInterna As typeContaCorrenteInt
Dim objContaCorrenteInterna As ClassContasCorrentesInternas

On Error GoTo Erro_ContasCorrentesInternas_Le_Todas

    'abre comando
    lComando = Comando_Abrir
    If lComando = 0 Then gError 103862
    
    'inicializa as strings
    tContaCorrenteInterna.sAgencia = String(STRING_CONTASCORRENTESINTERNAS_AGENCIA, 0)
    tContaCorrenteInterna.sContaContabil = String(STRING_CONTASCORRENTESINTERNAS_CONTACONTABIL, 0)
    tContaCorrenteInterna.sContato = String(STRING_CONTASCORRENTESINTERNAS_CONTATO, 0)
    tContaCorrenteInterna.sDescricao = String(STRING_CONTASCORRENTESINTERNAS_DESCRICAO, 0)
    tContaCorrenteInterna.sDVAgConta = String(STRING_CONTASCORRENTESINTERNAS_DVAGCONTA, 0)
    tContaCorrenteInterna.sDVAgencia = String(STRING_CONTASCORRENTESINTERNAS_DVAGENCIA, 0)
    tContaCorrenteInterna.sDVNumConta = String(STRING_CONTASCORRENTESINTERNAS_DVNUMCONTA, 0)
    tContaCorrenteInterna.sFax = String(STRING_CONTASCORRENTESINTERNAS_FAX, 0)
    tContaCorrenteInterna.sNomeReduzido = String(STRING_CONTASCORRENTESINTERNAS_NOMEREDUZIDO, 0)
    tContaCorrenteInterna.sNumConta = String(STRING_CONTASCORRENTESINTERNAS_NUMCONTA, 0)
    tContaCorrenteInterna.sTelefone = String(STRING_CONTASCORRENTESINTERNAS_TELEFONE, 0)
    
    'seleciona todos os registros da tabela
    lErro = Comando_Executar(lComando, "SELECT ChequePre, Codigo , FilialEmpresa, NomeReduzido, Descricao, CodBanco, Agencia, DVAgencia, NumConta, DVNumConta, DVAgConta, Contato, Telefone, Fax, SaldoInicial, DataSaldoInicial, ContaContabil, ProxSeqMov, NumMenorExtratoNaoConciliado, ProxBordero, ChequeBordero FROM ContasCorrentesInternas", _
    tContaCorrenteInterna.iChequePre, tContaCorrenteInterna.iCodigo, tContaCorrenteInterna.iFilialEmpresa, tContaCorrenteInterna.sNomeReduzido, tContaCorrenteInterna.sDescricao, tContaCorrenteInterna.iCodBanco, tContaCorrenteInterna.sAgencia, tContaCorrenteInterna.sDVAgencia, tContaCorrenteInterna.sNumConta, tContaCorrenteInterna.sDVNumConta, tContaCorrenteInterna.sDVAgConta, tContaCorrenteInterna.sContato, tContaCorrenteInterna.sTelefone, tContaCorrenteInterna.sFax, tContaCorrenteInterna.dSaldoInicial, tContaCorrenteInterna.dtDataSaldoInicial, tContaCorrenteInterna.sContaContabil, tContaCorrenteInterna.lProxSeq, tContaCorrenteInterna.iNumMenorExtratoNaoConciliado, tContaCorrenteInterna.lProxBordero, tContaCorrenteInterna.iChequeBordero)
    If lErro <> AD_SQL_SUCESSO Then gError 103863
    
    'busca o primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then gError 103864
    
    'se a tabela estiver vazia -> erro
    If lErro = AD_SQL_SEM_DADOS Then gError 103865
    
    'enquanto houver dado a ser lido
    Do While (lErro <> AD_SQL_SEM_DADOS)
        
        'cria um objeto novo
        Set objContaCorrenteInterna = New ClassContasCorrentesInternas
        
        'e o preenche com o registro lido da tabela
        With objContaCorrenteInterna
        
            .iChequePre = tContaCorrenteInterna.iChequePre
            .dSaldoInicial = tContaCorrenteInterna.dSaldoInicial
            .dtDataInicial = tContaCorrenteInterna.dtDataSaldoInicial
            .iChequeBordero = tContaCorrenteInterna.iChequeBordero
            .iCodBanco = tContaCorrenteInterna.iCodBanco
            .iCodigo = tContaCorrenteInterna.iCodigo
            .iFilialEmpresa = tContaCorrenteInterna.iFilialEmpresa
            .iNumMenorExtratoNaoConciliado = tContaCorrenteInterna.iNumMenorExtratoNaoConciliado
            .lProxBordero = tContaCorrenteInterna.lProxBordero
            .lProxSeqMov = tContaCorrenteInterna.lProxSeq
            .sAgencia = tContaCorrenteInterna.sAgencia
            .sContaContabil = tContaCorrenteInterna.sContaContabil
            .sContato = tContaCorrenteInterna.sContato
            .sDescricao = tContaCorrenteInterna.sDescricao
            .sDVAgConta = tContaCorrenteInterna.sDVAgConta
            .sDVAgencia = tContaCorrenteInterna.sDVAgencia
            .sDVNumConta = tContaCorrenteInterna.sDVNumConta
            .sFax = tContaCorrenteInterna.sFax
            .sNomeReduzido = tContaCorrenteInterna.sNomeReduzido
            .sNumConta = tContaCorrenteInterna.sNumConta
            .sTelefone = tContaCorrenteInterna.sTelefone
        
        End With
        
        'coloca-o na coleção
        colContasCorrentesInternas.Add objContaCorrenteInterna
        
        'busca o próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then gError 103866
        
    Loop
    
    'fecha comando
    Call Comando_Fechar(lComando)

    ContasCorrentesInternas_Le_Todas = SUCESSO
    
    Exit Function

Erro_ContasCorrentesInternas_Le_Todas:

    ContasCorrentesInternas_Le_Todas = gErr
    
    Select Case gErr
    
        Case 103862
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 103863, 103864, 103866
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTASCORRENTESINTERNAS_LEITURA", gErr, objContaCorrenteInterna.iCodigo)
        
        Case 103865
            'será tratado na rotina chamadora
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151423)
    
    End Select
    
    'fecha comando
    Call Comando_Fechar(lComando)

    Exit Function
    
End Function

Public Function ContasCorrentesInternas_Le(ByVal objContaCorrenteInterna As ClassContasCorrentesInternas) As Long
'Função que recebe um objContaCorrenteInterna com o código preenchido, busca na tabela e, achando e preenche os seus
'atributos.

Dim lErro As Long
Dim lComando As Long
Dim tContaCorrenteInterna As typeContaCorrenteInt

On Error GoTo Erro_ContasCorrentesInternas_Le

    'Abre comando
    lComando = Comando_Abrir
    If lComando = 0 Then gError 103935
    
    'inicializa as strings
    tContaCorrenteInterna.sAgencia = String(STRING_CONTASCORRENTESINTERNAS_AGENCIA, 0)
    tContaCorrenteInterna.sContaContabil = String(STRING_CONTASCORRENTESINTERNAS_CONTACONTABIL, 0)
    tContaCorrenteInterna.sContato = String(STRING_CONTASCORRENTESINTERNAS_CONTATO, 0)
    tContaCorrenteInterna.sDescricao = String(STRING_CONTASCORRENTESINTERNAS_DESCRICAO, 0)
    tContaCorrenteInterna.sDVAgConta = String(STRING_CONTASCORRENTESINTERNAS_DVAGCONTA, 0)
    tContaCorrenteInterna.sDVAgencia = String(STRING_CONTASCORRENTESINTERNAS_DVAGENCIA, 0)
    tContaCorrenteInterna.sDVNumConta = String(STRING_CONTASCORRENTESINTERNAS_DVNUMCONTA, 0)
    tContaCorrenteInterna.sFax = String(STRING_CONTASCORRENTESINTERNAS_FAX, 0)
    tContaCorrenteInterna.sNomeReduzido = String(STRING_CONTASCORRENTESINTERNAS_NOMEREDUZIDO, 0)
    tContaCorrenteInterna.sNumConta = String(STRING_CONTASCORRENTESINTERNAS_NUMCONTA, 0)
    tContaCorrenteInterna.sTelefone = String(STRING_CONTASCORRENTESINTERNAS_TELEFONE, 0)
    
    'seleciona a conta corrente passada por parâmetro
    lErro = Comando_Executar(lComando, "SELECT FilialEmpresa, NomeReduzido, Descricao, CodBanco, Agencia, DVAgencia, NumConta, DVNumConta, DVAgConta, Contato, Telefone, Fax, SaldoInicial, DataSaldoInicial, ContaContabil, ProxSeqMov, NumMenorExtratoNaoConciliado, ProxBordero, ChequeBordero FROM ContasCorrentesInternas WHERE Codigo = ?", _
    tContaCorrenteInterna.iFilialEmpresa, tContaCorrenteInterna.sNomeReduzido, tContaCorrenteInterna.sDescricao, tContaCorrenteInterna.iCodBanco, tContaCorrenteInterna.sAgencia, tContaCorrenteInterna.sDVAgencia, tContaCorrenteInterna.sNumConta, tContaCorrenteInterna.sDVNumConta, tContaCorrenteInterna.sDVAgConta, tContaCorrenteInterna.sContato, tContaCorrenteInterna.sTelefone, tContaCorrenteInterna.sFax, tContaCorrenteInterna.dSaldoInicial, tContaCorrenteInterna.dtDataSaldoInicial, tContaCorrenteInterna.sContaContabil, tContaCorrenteInterna.lProxSeq, tContaCorrenteInterna.iNumMenorExtratoNaoConciliado, tContaCorrenteInterna.lProxBordero, tContaCorrenteInterna.iChequeBordero, objContaCorrenteInterna.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 103936
    
    'busca a primeira
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 103937
    
    'se não encontrou->erro
    If lErro = AD_SQL_SEM_DADOS Then gError 103938
    
    'o preenche com o registro lido da tabela
    With objContaCorrenteInterna
    
        .dSaldoInicial = tContaCorrenteInterna.dSaldoInicial
        .dtDataInicial = tContaCorrenteInterna.dtDataSaldoInicial
        .iChequeBordero = tContaCorrenteInterna.iChequeBordero
        .iCodBanco = tContaCorrenteInterna.iCodBanco
        .iFilialEmpresa = tContaCorrenteInterna.iFilialEmpresa
        .iNumMenorExtratoNaoConciliado = tContaCorrenteInterna.iNumMenorExtratoNaoConciliado
        .lProxBordero = tContaCorrenteInterna.lProxBordero
        .lProxSeqMov = tContaCorrenteInterna.lProxSeq
        .sAgencia = tContaCorrenteInterna.sAgencia
        .sContaContabil = tContaCorrenteInterna.sContaContabil
        .sContato = tContaCorrenteInterna.sContato
        .sDescricao = tContaCorrenteInterna.sDescricao
        .sDVAgConta = tContaCorrenteInterna.sDVAgConta
        .sDVAgencia = tContaCorrenteInterna.sDVAgencia
        .sDVNumConta = tContaCorrenteInterna.sDVNumConta
        .sFax = tContaCorrenteInterna.sFax
        .sNomeReduzido = tContaCorrenteInterna.sNomeReduzido
        .sNumConta = tContaCorrenteInterna.sNumConta
        .sTelefone = tContaCorrenteInterna.sTelefone
    
    End With
    
    'fecha o comando
    Call Comando_Fechar(lComando)
    
    ContasCorrentesInternas_Le = SUCESSO

    Exit Function

Erro_ContasCorrentesInternas_Le:
    
    ContasCorrentesInternas_Le = gErr
    
    Select Case gErr
    
        Case 103935
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 103936, 103937
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTASCORRENTESINTERNAS_LEITURA", gErr, objContaCorrenteInterna.iCodigo)
            
        Case 103938
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151424)
    
    End Select

    'fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function ContasCorrentesInternas_Lock(ByVal objContaCorrenteInterna As ClassContasCorrentesInternas) As Long
'Função que recebe um objContaCorrenteInterna com o código preenchido, busca na tabela e, achando, preenche os seus
'atributos e o locka. Deve ser chamada dentro de transação.

Dim lErro As Long
Dim lComando As Long
Dim tContaCorrenteInterna As typeContaCorrenteInt

On Error GoTo Erro_ContasCorrentesInternas_Lock

    'Abre comando
    lComando = Comando_Abrir
    If lComando = 0 Then gError 103897
    
    'inicializa as strings
    tContaCorrenteInterna.sAgencia = String(STRING_CONTASCORRENTESINTERNAS_AGENCIA, 0)
    tContaCorrenteInterna.sContaContabil = String(STRING_CONTASCORRENTESINTERNAS_CONTACONTABIL, 0)
    tContaCorrenteInterna.sContato = String(STRING_CONTASCORRENTESINTERNAS_CONTATO, 0)
    tContaCorrenteInterna.sDescricao = String(STRING_CONTASCORRENTESINTERNAS_DESCRICAO, 0)
    tContaCorrenteInterna.sDVAgConta = String(STRING_CONTASCORRENTESINTERNAS_DVAGCONTA, 0)
    tContaCorrenteInterna.sDVAgencia = String(STRING_CONTASCORRENTESINTERNAS_DVAGENCIA, 0)
    tContaCorrenteInterna.sDVNumConta = String(STRING_CONTASCORRENTESINTERNAS_DVNUMCONTA, 0)
    tContaCorrenteInterna.sFax = String(STRING_CONTASCORRENTESINTERNAS_FAX, 0)
    tContaCorrenteInterna.sNomeReduzido = String(STRING_CONTASCORRENTESINTERNAS_NOMEREDUZIDO, 0)
    tContaCorrenteInterna.sNumConta = String(STRING_CONTASCORRENTESINTERNAS_NUMCONTA, 0)
    tContaCorrenteInterna.sTelefone = String(STRING_CONTASCORRENTESINTERNAS_TELEFONE, 0)
    
    'seleciona a conta corrente passada por parâmetro
    lErro = Comando_ExecutarPos(lComando, "SELECT FilialEmpresa, NomeReduzido, Descricao, CodBanco, Agencia, DVAgencia, NumConta, DVNumConta, DVAgConta, Contato, Telefone, Fax, SaldoInicial, DataSaldoInicial, ContaContabil, ProxSeqMov, NumMenorExtratoNaoConciliado, ProxBordero, ChequeBordero FROM ContasCorrentesInternas WHERE Codigo = ?", 0, _
    tContaCorrenteInterna.iFilialEmpresa, tContaCorrenteInterna.sNomeReduzido, tContaCorrenteInterna.sDescricao, tContaCorrenteInterna.iCodBanco, tContaCorrenteInterna.sAgencia, tContaCorrenteInterna.sDVAgencia, tContaCorrenteInterna.sNumConta, tContaCorrenteInterna.sDVNumConta, tContaCorrenteInterna.sDVAgConta, tContaCorrenteInterna.sContato, tContaCorrenteInterna.sTelefone, tContaCorrenteInterna.sFax, tContaCorrenteInterna.dSaldoInicial, tContaCorrenteInterna.dtDataSaldoInicial, tContaCorrenteInterna.sContaContabil, tContaCorrenteInterna.lProxSeq, tContaCorrenteInterna.iNumMenorExtratoNaoConciliado, tContaCorrenteInterna.lProxBordero, tContaCorrenteInterna.iChequeBordero, objContaCorrenteInterna.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 103898
    
    'busca a primeira
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 103899
    
    'se não encontrou->erro
    If lErro = AD_SQL_SEM_DADOS Then gError 103900
    
    'locka a conta corrente
    lErro = Comando_LockShared(lComando)
    If lErro <> AD_SQL_SUCESSO Then gError 103901
    
    'o preenche com o registro lido da tabela
    With objContaCorrenteInterna
    
        .dSaldoInicial = tContaCorrenteInterna.dSaldoInicial
        .dtDataInicial = tContaCorrenteInterna.dtDataSaldoInicial
        .iChequeBordero = tContaCorrenteInterna.iChequeBordero
        .iCodBanco = tContaCorrenteInterna.iCodBanco
        .iFilialEmpresa = tContaCorrenteInterna.iFilialEmpresa
        .iNumMenorExtratoNaoConciliado = tContaCorrenteInterna.iNumMenorExtratoNaoConciliado
        .lProxBordero = tContaCorrenteInterna.lProxBordero
        .lProxSeqMov = tContaCorrenteInterna.lProxSeq
        .sAgencia = tContaCorrenteInterna.sAgencia
        .sContaContabil = tContaCorrenteInterna.sContaContabil
        .sContato = tContaCorrenteInterna.sContato
        .sDescricao = tContaCorrenteInterna.sDescricao
        .sDVAgConta = tContaCorrenteInterna.sDVAgConta
        .sDVAgencia = tContaCorrenteInterna.sDVAgencia
        .sDVNumConta = tContaCorrenteInterna.sDVNumConta
        .sFax = tContaCorrenteInterna.sFax
        .sNomeReduzido = tContaCorrenteInterna.sNomeReduzido
        .sNumConta = tContaCorrenteInterna.sNumConta
        .sTelefone = tContaCorrenteInterna.sTelefone
    
    End With
    
    'fecha o comando
    Call Comando_Fechar(lComando)
    
    ContasCorrentesInternas_Lock = SUCESSO

    Exit Function

Erro_ContasCorrentesInternas_Lock:
    
    ContasCorrentesInternas_Lock = gErr
    
    Select Case gErr
    
        Case 103897
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 103898, 103899
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTASCORRENTESINTERNAS_LEITURA", gErr, objContaCorrenteInterna.iCodigo)
            
        Case 103900
            
        Case 103901
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTASCORRENTESINTERNAS_LOCK", objContaCorrenteInterna.iCodigo)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151425)
    
    End Select

    'fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function MovimentosCaixa_Le_NumIntDoc(ByVal objMovimentoCaixa As ClassMovimentoCaixa) As Long
'Função que recebe um objMovimento caixa com o seu número e filial preenchidos
'e gera um número sequencial para o movimento

Dim lErro As Long
Dim lComando As Long
Dim tMovimentoCaixa As typeMovimentoCaixa

On Error GoTo Erro_MovimentosCaixa_Le_NumIntDoc

    'abre o comando
    lComando = Comando_Abrir
    If lComando = 0 Then gError 103674

    'inicializa as strings
    tMovimentoCaixa.sFavorecido = String(STRING_MOVIMENTOCAIXA_FAVORECIDO, 0)
    tMovimentoCaixa.sHistorico = String(STRING_MOVIMENTOCAIXA_HISTORICO, 0)

    'seleciona o registro com o numintdoc passado
    lErro = Comando_Executar(lComando, "SELECT Caixa, Sequencial, Tipo, AdmMeioPagto, Parcelamento, TipoCartao, Numero, DataMovimento, Valor, Historico, Favorecido, CupomFiscal, NumRefInterna, MovtoTransf, MovtoEstorno, Gerente,CodConta,Hora FROM MovimentosCaixa WHERE FilialEmpresa = ? AND NumMovto = ? ", _
            tMovimentoCaixa.iCaixa, tMovimentoCaixa.lSequencial, tMovimentoCaixa.iTipo, tMovimentoCaixa.iAdmMeioPagto, tMovimentoCaixa.iParcelamento, tMovimentoCaixa.iTipoCartao, tMovimentoCaixa.lNumero, tMovimentoCaixa.dtDataMovimento, tMovimentoCaixa.dValor, tMovimentoCaixa.sHistorico, tMovimentoCaixa.sFavorecido, tMovimentoCaixa.lCupomFiscal, tMovimentoCaixa.lNumRefInterna, tMovimentoCaixa.lMovtoTransf, tMovimentoCaixa.lMovtoEstorno, tMovimentoCaixa.iGerente, tMovimentoCaixa.iCodConta, tMovimentoCaixa.dHora, _
            objMovimentoCaixa.iFilialEmpresa, objMovimentoCaixa.lNumMovto)
    If lErro <> AD_SQL_SUCESSO Then gError 103675

    'busca o primeiro na seleção
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 103676

    'se não encontrar--> erro
    If lErro = AD_SQL_SEM_DADOS Then gError 103677

    'preenche o objMovimentoCaixa
    objMovimentoCaixa.dtDataMovimento = tMovimentoCaixa.dtDataMovimento
    objMovimentoCaixa.dValor = tMovimentoCaixa.dValor
    objMovimentoCaixa.iAdmMeioPagto = tMovimentoCaixa.iAdmMeioPagto
    objMovimentoCaixa.iCaixa = tMovimentoCaixa.iCaixa
    objMovimentoCaixa.iGerente = tMovimentoCaixa.iGerente
    objMovimentoCaixa.iParcelamento = tMovimentoCaixa.iParcelamento
    objMovimentoCaixa.iTipo = tMovimentoCaixa.iTipo
    objMovimentoCaixa.lCupomFiscal = tMovimentoCaixa.lCupomFiscal
    objMovimentoCaixa.lMovtoEstorno = tMovimentoCaixa.lMovtoEstorno
    objMovimentoCaixa.lMovtoTransf = tMovimentoCaixa.lMovtoTransf
    objMovimentoCaixa.lNumero = tMovimentoCaixa.lNumero
    objMovimentoCaixa.lNumRefInterna = tMovimentoCaixa.lNumRefInterna
    objMovimentoCaixa.lSequencial = tMovimentoCaixa.lSequencial
    objMovimentoCaixa.sFavorecido = tMovimentoCaixa.sFavorecido
    objMovimentoCaixa.sHistorico = tMovimentoCaixa.sHistorico
    objMovimentoCaixa.iCodConta = tMovimentoCaixa.iCodConta
    objMovimentoCaixa.dHora = tMovimentoCaixa.dHora

    MovimentosCaixa_Le_NumIntDoc = SUCESSO

    'fecha os comandos
    Call Comando_Fechar(lComando)

    Exit Function

Erro_MovimentosCaixa_Le_NumIntDoc:

    MovimentosCaixa_Le_NumIntDoc = gErr

    Select Case gErr

        Case 103674
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 103675, 103676
            Call Rotina_Erro(vbOKOnly, "ERRO_MOVIMENTOSCAIXA_LEITURA", gErr)

        Case 103677
        'será tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151426)

    End Select

    'fecha os comandos
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function MovimentosCaixa_Le(ByVal objMovimentoCaixa As ClassMovimentoCaixa) As Long
'função que recebe um objmovimento caixa preenchido com seu codcaixa e seu sequencial preenchidos
'e o retorna preenchido com os dados da tabela de movimentosCaixa

Dim lErro As Long
Dim lComando As Long
Dim tMovimentoCaixa As typeMovimentoCaixa

On Error GoTo Erro_MovimentosCaixa_Le

    'abre comando
    lComando = Comando_Abrir
    If lComando = 0 Then gError 103741

    'inicializa as strings
    tMovimentoCaixa.sHistorico = String(STRING_MOVIMENTOCAIXA_HISTORICO, 0)
    tMovimentoCaixa.sFavorecido = String(STRING_MOVIMENTOCAIXA_FAVORECIDO, 0)

    'faz o select na tabela
    lErro = Comando_Executar(lComando, "SELECT FilialEmpresa, NumMovto, Tipo , AdmMeioPagto , Parcelamento, TipoCartao, Numero, DataMovimento, Valor, Historico, Favorecido,CupomFiscal,NumRefInterna,MovtoTransf,MovtoEstorno,Gerente,CodConta, Hora FROM MovimentosCaixa WHERE Caixa = ? AND Sequencial = ?", _
    tMovimentoCaixa.iFilialEmpresa, tMovimentoCaixa.lNumMovto, tMovimentoCaixa.iTipo, tMovimentoCaixa.iAdmMeioPagto, tMovimentoCaixa.iParcelamento, tMovimentoCaixa.iTipoCartao, tMovimentoCaixa.lNumero, tMovimentoCaixa.dtDataMovimento, tMovimentoCaixa.dValor, tMovimentoCaixa.sHistorico, tMovimentoCaixa.sFavorecido, tMovimentoCaixa.lCupomFiscal, tMovimentoCaixa.lNumRefInterna, tMovimentoCaixa.lMovtoTransf, tMovimentoCaixa.lMovtoEstorno, tMovimentoCaixa.iGerente, tMovimentoCaixa.iCodConta, tMovimentoCaixa.dHora, objMovimentoCaixa.iCaixa, objMovimentoCaixa.lSequencial)
    If lErro <> AD_SQL_SUCESSO Then gError 103742

    'busca o primeiro na seleção
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 103743

    'se não encontrou --> erro
    If lErro = AD_SQL_SEM_DADOS Then gError 103744

    'carrega o obj com os dados da tabela
    With objMovimentoCaixa

        .dtDataMovimento = tMovimentoCaixa.dtDataMovimento
        .dValor = tMovimentoCaixa.dValor
        .iAdmMeioPagto = tMovimentoCaixa.iAdmMeioPagto
        .iFilialEmpresa = tMovimentoCaixa.iFilialEmpresa
        .iGerente = tMovimentoCaixa.iGerente
        .iParcelamento = tMovimentoCaixa.iParcelamento
        .iTipo = tMovimentoCaixa.iTipo
        .iTipoCartao = tMovimentoCaixa.iTipoCartao
        .lCupomFiscal = tMovimentoCaixa.lCupomFiscal
        .lMovtoEstorno = tMovimentoCaixa.lMovtoEstorno
        .lMovtoTransf = tMovimentoCaixa.lMovtoTransf
        .lNumero = tMovimentoCaixa.lNumero
        .lNumMovto = tMovimentoCaixa.lNumMovto
        .lNumRefInterna = tMovimentoCaixa.lNumRefInterna
        .sFavorecido = tMovimentoCaixa.sFavorecido
        .sHistorico = tMovimentoCaixa.sHistorico
        .iCodConta = tMovimentoCaixa.iCodConta
        .dHora = tMovimentoCaixa.dHora

    End With

    'fecha comandos
    Call Comando_Fechar(lComando)

    MovimentosCaixa_Le = SUCESSO

    Exit Function

Erro_MovimentosCaixa_Le:

    MovimentosCaixa_Le = gErr

    Select Case gErr

        Case 103741
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 103742, 103743
            Call Rotina_Erro(vbOKOnly, "ERRO_MOVIMENTOSCAIXA_LEITURA", gErr)

        Case 103744
        'será tratada na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151427)

    End Select

    'fecha comandos
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function CCMov_Le(ByVal objCCMov As ClassCCMov) As Long
'Função que lê um movimento anual e preenche o objCCMov passado por parâmetro

Dim lErro As Long
Dim lComando As Long
Dim tCCMov As typeCCMov

On Error GoTo Erro_CCMov_Le

    'abre o comando
    lComando = Comando_Abrir
    If lComando = 0 Then gError 103730

    'seleciona o objCCMov passado como parametro
    lErro = Comando_Executar(lComando, "SELECT CodCaixa, FilialEmpresa, Ano, TipoMeioPagto, AdmMeioPagto, Parcelamento, SaldoInicial, Deb01, Cred01,Deb02, Cred02,Deb03, Cred03,Deb04, Cred04,Deb05, Cred05,Deb06, Cred06,Deb07, Cred07,Deb08, Cred08,Deb09, Cred09,Deb10, Cred10,Deb11, Cred11,Deb12, Cred12 FROM CCMov WHERE CodCaixa = ? AND FilialEmpresa = ? AND Ano = ? AND TipoMeioPagto = ? AND AdmMeioPagto = ? AND Parcelamento = ?", _
    tCCMov.iCodCaixa, tCCMov.iFilialEmpresa, tCCMov.iAno, tCCMov.iTipoMeioPagto, tCCMov.iAdmMeioPagto, tCCMov.iParcelamento, tCCMov.dSaldoInicial, tCCMov.dDeb01, tCCMov.dCred01, tCCMov.dDeb02, tCCMov.dCred02, tCCMov.dDeb03, tCCMov.dCred03, tCCMov.dDeb04, tCCMov.dCred04, tCCMov.dDeb05, tCCMov.dCred05, tCCMov.dDeb06, tCCMov.dCred06, tCCMov.dDeb07, tCCMov.dCred07, tCCMov.dDeb08, tCCMov.dCred08, tCCMov.dDeb09, tCCMov.dCred09, tCCMov.dDeb10, tCCMov.dCred10, tCCMov.dDeb11, tCCMov.dCred11, tCCMov.dDeb12, tCCMov.dCred12, objCCMov.iCodCaixa, objCCMov.iFilialEmpresa, objCCMov.iAno, objCCMov.iTipoMeioPagto, objCCMov.iAdmMeioPagto, objCCMov.iParcelamento)
    If lErro <> AD_SQL_SUCESSO Then gError 103731

    'seleciona o primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 103732

    'se não encontrou, erro
    If lErro = AD_SQL_SEM_DADOS Then gError 103733

    'preenche o objCCmov com os dados da tabela
    With objCCMov

        .iCodCaixa = tCCMov.iCodCaixa
        .iFilialEmpresa = tCCMov.iFilialEmpresa
        .iAno = tCCMov.iAno
        .iTipoMeioPagto = tCCMov.iTipoMeioPagto
        .iAdmMeioPagto = tCCMov.iAdmMeioPagto
        .iParcelamento = tCCMov.iParcelamento
        .dSaldoInicial = tCCMov.dSaldoInicial
        .dDeb01 = tCCMov.dDeb01
        .dCred01 = tCCMov.dCred01
        .dDeb02 = tCCMov.dDeb02
        .dCred02 = tCCMov.dCred02
        .dDeb03 = tCCMov.dDeb03
        .dCred03 = tCCMov.dCred03
        .dDeb04 = tCCMov.dDeb04
        .dCred04 = tCCMov.dCred04
        .dDeb05 = tCCMov.dDeb05
        .dCred05 = tCCMov.dCred05
        .dDeb06 = tCCMov.dDeb06
        .dCred06 = tCCMov.dCred06
        .dDeb07 = tCCMov.dDeb07
        .dCred07 = tCCMov.dCred07
        .dDeb08 = tCCMov.dDeb08
        .dCred08 = tCCMov.dCred08
        .dDeb09 = tCCMov.dDeb09
        .dCred09 = tCCMov.dCred09
        .dDeb10 = tCCMov.dDeb10
        .dCred10 = tCCMov.dCred10
        .dDeb11 = tCCMov.dDeb11
        .dCred11 = tCCMov.dCred11
        .dDeb12 = tCCMov.dDeb12
        .dCred12 = tCCMov.dCred12

    End With

    'fecha o comando
    Call Comando_Fechar(lComando)

    CCMov_Le = SUCESSO

    Exit Function

Erro_CCMov_Le:

    CCMov_Le = gErr

    Select Case gErr

        Case 103730
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 103731, 103732
            Call Rotina_Erro(vbOKOnly, "ERRO_CCMOV_LEITURA", gErr)

        Case 103733
            'será tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151428)

    End Select

    'fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function Caixa_Lock(ByVal objCaixa As ClassCaixa) As Long
'Função que recebe um objCaixa, o preenche completamente e o locka.
'Deve ser chamada dentro de transação

Dim alComando(1) As Long
Dim lErro As Long
Dim iIndice As Integer
Dim tCaixa As typeCaixa

On Error GoTo Erro_Caixa_Lock

    'abre comandos
    For iIndice = LBound(alComando) To UBound(alComando)
    
        alComando(iIndice) = Comando_Abrir
        If alComando(iIndice) = 0 Then gError 103773
        
    Next
     
    'inicializa as strings
    tCaixa.sNomeReduzido = String(STRING_CAIXA_NOME_REDUZIDO, 0)
    tCaixa.sDescricao = String(STRING_CAIXA_DESCRICAO, 0)
    
    'seleciona o caixa na tabela de caixas
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Codigo,FilialEmpresa,NomeReduzido,Descricao,Status,DataInicial,ProxSeqMov,POS,TEF,BoletoManual,Ativo  FROM Caixa WHERE Codigo = ? AND FilialEmpresa = ?", _
    0, tCaixa.iCodigo, tCaixa.iFilialEmpresa, tCaixa.sNomeReduzido, tCaixa.sDescricao, tCaixa.iStatus, tCaixa.dtDataInicial, tCaixa.lProxSeqMov, tCaixa.iPos, tCaixa.iTEF, tCaixa.iBoletoManual, tCaixa.iAtivo, objCaixa.iCodigo, objCaixa.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 103774
    
    'busca o primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 103775
    
    'se não encontrou o caixa--> erro
    If lErro = AD_SQL_SEM_DADOS Then gError 103776
    
    'locka o caixa selecionado
    lErro = Comando_LockShared(alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 103777
    
    'preenche o objcaixa
    With objCaixa
    
        .dtDataInicial = tCaixa.dtDataInicial
        .iCodigo = tCaixa.iCodigo
        .iFilialEmpresa = tCaixa.iFilialEmpresa
        .iStatus = tCaixa.iStatus
        .lProxSeqMov = tCaixa.lProxSeqMov
        .sDescricao = tCaixa.sDescricao
        .sNomeReduzido = tCaixa.sNomeReduzido
        .iPos = tCaixa.iPos
        .iTEF = tCaixa.iTEF
        .iBoletoManual = tCaixa.iBoletoManual
        .iAtivo = tCaixa.iAtivo
       
    End With
    
    'fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
    
        Call Comando_Fechar(alComando(iIndice))
        
    Next
    
    Caixa_Lock = SUCESSO
    
    Exit Function

Erro_Caixa_Lock:
    
    Caixa_Lock = gErr
    
    Select Case gErr
    
        Case 103773
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 103774, 103775
            Call Rotina_Erro(vbOKOnly, "ERRO_CAIXA_LEITURA", gErr)
        
        Case 103776
            'será tratado na rotina chamadora
        
        Case 103777
            Call Rotina_Erro(vbOKOnly, "ERRO_CAIXA_LOCK", gErr, objCaixa.iCodigo, objCaixa.iFilialEmpresa)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151429)

    End Select
    
    'fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
    
        Call Comando_Fechar(alComando(iIndice))
        
    Next
    
    Exit Function

End Function
'Rafael Menezes( Fim ) - Tela: SaqueCaixa

Function LojaConfig_Le(ByVal objConfiguraLoja As ClassLoja) As Long
'Função responsável pela leitura dos dados

Dim lErro As Long
Dim sConteudo As String
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim objLojaConfig As ClassLojaConfig
Dim colLojaConfig As New colLojaConfig

On Error GoTo Erro_LojaConfig_Le

    'Abre os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 80078
    Next

    lErro = CF("ConfiguraLoja_MoverCampos_ColLojaConfig", objConfiguraLoja, colLojaConfig)
    If lErro <> SUCESSO Then gError 80079

    For Each objLojaConfig In colLojaConfig

        sConteudo = String(STRING_CONTEUDO, 0)

        'Seleciona codigo e filial na tabela LojaConfig de acordo com o codigo passado como parametro
        lErro = Comando_Executar(alComando(0), "SELECT Conteudo FROM LojaConfig WHERE Codigo = ? And FilialEmpresa = ?", sConteudo, objLojaConfig.sCodigo, objLojaConfig.iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 80080

        lErro = Comando_BuscarPrimeiro(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 80081

        'Nao encontrou registro em LojaConfig
        If lErro = AD_SQL_SEM_DADOS Then
        
            'Seleciona codigo e filial na tabela LojaConfig de acordo com o codigo passado como parametro
            lErro = Comando_Executar(alComando(0), "SELECT Conteudo FROM LojaConfig WHERE Codigo = ? And FilialEmpresa = ?", sConteudo, objLojaConfig.sCodigo, 1)
            If lErro <> AD_SQL_SUCESSO Then gError 80080

            lErro = Comando_BuscarPrimeiro(alComando(0))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 80081
        
            If lErro = AD_SQL_SEM_DADOS Then gError 80082
        
        End If
        
        objLojaConfig.sConteudo = sConteudo

    Next

    'Preenche objeto objConfiguraLoja com os dados lidos do BD
    Call ColLojaConfig_MoverCampos_ConfiguraLoja(objConfiguraLoja, colLojaConfig)

    'Preenche o GridCartoes com os dados lidos do BD
    lErro = AliquotaICMS_Le(objConfiguraLoja.colAliquotaICMS)
    If lErro <> SUCESSO Then gError 80113

    'Fecha os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next

    LojaConfig_Le = SUCESSO

    Exit Function

Erro_LojaConfig_Le:

    LojaConfig_Le = gErr

    Select Case gErr

        Case 80078
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 80079

        Case 80080, 80081
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LOJACONFIG", gErr, objLojaConfig.sCodigo)

        Case 80082
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REGISTRO_LOJA_CONFIG_NAO_ENCONTRADO", gErr, objLojaConfig.sCodigo)
            
        Case 80113

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151430)

    End Select
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function ConfiguraLoja_MoverCampos_ColLojaConfig(ByVal objConfiguraLoja As ClassLoja, ByVal colLojaConfig As colLojaConfig) As Long
'Transmite os dados (EMPRESA TODA E A FILIAL ATUAL) do objConfiguraLoja para ColConfiguraLoja

Dim lErro As Long

On Error GoTo Erro_ConfiguraLoja_MoverCampos_ColLojaConfig

    lErro = CF("ConfiguraLoja_MoverCampos_ColLojaConfig_EmpToda", objConfiguraLoja, colLojaConfig)
    If lErro <> SUCESSO Then gError 109385
    
    If giFilialEmpresa <> EMPRESA_TODA Then
        lErro = CF("ConfiguraLoja_MoverCampos_ColLojaConfig_Filial", objConfiguraLoja, colLojaConfig)
        If lErro <> SUCESSO Then gError 109386
    End If

    ConfiguraLoja_MoverCampos_ColLojaConfig = SUCESSO
    
    Exit Function

Erro_ConfiguraLoja_MoverCampos_ColLojaConfig:

    ConfiguraLoja_MoverCampos_ColLojaConfig = Err

    Select Case gErr
        
        Case 109385, 109386

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 151431)

    End Select

    Exit Function

End Function

Public Function ConfiguraLoja_MoverCampos_ColLojaConfig_EmpToda(ByVal objConfiguraLoja As ClassLoja, ByVal colLojaConfig As colLojaConfig) As Long
'Transmite os dados REFERENTES À EMPRESA TODA do objConfiguraLoja para ColConfiguraLoja

On Error GoTo Erro_ConfiguraLoja_MoverCampos_ColLojaConfig_EmpToda
    
    colLojaConfig.Add "SIMBOLO_MOEDA", EMPRESA_TODA, "", 0, objConfiguraLoja.sSimboloMoeda
    colLojaConfig.Add "TRUNCAMENTO_ARREDONDAMENTO", EMPRESA_TODA, "", 0, objConfiguraLoja.sTruncamentoArredondamento
    colLojaConfig.Add "FTP_URL", EMPRESA_TODA, "", 0, objConfiguraLoja.sFTPURL
    colLojaConfig.Add "FTP_USERNAME", EMPRESA_TODA, "", 0, objConfiguraLoja.sFTPUserName
    colLojaConfig.Add "FTP_PASSWORD", EMPRESA_TODA, "", 0, objConfiguraLoja.sFTPPassword
    colLojaConfig.Add "FTP_DIRETORIO", EMPRESA_TODA, "", 0, objConfiguraLoja.sFTPDiretorio

    ConfiguraLoja_MoverCampos_ColLojaConfig_EmpToda = SUCESSO
    
    Exit Function
    
Erro_ConfiguraLoja_MoverCampos_ColLojaConfig_EmpToda:
    
    ConfiguraLoja_MoverCampos_ColLojaConfig_EmpToda = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 151432)

    End Select
    
    Exit Function

End Function

Public Function ConfiguraLoja_MoverCampos_ColLojaConfig_Filial(ByVal objConfiguraLoja As ClassLoja, ByVal colLojaConfig As colLojaConfig) As Long
'Transmite os dados REFERENTES À FILIAL ATUAL do objConfiguraLoja para ColConfiguraLoja

On Error GoTo Erro_ConfiguraLoja_MoverCampos_ColLojaConfig_Filial

    colLojaConfig.Add "CUPOM_DESCREVE_FORMA_PAGTO", giFilialEmpresa, "", 0, CStr(objConfiguraLoja.iCupomDescreveFormaPagto)
    colLojaConfig.Add "ESPACO_ENTRE_LINHAS", giFilialEmpresa, "", 0, CStr(objConfiguraLoja.lEspacoEntreLinhas)
    colLojaConfig.Add "LINHAS_ENTRE_CUPONS", giFilialEmpresa, "", 0, CStr(objConfiguraLoja.iLinhasEntreCupons)
    colLojaConfig.Add "HORARIO_VERAO", giFilialEmpresa, "", 0, CStr(objConfiguraLoja.iHorarioVerao)
    colLojaConfig.Add "IMPRIME_ITEM_A_ITEM", giFilialEmpresa, "", 0, CStr(objConfiguraLoja.iImprimeItemAItem)
    colLojaConfig.Add "MENSAGEM_CUPOM", giFilialEmpresa, "", 0, objConfiguraLoja.sMensagemCupom
    colLojaConfig.Add "NATUREZA_OPERACAO_PADRAO", giFilialEmpresa, "", 0, objConfiguraLoja.sNatOpPadrao
    colLojaConfig.Add "NUM_LIM_RO", giFilialEmpresa, "", 0, objConfiguraLoja.sNumLimRO
    colLojaConfig.Add "OPERADOR_IGUAL_VENDEDOR", giFilialEmpresa, "", 0, CStr(objConfiguraLoja.iOperadorIgualVendedor)
    colLojaConfig.Add "TABELA_PRECO_LOJA", giFilialEmpresa, "", 0, CStr(objConfiguraLoja.iTabelaPreco)
    colLojaConfig.Add "VENDEDOR_OBRIGATORIO", giFilialEmpresa, "", 0, CStr(objConfiguraLoja.iVendedorObrigatorio)
    colLojaConfig.Add "TELA_VENDA_M_P", giFilialEmpresa, "", 0, CStr(objConfiguraLoja.iSemTeclado)
    colLojaConfig.Add "OP_CAIXA_NECESSITA_GERENTE", giFilialEmpresa, "", 0, CStr(objConfiguraLoja.iGerenteAutoriza)
    colLojaConfig.Add "CONTACONTABIL_CAIXA_CENTRAL", giFilialEmpresa, "", 0, objConfiguraLoja.sContaContabil
    colLojaConfig.Add "ATUALIZACAO_ECF", giFilialEmpresa, "", 0, CStr(objConfiguraLoja.iAtualizacaoECF)
    colLojaConfig.Add "ABRE_APOS_FECHAMENTO", giFilialEmpresa, "", 0, CStr(objConfiguraLoja.iAbreAposFechamento)
    colLojaConfig.Add "INTERVALO_TRANSFERENCIA", giFilialEmpresa, "", 0, CStr(objConfiguraLoja.lIntervaloTrans)
    
    ConfiguraLoja_MoverCampos_ColLojaConfig_Filial = SUCESSO
    
    Exit Function
    
Erro_ConfiguraLoja_MoverCampos_ColLojaConfig_Filial:
    
    ConfiguraLoja_MoverCampos_ColLojaConfig_Filial = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 151433)

    End Select

End Function

Function ColLojaConfig_MoverCampos_ConfiguraLoja(ByVal objConfiguraLoja As ClassLoja, ByVal colLojaConfig As colLojaConfig)
'Preenche objConfiguraLoja com os dados de objLojaConfig

Dim objLojaConfig As ClassLojaConfig
Dim lErro As Long

On Error GoTo Erro_ColLojaConfig_MoverCampos_ConfiguraLoja

    For Each objLojaConfig In colLojaConfig

        Select Case objLojaConfig.sCodigo

            Case "ATUALIZACAO_ECF"
                objConfiguraLoja.iAtualizacaoECF = StrParaInt(objLojaConfig.sConteudo)
            
            Case "CONTACONTABIL_CAIXA_CENTRAL"
                objConfiguraLoja.sContaContabil = objLojaConfig.sConteudo
            
            Case "CUPOM_DESCREVE_FORMA_PAGTO"
                objConfiguraLoja.iCupomDescreveFormaPagto = StrParaInt(objLojaConfig.sConteudo)

            Case "ESPACO_ENTRE_LINHAS"
                objConfiguraLoja.lEspacoEntreLinhas = StrParaLong(objLojaConfig.sConteudo)

            Case "HORARIO_VERAO"
                objConfiguraLoja.iHorarioVerao = StrParaInt(objLojaConfig.sConteudo)
            
            Case "IMPRIME_ITEM_A_ITEM"
                objConfiguraLoja.iImprimeItemAItem = StrParaInt(objLojaConfig.sConteudo)
            
            Case "LINHAS_ENTRE_CUPONS"
                objConfiguraLoja.iLinhasEntreCupons = StrParaInt(objLojaConfig.sConteudo)

            Case "MENSAGEM_CUPOM"
                objConfiguraLoja.sMensagemCupom = objLojaConfig.sConteudo
            
            Case "NATUREZA_OPERACAO_PADRAO"
                objConfiguraLoja.sNatOpPadrao = objLojaConfig.sConteudo
            
            Case "NUM_LIM_RO"
                objConfiguraLoja.sNumLimRO = objLojaConfig.sConteudo
            
            Case "OP_CAIXA_NECESSITA_GERENTE"
                objConfiguraLoja.iGerenteAutoriza = StrParaInt(objLojaConfig.sConteudo)
            
            Case "OPERADOR_IGUAL_VENDEDOR"
                objConfiguraLoja.iOperadorIgualVendedor = StrParaInt(objLojaConfig.sConteudo)
            
            Case "SIMBOLO_MOEDA"
                objConfiguraLoja.sSimboloMoeda = objLojaConfig.sConteudo
           
           Case "TABELA_PRECO_LOJA"
                objConfiguraLoja.iTabelaPreco = StrParaInt(objLojaConfig.sConteudo)
            
            Case "TELA_VENDA_M_P"
                objConfiguraLoja.iTelaVendaMP = IIf(objLojaConfig.sConteudo = "M", 0, 1)

            Case "TRUNCAMENTO_ARREDONDAMENTO"
                objConfiguraLoja.sTruncamentoArredondamento = objLojaConfig.sConteudo

            Case "VENDEDOR_OBRIGATORIO"
                objConfiguraLoja.iVendedorObrigatorio = StrParaInt(objLojaConfig.sConteudo)
        
            Case "ABRE_APOS_FECHAMENTO"
                objConfiguraLoja.iAbreAposFechamento = StrParaInt(objLojaConfig.sConteudo)
       
            Case "FTP_URL"
                objConfiguraLoja.sFTPURL = objLojaConfig.sConteudo
                
            Case "FTP_USERNAME"
                objConfiguraLoja.sFTPUserName = objLojaConfig.sConteudo
       
            Case "FTP_PASSWORD"
                objConfiguraLoja.sFTPPassword = objLojaConfig.sConteudo
       
            Case "FTP_DIRETORIO"
                objConfiguraLoja.sFTPDiretorio = objLojaConfig.sConteudo
       
            Case "INTERVALO_TRANSFERENCIA"
                objConfiguraLoja.lIntervaloTrans = StrParaLong(objLojaConfig.sConteudo)
       
       End Select

    Next

    ColLojaConfig_MoverCampos_ConfiguraLoja = SUCESSO

    Exit Function

Erro_ColLojaConfig_MoverCampos_ConfiguraLoja:

    ColLojaConfig_MoverCampos_ConfiguraLoja = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 151434)

    End Select

    Exit Function

End Function

Function AliquotaICMS_Le(ByVal colAliquotaICMS As Collection) As Long
'Le os Dados da AliquotaICMS no BD

Dim lErro As Long
Dim lComando As Long
Dim objAliquotaICMS As ClassAliquotaICMS
Dim sSigla As String
Dim dAliquota As Double
Dim iISS As Integer

On Error GoTo Erro_AliquotaICMS_Le

    'Abre Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 80083

    'Inicializa String
    sSigla = String(STRING_SIGLA, 0)

    lErro = Comando_Executar(lComando, "SELECT Sigla, Aliquota, ISS FROM AliquotasICMS WHERE FilialEmpresa = ?", sSigla, dAliquota, iISS, giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 80084

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 50125
''''    If lErro = AD_SQL_SEM_DADOS Then gError 80086

    'Enquanto existirem linhas, carregar no Grid
    Do While lErro = AD_SQL_SUCESSO
        
        Set objAliquotaICMS = New ClassAliquotaICMS

        objAliquotaICMS.sSigla = sSigla
        objAliquotaICMS.iISS = iISS
        objAliquotaICMS.dAliquota = dAliquota

        colAliquotaICMS.Add objAliquotaICMS

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 50127

    Loop

    Call Comando_Fechar(lComando)

    AliquotaICMS_Le = SUCESSO

    Exit Function

Erro_AliquotaICMS_Le:

    AliquotaICMS_Le = gErr

    Select Case gErr

        Case 80083
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 80084, 80085, 80087
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ALIQUOTAICMS", gErr)

        Case 80086 'A ser tratado na rotina chamadora.

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151435)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function


''Function Rede_Le(ByVal objRede As ClassRede) As Long
'''Lê os dados de rede apartir do código.
''
''Dim lErro As Long
''Dim lComando As Long
''Dim tRede As typeRede
''
''On Error GoTo Erro_Rede_Le
''
''    'Abre o comando
''    lComando = Comando_Abrir()
''    If lComando = 0 Then gError 80588
''
''    'Inicializa nome
''    tRede.sNome = String(STRING_REDE_NOME, 0)
''
''    'Lê os dados da tabela Rede
''    lErro = Comando_Executar(lComando, "SELECT Nome FROM Redes WHERE Codigo = ?", tRede.sNome, objRede.iCodigo)
''    If lErro <> AD_SQL_SUCESSO Then gError 80589
''
''    lErro = Comando_BuscarPrimeiro(lComando)
''    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 80590
''
''    If lErro = AD_SQL_SEM_DADOS Then gError 80591
''
''    'Carrega nome em objRede
''    objRede.sNome = tRede.sNome
''
''    lErro = Comando_Fechar(lComando)
''
''    Rede_Le = SUCESSO
''
''    Exit Function
''
''Erro_Rede_Le:
''
''    Rede_Le = gErr
''
''    Select Case gErr
''
''        Case 80588
''            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
''
''        Case 80589, 80590
''            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REDE", gErr, objRede.iCodigo)
''
''        Case 80591
''        'Erro tratado na rotina chamadora
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151436)
''
''    End Select
''
''    Call Comando_Fechar(lComando)
''
''    Exit Function
''
''End Function

Function TipoMeioPagtoLoja_Le(ByVal objTipoMeioPagtoLoja As ClassTMPLoja) As Long
'Lê os dados de meio de pagamento apartir do código.
'Devolve os dados em objTipoMeioPagtoLoja

Dim lErro As Long
Dim lComando As Long
Dim tTipoMeioPagtoLoja As typeTipoMeioPagtoLoja

On Error GoTo Erro_TipoMeioPagtoLoja_Le

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 80570
    
    'Inicializa
    tTipoMeioPagtoLoja.sDescricao = String(STRING_TIPOMEIOPAGTOLOJA_DESCRICAO, 0)

    'Lê os dados da tabela TipoMeioPagtoLoja
    lErro = Comando_Executar(lComando, "SELECT Inativo, Descricao, ExigeNumero, EnvioPagamento, PossuiAdm, ExigeRede  FROM TipoMeioPagtoLoja WHERE Tipo = ? ", tTipoMeioPagtoLoja.iInativo, tTipoMeioPagtoLoja.sDescricao, tTipoMeioPagtoLoja.iExigeNumero, tTipoMeioPagtoLoja.iEnvioPagamento, tTipoMeioPagtoLoja.iPossuiAdm, tTipoMeioPagtoLoja.iExigeRede, objTipoMeioPagtoLoja.iTipo)

    If lErro <> AD_SQL_SUCESSO Then gError 80571

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 80572

    If lErro = AD_SQL_SEM_DADOS Then gError 80573
    
    'Carrega os dados em objTipoMeioPagtoLoja
    objTipoMeioPagtoLoja.iInativo = tTipoMeioPagtoLoja.iInativo
    objTipoMeioPagtoLoja.sDescricao = tTipoMeioPagtoLoja.sDescricao
    objTipoMeioPagtoLoja.iExigeNumero = tTipoMeioPagtoLoja.iExigeNumero
    objTipoMeioPagtoLoja.iEnvioPagamento = tTipoMeioPagtoLoja.iEnvioPagamento
    objTipoMeioPagtoLoja.iPossuiAdm = tTipoMeioPagtoLoja.iInativo
    objTipoMeioPagtoLoja.iExigeRede = tTipoMeioPagtoLoja.iExigeRede
    
    lErro = Comando_Fechar(lComando)

    TipoMeioPagtoLoja_Le = SUCESSO

    Exit Function

Erro_TipoMeioPagtoLoja_Le:

    TipoMeioPagtoLoja_Le = gErr

    Select Case gErr

        Case 80570
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 80571, 80572
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOMEIOPAGTOLOJA", gErr)

        Case 80573
        'Erro tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151437)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function
'Maristela(Fim) - Tela: AdmMeioPagto

'Maristela(Inicio) - Tela: Gerente
Function Gerente_Le_Todos(ByVal colGerente As Collection) As Long
'Lê todos os gerentes da gifilialempresa e guarda em colGerente

Dim lErro As Long
Dim lComando As Long
Dim sCodUsuario As String
Dim objGerente As ClassGerente
Dim iCodigo As Integer

On Error GoTo Erro_Gerente_Le_Todos

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 81073

    sCodUsuario = String(STRING_USUARIO_CODIGO, 0)

    'Seleciona todos os gerentes da giFilialEmpresa
    lErro = Comando_Executar(lComando, "SELECT Codigo, CodUsuario FROM Gerente WHERE FilialEmpresa = ?", iCodigo, sCodUsuario, giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 81074

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81075
    If lErro = AD_SQL_SEM_DADOS Then gError 81076

    Do While lErro = AD_SQL_SUCESSO

        'Guarda na coleção os nomes dos usuarios que têm gerentes
        Set objGerente = New ClassGerente

        objGerente.sCodUsuario = sCodUsuario
        objGerente.iCodigo = iCodigo
       
        colGerente.Add objGerente
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81077

    Loop

    Call Comando_Fechar(lComando)

    Gerente_Le_Todos = SUCESSO

    Exit Function

Erro_Gerente_Le_Todos:

    Gerente_Le_Todos = gErr

    Select Case gErr

        Case 81073
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 81074, 81075, 81077
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_GERENTE2", gErr)

        Case 81076 'A ser tratado na rotina chamadora.

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151438)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Gerente_Le_Usuario(ByVal objGerente As ClassGerente) As Long
'Le o Gerente de um usuario

Dim lErro As Long
Dim lComando As Long
Dim iCodigo As Integer

On Error GoTo Erro_Gerente_Le_Usuario

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 81086

    'Seleciona o Codigo do Gerente do Usuario em objGerente
    lErro = Comando_Executar(lComando, "SELECT Codigo FROM Gerente WHERE CodUsuario = ? AND FilialEmpresa = ?", iCodigo, objGerente.sCodUsuario, giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 81087

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81088

    If lErro = AD_SQL_SEM_DADOS Then gError 81084

    
    objGerente.iCodigo = iCodigo

    Call Comando_Fechar(lComando)

    Gerente_Le_Usuario = SUCESSO

    Exit Function

Erro_Gerente_Le_Usuario:

    Gerente_Le_Usuario = gErr

    Select Case gErr

        Case 81086
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 81087, 81088
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_GERENTE1", gErr, objGerente.sCodUsuario)

        Case 81084 'Não existe na tabela de Gerentes

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151439)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Gerente_Le(ByVal objGerente As ClassGerente) As Long
'Le um gerente

Dim lErro As Long
Dim lComando As Long
Dim sCodUsuario As String

On Error GoTo Erro_Gerente_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 81093

    sCodUsuario = String(STRING_USUARIO_CODIGO, 0)

    'Seleciona os CodUsuario onde o Codigo for igual ao do objGerente e FilialEmpresa=giFilialEmpresa
    lErro = Comando_Executar(lComando, "SELECT CodUsuario FROM Gerente WHERE Codigo = ? and FilialEmpresa=?", sCodUsuario, objGerente.iCodigo, giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 81094

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81095

    If lErro = AD_SQL_SEM_DADOS Then gError 81096

    objGerente.sCodUsuario = sCodUsuario

    Call Comando_Fechar(lComando)

    Gerente_Le = SUCESSO

    Exit Function

Erro_Gerente_Le:

    Gerente_Le = gErr

    Select Case gErr

        Case 81093
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 81094, 81095
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_GERENTE", gErr, objGerente.iCodigo)

        Case 81096 'Não existe na tabela de Gerentes

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151440)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Gerente_Existe_Relacionamento(ByVal objGerente As ClassGerente) As Long

Dim lErro As Long
Dim lComando As Long
Dim iCodGerente As Integer
Dim iStatus As Integer
Dim iCaixa As Integer

On Error GoTo Erro_Gerente_Existe_Relacionamento

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 81134

    'Sessão
    lErro = Comando_Executar(lComando, "SELECT Caixa FROM Sessao WHERE (GerenteAbertura = ? OR GerenteFechamento = ?) and filialempresa =?", iCaixa, objGerente.iCodigo, objGerente.iCodigo, objGerente.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 81135

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81136

    'Existe Relacionamento
    If lErro = AD_SQL_SUCESSO Then gError 81137

    'Cupom Fiscal
    lErro = Comando_Executar(lComando, "SELECT Status FROM CupomFiscal WHERE (CodUsuarioCancel = ? AND FilialEmpresa = ?)", iStatus, objGerente.sCodUsuario, objGerente.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 81139

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81140

    'Existe Relacionamento
    If lErro = AD_SQL_SUCESSO Then gError 81141
    
    Call Comando_Fechar(lComando)

    Gerente_Existe_Relacionamento = SUCESSO

    Exit Function

Erro_Gerente_Existe_Relacionamento:

    Gerente_Existe_Relacionamento = gErr

    Select Case gErr

        Case 81134, 81135, 81139

        Case 81136, 81140

        Case 81137
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GERENTE_VINCULADO_SESSAO", gErr, objGerente.iCodigo)
            
        Case 81141
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GERENTE_VINCULADO_CUPOMFISCAL", gErr, objGerente.sCodUsuario)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151441)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function
'Maristela(Fim) - Tela: Gerente

Function Caixas_Le(ByVal objCaixa As ClassCaixa) As Long
'Lê os dados do Caixa, cujo código foi passado como parâmetro

Dim lErro As Long
Dim lComando As Long
Dim tCaixa As typeCaixa

On Error GoTo Erro_Caixas_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 79406

    'Inicializa as strings que receberão dados do BD
    tCaixa.sDescricao = String(STRING_CAIXA_DESCRICAO, 0)
    tCaixa.sNomeReduzido = String(STRING_CAIXA_NOME_REDUZIDO, 0)
    
    'Alterado por cyntia para incluir FilialEmpresa como parâmetro
    'Busca o Caixa no BD através do código passado como parâmetro
    lErro = Comando_Executar(lComando, "SELECT  NomeReduzido, Descricao, Status, DataInicial, ProxSeqMov, POS, TEF, BoletoManual, Ativo, Teclado, SoOrcamento, UltimoOperador, SessaoStatus FROM Caixa WHERE Codigo = ? AND FilialEmpresa=?", tCaixa.sNomeReduzido, tCaixa.sDescricao, tCaixa.iStatus, tCaixa.dtDataInicial, tCaixa.lProxSeqMov, tCaixa.iPos, tCaixa.iTEF, tCaixa.iBoletoManual, tCaixa.iAtivo, tCaixa.iTeclado, tCaixa.iOrcamentoECF, tCaixa.iUltimoOperador, tCaixa.iSessaoStatus, objCaixa.iCodigo, objCaixa.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 79403

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 79404

    'Se não encontrou o Caixa => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 79405

    'Transfere para objCaixa os dados obtidos do BD
    objCaixa.sNomeReduzido = tCaixa.sNomeReduzido
    objCaixa.sDescricao = tCaixa.sDescricao
    objCaixa.iStatus = tCaixa.iStatus
    objCaixa.dtDataInicial = tCaixa.dtDataInicial
    objCaixa.iTeclado = tCaixa.iTeclado
    objCaixa.lProxSeqMov = tCaixa.lProxSeqMov
    objCaixa.iPos = tCaixa.iPos
    objCaixa.iTEF = tCaixa.iTEF
    objCaixa.iBoletoManual = tCaixa.iBoletoManual
    objCaixa.iAtivo = tCaixa.iAtivo
    objCaixa.iOrcamentoECF = tCaixa.iOrcamentoECF
    objCaixa.iUltimoOperador = tCaixa.iUltimoOperador
    objCaixa.iSessaoStatus = tCaixa.iSessaoStatus
    
    'Fecha o comando
    Call Comando_Fechar(lComando)

    Caixas_Le = SUCESSO

    Exit Function

Erro_Caixas_Le:

    Caixas_Le = gErr

    Select Case gErr

        Case 79406
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 79403, 79404
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CAIXAS", gErr, objCaixa.iCodigo)

        Case 79405
            'Caixa não encontrado. Esse erro será tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151442)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function MovCaixa_Le_PrimeiroMovimentoCaixa(ByVal objCaixa As ClassCaixa, ByVal objMovimentosCaixa As ClassMovimentoCaixa) As Long
'Procura no BD se já existe um movimento para o caixa passado como parâmetro e retorna o código e data do primeiro movimento, caso este exista

Dim lErro As Long
Dim lComando As Long
Dim lNumMovto As Long
Dim dtDataMovimento As Date

On Error GoTo Erro_MovCaixa_Le_PrimeiroMovimentoCaixa

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 79457
    
    'Alterado por cyntia para incluir FilialEmpresa como parâmetro
    'Procura no BD se há algum movimento de caixa para o caixa passado como parâmetro
    lErro = Comando_Executar(lComando, "SELECT NumMovto, DataMovimento FROM MovimentosCaixa WHERE Caixa = ?  and FilialEmpresa = ? ORDER BY DataMovimento", lNumMovto, dtDataMovimento, objCaixa.iCodigo, objCaixa.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 79455
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 79456
    
    'Se encontrou
    If lErro <> AD_SQL_SEM_DADOS Then
        
        'Guarda em objMovimentosCaixa a data e o número do primeiro movto
        objMovimentosCaixa.dtDataMovimento = dtDataMovimento
        objMovimentosCaixa.lNumMovto = lNumMovto
    Else
        gError 79495
    End If
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    MovCaixa_Le_PrimeiroMovimentoCaixa = SUCESSO
    
    Exit Function
    
Erro_MovCaixa_Le_PrimeiroMovimentoCaixa:

    MovCaixa_Le_PrimeiroMovimentoCaixa = gErr
    
    Select Case gErr
    
        Case 79495
        
        Case 79457
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 79455, 79456
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOS_CAIXA", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151443)
    
    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function Caixa_Le_NomeReduzido(ByVal objCaixa As ClassCaixa) As Long
'Lê o codigo do Caixa a partir do Nome Reduzido passado como parâmetro

Dim lComando As Long
Dim lErro As Long
Dim tCaixa As typeCaixa

On Error GoTo Erro_Caixa_Le_NomeReduzido

    'Inicia o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 79579
    
    'Inicializa as variáveis que recberão strings
    tCaixa.sNomeReduzido = String(STRING_CAIXA_NOME_REDUZIDO, 0)
    tCaixa.sDescricao = String(STRING_CAIXA_DESCRICAO, 0)
    
    'Alterado por cyntia para incluir FilialEmpresa como parâmetro
    'Procura o Caixa a partir do Nome Reduzido
    lErro = Comando_Executar(lComando, "SELECT Codigo, NomeReduzido, Descricao, Status, DataInicial, ProxSeqMov , POS, TEF, BoletoManual, Ativo, UltimoOperador, SessaoStatus  FROM Caixa WHERE NomeReduzido =? AND FilialEmpresa=?", tCaixa.iCodigo, tCaixa.sNomeReduzido, tCaixa.sDescricao, tCaixa.iStatus, tCaixa.dtDataInicial, tCaixa.lProxSeqMov, tCaixa.iPos, tCaixa.iTEF, tCaixa.iBoletoManual, tCaixa.iAtivo, tCaixa.iUltimoOperador, tCaixa.iSessaoStatus, objCaixa.sNomeReduzido, giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 79580
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 79581
    
    'Se não encontrou => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 79582
    
    'Passa para objCaixa os dados encontrados
    objCaixa.iCodigo = tCaixa.iCodigo
    objCaixa.sNomeReduzido = tCaixa.sNomeReduzido
    objCaixa.sDescricao = tCaixa.sDescricao
    objCaixa.iStatus = tCaixa.iStatus
    objCaixa.dtDataInicial = tCaixa.dtDataInicial
    objCaixa.lProxSeqMov = tCaixa.lProxSeqMov
    objCaixa.iPos = tCaixa.iPos
    objCaixa.iTEF = tCaixa.iTEF
    objCaixa.iBoletoManual = tCaixa.iBoletoManual
    objCaixa.iAtivo = tCaixa.iAtivo
    objCaixa.iUltimoOperador = tCaixa.iUltimoOperador
    objCaixa.iSessaoStatus = tCaixa.iSessaoStatus
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Caixa_Le_NomeReduzido = SUCESSO
    
    Exit Function
    
Erro_Caixa_Le_NomeReduzido:

    Caixa_Le_NomeReduzido = gErr
    
    Select Case gErr
    
        Case 79579
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 79580, 79581
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CAIXAS2", gErr, objCaixa.sNomeReduzido)
        
        Case 79582
        'Caixa não encontrado. Esse erro deve ser tratado na rotina chamadora
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151444)
            
    End Select
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function
'Maristela(Fim) - Tela: Caixas

'Maristela(Inicio) - Tela: ECF
Function POS_Le(ByVal objPos As ClassPOS)
'Lê todos os dados do POS cujo código foi passado como parâmetro

Dim lComando As Long
Dim lErro As Long
Dim tPOS As typePOS

On Error GoTo Erro_POS_Le

    'Inicia o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 79587
    
    'Inicializa a variável que receberá o código do POS
    tPOS.sCodigo = String(STRING_POS_CODIGO, 0)
    
    'Procura o POS no BD a partir do código
    lErro = Comando_Executar(lComando, "SELECT Codigo, FilialEmpresa, Rede FROM POS WHERE Codigo = ?", tPOS.sCodigo, tPOS.iFilialEmpresa, tPOS.iRede, objPos.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 79588
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 79589
    
    'Se não encontrou o POS => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 79590
    
    'Passa para objPOS os valores lidos do BD
    objPos.sCodigo = tPOS.sCodigo
    objPos.iFilialEmpresa = tPOS.iFilialEmpresa
    objPos.iRede = tPOS.iRede
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    POS_Le = SUCESSO
    
    Exit Function
    
Erro_POS_Le:

    POS_Le = gErr
    
    Select Case gErr
    
        Case 79587
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 79588, 79589
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_POS", gErr, objPos.sCodigo)
            
        Case 79590 'POS não encontrado. Erro a ser tratado na rotina chamadora
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151445)
            
    End Select
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

'Modificação na Tela de Emissor de Cupom Fiscal (ECM)por Sergio Data 06/05/02
Function ECF_Le_Todos(ByVal colECF As Collection) As Long
'Lê o código de todos os ECF's da giFilialEmpresa e guarda em colECF

Dim lErro As Long
Dim lComando As Long
Dim iCodigo As Integer
Dim objECF As ClassECF
Dim tECF As typeECF

On Error GoTo Erro_ECF_Le_Todos

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 79520

    'Seleciona todos os ECF da giFilialEmpresa empresa Ordena por código
    lErro = Comando_Executar(lComando, "SELECT Codigo , Caixa, ImpressoraECF, ImpressoraCheque, HorarioVerao FROM ECF WHERE FilialEmpresa = ? ORDER BY Codigo", tECF.iCodigo, tECF.iCaixa, tECF.iImpressoraECF, tECF.iImpressoraCheque, tECF.iHorarioVerao, giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 79521

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 79522
    
    'Se não encontrou nenhum ECF => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 79524
    
    'Enquanto houverem Ecf's
    Do While lErro = AD_SQL_SUCESSO

        'Cria um novo objECF
        Set objECF = New ClassECF

        'Guarda no objECF o Código do ECF atual
        objECF.iCodigo = tECF.iCodigo
        
        'Armazena o objECF a FilialEmpresa que está sendo Utilizada
        objECF.iFilialEmpresa = giFilialEmpresa
        
        'Armazena no objECF o Caixa Vinculdado ao que está sendo Utilizada
        objECF.iCaixa = tECF.iCaixa
        
        'Armazena no objECF a ImpressoraECf  que está sendo Utilizada
        objECF.iImpressoraECF = tECF.iImpressoraECF
        
        'Armazena no objECF a ImpressoraCheque  que está sendo Utilizada
        objECF.iImpressoraCheque = tECF.iImpressoraCheque
        
        'Armazena no objECF a ImpressoraCheque  que está sendo Utilizada
        objECF.iHorarioVerao = tECF.iHorarioVerao
        
        'Guarda o objECF em colECF
        colECF.Add objECF

        'Busca o código do próximo ECF na Tabela
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 79523

    Loop

    'Fecha o comando
    Call Comando_Fechar(lComando)

    ECF_Le_Todos = SUCESSO

    Exit Function

Erro_ECF_Le_Todos:

    ECF_Le_Todos = gErr

    Select Case gErr

        Case 79520
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 79521, 79522, 79523
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ECF1", gErr)
        
        Case 79524
            'Erro Sem menssagem, mesmo senão encontrar  abrir a Tela
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151446)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ImpressorasECF_Le_Todos(ByVal colImpressorasECF As Collection) As Long
'Lê o código de todos os Fabricantes de ECF e guarda em colFabricanteECF

Dim lErro As Long
Dim lComando As Long
Dim objImpressorasECF As ClassFabricanteECF
Dim tImpressorasECF As TypeImpressoras

On Error GoTo Erro_ImpressorasECF_Le_Todos

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 79531

    'Inicializa a string que receberá o nome do Fabricante
    tImpressorasECF.sNome = String(STRING_IMPRESSORA_NOME, 0)
    
    'Seleciona as Impressoras
    lErro = Comando_Executar(lComando, "SELECT Codigo, Nome FROM ImpressorasECF", tImpressorasECF.iCodigo, tImpressorasECF.sNome)
    If lErro <> AD_SQL_SUCESSO Then gError 79532

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 79533
    
    'Se não encontrou nenhum Impressora = > erro
    If lErro = AD_SQL_SEM_DADOS Then gError 79534
    
    'Enquanto houverem Fabricantes
    Do While lErro = AD_SQL_SUCESSO

        'Cria um novo objFabricanteECF
        Set objImpressorasECF = New ClassFabricanteECF

        'Guarda em objImpressorasECF o Código e o Nome dImpressora atual
        objImpressorasECF.iCodigo = tImpressorasECF.iCodigo
        objImpressorasECF.sNome = tImpressorasECF.sNome

        'Guarda o objImpressora em colFabricanteECF
        colImpressorasECF.Add objImpressorasECF

        'Busca o próximo Fabricante
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 79535

    Loop

    'Fecha o comando
    Call Comando_Fechar(lComando)

    ImpressorasECF_Le_Todos = SUCESSO

    Exit Function

Erro_ImpressorasECF_Le_Todos:

    ImpressorasECF_Le_Todos = gErr

    Select Case gErr

        Case 79531
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 79532, 79533, 79535
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FABRICANTEECF1", gErr)
        
        Case 79534
            'Erro Tratado na Função Chamadora
        
        Case Else
        
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151447)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Caixa_Le_Todos(ByVal colCaixa As Collection) As Long
'Lê o código e o NOmeRed todos os caixas da giFilialEmpresa e guarda em colCaixas

Dim lErro As Long
Dim lComando As Long
Dim objCaixa As ClassCaixa
Dim tCaixa As typeCaixa

On Error GoTo Erro_Caixa_Le_Todos

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 79407

    'Inicializa a string que receberá o nome reduzido do Caixa e a Descrição e a POS
    tCaixa.sNomeReduzido = String(STRING_CAIXA_NOME_REDUZIDO, 0)
    tCaixa.sDescricao = String(STRING_CAIXA_DESCRICAO, 0)
    
    'Seleciona todos os caixas da giFilialEmpresa
    lErro = Comando_Executar(lComando, "SELECT Codigo, NomeReduzido, Descricao, Status, DataInicial, ProxSeqMov, POS, TEF, BoletoManual, Ativo FROM Caixa WHERE FilialEmpresa = ? ORDER BY Codigo", tCaixa.iCodigo, tCaixa.sNomeReduzido, tCaixa.sDescricao, tCaixa.iStatus, tCaixa.dtDataInicial, tCaixa.lProxSeqMov, tCaixa.iPos, tCaixa.iTEF, tCaixa.iBoletoManual, tCaixa.iAtivo, giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 79408

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 79409
    
    'Se não encontrou nenhum caixa => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 79525
    
    'Guarda na coleção os dados sobre os Caixas encontrados
    Do While lErro = AD_SQL_SUCESSO

        Set objCaixa = New ClassCaixa

        'Guarda em objCaixa os Registros Lidos no Banco de Dados
        objCaixa.sNomeReduzido = tCaixa.sNomeReduzido
        objCaixa.iCodigo = tCaixa.iCodigo
        objCaixa.iFilialEmpresa = giFilialEmpresa
        objCaixa.dtDataInicial = tCaixa.dtDataInicial
        objCaixa.iStatus = tCaixa.iStatus
        objCaixa.lProxSeqMov = tCaixa.lProxSeqMov
        objCaixa.sDescricao = tCaixa.sDescricao
        objCaixa.iPos = tCaixa.iPos
        objCaixa.iTEF = tCaixa.iTEF
        objCaixa.iBoletoManual = tCaixa.iBoletoManual
        objCaixa.iAtivo = tCaixa.iAtivo
        
        'Guarda objCaixa em colCaixa
        colCaixa.Add objCaixa

        'Busca o próximo registro que será armazenado em colCaixa
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 79411

    Loop

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Caixa_Le_Todos = SUCESSO

    Exit Function

Erro_Caixa_Le_Todos:

    Caixa_Le_Todos = gErr

    Select Case gErr

        Case 79407
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 79408, 79409, 79411
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CAIXAS1", gErr)
        
        Case 79525
            'Erro Tratado Dentro da Função Chamadora
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151448)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ECF_Le(ByVal objECF As ClassECF) As Long
'Lê os dados do ECF, cujo código foi passado como parâmetro

Dim lErro As Long
Dim lComando As Long
Dim tECF As typeECF

On Error GoTo Erro_ECF_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 79570
    
    If giFilialEmpresa = EMPRESA_TODA Then
    
        'Busca o ECF no BD através do código passado como parâmetro
        lErro = Comando_Executar(lComando, "SELECT ImpressoraECF, Caixa, ImpressoraCheque,   HorarioVerao, Ativo FROM ECF WHERE Codigo = ?  ", tECF.iImpressoraECF, tECF.iCaixa, tECF.iImpressoraCheque, tECF.iHorarioVerao, tECF.iAtivo, objECF.iCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 79571
    
    Else
        
        'Busca o ECF no BD através do código passado como parâmetro
        lErro = Comando_Executar(lComando, "SELECT ImpressoraECF, Caixa, ImpressoraCheque,   HorarioVerao, Ativo FROM ECF WHERE Codigo = ?  and FilialEmpresa = ? ", tECF.iImpressoraECF, tECF.iCaixa, tECF.iImpressoraCheque, tECF.iHorarioVerao, tECF.iAtivo, objECF.iCodigo, objECF.iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 79571
    
    End If
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 79572

    'Se não encontrou o ECF => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 79573

    'Transfere para objECF os dados obtidos do BD
    objECF.iImpressoraECF = tECF.iImpressoraECF
    objECF.iCaixa = tECF.iCaixa
    objECF.iImpressoraCheque = tECF.iImpressoraCheque
    objECF.iHorarioVerao = tECF.iHorarioVerao
    objECF.iAtivo = tECF.iAtivo
    
    'Fecha o comando
    Call Comando_Fechar(lComando)

    ECF_Le = SUCESSO

    Exit Function

Erro_ECF_Le:

    ECF_Le = gErr

    Select Case gErr

        Case 79570
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 79571, 79572
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ECF", gErr, objECF.iCodigo)

        Case 79573
            'ECF não encontrado. Esse erro será tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151449)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function



Function FabricanteECF_Le_Todos(ByVal colFabricanteECF As Collection) As Long
'Lê o código de todos os Fabricantes de ECF e guarda em colFabricanteECF

Dim lErro As Long
Dim lComando As Long
Dim iCodigo As Integer
Dim sNome As String
Dim objFabricanteECF As ClassFabricanteECF

On Error GoTo Erro_FabricanteECF_Le_Todos

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 79531

    'Inicializa a string que receberá o nome do Fabricante
    sNome = String(STRING_FABRICANTE_NOME, 0)
    
    'Seleciona todos os Fabricantes da giFilialEmpresa
    lErro = Comando_Executar(lComando, "SELECT Codigo, Nome FROM FabricanteECF", iCodigo, sNome)
    If lErro <> AD_SQL_SUCESSO Then gError 79532

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 79533
    
    'Se não encontrou nenhum Fabricante = > erro
    If lErro = AD_SQL_SEM_DADOS Then gError 79534
    
    'Enquanto houverem Fabricantes
    Do While lErro = AD_SQL_SUCESSO

        'Cria um novo objFabricanteECF
        Set objFabricanteECF = New ClassFabricanteECF

        'Guarda em objFabricanteECF o Código e o Nome do Fabricante atual
        objFabricanteECF.iCodigo = iCodigo
        objFabricanteECF.sNome = sNome

        'Guarda o objFabricanteECF em colFabricanteECF
        colFabricanteECF.Add objFabricanteECF

        'Busca o próximo Fabricante
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 79535

    Loop

    'Fecha o comando
    Call Comando_Fechar(lComando)

    FabricanteECF_Le_Todos = SUCESSO

    Exit Function

Erro_FabricanteECF_Le_Todos:

    FabricanteECF_Le_Todos = gErr

    Select Case gErr

        Case 79531
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        '???? O código de quem vc vai estar passando para essa msg? *** Ok - Luiz
        Case 79532, 79533, 79535
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FABRICANTEECF1", gErr)
        
        Case 79534

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151450)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'Inicio Tela de Operador Dia 7/06/02 por Sergio Ricardo P da Vitoria
Function Operador_Le_Todos(ByVal colOperador As Collection) As Long
'Lê todos os operadores da gifilialempresa e guarda em colOperador

Dim lErro As Long
Dim lComando As Long
Dim objOperador As ClassOperador
Dim tOperador As typeOperador

On Error GoTo Erro_Operador_Le_Todos

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 81002
    
    'Inicilaização do Buffer
    tOperador.sSenha = String(STRING_OPERADOR_SENHA, 0)
    tOperador.sNome = String(STRING_OPERADOR_NOME, 0)
    
    'Seleciona todos os operadores da giFilialEmpresa
    lErro = Comando_Executar(lComando, "SELECT Codigo, Nome FROM Operador WHERE FilialEmpresa = ?", tOperador.iCodigo, tOperador.sNome, giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 81003

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81004

    If lErro = AD_SQL_SEM_DADOS Then gError 81005

    Do While lErro = AD_SQL_SUCESSO

        'Guarda na coleção os nomes dos usuarios que têm operadores
        Set objOperador = New ClassOperador

        objOperador.iCodigo = tOperador.iCodigo
        objOperador.sNome = tOperador.sNome
        
        'Adcionar na Coleção de Operador
        colOperador.Add objOperador

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81006

    Loop

    Call Comando_Fechar(lComando)

    Operador_Le_Todos = SUCESSO

    Exit Function

Erro_Operador_Le_Todos:

    Operador_Le_Todos = gErr

    Select Case gErr

        Case 81002
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 81003, 81004, 81006
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_OPERADOR2", gErr)

        Case 81005
            'Não Emite Menssagem de erro só desvia o código
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151451)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function


Function Vendedor_Le_Todos(colVendedor As Collection) As Long
'Lê todos os Vendedores Ativos da Empresa

Dim lErro As Long
Dim lComando As Long
Dim tVendedor As typeVendedor
Dim objVendedor As ClassVendedor

On Error GoTo Erro_Vendedor_Le_Todos

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 104284
    
    'Inicialização do Buffer
    tVendedor.sNomeReduzido = String(STRING_VENDEDOR_NOME_RED, 0)
    
    'Seleciona todos os Vendedores Ativos
    lErro = Comando_Executar(lComando, "SELECT Codigo, NomeReduzido " & _
                                       "FROM Vendedores ", _
                                       tVendedor.iCodigo, tVendedor.sNomeReduzido)
    If lErro <> AD_SQL_SUCESSO Then gError 104285

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 104286

    If lErro = AD_SQL_SEM_DADOS Then gError 104287
    
    Do While lErro = AD_SQL_SUCESSO

        'Guarda na coleção os nomes dos usuarios que têm operadores
        Set objVendedor = New ClassVendedor

        objVendedor.iCodigo = tVendedor.iCodigo
        objVendedor.sNomeReduzido = tVendedor.sNomeReduzido
        
        colVendedor.Add objVendedor

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 104314

    Loop

    Call Comando_Fechar(lComando)

    Vendedor_Le_Todos = SUCESSO

    Exit Function

Erro_Vendedor_Le_Todos:

    Vendedor_Le_Todos = gErr

    Select Case gErr

        Case 104284
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 104285, 104286, 104314
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_VENDEDORES", gErr)

        Case 104287
            'Só Desvia o Código não Emite Msg de Erro
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151452)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function VendedorFilial_Le_Todos(ByVal colVendedor As Collection) As Long
'Lê todos os Vendedores Ativos da Filial

Dim lErro As Long
Dim lComando As Long
Dim tVendedor As typeVendedor
Dim objVendedor As ClassVendedor

On Error GoTo Erro_VendedorFilial_Le_Todos

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 109487
    
    'Inicialização do Buffer
    tVendedor.sNomeReduzido = String(STRING_VENDEDOR_NOME_RED, 0)
    
    'Seleciona todos os Vendedores Ativos da FilialEmpresa
    lErro = Comando_Executar(lComando, "SELECT Codigo, NomeReduzido " & _
                                       "FROM Vendedores, VendedorLoja " & _
                                       "WHERE Vendedores.Codigo=VendedorLoja.Vendedor " & _
                                       "AND VendedorLoja.FilialEmpresa=?", _
                                       tVendedor.iCodigo, tVendedor.sNomeReduzido, _
                                       giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 109488

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 109489

    If lErro = AD_SQL_SEM_DADOS Then gError 109490
    
    Do While lErro = AD_SQL_SUCESSO

        'Guarda na coleção os nomes dos usuarios que têm operadores
        Set objVendedor = New ClassVendedor

        objVendedor.iCodigo = tVendedor.iCodigo
        objVendedor.sNomeReduzido = tVendedor.sNomeReduzido
        
        colVendedor.Add objVendedor

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 109491

    Loop

    Call Comando_Fechar(lComando)

    VendedorFilial_Le_Todos = SUCESSO

    Exit Function

Erro_VendedorFilial_Le_Todos:

    VendedorFilial_Le_Todos = gErr

    Select Case gErr

        Case 109487
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 109488, 109489, 109491
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_VENDEDORES", gErr)

        Case 109490
            'Só Desvia o Código não Emite Msg de Erro
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151453)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Operador_Le(ByVal objOperador As ClassOperador) As Long
'Le um operador

Dim lErro As Long
Dim lComando As Long
Dim tOperador As typeOperador
Dim iAtivo As Integer

On Error GoTo Erro_Operador_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 81023
    
    'Inicilaização do Buffer
    tOperador.sSenha = String(STRING_OPERADOR_SENHA, 0)
    tOperador.sNome = String(STRING_OPERADOR_NOME, 0)
    
    'Seleciona todos os operadores da giFilialEmpresa
    lErro = Comando_Executar(lComando, "SELECT  Nome , Senha , Desconto , LimiteDesconto , CodVendedor, Gerente, Ativo FROM Operador WHERE Codigo = ? AND FilialEmpresa = ?", tOperador.sNome, tOperador.sSenha, tOperador.iDesconto, tOperador.iLimiteDesconto, tOperador.iCodigoVendedor, tOperador.iGerente, iAtivo, objOperador.iCodigo, objOperador.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 81024

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81025

    If lErro = AD_SQL_SEM_DADOS Then gError 81026

    objOperador.sNome = tOperador.sNome
    objOperador.iDesconto = tOperador.iDesconto
    objOperador.iLimiteDesconto = tOperador.iLimiteDesconto
    objOperador.iGerente = tOperador.iGerente
    objOperador.sSenha = tOperador.sSenha
    objOperador.iCodigoVendedor = tOperador.iCodigoVendedor
    objOperador.iAtivo = iAtivo
    
    Call Comando_Fechar(lComando)

    Operador_Le = SUCESSO

    Exit Function

Erro_Operador_Le:

    Operador_Le = gErr

        Select Case gErr

        Case 81023
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 81024, 81025
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_OPERADOR", gErr, objOperador.iCodigo)

        Case 81026 'Não existe na tabela de Operadores

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151454)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function












'Maristela(Fim) - Tela: ECT

''Maristela(Inicio) - Tela: Operador
'Function Operador_Le_Todos(ByVal colOperador As Collection) As Long
''Lê todos os operadores da gifilialempresa e guarda em colOperador
'
'Dim lErro As Long
'Dim lComando As Long
'Dim sCodUsuario As String
'Dim objOperador As ClassOperador
'Dim iCodigo As Integer
'Dim iDesconto As Integer
'Dim iLimiteDesconto As Integer
'
'On Error GoTo Erro_Operador_Le_Todos
'
'    lComando = Comando_Abrir()
'    If lComando = 0 Then gError 81002
'
'    sCodUsuario = String(STRING_USUARIO_CODIGO, 0)
'
'    'Seleciona todos os operadores da giFilialEmpresa
'    lErro = Comando_Executar(lComando, "SELECT Codigo, Desconto, LimiteDesconto ,CodUsuario FROM Operador WHERE FilialEmpresa = ?", iCodigo, iDesconto, iLimiteDesconto, sCodUsuario, giFilialEmpresa)
'    If lErro <> AD_SQL_SUCESSO Then gError 81003
'
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81004
'
'    If lErro = AD_SQL_SEM_DADOS Then gError 81005
'
'    Do While lErro = AD_SQL_SUCESSO
'
'        'Guarda na coleção os nomes dos usuarios que têm operadores
'        Set objOperador = New ClassOperador
'
'        objOperador.sCodUsuario = sCodUsuario
'        objOperador.iCodigo = iCodigo
'        objOperador.iDesconto = iDesconto
'        objOperador.iLimiteDesconto = iLimiteDesconto
'
'        colOperador.Add objOperador
'
'        lErro = Comando_BuscarProximo(lComando)
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81006
'
'    Loop
'
'    Call Comando_Fechar(lComando)
'
'    Operador_Le_Todos = SUCESSO
'
'    Exit Function
'
'Erro_Operador_Le_Todos:
'
'    Operador_Le_Todos = gErr
'
'    Select Case gErr
'
'        Case 81002
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 81003, 81004, 81006
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_OPERADOR2", gErr)
'
'        Case 81005 'A ser tratado na rotina chamadora.
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151455)
'
'    End Select
'
'    Call Comando_Fechar(lComando)
'
'    Exit Function
'
'End Function

'Function Operador_Le_Usuario(ByVal objOperador As ClassOperador) As Long
''Le o operador de um usuario
'
'Dim lErro As Long
'Dim lComando As Long
'Dim iCodigo As Integer
'Dim iDesconto As Integer
'Dim iLimiteDesconto As Integer
'
'On Error GoTo Erro_Operador_Le_Usuario
'
'    lComando = Comando_Abrir()
'    If lComando = 0 Then gError 81010
'
'    'Seleciona o Codigo do Operador, desconto e limite do desconto do Usuario em objOperador
'    lErro = Comando_Executar(lComando, "SELECT Codigo, Desconto,LimiteDesconto FROM Operador WHERE CodUsuario = ? AND FilialEmpresa = ?", iCodigo, iDesconto, iLimiteDesconto, objOperador.sCodUsuario, giFilialEmpresa)
'    If lErro <> AD_SQL_SUCESSO Then gError 81011
'
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81012
'    If lErro = AD_SQL_SEM_DADOS Then gError 81013
'
'    objOperador.iDesconto = iDesconto
'    objOperador.iCodigo = iCodigo
'    objOperador.iLimiteDesconto = iLimiteDesconto
'
'    Call Comando_Fechar(lComando)
'
'    Operador_Le_Usuario = SUCESSO
'
'    Exit Function
'
'Erro_Operador_Le_Usuario:
'
'    Operador_Le_Usuario = gErr
'
'    Select Case gErr
'
'        Case 81010
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 81011, 81012
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_OPERADOR1", gErr, objOperador.sCodUsuario)
'
'        Case 81013 'Não existe na tabela de Operadores
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151456)
'
'    End Select
'
'    Call Comando_Fechar(lComando)
'
'    Exit Function
'
'End Function
'
'Function Operador_Le(ByVal objOperador As ClassOperador) As Long
''Le um operador
'
'Dim lErro As Long
'Dim lComando As Long
'Dim sCodUsuario As String
'Dim iDesconto As Integer
'Dim iLimiteDesconto As Integer
'
'On Error GoTo Erro_Operador_Le
'
'    lComando = Comando_Abrir()
'    If lComando = 0 Then gError 81023
'
'    sCodUsuario = String(STRING_USUARIO_CODIGO, 0)
'
'    'Seleciona os CodUsuario,Desconto, Lmite de Desconto onde o Codigo for igual ao do objOperador e FilialEmpresa=giFilialEmpresa
'    lErro = Comando_Executar(lComando, "SELECT CodUsuario, Desconto , LimiteDesconto FROM Operador WHERE Codigo = ? and FilialEmpresa=?", sCodUsuario, iDesconto, iLimiteDesconto, objOperador.iCodigo, giFilialEmpresa)
'    If lErro <> AD_SQL_SUCESSO Then gError 81024
'
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81025
'
'    If lErro = AD_SQL_SEM_DADOS Then gError 81026
'
'    objOperador.sCodUsuario = sCodUsuario
'    objOperador.iDesconto = iDesconto
'    objOperador.iLimiteDesconto = iLimiteDesconto
'
'    Call Comando_Fechar(lComando)
'
'    Operador_Le = SUCESSO
'
'    Exit Function
'
'Erro_Operador_Le:
'
'    Operador_Le = gErr
'
'    Select Case gErr
'
'        Case 81023
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 81024, 81025
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_OPERADOR", gErr, objOperador.iCodigo)
'
'        Case 81026 'Não existe na tabela de Operadores
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151457)
'
'    End Select
'
'    Call Comando_Fechar(lComando)
'
'    Exit Function
'
'End Function
'
'Function Operador_Existe_Relacionamento(ByVal objOperador As ClassOperador) As Long
''Verifica se existe relacionamento com a tabela Boleto
'
'Dim lErro As Long
'Dim lComando As Long
'Dim iCodOperador As Integer
'
'On Error GoTo Erro_Operador_Existe_Relacionamento
'
'    lComando = Comando_Abrir()
'    If lComando = 0 Then gError 81053
'
'    'Boleto
'    lErro = Comando_Executar(lComando, "SELECT Operador FROM Boleto WHERE Operador = ?", iCodOperador, objOperador.iCodigo)
'    If lErro <> AD_SQL_SUCESSO Then gError 81054
'
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81055
'
'    'Existe Relacionamento
'    If lErro = AD_SQL_SUCESSO Then gError 81056
'
'    'Valeticket
'    lErro = Comando_Executar(lComando, "SELECT Operador FROM ValeTicket WHERE Operador= ?", iCodOperador, objOperador.iCodigo)
'    If lErro <> AD_SQL_SUCESSO Then gError 81068
'
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81069
'
'    'Existe Relacionamento
'    If lErro = AD_SQL_SUCESSO Then gError 81070
'
'    Call Comando_Fechar(lComando)
'
'    Operador_Existe_Relacionamento = SUCESSO
'
'    Exit Function
'
'Erro_Operador_Existe_Relacionamento:
'
'    Operador_Existe_Relacionamento = gErr
'
'    Select Case gErr
'
'        Case 81053, 81054, 81068
'
'        Case 81055, 81069
'
'        Case 81056
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_OPERADOR_VINCULADO_BOLETO", gErr, objOperador.iCodigo)
'
'        Case 81070
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_OPERADOR_VINCULADO_VALETICKET", gErr, objOperador.iCodigo)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151458)
'
'    End Select
'
'    Call Comando_Fechar(lComando)
'
'    Exit Function
'
'End Function
''Maristela(Fim) - Tela: Operador

'Maristela(Inicio) - Tela: POS
Function POS_Le_Todos(ByVal colPOS As Collection) As Long
'Lê em colPOS os terminais cadastrados

Dim lErro As Long
Dim lComando As Long
Dim objPos As ClassPOS
Dim sCodigo As String

On Error GoTo Erro_POS_Le_Todos

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 80682
    
    'Inicializa código
    sCodigo = String(STRING_POS_CODIGO, 0)
    
    'Seleciona código da tabela POS
    lErro = Comando_Executar(lComando, "SELECT Codigo FROM POS WHERE FilialEmpresa = ? ", sCodigo, giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 80683
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 80684

    Do While lErro = AD_SQL_SUCESSO
        
        'Inicializa POS
        Set objPos = New ClassPOS

        'Carrega os dados em objPOS
        objPos.sCodigo = sCodigo
        
        colPOS.Add objPos

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 80685

     Loop

    'Fecha o comando
    Call Comando_Fechar(lComando)

    POS_Le_Todos = SUCESSO

    Exit Function

Erro_POS_Le_Todos:

   POS_Le_Todos = gErr

    Select Case gErr

        Case 80682
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 80683, 80684, 80685
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_POS1", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151459)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function
'Maristela(Fim) - Tela: POS


Function AliquotaICMS_Le_Todas(ByVal colAliquotaICMS As Collection) As Long
'Lê todos os tipos de Aliquotas ICMS e guarda em colAliquotaICMS

Dim lErro As Long
Dim objAliquotaICMS As ClassAliquotaICMS
Dim lComando As Long
Dim dAliquota As Double
Dim sSigla As String
Dim iISS As Integer

On Error GoTo Erro_AliquotaICMS_Le_Todas

    'abre o comando
    lComando = Comando_Abrir
    If lComando = 0 Then gError 98389
    
    sSigla = String(STRING_SIGLA, 0)
    
    lErro = Comando_Executar(lComando, "SELECT Sigla, Aliquota, ISS FROM AliquotasICMS WHERE FilialEmpresa=?", sSigla, dAliquota, iISS, giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 98390

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98391

    Do While lErro = AD_SQL_SUCESSO
    
        'Guarda na coleção a sigla e a aliquota
        Set objAliquotaICMS = New ClassAliquotaICMS

        objAliquotaICMS.sSigla = sSigla
        objAliquotaICMS.dAliquota = dAliquota
        objAliquotaICMS.iISS = iISS
    
        colAliquotaICMS.Add objAliquotaICMS

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98392

    Loop

    Call Comando_Fechar(lComando)

    AliquotaICMS_Le_Todas = SUCESSO

    Exit Function

Erro_AliquotaICMS_Le_Todas:

    AliquotaICMS_Le_Todas = gErr
    
    Select Case gErr
    
        Case 98389
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 98390, 98391, 98392
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ALIQUOTAICMS", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151460)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function


Function AdmMeioPagto_Le_Todas(ByVal colAdmMeioPagto As Collection) As Long
'Lê em colAdmMeioPagto as administradoras de meios de pagamentos cadastrados

Dim lErro As Long
Dim lComando As Long
Dim objAdmMeioPagto As ClassAdmMeioPagto
Dim tAdmMeioPagto As typeAdmMeioPagto

On Error GoTo Erro_AdmMeioPagto_Le_Todas

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 104029
    
    'Inicializa o buffer nome
    tAdmMeioPagto.sNome = String(STRING_ADMMEIOPAGTO_NOME, 0)

    'Seleciona código e nome dos meios de pagamentos da tabela AdmMeioPagto
    lErro = Comando_Executar(lComando, "SELECT Codigo, Nome, TipoMeioPagto FROM AdmMeioPagto Order by Codigo ", tAdmMeioPagto.iCodigo, tAdmMeioPagto.sNome, tAdmMeioPagto.iTipoMeioPagto)
    If lErro <> AD_SQL_SUCESSO Then gError 104030

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 104031

    Do While lErro = AD_SQL_SUCESSO
        
        'Inicializa  o obj relacionado a administradora
        Set objAdmMeioPagto = New ClassAdmMeioPagto
        
        objAdmMeioPagto.iCodigo = tAdmMeioPagto.iCodigo
        
        objAdmMeioPagto.sNome = tAdmMeioPagto.sNome
        objAdmMeioPagto.iTipoMeioPagto = tAdmMeioPagto.iTipoMeioPagto
        
        'Adcionar na Coleção para preenchimento da ListBox Admnistradoras
        colAdmMeioPagto.Add objAdmMeioPagto

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 104032

     Loop

    'Fecha o comando
    Call Comando_Fechar(lComando)

    AdmMeioPagto_Le_Todas = SUCESSO

    Exit Function

Erro_AdmMeioPagto_Le_Todas:

    AdmMeioPagto_Le_Todas = gErr

    Select Case gErr

        Case 104029
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 104030, 104031, 104032
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ADMMEIOPAGTO1", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151461)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function AdmMeioPagto_Le(ByVal objAdmMeioPagto As ClassAdmMeioPagto) As Long
'Lê AdmMeioPagto no Banco de Dados

Dim lErro As Long
Dim lComando As Long
Dim tAdmMeioPagto As typeAdmMeioPagto


On Error GoTo Erro_AdmMeioPagto_Le

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 104075
    
    'Inicializa código e nome
    tAdmMeioPagto.sNome = String(STRING_ADMMEIOPAGTO_NOME, 0)
    tAdmMeioPagto.sAgencia = String(STRING_ADMMEIOPAGTO_AGENCIA, 0)
    
    'Seleciona código e nome dos meios de pagamentos da tabela AdmMeioPagto
    lErro = Comando_Executar(lComando, "SELECT Nome, Rede, TaxaVista, TaxaParcelado, TipoMeioPagto, DefasagemPagtoVista, ContaCorrenteInterna, GeraTituloRec, Endereco, CodBanco, Agencia, Ativo, DataLog FROM AdmMeioPagto  WHERE  Codigo = ?  AND FilialEmpresa = ? ", tAdmMeioPagto.sNome, tAdmMeioPagto.iRede, tAdmMeioPagto.dTaxaVista, tAdmMeioPagto.dTaxaParcelado, tAdmMeioPagto.iTipoMeioPagto, tAdmMeioPagto.iDefasagemPagtoVista, tAdmMeioPagto.iContaCorrenteInterna, tAdmMeioPagto.iGeraTituloRec, tAdmMeioPagto.lEndereco, tAdmMeioPagto.iCodBanco, tAdmMeioPagto.sAgencia, tAdmMeioPagto.iAtivo, tAdmMeioPagto.dtDataLog, objAdmMeioPagto.iCodigo, objAdmMeioPagto.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 104073
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 104074
    
    'Se não encontrar Dados
    If lErro = AD_SQL_SEM_DADOS Then gError 104017
   
   'Encontrando Preencher objAdmMeioPagto
    objAdmMeioPagto.dTaxaParcelado = tAdmMeioPagto.dTaxaParcelado
    objAdmMeioPagto.sNome = tAdmMeioPagto.sNome
    objAdmMeioPagto.dTaxaVista = tAdmMeioPagto.dTaxaVista
    objAdmMeioPagto.iCodBanco = tAdmMeioPagto.iCodBanco
    objAdmMeioPagto.iContaCorrenteInterna = tAdmMeioPagto.iContaCorrenteInterna
    objAdmMeioPagto.iDefasagemPagtoVista = tAdmMeioPagto.iDefasagemPagtoVista
    objAdmMeioPagto.iGeraTituloRec = tAdmMeioPagto.iGeraTituloRec
    objAdmMeioPagto.iRede = tAdmMeioPagto.iRede
    objAdmMeioPagto.iTipoMeioPagto = tAdmMeioPagto.iTipoMeioPagto
    objAdmMeioPagto.lEndereco = tAdmMeioPagto.lEndereco
    objAdmMeioPagto.sAgencia = tAdmMeioPagto.sAgencia
    objAdmMeioPagto.iAtivo = tAdmMeioPagto.iAtivo
    objAdmMeioPagto.dtDataLog = tAdmMeioPagto.dtDataLog
    
    'Fecha o comando
    Call Comando_Fechar(lComando)

    AdmMeioPagto_Le = SUCESSO

    Exit Function

Erro_AdmMeioPagto_Le:

    AdmMeioPagto_Le = gErr

    Select Case gErr

        Case 104075
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 104073, 104074
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ADMMEIOPAGTO1", gErr)
    
        Case 104017

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151462)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function TipoMeioPagto_Le_Todas(ByVal colTipoMeioPagto As Collection) As Long
'Função que Lê todas os Tipos de Pagamentos
Dim lErro As Long
Dim lComando As Long
Dim tTipoMeioPagtoLoja As typeTipoMeioPagtoLoja
Dim objTipoMeioPagtoLoja As ClassTMPLoja

On Error GoTo Erro_TipoMeioPagto_Le_Todas

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 104029
    
    'Inicializa código e nome
    tTipoMeioPagtoLoja.sDescricao = String(STRING_TIPOMEIOPAGTOLOJA_DESCRICAO, 0)

    'Seleciona código e nome dos meios de pagamentos da tabela AdmMeioPagto
    lErro = Comando_Executar(lComando, "SELECT Tipo, Descricao, EnvioPagamento, Transferencia, PossuiAdm FROM TipoMeioPagtoLoja ORDER BY Tipo", tTipoMeioPagtoLoja.iTipo, tTipoMeioPagtoLoja.sDescricao, tTipoMeioPagtoLoja.iEnvioPagamento, tTipoMeioPagtoLoja.iTransferencia, tTipoMeioPagtoLoja.iPossuiAdm)
    If lErro <> AD_SQL_SUCESSO Then gError 104035
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 104036

    Do While lErro = AD_SQL_SUCESSO
        
        'Inicializa objTipoMeioPagtoLoja
        Set objTipoMeioPagtoLoja = New ClassTMPLoja
        
        objTipoMeioPagtoLoja.iTipo = tTipoMeioPagtoLoja.iTipo
        objTipoMeioPagtoLoja.sDescricao = tTipoMeioPagtoLoja.sDescricao
        objTipoMeioPagtoLoja.iEnvioPagamento = tTipoMeioPagtoLoja.iEnvioPagamento
        objTipoMeioPagtoLoja.iTransferencia = tTipoMeioPagtoLoja.iTransferencia
        objTipoMeioPagtoLoja.iPossuiAdm = tTipoMeioPagtoLoja.iPossuiAdm
        
        'Adcionar na Coleção
        colTipoMeioPagto.Add objTipoMeioPagtoLoja
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 104037

     Loop

    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    TipoMeioPagto_Le_Todas = SUCESSO
    
    Exit Function

Erro_TipoMeioPagto_Le_Todas:

    TipoMeioPagto_Le_Todas = gErr
    
   Select Case gErr

        Case 104029
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 104035, 104036, 104037
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SELECAO_TIPOMEIOPAGTOLOJA", gErr)
        
         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151463)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function AdmMeioPagtoCondPagto_Le(ByVal objAdmMeioPagto As ClassAdmMeioPagto) As Long
'Lê em colCodPagtoLoja as Codiçoes de Pagamento relacionados as Admnistradoras
Dim lErro As Long
Dim lComando As Long
Dim tAdmMeioPagtoCondPagto As typeAdmMeioPagtoCondPagto
Dim objAdmMeioPagtoCondPagto As ClassAdmMeioPagtoCondPagto

On Error GoTo Erro_AdmMeioPagtoCondPagto_Le

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 104073
    
    'Inicializa  nome do Parcelamento
    tAdmMeioPagtoCondPagto.sNomeParcelamento = String(STRING_NOME_PARCELAMENTO, 0)

    'Seleciona código e nome dos meios de pagamentos da tabela AdmMeioPagto
    lErro = Comando_Executar(lComando, "SELECT Parcelamento, NomeParcelamento, NumParcelas , ParcelasRecebto , Taxa , ValorMinimo , Desconto , JurosParcelamento , Juros, Saldo, Ativo, DataLog, PreDatado FROM AdmMeioPagtoCondPagto where AdmMeioPagto = ? AND FilialEmpresa = ? ", _
    tAdmMeioPagtoCondPagto.iParcelamento, tAdmMeioPagtoCondPagto.sNomeParcelamento, tAdmMeioPagtoCondPagto.iNumParcelas, tAdmMeioPagtoCondPagto.iParcelasRecebto, tAdmMeioPagtoCondPagto.dTaxa, tAdmMeioPagtoCondPagto.dValorMinimo, tAdmMeioPagtoCondPagto.dDesconto, tAdmMeioPagtoCondPagto.iJurosParcelamento, _
    tAdmMeioPagtoCondPagto.dJuros, tAdmMeioPagtoCondPagto.dSaldo, tAdmMeioPagtoCondPagto.iAtivo, tAdmMeioPagtoCondPagto.dtDataLog, tAdmMeioPagtoCondPagto.iPreDatado, objAdmMeioPagto.iCodigo, objAdmMeioPagto.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 104070

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 104071
    
    If lErro = AD_SQL_SEM_DADOS Then gError 104086
    
    Do While lErro = AD_SQL_SUCESSO
        
        'Inicializa o objAdmMeioPagtoCondPagto
        Set objAdmMeioPagtoCondPagto = New ClassAdmMeioPagtoCondPagto
        
        objAdmMeioPagtoCondPagto.iAdmMeioPagto = objAdmMeioPagto.iCodigo
        objAdmMeioPagtoCondPagto.dDesconto = tAdmMeioPagtoCondPagto.dDesconto
        objAdmMeioPagtoCondPagto.dJuros = tAdmMeioPagtoCondPagto.dJuros
        objAdmMeioPagtoCondPagto.dTaxa = tAdmMeioPagtoCondPagto.dTaxa
        objAdmMeioPagtoCondPagto.dValorMinimo = tAdmMeioPagtoCondPagto.dValorMinimo
        objAdmMeioPagtoCondPagto.iFilialEmpresa = objAdmMeioPagto.iFilialEmpresa
        objAdmMeioPagtoCondPagto.iJurosParcelamento = tAdmMeioPagtoCondPagto.iJurosParcelamento
        objAdmMeioPagtoCondPagto.iNumParcelas = tAdmMeioPagtoCondPagto.iNumParcelas
        objAdmMeioPagtoCondPagto.iParcelasRecebto = tAdmMeioPagtoCondPagto.iParcelasRecebto
        objAdmMeioPagtoCondPagto.sNomeParcelamento = tAdmMeioPagtoCondPagto.sNomeParcelamento
        objAdmMeioPagtoCondPagto.iParcelamento = tAdmMeioPagtoCondPagto.iParcelamento
        objAdmMeioPagtoCondPagto.dSaldo = tAdmMeioPagtoCondPagto.dSaldo
        objAdmMeioPagtoCondPagto.iAtivo = tAdmMeioPagtoCondPagto.iAtivo
        objAdmMeioPagtoCondPagto.dtDataLog = tAdmMeioPagtoCondPagto.dtDataLog
        objAdmMeioPagtoCondPagto.iPreDatado = tAdmMeioPagtoCondPagto.iPreDatado
        
        'Adcionar na Coleção de colCondPagtoLoja as Informações de objAdmMeioPagtoCondPagto
        objAdmMeioPagto.colCondPagtoLoja.Add objAdmMeioPagtoCondPagto

        'Lê as Parcelas no Banco de Dados
        lErro = CF("AdmMeioPagtoParcelas_Le", objAdmMeioPagtoCondPagto)
        If lErro <> SUCESSO Then gError 104084
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 104072

     Loop
        
        
    'Fecha o comando
    Call Comando_Fechar(lComando)

    AdmMeioPagtoCondPagto_Le = SUCESSO

    Exit Function

Erro_AdmMeioPagtoCondPagto_Le:

    AdmMeioPagtoCondPagto_Le = gErr

    Select Case gErr
        
        Case 104086
            'Erro Tratado Dentro da Função Chamadora
            
        Case 104073
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 104070, 104071, 1040732
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ADMMEIOPAGTO1", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151464)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function AdmMeioPagtoParcelas_Le(ByVal objAdmMeioPagtoCondPagto As ClassAdmMeioPagtoCondPagto) As Long
'Le todas as Pareclas Viculadas ao Parcelamento
Dim lErro As Long
Dim lComando As Long
Dim tAdmMeioPagtoParcelas As typeAdmMeioPagtoParcelas
Dim objAdmMeioPagtoParcelas As ClassAdmMeioPagtoParcelas

On Error GoTo Erro_AdmMeioPagtoParcelas_Le_Todas

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 104076
    
    'Seleciona código e nome dos meios de pagamentos da tabela AdmMeioPagto
    lErro = Comando_Executar(lComando, "SELECT Parcela , PercRecebimento, IntervaloRecebto FROM AdmMeioPagtoParcelas WHERE AdmMeioPagto = ? AND FilialEmpresa = ? AND Parcelamento =? ORDER BY Parcela", tAdmMeioPagtoParcelas.iParcela, tAdmMeioPagtoParcelas.dPercRecebimento, tAdmMeioPagtoParcelas.iIntervaloRecebto, objAdmMeioPagtoCondPagto.iAdmMeioPagto, objAdmMeioPagtoCondPagto.iFilialEmpresa, objAdmMeioPagtoCondPagto.iParcelamento)
    If lErro <> AD_SQL_SUCESSO Then gError 104077

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 104078

    Do While lErro = AD_SQL_SUCESSO
        
        'Inicializa  o Obj relacionado à parcelas
        Set objAdmMeioPagtoParcelas = New ClassAdmMeioPagtoParcelas
        
        objAdmMeioPagtoParcelas.iAdmMeioPagto = objAdmMeioPagtoCondPagto.iAdmMeioPagto
        objAdmMeioPagtoParcelas.dPercRecebimento = tAdmMeioPagtoParcelas.dPercRecebimento
        objAdmMeioPagtoParcelas.iFilialEmpresa = objAdmMeioPagtoCondPagto.iFilialEmpresa
        objAdmMeioPagtoParcelas.iParcela = tAdmMeioPagtoParcelas.iParcela
        objAdmMeioPagtoParcelas.iParcelamento = objAdmMeioPagtoCondPagto.iParcelamento
        objAdmMeioPagtoParcelas.iIntervaloRecebimento = tAdmMeioPagtoParcelas.iIntervaloRecebto
        
        'Adcionar na Coleção colParcelas o objAdmMeioPagtoParcelas
        objAdmMeioPagtoCondPagto.colParcelas.Add objAdmMeioPagtoParcelas
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 104079

     Loop
    
    'Fecha o comando
    Call Comando_Fechar(lComando)

    AdmMeioPagtoParcelas_Le = SUCESSO

    Exit Function

Erro_AdmMeioPagtoParcelas_Le_Todas:

    AdmMeioPagtoParcelas_Le = gErr

    Select Case gErr
        
        Case 104085
            Call Rotina_Erro(vbOKOnly, "ERRO_ADMMEIOPAGTOCONDPAGTO_SEM_PARCELAS", gErr, objAdmMeioPagtoCondPagto.sNomeParcelamento)
            
        Case 104076
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 104077, 104078, 104079
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ADMMEIOPAGTO1", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151465)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function POS_Le_Lockado(ByVal objPos As ClassPOS)
'Lê todos os dados do POS cujo código foi passado como parâmetro e LOCKA o registro

Dim lComando As Long
Dim lErro As Long
Dim tPOS As typePOS

On Error GoTo Erro_POS_Le_Lockado

    'Inicia o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 103039
    
    'Inicializa a variável que receberá o código do POS
    tPOS.sCodigo = String(STRING_POS_CODIGO, 0)
    
    'Procura o POS no BD a partir do código
    lErro = Comando_ExecutarPos(lComando, "SELECT Codigo, FilialEmpresa, Rede FROM POS WHERE Codigo = ?", 0, tPOS.sCodigo, tPOS.iFilialEmpresa, tPOS.iRede, objPos.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 103040
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 103041
    
    'Se não encontrou o POS => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 103042
    
    lErro = Comando_LockShared(lComando)
    If lErro <> SUCESSO Then gError 103043
    
    'Passa para objPOS os valores lidos do BD
    objPos.sCodigo = tPOS.sCodigo
    objPos.iFilialEmpresa = tPOS.iFilialEmpresa
    objPos.iRede = tPOS.iRede
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    POS_Le_Lockado = SUCESSO
    
    Exit Function
    
Erro_POS_Le_Lockado:

    POS_Le_Lockado = gErr
    
    Select Case gErr
    
        Case 103039
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 103040, 103041
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_POS", gErr, objPos.sCodigo)
            
        Case 103042 'POS não encontrado. Erro a ser tratado na rotina chamadora
            
        Case 103043
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_POS", gErr, objPos.sCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151466)
            
    End Select
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

'Tela Rede
Function Redes_Le_Todas(ByVal colRedes As Collection)
'Função que Lê todas os Registros no banco de Dados relacionados a Rede

Dim lErro As Long
Dim tRede As typeRede
Dim objRede As ClassRede
Dim lComando As Long

On Error GoTo Erro_Redes_Le_Todas

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 104209

    'Inicializa o buffer nome
    tRede.sNome = String(STRING_REDE_NOME, 0)
    tRede.sEstabelecimento = String(STRING_REDE_ESTABELECIMENTO, 0)
    tRede.sDirImportacaoExtrato = String(STRING_REDE_DIRIMPORTACAOEXTRATO, 0)

    'Seleciona código e nome das Redes no Banco de Dados da Mesma FilialEmpresa
    lErro = Comando_Executar(lComando, "SELECT Codigo, Nome, Cliente, FilialCli, Bandeira, Estabelecimento, CodConta, DirImportacaoExtrato FROM Redes WHERE FilialEmpresa =? ORDER BY Codigo ", tRede.iCodigo, tRede.sNome, tRede.lCliente, tRede.iFilialCli, tRede.iBandeira, tRede.sEstabelecimento, tRede.iCodConta, tRede.sDirImportacaoExtrato, giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 104210

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 104211

    Do While lErro = AD_SQL_SUCESSO

        'Inicializa  o objRede para Receber os Campos Reacionados no Banco de Dados
        Set objRede = New ClassRede

        objRede.iCodigo = tRede.iCodigo
        objRede.sNome = tRede.sNome
        objRede.lCliente = tRede.lCliente
        objRede.iFilialCli = tRede.iFilialCli
        objRede.iBandeira = tRede.iBandeira
        objRede.sEstabelecimento = tRede.sEstabelecimento
        objRede.iCodConta = tRede.iCodConta
        objRede.sDirImportacaoExtrato = tRede.sDirImportacaoExtrato
        
        'Adcionar na Coleção colRedes para adcionar o objRede
        colRedes.Add objRede

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 104212

     Loop

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Redes_Le_Todas = SUCESSO

    Exit Function

Erro_Redes_Le_Todas:

    Redes_Le_Todas = gErr

    Select Case gErr

        Case 104209
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 104210, 104211, 104212
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REDE1", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151467)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Rede_Le(ByVal objRede As ClassRede) As Long

Dim lErro As Long
Dim tRede As typeRede
Dim lComando As Long
Dim iAtivo As Integer

On Error GoTo Erro_Redes_Le

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 104217

    'Inicializa o buffer nome
    tRede.sNome = String(STRING_REDE_NOME, 0)
    tRede.sEstabelecimento = String(STRING_REDE_ESTABELECIMENTO, 0)
    tRede.sDirImportacaoExtrato = String(STRING_REDE_DIRIMPORTACAOEXTRATO, 0)

    'Seleciona código e nome das Redes no Banco de Dados da Mesma FilialEmpresa
    lErro = Comando_Executar(lComando, "SELECT Codigo, Nome , Cliente, FilialCli, Ativo, Bandeira, Estabelecimento, CodConta, DirImportacaoExtrato FROM Redes WHERE Codigo = ? AND FilialEmpresa =? ", tRede.iCodigo, tRede.sNome, tRede.lCliente, tRede.iFilialCli, iAtivo, tRede.iBandeira, tRede.sEstabelecimento, tRede.iCodConta, tRede.sDirImportacaoExtrato, objRede.iCodigo, giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 104218

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 104219

    'Se Não Encontrar Registros
    If lErro = AD_SQL_SEM_DADOS Then gError 104244

    objRede.iCodigo = tRede.iCodigo
    objRede.sNome = tRede.sNome
    objRede.lCliente = tRede.lCliente
    objRede.iFilialCli = tRede.iFilialCli
    objRede.iAtivo = iAtivo
    objRede.iBandeira = tRede.iBandeira
    objRede.sEstabelecimento = tRede.sEstabelecimento
    objRede.iCodConta = tRede.iCodConta
    objRede.sDirImportacaoExtrato = tRede.sDirImportacaoExtrato
    
    Rede_Le = SUCESSO
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
Erro_Redes_Le:

    Rede_Le = gErr
    
    Select Case gErr
        Case 104217
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 104218, 104219
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REDE", gErr, objRede.iCodigo)

        Case 104244
            'Erro sem menssagem de Erro, Cliente já sabe que Rede não Existe.

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151468)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function Teclado_Le(ByVal objTeclado As ClassTeclado) As Long

Dim lErro As Long
Dim tTeclado As typeTeclado
Dim lComando As Long

On Error GoTo Erro_teclado_Le

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 99456

    'Inicializa o buffer nome
    tTeclado.sDescricao = String(STRING_TECLADO_DESCRICAO, 0)

    'Seleciona código e Descricao dos teclados no Banco de Dados da Mesma FilialEmpresa
    lErro = Comando_Executar(lComando, "SELECT Codigo, Descricao  FROM teclado WHERE Codigo = ? AND FilialEmpresa =? ", tTeclado.iCodigo, tTeclado.sDescricao, objTeclado.iCodigo, giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 99457

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 99458

    'Se Não Encontrar Registros
    If lErro = AD_SQL_SEM_DADOS Then gError 99459

    objTeclado.iCodigo = tTeclado.iCodigo
    objTeclado.sDescricao = tTeclado.sDescricao
            
    Teclado_Le = SUCESSO
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
Erro_teclado_Le:

    Teclado_Le = gErr
    
    Select Case gErr
    
        Case 99456
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 99457, 99458
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TECLADO", gErr, objTeclado.iCodigo)

        Case 99459
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151469)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function Teclado_Le_Todos(ByVal colTeclado As Collection) As Long

Dim lErro As Long
Dim tTeclado As typeTeclado
Dim lComando As Long
Dim objTeclado As ClassTeclado

On Error GoTo Erro_teclado_Le_Todos

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 99511

    'Inicializa o buffer nome
    tTeclado.sDescricao = String(STRING_TECLADO_DESCRICAO, 0)

    'Seleciona código e Descricao dos teclados no Banco de Dados da Mesma FilialEmpresa
    lErro = Comando_Executar(lComando, "SELECT Codigo, Descricao  FROM Teclado WHERE FilialEmpresa =? ", tTeclado.iCodigo, tTeclado.sDescricao, giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 99512

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 99513
    
    'Se Não Encontrar Registros
    If lErro = AD_SQL_SEM_DADOS Then gError 99514
    
    Do While lErro = AD_SQL_SUCESSO
        
        Set objTeclado = New ClassTeclado
        
        objTeclado.iCodigo = tTeclado.iCodigo
        objTeclado.sDescricao = tTeclado.sDescricao
        
        colTeclado.Add objTeclado
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 99515
                
    Loop
    
    Teclado_Le_Todos = SUCESSO
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
Erro_teclado_Le_Todos:

    Teclado_Le_Todos = gErr
    
    Select Case gErr
    
        Case 99511
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 99512, 99513, 99515
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TECLADO", gErr, tTeclado.iCodigo)

        Case 99514
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151470)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function TecladoProduto_Le(ByVal objTecladoProduto As ClassTecladoProduto) As Long

Dim lErro As Long
Dim tTecladoProduto As typeTecladoProduto
Dim lComando As Long

On Error GoTo Erro_TecladoProduto_Le

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 99523

    'Inicializa o buffer nome
    tTecladoProduto.sDescricao = String(STRING_TECLADOPRODUTO_DESCRICAO, 0)

    'Seleciona código e Descricao dos TecladoProdutoItenss no Banco de Dados da Mesma FilialEmpresa
    lErro = Comando_Executar(lComando, "SELECT Descricao, Teclado,Padrao   FROM TecladoProduto WHERE Codigo = ? AND FilialEmpresa =? ", tTecladoProduto.sDescricao, tTecladoProduto.iTeclado, tTecladoProduto.iPadrao, objTecladoProduto.iCodigo, giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 99524

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 99525

    'Se Não Encontrar Registros
    If lErro = AD_SQL_SEM_DADOS Then gError 99526

    objTecladoProduto.iPadrao = tTecladoProduto.iPadrao
    objTecladoProduto.sDescricao = tTecladoProduto.sDescricao
    objTecladoProduto.iTeclado = tTecladoProduto.iTeclado
        
    TecladoProduto_Le = SUCESSO
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
Erro_TecladoProduto_Le:

    TecladoProduto_Le = gErr
    
    Select Case gErr
    
        Case 99523
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 99524, 99525
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TECLADOPRODUTO", gErr, objTecladoProduto.iCodigo)

        Case 99526
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151471)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function TecladoProdutoItens_Le(ByVal objTecladoProduto As ClassTecladoProduto) As Long

Dim lErro As Long
Dim tTecladoProdutoItens As typeTecladoProdutoItens
Dim lComando As Long
Dim objTecladoProdutoItens As ClassTecladoProdutoItem

On Error GoTo Erro_TecladoProdutoItens_Le

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 99527
    
    With tTecladoProdutoItens
    
    'Inicializa o buffer
    .sProduto = String(STRING_TECLADOPRODUTOITENS_PRODUTO, 0)
    .sTitulo = String(STRING_TECLADOPRODUTOITENS_TITULO, 0)
    .sArvoreKey = String(STRING_TECLADOPRODUTOITENS_ARVOREKEY, 0)
    
    'Seleciona todos os dados dos itens do TecladoProduto no Banco de Dados
    lErro = Comando_Executar(lComando, "SELECT Produto, Titulo, Cor, Tecla, ArvoreKey  FROM TecladoProdutoItem WHERE Teclado=? AND FilialEmpresa =? ORDER BY ArvoreKey ", .sProduto, .sTitulo, .lCor, .iTecla, .sArvoreKey, objTecladoProduto.iCodigo, giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 99528

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 99529
        
    Do While lErro = AD_SQL_SUCESSO
        
        Set objTecladoProdutoItens = New ClassTecladoProdutoItem
        
        objTecladoProdutoItens.sProduto = .sProduto
        objTecladoProdutoItens.sTitulo = .sTitulo
        objTecladoProdutoItens.iTecla = .iTecla
        objTecladoProdutoItens.lColor = .lCor
        objTecladoProdutoItens.sArvoreKey = .sArvoreKey
        
        objTecladoProduto.colTecladoProdutoItem.Add objTecladoProdutoItens, objTecladoProdutoItens.sArvoreKey
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 99530
                
    Loop
    
    End With
    
    TecladoProdutoItens_Le = SUCESSO
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
Erro_TecladoProdutoItens_Le:

    TecladoProdutoItens_Le = gErr
    
    Select Case gErr
    
        Case 99527
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 99528, 99529, 99530
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TECLADOPRODUTO", gErr, objTecladoProduto.iTeclado)
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151472)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function Cheque_Le(ByVal objCheque As ClassChequePre) As Long
'Função que lê as Informações do Cheque passado como parâmetro no banco de dados, Carregando no objCheque

Dim lErro As Long
Dim lComando As Long
Dim lCupomFiscal As Long
Dim tCheque As typeChequePre
Dim iTipoMovcx As Integer

On Error GoTo Erro_Cheque_Le

    'Inicia a Abertura de o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 104343

    'Inicializa o Buffer das variáveis
    tCheque.sAgencia = String(STRING_BUFFER_MAX_TEXTO, 0)
    tCheque.sContaCorrente = String(STRING_BUFFER_MAX_TEXTO, 0)
    tCheque.sCPFCGC = String(STRING_BUFFER_MAX_TEXTO, 0)
    
    'Verifica se Esta Trabalhando no caixa central
    If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL Then

        'Setar a Variável para consulta SQL, quando a Consulta vem mediante ao Browser
        objCheque.lSequencial = objCheque.lSequencialLoja

        'Procura o Cheque no BD a partir do código e a FilialEmpresaLoja
        lErro = Comando_Executar(lComando, "SELECT NumIntCheque, Cliente, Filial , Banco , ChequePre.Agencia , ContaCorrente , ChequePre.Numero, DataDeposito , ChequePre.Valor , NumBordero , ChequePre.CPFCGC , NumMovtoCaixa , Aprovado , NaoEspecificado , NumBorderoLoja , ChequePre.FilialEmpresa , SequencialLoja , SequencialBack,  ChequePre.Status , MovimentosCaixa.CupomFiscal, MovimentosCaixa.Tipo, ChequePre.Localizacao, ChequePre.NumBorderoLojaBanco, ChequePre.Caixa, ChequePre.FilialEmpresaLoja, ChequePre.NumMovtoSangria, ChequePre.SequencialCaixa, ChequePre.FilialEmpresa, ChequePre.DataEmissao  FROM ChequePre , MovimentosCaixa WHERE ChequePre.SequencialLoja = ? AND ChequePre.NumMovtoCaixa = MovimentosCaixa.NumMovto AND MovimentosCaixa.FilialEmpresa = ChequePre.FilialEmpresaLoja AND ChequePre.FilialEmpresaLoja = ? AND ChequePre.Status <> 5", _
        tCheque.lNumIntCheque, tCheque.lCliente, tCheque.iFilial, tCheque.iBanco, tCheque.sAgencia, tCheque.sContaCorrente, tCheque.lNumero, tCheque.dtDataDeposito, tCheque.dValor, tCheque.lNumBordero, tCheque.sCPFCGC, tCheque.lNumMovtoCaixa, tCheque.iAprovado, tCheque.iNaoEspecificado, tCheque.lNumBorderoLoja, tCheque.iFilialEmpresa, tCheque.lSequencialLoja, tCheque.lSequencialBack, tCheque.iStatus, lCupomFiscal, iTipoMovcx, tCheque.iLocalizacao, tCheque.lNumBorderoLojaBanco, tCheque.iCaixa, tCheque.iFilialEmpresaLoja, tCheque.lNumMovtoSangria, tCheque.lSequencialCaixa, tCheque.iFilialEmpresa, tCheque.dtDataEmissao, objCheque.lSequencial, objCheque.iFilialEmpresaLoja)
        If lErro <> AD_SQL_SUCESSO Then gError 104463

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 104464

        'Se não encontrou o Cheque => erro
        If lErro = AD_SQL_SEM_DADOS Then gError 104465

        
    Else
        
        objCheque.lSequencial = objCheque.lSequencialBack

        'Procura o Cheque no BD a partir do código e a FilialEmpresaLoja
        lErro = Comando_Executar(lComando, "SELECT NumIntCheque, Cliente, Filial , Banco , Agencia , ContaCorrente , ChequePre.Numero, DataDeposito , ChequePre.Valor , NumBordero , ChequePre.CPFCGC , NumMovtoCaixa , Aprovado , NaoEspecificado , NumBorderoLoja  ,  SequencialLoja , SequencialBack, ChequePre.Status ,ECF , MovimentosCaixa.CupomFiscal, ChequePre.Localizacao, ChequePre.NumBorderoLojaBanco, ChequePre.Caixa, ChequePre.FilialEmpresaLoja, ChequePre.NumMovtoSangria, ChequePre.SequencialCaixa, ChequePre.FilialEmpresa, ChequePre.DataEmissao  FROM CupomFiscal, ChequePre , MovimentosCaixa WHERE ChequePre.SequencialBack = ? AND ChequePre.NumMovtoCaixa = MovimentosCaixa.NumMovto AND MovimentosCaixa.CupomFiscal = CupomFiscal.NumIntDoc AND MovimentosCaixa.FilialEmpresa = ChequePre.FilialEmpresa AND ChequePre.Status <> 5", _
        tCheque.lNumIntCheque, tCheque.lCliente, tCheque.iFilial, tCheque.iBanco, tCheque.sAgencia, tCheque.sContaCorrente, tCheque.lNumero, tCheque.dtDataDeposito, tCheque.dValor, tCheque.lNumBordero, tCheque.sCPFCGC, tCheque.lNumMovtoCaixa, tCheque.iAprovado, tCheque.iNaoEspecificado, tCheque.lNumBorderoLoja, tCheque.lSequencialLoja, tCheque.lSequencialBack, tCheque.iStatus, tCheque.iECF, lCupomFiscal, tCheque.iLocalizacao, tCheque.lNumBorderoLojaBanco, tCheque.iCaixa, tCheque.iFilialEmpresaLoja, tCheque.lNumMovtoSangria, tCheque.lSequencialCaixa, tCheque.iFilialEmpresa, tCheque.dtDataEmissao, objCheque.lSequencialBack)
        If lErro <> AD_SQL_SUCESSO Then gError 104344

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 104345

        'Se não encontrou o Cheque => erro
        If lErro = AD_SQL_SEM_DADOS Then gError 104346
        
    End If

    'se tiver cupomfiscal envolvido ==> pega o ECF
    If lCupomFiscal <> 0 Then

        'Procura o Cheque no BD a partir do código e a FilialEmpresaLoja
        lErro = Comando_Executar(lComando, "SELECT ECF FROM CupomFiscal WHERE FilialEmpresa = ? AND NumIntDoc = ?", tCheque.iECF, objCheque.iFilialEmpresa, lCupomFiscal)
        If lErro <> AD_SQL_SUCESSO Then gError 126056

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 126057

        'Se não encontrou o Cheque => erro
        If lErro = AD_SQL_SEM_DADOS Then gError 126058

    End If

    'Passa para objCheque os valores lidos do BD
    objCheque.dtDataDeposito = tCheque.dtDataDeposito
    objCheque.dValor = tCheque.dValor
    objCheque.iAprovado = tCheque.iAprovado
    objCheque.iBanco = tCheque.iBanco
    objCheque.iFilial = tCheque.iFilial
    objCheque.iFilialEmpresa = tCheque.iFilialEmpresa
    objCheque.iNaoEspecificado = tCheque.iNaoEspecificado
    objCheque.lCliente = tCheque.lCliente
    objCheque.lNumBordero = tCheque.lNumBordero
    objCheque.lNumBorderoLoja = tCheque.lNumBorderoLoja
    objCheque.lNumero = tCheque.lNumero
    objCheque.lNumIntCheque = tCheque.lNumIntCheque
    objCheque.lNumMovtoCaixa = tCheque.lNumMovtoCaixa
    objCheque.lSequencialBack = tCheque.lSequencialBack
    objCheque.lSequencialLoja = tCheque.lSequencialLoja
    objCheque.sAgencia = tCheque.sAgencia
    objCheque.sContaCorrente = tCheque.sContaCorrente
    objCheque.sCPFCGC = tCheque.sCPFCGC
    objCheque.lCupomFiscal = lCupomFiscal
    objCheque.iECF = tCheque.iECF
    objCheque.iStatus = tCheque.iStatus
    objCheque.iLocalizacao = tCheque.iLocalizacao
    objCheque.lNumBorderoLojaBanco = tCheque.lNumBorderoLojaBanco
    objCheque.iCaixa = tCheque.iCaixa
    objCheque.iFilialEmpresaLoja = tCheque.iFilialEmpresaLoja
    objCheque.lNumMovtoSangria = tCheque.lNumMovtoSangria
    objCheque.lSequencialCaixa = tCheque.lSequencialCaixa
    objCheque.dtDataEmissao = tCheque.dtDataEmissao
    
    'Fecha o comando
    Call Comando_Fechar(lComando)

    Cheque_Le = SUCESSO

    Exit Function

Erro_Cheque_Le:

    Cheque_Le = gErr

    Select Case gErr

        Case 104343
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 104344, 104345, 104463, 104464
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CHEQUE1", gErr, objCheque.lSequencial)

        Case 104346, 104465

        Case 126056, 126057
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CUPOMFISCAL2", gErr, lCupomFiscal)

        Case 126058
            Call Rotina_Erro(vbOKOnly, "ERRO_CUPOM_FISCAL_NAO_CADASTRADO1", gErr, lCupomFiscal)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151473)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Cheque_Lock(ByVal objCheque As ClassChequePre) As Long
'Função Realiza lock na Tabela de ChequePre com o Sequencial e Filial Passado como parâmetro

Dim lErro As Long
Dim lComando As Long
Dim lCodigoMsgErro As Long
Dim tChequePre As typeChequePre

On Error GoTo Erro_Cheque_Lock

     'Abertura dos Comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 104349

    If giLocalOperacao = LOCALOPERACAO_BACKOFFICE Or giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL_BACKOFFICE Then

        lErro = Comando_ExecutarPos(lComando, "SELECT NumIntCheque, NumMovtoCaixa , Valor, Localizacao, NaoEspecificado, Status, DataEmissao FROM ChequePre WHERE SequencialBack = ?", 0, tChequePre.lNumIntCheque, tChequePre.lNumMovtoCaixa, tChequePre.dValor, tChequePre.iLocalizacao, tChequePre.iNaoEspecificado, tChequePre.iStatus, tChequePre.dtDataEmissao, objCheque.lSequencialBack)
        If lErro <> AD_SQL_SUCESSO Then gError 104484

        'Devido a Menssagem de Erro, Serve para as duas situações tanto no backoffice, quanto no Caixa Central
        lCodigoMsgErro = objCheque.lSequencialBack

    ElseIf giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL Then

        lErro = Comando_ExecutarPos(lComando, "SELECT NumMovtoCaixa , Valor,  Localizacao, NaoEspecificado, Status, DataEmissao FROM ChequePre WHERE SequencialLoja = ? AND FilialEmpresaLoja = ?", 0, tChequePre.lNumMovtoCaixa, tChequePre.dValor, tChequePre.iLocalizacao, tChequePre.iNaoEspecificado, tChequePre.iStatus, tChequePre.dtDataEmissao, objCheque.lSequencialLoja, objCheque.iFilialEmpresaLoja)
        If lErro <> AD_SQL_SUCESSO Then gError 104350
        
        'Devido a Menssagem de Erro, Serve para as duas situações tanto no backoffice, quanto no Caixa Central
        lCodigoMsgErro = objCheque.lSequencialLoja

    End If
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 104351

    If lErro = AD_SQL_SEM_DADOS Then gError 104352

    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then gError 104353
    
    objCheque.lNumIntCheque = tChequePre.lNumIntCheque
    objCheque.lNumMovtoCaixa = tChequePre.lNumMovtoCaixa
    objCheque.dValor = tChequePre.dValor
    objCheque.iLocalizacao = tChequePre.iLocalizacao
    objCheque.iNaoEspecificado = tChequePre.iNaoEspecificado
    objCheque.iStatus = tChequePre.iStatus
    objCheque.dtDataEmissao = tChequePre.dtDataEmissao
    
    Cheque_Lock = SUCESSO

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

Erro_Cheque_Lock:

    Cheque_Lock = gErr

    Select Case gErr

        Case 104349
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 104350, 104351, 104484
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CHEQUE1", gErr, lCodigoMsgErro)

        Case 104352

        Case 104353
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_CHEQUE", gErr, lCodigoMsgErro)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151474)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Cheque_Le_Nao_Especificado(ByVal objCheque As ClassChequePre) As Long
'Função que lê no Banco de Dados cheques não especificados

Dim lErro As Long
Dim tCheque As typeChequePre
Dim lComando As Long
Dim iECF As Integer
Dim iCupomFiscal As Integer
Dim lCodigoMsgErro As Long

On Error GoTo Erro_Cheque_Le_Nao_Especificado

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 104363

   'Inicializa o Buffer das variáveis
    tCheque.sAgencia = String(STRING_AGENCIA, 0)
    tCheque.sContaCorrente = String(STRING_CONTACORRENTE, 0)
    tCheque.sCPFCGC = String(STRING_CPFCGC, 0)

    'Flag para Marcar Cheque não especificado
    objCheque.iNaoEspecificado = CHEQUE_NAO_ESPECIFICADO

    'Se a Operação for no Back então o Sequencial
    If giLocalOperacao = LOCALOPERACAO_BACKOFFICE Then
    
        'Devido a Menssagem de Erro, Serve para as duas situações tanto no backoffice, quanto no Caixa Central
        lCodigoMsgErro = objCheque.lSequencialBack
    
    
        'Não Esquecer depois de Mudar a Tabela Atenção
        lErro = Comando_Executar(lComando, "SELECT NumIntCheque, Cliente, Filial , Banco , Agencia , ContaCorrente , Numero, DataDeposito , Valor , NumBordero , CPFCGC , Aprovado  , NumBorderoLoja , FilialEmpresa , SequencialLoja , SequencialBack , Status, DataEmissao FROM ChequePre  WHERE   NumMovtoCaixa = ?  AND NaoEspecificado = ? AND ChequePre.Status <> ? ", _
        tCheque.lNumIntCheque, tCheque.lCliente, tCheque.iFilial, tCheque.iBanco, tCheque.sAgencia, tCheque.sContaCorrente, tCheque.lNumero, tCheque.dtDataDeposito, tCheque.dValor, tCheque.lNumBordero, tCheque.sCPFCGC, tCheque.iAprovado, tCheque.lNumBorderoLoja, tCheque.iFilialEmpresa, tCheque.lSequencialLoja, tCheque.lSequencialBack, tCheque.iStatus, tCheque.dtDataEmissao, objCheque.lNumMovtoCaixa, objCheque.iNaoEspecificado, STATUS_EXCLUIDO)
        If lErro <> AD_SQL_SUCESSO Then gError 104486
    
        
    Else
    
        'Devido a Menssagem de Erro, Serve para as duas situações tanto no backoffice, quanto no Caixa Central
        lCodigoMsgErro = objCheque.lSequencialLoja
    
        lErro = Comando_Executar(lComando, "SELECT NumIntCheque, Cliente, Filial , Banco , Agencia , ContaCorrente , Numero, DataDeposito , Valor , NumBordero , CPFCGC , Aprovado  , NumBorderoLoja , FilialEmpresa , SequencialLoja , SequencialBack , Status, DataEmissao FROM ChequePre  WHERE   NumMovtoCaixa = ? AND FilialEmpresaLoja = ? AND NaoEspecificado = ? AND ChequePre.Status <> ?  ", _
        tCheque.lNumIntCheque, tCheque.lCliente, tCheque.iFilial, tCheque.iBanco, tCheque.sAgencia, tCheque.sContaCorrente, tCheque.lNumero, tCheque.dtDataDeposito, tCheque.dValor, tCheque.lNumBordero, tCheque.sCPFCGC, tCheque.iAprovado, tCheque.lNumBorderoLoja, tCheque.iFilialEmpresa, tCheque.lSequencialLoja, tCheque.lSequencialBack, tCheque.iStatus, tCheque.dtDataEmissao, objCheque.lNumMovtoCaixa, objCheque.iFilialEmpresaLoja, objCheque.iNaoEspecificado, STATUS_EXCLUIDO)
        If lErro <> AD_SQL_SUCESSO Then gError 104364
        
    End If
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 104365

    'Se não encontrou o Cheque não Especificado  => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 104366

    'Carrega o objCheque com os valores lidos do BD
    objCheque.dtDataDeposito = tCheque.dtDataDeposito
    objCheque.dValor = tCheque.dValor
    objCheque.iAprovado = tCheque.iAprovado
    objCheque.iBanco = tCheque.iBanco
    objCheque.iFilial = tCheque.iFilial
    objCheque.iFilialEmpresa = tCheque.iFilialEmpresa
    objCheque.iNaoEspecificado = tCheque.iNaoEspecificado
    objCheque.lCliente = tCheque.lCliente
    objCheque.lNumBordero = tCheque.lNumBordero
    objCheque.lNumBorderoLoja = tCheque.lNumBorderoLoja
    objCheque.lNumero = tCheque.lNumero
    objCheque.lNumIntCheque = tCheque.lNumIntCheque
    objCheque.lSequencialBack = tCheque.lSequencialBack
    objCheque.lSequencialLoja = tCheque.lSequencialLoja
    objCheque.sAgencia = tCheque.sAgencia
    objCheque.sContaCorrente = tCheque.sContaCorrente
    objCheque.sCPFCGC = tCheque.sCPFCGC
    objCheque.iStatus = tCheque.iStatus
    objCheque.dtDataEmissao = tCheque.dtDataEmissao

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Cheque_Le_Nao_Especificado = SUCESSO

    Exit Function

Erro_Cheque_Le_Nao_Especificado:

    Cheque_Le_Nao_Especificado = gErr

    Select Case gErr
    
        Case 104363
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 104364, 104365, 104486
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CHEQUE1", gErr, lCodigoMsgErro)

        Case 104366

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151475)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function BorderoCheque_Le(ByVal objBorderoCheque As ClassBorderoCheque) As Long
'Função que Lê do banco de Dados borderô Vinculado a Cheque

Dim lErro As Long
Dim tBorderoCheque As typeBorderoCheque
Dim lComando As Long

On Error GoTo Erro_BorderoCheque_Le

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 104382

    lErro = Comando_Executar(lComando, "SELECT CodNossaConta, DataImpressao, DataEnvio, DataBackoffice  FROM BorderoCheque WHERE  NumBordero  =  ? AND FilialEmpresa = ? ", _
    tBorderoCheque.iCodNossaConta, tBorderoCheque.dtDataImpressao, tBorderoCheque.dtDataEnvio, tBorderoCheque.dtDataBackoffice, objBorderoCheque.lNumBordero, objBorderoCheque.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 104383

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 104384

    'Se não encontrou o BorederoCheque   => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 104385

    'Mover os Dados Lidos para o objBorderoCheque
    objBorderoCheque.dtDataBackoffice = tBorderoCheque.dtDataBackoffice
    objBorderoCheque.dtDataEnvio = tBorderoCheque.dtDataEnvio
    objBorderoCheque.dtDataImpressao = tBorderoCheque.dtDataImpressao
    objBorderoCheque.iCodNossaConta = tBorderoCheque.iCodNossaConta
    
    'Fecha o comando
    Call Comando_Fechar(lComando)

    BorderoCheque_Le = SUCESSO

    Exit Function

Erro_BorderoCheque_Le:

    BorderoCheque_Le = gErr

    Select Case gErr
    
        Case 104382
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 104383, 104384
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BORDEROCHEQUE", gErr, objBorderoCheque.lNumBordero)

        Case 104385

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151476)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function CupomFiscal_Lock(ByVal objCupomFiscal As ClassCupomFiscal) As Long
'Função que lê no Banco e faz Lock no Cupom Fiscal Viculado ao Cheque

Dim lErro As Long
Dim tCupomFiscal As typeCupomFiscal
Dim lComando As Long
Dim lCupomFiscal As Long
Dim iECF As Integer

On Error GoTo Erro_CupomFiscal_Lock

    'Abertura dos Comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 104370

    'Inicializa o Buffer das variáveis
    tCupomFiscal.sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
    tCupomFiscal.sCPFCGC = String(STRING_CPFCGC1, 0)
    tCupomFiscal.sMotivoCancel = String(STRING_MOTIVOCANCEL, 0)
    
    'Leitura da tabela de cupom fiscal para fim de LockShared
    lErro = Comando_ExecutarPos(lComando, "SELECT NumIntDoc, Status, ECF , DataEmissao , HoraEmissao , CPFCGC , NumOrcamento , ValorTotal, ValorProdutos , ValorDesconto, ValorAcrescimo , TabelaPreco ,NaturezaOp , GerenteCancel , MotivoCancel , Vendedor, Caixa  FROM CupomFiscal  WHERE CupomFiscal.Numero = ?  AND CupomFiscal.FilialEmpresa = ? AND ECF = ?", 0, _
    tCupomFiscal.lNumIntDoc, tCupomFiscal.iStatus, tCupomFiscal.iECF, tCupomFiscal.dtDataEmissao, tCupomFiscal.dHoraEmissao, tCupomFiscal.sCPFCGC, tCupomFiscal.lNumOrcamento, tCupomFiscal.dValorTotal, tCupomFiscal.dValorProdutos, tCupomFiscal.dValorAcrescimo, tCupomFiscal.dValorAcrescimo, tCupomFiscal.iTabelaPreco, tCupomFiscal.sNaturezaOp, tCupomFiscal.lGerenteCancel, tCupomFiscal.sMotivoCancel, tCupomFiscal.iVendedor, tCupomFiscal.iCaixa, objCupomFiscal.lNumero, objCupomFiscal.iFilialEmpresa, objCupomFiscal.iECF)
    If lErro <> AD_SQL_SUCESSO Then gError 104371

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 104380

    'SE Não encontrar Registro
    If lErro = AD_SQL_SEM_DADOS Then gError 104375

    'Faz um Lock no Registro Encontrado
    lErro = Comando_LockShared(lComando)
    If lErro <> SUCESSO Then gError 104374

    'Traz os dados para o objCupomFiscal
    objCupomFiscal.dtDataEmissao = tCupomFiscal.dtDataEmissao
    objCupomFiscal.dValorAcrescimo = tCupomFiscal.dValorAcrescimo
    objCupomFiscal.dValorDesconto = tCupomFiscal.dValorDesconto
    objCupomFiscal.dValorProdutos = tCupomFiscal.dValorProdutos
    objCupomFiscal.dValorTotal = tCupomFiscal.dValorTotal
    objCupomFiscal.iStatus = tCupomFiscal.iStatus
    objCupomFiscal.iTabelaPreco = tCupomFiscal.iTabelaPreco
    objCupomFiscal.lNumIntDoc = tCupomFiscal.lNumIntDoc
    objCupomFiscal.lNumOrcamento = tCupomFiscal.lNumOrcamento
    objCupomFiscal.sCPFCGC = tCupomFiscal.sCPFCGC
    objCupomFiscal.sMotivoCancel = tCupomFiscal.sMotivoCancel
    objCupomFiscal.sNaturezaOp = tCupomFiscal.sNaturezaOp
    objCupomFiscal.iVendedor = tCupomFiscal.iVendedor
    objCupomFiscal.lGerenteCancel = tCupomFiscal.lGerenteCancel
    objCupomFiscal.iCodCaixa = tCupomFiscal.iCaixa
    
     'Fecha o comando
    Call Comando_Fechar(lComando)

    CupomFiscal_Lock = SUCESSO

    Exit Function

Erro_CupomFiscal_Lock:

    CupomFiscal_Lock = gErr

    Select Case gErr

        Case 104370
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 104371, 104380
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CUPOMFISCAL1", gErr, objCupomFiscal.lNumero)

        Case 104374
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_CUPOMFISCAL", gErr, objCupomFiscal.lNumero)

        Case 104375

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151477)

    End Select

     'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function MovimentosCaixaCupom_Lock(ByVal objMovimentosCaixa As ClassMovimentoCaixa) As Long
'Função que faz Lock no movimento de caixa correspondente ao Cheque

Dim lErro As Long
Dim tMovimentosCaixa As typeMovimentoCaixa
Dim lComando As Long

On Error GoTo Erro_MovimentosCaixaCupom_Lock

    'Abertura dos Comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 104376

    'Inicializa o Buffer das variáveis
    tMovimentosCaixa.sHistorico = String(STRING_HISTORICO1, 0)
    tMovimentosCaixa.sFavorecido = String(STRING_FAVORECIDO, 0)

    'Leitura da tabela de cupom fiscal para fim de LockShared
    lErro = Comando_ExecutarPos(lComando, "SELECT NumMovto, Caixa, Sequencial , AdmMeioPagto , Parcelamento , TipoCartao , Numero, DataMovimento , Valor, Historico , Favorecido ,NumRefInterna , MovtoTransf , MovtoEstorno , Gerente,Hora  FROM MovimentosCaixa WHERE Tipo = ? AND CupomFiscal = ? AND FilialEmpresa = ? ", 0, _
    tMovimentosCaixa.lNumMovto, tMovimentosCaixa.iCaixa, tMovimentosCaixa.lSequencial, tMovimentosCaixa.iAdmMeioPagto, tMovimentosCaixa.iParcelamento, tMovimentosCaixa.iTipoCartao, tMovimentosCaixa.lNumero, tMovimentosCaixa.dtDataMovimento, tMovimentosCaixa.dValor, tMovimentosCaixa.sHistorico, tMovimentosCaixa.sFavorecido, tMovimentosCaixa.lNumRefInterna, tMovimentosCaixa.lMovtoTransf, tMovimentosCaixa.lMovtoEstorno, tMovimentosCaixa.iGerente, tMovimentosCaixa.dHora, objMovimentosCaixa.iTipo, objMovimentosCaixa.lCupomFiscal, objMovimentosCaixa.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 104377

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 104381

    'SE Não encontrar Registro
    If lErro = AD_SQL_SEM_DADOS Then gError 104378

    'Faz um Lock no Registro Encontrado
    lErro = Comando_LockShared(lComando)
    If lErro <> SUCESSO Then gError 104379
    
    'Armazena o Resultado no objMovimentosCaixa
    objMovimentosCaixa.dtDataMovimento = tMovimentosCaixa.dtDataMovimento
    objMovimentosCaixa.dValor = tMovimentosCaixa.dValor
    objMovimentosCaixa.iAdmMeioPagto = tMovimentosCaixa.iAdmMeioPagto
    objMovimentosCaixa.iCaixa = tMovimentosCaixa.iCaixa
    objMovimentosCaixa.iGerente = tMovimentosCaixa.iGerente
    objMovimentosCaixa.iParcelamento = tMovimentosCaixa.iParcelamento
    objMovimentosCaixa.iTipoCartao = tMovimentosCaixa.iTipoCartao
    objMovimentosCaixa.lMovtoEstorno = tMovimentosCaixa.lMovtoEstorno
    objMovimentosCaixa.lMovtoTransf = tMovimentosCaixa.lMovtoTransf
    objMovimentosCaixa.lNumero = tMovimentosCaixa.lNumero
    objMovimentosCaixa.lNumMovto = tMovimentosCaixa.lNumMovto
    objMovimentosCaixa.lNumRefInterna = tMovimentosCaixa.lNumRefInterna
    objMovimentosCaixa.lSequencial = tMovimentosCaixa.lSequencial
    objMovimentosCaixa.sFavorecido = tMovimentosCaixa.sFavorecido
    objMovimentosCaixa.sHistorico = tMovimentosCaixa.sHistorico
    objMovimentosCaixa.dHora = tMovimentosCaixa.dHora

    'Fecha o comando
    Call Comando_Fechar(lComando)

    MovimentosCaixaCupom_Lock = SUCESSO

    Exit Function

Erro_MovimentosCaixaCupom_Lock:

     MovimentosCaixaCupom_Lock = gErr

    Select Case gErr

        Case 104376
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 104377, 104381
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOS_CAIXA1", gErr, objMovimentosCaixa.lNumMovto)

        Case 104378

        Case 104379
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_MOVIMENTOSCAIXA", gErr, objMovimentosCaixa.lNumMovto)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151478)

    End Select

     'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function MovimentosCaixa_Lock1(ByVal objMovimentosCaixa As ClassMovimentoCaixa) As Long
'Função que faz Lock no movimento de caixa correspondente ao Cheque

Dim lErro As Long
Dim tMovimentosCaixa As typeMovimentoCaixa
Dim lComando As Long

On Error GoTo Erro_MovimentosCaixa_Lock1

    'Abertura dos Comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 104431

    'Inicializa o Buffer das variáveis
    tMovimentosCaixa.sHistorico = String(STRING_HISTORICO1, 0)
    tMovimentosCaixa.sFavorecido = String(STRING_FAVORECIDO, 0)

    'Leitura da tabela de cupom fiscal para fim de LockShared
    '???? A chave está errada. Trocar a clausula de where. ok Sergio
    lErro = Comando_ExecutarPos(lComando, "SELECT Caixa, Sequencial , AdmMeioPagto , Parcelamento , TipoCartao , Numero, DataMovimento , Valor, Historico , Favorecido , NumRefInterna , MovtoTransf , MovtoEstorno , Gerente,Hora  FROM MovimentosCaixa WHERE NumMovto = ? AND FilialEmpresa = ? ", 0, _
    tMovimentosCaixa.iCaixa, tMovimentosCaixa.lSequencial, tMovimentosCaixa.iAdmMeioPagto, tMovimentosCaixa.iParcelamento, tMovimentosCaixa.iTipoCartao, tMovimentosCaixa.lNumero, tMovimentosCaixa.dtDataMovimento, tMovimentosCaixa.dValor, tMovimentosCaixa.sHistorico, tMovimentosCaixa.sFavorecido, tMovimentosCaixa.lNumRefInterna, tMovimentosCaixa.lMovtoTransf, tMovimentosCaixa.lMovtoEstorno, tMovimentosCaixa.iGerente, tMovimentosCaixa.dHora, objMovimentosCaixa.lNumMovto, objMovimentosCaixa.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 104432

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 104433

    'SE Não encontrar Registro
    If lErro = AD_SQL_SEM_DADOS Then gError 104434

    'Faz um Lock no Registro Encontrado
    lErro = Comando_LockShared(lComando)
    If lErro <> SUCESSO Then gError 104435
    
    'Armazena o Resultado no objMovimentosCaixa
    objMovimentosCaixa.dtDataMovimento = tMovimentosCaixa.dtDataMovimento
    objMovimentosCaixa.dValor = tMovimentosCaixa.dValor
    objMovimentosCaixa.iAdmMeioPagto = tMovimentosCaixa.iAdmMeioPagto
    objMovimentosCaixa.iCaixa = tMovimentosCaixa.iCaixa
    objMovimentosCaixa.iGerente = tMovimentosCaixa.iGerente
    objMovimentosCaixa.iParcelamento = tMovimentosCaixa.iParcelamento
    objMovimentosCaixa.iTipoCartao = tMovimentosCaixa.iTipoCartao
    objMovimentosCaixa.lMovtoEstorno = tMovimentosCaixa.lMovtoEstorno
    objMovimentosCaixa.lMovtoTransf = tMovimentosCaixa.lMovtoTransf
    objMovimentosCaixa.lNumero = tMovimentosCaixa.lNumero
    objMovimentosCaixa.lNumRefInterna = tMovimentosCaixa.lNumRefInterna
    objMovimentosCaixa.lSequencial = tMovimentosCaixa.lSequencial
    objMovimentosCaixa.sFavorecido = tMovimentosCaixa.sFavorecido
    objMovimentosCaixa.sHistorico = tMovimentosCaixa.sHistorico
    objMovimentosCaixa.dHora = tMovimentosCaixa.dHora

     'Fecha o comando
    Call Comando_Fechar(lComando)

    MovimentosCaixa_Lock1 = SUCESSO

    Exit Function

Erro_MovimentosCaixa_Lock1:

     MovimentosCaixa_Lock1 = gErr

    Select Case gErr

        Case 104431
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 104432, 104433
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOS_CAIXA1", gErr, objMovimentosCaixa.lNumMovto)

        Case 104434

        Case 104435
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_MOVIMENTOSCAIXA", gErr, objMovimentosCaixa.lNumMovto)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151479)

    End Select

     'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function RecebimentoCarne_Le_Parcelas(ByVal objRecebimentoCarne As ClassRecebimentoCarne) As Long

Dim lErro As Long
Dim lComando As Long
Dim sSelect As String
Dim tInfoRecebimentoCarneVar As TypeInfoRecebimentoCarneVar

On Error GoTo Erro_RecebimentoCarne_Le_Parcelas

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 99964
    
    'Monta o select
    Call RecebimentoCarne_Le_Parcelas1(sSelect, objRecebimentoCarne)
        
    'Preparação da parte fixa
    lErro = RecebimentoCarne_Le_Parcelas2(lComando, sSelect, objRecebimentoCarne, tInfoRecebimentoCarneVar)
    If lErro <> SUCESSO Then gError 99969
    
    'Preparação da parte dinâmica
    lErro = RecebimentoCarne_Le_Parcelas3(lComando, tInfoRecebimentoCarneVar, objRecebimentoCarne)
    If lErro <> SUCESSO Then gError 99970
    
    'Processa a leitura das parcelas
    lErro = RecebimentoCarne_Le_Parcelas4(lComando, tInfoRecebimentoCarneVar, objRecebimentoCarne)
    If lErro <> SUCESSO Then gError 99970
    
    Call Comando_Fechar(lComando)
    
    RecebimentoCarne_Le_Parcelas = SUCESSO
    
    Exit Function
    
Erro_RecebimentoCarne_Le_Parcelas:
    
    RecebimentoCarne_Le_Parcelas = gErr
    
    Select Case gErr
        
        Case 99964, 99968 To 99970
                        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 151480)
                        
    End Select
   
    Exit Function
    
End Function

Sub RecebimentoCarne_Le_Parcelas1(sSelect As String, ByVal objRecebimentoCarne As ClassRecebimentoCarne)
   
Dim sWhere As String
Dim sFrom As String

   'select fixo
    sSelect = "SELECT CodBarrasCarne, Parcela, DataVencimento, Valor, CarneParcelas.NumIntDoc, Carne.NumIntDoc"

    'cláusula FROM
    sFrom = " FROM Carne, CarneParcelas"
    
    'Claúsula WHERE
    sWhere = " WHERE Carne.FilialEmpresa=CarneParcelas.FilialEmpresa" & _
             " AND Carne.NumIntDoc=CarneParcelas.NumIntCarne" & _
             " AND Cliente=?" & _
             " AND CarneParcelas.Status=?" & _
             " AND Carne.FilialEmpresa=?"
            
    'Acrescenta filtro caso exista carne inicial
    If Len(Trim(objRecebimentoCarne.sCodCarneDe)) <> 0 Then sWhere = sWhere & " AND CodBarrasCarne >= ?"
    
    'Acrescenta filtro caso exista carne final
    If Len(Trim(objRecebimentoCarne.sCodCarneAte)) <> 0 Then sWhere = sWhere & " AND CodBarrasCarne <= ?"
      
    'Acrescenta filtro caso exista data inicial
    If objRecebimentoCarne.dtDataVenctoDe <> DATA_NULA Then sWhere = sWhere & " AND DataVencimento >= ?"
    
    'Acrescenta filtro caso exista data final
    If objRecebimentoCarne.dtDataVenctoAte <> DATA_NULA Then sWhere = sWhere & " AND DataVencimento <= ?"
    
    sSelect = sSelect & sFrom & sWhere
      
End Sub

Private Function RecebimentoCarne_Le_Parcelas2(ByVal lComando As Long, ByVal sSelect As String, ByVal objRecebimentoCarne As ClassRecebimentoCarne, tInfoRecebimentoCarneVar As TypeInfoRecebimentoCarneVar) As Long
'Faz o Bind do Select e das variáveis comuns a todos os selects que receberão os dados lidos

Dim lret As Long
Dim lErro As Long

On Error GoTo Erro_RecebimentoCarne_Le_Parcelas2

    With tInfoRecebimentoCarneVar

        .vsCarneNumero = String(STRING_CARNE_CODBARRASCARNE, 0)

        lret = Comando_PrepararInt(lComando, sSelect)
        If (lret <> AD_SQL_SUCESSO) Then gError 99971
        
        lret = Comando_BindVarInt(lComando, .vsCarneNumero)
        If (lret <> AD_SQL_SUCESSO) Then gError 99972

        .viParcelaNumero = CInt(.viParcelaNumero)
        lret = Comando_BindVarInt(lComando, .viParcelaNumero)
        If (lret <> AD_SQL_SUCESSO) Then gError 99981

        .vdtParcelaVencimento = CDate(.vdtParcelaVencimento)
        lret = Comando_BindVarInt(lComando, .vdtParcelaVencimento)
        If (lret <> AD_SQL_SUCESSO) Then gError 99973

        .vdParcelaValor = CDbl(.vdParcelaValor)
        lret = Comando_BindVarInt(lComando, .vdParcelaValor)
        If (lret <> AD_SQL_SUCESSO) Then gError 99974

        .vlParcelaNumIntDoc = CLng(.vlParcelaNumIntDoc)
        lret = Comando_BindVarInt(lComando, .vlParcelaNumIntDoc)
        If (lret <> AD_SQL_SUCESSO) Then gError 109517
            
        .vlNumIntDoc = CLng(.vlNumIntDoc)
        lret = Comando_BindVarInt(lComando, .vlNumIntDoc)
        If (lret <> AD_SQL_SUCESSO) Then gError 109518
            
        .vlCodCliente = objRecebimentoCarne.lCodCliente
        lret = Comando_BindVarInt(lComando, .vlCodCliente)
        If (lret <> AD_SQL_SUCESSO) Then gError 99978
        
        .viStatus = objRecebimentoCarne.iStatus
        .viStatus = CInt(.viStatus)
        lret = Comando_BindVarInt(lComando, .viStatus)
        If (lret <> AD_SQL_SUCESSO) Then gError 99979
        
        .viFilialEmpresa = giFilialEmpresa
        lret = Comando_BindVarInt(lComando, .viFilialEmpresa)
        If (lret <> AD_SQL_SUCESSO) Then gError 99980
        
    End With

    RecebimentoCarne_Le_Parcelas2 = SUCESSO

    Exit Function

Erro_RecebimentoCarne_Le_Parcelas2:

    RecebimentoCarne_Le_Parcelas2 = gErr

    Select Case gErr

        Case 99971 To 99981, 109517, 109518
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CARNE_REC_BAIXADAS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151481)

    End Select

    Exit Function

End Function

Private Function RecebimentoCarne_Le_Parcelas3(ByVal lComando As Long, tInfoRecebimentoCarneVar As TypeInfoRecebimentoCarneVar, ByVal objRecebimentoCarne As ClassRecebimentoCarne) As Long
'Faz o Bind dos parâmetros dinâmicos

Dim lErro As Long
Dim lret As Long

On Error GoTo Erro_RecebimentoCarne_Le_Parcelas3

    With tInfoRecebimentoCarneVar
        
        'Se Código inicial preenchido
        If (Len(Trim(objRecebimentoCarne.sCodCarneDe)) <> 0) Then
            .vsCodCarneDe = objRecebimentoCarne.sCodCarneDe
            lret = Comando_BindVarInt(lComando, .vsCodCarneDe)
            If (lret <> AD_SQL_SUCESSO) Then gError 99982
        End If

        'Se Código final preenchido
        If (Len(Trim(objRecebimentoCarne.sCodCarneAte)) <> 0) Then
            .vsCodCarneAte = objRecebimentoCarne.sCodCarneAte
            lret = Comando_BindVarInt(lComando, .vsCodCarneAte)
            If (lret <> AD_SQL_SUCESSO) Then gError 99983
        End If
        
        'Se data início preenchida
        If (objRecebimentoCarne.dtDataVenctoDe <> DATA_NULA) Then
            .vdtCodDataVenctoDe = objRecebimentoCarne.dtDataVenctoDe
            lret = Comando_BindVarInt(lComando, .vdtCodDataVenctoDe)
            If (lret <> AD_SQL_SUCESSO) Then gError 99984
        End If

        'Se data final preenchida
        If (objRecebimentoCarne.dtDataVenctoAte <> DATA_NULA) Then
            .vdtCodDataVenctoAte = objRecebimentoCarne.dtDataVenctoAte
            lret = Comando_BindVarInt(lComando, .vdtCodDataVenctoAte)
            If (lret <> AD_SQL_SUCESSO) Then gError 99985
        End If

    End With

    lret = Comando_ExecutarInt(lComando)
    If (lret <> AD_SQL_SUCESSO) Then gError 99986

    RecebimentoCarne_Le_Parcelas3 = SUCESSO

    Exit Function

Erro_RecebimentoCarne_Le_Parcelas3:

    RecebimentoCarne_Le_Parcelas3 = gErr

    Select Case gErr

        Case 99982 To 99986
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CARNE_REC_BAIXADAS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151482)

    End Select

    Exit Function

End Function

Private Function RecebimentoCarne_Le_Parcelas4(ByVal lComando As Long, tInfoRecebimentoCarneVar As TypeInfoRecebimentoCarneVar, ByVal objRecebimentoCarne As ClassRecebimentoCarne) As Long
'Busca os dados do select que foi executado

Dim lErro As Long
Dim lret As Long
Dim objRecebimentoCarneParc As ClassRecebimentoCarneParc

On Error GoTo Erro_RecebimentoCarne_Le_Parcelas4

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 99987

    If lErro = AD_SQL_SEM_DADOS Then gError 99988

    Do While lErro = AD_SQL_SUCESSO

        Set objRecebimentoCarneParc = New ClassRecebimentoCarneParc

        With objRecebimentoCarneParc
            .sCarneNumero = tInfoRecebimentoCarneVar.vsCarneNumero
            .iParcelaNumero = tInfoRecebimentoCarneVar.viParcelaNumero
            .dtParcelaVencimento = tInfoRecebimentoCarneVar.vdtParcelaVencimento
            .dParcelaValor = tInfoRecebimentoCarneVar.vdParcelaValor
            .lParcelaNumIntDoc = tInfoRecebimentoCarneVar.vlParcelaNumIntDoc
            .lNumIntDoc = tInfoRecebimentoCarneVar.vlNumIntDoc
        End With
                
        objRecebimentoCarne.colParcelas.Add objRecebimentoCarneParc

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 99989

    Loop

    RecebimentoCarne_Le_Parcelas4 = SUCESSO

    Exit Function

Erro_RecebimentoCarne_Le_Parcelas4:

    RecebimentoCarne_Le_Parcelas4 = gErr

    Select Case gErr

        Case 99987, 99989
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CARNE_REC_BAIXADAS", gErr)

        Case 99988 'Sem dados

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151483)

    End Select

    Exit Function

End Function


'******************************************************'
'Feitos por cyntia ainda não colocados no SGEControle
'subir para tela de produtos enxergar
Function Gera_CodigoBarras_EAN13(sCodigoPais As String, sEmpresa As String, sCodigo As String)
'Padrão do EAN-13 789 - empresa - código - Dígito Verificador

Dim iDigito As Integer
Dim sCodigoEAN As String

    sCodigoEAN = sCodigoPais & sEmpresa & sCodigo
    
    iDigito = Calculo_DigitoVerificador_EAN13(sCodigoEAN)
        
    sCodigoEAN = sCodigoEAN & iDigito
    
    Gera_CodigoBarras_EAN13 = sCodigoEAN
    
End Function
'subir para a tela de produtos enxergar
Function Calculo_DigitoVerificador_EAN13(sCodigo As String)

Dim iIndice As Integer
Dim iDigito As Integer
Dim iDiv As Integer
Dim iSoma As Integer
Dim iSomaPar As Integer
Dim iSomaImpar As Integer
Dim bAchou As Boolean

    iDigito = 0
        
    'Para o cálculo do dígito Verificador temos q somar os índices pares e os ímpares
    For iIndice = 1 To Len(sCodigo)
      If iIndice Mod 2 = 0 Then
        iSomaPar = iSomaPar + (StrParaInt(Mid(sCodigo, iIndice, 1)))
      Else
        iSomaImpar = iSomaImpar + (StrParaInt(Mid(sCodigo, iIndice, 1)))
      End If
    Next
    
    'Depois, multiplica os pares por 3 e soma com os ímpares
    iSoma = (iSomaPar * 3) + iSomaImpar
    
    iDiv = iSoma + 1
    bAchou = False
    
    'Depois temos que verificar o número multiplo de 10 superior mais próximo
    Do While Not (bAchou)
        'Se for múltiplo
        If iDiv Mod 10 = 0 Then
            bAchou = True
        Else
            iDiv = iDiv + 1
        End If
    Loop
    
    'Subtrai a soma do múltiplo de 10
    iDigito = iDiv - iSoma
    
    'se o dígito for igual a 10 --> o dígito é 0.
    If iDigito = 10 Then iDigito = 0
    
    Calculo_DigitoVerificador_EAN13 = iDigito
    
End Function
'rotinas ecf
Function Caixa_Carne_Imprime_ECF(ByVal colCarneParcelasImpressao As Collection) As Long

Dim lErro As Long
Dim objCarneParcImpressao As New ClassCarneParcelasImpressao

On Error GoTo Erro_Caixa_Carne_Imprime_ECF

    'Para cada carnê
    For Each objCarneParcImpressao In colCarneParcelasImpressao
        'Imprime o código de barras
        lErro = Imprime_Carne(objCarneParcImpressao.sCodCarne)
        If lErro <> SUCESSO Then gError 109573
    Next
        
    Caixa_Carne_Imprime_ECF = SUCESSO

    Exit Function

Erro_Caixa_Carne_Imprime_ECF:

    Caixa_Carne_Imprime_ECF = gErr

    Select Case gErr

        Case 109573

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151484)

    End Select

    Exit Function
    
End Function
'rotinas ecf
Function Imprime_Carne(ByVal sCarne As String) As Long
'impressão pela impressora fiscal
Dim lErro As Long
Dim objCarne As ClassCarne
Dim objCarneParc As ClassCarneParcelas
Dim dValor As Double
Dim sMsg As String

On Error GoTo Erro_Imprime_Carne

    'Abri um fiscal não vinculado
    lErro = AFRAC_AbrirNaoFiscalNaoVinculado()
    If lErro <> SUCESSO Then
        lErro = CF_ECF("Retorna_MSGErro_AFRAC")
        If lErro <> SUCESSO Then gError 109571
    End If
    
    For Each objCarne In gcolCarne
    '? se sgeprinc --> le do bd.
        'Se é o carne passado
        If objCarne.sCodBarrasCarne = sCarne Then
        
            dValor = 0
            'Verifica o valor do carnê
            For Each objCarneParc In objCarne.colParcelas
                dValor = dValor + objCarneParc.dValor
            Next
            
            '???objCarne.lCliente
            '???/total da compra--> dvalor
            
            For Each objCarneParc In objCarne.colParcelas
            '? se status aberto-->imprime
                '???nº.:
                '??objcarneparc.iparcela
                'vencimento
                '???objCarneParc.dtDataVencimento
                'Valor
                '???objCarneParc.dValor
                'imprime código de barras
                Call CF_ECF("Imprime_Codigo_Code39", objCarne.sCodBarrasCarne & objCarneParc.iParcela)
            Next
            
            'imprime tudo de uma vez????????
            'começa a imprimir
            lErro = AFRAC_RegistrarNaoFiscal(TIPOMEIOPAGTOLOJA_CARNE - 1, dValor, sMsg)
            If lErro <> SUCESSO Then
                lErro = CF_ECF("Retorna_MSGErro_AFRAC")
                If lErro <> SUCESSO Then gError 109572
            End If
            
        End If
    Next
    
    'Fecha cupom
'    lErro = AFRAC_FecharCupom(False)
'    If lErro <> SUCESSO Then
'        lErro = CF_ECF("Retorna_MSGErro_AFRAC")
'        If lErro <> SUCESSO Then gError 109574
'    End If
    
    Exit Function
    
Erro_Imprime_Carne:

End Function

'rotinas ecf
Sub Imprime_Codigo_Code39(sCodigo As String)
    
    'Formata o Código para a fonte c39p48dmtt.ttf
    sCodigo = "*" & sCodigo & "*"
    
    'Imprimi neste momento, seleciona a fonte
    '??????
    
End Sub

'tela de impressão ainda não feita
Sub Imprime_Codigo_EAN13(sCodigo As String)
    
    'Formata o Código para a fonte
        
    'Imprimi neste momento, seleciona a fonte
    '??????
    
End Sub

'******************************************************'


Function Carne_Autentica(ByVal colCarneParcelasImpressao As Collection) As Long

Dim lErro As Long
Dim sLinha As String
Dim vbMsgRes As VbMsgBoxResult
Dim bAutentica As Boolean
Dim objCarneParcelasImpressao As ClassCarneParcelasImpressao
    
    'Para cada parcela --> verifica se deseja autenticar
    For Each objCarneParcelasImpressao In colCarneParcelasImpressao
    
        'Envia aviso perguntando se pode autenticar
        vbMsgRes = Rotina_Aviso(vbOKCancel, "AVISO_AUTENTICAR_CARNE", objCarneParcelasImpressao.iParcelaNumero, objCarneParcelasImpressao.sCodCarne)
        If vbMsgRes = vbOK Then
            'Autentica
            lErro = AFRAC_Autenticar(sLinha)
            If lErro <> SUCESSO Then
                lErro = CF("Retorna_MSGErro_AFRAC")
                If lErro <> SUCESSO Then gError 109678
            End If
            bAutentica = True
            Do While bAutentica
                'Envia aviso perguntando se deseja autenticar denovo
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_REAUTENTICAR_CARNE", objCarneParcelasImpressao.iParcelaNumero, objCarneParcelasImpressao.sCodCarne)
                
                If vbMsgRes = vbYes Then
                    'Envia aviso perguntando se pode autenticar
                    vbMsgRes = Rotina_Aviso(vbOKCancel, "AVISO_AUTENTICAR_CARNE", objCarneParcelasImpressao.iParcelaNumero, objCarneParcelasImpressao.sCodCarne)
                    If vbMsgRes = vbOK Then
                        'AutenticaAFRAC_RepetirAutenticacao
                        lErro = AFRAC_RepetirAutenticacao()
                        If lErro <> SUCESSO Then
                            lErro = CF("Retorna_MSGErro_AFRAC")
                            If lErro <> SUCESSO Then gError 109689
                        End If
                    Else
                        bAutentica = False
                    End If
                Else
                    bAutentica = False
                End If
            Loop
        End If
    Next
    Exit Function

Erro_Carne_Autentica:

    Select Case gErr

        Case 109678, 109689
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151485)

    End Select

    Exit Function
    
End Function

Public Function MovimentosCaixa_Le_Transf(ByVal objTransfCaixa As ClassTransfCaixa) As Long
'le os 2 movimentos de caixa que participam de uma transferencia em questao e retorna-os em objTransfCaixa
'Numa dada transferencia, existe sempre um movimento de entrada e um movimento de saida, nesta ORDEM

Dim lErro As Long
Dim lComando As Long
Dim tMovimentoCaixa As typeMovimentoCaixa
Dim objAux As ClassMovimentoCaixa

On Error GoTo Erro_MovimentosCaixa_Le_Transf

    'abre o comando
    lComando = Comando_Abrir
    If lComando = 0 Then gError 101991

    'inicializa as strings
    tMovimentoCaixa.sFavorecido = String(STRING_MOVIMENTOCAIXA_FAVORECIDO, 0)
    tMovimentoCaixa.sHistorico = String(STRING_MOVIMENTOCAIXA_HISTORICO, 0)

    'seleciona o registro com o filtro passado
    lErro = Comando_Executar(lComando, "SELECT FilialEmpresa, NumMovto, Caixa, Operador, Sequencial, Tipo, AdmMeioPagto, Parcelamento, TipoCartao, Numero, DataMovimento, Valor, Historico, Favorecido, CupomFiscal, NumRefInterna, MovtoTransf, MovtoEstorno, Gerente, CodConta, Hora, NumIntExt, NumIntDocLog, QuantLog " & _
                                       "FROM MovimentosCaixa " & _
                                       "WHERE FilialEmpresa = ? AND Transferencia = ? " & _
                                       "ORDER BY NumMovTo", _
                                       tMovimentoCaixa.iFilialEmpresa, tMovimentoCaixa.lNumMovto, tMovimentoCaixa.iCaixa, tMovimentoCaixa.iCodOperador, tMovimentoCaixa.lSequencial, tMovimentoCaixa.iTipo, tMovimentoCaixa.iAdmMeioPagto, tMovimentoCaixa.iParcelamento, tMovimentoCaixa.iTipoCartao, tMovimentoCaixa.lNumero, tMovimentoCaixa.dtDataMovimento, tMovimentoCaixa.dValor, tMovimentoCaixa.sHistorico, tMovimentoCaixa.sFavorecido, tMovimentoCaixa.lCupomFiscal, tMovimentoCaixa.lNumRefInterna, tMovimentoCaixa.lMovtoTransf, tMovimentoCaixa.lMovtoEstorno, tMovimentoCaixa.iGerente, tMovimentoCaixa.iCodConta, tMovimentoCaixa.dHora, tMovimentoCaixa.lNumIntExt, tMovimentoCaixa.lNumIntDocLog, tMovimentoCaixa.iQuantLog, _
                                       objTransfCaixa.iFilialEmpresa, objTransfCaixa.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 101992

    'busca o primeiro na seleção
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 101993

    'se não encontrar--> erro
    If lErro = AD_SQL_SEM_DADOS Then gError 101994

    'preenche o movimento de origem
    objTransfCaixa.objMovCaixaDe.iFilialEmpresa = tMovimentoCaixa.iFilialEmpresa
    objTransfCaixa.objMovCaixaDe.lNumMovto = tMovimentoCaixa.lNumMovto
    objTransfCaixa.objMovCaixaDe.iCaixa = tMovimentoCaixa.iCaixa
    objTransfCaixa.objMovCaixaDe.iCodOperador = tMovimentoCaixa.iCodOperador
    objTransfCaixa.objMovCaixaDe.lSequencial = tMovimentoCaixa.lSequencial
    objTransfCaixa.objMovCaixaDe.iTipo = tMovimentoCaixa.iTipo
    objTransfCaixa.objMovCaixaDe.iAdmMeioPagto = tMovimentoCaixa.iAdmMeioPagto
    objTransfCaixa.objMovCaixaDe.iParcelamento = tMovimentoCaixa.iParcelamento
    objTransfCaixa.objMovCaixaDe.iTipoCartao = tMovimentoCaixa.iTipoCartao
    objTransfCaixa.objMovCaixaDe.lNumero = tMovimentoCaixa.lNumero
    objTransfCaixa.objMovCaixaDe.dtDataMovimento = tMovimentoCaixa.dtDataMovimento
    objTransfCaixa.objMovCaixaDe.dValor = tMovimentoCaixa.dValor
    objTransfCaixa.objMovCaixaDe.sHistorico = tMovimentoCaixa.sHistorico
    objTransfCaixa.objMovCaixaDe.sFavorecido = tMovimentoCaixa.sFavorecido
    objTransfCaixa.objMovCaixaDe.lCupomFiscal = tMovimentoCaixa.lCupomFiscal
    objTransfCaixa.objMovCaixaDe.lNumRefInterna = tMovimentoCaixa.lNumRefInterna
    objTransfCaixa.objMovCaixaDe.lMovtoTransf = tMovimentoCaixa.lMovtoTransf
    objTransfCaixa.objMovCaixaDe.lMovtoEstorno = tMovimentoCaixa.lMovtoEstorno
    objTransfCaixa.objMovCaixaDe.iGerente = tMovimentoCaixa.iGerente
    objTransfCaixa.objMovCaixaDe.iCodConta = tMovimentoCaixa.iCodConta
    objTransfCaixa.objMovCaixaDe.dHora = tMovimentoCaixa.dHora
    objTransfCaixa.objMovCaixaDe.lNumIntExt = tMovimentoCaixa.lNumIntExt
    objTransfCaixa.objMovCaixaDe.lNumIntDocLog = tMovimentoCaixa.lNumIntDocLog
    objTransfCaixa.objMovCaixaDe.iQuantLog = tMovimentoCaixa.iQuantLog

    'busca o proximo
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 101996

    'se não encontrar--> erro
    If lErro = AD_SQL_SEM_DADOS Then gError 101997

    'preenche o movimento de destino
    objTransfCaixa.objMovCaixaPara.iFilialEmpresa = tMovimentoCaixa.iFilialEmpresa
    objTransfCaixa.objMovCaixaPara.lNumMovto = tMovimentoCaixa.lNumMovto
    objTransfCaixa.objMovCaixaPara.iCaixa = tMovimentoCaixa.iCaixa
    objTransfCaixa.objMovCaixaPara.iCodOperador = tMovimentoCaixa.iCodOperador
    objTransfCaixa.objMovCaixaPara.lSequencial = tMovimentoCaixa.lSequencial
    objTransfCaixa.objMovCaixaPara.iTipo = tMovimentoCaixa.iTipo
    objTransfCaixa.objMovCaixaPara.iAdmMeioPagto = tMovimentoCaixa.iAdmMeioPagto
    objTransfCaixa.objMovCaixaPara.iParcelamento = tMovimentoCaixa.iParcelamento
    objTransfCaixa.objMovCaixaPara.iTipoCartao = tMovimentoCaixa.iTipoCartao
    objTransfCaixa.objMovCaixaPara.lNumero = tMovimentoCaixa.lNumero
    objTransfCaixa.objMovCaixaPara.dtDataMovimento = tMovimentoCaixa.dtDataMovimento
    objTransfCaixa.objMovCaixaPara.dValor = tMovimentoCaixa.dValor
    objTransfCaixa.objMovCaixaPara.sHistorico = tMovimentoCaixa.sHistorico
    objTransfCaixa.objMovCaixaPara.sFavorecido = tMovimentoCaixa.sFavorecido
    objTransfCaixa.objMovCaixaPara.lCupomFiscal = tMovimentoCaixa.lCupomFiscal
    objTransfCaixa.objMovCaixaPara.lNumRefInterna = tMovimentoCaixa.lNumRefInterna
    objTransfCaixa.objMovCaixaPara.lMovtoTransf = tMovimentoCaixa.lMovtoTransf
    objTransfCaixa.objMovCaixaPara.lMovtoEstorno = tMovimentoCaixa.lMovtoEstorno
    objTransfCaixa.objMovCaixaPara.iGerente = tMovimentoCaixa.iGerente
    objTransfCaixa.objMovCaixaPara.iCodConta = tMovimentoCaixa.iCodConta
    objTransfCaixa.objMovCaixaPara.dHora = tMovimentoCaixa.dHora
    objTransfCaixa.objMovCaixaPara.lNumIntExt = tMovimentoCaixa.lNumIntExt
    objTransfCaixa.objMovCaixaPara.lNumIntDocLog = tMovimentoCaixa.lNumIntDocLog
    objTransfCaixa.objMovCaixaPara.iQuantLog = tMovimentoCaixa.iQuantLog

    MovimentosCaixa_Le_Transf = SUCESSO

    'fecha os comandos
    Call Comando_Fechar(lComando)

    Exit Function

Erro_MovimentosCaixa_Le_Transf:

    MovimentosCaixa_Le_Transf = gErr

    Select Case gErr

        Case 101991
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 101992, 101993, 101996
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOSCAIXA", gErr)

        Case 101994, 101997

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151486)

    End Select

    'fecha os comandos
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function Teclado_Verifica(ByVal objTecladoProduto As ClassTecladoProduto) As Long

Dim lErro As Long
Dim lComando As Long
Dim iCod As Integer

On Error GoTo Erro_Teclado_Verifica

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 109822

    'Verifica se exite outro tecladoproduto para esta marca de teclado
    lErro = Comando_Executar(lComando, "SELECT Codigo  FROM TecladoProduto WHERE Teclado = ? AND FilialEmpresa =? AND Padrao = ?", iCod, objTecladoProduto.iTeclado, giFilialEmpresa, TECLADO_COMUM)
    If lErro <> AD_SQL_SUCESSO Then gError 109818

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 109819

    'Se Não Encontrar Registros
    If lErro = AD_SQL_SUCESSO Then gError 109820
            
    Teclado_Verifica = SUCESSO
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
Erro_Teclado_Verifica:

    Teclado_Verifica = gErr
    
    Select Case gErr
    
        Case 109822
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 109818, 109819
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TECLADO", gErr, iCod)

        Case 109820
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151487)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function Carne_Le(ByVal objCarne As ClassCarne) As Long

Dim lErro As Long
Dim lComando As Long
Dim lCliente As Long
Dim iStatus As Integer
Dim dtData As Date
Dim lNumIntDoc As Long

On Error GoTo Erro_Carne_Le

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 109838

    'Seleciona código e nome das Carnes no Banco de Dados da Mesma FilialEmpresa
    lErro = Comando_Executar(lComando, "SELECT Cliente, DataReferencia, Status, NumIntDoc FROM Carne WHERE CodBarrasCarne = ? AND FilialEmpresa =? ", lCliente, dtData, iStatus, lNumIntDoc, objCarne.sCodBarrasCarne, giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 109839

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 109840

    'Se Não Encontrar Registros
    If lErro = AD_SQL_SEM_DADOS Then gError 109841

    objCarne.iStatus = iStatus
    objCarne.lCliente = lCliente
    objCarne.dtDataReferencia = dtData
    objCarne.lNumIntDoc = lNumIntDoc
    
    Carne_Le = SUCESSO
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
Erro_Carne_Le:

    Carne_Le = gErr
    
    Select Case gErr
    
        Case 109838
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 109839, 109840
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CARNE", gErr)

        Case 109841
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151488)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function


Function BaixasCarne_Le(ByVal objCarne As ClassCarne) As Long

Dim lErro As Long
Dim lComando As Long
Dim lNumIntDoc As Long

On Error GoTo Erro_BaixasCarne_Le

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 109846

    'Seleciona código e nome das Carnes no Banco de Dados da Mesma FilialEmpresa
    lErro = Comando_Executar(lComando, "SELECT NumIntBaixa FROM BaixasCarneLoja WHERE CodBarrasCarne = ? AND FilialEmpresa =? AND DataBaixa = ?", lNumIntDoc, objCarne.sCodBarrasCarne, giFilialEmpresa, objCarne.dtDataReferencia)
    If lErro <> AD_SQL_SUCESSO Then gError 109847

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 109848

    'Se Não Encontrar Registros
    If lErro = AD_SQL_SEM_DADOS Then gError 109849

    objCarne.lNumIntDoc = lNumIntDoc
    
    BaixasCarne_Le = SUCESSO
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
Erro_BaixasCarne_Le:

    BaixasCarne_Le = gErr
    
    Select Case gErr
    
        Case 109846
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 109847, 109848
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CARNE", gErr)

        Case 109849
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151489)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function MovimentosCaixaCarne_Lock(ByVal objMovimentosCaixa As ClassMovimentoCaixa) As Long
'Função que faz Lock no movimento de caixa correspondente ao Cheque

Dim lErro As Long
Dim tMovimentosCaixa As typeMovimentoCaixa
Dim lComando As Long

On Error GoTo Erro_MovimentosCaixaCarne_Lock

    'Abertura dos Comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 109854

    'Inicializa o Buffer das variáveis
    tMovimentosCaixa.sHistorico = String(STRING_HISTORICO1, 0)
    tMovimentosCaixa.sFavorecido = String(STRING_FAVORECIDO, 0)

    'Leitura da tabela de Carne fiscal para fim de LockShared
    lErro = Comando_ExecutarPos(lComando, "SELECT NumMovto, Caixa, Sequencial , AdmMeioPagto , Parcelamento , TipoCartao , Numero, DataMovimento , Valor, Historico , Favorecido ,NumRefInterna , MovtoTransf , MovtoEstorno , Gerente,Hora  FROM MovimentosCaixa WHERE Tipo = ? AND NumRefInterna = ? AND FilialEmpresa = ? ", 0, _
    tMovimentosCaixa.lNumMovto, tMovimentosCaixa.iCaixa, tMovimentosCaixa.lSequencial, tMovimentosCaixa.iAdmMeioPagto, tMovimentosCaixa.iParcelamento, tMovimentosCaixa.iTipoCartao, tMovimentosCaixa.lNumero, tMovimentosCaixa.dtDataMovimento, tMovimentosCaixa.dValor, tMovimentosCaixa.sHistorico, tMovimentosCaixa.sFavorecido, tMovimentosCaixa.lNumRefInterna, tMovimentosCaixa.lMovtoTransf, tMovimentosCaixa.lMovtoEstorno, tMovimentosCaixa.iGerente, tMovimentosCaixa.dHora, objMovimentosCaixa.iTipo, objMovimentosCaixa.lNumRefInterna, objMovimentosCaixa.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 109855

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 109856

    'SE Não encontrar Registro
    If lErro = AD_SQL_SEM_DADOS Then gError 109857

    'Faz um Lock no Registro Encontrado
    lErro = Comando_LockShared(lComando)
    If lErro <> SUCESSO Then gError 109858
    
    'Armazena o Resultado no objMovimentosCaixa
    objMovimentosCaixa.dtDataMovimento = tMovimentosCaixa.dtDataMovimento
    objMovimentosCaixa.dValor = tMovimentosCaixa.dValor
    objMovimentosCaixa.iAdmMeioPagto = tMovimentosCaixa.iAdmMeioPagto
    objMovimentosCaixa.iCaixa = tMovimentosCaixa.iCaixa
    objMovimentosCaixa.iGerente = tMovimentosCaixa.iGerente
    objMovimentosCaixa.iParcelamento = tMovimentosCaixa.iParcelamento
    objMovimentosCaixa.iTipoCartao = tMovimentosCaixa.iTipoCartao
    objMovimentosCaixa.lMovtoEstorno = tMovimentosCaixa.lMovtoEstorno
    objMovimentosCaixa.lMovtoTransf = tMovimentosCaixa.lMovtoTransf
    objMovimentosCaixa.lNumero = tMovimentosCaixa.lNumero
    objMovimentosCaixa.lNumMovto = tMovimentosCaixa.lNumMovto
    objMovimentosCaixa.lNumRefInterna = tMovimentosCaixa.lNumRefInterna
    objMovimentosCaixa.lSequencial = tMovimentosCaixa.lSequencial
    objMovimentosCaixa.sFavorecido = tMovimentosCaixa.sFavorecido
    objMovimentosCaixa.sHistorico = tMovimentosCaixa.sHistorico
    objMovimentosCaixa.dHora = tMovimentosCaixa.dHora

    'Fecha o comando
    Call Comando_Fechar(lComando)

    MovimentosCaixaCarne_Lock = SUCESSO

    Exit Function

Erro_MovimentosCaixaCarne_Lock:

     MovimentosCaixaCarne_Lock = gErr

    Select Case gErr

        Case 109854
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 109855, 109856
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOS_CAIXA1", gErr, objMovimentosCaixa.lNumMovto)

        Case 109857

        Case 109858
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_MOVIMENTOSCAIXA", gErr, objMovimentosCaixa.lNumMovto)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151490)

    End Select

     'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function


'Funções Realizadas para a Tela de RelOpRanking
'Sergio Dia 26/11/02

Function SlDiaFat_Le_ProdutosRanking(ByVal objProdutoRankingTela As ClassProdutosRankingTela, ByVal colProdutosRanking As Collection) As Long
'Função que monta o Select que será executado posteriormente

Dim lErro As Long
Dim sSelect As String
Dim tProdutoRanking As TypeProdutosRankingVar
Dim lComando As Long

On Error GoTo Erro_SlDiaFat_Le_ProdutosRanking

    'abre o comando
    lComando = Comando_Abrir
    If lComando = 0 Then gError 113066

    'Função que Monta o select
    lErro = SldDiaFat_Le_ProdutosRanking1(objProdutoRankingTela, sSelect)
    If lErro <> SUCESSO Then gError 113067

    'Função que Prepara a Parte Fixa do Select
    lErro = SldDiaFat_Le_ProdutosRanking2(lComando, sSelect, tProdutoRanking)
    If lErro <> SUCESSO Then gError 113068

    'Prepara os parâmetros que variam de acordo com a seleção do usuário
    lErro = SldDiaFat_Le_ProdutosRanking3(lComando, tProdutoRanking, objProdutoRankingTela, sSelect)
    If lErro <> SUCESSO Then gError 113069

    'Função que preenche a Coleção com os Registros retornados pelo Select
    lErro = SldDiaFat_Le_ProdutosRanking4(lComando, tProdutoRanking, colProdutosRanking, sSelect)
    If lErro <> SUCESSO Then gError 113070

    'fecha o comando
    Call Comando_Fechar(lComando)

    SlDiaFat_Le_ProdutosRanking = SUCESSO

    Exit Function

Erro_SlDiaFat_Le_ProdutosRanking:

    SlDiaFat_Le_ProdutosRanking = gErr

    Select Case gErr

        Case 113066
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 113067 To 113070

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151491)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function


Function SldDiaFat_Le_ProdutosRanking1(ByVal objProdutoRankingTela As ClassProdutosRankingTela, sSelect As String) As Long
'Função que Guarda na String sSelect o comando que será utilizada para ler os registro em SldDiaFat

Dim lErro As Long
Dim sFrom As String
Dim sWhere As String

On Error GoTo Erro_SldDiaFat_Le_ProdutosRanking1

    'Verifica se será utilizado filtro categoria
    If Len(Trim(objProdutoRankingTela.sCategoria)) <> 0 Then

        sSelect = "SELECT Produtos.Codigo , Produtos.NomeReduzido , MAX(ProdutoCategoria.Item) AS ItemCategoria , SUM(SldDiaFat.QuantFaturada) AS Quantidade , SUM(SldDiaFat.ValorFaturado) AS Valor"
        sFrom = " FROM Produtos , ProdutoCategoria , SldDiaFatCust AS SldDiaFat"
        sWhere = " WHERE  Produtos.Codigo = SldDiaFat.Produto AND Produtos.Codigo = ProdutoCategoria.Produto "
    
    Else
    
        sSelect = "SELECT Produtos.Codigo , Produtos.NomeReduzido , '' AS ItemCategoria , SUM(SldDiaFat.QuantFaturada) AS Quantidade , SUM(SldDiaFat.ValorFaturado) AS Valor"
        sFrom = " FROM Produtos, SldDiaFatCust AS SldDiaFat"
        sWhere = " WHERE  Produtos.Codigo = SldDiaFat.Produto "
    
    End If

    'Verifica se o Filtro utilizado vai ser por filial ou não
    If objProdutoRankingTela.iFilialEmpresa <> EMPRESA_TODA Then

        sWhere = sWhere & " AND SldDiaFat.FilialEmpresa = ? "

    End If

    'Verifica se a Data De é diferente de data Nula se for  será utilizada como Filtro
    If objProdutoRankingTela.dtDataDe <> DATA_NULA Then

        sWhere = sWhere & " AND SldDiaFat.Data >= ? "

    End If

    'Verifica se a Data Até é diferente de data Nula se for  será utilizada como Filtro
    If objProdutoRankingTela.dtDataAte <> DATA_NULA Then

        sWhere = sWhere & " AND SldDiaFat.Data <= ? "

    End If

    'Verifica se será utilizado filtro por tipo de produto
    If objProdutoRankingTela.iTipoProdutoDe <> 0 Then

        sWhere = sWhere & " AND Produtos.Tipo >= ? "

    End If

    'Verifica se será utilizado filtro por tipo de produto
    If objProdutoRankingTela.iTipoProdutoAte <> 0 Then

        sWhere = sWhere & " AND Produtos.Tipo <= ? "

    End If

    'Verifica se será utilizado filtro por produto
    If objProdutoRankingTela.sProdutoDe <> "" Then

        sWhere = sWhere & " AND Produtos.Codigo >= ? "

    End If

    'Verifica se será utilizado filtro por tipo de produto
    If objProdutoRankingTela.sProdutoAte <> "" Then

        sWhere = sWhere & " AND Produtos.Codigo <= ? "

    End If

    'Verifica se será utilizado filtro categoria
    If Len(Trim(objProdutoRankingTela.sCategoria)) <> 0 Then

        sWhere = sWhere & " AND ProdutoCategoria.Categoria = ? "

    End If

    sSelect = sSelect & sFrom & sWhere & "Group By (Codigo),(Nomereduzido)"
    
    SldDiaFat_Le_ProdutosRanking1 = SUCESSO

    Exit Function

Erro_SldDiaFat_Le_ProdutosRanking1:

    SldDiaFat_Le_ProdutosRanking1 = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151492)

    End Select

    Exit Function

End Function

Private Function SldDiaFat_Le_ProdutosRanking2(ByVal lComando As Long, sSelect As String, tProdutoRanking As TypeProdutosRankingVar) As Long
'Função que Bind o as variáveis que serão recebidas pelo select

Dim lErro As Long

On Error GoTo Erro_SldDiaFat_Le_ProdutosRanking2

    With tProdutoRanking

        .vsProduto = String(STRING_PRODUTO, 0)
        .vsNomeReduzido = String(STRING_PRODUTO_NOME_REDUZIDO, 0)
        .vsItemCategoria = String(STRING_CATEGORIAPRODUTOITEM_ITEM, 0)
        
        lErro = Comando_PrepararInt(lComando, sSelect)
        If (lErro <> AD_SQL_SUCESSO) Then gError 113075

        lErro = Comando_BindVarInt(lComando, .vsProduto)
        If (lErro <> AD_SQL_SUCESSO) Then gError 113076


        lErro = Comando_BindVarInt(lComando, .vsNomeReduzido)
        If (lErro <> AD_SQL_SUCESSO) Then gError 113077

        lErro = Comando_BindVarInt(lComando, .vsItemCategoria)
        If (lErro <> AD_SQL_SUCESSO) Then gError 113078

        .vdQuantidade = CDbl(.vdQuantidade)
        lErro = Comando_BindVarInt(lComando, .vdQuantidade)
        If (lErro <> AD_SQL_SUCESSO) Then gError 113079

        .vdValor = CDbl(.vdValor)
        lErro = Comando_BindVarInt(lComando, .vdValor)
        If (lErro <> AD_SQL_SUCESSO) Then gError 113080

    End With

    SldDiaFat_Le_ProdutosRanking2 = SUCESSO

    Exit Function

Erro_SldDiaFat_Le_ProdutosRanking2:

    SldDiaFat_Le_ProdutosRanking2 = gErr

    Select Case gErr

        Case 113075 To 113080
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXECUCAO_COMANDO_SQL", gErr, sSelect)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151493)

    End Select

    Exit Function

End Function

Private Function SldDiaFat_Le_ProdutosRanking3(ByVal lComando As Long, tProdutoRanking As TypeProdutosRankingVar, ByVal objProdutoRankingTela As ClassProdutosRankingTela, sSelect As String)
'Função que Bind os filtros passados pelo usuário

Dim lErro As Long

On Error GoTo Erro_SldDiaFat_Le_ProdutosRanking3

    'Verifica se o filtro é por filial empresa
    If objProdutoRankingTela.iFilialEmpresa <> EMPRESA_TODA Then

        tProdutoRanking.viFilialEmpresa = CInt(giFilialEmpresa)
        lErro = Comando_BindVarInt(lComando, tProdutoRanking.viFilialEmpresa)
        If (lErro <> AD_SQL_SUCESSO) Then gError 113081


    End If

    'Verifica se o filtro também será por data
    If objProdutoRankingTela.dtDataDe <> DATA_NULA Then

        tProdutoRanking.vdtDataDe = CDate(objProdutoRankingTela.dtDataDe)
        lErro = Comando_BindVarInt(lComando, tProdutoRanking.vdtDataDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError 113082

    End If

    'Verifica se o filtro também será por data
    If objProdutoRankingTela.dtDataAte <> DATA_NULA Then

        tProdutoRanking.vdtDataAte = CDate(objProdutoRankingTela.dtDataAte)
        lErro = Comando_BindVarInt(lComando, tProdutoRanking.vdtDataAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError 113083

    End If

    'Verifica se o filtro também será por Tipo de Produto
    If objProdutoRankingTela.iTipoProdutoDe <> 0 Then

        tProdutoRanking.viTipoProdutoDe = CInt(objProdutoRankingTela.iTipoProdutoDe)
        lErro = Comando_BindVarInt(lComando, tProdutoRanking.viTipoProdutoDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError 113084

    End If

    'Verifica se o filtro também será por Tipo de Produto
    If objProdutoRankingTela.iTipoProdutoAte <> 0 Then

        tProdutoRanking.viTipoProdutoAte = CInt(objProdutoRankingTela.iTipoProdutoAte)
        lErro = Comando_BindVarInt(lComando, tProdutoRanking.viTipoProdutoAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError 113085

    End If

    'Verifica se o filtro também será por Produto
    If objProdutoRankingTela.sProdutoDe <> "" Then

        tProdutoRanking.vsProdutoDe = CStr(objProdutoRankingTela.sProdutoDe)
        lErro = Comando_BindVarInt(lComando, tProdutoRanking.vsProdutoDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError 113084

    End If

    'Verifica se o filtro também será por Produto
    If objProdutoRankingTela.sProdutoAte <> "" Then

        tProdutoRanking.vsProdutoAte = CStr(objProdutoRankingTela.sProdutoAte)
        lErro = Comando_BindVarInt(lComando, tProdutoRanking.vsProdutoAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError 113085

    End If

    'Verifica se o filtro também será por Categoria
    If Len(Trim(objProdutoRankingTela.sCategoria)) <> 0 Then

        tProdutoRanking.vsCategoria = CStr(objProdutoRankingTela.sCategoria)
        lErro = Comando_BindVarInt(lComando, tProdutoRanking.vsCategoria)
        If (lErro <> AD_SQL_SUCESSO) Then gError 113086

    End If

    lErro = Comando_ExecutarInt(lComando)
    If (lErro <> AD_SQL_SUCESSO) Then gError 113087

    SldDiaFat_Le_ProdutosRanking3 = SUCESSO

    Exit Function

Erro_SldDiaFat_Le_ProdutosRanking3:

    SldDiaFat_Le_ProdutosRanking3 = gErr

    Select Case gErr

        Case 113081 To 113086
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXECUCAO_COMANDO_SQL", gErr, sSelect)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151494)

    End Select

    Exit Function

End Function

Private Function SldDiaFat_Le_ProdutosRanking4(ByVal lComando As Long, tProdutoRanking As TypeProdutosRankingVar, ByVal colProdutosRanking As Collection, sSelect As String) As Long
'Busca no banco de dados e acumula as estimativas dos produtos lidos( como Valor Faturado , Quantidade faturado , a que item este produto corresponde na categoria )

Dim lErro As Long
Dim objProdutosRanking As ClassProdutosRanking

On Error GoTo Erro_SldDiaFat_Le_ProdutosRanking4

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 113088

    If lErro = AD_SQL_SEM_DADOS Then gError 113089

    Do While lErro = AD_SQL_SUCESSO

        Set objProdutosRanking = New ClassProdutosRanking

        With objProdutosRanking

            .sProduto = tProdutoRanking.vsProduto
            .dQuantidade = tProdutoRanking.vdQuantidade
            .dValor = tProdutoRanking.vdValor
            .sNomeReduzido = tProdutoRanking.vsNomeReduzido
            .sItemCategoria = tProdutoRanking.vsItemCategoria
        End With

        colProdutosRanking.Add objProdutosRanking

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 113090

    Loop

    SldDiaFat_Le_ProdutosRanking4 = SUCESSO

    Exit Function

Erro_SldDiaFat_Le_ProdutosRanking4:

    SldDiaFat_Le_ProdutosRanking4 = gErr

    Select Case gErr

        Case 113088 To 113090
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXECUCAO_COMANDO_SQL", gErr, sSelect)

        Case 113089
            'Só desvia o Código sem Msg

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151495)

    End Select

    Exit Function

End Function

Function ProdutosRanking_Preenche(lNumIntRel As Long, objProdutoRankingTela As ClassProdutosRankingTela) As Long

Dim lErro As Long
Dim iRanking As Integer
Dim dQuantTotal As Double
Dim dValorTotal As Double
Dim objProdutoRanking As ClassProdutosRanking
Dim colProdutosRanking As New Collection
Dim dQuantProdPerc As Double
Dim dValorProdPerc As Double

On Error GoTo Erro_ProdutosRanking_Preenche

    'Função que Lê os dados que serão inseridos em produtos Ranking
    lErro = CF("SlDiaFat_Le_ProdutosRanking", objProdutoRankingTela, colProdutosRanking)
    If lErro <> SUCESSO Then gError 113064

    'inicializar iRanking
    iRanking = 1

    'Acumula a Qdt Total de produtos vendidos e o Valor Total de produtos vendido
    For Each objProdutoRanking In colProdutosRanking

        dQuantTotal = dQuantTotal + objProdutoRanking.dQuantidade
        dValorTotal = dValorTotal + objProdutoRanking.dValor

    Next

    'guarda as Estatísticas de venda de cada produto no Obj
    For Each objProdutoRanking In colProdutosRanking

        objProdutoRanking.lRanking = iRanking
        dQuantProdPerc = ((objProdutoRanking.dQuantidade * 100) / dQuantTotal)
        objProdutoRanking.dQuantPercParticip = dQuantProdPerc
        dValorProdPerc = ((objProdutoRanking.dValor * 100) / dValorTotal)
        objProdutoRanking.dVlrPercParticip = dValorProdPerc
        iRanking = iRanking + 1

    Next
    
    'Função que grava os calculos de relatório
    lErro = CF("ProdutosRanking_Grava", colProdutosRanking, lNumIntRel)
    If lErro <> SUCESSO Then gError 113093
    
    ProdutosRanking_Preenche = SUCESSO

    Exit Function

Erro_ProdutosRanking_Preenche:

    ProdutosRanking_Preenche = gErr

    Select Case gErr
        
        Case 113064, 113093

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151496)

    End Select

    Exit Function

End Function




'Fim Sergio .....

'Função Utilizada na Tela de Desconto Item
'Função que Lê os Produtos Vinculado a Uma Filial e o Desconto para cada Produto
'Sergio dia 14/01/03

Function DescProdutosFilial_Le(ByVal objProduto As ClassProduto, ByVal objProdutoFilial As ClassProdutoFilial) As Long
'Lê os dados do produto e o desconto, usando o código do produto passado como parâmetro
'Filial empresa (gifilialEmpresa)

Dim lErro As Long
Dim lComando As Long
Dim tProduto As typeProduto
Dim tProdFilial As typeProdutoFilial

On Error GoTo Erro_DescProdutosFilial_Le

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 113314

    'Inicializa as variáveis do tipo String
    With tProduto
        
        .sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
        .sNomeReduzido = String(STRING_PRODUTO_NOME_REDUZIDO, 0)
        
        'Lê o Produto
        lErro = Comando_Executar(lComando, "SELECT Produtos.Descricao , Produtos.NomeReduzido , ProdutosFilial.DescontoItem, ProdutosFilial.DescontoValor FROM Produtos , ProdutosFilial WHERE ProdutosFilial.Produto = ? AND Produtos.Codigo = ProdutosFilial.Produto AND ProdutosFilial.FilialEmpresa = ? ", _
        .sDescricao, .sNomeReduzido, tProdFilial.dDescontoItem, tProdFilial.dDescontoValor, objProduto.sCodigo, giFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 113315

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 113316

        'Se não encontrou o Produto --> Erro
        If lErro = AD_SQL_SEM_DADOS Then gError 113313

        'Coloca os dados lidos no objProduto
        objProduto.sDescricao = .sDescricao
        objProduto.sNomeReduzido = .sNomeReduzido
        objProdutoFilial.dDescontoItem = tProdFilial.dDescontoItem
        objProdutoFilial.dDescontoValor = tProdFilial.dDescontoValor
        
    End With

    'Fecha o comando
    Call Comando_Fechar(lComando)

    DescProdutosFilial_Le = SUCESSO

    Exit Function

Erro_DescProdutosFilial_Le:

    DescProdutosFilial_Le = gErr

    Select Case gErr

        Case 113313
            'Produto não encontrado. Será tratado na rotina chamadora.

        Case 113314
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 113315, 113316
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 151497)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ChequeBordero_Le_NumIntCheque(ByVal objChequeBordero As ClassChequeBordero) As Long
'Le a associacao de cheque com bordero

Dim lComando1 As Long
Dim lErro As Long
Dim lNumIntDoc As Long
Dim lNumBordero As Long
Dim iTipoBordero As Integer
Dim iFilialEmpresa As Integer

On Error GoTo Erro_ChequeBordero_Le_NumIntCheque

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 110054
    
    'Pega os dados da tabela DevolucaoCheque
    lErro = Comando_Executar(lComando1, "SELECT NumIntDoc, NumBordero, TipoBordero, FilialEmpresa FROM ChequeBordero WHERE NumIntCheque = ?", lNumIntDoc, lNumBordero, iTipoBordero, iFilialEmpresa, objChequeBordero.lNumIntCheque)
    If lErro <> AD_SQL_SUCESSO Then gError 110055

    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 110056

    If lErro = AD_SQL_SEM_DADOS Then gError 110057
     
    objChequeBordero.lNumIntDoc = lNumIntDoc
    objChequeBordero.lNumBordero = lNumBordero
    objChequeBordero.iTipoBordero = iTipoBordero
    objChequeBordero.iFilialEmpresa = iFilialEmpresa
    
    Call Comando_Fechar(lComando1)

    ChequeBordero_Le_NumIntCheque = SUCESSO

    Exit Function

Erro_ChequeBordero_Le_NumIntCheque:

    ChequeBordero_Le_NumIntCheque = gErr

    Select Case gErr

        Case 110054
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 110055, 110056
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CHEQUEBORDERO", gErr)

        Case 110057

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151498)

    End Select

    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Function BaixasCarneDetalhes_Le(objBaixasCarne As ClassBaixasCarne) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim colParcelas As New Collection
Dim objCarneParcelasImpressao As ClassCarneParcelasImpressao
Dim lComando As Long
Dim tCarneParcelasImpressao As TypeCarneParcelasImpressao

On Error GoTo Erro_BaixasCarneDetalhes_Le

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 109706
        
    With tCarneParcelasImpressao
    
    .sCodCarne = String(Carne_sCarne, 0)
    
    'Le as parcelas q fazem parte da baixa
    lErro = Comando_Executar(lComando, "SELECT CodBarrasCarne, Cliente, Parcela, DataVencimento, Valor, ValorBaixado, ValorDesconto, ValorMulta, ValorJuros FROM BaixasCarneLoja  WHERE  Codigo = ? AND FilialEmpresa = ?", .sCodCarne, .lCodCliente, .iCarneNumero, .dtParcelaVencimento, .dValor, .dValorBaixado, .dDesconto, .dMulta, .dJuros, objBaixasCarne.lCodigo, giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 109707

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 109708

    'Se não encontrar Dados
    If lErro = AD_SQL_SEM_DADOS Then gError 109704
    
    'Enquanto tiver parcelas...
    Do While lErro = AD_SQL_SUCESSO
        
        Set objCarneParcelasImpressao = New ClassCarneParcelasImpressao
        
        'Encontrando Preencher objAdmMeioPagto
        objCarneParcelasImpressao.dDesconto = .dDesconto
        objCarneParcelasImpressao.dJuros = .dJuros
        objCarneParcelasImpressao.dMulta = .dMulta
        objCarneParcelasImpressao.dParcelaValor = .dValor
        objCarneParcelasImpressao.dtDataVencParcela = .dtParcelaVencimento
        objCarneParcelasImpressao.dValorBaixado = .dValorBaixado
        objCarneParcelasImpressao.iParcelaNumero = .iCarneNumero
        objCarneParcelasImpressao.lCodCliente = .lCodCliente
        objCarneParcelasImpressao.sCodCarne = .sCodCarne
        
        objBaixasCarne.colParcelas.Add objCarneParcelasImpressao
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 109709
    
    Loop
    
    End With
    
    'Fecha o comando
    Call Comando_Fechar(lComando)

    BaixasCarneDetalhes_Le = SUCESSO

    Exit Function

Erro_BaixasCarneDetalhes_Le:

    BaixasCarneDetalhes_Le = gErr

    Select Case gErr

        Case 109706
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 109707, 109708, 109709
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BAIXASCARNELOJA", gErr)

        Case 109704

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151499)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function
    
End Function


Function VendedorAtivo_Le_Todos(colVendedor As Collection) As Long
'Lê todos os Vendedores Ativos da Empresa

Dim lErro As Long
Dim lComando As Long
Dim tVendedor As typeVendedor
Dim objVendedor As ClassVendedor

On Error GoTo Erro_VendedorAtivo_Le_Todos

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 104284
    
    'Inicialização do Buffer
    tVendedor.sNomeReduzido = String(STRING_VENDEDOR_NOME_RED, 0)
    tVendedor.sAgencia = String(STRING_VENDEDOR_AGENCIA, 0)
    tVendedor.sContaCorrente = String(STRING_VENDEDOR_CONTA_CORRENTE, 0)
    tVendedor.sMatricula = String(STRING_VENDEDOR_MATRICULA, 0)
    tVendedor.sNome = String(STRING_VENDEDOR_NOME, 0)
    tVendedor.sCodUsuario = String(STRING_USUARIO_CODIGO, 0)
    
    'Seleciona todos os Vendedores Ativos
    lErro = Comando_Executar(lComando, "SELECT Codigo, NomeReduzido, Nome, Endereco, Tipo, Matricula, CodRegiao, SaldoComissao, PercComissao, PercComissaoBaixa, PercComissaoEmissao, ComissaoSobreTotal, ComissaoFrete, ComissaoICM, ComissaoIPI, ComissaoSeguro, Banco, Agencia, ContaCorrente, Ativo, CodUsuario " & _
                                       "FROM Vendedores WHERE Ativo = ?", _
                                       tVendedor.iCodigo, tVendedor.sNomeReduzido, tVendedor.sNome, tVendedor.lEndereco, tVendedor.iTipo, tVendedor.sMatricula, tVendedor.iCodRegiao, tVendedor.dSaldoComissao, tVendedor.dPercComissao, tVendedor.dPercComissaoBaixa, tVendedor.dPercComissaoEmissao, tVendedor.iComissaoSobreTotal, tVendedor.iComissaoFrete, tVendedor.iComissaoICM, tVendedor.iComissaoIPI, tVendedor.iComissaoSeguro, tVendedor.iBanco, tVendedor.sAgencia, tVendedor.sContaCorrente, tVendedor.iAtivo, tVendedor.sCodUsuario, VENDEDOR_ATIVO)
    If lErro <> AD_SQL_SUCESSO Then gError 104285

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 104286

    If lErro = AD_SQL_SEM_DADOS Then gError 104287
    
    Do While lErro = AD_SQL_SUCESSO

        'Guarda na coleção os nomes dos usuarios que têm operadores
        Set objVendedor = New ClassVendedor

        objVendedor.iCodigo = tVendedor.iCodigo
        objVendedor.sNomeReduzido = tVendedor.sNomeReduzido
        objVendedor.sNome = tVendedor.sNome
        objVendedor.lEndereco = tVendedor.lEndereco
        objVendedor.iTipo = tVendedor.iTipo
        objVendedor.sMatricula = tVendedor.sMatricula
        objVendedor.iCodRegiao = tVendedor.iCodRegiao
        objVendedor.dSaldoComissao = tVendedor.dSaldoComissao
        objVendedor.dPercComissao = tVendedor.dPercComissao
        objVendedor.dPercComissaoBaixa = tVendedor.dPercComissaoBaixa
        objVendedor.dPercComissaoEmissao = tVendedor.dPercComissaoEmissao
        objVendedor.iComissaoSobreTotal = tVendedor.iComissaoSobreTotal
        objVendedor.iComissaoFrete = tVendedor.iComissaoFrete
        objVendedor.iComissaoICM = tVendedor.iComissaoICM
        objVendedor.iComissaoIPI = tVendedor.iComissaoIPI
        objVendedor.iComissaoSeguro = tVendedor.iComissaoSeguro
        objVendedor.iBanco = tVendedor.iBanco
        objVendedor.sAgencia = tVendedor.sAgencia
        objVendedor.sContaCorrente = tVendedor.sContaCorrente
        objVendedor.iAtivo = tVendedor.iAtivo
        objVendedor.sCodUsuario = tVendedor.sCodUsuario
        
        colVendedor.Add objVendedor

        'le as filiais para as quais o vendedor pode atuar
        lErro = CF("VendedorLoja_Le_FilalEmpresa", objVendedor)
        If lErro <> SUCESSO And lErro <> 107503 Then gError 105930

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 104314

    Loop

    Call Comando_Fechar(lComando)

    VendedorAtivo_Le_Todos = SUCESSO

    Exit Function

Erro_VendedorAtivo_Le_Todos:

    VendedorAtivo_Le_Todos = gErr

    Select Case gErr

        Case 104284
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 104285, 104286, 104314
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_VENDEDORES", gErr)

        Case 104287
            'Só Desvia o Código não Emite Msg de Erro
        
        Case 105930
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151500)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Cliente_Le_Loja(ByVal objCliente As ClassCliente) As Long
'le o cliente ( e sua matriz ) a partir da chave(lCodigo) fornecida em objCliente. Devolve os dados em objCliente

Dim lComando As Long
Dim lComando2 As Long
Dim lErro As Long
Dim tCliente As typeCliente
Dim lCodigoCliente As Long

On Error GoTo Erro_Cliente_Le_Loja

    tCliente.sRazaoSocial = String(STRING_CLIENTE_RAZAO_SOCIAL, 0)
    tCliente.sNomeReduzido = String(STRING_CLIENTE_NOME_REDUZIDO, 0)
    tCliente.sNome = String(STRING_FILIAL_CLIENTE_NOME, 0)
    tCliente.sCgc = String(STRING_CGC, 0)
    tCliente.sRG = String(STRING_RG, 0)
    
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 112600

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then gError 112601
        
    'faço o select com codigoloja e FilialEmpresaloja
    lErro = Comando_Executar(lComando, "SELECT CodigoLoja, Codigo, RazaoSocial, NomeReduzido, Tipo, CGC, RG, Endereco FROM Clientes, FiliaisClientes WHERE CodigoLoja = ? AND Clientes.FilialEmpresaLoja=? AND Clientes.Codigo=FiliaisClientes.CodCliente", _
    tCliente.lCodigoLoja, tCliente.lCodigo, tCliente.sRazaoSocial, tCliente.sNomeReduzido, tCliente.iTipo, tCliente.sCgc, tCliente.sRG, tCliente.lEndereco, objCliente.lCodigoLoja, objCliente.iFilialEmpresaLoja)
    If lErro <> AD_SQL_SUCESSO Then gError 112602
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 112603
    
    If lErro = AD_SQL_SEM_DADOS Then
        
        tCliente.sRazaoSocial = String(STRING_CLIENTE_RAZAO_SOCIAL, 0)
        tCliente.sNomeReduzido = String(STRING_CLIENTE_NOME_REDUZIDO, 0)
        tCliente.sNome = String(STRING_FILIAL_CLIENTE_NOME, 0)
        tCliente.sCgc = String(STRING_CGC, 0)
        tCliente.sRG = String(STRING_RG, 0)
        
        'faço o select com codigo
        lErro = Comando_Executar(lComando2, "SELECT CodigoLoja, Codigo, RazaoSocial, NomeReduzido, Tipo, CGC, RG, Endereco FROM Clientes, FiliaisClientes WHERE Codigo = ? AND Clientes.Codigo=FiliaisClientes.CodCliente", _
        tCliente.lCodigoLoja, tCliente.lCodigo, tCliente.sRazaoSocial, tCliente.sNomeReduzido, tCliente.iTipo, tCliente.sCgc, tCliente.sRG, tCliente.lEndereco, objCliente.lCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 112604
        
        lErro = Comando_BuscarPrimeiro(lComando2)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 112605
    
        If lErro = AD_SQL_SEM_DADOS Then gError 112606
        
        objCliente.lCodigoLoja = tCliente.lCodigoLoja
        objCliente.iFilialEmpresaLoja = tCliente.iFilialEmpresaLoja
       
    End If
    
    objCliente.lCodigo = tCliente.lCodigo
    objCliente.sRazaoSocial = tCliente.sRazaoSocial
    objCliente.sNomeReduzido = tCliente.sNomeReduzido
    objCliente.iTipo = tCliente.iTipo
    objCliente.sCgc = tCliente.sCgc
    objCliente.sRG = tCliente.sRG
    objCliente.lEndereco = tCliente.lEndereco
    
    lErro = Comando_Fechar(lComando)
    lErro = Comando_Fechar(lComando2)

    Cliente_Le_Loja = SUCESSO

    Exit Function

Erro_Cliente_Le_Loja:

    Cliente_Le_Loja = gErr

    Select Case gErr

        Case 112600, 112601
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 112602 To 112605
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLIENTES", gErr)

        Case 112606 'cliente nao cadastrado em Clientes

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151501)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)

    Exit Function

End Function

Function FilialCliente_Le_Loja(ByVal objFilialCliente As ClassFilialCliente) As Long
'Le todos os dados da tabela FiliaisClientes, correspondente a chave passada por parametro

Dim lComando As Long
Dim lComando2 As Long
Dim lErro As Long
Dim tFilialCliente As typeFilialCliente
Dim iCodFilialLoja As Integer
Dim iFilialEmpresaLoja As Integer
Dim lCodClienteLoja As Long

On Error GoTo Erro_FilialCliente_Le_Loja

    tFilialCliente.sNome = String(STRING_FILIAL_CLIENTE_NOME, 0)
    tFilialCliente.sRG = String(STRING_RG, 0)
    tFilialCliente.sCgc = String(STRING_CGC, 0)
    
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 112608
    
    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then gError 112609
    
    'Pega os dados da tabela Filiais Clientes
    lErro = Comando_Executar(lComando, "SELECT Nome, CGC, Endereco, RG, CodFilialLoja, CodFilial, FilialEmpresaLoja, CodClienteLoja FROM FiliaisClientes WHERE CodClienteLoja = ? AND CodFilialLoja = ? AND FilialEmpresaLoja=?", _
    tFilialCliente.sNome, tFilialCliente.sCgc, tFilialCliente.lEndereco, tFilialCliente.sRG, iCodFilialLoja, tFilialCliente.iCodFilial, iFilialEmpresaLoja, lCodClienteLoja, objFilialCliente.lCodClienteLoja, objFilialCliente.iCodFilialLoja, objFilialCliente.iFilialEmpresaLoja)
    If lErro <> AD_SQL_SUCESSO Then gError 112610

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 112611
  
    If lErro = AD_SQL_SEM_DADOS Then
        
        tFilialCliente.sNome = String(STRING_FILIAL_CLIENTE_NOME, 0)
        tFilialCliente.sRG = String(STRING_RG, 0)
        tFilialCliente.sCgc = String(STRING_CGC, 0)
        
        'Pega os dados da tabela Filiais Clientes
        lErro = Comando_Executar(lComando2, "SELECT Nome, CGC, Endereco, RG, CodFilialLoja, CodFilial, FilialEmpresaLoja, CodClienteLoja FROM FiliaisClientes WHERE CodCliente = ? AND CodFilial = ?", _
        tFilialCliente.sNome, tFilialCliente.sCgc, tFilialCliente.lEndereco, tFilialCliente.sRG, iCodFilialLoja, tFilialCliente.iCodFilial, iFilialEmpresaLoja, lCodClienteLoja, objFilialCliente.lCodCliente, objFilialCliente.iCodFilial)
        If lErro <> AD_SQL_SUCESSO Then gError 112612
    
        lErro = Comando_BuscarPrimeiro(lComando2)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 112613
        
        If lErro = AD_SQL_SEM_DADOS Then gError 112607
        
    End If
        
    objFilialCliente.sNome = tFilialCliente.sNome
    objFilialCliente.sCgc = tFilialCliente.sCgc
    objFilialCliente.sRG = tFilialCliente.sRG
    objFilialCliente.lEndereco = tFilialCliente.lEndereco
    objFilialCliente.iCodFilial = tFilialCliente.iCodFilial
    objFilialCliente.iCodFilialLoja = iCodFilialLoja
    objFilialCliente.iFilialEmpresaLoja = iFilialEmpresaLoja
    objFilialCliente.lCodClienteLoja = lCodClienteLoja
    
    lErro = Comando_Fechar(lComando)
    lErro = Comando_Fechar(lComando2)
    
    FilialCliente_Le_Loja = SUCESSO

    Exit Function

Erro_FilialCliente_Le_Loja:

    FilialCliente_Le_Loja = gErr

    Select Case gErr

        Case 112608
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 112609, 112610
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISCLIENTES", gErr)

        Case 112607 'nao encontrou a filial (sem dados)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151502)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)

    Exit Function

End Function


Function ModeloECF_Le(ByVal objModeloECF As ClassModeloECF) As Long
'Lê os dados do ModeloECF_Le, cujo código foi passado como parâmetro

Dim lErro As Long
Dim lComando As Long
Dim tModeloECF As typeModeloECF

On Error GoTo Erro_ModeloECF_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 103462

    tModeloECF.sNome = String(STRING_MODELOECF_NOME, 0)
    tModeloECF.sFabricante = String(STRING_MODELOECF_FABRICANTE, 0)

    'Busca o ModeloECF no BD através do código e filial passados como parâmetro
    lErro = Comando_Executar(lComando, "SELECT Nome, Fabricante FROM ModeloECF WHERE Codigo = ?", tModeloECF.sNome, tModeloECF.sFabricante, objModeloECF.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 103463

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 103464

    'Se não encontrou o ImpressoraECF => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 103459

    'Transfere para objImpressoraECF os dados obtidos do BD
    objModeloECF.sFabricante = tModeloECF.sFabricante
    objModeloECF.sNome = tModeloECF.sNome

    'Fecha o comando
    Call Comando_Fechar(lComando)

    ModeloECF_Le = SUCESSO

    Exit Function

Erro_ModeloECF_Le:

    ModeloECF_Le = gErr

    Select Case gErr

        Case 103462
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 103463, 103464
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MODELOECF", gErr)

        Case 103459
            'ModeloECF não encontrado. Esse erro será tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151503)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function


Function ImpressoraECF_Le(ByVal objImpressoraECF As ClassImpressoraECF) As Long
'Lê os dados da ImpressoraECF, cujo código foi passado como parâmetro

Dim lErro As Long
Dim lComando As Long
Dim tImpressoraECF As typeImpressoraECF

On Error GoTo Erro_ImpressoraECF_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 103454

    tImpressoraECF.sNumSerie = String(STRING_IMPRESSORAECF_NUMSERIE, 0)

    'Busca o ImpressoraECF no BD através do código e filial passados como parâmetro
    lErro = Comando_Executar(lComando, "SELECT CodModelo, NumSerie FROM ImpressoraECF WHERE Codigo = ?  AND FilialEmpresa = ? ", tImpressoraECF.iCodModelo, tImpressoraECF.sNumSerie, objImpressoraECF.iCodigo, objImpressoraECF.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 103455

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 103456

    'Se não encontrou o ImpressoraECF => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 103447

    'Transfere para objImpressoraECF os dados obtidos do BD
    objImpressoraECF.iCodModelo = tImpressoraECF.iCodModelo
    objImpressoraECF.sNumSerie = tImpressoraECF.sNumSerie

    'Fecha o comando
    Call Comando_Fechar(lComando)

    ImpressoraECF_Le = SUCESSO

    Exit Function

Erro_ImpressoraECF_Le:

    ImpressoraECF_Le = gErr

    Select Case gErr

        Case 103454
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 103455, 103456
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_IMPRESSORAECF", gErr)

        Case 103447
            'ImpressoraECF não encontrado. Esse erro será tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151504)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ControleLogCCBack_Le_Ultimos(ByVal lUltimos As Long, colControleLogCCBack As Collection, ByVal iFilialEmpresa As Integer) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(1 To 2) As Long
Dim lSeq As Long
Dim lNumIntIni As Long
Dim lNumIntFim As Long
Dim dtData As Date
Dim objControleLogCCBack As ClassControleLogCCBack
Dim lRegistrosLidos As Long

On Error GoTo Erro_ControleLogCCBack_Le_Ultimos
    
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 118919
    Next
    
    'selecionar o ultimo registro que contem info sobre os dados do log que foram enviados para o backfoffice
    lErro = Comando_Executar(alComando(1), "SELECT  Sequencial, NumIntDocInicial, NumIntDocFinal, Data FROM ControleLogCCBack WHERE FilialEmpresa = ? ORDER BY Sequencial DESC ", lSeq, lNumIntIni, lNumIntFim, dtData, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 118920

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 118921
    
    Do While lErro = AD_SQL_SUCESSO And lRegistrosLidos < lUltimos
                
        lRegistrosLidos = lRegistrosLidos + 1
        
        Set objControleLogCCBack = New ClassControleLogCCBack
        
        objControleLogCCBack.dtData = dtData
        objControleLogCCBack.lNumIntDocInicial = lNumIntIni
        objControleLogCCBack.lNumIntDocFinal = lNumIntFim
        objControleLogCCBack.lSequencial = lSeq
        
        colControleLogCCBack.Add objControleLogCCBack
        
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 118922
        
    Loop
    
    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    ControleLogCCBack_Le_Ultimos = SUCESSO

    Exit Function

Erro_ControleLogCCBack_Le_Ultimos:

    ControleLogCCBack_Le_Ultimos = gErr

    Select Case gErr

        Case 118919
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 118920 To 118922
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTROLELOGCCBACK", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151505)

    End Select

    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Function ControleLogCCBack_Le_Ultimo(lSeq As Long) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(1 To 2) As Long

On Error GoTo Erro_ControleLogCCBack_Le_Ultimo
    
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 118928
    Next
    
    'selecionar o ultimo registro que contem info sobre os dados do log que foram enviados para o backfoffice
    lErro = Comando_Executar(alComando(1), "SELECT  Sequencial FROM ControleLogCCBack WHERE FilialEmpresa = ? ORDER BY Sequencial DESC", lSeq, giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 118929

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 118930
    
    If lErro = AD_SQL_SEM_DADOS Then lSeq = 0
    
    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    ControleLogCCBack_Le_Ultimo = SUCESSO

    Exit Function

Erro_ControleLogCCBack_Le_Ultimo:

    ControleLogCCBack_Le_Ultimo = gErr

    Select Case gErr

        Case 118928
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 118929, 118930
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTROLELOGCCBACK", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151506)

    End Select

    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Function TransferenciaLoja_Le(ByVal objTransfCaixa As ClassTransfCaixa) As Long
'Lê os dados da Transferencia cujo codigo/filialempresa foram passados como parametro

Dim lErro As Long
Dim lComando As Long
Dim tTransfCaixa As typeTransfCaixa

On Error GoTo Erro_TransferenciaLoja_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 105232

    'Busca a trasnferenciaLoja no BD através do código e filial passados como parâmetro
    lErro = Comando_Executar(lComando, "SELECT DataBackoffice, NumMovtoCxDe, NumMovtoCxPara FROM TransferenciaLoja WHERE FilialEmpresa = ? AND Codigo = ?", _
    tTransfCaixa.dtDataBackoffice, tTransfCaixa.lNumMovtoCxDe, tTransfCaixa.lNumMovtoCxPara, objTransfCaixa.iFilialEmpresa, objTransfCaixa.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 105233

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 105234

    'Se não encontrou a transferencialoja => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 105235

    'Transfere para objImpressoraECF os dados obtidos do BD
    objTransfCaixa.dtDataBackoffice = tTransfCaixa.dtDataBackoffice
    objTransfCaixa.lNumMovtoCxDe = tTransfCaixa.lNumMovtoCxDe
    objTransfCaixa.lNumMovtoCxPara = tTransfCaixa.lNumMovtoCxPara

    'Fecha o comando
    Call Comando_Fechar(lComando)

    TransferenciaLoja_Le = SUCESSO

    Exit Function

Erro_TransferenciaLoja_Le:

    TransferenciaLoja_Le = gErr

    Select Case gErr

        Case 105232
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 105233, 105234
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRANSFERENCIALOJA", gErr, objTransfCaixa.iFilialEmpresa, objTransfCaixa.lCodigo)

        Case 105235

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151507)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ChequePre_Le_Carne(objChequePre As ClassChequePre) As Long
'le o cheque se estiver vinculado a carne

Dim lErro As Long
Dim lComando As Long
Dim tChequePre As typeChequePre
Dim sComandoSQL As String

On Error GoTo Erro_ChequePre_Le_Carne

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 105252

    'Inicializa o Buffer das variáveis
    tChequePre.sAgencia = String(STRING_BUFFER_MAX_TEXTO, 0)
    tChequePre.sContaCorrente = String(STRING_BUFFER_MAX_TEXTO, 0)
    tChequePre.sCPFCGC = String(STRING_BUFFER_MAX_TEXTO, 0)
    tChequePre.sCarne = String(STRING_BUFFER_MAX_TEXTO, 0)

    sComandoSQL = "SELECT NumIntCheque, ChequePre.Cliente, Filial, Banco, Agencia, ContaCorrente, ChequePre.Numero,  DataDeposito,  ChequePre.Valor, CPFCGC, NaoEspecificado, SequencialBack, FilialEmpresaLoja, ChequePre.Status, ChequePre.Caixa, SequencialCaixa, Carne.CupomFiscal, ChequePre.DataEmissao" & _
        " Carne.CodBarrasCarne" & _
        " FROM ChequePre, MovimentosCaixa, BaixasParcCarne, CarneParcelas , Carne " & _
        " WHERE CHequePre.FilialEmpresa = MovimentosCaixa.FilialEmpresa AND " & _
        " ChequePre.NumMovtoCaixa = MovimentosCaixa.NumMovto AND " & _
        " MovimentosCaixa.FilialEmpresa = BaixasParcCarne.FilialEmpresa AND " & _
        " MovimentosCaixa.NumRefInterna = BaixasParcCarne.NumIntBaixa AND " & _
        " BaixasParcCarne.FilialEmpresa = CarneParcelas.FilialEmpresa AND " & _
        " BaixasParcCarne.NumIntParcelaCarne = CarneParcelas.NumIntDOC AND " & _
        " CarneParcelas.FilialEmpresa = Carne.FilialEmpresa AND " & _
        " CarneParcelas.NumIntCarne = Carne.NumIntDoc AND " & _
        " ChequePre.Status <> 5 AND "

    If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL Then

        objChequePre.lSequencial = objChequePre.lSequencialLoja

        'Pesquisa o cheque pelo codigo de barras associado
        lErro = Comando_Executar(lComando, sComandoSQL & _
            " ChequePre.FilialEmpresaLoja = ? AND ChequePre.SequencialLoja = ? ", _
            tChequePre.lNumIntCheque, tChequePre.lCliente, tChequePre.iFilial, tChequePre.iBanco, tChequePre.sAgencia, tChequePre.sContaCorrente, _
            tChequePre.lNumero, tChequePre.dtDataDeposito, tChequePre.dValor, tChequePre.sCPFCGC, tChequePre.iNaoEspecificado, _
            tChequePre.lSequencialBack, tChequePre.iFilialEmpresaLoja, tChequePre.iStatus, tChequePre.iCaixa, tChequePre.lSequencialCaixa, tChequePre.lCupomFiscal, tChequePre.dtDataEmissao, tChequePre.sCarne, _
            objChequePre.iFilialEmpresaLoja, objChequePre.lSequencialLoja)
        If lErro <> AD_SQL_SUCESSO Then gError 105253

    ElseIf giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL_BACKOFFICE Then
    
        objChequePre.lSequencial = objChequePre.lSequencialBack
    
        'Pesquisa o cheque pelo codigo de barras associado
        lErro = Comando_Executar(lComando, sComandoSQL & _
            " ChequePre.SequencialBack = ? ", _
            tChequePre.lNumIntCheque, tChequePre.lCliente, tChequePre.iFilial, tChequePre.iBanco, tChequePre.sAgencia, tChequePre.sContaCorrente, _
            tChequePre.lNumero, tChequePre.dtDataDeposito, tChequePre.dValor, tChequePre.sCPFCGC, tChequePre.iNaoEspecificado, _
            tChequePre.lSequencialBack, tChequePre.iFilialEmpresaLoja, tChequePre.iStatus, tChequePre.iCaixa, tChequePre.lSequencialCaixa, tChequePre.lCupomFiscal, tChequePre.dtDataEmissao, tChequePre.sCarne, _
            objChequePre.lSequencialBack)
        If lErro <> AD_SQL_SUCESSO Then gError 105254
    
    End If

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 105255

    'Se não encontrou o cheque => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 105256

    objChequePre.lNumIntCheque = tChequePre.lNumIntCheque
    objChequePre.lCliente = tChequePre.lCliente
    objChequePre.iFilial = tChequePre.iFilial
    objChequePre.iBanco = tChequePre.iBanco
    objChequePre.sAgencia = tChequePre.sAgencia
    objChequePre.sContaCorrente = tChequePre.sContaCorrente
    objChequePre.lNumero = tChequePre.lNumero
    objChequePre.dtDataDeposito = tChequePre.dtDataDeposito
    objChequePre.dValor = tChequePre.dValor
    objChequePre.sCPFCGC = tChequePre.sCPFCGC
    objChequePre.iNaoEspecificado = tChequePre.iNaoEspecificado
    objChequePre.lSequencialBack = tChequePre.lSequencialBack
    objChequePre.iFilialEmpresaLoja = tChequePre.iFilialEmpresaLoja
    objChequePre.iStatus = tChequePre.iStatus
    objChequePre.iCaixa = tChequePre.iCaixa
    objChequePre.lSequencialCaixa = tChequePre.lSequencialCaixa
    objChequePre.lCupomFiscal = tChequePre.lCupomFiscal
    objChequePre.sCarne = tChequePre.sCarne
    objChequePre.dtDataEmissao = tChequePre.dtDataEmissao
    
    'Fecha o comando
    Call Comando_Fechar(lComando)

    ChequePre_Le_Carne = SUCESSO

    Exit Function

Erro_ChequePre_Le_Carne:

    ChequePre_Le_Carne = gErr

    Select Case gErr

        Case 105252
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 105253, 105254, 105255
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CHEQUEPRE_LOJA", gErr, objChequePre.lSequencial, objChequePre.iFilialEmpresa)

        Case 105256

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151508)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function CupomFiscal_Le(ByVal objCupomFiscal As ClassCupomFiscal) As Long
'Função que lê o Cupom Fiscal

Dim lErro As Long
Dim tCupomFiscal As typeCupomFiscal
Dim lComando As Long
Dim lCupomFiscal As Long
Dim iECF As Integer

On Error GoTo Erro_CupomFiscal_Le

    'Abertura dos Comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 105259

    'Inicializa o Buffer das variáveis
    tCupomFiscal.sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
    tCupomFiscal.sCPFCGC = String(STRING_CPFCGC1, 0)
    tCupomFiscal.sMotivoCancel = String(STRING_MOTIVOCANCEL, 0)
    
    'Leitura da tabela de cupom fiscal para fim de LockShared
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Status, ECF , DataEmissao , HoraEmissao , CPFCGC , NumOrcamento , ValorTotal, ValorProdutos , ValorDesconto, ValorAcrescimo , TabelaPreco ,NaturezaOp , GerenteCancel , MotivoCancel , Vendedor, Caixa  FROM CupomFiscal  WHERE CupomFiscal.Numero = ?  AND CupomFiscal.FilialEmpresa = ? AND ECF = ?", _
    tCupomFiscal.lNumIntDoc, tCupomFiscal.iStatus, tCupomFiscal.iECF, tCupomFiscal.dtDataEmissao, tCupomFiscal.dHoraEmissao, tCupomFiscal.sCPFCGC, tCupomFiscal.lNumOrcamento, tCupomFiscal.dValorTotal, tCupomFiscal.dValorProdutos, tCupomFiscal.dValorAcrescimo, tCupomFiscal.dValorAcrescimo, tCupomFiscal.iTabelaPreco, tCupomFiscal.sNaturezaOp, tCupomFiscal.lGerenteCancel, tCupomFiscal.sMotivoCancel, tCupomFiscal.iVendedor, tCupomFiscal.iCaixa, objCupomFiscal.lNumero, objCupomFiscal.iFilialEmpresa, objCupomFiscal.iECF)
    If lErro <> AD_SQL_SUCESSO Then gError 105260

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 105261

    'SE Não encontrar Registro
    If lErro = AD_SQL_SEM_DADOS Then gError 105262

    'Traz os dados para o objCupomFiscal
    objCupomFiscal.dtDataEmissao = tCupomFiscal.dtDataEmissao
    objCupomFiscal.dValorAcrescimo = tCupomFiscal.dValorAcrescimo
    objCupomFiscal.dValorDesconto = tCupomFiscal.dValorDesconto
    objCupomFiscal.dValorProdutos = tCupomFiscal.dValorProdutos
    objCupomFiscal.dValorTotal = tCupomFiscal.dValorTotal
    objCupomFiscal.iStatus = tCupomFiscal.iStatus
    objCupomFiscal.iTabelaPreco = tCupomFiscal.iTabelaPreco
    objCupomFiscal.lNumIntDoc = tCupomFiscal.lNumIntDoc
    objCupomFiscal.lNumOrcamento = tCupomFiscal.lNumOrcamento
    objCupomFiscal.sCPFCGC = tCupomFiscal.sCPFCGC
    objCupomFiscal.sMotivoCancel = tCupomFiscal.sMotivoCancel
    objCupomFiscal.sNaturezaOp = tCupomFiscal.sNaturezaOp
    objCupomFiscal.iVendedor = tCupomFiscal.iVendedor
    objCupomFiscal.lGerenteCancel = tCupomFiscal.lGerenteCancel
    objCupomFiscal.iCodCaixa = tCupomFiscal.iCaixa
    
     'Fecha o comando
    Call Comando_Fechar(lComando)

    CupomFiscal_Le = SUCESSO

    Exit Function

Erro_CupomFiscal_Le:

    CupomFiscal_Le = gErr

    Select Case gErr

        Case 105259
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 105260, 105261
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CUPOMFISCAL1", gErr, objCupomFiscal.lNumero)

        Case 105262

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151509)

    End Select

     'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function CupomFiscal_Le_NumIntDoc(ByVal objCupomFiscal As ClassCupomFiscal) As Long
'Função que lê o Cupom Fiscal

Dim lErro As Long
Dim tCupomFiscal As typeCupomFiscal
Dim lComando As Long
Dim lCupomFiscal As Long
Dim iECF As Integer

On Error GoTo Erro_CupomFiscal_Le_NumIntDoc

    'Abertura dos Comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 105265

    'Inicializa o Buffer das variáveis
    tCupomFiscal.sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
    tCupomFiscal.sCPFCGC = String(STRING_CPFCGC1, 0)
    tCupomFiscal.sMotivoCancel = String(STRING_MOTIVOCANCEL, 0)
    
    'Leitura da tabela de cupom fiscal para fim de LockShared
    lErro = Comando_Executar(lComando, "SELECT Numero, Status, ECF , DataEmissao , HoraEmissao , CPFCGC , NumOrcamento , ValorTotal, ValorProdutos , ValorDesconto, ValorAcrescimo , TabelaPreco ,NaturezaOp , GerenteCancel , MotivoCancel , Vendedor, Caixa  FROM CupomFiscal  WHERE FilialEmpresa = ? AND NumIntDoc = ?", _
    tCupomFiscal.lNumero, tCupomFiscal.iStatus, tCupomFiscal.iECF, tCupomFiscal.dtDataEmissao, tCupomFiscal.dHoraEmissao, tCupomFiscal.sCPFCGC, tCupomFiscal.lNumOrcamento, tCupomFiscal.dValorTotal, tCupomFiscal.dValorProdutos, tCupomFiscal.dValorDesconto, tCupomFiscal.dValorAcrescimo, tCupomFiscal.iTabelaPreco, tCupomFiscal.sNaturezaOp, tCupomFiscal.lGerenteCancel, tCupomFiscal.sMotivoCancel, tCupomFiscal.iVendedor, tCupomFiscal.iCaixa, objCupomFiscal.iFilialEmpresa, objCupomFiscal.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 105266

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 105267

    'SE Não encontrar Registro
    If lErro = AD_SQL_SEM_DADOS Then gError 105268

    'Traz os dados para o objCupomFiscal
    objCupomFiscal.dtDataEmissao = tCupomFiscal.dtDataEmissao
    objCupomFiscal.dValorAcrescimo = tCupomFiscal.dValorAcrescimo
    objCupomFiscal.dValorDesconto = tCupomFiscal.dValorDesconto
    objCupomFiscal.dValorProdutos = tCupomFiscal.dValorProdutos
    objCupomFiscal.dValorTotal = tCupomFiscal.dValorTotal
    objCupomFiscal.iStatus = tCupomFiscal.iStatus
    objCupomFiscal.iTabelaPreco = tCupomFiscal.iTabelaPreco
    objCupomFiscal.lNumero = tCupomFiscal.lNumero
    objCupomFiscal.lNumOrcamento = tCupomFiscal.lNumOrcamento
    objCupomFiscal.sCPFCGC = tCupomFiscal.sCPFCGC
    objCupomFiscal.sMotivoCancel = tCupomFiscal.sMotivoCancel
    objCupomFiscal.sNaturezaOp = tCupomFiscal.sNaturezaOp
    objCupomFiscal.iVendedor = tCupomFiscal.iVendedor
    objCupomFiscal.lGerenteCancel = tCupomFiscal.lGerenteCancel
    objCupomFiscal.iCodCaixa = tCupomFiscal.iCaixa
    objCupomFiscal.iECF = tCupomFiscal.iECF
    
     'Fecha o comando
    Call Comando_Fechar(lComando)

    CupomFiscal_Le_NumIntDoc = SUCESSO

    Exit Function

Erro_CupomFiscal_Le_NumIntDoc:

    CupomFiscal_Le_NumIntDoc = gErr

    Select Case gErr

        Case 105265
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 105266, 105267
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CUPOMFISCAL1", gErr, objCupomFiscal.lNumIntDoc)

        Case 105268

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151510)

    End Select

     'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function BaixasCarne_Le_CodBarras(ByVal objCarne As ClassCarne, ByVal objBaixaCarne As ClassBaixasCarne) As Long
'le uma das baixas associada ao carne em questao

Dim lErro As Long
Dim lComando As Long
Dim lNumIntDoc As Long

On Error GoTo Erro_BaixasCarne_Le_CodBarras

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 105290

    'Seleciona código e nome das Carnes no Banco de Dados da Mesma FilialEmpresa
    lErro = Comando_Executar(lComando, "SELECT NumIntBaixa FROM BaixasCarneLoja WHERE CodBarrasCarne = ? AND FilialEmpresa =?", lNumIntDoc, objCarne.sCodBarrasCarne, objCarne.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 105291

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 105292

    'Se Não Encontrar Registros
    If lErro = AD_SQL_SEM_DADOS Then gError 105293

    objBaixaCarne.lNumIntDoc = lNumIntDoc
    objBaixaCarne.iFilialEmpresa = objCarne.iFilialEmpresa
    
    BaixasCarne_Le_CodBarras = SUCESSO
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
Erro_BaixasCarne_Le_CodBarras:

    BaixasCarne_Le_CodBarras = gErr
    
    Select Case gErr
    
        Case 105290
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 105291, 105292
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CARNE", gErr)

        Case 105293
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151511)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function BaixasCarne_Lock(ByVal objBaixaCarne As ClassBaixasCarne) As Long
'loca a baixa de carne em questao
'IMPORTANTE: TEM QUE SER CHAMADO EM TRANSACAO

Dim lErro As Long
Dim lComando As Long
Dim lNumIntDoc As Long

On Error GoTo Erro_BaixasCarne_Lock

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 105294

    'Seleciona código e nome das Carnes no Banco de Dados da Mesma FilialEmpresa
    lErro = Comando_ExecutarPos(lComando, "SELECT NumIntBaixa FROM BaixasCarne WHERE NumIntBaixa = ? AND FilialEmpresa =?", 0, lNumIntDoc, objBaixaCarne.lNumIntDoc, objBaixaCarne.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 105295

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 105296

    'Se Não Encontrar Registros
    If lErro = AD_SQL_SEM_DADOS Then gError 105297

    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then gError 105298

    BaixasCarne_Lock = SUCESSO
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
Erro_BaixasCarne_Lock:

    BaixasCarne_Lock = gErr
    
    Select Case gErr
    
        Case 105294
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 105295, 105296
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BAIXASCARNE1", gErr, objBaixaCarne.iFilialEmpresa, objBaixaCarne.lNumIntDoc)

        Case 105297

        Case 105298
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_BAIXASCARNE", gErr, objBaixaCarne.iFilialEmpresa, objBaixaCarne.lNumIntDoc)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151512)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function Cheque_Le_Excluido(ByVal objCheque As ClassChequePre) As Long
'Função que lê as Informações do Cheque (inclusive os excluidos logicamente) passado como parâmetro no banco de dados, Carregando no objCheque

Dim lErro As Long
Dim lComando As Long
Dim lCupomFiscal As Long
Dim tCheque As typeChequePre
Dim iTipoMovcx As Integer

On Error GoTo Erro_Cheque_Le_Excluido

    'Inicia a Abertura de o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 105307

    'Inicializa o Buffer das variáveis
    tCheque.sAgencia = String(STRING_BUFFER_MAX_TEXTO, 0)
    tCheque.sContaCorrente = String(STRING_BUFFER_MAX_TEXTO, 0)
    tCheque.sCPFCGC = String(STRING_BUFFER_MAX_TEXTO, 0)
    
    'Verifica se Esta Trabalhando no caixa central
    If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL Then

        'Setar a Variável para consulta SQL, quando a Consulta vem mediante ao Browser
        objCheque.lSequencial = objCheque.lSequencialLoja

        'Procura o Cheque no BD a partir do código e a FilialEmpresaLoja
        lErro = Comando_Executar(lComando, "SELECT NumIntCheque, Cliente, Filial , Banco , ChequePre.Agencia , ContaCorrente , ChequePre.Numero, DataDeposito , ChequePre.Valor , NumBordero , ChequePre.CPFCGC , NumMovtoCaixa , Aprovado , NaoEspecificado , NumBorderoLoja , ChequePre.FilialEmpresa , SequencialLoja , SequencialBack,  ChequePre.Status ,MovimentosCaixa.CupomFiscal, MovimentosCaixa.Tipo, ChequePre.Localizacao, ChequePre.NumBorderoLojaBanco, ChequePre.Caixa, ChequePre.FilialEmpresaLoja, ChequePre.NumMovtoSangria, ChequePre.SequencialCaixa, ChequePre.FilialEmpresa, MovimentosCaixa.NumRefInterna, ChequePre.DataEmissao  FROM ChequePre , MovimentosCaixa WHERE ChequePre.SequencialLoja = ? AND ChequePre.NumMovtoCaixa = MovimentosCaixa.NumMovto AND MovimentosCaixa.FilialEmpresa = ChequePre.FilialEmpresaLoja AND ChequePre.FilialEmpresaLoja = ? ", _
        tCheque.lNumIntCheque, tCheque.lCliente, tCheque.iFilial, tCheque.iBanco, tCheque.sAgencia, tCheque.sContaCorrente, tCheque.lNumero, tCheque.dtDataDeposito, tCheque.dValor, tCheque.lNumBordero, tCheque.sCPFCGC, tCheque.lNumMovtoCaixa, tCheque.iAprovado, tCheque.iNaoEspecificado, tCheque.lNumBorderoLoja, tCheque.iFilialEmpresa, tCheque.lSequencialLoja, tCheque.lSequencialBack, tCheque.iStatus, lCupomFiscal, iTipoMovcx, tCheque.iLocalizacao, tCheque.lNumBorderoLojaBanco, tCheque.iCaixa, tCheque.iFilialEmpresaLoja, tCheque.lNumMovtoSangria, tCheque.lSequencialCaixa, tCheque.iFilialEmpresa, tCheque.lNumIntExt, tCheque.dtDataEmissao, objCheque.lSequencial, objCheque.iFilialEmpresaLoja)
        If lErro <> AD_SQL_SUCESSO Then gError 105308

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 105309

        'Se não encontrou o Cheque => erro
        If lErro = AD_SQL_SEM_DADOS Then gError 105310

        
    Else
        
        objCheque.lSequencial = objCheque.lSequencialBack

        'Procura o Cheque no BD a partir do código e a FilialEmpresaLoja
        lErro = Comando_Executar(lComando, "SELECT NumIntCheque, Cliente, Filial , Banco , Agencia , ContaCorrente , ChequePre.Numero, DataDeposito , ChequePre.Valor , NumBordero , ChequePre.CPFCGC , NumMovtoCaixa , Aprovado , NaoEspecificado , NumBorderoLoja  ,  SequencialLoja , SequencialBack, ChequePre.Status ,MovimentosCaixa.CupomFiscal, ChequePre.Localizacao, ChequePre.NumBorderoLojaBanco, ChequePre.Caixa, ChequePre.FilialEmpresaLoja, ChequePre.NumMovtoSangria, ChequePre.SequencialCaixa, ChequePre.FilialEmpresa, MovimentosCaixa.NumRefInterna, ChequePre.DataEmissao FROM ChequePre , MovimentosCaixa WHERE ChequePre.SequencialBack = ? AND ChequePre.NumMovtoCaixa = MovimentosCaixa.NumMovto AND MovimentosCaixa.FilialEmpresa = ChequePre.FilialEmpresa ", _
        tCheque.lNumIntCheque, tCheque.lCliente, tCheque.iFilial, tCheque.iBanco, tCheque.sAgencia, tCheque.sContaCorrente, tCheque.lNumero, tCheque.dtDataDeposito, tCheque.dValor, tCheque.lNumBordero, tCheque.sCPFCGC, tCheque.lNumMovtoCaixa, tCheque.iAprovado, tCheque.iNaoEspecificado, tCheque.lNumBorderoLoja, tCheque.lSequencialLoja, tCheque.lSequencialBack, tCheque.iStatus, lCupomFiscal, tCheque.iLocalizacao, tCheque.lNumBorderoLojaBanco, tCheque.iCaixa, tCheque.iFilialEmpresaLoja, tCheque.lNumMovtoSangria, tCheque.lSequencialCaixa, tCheque.iFilialEmpresa, tCheque.lNumIntExt, tCheque.dtDataEmissao, objCheque.lSequencialBack)
        If lErro <> AD_SQL_SUCESSO Then gError 105311

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 105312

        'Se não encontrou o Cheque => erro
        If lErro = AD_SQL_SEM_DADOS Then gError 105313
        
    End If

    'Passa para objCheque os valores lidos do BD
    objCheque.dtDataDeposito = tCheque.dtDataDeposito
    objCheque.dValor = tCheque.dValor
    objCheque.iAprovado = tCheque.iAprovado
    objCheque.iBanco = tCheque.iBanco
    objCheque.iFilial = tCheque.iFilial
    objCheque.iFilialEmpresa = tCheque.iFilialEmpresa
    objCheque.iNaoEspecificado = tCheque.iNaoEspecificado
    objCheque.lCliente = tCheque.lCliente
    objCheque.lNumBordero = tCheque.lNumBordero
    objCheque.lNumBorderoLoja = tCheque.lNumBorderoLoja
    objCheque.lNumero = tCheque.lNumero
    objCheque.lNumIntCheque = tCheque.lNumIntCheque
    objCheque.lNumMovtoCaixa = tCheque.lNumMovtoCaixa
    objCheque.lSequencialBack = tCheque.lSequencialBack
    objCheque.lSequencialLoja = tCheque.lSequencialLoja
    objCheque.sAgencia = tCheque.sAgencia
    objCheque.sContaCorrente = tCheque.sContaCorrente
    objCheque.sCPFCGC = tCheque.sCPFCGC
    objCheque.lCupomFiscal = lCupomFiscal
    objCheque.iStatus = tCheque.iStatus
    objCheque.iLocalizacao = tCheque.iLocalizacao
    objCheque.lNumBorderoLojaBanco = tCheque.lNumBorderoLojaBanco
    objCheque.iCaixa = tCheque.iCaixa
    objCheque.iFilialEmpresaLoja = tCheque.iFilialEmpresaLoja
    objCheque.lNumMovtoSangria = tCheque.lNumMovtoSangria
    objCheque.lSequencialCaixa = tCheque.lSequencialCaixa
    objCheque.lNumIntExt = tCheque.lNumIntExt
    objCheque.dtDataEmissao = tCheque.dtDataEmissao
    
    'Fecha o comando
    Call Comando_Fechar(lComando)

    Cheque_Le_Excluido = SUCESSO

    Exit Function

Erro_Cheque_Le_Excluido:

    Cheque_Le_Excluido = gErr

    Select Case gErr

        Case 105307
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 105308, 105309, 105311, 105312
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CHEQUE1", gErr, objCheque.lSequencial)

        Case 105310, 105313

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151513)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Cheque_Le_SequencialCaixa1(ByVal objCheque As ClassChequePre) As Long
'Função que lê as Informações do Cheque passado como parâmetro no banco de dados, Carregando no objCheque

Dim lErro As Long
Dim lComando As Long
Dim lCupomFiscal As Long
Dim tCheque As typeChequePre
Dim iTipoMovcx As Integer

On Error GoTo Erro_Cheque_Le_SequencialCaixa1

    'Inicia a Abertura de o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 105478

    'Inicializa o Buffer das variáveis
    tCheque.sAgencia = String(STRING_BUFFER_MAX_TEXTO, 0)
    tCheque.sContaCorrente = String(STRING_BUFFER_MAX_TEXTO, 0)
    tCheque.sCPFCGC = String(STRING_BUFFER_MAX_TEXTO, 0)
    
    'Procura o Cheque no BD a partir do código e a FilialEmpresaLoja
    lErro = Comando_Executar(lComando, "SELECT NumIntCheque, Cliente, Filial , Banco , ChequePre.Agencia , ContaCorrente , ChequePre.Numero, DataDeposito , ChequePre.Valor , NumBordero , ChequePre.CPFCGC , NumMovtoCaixa , Aprovado , NaoEspecificado , NumBorderoLoja , ChequePre.FilialEmpresa , SequencialLoja , SequencialBack,  ChequePre.Status ,ECF , MovimentosCaixa.CupomFiscal, MovimentosCaixa.Tipo, ChequePre.Localizacao, ChequePre.NumBorderoLojaBanco, ChequePre.Caixa, ChequePre.FilialEmpresaLoja, ChequePre.NumMovtoSangria, ChequePre.SequencialCaixa, ChequePre.FilialEmpresa, ChequePre.DataEmissao FROM CupomFiscal , ChequePre , MovimentosCaixa WHERE ChequePre.FilialEmpresaLoja = ? AND ChequePre.Caixa = ? AND ChequePre.SequencialCaixa = ? AND ChequePre.NumMovtoCaixa = MovimentosCaixa.NumMovto AND MovimentosCaixa.CupomFiscal = CupomFiscal.NumIntDoc AND MovimentosCaixa.FilialEmpresa = ChequePre.FilialEmpresaLoja AND ChequePre.Status <> 5", _
    tCheque.lNumIntCheque, tCheque.lCliente, tCheque.iFilial, tCheque.iBanco, tCheque.sAgencia, tCheque.sContaCorrente, tCheque.lNumero, tCheque.dtDataDeposito, tCheque.dValor, tCheque.lNumBordero, tCheque.sCPFCGC, tCheque.lNumMovtoCaixa, tCheque.iAprovado, tCheque.iNaoEspecificado, tCheque.lNumBorderoLoja, tCheque.iFilialEmpresa, tCheque.lSequencialLoja, tCheque.lSequencialBack, tCheque.iStatus, tCheque.iECF, lCupomFiscal, iTipoMovcx, tCheque.iLocalizacao, tCheque.lNumBorderoLojaBanco, tCheque.iCaixa, tCheque.iFilialEmpresaLoja, tCheque.lNumMovtoSangria, tCheque.lSequencialCaixa, tCheque.iFilialEmpresa, tCheque.dtDataEmissao, objCheque.iFilialEmpresaLoja, objCheque.iCaixa, objCheque.lSequencialCaixa)
    If lErro <> AD_SQL_SUCESSO Then gError 105479

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 105480

    'Se não encontrou o Cheque => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 105481
        
    'Passa para objCheque os valores lidos do BD
    objCheque.dtDataDeposito = tCheque.dtDataDeposito
    objCheque.dValor = tCheque.dValor
    objCheque.iAprovado = tCheque.iAprovado
    objCheque.iBanco = tCheque.iBanco
    objCheque.iFilial = tCheque.iFilial
    objCheque.iFilialEmpresa = tCheque.iFilialEmpresa
    objCheque.iNaoEspecificado = tCheque.iNaoEspecificado
    objCheque.lCliente = tCheque.lCliente
    objCheque.lNumBordero = tCheque.lNumBordero
    objCheque.lNumBorderoLoja = tCheque.lNumBorderoLoja
    objCheque.lNumero = tCheque.lNumero
    objCheque.lNumIntCheque = tCheque.lNumIntCheque
    objCheque.lNumMovtoCaixa = tCheque.lNumMovtoCaixa
    objCheque.lSequencialBack = tCheque.lSequencialBack
    objCheque.lSequencialLoja = tCheque.lSequencialLoja
    objCheque.sAgencia = tCheque.sAgencia
    objCheque.sContaCorrente = tCheque.sContaCorrente
    objCheque.sCPFCGC = tCheque.sCPFCGC
    objCheque.lCupomFiscal = lCupomFiscal
    objCheque.iECF = tCheque.iECF
    objCheque.iStatus = tCheque.iStatus
    objCheque.iLocalizacao = tCheque.iLocalizacao
    objCheque.lNumBorderoLojaBanco = tCheque.lNumBorderoLojaBanco
    objCheque.iCaixa = tCheque.iCaixa
    objCheque.iFilialEmpresaLoja = tCheque.iFilialEmpresaLoja
    objCheque.lNumMovtoSangria = tCheque.lNumMovtoSangria
    objCheque.lSequencialCaixa = tCheque.lSequencialCaixa
    objCheque.dtDataEmissao = tCheque.dtDataEmissao
    
    'Fecha o comando
    Call Comando_Fechar(lComando)

    Cheque_Le_SequencialCaixa1 = SUCESSO

    Exit Function

Erro_Cheque_Le_SequencialCaixa1:

    Cheque_Le_SequencialCaixa1 = gErr

    Select Case gErr

        Case 105478
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 105479, 105480
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CHEQUE2", gErr, objCheque.iFilialEmpresaLoja, objCheque.iCaixa, objCheque.lSequencialCaixa)

        Case 105481

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151514)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Cheque_Lock1(ByVal objCheque As ClassChequePre) As Long
'Função Realiza lock na Tabela de ChequePre com o Sequencial e Filial Passado como parâmetro

Dim lErro As Long
Dim lComando As Long
Dim tChequePre As typeChequePre

On Error GoTo Erro_Cheque_Lock1

     'Abertura dos Comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 105581
    
    lErro = Comando_ExecutarPos(lComando, "SELECT SequencialBack, NumMovtoCaixa , Valor,  Localizacao, NaoEspecificado, Status, DataEmissao FROM ChequePre WHERE SequencialLoja = ? AND FilialEmpresaLoja = ?", 0, tChequePre.lSequencialBack, tChequePre.lNumMovtoCaixa, tChequePre.dValor, tChequePre.iLocalizacao, tChequePre.iNaoEspecificado, tChequePre.iStatus, tChequePre.dtDataEmissao, objCheque.lSequencialLoja, objCheque.iFilialEmpresaLoja)
    If lErro <> AD_SQL_SUCESSO Then gError 105582
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 105583

    If lErro = AD_SQL_SEM_DADOS Then gError 105584

    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then gError 105585
    
    objCheque.lNumIntCheque = tChequePre.lNumIntCheque
    objCheque.lNumMovtoCaixa = tChequePre.lNumMovtoCaixa
    objCheque.dValor = tChequePre.dValor
    objCheque.iLocalizacao = tChequePre.iLocalizacao
    objCheque.iNaoEspecificado = tChequePre.iNaoEspecificado
    objCheque.iStatus = tChequePre.iStatus
    objCheque.lSequencialBack = tChequePre.lSequencialBack
    objCheque.dtDataEmissao = tChequePre.dtDataEmissao
    
    Cheque_Lock1 = SUCESSO

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

Erro_Cheque_Lock1:

    Cheque_Lock1 = gErr

    Select Case gErr

        Case 105581
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 105582, 105583
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CHEQUE1", gErr, objCheque.lSequencialLoja)

        Case 105584

        Case 105585
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_CHEQUE", gErr, objCheque.lSequencialLoja)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151515)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function CCMov_Le_SldAnterior(dSaldo As Double, ByVal iCaixa As Integer, ByVal dtData As Date, ByVal iFilialEmpresa As Integer, Optional ByVal iTipoMeioPagto As Integer = 0, Optional ByVal sTabela As String = "CCMov") As Long
'Le o Saldo do caixa central na data passada como parametro, somente o que tem acumulado em CCMov.
'Se a data for o ultimo dia do mes em questao o saldo inclui o mes em questao
'caso contrario só inclui até o mes anterior

Dim lErro As Long
Dim dCredito(NUM_MAX_PERIODOS) As Double
Dim dDebito(NUM_MAX_PERIODOS) As Double
Dim dMovimento As Double
Dim dSldIni As Double
Dim iIndice As Integer
Dim iMes As Integer
Dim iDia As Integer
Dim iAno As Integer
Dim iAnoParam As Integer
Dim iIndice2 As Integer
Dim lComando As Long

On Error GoTo Erro_CCMov_Le_SldAnterior

    lComando = Comando_AbrirExt(GL_lConexaoRel)
    If lComando = 0 Then gError 105640

    If dtData = DATA_NULA Then

        'le os saldos do ccmov
        If iTipoMeioPagto = 0 Then
            If iCaixa <> 0 Then
                lErro = Comando_Executar(lComando, "SELECT SaldoInicial FROM " & sTabela & " WHERE FilialEmpresa = ? AND CodCaixa = ? ORDER BY Ano ", dMovimento, iFilialEmpresa, iCaixa)
            Else
                lErro = Comando_Executar(lComando, "SELECT SaldoInicial FROM " & sTabela & " WHERE FilialEmpresa = ? AND CodCaixa <> 9999 ORDER BY Ano ", dMovimento, iFilialEmpresa)
            End If
        Else
            If iCaixa <> 0 Then
                lErro = Comando_Executar(lComando, "SELECT SaldoInicial FROM " & sTabela & " WHERE FilialEmpresa = ? AND CodCaixa = ? AND TipoMeioPagto = ? ORDER BY Ano ", dMovimento, iFilialEmpresa, iCaixa, iTipoMeioPagto)
            Else
                lErro = Comando_Executar(lComando, "SELECT SaldoInicial FROM " & sTabela & " WHERE FilialEmpresa = ? AND CodCaixa <> 9999 AND TipoMeioPagto = ? ORDER BY Ano ", dMovimento, iFilialEmpresa, iTipoMeioPagto)
            End If
        End If
        If lErro <> AD_SQL_SUCESSO Then gError 105647

        'le o primeiro saldo do cliente
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 105648


    Else
    
        If Month(dtData) = 1 Then
            iAnoParam = Year(dtData) - 1
            iIndice2 = 12
        Else
            iAnoParam = Year(dtData)
            iIndice2 = Month(dtData) - 1
        End If
    
        'le os saldos do ccmov
        If iTipoMeioPagto = 0 Then
            If iCaixa <> 0 Then
                lErro = Comando_Executar(lComando, "SELECT Ano, SaldoInicial, Deb01, Deb02, Deb03, Deb04, Deb05, Deb06, Deb07, Deb08, Deb09, Deb10, Deb11, Deb12, Cred01, Cred02, Cred03, Cred04, Cred05, Cred06, Cred07, Cred08, Cred09, Cred10, Cred11, Cred12 FROM " & sTabela & " WHERE FilialEmpresa = ? AND CodCaixa = ? AND Ano = ? ORDER BY Ano DESC", iAno, dSldIni, dDebito(1), dDebito(2), dDebito(3), dDebito(4), dDebito(5), dDebito(6), dDebito(7), dDebito(8), dDebito(9), dDebito(10), dDebito(11), dDebito(12), dCredito(1), dCredito(2), dCredito(3), dCredito(4), dCredito(5), dCredito(6), dCredito(7), dCredito(8), dCredito(9), dCredito(10), dCredito(11), dCredito(12), iFilialEmpresa, iCaixa, iAnoParam)
            Else
                lErro = Comando_Executar(lComando, "SELECT Ano, SaldoInicial, Deb01, Deb02, Deb03, Deb04, Deb05, Deb06, Deb07, Deb08, Deb09, Deb10, Deb11, Deb12, Cred01, Cred02, Cred03, Cred04, Cred05, Cred06, Cred07, Cred08, Cred09, Cred10, Cred11, Cred12 FROM " & sTabela & " WHERE FilialEmpresa = ? AND CodCaixa <> 9999 AND Ano = ? ORDER BY Ano DESC", iAno, dSldIni, dDebito(1), dDebito(2), dDebito(3), dDebito(4), dDebito(5), dDebito(6), dDebito(7), dDebito(8), dDebito(9), dDebito(10), dDebito(11), dDebito(12), dCredito(1), dCredito(2), dCredito(3), dCredito(4), dCredito(5), dCredito(6), dCredito(7), dCredito(8), dCredito(9), dCredito(10), dCredito(11), dCredito(12), iFilialEmpresa, iAnoParam)
            End If
        Else
            If iCaixa <> 0 Then
                lErro = Comando_Executar(lComando, "SELECT Ano, SaldoInicial, Deb01, Deb02, Deb03, Deb04, Deb05, Deb06, Deb07, Deb08, Deb09, Deb10, Deb11, Deb12, Cred01, Cred02, Cred03, Cred04, Cred05, Cred06, Cred07, Cred08, Cred09, Cred10, Cred11, Cred12 FROM " & sTabela & " WHERE FilialEmpresa = ? AND CodCaixa = ? AND Ano = ? AND TipoMeioPagto = ? ORDER BY Ano DESC", iAno, dSldIni, dDebito(1), dDebito(2), dDebito(3), dDebito(4), dDebito(5), dDebito(6), dDebito(7), dDebito(8), dDebito(9), dDebito(10), dDebito(11), dDebito(12), dCredito(1), dCredito(2), dCredito(3), dCredito(4), dCredito(5), dCredito(6), dCredito(7), dCredito(8), dCredito(9), dCredito(10), dCredito(11), dCredito(12), iFilialEmpresa, iCaixa, iAnoParam, iTipoMeioPagto)
            Else
                lErro = Comando_Executar(lComando, "SELECT Ano, SaldoInicial, Deb01, Deb02, Deb03, Deb04, Deb05, Deb06, Deb07, Deb08, Deb09, Deb10, Deb11, Deb12, Cred01, Cred02, Cred03, Cred04, Cred05, Cred06, Cred07, Cred08, Cred09, Cred10, Cred11, Cred12 FROM " & sTabela & " WHERE FilialEmpresa = ? AND CodCaixa <> 9999 AND Ano = ? AND TipoMeioPagto = ? ORDER BY Ano DESC", iAno, dSldIni, dDebito(1), dDebito(2), dDebito(3), dDebito(4), dDebito(5), dDebito(6), dDebito(7), dDebito(8), dDebito(9), dDebito(10), dDebito(11), dDebito(12), dCredito(1), dCredito(2), dCredito(3), dCredito(4), dCredito(5), dCredito(6), dCredito(7), dCredito(8), dCredito(9), dCredito(10), dCredito(11), dCredito(12), iFilialEmpresa, iAnoParam, iTipoMeioPagto)
            End If
        End If
        If lErro <> AD_SQL_SUCESSO Then gError 105641
        
        'le o primeiro saldo do cliente
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 105642
        
        'Wagner
        'Descomentei porque é por tipo de meio de pagamento, não pode pegar simplesmente o 1o
        'a não ser no caso que passe o tipo (que eu implementei)
        'também não pode pegar o próprio mês se o for o último dia dele pois estaria considerando a movimentação do proprio dia o que é errado
        'pois é uma leitura de saldo inicial do dia
        dMovimento = 0
        Do While lErro = AD_SQL_SUCESSO
            
            dMovimento = Round(dMovimento + dSldIni, 2)
            'dMovimento = Round(dSldIni, 2)
        
'            If iAno = iAnoParam Then
'
'                iMes = Month(dtData)
''                iDia = Day(dtData)
''
''                'Verifica se é o ultimo dia do mes
''                If iDia = Dias_Mes(iMes, iAnoParam) Then
''                    'considera o mês da data informada
''                    iIndice2 = iMes
''                Else
'                    'considera até o mês anterior à data informada
'                    iIndice2 = iMes - 1
''                End If
'
'            Else
'                iIndice2 = 12
'            End If
            
            'acumula os saldos em dMovimento
            For iIndice = 1 To iIndice2
                dMovimento = Round(dMovimento + dCredito(iIndice) - dDebito(iIndice), 2)
            Next
                
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 105643

        Loop
        
    End If
        
    dSaldo = dMovimento
        
    Call Comando_Fechar(lComando)
        
    CCMov_Le_SldAnterior = SUCESSO
    
    Exit Function
 
Erro_CCMov_Le_SldAnterior:

    CCMov_Le_SldAnterior = gErr

    Select Case gErr
    
        Case 105640
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 105641, 105642, 105643, 105647, 105648
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CCMOV", gErr)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151516)
    
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function CCMovDia_Le_SldAnterior(dSaldoDiario As Double, ByVal iCaixa As Integer, ByVal dtData As Date, ByVal iFilialEmpresa As Integer, Optional ByVal iTipoMeioPagto As Integer = 0, Optional ByVal sTabela As String = "CCMovDia") As Long
'Le o Saldo do caixa central no mes passado como parametro

Dim lErro As Long
Dim dCredito As Double
Dim dDebito As Double
Dim dMovimento As Double
Dim dSldIni As Double
Dim iIndice As Integer
Dim iMes As Integer
Dim iAno As Integer
Dim dtDataAnterior As Date
Dim lComando As Long

On Error GoTo Erro_CCMovDia_Le_SldAnterior

    lComando = Comando_AbrirExt(GL_lConexaoRel)
    If lComando = 0 Then gError 105644
    
    iMes = Month(dtData)
    iAno = Year(dtData)
    dtDataAnterior = CDate("01/" & Month(dtData) & "/" & Year(dtData))
    
    'le os saldos do cliente, no intervalo de dias informado
    If iTipoMeioPagto = 0 Then
        If iCaixa <> 0 Then
            lErro = Comando_Executar(lComando, "SELECT SUM(Cred - Deb) FROM " & sTabela & " WHERE FilialEmpresa = ? AND CodCaixa = ? AND Data>=? AND Data<=?", dMovimento, iFilialEmpresa, iCaixa, dtDataAnterior, dtData - 1)
        Else
            lErro = Comando_Executar(lComando, "SELECT SUM(Cred - Deb) FROM " & sTabela & " WHERE FilialEmpresa = ? AND CodCaixa <> 9999 AND Data>=? AND Data<=?", dMovimento, iFilialEmpresa, dtDataAnterior, dtData - 1)
        End If
    Else
        If iCaixa <> 0 Then
            lErro = Comando_Executar(lComando, "SELECT SUM(Cred - Deb) FROM " & sTabela & " WHERE FilialEmpresa = ? AND CodCaixa = ? AND Data>=? AND Data<=? AND TipoMeioPagto = ?", dMovimento, iFilialEmpresa, iCaixa, dtDataAnterior, dtData - 1, iTipoMeioPagto)
        Else
            lErro = Comando_Executar(lComando, "SELECT SUM(Cred - Deb) FROM " & sTabela & " WHERE FilialEmpresa = ? AND CodCaixa <> 9999 AND Data>=? AND Data<=? AND TipoMeioPagto = ?", dMovimento, iFilialEmpresa, dtDataAnterior, dtData - 1, iTipoMeioPagto)
        End If
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 105645
    
    'le o primeiro saldo do cliente
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 105646
    
    dSaldoDiario = dMovimento
        
    Call Comando_Fechar(lComando)
        
    CCMovDia_Le_SldAnterior = SUCESSO
    
    Exit Function
 
Erro_CCMovDia_Le_SldAnterior:

    CCMovDia_Le_SldAnterior = gErr

    Select Case gErr
    
        Case 105644
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 105645, 105646
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CCMOVDIA", gErr)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151517)
    
    End Select
    
    Exit Function

End Function

Function CCMovDia_Le_DebCred(dDeb As Double, dCred As Double, ByVal iCaixa As Integer, ByVal dtData As Date, ByVal iFilialEmpresa As Integer, Optional ByVal iTipoMeioPagto As Integer = 0, Optional ByVal sTabela As String = "CCMovDia") As Long
'Le o Saldo do caixa central no mes passado como parametro

Dim lErro As Long
Dim dCredito As Double
Dim dDebito As Double
Dim lComando As Long

On Error GoTo Erro_CCMovDia_Le_DebCred

    lComando = Comando_AbrirExt(GL_lConexaoRel)
    If lComando = 0 Then gError 105644
    
    If iTipoMeioPagto = 0 Then
        If iCaixa <> 0 Then
            lErro = Comando_Executar(lComando, "SELECT SUM(Cred), SUM(Deb) FROM " & sTabela & " WHERE FilialEmpresa = ? AND CodCaixa = ? AND Data=? ", dCredito, dDebito, iFilialEmpresa, iCaixa, dtData)
        Else
            lErro = Comando_Executar(lComando, "SELECT SUM(Cred), SUM(Deb) FROM " & sTabela & " WHERE FilialEmpresa = ? AND CodCaixa <> 9999 AND Data=? ", dCredito, dDebito, iFilialEmpresa, dtData)
        End If
    Else
        If iCaixa <> 0 Then
            lErro = Comando_Executar(lComando, "SELECT SUM(Cred), SUM(Deb) FROM " & sTabela & " WHERE FilialEmpresa = ? AND CodCaixa = ? AND Data=? AND TipoMeioPagto = ?", dCredito, dDebito, iFilialEmpresa, iCaixa, dtData, iTipoMeioPagto)
        Else
            lErro = Comando_Executar(lComando, "SELECT SUM(Cred), SUM(Deb) FROM " & sTabela & " WHERE FilialEmpresa = ? AND CodCaixa <> 9999 AND Data=? AND TipoMeioPagto = ?", dCredito, dDebito, iFilialEmpresa, dtData, iTipoMeioPagto)
        End If
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 105645
    
    'le o primeiro saldo do cliente
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 105646
    
    dDeb = dDebito
    dCred = dCredito
        
    Call Comando_Fechar(lComando)
        
    CCMovDia_Le_DebCred = SUCESSO
    
    Exit Function
 
Erro_CCMovDia_Le_DebCred:

    CCMovDia_Le_DebCred = gErr

    Select Case gErr
    
        Case 105644
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 105645, 105646
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CCMOVDIA", gErr)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151517)
    
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function MovimentosCaixa_Le_MovAbertCx(ByVal objMovimentoCaixa As ClassMovimentoCaixa) As Long
'Lê em MovimentosCaixa os dados do último movimento abertura para o caixa passado como parâmetro na data e filial passadas.

Dim lErro As Long
Dim lComando As Long
Dim tMovCaixa As typeMovimentoCaixa

On Error GoTo Erro_MovimentosCaixa_Le_MovAbertCx

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 116232

    'Inicializa as strings que receberão dados do BD
    tMovCaixa.sFavorecido = String(STRING_FAVORECIDO, 0)
    tMovCaixa.sHistorico = String(STRING_HISTORICO1, 0)
    
    With tMovCaixa
        
        'faz um select no Bd
        lErro = Comando_Executar(lComando, "SELECT NumMovto, Caixa,  Operador, Sequencial, Tipo, AdmMeioPagto, Parcelamento, TipoCartao, Numero, Valor, Historico, Favorecido, CupomFiscal,NumRefInterna, MovtoTransf, MovtoEstorno, Gerente, Hora FROM MovimentosCaixa WHERE DataMovimento = ? AND FilialEmpresa = ? AND Caixa = ? AND Tipo = ? ORDER BY Hora DESC", _
        .lNumMovto, .iCaixa, .iCodOperador, .lSequencial, .iTipo, .iAdmMeioPagto, .iParcelamento, .iTipoCartao, .lNumero, .dValor, .sHistorico, .sFavorecido, .lCupomFiscal, .lNumRefInterna, .lMovtoTransf, .lMovtoEstorno, .iGerente, .dHora, objMovimentoCaixa.dtDataMovimento, objMovimentoCaixa.iFilialEmpresa, objMovimentoCaixa.iCaixa, MOVIMENTO_CAIXA_ABERTURA)
        If lErro <> AD_SQL_SUCESSO Then gError 116233
    
        'busca o 1º registro
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 116234
    
        'Se não encontrou os dados => erro
        If lErro = AD_SQL_SEM_DADOS Then gError 116235
    
        'guarda os dados lidos no objMovimentoCaixa
        objMovimentoCaixa.dHora = .dHora
        objMovimentoCaixa.dValor = .dValor
        objMovimentoCaixa.iAdmMeioPagto = .iAdmMeioPagto
        objMovimentoCaixa.iCaixa = .iCaixa
        objMovimentoCaixa.iCodOperador = .iCodOperador
        objMovimentoCaixa.iCodConta = .iCodConta
        objMovimentoCaixa.iGerente = .iGerente
        objMovimentoCaixa.iParcelamento = .iParcelamento
        objMovimentoCaixa.iTipo = .iTipo
        objMovimentoCaixa.iTipoCartao = .iTipoCartao
        objMovimentoCaixa.lCupomFiscal = .lCupomFiscal
        objMovimentoCaixa.lMovtoEstorno = .lMovtoEstorno
        objMovimentoCaixa.lMovtoTransf = .lMovtoTransf
        objMovimentoCaixa.lNumero = .lNumero
        objMovimentoCaixa.lNumMovto = .lNumMovto
        objMovimentoCaixa.lNumRefInterna = .lNumRefInterna
        objMovimentoCaixa.lSequencial = .lSequencial
        objMovimentoCaixa.sFavorecido = .sFavorecido
        objMovimentoCaixa.sHistorico = .sHistorico
        
    End With
    
    'Fecha o comando
    Call Comando_Fechar(lComando)

    MovimentosCaixa_Le_MovAbertCx = SUCESSO

    Exit Function

Erro_MovimentosCaixa_Le_MovAbertCx:

    MovimentosCaixa_Le_MovAbertCx = gErr

    Select Case gErr

        Case 116232
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 116233, 116234
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CAIXAS", gErr, objMovimentoCaixa.iCaixa)

        Case 116235
            'Registro não encontrado. Esse erro será tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 151518)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Caixas_Le1(ByVal objCaixa As ClassCaixa) As Long
'Lê os dados do Caixa, cujo código foi passado como parâmetro

Dim lErro As Long
Dim lComando As Long
Dim tCaixa As typeCaixa

On Error GoTo Erro_Caixas_Le1

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 105686

    'Inicializa as strings que receberão dados do BD
    tCaixa.sDescricao = String(STRING_CAIXA_DESCRICAO, 0)
    tCaixa.sNomeReduzido = String(STRING_CAIXA_NOME_REDUZIDO, 0)
    
    'Alterado por cyntia para incluir FilialEmpresa como parâmetro
    'Busca o Caixa no BD através do código passado como parâmetro
    lErro = Comando_Executar(lComando, "SELECT  FilialEmpresa, NomeReduzido, Descricao, Status, DataInicial, ProxSeqMov, POS, TEF, BoletoManual, Ativo, Teclado, SoOrcamento, UltimoOperador, SessaoStatus FROM Caixa WHERE Codigo = ?", tCaixa.iFilialEmpresa, tCaixa.sNomeReduzido, tCaixa.sDescricao, tCaixa.iStatus, tCaixa.dtDataInicial, tCaixa.lProxSeqMov, tCaixa.iPos, tCaixa.iTEF, tCaixa.iBoletoManual, tCaixa.iAtivo, tCaixa.iTeclado, tCaixa.iOrcamentoECF, tCaixa.iUltimoOperador, tCaixa.iSessaoStatus, objCaixa.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 105687

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 105688

    'Se não encontrou o Caixa => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 105689

    'Transfere para objCaixa os dados obtidos do BD
    objCaixa.iFilialEmpresa = tCaixa.iFilialEmpresa
    objCaixa.sNomeReduzido = tCaixa.sNomeReduzido
    objCaixa.sDescricao = tCaixa.sDescricao
    objCaixa.iStatus = tCaixa.iStatus
    objCaixa.dtDataInicial = tCaixa.dtDataInicial
    objCaixa.iTeclado = tCaixa.iTeclado
    objCaixa.lProxSeqMov = tCaixa.lProxSeqMov
    objCaixa.iPos = tCaixa.iPos
    objCaixa.iTEF = tCaixa.iTEF
    objCaixa.iBoletoManual = tCaixa.iBoletoManual
    objCaixa.iAtivo = tCaixa.iAtivo
    objCaixa.iOrcamentoECF = tCaixa.iOrcamentoECF
    objCaixa.iUltimoOperador = tCaixa.iUltimoOperador
    objCaixa.iSessaoStatus = tCaixa.iSessaoStatus
    
    'Fecha o comando
    Call Comando_Fechar(lComando)

    Caixas_Le1 = SUCESSO

    Exit Function

Erro_Caixas_Le1:

    Caixas_Le1 = gErr

    Select Case gErr

        Case 105686
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 105687, 105688
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CAIXAS", gErr, objCaixa.iCodigo)

        Case 105689
            'Caixa não encontrado. Esse erro será tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151519)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function


'???Transferir para RotinasLoja.ClassLojaSelect
Function Operador_Le1(ByVal objOperador As ClassOperador) As Long
'Le um operador sem filtrar pela filial

Dim lErro As Long
Dim lComando As Long
Dim tOperador As typeOperador

On Error GoTo Erro_Operador_Le1

    'Abertura de comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 102477
    
    'Inicilaização do Buffer
    tOperador.sSenha = String(STRING_OPERADOR_SENHA, 0)
    tOperador.sNome = String(STRING_OPERADOR_NOME, 0)
    
    'Seleciona o operador com código passado no obj
    'ATENÇÃO: Se existirem vários operadores com o mesmo código e em filiais diferentes, o sistema irá retornar o
    'operador da filial com código menor
    lErro = Comando_Executar(lComando, "SELECT  Codigo, FilialEmpresa, Nome, Senha, Desconto, LimiteDesconto, CodVendedor, Gerente FROM Operador WHERE Codigo = ? ORDER BY Codigo, FilialEmpresa ", tOperador.iCodigo, tOperador.iFilialEmpresa, tOperador.sNome, tOperador.sSenha, tOperador.iDesconto, tOperador.iLimiteDesconto, tOperador.iCodigoVendedor, tOperador.iGerente, objOperador.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 102478

    'Busca o primeiro operador encontrado
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 102479

    'Se não encontrou o operador =>erro
    If lErro = AD_SQL_SEM_DADOS Then gError 102480

    'Transfere para o obj os dados lidos do BD
    objOperador.iCodigo = tOperador.iCodigo
    objOperador.iFilialEmpresa = tOperador.iFilialEmpresa
    objOperador.sNome = tOperador.sNome
    objOperador.iDesconto = tOperador.iDesconto
    objOperador.iLimiteDesconto = tOperador.iLimiteDesconto
    objOperador.iGerente = tOperador.iGerente
    objOperador.sSenha = tOperador.sSenha
    objOperador.iCodigoVendedor = tOperador.iCodigoVendedor
        
    Call Comando_Fechar(lComando)

    Operador_Le1 = SUCESSO

    Exit Function

Erro_Operador_Le1:

    Operador_Le1 = gErr

        Select Case gErr

        Case 102477
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 102478, 102479
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_OPERADOR", gErr, objOperador.iCodigo)

        Case 102480 'Não existe na tabela de Operadores

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151520)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

'???Transferir para RotinasLoja.ClassLojaSelect
Function Operador_Le_NomeReduzido(ByVal objOperador As ClassOperador) As Long
'Lê o codigo do Operador a partir do Nome/FilialEmpresa passados como parâmetros

Dim lComando As Long
Dim lErro As Long
Dim tOperador As typeOperador

On Error GoTo Erro_Operador_Le_NomeReduzido

    'Inicia o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 116186
    
    'Inicializa as variáveis que recberão strings
    tOperador.sNome = String(STRING_OPERADOR_NOME, 0)
    tOperador.sSenha = String(STRING_OPERADOR_SENHA, 0)

    'Lê o operador a partir dos filtros Nome/FilialEmpresa passados no obj
    lErro = Comando_Executar(lComando, "SELECT  Codigo, Senha, Desconto, LimiteDesconto, CodVendedor, Gerente FROM Operador WHERE Nome = ? AND FilialEmpresa = ?", tOperador.iCodigo, tOperador.sSenha, tOperador.iDesconto, tOperador.iLimiteDesconto, tOperador.iCodigoVendedor, tOperador.iGerente, objOperador.sNome, objOperador.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 116187

    'Busca o primeiro operador
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 116188

    'se não encontrou, erro
    If lErro = AD_SQL_SEM_DADOS Then gError 116189

    'carrega o obj com os dados encontrados
    objOperador.iCodigo = tOperador.iCodigo
    objOperador.iDesconto = tOperador.iDesconto
    objOperador.iLimiteDesconto = tOperador.iLimiteDesconto
    objOperador.iGerente = tOperador.iGerente
    objOperador.sSenha = tOperador.sSenha
    objOperador.iCodigoVendedor = tOperador.iCodigoVendedor
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Operador_Le_NomeReduzido = SUCESSO
    
    Exit Function
    
Erro_Operador_Le_NomeReduzido:

    Operador_Le_NomeReduzido = gErr
    
    Select Case gErr
    
        Case 116186
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 116187, 116188
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_OPERADOR", gErr)
        
        Case 116189
        'Operador não encontrado. Esse erro deve ser tratado na rotina chamadora
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151521)
            
    End Select
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

'???Transferir para RotinasLoja.ClassLojaSelect
Function Operador_Le_NomeReduzido1(ByVal objOperador As ClassOperador) As Long
'Lê o codigo do Operador a partir do Nome passado como parâmetro
'Não faz filtro por filialempresa

Dim lComando As Long
Dim lErro As Long
Dim tOperador As typeOperador

On Error GoTo Erro_Operador_Le_NomeReduzido1

    'Inicia o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 102481
    
    'Inicializa as variáveis que recberão strings
    tOperador.sSenha = String(STRING_OPERADOR_SENHA, 0)

    'Lê o operador a partir do Nome passado no obj
    'ATENÇÃO: Se existirem vários operadores com o mesmo nome reduzido e em filiais diferentes, o sistema irá retornar o
    'operador da filial com código menor
    lErro = Comando_Executar(lComando, "SELECT Codigo, FilialEmpresa, Senha, Desconto, LimiteDesconto, CodVendedor, Gerente FROM Operador WHERE Nome = ?", tOperador.iCodigo, tOperador.iFilialEmpresa, tOperador.sSenha, tOperador.iDesconto, tOperador.iLimiteDesconto, tOperador.iCodigoVendedor, tOperador.iGerente, objOperador.sNome)
    If lErro <> AD_SQL_SUCESSO Then gError 102482

    'Busca o primeiro operador
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 102483

    'se não encontrou, erro
    If lErro = AD_SQL_SEM_DADOS Then gError 102484

    'carrega o obj com os dados encontrados
    objOperador.iCodigo = tOperador.iCodigo
    objOperador.iFilialEmpresa = tOperador.iFilialEmpresa
    objOperador.iDesconto = tOperador.iDesconto
    objOperador.iLimiteDesconto = tOperador.iLimiteDesconto
    objOperador.iGerente = tOperador.iGerente
    objOperador.sSenha = tOperador.sSenha
    objOperador.iCodigoVendedor = tOperador.iCodigoVendedor
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Operador_Le_NomeReduzido1 = SUCESSO
    
    Exit Function
    
Erro_Operador_Le_NomeReduzido1:

    Operador_Le_NomeReduzido1 = gErr
    
    Select Case gErr
    
        Case 102481
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 102482, 102483
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_OPERADOR", gErr)
        
        Case 102484
        'Operador não encontrado. Esse erro deve ser tratado na rotina chamadora
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151522)
            
    End Select
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function ControleLogCCBack_Le_Limites(ByVal iFilialEmpresa As Integer, lNumIntDocInicial As Long, lNumIntDocFinal As Long) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(1 To 2) As Long

On Error GoTo Erro_ControleLogCCBack_Le_Limites
    
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 118905
    Next
    
    'selecionar o ultimo registro que contem info sobre os dados do log que foram enviados para o backfoffice
    lErro = Comando_Executar(alComando(1), "SELECT  NumIntDocFinal FROM ControleLogCCBack WHERE FilialEmpresa = ? ORDER BY Sequencial DESC", lNumIntDocInicial, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 118906

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 118907

    If lErro = AD_SQL_SEM_DADOS Then lNumIntDocInicial = 0
    
    'selecionar o ultimo registro que contem info sobre os dados do log que foram enviados para o backfoffice
    lErro = Comando_Executar(alComando(2), "SELECT  NumIntDoc FROM Log WHERE NumIntDoc > ? ORDER BY NumIntDoc DESC", lNumIntDocFinal, lNumIntDocInicial)
    If lErro <> AD_SQL_SUCESSO Then gError 118908

    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 118909

    If lErro = AD_SQL_SEM_DADOS Then gError 118910
        
    lNumIntDocInicial = lNumIntDocInicial + 1
    
    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    ControleLogCCBack_Le_Limites = SUCESSO

    Exit Function

Erro_ControleLogCCBack_Le_Limites:

    ControleLogCCBack_Le_Limites = gErr

    Select Case gErr

        Case 118905
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 118906, 118907
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTROLELOGCCBACK", gErr)
        
        Case 118908, 118909
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LOG", gErr)

        Case 118910
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151523)

    End Select

    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Function Cheque_Le_SequencialLoja(ByVal objCheque As ClassChequePre) As Long
'Função que lê as Informações do Cheque passado como parâmetro no banco de dados, Carregando no objCheque

Dim lErro As Long
Dim lComando As Long
Dim lCupomFiscal As Long
Dim tCheque As typeChequePre
Dim iTipoMovcx As Integer

On Error GoTo Erro_Cheque_Le_SequencialLoja

    'Inicia a Abertura de o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 126046

    'Inicializa o Buffer das variáveis
    tCheque.sAgencia = String(STRING_BUFFER_MAX_TEXTO, 0)
    tCheque.sContaCorrente = String(STRING_BUFFER_MAX_TEXTO, 0)
    tCheque.sCPFCGC = String(STRING_BUFFER_MAX_TEXTO, 0)
    
    'Procura o Cheque no BD a partir do código e a FilialEmpresaLoja
    lErro = Comando_Executar(lComando, "SELECT NumIntCheque, Cliente, Filial , Banco , ChequePre.Agencia , ContaCorrente , ChequePre.Numero, DataDeposito , ChequePre.Valor , NumBordero , ChequePre.CPFCGC , NumMovtoCaixa , Aprovado , NaoEspecificado , NumBorderoLoja , ChequePre.FilialEmpresa , SequencialLoja , SequencialBack,  ChequePre.Status ,ECF , MovimentosCaixa.CupomFiscal, MovimentosCaixa.Tipo, ChequePre.Localizacao, ChequePre.NumBorderoLojaBanco, ChequePre.Caixa, ChequePre.FilialEmpresaLoja, ChequePre.NumMovtoSangria, ChequePre.SequencialCaixa, ChequePre.FilialEmpresa, ChequePre.DataEmissao  FROM CupomFiscal , ChequePre , MovimentosCaixa WHERE ChequePre.FilialEmpresaLoja = ? AND ChequePre.SequencialLoja = ? AND ChequePre.NumMovtoCaixa = MovimentosCaixa.NumMovto AND MovimentosCaixa.CupomFiscal = CupomFiscal.NumIntDoc AND MovimentosCaixa.FilialEmpresa = ChequePre.FilialEmpresaLoja AND ChequePre.Status <> 5", _
    tCheque.lNumIntCheque, tCheque.lCliente, tCheque.iFilial, tCheque.iBanco, tCheque.sAgencia, tCheque.sContaCorrente, tCheque.lNumero, tCheque.dtDataDeposito, tCheque.dValor, tCheque.lNumBordero, tCheque.sCPFCGC, tCheque.lNumMovtoCaixa, tCheque.iAprovado, tCheque.iNaoEspecificado, tCheque.lNumBorderoLoja, tCheque.iFilialEmpresa, tCheque.lSequencialLoja, tCheque.lSequencialBack, tCheque.iStatus, tCheque.iECF, lCupomFiscal, iTipoMovcx, tCheque.iLocalizacao, tCheque.lNumBorderoLojaBanco, tCheque.iCaixa, tCheque.iFilialEmpresaLoja, tCheque.lNumMovtoSangria, tCheque.lSequencialCaixa, tCheque.iFilialEmpresa, tCheque.dtDataEmissao, objCheque.iFilialEmpresaLoja, objCheque.lSequencialLoja)
    If lErro <> AD_SQL_SUCESSO Then gError 126047

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 126048

    'Se não encontrou o Cheque => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 126049
        
    'Passa para objCheque os valores lidos do BD
    objCheque.dtDataDeposito = tCheque.dtDataDeposito
    objCheque.dValor = tCheque.dValor
    objCheque.iAprovado = tCheque.iAprovado
    objCheque.iBanco = tCheque.iBanco
    objCheque.iFilial = tCheque.iFilial
    objCheque.iFilialEmpresa = tCheque.iFilialEmpresa
    objCheque.iNaoEspecificado = tCheque.iNaoEspecificado
    objCheque.lCliente = tCheque.lCliente
    objCheque.lNumBordero = tCheque.lNumBordero
    objCheque.lNumBorderoLoja = tCheque.lNumBorderoLoja
    objCheque.lNumero = tCheque.lNumero
    objCheque.lNumIntCheque = tCheque.lNumIntCheque
    objCheque.lNumMovtoCaixa = tCheque.lNumMovtoCaixa
    objCheque.lSequencialBack = tCheque.lSequencialBack
    objCheque.lSequencialLoja = tCheque.lSequencialLoja
    objCheque.sAgencia = tCheque.sAgencia
    objCheque.sContaCorrente = tCheque.sContaCorrente
    objCheque.sCPFCGC = tCheque.sCPFCGC
    objCheque.lCupomFiscal = lCupomFiscal
    objCheque.iECF = tCheque.iECF
    objCheque.iStatus = tCheque.iStatus
    objCheque.iLocalizacao = tCheque.iLocalizacao
    objCheque.lNumBorderoLojaBanco = tCheque.lNumBorderoLojaBanco
    objCheque.iCaixa = tCheque.iCaixa
    objCheque.iFilialEmpresaLoja = tCheque.iFilialEmpresaLoja
    objCheque.lNumMovtoSangria = tCheque.lNumMovtoSangria
    objCheque.lSequencialCaixa = tCheque.lSequencialCaixa
    objCheque.dtDataEmissao = tCheque.dtDataEmissao
    
    'Fecha o comando
    Call Comando_Fechar(lComando)

    Cheque_Le_SequencialLoja = SUCESSO

    Exit Function

Erro_Cheque_Le_SequencialLoja:

    Cheque_Le_SequencialLoja = gErr

    Select Case gErr

        Case 126046
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 126047, 126048
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CHEQUE4", gErr, objCheque.iFilialEmpresaLoja, objCheque.lSequencialLoja)

        Case 126049

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151524)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ControleLogCCBack_Le_De_Ate(ByVal lDe As Long, ByVal lAte As Long, colControleLogCCBack As Collection, ByVal iFilialEmpresa As Integer) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(1 To 2) As Long
Dim lSeq As Long
Dim lNumIntIni As Long
Dim lNumIntFim As Long
Dim dtData As Date
Dim objControleLogCCBack As ClassControleLogCCBack
Dim lRegistrosLidos As Long
Dim lTotal As Long

On Error GoTo Erro_ControleLogCCBack_Le_De_Ate
    
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 126081
    Next
    
    If lDe = 0 And lAte = 0 Then
    
        'selecionar o ultimo registro que contem info sobre os dados do log que foram enviados para o backfoffice
        lErro = Comando_Executar(alComando(1), "SELECT  Count(*) FROM ControleLogCCBack WHERE FilialEmpresa = ?", lTotal, iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 126106
    
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 126107
        
        If lTotal > NUM_MAX_SEQ_CONTROLELOGCCBACK Then gError 126108
        
        'selecionar o ultimo registro que contem info sobre os dados do log que foram enviados para o backfoffice
        lErro = Comando_Executar(alComando(1), "SELECT  Sequencial, NumIntDocInicial, NumIntDocFinal, Data FROM ControleLogCCBack WHERE FilialEmpresa = ? ORDER BY Sequencial DESC", lSeq, lNumIntIni, lNumIntFim, dtData, iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 126109
    
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 126110
    
    ElseIf lDe <> 0 And lAte <> 0 Then
    
        'selecionar o ultimo registro que contem info sobre os dados do log que foram enviados para o backfoffice
        lErro = Comando_Executar(alComando(1), "SELECT  Count(*) FROM ControleLogCCBack WHERE FilialEmpresa = ? AND Sequencial >= ? And Sequencial <= ?", lTotal, iFilialEmpresa, lDe, lAte)
        If lErro <> AD_SQL_SUCESSO Then gError 126082
    
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 126083
        
        If lTotal > NUM_MAX_SEQ_CONTROLELOGCCBACK Then gError 126084
        
        'selecionar o ultimo registro que contem info sobre os dados do log que foram enviados para o backfoffice
        lErro = Comando_Executar(alComando(1), "SELECT  Sequencial, NumIntDocInicial, NumIntDocFinal, Data FROM ControleLogCCBack WHERE FilialEmpresa = ? AND Sequencial >= ? AND Sequencial <= ? ORDER BY Sequencial DESC", lSeq, lNumIntIni, lNumIntFim, dtData, iFilialEmpresa, lDe, lAte)
        If lErro <> AD_SQL_SUCESSO Then gError 126085
    
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 126095
    
    ElseIf lDe <> 0 Then
    
        'selecionar o ultimo registro que contem info sobre os dados do log que foram enviados para o backfoffice
        lErro = Comando_Executar(alComando(1), "SELECT  Count(*) FROM ControleLogCCBack WHERE FilialEmpresa = ? AND Sequencial >= ? ", lTotal, iFilialEmpresa, lDe)
        If lErro <> AD_SQL_SUCESSO Then gError 126086
    
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 126087
        
        If lTotal > NUM_MAX_SEQ_CONTROLELOGCCBACK Then gError 126088
        
        'selecionar o ultimo registro que contem info sobre os dados do log que foram enviados para o backfoffice
        lErro = Comando_Executar(alComando(1), "SELECT  Sequencial, NumIntDocInicial, NumIntDocFinal, Data FROM ControleLogCCBack WHERE FilialEmpresa = ? AND Sequencial >= ? ORDER BY Sequencial DESC", lSeq, lNumIntIni, lNumIntFim, dtData, iFilialEmpresa, lDe)
        If lErro <> AD_SQL_SUCESSO Then gError 126089
    
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 126096
    
    ElseIf lAte <> 0 Then
    
        'selecionar o ultimo registro que contem info sobre os dados do log que foram enviados para o backfoffice
        lErro = Comando_Executar(alComando(1), "SELECT  Count(*) FROM ControleLogCCBack WHERE FilialEmpresa = ? AND Sequencial <= ? ", lTotal, iFilialEmpresa, lAte)
        If lErro <> AD_SQL_SUCESSO Then gError 126090
    
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 126091
        
        If lTotal > NUM_MAX_SEQ_CONTROLELOGCCBACK Then gError 126092
        
        'selecionar o ultimo registro que contem info sobre os dados do log que foram enviados para o backfoffice
        lErro = Comando_Executar(alComando(1), "SELECT  Sequencial, NumIntDocInicial, NumIntDocFinal, Data FROM ControleLogCCBack WHERE FilialEmpresa = ? AND Sequencial <= ? ORDER BY Sequencial DESC", lSeq, lNumIntIni, lNumIntFim, dtData, iFilialEmpresa, lAte)
        If lErro <> AD_SQL_SUCESSO Then gError 126093
    
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 126097
    
    End If
    
    
    Do While lErro = AD_SQL_SUCESSO
                
        Set objControleLogCCBack = New ClassControleLogCCBack
        
        objControleLogCCBack.dtData = dtData
        objControleLogCCBack.lNumIntDocInicial = lNumIntIni
        objControleLogCCBack.lNumIntDocFinal = lNumIntFim
        objControleLogCCBack.lSequencial = lSeq
        
        colControleLogCCBack.Add objControleLogCCBack
        
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 126094
        
    Loop
    
    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    ControleLogCCBack_Le_De_Ate = SUCESSO

    Exit Function

Erro_ControleLogCCBack_Le_De_Ate:

    ControleLogCCBack_Le_De_Ate = gErr

    Select Case gErr

        Case 126081
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 126082, 126083, 126085, 126086, 126087, 126089, 126090, 126091, 126093, 126094, 126095 To 126097, 126106, 126107, 126109, 126110
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTROLELOGCCBACK", gErr)
        
        Case 126084, 126088, 126092, 126108
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151525)

    End Select

    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Function ComissoesCF_Le(objCupomFiscal As ClassCupomFiscal, colComissoes As Collection) As Long

Dim lErro As Long
Dim lComando As Long
Dim tComissoesCF As typeComissoesCF
Dim objComissoesCF As ClassComissoesCF
Dim alComando(1 To 1) As Long
Dim iIndice As Integer

On Error GoTo Erro_ComissoesCF_Le

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 126321
    Next
    
    lErro = Comando_Executar(alComando(1), "SELECT CodVendedor, ValorBase, ValorComissao, Indireta FROM ComissoesCupomFiscal WHERE FilialEmpresa = ? AND NumIntDoc = ? ORDER BY CodVendedor", tComissoesCF.iCodVendedor, tComissoesCF.dValorBase, tComissoesCF.dValorComissao, tComissoesCF.iIndireta, objCupomFiscal.iFilialEmpresa, objCupomFiscal.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 126318
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 126319
    
    Do While lErro <> AD_SQL_SEM_DADOS
    
        'Instancia um novo objComissoesCF
        Set objComissoesCF = New ClassComissoesCF
        
        'Transfere os dados lidos para o obj
        With tComissoesCF
            
            objComissoesCF.iCodVendedor = .iCodVendedor
            objComissoesCF.dValorBase = .dValorBase
            objComissoesCF.dValorComissao = .dValorComissao
            objComissoesCF.iIndireta = .iIndireta

        End With
    
        'Adiciona o obj à coleção
        colComissoes.Add objComissoesCF
            
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 126320
        
    Loop
    
    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    ComissoesCF_Le = SUCESSO
    
    Exit Function
    
Erro_ComissoesCF_Le:

    ComissoesCF_Le = gErr
    
    Select Case gErr
    
        Case 126318, 126319, 126320
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMISSOESCF", gErr, objCupomFiscal.iFilialEmpresa, objCupomFiscal.lNumIntDoc)
            
        Case 126321
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 151526)
            
    End Select
    
    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Function LojaConfig_Le1(ByVal objLojaConfig As ClassLojaConfig) As Long
'Função responsável pela leitura dos dados

Dim lErro As Long
Dim sConteudo As String
Dim alComando(0 To 1) As Long
Dim iIndice As Integer

On Error GoTo Erro_LojaConfig_Le1

    'Abre os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 126358
    Next

    sConteudo = String(STRING_CONTEUDO, 0)

    'Seleciona codigo e filial na tabela LojaConfig de acordo com o codigo passado como parametro
    lErro = Comando_Executar(alComando(0), "SELECT Conteudo FROM LojaConfig WHERE Codigo = ? And FilialEmpresa = ?", sConteudo, objLojaConfig.sCodigo, objLojaConfig.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 126359

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 126360

    'Nao encontrou registro em LojaConfig
    If lErro = AD_SQL_SEM_DADOS Then gError 126361

    objLojaConfig.sConteudo = sConteudo

    'Fecha os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next

    LojaConfig_Le1 = SUCESSO

    Exit Function

Erro_LojaConfig_Le1:

    LojaConfig_Le1 = gErr

    Select Case gErr

        Case 126358
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 126359, 126360
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LOJACONFIG", gErr, objLojaConfig.sCodigo)

        Case 126361
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151527)

    End Select
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'Incluído por Luiz Nogueira em 22/04/04
Public Function MovimentosCaixa_Le_MovAbertSessao(ByVal objMovimentoCaixa As ClassMovimentoCaixa) As Long
'Lê em MovimentosCaixa os dados do último movimento de abertura de sessão para o caixa passado como parâmetro na data e filial passadas.

Dim lErro As Long
Dim lComando As Long
Dim tMovCaixa As typeMovimentoCaixa

On Error GoTo Erro_MovimentosCaixa_Le_MovAbertSessao

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 127101

    'Inicializa as strings que receberão dados do BD
    tMovCaixa.sFavorecido = String(STRING_FAVORECIDO, 0)
    tMovCaixa.sHistorico = String(STRING_HISTORICO1, 0)
    
    With tMovCaixa
        
        'faz um select no Bd
        lErro = Comando_Executar(lComando, "SELECT NumMovto, Caixa,  Operador, Sequencial, Tipo, AdmMeioPagto, Parcelamento, TipoCartao, Numero, Valor, Historico, Favorecido, CupomFiscal,NumRefInterna, MovtoTransf, MovtoEstorno, Gerente, Hora FROM MovimentosCaixa WHERE DataMovimento = ? AND FilialEmpresa = ? AND Caixa = ? AND Tipo = ? ORDER BY Hora DESC", _
        .lNumMovto, .iCaixa, .iCodOperador, .lSequencial, .iTipo, .iAdmMeioPagto, .iParcelamento, .iTipoCartao, .lNumero, .dValor, .sHistorico, .sFavorecido, .lCupomFiscal, .lNumRefInterna, .lMovtoTransf, .lMovtoEstorno, .iGerente, .dHora, objMovimentoCaixa.dtDataMovimento, objMovimentoCaixa.iFilialEmpresa, objMovimentoCaixa.iCaixa, MOVIMENTO_CAIXA_SESSAO_ABERTURA)
        If lErro <> AD_SQL_SUCESSO Then gError 127102
    
        'busca o 1º registro
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 127103
    
        'Se não encontrou os dados => erro
        If lErro = AD_SQL_SEM_DADOS Then gError 127104
    
        'guarda os dados lidos no objMovimentoCaixa
        objMovimentoCaixa.dHora = .dHora
        objMovimentoCaixa.dValor = .dValor
        objMovimentoCaixa.iAdmMeioPagto = .iAdmMeioPagto
        objMovimentoCaixa.iCaixa = .iCaixa
        objMovimentoCaixa.iCodOperador = .iCodOperador
        objMovimentoCaixa.iCodConta = .iCodConta
        objMovimentoCaixa.iGerente = .iGerente
        objMovimentoCaixa.iParcelamento = .iParcelamento
        objMovimentoCaixa.iTipo = .iTipo
        objMovimentoCaixa.iTipoCartao = .iTipoCartao
        objMovimentoCaixa.lCupomFiscal = .lCupomFiscal
        objMovimentoCaixa.lMovtoEstorno = .lMovtoEstorno
        objMovimentoCaixa.lMovtoTransf = .lMovtoTransf
        objMovimentoCaixa.lNumero = .lNumero
        objMovimentoCaixa.lNumMovto = .lNumMovto
        objMovimentoCaixa.lNumRefInterna = .lNumRefInterna
        objMovimentoCaixa.lSequencial = .lSequencial
        objMovimentoCaixa.sFavorecido = .sFavorecido
        objMovimentoCaixa.sHistorico = .sHistorico
        
    End With
    
    'Fecha o comando
    Call Comando_Fechar(lComando)

    MovimentosCaixa_Le_MovAbertSessao = SUCESSO

    Exit Function

Erro_MovimentosCaixa_Le_MovAbertSessao:

    MovimentosCaixa_Le_MovAbertSessao = gErr

    Select Case gErr

        Case 127101
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 127102, 1027103
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CAIXAS", gErr, objMovimentoCaixa.iCaixa)

        Case 127104
            'Registro não encontrado. Esse erro será tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 151528)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Verifica_Nome_Arquivo(ByVal sNome As String) As Long

Dim lComando As Long
Dim lComando1 As Long
Dim lErro As Long
Dim iCaixa As Integer
Dim iCaixa1 As Integer
Dim lSeq As Long
Dim lSeq1 As Long
Dim iUnder1 As Integer
Dim iUnder2 As Integer
Dim iUnder3 As Integer
Dim iPonto As Integer
Dim iFilialEmpresa As Integer
Dim lEmpresa As Long

On Error GoTo Erro_Verifica_Nome_Arquivo
    
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 133653
    
    'posiciona no _ do CC_
    iUnder1 = InStr(1, sNome, "_")
    If iUnder1 = 0 Then gError 133654
    
    'posiciona no _ depois da empresa
    iUnder2 = InStr(iUnder1 + 1, sNome, "_")
    If iUnder2 = 0 Then gError 133655
    
    lEmpresa = StrParaLong(left(sNome, iUnder2 - 1))
    
    If lEmpresa <> glEmpresa Then gError 133651
    
    'posiciona no _ depois da filial
    iUnder3 = InStr(iUnder2 + 1, sNome, "_")
    If iUnder3 = 0 Then gError 133656
    
    iFilialEmpresa = StrParaInt(Mid(sNome, iUnder2 + 1, iUnder3 - (iUnder2 + 1)))

    'posiciona no _ depois da filial
    iPonto = InStr(iUnder3 + 1, sNome, "_")
    If iPonto = 0 Then gError 133657

    lSeq = StrParaLong(Mid(sNome, iUnder3 + 1, iPonto - (iUnder3 + 1)))

    'selecionar o ultimo registro que contem info sobre os dados do log que foram enviados para o backfoffice
    lErro = Comando_Executar(lComando, "SELECT Sequencial FROM ControleLogBack WHERE FilialEmpresa = ? ORDER BY Sequencial DESC", lSeq1, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 133658

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 133659

    'se nao tem registro ainda cadastrado ==> proximo sequencial = 1 e ultimo registro processado = 0 (lRegistroFinal)
    If lErro = AD_SQL_SEM_DADOS Then
        If lSeq <> 1 Then gError 133660
    'se o ultimo registro cadastrado tem lRegistroFinalArq preenchido ==> um novo arquivo com novo sequencial deve ser lido
    Else
        If lSeq <> lSeq1 + 1 Then gError 133661
    End If

    Call Comando_Fechar(lComando)
    
    Verifica_Nome_Arquivo = SUCESSO

    Exit Function

Erro_Verifica_Nome_Arquivo:

    Verifica_Nome_Arquivo = gErr

    Select Case gErr

        Case 133653
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 133654 To 133657
            Call Rotina_Erro(vbOKOnly, "ERRO_ARQUIVO_INVALIDO1", gErr)
                
        Case 133660
            Call Rotina_Erro(vbOKOnly, "ERRO_SEQUENCIAL_DIFERE", gErr, lSeq, 1)
                
        Case 133661
            Call Rotina_Erro(vbOKOnly, "ERRO_SEQUENCIAL_DIFERE", gErr, lSeq, lSeq1 + 1)

        Case 133658, 133659
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTROLELOGBACK", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151529)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function


Function Verifica_Nome_Arquivo1(ByVal sNome As String) As Long

Dim lComando As Long
Dim lComando1 As Long
Dim lErro As Long
Dim iCaixa As Integer
Dim iCaixa1 As Integer
Dim lSeq As Long
Dim lSeq1 As Long
Dim iUnder1 As Integer
Dim iUnder2 As Integer
Dim iUnder3 As Integer
Dim iPonto As Integer
Dim iFilialEmpresa As Integer
Dim lEmpresa As Long

On Error GoTo Erro_Verifica_Nome_Arquivo1
    
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 133422
        
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 133423
        
    iUnder1 = InStr(1, sNome, "_")
    If iUnder1 = 0 Then gError 133424
    
    lEmpresa = StrParaLong(left(sNome, iUnder1 - 1))
    
    If lEmpresa <> glEmpresa Then gError 133645
    
    iUnder2 = InStr(iUnder1 + 1, sNome, "_")
    If iUnder2 = 0 Then gError 133425
    
    iFilialEmpresa = StrParaInt(Mid(sNome, iUnder1 + 1, iUnder2 - (iUnder1 + 1)))
    
    iUnder3 = InStr(iUnder2 + 1, sNome, "_")
    If iUnder3 = 0 Then gError 133426
    
    iCaixa = StrParaInt(Mid(sNome, iUnder2 + 1, iUnder3 - (iUnder2 + 1)))
    
    iPonto = InStr(iUnder3 + 1, sNome, ".")
    If iPonto = 0 Then gError 133427
    
    lSeq = StrParaLong(Mid(sNome, iUnder3 + 1, iPonto - (iUnder3 + 1)))
    
    'Pega os dados da tabela  de controle
    lErro = Comando_Executar(lComando1, "SELECT Sequencial FROM ControleLogCaixaCC WHERE Caixa = ? AND FilialEmpresa = ? Order By Sequencial desc", lSeq1, iCaixa, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 133431

    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 133432
    
    If lErro = AD_SQL_SUCESSO Then
        'se o sequencial do arquivo a ser lido não for o próximo --> erro.
        If lSeq <> lSeq1 + 1 Then gError 133433
    Else
        If lSeq <> 1 Then gError 133434
    End If
    
    lErro = Comando_Fechar(lComando)
    lErro = Comando_Fechar(lComando1)
    
    Verifica_Nome_Arquivo1 = SUCESSO

    Exit Function

Erro_Verifica_Nome_Arquivo1:

    Verifica_Nome_Arquivo1 = gErr

    Select Case gErr

        Case 133422, 133423
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 133424 To 133427, 133645
            Call Rotina_Erro(vbOKOnly, "ERRO_ARQUIVO_INVALIDO1", gErr)
                
        Case 133433
            Call Rotina_Erro(vbOKOnly, "ERRO_SEQUENCIAL_DIFERE", gErr, lSeq, lSeq1 + 1)

        Case 133434
            Call Rotina_Erro(vbOKOnly, "ERRO_SEQUENCIAL_DIFERE", gErr, lSeq, 1)
                
        Case 133431, 133432
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTROLELOGCAIXACC", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151530)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Function Caixa_Le_Prox_Seq(ByVal iFilialEmpresa As Integer, ByVal iCaixa As Integer, lSeq As Long) As Long

Dim lErro As Long
Dim alComando(1 To 1) As Long
Dim iIndice As Integer

On Error GoTo Erro_Caixa_Le_Prox_Seq
    
    'Abre os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 133514
    Next
    
    'Pega os dados da tabela  de controle
    lErro = Comando_Executar(alComando(1), "SELECT Sequencial FROM ControleLogCaixaCC WHERE Caixa = ? AND FilialEmpresa = ? Order By Sequencial desc", lSeq, iCaixa, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 133515

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 133516
    
    If lErro = AD_SQL_SUCESSO Then
        lSeq = lSeq + 1
    Else
        lSeq = 1
    End If
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Caixa_Le_Prox_Seq = SUCESSO

    Exit Function

Erro_Caixa_Le_Prox_Seq:

    Caixa_Le_Prox_Seq = gErr

    Select Case gErr

        Case 133514
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 133515, 133516
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTROLELOGCAIXACC1", gErr, iFilialEmpresa, iCaixa)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151531)

    End Select

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function AdmMeioPagto_Le_Todas1(ByVal iFilialEmpresa As Integer, ByVal colAdmMeioPagto As Collection) As Long
'Lê em colAdmMeioPagto as administradoras de meios de pagamentos cadastrados

Dim lErro As Long
Dim lComando As Long
Dim objAdmMeioPagto As ClassAdmMeioPagto
Dim tAdmMeioPagto As typeAdmMeioPagto

On Error GoTo Erro_AdmMeioPagto_Le_Todas1

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 1183118
    
    'Inicializa o buffer nome
    tAdmMeioPagto.sNome = String(STRING_ADMMEIOPAGTO_NOME, 0)

    'Seleciona código e nome dos meios de pagamentos da tabela AdmMeioPagto
    lErro = Comando_Executar(lComando, "SELECT Codigo, Nome, TipoMeioPagto, Rede FROM AdmMeioPagto WHERE FilialEmpresa = ? ORDER BY Codigo ", tAdmMeioPagto.iCodigo, tAdmMeioPagto.sNome, tAdmMeioPagto.iTipoMeioPagto, tAdmMeioPagto.iRede, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 183115

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 183116

    Do While lErro = AD_SQL_SUCESSO
        
        'Inicializa  o obj relacionado a administradora
        Set objAdmMeioPagto = New ClassAdmMeioPagto
        
        objAdmMeioPagto.iCodigo = tAdmMeioPagto.iCodigo
        
        objAdmMeioPagto.sNome = tAdmMeioPagto.sNome
        objAdmMeioPagto.iTipoMeioPagto = tAdmMeioPagto.iTipoMeioPagto
        objAdmMeioPagto.iRede = tAdmMeioPagto.iRede
        
        'Adcionar na Coleção para preenchimento da ListBox Admnistradoras
        colAdmMeioPagto.Add objAdmMeioPagto

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 183117

     Loop

    'Fecha o comando
    Call Comando_Fechar(lComando)

    AdmMeioPagto_Le_Todas1 = SUCESSO

    Exit Function

Erro_AdmMeioPagto_Le_Todas1:

    AdmMeioPagto_Le_Todas1 = gErr

    Select Case gErr

        Case 183115 To 183117
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ADMMEIOPAGTO1", gErr)

        Case 183118
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183119)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function RelProdRanking_Obter_ValorFrete(ByVal iFilialEmpresa As Integer, ByVal dtDataDe As Date, ByVal dtDataAte As Date, ByVal lNumIntRel As Long, dValorFrete As Double) As Long
'Para obter o valor frete de uma filialempresa em um periodo

Dim lErro As Long, iIndice As Integer
Dim alComando(0 To 1) As Long, sSQL As String
Dim dValorFreteNF As Double, dValorFreteCF As Double

On Error GoTo Erro_RelProdRanking_Obter_ValorFrete

    'Executa a abertura do Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 216153
    Next
    
    If dtDataAte = DATA_NULA Then dtDataAte = Date
    
'    sSQL = "SELECT SUM(TI.ValorFreteItem) FROM TiposDocInfo AS T, NFiscal AS N, TributacaoDocItem AS TI "
'    sSQL = sSQL & "WHERE T.Codigo = N.TipoNFiscal AND TI.TipoDoc = 0 AND TI.NumIntDoc = N.NumIntDoc AND "
'    sSQL = sSQL & "T.Faturamento=1 AND N.Status <> 7 AND N.FilialEmpresa = ? AND N.DataEmissao BETWEEN ? AND ? AND TI.Produto IN ("
'    sSQL = sSQL & "SELECT Produto FROM ProdutosRanking WHERE NumIntRel = ?)"
'
'    lErro = Comando_Executar(alComando(0), sSQL, dValorFreteNF, iFilialEmpresa, dtDataDe, dtDataAte, lNumIntRel)
'    If lErro <> AD_SQL_SUCESSO Then gError 216154
'
'    lErro = Comando_BuscarPrimeiro(alComando(0))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 216155
    
    sSQL = "SELECT SUM(C.ValorAcrescimo-C.ValorDesconto) FROM CFVendDtBase AS C "
    'sSQL = sSQL & "WHERE C.Tipo=2 AND C.Status = 2 AND C.FilialEmpresa = ? AND C.DataEmissao BETWEEN ? AND ?  "
    sSQL = sSQL & "WHERE ((C.Status = 2 AND C.Tipo = 2 ) OR (C.Status <> 7 AND C.Tipo = 1)) AND C.FilialEmpresa = ? AND C.DtBase BETWEEN ? AND ?  "
    
    lErro = Comando_Executar(alComando(1), sSQL, dValorFreteCF, iFilialEmpresa, dtDataDe, dtDataAte)
    If lErro <> AD_SQL_SUCESSO Then gError 216154
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 216155
    
    dValorFrete = Arredonda_Moeda(dValorFreteCF + dValorFreteNF)
    
    'fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    RelProdRanking_Obter_ValorFrete = SUCESSO

    Exit Function

Erro_RelProdRanking_Obter_ValorFrete:

    RelProdRanking_Obter_ValorFrete = gErr

    Select Case gErr

        Case 216153
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 216154, 216155
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", gErr)
        
'        Case 216156, 216157
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CUPOMFISCAL", gErr)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 216158)

    End Select

    'fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    Exit Function

End Function

Public Function CupomFiscal_Obter_ValorFrete(ByVal iFilialEmpresa As Integer, ByVal iCaixa As Integer, ByVal dtData As Date, dValorFrete As Double) As Long
'Para obter o valor frete de uma filialempresa em um periodo

Dim lErro As Long, iIndice As Integer
Dim alComando(0 To 1) As Long, sSQL As String
Dim dValorFreteNF As Double, dValorFreteCF As Double

On Error GoTo Erro_CupomFiscal_Obter_ValorFrete

    'Executa a abertura do Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 216153
    Next
    
    If iCaixa = 0 Then
        sSQL = "SELECT SUM(C.ValorAcrescimo-C.ValorDesconto) FROM CFVendDtBase AS C "
        sSQL = sSQL & "WHERE ((C.Status = 2 AND C.Tipo = 2 ) OR (C.Status <> 7 AND C.Tipo = 1)) AND C.FilialEmpresa = ? AND C.DtBase = ?  "
        
        lErro = Comando_Executar(alComando(1), sSQL, dValorFreteCF, iFilialEmpresa, dtData)
        If lErro <> AD_SQL_SUCESSO Then gError 216154
    Else
        sSQL = "SELECT SUM(C.ValorAcrescimo-C.ValorDesconto) FROM CFVendDtBase AS C "
        sSQL = sSQL & "WHERE ((C.Status = 2 AND C.Tipo = 2 ) OR (C.Status <> 7 AND C.Tipo = 1)) AND C.FilialEmpresa = ? AND C.Caixa = ? AND C.DtBase = ?  "
        
        lErro = Comando_Executar(alComando(1), sSQL, dValorFreteCF, iFilialEmpresa, iCaixa, dtData)
        If lErro <> AD_SQL_SUCESSO Then gError 216154
    End If
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 216155
    
    dValorFrete = Arredonda_Moeda(dValorFreteCF)
    
    'fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    CupomFiscal_Obter_ValorFrete = SUCESSO

    Exit Function

Erro_CupomFiscal_Obter_ValorFrete:

    CupomFiscal_Obter_ValorFrete = gErr

    Select Case gErr

        Case 216153
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 216154, 216155
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", gErr)
        
'        Case 216156, 216157
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CUPOMFISCAL", gErr)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 216158)

    End Select

    'fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    Exit Function

End Function

Function SldDiaMeioPagtoCx_Le_AcumPeriodo(ByVal iFilialEmpresa As Integer, ByVal iCxDe As Integer, ByVal iCxAte As Integer, ByVal iTipoDe As Integer, ByVal iTipoAte As Integer, ByVal dtDataDe As Date, ByVal dtDataAte As Date, dValor As Double, Optional ByVal bTabCust As Boolean = False) As Long

Dim lErro As Long, sSQL As String
Dim alComando(0 To 1) As Long, iIndice As Integer

On Error GoTo Erro_SldDiaMeioPagtoCx_Le_AcumPeriodo
   
    'abre o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = SUCESSO Then gError 216240
    Next
   
    lErro = SldDiaMeioPagtoCx_Le_AcumPeriodo_SQL(iFilialEmpresa, iCxDe, iCxAte, iTipoDe, iTipoAte, dtDataDe, dtDataAte, bTabCust, sSQL)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = SldDiaMeioPagtoCx_Le_AcumPeriodo_Int(alComando(0), sSQL, iFilialEmpresa, iCxDe, iCxAte, iTipoDe, iTipoAte, dtDataDe, dtDataAte, dValor)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = Comando_BuscarProximo(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 216241
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next
    
    SldDiaMeioPagtoCx_Le_AcumPeriodo = SUCESSO
     
    Exit Function
    
Erro_SldDiaMeioPagtoCx_Le_AcumPeriodo:

    SldDiaMeioPagtoCx_Le_AcumPeriodo = gErr
     
    Select Case gErr
          
        Case 216240
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 216241, 216243
            Call Rotina_Erro(vbOKOnly, "ERRO_PREPARACAO_COMANDO_SQL", gErr, sSQL)
        
         Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 216245)
     
    End Select
     
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next
        
    Exit Function

End Function

Public Function SldDiaMeioPagtoCx_Le_AcumPeriodo_SQL(ByVal iFilialEmpresa As Integer, ByVal iCxDe As Integer, ByVal iCxAte As Integer, ByVal iTipoDe As Integer, ByVal iTipoAte As Integer, ByVal dtDataDe As Date, ByVal dtDataAte As Date, ByVal bTabCust As Boolean, sSQL As String) As Long

Dim lErro As Long
Dim sWhere As String
Dim sFiltroProd As String, sFiltro As String
Dim colVendedores As New Collection, objVendedor As ClassVendedor

On Error GoTo Erro_SldDiaMeioPagtoCx_Le_AcumPeriodo_SQL

    sSQL = "SELECT SUM(ValorFaturado) FROM SldDiaMeioPagtoCx" & IIf(bTabCust, "Cust ", " ")
    
    If iFilialEmpresa <> EMPRESA_TODA Then
        sWhere = sWhere & "AND FilialEmpresa = ? "
    End If
        
    If iCxDe <> 0 Then
        sWhere = sWhere & "AND Caixa >= ? "
    End If
    
    If iCxAte <> 0 Then
        sWhere = sWhere & "AND Caixa <= ? "
    End If
    
    If iTipoDe <> 0 Then
        sWhere = sWhere & "AND TipoMeioPagto >= ? "
    End If
    
    If iTipoAte <> 0 Then
        sWhere = sWhere & "AND TipoMeioPagto <= ? "
    End If
    
    If dtDataDe <> DATA_NULA Then
        sWhere = sWhere & "AND Data >= ? "
    End If
    
    If dtDataAte <> DATA_NULA Then
        sWhere = sWhere & "AND Data <= ? "
    End If
       
    If Len(Trim(sWhere)) > 0 Then sWhere = "WHERE " & Mid(sWhere, 5)
    sSQL = sSQL & sWhere
    
    SldDiaMeioPagtoCx_Le_AcumPeriodo_SQL = SUCESSO
     
    Exit Function
    
Erro_SldDiaMeioPagtoCx_Le_AcumPeriodo_SQL:

    SldDiaMeioPagtoCx_Le_AcumPeriodo_SQL = gErr
     
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 216246)
     
    End Select
     
    Exit Function

End Function

Private Function SldDiaMeioPagtoCx_Le_AcumPeriodo_Int(ByVal lComando As Long, ByVal sSQL As String, ByVal viFilialEmpresa As Variant, ByVal viCxDe As Variant, ByVal viCxAte As Variant, ByVal viTipoDe As Variant, ByVal viTipoAte As Variant, ByVal vdtDataDe As Variant, ByVal vdtDataAte As Variant, vdValor As Variant) As Long

Dim lErro As Long
Dim viStatus As Variant

On Error GoTo Erro_SldDiaMeioPagtoCx_Le_AcumPeriodo_Int

    lErro = Comando_PrepararInt(lComando, sSQL)
    If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM

    lErro = Comando_BindVarInt(lComando, vdValor)
    If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM

    If viFilialEmpresa <> EMPRESA_TODA Then
        lErro = Comando_BindVarInt(lComando, viFilialEmpresa)
        If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    End If
    
    If viCxDe <> 0 Then
        lErro = Comando_BindVarInt(lComando, viCxDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    End If
    
    If viCxAte <> 0 Then
        lErro = Comando_BindVarInt(lComando, viCxAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    End If
       
    If viTipoDe <> 0 Then
        lErro = Comando_BindVarInt(lComando, viCxDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    End If
    
    If viTipoAte <> 0 Then
        lErro = Comando_BindVarInt(lComando, viCxAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    End If
    
    If vdtDataDe <> DATA_NULA Then
        lErro = Comando_BindVarInt(lComando, vdtDataDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    End If
        
    If vdtDataAte <> DATA_NULA Then
        lErro = Comando_BindVarInt(lComando, vdtDataAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    End If
       
    lErro = Comando_ExecutarInt(lComando)
    If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    
    SldDiaMeioPagtoCx_Le_AcumPeriodo_Int = SUCESSO
     
    Exit Function
    
Erro_SldDiaMeioPagtoCx_Le_AcumPeriodo_Int:

    SldDiaMeioPagtoCx_Le_AcumPeriodo_Int = gErr
     
    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            Call Rotina_Erro(vbOKOnly, "ERRO_PREPARACAO_COMANDO_SQL", gErr, sSQL)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 216247)
     
    End Select
     
    Exit Function

End Function



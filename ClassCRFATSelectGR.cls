VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCRFATSelectGR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function CondPagto_Critica_Usu(objCondicaoPagto) As Long
'Verifica a Existencia da condicao de pagamento na tabela CotacaoGR

Dim lErro As Long
Dim lComando As Long
Dim iCondPagto As Integer

On Error GoTo Erro_CondPagto_Critica_Usu

    'Abre o Comando
    lComando = Comando_Abrir
    If lComando = 0 Then gError 97901
    
    'Define a Busca pela Condicao de Pagamento passada como paramentro
    lErro = Comando_Executar(lComando, "SELECT CondicaoPagto FROM CotacaoGR WHERE CondicaoPagto = ?", iCondPagto, objCondicaoPagto.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 97902
    
    'Faz a busca do Primeiro Registro a ser  encontrado
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97903
    
    'Se encontrou => erro
    If lErro = AD_SQL_SUCESSO Then gError 97904
    
    'Fecha o Comando
    Call Comando_Fechar(lComando)
    
    CondPagto_Critica_Usu = SUCESSO
    
    Exit Function
    
Erro_CondPagto_Critica_Usu:

    CondPagto_Critica_Usu = gErr
    
    Select Case gErr
    
        Case 97901
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 97902, 97903
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COTACAOGR", gErr)
            
        Case 97904
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAOPAGTO_UTILIZADA", gErr, objCondicaoPagto.iCodigo, "CondicaoPagto")
    
    End Select
    
    'Fecha o Comando
    Call Comando_Fechar(lComando)

End Function

Function Cliente_Critica_Usu(objCliente As ClassCliente) As Long
'Verifica a Existencia do Cliente na Tabela SolicitacaoServico e TabPrecoGR

Dim lErro As Long
Dim alComando(1) As Long
Dim lCliente As Long
Dim iIndice As Integer

On Error GoTo Erro_Cliente_Critica_Usu

    'Abre o Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 97907
    Next
    
    '************************************************************************ _
    BUSCA NA TABELA DE SolicitacaoServico
    
    'Define a Busca
    lErro = Comando_Executar(alComando(0), "SELECT Cliente FROM SolicitacaoServico WHERE Cliente = ?", lCliente, objCliente.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 97908
    
    'Faz a busca do Primeiro Registro a ser  encontrado
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97909
    
    'Se encontrou => erro
    If lErro = AD_SQL_SUCESSO Then gError 97910
    
    '************************************************************************ _
    BUSCA NA TABELA DE TabPrecoGR
    
    'Define a Busca
    lErro = Comando_Executar(alComando(1), "SELECT Cliente FROM TabPrecoGR WHERE Cliente = ?", lCliente, objCliente.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 97911
    
    'Faz a busca do Primeiro Registro a ser  encontrado
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97912
    
    'Se encontrou => erro
    If lErro = AD_SQL_SUCESSO Then gError 97913
    
    'Fecha os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next
    
    Cliente_Critica_Usu = SUCESSO
    
    Exit Function
    
Erro_Cliente_Critica_Usu:

    Cliente_Critica_Usu = gErr
    
    Select Case gErr
    
        Case 97907
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 97908, 97909
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SOLICITACAOSERVICO", gErr)
            
        Case 97911, 97912
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABPRECOGR", gErr)
            
        Case 97904
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_UTILIZADO", gErr, objCliente.lCodigo, "SOLICITACAOSERVICO")
            
        Case 97913
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_UTILIZADO", gErr, objCliente.lCodigo, "TABPRECOGR")
            
    End Select
    
    'Fecha os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next

End Function

Function NFiscalComprovante_Le(ByVal objNFiscal As ClassNFiscal) As Long
'Lê os Comprovantes da Nota Fiscal passada em objNFiscal

Dim iIndice As Integer
Dim lErro As Long
Dim lComando As Long
Dim tCompNF As typeComprovanteNF
Dim objComprovante As ClassCompServ

On Error GoTo Erro_NFiscalComprovante_Le

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 99091

    With tCompNF

        .sProduto = String(STRING_PRODUTO, 0)
        
        'Pesquisa os ítens vinculados à Nota Fiscal
        lErro = Comando_Executar(lComando, "SELECT Codigo, Produto, Quantidade, FretePeso, DataEmissao, ValorMercadoria, ValorContainer, AdValoren, Pedagio FROM CompServGR WHERE NumIntNota = ? ORDER BY Codigo", .lCodigo, .sProduto, .dQuantidade, .dPreco, .dtData, .dValorMercadoria, .dValorContainer, .dAdValoren, .dPedagio, objNFiscal.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 99092

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 99093
        
        If lErro = AD_SQL_SEM_DADOS Then gError 99094
        
        Do While lErro <> AD_SQL_SEM_DADOS
            
            Set objComprovante = New ClassCompServ
            
            objComprovante.lCodigo = .lCodigo
            objComprovante.sProduto = .sProduto
            objComprovante.dtDataEmissao = .dtData
            objComprovante.dQuantidade = .dQuantidade
            objComprovante.dValorMercadoria = .dValorMercadoria
            objComprovante.dValorContainer = .dValorContainer
            objComprovante.dAdValoren = .dAdValoren
            objComprovante.dFretePeso = .dPreco
            objComprovante.dPedagio = .dPedagio
            
            'Guarda na coleção de ítens de objNFiscal
            objNFiscal.colComprovServ.Add objComprovante

            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 99095

        Loop

    End With

    'Fecha os comandos
    Call Comando_Fechar(lComando)

    NFiscalComprovante_Le = SUCESSO

    Exit Function

Erro_NFiscalComprovante_Le:

    NFiscalComprovante_Le = gErr

    Select Case gErr

        Case 99091
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 99092, 99093, 99095
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMPROVANTENFISCAL", gErr)
        
        Case 99094 'Tratado na rotina chamadora
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMPROVANTENFISCAL_NAO_CADASTRADO", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function


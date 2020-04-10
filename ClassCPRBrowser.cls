VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCPRBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function TituloPag_TodosLista_BotaoEdita(objTituloPagar As ClassTituloPagar, lErro As Long) As Long

On Error GoTo Erro_TituloPag_TodosLista_BotaoEdita

    If lErro = SUCESSO Then
        
        If objTituloPagar.sSiglaDocumento = TIPODOC_FATURA_A_PAGAR Then
    
            Call Chama_Tela("FaturasPag", objTituloPagar)
    
        Else
    
            Call Chama_Tela("NFFatPag", objTituloPagar)
    
        End If
    
    Else
    
        Call Chama_Tela("NFFatPag")
        
    End If
    
    TituloPag_TodosLista_BotaoEdita = SUCESSO
        
    Exit Function

Erro_TituloPag_TodosLista_BotaoEdita:

    TituloPag_TodosLista_BotaoEdita = Err
    
    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147887)

    End Select

    Exit Function

End Function

Function FiliaisFornecedoresCompleto_BotaoEdita(objFilialFornecedor As ClassFilialFornecedor, lErro As Long) As Long

Dim objFornecedor As ClassFornecedor

On Error GoTo Erro_FiliaisFornecedoresCompleto_BotaoEdita

    If Not (objFilialFornecedor Is Nothing) Then

        If objFilialFornecedor.iCodFilial = FILIAL_MATRIZ Then
        
            Set objFornecedor = New ClassFornecedor
            
            'passa para objFornecedor o codigo do fornecedor correspondente em objFilialFornecedor
            objFornecedor.lCodigo = objFilialFornecedor.lCodFornecedor
            
            'le os dados do fornecedor
            lErro = CF("Fornecedor_Le", objFornecedor)
            If lErro <> SUCESSO And lErro <> 12729 Then gError 123519
            If lErro = 12729 Then gError 123520
            
            Call Chama_Tela("Fornecedores", objFornecedor)
            
        Else
        
            Call Chama_Tela("FiliaisFornecedores", objFilialFornecedor)
        
        End If
    
    Else
    
            Call Chama_Tela("Fornecedores")
    
    End If
    
    FiliaisFornecedoresCompleto_BotaoEdita = SUCESSO
    
    Exit Function
    
Erro_FiliaisFornecedoresCompleto_BotaoEdita:

    FiliaisFornecedoresCompleto_BotaoEdita = gErr

    Select Case gErr
    
    Case 123519
    
    Case 123520
    
    Case Else
       Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147888)

    End Select

    Exit Function

End Function


Function TituloPag_TodosTipoLista_BotaoEdita(objTituloPagar As ClassTituloPagar, lErro As Long) As Long

On Error GoTo Erro_TituloPag_TodosTipoLista_BotaoEdita

    If lErro = SUCESSO Then
        
        If objTituloPagar.sSiglaDocumento = TIPODOC_FATURA_A_PAGAR Then
    
            Call Chama_Tela("FaturasPag", objTituloPagar)
    
        Else
    
            Call Chama_Tela("NFFatPag", objTituloPagar)
    
        End If
    
    Else
    
        Call Chama_Tela("NFFatPag")
        
    End If
    
    
    TituloPag_TodosTipoLista_BotaoEdita = SUCESSO
        
    Exit Function

Erro_TituloPag_TodosTipoLista_BotaoEdita:

    TituloPag_TodosTipoLista_BotaoEdita = Err
    
    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147889)

    End Select

    Exit Function

End Function
Function TitulosPagarLista_BotaoEdita(objTituloPagar As ClassTituloPagar, lErro As Long) As Long

On Error GoTo Erro_TitulosPagarLista_BotaoEdita

    If lErro = SUCESSO Then
        
        If objTituloPagar.sSiglaDocumento = TIPODOC_FATURA_A_PAGAR Then
    
            Call Chama_Tela("FaturasPag", objTituloPagar)
    
        Else
        
            Call Chama_Tela("NFFatPag", objTituloPagar)
    
        End If
    
    Else
    
        Call Chama_Tela("NFFatPag")
        
    End If
    
    
    TitulosPagarLista_BotaoEdita = SUCESSO
        
    Exit Function

Erro_TitulosPagarLista_BotaoEdita:

    TitulosPagarLista_BotaoEdita = Err
    
    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147890)

    End Select

    Exit Function

End Function
Function TitulosPagLista_Consulta_BotaoEdita(objTituloPagar As ClassTituloPagar, lErro As Long) As Long

On Error GoTo Erro_TitulosPagLista_Consulta_BotaoEdita

    If lErro = SUCESSO Then
        
        If objTituloPagar.sSiglaDocumento = TIPODOC_FATURA_A_PAGAR Then
    
            Call Chama_Tela("FaturasPag", objTituloPagar)
    
        Else
        
            Call Chama_Tela("NFFatPag", objTituloPagar)
    
        End If
    
    Else
    
        Call Chama_Tela("NFFatPag")
        
    End If
    
    
    TitulosPagLista_Consulta_BotaoEdita = SUCESSO
        
    Exit Function

Erro_TitulosPagLista_Consulta_BotaoEdita:

    TitulosPagLista_Consulta_BotaoEdita = Err
    
    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147891)

    End Select

    Exit Function

End Function

'Incluída por Luiz Nogueira em 18/05/04
Public Function ClientesLista_BotaoEdita(ByVal objCliente As ClassCliente, lErro As Long) As Long

Dim objFilialCliente As ClassFilialCliente

On Error GoTo Erro_ClientesLista_BotaoEdita

    '##########################################
    'Alterado por Wagner 11/01/2006
    If Not (objCliente Is Nothing) Then

        'Verifica se o código da filial editada é diferente de 1 (código da filial principal)
        If objCliente.iFilialEmpresaLoja <> 1 Then
            
            Set objFilialCliente = New ClassFilialCliente
            
            objFilialCliente.lCodCliente = objCliente.lCodigo
            objFilialCliente.iCodFilial = objCliente.iFilialEmpresaLoja
            
            Call Chama_Tela("FiliaisClientes", objFilialCliente)
        
        Else
        
            Call Chama_Tela("Clientes", objCliente)
        
        End If
        
    Else
    
        Call Chama_Tela("Clientes", objCliente)
    
    End If
    '##########################################
    
    ClientesLista_BotaoEdita = SUCESSO

    Exit Function
    
Erro_ClientesLista_BotaoEdita:

    ClientesLista_BotaoEdita = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147892)

    End Select

End Function

Function MovCtaCorrenteLista_BotaoEdita(objMovCtaCorrente As ClassMovContaCorrente, lErro As Long) As Long

Dim objAplicacao As New ClassAplicacao
Dim objResgate As New ClassResgate
Dim objAntecipPag As New ClassAntecipPag
Dim objAntecipReceb As New ClassAntecipReceb
Dim colSelecao As New Collection
Dim colTitulos As New Collection

On Error GoTo Erro_MovCtaCorrenteLista_BotaoEdita

    Select Case objMovCtaCorrente.iTipo
        
        Case MOVCCI_SAQUE
            Call Chama_Tela("Saque", objMovCtaCorrente)
    
        Case MOVCCI_DEPOSITO
            Call Chama_Tela("Deposito", objMovCtaCorrente)
    
        Case MOVCCI_APLICACAO
        
            objAplicacao.lNumMovto = objMovCtaCorrente.lNumMovto
        
            lErro = CF("Aplicacao_Le_NumMovto", objAplicacao)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 189337
        
            Call Chama_Tela("Aplicacao", objAplicacao)

        Case MOVCCI_RESGATE

            objResgate.lNumMovto = objMovCtaCorrente.lNumMovto
        
            lErro = CF("Resgate_Le_NumMovto", objResgate)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 189343

            Call Chama_Tela("Resgate", objResgate)
    
        Case MOVCCI_SAIDA_TRANSFERENCIA, MOVCCI_ENTRADA_TRANSFERENCIA
            Call Chama_Tela("Transferencia", objMovCtaCorrente)
    
        Case MOVCCI_PAGTO_ANTECIPADO
        
            objAntecipPag.lNumMovto = objMovCtaCorrente.lNumMovto
        
            lErro = CF("AntecipPag_Le_NumMovto", objAntecipPag)
            If lErro <> SUCESSO And lErro <> 42845 Then gError 189345
        
            Call Chama_Tela("AntecipPag", objAntecipPag)

        Case MOVCCI_RECEB_ANTECIPADO
            
            objAntecipReceb.lNumMovto = objMovCtaCorrente.lNumMovto
        
            lErro = CF("AntecipRec_Le_NumMovto", objAntecipReceb)
            If lErro <> SUCESSO And lErro <> 46252 Then gError 189346
            
            Call Chama_Tela("AntecipReceb", objAntecipReceb)
    
        Case MOVCCI_RECEBIMENTO_TITULO, MOVCCI_CREDITO_RETORNO_COBRANCA, MOVCCI_CRED_RET_COBRANCA
        
            lErro = CF("BaixaRec_Le_TitulosBaixados", objMovCtaCorrente.lNumMovto, colTitulos)
            If lErro <> SUCESSO Then gError 189361
            
            If colTitulos.Count <> 1 Then
        
                colSelecao.Add objMovCtaCorrente.lNumMovto
            
                Call Chama_Tela("BaixasRecLista", colSelecao, Nothing, Nothing, "NumMovCta = ?")
    
            Else
            
                Call Chama_Tela("TituloReceber_Consulta", colTitulos.Item(1))
            
            End If
    
        Case MOVCCI_PAGTO_TITULO_POR_CHEQUE, MOVCCI_PAGTO_TITULO_POR_BORDERO, MOVCCI_PAGTO_TITULO_POR_DINHEIRO
            
            lErro = CF("BaixaPag_Le_TitulosBaixados", objMovCtaCorrente.lNumMovto, colTitulos)
            If lErro <> SUCESSO Then gError 189362
            
            If colTitulos.Count <> 1 Then
            
                colSelecao.Add objMovCtaCorrente.lNumMovto
                
                Call Chama_Tela("BaixasPagLista", colSelecao, Nothing, Nothing, "NumMovCta = ?")
                
            Else
            
                Call Chama_Tela("TituloPagar_Consulta", colTitulos.Item(1))
            
            End If
    
        Case Else
            Call Rotina_Aviso(vbOKOnly, "AVISO_TIPO_SEM_TELA_DE_CONSULTA")
        
    End Select
    
    MovCtaCorrenteLista_BotaoEdita = SUCESSO
        
    Exit Function

Erro_MovCtaCorrenteLista_BotaoEdita:

    MovCtaCorrenteLista_BotaoEdita = gErr
    
    Select Case gErr
    
        Case 189337, 189343, 189345, 189346, 189361, 189362

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 189339)

    End Select

    Exit Function

End Function

Function FluxoContratoItemNFPagLista_BotaoEdita(objFluxoContratoItemNFPag As ClassFluxoContratoItemNFPag, lErro As Long) As Long

Dim objNFiscal As New ClassNFiscal
Dim objContrato As New ClassContrato
Dim objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Erro_FluxoContratoItemNFPagLista_BotaoEdita

    If lErro = SUCESSO Then
        
        If objFluxoContratoItemNFPag.lNumNotaFiscal <> 0 Then
    
            'lê a Nota Fiscal
            objNFiscal.lNumIntDoc = objFluxoContratoItemNFPag.lNumIntNF
            
            lErro = CF("NFiscal_Le", objNFiscal)
            If lErro <> SUCESSO And lErro <> 31442 Then gError 188284
            
            If lErro = 31442 Then gError 188285
            
            objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
        
            'lê o Tipo da Nota Fiscal
            lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
            If lErro <> SUCESSO And lErro <> 31415 Then gError 188286
            
            If lErro = 31415 Then gError 188287
        
            Call Chama_Tela(objTipoDocInfo.sNomeTelaNFiscal, objNFiscal)
    
        Else
    
            objContrato.lNumIntDoc = objFluxoContratoItemNFPag.lNumIntContrato
        
            Call Chama_Tela("ContratoPagar", objContrato)
    
        End If
    
    Else
    
        Call Chama_Tela("ContratoPagar")
        
    End If
    
    
    FluxoContratoItemNFPagLista_BotaoEdita = SUCESSO
        
    Exit Function

Erro_FluxoContratoItemNFPagLista_BotaoEdita:

    FluxoContratoItemNFPagLista_BotaoEdita = gErr
    
    Select Case gErr

        Case 188284, 188286
        
        Case 188287
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objTipoDocInfo.iCodigo)
            
        Case 188285
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA", gErr, objNFiscal.lNumIntDoc)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 188288)

    End Select

    Exit Function

End Function

Function FluxoContratoItemNFRecLista_BotaoEdita(objFluxoContratoItemNFRec As ClassFluxoContratoItemNFRec, lErro As Long) As Long

Dim objNFiscal As New ClassNFiscal
Dim objContrato As New ClassContrato
Dim objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Erro_FluxoContratoItemNFRecLista_BotaoEdita

    If lErro = SUCESSO Then
        
        If objFluxoContratoItemNFRec.lNumNotaFiscal <> 0 Then
    
            'lê a Nota Fiscal
            objNFiscal.lNumIntDoc = objFluxoContratoItemNFRec.lNumIntNF
            
            lErro = CF("NFiscal_Le", objNFiscal)
            If lErro <> SUCESSO And lErro <> 31442 Then gError 188289
            
            If lErro = 31442 Then gError 188290
            
            objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
        
            'lê o Tipo da Nota Fiscal
            lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
            If lErro <> SUCESSO And lErro <> 31415 Then gError 188291
            
            If lErro = 31415 Then gError 188292
        
            Call Chama_Tela(objTipoDocInfo.sNomeTelaNFiscal, objNFiscal)
    
        Else
    
            objContrato.lNumIntDoc = objFluxoContratoItemNFRec.lNumIntContrato
        
            Call Chama_Tela("ContratoCadastro", objContrato)
    
        End If
    
    Else
    
        Call Chama_Tela("ContratoCadastro")
        
    End If
    
    
    FluxoContratoItemNFRecLista_BotaoEdita = SUCESSO
        
    Exit Function

Erro_FluxoContratoItemNFRecLista_BotaoEdita:

    FluxoContratoItemNFRecLista_BotaoEdita = gErr
    
    Select Case gErr

        Case 188289, 188291
        
        Case 188290
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA", gErr, objNFiscal.lNumIntDoc)

        Case 188292
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objTipoDocInfo.iCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 188293)

    End Select

    Exit Function

End Function

Function RelFlCxCtbLista_BotaoEdita(objRelFlCxCtb As ClassRelFlCxCtb, lErro As Long) As Long

Dim objAplicacao As New ClassAplicacao
Dim objResgate As New ClassResgate
Dim objAntecipPag As New ClassAntecipPag
Dim objAntecipReceb As New ClassAntecipReceb
Dim colSelecao As New Collection
Dim colTitulos As New Collection
Dim objMovCtaCorrente As New ClassMovContaCorrente
Dim objTituloPagar As New ClassTituloPagar
Dim objTituloRec As New ClassTituloReceber
Dim objParcelaPagar As New ClassParcelaPagar
Dim objParcelaReceber As New ClassParcelaReceber

On Error GoTo Erro_RelFlCxCtbLista_BotaoEdita

    Select Case objRelFlCxCtb.sTipo
        
        Case "SAQ"
            
            objMovCtaCorrente.lNumMovto = objRelFlCxCtb.lNumMovto
        
            Call Chama_Tela("Saque", objMovCtaCorrente)
    
        Case "DEP"
        
            objMovCtaCorrente.lNumMovto = objRelFlCxCtb.lNumMovto
        
            Call Chama_Tela("Deposito", objMovCtaCorrente)
    
    
        Case "TRF"
            
            objMovCtaCorrente.lNumMovto = objRelFlCxCtb.lNumMovto
            
            'Le o movimento passado como parametro
            lErro = CF("MovContaCorrente_Le", objMovCtaCorrente)
            If lErro <> SUCESSO And lErro <> 11893 Then gError 197862
            
            Call Chama_Tela("Transferencia", objMovCtaCorrente)
    
        Case "ANP"
        
            If objRelFlCxCtb.sSiglaDoc = "" Then
            
                objAntecipPag.lNumMovto = objRelFlCxCtb.lNumMovto
            
                lErro = CF("AntecipPag_Le_NumMovto", objAntecipPag)
                If lErro <> SUCESSO And lErro <> 42845 Then gError 197863
            
                Call Chama_Tela("AntecipPag", objAntecipPag)
            Else
            
                objTituloPagar.lNumIntDoc = objRelFlCxCtb.lNumIntDoc
            
                Call Chama_Tela("TituloPagar_Consulta", objTituloPagar)
            
            End If

        Case "ANR"
            
            If objRelFlCxCtb.sSiglaDoc = "" Then
            
                objAntecipReceb.lNumMovto = objRelFlCxCtb.lNumMovto
            
                lErro = CF("AntecipRec_Le_NumMovto", objAntecipReceb)
                If lErro <> SUCESSO And lErro <> 46252 Then gError 197864
            
                Call Chama_Tela("AntecipReceb", objAntecipReceb)
    
            Else
            
                objTituloRec.lNumIntDoc = objRelFlCxCtb.lNumIntDoc
            
                Call Chama_Tela("TituloReceber_Consulta", objTituloRec)
            
            End If
            
    
        Case "TPG"
        
                objTituloPagar.lNumIntDoc = objRelFlCxCtb.lNumIntDoc
            
                Call Chama_Tela("TituloPagar_Consulta", objTituloPagar)
        
    
        Case "TRC"
            
                objTituloRec.lNumIntDoc = objRelFlCxCtb.lNumIntDoc
            
                Call Chama_Tela("TituloReceber_Consulta", objTituloRec)
    
    
        Case "PPG"
    
            objParcelaPagar.lNumIntDoc = objRelFlCxCtb.lNumIntDoc
    
            'Le o NumInterno do Titulo para passar no objTituloPag
            lErro = CF("ParcelaPagar_Le", objParcelaPagar)
            If lErro <> SUCESSO And lErro <> 60479 Then gError 197867
            If lErro <> SUCESSO Then
            
                'se não encontrou
                lErro = CF("ParcelaPagarBaixada_Le", objParcelaPagar)
                If lErro <> SUCESSO Then gError 197868
            
            End If
            
            objTituloPagar.lNumIntDoc = objParcelaPagar.lNumIntTitulo
            
            Call Chama_Tela("TituloPagar_Consulta", objTituloPagar)
    
    
        Case "PRC"
    
            objParcelaReceber.lNumIntDoc = objRelFlCxCtb.lNumIntDoc
            
            'Le o NumInterno do Titulo para passar no objParcelaReceber
            lErro = CF("ParcelaReceber_Le", objParcelaReceber)
            If lErro <> SUCESSO And lErro <> 19147 Then gError 197869
            If lErro <> SUCESSO Then
    
                'Se não encontrar
                lErro = CF("ParcelaReceber_Baixada_Le", objParcelaReceber)
                If lErro <> SUCESSO Then gError 197870
    
            End If
    
            objTituloRec.lNumIntDoc = objParcelaReceber.lNumIntTitulo
            
            Call Chama_Tela("TituloReceber_Consulta", objTituloRec)
    
    
        Case Else
            Call Rotina_Aviso(vbOKOnly, "AVISO_TIPO_SEM_TELA_DE_CONSULTA")
        
    End Select
    
    RelFlCxCtbLista_BotaoEdita = SUCESSO
        
    Exit Function

Erro_RelFlCxCtbLista_BotaoEdita:

    RelFlCxCtbLista_BotaoEdita = gErr
    
    Select Case gErr
    
        Case 197862, 197863, 197864, 197867 To 197870

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 197865)

    End Select

    Exit Function

End Function


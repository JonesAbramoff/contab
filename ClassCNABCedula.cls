VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCNABCedula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim glNumArquivoRemessa As Long
Dim gColFiliaisEmpresa As Collection
Dim gobjCobrancaEletronica As ClassCobrancaEletronica
Public gobjGeracaoArqCNABPag As ClassGeracaoArqCNABPag
Dim gobjFilialEmpresa As AdmFiliais

Function Criar_Arquivo_Remessa_Cedula(objCobrancaEletronica As ClassCobrancaEletronica) As Long
'Cria uma arquivo de remessa com base nos dados do BD

Dim iCancela As Integer
Dim lErro As Long
Dim sRegHeader As String
Dim lTransacao As Long
Dim alComando(0 To 5) As Long
Dim sNomeArquivo As String
Dim iIndice As Integer
Dim sRegDetalhe As String
Dim sRegTrailer As String, lSeqNomeArq As Long
Dim objBorderoCobranca As New ClassBorderoCobranca
Dim objCarteiraCobrador As New ClassCarteiraCobrador
Dim objOcorrRemParc As ClassOcorrRemParcRec

On Error GoTo Erro_Criar_Arquivo_Remessa_Cedula
    
    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 51754
    Next
        
    'ABre a transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 51755
    
    Set gobjCobrancaEletronica = objCobrancaEletronica
    Set gColFiliaisEmpresa = New Collection
    Set gobjFilialEmpresa = New AdmFiliais
    
    'Gera o nome do arquivo de acordo com o sequencial do Cobrador
    lErro = CF("CobrancaEletronica_Gera_NomeArquivo", objCobrancaEletronica.objCobrador, objCobrancaEletronica.sDiretorio, sNomeArquivo, lSeqNomeArq)
    If lErro <> SUCESSO Then gError 51756
    
    gobjCobrancaEletronica.lSeqNomeArq = lSeqNomeArq
    
    gobjFilialEmpresa.iCodFilial = giFilialEmpresa
    
    lErro = CF("FilialEmpresa_Le", gobjFilialEmpresa, True)
    If lErro <> SUCESSO And lErro <> 27378 Then gError 51814
    If lErro <> SUCESSO Then gError 90073
    
    'Cria o arquivo de remessa
    Open sNomeArquivo For Output As #1
        
    lErro = Montar_RegistroHeader_Cedula(sRegHeader)
    If lErro <> SUCESSO Then gError 51760
        
    'Inclui o registro de Header
    Print #1, DesacentuaTexto(sRegHeader)

    For Each objBorderoCobranca In objCobrancaEletronica.colBorderos
        'Lê o bordero de cobrança
        lErro = CF("BorderoCobranca_LeMod", alComando(2), objBorderoCobranca)
        If lErro <> SUCESSO Then gError 51757
            
        objCarteiraCobrador.iCobrador = objCobrancaEletronica.iCobrador
        objCarteiraCobrador.iCodCarteiraCobranca = objBorderoCobranca.iCodCarteiraCobranca
        
        'Lê a carteira do cobrador
        lErro = CF("CarteiraCobrador_LeMod", objCarteiraCobrador, alComando(4))
        If lErro <> SUCESSO Then gError 51759
               
        'Busca no BD as Ocorrências do Bordero em questão
        lErro = CF("CobrancaEletronica_Obter_OcorrRemParcRec", alComando(3), objBorderoCobranca, gobjCobrancaEletronica)
        If lErro <> SUCESSO Then gError 51662
        
        'Para cada ocorrência
        For Each objOcorrRemParc In objCobrancaEletronica.colOcorrRemParcRec
    
            sRegDetalhe = ""
            'Monta um registro de detalhe do arquivo de remessa
            lErro = Montar_RegistroDetalhe_Cedula(sRegDetalhe, objOcorrRemParc, objCarteiraCobrador)
            If lErro <> SUCESSO And lErro <> 32266 Then gError 51761
            
            If lErro = SUCESSO Then
            
                'Inclui o registro com os detalhes do Título(Parcela) a receber
                Print #1, DesacentuaTexto(sRegDetalhe)
                
            End If
            
        Next
        
        objBorderoCobranca.iProcessado = BORDERO_PROCESSADO
        objBorderoCobranca.sNomeArquivo = sNomeArquivo
        
        'Atualiza o Status do Bordero de Cobrança
        lErro = CF("BorderoCobranca_Atualiza_Processo", alComando(2), alComando(3), objBorderoCobranca)
        If lErro <> SUCESSO Then gError 51758
    
        'Atualiza o NossoNumero em carteiraCobrador
        lErro = CF("CarteiraCobrador_AtualizaNossoNumero", objCarteiraCobrador, alComando(4), alComando(5))
        If lErro <> SUCESSO Then gError 51707
    
        iCancela = 0
        
        Call gobjCobrancaEletronica.objTelaAtualizacao.Mostra_Evolucao(iCancela, 1)
        
        If iCancela = CANCELA_BATCH Then gError 79705
        
    Next
    
    
    'MOnta o registro de trailer
    lErro = Montar_RegistroTrailer_Cedula(sRegTrailer, objOcorrRemParc)
    If lErro <> SUCESSO Then gError 51762

    'Coloca o registro de trailer no arquivo
    Print #1, DesacentuaTexto(sRegTrailer)
    
    Close #1
    
    lErro = Transacao_Commit()
    If lErro <> SUCESSO Then gError 62010
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Criar_Arquivo_Remessa_Cedula = SUCESSO
    
    Exit Function

Erro_Criar_Arquivo_Remessa_Cedula:

    Criar_Arquivo_Remessa_Cedula = gErr
    
    Select Case gErr
    
        Case 51754
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 51755
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 51756, 51757, 51758, 51759, 51760, 51761, 51762, 79705
        
        Case 62010
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147220)
    
    End Select
    
    Close #1
    
    Call Transacao_Rollback

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Private Function Montar_RegistroHeader_Cedula(sRegHeader As String) As Long
'Monta o registro

Dim tRegHeaderCedula As typeRemHeaderCedula
Dim lErro As Long

On Error GoTo Erro_Montar_RegistroHeader_Cedula

    gobjCobrancaEletronica.lSequencialRegistro = 1

    'Carrega tRegHeaderCedula com os dados do Header do aquivo
    lErro = Obter_Dados_Header(tRegHeaderCedula)
    If lErro <> SUCESSO Then Error 51763
        
    With tRegHeaderCedula
    
        'Monta o registro de Header do Cedula
        sRegHeader = .sIdentificacaoReg & .sIdentificacaoArq & .sLiteralRemessa & .sCodTipoServico & .sLiteralServico
        sRegHeader = sRegHeader & .sCodigoEmpresa & .sNomeEmpresa & .sNumeroBanco & .sNomeBanco & .sDataEmissaoArq
        sRegHeader = sRegHeader & .sComplemento1 & .sIdentificacaoSistema & .sSequencialArq & .sComplemento2 & .sSequencialRegistro
        
    End With
    
    Montar_RegistroHeader_Cedula = SUCESSO

    Exit Function

Erro_Montar_RegistroHeader_Cedula:

    Montar_RegistroHeader_Cedula = Err

    Select Case Err

        Case 51763

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147221)

    End Select

    Exit Function

End Function

Private Function Obter_Dados_Header(tRegHeaderCedula As typeRemHeaderCedula) As Long

Dim lErro As Long
Dim objCCI As New ClassContasCorrentesInternas
Dim sCodEmpresaBanco As String

On Error GoTo Erro_Obter_Dados_Header
        
    'Lê a conta corrente associada ao cobrador
    lErro = CF("ContaCorrenteInt_Le", gobjCobrancaEletronica.objCobrador.iCodCCI, objCCI)
    If lErro <> SUCESSO And lErro <> 11807 Then Error 51764
    If lErro <> SUCESSO Then Error 51765
    
    If objCCI.iCodBanco <> gobjCobrancaEletronica.objCobrador.iCodBanco Then Error 62061
    If Len(Trim(objCCI.sAgencia)) = 0 Or Len(Trim(objCCI.sNumConta)) = 0 Then Error 62062
        
    Set gobjCobrancaEletronica.objCCI = objCCI
    
    'Carrega os campos do cabeçalho
    tRegHeaderCedula.sIdentificacaoReg = "0"
    tRegHeaderCedula.sIdentificacaoArq = "1"
    tRegHeaderCedula.sLiteralRemessa = FormataCpoTexto("REMESSA", 7)
    tRegHeaderCedula.sCodTipoServico = "01"
    tRegHeaderCedula.sLiteralServico = FormataCpoTexto("COBRANCA", 15)
    tRegHeaderCedula.sCodigoEmpresa = FormataCpoNum(Obter_Info_Cobrador(gobjCobrancaEletronica.objCobradorCNABInfo.colInformacoes, 1), 20)
    tRegHeaderCedula.sNomeEmpresa = FormataCpoTexto(gsNomeEmpresa, 30)
    tRegHeaderCedula.sNumeroBanco = "266"
    tRegHeaderCedula.sNomeBanco = FormataCpoTexto("CEDULA", 15)
    tRegHeaderCedula.sDataEmissaoArq = Format(gdtDataHoje, "DDMMYY")
    tRegHeaderCedula.sComplemento1 = FormataCpoTexto("", 8)
    tRegHeaderCedula.sIdentificacaoSistema = FormataCpoTexto("MX", 2)
    tRegHeaderCedula.sSequencialArq = FormataCpoNum(gobjCobrancaEletronica.objCobrador.lCNABProxSeqArqCobr, 7)
    tRegHeaderCedula.sComplemento2 = FormataCpoTexto("", 277)
    tRegHeaderCedula.sSequencialRegistro = FormataCpoNum(gobjCobrancaEletronica.lSequencialRegistro, 6)
    
    'Incrementa o sequencial de registro
    gobjCobrancaEletronica.lSequencialRegistro = gobjCobrancaEletronica.lSequencialRegistro + 1

    Obter_Dados_Header = SUCESSO

    Exit Function
    
Erro_Obter_Dados_Header:

    Obter_Dados_Header = Err
    
    Select Case Err
    
        Case 62061
            lErro = Rotina_Erro(vbOKOnly, "ERRO_BANCO_CCI_DIFERENTE_COBRADOR", Err, objCCI.iCodBanco, gobjCobrancaEletronica.objCobrador.iCodBanco)
    
        Case 62062
            lErro = Rotina_Erro(vbOKOnly, "ERRO_AGENCIA_CONTA_COBRADOR_NAO_PREENCHIDAS", Err)
    
        Case 51764
        
        Case 51765
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_COBRADOR_NAO_ENCONTRADA", Err, gobjCobrancaEletronica.iCobrador)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147222)
            
    End Select
    
    Exit Function

End Function

Private Function Montar_RegistroDetalhe_Cedula(sRegDetalhe As String, objOcorrRemParcRec As ClassOcorrRemParcRec, objCarteiraCobrador As ClassCarteiraCobrador) As Long

'Monta o registro

Dim tRegDetalheCedula As typeRemDetalheCedula
Dim lErro As Long

On Error GoTo Erro_Montar_RegistroDetalhe_Cedula
    
    'Obtem os dados do registro de detalhe
    lErro = Obter_Dados_Detalhe(tRegDetalheCedula, objOcorrRemParcRec, objCarteiraCobrador)
    If lErro <> SUCESSO And lErro <> 32265 Then Error 51766
    If lErro <> SUCESSO Then Error 32266
    
    With tRegDetalheCedula
        'Forma o registro de detalhe
        sRegDetalhe = .sIdentificacaoReg & .sAgenciaDebito & .sDigitoAgDebito & .sRazaoContaCorrente & .sContaCorrente
        sRegDetalhe = sRegDetalhe & .sDigitoContaCorrente & .sIdentifEmpresa & .sNumControle & .sCodigoBancoDebitado & .sZeros
        sRegDetalhe = sRegDetalhe & .sNossoNumero & .sDesconto & .sCondEmissaoPapel & .sIdentEmitePapel & .sIdentOpBanco
        sRegDetalhe = sRegDetalhe & .sIndicadorRateio & .sEnderecamento & .sBranco & .sIdentificacaoOcorr & .sNumDocto
        sRegDetalhe = sRegDetalhe & .sDataVencto & .sValorTitulo & .sBancoCobranca & .sAgenciaDepositaria & .sEspecie
        sRegDetalhe = sRegDetalhe & .sAceite & .sDataEmissao & .sInstrucao1 & .sInstrucao2 & .sValorJuros & .sDataDesconto
        sRegDetalhe = sRegDetalhe & .sValorDesconto & .sValorIOF & .sAbatimento & .sIdentfInscSacado & .sNumInscricaoSacado
        sRegDetalhe = sRegDetalhe & .sNomeSacado & .sEnderecoSacado & .sMensagem1 & .sCEP & " " & FormataCpoNum(gobjFilialEmpresa.sCgc, 14) & FormataCpoTexto("", 45) & .sNumSequencialRegistro
    End With
    
    Montar_RegistroDetalhe_Cedula = SUCESSO

    Exit Function

Erro_Montar_RegistroDetalhe_Cedula:

    Montar_RegistroDetalhe_Cedula = Err

    Select Case Err
    
        Case 51766, 32266

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147223)

    End Select

    Exit Function

End Function

Private Function Obter_Dados_Detalhe(tRegDetalheCedula As typeRemDetalheCedula, objOcorrRemParcRem As ClassOcorrRemParcRec, objCarteiraCobrador As ClassCarteiraCobrador) As Long
'Obtem os dados que são utilizados para preencher o registro de detalhe

Dim lErro As Long
Dim objParcelaRec As New ClassParcelaReceber
Dim objEndereco As New ClassEndereco
Dim sNossoNumero As String
Dim dValorDesconto As Double

On Error GoTo Erro_Obter_Dados_Detalhe

    objParcelaRec.lNumIntDoc = objOcorrRemParcRem.lNumIntParc

    'Busca os dados da parcela a receber no BD
    lErro = CF("CobrancaEletronica_Obter_DadosParcelaRec", objParcelaRec, gobjCobrancaEletronica)
    If lErro <> SUCESSO And lErro <> 51674 Then Error 51767
    If lErro <> SUCESSO Then Error 32265
    
    lErro = CF("CobrancaEletronica_Obter_Endereco", gobjCobrancaEletronica, objEndereco)
    If lErro <> SUCESSO Then Error 51769
        
    sNossoNumero = ""
    If objCarteiraCobrador.iImprimeBoleta = EMPRESA_IMPRIME_BOLETA Then
        If Len(Trim(objCarteiraCobrador.sFaixaNossoNumeroProx)) = 0 Then objCarteiraCobrador.sFaixaNossoNumeroProx = FormataCpoNum(1, 11)
        lErro = Calcula_NossoNumero_Cedula(objCarteiraCobrador, sNossoNumero)
        If lErro <> SUCESSO Then Error 51770
    End If
    
    With tRegDetalheCedula
        .sIdentificacaoReg = "1"
        .sAgenciaDebito = FormataCpoNum("", 5)
        .sDigitoAgDebito = FormataCpoTexto("", 1)
        .sRazaoContaCorrente = FormataCpoNum("", 5)
        .sContaCorrente = FormataCpoNum("", 7)
        .sDigitoContaCorrente = FormataCpoTexto("", 1)
        .sIdentifEmpresa = "0" & FormataCpoNum(objCarteiraCobrador.iNumCarteiraNoBanco, 3) & FormataCpoNum(gobjCobrancaEletronica.objCCI.sAgencia, 5) & FormataCpoNum(gobjCobrancaEletronica.objCCI.sNumConta, 7) & FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVNumConta, 1)
        .sNumControle = FormataCpoTexto(objOcorrRemParcRem.lNumIntDoc, 25)
        .sCodigoBancoDebitado = FormataCpoNum("", 3)
        .sZeros = FormataCpoNum("", 5)
        .sNossoNumero = FormataCpoNum(sNossoNumero, 12)
        .sDesconto = FormataCpoNum("", 10)
        If objCarteiraCobrador.iImprimeBoleta = BANCO_IMPRIME_BOLETA Then
            .sCondEmissaoPapel = "1"
        Else
            .sCondEmissaoPapel = "2"
        End If
        .sIdentEmitePapel = FormataCpoTexto("", 1)
        .sIdentOpBanco = FormataCpoTexto("", 10)
        .sIndicadorRateio = FormataCpoTexto("", 1)
        .sEnderecamento = "2"
        .sBranco = FormataCpoTexto("", 2)
        .sIdentificacaoOcorr = FormataCpoNum(objOcorrRemParcRem.iCodOcorrencia, 2)
        .sNumDocto = FormataCpoTexto(gobjCobrancaEletronica.lNumTitulo & "/" & objParcelaRec.iNumParcela, "10")
        .sDataVencto = Format(objParcelaRec.dtDataVencimento, "DDMMYY")
        .sValorTitulo = FormataCpoValor(objParcelaRec.dSaldo, 13)
        .sBancoCobranca = FormataCpoNum("", 3)
        .sAgenciaDepositaria = FormataCpoNum("", 5)
        .sEspecie = FormataCpoNum(1, 2) 'Duplicata
        If objParcelaRec.iAceite = 1 Then
            .sAceite = "A"
        Else
            .sAceite = "N"
        End If
        .sDataEmissao = Format(gobjCobrancaEletronica.dtDataEmissaoTitulo, "DDMMYY")
        .sInstrucao1 = FormataCpoNum(objOcorrRemParcRem.iInstrucao1, 2)
        .sInstrucao2 = FormataCpoNum(objOcorrRemParcRem.iInstrucao2, 2)
        .sValorJuros = FormataCpoValor(objParcelaRec.dSaldo * objOcorrRemParcRem.dJuros, 13)
        If objParcelaRec.dtDesconto1Ate <> DATA_NULA Then
            .sDataDesconto = Format(objParcelaRec.dtDesconto1Ate, "DDMMYY")
        Else
            .sDataDesconto = FormataCpoNum("", 6)
        End If
        
        'Calcula o valor do desconto quando for do tipo percentual
        dValorDesconto = 0
        If objParcelaRec.iDesconto1Codigo = PERC_ANT_DIA Or objParcelaRec.iDesconto1Codigo = PERC_ANT_DIA_UTIL Or objParcelaRec.iDesconto1Codigo = Percentual Then
            dValorDesconto = StrParaDbl(Format(objParcelaRec.dDesconto1Valor * objParcelaRec.dSaldo, "Fixed"))
        Else
            If objParcelaRec.iDesconto1Codigo = VALOR_FIXO Then
                dValorDesconto = objParcelaRec.dDesconto1Valor
            End If
        End If
        
        .sValorDesconto = FormataCpoValor(dValorDesconto, 13)
        .sValorIOF = FormataCpoValor(0, 13)
        .sAbatimento = FormataCpoValor(0, 13)
        If Len(gobjCobrancaEletronica.sCGCSacado) = 14 Then
            .sIdentfInscSacado = "02"
        Else
            .sIdentfInscSacado = "01"
        End If
        .sNumInscricaoSacado = FormataCpoNum(gobjCobrancaEletronica.sCGCSacado, 14)
        .sNomeSacado = FormataCpoTexto(gobjCobrancaEletronica.sNomeSacado, 40)
        .sEnderecoSacado = FormataCpoTexto(Trim(objEndereco.sEndereco), 40)
        .sMensagem1 = FormataCpoTexto(tRegDetalheCedula.sMensagem1, 12)
        .sCEP = Formata_Cep(objEndereco.sCEP, 8)
        .sSacadorAvalista = FormataCpoTexto("", 60)
        .sNumSequencialRegistro = FormataCpoNum(gobjCobrancaEletronica.lSequencialRegistro, 6)
    
    End With
    
    'Incrementa o sequencial de registro
    gobjCobrancaEletronica.lSequencialRegistro = gobjCobrancaEletronica.lSequencialRegistro + 1
    
    Obter_Dados_Detalhe = SUCESSO

    Exit Function
    
Erro_Obter_Dados_Detalhe:

    Obter_Dados_Detalhe = Err
    
    Select Case Err

        Case 51767, 51769, 51770, 32265
       
        Case 51768
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ENDERECO_FILIALCLIENTE_NAO_INFORMADO", Err, gobjCobrancaEletronica.lCliente, gobjCobrancaEletronica.iFilialCli)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147224)

    End Select

    Exit Function

End Function

Private Function Montar_RegistroTrailer_Cedula(sRegTrailer As String, objCobrador As ClassCobrador) As Long
'Monta o registro

Dim tRegTrailerCedula As typeRemTrailerCedula
Dim lErro As Long

On Error GoTo Erro_Montar_RegistroTrailer_Cedula

    'Preenche os dados do regsitro de trailer
    tRegTrailerCedula.sIdentificacaoReg = "9"
    tRegTrailerCedula.sComplemento1 = FormataCpoTexto("", 393)
    tRegTrailerCedula.sSequencialRegistro = FormataCpoNum(gobjCobrancaEletronica.lSequencialRegistro, 6)
    
    'Monta o Registro de trailer
    sRegTrailer = tRegTrailerCedula.sIdentificacaoReg & tRegTrailerCedula.sComplemento1 & tRegTrailerCedula.sSequencialRegistro

    Montar_RegistroTrailer_Cedula = SUCESSO

    Exit Function

Erro_Montar_RegistroTrailer_Cedula:

    Montar_RegistroTrailer_Cedula = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147225)

    End Select

    Exit Function

End Function

Function Processar_ArquivoRetorno_Cobranca_Cedula(objCobrancaEletronica As ClassCobrancaEletronica) As Long
'Faz o processamento do arquivo de retorno

Dim lErro As Long
Dim iIndice As Integer
Dim sRegistro As String
Dim iCodRegistro As Integer
Dim objBorderoRetCobr As New ClassBorderoRetCobr
Dim lComando As Long
Dim lComando1 As Long

On Error GoTo Erro_Processar_ArquivoRetorno_Cobranca_Cedula

    objBorderoRetCobr.iBanco = objCobrancaEletronica.objCobrador.iCodBanco
    
    'le as carteiras de cobranca do cobrador
    lErro = CF("CarteirasCobrador_Le", objCobrancaEletronica.objCobrador, objBorderoRetCobr.colCartCobr)
    If lErro <> SUCESSO Then Error 32265
    
    'Abre os comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 51772

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 51773

    'Abre o arquivo de retorno
    Open objCobrancaEletronica.sNomeArquivoRetorno For Input As #2
    
    Set objBorderoRetCobr.objMovCCI = New ClassMovContaCorrente
    Set objBorderoRetCobr.colOcorrRetParcRec = New Collection
    objBorderoRetCobr.objMovCCI.dValor = 0

    
    'Até chegar ao fim do arquivo
    Do While Not EOF(2)
    
        'Busca o próximo registro do arquivo (na 1a vez vai ser o de header)
        Line Input #2, sRegistro
    
        'Pega o código de identificação do registro
        iCodRegistro = StrParaInt(Left(sRegistro, 1))
        
        'Se o registro for o de Header
        If iCodRegistro = 0 Then
            'Processa os dados contidos no Header
            lErro = Processa_Registro_Header(sRegistro, objBorderoRetCobr, objCobrancaEletronica.objCobrador, objCobrancaEletronica.sNomeArquivoRetorno)
            If lErro <> SUCESSO Then Error 51774

        'Se o registro for de Detalhe ("Transação")
        ElseIf iCodRegistro = 1 Then
            'Processa os dados contidos no registro de detalhe
            lErro = Processa_Registro_Detalhe(sRegistro, objBorderoRetCobr, objCobrancaEletronica, lComando, lComando1)
            If lErro <> SUCESSO Then Error 51775
            
            Call objCobrancaEletronica.objTelaAtualizacao.Mostra_Evolucao(False, 1)
        
        'Se o registro for de Trailer
        ElseIf iCodRegistro = 9 Then
            'Processa os dados do trailer do arquivo
            lErro = Processa_Registro_Trailer(sRegistro)
            If lErro <> SUCESSO Then Error 51776
        
        Else
            Error 51779
        End If
        
    Loop
        
    Close #2

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    'Grava os dados recolhidos durante o processamento do arquivo de retorno
    lErro = CF("BorderoRetCobr_ConfirmaGrava", objBorderoRetCobr)
    If lErro <> SUCESSO Then Error 51777

    Processar_ArquivoRetorno_Cobranca_Cedula = SUCESSO
    
    Exit Function
    
Erro_Processar_ArquivoRetorno_Cobranca_Cedula:

    Processar_ArquivoRetorno_Cobranca_Cedula = Err
    
    Select Case Err
    
        Case 51771
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 51772, 51773
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 51774, 51775, 51776, 51777, 32265
        
        Case 51779
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORMATO_ARQUIVO_INCORRETO", Err, objCobrancaEletronica.sNomeArquivoRetorno)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147226)
    
    End Select

    Close #2

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Private Function Processa_Registro_Header(sRegistro As String, objBorderoRetCobr As ClassBorderoRetCobr, objCobrador As ClassCobrador, sNomeArquivo As String) As Long

Dim lErro As Long
Dim tRegHeaderCedula As typeRetHeaderCedula
Dim objContaCorrente As New ClassContasCorrentesInternas

On Error GoTo Erro_Processa_Registro_Header

    'Extrai da variável lida no arquivo os dados do Header
    lErro = Extrair_Header_Retorno(tRegHeaderCedula, sRegistro)
    If lErro <> SUCESSO Then Error 51780
    
    'Preenche o objBorderoRetCobr
    objBorderoRetCobr.dtDataRecepcao = gdtDataHoje
    objBorderoRetCobr.iCobrador = objCobrador.iCodigo
    objBorderoRetCobr.sNomeArq = sNomeArquivo
    
    'Pega a conta corrente do cobrador
    objContaCorrente.iCodigo = objCobrador.iCodCCI
    
    'Lê a conta corrente do cobrador
    lErro = CF("ContaCorrenteInt_Le", objContaCorrente.iCodigo, objContaCorrente)
    If lErro <> SUCESSO And lErro <> 11807 Then Error 51781
    
    'Se não achou a Conta Corrente --> erro
    If lErro <> SUCESSO Then Error 51782
    
    'Preenche parcialmente o objMovCCI
    objBorderoRetCobr.objMovCCI.iFilialEmpresa = objContaCorrente.iFilialEmpresa
    objBorderoRetCobr.objMovCCI.iCodConta = objContaCorrente.iCodigo
    objBorderoRetCobr.objMovCCI.iTipo = MOVCCI_RECEBIMENTO_TITULO
    objBorderoRetCobr.objMovCCI.iExcluido = NAO_EXCLUIDO
    objBorderoRetCobr.objMovCCI.iTipoMeioPagto = DINHEIRO
    objBorderoRetCobr.objMovCCI.dtDataBaixa = gdtDataHoje
    objBorderoRetCobr.objMovCCI.dtDataMovimento = Extrair_Data(tRegHeaderCedula.sDataCredito)
    objBorderoRetCobr.objMovCCI.iConciliado = NAO_CONCILIADO
           
    Processa_Registro_Header = SUCESSO
    
    Exit Function

Erro_Processa_Registro_Header:

    Processa_Registro_Header = Err

    Select Case Err
    
        Case 51780, 51781
    
        Case 51782
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_CORRENTE_NAO_ENCONTRADA", Err, objContaCorrente.iCodigo)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147227)
            
    End Select
    
    Exit Function

End Function

Private Function Extrair_Header_Retorno(tRegHeaderCedula As typeRetHeaderCedula, sRegistro As String) As Long
'Extrai do sRegistro os dados do header do arquido de retorno

Dim lErro As Long

On Error GoTo Erro_Extrair_Header_Retorno

    With tRegHeaderCedula
        .sIdentificacaoReg = Mid(sRegistro, 1, 1)
        .sIdentificacaoArq = Mid(sRegistro, 2, 1)
        .sLiteralRetorno = Mid(sRegistro, 3, 7)
        .sCodTipoServico = Mid(sRegistro, 10, 2)
        .sLiteralServico = Mid(sRegistro, 12, 15)
        .sCodEmpresa = Mid(sRegistro, 27, 20)
        .sNomeEmpresa = Mid(sRegistro, 47, 30)
        .sNumeroBanco = Mid(sRegistro, 77, 3)
        .sNomeBanco = Mid(sRegistro, 80, 15)
        .sDataEmissaoArq = Mid(sRegistro, 95, 6)
        .sZeros = Mid(sRegistro, 101, 8)
        .sNumAviso = Mid(sRegistro, 109, 5)
        .sComplemento1 = Mid(sRegistro, 114, 266)
        .sDataCredito = Mid(sRegistro, 380, 6)
        .sComplemento2 = Mid(sRegistro, 386, 9)
        .sSequencialRegistro = Mid(sRegistro, 395, 6)
    
    End With

    Extrair_Header_Retorno = SUCESSO

    Exit Function

Erro_Extrair_Header_Retorno:

    Extrair_Header_Retorno = Err
    
    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147228)
            
    End Select
    
    Exit Function

End Function


Private Function Processa_Registro_Detalhe(sRegistro As String, objBorderoRetCobr As ClassBorderoRetCobr, objCobrancaEletronica As ClassCobrancaEletronica, lComando As Long, lComando1 As Long) As Long

Dim tRegDetalhe As typeRetDetalheCedula
Dim iCodOcorrencia As Integer, iTipoOcorrencia As Integer
Dim lErro As Long
Dim objDetRetCobr As New ClassDetRetCobr
Dim objBorderoRetCobrCart As ClassBorderoRetCobrCart

On Error GoTo Erro_Processa_Registro_Detalhe

    'Extrai da variável lida no arquivo os dados do registro de detalhe
    lErro = Extrair_Detalhe_Retorno(tRegDetalhe, sRegistro)
    If lErro <> SUCESSO Then Error 51783
    
    'Recolhe o código da ocorrência
    iCodOcorrencia = StrParaInt(tRegDetalhe.sCodOcorrencia)
    
    Select Case iCodOcorrencia
        
        'Quando ocorre a baixa
        Case 6, 9, 10
            iTipoOcorrencia = RETCOBR_DET_BAIXA
    
        Case 3
            iTipoOcorrencia = RETCOBR_DET_REJEITADO
                                  
        Case Else
            iTipoOcorrencia = RETCOBR_DET_IGNORAR
    
    End Select
    
    If iTipoOcorrencia <> RETCOBR_DET_IGNORAR Then
    
        lErro = Prepara_CartCobrRet1(objCobrancaEletronica, objBorderoRetCobr, objBorderoRetCobrCart, objDetRetCobr, tRegDetalhe)
        If lErro <> SUCESSO Then Error 32257
        
        lErro = CF("Processa_Registro_Detalhe_Geral", objBorderoRetCobr, objCobrancaEletronica, iTipoOcorrencia, RETCOBR_TIPO_SEU_NUMERO4, objDetRetCobr, objBorderoRetCobrCart, lComando)
        If lErro <> SUCESSO Then Error 32356
                                  
    End If
    
    Processa_Registro_Detalhe = SUCESSO
    
    Exit Function

Erro_Processa_Registro_Detalhe:

    Processa_Registro_Detalhe = Err

    Select Case Err
    
        Case 32356, 32357 'Inserido por Wagner
    
        Case 51783, 51784, 51786, 51788
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147229)
            
    End Select
    
    Exit Function

End Function

Private Function Extrair_Detalhe_Retorno(tRegDetalheCedula As typeRetDetalheCedula, sRegistro As String) As Long
'Extrai do sRegistro os dados do Detalhe do arquido de retorno

Dim lErro As Long

On Error GoTo Erro_Extrair_Detalhe_Retorno

    With tRegDetalheCedula
        
        .sIdentificacaoReg = Mid(sRegistro, 1, 1)
        .sTipoInscricaoEmpresa = Mid(sRegistro, 2, 2)
        .sNumInscricaoEmpresa = Mid(sRegistro, 4, 14)
        .sZeros = Mid(sRegistro, 18, 3)
        .sIdentfEmpresa = Mid(sRegistro, 21, 17)
        .sNumControle = Mid(sRegistro, 38, 25)
        .sZeros1 = Mid(sRegistro, 63, 8)
        .sNossoNumero1 = Mid(sRegistro, 71, 12)
        .sUsoDoBanco1 = Mid(sRegistro, 83, 10)
        .sUsoDoBanco2 = Mid(sRegistro, 93, 12)
        .sIndicadorRateio = Mid(sRegistro, 105, 1)
        .sZeros2 = Mid(sRegistro, 106, 2)
        .sCodCarteira = Mid(sRegistro, 108, 1)
        .sCodOcorrencia = Mid(sRegistro, 109, 2)
        .sDataOcorrencia = Mid(sRegistro, 111, 6)
        .sNumDocto = Mid(sRegistro, 117, 10)
        .sNossoNumero2 = Mid(sRegistro, 127, 20)
        .sVencimento = Mid(sRegistro, 147, 6)
        .sValorTitulo = Mid(sRegistro, 153, 13)
        .sBancoCobrador = Mid(sRegistro, 166, 3)
        .sAgenciaCobradora = Mid(sRegistro, 169, 5)
        .sEspecie = Mid(sRegistro, 174, 2)
        .sTarifaCobranca = Mid(sRegistro, 176, 13)
        .sOutrasDespesas = Mid(sRegistro, 189, 13)
        .sJurosOp = Mid(sRegistro, 202, 13)
        .sValorIOF = Mid(sRegistro, 215, 13)
        .sAbatimento = Mid(sRegistro, 228, 13)
        .sDescontos = Mid(sRegistro, 241, 13)
        .sValorPrincipal = Mid(sRegistro, 254, 13)
        .sJuros = Mid(sRegistro, 267, 13)
        .sOutrosCreditos = Mid(sRegistro, 280, 13)
        .sComplemento1 = Mid(sRegistro, 293, 2)
        .sMotivoOcorr1 = Mid(sRegistro, 295, 1)
        .sDataCredito = Mid(sRegistro, 296, 6)
        .sComplemento2 = Mid(sRegistro, 302, 17)
        .sMotivoOcorr2 = Mid(sRegistro, 319, 10)
        .sComplemento3 = Mid(sRegistro, 329, 66)
        .sNumSequencialRegistro = Mid(sRegistro, 395, 6)
    
    End With
    
    Extrair_Detalhe_Retorno = SUCESSO

    Exit Function

Erro_Extrair_Detalhe_Retorno:

    Extrair_Detalhe_Retorno = Err
    
    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147230)
            
    End Select
    
    Exit Function

End Function
Function Processa_Registro_Trailer(sRegistroTrailer As String) As Long

Dim lErro As Long
Dim tRegTrailerCedula As typeRetTrailerCedula

On Error GoTo Erro_Processa_Registro_Trailer

    'Extrai da variável lida no arquivo os dados do registro de detalhe
    lErro = Extrair_Trailer_Retorno(tRegTrailerCedula, sRegistroTrailer)
    If lErro <> SUCESSO Then Error 51789
    
    Processa_Registro_Trailer = SUCESSO
    
    Exit Function
    
Erro_Processa_Registro_Trailer:

    Processa_Registro_Trailer = Err
    
    Select Case Err
    
        Case 51789
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147231)
            
    End Select
    
    Exit Function
    
End Function
Private Function Extrair_Trailer_Retorno(tRegTrailerCedula As typeRetTrailerCedula, sRegistro As String) As Long
'Extrai do sRegistro os dados do Trailer do arquido de retorno

Dim lErro As Long

On Error GoTo Erro_Extrair_Trailer_Retorno

    tRegTrailerCedula.sIdentificacaoReg = Mid(sRegistro, 1, 1)
    tRegTrailerCedula.sCodigoRetorno = Mid(sRegistro, 2, 1)
    tRegTrailerCedula.sIdentificacaoReg = Mid(sRegistro, 3, 2)
    tRegTrailerCedula.sNumeroBanco = Mid(sRegistro, 5, 3)
    tRegTrailerCedula.sComplemento1 = Mid(sRegistro, 8, 10)
    tRegTrailerCedula.sQuantTitulosCobr = Mid(sRegistro, 18, 8)
    tRegTrailerCedula.sValorTotalCobr = Mid(sRegistro, 26, 14)
    tRegTrailerCedula.sAvisoBancario = Mid(sRegistro, 40, 8)
    tRegTrailerCedula.sComplemento2 = Mid(sRegistro, 48, 10)
    tRegTrailerCedula.sQuantRegOcorr_02 = Mid(sRegistro, 58, 5)
    tRegTrailerCedula.sValorRegOcorr_02 = Mid(sRegistro, 63, 12)
    tRegTrailerCedula.sValorRegOcorr_06 = Mid(sRegistro, 75, 12)
    tRegTrailerCedula.sQuantRegOcorr_06 = Mid(sRegistro, 87, 5)
    tRegTrailerCedula.sValorRegOcorr_06_09_10 = Mid(sRegistro, 92, 12)
    tRegTrailerCedula.sQuantRegOcorr_09_10 = Mid(sRegistro, 104, 5)
    tRegTrailerCedula.sValorRegOcorr_09_10 = Mid(sRegistro, 109, 12)
    tRegTrailerCedula.sQuantRegOcorr_13 = Mid(sRegistro, 121, 5)
    tRegTrailerCedula.sValorRegOcorr_13 = Mid(sRegistro, 126, 12)
    tRegTrailerCedula.sQuantRegOcorr_14 = Mid(sRegistro, 138, 5)
    tRegTrailerCedula.sValorRegOcorr_14 = Mid(sRegistro, 143, 12)
    tRegTrailerCedula.sQuantRegOcorr_12 = Mid(sRegistro, 155, 5)
    tRegTrailerCedula.sValorRegOcorr_12 = Mid(sRegistro, 160, 12)
    tRegTrailerCedula.sQuantRegOcorr_19 = Mid(sRegistro, 172, 5)
    tRegTrailerCedula.sValorRegOcorr_19 = Mid(sRegistro, 177, 12)
    tRegTrailerCedula.sComplemento3 = Mid(sRegistro, 189, 174)
    tRegTrailerCedula.sValorTotalRateios = Mid(sRegistro, 363, 15)
    tRegTrailerCedula.sQuantTotalRateios = Mid(sRegistro, 378, 8)
    tRegTrailerCedula.sSequencialRegistro = Mid(sRegistro, 395, 6)

    Extrair_Trailer_Retorno = SUCESSO

    Exit Function

Erro_Extrair_Trailer_Retorno:

    Extrair_Trailer_Retorno = Err
    
    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147232)
            
    End Select
    
    Exit Function

End Function

Function Extrair_Valor(vValor As Variant) As Double
'Extrai o conteúdo numérico (double) do valor passado
Dim sValor As String

    If Len(Trim(vValor)) = 0 Then
        Extrair_Valor = 0
    Else
        
        sValor = Left(vValor, Len(vValor) - 2)
        Extrair_Valor = StrParaDbl(sValor & "," & Right(vValor, 2))
        
    End If

    Exit Function

End Function

Function Extrair_Data(vData As Variant) As Date
'Tenta extrair uma data no formato "DDMMYY" de vData.
'Caso não consiga, devolve DATA_NULA
    
Dim sData As String

    sData = (Mid(vData, 1, 2) & "/" & Mid(vData, 3, 2) & "/" & Mid(vData, 5, 2))
    
    If IsDate(sData) Then
        Extrair_Data = CDate(sData)
    Else
        Extrair_Data = DATA_NULA
    End If
    
    Exit Function

End Function

Function Obter_FilialEmpresa(iFilialEmpresa As Integer, objFilialEmpresa As AdmFiliais) As Long
'Retorna o objFilialEmpresa prenchido com os dados da filialempresa passada por parâmetro

Dim lErro As Long
Dim objFilialEmpresaAux As AdmFiliais
Dim iIndice As Integer
Dim iPosicao As Integer

On Error GoTo Erro_Obter_FilialEmpresa

    iPosicao = 0
    iIndice = 0
    
    'Procura a filial empresa passada na coleção das filiais já lidas
    For Each objFilialEmpresaAux In gColFiliaisEmpresa
        iIndice = iIndice + 1
        If objFilialEmpresaAux.iCodFilial = iFilialEmpresa Then
            Set objFilialEmpresa = objFilialEmpresaAux
            Exit For
        End If
    Next
    
    'Se não encontrou na coleção
    If iPosicao = 0 Then
        'Carrega o objFilialEmpresa com a sua chave
        objFilialEmpresa.lCodEmpresa = glEmpresa
        objFilialEmpresa.iCodFilial = iFilialEmpresa
        
        'Lê a FilialEmpresa
        lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 27378 Then Error 51666
        If lErro <> SUCESSO Then Error 51667
    
        'Adiciona na coleção das filiais lidas
        gColFiliaisEmpresa.Add objFilialEmpresa
    End If
    
    Obter_FilialEmpresa = SUCESSO
    
    Exit Function

Erro_Obter_FilialEmpresa:

    Obter_FilialEmpresa = Err
    
    Select Case Err
    
        Case 51666
        
        Case 51667
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", Err, objFilialEmpresa.iCodFilial)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147233)
            
    End Select
    
    Exit Function
    
End Function

Private Sub Extrair_Agencia_DV(sAgenciaDVAg As String, sAgencia As String, iTamanhoAg As Integer, sDV As String, iTamanhoDV As Integer)
'Extrai da string passada na forma "Agencia-DV", a agência e o DV separadamente

Dim iPosSeparador As Integer

    'Procura o separador da Ag e o DV
    iPosSeparador = InStr(sAgenciaDVAg, SEPARADOR)
    'Se encontrou
    If iPosSeparador > 0 Then
        'Recolhe a agência e o DV
        sAgencia = FormataCpoNum(Left(sAgenciaDVAg, iPosSeparador - 1), iTamanhoAg)
        sDV = Mid(sAgenciaDVAg, iPosSeparador + 1, iTamanhoDV)
    'Se o DV não foi informado
    Else
        'Recolhe a agência
        sAgencia = FormataCpoNum(sAgenciaDVAg, iTamanhoAg)
        sDV = FormataCpoTexto("", iTamanhoDV)
    End If
    
    Exit Sub
    
End Sub

Private Sub Extrair_Conta_DV(sContaDVCta As String, sConta As String, iTamanhoCta As Integer, sDV As String, iTamanhoDV As Integer)
'Extrai da string passada na forma "Conta-DV", a Conta e o DV separadamente

Dim iPosSeparador As Integer

    'Procura o separador da Conta e o DV
    iPosSeparador = InStr(sContaDVCta, SEPARADOR)
    'Se encontrou
    If iPosSeparador > 0 Then
        'Recolhe a conta e o DV
        sConta = FormataCpoNum(Left(sContaDVCta, iPosSeparador - 1), iTamanhoCta)
        sDV = Mid(sContaDVCta, iPosSeparador + 1, iTamanhoDV)
    'Se o DV não foi informado
    Else
        'Recolhe a agência
        sConta = FormataCpoNum(sContaDVCta, iTamanhoCta)
        sDV = FormataCpoTexto("", iTamanhoDV)
    End If
    
    Exit Sub
    
End Sub

Private Function Obter_Info_Cobrador(colInformacoes As Collection, iCodInformacao As Integer) As String

Dim objCodNome As AdmCodigoNome

    For Each objCodNome In colInformacoes
        If objCodNome.iCodigo = iCodInformacao Then
            Obter_Info_Cobrador = objCodNome.sNome
            Exit Function
            Exit For
        End If
    Next
    
    Obter_Info_Cobrador = ""
    
End Function

Function Calcula_NossoNumero_Cedula(objCarteiraCobrador As ClassCarteiraCobrador, sNossoNumero As String) As Long

Dim sProxNossoNumero As String
Dim iCarteira As Integer
Dim sDVNossoNumero As String

On Error GoTo Erro_Calcula_NossoNumero_Cedula
    
    iCarteira = objCarteiraCobrador.iNumCarteiraNoBanco
    
    If objCarteiraCobrador.iComRegistro = CARTEIRA_SEM_REGISTRO Then iCarteira = "06"

    sNossoNumero = objCarteiraCobrador.sFaixaNossoNumeroProx
    
    If Len(Trim(sNossoNumero)) = 0 Then gError 79703
    
    sProxNossoNumero = sNossoNumero + 1
    
    If Len(Trim(objCarteiraCobrador.sFaixaNossoNumeroFinal)) > 0 Then
        
        If Val(sNossoNumero) > Val(objCarteiraCobrador.sFaixaNossoNumeroFinal) Then sProxNossoNumero = FormataCpoNum(objCarteiraCobrador.sFaixaNossoNumeroInicial, 11)
    End If
    
    objCarteiraCobrador.sFaixaNossoNumeroProx = sProxNossoNumero
    
    Call Calcula_DV_NossoNumero_Cedula(FormataCpoNum(iCarteira, 2) & sNossoNumero, sDVNossoNumero, 11, 7)
    
    sNossoNumero = sNossoNumero & sDVNossoNumero
    
    Calcula_NossoNumero_Cedula = SUCESSO
    
    Exit Function
    
Erro_Calcula_NossoNumero_Cedula:

    Calcula_NossoNumero_Cedula = gErr
    
    Select Case gErr
    
        Case 79703
            Call Rotina_Erro(vbOKOnly, "ERRO_NOSSO_NUMERO_NAO_DEFINIDO", gErr, objCarteiraCobrador.iCodCarteiraCobranca, objCarteiraCobrador.iCobrador)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147234)
    
    End Select

    Exit Function

End Function

Private Sub Calcula_DV_NossoNumero_Cedula(sNossoNumero As String, sDVNossoNumero As String, iModulo As Integer, iBase As Integer)

Dim iIndice As Integer
Dim iTamNossoNum As Integer
Dim iDigito1 As Integer
Dim iVarModulo As Integer
Dim iProduto As Integer
Dim iSoma As Integer
Dim iResto As Integer
Dim iResultado As Integer
Dim sResultado As Integer

    iVarModulo = 2
    iSoma = 0
    
    iTamNossoNum = Len(sNossoNumero)
    
    For iIndice = iTamNossoNum To 1 Step -1

        iDigito1 = StrParaInt(Mid(sNossoNumero, iIndice, 1))
        iProduto = (iDigito1 * iVarModulo)
        
        iSoma = iSoma + iProduto
        
        iVarModulo = iVarModulo + 1
        
        If iVarModulo > iBase Then iVarModulo = 2
    
    Next
    
    iResto = iSoma Mod iModulo
    
    If iResto = 0 Then
        iResultado = 0
    Else
        iResultado = 11 - iResto
    End If
    
    If iResultado = 10 Then
        sDVNossoNumero = "P"
    Else
        sDVNossoNumero = FormataCpoNum(iResultado, 1)
    End If
    
    Exit Sub
    
End Sub

Private Sub RegDetCobr_ObterInfo(objDetRetCobr As ClassDetRetCobr, tRegDetalhe As typeRetDetalheCedula)
'copia dados do type para o objeto

    'Recolhe alguns dados do registro de detalhe
    With objDetRetCobr
    
        .dValorRecebido = Extrair_Valor(tRegDetalhe.sValorPrincipal)
        .dtDataVencimento = Extrair_Data(tRegDetalhe.sVencimento)
        .dtDataCredito = Extrair_Data(tRegDetalhe.sDataCredito)
        If .dtDataCredito = DATA_NULA Then
            .dtDataCredito = Extrair_Data(tRegDetalhe.sDataOcorrencia)
        End If
        .dValorJuros = Extrair_Valor(tRegDetalhe.sJuros)
        .dValorDesconto = Extrair_Valor(tRegDetalhe.sDescontos)
        .sNossoNumero = tRegDetalhe.sNossoNumero1
        .iCodOcorrencia = StrParaInt(tRegDetalhe.sCodOcorrencia)
        .iCodCarteiraNoBanco = StrParaInt(tRegDetalhe.sCodCarteira)
        .dValorTarifa = Extrair_Valor(tRegDetalhe.sTarifaCobranca)
        .sSeuNumero = tRegDetalhe.sNumControle
        .sUsoDaEmpresa = tRegDetalhe.sNumControle
        .lNumSequencial = CLng(tRegDetalhe.sNumSequencialRegistro)
        .dValorTitulo = Extrair_Valor(tRegDetalhe.sValorTitulo)
        
    End With
    
End Sub

Private Function Prepara_CartCobrRet1(objCobrancaEletronica As ClassCobrancaEletronica, objBorderoRetCobr As ClassBorderoRetCobr, objBorderoRetCobrCart As ClassBorderoRetCobrCart, objDetRetCobr As ClassDetRetCobr, tRegDetalhe As typeRetDetalheCedula) As Long
'codigo comum ao tratamento de registros detalhe de baixas, custas e tarifas
'preenche objBorderoRetCobrCart e objDetRetCobr

Dim lErro As Long

On Error GoTo Erro_Prepara_CartCobrRet1

    Call RegDetCobr_ObterInfo(objDetRetCobr, tRegDetalhe)
    
    objDetRetCobr.dtDataBaixa = Extrair_Data(tRegDetalhe.sDataOcorrencia)
    
    lErro = CF("Prepara_CartCobrRet", objCobrancaEletronica, objBorderoRetCobr, objBorderoRetCobrCart, objDetRetCobr)
    If lErro <> SUCESSO Then Error 32255
    
    Prepara_CartCobrRet1 = SUCESSO
    
    Exit Function
    
Erro_Prepara_CartCobrRet1:

    Prepara_CartCobrRet1 = Err
    
    Select Case Err
    
        Case 32255
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147235)
            
    End Select

    Exit Function

End Function


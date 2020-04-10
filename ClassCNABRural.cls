VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCNABRural"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Dim glNumArquivoRemessa As Long
Dim gColFiliaisEmpresa As Collection
Dim gobjCobrancaEletronica As ClassCobrancaEletronica
Public gobjGeracaoArqCNABPag As ClassGeracaoArqCNABPag

Function Criar_Arquivo_Remessa_Rural(objCobrancaEletronica As ClassCobrancaEletronica) As Long
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

On Error GoTo Erro_Criar_Arquivo_Remessa_Rural
    
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
    
    'Gera o nome do arquivo de acordo com o sequencial do Cobrador
    lErro = CF("CobrancaEletronica_Gera_NomeArquivo", objCobrancaEletronica.objCobrador, objCobrancaEletronica.sDiretorio, sNomeArquivo, lSeqNomeArq)
    If lErro <> SUCESSO Then gError 51756
    
    gobjCobrancaEletronica.lSeqNomeArq = lSeqNomeArq
    
    'Cria o arquivo de remessa
    Open sNomeArquivo For Output As #1
        
    lErro = Montar_RegistroHeader_Rural(sRegHeader)
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
            lErro = Montar_RegistroDetalhe_Rural(sRegDetalhe, objOcorrRemParc, objCarteiraCobrador)
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
    lErro = Montar_RegistroTrailer_Rural(sRegTrailer, objOcorrRemParc)
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
    
    Criar_Arquivo_Remessa_Rural = SUCESSO
    
    Exit Function

Erro_Criar_Arquivo_Remessa_Rural:

    Criar_Arquivo_Remessa_Rural = gErr
    
    Select Case gErr
    
        Case 51754
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 51755
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 51756, 51757, 51758, 51759, 51760, 51761, 51762, 79705
        
        Case 62010
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147357)
    
    End Select
    
    Close #1
    
    Call Transacao_Rollback

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function


Private Function Montar_RegistroHeader_Rural(sRegHeader As String) As Long
'Monta o registro

Dim tRegHeaderRural As typeRemHeaderRural
Dim lErro As Long

On Error GoTo Erro_Montar_RegistroHeader_Rural

    gobjCobrancaEletronica.lSequencialRegistro = 1

    'Carrega tRegHeaderRural com os dados do Header do aquivo
    lErro = Obter_Dados_Header(tRegHeaderRural)
    If lErro <> SUCESSO Then Error 51763
        
    With tRegHeaderRural
    
        'Monta o registro de Header do Rural
        sRegHeader = .sIdentificacaoReg & .sIdentificacaoArq & .sLiteralRemessa & .sCodTipoServico & .sLiteralServico
        sRegHeader = sRegHeader & .sCodigoEmpresa & .sBrancosCNAB1 & .sNomeEmpresa & .sNumeroBanco & .sNomeBanco & .sBrancosCNAB2
        sRegHeader = sRegHeader & .sCodigoVersao & .sBrancosCNAB3 & .sSequencialRegistro
        
    End With
    
    Montar_RegistroHeader_Rural = SUCESSO

    Exit Function

Erro_Montar_RegistroHeader_Rural:

    Montar_RegistroHeader_Rural = Err

    Select Case Err

        Case 51763

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147358)

    End Select

    Exit Function

End Function

Private Function Obter_Dados_Header(tRegHeaderRural As typeRemHeaderRural) As Long

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
    tRegHeaderRural.sIdentificacaoReg = "0"
    tRegHeaderRural.sIdentificacaoArq = "9"
    tRegHeaderRural.sLiteralRemessa = FormataCpoTexto("REMESSA", 7)
    tRegHeaderRural.sCodTipoServico = "01"
    
    If gobjCobrancaEletronica.colBorderos.Count <> 0 Then
        If gobjCobrancaEletronica.colBorderos(1).iCodCarteiraCobranca = CARTEIRA_DESCONTADA Then
            tRegHeaderRural.sLiteralServico = FormataCpoTexto("DESCONTO", 15)
        Else
            tRegHeaderRural.sLiteralServico = FormataCpoTexto("COBRANCA", 15)
        End If
    End If
    
    tRegHeaderRural.sCodigoEmpresa = FormataCpoNum(Obter_Info_Cobrador(gobjCobrancaEletronica.objCobradorCNABInfo.colInformacoes, 1), 14)
    tRegHeaderRural.sBrancosCNAB1 = FormataCpoTexto("", 6)
    tRegHeaderRural.sNomeEmpresa = FormataCpoTexto(gsNomeEmpresa, 30)
    tRegHeaderRural.sNumeroBanco = "453"
    tRegHeaderRural.sNomeBanco = FormataCpoTexto("RURAL", 15)
    tRegHeaderRural.sBrancosCNAB2 = FormataCpoTexto("", 15)
    tRegHeaderRural.sCodigoVersao = "CD0V01"
    tRegHeaderRural.sBrancosCNAB3 = FormataCpoTexto("", 279)
    tRegHeaderRural.sSequencialRegistro = FormataCpoNum(gobjCobrancaEletronica.lSequencialRegistro, 6)
    
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147359)
            
    End Select
    
    Exit Function

End Function

Private Function Montar_RegistroDetalhe_Rural(sRegDetalhe As String, objOcorrRemParcRec As ClassOcorrRemParcRec, objCarteiraCobrador As ClassCarteiraCobrador) As Long

'Monta o registro

Dim tRegDetalhe As typeRemDetalheRural
Dim lErro As Long

On Error GoTo Erro_Montar_RegistroDetalhe_Rural
    
    'Obtem os dados do registro de detalhe
    lErro = Obter_Dados_Detalhe(tRegDetalhe, objOcorrRemParcRec, objCarteiraCobrador)
    If lErro <> SUCESSO And lErro <> 32265 Then Error 51766
    If lErro <> SUCESSO Then Error 32266
    
    With tRegDetalhe
        'Forma o registro de detalhe
        sRegDetalhe = .sIdentificacaoReg & .sBrancosCNAB1 & .sCodigoEmpresa & .sPrazoProtesto & .sMoeda & .sTipoMora & .sIdentifMensagem
        sRegDetalhe = sRegDetalhe & .sNumControle & .sNossoNumero & .sDVNossoNumero & .sNumContrato & .sBrancosCNAB2 & .sSacadorAvalista & .sBrancosCNAB3
        sRegDetalhe = sRegDetalhe & .sCodCarteira & .sIdentificacaoOcorr & .sNumDocto & .sDataVencto & .sValorTitulo & .sBancoCobranca & .sBrancosCNAB4 & .sEspecie
        sRegDetalhe = sRegDetalhe & .sAceite & .sDataEmissao & .sInstrucao1 & .sValorJuros & .sDataDesconto
        sRegDetalhe = sRegDetalhe & .sValorDesconto & .sValorIOF & .sAbatimento & .sIdentfInscSacado & .sNumInscricaoSacado
        sRegDetalhe = sRegDetalhe & .sNomeSacado & .sEnderecoSacado & .sComplemento & .sCEP & .sCidadeSacado & .sEstado & .sSacadorAvalista2 & .sNumSequencialRegistro
    
    End With
    
    Montar_RegistroDetalhe_Rural = SUCESSO

    Exit Function

Erro_Montar_RegistroDetalhe_Rural:

    Montar_RegistroDetalhe_Rural = Err

    Select Case Err
    
        Case 51766, 32266

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147360)

    End Select

    Exit Function

End Function

Private Function Obter_Dados_Detalhe(tRegDetalhe As typeRemDetalheRural, objOcorrRemParcRem As ClassOcorrRemParcRec, objCarteiraCobrador As ClassCarteiraCobrador) As Long
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
''''        lErro = Calcula_NossoNumero_Rural(objCarteiraCobrador, sNossoNumero)
        If lErro <> SUCESSO Then Error 51770
    End If
    
    With tRegDetalhe
        .sIdentificacaoReg = "1"
        .sBrancosCNAB1 = FormataCpoTexto("", 16)
        .sCodigoEmpresa = FormataCpoNum(Obter_Info_Cobrador(gobjCobrancaEletronica.objCobradorCNABInfo.colInformacoes, 1), 14)
        .sPrazoProtesto = FormataCpoNum(Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 3), 2)
        .sMoeda = "00" 'Real
        .sTipoMora = "1" 'Em reais
        .sIdentifMensagem = " "
        .sNumControle = FormataCpoTexto(objOcorrRemParcRem.lNumIntDoc, 25)
        .sNossoNumero = FormataCpoNum(sNossoNumero, 7)
        .sDVNossoNumero = FormataCpoNum(sNossoNumero, 1)
    
        If objCarteiraCobrador.iCodCarteiraCobranca = CARTEIRA_DESCONTADA Then
            .sNumContrato = FormataCpoNum(0, 10)
        Else
            .sNumContrato = FormataCpoNum(Obter_Info_Cobrador(gobjCobrancaEletronica.objCobradorCNABInfo.colInformacoes, 2), 10)
        End If
        
        .sBrancosCNAB2 = FormataCpoTexto("", 2)
        .sSacadorAvalista = FormataCpoNum("", 16) '2+14
        .sBrancosCNAB3 = FormataCpoTexto("", 9)
        .sCodCarteira = FormataCpoTexto(objCarteiraCobrador.sCodCarteiraNoBanco, 1)
        .sIdentificacaoOcorr = FormataCpoNum(objOcorrRemParcRem.iCodOcorrencia, 2)
        .sNumDocto = FormataCpoTexto(gobjCobrancaEletronica.lNumTitulo & "/" & objParcelaRec.iNumParcela, "10")
        .sDataVencto = Format(objParcelaRec.dtDataVencimento, "DDMMYY")
        .sValorTitulo = FormataCpoValor(objParcelaRec.dSaldo, 13)
        .sBancoCobranca = FormataCpoNum("", 7)
        .sBrancosCNAB4 = FormataCpoTexto("", 1)
        
        If objCarteiraCobrador.iImprimeBoleta = EMPRESA_IMPRIME_BOLETA Then
            .sEspecie = FormataCpoNum(1, 2) 'Duplicata
        Else
            .sEspecie = FormataCpoNum(21, 2) 'Duplicata
        End If
        If objParcelaRec.iAceite = 1 Then
            .sAceite = "A"
        Else
            .sAceite = "B"
        End If
        .sDataEmissao = Format(gobjCobrancaEletronica.dtDataEmissaoTitulo, "DDMMYY")
        .sInstrucao1 = FormataCpoNum("", 4)
        
        If objCarteiraCobrador.iCodCarteiraCobranca = CARTEIRA_DESCONTADA Then
            .sValorJuros = FormataCpoValor(0, 13)
        Else
            .sValorJuros = FormataCpoValor(objParcelaRec.dSaldo * gobjCobrancaEletronica.dPercJurosDiario, 13)
        End If
        
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
        .sNomeSacado = FormataCpoTexto(UCase(gobjCobrancaEletronica.sNomeSacado), 40)
        .sEnderecoSacado = FormataCpoTexto(Trim(objEndereco.sEndereco), 40)
        .sComplemento = FormataCpoTexto("", 12)
        .sCEP = Formata_Cep(objEndereco.sCEP, 8)
        .sCidadeSacado = FormataCpoTexto(objEndereco.sCidade, 15)
        .sEstado = FormataCpoTexto(objEndereco.sSiglaEstado, 2)
        .sSacadorAvalista2 = FormataCpoTexto("1", 43) '1+42
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147361)

    End Select

    Exit Function

End Function


Private Function Montar_RegistroTrailer_Rural(sRegTrailer As String, objCobrador As ClassCobrador) As Long
'Monta o registro

Dim tRegTrailerRural As typeRemTrailerRural
Dim lErro As Long

On Error GoTo Erro_Montar_RegistroTrailer_Rural

    'Preenche os dados do regsitro de trailer
    tRegTrailerRural.sIdentificacaoReg = "9"
    tRegTrailerRural.sComplemento1 = FormataCpoTexto("", 393)
    tRegTrailerRural.sSequencialRegistro = FormataCpoNum(gobjCobrancaEletronica.lSequencialRegistro, 6)
    
    'Monta o Registro de trailer
    sRegTrailer = tRegTrailerRural.sIdentificacaoReg & tRegTrailerRural.sComplemento1 & tRegTrailerRural.sSequencialRegistro

    Montar_RegistroTrailer_Rural = SUCESSO

    Exit Function

Erro_Montar_RegistroTrailer_Rural:

    Montar_RegistroTrailer_Rural = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147362)

    End Select

    Exit Function

End Function


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

Function Processar_ArquivoRetorno_Cobranca_Rural(objCobrancaEletronica As ClassCobrancaEletronica) As Long
'Faz o processamento do arquivo de retorno

Dim lErro As Long
Dim iIndice As Integer
Dim sRegistro As String
Dim iCodRegistro As Integer
Dim objBorderoRetCobr As New ClassBorderoRetCobr
Dim lComando As Long
Dim lComando1 As Long

On Error GoTo Erro_Processar_ArquivoRetorno_Cobranca_Rural

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

    Processar_ArquivoRetorno_Cobranca_Rural = SUCESSO
    
    Exit Function
    
Erro_Processar_ArquivoRetorno_Cobranca_Rural:

    Processar_ArquivoRetorno_Cobranca_Rural = Err
    
    Select Case Err
    
        Case 51771
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 51772, 51773
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 51774, 51775, 51776, 51777, 32265
        
        Case 51779
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORMATO_ARQUIVO_INCORRETO", Err, objCobrancaEletronica.sNomeArquivoRetorno)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147363)
    
    End Select

    Close #2

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Private Function Processa_Registro_Header(sRegistro As String, objBorderoRetCobr As ClassBorderoRetCobr, objCobrador As ClassCobrador, sNomeArquivo As String) As Long

Dim lErro As Long
Dim tRegHeaderRural As typeRetHeaderRural
Dim objContaCorrente As New ClassContasCorrentesInternas

On Error GoTo Erro_Processa_Registro_Header

    'Extrai da variável lida no arquivo os dados do Header
    lErro = Extrair_Header_Retorno(tRegHeaderRural, sRegistro)
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
    objBorderoRetCobr.objMovCCI.dtDataMovimento = gdtDataHoje
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147364)
            
    End Select
    
    Exit Function

End Function

Private Function Extrair_Header_Retorno(tRegHeaderRural As typeRetHeaderRural, sRegistro As String) As Long
'Extrai do sRegistro os dados do header do arquido de retorno

Dim lErro As Long

On Error GoTo Erro_Extrair_Header_Retorno

    With tRegHeaderRural
        .sIdentificacaoReg = Mid(sRegistro, 1, 1)
        .sIdentificacaoArq = Mid(sRegistro, 2, 1)
        .sLiteralRetorno = Mid(sRegistro, 3, 7)
        .sCodTipoServico = Mid(sRegistro, 10, 2)
        .sLiteralServico = Mid(sRegistro, 12, 15)
        .sCodEmpresa = Mid(sRegistro, 27, 14)
        .sNomeEmpresa = Mid(sRegistro, 47, 30)
        .sNumeroBanco = Mid(sRegistro, 77, 3)
        .sNomeBanco = Mid(sRegistro, 80, 15)
        .sDataEmissaoArq = Mid(sRegistro, 95, 6)
        .sZeros = Mid(sRegistro, 101, 294)
        .sSequencialRegistro = Mid(sRegistro, 395, 6)
    
    End With

    Extrair_Header_Retorno = SUCESSO

    Exit Function

Erro_Extrair_Header_Retorno:

    Extrair_Header_Retorno = Err
    
    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147365)
            
    End Select
    
    Exit Function

End Function



Private Function Processa_Registro_Detalhe(sRegistro As String, objBorderoRetCobr As ClassBorderoRetCobr, objCobrancaEletronica As ClassCobrancaEletronica, lComando As Long, lComando1 As Long) As Long

Dim tRegDetalhe As typeRetDetalheRural
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
        Case 6, 9, 10, 15, 16
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
    
        Case 51783, 51784, 51786, 51788
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147366)
            
    End Select
    
    Exit Function

End Function

Private Function Extrair_Detalhe_Retorno(tRegDetalhe As typeRetDetalheRural, sRegistro As String) As Long
'Extrai do sRegistro os dados do Detalhe do arquido de retorno

Dim lErro As Long

On Error GoTo Erro_Extrair_Detalhe_Retorno

    With tRegDetalhe
        
        .sIdentificacaoReg = Mid(sRegistro, 1, 1)
        .sTipoInscricaoEmpresa = Mid(sRegistro, 2, 2)
        .sNumInscricaoEmpresa = Mid(sRegistro, 4, 14)
        .sIdentfEmpresa = Mid(sRegistro, 18, 14)
        .sDiasProtesto = Mid(sRegistro, 32, 2)
        .sMoeda = Mid(sRegistro, 34, 2)
        .sNumControle = Mid(sRegistro, 38, 25)
        .sNossoNumero1 = Mid(sRegistro, 63, 7)
        .sDVNossoNumero1 = Mid(sRegistro, 70, 1)
        .sNumContrato = Mid(sRegistro, 71, 10)
        .sUsoDoBanco2 = Mid(sRegistro, 93, 12)
        .sIndicadorRateio = Mid(sRegistro, 105, 1)
        .sZeros2 = Mid(sRegistro, 106, 2)
        .sCodCarteira = Mid(sRegistro, 108, 1)
        .sCodOcorrencia = Mid(sRegistro, 109, 2)
        .sDataOcorrencia = Mid(sRegistro, 111, 6)
        .sNumDocto = Mid(sRegistro, 117, 15)
        .sVencimento = Mid(sRegistro, 147, 6)
        .sValorTitulo = Mid(sRegistro, 153, 13)
        .sBancoCobrador = Mid(sRegistro, 166, 7)
        .sEspecie = Mid(sRegistro, 174, 2)
        .sTarifaCobranca = Mid(sRegistro, 176, 13)
        .sValorIOF = Mid(sRegistro, 215, 13)
        .sAbatimento = Mid(sRegistro, 228, 13)
        .sDescontos = Mid(sRegistro, 241, 13)
        .sValorPrincipal = Mid(sRegistro, 254, 13)
        .sJuros = Mid(sRegistro, 267, 13)
        .sOutrosCreditos = Mid(sRegistro, 280, 13)
        .sDataCredito = Mid(sRegistro, 296, 6)
        .sTabelaErros = Mid(sRegistro, 334, 61)
        .sNumSequencialRegistro = Mid(sRegistro, 395, 6)
    
    End With
    
    Extrair_Detalhe_Retorno = SUCESSO

    Exit Function

Erro_Extrair_Detalhe_Retorno:

    Extrair_Detalhe_Retorno = Err
    
    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147367)
            
    End Select
    
    Exit Function

End Function
Function Processa_Registro_Trailer(sRegistroTrailer As String) As Long

Dim lErro As Long
Dim tRegTrailerRural As typeRetTrailerRural

On Error GoTo Erro_Processa_Registro_Trailer

    'Extrai da variável lida no arquivo os dados do registro de detalhe
    lErro = Extrair_Trailer_Retorno(tRegTrailerRural, sRegistroTrailer)
    If lErro <> SUCESSO Then Error 51789
    
    Processa_Registro_Trailer = SUCESSO
    
    Exit Function
    
Erro_Processa_Registro_Trailer:

    Processa_Registro_Trailer = Err
    
    Select Case Err
    
        Case 51789
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147368)
            
    End Select
    
    Exit Function
    
End Function
Private Function Extrair_Trailer_Retorno(tRegTrailerRural As typeRetTrailerRural, sRegistro As String) As Long
'Extrai do sRegistro os dados do Trailer do arquido de retorno

Dim lErro As Long

On Error GoTo Erro_Extrair_Trailer_Retorno

    tRegTrailerRural.sIdentificacaoReg = Mid(sRegistro, 1, 1)
    tRegTrailerRural.sCodigoRetorno = Mid(sRegistro, 2, 1)
    tRegTrailerRural.sIdentificacaoReg = Mid(sRegistro, 3, 2)
    tRegTrailerRural.sNumeroBanco = Mid(sRegistro, 5, 3)
    tRegTrailerRural.sComplemento1 = Mid(sRegistro, 8, 10)
    tRegTrailerRural.sQuantTitulosCobr = Mid(sRegistro, 18, 8)
    tRegTrailerRural.sValorTotalCobr = Mid(sRegistro, 26, 14)
    tRegTrailerRural.sAvisoBancario = Mid(sRegistro, 40, 8)
    tRegTrailerRural.sQuantTitulosCobrDireta = Mid(sRegistro, 48, 8)
    tRegTrailerRural.sValorTotalCobrDireta = Mid(sRegistro, 56, 14)
    tRegTrailerRural.sQuantTitulosNaoCad = Mid(sRegistro, 70, 8)
    tRegTrailerRural.sValorTotalTitNaoCad = Mid(sRegistro, 78, 14)
    tRegTrailerRural.sSequencialRegistro = Mid(sRegistro, 395, 6)

    Extrair_Trailer_Retorno = SUCESSO

    Exit Function

Erro_Extrair_Trailer_Retorno:

    Extrair_Trailer_Retorno = Err
    
    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147369)
            
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

Function Obter_Info_CarteiraCobrador(iCodCarteiraCobranca As Integer, colCarteirasCNABInfo As Collection, iCodInformacao As Integer) As String

Dim objCarteiraInfo As ClassCNABInfo
Dim objCodNome As AdmCodigoNome

    For Each objCarteiraInfo In colCarteirasCNABInfo
    
        If objCarteiraInfo.iCarteiraCobrador = iCodCarteiraCobranca Then
            For Each objCodNome In objCarteiraInfo.colInformacoes
                If objCodNome.iCodigo = iCodInformacao Then
                    Obter_Info_CarteiraCobrador = objCodNome.sNome
                    Exit Function
                    Exit For
                End If
            Next
        End If
    Next

    Obter_Info_CarteiraCobrador = ""

    Exit Function

End Function

Private Sub RegDetCobr_ObterInfo(objDetRetCobr As ClassDetRetCobr, tRegDetalhe As typeRetDetalheRural)
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

Private Function Prepara_CartCobrRet1(objCobrancaEletronica As ClassCobrancaEletronica, objBorderoRetCobr As ClassBorderoRetCobr, objBorderoRetCobrCart As ClassBorderoRetCobrCart, objDetRetCobr As ClassDetRetCobr, tRegDetalhe As typeRetDetalheRural) As Long
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147370)
            
    End Select

    Exit Function

End Function


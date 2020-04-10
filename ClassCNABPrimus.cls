VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCNABPrimus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim gColFiliaisEmpresa As Collection
Dim gobjCobrancaEletronica As ClassCobrancaEletronica

Function Criar_Arquivo_Remessa(bCobrancaExpressa As Boolean, ParamArray alComando()) As Long
'Cria uma arquivo de remessa com base nos dados do BD

Dim iCancela As Integer
Dim lErro As Long
Dim sRegHeader As String
Dim sNomeArquivo As String
Dim iIndice As Integer
Dim sRegDetalhe As String
Dim sRegTrailer As String, lSeqNomeArq As Long
Dim objBorderoCobranca As New ClassBorderoCobranca
Dim objCarteiraCobrador As New ClassCarteiraCobrador
Dim objOcorrRemParc As ClassOcorrRemParcRec

On Error GoTo Erro_Criar_Arquivo_Remessa
    
    Set gColFiliaisEmpresa = New Collection
    
    'Gera o nome do arquivo de acordo com o sequencial do Cobrador
    lErro = CF("CobrancaEletronica_Gera_NomeArquivo", gobjCobrancaEletronica.objCobrador, gobjCobrancaEletronica.sDiretorio, sNomeArquivo, lSeqNomeArq)
    If lErro <> SUCESSO Then gError 51756
    
    gobjCobrancaEletronica.lSeqNomeArq = lSeqNomeArq
    
    'Cria o arquivo de remessa
    Open sNomeArquivo For Output As #1
        
    lErro = Montar_RegistroHeader_Primus(sRegHeader, bCobrancaExpressa)
    If lErro <> SUCESSO Then gError 51760
        
    'Inclui o registro de Header
    Print #1, DesacentuaTexto(sRegHeader)

    For Each objBorderoCobranca In gobjCobrancaEletronica.colBorderos
        'Lê o bordero de cobrança
        lErro = CF("BorderoCobranca_LeMod", alComando(0), objBorderoCobranca)
        If lErro <> SUCESSO Then gError 51757
            
        objCarteiraCobrador.iCobrador = gobjCobrancaEletronica.iCobrador
        objCarteiraCobrador.iCodCarteiraCobranca = objBorderoCobranca.iCodCarteiraCobranca
        
        'Lê a carteira do cobrador
        lErro = CF("CarteiraCobrador_LeMod", objCarteiraCobrador, alComando(1))
        If lErro <> SUCESSO Then gError 51759
       
        'Busca no BD as Ocorrências do Bordero em questão
        lErro = CF("CobrancaEletronica_Obter_OcorrRemParcRec", alComando(2), objBorderoCobranca, gobjCobrancaEletronica)
        If lErro <> SUCESSO Then gError 51662
        
        'Para cada ocorrência
        For Each objOcorrRemParc In gobjCobrancaEletronica.colOcorrRemParcRec
    
            sRegDetalhe = ""
            'Monta um registro de detalhe do arquivo de remessa
            lErro = Montar_RegistroDetalhe_Primus(sRegDetalhe, objOcorrRemParc, objCarteiraCobrador)
            If lErro <> SUCESSO And lErro <> 32266 Then gError 51761
            
            If lErro = SUCESSO Then
            
                'Inclui o registro com os detalhes do Título(Parcela) a receber
                Print #1, DesacentuaTexto(sRegDetalhe)
                
            End If
            
        Next
        
        objBorderoCobranca.iProcessado = BORDERO_PROCESSADO
        objBorderoCobranca.sNomeArquivo = sNomeArquivo
        
        'Atualiza o Status do Bordero de Cobrança
        lErro = CF("BorderoCobranca_Atualiza_Processo", alComando(0), alComando(2), objBorderoCobranca)
        If lErro <> SUCESSO Then gError 51758
    
        'Atualiza o NossoNumero em carteiraCobrador
        lErro = CF("CarteiraCobrador_AtualizaNossoNumero", objCarteiraCobrador, alComando(1), alComando(3))
        If lErro <> SUCESSO Then gError 51707
    
        iCancela = 0
        
        Call gobjCobrancaEletronica.objTelaAtualizacao.Mostra_Evolucao(iCancela, 1)
        
        If iCancela = CANCELA_BATCH Then gError 79705
        
    Next
    
    
    'MOnta o registro de trailer
    lErro = Montar_RegistroTrailer_Primus(sRegTrailer, objOcorrRemParc)
    If lErro <> SUCESSO Then gError 51762

    'Coloca o registro de trailer no arquivo
    Print #1, DesacentuaTexto(sRegTrailer)
    
    Close #1
    
    
    Criar_Arquivo_Remessa = SUCESSO
    
    Exit Function

Erro_Criar_Arquivo_Remessa:

    Criar_Arquivo_Remessa = gErr
    
    Select Case gErr
    
        Case 51756, 51757, 51758, 51759, 51760, 51761, 51762, 79705
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147310)
    
    End Select
    
    Close #1
    
    Exit Function
    
End Function

Private Function Montar_RegistroHeader_Primus(sRegHeader As String, bCobrancaExpressa As Boolean) As Long
'Monta o registro

Dim tRegHeaderPrimus As typeRemHeaderPrimus
Dim lErro As Long

On Error GoTo Erro_Montar_RegistroHeader_Primus

    gobjCobrancaEletronica.lSequencialRegistro = 1

    'Carrega tRegHeaderPrimus com os dados do Header do aquivo
    lErro = Obter_Dados_Header(tRegHeaderPrimus, bCobrancaExpressa)
    If lErro <> SUCESSO Then Error 51763
        
    With tRegHeaderPrimus
    
        'Monta o registro de Header do Primus
        sRegHeader = .sIdentificacaoReg & .sBrancosCNAB1 & .sCodigoEmpresa & .sBrancosCNAB2 & .sNomeEmpresa
        sRegHeader = sRegHeader & .sBrancosCNAB3 & .sDataEmissaoArq & .sBrancosCNAB4 & .sContaCobrancaDireta & .sBrancosCNAB5 & .sNumSequencial

    End With
    
    Montar_RegistroHeader_Primus = SUCESSO

    Exit Function

Erro_Montar_RegistroHeader_Primus:

    Montar_RegistroHeader_Primus = Err

    Select Case Err

        Case 51763

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147311)

    End Select

    Exit Function

End Function

Private Function Obter_Dados_Header(tRegHeaderPrimus As typeRemHeaderPrimus, bCobrancaExpressa As Boolean) As Long

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
    tRegHeaderPrimus.sIdentificacaoReg = "0"
    tRegHeaderPrimus.sBrancosCNAB1 = FormataCpoTexto("", 25)
    tRegHeaderPrimus.sCodigoEmpresa = FormataCpoNum(Obter_Info_Cobrador(gobjCobrancaEletronica.objCobradorCNABInfo.colInformacoes, 1), 12)
    tRegHeaderPrimus.sBrancosCNAB2 = FormataCpoTexto("", 8)
    tRegHeaderPrimus.sNomeEmpresa = FormataCpoTexto(gsNomeEmpresa, 30)
    tRegHeaderPrimus.sBrancosCNAB3 = FormataCpoTexto("", 18)
    tRegHeaderPrimus.sDataEmissaoArq = Format(gdtDataHoje, "DDMMYY")
    tRegHeaderPrimus.sBrancosCNAB4 = FormataCpoTexto("", 8)
    tRegHeaderPrimus.sContaCobrancaDireta = FormataCpoNum(IIf(bCobrancaExpressa, "", Obter_Info_Cobrador(gobjCobrancaEletronica.objCobradorCNABInfo.colInformacoes, 2)), 12)
    tRegHeaderPrimus.sBrancosCNAB5 = FormataCpoTexto("", 274)
    tRegHeaderPrimus.sNumSequencial = FormataCpoNum(1, 6)
    
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147312)
            
    End Select
    
    Exit Function

End Function

Private Function Montar_RegistroDetalhe_Primus(sRegDetalhe As String, objOcorrRemParcRec As ClassOcorrRemParcRec, objCarteiraCobrador As ClassCarteiraCobrador) As Long

'Monta o registro

Dim tRegDetalhe As typeRemDetalhePrimus
Dim lErro As Long

On Error GoTo Erro_Montar_RegistroDetalhe_Primus
    
    'Obtem os dados do registro de detalhe
    lErro = Obter_Dados_Detalhe(tRegDetalhe, objOcorrRemParcRec, objCarteiraCobrador)
    If lErro <> SUCESSO And lErro <> 32265 Then Error 51766
    If lErro <> SUCESSO Then Error 32266
    
    With tRegDetalhe
        'Forma o registro de detalhe
        sRegDetalhe = .sIdentificacaoReg & .sBrancosCNAB1 & .sCGCEmpresa & .sCodigoEmpresa & .sBrancosCNAB2 & .sUsoEmpresa
        sRegDetalhe = sRegDetalhe & .sNossoNumero & .sBrancosCNAB3 & .sBancoCobrancaDireta & .sBrancosCNAB4 & .sCodigoCarteira
        sRegDetalhe = sRegDetalhe & .sIdentificacaoOcorr & .sSeuNumero & .sDataVencto & .sValorTitulo & .sBrancosCNAB5 & .sEspecie
        sRegDetalhe = sRegDetalhe & .sAceite & .sDataEmissao & .sInstrucao1 & .sInstrucao2 & .sValorJuros & .sDataDesconto
        sRegDetalhe = sRegDetalhe & .sValorDesconto & .sBrancosCNAB6 & .sAbatimento & .sIdentfInscSacado & .sNumInscricaoSacado
        sRegDetalhe = sRegDetalhe & .sNomeSacado & .sBrancosCNAB7 & .sEnderecoSacado & .sBairroSacado & .sCEPSacado & .sCidadeSacado & .sUFSacado & .sMensagem1 & .sBrancosCNAB8
        sRegDetalhe = sRegDetalhe & .sDataJurosMora & .sPrazo & .sBrancosCNAB9 & .sNumSequencialRegistro & .sTipoEmitente
    
    End With
    
    Montar_RegistroDetalhe_Primus = SUCESSO

    Exit Function

Erro_Montar_RegistroDetalhe_Primus:

    Montar_RegistroDetalhe_Primus = Err

    Select Case Err
    
        Case 51766, 32266

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147313)

    End Select

    Exit Function

End Function

Private Function Obter_Dados_Detalhe(tRegDetalhe As typeRemDetalhePrimus, objOcorrRemParcRem As ClassOcorrRemParcRec, objCarteiraCobrador As ClassCarteiraCobrador) As Long
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
            
    With tRegDetalhe
        .sIdentificacaoReg = "1"
        .sBrancosCNAB1 = FormataCpoTexto("", 2)
        .sCGCEmpresa = FormataCpoNum(gobjCobrancaEletronica.sCGCSacado, 14)
        .sCodigoEmpresa = FormataCpoNum(Obter_Info_Cobrador(gobjCobrancaEletronica.objCobradorCNABInfo.colInformacoes, 4), 12)
        .sBrancosCNAB2 = FormataCpoTexto("", 8)
        .sUsoEmpresa = FormataCpoTexto(objOcorrRemParcRem.lNumIntDoc, "25")
        .sNossoNumero = FormataCpoTexto(sNossoNumero, 12)
        .sBrancosCNAB3 = FormataCpoTexto("", 8)
        .sBancoCobrancaDireta = FormataCpoNum(IIf(objCarteiraCobrador.iImprimeBoleta = BANCO_IMPRIME_BOLETA, 0, Obter_Info_Cobrador(gobjCobrancaEletronica.objCobradorCNABInfo.colInformacoes, 3)), 3)
        .sBrancosCNAB4 = FormataCpoTexto("", 21)
        
        If objCarteiraCobrador.iCodCarteiraCobranca = CARTEIRA_SIMPLES Then
        
            If objCarteiraCobrador.iImprimeBoleta = BANCO_IMPRIME_BOLETA Then
                .sCodigoCarteira = FormataCpoNum(1, 2)
                'express
            Else
                'direta
                .sCodigoCarteira = FormataCpoNum(7, 2)
            End If
        ElseIf objCarteiraCobrador.iCodCarteiraCobranca = CARTEIRA_CAUCIONADA Then
            If objCarteiraCobrador.iImprimeBoleta = BANCO_IMPRIME_BOLETA Then
                'express
                .sCodigoCarteira = FormataCpoNum(2, 2)
            Else
                'direta
                .sCodigoCarteira = FormataCpoNum(6, 2)
            End If
        Else
            .sCodigoCarteira = FormataCpoNum(objCarteiraCobrador.sCodCarteiraNoBanco, 2)
        End If
        
        .sIdentificacaoOcorr = FormataCpoNum(objOcorrRemParcRem.iCodOcorrencia, 2)
        .sSeuNumero = FormataCpoTexto(gobjCobrancaEletronica.lNumTitulo & "/" & objParcelaRec.iNumParcela, "10")
        .sDataVencto = Format(objParcelaRec.dtDataVencimento, "DDMMYY")
        .sValorTitulo = FormataCpoValor(objParcelaRec.dSaldo, 13)
        .sBrancosCNAB5 = FormataCpoTexto("", 8)
        .sEspecie = FormataCpoNum(1, 2) 'Duplicata Mercantil
        If objParcelaRec.iAceite = 1 Then
            .sAceite = "A"
        Else
            .sAceite = "N"
        End If
        .sDataEmissao = Format(gobjCobrancaEletronica.dtDataEmissaoTitulo, "DDMMYY")
        .sInstrucao1 = FormataCpoNum(objOcorrRemParcRem.iInstrucao1, 2)
        .sInstrucao2 = FormataCpoNum(objOcorrRemParcRem.iInstrucao2, 2)
        .sValorJuros = FormataCpoValor(objParcelaRec.dSaldo * gobjCobrancaEletronica.dPercJurosDiario, 13)
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
        .sBrancosCNAB6 = FormataCpoTexto("", 13)
        .sAbatimento = FormataCpoValor(0, 13)
        If Len(gobjCobrancaEletronica.sCGCSacado) = 14 Then
            .sIdentfInscSacado = "02"
        Else
            .sIdentfInscSacado = "01"
        End If
        .sNumInscricaoSacado = FormataCpoNum(gobjCobrancaEletronica.sCGCSacado, 14)
        .sNomeSacado = FormataCpoTexto(gobjCobrancaEletronica.sNomeSacado, 30)
        .sBrancosCNAB7 = FormataCpoTexto("", 10)
        .sEnderecoSacado = FormataCpoTexto(Trim(objEndereco.sEndereco), 40)
        .sBairroSacado = FormataCpoTexto(Trim(objEndereco.sBairro), 12)
        .sCEPSacado = Formata_Cep(objEndereco.sCEP, 8)
        .sCidadeSacado = FormataCpoTexto(Trim(objEndereco.sCidade), 15)
        .sUFSacado = FormataCpoTexto(Trim(objEndereco.sSiglaEstado), 2)
        .sMensagem1 = FormataCpoTexto("", 30)
        .sBrancosCNAB8 = FormataCpoTexto("", 4)
        .sDataJurosMora = "000000"
        .sPrazo = "00"
        .sBrancosCNAB9 = FormataCpoTexto("", 1)
        .sNumSequencialRegistro = FormataCpoNum(gobjCobrancaEletronica.lSequencialRegistro, 6)
        .sTipoEmitente = FormataCpoTexto("", 102)
'        .sNomeEmitente
'        .sEnderecoEmitente
'        .sCidadeEmitente
'        .sUFEmitente
    
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147314)

    End Select

    Exit Function

End Function

Private Function Montar_RegistroTrailer_Primus(sRegTrailer As String, objCobrador As ClassCobrador) As Long
'Monta o registro

Dim tRegTrailerPrimus As typeRemTrailerPrimus
Dim lErro As Long

On Error GoTo Erro_Montar_RegistroTrailer_Primus

    'Preenche os dados do regsitro de trailer
    tRegTrailerPrimus.sIdentificacaoReg = "9"
    tRegTrailerPrimus.sComplemento1 = FormataCpoTexto("", 393)
    tRegTrailerPrimus.sSequencialRegistro = FormataCpoNum(gobjCobrancaEletronica.lSequencialRegistro, 6)
    
    'Monta o Registro de trailer
    sRegTrailer = tRegTrailerPrimus.sIdentificacaoReg & tRegTrailerPrimus.sComplemento1 & tRegTrailerPrimus.sSequencialRegistro

    Montar_RegistroTrailer_Primus = SUCESSO

    Exit Function

Erro_Montar_RegistroTrailer_Primus:

    Montar_RegistroTrailer_Primus = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147315)

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

Function Processar_ArquivoRetorno_Cobranca_Primus(objCobrancaEletronica As ClassCobrancaEletronica) As Long
'Faz o processamento do arquivo de retorno

Dim lErro As Long
Dim iIndice As Integer
Dim sRegistro As String
Dim iCodRegistro As Integer
Dim objBorderoRetCobr As New ClassBorderoRetCobr
Dim lComando As Long
Dim lComando1 As Long

On Error GoTo Erro_Processar_ArquivoRetorno_Cobranca_Primus

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

    Processar_ArquivoRetorno_Cobranca_Primus = SUCESSO
    
    Exit Function
    
Erro_Processar_ArquivoRetorno_Cobranca_Primus:

    Processar_ArquivoRetorno_Cobranca_Primus = Err
    
    Select Case Err
    
        Case 51771
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 51772, 51773
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 51774, 51775, 51776, 51777, 32265
        
        Case 51779
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORMATO_ARQUIVO_INCORRETO", Err, objCobrancaEletronica.sNomeArquivoRetorno)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147316)
    
    End Select

    Close #2

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function


Private Function Processa_Registro_Header(sRegistro As String, objBorderoRetCobr As ClassBorderoRetCobr, objCobrador As ClassCobrador, sNomeArquivo As String) As Long

Dim lErro As Long
Dim tRegHeaderPrimus As typeRetHeaderPrimus
Dim objContaCorrente As New ClassContasCorrentesInternas

On Error GoTo Erro_Processa_Registro_Header

    'Extrai da variável lida no arquivo os dados do Header
    lErro = Extrair_Header_Retorno(tRegHeaderPrimus, sRegistro)
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147317)
            
    End Select
    
    Exit Function

End Function


Private Function Extrair_Header_Retorno(tRegHeaderPrimus As typeRetHeaderPrimus, sRegistro As String) As Long
'Extrai do sRegistro os dados do header do arquido de retorno

Dim lErro As Long

On Error GoTo Erro_Extrair_Header_Retorno

    With tRegHeaderPrimus
        .sIdentificacaoReg = Mid(sRegistro, 1, 1)
        .sIdentificacaoArq = Mid(sRegistro, 2, 1)
        .sCodEmpresa = Mid(sRegistro, 27, 12)
        .sNumeroBanco = Mid(sRegistro, 77, 3)
        .sContaCobranca = Mid(sRegistro, 109, 12)
        .sSequencialRegistro = Mid(sRegistro, 395, 6)
    
    End With

    Extrair_Header_Retorno = SUCESSO

    Exit Function

Erro_Extrair_Header_Retorno:

    Extrair_Header_Retorno = Err
    
    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147318)
            
    End Select
    
    Exit Function

End Function

Private Function Processa_Registro_Detalhe(sRegistro As String, objBorderoRetCobr As ClassBorderoRetCobr, objCobrancaEletronica As ClassCobrancaEletronica, lComando As Long, lComando1 As Long) As Long

Dim tRegDetalhe As typeRetDetalhePrimus
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
    
        Case 51783, 51784, 51786, 51788
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147319)
            
    End Select
    
    Exit Function

End Function

Private Function Extrair_Detalhe_Retorno(tRegDetalhe As typeRetDetalhePrimus, sRegistro As String) As Long
'Extrai do sRegistro os dados do Detalhe do arquido de retorno

Dim lErro As Long

On Error GoTo Erro_Extrair_Detalhe_Retorno

    With tRegDetalhe
        
        .sIdentificacaoReg = Mid(sRegistro, 1, 1)
        .sIdentfEmpresa = Mid(sRegistro, 18, 12)
        .sNumControle = Mid(sRegistro, 38, 25)
        .sNossoNumero1 = Mid(sRegistro, 63, 12)
        .sCodCarteira = Mid(sRegistro, 107, 2)
        .sCodOcorrencia = Mid(sRegistro, 109, 2)
        .sDataOcorrencia = Mid(sRegistro, 111, 6)
        .sNumDocto = Mid(sRegistro, 117, 10)
        .sVencimento = Mid(sRegistro, 147, 6)
        .sValorTitulo = Mid(sRegistro, 153, 13)
        .sBancoCobrador = Mid(sRegistro, 166, 3)
        .sTarifaCobranca = Mid(sRegistro, 176, 13)
        .sAbatimento = Mid(sRegistro, 228, 13)
        .sDescontos = Mid(sRegistro, 241, 13)
        .sValorPrincipal = Mid(sRegistro, 254, 13)
        .sJuros = Mid(sRegistro, 267, 13)
        .sOutrosCreditos = Mid(sRegistro, 280, 13)
        .sCodigoErro = Mid(sRegistro, 378, 8)
        .sNumSequencialRegistro = Mid(sRegistro, 395, 6)
    
    End With
    
    Extrair_Detalhe_Retorno = SUCESSO

    Exit Function

Erro_Extrair_Detalhe_Retorno:

    Extrair_Detalhe_Retorno = Err
    
    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147320)
            
    End Select
    
    Exit Function

End Function

Function Processa_Registro_Trailer(sRegistroTrailer As String) As Long

Dim lErro As Long
Dim tRegTrailerPrimus As typeRetTrailerPrimus

On Error GoTo Erro_Processa_Registro_Trailer

    'Extrai da variável lida no arquivo os dados do registro de detalhe
    lErro = Extrair_Trailer_Retorno(tRegTrailerPrimus, sRegistroTrailer)
    If lErro <> SUCESSO Then Error 51789
    
    Processa_Registro_Trailer = SUCESSO
    
    Exit Function
    
Erro_Processa_Registro_Trailer:

    Processa_Registro_Trailer = Err
    
    Select Case Err
    
        Case 51789
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147321)
            
    End Select
    
    Exit Function
    
End Function
Private Function Extrair_Trailer_Retorno(tRegTrailerPrimus As typeRetTrailerPrimus, sRegistro As String) As Long
'Extrai do sRegistro os dados do Trailer do arquido de retorno

Dim lErro As Long

On Error GoTo Erro_Extrair_Trailer_Retorno

    tRegTrailerPrimus.sIdentificacaoReg = Mid(sRegistro, 1, 1)
    tRegTrailerPrimus.sSequencialRegistro = Mid(sRegistro, 395, 6)

    Extrair_Trailer_Retorno = SUCESSO

    Exit Function

Erro_Extrair_Trailer_Retorno:

    Extrair_Trailer_Retorno = Err
    
    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147322)
            
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


Private Function Separa_Borderos_Direta_Express(colBorderos As Collection, colBorderosDireta As Collection, colBorderosExpress As Collection) As Long
'Separa os borderos das carteiras com registro das carteiras sem registro.

Dim lErro As Long
Dim lComando As Long
Dim iRegistro As Integer
Dim objCarteiraCob As New ClassCarteiraCobrador
Dim objBordero As New ClassBorderoCobranca

On Error GoTo Erro_Separa_Borderos_Direta_Express
    
    'Para cada bordero da coleção lida
    For Each objBordero In colBorderos
        
        objCarteiraCob.iCobrador = objBordero.iCobrador
        objCarteiraCob.iCodCarteiraCobranca = objBordero.iCodCarteiraCobranca
        'Lê a carteira do cobrdor
        lErro = CF("CarteiraCobrador_Le", objCarteiraCob)
        If lErro <> SUCESSO And lErro <> 23551 Then Error 62015
        If lErro <> SUCESSO Then Error 62016 'Não encontrou
        
        'Se a carteira é registrada
        If objCarteiraCob.iImprimeBoleta = BANCO_IMPRIME_BOLETA Then
            'Adiciona o bordero na coleção de borderos c\ registro
            colBorderosExpress.Add objBordero
        Else
            'Adiciona o bordero na coleção de borderos sem registro
            colBorderosDireta.Add objBordero
        End If
    
    Next
    
    Separa_Borderos_Direta_Express = SUCESSO
    
    Exit Function
    
Erro_Separa_Borderos_Direta_Express:
    
    Separa_Borderos_Direta_Express = Err
    
    Select Case Err
    
        Case 62015
        
        Case 62016
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CARTEIRACOBRADOR_NAO_CADASTRADA1", Err, objCarteiraCob.iCodCarteiraCobranca, objCarteiraCob.iCobrador)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147323)
            
    End Select

    Exit Function
    
End Function

Function Criar_Arquivo_Remessa_BancoPrimus(objCobrancaEletronica As ClassCobrancaEletronica) As Long
'Cria uma arquivo de remessa com base nos dados do BD

Dim lErro As Long
Dim lTransacao As Long
Dim alComando(0 To 7) As Long
Dim iIndice As Integer
Dim colBorderosDireta As New Collection
Dim colBorderosExpress As New Collection

On Error GoTo Erro_Criar_Arquivo_Remessa_BancoPrimus
    
    Set gobjCobrancaEletronica = objCobrancaEletronica
    Set gColFiliaisEmpresa = New Collection
     
     'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 51885
    Next
        
    'ABre a transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 51886
    
   
    'Separa os borderos com e sem registro
    lErro = Separa_Borderos_Direta_Express(gobjCobrancaEletronica.colBorderos, colBorderosDireta, colBorderosExpress)
    If lErro <> SUCESSO Then Error 51887
    
    If colBorderosDireta.Count > 0 Then
    
        Set gobjCobrancaEletronica.colBorderos = colBorderosDireta
    
        lErro = Criar_Arquivo_Remessa(False, alComando(0), alComando(1), alComando(2), alComando(3))
        If lErro <> SUCESSO Then Error 51888
    
    End If
    
    If colBorderosExpress.Count > 0 Then
        
        Set gobjCobrancaEletronica.colBorderos = colBorderosExpress
        
        lErro = Criar_Arquivo_Remessa(True, alComando(4), alComando(5), alComando(6), alComando(7))
        If lErro <> SUCESSO Then Error 51889
    
    End If
    
    Close #1
    
    lErro = Transacao_Commit()
    If lErro <> SUCESSO Then Error 51890
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Criar_Arquivo_Remessa_BancoPrimus = SUCESSO
    
    Exit Function

Erro_Criar_Arquivo_Remessa_BancoPrimus:

    Criar_Arquivo_Remessa_BancoPrimus = Err
    
    Select Case Err
    
        Case 51885
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 51886
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
        
        Case 51887, 51888, 51889
        
        Case 51890
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 147324)
    
    End Select

    Call gobjCobrancaEletronica.objTelaAtualizacao.Mostra_Evolucao(True, 1)
    
    Close #1

    Call Transacao_Rollback

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Private Sub RegDetCobr_ObterInfo(objDetRetCobr As ClassDetRetCobr, tRegDetalhe As typeRetDetalhePrimus)
'copia dados do type para o objeto

    'Recolhe alguns dados do registro de detalhe
    With objDetRetCobr
    
        .dValorRecebido = Extrair_Valor(tRegDetalhe.sValorPrincipal)
        .dtDataVencimento = Extrair_Data(tRegDetalhe.sVencimento)
        .dtDataCredito = Extrair_Data(tRegDetalhe.sDataOcorrencia)
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

Private Function Prepara_CartCobrRet1(objCobrancaEletronica As ClassCobrancaEletronica, objBorderoRetCobr As ClassBorderoRetCobr, objBorderoRetCobrCart As ClassBorderoRetCobrCart, objDetRetCobr As ClassDetRetCobr, tRegDetalhe As typeRetDetalhePrimus) As Long
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147325)
            
    End Select

    Exit Function

End Function


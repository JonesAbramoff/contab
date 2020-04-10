VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCNABBandeirantes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Dim glNumArquivoRemessa As Long
Dim gobjCobrancaEletronica As ClassCobrancaEletronica

Function Criar_Arquivo_Remessa_Bandeirantes(objCobrancaEletronica As ClassCobrancaEletronica) As Long
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

On Error GoTo Erro_Criar_Arquivo_Remessa_Bandeirantes
    
    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 51754
    Next
        
    'ABre a transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 51847
    
    Set gobjCobrancaEletronica = objCobrancaEletronica
    
    'Gera o nome do arquivo de acordo com o sequencial do Cobrador
    lErro = CF("CobrancaEletronica_Gera_NomeArquivo", objCobrancaEletronica.objCobrador, gobjCobrancaEletronica.sDiretorio, sNomeArquivo, lSeqNomeArq)
    If lErro <> SUCESSO Then gError 51848
    
    gobjCobrancaEletronica.lSeqNomeArq = lSeqNomeArq
    
    'Cria o arquivo de remessa
    Open sNomeArquivo For Output As #1
        
    lErro = Montar_RegistroHeader_Bandeirantes(sRegHeader, objCobrancaEletronica.objCobrador)
    If lErro <> SUCESSO Then gError 51849
        
    'Inclui o registro de Header
    Print #1, DesacentuaTexto(sRegHeader)

    For Each objBorderoCobranca In objCobrancaEletronica.colBorderos
        'Lê o bordero de cobrança
        lErro = CF("BorderoCobranca_LeMod", alComando(2), objBorderoCobranca)
        If lErro <> SUCESSO Then gError 51850
            
        objCarteiraCobrador.iCobrador = objCobrancaEletronica.iCobrador
        objCarteiraCobrador.iCodCarteiraCobranca = objBorderoCobranca.iCodCarteiraCobranca
        
        'Lê a carteira do cobrador
        lErro = CF("CarteiraCobrador_LeMod", objCarteiraCobrador, alComando(4))
        If lErro <> SUCESSO Then gError 51851
       
        'Busca no BD as Ocorrências do Bordero em questão
        lErro = CF("CobrancaEletronica_Obter_OcorrRemParcRec", alComando(3), objBorderoCobranca, gobjCobrancaEletronica)
        If lErro <> SUCESSO Then gError 51852
        
        'Para cada ocorrência
        For Each objOcorrRemParc In objCobrancaEletronica.colOcorrRemParcRec
    
            sRegDetalhe = ""
            'Monta um registro de detalhe do arquivo de remessa
            lErro = Montar_RegistroDetalhe_Bandeirantes(sRegDetalhe, objOcorrRemParc, objCarteiraCobrador)
            If lErro <> SUCESSO And lErro <> 32263 Then gError 51853
            
            If lErro = SUCESSO Then
            
                'Inclui o registro com os detalhes do Título(Parcela) a receber
                Print #1, DesacentuaTexto(sRegDetalhe)
            
            End If
            
        Next
        
        objBorderoCobranca.iProcessado = BORDERO_PROCESSADO
        objBorderoCobranca.sNomeArquivo = sNomeArquivo
        
        'Atualiza o Status do Bordero de Cobrança
        lErro = CF("BorderoCobranca_Atualiza_Processo", alComando(2), alComando(3), objBorderoCobranca)
        If lErro <> SUCESSO Then gError 51854
    
        'Atualiza o NossoNumero em carteiraCobrador
        lErro = CF("CarteiraCobrador_AtualizaNossoNumero", objCarteiraCobrador, alComando(4), alComando(5))
        If lErro <> SUCESSO Then gError 51855
        
        iCancela = 0
        
        Call gobjCobrancaEletronica.objTelaAtualizacao.Mostra_Evolucao(iCancela, 1)
        
        If iCancela = CANCELA_BATCH Then gError 79709
    
    Next
    
    'Monta o registro de trailer
    lErro = Montar_RegistroTrailer_Bandeirantes(sRegTrailer, objOcorrRemParc)
    If lErro <> SUCESSO Then gError 51856

    'Coloca o registro de trailer no arquivo
    Print #1, DesacentuaTexto(sRegTrailer)
    
    Close #1
    
    lErro = Transacao_Commit()
    If lErro <> SUCESSO Then gError 51857
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Criar_Arquivo_Remessa_Bandeirantes = SUCESSO
    
    Exit Function

Erro_Criar_Arquivo_Remessa_Bandeirantes:

    Criar_Arquivo_Remessa_Bandeirantes = gErr
    
    Select Case gErr
    
        Case 51754
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 51847
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 51848, 51850, 51854, 51851, 51852, 51855, 51849, 51853, 51856, 51857, 79709
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 147112)
    
    End Select
    
    Close #1

    Call Transacao_Rollback

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Private Function Montar_RegistroHeader_Bandeirantes(sRegHeader As String, objCobrador As ClassCobrador) As Long
'Monta o registro

Dim tRegHeaderBandeirantes As typeRemHeaderBandeirantes
Dim lErro As Long

On Error GoTo Erro_Montar_RegistroHeader_Bandeirantes

    gobjCobrancaEletronica.lSequencialRegistro = 1

    'Carrega tRegHeaderBandeirantes com os dados do Header do arquivo
    lErro = Obter_Dados_Header(tRegHeaderBandeirantes)
    If lErro <> SUCESSO Then Error 51858
        
    With tRegHeaderBandeirantes
    
        'Monta o registro de Header do Bandeirantes
        sRegHeader = .sIdentificacaoReg & .sIdentificacaoArq & .sLiteralRemessa & .sCodTipoServico & .sLiteralServico & .sCodigoEmpresa
        sRegHeader = sRegHeader & .sBrancos1 & .sNomeEmpresa & .sNumeroBanco & .sNomeBanco & .sDataGravacao & .sDensidade & .sLiteralDensidade
        sRegHeader = sRegHeader & .sBrancos2 & .sSequencialRegistro
        
    End With
    
    Montar_RegistroHeader_Bandeirantes = SUCESSO

    Exit Function

Erro_Montar_RegistroHeader_Bandeirantes:

    Montar_RegistroHeader_Bandeirantes = Err

    Select Case Err

        Case 51858

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147113)

    End Select

    Exit Function

End Function

Private Function Obter_Dados_Header(tRegHeaderBandeirantes As typeRemHeaderBandeirantes) As Long

Dim lErro As Long
Dim objCCI As New ClassContasCorrentesInternas
Dim sCodEmpresaBanco As String

On Error GoTo Erro_Obter_Dados_Header
        
    'Lê a conta corrente associada ao cobrador
    lErro = CF("ContaCorrenteInt_Le", gobjCobrancaEletronica.objCobrador.iCodCCI, objCCI)
    If lErro <> SUCESSO And lErro <> 11807 Then Error 51859
    If lErro <> SUCESSO Then Error 51860
    
    If objCCI.iCodBanco <> gobjCobrancaEletronica.objCobrador.iCodBanco Then Error 62065
    If Len(Trim(objCCI.sAgencia)) = 0 Or Len(Trim(objCCI.sNumConta)) = 0 Then Error 62066
    
    Set gobjCobrancaEletronica.objCCI = objCCI
    
    'Carrega os campos do cabeçalho
    tRegHeaderBandeirantes.sIdentificacaoReg = "0"
    tRegHeaderBandeirantes.sIdentificacaoArq = "1"
    tRegHeaderBandeirantes.sLiteralRemessa = FormataCpoTexto("REMESSA", 7)
    tRegHeaderBandeirantes.sCodTipoServico = "01"
    tRegHeaderBandeirantes.sLiteralServico = FormataCpoTexto("COBRANCA", 15)
    tRegHeaderBandeirantes.sCodigoEmpresa = FormataCpoNum(Obter_Info_Cobrador(gobjCobrancaEletronica.objCobradorCNABInfo.colInformacoes, 1), 13)
    tRegHeaderBandeirantes.sBrancos1 = FormataCpoTexto("", 7)
    tRegHeaderBandeirantes.sNomeEmpresa = FormataCpoTexto(gsNomeEmpresa, 30)
    tRegHeaderBandeirantes.sNumeroBanco = "230"
    tRegHeaderBandeirantes.sNomeBanco = FormataCpoTexto("BANDEIRANTES", 15)
    tRegHeaderBandeirantes.sDataGravacao = Format(gdtDataHoje, "DDMMYY")
    tRegHeaderBandeirantes.sDensidade = "01600"
    tRegHeaderBandeirantes.sLiteralDensidade = "BPI"
    tRegHeaderBandeirantes.sBrancos2 = FormataCpoTexto("", 286)
    tRegHeaderBandeirantes.sSequencialRegistro = FormataCpoNum(gobjCobrancaEletronica.lSequencialRegistro, 6)
    
    'Incrementa o sequencial de registro
    gobjCobrancaEletronica.lSequencialRegistro = gobjCobrancaEletronica.lSequencialRegistro + 1

    Obter_Dados_Header = SUCESSO

    Exit Function
    
Erro_Obter_Dados_Header:

    Obter_Dados_Header = Err
    
    Select Case Err
    
        Case 62065
            lErro = Rotina_Erro(vbOKOnly, "ERRO_BANCO_CCI_DIFERENTE_COBRADOR", Err, objCCI.iCodBanco, gobjCobrancaEletronica.objCobrador.iCodBanco)
    
        Case 62066
            lErro = Rotina_Erro(vbOKOnly, "ERRO_AGENCIA_CONTA_COBRADOR_NAO_PREENCHIDAS", Err)
    
        Case 51859
        
        Case 51860
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_COBRADOR_NAO_ENCONTRADA", Err, gobjCobrancaEletronica.iCobrador)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147114)
            
    End Select
    
    Exit Function

End Function

Private Function Montar_RegistroDetalhe_Bandeirantes(sRegDetalhe As String, objOcorrRemParcRec As ClassOcorrRemParcRec, objCarteiraCobrador As ClassCarteiraCobrador) As Long
'Monta o registro de detalhe do arquivo e remessa de títulos a receber

Dim tRegDetalhe As typeRemDetalheBandeirantes
Dim lErro As Long

On Error GoTo Erro_Montar_RegistroDetalhe_Bandeirantes
    
    'Obtem os dados do registro de detalhe
    lErro = Obter_Dados_Detalhe(tRegDetalhe, objOcorrRemParcRec, objCarteiraCobrador)
    If lErro <> SUCESSO And lErro <> 32262 Then Error 51861
    If lErro <> SUCESSO Then Error 32263
    
    With tRegDetalhe
        'Forma o registro de detalhe
        sRegDetalhe = .sIdentificacaoReg & .sTipoInscricaoEmpSac & .sNumInscricaoEmpSac & .sIdentificacaoEmp & .sBrancos1 & .sUsoDaEmpresa & .sNossoNumero & .sBrancos2 & .sDiasDevolucao & .sUsoDoBanco
        sRegDetalhe = sRegDetalhe & .sCodCarteira & .sCodOcorrencia & .sSeuNumero & .sVencimento & .sValorTitulo & .sNumeroBanco & .sAgenciaCobradora & .sEspecie & .sAceite & .sDataEmissao & .sInstrucao1
        sRegDetalhe = sRegDetalhe & .sInstrucao2 & .sJurosDiarios & .sDescontoAte & .sValorDesconto & .sValorIOF & .sAbatimento & .sCodigoInscSacado & .sNumInscricaoSacado & .sNome
        sRegDetalhe = sRegDetalhe & .sEndereco & .sBairro & .sCEP & .sCidade & .sEstado & .sMensagem & .sBrancos3 & .sCodigoMoeda & .sNumSequencial
    End With
    
    Montar_RegistroDetalhe_Bandeirantes = SUCESSO

    Exit Function

Erro_Montar_RegistroDetalhe_Bandeirantes:

    Montar_RegistroDetalhe_Bandeirantes = Err

    Select Case Err
    
        Case 51861, 32263

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147115)

    End Select

    Exit Function

End Function

Private Function Obter_Dados_Detalhe(tRegDetalhe As typeRemDetalheBandeirantes, objOcorrRemParcRem As ClassOcorrRemParcRec, objCarteiraCobrador As ClassCarteiraCobrador) As Long
'Obtem os dados que são utilizados para preencher o registro de detalhe

Dim lErro As Long
Dim objParcelaRec As New ClassParcelaReceber
Dim objEndereco As New ClassEndereco
Dim sNossoNumero As String
Dim dValorDesconto As Double
Dim sCodEmpresaBanco As String

On Error GoTo Erro_Obter_Dados_Detalhe

    objParcelaRec.lNumIntDoc = objOcorrRemParcRem.lNumIntParc

    'Busca os dados da parcela a receber no BD
    lErro = CF("CobrancaEletronica_Obter_DadosParcelaRec", objParcelaRec, gobjCobrancaEletronica)
    If lErro <> SUCESSO And lErro <> 51674 Then Error 51862
    If lErro <> SUCESSO Then Error 32262
    
    lErro = CF("CobrancaEletronica_Obter_Endereco", gobjCobrancaEletronica, objEndereco)
    If lErro <> SUCESSO Then Error 51863
        
    sNossoNumero = ""
    If objCarteiraCobrador.iImprimeBoleta = EMPRESA_IMPRIME_BOLETA Then
        lErro = Calcula_NossoNumero_Bandeirantes(objCarteiraCobrador, sNossoNumero)
        If lErro <> SUCESSO Then Error 51864
    End If
    
    tRegDetalhe.sIdentificacaoReg = "1"
    If Len(gobjCobrancaEletronica.sCGCSacado) = 14 Then
        tRegDetalhe.sTipoInscricaoEmpSac = "02"
    Else
        tRegDetalhe.sTipoInscricaoEmpSac = "01"
    End If
    tRegDetalhe.sNumInscricaoEmpSac = FormataCpoTexto(gobjCobrancaEletronica.sCGCCedente, 14)
    tRegDetalhe.sIdentificacaoEmp = FormataCpoNum(Obter_Info_Cobrador(gobjCobrancaEletronica.objCobradorCNABInfo.colInformacoes, 1), 13)
    tRegDetalhe.sBrancos1 = FormataCpoTexto("", 7)
    tRegDetalhe.sUsoDaEmpresa = FormataCpoTexto(objOcorrRemParcRem.lNumIntDoc, 25)
    tRegDetalhe.sNossoNumero = FormataCpoNum(sNossoNumero, 13)
    tRegDetalhe.sBrancos2 = FormataCpoTexto("", 7)
    tRegDetalhe.sDiasDevolucao = FormataCpoNum(objOcorrRemParcRem.iDiasDeProtesto1, 2)
    tRegDetalhe.sUsoDoBanco = FormataCpoTexto("", 23)
    tRegDetalhe.sCodCarteira = FormataCpoTexto(objCarteiraCobrador.sCodCarteiraNoBanco, 1)
    tRegDetalhe.sCodOcorrencia = FormataCpoNum(objOcorrRemParcRem.iCodOcorrencia, 2)
    tRegDetalhe.sSeuNumero = FormataCpoTexto(gobjCobrancaEletronica.lNumTitulo & "/" & objParcelaRec.iNumParcela, "10")
    tRegDetalhe.sVencimento = Format(objParcelaRec.dtDataVencimento, "DDMMYY")
    tRegDetalhe.sValorTitulo = FormataCpoValor(objParcelaRec.dSaldo, 13)
    tRegDetalhe.sNumeroBanco = "230"
    tRegDetalhe.sAgenciaCobradora = FormataCpoNum("", 5)
    tRegDetalhe.sEspecie = FormataCpoNum(gobjCobrancaEletronica.iEspecieTitulo, 2)
    If objParcelaRec.iAceite = 1 Then
        tRegDetalhe.sAceite = "A"
    Else
        tRegDetalhe.sAceite = "N"
    End If
    tRegDetalhe.sDataEmissao = Format(gobjCobrancaEletronica.dtDataEmissaoTitulo, "DDMMYY")
    tRegDetalhe.sInstrucao1 = FormataCpoNum(objOcorrRemParcRem.iInstrucao1, 2)
    tRegDetalhe.sInstrucao2 = FormataCpoNum(objOcorrRemParcRem.iInstrucao2, 2)
    tRegDetalhe.sJurosDiarios = FormataCpoValor(objParcelaRec.dSaldo * gobjCobrancaEletronica.dPercJurosDiario, 13)
    If objParcelaRec.dtDesconto1Ate <> DATA_NULA Then
        tRegDetalhe.sDescontoAte = Format(objParcelaRec.dtDesconto1Ate, "DDMMYY")
    Else
        tRegDetalhe.sDescontoAte = FormataCpoNum("", 6)
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
    
    tRegDetalhe.sValorDesconto = FormataCpoValor(dValorDesconto, 13)
    tRegDetalhe.sValorIOF = FormataCpoValor(0, 13)
    tRegDetalhe.sAbatimento = FormataCpoValor(0, 13)
    If Len(gobjCobrancaEletronica.sCGCSacado) = 14 Then
        tRegDetalhe.sCodigoInscSacado = "02"
    Else
        tRegDetalhe.sCodigoInscSacado = "01"
    End If
    tRegDetalhe.sNumInscricaoSacado = FormataCpoTexto(gobjCobrancaEletronica.sCGCSacado, 14)
    tRegDetalhe.sNome = FormataCpoTexto(gobjCobrancaEletronica.sNomeSacado, 40)
    tRegDetalhe.sEndereco = FormataCpoTexto(Trim(objEndereco.sEndereco), 40)
    tRegDetalhe.sBairro = FormataCpoTexto(Trim(objEndereco.sBairro), 12)
    tRegDetalhe.sCEP = Formata_Cep(objEndereco.sCEP, 8)
    tRegDetalhe.sCidade = FormataCpoTexto(objEndereco.sCidade, 15)
    tRegDetalhe.sEstado = FormataCpoTexto(objEndereco.sSiglaEstado, 2)
    tRegDetalhe.sMensagem = FormataCpoTexto("", 40)
    tRegDetalhe.sBrancos3 = FormataCpoTexto("", 2)
    tRegDetalhe.sCodigoMoeda = "9"
    tRegDetalhe.sNumSequencial = FormataCpoNum(gobjCobrancaEletronica.lSequencialRegistro, 6)
    
    'Incrementa o sequencial de registro
    gobjCobrancaEletronica.lSequencialRegistro = gobjCobrancaEletronica.lSequencialRegistro + 1
    
    Obter_Dados_Detalhe = SUCESSO

    Exit Function
    
Erro_Obter_Dados_Detalhe:

    Obter_Dados_Detalhe = Err
    
    Select Case Err

        Case 51862, 51863, 51864, 32262
       
        Case 51865
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ENDERECO_FILIALCLIENTE_NAO_INFORMADO", Err, gobjCobrancaEletronica.lCliente, gobjCobrancaEletronica.iFilialCli)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147116)

    End Select

    Exit Function

End Function

Private Function Montar_RegistroTrailer_Bandeirantes(sRegTrailer As String, objCobrador As ClassCobrador) As Long
'Monta o registro

Dim tRegTrailerBandeirantes As typeRemTrailerBandeirantes
Dim lErro As Long

On Error GoTo Erro_Montar_RegistroTrailer_Bandeirantes

    'Preenche os dados do regsitro de trailer
    tRegTrailerBandeirantes.sIdentificacaoReg = "9"
    tRegTrailerBandeirantes.sComplemento1 = FormataCpoTexto("", 393)
    tRegTrailerBandeirantes.sSequencialRegistro = FormataCpoNum(gobjCobrancaEletronica.lSequencialRegistro, 6)
    
    'Monta o Registro de trailer
    sRegTrailer = tRegTrailerBandeirantes.sIdentificacaoReg & tRegTrailerBandeirantes.sComplemento1 & tRegTrailerBandeirantes.sSequencialRegistro

    Montar_RegistroTrailer_Bandeirantes = SUCESSO

    Exit Function

Erro_Montar_RegistroTrailer_Bandeirantes:

    Montar_RegistroTrailer_Bandeirantes = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147117)

    End Select

    Exit Function

End Function

Function Processar_ArquivoRetorno_Cobranca_Bandeirantes(objCobrancaEletronica As ClassCobrancaEletronica) As Long
'Faz o processamento do arquivo de retorno

Dim lErro As Long
Dim iIndice As Integer
Dim sRegistro As String
Dim iCodRegistro As Integer
Dim objBorderoRetCobr As New ClassBorderoRetCobr
Dim lComando As Long

On Error GoTo Erro_Processar_ArquivoRetorno_Cobranca_Bandeirantes

    objBorderoRetCobr.iBanco = 230
    
    'le as carteiras de cobranca do cobrador
    lErro = CF("CarteirasCobrador_Le", objCobrancaEletronica.objCobrador, objBorderoRetCobr.colCartCobr)
    If lErro <> SUCESSO Then Error 32265
    
    'Abre os comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 51867

    'Abre o arquivo de retorno
    Open objCobrancaEletronica.sNomeArquivoRetorno For Input As #2
    
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
            If lErro <> SUCESSO Then Error 51869

        'Se o registro for de Detalhe ("Transação")
        ElseIf iCodRegistro = 1 Then
            'Processa os dados contidos no registro de detalhe
            lErro = Processa_Registro_Detalhe(sRegistro, objBorderoRetCobr, objCobrancaEletronica, lComando)
            If lErro <> SUCESSO Then Error 51870
        
        'Se o registro for de Trailer
        ElseIf iCodRegistro = 9 Then
            'Processa os dados do trailer do arquivo
            lErro = Processa_Registro_Trailer(sRegistro)
            If lErro <> SUCESSO Then Error 51871
        Else
            Error 51872
        End If
        
    Loop
    
    Close #2
    
    Call Comando_Fechar(lComando)
    
    'Grava os dados recolhidos durante o processamento do arquivo de retorno
    lErro = CF("BorderoRetCobr_ConfirmaGrava", objBorderoRetCobr)
    If lErro <> SUCESSO Then Error 51873

    Processar_ArquivoRetorno_Cobranca_Bandeirantes = SUCESSO
    
    Exit Function
    
Erro_Processar_ArquivoRetorno_Cobranca_Bandeirantes:

    Processar_ArquivoRetorno_Cobranca_Bandeirantes = Err
    
    Select Case Err
    
        Case 51867, 51868
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 51869, 51870, 51871, 51873
        
        Case 51872
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORMATO_ARQUIVO_INCORRETO", Err, objCobrancaEletronica.sNomeArquivoRetorno)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147118)
    
    End Select

    Close #2
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Private Function Processa_Registro_Header(sRegistro As String, objBorderoRetCobr As ClassBorderoRetCobr, objCobrador As ClassCobrador, sNomeArquivo As String) As Long

Dim lErro As Long
Dim tRegHeaderBandeirantes As typeRetHeaderBandeirantes
Dim objContaCorrente As New ClassContasCorrentesInternas

On Error GoTo Erro_Processa_Registro_Header

    'Extrai da variável lida no arquivo os dados do Header
    lErro = Extrair_Header_Retorno(tRegHeaderBandeirantes, sRegistro)
    If lErro <> SUCESSO Then Error 51875
    
    'Preenche o objBorderoRetCobr
    objBorderoRetCobr.dtDataRecepcao = gdtDataHoje
    objBorderoRetCobr.iCobrador = objCobrador.iCodigo
    objBorderoRetCobr.sNomeArq = sNomeArquivo
    
    'Pega a conta corrente do cobrador
    objContaCorrente.iCodigo = objCobrador.iCodCCI
    
    'Lê a conta corrente do cobrador
    lErro = CF("ContaCorrenteInt_Le", objContaCorrente.iCodigo, objContaCorrente)
    If lErro <> SUCESSO And lErro <> 11807 Then Error 51876
    
    'Se não achou a Conta Corrente --> erro
    If lErro <> SUCESSO Then Error 51877
    
    'Preenche parcialmente o objMovCCI
    objBorderoRetCobr.objMovCCI.iFilialEmpresa = objContaCorrente.iFilialEmpresa
    objBorderoRetCobr.objMovCCI.iCodConta = objContaCorrente.iCodigo
    objBorderoRetCobr.objMovCCI.iTipo = MOVCCI_RECEBIMENTO_TITULO
    objBorderoRetCobr.objMovCCI.iExcluido = NAO_EXCLUIDO
    objBorderoRetCobr.objMovCCI.iTipoMeioPagto = DINHEIRO
    objBorderoRetCobr.objMovCCI.dtDataBaixa = gdtDataHoje
    objBorderoRetCobr.objMovCCI.dtDataMovimento = Extrair_Data(tRegHeaderBandeirantes.sDataEmissaoArq)
    objBorderoRetCobr.objMovCCI.iConciliado = NAO_CONCILIADO
           
    Processa_Registro_Header = SUCESSO
    
    Exit Function

Erro_Processa_Registro_Header:

    Processa_Registro_Header = Err

    Select Case Err
    
        Case 51875, 51876
    
        Case 51877
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_CORRENTE_NAO_ENCONTRADA", Err, objContaCorrente.iCodigo)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147119)
            
    End Select
    
    Exit Function

End Function

Private Function Extrair_Header_Retorno(tRegHeaderBandeirantes As typeRetHeaderBandeirantes, sRegistro As String) As Long
'Extrai do sRegistro os dados do header do arquido de retorno

Dim lErro As Long

On Error GoTo Erro_Extrair_Header_Retorno

    With tRegHeaderBandeirantes
        .sIdentificacaoReg = Mid(sRegistro, 1, 1)
        .sIdentificacaoArq = Mid(sRegistro, 2, 1)
        .sLiteralRetorno = Mid(sRegistro, 3, 7)
        .sCodTipoServico = Mid(sRegistro, 10, 2)
        .sLiteralServico = Mid(sRegistro, 12, 15)
        .sCodEmpresa = Mid(sRegistro, 27, 13)
        .sNomeEmpresa = Mid(sRegistro, 47, 30)
        .sNumeroBanco = Mid(sRegistro, 77, 3)
        .sNomeBanco = Mid(sRegistro, 80, 15)
        .sDataEmissaoArq = Mid(sRegistro, 95, 6)
        .sDensidade = Mid(sRegistro, 101, 5)
        .sLiteralDensidade = Mid(sRegistro, 105, 3)
        .sNumSequencialArq = Mid(sRegistro, 109, 5)
        .sBrancos = Mid(sRegistro, 114, 281)
        .sSequencialRegistro = Mid(sRegistro, 395, 6)
    
    End With

    Extrair_Header_Retorno = SUCESSO

    Exit Function

Erro_Extrair_Header_Retorno:

    Extrair_Header_Retorno = Err
    
    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147120)
            
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

Private Function Processa_Registro_Detalhe(sRegistro As String, objBorderoRetCobr As ClassBorderoRetCobr, objCobrancaEletronica As ClassCobrancaEletronica, lComando As Long) As Long

Dim tRegDetalhe As typeRetDetalheBandeirantes
Dim iCodOcorrencia As Integer, iTipoOcorrencia As Integer
Dim lErro As Long
Dim objParcelaRec As New ClassParcelaReceber
Dim objDetRetCobr As New ClassDetRetCobr
Dim objBorderoRetCobrCart As ClassBorderoRetCobrCart

On Error GoTo Erro_Processa_Registro_Detalhe

    'Extrai da variável lida no arquivo os dados do registro de detalhe
    lErro = Extrair_Detalhe_Retorno(tRegDetalhe, sRegistro)
    If lErro <> SUCESSO Then Error 51878
    
    'Recolhe o código da ocorrência
    iCodOcorrencia = StrParaInt(tRegDetalhe.sCodOcorrencia)
    
    Select Case iCodOcorrencia
        
        'Quando ocorre a baixa
        Case 6, 8, 9, 10
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
    
        Case 51881, 51878, 51879, 51883
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147121)
            
    End Select
    
    Exit Function

End Function

Private Function Extrair_Detalhe_Retorno(tRegDetalhe As typeRetDetalheBandeirantes, sRegistro As String) As Long
'Extrai do sRegistro os dados do Detalhe do arquido de retorno

Dim lErro As Long

On Error GoTo Erro_Extrair_Detalhe_Retorno
    
    With tRegDetalhe

        .sIdentificacaoReg = Mid(sRegistro, 1, 1)
        .sTipoInscricaoEmpresa = Mid(sRegistro, 2, 2)
        .sNumInscricaoEmpresa = Mid(sRegistro, 4, 14)
        .sCodEmpresa = Mid(sRegistro, 18, 13)
        .sBrancos1 = Mid(sRegistro, 30, 7)
        .sUsoDaEmpresa = Mid(sRegistro, 38, 25)
        .sNossoNumero1 = Mid(sRegistro, 63, 13)
        .sBrancos2 = Mid(sRegistro, 76, 7)
        .sUsoDoBanco = Mid(sRegistro, 83, 25)
        .sCodCarteira = Mid(sRegistro, 108, 1)
        .sCodOcorrencia = Mid(sRegistro, 109, 2)
        .sDataOcorrencia = Mid(sRegistro, 111, 6)
        .sSeuNumero = Mid(sRegistro, 117, 10)
        .sNossoNumero2 = Mid(sRegistro, 127, 13)
        .sBrancos3 = Mid(sRegistro, 135, 12)
        .sVencimento = Mid(sRegistro, 147, 6)
        .sValorTitulo = Mid(sRegistro, 153, 13)
        .sNumeroBanco = Mid(sRegistro, 166, 3)
        .sAgenciaCobradora = Mid(sRegistro, 169, 5)
        .sEspecie = Mid(sRegistro, 174, 2)
        .sTarifaCobranca = Mid(sRegistro, 176, 13)
        .sBrancos4 = Mid(sRegistro, 189, 26)
        .sValorIOF = Mid(sRegistro, 215, 13)
        .sOutrasDeducoes = Mid(sRegistro, 228, 13)
        .sDescontos = Mid(sRegistro, 241, 13)
        .sValorPrincipal = Mid(sRegistro, 254, 13)
        .sJuros = Mid(sRegistro, 267, 13)
        .sCodigoMoeda = Mid(sRegistro, 280, 1)
        .sAgenciaRec = Mid(sRegistro, 281, 3)
        .sBrancos5 = Mid(sRegistro, 284, 101)
        .sTipoCobranca = Mid(sRegistro, 385, 1)
        .sDataCredito = Mid(sRegistro, 356, 6)
        .sBrancos6 = Mid(sRegistro, 392, 3)
        .sNumSequencial = Mid(sRegistro, 395, 6)
    End With
    
    Extrair_Detalhe_Retorno = SUCESSO

    Exit Function

Erro_Extrair_Detalhe_Retorno:

    Extrair_Detalhe_Retorno = Err
    
    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147122)
            
    End Select
    
    Exit Function

End Function
'??????????????

Function Processa_Registro_Trailer(sRegistroTrailer As String) As Long

Dim lErro As Long
Dim tRegTrailerBandeirantes As typeRetTrailerBandeirantes

On Error GoTo Erro_Processa_Registro_Trailer

    'Extrai da variável lida no arquivo os dados do registro de detalhe
    lErro = Extrair_Trailer_Retorno(tRegTrailerBandeirantes, sRegistroTrailer)
    If lErro <> SUCESSO Then Error 51884
    
    Processa_Registro_Trailer = SUCESSO
    
    Exit Function
    
Erro_Processa_Registro_Trailer:

    Processa_Registro_Trailer = Err
    
    Select Case Err
    
        Case 51884
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147123)
            
    End Select
    
    Exit Function
    
End Function

Private Function Extrair_Trailer_Retorno(tRegTrailerBandeirantes As typeRetTrailerBandeirantes, sRegistro As String) As Long
'Extrai do sRegistro os dados do Trailer do arquido de retorno

Dim lErro As Long

On Error GoTo Erro_Extrair_Trailer_Retorno

    tRegTrailerBandeirantes.sIdentificacaoReg = Mid(sRegistro, 1, 1)
    tRegTrailerBandeirantes.sCodigoRetorno = Mid(sRegistro, 2, 1)
    tRegTrailerBandeirantes.sCodigoServico = Mid(sRegistro, 3, 2)
    tRegTrailerBandeirantes.sNumeroBanco = Mid(sRegistro, 5, 3)
    tRegTrailerBandeirantes.sBrancos1 = Mid(sRegistro, 8, 10)
    tRegTrailerBandeirantes.sQuantTitulos = Mid(sRegistro, 18, 8)
    tRegTrailerBandeirantes.sValorTotal = Mid(sRegistro, 26, 14)
    tRegTrailerBandeirantes.sAvisoBancario = Mid(sRegistro, 40, 8)
    tRegTrailerBandeirantes.sBrancos2 = Mid(sRegistro, 48, 347)
    tRegTrailerBandeirantes.sSequencialRegistro = Mid(sRegistro, 395, 6)

    Extrair_Trailer_Retorno = SUCESSO

    Exit Function

Erro_Extrair_Trailer_Retorno:

    Extrair_Trailer_Retorno = Err
    
    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147124)
            
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
    
    Exit Function
    
End Function
    
Function Calcula_NossoNumero_Bandeirantes(objCarteiraCobrador As ClassCarteiraCobrador, sNossoNumero As String) As Long

Dim sProxNossoNumero As Long
Dim sDVNossoNumero As String

On Error GoTo Erro_Calcula_NossoNumero_Bandeirantes

    sNossoNumero = objCarteiraCobrador.sFaixaNossoNumeroProx
    
    sProxNossoNumero = sNossoNumero + 1
        
    If Val(sNossoNumero) > Val(objCarteiraCobrador.sFaixaNossoNumeroFinal) Then
        If objCarteiraCobrador.iGeraNossoNumero = BANCO_GERA_NOSSONUMERO Then
            Error 62067
        Else
            sProxNossoNumero = FormataCpoNum(1, 11)
        End If
    End If
    
    sNossoNumero = FormataCpoNum(sNossoNumero, 11)
    
    Call Calcula_DV_NossoNumero(sNossoNumero, sDVNossoNumero, 11, 9)
    
    objCarteiraCobrador.sFaixaNossoNumeroProx = sProxNossoNumero

    Calcula_NossoNumero_Bandeirantes = SUCESSO

    Exit Function
    
Erro_Calcula_NossoNumero_Bandeirantes:

    Calcula_NossoNumero_Bandeirantes = Err
    
    Select Case Err
    
        Case 62067
            Call Rotina_Erro(vbOKOnly, "ERRO_FAIXA_NOSSONUMERO_INSUFICIENTE", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, Err, 147125)
    
    End Select

    Exit Function

End Function

Private Sub Calcula_DV_NossoNumero(sNossoNumero As String, sDVNossoNumero As String, iModulo As Integer, iBase As Integer)

Dim iIndice As Integer
Dim iTamNossoNum As Integer
Dim iDigito1 As Integer
Dim iVarModulo As Integer
Dim iProduto As Integer
Dim iSoma As Integer
Dim iResto As Integer
Dim iResultado As Integer
    
    iVarModulo = 2
    iSoma = 0
    
    For iIndice = iTamNossoNum To 1 Step -1

        iDigito1 = StrParaInt(Mid(sNossoNumero, iIndice, 1))
        iProduto = (iDigito1 * iVarModulo)
        
        iSoma = iSoma + iProduto
        
        iVarModulo = iVarModulo + 1
        
        If iVarModulo > iModulo Then iVarModulo = 2
    
    Next
    
    iResto = iSoma Mod 11
    
    iResultado = 11 - iResto
    
    If iResultado >= 10 Then iResultado = iResultado - 10
    
    sDVNossoNumero = FormataCpoNum(iResultado, 1)
    
    Exit Sub
    
End Sub

Private Sub RegDetCobr_ObterInfo(objDetRetCobr As ClassDetRetCobr, tRegDetalhe As typeRetDetalheBandeirantes)
'copia dados do type para o objeto

    'Recolhe alguns dados do registro de detalhe
    With objDetRetCobr
        .dValorRecebido = Extrair_Valor(tRegDetalhe.sValorPrincipal)
        .dtDataVencimento = Extrair_Data(tRegDetalhe.sVencimento)
        .dtDataCredito = Extrair_Data(tRegDetalhe.sDataCredito)
        .dValorJuros = Extrair_Valor(tRegDetalhe.sJuros)
        .dValorDesconto = Extrair_Valor(tRegDetalhe.sDescontos)
        .sNossoNumero = tRegDetalhe.sNossoNumero1
        .iCodOcorrencia = StrParaInt(tRegDetalhe.sCodOcorrencia)
        .iCodCarteiraNoBanco = StrParaInt(tRegDetalhe.sCodCarteira)
        .dValorTarifa = Extrair_Valor(tRegDetalhe.sTarifaCobranca)
        .sSeuNumero = tRegDetalhe.sSeuNumero
        .sUsoDaEmpresa = tRegDetalhe.sUsoDaEmpresa
        .lNumSequencial = CLng(tRegDetalhe.sNumSequencial)
        .dValorTitulo = Extrair_Valor(tRegDetalhe.sValorTitulo)
        '??? .dValorEntregue = Extrair_Valor(tRegDetalhe.sValorEntregue)
    End With
    
End Sub

Private Function Prepara_CartCobrRet1(objCobrancaEletronica As ClassCobrancaEletronica, objBorderoRetCobr As ClassBorderoRetCobr, objBorderoRetCobrCart As ClassBorderoRetCobrCart, objDetRetCobr As ClassDetRetCobr, tRegDetalhe As typeRetDetalheBandeirantes) As Long
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147126)
            
    End Select

    Exit Function

End Function


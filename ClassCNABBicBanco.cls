VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCNABBicBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim gobjCobrancaEletronica As ClassCobrancaEletronica

Const BICBANCO_INFO_CODEMPRESA = 1

Function Criar_Arquivo_Remessa_BicBanco(objCobrancaEletronica As ClassCobrancaEletronica) As Long
'Cria uma arquivo de remessa com base nos dados do BD

Dim iCancela As Integer
Dim lErro As Long
Dim sRegHeader As String
Dim lTransacao As Long
Dim alComando(0 To 7) As Long
Dim sNomeArquivo As String
Dim iIndice As Integer
Dim sRegDetalhe As String
Dim sRegTrailer As String, lSeqNomeArq As Long
Dim objBorderoCobranca As New ClassBorderoCobranca
Dim objCarteiraCobrador As New ClassCarteiraCobrador
Dim objOcorrRemParc As ClassOcorrRemParcRec

On Error GoTo Erro_Criar_Arquivo_Remessa_BicBanco
    
    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 51754
    Next
        
    'ABre a transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 51847
    
    Set gobjCobrancaEletronica = objCobrancaEletronica
    Set gColFiliaisEmpresa = New Collection
    
    'Gera o nome do arquivo de acordo com o sequencial do Cobrador
    lErro = CF("CobrancaEletronica_Gera_NomeArquivo", objCobrancaEletronica.objCobrador, gobjCobrancaEletronica.sDiretorio, sNomeArquivo, lSeqNomeArq)
    If lErro <> SUCESSO Then gError 51848
    
    gobjCobrancaEletronica.lSeqNomeArq = lSeqNomeArq
    
    'Cria o arquivo de remessa
    Open sNomeArquivo For Output As #1
        
    lErro = Montar_RegistroHeader_BicBanco(sRegHeader, objCobrancaEletronica.objCobrador)
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
            lErro = Montar_RegistroDetalhe_BicBanco(sRegDetalhe, objOcorrRemParc, objCarteiraCobrador)
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
    lErro = Montar_RegistroTrailer_BicBanco(sRegTrailer, objOcorrRemParc)
    If lErro <> SUCESSO Then gError 51856

    'Coloca o registro de trailer no arquivo
    Print #1, DesacentuaTexto(sRegTrailer)
    
    Close #1
    
    'atualizar seq de arquivo de remessa no cobrador
    lErro = Cobrador_Atualizar(objCobrancaEletronica.objCobrador, alComando(6), alComando(7))
    If lErro <> SUCESSO Then gError 81699
    
    lErro = Transacao_Commit()
    If lErro <> SUCESSO Then gError 51857
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Criar_Arquivo_Remessa_BicBanco = SUCESSO
    
    Exit Function

Erro_Criar_Arquivo_Remessa_BicBanco:

    Criar_Arquivo_Remessa_BicBanco = gErr
    
    Select Case gErr
    
        Case 51754
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 51847
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 51848, 51850, 51854, 51851, 51852, 51855, 51849, 51853, 51856, 51857, 79709, 81699
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 147184)
    
    End Select
    
    Close #1

    Call Transacao_Rollback

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Private Function Montar_RegistroHeader_BicBanco(sRegHeader As String, objCobrador As ClassCobrador) As Long
'Monta o registro

Dim tRegHeaderBicBanco As typeRemHeaderBicBanco
Dim lErro As Long

On Error GoTo Erro_Montar_RegistroHeader_BicBanco

    gobjCobrancaEletronica.lSequencialRegistro = 1

    'Carrega tRegHeaderBicBanco com os dados do Header do arquivo
    lErro = Obter_Dados_Header(tRegHeaderBicBanco)
    If lErro <> SUCESSO Then Error 51858
        
    With tRegHeaderBicBanco
    
        'Monta o registro de Header do BicBanco
        sRegHeader = .sIdentificacaoReg & .sIdentificacaoArq & .sLiteralRemessa & .sCodTipoServico & .sLiteralServico & .sCodigoEmpresa
        sRegHeader = sRegHeader & .sBrancos1 & .sAgencia & .sNomeEmpresa & .sNumeroBanco & .sNomeBanco & .sDataGravacao & .sDensidade & .sLiteralDensidade & .sSequencialArq
        sRegHeader = sRegHeader & .sBrancos2 & .sSequencialRegistro
        
    End With
    
    Montar_RegistroHeader_BicBanco = SUCESSO

    Exit Function

Erro_Montar_RegistroHeader_BicBanco:

    Montar_RegistroHeader_BicBanco = Err

    Select Case Err

        Case 51858

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147185)

    End Select

    Exit Function

End Function

Private Function Obter_Dados_Header(tRegHeaderBicBanco As typeRemHeaderBicBanco) As Long

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
    tRegHeaderBicBanco.sIdentificacaoReg = "0"
    tRegHeaderBicBanco.sIdentificacaoArq = "1"
    tRegHeaderBicBanco.sLiteralRemessa = FormataCpoTexto("REMESSA", 7)
    tRegHeaderBicBanco.sCodTipoServico = "01"
    tRegHeaderBicBanco.sLiteralServico = FormataCpoTexto("COBRANCA", 15)
    tRegHeaderBicBanco.sCodigoEmpresa = FormataCpoNum(Obter_Info_Cobrador(gobjCobrancaEletronica.objCobradorCNABInfo.colInformacoes, BICBANCO_INFO_CODEMPRESA), 10)
    tRegHeaderBicBanco.sBrancos1 = FormataCpoTexto("", 7)
    tRegHeaderBicBanco.sAgencia = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sAgencia, 3)
    tRegHeaderBicBanco.sNomeEmpresa = FormataCpoTexto(gsNomeEmpresa, 30)
    tRegHeaderBicBanco.sNumeroBanco = "320"
    tRegHeaderBicBanco.sNomeBanco = FormataCpoTexto("BICBANCO", 15)
    tRegHeaderBicBanco.sDataGravacao = Format(gdtDataHoje, "DDMMYY")
    tRegHeaderBicBanco.sDensidade = "01600"
    tRegHeaderBicBanco.sLiteralDensidade = "BPI"
    tRegHeaderBicBanco.sSequencialArq = FormataCpoNum(gobjCobrancaEletronica.objCobrador.lCNABProxSeqArqCobr, 7)
    tRegHeaderBicBanco.sBrancos2 = FormataCpoTexto("", 279)
    tRegHeaderBicBanco.sSequencialRegistro = FormataCpoNum(gobjCobrancaEletronica.lSequencialRegistro, 6)
    
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147186)
            
    End Select
    
    Exit Function

End Function

Private Function Montar_RegistroDetalhe_BicBanco(sRegDetalhe As String, objOcorrRemParcRec As ClassOcorrRemParcRec, objCarteiraCobrador As ClassCarteiraCobrador) As Long
'Monta o registro de detalhe do arquivo e remessa de títulos a receber

Dim tRegDetalhe As typeRemDetalheBicBanco
Dim lErro As Long

On Error GoTo Erro_Montar_RegistroDetalhe_BicBanco
    
    'Obtem os dados do registro de detalhe
    lErro = Obter_Dados_Detalhe(tRegDetalhe, objOcorrRemParcRec, objCarteiraCobrador)
    If lErro <> SUCESSO And lErro <> 32262 Then Error 51861
    If lErro <> SUCESSO Then Error 32263
    
    With tRegDetalhe
        'Forma o registro de detalhe
        sRegDetalhe = .sIdentificacaoReg & .sTipoInscricaoEmpSac & .sNumInscricaoEmpSac & .sIdentificacaoEmp & .sBrancos1 & FormataCpoTexto("", 25) & .sNossoNumero & .sBrancos2 & .sDiasDevolucao & .sUsoDoBanco
        sRegDetalhe = sRegDetalhe & .sCodCarteira & .sCodOcorrencia & .sSeuNumero & .sVencimento & .sValorTitulo & .sNumeroBanco & .sAgenciaCobradora & .sEspecie & .sAceite & .sDataEmissao & .sInstrucao1
        sRegDetalhe = sRegDetalhe & .sInstrucao2 & .sJurosDiarios & .sDescontoAte & .sValorDesconto & .sValorIOF & .sAbatimento & .sCodigoInscSacado & .sNumInscricaoSacado & .sNome
        sRegDetalhe = sRegDetalhe & .sEndereco & .sBrancosCNAB2 & .sBairro & .sCEP & .sCidade & .sEstado & .sMensagem & "05" & .sCodigoMoeda & .sNumSequencial
    End With
    
    Montar_RegistroDetalhe_BicBanco = SUCESSO

    Exit Function

Erro_Montar_RegistroDetalhe_BicBanco:

    Montar_RegistroDetalhe_BicBanco = Err

    Select Case Err
    
        Case 51861, 32263

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147187)

    End Select

    Exit Function

End Function

Private Function Obter_Dados_Detalhe(tRegDetalhe As typeRemDetalheBicBanco, objOcorrRemParcRem As ClassOcorrRemParcRec, objCarteiraCobrador As ClassCarteiraCobrador) As Long
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
            
    tRegDetalhe.sIdentificacaoReg = "1"
    If Len(gobjCobrancaEletronica.sCGCSacado) = 14 Then
        tRegDetalhe.sTipoInscricaoEmpSac = "02"
    Else
        tRegDetalhe.sTipoInscricaoEmpSac = "01"
    End If
    tRegDetalhe.sNumInscricaoEmpSac = FormataCpoTexto(gobjCobrancaEletronica.sCGCCedente, 14)
    tRegDetalhe.sIdentificacaoEmp = FormataCpoNum(Obter_Info_Cobrador(gobjCobrancaEletronica.objCobradorCNABInfo.colInformacoes, BICBANCO_INFO_CODEMPRESA), 10)
    tRegDetalhe.sBrancos1 = FormataCpoTexto("", 10)
    tRegDetalhe.sUsoDaEmpresa = FormataCpoTexto(objOcorrRemParcRem.lNumIntDoc, 25)
    tRegDetalhe.sNossoNumero = FormataCpoNum("", 8)
    tRegDetalhe.sBrancos2 = FormataCpoTexto("", 37)
        
    If objCarteiraCobrador.iCodCarteiraCobranca = CARTEIRA_SIMPLES Then
        tRegDetalhe.sCodCarteira = FormataCpoTexto(1, 1)
    ElseIf objCarteiraCobrador.iCodCarteiraCobranca = CARTEIRA_CAUCIONADA Then
        tRegDetalhe.sCodCarteira = FormataCpoTexto(4, 1)
    Else
        tRegDetalhe.sCodCarteira = FormataCpoTexto(objCarteiraCobrador.sCodCarteiraNoBanco, 1)
    End If
    
    tRegDetalhe.sCodOcorrencia = FormataCpoNum(objOcorrRemParcRem.iCodOcorrencia, 2)
    tRegDetalhe.sSeuNumero = FormataCpoTexto(gobjCobrancaEletronica.lNumTitulo & "/" & objParcelaRec.iNumParcela, "10")
    tRegDetalhe.sVencimento = Format(objParcelaRec.dtDataVencimento, "DDMMYY")
    tRegDetalhe.sValorTitulo = FormataCpoValor(objParcelaRec.dSaldo, 13)
    tRegDetalhe.sNumeroBanco = FormataCpoNum(gobjCobrancaEletronica.objCobrador.iCodBanco, 3)
    tRegDetalhe.sAgenciaCobradora = FormataCpoNum("", 5)
    tRegDetalhe.sEspecie = FormataCpoNum(1, 2)
    If objParcelaRec.iAceite = 1 Then
        tRegDetalhe.sAceite = "A"
    Else
        tRegDetalhe.sAceite = "N"
    End If
    tRegDetalhe.sDataEmissao = Format(gobjCobrancaEletronica.dtDataEmissaoTitulo, "DDMMYY")
    tRegDetalhe.sInstrucao1 = FormataCpoNum(7, 2)
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
    tRegDetalhe.sEndereco = FormataCpoTexto(Trim(objEndereco.sEndereco), 34)
    tRegDetalhe.sBrancosCNAB2 = FormataCpoTexto("", 6)
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147188)

    End Select

    Exit Function

End Function

Private Function Montar_RegistroTrailer_BicBanco(sRegTrailer As String, objCobrador As ClassCobrador) As Long
'Monta o registro

Dim tRegTrailerBicBanco As typeRemTrailerBicBanco
Dim lErro As Long

On Error GoTo Erro_Montar_RegistroTrailer_BicBanco

    'Preenche os dados do regsitro de trailer
    tRegTrailerBicBanco.sIdentificacaoReg = "9"
    tRegTrailerBicBanco.sComplemento1 = FormataCpoTexto("", 393)
    tRegTrailerBicBanco.sSequencialRegistro = FormataCpoNum(gobjCobrancaEletronica.lSequencialRegistro, 6)
    
    'Monta o Registro de trailer
    sRegTrailer = tRegTrailerBicBanco.sIdentificacaoReg & tRegTrailerBicBanco.sComplemento1 & tRegTrailerBicBanco.sSequencialRegistro

    Montar_RegistroTrailer_BicBanco = SUCESSO

    Exit Function

Erro_Montar_RegistroTrailer_BicBanco:

    Montar_RegistroTrailer_BicBanco = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147189)

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


Function Processar_ArquivoRetorno_Cobranca_BicBanco(objCobrancaEletronica As ClassCobrancaEletronica) As Long
'Faz o processamento do arquivo de retorno

Dim lErro As Long
Dim iIndice As Integer
Dim sRegistro As String
Dim iCodRegistro As Integer
Dim objBorderoRetCobr As New ClassBorderoRetCobr
Dim lComando As Long
Dim lComando1 As Long

On Error GoTo Erro_Processar_ArquivoRetorno_Cobranca_BicBanco

    Set gobjCobrancaEletronica = objCobrancaEletronica
    objBorderoRetCobr.iBanco = 320
    
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

    Processar_ArquivoRetorno_Cobranca_BicBanco = SUCESSO
    
    Exit Function
    
Erro_Processar_ArquivoRetorno_Cobranca_BicBanco:

    Processar_ArquivoRetorno_Cobranca_BicBanco = Err
    
    Select Case Err
    
        Case 51867, 51868
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 51869, 51870, 51871, 51873
        
        Case 51872
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORMATO_ARQUIVO_INCORRETO", Err, objCobrancaEletronica.sNomeArquivoRetorno)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147190)
    
    End Select

    Close #2
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Private Function Processa_Registro_Header(sRegistro As String, objBorderoRetCobr As ClassBorderoRetCobr, objCobrador As ClassCobrador, sNomeArquivo As String) As Long

Dim lErro As Long
Dim tRegHeaderBicBanco As typeRetHeaderBicBanco
Dim objContaCorrente As New ClassContasCorrentesInternas

On Error GoTo Erro_Processa_Registro_Header

    'Extrai da variável lida no arquivo os dados do Header
    lErro = Extrair_Header_Retorno(tRegHeaderBicBanco, sRegistro)
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
    objBorderoRetCobr.objMovCCI.dtDataMovimento = Extrair_Data(tRegHeaderBicBanco.sDataEmissaoArq)
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147191)
            
    End Select
    
    Exit Function

End Function

Private Function Extrair_Header_Retorno(tRegHeaderBicBanco As typeRetHeaderBicBanco, sRegistro As String) As Long
'Extrai do sRegistro os dados do header do arquido de retorno

Dim lErro As Long

On Error GoTo Erro_Extrair_Header_Retorno

    With tRegHeaderBicBanco
        .sIdentificacaoReg = Mid(sRegistro, 1, 1)
        .sIdentificacaoArq = Mid(sRegistro, 2, 1)
        .sLiteralRetorno = Mid(sRegistro, 3, 7)
        .sCodTipoServico = Mid(sRegistro, 10, 2)
        .sLiteralServico = Mid(sRegistro, 12, 15)
        .sCodEmpresa = Mid(sRegistro, 27, 10)
        .sNomeEmpresa = Mid(sRegistro, 47, 30)
        .sNumeroBanco = Mid(sRegistro, 77, 3)
        .sNomeBanco = Mid(sRegistro, 80, 15)
        .sDataEmissaoArq = Mid(sRegistro, 95, 6)
        .sDensidade = Mid(sRegistro, 101, 5)
        .sLiteralDensidade = Mid(sRegistro, 106, 3)
        .sNumSequencialArq = Mid(sRegistro, 109, 7)
        .sBrancos = Mid(sRegistro, 116, 279)
        .sSequencialRegistro = Mid(sRegistro, 395, 6)
    
    End With

    Extrair_Header_Retorno = SUCESSO

    Exit Function

Erro_Extrair_Header_Retorno:

    Extrair_Header_Retorno = Err
    
    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147192)
            
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

Dim tRegDetalhe As typeRetDetalheBicBanco
Dim iCodOcorrencia As Integer, iTipoOcorrencia As Integer
Dim iCodCarteiraNoBanco As Integer
Dim lErro As Long
Dim objBorderoRetCobrCart As New ClassBorderoRetCobrCart
Dim objDetRetCobr As New ClassDetRetCobr

On Error GoTo Erro_Processa_Registro_Detalhe

    'Extrai da variável lida no arquivo os dados do registro de detalhe
    lErro = Extrair_Detalhe_Retorno(tRegDetalhe, sRegistro)
    If lErro <> SUCESSO Then Error 51878
    
    'Recolhe o código da ocorrência
    iCodOcorrencia = StrParaInt(tRegDetalhe.sCodOcorrencia)
    
    Select Case iCodOcorrencia
        
        'Quando ocorre a baixa
        Case 6, 7, 8, 9, 10, 15, 19
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
    
        Case 51878, 51883
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147193)
            
    End Select
    
    Exit Function

End Function

Private Function Extrair_Detalhe_Retorno(tRegDetalhe As typeRetDetalheBicBanco, sRegistro As String) As Long
'Extrai do sRegistro os dados do Detalhe do arquido de retorno

Dim lErro As Long

On Error GoTo Erro_Extrair_Detalhe_Retorno
    
    With tRegDetalhe

        .sIdentificacaoReg = Mid(sRegistro, 1, 1)
        .sTipoInscricaoEmpresa = Mid(sRegistro, 2, 2)
        .sNumInscricaoEmpresa = Mid(sRegistro, 4, 14)
        .sCodEmpresa = Mid(sRegistro, 18, 10)
        .sBrancos1 = Mid(sRegistro, 20, 10)
        .sUsoDaEmpresa = Mid(sRegistro, 38, 25)
        .sNossoNumero1 = Mid(sRegistro, 63, 8)
        .sBrancos2 = Mid(sRegistro, 76, 37)
        .sCodCarteira = Mid(sRegistro, 108, 1)
        .sCodOcorrencia = Mid(sRegistro, 109, 2)
        .sDataOcorrencia = Mid(sRegistro, 111, 6)
        .sSeuNumero = Mid(sRegistro, 117, 10)
        .sNossoNumero2 = Mid(sRegistro, 127, 8)
        .sCodRejeicao = Mid(sRegistro, 135, 2)
        .sBrancos3 = Mid(sRegistro, 137, 10)
        .sVencimento = Mid(sRegistro, 147, 6)
        .sValorTitulo = Mid(sRegistro, 153, 13)
        .sNumeroBanco = Mid(sRegistro, 166, 3)
        .sAgenciaCobradora = Mid(sRegistro, 169, 5)
        .sEspecie = Mid(sRegistro, 174, 2)
        .sTarifaCobranca = Mid(sRegistro, 176, 13)
        .sOutrasDespesas = Mid(sRegistro, 189, 13)
        .sDescontos = Mid(sRegistro, 202, 13)
        .sValorIOF = Mid(sRegistro, 215, 13)
        .sOutrasDeducoes = Mid(sRegistro, 228, 13)
        .sValorPrincipal = Mid(sRegistro, 254, 13)
        .sJuros = Mid(sRegistro, 267, 13)
        .sOutrosCreditos = Mid(sRegistro, 280, 13)
        .sDataCredito = Mid(sRegistro, 293, 6)
        .sValorLiquido = Mid(sRegistro, 299, 13)
        .sNomeSacado = Mid(sRegistro, 312, 40)
        .sMoeda = Mid(sRegistro, 394, 1)
        .sNumSequencial = Mid(sRegistro, 395, 6)
    End With
    
    Extrair_Detalhe_Retorno = SUCESSO

    Exit Function

Erro_Extrair_Detalhe_Retorno:

    Extrair_Detalhe_Retorno = Err
    
    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147194)
            
    End Select
    
    Exit Function

End Function

Function Processa_Registro_Trailer(sRegistroTrailer As String) As Long

Dim lErro As Long
Dim tRegTrailerBicBanco As typeRetTrailerBicBanco

On Error GoTo Erro_Processa_Registro_Trailer

    'Extrai da variável lida no arquivo os dados do registro de detalhe
    lErro = Extrair_Trailer_Retorno(tRegTrailerBicBanco, sRegistroTrailer)
    If lErro <> SUCESSO Then Error 51884
    
    Processa_Registro_Trailer = SUCESSO
    
    Exit Function
    
Erro_Processa_Registro_Trailer:

    Processa_Registro_Trailer = Err
    
    Select Case Err
    
        Case 51884
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147195)
            
    End Select
    
    Exit Function
    
End Function

Private Function Extrair_Trailer_Retorno(tRegTrailerBicBanco As typeRetTrailerBicBanco, sRegistro As String) As Long
'Extrai do sRegistro os dados do Trailer do arquido de retorno

Dim lErro As Long

On Error GoTo Erro_Extrair_Trailer_Retorno

    tRegTrailerBicBanco.sIdentificacaoReg = Mid(sRegistro, 1, 1)
    tRegTrailerBicBanco.sCodigoRetorno = Mid(sRegistro, 2, 1)
    tRegTrailerBicBanco.sCodigoServico = Mid(sRegistro, 3, 2)
    tRegTrailerBicBanco.sNumeroBanco = Mid(sRegistro, 5, 3)
    tRegTrailerBicBanco.sBrancos1 = Mid(sRegistro, 8, 10)
    tRegTrailerBicBanco.sQuantTitulos = Mid(sRegistro, 18, 8)
    tRegTrailerBicBanco.sValorTotal = Mid(sRegistro, 26, 14)
    tRegTrailerBicBanco.sAvisoBancario = Mid(sRegistro, 40, 8)
    tRegTrailerBicBanco.sDataCredito = Mid(sRegistro, 48, 6)
    tRegTrailerBicBanco.sBrancos2 = Mid(sRegistro, 54, 341)
    tRegTrailerBicBanco.sSequencialRegistro = Mid(sRegistro, 395, 6)

    Extrair_Trailer_Retorno = SUCESSO

    Exit Function

Erro_Extrair_Trailer_Retorno:

    Extrair_Trailer_Retorno = Err
    
    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147196)
            
    End Select
    
    Exit Function

End Function

Private Function Prepara_CartCobrRet1(objCobrancaEletronica As ClassCobrancaEletronica, objBorderoRetCobr As ClassBorderoRetCobr, objBorderoRetCobrCart As ClassBorderoRetCobrCart, objDetRetCobr As ClassDetRetCobr, tRegDetalhe As typeRetDetalheBicBanco) As Long
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147197)
            
    End Select

    Exit Function

End Function

Private Sub RegDetCobr_ObterInfo(objDetRetCobr As ClassDetRetCobr, tRegDetalhe As typeRetDetalheBicBanco)
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
        
    End With
    
End Sub

Private Function Cobrador_Atualizar(objCobrador As ClassCobrador, lComando1 As Long, lComando2 As Long) As Long
'Atualiza o sequencial do proximo arquivo remessa a ser gerado

Dim lErro As Long, lCNABProxSeqArqCobr As Long

On Error GoTo Erro_Cobrador_Atualizar

    lErro = Comando_ExecutarPos(lComando1, "SELECT CNABProxSeqArqCobr FROM Cobradores WHERE Codigo = ?", 0, lCNABProxSeqArqCobr, objCobrador.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 81700
    
    lErro = Comando_BuscarProximo(lComando1)
    If lErro <> AD_SQL_SUCESSO Then gError 81701
    
    lCNABProxSeqArqCobr = lCNABProxSeqArqCobr + 1
    
    lErro = Comando_ExecutarPos(lComando2, "UPDATE Cobradores SET CNABProxSeqArqCobr = ?", lComando1, lCNABProxSeqArqCobr)
    If lErro <> AD_SQL_SUCESSO Then gError 81702
    
    Cobrador_Atualizar = SUCESSO
     
    Exit Function
    
Erro_Cobrador_Atualizar:

    Cobrador_Atualizar = gErr
     
    Select Case gErr
          
        Case 81700, 81701
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COBRADOR", gErr, objCobrador.iCodigo)
        
        Case 81702
            Call Rotina_Erro(vbOKOnly, "ERRO_MODIFICACAO_COBRADOR", gErr, objCobrador.iCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147198)
     
    End Select
     
    Exit Function

End Function


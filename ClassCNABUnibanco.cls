VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCNABUnibanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Dim gobjCobrancaEletronica As ClassCobrancaEletronica
Public gobjGeracaoArqCNABPag As ClassGeracaoArqCNABPag

Function Criar_Arquivo_Remessa_Unibanco(objCobrancaEletronica As ClassCobrancaEletronica) As Long
'Cria uma arquivo de remessa com base nos dados do BD

Dim lErro As Long
Dim lTransacao As Long
Dim alComando(0 To 7) As Long
Dim iIndice As Integer
Dim colBorderosComReg As New Collection
Dim colBorderosSemReg As New Collection

On Error GoTo Erro_Criar_Arquivo_Remessa_Unibanco
    
    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 51885
    Next
        
    'ABre a transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 51886
    
    Set gobjCobrancaEletronica = objCobrancaEletronica
    Set gColFiliaisEmpresa = New Collection
    
    'Separa os borderos com e sem registro
    lErro = Separa_Borderos_ComSem_Registro(gobjCobrancaEletronica.colBorderos, colBorderosComReg, colBorderosSemReg)
    If lErro <> SUCESSO Then Error 51887
    
    If colBorderosComReg.Count > 0 Then
    
        Set gobjCobrancaEletronica.colBorderos = colBorderosComReg
    
        lErro = Gera_Arquivo_CobrancaComRegistro(alComando(0), alComando(1), alComando(2), alComando(3))
        If lErro <> SUCESSO Then Error 51888
    
    End If
    
    If colBorderosSemReg.Count > 0 Then
        
        Set gobjCobrancaEletronica.colBorderos = colBorderosSemReg
        
        lErro = Gera_Arquivo_CobrancaSemRegistro(alComando(4), alComando(5), alComando(6), alComando(7))
        If lErro <> SUCESSO Then Error 51889
    
    End If
    
    Close #1
    
    lErro = Transacao_Commit()
    If lErro <> SUCESSO Then Error 51890
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Criar_Arquivo_Remessa_Unibanco = SUCESSO
    
    Exit Function

Erro_Criar_Arquivo_Remessa_Unibanco:

    Criar_Arquivo_Remessa_Unibanco = Err
    
    Select Case Err
    
        Case 51885
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 51886
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
        
        Case 51887, 51888, 51889
        
        Case 51890
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 147397)
    
    End Select

    Call gobjCobrancaEletronica.objTelaAtualizacao.Mostra_Evolucao(True, 1)
    
    Close #1

    Call Transacao_Rollback

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Private Function Montar_RegistroHeader_ComReg_Unibanco(sRegHeader As String) As Long
'Monta o registro

Dim tRegHeaderUnibanco As typeRemHeaderUnibanco
Dim lErro As Long

On Error GoTo Erro_Montar_RegistroHeader_ComReg_Unibanco

    gobjCobrancaEletronica.lSequencialRegistro = 1

    'Carrega tRegHeaderUnibanco com os dados do Header do arquivo
    lErro = Obter_Dados_Header(tRegHeaderUnibanco)
    If lErro <> SUCESSO Then Error 51891
    
    'Monta o registro de Header do Unibanco
    With tRegHeaderUnibanco
        sRegHeader = .sIdentificacaoReg & .sIdentificacaoArq & .sLiteralRemessa & .sCodTipoServico & .sLiteralServico & .sBrancos1
        sRegHeader = sRegHeader & .sAgCredito & .sContaCredito & .sDVContaCredito & .sZeros1 & .sNomeEmpresa & .sNumeroBanco & .sNomeBanco & .sBrancos2 & .sDataEmissaoArq
        sRegHeader = sRegHeader & .sDensidade & .sLiteralDensidade & .sZeros2 & .sSequencialRegistro
    End With
    
    Montar_RegistroHeader_ComReg_Unibanco = SUCESSO

    Exit Function

Erro_Montar_RegistroHeader_ComReg_Unibanco:

    Montar_RegistroHeader_ComReg_Unibanco = Err

    Select Case Err

        Case 51891

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147398)

    End Select

    Exit Function

End Function

Private Function Obter_Dados_Header(tRegHeaderUnibanco As typeRemHeaderUnibanco) As Long

Dim lErro As Long
Dim objCCI As New ClassContasCorrentesInternas

On Error GoTo Erro_Obter_Dados_Header
        
    'Lê a conta corrente associada ao cobrador
    lErro = CF("ContaCorrenteInt_Le", gobjCobrancaEletronica.objCobrador.iCodCCI, objCCI)
    If lErro <> SUCESSO And lErro <> 11807 Then Error 51892
    If lErro <> SUCESSO Then Error 51893
    
    If objCCI.iCodBanco <> gobjCobrancaEletronica.objCobrador.iCodBanco Then Error 62013
    If Len(Trim(objCCI.sAgencia)) = 0 Or Len(Trim(objCCI.sNumConta)) = 0 Then Error 62014
        
    Set gobjCobrancaEletronica.objCCI = objCCI
        
    'Carrega os campos do cabeçalho
    tRegHeaderUnibanco.sIdentificacaoReg = "0"
    tRegHeaderUnibanco.sIdentificacaoArq = "1"
    tRegHeaderUnibanco.sLiteralRemessa = FormataCpoTexto("REMESSA", 7)
    tRegHeaderUnibanco.sCodTipoServico = "01"
    tRegHeaderUnibanco.sLiteralServico = FormataCpoTexto("COBRANCA", 8)
    tRegHeaderUnibanco.sBrancos1 = FormataCpoTexto("", 7)
    tRegHeaderUnibanco.sAgCredito = FormataCpoNum(objCCI.sAgencia, 4)
    tRegHeaderUnibanco.sContaCredito = FormataCpoNum(objCCI.sNumConta, 6)
    tRegHeaderUnibanco.sDVContaCredito = FormataCpoNum(objCCI.sDVNumConta, 1)
    tRegHeaderUnibanco.sZeros1 = FormataCpoNum("", 9)
    tRegHeaderUnibanco.sNomeEmpresa = FormataCpoTexto(gsNomeEmpresa, 30)
    tRegHeaderUnibanco.sNumeroBanco = "409"
    tRegHeaderUnibanco.sNomeBanco = "UNIBANCO"
    tRegHeaderUnibanco.sBrancos2 = FormataCpoTexto("", 7)
    tRegHeaderUnibanco.sDataEmissaoArq = Format(gdtDataHoje, "DDMMYY")
    tRegHeaderUnibanco.sDensidade = "01600"
    tRegHeaderUnibanco.sLiteralDensidade = "BPI"
    tRegHeaderUnibanco.sZeros2 = FormataCpoNum("", 286)
    tRegHeaderUnibanco.sSequencialRegistro = FormataCpoNum(gobjCobrancaEletronica.lSequencialRegistro, 6)
    
    'Incrementa o seuqencial do arquivo
    gobjCobrancaEletronica.lSequencialRegistro = gobjCobrancaEletronica.lSequencialRegistro + 1

    Obter_Dados_Header = SUCESSO

    Exit Function
    
Erro_Obter_Dados_Header:

    Obter_Dados_Header = Err
    
    Select Case Err
    
        Case 51892
        
        Case 51893
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_COBRADOR_NAO_ENCONTRADA", Err, gobjCobrancaEletronica.iCobrador)
        
        Case 62013
            lErro = Rotina_Erro(vbOKOnly, "ERRO_BANCO_CCI_DIFERENTE_COBRADOR", Err, objCCI.iCodBanco, gobjCobrancaEletronica.objCobrador.iCodBanco)
    
        Case 62014
            lErro = Rotina_Erro(vbOKOnly, "ERRO_AGENCIA_CONTA_COBRADOR_NAO_PREENCHIDAS", Err)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147399)
            
    End Select
    
    Exit Function

End Function

Private Function Montar_RegistroDetalhe_ComReg_Unibanco(sRegDetalhe As String, objOcorrRemParcRec As ClassOcorrRemParcRec, objCarteiraCobrador As ClassCarteiraCobrador) As Long

'Monta o registro

Dim tRegDetalhe As typeRemDetalheUnibanco
Dim lErro As Long

On Error GoTo Erro_Montar_RegistroDetalhe_ComReg_Unibanco
    
    'Obtem os dados do registro de detalhe
    lErro = Obter_Dados_Detalhe(tRegDetalhe, objOcorrRemParcRec, objCarteiraCobrador)
    If lErro <> SUCESSO And lErro <> 32267 Then Error 51894
    If lErro <> SUCESSO Then Error 32270
    
    With tRegDetalhe
        'Forma o registro de detalhe
        sRegDetalhe = .sIdentificacaoReg & .sCodigoInscricao & .sNumeroInscricao & .sAgenciaDebito & .sContaCorrente & .sDigitoContaCorrente
        sRegDetalhe = sRegDetalhe & .sZeros1 & .sUsoEmpresa & .sNossoNumero & .sDVNossoNumero & .sMensagem & .sMoeda & .sCarteira & .sTipoTransacao
        sRegDetalhe = sRegDetalhe & .sSeuNumero & .sDataVencto & .sValorTitulo & .sBancoCobranca & .sZeros2 & .sAgenciaCobradora & .sEspecie & .sAceite
        sRegDetalhe = sRegDetalhe & .sDataEmissao & .sInstrucao1 & .sInstrucao2 & .sValorJuros & .sDataDesconto & .sValorDesconto & .sZeros3 & .sAbatimento
        sRegDetalhe = sRegDetalhe & .sIdentfInscSacado & .sNumInscricaoSacado & .sNomeSacado & .sZeros4 & .sEnderecoSacado & .sEnderecoCompl & .sBairro
        sRegDetalhe = sRegDetalhe & .sCEP & .sCidade & .sEstado & .sSacadorAvalista & .sZeros5 & .sPrazoProtesto & .sZeros6 & .sNumSequencialRegistro
    End With
    
    Montar_RegistroDetalhe_ComReg_Unibanco = SUCESSO

    Exit Function

Erro_Montar_RegistroDetalhe_ComReg_Unibanco:

    Montar_RegistroDetalhe_ComReg_Unibanco = Err

    Select Case Err
    
        Case 51894, 32270

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147400)

    End Select

    Exit Function

End Function

Private Function Obter_Dados_Detalhe(tRegDetalhe As typeRemDetalheUnibanco, objOcorrRemParcRem As ClassOcorrRemParcRec, objCarteiraCobrador As ClassCarteiraCobrador) As Long
'Obtem os dados que são utilizados para preencher o registro de detalhe

Dim lErro As Long
Dim objParcelaRec As New ClassParcelaReceber
Dim objEndereco As New ClassEndereco
Dim sNossoNumero As String
Dim sDVNossoNumero As String
Dim dValorDesconto As Double

On Error GoTo Erro_Obter_Dados_Detalhe

    objParcelaRec.lNumIntDoc = objOcorrRemParcRem.lNumIntParc

    'Busca os dados da parcela a receber no BD
    lErro = CF("CobrancaEletronica_Obter_DadosParcelaRec", objParcelaRec, gobjCobrancaEletronica)
    If lErro <> SUCESSO And lErro <> 51674 Then Error 51895
    If lErro <> SUCESSO Then Error 32267
    
    lErro = CF("CobrancaEletronica_Obter_Endereco", gobjCobrancaEletronica, objEndereco)
    If lErro <> SUCESSO Then Error 51897
        
    lErro = Calcula_NossoNumero_Unibanco(objCarteiraCobrador, sNossoNumero, sDVNossoNumero)
    If lErro <> SUCESSO Then Error 51898
    
    With tRegDetalhe
        .sIdentificacaoReg = "1"
        .sCodigoInscricao = "02"
        .sNumeroInscricao = FormataCpoNum(gobjCobrancaEletronica.sCGCCedente, 14)
        .sAgenciaDebito = FormataCpoNum(gobjCobrancaEletronica.objCCI.sAgencia, 4)
        .sContaCorrente = FormataCpoNum(gobjCobrancaEletronica.objCCI.sNumConta, 6)
        .sDigitoContaCorrente = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVNumConta, 1)
        .sZeros1 = FormataCpoNum("", 9)
        .sUsoEmpresa = FormataCpoTexto(objOcorrRemParcRem.lNumIntDoc, 25)
        .sNossoNumero = FormataCpoNum(sNossoNumero, 10)
        .sDVNossoNumero = FormataCpoNum(sDVNossoNumero, 1)
        .sMensagem = FormataCpoTexto("", 30)
        .sMoeda = FormataCpoTexto("", 4)
        If objCarteiraCobrador.iCodCarteiraCobranca = CARTEIRA_SIMPLES Then
            .sCarteira = "1"
        ElseIf objCarteiraCobrador.iCodCarteiraCobranca = CARTEIRA_CAUCIONADA Then
            .sCarteira = "6"
        ElseIf objCarteiraCobrador.iCodCarteiraCobranca = CARTEIRA_DESCONTADA Then
            .sCarteira = "7"
        End If
        
        .sTipoTransacao = FormataCpoNum(objOcorrRemParcRem.iCodOcorrencia, 2)
        .sSeuNumero = FormataCpoTexto(gobjCobrancaEletronica.lNumTitulo & "/" & objParcelaRec.iNumParcela, "10")
        .sDataVencto = Format(objParcelaRec.dtDataVencimento, "DDMMYY")
        .sValorTitulo = FormataCpoValor(objParcelaRec.dSaldo, 13)
        .sBancoCobranca = "409"
        .sZeros2 = FormataCpoNum("", 1)
        .sAgenciaCobradora = FormataCpoNum("", 4)
        .sEspecie = FormataCpoNum(gobjCobrancaEletronica.iEspecieTitulo, 2)
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
        .sZeros3 = FormataCpoValor(0, 13)
        .sAbatimento = FormataCpoValor(0, 13)
        If Len(gobjCobrancaEletronica.sCGCSacado) = 14 Then
            .sIdentfInscSacado = "02"
        Else
            .sIdentfInscSacado = "01"
        End If
        .sNumInscricaoSacado = FormataCpoNum(gobjCobrancaEletronica.sCGCSacado, 14)
        .sNomeSacado = FormataCpoTexto(gobjCobrancaEletronica.sNomeSacado, 30)
        .sZeros4 = FormataCpoNum("", 10)
        .sEnderecoSacado = FormataCpoTexto(Trim(objEndereco.sEndereco), 40)
        .sEnderecoCompl = ""
        .sBairro = FormataCpoTexto(objEndereco.sBairro, 12)
        .sCEP = Formata_Cep(objEndereco.sCEP, 8)
        .sCidade = FormataCpoTexto(objEndereco.sCidade, 15)
        .sEstado = FormataCpoTexto(objEndereco.sSiglaEstado, 2)
        .sSacadorAvalista = FormataCpoTexto("", 2)
        If objParcelaRec.dtDesconto2Ate <> DATA_NULA Then
            .sSacadorAvalista = .sSacadorAvalista & Format(objParcelaRec.dtDesconto2Ate, "DDMMYY")
            'Calcula o valor do desconto quando for do tipo percentual
            dValorDesconto = 0
            If objParcelaRec.iDesconto2Codigo = PERC_ANT_DIA Or objParcelaRec.iDesconto2Codigo = PERC_ANT_DIA_UTIL Or objParcelaRec.iDesconto2Codigo = Percentual Then
                dValorDesconto = StrParaDbl(Format(objParcelaRec.dDesconto2Valor * objParcelaRec.dSaldo, "Fixed"))
            Else
                If objParcelaRec.iDesconto2Codigo = VALOR_FIXO Then
                    dValorDesconto = objParcelaRec.dDesconto2Valor
                End If
            End If
            
            .sSacadorAvalista = .sSacadorAvalista & FormataCpoValor(dValorDesconto, 13)
            
            If objParcelaRec.dtDesconto3Ate <> DATA_NULA Then
                .sSacadorAvalista = .sSacadorAvalista & Format(objParcelaRec.dtDesconto3Ate, "DDMMYY")
                'Calcula o valor do desconto quando for do tipo percentual
                dValorDesconto = 0
                If objParcelaRec.iDesconto3Codigo = PERC_ANT_DIA Or objParcelaRec.iDesconto3Codigo = PERC_ANT_DIA_UTIL Or objParcelaRec.iDesconto3Codigo = Percentual Then
                    dValorDesconto = StrParaDbl(Format(objParcelaRec.dDesconto3Valor * objParcelaRec.dSaldo, "Fixed"))
                Else
                    If objParcelaRec.iDesconto3Codigo = VALOR_FIXO Then
                        dValorDesconto = objParcelaRec.dDesconto3Valor
                    End If
                End If
                
                .sSacadorAvalista = .sSacadorAvalista & FormataCpoValor(dValorDesconto, 13)
            Else
                .sSacadorAvalista = .sSacadorAvalista & FormataCpoTexto("", 19)
            End If
        Else
            .sSacadorAvalista = .sSacadorAvalista & FormataCpoTexto("", 38)
        End If
        
        If objOcorrRemParcRem.iInstrucao1 = 6 Then
            .sPrazoProtesto = FormataCpoNum(objOcorrRemParcRem.iDiasDeProtesto1, 2)
        ElseIf objOcorrRemParcRem.iInstrucao2 = 6 Then
            .sPrazoProtesto = FormataCpoNum(objOcorrRemParcRem.iDiasDeProtesto2, 2)
        Else
            .sPrazoProtesto = FormataCpoNum("", 2)
        End If
        
        .sZeros6 = FormataCpoNum("", 1)
        .sNumSequencialRegistro = FormataCpoNum(gobjCobrancaEletronica.lSequencialRegistro, 6)
    
    End With
    
    'Incrementa o sequencial de registro
    gobjCobrancaEletronica.lSequencialRegistro = gobjCobrancaEletronica.lSequencialRegistro + 1
    
    Obter_Dados_Detalhe = SUCESSO

    Exit Function
    
Erro_Obter_Dados_Detalhe:

    Obter_Dados_Detalhe = Err
    
    Select Case Err

        Case 51895, 51897, 51898, 32267
       
        Case 51896
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ENDERECO_FILIALCLIENTE_NAO_INFORMADO", Err, gobjCobrancaEletronica.lCliente, gobjCobrancaEletronica.iFilialCli)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147401)

    End Select

    Exit Function

End Function

Private Function Montar_RegistroTrailer_ComReg_Unibanco(sRegTrailer As String, objCobrador As ClassCobrador) As Long
'Monta o registro

Dim tRegTrailerUnibanco As typeRemTrailerUnibanco
Dim lErro As Long

On Error GoTo Erro_Montar_RegistroTrailer_ComReg_Unibanco

    'Preenche os dados do regsitro de trailer
    tRegTrailerUnibanco.sIdentificacaoReg = "9"
    tRegTrailerUnibanco.sZeros = FormataCpoTexto("", 390)
    tRegTrailerUnibanco.sSequencialArq = FormataCpoTexto(gobjCobrancaEletronica.lSeqNomeArq, 3)
    tRegTrailerUnibanco.sSequencialRegistro = FormataCpoNum(gobjCobrancaEletronica.lSequencialRegistro, 6)
    
    'Monta o Registro de trailer
    sRegTrailer = tRegTrailerUnibanco.sIdentificacaoReg & tRegTrailerUnibanco.sZeros & tRegTrailerUnibanco.sSequencialArq & tRegTrailerUnibanco.sSequencialRegistro

    Montar_RegistroTrailer_ComReg_Unibanco = SUCESSO

    Exit Function

Erro_Montar_RegistroTrailer_ComReg_Unibanco:

    Montar_RegistroTrailer_ComReg_Unibanco = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147402)

    End Select

    Exit Function

End Function

Private Function Separa_Borderos_ComSem_Registro(colBorderos As Collection, colBorderosComReg As Collection, colBorderosSemReg As Collection) As Long
'Separa os borderos das carteiras com registro das carteiras sem registro.

Dim lErro As Long
Dim lComando As Long
Dim iRegistro As Integer
Dim objCarteiraCob As New ClassCarteiraCobrador
Dim objBordero As New ClassBorderoCobranca

On Error GoTo Erro_Separa_Borderos_ComSem_Registro
    
    'Para cada bordero da coleção lida
    For Each objBordero In colBorderos
        
        objCarteiraCob.iCobrador = objBordero.iCobrador
        objCarteiraCob.iCodCarteiraCobranca = objBordero.iCodCarteiraCobranca
        'Lê a carteira do cobrdor
        lErro = CF("CarteiraCobrador_Le", objCarteiraCob)
        If lErro <> SUCESSO And lErro <> 23551 Then Error 62015
        If lErro <> SUCESSO Then Error 62016 'Não encontrou
        
        'Se a carteira é registrada
        If objCarteiraCob.iComRegistro = CARTEIRA_COM_REGISTRO Then
            'Adiciona o bordero na coleção de borderos c\ registro
            colBorderosComReg.Add objBordero
        Else
            'Adiciona o bordero na coleção de borderos sem registro
            colBorderosSemReg.Add objBordero
        End If
    
    Next
    
    Separa_Borderos_ComSem_Registro = SUCESSO
    
    Exit Function
    
Erro_Separa_Borderos_ComSem_Registro:
    
    Separa_Borderos_ComSem_Registro = Err
    
    Select Case Err
    
        Case 62015
        
        Case 62016
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CARTEIRACOBRADOR_NAO_CADASTRADA1", Err, objCarteiraCob.iCodCarteiraCobranca, objCarteiraCob.iCobrador)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147403)
            
    End Select

    Exit Function
    
End Function

Function Gera_Arquivo_CobrancaComRegistro(lComando As Long, lComando1 As Long, lComando2 As Long, lComando3 As Long) As Long
'Gera o arquivo de remessa de cobrança eletrônica para as cobranças Com Registro
        
Dim iCancela As Integer
Dim sNomeArquivo As String
Dim lErro As Long
Dim sRegHeader As String
Dim sRegDetalhe As String
Dim sRegTrailer As String
Dim objBorderoCobranca As New ClassBorderoCobranca
Dim objCarteiraCobrador As New ClassCarteiraCobrador
Dim objOcorrRemParc As ClassOcorrRemParcRec, lSeqNomeArq As Long

On Error GoTo Erro_Gera_Arquivo_CobrancaComRegistro

    'Gera o nome do arquivo de acordo com o sequencial do Cobrador
    lErro = CF("CobrancaEletronica_Gera_NomeArquivo", gobjCobrancaEletronica.objCobrador, gobjCobrancaEletronica.sDiretorio, sNomeArquivo, lSeqNomeArq)
    If lErro <> SUCESSO Then gError 51899
    
    gobjCobrancaEletronica.lSeqNomeArq = lSeqNomeArq
    
    'Cria o arquivo de remessa
    Open sNomeArquivo For Output As #1
        
    lErro = Montar_RegistroHeader_ComReg_Unibanco(sRegHeader)
    If lErro <> SUCESSO Then gError 51900
        
    'Inclui o registro de Header
    Print #1, DesacentuaTexto(sRegHeader)

    For Each objBorderoCobranca In gobjCobrancaEletronica.colBorderos
        'Lê o bordero de cobrança
        lErro = CF("BorderoCobranca_LeMod", lComando, objBorderoCobranca)
        If lErro <> SUCESSO Then gError 51901
            
        objCarteiraCobrador.iCobrador = gobjCobrancaEletronica.iCobrador
        objCarteiraCobrador.iCodCarteiraCobranca = objBorderoCobranca.iCodCarteiraCobranca
        
        'Lê a carteira do cobrador
        lErro = CF("CarteiraCobrador_LeMod", objCarteiraCobrador, lComando1)
        If lErro <> SUCESSO Then gError 51902
       
        'Busca no BD as Ocorrências do Bordero em questão
        lErro = CF("CobrancaEletronica_Obter_OcorrRemParcRec", lComando2, objBorderoCobranca, gobjCobrancaEletronica)
        If lErro <> SUCESSO Then gError 51903
        
        'Para cada ocorrência
        For Each objOcorrRemParc In gobjCobrancaEletronica.colOcorrRemParcRec
    
            sRegDetalhe = ""
            'Monta um registro de detalhe do arquivo de remessa
            lErro = Montar_RegistroDetalhe_ComReg_Unibanco(sRegDetalhe, objOcorrRemParc, objCarteiraCobrador)
            If lErro <> SUCESSO And lErro <> 32270 Then gError 51904
            
            If lErro = SUCESSO Then
            
                'Inclui o registro com os detalhes do Título(Parcela) a receber
                Print #1, DesacentuaTexto(sRegDetalhe)
            
            End If
            
        Next
        
        objBorderoCobranca.iProcessado = BORDERO_PROCESSADO
        objBorderoCobranca.sNomeArquivo = sNomeArquivo
        
        'Atualiza o Status do Bordero de Cobrança
        lErro = CF("BorderoCobranca_Atualiza_Processo", lComando, lComando2, objBorderoCobranca)
        If lErro <> SUCESSO Then gError 51905
    
        'Atualiza o NossoNumero em carteiraCobrador
        lErro = CF("CarteiraCobrador_AtualizaNossoNumero", objCarteiraCobrador, lComando1, lComando3)
        If lErro <> SUCESSO Then gError 51906
    
        iCancela = 0
                
        Call gobjCobrancaEletronica.objTelaAtualizacao.Mostra_Evolucao(iCancela, 1)
        
        If iCancela = CANCELA_BATCH Then gError 79706
        
    Next
    'MOnta o registro de trailer
    lErro = Montar_RegistroTrailer_ComReg_Unibanco(sRegTrailer, objOcorrRemParc)
    If lErro <> SUCESSO Then gError 51907

    'Coloca o registro de trailer no arquivo
    Print #1, DesacentuaTexto(sRegTrailer)
    
    Close #1
    
    Gera_Arquivo_CobrancaComRegistro = SUCESSO
    
    Exit Function

Erro_Gera_Arquivo_CobrancaComRegistro:

    Gera_Arquivo_CobrancaComRegistro = gErr
    
    Select Case gErr
    
        Case 51899, 51901, 51903, 51905, 51902, 51900, 51904, 51906, 51907, 79706
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147404)
            
    End Select

    Exit Function

End Function
Function Gera_Arquivo_CobrancaSemRegistro(lComando As Long, lComando1 As Long, lComando2 As Long, lComando3 As Long) As Long
        
Dim iCancela As Integer
Dim sNomeArquivo As String
Dim lErro As Long
Dim sRegHeader As String
Dim sRegDetalhe As String
Dim sRegTrailer As String, lSeqNomeArq As Long
Dim objBorderoCobranca As New ClassBorderoCobranca
Dim objCarteiraCobrador As New ClassCarteiraCobrador
Dim objOcorrRemParc As ClassOcorrRemParcRec

On Error GoTo Erro_Gera_Arquivo_CobrancaSemRegistro

    'Gera o nome do arquivo de acordo com o sequencial do Cobrador
    lErro = CF("CobrancaEletronica_Gera_NomeArquivo", gobjCobrancaEletronica.objCobrador, gobjCobrancaEletronica.sDiretorio, sNomeArquivo, lSeqNomeArq)
    If lErro <> SUCESSO Then gError 51908
    
    gobjCobrancaEletronica.lSeqNomeArq = lSeqNomeArq
    
    'Cria o arquivo de remessa
    Open sNomeArquivo For Output As #3
        
    lErro = Montar_RegistroHeader_SemReg_Unibanco(sRegHeader)
    If lErro <> SUCESSO Then gError 51909
        
    'Inclui o registro de Header
    Print #3, sRegHeader

    For Each objBorderoCobranca In gobjCobrancaEletronica.colBorderos
        'Lê o bordero de cobrança
        lErro = CF("BorderoCobranca_LeMod", lComando, objBorderoCobranca)
        If lErro <> SUCESSO Then gError 51910
            
        objCarteiraCobrador.iCobrador = gobjCobrancaEletronica.iCobrador
        objCarteiraCobrador.iCodCarteiraCobranca = objBorderoCobranca.iCodCarteiraCobranca
        
        'Lê a carteira do cobrador
        lErro = CF("CarteiraCobrador_LeMod", objCarteiraCobrador, lComando1)
        If lErro <> SUCESSO Then gError 51911
       
        'Busca no BD as Ocorrências do Bordero em questão
        lErro = CF("CobrancaEletronica_Obter_OcorrRemParcRec", lComando2, objBorderoCobranca, gobjCobrancaEletronica)
        If lErro <> SUCESSO Then gError 51912
        
        'Para cada ocorrência
        For Each objOcorrRemParc In gobjCobrancaEletronica.colOcorrRemParcRec
    
            sRegDetalhe = ""
            'Monta um registro de detalhe do arquivo de remessa
            lErro = Montar_RegistroDetalhe_SemReg_Unibanco(sRegDetalhe, objOcorrRemParc, objCarteiraCobrador)
            If lErro <> SUCESSO And lErro <> 32269 Then gError 51913
            
            If lErro = SUCESSO Then
            
                'Inclui o registro com os detalhes do Título(Parcela) a receber
                Print #3, sRegDetalhe
                
            End If
            
        Next
        
        objBorderoCobranca.iProcessado = BORDERO_PROCESSADO
        objBorderoCobranca.sNomeArquivo = sNomeArquivo
        
        'Atualiza o Status do Bordero de Cobrança
        lErro = CF("BorderoCobranca_Atualiza_Processo", lComando, lComando2, objBorderoCobranca)
        If lErro <> SUCESSO Then gError 51914
    
        'Atualiza o NossoNumero em carteiraCobrador
        lErro = CF("CarteiraCobrador_AtualizaNossoNumero", objCarteiraCobrador, lComando1, lComando3)
        If lErro <> SUCESSO Then gError 51915
            
        iCancela = 0
        
        Call gobjCobrancaEletronica.objTelaAtualizacao.Mostra_Evolucao(iCancela, 1)
        
        If iCancela = CANCELA_BATCH Then gError 79704
    
    Next
    
    'MOnta o registro de trailer
    lErro = Montar_RegistroTrailer_SemReg_Unibanco(sRegTrailer, objOcorrRemParc)
    If lErro <> SUCESSO Then gError 51916

    'Coloca o registro de trailer no arquivo
    Print #3, sRegTrailer
    
    Close #3
    
    Gera_Arquivo_CobrancaSemRegistro = SUCESSO
    
    Exit Function

Erro_Gera_Arquivo_CobrancaSemRegistro:

    Gera_Arquivo_CobrancaSemRegistro = gErr
    
    Select Case gErr
    
        Case 51908, 51910, 51912, 51914, 51911, 51909, 51913, 51915, 51916, 79704
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147405)
            
    End Select
    
    Close #3

    Exit Function
    
End Function

Private Function Montar_RegistroHeader_SemReg_Unibanco(sRegHeader As String) As Long
'Monta o registro

Dim tRegHeaderUnibanco As typeRemHeaderUnibanco
Dim lErro As Long

On Error GoTo Erro_Montar_RegistroHeader_SemReg_Unibanco

    gobjCobrancaEletronica.lSequencialRegistro = 1

    'Carrega tRegHeaderUnibanco com os dados do Header do arquivo
    lErro = Obter_Dados_Header_SemReg(tRegHeaderUnibanco)
    If lErro <> SUCESSO Then Error 51917
    
    'Monta o registro de Header do Unibanco
    With tRegHeaderUnibanco
        sRegHeader = .sIdentificacaoReg & .sIdentificacaoArq & .sLiteralRemessa & .sCodTipoServico & .sLiteralServico & .sAgCredito
        sRegHeader = sRegHeader & .sGrupoEmpresarial1 & .sTipoFormulario & .sTipoCritica & .sTipoPostagem & .sGrupoEmpresarial2 & .sBrancos1
        sRegHeader = sRegHeader & .sDataEmissaoArq & .sDensidade & .sLiteralDensidade & .sBrancos2 & .sZeros1 & .sVersaoArquivo & .sSequencialRegistro
    End With
    
    Montar_RegistroHeader_SemReg_Unibanco = SUCESSO

    Exit Function

Erro_Montar_RegistroHeader_SemReg_Unibanco:

    Montar_RegistroHeader_SemReg_Unibanco = Err

    Select Case Err

        Case 51917

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147406)

    End Select

    Exit Function

End Function

Private Function Obter_Dados_Header_SemReg(tRegHeaderUnibanco As typeRemHeaderUnibanco) As Long

Dim lErro As Long
Dim objCCI As New ClassContasCorrentesInternas

On Error GoTo Erro_Obter_Dados_Header_SemReg
        
    'Lê a conta corrente associada ao cobrador
    lErro = CF("ContaCorrenteInt_Le", gobjCobrancaEletronica.objCobrador.iCodCCI, objCCI)
    If lErro <> SUCESSO And lErro <> 11807 Then Error 51918
    If lErro <> SUCESSO Then Error 51919
    
    Set gobjCobrancaEletronica.objCCI = objCCI
        
    With tRegHeaderUnibanco
        'Carrega os campos do cabeçalho
        .sIdentificacaoReg = "0"
        .sIdentificacaoArq = "1"
        .sLiteralRemessa = FormataCpoTexto("REMESSA", 7)
        .sCodTipoServico = "03"
        .sLiteralServico = FormataCpoTexto("COBR.  ESPECIAL", 15)
        .sAgCredito = FormataCpoNum(objCCI.sAgencia, 4) & FormataCpoNum(objCCI.sNumConta, 6) & FormataCpoNum(objCCI.sDVAgConta, 1)
        .sGrupoEmpresarial1 = FormataCpoNum(Obter_Info_Cobrador(gobjCobrancaEletronica.objCobradorCNABInfo.colInformacoes, 1), 4)
        .sTipoFormulario = FormataCpoNum(Obter_Info_Cobrador(gobjCobrancaEletronica.objCobradorCNABInfo.colInformacoes, 3), 1)
        .sTipoCritica = FormataCpoNum(Obter_Info_Cobrador(gobjCobrancaEletronica.objCobradorCNABInfo.colInformacoes, 4), 1)
        .sTipoPostagem = FormataCpoNum(Obter_Info_Cobrador(gobjCobrancaEletronica.objCobradorCNABInfo.colInformacoes, 5), 1)
        .sGrupoEmpresarial2 = FormataCpoTexto(Obter_Info_Cobrador(gobjCobrancaEletronica.objCobradorCNABInfo.colInformacoes, 2), 7)
        .sBrancos1 = FormataCpoTexto("", 43)
        .sDataEmissaoArq = Format(gdtDataHoje, "DDMMYY")
        .sDensidade = "01600"
        .sLiteralDensidade = "BPI"
        .sBrancos2 = FormataCpoTexto("", 116)
        .sZeros1 = FormataCpoNum("", 167)
        .sVersaoArquivo = FormataCpoNum(gobjCobrancaEletronica.lSeqNomeArq, 3)
        .sSequencialRegistro = FormataCpoNum(gobjCobrancaEletronica.lSequencialRegistro, 6)
    
    End With
    'Incrementa o seuqencial do arquivo
    gobjCobrancaEletronica.lSequencialRegistro = gobjCobrancaEletronica.lSequencialRegistro + 1

    Obter_Dados_Header_SemReg = SUCESSO

    Exit Function
    
Erro_Obter_Dados_Header_SemReg:

    Obter_Dados_Header_SemReg = Err
    
    Select Case Err
    
        Case 51918
        
        Case 51919
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_COBRADOR_NAO_ENCONTRADA", Err, gobjCobrancaEletronica.iCobrador)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147407)
            
    End Select
    
    Exit Function

End Function

Private Function Montar_RegistroDetalhe_SemReg_Unibanco(sRegDetalhe As String, objOcorrRemParcRec As ClassOcorrRemParcRec, objCarteiraCobrador As ClassCarteiraCobrador) As Long
'Monta o registro

Dim tRegDetalhe As typeRemDetalheUnibanco
Dim lErro As Long

On Error GoTo Erro_Montar_RegistroDetalhe_SemReg_Unibanco
    
    'Obtem os dados do registro de detalhe
    lErro = Obter_Dados_Detalhe_SemReg(tRegDetalhe, objOcorrRemParcRec, objCarteiraCobrador)
    If lErro <> SUCESSO And lErro <> 32268 Then Error 51920
    If lErro <> SUCESSO Then Error 32269
    
    With tRegDetalhe
        'Forma o registro de detalhe
        sRegDetalhe = .sIdentificacaoReg & .sNumReferenciaCli & .sDVReferencia & .sDataVencto & .sAgenciaDepositaria & .sAgenciaDebito & .sNomeSacado & .sEnderecoSacado
        sRegDetalhe = sRegDetalhe & .sBairro & .sCidade & .sEstado & .sCEP & .sDataEmissao & .sBrancos1 & .sQuantMoedas & .sValorTitulo & .sValorDesconto & .sValorJuros & .sValorMulta
        sRegDetalhe = sRegDetalhe & .sNumParcela & .sBrancos2 & .sNumTitulo & .sEspecie & .sAceite & .sDataProcessamento & .sCarteira & .sIndicadorMsg & .sDataDesconto & .sDataMulta & .sPrazoMora
        sRegDetalhe = sRegDetalhe & .sMoeda & .sBrancos3 & .sEndereco2 & .sBrancos4 & .sNumSequencialRegistro
    End With
    
    Montar_RegistroDetalhe_SemReg_Unibanco = SUCESSO

    Exit Function

Erro_Montar_RegistroDetalhe_SemReg_Unibanco:

    Montar_RegistroDetalhe_SemReg_Unibanco = Err

    Select Case Err
    
        Case 51920, 32269

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147408)

    End Select

    Exit Function

End Function

Private Function Obter_Dados_Detalhe_SemReg(tRegDetalhe As typeRemDetalheUnibanco, objOcorrRemParcRem As ClassOcorrRemParcRec, objCarteiraCobrador As ClassCarteiraCobrador) As Long
'Obtem os dados que são utilizados para preencher o registro de detalhe

Dim lErro As Long
Dim objParcelaRec As New ClassParcelaReceber
Dim objEndereco As New ClassEndereco
Dim sNossoNumero As String
Dim sDVNossoNumero As String
Dim dValorDesconto As Double

On Error GoTo Erro_Obter_Dados_Detalhe_SemReg

    objParcelaRec.lNumIntDoc = objOcorrRemParcRem.lNumIntParc

    'Busca os dados da parcela a receber no BD
    lErro = CF("CobrancaEletronica_Obter_DadosParcelaRec", objParcelaRec, gobjCobrancaEletronica)
    If lErro <> SUCESSO And lErro <> 51674 Then Error 51921
    If lErro <> SUCESSO Then Error 32268
    
    lErro = CF("CobrancaEletronica_Obter_Endereco", gobjCobrancaEletronica, objEndereco)
    If lErro <> SUCESSO Then Error 51924
            
    lErro = Calcula_NossoNumero_Unibanco(objCarteiraCobrador, sNossoNumero, sDVNossoNumero)
    If lErro <> SUCESSO Then Error 51925
    
    With tRegDetalhe
        .sIdentificacaoReg = "2"
        .sNumReferenciaCli = FormataCpoNum(sNossoNumero, 14)
        .sDVReferencia = FormataCpoNum(sDVNossoNumero, 1)
        .sDataVencto = Format(objParcelaRec.dtDataVencimento, "DDMMYY")
        .sAgenciaDepositaria = "9999" '????? problema
        .sAgenciaDebito = FormataCpoNum(gobjCobrancaEletronica.objCCI.sAgencia, 4) & FormataCpoNum(gobjCobrancaEletronica.objCCI.sNumConta, 6) & FormataCpoNum(gobjCobrancaEletronica.objCCI.sDVAgConta, 1)
        .sNomeSacado = FormataCpoTexto(gobjCobrancaEletronica.sNomeSacado, 30)
        .sEnderecoSacado = FormataCpoTexto("", 30)
        .sBairro = FormataCpoTexto(objEndereco.sBairro, 20)
        .sCidade = FormataCpoTexto(objEndereco.sCidade, 20)
        .sEstado = FormataCpoTexto(objEndereco.sSiglaEstado, 2)
        .sCEP = Formata_Cep(objEndereco.sCEP, 8)
        .sDataEmissao = Format(gobjCobrancaEletronica.dtDataEmissaoTitulo, "DDMMYY")
        .sBrancos1 = FormataCpoTexto("", 2)
        .sQuantMoedas = FormataCpoNum("", 10)
        .sValorTitulo = FormataCpoValor(objParcelaRec.dSaldo, 15)
        'Calcula o valor do desconto quando for do tipo percentual
        dValorDesconto = 0
        If objParcelaRec.iDesconto1Codigo = PERC_ANT_DIA Or objParcelaRec.iDesconto1Codigo = PERC_ANT_DIA_UTIL Or objParcelaRec.iDesconto1Codigo = Percentual Then
            dValorDesconto = StrParaDbl(Format(objParcelaRec.dDesconto1Valor * objParcelaRec.dSaldo, "Fixed"))
        Else
            If objParcelaRec.iDesconto1Codigo = VALOR_FIXO Then
                dValorDesconto = objParcelaRec.dDesconto1Valor
            End If
        End If
        .sValorDesconto = FormataCpoValor(dValorDesconto, 15)
        .sValorJuros = FormataCpoValor(objParcelaRec.dSaldo * objOcorrRemParcRem.dJuros, 12)
        .sValorMulta = FormataCpoNum("", 12)
        .sNumParcela = FormataCpoNum(objParcelaRec.iNumParcela, 3)
        .sBrancos2 = FormataCpoTexto("", 42)
        .sNumTitulo = FormataCpoNum(gobjCobrancaEletronica.lNumTitulo, 18)
        .sEspecie = FormataCpoNum(gobjCobrancaEletronica.iEspecieTitulo, 6) '??????olhar
        If objParcelaRec.iAceite = 1 Then
            .sAceite = "A "
        Else
            .sAceite = "N "
        End If
        .sDataProcessamento = Format(gdtDataHoje, "DDMMYY")
        .sCarteira = "20"
        .sIndicadorMsg = "0"
        If objParcelaRec.dtDesconto1Ate <> DATA_NULA Then
            .sDataDesconto = Format(objParcelaRec.dtDesconto1Ate, "DDMMYY")
        Else
            .sDataDesconto = FormataCpoNum("", 6)
        End If
        .sDataMulta = FormataCpoNum("", 6)
        .sPrazoMora = FormataCpoNum("", 3)
        .sMoeda = FormataCpoTexto("", 4)
        .sBrancos3 = FormataCpoTexto("", 13)
        .sEndereco2 = FormataCpoTexto(objEndereco.sEndereco, 60)
        .sBrancos4 = FormataCpoTexto("", 3)
        .sNumSequencialRegistro = FormataCpoNum(gobjCobrancaEletronica.lSequencialRegistro, 6)
    
    End With
    
    'Incrementa o sequencial de registro
    gobjCobrancaEletronica.lSequencialRegistro = gobjCobrancaEletronica.lSequencialRegistro + 1
    
    Obter_Dados_Detalhe_SemReg = SUCESSO

    Exit Function
    
Erro_Obter_Dados_Detalhe_SemReg:

    Obter_Dados_Detalhe_SemReg = Err
    
    Select Case Err

        Case 51921, 51924, 51925, 51922, 32268
       
        Case 51923
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ENDERECO_FILIALCLIENTE_NAO_INFORMADO", Err, gobjCobrancaEletronica.lCliente, gobjCobrancaEletronica.iFilialCli)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147409)

    End Select

    Exit Function

End Function

Private Function Montar_RegistroTrailer_SemReg_Unibanco(sRegTrailer As String, objCobrador As ClassCobrador) As Long
'Monta o registro

Dim tRegTrailerUnibanco As typeRemTrailerUnibanco
Dim lErro As Long

On Error GoTo Erro_Montar_RegistroTrailer_SemReg_Unibanco

    'Preenche os dados do registro de trailer
    tRegTrailerUnibanco.sIdentificacaoReg = "9"
    tRegTrailerUnibanco.sBrancos1 = FormataCpoTexto("", 25)
    tRegTrailerUnibanco.sAgCtaDVCedente = FormataCpoNum(gobjCobrancaEletronica.objCCI.sAgencia, 4) & FormataCpoNum(gobjCobrancaEletronica.objCCI.sNumConta, 6) & FormataCpoNum(gobjCobrancaEletronica.objCCI.sDVAgConta, 1)
    tRegTrailerUnibanco.sBrancos2 = FormataCpoTexto("", 334)
    tRegTrailerUnibanco.sQuantRegistros = FormataCpoNum(gobjCobrancaEletronica.lSequencialRegistro, 6)
    tRegTrailerUnibanco.sTotalTitulos = FormataCpoNum(gobjCobrancaEletronica.lSequencialRegistro - 2, 17)
    tRegTrailerUnibanco.sSequencialRegistro = FormataCpoNum(gobjCobrancaEletronica.lSequencialRegistro, 6)
    
    'Monta o Registro de trailer
    sRegTrailer = tRegTrailerUnibanco.sIdentificacaoReg & tRegTrailerUnibanco.sBrancos1 & tRegTrailerUnibanco.sAgCtaDVCedente & tRegTrailerUnibanco.sBrancos2 & tRegTrailerUnibanco.sQuantRegistros & tRegTrailerUnibanco.sTotalTitulos & tRegTrailerUnibanco.sSequencialRegistro

    Montar_RegistroTrailer_SemReg_Unibanco = SUCESSO

    Exit Function

Erro_Montar_RegistroTrailer_SemReg_Unibanco:

    Montar_RegistroTrailer_SemReg_Unibanco = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147410)

    End Select

    Exit Function

End Function

Function Processar_ArquivoRetorno_Cobranca_Unibanco(objCobrancaEletronica As ClassCobrancaEletronica) As Long
'Faz o processamento do arquivo de retorno

Dim lErro As Long
Dim iIndice As Integer
Dim sRegistro As String
Dim iCodRegistro As Integer
Dim objBorderoRetCobr As New ClassBorderoRetCobr
Dim lComando As Long
Dim lComando1 As Long
Dim sCodigoOperacao As String

On Error GoTo Erro_Processar_ArquivoRetorno_Cobranca_Unibanco

    objBorderoRetCobr.iBanco = objCobrancaEletronica.objCobrador.iCodBanco
    
    'le as carteiras de cobranca do cobrador
    lErro = CF("CarteirasCobrador_Le", objCobrancaEletronica.objCobrador, objBorderoRetCobr.colCartCobr)
    If lErro <> SUCESSO Then Error 32265
    
    'Abre os comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 51927

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 51928

    'Abre o arquivo de retorno
    Open objCobrancaEletronica.sNomeArquivoRetorno For Input As #2
    
    sCodigoOperacao = ""
    objBorderoRetCobr.objMovCCI.dValor = 0
    
    'Até chegar ao fim do arquivo
    Do While Not EOF(2)
    
        'Busca o próximo registro do arquivo (na 1a vez vai ser o de header)
        Line Input #2, sRegistro
    
        'Pega o código de identificação do registro
        iCodRegistro = StrParaInt(Left(sRegistro, 1))
        sCodigoOperacao = Trim(Mid(sRegistro, 115, 5))
        
        'Se o registro for o de Header
        If iCodRegistro = 0 Then
            'Processa os dados contidos no Header
            lErro = Processa_Registro_Header(sRegistro, objBorderoRetCobr, objCobrancaEletronica.objCobrador, objCobrancaEletronica.sNomeArquivoRetorno, sCodigoOperacao)
            If lErro <> SUCESSO Then Error 51929

        'Se o registro for de Detalhe ("Transação")
        ElseIf iCodRegistro = 1 Then
            'Processa os dados contidos no registro de detalhe
            lErro = Processa_Registro_Detalhe(sRegistro, objBorderoRetCobr, objCobrancaEletronica, lComando, lComando1, sCodigoOperacao)
            If lErro <> SUCESSO Then Error 51930
        
        'Se o registro for de Trailer
        ElseIf iCodRegistro = 9 Then
            'Processa os dados do trailer do arquivo
            lErro = Processa_Registro_Trailer(sRegistro)
            If lErro <> SUCESSO Then Error 51931
        
        Else
            Error 51932
        End If
        
    Loop
        
    Close #2

    'Grava os dados recolhidos durante o processamento do arquivo de retorno
    lErro = CF("BorderoRetCobr_ConfirmaGrava", objBorderoRetCobr)
    If lErro <> SUCESSO Then Error 51933

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    Processar_ArquivoRetorno_Cobranca_Unibanco = SUCESSO
    
    Exit Function
    
Erro_Processar_ArquivoRetorno_Cobranca_Unibanco:

    Processar_ArquivoRetorno_Cobranca_Unibanco = Err
    
    Select Case Err
    
        Case 51926
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 51927, 51928
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 51929, 51930, 51931, 51933, 32265
        
        Case 51932
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORMATO_ARQUIVO_INCORRETO", Err, objCobrancaEletronica.sNomeArquivoRetorno)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147411)
    
    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Close #2

    Exit Function

End Function

Private Function Processa_Registro_Header(sRegistro As String, objBorderoRetCobr As ClassBorderoRetCobr, objCobrador As ClassCobrador, sNomeArquivo As String, sCodigoOperacao As String) As Long

Dim lErro As Long
Dim tRegHeaderUnibanco As typeRetHeaderUnibanco
Dim objContaCorrente As New ClassContasCorrentesInternas

On Error GoTo Erro_Processa_Registro_Header

    If UCase(sCodigoOperacao) <> "CERET" Then
    
        'Extrai da variável lida no arquivo os dados do Header
        lErro = Extrair_Header_Retorno(tRegHeaderUnibanco, sRegistro)
        If lErro <> SUCESSO Then Error 51935
    
    Else
        lErro = Extrair_Header_Retorno_SemReg(tRegHeaderUnibanco, sRegistro)
        If lErro <> SUCESSO Then Error 51936
            
    End If
    'Preenche o objBorderoRetCobr
    objBorderoRetCobr.dtDataRecepcao = gdtDataHoje
    objBorderoRetCobr.iCobrador = objCobrador.iCodigo
    objBorderoRetCobr.sNomeArq = sNomeArquivo
    
    'Pega a conta corrente do cobrador
    objContaCorrente.iCodigo = objCobrador.iCodCCI
    
    'Lê a conta corrente do cobrador
    lErro = CF("ContaCorrenteInt_Le", objContaCorrente.iCodigo, objContaCorrente)
    If lErro <> SUCESSO And lErro <> 11807 Then Error 51937
    
    'Se não achou a Conta Corrente --> erro
    If lErro <> SUCESSO Then Error 51938
    
    'Preenche parcialmente o objMovCCI
    objBorderoRetCobr.objMovCCI.iFilialEmpresa = objContaCorrente.iFilialEmpresa
    objBorderoRetCobr.objMovCCI.iCodConta = objContaCorrente.iCodigo
    objBorderoRetCobr.objMovCCI.iTipo = MOVCCI_RECEBIMENTO_TITULO
    objBorderoRetCobr.objMovCCI.iExcluido = NAO_EXCLUIDO
    objBorderoRetCobr.objMovCCI.iTipoMeioPagto = DINHEIRO
    objBorderoRetCobr.objMovCCI.dtDataBaixa = gdtDataHoje
    objBorderoRetCobr.objMovCCI.dtDataMovimento = Extrair_Data(tRegHeaderUnibanco.sDataMovto)
    objBorderoRetCobr.objMovCCI.iConciliado = NAO_CONCILIADO
           
    Processa_Registro_Header = SUCESSO
    
    Exit Function

Erro_Processa_Registro_Header:

    Processa_Registro_Header = Err

    Select Case Err
    
        Case 51935, 51937, 51936
    
        Case 51938
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_CORRENTE_NAO_ENCONTRADA", Err, objContaCorrente.iCodigo)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147412)
            
    End Select
    
    Exit Function

End Function

Private Function Extrair_Header_Retorno(tRegHeaderUnibanco As typeRetHeaderUnibanco, sRegistro As String) As Long
'Extrai do sRegistro os dados do header do arquido de retorno

Dim lErro As Long

On Error GoTo Erro_Extrair_Header_Retorno

    With tRegHeaderUnibanco
        .sIdentificacaoReg = Mid(sRegistro, 1, 1)
        .sIdentificacaoArq = Mid(sRegistro, 2, 1)
        .sLiteralRetorno = Mid(sRegistro, 3, 7)
        .sCodTipoServico = Mid(sRegistro, 10, 2)
        .sLiteralServico = Mid(sRegistro, 12, 8)
        .sBrancos1 = Mid(sRegistro, 20, 7)
        .sCodEmpresa = Mid(sRegistro, 27, 11)
        .sZeros = Mid(sRegistro, 38, 9)
        .sNomeEmpresa = Mid(sRegistro, 47, 30)
        .sNumeroBanco = Mid(sRegistro, 77, 3)
        .sNomeBanco = Mid(sRegistro, 80, 8)
        .sBrancos2 = Mid(sRegistro, 88, 7)
        .sDataEmissaoArq = Mid(sRegistro, 95, 6)
        .sDensidade = Mid(sRegistro, 101, 5)
        .sLiteralDensidade = Mid(sRegistro, 106, 3)
        .sBrancos3 = Mid(sRegistro, 109, 17)
        .sDataMovto = Mid(sRegistro, 126, 6)
        .sZeros1 = Mid(sRegistro, 132, 254)
        .sMoeda = Mid(sRegistro, 386, 2)
        .sBrancos4 = Mid(sRegistro, 388, 4)
        .sNumGeracaoArq = Mid(sRegistro, 392, 3)
        .sSequencialRegistro = Mid(sRegistro, 395, 6)
    
    End With

    Extrair_Header_Retorno = SUCESSO

    Exit Function

Erro_Extrair_Header_Retorno:

    Extrair_Header_Retorno = Err
    
    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147413)
            
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

Private Function Processa_Registro_Detalhe(sRegistro As String, objBorderoRetCobr As ClassBorderoRetCobr, objCobrancaEletronica As ClassCobrancaEletronica, lComando As Long, lComando1 As Long, sCodigoOperacao As String) As Long

Dim tRegDetalhe As typeRetDetalheUnibanco
Dim iCodOcorrencia As Integer, iTipoOcorrencia As Integer
Dim lErro As Long
Dim objDetRetCobr As New ClassDetRetCobr
Dim objBorderoRetCobrCart As ClassBorderoRetCobrCart

On Error GoTo Erro_Processa_Registro_Detalhe

    If UCase(sCodigoOperacao) <> "CERET" Then
        'Extrai da variável lida no arquivo os dados do registro de detalhe
        lErro = Extrair_Detalhe_Retorno(tRegDetalhe, sRegistro)
        If lErro <> SUCESSO Then Error 51939
    Else
        lErro = Extrair_Detalhe_Retorno_SemReg(tRegDetalhe, sRegistro)
        If lErro <> SUCESSO Then Error 51940
    End If
    
    'Recolhe o código da ocorrência
    iCodOcorrencia = StrParaInt(tRegDetalhe.sCodOcorrencia)
    
    Select Case iCodOcorrencia
        
        'Quando ocorre a baixa
        Case 6, 7, 8, 9, 10
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
    
        Case 51939, 51941, 51943, 51945, 51940
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147414)
            
    End Select
    
    Exit Function

End Function

Private Function Extrair_Detalhe_Retorno(tRegDetalhe As typeRetDetalheUnibanco, sRegistro As String) As Long
'Extrai do sRegistro os dados do Detalhe do arquido de retorno

Dim lErro As Long

On Error GoTo Erro_Extrair_Detalhe_Retorno

    With tRegDetalhe
        
        .sIdentificacaoReg = Mid(sRegistro, 1, 1)
        .sTipoInscricaoEmpresa = Mid(sRegistro, 2, 2)
        .sNumInscricaoEmpresa = Mid(sRegistro, 4, 14)
        .sCodEmpresa = Mid(sRegistro, 18, 11)
        .sZeros1 = Mid(sRegistro, 29, 9)
        .sUsoDaEmpresa = Mid(sRegistro, 38, 25)
        .sNossoNumero1 = Mid(sRegistro, 63, 11)
        .sZeros2 = Mid(sRegistro, 74, 9)
        .sIdenfOpBanco = Mid(sRegistro, 83, 25)
        .sCarteira = Mid(sRegistro, 108, 1)
        .sCodOcorrencia = Mid(sRegistro, 109, 2)
        .sDataOcorrencia = Mid(sRegistro, 111, 6)
        .sSeuNumero = Mid(sRegistro, 117, 10)
        .sNossoNumero2 = Mid(sRegistro, 127, 11)
        .sZeros3 = Mid(sRegistro, 138, 9)
        .sVencimento = Mid(sRegistro, 147, 6)
        .sValorTitulo = Mid(sRegistro, 153, 13)
        .sBancoCobrador = Mid(sRegistro, 166, 3)
        .sAgenciaCobradora = Mid(sRegistro, 169, 5)
        .sEspecie = Mid(sRegistro, 174, 2)
        .sTarifaCobranca = Mid(sRegistro, 176, 13)
        .sOutrasDespesas = Mid(sRegistro, 189, 13)
        .sZeros4 = Mid(sRegistro, 202, 26)
        .sAbatimento = Mid(sRegistro, 228, 13)
        .sDescontos = Mid(sRegistro, 241, 13)
        .sValorPrincipal = Mid(sRegistro, 254, 13)
        .sJuros = Mid(sRegistro, 267, 13)
        .sZeros5 = Mid(sRegistro, 280, 54)
        .sValorOriginalTitulo = Mid(sRegistro, 334, 13)
        .sNomeSacado = Mid(sRegistro, 347, 30)
        .sMoeda = Mid(sRegistro, 377, 2)
        .sZeros6 = Mid(sRegistro, 379, 7)
        .sDataEmissaoArq = Mid(sRegistro, 386, 6)
        .sNumGeracaoArq = Mid(sRegistro, 392, 3)
        .sNumSequencialRegistro = Mid(sRegistro, 395, 6)
    
    End With
    
    Extrair_Detalhe_Retorno = SUCESSO

    Exit Function

Erro_Extrair_Detalhe_Retorno:

    Extrair_Detalhe_Retorno = Err
    
    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147415)
            
    End Select
    
    Exit Function

End Function

Function Processa_Registro_Trailer(sRegistroTrailer As String) As Long

Dim lErro As Long
Dim tRegTrailerUnibanco As typeRetTrailerUnibanco

On Error GoTo Erro_Processa_Registro_Trailer

    'Extrai da variável lida no arquivo os dados do registro de detalhe
    lErro = Extrair_Trailer_Retorno(tRegTrailerUnibanco, sRegistroTrailer)
    If lErro <> SUCESSO Then Error 51946
    
    Processa_Registro_Trailer = SUCESSO
    
    Exit Function
    
Erro_Processa_Registro_Trailer:

    Processa_Registro_Trailer = Err
    
    Select Case Err
    
        Case 51946
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147416)
            
    End Select
    
    Exit Function
    
End Function

Private Function Extrair_Trailer_Retorno(tRegTrailerUnibanco As typeRetTrailerUnibanco, sRegistro As String) As Long
'Extrai do sRegistro os dados do Trailer do arquido de retorno

Dim lErro As Long

On Error GoTo Erro_Extrair_Trailer_Retorno

    tRegTrailerUnibanco.sIdentificacaoReg = Mid(sRegistro, 1, 1)
    tRegTrailerUnibanco.sQuantTitulos = Mid(sRegistro, 2, 8)
    tRegTrailerUnibanco.sValorTotal = Mid(sRegistro, 10, 14)
    tRegTrailerUnibanco.sAvisoBancario = Mid(sRegistro, 24, 8)
    tRegTrailerUnibanco.sZeros = Mid(sRegistro, 32, 360)
    tRegTrailerUnibanco.sNumGeracaoArq = Mid(sRegistro, 392, 3)
    tRegTrailerUnibanco.sSequencialRegistro = Mid(sRegistro, 395, 6)

    Extrair_Trailer_Retorno = SUCESSO

    Exit Function

Erro_Extrair_Trailer_Retorno:

    Extrair_Trailer_Retorno = Err
    
    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147417)
            
    End Select
    
    Exit Function

End Function

Function BorderoPagto_Criar_ArquivoCNAB_Unibanco(objGeracaoArqCNABPag As ClassGeracaoArqCNABPag) As Long
'Cria o arquivo de remessa de títulos a pagar nos padrões do Unibanco

Dim lErro As Long
Dim sNomeArquivo As String
Dim lTransacao As Long
Dim iIndice As Integer
Dim iTipoCobranca As Integer
Dim sRegHeaderArq As String
Dim sRegDetalhe As String
Dim sTrailerArq As String
Dim objBaixaParcPag As New ClassBaixaParcPagar
Dim alComando(0 To 2) As Long

On Error GoTo Erro_BorderoPagto_Criar_ArquivoCNAB_Unibanco

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 51947
    Next

    'Abre a transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 51948
    
    'Inicializa os obj globais da tela
    Set gobjGeracaoArqCNABPag = objGeracaoArqCNABPag
    Set gColFiliaisEmpresa = New Collection
    
    'faz o lock no bordero de pagamento
    lErro = CF("BorderoPagto_Lock", alComando(0), gobjGeracaoArqCNABPag.objBorderoPagto.lNumIntBordero)
    If lErro <> SUCESSO Then Error 51949
    
    'Gera o nome do arquivo
    lErro = CF("PagamentoEletronico_Gera_NomeArquivo", sNomeArquivo, gobjGeracaoArqCNABPag)
    If lErro <> SUCESSO Then Error 51950
 
    'Cria o arquivo de remessa
    Open sNomeArquivo For Output As #1
 
    gobjGeracaoArqCNABPag.lQuantRegistrosArq = 1
    
    'Monta o registro de Header do arquivo
    lErro = Montar_RegistroHeaderCNABPag_Unibanco(sRegHeaderArq)
    If lErro <> SUCESSO Then Error 51951

    
    'Inclui no arquivo
    Print #1, DesacentuaTexto(sRegHeaderArq)
    
    iTipoCobranca = gobjGeracaoArqCNABPag.objBorderoPagto.iTipoDeCobranca
    gobjGeracaoArqCNABPag.lSequencialRegDetalhe = 1
    gobjGeracaoArqCNABPag.dValorTotalLote = 0
    
    'Lê as baixas ligadas ao bordero
    lErro = CF("BaixasParcPag_Le_BorderoPagto", alComando(2), gobjGeracaoArqCNABPag.objBorderoPagto.lNumIntBordero, gobjGeracaoArqCNABPag.colBaixasParcPag)
    If lErro <> SUCESSO Then Error 51952
    
    'Para cada baixa executada
    For Each objBaixaParcPag In gobjGeracaoArqCNABPag.colBaixasParcPag
        
        lErro = CF("BaixaParcPag_ObterDados_ParcelaTituloPag", alComando(2), objBaixaParcPag.lNumIntParcela, gobjGeracaoArqCNABPag)
        If lErro <> SUCESSO Then Error 51953
        
        If iTipoCobranca = TIPO_COBRANCA_DEP_CONTA Then
        
            'Monta um registro de
            lErro = Montar_Registro_DetalhePagto_CCI(sRegDetalhe, objBaixaParcPag)
            If lErro <> SUCESSO Then Error 51954
        
        ElseIf iTipoCobranca = TIPO_COBRANCA_DOC Then
            
            lErro = Montar_Registro_DetalhePagto_DOC(sRegDetalhe, objBaixaParcPag)
            If lErro <> SUCESSO Then Error 51955
        
        ElseIf iTipoCobranca = TIPO_COBRANCA_OP Then
            lErro = Montar_Registro_DetalhePagto_OP(sRegDetalhe, objBaixaParcPag)
            If lErro <> SUCESSO Then Error 51956
        End If
        
        'incluir no arq
        Print #1, DesacentuaTexto(sRegDetalhe)
    
        gobjGeracaoArqCNABPag.lSequencialRegDetalhe = gobjGeracaoArqCNABPag.lSequencialRegDetalhe + 1
        gobjGeracaoArqCNABPag.dValorTotalLote = gobjGeracaoArqCNABPag.dValorTotalLote + objBaixaParcPag.dValorBaixado - objBaixaParcPag.dValorDesconto + objBaixaParcPag.dValorJuros
        gobjGeracaoArqCNABPag.lQuantRegistrosArq = gobjGeracaoArqCNABPag.lQuantRegistrosArq + 1
        Call objGeracaoArqCNABPag.objTelaAtualizacao.Mostra_Evolucao(True, 1)
    
    Next

    'Monta o registro de trailer de arquivo
    lErro = Montar_Registro_TrailerArq(sTrailerArq)
    If lErro <> SUCESSO Then Error 51957
    
    'Inclui no arquivo
    Print #1, DesacentuaTexto(sTrailerArq)
    
    'Fecha arquivo
    Close #1
    
    'Atualiza em BorderosPagto o Nome e o sequencial do arquivo
    lErro = Comando_ExecutarPos(alComando(1), "UPDATE BorderosPagto SET NomeArq = ? , NumArqRemessa = ?", alComando(0), sNomeArquivo, gobjGeracaoArqCNABPag.lSequencialArq + 1)
    If lErro <> AD_SQL_SUCESSO Then Error 51958
    
    'Faz o commit na transacao
    lErro = Transacao_Commit()
    If lErro <> SUCESSO Then Error 51959
        
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    BorderoPagto_Criar_ArquivoCNAB_Unibanco = SUCESSO
    
    Exit Function
    
Erro_BorderoPagto_Criar_ArquivoCNAB_Unibanco:

    BorderoPagto_Criar_ArquivoCNAB_Unibanco = Err
    
    Select Case Err
    
        Case 51947
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 51948
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
        
        Case 51949, 51951, 51725, 51952, 51953, 51954, 51955, 51956, 51730, 51731, 51957, 51950
        
        Case 51958
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_BORDERO_PAGTO", Err)
        
        Case 51959
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147418)
            
    End Select
    
    'Faz um RollBack na transacao
    Call Transacao_Rollback
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function



Function Montar_RegistroHeaderCNABPag_Unibanco(sRegHeader As String) As Long
'Monta o registro de header do arquivo de contas a pagar

Dim lErro As Long
Dim tRegHeaderArq As typePagtoHeaderArqUnibanco

On Error GoTo Erro_Montar_RegistroHeaderCNABPag_Unibanco
    
    'Inicializa o contador do número de registros do arquivo
    gobjGeracaoArqCNABPag.lQuantRegistrosArq = 1

    'Recolhe os dados do cabeçalho
    tRegHeaderArq.sIdentificacaoReg = "0"
    tRegHeaderArq.sIdentificacaoArq = "1"
    tRegHeaderArq.sLiteralRemessa = "REMESSA"
    tRegHeaderArq.sCodTipoServico = "08"
    tRegHeaderArq.sLiteralServico = "CONTAS A PAGAR"
    tRegHeaderArq.sAgenciaCliente = FormataCpoNum(gobjGeracaoArqCNABPag.objCCI.sAgencia, 4)
    tRegHeaderArq.sContaCliente = FormataCpoNum(gobjGeracaoArqCNABPag.objCCI.sNumConta, 6)
    tRegHeaderArq.sDVCliente = FormataCpoNum(gobjGeracaoArqCNABPag.objCCI.sDVNumConta, 1)
    tRegHeaderArq.sNomeCliente = FormataCpoTexto(gsNomeEmpresa, 30)
    tRegHeaderArq.sNumeroBanco = "409"
    tRegHeaderArq.sNomeBanco = "UNIBANCO"
    tRegHeaderArq.sDataEmissaoArq = Format(gdtDataHoje, "DDMMYY")
    tRegHeaderArq.sBrancos = FormataCpoTexto("", 224)
    tRegHeaderArq.sSequencialRegistro = FormataCpoNum(gobjGeracaoArqCNABPag.lQuantRegistrosArq, 6)
    
    'Forma o registro com os dados recolhidos p\ o cabeçalho
    With tRegHeaderArq
        sRegHeader = .sIdentificacaoReg & .sIdentificacaoArq & .sLiteralRemessa & .sCodTipoServico & .sLiteralServico & .sAgenciaCliente & .sContaCliente & .sDVCliente
        sRegHeader = sRegHeader & .sNomeCliente & .sNumeroBanco & .sNomeBanco & .sDataEmissaoArq & .sBrancos & .sSequencialRegistro
    End With
    
    gobjGeracaoArqCNABPag.lQuantRegistrosArq = gobjGeracaoArqCNABPag.lQuantRegistrosArq + 1
    
    Montar_RegistroHeaderCNABPag_Unibanco = SUCESSO
    
    Exit Function
    
Erro_Montar_RegistroHeaderCNABPag_Unibanco:

    Montar_RegistroHeaderCNABPag_Unibanco = Err
    
    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147419)
            
    End Select
    
    Exit Function
    
End Function


Function Montar_Registro_DetalhePagto_CCI(sRegDetalhe As String, objBaixaParcPag As ClassBaixaParcPagar) As Long
'Monta o registro de transação do arquivo com os dados obtidos no BD

Dim lErro As Long
Dim tRegDetalhe As typePagtoDetalheUnibanco
Dim objEndereco As New ClassEndereco
Dim sAgencia As String
Dim sConta As String
Dim sDV As String
Dim sCheckHorizontal As String

On Error GoTo Erro_Montar_Registro_DetalhePagto_CCI

    objEndereco.lCodigo = gobjGeracaoArqCNABPag.lEnderecoFavorecido
    'Lê o endereço do favorecido
    lErro = CF("Endereco_Le", objEndereco)
    If lErro <> SUCESSO And lErro <> 12309 Then Error 51960
    If lErro <> SUCESSO Then Error 51961 'Não encontrou

    'Recolhe os dados do registro de detalhe
    With tRegDetalhe
        .sIdentificacaoReg = "2"
        .sAgenciaCliente = FormataCpoNum(gobjGeracaoArqCNABPag.objCCI.sAgencia, 4)
        .sContaCliente = FormataCpoNum(gobjGeracaoArqCNABPag.objCCI.sNumConta, 6)
        .sDVCliente = FormataCpoNum(gobjGeracaoArqCNABPag.objCCI.sDVNumConta, 1)
        .sBrancos1 = FormataCpoNum("", 20)
        .sBancoForn = FormataCpoTexto(gobjGeracaoArqCNABPag.iBancoFavorecido, 3)
        Call Extrair_Agencia_DV(gobjGeracaoArqCNABPag.sAgenciaFavorecido, sAgencia, 4, sDV, 1)
        .sAgenciaForn = sAgencia
        Call Extrair_Conta_DV(gobjGeracaoArqCNABPag.sContaFavorecido, sConta, 10, sDV, 1)
        .sContaCorrenteForn = sConta
        .sMeioRepasse = "4"
        .sEmissaoAviso = "0"
        .sTipoOperacao = "5"
        .sTipoServico = "4"
        .sValorCredito = FormataCpoValor(objBaixaParcPag.dValorBaixado, 13)
        .sIdentfFavorecido = FormataCpoNum("", 15)
        .sDataGravacao = Format(gobjGeracaoArqCNABPag.dtDataEmissaoTitulo, "DDMMYY")
        .sDataCredito = Format(gobjGeracaoArqCNABPag.dtDataVenctoParcela, "DDMMYY")
        .sNomeFavorecido = FormataCpoValor(gobjGeracaoArqCNABPag.sNomeFavorecido, 30)
        .sBrancos2 = FormataCpoTexto("", 95)
        .sHistorico = FormataCpoTexto("", 15)
        .sMoeda = "14"
        .sBrancos3 = FormataCpoTexto("", 11)
        .sNossoNumero = FormataCpoTexto(gobjGeracaoArqCNABPag.sNossoNumero, 11)
        .sCodOcorrencia = "01"
        .sMesmaTitularidade = "0"
        .sBrancos4 = FormataCpoTexto("", 1)
        Call Calcula_CheckHorizontal(sCheckHorizontal, .sBancoForn & .sAgenciaForn & .sContaCorrenteForn, .sValorCredito, .sTipoOperacao)
        .sCheckHorizontal = sCheckHorizontal
        .sBrancos5 = FormataCpoTexto("", 2)
        .sDVCCFavorecido = FormataCpoTexto(sDV, 1)
        .sNumSequencialRegistro = FormataCpoNum(gobjGeracaoArqCNABPag.lSequencialRegDetalhe, 6)
            
        'Monta o registro de detalhe com os dados obtidos
        sRegDetalhe = .sIdentificacaoReg & .sAgenciaCliente & .sContaCliente & .sDVCliente & .sBrancos1 & .sBancoForn & .sAgenciaForn & .sContaCorrenteForn
        sRegDetalhe = sRegDetalhe & .sMeioRepasse & .sEmissaoAviso & .sTipoOperacao & .sTipoServico & .sValorCredito & .sIdentfFavorecido & .sDataGravacao
        sRegDetalhe = sRegDetalhe & .sDataCredito & .sNomeFavorecido & .sBrancos2 & .sHistorico & .sMoeda & .sBrancos3 & .sNossoNumero & .sCodOcorrencia
        sRegDetalhe = sRegDetalhe & .sMesmaTitularidade & .sBrancos4 & .sCheckHorizontal & .sBrancos5 & .sDVCCFavorecido & .sNumSequencialRegistro

    End With

    Montar_Registro_DetalhePagto_CCI = SUCESSO
    
    Exit Function

Erro_Montar_Registro_DetalhePagto_CCI:

    Montar_Registro_DetalhePagto_CCI = Err
    
    Select Case Err
    
        Case 51960
        
        Case 51961
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ENDERECO_NAO_CADASTRADO", Err, objEndereco.lCodigo)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147420)
            
    End Select

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


Function Montar_Registro_DetalhePagto_DOC(sRegDetalhe As String, objBaixaParcPag As ClassBaixaParcPagar) As Long
'Monta o registro de transação do arquivo com os dados obtidos no BD

Dim lErro As Long
Dim tRegDetalhe As typePagtoDetalheUnibanco
Dim objEndereco As New ClassEndereco
Dim sAgencia As String
Dim sConta As String
Dim sDVAg As String
Dim sDVCta As String
Dim sCheckHorizontal As String
Dim sCEP As String

On Error GoTo Erro_Montar_Registro_DetalhePagto_DOC

    objEndereco.lCodigo = gobjGeracaoArqCNABPag.lEnderecoFavorecido
    'Lê o endereço do favorecido
    lErro = CF("Endereco_Le", objEndereco)
    If lErro <> SUCESSO And lErro <> 12309 Then Error 51962
    If lErro <> SUCESSO Then Error 51963 'Não encontrou

    'Recolhe os dados do registro de detalhe
    With tRegDetalhe
        .sIdentificacaoReg = "2"
        .sAgenciaCliente = FormataCpoNum(gobjGeracaoArqCNABPag.objCCI.sAgencia, 4)
        .sContaCliente = FormataCpoNum(gobjGeracaoArqCNABPag.objCCI.sNumConta, 6)
        .sDVCliente = FormataCpoNum(gobjGeracaoArqCNABPag.objCCI.sDVNumConta, 1)
        .sBrancos1 = FormataCpoNum("", 17)
        .sComplCepForn = Right(Formata_Cep(objEndereco.sCEP, 8), 3)
        .sBancoForn = FormataCpoTexto(gobjGeracaoArqCNABPag.iBancoFavorecido, 3)
        Call Extrair_Agencia_DV(gobjGeracaoArqCNABPag.sAgenciaFavorecido, sAgencia, 4, sDVAg, 1)
        .sAgenciaForn = sAgencia
        Call Extrair_Conta_DV(gobjGeracaoArqCNABPag.sContaFavorecido, sConta, 10, sDVCta, 1)
        .sContaCorrenteForn = sConta
        .sMeioRepasse = "4"
        .sEmissaoAviso = "0"
        .sTipoOperacao = "5"
        .sTipoServico = "4"
        .sValorCredito = FormataCpoValor(objBaixaParcPag.dValorBaixado, 13)
        .sIdentfFavorecido = FormataCpoNum("", 15)
        .sDataGravacao = Format(gobjGeracaoArqCNABPag.dtDataEmissaoTitulo, "DDMMYY")
        .sDataCredito = Format(gobjGeracaoArqCNABPag.dtDataVenctoParcela, "DDMMYY")
        .sNomeFavorecido = FormataCpoValor(gobjGeracaoArqCNABPag.sNomeFavorecido, 30)
        .sEndereco = FormataCpoTexto(objEndereco.sEndereco, 30)
        .sCidade = FormataCpoTexto(objEndereco.sCidade, 15)
        .sEstado = FormataCpoTexto(objEndereco.sSiglaEstado, 2)
        .sCEP = Mid(Formata_Cep(objEndereco.sCEP, 8), 1, 5)
        .sBrancos2 = FormataCpoTexto("", 20)
        .sNomeAgencia = FormataCpoTexto("", 20)
        .sPracaCompensacao = FormataCpoNum("", 3)
        .sHistorico = FormataCpoTexto("", 15)
        .sMoeda = "14"
        .sBrancos3 = FormataCpoTexto("", 11)
        .sNossoNumero = FormataCpoTexto(gobjGeracaoArqCNABPag.sNossoNumero, 11)
        .sCodOcorrencia = "01"
        .sMesmaTitularidade = "0"
        .sBrancos4 = FormataCpoTexto("", 1)
        Call Calcula_CheckHorizontal(sCheckHorizontal, .sBancoForn & .sAgenciaForn & .sContaCorrenteForn, .sValorCredito, .sTipoOperacao)
        .sCheckHorizontal = sCheckHorizontal
        .sBrancos5 = FormataCpoTexto("", 2)
        .sCepForn = Left(Formata_Cep(objEndereco.sCEP, 8), 5)
        .sBrancos6 = FormataCpoTexto("", 1)
        .sDVAgenciaForn = FormataCpoTexto(sDVAg, 1)
        .sDVCCFavorecido = FormataCpoTexto(sDVCta, 1)
        .sNumSequencialRegistro = FormataCpoNum(gobjGeracaoArqCNABPag.lSequencialRegDetalhe, 6)
            
        'Monta o registro de detalhe com os dados obtidos
        sRegDetalhe = .sIdentificacaoReg & .sAgenciaCliente & .sContaCliente & .sDVCliente & .sBrancos1 & .sComplCepForn & .sBancoForn & .sAgenciaForn
        sRegDetalhe = sRegDetalhe & .sContaCorrenteForn & .sMeioRepasse & .sEmissaoAviso & .sTipoOperacao & .sTipoServico & .sValorCredito & .sIdentfFavorecido
        sRegDetalhe = sRegDetalhe & .sDataGravacao & .sDataCredito & .sNomeFavorecido & .sEndereco & .sCidade & .sEstado & .sCEP & .sBrancos2 & .sNomeAgencia
        sRegDetalhe = sRegDetalhe & .sPracaCompensacao & .sHistorico & .sMoeda & .sBrancos3 & .sNossoNumero & .sCodOcorrencia & .sMesmaTitularidade & .sBrancos4
        sRegDetalhe = sRegDetalhe & .sCheckHorizontal & .sBrancos5 & .sCepForn & .sBrancos6 & .sDVAgenciaForn & .sDVCCFavorecido & .sNumSequencialRegistro

    End With

    Montar_Registro_DetalhePagto_DOC = SUCESSO
    
    Exit Function

Erro_Montar_Registro_DetalhePagto_DOC:

    Montar_Registro_DetalhePagto_DOC = Err
    
    Select Case Err
    
        Case 51962
        
        Case 51963
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ENDERECO_NAO_CADASTRADO", Err, objEndereco.lCodigo)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147421)
            
    End Select

End Function

Function Montar_Registro_DetalhePagto_OP(sRegDetalhe As String, objBaixaParcPag As ClassBaixaParcPagar) As Long
'Monta o registro de transação do arquivo com os dados obtidos no BD

Dim lErro As Long
Dim tRegDetalhe As typePagtoDetalheUnibanco
Dim objEndereco As New ClassEndereco
Dim sAgencia As String
Dim sConta As String
Dim sDVAg As String
Dim sDVCta As String
Dim sCheckHorizontal As String

On Error GoTo Erro_Montar_Registro_DetalhePagto_OP

    objEndereco.lCodigo = gobjGeracaoArqCNABPag.lEnderecoFavorecido
    'Lê o endereço do favorecido
    lErro = CF("Endereco_Le", objEndereco)
    If lErro <> SUCESSO And lErro <> 12309 Then Error 51964
    If lErro <> SUCESSO Then Error 51965 'Não encontrou

    'Recolhe os dados do registro de detalhe
    With tRegDetalhe
        .sIdentificacaoReg = "2"
        .sAgenciaCliente = FormataCpoNum(gobjGeracaoArqCNABPag.objCCI.sAgencia, 4)
        .sContaCliente = FormataCpoNum(gobjGeracaoArqCNABPag.objCCI.sNumConta, 6)
        .sDVCliente = FormataCpoNum(gobjGeracaoArqCNABPag.objCCI.sDVNumConta, 1)
        .sBrancos1 = FormataCpoNum("", 17)
        .sComplCepForn = Mid(Formata_Cep(objEndereco.sCEP, 8), 6, 3)
        .sBancoForn = FormataCpoTexto(gobjGeracaoArqCNABPag.iBancoFavorecido, 3)
        Call Extrair_Agencia_DV(gobjGeracaoArqCNABPag.sAgenciaFavorecido, sAgencia, 4, sDVAg, 1)
        .sAgenciaForn = sAgencia
        .sBrancos2 = FormataCpoNum("", 10)
        .sMeioRepasse = "4"
        .sEmissaoAviso = "0"
        .sTipoOperacao = "5"
        .sTipoServico = "7"
        .sValorCredito = FormataCpoValor(objBaixaParcPag.dValorBaixado, 13)
        .sIdentfFavorecido = FormataCpoNum("", 15)
        .sDataGravacao = Format(gobjGeracaoArqCNABPag.dtDataEmissaoTitulo, "DDMMYY")
        .sDataCredito = Format(gobjGeracaoArqCNABPag.dtDataVenctoParcela, "DDMMYY")
        .sNomeFavorecido = FormataCpoValor(gobjGeracaoArqCNABPag.sNomeFavorecido, 30)
        .sEndereco = FormataCpoTexto(objEndereco.sEndereco, 30)
        .sCidade = FormataCpoTexto(objEndereco.sCidade, 15)
        .sEstado = FormataCpoTexto(objEndereco.sSiglaEstado, 2)
        .sCEP = Mid(Formata_Cep(objEndereco.sCEP, 8), 1, 5)
        .sBrancos3 = FormataCpoTexto("", 43)
        .sHistorico = FormataCpoTexto("", 15)
        .sMoeda = "14"
        .sBrancos4 = FormataCpoTexto("", 11)
        .sNossoNumero = FormataCpoTexto(gobjGeracaoArqCNABPag.sNossoNumero, 11)
        .sCodOcorrencia = "01"
        .sBrancos5 = FormataCpoTexto("", 2)
        Call Calcula_CheckHorizontal(sCheckHorizontal, .sBancoForn & .sAgenciaForn & .sContaCorrenteForn, .sValorCredito, .sTipoOperacao)
        .sCheckHorizontal = FormataCpoNum(sCheckHorizontal, 18)
        .sBrancos5 = FormataCpoTexto("", 34)
        .sNumSequencialRegistro = FormataCpoNum(gobjGeracaoArqCNABPag.lSequencialRegDetalhe, 6)
            
        'Monta o registro de detalhe com os dados obtidos
        sRegDetalhe = .sIdentificacaoReg & .sAgenciaCliente & .sContaCliente & .sDVCliente & .sBrancos1 & .sComplCepForn & .sBancoForn & .sAgenciaForn & .sBrancos2 & .sMeioRepasse
        sRegDetalhe = sRegDetalhe & .sEmissaoAviso & .sTipoOperacao & .sTipoServico & .sValorCredito & .sIdentfFavorecido & .sDataGravacao & .sDataCredito & .sNomeFavorecido & .sEndereco
        sRegDetalhe = sRegDetalhe & .sCidade & .sEstado & .sCEP & .sBrancos3 & .sHistorico & .sMoeda & .sBrancos4 & .sNossoNumero & .sCodOcorrencia & .sBrancos5 & .sCheckHorizontal & .sBrancos5 & .sNumSequencialRegistro
    End With

    Montar_Registro_DetalhePagto_OP = SUCESSO
    
    Exit Function

Erro_Montar_Registro_DetalhePagto_OP:

    Montar_Registro_DetalhePagto_OP = Err
    
    Select Case Err
    
        Case 51964
        
        Case 51965
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ENDERECO_NAO_CADASTRADO", Err, objEndereco.lCodigo)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147422)
            
    End Select
    
    Exit Function

End Function

Function Montar_Registro_TrailerArq(sTrailerArq As String) As Long
'Monta o registro de trailer de arquivo

Dim lErro As Long
Dim tRegTrailer As typePagtoTrailerArqUnibanco

On Error GoTo Erro_Montar_Registro_TrailerArq
    
    'Incrementa a quantidade de registros do arquivo
    gobjGeracaoArqCNABPag.lQuantRegistrosArq = gobjGeracaoArqCNABPag.lQuantRegistrosArq + 1

    'Recolhe os dados do trailer
    With tRegTrailer
        .sIdentificacaoReg = "9"
        .sTotalQtdRegistros = FormataCpoNum(gobjGeracaoArqCNABPag.lQuantRegistrosLote, 6)
        .sTotalValorPagtos = FormataCpoValor(gobjGeracaoArqCNABPag.dValorTotalLote, 13)
        .sQuantRegistrosArq = FormataCpoNum(gobjGeracaoArqCNABPag.lQuantRegistrosArq, 6)
        .sBrancos = FormataCpoTexto("", 288)
        .sNumSequencialRegistro = FormataCpoNum(gobjGeracaoArqCNABPag.lQuantRegistrosArq, 6)
    
        'Monta o registro com os dados recolhidos
        sTrailerArq = .sIdentificacaoReg & .sTotalQtdRegistros & .sTotalValorPagtos & .sQuantRegistrosArq & .sBrancos & .sNumSequencialRegistro
    
    End With

    Montar_Registro_TrailerArq = SUCESSO
    
    Exit Function
    
Erro_Montar_Registro_TrailerArq:

    Montar_Registro_TrailerArq = Err
    
    Select Case Err
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147423)
            
    End Select
    
    Exit Function

End Function
Private Function Extrair_Header_Retorno_SemReg(tRegHeaderUnibanco As typeRetHeaderUnibanco, sRegistro As String) As Long
'Extrai do sRegistro os dados do header do arquido de retorno

Dim lErro As Long

On Error GoTo Erro_Extrair_Header_Retorno_SemReg

    With tRegHeaderUnibanco
        .sIdentificacaoReg = Mid(sRegistro, 1, 1)
        .sIdentificacaoArq = Mid(sRegistro, 2, 1)
        .sLiteralRetorno = Mid(sRegistro, 3, 7)
        .sCodTipoServico = Mid(sRegistro, 10, 2)
        .sLiteralServico = Mid(sRegistro, 12, 8)
        .sBrancos1 = Mid(sRegistro, 20, 7)
        .sCodEmpresa = Mid(sRegistro, 27, 11)
        .sZeros = Mid(sRegistro, 38, 9)
        .sNomeEmpresa = Mid(sRegistro, 47, 30)
        .sNumeroBanco = Mid(sRegistro, 77, 3)
        .sNomeBanco = Mid(sRegistro, 80, 12)
        .sBrancos2 = Mid(sRegistro, 88, 3)
        .sDataEmissaoArq = Mid(sRegistro, 95, 6)
        .sDensidade = Mid(sRegistro, 101, 5)
        .sLiteralDensidade = Mid(sRegistro, 106, 3)
        .sDataMovto = Mid(sRegistro, 109, 6)
        .sCodOperacao = Mid(sRegistro, 115, 5)
        .sBrancos3 = Mid(sRegistro, 120, 272)
        .sNumGeracaoArq = Mid(sRegistro, 392, 3)
        .sSequencialRegistro = Mid(sRegistro, 395, 6)
    
    End With

    Extrair_Header_Retorno_SemReg = SUCESSO

    Exit Function

Erro_Extrair_Header_Retorno_SemReg:

    Extrair_Header_Retorno_SemReg = Err
    
    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147424)
            
    End Select
    
    Exit Function

End Function
Private Function Extrair_Detalhe_Retorno_SemReg(tRegDetalhe As typeRetDetalheUnibanco, sRegistro As String) As Long
'Extrai do sRegistro alguns dados do Detalhe do arquido de retorno

Dim lErro As Long

On Error GoTo Erro_Extrair_Detalhe_Retorno_SemReg

    With tRegDetalhe
        
        .sIdentificacaoReg = Mid(sRegistro, 1, 1)
        .sUsoDaEmpresa = Mid(sRegistro, 38, 25)
        .sNossoNumero1 = Mid(sRegistro, 63, 20)
        .sCodOcorrencia = Mid(sRegistro, 109, 2)
        .sDataOcorrencia = Mid(sRegistro, 111, 6)
        .sSeuNumero = Mid(sRegistro, 117, 10)
        .sVencimento = Mid(sRegistro, 147, 6)
        .sValorTitulo = Mid(sRegistro, 153, 13)
        .sValorPrincipal = Mid(sRegistro, 254, 13)
        .sNumSequencialRegistro = Mid(sRegistro, 395, 6)
    
    End With
    
    Extrair_Detalhe_Retorno_SemReg = SUCESSO

    Exit Function

Erro_Extrair_Detalhe_Retorno_SemReg:

    Extrair_Detalhe_Retorno_SemReg = Err
    
    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147425)
            
    End Select
    
    Exit Function

End Function


Private Sub Calcula_CheckHorizontal(sCheckHorizontal As String, sDadosFornecedor As String, sValorTransacao As String, sTipoOperacao As String)
'Calcula o conteúdo do campo CheckHorizontal

Dim iTipoOperacao As Integer
Dim iIndice As Integer
Dim iDigito1 As Integer
Dim iDigito2 As Integer
Dim iSoma As Integer
Dim iVaiUm As Integer
Dim sResultado As String
Dim iProduto As Integer

    sResultado = ""
    iVaiUm = 0
    
    'Soma o Valor da transacao aos Dados do Cliente
    For iIndice = 13 To 1 Step -1
        
        iDigito1 = StrParaInt(Mid(sDadosFornecedor, iIndice + 5, 1))
        iDigito2 = StrParaInt(Mid(sValorTransacao, iIndice, 1))
        
        iSoma = iDigito1 + iDigito2 + iVaiUm
                
        iVaiUm = 0
        
        If iSoma > 9 Then iVaiUm = 1
        
        sResultado = Right(Str(iSoma), 1) & sResultado

    Next
    
    'Termina de completar o Resultado
    If iVaiUm > 0 Then
        
        iIndice = 5
        
        Do While iVaiUm > 0 And iIndice > 0
        
            iDigito1 = StrParaInt(Mid(sDadosFornecedor, iIndice, 1))
            
            iSoma = iDigito1 + iVaiUm
            
            iVaiUm = 0
            
            If iSoma > 9 Then iVaiUm = 1
            
            sResultado = Right(Str(iSoma), 1) & sResultado
                    
            iIndice = iIndice - 1
        
        Loop
        
    End If
    
    If Len(sResultado) < 18 Then sResultado = Left(sDadosFornecedor, 18 - Len(sResultado)) & sResultado
    
    iTipoOperacao = StrParaInt(sTipoOperacao)
    sCheckHorizontal = ""
    iVaiUm = 0
    
    For iIndice = 18 To 1 Step -1
    
        iDigito1 = StrParaInt(Mid(sResultado, iIndice, 1))
        
        iProduto = (iDigito1 * iTipoOperacao) + iVaiUm
                
        iVaiUm = 0
        
        If iProduto > 9 Then iVaiUm = StrParaInt(Left(iProduto, 1))
        
        sCheckHorizontal = Right(Str(iProduto), 1) & sCheckHorizontal

    Next
        
    Exit Sub
    
End Sub

Private Function Obter_Info_Cobrador(colInformacoes As Collection, iCodInformacao As Integer) As String
'Procura a informação com o código passado na coleção de informações do cobrador

Dim objCodNome As AdmCodigoNome

    'para cada informação armazenada
    For Each objCodNome In colInformacoes
        'procura aquela com o código pedido
        If objCodNome.iCodigo = iCodInformacao Then
            Obter_Info_Cobrador = objCodNome.sNome
            Exit Function
            Exit For
        End If
    Next
    
    Obter_Info_Cobrador = ""
    
End Function

Function Calcula_NossoNumero_Unibanco(objCarteiraCobrador As ClassCarteiraCobrador, sNossoNumero As String, sDVNossoNumero As String) As Long
'Calcula o campo Nosso Numero do Unibanco

Dim sProxNossoNumero As String

On Error GoTo Erro_Calcula_NossoNumero_Unibanco

    sNossoNumero = 0
    
    If objCarteiraCobrador.iGeraNossoNumero = BANCO_GERA_NOSSONUMERO Then
        sNossoNumero = objCarteiraCobrador.sFaixaNossoNumeroProx
        If Val(sNossoNumero) > Val(objCarteiraCobrador.sFaixaNossoNumeroProx) Then gError 62017
        
        'se o próximo número não foi preenchido = > erro
        If Len(Trim(sNossoNumero)) = 0 Then gError 79702
        
        objCarteiraCobrador.sFaixaNossoNumeroProx = sNossoNumero + 1
    End If
    
    Calcula_NossoNumero_Unibanco = SUCESSO
    
    Exit Function
    
Erro_Calcula_NossoNumero_Unibanco:

    Calcula_NossoNumero_Unibanco = gErr
    
    Select Case gErr
    
        Case 62017
            Call Rotina_Erro(vbOKOnly, "ERRO_FAIXA_NOSSONUMERO_INSUFICIENTE", gErr)
        
        Case 79702
            Call Rotina_Erro(vbOKOnly, "ERRO_NOSSO_NUMERO_NAO_DEFINIDO", gErr, objCarteiraCobrador.iCodCarteiraCobranca, objCarteiraCobrador.iCobrador)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147426)
    
    End Select

    Exit Function

End Function

Private Sub RegDetCobr_ObterInfo(objDetRetCobr As ClassDetRetCobr, tRegDetalhe As typeRetDetalheUnibanco)
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
        .iCodCarteiraNoBanco = StrParaInt(tRegDetalhe.sCarteira)
        .dValorTarifa = Extrair_Valor(tRegDetalhe.sTarifaCobranca)
        .sSeuNumero = tRegDetalhe.sSeuNumero
        .sUsoDaEmpresa = tRegDetalhe.sSeuNumero
        .lNumSequencial = CLng(tRegDetalhe.sNumSequencialRegistro)
        .dValorTitulo = Extrair_Valor(tRegDetalhe.sValorTitulo)
        
    End With
    
End Sub

Private Function Prepara_CartCobrRet1(objCobrancaEletronica As ClassCobrancaEletronica, objBorderoRetCobr As ClassBorderoRetCobr, objBorderoRetCobrCart As ClassBorderoRetCobrCart, objDetRetCobr As ClassDetRetCobr, tRegDetalhe As typeRetDetalheUnibanco) As Long
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147427)
            
    End Select

    Exit Function

End Function


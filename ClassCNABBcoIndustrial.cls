VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCNABBcoIndustrial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'O BANCO INDUSTRIAL JÁ SEGUE O PADRÃO CNAB 240

Dim gobjCobrancaEletronica As ClassCobrancaEletronica
Dim giNumLotes As Integer
Dim glNumRegistros As Long

Dim gobjFilialEmpresa As AdmFiliais

Const BANCOINDUSTRIAL_INFO_CONVENIO = 1

Function Criar_Arquivo_Remessa_BancoIndustrial(objCobrancaEletronica As ClassCobrancaEletronica) As Long
'Cria uma arquivo de remessa de cobranca com base nos dados do BD para o Banco Industrial

Dim lErro As Long
Dim iIndice As Integer
Dim lTransacao As Long
Dim lSeqNomeArq As Long
Dim sNomeArquivo As String
Dim alComando(0 To 7) As Long
Dim objOcorrRemParc As ClassOcorrRemParcRec
Dim objBorderoCobranca As New ClassBorderoCobranca
Dim objCarteiraCobrador As New ClassCarteiraCobrador
Dim sRegHeader As String
Dim sRegDetalheSegP As String
Dim sRegDetalheSegQ As String
Dim sRegDetalheSegR As String
Dim sRegTrailer As String

On Error GoTo Erro_Criar_Arquivo_Remessa_BancoIndustrial

    'Abre os comandos para acesso ao Banco de Dados
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 86413
    Next

    'Abre a transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 86417

    Set gobjCobrancaEletronica = objCobrancaEletronica

    'Gera o nome do arquivo de acordo com o sequencial do Cobrador
    lErro = CF("CobrancaEletronica_Gera_NomeArquivo", objCobrancaEletronica.objCobrador, objCobrancaEletronica.sDiretorio, sNomeArquivo, lSeqNomeArq)
    If lErro <> SUCESSO Then gError 86418

    gobjCobrancaEletronica.lSeqNomeArq = lSeqNomeArq

    'Lê a conta corrente associada ao cobrador
    lErro = CF("ContaCorrenteInt_Le", gobjCobrancaEletronica.objCobrador.iCodCCI, gobjCobrancaEletronica.objCCI)
    If lErro <> SUCESSO And lErro <> 11807 Then gError 86419
    If lErro <> SUCESSO Then gError 86420

    If gobjCobrancaEletronica.objCCI.iCodBanco <> gobjCobrancaEletronica.objCobrador.iCodBanco Then gError 86421
    If Len(Trim(gobjCobrancaEletronica.objCCI.sAgencia)) = 0 Or Len(Trim(gobjCobrancaEletronica.objCCI.sNumConta)) = 0 Then gError 86422

    Set gobjFilialEmpresa = New AdmFiliais
    gobjFilialEmpresa.iCodFilial = giFilialEmpresa

    lErro = CF("FilialEmpresa_Le", gobjFilialEmpresa, True)
    If lErro <> SUCESSO And lErro <> 27378 Then gError 86423
    If lErro <> SUCESSO Then gError 86424

    'Cria o arquivo de remessa
    Open sNomeArquivo For Output As #1

    lErro = Montar_RegistroHeaderArquivo_BancoIndustrial(sRegHeader)
    If lErro <> SUCESSO Then gError 86425

    'Inclui o registro de Header de arquivo
    Print #1, DesacentuaTexto(sRegHeader)

    glNumRegistros = glNumRegistros + 1

    iIndice = 0
    For Each objBorderoCobranca In objCobrancaEletronica.colBorderos

        'Lê o bordero de cobrança
        lErro = CF("BorderoCobranca_LeMod", alComando(2), objBorderoCobranca)
        If lErro <> SUCESSO Then gError 86426

        objCarteiraCobrador.iCobrador = objCobrancaEletronica.iCobrador
        objCarteiraCobrador.iCodCarteiraCobranca = objBorderoCobranca.iCodCarteiraCobranca

        'Lê a carteira do cobrador
        lErro = CF("CarteiraCobrador_LeMod", objCarteiraCobrador, alComando(4))
        If lErro <> SUCESSO Then gError 86427

        'Busca no BD as Ocorrências do Bordero em questão
        lErro = CF("CobrancaEletronica_Obter_OcorrRemParcRec", alComando(3), objBorderoCobranca, gobjCobrancaEletronica)
        If lErro <> SUCESSO Then gError 86428

        If iIndice <> 0 Then
            If objCobrancaEletronica.colBorderos(iIndice).iCodCarteiraCobranca <> objBorderoCobranca.iCodCarteiraCobranca Then giNumLotes = giNumLotes + 1
        Else
            giNumLotes = giNumLotes + 1
        End If

        'incluir header de lote
        lErro = Montar_RegistroHeaderLote_BancoIndustrial(sRegHeader, objCarteiraCobrador)
        If lErro <> SUCESSO Then gError 86429

        'Inclui o registro de Header de lote
        Print #1, DesacentuaTexto(sRegHeader)

        glNumRegistros = glNumRegistros + 1

        'Número de Registros dentro do lote
        gobjCobrancaEletronica.lSequencialRegistro = 0

        'Para cada ocorrência
         For Each objOcorrRemParc In objCobrancaEletronica.colOcorrRemParcRec

            sRegDetalheSegP = ""
            sRegDetalheSegQ = ""
            sRegDetalheSegR = ""

            'Monta um registro de detalhe do arquivo de remessa
            lErro = Montar_RegistroDetalhe_BancoIndustrial(sRegDetalheSegP, sRegDetalheSegQ, sRegDetalheSegR, objOcorrRemParc, objCarteiraCobrador)
            If lErro <> SUCESSO Then gError 86430

            'Inclui o registro com os detalhes do Título(Parcela) a receber
            Print #1, DesacentuaTexto(sRegDetalheSegP)
            Print #1, DesacentuaTexto(sRegDetalheSegQ)

            glNumRegistros = glNumRegistros + 2

            If sRegDetalheSegR <> "" Then

                Print #1, DesacentuaTexto(sRegDetalheSegR)
                glNumRegistros = glNumRegistros + 1

            End If

        Next

        'incluir trailer de lote

        'MOnta o registro de trailer de lote
        lErro = Montar_RegistroTrailerLote_BancoIndustrial(sRegTrailer)
        If lErro <> SUCESSO Then gError 86431

        'grava o registro trailer de lote
        Print #1, DesacentuaTexto(sRegTrailer)

        glNumRegistros = glNumRegistros + 1

        objBorderoCobranca.iProcessado = BORDERO_PROCESSADO
        objBorderoCobranca.sNomeArquivo = sNomeArquivo

        'Atualiza o Status do Bordero de Cobrança
        lErro = CF("BorderoCobranca_Atualiza_Processo", alComando(2), alComando(3), objBorderoCobranca)
        If lErro <> SUCESSO Then gError 86432

        'Atualiza o NossoNumero em carteiraCobrador
        lErro = CF("CarteiraCobrador_AtualizaNossoNumero", objCarteiraCobrador, alComando(4), alComando(5))
        If lErro <> SUCESSO Then gError 86433

        iIndice = iIndice + 1
    Next

    'MOnta o registro de trailer de arquivo
    lErro = Montar_RegistroTrailerArq_BancoIndustrial(sRegTrailer)
    If lErro <> SUCESSO Then gError 86434

    'grava o registro de trailer no arquivo
    Print #1, DesacentuaTexto(sRegTrailer)

    Close #1

    'atualizar seq de arquivo de remessa no cobrador
    lErro = Cobrador_Atualizar(objCobrancaEletronica.objCobrador, alComando(6), alComando(7))
    If lErro <> SUCESSO Then gError 86435

    lErro = Transacao_Commit()
    If lErro <> SUCESSO Then gError 86436

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Criar_Arquivo_Remessa_BancoIndustrial = SUCESSO

    Exit Function

Erro_Criar_Arquivo_Remessa_BancoIndustrial:

    Criar_Arquivo_Remessa_BancoIndustrial = gErr

    Close #1

    Select Case gErr

        Case 86413
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 86417
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 86418, 86419, 86423, 86425, 86426, 86427, 86428, 86429, 86430, 86431, 86432, _
             86433, 86434, 86435
        
        Case 86420
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTA_CORRENTE_NAO_ENCONTRADA", gErr, gobjCobrancaEletronica.objCobrador.iCodCCI)
        
        Case 86421
            Call Rotina_Erro(vbOKOnly, "ERRO_BANCO_CCI_DIFERENTE_COBRADOR", gErr, gobjCobrancaEletronica.objCCI.iCodBanco, gobjCobrancaEletronica.objCobrador.iCodBanco)
        
        Case 86422
            Call Rotina_Erro(vbOKOnly, "ERRO_AGENCIA_CONTA_COBRADOR_NAO_PREENCHIDAS", gErr)
        
        Case 86424
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, giFilialEmpresa)
        
        Case 86436
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT_TRANSACAO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147164)

    End Select

    Call Transacao_Rollback

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Private Function Montar_RegistroHeaderArquivo_BancoIndustrial(sRegHeader As String) As Long
'Monta o registro

Dim tRegHeaderBancoIndustrial240 As typeRemHeaderCNAB240
Dim lErro As Long

On Error GoTo Erro_Montar_RegistroHeaderArquivo_BancoIndustrial

    'Carrega tRegHeaderBancoIndustrial240 com os dados do Header do aquivo
    lErro = Obter_Dados_Header_Arquivo(tRegHeaderBancoIndustrial240)
    If lErro <> SUCESSO Then gError 86415

    With tRegHeaderBancoIndustrial240

        'Monta o registro de Header do Banco Industrial
        sRegHeader = .sCodBancoComp & .sLoteServico & .sRegHeaderLote & .sBrancosCNAB1
        sRegHeader = sRegHeader & .sTipoInscEmpresa & .sCodInsEmpresa & .sCodConvBanco & .sCodConvBanco 'o convenio substitui a conta
        sRegHeader = sRegHeader & .sNomeEmpresa & .sMensagem1 & .sMensagem2 & "1" & .sDataGravRem & "120000" & .sNumRemessa & .sNVersaoLayLote & "00000" & FormataCpoTexto("", 54) & "000" & "  " & FormataCpoTexto("", 10)

    End With

    Montar_RegistroHeaderArquivo_BancoIndustrial = SUCESSO

    Exit Function

Erro_Montar_RegistroHeaderArquivo_BancoIndustrial:

    Montar_RegistroHeaderArquivo_BancoIndustrial = gErr

    Select Case gErr

        Case 86415

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147165)

    End Select

    Exit Function

End Function

Private Function Obter_Dados_Header_Arquivo(tRegHeaderBancoIndustrial240 As typeRemHeaderCNAB240) As Long

On Error GoTo Erro_Obter_Dados_Header_Arquivo

    'Carrega os campos do cabeçalho
    With tRegHeaderBancoIndustrial240

        .sCodBancoComp = "604"
        .sLoteServico = "0000"
        .sRegHeaderLote = "0"
        .sBrancosCNAB1 = FormataCpoTexto("", 9)

        If Len(gobjFilialEmpresa.sCgc) = 14 Then
            .sTipoInscEmpresa = "2"
        Else
            .sTipoInscEmpresa = "1"
        End If

        .sCodInsEmpresa = FormataCpoNum(gobjFilialEmpresa.sCgc, 14)
        .sCodConvBanco = FormataCpoTexto(Obter_Info_Cobrador(gobjCobrancaEletronica.objCobradorCNABInfo.colInformacoes, BANCOINDUSTRIAL_INFO_CONVENIO), 20)

        .sAgMantConta = FormataCpoNum(gobjCobrancaEletronica.objCCI.sAgencia, 5)
        .sDVAgencia = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVAgencia, 1)
        .sNumContaCorrente = FormataCpoNum(gobjCobrancaEletronica.objCCI.sNumConta, 12)
        .sDVConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVNumConta, 1)
        .sDVAgenciaConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVAgConta, 1)

        .sNomeEmpresa = FormataCpoTexto(gsNomeEmpresa, 30)
        .sMensagem1 = FormataCpoTexto("BANCO INDUSTRIAL", 30)
        .sMensagem2 = FormataCpoTexto("", 10)

        .sDataGravRem = Format(gdtDataHoje, "DDMMYYYY")
        .sNumRemessa = FormataCpoNum(gobjCobrancaEletronica.objCobrador.lCNABProxSeqArqCobr, 6)
        .sNVersaoLayLote = "040"

    End With

    Obter_Dados_Header_Arquivo = SUCESSO

    Exit Function

Erro_Obter_Dados_Header_Arquivo:

    Obter_Dados_Header_Arquivo = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147166)

    End Select

    Exit Function

End Function

Private Function Montar_RegistroHeaderLote_BancoIndustrial(sRegHeader As String, objCarteiraCobrador As ClassCarteiraCobrador) As Long
'Monta o registro

Dim lErro As Long
Dim tRegHeaderBancoIndustrial240 As typeRemHeaderCNAB240

On Error GoTo Erro_Montar_RegistroHeaderLote_BancoIndustrial

    'Carrega tRegHeaderBco com os dados do Header do aquivo
    lErro = Obter_Dados_Header_Lote(tRegHeaderBancoIndustrial240, objCarteiraCobrador)
    If lErro <> SUCESSO Then gError 86416

    With tRegHeaderBancoIndustrial240

        'Monta o registro de Header do Bco
        sRegHeader = .sCodBancoComp & .sLoteServico & .sRegHeaderLote & .sTipoOperacao & .sTipoServico & .sFormaLancamento & .sNVersaoLayLote & .sBrancosCNAB1
        sRegHeader = sRegHeader & .sTipoInscEmpresa & .sCodInsEmpresa & .sCodConvBanco & .sCodConvBanco 'o convenio substitui a conta
        sRegHeader = sRegHeader & .sNomeEmpresa & .sMensagem1 & .sMensagem2 & .sNumRemessa & .sDataGravRem & .sDataCredito & .sBrancosCNAB2

    End With

    Montar_RegistroHeaderLote_BancoIndustrial = SUCESSO

    Exit Function

Erro_Montar_RegistroHeaderLote_BancoIndustrial:

    Montar_RegistroHeaderLote_BancoIndustrial = gErr

    Select Case gErr

        Case 86416

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147167)

    End Select

    Exit Function

End Function

Private Function Obter_Dados_Header_Lote(tRegHeaderBancoIndustrial240 As typeRemHeaderCNAB240, objCarteiraCobrador As ClassCarteiraCobrador) As Long

On Error GoTo Erro_Obter_Dados_Header_Lote

    'Carrega os campos do cabeçalho
    With tRegHeaderBancoIndustrial240

        .sCodBancoComp = "604"
        .sLoteServico = FormataCpoNum(giNumLotes, 4)
        .sRegHeaderLote = "1"
        .sTipoOperacao = "R"
        .sTipoServico = "01"
        .sFormaLancamento = FormataCpoNum("", 2)
        .sNVersaoLayLote = "030"
        .sBrancosCNAB1 = FormataCpoTexto("", 1)

        If Len(gobjFilialEmpresa.sCgc) = 14 Then
            .sTipoInscEmpresa = "2"
        Else
            .sTipoInscEmpresa = "1"
        End If

        .sCodInsEmpresa = FormataCpoNum(gobjFilialEmpresa.sCgc, 15)
        .sCodConvBanco = FormataCpoTexto(Obter_Info_Cobrador(gobjCobrancaEletronica.objCobradorCNABInfo.colInformacoes, BANCOINDUSTRIAL_INFO_CONVENIO), 20)

        .sAgMantConta = FormataCpoNum(gobjCobrancaEletronica.objCCI.sAgencia, 5)
        .sDVAgencia = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVAgencia, 1)
        .sNumContaCorrente = FormataCpoNum(gobjCobrancaEletronica.objCCI.sNumConta, 12)
        .sDVConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVNumConta, 1)
        .sDVAgenciaConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVAgConta, 1)

        .sNomeEmpresa = FormataCpoTexto(gsNomeEmpresa, 30)

        .sMensagem1 = FormataCpoTexto("", 40)
        .sMensagem2 = FormataCpoTexto("", 40)
        .sNumRemessa = FormataCpoNum(gobjCobrancaEletronica.objCobrador.lCNABProxSeqArqCobr, 8)
        .sDataGravRem = Format(gdtDataHoje, "DDMMYYYY")
        .sDataCredito = FormataCpoNum(0, 8)
        .sBrancosCNAB2 = FormataCpoTexto("", 33)

    End With

    Obter_Dados_Header_Lote = SUCESSO

    Exit Function

Erro_Obter_Dados_Header_Lote:

    Obter_Dados_Header_Lote = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147168)

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

Private Function Montar_RegistroDetalhe_BancoIndustrial(sRegDetalheSegP As String, sRegDetalheSegQ As String, sRegDetalheSegR As String, objOcorrRemParcRec As ClassOcorrRemParcRec, objCarteiraCobrador As ClassCarteiraCobrador) As Long
'Monta o registro

Dim tRegDetalheBancoIndustrial240SegP As typeRemDetalheCNAB240SegP
Dim tRegDetalheBancoIndustrial240SegQ As typeRemDetalheCNAB240SegQ
Dim tRegDetalheBancoIndustrial240SegR As typeRemDetalheCNAB240SegR
Dim lErro As Long
Dim objParcelaRec As New ClassParcelaReceber

On Error GoTo Erro_Montar_RegistroDetalhe_BancoIndustrial

    gobjCobrancaEletronica.lQuantTitulos = gobjCobrancaEletronica.lQuantTitulos + 1

    objParcelaRec.lNumIntDoc = objOcorrRemParcRec.lNumIntParc

    'Busca os dados da parcela a receber no BD
    lErro = CF("CobrancaEletronica_Obter_DadosParcelaRec", objParcelaRec, gobjCobrancaEletronica)
    If lErro <> SUCESSO Then Error 51817

    gobjCobrancaEletronica.lSequencialRegistro = gobjCobrancaEletronica.lSequencialRegistro + 1

    'Obtem os dados do registro de detalhe SegmentoP
    lErro = Obter_Dados_DetalheSegP(tRegDetalheBancoIndustrial240SegP, objOcorrRemParcRec, objCarteiraCobrador, objParcelaRec)
    If lErro <> SUCESSO Then Error 51816

    With tRegDetalheBancoIndustrial240SegP
        'Forma o registro de detalhe
        sRegDetalheSegP = .sCodBancoComp & .sLoteServico & .sRegDetalhe & .sNumSequencialReg & .sCodSegRegDetalhe & .sBrancosCNAB1 & .sCodMovimento
        sRegDetalheSegP = sRegDetalheSegP & FormataCpoTexto(Obter_Info_Cobrador(gobjCobrancaEletronica.objCobradorCNABInfo.colInformacoes, BANCOINDUSTRIAL_INFO_CONVENIO), 20) 'o convenio substitui a conta
        sRegDetalheSegP = sRegDetalheSegP & .sDirecionamentoCobranca & .sModalidadeCorrespondentes & .sBrancosCNAB3 & .sModalidadeBcoCobranca & .sIdentTituloBanco & .sCodCarteira & .sFormaCadTituloBanco
        sRegDetalheSegP = sRegDetalheSegP & .sTipoDocumento & .sIdentEmissaoBloq & .sIdentDistrib & .sNumDocCobranca & .sDataVencTitulo & .sValorNominalTitulo & .sAgEncCobranca & .sDVAgencia & .sEspecieTitulo
        sRegDetalheSegP = sRegDetalheSegP & .sIdentTituloAceite & .sDataEmissaoTitulo & .sCodJurosMora & .sDataJurosMora & .sJurosMoraDiaTaxa & .sCodDesconto1 & .sDataDesconto1
        sRegDetalheSegP = sRegDetalheSegP & .sValorPerConcedido & .sValorIOF & .sValorAbatimento & .sIdentTituloEmp & .sCodProtesto & .sNumDiasProtesto
        sRegDetalheSegP = sRegDetalheSegP & .sCodBaixaDevolucao & .sNumDiasBaixaDevol & .sCodMoeda & .sNumContratoOpCred & .sBrancosCNAB2

    End With

    gobjCobrancaEletronica.lSequencialRegistro = gobjCobrancaEletronica.lSequencialRegistro + 1
    
    'Obtem os dados do registro de detalhe SegmentoQ
    lErro = Obter_Dados_DetalheSegQ(tRegDetalheBancoIndustrial240SegQ, objOcorrRemParcRec, objCarteiraCobrador, objParcelaRec)
    If lErro <> SUCESSO Then Error 51816
    
    With tRegDetalheBancoIndustrial240SegQ
        'Forma o registro de detalhe
        sRegDetalheSegQ = .sCodBancoComp & .sLoteServico & .sRegDetalhe & .sNumSequencialReg & .sCodSegRegDetalhe & .sBrancosCNAB1 & .sCodMovimento
        sRegDetalheSegQ = sRegDetalheSegQ & .sTipoInscricao & .sNumInscricao & .sNome & .sEndereco & .sBairro & .sCEP & .sSufixoCEP & .sCidade & .sUnidFederacao
        sRegDetalheSegQ = sRegDetalheSegQ & .sTipoInscricaoSacAval & .sNumInscricaoSacAval & .sNomeSacadorAvalista & .sBancoCompensacao & .sNossoNumBancoCorresp & .sBrancosCNAB2
        
    End With
    
    gobjCobrancaEletronica.lSequencialRegistro = gobjCobrancaEletronica.lSequencialRegistro + 1
        
    'Obtem os dados do registro de detalhe SegmentoR
    lErro = Obter_Dados_DetalheSegR(tRegDetalheBancoIndustrial240SegR, objOcorrRemParcRec, objCarteiraCobrador, objParcelaRec)
    If lErro <> SUCESSO Then Error 51816
    
    With tRegDetalheBancoIndustrial240SegR
        'Forma o registro de detalhe
        sRegDetalheSegR = .sCodBancoComp & .sLoteServico & .sRegDetalhe & .sNumSequencialReg & .sCodSegRegDetalhe & .sBrancosCNAB1 & .sCodMovimento
        sRegDetalheSegR = sRegDetalheSegR & .sCodDesconto2 & .sDataDesconto2 & .sValorPerConcedido2 & .sCodDesconto3 & .sDataDesconto3 & .sValorPerConcedido3
        sRegDetalheSegR = sRegDetalheSegR & .sCodMulta & .sDataMulta & .sValorPerAplicado & .sInfoBancoSacado & .sMensagem3 & .sMensagem4 & .sBrancosCNAB3
        sRegDetalheSegR = sRegDetalheSegR & .sCodOcorrenciaSac & .sCodBancoContaDeb & .sCodAgContaDeb & .sCodAgDVDeb & .sCodContaDVDeb & .sCodAgContaDVDeb & .sBrancosCNAB2

    End With
    
    Montar_RegistroDetalhe_BancoIndustrial = SUCESSO

    Exit Function

Erro_Montar_RegistroDetalhe_BancoIndustrial:

     Montar_RegistroDetalhe_BancoIndustrial = Err

    Select Case Err

        Case 51816, 51817

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147169)

    End Select

    Exit Function

End Function

Private Function Obter_Dados_DetalheSegP(tRegDetalheBancoIndustrial240SegP As typeRemDetalheCNAB240SegP, objOcorrRemParcRem As ClassOcorrRemParcRec, objCarteiraCobrador As ClassCarteiraCobrador, objParcelaRec As ClassParcelaReceber) As Long
'Obtem os dados que são utilizados para preencher o registro de detalhe

Dim dValorDesconto As Double
Dim iDias As Integer

On Error GoTo Erro_Obter_Dados_DetalheSegP

    With tRegDetalheBancoIndustrial240SegP
    
        .sCodBancoComp = "604"
        .sLoteServico = FormataCpoNum(giNumLotes, 4)
        .sRegDetalhe = "3"
        .sNumSequencialReg = FormataCpoNum(gobjCobrancaEletronica.lSequencialRegistro, 5)
        .sCodSegRegDetalhe = "P"
        .sBrancosCNAB1 = FormataCpoTexto("", 1)
        .sCodMovimento = FormataCpoNum(objOcorrRemParcRem.iCodOcorrencia, 2)
        .sAgMantConta = FormataCpoNum(gobjCobrancaEletronica.objCCI.sAgencia, 5)
        .sDVAgMantConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVAgencia, 1)
        .sNumContaCorrente = FormataCpoNum(gobjCobrancaEletronica.objCCI.sNumConta, 12)
        .sDVConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVNumConta, 1)
        .sDVAgenciaConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVAgConta, 1)
        .sDirecionamentoCobranca = "0" 'Cobrança a classificar pelo CEP. Banco se encarrega.
        .sModalidadeCorrespondentes = "000" '???
        .sBrancosCNAB3 = "00"
        .sModalidadeBcoCobranca = FormataCpoNum(objCarteiraCobrador.iNumCarteiraNoBanco, 3)
        
        If objCarteiraCobrador.iGeraNossoNumero Then
            .sIdentTituloBanco = FormataCpoNum("", 11)
        Else
            .sIdentTituloBanco = FormataCpoNum(Calcula_NossoNumero_BancoIndustrial(objCarteiraCobrador, .sAgMantConta, .sModalidadeBcoCobranca), 11)
        End If
        
        If objCarteiraCobrador.iCodCarteiraCobranca = CARTEIRA_SIMPLES Then
            .sCodCarteira = FormataCpoNum(1, 1)
        ElseIf objCarteiraCobrador.iCodCarteiraCobranca = CARTEIRA_VINCULADA Then
            .sCodCarteira = FormataCpoNum(2, 1)
        ElseIf objCarteiraCobrador.iCodCarteiraCobranca = CARTEIRA_CAUCIONADA Then
            .sCodCarteira = FormataCpoNum(3, 1)
        ElseIf objCarteiraCobrador.iCodCarteiraCobranca = CARTEIRA_DESCONTADA Then
            .sCodCarteira = FormataCpoNum(4, 1)
        Else
            .sCodCarteira = FormataCpoNum(objCarteiraCobrador.sCodCarteiraNoBanco, 1)
        End If
        
        .sFormaCadTituloBanco = FormataCpoNum(IIf(objCarteiraCobrador.iComRegistro = 0, "2", "1"), 1)
        .sTipoDocumento = "2" 'Escritural
        .sIdentEmissaoBloq = FormataCpoNum(IIf(objCarteiraCobrador.iImprimeBoleta = 0, "1", "2"), 1)
        .sIdentDistrib = "1" 'Banco Distribui
        .sNumDocCobranca = FormataCpoTexto(gobjCobrancaEletronica.lNumTitulo & SEPARADOR & objParcelaRec.iNumParcela, 15)
        
        .sDataVencTitulo = Format(objParcelaRec.dtDataVencimento, "DDMMYYYY")
        .sValorNominalTitulo = FormataCpoValor(objOcorrRemParcRem.dValorCobrado, 15)
        .sAgEncCobranca = FormataCpoNum("", 5)
        .sDVAgencia = FormataCpoNum("0", 1)
        
        .sEspecieTitulo = FormataCpoNum(2, 2) 'Duplicata Mercantil
        
        If objParcelaRec.iAceite = 1 Then
            .sIdentTituloAceite = "A"
        Else
            .sIdentTituloAceite = "N"
        End If

        .sDataEmissaoTitulo = Format(gobjCobrancaEletronica.dtDataEmissaoTitulo, "DDMMYYYY")
        .sCodJurosMora = "1"
        .sDataJurosMora = Format(objParcelaRec.dtDataVencimento + 1, "DDMMYYYY")
        .sJurosMoraDiaTaxa = FormataCpoValor(Round(objOcorrRemParcRem.dValorCobrado * gobjCobrancaEletronica.dPercJurosDiario, 2), 15)
        .sCodDesconto1 = FormataCpoNum(objParcelaRec.iDesconto1Codigo, 1)
        
        If objParcelaRec.dtDesconto1Ate <> DATA_NULA Then
                .sDataDesconto1 = Format(objParcelaRec.dtDesconto1Ate, "DDMMYYYY")
            Else
                .sDataDesconto1 = FormataCpoNum(0, 8)
        End If
        
        'Calcula o valor do desconto quando for do tipo percentual
        dValorDesconto = 0
        If objParcelaRec.iDesconto1Codigo = PERC_ANT_DIA Or objParcelaRec.iDesconto1Codigo = PERC_ANT_DIA_UTIL Or objParcelaRec.iDesconto1Codigo = Percentual Then
            dValorDesconto = StrParaDbl(Format(objParcelaRec.dDesconto1Valor * 100))
        Else
            If objParcelaRec.iDesconto1Codigo = VALOR_FIXO Then
                dValorDesconto = objParcelaRec.dDesconto1Valor
            End If
        End If
        
        .sValorPerConcedido = FormataCpoValor(dValorDesconto, 15)
        .sValorIOF = FormataCpoValor(0, 15)
        .sValorAbatimento = FormataCpoValor(0, 15)
        .sIdentTituloEmp = FormataCpoTexto(objOcorrRemParcRem.lNumIntDoc, 25)
        
        If objOcorrRemParcRem.iCodOcorrencia = 9 Then
            'Dias Corridos
            If objOcorrRemParcRem.iInstrucao1 = 81 Then
                .sCodProtesto = "1"
                .sNumDiasProtesto = FormataCpoNum(objOcorrRemParcRem.iDiasDeProtesto1, 2)
            
            'Dias Úteis
            ElseIf objOcorrRemParcRem.iInstrucao1 = 82 Then
                .sCodProtesto = "2"
                .sNumDiasProtesto = FormataCpoNum(objOcorrRemParcRem.iDiasDeProtesto1, 2)
            
            'Dias Corridos
            ElseIf objOcorrRemParcRem.iInstrucao2 = 81 Then
                .sCodProtesto = "1"
                .sNumDiasProtesto = FormataCpoNum(objOcorrRemParcRem.iDiasDeProtesto2, 2)
            
            'Dias Úteis
            ElseIf objOcorrRemParcRem.iInstrucao1 = 82 Then
                .sCodProtesto = "2"
                .sNumDiasProtesto = FormataCpoNum(objOcorrRemParcRem.iDiasDeProtesto2, 2)
            
            Else
                .sCodProtesto = "3"
                .sNumDiasProtesto = FormataCpoNum(0, 2)
            End If
            
        Else
            .sCodProtesto = "3"
            .sNumDiasProtesto = FormataCpoNum(0, 2)
        End If
        
        If (objOcorrRemParcRem.iInstrucao1 >= 17 And objOcorrRemParcRem.iInstrucao1 <= 29) Or _
           (objOcorrRemParcRem.iInstrucao2 >= 17 And objOcorrRemParcRem.iInstrucao2 <= 29) Or _
            objOcorrRemParcRem.iInstrucao1 = 92 Or objOcorrRemParcRem.iInstrucao1 = 92 Then
        
            .sCodBaixaDevolucao = "1"
            If (objOcorrRemParcRem.iInstrucao1 >= 17 And objOcorrRemParcRem.iInstrucao1 < 29) Then
                
                iDias = 5 * (objOcorrRemParcRem.iInstrucao1 - 17)
            
            ElseIf (objOcorrRemParcRem.iInstrucao2 >= 17 And objOcorrRemParcRem.iInstrucao2 < 29) Then
                
                iDias = 5 * (objOcorrRemParcRem.iInstrucao2 - 17)
            
            ElseIf objOcorrRemParcRem.iInstrucao2 = 29 Or objOcorrRemParcRem.iInstrucao2 = 29 Then
                
                iDias = 90
            
            ElseIf objOcorrRemParcRem.iInstrucao1 = 92 Then
                
                iDias = objOcorrRemParcRem.iDiasDeProtesto1
            
            Else
                
                iDias = objOcorrRemParcRem.iDiasDeProtesto2
            
            End If
            
        Else
        
            .sCodBaixaDevolucao = "2"
            .sNumDiasBaixaDevol = FormataCpoNum(0, 3)
        End If
        
                
        .sCodMoeda = "09"
        .sNumContratoOpCred = FormataCpoNum(0, 10)
        .sBrancosCNAB2 = FormataCpoTexto("", 1)
        
                   
    End With

    Obter_Dados_DetalheSegP = SUCESSO

    Exit Function

Erro_Obter_Dados_DetalheSegP:

Obter_Dados_DetalheSegP = Err

    Select Case Err


        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147170)

    End Select

    Exit Function

End Function


Private Function Obter_Dados_DetalheSegQ(tRegDetalheBancoIndustrial240SegQ As typeRemDetalheCNAB240SegQ, objOcorrRemParcRem As ClassOcorrRemParcRec, objCarteiraCobrador As ClassCarteiraCobrador, objParcelaRec As ClassParcelaReceber) As Long
'Obtem os dados que são utilizados para preencher o registro de detalhe

Dim lErro As Long
Dim objEndereco As New ClassEndereco

On Error GoTo Erro_Obter_Dados_DetalheSegQ

    lErro = CF("CobrancaEletronica_Obter_Endereco", gobjCobrancaEletronica, objEndereco)
    If lErro <> SUCESSO Then Error 51818
        
    With tRegDetalheBancoIndustrial240SegQ
           
        .sCodBancoComp = "604"
        .sLoteServico = FormataCpoNum(giNumLotes, 4)
        .sRegDetalhe = "3"
        .sNumSequencialReg = FormataCpoNum(gobjCobrancaEletronica.lSequencialRegistro, 5)
        .sCodSegRegDetalhe = "Q"
        .sBrancosCNAB1 = FormataCpoTexto("", 1)
        .sCodMovimento = FormataCpoNum(objOcorrRemParcRem.iCodOcorrencia, 2)
    
        If Len(gobjCobrancaEletronica.sCGCSacado) = 14 Then
            .sTipoInscricao = "2"
        Else
            .sTipoInscricao = "1"
        End If
        
        .sNumInscricao = FormataCpoNum(gobjCobrancaEletronica.sCGCSacado, 15)
        .sNome = FormataCpoTexto(gobjCobrancaEletronica.sNomeSacado, 40)
        .sEndereco = FormataCpoTexto(Trim(objEndereco.sEndereco), 40)
        .sBairro = FormataCpoTexto(Trim(objEndereco.sBairro), 15)
        .sCEP = Formata_Cep(objEndereco.sCEP, 8) 'Incluido o sufixo.
        .sCidade = FormataCpoTexto(objEndereco.sCidade, 15)
        .sUnidFederacao = FormataCpoTexto(objEndereco.sSiglaEstado, 2)
        .sTipoInscricaoSacAval = FormataCpoNum(0, 1)
        .sNumInscricaoSacAval = FormataCpoNum(0, 15)
        .sNomeSacadorAvalista = FormataCpoTexto("", 40)
        .sBancoCompensacao = FormataCpoNum(0, 3)
        .sNossoNumBancoCorresp = FormataCpoTexto("", 20)
        .sBrancosCNAB2 = FormataCpoTexto("", 8)
        
    End With
    
    Obter_Dados_DetalheSegQ = SUCESSO

    Exit Function
    
Erro_Obter_Dados_DetalheSegQ:

Obter_Dados_DetalheSegQ = Err
    
    Select Case Err

        Case 51818
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147171)

    End Select

    Exit Function

End Function


Private Function Obter_Dados_DetalheSegR(tRegDetalheBancoIndustrial240SegR As typeRemDetalheCNAB240SegR, objOcorrRemParcRem As ClassOcorrRemParcRec, objCarteiraCobrador As ClassCarteiraCobrador, objParcelaRec As ClassParcelaReceber) As Long
'Obtem os dados que são utilizados para preencher o registro de detalhe

Dim dValorDesconto As Double

On Error GoTo Erro_Obter_Dados_DetalheSegR

    With tRegDetalheBancoIndustrial240SegR
       
        .sCodBancoComp = "604"
        .sLoteServico = FormataCpoNum(giNumLotes, 4)
        .sRegDetalhe = "3"
        .sNumSequencialReg = FormataCpoNum(gobjCobrancaEletronica.lSequencialRegistro, 5)
        .sCodSegRegDetalhe = "R"
        .sBrancosCNAB1 = FormataCpoTexto("", 1)
        .sCodMovimento = FormataCpoNum(objOcorrRemParcRem.iCodOcorrencia, 2)
        .sCodDesconto2 = FormataCpoNum(objParcelaRec.iDesconto2Codigo, 1)
        
        If objParcelaRec.dtDesconto2Ate <> DATA_NULA Then
            .sDataDesconto2 = Format(objParcelaRec.dtDesconto2Ate, "DDMMYYYY")
        Else
            .sDataDesconto2 = FormataCpoNum(0, 8)
        End If
        
        'Calcula o valor do desconto quando for do tipo percentual
        dValorDesconto = 0
        If objParcelaRec.iDesconto2Codigo = PERC_ANT_DIA Or objParcelaRec.iDesconto1Codigo = PERC_ANT_DIA_UTIL Or objParcelaRec.iDesconto2Codigo = Percentual Then
            dValorDesconto = StrParaDbl(Format(objParcelaRec.dDesconto2Valor * 100))
        Else
            If objParcelaRec.iDesconto2Codigo = VALOR_FIXO Then
                dValorDesconto = objParcelaRec.dDesconto2Valor
            End If
        End If
        
        .sValorPerConcedido2 = FormataCpoValor(dValorDesconto, 15)
        .sCodDesconto3 = FormataCpoNum(objParcelaRec.iDesconto3Codigo, 1)
        
        If objParcelaRec.dtDesconto3Ate <> DATA_NULA Then
            .sDataDesconto3 = Format(objParcelaRec.dtDesconto3Ate, "DDMMYYYY")
        Else
            .sDataDesconto3 = FormataCpoNum(0, 8)
        End If
        
        'Calcula o valor do desconto quando for do tipo percentual
        dValorDesconto = 0
        If objParcelaRec.iDesconto3Codigo = PERC_ANT_DIA Or objParcelaRec.iDesconto1Codigo = PERC_ANT_DIA_UTIL Or objParcelaRec.iDesconto3Codigo = Percentual Then
            dValorDesconto = StrParaDbl(Format(objParcelaRec.dDesconto3Valor * 100))
        Else
            If objParcelaRec.iDesconto3Codigo = VALOR_FIXO Then
                dValorDesconto = objParcelaRec.dDesconto3Valor
            End If
        End If
        
        .sValorPerConcedido3 = FormataCpoValor(dValorDesconto, 15)
        .sCodMulta = IIf(gobjCobrancaEletronica.dPercMulta = 0, "0", "2") '1: valor fixo, 2:Percentual
        .sDataMulta = Format(objParcelaRec.dtDataVencimento + 1, "DDMMYYYY")
        .sValorPerAplicado = FormataCpoValor(gobjCobrancaEletronica.dPercMulta * 10000, 15)
        .sInfoBancoSacado = FormataCpoTexto("", 10)
        .sMensagem3 = FormataCpoTexto("", 40)
        .sMensagem4 = FormataCpoTexto("", 40)
        .sBrancosCNAB3 = FormataCpoTexto("", 20)
        .sCodOcorrenciaSac = FormataCpoTexto("", 8)
        .sCodBancoContaDeb = FormataCpoTexto("", 3)
        .sCodAgContaDeb = FormataCpoTexto("", 5)
        .sCodAgDVDeb = FormataCpoTexto("", 1)
        .sCodContaDVDeb = FormataCpoTexto("", 13) '12 + 1
        .sCodAgContaDVDeb = FormataCpoTexto("", 1)
        .sBrancosCNAB2 = FormataCpoTexto("", 10) '1 + 9
        
    End With
    
    Obter_Dados_DetalheSegR = SUCESSO

    Exit Function
    
Erro_Obter_Dados_DetalheSegR:

Obter_Dados_DetalheSegR = Err
    
    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147172)

    End Select

    Exit Function

End Function

Private Function Montar_RegistroTrailerLote_BancoIndustrial(sRegTrailer As String) As Long
'Monta o registro

Dim tRegTrailerBancoIndustrial240 As typeRemTrailerCNAB240

On Error GoTo Erro_Montar_RegistroTrailerLote_BancoIndustrial
    
    With tRegTrailerBancoIndustrial240
        .sCodBancoComp = "604"
        .sLoteServico = FormataCpoNum(giNumLotes, 4)
        .sRegTrailer = "5"
        .sBrancosCNAB1 = FormataCpoTexto("", 9)
        .sQuantRegLote = FormataCpoNum((gobjCobrancaEletronica.lSequencialRegistro + 2), 6)
        .sQuantTitCobS = FormataCpoNum(0, 6) 'Informação que virão no Arq de Retorno
        .sValTotalTitCartS = FormataCpoNum(0, 17) 'Informação que virão no Arq de Retorno
        .sQuantTitCobV = FormataCpoNum(0, 6) 'Informação que virão no Arq de Retorno
        .sValTotalTitCartV = FormataCpoNum(0, 17) 'Informação que virão no Arq de Retorno
        .sQuantTitCobC = FormataCpoNum(0, 6) 'Informação que virão no Arq de Retorno
        .sValTotalTitCartC = FormataCpoNum(0, 17) 'Informação que virão no Arq de Retorno
        .sQuantTitCobD = FormataCpoNum(0, 6) 'Informação que virão no Arq de Retorno
        .sValTotalTitCartD = FormataCpoNum(0, 17) 'Informação que virão no Arq de Retorno
        .sBrancosCNAB2 = FormataCpoTexto("", 117 + 8)
    End With
    
    With tRegTrailerBancoIndustrial240
        'Monta o Registro de trailer
        sRegTrailer = .sCodBancoComp & .sLoteServico & .sRegTrailer & .sBrancosCNAB1
        sRegTrailer = sRegTrailer & .sQuantRegLote & .sQuantTitCobS & .sValTotalTitCartS
        sRegTrailer = sRegTrailer & .sQuantTitCobV & .sValTotalTitCartV & .sQuantTitCobC
        sRegTrailer = sRegTrailer & .sValTotalTitCartC & .sQuantTitCobD & .sValTotalTitCartD & .sBrancosCNAB2
    
    End With
    
    Montar_RegistroTrailerLote_BancoIndustrial = SUCESSO

    Exit Function

Erro_Montar_RegistroTrailerLote_BancoIndustrial:

    Montar_RegistroTrailerLote_BancoIndustrial = Err

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147173)

    End Select

    Exit Function

End Function

Private Function Montar_RegistroTrailerArq_BancoIndustrial(sRegTrailer As String) As Long
'Monta o registro

Dim tRegTrailerBancoIndustrial240 As typeTrailerArqCNAB240

On Error GoTo Erro_Montar_RegistroTrailerArq_BancoIndustrial

    glNumRegistros = glNumRegistros + 1
    
    With tRegTrailerBancoIndustrial240
        .sNumeroBanco = "604"
        .sLoteServico = "9999"
        .sIdentificacaoReg = "9"
        .sReservadoCNAB1 = "         "
        .sQtdeLotes = FormataCpoNum(giNumLotes, 6)
        .sQtdeRegistros = FormataCpoNum(glNumRegistros, 6)
        .sQtdeContasConciliacao = FormataCpoNum(0, 6)
        .sReservadoCNAB2 = FormataCpoTexto("", 205)
    
        sRegTrailer = .sNumeroBanco & .sLoteServico & .sIdentificacaoReg & .sReservadoCNAB1 & .sQtdeLotes & .sQtdeRegistros & .sQtdeContasConciliacao & .sReservadoCNAB2
    
    End With
        
    Montar_RegistroTrailerArq_BancoIndustrial = SUCESSO

    Exit Function

Erro_Montar_RegistroTrailerArq_BancoIndustrial:

    Montar_RegistroTrailerArq_BancoIndustrial = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147174)

    End Select

    Exit Function

End Function

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147175)
     
    End Select
     
    Exit Function

End Function

Function Processar_ArquivoRetorno_Cobranca_BancoIndustrial(objCobrancaEletronica As ClassCobrancaEletronica) As Long
'Faz o processamento do arquivo de retorno

Dim lErro As Long
Dim iIndice As Integer
Dim sRegistro As String
Dim iCodRegistro As Integer, objLoteServico As New ClassCNABLoteServico
Dim objBorderoRetCobr As New ClassBorderoRetCobr
Dim lComando As Long, tRegDetalhe As typeRetCobrDetalheCNAB240
Dim lComando1 As Long, tRegControle As typeRegControle
Dim bDesistiu As Boolean, dValorMov As Double
Dim objBorderoRetCobrCart As ClassBorderoRetCobrCart

On Error GoTo Erro_Processar_ArquivoRetorno_Cobranca_BancoIndustrial

    objBorderoRetCobr.iBanco = 604
    
    'le as carteiras de cobranca do cobrador
    lErro = CF("CarteirasCobrador_Le", objCobrancaEletronica.objCobrador, objBorderoRetCobr.colCartCobr)
    If lErro <> SUCESSO Then Error 32265
    
    'Abre os comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 51824

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 51825

    'Abre o arquivo de retorno
    Open objCobrancaEletronica.sNomeArquivoRetorno For Input As #2
    
    'Até chegar ao fim do arquivo
    Do While Not EOF(2)
    
        'Busca o próximo registro do arquivo (na 1a vez vai ser o de header)
        Line Input #2, sRegistro
    
        If Len(Trim(sRegistro)) <> 0 Then
        
            'Pega o "controle" do registro
            Call RegArqCNAB_Extrai_Controle(sRegistro, tRegControle, objCobrancaEletronica)
            
            Select Case tRegControle.sIdentificacaoReg
                
                Case "0" 'Header de Arquivo
                
                    'Processa os dados contidos no Header do arquivo
                    lErro = Processa_Registro_Header_Arq(sRegistro, objBorderoRetCobr, objCobrancaEletronica.objCobrador, objCobrancaEletronica.sNomeArquivoRetorno)
                    If lErro <> SUCESSO Then Error 51826
                
                Case "1" 'header de lote
                    'Processa os dados contidos no Header do lote
                    lErro = Processa_Registro_Header_Lote(objLoteServico, sRegistro, objBorderoRetCobr, objCobrancaEletronica.objCobrador)
                    If lErro <> SUCESSO Then Error 51826
                
                Case "3" 'Detalhe ("Transação")
                    'Processa os dados contidos no registro de detalhe
                    lErro = Processa_Registro_Detalhe(objLoteServico, sRegistro, objBorderoRetCobr, objCobrancaEletronica, tRegDetalhe, lComando, lComando1)
                    If lErro <> SUCESSO Then Error 51827
            
                Case "5" 'Trailer de lote
                    'Processa os dados do trailer de Lote
                    lErro = Processa_Registro_Trailer_Lote(objLoteServico, sRegistro)
                    If lErro <> SUCESSO Then Error 51828
                
                Case "9" 'Trailer de arquivo
                    'Processa os dados do trailer do arquivo
                    lErro = Processa_Registro_Trailer_Arq(sRegistro)
                    If lErro <> SUCESSO Then Error 51828
            
                Case Else
                    Error 51831
                    
            End Select
        
        End If
        
    Loop
        
    Close #2
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    'Grava os dados recolhidos durante o processamento do arquivo de retorno
    lErro = CF("BorderoRetCobr_ConfirmaGrava", objBorderoRetCobr)
    If lErro <> SUCESSO Then Error 51829
    
    Processar_ArquivoRetorno_Cobranca_BancoIndustrial = SUCESSO
    
    Exit Function
    
Erro_Processar_ArquivoRetorno_Cobranca_BancoIndustrial:

    Processar_ArquivoRetorno_Cobranca_BancoIndustrial = Err
    
    Select Case Err
    
        Case 51823
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 51824, 51825
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 32260, 32265, 51826, 51827, 51828, 51829
        
        Case 51831
            Call Rotina_Erro(vbOKOnly, "ERRO_FORMATO_ARQUIVO_INCORRETO", Err, objCobrancaEletronica.sNomeArquivoRetorno)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147176)
    
    End Select

    Close #2
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Sub RegArqCNAB_Extrai_Controle(sRegistro As String, tRegControle As typeRegControle, objCobrancaEletronica As ClassCobrancaEletronica)

    With tRegControle
    
        .sNumeroBanco = Mid(sRegistro, 1, 3)
        .sLoteServico = Mid(sRegistro, 4, 4)
        .sIdentificacaoReg = Mid(sRegistro, 8, 1)
        
    End With
    
End Sub

Private Function Processa_Registro_Header_Arq(sRegistro As String, objBorderoRetCobr As ClassBorderoRetCobr, objCobrador As ClassCobrador, sNomeArquivo As String) As Long

On Error GoTo Erro_Processa_Registro_Header_Arq

    'Preenche o objBorderoRetCobr
    objBorderoRetCobr.dtDataRecepcao = gdtDataHoje
    objBorderoRetCobr.iCobrador = objCobrador.iCodigo
    objBorderoRetCobr.sNomeArq = sNomeArquivo
    
    '??? verificar se o bco é o do cobrador
    
    Processa_Registro_Header_Arq = SUCESSO
    
    Exit Function

Erro_Processa_Registro_Header_Arq:

    Processa_Registro_Header_Arq = Err

    Select Case Err
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147177)
            
    End Select
    
    Exit Function

End Function

Function Processa_Registro_Header_Lote(objLoteServico As ClassCNABLoteServico, sRegistro As String, objBorderoRetCobr As ClassBorderoRetCobr, objCobrador As ClassCobrador) As Long

Dim lErro As Long

On Error GoTo Erro_Processa_Registro_Header_Lote

    With objLoteServico
        .sLote = Mid(sRegistro, 4, 4)
        .sTipoOperacao = Mid(sRegistro, 9, 1)
        .sTipoServico = Mid(sRegistro, 10, 2)
        .sFormaLcto = Mid(sRegistro, 12, 2)
    End With
    
    'se for retorno de cobranca
    If objLoteServico.sTipoOperacao = "T" And objLoteServico.sTipoServico = "01" Then
    
        'pode obter identificacao da empresa, num ret, data de credito, etc
    
    End If
        
    Processa_Registro_Header_Lote = SUCESSO
     
    Exit Function
    
Erro_Processa_Registro_Header_Lote:

    Processa_Registro_Header_Lote = Err
     
    Select Case Err
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147178)
     
    End Select
     
    Exit Function

End Function


Private Function Processa_Registro_Detalhe(objLoteServico As ClassCNABLoteServico, sRegistro As String, objBorderoRetCobr As ClassBorderoRetCobr, objCobrancaEletronica As ClassCobrancaEletronica, tRegDetalhe As typeRetCobrDetalheCNAB240, lComando As Long, lComando1 As Long) As Long

Dim objOcorrRemParcRec As New ClassOcorrRemParcRec
Dim iCodOcorrencia As Integer, iTipoOcorrencia As Integer
Dim iCodCarteiraNoBanco As Integer
Dim lErro As Long
Dim objParcelaRec As New ClassParcelaReceber
Dim dValorPrincipal As Double
Dim dtDataVencimento As Date
Dim dtDataCredito As Date, objDetRetCobr As New ClassDetRetCobr
Dim objBorderoRetCobrCart As ClassBorderoRetCobrCart
Dim objInfoParcRec As ClassInfoParcRec

On Error GoTo Erro_Processa_Registro_Detalhe

    'se for retorno de cobranca
    If objLoteServico.sTipoOperacao = "T" And objLoteServico.sTipoServico = "01" Then
    
        'Extrai da variável lida no arquivo os dados do registro de detalhe
        lErro = Extrair_Detalhe_Retorno(tRegDetalhe, sRegistro)
        If lErro <> SUCESSO Then Error 51835
        
        'dependendo do tipo de segmento
        If Mid(sRegistro, 14, 1) = "U" Then
               
            'Recolhe o código da ocorrência
            iCodOcorrencia = StrParaInt(tRegDetalhe.tU030.sCodMovto)
            
            Select Case iCodOcorrencia
                
                'Quando ocorre baixa p/liquidacao
                Case 6
                    iTipoOcorrencia = RETCOBR_DET_BAIXA
                
                'cobranca de tarifas e custas
                Case 2, 12, 20, 23, 28
                    iTipoOcorrencia = RETCOBR_DET_TARIFAS
                                
                Case 3
                    iTipoOcorrencia = RETCOBR_DET_REJEITADO
                                
                Case Else
                    iTipoOcorrencia = RETCOBR_DET_IGNORAR
                    
            End Select
                
            If iTipoOcorrencia <> RETCOBR_DET_IGNORAR Then
            
                lErro = Prepara_CartCobrRet1(objCobrancaEletronica, objBorderoRetCobr, objBorderoRetCobrCart, objDetRetCobr, tRegDetalhe)
                If lErro <> SUCESSO Then Error 32257
                
                lErro = CF("Processa_Registro_Detalhe_Geral", objBorderoRetCobr, objCobrancaEletronica, iTipoOcorrencia, RETCOBR_TIPO_SEU_NUMERO2, objDetRetCobr, objBorderoRetCobrCart, lComando)
                If lErro <> SUCESSO Then Error 32356
                                          
            End If
            
        End If
    
    End If
    
    Processa_Registro_Detalhe = SUCESSO
    
    Exit Function

Erro_Processa_Registro_Detalhe:

    Processa_Registro_Detalhe = Err

    Select Case Err
    
        Case 32356, 51835, 51836, 51837, 51838
    
        Case 51839
            lErro = Rotina_Erro(vbOKOnly, "ERRO_OCORRREMPARCREC_NAO_CADASTRADA", Err, objOcorrRemParcRec.lNumIntDoc)
            
        Case 51840
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PARCELAREC_NAO_CADASTRADA", Err, objOcorrRemParcRec.lNumIntParc)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147179)
            
    End Select
    
    Exit Function

End Function


Private Function Extrair_Detalhe_Retorno(tRegDetalheBancoIndustrial As typeRetCobrDetalheCNAB240, sRegistro As String) As Long
'Extrai do sRegistro os dados do Detalhe do arquido de retorno

On Error GoTo Erro_Extrair_Detalhe_Retorno

    Select Case Mid(sRegistro, 14, 1)
    
        Case "T"
            
            With tRegDetalheBancoIndustrial.tT030
    
                .sNumeroBanco = Mid(sRegistro, 1, 3)
                .sLoteServico = Mid(sRegistro, 4, 4)
                .sIdentificacaoReg = Mid(sRegistro, 8, 1)
                .sNumSeqRegLote = Mid(sRegistro, 9, 5)
                .sCodSegmento = Mid(sRegistro, 14, 1)
                .sReservadoCNAB1 = Mid(sRegistro, 15, 1)
                .sCodMovto = Mid(sRegistro, 16, 2)
                .sAgencia = Mid(sRegistro, 18, 5)
                .sDVAgencia = Mid(sRegistro, 23, 1)
                .sContaCorrente = Mid(sRegistro, 24, 12)
                .sDVContaCorrente = Mid(sRegistro, 36, 1)
                .sDVAgenciaConta = Mid(sRegistro, 37, 1)
                .sNossoNumero = Mid(sRegistro, 38, 20) '1+3+2+3+11
                .sCodCarteira = Mid(sRegistro, 58, 1)
                .sNumDoc = Mid(sRegistro, 59, 15)
                .sVencimento = Mid(sRegistro, 74, 8)
                .sValorTitulo = Mid(sRegistro, 82, 15)
                .sBcoCobrRec = Mid(sRegistro, 97, 3)
                .sAgCobrRec = Mid(sRegistro, 100, 5)
                .sDVAgCobrRec = Mid(sRegistro, 105, 1)
                .sUsoEmpresa = Mid(sRegistro, 106, 25)
                .sCodMoeda = Mid(sRegistro, 131, 2)
                .sSacadoTipoInscr = Mid(sRegistro, 133, 1)
                .sSacadoNumInscr = Mid(sRegistro, 134, 15)
                .sSacadoNome = Mid(sRegistro, 149, 40)
                .sNumContrato = Mid(sRegistro, 189, 10)
                .sValorTarifa = Mid(sRegistro, 199, 15)
                .sMotivoOcorrencia = Mid(sRegistro, 214, 10)
                .sReservadoCNAB2 = Mid(sRegistro, 224, 17)
            
            End With
            
        Case "U"
        
            With tRegDetalheBancoIndustrial.tU030
    
                .sNumeroBanco = Mid(sRegistro, 1, 3)
                .sLoteServico = Mid(sRegistro, 4, 4)
                .sIdentificacaoReg = Mid(sRegistro, 8, 1)
                .sNumSeqRegLote = Mid(sRegistro, 9, 5)
                .sCodSegmento = Mid(sRegistro, 14, 1)
                .sReservadoCNAB1 = Mid(sRegistro, 15, 1)
                .sCodMovto = Mid(sRegistro, 16, 2)
                .sJurosMulta = Mid(sRegistro, 18, 15)
                .sDesconto = Mid(sRegistro, 33, 15)
                .sAbatimento = Mid(sRegistro, 48, 15)
                .sIOFRecolhido = Mid(sRegistro, 63, 15)
                .sValorPagoSacado = Mid(sRegistro, 78, 15)
                .sValorCreditado = Mid(sRegistro, 93, 15)
                .sOutrasDespesas = Mid(sRegistro, 108, 15)
                .sOutrosCreditos = Mid(sRegistro, 123, 15)
                .sDataOcorrencia = Mid(sRegistro, 138, 8)
                .sDataCredito = Mid(sRegistro, 146, 8)
                .sOcorrSacCodigo = Mid(sRegistro, 154, 4)
                .sOcorrSacData = Mid(sRegistro, 158, 8)
                .sOcorrSacValor = Mid(sRegistro, 166, 15)
                .sOcorrSacCompl = Mid(sRegistro, 181, 30)
                .sCodBcoCorresp = Mid(sRegistro, 211, 3)
                .sNossoNumBcoCorresp = Mid(sRegistro, 214, 20)
                .sReservadoCNAB2 = Mid(sRegistro, 234, 7)
            
            End With
    
    End Select
    
    Extrair_Detalhe_Retorno = SUCESSO

    Exit Function

Erro_Extrair_Detalhe_Retorno:

    Extrair_Detalhe_Retorno = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147180)
            
    End Select
    
    Exit Function

End Function


Private Function Prepara_CartCobrRet1(objCobrancaEletronica As ClassCobrancaEletronica, objBorderoRetCobr As ClassBorderoRetCobr, objBorderoRetCobrCart As ClassBorderoRetCobrCart, objDetRetCobr As ClassDetRetCobr, tRegDetalhe As typeRetCobrDetalheCNAB240) As Long
'codigo comum ao tratamento de registros detalhe de baixas, custas e tarifas
'preenche objBorderoRetCobrCart e objDetRetCobr

Dim lErro As Long

On Error GoTo Erro_Prepara_CartCobrRet1

    Call RegDetCobr_ObterInfo(objDetRetCobr, tRegDetalhe)
    
    objDetRetCobr.dtDataBaixa = Extrair_Data(tRegDetalhe.tU030.sDataOcorrencia)
    
    lErro = CF("Prepara_CartCobrRet", objCobrancaEletronica, objBorderoRetCobr, objBorderoRetCobrCart, objDetRetCobr)
    If lErro <> SUCESSO Then Error 32255
    
    Prepara_CartCobrRet1 = SUCESSO
    
    Exit Function
    
Erro_Prepara_CartCobrRet1:

    Prepara_CartCobrRet1 = Err
    
    Select Case Err
    
        Case 32255
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147181)
            
    End Select

    Exit Function

End Function

Private Sub RegDetCobr_ObterInfo(objDetRetCobr As ClassDetRetCobr, tRegDetalhe As typeRetCobrDetalheCNAB240)
'copia dados do type para o objeto

    'Recolhe alguns dados do registro de detalhe
    With objDetRetCobr
    
        .dValorRecebido = Extrair_Valor(tRegDetalhe.tU030.sValorPagoSacado)
        .dtDataVencimento = Extrair_Data(tRegDetalhe.tT030.sVencimento)
        .dtDataCredito = Extrair_Data(tRegDetalhe.tU030.sDataCredito)
        If .dtDataCredito = DATA_NULA Then
            .dtDataCredito = Extrair_Data(tRegDetalhe.tU030.sDataOcorrencia)
        End If
        .dValorJuros = Extrair_Valor(tRegDetalhe.tU030.sJurosMulta)
        .dValorDesconto = Extrair_Valor(tRegDetalhe.tU030.sDesconto)
        .sNossoNumero = Trim(tRegDetalhe.tT030.sNossoNumero)
        .iCodOcorrencia = StrParaInt(tRegDetalhe.tT030.sCodMovto)
        .iCodCarteiraNoBanco = StrParaInt(tRegDetalhe.tT030.sCodCarteira)
        .dValorTarifa = Extrair_Valor(tRegDetalhe.tT030.sValorTarifa)
        .sSeuNumero = tRegDetalhe.tT030.sNumDoc
        .sUsoDaEmpresa = tRegDetalhe.tT030.sUsoEmpresa
        .lNumSequencial = CLng(tRegDetalhe.tT030.sNumSeqRegLote)
        .dValorTitulo = Extrair_Valor(tRegDetalhe.tT030.sValorTitulo)
        .iCodOcorrencia1 = StrParaInt(Mid(tRegDetalhe.tT030.sMotivoOcorrencia, 1, 2))
        .iCodOcorrencia2 = StrParaInt(Mid(tRegDetalhe.tT030.sMotivoOcorrencia, 3, 2))
        .iCodOcorrencia3 = StrParaInt(Mid(tRegDetalhe.tT030.sMotivoOcorrencia, 5, 2))
        .iCodOcorrencia4 = StrParaInt(Mid(tRegDetalhe.tT030.sMotivoOcorrencia, 7, 2))
        .iCodOcorrencia5 = StrParaInt(Mid(tRegDetalhe.tT030.sMotivoOcorrencia, 9, 2))
        
    End With
    
End Sub

Sub RetCobrCart_Obtem(objDetRetCobr As ClassDetRetCobr, objBorderoRetCobr As ClassBorderoRetCobr, objBorderoRetCobrCart As ClassBorderoRetCobrCart)
'obtem dentro do bordero de retorno objBorderoRetCobrCart correspondente à carteira passada
'Se ainda nao houver em objBorderoRetCobr retorna nothing.

Dim dtDataCredito As Date, iCodCarteiraNoBanco

    iCodCarteiraNoBanco = objDetRetCobr.iCodCarteiraNoBanco
    dtDataCredito = IIf(objDetRetCobr.dtDataCredito <> DATA_NULA, objDetRetCobr.dtDataCredito, objBorderoRetCobr.dtDataCredito)
    
    Set objBorderoRetCobrCart = Nothing
    
    For Each objBorderoRetCobrCart In objBorderoRetCobr.colBorderoRetCobrCart
    
        If objBorderoRetCobrCart.iCodCarteiraNoBanco = iCodCarteiraNoBanco And objBorderoRetCobrCart.dtDataCredito = dtDataCredito Then Exit For
        
    Next
    
End Sub

Function Processa_Registro_Trailer_Lote(objLoteServico As ClassCNABLoteServico, sRegistro As String) As Long

Dim lErro As Long

On Error GoTo Erro_Processa_Registro_Trailer_Lote

    Processa_Registro_Trailer_Lote = SUCESSO
     
    Exit Function
    
Erro_Processa_Registro_Trailer_Lote:

    Processa_Registro_Trailer_Lote = Err
     
    Select Case Err
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147182)
     
    End Select
     
    Exit Function

End Function

Function Processa_Registro_Trailer_Arq(sRegistro As String) As Long

Dim lErro As Long

On Error GoTo Erro_Processa_Registro_Trailer_Arq

    Processa_Registro_Trailer_Arq = SUCESSO
     
    Exit Function
    
Erro_Processa_Registro_Trailer_Arq:

    Processa_Registro_Trailer_Arq = Err
     
    Select Case Err
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147183)
     
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
'Tenta extrair uma data no formato "DDMMYYYY" de vData.
'Caso não consiga, devolve DATA_NULA
    
Dim sData As String

    sData = (Mid(vData, 1, 2) & "/" & Mid(vData, 3, 2) & "/" & Mid(vData, 5, 4))
    
    If IsDate(sData) Then
        Extrair_Data = CDate(sData)
    Else
        Extrair_Data = DATA_NULA
    End If
    
    Exit Function

End Function

Sub Calcula_DV_NossoNumero_BancoIndustrial(sAgMantConta As String, sModalidadeBcoCobranca As String, sNossoNumero As String, sDVNossoNumero As String)

Dim aiNumeros(1 To 17) As Integer
Dim iFator As Integer
Dim iIndice As Integer
Dim iIndiceAtual As Integer
Dim iAcumulador As Integer

    
    iIndiceAtual = 0

    For iIndice = 1 To 4
        
        iIndiceAtual = iIndiceAtual + 1
        
        aiNumeros(iIndiceAtual) = StrParaInt(Mid(sAgMantConta, iIndice, 1))
                
    Next
    
    For iIndice = 1 To Len(sModalidadeBcoCobranca)
        
        iIndiceAtual = iIndiceAtual + 1
        
        aiNumeros(iIndiceAtual) = StrParaInt(Mid(sModalidadeBcoCobranca, iIndice, 1))
                
    Next
        
    For iIndice = 1 To Len(sNossoNumero)
        
        iIndiceAtual = iIndiceAtual + 1
        
        aiNumeros(iIndiceAtual) = StrParaInt(Mid(sNossoNumero, iIndice, 1))
                
    Next

    For iIndice = 1 To UBound(aiNumeros)
        If iIndice Mod 2 Then
            iFator = 2
        Else
            iFator = 1
        End If
        
        aiNumeros(iIndice) = aiNumeros(iIndice) * iFator
        
        iAcumulador = iAcumulador + IIf(aiNumeros(iIndice) > 10, (1 + aiNumeros(iIndice) Mod 10), aiNumeros(iIndice))
        
    Next
    
    iAcumulador = (10 - (iAcumulador Mod 10)) Mod 10
    
    sDVNossoNumero = FormataCpoNum(iAcumulador, 1)
    
    Exit Sub
    
End Sub

Function Calcula_NossoNumero_BancoIndustrial(objCarteiraCobrador As ClassCarteiraCobrador, sAgMantConta As String, sModalidadeBcoCobranca As String) As String

Dim sDVNossoNumero As String

    If objCarteiraCobrador.iComRegistro = MARCADO Then
    
        If objCarteiraCobrador.iNumCarteiraNoBanco = 112 Then
        
            If Len(Trim(objCarteiraCobrador.sFaixaNossoNumeroProx)) = 0 Then
                Calcula_NossoNumero_BancoIndustrial = FormataCpoNum("", 11)
            Else
                If Val(objCarteiraCobrador.sFaixaNossoNumeroProx) + 1 > objCarteiraCobrador.sFaixaNossoNumeroFinal Then gError 11111
                               
                objCarteiraCobrador.sFaixaNossoNumeroProx = FormataCpoNum(Val(objCarteiraCobrador.sFaixaNossoNumeroProx + 1), 10)
                
                Call Calcula_DV_NossoNumero_BancoIndustrial(sAgMantConta, sModalidadeBcoCobranca, FormataCpoNum(objCarteiraCobrador.sFaixaNossoNumeroProx, 10), sDVNossoNumero)
                
                Calcula_NossoNumero_BancoIndustrial = FormataCpoNum(objCarteiraCobrador.sFaixaNossoNumeroProx & sDVNossoNumero, 11)
                
            End If
            
        Else
            Calcula_NossoNumero_BancoIndustrial = FormataCpoNum("", 11)
        
        End If
    Else
    
        objCarteiraCobrador.sFaixaNossoNumeroProx = FormataCpoNum(Val(objCarteiraCobrador.sFaixaNossoNumeroProx) + 1, 10)
    
        Call Calcula_DV_NossoNumero_BancoIndustrial(sAgMantConta, sModalidadeBcoCobranca, FormataCpoNum(objCarteiraCobrador.sFaixaNossoNumeroProx, 10), sDVNossoNumero)
        
        Calcula_NossoNumero_BancoIndustrial = FormataCpoNum(objCarteiraCobrador.sFaixaNossoNumeroProx & sDVNossoNumero, 11)
    
    End If

    Exit Function

End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCNABBcoBrasil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Dim glNumArquivoRemessa As Long
Dim gColFiliaisEmpresa As Collection
Dim gobjCobrancaEletronica As ClassCobrancaEletronica
Public gobjGeracaoArqCNABPag As ClassGeracaoArqCNABPag
Dim gobjFilialEmpresa As AdmFiliais

Dim giTotCarteiraSimples As Integer
Dim giTotCarteiraVinculada  As Integer
Dim giTotCarteiraCaucionada  As Integer
Dim giTotCarteiraDescontada  As Integer
Dim gdValorTotCarteiraSimples As Double
Dim gdValorTotCarteiraVinculada  As Double
Dim gdValorTotCarteiraCaucionada  As Double
Dim gdValorTotCarteiraDescontada  As Double

Dim giNumLotes As Integer
Dim glNumRegistros As Long

Public Type typeRegControle
    sNumeroBanco As String
    sLoteServico As String
    sIdentificacaoReg As String
End Type

'====================== BANCO DO BRASIL ===============================
'240 bytes
'------------ Remessa de Títulos a Receber ----------------------------

'Registro de Header do arquivo de remessa titulos a receber
Public Type typeRemHeaderBcoBrasil240
    sCodBancoComp As String
    sLoteServico As String
    sRegHeaderLote As String
    sTipoOperacao As String
    sTipoServico As String
    sFormaLancamento As String
    sNVersaoLayLote As String
    sBrancosCNAB1 As String
    sTipoInscEmpresa As String
    sCodInsEmpresa As String
    sCodConvBanco As String
    sAgMantConta As String
    sDVAgencia As String
    sNumContaCorrente As String
    sDVConta As String
    sDVAgenciaConta As String
    sNomeEmpresa As String
    sMensagem1 As String
    sMensagem2 As String
    sNumRemessa As String
    sDataGravRem As String
    sDataCredito As String
    sBrancosCNAB2 As String
End Type

'Registro de Detalhe do arquivo de remessa titulos a receber
Public Type typeRemDetalheBcoBrasil240SegP
    sCodBancoComp As String
    sLoteServico As String
    sRegDetalhe As String
    sNumSequencialReg As String
    sCodSegRegDetalhe As String
    sBrancosCNAB1 As String
    sCodMovimento As String
    sAgMantConta As String
    sDVAgMantConta As String
    sNumContaCorrente As String
    sDVConta As String
    sDVAgenciaConta As String
    sIdentTituloBanco As String
    sCodCarteira As String
    sFormaCadTituloBanco As String
    sTipoDocumento As String
    sIdentEmissaoBloq As String
    sIdentDistrib As String
    sNumDocCobranca As String
    sDataVencTitulo As String
    sValorNominalTitulo As String
    sAgEncCobranca As String
    sDVAgencia As String
    sEspecieTitulo As String
    sIdentTituloAceite As String
    sDataEmissaoTitulo As String
    sCodJurosMora As String
    sDataJurosMora As String
    sJurosMoraDiaTaxa As String
    sCodDesconto1 As String
    sDataDesconto1 As String
    sValorPerConcedido As String
    sValorIOF As String
    sValorAbatimento As String
    sIdentTituloEmp As String
    sCodProtesto As String
    sNumDiasProtesto As String
    sCodBaixaDevolucao As String
    sNumDiasBaixaDevol As String
    sCodMoeda As String
    sNumContratoOpCred As String
    sBrancosCNAB2 As String

End Type

'Registro de Detalhe do arquivo de remessa titulos a receber
Public Type typeRemDetalheBcoBrasil240SegQ
    sCodBancoComp As String
    sLoteServico As String
    sRegDetalhe As String
    sNumSequencialReg As String
    sCodSegRegDetalhe As String
    sBrancosCNAB1 As String
    sCodMovimento As String
    sTipoInscricao As String
    sNumInscricao As String
    sNome As String
    sEndereco As String
    sBairro As String
    sCEP  As String
    sSufixoCEP As String
    sCidade As String
    sUnidFederacao As String
    sTipoInscricaoSacAval As String
    sNumInscricaoSacAval As String
    sNomeSacadorAvalista As String
    sBancoCompensacao As String
    sNossoNumBancoCorresp As String
    sBrancosCNAB2 As String
  
End Type

'Registro de Detalhe do arquivo de remessa titulos a receber
Public Type typeRemDetalheBcoBrasil240SegR
    sCodBancoComp As String
    sLoteServico As String
    sRegDetalhe As String
    sNumSequencialReg As String
    sCodSegRegDetalhe As String
    sBrancosCNAB1 As String
    sCodMovimento As String
    sCodDesconto2 As String
    sDataDesconto2 As String
    sValorPerConcedido2 As String
    sCodDesconto3 As String
    sDataDesconto3 As String
    sValorPerConcedido3 As String
    sCodMulta As String
    sDataMulta As String
    sValorPerAplicado As String
    sInfoBancoSacado As String
    sMensagem3 As String
    sMensagem4 As String
    sCodBancoContaDeb As String
    sCodAgContaDeb As String
    sCodContaDVDeb As String
    sCodOcorrenciaSac As String
    sBrancosCNAB2 As String
    
End Type

'Registro de Trailer do arquivo de remessa titulos a receber
Type typeRemTrailerBcoBrasil240
    
    sCodBancoComp As String
    sLoteServico As String
    sRegTrailer As String
    sBrancosCNAB1 As String
    sQuantRegLote As String
    sQuantTitCobS As String
    sValTotalTitCartS As String
    sQuantTitCobV As String
    sValTotalTitCartV As String
    sQuantTitCobC As String
    sValTotalTitCartC As String
    sQuantTitCobD As String
    sValTotalTitCartD As String
    sNumAvisoLancamento As String
    sBrancosCNAB2 As String

End Type

'------------------ Retorno de Títulos a Receber -------------------------

'Registro de Detalhe do arquivo de retorno titulos a receber
Public Type typeRetDetalheBcoBrasil240SegT
    
    sCodBancoComp As String
    sLoteServico As String
    sRegDetalhe As String
    sNumSequencialReg As String
    sCodSegRegDetalhe As String
    sBrancosCNAB1 As String
    sCodMovimento As String
    sAgMantConta As String
    sDVAgencia As String
    sNumContaCorrente As String
    sDVConta As String
    sDVAgenciaConta As String
    sIdentTituloBanco As String
    sCodCarteira As String
    sNumDocCobranca As String
    sDataVencTitulo As String
    sValorNominalTitulo As String
    sNumBanco As String
    sAgCobradoraRecebedora As String
    sDVAgencia1 As String
    sIdentTituloEmp As String
    sCodMoeda As String
    sTipoInscricao As String
    sNumInscricao As String
    sNome As String
    sNumContrato As String
    sValorTarifa As String
    sMotivoOcorrencia As String
    sBrancosCNAB2 As String
    
End Type

'Registro de Detalhe do arquivo de retorno titulos a receber
Public Type typeRetDetalheBcoBrasil240SegU
    
    sCodBancoComp As String
    sLoteServico As String
    sRegDetalhe As String
    sNumSequencialReg As String
    sCodSegRegDetalhe As String
    sBrancosCNAB1 As String
    sCodMovimento As String
    sJurosMultaEncargos As String
    sValorDescConcedido As String
    sValorAbatConcedCancel As String
    sValorIOFRecolhido As String
    sValorPagoSacado As String
    sValorLiqSerCreditado As String
    sValorOutrasDespesas As String
    sValorOutrosCreditos As String
    sDataOcorrencia As String
    sDataEfetCredito As String
    sCodOcorSacado As String
    sDataOcorSacado As String
    sValorOcorSacado As String
    sCompOcorSacado As String
    sCodBancoCorrespCompens As String
    sNossoNumBancoCorresp As String
    sBrancosCNAB2 As String

End Type

'Registro de trailer do arquivo de retorno titulos a receber
Public Type typeRetTrailerBcoBrasil
    sIdentificacaoReg As String
    sCodigoRetorno As String
    sCodigoServico As String
    sNumeroBanco As String
    sBrancos1 As String
    sQuantTitulosCobrSimp As String
    sValorTotalSimp As String
    sAvisoBancario As String
    sBrancos2 As String
    sQuantTitulosCobrVinc As String
    sValorTotalVinc As String
    sAvisoBancario2 As String
    sBrancos3 As String
    sQuantTitulosCobrCauc As String
    sValorTotalCauc As String
    sAvisoBancario3 As String
    sBrancos4 As String
    sQuantTitulosCobrDesc As String
    sValorTotalDesc As String
    sAvisoBancario4 As String
    sBrancos5 As String
    sSequencialRegistro As String
    
End Type

Function Criar_Arquivo_Remessa_BcoBrasil(objCobrancaEletronica As ClassCobrancaEletronica) As Long
'Cria uma arquivo de remessa com base nos dados do BD

Dim lErro As Long
Dim lTransacao As Long
Dim alComando(0 To 7) As Long
Dim sNomeArquivo As String
Dim iIndice As Integer, lSeqNomeArq As Long

Dim sRegHeader As String
Dim sRegDetalheSegP As String
Dim sRegDetalheSegQ As String
Dim sRegDetalheSegR As String
Dim sRegTrailer As String

Dim objBorderoCobranca As New ClassBorderoCobranca
Dim objCarteiraCobrador As New ClassCarteiraCobrador
Dim objOcorrRemParc As ClassOcorrRemParcRec

On Error GoTo Erro_Criar_Arquivo_Remessa_BcoBrasil
    
    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 51804
    Next
        
    'ABre a transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 51805
    
    Set gobjCobrancaEletronica = objCobrancaEletronica
    Set gobjFilialEmpresa = New AdmFiliais
    
    'Gera o nome do arquivo de acordo com o sequencial do Cobrador
    lErro = cf("CobrancaEletronica_Gera_NomeArquivo", objCobrancaEletronica.objCobrador, objCobrancaEletronica.sDiretorio, sNomeArquivo, lSeqNomeArq)
    If lErro <> SUCESSO Then gError 51806
    
    gobjCobrancaEletronica.lSeqNomeArq = lSeqNomeArq
    
    'Lê a conta corrente associada ao cobrador
    lErro = cf("ContaCorrenteInt_Le", gobjCobrancaEletronica.objCobrador.iCodCCI, gobjCobrancaEletronica.objCCI)
    If lErro <> SUCESSO And lErro <> 11807 Then gError 51814
    If lErro <> SUCESSO Then gError 51815

    If gobjCobrancaEletronica.objCCI.iCodBanco <> gobjCobrancaEletronica.objCobrador.iCodBanco Then gError 62063
    If Len(Trim(gobjCobrancaEletronica.objCCI.sAgencia)) = 0 Or Len(Trim(gobjCobrancaEletronica.objCCI.sNumConta)) = 0 Then gError 62064
    
    gobjFilialEmpresa.iCodFilial = giFilialEmpresa
    
    lErro = cf("FilialEmpresa_Le", gobjFilialEmpresa, True)
    If lErro <> SUCESSO And lErro <> 27378 Then gError 51814
    If lErro <> SUCESSO Then gError 90073
    
    'Cria o arquivo de remessa
    Open sNomeArquivo For Output As #1
           
    lErro = Montar_RegistroHeaderArquivo_BcoBrasil(sRegHeader)
    If lErro <> SUCESSO Then gError 51810
        
    'Inclui o registro de Header de arquivo
    Print #1, DesacentuaTexto(sRegHeader)
    
    glNumRegistros = glNumRegistros + 1
    
    iIndice = 0
    For Each objBorderoCobranca In objCobrancaEletronica.colBorderos

        'Lê o bordero de cobrança
        lErro = cf("BorderoCobranca_LeMod", alComando(2), objBorderoCobranca)
        If lErro <> SUCESSO Then gError 51807

        objCarteiraCobrador.iCobrador = objCobrancaEletronica.iCobrador
        objCarteiraCobrador.iCodCarteiraCobranca = objBorderoCobranca.iCodCarteiraCobranca

        'Lê a carteira do cobrador
        lErro = cf("CarteiraCobrador_LeMod", objCarteiraCobrador, alComando(4))
        If lErro <> SUCESSO Then gError 51809

        'Busca no BD as Ocorrências do Bordero em questão
        lErro = cf("CobrancaEletronica_Obter_OcorrRemParcRec", alComando(3), objBorderoCobranca, gobjCobrancaEletronica)
        If lErro <> SUCESSO Then gError 51662
        
'        'Alterado por Cyntia
'        If iIndice <> 0 Then
'            If objCobrancaEletronica.colBorderos(iIndice).iCodCarteiraCobranca <> objBorderoCobranca.iCodCarteiraCobranca Then giNumLotes = giNumLotes + 1
'        Else
            giNumLotes = giNumLotes + 1
'        End If
        
        'incluir header de lote
        lErro = Montar_RegistroHeaderLote_BcoBrasil(sRegHeader, objCarteiraCobrador)
        If lErro <> SUCESSO Then gError 81697
        
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
            lErro = Montar_RegistroDetalhe_BcoBrasil(sRegDetalheSegP, sRegDetalheSegQ, sRegDetalheSegR, objOcorrRemParc, objCarteiraCobrador)
            If lErro <> SUCESSO And lErro <> 51674 Then gError 51811

            If lErro = SUCESSO Then
            
                'Inclui o registro com os detalhes do Título(Parcela) a receber
                Print #1, DesacentuaTexto(sRegDetalheSegP)
                Print #1, DesacentuaTexto(sRegDetalheSegQ)
                Print #1, DesacentuaTexto(sRegDetalheSegR)
    
                glNumRegistros = glNumRegistros + 3
                
            End If
            
        Next

        'incluir trailer de lote
    
        'MOnta o registro de trailer de lote
        lErro = Montar_RegistroTrailerLote_BcoBrasil(sRegTrailer)
        If lErro <> SUCESSO Then gError 81698
    
        'grava o registro trailer de lote
        Print #1, DesacentuaTexto(sRegTrailer)
        
        glNumRegistros = glNumRegistros + 1
        
        objBorderoCobranca.iProcessado = BORDERO_PROCESSADO
        objBorderoCobranca.sNomeArquivo = sNomeArquivo

        'Atualiza o Status do Bordero de Cobrança
        lErro = cf("BorderoCobranca_Atualiza_Processo", alComando(2), alComando(3), objBorderoCobranca)
        If lErro <> SUCESSO Then gError 51808

        'Atualiza o NossoNumero em carteiraCobrador
        lErro = cf("CarteiraCobrador_AtualizaNossoNumero", objCarteiraCobrador, alComando(4), alComando(5))
        If lErro <> SUCESSO Then gError 51842
        
        iIndice = iIndice + 1
    Next

    'MOnta o registro de trailer de arquivo
    lErro = Montar_RegistroTrailerArq_BcoBrasil(sRegTrailer)
    If lErro <> SUCESSO Then gError 51812

    'grava o registro de trailer no arquivo
    Print #1, DesacentuaTexto(sRegTrailer)
    
    Close #1
    
    'atualizar seq de arquivo de remessa no cobrador
    lErro = Cobrador_Atualizar(objCobrancaEletronica.objCobrador, alComando(6), alComando(7))
    If lErro <> SUCESSO Then gError 81699
    
    lErro = Transacao_Commit()
    If lErro <> SUCESSO Then gError 62011
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Criar_Arquivo_Remessa_BcoBrasil = SUCESSO
    
    Exit Function

Erro_Criar_Arquivo_Remessa_BcoBrasil:

    Criar_Arquivo_Remessa_BcoBrasil = gErr
    
    Select Case gErr
    
        Case 51804
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 51805
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 51806, 51807, 51808, 51809, 51810, 51811, 51812, 51814, 51842, 81697, 81698
            MsgBox (CStr(gErr))
        
        Case 62011
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case 90073
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, giFilialEmpresa)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 147127)
    
    End Select

    Close #1

    Call Transacao_Rollback

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Private Function Montar_RegistroHeaderLote_BcoBrasil(sRegHeader As String, objCarteiraCobrador As ClassCarteiraCobrador) As Long
'Monta o registro

Dim tRegHeaderBcoBrasil240 As typeRemHeaderBcoBrasil240
Dim lErro As Long

On Error GoTo Erro_Montar_RegistroHeaderLote_BcoBrasil

    'Carrega tRegHeaderBcoBrasil com os dados do Header do aquivo
    lErro = Obter_Dados_Header_Lote(tRegHeaderBcoBrasil240, objCarteiraCobrador)
    If lErro <> SUCESSO Then Error 51813
        
    With tRegHeaderBcoBrasil240
    
        'Monta o registro de Header do BcoBrasil
        sRegHeader = .sCodBancoComp & .sLoteServico & .sRegHeaderLote & .sTipoOperacao & .sTipoServico & .sFormaLancamento & .sNVersaoLayLote & .sBrancosCNAB1
        sRegHeader = sRegHeader & .sTipoInscEmpresa & .sCodInsEmpresa & .sCodConvBanco & .sAgMantConta & .sDVAgencia & .sNumContaCorrente & .sDVConta & .sDVAgenciaConta
        sRegHeader = sRegHeader & .sNomeEmpresa & .sMensagem1 & .sMensagem2 & .sNumRemessa & .sDataGravRem & .sDataCredito & .sBrancosCNAB2
        
    End With
    
    Montar_RegistroHeaderLote_BcoBrasil = SUCESSO

    Exit Function

Erro_Montar_RegistroHeaderLote_BcoBrasil:

    Montar_RegistroHeaderLote_BcoBrasil = Err

    Select Case Err

        Case 51813

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147128)

    End Select

    Exit Function

End Function

Private Function Obter_Dados_Header_Lote(tRegHeaderBcoBrasil240 As typeRemHeaderBcoBrasil240, objCarteiraCobrador As ClassCarteiraCobrador) As Long

Dim lErro As Long

On Error GoTo Erro_Obter_Dados_Header_Lote
      
    'Carrega os campos do cabeçalho
    With tRegHeaderBcoBrasil240
   
        .sCodBancoComp = FormataCpoNum("001", 3)
        .sLoteServico = FormataCpoNum(giNumLotes, 4)
        .sRegHeaderLote = "1"
        .sTipoOperacao = "R"
        .sTipoServico = "01"
        .sFormaLancamento = FormataCpoNum("0", 2)
        .sNVersaoLayLote = "020"
        .sBrancosCNAB1 = FormataCpoTexto("", 1)
        
        If Len(gobjFilialEmpresa.sCgc) = 14 Then
            .sTipoInscEmpresa = "2"
        Else
            .sTipoInscEmpresa = "1"
        End If
        
        .sCodInsEmpresa = FormataCpoNum(gobjFilialEmpresa.sCgc, 15)
        
        .sCodConvBanco = FormataCpoTexto(Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 7), 20)
                
        Call cf("CodConvBanco_Obtem_Customizado", objCarteiraCobrador, .sCodConvBanco)
        
        .sAgMantConta = FormataCpoNum(gobjCobrancaEletronica.objCCI.sAgencia, 5)
        .sDVAgencia = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVAgencia, 1)
        .sNumContaCorrente = FormataCpoNum(gobjCobrancaEletronica.objCCI.sNumConta, 12)
        .sDVConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVNumConta, 1)
        .sDVAgenciaConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVAgConta, 1)
        
        .sNomeEmpresa = FormataCpoTexto(gsNomeEmpresa, 30)
        
        Call cf("Empresa_Le_Nome_Customizado", .sNomeEmpresa)
        
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

    Obter_Dados_Header_Lote = Err
    
    Select Case Err
    
        Case 62063
            lErro = Rotina_Erro(vbOKOnly, "ERRO_BANCO_CCI_DIFERENTE_COBRADOR", Err, gobjCobrancaEletronica.objCCI.iCodBanco, gobjCobrancaEletronica.objCobrador.iCodBanco)
    
        Case 62064
            lErro = Rotina_Erro(vbOKOnly, "ERRO_AGENCIA_CONTA_COBRADOR_NAO_PREENCHIDAS", Err)
    
        Case 51815
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_COBRADOR_NAO_ENCONTRADA", Err, gobjCobrancaEletronica.iCobrador)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147129)
            
    End Select
    
    Exit Function

End Function

Private Function Montar_RegistroDetalhe_BcoBrasil(sRegDetalheSegP As String, sRegDetalheSegQ As String, sRegDetalheSegR As String, objOcorrRemParcRec As ClassOcorrRemParcRec, objCarteiraCobrador As ClassCarteiraCobrador) As Long
'Monta o registro

Dim tRegDetalheBcoBrasil240SegP As typeRemDetalheBcoBrasil240SegP
Dim tRegDetalheBcoBrasil240SegQ As typeRemDetalheBcoBrasil240SegQ
Dim tRegDetalheBcoBrasil240SegR As typeRemDetalheBcoBrasil240SegR
Dim lErro As Long
Dim objParcelaRec As New ClassParcelaReceber
 
On Error GoTo Erro_Montar_RegistroDetalhe_BcoBrasil
    
    objParcelaRec.lNumIntDoc = objOcorrRemParcRec.lNumIntParc

    'Busca os dados da parcela a receber no BD
    lErro = cf("CobrancaEletronica_Obter_DadosParcelaRec", objParcelaRec, gobjCobrancaEletronica)
    If lErro <> SUCESSO And lErro <> 51674 Then Error 51817
    If lErro <> SUCESSO Then Error 51674
                    
    gobjCobrancaEletronica.lQuantTitulos = gobjCobrancaEletronica.lQuantTitulos + 1
      
    gobjCobrancaEletronica.lSequencialRegistro = gobjCobrancaEletronica.lSequencialRegistro + 1
    
    'Obtem os dados do registro de detalhe SegmentoP
    lErro = Obter_Dados_DetalheSegP(tRegDetalheBcoBrasil240SegP, objOcorrRemParcRec, objCarteiraCobrador, objParcelaRec)
    If lErro <> SUCESSO Then Error 51816
    
    With tRegDetalheBcoBrasil240SegP
        'Forma o registro de detalhe
        sRegDetalheSegP = .sCodBancoComp & .sLoteServico & .sRegDetalhe & .sNumSequencialReg & .sCodSegRegDetalhe & .sBrancosCNAB1 & .sCodMovimento
        sRegDetalheSegP = sRegDetalheSegP & .sAgMantConta & .sDVAgMantConta & .sNumContaCorrente & .sDVConta & .sDVAgenciaConta & .sIdentTituloBanco & .sCodCarteira & .sFormaCadTituloBanco
        sRegDetalheSegP = sRegDetalheSegP & .sTipoDocumento & .sIdentEmissaoBloq & .sIdentDistrib & .sNumDocCobranca & .sDataVencTitulo & .sValorNominalTitulo & .sAgEncCobranca & .sDVAgencia & .sEspecieTitulo
        sRegDetalheSegP = sRegDetalheSegP & .sIdentTituloAceite & .sDataEmissaoTitulo & .sCodJurosMora & .sDataJurosMora & .sJurosMoraDiaTaxa & .sCodDesconto1 & .sDataDesconto1
        sRegDetalheSegP = sRegDetalheSegP & .sValorPerConcedido & .sValorIOF & .sValorAbatimento & .sIdentTituloEmp & .sCodProtesto & .sNumDiasProtesto
        sRegDetalheSegP = sRegDetalheSegP & .sCodBaixaDevolucao & .sNumDiasBaixaDevol & .sCodMoeda & .sNumContratoOpCred & .sBrancosCNAB2
        
    End With
    
    gobjCobrancaEletronica.lSequencialRegistro = gobjCobrancaEletronica.lSequencialRegistro + 1
    
    'Obtem os dados do registro de detalhe SegmentoQ
    lErro = Obter_Dados_DetalheSegQ(tRegDetalheBcoBrasil240SegQ, objOcorrRemParcRec, objCarteiraCobrador, objParcelaRec)
    If lErro <> SUCESSO Then Error 51816
    
    With tRegDetalheBcoBrasil240SegQ
        'Forma o registro de detalhe
        sRegDetalheSegQ = .sCodBancoComp & .sLoteServico & .sRegDetalhe & .sNumSequencialReg & .sCodSegRegDetalhe & .sBrancosCNAB1 & .sCodMovimento
        sRegDetalheSegQ = sRegDetalheSegQ & .sTipoInscricao & .sNumInscricao & .sNome & .sEndereco & .sBairro & .sCEP & .sSufixoCEP & .sCidade & .sUnidFederacao
        sRegDetalheSegQ = sRegDetalheSegQ & .sTipoInscricaoSacAval & .sNumInscricaoSacAval & .sNomeSacadorAvalista & .sBancoCompensacao & .sNossoNumBancoCorresp & .sBrancosCNAB2
        
    End With
    
    gobjCobrancaEletronica.lSequencialRegistro = gobjCobrancaEletronica.lSequencialRegistro + 1
        
    'Obtem os dados do registro de detalhe SegmentoR
    lErro = Obter_Dados_DetalheSegR(tRegDetalheBcoBrasil240SegR, objOcorrRemParcRec, objCarteiraCobrador, objParcelaRec)
    If lErro <> SUCESSO Then Error 51816
    
    With tRegDetalheBcoBrasil240SegR
        'Forma o registro de detalhe
        sRegDetalheSegR = .sCodBancoComp & .sLoteServico & .sRegDetalhe & .sNumSequencialReg & .sCodSegRegDetalhe & .sBrancosCNAB1 & .sCodMovimento
        sRegDetalheSegR = sRegDetalheSegR & .sCodDesconto2 & .sDataDesconto2 & .sValorPerConcedido2 & .sCodDesconto3 & .sDataDesconto3 & .sValorPerConcedido3
        sRegDetalheSegR = sRegDetalheSegR & .sCodMulta & .sDataMulta & .sValorPerAplicado & .sInfoBancoSacado & .sMensagem3 & .sMensagem4
        sRegDetalheSegR = sRegDetalheSegR & .sCodBancoContaDeb & .sCodAgContaDeb & .sCodContaDVDeb & .sCodOcorrenciaSac & .sBrancosCNAB2

    End With
    
    Montar_RegistroDetalhe_BcoBrasil = SUCESSO

    Exit Function

Erro_Montar_RegistroDetalhe_BcoBrasil:

     Montar_RegistroDetalhe_BcoBrasil = Err

    Select Case Err
    
        Case 51816, 51674
            MsgBox (CStr(Err))

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147130)

    End Select
    
    Exit Function

End Function

'Private Function Obter_Dados_Detalhe(tRegDetalheBcoBrasil As typeRemDetalheBcoBrasil, objOcorrRemParcRem As ClassOcorrRemParcRec, objCarteiraCobrador As ClassCarteiraCobrador) As Long
''Obtem os dados que são utilizados para preencher o registro de detalhe
'
'Dim lErro As Long
'Dim objParcelaRec As New ClassParcelaReceber
'Dim objEndereco As New ClassEndereco
'Dim sNossoNumero As String
'Dim dValorDesconto As Double
'Dim sDVNossoNumero As String
'
'On Error GoTo Erro_Obter_Dados_Detalhe
'
'    objParcelaRec.lNumIntDoc = objOcorrRemParcRem.lNumIntParc
'
'    'Busca os dados da parcela a receber no BD
'    lErro = CF("CobrancaEletronica_Obter_DadosParcelaRec",objParcelaRec, gobjCobrancaEletronica)
'    If lErro <> SUCESSO Then Error 51817
'
'    'Se o endereço de cobrança estiver preenchido
'    If gobjCobrancaEletronica.lEnderecoCobrancaSacado > 0 Then
'        'Preenche o objEndereço com o endereço de cobrança
'        objEndereco.lCodigo = gobjCobrancaEletronica.lEnderecoCobrancaSacado
'    'Se o endereço principal estiver preenchido
'    ElseIf gobjCobrancaEletronica.lEnderecoSacado > 0 Then
'        'Preenche o objEndereço com o endereço principal
'        objEndereco.lCodigo = gobjCobrancaEletronica.lEnderecoSacado
'    'Se nenhum endereço estiver preenchido e o banco for o responsável pela impressão da boleta
'    ElseIf objCarteiraCobrador.iImprimeBoleta = BANCO_IMPRIME_BOLETA Then
'        'Erro
'        Error 51820
'    End If
'
'    If objEndereco.lCodigo > 0 Then
'        'Lê o endereço
'        lErro = CF("Endereco_Le",objEndereco)
'        If lErro <> SUCESSO And lErro <> 12309 Then Error 51818
'    End If
'
'    With tRegDetalheBcoBrasil
'
'        .sIdentificacaoReg = "1"
'        If Len(gobjCobrancaEletronica.sCGCCedente) = 14 Then
'            .sTipoInscricaoCedente = "02"
'        Else
'            .sTipoInscricaoCedente = "01"
'        End If
'        .sNumInscricaoCedente = FormataCpoNum(gobjCobrancaEletronica.sCGCCedente, 14)
'        .sAgencia = FormataCpoNum(gobjCobrancaEletronica.objCCI.sAgencia, 4)
'        .sDVAgencia = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVAgencia, 1)
'        .sCodigoCedente = FormataCpoNum(Obter_Info_Cobrador(gobjCobrancaEletronica.objCobradorCNABInfo.colInformacoes, 1), 8)
'        .sDVCodCedente = FormataCpoTexto(Obter_Info_Cobrador(gobjCobrancaEletronica.objCobradorCNABInfo.colInformacoes, 2), 1)
'        .sNumConvenio = FormataCpoNum(Obter_Info_Cobrador(gobjCobrancaEletronica.objCobradorCNABInfo.colInformacoes, 7), 6)
'        .sNumControleCli = FormataCpoTexto(objOcorrRemParcRem.lNumIntDoc, 25)
'        '
'        If objCarteiraCobrador.iNumCarteiraNoBanco = 13 Then
'            .sNumPrestacao = FormataCpoNum(objParcelaRec.iNumParcela, 2)
'        Else
'            .sNumPrestacao = FormataCpoNum("", 2)
'        End If
'        .sIndGrupoValor = FormataCpoNum("", 2)
'        .sBrancos1 = FormataCpoTexto("", 4)
'
'        Select Case objCarteiraCobrador.iNumCarteiraNoBanco
'
'            Case 11, 13, 15, 17
'                .sPrefixoTitulo = "AI "
'
'            Case 12
'                .sPrefixoTitulo = "AIU"
'
'            Case 31, 51
'                .sPrefixoTitulo = "SD "
'
'            Case Else
'                .sPrefixoTitulo = "   "
'
'        End Select
'
'        .sVariacao = FormataCpoNum(Obter_Info_Cobrador(gobjCobrancaEletronica.objCobradorCNABInfo.colInformacoes, 5), 3)
'
'        Select Case objCarteiraCobrador.iNumCarteiraNoBanco
'
'            Case 31
'                If ((objParcelaRec.dtDataVencimento < (gdtDataHoje + 180)) And (objParcelaRec.dtDataVencimento > (gdtDataHoje - 14))) Then
'                    .sContaCaucao = "1"
'                Else
'                    .sContaCaucao = "2"
'                End If
'
'            Case Else
'                .sContaCaucao = "0"
'
'        End Select
'
'        Select Case objCarteiraCobrador.iNumCarteiraNoBanco
'
'            Case 51, 54
'                .sCodRespons = FormataCpoNum(Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 3), 5)
'                .sDVCodRespons = FormataCpoNum(Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 4), 1)
'
'            Case Else
'                .sCodRespons = FormataCpoNum("", 5)
'                .sDVCodRespons = FormataCpoTexto("", 1)
'
'        End Select
'
'        .sNumBordero = FormataCpoNum(0, 6) '??? na inpal estao usando 900xxx ver se será necessario
'
'        .sBrancos2 = FormataCpoTexto("", 5)
'        .sNumCarteiraBanco = FormataCpoNum(objCarteiraCobrador.iNumCarteiraNoBanco, 2)
'        .sComando = FormataCpoNum(objOcorrRemParcRem.iCodOcorrencia, 2)
'        .sNumTitulo = FormataCpoTexto(gobjCobrancaEletronica.lNumTitulo, "10")
'        .sVencimento = Format(objParcelaRec.dtDataVencimento, "DDMMYY")
'        .sValorTitulo = FormataCpoValor(objParcelaRec.dSaldo, 13)
'        .sNumeroBanco = "001"
'        .sAgenciaCobradora = FormataCpoNum("", 4)
'        .sDVAgenciaCobradora = FormataCpoTexto("", 1)
'        .sEspecie = FormataCpoNum(gobjCobrancaEletronica.iEspecieTitulo, 2)
'        If objParcelaRec.iAceite = 1 Then
'            .sAceite = "A"
'        Else
'            .sAceite = "N"
'        End If
'        .sDataEmissao = Format(gobjCobrancaEletronica.dtDataEmissaoTitulo, "DDMMYY")
'        .sInstrucao1 = FormataCpoNum(objOcorrRemParcRem.iInstrucao1, 2)
'        .sInstrucao2 = FormataCpoNum(objOcorrRemParcRem.iInstrucao2, 2)
'        .sJurosDiarios = FormataCpoValor(Round(objParcelaRec.dSaldo * gobjCobrancaEletronica.dPercJurosDiario, 2), 13)
'        If objParcelaRec.dtDesconto1Ate <> DATA_NULA Then
'            .sDescontoAte = Format(objParcelaRec.dtDesconto1Ate, "DDMMYY")
'        Else
'            .sDescontoAte = FormataCpoNum("", 6)
'        End If
'
'        'Calcula o valor do desconto quando for do tipo percentual
'        dValorDesconto = 0
'        If objParcelaRec.iDesconto1Codigo = PERC_ANT_DIA Or objParcelaRec.iDesconto1Codigo = PERC_ANT_DIA_UTIL Or objParcelaRec.iDesconto1Codigo = Percentual Then
'            dValorDesconto = StrParaDbl(Format(objParcelaRec.dDesconto1Valor * objParcelaRec.dSaldo, "Fixed"))
'        End If
'
'        .sValorDesconto = FormataCpoValor(dValorDesconto, 13)
'        .sValorIOF = FormataCpoValor(0, 13)
'        .sQuantUnidVar = FormataCpoNum("", 12)
'        .sEspecieValor = FormataCpoTexto("", 1)
'        .sAbatimento = FormataCpoValor(0, 13)
'        If Len(gobjCobrancaEletronica.sCGCSacado) = 14 Then
'            .sTipoInscSacado = "02"
'        Else
'            .sTipoInscSacado = "01"
'        End If
'
'        .sNumInscricaoSacado = FormataCpoNum(gobjCobrancaEletronica.sCGCSacado, 14)
'        .sNomeSacado = FormataCpoTexto(gobjCobrancaEletronica.sNomeSacado, 37)
'        .sBrancos3 = FormataCpoTexto("", 3)
'        .sEndereco = FormataCpoTexto(Trim(objEndereco.sEndereco), 37)
'        .sBrancos4 = FormataCpoTexto("", 3)
'        .sCEP = Formata_Cep(objEndereco.sCEP, 8)
'        .sCidade = FormataCpoTexto(objEndereco.sCidade, 15)
'        .sEstado = FormataCpoTexto(objEndereco.sSiglaEstado, 2)
'        .sObservacoes = FormataCpoTexto("", 40)
'        .sBrancos5 = FormataCpoTexto("", 3)
'        .sNumSequencialReg = FormataCpoNum(gobjCobrancaEletronica.lSequencialRegistro, 6)
'
'        lErro = Calcula_NossoNumero_BcoBrasil(objCarteiraCobrador, sNossoNumero, sDVNossoNumero, .sVariacao, .sEspecie, .sNumConvenio)
'        If lErro <> SUCESSO Then Error 51819
'
'        .sNossoNumero = FormataCpoNum(sNossoNumero, 11)
'        .sDVNossoNumero = sDVNossoNumero
'
'    End With
'
'    'Incrementa o sequencial de registro
'    gobjCobrancaEletronica.lSequencialRegistro = gobjCobrancaEletronica.lSequencialRegistro + 1
'
'    Obter_Dados_Detalhe = SUCESSO
'
'    Exit Function
'
'Erro_Obter_Dados_Detalhe:
'
'    Obter_Dados_Detalhe = Err
'
'    Select Case Err
'
'        Case 51817, 51818, 51819
'
'        Case 51820
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ENDERECO_FILIALCLIENTE_NAO_INFORMADO", Err, gobjCobrancaEletronica.lCliente, gobjCobrancaEletronica.iFilialCli)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147131)
'
'    End Select
'
'    Exit Function
'
'End Function

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
        lErro = cf("FilialEmpresa_Le", objFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 27378 Then Error 51821
        If lErro <> SUCESSO Then Error 51822
    
        'Adiciona na coleção das filiais lidas
        gColFiliaisEmpresa.Add objFilialEmpresa
    End If
    
    Obter_FilialEmpresa = SUCESSO
    
    Exit Function

Erro_Obter_FilialEmpresa:

    Obter_FilialEmpresa = Err
    
    Select Case Err
    
        Case 51821
        
        Case 51822
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", Err, objFilialEmpresa.iCodFilial)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147132)
            
    End Select
    
    Exit Function
    
End Function

Private Function Montar_RegistroTrailerLote_BcoBrasil(sRegTrailer As String) As Long
'Monta o registro

Dim tRegTrailerBcoBrasil240 As typeRemTrailerBcoBrasil240
Dim lErro As Long

On Error GoTo Erro_Montar_RegistroTrailerLote_BcoBrasil
    
    With tRegTrailerBcoBrasil240
        .sCodBancoComp = "001"
        .sLoteServico = FormataCpoNum(giNumLotes, 4)
        .sRegTrailer = "5"
        .sBrancosCNAB1 = FormataCpoTexto("", 9)
        .sQuantRegLote = FormataCpoNum((gobjCobrancaEletronica.lSequencialRegistro + 2), 6)
        .sQuantTitCobS = FormataCpoNum(0, 6)
        .sValTotalTitCartS = FormataCpoNum(0, 17)
        .sQuantTitCobV = FormataCpoNum(giTotCarteiraVinculada, 6)
        .sValTotalTitCartV = FormataCpoNum(gdValorTotCarteiraVinculada, 17)
        .sQuantTitCobC = FormataCpoNum(giTotCarteiraCaucionada, 6)
        .sValTotalTitCartC = FormataCpoNum(gdValorTotCarteiraCaucionada, 17)
        .sQuantTitCobD = FormataCpoNum(giTotCarteiraDescontada, 6)
        .sValTotalTitCartD = FormataCpoNum(gdValorTotCarteiraDescontada, 17)
        .sNumAvisoLancamento = FormataCpoTexto("", 8)
        .sBrancosCNAB2 = FormataCpoTexto("", 117)
    End With
    
    With tRegTrailerBcoBrasil240
        'Monta o Registro de trailer
        sRegTrailer = .sCodBancoComp & .sLoteServico & .sRegTrailer & .sBrancosCNAB1
        sRegTrailer = sRegTrailer & .sQuantRegLote & .sQuantTitCobS & .sValTotalTitCartS
        sRegTrailer = sRegTrailer & .sQuantTitCobV & .sValTotalTitCartV & .sQuantTitCobC
        sRegTrailer = sRegTrailer & .sValTotalTitCartC & .sQuantTitCobD & .sValTotalTitCartD & .sNumAvisoLancamento & .sBrancosCNAB2
    
    End With
    
    Montar_RegistroTrailerLote_BcoBrasil = SUCESSO

    Exit Function

Erro_Montar_RegistroTrailerLote_BcoBrasil:

    Montar_RegistroTrailerLote_BcoBrasil = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147133)

    End Select

    Exit Function

End Function

Function Processar_ArquivoRetorno_Cobranca_BcoBrasil(objCobrancaEletronica As ClassCobrancaEletronica) As Long
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

On Error GoTo Erro_Processar_ArquivoRetorno_Cobranca_BcoBrasil

    objBorderoRetCobr.iBanco = 1
    
    'le as carteiras de cobranca do cobrador
    lErro = cf("CarteirasCobrador_Le", objCobrancaEletronica.objCobrador, objBorderoRetCobr.colCartCobr)
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
    
    Call RetCobr_ConfirmaAtualizacao(objBorderoRetCobr, bDesistiu)
    If lErro <> SUCESSO Then Error 32260

    If objBorderoRetCobr.colRetCobrErros.Count <> 0 Then
        '???
        If MsgBox("Grava informações dos títulos não encontrados ?", vbYesNo, "Bordero de Retorno de Cobrança") = vbYes Then Call RetCobr_Grava_Erros(objBorderoRetCobr)
    End If

    If bDesistiu = False Then

        'atualiza o valor a ser debitado ou creditado por carteira
        For Each objBorderoRetCobrCart In objBorderoRetCobr.colBorderoRetCobrCart

            dValorMov = Round(objBorderoRetCobrCart.dValorRecebido - objBorderoRetCobrCart.dValorCustas - objBorderoRetCobrCart.dValorTarifas, 2)
            objBorderoRetCobrCart.objMovCCI.dValor = Abs(dValorMov)

            If dValorMov >= 0 Then

                objBorderoRetCobrCart.objMovCCI.iTipo = MOVCCI_CRED_RET_COBRANCA

            Else

                objBorderoRetCobrCart.objMovCCI.iTipo = MOVCCI_DEB_RET_COBRANCA

            End If

        Next

        'Grava os dados recolhidos durante o processamento do arquivo de retorno
        lErro = cf("BorderoRetCobr_Grava", objBorderoRetCobr)
        If lErro <> SUCESSO Then Error 51829

    End If
    
    Processar_ArquivoRetorno_Cobranca_BcoBrasil = SUCESSO
    
    Exit Function
    
Erro_Processar_ArquivoRetorno_Cobranca_BcoBrasil:

    Processar_ArquivoRetorno_Cobranca_BcoBrasil = Err
    
    Select Case Err
    
        Case 51823
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 51824, 51825
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 32260, 32265, 51826, 51827, 51828, 51829
        
        Case 51831
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORMATO_ARQUIVO_INCORRETO", Err, objCobrancaEletronica.sNomeArquivoRetorno)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147134)
    
    End Select

    Close #2
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Private Function Processa_Registro_Header_Arq(sRegistro As String, objBorderoRetCobr As ClassBorderoRetCobr, objCobrador As ClassCobrador, sNomeArquivo As String) As Long

Dim lErro As Long
Dim sConvenio As String
Dim objCarteiraCobrador As ClassCarteiraCobrador
Dim objCNABCarteiraInfo As ClassCNABInfo
Dim objCodigoTexto As AdmCodigoNome
Dim iAchou As Integer

On Error GoTo Erro_Processa_Registro_Header_Arq

    'Preenche o objBorderoRetCobr
    objBorderoRetCobr.dtDataRecepcao = gdtDataHoje
    objBorderoRetCobr.iCobrador = objCobrador.iCodigo
    objBorderoRetCobr.sNomeArq = sNomeArquivo
    
    sConvenio = Trim(Mid(sRegistro, 33, 13)) '??? no retorno só vem parte do convenio
    
    For Each objCarteiraCobrador In objBorderoRetCobr.colCartCobr
    
        Set objCNABCarteiraInfo = New ClassCNABInfo
        
        objCNABCarteiraInfo.iCodCobrador = objCarteiraCobrador.iCobrador
        objCNABCarteiraInfo.iCarteiraCobrador = objCarteiraCobrador.iCodCarteiraCobranca
    
        lErro = cf("CarteiraCobradorInfo_Le", objCNABCarteiraInfo.iCodCobrador, objCNABCarteiraInfo.iCarteiraCobrador, objCNABCarteiraInfo.colInformacoes)
        If lErro <> SUCESSO Then gError 126798
    
        iAchou = 0
    
        For Each objCodigoTexto In objCNABCarteiraInfo.colInformacoes
        
            'se for o codigo do convencio (7)
            If objCodigoTexto.iCodigo = 7 Then
                If Trim(Mid(objCodigoTexto.sNome, 1, 13)) = sConvenio Then
                    iAchou = 1
                    Exit For
                End If
            End If
        
        Next
    
        If iAchou = 1 Then Exit For
        
    Next
    
    'se nao achou o codigo do convenio associado as carteiras de cobranca do banco cobrador ==> erro
    If iAchou = 0 Then gError 126799
    
    Processa_Registro_Header_Arq = SUCESSO
    
    Exit Function

Erro_Processa_Registro_Header_Arq:

    Processa_Registro_Header_Arq = gErr

    Select Case gErr
    
        Case 126798
        
        Case 126799
            Call Rotina_Erro(vbOKOnly, "ERRO_CODCONVENIOCNAB", gErr, sConvenio)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147135)
            
    End Select
    
    Exit Function

End Function

Private Function Extrair_Header_Retorno(tRegHeaderBcoBrasil As typeRetHeaderBcoBrasil, sRegistro As String) As Long
'Extrai do sRegistro os dados do header do arquido de retorno

Dim lErro As Long

On Error GoTo Erro_Extrair_Header_Retorno

    With tRegHeaderBcoBrasil
        .sIdentificacaoReg = Mid(sRegistro, 1, 1)
        .sIdentificacaoArq = Mid(sRegistro, 2, 1)
        .sLiteralRetorno = Mid(sRegistro, 3, 7)
        .sCodTipoServico = Mid(sRegistro, 10, 2)
        .sLiteralServico = Mid(sRegistro, 12, 8)
        .sBrancos1 = Mid(sRegistro, 20, 7)
        .sAgencia = Mid(sRegistro, 27, 4)
        .sDVAgencia = Mid(sRegistro, 31, 1)
        .sCodigoCedente = Mid(sRegistro, 32, 8)
        .sDVCodCedente = Mid(sRegistro, 40, 1)
        .sNumCovenente = Mid(sRegistro, 41, 6)
        .sNomeCedente = Mid(sRegistro, 47, 30)
        .sNumeroBanco = Mid(sRegistro, 77, 3)
        .sNomeBanco = Mid(sRegistro, 80, 15)
        .sDataGravacao = Mid(sRegistro, 95, 6)
        .sSeqRetorno = Mid(sRegistro, 101, 7)
        .sBrancos2 = Mid(sRegistro, 108, 287)
        .sSequencialReg = Mid(sRegistro, 395, 6)
    
    End With

    Extrair_Header_Retorno = SUCESSO

    Exit Function

Erro_Extrair_Header_Retorno:

    Extrair_Header_Retorno = Err
    
    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147136)
            
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

Private Function Processa_Registro_Detalhe(objLoteServico As ClassCNABLoteServico, sRegistro As String, objBorderoRetCobr As ClassBorderoRetCobr, objCobrancaEletronica As ClassCobrancaEletronica, tRegDetalhe As typeRetCobrDetalheCNAB240, lComando As Long, lComando1 As Long) As Long

Dim objOcorrRemParcRec As New ClassOcorrRemParcRec
Dim iCodOcorrencia As Integer
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
    
        'dependendo do tipo de segmento
        Select Case Mid(sRegistro, 14, 1)
            
            Case "T"
                'Extrai da variável lida no arquivo os dados do registro de detalhe
                lErro = Extrair_Detalhe_Retorno(tRegDetalhe, sRegistro)
                If lErro <> SUCESSO Then Error 51835
            
            Case "U"
                'Extrai da variável lida no arquivo os dados do registro de detalhe
                lErro = Extrair_Detalhe_Retorno(tRegDetalhe, sRegistro)
                If lErro <> SUCESSO Then Error 51835
        
                'Recolhe o código da ocorrência
                iCodOcorrencia = StrParaInt(tRegDetalhe.tT030.sCodMovto)
                
                Select Case iCodOcorrencia
                    
                    'Quando ocorre baixa p/liquidacao
                    Case 6
                    
                        lErro = Prepara_CartCobrRet(objCobrancaEletronica, objBorderoRetCobr, objBorderoRetCobrCart, objDetRetCobr, tRegDetalhe)
                        If lErro <> SUCESSO Then Error 32356
        
                        '??? se o nosso sistema estiver gerando o numero p/uso da empresa
                        If False Then
                        
                            objOcorrRemParcRec.lNumIntDoc = StrParaLong(tRegDetalhe.tT030.sUsoEmpresa)
                            
                            'Lê os dados da remessa em ocorremparcrec
                            lErro = cf("OcorrRemParcRec_Le_NumIntOcorr", objOcorrRemParcRec)
                            If lErro <> SUCESSO And lErro <> 51719 Then Error 51694
                            If lErro <> SUCESSO Then Error 51696
                    
                            objParcelaRec.lNumIntDoc = objOcorrRemParcRec.lNumIntParc
                        
                            lErro = cf("RetCobr_Preenche_InfoParcRec", objParcelaRec, objInfoParcRec)
                            If lErro <> SUCESSO Then Error 51695
            
                            'Processa a parcela a receber
                            lErro = Processa_ParcelaRec(objParcelaRec, objBorderoRetCobr, objBorderoRetCobrCart, objDetRetCobr, objInfoParcRec)
                            If lErro <> SUCESSO Then Error 51695
                        
                        Else
                                   
                            lErro = RetCobr_Obtem_Parcela(objDetRetCobr, objParcelaRec, tRegDetalhe, lComando)
                            If lErro <> SUCESSO And lErro <> 32269 And lErro <> 32271 Then Error 32264
                            'se nao encontrou a parcela
                            If lErro <> SUCESSO Then
                        
                                'incluir ocorrencia de erro do retorno
                                Call objBorderoRetCobr.colRetCobrErros.Add(objDetRetCobr)
                        
                            Else
                        
                                lErro = cf("RetCobr_Preenche_InfoParcRec", objParcelaRec, objInfoParcRec)
                                If lErro <> SUCESSO Then Error 51695
            
                                'Processa a parcela a receber
                                lErro = Processa_ParcelaRec(objParcelaRec, objBorderoRetCobr, objBorderoRetCobrCart, objDetRetCobr, objInfoParcRec)
                                If lErro <> SUCESSO Then Error 51695
                                              
                            End If
                            
                        End If
                        
                    'cobranca de tarifas e custas
                    Case 2, 12, 20, 23, 28
                    
                        lErro = Prepara_CartCobrRet(objCobrancaEletronica, objBorderoRetCobr, objBorderoRetCobrCart, objDetRetCobr, tRegDetalhe)
                        If lErro <> SUCESSO Then Error 32257
        
                        objBorderoRetCobrCart.dValorTarifas = Round(objBorderoRetCobrCart.dValorTarifas + objDetRetCobr.dValorTarifa, 2)
                                            
                End Select
                
        End Select
    
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147137)
            
    End Select
    
    Exit Function

End Function

Private Function Extrair_Detalhe_Retorno(tRegDetalheBcoBrasil As typeRetCobrDetalheCNAB240, sRegistro As String) As Long
'Extrai do sRegistro os dados do Detalhe do arquido de retorno

Dim lErro As Long

On Error GoTo Erro_Extrair_Detalhe_Retorno

    Select Case Mid(sRegistro, 14, 1)
    
        Case "T"
            
            With tRegDetalheBcoBrasil.tT030
    
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
                .sNossoNumero = Mid(sRegistro, 38, 20)
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
        
            With tRegDetalheBcoBrasil.tU030
    
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
                .sValorCreditado = Mid(sRegistro, 98, 15)
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

    Extrair_Detalhe_Retorno = Err
    
    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147138)
            
    End Select
    
    Exit Function

End Function

Function Processa_Registro_Trailer(sRegistroTrailer As String) As Long

Dim lErro As Long
Dim tRegTrailerBcoBrasil As typeRetTrailerBcoBrasil

On Error GoTo Erro_Processa_Registro_Trailer

    'Extrai da variável lida no arquivo os dados do registro de detalhe
    lErro = Extrair_Trailer_Retorno(tRegTrailerBcoBrasil, sRegistroTrailer)
    If lErro <> SUCESSO Then Error 51841
    
    Processa_Registro_Trailer = SUCESSO
    
    Exit Function
    
Erro_Processa_Registro_Trailer:

    Processa_Registro_Trailer = Err
    
    Select Case Err
    
        Case 51841
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147139)
            
    End Select
    
    Exit Function
    
End Function

Private Function Extrair_Trailer_Retorno(tRegTrailerBcoBrasil As typeRetTrailerBcoBrasil, sRegistro As String) As Long
'Extrai do sRegistro os dados do Trailer do arquido de retorno

Dim lErro As Long

On Error GoTo Erro_Extrair_Trailer_Retorno

    tRegTrailerBcoBrasil.sIdentificacaoReg = Mid(sRegistro, 1, 1)
    tRegTrailerBcoBrasil.sCodigoRetorno = Mid(sRegistro, 2, 1)
    tRegTrailerBcoBrasil.sCodigoServico = Mid(sRegistro, 3, 2)
    tRegTrailerBcoBrasil.sNumeroBanco = Mid(sRegistro, 5, 3)
    tRegTrailerBcoBrasil.sBrancos1 = Mid(sRegistro, 8, 10)
    tRegTrailerBcoBrasil.sQuantTitulosCobrSimp = Mid(sRegistro, 18, 8)
    tRegTrailerBcoBrasil.sValorTotalSimp = Mid(sRegistro, 26, 14)
    tRegTrailerBcoBrasil.sAvisoBancario = Mid(sRegistro, 40, 8)
    tRegTrailerBcoBrasil.sBrancos2 = Mid(sRegistro, 48, 10)
    tRegTrailerBcoBrasil.sQuantTitulosCobrVinc = Mid(sRegistro, 58, 8)
    tRegTrailerBcoBrasil.sValorTotalVinc = Mid(sRegistro, 66, 14)
    tRegTrailerBcoBrasil.sAvisoBancario2 = Mid(sRegistro, 80, 8)
    tRegTrailerBcoBrasil.sBrancos3 = Mid(sRegistro, 88, 10)
    tRegTrailerBcoBrasil.sQuantTitulosCobrCauc = Mid(sRegistro, 98, 8)
    tRegTrailerBcoBrasil.sValorTotalCauc = Mid(sRegistro, 106, 14)
    tRegTrailerBcoBrasil.sAvisoBancario3 = Mid(sRegistro, 120, 8)
    tRegTrailerBcoBrasil.sBrancos4 = Mid(sRegistro, 128, 10)
    tRegTrailerBcoBrasil.sQuantTitulosCobrDesc = Mid(sRegistro, 138, 8)
    tRegTrailerBcoBrasil.sValorTotalDesc = Mid(sRegistro, 146, 14)
    tRegTrailerBcoBrasil.sAvisoBancario4 = Mid(sRegistro, 160, 8)
    tRegTrailerBcoBrasil.sBrancos5 = Mid(sRegistro, 168, 10)
    tRegTrailerBcoBrasil.sSequencialRegistro = Mid(sRegistro, 395, 6)

    Extrair_Trailer_Retorno = SUCESSO

    Exit Function

Erro_Extrair_Trailer_Retorno:

    Extrair_Trailer_Retorno = Err
    
    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147140)
            
    End Select
    
    Exit Function

End Function

Function BorderoPagto_Criar_ArquivoCNAB_BcoBrasil(objGeracaoArqCNABPag As ClassGeracaoArqCNABPag) As Long
'Cria o arquivo de remessa de títulos a pagar nos padrões do BcoBrasil

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

On Error GoTo Erro_BorderoPagto_Criar_ArquivoCNAB_BcoBrasil

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 51790
    Next

    'Abre a transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 51791
    
    'Inicializa os obj globais da tela
    Set gobjGeracaoArqCNABPag = objGeracaoArqCNABPag
    Set gColFiliaisEmpresa = New Collection
    
    'faz o lock no bordero de pagamento
    lErro = cf("BorderoPagto_Lock", alComando(0), gobjGeracaoArqCNABPag.objBorderoPagto.lNumIntBordero)
    If lErro <> SUCESSO Then Error 51792
    
    'Gera o nome do arquivo
    lErro = cf("PagamentoEletronico_Gera_NomeArquivo", sNomeArquivo, gobjGeracaoArqCNABPag)
    If lErro <> SUCESSO Then Error 51793
 
    'Cria o arquivo de remessa
    Open sNomeArquivo For Output As #1
 
    gobjGeracaoArqCNABPag.lQuantRegistrosArq = 1
    
    'Monta o registro de Header do arquivo
    lErro = Montar_RegistroHeaderCNABPag_BcoBrasil(sRegHeaderArq)
    If lErro <> SUCESSO Then Error 51794
    
    'Inclui no arquivo
    Print #1, DesacentuaTexto(sRegHeaderArq)
    
    iTipoCobranca = gobjGeracaoArqCNABPag.objBorderoPagto.iTipoDeCobranca
    gobjGeracaoArqCNABPag.lSequencialRegDetalhe = 1
    gobjGeracaoArqCNABPag.dValorTotalLote = 0
    
    'Lê as baixas ligadas ao bordero
    lErro = cf("BaixasParcPag_Le_BorderoPagto", alComando(2), gobjGeracaoArqCNABPag.objBorderoPagto.lNumIntBordero, gobjGeracaoArqCNABPag.colBaixasParcPag)
    If lErro <> SUCESSO Then Error 51795
    
    'Para cada baixa executada
    For Each objBaixaParcPag In gobjGeracaoArqCNABPag.colBaixasParcPag
        
        lErro = cf("BaixaParcPag_ObterDados_ParcelaTituloPag", alComando(2), objBaixaParcPag.lNumIntParcela, gobjGeracaoArqCNABPag)
        If lErro <> SUCESSO Then Error 51796
        
        If iTipoCobranca = TIPO_COBRANCA_OP Or iTipoCobranca = TIPO_COBRANCA_DOC Or iTipoCobranca = TIPO_COBRANCA_DEP_CONTA Then
        
            'Monta um registro de
            lErro = Montar_Registro_DetalhePagto(sRegDetalhe, objBaixaParcPag)
            If lErro <> SUCESSO Then Error 51797
                
            'incluir no arq
            Print #1, DesacentuaTexto(sRegDetalhe)
        End If
        
        Call objGeracaoArqCNABPag.objTelaAtualizacao.Mostra_Evolucao(True, 1)
        
        gobjGeracaoArqCNABPag.lSequencialRegDetalhe = gobjGeracaoArqCNABPag.lSequencialRegDetalhe + 1
        gobjGeracaoArqCNABPag.dValorTotalLote = gobjGeracaoArqCNABPag.dValorTotalLote + objBaixaParcPag.dValorBaixado - objBaixaParcPag.dValorDesconto + objBaixaParcPag.dValorJuros
        gobjGeracaoArqCNABPag.lQuantRegistrosArq = gobjGeracaoArqCNABPag.lQuantRegistrosArq + 1
    
    Next

    'Monta o registro de trailer de arquivo
    lErro = Montar_Registro_TrailerArq(sTrailerArq)
    If lErro <> SUCESSO Then Error 51798
    
    'Inclui no arquivo
    Print #1, DesacentuaTexto(sTrailerArq)
    
    'Fecha arquivo
    Close #1
    
    'Atualiza em BorderosPagto o Nome e o sequencial do arquivo
    lErro = Comando_ExecutarPos(alComando(1), "UPDATE BorderosPagto SET NomeArq = ? , NumArqRemessa = ?", alComando(0), sNomeArquivo, gobjGeracaoArqCNABPag.lSequencialArq + 1)
    If lErro <> AD_SQL_SUCESSO Then Error 51799
    
    'Faz o commit na transacao
    lErro = Transacao_Commit()
    If lErro <> SUCESSO Then Error 51800
        
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    BorderoPagto_Criar_ArquivoCNAB_BcoBrasil = SUCESSO
    
    Exit Function
    
Erro_BorderoPagto_Criar_ArquivoCNAB_BcoBrasil:

    BorderoPagto_Criar_ArquivoCNAB_BcoBrasil = Err
    
    Select Case Err
    
        Case 51790
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 51791
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
        
        Case 51792, 51794, 51725, 51795, 51796, 51797, 51730, 51731, 51798, 51793
        
        Case 51799
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_BORDERO_PAGTO", Err)
        
        Case 51800
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147141)
            
    End Select
    
    Close #1
    
    'Faz um RollBack na transacao
    Call Transacao_Rollback
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Function Montar_RegistroHeaderCNABPag_BcoBrasil(sRegHeader As String) As Long
'Monta o registro de header do arquivo de contas a pagar

Dim lErro As Long
Dim tRegHeaderArq As typePagtoHeaderArqCNAB240
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_Montar_RegistroHeaderCNABPag_BcoBrasil
  
    
    'incrementa o contador do número de registros do arquivo
    gobjGeracaoArqCNABPag.lQuantRegistrosArq = gobjGeracaoArqCNABPag.lQuantRegistrosArq + 1

    'Busca os dados da nossa filial empresa
    lErro = Obter_FilialEmpresa(giFilialEmpresa, objFilialEmpresa)
    If lErro <> SUCESSO Then Error 51801
    
    'Recolhe os dados do cabeçalho
    With tRegHeaderArq
        .sIdentificacaoReg = "0"
        .sCodigoRemessa = "1"
        .sBrancos1 = FormataCpoTexto("", 7)
        .sCodTipoServico = "03"
        .sindicadorCGC = FormataCpoTexto("", 1)
        .sValorTarifa = FormataCpoNum(0, 5)
        .sBrancos2 = FormataCpoTexto("", 9)
        .sCGCouPrefixoAg = Format(gobjGeracaoArqCNABPag.objCCI.sAgencia, 4)
        .sDVAgencia = FormataCpoTexto(gobjGeracaoArqCNABPag.objCCI.sDVAgencia, 1)
        .sContaCorrente = FormataCpoNum(gobjGeracaoArqCNABPag.objCCI.sNumConta, 9)
        .sDVContaCorrente = FormataCpoTexto(gobjGeracaoArqCNABPag.objCCI.sDVNumConta, 1)
        .sBrancos3 = FormataCpoTexto("", 5)
        
        .sNomeEmpresa = FormataCpoTexto(gsNomeEmpresa, 30)

        Call cf("Empresa_Le_Nome_Customizado", .sNomeEmpresa)

        .sNumeroBanco = "001"
        .sCodigoConvenente = "      "
''''        .sCodigoConvenente = FormataCpoNum(Obter_Info_Cobrador(gobjGeracaoArqCNABPag.objCobradorCNABInfo.colInformacoes, 6), 6)
        .sTipoRetorno = "RET"
        .sUsoDaEmpresa = FormataCpoTexto("", 10)
        .sMeioFisicoRet = "08"
        .sContadorRemessa = FormataCpoNum("", 3)
        .sBrancos4 = FormataCpoTexto("", 63)
        .sOpcaoLayOut = "NOVO"
        .sBrancos5 = FormataCpoTexto("", 15)
        .sTipoRetornoEnv = FormataCpoTexto("", 9)
        .sSequencialRegistro = FormataCpoNum(gobjGeracaoArqCNABPag.lQuantRegistrosArq, 6)
    
        'Forma o registro com os dados recolhidos p/o cabeçalho
        sRegHeader = .sIdentificacaoReg & .sCodigoRemessa & .sBrancos1 & .sCodTipoServico & .sindicadorCGC & .sValorTarifa & .sBrancos2 & .sCGCouPrefixoAg & .sDVAgencia
        sRegHeader = sRegHeader & .sContaCorrente & .sDVContaCorrente & .sBrancos3 & .sNomeEmpresa & .sNumeroBanco & .sCodigoConvenente & .sTipoRetorno & .sUsoDaEmpresa
        sRegHeader = sRegHeader & .sMeioFisicoRet & .sContadorRemessa & .sBrancos4 & .sOpcaoLayOut & .sBrancos5 & .sTipoRetornoEnv & .sSequencialRegistro
    End With
    
    gobjGeracaoArqCNABPag.lQuantRegistrosArq = gobjGeracaoArqCNABPag.lQuantRegistrosArq + 1
    
    Montar_RegistroHeaderCNABPag_BcoBrasil = SUCESSO
    
    Exit Function
    
Erro_Montar_RegistroHeaderCNABPag_BcoBrasil:

    Montar_RegistroHeaderCNABPag_BcoBrasil = Err
    
    Select Case Err
    
        Case 51801
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147142)
            
    End Select
    
    Exit Function
    
End Function

Function Montar_Registro_DetalhePagto(sRegDetalhe As String, objBaixaParcPag As ClassBaixaParcRec) As Long
'Monta o registro de transação do arquivo com os dados obtidos no BD

Dim lErro As Long
Dim tRegDetalhe As typePagtoDetalheCNAB240
Dim objEndereco As New ClassEndereco
Dim sAgencia As String
Dim sConta As String
Dim sDV As String

On Error GoTo Erro_Montar_Registro_DetalhePagto

    objEndereco.lCodigo = gobjGeracaoArqCNABPag.lEnderecoFavorecido
    'Lê o endereço do favorecido
    lErro = cf("Endereco_Le", objEndereco)
    If lErro <> SUCESSO And lErro <> 12309 Then Error 51802
    If lErro <> SUCESSO Then Error 51803 'Não encontrou

    'Recolhe os dados do registro de detalhe
    With tRegDetalhe
        .sIdentificacaoReg = "1"
        .sBrancos1 = FormataCpoTexto("", 1)
        .sIndicadorConf = "0"
        .sCGCouCPFConferir = FormataCpoNum("", 14)
        .sDVCGCouCPF = FormataCpoTexto("", 2)
        .sCGCouPrefixoEmpresa = FormataCpoNum(gobjGeracaoArqCNABPag.objCCI.sAgencia, 4)
        .sDVPrefixo = FormataCpoTexto(gobjGeracaoArqCNABPag.objCCI.sDVAgencia, 1)
        .sNumContaEmpresa = FormataCpoNum(gobjGeracaoArqCNABPag.objCCI.sNumConta, 9)
        .sDVNumConta = FormataCpoTexto(gobjGeracaoArqCNABPag.objCCI.sDVNumConta, 1)
        .sUsoDaEmpresa = FormataCpoTexto(objBaixaParcPag.lNumIntParcela & SEPARADOR & objBaixaParcPag.iSequencial, 10)
        .sBrancos2 = FormataCpoTexto("", 14)
        .sCodCamara = "000"
        .sCodBancoDest = FormataCpoNum(gobjGeracaoArqCNABPag.iBancoFavorecido, 3)
        Call Extrair_Agencia_DV(gobjGeracaoArqCNABPag.sAgenciaFavorecido, sAgencia, 4, sDV, 1)
        .sAgenciaFavorecido = sAgencia
        .sDVAgenciaFavorecido = sDV
        Call Extrair_Conta_DV(gobjGeracaoArqCNABPag.sContaFavorecido, sConta, 12, sDV, 1)
        .sNumContaFavorecido = sConta
        .sDVContaFavorecido = sDV
        .sBrancos3 = FormataCpoTexto("", 2)
        .sNomeFavorecido = FormataCpoTexto(gobjGeracaoArqCNABPag.sNomeFavorecido, 30)
        .sDataProcessamento = Format(gdtDataHoje, "DDMMAA")
        .sValor = FormataCpoValor(objBaixaParcPag.dValorBaixado, 13)
''''        .sCodServico = ??? Verificar tabela FCC96001.
        .sMensagem = FormataCpoTexto("", 40)
        .sOcorrenciasRetorno = FormataCpoTexto("", 10)
        .sNumSequencialRegistro = FormataCpoNum(gobjGeracaoArqCNABPag.lQuantRegistrosArq, 6)
            
        'Monta o registro de detalhe com os dados obtidos
        sRegDetalhe = .sIdentificacaoReg & .sBrancos1 & .sIndicadorConf & .sCGCouCPFConferir & .sDVCGCouCPF & .sCGCouPrefixoEmpresa & .sDVPrefixo & .sNumContaEmpresa & .sDVNumConta & .sUsoDaEmpresa & .sBrancos2
        sRegDetalhe = sRegDetalhe & .sCodCamara & .sCodBancoDest & .sAgenciaFavorecido & .sDVAgenciaFavorecido & .sNumContaFavorecido & .sDVContaFavorecido & .sBrancos3 & .sNomeFavorecido & .sDataProcessamento
        sRegDetalhe = sRegDetalhe & .sValor & .sCodServico & .sMensagem & .sOcorrenciasRetorno & .sNumSequencialRegistro

    End With

    Montar_Registro_DetalhePagto = SUCESSO
    
    Exit Function

Erro_Montar_Registro_DetalhePagto:

    Montar_Registro_DetalhePagto = Err
    
    Select Case Err
    
        Case 51802
        
        Case 51803
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ENDERECO_NAO_CADASTRADO", Err, objEndereco.lCodigo)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147143)
            
    End Select

End Function

Function Montar_Registro_TrailerArq(sTrailerArq As String) As Long
'Monta o registro de trailer de arquivo

Dim lErro As Long
Dim tRegTrailer As typePagtoTrailerArqCNAB240

On Error GoTo Erro_Montar_Registro_TrailerArq
    
    'Incrementa a quantidade de registros do arquivo
    gobjGeracaoArqCNABPag.lQuantRegistrosArq = gobjGeracaoArqCNABPag.lQuantRegistrosArq + 1

    'Recolhe os dados do trailer
    With tRegTrailer
        .sIdentificacaoReg = "9"
        .sBrancos = FormataCpoTexto("", 197)
        .sNumSequencialRegistro = FormataCpoNum(gobjGeracaoArqCNABPag.lSequencialRegDetalhe, 6)
    
        'Monta o registro com os dados recolhidos
        sTrailerArq = .sIdentificacaoReg & .sBrancos & .sNumSequencialRegistro
    
    End With

    Montar_Registro_TrailerArq = SUCESSO
    
    Exit Function
    
Erro_Montar_Registro_TrailerArq:

    Montar_Registro_TrailerArq = Err
    
    Select Case Err
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147144)
            
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

Function Calcula_NossoNumero_BcoBrasil(objCarteiraCobrador As ClassCarteiraCobrador, sNossoNumero As String, sDVNossoNumero As String, sVariacao As String, sEspecie As String, sConvenio As String) As Long

Dim sProxNossoNumero As String

On Error GoTo Erro_Calcula_NossoNumero_BcoBrasil

    sNossoNumero = ""
    sDVNossoNumero = " "

    If ((objCarteiraCobrador.iNumCarteiraNoBanco = 12) And StrParaInt(sVariacao) = 906 Or StrParaInt(sVariacao) = 922 Or StrParaInt(sVariacao) = 949) Or (objCarteiraCobrador.iNumCarteiraNoBanco = 15) Or (objCarteiraCobrador.iNumCarteiraNoBanco = 17) Then
        sNossoNumero = FormataCpoNum(sConvenio, 6) & FormataCpoNum(gobjCobrancaEletronica.lQuantTitulos, 5)
    End If

    If Len(sNossoNumero) > 0 Then
        Call Calcula_DV_NossoNumero(sNossoNumero, sDVNossoNumero, 11, 9)
    End If

    Calcula_NossoNumero_BcoBrasil = SUCESSO
    
    Exit Function
    
Erro_Calcula_NossoNumero_BcoBrasil:

    Calcula_NossoNumero_BcoBrasil = Err
    
    Select Case Err
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, Err, 147145)
    
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

Sub RegArqCNAB_Extrai_Controle(sRegistro As String, tRegControle As typeRegControle, objCobrancaEletronica As ClassCobrancaEletronica)

    With tRegControle
    
        .sNumeroBanco = Mid(sRegistro, 1, 3)
        .sLoteServico = Mid(sRegistro, 4, 4)
        .sIdentificacaoReg = Mid(sRegistro, 8, 1)
        
    End With
    
End Sub

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147146)
     
    End Select
     
    Exit Function

End Function

Function Processa_Registro_Trailer_Lote(objLoteServico As ClassCNABLoteServico, sRegistro As String) As Long

Dim lErro As Long

On Error GoTo Erro_Processa_Registro_Trailer_Lote

    Processa_Registro_Trailer_Lote = SUCESSO
     
    Exit Function
    
Erro_Processa_Registro_Trailer_Lote:

    Processa_Registro_Trailer_Lote = Err
     
    Select Case Err
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147147)
     
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147148)
     
    End Select
     
    Exit Function

End Function

Private Function Prepara_CartCobrRet(objCobrancaEletronica As ClassCobrancaEletronica, objBorderoRetCobr As ClassBorderoRetCobr, objBorderoRetCobrCart As ClassBorderoRetCobrCart, objDetRetCobr As ClassDetRetCobr, tRegDetalhe As typeRetCobrDetalheCNAB240) As Long
'codigo comum ao tratamento de registros detalhe de baixas, custas e tarifas
'preenche objBorderoRetCobrCart e objDetRetCobr

Dim lErro As Long

On Error GoTo Erro_Prepara_CartCobrRet

    Call RegDetCobr_ObterInfo(objDetRetCobr, tRegDetalhe)
    
    objDetRetCobr.dtDataBaixa = Extrair_Data(tRegDetalhe.tU030.sDataOcorrencia)
    
    Call RetCobrCart_Obtem(objDetRetCobr, objBorderoRetCobr, objBorderoRetCobrCart)

    If objBorderoRetCobrCart Is Nothing Then
    
        Set objBorderoRetCobrCart = New ClassBorderoRetCobrCart
        
        lErro = RetCobrCart_Inicializa(objCobrancaEletronica, objBorderoRetCobr, objBorderoRetCobrCart, objDetRetCobr)
        If lErro <> SUCESSO Then Error 32255
        
        Call objBorderoRetCobr.colBorderoRetCobrCart.Add(objBorderoRetCobrCart)
                
    End If
    
    Prepara_CartCobrRet = SUCESSO
    
    Exit Function
    
Erro_Prepara_CartCobrRet:

    Prepara_CartCobrRet = Err
    
    Select Case Err
    
        Case 32255
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147149)
            
    End Select

    Exit Function

End Function

Private Function RetCobr_ConfirmaAtualizacao(objBorderoRetCobr As ClassBorderoRetCobr, bDesistiu As Boolean) As Long
'Dá chance ao usuario para desistir de atualizar bordero de retorno de cobranca conferindo os valores do creditos, custas e tarifas que serao atualizados

Dim lErro As Long, objBorderoRetCobrCart As ClassBorderoRetCobrCart
Dim dValorRecebido As Double, dValorTarifas As Double, dValorCustas As Double

On Error GoTo Erro_RetCobr_ConfirmaAtualizacao

    For Each objBorderoRetCobrCart In objBorderoRetCobr.colBorderoRetCobrCart
    
        dValorRecebido = Round(dValorRecebido + objBorderoRetCobrCart.dValorRecebido, 2)
        dValorTarifas = Round(dValorTarifas + objBorderoRetCobrCart.dValorTarifas, 2)
        dValorCustas = Round(dValorCustas + objBorderoRetCobrCart.dValorCustas, 2)
        
    Next

    bDesistiu = False
    
    '??? abrir form pedindo confirmacao
    If MsgBox("Atenção: confirma os valores recebido " & Format(dValorRecebido, "standard") & " , de tarifas " & Format(dValorTarifas, "standard") & " , de custas " & Format(dValorCustas, "standard") & " e prossegue com a atualização ?", vbYesNo, "Bordero de Retorno de Cobrança") <> vbYes Then
    
        bDesistiu = True
        
    End If
    
    RetCobr_ConfirmaAtualizacao = SUCESSO
    
    Exit Function
    
Erro_RetCobr_ConfirmaAtualizacao:

    RetCobr_ConfirmaAtualizacao = Err
    
    Select Case Err
    
        Case 32259
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147150)
            
    End Select

    Exit Function

End Function

Private Function RetCobr_Obtem_Parcela(objDetRetCobr As ClassDetRetCobr, objParcelaRec As ClassParcelaReceber, tRegDetalhe As typeRetCobrDetalheCNAB240, lComando As Long) As Long
'preenche objParcelaRec lendo do bd identificando a parcela por dados em objDetRetCobr

Dim lErro As Long, dValorTitulo As Double, lNumTitulo As Long, iNumParcela As Integer, sParcela As String
Dim lNumIntDoc As Long, iPosSep As Integer
Dim sFiltro As String

On Error GoTo Erro_RetCobr_Obtem_Parcela

    iPosSep = InStr(tRegDetalhe.tT030.sNumDoc, SEPARADOR)
    If iPosSep <> 0 Then
    
        lNumTitulo = StrParaLong(Mid(tRegDetalhe.tT030.sNumDoc, 1, iPosSep - 1))
        sParcela = Mid(tRegDetalhe.tT030.sNumDoc, iPosSep + 1, 2)
        
    Else
    
        If IsNumeric(tRegDetalhe.tT030.sNumDoc) Then
        
            lNumTitulo = StrParaLong(Mid(tRegDetalhe.tT030.sNumDoc, 1, 6))
            sParcela = "1"
            
        Else
        
            objDetRetCobr.iTipoCritica = RETCOBR_CRITICA_SEM_PARC
            Error 32269
        
        End If
        
    End If
    
    Call cf("FilialEmpresa_Le_Filtro_Filial", sFiltro)
    
    lErro = Comando_Executar(lComando, "SELECT ParcelasRec.NumIntDoc FROM ParcelasRec, TitulosRec WHERE ParcelasRec.NumIntTitulo = TitulosRec.NumIntDoc AND TitulosRec.NumTitulo = ? AND DataVencimento = ? AND ParcelasRec.Status <> ? " & sFiltro, lNumIntDoc, lNumTitulo, objDetRetCobr.dtDataVencimento, STATUS_EXCLUIDO)
    If lErro <> AD_SQL_SUCESSO Then Error 32267

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 32268
    
    If lErro = AD_SQL_SEM_DADOS Then
    
        '??? poderia abrir janela p/selecionar parcela passando as informacoes do registro
        
        objDetRetCobr.iTipoCritica = RETCOBR_CRITICA_SEM_PARC
        Error 32269
    
    Else
    
        objParcelaRec.lNumIntDoc = lNumIntDoc
    
        'Lê a parcela a receber
        lErro = cf("ParcelaReceber_Le", objParcelaRec)
        If lErro <> SUCESSO And lErro <> 19147 Then Error 51683
        If lErro <> SUCESSO Then Error 51697
    
        'poderia testar se a parcela está no cobrador/carteira devidos
        '???
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 32270
        
        'se há mais de uma parcela
        If lErro <> AD_SQL_SEM_DADOS Then
        
            If IsNumeric(sParcela) Then
                iNumParcela = CInt(sParcela)
            Else
                If sParcela >= "A" And sParcela <= "Z" Then
                    iNumParcela = Asc(sParcela) - Asc("A") + 1
                Else
                    If sParcela >= "a" And sParcela <= "z" Then
                        iNumParcela = Asc(sParcela) - Asc("a") + 1
                    Else
                        iNumParcela = 0
                    End If
                End If
            End If
        
            'se sei o # da parcela
            If iNumParcela <> 0 Then
                
                lErro = Comando_Executar(lComando, "SELECT ParcelasRec.NumIntDoc FROM ParcelasRec, TitulosRec WHERE ParcelasRec.NumIntTitulo = TitulosRec.NumIntDoc AND TitulosRec.NumTitulo = ? AND DataVencimento = ? AND NumParcela = ? AND ParcelasRec.Status <> ? " & sFiltro, lNumIntDoc, lNumTitulo, objDetRetCobr.dtDataVencimento, iNumParcela, STATUS_EXCLUIDO)
                If lErro <> AD_SQL_SUCESSO Then Error 32267
                
                lErro = Comando_BuscarPrimeiro(lComando)
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 32268
        
                objParcelaRec.lNumIntDoc = lNumIntDoc
            
                'Lê a parcela a receber
                lErro = cf("ParcelaReceber_Le", objParcelaRec)
                If lErro <> SUCESSO And lErro <> 19147 Then Error 51683
                If lErro <> SUCESSO Then Error 51697
    
                lErro = Comando_BuscarProximo(lComando)
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 32270
                
                'se há mais de uma parcela
                If lErro <> AD_SQL_SEM_DADOS Then
                
                    '??? poderia abrir janela p/selecionar parcela passando as informacoes do registro
                    objDetRetCobr.iTipoCritica = RETCOBR_CRITICA_VARIAS_PARC
                    Error 32271
                    
                End If
                
            Else
            
                '??? poderia abrir janela p/selecionar parcela passando as informacoes do registro
                objDetRetCobr.iTipoCritica = RETCOBR_CRITICA_VARIAS_PARC
                Error 32271
            
            End If
            
        End If
        
    End If
    
    RetCobr_Obtem_Parcela = SUCESSO

    Exit Function

Erro_RetCobr_Obtem_Parcela:

    RetCobr_Obtem_Parcela = Err

    Select Case Err

        Case 32269
            'Call Rotina_Erro(vbOKOnly, "ERRO_PARCBORDRETCOBR_NENHUMA", Err, tRegDetalheItau.sSeuNumero)
        
        Case 32271
            'Call Rotina_Erro(vbOKOnly, "ERRO_PARCBORDRETCOBR_VARIAS", Err, tRegDetalheItau.sSeuNumero)

        Case 32267, 32268, 32270
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PARCREC_RETCOBR", Err)
        
        Case 51683
        
        Case 51697
            Call Rotina_Erro(vbOKOnly, "ERRO_PARCELAREC_NAO_CADASTRADA", Err, lNumIntDoc)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147151)
            
    End Select

    Exit Function

End Function

Function RetCobr_Grava_Erros_Trans(objBorderoRetCobr As ClassBorderoRetCobr) As Long

Dim lErro As Long, objDetRetCobr As ClassDetRetCobr, lComando As Long

On Error GoTo Erro_RetCobr_Grava_Erros_Trans

    'Abre os comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 32277
    
    lErro = Comando_Executar(lComando, "DELETE FROM RetCobrErros WHERE Banco = ? AND NomeArq = ? AND DataArq = ?", objBorderoRetCobr.iBanco, objBorderoRetCobr.sNomeArq, objBorderoRetCobr.dtDataRecepcao)
    If lErro <> AD_SQL_SUCESSO Then Error 32278
    
    For Each objDetRetCobr In objBorderoRetCobr.colRetCobrErros
    
        lErro = Comando_Executar(lComando, "INSERT INTO RetCobrErros (Banco, NomeArq, DataArq, SeqReg, CodCobrador, NumCarteiraBanco, Ocorrencia, DataVencimento, ValorTitulo, ValorRecebido, ValorJuros, ValorDesconto, ValorTarifa, SeuNumero, NossoNumero, UsoDaEmpresa, DataCredito, TipoCritica) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
            objBorderoRetCobr.iBanco, objBorderoRetCobr.sNomeArq, objBorderoRetCobr.dtDataCredito, objDetRetCobr.lNumSequencial, objBorderoRetCobr.iCobrador, objDetRetCobr.iCodCarteiraNoBanco, objDetRetCobr.iCodOcorrencia, objDetRetCobr.dtDataVencimento, objDetRetCobr.dValorTitulo, objDetRetCobr.dValorRecebido, objDetRetCobr.dValorJuros, objDetRetCobr.dValorDesconto, objDetRetCobr.dValorTarifa, objDetRetCobr.sSeuNumero, objDetRetCobr.sNossoNumero, objDetRetCobr.sUsoDaEmpresa, objDetRetCobr.dtDataCredito, objDetRetCobr.iTipoCritica)
        If lErro <> AD_SQL_SUCESSO Then Error 32279
            
    Next
    
    Call Comando_Fechar(lComando)
    
    RetCobr_Grava_Erros_Trans = SUCESSO

    Exit Function

Erro_RetCobr_Grava_Erros_Trans:

    RetCobr_Grava_Erros_Trans = Err

    Select Case Err

        Case 32277
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 32278
            Call Rotina_Erro(vbOKOnly, "ERRO_RETCOBR_EXCLUSAO_ERROS", Err)

        Case 32279
            Call Rotina_Erro(vbOKOnly, "ERRO_RETCOBR_INCLUSAO_ERRO", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147152)
            
    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function RetCobr_Grava_Erros(objBorderoRetCobr As ClassBorderoRetCobr) As Long

Dim lErro As Long, lTransacao As Long

On Error GoTo Erro_RetCobr_Grava_Erros

    'Entra em transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 32274

    lErro = RetCobr_Grava_Erros_Trans(objBorderoRetCobr)
    If lErro <> SUCESSO Then Error 32275
    
    'Confirma transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 32276

    RetCobr_Grava_Erros = SUCESSO

    Exit Function

Erro_RetCobr_Grava_Erros:

    RetCobr_Grava_Erros = Err

    Select Case Err

        Case 32275
        
        Case 32274
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 32276
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147153)
            
    End Select

    Call Transacao_Rollback
    
    Exit Function

End Function

'??? tem que voltar a ser funcao comum em classcnabcobrret
Private Function Processa_ParcelaRec(objParcelaRec As ClassParcelaReceber, objBorderoRetCobr As ClassBorderoRetCobr, objBorderoRetCobrCart As ClassBorderoRetCobrCart, objDetRetCobr As ClassDetRetCobr, objInfoParcRec As ClassInfoParcRec) As Long
'Faz o processamento da parcela a receber

Dim lErro As Long, dValorBaixado As Double
Dim objOcorrRetParcRec As New ClassOcorrRetParcRec

On Error GoTo Erro_Processa_ParcelaRec

    'Acumula-se os valores creditados para o criação do movimento de conta corrente
    If objParcelaRec.iCarteiraCobranca <> CARTEIRA_DESCONTADA And objBorderoRetCobrCart.dtDataCredito >= objBorderoRetCobr.dtDataCredito Then
        objBorderoRetCobrCart.dValorRecebido = Round(objBorderoRetCobrCart.dValorRecebido + objDetRetCobr.dValorRecebido, 2)
        objBorderoRetCobrCart.dValorJuros = Round(objBorderoRetCobrCart.dValorJuros + objDetRetCobr.dValorJuros, 2)
        objBorderoRetCobrCart.dValorDescontos = Round(objBorderoRetCobrCart.dValorDescontos + objDetRetCobr.dValorDesconto, 2)
    End If
    
    dValorBaixado = Round(objDetRetCobr.dValorRecebido + objDetRetCobr.dValorDesconto - objDetRetCobr.dValorJuros, 2)
           
    'se a baixa nao foi total incluir o reg p/sair no relatorio de criticas
    If Abs(dValorBaixado - objParcelaRec.dSaldo) > 0.01 Then
    
        objDetRetCobr.iTipoCritica = RETCOBR_CRITICA_BAIXA_PARCIAL
        Call objBorderoRetCobr.colRetCobrErros.Add(objDetRetCobr)
        
    End If
    
    'Gera uma baixaparcrec p\ a parcela a receber em questão
    objBorderoRetCobrCart.ColBaixaParcRec.Add 0, 0, objParcelaRec.lNumIntDoc, 0, STATUS_LANCADO, 0, objDetRetCobr.dValorJuros, objDetRetCobr.dValorDesconto, dValorBaixado, objDetRetCobr.dValorRecebido, objBorderoRetCobr.iCobrador
    
    'guarda a informacao que será util p/gerar a contabilizacao
    objBorderoRetCobrCart.colInfoParcRec.Add objInfoParcRec
    
    'Gera um ocorrenciaRetParcrec de acordo com os dados retornados
    objOcorrRetParcRec.dtDataComplementar = DATA_NULA
    objOcorrRetParcRec.lNumIntParc = objParcelaRec.lNumIntDoc
    objOcorrRetParcRec.lNumSeqOcorr = 0
    objOcorrRetParcRec.sNumTitCobrador = objDetRetCobr.sNossoNumero
    objOcorrRetParcRec.iCodOcorrencia1 = objDetRetCobr.iCodOcorrencia
    objOcorrRetParcRec.dtDataComplementar = DATA_NULA
        
    'Adiciona a ocorrência na coleção
    objBorderoRetCobrCart.colOcorrRetParcRec.Add objOcorrRetParcRec
        
    Processa_ParcelaRec = SUCESSO
    
    Exit Function

Erro_Processa_ParcelaRec:

    Processa_ParcelaRec = Err
    
    Select Case Err
    
        Case 51691, 51692
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147154)
            
    End Select

    Exit Function

End Function

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

Function RetCobrCart_Inicializa(objCobrancaEletronica As ClassCobrancaEletronica, objBorderoRetCobr As ClassBorderoRetCobr, objBorderoRetCobrCart As ClassBorderoRetCobrCart, objDetRetCobr As ClassDetRetCobr)
'inicializa carteira de cobranca

Dim lErro As Long, objCarteiraCobrador As ClassCarteiraCobrador
Dim objContaCorrente As New ClassContasCorrentesInternas
    
On Error GoTo Erro_RetCobrCart_Inicializa

    objBorderoRetCobrCart.iCobrador = objBorderoRetCobr.iCobrador
    objBorderoRetCobrCart.iCodCarteiraNoBanco = objDetRetCobr.iCodCarteiraNoBanco

    Call objBorderoRetCobr.CarteiraCobranca_Obter(objBorderoRetCobrCart.iCodCarteiraNoBanco, objCarteiraCobrador)
    If objCarteiraCobrador Is Nothing Then Error 32266
    
    objBorderoRetCobrCart.iCarteiraCobr = objCarteiraCobrador.iCodCarteiraCobranca
    
    'Pega a conta corrente do cobrador
    objContaCorrente.iCodigo = objCobrancaEletronica.objCobrador.iCodCCI
    
    'Lê a conta corrente do cobrador
    lErro = cf("ContaCorrenteInt_Le", objContaCorrente.iCodigo, objContaCorrente)
    If lErro <> SUCESSO And lErro <> 11807 Then Error 51711
    
    'Se não achou a Conta Corrente --> erro
    If lErro <> SUCESSO Then Error 51712
    
    'Preenche parcialmente o objMovCCI
    With objBorderoRetCobrCart.objMovCCI
    
        .iFilialEmpresa = objContaCorrente.iFilialEmpresa
        .iCodConta = objContaCorrente.iCodigo
        .iExcluido = NAO_EXCLUIDO
        .iTipoMeioPagto = DINHEIRO
        .dtDataBaixa = objDetRetCobr.dtDataBaixa
        .dtDataMovimento = IIf(objDetRetCobr.dtDataCredito <> DATA_NULA, objDetRetCobr.dtDataCredito, objBorderoRetCobr.dtDataCredito)
        .dtDataContabil = IIf(objDetRetCobr.dtDataCredito <> DATA_NULA, objDetRetCobr.dtDataCredito, objBorderoRetCobr.dtDataCredito)
        .iConciliado = NAO_CONCILIADO
           
    End With
    
    objBorderoRetCobrCart.dtDataCredito = objBorderoRetCobrCart.objMovCCI.dtDataMovimento
    
    RetCobrCart_Inicializa = SUCESSO
    
    Exit Function
    
Erro_RetCobrCart_Inicializa:

    RetCobrCart_Inicializa = Err
    
    Select Case Err

        Case 32266
            Call Rotina_Erro(vbOKOnly, "ERRO_CARTEIRA_COBR_INVALIDA", Err, objBorderoRetCobrCart.iCodCarteiraNoBanco, objBorderoRetCobr.iCobrador)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147155)

    End Select

    Exit Function

End Function

Public Function CobrancaEletronica_ObtemNumTitulos(objCobrancaEletronica As ClassCobrancaEletronica, lQuantRegistros As Long) As Long

Dim lErro As Long
Dim sRegistro As String, objLoteServico As New ClassCNABLoteServico
Dim tRegControle As typeRegControle

On Error GoTo Erro_CobrancaEletronica_ObtemNumTitulos

    'Abre o arquivo de retorno
    Open objCobrancaEletronica.sNomeArquivoRetorno For Input As #1
    
    'Até chegar ao fim do arquivo
    Do While Not EOF(1)
    
        'Busca o próximo registro do arquivo (na 1a vez vai ser o de header)
        Line Input #1, sRegistro
    
        'Pega o "controle" do registro
        Call RegArqCNAB_Extrai_Controle(sRegistro, tRegControle, objCobrancaEletronica)
        
        Select Case tRegControle.sIdentificacaoReg
            
            Case "0" 'Header de Arquivo
            
            Case "1" 'header de lote
            
                With objLoteServico
                    .sLote = Mid(sRegistro, 4, 4)
                    .sTipoOperacao = Mid(sRegistro, 9, 1)
                    .sTipoServico = Mid(sRegistro, 10, 2)
                    .sFormaLcto = Mid(sRegistro, 12, 2)
                End With
    
            Case "3" 'Detalhe ("Transação")
                
                'se for retorno de cobranca
                If objLoteServico.sTipoOperacao = "T" And objLoteServico.sTipoServico = "01" Then
                
                    'dependendo do tipo de segmento
                    If Mid(sRegistro, 14, 1) = "T" Then
                        
                        lQuantRegistros = lQuantRegistros + 1
                        
                        If lQuantRegistros = 1 Then
                        
                            If StrParaInt(Mid(sRegistro, 1, 3)) <> objCobrancaEletronica.objCobrador.iCodBanco Then gError 32249
                            
                        End If
        
                    End If
                    
                End If
                
            Case "5" 'Trailer de lote
            
            Case "9" 'Trailer de arquivo
        
        End Select
        
    Loop
        
    Close #1
    
    CobrancaEletronica_ObtemNumTitulos = SUCESSO
     
    Exit Function
    
Erro_CobrancaEletronica_ObtemNumTitulos:

    CobrancaEletronica_ObtemNumTitulos = gErr
     
    Select Case gErr
          
        Case 32250
            Call Rotina_Erro(vbOKOnly, "ERRO_ARQRETCOBR_ARQ_INVALIDO", gErr)
    
        Case 32249
            Call Rotina_Erro(vbOKOnly, "ERRO_ARQRETCOBR_BCO_NAO_CONFERE", gErr, StrParaInt(Mid(sRegistro, 1, 3)), objCobrancaEletronica.objCobrador.iCodBanco)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147156)
     
    End Select
     
    Close #1
    
    Exit Function

End Function

Private Function Obter_Dados_DetalheSegP(tRegDetalheBcoBrasil240SegP As typeRemDetalheBcoBrasil240SegP, objOcorrRemParcRem As ClassOcorrRemParcRec, objCarteiraCobrador As ClassCarteiraCobrador, objParcelaRec As ClassParcelaReceber) As Long
'Obtem os dados que são utilizados para preencher o registro de detalhe

Dim lErro As Long
Dim sNossoNumero As String
Dim dValorDesconto As Double
Dim sDVNossoNumero As String

On Error GoTo Erro_Obter_Dados_DetalheSegP

With tRegDetalheBcoBrasil240SegP
        
    .sCodBancoComp = "001"
    .sLoteServico = FormataCpoNum(giNumLotes, 4)
    .sRegDetalhe = "3"
    .sNumSequencialReg = FormataCpoNum(gobjCobrancaEletronica.lSequencialRegistro, 5)
    .sCodSegRegDetalhe = "P"
    .sBrancosCNAB1 = FormataCpoTexto("", 1)
    .sCodMovimento = "01"
    .sAgMantConta = FormataCpoNum(gobjCobrancaEletronica.objCCI.sAgencia, 5)
    .sDVAgMantConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVAgencia, 1)
    .sNumContaCorrente = FormataCpoNum(gobjCobrancaEletronica.objCCI.sNumConta, 12)
    .sDVConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVNumConta, 1)
    .sDVAgenciaConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVAgConta, 1)
    '.sIdentTituloBanco Calculado depois.
    .sCodCarteira = FormataCpoNum(objCarteiraCobrador.iNumCarteiraNoBanco, 1)
    .sFormaCadTituloBanco = FormataCpoNum(IIf(objCarteiraCobrador.iComRegistro = 0, "2", "1"), 1)
    .sTipoDocumento = "2"
    .sIdentEmissaoBloq = FormataCpoNum(IIf(objCarteiraCobrador.iImprimeBoleta = 0, "1", "2"), 1)
    .sIdentDistrib = "1"
    
    .sNumDocCobranca = FormataCpoTexto(gobjCobrancaEletronica.lNumTitulo & SEPARADOR & objParcelaRec.iNumParcela, 15)
    
    .sDataVencTitulo = Format(objParcelaRec.dtDataVencimento, "DDMMYYYY")
    .sValorNominalTitulo = FormataCpoValor(objOcorrRemParcRem.dValorCobrado, 15)
    .sAgEncCobranca = FormataCpoNum("", 5)
    .sDVAgencia = FormataCpoTexto("0", 1)
    
    .sEspecieTitulo = FormataCpoNum(2, 2)
    
    If objParcelaRec.iAceite = 1 Then
        .sIdentTituloAceite = "A"
    Else
        .sIdentTituloAceite = "N"
    End If
    
    .sDataEmissaoTitulo = Format(gobjCobrancaEletronica.dtDataEmissaoTitulo, "DDMMYYYY")
    .sCodJurosMora = "1"
    .sDataJurosMora = FormataCpoNum(0, 8)
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
        dValorDesconto = objParcelaRec.dDesconto1Valor
    End If
    
    .sValorPerConcedido = FormataCpoValor(dValorDesconto, 15)
    .sValorIOF = FormataCpoValor(0, 15)
    .sValorAbatimento = FormataCpoValor(0, 15)
    .sIdentTituloEmp = FormataCpoTexto(gobjCobrancaEletronica.lNumTitulo & SEPARADOR & objParcelaRec.iNumParcela, 25)
        
    'se for cobranca caucionada => nao protesta
    If objCarteiraCobrador.iNumCarteiraNoBanco = 3 Then
    
        .sCodProtesto = "3"
        .sNumDiasProtesto = FormataCpoNum(0, 2)
        
    End If
    
    'se for cobranca simples => protesta apos 5 dias corridos
    If objCarteiraCobrador.iNumCarteiraNoBanco = 1 Then
    
        .sCodProtesto = "1"
        .sNumDiasProtesto = FormataCpoNum(5, 2)
        
    End If
    
    Call cf("BB_Obter_Dados_DetalheSegP_Customizado", gobjCobrancaEletronica, objOcorrRemParcRem, objCarteiraCobrador, objParcelaRec, .sTipoDocumento, .sNumDocCobranca, .sDVAgencia, .sCodProtesto, .sNumDiasProtesto)
    
    .sCodBaixaDevolucao = "2"
    .sNumDiasBaixaDevol = FormataCpoNum(0, 3)
    
'    'Verifica qual será o Código do Protesto
'    If objOcorrRemParcRem.iInstrucao1 = 81 Or objOcorrRemParcRem.iInstrucao1 = 82 Then
'
'        .sCodProtesto = IIf(objOcorrRemParcRem.iInstrucao1 = 82, "1", "2")
'        .sNumDiasProtesto = FormataCpoNum(objOcorrRemParcRem.iDiasDeProtesto1, 2)
'
'    ElseIf objOcorrRemParcRem.iInstrucao2 = 81 Or objOcorrRemParcRem.iInstrucao2 = 82 Then
'
'        .sCodProtesto = IIf(objOcorrRemParcRem.iInstrucao2 = 82, "1", "2")
'        .sNumDiasProtesto = FormataCpoNum(objOcorrRemParcRem.iDiasDeProtesto2, 2)
'    Else
'        .sCodProtesto = "3"
'        .sNumDiasProtesto = FormataCpoNum(0, 2)
'    End If
'
'    'Verifica qual será o Código para Baixa/Devolução
'    If objOcorrRemParcRem.iInstrucao1 >= 17 And objOcorrRemParcRem.iInstrucao1 <= 29 Or objOcorrRemParcRem.iInstrucao1 = 92 Then
'        .sCodBaixaDevolucao = "1"
'        .sNumDiasBaixaDevol = FormataCpoNum(objOcorrRemParcRem.iDiasDeProtesto1, 3)
'
'    ElseIf objOcorrRemParcRem.iInstrucao2 >= 17 And objOcorrRemParcRem.iInstrucao2 <= 29 Or objOcorrRemParcRem.iInstrucao2 = 92 Then
'        .sCodBaixaDevolucao = "1"
'        .sNumDiasBaixaDevol = FormataCpoNum(objOcorrRemParcRem.iDiasDeProtesto2, 3)
'    Else
'        .sCodBaixaDevolucao = "2"
'        .sNumDiasBaixaDevol = FormataCpoNum(objOcorrRemParcRem.iDiasDeProtesto2, 3)
'    End If
'
    'FIM INPAL
        
    .sCodMoeda = "09"
    .sNumContratoOpCred = FormataCpoNum(0, 10)
    .sBrancosCNAB2 = FormataCpoTexto("", 1)
    
    'Está faltando o campo Varição e o DVNossoNumero
    'lErro = Calcula_NossoNumero_BcoBrasil(objCarteiraCobrador, sNossoNumero, sDVNossoNumero, .sVariacao, .sEspecieTitulo, gsNumConvenio)
       
    .sIdentTituloBanco = FormataCpoNum("", 20) '?????
    
    End With
    
    Obter_Dados_DetalheSegP = SUCESSO

    Exit Function
    
Erro_Obter_Dados_DetalheSegP:

Obter_Dados_DetalheSegP = Err
    
    Select Case Err

        Case 51817, 51818, 51819
            MsgBox (CStr(Err))
       
        Case 51820
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ENDERECO_FILIALCLIENTE_NAO_INFORMADO", Err, gobjCobrancaEletronica.lCliente, gobjCobrancaEletronica.iFilialCli)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147157)

    End Select

    Exit Function

End Function

Private Function Obter_Dados_DetalheSegQ(tRegDetalheBcoBrasil240SegQ As typeRemDetalheBcoBrasil240SegQ, objOcorrRemParcRem As ClassOcorrRemParcRec, objCarteiraCobrador As ClassCarteiraCobrador, objParcelaRec As ClassParcelaReceber) As Long
'Obtem os dados que são utilizados para preencher o registro de detalhe

Dim lErro As Long
Dim objEndereco As New ClassEndereco
Dim sNossoNumero As String
Dim dValorDesconto As Double
Dim sDVNossoNumero As String

On Error GoTo Erro_Obter_Dados_DetalheSegQ

    lErro = cf("CobrancaEletronica_Obter_Endereco", gobjCobrancaEletronica, objEndereco)
    If lErro <> SUCESSO Then Error 51818
        
    With tRegDetalheBcoBrasil240SegQ
           
        .sCodBancoComp = "001"
        .sLoteServico = FormataCpoNum(giNumLotes, 4)
        .sRegDetalhe = "3"
        .sNumSequencialReg = FormataCpoNum(gobjCobrancaEletronica.lSequencialRegistro, 5)
        .sCodSegRegDetalhe = "Q"
        .sBrancosCNAB1 = FormataCpoTexto("", 1)
        .sCodMovimento = "01"
    
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

        Case 51817, 51818, 51819
            MsgBox (CStr(Err))
       
        Case 51820
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ENDERECO_FILIALCLIENTE_NAO_INFORMADO", Err, gobjCobrancaEletronica.lCliente, gobjCobrancaEletronica.iFilialCli)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147158)

    End Select

    Exit Function

End Function

Private Function Obter_Dados_DetalheSegR(tRegDetalheBcoBrasil240SegR As typeRemDetalheBcoBrasil240SegR, objOcorrRemParcRem As ClassOcorrRemParcRec, objCarteiraCobrador As ClassCarteiraCobrador, objParcelaRec As ClassParcelaReceber) As Long
'Obtem os dados que são utilizados para preencher o registro de detalhe

Dim lErro As Long
Dim objEndereco As New ClassEndereco
Dim sNossoNumero As String
Dim dValorDesconto As Double
Dim sDVNossoNumero As String
Dim o As ClassTituloReceber

On Error GoTo Erro_Obter_Dados_DetalheSegR

    With tRegDetalheBcoBrasil240SegR
       
        .sCodBancoComp = "001"
        .sLoteServico = FormataCpoNum(giNumLotes, 4)
        .sRegDetalhe = "3"
        .sNumSequencialReg = FormataCpoNum(gobjCobrancaEletronica.lSequencialRegistro, 5)
        .sCodSegRegDetalhe = "R"
        .sBrancosCNAB1 = FormataCpoTexto("", 1)
        .sCodMovimento = "01"
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
            dValorDesconto = objParcelaRec.dDesconto2Valor
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
            dValorDesconto = objParcelaRec.dDesconto3Valor
        End If
        
        .sValorPerConcedido3 = FormataCpoValor(dValorDesconto, 15)
        .sCodMulta = "2"
        .sDataMulta = FormataCpoNum(0, 8)
        .sValorPerAplicado = FormataCpoValor(gobjCobrancaEletronica.dPercMulta * 100, 15)
        .sInfoBancoSacado = FormataCpoTexto("", 10)
        .sMensagem3 = FormataCpoTexto("", 40)
        .sMensagem4 = FormataCpoTexto("", 40)
        'Alterado por cyntia em 28/05/02
        .sCodBancoContaDeb = FormataCpoTexto("", 3)
        .sCodAgContaDeb = FormataCpoTexto("", 4)
        .sCodContaDVDeb = FormataCpoTexto("", 13)
        .sCodOcorrenciaSac = FormataCpoTexto("", 8)
        .sBrancosCNAB2 = FormataCpoTexto("", 33)
        
    End With
    
    Obter_Dados_DetalheSegR = SUCESSO

    Exit Function
    
Erro_Obter_Dados_DetalheSegR:

Obter_Dados_DetalheSegR = Err
    
    Select Case Err

        Case 51817, 51818, 51819
            MsgBox (CStr(Err))
       
        Case 51820
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ENDERECO_FILIALCLIENTE_NAO_INFORMADO", Err, gobjCobrancaEletronica.lCliente, gobjCobrancaEletronica.iFilialCli)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147159)

    End Select

    Exit Function

End Function

Private Function Obter_Dados_Header_Arquivo(tRegHeaderBcoBrasil240 As typeRemHeaderBcoBrasil240) As Long

Dim lErro As Long

On Error GoTo Erro_Obter_Dados_Header_Arquivo
      
    'Carrega os campos do cabeçalho
    With tRegHeaderBcoBrasil240
   
        .sCodBancoComp = FormataCpoNum("001", 3)
        .sLoteServico = FormataCpoNum(giNumLotes, 4)
        .sRegHeaderLote = "0"
        .sNVersaoLayLote = "030"
        .sBrancosCNAB1 = FormataCpoTexto("", 9)
        
        If Len(gobjFilialEmpresa.sCgc) = 14 Then
            .sTipoInscEmpresa = "2"
        Else
            .sTipoInscEmpresa = "1"
        End If
        
        .sCodInsEmpresa = FormataCpoNum(gobjFilialEmpresa.sCgc, 14)
        .sCodConvBanco = FormataCpoTexto(Obter_Info_CarteiraCobrador(CARTEIRA_SIMPLES, gobjCobrancaEletronica.colCarteirasCNABInfo, 7), 20)
        .sAgMantConta = FormataCpoNum(gobjCobrancaEletronica.objCCI.sAgencia, 5)
        .sDVAgencia = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVAgencia, 1)
        .sNumContaCorrente = FormataCpoNum(gobjCobrancaEletronica.objCCI.sNumConta, 12)
        .sDVConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVNumConta, 1)
        .sDVAgenciaConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVAgConta, 1)
        
        .sNomeEmpresa = FormataCpoTexto(gsNomeEmpresa, 30)
        
        Call cf("Empresa_Le_Nome_Customizado", .sNomeEmpresa)
        
        .sMensagem1 = FormataCpoTexto("BANCO DO BRASIL", 30)
        .sMensagem2 = FormataCpoTexto("", 10)
        .sNumRemessa = FormataCpoNum(gobjCobrancaEletronica.objCobrador.lCNABProxSeqArqCobr, 6)
        .sDataGravRem = Format(gdtDataHoje, "DDMMYYYY")
    
    End With
   
    Obter_Dados_Header_Arquivo = SUCESSO

    Exit Function
    
Erro_Obter_Dados_Header_Arquivo:

    Obter_Dados_Header_Arquivo = Err
    
    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147160)
            
    End Select
    
    Exit Function

End Function

Private Function Montar_RegistroHeaderArquivo_BcoBrasil(sRegHeader As String) As Long
'Monta o registro

Dim tRegHeaderBcoBrasil240 As typeRemHeaderBcoBrasil240
Dim lErro As Long

On Error GoTo Erro_Montar_RegistroHeaderArquivo_BcoBrasil

    'Carrega tRegHeaderBcoBrasil com os dados do Header do aquivo
    lErro = Obter_Dados_Header_Arquivo(tRegHeaderBcoBrasil240)
    If lErro <> SUCESSO Then Error 51813
        
    With tRegHeaderBcoBrasil240
    
        'Monta o registro de Header do BcoBrasil
        sRegHeader = .sCodBancoComp & .sLoteServico & .sRegHeaderLote & .sBrancosCNAB1
        sRegHeader = sRegHeader & .sTipoInscEmpresa & .sCodInsEmpresa & .sCodConvBanco & .sAgMantConta & .sDVAgencia & .sNumContaCorrente & .sDVConta & .sDVAgenciaConta
        sRegHeader = sRegHeader & .sNomeEmpresa & .sMensagem1 & .sMensagem2 & "1" & .sDataGravRem & "120000" & .sNumRemessa & .sNVersaoLayLote & "00000" & FormataCpoTexto("", 54) & "000" & "  " & FormataCpoTexto("", 10)
        
    End With
    
    Montar_RegistroHeaderArquivo_BcoBrasil = SUCESSO

    Exit Function

Erro_Montar_RegistroHeaderArquivo_BcoBrasil:

    Montar_RegistroHeaderArquivo_BcoBrasil = Err

    Select Case Err

        Case 51813

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147161)

    End Select

    Exit Function

End Function

Private Function Montar_RegistroTrailerArq_BcoBrasil(sRegTrailer As String) As Long
'Monta o registro

    glNumRegistros = glNumRegistros + 1
        
    sRegTrailer = "00199999         " & FormataCpoNum(giNumLotes, 6) & FormataCpoNum(glNumRegistros, 6) & FormataCpoNum(0, 6) & FormataCpoTexto("", 205)
    
    Montar_RegistroTrailerArq_BcoBrasil = SUCESSO

    Exit Function

Erro_Montar_RegistroTrailerArq_BcoBrasil:

    Montar_RegistroTrailerArq_BcoBrasil = Err

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147162)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147163)
     
    End Select
     
    Exit Function

End Function

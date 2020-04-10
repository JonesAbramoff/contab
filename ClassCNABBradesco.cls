VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCNABBradesco"
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

Function Criar_Arquivo_Remessa_Bradesco(objCobrancaEletronica As ClassCobrancaEletronica) As Long
'Cria uma arquivo de remessa com base nos dados do BD

Dim lErro As Long
Dim lTransacao As Long
Dim colBorderos As New Collection
Dim colBorderosDesc As New Collection

On Error GoTo Erro_Criar_Arquivo_Remessa_Bradesco
           
    'ABre a transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 51886
    
    glNumArquivoRemessa = 0
    
    'Separa os borderos
    lErro = Separa_Borderos(objCobrancaEletronica.colBorderos, colBorderos, colBorderosDesc)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If colBorderos.Count > 0 Then
    
        Set objCobrancaEletronica.colBorderos = colBorderos
    
        objCobrancaEletronica.iCarteira = CARTEIRA_SIMPLES
    
        lErro = Criar_Arquivo_Remessa_BradescoAux(objCobrancaEletronica)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    End If
    
    If colBorderosDesc.Count > 0 Then
        
        Set objCobrancaEletronica.colBorderos = colBorderosDesc
        
        objCobrancaEletronica.iCarteira = CARTEIRA_DESCONTADA
        
        lErro = Criar_Arquivo_Remessa_BradescoAux(objCobrancaEletronica)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    End If
    
    lErro = Transacao_Commit()
    If lErro <> SUCESSO Then gError 51890
    
    Criar_Arquivo_Remessa_Bradesco = SUCESSO
    
    Exit Function

Erro_Criar_Arquivo_Remessa_Bradesco:

    Criar_Arquivo_Remessa_Bradesco = gErr
    
    Select Case gErr
    
        Case 51886
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 51890
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 147397)
    
    End Select

    Call Transacao_Rollback
   
    Exit Function
    
End Function

Private Function Separa_Borderos(colBorderos As Collection, colBorderosRet As Collection, colBorderosRetDesc As Collection) As Long
'Separa os borderos das carteiras de desconto das outras

Dim lErro As Long
Dim lComando As Long
Dim iRegistro As Integer
Dim objCarteiraCob As New ClassCarteiraCobrador
Dim objBordero As New ClassBorderoCobranca

On Error GoTo Erro_Separa_Borderos
    
    'Para cada bordero da coleção lida
    For Each objBordero In colBorderos
        
        objCarteiraCob.iCobrador = objBordero.iCobrador
        objCarteiraCob.iCodCarteiraCobranca = objBordero.iCodCarteiraCobranca
        'Lê a carteira do cobrdor
        lErro = CF("CarteiraCobrador_Le", objCarteiraCob)
        If lErro <> SUCESSO And lErro <> 23551 Then gError ERRO_SEM_MENSAGEM
        If lErro <> SUCESSO Then gError 62016 'Não encontrou
        
        If objCarteiraCob.iCodCarteiraCobranca = CARTEIRA_DESCONTADA Then
            colBorderosRetDesc.Add objBordero
        Else
            colBorderosRet.Add objBordero
        End If
    
    Next
    
    Separa_Borderos = SUCESSO
    
    Exit Function
    
Erro_Separa_Borderos:
    
    Separa_Borderos = gErr
    
    Select Case gErr
        
        Case 62016
            Call Rotina_Erro(vbOKOnly, "ERRO_CARTEIRACOBRADOR_NAO_CADASTRADA1", gErr, objCarteiraCob.iCodCarteiraCobranca, objCarteiraCob.iCobrador)
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147403)
            
    End Select

    Exit Function
    
End Function

Function Criar_Arquivo_Remessa_BradescoAux(objCobrancaEletronica As ClassCobrancaEletronica) As Long
'Cria uma arquivo de remessa com base nos dados do BD

Dim iCancela As Integer
Dim lErro As Long
Dim sRegHeader As String
'Dim lTransacao As Long
Dim alComando(0 To 9) As Long
Dim sNomeArquivo As String
Dim iIndice As Integer
Dim sRegDetalhe As String, sRegInstrucoes As String
Dim sRegTrailer As String, lSeqNomeArq As Long
Dim objBorderoCobranca As New ClassBorderoCobranca
Dim objCarteiraCobrador As New ClassCarteiraCobrador
Dim objOcorrRemParc As ClassOcorrRemParcRec

On Error GoTo Erro_Criar_Arquivo_Remessa_BradescoAux
    
    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 51754
    Next
        
'    'ABre a transação
'    lTransacao = Transacao_Abrir()
'    If lTransacao = 0 Then gError 51755
'
    Set gobjCobrancaEletronica = objCobrancaEletronica
    Set gColFiliaisEmpresa = New Collection
    
    'Gera o nome do arquivo de acordo com o sequencial do Cobrador
    lErro = CF("CobrancaEletronica_Gera_NomeArquivo", objCobrancaEletronica.objCobrador, objCobrancaEletronica.sDiretorio, sNomeArquivo, lSeqNomeArq, objCobrancaEletronica.iCarteira)
    If lErro <> SUCESSO Then gError 51756
    
    gobjCobrancaEletronica.lSeqNomeArq = lSeqNomeArq
    
    'Cria o arquivo de remessa
    Open sNomeArquivo For Output As #1
        
    lErro = Montar_RegistroHeader_Bradesco(sRegHeader)
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
            sRegInstrucoes = ""
            'Monta um registro de detalhe do arquivo de remessa
            lErro = Montar_RegistroDetalhe_Bradesco(sRegDetalhe, objOcorrRemParc, objCarteiraCobrador, sRegInstrucoes)
            If lErro <> SUCESSO And lErro <> 32266 Then gError 51761
            
            If lErro = SUCESSO Then
            
                'Inclui o registro com os detalhes do Título(Parcela) a receber
                Print #1, DesacentuaTexto(sRegDetalhe)
                
            End If
            
            If sRegInstrucoes <> "" Then
            
                'Inclui o registro com as instrucoes para o boleto do Título(Parcela) a receber
                Print #1, DesacentuaTexto(sRegInstrucoes)
            
            End If
            
            sRegDetalhe = ""

            'Monta um registro de xmls do arquivo de remessa
            lErro = Montar_Registro4_Bradesco(sRegDetalhe, objOcorrRemParc, objCarteiraCobrador)
            If lErro = SUCESSO And sRegDetalhe <> "" Then
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
    lErro = Montar_RegistroTrailer_Bradesco(sRegTrailer, objOcorrRemParc)
    If lErro <> SUCESSO Then gError 51762

    'Coloca o registro de trailer no arquivo
    Print #1, DesacentuaTexto(sRegTrailer)
    
    Close #1
    
    'atualizar seq de arquivo de remessa no cobrador
    lErro = Cobrador_Atualizar(objCobrancaEletronica.objCobrador, alComando(6), alComando(7))
    If lErro <> SUCESSO Then gError 81699
    
    lErro = CCI_Atualizar(objCobrancaEletronica.objCCI, alComando(8), alComando(9))
    If lErro <> SUCESSO Then gError 81699
    
'    lErro = Transacao_Commit()
'    If lErro <> SUCESSO Then gError 62010
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Criar_Arquivo_Remessa_BradescoAux = SUCESSO
    
    Exit Function

Erro_Criar_Arquivo_Remessa_BradescoAux:

    Criar_Arquivo_Remessa_BradescoAux = gErr
    
    Select Case gErr
    
        Case 51754
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 51755
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 51756, 51757, 51758, 51759, 51760, 51761, 51762, 79705, 81699
        
        Case 62010
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147199)
    
    End Select
    
    Close #1
    
'    Call Transacao_Rollback

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Private Function Montar_RegistroHeader_Bradesco(sRegHeader As String) As Long
'Monta o registro

Dim tRegHeaderBradesco As typeRemHeaderBradesco
Dim lErro As Long

On Error GoTo Erro_Montar_RegistroHeader_Bradesco

    gobjCobrancaEletronica.lSequencialRegistro = 1

    'Carrega tRegHeaderBradesco com os dados do Header do aquivo
    lErro = Obter_Dados_Header(tRegHeaderBradesco)
    If lErro <> SUCESSO Then Error 51763
        
    With tRegHeaderBradesco
    
        'Monta o registro de Header do Bradesco
        sRegHeader = .sIdentificacaoReg & .sIdentificacaoArq & .sLiteralRemessa & .sCodTipoServico & .sLiteralServico
        sRegHeader = sRegHeader & .sCodigoEmpresa & .sNomeEmpresa & .sNumeroBanco & .sNomeBanco & .sDataEmissaoArq
        sRegHeader = sRegHeader & .sComplemento1 & .sIdentificacaoSistema & .sSequencialArq & .sComplemento2 & .sSequencialRegistro
        
    End With
    
    Montar_RegistroHeader_Bradesco = SUCESSO

    Exit Function

Erro_Montar_RegistroHeader_Bradesco:

    Montar_RegistroHeader_Bradesco = Err

    Select Case Err

        Case 51763

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147200)

    End Select

    Exit Function

End Function

Private Function Obter_Dados_Header(tRegHeaderBradesco As typeRemHeaderBradesco) As Long

Dim lErro As Long
Dim objCCI As New ClassContasCorrentesInternas
Dim sCodEmpresaBanco As String, lCNABProxSeqArqCobr As Long

On Error GoTo Erro_Obter_Dados_Header
        
    'Lê a conta corrente associada ao cobrador
    lErro = CF("ContaCorrenteInt_Le", gobjCobrancaEletronica.objCobrador.iCodCCI, objCCI)
    If lErro <> SUCESSO And lErro <> 11807 Then Error 51764
    If lErro <> SUCESSO Then Error 51765
    
    If objCCI.iCodBanco <> gobjCobrancaEletronica.objCobrador.iCodBanco Then Error 62061
    If Len(Trim(objCCI.sAgencia)) = 0 Or Len(Trim(objCCI.sNumConta)) = 0 Then Error 62062
        
    Set gobjCobrancaEletronica.objCCI = objCCI
    
    'Carrega os campos do cabeçalho
    tRegHeaderBradesco.sIdentificacaoReg = "0"
    tRegHeaderBradesco.sIdentificacaoArq = "1"
    tRegHeaderBradesco.sLiteralRemessa = FormataCpoTexto("REMESSA", 7)
    tRegHeaderBradesco.sCodTipoServico = "01"
    tRegHeaderBradesco.sLiteralServico = FormataCpoTexto("COBRANCA", 15)
    tRegHeaderBradesco.sCodigoEmpresa = FormataCpoNum(Obter_Info_Cobrador(gobjCobrancaEletronica.objCobradorCNABInfo.colInformacoes, 1), 20)
    
    'Pega o próximo número do arquivo da conta se tiver, senão pega do cobrador
    If gobjCobrancaEletronica.objCCI.lCNABProxSeqArqCobr <> 0 Then
        lCNABProxSeqArqCobr = gobjCobrancaEletronica.objCCI.lCNABProxSeqArqCobr
    Else
        lCNABProxSeqArqCobr = gobjCobrancaEletronica.objCobrador.lCNABProxSeqArqCobr
    End If
    
    tRegHeaderBradesco.sNomeEmpresa = FormataCpoTexto(gsNomeEmpresa, 30)
    tRegHeaderBradesco.sNumeroBanco = "237"
    tRegHeaderBradesco.sNomeBanco = FormataCpoTexto("BRADESCO", 15)
    tRegHeaderBradesco.sDataEmissaoArq = Format(gdtDataHoje, "DDMMYY")
    tRegHeaderBradesco.sComplemento1 = FormataCpoTexto("", 8)
    tRegHeaderBradesco.sIdentificacaoSistema = FormataCpoTexto("MX", 2)
    tRegHeaderBradesco.sSequencialArq = FormataCpoNum(lCNABProxSeqArqCobr, 7)
    tRegHeaderBradesco.sComplemento2 = FormataCpoTexto("", 277)
    tRegHeaderBradesco.sSequencialRegistro = FormataCpoNum(gobjCobrancaEletronica.lSequencialRegistro, 6)
    
    If gobjCobrancaEletronica.iCarteira = CARTEIRA_DESCONTADA Then
        tRegHeaderBradesco.sComplemento2 = FormataCpoTexto("", 249) & "DESC" & FormataCpoNum(Obter_Info_Cobrador(gobjCobrancaEletronica.objCobradorCNABInfo.colInformacoes, 31), 4) & FormataCpoNum(objCCI.sAgencia, 5) & FormataCpoNum(objCCI.sNumConta, 7) & FormataCpoNum(objCCI.sDVNumConta, 1) & FormataCpoTexto("", 7)
    End If
        
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147201)
            
    End Select
    
    Exit Function

End Function

Private Function Montar_RegistroDetalhe_Bradesco(sRegDetalhe As String, objOcorrRemParcRec As ClassOcorrRemParcRec, objCarteiraCobrador As ClassCarteiraCobrador, sRegInstrucoes As String) As Long

'Monta o registro

Dim tRegDetalhe As typeRemDetalheBradesco
Dim lErro As Long, sInstrucoes As String
Dim sAux As String

On Error GoTo Erro_Montar_RegistroDetalhe_Bradesco
    
    'Obtem os dados do registro de detalhe
    lErro = Obter_Dados_Detalhe(tRegDetalhe, objOcorrRemParcRec, objCarteiraCobrador, sInstrucoes)
    If lErro <> SUCESSO And lErro <> 32265 Then Error 51766
    If lErro <> SUCESSO Then Error 32266
    
    With tRegDetalhe
        
        Select Case gobjCobrancaEletronica.iCarteira
        
            Case CARTEIRA_DESCONTADA
            
                sInstrucoes = ""
                
                If Len(Trim(.sIdentOpBanco)) = 0 Then
                    sAux = FormataCpoTexto("", 16)
                Else
                    sAux = "0" & FormataCpoTexto("", 1) & .sIdentOpBanco & FormataCpoTexto("", 4)
                End If
                
                sRegDetalhe = .sIdentificacaoReg & .sIdentifEmpresa & .sNumControle & FormataCpoNum("", 8)
                sRegDetalhe = sRegDetalhe & .sNossoNumero & FormataCpoNum("", 10) & sAux & .sIdentificacaoOcorr & .sNumDocto
                sRegDetalhe = sRegDetalhe & .sDataVencto & .sValorTitulo & FormataCpoTexto("", 8) & .sEspecie & " "
                sRegDetalhe = sRegDetalhe & .sDataEmissao & FormataCpoNum("", 62) & .sIdentfInscSacado & .sNumInscricaoSacado
                sRegDetalhe = sRegDetalhe & .sNomeSacado & .sEnderecoSacado & FormataCpoTexto("", 12) & .sCEP & .sSacadorAvalista & .sNumSequencialRegistro
           
            Case Else
            
                sRegDetalhe = .sIdentificacaoReg & .sAgenciaDebito & .sDigitoAgDebito & .sRazaoContaCorrente & .sContaCorrente
                sRegDetalhe = sRegDetalhe & .sDigitoContaCorrente & .sIdentifEmpresa & .sNumControle & .sCodigoBancoDebitado & .sZeros
                sRegDetalhe = sRegDetalhe & .sNossoNumero & .sDesconto & .sCondEmissaoPapel & .sIdentEmitePapel & .sIdentOpBanco
                sRegDetalhe = sRegDetalhe & .sIndicadorRateio & .sEnderecamento & .sBranco & .sIdentificacaoOcorr & .sNumDocto
                sRegDetalhe = sRegDetalhe & .sDataVencto & .sValorTitulo & .sBancoCobranca & .sAgenciaDepositaria & .sEspecie
                sRegDetalhe = sRegDetalhe & .sAceite & .sDataEmissao & .sInstrucao1 & .sInstrucao2 & .sValorJuros & .sDataDesconto
                sRegDetalhe = sRegDetalhe & .sValorDesconto & .sValorIOF & .sAbatimento & .sIdentfInscSacado & .sNumInscricaoSacado
                sRegDetalhe = sRegDetalhe & .sNomeSacado & .sEnderecoSacado & .sMensagem1 & .sCEP & .sSacadorAvalista & .sNumSequencialRegistro
            
        End Select
        
    End With
    
    sInstrucoes = Trim(sInstrucoes)
    If sInstrucoes <> "" Then
        sRegInstrucoes = "2" & FormataCpoTexto(left(sInstrucoes, 80), 80) & FormataCpoTexto(Mid(sInstrucoes, 80 + 1, 80), 80) & FormataCpoTexto(Mid(sInstrucoes, (2 * 80) + 1, 80), 80) & FormataCpoTexto(Mid(sInstrucoes, (3 * 80) + 1, 80), 80)
        sRegInstrucoes = sRegInstrucoes & FormataCpoNum(0, 6 + 13 + 6 + 13) & FormataCpoTexto("", 45 - (6 + 13 + 6 + 13)) & Mid(tRegDetalhe.sIdentifEmpresa, 2) & tRegDetalhe.sNossoNumero & FormataCpoNum(gobjCobrancaEletronica.lSequencialRegistro, 6)
        gobjCobrancaEletronica.lSequencialRegistro = gobjCobrancaEletronica.lSequencialRegistro + 1
    End If
    
    Montar_RegistroDetalhe_Bradesco = SUCESSO

    Exit Function

Erro_Montar_RegistroDetalhe_Bradesco:

    Montar_RegistroDetalhe_Bradesco = Err

    Select Case Err
    
        Case 51766, 32266

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147202)

    End Select

    Exit Function

End Function

Private Function Obter_Dados_Detalhe(tRegDetalhe As typeRemDetalheBradesco, objOcorrRemParcRem As ClassOcorrRemParcRec, objCarteiraCobrador As ClassCarteiraCobrador, sInstrucoes As String) As Long
'Obtem os dados que são utilizados para preencher o registro de detalhe

Dim lErro As Long
Dim objParcelaRec As New ClassParcelaReceber
Dim objEndereco As New ClassEndereco
Dim sNossoNumero As String, objTituloReceber As New ClassTituloReceber
Dim dValorDesconto As Double, sEndAux As String, sDVNN As String, iCarteira As Integer
Dim bCalcNossoNumero As Boolean, sAux As String, iGrupoSifra As Integer, iCNAB444 As Integer
Dim objNF As New ClassNFiscal, sTela As String

On Error GoTo Erro_Obter_Dados_Detalhe

    objParcelaRec.lNumIntDoc = objOcorrRemParcRem.lNumIntParc

    'Busca os dados da parcela a receber no BD
    lErro = CF("CobrancaEletronica_Obter_DadosParcelaRec", objParcelaRec, gobjCobrancaEletronica)
    If lErro <> SUCESSO And lErro <> 51674 Then gError 51767
    If lErro <> SUCESSO Then gError 32265
    
    lErro = CF("CobrancaEletronica_Obter_Endereco", gobjCobrancaEletronica, objEndereco)
    If lErro <> SUCESSO Then gError 51769
        
    sNossoNumero = objParcelaRec.sNumTitCobrador
    bCalcNossoNumero = False
    'Se é para imprimir o boleto e força para ser o número do título ou não tem o nosso número na parcela, senão respeita o nosso número da parcela
    If objCarteiraCobrador.iImprimeBoleta = EMPRESA_IMPRIME_BOLETA And (Len(Trim(sNossoNumero)) = 0 Or (right(objCarteiraCobrador.sFaixaNossoNumeroProx, 10) = "9999999999" And objCarteiraCobrador.iNumCarteiraNoBanco = 9)) Then
        If Len(Trim(objCarteiraCobrador.sFaixaNossoNumeroProx)) = 0 Then objCarteiraCobrador.sFaixaNossoNumeroProx = FormataCpoNum(1, 11)
        lErro = Calcula_NossoNumero_Bradesco(objCarteiraCobrador, sNossoNumero)
        If lErro <> SUCESSO Then gError 51770
        bCalcNossoNumero = True
    End If
    
    iCarteira = objCarteiraCobrador.iNumCarteiraNoBanco
    
    If objCarteiraCobrador.iComRegistro = CARTEIRA_SEM_REGISTRO Then iCarteira = "06"
    
    With tRegDetalhe
        .sIdentificacaoReg = "1"
        .sAgenciaDebito = FormataCpoNum("", 5)
        .sDigitoAgDebito = FormataCpoTexto("0", 1)
        .sRazaoContaCorrente = FormataCpoNum("", 5)
        .sContaCorrente = FormataCpoNum("", 7)
        .sDigitoContaCorrente = FormataCpoTexto("0", 1)
        .sIdentifEmpresa = "0" & FormataCpoNum(objCarteiraCobrador.iNumCarteiraNoBanco, 3) & FormataCpoNum(gobjCobrancaEletronica.objCCI.sAgencia, 5) & FormataCpoNum(gobjCobrancaEletronica.objCCI.sNumConta, 7) & FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVNumConta, 1)
        
        iGrupoSifra = StrParaInt(Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 35))
    
        If iGrupoSifra = MARCADO Then
            .sNumControle = FormataCpoTexto(objEndereco.sTelefone1, 25)
        Else
            .sNumControle = FormataCpoTexto(objOcorrRemParcRem.lNumIntDoc, 25)
        End If
        
        .sCodigoBancoDebitado = FormataCpoNum("", 3)
        If gobjCobrancaEletronica.dPercMulta = 0 Then
            .sZeros = FormataCpoNum("", 5)
        Else
            .sZeros = "2" & FormataCpoValor(gobjCobrancaEletronica.dPercMulta * 100, 4)
        End If
        
        If bCalcNossoNumero Then
            .sNossoNumero = FormataCpoNum(sNossoNumero, 12)
        Else
            'Nosso número veio da Parcela
            Call Calcula_DV_NossoNumero_Bradesco(FormataCpoNum(iCarteira, 2) & sNossoNumero, sDVNN, 11, 7)
            
            .sNossoNumero = FormataCpoNum(sNossoNumero & sDVNN, 12)
        End If
        .sDesconto = FormataCpoNum("", 10)
        If objCarteiraCobrador.iImprimeBoleta = BANCO_IMPRIME_BOLETA And objParcelaRec.lIdImpressaoBoleto = 0 Then
            .sCondEmissaoPapel = "1"
        Else
            .sCondEmissaoPapel = "2"
        End If
        .sIdentEmitePapel = FormataCpoTexto("S", 1)
        
        If iGrupoSifra = MARCADO Then
        
            lErro = CF("Titulo_Le_DocumentoOriginal", objParcelaRec.lNumIntTitulo, CPR_TITULO_RECEBER, objNF, sTela)
            If lErro <> SUCESSO And lErro <> 58942 Then gError ERRO_SEM_MENSAGEM
        
            If lErro = SUCESSO Then
            
                lErro = CF("NFiscal_Le", objNF)
                If lErro <> SUCESSO And lErro <> 31442 Then gError ERRO_SEM_MENSAGEM
        
                .sIdentOpBanco = FormataCpoNum(objNF.lNumNotaFiscal, 10)
            Else
                .sIdentOpBanco = FormataCpoTexto("", 10)
            End If
        Else
            .sIdentOpBanco = FormataCpoTexto("", 10)
        End If
        
        .sIndicadorRateio = FormataCpoTexto("", 1)
        
        sEndAux = Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 2)
        If Len(Trim(sEndAux)) > 0 Then
            .sEnderecamento = left(sEndAux, 1)
        Else
            .sEnderecamento = "2"
        End If
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
        
        If objOcorrRemParcRem.iInstrucao1 <> 0 Then
            .sInstrucao1 = FormataCpoNum(objOcorrRemParcRem.iInstrucao1, 2)
        Else
            .sInstrucao1 = FormataCpoNum(Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 3), 2)
        End If
        
        If objOcorrRemParcRem.iInstrucao2 <> 0 Then
            .sInstrucao2 = FormataCpoNum(objOcorrRemParcRem.iInstrucao2, 2)
        Else
            .sInstrucao2 = FormataCpoNum(Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 4), 2)
        End If
        
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
        .sValorIOF = FormataCpoValor(0, 13)
        .sAbatimento = FormataCpoValor(0, 13)
        If Len(gobjCobrancaEletronica.sCGCSacado) = 14 Then
            .sIdentfInscSacado = "02"
        Else
            .sIdentfInscSacado = "01"
        End If
        .sNumInscricaoSacado = FormataCpoNum(gobjCobrancaEletronica.sCGCSacado, 14)
        
        If .sNumInscricaoSacado = FormataCpoNum(0, 14) Then
            .sIdentfInscSacado = "98"
        End If
        
        .sNomeSacado = FormataCpoTexto(gobjCobrancaEletronica.sNomeSacado, 40)
        .sEnderecoSacado = FormataCpoTexto(Trim(objEndereco.sEndereco), 40)
        .sMensagem1 = FormataCpoTexto(tRegDetalhe.sMensagem1, 12)
        .sCEP = Formata_Cep(objEndereco.sCEP, 8)
        
        If iGrupoSifra = MARCADO Then
            .sSacadorAvalista = FormataCpoTexto(left(objEndereco.sCidade, 15), 15) & FormataCpoTexto(left(objEndereco.sSiglaEstado, 2), 2) & FormataCpoTexto(objOcorrRemParcRem.lNumIntDoc, 25) & FormataCpoTexto("", 18)
        Else
            .sSacadorAvalista = FormataCpoTexto("", 60)
        End If
        
        .sNumSequencialRegistro = FormataCpoNum(gobjCobrancaEletronica.lSequencialRegistro, 6)
    
        iCNAB444 = StrParaInt(Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 36))
        
        If iCNAB444 = MARCADO Then
       
            lErro = CF("Titulo_Le_DocumentoOriginal", objParcelaRec.lNumIntTitulo, CPR_TITULO_RECEBER, objNF, sTela)
            If lErro <> SUCESSO And lErro <> 58942 Then gError ERRO_SEM_MENSAGEM
            
            lErro = CF("NFiscal_Le_ChvNFe", objNF)
            If lErro <> SUCESSO And lErro <> 58942 Then gError ERRO_SEM_MENSAGEM
            
            .sNumSequencialRegistro = .sNumSequencialRegistro & FormataCpoNum(objNF.sChvNFe, 44)
        
        End If
    
        If gobjCobrancaEletronica.iCarteira = CARTEIRA_DESCONTADA Then
            sAux = ""
            If Len(gobjCobrancaEletronica.sCGCCedente) = 14 Then
                sAux = "02"
            Else
                sAux = "01"
            End If
            sAux = sAux & FormataCpoNum(gobjCobrancaEletronica.sCGCCedente, 14)
    
            .sIdentifEmpresa = sAux & "0000" & FormataCpoNum(objCarteiraCobrador.iNumCarteiraNoBanco, 3) & FormataCpoNum(gobjCobrancaEletronica.objCCI.sAgencia, 5) & FormataCpoNum(gobjCobrancaEletronica.objCCI.sNumConta, 7) & FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVNumConta, 1)
        
        End If
        
    End With

    'Incrementa o sequencial de registro
    gobjCobrancaEletronica.lSequencialRegistro = gobjCobrancaEletronica.lSequencialRegistro + 1
    
    'obter instruçoes para o boleto
    objTituloReceber.lNumIntDoc = objParcelaRec.lNumIntTitulo
    lErro = CF("TituloReceber_Le", objTituloReceber)
    If lErro <> SUCESSO And lErro <> 26061 Then gError 184423
    If lErro <> SUCESSO Then
        lErro = CF("TituloReceberBaixado_Le", objTituloReceber)
        If lErro <> SUCESSO And lErro <> 56570 Then gError 184424
        If lErro <> SUCESSO Then gError 184426
    End If
    lErro = CF("Preenche_Instrucoes_Boleto", objTituloReceber, objParcelaRec, "R$", sInstrucoes)
    If lErro <> SUCESSO Then gError 184425
    
    Obter_Dados_Detalhe = SUCESSO

    Exit Function
    
Erro_Obter_Dados_Detalhe:

    Obter_Dados_Detalhe = gErr
    
    Select Case gErr

        Case 51767, 51769, 51770, 32265, 184423, 184424
       
        Case 184425, 184426
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITREC_INSTRCNAB", gErr)
        
        Case 51768
            Call Rotina_Erro(vbOKOnly, "ERRO_ENDERECO_FILIALCLIENTE_NAO_INFORMADO", gErr, gobjCobrancaEletronica.lCliente, gobjCobrancaEletronica.iFilialCli)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147203)

    End Select

    Exit Function

End Function

Private Function Montar_RegistroTrailer_Bradesco(sRegTrailer As String, objCobrador As ClassCobrador) As Long
'Monta o registro

Dim tRegTrailerBradesco As typeRemTrailerBradesco
Dim lErro As Long

On Error GoTo Erro_Montar_RegistroTrailer_Bradesco

    'Preenche os dados do regsitro de trailer
    tRegTrailerBradesco.sIdentificacaoReg = "9"
    tRegTrailerBradesco.sComplemento1 = FormataCpoTexto("", 393)
    tRegTrailerBradesco.sSequencialRegistro = FormataCpoNum(gobjCobrancaEletronica.lSequencialRegistro, 6)
    
    'Monta o Registro de trailer
    sRegTrailer = tRegTrailerBradesco.sIdentificacaoReg & tRegTrailerBradesco.sComplemento1 & tRegTrailerBradesco.sSequencialRegistro

    Montar_RegistroTrailer_Bradesco = SUCESSO

    Exit Function

Erro_Montar_RegistroTrailer_Bradesco:

    Montar_RegistroTrailer_Bradesco = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147204)

    End Select

    Exit Function

End Function

Function Processar_ArquivoRetorno_Cobranca_Bradesco(objCobrancaEletronica As ClassCobrancaEletronica) As Long
'Faz o processamento do arquivo de retorno

Dim lErro As Long
Dim iIndice As Integer
Dim sRegistro As String
Dim iCodRegistro As Integer
Dim objBorderoRetCobr As New ClassBorderoRetCobr
Dim lComando As Long
Dim lComando1 As Long

On Error GoTo Erro_Processar_ArquivoRetorno_Cobranca_Bradesco

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
        iCodRegistro = StrParaInt(left(sRegistro, 1))
        
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

    Processar_ArquivoRetorno_Cobranca_Bradesco = SUCESSO
    
    Exit Function
    
Erro_Processar_ArquivoRetorno_Cobranca_Bradesco:

    Processar_ArquivoRetorno_Cobranca_Bradesco = Err
    
    Select Case Err
    
        Case 51771
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 51772, 51773
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 51774, 51775, 51776, 51777, 32265
        
        Case 51779
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORMATO_ARQUIVO_INCORRETO", Err, objCobrancaEletronica.sNomeArquivoRetorno)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147205)
    
    End Select

    Close #2

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Private Function Processa_Registro_Header(sRegistro As String, objBorderoRetCobr As ClassBorderoRetCobr, objCobrador As ClassCobrador, sNomeArquivo As String) As Long

Dim lErro As Long
Dim tRegHeaderBradesco As typeRetHeaderBradesco
Dim objContaCorrente As New ClassContasCorrentesInternas

On Error GoTo Erro_Processa_Registro_Header

    'Extrai da variável lida no arquivo os dados do Header
    lErro = Extrair_Header_Retorno(tRegHeaderBradesco, sRegistro)
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
    objBorderoRetCobr.objMovCCI.dtDataMovimento = Extrair_Data(tRegHeaderBradesco.sDataCredito)
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147206)
            
    End Select
    
    Exit Function

End Function

Private Function Extrair_Header_Retorno(tRegHeaderBradesco As typeRetHeaderBradesco, sRegistro As String) As Long
'Extrai do sRegistro os dados do header do arquido de retorno

Dim lErro As Long

On Error GoTo Erro_Extrair_Header_Retorno

    With tRegHeaderBradesco
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147207)
            
    End Select
    
    Exit Function

End Function


Private Function Processa_Registro_Detalhe(sRegistro As String, objBorderoRetCobr As ClassBorderoRetCobr, objCobrancaEletronica As ClassCobrancaEletronica, lComando As Long, lComando1 As Long) As Long

Dim tRegDetalhe As typeRetDetalheBradesco
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
    
'    Select Case iCodOcorrencia
'
'        'Quando ocorre a baixa
'        Case 6, 9, 10, 17
'            iTipoOcorrencia = RETCOBR_DET_BAIXA
'
'        Case 3
'            iTipoOcorrencia = RETCOBR_DET_REJEITADO
'
'        Case Else
'            iTipoOcorrencia = RETCOBR_DET_IGNORAR
'
'    End Select

    Select Case iCodOcorrencia
        
        'Quando ocorre baixa p/liquidacao
        Case 6, 9, 15, 17
            iTipoOcorrencia = RETCOBR_DET_BAIXA
            
        'cobranca de tarifas e custas
        Case 2, 12, 19, 20, 23, 28, 33
            iTipoOcorrencia = RETCOBR_DET_TARIFAS
                                
        Case 3, 24, 32
            iTipoOcorrencia = RETCOBR_DET_REJEITADO
    
        Case 25 'baixa por ter sido protestado
            iTipoOcorrencia = RETCOBR_DET_BAIXA_POR_PROTESTO
        
        Case Else
            iTipoOcorrencia = RETCOBR_DET_IGNORAR
            
    End Select
    
    If iTipoOcorrencia <> RETCOBR_DET_IGNORAR Then
    
        lErro = Prepara_CartCobrRet1(objCobrancaEletronica, objBorderoRetCobr, objBorderoRetCobrCart, objDetRetCobr, tRegDetalhe)
        If lErro <> SUCESSO Then Error 32257
        
        If Len(Trim(objDetRetCobr.sSeuNumero)) = 0 Then
            objDetRetCobr.sSeuNumero = objDetRetCobr.sNossoNumero & SEPARADOR & "1"
            lErro = CF("Processa_Registro_Detalhe_Geral", objBorderoRetCobr, objCobrancaEletronica, iTipoOcorrencia, RETCOBR_TIPO_SEU_NUMERO1, objDetRetCobr, objBorderoRetCobrCart, lComando)
        Else
            lErro = CF("Processa_Registro_Detalhe_Geral", objBorderoRetCobr, objCobrancaEletronica, iTipoOcorrencia, RETCOBR_TIPO_SEU_NUMERO4, objDetRetCobr, objBorderoRetCobrCart, lComando)
        End If
        If lErro <> SUCESSO Then Error 32356
                                  
    End If
    
    Processa_Registro_Detalhe = SUCESSO
    
    Exit Function

Erro_Processa_Registro_Detalhe:

    Processa_Registro_Detalhe = Err

    Select Case Err
    
        Case 51783, 51784, 51786, 51788
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147208)
            
    End Select
    
    Exit Function

End Function

Private Function Extrair_Detalhe_Retorno(tRegDetalhe As typeRetDetalheBradesco, sRegistro As String) As Long
'Extrai do sRegistro os dados do Detalhe do arquido de retorno

Dim lErro As Long

On Error GoTo Erro_Extrair_Detalhe_Retorno

    With tRegDetalhe
        
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147209)
            
    End Select
    
    Exit Function

End Function
Function Processa_Registro_Trailer(sRegistroTrailer As String) As Long

Dim lErro As Long
Dim tRegTrailerBradesco As typeRetTrailerBradesco

On Error GoTo Erro_Processa_Registro_Trailer

    'Extrai da variável lida no arquivo os dados do registro de detalhe
    lErro = Extrair_Trailer_Retorno(tRegTrailerBradesco, sRegistroTrailer)
    If lErro <> SUCESSO Then Error 51789
    
    Processa_Registro_Trailer = SUCESSO
    
    Exit Function
    
Erro_Processa_Registro_Trailer:

    Processa_Registro_Trailer = Err
    
    Select Case Err
    
        Case 51789
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147210)
            
    End Select
    
    Exit Function
    
End Function
Private Function Extrair_Trailer_Retorno(tRegTrailerBradesco As typeRetTrailerBradesco, sRegistro As String) As Long
'Extrai do sRegistro os dados do Trailer do arquido de retorno

Dim lErro As Long

On Error GoTo Erro_Extrair_Trailer_Retorno

    tRegTrailerBradesco.sIdentificacaoReg = Mid(sRegistro, 1, 1)
    tRegTrailerBradesco.sCodigoRetorno = Mid(sRegistro, 2, 1)
    tRegTrailerBradesco.sIdentificacaoReg = Mid(sRegistro, 3, 2)
    tRegTrailerBradesco.sNumeroBanco = Mid(sRegistro, 5, 3)
    tRegTrailerBradesco.sComplemento1 = Mid(sRegistro, 8, 10)
    tRegTrailerBradesco.sQuantTitulosCobr = Mid(sRegistro, 18, 8)
    tRegTrailerBradesco.sValorTotalCobr = Mid(sRegistro, 26, 14)
    tRegTrailerBradesco.sAvisoBancario = Mid(sRegistro, 40, 8)
    tRegTrailerBradesco.sComplemento2 = Mid(sRegistro, 48, 10)
    tRegTrailerBradesco.sQuantRegOcorr_02 = Mid(sRegistro, 58, 5)
    tRegTrailerBradesco.sValorRegOcorr_02 = Mid(sRegistro, 63, 12)
    tRegTrailerBradesco.sValorRegOcorr_06 = Mid(sRegistro, 75, 12)
    tRegTrailerBradesco.sQuantRegOcorr_06 = Mid(sRegistro, 87, 5)
    tRegTrailerBradesco.sValorRegOcorr_06_09_10 = Mid(sRegistro, 92, 12)
    tRegTrailerBradesco.sQuantRegOcorr_09_10 = Mid(sRegistro, 104, 5)
    tRegTrailerBradesco.sValorRegOcorr_09_10 = Mid(sRegistro, 109, 12)
    tRegTrailerBradesco.sQuantRegOcorr_13 = Mid(sRegistro, 121, 5)
    tRegTrailerBradesco.sValorRegOcorr_13 = Mid(sRegistro, 126, 12)
    tRegTrailerBradesco.sQuantRegOcorr_14 = Mid(sRegistro, 138, 5)
    tRegTrailerBradesco.sValorRegOcorr_14 = Mid(sRegistro, 143, 12)
    tRegTrailerBradesco.sQuantRegOcorr_12 = Mid(sRegistro, 155, 5)
    tRegTrailerBradesco.sValorRegOcorr_12 = Mid(sRegistro, 160, 12)
    tRegTrailerBradesco.sQuantRegOcorr_19 = Mid(sRegistro, 172, 5)
    tRegTrailerBradesco.sValorRegOcorr_19 = Mid(sRegistro, 177, 12)
    tRegTrailerBradesco.sComplemento3 = Mid(sRegistro, 189, 174)
    tRegTrailerBradesco.sValorTotalRateios = Mid(sRegistro, 363, 15)
    tRegTrailerBradesco.sQuantTotalRateios = Mid(sRegistro, 378, 8)
    tRegTrailerBradesco.sSequencialRegistro = Mid(sRegistro, 395, 6)

    Extrair_Trailer_Retorno = SUCESSO

    Exit Function

Erro_Extrair_Trailer_Retorno:

    Extrair_Trailer_Retorno = Err
    
    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147211)
            
    End Select
    
    Exit Function

End Function

Function Extrair_Valor(vValor As Variant) As Double
'Extrai o conteúdo numérico (double) do valor passado
Dim sValor As String

    If Len(Trim(vValor)) = 0 Then
        Extrair_Valor = 0
    Else
        
        sValor = left(vValor, Len(vValor) - 2)
        Extrair_Valor = StrParaDbl(sValor & "," & right(vValor, 2))
        
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

Function BorderoPagto_Criar_ArquivoCNAB_Bradesco(objGeracaoArqCNABPag As ClassGeracaoArqCNABPag) As Long
'Cria o arquivo de remessa de títulos a pagar nos padrões do Bradesco

Dim lErro As Long
Dim sNomeArquivo As String
Dim lTransacao As Long
Dim iIndice As Integer
Dim sRegHeaderArq As String
Dim sRegDetalhe As String
Dim sTrailerArq As String
Dim objBaixaParcPag As New ClassBaixaParcPagar
Dim alComando(0 To 2) As Long

On Error GoTo Erro_BorderoPagto_Criar_ArquivoCNAB_Bradesco

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
    lErro = CF("BorderoPagto_Lock", alComando(0), gobjGeracaoArqCNABPag.objBorderoPagto.lNumIntBordero)
    If lErro <> SUCESSO Then Error 51792

    'Gera o nome do arquivo
    lErro = CF("PagamentoEletronico_Gera_NomeArquivo", sNomeArquivo, gobjGeracaoArqCNABPag)
    If lErro <> SUCESSO Then Error 51793

    'Cria o arquivo de remessa
    Open sNomeArquivo For Output As #1
 
    gobjGeracaoArqCNABPag.lQuantRegistrosArq = 0
    
    'Monta o registro de Header do arquivo
    lErro = Montar_RegistroHeaderCNABPag_Bradesco(sRegHeaderArq)
    If lErro <> SUCESSO Then Error 51794

    
    'Inclui no arquivo
    Print #1, DesacentuaTexto(sRegHeaderArq)
    
    gobjGeracaoArqCNABPag.lSequencialRegDetalhe = 1
    gobjGeracaoArqCNABPag.dValorTotalLote = 0
    
    'Lê as baixas ligadas ao bordero
    lErro = CF("BaixasParcPag_Le_BorderoPagto", alComando(2), gobjGeracaoArqCNABPag.objBorderoPagto.lNumIntBordero, gobjGeracaoArqCNABPag.colBaixasParcPag)
    If lErro <> SUCESSO Then Error 51795
    
    'Para cada baixa executada
    For Each objBaixaParcPag In gobjGeracaoArqCNABPag.colBaixasParcPag
        
        lErro = CF("BaixaParcPag_ObterDados_ParcelaTituloPag", alComando(2), objBaixaParcPag.lNumIntParcela, gobjGeracaoArqCNABPag)
        If lErro <> SUCESSO Then Error 51796
        
        'Monta um registro de
        lErro = Montar_Registro_DetalhePagto(sRegDetalhe, objBaixaParcPag)
        If lErro <> SUCESSO Then Error 51797
            
        'incluir no arq
        Print #1, DesacentuaTexto(sRegDetalhe)
    
        Call objGeracaoArqCNABPag.objTelaAtualizacao.Mostra_Evolucao(False, 1)
    
        gobjGeracaoArqCNABPag.lSequencialRegDetalhe = gobjGeracaoArqCNABPag.lSequencialRegDetalhe + 1
        gobjGeracaoArqCNABPag.dValorTotalLote = gobjGeracaoArqCNABPag.dValorTotalLote + objBaixaParcPag.dValorBaixado - objBaixaParcPag.dValorDesconto + objBaixaParcPag.dValorJuros
    
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
    
    Set gobjGeracaoArqCNABPag = Nothing
    Set gColFiliaisEmpresa = Nothing
    
    MsgBox ("Envie para o banco o arquivo " & sNomeArquivo)
    
    BorderoPagto_Criar_ArquivoCNAB_Bradesco = SUCESSO
    
    Exit Function
    
Erro_BorderoPagto_Criar_ArquivoCNAB_Bradesco:

    BorderoPagto_Criar_ArquivoCNAB_Bradesco = Err
    
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147212)
            
    End Select
    
    Close #1
    
    'Faz um RollBack na transacao
    Call Transacao_Rollback
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Set gobjGeracaoArqCNABPag = Nothing
    Set gColFiliaisEmpresa = Nothing
    
    Exit Function
    
End Function


Function Montar_RegistroHeaderCNABPag_Bradesco(sRegHeader As String) As Long
'Monta o registro de header do arquivo de contas a pagar

Dim lErro As Long
Dim tRegHeaderArq As typePagtoHeaderArqBradesco
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_Montar_RegistroHeaderCNABPag_Bradesco
  
    
    'incrementa o contador do número de registros do arquivo
    gobjGeracaoArqCNABPag.lQuantRegistrosArq = gobjGeracaoArqCNABPag.lQuantRegistrosArq + 1

    'Busca os dados da nossa filial empresa
    lErro = Obter_FilialEmpresa(giFilialEmpresa, objFilialEmpresa)
    If lErro <> SUCESSO Then Error 51801
    
    'Recolhe os dados do cabeçalho
    tRegHeaderArq.sTipoRegistro = "0"
    tRegHeaderArq.sCodComunicacao = FormataCpoNum(gobjGeracaoArqCNABPag.sCodEmpresaBanco, 8)
    If Len(objFilialEmpresa.sCgc) = 14 Then
        tRegHeaderArq.sTipoInscEmpresa = "2"
    Else
        tRegHeaderArq.sTipoInscEmpresa = "1"
    End If
    tRegHeaderArq.sCGCEmpresaDebitada = FormataCpoNum(objFilialEmpresa.sCgc, 15)
    tRegHeaderArq.sNomeEmpresa = FormataCpoTexto(gsNomeEmpresa, 40)
    tRegHeaderArq.sTipoServico = "20"
    tRegHeaderArq.sCodigoOrigem = "1"
    tRegHeaderArq.sNumRemessa = FormataCpoNum(gobjGeracaoArqCNABPag.lSequencialArq, 5)
    tRegHeaderArq.sNumRetorno = FormataCpoNum("", 5)
    tRegHeaderArq.sDataGeracao = Format(gdtDataHoje, "DDMMYYYY")
    tRegHeaderArq.sHoraGeracao = Format(Time, "HHMMSS")
    tRegHeaderArq.sDensidade = FormataCpoTexto("", 5)
    tRegHeaderArq.sUnidadeDensidade = FormataCpoTexto("", 3)
    tRegHeaderArq.sIdentifModulo = FormataCpoTexto("", 5)
    tRegHeaderArq.sTipoProcessamento = FormataCpoNum("", 1)
    tRegHeaderArq.sReservadoEmpresa = FormataCpoTexto("", 74)
    tRegHeaderArq.sReservadoBanco = FormataCpoTexto("", 80)
    tRegHeaderArq.sReservadoExpansao = FormataCpoTexto("", 234)
    tRegHeaderArq.sNumSequencialRegistro = FormataCpoNum(gobjGeracaoArqCNABPag.lQuantRegistrosArq, 6)
    
    'Forma o registro com os dados recolhidos p\ o cabeçalho
    With tRegHeaderArq
        sRegHeader = .sTipoRegistro & .sCodComunicacao & .sTipoInscEmpresa & .sCGCEmpresaDebitada & .sNomeEmpresa & .sTipoServico
        sRegHeader = sRegHeader & .sCodigoOrigem & .sNumRemessa & .sNumRetorno & .sDataGeracao & .sHoraGeracao & .sDensidade
        sRegHeader = sRegHeader & .sUnidadeDensidade & .sIdentifModulo & .sTipoProcessamento & .sReservadoEmpresa & .sReservadoBanco
        sRegHeader = sRegHeader & .sReservadoExpansao & .sNumSequencialRegistro
    End With
    
    Montar_RegistroHeaderCNABPag_Bradesco = SUCESSO
    
    Exit Function
    
Erro_Montar_RegistroHeaderCNABPag_Bradesco:

    Montar_RegistroHeaderCNABPag_Bradesco = Err
    
    Select Case Err
    
        Case 51801
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147213)
            
    End Select
    
    Exit Function
    
End Function

Function Montar_Registro_DetalhePagto(sRegDetalhe As String, objBaixaParcPag As ClassBaixaParcPagar) As Long
'Monta o registro de transação do arquivo com os dados obtidos no BD

Dim lErro As Long
Dim tRegDetalhe As typePagtoDetalheBradesco
Dim objEndereco As New ClassEndereco
Dim sAgencia As String, sNossoNumeroCodBarras As String
Dim sConta As String, sFatorVencto As String
Dim sDV As String, sCodigoDeBarras As String
Dim sAgenciaForn As String, sDVAgenciaForn As String, sContaCorrenteForn As String, sDVContaCorrenteForn As String

On Error GoTo Erro_Montar_Registro_DetalhePagto

    gobjGeracaoArqCNABPag.lQuantRegistrosArq = gobjGeracaoArqCNABPag.lQuantRegistrosArq + 1
    
    objEndereco.lCodigo = gobjGeracaoArqCNABPag.lEnderecoFavorecido
    'Lê o endereço do favorecido
    lErro = CF("Endereco_Le", objEndereco)
    If lErro <> SUCESSO And lErro <> 12309 Then Error 51802
    If lErro <> SUCESSO Then Error 51803 'Não encontrou

    'Recolhe os dados do registro de detalhe
    With tRegDetalhe
        .sIdentificacao = "1"
        If Len(gobjGeracaoArqCNABPag.sCGCFavorecido) = 14 Then
            .sTipoInscrForn = "2"
        Else
            .sTipoInscrForn = "1"
        End If
        .sInscricaoForn = FormataCpoNum(gobjGeracaoArqCNABPag.sCGCFavorecido, 15)
        .sNomeFornecedor = FormataCpoTexto(gobjGeracaoArqCNABPag.sNomeFavorecido, 30)
        .sEnderecoForn = FormataCpoTexto(objEndereco.sEndereco, 40)
        .sCepForn = Formata_Cep(objEndereco.sCEP, 8)
        .sCepComplementoForn = ""
        
        .sDataVencimento = Format(gobjGeracaoArqCNABPag.dtDataVenctoParcela, "YYYYMMDD")
        .sDataEmissao = Format(gobjGeracaoArqCNABPag.dtDataEmissaoTitulo, "YYYYMMDD")
        .sDataDesconto = FormataCpoNum("", 8)
        .sValorDocto = FormataCpoValor(objBaixaParcPag.dValorBaixado, 10)
        .sValorPagto = FormataCpoValor(objBaixaParcPag.dValorBaixado - objBaixaParcPag.dValorDesconto + objBaixaParcPag.dValorJuros, 15)
        .sValorDesconto = FormataCpoValor(objBaixaParcPag.dValorDesconto, 15)
        .sValorAcrescimo = FormataCpoValor(objBaixaParcPag.dValorJuros, 15)
        If gobjGeracaoArqCNABPag.sSiglaTitulo = TIPODOC_NF_FATURA_PAGAR Then
            .sTipoDocumento = "01"
        ElseIf gobjGeracaoArqCNABPag.sSiglaTitulo = TIPODOC_FATURA_A_PAGAR Then
            .sTipoDocumento = "02"
        Else
            .sTipoDocumento = "05"
        End If
        .sNumeroTitulo = FormataCpoNum(gobjGeracaoArqCNABPag.lNumTitulo, 10)
        .sSerieDocumento = FormataCpoTexto("  ", 2)
        
        Call Extrair_Agencia_DV(gobjGeracaoArqCNABPag.sAgenciaFavorecido, sAgencia, 5, sDV, 1)
        .sAgenciaForn = FormataCpoNum(sAgencia, 5)
        .sDVAgenciaForn = FormataCpoTexto(sDV, 1)
        Call Extrair_Conta_DV(gobjGeracaoArqCNABPag.sContaFavorecido, sConta, 13, sDV, 1)
        .sContaCorrenteForn = FormataCpoNum(sConta, 13)
        .sDVContaCorrenteForn = FormataCpoTexto(sDV, 2)
        .sContaComplementar = FormataCpoNum(gobjGeracaoArqCNABPag.objCCI.sNumConta, 7)
            
        .sSeuNumero = FormataCpoNum(0, 15)
        
        'Seleciona a modalidade de pagamento
        Select Case gobjGeracaoArqCNABPag.iTipoCobranca
        
            Case TIPO_COBRANCA_DEP_CONTA
                .sModalidadePagto = "01"
                .sBancoForn = "237"
                .sInformacoesCompl = FormataCpoTexto("", 40)
                sFatorVencto = FormataCpoNum("", 4)
                .sCarteira = FormataCpoNum(0, 3)
                .sAnoNossoNumero = FormataCpoNum(0, 3)
                .sNossoNumero = FormataCpoNum(0, 9)
                                
            Case TIPO_COBRANCA_OP
                .sModalidadePagto = "02"
                .sBancoForn = "237"
                .sInformacoesCompl = FormataCpoTexto("", 40)
                sFatorVencto = FormataCpoNum("", 4)
                .sCarteira = FormataCpoNum(0, 3)
                .sAnoNossoNumero = FormataCpoNum(0, 3)
                .sNossoNumero = FormataCpoNum(0, 9)
                
            Case TIPO_COBRANCA_DOC
                .sModalidadePagto = "03"
                .sBancoForn = FormataCpoNum(gobjGeracaoArqCNABPag.iBancoFavorecido, 3)
                .sInformacoesCompl = FormataCpoTexto("C0000000101", 40)
                sFatorVencto = FormataCpoNum("", 4)
                .sCarteira = FormataCpoNum(0, 3)
                .sAnoNossoNumero = FormataCpoNum(0, 3)
                .sNossoNumero = FormataCpoNum(0, 9)
                
            Case TIPO_COBRANCA_BANCARIA
'                If gobjGeracaoArqCNABPag.iBancoFavorecido = gobjGeracaoArqCNABPag.iBancoCobrador Then
'                    .sModalidadePagto = "30"
'                    .sBancoForn = FormataCpoNum(gobjGeracaoArqCNABPag.iBancoFavorecido, 3)
'                Else
                    .sModalidadePagto = "31"
'                    .sBancoForn = FormataCpoNum(gobjGeracaoArqCNABPag.iBancoFavorecido, 3)
'                End If
                        
                sCodigoDeBarras = gobjGeracaoArqCNABPag.sCodigoDeBarras
                .sBancoForn = FormataCpoNum(left(sCodigoDeBarras, 3), 3)
                
                'se o codigo de barras entrou por leitura optica
                If Len(Trim(sCodigoDeBarras)) = 44 Then
                    'campo livre do cod de barras (25) pos 20 a 44
                    'digito do codigo de barras (1) pos 5
                    'codigo da moeda (1) pos 4
                    'brancos (13)
                    .sInformacoesCompl = FormataCpoTexto(Mid(sCodigoDeBarras, 20, 25) & Mid(sCodigoDeBarras, 5, 1) & Mid(sCodigoDeBarras, 4, 1), 40)
                    sFatorVencto = FormataCpoNum(Mid(sCodigoDeBarras, 6, 4), 4)
                    .sValorDocto = FormataCpoNum(Mid(sCodigoDeBarras, 10, 10), 10)
                Else
                    .sInformacoesCompl = FormataCpoTexto(Mid(sCodigoDeBarras, 5, 5) & Mid(sCodigoDeBarras, 11, 10) & Mid(sCodigoDeBarras, 22, 10) & Mid(sCodigoDeBarras, 33, 1) & Mid(sCodigoDeBarras, 4, 1), 40)
                    sFatorVencto = FormataCpoNum(Mid(sCodigoDeBarras, 34, 4), 4)
                    .sValorDocto = FormataCpoNum(Mid(sCodigoDeBarras, 38, 10), 10)
                End If
                
                If .sBancoForn = "237" Then
                
                    If Len(Trim(sCodigoDeBarras)) = 44 Then
                    
                        .sCarteira = FormataCpoNum(Mid(sCodigoDeBarras, 24, 2), 3)
                        sNossoNumeroCodBarras = FormataCpoNum(Mid(sCodigoDeBarras, 26, 11), 12)
                        sAgenciaForn = Mid(sCodigoDeBarras, 20, 4)
                        sContaCorrenteForn = Mid(sCodigoDeBarras, 37, 7)
                        
                    Else
                    
                        .sCarteira = FormataCpoNum(Mid(sCodigoDeBarras, 9, 1) & Mid(sCodigoDeBarras, 11, 1), 3)
                        sNossoNumeroCodBarras = FormataCpoNum(Mid(sCodigoDeBarras, 12, 9) & Mid(sCodigoDeBarras, 22, 2), 12)
                        sAgenciaForn = Mid(sCodigoDeBarras, 5, 4)
                        sContaCorrenteForn = Mid(sCodigoDeBarras, 24, 7)
                    
                    End If
                    Call Calcula_DV11(sAgenciaForn, 7, sDVAgenciaForn)
                    Call Calcula_DV11(sContaCorrenteForn, 7, sDVContaCorrenteForn)
                    .sAnoNossoNumero = FormataCpoNum(left(sNossoNumeroCodBarras, 3), 3)
                    .sNossoNumero = FormataCpoNum(Mid(sNossoNumeroCodBarras, 4, 9), 9)
                    .sAgenciaForn = FormataCpoNum(sAgenciaForn, 5)
                    .sDVAgenciaForn = FormataCpoTexto(sDVAgenciaForn, 1)
                    .sContaCorrenteForn = FormataCpoNum(sContaCorrenteForn, 13)
                    .sDVContaCorrenteForn = FormataCpoTexto(sDVContaCorrenteForn, 2)
                    
                Else
                
                    .sCarteira = FormataCpoNum(0, 3)
                    .sAnoNossoNumero = FormataCpoNum(0, 3)
                    .sNossoNumero = FormataCpoNum(0, 9)
                    
                End If
                
            Case Else
                Error 32316
                
        End Select
            
        .sNumeroPagto = FormataCpoNum(objBaixaParcPag.lNumIntParcela, 16)
        .sDataPagto = .sDataVencimento
        .sMoeda = FormataCpoTexto("", 3)
        .sSituacaoAgendamento = "01"
        .sInformacaoRetorno = FormataCpoTexto("", 10)
        .sTipoMovimento = "0"
        .sCodigoMovimento = "00"
        .sEnderecoSacado = FormataCpoTexto("", 40)
        .sSacadorAvalista = FormataCpoTexto("", 40)
        .sReserva1 = FormataCpoTexto("", 1)
        .sNivelInformacao = FormataCpoTexto("", 1)
        
        .sCodigoAreaEmpresa = FormataCpoTexto("", 2)
        .sUsoDaEmpresa = FormataCpoTexto(objBaixaParcPag.lNumIntParcela & SEPARADOR & objBaixaParcPag.iSequencial, 35)
        .sReserva2 = FormataCpoTexto("", 22)
        .sCodLancamento = FormataCpoNum("", 5) '????
        .sReserva3 = FormataCpoTexto("", 1)
        .sTipoContaForn = "1"
        .sReserva4 = FormataCpoTexto("", 8)
        .sNumSequencialRegistro = FormataCpoNum(gobjGeracaoArqCNABPag.lQuantRegistrosArq, 6)
            
        'Monta o registro de detalhe com os dados obtidos
        sRegDetalhe = .sIdentificacao & .sTipoInscrForn & .sInscricaoForn & .sNomeFornecedor & .sEnderecoForn & .sCepForn & .sCepComplementoForn & .sBancoForn & .sAgenciaForn & .sDVAgenciaForn
        sRegDetalhe = sRegDetalhe & .sContaCorrenteForn & .sDVContaCorrenteForn & .sNumeroPagto & .sCarteira & .sAnoNossoNumero & .sNossoNumero & .sSeuNumero & .sDataVencimento & .sDataEmissao
        sRegDetalhe = sRegDetalhe & .sDataDesconto & "0" & sFatorVencto & .sValorDocto & .sValorPagto & .sValorDesconto & .sValorAcrescimo & .sTipoDocumento & .sNumeroTitulo & .sSerieDocumento & .sModalidadePagto & .sDataPagto
        sRegDetalhe = sRegDetalhe & .sMoeda & .sSituacaoAgendamento & .sInformacaoRetorno & .sTipoMovimento & .sCodigoMovimento & .sEnderecoSacado & .sSacadorAvalista & .sReserva1 & .sNivelInformacao
        sRegDetalhe = sRegDetalhe & .sInformacoesCompl & .sCodigoAreaEmpresa & .sUsoDaEmpresa & .sReserva2 & .sCodLancamento & .sReserva3 & .sTipoContaForn & .sContaComplementar & .sReserva4 & .sNumSequencialRegistro

    End With

    Montar_Registro_DetalhePagto = SUCESSO
    
    Exit Function

Erro_Montar_Registro_DetalhePagto:

    Montar_Registro_DetalhePagto = Err
    
    Select Case Err
    
        Case 32316
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_COBR_PAG_INVALIDO", Err)
            
        Case 51802
        
        Case 51803
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ENDERECO_NAO_CADASTRADO", Err, objEndereco.lCodigo)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147214)
            
    End Select

End Function

Function Montar_Registro_TrailerArq(sTrailerArq As String) As Long
'Monta o registro de trailer de arquivo

Dim lErro As Long
Dim tRegTrailer As typePagtoTrailerArqBradesco

On Error GoTo Erro_Montar_Registro_TrailerArq
    
    'Incrementa a quantidade de registros do arquivo
    gobjGeracaoArqCNABPag.lQuantRegistrosArq = gobjGeracaoArqCNABPag.lQuantRegistrosArq + 1

    'Recolhe os dados do trailer
    With tRegTrailer
        .sIdentificacao = "9"
        .sTotalQtdRegistros = FormataCpoNum(gobjGeracaoArqCNABPag.lQuantRegistrosArq, 6)
        .sTotalValorPagtos = FormataCpoValor(gobjGeracaoArqCNABPag.dValorTotalLote, 17)
        .sReserva = FormataCpoTexto("", 470)
        .sNumSequencialRegistro = FormataCpoNum(gobjGeracaoArqCNABPag.lQuantRegistrosArq, 6)
    
        'Monta o registro com os dados recolhidos
        sTrailerArq = .sIdentificacao & .sTotalQtdRegistros & .sTotalValorPagtos & .sReserva & .sNumSequencialRegistro
    
    End With

    Montar_Registro_TrailerArq = SUCESSO
    
    Exit Function
    
Erro_Montar_Registro_TrailerArq:

    Montar_Registro_TrailerArq = Err
    
    Select Case Err
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147215)
            
    End Select
    
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147216)
            
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
        sAgencia = FormataCpoNum(left(sAgenciaDVAg, iPosSeparador - 1), iTamanhoAg)
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
        sConta = FormataCpoNum(left(sContaDVCta, iPosSeparador - 1), iTamanhoCta)
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

Function Calcula_NossoNumero_Bradesco(objCarteiraCobrador As ClassCarteiraCobrador, sNossoNumero As String) As Long

Dim sProxNossoNumero As String
Dim iCarteira As Integer
Dim sDVNossoNumero As String

On Error GoTo Erro_Calcula_NossoNumero_Bradesco
    
    iCarteira = objCarteiraCobrador.iNumCarteiraNoBanco
    
    If objCarteiraCobrador.iComRegistro = CARTEIRA_SEM_REGISTRO Then iCarteira = "06"

    'se for para forçar o nosso numero com o numero do titulo, com a carteira na frente
    If right(objCarteiraCobrador.sFaixaNossoNumeroProx, 10) = "9999999999" And iCarteira = 9 Then
    
        sNossoNumero = FormataCpoNum(gobjCobrancaEletronica.lNumTitulo, 11)
    
    Else
    'Se não tem o nosso número na parcela calcula
        
        If Len(Trim(Replace(objCarteiraCobrador.sFaixaNossoNumeroProx, "0", ""))) = 0 Then gError 79703
        
        sNossoNumero = FormataCpoNum(objCarteiraCobrador.sFaixaNossoNumeroProx, 11)

        sProxNossoNumero = sNossoNumero + 1

        If Len(Trim(objCarteiraCobrador.sFaixaNossoNumeroFinal)) > 0 Then

            If Val(sNossoNumero) > Val(objCarteiraCobrador.sFaixaNossoNumeroFinal) Then sProxNossoNumero = FormataCpoNum(objCarteiraCobrador.sFaixaNossoNumeroInicial, 11)
        End If

        objCarteiraCobrador.sFaixaNossoNumeroProx = sProxNossoNumero
    
    End If
    
    Call Calcula_DV_NossoNumero_Bradesco(FormataCpoNum(iCarteira, 2) & sNossoNumero, sDVNossoNumero, 11, 7)
    
    sNossoNumero = sNossoNumero & sDVNossoNumero
    
    Calcula_NossoNumero_Bradesco = SUCESSO
    
    Exit Function
    
Erro_Calcula_NossoNumero_Bradesco:

    Calcula_NossoNumero_Bradesco = gErr
    
    Select Case gErr
    
        Case 79703
            Call Rotina_Erro(vbOKOnly, "ERRO_NOSSO_NUMERO_NAO_DEFINIDO", gErr, objCarteiraCobrador.iCodCarteiraCobranca, objCarteiraCobrador.iCobrador)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147217)
    
    End Select

    Exit Function

End Function

Private Sub Calcula_DV_NossoNumero_Bradesco(sNossoNumero As String, sDVNossoNumero As String, iModulo As Integer, iBase As Integer)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147218)
     
    End Select
     
    Exit Function

End Function

Private Function CCI_Atualizar(objCCI As ClassContasCorrentesInternas, lComando1 As Long, lComando2 As Long) As Long
'Atualiza o sequencial do proximo arquivo remessa a ser gerado

Dim lErro As Long, lCNABProxSeqArqCobr As Long

On Error GoTo Erro_CCI_Atualizar

    If objCCI.lCNABProxSeqArqCobr <> 0 Then

        lErro = Comando_ExecutarPos(lComando1, "SELECT CNABProxSeqArqCobr FROM ContasCorrentesInternas WHERE Codigo = ?", 0, lCNABProxSeqArqCobr, objCCI.iCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 81700
        
        lErro = Comando_BuscarProximo(lComando1)
        If lErro <> AD_SQL_SUCESSO Then gError 81701
        
        lCNABProxSeqArqCobr = lCNABProxSeqArqCobr + 1
        
        lErro = Comando_ExecutarPos(lComando2, "UPDATE ContasCorrentesInternas SET CNABProxSeqArqCobr = ?", lComando1, lCNABProxSeqArqCobr)
        If lErro <> AD_SQL_SUCESSO Then gError 81702
        
    End If
    
    CCI_Atualizar = SUCESSO
     
    Exit Function
    
Erro_CCI_Atualizar:

    CCI_Atualizar = gErr
     
    Select Case gErr
          
        Case 81700, 81701
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COBRADOR", gErr, objCCI.iCodigo)
        
        Case 81702
            Call Rotina_Erro(vbOKOnly, "ERRO_MODIFICACAO_COBRADOR", gErr, objCCI.iCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147218)
     
    End Select
     
    Exit Function

End Function

Private Sub RegDetCobr_ObterInfo(objDetRetCobr As ClassDetRetCobr, tRegDetalhe As typeRetDetalheBradesco)
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
        .sNossoNumero = left(tRegDetalhe.sNossoNumero1, 11) 'para excluir o dv
        .iCodOcorrencia = StrParaInt(tRegDetalhe.sCodOcorrencia)
        .iCodCarteiraNoBanco = StrParaInt(tRegDetalhe.sCodCarteira)
        .dValorTarifa = Extrair_Valor(tRegDetalhe.sTarifaCobranca)
        .sSeuNumero = tRegDetalhe.sNumControle
        .sUsoDaEmpresa = tRegDetalhe.sNumControle
        .lNumSequencial = CLng(tRegDetalhe.sNumSequencialRegistro)
        .dValorTitulo = Extrair_Valor(tRegDetalhe.sValorTitulo)
        
    End With
    
End Sub

Private Function Prepara_CartCobrRet1(objCobrancaEletronica As ClassCobrancaEletronica, objBorderoRetCobr As ClassBorderoRetCobr, objBorderoRetCobrCart As ClassBorderoRetCobrCart, objDetRetCobr As ClassDetRetCobr, tRegDetalhe As typeRetDetalheBradesco) As Long
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147219)
            
    End Select

    Exit Function

End Function

Private Function Calcula_DV11(ByVal sSequencia As String, ByVal iBase As Integer, sDigito As String, Optional ByVal bAceitaP As Boolean = False) As Long
'Calcula o Dígito Verificador do Nosso Numero - Alterado por Jorge Specian - 15/03/2005
'Cálculo através do módulo 11

Dim lErro As Long
Dim iContador As Integer
Dim iNumero As Integer
Dim iTotalNumero As Integer
Dim iMultiplicador As Integer
Dim iResto As Integer
Dim iResultado As Integer
Dim sCaracter As String

On Error GoTo Erro_Calcula_DV11

    iMultiplicador = 2
    
    For iContador = 1 To Len(sSequencia)
        sCaracter = Mid(right(sSequencia, iContador), 1, 1)
        If iMultiplicador > iBase Then
            iMultiplicador = 2
        End If
        iNumero = sCaracter * iMultiplicador
        iTotalNumero = iTotalNumero + iNumero
        iMultiplicador = iMultiplicador + 1
    Next
    
    iResto = iTotalNumero Mod 11
    
    iResultado = 11 - iResto
    
    If iResultado = 10 Then
        If bAceitaP Then
            sDigito = "P"
        Else
            sDigito = "0"   'zero
        End If
    ElseIf iResultado = 11 Then
        sDigito = "0"   'zero
    Else
        sDigito = CStr(iResultado)
    End If

    Calcula_DV11 = SUCESSO

    Exit Function

Erro_Calcula_DV11:

    Calcula_DV11 = gErr
    
    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 148442)
    
    End Select
    
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

Private Function Montar_Registro4_Bradesco(sRegDetalhe As String, objOcorrRemParcRec As ClassOcorrRemParcRec, objCarteiraCobrador As ClassCarteiraCobrador) As Long

'Monta o registro
Dim lErro As Long
Dim objParcRec As New ClassParcelaReceber
Dim objNF As New ClassNFiscal
Dim sTela As String, iExibeReg4 As Integer

On Error GoTo Erro_Montar_Registro4_Bradesco

    iExibeReg4 = StrParaInt(Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 35))

    If iExibeReg4 = MARCADO Then

        objParcRec.lNumIntDoc = objOcorrRemParcRec.lNumIntParc
        
        lErro = CF("ParcelaReceber_Le", objParcRec)
        If lErro <> SUCESSO And lErro <> 19147 Then gError ERRO_SEM_MENSAGEM
        If lErro <> SUCESSO Then
            lErro = CF("ParcelaReceber_Baixada_Le", objParcRec)
            If lErro <> SUCESSO And lErro <> 58559 Then gError ERRO_SEM_MENSAGEM
            
            If lErro <> SUCESSO Then gError 213224
        End If
        
        lErro = CF("Titulo_Le_DocumentoOriginal", objParcRec.lNumIntTitulo, CPR_TITULO_RECEBER, objNF, sTela)
        If lErro <> SUCESSO And lErro <> 58942 Then gError ERRO_SEM_MENSAGEM
        
        lErro = CF("NFiscal_Le_ChvNFe", objNF)
        If lErro <> SUCESSO And lErro <> 58942 Then gError ERRO_SEM_MENSAGEM
        
        If Len(objNF.sChvNFe) = 44 Then
            sRegDetalhe = "4" & objNF.sChvNFe & "0101" & FormataCpoTexto("", 345) & FormataCpoNum(gobjCobrancaEletronica.lSequencialRegistro, 6)
            gobjCobrancaEletronica.lSequencialRegistro = gobjCobrancaEletronica.lSequencialRegistro + 1
        End If
        
    End If
           
    Montar_Registro4_Bradesco = SUCESSO

    Exit Function

Erro_Montar_Registro4_Bradesco:

    Montar_Registro4_Bradesco = gErr

    Select Case gErr
    
        Case 213224
            Call Rotina_Erro(vbOKOnly, "ERRO_PARCELA_RECEBER_NAO_CADASTRADA1", gErr, objParcRec.lNumIntDoc)
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 213225)

    End Select

    Exit Function

End Function

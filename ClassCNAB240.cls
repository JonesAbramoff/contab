VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCNAB240"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'inclui tratamento dos bancos 1 (ainda existe o original), 104, 479, 275, 356, 409, 399, 341, 745
'procurar texto abaixo para identificar as particularidades de cada bco
'??? com variacao de acordo com o banco

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

Function Criar_Arquivo_Remessa_CNAB240(objCobrancaEletronica As ClassCobrancaEletronica) As Long
'Cria uma arquivo de remessa com base nos dados do BD
'de cobranca
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

Dim bQuebrarLote As Boolean

On Error GoTo Erro_Criar_Arquivo_Remessa_CNAB240
    
    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 51804
    Next
        
    'ABre a transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 51805
    
    Call Zera_Var
    
    Set gobjCobrancaEletronica = objCobrancaEletronica
    Set gobjFilialEmpresa = New AdmFiliais
    
    'Gera o nome do arquivo de acordo com o sequencial do Cobrador
    lErro = CF("CobrancaEletronica_Gera_NomeArquivo", objCobrancaEletronica.objCobrador, objCobrancaEletronica.sDiretorio, sNomeArquivo, lSeqNomeArq)
    If lErro <> SUCESSO Then gError 51806
    
    gobjCobrancaEletronica.lSeqNomeArq = lSeqNomeArq
    
    'Lê a conta corrente associada ao cobrador
    lErro = CF("ContaCorrenteInt_Le", gobjCobrancaEletronica.objCobrador.iCodCCI, gobjCobrancaEletronica.objCCI)
    If lErro <> SUCESSO And lErro <> 11807 Then gError 51814
    If lErro <> SUCESSO Then gError 51815

    If gobjCobrancaEletronica.objCCI.iCodBanco <> gobjCobrancaEletronica.objCobrador.iCodBanco Then gError 62063
    If Len(Trim(gobjCobrancaEletronica.objCCI.sAgencia)) = 0 Or Len(Trim(gobjCobrancaEletronica.objCCI.sNumConta)) = 0 Then gError 62064
    
    gobjFilialEmpresa.iCodFilial = giFilialEmpresa
    
    lErro = CF("FilialEmpresa_Le", gobjFilialEmpresa, True)
    If lErro <> SUCESSO And lErro <> 27378 Then gError 51814
    If lErro <> SUCESSO Then gError 90073
    
    If gobjCobrancaEletronica.objCobrador.iCodBanco = 341 Then
        bQuebrarLote = False
    Else
        bQuebrarLote = True
    End If
    
    'Cria o arquivo de remessa
    Open sNomeArquivo For Output As #1
           
    lErro = Montar_RegistroHeaderArquivo_BcoBrasil(sRegHeader)
    If lErro <> SUCESSO Then gError 51810
        
    If Len(sRegHeader) <> 240 Then gError 206433
    
    If gobjCobrancaEletronica.objCCI.iCodBanco = 453 Or gobjCobrancaEletronica.objCCI.iCodBanco = 749 Then
        sRegHeader = UCase(sRegHeader)
    End If
        
    'Inclui o registro de Header de arquivo
    Print #1, DesacentuaTexto(sRegHeader)
    
    glNumRegistros = glNumRegistros + 1
    
    iIndice = 0
    
    For Each objBorderoCobranca In objCobrancaEletronica.colBorderos

        iIndice = iIndice + 1
        
        'Lê o bordero de cobrança
        lErro = CF("BorderoCobranca_LeMod", alComando(2), objBorderoCobranca)
        If lErro <> SUCESSO Then gError 51807

        objCarteiraCobrador.iCobrador = objCobrancaEletronica.iCobrador
        objCarteiraCobrador.iCodCarteiraCobranca = objBorderoCobranca.iCodCarteiraCobranca

        'Lê a carteira do cobrador
        lErro = CF("CarteiraCobrador_LeMod", objCarteiraCobrador, alComando(4))
        If lErro <> SUCESSO Then gError 51809

        'Busca no BD as Ocorrências do Bordero em questão
        lErro = CF("CobrancaEletronica_Obter_OcorrRemParcRec", alComando(3), objBorderoCobranca, gobjCobrancaEletronica)
        If lErro <> SUCESSO Then gError 51662
        
        If bQuebrarLote Or iIndice = 1 Then
            
            giNumLotes = giNumLotes + 1
            
            'incluir header de lote
            lErro = Montar_RegistroHeaderLote_BcoBrasil(sRegHeader, objCarteiraCobrador)
            If lErro <> SUCESSO Then gError 81697
            
            'Inclui o registro de Header de lote
            If Len(sRegHeader) <> 240 Then gError 206434
            
            If gobjCobrancaEletronica.objCCI.iCodBanco = 453 Or gobjCobrancaEletronica.objCCI.iCodBanco = 749 Then
                sRegHeader = UCase(sRegHeader)
            End If
            
            Print #1, DesacentuaTexto(sRegHeader)
        
            glNumRegistros = glNumRegistros + 1
            
            'Número de Registros dentro do lote
            gobjCobrancaEletronica.lSequencialRegistro = 0
    
        End If
        
        'Para cada ocorrência
         For Each objOcorrRemParc In objCobrancaEletronica.colOcorrRemParcRec

            sRegDetalheSegP = ""
            sRegDetalheSegQ = ""
            sRegDetalheSegR = ""
            
            'Monta um registro de detalhe do arquivo de remessa
            lErro = Montar_RegistroDetalhe_BcoBrasil(sRegDetalheSegP, sRegDetalheSegQ, sRegDetalheSegR, objOcorrRemParc, objCarteiraCobrador)
            If lErro <> SUCESSO And lErro <> 51674 Then gError 51811

            If lErro = SUCESSO Then
            
                If Len(sRegDetalheSegP) <> 240 Then gError 206435
                If Len(sRegDetalheSegQ) <> 240 Then gError 206436
                If Len(sRegDetalheSegR) <> 240 Then gError 206437
            
                If gobjCobrancaEletronica.objCCI.iCodBanco = 1 Or gobjCobrancaEletronica.objCCI.iCodBanco = 453 Or gobjCobrancaEletronica.objCCI.iCodBanco = 749 Then
                    sRegDetalheSegP = UCase(sRegDetalheSegP)
                    sRegDetalheSegQ = UCase(sRegDetalheSegQ)
                    sRegDetalheSegR = UCase(sRegDetalheSegR)
                End If
                
                'Inclui o registro com os detalhes do Título(Parcela) a receber
                Print #1, DesacentuaTexto(sRegDetalheSegP)
                Print #1, DesacentuaTexto(sRegDetalheSegQ)
                Print #1, DesacentuaTexto(sRegDetalheSegR)
    
                glNumRegistros = glNumRegistros + 3
                
            End If
            
        Next

        If bQuebrarLote Or (iIndice = objCobrancaEletronica.colBorderos.Count) Then
        
            'incluir trailer de lote
        
            'MOnta o registro de trailer de lote
            lErro = Montar_RegistroTrailerLote_BcoBrasil(sRegTrailer)
            If lErro <> SUCESSO Then gError 81698
        
            If Len(sRegTrailer) <> 240 Then gError 206438
            
            If gobjCobrancaEletronica.objCCI.iCodBanco = 453 Or gobjCobrancaEletronica.objCCI.iCodBanco = 749 Then
                sRegTrailer = UCase(sRegTrailer)
            End If
        
            'grava o registro trailer de lote
            Print #1, DesacentuaTexto(sRegTrailer)
            
            glNumRegistros = glNumRegistros + 1
        
        End If
        
        objBorderoCobranca.iProcessado = BORDERO_PROCESSADO
        objBorderoCobranca.sNomeArquivo = sNomeArquivo

        'Atualiza o Status do Bordero de Cobrança
        lErro = CF("BorderoCobranca_Atualiza_Processo", alComando(2), alComando(3), objBorderoCobranca)
        If lErro <> SUCESSO Then gError 51808

'        'Atualiza o NossoNumero em carteiraCobrador
'        lErro = CF("CarteiraCobrador_AtualizaNossoNumero", objCarteiraCobrador, alComando(4), alComando(5))
'        If lErro <> SUCESSO Then gError 51842
        
    Next

    'MOnta o registro de trailer de arquivo
    lErro = Montar_RegistroTrailerArq_BcoBrasil(sRegTrailer)
    If lErro <> SUCESSO Then gError 51812

    If Len(sRegTrailer) <> 240 Then gError 206439
    
    If gobjCobrancaEletronica.objCCI.iCodBanco = 453 Or gobjCobrancaEletronica.objCCI.iCodBanco = 749 Then
        sRegTrailer = UCase(sRegTrailer)
    End If

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
    
    Criar_Arquivo_Remessa_CNAB240 = SUCESSO
    
    Exit Function

Erro_Criar_Arquivo_Remessa_CNAB240:

    Criar_Arquivo_Remessa_CNAB240 = gErr
    
    Select Case gErr
    
        Case 51662, 51806, 51807, 51808, 51809, 51810, 51811, 51812, 51814, 51842, 81697, 81698, 81699
    
        Case 51804
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 51805
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
                    
        Case 51815
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_COBRADOR_NAO_ENCONTRADA", gErr, gobjCobrancaEletronica.iCobrador)
        
        Case 62011
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case 62063
            Call Rotina_Erro(vbOKOnly, "ERRO_BANCO_CCI_DIFERENTE_COBRADOR", Err, gobjCobrancaEletronica.objCCI.iCodBanco, gobjCobrancaEletronica.objCobrador.iCodBanco)
    
        Case 62064
            Call Rotina_Erro(vbOKOnly, "ERRO_AGENCIA_CONTA_COBRADOR_NAO_PREENCHIDAS", Err)
        
        Case 90073
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, giFilialEmpresa)
            
        Case 206433
            Call Rotina_Erro(vbOKOnly, "ERRO_REG_HEADER_ARQ_TAM_ERRADO", gErr)
        
        Case 206434
            Call Rotina_Erro(vbOKOnly, "ERRO_REG_HEADER_LOTE_TAM_ERRADO", gErr)
        
        Case 206435
            Call Rotina_Erro(vbOKOnly, "ERRO_REG_DET_P_TAM_ERRADO", gErr)
        
        Case 206436
            Call Rotina_Erro(vbOKOnly, "ERRO_REG_DET_R_TAM_ERRADO", gErr)
        
        Case 206437
            Call Rotina_Erro(vbOKOnly, "ERRO_REG_DET_Q_TAM_ERRADO", gErr)
        
       Case 206438
            Call Rotina_Erro(vbOKOnly, "ERRO_REG_TRAILLER_LOTE_TAM_ERRADO", gErr)
        
        Case 206439
            Call Rotina_Erro(vbOKOnly, "ERRO_REG_TRAILLER_ARQ_TAM_ERRADO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147077)
    
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
'para cobrança
Dim tRegHeaderCNAB240 As typeRemHeaderCNAB240
Dim lErro As Long, sCodigoCedente As String, lCodigoModeloPersonalizado As Long

On Error GoTo Erro_Montar_RegistroHeaderLote_BcoBrasil

    'Carrega tRegHeaderBcoBrasil com os dados do Header do aquivo
    lErro = Obter_Dados_Header_Lote(tRegHeaderCNAB240, objCarteiraCobrador)
    If lErro <> SUCESSO Then Error 51813
        
    With tRegHeaderCNAB240
    
        'Monta o registro
        sRegHeader = .sCodBancoComp & .sLoteServico & .sRegHeaderLote & .sTipoOperacao & .sTipoServico & .sFormaLancamento & .sNVersaoLayLote & .sBrancosCNAB1 & .sTipoInscEmpresa & .sCodInsEmpresa
        
        '??? com variacao de acordo com o banco
        Select Case gobjCobrancaEletronica.objCobrador.iCodBanco
        
            Case 1
                 sRegHeader = sRegHeader & .sCodConvBanco & .sAgMantConta & .sDVAgencia & .sNumContaCorrente & .sDVConta & .sDVAgenciaConta
                             
            Case 8, 33, 353
                 sRegHeader = sRegHeader & FormataCpoTexto("", 20) & .sCodConvBanco & FormataCpoTexto("", 5)
                             
            Case 104
                If gobjCobrancaEletronica.sVersaoLayout = "050" Then
                    lCodigoModeloPersonalizado = 0 '???
                    sRegHeader = sRegHeader & .sCodConvBanco & FormataCpoNum(0, 14) & .sAgMantConta & .sDVAgencia & .sCodConvBanco & FormataCpoNum(lCodigoModeloPersonalizado, 7) & "0"
                Else
                    sRegHeader = sRegHeader & .sCodConvBanco & FormataCpoTexto("", 4) & .sAgMantConta & .sDVAgencia & .sNumContaCorrente & .sDVConta & .sDVAgenciaConta
                End If
                
            Case 479
                sRegHeader = sRegHeader & "1" & .sCodConvBanco & "014        " & .sAgMantConta & .sDVAgencia & "000000" & FormataCpoNum(gobjCobrancaEletronica.objCCI.sNumConta, 8)
                
            Case 275, 356
                sRegHeader = sRegHeader & .sCodConvBanco & .sAgMantConta & " " & .sNumContaCorrente & "  "
                
            Case 409
                sRegHeader = sRegHeader & .sCodConvBanco & .sAgMantConta & .sDVAgencia & .sNumContaCorrente & .sDVConta & " "
                
            Case 399
                sRegHeader = sRegHeader & .sCodConvBanco & .sAgMantConta & "0" & .sNumContaCorrente & .sDVConta & "0"
                
            Case 341
                sRegHeader = sRegHeader & .sCodConvBanco & .sAgMantConta & " " & .sNumContaCorrente & " " & .sDVConta
                
            Case 745
                sRegHeader = sRegHeader & .sCodConvBanco & .sAgMantConta & .sDVAgencia & .sNumContaCorrente & .sDVConta & .sDVConta
            
            Case 246
                 sRegHeader = sRegHeader & .sCodConvBanco & .sAgMantConta & .sDVAgencia & .sNumContaCorrente & .sDVConta & .sDVAgenciaConta

            Case 41
                 sRegHeader = sRegHeader & .sCodConvBanco & .sAgMantConta & .sDVAgencia & .sNumContaCorrente & .sDVConta & .sDVAgenciaConta

            Case 453, 749
                 sRegHeader = sRegHeader & .sCodConvBanco & .sAgMantConta & .sDVAgencia & .sNumContaCorrente & .sDVConta & .sDVAgenciaConta

            Case 422
                 sRegHeader = sRegHeader & .sCodConvBanco & .sAgMantConta & .sDVAgencia & .sNumContaCorrente & .sDVConta & .sDVAgenciaConta

            Case 748
                 sRegHeader = sRegHeader & .sCodConvBanco & .sAgMantConta & .sDVAgencia & .sNumContaCorrente & .sDVConta & .sDVAgenciaConta

            Case 637
                 sRegHeader = sRegHeader & .sCodConvBanco & .sAgMantConta & .sDVAgencia & .sNumContaCorrente & .sDVConta & .sDVAgenciaConta

            Case 85
                 sRegHeader = sRegHeader & .sCodConvBanco & .sAgMantConta & .sDVAgencia & .sNumContaCorrente & .sDVConta & .sDVAgenciaConta

        End Select
        
        sRegHeader = sRegHeader & .sNomeEmpresa & .sMensagem1 & .sMensagem2 & .sNumRemessa & .sDataGravRem & .sDataCredito & .sBrancosCNAB2
        
    End With
    
    Montar_RegistroHeaderLote_BcoBrasil = SUCESSO

    Exit Function

Erro_Montar_RegistroHeaderLote_BcoBrasil:

    Montar_RegistroHeaderLote_BcoBrasil = Err

    Select Case Err

        Case 51813

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147078)

    End Select

    Exit Function

End Function

Private Function Obter_Dados_Header_Lote(tRegHeaderCNAB240 As typeRemHeaderCNAB240, objCarteiraCobrador As ClassCarteiraCobrador) As Long
'para cobranca
Dim lErro As Long, sDVConta As String
Dim sCodConvBco As String

On Error GoTo Erro_Obter_Dados_Header_Lote
      
    'Carrega os campos do cabeçalho
    With tRegHeaderCNAB240
   
        .sCodBancoComp = FormataCpoNum(gobjCobrancaEletronica.objCobrador.iCodBanco, 3)
        .sLoteServico = FormataCpoNum(giNumLotes, 4)
        .sRegHeaderLote = "1"
        .sTipoOperacao = "R"
        .sTipoServico = "01"
        .sFormaLancamento = FormataCpoNum("0", 2)
        .sDataCredito = FormataCpoNum(0, 8)
        .sAgMantConta = FormataCpoNum(gobjCobrancaEletronica.objCCI.sAgencia, 5)
        .sDVAgencia = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVAgencia, 1)
        .sDVAgenciaConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVAgConta, 1)
        
        '??? com variacao de acordo com o banco
        Select Case gobjCobrancaEletronica.objCobrador.iCodBanco
        
            Case 1
                .sNVersaoLayLote = "020"
                .sCodConvBanco = FormataCpoTexto(Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 7), 20)
                .sNumContaCorrente = FormataCpoNum(gobjCobrancaEletronica.objCCI.sNumConta, 12)
                .sDVConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVNumConta, 1)
        
            Case 8, 33, 353
                .sFormaLancamento = FormataCpoTexto("", 2)
                .sNVersaoLayLote = "030"
                sCodConvBco = Obter_Info_CarteiraCobrador(CARTEIRA_SIMPLES, gobjCobrancaEletronica.colCarteirasCNABInfo, 7)
                If Len(Trim(sCodConvBco)) <= 7 Then sCodConvBco = FormataCpoNum(gobjCobrancaEletronica.objCCI.sAgencia, 4) & "0000" & FormataCpoNum(Trim(sCodConvBco), 7)
                .sCodConvBanco = FormataCpoNum(sCodConvBco, 15)
                .sDataCredito = FormataCpoTexto("", 8)
                
            Case 104
                
                If gobjCobrancaEletronica.sVersaoLayout = "050" Then
                    .sNVersaoLayLote = "030"
                    .sCodConvBanco = FormataCpoNum(Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 7), 6)
                
                Else 'para versao 030 de arquivo
                    .sNVersaoLayLote = "020"
                    .sCodConvBanco = FormataCpoNum(Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 7), 16)
                    .sAgMantConta = FormataCpoNum(Mid(.sCodConvBanco, 1, 4), 5)
                    Call Calcula_DV11(.sAgMantConta, 9, sDVConta)
                    .sDVAgencia = sDVConta
                    
                    .sNumContaCorrente = FormataCpoNum(Mid(.sCodConvBanco, 8, 8), 12) '???? tem que pegar do cedente
                    'calcular dv da conta
                    'Call Calcula_DV11(.sNumContaCorrente, 9, sDVConta)
                    .sDVConta = "0" 'sDVConta
                    
                    Call Calcula_DV11(left(.sCodConvBanco, 4) & .sNumContaCorrente, 9, sDVConta)
                    .sDVAgenciaConta = sDVConta
                End If
                
            Case 479
                .sNVersaoLayLote = "020"
                .sCodConvBanco = FormataCpoNum(Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 7), 8)
                .sNumContaCorrente = FormataCpoNum(gobjCobrancaEletronica.objCCI.sNumConta, 12)
                .sDVConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVNumConta, 1)
        
            Case 275, 356
                .sNVersaoLayLote = "040"
                .sCodConvBanco = FormataCpoTexto("", 20)
                .sNumContaCorrente = FormataCpoNum(gobjCobrancaEletronica.objCCI.sNumConta, 12)
                .sDVConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVNumConta, 1)
        
            Case 409
                .sFormaLancamento = FormataCpoTexto("", 2) '???
                .sNVersaoLayLote = "041"
                .sCodConvBanco = FormataCpoTexto("", 20)
                .sNumContaCorrente = FormataCpoNum(gobjCobrancaEletronica.objCCI.sNumConta, 12)
                .sDVConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVNumConta, 1)
        
            Case 399
                .sNVersaoLayLote = "010"
                .sCodConvBanco = FormataCpoTexto(Obter_Info_CarteiraCobrador(CARTEIRA_SIMPLES, gobjCobrancaEletronica.colCarteirasCNABInfo, 22), 20)
                .sNumContaCorrente = FormataCpoNum(gobjCobrancaEletronica.objCCI.sNumConta, 12)
                '.sDVConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVNumConta, 1)
                .sDVConta = "0"
        
            Case 341
                .sNVersaoLayLote = "030"
                .sCodConvBanco = FormataCpoTexto("", 20)
                .sNumContaCorrente = FormataCpoNum(gobjCobrancaEletronica.objCCI.sNumConta, 12)
                .sDVConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVNumConta, 1)
                
            Case 745
                .sNVersaoLayLote = "041"
                .sCodConvBanco = FormataCpoTexto("", 20)
                .sNumContaCorrente = FormataCpoNum(gobjCobrancaEletronica.objCCI.sNumConta, 12)
                .sDVConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVNumConta, 1)
                
            Case 246
                .sNVersaoLayLote = "030"
                If Len(Trim(Obter_Info_CarteiraCobrador(CARTEIRA_SIMPLES, gobjCobrancaEletronica.colCarteirasCNABInfo, 7))) = 0 Then
                    .sCodConvBanco = FormataCpoNum(gobjCobrancaEletronica.objCCI.sAgencia, 5) & FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVAgencia, 1) & FormataCpoNum(gobjCobrancaEletronica.objCCI.sNumConta, 12) & FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVNumConta, 1) & FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVAgConta, 1)
                    .sNumContaCorrente = FormataCpoNum(gobjCobrancaEletronica.objCCI.sNumConta, 12)
                    .sDVConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVNumConta, 1)
                    .sAgMantConta = FormataCpoNum(gobjCobrancaEletronica.objCCI.sAgencia, 5)
                    .sDVAgencia = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVAgencia, 1)
                    .sDVAgenciaConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVAgConta, 1)
                Else
                    .sCodConvBanco = FormataCpoTexto(Obter_Info_CarteiraCobrador(CARTEIRA_SIMPLES, gobjCobrancaEletronica.colCarteirasCNABInfo, 7), 20)
                    .sAgMantConta = Mid(.sCodConvBanco, 1, 5)
                    .sDVAgencia = Mid(.sCodConvBanco, 6, 1)
                    .sNumContaCorrente = Mid(.sCodConvBanco, 7, 12)
                    .sDVConta = Mid(.sCodConvBanco, 19, 1)
                    .sDVAgenciaConta = Mid(.sCodConvBanco, 20, 1)
                End If
                
            Case 41
                .sNVersaoLayLote = "020"
                .sCodConvBanco = FormataCpoTexto(Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 7), 20)
                .sNumContaCorrente = FormataCpoNum(gobjCobrancaEletronica.objCCI.sNumConta, 12)
                .sDVConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVNumConta, 1)
            
            Case 453, 749
                .sNVersaoLayLote = "020"
                .sCodConvBanco = FormataCpoTexto("Cobrança", 20)
                .sNumContaCorrente = "000" & FormataCpoNum(gobjCobrancaEletronica.objCCI.sNumConta, 10)
                .sDVConta = "0"
                .sDVAgenciaConta = ""
                
            Case 422
                .sNVersaoLayLote = "045"
                .sNumContaCorrente = FormataCpoNum(gobjCobrancaEletronica.objCCI.sNumConta, 12)
                .sDVConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVNumConta, 1)
                .sCodConvBanco = FormataCpoTexto(Obter_Info_Cobrador(gobjCobrancaEletronica.objCobradorCNABInfo.colInformacoes, 1), 20)
            
            Case 748
                .sNVersaoLayLote = "040"
                .sNumContaCorrente = FormataCpoNum(gobjCobrancaEletronica.objCCI.sNumConta, 12)
                .sDVConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVNumConta, 1)
                .sCodConvBanco = FormataCpoTexto(Obter_Info_Cobrador(gobjCobrancaEletronica.objCobradorCNABInfo.colInformacoes, 1), 20)
            
            Case 637
                .sNVersaoLayLote = "042"
                .sNumContaCorrente = FormataCpoNum(gobjCobrancaEletronica.objCCI.sNumConta, 12)
                .sDVConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVNumConta, 1)
                .sCodConvBanco = FormataCpoTexto(Obter_Info_Cobrador(gobjCobrancaEletronica.objCobradorCNABInfo.colInformacoes, 1), 20)
            
            Case 85
                .sNVersaoLayLote = "045"
                .sNumContaCorrente = FormataCpoNum(gobjCobrancaEletronica.objCCI.sNumConta, 12)
                .sDVConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVNumConta, 1)
                .sCodConvBanco = FormataCpoTexto(Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 7), 20)
                .sFormaLancamento = FormataCpoTexto("", 2)
            
        End Select
        
        .sBrancosCNAB1 = FormataCpoTexto("", 1)
        
        If Len(gobjCobrancaEletronica.sCGCCedente) = 14 Then
            .sTipoInscEmpresa = "2"
        Else
            .sTipoInscEmpresa = "1"
        End If
        
        .sCodInsEmpresa = FormataCpoNum(gobjCobrancaEletronica.sCGCCedente, 15)
                
        .sNomeEmpresa = FormataCpoTexto(gobjCobrancaEletronica.sCedente, 30)
        
        Call CF("Empresa_Le_Nome_Customizado", .sNomeEmpresa)
        
        .sMensagem1 = FormataCpoTexto("", 40)
        .sMensagem2 = FormataCpoTexto("", 40)
        .sNumRemessa = FormataCpoNum(gobjCobrancaEletronica.objCobrador.lCNABProxSeqArqCobr, 8)
        .sDataGravRem = Format(gdtDataHoje, "DDMMYYYY")
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147079)
            
    End Select
    
    Exit Function

End Function

Private Function Montar_RegistroDetalhe_BcoBrasil(sRegDetalheSegP As String, sRegDetalheSegQ As String, sRegDetalheSegR As String, objOcorrRemParcRec As ClassOcorrRemParcRec, objCarteiraCobrador As ClassCarteiraCobrador) As Long
'Monta o registro
'para cobranca
Dim tRegDetalheCNAB240SegP As typeRemDetalheCNAB240SegP
Dim tRegDetalheCNAB240SegQ As typeRemDetalheCNAB240SegQ
Dim tRegDetalheCNAB240SegR As typeRemDetalheCNAB240SegR
Dim lErro As Long, sCodConvBanco As String
Dim objParcelaRec As New ClassParcelaReceber
 
On Error GoTo Erro_Montar_RegistroDetalhe_BcoBrasil
    
    objParcelaRec.lNumIntDoc = objOcorrRemParcRec.lNumIntParc

    'Busca os dados da parcela a receber no BD
    lErro = CF("CobrancaEletronica_Obter_DadosParcelaRec", objParcelaRec, gobjCobrancaEletronica)
    If lErro <> SUCESSO And lErro <> 51674 Then Error 51817
    If lErro <> SUCESSO Then Error 51674
                    
    gobjCobrancaEletronica.lQuantTitulos = gobjCobrancaEletronica.lQuantTitulos + 1
      
    gobjCobrancaEletronica.lSequencialRegistro = gobjCobrancaEletronica.lSequencialRegistro + 1
    
    'Obtem os dados do registro de detalhe SegmentoP
    lErro = Obter_Dados_DetalheSegP(tRegDetalheCNAB240SegP, objOcorrRemParcRec, objCarteiraCobrador, objParcelaRec)
    If lErro <> SUCESSO Then Error 51816
    
    With tRegDetalheCNAB240SegP
        'Forma o registro de detalhe
        sRegDetalheSegP = .sCodBancoComp & .sLoteServico & .sRegDetalhe & .sNumSequencialReg & .sCodSegRegDetalhe & .sBrancosCNAB1 & .sCodMovimento & .sAgMantConta
        
        '??? com variacao de acordo com o banco
        Select Case gobjCobrancaEletronica.objCobrador.iCodBanco
        
            Case 1
                sRegDetalheSegP = sRegDetalheSegP & .sDVAgMantConta & .sNumContaCorrente & .sDVConta & .sDVAgenciaConta & .sIdentTituloBanco & .sCodCarteira & .sFormaCadTituloBanco & .sTipoDocumento
                
            Case 8, 33, 353
                sRegDetalheSegP = sRegDetalheSegP & .sDVAgMantConta & .sNumContaCorrente & .sDVConta & .sContaCobranca & .sDVContaCobranca & FormataCpoTexto("", 2) & .sIdentTituloBanco & .sCodCarteira & .sFormaCadTituloBanco & .sTipoDocumento
                
            Case 104
                If gobjCobrancaEletronica.sVersaoLayout = "050" Then
                    sRegDetalheSegP = sRegDetalheSegP & .sDVAgMantConta & .sNumContaCorrente & FormataCpoNum(0, 11) & .sIdentTituloBanco & .sCodCarteira & .sFormaCadTituloBanco & .sTipoDocumento
                Else 'arquivo com versao 030
                    sRegDetalheSegP = sRegDetalheSegP & .sDVAgMantConta & .sNumContaCorrente & .sDVConta & .sDVAgenciaConta & FormataCpoTexto("", 9) & .sIdentTituloBanco & .sCodCarteira & .sFormaCadTituloBanco & .sTipoDocumento
                End If
                
            Case 479
                sRegDetalheSegP = sRegDetalheSegP & .sDVAgMantConta & "000000" & .sNumContaCorrente & .sIdentTituloBanco & FormataCpoTexto("", 11) & .sCodCarteira & .sFormaCadTituloBanco & .sTipoDocumento
                
            Case 275, 356
                sRegDetalheSegP = sRegDetalheSegP & " " & .sNumContaCorrente & "  0000000" & .sIdentTituloBanco & "000"
        
            Case 409
                sRegDetalheSegP = sRegDetalheSegP & .sDVAgMantConta & .sNumContaCorrente & .sDVConta & .sDVAgenciaConta & .sIdentTituloBanco & .sCodCarteira & .sFormaCadTituloBanco & .sTipoDocumento
        
            Case 399
                sRegDetalheSegP = sRegDetalheSegP & .sDVAgMantConta & .sNumContaCorrente & .sDVConta & .sDVAgenciaConta & .sIdentTituloBanco & FormataCpoTexto("", 9) & .sCodCarteira & .sFormaCadTituloBanco & .sTipoDocumento
        
            Case 341
                sRegDetalheSegP = sRegDetalheSegP & " " & .sNumContaCorrente & " " & .sDVConta & .sCodCarteira & .sIdentTituloBanco & FormataCpoTexto("", 8) & "00000"
                
            Case 745
                sRegDetalheSegP = sRegDetalheSegP & .sDVAgMantConta & .sNumContaCorrente & .sDVConta & .sDVAgenciaConta & .sIdentTituloBanco & .sCodCarteira & .sFormaCadTituloBanco & .sTipoDocumento
            
            Case 246
                sRegDetalheSegP = sRegDetalheSegP & .sDVConta & .sNumContaCorrente & .sDVConta & .sDVAgenciaConta & "000000112" & .sIdentTituloBanco & .sCodCarteira & .sFormaCadTituloBanco & .sTipoDocumento
                
            Case 41
                sRegDetalheSegP = sRegDetalheSegP & .sDVAgMantConta & .sNumContaCorrente & .sDVConta & .sDVAgenciaConta & .sIdentTituloBanco & .sCodCarteira & .sFormaCadTituloBanco & .sTipoDocumento
                
            Case 453, 749
                sRegDetalheSegP = sRegDetalheSegP & .sDVAgMantConta & .sNumContaCorrente & .sDVConta & .sDVAgenciaConta & .sIdentTituloBanco & .sCodCarteira & .sFormaCadTituloBanco & .sTipoDocumento
            
            Case 422
                sRegDetalheSegP = sRegDetalheSegP & .sDVAgMantConta & .sNumContaCorrente & .sDVConta & .sDVAgenciaConta & .sIdentTituloBanco & .sCodCarteira & .sFormaCadTituloBanco & .sTipoDocumento
                
            Case 748
                sRegDetalheSegP = sRegDetalheSegP & .sDVAgMantConta & .sNumContaCorrente & .sDVConta & .sDVAgenciaConta & .sIdentTituloBanco & .sCodCarteira & .sFormaCadTituloBanco & .sTipoDocumento
                
            Case 637
                sRegDetalheSegP = sRegDetalheSegP & .sDVAgMantConta & .sNumContaCorrente & .sDVConta & .sDVAgenciaConta & .sIdentTituloBanco & .sCodCarteira & .sFormaCadTituloBanco & .sTipoDocumento
                
            Case 85
                sRegDetalheSegP = sRegDetalheSegP & .sDVAgMantConta & .sNumContaCorrente & .sDVConta & .sDVAgenciaConta & .sIdentTituloBanco & .sCodCarteira & .sFormaCadTituloBanco & .sTipoDocumento
                
        End Select
        
        sRegDetalheSegP = sRegDetalheSegP & .sIdentEmissaoBloq & .sIdentDistrib & .sNumDocCobranca & .sDataVencTitulo & .sValorNominalTitulo & .sAgEncCobranca & .sDVAgencia & .sEspecieTitulo
        sRegDetalheSegP = sRegDetalheSegP & .sIdentTituloAceite & .sDataEmissaoTitulo & .sCodJurosMora & .sDataJurosMora & .sJurosMoraDiaTaxa & .sCodDesconto1 & .sDataDesconto1
        sRegDetalheSegP = sRegDetalheSegP & .sValorPerConcedido & .sValorIOF & .sValorAbatimento & .sIdentTituloEmp & .sCodProtesto & .sNumDiasProtesto
        sRegDetalheSegP = sRegDetalheSegP & .sCodBaixaDevolucao & .sNumDiasBaixaDevol & .sCodMoeda & .sNumContratoOpCred & .sBrancosCNAB2
        
    End With
    
    gobjCobrancaEletronica.lSequencialRegistro = gobjCobrancaEletronica.lSequencialRegistro + 1
    
    'Obtem os dados do registro de detalhe SegmentoQ
    lErro = Obter_Dados_DetalheSegQ(tRegDetalheCNAB240SegQ, objOcorrRemParcRec, objCarteiraCobrador, objParcelaRec)
    If lErro <> SUCESSO Then Error 51816
    
    With tRegDetalheCNAB240SegQ
        'Forma o registro de detalhe
        sRegDetalheSegQ = .sCodBancoComp & .sLoteServico & .sRegDetalhe & .sNumSequencialReg & .sCodSegRegDetalhe & .sBrancosCNAB1 & .sCodMovimento
        sRegDetalheSegQ = sRegDetalheSegQ & .sTipoInscricao & .sNumInscricao & .sNome & .sEndereco & .sBairro & .sCEP & .sSufixoCEP & .sCidade & .sUnidFederacao
        sRegDetalheSegQ = sRegDetalheSegQ & .sTipoInscricaoSacAval & .sNumInscricaoSacAval & .sNomeSacadorAvalista & .sBancoCompensacao & .sNossoNumBancoCorresp & .sBrancosCNAB2
                
    End With
    
    gobjCobrancaEletronica.lSequencialRegistro = gobjCobrancaEletronica.lSequencialRegistro + 1
        
    'Obtem os dados do registro de detalhe SegmentoR
    lErro = Obter_Dados_DetalheSegR(tRegDetalheCNAB240SegR, objOcorrRemParcRec, objCarteiraCobrador, objParcelaRec)
    If lErro <> SUCESSO Then Error 51816
    
    With tRegDetalheCNAB240SegR
        'Forma o registro de detalhe
        sRegDetalheSegR = .sCodBancoComp & .sLoteServico & .sRegDetalhe & .sNumSequencialReg & .sCodSegRegDetalhe & .sBrancosCNAB1 & .sCodMovimento
        sRegDetalheSegR = sRegDetalheSegR & .sCodDesconto2 & .sDataDesconto2 & .sValorPerConcedido2 & .sCodDesconto3 & .sDataDesconto3 & .sValorPerConcedido3
        sRegDetalheSegR = sRegDetalheSegR & .sCodMulta & .sDataMulta & .sValorPerAplicado & .sInfoBancoSacado & .sMensagem3 & .sMensagem4
        
        '??? com variacao de acordo com o banco
        Select Case gobjCobrancaEletronica.objCobrador.iCodBanco
        
            Case 1
                sRegDetalheSegR = sRegDetalheSegR & .sCodBancoContaDeb & .sCodAgContaDeb & .sCodContaDVDeb & .sCodOcorrenciaSac & .sBrancosCNAB2
                
            Case 8, 33, 353
                sRegDetalheSegR = sRegDetalheSegR & FormataCpoTexto("", 50 + 11)
                
            Case 104
                sRegDetalheSegR = sRegDetalheSegR & FormataCpoTexto("", 50 + 11)
        
            Case 479
                sRegDetalheSegR = sRegDetalheSegR & "479" & FormataCpoNum(0, 4 + 3 + 2 + 8) & FormataCpoTexto("", 41)
        
            Case 275, 356
                sRegDetalheSegR = sRegDetalheSegR & FormataCpoTexto("", 20) & FormataCpoNum(0, 16) & " " & FormataCpoNum(0, 12) & "  0" & FormataCpoTexto("", 9)
        
            Case 409
                sRegDetalheSegR = sRegDetalheSegR & FormataCpoTexto("", 20) & FormataCpoNum(0, 8) & FormataCpoNum(gobjCobrancaEletronica.objCobrador.iCodBanco, 3) & FormataCpoNum(gobjCobrancaEletronica.objCCI.sAgencia, 5) & FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVAgencia, 1) & _
                    FormataCpoNum(gobjCobrancaEletronica.objCCI.sNumConta, 12) & FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVNumConta, 1) & FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVAgConta, 1) & "0" & FormataCpoTexto("", 9)
                
            Case 399
                sRegDetalheSegR = sRegDetalheSegR & FormataCpoTexto("", 61)
        
            Case 341
                sRegDetalheSegR = sRegDetalheSegR & FormataCpoTexto("", 20) & FormataCpoNum(0, 16) & " " & FormataCpoNum(0, 12) & "  0" & FormataCpoTexto("", 9)
                
            Case 745
                sRegDetalheSegR = sRegDetalheSegR & .sCodBancoContaDeb & .sCodAgContaDeb & .sCodContaDVDeb & .sCodOcorrenciaSac & .sBrancosCNAB2
            
            Case 246
                sRegDetalheSegR = sRegDetalheSegR & FormataCpoTexto("", 20) & .sCodOcorrenciaSac & FormataCpoTexto("", 23) & "0" & FormataCpoTexto("", 9)
            
            Case 41
                sRegDetalheSegR = sRegDetalheSegR & FormataCpoTexto("", 50 + 11)
            
            Case 453, 749
                sRegDetalheSegR = sRegDetalheSegR & FormataCpoTexto("", 61)
            
            Case 422
                sRegDetalheSegR = sRegDetalheSegR & FormataCpoTexto("", 61)
            
            Case 748
                sRegDetalheSegR = sRegDetalheSegR & FormataCpoTexto("", 61)
            
            Case 637
                sRegDetalheSegR = sRegDetalheSegR & FormataCpoTexto("", 20) & FormataCpoNum(0, 16) & " " & FormataCpoNum(0, 12) & "  0" & FormataCpoTexto("", 9)
            
            Case 85
                sRegDetalheSegR = sRegDetalheSegR & FormataCpoTexto("", 20) & FormataCpoNum(0, 16) & " " & FormataCpoNum(0, 12) & "  0" & FormataCpoTexto("", 9)
            
        End Select

    End With
    
    Montar_RegistroDetalhe_BcoBrasil = SUCESSO

    Exit Function

Erro_Montar_RegistroDetalhe_BcoBrasil:

     Montar_RegistroDetalhe_BcoBrasil = Err

    Select Case Err
    
        Case 51816, 51674, 51817

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147080)

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
            iPosicao = iIndice
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147081)
            
    End Select
    
    Exit Function
    
End Function

Private Function Montar_RegistroTrailerLote_BcoBrasil(sRegTrailer As String) As Long
'Monta o registro
'para cobranca
Dim tRegTrailerCNAB240 As typeRemTrailerCNAB240
Dim lErro As Long
Dim lQtdeTit As Long, dVlrTit As Double
Dim objOcorrRemParc As ClassOcorrRemParcRec

On Error GoTo Erro_Montar_RegistroTrailerLote_BcoBrasil
    
    With tRegTrailerCNAB240
        .sCodBancoComp = FormataCpoNum(gobjCobrancaEletronica.objCobrador.iCodBanco, 3)
        .sLoteServico = FormataCpoNum(giNumLotes, 4)
        .sRegTrailer = "5"
        .sBrancosCNAB1 = FormataCpoTexto("", 9)
        .sQuantRegLote = FormataCpoNum((gobjCobrancaEletronica.lSequencialRegistro + 2), 6)
        
        
        '??? com variacao de acordo com o banco
        Select Case gobjCobrancaEletronica.objCobrador.iCodBanco
        
            Case 8, 33, 353
                .sQuantTitCobS = FormataCpoTexto("", 6)
                .sValTotalTitCartS = FormataCpoTexto("", 17)
                .sQuantTitCobV = FormataCpoTexto("", 6)
                .sValTotalTitCartV = FormataCpoTexto("", 17)
                .sQuantTitCobC = FormataCpoTexto("", 6)
                .sValTotalTitCartC = FormataCpoTexto("", 17)
                .sQuantTitCobD = FormataCpoTexto("", 6)
                .sValTotalTitCartD = FormataCpoTexto("", 17)
                
            Case 104
            
                If gobjCobrancaEletronica.sVersaoLayout = "050" Then
            
                    lQtdeTit = 0
                    dVlrTit = 0
            
                    For Each objOcorrRemParc In gobjCobrancaEletronica.colOcorrRemParcRec
                        lQtdeTit = lQtdeTit + 1
                        dVlrTit = dVlrTit + objOcorrRemParc.dValorCobrado
                    Next
           
                    .sQuantTitCobS = FormataCpoNum(lQtdeTit, 6)
                    .sValTotalTitCartS = FormataCpoValor(dVlrTit, 17)
            
                Else
            
                    .sQuantTitCobS = FormataCpoNum(0, 6)
                    .sValTotalTitCartS = FormataCpoNum(0, 17)
                
                End If
                
                .sQuantTitCobV = FormataCpoNum(0, 6)
                .sValTotalTitCartV = FormataCpoNum(0, 17)
                
                .sQuantTitCobC = FormataCpoNum(0, 6)
                .sValTotalTitCartC = FormataCpoNum(0, 17)
                
                .sQuantTitCobD = FormataCpoTexto("", 6)
                .sValTotalTitCartD = FormataCpoTexto("", 17)
            
            Case Else
                .sQuantTitCobS = FormataCpoNum(0, 6)
                .sValTotalTitCartS = FormataCpoNum(0, 17)
                .sQuantTitCobV = FormataCpoNum(0, 6)
                .sValTotalTitCartV = FormataCpoNum(0, 17)
                .sQuantTitCobC = FormataCpoNum(0, 6)
                .sValTotalTitCartC = FormataCpoNum(0, 17)
                .sQuantTitCobD = FormataCpoNum(0, 6)
                .sValTotalTitCartD = FormataCpoNum(0, 17)
        
        End Select
        
        .sNumAvisoLancamento = FormataCpoTexto("", 8)
        .sBrancosCNAB2 = FormataCpoTexto("", 117)
    End With
    
    With tRegTrailerCNAB240
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147082)

    End Select

    Exit Function

End Function

Function Processar_ArquivoRetorno_Cobranca(objCobrancaEletronica As ClassCobrancaEletronica) As Long
'Faz o processamento do arquivo de retorno

Dim lErro As Long
Dim iIndice As Integer
Dim sRegistro As String
Dim iCodRegistro As Integer, objLoteServico As New ClassCNABLoteServico
Dim objBorderoRetCobr As New ClassBorderoRetCobr
Dim lComando As Long, tRegDetalhe As typeRetCobrDetalheCNAB240
Dim tRegControle As typeRegControle

On Error GoTo Erro_Processar_ArquivoRetorno_Cobranca

    Call Zera_Var

    Set gobjCobrancaEletronica = objCobrancaEletronica
   
    objBorderoRetCobr.iBanco = objCobrancaEletronica.objCobrador.iCodBanco
    
    'le as carteiras de cobranca do cobrador
    lErro = CF("CarteirasCobrador_Le", objCobrancaEletronica.objCobrador, objBorderoRetCobr.colCartCobr)
    If lErro <> SUCESSO Then Error 32265
    
    'Abre os comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 51824

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
                    lErro = Processa_Registro_Detalhe(objLoteServico, sRegistro, objBorderoRetCobr, objCobrancaEletronica, tRegDetalhe, lComando)
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
    
    'Grava os dados recolhidos durante o processamento do arquivo de retorno
    lErro = CF("BorderoRetCobr_ConfirmaGrava", objBorderoRetCobr)
    If lErro <> SUCESSO Then Error 51829
    
    Processar_ArquivoRetorno_Cobranca = SUCESSO
    
    Exit Function
    
Erro_Processar_ArquivoRetorno_Cobranca:

    Processar_ArquivoRetorno_Cobranca = Err
    
    Select Case Err
    
        Case 51823
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 51824, 51825
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 32260, 32265, 51826, 51827, 51828, 51829
        
        Case 51831
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORMATO_ARQUIVO_INCORRETO", Err, objCobrancaEletronica.sNomeArquivoRetorno)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147083)
    
    End Select

    Close #2
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Private Function Processa_Registro_Header_Arq(sRegistro As String, objBorderoRetCobr As ClassBorderoRetCobr, objCobrador As ClassCobrador, sNomeArquivo As String) As Long
'para cobranca
Dim lErro As Long, lComando As Long, lNumSeqArquivo As Long, iCobrador As Integer
Dim sNomeArq As String, dtDataRecepcao As Date
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Processa_Registro_Header_Arq

    'Abre os comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 130457

    'Preenche o objBorderoRetCobr
    objBorderoRetCobr.dtDataRecepcao = gdtDataHoje
    objBorderoRetCobr.iCobrador = objCobrador.iCodigo
    objBorderoRetCobr.sNomeArq = sNomeArquivo
        
    lNumSeqArquivo = StrParaLong(Mid(sRegistro, 158, 6))
    objBorderoRetCobr.lNumBorderoRet = lNumSeqArquivo
    
    'verificar se o bco é o do cobrador
    If gobjCobrancaEletronica.objCobrador.iCodBanco <> CInt(left(sRegistro, 3)) Then gError 130011
    
    'se o arquivo já foi processado antes, entao dar erro
    sNomeArq = String(STRING_NOME_ARQ_COMPLETO, 0)
    lErro = Comando_Executar(lComando, "SELECT Cobrador, NomeArq, DataRecepcao FROM BorderosRetCobr WHERE BorderosRetCobr.Cobrador = ? AND NumBorderoRet = ? ORDER BY DataRecepcao DESC", _
        iCobrador, sNomeArq, dtDataRecepcao, gobjCobrancaEletronica.objCobrador.iCodigo, lNumSeqArquivo)
    If lErro <> AD_SQL_SUCESSO Then gError 130458

    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130459
    If lErro <> AD_SQL_SEM_DADOS Then
        vbMsgRes = Rotina_Aviso(vbYesNo, "ERRO_BORDERORETCOBR_PROCESSADO", lNumSeqArquivo, gobjCobrancaEletronica.objCobrador.iCodBanco, dtDataRecepcao, sNomeArq, objCobrador.iCodigo)
        If vbMsgRes <> vbYes Then gError 130460
    End If
        
    Call Comando_Fechar(lComando)
    
    gobjCobrancaEletronica.sVersaoLayout = Mid(sRegistro, 164, 3)
    
    Processa_Registro_Header_Arq = SUCESSO
    
    Exit Function

Erro_Processa_Registro_Header_Arq:

    Processa_Registro_Header_Arq = gErr

    Select Case gErr
    
        Case 130011
            Call Rotina_Erro(vbOKOnly, "ERRO_BANCO_ARQRET_NAO_CONFERE", gErr)
            
        Case 130457
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 130458, 130459
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BORDEROSRETCOBR", gErr)
            
        Case 130460
            'Call Rotina_Erro(vbOKOnly, "ERRO_BORDERORETCOBR_PROCESSADO", gErr, lNumSeqArquivo, gobjCobrancaEletronica.objCobrador.iCodBanco, dtDataRecepcao, sNomeArq, objCobrador.iCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147084)
            
    End Select
    
    Call Comando_Fechar(lComando)
    
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

Private Function Processa_Registro_Detalhe(objLoteServico As ClassCNABLoteServico, sRegistro As String, objBorderoRetCobr As ClassBorderoRetCobr, objCobrancaEletronica As ClassCobrancaEletronica, tRegDetalhe As typeRetCobrDetalheCNAB240, lComando As Long) As Long
'para cobranca
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
                    Case 6, 8, 17
                        If iCodOcorrencia = 8 Then
                            If gobjCobrancaEletronica.objCobrador.iCodBanco <> 341 Then
                                iTipoOcorrencia = RETCOBR_DET_IGNORAR
                            Else
                                iTipoOcorrencia = RETCOBR_DET_BAIXA 'No Itaú a Baixa em cartório é o iCodOcorrencia 8, enquanto nos outros é a iCodOcorrencia 6 com Detalhe 1 = 8
                            End If
                        Else
                            iTipoOcorrencia = RETCOBR_DET_BAIXA
                        End If
                        
                    'cobranca de tarifas e custas
                    Case 2, 9, 12, 19, 20, 23, 28, 14
                        iTipoOcorrencia = RETCOBR_DET_TARIFAS
                                            
                    Case 3
                        iTipoOcorrencia = RETCOBR_DET_REJEITADO
                
                    Case 25 'baixa por ter sido protestado
                        iTipoOcorrencia = RETCOBR_DET_BAIXA_POR_PROTESTO
                    
                    Case Else
                        iTipoOcorrencia = RETCOBR_DET_IGNORAR
                        
                End Select
                
                If iTipoOcorrencia <> RETCOBR_DET_IGNORAR Then
                
                    lErro = Prepara_CartCobrRet1(objCobrancaEletronica, objBorderoRetCobr, objBorderoRetCobrCart, objDetRetCobr, tRegDetalhe)
                    If lErro <> SUCESSO Then Error 32257
                    
                    lErro = CF("Processa_Registro_Detalhe_Geral", objBorderoRetCobr, objCobrancaEletronica, iTipoOcorrencia, RETCOBR_TIPO_SEU_NUMERO2, objDetRetCobr, objBorderoRetCobrCart, lComando)
                    If lErro <> SUCESSO Then Error 32356
                                              
                End If
        
        End Select
    
    End If
    
    Processa_Registro_Detalhe = SUCESSO
    
    Exit Function

Erro_Processa_Registro_Detalhe:

    Processa_Registro_Detalhe = Err

    Select Case Err
    
        Case 32356, 32257, 51835, 51836, 51837, 51838
    
        Case 51839
            lErro = Rotina_Erro(vbOKOnly, "ERRO_OCORRREMPARCREC_NAO_CADASTRADA", Err, objOcorrRemParcRec.lNumIntDoc)
            
        Case 51840
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PARCELAREC_NAO_CADASTRADA", Err, objOcorrRemParcRec.lNumIntParc)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147085)
            
    End Select
    
    Exit Function

End Function

Private Function Extrair_Detalhe_Retorno(tRegDetalheBcoBrasil As typeRetCobrDetalheCNAB240, sRegistro As String) As Long
'Extrai do sRegistro os dados do Detalhe do arquido de retorno
'para cobranca
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
                
                '??? com variacao de acordo com o banco
                Select Case gobjCobrancaEletronica.objCobrador.iCodBanco
                
                    Case 1
                        .sCodCarteira = Mid(sRegistro, 58, 1)
                        .sNossoNumero = Mid(sRegistro, 38, 20)
                        
                    Case 8, 33, 353
                        .sCodCarteira = Mid(sRegistro, 54, 1)
                        .sNossoNumero = Mid(sRegistro, 41, 12) 'sem o dv
                        
                        .sAgencia = Mid(sRegistro, 18, 4)
                        .sDVAgencia = Mid(sRegistro, 22, 1)
                        .sContaCorrente = Mid(sRegistro, 23, 9)
                        .sDVContaCorrente = Mid(sRegistro, 32, 1)
                        .sDVAgenciaConta = "0"
                        .sNumDoc = Mid(sRegistro, 55, 15)
                        .sVencimento = Mid(sRegistro, 70, 8)
                        .sValorTitulo = Mid(sRegistro, 78, 15)
                        .sBcoCobrRec = Mid(sRegistro, 93, 3)
                        .sAgCobrRec = Mid(sRegistro, 96, 4)
                        .sDVAgCobrRec = Mid(sRegistro, 100, 1)
                        .sUsoEmpresa = Mid(sRegistro, 101, 25)
                        .sCodMoeda = Mid(sRegistro, 126, 2)
                        .sSacadoTipoInscr = Mid(sRegistro, 128, 1)
                        .sSacadoNumInscr = Mid(sRegistro, 129, 15)
                        .sSacadoNome = Mid(sRegistro, 144, 40)
                        .sNumContrato = Mid(sRegistro, 184, 10)
                        .sValorTarifa = Mid(sRegistro, 194, 15)
                        .sMotivoOcorrencia = Mid(sRegistro, 209, 10)
                        .sReservadoCNAB2 = Mid(sRegistro, 219, 22)
                        
                    Case 104
                        .sCodCarteira = Mid(sRegistro, 58, 1)
                        If gobjCobrancaEletronica.sVersaoLayout = "050" Then
                            .sNossoNumero = Mid(sRegistro, 42, 15)
                        Else 'versao 030 do arquivo
                            .sNossoNumero = Mid(sRegistro, 39, 18)
                        End If
                        
                    Case 479
                        .sCodCarteira = Mid(sRegistro, 58, 1)
                        .sNossoNumero = Mid(sRegistro, 38, 9)
                        
                    Case 275, 356
                        .sCodCarteira = Mid(sRegistro, 58, 1)
                        .sNossoNumero = Mid(sRegistro, 45, 13) 'os ultimos 7 sao o nosso numero. nao tem dv
                        
                    Case 409
                        .sCodCarteira = Mid(sRegistro, 58, 1)
                        .sNossoNumero = Mid(sRegistro, 38, 20)
                        
                    Case 399
                        .sCodCarteira = Mid(sRegistro, 58, 1)
                        If Len(Trim(Mid(sRegistro, 38, 20))) = 16 Then
                            .sNossoNumero = Mid(sRegistro, 38, 13) 'cob nao registrada sem os 3 dvs
                        Else
                            .sNossoNumero = Mid(sRegistro, 38, 10) 'cob registrada sem o dv
                        End If
                
                    Case 341
                        .sCodCarteira = Mid(sRegistro, 38, 3)
                        .sNossoNumero = Mid(sRegistro, 41, 8) 'sem o dv
                        
                    Case 745
                        .sCodCarteira = Mid(sRegistro, 58, 1)
                        .sNossoNumero = Mid(sRegistro, 38, 20)
                    
                    Case 453, 749 'Rural
                        .sCodCarteira = Mid(sRegistro, 58, 1)
                        .sNossoNumero = Mid(sRegistro, 38, 7)
                        
''                    Case 246 'Banco ABC Brasil
''                        .sCodCarteira = Mid(sRegistro, 58, 1)
''                        .sNossoNumero = Mid(sRegistro, 47, 11)
                    
''                    Case 41 'Banrisul
''                        .sCodCarteira = Mid(sRegistro, 58, 1)
''                        .sNossoNumero = Mid(sRegistro, 38, 20)
                    
                    Case 85
                        .sCodCarteira = Mid(sRegistro, 58, 1)
                        .sNossoNumero = Trim(Mid(sRegistro, 38, 20))
                    
                    Case Else
                        Error 28000
                    
                End Select
            
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

    Extrair_Detalhe_Retorno = Err
    
    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147086)
            
    End Select
    
    Exit Function

End Function

Function Processa_Registro_Trailer(sRegistroTrailer As String) As Long

Dim lErro As Long
Dim tRegTrailerBcoBrasil As typeRetTrailerCNAB240

On Error GoTo Erro_Processa_Registro_Trailer

    'Extrai da variável lida no arquivo os dados do registro de detalhe
    lErro = Extrair_Trailer_Retorno(tRegTrailerBcoBrasil, sRegistroTrailer)
    If lErro <> SUCESSO Then Error 51841
    
    '??? poderia verificar a qtde de registros
    
    Processa_Registro_Trailer = SUCESSO
    
    Exit Function
    
Erro_Processa_Registro_Trailer:

    Processa_Registro_Trailer = Err
    
    Select Case Err
    
        Case 51841
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147087)
            
    End Select
    
    Exit Function
    
End Function

Private Function Extrair_Trailer_Retorno(tRegTrailerBcoBrasil As typeRetTrailerCNAB240, sRegistro As String) As Long
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147088)
            
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
        Call Calcula_DV11(sNossoNumero, 9, sDVNossoNumero, True)
        
        '• MÉTODOLOGIA
        'a) O módulo 11 é utilizado para montagem dos campos nosso-número, código do cedente e
        'prefixo da agência, que serão impressos nos respectivos campos do bloqueto de cobrança;
        'b) Cada algarismo que compõe o nosso-número é multiplicado pelo respectivo multiplicador
        '(peso).
        'I - os multiplicadores (pesos) variam de 9 a 2, da direita para a esquerda;
        'II - o primeiro dígito da direita para a esquerda deverá ser multiplicado por 9, o
        'segundo por 8 e assim sucessivamente;
        'c) Os resultados das multiplicacoes são somados;
        'd) O total da soma é dívido por onze;
        'I - se o resto for menor que 10 (dez) o DV será igual ao resto;
'------>'II - se o resto for igual a 10 /dez/ o DV sera igual a X;
        'III - se o resto for igual a 0 /zero/ o DV sera igual a 0;
        'IV - no exemplo acima o digito verificador sera igual a 3;
        
        If sDVNossoNumero = "P" Then sDVNossoNumero = "X"
    End If

    Calcula_NossoNumero_BcoBrasil = SUCESSO
    
    Exit Function
    
Erro_Calcula_NossoNumero_BcoBrasil:

    Calcula_NossoNumero_BcoBrasil = Err
    
    Select Case Err
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, Err, 147089)
    
    End Select

    Exit Function

End Function

'Private Sub Calcula_DV_NossoNumero(sNossoNumero As String, sDVNossoNumero As String, iModulo As Integer, iBase As Integer)
'
'Dim iIndice As Integer
'Dim iTamNossoNum As Integer
'Dim iDigito1 As Integer
'Dim iVarModulo As Integer
'Dim iProduto As Integer
'Dim iSoma As Integer
'Dim iResto As Integer
'Dim iResultado As Integer
'
'    iVarModulo = 2
'    iSoma = 0
'    iTamNossoNum = Len(sNossoNumero)
'
'    For iIndice = iTamNossoNum To 1 Step -1
'
'        iDigito1 = StrParaInt(Mid(sNossoNumero, iIndice, 1))
'        iProduto = (iDigito1 * iVarModulo)
'
'        iSoma = iSoma + iProduto
'
'        iVarModulo = iVarModulo + 1
'
'        If iVarModulo > iModulo Then iVarModulo = 2
'
'    Next
'
'    iResto = iSoma Mod 11
'
'    iResultado = 11 - iResto
'
'    If iResultado >= 10 Then iResultado = iResultado - 10
'
'    sDVNossoNumero = FormataCpoNum(iResultado, 1)
'
'    Exit Sub
'
'End Sub

Sub RegArqCNAB_Extrai_Controle(sRegistro As String, tRegControle As typeRegControle, objCobrancaEletronica As ClassCobrancaEletronica)

    With tRegControle
    
        .sNumeroBanco = Mid(sRegistro, 1, 3)
        .sLoteServico = Mid(sRegistro, 4, 4)
        .sIdentificacaoReg = Mid(sRegistro, 8, 1)
        
    End With
    
End Sub

Function Processa_Registro_Header_Lote(objLoteServico As ClassCNABLoteServico, sRegistro As String, objBorderoRetCobr As ClassBorderoRetCobr, objCobrador As ClassCobrador) As Long

Dim lErro As Long
Dim objContaCorrente As New ClassContasCorrentesInternas
Dim vbMsgRes As VbMsgBoxResult

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
        
        With objLoteServico
        
            If Mid(sRegistro, 1, 3) = "033" Or Mid(sRegistro, 1, 3) = "003" Or Mid(sRegistro, 1, 3) = "353" Then
                .sAgencia = Mid(sRegistro, 54, 4)
                .sDVAgencia = Mid(sRegistro, 58, 1)
                .sNumCtaCorrente = Mid(sRegistro, 59, 9)
                .sDVCtaCorrente = Mid(sRegistro, 68, 1)
            ElseIf Mid(sRegistro, 1, 3) = "453" Or Mid(sRegistro, 1, 3) = "749" Then
                .sAgencia = Mid(sRegistro, 54, 5)
                .sDVAgencia = Mid(sRegistro, 59, 1)
                .sNumCtaCorrente = Mid(sRegistro, 63, 10)
                .sDVCtaCorrente = ""
            Else
                .sAgencia = Mid(sRegistro, 54, 5)
                .sDVAgencia = Mid(sRegistro, 59, 1)
                .sNumCtaCorrente = Mid(sRegistro, 60, 12)
                .sDVCtaCorrente = Mid(sRegistro, 72, 1)
            End If
        End With
    
        'Pega a conta corrente do cobrador
        objContaCorrente.iCodigo = objCobrador.iCodCCI
        
        'Lê a conta corrente do cobrador
        lErro = CF("ContaCorrenteInt_Le", objContaCorrente.iCodigo, objContaCorrente)
        If lErro <> SUCESSO And lErro <> 11807 Then gError 184420
        
        'Se não achou a Conta Corrente --> erro
        If lErro <> SUCESSO Then gError 184421
    
        'ABC não passa a conta corrente e sim o código do contrato
        'CEF usa uma conta de cobranca diferente da conta movimentada
        If objContaCorrente.iCodBanco <> 246 And objContaCorrente.iCodBanco <> 104 Then
            If StrParaLong(objLoteServico.sNumCtaCorrente) <> 0 And StrParaLong(objLoteServico.sNumCtaCorrente) <> StrParaLong(objContaCorrente.sNumConta) Then
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_RETCOBR_CONTA_OUTRO_COBRADOR", StrParaLong(objLoteServico.sNumCtaCorrente), StrParaLong(objContaCorrente.sNumConta))
                If vbMsgRes <> vbYes Then gError 184422
            End If
        End If
            
    End If
        
    Processa_Registro_Header_Lote = SUCESSO
     
    Exit Function
    
Erro_Processa_Registro_Header_Lote:

    Processa_Registro_Header_Lote = gErr
     
    Select Case gErr
          
        Case 184420, 184421, 184422
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147090)
     
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147091)
     
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147092)
     
    End Select
     
    Exit Function

End Function

Private Function Prepara_CartCobrRet1(objCobrancaEletronica As ClassCobrancaEletronica, objBorderoRetCobr As ClassBorderoRetCobr, objBorderoRetCobrCart As ClassBorderoRetCobrCart, objDetRetCobr As ClassDetRetCobr, tRegDetalhe As typeRetCobrDetalheCNAB240) As Long
'codigo comum ao tratamento de registros detalhe de baixas, custas e tarifas
'preenche objBorderoRetCobrCart e objDetRetCobr

Dim lErro As Long
Dim dtDataBaixaAjustada As Date

On Error GoTo Erro_Prepara_CartCobrRet1

    Call RegDetCobr_ObterInfo(objBorderoRetCobr, objDetRetCobr, tRegDetalhe)
    
    objDetRetCobr.dtDataBaixa = Extrair_Data(tRegDetalhe.tU030.sDataOcorrencia)
    
    lErro = RetCobr_Obtem_DataBaixa(objCobrancaEletronica, objBorderoRetCobr, objBorderoRetCobrCart, objDetRetCobr, dtDataBaixaAjustada)
    If lErro <> SUCESSO Then Error 32255
    
    objDetRetCobr.dtDataBaixa = dtDataBaixaAjustada
    
    lErro = CF("Prepara_CartCobrRet", objCobrancaEletronica, objBorderoRetCobr, objBorderoRetCobrCart, objDetRetCobr)
    If lErro <> SUCESSO Then Error 32255
    
    Prepara_CartCobrRet1 = SUCESSO
    
    Exit Function
    
Erro_Prepara_CartCobrRet1:

    Prepara_CartCobrRet1 = Err
    
    Select Case Err
    
        Case 32255
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147093)
            
    End Select

    Exit Function

End Function

Private Function RetCobr_Obtem_DataBaixa(ByVal objCobrancaEletronica As ClassCobrancaEletronica, ByVal objBorderoRetCobr As ClassBorderoRetCobr, ByVal objBorderoRetCobrCart As ClassBorderoRetCobrCart, ByVal objDetRetCobr As ClassDetRetCobr, dtDataBaixaAjustada As Date) As Long

Dim lErro As Long
Dim dtDataAux As Date
Dim dtDataAux2 As Date
Dim iDiasAjusteBaixa As Integer

On Error GoTo Erro_RetCobr_Obtem_DataBaixa

    If objCobrancaEletronica.objCobradorCNABInfo.colInformacoes.Count = 0 Then
        lErro = CF("CobradorInfo_Le", objCobrancaEletronica.objCobrador.iCodigo, objCobrancaEletronica.objCobradorCNABInfo.colInformacoes)
        If lErro <> SUCESSO Then gError 51681
    End If

    iDiasAjusteBaixa = StrParaInt(Obter_Info_Cobrador(objCobrancaEletronica.objCobradorCNABInfo.colInformacoes, 99))
    
    If iDiasAjusteBaixa <> 0 Then
    
        dtDataAux = objDetRetCobr.dtDataBaixa
        
        'Antecipa iDiasAjusteBaixa dias úteis
        Do While iDiasAjusteBaixa > 0
            
            dtDataAux = DateAdd("d", -1, dtDataAux)
        
            lErro = CF("DataVencto_Real", dtDataAux, dtDataAux2)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                        
            If dtDataAux = dtDataAux2 Then iDiasAjusteBaixa = iDiasAjusteBaixa - 1
        
        Loop
        
        dtDataBaixaAjustada = dtDataAux
        
    Else
    
        dtDataBaixaAjustada = objDetRetCobr.dtDataBaixa
    
    End If
    
    objBorderoRetCobr.iQuebraMovCCPorStatusRecurso = StrParaInt(Obter_Info_Cobrador(objCobrancaEletronica.objCobradorCNABInfo.colInformacoes, 37))
    
    objBorderoRetCobr.iVerificaCreditosTarifas = StrParaInt(Obter_Info_Cobrador(objCobrancaEletronica.objCobradorCNABInfo.colInformacoes, 38))
    
    RetCobr_Obtem_DataBaixa = SUCESSO
    
    Exit Function
    
Erro_RetCobr_Obtem_DataBaixa:

    RetCobr_Obtem_DataBaixa = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147093)
            
    End Select

    Exit Function
    
End Function

Private Sub RegDetCobr_ObterInfo(ByVal objBorderoRetCobr As ClassBorderoRetCobr, objDetRetCobr As ClassDetRetCobr, tRegDetalhe As typeRetCobrDetalheCNAB240)
'copia dados do type para o objeto
Dim objCarteiraCobrador As ClassCarteiraCobrador

    'Recolhe alguns dados do registro de detalhe
    With objDetRetCobr
            
        .dValorRecebido = Extrair_Valor(tRegDetalhe.tU030.sValorPagoSacado)
        .dValorCreditado = Extrair_Valor(tRegDetalhe.tU030.sValorCreditado)
        
        .dtDataVencimento = Extrair_Data(tRegDetalhe.tT030.sVencimento)
        .dtDataCredito = Extrair_Data(tRegDetalhe.tU030.sDataCredito)
        If .dtDataCredito = DATA_NULA Then
            .dtDataCredito = Extrair_Data(tRegDetalhe.tU030.sDataOcorrencia)
        End If
        .dValorJuros = Arredonda_Moeda(Extrair_Valor(tRegDetalhe.tU030.sJurosMulta) + Extrair_Valor(tRegDetalhe.tU030.sOutrosCreditos))
        .dValorDesconto = Extrair_Valor(tRegDetalhe.tU030.sDesconto)
        .sNossoNumero = Trim(tRegDetalhe.tT030.sNossoNumero)
        .iCodOcorrencia = StrParaInt(tRegDetalhe.tT030.sCodMovto)
        
        If tRegDetalhe.tT030.sCodCarteira = "F" And (tRegDetalhe.tT030.sNumeroBanco = "453" Or tRegDetalhe.tT030.sNumeroBanco = "749") Then
        
            For Each objCarteiraCobrador In objBorderoRetCobr.colCartCobr
                If objCarteiraCobrador.iCodCarteiraCobranca = CARTEIRA_DESCONTADA Then
                    .iCodCarteiraNoBanco = objCarteiraCobrador.iNumCarteiraNoBanco
                    Exit For
                End If
            Next
        
        Else
        
            .iCodCarteiraNoBanco = StrParaInt(tRegDetalhe.tT030.sCodCarteira)
        
        End If
        
        .dValorTarifa = Extrair_Valor(tRegDetalhe.tT030.sValorTarifa) + Extrair_Valor(tRegDetalhe.tU030.sOutrasDespesas)
        .sSeuNumero = tRegDetalhe.tT030.sNumDoc
        .sUsoDaEmpresa = tRegDetalhe.tT030.sUsoEmpresa
        .lNumSequencial = CLng(tRegDetalhe.tT030.sNumSeqRegLote)
        .dValorTitulo = Extrair_Valor(tRegDetalhe.tT030.sValorTitulo)
        
'        .iCodOcorrencia1 = StrParaInt(Replace(Mid(tRegDetalhe.tT030.sMotivoOcorrencia, 1, 2), "A", "10"))
'        .iCodOcorrencia2 = StrParaInt(Replace(Mid(tRegDetalhe.tT030.sMotivoOcorrencia, 3, 2), "A", "10"))
'        .iCodOcorrencia3 = StrParaInt(Replace(Mid(tRegDetalhe.tT030.sMotivoOcorrencia, 5, 2), "A", "10"))
'        .iCodOcorrencia4 = StrParaInt(Replace(Mid(tRegDetalhe.tT030.sMotivoOcorrencia, 7, 2), "A", "10"))
        .iCodOcorrencia1 = RetCobr_Extrai_CodOcor(Mid(tRegDetalhe.tT030.sMotivoOcorrencia, 1, 2))
        .iCodOcorrencia2 = RetCobr_Extrai_CodOcor(Mid(tRegDetalhe.tT030.sMotivoOcorrencia, 3, 2))
        .iCodOcorrencia3 = RetCobr_Extrai_CodOcor(Mid(tRegDetalhe.tT030.sMotivoOcorrencia, 5, 2))
        .iCodOcorrencia4 = RetCobr_Extrai_CodOcor(Mid(tRegDetalhe.tT030.sMotivoOcorrencia, 7, 2))
        
        If tRegDetalhe.tT030.sNumeroBanco <> "341" Then
            '.iCodOcorrencia5 = StrParaInt(Replace(Mid(tRegDetalhe.tT030.sMotivoOcorrencia, 9, 2), "A", "10"))
            .iCodOcorrencia5 = RetCobr_Extrai_CodOcor(Mid(tRegDetalhe.tT030.sMotivoOcorrencia, 9, 2))
        Else
            .sCodLiquidacao = Mid(tRegDetalhe.tT030.sMotivoOcorrencia, 9, 2)
        End If
                
    End With
    
End Sub

Private Function RetCobr_Extrai_CodOcor(ByVal sValor As String) As Integer

Dim sValorAux As String

    sValorAux = ""
    
    If Len(Trim(sValor)) > 0 Then

        sValorAux = Replace(sValor, "A", "10")
        
        If Not IsNumeric(sValorAux) Then sValorAux = "" 'Paliativo para não dar erro, existem vários valores não numéricos possíveis e não só "A"
    
    End If
            
    RetCobr_Extrai_CodOcor = StrParaInt(sValorAux)
    
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147094)
     
    End Select
     
    Close #1
    
    Exit Function

End Function

Private Function Obter_Dados_DetalheSegP(tRegDetalheCNAB240SegP As typeRemDetalheCNAB240SegP, objOcorrRemParcRem As ClassOcorrRemParcRec, objCarteiraCobrador As ClassCarteiraCobrador, objParcelaRec As ClassParcelaReceber) As Long
'Obtem os dados que são utilizados para preencher o registro de detalhe

Dim lErro As Long
Dim sNossoNumero As String, sCampoLivre As String
Dim dValorDesconto As Double
Dim sDVNossoNumero As String, sCodConvBanco As String, sDVAgCta As String, sDVConta As String
Dim sTipoDoc As String, sTipoMora As String

On Error GoTo Erro_Obter_Dados_DetalheSegP

With tRegDetalheCNAB240SegP
        
    .sCodBancoComp = FormataCpoNum(gobjCobrancaEletronica.objCobrador.iCodBanco, 3)
    .sLoteServico = FormataCpoNum(giNumLotes, 4)
    .sRegDetalhe = "3"
    .sNumSequencialReg = FormataCpoNum(gobjCobrancaEletronica.lSequencialRegistro, 5)
    .sCodSegRegDetalhe = "P"
    .sBrancosCNAB1 = FormataCpoTexto("", 1)
    Select Case objOcorrRemParcRem.iCodOcorrencia
        Case 1
            .sCodMovimento = "01"
        Case 2
            .sCodMovimento = "02"
        Case 6
            .sCodMovimento = "06"
            'obter valor cobrado, instrucao1, dias de protesto1, instrucao2, dias de protesto2 da remessa anterior do titulo
            lErro = InstrucaoRem_ObterInfo(objOcorrRemParcRem)
            If lErro <> SUCESSO And lErro <> 184164 Then gError 51822
            
            If lErro = 184164 Then gError 202346
            
        Case Else
            gError 51821
    End Select
    
    .sAgMantConta = FormataCpoNum(gobjCobrancaEletronica.objCCI.sAgencia, 5)
    
    '.sIdentTituloBanco Calculado depois.
    .sFormaCadTituloBanco = FormataCpoNum(IIf(objCarteiraCobrador.iComRegistro = 0, "2", "1"), 1)
    
    sTipoDoc = Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 11)

    If Len(Trim(sTipoDoc)) = 0 Then
        .sTipoDocumento = "2"
    Else
        .sTipoDocumento = left(sTipoDoc, 1)
    End If
    
    .sCodBaixaDevolucao = "2"
    .sNumDiasBaixaDevol = FormataCpoNum(0, 3)
    
    '??? com variacao de acordo com o banco
    Select Case gobjCobrancaEletronica.objCobrador.iCodBanco
        
        Case 1
'            .sCodBaixaDevolucao = "0"
            .sCodBaixaDevolucao = FormataCpoNum(Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 12), 1)
            .sCodMoeda = "09"
            .sCodJurosMora = "1"
            .sEspecieTitulo = FormataCpoNum(2, 2)
            .sDVAgMantConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVAgencia, 1)
            .sDVConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVNumConta, 1)
            .sDVAgenciaConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVAgConta, 1)
            .sDVAgencia = " "
            .sIdentEmissaoBloq = FormataCpoNum(IIf(objCarteiraCobrador.iImprimeBoleta = 0 And objParcelaRec.lIdImpressaoBoleto = 0, "1", "2"), 1)
            .sIdentDistrib = .sIdentEmissaoBloq '"1"
            Select Case objCarteiraCobrador.iNumCarteiraNoBanco
                Case 17
                    .sCodCarteira = "7"
                Case Else
                    Select Case objCarteiraCobrador.iCodCarteiraCobranca
                        Case CARTEIRA_SIMPLES
                            .sCodCarteira = "1"
                        Case CARTEIRA_VINCULADA
                            .sCodCarteira = "2"
                        Case CARTEIRA_CAUCIONADA
                            .sCodCarteira = "3"
                        Case CARTEIRA_DESCONTADA
                            .sCodCarteira = "4"
                    End Select
            End Select
            .sNumContaCorrente = FormataCpoNum(gobjCobrancaEletronica.objCCI.sNumConta, 12)
            'Está faltando o campo Varição e o DVNossoNumero
            'lErro = Calcula_NossoNumero_BcoBrasil(objCarteiraCobrador, sNossoNumero, sDVNossoNumero, .sVariacao, .sEspecieTitulo, gsNumConvenio)
            
            'A parcela já tem que estar com o nosso número certo para emissão do boleto
            'sCodConvBanco = FormataCpoTexto(Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 7), 20)
            '.sIdentTituloBanco = FormataCpoNum(left(sCodConvBanco, 7), 7) & FormataCpoNum(right(objParcelaRec.sNumTitCobrador, 10), 10) & "   "
    
    
            '1. Caso seja o Banco do Brasil quem gera o "Nosso Número", informar 'brancos' (espaços) ou Zeros.
            '2. Caso seja a empresa quem gera o "Nosso Número", informar de acordo com o número do convênio de
            'cobrança conforme as seguintes regras:
            'a) Para número de convênio de 4 posições (carteira 17 até 9.999) informar o nosso número com
            '11 posições mais o DV (dígito verificador), sendo as 4 primeiras posições o número do convênio,
            'as 7 posições seguintes um número sequencial para controle e mais o DV.
            'Exemplo: CONVÊNIOS DE 0001 ATÉ 9.999
            '123412345671
            'CCCCSSSSSSSD
            'Onde: C = Convênio S = Sequencial D = dígito verificador
            'b) Para número de convênio de 6 posições (de 10.000 a 999.999), informar o nosso número com
            '11 posições mais o DV, sendo as 6 primeiras posições o número do convênio, as 5 posições
            'seguintes um número sequencial para controle e mais o DV.
            'Exemplo: CONVÊNIOS DE 010.000 ATÉ 999.999
            '123456123451
            'CCCCCCSSSSSD
            'Onde: C = Convênio S = Sequencial D = dígito verificador
            'c) Para número de convênio de 7 posições (acima de 1.000.000) informar o nosso número com
            '17 posições sem DV, sendo as 7 primeiras posições o número do convênio e as 10 posições seguintes
            'um número sequencial para controle. Esse é o padrão mais utilizado atualmente.
            'Exemplo: CONVÊNIOS DE 1.000.000 ATÉ 9.999.999
            '12345671234567890
            'CCCCCCCSSSSSSSSSS
            'Onde: C = Convênio S = Sequencial
            'Importante:
            'todos os "nosso número" devem ser alinhados à esquerda com brancos à direita.
            If Len(Trim(objParcelaRec.sNumTitCobrador)) = 0 Then
            'o Banco do Brasil quem gera o "Nosso Número", informar 'brancos' (espaços) ou Zeros.
                .sIdentTituloBanco = FormataCpoNum(objParcelaRec.sNumTitCobrador, 20)
            ElseIf Len(Trim(objParcelaRec.sNumTitCobrador)) = 12 Then
            'CCCCSSSSSSSD (convênio de 4 posições) ou CCCCCCSSSSSD (convênio de 6 posições)
                .sIdentTituloBanco = FormataCpoNum(objParcelaRec.sNumTitCobrador, 12) & FormataCpoTexto("", 8)
            Else
            'CCCCCCCSSSSSSSSSS (convênio de 7 posições)
                .sIdentTituloBanco = FormataCpoNum(objParcelaRec.sNumTitCobrador, 17) & "   "
            End If
    
            'Verifica qual será o Código do Protesto
            If objOcorrRemParcRem.iInstrucao1 = 81 Or objOcorrRemParcRem.iInstrucao1 = 82 Then
        
                .sCodProtesto = IIf(objOcorrRemParcRem.iInstrucao1 = 81, "1", "2")
                .sNumDiasProtesto = FormataCpoNum(objOcorrRemParcRem.iDiasDeProtesto1, 2)
        
            ElseIf objOcorrRemParcRem.iInstrucao2 = 81 Or objOcorrRemParcRem.iInstrucao2 = 82 Then
        
                .sCodProtesto = IIf(objOcorrRemParcRem.iInstrucao2 = 81, "1", "2")
                .sNumDiasProtesto = FormataCpoNum(objOcorrRemParcRem.iDiasDeProtesto2, 2)
            Else
                .sCodProtesto = FormataCpoTexto(Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 8), 1)
                .sNumDiasProtesto = FormataCpoNum(Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 9), 2)
            End If
        
            If .sCodProtesto = " " Then
            
                'se for cobranca caucionada => nao protesta
                If objCarteiraCobrador.iNumCarteiraNoBanco = 3 Then
                
                    .sCodProtesto = "3"
                    .sNumDiasProtesto = FormataCpoNum(0, 2)
                    
                Else
                
                    .sCodProtesto = "2" 'dias uteis
                    .sNumDiasProtesto = FormataCpoNum(5, 2)
                    
                End If
                
            End If
                       
            .sNumContratoOpCred = FormataCpoNum(Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 10), 10)

    
        Case 8, 33, 353
            .sCodMoeda = "00"
            .sCodJurosMora = "1"
            .sEspecieTitulo = " " & FormataCpoNum(2, 2)
            .sAgMantConta = FormataCpoNum(gobjCobrancaEletronica.objCCI.sAgencia, 4)
            .sDVAgMantConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVAgencia, 1)
            .sNumContaCorrente = FormataCpoNum(gobjCobrancaEletronica.objCCI.sNumConta, 9)
            .sDVConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVNumConta, 1)
            .sDVAgencia = ""
            .sAgEncCobranca = FormataCpoTexto("0", 4)
            .sCodCarteira = FormataCpoNum(IIf(Len(CStr(objCarteiraCobrador.iNumCarteiraNoBanco)) > 1, objCarteiraCobrador.sCodCarteiraNoBanco, objCarteiraCobrador.iNumCarteiraNoBanco), 1)
            .sIdentEmissaoBloq = " "
            .sIdentDistrib = " "
            .sCodProtesto = FormataCpoTexto(Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 8), 1)
            .sNumDiasProtesto = FormataCpoNum(Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 9), 2)
            
            If Len(Trim(objParcelaRec.sNumTitCobrador)) <> 0 Then
                If Len(Trim(objParcelaRec.sNumTitCobrador)) = 13 Then
                    .sIdentTituloBanco = Trim(objParcelaRec.sNumTitCobrador)
                Else
                    sNossoNumero = FormataCpoNum(right(objParcelaRec.sNumTitCobrador, 12), 12)
                    Call Calcula_DV11(sNossoNumero, 9, sDVNossoNumero)
                    .sIdentTituloBanco = sNossoNumero & sDVNossoNumero
                End If
            Else
                .sIdentTituloBanco = FormataCpoNum("", 13)
            End If
            '.sIdentTituloBanco = FormataCpoNum(objParcelaRec.sNumTitCobrador, 13)
            
            .sNumContratoOpCred = FormataCpoTexto("", 10)
            .sContaCobranca = FormataCpoNum(Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 15), 9)
            .sDVContaCobranca = FormataCpoNum(Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 16), 1)
   
        Case 104
            
            .sCodMoeda = "09"
            .sCodJurosMora = "1"
            .sEspecieTitulo = FormataCpoNum(2, 2)
                
            .sIdentEmissaoBloq = FormataCpoNum(IIf(objCarteiraCobrador.iImprimeBoleta = 0 And objParcelaRec.lIdImpressaoBoleto = 0, "1", "2"), 1)
            
            If gobjCobrancaEletronica.sVersaoLayout = "050" Then
                
                '.sDVAgencia = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVAgencia, 1)
                .sDVAgencia = FormataCpoTexto("0", 1)

                .sDVAgMantConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVAgencia, 1)
                .sNumContaCorrente = FormataCpoNum(Obter_Info_CarteiraCobrador(CARTEIRA_SIMPLES, gobjCobrancaEletronica.colCarteirasCNABInfo, 7), 6)
            
                If Len(Trim(objParcelaRec.sNumTitCobrador)) <> 0 Then
                    
                    sNossoNumero = FormataCpoNum(right(objParcelaRec.sNumTitCobrador, 10), 10)
                    Call Calcula_DV11(sNossoNumero, 9, sDVNossoNumero)
                    .sIdentTituloBanco = IIf(objCarteiraCobrador.iComRegistro <> 0, "14", "24") & FormataCpoNum(sNossoNumero, 14) & sDVNossoNumero
                
                Else
                    
                    .sIdentTituloBanco = FormataCpoNum("", 17)
                
                End If
                
                .sNumContratoOpCred = FormataCpoNum("", 10)
            
            Else 'versao 030 do arquivo
            
                .sCodMoeda = "09"
                .sCodJurosMora = "1"
                .sEspecieTitulo = FormataCpoNum(2, 2)
                .sDVAgencia = FormataCpoTexto("0", 1)
                sCodConvBanco = FormataCpoNum(Obter_Info_CarteiraCobrador(CARTEIRA_SIMPLES, gobjCobrancaEletronica.colCarteirasCNABInfo, 7), 16)
                .sAgMantConta = FormataCpoNum(Mid(sCodConvBanco, 1, 4), 5)
                Call Calcula_DV11(.sAgMantConta, 9, sDVConta)
                .sDVAgMantConta = sDVConta
    
                .sNumContaCorrente = FormataCpoNum(Mid(sCodConvBanco, 8, 8), 12) '???? tem que pegar do cedente
                'calcular dv da conta
                'Call Calcula_DV11(.sNumContaCorrente, 9, sDVConta)
                .sDVConta = "0" 'sDVConta
                
                'Call Calcula_DV_AgCta_104(left(sCodConvBanco, 4) & .sNumContaCorrente, sDVAgCta)
                Call Calcula_DV11(left(sCodConvBanco, 4) & .sNumContaCorrente, 9, sDVAgCta)
                .sDVAgenciaConta = sDVAgCta
                '.sNumContaCorrente = Mid(sCodConvBanco, 6, 1) & FormataCpoNum(Mid(sCodConvBanco, 8, 8), 11) '???? tem que pegar do cedente
                
                If Len(Trim(objParcelaRec.sNumTitCobrador)) <> 0 Then
                    sNossoNumero = FormataCpoNum(right(objParcelaRec.sNumTitCobrador, 10), 10)
                    Call Calcula_DV11(sNossoNumero, 9, sDVNossoNumero)
                    .sIdentTituloBanco = sNossoNumero & sDVNossoNumero
                Else
                    .sIdentTituloBanco = FormataCpoNum("", 11)
                End If
            
                .sNumContratoOpCred = FormataCpoTexto("", 10)

            End If
            
            .sIdentDistrib = .sIdentEmissaoBloq ''??????
            If Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 37) <> "" Then
                'C009 Identificação da Emissão do Bloqueto
                'Código adotado pela FEBRABAN para identificar o responsável e a forma de emissão do bloqueto.
                '‘1’ = Banco Emite
                '‘2’ = Cliente Emite ou para Bloqueto Pré-Impresso Registrado (entrega do bloqueto pelo Cedente)
                '‘4’ = Banco Reemite
                '‘5’ = Banco Não Reemite
                'Os códigos ‘4’ e ‘5’ só serão aceitos para código de movimento para remessa ‘31’
                'C009
                'C010 Identificação da Entrega / Distribuição do Bloqueto
                'Código adotado pela FEBRABAN para identificar o responsável pela distribuição do bloqueto.
                'Id Entrega do Bloqueto
                '‘0’ = Postagem pelo Cedente
                '‘1’ = Sacado via Correios
                '‘2’ = Cedente via Agência CAIXA
                .sIdentDistrib = FormataCpoNum(Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 37), 1)
            End If
            
            
            If objCarteiraCobrador.iCodCarteiraCobranca = CARTEIRA_SIMPLES Then
                .sCodCarteira = "1"
            Else
                If objCarteiraCobrador.iCodCarteiraCobranca = CARTEIRA_CAUCIONADA Then
                    .sCodCarteira = "3"
                Else
                    If objCarteiraCobrador.iCodCarteiraCobranca = CARTEIRA_DESCONTADA And gobjCobrancaEletronica.sVersaoLayout = "050" Then
                        .sCodCarteira = "4"
                    Else
                        .sCodCarteira = "1" '???? nao deveria acontecer
                    End If
                End If
            End If
            
            'Verifica qual será o Código do Protesto
            If objOcorrRemParcRem.iInstrucao1 = 81 Or objOcorrRemParcRem.iInstrucao1 = 82 Then
        
                .sCodProtesto = IIf(objOcorrRemParcRem.iInstrucao1 = 81, "1", "2")
                .sNumDiasProtesto = FormataCpoNum(objOcorrRemParcRem.iDiasDeProtesto1, 2)
        
            ElseIf objOcorrRemParcRem.iInstrucao2 = 81 Or objOcorrRemParcRem.iInstrucao2 = 82 Then
        
                .sCodProtesto = IIf(objOcorrRemParcRem.iInstrucao2 = 81, "1", "2")
                .sNumDiasProtesto = FormataCpoNum(objOcorrRemParcRem.iDiasDeProtesto2, 2)
            
            ElseIf objOcorrRemParcRem.iInstrucao1 = 99 Or objOcorrRemParcRem.iInstrucao2 = 99 Then
            
                .sCodProtesto = "3"
                .sNumDiasProtesto = FormataCpoNum(0, 2)
                
            Else
                .sCodProtesto = FormataCpoTexto(Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 8), 1)
                .sNumDiasProtesto = FormataCpoNum(Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 9), 2)
            End If
        
            '.sNumContratoOpCred = FormataCpoTexto("", 10)
            
            If Len(Trim(Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 13))) <> 0 Then
                .sCodBaixaDevolucao = left(Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 13), 1)
                .sNumDiasBaixaDevol = FormataCpoNum(Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 14), 3)
            End If
            
            If .sCodBaixaDevolucao = "1" Then 'baixar/devolver
                .sCodProtesto = "3"
                .sNumDiasProtesto = FormataCpoNum(0, 2)
            End If
        
        
        Case 479
            .sCodMoeda = "09"
            .sCodJurosMora = "1"
            .sEspecieTitulo = FormataCpoNum(2, 2)
            .sDVAgMantConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVAgencia, 1)
            .sDVConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVNumConta, 1)
            .sDVAgenciaConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVAgConta, 1)
            .sDVAgencia = FormataCpoTexto("0", 1)
            .sCodCarteira = FormataCpoNum(objCarteiraCobrador.iNumCarteiraNoBanco, 1)
            .sNumContaCorrente = FormataCpoNum(gobjCobrancaEletronica.objCCI.sNumConta, 8)
            .sIdentEmissaoBloq = FormataCpoNum(IIf(objCarteiraCobrador.iImprimeBoleta = 0 And objParcelaRec.lIdImpressaoBoleto = 0, "1", "2"), 1)
            If objCarteiraCobrador.iNumCarteiraNoBanco = 6 Then
                .sIdentDistrib = "2"
            Else
                .sIdentDistrib = "1"
            End If
            
            .sCodProtesto = FormataCpoTexto(Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 8), 1)
            .sNumDiasProtesto = FormataCpoNum(Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 9), 2)
            'Está faltando o campo Varição e o DVNossoNumero
            'lErro = Calcula_NossoNumero_BcoBrasil(objCarteiraCobrador, sNossoNumero, sDVNossoNumero, .sVariacao, .sEspecieTitulo, gsNumConvenio)
               
            .sIdentTituloBanco = FormataCpoNum(objParcelaRec.sNumTitCobrador, 9) '?????
            .sNumContratoOpCred = FormataCpoNum(Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 10), 10)
            
        Case 409
            .sCodMoeda = "09"
            .sCodJurosMora = "1"
            .sEspecieTitulo = FormataCpoNum(2, 2)
            .sDVAgMantConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVAgencia, 1)
            .sDVConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVNumConta, 1)
            .sDVAgenciaConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVAgConta, 1)
            .sDVAgencia = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVAgencia, 1)
            
            Select Case objCarteiraCobrador.iCodCarteiraCobranca
            
                Case CARTEIRA_SIMPLES
                    .sCodCarteira = "1"
                    
                Case CARTEIRA_VINCULADA
                    .sCodCarteira = "2"
                
                Case CARTEIRA_CAUCIONADA
                    .sCodCarteira = "3"
                
                Case CARTEIRA_DESCONTADA
                    .sCodCarteira = "4"
                
                '??? faltou tratar Vendor ("5")
                
                Case Else
                    .sCodCarteira = "1" '???? nao deveria acontecer
            
            End Select
            
            .sIdentEmissaoBloq = FormataCpoNum(IIf(objCarteiraCobrador.iImprimeBoleta = 0 And objParcelaRec.lIdImpressaoBoleto = 0, "1", "2"), 1)
            .sIdentDistrib = .sIdentEmissaoBloq
            .sNumContaCorrente = FormataCpoNum(gobjCobrancaEletronica.objCCI.sNumConta, 12)
            'Está faltando o campo Varição e o DVNossoNumero
            'lErro = Calcula_NossoNumero_BcoBrasil(objCarteiraCobrador, sNossoNumero, sDVNossoNumero, .sVariacao, .sEspecieTitulo, gsNumConvenio)
               
            sNossoNumero = FormataCpoNum(right(objParcelaRec.sNumTitCobrador, 10), 10)
            Call Calcula_DV11(sNossoNumero, 9, sDVNossoNumero)
            .sIdentTituloBanco = FormataCpoTexto(sNossoNumero & sDVNossoNumero, 20)
    
            'Verifica qual será o Código do Protesto
            If objOcorrRemParcRem.iInstrucao1 = 81 Or objOcorrRemParcRem.iInstrucao1 = 82 Then
        
                .sCodProtesto = IIf(objOcorrRemParcRem.iInstrucao1 = 81, "1", "2")
                .sNumDiasProtesto = FormataCpoNum(objOcorrRemParcRem.iDiasDeProtesto1, 2)
        
            ElseIf objOcorrRemParcRem.iInstrucao2 = 81 Or objOcorrRemParcRem.iInstrucao2 = 82 Then
        
                .sCodProtesto = IIf(objOcorrRemParcRem.iInstrucao2 = 81, "1", "2")
                .sNumDiasProtesto = FormataCpoNum(objOcorrRemParcRem.iDiasDeProtesto2, 2)
            
            ElseIf objOcorrRemParcRem.iInstrucao1 = 99 Or objOcorrRemParcRem.iInstrucao2 = 99 Then
            
                .sCodProtesto = "3"
                .sNumDiasProtesto = FormataCpoNum(0, 2)
                
            Else
                .sCodProtesto = FormataCpoTexto(Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 8), 1)
                .sNumDiasProtesto = FormataCpoNum(Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 9), 2)
            End If
        
            .sNumContratoOpCred = FormataCpoTexto("", 10)
        
        Case 275, 356
            .sCodMoeda = "09"
            sTipoMora = Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 12)
        
            If Len(Trim(sTipoMora)) = 0 Then
                .sCodJurosMora = "1"
            Else
                .sCodJurosMora = left(sTipoMora, 1)
            End If
            .sEspecieTitulo = FormataCpoNum(2, 2)
            .sDVAgMantConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVAgencia, 1)
            .sCodBaixaDevolucao = "0"
            .sNumDiasBaixaDevol = FormataCpoTexto("", 3)
            .sDVConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVNumConta, 1)
            .sDVAgenciaConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVAgConta, 1)
            .sDVAgencia = FormataCpoTexto(" ", 1)
            .sCodCarteira = "0"
            .sIdentEmissaoBloq = FormataCpoNum(IIf(objCarteiraCobrador.iImprimeBoleta = 0 And objParcelaRec.lIdImpressaoBoleto = 0, "1", "2"), 1)
            .sIdentDistrib = " "
            .sNumContaCorrente = FormataCpoNum(gobjCobrancaEletronica.objCCI.sNumConta, 12)
            'Está faltando o campo Varição e o DVNossoNumero
            'lErro = Calcula_NossoNumero_BcoBrasil(objCarteiraCobrador, sNossoNumero, sDVNossoNumero, .sVariacao, .sEspecieTitulo, gsNumConvenio)
               
            .sIdentTituloBanco = FormataCpoNum(objParcelaRec.sNumTitCobrador, 13) '?????
    
            'Verifica qual será o Código do Protesto
            If objOcorrRemParcRem.iInstrucao1 = 81 Or objOcorrRemParcRem.iInstrucao1 = 82 Then
        
                .sCodProtesto = IIf(objOcorrRemParcRem.iInstrucao1 = 81, "1", "2")
                .sNumDiasProtesto = FormataCpoNum(IIf(objOcorrRemParcRem.iDiasDeProtesto1 > 55, 55, objOcorrRemParcRem.iDiasDeProtesto1), 2)
        
            ElseIf objOcorrRemParcRem.iInstrucao2 = 81 Or objOcorrRemParcRem.iInstrucao2 = 82 Then
        
                .sCodProtesto = IIf(objOcorrRemParcRem.iInstrucao2 = 81, "1", "2")
                .sNumDiasProtesto = FormataCpoNum(IIf(objOcorrRemParcRem.iDiasDeProtesto2 > 55, 55, objOcorrRemParcRem.iDiasDeProtesto2), 2)
            
            ElseIf objOcorrRemParcRem.iInstrucao1 = 99 Or objOcorrRemParcRem.iInstrucao2 = 99 Then
            
                .sCodProtesto = "3"
                .sNumDiasProtesto = FormataCpoNum(0, 2)
                
            Else
            
                .sCodProtesto = FormataCpoTexto(Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 8), 1)
                .sNumDiasProtesto = FormataCpoNum(Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 9), 2)
                
            End If
        
            .sNumContratoOpCred = FormataCpoNum(Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 10), 10)
    
        Case 399
            .sCodMoeda = "09"
            .sCodJurosMora = "1"
            .sEspecieTitulo = FormataCpoNum(2, 2)
            .sDVAgMantConta = FormataCpoTexto(0, 1)
            '.sDVConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVNumConta, 1)
            .sDVConta = "0"
            .sDVAgenciaConta = FormataCpoTexto("0", 1)
            .sDVAgencia = FormataCpoTexto("0", 1)
            .sCodCarteira = FormataCpoNum(objCarteiraCobrador.iNumCarteiraNoBanco, 1)
            .sNumContaCorrente = FormataCpoNum(gobjCobrancaEletronica.objCCI.sNumConta, 12)
            .sIdentEmissaoBloq = FormataCpoNum(IIf(objCarteiraCobrador.iImprimeBoleta = 0 And objParcelaRec.lIdImpressaoBoleto = 0, "9", "2"), 1) '2:cliente, 4: banco com comprovante e 9: banco sem comprovante
            If objCarteiraCobrador.iImprimeBoleta = 0 And objParcelaRec.lIdImpressaoBoleto = 0 Then
                .sIdentDistrib = "1" 'banco distribui
            Else
                .sIdentDistrib = "2" 'cliente distribui
            End If
            
            .sCodProtesto = FormataCpoTexto(Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 8), 1)
            .sNumDiasProtesto = FormataCpoNum(Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 9), 2)
            'Está faltando o campo Varição e o DVNossoNumero
            'lErro = Calcula_NossoNumero_BcoBrasil(objCarteiraCobrador, sNossoNumero, sDVNossoNumero, .sVariacao, .sEspecieTitulo, gsNumConvenio)
               
            If objCarteiraCobrador.iComRegistro = 0 Or (Len(Trim(objParcelaRec.sNumTitCobrador)) = 0 And objCarteiraCobrador.iImprimeBoleta = 0 And objParcelaRec.lIdImpressaoBoleto = 0) Then
                .sIdentTituloBanco = FormataCpoNum(objParcelaRec.sNumTitCobrador, 11) '?????
            Else
                'obter dv do nosso numero
                lErro = CF("Preenche_Boleto_HSBC", gobjCobrancaEletronica.objCCI, objParcelaRec, sCampoLivre, objCarteiraCobrador, False)
                If lErro <> SUCESSO Then gError 51819
                    
                .sIdentTituloBanco = Mid(sCampoLivre, 1, 11)
            End If
            .sNumContratoOpCred = FormataCpoNum(Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 10), 10)
    
        Case 341
            .sCodMoeda = ""
            .sCodJurosMora = "0"
            .sEspecieTitulo = FormataCpoNum(1, 2)
            .sDVAgMantConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVAgencia, 1)
            .sCodBaixaDevolucao = "0"
            .sNumDiasBaixaDevol = FormataCpoNum(0, 2)
            .sDVConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVNumConta, 1)
            .sDVAgenciaConta = FormataCpoTexto("", 1)
            .sDVAgencia = FormataCpoTexto("0", 1)
            .sCodCarteira = FormataCpoNum(objCarteiraCobrador.iNumCarteiraNoBanco, 3)
            .sIdentEmissaoBloq = ""
            .sIdentDistrib = ""
            .sNumContaCorrente = FormataCpoNum(gobjCobrancaEletronica.objCCI.sNumConta, 12)
    
            If Len(Trim(objParcelaRec.sNumTitCobrador)) = 0 And objCarteiraCobrador.iImprimeBoleta = 0 And objParcelaRec.lIdImpressaoBoleto = 0 Then
                .sIdentTituloBanco = FormataCpoNum("", 9)
            Else
                'obter dv do nosso numero
                lErro = CF("Preenche_Boleto_ITAU", gobjCobrancaEletronica.objCCI, objParcelaRec, sCampoLivre, objCarteiraCobrador, False)
                If lErro <> SUCESSO Then gError 51819
                
                .sIdentTituloBanco = FormataCpoNum(objParcelaRec.sNumTitCobrador, 8) & Mid(sCampoLivre, 12, 1)
            End If
            
            'Verifica qual será o Código do Protesto
            If objOcorrRemParcRem.iInstrucao1 = 81 Or objOcorrRemParcRem.iInstrucao1 = 82 Then
        
                .sCodProtesto = IIf(objOcorrRemParcRem.iInstrucao1 = 81, "1", "2")
                .sNumDiasProtesto = FormataCpoNum(IIf(objOcorrRemParcRem.iDiasDeProtesto1 > 55, 55, objOcorrRemParcRem.iDiasDeProtesto1), 2)
        
            ElseIf objOcorrRemParcRem.iInstrucao2 = 81 Or objOcorrRemParcRem.iInstrucao2 = 82 Then
        
                .sCodProtesto = IIf(objOcorrRemParcRem.iInstrucao2 = 81, "1", "2")
                .sNumDiasProtesto = FormataCpoNum(IIf(objOcorrRemParcRem.iDiasDeProtesto2 > 55, 55, objOcorrRemParcRem.iDiasDeProtesto2), 2)
            
            ElseIf objOcorrRemParcRem.iInstrucao1 = 99 Or objOcorrRemParcRem.iInstrucao2 = 99 Then
            
                .sCodProtesto = "3"
                .sNumDiasProtesto = FormataCpoNum(0, 2)
                
            Else
            
                .sCodProtesto = FormataCpoNum(Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 8), 1)
                .sNumDiasProtesto = FormataCpoNum(Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 9), 2)
                
            End If
        
            .sNumContratoOpCred = FormataCpoNum(0, 13)
            
        Case 745
            .sCodMoeda = "09"
            .sCodJurosMora = "1"
            .sEspecieTitulo = FormataCpoNum(3, 2) 'DMI
            .sDVAgMantConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVAgencia, 1)
            .sDVConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVNumConta, 1)
            .sDVAgenciaConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVAgConta, 1)
            .sDVAgencia = FormataCpoTexto("0", 1)
            .sIdentEmissaoBloq = FormataCpoNum(IIf(objCarteiraCobrador.iImprimeBoleta = 0 And objParcelaRec.lIdImpressaoBoleto = 0, "1", "2"), 1)
            .sIdentDistrib = "1"
            Select Case objCarteiraCobrador.iNumCarteiraNoBanco
                Case Else
                    Select Case objCarteiraCobrador.iCodCarteiraCobranca
                        Case CARTEIRA_SIMPLES
                            .sCodCarteira = "1"
                        Case CARTEIRA_VINCULADA
                            .sCodCarteira = "2"
                        Case CARTEIRA_CAUCIONADA
                            .sCodCarteira = "3"
                        Case CARTEIRA_DESCONTADA
                            .sCodCarteira = "4"
                    End Select
            End Select
            .sNumContaCorrente = FormataCpoNum(gobjCobrancaEletronica.objCCI.sNumConta, 12)
            'Está faltando o campo Varição e o DVNossoNumero
            'lErro = Calcula_NossoNumero_BcoBrasil(objCarteiraCobrador, sNossoNumero, sDVNossoNumero, .sVariacao, .sEspecieTitulo, gsNumConvenio)
               
            .sIdentTituloBanco = FormataCpoNum(objParcelaRec.sNumTitCobrador, 20) '?????
    
            'Verifica qual será o Código do Protesto
            If objOcorrRemParcRem.iInstrucao1 = 81 Or objOcorrRemParcRem.iInstrucao1 = 82 Then
        
                .sCodProtesto = IIf(objOcorrRemParcRem.iInstrucao1 = 81, "1", "2")
                .sNumDiasProtesto = FormataCpoNum(objOcorrRemParcRem.iDiasDeProtesto1, 2)
        
            ElseIf objOcorrRemParcRem.iInstrucao2 = 81 Or objOcorrRemParcRem.iInstrucao2 = 82 Then
        
                .sCodProtesto = IIf(objOcorrRemParcRem.iInstrucao2 = 81, "1", "2")
                .sNumDiasProtesto = FormataCpoNum(objOcorrRemParcRem.iDiasDeProtesto2, 2)
            Else
                .sCodProtesto = FormataCpoTexto(Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 8), 1)
                .sNumDiasProtesto = FormataCpoNum(Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 9), 2)
            End If
        
            If .sCodProtesto = " " Then
            
                'se for cobranca caucionada => nao protesta
                If objCarteiraCobrador.iNumCarteiraNoBanco = 3 Then
                
                    .sCodProtesto = "3"
                    .sNumDiasProtesto = FormataCpoNum(0, 2)
                    
                Else
                
                    .sCodProtesto = "2" 'dias uteis
                    .sNumDiasProtesto = FormataCpoNum(5, 2)
                    
                End If
                
            End If
            
            .sNumContratoOpCred = FormataCpoNum(0, 10)
            
        Case 246
        
            If Len(Trim(Obter_Info_CarteiraCobrador(CARTEIRA_SIMPLES, gobjCobrancaEletronica.colCarteirasCNABInfo, 7))) = 0 Then
                .sNumContaCorrente = FormataCpoNum(gobjCobrancaEletronica.objCCI.sNumConta, 12)
                .sDVConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVNumConta, 1)
                .sAgMantConta = FormataCpoNum(gobjCobrancaEletronica.objCCI.sAgencia, 5)
                .sDVAgencia = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVAgencia, 1)
                .sDVAgenciaConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVAgConta, 1)
            Else
                sCodConvBanco = Obter_Info_CarteiraCobrador(CARTEIRA_SIMPLES, gobjCobrancaEletronica.colCarteirasCNABInfo, 7)
                .sAgMantConta = Mid(sCodConvBanco, 1, 5)
                .sDVAgencia = Mid(sCodConvBanco, 6, 1)
                .sNumContaCorrente = Mid(sCodConvBanco, 7, 12)
                .sDVConta = Mid(sCodConvBanco, 19, 1)
                .sDVAgenciaConta = Mid(sCodConvBanco, 20, 1)
            End If
            
'            .sDVConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVNumConta, 1)
'            .sDVAgenciaConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVAgConta, 1)
'            .sDVAgencia = FormataCpoTexto(" ", 1)
'            .sDVAgMantConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVAgencia, 1)
            .sIdentTituloBanco = FormataCpoNum(objParcelaRec.sNumTitCobrador, 11)
            .sCodCarteira = FormataCpoNum(objCarteiraCobrador.iNumCarteiraNoBanco, 1)
'            .sNumContaCorrente = FormataCpoNum(gobjCobrancaEletronica.objCCI.sNumConta, 12)
            
            .sCodMoeda = "09"
            
            .sCodJurosMora = "1"
            .sEspecieTitulo = FormataCpoNum(2, 2)
            .sIdentEmissaoBloq = FormataCpoNum(IIf(objCarteiraCobrador.iImprimeBoleta = 0 And objParcelaRec.lIdImpressaoBoleto = 0, "1", "2"), 1)
            .sIdentDistrib = "1"
            
            .sCodProtesto = FormataCpoTexto(Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 8), 1)
            .sNumDiasProtesto = FormataCpoNum(Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 9), 2)

            If .sCodProtesto = "0" Then .sCodProtesto = "3"

            .sNumContratoOpCred = FormataCpoNum(Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 10), 10)
        
        Case 41
            .sCodMoeda = "09"
            .sCodJurosMora = "0"
            .sEspecieTitulo = FormataCpoTexto(Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 32), 2)
            .sDVAgMantConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVAgencia, 1)
            .sCodBaixaDevolucao = "0"
            .sNumDiasBaixaDevol = FormataCpoNum(0, 3)
            .sDVConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVNumConta, 1)
            .sDVAgenciaConta = FormataCpoTexto("", 1)
            .sDVAgencia = FormataCpoTexto("0", 1)
            .sCodCarteira = FormataCpoNum(objCarteiraCobrador.iNumCarteiraNoBanco, 1)
            .sIdentEmissaoBloq = ""
            .sIdentDistrib = ""
            .sNumContaCorrente = FormataCpoNum(gobjCobrancaEletronica.objCCI.sNumConta, 12)
            
            lErro = CF("NossoNumero_ObtemDV_Banrisul", objParcelaRec, sDVNossoNumero)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            .sIdentTituloBanco = FormataCpoNum(objParcelaRec.sNumTitCobrador, 8) & sDVNossoNumero & FormataCpoTexto("", 10)
            .sIdentEmissaoBloq = FormataCpoNum(IIf(objCarteiraCobrador.iImprimeBoleta = 0 And objParcelaRec.lIdImpressaoBoleto = 0, "1", "2"), 1)
            .sIdentDistrib = .sIdentEmissaoBloq '"1"
            
            'Verifica qual será o Código do Protesto
            If objOcorrRemParcRem.iInstrucao1 = 81 Or objOcorrRemParcRem.iInstrucao1 = 82 Then
        
                .sCodProtesto = IIf(objOcorrRemParcRem.iInstrucao1 = 81, "1", "2")
                .sNumDiasProtesto = FormataCpoNum(IIf(objOcorrRemParcRem.iDiasDeProtesto1 > 55, 55, objOcorrRemParcRem.iDiasDeProtesto1), 2)
        
            ElseIf objOcorrRemParcRem.iInstrucao2 = 81 Or objOcorrRemParcRem.iInstrucao2 = 82 Then
        
                .sCodProtesto = IIf(objOcorrRemParcRem.iInstrucao2 = 81, "1", "2")
                .sNumDiasProtesto = FormataCpoNum(IIf(objOcorrRemParcRem.iDiasDeProtesto2 > 55, 55, objOcorrRemParcRem.iDiasDeProtesto2), 2)
            
            ElseIf objOcorrRemParcRem.iInstrucao1 = 99 Or objOcorrRemParcRem.iInstrucao2 = 99 Then
            
                .sCodProtesto = "3"
                .sNumDiasProtesto = FormataCpoNum(0, 2)
                
            Else
            
                .sCodProtesto = FormataCpoNum(Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 8), 1)
                .sNumDiasProtesto = FormataCpoNum(Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 9), 2)
                
            End If
        
            .sNumContratoOpCred = FormataCpoNum(0, 10)
            
        Case 453, 749
        
            .sCodMoeda = "09" '09 - REAL
            .sCodJurosMora = "1" '1 - Valor por dia; 2 - Taxa mensal; 3 - Isento
            .sEspecieTitulo = "02" '02 - DM DUPLICATA MERCANTIL
            '01 - CH CHEQUE
            '02 - DM DUPLICATA MERCANTIL
            '03 - DMI  DUPLICATA MERCANTIL P/ INDICAÇÃO
            '04 - DS DUPLICATA DE SERVIÇO
            '05 - DSI DUPLICATA DE SERVIÇO P/ INDICAÇÃO
            '06 - DR DUPLICATA RURAL
            '07 - LC LETRA DE CÂMBIO
            '08 - NCC NOTA DE CRÉDITO COMERCIAL
            '09 - NCE NOTA DE CRÉDITO A EXPORTAÇÃO
            '10 - NCI  NOTA DE CRÉDITO INDUSTRIAL
            '11 - NCR NOTA DE CRÉDITO RURAL
            '12 - NP NOTA PROMISSÓRIA
            '13 - NPR NOTA PROMISSÓRIA RURAL
            '14 - TM TRIPLICATA MERCANTIL
            '15 - TS TRIPLICATA DE SERVIÇO
            '16 - NS NOTA DE SEGURO
            '17 - RC RECIBO
            '18 - FAT FATURA
            '19 - ND NOTA DE CRÉDITO
            '20 - AP APÓLICE DE SEGURO
            '21 - ME MENSALIDADE ESCOLAR
            '22 - PC PARCELA DE CONSÓRCIO
            '80 - DP USO ESPECÍFICOS BANCO
            '81 - ND USO ESPECÍFICOS BANCO
            '82 - FR USO ESPECÍFICOS BANCO
            '99 - OUTROS

            .sDVAgMantConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVAgencia, 1)
            .sIdentEmissaoBloq = FormataCpoNum(IIf(objCarteiraCobrador.iImprimeBoleta = 0 And objParcelaRec.lIdImpressaoBoleto = 0, "1", "2"), 1)
            '1 - Banco emite
            '2 - Cliente emite
            '3 - Banco pré-emite e cliente complementa
            
            .sNumContaCorrente = "000" & FormataCpoNum(gobjCobrancaEletronica.objCCI.sNumConta, 10)
            .sDVConta = "0"
           
            sNossoNumero = FormataCpoNum(objParcelaRec.sNumTitCobrador, 7)
            If Len(Trim(sNossoNumero)) > 0 Then
                Call Calcula_DV11(sNossoNumero, 9, sDVNossoNumero)
                .sIdentTituloBanco = sNossoNumero & sDVNossoNumero & FormataCpoNum(0, 12)
            Else
                .sIdentTituloBanco = FormataCpoNum(0, 20)
            End If
            .sCodCarteira = CStr(objCarteiraCobrador.iNumCarteiraNoBanco)
            '1 - COBRANÇA SIMPLES EM REAIS
            '2 - COBRANÇA CAUCIONADA EM REAIS
            '3 - COBRANÇA DIRETA
            '4 - AP. TRD. C/R. S/IOS
            '5 - COBRANÇA VINCULADA CÂMBIO
            '6 - COBRANÇA VINCULADA EM REAIS
            '7 - AP. SEG. S/R. C/ IOS
            '8 - AP. TRD. C/R. C/ IOS
            '9 - AP. R$. C/R. C/ IOS
            'A - COBRANÇA SIMPLES ESPECIAL
            'i -AP.FAJ.C / r.S / IOS
            'J -AP.FAJ.C / r.C / IOS
            'L - COBRANÇA SIMPLES EM DÓLAR
            'Q - COBRANÇA VINCULADA EM DÓLAR
            'R - COBRANÇA CAUCIONADA EM DÓLAR

            .sCodBaixaDevolucao = "0"
            .sNumDiasBaixaDevol = FormataCpoNum(0, 3)
            .sDVAgenciaConta = ""
            .sDVAgencia = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVAgencia, 1)
            .sIdentDistrib = .sIdentEmissaoBloq
            .sDataJurosMora = FormataCpoNum(0, 8)
            
            'Verifica qual será o Código do Protesto
            If objOcorrRemParcRem.iInstrucao1 = 81 Or objOcorrRemParcRem.iInstrucao1 = 82 Then
        
                .sCodProtesto = IIf(objOcorrRemParcRem.iInstrucao1 = 81, "1", "2")
                .sNumDiasProtesto = FormataCpoNum(IIf(objOcorrRemParcRem.iDiasDeProtesto1 > 55, 55, objOcorrRemParcRem.iDiasDeProtesto1), 2)
        
            ElseIf objOcorrRemParcRem.iInstrucao2 = 81 Or objOcorrRemParcRem.iInstrucao2 = 82 Then
        
                .sCodProtesto = IIf(objOcorrRemParcRem.iInstrucao2 = 81, "1", "2")
                .sNumDiasProtesto = FormataCpoNum(IIf(objOcorrRemParcRem.iDiasDeProtesto2 > 55, 55, objOcorrRemParcRem.iDiasDeProtesto2), 2)
            
            ElseIf objOcorrRemParcRem.iInstrucao1 = 99 Or objOcorrRemParcRem.iInstrucao2 = 99 Then
            
                .sCodProtesto = "3"
                .sNumDiasProtesto = FormataCpoNum(0, 2)
                
            Else
            
                .sCodProtesto = FormataCpoNum(Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 8), 1)
                .sNumDiasProtesto = FormataCpoNum(Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 9), 2)
                
            End If
        
            .sNumContratoOpCred = FormataCpoNum(Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 10), 10)
        
        Case 422
            .sCodMoeda = "09"
            .sCodJurosMora = "1"
            .sEspecieTitulo = FormataCpoNum(2, 2)
            .sDVAgMantConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVAgencia, 1)
            .sDVConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVNumConta, 1)
            .sDVAgenciaConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVAgConta, 1)
            .sDVAgencia = FormataCpoTexto("0", 1)
            .sCodCarteira = FormataCpoNum(objCarteiraCobrador.iNumCarteiraNoBanco, 1)
            .sNumContaCorrente = FormataCpoNum(gobjCobrancaEletronica.objCCI.sNumConta, 12)
            .sIdentEmissaoBloq = FormataCpoNum(IIf(objCarteiraCobrador.iImprimeBoleta = 0 And objParcelaRec.lIdImpressaoBoleto = 0, "1", "2"), 1)
           
            .sCodProtesto = FormataCpoTexto(Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 8), 1)
            .sNumDiasProtesto = FormataCpoNum(Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 9), 2)
               
            If Len(Trim(objParcelaRec.sNumTitCobrador)) = 8 Then
                            
                sNossoNumero = Trim(objParcelaRec.sNumTitCobrador)
                Call CF("Calcula_DV11_Safra", sNossoNumero, 9, sDVNossoNumero)
                
                .sIdentTituloBanco = sNossoNumero & sDVNossoNumero
            Else
                .sIdentTituloBanco = Trim(objParcelaRec.sNumTitCobrador)
            End If
                
            '.sIdentTituloBanco = FormataCpoNum(.sIdentTituloBanco ,20)
            '.sIdentTituloBanco = .sIdentTituloBanco & String(20 - Len(.sIdentTituloBanco), "0")
            '.sIdentTituloBanco = FormataCpoTexto(.sIdentTituloBanco, 20)
            .sIdentTituloBanco = FormataCpoNum(.sIdentTituloBanco, 9) & FormataCpoTexto("", 11)
            
            .sNumContratoOpCred = FormataCpoNum(Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 10), 10)
            .sIdentDistrib = .sIdentEmissaoBloq
            
        Case 748
            .sCodMoeda = "09"
            .sCodJurosMora = "1"
            .sEspecieTitulo = FormataCpoNum(2, 2)
            .sDVAgMantConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVAgencia, 1)
            .sDVConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVNumConta, 1)
            .sDVAgenciaConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVAgConta, 1)
            .sDVAgencia = FormataCpoTexto("0", 1)
            .sCodCarteira = FormataCpoNum(objCarteiraCobrador.sCodCarteiraNoBanco, 1)
            .sNumContaCorrente = FormataCpoNum(gobjCobrancaEletronica.objCCI.sNumConta, 12)
            .sIdentEmissaoBloq = FormataCpoNum(IIf(objCarteiraCobrador.iImprimeBoleta = 0 And objParcelaRec.lIdImpressaoBoleto = 0, "1", "2"), 1)
           
            .sCodProtesto = FormataCpoTexto(Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 8), 1)
            .sNumDiasProtesto = FormataCpoNum(Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 9), 2)
               
'            Se a impressão for pela Sicredi (A) é possível deixar em branco (sem
'            preenchimento - gerado automaticamente pelo Banco) ou informar "Nosso
'            número " devidamente preenchido. Se for impressão pelo Beneficiário (B) -"
'            informar o "Nosso Número" conforme informações abaixo:
'            13 - ano atual
'            2 a 9 - byte de geração "somente será "1" se forem boletos pré-impressos".
'            xxxxx - número sequencial
'            d - dígito verificador calculado
'            ou seja, a nomenclatura correta é: 132xxxxxD

            sCodConvBanco = FormataCpoNum(Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 7), 11)
            sCodConvBanco = Replace(sCodConvBanco, ".", "")
          
            If Len(Trim(objParcelaRec.sNumTitCobrador)) = 5 Then 'Só sequencial, adiciona ano e byte
                            
                sNossoNumero = Format(objParcelaRec.dtDataVencimento, "YY") & "2" & Trim(objParcelaRec.sNumTitCobrador)
                lErro = CF("Calcula_DV11_SiCred", sCodConvBanco & sNossoNumero, 9, sDVNossoNumero)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

                .sIdentTituloBanco = sNossoNumero & sDVNossoNumero
                
            ElseIf Len(Trim(objParcelaRec.sNumTitCobrador)) = 6 Then 'Sem ano
                            
                sNossoNumero = Format(objParcelaRec.dtDataVencimento, "YY") & Trim(objParcelaRec.sNumTitCobrador)
                lErro = CF("Calcula_DV11_SiCred", sCodConvBanco & sNossoNumero, 9, sDVNossoNumero)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

                .sIdentTituloBanco = sNossoNumero & sDVNossoNumero
                
            ElseIf Len(Trim(objParcelaRec.sNumTitCobrador)) = 8 Then 'Sem DV
            
                sNossoNumero = Trim(objParcelaRec.sNumTitCobrador)
                lErro = CF("Calcula_DV11_SiCred", sCodConvBanco & sNossoNumero, 9, sDVNossoNumero)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                .sIdentTituloBanco = sNossoNumero & sDVNossoNumero
                
            Else
                .sIdentTituloBanco = Trim(objParcelaRec.sNumTitCobrador)
            End If
                
            .sIdentTituloBanco = FormataCpoTexto(.sIdentTituloBanco, 20)
            
            .sNumContratoOpCred = FormataCpoNum("", 10)
            .sIdentDistrib = .sIdentEmissaoBloq
            
        Case 637
            .sCodMoeda = "09"
            .sCodJurosMora = "1"
            .sEspecieTitulo = FormataCpoNum(2, 2)
            .sDVAgMantConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVAgencia, 1)
            .sDVConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVNumConta, 1)
            .sDVAgenciaConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVAgConta, 1)
            .sDVAgencia = FormataCpoTexto("0", 1)
            .sCodCarteira = FormataCpoNum(objCarteiraCobrador.sCodCarteiraNoBanco, 1)
            .sNumContaCorrente = FormataCpoNum(gobjCobrancaEletronica.objCCI.sNumConta, 12)
            .sIdentEmissaoBloq = FormataCpoNum(IIf(objCarteiraCobrador.iImprimeBoleta = 0 And objParcelaRec.lIdImpressaoBoleto = 0, "1", "2"), 1)
           
            .sCodProtesto = FormataCpoTexto(Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 8), 1)
            .sNumDiasProtesto = FormataCpoNum(Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 9), 2)

            .sIdentTituloBanco = Trim(objParcelaRec.sNumTitCobrador)
            .sIdentTituloBanco = FormataCpoNum(.sIdentTituloBanco, 20)
            .sIdentDistrib = .sIdentEmissaoBloq
            .sNumContratoOpCred = FormataCpoNum(Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 10), 10)
            
        Case 85
            .sCodMoeda = "09"
            .sCodJurosMora = "1"
            .sEspecieTitulo = FormataCpoNum(2, 2)
            .sDVAgMantConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVAgencia, 1)
            .sDVConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVNumConta, 1)
            .sDVAgenciaConta = " "
            .sDVAgencia = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVAgencia, 1)
            .sCodCarteira = FormataCpoNum(objCarteiraCobrador.iNumCarteiraNoBanco, 1)
            .sNumContaCorrente = FormataCpoNum(gobjCobrancaEletronica.objCCI.sNumConta, 12)
            .sIdentEmissaoBloq = FormataCpoNum(IIf(objCarteiraCobrador.iImprimeBoleta = 0 And objParcelaRec.lIdImpressaoBoleto = 0, "1", "2"), 1)
           
            .sCodProtesto = FormataCpoTexto(Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 8), 1)
            .sNumDiasProtesto = FormataCpoNum(Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 9), 2)
            .sIdentTituloBanco = FormataCpoTexto(objParcelaRec.sNumTitCobrador, 20)
            
            .sNumContratoOpCred = FormataCpoNum(Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 10), 10)
            .sIdentDistrib = .sIdentEmissaoBloq
            .sTipoDocumento = "1"
            .sNumDiasBaixaDevol = "   "

    End Select
    
    .sNumDocCobranca = FormataCpoTexto(gobjCobrancaEletronica.lNumTitulo & SEPARADOR & objParcelaRec.iNumParcela, 15)
    
    If objOcorrRemParcRem.iCodOcorrencia = 6 Then
        .sDataVencTitulo = Format(objOcorrRemParcRem.dtNovaDataVcto, "DDMMYYYY")
    Else
        .sDataVencTitulo = Format(objParcelaRec.dtDataVencimento, "DDMMYYYY")
    End If
    
    .sValorNominalTitulo = FormataCpoValor(objOcorrRemParcRem.dValorCobrado, 15)
    .sAgEncCobranca = FormataCpoNum(0, 5)
    
    If objParcelaRec.iAceite = 1 Then
        .sIdentTituloAceite = "A"
    Else
        .sIdentTituloAceite = "N"
    End If
    
    .sDataEmissaoTitulo = Format(gobjCobrancaEletronica.dtDataEmissaoTitulo, "DDMMYYYY")
    .sDataJurosMora = .sDataVencTitulo
    If .sCodJurosMora = "1" Or gobjCobrancaEletronica.objCobrador.iCodBanco <> 356 Then
        .sJurosMoraDiaTaxa = FormataCpoValor(Round(objOcorrRemParcRem.dValorCobrado * gobjCobrancaEletronica.dPercJurosDiario, 2), 15)
    ElseIf .sCodJurosMora = "2" And gobjCobrancaEletronica.objCobrador.iCodBanco = 356 Then
        .sJurosMoraDiaTaxa = FormataCpoValor(Round(gobjCobrancaEletronica.dPercJurosDiario * 3000, 2), 15)
    Else
        .sJurosMoraDiaTaxa = FormataCpoValor(0, 15)
    End If
    .sCodDesconto1 = FormataCpoNum(objParcelaRec.iDesconto1Codigo, 1)
    
    If objParcelaRec.dtDesconto1Ate <> DATA_NULA Then
        .sDataDesconto1 = Format(objParcelaRec.dtDesconto1Ate, "DDMMYYYY")
    Else
        'If gobjCobrancaEletronica.objCobrador.iCodBanco <> 104 Then
            .sDataDesconto1 = FormataCpoNum(0, 8)
        'Else
            '.sDataDesconto1 = .sDataVencTitulo
        'End If
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
    .sIdentTituloEmp = FormataCpoTexto(gobjCobrancaEletronica.lNumTitulo & SEPARADOR & objParcelaRec.iNumParcela, 25)
        
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
        
    .sBrancosCNAB2 = FormataCpoTexto("", 1)
    
    End With
    
    Obter_Dados_DetalheSegP = SUCESSO

    Exit Function
    
Erro_Obter_Dados_DetalheSegP:

Obter_Dados_DetalheSegP = gErr
    
    Select Case gErr

        Case 51817, 51818, 51819
       
        Case 51820
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ENDERECO_FILIALCLIENTE_NAO_INFORMADO", gErr, gobjCobrancaEletronica.lCliente, gobjCobrancaEletronica.iFilialCli)

        Case 51821
            lErro = Rotina_Erro(vbOKOnly, "ERRO_OCORRENCIA_NAO_TRATADA", gErr, objOcorrRemParcRem.iCodOcorrencia, objParcelaRec.sNumTitCobrador)

        Case 51822
        
        Case ERRO_SEM_MENSAGEM
        
        Case 202346
            lErro = Rotina_Erro(vbOKOnly, "ERRO_OCORRENCIA_FALTA_INFO_LIQ", gErr, objParcelaRec.sNumTitCobrador)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147095)

    End Select

    Exit Function

End Function

Private Function Obter_Dados_DetalheSegQ(tRegDetalheCNAB240SegQ As typeRemDetalheCNAB240SegQ, objOcorrRemParcRem As ClassOcorrRemParcRec, objCarteiraCobrador As ClassCarteiraCobrador, objParcelaRec As ClassParcelaReceber) As Long
'Obtem os dados que são utilizados para preencher o registro de detalhe

Dim lErro As Long
Dim objEndereco As New ClassEndereco
Dim sNossoNumero As String
Dim dValorDesconto As Double
Dim sDVNossoNumero As String

On Error GoTo Erro_Obter_Dados_DetalheSegQ

    lErro = CF("CobrancaEletronica_Obter_Endereco", gobjCobrancaEletronica, objEndereco)
    If lErro <> SUCESSO Then Error 51818
        
    With tRegDetalheCNAB240SegQ
           
        .sCodBancoComp = FormataCpoNum(gobjCobrancaEletronica.objCobrador.iCodBanco, 3)
        .sLoteServico = FormataCpoNum(giNumLotes, 4)
        .sRegDetalhe = "3"
        .sNumSequencialReg = FormataCpoNum(gobjCobrancaEletronica.lSequencialRegistro, 5)
        .sCodSegRegDetalhe = "Q"
        .sBrancosCNAB1 = FormataCpoTexto("", 1)
        Select Case objOcorrRemParcRem.iCodOcorrencia
            Case 1
                .sCodMovimento = "01"
            Case 2
                .sCodMovimento = "02"
            Case 6
                .sCodMovimento = "06"
            Case Else
                'Error
        End Select
    
        If Len(gobjCobrancaEletronica.sCGCSacado) = 14 Then
            .sTipoInscricao = "2"
            .sNumInscricao = FormataCpoNum(gobjCobrancaEletronica.sCGCSacado, 15)
        Else
            .sTipoInscricao = "1"
            ' .sNumInscricao = left(FormataCpoNum(gobjCobrancaEletronica.sCGCSacado, 11), 9) & "0000" & FormataCpoNum(right(gobjCobrancaEletronica.sCGCSacado, 2), 2)
            .sNumInscricao = FormataCpoNum(gobjCobrancaEletronica.sCGCSacado, 15)
        End If
        
        If gobjCobrancaEletronica.objCobrador.iCodBanco = 341 Then
            .sNome = FormataCpoTexto(Trata_Texto(gobjCobrancaEletronica.sNomeSacado, 30), 40)
        Else
            .sNome = FormataCpoTexto(Trata_Texto(gobjCobrancaEletronica.sNomeSacado), 40)
        End If
        
        .sEndereco = FormataCpoTexto(Trata_Texto(objEndereco.sEndereco), 40)
        .sBairro = FormataCpoTexto(Trata_Texto(objEndereco.sBairro), 15)
        .sCEP = Formata_Cep(objEndereco.sCEP, 8) 'Incluido o sufixo.
        .sCidade = FormataCpoTexto(Trata_Texto(objEndereco.sCidade), 15)
        .sUnidFederacao = FormataCpoTexto(objEndereco.sSiglaEstado, 2)
        
'        If Len(Trim(gobjCobrancaEletronica.sSacadorAvalista)) = 0 Then
            'If gobjCobrancaEletronica.objCobrador.iCodBanco <> 104 Or gobjCobrancaEletronica.sVersaoLayout = "030" Then
            'Banco 104 - CEF Versão 050 tem que ir 0 - Não Informado e não 2 - CNPJ
                .sTipoInscricaoSacAval = FormataCpoNum(0, 1)
            'Else
                '.sTipoInscricaoSacAval = FormataCpoNum(2, 1)
            'End If
            .sNumInscricaoSacAval = FormataCpoNum(0, 15)
            .sNomeSacadorAvalista = FormataCpoTexto("", 40)
                
'        Else
'
'            .sTipoInscricaoSacAval = FormataCpoNum(2, 1)
'            .sNumInscricaoSacAval = FormataCpoNum(gobjCobrancaEletronica.sCGCSacadorAvalista, 15)
'            .sNomeSacadorAvalista = FormataCpoTexto(gobjCobrancaEletronica.sSacadorAvalista, 40)
'
'        End If
        
        If gobjCobrancaEletronica.objCobrador.iCodBanco <> 104 Then
            .sBancoCompensacao = FormataCpoNum(0, 3)
        Else
            .sBancoCompensacao = FormataCpoTexto("", 3)
        End If
        
        .sNossoNumBancoCorresp = FormataCpoTexto("", 20)
        .sBrancosCNAB2 = FormataCpoTexto("", 8)
        
        Select Case gobjCobrancaEletronica.objCobrador.iCodBanco
            Case 8, 33, 353
                .sBancoCompensacao = FormataCpoNum(0, 12)
                .sNossoNumBancoCorresp = FormataCpoTexto("", 11)
            Case 356
                'cpf
                If .sTipoInscricao = "1" Then
                    .sNumInscricao = left(FormataCpoNum(gobjCobrancaEletronica.sCGCSacado, 11), 9) & "0000" & FormataCpoNum(right(gobjCobrancaEletronica.sCGCSacado, 2), 2)
                End If
        End Select
        
    End With
    
    Obter_Dados_DetalheSegQ = SUCESSO

    Exit Function
    
Erro_Obter_Dados_DetalheSegQ:

Obter_Dados_DetalheSegQ = Err
    
    Select Case Err

        Case 51817, 51818, 51819
       
        Case 51820
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ENDERECO_FILIALCLIENTE_NAO_INFORMADO", Err, gobjCobrancaEletronica.lCliente, gobjCobrancaEletronica.iFilialCli)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147096)

    End Select

    Exit Function

End Function

Private Function Obter_Dados_DetalheSegR(tRegDetalheCNAB240SegR As typeRemDetalheCNAB240SegR, objOcorrRemParcRem As ClassOcorrRemParcRec, objCarteiraCobrador As ClassCarteiraCobrador, objParcelaRec As ClassParcelaReceber) As Long
'Obtem os dados que são utilizados para preencher o registro de detalhe

Dim lErro As Long
Dim objEndereco As New ClassEndereco
Dim sNossoNumero As String
Dim dValorDesconto As Double
Dim sDVNossoNumero As String
Dim o As ClassTituloReceber

On Error GoTo Erro_Obter_Dados_DetalheSegR

    With tRegDetalheCNAB240SegR
       
        .sCodBancoComp = FormataCpoNum(gobjCobrancaEletronica.objCobrador.iCodBanco, 3)
        .sLoteServico = FormataCpoNum(giNumLotes, 4)
        .sRegDetalhe = "3"
        .sNumSequencialReg = FormataCpoNum(gobjCobrancaEletronica.lSequencialRegistro, 5)
        .sCodSegRegDetalhe = "R"
        .sBrancosCNAB1 = FormataCpoTexto("", 1)
        Select Case objOcorrRemParcRem.iCodOcorrencia
            Case 1
                .sCodMovimento = "01"
            Case 2
                .sCodMovimento = "02"
            Case 6
                .sCodMovimento = "06"
            Case Else
                'Error
        End Select
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
            If objParcelaRec.iDesconto1Codigo = VALOR_FIXO Then
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
            If objParcelaRec.iDesconto1Codigo = VALOR_FIXO Then
                dValorDesconto = objParcelaRec.dDesconto3Valor
            End If
        End If
        
        .sValorPerConcedido3 = FormataCpoValor(dValorDesconto, 15)
        
        .sCodMulta = "0"
        .sDataMulta = FormataCpoNum(0, 8)
        .sValorPerAplicado = FormataCpoValor(0, 15) 'deixar ser determinada no convenio
        
'        .sCodMulta = "2"
'        .sDataMulta = FormataCpoNum(0, 8)
'        .sValorPerAplicado = FormataCpoValor(gobjCobrancaEletronica.dPercMulta * 100, 15)
        
        .sInfoBancoSacado = FormataCpoTexto("", 10)
        .sMensagem3 = FormataCpoTexto("", 40)
        .sMensagem4 = FormataCpoTexto("", 40)
        'Alterado por cyntia em 28/05/02
        
        .sCodBancoContaDeb = FormataCpoTexto("", 3)
        .sCodAgContaDeb = FormataCpoTexto("", 4)
        .sCodContaDVDeb = FormataCpoTexto("", 13)
        .sCodOcorrenciaSac = FormataCpoTexto("", 8)
        
        .sBrancosCNAB2 = FormataCpoTexto("", 33)
        
        Select Case gobjCobrancaEletronica.objCobrador.iCodBanco
            
            Case 1
                .sCodBancoContaDeb = FormataCpoNum(0, 3)
                .sCodAgContaDeb = FormataCpoNum(0, 4)
                .sCodContaDVDeb = FormataCpoNum(0, 13)
                .sCodOcorrenciaSac = FormataCpoNum(0, 8)
                
                .sCodMulta = FormataCpoNum(Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 23), 1)
                If .sCodMulta = "1" Then
                    .sDataMulta = Format(objParcelaRec.dtDataVencimento, "DDMMYYYY")
                    .sValorPerAplicado = FormataCpoValor(Round(objOcorrRemParcRem.dValorCobrado * gobjCobrancaEletronica.dPercMulta, 2), 15)
                End If
                
                If .sCodMulta = "2" Then
                    .sDataMulta = Format(objParcelaRec.dtDataVencimento, "DDMMYYYY")
                    .sValorPerAplicado = FormataCpoValor(gobjCobrancaEletronica.dPercMulta * 100, 15)
                End If
                
            Case 341
                .sCodMulta = FormataCpoNum(Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 23), 1)
                If .sCodMulta = "1" Then
                    .sDataMulta = Format(objParcelaRec.dtDataVencimento, "DDMMYYYY")
                    .sValorPerAplicado = FormataCpoValor(Round(objOcorrRemParcRem.dValorCobrado * gobjCobrancaEletronica.dPercMulta, 2), 15)
                End If
                
                If .sCodMulta = "2" Then
                    .sDataMulta = Format(objParcelaRec.dtDataVencimento, "DDMMYYYY")
                    .sValorPerAplicado = FormataCpoValor(gobjCobrancaEletronica.dPercMulta * 100, 15)
                End If
                            
            Case 8, 33, 353
                .sCodDesconto3 = FormataCpoTexto("", 1)
                .sDataDesconto3 = FormataCpoTexto("", 8)
                .sValorPerConcedido3 = FormataCpoTexto("", 15)
                
                .sCodMulta = FormataCpoNum(Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 12), 1)
                If .sCodMulta = "1" Then
                    .sDataMulta = Format(objParcelaRec.dtDataVencimento, "DDMMYYYY")
                    .sValorPerAplicado = FormataCpoValor(Round(objOcorrRemParcRem.dValorCobrado * gobjCobrancaEletronica.dPercMulta, 2), 15)
                Else
                
                    If gobjCobrancaEletronica.dPercMulta <> 0 Then
                        .sCodMulta = "2"
                        .sDataMulta = FormataCpoNum(0, 8)
                        .sValorPerAplicado = FormataCpoValor(gobjCobrancaEletronica.dPercMulta * 100, 15)
                    End If
                
                End If
                
            Case 85
                .sCodMulta = FormataCpoNum(Obter_Info_CarteiraCobrador(objCarteiraCobrador.iCodCarteiraCobranca, gobjCobrancaEletronica.colCarteirasCNABInfo, 23), 1)
                If .sCodMulta = "1" Then
                    .sDataMulta = Format(objParcelaRec.dtDataVencimento, "DDMMYYYY")
                    .sValorPerAplicado = FormataCpoValor(Round(objOcorrRemParcRem.dValorCobrado * gobjCobrancaEletronica.dPercMulta, 2), 15)
                End If
                
                If .sCodMulta = "2" Then
                    .sDataMulta = Format(objParcelaRec.dtDataVencimento, "DDMMYYYY")
                    .sValorPerAplicado = FormataCpoValor(gobjCobrancaEletronica.dPercMulta * 100, 15)
                End If
                
        End Select
        
    End With
    
    Obter_Dados_DetalheSegR = SUCESSO

    Exit Function
    
Erro_Obter_Dados_DetalheSegR:

Obter_Dados_DetalheSegR = Err
    
    Select Case Err

        Case 51817, 51818, 51819
       
        Case 51820
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ENDERECO_FILIALCLIENTE_NAO_INFORMADO", Err, gobjCobrancaEletronica.lCliente, gobjCobrancaEletronica.iFilialCli)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147097)

    End Select

    Exit Function

End Function

Private Function Obter_Dados_Header_Arquivo(tRegHeaderCNAB240 As typeRemHeaderCNAB240) As Long

Dim lErro As Long, sDVConta As String, sAux As String
Dim sCedente As String, sCGCCedente As String
Dim colCobradorInfo As New Collection, objCodigoTexto As AdmCodigoNome
Dim sCodConvBco As String

On Error GoTo Erro_Obter_Dados_Header_Arquivo
      
    'Carrega os campos do cabeçalho
    With tRegHeaderCNAB240
   
        .sCodBancoComp = FormataCpoNum(gobjCobrancaEletronica.objCobrador.iCodBanco, 3)
        .sLoteServico = FormataCpoNum(giNumLotes, 4)
        .sRegHeaderLote = "0"
                
        .sBrancosCNAB1 = FormataCpoTexto("", 9)
        
        gobjCobrancaEletronica.sCedente = gsNomeEmpresa
        gobjCobrancaEletronica.sCGCCedente = gobjFilialEmpresa.sCgc
        gobjCobrancaEletronica.sSacadorAvalista = ""
        gobjCobrancaEletronica.sCGCSacadorAvalista = ""
        
        If gobjCobrancaEletronica.objCCI.iCodBanco = 33 Then
        
            lErro = CF("CobradorInfo_Le", gobjCobrancaEletronica.objCobrador.iCodigo, colCobradorInfo)
            If lErro <> SUCESSO Then gError 134224
        
            For Each objCodigoTexto In colCobradorInfo
            
                If objCodigoTexto.iCodigo = 18 And Len(Trim(objCodigoTexto.sNome)) <> 0 Then
                    gobjCobrancaEletronica.sSacadorAvalista = gobjCobrancaEletronica.sCedente
                    gobjCobrancaEletronica.sCedente = left(Trim(objCodigoTexto.sNome), 50)
                End If
                
                If objCodigoTexto.iCodigo = 19 And Len(Trim(objCodigoTexto.sNome)) <> 0 Then
                    gobjCobrancaEletronica.sCGCSacadorAvalista = gobjCobrancaEletronica.sCGCCedente
                    gobjCobrancaEletronica.sCGCCedente = left(Trim(objCodigoTexto.sNome), 14)
                End If
                
            Next
        
        End If
        
        If gobjCobrancaEletronica.objCCI.iCodBanco = 104 Then
        
            lErro = CF("CobradorInfo_Le", gobjCobrancaEletronica.objCobrador.iCodigo, colCobradorInfo)
            If lErro <> SUCESSO Then gError 134224
        
            For Each objCodigoTexto In colCobradorInfo
            
                If objCodigoTexto.iCodigo = 20 Then
                    gobjCobrancaEletronica.sVersaoLayout = left(Trim(objCodigoTexto.sNome), 3)
                End If
                
                If objCodigoTexto.iCodigo = 24 Then
                
                    .sReservadoEmpresa = UCase(Trim(objCodigoTexto.sNome))
                
                End If
            
            Next
        
            .sReservadoEmpresa = FormataCpoTexto(.sReservadoEmpresa, 20)
        
        End If
        
        If Len(gobjCobrancaEletronica.sCGCCedente) = 14 Then
            .sTipoInscEmpresa = "2"
        Else
            .sTipoInscEmpresa = "1"
        End If
        
        .sCodInsEmpresa = FormataCpoNum(gobjCobrancaEletronica.sCGCCedente, 14)
        
        .sAgMantConta = FormataCpoNum(gobjCobrancaEletronica.objCCI.sAgencia, 5)
        .sDVAgencia = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVAgencia, 1)
        .sDVAgenciaConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVAgConta, 1)
        
        '??? com variacao de acordo com o banco
        Select Case gobjCobrancaEletronica.objCobrador.iCodBanco
        
            Case 1
                .sNVersaoLayLote = "030"
                .sCodConvBanco = FormataCpoTexto(Obter_Info_CarteiraCobrador(CARTEIRA_SIMPLES, gobjCobrancaEletronica.colCarteirasCNABInfo, 7), 20)
                .sMensagem1 = FormataCpoTexto("BANCO DO BRASIL", 30)
                .sNumContaCorrente = FormataCpoNum(gobjCobrancaEletronica.objCCI.sNumConta, 12)
                .sDVConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVNumConta, 1)
                
            Case 8, 33, 353
                .sCodInsEmpresa = FormataCpoNum(gobjCobrancaEletronica.sCGCCedente, 15)
                .sNVersaoLayLote = "040"
                .sMensagem1 = FormataCpoTexto("BANCO SANTANDER", 30)
                sCodConvBco = Obter_Info_CarteiraCobrador(CARTEIRA_SIMPLES, gobjCobrancaEletronica.colCarteirasCNABInfo, 7)
                If Len(Trim(sCodConvBco)) <= 7 Then sCodConvBco = FormataCpoNum(gobjCobrancaEletronica.objCCI.sAgencia, 4) & "0000" & FormataCpoNum(Trim(sCodConvBco), 7)
                .sCodConvBanco = FormataCpoNum(sCodConvBco, 15)
        
            Case 409
                .sNVersaoLayLote = "082"
                .sCodConvBanco = FormataCpoTexto("", 20)
                .sMensagem1 = FormataCpoTexto("UNIBANCO", 30)
                .sNumContaCorrente = FormataCpoNum(gobjCobrancaEletronica.objCCI.sNumConta, 12)
                .sDVConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVNumConta, 1)
            
            Case 104
            
                If gobjCobrancaEletronica.sVersaoLayout = "050" Then
                
                    .sNVersaoLayLote = "050"
                    .sCodConvBanco = FormataCpoNum(Obter_Info_CarteiraCobrador(CARTEIRA_SIMPLES, gobjCobrancaEletronica.colCarteirasCNABInfo, 7), 6)
                    .sMensagem1 = FormataCpoTexto("CAIXA ECONOMICA FEDERAL", 30)
                
                Else
                
                    .sNVersaoLayLote = "030"
                    .sCodConvBanco = FormataCpoNum(Obter_Info_CarteiraCobrador(CARTEIRA_SIMPLES, gobjCobrancaEletronica.colCarteirasCNABInfo, 7), 16)
                    .sAgMantConta = FormataCpoNum(Mid(.sCodConvBanco, 1, 4), 5)
                    Call Calcula_DV11(.sAgMantConta, 9, sDVConta)
                    .sDVAgencia = sDVConta
                    .sNumContaCorrente = FormataCpoNum(Mid(.sCodConvBanco, 8, 8), 12) '???? tem que pegar do cedente
                    'calcular dv da conta
                    Call Calcula_DV11(.sNumContaCorrente, 9, sDVConta)
                    .sDVConta = sDVConta
                    .sMensagem1 = FormataCpoTexto("CAIXA ECONOMICA FEDERAL", 30)
                    Call Calcula_DV11(Mid(.sCodConvBanco, 1, 4) & Mid(.sCodConvBanco, 8, 8), 9, sDVConta)
                    .sDVAgenciaConta = sDVConta
                
                End If
                
            Case 479
                .sNVersaoLayLote = "030"
                .sCodConvBanco = FormataCpoNum(Obter_Info_CarteiraCobrador(CARTEIRA_SIMPLES, gobjCobrancaEletronica.colCarteirasCNABInfo, 7), 8)
                .sMensagem1 = FormataCpoTexto("BANKBOSTON", 30)
                .sNumContaCorrente = FormataCpoNum(gobjCobrancaEletronica.objCCI.sNumConta, 8)
                .sDVConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVNumConta, 1)
            
            Case 275, 356
                .sNVersaoLayLote = "040"
                .sCodConvBanco = FormataCpoTexto(Obter_Info_CarteiraCobrador(CARTEIRA_SIMPLES, gobjCobrancaEletronica.colCarteirasCNABInfo, 7), 20)
                .sMensagem1 = FormataCpoTexto("ABN AMRO SA", 30)
                .sNumContaCorrente = FormataCpoNum(gobjCobrancaEletronica.objCCI.sNumConta, 12)
                .sDVConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVNumConta, 1)
                
            Case 399
                .sNVersaoLayLote = "010"
                .sCodConvBanco = FormataCpoTexto(Obter_Info_CarteiraCobrador(CARTEIRA_SIMPLES, gobjCobrancaEletronica.colCarteirasCNABInfo, 21), 20)
                .sMensagem1 = FormataCpoTexto("HSBC BANK BRASIL SA", 30)
                .sNumContaCorrente = FormataCpoNum(gobjCobrancaEletronica.objCCI.sNumConta, 12)
                '.sDVConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVNumConta, 1)
                .sDVConta = "0"
        
            Case 341
                .sNVersaoLayLote = "040"
                .sCodConvBanco = FormataCpoTexto("", 20)
                .sMensagem1 = FormataCpoTexto("BANCO ITAU SA", 30)
                .sNumContaCorrente = FormataCpoNum(gobjCobrancaEletronica.objCCI.sNumConta, 12)
                .sDVConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVNumConta, 1)
                
            Case 745
                .sNVersaoLayLote = "082"
                .sCodConvBanco = FormataCpoTexto("", 20)
                .sMensagem1 = FormataCpoTexto("CITIBANK", 30)
                .sNumContaCorrente = FormataCpoNum(gobjCobrancaEletronica.objCCI.sNumConta, 12)
                .sDVConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVNumConta, 1)
                
            Case 246
                .sNVersaoLayLote = "040"
                .sMensagem1 = FormataCpoTexto("BCO ABC BRASIL", 30)
                
                If Len(Trim(Obter_Info_CarteiraCobrador(CARTEIRA_SIMPLES, gobjCobrancaEletronica.colCarteirasCNABInfo, 7))) = 0 Then
                    .sCodConvBanco = FormataCpoNum(gobjCobrancaEletronica.objCCI.sAgencia, 5) & FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVAgencia, 1) & FormataCpoNum(gobjCobrancaEletronica.objCCI.sNumConta, 12) & FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVNumConta, 1) & FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVAgConta, 1)
                    .sNumContaCorrente = FormataCpoNum(gobjCobrancaEletronica.objCCI.sNumConta, 12)
                    .sDVConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVNumConta, 1)
                    .sAgMantConta = FormataCpoNum(gobjCobrancaEletronica.objCCI.sAgencia, 5)
                    .sDVAgencia = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVAgencia, 1)
                    .sDVAgenciaConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVAgConta, 1)
                Else
                    .sCodConvBanco = FormataCpoTexto(Obter_Info_CarteiraCobrador(CARTEIRA_SIMPLES, gobjCobrancaEletronica.colCarteirasCNABInfo, 7), 20)
                    .sAgMantConta = Mid(.sCodConvBanco, 1, 5)
                    .sDVAgencia = Mid(.sCodConvBanco, 6, 1)
                    .sNumContaCorrente = Mid(.sCodConvBanco, 7, 12)
                    .sDVConta = Mid(.sCodConvBanco, 19, 1)
                    .sDVAgenciaConta = Mid(.sCodConvBanco, 20, 1)
                End If
                
            Case 41
                .sNVersaoLayLote = "040"
                .sCodConvBanco = FormataCpoTexto(Obter_Info_CarteiraCobrador(CARTEIRA_SIMPLES, gobjCobrancaEletronica.colCarteirasCNABInfo, 7), 20)
                .sMensagem1 = FormataCpoTexto("BANRISUL", 30)
                .sNumContaCorrente = FormataCpoNum(gobjCobrancaEletronica.objCCI.sNumConta, 12)
                .sDVConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVNumConta, 1)
                
            Case 453
                .sNVersaoLayLote = "030"
                .sCodConvBanco = "Cobrança" + FormataCpoTexto("", 12)
                .sMensagem1 = FormataCpoTexto("Rural", 30)
                .sNumContaCorrente = "000" & FormataCpoNum(gobjCobrancaEletronica.objCCI.sNumConta, 10)
                .sDVConta = ""
                .sDVAgenciaConta = "0"
                
             Case 749
                .sNVersaoLayLote = "030"
                .sCodConvBanco = "Cobrança" + FormataCpoTexto("", 12)
                .sMensagem1 = FormataCpoTexto("BR Mercantil", 30)
                .sNumContaCorrente = "000" & FormataCpoNum(gobjCobrancaEletronica.objCCI.sNumConta, 10)
                .sDVConta = ""
                .sDVAgenciaConta = "0"
            
             Case 422
                .sNVersaoLayLote = "087"
                .sCodConvBanco = FormataCpoTexto(Obter_Info_Cobrador(gobjCobrancaEletronica.objCobradorCNABInfo.colInformacoes, 1), 20)
                .sMensagem1 = FormataCpoTexto("BANCO SAFRA", 30)
                .sNumContaCorrente = FormataCpoNum(gobjCobrancaEletronica.objCCI.sNumConta, 12)
                .sDVConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVNumConta, 1)
                
            Case 748
                .sNVersaoLayLote = "081"
                .sCodConvBanco = FormataCpoTexto(Obter_Info_CarteiraCobrador(CARTEIRA_SIMPLES, gobjCobrancaEletronica.colCarteirasCNABInfo, 7), 20)
                .sMensagem1 = FormataCpoTexto("SICRED", 30)
                .sNumContaCorrente = FormataCpoNum(gobjCobrancaEletronica.objCCI.sNumConta, 12)
                .sDVConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVNumConta, 1)
                
            Case 637
                .sNVersaoLayLote = "083"
                .sCodConvBanco = FormataCpoTexto(Obter_Info_CarteiraCobrador(CARTEIRA_SIMPLES, gobjCobrancaEletronica.colCarteirasCNABInfo, 7), 20)
                .sMensagem1 = FormataCpoTexto("BANCO SOFISA SA", 30)
                .sNumContaCorrente = FormataCpoNum(gobjCobrancaEletronica.objCCI.sNumConta, 12)
                .sDVConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVNumConta, 1)
                
            Case 85
                .sNVersaoLayLote = "087"
                .sCodConvBanco = FormataCpoTexto(Obter_Info_CarteiraCobrador(CARTEIRA_SIMPLES, gobjCobrancaEletronica.colCarteirasCNABInfo, 7), 20)
                .sMensagem1 = FormataCpoTexto("CECRED", 30)
                .sNumContaCorrente = FormataCpoNum(gobjCobrancaEletronica.objCCI.sNumConta, 12)
                .sDVConta = FormataCpoTexto(gobjCobrancaEletronica.objCCI.sDVNumConta, 1)
                
        End Select
        
        .sMensagem2 = FormataCpoTexto("", 10)
        
        .sNomeEmpresa = FormataCpoTexto(UCase(gobjCobrancaEletronica.sCedente), 30)
        
        Call CF("Empresa_Le_Nome_Customizado", .sNomeEmpresa)

        .sNumRemessa = FormataCpoNum(gobjCobrancaEletronica.objCobrador.lCNABProxSeqArqCobr, 6)
        .sDataGravRem = Format(gdtDataHoje, "DDMMYYYY")
    
    End With
   
    Obter_Dados_Header_Arquivo = SUCESSO

    Exit Function
    
Erro_Obter_Dados_Header_Arquivo:

    Obter_Dados_Header_Arquivo = gErr
    
    Select Case gErr
    
        Case 134224
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147098)
            
    End Select
    
    Exit Function

End Function

Private Function Montar_RegistroHeaderArquivo_BcoBrasil(sRegHeader As String) As Long
'Monta o registro

Dim tRegHeaderCNAB240 As typeRemHeaderCNAB240
Dim lErro As Long

On Error GoTo Erro_Montar_RegistroHeaderArquivo_BcoBrasil

    'Carrega tRegHeaderBcoBrasil com os dados do Header do aquivo
    lErro = Obter_Dados_Header_Arquivo(tRegHeaderCNAB240)
    If lErro <> SUCESSO Then Error 51813
        
    With tRegHeaderCNAB240
    
        'Monta o registro de Header do BcoBrasil
        sRegHeader = .sCodBancoComp & .sLoteServico & .sRegHeaderLote & .sBrancosCNAB1 & .sTipoInscEmpresa & .sCodInsEmpresa
        
        '??? com variacao de acordo com o banco
        Select Case gobjCobrancaEletronica.objCobrador.iCodBanco
        
            Case 1
                sRegHeader = sRegHeader & .sCodConvBanco & .sAgMantConta & .sDVAgencia & .sNumContaCorrente & .sDVConta & .sDVAgenciaConta
                sRegHeader = sRegHeader & .sNomeEmpresa & .sMensagem1 & .sMensagem2 & "1" & .sDataGravRem & "120000" & .sNumRemessa & .sNVersaoLayLote & "00000" & FormataCpoTexto("", 20 + 20 + 11) & "CSP" & "000" & "  " & FormataCpoTexto("", 10)
                
            Case 8, 33, 353
                sRegHeader = .sCodBancoComp & .sLoteServico & .sRegHeaderLote & FormataCpoTexto("", 8) & .sTipoInscEmpresa & .sCodInsEmpresa & .sCodConvBanco & FormataCpoTexto("", 25)
                sRegHeader = sRegHeader & .sNomeEmpresa & .sMensagem1 & FormataCpoTexto("", 10) & "1" & .sDataGravRem & FormataCpoTexto("", 6) & .sNumRemessa & .sNVersaoLayLote & FormataCpoTexto("", 74)
                
            Case 104
                If gobjCobrancaEletronica.sVersaoLayout = "050" Then
                    sRegHeader = sRegHeader & FormataCpoNum(0, 20) & .sAgMantConta & .sDVAgencia & .sCodConvBanco & FormataCpoNum(0, 7 + 1)
                Else 'versao 030
                    sRegHeader = sRegHeader & .sCodConvBanco & FormataCpoTexto("", 4) & .sAgMantConta & .sDVAgencia & .sNumContaCorrente & .sDVConta & .sDVAgenciaConta
                End If
                
                sRegHeader = sRegHeader & .sNomeEmpresa & .sMensagem1 & .sMensagem2 & "1" & .sDataGravRem & "120000" & .sNumRemessa & .sNVersaoLayLote & "00000" & FormataCpoTexto("", 20) & .sReservadoEmpresa & FormataCpoTexto("", 29)
            
            Case 409
                sRegHeader = sRegHeader & .sCodConvBanco & .sAgMantConta & .sDVAgencia & .sNumContaCorrente & .sDVConta & .sDVAgenciaConta
                sRegHeader = sRegHeader & .sNomeEmpresa & .sMensagem1 & .sMensagem2 & "1" & .sDataGravRem & "120000" & .sNumRemessa & .sNVersaoLayLote & "01600" & FormataCpoTexto("", 20 + 20 + 29)
            
            Case 479
                sRegHeader = sRegHeader & "1" & .sCodConvBanco & "014        " & .sAgMantConta & .sDVAgencia & "000000" & .sNumContaCorrente
                sRegHeader = sRegHeader & .sNomeEmpresa & .sMensagem1 & .sMensagem2 & "1" & .sDataGravRem & "120000" & .sNumRemessa & .sNVersaoLayLote & "01600BKB001" & FormataCpoTexto("", 14 + 20 + 17 + 2 + 10)
        
            Case 275, 356
                sRegHeader = sRegHeader & .sCodConvBanco & .sAgMantConta & .sDVAgencia & .sNumContaCorrente & .sDVConta & .sDVAgenciaConta
                sRegHeader = sRegHeader & .sNomeEmpresa & .sMensagem1 & .sMensagem2 & "1" & .sDataGravRem & "120000" & .sNumRemessa & .sNVersaoLayLote & "00000" & FormataCpoTexto("", 54) & "000" & "  " & FormataCpoTexto("", 10)
        
            Case 399
                sRegHeader = sRegHeader & .sCodConvBanco & .sAgMantConta & "0" & .sNumContaCorrente & .sDVConta & "0"
                sRegHeader = sRegHeader & .sNomeEmpresa & .sMensagem1 & .sMensagem2 & "1" & .sDataGravRem & "120000" & .sNumRemessa & .sNVersaoLayLote & "00000" & "N" & FormataCpoTexto("", 11 + 8 + 20 + 29)
        
            Case 341
                sRegHeader = sRegHeader & .sCodConvBanco & .sAgMantConta & .sDVAgencia & .sNumContaCorrente & " " & .sDVConta
                sRegHeader = sRegHeader & .sNomeEmpresa & .sMensagem1 & .sMensagem2 & "1" & .sDataGravRem & "120000" & .sNumRemessa & .sNVersaoLayLote & "00000" & FormataCpoTexto("", 54) & "000" & FormataCpoTexto("", 12)
                
            Case 745
                sRegHeader = sRegHeader & .sCodConvBanco & .sAgMantConta & .sDVAgencia & .sNumContaCorrente & .sDVConta & .sDVAgenciaConta
                sRegHeader = sRegHeader & .sNomeEmpresa & .sMensagem1 & .sMensagem2 & "1" & .sDataGravRem & "120000" & .sNumRemessa & .sNVersaoLayLote & "01600" & FormataCpoTexto("", 20 + 20 + 29)
            
            Case 246
                sRegHeader = sRegHeader & .sCodConvBanco & .sAgMantConta & .sDVAgencia & .sNumContaCorrente & .sDVConta & .sDVAgenciaConta
                sRegHeader = sRegHeader & .sNomeEmpresa & .sMensagem1 & .sMensagem2 & "1" & .sDataGravRem & "120000" & .sNumRemessa & .sNVersaoLayLote & "00000" & FormataCpoTexto("", 54) & "000" & "  " & FormataCpoTexto("", 10)
            
            Case 41
                sRegHeader = sRegHeader & .sCodConvBanco & .sAgMantConta & .sDVAgencia & .sNumContaCorrente & .sDVConta & .sDVAgenciaConta
                sRegHeader = sRegHeader & .sNomeEmpresa & .sMensagem1 & .sMensagem2 & "1" & .sDataGravRem & "120000" & .sNumRemessa & .sNVersaoLayLote & "00000" & FormataCpoTexto("", 8) & "BE" & FormataCpoTexto("", 10 + 20 + 11) & "CSP" & "000" & "  " & FormataCpoTexto("", 10)
            
            Case 453, 749
                sRegHeader = sRegHeader & .sCodConvBanco & .sAgMantConta & .sDVAgencia & .sNumContaCorrente & .sDVConta & .sDVAgenciaConta
                sRegHeader = sRegHeader & .sNomeEmpresa & .sMensagem1 & .sMensagem2 & "1" & .sDataGravRem & "120000" & .sNumRemessa & .sNVersaoLayLote & "00000" & FormataCpoTexto("", 20 + 20 + 29)
       
            Case 422
                sRegHeader = sRegHeader & .sCodConvBanco & .sAgMantConta & .sDVAgencia & .sNumContaCorrente & .sDVConta & .sDVAgenciaConta
                sRegHeader = sRegHeader & .sNomeEmpresa & .sMensagem1 & .sMensagem2 & "1" & .sDataGravRem & "120000" & .sNumRemessa & .sNVersaoLayLote & "00000" & FormataCpoTexto("", 20 + 20 + 29)
       
            Case 748 'Sicred
                sRegHeader = sRegHeader & .sCodConvBanco & .sAgMantConta & .sDVAgencia & .sNumContaCorrente & .sDVConta & .sDVAgenciaConta
                sRegHeader = sRegHeader & .sNomeEmpresa & .sMensagem1 & .sMensagem2 & "1" & .sDataGravRem & "120000" & .sNumRemessa & .sNVersaoLayLote & "00000" & FormataCpoTexto("", 20 + 20 + 11) & "CSP" & "000" & "  " & FormataCpoTexto("", 10)
            
            Case 637 'Sofisa
                sRegHeader = sRegHeader & .sCodConvBanco & .sAgMantConta & .sDVAgencia & .sNumContaCorrente & .sDVConta & .sDVAgenciaConta
                sRegHeader = sRegHeader & .sNomeEmpresa & .sMensagem1 & .sMensagem2 & "1" & .sDataGravRem & "120000" & .sNumRemessa & .sNVersaoLayLote & "00000" & FormataCpoTexto("", 20 + 20 + 29)
       
            Case 85 'Cecred
                sRegHeader = sRegHeader & .sCodConvBanco & .sAgMantConta & .sDVAgencia & .sNumContaCorrente & .sDVConta & .sDVAgenciaConta
                sRegHeader = sRegHeader & .sNomeEmpresa & .sMensagem1 & .sMensagem2 & "1" & .sDataGravRem & "120000" & .sNumRemessa & .sNVersaoLayLote & "00000" & FormataCpoTexto("", 20 + 20 + 29)
       
        End Select
        
    End With
    
    Montar_RegistroHeaderArquivo_BcoBrasil = SUCESSO

    Exit Function

Erro_Montar_RegistroHeaderArquivo_BcoBrasil:

    Montar_RegistroHeaderArquivo_BcoBrasil = Err

    Select Case Err

        Case 51813

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147099)

    End Select

    Exit Function

End Function

Private Function Montar_RegistroTrailerArq_BcoBrasil(sRegTrailer As String) As Long
'Monta o registro

    glNumRegistros = glNumRegistros + 1
        
    sRegTrailer = FormataCpoNum(gobjCobrancaEletronica.objCobrador.iCodBanco, 3) & "99999         " & FormataCpoNum(giNumLotes, 6) & FormataCpoNum(glNumRegistros, 6)
    
    '??? com variacao de acordo com o banco
    Select Case gobjCobrancaEletronica.objCobrador.iCodBanco
    
        Case 8, 33, 353
             sRegTrailer = sRegTrailer & FormataCpoTexto("", 6)

        Case 104
             sRegTrailer = sRegTrailer & FormataCpoTexto("", 6)

        Case Else
             sRegTrailer = sRegTrailer & FormataCpoNum(0, 6)
             
    End Select
    
    sRegTrailer = sRegTrailer & FormataCpoTexto("", 205)
    
    Montar_RegistroTrailerArq_BcoBrasil = SUCESSO

    Exit Function

Erro_Montar_RegistroTrailerArq_BcoBrasil:

    Montar_RegistroTrailerArq_BcoBrasil = Err

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147100)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147101)
     
    End Select
     
    Exit Function

End Function

Private Sub Calcula_DV_AgCta_104(ByVal sTexto As String, sDVTexto As String)

Dim iIndice As Integer
Dim iTamTexto As Integer
Dim iDigito1 As Integer
Dim iVarModulo As Integer
Dim iProduto As Integer
Dim iSoma As Integer
Dim iResto As Integer
Dim iResultado As Integer
Dim iModulo As Integer, iBase As Integer

    iModulo = 9
    iBase = 9
    iVarModulo = 2
    iSoma = 0
    iTamTexto = Len(sTexto)
    For iIndice = iTamTexto To 1 Step -1

        iDigito1 = StrParaInt(Mid(sTexto, iIndice, 1))
        iProduto = (iDigito1 * iVarModulo)
        
        iSoma = iSoma + iProduto
        
        iVarModulo = iVarModulo + 1
        
        If iVarModulo > iModulo Then iVarModulo = 2
    
    Next
    
    iResto = iSoma Mod 11
    
    iResultado = 11 - iResto
    
    '??? alteracao
    If iResultado > 9 Then iResultado = 0
    
    sDVTexto = FormataCpoNum(iResultado, 1)
    
    Exit Sub
    
End Sub

Function Calcula_DV11(ByVal sSequencia As String, ByVal iBase As Integer, sDigito As String, Optional bAceitaP As Boolean = False) As Long
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

    Calcula_DV11 = 0

    Exit Function

Erro_Calcula_DV11:

    Calcula_DV11 = Err
    
    Select Case Err
        
        Case Else
            Call MsgBox(Error$)
    
    End Select
    
    Exit Function

End Function

Function BorderoPagto_ProcessarDetPag(ByVal objDetArqCNABPag As ClassDetArqCNABPag) As Long

Dim lErro As Long

On Error GoTo Erro_BorderoPagto_ProcessarDetPag

    Select Case objDetArqCNABPag.sFormaPagamento
    
        Case "11" 'se pagamento de contas de concessionarias (luz, telefone,...)
    
            lErro = Montar_Registro_SegmentoO(objDetArqCNABPag)
            If lErro <> SUCESSO Then gError 193338

        Case "10", "03", "01", "41" 'OP, DOC ou deposito em conta ou TED
    
            'Monta um registro de Segmento A
            lErro = Montar_Registro_SegmentoA(objDetArqCNABPag)
            If lErro <> SUCESSO Then gError 51729
                
            'Monta um registro de Segmento B
            lErro = Montar_Registro_SegmentoB(objDetArqCNABPag)
            If lErro <> SUCESSO Then gError 51729
            
        Case "30", "31" 'liquidacao de titulos do proprio banco ou de outro
        
            'Monta um registro de Segmento J
            lErro = Montar_Registro_SegmentoJ(objDetArqCNABPag)
            If lErro <> SUCESSO Then gError 51730
    
            'santander
            If gobjGeracaoArqCNABPag.objCCI.iCodBanco = 33 Then
            
                'Monta um registro de Segmento J52
                lErro = Montar_Registro_SegmentoJ52(objDetArqCNABPag)
                If lErro <> SUCESSO Then gError 51730
        
            End If
        
        Case Else
        
            gError 184293
    
    End Select
    
    BorderoPagto_ProcessarDetPag = SUCESSO
     
    Exit Function
    
Erro_BorderoPagto_ProcessarDetPag:

    BorderoPagto_ProcessarDetPag = gErr
     
    Select Case gErr
          
        Case 51729, 51730, 193338
        
        Case 184293
            lErro = Rotina_Erro(vbOKOnly, "ERRO_BORDPAG_FORMAPAGTO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147102)
     
    End Select
     
    Exit Function

End Function

Function BorderoPagto_Criar_ArquivoCNAB_240(objGeracaoArqCNABPag As ClassGeracaoArqCNABPag) As Long
'Cria o arquivo de remessa de títulos a pagar

Dim lErro As Long
Dim sNomeArquivo As String
Dim lTransacao As Long
Dim iIndice As Integer
Dim objDetArqCNABPag As ClassDetArqCNABPag
Dim alComando(0 To 2) As Long
Dim objLote As ClassLoteArqCNABPag
Dim objArq As New ClassArqCNABPag

On Error GoTo Erro_BorderoPagto_Criar_ArquivoCNAB_240

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 51721
    Next

    'Abre a transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 51722
    
    Call Zera_Var
    
    Set gobjGeracaoArqCNABPag = objGeracaoArqCNABPag
    Set gobjGeracaoArqCNABPag.objArq = objArq
    Set gColFiliaisEmpresa = New Collection
    
    'Gera o nome do arquivo
    lErro = CF("PagamentoEletronico_Gera_NomeArquivo", sNomeArquivo, gobjGeracaoArqCNABPag)
    If lErro <> SUCESSO Then gError 51735
 
    'Cria o arquivo de remessa
    Open sNomeArquivo For Output As #1
 
    'faz o lock no bordero de pagamento
    lErro = CF("BorderoPagto_Lock", alComando(0), gobjGeracaoArqCNABPag.objBorderoPagto.lNumIntBordero)
    If lErro <> SUCESSO Then gError 51723
    
    lErro = objArq.Iniciar(Me, gobjGeracaoArqCNABPag.objBorderoPagto.lNumIntBordero)
    If lErro <> SUCESSO Then gError 51724
    
    'Para cada baixa
    For Each objDetArqCNABPag In objArq.colDetArqCNABPag
        
        If objLote Is Nothing Then
        
            Set objLote = New ClassLoteArqCNABPag
            
            lErro = objLote.Iniciar(objDetArqCNABPag, objArq)
            If lErro <> SUCESSO Then gError 51730
            
        Else
        
            If objLote.Trocou(objDetArqCNABPag) Then
            
                lErro = objLote.Reiniciar(objDetArqCNABPag)
                If lErro <> SUCESSO Then gError 51730
            
            End If
            
        End If
        
        lErro = objLote.ProcessarDetalhe(objDetArqCNABPag)
        If lErro <> SUCESSO Then gError 51730
        
        Call objGeracaoArqCNABPag.objTelaAtualizacao.Mostra_Evolucao(True, 1)
    
    Next

    lErro = objLote.Finalizar
    If lErro <> SUCESSO Then gError 51731
    
    lErro = objArq.Finalizar
    If lErro <> SUCESSO Then gError 51732
    
    'Fecha arquivo
    Close #1
    
    'Atualiza em BorderosPagto o Nome e o sequencial do arquivo
    lErro = Comando_ExecutarPos(alComando(1), "UPDATE BorderosPagto SET NomeArq = ? , NumArqRemessa = ?", alComando(0), sNomeArquivo, gobjGeracaoArqCNABPag.lSequencialArq)
    If lErro <> AD_SQL_SUCESSO Then gError 51733
    
    'Faz o commit na transacao
    lErro = Transacao_Commit()
    If lErro <> SUCESSO Then gError 51734
        
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    BorderoPagto_Criar_ArquivoCNAB_240 = SUCESSO
    
    Exit Function
    
Erro_BorderoPagto_Criar_ArquivoCNAB_240:

    BorderoPagto_Criar_ArquivoCNAB_240 = gErr
    
    Select Case gErr
    
        Case 51721
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 51722
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 51723, 51724, 51725, 51726, 51728, 51729, 51730, 51731, 51732, 51735
        
        Case 51733
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_BORDERO_PAGTO", gErr)
        
        Case 51734
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147103)
            
    End Select
    
    'Faz um RollBack na transacao
    Call Transacao_Rollback
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    'Fecha arquivo
    Close #1
    
    Exit Function
    
End Function
        
Function Montar_RegistroHeaderCNABPag(ByVal objArq As ClassArqCNABPag) As Long
'Monta o registro de header do arquivo de contas a pagar

Dim lErro As Long, sRegistro As String
Dim tRegHeaderArq As typePagto240HeaderArq
Dim objFilialEmpresa As New AdmFiliais, sConvenio As String

On Error GoTo Erro_Montar_RegistroHeaderCNABPag

    Call objArq.ObterProxReg

    'Busca os dados da nossa filial empresa
    lErro = Obter_FilialEmpresa(giFilialEmpresa, objFilialEmpresa)
    If lErro <> SUCESSO Then gError 51736
    
    'Recolhe os dados do cabeçalho
    tRegHeaderArq.sNumeroBanco = FormataCpoNum(gobjGeracaoArqCNABPag.objCCI.iCodBanco, 3)
    tRegHeaderArq.sCodigoLote = "0000"
    tRegHeaderArq.sTipoRegistro = "0"
    tRegHeaderArq.sComplemento1 = FormataCpoTexto("", 9)
    If Len(objFilialEmpresa.sCgc) = 14 Then
        tRegHeaderArq.sTipoInscEmpresa = "2"
    Else
        tRegHeaderArq.sTipoInscEmpresa = "1"
    End If
    tRegHeaderArq.sCGCEmpresaDebitada = FormataCpoNum(objFilialEmpresa.sCgc, 14)
    tRegHeaderArq.sAgencia = FormataCpoNum(gobjGeracaoArqCNABPag.objCCI.sAgencia, 5)
    tRegHeaderArq.sDVAgencia = FormataCpoTexto(gobjGeracaoArqCNABPag.objCCI.sDVAgencia, 1)
    tRegHeaderArq.sConta = FormataCpoNum(gobjGeracaoArqCNABPag.objCCI.sNumConta, 12)
    
    tRegHeaderArq.sNomeEmpresa = FormataCpoTexto(gsNomeEmpresa, 30)
    
    Call CF("Empresa_Le_Nome_Customizado", tRegHeaderArq.sNomeEmpresa)
    
    tRegHeaderArq.sComplemento2 = FormataCpoTexto("", 10)
    tRegHeaderArq.sArquivoCodigo = "1"
    tRegHeaderArq.sDataGeracao = Format(gdtDataHoje, "DDMMYYYY")
    tRegHeaderArq.sHoraGeracao = Format(Time, "HHMMSS")
    tRegHeaderArq.sSequencialArquivo = FormataCpoNum(gobjGeracaoArqCNABPag.lSequencialArq, 6)
    tRegHeaderArq.sUnidadeDensidade = FormataCpoNum("1600", 5)
    tRegHeaderArq.sComplemento3 = FormataCpoTexto("", 69)
    
    Select Case gobjGeracaoArqCNABPag.objCCI.iCodBanco
    
        Case 1
            tRegHeaderArq.sDVNumConta = FormataCpoTexto(gobjGeracaoArqCNABPag.objCCI.sDVNumConta, 1)
            tRegHeaderArq.sDVAgConta = FormataCpoTexto(gobjGeracaoArqCNABPag.objCCI.sDVAgConta, 1)
            tRegHeaderArq.sNomeBanco = FormataCpoTexto("BANCO DO BRASIL", 30)
            tRegHeaderArq.sConvenio = FormataCpoTexto(gobjGeracaoArqCNABPag.objCCI.sConvenioPagto, 20)
            tRegHeaderArq.sLayoutArq = "030"
        
        Case 275, 356
            tRegHeaderArq.sDVNumConta = FormataCpoTexto(gobjGeracaoArqCNABPag.objCCI.sDVNumConta, 1)
            tRegHeaderArq.sDVAgConta = FormataCpoTexto(gobjGeracaoArqCNABPag.objCCI.sDVAgConta, 1)
            tRegHeaderArq.sNomeBanco = FormataCpoTexto("ABN AMRO SA", 30)
            sConvenio = FormataCpoNum(gobjGeracaoArqCNABPag.objCCI.sAgencia, 4) & FormataCpoNum(gobjGeracaoArqCNABPag.objCCI.sNumConta, 7) & "PG"
            tRegHeaderArq.sConvenio = FormataCpoTexto(sConvenio, 20)
            tRegHeaderArq.sLayoutArq = "040"
            
        Case 341
            tRegHeaderArq.sDVNumConta = FormataCpoTexto(gobjGeracaoArqCNABPag.objCCI.sDVAgConta, 1)
            tRegHeaderArq.sDVAgConta = FormataCpoTexto(gobjGeracaoArqCNABPag.objCCI.sDVNumConta, 1)
            tRegHeaderArq.sNomeBanco = FormataCpoTexto("ITAU", 30)
            tRegHeaderArq.sConvenio = FormataCpoTexto("", 20)
            tRegHeaderArq.sSequencialArquivo = FormataCpoNum(0, 6)
            tRegHeaderArq.sLayoutArq = "000"
            tRegHeaderArq.sComplemento1 = FormataCpoTexto("", 6) + "080"
            tRegHeaderArq.sUnidadeDensidade = FormataCpoNum("0", 5)
            
        Case 399
            tRegHeaderArq.sDVNumConta = FormataCpoTexto(gobjGeracaoArqCNABPag.objCCI.sDVNumConta, 1)
            tRegHeaderArq.sDVAgConta = FormataCpoTexto(gobjGeracaoArqCNABPag.objCCI.sDVAgConta, 1)
            tRegHeaderArq.sNomeBanco = FormataCpoTexto("HSBC BANK BRASIL SA", 30)
            tRegHeaderArq.sConvenio = FormataCpoTexto(gobjGeracaoArqCNABPag.objCCI.sConvenioPagto, 20)
            tRegHeaderArq.sLayoutArq = "020"
            tRegHeaderArq.sComplemento3 = FormataCpoTexto("CPGY2K", 69)
            
        Case 33, 353, 8
            tRegHeaderArq.sDVNumConta = FormataCpoTexto(gobjGeracaoArqCNABPag.objCCI.sDVNumConta, 1)
            tRegHeaderArq.sDVAgConta = FormataCpoTexto(gobjGeracaoArqCNABPag.objCCI.sDVAgConta, 1)
            tRegHeaderArq.sNomeBanco = FormataCpoTexto("Banco Santander", 30)
            tRegHeaderArq.sConvenio = FormataCpoTexto(gobjGeracaoArqCNABPag.objCCI.sConvenioPagto, 20)
            tRegHeaderArq.sLayoutArq = "060"
            tRegHeaderArq.sUnidadeDensidade = FormataCpoNum("0", 5)
            
        Case Else
            'falta conferir o layout dos outros
            
            'tRegHeaderArq.sNomeBanco = FormataCpoTexto("ABN AMRO SA", 30)
            'tRegHeaderArq.sConvenio = FormataCpoTexto(sConvenio, 20) '??? obter codigo do convenio
            
            gError 130412
            
    End Select
    
    'Forma o registro com os dados recolhidos p\ o cabeçalho
    With tRegHeaderArq
        sRegistro = .sNumeroBanco & .sCodigoLote & .sTipoRegistro & .sComplemento1 & .sTipoInscEmpresa & .sCGCEmpresaDebitada & .sConvenio
        sRegistro = sRegistro & .sAgencia & .sDVAgencia & .sConta
        sRegistro = sRegistro & .sDVNumConta & .sDVAgConta & .sNomeEmpresa & .sNomeBanco & .sComplemento2 & .sArquivoCodigo & .sDataGeracao
        sRegistro = sRegistro & .sHoraGeracao & .sSequencialArquivo & .sLayoutArq & .sUnidadeDensidade & .sComplemento3
    End With
    
    'Inclui no arquivo
    If Len(sRegistro) <> 240 Then gError 206433
    Print #1, DesacentuaTexto(sRegistro)
    
    Montar_RegistroHeaderCNABPag = SUCESSO
    
    Exit Function
    
Erro_Montar_RegistroHeaderCNABPag:

    Montar_RegistroHeaderCNABPag = gErr
    
    Select Case gErr
    
        Case 130412
            Call Rotina_Erro(vbOKOnly, "ERRO_CONFERIR_VERSAO_LAYOUT_CNAB", gErr)
            
        Case 51736
            
        Case 206433
            Call Rotina_Erro(vbOKOnly, "ERRO_REG_HEADER_ARQ_TAM_ERRADO", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147104)
            
    End Select
    
    Exit Function
    
End Function

Function Montar_Registro_HeaderLote(ByVal objLote As ClassLoteArqCNABPag) As Long
'Monta o registro de header de um lote de contas a pagar

Dim lErro As Long, sRegistro As String
Dim tRegHeaderLote As typePagto240HeaderLote
Dim objFilialEmpresa As New AdmFiliais, sConvenio As String

On Error GoTo Erro_Montar_Registro_HeaderLote
    
    Call objLote.ObterProxReg
    
    'Busca os dados da filialempresa
    lErro = Obter_FilialEmpresa(giFilialEmpresa, objFilialEmpresa)
    If lErro <> SUCESSO Then gError 51737

    'Recolhe os dados do header do Lote
    tRegHeaderLote.sNumeroBanco = FormataCpoNum(gobjGeracaoArqCNABPag.objCCI.iCodBanco, 3)
    tRegHeaderLote.sCodigoLote = FormataCpoNum(objLote.lSequencial, 4)
    tRegHeaderLote.sTipoRegistro = "1"
    tRegHeaderLote.sTipoOperacao = "C"
        
    tRegHeaderLote.sComplemento1 = FormataCpoTexto("", 1)
    If Len(objFilialEmpresa.sCgc) = 14 Then
        tRegHeaderLote.sTipoInscEmpresa = "2"
    Else
        tRegHeaderLote.sTipoInscEmpresa = "1"
    End If
    tRegHeaderLote.sCGCEmpresaDebitada = FormataCpoNum(objFilialEmpresa.sCgc, 14)
    
    'tRegHeaderLote.sComplemento2 = FormataCpoTexto("", 20)
    tRegHeaderLote.sComplemento2 = FormataCpoTexto("", 8)
    
    Select Case gobjGeracaoArqCNABPag.objCCI.iCodBanco
    
        Case 1
            If objLote.sFormaPagamento = "11" Then
                tRegHeaderLote.sTipoPagamento = "22"
            Else
                tRegHeaderLote.sTipoPagamento = "98"
            End If
            tRegHeaderLote.sFormaPagamento = objLote.sFormaPagamento
    
            tRegHeaderLote.sDVNumConta = FormataCpoTexto(gobjGeracaoArqCNABPag.objCCI.sDVNumConta, 1)
            tRegHeaderLote.sDVAgConta = FormataCpoTexto(gobjGeracaoArqCNABPag.objCCI.sDVAgConta, 1)
            tRegHeaderLote.sConvenio = FormataCpoTexto(gobjGeracaoArqCNABPag.objCCI.sConvenioPagto, 20)
            tRegHeaderLote.sLayoutLote = "020"
        
        Case 275, 356
            If objLote.sFormaPagamento = "11" Then
                tRegHeaderLote.sTipoPagamento = "22"
            Else
                tRegHeaderLote.sTipoPagamento = "98"
            End If
            tRegHeaderLote.sFormaPagamento = objLote.sFormaPagamento
            tRegHeaderLote.sDVNumConta = FormataCpoTexto(gobjGeracaoArqCNABPag.objCCI.sDVNumConta, 1)
            tRegHeaderLote.sDVAgConta = FormataCpoTexto(gobjGeracaoArqCNABPag.objCCI.sDVAgConta, 1)
            sConvenio = FormataCpoNum(gobjGeracaoArqCNABPag.objCCI.sAgencia, 4) & FormataCpoNum(gobjGeracaoArqCNABPag.objCCI.sNumConta, 7) & "PG"
            tRegHeaderLote.sConvenio = FormataCpoTexto(sConvenio, 20)
            tRegHeaderLote.sLayoutLote = "030"
            If objLote.sFormaPagamento = "11" Then tRegHeaderLote.sLayoutLote = "010"
            
        Case 341
            tRegHeaderLote.sTipoPagamento = "20"
            If objLote.sFormaPagamento = "11" Then
                tRegHeaderLote.sFormaPagamento = "13"
            Else
                tRegHeaderLote.sFormaPagamento = objLote.sFormaPagamento
            End If
            tRegHeaderLote.sDVNumConta = FormataCpoTexto(gobjGeracaoArqCNABPag.objCCI.sDVAgConta, 1)
            tRegHeaderLote.sDVAgConta = FormataCpoTexto(gobjGeracaoArqCNABPag.objCCI.sDVNumConta, 1)
            tRegHeaderLote.sConvenio = FormataCpoTexto("", 20)
            tRegHeaderLote.sLayoutLote = "040"
            
        Case 399
            '10 = OP, só que não existe no HSBC
            'Mudança para crédito administrativo
            If objLote.sFormaPagamento = "10" Then
                tRegHeaderLote.sFormaPagamento = "02"
            '41 = TED, só que não existe no HSBC
            'Mudança para DOC
            ElseIf objLote.sFormaPagamento = "41" Then
                tRegHeaderLote.sFormaPagamento = "03"
            Else
                '30 - Código de Barras Mesmo Banco
                '31 - Código de barras banco diferente
                '11 - Tributos
                '03 - DOC
                'Permanecem com a configuração original
                tRegHeaderLote.sFormaPagamento = objLote.sFormaPagamento
            End If
            If tRegHeaderLote.sFormaPagamento = "11" Then
                '22 - Pagamento de Contas,Tributos e Impostos barras e DARF/DARF Simples campo a campo
                tRegHeaderLote.sTipoPagamento = "22"
            Else
                '20 - Pagamento Fornecedores
                tRegHeaderLote.sTipoPagamento = "20"
            End If
            'tRegHeaderLote.sTipoPagamento = objLote.sTipoPagamento
    
            tRegHeaderLote.sDVNumConta = FormataCpoTexto(gobjGeracaoArqCNABPag.objCCI.sDVNumConta, 1)
            tRegHeaderLote.sDVAgConta = FormataCpoTexto(gobjGeracaoArqCNABPag.objCCI.sDVAgConta, 1)
            tRegHeaderLote.sConvenio = FormataCpoTexto(gobjGeracaoArqCNABPag.objCCI.sConvenioPagto, 20)
            tRegHeaderLote.sLayoutLote = "020"
            tRegHeaderLote.sComplemento2 = FormataCpoTexto("S", 8)
            
        Case 33, 353, 8
            'Não tem 41 para TED, TED e DOC são 03
            tRegHeaderLote.sFormaPagamento = objLote.sFormaPagamento
            If tRegHeaderLote.sFormaPagamento = "41" Then tRegHeaderLote.sFormaPagamento = "03"
            tRegHeaderLote.sTipoPagamento = "20"
            tRegHeaderLote.sDVNumConta = FormataCpoTexto(gobjGeracaoArqCNABPag.objCCI.sDVNumConta, 1)
            tRegHeaderLote.sDVAgConta = FormataCpoTexto(gobjGeracaoArqCNABPag.objCCI.sDVAgConta, 1)
            tRegHeaderLote.sConvenio = FormataCpoTexto(gobjGeracaoArqCNABPag.objCCI.sConvenioPagto, 20)
            tRegHeaderLote.sLayoutLote = "031"
            
        Case Else
            'falta conferir o layout dos outros
            
            'tRegHeaderArq.sConvenio = FormataCpoTexto(sConvenio, 20) '??? obter codigo do convenio
            
            gError 130413
            
    End Select
    
    tRegHeaderLote.sAgencia = FormataCpoNum(gobjGeracaoArqCNABPag.objCCI.sAgencia, 5)
    tRegHeaderLote.sDVAgencia = FormataCpoTexto(gobjGeracaoArqCNABPag.objCCI.sDVAgencia, 1)
    tRegHeaderLote.sConta = FormataCpoNum(gobjGeracaoArqCNABPag.objCCI.sNumConta, 12)
    
    tRegHeaderLote.sNomeEmpresa = FormataCpoTexto(gsNomeEmpresa, 30)
    
    Call CF("Empresa_Le_Nome_Customizado", tRegHeaderLote.sNomeEmpresa)
    
    tRegHeaderLote.sMensagem = FormataCpoTexto("", 40)

    tRegHeaderLote.sEnderecoEmpresa = FormataCpoTexto("", 30)
    tRegHeaderLote.sNumeroLocal = FormataCpoNum("", 5)
    tRegHeaderLote.sComplementoEndereco = FormataCpoTexto("", 15)
    tRegHeaderLote.sCidade = FormataCpoTexto(objFilialEmpresa.objEndereco.sCidade, 20)
    tRegHeaderLote.sCEP = Formata_Cep(objFilialEmpresa.objEndereco.sCEP, 8)
    tRegHeaderLote.sEstado = FormataCpoTexto(objFilialEmpresa.objEndereco.sSiglaEstado, 2)
    tRegHeaderLote.sOcorrenciasRetorno = FormataCpoTexto("", 10)
    
    'Monta o registro de acordo com os dados recolhidos
    With tRegHeaderLote
    
        'alteracao por tulio301202... a atribuicao estava sendo feita num momento improprio..
        'gerando problemas na hora de gravar o arquivo
        .sCGCEmpresaDebitada = FormataCpoNum(objFilialEmpresa.sCgc, 14)
    
        sRegistro = .sNumeroBanco & .sCodigoLote & .sTipoRegistro & "C" & .sTipoPagamento & .sFormaPagamento & .sLayoutLote
        sRegistro = sRegistro & .sComplemento1 & .sTipoInscEmpresa & .sCGCEmpresaDebitada
        sRegistro = sRegistro & .sConvenio & .sAgencia & .sDVAgencia & .sConta & .sDVNumConta & .sDVAgConta
        sRegistro = sRegistro & .sNomeEmpresa & .sMensagem & .sEnderecoEmpresa & .sNumeroLocal & .sComplementoEndereco
        sRegistro = sRegistro & .sCidade & .sCEP & .sEstado & .sComplemento2 & .sOcorrenciasRetorno

    End With
    
    'Inclui no arquivo
    If Len(sRegistro) <> 240 Then gError 206434
    Print #1, DesacentuaTexto(sRegistro)
       
    Montar_Registro_HeaderLote = SUCESSO
    
    Exit Function

Erro_Montar_Registro_HeaderLote:

    Montar_Registro_HeaderLote = gErr

    Select Case gErr
    
        Case 51737, 130413
            
        Case 206434
            Call Rotina_Erro(vbOKOnly, "ERRO_REG_HEADER_LOTE_TAM_ERRADO", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147105)
            
    End Select

    Exit Function

End Function

Function Montar_Registro_SegmentoA(ByVal objDet As ClassDetArqCNABPag) As Long
'Monta o registro de segmento A de pagto

Dim lErro As Long, sRegistro As String
Dim tRegSegmentoA As typePagto240SegmentoA, sDVAgConta As String
Dim sDVAgencia As String, sDVConta As String, sConta As String, sAgencia As String
Dim sParteErro As String

On Error GoTo Erro_Montar_Registro_SegmentoA

    If (Len(Trim(objDet.sAgenciaFavorecido)) = 0 Or Len(Trim(objDet.sContaFavorecido)) = 0 Or objDet.iBancoFavorecido = 0) Then gError 106575

    Call objDet.objLote.ObterProxReg
    
    'Recolhe os dados do Segmento A do detalhe do lote
    tRegSegmentoA.sNumeroBanco = FormataCpoNum(gobjGeracaoArqCNABPag.objCCI.iCodBanco, 3)
    tRegSegmentoA.sCodigoLote = FormataCpoNum(objDet.objLote.lSequencial, 4)
    tRegSegmentoA.sTipoRegistro = "3"
    tRegSegmentoA.sNumeroRegistro = FormataCpoNum(objDet.lSequencial, 5)
    tRegSegmentoA.sSegmento = "A"
    tRegSegmentoA.sTipoMovimento = "0"
    tRegSegmentoA.sCodigoInstrucaoMovto = "00"
    tRegSegmentoA.sCodCamaraCentralizadora = FormataCpoNum(0, 3)
    tRegSegmentoA.sBancoFavorecido = FormataCpoNum(objDet.iBancoFavorecido, 3)
    
sParteErro = "01-" & CStr(objDet.lNumTitulo) & "-" & objDet.sAgenciaFavorecido
        
    Call Extrair_Agencia_DV(objDet.sAgenciaFavorecido, sAgencia, 5, sDVAgencia, 1)
    
sParteErro = "02-" & CStr(objDet.lNumTitulo)

    Call Extrair_Conta_DV(objDet.sContaFavorecido, sConta, 12, sDVConta, 1)
'    sDVAgConta = " " '??? nao tenho
'    tRegSegmentoA.sAgenciaConta = sAgencia & sDVAgencia & sConta & sDVConta & sDVAgConta

sParteErro = "03-" & CStr(objDet.lNumTitulo)
    

    tRegSegmentoA.sNomeFavorecido = FormataCpoTexto(objDet.sNomeFavorecido, 30)
    tRegSegmentoA.sSeuNumero = FormataCpoTexto(CStr(objDet.lNumTitulo) & " " & CStr(objDet.iNumParcela) & " " & CStr(objDet.lFornecedor) & " " & CStr(objDet.iSeqBaixaParcPag), 20)
    tRegSegmentoA.sDataPagamento = Format(IIf(gobjGeracaoArqCNABPag.objBorderoPagto.dtDataVencimento >= gdtDataAtual, gobjGeracaoArqCNABPag.objBorderoPagto.dtDataVencimento, gdtDataAtual), "DDMMYYYY")
    tRegSegmentoA.sTipoMoeda = "BRL"
    tRegSegmentoA.sQtdeMoeda = FormataCpoNum("", 15)
    tRegSegmentoA.sValorPagamento = FormataCpoValor(objDet.dValorPagto, 15)
    tRegSegmentoA.sNossoNumero = FormataCpoTexto("", 20)
    tRegSegmentoA.sDataEfetiva = FormataCpoNum("", 8)
    tRegSegmentoA.sValorEfetivo = FormataCpoNum("", 15)
    tRegSegmentoA.sMensagem2 = FormataCpoTexto("", 40)
    tRegSegmentoA.sFinalidadeDetalhe = "03"
    tRegSegmentoA.sComplemento1 = FormataCpoTexto("", 10)
    tRegSegmentoA.sAviso = "0"
    tRegSegmentoA.sOcorrenciasRetorno = FormataCpoTexto("", 10)
    
    sDVAgConta = " " '??? nao tenho
    Select Case gobjGeracaoArqCNABPag.objCCI.iCodBanco
        
        Case 1
            tRegSegmentoA.sAgenciaConta = sAgencia & sDVAgencia & sConta & sDVConta & sDVAgConta
            tRegSegmentoA.sFinalidadeDetalhe = "  "
            
        Case 33, 353, 8
        
            tRegSegmentoA.sAgenciaConta = sAgencia & sDVAgencia & sConta & sDVConta & sDVAgConta
            tRegSegmentoA.sFinalidadeDetalhe = "03"

            If Not (objDet.objLote Is Nothing) Then
                Select Case objDet.objLote.sFormaPagamento
'                14. Código da Câmara Centralizadora
'                018 = TED CIP
'                810 = TED STR
'                700 = DOC
'                000 = CC
                    Case "41"
                        tRegSegmentoA.sCodCamaraCentralizadora = "018"
                    Case "03"
                        tRegSegmentoA.sCodCamaraCentralizadora = "700"
                End Select
            End If
        
        Case 341
            tRegSegmentoA.sAgenciaConta = sAgencia & " " & sConta & " " & sDVConta
            tRegSegmentoA.sMensagem2 = FormataCpoTexto("", 18 + 2) + FormataCpoNum("", 6 + 14)
            tRegSegmentoA.sFinalidadeDetalhe = "  "
            tRegSegmentoA.sTipoMoeda = "REA"
            
        Case 356
            tRegSegmentoA.sAgenciaConta = sAgencia & sDVAgencia & sConta & sDVConta & sDVAgConta
            tRegSegmentoA.sFinalidadeDetalhe = "03"

            If Not (objDet.objLote Is Nothing) Then
                If objDet.objLote.sFormaPagamento = "41" Or objDet.objLote.sFormaPagamento = "03" Or objDet.objLote.sFormaPagamento = "43" Then
                    tRegSegmentoA.sCodCamaraCentralizadora = "018"
                End If
            End If
            
        Case 399
            tRegSegmentoA.sAgenciaConta = sAgencia & " " & sConta & sDVConta & sDVAgConta
            tRegSegmentoA.sFinalidadeDetalhe = "03"

            If Not (objDet.objLote Is Nothing) Then
                If objDet.objLote.sFormaPagamento = "03" Then
                    tRegSegmentoA.sCodCamaraCentralizadora = "700"
                End If
            End If
        
        Case Else
            tRegSegmentoA.sAgenciaConta = sAgencia & sDVAgencia & sConta & sDVConta & sDVAgConta
            tRegSegmentoA.sFinalidadeDetalhe = "03"
            
    End Select

    'Monta o registro A
    With tRegSegmentoA
        
        sRegistro = .sNumeroBanco & .sCodigoLote & .sTipoRegistro & .sNumeroRegistro & .sSegmento & .sTipoMovimento
        sRegistro = sRegistro & .sCodigoInstrucaoMovto & .sCodCamaraCentralizadora & .sBancoFavorecido & .sAgenciaConta & .sNomeFavorecido & .sSeuNumero
        sRegistro = sRegistro & .sDataPagamento & .sTipoMoeda & .sQtdeMoeda & .sValorPagamento & .sNossoNumero
        sRegistro = sRegistro & .sDataEfetiva & .sValorEfetivo & .sMensagem2 & .sFinalidadeDetalhe & .sComplemento1
        sRegistro = sRegistro & .sAviso & .sOcorrenciasRetorno
    
    End With

    'Inclui no arquivo
    If Len(sRegistro) <> 240 Then gError 206435
    Print #1, DesacentuaTexto(sRegistro)
       
    Montar_Registro_SegmentoA = SUCESSO
    
    Exit Function
    
Erro_Montar_Registro_SegmentoA:

    Montar_Registro_SegmentoA = gErr
    
    Call Rotina_Erro(vbOKOnly, sParteErro, gErr)
    
    Select Case gErr
    
        Case 106575
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORN_SEM_INFO_CTA", gErr, objDet.lFornecedor, objDet.iFilialForn)
        
        Case 206435
            Call Rotina_Erro(vbOKOnly, "ERRO_REG_DET_A_TAM_ERRADO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147106)
    
    End Select
    
    Exit Function
    
End Function


Function Montar_Registro_SegmentoJ(ByVal objDet As ClassDetArqCNABPag) As Long
'Recolhe os dados do Segmento J do detalhe do lote

Dim lErro As Long, sRegistro As String, sCodigoDeBarras As String
Dim tRegSegmentoJ As typePagto240SegmentoJ, bDigitacao As Boolean, iDVTeste As Integer

On Error GoTo Erro_Montar_Registro_SegmentoJ

    Call objDet.objLote.ObterProxReg
    
    sCodigoDeBarras = objDet.sCodigoDeBarras
 '??? apenas para teste
'    If Len(Trim(sCodigoDeBarras)) < 33 Then sCodigoDeBarras = "00198287480000000000000020571212219410000016500"
    
    If Len(Trim(sCodigoDeBarras)) < 33 Then gError 106574

    bDigitacao = True
    If Len(Trim(sCodigoDeBarras)) = 44 Then
    
        Call CalculaDVGeralMod11(Mid(sCodigoDeBarras, 1, 4) & Mid(sCodigoDeBarras, 6), iDVTeste)
    
        If iDVTeste = CInt(Mid(sCodigoDeBarras, 5, 1)) Then bDigitacao = False
        
    End If
    
    'Recolhe os dados do Segmento J do detalhe do lote
    With tRegSegmentoJ
    
        .sNumeroBanco = FormataCpoNum(gobjGeracaoArqCNABPag.objCCI.iCodBanco, 3)
        .sCodigoLote = FormataCpoNum(objDet.objLote.lSequencial, 4)
        .sTipoRegistro = "3"
        .sNumeroRegistro = FormataCpoNum(objDet.lSequencial, 5)
        .sSegmento = "J"
        .sTipoMovimento = "000"
        .sBancoFavorecido = Mid(sCodigoDeBarras, 1, 3)
        .sMoeda = "9"
        
        If bDigitacao = True Then
            .sDV = Mid(sCodigoDeBarras, 33, 1)
            If .sDV = "" Then .sDV = "0"
            If Len(Trim(sCodigoDeBarras)) <> 47 Then
                .sValor = FormataCpoValor(StrParaDbl(Mid(sCodigoDeBarras, 34, 14)) / 100, 14)
            Else
                .sValor = Mid(sCodigoDeBarras, 34, 14)
            End If
            .sCampoLivre = Mid(sCodigoDeBarras, 5, 5) & Mid(sCodigoDeBarras, 11, 10) & Mid(sCodigoDeBarras, 22, 10)
        Else
            .sDV = Mid(sCodigoDeBarras, 5, 1)
            .sValor = Mid(sCodigoDeBarras, 6, 14)
            .sCampoLivre = Mid(sCodigoDeBarras, 20, 25)
        End If
        
        .sNomeFavorecido = FormataCpoTexto(objDet.sNomeFavorecido, 30)
        .sDataVencto = Format(objDet.dtDataVenctoParcela, "DDMMYYYY")
        .sValorTitulo = FormataCpoValor(objDet.dValorTitulo, 15)
        .sDescontos = FormataCpoValor(objDet.dValorDesconto, 15)
        .sAcrescimos = FormataCpoValor(objDet.dValorJuros + objDet.dValorMulta, 15)
        .sDataPagamento = Format(gobjGeracaoArqCNABPag.objBorderoPagto.dtDataVencimento, "DDMMYYYY")
        .sValorPagamento = FormataCpoValor(objDet.dValorPagto, 15)
        .sZeros = FormataCpoNum("", 15)
        .sSeuNumero = FormataCpoTexto(CStr(objDet.lNumTitulo) & " " & CStr(objDet.iNumParcela) & " " & CStr(objDet.lFornecedor) & " " & CStr(objDet.iSeqBaixaParcPag), 20)
        .sComplemento = FormataCpoTexto("", 13)
        .sNossoNumero = FormataCpoTexto("", 15)
        .sOcorrenciasRetorno = FormataCpoTexto("", 10)
    
        'Monta o registro J
        sRegistro = .sNumeroBanco & .sCodigoLote & .sTipoRegistro & .sNumeroRegistro & .sSegmento & .sTipoMovimento & .sBancoFavorecido & .sMoeda
        sRegistro = sRegistro & .sDV & .sValor & .sCampoLivre & .sNomeFavorecido & .sDataVencto & .sValorTitulo & .sDescontos & .sAcrescimos
        sRegistro = sRegistro & .sDataPagamento & .sValorPagamento & .sZeros & .sSeuNumero & .sComplemento & .sNossoNumero & .sOcorrenciasRetorno

    End With
    
    'Inclui no arquivo
    If Len(sRegistro) <> 240 Then gError 206437
    Print #1, DesacentuaTexto(sRegistro)
       
    Montar_Registro_SegmentoJ = SUCESSO
    
    Exit Function
    
Erro_Montar_Registro_SegmentoJ:

    Montar_Registro_SegmentoJ = gErr
    
    Select Case gErr
    
        Case 206437
            Call Rotina_Erro(vbOKOnly, "ERRO_REG_DET_J_TAM_ERRADO", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_TITPAG_SEM_CODBARRAS", gErr, objDet.lNumTitulo, objDet.iNumParcela, objDet.lFornecedor)
        
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147107)
            
    End Select
    
    Exit Function
    
End Function

Function Montar_Registro_TrailerLote(ByVal objLote As ClassLoteArqCNABPag) As Long
'MOnta o regsitro de trailer do lote

Dim lErro As Long, sRegistro As String
Dim tRegTrailerLote As typePagto240TrailerLote

On Error GoTo Erro_Montar_Registro_TrailerLote

    Call objLote.ObterProxReg
    
    'Reune os dados do trailer
    With tRegTrailerLote
    
        .sNumeroBanco = FormataCpoNum(gobjGeracaoArqCNABPag.objCCI.iCodBanco, 3)
        .sCodigoLote = FormataCpoNum(objLote.lSequencial, 4)
        .sTipoRegistro = "5"
        .sComplemento1 = FormataCpoTexto("", 9)
        .sTotalQtdRegistros = FormataCpoNum(objLote.lQuantRegistros, 6)
        .sTotalValorPagtos = FormataCpoValor(objLote.dValorTotal, 18)
        .sZeros = FormataCpoNum("", 18)
        .sComplemento2 = FormataCpoTexto("", 171)
        .sOcorrenciasRetorno = FormataCpoTexto("", 10)
        
        'Monta o registro de trailer de lote
        sRegistro = .sNumeroBanco & .sCodigoLote & .sTipoRegistro & .sComplemento1 & .sTotalQtdRegistros & .sTotalValorPagtos & .sZeros & .sComplemento2 & .sOcorrenciasRetorno
    
    End With

    'Inclui no arquivo
    Print #1, DesacentuaTexto(sRegistro)
       
    Montar_Registro_TrailerLote = SUCESSO

    Exit Function

Erro_Montar_Registro_TrailerLote:

    Montar_Registro_TrailerLote = Err
    
    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147108)
            
    End Select

    Exit Function

End Function

Function Montar_Registro_TrailerArq(objArq As ClassArqCNABPag) As Long
'MOnta o regsitro de trailer do arquivo de pagtos

Dim lErro As Long, sRegistro As String
Dim tRegTrailerArq As typePagto240TrailerArq

On Error GoTo Erro_Montar_Registro_TrailerArq

    Call objArq.ObterProxReg
    
    'Reune os dados do trailer
    With tRegTrailerArq
        
        .sNumeroBanco = FormataCpoNum(gobjGeracaoArqCNABPag.objCCI.iCodBanco, 3)
        .sCodigoLote = "9999"
        .sTipoRegistro = "9"
        .sComplemento1 = FormataCpoTexto("", 9)
        .sTotalQtdLotes = FormataCpoNum(objArq.lSequencialLote, 6)
        .sTotalQtdRegistros = FormataCpoNum(objArq.lQuantRegistros, 6)
        
        If gobjGeracaoArqCNABPag.objCCI.iCodBanco <> 341 Then
            .sTotalContasConc = FormataCpoNum(0, 6)
        Else
            .sTotalContasConc = FormataCpoTexto("", 6)
        End If
    
        .sComplemento2 = FormataCpoTexto("", 205)
        
        'Monta o registro de trailer de lote
        sRegistro = .sNumeroBanco & .sCodigoLote & .sTipoRegistro & .sComplemento1 & .sTotalQtdLotes & .sTotalQtdRegistros & .sTotalContasConc & .sComplemento2
        
    End With

    'Inclui no arquivo
    Print #1, DesacentuaTexto(sRegistro)
       
    Montar_Registro_TrailerArq = SUCESSO
    
    Exit Function
    
Erro_Montar_Registro_TrailerArq:

    Montar_Registro_TrailerArq = Err
    
    Select Case Err
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147109)
            
    End Select
    
    Exit Function
    
End Function

Function Montar_Registro_SegmentoB(ByVal objDet As ClassDetArqCNABPag) As Long
'Monta o registro de segmento B de pagto

Dim lErro As Long, sRegistro As String, sEnderecoFavorecido As String
Dim tRegSegmentoB As typePagto240SegmentoB

On Error GoTo Erro_Montar_Registro_SegmentoB

    Call objDet.objLote.ObterProxReg
    
    'Recolhe os dados do Segmento B do detalhe do lote
    tRegSegmentoB.sNumeroBanco = FormataCpoNum(gobjGeracaoArqCNABPag.objCCI.iCodBanco, 3)
    tRegSegmentoB.sCodigoLote = FormataCpoNum(objDet.objLote.lSequencial, 4)
    tRegSegmentoB.sTipoRegistro = "3"
    tRegSegmentoB.sNumeroRegistro = FormataCpoNum(objDet.lSequencial, 5)
    tRegSegmentoB.sSegmento = "B"
    tRegSegmentoB.sComplemento1 = FormataCpoTexto("", 3)
    
    If Len(objDet.sCGCFavorecido) = 14 Then
        tRegSegmentoB.sTipoInscEmpresa = "2"
        'tRegSegmentoB.sNumInscFavorecido = objDet.sCGCFavorecido
    Else
        tRegSegmentoB.sTipoInscEmpresa = "1"
        'tRegSegmentoB.sNumInscFavorecido = Left(FormataCpoNum(objDet.sCGCFavorecido, 11), 9) & "000" & FormataCpoNum(Right(objDet.sCGCFavorecido, 2), 2)
    End If
    'Alterado por Wagner 02/05/2007
    tRegSegmentoB.sNumInscFavorecido = FormataCpoNum(objDet.sCGCFavorecido, 14)
    
    sEnderecoFavorecido = FormataCpoTexto(left(objDet.sEnderecoFavorecido, 50), 50)
    tRegSegmentoB.sEndereco = Mid(sEnderecoFavorecido, 1, 30)
    tRegSegmentoB.sNumeroLocal = Mid(sEnderecoFavorecido, 31, 5)
    tRegSegmentoB.sComplementoEndereco = Mid(sEnderecoFavorecido, 36, 15)
    tRegSegmentoB.sBairro = FormataCpoTexto(left(objDet.sBairroFavorecido, 15), 15)
    tRegSegmentoB.sCidade = FormataCpoTexto(left(objDet.sCidadeFavorecido, 20), 20)
    tRegSegmentoB.sCEP = left(objDet.sCEPFavorecido & "00000000", 8)
    tRegSegmentoB.sEstado = FormataCpoTexto(objDet.sEstadoFavorecido, 2)
    tRegSegmentoB.sDataVencimento = Format(objDet.dtDataVenctoParcela, "DDMMYYYY")
    tRegSegmentoB.sValorDocumento = FormataCpoValor(objDet.dValorPagto, 15) '???
    tRegSegmentoB.sValorAbatimento = FormataCpoValor(0, 15)
    tRegSegmentoB.sValorDesconto = FormataCpoValor(0, 15)
    tRegSegmentoB.sValorMora = FormataCpoValor(0, 15)
    tRegSegmentoB.sValorMulta = FormataCpoValor(0, 15)
    tRegSegmentoB.sCodDocFavorecido = FormataCpoTexto("", 15)
    tRegSegmentoB.sAvisoFavorecido = "0"
    tRegSegmentoB.sComplemento2 = FormataCpoTexto("", 14)
    
    Select Case gobjGeracaoArqCNABPag.objCCI.iCodBanco
        Case 1
            tRegSegmentoB.sAvisoFavorecido = " "
            
        Case 399
            tRegSegmentoB.sDataVencimento = ""
            tRegSegmentoB.sValorDocumento = ""
            tRegSegmentoB.sValorAbatimento = ""
            tRegSegmentoB.sValorDesconto = ""
            tRegSegmentoB.sValorMora = ""
            tRegSegmentoB.sValorMulta = ""
            tRegSegmentoB.sCodDocFavorecido = ""
            tRegSegmentoB.sAvisoFavorecido = ""
            tRegSegmentoB.sComplemento2 = FormataCpoTexto("", 113)
        
        Case 341
            tRegSegmentoB.sNumeroLocal = "00000"
            tRegSegmentoB.sComplementoEndereco = Mid(sEnderecoFavorecido, 31, 15)
            
        Case 33, 353, 8
            tRegSegmentoB.sCodDocFavorecido = ""
            tRegSegmentoB.sAvisoFavorecido = ""
            tRegSegmentoB.sComplemento2 = "0000" & FormataCpoTexto("", 11) & "00000" & FormataCpoTexto("", 10)
        
        Case Else
            tRegSegmentoB.sAvisoFavorecido = "0"
    End Select
    
    'Monta o registro B
    With tRegSegmentoB
        
        sRegistro = .sNumeroBanco & .sCodigoLote & .sTipoRegistro & .sNumeroRegistro & .sSegmento & .sComplemento1
        sRegistro = sRegistro & .sTipoInscEmpresa & .sNumInscFavorecido & .sEndereco & .sNumeroLocal & .sComplementoEndereco
        sRegistro = sRegistro & .sBairro & .sCidade & .sCEP & .sEstado & .sDataVencimento & .sValorDocumento & .sValorAbatimento
        sRegistro = sRegistro & .sValorDesconto & .sValorMora & .sValorMulta & .sCodDocFavorecido & .sAvisoFavorecido & .sComplemento2
    
    End With

    'Inclui no arquivo
    If Len(sRegistro) <> 240 Then gError 206436
    Print #1, DesacentuaTexto(sRegistro)
       
    Montar_Registro_SegmentoB = SUCESSO
    
    Exit Function
    
Erro_Montar_Registro_SegmentoB:

    Montar_Registro_SegmentoB = gErr
    
    Select Case gErr
    
        Case 106575
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORN_SEM_INFO_CTA", gErr, objDet.lFornecedor, objDet.iFilialForn)
        
        Case 206436
            Call Rotina_Erro(vbOKOnly, "ERRO_REG_DET_B_TAM_ERRADO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147110)
    
    End Select
    
    Exit Function
    
End Function

Private Function CalculaDVGeralMod11(ByVal sCodBarras As String, iDV As Integer) As Long
'Calcula o dv geral do codigo de barras, modulo 11
'sCodBarras eh a string que traz o codigo de barras (INPUT)
'iDv eh o inteiro que retorna o dv (OUTPUT)

Dim iCont As String
Dim lACC As Long
Dim iIndex As Integer
Dim iNumCorrente As Integer
Dim iIndice As Integer

On Error GoTo Erro_CalculaDVGeralMod11

    'inicializa a variavel multiplicadora
    iCont = 0
    
    'inicializa o acumulador
    lACC = 0
    
    'inicializa o indice do mod
    iIndice = 0

    'percorre o codbarras de traz pra frente
    For iIndex = Len(sCodBarras) To 1 Step -1
    
        'obtem o numero corrente
        iNumCorrente = StrParaInt(Mid(sCodBarras, iIndex, 1))
    
        'atualiza a variavel que tera de rodar a sequencia
        '234567892345678923456789.......
        iCont = (iIndice Mod 8) + 2
    
        'acumula o numero corrente * a variavel multiplicadora
        lACC = lACC + (iNumCorrente * iCont)
        
        'incrementa o indice
        iIndice = iIndice + 1
        
    Next

    'calcula finalmente o dv
    iDV = 11 - (lACC Mod 11)
    
    'se o dv for 10 ou 11, fazer dv = 1
    If iDV = 10 Or iDV = 11 Then iDV = 1
    
    CalculaDVGeralMod11 = SUCESSO

    Exit Function

Erro_CalculaDVGeralMod11:

    CalculaDVGeralMod11 = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147111)
            
    End Select

End Function

Private Function InstrucaoRem_ObterInfo(ByVal objOcorrRemParcRem As ClassOcorrRemParcRec) As Long
'obter valor cobrado, instrucao1, dias de protesto1, instrucao2, dias de protesto2 da remessa anterior do titulo

Dim lErro As Long
Dim lComando As Long
Dim tOcorrRemParcRec As typeTipoOcorrRemParcRec

On Error GoTo Erro_InstrucaoRem_ObterInfo

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 184161
    
    With tOcorrRemParcRec
        lErro = Comando_Executar(lComando, "SELECT Juros, Instrucao1, DiasDeProtesto1, Instrucao2, DiasDeProtesto2, ValorCobrado FROM OcorrenciasRemParcRec WHERE NumIntParc = ? AND Cobrador = ? AND NumSeqOCorr <= ? AND CodOcorrencia = 1 ORDER BY NumSeqOCorr DESC", _
            .dJuros, .iInstrucao1, .iDiasDeProtesto1, .iInstrucao2, .iDiasDeProtesto2, .dValorCobrado, objOcorrRemParcRem.lNumIntParc, objOcorrRemParcRem.iCobrador, objOcorrRemParcRem.iNumSeqOcorr)
    End With
    If lErro <> AD_SQL_SUCESSO Then gError 184162

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 184163
    If lErro = AD_SQL_SEM_DADOS Then
        gError 184164
    End If
    
    With tOcorrRemParcRec
        objOcorrRemParcRem.dJuros = .dJuros
        objOcorrRemParcRem.iInstrucao1 = .iInstrucao1
        objOcorrRemParcRem.iDiasDeProtesto1 = .iDiasDeProtesto1
        objOcorrRemParcRem.iInstrucao2 = .iInstrucao2
        objOcorrRemParcRem.iDiasDeProtesto2 = .iDiasDeProtesto2
        objOcorrRemParcRem.dValorCobrado = .dValorCobrado
    End With
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    InstrucaoRem_ObterInfo = SUCESSO
    
    Exit Function

Erro_InstrucaoRem_ObterInfo:

    InstrucaoRem_ObterInfo = gErr
    
    Select Case gErr
    
        Case 184161
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 184162, 184163
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PARCELAS_REC", gErr)

        Case 184164
            'nao encontrou
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184165)
    
    End Select
        
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function Montar_Registro_SegmentoO(ByVal objDet As ClassDetArqCNABPag) As Long
'Recolhe os dados do Segmento O do detalhe do lote

Dim lErro As Long, sRegistro As String, sCodigoDeBarras As String, i As Integer, iDV As Integer, sCodigoDeBarrasAux As String
Dim tRegSegmentoO As typePagto240SegmentoO, bDigitacao As Boolean, iDVTeste As Integer, sTrecho As String

On Error GoTo Erro_Montar_Registro_SegmentoO

    Call objDet.objLote.ObterProxReg
    
    sCodigoDeBarras = objDet.sCodigoDeBarras
 '??? apenas para teste
'    If Len(Trim(sCodigoDeBarras)) < 33 Then sCodigoDeBarras = "00198287480000000000000020571212219410000016500"
    
    If Len(Trim(sCodigoDeBarras)) <> 44 And Len(Trim(sCodigoDeBarras)) <> 48 Then gError 193339

    If gobjGeracaoArqCNABPag.objCCI.iCodBanco = 341 Then
        If Len(Trim(sCodigoDeBarras)) = 44 Then
            sCodigoDeBarrasAux = sCodigoDeBarras
            sCodigoDeBarras = ""
            For i = 1 To 4
                iDV = 0
                sTrecho = Mid(sCodigoDeBarrasAux, (i - 1) * 11 + 1, 11)
                Call Calcula_DV10(sTrecho, iDV)
                sCodigoDeBarras = sCodigoDeBarras & sTrecho & CStr(iDV)
            Next
        End If
    Else
        If Len(Trim(sCodigoDeBarras)) = 48 Then
            sCodigoDeBarras = Mid(objDet.sCodigoDeBarras, 1, 11) & Mid(objDet.sCodigoDeBarras, 13, 11) & _
                              Mid(objDet.sCodigoDeBarras, 25, 11) & Mid(objDet.sCodigoDeBarras, 37, 11)
        End If
    End If
    
    'Recolhe os dados do Segmento O do detalhe do lote
    With tRegSegmentoO
        .sNumeroBanco = FormataCpoNum(gobjGeracaoArqCNABPag.objCCI.iCodBanco, 3)
        .sCodigoLote = FormataCpoNum(objDet.objLote.lSequencial, 4)
        .sTipoRegistro = "3"
        .sNumeroRegistro = FormataCpoNum(objDet.lSequencial, 5)
        .sSegmento = "O"
        .sTipoMovimento = "000"
        .sCodigoBarras = sCodigoDeBarras
        .sNomeConcessionaria = FormataCpoTexto(objDet.sNomeFavorecido, 30)
        .sDataVencto = Format(objDet.dtDataVenctoParcela, "DDMMYYYY")
        .sDataPagamento = Format(gobjGeracaoArqCNABPag.objBorderoPagto.dtDataVencimento, "DDMMYYYY")
        .sValorPagamento = FormataCpoValor(objDet.dValorPagto, 15)
        .sNumDoctoEmpresa = FormataCpoTexto(CStr(objDet.lNumTitulo) & " " & CStr(objDet.iNumParcela) & " " & CStr(objDet.lFornecedor) & " " & CStr(objDet.iSeqBaixaParcPag), 20)
        .sNumDoctoBanco = FormataCpoTexto("", 20)
        .sCNAB = FormataCpoTexto("", 68)
        .sOcorrenciasRetorno = FormataCpoTexto("", 10)
                
        If gobjGeracaoArqCNABPag.objCCI.iCodBanco = 341 Then
            'Monta o registro 0
            sRegistro = .sNumeroBanco & .sCodigoLote & .sTipoRegistro & .sNumeroRegistro & .sSegmento & .sTipoMovimento
            sRegistro = sRegistro & .sCodigoBarras & .sNomeConcessionaria & .sDataVencto & "REA" & FormataCpoValor(0, 15)
            sRegistro = sRegistro & .sValorPagamento & .sDataPagamento & FormataCpoValor(0, 15) & FormataCpoTexto("", 15) & .sNumDoctoEmpresa & FormataCpoTexto("", 36) & .sOcorrenciasRetorno
        Else
            'Monta o registro 0
            sRegistro = .sNumeroBanco & .sCodigoLote & .sTipoRegistro & .sNumeroRegistro & .sSegmento & .sTipoMovimento
            sRegistro = sRegistro & .sCodigoInstrucaoMovto & .sCodigoBarras & .sNomeConcessionaria & .sDataVencto & .sDataPagamento
            sRegistro = sRegistro & .sValorPagamento & .sNumDoctoEmpresa & .sNumDoctoBanco & .sCNAB & .sOcorrenciasRetorno
        End If
    End With
    
    'Inclui no arquivo
    If Len(sRegistro) <> 240 Then gError 206438
    Print #1, DesacentuaTexto(sRegistro)
       
    Montar_Registro_SegmentoO = SUCESSO
    
    Exit Function
    
Erro_Montar_Registro_SegmentoO:

    Montar_Registro_SegmentoO = gErr
    
    Select Case gErr
    
        Case 193339
            Call Rotina_Erro(vbOKOnly, "ERRO_TITPAG_SEM_CODBARRAS", gErr, objDet.lNumTitulo, objDet.iNumParcela, objDet.lFornecedor)
        
        Case 206438
            Call Rotina_Erro(vbOKOnly, "ERRO_REG_DET_J_TAM_ERRADO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193340)
            
    End Select
    
    Exit Function
    
End Function

Function Calcula_DV10(ByVal sSequencia As String, iDigito As Integer) As Long
'Calcula o Digito Verificador no Módulo 10 para Linha Digitável de um Boleto Bancário
'Alterado por Jorge Specian - 09/03/2005

Dim lErro As Long
Dim iContador As Integer
Dim iNumero As Integer
Dim iTotalNumero As Integer
Dim iMultiplicador As Integer
Dim DezenaSuperior As Integer

On Error GoTo Erro_Calcula_DV10

    'Se nao for um valor numerico -> erro
    If Not IsNumeric(sSequencia) Then gError 134236
        
    'Inicia o multiplicador
    iMultiplicador = 2
    
    'Pega cada caracter do numero a partir da direita
    For iContador = Len(sSequencia) To 1 Step -1
        
        'Extrai o caracter e multiplica pelo multiplicador
        iNumero = Val(Mid(sSequencia, iContador, 1)) * iMultiplicador
        
        'Se o resultado for maior que nove soma os algarismos do resultado
        If iNumero > 9 Then
            
            iNumero = Val(left(iNumero, 1)) + Val(right(iNumero, 1))
        
        End If
        
        'Soma o resultado para totalização
        iTotalNumero = iTotalNumero + iNumero
        
        'Se o multiplicador for igual a 2 atribuir valor 1 se for 1 atribui 2
        iMultiplicador = IIf(iMultiplicador = 2, 1, 2)
        
    Next

    If iTotalNumero < 10 Then
        DezenaSuperior = 10
    Else
        DezenaSuperior = 10 * (Val(left(CStr(iTotalNumero), 1)) + 1)
    End If
    
    iDigito = DezenaSuperior - iTotalNumero

    'verifica as exceções ( 10 -> DV=0 )
    If iDigito = 10 Then iDigito = 0
    
    Calcula_DV10 = SUCESSO

    Exit Function

Erro_Calcula_DV10:

    Calcula_DV10 = gErr
    
    Select Case gErr
        
        Case 134236
            Call Rotina_Erro(vbOKOnly, "ERRO_SEQUENCIA_NAO_NUMERICO", gErr, sSequencia)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 148437)
    
    End Select
    
    Exit Function

End Function

Private Sub Zera_Var()
    
    glNumArquivoRemessa = 0
    giTotCarteiraSimples = 0
    giTotCarteiraVinculada = 0
    giTotCarteiraCaucionada = 0
    giTotCarteiraDescontada = 0
    gdValorTotCarteiraSimples = 0
    gdValorTotCarteiraVinculada = 0
    gdValorTotCarteiraCaucionada = 0
    gdValorTotCarteiraDescontada = 0
    giNumLotes = 0
    glNumRegistros = 0

    Set gColFiliaisEmpresa = Nothing
    Set gobjCobrancaEletronica = Nothing
    Set gobjGeracaoArqCNABPag = Nothing
    Set gobjFilialEmpresa = Nothing
    
End Sub

Private Function Trata_Texto(ByVal sTexto As String, Optional ByVal lTamMax As Long = 0, Optional bFazTrim As Boolean = True, Optional bRemoveEspacoDuplo As Boolean = True, Optional bRemoveAcentos As Boolean = True, Optional bUpperCase As Boolean = True) As String
Dim sTextoAux As String
Dim iIndice As Integer
Dim iPOS As Integer

    sTextoAux = sTexto
    
    sTextoAux = Replace(sTextoAux, "|", " ")
    sTextoAux = Replace(sTextoAux, "'", " ")
    sTextoAux = Replace(sTextoAux, "#", " ")
    sTextoAux = Replace(sTextoAux, "@", " ")
    sTextoAux = Replace(sTextoAux, "$", " ")
    sTextoAux = Replace(sTextoAux, "%", " ")
    sTextoAux = Replace(sTextoAux, "&", " ")
    sTextoAux = Replace(sTextoAux, "*", " ")
    sTextoAux = Replace(sTextoAux, "(", " ")
    sTextoAux = Replace(sTextoAux, ")", " ")
    sTextoAux = Replace(sTextoAux, "-", " ")
    sTextoAux = Replace(sTextoAux, "_", " ")
    sTextoAux = Replace(sTextoAux, "+", " ")
    sTextoAux = Replace(sTextoAux, "*", " ")
    sTextoAux = Replace(sTextoAux, "=", " ")
    sTextoAux = Replace(sTextoAux, "\", " ")
    sTextoAux = Replace(sTextoAux, "/", " ")
    sTextoAux = Replace(sTextoAux, ">", " ")
    sTextoAux = Replace(sTextoAux, "<", " ")
    sTextoAux = Replace(sTextoAux, "?", " ")
    sTextoAux = Replace(sTextoAux, "!", " ")
    sTextoAux = Replace(sTextoAux, "{", " ")
    sTextoAux = Replace(sTextoAux, "}", " ")
    sTextoAux = Replace(sTextoAux, "[", " ")
    sTextoAux = Replace(sTextoAux, "]", " ")
    
    sTextoAux = Replace(sTextoAux, ",", " ")
    sTextoAux = Replace(sTextoAux, ".", " ")
    sTextoAux = Replace(sTextoAux, ":", " ")
    sTextoAux = Replace(sTextoAux, ";", " ")
    
    sTextoAux = Replace(sTextoAux, "ª", "a")
    sTextoAux = Replace(sTextoAux, "º", "o")
    
    For iIndice = 0 To 31 'Remove caracteres não-imprimíveis
        sTextoAux = Replace(sTextoAux, Chr(iIndice), "")
    Next
    
    If bFazTrim Then 'Remove espaços no início e no final (OPCIONAL)
        sTextoAux = Trim(sTextoAux)
    End If
    
    If bRemoveEspacoDuplo Then 'Remove espaços duplos "  " (OPCIONAL)
        iPOS = InStr(1, sTextoAux, "  ")
        Do While iPOS <> 0
            sTextoAux = Replace(sTextoAux, "  ", " ")
            iPOS = InStr(1, sTextoAux, "  ")
        Loop
    End If
    
    If bRemoveAcentos Then 'Tira acentuação (OPCIONAL)
        sTextoAux = DesacentuaTexto(sTextoAux)
    End If
    
    If bUpperCase Then 'Muda para maiúsculo (OPCIONAL)
        sTextoAux = UCase(sTextoAux)
    End If
    
    If lTamMax > 0 Then 'Reduz o tamanho do texto (OPCIONAL)
        sTextoAux = left(sTextoAux, lTamMax)
    End If
    
    Trata_Texto = sTextoAux

End Function

Function Montar_Registro_SegmentoJ52(ByVal objDet As ClassDetArqCNABPag) As Long
'Recolhe os dados do Segmento J52 do detalhe do lote

Dim lErro As Long, sRegistro As String, sCodigoDeBarras As String
Dim tRegSegmentoJ52 As typePagto240SegmentoJ52, bDigitacao As Boolean, iDVTeste As Integer
Dim objFilialEmpresa As New AdmFiliais, sNomeEmpresa As String

On Error GoTo Erro_Montar_Registro_SegmentoJ52

    Call objDet.objLote.ObterProxReg
    
    'Busca os dados da filialempresa
    lErro = Obter_FilialEmpresa(giFilialEmpresa, objFilialEmpresa)
    If lErro <> SUCESSO Then gError 51737
    
    'Recolhe os dados do Segmento J52 do detalhe do lote
    With tRegSegmentoJ52
    
        .sNumeroBanco = FormataCpoNum(gobjGeracaoArqCNABPag.objCCI.iCodBanco, 3)
        .sCodigoLote = FormataCpoNum(objDet.objLote.lSequencial, 4)
        .sTipoRegistro = "3"
        .sNumeroRegistro = FormataCpoNum(objDet.lSequencial, 5)
        .sSegmento = "J"
        .sFiller = " "
        .sTipoMovimento = "00"
        .sIdentificacaoRegistroOpcional = "52"
        .sSacadoTipoInscricao = "2"
        .sSacadoCNPJ = FormataCpoNum(objFilialEmpresa.sCgc, 15)
        
        sNomeEmpresa = FormataCpoTexto(gsNomeEmpresa, 40)
        Call CF("Empresa_Le_Nome_Customizado", sNomeEmpresa)
        .sSacadoNome = sNomeEmpresa
        
        If Len(objDet.sCGCFavorecido) = 14 Then
            .sCedenteTipoInscricao = "2"
        Else
            .sCedenteTipoInscricao = "1"
        End If
        .sCedenteCNPJ = FormataCpoNum(objDet.sCGCFavorecido, 15)
        .sCedenteNome = FormataCpoTexto(objDet.sNomeFavorecido, 40)
        .sSacadorTipoInscricao = .sSacadoTipoInscricao
        .sSacadorCNPJ = .sSacadoCNPJ
        .sSacadorNome = .sSacadoNome
        .sFiller2 = String$(53, " ")

        'Monta o registro J52
        sRegistro = .sNumeroBanco & .sCodigoLote & .sTipoRegistro & .sNumeroRegistro & .sSegmento & .sFiller & .sTipoMovimento
        sRegistro = sRegistro & .sIdentificacaoRegistroOpcional & .sSacadoTipoInscricao & .sSacadoCNPJ & .sSacadoNome
        sRegistro = sRegistro & .sCedenteTipoInscricao & .sCedenteCNPJ & .sCedenteNome & .sSacadorTipoInscricao & .sSacadorCNPJ
        sRegistro = sRegistro & .sSacadorNome & .sFiller2
    End With
    
    'Inclui no arquivo
    If Len(sRegistro) <> 240 Then gError 206437
    Print #1, DesacentuaTexto(sRegistro)
       
    Montar_Registro_SegmentoJ52 = SUCESSO
    
    Exit Function
    
Erro_Montar_Registro_SegmentoJ52:

    Montar_Registro_SegmentoJ52 = gErr
    
    Select Case gErr
    
        Case 206437
            Call Rotina_Erro(vbOKOnly, "ERRO_REG_DET_J_TAM_ERRADO", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_TITPAG_SEM_CODBARRAS", gErr, objDet.lNumTitulo, objDet.iNumParcela, objDet.lFornecedor)
        
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147107)
            
    End Select
    
    Exit Function
    
End Function


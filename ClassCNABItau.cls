VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCNABItau"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Algumas Pendências:
    '- Criação de um campo em cobradores que faça com que o sequencial do arquivo seja reiniciado.
    '- Criação de um campo em Bordero Cobrança que indique que aquele bordero já gerou arquivo.
    '- Criação de um campo em carteira cobrador para indicar se vai ser o banco ou a empresa quem vai imprimir a boleta. (Motivo: Critica de endereço)
    
'Observacoes:
    '- a ocorrencia 18 deve ser tratada: nao deve
    '- Em que campo vem as custas: o mesmo das tarifas
    
'Dim gColFiliaisEmpresa As Collection
Dim gobjCobrancaEletronica As ClassCobrancaEletronica
Public gobjGeracaoArqCNABPag As ClassGeracaoArqCNABPag

Function Criar_Arquivo_Remessa_Itau(objCobrancaEletronica As ClassCobrancaEletronica) As Long
'Cria uma arquivo de remessa de cobranca com base nos dados do BD

Dim lErro As Long
Dim lTransacao As Long
Dim alComando(0 To 7) As Long
Dim iIndice As Integer
Dim colBorderosImpr As New Collection
Dim colBorderosSemImpr As New Collection

On Error GoTo Erro_Criar_Arquivo_Remessa_Itau
    
    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 51636
    Next
        
    'ABre a transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 51638
    
    Set gobjCobrancaEletronica = objCobrancaEletronica
    Set gColFiliaisEmpresa = New Collection
    
    lErro = Separa_Borderos_ComSem_Emissao(gobjCobrancaEletronica.colBorderos, colBorderosImpr, colBorderosSemImpr)
    If lErro <> SUCESSO Then Error 51887
    
    If colBorderosSemImpr.Count > 0 Then
    
        Set gobjCobrancaEletronica.colBorderos = colBorderosSemImpr
    
        'gera arquivo com layout do anexo A
        lErro = Gera_Arquivo_CobrancaSemEmiss(alComando(0), alComando(1), alComando(2), alComando(3))
        If lErro <> SUCESSO Then Error 62018
        
    End If
    
    
    If colBorderosImpr.Count > 0 Then
        
        Set gobjCobrancaEletronica.colBorderos = colBorderosImpr
        
        'gera arquivo para layout normal (nao o do anexo A)
        lErro = Gera_Arquivo_CobrancaComEmiss(alComando(0), alComando(1), alComando(2), alComando(3))
        If lErro <> SUCESSO Then Error 62019
        
    End If
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    lErro = Transacao_Commit()
    If lErro <> SUCESSO Then Error 62009
    
    Criar_Arquivo_Remessa_Itau = SUCESSO
    
    Exit Function

Erro_Criar_Arquivo_Remessa_Itau:

    Criar_Arquivo_Remessa_Itau = Err
    
    Select Case Err
    
        Case 51636
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 51638
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
            
        Case 62009
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
        
        Case 62018, 62019
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147258)
    
    End Select
    
    Close #1

    Call Transacao_Rollback

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Private Function Montar_RegistroTrailer_Itau(sRegTrailer As String, objCobrador As ClassCobrador) As Long
'Monta o registro

Dim tRegTrailerItau As typeRemTrailerItau
Dim lErro As Long

On Error GoTo Erro_Montar_RegistroTrailer_Itau

    'Preenche os dados do regsitro de trailer
    tRegTrailerItau.sIdentificacaoReg = "9"
    tRegTrailerItau.sComplemento1 = FormataCpoTexto("", 393)
    tRegTrailerItau.sSequencialRegistro = FormataCpoNum(gobjCobrancaEletronica.lSequencialRegistro, 6)
    
    'Monta o Registro de trailer
    sRegTrailer = tRegTrailerItau.sIdentificacaoReg & tRegTrailerItau.sComplemento1 & tRegTrailerItau.sSequencialRegistro

    Montar_RegistroTrailer_Itau = SUCESSO

    Exit Function

Erro_Montar_RegistroTrailer_Itau:

    Montar_RegistroTrailer_Itau = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147259)

    End Select

    Exit Function

End Function

Private Function Montar_RegistroHeader_Itau(sRegHeader As String) As Long
'Monta o registro

Dim tRegHeaderItau As typeRemHeaderItau
Dim lErro As Long

On Error GoTo Erro_Montar_RegistroHeader_Itau

    gobjCobrancaEletronica.lSequencialRegistro = 1

    'Carrega tRegHeaderItau com os dados do Header do aquivo
    lErro = Obter_Dados_Header(tRegHeaderItau)
    If lErro <> SUCESSO Then Error 51656
    
    'Monta o registro de Header do Itau
    sRegHeader = tRegHeaderItau.sIdentificacaoReg & tRegHeaderItau.sIdentificacaoArq & tRegHeaderItau.sIdentificacaoMov & tRegHeaderItau.sCodTipoServico & tRegHeaderItau.sIdentificacaoTipoServ
    sRegHeader = sRegHeader & tRegHeaderItau.sCodEmpresa & tRegHeaderItau.sComplemento1 & tRegHeaderItau.sNomeEmpresa
    sRegHeader = sRegHeader & tRegHeaderItau.sNumeroBanco & tRegHeaderItau.sNomeBanco & tRegHeaderItau.sDataEmissaoArq
    sRegHeader = sRegHeader & tRegHeaderItau.sComplemento2 & tRegHeaderItau.sSequencialRegistro
    
    Montar_RegistroHeader_Itau = SUCESSO

    Exit Function

Erro_Montar_RegistroHeader_Itau:

    Montar_RegistroHeader_Itau = Err

    Select Case Err

        Case 51656

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147260)

    End Select

    Exit Function

End Function

Private Function Montar_RegistroDetalhe_Itau(sRegDetalhe As String, objOcorrRemParcRec As ClassOcorrRemParcRec, objCarteiraCobrador As ClassCarteiraCobrador) As Long

'Monta o registro

Dim tRegDetalhe As typeRemDetalheItau
Dim lErro As Long

On Error GoTo Erro_Montar_RegistroDetalhe_Itau
    
    'Obtem os dados do registro de detalhe
    lErro = Obter_Dados_Detalhe(tRegDetalhe, objOcorrRemParcRec, objCarteiraCobrador)
    If lErro <> SUCESSO And lErro <> 32260 Then Error 51682
    If lErro <> SUCESSO Then Error 32261
    
    'Forma o registro de detalhe
    sRegDetalhe = tRegDetalhe.sIdentificacaoReg & tRegDetalhe.sTipoInscricaoEmpSac & tRegDetalhe.sNumInscricaoEmpSac & tRegDetalhe.sIdentificacaoEmp & tRegDetalhe.sComplemento1
    sRegDetalhe = sRegDetalhe & tRegDetalhe.sUsoDaEmpresa & tRegDetalhe.sNossoNumero & tRegDetalhe.sQuantMoeda
    sRegDetalhe = sRegDetalhe & tRegDetalhe.sNumCarteiraBanco & tRegDetalhe.sUsoDoBanco & tRegDetalhe.sCodCarteira & tRegDetalhe.sCodOcorrencia
    sRegDetalhe = sRegDetalhe & tRegDetalhe.sSeuNumero & tRegDetalhe.sVencimento & tRegDetalhe.sValorTitulo & tRegDetalhe.sNumeroBanco & tRegDetalhe.sAgenciaCobradora
    sRegDetalhe = sRegDetalhe & tRegDetalhe.sEspecie & tRegDetalhe.sAceite & tRegDetalhe.sDataEmissao & tRegDetalhe.sInstrucao1 & tRegDetalhe.sInstrucao2
    sRegDetalhe = sRegDetalhe & tRegDetalhe.sJurosDiarios & tRegDetalhe.sDescontoAte & tRegDetalhe.sValorDesconto & tRegDetalhe.sValorIOF & tRegDetalhe.sAbatimento
    sRegDetalhe = sRegDetalhe & tRegDetalhe.sCodigoInscSacado & tRegDetalhe.sNumInscricaoSacado & tRegDetalhe.sNome & tRegDetalhe.sComplemento2 & tRegDetalhe.sLogradouro
    sRegDetalhe = sRegDetalhe & tRegDetalhe.sBairro & tRegDetalhe.sCEP & tRegDetalhe.sCidade & tRegDetalhe.sEstado & tRegDetalhe.sSacadorAvalista & tRegDetalhe.sComplemento3
    sRegDetalhe = sRegDetalhe & tRegDetalhe.sDataDeMora & tRegDetalhe.sPrazo & tRegDetalhe.sComplemento4 & tRegDetalhe.sNumSequencial
    
    Montar_RegistroDetalhe_Itau = SUCESSO

    Exit Function

Erro_Montar_RegistroDetalhe_Itau:

    Montar_RegistroDetalhe_Itau = Err

    Select Case Err
    
        Case 51682, 32261

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147261)

    End Select

    Exit Function

End Function


Private Function Obter_Dados_Header(tRegHeaderItau As typeRemHeaderItau) As Long

Dim lErro As Long
Dim objCCI As New ClassContasCorrentesInternas
Dim sCodEmpresaBanco As String

On Error GoTo Erro_Obter_Dados_Header
        
    'Lê a conta corrente associada ao cobrador
    lErro = CF("ContaCorrenteInt_Le", gobjCobrancaEletronica.objCobrador.iCodCCI, objCCI)
    If lErro <> SUCESSO And lErro <> 11807 Then Error 51654
    If lErro <> SUCESSO Then Error 51655
    
    If objCCI.iCodBanco <> gobjCobrancaEletronica.objCobrador.iCodBanco Then Error 62054
    If Len(Trim(objCCI.sAgencia)) = 0 Or Len(Trim(objCCI.sNumConta)) = 0 Then Error 62055
    
    Set gobjCobrancaEletronica.objCCI = objCCI
    
    'gera o código que identifica a empresa no banco Itau
    'sCodEmpresaBanco = FormataCpoNum(Trim(objCCI.sAgencia), 4) & FormataCpoNum("", 2) & FormataCpoNum(Trim(objCCI.sNumConta), 5) & FormataCpoNum(Trim(objCCI.sDVAgConta), 1)
    'gobjCobrancaEletronica.sCodEmpresaBanco = gobjCobrancaEletronica.sCodEmpresaBanco
    gobjCobrancaEletronica.sCodEmpresaBanco = Obter_Info_Cobrador(gobjCobrancaEletronica.objCobradorCNABInfo.colInformacoes, 1)
    
    'Carrega os campos do cabeçalho
    tRegHeaderItau.sIdentificacaoReg = "0"
    tRegHeaderItau.sIdentificacaoArq = "1"
    tRegHeaderItau.sIdentificacaoMov = FormataCpoTexto("REMESSA", 7)
    tRegHeaderItau.sCodTipoServico = "01"
    tRegHeaderItau.sIdentificacaoTipoServ = FormataCpoTexto("COBRANCA", 15)
    tRegHeaderItau.sCodEmpresa = FormataCpoNum(gobjCobrancaEletronica.sCodEmpresaBanco, 12)
    tRegHeaderItau.sComplemento1 = FormataCpoTexto("", 8)
    tRegHeaderItau.sNomeEmpresa = FormataCpoTexto(gsNomeEmpresa, 30)
    
    Call CF("Empresa_Le_Nome_Customizado", tRegHeaderItau.sNomeEmpresa)
    
    tRegHeaderItau.sNumeroBanco = "341"
    tRegHeaderItau.sNomeBanco = FormataCpoTexto("ITAU", 15)
    tRegHeaderItau.sDataEmissaoArq = Format(gdtDataHoje, "DDMMYY")
    tRegHeaderItau.sComplemento2 = FormataCpoTexto("", 294)
    tRegHeaderItau.sSequencialRegistro = FormataCpoNum(gobjCobrancaEletronica.lSequencialRegistro, 6)
    
    'Incrementa o sequencial do arquivo
    gobjCobrancaEletronica.lSequencialRegistro = gobjCobrancaEletronica.lSequencialRegistro + 1

    Obter_Dados_Header = SUCESSO

    Exit Function
    
Erro_Obter_Dados_Header:

    Obter_Dados_Header = Err
    
    Select Case Err
    
        Case 62054
            lErro = Rotina_Erro(vbOKOnly, "ERRO_BANCO_CCI_DIFERENTE_COBRADOR", Err, objCCI.iCodBanco, gobjCobrancaEletronica.objCobrador.iCodBanco)
    
        Case 62055
            lErro = Rotina_Erro(vbOKOnly, "ERRO_AGENCIA_CONTA_COBRADOR_NAO_PREENCHIDAS", Err)
    
        Case 51654
        
        Case 51655
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_COBRADOR_NAO_ENCONTRADA", Err, gobjCobrancaEletronica.iCobrador)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147262)
            
    End Select
    
    Exit Function

End Function

Private Function Obter_Dados_Detalhe(tRegDetalhe As typeRemDetalheItau, objOcorrRemParcRem As ClassOcorrRemParcRec, objCarteiraCobrador As ClassCarteiraCobrador) As Long
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
    If lErro <> SUCESSO And lErro <> 51674 Then Error 51670
    If lErro <> SUCESSO Then Error 32260
    
    lErro = CF("CobrancaEletronica_Obter_Endereco", gobjCobrancaEletronica, objEndereco)
    If lErro <> SUCESSO Then Error 51671
    
    If objOcorrRemParcRem.iCodOcorrencia = COBRANCA_OCORR_INC_TITULO Then
    
        lErro = Calcula_NossoNumero_Itau(objCarteiraCobrador, sNossoNumero)
        If lErro <> SUCESSO Then Error 51709
        
    Else
    
        sNossoNumero = objParcelaRec.sNumTitCobrador
    
    End If
    
    tRegDetalhe.sIdentificacaoReg = "1"
    tRegDetalhe.sTipoInscricaoEmpSac = "02"
    tRegDetalhe.sNumInscricaoEmpSac = FormataCpoNum(Trim(gobjCobrancaEletronica.sCGCCedente), 14)
    tRegDetalhe.sIdentificacaoEmp = FormataCpoNum(gobjCobrancaEletronica.sCodEmpresaBanco, 12)
    tRegDetalhe.sComplemento1 = FormataCpoTexto("", 8)
    tRegDetalhe.sUsoDaEmpresa = FormataCpoTexto(objOcorrRemParcRem.lNumIntDoc, 25)
    tRegDetalhe.sNossoNumero = FormataCpoNum(sNossoNumero, 8)
    tRegDetalhe.sQuantMoeda = FormataCpoNum("", 13)
    tRegDetalhe.sNumCarteiraBanco = FormataCpoNum(objCarteiraCobrador.iNumCarteiraNoBanco, 3)
    tRegDetalhe.sUsoDoBanco = FormataCpoTexto("", 21)
    tRegDetalhe.sCodCarteira = FormataCpoTexto(objCarteiraCobrador.sCodCarteiraNoBanco, 1)
    tRegDetalhe.sCodOcorrencia = FormataCpoNum(objOcorrRemParcRem.iCodOcorrencia, 2)
            
    '??? hicare
    'tRegDetalhe.sSeuNumero = FormataCpoNum(gobjCobrancaEletronica.iFilialEmpTitulo, 2) & FormataCpoNum(objParcelaRec.iNumParcela, 2) & FormataCpoNum(gobjCobrancaEletronica.lNumTitulo, 6)
    tRegDetalhe.sSeuNumero = FormataCpoTexto(CStr(gobjCobrancaEletronica.lNumTitulo) & SEPARADOR & CStr(objParcelaRec.iNumParcela), 10)

    tRegDetalhe.sVencimento = Format(objParcelaRec.dtDataVencimento, "DDMMYY")
    tRegDetalhe.sValorTitulo = FormataCpoValor(objOcorrRemParcRem.dValorCobrado, 13)
    tRegDetalhe.sNumeroBanco = "341"
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
    
    Call CF("Itau_Obter_Dados_DetalheSegP_Customizado", gobjCobrancaEletronica, tRegDetalhe.sSeuNumero, objParcelaRec, objOcorrRemParcRem, tRegDetalhe.sInstrucao1, tRegDetalhe.sInstrucao2)
    
    tRegDetalhe.sJurosDiarios = FormataCpoValor(Round(objOcorrRemParcRem.dValorCobrado * gobjCobrancaEletronica.dPercJurosDiario, 2), 13)
    
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
    tRegDetalhe.sNome = FormataCpoTexto(gobjCobrancaEletronica.sNomeSacado, 30)
    tRegDetalhe.sComplemento2 = FormataCpoTexto("", 10)
    tRegDetalhe.sLogradouro = FormataCpoTexto(Trim(objEndereco.sEndereco), 40)
    tRegDetalhe.sBairro = FormataCpoTexto(Trim(objEndereco.sBairro), 12)
    tRegDetalhe.sCEP = Formata_Cep(objEndereco.sCEP, 8)
    tRegDetalhe.sCidade = FormataCpoTexto(objEndereco.sCidade, 15)
    tRegDetalhe.sEstado = FormataCpoTexto(objEndereco.sSiglaEstado, 2)
    
    tRegDetalhe.sSacadorAvalista = FormataCpoTexto("", 2)
    If objParcelaRec.dtDesconto2Ate <> DATA_NULA Then
        tRegDetalhe.sSacadorAvalista = tRegDetalhe.sSacadorAvalista & Format(objParcelaRec.dtDesconto2Ate, "DDMMYY")
        'Calcula o valor do desconto quando for do tipo percentual
        dValorDesconto = 0
        If objParcelaRec.iDesconto2Codigo = PERC_ANT_DIA Or objParcelaRec.iDesconto2Codigo = PERC_ANT_DIA_UTIL Or objParcelaRec.iDesconto2Codigo = Percentual Then
            dValorDesconto = StrParaDbl(Format(objParcelaRec.dDesconto2Valor * objParcelaRec.dSaldo, "Fixed"))
        Else
            If objParcelaRec.iDesconto2Codigo = VALOR_FIXO Then
                dValorDesconto = objParcelaRec.dDesconto2Valor
            End If
        End If
        
        tRegDetalhe.sSacadorAvalista = tRegDetalhe.sSacadorAvalista & FormataCpoValor(dValorDesconto, 13)
        
        If objParcelaRec.dtDesconto3Ate <> DATA_NULA Then
            tRegDetalhe.sSacadorAvalista = tRegDetalhe.sSacadorAvalista & Format(objParcelaRec.dtDesconto3Ate, "DDMMYY")
            'Calcula o valor do desconto quando for do tipo percentual
            dValorDesconto = 0
            If objParcelaRec.iDesconto3Codigo = PERC_ANT_DIA Or objParcelaRec.iDesconto3Codigo = PERC_ANT_DIA_UTIL Or objParcelaRec.iDesconto3Codigo = Percentual Then
                dValorDesconto = StrParaDbl(Format(objParcelaRec.dDesconto3Valor * objParcelaRec.dSaldo, "Fixed"))
            Else
                If objParcelaRec.iDesconto3Codigo = VALOR_FIXO Then
                    dValorDesconto = objParcelaRec.dDesconto3Valor
                End If
            End If
            
            tRegDetalhe.sSacadorAvalista = tRegDetalhe.sSacadorAvalista & FormataCpoValor(dValorDesconto, 13)
        Else
            tRegDetalhe.sSacadorAvalista = tRegDetalhe.sSacadorAvalista & FormataCpoTexto("", 19)
        End If
    Else
        tRegDetalhe.sSacadorAvalista = tRegDetalhe.sSacadorAvalista & FormataCpoTexto("", 38)
    End If
    
    'Isso porque a gente trabalha c\ os três descontos na parcela
    tRegDetalhe.sComplemento3 = ""
    tRegDetalhe.sDataDeMora = ""
    
    tRegDetalhe.sPrazo = FormataCpoNum(0, 2)
    tRegDetalhe.sComplemento4 = FormataCpoTexto("", 1)
    tRegDetalhe.sNumSequencial = FormataCpoNum(gobjCobrancaEletronica.lSequencialRegistro, 6)
    
    'Incrementa o sequencial de registro
    gobjCobrancaEletronica.lSequencialRegistro = gobjCobrancaEletronica.lSequencialRegistro + 1
    
    Obter_Dados_Detalhe = SUCESSO

    Exit Function
    
Erro_Obter_Dados_Detalhe:

    Obter_Dados_Detalhe = Err
    
    Select Case Err

        Case 51670, 51671, 51709, 32260
       
        Case 51708
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ENDERECO_FILIALCLIENTE_NAO_INFORMADO", Err, gobjCobrancaEletronica.lCliente, gobjCobrancaEletronica.iFilialCli)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147263)

    End Select

    Exit Function

End Function

Function Obter_FilialEmpresa(iFilialEmpresa As Integer, objFilialEmpresa As AdmFiliais) As Long
'Retorna o objFilialEmpresa prenchido com os dados da filialempresa passada por parâmetro

Dim lErro As Long
Dim objFilialEmpresaAux As AdmFiliais

On Error GoTo Erro_Obter_FilialEmpresa

    'Carrega o objFilialEmpresa com a sua chave
    objFilialEmpresa.lCodEmpresa = glEmpresa
    objFilialEmpresa.iCodFilial = iFilialEmpresa

    'Lê a FilialEmpresa
    lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
    If lErro <> SUCESSO And lErro <> 27378 Then Error 51666
    If lErro <> SUCESSO Then Error 51667

    Obter_FilialEmpresa = SUCESSO

    Exit Function

Erro_Obter_FilialEmpresa:

    Obter_FilialEmpresa = Err

    Select Case Err

        Case 51666

        Case 51667
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", Err, objFilialEmpresa.iCodFilial)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147264)

    End Select

    Exit Function

End Function

Function Processar_ArquivoRetorno_Cobranca_Itau(objCobrancaEletronica As ClassCobrancaEletronica) As Long
'Faz o processamento do arquivo de retorno

Dim lTransacao As Long
Dim lErro As Long
Dim iIndice As Integer
Dim sRegistro As String
Dim iCodRegistro As Integer
Dim objBorderoRetCobr As New ClassBorderoRetCobr
Dim sIdentifEmpresa As String
Dim objBorderoRetCobrCart As ClassBorderoRetCobrCart
Dim lComando As Long, dValorMov As Double, bArquivoAberto As Boolean
Dim lComando1 As Long, bDesistiu As Boolean, lComando2 As Long

On Error GoTo Erro_Processar_ArquivoRetorno_Cobranca_Itau

    bArquivoAberto = False
    
    objBorderoRetCobr.iBanco = 341
    
    'le as carteiras de cobranca do cobrador
    lErro = CF("CarteirasCobrador_Le", objCobrancaEletronica.objCobrador, objBorderoRetCobr.colCartCobr)
    If lErro <> SUCESSO Then Error 32265
    
    'Abre os comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 51824

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 51825

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 51825

    'Abre o arquivo de retorno
    Open objCobrancaEletronica.sNomeArquivoRetorno For Input As #2
    
    bArquivoAberto = True
    
'    objBorderoRetCobr.objMovCCI.dValor = 0
'
    'Abre a transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 51722
    
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
            If lErro <> SUCESSO Then Error 51686

            'obtem codigo da empresa
            sIdentifEmpresa = objCobrancaEletronica.objCobradorCNABInfo.colInformacoes(1).sNome

        'Se o registro for de Detalhe ("Transação")
        ElseIf iCodRegistro = 1 Then
        
            'Processa os dados contidos no registro de detalhe
            lErro = Processa_Registro_Detalhe(sRegistro, objBorderoRetCobr, objCobrancaEletronica, sIdentifEmpresa, lComando, lComando1, lComando2)
            If lErro <> SUCESSO Then Error 51687
        
        'Se o registro for de Trailer
        ElseIf iCodRegistro = 9 Then
        
            'Processa os dados do trailer do arquivo
            lErro = Processa_Registro_Trailer(sRegistro)
            If lErro <> SUCESSO Then Error 51688
        
        Else
            Error 51710
        End If
        
    Loop
        
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    
    'Faz o commit na transacao
    lErro = Transacao_Commit()
    If lErro <> SUCESSO Then Error 51734
        
    Close #2

    bArquivoAberto = False
    
    'Grava os dados recolhidos durante o processamento do arquivo de retorno
    lErro = CF("BorderoRetCobr_ConfirmaGrava", objBorderoRetCobr)
    If lErro <> SUCESSO Then Error 51689
    
    Processar_ArquivoRetorno_Cobranca_Itau = SUCESSO
    
    Exit Function
    
Erro_Processar_ArquivoRetorno_Cobranca_Itau:

    Processar_ArquivoRetorno_Cobranca_Itau = Err
    
    Select Case Err
    
        Case 51722
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
        
        Case 51734
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
        
        Case 51686, 51687, 51688, 51689, 32260, 32265
        
        Case 51824, 51825
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 51710
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORMATO_ARQUIVO_INCORRETO", Err, objCobrancaEletronica.sNomeArquivoRetorno)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147265)
    
    End Select

    'Faz um RollBack na transacao
    Call Transacao_Rollback
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    
    If bArquivoAberto Then Close #2

    Exit Function

End Function

Private Function Processa_Registro_Header(sRegistro As String, objBorderoRetCobr As ClassBorderoRetCobr, objCobrador As ClassCobrador, sNomeArquivo As String) As Long

Dim lErro As Long
Dim tRegHeaderItau As typeRetHeaderItau
Dim objContaCorrente As New ClassContasCorrentesInternas

On Error GoTo Erro_Processa_Registro_Header

    'Extrai da variável lida no arquivo os dados do Header
    lErro = Extrair_Header_Retorno(tRegHeaderItau, sRegistro)
    If lErro <> SUCESSO Then Error 51699
    
    'Preenche o objBorderoRetCobr
    objBorderoRetCobr.dtDataRecepcao = gdtDataHoje
    objBorderoRetCobr.iCobrador = objCobrador.iCodigo
    objBorderoRetCobr.sNomeArq = sNomeArquivo
    objBorderoRetCobr.dtDataCredito = Extrair_Data(tRegHeaderItau.sDataCredito)
    
    Processa_Registro_Header = SUCESSO
    
    Exit Function

Erro_Processa_Registro_Header:

    Processa_Registro_Header = Err

    Select Case Err
    
        Case 51699, 51711
    
        Case 51712
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_CORRENTE_NAO_ENCONTRADA", Err, objContaCorrente.iCodigo)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147266)
            
    End Select
    
    Exit Function

End Function

Private Function Extrair_Header_Retorno(tRegHeaderItau As typeRetHeaderItau, sRegistro As String) As Long
'Extrai do sRegistro os dados do header do arquido de retorno

Dim lErro As Long

On Error GoTo Erro_Extrair_Header_Retorno

    tRegHeaderItau.sIdentificacaoReg = Mid(sRegistro, 1, 1)
    tRegHeaderItau.sIdentificacaoArq = Mid(sRegistro, 2, 1)
    tRegHeaderItau.sIdentificacaoMov = Mid(sRegistro, 3, 7)
    tRegHeaderItau.sCodTipoServico = Mid(sRegistro, 10, 2)
    tRegHeaderItau.sIdentificacaoTipoServ = Mid(sRegistro, 12, 15)
    tRegHeaderItau.sCodEmpresa = Mid(sRegistro, 27, 12)
    tRegHeaderItau.sComplemento1 = Mid(sRegistro, 39, 8)
    tRegHeaderItau.sNomeEmpresa = Mid(sRegistro, 47, 30)
    tRegHeaderItau.sNumeroBanco = Mid(sRegistro, 77, 3)
    tRegHeaderItau.sNomeBanco = Mid(sRegistro, 80, 15)
    tRegHeaderItau.sDataEmissaoArq = Mid(sRegistro, 95, 6)
    tRegHeaderItau.sDensidade = Mid(sRegistro, 101, 5)
    tRegHeaderItau.sUnidadeDensid = Mid(sRegistro, 106, 3)
    tRegHeaderItau.sNumSequencialArqRet = Mid(sRegistro, 109, 5)
    tRegHeaderItau.sComplemento2 = Mid(sRegistro, 120, 275)
    tRegHeaderItau.sSequencialRegistro = Mid(sRegistro, 395, 6)
    tRegHeaderItau.sDataCredito = Mid(sRegistro, 114, 6)
    
    Extrair_Header_Retorno = SUCESSO

    Exit Function

Erro_Extrair_Header_Retorno:

    Extrair_Header_Retorno = Err
    
    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147267)
            
    End Select
    
    Exit Function

End Function

Private Function Processa_Registro_Detalhe(sRegistro As String, objBorderoRetCobr As ClassBorderoRetCobr, objCobrancaEletronica As ClassCobrancaEletronica, sIdentifEmpresa As String, lComando As Long, lComando1 As Long, lComando2 As Long) As Long

Dim tRegDetalhe As typeRetDetalheItau
Dim objOcorrRemParcRec As New ClassOcorrRemParcRec
Dim iCodOcorrencia As Integer, iTipoOcorrencia As Integer
Dim iCodCarteiraNoBanco As Integer
Dim lErro As Long, objDetRetCobr As New ClassDetRetCobr
Dim objParcelaRec As New ClassParcelaReceber
Dim objBorderoRetCobrCart As ClassBorderoRetCobrCart
Dim objInfoParcRec As ClassInfoParcRec

On Error GoTo Erro_Processa_Registro_Detalhe

    'Extrai da variável lida no arquivo os dados do registro de detalhe
    lErro = Extrair_Detalhe_Retorno(tRegDetalhe, sRegistro)
    If lErro <> SUCESSO Then Error 51693
    
    If (sIdentifEmpresa = tRegDetalhe.sCodEmpresa) Then
    
        'Recolhe o código da ocorrência
        iCodOcorrencia = StrParaInt(tRegDetalhe.sCodOcorrencia)
        
        Select Case iCodOcorrencia
            
            'Quando ocorre baixa
            Case 6, 7, 8
                iTipoOcorrencia = RETCOBR_DET_BAIXA
                lErro = Prepara_CartCobrRet1(objCobrancaEletronica, objBorderoRetCobr, objBorderoRetCobrCart, objDetRetCobr, tRegDetalhe)
                If lErro <> SUCESSO Then Error 32356

                lErro = RetCobr_Obtem_ParcelaItau(objBorderoRetCobr, objDetRetCobr, objParcelaRec, tRegDetalhe, lComando)
                If lErro <> SUCESSO Then Error 32264
                'se houve erro
                If objDetRetCobr.iTipoCritica <> RETCOBR_CRITICA_SEM_ERRO Then
            
                    'incluir ocorrencia de erro do retorno
                    Call objBorderoRetCobr.colRetCobrErros.Add(objDetRetCobr)
            
                Else
            
                    lErro = CF("RetCobr_Preenche_InfoParcRec", objParcelaRec, objInfoParcRec)
                    '##############################
                    'Alterado por Wagner
                    'If lErro <> SUCESSO Then Error 51695
                    
                    'Se deu erro ao ler a Parcela ou o Titulo não processa e inclui erro
                    If lErro <> SUCESSO Then
                    
                        objDetRetCobr.iTipoCritica = RETCOBR_CRITICA_SEM_PARC
                    
                        'incluir ocorrencia de erro do retorno
                        Call objBorderoRetCobr.colRetCobrErros.Add(objDetRetCobr)
                    
                    Else
                        'Processa a parcela a receber
                        lErro = CF("Processa_ParcelaRec", objParcelaRec, objBorderoRetCobr, objBorderoRetCobrCart, objDetRetCobr, objInfoParcRec)
                        If lErro <> SUCESSO Then Error 51695
                    End If
                    '##############################

                End If
                
            'cobranca de tarifas
            Case 12, 20, 23, 24, 26 To 30, 37 To 46, 51 To 55
                iTipoOcorrencia = RETCOBR_DET_TARIFAS
                                    
            'cobranca de custas
            Case 33 To 36, 48, 56
                iTipoOcorrencia = RETCOBR_DET_CUSTAS
                
            'entrada confirmada
            Case 2
                iTipoOcorrencia = RETCOBR_DET_CONFIRMADO
                Call RegDetCobr_ObterInfo(objDetRetCobr, tRegDetalhe)

                lErro = RetCobr_Obtem_ParcelaItau(objBorderoRetCobr, objDetRetCobr, objParcelaRec, tRegDetalhe, lComando)
                If lErro <> SUCESSO Then Error 32264
                'se houve erro
                If objDetRetCobr.iTipoCritica <> RETCOBR_CRITICA_SEM_ERRO Then
            
                    'incluir ocorrencia de erro do retorno
                    Call objBorderoRetCobr.colRetCobrErros.Add(objDetRetCobr)
            
                Else
            
                    'Processa a parcela a receber confirmada
                    lErro = Processa_ParcelaRecConfirmada(objParcelaRec, objBorderoRetCobr, objBorderoRetCobrCart, objDetRetCobr, tRegDetalhe, lComando1, lComando2)
                    If lErro <> SUCESSO Then Error 51695
                                  
                End If
                
            Case 3
                iTipoOcorrencia = RETCOBR_DET_REJEITADO
        
            Case Else
                iTipoOcorrencia = RETCOBR_DET_IGNORAR
        
        End Select
    
        If iTipoOcorrencia = RETCOBR_DET_REJEITADO Or iTipoOcorrencia = RETCOBR_DET_CUSTAS Or iTipoOcorrencia = RETCOBR_DET_TARIFAS Then
        
            lErro = Prepara_CartCobrRet1(objCobrancaEletronica, objBorderoRetCobr, objBorderoRetCobrCart, objDetRetCobr, tRegDetalhe)
            If lErro <> SUCESSO Then Error 32257
            
            lErro = CF("Processa_Registro_Detalhe_Geral", objBorderoRetCobr, objCobrancaEletronica, iTipoOcorrencia, RETCOBR_TIPO_SEU_NUMERO1, objDetRetCobr, objBorderoRetCobrCart, lComando)
            If lErro <> SUCESSO Then Error 32356
                                      
        End If
    
    End If
    
    Processa_Registro_Detalhe = SUCESSO
    
    Exit Function

Erro_Processa_Registro_Detalhe:

    Processa_Registro_Detalhe = Err

    Select Case Err
    
        Case 51683, 51693, 51694, 51695, 32356, 32257, 32258, 32264
    
        Case 51696
            lErro = Rotina_Erro(vbOKOnly, "ERRO_OCORRREMPARCREC_NAO_CADASTRADA", Err, objOcorrRemParcRec.lNumIntDoc)
            
        Case 51697
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PARCELAREC_NAO_CADASTRADA", Err, objOcorrRemParcRec.lNumIntParc)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147268)
            
    End Select
    
    Exit Function

End Function

Private Function Extrair_Detalhe_Retorno(tRegDetalhe As typeRetDetalheItau, sRegistro As String) As Long
'Extrai do sRegistro os dados do Detalhe do arquido de retorno

Dim lErro As Long

On Error GoTo Erro_Extrair_Detalhe_Retorno

    tRegDetalhe.sIdentificacaoReg = Mid(sRegistro, 1, 1)
    tRegDetalhe.sTipoInscricaoEmpresa = Mid(sRegistro, 2, 2)
    tRegDetalhe.sNumInscricaoEmpresa = Mid(sRegistro, 4, 14)
    tRegDetalhe.sCodEmpresa = Mid(sRegistro, 18, 12)
    tRegDetalhe.sComplemento1 = Mid(sRegistro, 30, 8)
    tRegDetalhe.sUsoDaEmpresa = Mid(sRegistro, 38, 25)
    tRegDetalhe.sNossoNumero1 = Mid(sRegistro, 63, 8)
    tRegDetalhe.sComplemento2 = Mid(sRegistro, 71, 12)
    tRegDetalhe.sNumCarteira = Mid(sRegistro, 83, 3)
    tRegDetalhe.sNossoNumero2 = Mid(sRegistro, 86, 8)
    tRegDetalhe.sDACNossoNumero2 = Mid(sRegistro, 94, 1)
    tRegDetalhe.sComplemento3 = Mid(sRegistro, 95, 13)
    tRegDetalhe.sCodCarteira = Mid(sRegistro, 108, 1)
    tRegDetalhe.sCodOcorrencia = Mid(sRegistro, 109, 2)
    tRegDetalhe.sDataOcorrencia = Mid(sRegistro, 111, 6)
    tRegDetalhe.sSeuNumero = Mid(sRegistro, 117, 10)
    tRegDetalhe.sNossoNumero3 = Mid(sRegistro, 127, 8)
    tRegDetalhe.sComplemento4 = Mid(sRegistro, 135, 12)
    tRegDetalhe.sVencimento = Mid(sRegistro, 147, 6)
    tRegDetalhe.sValorTitulo = Mid(sRegistro, 153, 13)
    tRegDetalhe.sNumeroBanco = Mid(sRegistro, 166, 3)
    tRegDetalhe.sAgenciaCobradora = Mid(sRegistro, 169, 4)
    tRegDetalhe.sDACAgCobradora = Mid(sRegistro, 173, 1)
    tRegDetalhe.sEspecie = Mid(sRegistro, 174, 2)
    tRegDetalhe.sTarifaCobranca = Mid(sRegistro, 176, 13)
    tRegDetalhe.sComplemento5 = Mid(sRegistro, 189, 26)
    tRegDetalhe.sValorIOF = Mid(sRegistro, 215, 13)
    tRegDetalhe.sAbatimento = Mid(sRegistro, 228, 13)
    tRegDetalhe.sDescontos = Mid(sRegistro, 241, 13)
    tRegDetalhe.sValorPrincipal = Mid(sRegistro, 254, 13)
    tRegDetalhe.sJuros = Mid(sRegistro, 267, 13)
    tRegDetalhe.sComplemento6 = Mid(sRegistro, 280, 16)
    tRegDetalhe.sDataCredito = Mid(sRegistro, 296, 6)
    tRegDetalhe.sComplemento7 = Mid(sRegistro, 302, 23)
    tRegDetalhe.sNomeSacado = Mid(sRegistro, 325, 30)
    tRegDetalhe.sComplemento8 = Mid(sRegistro, 355, 23)
    tRegDetalhe.sErros = Mid(sRegistro, 378, 8)
    tRegDetalhe.sComplemento9 = Mid(sRegistro, 86, 9)
    tRegDetalhe.sNumSequencial = Mid(sRegistro, 395, 6)
    
    Extrair_Detalhe_Retorno = SUCESSO

    Exit Function

Erro_Extrair_Detalhe_Retorno:

    Extrair_Detalhe_Retorno = Err
    
    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147269)
            
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

Function Processa_Registro_Trailer(sRegistroTrailer As String) As Long

Dim lErro As Long
Dim tRegTrailerItau As typeRetTrailerItau

On Error GoTo Erro_Processa_Registro_Trailer

    'Extrai da variável lida no arquivo os dados do registro de detalhe
    lErro = Extrair_Trailer_Retorno(tRegTrailerItau, sRegistroTrailer)
    If lErro <> SUCESSO Then Error 51713
    
    Processa_Registro_Trailer = SUCESSO
    
    Exit Function
    
Erro_Processa_Registro_Trailer:

    Processa_Registro_Trailer = Err
    
    Select Case Err
    
        Case 51713
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147270)
            
    End Select
    
    Exit Function
    
End Function
Private Function Extrair_Trailer_Retorno(tRegTrailerItau As typeRetTrailerItau, sRegistro As String) As Long
'Extrai do sRegistro os dados do Trailer do arquido de retorno

Dim lErro As Long

On Error GoTo Erro_Extrair_Trailer_Retorno

    tRegTrailerItau.sIdentificacaoReg = Mid(sRegistro, 1, 1)
    tRegTrailerItau.sCodigoRetorno = Mid(sRegistro, 2, 1)
    tRegTrailerItau.sCodigoServico = Mid(sRegistro, 3, 2)
    tRegTrailerItau.sNumeroBanco = Mid(sRegistro, 5, 3)
    tRegTrailerItau.sComplemento1 = Mid(sRegistro, 8, 10)
    tRegTrailerItau.sQuantTitulosCobrSimp = Mid(sRegistro, 18, 8)
    tRegTrailerItau.sValorTotalSimp = Mid(sRegistro, 26, 14)
    tRegTrailerItau.sAvisoBancario = Mid(sRegistro, 40, 8)
    tRegTrailerItau.sComplemento2 = Mid(sRegistro, 48, 10)
    tRegTrailerItau.sQuantTitulosCobrVinc = Mid(sRegistro, 58, 8)
    tRegTrailerItau.sValorTotalVinc = Mid(sRegistro, 66, 14)
    tRegTrailerItau.sAvisoBancario2 = Mid(sRegistro, 80, 8)
    tRegTrailerItau.sComplemento3 = Mid(sRegistro, 88, 90)
    tRegTrailerItau.sQuantTitulosCobrEscr = Mid(sRegistro, 178, 8)
    tRegTrailerItau.sValorTotalEscr = Mid(sRegistro, 186, 13)
    tRegTrailerItau.sAvisoBancario3 = Mid(sRegistro, 200, 8)
    tRegTrailerItau.sControleArquivo = Mid(sRegistro, 208, 5)
    tRegTrailerItau.sQuantDetalhes = Mid(sRegistro, 213, 8)
    tRegTrailerItau.sValorTotalInformado = Mid(sRegistro, 221, 14)
    tRegTrailerItau.sComplemento4 = Mid(sRegistro, 235, 160)
    tRegTrailerItau.sSequencialRegistro = Mid(sRegistro, 395, 6)

    Extrair_Trailer_Retorno = SUCESSO

    Exit Function

Erro_Extrair_Trailer_Retorno:

    Extrair_Trailer_Retorno = Err
    
    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147271)
            
    End Select
    
    Exit Function

End Function

Function BorderoPagto_ProcessarDetPag(ByVal objDetArqCNABPag As ClassDetArqCNABPag) As Long

Dim lErro As Long

On Error GoTo Erro_BorderoPagto_ProcessarDetPag

    If objDetArqCNABPag.iTipoCobranca <> TIPO_COBRANCA_BANCARIA Then
    
        'Monta um registro de Segmento A
        lErro = Montar_Registro_SegmentoA(objDetArqCNABPag)
        If lErro <> SUCESSO Then gError 51729
            
    Else
        'Monta um registro de Segmento J
        lErro = Montar_Registro_SegmentoJ(objDetArqCNABPag)
        If lErro <> SUCESSO Then gError 51730
    
    End If
        
    BorderoPagto_ProcessarDetPag = SUCESSO
     
    Exit Function
    
Erro_BorderoPagto_ProcessarDetPag:

    BorderoPagto_ProcessarDetPag = gErr
     
    Select Case gErr
          
        Case 51729, 51730
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147272)
     
    End Select
     
    Exit Function

End Function

Function BorderoPagto_Criar_ArquivoCNAB_Itau(objGeracaoArqCNABPag As ClassGeracaoArqCNABPag) As Long
'Cria o arquivo de remessa de títulos a pagar nos padrões do ITAU

Dim lErro As Long
Dim sNomeArquivo As String
Dim lTransacao As Long
Dim iIndice As Integer
Dim objDetArqCNABPag As ClassDetArqCNABPag
Dim alComando(0 To 2) As Long
Dim objLote As ClassLoteArqCNABPag
Dim objArq As New ClassArqCNABPag

On Error GoTo Erro_BorderoPagto_Criar_ArquivoCNAB_Itau

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 51721
    Next

    'Abre a transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 51722
    
    Set gobjGeracaoArqCNABPag = objGeracaoArqCNABPag
    Set gobjGeracaoArqCNABPag.objArq = objArq
    
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
    lErro = Comando_ExecutarPos(alComando(1), "UPDATE BorderosPagto SET NomeArq = ? , NumArqRemessa = ?", alComando(0), sNomeArquivo, gobjGeracaoArqCNABPag.lSequencialArq + 1)
    If lErro <> AD_SQL_SUCESSO Then gError 51733
    
    'Faz o commit na transacao
    lErro = Transacao_Commit()
    If lErro <> SUCESSO Then gError 51734
        
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    BorderoPagto_Criar_ArquivoCNAB_Itau = SUCESSO
    
    Exit Function
    
Erro_BorderoPagto_Criar_ArquivoCNAB_Itau:

    BorderoPagto_Criar_ArquivoCNAB_Itau = gErr
    
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147273)
            
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
Dim tRegHeaderArq As typePagtoHeaderArqItau
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_Montar_RegistroHeaderCNABPag

    Call objArq.ObterProxReg

    'Busca os dados da nossa filial empresa
    lErro = Obter_FilialEmpresa(giFilialEmpresa, objFilialEmpresa)
    If lErro <> SUCESSO Then Error 51736
    
    'Recolhe os dados do cabeçalho
    tRegHeaderArq.sNumeroBanco = "341"
    tRegHeaderArq.sCodigoLote = "0000"
    tRegHeaderArq.sTipoRegistro = "0"
    tRegHeaderArq.sComplemento1 = FormataCpoTexto("", 6)
    tRegHeaderArq.sLayoutArq = "020"
    If Len(objFilialEmpresa.sCgc) = 14 Then
        tRegHeaderArq.sTipoInscEmpresa = "2"
    Else
        tRegHeaderArq.sTipoInscEmpresa = "1"
    End If
    tRegHeaderArq.sCGCEmpresaDebitada = FormataCpoNum(objFilialEmpresa.sCgc, 14)
    tRegHeaderArq.sComplemento2 = FormataCpoTexto("", 20)
    tRegHeaderArq.sAgencia = FormataCpoNum(gobjGeracaoArqCNABPag.objCCI.sAgencia, 5)
    tRegHeaderArq.sComplemento3 = FormataCpoTexto("", 1)
    tRegHeaderArq.sConta = FormataCpoNum(gobjGeracaoArqCNABPag.objCCI.sNumConta, 12)
    tRegHeaderArq.sComplemento4 = FormataCpoTexto("", 1)
    tRegHeaderArq.sDACAgDebitada = FormataCpoTexto(gobjGeracaoArqCNABPag.objCCI.sDVNumConta, "1")
    
    tRegHeaderArq.sNomeEmpresa = FormataCpoTexto(gsNomeEmpresa, 30)

    Call CF("Empresa_Le_Nome_Customizado", tRegHeaderArq.sNomeEmpresa)
    
    tRegHeaderArq.sNomeBanco = FormataCpoTexto("ITAU", 30)
    tRegHeaderArq.sComplemento5 = FormataCpoTexto("", 10)
    tRegHeaderArq.sArquivoCodigo = "1"
    tRegHeaderArq.sDataGeracao = Format(gdtDataHoje, "DDMMYYYY")
    tRegHeaderArq.sHoraGeracao = Format(Time, "HHMMSS")
    tRegHeaderArq.sZeros = FormataCpoNum("", 9)
    tRegHeaderArq.sUnidadeDensidade = FormataCpoNum("", 5)
    tRegHeaderArq.sComplemento6 = FormataCpoTexto("", 69)
    
    'Forma o registro com os dados recolhidos p\ o cabeçalho
    With tRegHeaderArq
        sRegistro = .sNumeroBanco & .sCodigoLote & .sTipoRegistro & .sComplemento1 & .sLayoutArq & .sTipoInscEmpresa & .sCGCEmpresaDebitada & .sComplemento2
        sRegistro = sRegistro & .sAgencia & .sComplemento3 & .sConta
        sRegistro = sRegistro & .sComplemento4 & .sDACAgDebitada & .sNomeEmpresa & .sNomeBanco & .sComplemento5 & .sArquivoCodigo & .sDataGeracao
        sRegistro = sRegistro & .sHoraGeracao & .sZeros & .sUnidadeDensidade & .sComplemento6
    End With
    
    'Inclui no arquivo
    Print #1, DesacentuaTexto(sRegistro)
    
    Montar_RegistroHeaderCNABPag = SUCESSO
    
    Exit Function
    
Erro_Montar_RegistroHeaderCNABPag:

    Montar_RegistroHeaderCNABPag = Err
    
    Select Case Err
    
        Case 51736
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147274)
            
    End Select
    
    Exit Function
    
End Function

Function Montar_Registro_HeaderLote(ByVal objLote As ClassLoteArqCNABPag) As Long
'Monta o registro de header de um lote de contas a pagar

Dim lErro As Long, sRegistro As String
Dim tRegHeaderLote As typePagtoHeaderLote
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_Montar_Registro_HeaderLote
    
    Call objLote.ObterProxReg
    
    'Busca os dados da filialempresa
    lErro = Obter_FilialEmpresa(giFilialEmpresa, objFilialEmpresa)
    If lErro <> SUCESSO Then Error 51737

    'Recolhe os dados do header do Lote
    tRegHeaderLote.sNumeroBanco = "341"
    tRegHeaderLote.sCodigoLote = FormataCpoNum(objLote.lSequencial, 4)
    tRegHeaderLote.sTipoRegistro = "1"
    tRegHeaderLote.sTipoOperacao = "C"
    tRegHeaderLote.sTipoPagamento = "98"
     
    If objLote.iTipoCobranca = TIPO_COBRANCA_OP Then
        tRegHeaderLote.sFormaPagamento = "10"
    ElseIf objLote.iTipoCobranca = TIPO_COBRANCA_DOC Then
        tRegHeaderLote.sFormaPagamento = "03"
    ElseIf objLote.iTipoCobranca = TIPO_COBRANCA_DEP_CONTA Or objLote.iTipoCobranca = TIPO_COBRANCA_CARTEIRA Then
        tRegHeaderLote.sFormaPagamento = "01"
    ElseIf objLote.iTipoCobranca = TIPO_COBRANCA_BANCARIA Then
        If objLote.iBancoCobrador = gobjGeracaoArqCNABPag.objCCI.iCodBanco Then
            tRegHeaderLote.sFormaPagamento = "30"
        Else
            tRegHeaderLote.sFormaPagamento = "31"
        End If
    End If
    
    tRegHeaderLote.sLayoutLote = "020"
    tRegHeaderLote.sComplemento1 = FormataCpoTexto("", 1)
    If Len(objFilialEmpresa.sCgc) = 14 Then
        tRegHeaderLote.sTipoInscEmpresa = "2"
    Else
        tRegHeaderLote.sTipoInscEmpresa = "1"
    End If
    tRegHeaderLote.sCGCEmpresaDebitada = FormataCpoNum(objFilialEmpresa.sCgc, 14)
    tRegHeaderLote.sComplemento2 = FormataCpoTexto("", 20)
    tRegHeaderLote.sAgencia = FormataCpoNum(gobjGeracaoArqCNABPag.objCCI.sAgencia, 5)
    tRegHeaderLote.sComplemento3 = FormataCpoTexto("", 1)
    tRegHeaderLote.sContaDebitada = FormataCpoNum(gobjGeracaoArqCNABPag.objCCI.sNumConta, 12)
    tRegHeaderLote.sComplemento4 = FormataCpoTexto("", 1)
    tRegHeaderLote.sDACAgDebitada = FormataCpoTexto(gobjGeracaoArqCNABPag.objCCI.sDVNumConta, "1")
    
    tRegHeaderLote.sNomeEmpresa = FormataCpoTexto(gsNomeEmpresa, 30)
    
    tRegHeaderLote.sFinalidadeLote = FormataCpoTexto("", 30)
    tRegHeaderLote.sHistorico = FormataCpoTexto("", 10)
    tRegHeaderLote.sEnderecoEmpresa = FormataCpoTexto("", 30)
    tRegHeaderLote.sNumeroLocal = FormataCpoNum("", 5)
    tRegHeaderLote.sComplementoEndereco = FormataCpoTexto("", 15)
    tRegHeaderLote.sCidade = FormataCpoTexto(objFilialEmpresa.objEndereco.sCidade, 20)
    tRegHeaderLote.sCEP = Formata_Cep(objFilialEmpresa.objEndereco.sCEP, 8)
    tRegHeaderLote.sEstado = FormataCpoTexto(objFilialEmpresa.objEndereco.sSiglaEstado, 2)
    tRegHeaderLote.sComplemento5 = FormataCpoTexto("", 8)
    tRegHeaderLote.sOcorrenciasRetorno = FormataCpoTexto("", 10)
    
    'Monta o registro de acordo com os dados recolhidos
    With tRegHeaderLote
    
        'alteracao por tulio301202... a atribuicao estava sendo feita num momento improprio..
        'gerando problemas na hora de gravar o arquivo
        .sCGCEmpresaDebitada = FormataCpoNum(objFilialEmpresa.sCgc, 14)
    
        sRegistro = .sNumeroBanco & .sCodigoLote & .sTipoRegistro & "C" & .sTipoPagamento & .sFormaPagamento & .sLayoutLote
        sRegistro = sRegistro & .sComplemento1 & .sTipoInscEmpresa & .sCGCEmpresaDebitada
        sRegistro = sRegistro & .sComplemento2 & .sAgencia & .sComplemento3 & .sContaDebitada & .sComplemento4 & .sDACAgDebitada
        sRegistro = sRegistro & .sNomeEmpresa & .sFinalidadeLote & .sHistorico & .sEnderecoEmpresa & .sNumeroLocal & .sComplementoEndereco
        sRegistro = sRegistro & .sCidade & .sCEP & .sEstado & .sComplemento5 & .sOcorrenciasRetorno

    End With
    
    'Inclui no arquivo
    Print #1, DesacentuaTexto(sRegistro)
       
    Montar_Registro_HeaderLote = SUCESSO
    
    Exit Function

Erro_Montar_Registro_HeaderLote:

    Montar_Registro_HeaderLote = Err

    Select Case Err
    
        Case 51737
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147275)
            
    End Select

    Exit Function

End Function

Function Montar_Registro_SegmentoA(ByVal objDet As ClassDetArqCNABPag) As Long
'Monta o registro de header de um lote de contas a pagar

Dim lErro As Long, sRegistro As String
Dim tRegSegmentoA As typePagtoSegmentoA
Dim sDV As String, sConta As String, sAgencia As String

On Error GoTo Erro_Montar_Registro_SegmentoA

    If (Len(Trim(objDet.sAgenciaFavorecido)) = 0 Or Len(Trim(objDet.sContaFavorecido)) = 0 Or objDet.iBancoFavorecido = 0) Then gError 106575
    
    Call objDet.objLote.ObterProxReg
    
    'Recolhe os dados do Segmento A do detalhe do lote
    tRegSegmentoA.sNumeroBanco = "341"
    tRegSegmentoA.sCodigoLote = FormataCpoNum(objDet.objLote.lSequencial, 4)
    tRegSegmentoA.sTipoRegistro = "3"
    tRegSegmentoA.sNumeroRegistro = FormataCpoNum(objDet.lSequencial, 5)
    tRegSegmentoA.sSegmento = "A"
    tRegSegmentoA.sTipoMovimento = "000"
    tRegSegmentoA.sZeros1 = FormataCpoNum(0, 3)
    tRegSegmentoA.sBancoFavorecido = FormataCpoNum(objDet.iBancoFavorecido, 3)
    
    
    If objDet.iBancoFavorecido = 341 Then
        Call Extrair_Agencia_DV(objDet.sAgenciaFavorecido, sAgencia, 4, sDV, 1)
        Call Extrair_Conta_DV(objDet.sContaFavorecido, sConta, 5, sDV, 1)
        tRegSegmentoA.sAgenciaConta = "0" & FormataCpoTexto(sAgencia, 4) & FormataCpoTexto("", 1) & FormataCpoNum("", 7) & FormataCpoNum(sConta, 5) & " " & FormataCpoTexto(sDV, 1)
    Else
        Call Extrair_Agencia_DV(objDet.sAgenciaFavorecido, sAgencia, 5, sDV, 1)
        Call Extrair_Conta_DV(objDet.sContaFavorecido, sConta, 12, sDV, 1)
        tRegSegmentoA.sAgenciaConta = FormataCpoTexto(sAgencia, 5) & FormataCpoTexto("", 1) & FormataCpoNum(sConta, 12) & " " & FormataCpoTexto(sDV, 1)
    End If
    tRegSegmentoA.sNomeFavorecido = FormataCpoTexto(objDet.sNomeFavorecido, 30)
    tRegSegmentoA.sSeuNumero = FormataCpoTexto(CStr(objDet.lNumTitulo) & SEPARADOR & CStr(objDet.iNumParcela) & SEPARADOR & CStr(objDet.lFornecedor) & SEPARADOR & CStr(objDet.iSeqBaixaParcPag), 20)
    tRegSegmentoA.sDataPagamento = Format(gobjGeracaoArqCNABPag.objBorderoPagto.dtDataVencimento, "DDMMYYYY")
    tRegSegmentoA.sTipoMoeda = "REA"
    tRegSegmentoA.sZeros2 = FormataCpoNum("", 15)
    tRegSegmentoA.sValorPagamento = FormataCpoValor(objDet.dValorPagto, 15)
    tRegSegmentoA.sNossoNumero = FormataCpoTexto("", 15)
    tRegSegmentoA.sComplemento1 = FormataCpoTexto("", 5)
    tRegSegmentoA.sDataEfetiva = FormataCpoNum("", 8)
    tRegSegmentoA.sValorEfetivo = FormataCpoNum("", 15)
    tRegSegmentoA.sFinalidadeDetalhe = FormataCpoTexto("", 18)
    tRegSegmentoA.sComplemento2 = FormataCpoTexto("", 34)
    tRegSegmentoA.sAviso = FormataCpoTexto("", 1)
    tRegSegmentoA.sOcorrenciasRetorno = FormataCpoTexto("", 10)

    'Monta o registro A
    With tRegSegmentoA
        
        sRegistro = .sNumeroBanco & .sCodigoLote & .sTipoRegistro & .sNumeroRegistro & .sSegmento & .sTipoMovimento
        sRegistro = sRegistro & .sZeros1 & .sBancoFavorecido & .sAgenciaConta & .sNomeFavorecido & .sSeuNumero
        sRegistro = sRegistro & .sDataPagamento & .sTipoMoeda & .sZeros2 & .sValorPagamento & .sNossoNumero
        sRegistro = sRegistro & .sComplemento1 & .sDataEfetiva & .sValorEfetivo & .sFinalidadeDetalhe & .sComplemento2
        sRegistro = sRegistro & .sAviso & .sOcorrenciasRetorno
    
    End With

    'Inclui no arquivo
    Print #1, DesacentuaTexto(sRegistro)
       
    Montar_Registro_SegmentoA = SUCESSO
    
    Exit Function
    
Erro_Montar_Registro_SegmentoA:

    Montar_Registro_SegmentoA = gErr
    
    Select Case gErr
    
        Case 106575
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORN_SEM_INFO_CTA", gErr, objDet.lFornecedor, objDet.iFilialForn)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147276)
    
    End Select
    
    Exit Function
    
End Function


Function Montar_Registro_SegmentoJ(ByVal objDet As ClassDetArqCNABPag) As Long
'Recolhe os dados do Segmento J do detalhe do lote

Dim lErro As Long, sRegistro As String, sCodigoDeBarras As String
Dim tRegSegmentoJ As typePagtoSegmentoJ, bDigitacao As Boolean, iDVTeste As Integer

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
        .sNumeroBanco = "341"
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
                .sValor = FormataCpoValor(StrParaDbl(Mid(sCodigoDeBarras, 34)) / 100, 14)
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
        .sSeuNumero = FormataCpoTexto(CStr(objDet.lNumTitulo) & SEPARADOR & CStr(objDet.iNumParcela) & SEPARADOR & CStr(objDet.lFornecedor) & SEPARADOR & CStr(objDet.iSeqBaixaParcPag), 20)
        .sComplemento = FormataCpoTexto("", 13)
        .sNossoNumero = FormataCpoTexto("", 15)
        .sOcorrenciasRetorno = FormataCpoTexto("", 10)
    
        'Monta o registro J
        sRegistro = .sNumeroBanco & .sCodigoLote & .sTipoRegistro & .sNumeroRegistro & .sSegmento & .sTipoMovimento & .sBancoFavorecido & .sMoeda
        sRegistro = sRegistro & .sDV & .sValor & .sCampoLivre & .sNomeFavorecido & .sDataVencto & .sValorTitulo & .sDescontos & .sAcrescimos
        sRegistro = sRegistro & .sDataPagamento & .sValorPagamento & .sZeros & .sSeuNumero & .sComplemento & .sNossoNumero & .sOcorrenciasRetorno

    End With
    
    'Inclui no arquivo
    Print #1, DesacentuaTexto(sRegistro)
       
    Montar_Registro_SegmentoJ = SUCESSO
    
    Exit Function
    
Erro_Montar_Registro_SegmentoJ:

    Montar_Registro_SegmentoJ = gErr
    
    Select Case gErr
    
        Case 106574
            Call Rotina_Erro(vbOKOnly, "ERRO_TITPAG_SEM_CODBARRAS", gErr, objDet.lNumTitulo, objDet.iNumParcela, objDet.lFornecedor)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147277)
            
    End Select
    
    Exit Function
    
End Function

Function Montar_Registro_TrailerLote(ByVal objLote As ClassLoteArqCNABPag) As Long
'MOnta o regsitro de trailer do lote

Dim lErro As Long, sRegistro As String
Dim tRegTrailerLote As typePagtoTrailerLoteItau

On Error GoTo Erro_Montar_Registro_TrailerLote

    Call objLote.ObterProxReg
    
    'Reune os dados do trailer
    With tRegTrailerLote
    
        .sNumeroBanco = "341"
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147278)
            
    End Select

    Exit Function

End Function

Function Montar_Registro_TrailerArq(objArq As ClassArqCNABPag) As Long
'MOnta o regsitro de trailer do arquivo de pagtos

Dim lErro As Long, sRegistro As String
Dim tRegTrailerArq As typePagtoTrailerArqItau

On Error GoTo Erro_Montar_Registro_TrailerArq

    Call objArq.ObterProxReg
    
    'Reune os dados do trailer
    With tRegTrailerArq
        
        .sNumeroBanco = "341"
        .sCodigoLote = "9999"
        .sTipoRegistro = "9"
        .sComplemento1 = FormataCpoTexto("", 9)
        .sTotalQtdLotes = FormataCpoNum(objArq.lSequencialLote, 6)
        .sTotalQtdRegistros = FormataCpoValor(objArq.lQuantRegistros, 6)
        .sComplemento2 = FormataCpoTexto("", 211)
    
        'Monta o registro de trailer de lote
        sRegistro = .sNumeroBanco & .sCodigoLote & .sTipoRegistro & .sComplemento1 & .sTotalQtdLotes & .sTotalQtdRegistros & .sComplemento2
        
    End With

    'Inclui no arquivo
    Print #1, DesacentuaTexto(sRegistro)
       
    Montar_Registro_TrailerArq = SUCESSO
    
    Exit Function
    
Erro_Montar_Registro_TrailerArq:

    Montar_Registro_TrailerArq = Err
    
    Select Case Err
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147279)
            
    End Select
    
    Exit Function
    
End Function
Private Sub Extrair_Agencia_DV(sAgenciaDVAg As String, sAgencia As String, iTamanhoAg As Integer, sDV As String, iTamanhoDV As Integer)
'Extrai da string passada na forma "Agencia-DV", a agência e o DV separadamente

Dim iPosSeparador As Integer

    'Procura o separador da Ag e o DV
    iPosSeparador = InStr(1, sAgenciaDVAg, SEPARADOR)
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

Function Calcula_NossoNumero_Itau(objCarteiraCobrador As ClassCarteiraCobrador, sNossoNumero As String) As Long

Dim sProxNossoNumero As String

On Error GoTo Erro_Calcula_NossoNumero_Itau

    Select Case objCarteiraCobrador.iNumCarteiraNoBanco
    
        Case 112, 212, 104, 147, 105, 114, 166
            sNossoNumero = 0
        
        Case 108, 110, 111, 210, 168, 150, 126, 121, 221, 180
            sNossoNumero = objCarteiraCobrador.sFaixaNossoNumeroProx
            If Val(sNossoNumero) > Val(objCarteiraCobrador.sFaixaNossoNumeroProx) Then Error 62056
            objCarteiraCobrador.sFaixaNossoNumeroProx = sNossoNumero + 1
        
        Case Else
            sNossoNumero = objCarteiraCobrador.sFaixaNossoNumeroProx
            sProxNossoNumero = Val(sNossoNumero) + 1
            If sProxNossoNumero > objCarteiraCobrador.sFaixaNossoNumeroFinal Then sProxNossoNumero = FormataCpoNum(1, 8)
            objCarteiraCobrador.sFaixaNossoNumeroProx = sProxNossoNumero
        
    End Select

    Calcula_NossoNumero_Itau = SUCESSO
    
    Exit Function
    
Erro_Calcula_NossoNumero_Itau:

    Calcula_NossoNumero_Itau = Err
    
    Select Case Err
    
        Case 62056
            Call Rotina_Erro(vbOKOnly, "ERRO_FAIXA_NOSSONUMERO_INSUFICIENTE", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, Err, 147280)
    
    End Select

    Exit Function

End Function

Private Function Separa_Borderos_ComSem_Emissao(colBorderos As Collection, colBorderosComEmiss As Collection, colBorderosSemEmiss As Collection) As Long
'Separa os borderos das carteiras com registro das carteiras sem registro.

Dim lErro As Long
Dim objBordero As ClassBorderoCobranca
Dim colCodCarteiras As New Collection
Dim iCodigo As Integer
Dim bAchou As Boolean
Dim iIndice As Integer

On Error GoTo Erro_Separa_Borderos_ComSem_Emissao
    
    lErro = CF("CarteirasItauEmissao_Le_Cobrador", gobjCobrancaEletronica.iCobrador, colCodCarteiras)
    If lErro <> SUCESSO Then Error 62053
    
    'Para cada bordero da coleção lida
    For Each objBordero In colBorderos
        bAchou = False
        
        For iIndice = 1 To colCodCarteiras.Count
            iCodigo = colCodCarteiras(iIndice)
            
            If objBordero.iCodCarteiraCobranca = iCodigo Then
                bAchou = True
                colBorderosSemEmiss.Add objBordero
            End If
        Next
        
        If Not bAchou Then
            colBorderosComEmiss.Add objBordero
        End If
    Next
    
    Separa_Borderos_ComSem_Emissao = SUCESSO
    
    Exit Function
    
Erro_Separa_Borderos_ComSem_Emissao:
    
    Separa_Borderos_ComSem_Emissao = Err
    
    Select Case Err
    
        Case 62015, 62053
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147281)
            
    End Select

    Exit Function
    
End Function

Function CarteirasItauEmissao_Le_Cobrador(ByVal iCobrador As Integer, ByVal colCodCarteiras As Collection) As Long
'obtem os codigos das carteiras sem registro com emissao de bloquetos p/Itau com layout de acordo com anexo A

Dim lComando As Long
Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_CarteirasItauEmissao_Le_Cobrador

    lComando = Comando_Abrir
    If lComando = 0 Then Error 62057
    
    lErro = Comando_Executar(lComando, "SELECT CodCarteiraCobranca FROM CarteirasCobrador WHERE Cobrador = ? AND NumCarteiraNoBanco IN (106,196,103,173,195,102,172) ", iCodigo, iCobrador)
    If lErro <> AD_SQL_SUCESSO Then Error 62058

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 62059
        
    Do While lErro <> AD_SQL_SEM_DADOS
    
        colCodCarteiras.Add iCodigo
    
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 62060
        
    Loop
    
    Call Comando_Fechar(lComando)
    
    CarteirasItauEmissao_Le_Cobrador = SUCESSO
    
    Exit Function
    
Erro_CarteirasItauEmissao_Le_Cobrador:

    CarteirasItauEmissao_Le_Cobrador = Err
    
    Select Case Err
    
        Case 62057
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 62058, 62059, 62060
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CARTEIRAS_COBRADOR1", Err, iCobrador)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147282)
    
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function Gera_Arquivo_CobrancaSemEmiss(lComando As Long, lComando1 As Long, lComando2 As Long, lComando3 As Long)
'gera arquivo com layout do anexo A

Dim iCancela As Integer
Dim sNomeArquivo As String
Dim lErro As Long
Dim sRegHeader As String
Dim sRegDetalhe As String
Dim sRegTrailer As String, lSeqNomeArq As Long
Dim objBorderoCobranca As New ClassBorderoCobranca
Dim objCarteiraCobrador As New ClassCarteiraCobrador
Dim objOcorrRemParc As ClassOcorrRemParcRec

On Error GoTo Erro_Gera_Arquivo_CobrancaSemEmiss

    'Gera o nome do arquivo de acordo com o sequencial do Cobrador
    lErro = CF("CobrancaEletronica_Gera_NomeArquivo", gobjCobrancaEletronica.objCobrador, gobjCobrancaEletronica.sDiretorio, sNomeArquivo, lSeqNomeArq)
    If lErro <> SUCESSO Then gError 51652
    
    gobjCobrancaEletronica.lSeqNomeArq = lSeqNomeArq
    
    'Cria o arquivo de remessa
    Open sNomeArquivo For Output As #1
        
    lErro = Montar_RegistroHeader_Itau(sRegHeader)
    If lErro <> SUCESSO Then gError 51676
        
    'Inclui o registro de Header
    Print #1, DesacentuaTexto(sRegHeader)

    For Each objBorderoCobranca In gobjCobrancaEletronica.colBorderos
        'Lê o bordero de cobrança
        
        lErro = CF("BorderoCobranca_LeMod", lComando, objBorderoCobranca)
        If lErro <> SUCESSO Then gError 51657
            
        objCarteiraCobrador.iCobrador = gobjCobrancaEletronica.iCobrador
        objCarteiraCobrador.iCodCarteiraCobranca = objBorderoCobranca.iCodCarteiraCobranca
        
        'Lê a carteira do cobrador
        lErro = CF("CarteiraCobrador_LeMod", objCarteiraCobrador, lComando1)
        If lErro <> SUCESSO Then gError 51669
       
        'Busca no BD as Ocorrências do Bordero em questão
        lErro = CF("CobrancaEletronica_Obter_OcorrRemParcRec", lComando2, objBorderoCobranca, gobjCobrancaEletronica)
        If lErro <> SUCESSO Then gError 51662
        
        'Para cada ocorrência
        For Each objOcorrRemParc In gobjCobrancaEletronica.colOcorrRemParcRec
    
            sRegDetalhe = ""
            
            'Monta um registro de detalhe do arquivo de remessa
            lErro = Montar_RegistroDetalheEmiss1_Itau(sRegDetalhe, objOcorrRemParc, objCarteiraCobrador)
            If lErro <> SUCESSO Then gError 51677
            
            'Inclui o registro com os instrucoes do Título(Parcela) a receber
            Print #1, DesacentuaTexto(sRegDetalhe)
            
            If objOcorrRemParc.iCodOcorrencia = COBRANCA_OCORR_INC_TITULO And (objOcorrRemParc.dValorCobrado * gobjCobrancaEletronica.dPercJurosDiario) <> 0 Then
            
                sRegDetalhe = "62APOS VCTO MORA DIARIA DE R$ " & Format(Round(objOcorrRemParc.dValorCobrado * gobjCobrancaEletronica.dPercJurosDiario, 2), "Standard")
                sRegDetalhe = FormataCpoTexto(sRegDetalhe, 2 + (69 * 5) + 47) & FormataCpoNum(gobjCobrancaEletronica.lSequencialRegistro, 6)
                
                'Incrementa o sequencial de registro
                gobjCobrancaEletronica.lSequencialRegistro = gobjCobrancaEletronica.lSequencialRegistro + 1
        
                'Inclui o registro com os detalhes do Título(Parcela) a receber
                Print #1, DesacentuaTexto(sRegDetalhe)
                
            End If
            
        Next
        
        objBorderoCobranca.iProcessado = BORDERO_PROCESSADO
        objBorderoCobranca.sNomeArquivo = sNomeArquivo
        
        'Atualiza o Status do Bordero de Cobrança
        lErro = CF("BorderoCobranca_Atualiza_Processo", lComando, lComando2, objBorderoCobranca)
        If lErro <> SUCESSO Then gError 51663
        
        'Atualiza o NossoNumero em carteiraCobrador
        lErro = CF("CarteiraCobrador_AtualizaNossoNumero", objCarteiraCobrador, lComando1, lComando3)
        If lErro <> SUCESSO Then gError 51707

        iCancela = 0
        
        Call gobjCobrancaEletronica.objTelaAtualizacao.Mostra_Evolucao(iCancela, 1)
        
        If iCancela = CANCELA_BATCH Then gError 79707
    
    Next

    'MOnta o registro de trailer
    lErro = Montar_RegistroTrailer_Itau(sRegTrailer, objOcorrRemParc)
    If lErro <> SUCESSO Then gError 51678

    'Coloca o registro de trailer no arquivo
    Print #1, DesacentuaTexto(sRegTrailer)
    
    Close #1
    
    Gera_Arquivo_CobrancaSemEmiss = SUCESSO
    
    Exit Function

Erro_Gera_Arquivo_CobrancaSemEmiss:

    Gera_Arquivo_CobrancaSemEmiss = gErr
    
    Select Case gErr
    
        Case 51652, 51676, 51657, 51669, 51662, 51663, 51677, 51707, 51678, 79707
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147283)
            
    End Select
        
    Exit Function
        
End Function
Function Gera_Arquivo_CobrancaComEmiss(lComando As Long, lComando1 As Long, lComando2 As Long, lComando3 As Long)
'gera arquivo para layout normal (nao o do anexo A)

Dim iCancela As Integer
Dim sNomeArquivo As String
Dim lErro As Long
Dim sRegHeader As String
Dim sRegDetalhe As String
Dim sRegTrailer As String, lSeqNomeArq As Long
Dim objBorderoCobranca As New ClassBorderoCobranca
Dim objCarteiraCobrador As New ClassCarteiraCobrador
Dim objOcorrRemParc As ClassOcorrRemParcRec

On Error GoTo Erro_Gera_Arquivo_CobrancaComEmiss

    'Gera o nome do arquivo de acordo com o sequencial do Cobrador
    lErro = CF("CobrancaEletronica_Gera_NomeArquivo", gobjCobrancaEletronica.objCobrador, gobjCobrancaEletronica.sDiretorio, sNomeArquivo, lSeqNomeArq)
    If lErro <> SUCESSO Then gError 62023
    
    gobjCobrancaEletronica.lSeqNomeArq = lSeqNomeArq
    
    'Cria o arquivo de remessa
    Open sNomeArquivo For Output As #1
        
    lErro = Montar_RegistroHeader_Itau(sRegHeader)
    If lErro <> SUCESSO Then gError 62024
        
    'Inclui o registro de Header
    Print #1, DesacentuaTexto(sRegHeader)

    For Each objBorderoCobranca In gobjCobrancaEletronica.colBorderos
        'Lê o bordero de cobrança
        
        lErro = CF("BorderoCobranca_LeMod", lComando, objBorderoCobranca)
        If lErro <> SUCESSO Then gError 62025
            
        objCarteiraCobrador.iCobrador = gobjCobrancaEletronica.iCobrador
        objCarteiraCobrador.iCodCarteiraCobranca = objBorderoCobranca.iCodCarteiraCobranca
        
        'Lê a carteira do cobrador
        lErro = CF("CarteiraCobrador_LeMod", objCarteiraCobrador, lComando1)
        If lErro <> SUCESSO Then gError 62026
       
        'Busca no BD as Ocorrências do Bordero em questão
        lErro = CF("CobrancaEletronica_Obter_OcorrRemParcRec", lComando2, objBorderoCobranca, gobjCobrancaEletronica)
        If lErro <> SUCESSO Then gError 62027
        
        'Para cada ocorrência
        For Each objOcorrRemParc In gobjCobrancaEletronica.colOcorrRemParcRec
    
            sRegDetalhe = ""
            'Monta um registro de detalhe do arquivo de remessa
            lErro = Montar_RegistroDetalhe_Itau(sRegDetalhe, objOcorrRemParc, objCarteiraCobrador)
            If lErro <> SUCESSO And lErro <> 32261 Then gError 62029
            
            If lErro = SUCESSO Then
            
                'Inclui o registro com os detalhes do Título(Parcela) a receber
                Print #1, DesacentuaTexto(sRegDetalhe)
                
            End If
            
        Next
        
        objBorderoCobranca.iProcessado = BORDERO_PROCESSADO
        objBorderoCobranca.sNomeArquivo = sNomeArquivo
        
        'Atualiza o Status do Bordero de Cobrança
        lErro = CF("BorderoCobranca_Atualiza_Processo", lComando, lComando2, objBorderoCobranca)
        If lErro <> SUCESSO Then gError 62028
    
        'Atualiza o NossoNumero em carteiraCobrador
        lErro = CF("CarteiraCobrador_AtualizaNossoNumero", objCarteiraCobrador, lComando1, lComando3)
        If lErro <> SUCESSO Then gError 62030
    
        iCancela = 0
        
        Call gobjCobrancaEletronica.objTelaAtualizacao.Mostra_Evolucao(iCancela, 1)
        
        If iCancela = CANCELA_BATCH Then gError 79708
        
    Next
    
    'MOnta o registro de trailer
    lErro = Montar_RegistroTrailer_Itau(sRegTrailer, objOcorrRemParc)
    If lErro <> SUCESSO Then gError 62031

    'Coloca o registro de trailer no arquivo
    Print #1, DesacentuaTexto(sRegTrailer)
    
    Close #1
    
    Gera_Arquivo_CobrancaComEmiss = SUCESSO
    
    Exit Function

Erro_Gera_Arquivo_CobrancaComEmiss:

    Gera_Arquivo_CobrancaComEmiss = gErr
    
    Select Case gErr
    
        Case 62023, 62024, 62025, 62026, 62027, 62028, 62029, 62030, 62031, 79708
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147284)
            
    End Select

    Exit Function

End Function

Private Function Montar_RegistroDetalheEmiss1_Itau(sRegDetalhe As String, objOcorrRemParcRec As ClassOcorrRemParcRec, objCarteiraCobrador As ClassCarteiraCobrador) As Long
'Monta os registros detalhe p/anexo A

Dim tRegDetalhe As typeRemDetalheItau
Dim lErro As Long

On Error GoTo Erro_Montar_RegistroDetalheEmiss1_Itau
    
    'Obtem os dados do registro de detalhe
    lErro = Obter_Dados_Detalhe_Emiss1(tRegDetalhe, objOcorrRemParcRec, objCarteiraCobrador)
    If lErro <> SUCESSO Then Error 51682
    
    'Forma o registro de detalhe
    With tRegDetalhe
        sRegDetalhe = .sIdentificacaoReg & .sIdentificacaoEmp & .sNumCarteiraBanco & .sNossoNumero & .sQuantMoeda
        sRegDetalhe = sRegDetalhe & .sValorTitulo & .sSeuNumero & .sVencimento & .sEspecie & tRegDetalhe.sAceite
        sRegDetalhe = sRegDetalhe & .sDataEmissao & .sCodigoInscSacado & .sNumInscricaoSacado & .sNome & .sComplemento2
        sRegDetalhe = sRegDetalhe & .sLogradouro & .sBairro & .sCEP & .sCidade & .sEstado & .sSacadorAvalista
        sRegDetalhe = sRegDetalhe & .sTipoInscricaoEmpSac & .sNumInscricaoEmpSac & .sComplemento4 & .sNumSequencial

    End With
    
    Montar_RegistroDetalheEmiss1_Itau = SUCESSO

    Exit Function

Erro_Montar_RegistroDetalheEmiss1_Itau:

    Montar_RegistroDetalheEmiss1_Itau = Err

    Select Case Err
    
        Case 51682

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147285)

    End Select

    Exit Function

End Function

Private Function Obter_Dados_Detalhe_Emiss1(tRegDetalhe As typeRemDetalheItau, objOcorrRemParcRem As ClassOcorrRemParcRec, objCarteiraCobrador As ClassCarteiraCobrador) As Long
'Obtem os dados que são utilizados para preencher o registro de detalhe de cobranca segundo o anexo A

Dim lErro As Long
Dim objParcelaRec As New ClassParcelaReceber
Dim objEndereco As New ClassEndereco
Dim sNossoNumero As String
Dim dValorDesconto As Double
Dim sDac As String, sAgencia As String

On Error GoTo Erro_Obter_Dados_Detalhe_Emiss1

    objParcelaRec.lNumIntDoc = objOcorrRemParcRem.lNumIntParc

    'Busca os dados da parcela a receber no BD
    lErro = CF("CobrancaEletronica_Obter_DadosParcelaRec", objParcelaRec, gobjCobrancaEletronica)
    If lErro <> SUCESSO Then Error 62020
    
    lErro = CF("CobrancaEletronica_Obter_Endereco", gobjCobrancaEletronica, objEndereco)
    If lErro <> SUCESSO Then Error 62021
    
    'nosso numero = filialempresa + num parcela + num titulo
    sNossoNumero = CStr((gobjCobrancaEletronica.iFilialEmpTitulo * 10000000) + (objParcelaRec.iNumParcela * 1000000) + gobjCobrancaEletronica.lNumTitulo)
    
    tRegDetalhe.sIdentificacaoReg = "61"
    tRegDetalhe.sIdentificacaoEmp = FormataCpoNum(gobjCobrancaEletronica.sCodEmpresaBanco, 12)
    tRegDetalhe.sNumCarteiraBanco = FormataCpoNum(objCarteiraCobrador.iNumCarteiraNoBanco, 3)
    tRegDetalhe.sNossoNumero = FormataCpoNum(sNossoNumero, 8)
    sAgencia = FormataCpoNum(gobjCobrancaEletronica.objCCI.sAgencia, 4)
    Call Calcula_DAC(sAgencia & tRegDetalhe.sNumCarteiraBanco & tRegDetalhe.sNossoNumero, sDac)
    tRegDetalhe.sNossoNumero = FormataCpoNum(sNossoNumero, 8) & sDac
    tRegDetalhe.sQuantMoeda = "0    "
    tRegDetalhe.sValorTitulo = FormataCpoValor(objOcorrRemParcRem.dValorCobrado, 13)
    
    '??? hicare
    'tRegDetalhe.sSeuNumero = FormataCpoNum(gobjCobrancaEletronica.iFilialEmpTitulo, 2) & FormataCpoNum(objParcelaRec.iNumParcela, 2) & FormataCpoNum(gobjCobrancaEletronica.lNumTitulo, 6)
    tRegDetalhe.sSeuNumero = FormataCpoTexto(CStr(gobjCobrancaEletronica.lNumTitulo) & SEPARADOR & CStr(objParcelaRec.iNumParcela), 10)

    Call CF("Itau_Obter_Dados_DetalheSegP_Customizado", gobjCobrancaEletronica, tRegDetalhe.sSeuNumero, objParcelaRec)

    tRegDetalhe.sVencimento = Format(objParcelaRec.dtDataVencimento, "DDMMYY")
    tRegDetalhe.sEspecie = FormataCpoNum(gobjCobrancaEletronica.iEspecieTitulo, 2)
    If objParcelaRec.iAceite = 1 Then
        tRegDetalhe.sAceite = "A"
    Else
        tRegDetalhe.sAceite = "N"
    End If
    tRegDetalhe.sDataEmissao = Format(gobjCobrancaEletronica.dtDataEmissaoTitulo, "DDMMYY")
    If Len(gobjCobrancaEletronica.sCGCSacado) = 14 Then
        tRegDetalhe.sCodigoInscSacado = "02"
    Else
        tRegDetalhe.sCodigoInscSacado = "01"
    End If
    tRegDetalhe.sNumInscricaoSacado = FormataCpoTexto(gobjCobrancaEletronica.sCGCSacado, 15)
    tRegDetalhe.sNome = FormataCpoTexto(gobjCobrancaEletronica.sNomeSacado, 30)
    tRegDetalhe.sComplemento2 = FormataCpoTexto("", 9)
    tRegDetalhe.sLogradouro = FormataCpoTexto(Trim(objEndereco.sEndereco), 40)
    tRegDetalhe.sBairro = FormataCpoTexto(Trim(objEndereco.sBairro), 12)
    tRegDetalhe.sCEP = Formata_Cep(objEndereco.sCEP, 8)
    tRegDetalhe.sCidade = FormataCpoTexto(objEndereco.sCidade, 15)
    tRegDetalhe.sEstado = FormataCpoTexto(objEndereco.sSiglaEstado, 2)
    tRegDetalhe.sSacadorAvalista = FormataCpoTexto("", 144)
    tRegDetalhe.sTipoInscricaoEmpSac = "02"
    tRegDetalhe.sNumInscricaoEmpSac = FormataCpoNum(Trim(gobjCobrancaEletronica.sCGCCedente), 15)
    tRegDetalhe.sComplemento4 = FormataCpoTexto("", 31)
    tRegDetalhe.sNumSequencial = FormataCpoNum(gobjCobrancaEletronica.lSequencialRegistro, 6)
    
    'Incrementa o sequencial de registro
    gobjCobrancaEletronica.lSequencialRegistro = gobjCobrancaEletronica.lSequencialRegistro + 1
    
    Obter_Dados_Detalhe_Emiss1 = SUCESSO

    Exit Function
    
Erro_Obter_Dados_Detalhe_Emiss1:

    Obter_Dados_Detalhe_Emiss1 = Err
    
    Select Case Err

        Case 62020, 62021, 62022
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147286)

    End Select

    Exit Function

End Function

Private Sub Calcula_DAC(sCampo As String, sDac As String)

Dim iIndice As Integer
Dim iTamCampo As Integer
Dim iDigito1 As Integer
Dim iVarModulo As Integer
Dim iProduto As Integer
Dim iSoma As Integer
Dim sSoma As String
Dim iResto As Integer
Dim iResultado As Integer
    
    iVarModulo = 2
    iSoma = 0
    sSoma = ""
    
    iTamCampo = Len(sCampo)
    
    For iIndice = iTamCampo To 1 Step -1

        iDigito1 = StrParaInt(Mid(sCampo, iIndice, 1))
        iProduto = (iDigito1 * iVarModulo)
        sSoma = sSoma & iProduto
                
        iVarModulo = iVarModulo - 1
        
        If iVarModulo = 0 Then iVarModulo = 2
    
    Next
    
    For iIndice = 1 To Len(sSoma)
        iSoma = iSoma + StrParaInt(Mid(sSoma, iIndice, 1))
    Next
    
    iResto = iSoma Mod 10
    
    iResultado = 10 - iResto
    
    If iResto = 0 Then iResultado = 0
    
    sDac = iResultado
    
    Exit Sub
    
End Sub

Private Sub RegDetCobr_ObterInfo(objDetRetCobr As ClassDetRetCobr, tRegDetalhe As typeRetDetalheItau)
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
        .iCodCarteiraNoBanco = StrParaInt(tRegDetalhe.sNumCarteira)
        .dValorTarifa = Extrair_Valor(tRegDetalhe.sTarifaCobranca)
        .sSeuNumero = tRegDetalhe.sSeuNumero
        .sUsoDaEmpresa = tRegDetalhe.sUsoDaEmpresa
        .lNumSequencial = CLng(tRegDetalhe.sNumSequencial)
        .dValorTitulo = Extrair_Valor(tRegDetalhe.sValorTitulo)
        .dValorEntregue = Extrair_Valor(tRegDetalhe.sValorEntregue)
    End With
    
End Sub

Private Function Prepara_CartCobrRet1(objCobrancaEletronica As ClassCobrancaEletronica, objBorderoRetCobr As ClassBorderoRetCobr, objBorderoRetCobrCart As ClassBorderoRetCobrCart, objDetRetCobr As ClassDetRetCobr, tRegDetalhe As typeRetDetalheItau) As Long
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147287)

    End Select

    Exit Function

End Function

Private Function RetCobr_Obtem_ParcelaItau(objBorderoRetCobr As ClassBorderoRetCobr, objDetRetCobr As ClassDetRetCobr, objParcelaRec As ClassParcelaReceber, tRegDetalhe As typeRetDetalheItau, lComando As Long) As Long
'preenche objParcelaRec lendo do bd identificando a parcela por dados em objDetRetCobr

Dim lErro As Long, sParcela As String
Dim sNumTitulo As String, sFilialEmpTit As String

On Error GoTo Erro_RetCobr_Obtem_ParcelaItau

    lErro = CF("RetCobr_Obtem_Parcela4", objBorderoRetCobr, objDetRetCobr, objParcelaRec, lComando)
    If lErro <> SUCESSO And lErro <> 130041 Then gError 81769
    If lErro <> SUCESSO Then
    
        'se é de uma carteira sem registro com remessa segundo o anexo A
        If InStr("106,196,103,173,195,102,172", tRegDetalhe.sNumCarteira) <> 0 And InStr(tRegDetalhe.sNossoNumero1, "-") = 0 Then
            
            sFilialEmpTit = Mid(tRegDetalhe.sNossoNumero1, 1, 1)
            sParcela = Mid(tRegDetalhe.sNossoNumero1, 2, 1)
            sNumTitulo = Mid(tRegDetalhe.sNossoNumero1, 3, 6)
            
        Else
        
            lErro = CF("RetCobr_Obtem_Parcela1", objDetRetCobr, sNumTitulo, sParcela, sFilialEmpTit, RETCOBR_TIPO_SEU_NUMERO1)
            If lErro <> SUCESSO Then gError 130010
                
        End If
        
        lErro = CF("RetCobr_Obtem_Parcela2", objDetRetCobr, objParcelaRec, sNumTitulo, sParcela, sFilialEmpTit, lComando)
        If lErro <> SUCESSO Then gError 81769
        
    End If
    
    RetCobr_Obtem_ParcelaItau = SUCESSO

    Exit Function

Erro_RetCobr_Obtem_ParcelaItau:

    RetCobr_Obtem_ParcelaItau = gErr

    Select Case gErr

        Case 81769, 130010
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147288)
            
    End Select

    Exit Function

End Function

Private Function Processa_ParcelaRecConfirmada(objParcelaRec As ClassParcelaReceber, objBorderoRetCobr As ClassBorderoRetCobr, objBorderoRetCobrCart As ClassBorderoRetCobrCart, objDetRetCobr As ClassDetRetCobr, tRegDetalhe As typeRetDetalheItau, lComando As Long, lComando1 As Long) As Long
'Faz o processamento da parcela a receber que recebe a confirmacao de entrada

Dim lErro As Long, dValorBaixado As Double, sNumTitCobrador As String
Dim objOcorrRetParcRec As New ClassOcorrRetParcRec
Dim iCobrador As Integer, iCarteiraCobranca As Integer

On Error GoTo Erro_Processa_ParcelaRecConfirmada

    Select Case StrParaInt(tRegDetalhe.sNumCarteira)
    
        'carteiras em que o Nosso Numero é dado pelo banco
        Case 112, 212, 104, 147, 105, 114, 166
            'atualizar nosso numero
            
            sNumTitCobrador = String(STRING_NUMTITCOBRADOR, 0)
            
            lErro = Comando_ExecutarPos(lComando, "SELECT NumTitCobrador, Cobrador, CarteiraCobranca FROM ParcelasRec WHERE NumIntDoc = ?", 0, sNumTitCobrador, iCobrador, iCarteiraCobranca, objParcelaRec.lNumIntDoc)
            If lErro <> AD_SQL_SUCESSO Then gError 81733
            
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81734
                
            'se o nosso numero informado pelo bco é <> do que estava na parcela
            If Trim(sNumTitCobrador) <> Trim(tRegDetalhe.sNossoNumero1) Then
            
                lErro = Comando_LockExclusive(lComando)
                If lErro <> AD_SQL_SUCESSO Then gError 81735
            
                'se ainda está nesse cobrador (um teste mais preciso checaria tb a carteira)
                If iCobrador = objBorderoRetCobr.iCobrador Then
                
                    'atualiza o nosso numero na parcela
                    lErro = Comando_ExecutarPos(lComando1, "UPDATE ParcelasRec SET NumTitCobrador = ?", lComando, Trim(tRegDetalhe.sNossoNumero1))
                    If lErro <> AD_SQL_SUCESSO Then gError 81736
                
                End If
                
            End If
                    
    End Select
        
    Processa_ParcelaRecConfirmada = SUCESSO
    
    Exit Function

Erro_Processa_ParcelaRecConfirmada:

    Processa_ParcelaRecConfirmada = gErr
    
    Select Case gErr
    
        Case 81733, 81734
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PARCELAS_REC", gErr)
        
        Case 81735
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_PARCELAS_REC", gErr)
        
        Case 81736
            Call Rotina_Erro(vbOKOnly, "ERRO_MODIFICACAO_PARCELAS_REC", gErr)
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147289)
            
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147290)
            
    End Select

End Function

Private Function CalculaDVCampoMod10(sCampo As String, iDV As Integer) As Long
'Calcula o dv de um campo especifico do codigo de barras
'O Codigo de barras tem 5 campos
'Os campos passiveis de serem calculados sao 1, 2, 3
'sCampo eh a string que traz o campo em questao (INPUT)
'iDv eh o inteiro que retorna o dv (OUTPUT)

Dim iCont As String
Dim lACC As Long
Dim iIndex As Integer
Dim iNumCorrente As Integer
Dim iIndice As Integer

On Error GoTo Erro_CalculaDVCampoMod10

    'inicializa a variavel multiplicadora
    iCont = 0
    
    'inicializa o acumulador
    lACC = 0
    
    'inicializa o indice do mod
    iIndice = 1

    'percorre o codbarras de traz pra frente
    For iIndex = Len(sCampo) - 1 To 1 Step -1
    
        'obtem o numero corrente
        iNumCorrente = StrParaInt(Mid(sCampo, iIndex, iIndex + 1))
    
        'atualiza a variavel que tera de rodar a sequencia
        '21212121.......
        iCont = (iIndice Mod 2) + 1
    
        'acumula o numero corrente * a variavel multiplicadora
        lACC = lACC + (iNumCorrente * iCont)
        
        'incrementa o indice
        iIndice = iIndice + 1
        
    Next

    'calcula finalmente o dv
    iDV = 10 - (lACC Mod 10)
    
    CalculaDVCampoMod10 = SUCESSO

    Exit Function

Erro_CalculaDVCampoMod10:

    CalculaDVCampoMod10 = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147291)
            
    End Select

End Function


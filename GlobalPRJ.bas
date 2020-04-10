Attribute VB_Name = "GlobalPRJ"
Option Explicit

Public Const MNEMONICO_MALADIRETA_TIPO_PROPOSTA = 1
Public Const MNEMONICO_MALADIRETA_TIPO_CONTRATO = 2
Public Const MNEMONICO_MALADIRETA_TIPO_OV = 3

Public Const MNEMONICO_MALADIRETA_TIPOOBJ_OUTROS = 0
Public Const MNEMONICO_MALADIRETA_TIPOOBJ_PROJETO = 1
Public Const MNEMONICO_MALADIRETA_TIPOOBJ_CLIENTE = 2
Public Const MNEMONICO_MALADIRETA_TIPOOBJ_PROPOSTA = 3
Public Const MNEMONICO_MALADIRETA_TIPOOBJ_ESCOPO = 4
Public Const MNEMONICO_MALADIRETA_TIPOOBJ_ENDERECO_CLIENTE = 5
Public Const MNEMONICO_MALADIRETA_TIPOOBJ_RECEBIMENTO = 6
Public Const MNEMONICO_MALADIRETA_TIPOOBJ_CONTRATO = 7
Public Const MNEMONICO_MALADIRETA_TIPOOBJ_FILIALCLIENTE = 8
Public Const MNEMONICO_MALADIRETA_TIPOOBJ_OV = 9
Public Const MNEMONICO_MALADIRETA_TIPOOBJ_TRIBUTACAODOC = 10
Public Const MNEMONICO_MALADIRETA_TIPOOBJ_INFOADIC = 11
Public Const MNEMONICO_MALADIRETA_TIPOOBJ_INFOADIC_COMPRA = 12
Public Const MNEMONICO_MALADIRETA_TIPOOBJ_INFOADIC_EXPORT = 13
Public Const MNEMONICO_MALADIRETA_TIPOOBJ_INFOADIC_ENDENT = 14
Public Const MNEMONICO_MALADIRETA_TIPOOBJ_INFOADIC_ENDRET = 15
Public Const MNEMONICO_MALADIRETA_TIPOOBJ_CONDPAGTO = 16

Public Const STRING_MNEMONICO_MALADIRETA_DESCRICAO = 255
Public Const STRING_MNEMONICO_MALADIRETA = 50
Public Const STRING_MNEMONICO_MALADIRETA_NOMECAMPOOBJ = 50

Public Const STRING_PROPOSTAETAPA_DESCRICAO = 250
Public Const STRING_PROPOSTAETAPA_OBSERVACAO = 500

Public Const STRING_CONTRATOETAPA_DESCRICAO = 250
Public Const STRING_CONTRATOETAPA_OBSERVACAO = 500

Public Const STRING_PROPOSTAPRJ_CODIGO = 20
Public Const STRING_CONTRATOPRJ_CODIGO = 20

Public Const STRING_PRJ_CODIGO = 20
Public Const STRING_PRJ_NOMEREDUZIDO = 20
Public Const STRING_PRJ_DESCRICAO = 255
Public Const STRING_PRJ_RESPONSAVEL = 50
Public Const STRING_PRJ_OBJETIVO = 255
Public Const STRING_PRJ_JUSTIFICATIVA = 255
Public Const STRING_PRJ_OBSERVACAO = 255
Public Const STRING_PRJ_REFERENCIA = 20

Public Const STRING_ETAPAPRJ_CODIGO = 20
Public Const STRING_ETAPAPRJ_NOMEREDUZIDO = 50

Public Const STRING_PRJ_ESC_DESCRICAO = 255
Public Const STRING_PRJ_ESC_EXPECTATIVA = 255
Public Const STRING_PRJ_ESC_FATORES = 255
Public Const STRING_PRJ_ESC_RESTRICOES = 255
Public Const STRING_PRJ_ESC_PREMISSAS = 255
Public Const STRING_PRJ_ESC_EXCLUSOES = 255

Public Const STRING_CAMPO_CUST_TEXTO = 255

Public Const STRING_TIPO_CAMPO_CUST_NOMETELA = 50
Public Const STRING_TIPO_CAMPO_CUST_NOMETABELA = 50

Public Const CAMPO_CUSTOMIZADO_TIPO_PROJETO = 1
Public Const CAMPO_CUSTOMIZADO_TIPO_ETAPA = 2
Public Const CAMPO_CUSTOMIZADO_TIPO_PROPOSTA = 3
Public Const CAMPO_CUSTOMIZADO_TIPO_CONTRATO = 4

Public Const PRJ_TIPO_PAGTO = 1
Public Const PRJ_TIPO_RECEB = 2

Public Const INDICE_INF_PREV = 1
Public Const INDICE_CALC_PREV = 2
Public Const INDICE_INF_REAL = 3
Public Const INDICE_CALC_REAL = 4

Public Const STRING_PRJ_CR_OBSERVACAO = 255
Public Const STRING_PRJ_CR_ITEM = 5

Public Const PRJ_CR_TIPO_NF = 1
Public Const PRJ_CR_TIPO_SAQUE = 2
Public Const PRJ_CR_TIPO_DEPOSITO = 3
Public Const PRJ_CR_TIPO_TITREC = 4
Public Const PRJ_CR_TIPO_TITPAG = 5
Public Const PRJ_CR_TIPO_OV = 6
Public Const PRJ_CR_TIPO_PV = 7
Public Const PRJ_CR_TIPO_NFPAG = 8
Public Const PRJ_CR_TIPO_OP = 9
Public Const PRJ_CR_TIPO_PRODENTRADA = 10
Public Const PRJ_CR_TIPO_REQPROD = 11
Public Const PRJ_CR_TIPO_ORCSRV = 12
Public Const PRJ_CR_TIPO_OVHIST = 13
Public Const PRJ_CR_TIPO_PC = 14

Public Const PRJ_TIPO_VALID_VLR_MAIOR = 0
Public Const PRJ_TIPO_VALID_VLR_MENOR_GRAVACAO = 1
Public Const PRJ_TIPO_VALID_VLR_MENOR_TELA = 2
Public Const PRJ_TIPO_VALID_VLR_MENOR_AMBOS = 3

Public Const CAMPO_CUSTOMIZADO_QTD_REPETICOES = 5

Public Const CAMPO_CUSTOMIZADO_VARIACAO_INDEX_DATA = 1000
Public Const CAMPO_CUSTOMIZADO_VARIACAO_INDEX_VALOR = 2000
Public Const CAMPO_CUSTOMIZADO_VARIACAO_INDEX_NUMERO = 3000
Public Const CAMPO_CUSTOMIZADO_VARIACAO_INDEX_TEXTO = 4000

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeProjetos
    lNumIntDoc As Long
    sCodigo As String
    iFilialEmpresa As Integer
    sNomeReduzido As String
    sDescricao As String
    dtDataCriacao As Date
    lCliente As Long
    iFilialCliente As Integer
    sResponsavel As String
    sObjetivo As String
    sJustificativa As String
    sObservacao As String
    dtDataInicio As Date
    dtDataFim As Date
    dtDataInicioReal As Date
    dtDataFimReal As Date
    dPercentualComplet As Double
    lNumIntDocEscopo As Long
    sSegmento As String
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typePRJEscopo
    lNumIntDoc As Long
    sDescricao As String
    sExpectativa As String
    sFatoresSucesso As String
    sRestricoes As String
    sPremissas As String
    sExclusoesEspecificas As String
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeCamposCustomizados
    lNumIntDoc As Long
    lNumIntDocOrigem As Long
    iTipoNumIntDocOrigem As Integer
    dtData(1 To CAMPO_CUSTOMIZADO_QTD_REPETICOES) As Date
    sTexto(1 To CAMPO_CUSTOMIZADO_QTD_REPETICOES) As String
    lNumero(1 To CAMPO_CUSTOMIZADO_QTD_REPETICOES) As Long
    dValor(1 To CAMPO_CUSTOMIZADO_QTD_REPETICOES) As Double
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeTiposCamposCust
    iCodigo As Integer
    sNomeTela As String
    sNomeTabela As String
    iDatasPreenchida As Integer
    iTextosPreenchidos As Integer
    iNumerosPreenchidos As Integer
    iValoresPreenchidos As Integer
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeProjetoInfo
    lNumIntDoc As Long
    lNumIntDocPRJ As Long
    lNumIntDocEtapa As Long
    iTipoOrigem As Integer
    lNumIntDocOrigem As Long
    sCodigoOP As String
    iFilialEmpresa As Integer
End Type

Type typePRJCR
    lNumIntDoc As Long
    lNumIntDocPRJInfo As Long
    lNumIntDocPRJ As Long
    lNumIntDocEtapa As Long
    dValor As Double
    dQuantidade As Double
    dPercentual As Double
    iTipoValor As Integer
    sObservacao As String
    sItem As String
    iCalcAuto As Integer
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typePRJEtapas
    lNumIntDoc As Long
    lNumIntDocPRJ As Long
    sCodigo As String
    sReferencia As String
    sNomeReduzido As String
    sDescricao As String
    lCliente As Long
    iFilialCliente As Integer
    sResponsavel As String
    sObjetivo As String
    sJustificativa As String
    sObservacao As String
    dtDataInicio As Date
    dtDataFim As Date
    lNumIntDocEtapaPaiOrg As Long
    dtDataInicioReal As Date
    dtDataFimReal As Date
    dPercentualComplet As Double
    lNumIntDocEscopo As Long
    iNivel As Integer
    iSeq As Integer
    iPosicao As Integer
    dtDataVistoria As Date
    dtValidadeVistoria As Date
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typePRJEtapaItensProd
    lNumIntDoc As Long
    lNumIntDocEtapaPRJ As Long
    iSeq As Integer
    sProduto As String
    sDescricao As String
    sVersao As String
    sUM As String
    dQuantidade As Double
    sObservacao As String
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typePRJEtapaMaquinas
    lNumIntDoc As Long
    lNumIntDocEtapa As Long
    iSeq As Integer
    lNumIntDocMaq As Long
    sDescricao As String
    iQuantidade As Integer
    dHoras As Double
    dCusto As Double
    iTipo As Integer
    dtData As Date
    sObservacao As String
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typePRJEtapaMO
    lNumIntDoc As Long
    lNumIntDocEtapa As Long
    iSeq As Integer
    iMaoDeObra As Integer
    sDescricao As String
    iQuantidade As Integer
    dHoras As Double
    dCusto As Double
    iTipo As Integer
    dtData As Date
    sObservacao As String
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typePRJEtapasPredecessoras
    lNumIntDoc As Long
    lNumIntDocEtapa As Long
    lNumIntDocEtapaPre As Long
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typePRJEtapaMateriais
    lNumIntDoc As Long
    lNumIntDocEtapa As Long
    iSeq As Integer
    sProduto As String
    sVersao As String
    sDescricao As String
    sUM As String
    dQuantidade As Double
    dCusto As Double
    iTipo As Integer
    dtData As Date
    sObservacao As String
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typePRJRecebPagto
    lNumIntDoc As Long
    lNumIntDocPRJ As Long
    iTipo As Integer
    lNumero As Long
    dValor As Double
    lCliForn As Long
    iFilial As Integer
    lNumIntDocProposta As Long
    lNumIntDocContrato As Long
    iIncluiCFF As Integer
    iFilialEmpresa As Integer
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typePRJRecebPagtoRegras
    lNumIntDoc As Long
    lNumIntDocRecebPagto As Long
    sRegra As String
    dPercentual As Double
    iCondPagto As Integer
    sObservacao As String
End Type

Type typeMnemonicoPRJ
    sMnemonico As String
    iTipo As Integer
    iNumParam As Integer
    iParam1 As Integer
    iParam2 As Integer
    iParam3 As Integer
    sNomeGrid As String
    sMnemonicoCombo As String
    sMnemonicoDesc As String
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typePRJPropostas
    lNumIntDoc As Long
    lNumIntDocPRJ As Long
    sCodigo As String
    dtData As Date
    sObservacao As String
    lCliente As Long
    iFilialCliente As Integer
    dValorTotal As Double
    dValorProdutos As Double
    dValorFrete As Double
    dValorDesconto As Double
    dValorSeguro As Double
    dValorOutrasDespesas As Double
    dCustoInformado As Double
    dCustoCalculado As Double
    iExibirProdutos As Integer
    iExibirPreco As Integer
    iExibirCustoCalc As Integer
    iExibirCustoInfo As Integer
    sNaturezaOp As String
    lNumIntDocContrato As Long
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typePRJPropostaItem
    lNumIntDoc As Long
    iFilialEmpresa As Integer
    iItem As Integer
    lNumIntDocProposta As Long
    sProduto As String
    sDescProd As String
    lNumIntDocEtapa As Long
    sDescEtapa As String
    dQuantidade As Double
    dPrecoUnitario As Double
    dPrecoTotal As Double
    sUM As String
    dValorDesconto As Double
    dtDataEntrega As Date
    sObservacao As String
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typePRJPropostaEtapa
    lNumIntDoc As Long
    lNumIntDocProposta As Long
    lNumIntDocEtapa As Long
    lNumIntDocEtapaItemProd As Long
    iSelecionado As Integer
    dCustoInformado As Double
    dPreco As Double
    iImprimir As Integer
    sObservacao As String
    sDescricao As String
End Type

Type typeTribComplPRJProp
    lNumIntDoc As Long
    iTipo As Integer
    sNaturezaOp As String
    iTipoTributacao As Integer
    iIPITipo As Integer
    dIPIBaseCalculo As Double
    dIPIPercRedBase As Double
    dIPIAliquota As Double
    dIPIValor As Double
    iICMSTipo As Integer
    dICMSBase As Double
    dICMSPercRedBase As Double
    dICMSAliquota As Double
    dICMSValor As Double
    dICMSSubstBase As Double
    dICMSSubstAliquota As Double
    dICMSSubstValor As Double
    dICMSCredito As Double
    dPISCredito As Double
    dCOFINSCredito As Double
    dIPICredito As Double
End Type

Type typeTributacaoPRJProp
    iFilialEmpresa As Integer
    sCodProposta As String
    iTaxacaoAutomatica As Integer
    iTipoTributacao As Integer
    iTipoTributacaoManual As Integer
    dIPIBase As Double
    iIPIBaseManual As Integer
    dIPIValor As Double
    iIPIValorManual As Integer
    dICMSBase As Double
    iICMSBaseManual As Integer
    dICMSValor As Double
    iICMSValorManual As Integer
    dICMSSubstBase As Double
    iICMSSubstBaseManual As Integer
    dICMSSubstValor As Double
    iICMSSubstValorManual As Integer
    iISSIncluso As Integer
    dISSBase As Double
    dISSAliquota As Double
    iISSAliquotaManual As Integer
    dISSValor As Double
    iISSValorManual As Integer
    dIRRFBase As Double
    dIRRFAliquota As Double
    iIRRFAliquotaManual As Integer
    dIRRFValor As Double
    iIRRFValorManual As Integer
    iPISRetidoManual As Integer
    iISSRetidoManual As Integer
    iCOFINSRetidoManual As Integer
    iCSLLRetidoManual As Integer
    dPISRetido As Double
    dISSRetido As Double
    dCOFINSRetido As Double
    dCSLLRetido As Double
    dValorINSS As Double
    iINSSValorManual As Integer
    iINSSRetido As Integer
    iINSSRetidoManual As Integer
    dINSSBase As Double
    iINSSBaseManual As Integer
    dINSSDeducoes As Double
    iINSSDeducoesManual As Integer
    dPISCredito As Double
    iPISCreditoManual As Integer
    dCOFINSCredito As Double
    iCOFINSCreditoManual As Integer
    dICMSCredito As Double
    iICMSCreditoManual As Integer
    dIPICredito As Double
    iIPICreditoManual As Integer
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typePRJContratos
    lNumIntDoc As Long
    lNumIntDocPRJ As Long
    lNumIntDocProposta As Long
    sCodigo As String
    dtData As Date
    sObservacao As String
    lCliente As Long
    iFilialCliente As Integer
    dValorTotal As Double
    dValorProdutos As Double
    dValorFrete As Double
    dValorDesconto As Double
    dValorSeguro As Double
    dValorOutrasDespesas As Double
    dCustoInformado As Double
    dCustoCalculado As Double
    iExibirProdutos As Integer
    iExibirPreco As Integer
    iExibirCustoCalc As Integer
    iExibirCustoInfo As Integer
    sNaturezaOp As String
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typePRJContratoItem
    lNumIntDoc As Long
    iFilialEmpresa As Integer
    iItem As Integer
    lNumIntDocContrato As Long
    sProduto As String
    sDescProd As String
    lNumIntDocEtapa As Long
    sDescEtapa As String
    dQuantidade As Double
    dPrecoUnitario As Double
    dPrecoTotal As Double
    sUM As String
    dValorDesconto As Double
    dtDataEntrega As Date
    sObservacao As String
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typePRJContratoEtapa
    lNumIntDoc As Long
    lNumIntDocContrato As Long
    lNumIntDocEtapa As Long
    lNumIntDocEtapaItemProd As Long
    iSelecionado As Integer
    dCustoInformado As Double
    dPreco As Double
    iImprimir As Integer
    sObservacao As String
    sDescricao As String
End Type

Type typeTribComplPRJCTR
    lNumIntDoc As Long
    iTipo As Integer
    sNaturezaOp As String
    iTipoTributacao As Integer
    iIPITipo As Integer
    dIPIBaseCalculo As Double
    dIPIPercRedBase As Double
    dIPIAliquota As Double
    dIPIValor As Double
    iICMSTipo As Integer
    dICMSBase As Double
    dICMSPercRedBase As Double
    dICMSAliquota As Double
    dICMSValor As Double
    dICMSSubstBase As Double
    dICMSSubstAliquota As Double
    dICMSSubstValor As Double
    dICMSCredito As Double
    dPISCredito As Double
    dCOFINSCredito As Double
    dIPICredito As Double
End Type

Type typeTributacaoPRJCTR
    iFilialEmpresa As Integer
    sCodContrato As String
    iTaxacaoAutomatica As Integer
    iTipoTributacao As Integer
    iTipoTributacaoManual As Integer
    dIPIBase As Double
    iIPIBaseManual As Integer
    dIPIValor As Double
    iIPIValorManual As Integer
    dICMSBase As Double
    iICMSBaseManual As Integer
    dICMSValor As Double
    iICMSValorManual As Integer
    dICMSSubstBase As Double
    iICMSSubstBaseManual As Integer
    dICMSSubstValor As Double
    iICMSSubstValorManual As Integer
    iISSIncluso As Integer
    dISSBase As Double
    dISSAliquota As Double
    iISSAliquotaManual As Integer
    dISSValor As Double
    iISSValorManual As Integer
    dIRRFBase As Double
    dIRRFAliquota As Double
    iIRRFAliquotaManual As Integer
    dIRRFValor As Double
    iIRRFValorManual As Integer
    iPISRetidoManual As Integer
    iISSRetidoManual As Integer
    iCOFINSRetidoManual As Integer
    iCSLLRetidoManual As Integer
    dPISRetido As Double
    dISSRetido As Double
    dCOFINSRetido As Double
    dCSLLRetido As Double
    dValorINSS As Double
    iINSSValorManual As Integer
    iINSSRetido As Integer
    iINSSRetidoManual As Integer
    dINSSBase As Double
    iINSSBaseManual As Integer
    dINSSDeducoes As Double
    iINSSDeducoesManual As Integer
    dPISCredito As Double
    iPISCreditoManual As Integer
    dCOFINSCredito As Double
    iCOFINSCreditoManual As Integer
    dICMSCredito As Double
    iICMSCreditoManual As Integer
    dIPICredito As Double
    iIPICreditoManual As Integer
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeMnemonicoMalaDireta
    sMnemonico As String
    sDescricao As String
    iTipoObj As Integer
    sNomeCampoObj As String
    iTipo As Integer
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeApontProdPRJ
    lNumIntDoc As Long
    lNumIntDocApont As Long
    iSeq As Integer
    sProduto As String
    sUM As String
    dQtd As Double
    dCusto As Double
    sObs As String
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeApontMOPRJ
    lNumIntDoc As Long
    lNumIntDocApont As Long
    iSeq As Integer
    iCodMO As Integer
    dHoras As Double
    iQtd As Integer
    dCusto As Double
    sObs As String
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeApontMaqPRJ
    lNumIntDoc As Long
    lNumIntDocApont As Long
    iSeq As Integer
    iCodMaq As Integer
    dHoras As Double
    iQtd As Integer
    dCusto As Double
    sObs As String
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeApontPRJ
    lNumIntDoc As Long
    lCodigo As Long
    lNumIntDocPRJ As Long
    lNumIntDocEtapa As Long
    dtData As Date
    sDescricao As String
    sObservacao As String
    sUsuario As String
    dtDataRegistro As Date
    dHoraRegistro As Double
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typePRJEtapaVistorias
    lNumIntDoc As Long
    lNumIntPRJEtapa As Long
    lCodigo As Long
    dtData As Date
    dtDataValidade As Date
    sResponsavel As String
    sLaudo As String
End Type

Public Function Inicializa_Mascara_Projeto(ByVal objControle As Object) As Long
'inicializa a mascara da Projeto

Dim sMascaraProjeto As String
Dim lErro As Long

On Error GoTo Erro_Inicializa_Mascara_Projeto

    If Controle_ObterNomeClasse(objControle) = "MaskEdBox" Then

        'Inicializa a máscara da Projeto
        sMascaraProjeto = String(STRING_PRJ_CODIGO, 0)
        
        'Armazena em sMascaraProjeto a mascara a ser a ser exibida no campo Projeto
        lErro = MascaraItem(SEGMENTO_PROJETO, sMascaraProjeto)
        If lErro <> SUCESSO Then gError 189038
        
        'coloca a mascara na tela.
        objControle.Mask = sMascaraProjeto
        
    End If
    
    Inicializa_Mascara_Projeto = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Mascara_Projeto:

    Inicializa_Mascara_Projeto = gErr
    
    Select Case gErr
    
        Case 189038
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARAITEM", gErr)
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 189039)
        
    End Select

    Exit Function

End Function

Public Function Inicializa_Mascara_RefEtapa(ByVal objControle As Object) As Long
'inicializa a mascara da RefEtapa

Dim sMascaraRefEtapa As String
Dim lErro As Long

On Error GoTo Erro_Inicializa_Mascara_RefEtapa

    'Inicializa a máscara da RefEtapa
    sMascaraRefEtapa = String(STRING_PRJ_REFERENCIA, 0)
    
    'Armazena em sMascaraRefEtapa a mascara a ser a ser exibida no campo RefEtapa
    lErro = MascaraItem(SEGMENTO_REFETAPA, sMascaraRefEtapa)
    If lErro <> SUCESSO Then gError 189040
    
    'coloca a mascara na tela.
    objControle.Mask = sMascaraRefEtapa
    
    Inicializa_Mascara_RefEtapa = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Mascara_RefEtapa:

    Inicializa_Mascara_RefEtapa = gErr
    
    Select Case gErr
    
        Case 189040
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARAITEM", gErr)
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 189041)
        
    End Select

    Exit Function

End Function

Public Function Retorno_Projeto_Tela(ByVal objControle As Object, ByVal sValor As String) As Long
'inicializa a mascara da Projeto

Dim sMascaraProjeto As String
Dim lErro As Long

On Error GoTo Erro_Retorno_Projeto_Tela

    If Len(Trim(sValor)) <> 0 Then
    
        sMascaraProjeto = String(STRING_PRJ_CODIGO, 0)
    
        lErro = Mascara_RetornaItemEnxuto(SEGMENTO_PROJETO, sValor, sMascaraProjeto)
        If lErro <> SUCESSO Then gError 189043
    
        objControle.PromptInclude = False
        objControle.Text = sMascaraProjeto
        objControle.PromptInclude = True
        
    Else
    
        objControle.PromptInclude = False
        objControle.Text = ""
        objControle.PromptInclude = True
        
    End If
    
    Retorno_Projeto_Tela = SUCESSO
    
    Exit Function
    
Erro_Retorno_Projeto_Tela:

    Retorno_Projeto_Tela = gErr
    
    Select Case gErr
    
        Case 189043
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARAITEM", gErr)
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 189044)
        
    End Select

    Exit Function

End Function

Public Function Retorno_Projeto_Tela2(ByVal sValor As String, sValorMasc As String) As Long
'inicializa a mascara da Projeto

Dim lErro As Long

On Error GoTo Erro_Retorno_Projeto_Tela2

    If Len(Trim(sValor)) <> 0 Then
    
        sValorMasc = String(STRING_PRJ_CODIGO, 0)
    
        lErro = Mascara_RetornaItemTela(SEGMENTO_PROJETO, sValor, sValorMasc)
        If lErro <> SUCESSO Then gError 189247
        
    Else
    
        sValorMasc = ""
        
    End If
    
    Retorno_Projeto_Tela2 = SUCESSO
    
    Exit Function
    
Erro_Retorno_Projeto_Tela2:

    Retorno_Projeto_Tela2 = gErr
    
    Select Case gErr
    
        Case 189247
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARAITEM", gErr)
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 189248)
        
    End Select

    Exit Function

End Function

Public Function Retorno_RefEtapa_Tela(ByVal objControle As Object, ByVal sValor As String) As Long
'inicializa a mascara da RefEtapa

Dim sMascaraRefEtapa As String
Dim lErro As Long

On Error GoTo Erro_Retorno_RefEtapa_Tela

    If Len(Trim(sValor)) <> 0 Then
    
        sMascaraRefEtapa = String(STRING_ETAPAPRJ_CODIGO, 0)
    
        lErro = Mascara_RetornaItemEnxuto(SEGMENTO_REFETAPA, sValor, sMascaraRefEtapa)
        If lErro <> SUCESSO Then gError 189045
    
        objControle.PromptInclude = False
        objControle.Text = sMascaraRefEtapa
        objControle.PromptInclude = True
        
    Else
    
        objControle.PromptInclude = False
        objControle.Text = ""
        objControle.PromptInclude = True
        
    End If
    
    Retorno_RefEtapa_Tela = SUCESSO
    
    Exit Function
    
Erro_Retorno_RefEtapa_Tela:

    Retorno_RefEtapa_Tela = gErr
    
    Select Case gErr
    
        Case 189045
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARAITEM", gErr)
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 189046)
        
    End Select

    Exit Function

End Function

Public Function Projeto_Formata(ByVal sValor As String, sValorFormatado, iCampoPreenchido As Integer) As Long
'inicializa a mascara da Projeto

Dim lErro As Long

On Error GoTo Erro_Projeto_Formata

    sValorFormatado = String(STRING_PRJ_CODIGO, 0)
    
    'Coloca no formato do BD
    lErro = CF("Item_Formata", SEGMENTO_PROJETO, sValor, sValorFormatado, iCampoPreenchido)
    If lErro <> SUCESSO Then gError 189047
    
    Projeto_Formata = SUCESSO
    
    Exit Function
    
Erro_Projeto_Formata:

    Projeto_Formata = gErr
    
    Select Case gErr
    
        Case 189047
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 189048)
        
    End Select

    Exit Function

End Function

Public Function RefEtapa_Formata(ByVal sValor As String, sValorFormatado, iCampoPreenchido As Integer) As Long
'inicializa a mascara da RefEtapa

Dim lErro As Long

On Error GoTo Erro_RefEtapa_Formata

    sValorFormatado = String(STRING_PRJ_CODIGO, 0)
    
    'Coloca no formato do BD
    lErro = CF("Item_Formata", SEGMENTO_REFETAPA, sValor, sValorFormatado, iCampoPreenchido)
    If lErro <> SUCESSO Then gError 189049
    
    RefEtapa_Formata = SUCESSO
    
    Exit Function
    
Erro_RefEtapa_Formata:

    RefEtapa_Formata = gErr
    
    Select Case gErr
    
        Case 189049
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 189050)
        
    End Select

    Exit Function

End Function


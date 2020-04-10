VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPRJPropostas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

Implements ClassTributoDoc 'Inserido por Leo em 03/05/02
'Polimorfismo que implementando a ClassTributoDoc
Private mvarobjTributacaoDoc As New ClassTributacaoDocAux

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDoc As Long
Private mvarlNumIntDocPRJ As Long
Private mvarsCodigo As String
Private mvardtData As Date
Private mvarsObservacao As String
Private mvarlCliente As Long
Private mvariFilialCliente As Integer
Private mvardValorTotal As Double
Private mvardValorProdutos As Double
Private mvardValorFrete As Double
Private mvardValorDesconto As Double
Private mvardValorSeguro As Double
Private mvardValorOutrasDespesas As Double
Private mvardCustoInformado As Double
Private mvardCustoCalculado As Double
Private mvarsNaturezaOp As String
Private mvarcolItens As New Collection
Private mvarobjTributacaoPRJProp As New ClassTributacaoDoc
Private mvariNaturezaOpManual As Integer
Private mvariFilialEmpresa As Integer
Private mvarobjTiposCamposCust As New ClassTiposCamposCust
Private mvarobjCamposCust As New ClassCamposCustomizados
Private mvarcolEtapas As New Collection
Private mvariExibirProdutos As Integer
Private mvariExibirPreco As Integer
Private mvariExibirCustoCalc As Integer
Private mvariExibirCustoInfo As Integer
Private mvarlNumIntDocContrato As Long
Private mvardValorDescontoItens As Double
Private mvardValorItens As Double

Public Property Let dValorItens(ByVal vData As Double)
    mvardValorItens = vData
End Property

Public Property Get dValorItens() As Double
    dValorItens = mvardValorItens
End Property

Public Property Let dValorDescontoItens(ByVal vData As Double)
    mvardValorDescontoItens = vData
End Property

Public Property Get dValorDescontoItens() As Double
    dValorDescontoItens = mvardValorDescontoItens
End Property

Public Property Let lNumIntDocContrato(ByVal vData As Long)
    mvarlNumIntDocContrato = vData
End Property

Public Property Get lNumIntDocContrato() As Long
    lNumIntDocContrato = mvarlNumIntDocContrato
End Property

Public Property Set objTributacaoPRJProp(ByVal vData As Object)
    Set mvarobjTributacaoPRJProp = vData
End Property

Public Property Get objTributacaoPRJProp() As ClassTributacaoDoc
    Set objTributacaoPRJProp = mvarobjTributacaoPRJProp
End Property

Public Property Let iExibirCustoInfo(ByVal vData As Integer)
    mvariExibirCustoInfo = vData
End Property

Public Property Get iExibirCustoInfo() As Integer
    iExibirCustoInfo = mvariExibirCustoInfo
End Property

Public Property Let iExibirCustoCalc(ByVal vData As Integer)
    mvariExibirCustoCalc = vData
End Property

Public Property Get iExibirCustoCalc() As Integer
    iExibirCustoCalc = mvariExibirCustoCalc
End Property

Public Property Let iExibirPreco(ByVal vData As Integer)
    mvariExibirPreco = vData
End Property

Public Property Get iExibirPreco() As Integer
    iExibirPreco = mvariExibirPreco
End Property

Public Property Let iExibirProdutos(ByVal vData As Integer)
    mvariExibirProdutos = vData
End Property

Public Property Get iExibirProdutos() As Integer
    iExibirProdutos = mvariExibirProdutos
End Property

Public Property Set colEtapas(ByVal vData As Collection)
    Set mvarcolEtapas = vData
End Property

Public Property Get colEtapas() As Collection
    Set colEtapas = mvarcolEtapas
End Property

Public Property Let objCamposCust(ByVal vData As ClassCamposCustomizados)
    Set mvarobjCamposCust = vData
End Property

Public Property Get objCamposCust() As ClassCamposCustomizados
    Set objCamposCust = mvarobjCamposCust
End Property

Public Property Let objTiposCamposCust(ByVal vData As ClassTiposCamposCust)
    Set mvarobjTiposCamposCust = vData
End Property

Public Property Get objTiposCamposCust() As ClassTiposCamposCust
    Set objTiposCamposCust = mvarobjTiposCamposCust
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Set colItens(ByVal vData As Collection)
    Set mvarcolItens = vData
End Property

Public Property Get colItens() As Collection
    Set colItens = mvarcolItens
End Property

Public Property Let sNaturezaOp(ByVal vData As String)
    mvarsNaturezaOp = vData
End Property

Public Property Get sNaturezaOp() As String
    sNaturezaOp = mvarsNaturezaOp
End Property

Public Property Let iNaturezaOpManual(ByVal vData As Integer)
    mvariNaturezaOpManual = vData
End Property

Public Property Get iNaturezaOpManual() As Integer
    iNaturezaOpManual = mvariNaturezaOpManual
End Property
Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let lNumIntDocPRJ(ByVal vData As Long)
    mvarlNumIntDocPRJ = vData
End Property

Public Property Get lNumIntDocPRJ() As Long
    lNumIntDocPRJ = mvarlNumIntDocPRJ
End Property

Public Property Let sCodigo(ByVal vData As String)
    mvarsCodigo = vData
End Property

Public Property Get sCodigo() As String
    sCodigo = mvarsCodigo
End Property

Public Property Let dtData(ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let sObservacao(ByVal vData As String)
    mvarsObservacao = vData
End Property

Public Property Get sObservacao() As String
    sObservacao = mvarsObservacao
End Property

Public Property Let lCliente(ByVal vData As Long)
    mvarlCliente = vData
End Property

Public Property Get lCliente() As Long
    lCliente = mvarlCliente
End Property

Public Property Let iFilialCliente(ByVal vData As Integer)
    mvariFilialCliente = vData
End Property

Public Property Get iFilialCliente() As Integer
    iFilialCliente = mvariFilialCliente
End Property

Public Property Let dValorTotal(ByVal vData As Double)
    mvardValorTotal = vData
End Property

Public Property Get dValorTotal() As Double
    dValorTotal = mvardValorTotal
End Property

Public Property Let dValorProdutos(ByVal vData As Double)
    mvardValorProdutos = vData
End Property

Public Property Get dValorProdutos() As Double
    dValorProdutos = mvardValorProdutos
End Property

Public Property Let dValorFrete(ByVal vData As Double)
    mvardValorFrete = vData
End Property

Public Property Get dValorFrete() As Double
    dValorFrete = mvardValorFrete
End Property

Public Property Let dValorDesconto(ByVal vData As Double)
    mvardValorDesconto = vData
End Property

Public Property Get dValorDesconto() As Double
    dValorDesconto = mvardValorDesconto
End Property

Public Property Let dValorSeguro(ByVal vData As Double)
    mvardValorSeguro = vData
End Property

Public Property Get dValorSeguro() As Double
    dValorSeguro = mvardValorSeguro
End Property

Public Property Let dValorOutrasDespesas(ByVal vData As Double)
    mvardValorOutrasDespesas = vData
End Property

Public Property Get dValorOutrasDespesas() As Double
    dValorOutrasDespesas = mvardValorOutrasDespesas
End Property

Public Property Let dCustoInformado(ByVal vData As Double)
    mvardCustoInformado = vData
End Property

Public Property Get dCustoInformado() As Double
    dCustoInformado = mvardCustoInformado
End Property

Public Property Let dCustoCalculado(ByVal vData As Double)
    mvardCustoCalculado = vData
End Property

Public Property Get dCustoCalculado() As Double
    dCustoCalculado = mvardCustoCalculado
End Property

Private Function ClassTributoDoc_Ativar() As Long

Dim lErro As Long
Dim objTributoDocItem As ClassTributoDocItem
Dim objPRJPropostaItem As ClassPRJPropostaItem

On Error GoTo Erro_ClassTributoDoc_Ativar

    lErro = gobjTributacao.Tributacao_Doc_Iniciar(Me)
    If lErro <> SUCESSO Then gError 185890
    
    Set mvarobjTributacaoDoc = New ClassTributacaoDocAux
    
    'ativar o ov em si
    If mvarobjTributacaoDoc.Ativar(Me) <> SUCESSO Then
        
        Set mvarobjTributacaoDoc = Nothing
        gError 185891
    
    End If
    
    'ativar os seus itens
    For Each objPRJPropostaItem In colItens

        Set objTributoDocItem = objPRJPropostaItem

        lErro = objTributoDocItem.Ativar(Me)
        If lErro <> SUCESSO Then gError 185892

    Next
        
    ClassTributoDoc_Ativar = SUCESSO

    Exit Function

Erro_ClassTributoDoc_Ativar:

    ClassTributoDoc_Ativar = gErr

    Select Case gErr

        Case 185890, 185891, 185892
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185893)

    End Select

    Exit Function

End Function

Private Function ClassTributoDoc_Desativar() As Long

Dim lErro As Long
Dim objTributoDocItem As ClassTributoDocItem
Dim objPRJPropostaItem As ClassPRJPropostaItem

On Error GoTo Erro_ClassTributoDoc_Desativar
    
    'desativar os seus itens
    For Each objPRJPropostaItem In colItens
        
        Set objTributoDocItem = objPRJPropostaItem
        
        lErro = objTributoDocItem.Desativar()
        If lErro <> SUCESSO Then gError 185894
    
    Next
    
    'desativar o ov em si
    If Not (mvarobjTributacaoDoc Is Nothing) Then
        
        lErro = mvarobjTributacaoDoc.Desativar()
        If lErro <> SUCESSO Then gError 185895
        
        Set mvarobjTributacaoDoc = Nothing
    
    End If
     
    lErro = gobjTributacao.Tributacao_Doc_Terminar(Me)
    If lErro <> SUCESSO Then gError 185896
    
    ClassTributoDoc_Desativar = SUCESSO

    Exit Function

Erro_ClassTributoDoc_Desativar:

    ClassTributoDoc_Desativar = gErr

    Select Case gErr

        Case 185894, 185895, 185896
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185897)

    End Select

    Exit Function

End Function

Private Function ClassTributoDoc_GuardarVarTela(sIdentificador As String, vValor As Variant) As Long

Dim lErro As Long
Dim objVariavelCalculo As New ClassVariavelCalculo

On Error GoTo Erro_ClassTributoDoc_GuardarVarTela

    objVariavelCalculo.sIdentificador = sIdentificador
    objVariavelCalculo.vValor = vValor
    objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL
    
    ClassTributoDoc_GuardarVarTela = ClassTributoDoc_GuardarVar(objVariavelCalculo)

    Exit Function

Erro_ClassTributoDoc_GuardarVarTela:

    ClassTributoDoc_GuardarVarTela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185898)

    End Select

    Exit Function

End Function

Private Property Get ClassTributoDoc_objTributacaoDoc() As GlobaisTRB.ClassTributacaoDocAux
    Set ClassTributoDoc_objTributacaoDoc = mvarobjTributacaoDoc
End Property

Private Function ClassTributoDoc_ObterColDocCompl(colDocCompl As Collection) As Long

'Dim objPRJPropTribCompl As ClassPVTribCompl
'
'    Set colDocCompl = New Collection
'
'    Set objPRJPropTribCompl = New ClassPVTribCompl
'    Call objPRJPropTribCompl.Iniciar(Me, mvarobjTributacaoPRJProp.objTributacaoFrete)
'    Call colDocCompl.Add(objPRJPropTribCompl)
'
'    Set objPRJPropTribCompl = New ClassPVTribCompl
'    Call objPRJPropTribCompl.Iniciar(Me, mvarobjTributacaoPRJProp.objTributacaoSeguro)
'    Call colDocCompl.Add(objPRJPropTribCompl)
'
'    Set objPRJPropTribCompl = New ClassPVTribCompl
'    Call objPRJPropTribCompl.Iniciar(Me, mvarobjTributacaoPRJProp.objTributacaoDesconto)
'    Call colDocCompl.Add(objPRJPropTribCompl)
'
'    Set objPRJPropTribCompl = New ClassPVTribCompl
'    Call objPRJPropTribCompl.Iniciar(Me, mvarobjTributacaoPRJProp.objTributacaoOutras)
'    Call colDocCompl.Add(objPRJPropTribCompl)
    
    ClassTributoDoc_ObterColDocCompl = SUCESSO
    
End Function

Private Function ClassTributoDoc_ObterColDocItens(colDocItens As Collection) As Long
'retorna colecao de itens que implementam interface como TaxDocItemExemplo

Dim objPRJPropostaItem As ClassPRJPropostaItem

   Set colDocItens = New Collection
   
   For Each objPRJPropostaItem In mvarcolItens
        Call colDocItens.Add(objPRJPropostaItem)
   Next
   
   ClassTributoDoc_ObterColDocItens = SUCESSO
   
End Function

Private Function ClassTributoDoc_ObterDocAux(objTributacaoDoc As ClassTributacaoDocAux) As Long
'retorna objeto auxiliar vinculado ao Doc
    
    Set objTributacaoDoc = mvarobjTributacaoDoc
    ClassTributoDoc_ObterDocAux = SUCESSO

End Function

Private Function ClassTributoDoc_ObterNatOpPadrao(sNatOp As String) As Long

Dim lErro As Long

On Error GoTo Erro_ClassTributoDoc_ObterNatOpPadrao

    lErro = mvarobjTributacaoDoc.ObterNaturezaOpPadrao(sNatOp)
    If lErro <> SUCESSO Then gError 185899
    
    ClassTributoDoc_ObterNatOpPadrao = SUCESSO

    Exit Function

Erro_ClassTributoDoc_ObterNatOpPadrao:

    ClassTributoDoc_ObterNatOpPadrao = gErr

    Select Case gErr

        Case 185899
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185900)

    End Select

    Exit Function

End Function

Private Function ClassTributoDoc_ObterTribPadrao(iTipoTrib As Integer) As Long

Dim lErro As Long

On Error GoTo Erro_ClassTributoDoc_ObterTribPadrao

    lErro = mvarobjTributacaoDoc.ObterTribPadrao(iTipoTrib)
    If lErro <> SUCESSO Then gError 185901
    
    ClassTributoDoc_ObterTribPadrao = SUCESSO

    Exit Function

Erro_ClassTributoDoc_ObterTribPadrao:

    ClassTributoDoc_ObterTribPadrao = gErr

    Select Case gErr

        Case 185901
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185902)

    End Select

    Exit Function

End Function

Private Function ClassTributoDoc_ObterVar(sIdentificador As String, objVariavelCalculo As ClassVariavelCalculo) As Long
'retorna o valor solicitado se este estiver "disponivel" em seu "contexto"
'objVariavelCalculo.sIdentificador tem que vir preenchida
Dim lErro As Long
On Error GoTo Erro_ClassTributoDoc_ObterVar
    
    If objVariavelCalculo Is Nothing Then
        Set objVariavelCalculo = New ClassVariavelCalculo
    End If
    
    objVariavelCalculo.sIdentificador = sIdentificador
    
    Select Case sIdentificador
        
        Case "SERIE"
            objVariavelCalculo.vValor = ""
            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO

        Case "DATA_EMISSAO"
            objVariavelCalculo.vValor = dtData
            If dtData <> DATA_NULA Then
                objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL
            Else
                objVariavelCalculo.iStatus = VAR_PREENCH_VAZIO
            End If
        
        Case "FILIAL_FATURAMENTO"
            objVariavelCalculo.vValor = iFilialEmpresa 'mvariFilialEmpresaFaturamento
            objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL
'
'        Case "NAT_OPERACAO"
'            objVariavelCalculo.vValor = mvarsNaturezaOp
'            objVariavelCalculo.iStatus = mvariNaturezaOpManual
'
'        Case "TIPO_TRIB"
'            objVariavelCalculo.vValor = mvarobjTributacaoPV.iTipoTributacao
'            objVariavelCalculo.iStatus = mvarobjTributacaoPV.iTipoTributacaoManual
        
        Case "TIPO_DOC_SIGLA"
            objVariavelCalculo.vValor = "" 'SIGLA_PV_NORMAL
            objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL
        
        Case "TIPO_DOC_CODIGO"
            objVariavelCalculo.vValor = CODIGO_PV_NORMAL
            objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL
        
        Case "CLIENTE_CODIGO"
            objVariavelCalculo.vValor = mvarlCliente
            objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL
            
        Case "FILIAL_CLIENTE_CODIGO"
            objVariavelCalculo.vValor = mvariFilialCliente 'a outra filial do pedido nao interessa p/efeito de tributacao
            objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL
        
        Case "FRETE_VALOR"
            objVariavelCalculo.vValor = mvardValorFrete
            objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL
        
        Case "SEGURO_VALOR"
            objVariavelCalculo.vValor = mvardValorSeguro
            objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL
        
        Case "DESP_ACESS_VALOR"
            objVariavelCalculo.vValor = mvardValorOutrasDespesas
            objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL
            
        Case "DESCONTO_VALOR"
            objVariavelCalculo.vValor = mvardValorDesconto
            objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL

'        'para ICMS
'        Case "ICMS_BASE"
'            objVariavelCalculo.vValor = mvarobjTributacaoPV.dICMSBase
'            objVariavelCalculo.iStatus = mvarobjTributacaoPV.iICMSBaseManual
'
'        Case "ICMS_VALOR"
'            objVariavelCalculo.vValor = mvarobjTributacaoPV.dICMSValor
'            objVariavelCalculo.iStatus = mvarobjTributacaoPV.iICMSValorManual
'
'        Case "ICMS_SUBST_BASE"
'            objVariavelCalculo.vValor = mvarobjTributacaoPV.dICMSSubstBase
'            objVariavelCalculo.iStatus = mvarobjTributacaoPV.iICMSSubstBaseManual
'
'        Case "ICMS_SUBST_VALOR"
'            objVariavelCalculo.vValor = mvarobjTributacaoPV.dICMSSubstValor
'            objVariavelCalculo.iStatus = mvarobjTributacaoPV.iICMSSubstValorManual
'
'        Case "ICMS_VALOR_CREDITO"
'            objVariavelCalculo.vValor = mvarobjTributacaoPV.dICMSCredito
'            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO
'
'        'para IPI
'        Case "IPI_BASE"
'            objVariavelCalculo.vValor = mvarobjTributacaoPV.dIPIBase
'            objVariavelCalculo.iStatus = mvarobjTributacaoPV.iIPIBaseManual
'
'        Case "IPI_VALOR"
'            objVariavelCalculo.vValor = mvarobjTributacaoPV.dIPIValor
'            objVariavelCalculo.iStatus = mvarobjTributacaoPV.iIPIValorManual
'
'        Case "IPI_VALOR_CREDITO"
'            objVariavelCalculo.vValor = mvarobjTributacaoPV.dIPICredito
'            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO
'
'        'para ISS
'        Case "ISS_INCLUSO"
'            objVariavelCalculo.vValor = mvarobjTributacaoPV.iISSIncluso
'            'objVariavelCalculo.iStatus = mvarobjTributacaoPV.iISSInclusoManual
'
'        Case "ISS_BASE"
'            objVariavelCalculo.vValor = mvarobjTributacaoPV.dISSBase
'            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO
'
'        Case "ISS_ALIQUOTA"
'            objVariavelCalculo.vValor = mvarobjTributacaoPV.dISSAliquota
'            objVariavelCalculo.iStatus = mvarobjTributacaoPV.iISSAliquotaManual
'
'        Case "ISS_VALOR"
'            objVariavelCalculo.vValor = mvarobjTributacaoPV.dISSValor
'            objVariavelCalculo.iStatus = mvarobjTributacaoPV.iISSValorManual
'
'        'para IR
'        Case "IR_BASE"
'            objVariavelCalculo.vValor = mvarobjTributacaoPV.dIRRFBase
'            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO
'
'        Case "IR_ALIQUOTA"
'            objVariavelCalculo.vValor = mvarobjTributacaoPV.dIRRFAliquota
'            objVariavelCalculo.iStatus = mvarobjTributacaoPV.iIRRFAliquotaManual
'
'        Case "IR_VALOR"
'            objVariavelCalculo.vValor = mvarobjTributacaoPV.dIRRFValor
'            objVariavelCalculo.iStatus = mvarobjTributacaoPV.iIRRFValorManual
'
'        Case "PIS_VALOR_CREDITO"
'            objVariavelCalculo.vValor = mvarobjTributacaoPV.dPISCredito
'            objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL
'
'        Case "COFINS_VALOR_CREDITO"
'            objVariavelCalculo.vValor = mvarobjTributacaoPV.dCOFINSCredito
'            objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL
'
'        Case "PIS_RETIDO"
'            objVariavelCalculo.vValor = mvarobjTributacaoPV.dPISRetido
'            objVariavelCalculo.iStatus = mvarobjTributacaoPV.iPISRetidoManual
'
'        Case "ISS_RETIDO"
'            objVariavelCalculo.vValor = mvarobjTributacaoPV.dISSRetido
'            objVariavelCalculo.iStatus = mvarobjTributacaoPV.iISSRetidoManual
'
'        Case "COFINS_RETIDO"
'            objVariavelCalculo.vValor = mvarobjTributacaoPV.dCOFINSRetido
'            objVariavelCalculo.iStatus = mvarobjTributacaoPV.iCOFINSRetidoManual
'
'        Case "CSLL_RETIDO"
'            objVariavelCalculo.vValor = mvarobjTributacaoPV.dCSLLRetido
'            objVariavelCalculo.iStatus = mvarobjTributacaoPV.iCSLLRetidoManual
            
        Case Else
'        Case "FILIAL_CLIENTE_OBJ"
'        Case "FILIAL_CLIENTE_INSC_EST"
'        Case "CLIENTE_ENDERECO_OBJ"
'        Case "CLIENTE_UF" '(EX p/exterior)
            
            lErro = objTributacao.ObterVar_Padrao(sIdentificador, objVariavelCalculo)
            If lErro <> SUCESSO Then
                'delegar p/tributacaoDoc
                lErro = mvarobjTributacaoDoc.ObterVar(sIdentificador, objVariavelCalculo)
                If lErro <> SUCESSO Then Error 27368
            End If
            
    End Select
    
    ClassTributoDoc_ObterVar = SUCESSO

    Exit Function

Erro_ClassTributoDoc_ObterVar:

    ClassTributoDoc_ObterVar = Err

    Select Case Err

        Case 27368
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152497)

    End Select

    Exit Function

End Function

Private Function ClassTributoDoc_GuardarVar(objVariavelCalculo As ClassVariavelCalculo) As Long
'armazena info de variavel guardando se esta foi alterada ou nao desde ...
Dim lErro As Long
On Error GoTo Erro_ClassTributoDoc_GuardarVar
        
    Select Case objVariavelCalculo.sIdentificador
        
'        Case "NAT_OPERACAO"
'            mvarsNaturezaOp = objVariavelCalculo.vValor
'            objVariavelCalculo.iStatus = mvariNaturezaOpManual
'
'        Case "TIPO_TRIB"
'            mvarobjTributacaoPV.iTipoTributacao = objVariavelCalculo.vValor
'            mvarobjTributacaoPV.iTipoTributacaoManual = objVariavelCalculo.iStatus
        
        Case "TIPO_DOC_SIGLA"
        Case "TIPO_DOC_CODIGO"
        
        Case "CLIENTE_CODIGO"
        
        Case "FILIAL_CLIENTE_OBJ"
        
        Case "FILIAL_CLIENTE_CODIGO"
        
        Case "FILIAL_CLIENTE_INSC_EST"

        Case "CLIENTE_ENDERECO_OBJ"
        
        Case "CLIENTE_UF" '(EX p/exterior)
        
        Case "FRETE_VALOR"
            mvardValorFrete = objVariavelCalculo.vValor
        
        Case "SEGURO_VALOR"
            mvardValorSeguro = objVariavelCalculo.vValor
        
        Case "DESP_ACESS_VALOR"
            mvardValorOutrasDespesas = objVariavelCalculo.vValor

'        'para ICMS
'        Case "ICMS_BASE"
'            mvarobjTributacaoPV.dICMSBase = objVariavelCalculo.vValor
'            mvarobjTributacaoPV.iICMSBaseManual = objVariavelCalculo.iStatus
'
'        Case "ICMS_VALOR"
'            mvarobjTributacaoPV.dICMSValor = objVariavelCalculo.vValor
'            mvarobjTributacaoPV.iICMSValorManual = objVariavelCalculo.iStatus
'
'        Case "ICMS_SUBST_BASE"
'            mvarobjTributacaoPV.dICMSSubstBase = objVariavelCalculo.vValor
'            mvarobjTributacaoPV.iICMSSubstBaseManual = objVariavelCalculo.iStatus
'
'        Case "ICMS_SUBST_VALOR"
'            mvarobjTributacaoPV.dICMSSubstValor = objVariavelCalculo.vValor
'            mvarobjTributacaoPV.iICMSSubstValorManual = objVariavelCalculo.iStatus
'
'        Case "ICMS_VALOR_CREDITO"
'             mvarobjTributacaoPV.dICMSCredito = objVariavelCalculo.vValor
'
'        'para IPI
'        Case "IPI_BASE"
'            mvarobjTributacaoPV.dIPIBase = objVariavelCalculo.vValor
'            mvarobjTributacaoPV.iIPIBaseManual = objVariavelCalculo.iStatus
'
'        Case "IPI_VALOR"
'            mvarobjTributacaoPV.dIPIValor = objVariavelCalculo.vValor
'            mvarobjTributacaoPV.iIPIValorManual = objVariavelCalculo.iStatus
'
'        Case "IPI_VALOR_CREDITO"
'            mvarobjTributacaoPV.dIPICredito = objVariavelCalculo.vValor
'
'        'para ISS
'        Case "ISS_INCLUSO"
'            mvarobjTributacaoPV.iISSIncluso = objVariavelCalculo.vValor
'            'mvarobjTributacaoPV.iISSInclusoManual = objVariavelCalculo.iStatus
'
'        Case "ISS_BASE"
'            mvarobjTributacaoPV.dISSBase = objVariavelCalculo.vValor
'
'        Case "ISS_ALIQUOTA"
'            mvarobjTributacaoPV.dISSAliquota = objVariavelCalculo.vValor
'            mvarobjTributacaoPV.iISSAliquotaManual = objVariavelCalculo.iStatus
'
'        Case "ISS_VALOR"
'            mvarobjTributacaoPV.dISSValor = objVariavelCalculo.vValor
'            mvarobjTributacaoPV.iISSValorManual = objVariavelCalculo.iStatus
'
'        'para IR
'        Case "IR_BASE"
'            mvarobjTributacaoPV.dIRRFBase = objVariavelCalculo.vValor
'
'        Case "IR_ALIQUOTA"
'            mvarobjTributacaoPV.dIRRFAliquota = objVariavelCalculo.vValor
'            mvarobjTributacaoPV.iIRRFAliquotaManual = objVariavelCalculo.iStatus
'
'        Case "IR_VALOR"
'            mvarobjTributacaoPV.dIRRFValor = objVariavelCalculo.vValor
'            mvarobjTributacaoPV.iIRRFValorManual = objVariavelCalculo.iStatus
'
'        Case "PIS_VALOR_CREDITO"
'            mvarobjTributacaoPV.dPISRetido = objVariavelCalculo.vValor
'
'        Case "COFINS_VALOR_CREDITO"
'            mvarobjTributacaoPV.dCOFINSRetido = objVariavelCalculo.vValor
'
'        Case "PIS_RETIDO"
'            mvarobjTributacaoPV.dPISRetido = objVariavelCalculo.vValor
'            mvarobjTributacaoPV.iPISRetidoManual = objVariavelCalculo.iStatus
'
'        Case "ISS_RETIDO"
'            mvarobjTributacaoPV.dISSRetido = objVariavelCalculo.vValor
'            mvarobjTributacaoPV.iISSRetidoManual = objVariavelCalculo.iStatus
'
'        Case "COFINS_RETIDO"
'            mvarobjTributacaoPV.dCOFINSRetido = objVariavelCalculo.vValor
'            mvarobjTributacaoPV.iCOFINSRetidoManual = objVariavelCalculo.iStatus
'
'        Case "CSLL_RETIDO"
'            mvarobjTributacaoPV.dCSLLRetido = objVariavelCalculo.vValor
'            mvarobjTributacaoPV.iCSLLRetidoManual = objVariavelCalculo.iStatus
        
        Case Else
            lErro = objTributacao.GuardarVar_Padrao(objVariavelCalculo)
            If lErro <> SUCESSO Then
                'delegar p/tributacaoDoc
                lErro = mvarobjTributacaoDoc.GuardarVar(objVariavelCalculo)
                If lErro <> SUCESSO Then Error 27369
            End If
            
    End Select

    ClassTributoDoc_GuardarVar = SUCESSO

    Exit Function

Erro_ClassTributoDoc_GuardarVar:

    ClassTributoDoc_GuardarVar = Err

    Select Case Err

        Case 27369
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152498)

    End Select

    Exit Function

End Function

Public Function RemoverItem(vntIndexKey As Variant) As Long
'exclui um item do OV identificando-o como na colecao de itens

Dim lErro As Long
Dim objTributoDocItem As ClassTributoDocItem
Dim objPropostaItem As ClassPRJPropostaItem

On Error GoTo Erro_RemoverItem

    Set objPropostaItem = colItens.Item(vntIndexKey)
    
    Set objTributoDocItem = objPropostaItem
    
    lErro = objTributoDocItem.Desativar()
    If lErro <> SUCESSO Then gError 103023
        
    Call colItens.Remove(vntIndexKey)
    
    RemoverItem = SUCESSO

    Exit Function

Erro_RemoverItem:

    RemoverItem = gErr

    Select Case gErr

        Case 103023
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152222)

    End Select

    Exit Function

End Function

Public Property Set objTributacao(ByVal vData As Object)
    Set mvarobjTributacaoPRJProp = vData
End Property

Public Property Get objTributacao() As ClassTributacaoDoc
    Set objTributacao = mvarobjTributacaoPRJProp
End Property

Public Property Let iFilial(ByVal vData As Integer)
    mvariFilialCliente = vData
End Property

Public Property Get iFilial() As Integer
    iFilial = mvariFilialCliente
End Property

Public Property Let dtDataEmissao(ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtDataEmissao() As Date
    dtDataEmissao = mvardtData
End Property

Private Function ClassTributoDoc_Alterado() As Long
    ClassTributoDoc_Alterado = mvarobjTributacaoDoc.Alterado
End Function

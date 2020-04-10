VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPRJPropostaItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"colReserva"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Implements ClassTributoDocItem

Private mvarobjTributacaoDocItem As ClassTributacaoDocItemAux
Private mvarobjDoc As ClassTributoDoc 'link p/Doc
Private mvarobjTributacaoPRJPropItem As New ClassTributacaoDocItem

Private mvardQuantidade As Double
Private mvardPrecoUnitario As Double
Private mvardPrecoTotal As Double
Private mvariFilialEmpresa As Integer
Private mvariItem As Integer
Private mvarlNumIntDocProposta As Long
Private mvarlNumIntDocEtapa As Long
Private mvarlNumIntDoc As Long
Private mvarsDescProd As String
Private mvarsUM As String
Private mvarsDescEtapa As String
Private mvarsCodEtapa As String
Private mvarsProduto As String
Private mvardtDataEntrega As Date
Private mvardValorDesconto As Double
Private mvardValorDescGlobal As Double
Private mvardValorDespGlobal As Double
Private mvarsObservacao As String

Public Property Let sObservacao(ByVal vData As String)
    mvarsObservacao = vData
End Property

Public Property Get sObservacao() As String
    sObservacao = mvarsObservacao
End Property

Public Property Let dValorDescGlobal(ByVal vData As Double)
    mvardValorDescGlobal = vData
End Property

Public Property Get dValorDescGlobal() As Double
    dValorDescGlobal = mvardValorDescGlobal
End Property

Public Property Let dValorDespGlobal(ByVal vData As Double)
    mvardValorDespGlobal = vData
End Property

Public Property Get dValorDespGlobal() As Double
    dValorDespGlobal = mvardValorDespGlobal
End Property

Public Property Let dValorDesconto(ByVal vData As Double)
    mvardValorDesconto = vData
End Property

Public Property Get dValorDesconto() As Double
    dValorDesconto = mvardValorDesconto
End Property

Public Property Let sCodEtapa(ByVal vData As String)
    mvarsCodEtapa = vData
End Property

Public Property Get sCodEtapa() As String
    sCodEtapa = mvarsCodEtapa
End Property

Public Property Let dtDataEntrega(ByVal vData As Date)
    mvardtDataEntrega = vData
End Property

Public Property Get dtDataEntrega() As Date
    dtDataEntrega = mvardtDataEntrega
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let sDescEtapa(ByVal vData As String)
    mvarsDescEtapa = vData
End Property

Public Property Get sDescEtapa() As String
    sDescEtapa = mvarsDescEtapa
End Property

Public Property Let sUM(ByVal vData As String)
    mvarsUM = vData
End Property

Public Property Get sUM() As String
    sUM = mvarsUM
End Property

Public Property Let sDescProd(ByVal vData As String)
    mvarsDescProd = vData
End Property

Public Property Get sDescProd() As String
    sDescProd = mvarsDescProd
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let lNumIntDocProposta(ByVal vData As Long)
    mvarlNumIntDocProposta = vData
End Property

Public Property Get lNumIntDocProposta() As Long
    lNumIntDocProposta = mvarlNumIntDocProposta
End Property

Public Property Let iItem(ByVal vData As Integer)
    mvariItem = vData
End Property

Public Property Get iItem() As Integer
    iItem = mvariItem
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let dPrecoTotal(ByVal vData As Double)
    mvardPrecoTotal = vData
End Property

Public Property Get dPrecoTotal() As Double
    dPrecoTotal = mvardPrecoTotal
End Property

Public Property Let dPrecoUnitario(ByVal vData As Double)
    mvardPrecoUnitario = vData
End Property

Public Property Get dPrecoUnitario() As Double
    dPrecoUnitario = mvardPrecoUnitario
End Property

Public Property Let dQuantidade(ByVal vData As Double)
    mvardQuantidade = vData
End Property

Public Property Get dQuantidade() As Double
    dQuantidade = mvardQuantidade
End Property

Public Property Let lNumIntDocEtapa(ByVal vData As Long)
    mvarlNumIntDocEtapa = vData
End Property

Public Property Get lNumIntDocEtapa() As Long
    lNumIntDocEtapa = mvarlNumIntDocEtapa
End Property

Private Function ClassTributoDocItem_Ativar(objDoc As ClassTributoDoc) As Long
Dim lErro As Long
On Error GoTo Erro_ClassTributoDocItem_Ativar

    Set mvarobjDoc = objDoc
    
    Set mvarobjTributacaoDocItem = New ClassTributacaoDocItemAux

    If mvarobjTributacaoDocItem.Ativar(Me) <> SUCESSO Then
        Set mvarobjTributacaoDocItem = Nothing
        gError 185880
    End If

    ClassTributoDocItem_Ativar = SUCESSO

    Exit Function

Erro_ClassTributoDocItem_Ativar:

    ClassTributoDocItem_Ativar = gErr

    Select Case gErr

        Case 185880

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185884)

    End Select

    Exit Function

End Function

Private Function ClassTributoDocItem_Desativar() As Long
Dim lErro As Long
On Error GoTo Erro_ClassTributoDocItem_Desativar

    If Not (mvarobjTributacaoDocItem Is Nothing) Then
        lErro = mvarobjTributacaoDocItem.Desativar()
        Set mvarobjTributacaoDocItem = Nothing
    End If

    ClassTributoDocItem_Desativar = SUCESSO

    Exit Function

Erro_ClassTributoDocItem_Desativar:

    ClassTributoDocItem_Desativar = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185883)

    End Select

    Exit Function

End Function

Private Function ClassTributoDocItem_GuardarVarTela(sIdentificador As String, vValor As Variant) As Long

Dim lErro As Long, objVariavelCalculo As New ClassVariavelCalculo

On Error GoTo Erro_ClassTributoDocItem_GuardarVarTela

    objVariavelCalculo.sIdentificador = sIdentificador
    objVariavelCalculo.vValor = vValor
    objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL
    
    ClassTributoDocItem_GuardarVarTela = ClassTributoDocItem_GuardarVar(objVariavelCalculo)

    Exit Function

Erro_ClassTributoDocItem_GuardarVarTela:

    ClassTributoDocItem_GuardarVarTela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185882)

    End Select

    Exit Function

End Function

Private Property Get ClassTributoDocItem_objTributacaoDocItem() As ClassTributacaoDocItemAux
    Set ClassTributoDocItem_objTributacaoDocItem = mvarobjTributacaoDocItem
End Property


Private Function ClassTributoDocItem_ObterVar(sIdentificador As String, objVariavelCalculo As ClassVariavelCalculo) As Long
'retorna o valor solicitado se este estiver "disponivel" em seu "contexto"
'objVariavelCalculo.sIdentificador tem que vir preenchida
Dim lErro As Long, objProduto As New ClassProduto
On Error GoTo Erro_ClassTributoDocItem_ObterVar

    If objVariavelCalculo Is Nothing Then
        Set objVariavelCalculo = New ClassVariavelCalculo
    End If
    
    objVariavelCalculo.sIdentificador = sIdentificador
    
    Select Case sIdentificador

'        Case "NAT_OPERACAO"
'            objVariavelCalculo.vValor = mvarobjTributacaoItemPV.sNaturezaOp
'            objVariavelCalculo.iStatus = mvarobjTributacaoItemPV.iNaturezaOpManual
'
'        Case "TIPO_TRIB"
'            objVariavelCalculo.vValor = mvarobjTributacaoItemPV.iTipoTributacao
'            objVariavelCalculo.iStatus = mvarobjTributacaoItemPV.iTipoTributacaoManual
'
'        Case "PRODUTO_CODIGO"
'            objVariavelCalculo.vValor = mvarsProduto
'            objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL
'
        Case "PRODUTO_VALOR"
            objVariavelCalculo.vValor = mvardPrecoTotal
            objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL

        Case "PRODUTO_DESC_GLOBAL" 'valor do desconto global rateado para o item
            objVariavelCalculo.vValor = mvardValorDescGlobal
            objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL

        Case "PRODUTO_DESP_GLOBAL" 'valor do (frete+seguro+outras desp) rateado para o item
            objVariavelCalculo.vValor = mvardValorDespGlobal
            objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL

        Case "PRODUTO_VALOR_BRUTO"
            objVariavelCalculo.vValor = Round(mvardPrecoTotal + mvardValorDesconto, 2)
            objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL

        Case "PRODUTO_QTDE"
            objVariavelCalculo.vValor = mvardQuantidade
            objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL
'
'        Case "PRODUTO_OBJ"
'
'            objProduto.sCodigo = mvarsProduto
'            lErro = CF("Produto_Le", objProduto)
'            If lErro <> SUCESSO Then Error 27598
'
'            Set objVariavelCalculo.vValor = objProduto
'            objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL
'
'        'para ICMS
'        Case "ICMS_TIPO"
'            objVariavelCalculo.vValor = mvarobjTributacaoItemPV.iICMSTipo
'            objVariavelCalculo.iStatus = mvarobjTributacaoItemPV.iICMSTipoManual
'
'        Case "ICMS_BASE"
'            objVariavelCalculo.vValor = mvarobjTributacaoItemPV.dICMSBase
'            objVariavelCalculo.iStatus = mvarobjTributacaoItemPV.iICMSBaseManual
'
'        Case "ICMS_RED_BASE"
'            objVariavelCalculo.vValor = mvarobjTributacaoItemPV.dICMSPercRedBase
'            objVariavelCalculo.iStatus = mvarobjTributacaoItemPV.iICMSPercRedBaseManual
'
'        Case "ICMS_ALIQUOTA"
'            objVariavelCalculo.vValor = mvarobjTributacaoItemPV.dICMSAliquota
'            objVariavelCalculo.iStatus = mvarobjTributacaoItemPV.iICMSAliquotaManual
'
'        Case "ICMS_VALOR"
'            objVariavelCalculo.vValor = mvarobjTributacaoItemPV.dICMSValor
'            objVariavelCalculo.iStatus = mvarobjTributacaoItemPV.iICMSValorManual
'
'        Case "ICMS_BASE_SUBST"
'            objVariavelCalculo.vValor = mvarobjTributacaoItemPV.dICMSSubstBase
'            objVariavelCalculo.iStatus = mvarobjTributacaoItemPV.iICMSSubstBaseManual
'
'        Case "ICMS_RED_BASE_SUBST"
'            objVariavelCalculo.vValor = mvarobjTributacaoItemPV.dICMSSubstPercRedBase
'            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO
'
'        Case "ICMS_ALIQUOTA_SUBST"
'            objVariavelCalculo.vValor = mvarobjTributacaoItemPV.dICMSSubstAliquota
'            objVariavelCalculo.iStatus = mvarobjTributacaoItemPV.iICMSSubstAliquotaManual
'
'        Case "ICMS_VALOR_SUBST"
'            objVariavelCalculo.vValor = mvarobjTributacaoItemPV.dICMSSubstValor
'            objVariavelCalculo.iStatus = mvarobjTributacaoItemPV.iICMSSubstValorManual
'
'        Case "ICMS_MARG_SUBST"
'            objVariavelCalculo.vValor = mvarobjTributacaoItemPV.dICMSSubstPercMVA
'            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO
'
'        Case "ICMS_VALOR_CREDITO"
'            objVariavelCalculo.vValor = mvarobjTributacaoItemPV.dICMSCredito
'
'        'para IPI
'        Case "IPI_TIPO"
'            objVariavelCalculo.vValor = mvarobjTributacaoItemPV.iIPITipo
'            objVariavelCalculo.iStatus = mvarobjTributacaoItemPV.iIPITipoManual
'
'        Case "IPI_BASE"
'            objVariavelCalculo.vValor = mvarobjTributacaoItemPV.dIPIBaseCalculo
'            objVariavelCalculo.iStatus = mvarobjTributacaoItemPV.iIPIBaseCalculoManual
'
'        Case "IPI_RED_BASE"
'            objVariavelCalculo.vValor = mvarobjTributacaoItemPV.dIPIPercRedBase
'            objVariavelCalculo.iStatus = mvarobjTributacaoItemPV.iIPIPercRedBaseManual
'
'        Case "IPI_ALIQUOTA"
'            objVariavelCalculo.vValor = mvarobjTributacaoItemPV.dIPIAliquota
'            objVariavelCalculo.iStatus = mvarobjTributacaoItemPV.iIPIAliquotaManual
'
'        Case "IPI_VALOR"
'            objVariavelCalculo.vValor = mvarobjTributacaoItemPV.dIPIValor
'            objVariavelCalculo.iStatus = mvarobjTributacaoItemPV.iIPIValorManual
'
'        Case "IPI_VALOR_CREDITO"
'            objVariavelCalculo.vValor = mvarobjTributacaoItemPV.dIPICredito
'
'        Case "PIS_VALOR_CREDITO"
'            objVariavelCalculo.vValor = mvarobjTributacaoItemPV.dPISCredito
'
'        Case "COFINS_VALOR_CREDITO"
'            objVariavelCalculo.vValor = mvarobjTributacaoItemPV.dCOFINSCredito
        
        Case Else
            'delegar p/tributacaoDocItem
            '"ICMS_MARG_SUBST"
            lErro = objTributacao.ObterVar_Padrao(sIdentificador, objVariavelCalculo)
            If lErro <> SUCESSO Then
            
                lErro = mvarobjTributacaoDocItem.ObterVar(sIdentificador, objVariavelCalculo)
                If lErro <> SUCESSO Then Error 27408
                
            End If

    End Select

    ClassTributoDocItem_ObterVar = SUCESSO

    Exit Function

Erro_ClassTributoDocItem_ObterVar:

    ClassTributoDocItem_ObterVar = Err

    Select Case Err

        Case 27408, 27598
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150905)

    End Select

    Exit Function

End Function

Private Function ClassTributoDocItem_GuardarVar(objVariavelCalculo As ClassVariavelCalculo) As Long
'armazena info de variavel guardando se esta foi alterada ou nao desde ...
Dim lErro As Long
On Error GoTo Erro_ClassTributoDocItem_GuardarVar

    Select Case objVariavelCalculo.sIdentificador

'        Case "NAT_OPERACAO"
'            mvarobjTributacaoItemPV.sNaturezaOp = objVariavelCalculo.vValor
'            mvarobjTributacaoItemPV.iNaturezaOpManual = objVariavelCalculo.iStatus
'
'        Case "TIPO_TRIB"
'            mvarobjTributacaoItemPV.iTipoTributacao = objVariavelCalculo.vValor
'            mvarobjTributacaoItemPV.iTipoTributacaoManual = objVariavelCalculo.iStatus
'
'        Case "PRODUTO_CODIGO"
        Case "PRODUTO_VALOR"
        
        Case "PRODUTO_VALOR_BRUTO"
        
        Case "PRODUTO_QTDE"

        Case "PRODUTO_DESC_GLOBAL" 'valor do desconto global rateado para o item
            mvardValorDescGlobal = objVariavelCalculo.vValor

        Case "PRODUTO_DESP_GLOBAL" 'valor do (frete+seguro+outras desp) rateado para o item
            mvardValorDespGlobal = objVariavelCalculo.vValor
'
'        'para ICMS
'        Case "ICMS_TIPO"
'            mvarobjTributacaoItemPV.iICMSTipo = objVariavelCalculo.vValor
'            mvarobjTributacaoItemPV.iICMSTipoManual = objVariavelCalculo.iStatus
'
'        Case "ICMS_BASE"
'            mvarobjTributacaoItemPV.dICMSBase = objVariavelCalculo.vValor
'            mvarobjTributacaoItemPV.iICMSBaseManual = objVariavelCalculo.iStatus
'
'        Case "ICMS_RED_BASE"
'            mvarobjTributacaoItemPV.dICMSPercRedBase = objVariavelCalculo.vValor
'            mvarobjTributacaoItemPV.iICMSPercRedBaseManual = objVariavelCalculo.iStatus
'
'        Case "ICMS_ALIQUOTA"
'            mvarobjTributacaoItemPV.dICMSAliquota = objVariavelCalculo.vValor
'            mvarobjTributacaoItemPV.iICMSAliquotaManual = objVariavelCalculo.iStatus
'
'        Case "ICMS_VALOR"
'            mvarobjTributacaoItemPV.dICMSValor = objVariavelCalculo.vValor
'            mvarobjTributacaoItemPV.iICMSValorManual = objVariavelCalculo.iStatus
'
'        Case "ICMS_BASE_SUBST"
'            mvarobjTributacaoItemPV.dICMSSubstBase = objVariavelCalculo.vValor
'            mvarobjTributacaoItemPV.iICMSSubstBaseManual = objVariavelCalculo.iStatus
'
'        Case "ICMS_RED_BASE_SUBST"
'            mvarobjTributacaoItemPV.dICMSSubstPercRedBase = objVariavelCalculo.vValor
'
'        Case "ICMS_ALIQUOTA_SUBST"
'            mvarobjTributacaoItemPV.dICMSSubstAliquota = objVariavelCalculo.vValor
'            mvarobjTributacaoItemPV.iICMSSubstAliquotaManual = objVariavelCalculo.iStatus
'
'        Case "ICMS_VALOR_SUBST"
'            mvarobjTributacaoItemPV.dICMSSubstValor = objVariavelCalculo.vValor
'            mvarobjTributacaoItemPV.iICMSSubstValorManual = objVariavelCalculo.iStatus
'
'        Case "ICMS_MARG_SUBST"
'            mvarobjTributacaoItemPV.dICMSSubstPercMVA = objVariavelCalculo.vValor
'
'        Case "ICMS_VALOR_CREDITO"
'            mvarobjTributacaoItemPV.dICMSCredito = objVariavelCalculo.vValor
'
'        'para IPI
'        Case "IPI_TIPO"
'            mvarobjTributacaoItemPV.iIPITipo = objVariavelCalculo.vValor
'            mvarobjTributacaoItemPV.iIPITipoManual = objVariavelCalculo.iStatus
'
'        Case "IPI_BASE"
'            mvarobjTributacaoItemPV.dIPIBaseCalculo = objVariavelCalculo.vValor
'            mvarobjTributacaoItemPV.iIPIBaseCalculoManual = objVariavelCalculo.iStatus
'
'        Case "IPI_RED_BASE"
'            mvarobjTributacaoItemPV.dIPIPercRedBase = objVariavelCalculo.vValor
'            mvarobjTributacaoItemPV.iIPIPercRedBaseManual = objVariavelCalculo.iStatus
'
'        Case "IPI_ALIQUOTA"
'            mvarobjTributacaoItemPV.dIPIAliquota = objVariavelCalculo.vValor
'            mvarobjTributacaoItemPV.iIPIAliquotaManual = objVariavelCalculo.iStatus
'
'        Case "IPI_VALOR"
'            mvarobjTributacaoItemPV.dIPIValor = objVariavelCalculo.vValor
'            mvarobjTributacaoItemPV.iIPIValorManual = objVariavelCalculo.iStatus
'
'        Case "IPI_VALOR_CREDITO"
'            mvarobjTributacaoItemPV.dIPICredito = objVariavelCalculo.vValor
'
'        Case "PIS_VALOR_CREDITO"
'            mvarobjTributacaoItemPV.dPISCredito = objVariavelCalculo.vValor
'
'        Case "COFINS_VALOR_CREDITO"
'            mvarobjTributacaoItemPV.dCOFINSCredito = objVariavelCalculo.vValor
'
        Case Else
                   
            lErro = objTributacao.GuardarVar_Padrao(objVariavelCalculo)
            If lErro <> SUCESSO Then
            
                lErro = mvarobjTributacaoDocItem.GuardarVar(objVariavelCalculo)
                If lErro <> SUCESSO Then Error 27407
                
            End If

    End Select

    ClassTributoDocItem_GuardarVar = SUCESSO

    Exit Function

Erro_ClassTributoDocItem_GuardarVar:

    ClassTributoDocItem_GuardarVar = Err

    Select Case Err

        Case 27407
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150906)

    End Select

    Exit Function

End Function

Private Property Get ClassTributoDocItem_objDoc() As ClassTributoDoc
    Set ClassTributoDocItem_objDoc = mvarobjDoc
End Property

Public Property Set objTributacaoPRJPropItem(ByVal vData As Object)
    Set mvarobjTributacaoPRJPropItem = vData
End Property

Public Property Get objTributacaoPRJPropItem() As ClassTributacaoDocItem
    Set objTributacaoPRJPropItem = mvarobjTributacaoPRJPropItem
End Property

Public Property Set objTributacao(ByVal vData As Object)
    Set mvarobjTributacaoPRJPropItem = vData
End Property

Public Property Get objTributacao() As ClassTributacaoDocItem
    Set objTributacao = mvarobjTributacaoPRJPropItem
End Property

Public Property Let sUnidadeMed(ByVal vData As String)
    mvarsUM = vData
End Property

Public Property Get sUnidadeMed() As String
    sUnidadeMed = mvarsUM
End Property

Public Property Let sProdutoDescricao(ByVal vData As String)
    mvarsDescProd = vData
End Property

Public Property Get sProdutoDescricao() As String
    sProdutoDescricao = mvarsDescProd
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescProd = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescProd
End Property

Private Function ClassTributoDocItem_Alterado() As Long
    ClassTributoDocItem_Alterado = mvarobjTributacaoDocItem.Alterado
End Function

Private Function ClassTributoDocItem_Calculado() As Long
    ClassTributoDocItem_Calculado = mvarobjTributacaoDocItem.Calculado
End Function

Private Function ClassTributoDocItem_iJaCalculado() As Integer
    ClassTributoDocItem_iJaCalculado = mvarobjTributacaoDocItem.iJaCalculado
End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassItemOV"
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
Private mvarobjTributacaoItemOV As New ClassTributacaoDocItem
Private mvarobjInfoAdicDocItem As New ClassInfoAdicDocItem

Private mvardQuantidade As Double 'local copy
Private mvardPrecoUnitario As Double 'local copy
Private mvardPrecoTotal As Double 'local copy
Private mvardValorDesconto As Double 'local copy
Private mvardValorDescGlobal As Double
Private mvardValorDespGlobal As Double
Private mvariFilialEmpresa As Integer 'local copy
Private mvariItem As Integer 'local copy
Private mvariClasseUM As Integer 'local copy
Private mvariStatus As Integer 'local copy
Private mvarlCodOrcamento As Long 'local copy
Private mvarlNumIntDoc As Long 'local copy
Private mvarsProdutoDescricao As String 'local copy
Private mvarsUnidadeMed As String 'local copy
Private mvarsDescricao As String 'local copy

'local variable(s) to hold property value(s)
Private mvarsProduto As String 'local copy
Private mvardtDataEntrega As Date 'local copy
Private mvariPrazoEntrega As Integer

'Indica se o Item de Pedido de Venda é um pai de grade
Private mvariPossuiGrade As Integer
'Guarda os itens filhos de grade
Private mvarcolItensRomaneioGrade As New Collection

'#############################################
'INSERIDO POR WAGNER
Private mvarsVersaoKit As String
Private mvarsVersaoKitBase As String

'####################################
'Inserido por Wagner
Public objInfoUsu As Object

Private mvarlStatus As Long
Private mvarlMotivoPerda As Long
Private mvarsObservacao As String

Private mvariPCSituacao As Integer 'situacao para calculo de preço
Private mvardPCPrecoUnitCalc 'preço unitario calculado (na aba de formação de precos)
Private mvarcolItensFormPreco As New Collection

Private mvarsProdutoSolicit As String
Private mvarsDescProdSolicit As String
Private mvarlMotivoNaoCobranca As Long
Private mvarsSRVProduto As String
Private mvarlNumIntGarantia As Long
Private mvarlNumIntItemContrato As Long
Private mvarlCodGarantia As Long
Private mvarlCodContrato As Long
Private mvariPeca As Integer
Private mvarlNumIntDocOV As Long
Private mvariHistorico As Integer
Private mvariMarcado As Integer

Public Property Let iMarcado(ByVal vData As Integer)
    mvariMarcado = vData
End Property

Public Property Get iMarcado() As Integer
    iMarcado = mvariMarcado
End Property

Public Property Let iPrazoEntrega(ByVal vData As Integer)
    mvariPrazoEntrega = vData
End Property

Public Property Get iPrazoEntrega() As Integer
    iPrazoEntrega = mvariPrazoEntrega
End Property

Public Property Let iHistorico(ByVal vData As Integer)
    mvariHistorico = vData
End Property

Public Property Get iHistorico() As Integer
    iHistorico = mvariHistorico
End Property

Public Property Let lNumIntDocOV(ByVal vData As Long)
    mvarlNumIntDocOV = vData
End Property

Public Property Get lNumIntDocOV() As Long
    lNumIntDocOV = mvarlNumIntDocOV
End Property

Public Property Let iPeca(ByVal vData As Integer)
    mvariPeca = vData
End Property

Public Property Get iPeca() As Integer
    iPeca = mvariPeca
End Property

Public Property Let lCodContrato(ByVal vData As Long)
    mvarlCodContrato = vData
End Property

Public Property Get lCodContrato() As Long
    lCodContrato = mvarlCodContrato
End Property

Public Property Let lCodGarantia(ByVal vData As Long)
    mvarlCodGarantia = vData
End Property

Public Property Get lCodGarantia() As Long
    lCodGarantia = mvarlCodGarantia
End Property

Public Property Let lNumIntItemContrato(ByVal vData As Long)
    mvarlNumIntItemContrato = vData
End Property

Public Property Get lNumIntItemContrato() As Long
    lNumIntItemContrato = mvarlNumIntItemContrato
End Property

Public Property Let lNumIntGarantia(ByVal vData As Long)
    mvarlNumIntGarantia = vData
End Property

Public Property Get lNumIntGarantia() As Long
    lNumIntGarantia = mvarlNumIntGarantia
End Property

Public Property Let sSRVProduto(ByVal vData As String)
    mvarsSRVProduto = vData
End Property

Public Property Get sSRVProduto() As String
    sSRVProduto = mvarsSRVProduto
End Property

Public Property Let lMotivoNaoCobranca(ByVal vData As Long)
    mvarlMotivoNaoCobranca = vData
End Property

Public Property Get lMotivoNaoCobranca() As Long
    lMotivoNaoCobranca = mvarlMotivoNaoCobranca
End Property

Public Property Let sDescProdSolicit(ByVal vData As String)
    mvarsDescProdSolicit = vData
End Property

Public Property Get sDescProdSolicit() As String
    sDescProdSolicit = mvarsDescProdSolicit
End Property

Public Property Let sProdutoSolicit(ByVal vData As String)
    mvarsProdutoSolicit = vData
End Property

Public Property Get sProdutoSolicit() As String
    sProdutoSolicit = mvarsProdutoSolicit
End Property

Public Property Set colItensFormPreco(ByVal vData As Collection)
    Set mvarcolItensFormPreco = vData
End Property

Public Property Get colItensFormPreco() As Collection
    Set colItensFormPreco = mvarcolItensFormPreco
End Property

Public Property Get iPCSituacao() As Integer
    iPCSituacao = mvariPCSituacao
End Property

Public Property Let iPCSituacao(ByVal vData As Integer)
    mvariPCSituacao = vData
End Property

Public Property Let dPCPrecoUnitCalc(ByVal vData As Double)
    mvardPCPrecoUnitCalc = vData
End Property

Public Property Get dPCPrecoUnitCalc() As Double
    dPCPrecoUnitCalc = mvardPCPrecoUnitCalc
End Property

Public Property Let sObservacao(ByVal vData As String)
    mvarsObservacao = vData
End Property

Public Property Get sObservacao() As String
    sObservacao = mvarsObservacao
End Property

Public Property Let lMotivoPerda(ByVal vData As Long)
    mvarlMotivoPerda = vData
End Property

Public Property Get lMotivoPerda() As Long
    lMotivoPerda = mvarlMotivoPerda
End Property

Public Property Let lStatus(ByVal vData As Long)
    mvarlStatus = vData
End Property

Public Property Get lStatus() As Long
    lStatus = mvarlStatus
End Property
'####################################

Public Property Let sVersaoKit(ByVal vData As String)
    mvarsVersaoKit = vData
End Property

Public Property Get sVersaoKit() As String
    sVersaoKit = mvarsVersaoKit
End Property

Public Property Let sVersaoKitBase(ByVal vData As String)
    mvarsVersaoKitBase = vData
End Property

Public Property Get sVersaoKitBase() As String
    sVersaoKitBase = mvarsVersaoKitBase
End Property
'FIM
'##############################################

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

Public Property Set colItensRomaneioGrade(ByVal vData As Collection)
    Set mvarcolItensRomaneioGrade = vData
End Property

Public Property Get colItensRomaneioGrade() As Collection
    Set colItensRomaneioGrade = mvarcolItensRomaneioGrade
End Property

Public Property Get iPossuiGrade() As Integer
    iPossuiGrade = mvariPossuiGrade
End Property

Public Property Let iPossuiGrade(ByVal vData As Integer)
    mvariPossuiGrade = vData
End Property


Public Property Let dtDataEntrega(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtDataEntrega = 5
    mvardtDataEntrega = vData
End Property


Public Property Get dtDataEntrega() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtDataEntrega
    dtDataEntrega = mvardtDataEntrega
End Property



Public Property Let sProduto(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sProduto = 5
    mvarsProduto = vData
End Property


Public Property Get sProduto() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sProduto
    sProduto = mvarsProduto
End Property



Public Property Let sDescricao(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sDescricao = 5
    mvarsDescricao = vData
End Property


Public Property Get sDescricao() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sDescricao
    sDescricao = mvarsDescricao
End Property



Public Property Let sUnidadeMed(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sUnidadeMed = 5
    mvarsUnidadeMed = vData
End Property


Public Property Get sUnidadeMed() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sUnidadeMed
    sUnidadeMed = mvarsUnidadeMed
End Property



Public Property Let sProdutoDescricao(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sProdutoDescricao = 5
    mvarsProdutoDescricao = vData
End Property


Public Property Get sProdutoDescricao() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sProdutoDescricao
    sProdutoDescricao = mvarsProdutoDescricao
End Property



Public Property Let lNumIntDoc(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lNumIntDoc = 5
    mvarlNumIntDoc = vData
End Property


Public Property Get lNumIntDoc() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lNumIntDoc
    lNumIntDoc = mvarlNumIntDoc
End Property



Public Property Let lCodOrcamento(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lCodPedido = 5
    mvarlCodOrcamento = vData
End Property


Public Property Get lCodOrcamento() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lCodPedido
    lCodOrcamento = mvarlCodOrcamento
End Property



Public Property Let iStatus(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iStatus = 5
    mvariStatus = vData
End Property


Public Property Get iStatus() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iStatus
    iStatus = mvariStatus
End Property



Public Property Let iClasseUM(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iClasseUM = 5
    mvariClasseUM = vData
End Property


Public Property Get iClasseUM() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iClasseUM
    iClasseUM = mvariClasseUM
End Property



Public Property Let iItem(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iItem = 5
    mvariItem = vData
End Property


Public Property Get iItem() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iItem
    iItem = mvariItem
End Property



Public Property Let iFilialEmpresa(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iFilialEmpresa = 5
    mvariFilialEmpresa = vData
End Property


Public Property Get iFilialEmpresa() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iFilialEmpresa
    iFilialEmpresa = mvariFilialEmpresa
End Property



Public Property Let dValorDesconto(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dValorDesconto = 5
    mvardValorDesconto = vData
End Property


Public Property Get dValorDesconto() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dValorDesconto
    dValorDesconto = mvardValorDesconto
End Property



Public Property Let dPrecoTotal(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dPrecoTotal = 5
    mvardPrecoTotal = vData
End Property


Public Property Get dPrecoTotal() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dPrecoTotal
    dPrecoTotal = mvardPrecoTotal
End Property



Public Property Let dPrecoUnitario(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dPrecoUnitario = 5
    mvardPrecoUnitario = vData
End Property


Public Property Get dPrecoUnitario() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dPrecoUnitario
    dPrecoUnitario = mvardPrecoUnitario
End Property



Public Property Let dQuantidade(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dQuantidade = 5
    mvardQuantidade = vData
End Property


Public Property Get dQuantidade() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dQuantidade
    dQuantidade = mvardQuantidade
End Property


Private Function ClassTributoDocItem_Ativar(objDoc As ClassTributoDoc) As Long
Dim lErro As Long
On Error GoTo Erro_ClassTributoDocItem_Ativar

    Set mvarobjDoc = objDoc
    
    Set mvarobjTributacaoDocItem = New ClassTributacaoDocItemAux

    If mvarobjTributacaoDocItem.Ativar(Me) <> SUCESSO Then
        Set mvarobjTributacaoDocItem = Nothing
        Error 27367
    End If

    ClassTributoDocItem_Ativar = SUCESSO

    Exit Function

Erro_ClassTributoDocItem_Ativar:

    ClassTributoDocItem_Ativar = Err

    Select Case Err

        Case 27367

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150897)

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

    ClassTributoDocItem_Desativar = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150898)

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

    ClassTributoDocItem_GuardarVarTela = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150899)

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
'            objVariavelCalculo.vValor = mvarobjTributacaoItemOV.sNaturezaOp
'            objVariavelCalculo.iStatus = mvarobjTributacaoItemOV.iNaturezaOpManual
'
'        Case "TIPO_TRIB"
'            objVariavelCalculo.vValor = mvarobjTributacaoItemOV.iTipoTributacao
'            objVariavelCalculo.iStatus = mvarobjTributacaoItemOV.iTipoTributacaoManual
'
'        Case "PRODUTO_CODIGO"
'            objVariavelCalculo.vValor = mvarsProduto
'            objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL
'
        Case "PRODUTO_VALOR"
            If mvardQuantidade <> 0 Then
                objVariavelCalculo.vValor = mvardPrecoTotal
                objVariavelCalculo.vValor = CDbl(Format(objVariavelCalculo.vValor, "0.00"))
            Else
                objVariavelCalculo.vValor = mvardPrecoTotal
            End If
            objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL

        Case "PRODUTO_DESC_GLOBAL" 'valor do desconto global rateado para o item
            objVariavelCalculo.vValor = mvardValorDescGlobal
            objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL

        Case "PRODUTO_DESP_GLOBAL" 'valor do (frete+seguro+outras desp) rateado para o item
            objVariavelCalculo.vValor = mvardValorDespGlobal
            objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL

        Case "PRODUTO_VALOR_BRUTO"
            If mvardQuantidade <> 0 Then
                objVariavelCalculo.vValor = (mvardPrecoTotal + mvardValorDesconto)
                objVariavelCalculo.vValor = CDbl(Format(objVariavelCalculo.vValor, "0.00"))
            Else
                objVariavelCalculo.vValor = Round(mvardPrecoTotal + mvardValorDesconto, 2)
            End If
            objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL

        Case "PRODUTO_QTDE"
            If mvardQuantidade <> 0 Then
                objVariavelCalculo.vValor = mvardQuantidade
            Else
                objVariavelCalculo.vValor = CDbl(Format(objVariavelCalculo.vValor, "0.00"))
            End If
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
'            objVariavelCalculo.vValor = mvarobjTributacaoItemOV.iICMSTipo
'            objVariavelCalculo.iStatus = mvarobjTributacaoItemOV.iICMSTipoManual
'
'        Case "ICMS_BASE"
'            objVariavelCalculo.vValor = mvarobjTributacaoItemOV.dICMSBase
'            objVariavelCalculo.iStatus = mvarobjTributacaoItemOV.iICMSBaseManual
'
'        Case "ICMS_RED_BASE"
'            objVariavelCalculo.vValor = mvarobjTributacaoItemOV.dICMSPercRedBase
'            objVariavelCalculo.iStatus = mvarobjTributacaoItemOV.iICMSPercRedBaseManual
'
'        Case "ICMS_ALIQUOTA"
'            objVariavelCalculo.vValor = mvarobjTributacaoItemOV.dICMSAliquota
'            objVariavelCalculo.iStatus = mvarobjTributacaoItemOV.iICMSAliquotaManual
'
'        Case "ICMS_VALOR"
'            objVariavelCalculo.vValor = mvarobjTributacaoItemOV.dICMSValor
'            objVariavelCalculo.iStatus = mvarobjTributacaoItemOV.iICMSValorManual
'
'        Case "ICMS_BASE_SUBST"
'            objVariavelCalculo.vValor = mvarobjTributacaoItemOV.dICMSSubstBase
'            objVariavelCalculo.iStatus = mvarobjTributacaoItemOV.iICMSSubstBaseManual
'
'        Case "ICMS_RED_BASE_SUBST"
'            objVariavelCalculo.vValor = mvarobjTributacaoItemOV.dICMSSubstPercRedBase
'            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO
'
'        Case "ICMS_ALIQUOTA_SUBST"
'            objVariavelCalculo.vValor = mvarobjTributacaoItemOV.dICMSSubstAliquota
'            objVariavelCalculo.iStatus = mvarobjTributacaoItemOV.iICMSSubstAliquotaManual
'
'        Case "ICMS_VALOR_SUBST"
'            objVariavelCalculo.vValor = mvarobjTributacaoItemOV.dICMSSubstValor
'            objVariavelCalculo.iStatus = mvarobjTributacaoItemOV.iICMSSubstValorManual
'
'        Case "ICMS_MARG_SUBST"
'            objVariavelCalculo.vValor = mvarobjTributacaoItemOV.dICMSSubstPercMVA
'            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO
'
'        Case "ICMS_VALOR_CREDITO"
'            objVariavelCalculo.vValor = mvarobjTributacaoItemOV.dICMSCredito
'
'        'para IPI
'        Case "IPI_TIPO"
'            objVariavelCalculo.vValor = mvarobjTributacaoItemOV.iIPITipo
'            objVariavelCalculo.iStatus = mvarobjTributacaoItemOV.iIPITipoManual
'
'        Case "IPI_BASE"
'            objVariavelCalculo.vValor = mvarobjTributacaoItemOV.dIPIBaseCalculo
'            objVariavelCalculo.iStatus = mvarobjTributacaoItemOV.iIPIBaseCalculoManual
'
'        Case "IPI_RED_BASE"
'            objVariavelCalculo.vValor = mvarobjTributacaoItemOV.dIPIPercRedBase
'            objVariavelCalculo.iStatus = mvarobjTributacaoItemOV.iIPIPercRedBaseManual
'
'        Case "IPI_ALIQUOTA"
'            objVariavelCalculo.vValor = mvarobjTributacaoItemOV.dIPIAliquota
'            objVariavelCalculo.iStatus = mvarobjTributacaoItemOV.iIPIAliquotaManual
'
'        Case "IPI_VALOR"
'            objVariavelCalculo.vValor = mvarobjTributacaoItemOV.dIPIValor
'            objVariavelCalculo.iStatus = mvarobjTributacaoItemOV.iIPIValorManual
'
'        Case "IPI_VALOR_CREDITO"
'            objVariavelCalculo.vValor = mvarobjTributacaoItemOV.dIPICredito
'
'        Case "PIS_VALOR_CREDITO"
'            objVariavelCalculo.vValor = mvarobjTributacaoItemOV.dPISCredito
'
'        Case "COFINS_VALOR_CREDITO"
'            objVariavelCalculo.vValor = mvarobjTributacaoItemOV.dCOFINSCredito
        
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150900)

    End Select

    Exit Function

End Function

Private Function ClassTributoDocItem_GuardarVar(objVariavelCalculo As ClassVariavelCalculo) As Long
'armazena info de variavel guardando se esta foi alterada ou nao desde ...
Dim lErro As Long
On Error GoTo Erro_ClassTributoDocItem_GuardarVar

    Select Case objVariavelCalculo.sIdentificador

'        Case "NAT_OPERACAO"
'            mvarobjTributacaoItemOV.sNaturezaOp = objVariavelCalculo.vValor
'            mvarobjTributacaoItemOV.iNaturezaOpManual = objVariavelCalculo.iStatus
'
'        Case "TIPO_TRIB"
'            mvarobjTributacaoItemOV.iTipoTributacao = objVariavelCalculo.vValor
'            mvarobjTributacaoItemOV.iTipoTributacaoManual = objVariavelCalculo.iStatus
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
'            mvarobjTributacaoItemOV.iICMSTipo = objVariavelCalculo.vValor
'            mvarobjTributacaoItemOV.iICMSTipoManual = objVariavelCalculo.iStatus
'
'        Case "ICMS_BASE"
'            mvarobjTributacaoItemOV.dICMSBase = objVariavelCalculo.vValor
'            mvarobjTributacaoItemOV.iICMSBaseManual = objVariavelCalculo.iStatus
'
'        Case "ICMS_RED_BASE"
'            mvarobjTributacaoItemOV.dICMSPercRedBase = objVariavelCalculo.vValor
'            mvarobjTributacaoItemOV.iICMSPercRedBaseManual = objVariavelCalculo.iStatus
'
'        Case "ICMS_ALIQUOTA"
'            mvarobjTributacaoItemOV.dICMSAliquota = objVariavelCalculo.vValor
'            mvarobjTributacaoItemOV.iICMSAliquotaManual = objVariavelCalculo.iStatus
'
'        Case "ICMS_VALOR"
'            mvarobjTributacaoItemOV.dICMSValor = objVariavelCalculo.vValor
'            mvarobjTributacaoItemOV.iICMSValorManual = objVariavelCalculo.iStatus
'
'        Case "ICMS_BASE_SUBST"
'            mvarobjTributacaoItemOV.dICMSSubstBase = objVariavelCalculo.vValor
'            mvarobjTributacaoItemOV.iICMSSubstBaseManual = objVariavelCalculo.iStatus
'
'        Case "ICMS_RED_BASE_SUBST"
'            mvarobjTributacaoItemOV.dICMSSubstPercRedBase = objVariavelCalculo.vValor
'
'        Case "ICMS_ALIQUOTA_SUBST"
'            mvarobjTributacaoItemOV.dICMSSubstAliquota = objVariavelCalculo.vValor
'            mvarobjTributacaoItemOV.iICMSSubstAliquotaManual = objVariavelCalculo.iStatus
'
'        Case "ICMS_VALOR_SUBST"
'            mvarobjTributacaoItemOV.dICMSSubstValor = objVariavelCalculo.vValor
'            mvarobjTributacaoItemOV.iICMSSubstValorManual = objVariavelCalculo.iStatus
'
'        Case "ICMS_MARG_SUBST"
'            mvarobjTributacaoItemOV.dICMSSubstPercMVA = objVariavelCalculo.vValor
'
'        Case "ICMS_VALOR_CREDITO"
'            mvarobjTributacaoItemOV.dICMSCredito = objVariavelCalculo.vValor
'
'        'para IPI
'        Case "IPI_TIPO"
'            mvarobjTributacaoItemOV.iIPITipo = objVariavelCalculo.vValor
'            mvarobjTributacaoItemOV.iIPITipoManual = objVariavelCalculo.iStatus
'
'        Case "IPI_BASE"
'            mvarobjTributacaoItemOV.dIPIBaseCalculo = objVariavelCalculo.vValor
'            mvarobjTributacaoItemOV.iIPIBaseManual = objVariavelCalculo.iStatus
'
'        Case "IPI_RED_BASE"
'            mvarobjTributacaoItemOV.dIPIPercRedBase = objVariavelCalculo.vValor
'            mvarobjTributacaoItemOV.iIPIPercRedBaseManual = objVariavelCalculo.iStatus
'
'        Case "IPI_ALIQUOTA"
'            mvarobjTributacaoItemOV.dIPIAliquota = objVariavelCalculo.vValor
'            mvarobjTributacaoItemOV.iIPIAliquotaManual = objVariavelCalculo.iStatus
'
'        Case "IPI_VALOR"
'            mvarobjTributacaoItemOV.dIPIValor = objVariavelCalculo.vValor
'            mvarobjTributacaoItemOV.iIPIValorManual = objVariavelCalculo.iStatus
'
'        Case "IPI_VALOR_CREDITO"
'            mvarobjTributacaoItemOV.dIPICredito = objVariavelCalculo.vValor
'
'        Case "PIS_VALOR_CREDITO"
'            mvarobjTributacaoItemOV.dPISCredito = objVariavelCalculo.vValor
'
'        Case "COFINS_VALOR_CREDITO"
'            mvarobjTributacaoItemOV.dCOFINSCredito = objVariavelCalculo.vValor
'
        Case Else
            'delegar p/tributacaoDocItem
            '"ICMS_MARG_SUBST"
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150901)

    End Select

    Exit Function

End Function

Private Property Get ClassTributoDocItem_objDoc() As ClassTributoDoc
    Set ClassTributoDocItem_objDoc = mvarobjDoc
End Property

Public Property Set objTributacaoItemOV(ByVal vData As Object)
    Set mvarobjTributacaoItemOV = vData
End Property


Public Property Get objTributacaoItemOV() As ClassTributacaoDocItem
    Set objTributacaoItemOV = mvarobjTributacaoItemOV
End Property

Public Property Set objTributacao(ByVal vData As Object)
    Set mvarobjTributacaoItemOV = vData
End Property

Public Property Get objTributacao() As ClassTributacaoDocItem
    Set objTributacao = mvarobjTributacaoItemOV
End Property

Public Property Set objInfoAdicDocItem(ByVal vData As Object)
    Set mvarobjInfoAdicDocItem = vData
End Property

Public Property Get objInfoAdicDocItem() As ClassInfoAdicDocItem
    Set objInfoAdicDocItem = mvarobjInfoAdicDocItem
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

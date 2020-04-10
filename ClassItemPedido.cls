VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassItemPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"colReserva"
Option Explicit

Implements ClassTributoDocItem

Public objInfoUsu As Object 'Inserido por Wagner 20/03/2006

Private mvarobjTributacaoDocItem As ClassTributacaoDocItemAux
Private mvarobjDoc As ClassTributoDoc 'link p/Doc

Private mvarobjInfoAdicDocItem As New ClassInfoAdicDocItem

Private mvariFilialEmpresa As Integer
Private mvarlCodPedido As Long
Private mvariItem As Integer
Private mvarsProduto As String
Private mvardQuantidade As Double
Private mvardPrecoUnitario As Double
Private mvardPrecoTotal As Double
Private mvarsUnidadeMed As String
Private mvardValorDesconto As Double
Private mvardValorDescGlobal As Double
Private mvardValorDespGlobal As Double
Private mvardtDataEntrega As Date
Private mvarsProdutoDescricao As String
Private mvarsLote As String
Private mvardValorAbatComissao As Double
Private mvardQuantCancelada As Double
Private mvardQuantReservada As Double
Private mvarsProdutoNomeReduzido As String
Private mvarsUMEstoque As String
Private mvariClasseUM As Integer
Private mvardQuantFaturada As Double
Private mvardQuantAFaturar As Double
Private mvardQuantOP As Double
Private mvardQuantSC As Double
Private mvarsDescricao As String
Private mvariStatus As Integer
Private mvariControleEstoque As Integer
Private mvarobjTributacaoItemPV As New ClassTributacaoDocItem
Private mvarlNumIntDoc As Long
Private mvarcolReserva As New colReserva
Private mvarcolEmbalagens As New Collection
Private mvariApropriacaoProd As Integer 'apropriacao de custo do produto
Private mvariPossuiGrade As Integer 'Indica se o Item de Pedido de Venda é um pai de grade
Private mvarcolItensRomaneioGrade As New Collection 'Guarda os itens filhos de grade
Private mvardPercDesc1 As Double
Private mvariTipoDesc1 As Integer
Private mvardPercDesc2 As Double
Private mvariTipoDesc2 As Integer
Private mvardPercDesc3 As Double
Private mvariTipoDesc3 As Integer
Private mvarsProdutoPai As String 'Codigo do produto pai de um item de grade
Private mvariOrigemPedido As Integer
Private mvariTipoOP As Integer
Private mvarcolItemPVEntrega As New Collection
Private mvariPeca As Integer
Private mvariRecebForaFaixa As Integer
Private mvardPercentMaisReceb As Double
Private mvardPercentMenosReceb As Double
Private mvardQuantFaturadaAMais As Double
Private mvariPrioridade As Integer
Private mvarlNumIntDocPV As Long
Private mvardPrecoBase As Double
Private mvariTabelaPreco As Integer

Public dComissao As Double

Private mvariTratamentoFaltaEstoque As Integer
Private mvarbTelaFaltaEstExibeBtnAplicarEmTodos As Boolean

Public Property Get bTelaFaltaEstExibeBtnAplicarEmTodos() As Boolean
    bTelaFaltaEstExibeBtnAplicarEmTodos = mvarbTelaFaltaEstExibeBtnAplicarEmTodos
End Property

Public Property Let bTelaFaltaEstExibeBtnAplicarEmTodos(ByVal vData As Boolean)
    mvarbTelaFaltaEstExibeBtnAplicarEmTodos = vData
End Property

Public Property Get iTratamentoFaltaEstoque() As Integer
    iTratamentoFaltaEstoque = mvariTratamentoFaltaEstoque
End Property

Public Property Let iTratamentoFaltaEstoque(ByVal vData As Integer)
    mvariTratamentoFaltaEstoque = vData
End Property

Public Property Set objInfoAdicDocItem(ByVal vData As Object)
    Set mvarobjInfoAdicDocItem = vData
End Property

Public Property Get objInfoAdicDocItem() As ClassInfoAdicDocItem
    Set objInfoAdicDocItem = mvarobjInfoAdicDocItem
End Property

Public Property Get iTabelaPreco() As Integer
    iTabelaPreco = mvariTabelaPreco
End Property

Public Property Let iTabelaPreco(ByVal vData As Integer)
    mvariTabelaPreco = vData
End Property

Public Property Let dPrecoBase(ByVal vData As Double)
    mvardPrecoBase = vData
End Property

Public Property Get dPrecoBase() As Double
    dPrecoBase = mvardPrecoBase
End Property

Public Property Let lNumIntDocPV(ByVal vData As Long)
    mvarlNumIntDocPV = vData
End Property

Public Property Get lNumIntDocPV() As Long
    lNumIntDocPV = mvarlNumIntDocPV
End Property

Public Property Get dQuantRealPV() As Double
    dQuantRealPV = mvardQuantidade - mvardQuantFaturada - mvardQuantCancelada
End Property

Public Property Get iPrioridade() As Integer
    iPrioridade = mvariPrioridade
End Property

Public Property Let iPrioridade(ByVal vData As Integer)
    mvariPrioridade = vData
End Property

Public Property Get dQuantFaturadaAMais() As Double
    dQuantFaturadaAMais = mvardQuantFaturadaAMais
End Property

Public Property Let dQuantFaturadaAMais(ByVal vData As Double)
    mvardQuantFaturadaAMais = vData
End Property

Public Property Get iRecebForaFaixa() As Integer
    iRecebForaFaixa = mvariRecebForaFaixa
End Property

Public Property Let iRecebForaFaixa(ByVal vData As Integer)
    mvariRecebForaFaixa = vData
End Property

Public Property Get dPercentMaisReceb() As Double
    dPercentMaisReceb = mvardPercentMaisReceb
End Property

Public Property Let dPercentMaisReceb(ByVal vData As Double)
    mvardPercentMaisReceb = vData
End Property

Public Property Get dPercentMenosReceb() As Double
    dPercentMenosReceb = mvardPercentMenosReceb
End Property

Public Property Let dPercentMenosReceb(ByVal vData As Double)
    mvardPercentMenosReceb = vData
End Property

Public Property Get iPeca() As Integer
    iPeca = mvariPeca
End Property

Public Property Let iPeca(ByVal vData As Integer)
    mvariPeca = vData
End Property

Public Property Set colItemPVEntrega(ByVal vData As Object)
    Set mvarcolItemPVEntrega = vData
End Property

Public Property Get colItemPVEntrega() As Collection
    Set colItemPVEntrega = mvarcolItemPVEntrega
End Property

Public Property Get iTipoOP() As Integer
    iTipoOP = mvariTipoOP
End Property

Public Property Let iTipoOP(ByVal vData As Integer)
    mvariTipoOP = vData
End Property

Public Property Get iOrigemPedido() As Integer
    iOrigemPedido = mvariOrigemPedido
End Property

Public Property Let iOrigemPedido(ByVal vData As Integer)
    mvariOrigemPedido = vData
End Property

Public Property Let sProdutoPai(ByVal vData As String)
    mvarsProdutoPai = vData
End Property

Public Property Get sProdutoPai() As String
    sProdutoPai = mvarsProdutoPai
End Property

Public Property Let iTipoDesc3(ByVal vData As Integer)
    mvariTipoDesc3 = vData
End Property

Public Property Get iTipoDesc3() As Integer
    iTipoDesc3 = mvariTipoDesc3
End Property

Public Property Let dPercDesc3(ByVal vData As Double)
    mvardPercDesc3 = vData
End Property

Public Property Get dPercDesc3() As Double
    dPercDesc3 = mvardPercDesc3
End Property

Public Property Let iTipoDesc2(ByVal vData As Integer)
    mvariTipoDesc2 = vData
End Property

Public Property Get iTipoDesc2() As Integer
    iTipoDesc2 = mvariTipoDesc2
End Property

Public Property Let dPercDesc2(ByVal vData As Double)
    mvardPercDesc2 = vData
End Property

Public Property Get dPercDesc2() As Double
    dPercDesc2 = mvardPercDesc2
End Property

Public Property Let iTipoDesc1(ByVal vData As Integer)
    mvariTipoDesc1 = vData
End Property

Public Property Get iTipoDesc1() As Integer
    iTipoDesc1 = mvariTipoDesc1
End Property

Public Property Let dPercDesc1(ByVal vData As Double)
    mvardPercDesc1 = vData
End Property

Public Property Get dPercDesc1() As Double
    dPercDesc1 = mvardPercDesc1
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

Public Property Get iApropriacaoProd() As Integer
    iApropriacaoProd = mvariApropriacaoProd
End Property

Public Property Let iApropriacaoProd(ByVal vData As Integer)
    mvariApropriacaoProd = vData
End Property

Public Property Set colReserva(ByVal vData As Object)
    Set mvarcolReserva = vData
End Property

Public Property Get colReserva() As colReserva
    Set colReserva = mvarcolReserva
End Property
'Janaina
Public Property Set colEmbalagens(ByVal vData As Collection)
    Set mvarcolEmbalagens = vData
End Property
'Janaina
Public Property Get colEmbalagens() As Collection
    Set colEmbalagens = mvarcolEmbalagens
End Property
Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let dQuantOP(ByVal vData As Double)
    mvardQuantOP = vData
End Property

Public Property Get dQuantOP() As Double
    dQuantOP = mvardQuantOP
End Property

Public Property Let dQuantSC(ByVal vData As Double)
    mvardQuantSC = vData
End Property

Public Property Get dQuantSC() As Double
    dQuantSC = mvardQuantSC
End Property

Public Property Set objTributacaoItemPV(ByVal vData As Object)
    Set mvarobjTributacaoItemPV = vData
End Property

Public Property Get objTributacaoItemPV() As ClassTributacaoDocItem
    Set objTributacaoItemPV = mvarobjTributacaoItemPV
End Property

Public Property Let iControleEstoque(ByVal vData As Integer)
    mvariControleEstoque = vData
End Property

Public Property Get iControleEstoque() As Integer
    iControleEstoque = mvariControleEstoque
End Property

Public Property Let iStatus(ByVal vData As Integer)
    mvariStatus = vData
End Property

Public Property Get iStatus() As Integer
    iStatus = mvariStatus
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let dQuantFaturada(ByVal vData As Double)
    mvardQuantFaturada = vData
End Property

Public Property Get dQuantFaturada() As Double
    dQuantFaturada = mvardQuantFaturada
End Property

Public Property Let dQuantAFaturar(ByVal vData As Double)
    mvardQuantAFaturar = vData
End Property

Public Property Get dQuantAFaturar() As Double
    dQuantAFaturar = mvardQuantAFaturar
End Property

Public Property Let iClasseUM(ByVal vData As Integer)
    mvariClasseUM = vData
End Property

Public Property Get iClasseUM() As Integer
    iClasseUM = mvariClasseUM
End Property

Public Property Let sUMEstoque(ByVal vData As String)
    mvarsUMEstoque = vData
End Property

Public Property Get sUMEstoque() As String
    sUMEstoque = mvarsUMEstoque
End Property

Public Property Let sLote(ByVal vData As String)
    mvarsLote = vData
End Property

Public Property Get sLote() As String
    sLote = mvarsLote
End Property

Public Property Let sUnidadeMed(ByVal vData As String)
    mvarsUnidadeMed = vData
End Property

Public Property Get sUnidadeMed() As String
    sUnidadeMed = mvarsUnidadeMed
End Property

Public Property Let sProdutoNomeReduzido(ByVal vData As String)
    mvarsProdutoNomeReduzido = vData
End Property

Public Property Get sProdutoNomeReduzido() As String
    sProdutoNomeReduzido = mvarsProdutoNomeReduzido
End Property
'
'Public Property Set colReservaItem(ByVal vData As Object)
'    Set mvarcolReservaItem = vData
'End Property
'
'Public Property Get colReservaItem() As colReservaItem
'    Set colReservaItem = mvarcolReservaItem
'End Property

Public Property Let dQuantReservada(ByVal vData As Double)
    mvardQuantReservada = vData
End Property

Public Property Get dQuantReservada() As Double
    dQuantReservada = mvardQuantReservada
End Property

Public Property Let dQuantCancelada(ByVal vData As Double)
    mvardQuantCancelada = vData
End Property

Public Property Get dQuantCancelada() As Double
    dQuantCancelada = mvardQuantCancelada
End Property

Public Property Let dValorAbatComissao(ByVal vData As Double)
    mvardValorAbatComissao = vData
End Property

Public Property Get dValorAbatComissao() As Double
    dValorAbatComissao = mvardValorAbatComissao
End Property

Public Property Let sProdutoDescricao(ByVal vData As String)
    mvarsProdutoDescricao = vData
End Property

Public Property Get sProdutoDescricao() As String
    sProdutoDescricao = mvarsProdutoDescricao
End Property

Public Property Let dtDataEntrega(ByVal vData As Date)
    mvardtDataEntrega = vData
End Property

Public Property Get dtDataEntrega() As Date
    dtDataEntrega = mvardtDataEntrega
End Property

Public Property Let dValorDesconto(ByVal vData As Double)
    mvardValorDesconto = vData
End Property

Public Property Get dValorDesconto() As Double
    dValorDesconto = mvardValorDesconto
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

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let iItem(ByVal vData As Integer)
    mvariItem = vData
End Property

Public Property Get iItem() As Integer
    iItem = mvariItem
End Property

Public Property Let lCodPedido(ByVal vData As Long)
    mvarlCodPedido = vData
End Property

Public Property Get lCodPedido() As Long
    lCodPedido = mvarlCodPedido
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Private Sub Class_Initialize()
    mvarobjInfoAdicDocItem.iTipoDoc = TIPODOC_INFOADIC_ITEMPV
    dComissao = -1
End Sub

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150902)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150903)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150904)

    End Select

    Exit Function

End Function

Private Property Get ClassTributoDocItem_objDoc() As ClassTributoDoc
    Set ClassTributoDocItem_objDoc = mvarobjDoc
End Property

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
            If mvardQuantAFaturar <> 0 Then
                objVariavelCalculo.vValor = mvardPrecoTotal * mvardQuantAFaturar / (mvardQuantidade - mvardQuantCancelada)
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
            If mvardQuantAFaturar <> 0 Then
                objVariavelCalculo.vValor = (mvardPrecoTotal + mvardValorDesconto) * mvardQuantAFaturar / (mvardQuantidade - mvardQuantCancelada)
                objVariavelCalculo.vValor = CDbl(Format(objVariavelCalculo.vValor, "0.00"))
            Else
                objVariavelCalculo.vValor = Round(mvardPrecoTotal + mvardValorDesconto, 2)
            End If
            objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL

        Case "PRODUTO_QTDE"
            If mvardQuantAFaturar <> 0 Then
                objVariavelCalculo.vValor = mvardQuantAFaturar
            Else
                objVariavelCalculo.vValor = mvardQuantidade - mvardQuantCancelada
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
            If objTributacao.sProduto = "" Then objTributacao.sProduto = mvarsProduto
            
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

Public Property Set objTributacao(ByVal vData As Object)
    Set mvarobjTributacaoItemPV = vData
End Property

Public Property Get objTributacao() As ClassTributacaoDocItem
    Set objTributacao = mvarobjTributacaoItemPV
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



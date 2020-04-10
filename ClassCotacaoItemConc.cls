VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCotacaoItemConc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarlNumIntDoc As Long 'local copy
Private mvardValorPresente As Double 'local copy
Private mvariEscolhido As Integer 'local copy
Private mvarsMotivoEscolha As String 'local copy
Private mvardQuantidadeComprar As Double 'local copy
Private mvardtDataEntrega As Date 'local copy
Private mvardPrecoAjustado As Double 'local copy
'local variable(s) to hold property value(s)
Private mvarsFornecedor As String 'local copy
Private mvarsFilial As String 'local copy
Private mvarsCondPagto As String 'local copy
Private mvardPrecoUnitario As Double 'local copy
Private mvardCreditoICMS As Double 'local copy
Private mvardCreditoIPI As Double 'local copy
Private mvarlPedCotacao As Long 'local copy
Private mvardtDataValidade As Date 'local copy
Private mvariPrazoEntrega As Integer 'local copy
'local variable(s) to hold property value(s)
Private mvardQuantEntrega As Double 'local copy
Private mvardPreferencia As Double 'local copy
'local variable(s) to hold property value(s)
Private mvarlItemCotacao As Long 'local copy
Private mvariSelecionada As Integer
'local variable(s) to hold property value(s)
Private mvardQuantCotada As Double 'local copy
Private mvarsUMCompra As String 'local copy
Private mvardAliquotaIPI As Double 'local copy
Private mvardAliquotaICMS As Double 'local copy
Private mvariTipoTributacao As Integer
'local variable(s) to hold property value(s)
Private mvardtDataPedidoCotacao As Date 'local copy
Private mvariMoeda As Integer
Private mvardTaxa As Double
Private mvardtDataVencPriParc As Date
Private mvardPrazoMedio As Double

Public Property Let dPrazoMedio(ByVal vData As Double)
    mvardPrazoMedio = vData
End Property

Public Property Get dPrazoMedio() As Double
    dPrazoMedio = mvardPrazoMedio
End Property

Public Property Let dtDataVencPriParc(ByVal vData As Date)
    mvardtDataVencPriParc = vData
End Property

Public Property Get dtDataVencPriParc() As Date
    dtDataVencPriParc = mvardtDataVencPriParc
End Property

Public Property Let iMoeda(ByVal vData As Integer)
    mvariMoeda = vData
End Property

Public Property Get iMoeda() As Integer
    iMoeda = mvariMoeda
End Property

Public Property Let dTaxa(ByVal vData As Double)
    mvardTaxa = vData
End Property

Public Property Get dTaxa() As Double
    dTaxa = mvardTaxa
End Property

Public Property Let dtDataPedidoCotacao(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtDataPedidoCotacao = 5
    mvardtDataPedidoCotacao = vData
End Property


Public Property Get dtDataPedidoCotacao() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtDataPedidoCotacao
    dtDataPedidoCotacao = mvardtDataPedidoCotacao
End Property



Public Property Let iTipoTributacao(ByVal vData As Integer)
    mvariTipoTributacao = vData
End Property

Public Property Get iTipoTributacao() As Integer
    iTipoTributacao = mvariTipoTributacao
End Property

Public Property Let dAliquotaICMS(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dAliquotaICMS = 5
    mvardAliquotaICMS = vData
End Property


Public Property Get dAliquotaICMS() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dAliquotaICMS
    dAliquotaICMS = mvardAliquotaICMS
End Property



Public Property Let dAliquotaIPI(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dAliquotaIPI = 5
    mvardAliquotaIPI = vData
End Property


Public Property Get dAliquotaIPI() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dAliquotaIPI
    dAliquotaIPI = mvardAliquotaIPI
End Property



Public Property Let sUMCompra(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sUMCompra = 5
    mvarsUMCompra = vData
End Property


Public Property Get sUMCompra() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sUMCompra
    sUMCompra = mvarsUMCompra
End Property



Public Property Let dQuantCotada(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dQuantCotada = 5
    mvardQuantCotada = vData
End Property


Public Property Get dQuantCotada() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dQuantCotada
    dQuantCotada = mvardQuantCotada
End Property

Public Property Let iSelecionada(ByVal vData As Integer)
    mvariSelecionada = vData
End Property

Public Property Get iSelecionada() As Integer
    iSelecionada = mvariSelecionada
End Property

Public Property Let lItemCotacao(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lItemCotacao = 5
    mvarlItemCotacao = vData
End Property


Public Property Get lItemCotacao() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lItemCotacao
    lItemCotacao = mvarlItemCotacao
End Property



Public Property Let dPreferencia(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dPreferencia = 5
    mvardPreferencia = vData
End Property


Public Property Get dPreferencia() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dPreferencia
    dPreferencia = mvardPreferencia
End Property



Public Property Let dQuantEntrega(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dQuantEntrega = 5
    mvardQuantEntrega = vData
End Property


Public Property Get dQuantEntrega() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dQuantEntrega
    dQuantEntrega = mvardQuantEntrega
End Property



Public Property Let iPrazoEntrega(ByVal vData As Integer)
Attribute iPrazoEntrega.VB_Description = "Prazo de Entrega em dias."
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iPrazoEntrega = 5
    mvariPrazoEntrega = vData
End Property


Public Property Get iPrazoEntrega() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iPrazoEntrega
    iPrazoEntrega = mvariPrazoEntrega
End Property



Public Property Let dtDataValidade(ByVal vData As Date)
Attribute dtDataValidade.VB_Description = "Data da Validade do Pedido de Cotação."
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtDataValidade = 5
    mvardtDataValidade = vData
End Property


Public Property Get dtDataValidade() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtDataValidade
    dtDataValidade = mvardtDataValidade
End Property



Public Property Let lPedCotacao(ByVal vData As Long)
Attribute lPedCotacao.VB_Description = "Código do Pedido de Cotação."
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lPedCotacao = 5
    mvarlPedCotacao = vData
End Property


Public Property Get lPedCotacao() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lPedCotacao
    lPedCotacao = mvarlPedCotacao
End Property



Public Property Let dCreditoIPI(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dCreditoIPI = 5
    mvardCreditoIPI = vData
End Property


Public Property Get dCreditoIPI() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dCreditoIPI
    dCreditoIPI = mvardCreditoIPI
End Property



Public Property Let dCreditoICMS(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dCreditoICMS = 5
    mvardCreditoICMS = vData
End Property


Public Property Get dCreditoICMS() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dCreditoICMS
    dCreditoICMS = mvardCreditoICMS
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



Public Property Let sCondPagto(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sCondPagto = 5
    mvarsCondPagto = vData
End Property


Public Property Get sCondPagto() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sCondPagto
    sCondPagto = mvarsCondPagto
End Property



Public Property Let sFilial(ByVal vData As String)
Attribute sFilial.VB_Description = "Filial do Fornecedor."
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sFilial = 5
    mvarsFilial = vData
End Property


Public Property Get sFilial() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sFilial
    sFilial = mvarsFilial
End Property



Public Property Let sFornecedor(ByVal vData As String)
Attribute sFornecedor.VB_Description = "Nome Reduzido do Fornecedor."
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sFornecedor = 5
    mvarsFornecedor = vData
End Property


Public Property Get sFornecedor() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sFornecedor
    sFornecedor = mvarsFornecedor
End Property



Public Property Let dPrecoAjustado(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dPrecoAjustado = 5
    mvardPrecoAjustado = vData
End Property


Public Property Get dPrecoAjustado() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dPrecoAjustado
    dPrecoAjustado = mvardPrecoAjustado
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



Public Property Let dQuantidadeComprar(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dQuantidadeComprar = 5
    mvardQuantidadeComprar = vData
End Property


Public Property Get dQuantidadeComprar() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dQuantidadeComprar
    dQuantidadeComprar = mvardQuantidadeComprar
End Property



Public Property Let sMotivoEscolha(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sMotivoEscolha = 5
    mvarsMotivoEscolha = vData
End Property


Public Property Get sMotivoEscolha() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sMotivoEscolha
    sMotivoEscolha = mvarsMotivoEscolha
End Property



Public Property Let iEscolhido(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iEscolhido = 5
    mvariEscolhido = vData
End Property


Public Property Get iEscolhido() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iEscolhido
    iEscolhido = mvariEscolhido
End Property



Public Property Let dValorPresente(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dValorPresente = 5
    mvardValorPresente = vData
End Property


Public Property Get dValorPresente() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dValorPresente
    dValorPresente = mvardValorPresente
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

Private Sub Class_Initialize()
    dtDataVencPriParc = DATA_NULA
End Sub

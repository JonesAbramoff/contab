VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassGeracaoOPItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarobjItemOP As New ClassItemOP
Private mvardQtdeFalta As Double
Private mvarsClienteReduzido As String
Private mvariFilialCliente As Integer
Private mvardtEmissaoPV As Date
Private mvardtEntregaItemPV As Date
Private mvariSelecionado As Integer
Private mvarsProdutoDescricao As String
Private mvariPrioridade As Integer

Public Property Let iPrioridade(ByVal vData As Integer)
    mvariPrioridade = vData
End Property

Public Property Get iPrioridade() As Integer
    iPrioridade = mvariPrioridade
End Property

Public Property Let sProdutoDescricao(ByVal vData As String)
    mvarsProdutoDescricao = vData
End Property

Public Property Get sProdutoDescricao() As String
    sProdutoDescricao = mvarsProdutoDescricao
End Property

Public Property Let iSelecionado(ByVal vData As Integer)
    mvariSelecionado = vData
End Property

Public Property Get iSelecionado() As Integer
    iSelecionado = mvariSelecionado
End Property

Public Property Let dtEntregaItemPV(ByVal vData As Date)
    mvardtEntregaItemPV = vData
End Property

Public Property Get dtEntregaItemPV() As Date
    dtEntregaItemPV = mvardtEntregaItemPV
End Property

Public Property Let dtEmissaoPV(ByVal vData As Date)
    mvardtEmissaoPV = vData
End Property

Public Property Get dtEmissaoPV() As Date
    dtEmissaoPV = mvardtEmissaoPV
End Property

Public Property Let iFilialCliente(ByVal vData As Integer)
    mvariFilialCliente = vData
End Property

Public Property Get iFilialCliente() As Integer
    iFilialCliente = mvariFilialCliente
End Property

Public Property Let sClienteReduzido(ByVal vData As String)
    mvarsClienteReduzido = vData
End Property

Public Property Get sClienteReduzido() As String
    sClienteReduzido = mvarsClienteReduzido
End Property

Public Property Let dQtdeFalta(ByVal vData As Double)
    mvardQtdeFalta = vData
End Property

Public Property Get dQtdeFalta() As Double
    dQtdeFalta = mvardQtdeFalta
End Property

Public Property Set objItemOP(ByVal vData As Object)
    Set mvarobjItemOP = vData
End Property

Public Property Get objItemOP() As ClassItemOP
    Set objItemOP = mvarobjItemOP
End Property




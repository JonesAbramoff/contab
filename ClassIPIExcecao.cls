VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassIPIExcecao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarsCategoriaProduto As String
Private mvarsCategoriaProdutoItem As String
Private mvarsCategoriaCliente As String
Private mvarsCategoriaClienteItem As String
Private mvariTipo As Integer
Private mvardPercRedBaseCalculo As Double
Private mvardAliquota As Double
Private mvardPercMargemLucro As Double
Private mvarsFundamentacao As String
Private mvariPrioridade As Integer
Private mvariTipoCalculo As Integer
Private mvardAliquotaRS As Double

Public Property Let dAliquotaRS(ByVal vData As Double)
    mvardAliquotaRS = vData
End Property

Public Property Get dAliquotaRS() As Double
    dAliquotaRS = mvardAliquotaRS
End Property

Public Property Let iTipoCalculo(ByVal vData As Integer)
    mvariTipoCalculo = vData
End Property

Public Property Get iTipoCalculo() As Integer
    iTipoCalculo = mvariTipoCalculo
End Property

Public Property Let iPrioridade(ByVal vData As Integer)
    mvariPrioridade = vData
End Property

Public Property Get iPrioridade() As Integer
    iPrioridade = mvariPrioridade
End Property

Public Property Let sFundamentacao(ByVal vData As String)
    mvarsFundamentacao = vData
End Property

Public Property Get sFundamentacao() As String
    sFundamentacao = mvarsFundamentacao
End Property

Public Property Let dPercMargemLucro(ByVal vData As Double)
    mvardPercMargemLucro = vData
End Property

Public Property Get dPercMargemLucro() As Double
    dPercMargemLucro = mvardPercMargemLucro
End Property

Public Property Let dAliquota(ByVal vData As Double)
    mvardAliquota = vData
End Property

Public Property Get dAliquota() As Double
    dAliquota = mvardAliquota
End Property

Public Property Let dPercRedBaseCalculo(ByVal vData As Double)
    mvardPercRedBaseCalculo = vData
End Property

Public Property Get dPercRedBaseCalculo() As Double
    dPercRedBaseCalculo = mvardPercRedBaseCalculo
End Property

Public Property Let iTipo(ByVal vData As Integer)
    mvariTipo = vData
End Property

Public Property Get iTipo() As Integer
    iTipo = mvariTipo
End Property

Public Property Let sCategoriaClienteItem(ByVal vData As String)
    mvarsCategoriaClienteItem = vData
End Property

Public Property Get sCategoriaClienteItem() As String
    sCategoriaClienteItem = mvarsCategoriaClienteItem
End Property

Public Property Let sCategoriaCliente(ByVal vData As String)
    mvarsCategoriaCliente = vData
End Property

Public Property Get sCategoriaCliente() As String
    sCategoriaCliente = mvarsCategoriaCliente
End Property

Public Property Let sCategoriaProdutoItem(ByVal vData As String)
    mvarsCategoriaProdutoItem = vData
End Property

Public Property Get sCategoriaProdutoItem() As String
    sCategoriaProdutoItem = mvarsCategoriaProdutoItem
End Property

Public Property Let sCategoriaProduto(ByVal vData As String)
    mvarsCategoriaProduto = vData
End Property

Public Property Get sCategoriaProduto() As String
    sCategoriaProduto = mvarsCategoriaProduto
End Property




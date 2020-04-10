VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassMaquinas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariCodigo As Integer
Private mvariFilialEmpresa As Integer
Private mvarlNumIntDoc As Long
Private mvarsNomeReduzido As String
Private mvarsDescricao As String
'------------------------------------------------
'Incluido por Jorge Specian
Private mvardTempoMovimentacao As Double
Private mvardTempoPreparacao As Double
Private mvardTempoDescarga As Double
Private mvariRecurso As Integer
Private mvarcolProdutos As New Collection
Private mvarcolTipoOperadores As New Collection
Private mvardCustoHora As Double
Private mvarsProduto As String
Private mvardPeso As Double
Private mvardComprimento As Double
Private mvardEspessura As Double
Private mvardLargura As Double

Public Property Let dLargura(ByVal vData As Double)
    mvardLargura = vData
End Property

Public Property Get dLargura() As Double
    dLargura = mvardLargura
End Property

Public Property Let dEspessura(ByVal vData As Double)
    mvardEspessura = vData
End Property

Public Property Get dEspessura() As Double
    dEspessura = mvardEspessura
End Property

Public Property Let dComprimento(ByVal vData As Double)
    mvardComprimento = vData
End Property

Public Property Get dComprimento() As Double
    dComprimento = mvardComprimento
End Property

Public Property Let dPeso(ByVal vData As Double)
    mvardPeso = vData
End Property

Public Property Get dPeso() As Double
    dPeso = mvardPeso
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let dCustoHora(ByVal vData As Double)
    mvardCustoHora = vData
End Property

Public Property Get dCustoHora() As Double
    dCustoHora = mvardCustoHora
End Property

Public Property Let dTempoMovimentacao(ByVal vData As Double)
    mvardTempoMovimentacao = vData
End Property

Public Property Get dTempoMovimentacao() As Double
    dTempoMovimentacao = mvardTempoMovimentacao
End Property

Public Property Let dTempoPreparacao(ByVal vData As Double)
    mvardTempoPreparacao = vData
End Property

Public Property Get dTempoPreparacao() As Double
    dTempoPreparacao = mvardTempoPreparacao
End Property

Public Property Let dTempoDescarga(ByVal vData As Double)
    mvardTempoDescarga = vData
End Property

Public Property Get dTempoDescarga() As Double
    dTempoDescarga = mvardTempoDescarga
End Property

Public Property Let iRecurso(ByVal vData As Integer)
    mvariRecurso = vData
End Property

Public Property Get iRecurso() As Integer
    iRecurso = mvariRecurso
End Property

Public Property Set colProdutos(ByVal vData As Collection)
    Set mvarcolProdutos = vData
End Property

Public Property Get colProdutos() As Collection
    Set colProdutos = mvarcolProdutos
End Property

Public Property Set colTipoOperadores(ByVal vData As Collection)
    Set mvarcolTipoOperadores = vData
End Property

Public Property Get colTipoOperadores() As Collection
    Set colTipoOperadores = mvarcolTipoOperadores
End Property
'------------------------------------------------

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let sNomeReduzido(ByVal vData As String)
    mvarsNomeReduzido = vData
End Property

Public Property Get sNomeReduzido() As String
    sNomeReduzido = mvarsNomeReduzido
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let iCodigo(ByVal vData As Integer)
    mvariCodigo = vData
End Property

Public Property Get iCodigo() As Integer
    iCodigo = mvariCodigo
End Property




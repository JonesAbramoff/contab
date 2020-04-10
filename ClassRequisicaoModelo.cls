VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRequisicaoModelo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"

Private mvarlNumIntDoc As Long
Private mvariFilialEmpresa As Integer
Private mvarlCodigo As Long
Private mvarsDescricao As String
Private mvariUrgente As Integer
Private mvarlRequisitante As Long
Private mvarsCcl As String
Private mvariFilialCompra As Integer
Private mvariTipoDestido As Integer
Private mvarlFornCliDestino As Long
Private mvariFilialDestino As Integer
Private mvarlObservacao As Long
Private mvarsObservacao As String
Private mvarcolItensReqModelo As New Collection
Private mvariTipoTributacao As Integer 'local copy
'local variable(s) to hold property value(s)
Private mvarsNomeRequisitante As String 'local copy
Private mvarsNomeFilialEmpresa As String 'local copy
Private mvarsNomeFilialCompra As String 'local copy
Public Property Let sNomeFilialCompra(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sNomeFilialCompra = 5
    mvarsNomeFilialCompra = vData
End Property


Public Property Get sNomeFilialCompra() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sNomeFilialCompra
    sNomeFilialCompra = mvarsNomeFilialCompra
End Property



Public Property Let sNomeFilialEmpresa(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sNomeFilialEmpresa = 5
    mvarsNomeFilialEmpresa = vData
End Property


Public Property Get sNomeFilialEmpresa() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sNomeFilialEmpresa
    sNomeFilialEmpresa = mvarsNomeFilialEmpresa
End Property



Public Property Let sNomeRequisitante(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sNomeRequisitante = 5
    mvarsNomeRequisitante = vData
End Property


Public Property Get sNomeRequisitante() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sNomeRequisitante
    sNomeRequisitante = mvarsNomeRequisitante
End Property



Public Property Let iTipoTributacao(ByVal vData As Integer)
    mvariTipoTributacao = vData
End Property


Public Property Get iTipoTributacao() As Integer
    iTipoTributacao = mvariTipoTributacao
End Property



Public Property Set colItensReqModelo(ByVal vData As Object)
    Set mvarcolItensReqModelo = vData
End Property

Public Property Get colItensReqModelo() As Collection
    Set colItensReqModelo = mvarcolItensReqModelo
End Property

Public Property Let lObservacao(ByVal vData As Long)
    mvarlObservacao = vData
End Property

Public Property Get lObservacao() As Long
    lObservacao = mvarlObservacao
End Property

Public Property Let iFilialDestino(ByVal vData As Integer)
    mvariFilialDestino = vData
End Property

Public Property Get iFilialDestino() As Integer
    iFilialDestino = mvariFilialDestino
End Property

Public Property Let lFornCliDestino(ByVal vData As Long)
    mvarlFornCliDestino = vData
End Property

Public Property Get lFornCliDestino() As Long
    lFornCliDestino = mvarlFornCliDestino
End Property

Public Property Let iTipoDestido(ByVal vData As Integer)
    mvariTipoDestido = vData
End Property

Public Property Get iTipoDestido() As Integer
    iTipoDestido = mvariTipoDestido
End Property

Public Property Let iFilialCompra(ByVal vData As Integer)
    mvariFilialCompra = vData
End Property

Public Property Get iFilialCompra() As Integer
    iFilialCompra = mvariFilialCompra
End Property

Public Property Let sCcl(ByVal vData As String)
    mvarsCcl = vData
End Property

Public Property Get sCcl() As String
    sCcl = mvarsCcl
End Property

Public Property Let lRequisitante(ByVal vData As Long)
    mvarlRequisitante = vData
End Property

Public Property Get lRequisitante() As Long
    lRequisitante = mvarlRequisitante
End Property

Public Property Let iUrgente(ByVal vData As Integer)
    mvariUrgente = vData
End Property

Public Property Get iUrgente() As Integer
    iUrgente = mvariUrgente
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let lCodigo(ByVal vData As Long)
    mvarlCodigo = vData
End Property

Public Property Get lCodigo() As Long
    lCodigo = mvarlCodigo
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let sObservacao(ByVal vData As String)
    mvarsObservacao = vData
End Property

Public Property Get sObservacao() As String
    sObservacao = mvarsObservacao
End Property


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassApontPRJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDoc As Long
Private mvarlCodigo As Long
Private mvarlNumIntDocPRJ As Long
Private mvarlNumIntDocEtapa As Long
Private mvardtData As Date
Private mvarsDescricao As String
Private mvarsObservacao As String
Private mvarsUsuario As String
Private mvardtDataRegistro As Date
Private mvardHoraRegistro As Double
Private mvarcolMaquinas As New Collection
Private mvarcolMaoDeObra As New Collection
Private mvarcolMateriaPrima As New Collection

Public Property Set colMateriaPrima(ByVal vData As Collection)
    Set mvarcolMateriaPrima = vData
End Property

Public Property Get colMateriaPrima() As Collection
    Set colMateriaPrima = mvarcolMateriaPrima
End Property

Public Property Set colMaoDeObra(ByVal vData As Collection)
    Set mvarcolMaoDeObra = vData
End Property

Public Property Get colMaoDeObra() As Collection
    Set colMaoDeObra = mvarcolMaoDeObra
End Property

Public Property Set colMaquinas(ByVal vData As Collection)
    Set mvarcolMaquinas = vData
End Property

Public Property Get colMaquinas() As Collection
    Set colMaquinas = mvarcolMaquinas
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let lCodigo(ByVal vData As Long)
    mvarlCodigo = vData
End Property

Public Property Get lCodigo() As Long
    lCodigo = mvarlCodigo
End Property

Public Property Let lNumIntDocPRJ(ByVal vData As Long)
    mvarlNumIntDocPRJ = vData
End Property

Public Property Get lNumIntDocPRJ() As Long
    lNumIntDocPRJ = mvarlNumIntDocPRJ
End Property

Public Property Let lNumIntDocEtapa(ByVal vData As Long)
    mvarlNumIntDocEtapa = vData
End Property

Public Property Get lNumIntDocEtapa() As Long
    lNumIntDocEtapa = mvarlNumIntDocEtapa
End Property

Public Property Let dtData(ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let sObservacao(ByVal vData As String)
    mvarsObservacao = vData
End Property

Public Property Get sObservacao() As String
    sObservacao = mvarsObservacao
End Property

Public Property Let sUsuario(ByVal vData As String)
    mvarsUsuario = vData
End Property

Public Property Get sUsuario() As String
    sUsuario = mvarsUsuario
End Property

Public Property Let dtDataRegistro(ByVal vData As Date)
    mvardtDataRegistro = vData
End Property

Public Property Get dtDataRegistro() As Date
    dtDataRegistro = mvardtDataRegistro
End Property

Public Property Let dHoraRegistro(ByVal vData As Double)
    mvardHoraRegistro = vData
End Property

Public Property Get dHoraRegistro() As Double
    dHoraRegistro = mvardHoraRegistro
End Property

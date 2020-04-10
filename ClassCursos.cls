VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCursos"
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
Private mvariFilialEmpresa As Integer
Private mvarsDetalhamento As String
Private mvarsResponsavel As String
Private mvardtDataInicio As Date
Private mvardtDataConclusao As Date
Private mvarcolCertificados As New Collection
Private mvarcolMOCursos As New Collection

Public Property Set colMOCursos(ByVal vData As Collection)
    Set mvarcolMOCursos = vData
End Property

Public Property Get colMOCursos() As Collection
    Set colMOCursos = mvarcolMOCursos
End Property

Public Property Set colCertificados(ByVal vData As Collection)
    Set mvarcolCertificados = vData
End Property

Public Property Get colCertificados() As Collection
    Set colCertificados = mvarcolCertificados
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

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let sDetalhamento(ByVal vData As String)
    mvarsDetalhamento = vData
End Property

Public Property Get sDetalhamento() As String
    sDetalhamento = mvarsDetalhamento
End Property

Public Property Let sResponsavel(ByVal vData As String)
    mvarsResponsavel = vData
End Property

Public Property Get sResponsavel() As String
    sResponsavel = mvarsResponsavel
End Property

Public Property Let dtDataInicio(ByVal vData As Date)
    mvardtDataInicio = vData
End Property

Public Property Get dtDataInicio() As Date
    dtDataInicio = mvardtDataInicio
End Property

Public Property Let dtDataConclusao(ByVal vData As Date)
    mvardtDataConclusao = vData
End Property

Public Property Get dtDataConclusao() As Date
    dtDataConclusao = mvardtDataConclusao
End Property

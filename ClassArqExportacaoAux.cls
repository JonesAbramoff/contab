VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassArqExportacaoAux"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntGeracao As Long
Private mvardtDataGeracao As Date
Private mvardHoraGeracao As Double
Private mvarsUsuario As String
Private mvariExportar As Integer
Private mvardtExpDataDe As Date
Private mvardtExpDataAte As Date
Private mvarcolTiposArq As New Collection
Private mvarsDiretorio As String

Public Property Let sDiretorio(ByVal vData As String)
    mvarsDiretorio = vData
End Property

Public Property Get sDiretorio() As String
    sDiretorio = mvarsDiretorio
End Property

Public Property Set colTiposArq(ByVal vData As Object)
    Set mvarcolTiposArq = vData
End Property

Public Property Get colTiposArq() As Collection
    Set colTiposArq = mvarcolTiposArq
End Property

Public Property Let lNumIntGeracao(ByVal vData As Long)
    mvarlNumIntGeracao = vData
End Property

Public Property Get lNumIntGeracao() As Long
    lNumIntGeracao = mvarlNumIntGeracao
End Property

Public Property Let dtDataGeracao(ByVal vData As Date)
    mvardtDataGeracao = vData
End Property

Public Property Get dtDataGeracao() As Date
    dtDataGeracao = mvardtDataGeracao
End Property

Public Property Let dHoraGeracao(ByVal vData As Double)
    mvardHoraGeracao = vData
End Property

Public Property Get dHoraGeracao() As Double
    dHoraGeracao = mvardHoraGeracao
End Property

Public Property Let sUsuario(ByVal vData As String)
    mvarsUsuario = vData
End Property

Public Property Get sUsuario() As String
    sUsuario = mvarsUsuario
End Property

Public Property Let iExportar(ByVal vData As Integer)
    mvariExportar = vData
End Property

Public Property Get iExportar() As Integer
    iExportar = mvariExportar
End Property

Public Property Let dtExpDataDe(ByVal vData As Date)
    mvardtExpDataDe = vData
End Property

Public Property Get dtExpDataDe() As Date
    dtExpDataDe = mvardtExpDataDe
End Property

Public Property Let dtExpDataAte(ByVal vData As Date)
    mvardtExpDataAte = vData
End Property

Public Property Get dtExpDataAte() As Date
    dtExpDataAte = mvardtExpDataAte
End Property

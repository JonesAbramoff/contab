VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassInvCliForn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDoc As Long
Private mvariFilialEmpresa As Integer
Private mvardtData As Date
Private mvariEscaninho As Integer
Private mvariTipoCliForn As Integer
Private mvarlCliForn As Long
Private mvariFilial As Integer
Private mvarsUsuario As String
Private mvardtDataGravacao As Date
Private mvardHoraGravacao As Double
Private mvarsObs As String
Private mvarcolItens As New Collection

Public Property Set colItens(ByVal vData As Object)
    Set mvarcolItens = vData
End Property

Public Property Get colItens() As Collection
    Set colItens = mvarcolItens
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let dtData(ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let iEscaninho(ByVal vData As Integer)
    mvariEscaninho = vData
End Property

Public Property Get iEscaninho() As Integer
    iEscaninho = mvariEscaninho
End Property

Public Property Let iTipoCliForn(ByVal vData As Integer)
    mvariTipoCliForn = vData
End Property

Public Property Get iTipoCliForn() As Integer
    iTipoCliForn = mvariTipoCliForn
End Property

Public Property Let lCliForn(ByVal vData As Long)
    mvarlCliForn = vData
End Property

Public Property Get lCliForn() As Long
    lCliForn = mvarlCliForn
End Property

Public Property Let iFilial(ByVal vData As Integer)
    mvariFilial = vData
End Property

Public Property Get iFilial() As Integer
    iFilial = mvariFilial
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let sUsuario(ByVal vData As String)
    mvarsUsuario = vData
End Property

Public Property Get sUsuario() As String
    sUsuario = mvarsUsuario
End Property

Public Property Let dtDataGravacao(ByVal vData As Date)
    mvardtDataGravacao = vData
End Property

Public Property Get dtDataGravacao() As Date
    dtDataGravacao = mvardtDataGravacao
End Property

Public Property Let dHoraGravacao(ByVal vData As Double)
    mvardHoraGravacao = vData
End Property

Public Property Get dHoraGravacao() As Double
    dHoraGravacao = mvardHoraGravacao
End Property

Public Property Let sObs(ByVal vData As String)
    mvarsObs = vData
End Property

Public Property Get sObs() As String
    sObs = mvarsObs
End Property

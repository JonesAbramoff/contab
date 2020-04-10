VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFluxoPRJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarcolFluxo As New Collection
Private mvardtDataDe As Date
Private mvardtDataAte As Date
Private mvaralComando(0 To 25) As Long
Private mvarobjProjeto As New ClassProjetos
Private mvarlNumIntRel As Long
Private mvariFilialEmpresa As Integer
Private mvarlNumIntEtapa As Long
Private mvarlNumIntPRJ As Long
Private mvarsDiretorio As String
Private mvarsNomeArquivo As String
Private mvarsModelo As String
Private mvariTipo As Integer

Public Property Let iTipo(ByVal vData As Integer)
    mvariTipo = vData
End Property

Public Property Get iTipo() As Integer
    iTipo = mvariTipo
End Property

Public Property Let sModelo(ByVal vData As String)
    mvarsModelo = vData
End Property

Public Property Get sModelo() As String
    sModelo = mvarsModelo
End Property

Public Property Let sNomeArquivo(ByVal vData As String)
    mvarsNomeArquivo = vData
End Property

Public Property Get sNomeArquivo() As String
    sNomeArquivo = mvarsNomeArquivo
End Property

Public Property Let sDiretorio(ByVal vData As String)
    mvarsDiretorio = vData
End Property

Public Property Get sDiretorio() As String
    sDiretorio = mvarsDiretorio
End Property

Public Property Let lNumIntPRJ(ByVal vData As Long)
    mvarlNumIntPRJ = vData
End Property

Public Property Get lNumIntPRJ() As Long
    lNumIntPRJ = mvarlNumIntPRJ
End Property

Public Property Let lNumIntEtapa(ByVal vData As Long)
    mvarlNumIntEtapa = vData
End Property

Public Property Get lNumIntEtapa() As Long
    lNumIntEtapa = mvarlNumIntEtapa
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let lNumIntRel(ByVal vData As Long)
    mvarlNumIntRel = vData
End Property

Public Property Get lNumIntRel() As Long
    lNumIntRel = mvarlNumIntRel
End Property

Public Property Get alComando(ByVal iIndice As Integer) As Long
    alComando = mvaralComando(iIndice)
End Property

Public Property Let alComando(ByVal iIndice As Integer, ByVal vData As Long)
    mvaralComando(iIndice) = vData
End Property

Public Property Let dtDataAte(ByVal vData As Date)
    mvardtDataAte = vData
End Property

Public Property Get dtDataAte() As Date
    dtDataAte = mvardtDataAte
End Property

Public Property Let dtDataDe(ByVal vData As Date)
    mvardtDataDe = vData
End Property

Public Property Get dtDataDe() As Date
    dtDataDe = mvardtDataDe
End Property

Public Property Set colFluxo(ByVal vData As Collection)
    Set mvarcolFluxo = vData
End Property

Public Property Get colFluxo() As Collection
    Set colFluxo = mvarcolFluxo
End Property

Public Property Set objProjeto(ByVal vData As ClassProjetos)
    Set mvarobjProjeto = vData
End Property

Public Property Get objProjeto() As ClassProjetos
    Set objProjeto = mvarobjProjeto
End Property


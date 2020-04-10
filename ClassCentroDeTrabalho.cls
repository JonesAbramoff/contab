VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCentrodeTrabalho"
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
Private mvarsNomeReduzido As String
Private mvarsDescricao As String
Private mvardCargaMin As Double
Private mvardCargaMax As Double
Private mvariTurnos As Integer
Private mvardHorasTurno As Double
Private mvarcolMaquinas As New Collection
Private mvarcolCompetencias As New Collection
Private mvariDiaisUteis(1 To 7) As Integer
Private mvarcolTurnos As New Collection
Private mvarcolOperadores As New Collection

Public Property Set colOperadores(ByVal vData As Collection)
    Set mvarcolOperadores = vData
End Property

Public Property Get colOperadores() As Collection
    Set colOperadores = mvarcolOperadores
End Property

Public Property Set colTurnos(ByVal vData As Collection)
    Set mvarcolTurnos = vData
End Property

Public Property Get colTurnos() As Collection
    Set colTurnos = mvarcolTurnos
End Property

Public Property Let iDiaisUteis(ByVal iIndice As Integer, ByVal vData As Integer)
    mvariDiaisUteis(iIndice) = vData
End Property

Public Property Get iDiaisUteis(ByVal iIndice As Integer) As Integer
    iDiaisUteis = mvariDiaisUteis(iIndice)
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

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let dCargaMin(ByVal vData As Double)
    mvardCargaMin = vData
End Property

Public Property Get dCargaMin() As Double
    dCargaMin = mvardCargaMin
End Property

Public Property Let dCargaMax(ByVal vData As Double)
    mvardCargaMax = vData
End Property

Public Property Get dCargaMax() As Double
    dCargaMax = mvardCargaMax
End Property

Public Property Let iTurnos(ByVal vData As Integer)
    mvariTurnos = vData
End Property

Public Property Get iTurnos() As Integer
    iTurnos = mvariTurnos
End Property

Public Property Let dHorasTurno(ByVal vData As Double)
    mvardHorasTurno = vData
End Property

Public Property Get dHorasTurno() As Double
    dHorasTurno = mvardHorasTurno
End Property

Public Property Let sNomeReduzido(ByVal vData As String)
    mvarsNomeReduzido = vData
End Property

Public Property Get sNomeReduzido() As String
    sNomeReduzido = mvarsNomeReduzido
End Property

Public Property Set colMaquinas(ByVal vData As Collection)
    Set mvarcolMaquinas = vData
End Property

Public Property Get colMaquinas() As Collection
    Set colMaquinas = mvarcolMaquinas
End Property

Public Property Set colCompetencias(ByVal vData As Collection)
    Set mvarcolCompetencias = vData
End Property

Public Property Get colCompetencias() As Collection
    Set colCompetencias = mvarcolCompetencias
End Property


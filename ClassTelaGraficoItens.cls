VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTelaGraficoItens"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvariQtdDias As Integer
Private mvarsTextoExibicao As String
Private mvarlCor As Long
Private mvariIndiceCor As Integer
Private mvarsNomeTela As String
Private mvarcolobj As New Collection
Private mvardtDataInicio As Date
Private mvardtDataFim As Date
Private mvarsNome As String
Private mvariIcone As Integer
Private mvarcolPredecessores As New Collection
Private mvarobjControle As Object
Private mvarsDescricao As String

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Set objControle(ByVal vData As Object)
    Set mvarobjControle = vData
End Property

Public Property Get objControle() As Object
    Set objControle = mvarobjControle
End Property

Public Property Set colPredecessores(ByVal vData As Collection)
    Set mvarcolPredecessores = vData
End Property

Public Property Get colPredecessores() As Collection
    Set colPredecessores = mvarcolPredecessores
End Property

Public Property Let iIcone(ByVal vData As Integer)
    mvariIcone = vData
End Property

Public Property Get iIcone() As Integer
    iIcone = mvariIcone
End Property

Public Property Let dtDataInicio(ByVal vData As Date)
    mvardtDataInicio = vData
End Property

Public Property Get dtDataInicio() As Date
    dtDataInicio = mvardtDataInicio
End Property

Public Property Let dtDataFim(ByVal vData As Date)
    mvardtDataFim = vData
End Property

Public Property Get dtDataFim() As Date
    dtDataFim = mvardtDataFim
End Property

Public Property Set colobj(ByVal vData As Collection)
    Set mvarcolobj = vData
End Property

Public Property Get colobj() As Collection
    Set colobj = mvarcolobj
End Property

Public Property Let iQtdDias(ByVal vData As Integer)
    mvariQtdDias = vData
End Property

Public Property Get iQtdDias() As Integer
    iQtdDias = mvariQtdDias
End Property

Public Property Let iIndiceCor(ByVal vData As Integer)
    mvariIndiceCor = vData
End Property

Public Property Get iIndiceCor() As Integer
    iIndiceCor = mvariIndiceCor
End Property

Public Property Let sTextoExibicao(ByVal vData As String)
    mvarsTextoExibicao = vData
End Property

Public Property Get sTextoExibicao() As String
    sTextoExibicao = mvarsTextoExibicao
End Property

Public Property Let lCor(ByVal vData As Long)
    mvarlCor = vData
End Property

Public Property Get lCor() As Long
    lCor = mvarlCor
End Property

Public Property Let sNomeTela(ByVal vData As String)
    mvarsNomeTela = vData
End Property

Public Property Get sNomeTela() As String
    sNomeTela = mvarsNomeTela
End Property

Public Property Let sNome(ByVal vData As String)
    mvarsNome = vData
End Property

Public Property Get sNome() As String
    sNome = mvarsNome
End Property

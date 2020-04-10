VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRotasPontos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDoc As Long
Private mvarlNumIntDocRota As Long
Private mvariSeq As Integer
Private mvarlCliente As Long
Private mvarsObservacao As String
Private mvarlMeio As Long
Private mvardTempo As Double
Private mvardDistancia As Double
Private mvariSelecionado As Integer
Private mvarobjPontoAnt As Object
Private mvarobjCliente As New ClassCliente
Private mvarobjEndereco As New ClassEndereco
Private mvariFilialCliente As Integer

Public Property Let iFilialCliente(ByVal vData As Integer)
    mvariFilialCliente = vData
End Property

Public Property Get iFilialCliente() As Integer
    iFilialCliente = mvariFilialCliente
End Property

Public Property Set objEndereco(ByVal vData As Object)
    Set mvarobjEndereco = vData
End Property

Public Property Get objEndereco() As ClassEndereco
    Set objEndereco = mvarobjEndereco
End Property

Public Property Set objCliente(ByVal vData As Object)
    Set mvarobjCliente = vData
End Property

Public Property Get objCliente() As ClassCliente
    Set objCliente = mvarobjCliente
End Property

Public Property Set objPontoAnt(ByVal vData As Object)
    Set mvarobjPontoAnt = vData
End Property

Public Property Get objPontoAnt() As Object
    Set objPontoAnt = mvarobjPontoAnt
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let lNumIntDocRota(ByVal vData As Long)
    mvarlNumIntDocRota = vData
End Property

Public Property Get lNumIntDocRota() As Long
    lNumIntDocRota = mvarlNumIntDocRota
End Property

Public Property Let iSeq(ByVal vData As Integer)
    mvariSeq = vData
End Property

Public Property Get iSeq() As Integer
    iSeq = mvariSeq
End Property

Public Property Let lCliente(ByVal vData As Long)
    mvarlCliente = vData
End Property

Public Property Get lCliente() As Long
    lCliente = mvarlCliente
End Property

Public Property Let sObservacao(ByVal vData As String)
    mvarsObservacao = vData
End Property

Public Property Get sObservacao() As String
    sObservacao = mvarsObservacao
End Property

Public Property Let lMeio(ByVal vData As Long)
    mvarlMeio = vData
End Property

Public Property Get lMeio() As Long
    lMeio = mvarlMeio
End Property

Public Property Let dTempo(ByVal vData As Double)
    mvardTempo = vData
End Property

Public Property Get dTempo() As Double
    dTempo = mvardTempo
End Property

Public Property Let dDistancia(ByVal vData As Double)
    mvardDistancia = vData
End Property

Public Property Get dDistancia() As Double
    dDistancia = mvardDistancia
End Property

Public Property Let iSelecionado(ByVal vData As Integer)
    mvariSelecionado = vData
End Property

Public Property Get iSelecionado() As Integer
    iSelecionado = mvariSelecionado
End Property

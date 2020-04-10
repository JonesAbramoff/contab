VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassMapaDeEntrega"
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
Private mvarlCodigo As Long
Private mvardtData As Date
Private mvariRegiao As Integer
Private mvarlVeiculo As Long
Private mvardVolumeTotal As Double
Private mvardPesoTotal As Double
Private mvariNumViagens As Integer
Private mvardHoraSaida As Double
Private mvardHoraRetorno As Double
Private mvarsResponsavel As String
Private mvarcolDocs As New Collection
Private mvarcolMapaDoc As New Collection
Private mvariTransportadora As Integer
Private mvariTipoDoc As Integer

Public Property Let iTransportadora(ByVal vData As Integer)
    mvariTransportadora = vData
End Property

Public Property Get iTransportadora() As Integer
    iTransportadora = mvariTransportadora
End Property

Public Property Let iTipoDoc(ByVal vData As Integer)
    mvariTipoDoc = vData
End Property

Public Property Get iTipoDoc() As Integer
    iTipoDoc = mvariTipoDoc
End Property

Public Property Set colMapaDoc(ByVal vData As Object)
    Set mvarcolMapaDoc = vData
End Property

Public Property Get colMapaDoc() As Collection
    Set colMapaDoc = mvarcolMapaDoc
End Property

Public Property Set colDocs(ByVal vData As Object)
    Set mvarcolDocs = vData
End Property

Public Property Get colDocs() As Collection
    Set colDocs = mvarcolDocs
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

Public Property Let lCodigo(ByVal vData As Long)
    mvarlCodigo = vData
End Property

Public Property Get lCodigo() As Long
    lCodigo = mvarlCodigo
End Property

Public Property Let dtData(ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let iRegiao(ByVal vData As Integer)
    mvariRegiao = vData
End Property

Public Property Get iRegiao() As Integer
    iRegiao = mvariRegiao
End Property

Public Property Let lVeiculo(ByVal vData As Long)
    mvarlVeiculo = vData
End Property

Public Property Get lVeiculo() As Long
    lVeiculo = mvarlVeiculo
End Property

Public Property Let dVolumeTotal(ByVal vData As Double)
    mvardVolumeTotal = vData
End Property

Public Property Get dVolumeTotal() As Double
    dVolumeTotal = mvardVolumeTotal
End Property

Public Property Let dPesoTotal(ByVal vData As Double)
    mvardPesoTotal = vData
End Property

Public Property Get dPesoTotal() As Double
    dPesoTotal = mvardPesoTotal
End Property

Public Property Let iNumViagens(ByVal vData As Integer)
    mvariNumViagens = vData
End Property

Public Property Get iNumViagens() As Integer
    iNumViagens = mvariNumViagens
End Property

Public Property Let dHoraSaida(ByVal vData As Double)
    mvardHoraSaida = vData
End Property

Public Property Get dHoraSaida() As Double
    dHoraSaida = mvardHoraSaida
End Property

Public Property Let dHoraRetorno(ByVal vData As Double)
    mvardHoraRetorno = vData
End Property

Public Property Get dHoraRetorno() As Double
    dHoraRetorno = mvardHoraRetorno
End Property

Public Property Let sResponsavel(ByVal vData As String)
    mvarsResponsavel = vData
End Property

Public Property Get sResponsavel() As String
    sResponsavel = mvarsResponsavel
End Property

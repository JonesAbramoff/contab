VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTransportadora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

Private mvariCodigo As Integer
Private mvarsNome As String
Private mvarsNomeReduzido As String
Private mvarlEndereco As Long
Private mvarsCGC As String
Private mvarsInscricaoEstadual As String
Private mvariViaTransporte As Integer
Private mvarsInscricaoMunicipal As String
Private mvarsGuia As String
Private mvardPesoMinimo As Double
Private mvarsObservacao As String
Private mvariIEIsento As Integer
Private mvariIENaoContrib As Integer

Public Property Let iIENaoContrib(ByVal vData As Integer)
    mvariIENaoContrib = vData
End Property

Public Property Get iIENaoContrib() As Integer
    If mvariIENaoContrib = -1 Then
        iIENaoContrib = IIf(Len(Trim(mvarsInscricaoEstadual)) = 0, MARCADO, DESMARCADO)
    Else
        iIENaoContrib = mvariIENaoContrib
    End If
End Property

Public Property Let iIEIsento(ByVal vData As Integer)
    mvariIEIsento = vData
End Property

Public Property Get iIEIsento() As Integer
    If mvariIEIsento = -1 Then
        iIEIsento = IIf(Len(Trim(mvarsInscricaoEstadual)) = 0, MARCADO, DESMARCADO)
    Else
        iIEIsento = mvariIEIsento
    End If
End Property

Public Property Let sObservacao(ByVal vData As String)
    mvarsObservacao = vData
End Property

Public Property Get sObservacao() As String
    sObservacao = mvarsObservacao
End Property

Public Property Let dPesoMinimo(ByVal vData As Double)
    mvardPesoMinimo = vData
End Property

Public Property Get dPesoMinimo() As Double
    dPesoMinimo = mvardPesoMinimo
End Property

Public Property Let sGuia(ByVal vData As String)
    mvarsGuia = vData
End Property

Public Property Get sGuia() As String
    sGuia = mvarsGuia
End Property

Public Property Let sInscricaoMunicipal(ByVal vData As String)
    mvarsInscricaoMunicipal = vData
End Property

Public Property Get sInscricaoMunicipal() As String
    sInscricaoMunicipal = mvarsInscricaoMunicipal
End Property

Public Property Let iViaTransporte(ByVal vData As Integer)
    mvariViaTransporte = vData
End Property

Public Property Get iViaTransporte() As Integer
    iViaTransporte = mvariViaTransporte
End Property

Public Property Let lEndereco(ByVal vData As Long)
    mvarlEndereco = vData
End Property

Public Property Get lEndereco() As Long
    lEndereco = mvarlEndereco
End Property

Public Property Let sNomeReduzido(ByVal vData As String)
    mvarsNomeReduzido = vData
End Property

Public Property Get sNomeReduzido() As String
    sNomeReduzido = mvarsNomeReduzido
End Property

Public Property Let sNome(ByVal vData As String)
    mvarsNome = vData
End Property

Public Property Get sNome() As String
    sNome = mvarsNome
End Property

Public Property Let iCodigo(ByVal vData As Integer)
    mvariCodigo = vData
End Property

Public Property Get iCodigo() As Integer
    iCodigo = mvariCodigo
End Property

Public Property Let sCgc(ByVal vData As String)
    mvarsCGC = vData
End Property

Public Property Get sCgc() As String
    sCgc = mvarsCGC
End Property

Public Property Let sInscricaoEstadual(ByVal vData As String)
    mvarsInscricaoEstadual = vData
End Property

Public Property Get sInscricaoEstadual() As String
    sInscricaoEstadual = mvarsInscricaoEstadual
End Property

Private Sub Class_Initialize()
    mvariIEIsento = -1
    mvariIENaoContrib = -1

End Sub

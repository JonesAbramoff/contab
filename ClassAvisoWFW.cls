VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassAvisoWFW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarlNumIntDoc As Long
Private mvarsMsg As String
Private mvardtData As Date
Private mvardHora As Double
Private mvarsUsuario As String
Private mvariTransacao As Integer
Private mvardtDataUltAviso As Date
Private mvardHoraUltAviso As Double
Private mvardIntervalo As Double
Private mvarsUsuarioOrig As String
Private mvarsTransacaoTela As String
Private mvariUMIntervalo As Integer

Public Property Let iUMIntervalo(ByVal vData As Integer)
    mvariUMIntervalo = vData
End Property

Public Property Get iUMIntervalo() As Integer
    iUMIntervalo = mvariUMIntervalo
End Property

Public Property Let sTransacaoTela(ByVal vData As String)
    mvarsTransacaoTela = vData
End Property

Public Property Get sTransacaoTela() As String
    sTransacaoTela = mvarsTransacaoTela
End Property

Public Property Let sUsuarioOrig(ByVal vData As String)
    mvarsUsuarioOrig = vData
End Property

Public Property Get sUsuarioOrig() As String
    sUsuarioOrig = mvarsUsuarioOrig
End Property

Public Property Let dIntervalo(ByVal vData As Double)
    mvardIntervalo = vData
End Property

Public Property Get dIntervalo() As Double
    dIntervalo = mvardIntervalo
End Property

Public Property Let dHoraUltAviso(ByVal vData As Double)
    mvardHoraUltAviso = vData
End Property

Public Property Get dHoraUltAviso() As Double
    dHoraUltAviso = mvardHoraUltAviso
End Property

Public Property Let dtDataUltAviso(ByVal vData As Date)
    mvardtDataUltAviso = vData
End Property

Public Property Get dtDataUltAviso() As Date
    dtDataUltAviso = mvardtDataUltAviso
End Property

Public Property Let iTransacao(ByVal vData As Integer)
    mvariTransacao = vData
End Property

Public Property Get iTransacao() As Integer
    iTransacao = mvariTransacao
End Property

Public Property Let sUsuario(ByVal vData As String)
    mvarsUsuario = vData
End Property

Public Property Get sUsuario() As String
    sUsuario = mvarsUsuario
End Property

Public Property Let dHora(ByVal vData As Double)
    mvardHora = vData
End Property

Public Property Get dHora() As Double
    dHora = mvardHora
End Property

Public Property Let dtData(ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let sMsg(ByVal vData As String)
    mvarsMsg = vData
End Property

Public Property Get sMsg() As String
    sMsg = mvarsMsg
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property






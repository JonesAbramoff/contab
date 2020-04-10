VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTRVOcrCasosSrv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDocOcrCaso As Long
Private mvariSeq As Integer
Private mvariAutorizado As Integer
Private mvariSolicitado As Integer
Private mvarsDescricao As String
Private mvardValorLimite As Double
Private mvariMoeda As Integer
Private mvariTipo As Integer
Private mvardValorSolicitadoRS As Double
Private mvardValorSolicitadoUS As Double
Private mvardValorAutorizadoRS As Double
Private mvardValorAutorizadoUS As Double
Private mvarlCodigoServ As Long

Public Property Let lCodigoServ(ByVal vData As Long)
    mvarlCodigoServ = vData
End Property

Public Property Get lCodigoServ() As Long
    lCodigoServ = mvarlCodigoServ
End Property

Public Property Let lNumIntDocOcrCaso(ByVal vData As Long)
    mvarlNumIntDocOcrCaso = vData
End Property

Public Property Get lNumIntDocOcrCaso() As Long
    lNumIntDocOcrCaso = mvarlNumIntDocOcrCaso
End Property

Public Property Let iSeq(ByVal vData As Integer)
    mvariSeq = vData
End Property

Public Property Get iSeq() As Integer
    iSeq = mvariSeq
End Property

Public Property Let iSolicitado(ByVal vData As Integer)
    mvariSolicitado = vData
End Property

Public Property Get iSolicitado() As Integer
    iSolicitado = mvariSolicitado
End Property

Public Property Let iAutorizado(ByVal vData As Integer)
    mvariAutorizado = vData
End Property

Public Property Get iAutorizado() As Integer
    iAutorizado = mvariAutorizado
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let dValorLimite(ByVal vData As Double)
    mvardValorLimite = vData
End Property

Public Property Get dValorLimite() As Double
    dValorLimite = mvardValorLimite
End Property

Public Property Let iMoeda(ByVal vData As Integer)
    mvariMoeda = vData
End Property

Public Property Get iMoeda() As Integer
    iMoeda = mvariMoeda
End Property

Public Property Let iTipo(ByVal vData As Integer)
    mvariTipo = vData
End Property

Public Property Get iTipo() As Integer
    iTipo = mvariTipo
End Property

Public Property Let dValorSolicitadoRS(ByVal vData As Double)
    mvardValorSolicitadoRS = vData
End Property

Public Property Get dValorSolicitadoRS() As Double
    dValorSolicitadoRS = mvardValorSolicitadoRS
End Property

Public Property Let dValorSolicitadoUS(ByVal vData As Double)
    mvardValorSolicitadoUS = vData
End Property

Public Property Get dValorSolicitadoUS() As Double
    dValorSolicitadoUS = mvardValorSolicitadoUS
End Property

Public Property Let dValorAutorizadoRS(ByVal vData As Double)
    mvardValorAutorizadoRS = vData
End Property

Public Property Get dValorAutorizadoRS() As Double
    dValorAutorizadoRS = mvardValorAutorizadoRS
End Property

Public Property Let dValorAutorizadoUS(ByVal vData As Double)
    mvardValorAutorizadoUS = vData
End Property

Public Property Get dValorAutorizadoUS() As Double
    dValorAutorizadoUS = mvardValorAutorizadoUS
End Property

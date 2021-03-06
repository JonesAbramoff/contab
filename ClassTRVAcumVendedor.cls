VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTRVAcumVendedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvardValorAcum As Double
Private mvardValorDolarAcum As Double
Private mvariVendedor As Integer
Private mvariSuperior As Integer
Private mvarlNumIntInfoGer As Long
Private mvariCargo As Integer
Private mvarlNumIntCom As Long
Private mvarcolAcumAporte As New Collection
Private mvardCambioMedio As Double
Private mvarlNumIntRelComiInt As Long
Private mvarlNumIntInfoGerGrava As Long
Private mvarcolTRVAcumAux As New Collection
Private mvarlNumIntRelComiIntGrava As Long
Private mvardValorAcumCallCenter As Double
Private mvardPercCallCenter As Double
Private mvarlNumIntRelComiIntCallCenter As Long
Private mvarlNumIntInfoGerCallCenter As Long
Private mvariAporte As Integer  'serve para indicar se um determinado voucher esta associado a aporte
Private mvarsNomeReduzido As String
Private mvardValorAcumMyAssistance As Double
Private mvarlNumIntRelComiIntMyAssistance As Long
Private mvarlNumIntInfoGerMyAssistance As Long

Public Property Let lNumIntInfoGerMyAssistance(ByVal vData As Long)
    mvarlNumIntInfoGerMyAssistance = vData
End Property

Public Property Get lNumIntInfoGerMyAssistance() As Long
    lNumIntInfoGerMyAssistance = mvarlNumIntInfoGerMyAssistance
End Property

Public Property Let lNumIntRelComiIntMyAssistance(ByVal vData As Long)
    mvarlNumIntRelComiIntMyAssistance = vData
End Property

Public Property Get lNumIntRelComiIntMyAssistance() As Long
    lNumIntRelComiIntMyAssistance = mvarlNumIntRelComiIntMyAssistance
End Property

Public Property Let dValorAcumMyAssistance(ByVal vData As Double)
    mvardValorAcumMyAssistance = vData
End Property

Public Property Get dValorAcumMyAssistance() As Double
    dValorAcumMyAssistance = mvardValorAcumMyAssistance
End Property

Public Property Let sNomeReduzido(ByVal vData As String)
    mvarsNomeReduzido = vData
End Property

Public Property Get sNomeReduzido() As String
    sNomeReduzido = mvarsNomeReduzido
End Property

Public Property Let iAporte(ByVal vData As Integer)
    mvariAporte = vData
End Property

Public Property Get iAporte() As Integer
    iAporte = mvariAporte
End Property

Public Property Let lNumIntInfoGerCallCenter(ByVal vData As Long)
    mvarlNumIntInfoGerCallCenter = vData
End Property

Public Property Get lNumIntInfoGerCallCenter() As Long
    lNumIntInfoGerCallCenter = mvarlNumIntInfoGerCallCenter
End Property


Public Property Let lNumIntRelComiIntCallCenter(ByVal vData As Long)
    mvarlNumIntRelComiIntCallCenter = vData
End Property

Public Property Get lNumIntRelComiIntCallCenter() As Long
    lNumIntRelComiIntCallCenter = mvarlNumIntRelComiIntCallCenter
End Property

Public Property Let dPercCallCenter(ByVal vData As Double)
    mvardPercCallCenter = vData
End Property

Public Property Get dPercCallCenter() As Double
    dPercCallCenter = mvardPercCallCenter
End Property

Public Property Let dValorAcumCallCenter(ByVal vData As Double)
    mvardValorAcumCallCenter = vData
End Property

Public Property Get dValorAcumCallCenter() As Double
    dValorAcumCallCenter = mvardValorAcumCallCenter
End Property

Public Property Let lNumIntRelComiIntGrava(ByVal vData As Long)
    mvarlNumIntRelComiIntGrava = vData
End Property

Public Property Get lNumIntRelComiIntGrava() As Long
    lNumIntRelComiIntGrava = mvarlNumIntRelComiIntGrava
End Property

Public Property Set colTRVAcumAux(ByVal vData As Collection)
    Set mvarcolTRVAcumAux = vData
End Property

Public Property Get colTRVAcumAux() As Collection
    Set colTRVAcumAux = mvarcolTRVAcumAux
End Property

Public Property Let lNumIntInfoGerGrava(ByVal vData As Long)
    mvarlNumIntInfoGerGrava = vData
End Property

Public Property Get lNumIntInfoGerGrava() As Long
    lNumIntInfoGerGrava = mvarlNumIntInfoGerGrava
End Property

Public Property Let lNumIntRelComiInt(ByVal vData As Long)
    mvarlNumIntRelComiInt = vData
End Property

Public Property Get lNumIntRelComiInt() As Long
    lNumIntRelComiInt = mvarlNumIntRelComiInt
End Property

Public Property Let dCambioMedio(ByVal vData As Double)
    mvardCambioMedio = vData
End Property

Public Property Get dCambioMedio() As Double
    dCambioMedio = mvardCambioMedio
End Property

Public Property Set colAcumAporte(ByVal vData As Collection)
    Set mvarcolAcumAporte = vData
End Property

Public Property Get colAcumAporte() As Collection
    Set colAcumAporte = mvarcolAcumAporte
End Property

Public Property Let lNumIntCom(ByVal vData As Long)
    mvarlNumIntCom = vData
End Property

Public Property Get lNumIntCom() As Long
    lNumIntCom = mvarlNumIntCom
End Property

Public Property Let iCargo(ByVal vData As Integer)
    mvariCargo = vData
End Property

Public Property Get iCargo() As Integer
    iCargo = mvariCargo
End Property

Public Property Let lNumIntInfoGer(ByVal vData As Long)
    mvarlNumIntInfoGer = vData
End Property

Public Property Get lNumIntInfoGer() As Long
    lNumIntInfoGer = mvarlNumIntInfoGer
End Property

Public Property Let iSuperior(ByVal vData As Integer)
    mvariSuperior = vData
End Property

Public Property Get iSuperior() As Integer
    iSuperior = mvariSuperior
End Property

Public Property Let iVendedor(ByVal vData As Integer)
    mvariVendedor = vData
End Property

Public Property Get iVendedor() As Integer
    iVendedor = mvariVendedor
End Property

Public Property Let dValorDolarAcum(ByVal vData As Double)
    mvardValorDolarAcum = vData
End Property

Public Property Get dValorDolarAcum() As Double
    dValorDolarAcum = mvardValorDolarAcum
End Property

Public Property Let dValorAcum(ByVal vData As Double)
    mvardValorAcum = vData
End Property

Public Property Get dValorAcum() As Double
    dValorAcum = mvardValorAcum
End Property



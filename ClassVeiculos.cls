VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassVeiculos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlCodigo As Long
Private mvarsDescricao As String
Private mvarlTipo As Long
Private mvariProprio As Integer
Private mvarsPlaca As String
Private mvarsPlacaUF As String
Private mvardCapacidadeKg As Double
Private mvardVolumeM3 As Double
Private mvardCustoHora As Double
Private mvardDispPadraoDe As Double
Private mvardDispPadraoAte As Double

Public Property Let lCodigo(ByVal vData As Long)
    mvarlCodigo = vData
End Property

Public Property Get lCodigo() As Long
    lCodigo = mvarlCodigo
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let lTipo(ByVal vData As Long)
    mvarlTipo = vData
End Property

Public Property Get lTipo() As Long
    lTipo = mvarlTipo
End Property

Public Property Let iProprio(ByVal vData As Integer)
    mvariProprio = vData
End Property

Public Property Get iProprio() As Integer
    iProprio = mvariProprio
End Property

Public Property Let sPlaca(ByVal vData As String)
    mvarsPlaca = vData
End Property

Public Property Get sPlaca() As String
    sPlaca = mvarsPlaca
End Property

Public Property Let sPlacaUF(ByVal vData As String)
    mvarsPlacaUF = vData
End Property

Public Property Get sPlacaUF() As String
    sPlacaUF = mvarsPlacaUF
End Property

Public Property Let dCapacidadeKg(ByVal vData As Double)
    mvardCapacidadeKg = vData
End Property

Public Property Get dCapacidadeKg() As Double
    dCapacidadeKg = mvardCapacidadeKg
End Property

Public Property Let dVolumeM3(ByVal vData As Double)
    mvardVolumeM3 = vData
End Property

Public Property Get dVolumeM3() As Double
    dVolumeM3 = mvardVolumeM3
End Property

Public Property Let dCustoHora(ByVal vData As Double)
    mvardCustoHora = vData
End Property

Public Property Get dCustoHora() As Double
    dCustoHora = mvardCustoHora
End Property

Public Property Let dDispPadraoDe(ByVal vData As Double)
    mvardDispPadraoDe = vData
End Property

Public Property Get dDispPadraoDe() As Double
    dDispPadraoDe = mvardDispPadraoDe
End Property

Public Property Let dDispPadraoAte(ByVal vData As Double)
    mvardDispPadraoAte = vData
End Property

Public Property Get dDispPadraoAte() As Double
    dDispPadraoAte = mvardDispPadraoAte
End Property

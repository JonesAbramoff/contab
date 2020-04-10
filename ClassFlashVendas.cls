VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFlashVendas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarsProduto As String
Private mvarsUMVenda As String
Private mvariFilialEmpresa As Integer
Private mvardtData As Date
Private mvardHoraIni As Double
Private mvardHoraFim As Double
Private mvariCaixaIni As Integer
Private mvariCaixaFim As Integer
Private mvarlNumVendas As Long
Private mvardQuantVendida As Double
Private mvardValorVendido As Double

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let sUMVenda(ByVal vData As String)
    mvarsUMVenda = vData
End Property

Public Property Get sUMVenda() As String
    sUMVenda = mvarsUMVenda
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let dtData(ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let dHoraIni(ByVal vData As Double)
    mvardHoraIni = vData
End Property

Public Property Get dHoraIni() As Double
    dHoraIni = mvardHoraIni
End Property

Public Property Let dHoraFim(ByVal vData As Double)
    mvardHoraFim = vData
End Property

Public Property Get dHoraFim() As Double
    dHoraFim = mvardHoraFim
End Property

Public Property Let iCaixaIni(ByVal vData As Integer)
    mvariCaixaIni = vData
End Property

Public Property Get iCaixaIni() As Integer
    iCaixaIni = mvariCaixaIni
End Property

Public Property Let iCaixaFim(ByVal vData As Integer)
    mvariCaixaFim = vData
End Property

Public Property Get iCaixaFim() As Integer
    iCaixaFim = mvariCaixaFim
End Property

Public Property Let lNumVendas(ByVal vData As Long)
    mvarlNumVendas = vData
End Property

Public Property Get lNumVendas() As Long
    lNumVendas = mvarlNumVendas
End Property

Public Property Let dQuantVendida(ByVal vData As Double)
    mvardQuantVendida = vData
End Property

Public Property Get dQuantVendida() As Double
    dQuantVendida = mvardQuantVendida
End Property

Public Property Let dValorVendido(ByVal vData As Double)
    mvardValorVendido = vData
End Property

Public Property Get dValorVendido() As Double
    dValorVendido = mvardValorVendido
End Property


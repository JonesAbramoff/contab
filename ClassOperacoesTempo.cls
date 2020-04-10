VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassOperacoesTempo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlNumIntDocOperacao As Long
Private mvardLoteMax As Double
Private mvardLoteMin As Double
Private mvardLotePadrao As Double
Private mvardTempoPreparacao As Double
Private mvardTempoMovimentacao As Double
Private mvardTempoDescarga As Double
Private mvardTempoOperacao As Double
Private mvarsUMTempo As String
Private mvariTipo As Integer
Private mvarlNumIntDocMaq As Long

Public Property Let lNumIntDocMaq(ByVal vData As Long)
    mvarlNumIntDocMaq = vData
End Property

Public Property Get lNumIntDocMaq() As Long
    lNumIntDocMaq = mvarlNumIntDocMaq
End Property

Public Property Let iTipo(ByVal vData As Integer)
    mvariTipo = vData
End Property

Public Property Get iTipo() As Integer
    iTipo = mvariTipo
End Property

Public Property Let sUMTempo(ByVal vData As String)
    mvarsUMTempo = vData
End Property

Public Property Get sUMTempo() As String
    sUMTempo = mvarsUMTempo
End Property

Public Property Let dTempoOperacao(ByVal vData As Double)
    mvardTempoOperacao = vData
End Property

Public Property Get dTempoOperacao() As Double
    dTempoOperacao = mvardTempoOperacao
End Property

Public Property Let dTempoDescarga(ByVal vData As Double)
    mvardTempoDescarga = vData
End Property

Public Property Get dTempoDescarga() As Double
    dTempoDescarga = mvardTempoDescarga
End Property

Public Property Let dTempoMovimentacao(ByVal vData As Double)
    mvardTempoMovimentacao = vData
End Property

Public Property Get dTempoMovimentacao() As Double
    dTempoMovimentacao = mvardTempoMovimentacao
End Property

Public Property Let dTempoPreparacao(ByVal vData As Double)
    mvardTempoPreparacao = vData
End Property

Public Property Get dTempoPreparacao() As Double
    dTempoPreparacao = mvardTempoPreparacao
End Property

Public Property Let dLotePadrao(ByVal vData As Double)
    mvardLotePadrao = vData
End Property

Public Property Get dLotePadrao() As Double
    dLotePadrao = mvardLotePadrao
End Property

Public Property Let dLoteMin(ByVal vData As Double)
    mvardLoteMin = vData
End Property

Public Property Get dLoteMin() As Double
    dLoteMin = mvardLoteMin
End Property

Public Property Let dLoteMax(ByVal vData As Double)
    mvardLoteMax = vData
End Property

Public Property Get dLoteMax() As Double
    dLoteMax = mvardLoteMax
End Property

Public Property Let lNumIntDocOperacao(ByVal vData As Long)
    mvarlNumIntDocOperacao = vData
End Property

Public Property Get lNumIntDocOperacao() As Long
    lNumIntDocOperacao = mvarlNumIntDocOperacao
End Property


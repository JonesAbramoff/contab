VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassVNDReducaoTRP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarlNumIntDoc As Long
Private mvariVendedor As Integer
Private mvariSeq As Integer
Private mvardValorDe As Double
Private mvardValorAte As Double
Private mvariMoeda As Integer
Private mvardPercComissaoMax As Double

Public Property Let dPercComissaoMax(ByVal vData As Double)
    mvardPercComissaoMax = vData
End Property

Public Property Get dPercComissaoMax() As Double
    dPercComissaoMax = mvardPercComissaoMax
End Property

Public Property Let iMoeda(ByVal vData As Integer)
    mvariMoeda = vData
End Property

Public Property Get iMoeda() As Integer
    iMoeda = mvariMoeda
End Property

Public Property Let dValorAte(ByVal vData As Double)
    mvardValorAte = vData
End Property

Public Property Get dValorAte() As Double
    dValorAte = mvardValorAte
End Property

Public Property Let dValorDe(ByVal vData As Double)
    mvardValorDe = vData
End Property

Public Property Get dValorDe() As Double
    dValorDe = mvardValorDe
End Property

Public Property Let iSeq(ByVal vData As Integer)
    mvariSeq = vData
End Property

Public Property Get iSeq() As Integer
    iSeq = mvariSeq
End Property

Public Property Let iVendedor(ByVal vData As Integer)
    mvariVendedor = vData
End Property

Public Property Get iVendedor() As Integer
    iVendedor = mvariVendedor
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property




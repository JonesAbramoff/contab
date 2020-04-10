VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassOVAcomp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvardtDataProxDe As Date
Private mvardtDataProxAte As Date
Private mvardtDataPrevDe As Date
Private mvardtDataPrevAte As Date
Private mvariSoEmPV As Integer
Private mvariVendedor As Integer
Private mvariSoNaoFaturado As Integer
Private mvariSoNaoPerdido As Integer
Private mvardtDataEmiDe As Date
Private mvardtDataEmiAte As Date

Public Property Let dtDataEmiAte(ByVal vData As Date)
    mvardtDataEmiAte = vData
End Property

Public Property Get dtDataEmiAte() As Date
    dtDataEmiAte = mvardtDataEmiAte
End Property

Public Property Let dtDataEmiDe(ByVal vData As Date)
    mvardtDataEmiDe = vData
End Property

Public Property Get dtDataEmiDe() As Date
    dtDataEmiDe = mvardtDataEmiDe
End Property

Public Property Let iSoNaoFaturado(ByVal vData As Integer)
    mvariSoNaoFaturado = vData
End Property

Public Property Get iSoNaoFaturado() As Integer
    iSoNaoFaturado = mvariSoNaoFaturado
End Property

Public Property Let iSoNaoPerdido(ByVal vData As Integer)
    mvariSoNaoPerdido = vData
End Property

Public Property Get iSoNaoPerdido() As Integer
    iSoNaoPerdido = mvariSoNaoPerdido
End Property

Public Property Let iVendedor(ByVal vData As Integer)
    mvariVendedor = vData
End Property

Public Property Get iVendedor() As Integer
    iVendedor = mvariVendedor
End Property

Public Property Let iSoEmPV(ByVal vData As Integer)
    mvariSoEmPV = vData
End Property

Public Property Get iSoEmPV() As Integer
    iSoEmPV = mvariSoEmPV
End Property

Public Property Let dtDataPrevAte(ByVal vData As Date)
    mvardtDataPrevAte = vData
End Property

Public Property Get dtDataPrevAte() As Date
    dtDataPrevAte = mvardtDataPrevAte
End Property

Public Property Let dtDataPrevDe(ByVal vData As Date)
    mvardtDataPrevDe = vData
End Property

Public Property Get dtDataPrevDe() As Date
    dtDataPrevDe = mvardtDataPrevDe
End Property

Public Property Let dtDataProxAte(ByVal vData As Date)
    mvardtDataProxAte = vData
End Property

Public Property Get dtDataProxAte() As Date
    dtDataProxAte = mvardtDataProxAte
End Property

Public Property Let dtDataProxDe(ByVal vData As Date)
    mvardtDataProxDe = vData
End Property

Public Property Get dtDataProxDe() As Date
    dtDataProxDe = mvardtDataProxDe
End Property


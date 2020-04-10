VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTRVLiberaOCRSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarcolOcorrenciais As New Collection
Private mvardtDataEmissaoDe As Date
Private mvarlCodigoDe As Long
Private mvardtDataEmissaoAte As Date
Private mvarlCodigoAte As Long
Private mvariPago As Integer

Public Property Let iPago(ByVal vData As Integer)
    mvariPago = vData
End Property

Public Property Get iPago() As Integer
    iPago = mvariPago
End Property

Public Property Let lCodigoDe(ByVal vData As Long)
    mvarlCodigoDe = vData
End Property

Public Property Get lCodigoDe() As Long
    lCodigoDe = mvarlCodigoDe
End Property

Public Property Let lCodigoAte(ByVal vData As Long)
    mvarlCodigoAte = vData
End Property

Public Property Get lCodigoAte() As Long
    lCodigoAte = mvarlCodigoAte
End Property

Public Property Let dtDataEmissaoDe(ByVal vData As Date)
    mvardtDataEmissaoDe = vData
End Property

Public Property Get dtDataEmissaoDe() As Date
    dtDataEmissaoDe = mvardtDataEmissaoDe
End Property

Public Property Let dtDataEmissaoAte(ByVal vData As Date)
    mvardtDataEmissaoAte = vData
End Property

Public Property Get dtDataEmissaoAte() As Date
    dtDataEmissaoAte = mvardtDataEmissaoAte
End Property


Public Property Set colOcorrenciais(ByVal vData As Collection)
    Set mvarcolOcorrenciais = vData
End Property

Public Property Get colOcorrenciais() As Collection
    Set colOcorrenciais = mvarcolOcorrenciais
End Property

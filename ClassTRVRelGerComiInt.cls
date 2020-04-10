VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTRVRelGerComiInt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarlNumIntRel As Long
Private mvarlCodigo As Long
Private mvardtDataGeracao As Date
Private mvardHoraGeracao As Double
Private mvarsUsuario As String
Private mvardtDataPagtoDe As Date
Private mvardtDataPagtoAte As Date
Private mvardtDataEmiDe As Date
Private mvardtDataEmiAte As Date

Public Property Let lNumIntRel(ByVal vData As Long)
    mvarlNumIntRel = vData
End Property

Public Property Get lNumIntRel() As Long
    lNumIntRel = mvarlNumIntRel
End Property

Public Property Let lCodigo(ByVal vData As Long)
    mvarlCodigo = vData
End Property

Public Property Get lCodigo() As Long
    lCodigo = mvarlCodigo
End Property

Public Property Let dtDataGeracao(ByVal vData As Date)
    mvardtDataGeracao = vData
End Property

Public Property Get dtDataGeracao() As Date
    dtDataGeracao = mvardtDataGeracao
End Property

Public Property Let dHoraGeracao(ByVal vData As Double)
    mvardHoraGeracao = vData
End Property

Public Property Get dHoraGeracao() As Double
    dHoraGeracao = mvardHoraGeracao
End Property

Public Property Let sUsuario(ByVal vData As String)
    mvarsUsuario = vData
End Property

Public Property Get sUsuario() As String
    sUsuario = mvarsUsuario
End Property

Public Property Let dtDataPagtoDe(ByVal vData As Date)
    mvardtDataPagtoDe = vData
End Property

Public Property Get dtDataPagtoDe() As Date
    dtDataPagtoDe = mvardtDataPagtoDe
End Property

Public Property Let dtDataPagtoAte(ByVal vData As Date)
    mvardtDataPagtoAte = vData
End Property

Public Property Get dtDataPagtoAte() As Date
    dtDataPagtoAte = mvardtDataPagtoAte
End Property

Public Property Let dtDataEmiDe(ByVal vData As Date)
    mvardtDataEmiDe = vData
End Property

Public Property Get dtDataEmiDe() As Date
    dtDataEmiDe = mvardtDataEmiDe
End Property

Public Property Let dtDataEmiAte(ByVal vData As Date)
    mvardtDataEmiAte = vData
End Property

Public Property Get dtDataEmiAte() As Date
    dtDataEmiAte = mvardtDataEmiAte
End Property


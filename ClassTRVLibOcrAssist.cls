VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTRVLibOcrAssist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarcolOcorrenciais As New Collection
Private mvardtDataDe As Date
Private mvardtDataAte As Date
Private mvarsCodigoDe As String
Private mvarsCodigoAte As String
Private mvardtDataProgDe As Date
Private mvardtDataProgAte As Date
Private mvariTipo As Integer

Public Property Let iTipo(ByVal vData As Integer)
    mvariTipo = vData
End Property

Public Property Get iTipo() As Integer
    iTipo = mvariTipo
End Property

Public Property Let dtDataProgDe(ByVal vData As Date)
    mvardtDataProgDe = vData
End Property

Public Property Get dtDataProgDe() As Date
    dtDataProgDe = mvardtDataProgDe
End Property

Public Property Let dtDataProgAte(ByVal vData As Date)
    mvardtDataProgAte = vData
End Property

Public Property Get dtDataProgAte() As Date
    dtDataProgAte = mvardtDataProgAte
End Property

Public Property Let sCodigoDe(ByVal vData As String)
    mvarsCodigoDe = vData
End Property

Public Property Get sCodigoDe() As String
    sCodigoDe = mvarsCodigoDe
End Property

Public Property Let sCodigoAte(ByVal vData As String)
    mvarsCodigoAte = vData
End Property

Public Property Get sCodigoAte() As String
    sCodigoAte = mvarsCodigoAte
End Property

Public Property Let dtDataDe(ByVal vData As Date)
    mvardtDataDe = vData
End Property

Public Property Get dtDataDe() As Date
    dtDataDe = mvardtDataDe
End Property

Public Property Let dtDataAte(ByVal vData As Date)
    mvardtDataAte = vData
End Property

Public Property Get dtDataAte() As Date
    dtDataAte = mvardtDataAte
End Property

Public Property Set colOcorrenciais(ByVal vData As Collection)
    Set mvarcolOcorrenciais = vData
End Property

Public Property Get colOcorrenciais() As Collection
    Set colOcorrenciais = mvarcolOcorrenciais
End Property

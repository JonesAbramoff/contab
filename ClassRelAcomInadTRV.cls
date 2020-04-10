VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelAcomInadTRV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvaradtDatasRef(1 To 100) As Date
Private mvaraiFiliais(1 To 100) As Integer
Private mvarsTipoDoc As String
Private mvariNumFiliais As Integer
Private mvariNumDatas As Integer
Private mvardtDataVencDe As Date
Private mvardtDataVencAte As Date

Public Property Let dtDataVencDe(ByVal vData As Date)
    mvardtDataVencDe = vData
End Property

Public Property Get dtDataVencDe() As Date
    dtDataVencDe = mvardtDataVencDe
End Property

Public Property Let dtDataVencAte(ByVal vData As Date)
    mvardtDataVencAte = vData
End Property

Public Property Get dtDataVencAte() As Date
    dtDataVencAte = mvardtDataVencAte
End Property

Public Property Let iNumDatas(ByVal vData As Integer)
    mvariNumDatas = vData
End Property

Public Property Get iNumDatas() As Integer
    iNumDatas = mvariNumDatas
End Property

Public Property Let iNumFiliais(ByVal vData As Integer)
    mvariNumFiliais = vData
End Property

Public Property Get iNumFiliais() As Integer
    iNumFiliais = mvariNumFiliais
End Property

Public Property Let stipodoc(ByVal vData As String)
    mvarsTipoDoc = vData
End Property

Public Property Get stipodoc() As String
    stipodoc = mvarsTipoDoc
End Property

Public Property Let aiFiliais(ByVal iIndice As Integer, ByVal vData As Integer)
    mvaraiFiliais(iIndice) = vData
End Property

Public Property Get aiFiliais(ByVal iIndice As Integer) As Integer
    aiFiliais = mvaraiFiliais(iIndice)
End Property

Public Property Let adtDatasRef(ByVal iIndice As Integer, ByVal vData As Date)
    mvaradtDatasRef(iIndice) = vData
End Property

Public Property Get adtDatasRef(ByVal iIndice As Integer) As Date
    adtDatasRef = mvaradtDatasRef(iIndice)
End Property

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelProdConsumo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarlNumIntRel As Long
Private mvarsProdKit As String
Private mvarsVersaoKit As String
Private mvarsProdDe As String
Private mvarsProdAte As String
Private mvardtDataDe As Date
Private mvardtDataAte As Date
Private mvarcolNaturezas As New Collection
Private mvarsCategoria As String
Private mvarsCatItemDe As String
Private mvarsCatItemAte As String
Private mvarsFaixaData As String
Private mvarsFaixaProd As String
Private mvarsFaixaCat As String
Private mvarsFaixaKit As String
Private mvarsFaixaNat As String

Public Property Let lNumIntRel(ByVal vData As Long)
    mvarlNumIntRel = vData
End Property

Public Property Get lNumIntRel() As Long
    lNumIntRel = mvarlNumIntRel
End Property

Public Property Let sFaixaData(ByVal vData As String)
    mvarsFaixaData = vData
End Property

Public Property Get sFaixaData() As String
    sFaixaData = mvarsFaixaData
End Property

Public Property Let sFaixaProd(ByVal vData As String)
    mvarsFaixaProd = vData
End Property

Public Property Get sFaixaProd() As String
    sFaixaProd = mvarsFaixaProd
End Property

Public Property Let sFaixaCat(ByVal vData As String)
    mvarsFaixaCat = vData
End Property

Public Property Get sFaixaCat() As String
    sFaixaCat = mvarsFaixaCat
End Property

Public Property Let sFaixaKit(ByVal vData As String)
    mvarsFaixaKit = vData
End Property

Public Property Get sFaixaKit() As String
    sFaixaKit = mvarsFaixaKit
End Property

Public Property Let sFaixaNat(ByVal vData As String)
    mvarsFaixaNat = vData
End Property

Public Property Get sFaixaNat() As String
    sFaixaNat = mvarsFaixaNat
End Property

Public Property Let sCatItemAte(ByVal vData As String)
    mvarsCatItemAte = vData
End Property

Public Property Get sCatItemAte() As String
    sCatItemAte = mvarsCatItemAte
End Property

Public Property Let sCatItemDe(ByVal vData As String)
    mvarsCatItemDe = vData
End Property

Public Property Get sCatItemDe() As String
    sCatItemDe = mvarsCatItemDe
End Property

Public Property Let sCategoria(ByVal vData As String)
    mvarsCategoria = vData
End Property

Public Property Get sCategoria() As String
    sCategoria = mvarsCategoria
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

Public Property Let sProdAte(ByVal vData As String)
    mvarsProdAte = vData
End Property

Public Property Get sProdAte() As String
    sProdAte = mvarsProdAte
End Property

Public Property Let sProdDe(ByVal vData As String)
    mvarsProdDe = vData
End Property

Public Property Get sProdDe() As String
    sProdDe = mvarsProdDe
End Property

Public Property Let sProdKit(ByVal vData As String)
    mvarsProdKit = vData
End Property

Public Property Get sProdKit() As String
    sProdKit = mvarsProdKit
End Property

Public Property Let sVersaoKit(ByVal vData As String)
    mvarsVersaoKit = vData
End Property

Public Property Get sVersaoKit() As String
    sVersaoKit = mvarsVersaoKit
End Property

Public Property Set colNaturezas(ByVal vData As Object)
    Set mvarcolNaturezas = vData
End Property

Public Property Get colNaturezas() As Collection
    Set colNaturezas = mvarcolNaturezas
End Property


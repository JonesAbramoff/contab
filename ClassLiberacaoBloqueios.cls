VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassLiberacaoBloqueios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarcolCodBloqueios As New Collection
Private mvarlPedVendasDe As Long
Private mvarlPedVendasAte As Long
Private mvardtBloqueioDe As Date
Private mvardtBloqueioAte As Date
Private mvarsOrdenacao As String
Private mvarcolBloqueioLiberacaoInfo As New Collection
Private mvariFilialEmpresa As Integer

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Set colBloqueioLiberacaoInfo(ByVal vData As Object)
    Set mvarcolBloqueioLiberacaoInfo = vData
End Property

Public Property Get colBloqueioLiberacaoInfo() As Collection
    Set colBloqueioLiberacaoInfo = mvarcolBloqueioLiberacaoInfo
End Property

Public Property Let sOrdenacao(ByVal vData As String)
    mvarsOrdenacao = vData
End Property

Public Property Get sOrdenacao() As String
    sOrdenacao = mvarsOrdenacao
End Property

Public Property Let dtBloqueioAte(ByVal vData As Date)
    mvardtBloqueioAte = vData
End Property

Public Property Get dtBloqueioAte() As Date
    dtBloqueioAte = mvardtBloqueioAte
End Property

Public Property Let dtBloqueioDe(ByVal vData As Date)
    mvardtBloqueioDe = vData
End Property

Public Property Get dtBloqueioDe() As Date
    dtBloqueioDe = mvardtBloqueioDe
End Property

Public Property Let lPedVendasAte(ByVal vData As Long)
    mvarlPedVendasAte = vData
End Property

Public Property Get lPedVendasAte() As Long
    lPedVendasAte = mvarlPedVendasAte
End Property

Public Property Let lPedVendasDe(ByVal vData As Long)
    mvarlPedVendasDe = vData
End Property

Public Property Get lPedVendasDe() As Long
    lPedVendasDe = mvarlPedVendasDe
End Property

Public Property Set colCodBloqueios(ByVal vData As Object)
    Set mvarcolCodBloqueios = vData
End Property

Public Property Get colCodBloqueios() As Collection
    Set colCodBloqueios = mvarcolCodBloqueios
End Property




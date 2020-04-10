VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassLibBloqGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarcolCodBloqueios As New Collection
Private mvarlCodigoDe As Long
Private mvarlCodigoAte As Long
Private mvardtBloqueioDe As Date
Private mvardtBloqueioAte As Date
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

Public Property Let lCodigoAte(ByVal vData As Long)
    mvarlCodigoAte = vData
End Property

Public Property Get lCodigoAte() As Long
    lCodigoAte = mvarlCodigoAte
End Property

Public Property Let lCodigoDe(ByVal vData As Long)
    mvarlCodigoDe = vData
End Property

Public Property Get lCodigoDe() As Long
    lCodigoDe = mvarlCodigoDe
End Property

Public Property Set colCodBloqueios(ByVal vData As Object)
    Set mvarcolCodBloqueios = vData
End Property

Public Property Get colCodBloqueios() As Collection
    Set colCodBloqueios = mvarcolCodBloqueios
End Property




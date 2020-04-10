VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCamposGenericos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarlCodigo As Long
Private mvarsDescricao As String
Private mvarsComentarios As String
Private mvarlProxCodValor As Long
Private mvarsValidaExclusao As String
Private mvarcolCamposGenericosValores As New Collection

Public Property Let lCodigo(ByVal vData As Long)
    mvarlCodigo = vData
End Property

Public Property Get lCodigo() As Long
    lCodigo = mvarlCodigo
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let sComentarios(ByVal vData As String)
    mvarsComentarios = vData
End Property

Public Property Get sComentarios() As String
    sComentarios = mvarsComentarios
End Property

Public Property Let lProxCodValor(ByVal vData As Long)
    mvarlProxCodValor = vData
End Property

Public Property Get lProxCodValor() As Long
    lProxCodValor = mvarlProxCodValor
End Property

Public Property Let sValidaExclusao(ByVal vData As String)
    mvarsValidaExclusao = vData
End Property

Public Property Get sValidaExclusao() As String
    sValidaExclusao = mvarsValidaExclusao
End Property

Public Property Set colCamposGenericosValores(ByVal vData As Object)
    Set mvarcolCamposGenericosValores = vData
End Property

Public Property Get colCamposGenericosValores() As Collection
    Set colCamposGenericosValores = mvarcolCamposGenericosValores
End Property


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassIdiomaTabCampo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarsNomeCampo As String
Private mvarcolTextos As New Collection

Public Property Set colTextos(ByVal vData As Collection)
    Set mvarcolTextos = vData
End Property

Public Property Get colTextos() As Collection
    Set colTextos = mvarcolTextos
End Property

Public Property Let sNomeCampo(ByVal vData As String)
    mvarsNomeCampo = vData
End Property

Public Property Get sNomeCampo() As String
    sNomeCampo = mvarsNomeCampo
End Property

Public Function Reiniciar() As Long
Dim objTexto As ClassIdiomaTabCampoTexto
    For Each objTexto In mvarcolTextos
        Call objTexto.Reiniciar
    Next
End Function


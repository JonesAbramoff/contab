VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassOpcoesTelas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarlCodigo As Long
Private mvarsOpcao As String
Private mvarsNomeTela As String
Private mvariPadrao As Integer
Private mvarcolOpcoesTelasValores As New Collection

Public Property Let lCodigo(ByVal vData As Long)
    mvarlCodigo = vData
End Property

Public Property Get lCodigo() As Long
    lCodigo = mvarlCodigo
End Property

Public Property Let sOpcao(ByVal vData As String)
    mvarsOpcao = vData
End Property

Public Property Get sOpcao() As String
    sOpcao = mvarsOpcao
End Property

Public Property Let sNomeTela(ByVal vData As String)
    mvarsNomeTela = vData
End Property

Public Property Get sNomeTela() As String
    sNomeTela = mvarsNomeTela
End Property

Public Property Let iPadrao(ByVal vData As Integer)
    mvariPadrao = vData
End Property

Public Property Get iPadrao() As Integer
    iPadrao = mvariPadrao
End Property

Public Property Set colOpcoesTelasValores(ByVal vData As Object)
    Set mvarcolOpcoesTelasValores = vData
End Property

Public Property Get colOpcoesTelasValores() As Collection
    Set colOpcoesTelasValores = mvarcolOpcoesTelasValores
End Property


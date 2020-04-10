VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassGrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarsCodigo As String
Private mvarsDescricao As String
Private mvarcolCategoria As New Collection
Private mvariLayout As Integer

Public Property Let iLayout(ByVal vData As Integer)
    mvariLayout = vData
End Property

Public Property Get iLayout() As Integer
    iLayout = mvariLayout
End Property

Public Property Set colCategoria(ByVal vData As Collection)
    Set mvarcolCategoria = vData
End Property

Public Property Get colCategoria() As Collection
    Set colCategoria = mvarcolCategoria
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let sCodigo(ByVal vData As String)
    mvarsCodigo = vData
End Property

Public Property Get sCodigo() As String
    sCodigo = mvarsCodigo
End Property


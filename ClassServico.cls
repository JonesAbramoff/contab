VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassServico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarsProduto As String
Private mvarsDescricao As String
Private mvarcolServItemServ As New Collection
Private mvardQuant As Double

Public Property Let dQuant(ByVal vData As Double)
    mvardQuant = vData
End Property

Public Property Get dQuant() As Double
    dQuant = mvardQuant
End Property

Public Property Set colServItemServ(ByVal vData As Object)
    Set mvarcolServItemServ = vData
End Property

Public Property Get colServItemServ() As Collection
    Set colServItemServ = mvarcolServItemServ
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property




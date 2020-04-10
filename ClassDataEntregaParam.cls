VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassDataEntregaParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarsProduto As String
Private mvardQuantidade As Double
Private mvarcolDataEntrega As Collection
Private mvarcolDataEntregaPV As Collection

Public Property Set colDataEntregaPV(ByVal vData As Object)
    Set mvarcolDataEntregaPV = vData
End Property

Public Property Get colDataEntregaPV() As Collection
    Set colDataEntregaPV = mvarcolDataEntregaPV
End Property

Public Property Set colDataEntrega(ByVal vData As Object)
    Set mvarcolDataEntrega = vData
End Property

Public Property Get colDataEntrega() As Collection
    Set colDataEntrega = mvarcolDataEntrega
End Property

Public Property Let dQuantidade(ByVal vData As Double)
    mvardQuantidade = vData
End Property

Public Property Get dQuantidade() As Double
    dQuantidade = mvardQuantidade
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property


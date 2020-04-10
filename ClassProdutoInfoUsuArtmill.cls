VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassProdutoInfoUsu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarsCodigo As String

Private mvarsDetalheCor As String
Private mvarsDimEmbalagem As String
Private mvarsCodAnterior As String

Public Property Let sCodAnterior(ByVal vData As String)
    mvarsCodAnterior = vData
End Property

Public Property Get sCodAnterior() As String
    sCodAnterior = mvarsCodAnterior
End Property

Public Property Let sDimEmbalagem(ByVal vData As String)
    mvarsDimEmbalagem = vData
End Property

Public Property Get sDimEmbalagem() As String
    sDimEmbalagem = mvarsDimEmbalagem
End Property

Public Property Let sDetalheCor(ByVal vData As String)
    mvarsDetalheCor = vData
End Property

Public Property Get sDetalheCor() As String
    sDetalheCor = mvarsDetalheCor
End Property

Public Property Let sCodigo(ByVal vData As String)
    mvarsCodigo = vData
End Property

Public Property Get sCodigo() As String
    sCodigo = mvarsCodigo
End Property


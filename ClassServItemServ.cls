VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassServItemServ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarsProduto As String
Private mvariCodItemServico As Integer
Private mvariOrdem As Integer
Private mvarsDescricao As String

Public Property Let iOrdem(ByVal vData As Integer)
    mvariOrdem = vData
End Property

Public Property Get iOrdem() As Integer
    iOrdem = mvariOrdem
End Property

Public Property Let iCodItemServico(ByVal vData As Integer)
    mvariCodItemServico = vData
End Property

Public Property Get iCodItemServico() As Integer
    iCodItemServico = mvariCodItemServico
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property



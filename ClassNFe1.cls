VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassNFe1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarobjNFiscal As New ClassNFiscal
Private mvarsClienteNomeReduzido As String
Private mvarsFornecedorNomeReduzido As String
Private mvarlLote As Long
Private mvarsStatus As String
Private mvarsTipoNFiscal As String

Public Property Let sTipoNFiscal(ByVal vData As String)
    mvarsTipoNFiscal = vData
End Property

Public Property Get sTipoNFiscal() As String
    sTipoNFiscal = mvarsTipoNFiscal
End Property

Public Property Let sStatus(ByVal vData As String)
    mvarsStatus = vData
End Property

Public Property Get sStatus() As String
    sStatus = mvarsStatus
End Property

Public Property Let lLote(ByVal vData As Long)
    mvarlLote = vData
End Property

Public Property Get lLote() As Long
    lLote = mvarlLote
End Property

Public Property Let sFornecedorNomeReduzido(ByVal vData As String)
    mvarsFornecedorNomeReduzido = vData
End Property

Public Property Get sFornecedorNomeReduzido() As String
    sFornecedorNomeReduzido = mvarsFornecedorNomeReduzido
End Property

Public Property Let sClienteNomeReduzido(ByVal vData As String)
    mvarsClienteNomeReduzido = vData
End Property

Public Property Get sClienteNomeReduzido() As String
    sClienteNomeReduzido = mvarsClienteNomeReduzido
End Property

Public Property Set objNFiscal(ByVal vData As Object)
    Set mvarobjNFiscal = vData
End Property

Public Property Get objNFiscal() As ClassNFiscal
    Set objNFiscal = mvarobjNFiscal
End Property


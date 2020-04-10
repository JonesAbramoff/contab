VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassDespachante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvariCodigo As Integer
Private mvarsCGC As String
Private mvarsNome As String
Private mvarsNomeReduzido As String
Private mvarlEndereco As Long
Private mvarobjEndereco As New ClassEndereco
Private mvarcolContato As New Collection

Public Property Set colContato(ByVal vData As Object)
    Set mvarcolContato = vData
End Property

Public Property Get colContato() As Collection
    Set colContato = mvarcolContato
End Property

Public Property Set objEndereco(ByVal vData As Object)
    Set mvarobjEndereco = vData
End Property

Public Property Get objEndereco() As ClassEndereco
    Set objEndereco = mvarobjEndereco
End Property

Public Property Let lEndereco(ByVal vData As Long)
    mvarlEndereco = vData
End Property

Public Property Get lEndereco() As Long
    lEndereco = mvarlEndereco
End Property

Public Property Let sNomeReduzido(ByVal vData As String)
    mvarsNomeReduzido = vData
End Property

Public Property Get sNomeReduzido() As String
    sNomeReduzido = mvarsNomeReduzido
End Property

Public Property Let sNome(ByVal vData As String)
    mvarsNome = vData
End Property

Public Property Get sNome() As String
    sNome = mvarsNome
End Property

Public Property Let sCgc(ByVal vData As String)
    mvarsCGC = vData
End Property

Public Property Get sCgc() As String
    sCgc = mvarsCGC
End Property

Public Property Let iCodigo(ByVal vData As Integer)
    mvariCodigo = vData
End Property

Public Property Get iCodigo() As Integer
    iCodigo = mvariCodigo
End Property





VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRegiaoVenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariCodigo As Integer
Private mvarsDescricao As String
Private mvariCodigoPais As Integer
Private mvarsGerente As String
Private mvarsUsuarioCobrador As String
Private mvarobjInfoUsu As Object
Private mvarsUsuRespCallCenter As String

Public Property Let sUsuRespCallCenter(ByVal vData As String)
    mvarsUsuRespCallCenter = vData
End Property

Public Property Get sUsuRespCallCenter() As String
    sUsuRespCallCenter = mvarsUsuRespCallCenter
End Property

Public Property Set objInfoUsu(ByVal vData As Object)
    Set mvarobjInfoUsu = vData
End Property

Public Property Get objInfoUsu() As Object
    Set objInfoUsu = mvarobjInfoUsu
End Property

Public Property Let sUsuarioCobrador(ByVal vData As String)
    mvarsUsuarioCobrador = vData
End Property

Public Property Get sUsuarioCobrador() As String
    sUsuarioCobrador = mvarsUsuarioCobrador
End Property

Public Property Let sGerente(ByVal vData As String)
    mvarsGerente = vData
End Property

Public Property Get sGerente() As String
    sGerente = mvarsGerente
End Property

Public Property Let iCodigoPais(ByVal vData As Integer)
    mvariCodigoPais = vData
End Property

Public Property Get iCodigoPais() As Integer
    iCodigoPais = mvariCodigoPais
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let iCodigo(ByVal vData As Integer)
    mvariCodigo = vData
End Property

Public Property Get iCodigo() As Integer
    iCodigo = mvariCodigo
End Property




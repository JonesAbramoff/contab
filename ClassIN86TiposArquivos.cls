VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassIN86TiposArquivos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariCodigo As Integer
Private mvarsDescricao As String
Private mvarsPrefixoNome As String
Private mvarsRotinaGeradora As String
Private mvariAuxiliar As Integer
Private mvarsLayoutArquivo As String 'Por Leo em 09/05/02

'Por Leo em 09/05/02
Public Property Let sLayoutArquivo(ByVal vData As String)
    mvarsLayoutArquivo = vData
End Property

'Por Leo em 09/05/02
Public Property Get sLayoutArquivo() As String
    sLayoutArquivo = mvarsLayoutArquivo
End Property

Public Property Let iAuxiliar(ByVal vData As Integer)
    mvariAuxiliar = vData
End Property

Public Property Get iAuxiliar() As Integer
    iAuxiliar = mvariAuxiliar
End Property

Public Property Let sRotinaGeradora(ByVal vData As String)
    mvarsRotinaGeradora = vData
End Property

Public Property Get sRotinaGeradora() As String
    sRotinaGeradora = mvarsRotinaGeradora
End Property

Public Property Let sPrefixoNome(ByVal vData As String)
    mvarsPrefixoNome = vData
End Property

Public Property Get sPrefixoNome() As String
    sPrefixoNome = mvarsPrefixoNome
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

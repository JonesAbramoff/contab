VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFornMapaCot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlCodigo As Long
Private mvarlCodFornecedor As Long
Private mvarsNome As String
Private mvarsTel1 As String
Private mvarsTel2 As String
Private mvarsContato As String
Private mvariCodFilial As Integer
Private mvardQuantidade As Double
Private mvarlCodPedCotacao As Long

Public Property Let lCodPedCotacao(ByVal vData As Long)
    mvarlCodPedCotacao = vData
End Property

Public Property Get lCodPedCotacao() As Long
    lCodPedCotacao = mvarlCodPedCotacao
End Property

Public Property Let dQuantidade(ByVal vData As Double)
    mvardQuantidade = vData
End Property

Public Property Get dQuantidade() As Double
    dQuantidade = mvardQuantidade
End Property

Public Property Let lCodigo(ByVal vData As Long)
    mvarlCodigo = vData
End Property

Public Property Get lCodigo() As Long
    lCodigo = mvarlCodigo
End Property

Public Property Let lCodFornecedor(ByVal vData As Long)
    mvarlCodFornecedor = vData
End Property

Public Property Get lCodFornecedor() As Long
    lCodFornecedor = mvarlCodFornecedor
End Property

Public Property Let iCodFilial(ByVal vData As Integer)
    mvariCodFilial = vData
End Property

Public Property Get iCodFilial() As Integer
    iCodFilial = mvariCodFilial
End Property

Public Property Let sTel1(ByVal vData As String)
    mvarsTel1 = vData
End Property

Public Property Get sTel1() As String
    sTel1 = mvarsTel1
End Property

Public Property Let sTel2(ByVal vData As String)
    mvarsTel2 = vData
End Property

Public Property Get sTel2() As String
    sTel2 = mvarsTel2
End Property

Public Property Let sNome(ByVal vData As String)
    mvarsNome = vData
End Property

Public Property Get sNome() As String
    sNome = mvarsNome
End Property

Public Property Let sContato(ByVal vData As String)
    mvarsContato = vData
End Property

Public Property Get sContato() As String
    sContato = mvarsContato
End Property

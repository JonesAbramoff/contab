VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassVendedorPur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvariVendedor As Integer
Private mvardtDataInicio As Date
Private mvarsObservacao As String
Private mvarcolExcecoes As New Collection

Public Property Set colExcecoes(ByVal vData As Collection)
    Set mvarcolExcecoes = vData
End Property

Public Property Get colExcecoes() As Collection
    Set colExcecoes = mvarcolExcecoes
End Property

Public Property Let iVendedor(ByVal vData As Integer)
    mvariVendedor = vData
End Property

Public Property Get iVendedor() As Integer
    iVendedor = mvariVendedor
End Property

Public Property Let dtDataInicio(ByVal vData As Date)
    mvardtDataInicio = vData
End Property

Public Property Get dtDataInicio() As Date
    dtDataInicio = mvardtDataInicio
End Property

Public Property Let sObservacao(ByVal vData As String)
    mvarsObservacao = vData
End Property

Public Property Get sObservacao() As String
    sObservacao = mvarsObservacao
End Property

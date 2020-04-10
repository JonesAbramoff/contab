VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassProdEstSeleciona"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarsProdutoInicial As String
Private mvarsProdutoFinal As String
Private mvariFilialEmpresa As Integer
Private mvardtData As Date

Public Property Let dtData(ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let sProdutoFinal(ByVal vData As String)
    mvarsProdutoFinal = vData
End Property

Public Property Get sProdutoFinal() As String
    sProdutoFinal = mvarsProdutoFinal
End Property

Public Property Let sProdutoInicial(ByVal vData As String)
    mvarsProdutoInicial = vData
End Property

Public Property Get sProdutoInicial() As String
    sProdutoInicial = mvarsProdutoInicial
End Property



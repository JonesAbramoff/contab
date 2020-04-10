VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelDREConta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariCodigo As Integer
Private mvariItem As Integer
Private mvarsContaInicial As String
Private mvarsContaFinal As String
Private mvarsModelo As String
Private mvarsCclInicial As String
Private mvarsCclFinal As String

Public Property Let sCclFinal(ByVal vData As String)
    mvarsCclFinal = vData
End Property

Public Property Get sCclFinal() As String
    sCclFinal = mvarsCclFinal
End Property

Public Property Let sCclInicial(ByVal vData As String)
    mvarsCclInicial = vData
End Property

Public Property Get sCclInicial() As String
    sCclInicial = mvarsCclInicial
End Property

Public Property Let sModelo(ByVal vData As String)
Attribute sModelo.VB_Description = "Identificador do Modelo"
    mvarsModelo = vData
End Property

Public Property Get sModelo() As String
    sModelo = mvarsModelo
End Property

Public Property Let sContaFinal(ByVal vData As String)
Attribute sContaFinal.VB_Description = "Conta Final"
    mvarsContaFinal = vData
End Property

Public Property Get sContaFinal() As String
    sContaFinal = mvarsContaFinal
End Property

Public Property Let sContaInicial(ByVal vData As String)
Attribute sContaInicial.VB_Description = "Conta Inicial"
    mvarsContaInicial = vData
End Property

Public Property Get sContaInicial() As String
    sContaInicial = mvarsContaInicial
End Property

Public Property Let iItem(ByVal vData As Integer)
Attribute iItem.VB_Description = "Item identificador do grupo de contas"
    mvariItem = vData
End Property

Public Property Get iItem() As Integer
    iItem = mvariItem
End Property

Public Property Let iCodigo(ByVal vData As Integer)
Attribute iCodigo.VB_Description = "Codigo Identificador do Elemento"
    mvariCodigo = vData
End Property

Public Property Get iCodigo() As Integer
    iCodigo = mvariCodigo
End Property




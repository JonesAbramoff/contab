VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassMRPSeleciona"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlClienteInicial As Long
Private mvarlClienteFinal As Long
Private mvarsOPInicial As String
Private mvarsOPFinal As String
Private mvarsProdutoInicial As String
Private mvarsProdutoFinal As String
Private mvardtDataOPInicial As Date
Private mvardtDataOPFinal As Date

Public Property Let dtDataOPFinal(ByVal vData As Date)
    mvardtDataOPFinal = vData
End Property

Public Property Get dtDataOPFinal() As Date
    dtDataOPFinal = mvardtDataOPFinal
End Property

Public Property Let dtDataOPInicial(ByVal vData As Date)
    mvardtDataOPInicial = vData
End Property

Public Property Get dtDataOPInicial() As Date
    dtDataOPInicial = mvardtDataOPInicial
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

Public Property Let sOPFinal(ByVal vData As String)
    mvarsOPFinal = vData
End Property

Public Property Get sOPFinal() As String
    sOPFinal = mvarsOPFinal
End Property

Public Property Let sOPInicial(ByVal vData As String)
    mvarsOPInicial = vData
End Property

Public Property Get sOPInicial() As String
    sOPInicial = mvarsOPInicial
End Property

Public Property Let lClienteFinal(ByVal vData As Long)
    mvarlClienteFinal = vData
End Property

Public Property Get lClienteFinal() As Long
    lClienteFinal = mvarlClienteFinal
End Property

Public Property Let lClienteInicial(ByVal vData As Long)
    mvarlClienteInicial = vData
End Property

Public Property Get lClienteInicial() As Long
    lClienteInicial = mvarlClienteInicial
End Property



VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassProdutoCNAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarsProduto As String
Private mvarsCNAE As String
Private mvarcolCidades As Collection
Private mvariLocServCliente As Integer
Private mvariLocIncidImpCliente As Integer

Public Property Let iLocIncidImpCliente(ByVal vData As Integer)
    mvariLocIncidImpCliente = vData
End Property

Public Property Get iLocIncidImpCliente() As Integer
    iLocIncidImpCliente = mvariLocIncidImpCliente
End Property

Public Property Let iLocServCliente(ByVal vData As Integer)
    mvariLocServCliente = vData
End Property

Public Property Get iLocServCliente() As Integer
    iLocServCliente = mvariLocServCliente
End Property

Public Property Set colCidades(ByVal vData As Collection)
    Set mvarcolCidades = vData
End Property

Public Property Get colCidades() As Collection
    Set colCidades = mvarcolCidades
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let sCNAE(ByVal vData As String)
    mvarsCNAE = vData
End Property

Public Property Get sCNAE() As String
    sCNAE = mvarsCNAE
End Property

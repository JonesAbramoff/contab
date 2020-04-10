VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassDan_OS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarsOS As String
Private mvarlCliente As Long
Private mvarsModelo As String
Private mvarsNumSerie As String
Private mvarcolItens As New Collection

Public Property Set colItens(ByVal vData As Collection)
    Set mvarcolItens = vData
End Property

Public Property Get colItens() As Collection
    Set colItens = mvarcolItens
End Property

Public Property Let sOS(ByVal vData As String)
    mvarsOS = vData
End Property

Public Property Get sOS() As String
    sOS = mvarsOS
End Property

Public Property Let lCliente(ByVal vData As Long)
    mvarlCliente = vData
End Property

Public Property Get lCliente() As Long
    lCliente = mvarlCliente
End Property

Public Property Let sModelo(ByVal vData As String)
    mvarsModelo = vData
End Property

Public Property Get sModelo() As String
    sModelo = mvarsModelo
End Property

Public Property Let sNumSerie(ByVal vData As String)
    mvarsNumSerie = vData
End Property

Public Property Get sNumSerie() As String
    sNumSerie = mvarsNumSerie
End Property

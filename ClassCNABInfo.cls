VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCNABInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvariCodCobrador As Integer 'local copy
Private mvariCarteiraCobrador As Integer 'local copy
Private mvarcolInformacoes As New Collection 'local copy
Public Property Set colInformacoes(ByVal vData As Collection)
    Set mvarcolInformacoes = vData
End Property

Public Property Get colInformacoes() As Collection
    Set colInformacoes = mvarcolInformacoes
End Property

Public Property Let iCarteiraCobrador(ByVal vData As Integer)
    mvariCarteiraCobrador = vData
End Property

Public Property Get iCarteiraCobrador() As Integer
    iCarteiraCobrador = mvariCarteiraCobrador
End Property

Public Property Let iCodCobrador(ByVal vData As Integer)
    mvariCodCobrador = vData
End Property

Public Property Get iCodCobrador() As Integer
    iCodCobrador = mvariCodCobrador
End Property

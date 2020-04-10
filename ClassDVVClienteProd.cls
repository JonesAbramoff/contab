VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassDVVClienteProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariFilialEmpresa As Integer
Private mvarlCodCliente As Long
Private mvariCodFilial As Integer
Private mvarsProduto As String
Private mvardPercDVV As Double
Private mvariPaletizacao As Integer

Public Property Let iPaletizacao(ByVal vData As Integer)
    mvariPaletizacao = vData
End Property

Public Property Get iPaletizacao() As Integer
    iPaletizacao = mvariPaletizacao
End Property

Public Property Let dPercDVV(ByVal vData As Double)
    mvardPercDVV = vData
End Property

Public Property Get dPercDVV() As Double
    dPercDVV = mvardPercDVV
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let iCodFilial(ByVal vData As Integer)
    mvariCodFilial = vData
End Property

Public Property Get iCodFilial() As Integer
    iCodFilial = mvariCodFilial
End Property

Public Property Let lCodCliente(ByVal vData As Long)
    mvarlCodCliente = vData
End Property

Public Property Get lCodCliente() As Long
    lCodCliente = mvarlCodCliente
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

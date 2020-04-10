VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassNaturezaOp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarsCodigo As String
Private mvarsDescricao As String
Private mvarsDescrNF As String
Private mvarsNatBCCredPadrao As String

'nfe 3.10
Public iFinalidadeEspec As Integer

Public Property Let sNatBCCredPadrao(ByVal vData As String)
    mvarsNatBCCredPadrao = vData
End Property

Public Property Get sNatBCCredPadrao() As String
    sNatBCCredPadrao = mvarsNatBCCredPadrao
End Property

Public Property Let sDescrNF(ByVal vData As String)
    mvarsDescrNF = vData
End Property

Public Property Get sDescrNF() As String
    sDescrNF = mvarsDescrNF
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let sCodigo(ByVal vData As String)
    mvarsCodigo = vData
End Property

Public Property Get sCodigo() As String
    sCodigo = mvarsCodigo
End Property




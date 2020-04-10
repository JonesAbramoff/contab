VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassEstoqueTerc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarsProduto As String
Private mvariEscaninho As Integer
Private mvariTipoCliForn As Integer
Private mvarlCliForn As Long
Private mvariFilial As Integer
Private mvardQuantidade As Double
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

Public Property Let iEscaninho(ByVal vData As Integer)
    mvariEscaninho = vData
End Property

Public Property Get iEscaninho() As Integer
    iEscaninho = mvariEscaninho
End Property

Public Property Let iTipoCliForn(ByVal vData As Integer)
    mvariTipoCliForn = vData
End Property

Public Property Get iTipoCliForn() As Integer
    iTipoCliForn = mvariTipoCliForn
End Property

Public Property Let lCliForn(ByVal vData As Long)
    mvarlCliForn = vData
End Property

Public Property Get lCliForn() As Long
    lCliForn = mvarlCliForn
End Property

Public Property Let iFilial(ByVal vData As Integer)
    mvariFilial = vData
End Property

Public Property Get iFilial() As Integer
    iFilial = mvariFilial
End Property

Public Property Let dQuantidade(ByVal vData As Double)
    mvardQuantidade = vData
End Property

Public Property Get dQuantidade() As Double
    dQuantidade = mvardQuantidade
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

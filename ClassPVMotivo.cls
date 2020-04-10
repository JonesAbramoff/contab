VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPVMotivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarsProduto As String
Private mvardPrecoTabela As Double
Private mvardPrecoInformado As Double
Private mvarsMotivo As String

Public Property Let sMotivo(ByVal vData As String)
    mvarsMotivo = vData
End Property

Public Property Get sMotivo() As String
    sMotivo = mvarsMotivo
End Property

Public Property Let dPrecoInformado(ByVal vData As Double)
    mvardPrecoInformado = vData
End Property

Public Property Get dPrecoInformado() As Double
    dPrecoInformado = mvardPrecoInformado
End Property

Public Property Let dPrecoTabela(ByVal vData As Double)
    mvardPrecoTabela = vData
End Property

Public Property Get dPrecoTabela() As Double
    dPrecoTabela = mvardPrecoTabela
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property


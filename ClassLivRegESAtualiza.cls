VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassLivRegESAtualiza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarlFornecedorDe As Long
Private mvarlFornecedorAte As Long
Private mvarlClienteDe As Long
Private mvarlClienteAte As Long
Private mvardtData As Date
Private mvariIncluiLivRegFechados As Integer
Private mvariAtualizaCliente As Integer

Public Property Let iAtualizaCliente(ByVal vData As Integer)
    mvariAtualizaCliente = vData
End Property

Public Property Get iAtualizaCliente() As Integer
    iAtualizaCliente = mvariAtualizaCliente
End Property

Public Property Let iIncluiLivRegFechados(ByVal vData As Integer)
    mvariIncluiLivRegFechados = vData
End Property

Public Property Get iIncluiLivRegFechados() As Integer
    iIncluiLivRegFechados = mvariIncluiLivRegFechados
End Property

Public Property Let dtData(ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let lClienteAte(ByVal vData As Long)
    mvarlClienteAte = vData
End Property

Public Property Get lClienteAte() As Long
    lClienteAte = mvarlClienteAte
End Property

Public Property Let lClienteDe(ByVal vData As Long)
    mvarlClienteDe = vData
End Property

Public Property Get lClienteDe() As Long
    lClienteDe = mvarlClienteDe
End Property

Public Property Let lFornecedorAte(ByVal vData As Long)
    mvarlFornecedorAte = vData
End Property

Public Property Get lFornecedorAte() As Long
    lFornecedorAte = mvarlFornecedorAte
End Property

Public Property Let lFornecedorDe(ByVal vData As Long)
    mvarlFornecedorDe = vData
End Property

Public Property Get lFornecedorDe() As Long
    lFornecedorDe = mvarlFornecedorDe
End Property

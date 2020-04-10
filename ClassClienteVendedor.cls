VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassClienteFaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlClienteDe As Long
Private mvarlClienteAte As Long
Private mvariVendedorDe As Integer
Private mvariVendedorAte As Integer
Private mvardtDataDe As Date
Private mvardtDataAte As Date

Public Property Let dtDataAte(ByVal vData As Date)
    mvardtDataAte = vData
End Property

Public Property Get dtDataAte() As Date
    dtDataAte = mvardtDataAte
End Property

Public Property Let dtDataDe(ByVal vData As Date)
    mvardtDataDe = vData
End Property

Public Property Get dtDataDe() As Date
    dtDataDe = mvardtDataDe
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

Public Property Let iVendedorDe(ByVal vData As Integer)
    mvariVendedorDe = vData
End Property

Public Property Get iVendedorDe() As Integer
    iVendedorDe = mvariVendedorDe
End Property

Public Property Let iVendedorAte(ByVal vData As Integer)
    mvariVendedorAte = vData
End Property

Public Property Get iVendedorAte() As Integer
    iVendedorAte = mvariVendedorAte
End Property

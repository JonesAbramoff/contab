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
Private mvarlNumIntRelFiltroFilCliCat As Long
Private mvariComprou As Integer

Public Property Let iComprou(ByVal vData As Integer)
    mvariComprou = vData
End Property

Public Property Get iComprou() As Integer
    iComprou = mvariComprou
End Property

Public Property Let lNumIntRelFiltroFilCliCat(ByVal vData As Long)
    mvarlNumIntRelFiltroFilCliCat = vData
End Property

Public Property Get lNumIntRelFiltroFilCliCat() As Long
    lNumIntRelFiltroFilCliCat = mvarlNumIntRelFiltroFilCliCat
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

Private Sub Class_Initialize()
    mvariComprou = -1
End Sub

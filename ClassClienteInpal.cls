VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassClienteInpal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariPadraoTaxaFin As Integer
Private mvariPadraoJuros As Integer
Private mvardTaxaFinanceira As Double
Private mvardJuros As Double

Public Property Let dJuros(ByVal vData As Double)
    mvardJuros = vData
End Property

Public Property Get dJuros() As Double
    dJuros = mvardJuros
End Property

Public Property Let dTaxaFinanceira(ByVal vData As Double)
    mvardTaxaFinanceira = vData
End Property

Public Property Get dTaxaFinanceira() As Double
    dTaxaFinanceira = mvardTaxaFinanceira
End Property

Public Property Let iPadraoJuros(ByVal vData As Integer)
    mvariPadraoJuros = vData
End Property

Public Property Get iPadraoJuros() As Integer
    iPadraoJuros = mvariPadraoJuros
End Property

Public Property Let iPadraoTaxaFin(ByVal vData As Integer)
    mvariPadraoTaxaFin = vData
End Property

Public Property Get iPadraoTaxaFin() As Integer
    iPadraoTaxaFin = mvariPadraoTaxaFin
End Property

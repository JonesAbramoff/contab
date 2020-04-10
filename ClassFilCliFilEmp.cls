VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassFilCliFilEmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariFilialEmpresa As Integer
Private mvarlCodCliente As Long
Private mvariCodFilial As Integer
Private mvarlNumeroCompras As Long
Private mvardMediaCompra As Double
Private mvardtDataPrimeiraCompra As Date
Private mvardtDataUltimaCompra As Date
Private mvardValorAcumuladoCompras As Double


Public Property Let dValorAcumuladoCompras(ByVal vData As Double)
    mvardValorAcumuladoCompras = vData
End Property

Public Property Get dValorAcumuladoCompras() As Double
    dValorAcumuladoCompras = mvardValorAcumuladoCompras
End Property


Public Property Let dtDataUltimaCompra(ByVal vData As Date)
    mvardtDataUltimaCompra = vData
End Property

Public Property Get dtDataUltimaCompra() As Date
    dtDataUltimaCompra = mvardtDataUltimaCompra
End Property

Public Property Let dtDataPrimeiraCompra(ByVal vData As Date)
    mvardtDataPrimeiraCompra = vData
End Property

Public Property Get dtDataPrimeiraCompra() As Date
    dtDataPrimeiraCompra = mvardtDataPrimeiraCompra
End Property

Public Property Let dMediaCompra(ByVal vData As Double)
    mvardMediaCompra = vData
End Property

Public Property Get dMediaCompra() As Double
    dMediaCompra = mvardMediaCompra
End Property

Public Property Let lNumeroCompras(ByVal vData As Long)
    mvarlNumeroCompras = vData
End Property

Public Property Get lNumeroCompras() As Long
    lNumeroCompras = mvarlNumeroCompras
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


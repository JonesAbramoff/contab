VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelRCPEDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarobjItemMovEst As Object
Private mvarobjTipoMovEst As Object
Private mvarobjNFiscal As Object
Private mvarobjItemNF As Object
Private mvariQuebra As Integer
Private mvariCodigo As Integer
Private mvariAno As Integer
Private mvariMes As Integer
Private mvariDia As Integer
Private mvarsEntradaouSaida As String
Private mvardTotalEntrada As Double
Private mvardTotalSaida As Double

Public Property Let dTotalSaida(ByVal vData As Double)
    mvardTotalSaida = vData
End Property

Public Property Get dTotalSaida() As Double
    dTotalSaida = mvardTotalSaida
End Property

Public Property Let dTotalEntrada(ByVal vData As Double)
    mvardTotalEntrada = vData
End Property

Public Property Get dTotalEntrada() As Double
    dTotalEntrada = mvardTotalEntrada
End Property

Public Property Let sEntradaouSaida(ByVal vData As String)
    mvarsEntradaouSaida = vData
End Property

Public Property Get sEntradaouSaida() As String
    sEntradaouSaida = mvarsEntradaouSaida
End Property

Public Property Let iDia(ByVal vData As Integer)
    mvariDia = vData
End Property

Public Property Get iDia() As Integer
    iDia = mvariDia
End Property

Public Property Let iMes(ByVal vData As Integer)
    mvariMes = vData
End Property

Public Property Get iMes() As Integer
    iMes = mvariMes
End Property

Public Property Let iAno(ByVal vData As Integer)
    mvariAno = vData
End Property

Public Property Get iAno() As Integer
    iAno = mvariAno
End Property

Public Property Let iCodigo(ByVal vData As Integer)
    mvariCodigo = vData
End Property

Public Property Get iCodigo() As Integer
    iCodigo = mvariCodigo
End Property

Public Property Let iQuebra(ByVal vData As Integer)
    mvariQuebra = vData
End Property

Public Property Get iQuebra() As Integer
    iQuebra = mvariQuebra
End Property

Public Property Set objItemNF(ByVal vData As Object)
    Set mvarobjItemNF = vData
End Property

Public Property Get objItemNF() As Object
    Set objItemNF = mvarobjItemNF
End Property

Public Property Set objNFiscal(ByVal vData As Object)
    Set mvarobjNFiscal = vData
End Property

Public Property Get objNFiscal() As Object
    Set objNFiscal = mvarobjNFiscal
End Property

Public Property Set objTipoMovEst(ByVal vData As Object)
    Set mvarobjTipoMovEst = vData
End Property

Public Property Get objTipoMovEst() As Object
    Set objTipoMovEst = mvarobjTipoMovEst
End Property

Public Property Set objItemMovEst(ByVal vData As Object)
    Set mvarobjItemMovEst = vData
End Property

Public Property Get objItemMovEst() As Object
    Set objItemMovEst = mvarobjItemMovEst
End Property




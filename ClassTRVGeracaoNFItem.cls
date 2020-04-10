VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTRVGeracaoNFItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvariFilialEmpresa As Integer
Private mvarsFilialEmpresa As String
Private mvariNumNF As Integer
Private mvariNumTitulos As Integer
Private mvardValor As Double
Private mvardValorR As Double
Private mvarlNFDe As Long
Private mvarlNFAte As Long
Private mvarlNFProx As Long

Public Property Let lNFProx(ByVal vData As Long)
    mvarlNFProx = vData
End Property

Public Property Get lNFProx() As Long
    lNFProx = mvarlNFProx
End Property

Public Property Let lNFDe(ByVal vData As Long)
    mvarlNFDe = vData
End Property

Public Property Get lNFDe() As Long
    lNFDe = mvarlNFDe
End Property

Public Property Let lNFAte(ByVal vData As Long)
    mvarlNFAte = vData
End Property

Public Property Get lNFAte() As Long
    lNFAte = mvarlNFAte
End Property

Public Property Let dValor(ByVal vData As Double)
    mvardValor = vData
End Property

Public Property Get dValor() As Double
    dValor = mvardValor
End Property

Public Property Let dValorR(ByVal vData As Double)
    mvardValorR = vData
End Property

Public Property Get dValorR() As Double
    dValorR = mvardValorR
End Property

Public Property Let iNumTitulos(ByVal vData As Integer)
    mvariNumTitulos = vData
End Property

Public Property Get iNumTitulos() As Integer
    iNumTitulos = mvariNumTitulos
End Property

Public Property Let iNumNF(ByVal vData As Integer)
    mvariNumNF = vData
End Property

Public Property Get iNumNF() As Integer
    iNumNF = mvariNumNF
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let sFilialEmpresa(ByVal vData As String)
    mvarsFilialEmpresa = vData
End Property

Public Property Get sFilialEmpresa() As String
    sFilialEmpresa = mvarsFilialEmpresa
End Property

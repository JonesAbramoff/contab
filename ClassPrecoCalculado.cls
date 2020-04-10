VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPrecoCalculado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarsCodProduto As String
Private mvariFilialEmpresa As Integer
Private mvardtDataReferencia As Date
Private mvardPrecoCalculado As Double
Private mvardPrecoInformado As Double
Private mvardtDataVigencia As Date
Private mvariCodTabela As Integer

Public Property Let dPrecoCalculado(ByVal vData As Double)
    mvardPrecoCalculado = vData
End Property

Public Property Get dPrecoCalculado() As Double
    dPrecoCalculado = mvardPrecoCalculado
End Property

Public Property Let dPrecoInformado(ByVal vData As Double)
    mvardPrecoInformado = vData
End Property

Public Property Get dPrecoInformado() As Double
    dPrecoInformado = mvardPrecoInformado
End Property

Public Property Let dtDataReferencia(ByVal vData As Date)
    mvardtDataReferencia = vData
End Property

Public Property Get dtDataReferencia() As Date
    dtDataReferencia = mvardtDataReferencia
End Property

Public Property Let sCodProduto(ByVal vData As String)
    mvarsCodProduto = vData
End Property

Public Property Get sCodProduto() As String
    sCodProduto = mvarsCodProduto
End Property

Public Property Let dtDataVigencia(ByVal vData As Date)
    mvardtDataVigencia = vData
End Property

Public Property Get dtDataVigencia() As Date
    dtDataVigencia = mvardtDataVigencia
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property


Public Property Let iCodTabela(ByVal vData As Integer)
    mvariCodTabela = vData
End Property

Public Property Get iCodTabela() As Integer
    iCodTabela = mvariCodTabela
End Property


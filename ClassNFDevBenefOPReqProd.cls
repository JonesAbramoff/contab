VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassNFDevBenefReqProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarsProduto As String
Private mvarsUM As String
Private mvardQuantReqProd As Double
Private mvardtDataReqProd As Date
Private mvarlNumIntMovto As Long
Private mvarsOP As String
Private mvarsLoteOP As String
Private mvarlNumIntItemOP As Long
Private mvarlNumIntItemOPPrinc As Long
Private mvardFatorProd As Double

Public Property Let dFatorProd(ByVal vData As Double)
    mvardFatorProd = vData
End Property

Public Property Get dFatorProd() As Double
    dFatorProd = mvardFatorProd
End Property

Public Property Let lNumIntItemOPPrinc(ByVal vData As Long)
    mvarlNumIntItemOPPrinc = vData
End Property

Public Property Get lNumIntItemOPPrinc() As Long
    lNumIntItemOPPrinc = mvarlNumIntItemOPPrinc
End Property

Public Property Let lNumIntItemOP(ByVal vData As Long)
    mvarlNumIntItemOP = vData
End Property

Public Property Get lNumIntItemOP() As Long
    lNumIntItemOP = mvarlNumIntItemOP
End Property

Public Property Let sLoteOP(ByVal vData As String)
    mvarsLoteOP = vData
End Property

Public Property Get sLoteOP() As String
    sLoteOP = mvarsLoteOP
End Property

Public Property Let sOP(ByVal vData As String)
    mvarsOP = vData
End Property

Public Property Get sOP() As String
    sOP = mvarsOP
End Property

Public Property Let dQuantReqProd(ByVal vData As Double)
    mvardQuantReqProd = vData
End Property

Public Property Get dQuantReqProd() As Double
    dQuantReqProd = mvardQuantReqProd
End Property

Public Property Let lNumIntMovto(ByVal vData As Long)
    mvarlNumIntMovto = vData
End Property

Public Property Get lNumIntMovto() As Long
    lNumIntMovto = mvarlNumIntMovto
End Property

Public Property Let dtDataReqProd(ByVal vData As Date)
    mvardtDataReqProd = vData
End Property

Public Property Get dtDataReqProd() As Date
    dtDataReqProd = mvardtDataReqProd
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let sUM(ByVal vData As String)
    mvarsUM = vData
End Property

Public Property Get sUM() As String
    sUM = mvarsUM
End Property

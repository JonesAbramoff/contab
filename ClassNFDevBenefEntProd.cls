VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassNFDevBenefEntProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarsOP As String
Private mvarsProdutoOP As String
Private mvarsUM As String
Private mvardQuantEntrada As Double
Private mvardtDataEntrada As Date
Private mvarlNumIntMovto As Long
Private mvariAlmoxarifado As Integer
Private mvarlNumIntItemOP As Long

Public Property Let lNumIntItemOP(ByVal vData As Long)
    mvarlNumIntItemOP = vData
End Property

Public Property Get lNumIntItemOP() As Long
    lNumIntItemOP = mvarlNumIntItemOP
End Property

Public Property Let iAlmoxarifado(ByVal vData As Integer)
    mvariAlmoxarifado = vData
End Property

Public Property Get iAlmoxarifado() As Integer
    iAlmoxarifado = mvariAlmoxarifado
End Property

Public Property Let dQuantEntrada(ByVal vData As Double)
    mvardQuantEntrada = vData
End Property

Public Property Get dQuantEntrada() As Double
    dQuantEntrada = mvardQuantEntrada
End Property

Public Property Let lNumIntMovto(ByVal vData As Long)
    mvarlNumIntMovto = vData
End Property

Public Property Get lNumIntMovto() As Long
    lNumIntMovto = mvarlNumIntMovto
End Property

Public Property Let dtDataEntrada(ByVal vData As Date)
    mvardtDataEntrada = vData
End Property

Public Property Get dtDataEntrada() As Date
    dtDataEntrada = mvardtDataEntrada
End Property

Public Property Let sProdutoOP(ByVal vData As String)
    mvarsProdutoOP = vData
End Property

Public Property Get sProdutoOP() As String
    sProdutoOP = mvarsProdutoOP
End Property

Public Property Let sUM(ByVal vData As String)
    mvarsUM = vData
End Property

Public Property Get sUM() As String
    sUM = mvarsUM
End Property

Public Property Let sOP(ByVal vData As String)
    mvarsOP = vData
End Property

Public Property Get sOP() As String
    sOP = mvarsOP
End Property


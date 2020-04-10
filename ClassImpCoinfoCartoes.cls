VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassImpCoinfoCartoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarsCiaCart As String
Private mvardtDataDep As Date
Private mvariQuantParc As Integer
Private mvardValor As Double
Private mvarlNumIntTitulo As Long
Private mvarlNumTitulo As Long
Private mvarcolItens As New Collection
Private mvariFilialEmpresa As Integer
Private mvarsNumAuto As String
Private mvarsValidadeCC As String
Private mvarsNumCredCart As String

Public dValorTaxa As Double
Public dDeducaoValorPresente As Double
Public dValorTitulo As Double
Private mvarlCliente As Long

Public Property Let lCliente(ByVal vData As Long)
    mvarlCliente = vData
End Property

Public Property Get lCliente() As Long
    lCliente = mvarlCliente
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Set colItens(ByVal vData As Collection)
    Set mvarcolItens = vData
End Property

Public Property Get colItens() As Collection
    Set colItens = mvarcolItens
End Property

Public Property Let lNumIntTitulo(ByVal vData As Long)
    mvarlNumIntTitulo = vData
End Property

Public Property Get lNumIntTitulo() As Long
    lNumIntTitulo = mvarlNumIntTitulo
End Property

Public Property Let lNumTitulo(ByVal vData As Long)
    mvarlNumTitulo = vData
End Property

Public Property Get lNumTitulo() As Long
    lNumTitulo = mvarlNumTitulo
End Property

Public Property Let dValor(ByVal vData As Double)
    mvardValor = vData
End Property

Public Property Get dValor() As Double
    dValor = mvardValor
End Property

Public Property Let iQuantParc(ByVal vData As Integer)
    mvariQuantParc = vData
End Property

Public Property Get iQuantParc() As Integer
    iQuantParc = mvariQuantParc
End Property

Public Property Let dtDataDep(ByVal vData As Date)
    mvardtDataDep = vData
End Property

Public Property Get dtDataDep() As Date
    dtDataDep = mvardtDataDep
End Property

Public Property Let sCiaCart(ByVal vData As String)
    mvarsCiaCart = vData
End Property

Public Property Get sCiaCart() As String
    sCiaCart = mvarsCiaCart
End Property

Public Property Let sNumAuto(ByVal vData As String)
    mvarsNumAuto = vData
End Property

Public Property Get sNumAuto() As String
    sNumAuto = mvarsNumAuto
End Property

Public Property Let sValidadeCC(ByVal vData As String)
    mvarsValidadeCC = vData
End Property

Public Property Get sValidadeCC() As String
    sValidadeCC = mvarsValidadeCC
End Property

Public Property Let sNumCredCart(ByVal vData As String)
    mvarsNumCredCart = vData
End Property

Public Property Get sNumCredCart() As String
    sNumCredCart = mvarsNumCredCart
End Property

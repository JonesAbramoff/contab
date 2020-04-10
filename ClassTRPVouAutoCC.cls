VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTRPVouAutoCC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvardtDataAutoCC As Date
Private mvarsValidadeCC As String
Private mvarsBandeira As String
Private mvarsNumCCred As String
Private mvarlNumAuto As Long
Private mvariQuantParc As Integer
Private mvardValorB As Double
Private mvardValorL As Double
Private mvardTarifa As Double
Private mvarsTitular As String
Private mvarsTitularCPF As String
Private mvarsClienteFat As String
Private mvariCodSeguranca As Integer
Private mvarlClienteFat As Long
Private mvarcolVou As New Collection
Private mvarlNumTitulo As Long
Private mvarlNumIntTitulo As Long
Private mvardValorTarifa As Double
Private mvardValorDeducoes As Double
Private mvariFilialEmpresa As Integer

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let lNumTitulo(ByVal vData As Long)
    mvarlNumTitulo = vData
End Property

Public Property Get lNumTitulo() As Long
    lNumTitulo = mvarlNumTitulo
End Property

Public Property Let lNumIntTitulo(ByVal vData As Long)
    mvarlNumIntTitulo = vData
End Property

Public Property Get lNumIntTitulo() As Long
    lNumIntTitulo = mvarlNumIntTitulo
End Property

Public Property Let dValorTarifa(ByVal vData As Double)
    mvardValorTarifa = vData
End Property

Public Property Get dValorTarifa() As Double
    dValorTarifa = mvardValorTarifa
End Property

Public Property Let dValorDeducoes(ByVal vData As Double)
    mvardValorDeducoes = vData
End Property

Public Property Get dValorDeducoes() As Double
    dValorDeducoes = mvardValorDeducoes
End Property

Public Property Set colVou(ByVal vData As Collection)
    Set mvarcolVou = vData
End Property

Public Property Get colVou() As Collection
    Set colVou = mvarcolVou
End Property

Public Property Let sValidadeCC(ByVal vData As String)
    mvarsValidadeCC = vData
End Property

Public Property Get sValidadeCC() As String
    sValidadeCC = mvarsValidadeCC
End Property

Public Property Let dtDataAutoCC(ByVal vDataUltima As Date)
    mvardtDataAutoCC = vDataUltima
End Property

Public Property Get dtDataAutoCC() As Date
    dtDataAutoCC = mvardtDataAutoCC
End Property

Public Property Let lClienteFat(ByVal vData As Long)
    mvarlClienteFat = vData
End Property

Public Property Get lClienteFat() As Long
    lClienteFat = mvarlClienteFat
End Property

Public Property Let sTitular(ByVal vData As String)
    mvarsTitular = vData
End Property

Public Property Get sTitular() As String
    sTitular = mvarsTitular
End Property

Public Property Let sTitularCPF(ByVal vData As String)
    mvarsTitularCPF = vData
End Property

Public Property Get sTitularCPF() As String
    sTitularCPF = mvarsTitularCPF
End Property

Public Property Let sBandeira(ByVal vData As String)
    mvarsBandeira = vData
End Property

Public Property Get sBandeira() As String
    sBandeira = mvarsBandeira
End Property

Public Property Let sNumCCred(ByVal vData As String)
    mvarsNumCCred = vData
End Property

Public Property Get sNumCCred() As String
    sNumCCred = mvarsNumCCred
End Property

Public Property Let sClienteFat(ByVal vData As String)
    mvarsClienteFat = vData
End Property

Public Property Get sClienteFat() As String
    sClienteFat = mvarsClienteFat
End Property

Public Property Let lNumAuto(ByVal vData As Long)
    mvarlNumAuto = vData
End Property

Public Property Get lNumAuto() As Long
    lNumAuto = mvarlNumAuto
End Property

Public Property Let iQuantParc(ByVal vData As Integer)
    mvariQuantParc = vData
End Property

Public Property Get iQuantParc() As Integer
    iQuantParc = mvariQuantParc
End Property

Public Property Let iCodSeguranca(ByVal vData As Integer)
    mvariCodSeguranca = vData
End Property

Public Property Get iCodSeguranca() As Integer
    iCodSeguranca = mvariCodSeguranca
End Property

Public Property Let dValorL(ByVal vData As Double)
    mvardValorL = vData
End Property

Public Property Get dValorL() As Double
    dValorL = mvardValorL
End Property

Public Property Let dValorB(ByVal vData As Double)
    mvardValorB = vData
End Property

Public Property Get dValorB() As Double
    dValorB = mvardValorB
End Property

Public Property Let dTarifa(ByVal vData As Double)
    mvardTarifa = vData
End Property

Public Property Get dTarifa() As Double
    dTarifa = mvardTarifa
End Property

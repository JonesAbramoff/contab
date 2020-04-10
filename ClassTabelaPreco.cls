VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTabelaPreco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarsDescricao As String
Private mvariCodigo As Integer
Private mvariAtivo As Integer
Private mvardtDataLog As Date
Private mvariAjustaTabVlrCompCoefMaior As Integer
Private mvardVlrCompCoef As Double
Private mvardMarkUp As Double
Private mvardAliquotaICMS As Double

Private mvariCargoMinimo As Integer
Private mvariMoeda As Integer

Public dComissao As Double

Public iTipo As Integer
Public iDescontoLimitado As Integer
Public dDescontoMaximo As Double

Public Property Let iMoeda(ByVal vData As Integer)
    mvariMoeda = vData
End Property

Public Property Get iMoeda() As Integer
    iMoeda = mvariMoeda
End Property

Public Property Let iCargoMinimo(ByVal vData As Integer)
    mvariCargoMinimo = vData
End Property

Public Property Get iCargoMinimo() As Integer
    iCargoMinimo = mvariCargoMinimo
End Property

Public Property Let dAliquotaICMS(ByVal vData As Double)
    mvardAliquotaICMS = vData
End Property

Public Property Get dAliquotaICMS() As Double
    dAliquotaICMS = mvardAliquotaICMS
End Property

Public Property Let dMarkUp(ByVal vData As Double)
    mvardMarkUp = vData
End Property

Public Property Get dMarkUp() As Double
    dMarkUp = mvardMarkUp
End Property

Public Property Let dVlrCompCoef(ByVal vData As Double)
    mvardVlrCompCoef = vData
End Property

Public Property Get dVlrCompCoef() As Double
    dVlrCompCoef = mvardVlrCompCoef
End Property

Public Property Let iAjustaTabVlrCompCoefMaior(ByVal vData As Integer)
    mvariAjustaTabVlrCompCoefMaior = vData
End Property

Public Property Get iAjustaTabVlrCompCoefMaior() As Integer
    iAjustaTabVlrCompCoefMaior = mvariAjustaTabVlrCompCoefMaior
End Property

Public Property Let dtDataLog(ByVal vData As Date)
    mvardtDataLog = vData
End Property

Public Property Get dtDataLog() As Date
    dtDataLog = mvardtDataLog
End Property

Public Property Let iAtivo(ByVal vData As Integer)
    mvariAtivo = vData
End Property

Public Property Get iAtivo() As Integer
    iAtivo = mvariAtivo
End Property

Public Property Let iCodigo(ByVal vData As Integer)
    mvariCodigo = vData
End Property

Public Property Get iCodigo() As Integer
    iCodigo = mvariCodigo
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPRJRecebPagtoRegras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDoc As Long
Private mvarlNumIntDocRecebPagto As Long
Private mvarsRegra As String
Private mvardPercentual As Double
Private mvariCondPagto As Integer
Private mvarsObservacao As String
Private mvardtRegraValor As Date

'PROPERTY AUXILIARES PARA CALCULOS
Private mvariTipo As Integer
Private mvardValor As Double
Private mvarlNumero As Long
Private mvarcolParcelas As New Collection
Private mvariParcela As Integer
Private mvardValorParcela As Double
Private mvardtDataVencimentoParcela As Date
Private mvariTipoNegativo As Integer

Public Property Let iTipoNegativo(ByVal vData As Integer)
    mvariTipoNegativo = vData
End Property

Public Property Get iTipoNegativo() As Integer
    iTipoNegativo = mvariTipoNegativo
End Property

Public Property Let dtDataVencimentoParcela(ByVal vData As Date)
    mvardtDataVencimentoParcela = vData
End Property

Public Property Get dtDataVencimentoParcela() As Date
    dtDataVencimentoParcela = mvardtDataVencimentoParcela
End Property

Public Property Let dValorParcela(ByVal vData As Double)
    mvardValorParcela = vData
End Property

Public Property Get dValorParcela() As Double
    dValorParcela = mvardValorParcela
End Property

Public Property Let iParcela(ByVal vData As Integer)
    mvariParcela = vData
End Property

Public Property Get iParcela() As Integer
    iParcela = mvariParcela
End Property

Public Property Let lNumero(ByVal vData As Long)
    mvarlNumero = vData
End Property

Public Property Get lNumero() As Long
    lNumero = mvarlNumero
End Property

Public Property Set colParcelas(ByVal vData As Collection)
    Set mvarcolParcelas = vData
End Property

Public Property Get colParcelas() As Collection
    Set colParcelas = mvarcolParcelas
End Property

Public Property Let dValor(ByVal vData As Double)
    mvardValor = vData
End Property

Public Property Get dValor() As Double
    dValor = mvardValor
End Property

Public Property Let iTipo(ByVal vData As Integer)
    mvariTipo = vData
End Property

Public Property Get iTipo() As Integer
    iTipo = mvariTipo
End Property
'PROPERTY AUXILIARES PARA CALCULOS

Public Property Let dtRegraValor(ByVal vData As Date)
    mvardtRegraValor = vData
End Property

Public Property Get dtRegraValor() As Date
    dtRegraValor = mvardtRegraValor
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let lNumIntDocRecebPagto(ByVal vData As Long)
    mvarlNumIntDocRecebPagto = vData
End Property

Public Property Get lNumIntDocRecebPagto() As Long
    lNumIntDocRecebPagto = mvarlNumIntDocRecebPagto
End Property

Public Property Let sRegra(ByVal vData As String)
    mvarsRegra = vData
End Property

Public Property Get sRegra() As String
    sRegra = mvarsRegra
End Property

Public Property Let dPercentual(ByVal vData As Double)
    mvardPercentual = vData
End Property

Public Property Get dPercentual() As Double
    dPercentual = mvardPercentual
End Property

Public Property Let iCondPagto(ByVal vData As Integer)
    mvariCondPagto = vData
End Property

Public Property Get iCondPagto() As Integer
    iCondPagto = mvariCondPagto
End Property

Public Property Let sObservacao(ByVal vData As String)
    mvarsObservacao = vData
End Property

Public Property Get sObservacao() As String
    sObservacao = mvarsObservacao
End Property

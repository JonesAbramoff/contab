VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassGeracaoArqICMSAux"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

Private mvarsNatOperacao As String
Private mvardAliquota As Double
Private mvardValorTotal As Double
Private mvardBaseICMS As Double
Private mvardValorICMS As Double
Private mvardIsentoNaoTribICMS As Double
Private mvardOutrasICMS As Double
Private mvardValorIPI As Double
Private mvardIsentoNaoTribIPI As Double
Private mvardOutrasIPI As Double

Public Property Let dOutrasIPI(ByVal vData As Double)
    mvardOutrasIPI = vData
End Property

Public Property Get dOutrasIPI() As Double
    dOutrasIPI = mvardOutrasIPI
End Property

Public Property Let dIsentoNaoTribIPI(ByVal vData As Double)
    mvardIsentoNaoTribIPI = vData
End Property

Public Property Get dIsentoNaoTribIPI() As Double
    dIsentoNaoTribIPI = mvardIsentoNaoTribIPI
End Property

Public Property Let dValorIPI(ByVal vData As Double)
    mvardValorIPI = vData
End Property

Public Property Get dValorIPI() As Double
    dValorIPI = mvardValorIPI
End Property

Public Property Let dOutrasICMS(ByVal vData As Double)
    mvardOutrasICMS = vData
End Property

Public Property Get dOutrasICMS() As Double
    dOutrasICMS = mvardOutrasICMS
End Property

Public Property Let dIsentoNaoTribICMS(ByVal vData As Double)
    mvardIsentoNaoTribICMS = vData
End Property

Public Property Get dIsentoNaoTribICMS() As Double
    dIsentoNaoTribICMS = mvardIsentoNaoTribICMS
End Property

Public Property Let dValorICMS(ByVal vData As Double)
    mvardValorICMS = vData
End Property

Public Property Get dValorICMS() As Double
    dValorICMS = mvardValorICMS
End Property

Public Property Let dBaseICMS(ByVal vData As Double)
    mvardBaseICMS = vData
End Property

Public Property Get dBaseICMS() As Double
    dBaseICMS = mvardBaseICMS
End Property

Public Property Let dValorTotal(ByVal vData As Double)
    mvardValorTotal = vData
End Property

Public Property Get dValorTotal() As Double
    dValorTotal = mvardValorTotal
End Property

Public Property Let dAliquota(ByVal vData As Double)
    mvardAliquota = vData
End Property

Public Property Get dAliquota() As Double
    dAliquota = mvardAliquota
End Property

Public Property Let sNatOperacao(ByVal vData As String)
    mvarsNatOperacao = vData
End Property

Public Property Get sNatOperacao() As String
    sNatOperacao = mvarsNatOperacao
End Property

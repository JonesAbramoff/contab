VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassProcReajTitRecTit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDocProc As Long
Private mvarlNumIntTitulo As Long
Private mvardSaldoAnterior As Double
Private mvardValorReajuste As Double
Private mvarlTipoCritica As Long

Public Property Let lTipoCritica(ByVal vData As Long)
    mvarlTipoCritica = vData
End Property

Public Property Get lTipoCritica() As Long
    lTipoCritica = mvarlTipoCritica
End Property

Public Property Get dSaldoReajustado() As Double
    dSaldoReajustado = Arredonda_Moeda(mvardSaldoAnterior + mvardValorReajuste)
End Property

Public Property Let lNumIntDocProc(ByVal vData As Long)
    mvarlNumIntDocProc = vData
End Property

Public Property Get lNumIntDocProc() As Long
    lNumIntDocProc = mvarlNumIntDocProc
End Property

Public Property Let lNumIntTitulo(ByVal vData As Long)
    mvarlNumIntTitulo = vData
End Property

Public Property Get lNumIntTitulo() As Long
    lNumIntTitulo = mvarlNumIntTitulo
End Property

Public Property Let dSaldoAnterior(ByVal vData As Double)
    mvardSaldoAnterior = vData
End Property

Public Property Get dSaldoAnterior() As Double
    dSaldoAnterior = mvardSaldoAnterior
End Property

Public Property Let dValorReajuste(ByVal vData As Double)
    mvardValorReajuste = vData
End Property

Public Property Get dValorReajuste() As Double
    dValorReajuste = mvardValorReajuste
End Property

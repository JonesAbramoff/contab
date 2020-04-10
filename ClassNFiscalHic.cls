VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassNFiscalHic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvardValorFrete1 As Double
Private mvardValorSeguro1 As Double
Private mvardValorOutrasDespesas1 As Double
Private mvardValorFrete2 As Double
Private mvardValorSeguro2 As Double
Private mvardValorOutrasDespesas2 As Double
Private mvariFlagCompl1 As Integer
Private mvariFlagCompl2 As Integer

'#################################
'Inserido por Wagner
Private mvarlHistorico As Long
Private mvarlSubConta As Long


Public Property Let lHistorico(ByVal vData As Long)
    mvarlHistorico = vData
End Property

Public Property Get lHistorico() As Long
    lHistorico = mvarlHistorico
End Property

Public Property Let lSubConta(ByVal vData As Long)
    mvarlSubConta = vData
End Property

Public Property Get lSubConta() As Long
    lSubConta = mvarlSubConta
End Property
'#################################

Public Property Let dValorFrete1(ByVal vData As Double)
    mvardValorFrete1 = vData
End Property

Public Property Get dValorFrete1() As Double
    dValorFrete1 = mvardValorFrete1
End Property

Public Property Let dValorSeguro1(ByVal vData As Double)
    mvardValorSeguro1 = vData
End Property

Public Property Get dValorSeguro1() As Double
    dValorSeguro1 = mvardValorSeguro1
End Property

Public Property Let dValorOutrasDespesas1(ByVal vData As Double)
    mvardValorOutrasDespesas1 = vData
End Property

Public Property Get dValorOutrasDespesas1() As Double
    dValorOutrasDespesas1 = mvardValorOutrasDespesas1
End Property

Public Property Let iFlagCompl1(ByVal vData As Integer)
    mvariFlagCompl1 = vData
End Property

Public Property Get iFlagCompl1() As Integer
    iFlagCompl1 = mvariFlagCompl1
End Property

Public Property Let dValorFrete2(ByVal vData As Double)
    mvardValorFrete2 = vData
End Property

Public Property Get dValorFrete2() As Double
    dValorFrete2 = mvardValorFrete2
End Property

Public Property Let dValorSeguro2(ByVal vData As Double)
    mvardValorSeguro2 = vData
End Property

Public Property Get dValorSeguro2() As Double
    dValorSeguro2 = mvardValorSeguro2
End Property

Public Property Let dValorOutrasDespesas2(ByVal vData As Double)
    mvardValorOutrasDespesas2 = vData
End Property

Public Property Get dValorOutrasDespesas2() As Double
    dValorOutrasDespesas2 = mvardValorOutrasDespesas2
End Property

Public Property Let iFlagCompl2(ByVal vData As Integer)
    mvariFlagCompl2 = vData
End Property

Public Property Get iFlagCompl2() As Integer
    iFlagCompl2 = mvariFlagCompl2
End Property




VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassFluxoSint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlFluxoId As Long
Private mvardtData As Date
Private mvardRecValorSistema As Double
Private mvardRecValorAjustado As Double
Private mvardRecValorReal As Double
Private mvardPagValorSistema As Double
Private mvardPagValorAjustado As Double
Private mvardPagValorReal As Double
Private mvardTesValorSistema As Double
Private mvardTesValorAjustado As Double
Private mvardTesValorReal As Double
Private mvardSaldoValorSistema As Double
Private mvardSaldoValorAjustado As Double
Private mvardSaldoValorReal As Double

Public Property Let dSaldoValorReal(ByVal vData As Double)
Attribute dSaldoValorReal.VB_Description = "Saldo realizado"
    mvardSaldoValorReal = vData
End Property

Public Property Get dSaldoValorReal() As Double
    dSaldoValorReal = mvardSaldoValorReal
End Property

Public Property Let dSaldoValorAjustado(ByVal vData As Double)
Attribute dSaldoValorAjustado.VB_Description = "Saldo ajustado pelo Usuario"
    mvardSaldoValorAjustado = vData
End Property

Public Property Get dSaldoValorAjustado() As Double
    dSaldoValorAjustado = mvardSaldoValorAjustado
End Property

Public Property Let dSaldoValorSistema(ByVal vData As Double)
Attribute dSaldoValorSistema.VB_Description = "Saldo apurado pelo Sistema"
    mvardSaldoValorSistema = vData
End Property

Public Property Get dSaldoValorSistema() As Double
    dSaldoValorSistema = mvardSaldoValorSistema
End Property

Public Property Let dTesValorReal(ByVal vData As Double)
    mvardTesValorReal = vData
End Property

Public Property Get dTesValorReal() As Double
    dTesValorReal = mvardTesValorReal
End Property

Public Property Let dTesValorAjustado(ByVal vData As Double)
    mvardTesValorAjustado = vData
End Property

Public Property Get dTesValorAjustado() As Double
    dTesValorAjustado = mvardTesValorAjustado
End Property

Public Property Let dTesValorSistema(ByVal vData As Double)
    mvardTesValorSistema = vData
End Property

Public Property Get dTesValorSistema() As Double
    dTesValorSistema = mvardTesValorSistema
End Property

Public Property Let dPagValorReal(ByVal vData As Double)
Attribute dPagValorReal.VB_Description = "Total de Pagamentos realizados"
    mvardPagValorReal = vData
End Property

Public Property Get dPagValorReal() As Double
    dPagValorReal = mvardPagValorReal
End Property

Public Property Let dPagValorAjustado(ByVal vData As Double)
Attribute dPagValorAjustado.VB_Description = "Total de Pagamentos ajustado pelo usuário"
    mvardPagValorAjustado = vData
End Property

Public Property Get dPagValorAjustado() As Double
    dPagValorAjustado = mvardPagValorAjustado
End Property

Public Property Let dPagValorSistema(ByVal vData As Double)
Attribute dPagValorSistema.VB_Description = "Total de Pagamentos apurados pelo Sistema"
    mvardPagValorSistema = vData
End Property

Public Property Get dPagValorSistema() As Double
    dPagValorSistema = mvardPagValorSistema
End Property

Public Property Let dRecValorReal(ByVal vData As Double)
Attribute dRecValorReal.VB_Description = "Total de Recebimentos realizados"
    mvardRecValorReal = vData
End Property

Public Property Get dRecValorReal() As Double
    dRecValorReal = mvardRecValorReal
End Property

Public Property Let dRecValorAjustado(ByVal vData As Double)
Attribute dRecValorAjustado.VB_Description = "Total de Recebimentos ajustado pelo usuário"
    mvardRecValorAjustado = vData
End Property

Public Property Get dRecValorAjustado() As Double
    dRecValorAjustado = mvardRecValorAjustado
End Property

Public Property Let dRecValorSistema(ByVal vData As Double)
Attribute dRecValorSistema.VB_Description = "Total de Recebimentos apurados pelo Sistema"
    mvardRecValorSistema = vData
End Property

Public Property Get dRecValorSistema() As Double
    dRecValorSistema = mvardRecValorSistema
End Property

Public Property Let dtData(ByVal vData As Date)
Attribute dtData.VB_Description = "Data de Apuração dos valores"
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let lFluxoId(ByVal vData As Long)
Attribute lFluxoId.VB_Description = "Identificador do Fluxo de Caixa"
    mvarlFluxoId = vData
End Property

Public Property Get lFluxoId() As Long
    lFluxoId = mvarlFluxoId
End Property




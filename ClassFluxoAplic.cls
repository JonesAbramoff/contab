VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassFluxoAplic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlFluxoId As Long
Private mvarlCodigo As Long
Private mvardtDataResgatePrevista As Date
Private mvardSaldoAplicado As Double
Private mvardValorResgatePrevisto As Double

Public Property Let dValorResgatePrevisto(ByVal vData As Double)
Attribute dValorResgatePrevisto.VB_Description = "Valor previsto para o resgate da aplicação."
    mvardValorResgatePrevisto = vData
End Property

Public Property Get dValorResgatePrevisto() As Double
    dValorResgatePrevisto = mvardValorResgatePrevisto
End Property

Public Property Let dSaldoAplicado(ByVal vData As Double)
Attribute dSaldoAplicado.VB_Description = "valor aplicado menos valor resgatado"
    mvardSaldoAplicado = vData
End Property

Public Property Get dSaldoAplicado() As Double
    dSaldoAplicado = mvardSaldoAplicado
End Property

Public Property Let dtDataResgatePrevista(ByVal vData As Date)
Attribute dtDataResgatePrevista.VB_Description = "Data prevista para o resgate da aplicação."
    mvardtDataResgatePrevista = vData
End Property

Public Property Get dtDataResgatePrevista() As Date
    dtDataResgatePrevista = mvardtDataResgatePrevista
End Property

Public Property Let lCodigo(ByVal vData As Long)
Attribute lCodigo.VB_Description = "Identificador da aplicação"
    mvarlCodigo = vData
End Property

Public Property Get lCodigo() As Long
    lCodigo = mvarlCodigo
End Property

Public Property Let lFluxoId(ByVal vData As Long)
Attribute lFluxoId.VB_Description = "identificador do fluxo de caixa"
    mvarlFluxoId = vData
End Property

Public Property Get lFluxoId() As Long
    lFluxoId = mvarlFluxoId
End Property




VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassFluxoTipoAplic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlFluxoId As Long
Private mvariTipoAplicacao As Integer
Private mvarsDescricao As String
Private mvardtData As Date
Private mvardTotalSistema As Double
Private mvardTotalAjustado As Double
Private mvardTotalReal As Double
Private mvariUsuario As Date


Public Property Let dTotalReal(ByVal vData As Double)
Attribute dTotalReal.VB_Description = "valor total que realmente ocorreu"
    mvardTotalReal = vData
End Property

Public Property Get dTotalReal() As Double
    dTotalReal = mvardTotalReal
End Property

Public Property Let dTotalAjustado(ByVal vData As Double)
Attribute dTotalAjustado.VB_Description = "valor total ajustado pelo usuario"
    mvardTotalAjustado = vData
End Property

Public Property Get dTotalAjustado() As Double
    dTotalAjustado = mvardTotalAjustado
End Property

Public Property Let dTotalSistema(ByVal vData As Double)
Attribute dTotalSistema.VB_Description = "valor total projetado pelo sistema"
    mvardTotalSistema = vData
End Property

Public Property Get dTotalSistema() As Double
    dTotalSistema = mvardTotalSistema
End Property

Public Property Let dtData(ByVal vData As Date)
Attribute dtData.VB_Description = "titulos desta data estão sendo acumulados"
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let iTipoAplicacao(ByVal vData As Integer)
Attribute iTipoAplicacao.VB_Description = "identificador do Tipo de Aplicação."
    mvariTipoAplicacao = vData
End Property

Public Property Get iTipoAplicacao() As Integer
    iTipoAplicacao = mvariTipoAplicacao
End Property

Public Property Let lFluxoId(ByVal vData As Long)
Attribute lFluxoId.VB_Description = "Identicador do Fluxo de Caixa"
    mvarlFluxoId = vData
End Property

Public Property Get lFluxoId() As Long
    lFluxoId = mvarlFluxoId
End Property
Public Property Let iUsuario(ByVal vData As Integer)
    mvariUsuario = vData
End Property

Public Property Get iUsuario() As Integer
    iUsuario = mvariUsuario
End Property
Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property




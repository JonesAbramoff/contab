VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassFluxoForn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlFluxoId As Long
Private mvariTipoReg As Integer
Private mvarlFornecedor As Long
Private mvardtData As Date
Private mvardTotalSistema As Double
Private mvardTotalAjustado As Double
Private mvardTotalReal As Double
Private mvariTipoFornecedor As Integer
Private mvarsNomeReduzido As String
Private mvariUsuario As Integer

Public Property Let iUsuario(ByVal vData As Integer)
Attribute iUsuario.VB_Description = "0=foi gerado pelo sistema 1=foi criado pelo usuário"
    mvariUsuario = vData
End Property

Public Property Get iUsuario() As Integer
    iUsuario = mvariUsuario
End Property

Public Property Let sNomeReduzido(ByVal vData As String)
Attribute sNomeReduzido.VB_Description = "Nome Reduzido do Fornecedor ou Cliente"
    mvarsNomeReduzido = vData
End Property

Public Property Get sNomeReduzido() As String
    sNomeReduzido = mvarsNomeReduzido
End Property

Public Property Let iTipoFornecedor(ByVal vData As Integer)
Attribute iTipoFornecedor.VB_Description = "ou tipo de cliente ou tipo de aplicacao"
    mvariTipoFornecedor = vData
End Property

Public Property Get iTipoFornecedor() As Integer
    iTipoFornecedor = mvariTipoFornecedor
End Property

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

Public Property Let lFornecedor(ByVal vData As Long)
Attribute lFornecedor.VB_Description = "Fornecedor ou cliente ou cobrador ou cta"
    mvarlFornecedor = vData
End Property

Public Property Get lFornecedor() As Long
    lFornecedor = mvarlFornecedor
End Property

Public Property Let iTipoReg(ByVal vData As Integer)
Attribute iTipoReg.VB_Description = "0:forn/1:cliente/2:cobrador/3:pedidos de venda/4:pedidos de compra/5:cheques-pre"
    mvariTipoReg = vData
End Property

Public Property Get iTipoReg() As Integer
    iTipoReg = mvariTipoReg
End Property

Public Property Let lFluxoId(ByVal vData As Long)
Attribute lFluxoId.VB_Description = "Identicador do Fluxo de Caixa"
    mvarlFluxoId = vData
End Property

Public Property Get lFluxoId() As Long
    lFluxoId = mvarlFluxoId
End Property




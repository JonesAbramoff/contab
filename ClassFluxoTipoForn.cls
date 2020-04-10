VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassFluxoTipoForn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlFluxoId As Long
Private mvariTipoReg As Integer
Private mvariTipoFornecedor As Integer
Private mvardtData As Date
Private mvardTotalSistema As Double
Private mvardTotalAjustado As Double
Private mvardTotalReal As Double
Private mvarsDescricao As String

Public Property Let sDescricao(ByVal vData As String)
Attribute sDescricao.VB_Description = "Descricao do Tipo do Fornecedor/Cliente"
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let dTotalReal(ByVal vData As Double)
Attribute dTotalReal.VB_Description = "valor total realizado"
    mvardTotalReal = vData
End Property

Public Property Get dTotalReal() As Double
    dTotalReal = mvardTotalReal
End Property

Public Property Let dTotalAjustado(ByVal vData As Double)
Attribute dTotalAjustado.VB_Description = "valor total ajustado pelo usuário"
    mvardTotalAjustado = vData
End Property

Public Property Get dTotalAjustado() As Double
    dTotalAjustado = mvardTotalAjustado
End Property

Public Property Let dTotalSistema(ByVal vData As Double)
Attribute dTotalSistema.VB_Description = "valor total apurado pelo sistema"
    mvardTotalSistema = vData
End Property

Public Property Get dTotalSistema() As Double
    dTotalSistema = mvardTotalSistema
End Property

Public Property Let dtData(ByVal vData As Date)
Attribute dtData.VB_Description = "Data da Apuração dos valores"
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let iTipoFornecedor(ByVal vData As Integer)
Attribute iTipoFornecedor.VB_Description = "Tipo do Fornecedor ou tipo de cliente ou tipo de aplicacao"
    mvariTipoFornecedor = vData
End Property

Public Property Get iTipoFornecedor() As Integer
    iTipoFornecedor = mvariTipoFornecedor
End Property

Public Property Let iTipoReg(ByVal vData As Integer)
Attribute iTipoReg.VB_Description = "0:forn/1:cli/2:cobradores/3:aplicacoes"
    mvariTipoReg = vData
End Property

Public Property Get iTipoReg() As Integer
    iTipoReg = mvariTipoReg
End Property

Public Property Let lFluxoId(ByVal vData As Long)
Attribute lFluxoId.VB_Description = "Identificador do fluxo de caixa"
    mvarlFluxoId = vData
End Property

Public Property Get lFluxoId() As Long
    lFluxoId = mvarlFluxoId
End Property




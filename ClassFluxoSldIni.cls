VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassFluxoSldIni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlFluxoId As Long
Private mvarsNomeReduzido As String
Private mvariCodConta As Integer
Private mvariUsuario As Integer
Private mvardSaldoSistema As Double
Private mvardSaldoAjustado As Double
Private mvardSaldoReal As Double

Public Property Let dSaldoReal(ByVal vData As Double)
    mvardSaldoReal = vData
End Property

Public Property Get dSaldoReal() As Double
    dSaldoReal = mvardSaldoReal
End Property

Public Property Let dSaldoAjustado(ByVal vData As Double)
Attribute dSaldoAjustado.VB_Description = "saldo ajustado pelo usuário"
    mvardSaldoAjustado = vData
End Property

Public Property Get dSaldoAjustado() As Double
    dSaldoAjustado = mvardSaldoAjustado
End Property

Public Property Let dSaldoSistema(ByVal vData As Double)
Attribute dSaldoSistema.VB_Description = "saldo fornecido pelo sistema"
    mvardSaldoSistema = vData
End Property

Public Property Get dSaldoSistema() As Double
    dSaldoSistema = mvardSaldoSistema
End Property

Public Property Let iCodConta(ByVal vData As Integer)
Attribute iCodConta.VB_Description = "codigo da conta corrente interna"
    mvariCodConta = vData
End Property

Public Property Get iCodConta() As Integer
    iCodConta = mvariCodConta
End Property

Public Property Let sNomeReduzido(ByVal vData As String)
Attribute sNomeReduzido.VB_Description = "Nome Reduzido da Conta Corrente Interna"
    mvarsNomeReduzido = vData
End Property

Public Property Get sNomeReduzido() As String
    sNomeReduzido = mvarsNomeReduzido
End Property

Public Property Let lFluxoId(ByVal vData As Long)
Attribute lFluxoId.VB_Description = "Identificador do Fluxo"
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



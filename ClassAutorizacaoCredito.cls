VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassAutorizacaoCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlCliente As Long
Private mvardValor As Double
Private mvariCreditoAutorizado As Integer
Private mvarsCodUsuario As String

Public Property Let sCodUsuario(ByVal vData As String)
Attribute sCodUsuario.VB_Description = "Codigo do Usuario que autorizou o crédito (se preenchido)"
    mvarsCodUsuario = vData
End Property

Public Property Get sCodUsuario() As String
    sCodUsuario = mvarsCodUsuario
End Property

Public Property Let iCreditoAutorizado(ByVal vData As Integer)
Attribute iCreditoAutorizado.VB_Description = "1=Credito Autorizado 0=Em caso contrário"
    mvariCreditoAutorizado = vData
End Property

Public Property Get iCreditoAutorizado() As Integer
    iCreditoAutorizado = mvariCreditoAutorizado
End Property

Public Property Let dValor(ByVal vData As Double)
Attribute dValor.VB_Description = "Valor de credito solicitado"
    mvardValor = vData
End Property

Public Property Get dValor() As Double
    dValor = mvardValor
End Property

Public Property Let lCliente(ByVal vData As Long)
Attribute lCliente.VB_Description = "Codigo identificador do Cliente"
    mvarlCliente = vData
End Property

Public Property Get lCliente() As Long
    lCliente = mvarlCliente
End Property




VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassLiberacaoCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarsCodUsuario As String
Private mvardLimiteOperacao As Double
Private mvardLimiteMensal As Double

Public Property Let dLimiteMensal(ByVal vData As Double)
Attribute dLimiteMensal.VB_Description = "Valor limite que o usuario pode autorizar mensalmente"
    mvardLimiteMensal = vData
End Property

Public Property Get dLimiteMensal() As Double
    dLimiteMensal = mvardLimiteMensal
End Property

Public Property Let dLimiteOperacao(ByVal vData As Double)
Attribute dLimiteOperacao.VB_Description = "Valor limite que o usuario pode autorizar por operação"
    mvardLimiteOperacao = vData
End Property

Public Property Get dLimiteOperacao() As Double
    dLimiteOperacao = mvardLimiteOperacao
End Property

Public Property Let sCodUsuario(ByVal vData As String)
Attribute sCodUsuario.VB_Description = "Codigo do Usuario"
    mvarsCodUsuario = vData
End Property

Public Property Get sCodUsuario() As String
    sCodUsuario = mvarsCodUsuario
End Property




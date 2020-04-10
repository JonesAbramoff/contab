VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassCarteiraBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariBanco As Integer
Private mvariCodCarteiraCobranca As Integer
Private mvarsNome As String
Private mvarsCodCarteiraNoBanco As String
Private mvariNumCarteiraNoBanco As Integer
Private mvariPermiteTaxaDesconto As Integer
Private mvariExigeFaixaNossoNumero As Integer

Public Property Let iExigeFaixaNossoNumero(ByVal vData As Integer)
Attribute iExigeFaixaNossoNumero.VB_Description = "0 ou 1 p/saber se o ""Nosso Numero"" sera escolhido pela Empresa dentro de uma faixa liberada pelo Banco. Tipicamente usada na cobranca sem registro"
    mvariExigeFaixaNossoNumero = vData
End Property

Public Property Get iExigeFaixaNossoNumero() As Integer
    iExigeFaixaNossoNumero = mvariExigeFaixaNossoNumero
End Property

Public Property Let iPermiteTaxaDesconto(ByVal vData As Integer)
Attribute iPermiteTaxaDesconto.VB_Description = "0 ou 1"
    mvariPermiteTaxaDesconto = vData
End Property

Public Property Get iPermiteTaxaDesconto() As Integer
    iPermiteTaxaDesconto = mvariPermiteTaxaDesconto
End Property

Public Property Let iNumCarteiraNoBanco(ByVal vData As Integer)
    mvariNumCarteiraNoBanco = vData
End Property

Public Property Get iNumCarteiraNoBanco() As Integer
    iNumCarteiraNoBanco = mvariNumCarteiraNoBanco
End Property

Public Property Let sCodCarteiraNoBanco(ByVal vData As String)
Attribute sCodCarteiraNoBanco.VB_Description = "Itau usa codigo+num, p/outros basta o num abaixo"
    mvarsCodCarteiraNoBanco = vData
End Property

Public Property Get sCodCarteiraNoBanco() As String
    sCodCarteiraNoBanco = mvarsCodCarteiraNoBanco
End Property

Public Property Let sNome(ByVal vData As String)
Attribute sNome.VB_Description = "nome dado pelo banco p/a carteira"
    mvarsNome = vData
End Property

Public Property Get sNome() As String
    sNome = mvarsNome
End Property

Public Property Let iCodCarteiraCobranca(ByVal vData As Integer)
Attribute iCodCarteiraCobranca.VB_Description = "chave p/CarteirasCobranca"
    mvariCodCarteiraCobranca = vData
End Property

Public Property Get iCodCarteiraCobranca() As Integer
    iCodCarteiraCobranca = mvariCodCarteiraCobranca
End Property

Public Property Let iBanco(ByVal vData As Integer)
    mvariBanco = vData
End Property

Public Property Get iBanco() As Integer
    iBanco = mvariBanco
End Property




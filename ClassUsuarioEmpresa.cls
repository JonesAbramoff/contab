VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassUsuarioEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarsCodUsuario As String
Private mvarsSenha As String
Private mvarlCodEmpresa As Long
Private mvarsNomeEmpresa As String
Private mvariCodFilial As Integer
Private mvarsNomeFilial As String
Private mvariTelaOK As Integer

Public Property Let iTelaOK(ByVal vData As Integer)
Attribute iTelaOK.VB_Description = "Indica se a tela foi processada com sucesso. True = sucesso  False = insucesso"
    mvariTelaOK = vData
End Property

Public Property Get iTelaOK() As Integer
    iTelaOK = mvariTelaOK
End Property

Public Property Let sNomeFilial(ByVal vData As String)
Attribute sNomeFilial.VB_Description = "Nome da Filial selecionada"
    mvarsNomeFilial = vData
End Property

Public Property Get sNomeFilial() As String
    sNomeFilial = mvarsNomeFilial
End Property

Public Property Let iCodFilial(ByVal vData As Integer)
Attribute iCodFilial.VB_Description = "Codigo da Filial"
    mvariCodFilial = vData
End Property

Public Property Get iCodFilial() As Integer
    iCodFilial = mvariCodFilial
End Property

Public Property Let sNomeEmpresa(ByVal vData As String)
Attribute sNomeEmpresa.VB_Description = "Nome da Empresa Selecionada"
    mvarsNomeEmpresa = vData
End Property

Public Property Get sNomeEmpresa() As String
    sNomeEmpresa = mvarsNomeEmpresa
End Property

Public Property Let lCodEmpresa(ByVal vData As Long)
Attribute lCodEmpresa.VB_Description = "Codigo da Empresa selecionada"
    mvarlCodEmpresa = vData
End Property

Public Property Get lCodEmpresa() As Long
    lCodEmpresa = mvarlCodEmpresa
End Property

Public Property Let sSenha(ByVal vData As String)
Attribute sSenha.VB_Description = "Senha do Usuario"
    mvarsSenha = vData
End Property

Public Property Get sSenha() As String
    sSenha = mvarsSenha
End Property

Public Property Let sCodUsuario(ByVal vData As String)
Attribute sCodUsuario.VB_Description = "Codigo do Usuario"
    mvarsCodUsuario = vData
End Property

Public Property Get sCodUsuario() As String
    sCodUsuario = mvarsCodUsuario
End Property


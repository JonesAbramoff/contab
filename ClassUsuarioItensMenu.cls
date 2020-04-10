VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassUsuarioItensMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarsCodUsuario As String
Private mvarlCodEmpresa As Long
Private mvariCodFilial As Integer
Private mvardtDataValidade As Date
Private mvariTipoDeAcesso As Integer
Private mvariFilialEmpresa As Integer
Private mvarsNomeControle As String
Private mvariIndiceControle As Integer
Private mvarsNomeControlePai As String
Private mvariIndiceControlePai As Integer
Private mvarsNomeTela As String

Public Property Let sNomeTela(ByVal vData As String)
    mvarsNomeTela = vData
End Property

Public Property Get sNomeTela() As String
    sNomeTela = mvarsNomeTela
End Property

Public Property Let iIndiceControlePai(ByVal vData As Integer)
Attribute iIndiceControlePai.VB_Description = "Indice do controle pai do item do menu"
    mvariIndiceControlePai = vData
End Property

Public Property Get iIndiceControlePai() As Integer
    iIndiceControlePai = mvariIndiceControlePai
End Property

Public Property Let sNomeControlePai(ByVal vData As String)
Attribute sNomeControlePai.VB_Description = "Nome do controle pai do item do menu"
    mvarsNomeControlePai = vData
End Property

Public Property Get sNomeControlePai() As String
    sNomeControlePai = mvarsNomeControlePai
End Property

Public Property Let iIndiceControle(ByVal vData As Integer)
Attribute iIndiceControle.VB_Description = "Indice do Item do Menu"
    mvariIndiceControle = vData
End Property

Public Property Get iIndiceControle() As Integer
    iIndiceControle = mvariIndiceControle
End Property

Public Property Let sNomeControle(ByVal vData As String)
Attribute sNomeControle.VB_Description = "Nome do Item do Menu"
    mvarsNomeControle = vData
End Property

Public Property Get sNomeControle() As String
    sNomeControle = mvarsNomeControle
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
Attribute iFilialEmpresa.VB_Description = "Indica se a empresa e a filial tem acesso. 0=acesso somente à  empresa toda 1= acesso somente à filila 2= acesso à ambos"
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let iTipoDeAcesso(ByVal vData As Integer)
Attribute iTipoDeAcesso.VB_Description = "Tipo de Acesso. 0=Nao tem Acesso 1=Tem Acesso"
    mvariTipoDeAcesso = vData
End Property

Public Property Get iTipoDeAcesso() As Integer
    iTipoDeAcesso = mvariTipoDeAcesso
End Property

Public Property Let dtDataValidade(ByVal vData As Date)
Attribute dtDataValidade.VB_Description = "Data de Validade"
    mvardtDataValidade = vData
End Property

Public Property Get dtDataValidade() As Date
    dtDataValidade = mvardtDataValidade
End Property

Public Property Let iCodFilial(ByVal vData As Integer)
Attribute iCodFilial.VB_Description = "Codigo da Filial"
    mvariCodFilial = vData
End Property

Public Property Get iCodFilial() As Integer
    iCodFilial = mvariCodFilial
End Property

Public Property Let lCodEmpresa(ByVal vData As Long)
Attribute lCodEmpresa.VB_Description = "Codigo da Empresa"
    mvarlCodEmpresa = vData
End Property

Public Property Get lCodEmpresa() As Long
    lCodEmpresa = mvarlCodEmpresa
End Property

Public Property Let sCodUsuario(ByVal vData As String)
Attribute sCodUsuario.VB_Description = "Codigo do Usuario"
    mvarsCodUsuario = vData
End Property

Public Property Get sCodUsuario() As String
    sCodUsuario = mvarsCodUsuario
End Property




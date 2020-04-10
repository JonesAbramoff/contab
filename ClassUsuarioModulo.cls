VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassUsuarioModulo"
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
Private mvarsSiglaModulo As String
Private mvarsNomeModulo As String
Private mvarsVersaoModulo As String
Private mvariMenuModulo As Integer

Public Property Let iMenuModulo(ByVal vData As Integer)
Attribute iMenuModulo.VB_Description = "Número do Menu no Sistema. Serve para que se efetue a apresentação do menu correspondente."
    mvariMenuModulo = vData
End Property

Public Property Get iMenuModulo() As Integer
    iMenuModulo = mvariMenuModulo
End Property

Public Property Let sVersaoModulo(ByVal vData As String)
Attribute sVersaoModulo.VB_Description = "Versao do Modulo"
    mvarsVersaoModulo = vData
End Property

Public Property Get sVersaoModulo() As String
    sVersaoModulo = mvarsVersaoModulo
End Property

Public Property Let sNomeModulo(ByVal vData As String)
Attribute sNomeModulo.VB_Description = "Nome do Modulo"
    mvarsNomeModulo = vData
End Property

Public Property Get sNomeModulo() As String
    sNomeModulo = mvarsNomeModulo
End Property

Public Property Let sSiglaModulo(ByVal vData As String)
Attribute sSiglaModulo.VB_Description = "Sigla do Modulo"
    mvarsSiglaModulo = vData
End Property

Public Property Get sSiglaModulo() As String
    sSiglaModulo = mvarsSiglaModulo
End Property

Public Property Let dtDataValidade(ByVal vData As Date)
Attribute dtDataValidade.VB_Description = "Data de Validade do Modulo para o Usuario em questao"
    mvardtDataValidade = vData
End Property

Public Property Get dtDataValidade() As Date
    dtDataValidade = mvardtDataValidade
End Property

Public Property Let iCodFilial(ByVal vData As Integer)
Attribute iCodFilial.VB_Description = "Codigo da FilialEmpresa"
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


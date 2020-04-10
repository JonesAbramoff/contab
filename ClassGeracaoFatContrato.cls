VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassGeracaoFatContrato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlClienteIni As Long
Private mvarlClienteFim As Long
Private mvarsContratoIni As String
Private mvarsContratoFim As String
Private mvardtDataEmissao As Date
Private mvardtDataGeracao As Date
Private mvardtDataRefVencimento As Date
Private mvariFilialEmpresa As Integer
Private mvariTipoNFiscal As Integer
Private mvarcolItens As New Collection
Private mvardtDataCobrIni As Date
Private mvardtDataCobrFim As Date
Private mvarlNumIntDoc As Long
Private mvarsUsuario As String

Public Property Let sUsuario(ByVal vData As String)
    mvarsUsuario = vData
End Property

Public Property Get sUsuario() As String
    sUsuario = mvarsUsuario
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let dtDataCobrFim(ByVal vData As Date)
    mvardtDataCobrFim = vData
End Property

Public Property Get dtDataCobrFim() As Date
    dtDataCobrFim = mvardtDataCobrFim
End Property

Public Property Let dtDataCobrIni(ByVal vData As Date)
    mvardtDataCobrIni = vData
End Property

Public Property Get dtDataCobrIni() As Date
    dtDataCobrIni = mvardtDataCobrIni
End Property

Public Property Set colItens(ByVal vData As Collection)
    Set mvarcolItens = vData
End Property

Public Property Get colItens() As Collection
    Set colItens = mvarcolItens
End Property

Public Property Let iTipoNFiscal(ByVal vData As Integer)
    mvariTipoNFiscal = vData
End Property

Public Property Get iTipoNFiscal() As Integer
    iTipoNFiscal = mvariTipoNFiscal
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let dtDataRefVencimento(ByVal vData As Date)
    mvardtDataRefVencimento = vData
End Property

Public Property Get dtDataRefVencimento() As Date
    dtDataRefVencimento = mvardtDataRefVencimento
End Property

Public Property Let dtDataEmissao(ByVal vData As Date)
    mvardtDataEmissao = vData
End Property

Public Property Get dtDataEmissao() As Date
    dtDataEmissao = mvardtDataEmissao
End Property

Public Property Let dtDataGeracao(ByVal vData As Date)
    mvardtDataGeracao = vData
End Property

Public Property Get dtDataGeracao() As Date
    dtDataGeracao = mvardtDataGeracao
End Property

Public Property Let sContratoFim(ByVal vData As String)
    mvarsContratoFim = vData
End Property

Public Property Get sContratoFim() As String
    sContratoFim = mvarsContratoFim
End Property

Public Property Let sContratoIni(ByVal vData As String)
    mvarsContratoIni = vData
End Property

Public Property Get sContratoIni() As String
    sContratoIni = mvarsContratoIni
End Property

Public Property Let lClienteFim(ByVal vData As Long)
    mvarlClienteFim = vData
End Property

Public Property Get lClienteFim() As Long
    lClienteFim = mvarlClienteFim
End Property

Public Property Let lClienteIni(ByVal vData As Long)
    mvarlClienteIni = vData
End Property

Public Property Get lClienteIni() As Long
    lClienteIni = mvarlClienteIni
End Property




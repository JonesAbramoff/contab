VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassContrato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariFilialEmpresa As Integer
Private mvarlNumIntDoc As Long
Private mvarsCodigo As String
Private mvarsCcl As String
Private mvarsDescricao As String
Private mvariAtivo As Integer
Private mvarlCliente As Long
Private mvariFilCli As Integer
Private mvarlFornecedor As Long
Private mvariFilForn As Integer
Private mvarsObservacao As String
Private mvardtDataIniContrato As Date
Private mvardtDataFimContrato As Date
Private mvardtDataRenovContrato As Date
Private mvarsContaContabil As String
Private mvarsHistoricoContabil As String
Private mvardtDataIniCobrancaPadrao As Date
Private mvariPeriodicidadePadrao As Integer
Private mvariCondPagtoPadrao As Integer
Private mvarcolItens As New Collection
Private mvariTipoTributacao As Integer
Private mvarsNaturezaOp As String
Private mvariTipo As Integer
Private mvariRecibo As Integer
Private mvariNFe As Integer
Private mvarsSerie As String

Public Property Let sSerie(ByVal vData As String)
    mvarsSerie = vData
End Property

Public Property Get sSerie() As String
    sSerie = mvarsSerie
End Property

Public Property Let iNFe(ByVal vData As Integer)
    mvariNFe = vData
End Property

Public Property Get iNFe() As Integer
    iNFe = mvariNFe
End Property

Public Property Let iRecibo(ByVal vData As Integer)
    mvariRecibo = vData
End Property

Public Property Get iRecibo() As Integer
    iRecibo = mvariRecibo
End Property

Public Property Let iFilForn(ByVal vData As Integer)
    mvariFilForn = vData
End Property

Public Property Get iFilForn() As Integer
     iFilForn = mvariFilForn
End Property

Public Property Let lFornecedor(ByVal vData As Long)
    mvarlFornecedor = vData
End Property

Public Property Get lFornecedor() As Long
     lFornecedor = mvarlFornecedor
End Property

Public Property Let iTipo(ByVal vData As Integer)
    mvariTipo = vData
End Property

Public Property Get iTipo() As Integer
    iTipo = mvariTipo
End Property

Public Property Let sNaturezaOp(ByVal vData As String)
    mvarsNaturezaOp = vData
End Property

Public Property Get sNaturezaOp() As String
    sNaturezaOp = mvarsNaturezaOp
End Property

Public Property Let iTipoTributacao(ByVal vData As Integer)
    mvariTipoTributacao = vData
End Property

Public Property Get iTipoTributacao() As Integer
    iTipoTributacao = mvariTipoTributacao
End Property

Public Property Set colItens(ByVal vData As Collection)
    Set mvarcolItens = vData
End Property

Public Property Get colItens() As Collection
    Set colItens = mvarcolItens
End Property

Public Property Let iCondPagtoPadrao(ByVal vData As Integer)
    mvariCondPagtoPadrao = vData
End Property

Public Property Get iCondPagtoPadrao() As Integer
     iCondPagtoPadrao = mvariCondPagtoPadrao
End Property

Public Property Let iPeriodicidadePadrao(ByVal vData As Integer)
    mvariPeriodicidadePadrao = vData
End Property

Public Property Get iPeriodicidadePadrao() As Integer
     iPeriodicidadePadrao = mvariPeriodicidadePadrao
End Property

Public Property Let dtDataIniCobrancaPadrao(ByVal vData As Date)
    mvardtDataIniCobrancaPadrao = vData
End Property

Public Property Get dtDataIniCobrancaPadrao() As Date
     dtDataIniCobrancaPadrao = mvardtDataIniCobrancaPadrao
End Property

Public Property Let sHistoricoContabil(ByVal vData As String)
    mvarsHistoricoContabil = vData
End Property

Public Property Get sHistoricoContabil() As String
     sHistoricoContabil = mvarsHistoricoContabil
End Property

Public Property Let sContaContabil(ByVal vData As String)
    mvarsContaContabil = vData
End Property

Public Property Get sContaContabil() As String
     sContaContabil = mvarsContaContabil
End Property

Public Property Let dtDataRenovContrato(ByVal vData As Date)
    mvardtDataRenovContrato = vData
End Property

Public Property Get dtDataRenovContrato() As Date
     dtDataRenovContrato = mvardtDataRenovContrato
End Property

Public Property Let dtDataFimContrato(ByVal vData As Date)
    mvardtDataFimContrato = vData
End Property

Public Property Get dtDataFimContrato() As Date
     dtDataFimContrato = mvardtDataFimContrato
End Property

Public Property Let dtDataIniContrato(ByVal vData As Date)
    mvardtDataIniContrato = vData
End Property

Public Property Get dtDataIniContrato() As Date
     dtDataIniContrato = mvardtDataIniContrato
End Property

Public Property Let sObservacao(ByVal vData As String)
    mvarsObservacao = vData
End Property

Public Property Get sObservacao() As String
     sObservacao = mvarsObservacao
End Property

Public Property Let iFilCli(ByVal vData As Integer)
    mvariFilCli = vData
End Property

Public Property Get iFilCli() As Integer
     iFilCli = mvariFilCli
End Property

Public Property Let lCliente(ByVal vData As Long)
    mvarlCliente = vData
End Property

Public Property Get lCliente() As Long
     lCliente = mvarlCliente
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
     iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
     lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let iAtivo(ByVal vData As Integer)
    mvariAtivo = vData
End Property

Public Property Get iAtivo() As Integer
     iAtivo = mvariAtivo
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
     sDescricao = mvarsDescricao
End Property

Public Property Let sCodigo(ByVal vData As String)
    mvarsCodigo = vData
End Property

Public Property Get sCodigo() As String
     sCodigo = mvarsCodigo
End Property

Public Property Let sCcl(ByVal vData As String)
    mvarsCcl = vData
End Property

Public Property Get sCcl() As String
     sCcl = mvarsCcl
End Property

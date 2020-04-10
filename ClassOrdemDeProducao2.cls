VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassOrdemDeProducao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

Public objInfoUsu As Object

Private mvariFilialEmpresa As Integer
Private mvarsCodigo As String
Private mvardtDataEmissao As Date
Private mvarcolItens As New Collection
Private mvariNumItens As Integer
Private mvariNumItensBaixados As Integer
Private mvariGeraReqCompra As Integer
Private mvariGeraOP As Integer
Private mvarsCodOPGerar As String
Private mvarsOPGeradora As String
Private mvariStatusOP As Integer
Private mvarlCodPrestador As Long
Private mvariTipo As Integer '0 = Ordem de Producao, 1 = Ordem de Corte
Private mvariGeraOPsArvore As Integer
'Inserido por Jorge Specian
'----------------------------------
Private mvarlCodigoNumerico As Long
Public mvarobjTela As Object
Private mvarobjProjetoInfo As Object
Private mvariTipoTerc As Integer
Private mvariEscaninhoTerc As Integer
Private mvariFilialTerc As Integer
Private mvarlCodTerc As Long
Private mvariIgnoraEst As Integer

Public Property Let iIgnoraEst(ByVal vData As Integer)
    mvariIgnoraEst = vData
End Property

Public Property Get iIgnoraEst() As Integer
    iIgnoraEst = mvariIgnoraEst
End Property

Public Property Let lCodTerc(ByVal vData As Long)
    mvarlCodTerc = vData
End Property

Public Property Get lCodTerc() As Long
    lCodTerc = mvarlCodTerc
End Property

Public Property Let iFilialTerc(ByVal vData As Integer)
    mvariFilialTerc = vData
End Property

Public Property Get iFilialTerc() As Integer
    iFilialTerc = mvariFilialTerc
End Property

Public Property Let iEscaninhoTerc(ByVal vData As Integer)
    mvariEscaninhoTerc = vData
End Property

Public Property Get iEscaninhoTerc() As Integer
    iEscaninhoTerc = mvariEscaninhoTerc
End Property

Public Property Let iTipoTerc(ByVal vData As Integer)
    mvariTipoTerc = vData
End Property

Public Property Get iTipoTerc() As Integer
    iTipoTerc = mvariTipoTerc
End Property

Public Property Set objProjetoInfo(ByVal vData As Object)
    Set mvarobjProjetoInfo = vData
End Property

Public Property Get objProjetoInfo() As Object
    Set objProjetoInfo = mvarobjProjetoInfo
End Property

Public Property Set objTela(ByVal vData As Object)
    Set mvarobjTela = vData
End Property

Public Property Get objTela() As Object
    Set objTela = mvarobjTela
End Property

Public Property Let lCodigoNumerico(ByVal vData As Long)
    mvarlCodigoNumerico = vData
End Property

Public Property Get lCodigoNumerico() As Long
    lCodigoNumerico = mvarlCodigoNumerico
End Property

'----------------------------------

Public Property Let iGeraOPsArvore(ByVal vData As Integer)
    mvariGeraOPsArvore = vData
End Property

Public Property Get iGeraOPsArvore() As Integer
    iGeraOPsArvore = mvariGeraOPsArvore
End Property

Public Property Let iTipo(ByVal vData As Integer)
    mvariTipo = vData
End Property

Public Property Get iTipo() As Integer
    iTipo = mvariTipo
End Property

Public Property Let lCodPrestador(ByVal vData As Long)
    mvarlCodPrestador = vData
End Property

Public Property Get lCodPrestador() As Long
    lCodPrestador = mvarlCodPrestador
End Property

Public Property Let iStatusOP(ByVal vData As Integer)
    mvariStatusOP = vData
End Property

Public Property Get iStatusOP() As Integer
    iStatusOP = mvariStatusOP
End Property

Public Property Let sOPGeradora(ByVal vData As String)
Attribute sOPGeradora.VB_Description = "Código da OP que gerou automaticamente essa ordem de produção."
    mvarsOPGeradora = vData
End Property

Public Property Get sOPGeradora() As String
    sOPGeradora = mvarsOPGeradora
End Property

Public Property Let sCodOPGerar(ByVal vData As String)
Attribute sCodOPGerar.VB_Description = "O código da OP filha que deve ser garada a partir da gravação dessa OP."
    mvarsCodOPGerar = vData
End Property

Public Property Get sCodOPGerar() As String
    sCodOPGerar = mvarsCodOPGerar
End Property

Public Property Let iGeraOP(ByVal vData As Integer)
    mvariGeraOP = vData
End Property

Public Property Get iGeraOP() As Integer
    iGeraOP = mvariGeraOP
End Property

Public Property Let iGeraReqCompra(ByVal vData As Integer)
    mvariGeraReqCompra = vData
End Property

Public Property Get iGeraReqCompra() As Integer
    iGeraReqCompra = mvariGeraReqCompra
End Property

Public Property Let iNumItens(ByVal vData As Integer)
    mvariNumItens = vData
End Property

Public Property Get iNumItens() As Integer
    iNumItens = mvariNumItens
End Property

Public Property Let iNumItensBaixados(ByVal vData As Integer)
    mvariNumItensBaixados = vData
End Property

Public Property Get iNumItensBaixados() As Integer
    iNumItensBaixados = mvariNumItensBaixados
End Property

Public Property Get colItens() As Collection
    Set colItens = mvarcolItens
End Property

Public Property Let dtDataEmissao(ByVal vData As Date)
    mvardtDataEmissao = vData
End Property

Public Property Get dtDataEmissao() As Date
    dtDataEmissao = mvardtDataEmissao
End Property

Public Property Let sCodigo(ByVal vData As String)
    mvarsCodigo = vData
End Property

Public Property Get sCodigo() As String
    sCodigo = mvarsCodigo
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property




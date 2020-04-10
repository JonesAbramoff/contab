VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFluxoAnalitico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

Private mvarlFluxoId As Long
Private mvariTipoReg As Integer
Private mvarlFornecedor As Long
Private mvariFilial As Integer
Private mvarsSiglaDocumento As String
Private mvariNumParcela As Integer
Private mvardtData As Date
Private mvardValor As Double
Private mvariTipoFornecedor As Integer
Private mvarsNomeReduzido As String
Private mvarsDescricao As String
Private mvarlNumIntDoc As Long
Private mvarsTitulo As String
Private mvariItem As Integer
Private mvardtDataReferencia As Date
Private mvariFilialEmpresa As Integer

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let dtDataReferencia(ByVal vData As Date)
    mvardtDataReferencia = vData
End Property

Public Property Get dtDataReferencia() As Date
    dtDataReferencia = mvardtDataReferencia
End Property

Public Property Let iItem(ByVal vData As Integer)
    mvariItem = vData
End Property

Public Property Get iItem() As Integer
    iItem = mvariItem
End Property

Public Property Let sTitulo(ByVal vData As String)
    mvarsTitulo = vData
End Property

Public Property Get sTitulo() As String
    sTitulo = mvarsTitulo
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let sDescricao(ByVal vData As String)
Attribute sDescricao.VB_Description = "Descricao do Tipo do Fornecedor/Cliente"
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let sNomeReduzido(ByVal vData As String)
Attribute sNomeReduzido.VB_Description = "Nome Reduzido do Cliente ou Fornecedor"
    mvarsNomeReduzido = vData
End Property

Public Property Get sNomeReduzido() As String
    sNomeReduzido = mvarsNomeReduzido
End Property

Public Property Let iTipoFornecedor(ByVal vData As Integer)
Attribute iTipoFornecedor.VB_Description = "Tipo do Fornecedor/Cliente"
    mvariTipoFornecedor = vData
End Property

Public Property Get iTipoFornecedor() As Integer
    iTipoFornecedor = mvariTipoFornecedor
End Property

Public Property Let dValor(ByVal vData As Double)
Attribute dValor.VB_Description = "Valor do Titulo"
    mvardValor = vData
End Property

Public Property Get dValor() As Double
    dValor = mvardValor
End Property

Public Property Let dtData(ByVal vData As Date)
Attribute dtData.VB_Description = "data ou data prev p/deposito de cheque-pre"
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let iNumParcela(ByVal vData As Integer)
Attribute iNumParcela.VB_Description = "Numero da Parcela ou cta p/cheque pre"
    mvariNumParcela = vData
End Property

Public Property Get iNumParcela() As Integer
    iNumParcela = mvariNumParcela
End Property

Public Property Let sSiglaDocumento(ByVal vData As String)
Attribute sSiglaDocumento.VB_Description = "Tipo do Titulo"
    mvarsSiglaDocumento = vData
End Property

Public Property Get sSiglaDocumento() As String
    sSiglaDocumento = mvarsSiglaDocumento
End Property

Public Property Let iFilial(ByVal vData As Integer)
Attribute iFilial.VB_Description = "Filial da Empresa"
    mvariFilial = vData
End Property

Public Property Get iFilial() As Integer
    iFilial = mvariFilial
End Property

Public Property Let lFornecedor(ByVal vData As Long)
Attribute lFornecedor.VB_Description = "fornecedor ou cliente ou cobrador ou zero p/aplicacao ou # do sequencial p/cheque-pre"
    mvarlFornecedor = vData
End Property

Public Property Get lFornecedor() As Long
    lFornecedor = mvarlFornecedor
End Property

Public Property Let iTipoReg(ByVal vData As Integer)
Attribute iTipoReg.VB_Description = "0:pagto/1:recebto/2:bordero de cobranca descontado/3:aplicacao/4:cheques-pre"
    mvariTipoReg = vData
End Property

Public Property Get iTipoReg() As Integer
    iTipoReg = mvariTipoReg
End Property

Public Property Let lFluxoId(ByVal vData As Long)
Attribute lFluxoId.VB_Description = "identificador do fluxo de caixa"
    mvarlFluxoId = vData
End Property

Public Property Get lFluxoId() As Long
    lFluxoId = mvarlFluxoId
End Property


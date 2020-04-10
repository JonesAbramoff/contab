VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassBorderoCobrancaEmissao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

Private mvariCobrador As Integer
Private mvariCarteira As Integer
Private mvardtDataEmissao As Date
Private mvardtDataVencimento As Date
Private mvardValorMaximo As Double
Private mvarsSiglaEstado As String
Private mvariTipoCliente As Integer
Private mvardValorParcelasSelecionadas As Double
Private mvariQtdeParcelasSelecionadas As Integer
Private mvarcolInfoParcRec As New Collection
Private mvarobjTelaAtualizacao As Object 'tela de acompanhamento da atualizacao de arquivos. Necessaria para a contabilizacao.
Private mvardtDataContabil As Date
Private mvarlNumero As Long 'local copy
Private mvarlNumIntBordero As Long 'local copy
Private mvariTipoAceite As Integer

Public Property Let iTipoAceite(ByVal vData As Integer)
    mvariTipoAceite = vData
End Property

Public Property Get iTipoAceite() As Integer
    iTipoAceite = mvariTipoAceite
End Property

Public Property Let lNumIntBordero(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lNumIntBordero = 5
    mvarlNumIntBordero = vData
End Property


Public Property Get lNumIntBordero() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lNumIntBordero
    lNumIntBordero = mvarlNumIntBordero
End Property



Public Property Let lNumero(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lNumero = 5
    mvarlNumero = vData
End Property


Public Property Get lNumero() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lNumero
    lNumero = mvarlNumero
End Property




Public Property Let dtDataContabil(ByVal vData As Date)
    mvardtDataContabil = vData
End Property

Public Property Get dtDataContabil() As Date
    dtDataContabil = mvardtDataContabil
End Property

Public Property Let objTelaAtualizacao(ByVal vData As Object)
    Set mvarobjTelaAtualizacao = vData
End Property

Public Property Get objTelaAtualizacao() As Object
    Set objTelaAtualizacao = mvarobjTelaAtualizacao
End Property

Public Property Get colInfoParcRec() As Collection
    Set colInfoParcRec = mvarcolInfoParcRec
End Property

Public Property Let iQtdeParcelasSelecionadas(ByVal vData As Integer)
    mvariQtdeParcelasSelecionadas = vData
End Property

Public Property Get iQtdeParcelasSelecionadas() As Integer
    iQtdeParcelasSelecionadas = mvariQtdeParcelasSelecionadas
End Property

Public Property Let dValorParcelasSelecionadas(ByVal vData As Double)
    mvardValorParcelasSelecionadas = vData
End Property

Public Property Get dValorParcelasSelecionadas() As Double
    dValorParcelasSelecionadas = mvardValorParcelasSelecionadas
End Property

Public Property Let iTipoCliente(ByVal vData As Integer)
    mvariTipoCliente = vData
End Property

Public Property Get iTipoCliente() As Integer
    iTipoCliente = mvariTipoCliente
End Property

Public Property Let sSiglaEstado(ByVal vData As String)
    mvarsSiglaEstado = vData
End Property

Public Property Get sSiglaEstado() As String
    sSiglaEstado = mvarsSiglaEstado
End Property

Public Property Let dValorMaximo(ByVal vData As Double)
    mvardValorMaximo = vData
End Property

Public Property Get dValorMaximo() As Double
    dValorMaximo = mvardValorMaximo
End Property

Public Property Let dtDataVencimento(ByVal vData As Date)
    mvardtDataVencimento = vData
End Property

Public Property Get dtDataVencimento() As Date
    dtDataVencimento = mvardtDataVencimento
End Property

Public Property Let dtDataEmissao(ByVal vData As Date)
    mvardtDataEmissao = vData
End Property

Public Property Get dtDataEmissao() As Date
    dtDataEmissao = mvardtDataEmissao
End Property

Public Property Let iCarteira(ByVal vData As Integer)
    mvariCarteira = vData
End Property

Public Property Get iCarteira() As Integer
    iCarteira = mvariCarteira
End Property

Public Property Let iCobrador(ByVal vData As Integer)
    mvariCobrador = vData
End Property

Public Property Get iCobrador() As Integer
    iCobrador = mvariCobrador
End Property

Function GeraContabilizacao(objContabAutomatica As ClassContabAutomatica, vParams As Variant) As Long

    GeraContabilizacao = mvarobjTelaAtualizacao.GeraContabilizacao(objContabAutomatica, vParams)
    
End Function



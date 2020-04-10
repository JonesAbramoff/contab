VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassBorderoDescChq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvariCobrador As Integer
Private mvariCarteiraCobranca As Integer
Private mvariContaCorrente As Integer
Private mvardtDataEmissao As Date
Private mvardtDataContabil As Date
Private mvardtDataDeposito As Date
Private mvardtDataBomParaAte As Date
Private mvariBanco As Integer
Private mvarsAgencia As String
Private mvarlNumBordero As Long
Private mvarcolChequePre As New Collection
Private mvariQuantChequesSel As Integer
Private mvarobjTelaAtualizacao As Object 'tela de acompanhamento da atualizacao de arquivos. Necessaria para a contabilizacao.
Private mvardValorChequesSel As Double
Private mvardValorCredito As Double
Private mvariFilialEmpresa As Integer
Private mvarsCarteiraCobranca As String
Private mvarsCobrador As String
Private mvarsContaCorrente As String
Private mvarcolInfoParcRec As New Collection

Public Property Get colInfoParcRec() As Collection
    Set colInfoParcRec = mvarcolInfoParcRec
End Property

Public Property Let sContaCorrente(ByVal vData As String)
    mvarsContaCorrente = vData
End Property

Public Property Get sContaCorrente() As String
    sContaCorrente = mvarsContaCorrente
End Property

Public Property Let sCobrador(ByVal vData As String)
    mvarsCobrador = vData
End Property

Public Property Get sCobrador() As String
    sCobrador = mvarsCobrador
End Property

Public Property Let sCarteiraCobranca(ByVal vData As String)
    mvarsCarteiraCobranca = vData
End Property

Public Property Get sCarteiraCobranca() As String
    sCarteiraCobranca = mvarsCarteiraCobranca
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let dValorCredito(ByVal vData As Double)
    mvardValorCredito = vData
End Property

Public Property Get dValorCredito() As Double
    dValorCredito = mvardValorCredito
End Property

Public Property Let dValorChequesSel(ByVal vData As Double)
    mvardValorChequesSel = vData
End Property

Public Property Get dValorChequesSel() As Double
    dValorChequesSel = mvardValorChequesSel
End Property

Public Property Let objTelaAtualizacao(ByVal vData As Object)
    Set mvarobjTelaAtualizacao = vData
End Property

Public Property Get objTelaAtualizacao() As Object
    Set objTelaAtualizacao = mvarobjTelaAtualizacao
End Property

Public Property Let iQuantChequesSel(ByVal vData As Integer)
    mvariQuantChequesSel = vData
End Property

Public Property Get iQuantChequesSel() As Integer
    iQuantChequesSel = mvariQuantChequesSel
End Property

Public Property Get colChequePre() As Collection
    Set colChequePre = mvarcolChequePre
End Property

Public Property Let lNumBordero(ByVal vData As Long)
    mvarlNumBordero = vData
End Property

Public Property Get lNumBordero() As Long
    lNumBordero = mvarlNumBordero
End Property

Public Property Let sAgencia(ByVal vData As String)
    mvarsAgencia = vData
End Property

Public Property Get sAgencia() As String
    sAgencia = mvarsAgencia
End Property

Public Property Let iBanco(ByVal vData As Integer)
    mvariBanco = vData
End Property

Public Property Get iBanco() As Integer
    iBanco = mvariBanco
End Property

Public Property Let dtDataBomParaAte(ByVal vData As Date)
    mvardtDataBomParaAte = vData
End Property

Public Property Get dtDataBomParaAte() As Date
    dtDataBomParaAte = mvardtDataBomParaAte
End Property

Public Property Let dtDataDeposito(ByVal vData As Date)
    mvardtDataDeposito = vData
End Property

Public Property Get dtDataDeposito() As Date
    dtDataDeposito = mvardtDataDeposito
End Property

Public Property Let dtDataContabil(ByVal vData As Date)
    mvardtDataContabil = vData
End Property

Public Property Get dtDataContabil() As Date
    dtDataContabil = mvardtDataContabil
End Property

Public Property Let dtDataEmissao(ByVal vData As Date)
    mvardtDataEmissao = vData
End Property

Public Property Get dtDataEmissao() As Date
    dtDataEmissao = mvardtDataEmissao
End Property

Public Property Let iContaCorrente(ByVal vData As Integer)
    mvariContaCorrente = vData
End Property

Public Property Get iContaCorrente() As Integer
    iContaCorrente = mvariContaCorrente
End Property

Public Property Let iCarteiraCobranca(ByVal vData As Integer)
    mvariCarteiraCobranca = vData
End Property

Public Property Get iCarteiraCobranca() As Integer
    iCarteiraCobranca = mvariCarteiraCobranca
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


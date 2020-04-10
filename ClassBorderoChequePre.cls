VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassBorderoChequePre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

Private mvarlNumBordero As Long
Private mvariCodNossaConta As Integer
Private mvardtDataEmissao As Date
Private mvardtDataDeposito As Date
Private mvardtDataContabil As Date
Private mvarobjTelaAtualizacao As Object 'tela de acompanhamento da atualizacao de arquivos. Necessaria para a contabilizacao.
Private mvarcolChequePre As New Collection
Private mvariQuantChequesSel As Integer 'local copy
Private mvardValorChequesSelecionados As Double 'local copy
Private mvarlNumIntBordero As Long 'local copy
Private mvarcolInfoParcRec As New Collection
Private mvariTipoBordero As Integer
Private mvariQuantParcelas As Integer
Private mvariFilialEmpresa As Integer

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let iTipoBordero(ByVal vData As Integer)
    mvariTipoBordero = vData
End Property

Public Property Get iTipoBordero() As Integer
    iTipoBordero = mvariTipoBordero
End Property

Public Property Get colInfoParcRec() As Collection
    Set colInfoParcRec = mvarcolInfoParcRec
End Property

Public Property Let iQuantParcelas(ByVal vData As Integer)
    mvariQuantParcelas = vData
End Property

Public Property Get iQuantParcelas() As Integer
    iQuantParcelas = mvariQuantParcelas
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



Public Property Let dValorChequesSelecionados(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dValorChequesSelecionados = 5
    mvardValorChequesSelecionados = vData
End Property


Public Property Get dValorChequesSelecionados() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dValorChequesSelecionados
    dValorChequesSelecionados = mvardValorChequesSelecionados
End Property



Public Property Let iQuantChequesSel(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iQuantChequesSel = 5
    mvariQuantChequesSel = vData
End Property


Public Property Get iQuantChequesSel() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iQuantChequesSel
    iQuantChequesSel = mvariQuantChequesSel
End Property




Public Property Get colChequePre() As Collection
    Set colChequePre = mvarcolChequePre
End Property

Public Property Let objTelaAtualizacao(ByVal vData As Object)
    Set mvarobjTelaAtualizacao = vData
End Property

Public Property Get objTelaAtualizacao() As Object
    Set objTelaAtualizacao = mvarobjTelaAtualizacao
End Property

Public Property Let dtDataContabil(ByVal vData As Date)
    mvardtDataContabil = vData
End Property

Public Property Get dtDataContabil() As Date
    dtDataContabil = mvardtDataContabil
End Property

Public Property Let dtDataDeposito(ByVal vData As Date)
    mvardtDataDeposito = vData
End Property

Public Property Get dtDataDeposito() As Date
    dtDataDeposito = mvardtDataDeposito
End Property

Public Property Let dtDataEmissao(ByVal vData As Date)
    mvardtDataEmissao = vData
End Property

Public Property Get dtDataEmissao() As Date
    dtDataEmissao = mvardtDataEmissao
End Property

Public Property Let iCodNossaConta(ByVal vData As Integer)
Attribute iCodNossaConta.VB_Description = "identifica a conta corrente onde foi efetuado o deposito"
    mvariCodNossaConta = vData
End Property

Public Property Get iCodNossaConta() As Integer
    iCodNossaConta = mvariCodNossaConta
End Property

Public Property Let lNumBordero(ByVal vData As Long)
Attribute lNumBordero.VB_Description = "numero do bordero, unico independente do cobrador"
    mvarlNumBordero = vData
End Property

Public Property Get lNumBordero() As Long
    lNumBordero = mvarlNumBordero
End Property

Function GeraContabilizacao(objContabAutomatica As ClassContabAutomatica, vParams As Variant) As Long

    GeraContabilizacao = mvarobjTelaAtualizacao.GeraContabilizacao(objContabAutomatica, vParams)
    
End Function

Sub LimpaColInfoParcRec()
    Set mvarcolInfoParcRec = New Collection
End Sub

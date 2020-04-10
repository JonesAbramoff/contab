VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassBorderoRetCobr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"colBaixaParcRec"
Option Explicit

Private mvarlNumIntBordero As Long
Private mvariBanco As Integer
Private mvariCobrador As Integer
Private mvarlNumBorderoRet As Long
Private mvarsNomeArq As String
Private mvardtDataRecepcao As Date
Private mvardtDataCredito As Date

Private mvarobjBaixaReceber As New ClassBaixaReceber
Private mvarcolBaixaParcRec As New colBaixaParcRec
Private mvarobjMovCCI As New ClassMovContaCorrente
Private mvarcolOcorrRetParcRec As New Collection

Private mvarcolBorderoRetCobrCart As New Collection

Private mvarcolRetCobrErros As New Collection

Private mvarcolCartCobr As New Collection
Private mvaralComando(1 To 20) As Long
Private mvariQuebraMovCCPorStatusRecurso As Integer
Private mvariVerificaCreditosTarifas As Integer

Public Property Let iVerificaCreditosTarifas(ByVal vData As Integer)
    mvariVerificaCreditosTarifas = vData
End Property

Public Property Get iVerificaCreditosTarifas() As Integer
    iVerificaCreditosTarifas = mvariVerificaCreditosTarifas
End Property

Public Property Let iQuebraMovCCPorStatusRecurso(ByVal vData As Integer)
    mvariQuebraMovCCPorStatusRecurso = vData
End Property

Public Property Get iQuebraMovCCPorStatusRecurso() As Integer
    iQuebraMovCCPorStatusRecurso = mvariQuebraMovCCPorStatusRecurso
End Property

Public Property Let alComando(ByVal iIndice As Integer, ByVal vData As Long)
    mvaralComando(iIndice) = vData
End Property

Public Property Get alComando(ByVal iIndice As Integer) As Long
    alComando = mvaralComando(iIndice)
End Property

Public Property Let iBanco(ByVal vData As Integer)
    mvariBanco = vData
End Property

Public Property Get iBanco() As Integer
    iBanco = mvariBanco
End Property

Public Property Set colRetCobrErros(ByVal vData As Collection)
    Set mvarcolRetCobrErros = vData
End Property

Public Property Get colRetCobrErros() As Collection
    Set colRetCobrErros = mvarcolRetCobrErros
End Property

Public Property Set colCartCobr(ByVal vData As Collection)
    Set mvarcolCartCobr = vData
End Property

Public Property Get colCartCobr() As Collection
    Set colCartCobr = mvarcolCartCobr
End Property

'Private mvarobjTelaAtualizacao As Object 'tela de acompanhamento da atualizacao de arquivos. Necessaria para a contabilizacao.
'
'Public Property Let objTelaAtualizacao(ByVal vData As Object)
'    Set mvarobjTelaAtualizacao = vData
'End Property
'
'Public Property Get objTelaAtualizacao() As Object
'    Set objTelaAtualizacao = mvarobjTelaAtualizacao
'End Property

Public Property Let dtDataCredito(ByVal vData As Date)
    mvardtDataCredito = vData
End Property

Public Property Get dtDataCredito() As Date
    dtDataCredito = mvardtDataCredito
End Property

Public Property Set colOcorrRetParcRec(ByVal vData As Collection)
    Set mvarcolOcorrRetParcRec = vData
End Property

Public Property Get colOcorrRetParcRec() As Collection
    Set colOcorrRetParcRec = mvarcolOcorrRetParcRec
End Property

Public Property Set colBorderoRetCobrCart(ByVal vData As Collection)
    Set mvarcolBorderoRetCobrCart = vData
End Property

Public Property Get colBorderoRetCobrCart() As Collection
    Set colBorderoRetCobrCart = mvarcolBorderoRetCobrCart
End Property

Public Property Set objMovCCI(ByVal vData As ClassMovContaCorrente)
    Set mvarobjMovCCI = vData
End Property

Public Property Get objMovCCI() As ClassMovContaCorrente
    Set objMovCCI = mvarobjMovCCI
End Property

Public Property Set colBaixaParcRec(ByVal vData As colBaixaParcRec)
    Set mvarcolBaixaParcRec = vData
End Property

Public Property Get colBaixaParcRec() As colBaixaParcRec
    Set colBaixaParcRec = mvarcolBaixaParcRec
End Property

Public Property Set objBaixaReceber(ByVal vData As ClassBaixaReceber)
    Set mvarobjBaixaReceber = vData
End Property

Public Property Get objBaixaReceber() As ClassBaixaReceber
    Set objBaixaReceber = mvarobjBaixaReceber
End Property

Public Property Let dtDataRecepcao(ByVal vData As Date)
    mvardtDataRecepcao = vData
End Property

Public Property Get dtDataRecepcao() As Date
    dtDataRecepcao = mvardtDataRecepcao
End Property

Public Property Let sNomeArq(ByVal vData As String)
    mvarsNomeArq = vData
End Property

Public Property Get sNomeArq() As String
    sNomeArq = mvarsNomeArq
End Property

Public Property Let lNumBorderoRet(ByVal vData As Long)
    mvarlNumBorderoRet = vData
End Property

Public Property Get lNumBorderoRet() As Long
    lNumBorderoRet = mvarlNumBorderoRet
End Property

Public Property Let iCobrador(ByVal vData As Integer)
    mvariCobrador = vData
End Property

Public Property Get iCobrador() As Integer
    iCobrador = mvariCobrador
End Property

Public Property Let lNumIntBordero(ByVal vData As Long)
    mvarlNumIntBordero = vData
End Property

Public Property Get lNumIntBordero() As Long
    lNumIntBordero = mvarlNumIntBordero
End Property

Public Sub CarteiraCobranca_Obter(iCodCarteiraNoBanco As Integer, objCarteiraCobrador As ClassCarteiraCobrador)

Dim objCarteiraCobradorAux As ClassCarteiraCobrador

    Set objCarteiraCobrador = Nothing
    
    For Each objCarteiraCobradorAux In colCartCobr
    
        If objCarteiraCobradorAux.iNumCarteiraNoBanco = iCodCarteiraNoBanco Then
            Set objCarteiraCobrador = objCarteiraCobradorAux
            Exit For
        End If
        
    Next
    
End Sub

Public Sub Abrir_Comando()
Dim iIndice As Integer
    For iIndice = LBound(mvaralComando) To UBound(mvaralComando)
        mvaralComando(iIndice) = Comando_Abrir()
    Next
End Sub

Public Sub Fechar_Comando()
Dim iIndice As Integer
    For iIndice = LBound(mvaralComando) To UBound(mvaralComando)
        Call Comando_Fechar(mvaralComando(iIndice))
    Next
End Sub

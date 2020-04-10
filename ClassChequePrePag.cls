VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassChequePrePag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDoc As Long
Private mvariFilialEmpresa As Integer
Private mvariContaCorrente As Integer
Private mvarlNumero As Long
Private mvardtDataEmissao As Date
Private mvardtDataBomPara As Date
Private mvardtDataDeposito As Date
Private mvardValor As Double
Private mvariStatus As Integer
Private mvarsObservacao As String
Private mvarsFavorecido As String
Private mvarcolParcelas As New Collection
Private mvarlFornecedor As Long
Private mvariFilial As Integer
Private mvarsNomeFornecedor As String
Private mvarsNomeContaCorrente As String
Private mvariIndiceChequeProc As Integer 'indice na colecao mvarcolInfoChequePag do cheque que está sendo processado na gravacao
Private mvarobjTelaAtualizacao As Object 'tela de acompanhamento da atualizacao de arquivos. Necessaria para a contabilizacao.
'Private mvardValorPago As Double
'Private mvardValorBaixado As Double
'Private mvardValorJuros As Double
'Private mvardValorMulta As Double
'Private mvardValorDesconto As Double
Private mvarobjFornecedor As New ClassFornecedor
Private mvarobjFilialFornecedor As New ClassFilialFornecedor

Private mvaradValorDiferenca() As Double

Private mvarobjColBaixaParcPagar As ColBaixaParcPagar

Public Property Get objColBaixaParcPagar() As ColBaixaParcPagar
    Set objColBaixaParcPagar = mvarobjColBaixaParcPagar
End Property

Public Property Set objColBaixaParcPagar(ByVal vData As ColBaixaParcPagar)
    Set mvarobjColBaixaParcPagar = vData
End Property

Public Property Get adValorDiferenca_UBound() As Integer
    adValorDiferenca_UBound = UBound(mvaradValorDiferenca)
End Property

Public Property Get adValorDiferenca_LBound() As Integer
    adValorDiferenca_LBound = LBound(mvaradValorDiferenca)
End Property

Public Property Let adValorDiferenca(ByVal vIndex As Variant, ByVal vData As Double)
    If vIndex > UBound(mvaradValorDiferenca) Then ReDim Preserve mvaradValorDiferenca(vIndex + 1)
    mvaradValorDiferenca(vIndex) = vData
End Property

Public Property Get adValorDiferenca(ByVal vIndex As Variant) As Double
    adValorDiferenca = mvaradValorDiferenca(vIndex)
End Property

Public Property Let dValorPago(ByVal vData As Double)
End Property
Public Property Let dValorBaixado(ByVal vData As Double)
End Property
Public Property Let dValorJuros(ByVal vData As Double)
End Property
Public Property Let dValorMulta(ByVal vData As Double)
End Property
Public Property Let dValorDesconto(ByVal vData As Double)
End Property

Public Property Set objFilialFornecedor(ByVal vData As ClassFilialFornecedor)
    Set mvarobjFilialFornecedor = vData
End Property

Public Property Get objFilialFornecedor() As ClassFilialFornecedor
    Set objFilialFornecedor = mvarobjFilialFornecedor
End Property

Public Property Set objFornecedor(ByVal vData As ClassFornecedor)
    Set mvarobjFornecedor = vData
End Property

Public Property Get objFornecedor() As ClassFornecedor
    Set objFornecedor = mvarobjFornecedor
End Property

Public Property Get dValorPago() As Double
Dim objParc As ClassChequePrePagParc
Dim dValorAux As Double
    For Each objParc In mvarcolParcelas
        dValorAux = dValorAux + objParc.dValorPago
    Next
    dValorPago = dValorAux
End Property

Public Property Get dValorBaixado() As Double
Dim objParc As ClassChequePrePagParc
Dim dValorAux As Double
    For Each objParc In mvarcolParcelas
        dValorAux = dValorAux + objParc.dValorBaixado
    Next
    dValorBaixado = dValorAux
End Property

Public Property Get dValorJuros() As Double
Dim objParc As ClassChequePrePagParc
Dim dValorAux As Double
    For Each objParc In mvarcolParcelas
        dValorAux = dValorAux + objParc.dValorJuros
    Next
    dValorJuros = dValorAux
End Property

Public Property Get dValorMulta() As Double
Dim objParc As ClassChequePrePagParc
Dim dValorAux As Double
    For Each objParc In mvarcolParcelas
        dValorAux = dValorAux + objParc.dValorMulta
    Next
    dValorMulta = dValorAux
End Property

Public Property Get dValorDesconto() As Double
Dim objParc As ClassChequePrePagParc
Dim dValorAux As Double
    For Each objParc In mvarcolParcelas
        dValorAux = dValorAux + objParc.dValorDesconto
    Next
    dValorDesconto = dValorAux
End Property

Public Property Let objTelaAtualizacao(ByVal vData As Object)
    Set mvarobjTelaAtualizacao = vData
End Property

Public Property Get objTelaAtualizacao() As Object
    Set objTelaAtualizacao = mvarobjTelaAtualizacao
End Property

Public Property Let iIndiceChequeProc(ByVal vData As Integer)
    mvariIndiceChequeProc = vData
End Property

Public Property Get iIndiceChequeProc() As Integer
    iIndiceChequeProc = mvariIndiceChequeProc
End Property

Public Property Let sNomeContaCorrente(ByVal vData As String)
    mvarsNomeContaCorrente = vData
End Property

Public Property Get sNomeContaCorrente() As String
    sNomeContaCorrente = mvarsNomeContaCorrente
End Property

Public Property Let sNomeFornecedor(ByVal vData As String)
    mvarsNomeFornecedor = vData
End Property

Public Property Get sNomeFornecedor() As String
    sNomeFornecedor = mvarsNomeFornecedor
End Property

Public Property Let iFilial(ByVal vData As Integer)
    mvariFilial = vData
End Property

Public Property Get iFilial() As Integer
    iFilial = mvariFilial
End Property

Public Property Let lFornecedor(ByVal vData As Long)
    mvarlFornecedor = vData
End Property

Public Property Get lFornecedor() As Long
    lFornecedor = mvarlFornecedor
End Property

Public Property Get colParcelas() As Collection
    Set colParcelas = mvarcolParcelas
End Property

Public Property Set colParcelas(ByVal vData As Collection)
    Set mvarcolParcelas = vData
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let iContaCorrente(ByVal vData As Integer)
    mvariContaCorrente = vData
End Property

Public Property Get iContaCorrente() As Integer
    iContaCorrente = mvariContaCorrente
End Property

Public Property Let lNumero(ByVal vData As Long)
    mvarlNumero = vData
End Property

Public Property Get lNumero() As Long
    lNumero = mvarlNumero
End Property

Public Property Let dtDataEmissao(ByVal vData As Date)
    mvardtDataEmissao = vData
End Property

Public Property Get dtDataEmissao() As Date
    dtDataEmissao = mvardtDataEmissao
End Property

Public Property Let dtDataBomPara(ByVal vData As Date)
    mvardtDataBomPara = vData
End Property

Public Property Get dtDataBomPara() As Date
    dtDataBomPara = mvardtDataBomPara
End Property

Public Property Let dtDataDeposito(ByVal vData As Date)
    mvardtDataDeposito = vData
End Property

Public Property Get dtDataDeposito() As Date
    dtDataDeposito = mvardtDataDeposito
End Property

Public Property Let dValor(ByVal vData As Double)
    mvardValor = vData
End Property

Public Property Get dValor() As Double
    dValor = mvardValor
End Property

Public Property Let iStatus(ByVal vData As Integer)
    mvariStatus = vData
End Property

Public Property Get iStatus() As Integer
    iStatus = mvariStatus
End Property

Public Property Let sObservacao(ByVal vData As String)
    mvarsObservacao = vData
End Property

Public Property Get sObservacao() As String
    sObservacao = mvarsObservacao
End Property

Public Property Let sFavorecido(ByVal vData As String)
    mvarsFavorecido = vData
End Property

Public Property Get sFavorecido() As String
    sFavorecido = mvarsFavorecido
End Property

Function GeraContabilizacao(objContabAutomatica As ClassContabAutomatica, vParams As Variant) As Long
    GeraContabilizacao = mvarobjTelaAtualizacao.GeraContabilizacao(objContabAutomatica, vParams)
End Function

Private Sub Class_Initialize()
    ReDim Preserve mvaradValorDiferenca(1)
End Sub


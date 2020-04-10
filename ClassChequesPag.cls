VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassChequesPag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"ColInfoChequePag"
Attribute VB_Ext_KEY = "Member1" ,"colInfoChequePag"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member2" ,"colInfoParcPag"
Option Explicit

Private mvariCta As Integer
Private mvarlNumCheque As Long
Private mvardtEmissao As Date
Private mvardtVenctoDe As Date
Private mvardtVencto As Date
Private mvariPortador As Integer
Private mvarcolInfoParcPag As New Collection
Private mvarcolInfoChequePag As New colInfoChequePag
Private mvardtContabil As Date
Private mvariQualquerPortador As Integer
Private mvarlNumImpressao As Long
Private mvarsLayoutCheque As String
Private mvariFilialEmpresaCta As Integer
Private mvarobjEvolucao As Object
Private mvariIndiceChequeProc As Integer 'indice na colecao mvarcolInfoChequePag do cheque que está sendo processado na gravacao
Private mvarobjTelaAtualizacao As Object 'tela de acompanhamento da atualizacao de arquivos. Necessaria para a contabilizacao.

Private mvardValorBaixado As Double
Private mvardValorMulta As Double
Private mvardValorJuros As Double
Private mvardValorDesconto As Double
Private mvaradValorDiferenca() As Double

Private mvarobjColBaixaParcPagar As ColBaixaParcPagar
Private mvarlNumIntDocChequePre As Long

Public Property Let lNumIntDocChequePre(ByVal vData As Long)
    mvarlNumIntDocChequePre = vData
End Property

Public Property Get lNumIntDocChequePre() As Long
    lNumIntDocChequePre = mvarlNumIntDocChequePre
End Property

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

Public Property Let dValorBaixado(ByVal vData As Double)
    mvardValorBaixado = vData
End Property

Public Property Get dValorBaixado() As Double
    dValorBaixado = mvardValorBaixado
End Property

Public Property Let dValorMulta(ByVal vData As Double)
    mvardValorMulta = vData
End Property

Public Property Get dValorMulta() As Double
    dValorMulta = mvardValorMulta
End Property

Public Property Let dValorJuros(ByVal vData As Double)
    mvardValorJuros = vData
End Property

Public Property Get dValorJuros() As Double
    dValorJuros = mvardValorJuros
End Property

Public Property Let dValorDesconto(ByVal vData As Double)
    mvardValorDesconto = vData
End Property

Public Property Get dValorDesconto() As Double
    dValorDesconto = mvardValorDesconto
End Property


Public Property Let objTelaAtualizacao(ByVal vData As Object)
    Set mvarobjTelaAtualizacao = vData
End Property

Public Property Get objTelaAtualizacao() As Object
    Set objTelaAtualizacao = mvarobjTelaAtualizacao
End Property

Public Property Set objEvolucao(ByVal vData As Object)
    Set mvarobjEvolucao = vData
End Property

Public Property Let objEvolucao(ByVal vData As Object)
    Set mvarobjEvolucao = vData
End Property

Public Property Get objEvolucao() As Object
    Set objEvolucao = mvarobjEvolucao
End Property

Public Property Let iFilialEmpresaCta(ByVal vData As Integer)
    mvariFilialEmpresaCta = vData
End Property

Public Property Get iFilialEmpresaCta() As Integer
    iFilialEmpresaCta = mvariFilialEmpresaCta
End Property

Public Property Let sLayoutCheque(ByVal vData As String)
    mvarsLayoutCheque = vData
End Property

Public Property Get sLayoutCheque() As String
    sLayoutCheque = mvarsLayoutCheque
End Property

Public Property Let iQualquerPortador(ByVal vData As Integer)
    mvariQualquerPortador = vData
End Property

Public Property Get iQualquerPortador() As Integer
    iQualquerPortador = mvariQualquerPortador
End Property

Public Property Let dtContabil(ByVal vData As Date)
    mvardtContabil = vData
End Property

Public Property Get dtContabil() As Date
    dtContabil = mvardtContabil
End Property

'Public Property Set ColInfoChequePag(ByVal vData As Object)
'    Set mvarcolInfoChequePag = vData
'End Property

Public Property Get colInfoChequePag() As colInfoChequePag
    Set colInfoChequePag = mvarcolInfoChequePag
End Property

''Public Property Set colInfoParcPag(ByVal vData As Object)
''    Set mvarcolInfoParcPag = vData
''End Property

Public Property Get colInfoParcPag() As Collection
    Set colInfoParcPag = mvarcolInfoParcPag
End Property

Public Property Let iPortador(ByVal vData As Integer)
    mvariPortador = vData
End Property

Public Property Get iPortador() As Integer
    iPortador = mvariPortador
End Property

Public Property Let dtVencto(ByVal vData As Date)
    mvardtVencto = vData
End Property

Public Property Get dtVencto() As Date
    dtVencto = mvardtVencto
End Property

Public Property Let dtVenctoDe(ByVal vData As Date)
    mvardtVenctoDe = vData
End Property

Public Property Get dtVenctoDe() As Date
    dtVenctoDe = mvardtVenctoDe
End Property

Public Property Let dtEmissao(ByVal vData As Date)
    mvardtEmissao = vData
End Property

Public Property Get dtEmissao() As Date
    dtEmissao = mvardtEmissao
End Property

Public Property Let lNumCheque(ByVal vData As Long)
    mvarlNumCheque = vData
End Property

Public Property Get lNumCheque() As Long
    lNumCheque = mvarlNumCheque
End Property

Public Property Let iCta(ByVal vData As Integer)
    mvariCta = vData
End Property

Public Property Get iCta() As Integer
    iCta = mvariCta
End Property

Public Property Let lNumImpressao(ByVal vData As Long)
    mvarlNumImpressao = vData
End Property

Public Property Get lNumImpressao() As Long
    lNumImpressao = mvarlNumImpressao
End Property

Public Property Let iIndiceChequeProc(ByVal vData As Integer)
    mvariIndiceChequeProc = vData
End Property

Public Property Get iIndiceChequeProc() As Integer
    iIndiceChequeProc = mvariIndiceChequeProc
End Property

Function GeraContabilizacao(objContabAutomatica As ClassContabAutomatica, vParams As Variant) As Long

    GeraContabilizacao = mvarobjTelaAtualizacao.GeraContabilizacao(objContabAutomatica, vParams)
    
End Function


Private Sub Class_Initialize()
    ReDim Preserve mvaradValorDiferenca(1)
End Sub

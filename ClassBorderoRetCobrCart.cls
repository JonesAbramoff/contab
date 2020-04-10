VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassBorderoRetCobrCart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvariCobrador As Integer
Private mvariCarteiraCobr As Integer
Private mvariCodCarteiraNoBanco As Integer
Private mvarobjBaixaReceber As New ClassBaixaReceber
Private mvarcolBaixaParcRec As New colBaixaParcRec
Private mvarobjMovCCI As New ClassMovContaCorrente
Private mvarcolOcorrRetParcRec As New Collection
Private mvardtDataCredito As Date
Private mvardValorRecebido As Double
Private mvardValorDescontos As Double
Private mvardValorJuros As Double
Private mvardValorTarifas As Double
Private mvardValorCustas As Double
Private mvarcolInfoParcRec As New Collection
Private mvardValorBaixado As Double
Private mvariStatusRecurso As Integer
Private mvardValorTarifasNaoCred As Double

Public Property Let dValorTarifasNaoCred(ByVal vData As Double)
    mvardValorTarifasNaoCred = vData
End Property

Public Property Get dValorTarifasNaoCred() As Double
    dValorTarifasNaoCred = mvardValorTarifasNaoCred
End Property

Public Property Let iStatusRecurso(ByVal vData As Integer)
    mvariStatusRecurso = vData
End Property

Public Property Get iStatusRecurso() As Integer
    iStatusRecurso = mvariStatusRecurso
End Property

Public Property Let dValorBaixado(ByVal vData As Double)
    mvardValorBaixado = vData
End Property

Public Property Get dValorBaixado() As Double
    dValorBaixado = mvardValorBaixado
End Property

Public Property Get colInfoParcRec() As Collection
    Set colInfoParcRec = mvarcolInfoParcRec
End Property

Public Property Let iCodCarteiraNoBanco(ByVal vData As Integer)
    mvariCodCarteiraNoBanco = vData
End Property

Public Property Get iCodCarteiraNoBanco() As Integer
    iCodCarteiraNoBanco = mvariCodCarteiraNoBanco
End Property

Public Property Let iCobrador(ByVal vData As Integer)
    mvariCobrador = vData
End Property

Public Property Get iCobrador() As Integer
    iCobrador = mvariCobrador
End Property

Public Property Let dValorRecebido(ByVal vData As Double)
    mvardValorRecebido = vData
End Property

Public Property Get dValorRecebido() As Double
    dValorRecebido = mvardValorRecebido
End Property

Public Property Let dValorCustas(ByVal vData As Double)
    mvardValorCustas = vData
End Property

Public Property Get dValorCustas() As Double
    dValorCustas = mvardValorCustas
End Property

Public Property Let dValorJuros(ByVal vData As Double)
    mvardValorJuros = vData
End Property

Public Property Get dValorJuros() As Double
    dValorJuros = mvardValorJuros
End Property

Public Property Let dValorDescontos(ByVal vData As Double)
    mvardValorDescontos = vData
End Property

Public Property Get dValorDescontos() As Double
    dValorDescontos = mvardValorDescontos
End Property

Public Property Let dValorTarifas(ByVal vData As Double)
    mvardValorTarifas = vData
End Property

Public Property Get dValorTarifas() As Double
    dValorTarifas = mvardValorTarifas
End Property

Public Property Let iCarteiraCobr(ByVal vData As Integer)
    mvariCarteiraCobr = vData
End Property

Public Property Get iCarteiraCobr() As Integer
    iCarteiraCobr = mvariCarteiraCobr
End Property

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


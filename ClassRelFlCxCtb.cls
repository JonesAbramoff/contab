VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelFlCxCtb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarlNumIntDoc As Long
Private mvarlNumIntRel As Long
Private mvarsSiglaBanco As String
Private mvardtData As Date
Private mvarsDescricao As String
Private mvarsDescConta As String
Private mvariCOAno As Integer
Private mvariCOMes As Integer
Private mvarlNumCheque As Long
Private mvarsDoc As String
Private mvardCredito As Double
Private mvardDebito As Double
Private mvarlNumMovto As Long
Private mvarsTipo As String
Private mvarsTipoMeioPagto As String
Private mvardSaldo As Double
Private mvardRotativo As Double
Private mvarsSiglaDoc As String
Private mvarsHistorico As String
Private mvarsDescContaPai As String
Private mvarlOrdem As Long
Private mvarlSeq As Long

Public Property Let lSeq(ByVal vData As Long)
    mvarlSeq = vData
End Property

Public Property Get lSeq() As Long
    lSeq = mvarlSeq
End Property

Public Property Let lOrdem(ByVal vData As Long)
    mvarlOrdem = vData
End Property

Public Property Get lOrdem() As Long
    lOrdem = mvarlOrdem
End Property

Public Property Let sDescContaPai(ByVal vData As String)
    mvarsDescContaPai = vData
End Property

Public Property Get sDescContaPai() As String
    sDescContaPai = mvarsDescContaPai
End Property

Public Property Let sHistorico(ByVal vData As String)
    mvarsHistorico = vData
End Property

Public Property Get sHistorico() As String
    sHistorico = mvarsHistorico
End Property

Public Property Let sSiglaDoc(ByVal vData As String)
    mvarsSiglaDoc = vData
End Property

Public Property Get sSiglaDoc() As String
    sSiglaDoc = mvarsSiglaDoc
End Property

Public Property Let dRotativo(ByVal vData As Double)
    mvardRotativo = vData
End Property

Public Property Get dRotativo() As Double
    dRotativo = mvardRotativo
End Property

Public Property Let dSaldo(ByVal vData As Double)
    mvardSaldo = vData
End Property

Public Property Get dSaldo() As Double
    dSaldo = mvardSaldo
End Property

Public Property Let sTipoMeioPagto(ByVal vData As String)
    mvarsTipoMeioPagto = vData
End Property

Public Property Get sTipoMeioPagto() As String
    sTipoMeioPagto = mvarsTipoMeioPagto
End Property

Public Property Let sTipo(ByVal vData As String)
    mvarsTipo = vData
End Property

Public Property Get sTipo() As String
    sTipo = mvarsTipo
End Property

Public Property Let lNumMovto(ByVal vData As Long)
    mvarlNumMovto = vData
End Property

Public Property Get lNumMovto() As Long
    lNumMovto = mvarlNumMovto
End Property

Public Property Let dDebito(ByVal vData As Double)
    mvardDebito = vData
End Property

Public Property Get dDebito() As Double
    dDebito = mvardDebito
End Property

Public Property Let dCredito(ByVal vData As Double)
    mvardCredito = vData
End Property

Public Property Get dCredito() As Double
    dCredito = mvardCredito
End Property

Public Property Let sDoc(ByVal vData As String)
    mvarsDoc = vData
End Property

Public Property Get sDoc() As String
    sDoc = mvarsDoc
End Property

Public Property Let lNumCheque(ByVal vData As Long)
    mvarlNumCheque = vData
End Property

Public Property Get lNumCheque() As Long
    lNumCheque = mvarlNumCheque
End Property

Public Property Let iCOMes(ByVal vData As Integer)
    mvariCOMes = vData
End Property

Public Property Get iCOMes() As Integer
    iCOMes = mvariCOMes
End Property

Public Property Let iCOAno(ByVal vData As Integer)
    mvariCOAno = vData
End Property

Public Property Get iCOAno() As Integer
    iCOAno = mvariCOAno
End Property

Public Property Let sDescConta(ByVal vData As String)
    mvarsDescConta = vData
End Property

Public Property Get sDescConta() As String
    sDescConta = mvarsDescConta
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let dtData(ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let sSiglaBanco(ByVal vData As String)
    mvarsSiglaBanco = vData
End Property

Public Property Get sSiglaBanco() As String
    sSiglaBanco = mvarsSiglaBanco
End Property

Public Property Let lNumIntRel(ByVal vData As Long)
    mvarlNumIntRel = vData
End Property

Public Property Get lNumIntRel() As Long
    lNumIntRel = mvarlNumIntRel
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property



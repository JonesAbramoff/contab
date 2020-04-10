VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassNecesProdInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarsProduto As String
Private mvarsVersao As String
Private mvariCompras As Integer
Private mvardQuantNecesInfo As Double
Private mvariClasseUM As Integer
Private mvarsUMNecesInfo As String
Private mvarsUMEstoque As String
Private mvarsUMCompra As String
Private mvardtData1 As Date
Private mvardtData2 As Date

Private mvardQuantNecesCalc As Double
Private mvardQuantNecesCalcEst As Double
Private mvardQuantSaldoLivre As Double

Private mvardQuantSaldoLivreInicial As Double
Private mvardSaldoDisponivel As Double
Private mvardSaldoReservado As Double
Private mvardSaldoPedCompras As Double
Private mvardSaldoReqCompras As Double
Private mvardSaldoOPs As Double
Private mvardSaldoItensPV As Double
Private mvardSaldoEmpenhosDiretos As Double
Private mvardSaldoEmpenhosIndiretos As Double
Private mvardSaldoTransferencias As Double

'##################################################
'INSERIDO POR WAGNER
Private mvariSeq As Integer
Private mvariNivel As Integer
Private mvarsProdutoRaiz As String
Private mvarsVersaoPai As String
Private mvarlNumIntItemOP As Long
Private mvarobjItemOP As New ClassItemOP

Public Property Set objItemOP(ByVal vData As ClassItemOP)
    Set mvarobjItemOP = vData
End Property

Public Property Get objItemOP() As ClassItemOP
    Set objItemOP = mvarobjItemOP
End Property

Public Property Let lNumIntItemOP(ByVal vData As Long)
    mvarlNumIntItemOP = vData
End Property

Public Property Get lNumIntItemOP() As Long
    lNumIntItemOP = mvarlNumIntItemOP
End Property

Public Property Let iSeq(ByVal vData As Integer)
    mvariSeq = vData
End Property

Public Property Get iSeq() As Integer
    iSeq = mvariSeq
End Property

Public Property Let iNivel(ByVal vData As Integer)
    mvariNivel = vData
End Property

Public Property Get iNivel() As Integer
    iNivel = mvariNivel
End Property

Public Property Let sProdutoRaiz(ByVal vData As String)
    mvarsProdutoRaiz = vData
End Property

Public Property Get sProdutoRaiz() As String
    sProdutoRaiz = mvarsProdutoRaiz
End Property

Public Property Let sVersaoPai(ByVal vData As String)
    mvarsVersaoPai = vData
End Property

Public Property Get sVersaoPai() As String
    sVersaoPai = mvarsVersaoPai
End Property
'##################################################

Public Property Let dtData1(ByVal vData As Date)
    mvardtData1 = vData
End Property

Public Property Get dtData1() As Date
    dtData1 = mvardtData1
End Property

Public Property Let dtData2(ByVal vData As Date)
    mvardtData2 = vData
End Property

Public Property Get dtData2() As Date
    dtData2 = mvardtData2
End Property

Public Property Let dQuantNecesCalcEst(ByVal vData As Double)
    mvardQuantNecesCalcEst = vData
End Property

Public Property Get dQuantNecesCalcEst() As Double
    dQuantNecesCalcEst = mvardQuantNecesCalcEst
End Property

Public Property Let dSaldoReservado(ByVal vData As Double)
    mvardSaldoReservado = vData
End Property

Public Property Get dSaldoReservado() As Double
    dSaldoReservado = mvardSaldoReservado
End Property

Public Property Let dQuantSaldoLivreInicial(ByVal vData As Double)
    mvardQuantSaldoLivreInicial = vData
End Property

Public Property Get dQuantSaldoLivreInicial() As Double
    dQuantSaldoLivreInicial = mvardQuantSaldoLivreInicial
End Property

Public Property Let dQuantSaldoLivre(ByVal vData As Double)
    mvardQuantSaldoLivre = vData
End Property

Public Property Get dQuantSaldoLivre() As Double
    dQuantSaldoLivre = mvardQuantSaldoLivre
End Property

Public Property Let iClasseUM(ByVal vData As Integer)
    mvariClasseUM = vData
End Property

Public Property Get iClasseUM() As Integer
    iClasseUM = mvariClasseUM
End Property

Public Property Let iCompras(ByVal vData As Integer)
    mvariCompras = vData
End Property

Public Property Get iCompras() As Integer
    iCompras = mvariCompras
End Property

Public Property Let sUMNecesInfo(ByVal vData As String)
    mvarsUMNecesInfo = vData
End Property

Public Property Get sUMNecesInfo() As String
    sUMNecesInfo = mvarsUMNecesInfo
End Property

Public Property Let sUMEstoque(ByVal vData As String)
    mvarsUMEstoque = vData
End Property

Public Property Get sUMEstoque() As String
    sUMEstoque = mvarsUMEstoque
End Property

Public Property Let sUMCompra(ByVal vData As String)
    mvarsUMCompra = vData
End Property

Public Property Get sUMCompra() As String
    sUMCompra = mvarsUMCompra
End Property

Public Property Let dSaldoTransferencias(ByVal vData As Double)
    mvardSaldoTransferencias = vData
End Property

Public Property Get dSaldoTransferencias() As Double
    dSaldoTransferencias = mvardSaldoTransferencias
End Property

Public Property Let dSaldoEmpenhosIndiretos(ByVal vData As Double)
    mvardSaldoEmpenhosIndiretos = vData
End Property

Public Property Get dSaldoEmpenhosIndiretos() As Double
    dSaldoEmpenhosIndiretos = mvardSaldoEmpenhosIndiretos
End Property

Public Property Let dSaldoEmpenhosDiretos(ByVal vData As Double)
    mvardSaldoEmpenhosDiretos = vData
End Property

Public Property Get dSaldoEmpenhosDiretos() As Double
    dSaldoEmpenhosDiretos = mvardSaldoEmpenhosDiretos
End Property

Public Property Let dSaldoItensPV(ByVal vData As Double)
    mvardSaldoItensPV = vData
End Property

Public Property Get dSaldoItensPV() As Double
    dSaldoItensPV = mvardSaldoItensPV
End Property

Public Property Let dSaldoOPs(ByVal vData As Double)
    mvardSaldoOPs = vData
End Property

Public Property Get dSaldoOPs() As Double
    dSaldoOPs = mvardSaldoOPs
End Property

Public Property Let dSaldoReqCompras(ByVal vData As Double)
    mvardSaldoReqCompras = vData
End Property

Public Property Get dSaldoReqCompras() As Double
    dSaldoReqCompras = mvardSaldoReqCompras
End Property

Public Property Let dSaldoPedCompras(ByVal vData As Double)
    mvardSaldoPedCompras = vData
End Property

Public Property Get dSaldoPedCompras() As Double
    dSaldoPedCompras = mvardSaldoPedCompras
End Property

Public Property Let dSaldoDisponivel(ByVal vData As Double)
    mvardSaldoDisponivel = vData
End Property

Public Property Get dSaldoDisponivel() As Double
    dSaldoDisponivel = mvardSaldoDisponivel
End Property

Public Property Let dQuantNecesCalc(ByVal vData As Double)
    mvardQuantNecesCalc = vData
End Property

Public Property Get dQuantNecesCalc() As Double
    dQuantNecesCalc = mvardQuantNecesCalc
End Property

Public Property Let dQuantNecesInfo(ByVal vData As Double)
    mvardQuantNecesInfo = vData
End Property

Public Property Get dQuantNecesInfo() As Double
    dQuantNecesInfo = mvardQuantNecesInfo
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let sVersao(ByVal vData As String)
    mvarsVersao = vData
End Property

Public Property Get sVersao() As String
    sVersao = mvarsVersao
End Property

Private Sub Class_Initialize()
    mvardtData1 = DATA_NULA
    mvardtData2 = DATA_NULA
End Sub

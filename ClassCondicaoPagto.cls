VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCondicaoPagto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarsDescricao As String
Private mvariEmPagamento As Integer
Private mvariEmRecebimento As Integer
Private mvariNumeroParcelas As Integer
Private mvariDiasParaPrimeiraParcela As Integer
Private mvariIntervaloParcelas As Integer
Private mvariMensal As Integer
Private mvariDiaDoMes As Integer
Private mvardAcrescimoFinanceiro As Double
Private mvariCodigo As Integer
Private mvarsDescReduzida As String
Private mvariModificador As Integer
Private mvariDataFixa As Integer
Private mvariPreCadastrado As Integer

Private mvarcolParcelas As New Collection

'parametros p/calculo das parcelas
Private mvardtDataEmissao As Date
Private mvardtDataEntrega As Date
Private mvardtDataRef As Date
Private mvardValorTotal As Double 'a ser rateado nas parcelas

Private mvariCargoMinimo As Integer
Private mvariFormaPagamento As Integer
Private mvardTetoParcela As Double

Public Property Let dTetoParcela(ByVal vData As Double)
    mvardTetoParcela = vData
End Property

Public Property Get dTetoParcela() As Double
    dTetoParcela = mvardTetoParcela
End Property

Public Property Let iFormaPagamento(ByVal vData As Integer)
    mvariFormaPagamento = vData
End Property

Public Property Get iFormaPagamento() As Integer
    iFormaPagamento = mvariFormaPagamento
End Property

Public Property Let iCargoMinimo(ByVal vData As Integer)
    mvariCargoMinimo = vData
End Property

Public Property Get iCargoMinimo() As Integer
    iCargoMinimo = mvariCargoMinimo
End Property

Public Property Let dValorTotal(ByVal vData As Double)
    mvardValorTotal = vData
End Property

Public Property Get dValorTotal() As Double
    dValorTotal = mvardValorTotal
End Property

Public Property Let dtDataRef(ByVal vData As Date)
    mvardtDataRef = vData
End Property

Public Property Get dtDataRef() As Date
    dtDataRef = mvardtDataRef
End Property

Public Property Let dtDataEmissao(ByVal vData As Date)
    mvardtDataEmissao = vData
End Property

Public Property Get dtDataEmissao() As Date
    dtDataEmissao = mvardtDataEmissao
End Property

Public Property Let dtDataEntrega(ByVal vData As Date)
    mvardtDataEntrega = vData
End Property

Public Property Get dtDataEntrega() As Date
    dtDataEntrega = mvardtDataEntrega
End Property

Public Property Set colParcelas(ByVal vData As Object)
    Set mvarcolParcelas = vData
End Property

Public Property Get colParcelas() As Collection
    Set colParcelas = mvarcolParcelas
End Property

Public Property Let iPreCadastrado(ByVal vData As Integer)
    mvariPreCadastrado = vData
End Property

Public Property Get iPreCadastrado() As Integer
    iPreCadastrado = mvariPreCadastrado
End Property

Public Property Let iDataFixa(ByVal vData As Integer)
    mvariDataFixa = vData
End Property

Public Property Get iDataFixa() As Integer
    iDataFixa = mvariDataFixa
End Property

Public Property Let iModificador(ByVal vData As Integer)
    mvariModificador = vData
End Property

Public Property Get iModificador() As Integer
    iModificador = mvariModificador
End Property

Public Property Let sDescReduzida(ByVal vData As String)
    mvarsDescReduzida = vData
End Property

Public Property Get sDescReduzida() As String
    sDescReduzida = mvarsDescReduzida
End Property

Public Property Let iCodigo(ByVal vData As Integer)
    mvariCodigo = vData
End Property

Public Property Get iCodigo() As Integer
    iCodigo = mvariCodigo
End Property

Public Property Let dAcrescimoFinanceiro(ByVal vData As Double)
    mvardAcrescimoFinanceiro = vData
End Property

Public Property Get dAcrescimoFinanceiro() As Double
    dAcrescimoFinanceiro = mvardAcrescimoFinanceiro
End Property

Public Property Let iDiaDoMes(ByVal vData As Integer)
    mvariDiaDoMes = vData
End Property

Public Property Get iDiaDoMes() As Integer
    iDiaDoMes = mvariDiaDoMes
End Property

Public Property Let iMensal(ByVal vData As Integer)
    mvariMensal = vData
End Property

Public Property Get iMensal() As Integer
    iMensal = mvariMensal
End Property

Public Property Let iIntervaloParcelas(ByVal vData As Integer)
    mvariIntervaloParcelas = vData
End Property

Public Property Get iIntervaloParcelas() As Integer
    iIntervaloParcelas = mvariIntervaloParcelas
End Property

Public Property Let iDiasParaPrimeiraParcela(ByVal vData As Integer)
    mvariDiasParaPrimeiraParcela = vData
End Property

Public Property Get iDiasParaPrimeiraParcela() As Integer
    iDiasParaPrimeiraParcela = mvariDiasParaPrimeiraParcela
End Property

Public Property Let iNumeroParcelas(ByVal vData As Integer)
    mvariNumeroParcelas = vData
End Property

Public Property Get iNumeroParcelas() As Integer
    iNumeroParcelas = mvariNumeroParcelas
End Property

Public Property Let iEmRecebimento(ByVal vData As Integer)
    mvariEmRecebimento = vData
End Property

Public Property Get iEmRecebimento() As Integer
    iEmRecebimento = mvariEmRecebimento
End Property

Public Property Let iEmPagamento(ByVal vData As Integer)
    mvariEmPagamento = vData
End Property

Public Property Get iEmPagamento() As Integer
    iEmPagamento = mvariEmPagamento
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Private Sub Class_Initialize()

    mvardtDataEmissao = DATA_NULA
    mvardtDataEntrega = DATA_NULA
    mvardtDataRef = DATA_NULA

End Sub

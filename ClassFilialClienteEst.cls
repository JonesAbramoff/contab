VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFilialClienteEst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariFilialEmpresa As Integer
Private mvarlCodCliente As Long
Private mvariCodFilial As Integer
Private mvarlNumeroCompras As Long
Private mvardMediaCompra As Double
Private mvardtDataPrimeiraCompra As Date
Private mvardtDataUltimaCompra As Date
Private mvardtDataUltimoContato As Date
Private mvardValorAcumuladoCompras As Double
Private mvarlMediaAtraso As Long
Private mvarlMaiorAtraso As Long
Private mvardSaldoTitulos As Double
Private mvardSaldoPedidosLiberados As Double
Private mvardSaldoAtrasados As Double
Private mvardValPagtosAtraso As Double
Private mvariNumChequesDevolvidos As Integer
Private mvardtDataUltChequeDevolvido As Date
Private mvardSaldoDuplicatas As Double
Private mvardValorNFsNaoFaturadas As Double
Private mvardValorTotalCR As Double
Private mvardValorTotalCRComProtesto As Double
Private mvardValorTotalCREmCartorio As Double
Private mvardValorTotalNaoPagoAteVenc As Double

Public Property Let dValorTotalNaoPagoAteVenc(ByVal vData As Double)
    mvardValorTotalNaoPagoAteVenc = vData
End Property

Public Property Get dValorTotalNaoPagoAteVenc() As Double
    dValorTotalNaoPagoAteVenc = mvardValorTotalNaoPagoAteVenc
End Property

Public Property Let dValorTotalCREmCartorio(ByVal vData As Double)
    mvardValorTotalCREmCartorio = vData
End Property

Public Property Get dValorTotalCREmCartorio() As Double
    dValorTotalCREmCartorio = mvardValorTotalCREmCartorio
End Property

Public Property Let dValorTotalCRComProtesto(ByVal vData As Double)
    mvardValorTotalCRComProtesto = vData
End Property

Public Property Get dValorTotalCRComProtesto() As Double
    dValorTotalCRComProtesto = mvardValorTotalCRComProtesto
End Property

Public Property Let dValorTotalCR(ByVal vData As Double)
    mvardValorTotalCR = vData
End Property

Public Property Get dValorTotalCR() As Double
    dValorTotalCR = mvardValorTotalCR
End Property

Public Property Let dtDataUltimoContato(ByVal vData As Date)
    mvardtDataUltimoContato = vData
End Property

Public Property Get dtDataUltimoContato() As Date
    dtDataUltimoContato = mvardtDataUltimoContato
End Property

Public Property Let dValorNFsNaoFaturadas(ByVal vData As Double)
    mvardValorNFsNaoFaturadas = vData
End Property

Public Property Get dValorNFsNaoFaturadas() As Double
    dValorNFsNaoFaturadas = mvardValorNFsNaoFaturadas
End Property

Public Property Let dValorAcumuladoCompras(ByVal vData As Double)
    mvardValorAcumuladoCompras = vData
End Property

Public Property Get dValorAcumuladoCompras() As Double
    dValorAcumuladoCompras = mvardValorAcumuladoCompras
End Property

Public Property Let dtDataUltimaCompra(ByVal vData As Date)
    mvardtDataUltimaCompra = vData
End Property

Public Property Get dtDataUltimaCompra() As Date
    dtDataUltimaCompra = mvardtDataUltimaCompra
End Property

Public Property Let dtDataPrimeiraCompra(ByVal vData As Date)
    mvardtDataPrimeiraCompra = vData
End Property

Public Property Get dtDataPrimeiraCompra() As Date
    dtDataPrimeiraCompra = mvardtDataPrimeiraCompra
End Property

Public Property Let dMediaCompra(ByVal vData As Double)
    mvardMediaCompra = vData
End Property

Public Property Get dMediaCompra() As Double
    dMediaCompra = mvardMediaCompra
End Property

Public Property Let lNumeroCompras(ByVal vData As Long)
    mvarlNumeroCompras = vData
End Property

Public Property Get lNumeroCompras() As Long
    lNumeroCompras = mvarlNumeroCompras
End Property

Public Property Let dSaldoDuplicatas(ByVal vData As Double)
    mvardSaldoDuplicatas = vData
End Property

Public Property Get dSaldoDuplicatas() As Double
    dSaldoDuplicatas = mvardSaldoDuplicatas
End Property

Public Property Let dtDataUltChequeDevolvido(ByVal vData As Date)
    mvardtDataUltChequeDevolvido = vData
End Property

Public Property Get dtDataUltChequeDevolvido() As Date
    dtDataUltChequeDevolvido = mvardtDataUltChequeDevolvido
End Property

Public Property Let iNumChequesDevolvidos(ByVal vData As Integer)
    mvariNumChequesDevolvidos = vData
End Property

Public Property Get iNumChequesDevolvidos() As Integer
    iNumChequesDevolvidos = mvariNumChequesDevolvidos
End Property

Public Property Let dValPagtosAtraso(ByVal vData As Double)
    mvardValPagtosAtraso = vData
End Property

Public Property Get dValPagtosAtraso() As Double
    dValPagtosAtraso = mvardValPagtosAtraso
End Property

Public Property Let dSaldoAtrasados(ByVal vData As Double)
    mvardSaldoAtrasados = vData
End Property

Public Property Get dSaldoAtrasados() As Double
    dSaldoAtrasados = mvardSaldoAtrasados
End Property

Public Property Let dSaldoPedidosLiberados(ByVal vData As Double)
    mvardSaldoPedidosLiberados = vData
End Property

Public Property Get dSaldoPedidosLiberados() As Double
    dSaldoPedidosLiberados = mvardSaldoPedidosLiberados
End Property

Public Property Let dSaldoTitulos(ByVal vData As Double)
    mvardSaldoTitulos = vData
End Property

Public Property Get dSaldoTitulos() As Double
    dSaldoTitulos = mvardSaldoTitulos
End Property

Public Property Let lMaiorAtraso(ByVal vData As Long)
    mvarlMaiorAtraso = vData
End Property

Public Property Get lMaiorAtraso() As Long
    lMaiorAtraso = mvarlMaiorAtraso
End Property

Public Property Let lMediaAtraso(ByVal vData As Long)
    mvarlMediaAtraso = vData
End Property

Public Property Get lMediaAtraso() As Long
    lMediaAtraso = mvarlMediaAtraso
End Property

Public Property Let iCodFilial(ByVal vData As Integer)
    mvariCodFilial = vData
End Property

Public Property Get iCodFilial() As Integer
    iCodFilial = mvariCodFilial
End Property

Public Property Let lCodCliente(ByVal vData As Long)
    mvarlCodCliente = vData
End Property

Public Property Get lCodCliente() As Long
    lCodCliente = mvarlCodCliente
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property



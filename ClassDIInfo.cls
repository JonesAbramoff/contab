VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassDIInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDoc As Long
Private mvarsNumero As String
Private mvardtData As Date
Private mvariFilialEmpresa As Integer
Private mvariStatus As Integer
Private mvarsDescricao As String
Private mvarlFornTrading As Long
Private mvariFilialFornTrading As Integer
Private mvarsProcessoTrading As String
Private mvariMoeda1 As Integer
Private mvardTaxaMoeda1 As Double
Private mvariMoeda2 As Integer
Private mvardTaxaMoeda2 As Double
Private mvardPesoBrutoKG As Double
Private mvardPesoLiqKG As Double
Private mvardValorMercadoriaMoeda As Double
Private mvardValorFreteInternacMoeda As Double
Private mvardValorSeguroInternacMoeda As Double
Private mvardValorMercadoriaEmReal As Double
Private mvardValorFreteInternacEmReal As Double
Private mvardValorSeguroInternacEmReal As Double

Public dIIValor As Double
Public dIPIValor As Double
Public dPISValor As Double
Public dCOFINSValor As Double
Public dICMSValor As Double

Public dValorDespesas As Double

Public colAdicoesDI As New Collection
Public colDespesasDI As New Collection 'de ClassImportCompl
Private mvardtDataDesembaraco As Date
Private mvarsUFDesembaraco As String
Private mvarsLocalDesembaraco As String
Private mvariMoedaMercadoria As Integer
Private mvariMoedaFrete As Integer
Private mvariMoedaSeguro As Integer
Private mvariMoedaItens As Integer
Private mvarsCodExportador As String
Public colItensPC As New Collection

'nfe 3.10
Public iViaTransp As Integer
Public iIntermedio As Integer
Public sCNPJAdquir As String
Public sUFAdquir As String

Public Property Let sCodExportador(ByVal vData As String)
    mvarsCodExportador = vData
End Property

Public Property Get sCodExportador() As String
    sCodExportador = mvarsCodExportador
End Property

Public Property Let iMoedaItens(ByVal vData As Integer)
    mvariMoedaItens = vData
End Property

Public Property Get iMoedaItens() As Integer
    iMoedaItens = mvariMoedaItens
End Property

Public Property Let iMoedaSeguro(ByVal vData As Integer)
    mvariMoedaSeguro = vData
End Property

Public Property Get iMoedaSeguro() As Integer
    iMoedaSeguro = mvariMoedaSeguro
End Property

Public Property Let iMoedaFrete(ByVal vData As Integer)
    mvariMoedaFrete = vData
End Property

Public Property Get iMoedaFrete() As Integer
    iMoedaFrete = mvariMoedaFrete
End Property

Public Property Let iMoedaMercadoria(ByVal vData As Integer)
    mvariMoedaMercadoria = vData
End Property

Public Property Get iMoedaMercadoria() As Integer
    iMoedaMercadoria = mvariMoedaMercadoria
End Property

Public Property Let sLocalDesembaraco(ByVal vData As String)
    mvarsLocalDesembaraco = vData
End Property

Public Property Get sLocalDesembaraco() As String
    sLocalDesembaraco = mvarsLocalDesembaraco
End Property

Public Property Let sUFDesembaraco(ByVal vData As String)
    mvarsUFDesembaraco = vData
End Property

Public Property Get sUFDesembaraco() As String
    sUFDesembaraco = mvarsUFDesembaraco
End Property

Public Property Let dtDataDesembaraco(ByVal vDataDesembaraco As Date)
    mvardtDataDesembaraco = vDataDesembaraco
End Property

Public Property Get dtDataDesembaraco() As Date
    dtDataDesembaraco = mvardtDataDesembaraco
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let sNumero(ByVal vData As String)
    mvarsNumero = vData
End Property

Public Property Get sNumero() As String
    sNumero = mvarsNumero
End Property

Public Property Let dtData(ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let iStatus(ByVal vData As Integer)
    mvariStatus = vData
End Property

Public Property Get iStatus() As Integer
    iStatus = mvariStatus
End Property

Public Property Let lFornTrading(ByVal vData As Long)
    mvarlFornTrading = vData
End Property

Public Property Get lFornTrading() As Long
    lFornTrading = mvarlFornTrading
End Property

Public Property Let iFilialFornTrading(ByVal vData As Integer)
    mvariFilialFornTrading = vData
End Property

Public Property Get iFilialFornTrading() As Integer
    iFilialFornTrading = mvariFilialFornTrading
End Property

Public Property Let sProcessoTrading(ByVal vData As String)
    mvarsProcessoTrading = vData
End Property

Public Property Get sProcessoTrading() As String
    sProcessoTrading = mvarsProcessoTrading
End Property

Public Property Let iMoeda2(ByVal vData As Integer)
    mvariMoeda2 = vData
End Property

Public Property Get iMoeda2() As Integer
    iMoeda2 = mvariMoeda2
End Property

Public Property Let dTaxaMoeda2(ByVal vData As Double)
    mvardTaxaMoeda2 = vData
End Property

Public Property Get dTaxaMoeda2() As Double
    dTaxaMoeda2 = mvardTaxaMoeda2
End Property

Public Property Let iMoeda1(ByVal vData As Integer)
    mvariMoeda1 = vData
End Property

Public Property Get iMoeda1() As Integer
    iMoeda1 = mvariMoeda1
End Property

Public Property Let dTaxaMoeda1(ByVal vData As Double)
    mvardTaxaMoeda1 = vData
End Property

Public Property Get dTaxaMoeda1() As Double
    dTaxaMoeda1 = mvardTaxaMoeda1
End Property

Public Property Let dPesoBrutoKG(ByVal vData As Double)
    mvardPesoBrutoKG = vData
End Property

Public Property Get dPesoBrutoKG() As Double
    dPesoBrutoKG = mvardPesoBrutoKG
End Property

Public Property Let dPesoLiqKG(ByVal vData As Double)
    mvardPesoLiqKG = vData
End Property

Public Property Get dPesoLiqKG() As Double
    dPesoLiqKG = mvardPesoLiqKG
End Property

Public Property Let dValorMercadoriaMoeda(ByVal vData As Double)
    mvardValorMercadoriaMoeda = vData
End Property

Public Property Get dValorMercadoriaMoeda() As Double
    dValorMercadoriaMoeda = mvardValorMercadoriaMoeda
End Property

Public Property Let dValorFreteInternacMoeda(ByVal vData As Double)
    mvardValorFreteInternacMoeda = vData
End Property

Public Property Get dValorFreteInternacMoeda() As Double
    dValorFreteInternacMoeda = mvardValorFreteInternacMoeda
End Property

Public Property Let dValorSeguroInternacMoeda(ByVal vData As Double)
    mvardValorSeguroInternacMoeda = vData
End Property

Public Property Get dValorSeguroInternacMoeda() As Double
    dValorSeguroInternacMoeda = mvardValorSeguroInternacMoeda
End Property

Public Property Let dValorMercadoriaEmReal(ByVal vData As Double)
    mvardValorMercadoriaEmReal = vData
End Property

Public Property Get dValorMercadoriaEmReal() As Double
    dValorMercadoriaEmReal = mvardValorMercadoriaEmReal
End Property

Public Property Let dValorFreteInternacEmReal(ByVal vData As Double)
    mvardValorFreteInternacEmReal = vData
End Property

Public Property Get dValorFreteInternacEmReal() As Double
    dValorFreteInternacEmReal = mvardValorFreteInternacEmReal
End Property

Public Property Let dValorSeguroInternacEmReal(ByVal vData As Double)
    mvardValorSeguroInternacEmReal = vData
End Property

Public Property Get dValorSeguroInternacEmReal() As Double
    dValorSeguroInternacEmReal = mvardValorSeguroInternacEmReal
End Property

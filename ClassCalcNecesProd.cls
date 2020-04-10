VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCalcNecesProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvariFilialEmpresa As Integer

Private mvariSemSaldoProdComprado As Integer 'desconsidera saldo existente p/produtos comprados
Private mvariSemDatasProdComprado As Integer
Private mvariSemSaldoProdProduzido As Integer 'desconsidera saldo existente p/produtos produzidos
Private mvariSemDatasProdProduzido As Integer
Private mvariSemSaldoPedCompras As Integer 'desconsidera saldo de pedidos de compra
Private mvariSemDatasPedCompras As Integer
Private mvariSemSaldoReqCompras As Integer 'desconsidera saldo de requisicoes de compra
Private mvariSemDatasReqCompras As Integer
Private mvariSemSaldoOrdensProd As Integer 'desconsidera saldo de ops
Private mvariSemDatasOrdensProd As Integer
Private mvariSemSaldoItensPV As Integer 'desconsidera saldo de itens de pv
Private mvariSemDatasItensPV As Integer
Private mvariSemSaldoEmpenhosDiretos As Integer 'desconsidera saldo de empenhos diretos
Private mvariSemDatasEmpenhosDiretos As Integer
Private mvariSemSaldoEmpenhosIndiretos As Integer 'desconsidera empenhos indiretos
Private mvariSemDatasEmpenhosIndiretos As Integer
Private mvariSemSaldoTransferencias As Integer 'desconsidera saldo de produtos em transferencia
Private mvariSemDatasTransferencias As Integer

Public colNecesInfProd As New Collection 'necessidades informadas de produção
Public colNecesCalcProd As New Collection 'necessidades calculadas de producao ou compra
Private mvariIgnoraEst As Integer

Public Property Let iIgnoraEst(ByVal vData As Integer)
    mvariIgnoraEst = vData
End Property

Public Property Get iIgnoraEst() As Integer
    iIgnoraEst = mvariIgnoraEst
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let iSemDatasTransferencias(ByVal vData As Integer)
    mvariSemDatasTransferencias = vData
End Property

Public Property Get iSemDatasTransferencias() As Integer
    iSemDatasTransferencias = mvariSemDatasTransferencias
End Property

Public Property Let iSemSaldoTransferencias(ByVal vData As Integer)
    mvariSemSaldoTransferencias = vData
End Property

Public Property Get iSemSaldoTransferencias() As Integer
    iSemSaldoTransferencias = mvariSemSaldoTransferencias
End Property

Public Property Let iSemDatasEmpenhosIndiretos(ByVal vData As Integer)
    mvariSemDatasEmpenhosIndiretos = vData
End Property

Public Property Get iSemDatasEmpenhosIndiretos() As Integer
    iSemDatasEmpenhosIndiretos = mvariSemDatasEmpenhosIndiretos
End Property

Public Property Let iSemSaldoEmpenhosIndiretos(ByVal vData As Integer)
    mvariSemSaldoEmpenhosIndiretos = vData
End Property

Public Property Get iSemSaldoEmpenhosIndiretos() As Integer
    iSemSaldoEmpenhosIndiretos = mvariSemSaldoEmpenhosIndiretos
End Property

Public Property Let iSemDatasEmpenhosDiretos(ByVal vData As Integer)
    mvariSemDatasEmpenhosDiretos = vData
End Property

Public Property Get iSemDatasEmpenhosDiretos() As Integer
    iSemDatasEmpenhosDiretos = mvariSemDatasEmpenhosDiretos
End Property

Public Property Let iSemSaldoEmpenhosDiretos(ByVal vData As Integer)
    mvariSemSaldoEmpenhosDiretos = vData
End Property

Public Property Get iSemSaldoEmpenhosDiretos() As Integer
    iSemSaldoEmpenhosDiretos = mvariSemSaldoEmpenhosDiretos
End Property

Public Property Let iSemDatasItensPV(ByVal vData As Integer)
    mvariSemDatasItensPV = vData
End Property

Public Property Get iSemDatasItensPV() As Integer
    iSemDatasItensPV = mvariSemDatasItensPV
End Property

Public Property Let iSemSaldoItensPV(ByVal vData As Integer)
    mvariSemSaldoItensPV = vData
End Property

Public Property Get iSemSaldoItensPV() As Integer
    iSemSaldoItensPV = mvariSemSaldoItensPV
End Property

Public Property Let iSemDatasOrdensProd(ByVal vData As Integer)
    mvariSemDatasOrdensProd = vData
End Property

Public Property Get iSemDatasOrdensProd() As Integer
    iSemDatasOrdensProd = mvariSemDatasOrdensProd
End Property

Public Property Let iSemSaldoOrdensProd(ByVal vData As Integer)
    mvariSemSaldoOrdensProd = vData
End Property

Public Property Get iSemSaldoOrdensProd() As Integer
    iSemSaldoOrdensProd = mvariSemSaldoOrdensProd
End Property

Public Property Let iSemDatasReqCompras(ByVal vData As Integer)
    mvariSemDatasReqCompras = vData
End Property

Public Property Get iSemDatasReqCompras() As Integer
    iSemDatasReqCompras = mvariSemDatasReqCompras
End Property

Public Property Let iSemSaldoReqCompras(ByVal vData As Integer)
    mvariSemSaldoReqCompras = vData
End Property

Public Property Get iSemSaldoReqCompras() As Integer
    iSemSaldoReqCompras = mvariSemSaldoReqCompras
End Property

Public Property Let iSemDatasPedCompras(ByVal vData As Integer)
    mvariSemDatasPedCompras = vData
End Property

Public Property Get iSemDatasPedCompras() As Integer
    iSemDatasPedCompras = mvariSemDatasPedCompras
End Property

Public Property Let iSemSaldoPedCompras(ByVal vData As Integer)
    mvariSemSaldoPedCompras = vData
End Property

Public Property Get iSemSaldoPedCompras() As Integer
    iSemSaldoPedCompras = mvariSemSaldoPedCompras
End Property

Public Property Let iSemDatasProdProduzido(ByVal vData As Integer)
    mvariSemDatasProdProduzido = vData
End Property

Public Property Get iSemDatasProdProduzido() As Integer
    iSemDatasProdProduzido = mvariSemDatasProdProduzido
End Property

Public Property Let iSemSaldoProdProduzido(ByVal vData As Integer)
    mvariSemSaldoProdProduzido = vData
End Property

Public Property Get iSemSaldoProdProduzido() As Integer
    iSemSaldoProdProduzido = mvariSemSaldoProdProduzido
End Property

Public Property Let iSemDatasProdComprado(ByVal vData As Integer)
    mvariSemDatasProdComprado = vData
End Property

Public Property Get iSemDatasProdComprado() As Integer
    iSemDatasProdComprado = mvariSemDatasProdComprado
End Property

Public Property Let iSemSaldoProdComprado(ByVal vData As Integer)
    mvariSemSaldoProdComprado = vData
End Property

Public Property Get iSemSaldoProdComprado() As Integer
    iSemSaldoProdComprado = mvariSemSaldoProdComprado
End Property



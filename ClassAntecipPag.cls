VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassAntecipPag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlFornecedor As Long
Private mvariFilial As Integer
Private mvariFilialEmpresa As Integer
Private mvariCodConta As Integer
Private mvarlSequencial As Long
Private mvardtData As Date
Private mvardValor As Double
Private mvariTipoMeioPagto As Integer
Private mvarlNumero As Long
Private mvarsHistorico As String
Private mvarlNumMovto As Long
Private mvardSaldoNaoApropriado As Double
Private mvarlNumIntPag As Long
Private mvarsDescricaoMeioPagto As String
Private mvardValorBaixado As Double
Private mvariExcluido As Integer
Private mvariFilialPedCompra As Integer
Private mvarlNumPedCompra As Long
Private mvarsNatureza As String

Public Property Let sNatureza(ByVal vData As String)
    mvarsNatureza = vData
End Property

Public Property Get sNatureza() As String
    sNatureza = mvarsNatureza
End Property

Public Property Let lNumPedCompra(ByVal vData As Long)
    mvarlNumPedCompra = vData
End Property

Public Property Get lNumPedCompra() As Long
    lNumPedCompra = mvarlNumPedCompra
End Property

Public Property Let iFilialPedCompra(ByVal vData As Integer)
    mvariFilialPedCompra = vData
End Property

Public Property Get iFilialPedCompra() As Integer
    iFilialPedCompra = mvariFilialPedCompra
End Property

Public Property Let iExcluido(ByVal vData As Integer)
Attribute iExcluido.VB_Description = "0=Valido 1=Excluido"
    mvariExcluido = vData
End Property

Public Property Get iExcluido() As Integer
    iExcluido = mvariExcluido
End Property

Public Property Let dValorBaixado(ByVal vData As Double)
Attribute dValorBaixado.VB_Description = "Valor a Baixar quando  Baixa a Pagar é feita por Pagto Antecipado."
    mvardValorBaixado = vData
End Property

Public Property Get dValorBaixado() As Double
    dValorBaixado = mvardValorBaixado
End Property

Public Property Let sDescricaoMeioPagto(ByVal vData As String)
    mvarsDescricaoMeioPagto = vData
End Property

Public Property Get sDescricaoMeioPagto() As String
    sDescricaoMeioPagto = mvarsDescricaoMeioPagto
End Property

Public Property Let lNumIntPag(ByVal vData As Long)
    mvarlNumIntPag = vData
End Property

Public Property Get lNumIntPag() As Long
    lNumIntPag = mvarlNumIntPag
End Property

Public Property Let dSaldoNaoApropriado(ByVal vData As Double)
    mvardSaldoNaoApropriado = vData
End Property

Public Property Get dSaldoNaoApropriado() As Double
    dSaldoNaoApropriado = mvardSaldoNaoApropriado
End Property

Public Property Let lNumMovto(ByVal vData As Long)
    mvarlNumMovto = vData
End Property

Public Property Get lNumMovto() As Long
    lNumMovto = mvarlNumMovto
End Property

Public Property Let sHistorico(ByVal vData As String)
    mvarsHistorico = vData
End Property

Public Property Get sHistorico() As String
    sHistorico = mvarsHistorico
End Property

Public Property Let lNumero(ByVal vData As Long)
    mvarlNumero = vData
End Property

Public Property Get lNumero() As Long
    lNumero = mvarlNumero
End Property

Public Property Let iTipoMeioPagto(ByVal vData As Integer)
    mvariTipoMeioPagto = vData
End Property

Public Property Get iTipoMeioPagto() As Integer
    iTipoMeioPagto = mvariTipoMeioPagto
End Property

Public Property Let dValor(ByVal vData As Double)
    mvardValor = vData
End Property

Public Property Get dValor() As Double
    dValor = mvardValor
End Property

Public Property Let dtData(ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let lSequencial(ByVal vData As Long)
    mvarlSequencial = vData
End Property

Public Property Get lSequencial() As Long
    lSequencial = mvarlSequencial
End Property

Public Property Let iCodConta(ByVal vData As Integer)
    mvariCodConta = vData
End Property

Public Property Get iCodConta() As Integer
    iCodConta = mvariCodConta
End Property

Public Property Let iFilial(ByVal vData As Integer)
    mvariFilial = vData
End Property

Public Property Get iFilial() As Integer
    iFilial = mvariFilial
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let lFornecedor(ByVal vData As Long)
    mvarlFornecedor = vData
End Property

Public Property Get lFornecedor() As Long
    lFornecedor = mvarlFornecedor
End Property




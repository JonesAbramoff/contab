VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassItensDeContrato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlNumIntDoc As Long
Private mvarlNumIntContrato As Long
Private mvariSeq As Integer
Private mvariCobrar As Integer
Private mvarsProduto As String
Private mvarsDescProd As String
Private mvardQuantidade As Double
Private mvarsUM As String
Private mvardValor As Double
Private mvariMedicao As Integer
Private mvardtDataIniCobranca As Date
Private mvardtDataProxCobranca As Date
Private mvariPeriodicidade As Integer
Private mvariCondPagto As Integer
Private mvarsObservacao As String
Private mvarsContrato As String
Private mvardtDataRefIni As Date
Private mvardtDataRefFim As Date
Private mvariFilialEmpresa As Integer
Private mvariQtdeParcelas As Integer
Private mvariUltParcCobrada As Integer

Public Property Let iUltParcCobrada(ByVal vData As Integer)
    mvariUltParcCobrada = vData
End Property

Public Property Get iUltParcCobrada() As Integer
     iUltParcCobrada = mvariUltParcCobrada
End Property

Public Property Let iQtdeParcelas(ByVal vData As Integer)
    mvariQtdeParcelas = vData
End Property

Public Property Get iQtdeParcelas() As Integer
     iQtdeParcelas = mvariQtdeParcelas
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
     iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let dtDataRefFim(ByVal vData As Date)
    mvardtDataRefFim = vData
End Property

Public Property Get dtDataRefFim() As Date
    dtDataRefFim = mvardtDataRefFim
End Property

Public Property Let dtDataRefIni(ByVal vData As Date)
    mvardtDataRefIni = vData
End Property

Public Property Get dtDataRefIni() As Date
    dtDataRefIni = mvardtDataRefIni
End Property

Public Property Let sContrato(ByVal vData As String)
    mvarsContrato = vData
End Property

Public Property Get sContrato() As String
    sContrato = mvarsContrato
End Property

Public Property Let sObservacao(ByVal vData As String)
    mvarsObservacao = vData
End Property

Public Property Get sObservacao() As String
     sObservacao = mvarsObservacao
End Property

Public Property Let iCondPagto(ByVal vData As Integer)
    mvariCondPagto = vData
End Property

Public Property Get iCondPagto() As Integer
     iCondPagto = mvariCondPagto
End Property

Public Property Let iSeq(ByVal vData As Integer)
    mvariSeq = vData
End Property

Public Property Get iSeq() As Integer
     iSeq = mvariSeq
End Property

Public Property Let iPeriodicidade(ByVal vData As Integer)
    mvariPeriodicidade = vData
End Property

Public Property Get iPeriodicidade() As Integer
     iPeriodicidade = mvariPeriodicidade
End Property

Public Property Let dtDataIniCobranca(ByVal vData As Date)
    mvardtDataIniCobranca = vData
End Property

Public Property Get dtDataIniCobranca() As Date
    dtDataIniCobranca = mvardtDataIniCobranca
End Property

Public Property Let dtDataProxCobranca(ByVal vData As Date)
    mvardtDataProxCobranca = vData
End Property

Public Property Get dtDataProxCobranca() As Date
    dtDataProxCobranca = mvardtDataProxCobranca
End Property

Public Property Let iMedicao(ByVal vData As Integer)
    mvariMedicao = vData
End Property

Public Property Get iMedicao() As Integer
     iMedicao = mvariMedicao
End Property

Public Property Let dValor(ByVal vData As Double)
    mvardValor = vData
End Property

Public Property Get dValor() As Double
     dValor = mvardValor
End Property

Public Property Let sUM(ByVal vData As String)
    mvarsUM = vData
End Property

Public Property Get sUM() As String
    sUM = mvarsUM
End Property

Public Property Let dQuantidade(ByVal vData As Double)
    mvardQuantidade = vData
End Property

Public Property Get dQuantidade() As Double
    dQuantidade = mvardQuantidade
End Property

Public Property Let sDescProd(ByVal vData As String)
    mvarsDescProd = vData
End Property

Public Property Get sDescProd() As String
    sDescProd = mvarsDescProd
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
     sProduto = mvarsProduto
End Property

Public Property Let iCobrar(ByVal vData As Integer)
    mvariCobrar = vData
End Property

Public Property Get iCobrar() As Integer
    iCobrar = mvariCobrar
End Property

Public Property Let lNumIntContrato(ByVal vData As Long)
    mvarlNumIntContrato = vData
End Property

Public Property Get lNumIntContrato() As Long
    lNumIntContrato = mvarlNumIntContrato
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property




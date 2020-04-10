VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRastreamentoLoteTeste"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntRastroLote As Long
Private mvariTesteCodigo As Integer
Private mvariSeqGrid As Integer
Private mvarsTesteEspecificacao As String
Private mvariTesteTipoResultado As Integer
Private mvardTesteLimiteDe As Double
Private mvardTesteLimiteAte As Double
Private mvarsTesteMetodoUsado As String
Private mvarsTesteObservacao As String
Private mvarsRegistroAnaliseID As String
Private mvardtRegistroAnaliseData As Date
Private mvariResultadoNaoConforme As Integer
Private mvarsResultadoValor As String
Private mvarsResultadoObservacao As String
Private mvariTesteNoCertificado As Integer

Public Property Let iTesteNoCertificado(ByVal vData As Integer)
    mvariTesteNoCertificado = vData
End Property

Public Property Get iTesteNoCertificado() As Integer
    iTesteNoCertificado = mvariTesteNoCertificado
End Property

Public Property Let iSeqGrid(ByVal vData As Integer)
    mvariSeqGrid = vData
End Property

Public Property Get iSeqGrid() As Integer
    iSeqGrid = mvariSeqGrid
End Property

Public Property Let lNumIntRastroLote(ByVal vData As Long)
    mvarlNumIntRastroLote = vData
End Property

Public Property Get lNumIntRastroLote() As Long
    lNumIntRastroLote = mvarlNumIntRastroLote
End Property

Public Property Let iTesteCodigo(ByVal vData As Integer)
    mvariTesteCodigo = vData
End Property

Public Property Get iTesteCodigo() As Integer
    iTesteCodigo = mvariTesteCodigo
End Property

Public Property Let sTesteEspecificacao(ByVal vData As String)
    mvarsTesteEspecificacao = vData
End Property

Public Property Get sTesteEspecificacao() As String
    sTesteEspecificacao = mvarsTesteEspecificacao
End Property

Public Property Let iTesteTipoResultado(ByVal vData As Integer)
    mvariTesteTipoResultado = vData
End Property

Public Property Get iTesteTipoResultado() As Integer
    iTesteTipoResultado = mvariTesteTipoResultado
End Property

Public Property Let dTesteLimiteDe(ByVal vData As Double)
    mvardTesteLimiteDe = vData
End Property

Public Property Get dTesteLimiteDe() As Double
    dTesteLimiteDe = mvardTesteLimiteDe
End Property

Public Property Let dTesteLimiteAte(ByVal vData As Double)
    mvardTesteLimiteAte = vData
End Property

Public Property Get dTesteLimiteAte() As Double
    dTesteLimiteAte = mvardTesteLimiteAte
End Property

Public Property Let sTesteMetodoUsado(ByVal vData As String)
    mvarsTesteMetodoUsado = vData
End Property

Public Property Get sTesteMetodoUsado() As String
    sTesteMetodoUsado = mvarsTesteMetodoUsado
End Property

Public Property Let sTesteObservacao(ByVal vData As String)
    mvarsTesteObservacao = vData
End Property

Public Property Get sTesteObservacao() As String
    sTesteObservacao = mvarsTesteObservacao
End Property

Public Property Let sRegistroAnaliseID(ByVal vData As String)
    mvarsRegistroAnaliseID = vData
End Property

Public Property Get sRegistroAnaliseID() As String
    sRegistroAnaliseID = mvarsRegistroAnaliseID
End Property

Public Property Let dtRegistroAnaliseData(ByVal vData As Date)
    mvardtRegistroAnaliseData = vData
End Property

Public Property Get dtRegistroAnaliseData() As Date
    dtRegistroAnaliseData = mvardtRegistroAnaliseData
End Property

Public Property Let iResultadoNaoConforme(ByVal vData As Integer)
    mvariResultadoNaoConforme = vData
End Property

Public Property Get iResultadoNaoConforme() As Integer
    iResultadoNaoConforme = mvariResultadoNaoConforme
End Property

Public Property Let sResultadoValor(ByVal vData As String)
    mvarsResultadoValor = vData
End Property

Public Property Get sResultadoValor() As String
    sResultadoValor = mvarsResultadoValor
End Property

Public Property Let sResultadoObservacao(ByVal vData As String)
    mvarsResultadoObservacao = vData
End Property

Public Property Get sResultadoObservacao() As String
    sResultadoObservacao = mvarsResultadoObservacao
End Property

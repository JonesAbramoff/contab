VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRastreamentoMovto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public objInfoUsu As Object

Private mvarlNumIntDoc As Long
Private mvarsProduto As String
Private mvariTipoDocOrigem As Integer
Private mvarlNumIntDocOrigem As Long
Private mvarsLote As String
Private mvariFilialOP As Integer
Private mvardQuantidade As Double
Private mvarlNumIntDocLote As Long
Private mvarsSiglaUM As String  'sigla da UM com que o movimento de rastreamento foi cadastrado
Private mvardQuantidadeEst As Double ' quantidade em unidade de medida do estoque

'#############################################
'Inserido por Wagner 10/03/2006
Private mvardQuantidadeSerie As Double
Private mvarlNumIntDocLoteSerieIni As Long
Private mvarcolRastreamentoSerie As New Collection
Private mvarsLocalizacao As String

Public Property Let sLocalizacao(ByVal vData As String)
    mvarsLocalizacao = vData
End Property

Public Property Get sLocalizacao() As String
    sLocalizacao = mvarsLocalizacao
End Property

Public Property Set colRastreamentoSerie(ByVal vData As Object)
    Set mvarcolRastreamentoSerie = vData
End Property

Public Property Get colRastreamentoSerie() As Collection
    Set colRastreamentoSerie = mvarcolRastreamentoSerie
End Property

Public Property Let lNumIntDocLoteSerieIni(ByVal vData As Long)
    mvarlNumIntDocLoteSerieIni = vData
End Property

Public Property Get lNumIntDocLoteSerieIni() As Long
    lNumIntDocLoteSerieIni = mvarlNumIntDocLoteSerieIni
End Property

Public Property Let dQuantidadeSerie(ByVal vData As Double)
    mvardQuantidadeSerie = vData
End Property

Public Property Get dQuantidadeSerie() As Double
    dQuantidadeSerie = mvardQuantidadeSerie
End Property
'#############################################

Public Property Let dQuantidadeEst(ByVal vData As Double)
    mvardQuantidadeEst = vData
End Property

Public Property Get dQuantidadeEst() As Double
    dQuantidadeEst = mvardQuantidadeEst
End Property

Public Property Let sSiglaUM(ByVal vData As String)
    mvarsSiglaUM = vData
End Property

Public Property Get sSiglaUM() As String
    sSiglaUM = mvarsSiglaUM
End Property

Public Property Let lNumIntDocLote(ByVal vData As Long)
    mvarlNumIntDocLote = vData
End Property

Public Property Get lNumIntDocLote() As Long
    lNumIntDocLote = mvarlNumIntDocLote
End Property

Public Property Let dQuantidade(ByVal vData As Double)
    mvardQuantidade = vData
End Property

Public Property Get dQuantidade() As Double
    dQuantidade = mvardQuantidade
End Property

Public Property Let iFilialOP(ByVal vData As Integer)
    mvariFilialOP = vData
End Property

Public Property Get iFilialOP() As Integer
    iFilialOP = mvariFilialOP
End Property

Public Property Let sLote(ByVal vData As String)
    mvarsLote = vData
End Property

Public Property Get sLote() As String
    sLote = mvarsLote
End Property

Public Property Let lNumIntDocOrigem(ByVal vData As Long)
    mvarlNumIntDocOrigem = vData
End Property

Public Property Get lNumIntDocOrigem() As Long
    lNumIntDocOrigem = mvarlNumIntDocOrigem
End Property

Public Property Let iTipoDocOrigem(ByVal vData As Integer)
    mvariTipoDocOrigem = vData
End Property

Public Property Get iTipoDocOrigem() As Integer
    iTipoDocOrigem = mvariTipoDocOrigem
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

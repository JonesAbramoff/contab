VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassGeracaoFatura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlCliente As Long
Private mvarsSerie As String
Private mvarlNumeroNFDe As Long
Private mvarlNumeroNFAte As Long
Private mvardtEmissaoNFDe As Date
Private mvardtEmissaoNFAte As Date
Private mvarsOrdenacao As String
Private mvarlNumTitulo As Long
Private mvardValorNFsSelecionadas As Double
Private mvardValorIRRF As Double
Private mvardValorISS As Double
Private mvariISSRetido As Integer
Private mvariFilialCli As Integer
Private mvardtDataEmissao As Date
Private mvarcolNFiscalInfo As New Collection 'colecao de objNFiscalInfo
Private mvarcolParcelas As New ColParcelaReceber
Private mvariCondicaoPagto As Integer
Private mvardValorINSS As Double

Public Property Let iCondicaoPagto(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iCondicaoPagto = 5
    mvariCondicaoPagto = vData
End Property


Public Property Get iCondicaoPagto() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iCondicaoPagto
    iCondicaoPagto = mvariCondicaoPagto
End Property

Public Property Let dValorINSS(ByVal vData As Double)
    mvardValorINSS = vData
End Property

Public Property Get dValorINSS() As Double
    dValorINSS = mvardValorINSS
End Property

Public Property Let dValorISS(ByVal vData As Double)
    mvardValorISS = vData
End Property

Public Property Get dValorISS() As Double
    dValorISS = mvardValorISS
End Property

Public Property Let iISSRetido(ByVal vData As Integer)
    mvariISSRetido = vData
End Property

Public Property Get iISSRetido() As Integer
    iISSRetido = mvariISSRetido
End Property

Public Property Let dValorIRRF(ByVal vData As Double)
    mvardValorIRRF = vData
End Property

Public Property Get dValorIRRF() As Double
    dValorIRRF = mvardValorIRRF
End Property

Public Property Let dValorNFsSelecionadas(ByVal vData As Double)
    mvardValorNFsSelecionadas = vData
End Property

Public Property Get dValorNFsSelecionadas() As Double
    dValorNFsSelecionadas = mvardValorNFsSelecionadas
End Property

Public Property Set colParcelas(ByVal vData As Object)
    Set mvarcolParcelas = vData
End Property

Public Property Get colParcelas() As ColParcelaReceber
    Set colParcelas = mvarcolParcelas
End Property

Public Property Let dtDataEmissao(ByVal vData As Date)
    mvardtDataEmissao = vData
End Property

Public Property Get dtDataEmissao() As Date
    dtDataEmissao = mvardtDataEmissao
End Property

Public Property Set colNFiscalInfo(ByVal vData As Object)
    Set mvarcolNFiscalInfo = vData
End Property

Public Property Get colNFiscalInfo() As Collection
    Set colNFiscalInfo = mvarcolNFiscalInfo
End Property

Public Property Let sOrdenacao(ByVal vData As String)
    mvarsOrdenacao = vData
End Property

Public Property Get sOrdenacao() As String
    sOrdenacao = mvarsOrdenacao
End Property

Public Property Let dtEmissaoNFAte(ByVal vData As Date)
    mvardtEmissaoNFAte = vData
End Property

Public Property Get dtEmissaoNFAte() As Date
    dtEmissaoNFAte = mvardtEmissaoNFAte
End Property

Public Property Let dtEmissaoNFDe(ByVal vData As Date)
    mvardtEmissaoNFDe = vData
End Property

Public Property Get dtEmissaoNFDe() As Date
    dtEmissaoNFDe = mvardtEmissaoNFDe
End Property

Public Property Let lNumeroNFAte(ByVal vData As Long)
    mvarlNumeroNFAte = vData
End Property

Public Property Get lNumeroNFAte() As Long
    lNumeroNFAte = mvarlNumeroNFAte
End Property

Public Property Let lNumeroNFDe(ByVal vData As Long)
    mvarlNumeroNFDe = vData
End Property

Public Property Get lNumeroNFDe() As Long
    lNumeroNFDe = mvarlNumeroNFDe
End Property

Public Property Let sSerie(ByVal vData As String)
    mvarsSerie = vData
End Property

Public Property Get sSerie() As String
    sSerie = mvarsSerie
End Property

Public Property Let lCliente(ByVal vData As Long)
    mvarlCliente = vData
End Property

Public Property Get lCliente() As Long
    lCliente = mvarlCliente
End Property

Public Property Let lNumTitulo(ByVal vData As Long)
    mvarlNumTitulo = vData
End Property

Public Property Get lNumTitulo() As Long
    lNumTitulo = mvarlNumTitulo
End Property

Public Property Let iFilialCli(ByVal vData As Integer)
    mvariFilialCli = vData
End Property

Public Property Get iFilialCli() As Integer
    iFilialCli = mvariFilialCli
End Property


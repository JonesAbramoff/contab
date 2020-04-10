VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassDetArqCNABPag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarobjLote As ClassLoteArqCNABPag

Private mvariBancoFavorecido As Integer
Private mvarsNomeFavorecido As String
Private mvarsAgenciaFavorecido As String
Private mvarsCGCFavorecido As String
Private mvardtDataVenctoParcela As Date
Private mvardValorPagto As Double
Private mvarsContaFavorecido As String
Private mvarsSiglaTitulo As String
Private mvarlNumTitulo As Long
Private mvardtDataEmissaoTitulo As Date
Private mvarsNossoNumero As String
Private mvarsCodigoDeBarras As String 'Guarda o Numero Referênte ao Código da Barras

Private mvarsEnderecoFavorecido As String
Private mvarsBairroFavorecido As String
Private mvarsCidadeFavorecido As String
Private mvarsCEPFavorecido As String
Private mvarsEstadoFavorecido As String

Private mvarlFornecedor As Long
Private mvariFilialForn As Integer
Private mvariNumParcela As Integer
Private mvariSeqBaixaParcPag As Integer
Private mvardValorMulta As Double
Private mvardValorJuros As Double
Private mvardValorDesconto As Double
Private mvardValorTitulo As Double

Private mvariTipoCobranca As Integer
Private mvariBancoCobrador As Integer
Private mvarsFormaPagamento As String
Private mvarlNumIntParcela As Long

Public Property Let lNumIntParcela(ByVal vData As Long)
    mvarlNumIntParcela = vData
End Property

Public Property Get lNumIntParcela() As Long
    lNumIntParcela = mvarlNumIntParcela
End Property

Public Property Let sFormaPagamento(ByVal vData As String)
    mvarsFormaPagamento = vData
End Property

Public Property Get sFormaPagamento() As String
    sFormaPagamento = mvarsFormaPagamento
End Property

Public Property Let sEnderecoFavorecido(ByVal vData As String)
    mvarsEnderecoFavorecido = vData
End Property

Public Property Get sEnderecoFavorecido() As String
    sEnderecoFavorecido = mvarsEnderecoFavorecido
End Property

Public Property Let sBairroFavorecido(ByVal vData As String)
    mvarsBairroFavorecido = vData
End Property

Public Property Get sBairroFavorecido() As String
    sBairroFavorecido = mvarsBairroFavorecido
End Property

Public Property Let sCidadeFavorecido(ByVal vData As String)
    mvarsCidadeFavorecido = vData
End Property

Public Property Get sCidadeFavorecido() As String
    sCidadeFavorecido = mvarsCidadeFavorecido
End Property

Public Property Let sCEPFavorecido(ByVal vData As String)
    mvarsCEPFavorecido = vData
End Property

Public Property Get sCEPFavorecido() As String
    sCEPFavorecido = mvarsCEPFavorecido
End Property

Public Property Let sEstadoFavorecido(ByVal vData As String)
    mvarsEstadoFavorecido = vData
End Property

Public Property Get sEstadoFavorecido() As String
    sEstadoFavorecido = mvarsEstadoFavorecido
End Property

Public Property Let sCodigoDeBarras(ByVal vData As String)
    mvarsCodigoDeBarras = vData
End Property

Public Property Get sCodigoDeBarras() As String
    sCodigoDeBarras = mvarsCodigoDeBarras
End Property

Public Property Let iBancoCobrador(ByVal vData As Integer)
    mvariBancoCobrador = vData
End Property

Public Property Get iBancoCobrador() As Integer
    iBancoCobrador = mvariBancoCobrador
End Property

Public Property Let iTipoCobranca(ByVal vData As Integer)
    mvariTipoCobranca = vData
End Property

Public Property Get iTipoCobranca() As Integer
    iTipoCobranca = mvariTipoCobranca
End Property

Public Property Let sNossoNumero(ByVal vData As String)
    mvarsNossoNumero = vData
End Property

Public Property Get sNossoNumero() As String
    sNossoNumero = mvarsNossoNumero
End Property

Public Property Let dtDataEmissaoTitulo(ByVal vData As Date)
    mvardtDataEmissaoTitulo = vData
End Property

Public Property Get dtDataEmissaoTitulo() As Date
    dtDataEmissaoTitulo = mvardtDataEmissaoTitulo
End Property

Public Property Let lNumTitulo(ByVal vData As Long)
    mvarlNumTitulo = vData
End Property

Public Property Get lNumTitulo() As Long
    lNumTitulo = mvarlNumTitulo
End Property

Public Property Let sSiglaTitulo(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sSiglaTitulo = 5
    mvarsSiglaTitulo = vData
End Property

Public Property Get sSiglaTitulo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sSiglaTitulo
    sSiglaTitulo = mvarsSiglaTitulo
End Property

Public Property Let sContaFavorecido(ByVal vData As String)
    mvarsContaFavorecido = vData
End Property

Public Property Get sContaFavorecido() As String
    sContaFavorecido = mvarsContaFavorecido
End Property

Public Property Let dValorPagto(ByVal vData As Double)
    mvardValorPagto = vData
End Property

Public Property Get dValorPagto() As Double
    dValorPagto = mvardValorPagto
End Property

Public Property Let dtDataVenctoParcela(ByVal vData As Date)
    mvardtDataVenctoParcela = vData
End Property

Public Property Get dtDataVenctoParcela() As Date
    dtDataVenctoParcela = mvardtDataVenctoParcela
End Property

Public Property Let sCGCFavorecido(ByVal vData As String)
    mvarsCGCFavorecido = vData
End Property

Public Property Get sCGCFavorecido() As String
    sCGCFavorecido = mvarsCGCFavorecido
End Property

Public Property Let sAgenciaFavorecido(ByVal vData As String)
    mvarsAgenciaFavorecido = vData
End Property

Public Property Get sAgenciaFavorecido() As String
    sAgenciaFavorecido = mvarsAgenciaFavorecido
End Property

Public Property Let sNomeFavorecido(ByVal vData As String)
    mvarsNomeFavorecido = vData
End Property

Public Property Get sNomeFavorecido() As String
    sNomeFavorecido = mvarsNomeFavorecido
End Property

Public Property Let iBancoFavorecido(ByVal vData As Integer)
    mvariBancoFavorecido = vData
End Property

Public Property Get iBancoFavorecido() As Integer
    iBancoFavorecido = mvariBancoFavorecido
End Property

Public Property Get objLote() As ClassLoteArqCNABPag
    Set objLote = mvarobjLote
End Property

Public Function Iniciar(ByVal objLotePar As ClassLoteArqCNABPag) As Long

    Set mvarobjLote = objLotePar
    Call objLotePar.AcumulaValor(mvardValorPagto)
    
End Function

Public Property Get lSequencial() As Long

    lSequencial = objLote.lSequencialDet
    
End Property

Public Property Let dValorTitulo(ByVal vData As Double)
    mvardValorTitulo = vData
End Property

Public Property Get dValorTitulo() As Double
    dValorTitulo = mvardValorTitulo
End Property

Public Property Let dValorDesconto(ByVal vData As Double)
    mvardValorDesconto = vData
End Property

Public Property Get dValorDesconto() As Double
    dValorDesconto = mvardValorDesconto
End Property

Public Property Let dValorJuros(ByVal vData As Double)
    mvardValorJuros = vData
End Property

Public Property Get dValorJuros() As Double
    dValorJuros = mvardValorJuros
End Property

Public Property Let dValorMulta(ByVal vData As Double)
    mvardValorMulta = vData
End Property

Public Property Get dValorMulta() As Double
    dValorMulta = mvardValorMulta
End Property

Public Property Let iSeqBaixaParcPag(ByVal vData As Integer)
    mvariSeqBaixaParcPag = vData
End Property

Public Property Get iSeqBaixaParcPag() As Integer
    iSeqBaixaParcPag = mvariSeqBaixaParcPag
End Property

Public Property Let iNumParcela(ByVal vData As Integer)
    mvariNumParcela = vData
End Property

Public Property Get iNumParcela() As Integer
    iNumParcela = mvariNumParcela
End Property

Public Property Let lFornecedor(ByVal vData As Long)
    mvarlFornecedor = vData
End Property

Public Property Get lFornecedor() As Long
    lFornecedor = mvarlFornecedor
End Property

Public Property Let iFilialForn(ByVal vData As Integer)
    mvariFilialForn = vData
End Property

Public Property Get iFilialForn() As Integer
    iFilialForn = mvariFilialForn
End Property


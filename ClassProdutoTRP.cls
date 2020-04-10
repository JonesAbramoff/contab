VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassProdutoTRP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvariVendedor As Integer
Private mvardPercComissao As Double
Private mvardDescontoTitular As Double
Private mvardDescontoAcompanhante As Double
Private mvariLimiteIdade As Integer
Private mvariMaxAcompanhante As Integer
Private mvariMoeda As Integer
Private mvariExibeWeb As Integer
Private mvarcolTarifas As New Collection
Private mvarsProduto As String
Private mvariAceitaIdadeMaiorLimite As Integer
Private mvardPercAcrescIdadeMaiorLimite As Double
Private mvariVigenciaFixa As Integer
Private mvariTodosDestValidos As Integer
Private mvarcolDestinos As New Collection
Private mvarsTextoVouEN As String
Private mvarsTextoVouPT As String

'AUXILIAR
Private mvardtDataVigencia As Date

Public Property Let dtDataVigencia(ByVal vData As Date)
    mvardtDataVigencia = vData
End Property

Public Property Get dtDataVigencia() As Date
    dtDataVigencia = mvardtDataVigencia
End Property
'FIM AUXILIAR

Public Property Let sTextoVouPT(ByVal vData As String)
    mvarsTextoVouPT = vData
End Property

Public Property Get sTextoVouPT() As String
    sTextoVouPT = mvarsTextoVouPT
End Property

Public Property Let sTextoVouEN(ByVal vData As String)
    mvarsTextoVouEN = vData
End Property

Public Property Get sTextoVouEN() As String
    sTextoVouEN = mvarsTextoVouEN
End Property

Public Property Set colDestinos(ByVal vData As Collection)
    Set mvarcolDestinos = vData
End Property

Public Property Get colDestinos() As Collection
    Set colDestinos = mvarcolDestinos
End Property

Public Property Let iTodosDestValidos(ByVal vData As Integer)
    mvariTodosDestValidos = vData
End Property

Public Property Get iTodosDestValidos() As Integer
    iTodosDestValidos = mvariTodosDestValidos
End Property

Public Property Let iVigenciaFixa(ByVal vData As Integer)
    mvariVigenciaFixa = vData
End Property

Public Property Get iVigenciaFixa() As Integer
    iVigenciaFixa = mvariVigenciaFixa
End Property

Public Property Let dPercAcrescIdadeMaiorLimite(ByVal vData As Double)
    mvardPercAcrescIdadeMaiorLimite = vData
End Property

Public Property Get dPercAcrescIdadeMaiorLimite() As Double
    dPercAcrescIdadeMaiorLimite = mvardPercAcrescIdadeMaiorLimite
End Property

Public Property Let iAceitaIdadeMaiorLimite(ByVal vData As Integer)
    mvariAceitaIdadeMaiorLimite = vData
End Property

Public Property Get iAceitaIdadeMaiorLimite() As Integer
    iAceitaIdadeMaiorLimite = mvariAceitaIdadeMaiorLimite
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Set colTarifas(ByVal vData As Collection)
    Set mvarcolTarifas = vData
End Property

Public Property Get colTarifas() As Collection
    Set colTarifas = mvarcolTarifas
End Property

Public Property Let iExibeWeb(ByVal vData As Integer)
    mvariExibeWeb = vData
End Property

Public Property Get iExibeWeb() As Integer
    iExibeWeb = mvariExibeWeb
End Property

Public Property Let iMoeda(ByVal vData As Integer)
    mvariMoeda = vData
End Property

Public Property Get iMoeda() As Integer
    iMoeda = mvariMoeda
End Property

Public Property Let iMaxAcompanhante(ByVal vData As Integer)
    mvariMaxAcompanhante = vData
End Property

Public Property Get iMaxAcompanhante() As Integer
    iMaxAcompanhante = mvariMaxAcompanhante
End Property

Public Property Let iLimiteIdade(ByVal vData As Integer)
    mvariLimiteIdade = vData
End Property

Public Property Get iLimiteIdade() As Integer
    iLimiteIdade = mvariLimiteIdade
End Property

Public Property Let dDescontoAcompanhante(ByVal vData As Double)
    mvardDescontoAcompanhante = vData
End Property

Public Property Get dDescontoAcompanhante() As Double
    dDescontoAcompanhante = mvardDescontoAcompanhante
End Property

Public Property Let dDescontoTitular(ByVal vData As Double)
    mvardDescontoTitular = vData
End Property

Public Property Get dDescontoTitular() As Double
    dDescontoTitular = mvardDescontoTitular
End Property

Public Property Let dPercComissao(ByVal vData As Double)
    mvardPercComissao = vData
End Property

Public Property Get dPercComissao() As Double
    dPercComissao = mvardPercComissao
End Property

Public Property Let iVendedor(ByVal vData As Integer)
    mvariVendedor = vData
End Property

Public Property Get iVendedor() As Integer
    iVendedor = mvariVendedor
End Property




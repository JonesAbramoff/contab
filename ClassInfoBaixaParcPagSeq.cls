VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassInfoBaixaParcPagSeq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlCodigo As Long
Private mvarsNomeReduzido As String
Private mvariCodFilial As Integer
Private mvarsNome As String
Private mvarsSiglaDocumento As String
Private mvarlNumTitulo As Long
Private mvardtDataEmissao As Date
Private mvariNumParcelas As Integer
Private mvardtDataVencimento As Date
Private mvardSaldo As Double
Private mvardValor As Double
Private mvariSequencial As Integer
Private mvardValorMulta As Double
Private mvardValorJuros As Double
Private mvardValorDesconto As Double
Private mvardValorBaixado As Double
Private mvardtData As Date

Public Property Let dtData(ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let dValorBaixado(ByVal vData As Double)
    mvardValorBaixado = vData
End Property

Public Property Get dValorBaixado() As Double
    dValorBaixado = mvardValorBaixado
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

Public Property Let iSequencial(ByVal vData As Integer)
    mvariSequencial = vData
End Property

Public Property Get iSequencial() As Integer
    iSequencial = mvariSequencial
End Property

Public Property Let dValor(ByVal vData As Double)
    mvardValor = vData
End Property

Public Property Get dValor() As Double
    dValor = mvardValor
End Property

Public Property Let dSaldo(ByVal vData As Double)
    mvardSaldo = vData
End Property

Public Property Get dSaldo() As Double
    dSaldo = mvardSaldo
End Property

Public Property Let dtDataVencimento(ByVal vData As Date)
    mvardtDataVencimento = vData
End Property

Public Property Get dtDataVencimento() As Date
    dtDataVencimento = mvardtDataVencimento
End Property

Public Property Let iNumParcelas(ByVal vData As Integer)
    mvariNumParcelas = vData
End Property

Public Property Get iNumParcelas() As Integer
    iNumParcelas = mvariNumParcelas
End Property

Public Property Let dtDataEmissao(ByVal vData As Date)
    mvardtDataEmissao = vData
End Property

Public Property Get dtDataEmissao() As Date

'Syntax: Debug.Print X.dtDataEmissao
    dtDataEmissao = mvardtDataEmissao
End Property



Public Property Let lNumTitulo(ByVal vData As Long)

'Syntax: X.lNumTitulo = 5
    mvarlNumTitulo = vData
End Property


Public Property Get lNumTitulo() As Long

'Syntax: Debug.Print X.lNumTitulo
    lNumTitulo = mvarlNumTitulo
End Property



Public Property Let sSiglaDocumento(ByVal vData As String)

'Syntax: X.sSiglaDocumento = 5
    mvarsSiglaDocumento = vData
End Property


Public Property Get sSiglaDocumento() As String

'Syntax: Debug.Print X.sSiglaDocumento
    sSiglaDocumento = mvarsSiglaDocumento
End Property



Public Property Let sNome(ByVal vData As String)

'Syntax: X.sNome = 5
    mvarsNome = vData
End Property


Public Property Get sNome() As String

'Syntax: Debug.Print X.sNome
    sNome = mvarsNome
End Property



Public Property Let iCodFilial(ByVal vData As Integer)

'Syntax: X.iCodFilial = 5
    mvariCodFilial = vData
End Property


Public Property Get iCodFilial() As Integer

'Syntax: Debug.Print X.iCodFilial
    iCodFilial = mvariCodFilial
End Property



Public Property Let sNomeReduzido(ByVal vData As String)

'Syntax: X.sNomeReduzido = 5
    mvarsNomeReduzido = vData
End Property


Public Property Get sNomeReduzido() As String

'Syntax: Debug.Print X.sNomeReduzido
    sNomeReduzido = mvarsNomeReduzido
End Property



Public Property Let lCodigo(ByVal vData As Long)

'Syntax: X.lCodigo = 5
    mvarlCodigo = vData
End Property


Public Property Get lCodigo() As Long

'Syntax: Debug.Print X.lCodigo
    lCodigo = mvarlCodigo
End Property




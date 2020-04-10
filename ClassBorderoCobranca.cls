VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassBorderoCobranca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

Private mvariCobrador As Integer
Private mvarlNumBordero As Long
Private mvardtDataEmissao As Date
Private mvariCodNossaConta As Integer
Private mvariDiasDeRetencao As Integer
Private mvardTaxaCobranca As Double
Private mvardTaxaDesconto As Double
Private mvardValor As Double
Private mvardValorDesconto As Double
Private mvarsConvenio As String
Private mvariStatus As Integer 'local copy
Private mvardtDataCancelamento As Date 'local copy
Private mvardtDataContabilCancelamento As Date 'local copy
Private mvariCodCarteiraCobranca As Integer 'local copy
Private mvariProcessado As Integer 'local copy
'local variable(s) to hold property value(s)
Private mvarsNomeArquivo As String 'local copy
Public Property Let sNomeArquivo(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sNomeArquivo = 5
    mvarsNomeArquivo = vData
End Property


Public Property Get sNomeArquivo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sNomeArquivo
    sNomeArquivo = mvarsNomeArquivo
End Property



Public Property Let iProcessado(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iProcessado = 5
    mvariProcessado = vData
End Property


Public Property Get iProcessado() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iProcessado
    iProcessado = mvariProcessado
End Property



Public Property Let iCodCarteiraCobranca(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iCodCarteiraCobranca = 5
    mvariCodCarteiraCobranca = vData
End Property


Public Property Get iCodCarteiraCobranca() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iCodCarteiraCobranca
    iCodCarteiraCobranca = mvariCodCarteiraCobranca
End Property



Public Property Let dtDataContabilCancelamento(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtDataContabilCancelamento = 5
    mvardtDataContabilCancelamento = vData
End Property


Public Property Get dtDataContabilCancelamento() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtDataContabilCancelamento
    dtDataContabilCancelamento = mvardtDataContabilCancelamento
End Property



Public Property Let dtDataCancelamento(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtDataCancelamento = 5
    mvardtDataCancelamento = vData
End Property


Public Property Get dtDataCancelamento() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtDataCancelamento
    dtDataCancelamento = mvardtDataCancelamento
End Property



Public Property Let iStatus(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iStatus = 5
    mvariStatus = vData
End Property


Public Property Get iStatus() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iStatus
    iStatus = mvariStatus
End Property




Public Property Let sConvenio(ByVal vData As String)
Attribute sConvenio.VB_Description = "qual o convenio utilizado"
    mvarsConvenio = vData
End Property

Public Property Get sConvenio() As String
    sConvenio = mvarsConvenio
End Property

Public Property Let dValorDesconto(ByVal vData As Double)
    mvardValorDesconto = vData
End Property

Public Property Get dValorDesconto() As Double
    dValorDesconto = mvardValorDesconto
End Property

Public Property Let dValor(ByVal vData As Double)
    mvardValor = vData
End Property

Public Property Get dValor() As Double
    dValor = mvardValor
End Property

Public Property Let dTaxaDesconto(ByVal vData As Double)
    mvardTaxaDesconto = vData
End Property

Public Property Get dTaxaDesconto() As Double
    dTaxaDesconto = mvardTaxaDesconto
End Property

Public Property Let dTaxaCobranca(ByVal vData As Double)
    mvardTaxaCobranca = vData
End Property

Public Property Get dTaxaCobranca() As Double
    dTaxaCobranca = mvardTaxaCobranca
End Property

Public Property Let iDiasDeRetencao(ByVal vData As Integer)
    mvariDiasDeRetencao = vData
End Property

Public Property Get iDiasDeRetencao() As Integer
    iDiasDeRetencao = mvariDiasDeRetencao
End Property

Public Property Let iCodNossaConta(ByVal vData As Integer)
Attribute iCodNossaConta.VB_Description = "cta onde será efetuado o deposito"
    mvariCodNossaConta = vData
End Property

Public Property Get iCodNossaConta() As Integer
    iCodNossaConta = mvariCodNossaConta
End Property

Public Property Let dtDataEmissao(ByVal vData As Date)
    mvardtDataEmissao = vData
End Property

Public Property Get dtDataEmissao() As Date
    dtDataEmissao = mvardtDataEmissao
End Property

Public Property Let lNumBordero(ByVal vData As Long)
    mvarlNumBordero = vData
End Property

Public Property Get lNumBordero() As Long
    lNumBordero = mvarlNumBordero
End Property

Public Property Let iCobrador(ByVal vData As Integer)
    mvariCobrador = vData
End Property

Public Property Get iCobrador() As Integer
    iCobrador = mvariCobrador
End Property




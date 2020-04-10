VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassNFProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarsProduto As String 'local copy
Private mvarsDescricao As String 'local copy
Private mvardQuatidade As Double 'local copy
Private mvardPrecoUnitario As Double 'local copy
Private mvardPrecoTotal As Double 'local copy
Private mvardValorIPI As Double 'local copy
Private mvardValorISS As Double 'local copy
Private mvarlNumIntDoc As Long 'local copy
Private mvariLinha As Integer 'local copy
Private mvardIPIAliquota As Double 'local copy
Private mvardICMSAliquota As Double 'local copy
Private mvariTipoProduto As Integer 'local copy
Private mvarsClassificacaoFiscal As String 'local copy
Private mvarsSituacaoTributaria As String 'local copy
'local variable(s) to hold property value(s)
'local variable(s) to hold property value(s)
Private mvariOrigemMercadoria As Integer 'local copy
'local variable(s) to hold property value(s)
Private mvarsUnidadeMed As String 'local copy
Public Property Let sUnidadeMed(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sUnidadeMed = 5
    mvarsUnidadeMed = vData
End Property


Public Property Get sUnidadeMed() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sUnidadeMed
    sUnidadeMed = mvarsUnidadeMed
End Property



Public Property Let iOrigemMercadoria(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iOrigemMercadoria = 5
    mvariOrigemMercadoria = vData
End Property


Public Property Get iOrigemMercadoria() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iOrigemMercadoria
    iOrigemMercadoria = mvariOrigemMercadoria
End Property




Public Property Let sSituacaoTributaria(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sIPICodigo = 5
    mvarsSituacaoTributaria = vData
End Property


Public Property Get sSituacaoTributaria() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sIPICodigo
    sSituacaoTributaria = mvarsSituacaoTributaria
End Property

Public Property Let sClassificacaoFiscal(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sIPICodigo = 5
    mvarsClassificacaoFiscal = vData
End Property


Public Property Get sClassificacaoFiscal() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sIPICodigo
    sClassificacaoFiscal = mvarsClassificacaoFiscal
End Property




Public Property Let iTipoProduto(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dICMSAliquota = 5
    mvariTipoProduto = vData
End Property


Public Property Get iTipoProduto() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dICMSAliquota
    iTipoProduto = mvariTipoProduto
End Property

Public Property Let dICMSAliquota(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dICMSAliquota = 5
    mvardICMSAliquota = vData
End Property


Public Property Get dICMSAliquota() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dICMSAliquota
    dICMSAliquota = mvardICMSAliquota
End Property



Public Property Let dIPIAliquota(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dIPIAliquota = 5
    mvardIPIAliquota = vData
End Property


Public Property Get dIPIAliquota() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dIPIAliquota
    dIPIAliquota = mvardIPIAliquota
End Property



Public Property Let iLinha(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iLinha = 5
    mvariLinha = vData
End Property


Public Property Get iLinha() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iLinha
    iLinha = mvariLinha
End Property



Public Property Let lNumIntDoc(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lNumIntDoc = 5
    mvarlNumIntDoc = vData
End Property


Public Property Get lNumIntDoc() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lNumIntDoc
    lNumIntDoc = mvarlNumIntDoc
End Property



Public Property Let dValorISS(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dValorISS = 5
    mvardValorISS = vData
End Property


Public Property Get dValorISS() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dValorISS
    dValorISS = mvardValorISS
End Property



Public Property Let dValorIPI(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dValorIPI = 5
    mvardValorIPI = vData
End Property


Public Property Get dValorIPI() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dValorIPI
    dValorIPI = mvardValorIPI
End Property



Public Property Let dPrecoTotal(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dPrecoTotal = 5
    mvardPrecoTotal = vData
End Property


Public Property Get dPrecoTotal() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dPrecoTotal
    dPrecoTotal = mvardPrecoTotal
End Property



Public Property Let dPrecoUnitario(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dPrecoUnitario = 5
    mvardPrecoUnitario = vData
End Property


Public Property Get dPrecoUnitario() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dPrecoUnitario
    dPrecoUnitario = mvardPrecoUnitario
End Property



Public Property Let dQuatidade(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dQuatidade = 5
    mvardQuatidade = vData
End Property


Public Property Get dQuatidade() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dQuatidade
    dQuatidade = mvardQuatidade
End Property



Public Property Let sDescricao(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sDescricao = 5
    mvarsDescricao = vData
End Property


Public Property Get sDescricao() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sDescricao
    sDescricao = mvarsDescricao
End Property



Public Property Let sProduto(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sProduto = 5
    mvarsProduto = vData
End Property


Public Property Get sProduto() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sProduto
    sProduto = mvarsProduto
End Property




VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassIN86CTBClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarsContaContabil As String 'local copy
Private mvarlCliente As Long 'local copy
Private mvariFilial As Integer 'local copy
Private mvardtDataOperacao As Date 'local copy
Private mvarsHistorico As String 'local copy
Private mvardValorOperacao As Double 'local copy
Private mvarsTipoOperacao As String 'local copy
Private mvarsTipoDocumento As String 'local copy
Private mvarlDocumento As Long 'local copy
Private mvardValorTitulo As Double 'local copy
Private mvardtDataEmissao As Date 'local copy
Private mvardtDataVencimento As Date 'local copy
Private mvarsArquivamento As String 'local copy
'local variable(s) to hold property value(s)
Private mvariNumParcela As Integer 'local copy
'local variable(s) to hold property value(s)
Private mvariFilialEmpresa As Integer 'local copy
Public Property Let iFilialEmpresa(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iFilialEmpresa = 5
    mvariFilialEmpresa = vData
End Property


Public Property Get iFilialEmpresa() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iFilialEmpresa
    iFilialEmpresa = mvariFilialEmpresa
End Property



Public Property Let iNumParcela(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iNumParcela = 5
    mvariNumParcela = vData
End Property


Public Property Get iNumParcela() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iNumParcela
    iNumParcela = mvariNumParcela
End Property




Public Property Let sArquivamento(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sArquivamento = 5
    mvarsArquivamento = vData
End Property


Public Property Get sArquivamento() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sArquivamento
    sArquivamento = mvarsArquivamento
End Property



Public Property Let dtDataVencimento(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtDataVencimento = 5
    mvardtDataVencimento = vData
End Property


Public Property Get dtDataVencimento() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtDataVencimento
    dtDataVencimento = mvardtDataVencimento
End Property



Public Property Let dtDataEmissao(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtDataEmissao = 5
    mvardtDataEmissao = vData
End Property


Public Property Get dtDataEmissao() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtDataEmissao
    dtDataEmissao = mvardtDataEmissao
End Property



Public Property Let dValorTitulo(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dValorTitulo = 5
    mvardValorTitulo = vData
End Property


Public Property Get dValorTitulo() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dValorTitulo
    dValorTitulo = mvardValorTitulo
End Property



Public Property Let lDocumento(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lDocumento = 5
    mvarlDocumento = vData
End Property


Public Property Get lDocumento() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lDocumento
    lDocumento = mvarlDocumento
End Property



Public Property Let sTipoDocumento(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sTipoDocumento = 5
    mvarsTipoDocumento = vData
End Property


Public Property Get sTipoDocumento() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sTipoDocumento
    sTipoDocumento = mvarsTipoDocumento
End Property



Public Property Let sTipoOperacao(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sTipoOperacao = 5
    mvarsTipoOperacao = vData
End Property


Public Property Get sTipoOperacao() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sTipoOperacao
    sTipoOperacao = mvarsTipoOperacao
End Property



Public Property Let dValorOperacao(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dValorOperacao = 5
    mvardValorOperacao = vData
End Property


Public Property Get dValorOperacao() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dValorOperacao
    dValorOperacao = mvardValorOperacao
End Property



Public Property Let sHistorico(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sHistorico = 5
    mvarsHistorico = vData
End Property


Public Property Get sHistorico() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sHistorico
    sHistorico = mvarsHistorico
End Property



Public Property Let dtDataOperacao(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtDataOperacao = 5
    mvardtDataOperacao = vData
End Property


Public Property Get dtDataOperacao() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtDataOperacao
    dtDataOperacao = mvardtDataOperacao
End Property



Public Property Let iFilial(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iFilial = 5
    mvariFilial = vData
End Property


Public Property Get iFilial() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iFilial
    iFilial = mvariFilial
End Property



Public Property Let lCliente(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lCliente = 5
    mvarlCliente = vData
End Property


Public Property Get lCliente() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lCliente
    lCliente = mvarlCliente
End Property



Public Property Let sContaContabil(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sContaContabil = 5
    mvarsContaContabil = vData
End Property


Public Property Get sContaContabil() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sContaContabil
    sContaContabil = mvarsContaContabil
End Property




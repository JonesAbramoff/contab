VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCusteioRoteiro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarlNumIntDoc As Long 'local copy
Private mvarlCodigo As Long 'local copy
Private mvarsNomeReduzido As String 'local copy
Private mvarsDescricao As String 'local copy
Private mvarsProduto As String 'local copy
Private mvarsVersao As String 'local copy
Private mvarsUMedida As String 'local copy
Private mvardQuantidade As Double 'local copy
Private mvardtDataCusteio As Date 'local copy
Private mvardtDataValidade As Date 'local copy
Private mvardCustoTotalInsumosKit As Double 'local copy
Private mvardCustoTotalInsumosMaq As Double 'local copy
Private mvardCustoTotalMaoDeObra As Double 'local copy
Private mvardPrecoTotalRoteiro As Double 'local copy
Private mvarsObservacao As String 'local copy
Private mvarcolCusteioRotInsumosKit As New Collection
Private mvarcolCusteioRotInsumosMaq As New Collection
Private mvarcolCusteioRotMaoDeObra As New Collection

Public Property Set colCusteioRotInsumosKit(ByVal vData As Collection)
    Set mvarcolCusteioRotInsumosKit = vData
End Property

Public Property Get colCusteioRotInsumosKit() As Collection
    Set colCusteioRotInsumosKit = mvarcolCusteioRotInsumosKit
End Property

Public Property Set colCusteioRotInsumosMaq(ByVal vData As Collection)
    Set mvarcolCusteioRotInsumosMaq = vData
End Property

Public Property Get colCusteioRotInsumosMaq() As Collection
    Set colCusteioRotInsumosMaq = mvarcolCusteioRotInsumosMaq
End Property

Public Property Set colCusteioRotMaoDeObra(ByVal vData As Collection)
    Set mvarcolCusteioRotMaoDeObra = vData
End Property

Public Property Get colCusteioRotMaoDeObra() As Collection
    Set colCusteioRotMaoDeObra = mvarcolCusteioRotMaoDeObra
End Property

Public Property Let sObservacao(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sObservacao = 5
    mvarsObservacao = vData
End Property

Public Property Get sObservacao() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sObservacao
    sObservacao = mvarsObservacao
End Property

Public Property Let dPrecoTotalRoteiro(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dPrecoTotalRoteiro = 5
    mvardPrecoTotalRoteiro = vData
End Property

Public Property Get dPrecoTotalRoteiro() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dPrecoTotalRoteiro
    dPrecoTotalRoteiro = mvardPrecoTotalRoteiro
End Property

Public Property Let dCustoTotalMaoDeObra(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dCustoTotalMaoDeObra = 5
    mvardCustoTotalMaoDeObra = vData
End Property

Public Property Get dCustoTotalMaoDeObra() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dCustoTotalMaoDeObra
    dCustoTotalMaoDeObra = mvardCustoTotalMaoDeObra
End Property

Public Property Let dCustoTotalInsumosMaq(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dCustoTotalInsumosMaq = 5
    mvardCustoTotalInsumosMaq = vData
End Property

Public Property Get dCustoTotalInsumosMaq() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dCustoTotalInsumosMaq
    dCustoTotalInsumosMaq = mvardCustoTotalInsumosMaq
End Property

Public Property Let dCustoTotalInsumosKit(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dCustoTotalInsumosKit = 5
    mvardCustoTotalInsumosKit = vData
End Property

Public Property Get dCustoTotalInsumosKit() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dCustoTotalInsumosKit
    dCustoTotalInsumosKit = mvardCustoTotalInsumosKit
End Property

Public Property Let dtDataValidade(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtDataValidade = 5
    mvardtDataValidade = vData
End Property

Public Property Get dtDataValidade() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtDataValidade
    dtDataValidade = mvardtDataValidade
End Property

Public Property Let dtDataCusteio(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtDataCusteio = 5
    mvardtDataCusteio = vData
End Property

Public Property Get dtDataCusteio() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtDataCusteio
    dtDataCusteio = mvardtDataCusteio
End Property

Public Property Let dQuantidade(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dQuantidade = 5
    mvardQuantidade = vData
End Property

Public Property Get dQuantidade() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dQuantidade
    dQuantidade = mvardQuantidade
End Property

Public Property Let sUMedida(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sUMedida = 5
    mvarsUMedida = vData
End Property

Public Property Get sUMedida() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sUMedida
    sUMedida = mvarsUMedida
End Property

Public Property Let sVersao(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sVersao = 5
    mvarsVersao = vData
End Property

Public Property Get sVersao() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sVersao
    sVersao = mvarsVersao
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

Public Property Let sNomeReduzido(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sNomeReduzido = 5
    mvarsNomeReduzido = vData
End Property

Public Property Get sNomeReduzido() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sNomeReduzido
    sNomeReduzido = mvarsNomeReduzido
End Property

Public Property Let lCodigo(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lCodigo = 5
    mvarlCodigo = vData
End Property

Public Property Get lCodigo() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lCodigo
    lCodigo = mvarlCodigo
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


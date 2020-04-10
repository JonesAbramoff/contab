VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassItemConcorrencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarlNumIntDoc As Long 'local copy
Private mvarsProduto As String 'local copy
Private mvarlFornecedor As Long 'local copy
Private mvariFilial As Integer 'local copy
Private mvardQuantidade As Double 'local copy
Private mvarsUM As String 'local copy
Private mvarcolQuantSuplementar As New Collection 'local copy
Private mvarcolCotacaoItemConc As New Collection 'local copy
Private mvarsDescricao As String 'local copy
Private mvardtDataNecessidade As Date 'local copy
Private mvarcolItemRCItemConcorrencia As New Collection 'local copy
Private mvariEscolhido As Integer
Private mvardQuantUrgente As Double
Public Property Let dQuantUrgente(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dQuantUrgente = 5
    mvardQuantUrgente = vData
End Property


Public Property Get dQuantUrgente() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dQuantUrgente
    dQuantUrgente = mvardQuantUrgente
End Property



Public Property Let iEscolhido(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iEscolhido = 5
    mvariEscolhido = vData
End Property


Public Property Get iEscolhido() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iEscolhido
    iEscolhido = mvariEscolhido
End Property



Public Property Set colItemRCItemConcorrencia(ByVal vData As Object)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.colItemRCItemConcorrencia = Form1
    Set mvarcolItemRCItemConcorrencia = vData
End Property


Public Property Get colItemRCItemConcorrencia() As Collection
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.colItemRCItemConcorrencia
    Set colItemRCItemConcorrencia = mvarcolItemRCItemConcorrencia
End Property



Public Property Let dtDataNecessidade(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtDataNecessidade = 5
    mvardtDataNecessidade = vData
End Property


Public Property Get dtDataNecessidade() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtDataNecessidade
    dtDataNecessidade = mvardtDataNecessidade
End Property



Public Property Let sDescricao(ByVal vData As String)
Attribute sDescricao.VB_Description = "Descrição do Produto."
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sDescricao = 5
    mvarsDescricao = vData
End Property


Public Property Get sDescricao() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sDescricao
    sDescricao = mvarsDescricao
End Property



Public Property Set colCotacaoItemConc(ByVal vData As Object)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.colCotacaoItemConc = Form1
    Set mvarcolCotacaoItemConc = vData
End Property


Public Property Get colCotacaoItemConc() As Collection
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.colCotacaoItemConc
    Set colCotacaoItemConc = mvarcolCotacaoItemConc
End Property



Public Property Set colQuantSuplementar(ByVal vData As Object)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.colQuantSuplementar = Form1
    Set mvarcolQuantSuplementar = vData
End Property


Public Property Get colQuantSuplementar() As Collection
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.colQuantSuplementar
    Set colQuantSuplementar = mvarcolQuantSuplementar
End Property



Public Property Let sUM(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sUM = 5
    mvarsUM = vData
End Property


Public Property Get sUM() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sUM
    sUM = mvarsUM
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



Public Property Let lFornecedor(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lFornecedor = 5
    mvarlFornecedor = vData
End Property


Public Property Get lFornecedor() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lFornecedor
    lFornecedor = mvarlFornecedor
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




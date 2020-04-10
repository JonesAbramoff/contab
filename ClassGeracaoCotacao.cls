VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassGeracaoCotacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvardtDataDe As Date 'local copy
Private mvardtDataAte As Date 'local copy
Private mvariTipoDestino As Integer 'local copy
Private mvarlFornCliDestino As Long 'local copy
Private mvariFilialDestino As Integer 'local copy
Private mvarlCodigoDe As Long 'local copy
Private mvarlCodigoAte As Long 'local copy
Private mvariSelecionaDestino As Integer 'local copy
Private mvarsOrdenacaoReq As String 'local copy
Private mvarcolReqCompra As New Collection 'local copy
Private mvardtDataLimiteDe As Date 'local copy
Private mvardtDataLimiteAte As Date 'local copy
Private mvariExibeReqCotadas As Integer 'local copy
Private mvarcolTipoProduto As New Collection 'local copy
Private mvarsOrdenacaoItensReq As String 'local copy
Public Property Let sOrdenacaoItensReq(ByVal vData As String)
Attribute sOrdenacaoItensReq.VB_Description = "Ordenação de cotações."
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sOrdenacaoItensReq = 5
    mvarsOrdenacaoItensReq = vData
End Property


Public Property Get sOrdenacaoItensReq() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sOrdenacaoItensReq
    sOrdenacaoItensReq = mvarsOrdenacaoItensReq
End Property



Public Property Set colTipoProduto(ByVal vData As Object)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.colTipoProduto = Form1
    Set mvarcolTipoProduto = vData
End Property


Public Property Get colTipoProduto() As Collection
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.colTipoProduto
    Set colTipoProduto = mvarcolTipoProduto
End Property



Public Property Let iExibeReqCotadas(ByVal vData As Integer)
Attribute iExibeReqCotadas.VB_Description = "Concorrência selecionada."
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iExibeReqCotadas = 5
    mvariExibeReqCotadas = vData
End Property


Public Property Get iExibeReqCotadas() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iExibeReqCotadas
    iExibeReqCotadas = mvariExibeReqCotadas
End Property



Public Property Let dtDataLimiteAte(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtDataLimiteAte = 5
    mvardtDataLimiteAte = vData
End Property


Public Property Get dtDataLimiteAte() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtDataLimiteAte
    dtDataLimiteAte = mvardtDataLimiteAte
End Property



Public Property Let dtDataLimiteDe(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtDataLimiteDe = 5
    mvardtDataLimiteDe = vData
End Property


Public Property Get dtDataLimiteDe() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtDataLimiteDe
    dtDataLimiteDe = mvardtDataLimiteDe
End Property



Public Property Set colReqCompra(ByVal vData As Object)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.colReqCompra = Form1
    Set mvarcolReqCompra = vData
End Property


Public Property Get colReqCompra() As Collection
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.colReqCompra
    Set colReqCompra = mvarcolReqCompra
End Property



Public Property Let sOrdenacaoReq(ByVal vData As String)
Attribute sOrdenacaoReq.VB_Description = "Ordenação de requisições."
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sOrdenacaoReq = 5
    mvarsOrdenacaoReq = vData
End Property


Public Property Get sOrdenacaoReq() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sOrdenacaoReq
    sOrdenacaoReq = mvarsOrdenacaoReq
End Property



Public Property Let iSelecionaDestino(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iSelecionaDestino = 5
    mvariSelecionaDestino = vData
End Property


Public Property Get iSelecionaDestino() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iSelecionaDestino
    iSelecionaDestino = mvariSelecionaDestino
End Property



Public Property Let lCodigoAte(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lCodigoAte = 5
    mvarlCodigoAte = vData
End Property


Public Property Get lCodigoAte() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lCodigoAte
    lCodigoAte = mvarlCodigoAte
End Property



Public Property Let lCodigoDe(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lCodigoDe = 5
    mvarlCodigoDe = vData
End Property


Public Property Get lCodigoDe() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lCodigoDe
    lCodigoDe = mvarlCodigoDe
End Property



Public Property Let iFilialDestino(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iFilialDestino = 5
    mvariFilialDestino = vData
End Property


Public Property Get iFilialDestino() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iFilialDestino
    iFilialDestino = mvariFilialDestino
End Property



Public Property Let lFornCliDestino(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lFornCliDestino = 5
    mvarlFornCliDestino = vData
End Property


Public Property Get lFornCliDestino() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lFornCliDestino
    lFornCliDestino = mvarlFornCliDestino
End Property



Public Property Let iTipoDestino(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iTipoDestino = 5
    mvariTipoDestino = vData
End Property


Public Property Get iTipoDestino() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iTipoDestino
    iTipoDestino = mvariTipoDestino
End Property



Public Property Let dtDataAte(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtDataAte = 5
    mvardtDataAte = vData
End Property


Public Property Get dtDataAte() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtDataAte
    dtDataAte = mvardtDataAte
End Property



Public Property Let dtDataDe(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtDataDe = 5
    mvardtDataDe = vData
End Property


Public Property Get dtDataDe() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtDataDe
    dtDataDe = mvardtDataDe
End Property




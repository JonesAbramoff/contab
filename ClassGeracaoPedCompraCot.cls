VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassGeracaoPedCompraCot"
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
Private mvarsOrdenacaoCot As String 'local copy
Private mvarcolCotacao As Collection 'local copy
Private mvariCotacaoSel As Integer 'local copy
Private mvarsOrdenacaoGerPedCot As String 'local copy
Public Property Let sOrdenacaoGerPedCot(ByVal vData As String)
Attribute sOrdenacaoGerPedCot.VB_Description = "Ordenação de concorrências."
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sOrdenacaoGerPedCot = 5
    mvarsOrdenacaoGerPedCot = vData
End Property


Public Property Get sOrdenacaoGerPedCot() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sOrdenacaoGerPedCot
    sOrdenacaoGerPedCot = mvarsOrdenacaoGerPedCot
End Property



Public Property Let iCotacaoSel(ByVal vData As Integer)
Attribute iCotacaoSel.VB_Description = "Concorrência selecionada."
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iCotacaoSel = 5
    mvariCotacaoSel = vData
End Property


Public Property Get iCotacaoSel() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iCotacaoSel
    iCotacaoSel = mvariCotacaoSel
End Property



Public Property Set colCotacao(ByVal vData As Object)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.colCotacao = Form1
    Set mvarcolCotacao = vData
End Property


Public Property Get colCotacao() As Collection
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.colCotacao
    Set colCotacao = mvarcolCotacao
End Property



Public Property Let sOrdenacaoCot(ByVal vData As String)
Attribute sOrdenacaoCot.VB_Description = "Ordenação de cotações."
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sOrdenacaoCot = 5
    mvarsOrdenacaoCot = vData
End Property


Public Property Get sOrdenacaoCot() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sOrdenacaoCot
    sOrdenacaoCot = mvarsOrdenacaoCot
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




VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassConcorrencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarlNumIntDoc As Long 'local copy
Private mvariFilialEmpresa As Integer 'local copy
Private mvarlCodigo As Long 'local copy
Private mvardtData As Date 'local copy
Private mvariComprador As Integer 'local copy
Private mvardTaxaFinanceira As Double 'local copy
Private mvariTipoDestino As Integer 'local copy
Private mvarlFornCliDestino As Long 'local copy
Private mvariFilialDestino As Integer 'local copy
Private mvarcolItens As New Collection 'local copy
Private mvarsDescricao As String 'local copy
'local variable(s) to hold property value(s)
Private mvarsNomeComprador As String 'local copy
Private mvarsFilialEmpNomeReduzido As String 'local copy
Public Property Let sFilialEmpNomeReduzido(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sFilialEmpNomeReduzido = 5
    mvarsFilialEmpNomeReduzido = vData
End Property


Public Property Get sFilialEmpNomeReduzido() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sFilialEmpNomeReduzido
    sFilialEmpNomeReduzido = mvarsFilialEmpNomeReduzido
End Property



Public Property Let sNomeComprador(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sNomeComprador = 5
    mvarsNomeComprador = vData
End Property


Public Property Get sNomeComprador() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sNomeComprador
    sNomeComprador = mvarsNomeComprador
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



Public Property Set colItens(ByVal vData As Object)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.colItens = Form1
    Set mvarcolItens = vData
End Property


Public Property Get colItens() As Collection
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.colItens
    Set colItens = mvarcolItens
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



Public Property Let dTaxaFinanceira(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dTaxaFinanceira = 5
    mvardTaxaFinanceira = vData
End Property


Public Property Get dTaxaFinanceira() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dTaxaFinanceira
    dTaxaFinanceira = mvardTaxaFinanceira
End Property



Public Property Let iComprador(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iComprador = 5
    mvariComprador = vData
End Property


Public Property Get iComprador() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iComprador
    iComprador = mvariComprador
End Property



Public Property Let dtData(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtData = 5
    mvardtData = vData
End Property


Public Property Get dtData() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtData
    dtData = mvardtData
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




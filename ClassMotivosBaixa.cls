VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassMotivosBaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvariCodigo As Integer 'local copy
Private mvarsDescricao As String 'local copy
Private mvariPagamento As Integer 'local copy
Private mvariRecebimento As Integer 'local copy
Private mvariPagaComissaoVendas As Integer 'local copy
Private mvariSubtipo As Integer 'local copy
Public Property Let iSubtipo(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iSubtipo = 5
    mvariSubtipo = vData
End Property


Public Property Get iSubtipo() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iSubtipo
    iSubtipo = mvariSubtipo
End Property



Public Property Let iPagaComissaoVendas(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iPagaComissaoVendas = 5
    mvariPagaComissaoVendas = vData
End Property


Public Property Get iPagaComissaoVendas() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iPagaComissaoVendas
    iPagaComissaoVendas = mvariPagaComissaoVendas
End Property



Public Property Let iRecebimento(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iRecebimento = 5
    mvariRecebimento = vData
End Property


Public Property Get iRecebimento() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iRecebimento
    iRecebimento = mvariRecebimento
End Property



Public Property Let iPagamento(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iPagamento = 5
    mvariPagamento = vData
End Property


Public Property Get iPagamento() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iPagamento
    iPagamento = mvariPagamento
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



Public Property Let iCodigo(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iCodigo = 5
    mvariCodigo = vData
End Property


Public Property Get iCodigo() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iCodigo
    iCodigo = mvariCodigo
End Property




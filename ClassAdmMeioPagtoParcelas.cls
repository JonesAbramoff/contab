VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassAdmMeioPagtoParcelas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvariAdmMeioPagto As Integer 'local copy
Private mvariFilialEmpresa As Integer 'local copy
Private mvariParcelamento As Integer 'local copy
Private mvardPercRecebimento As Double 'local copy
Private mvariIntervaloRecebimento As Integer 'local copy
Private mvariParcela As Integer 'local copy
Public Property Let iParcela(ByVal vData As Integer)
Attribute iParcela.VB_Description = "Número da Parcela"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iParcela = 5
    mvariParcela = vData
End Property


Public Property Get iParcela() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iParcela
    iParcela = mvariParcela
End Property



Public Property Let iIntervaloRecebimento(ByVal vData As Integer)
Attribute iIntervaloRecebimento.VB_Description = "Intervalo de recebimento entre essa parcela e a anterior"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iIntervaloRecebimento = 5
    mvariIntervaloRecebimento = vData
End Property


Public Property Get iIntervaloRecebimento() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iIntervaloRecebimento
    iIntervaloRecebimento = mvariIntervaloRecebimento
End Property



Public Property Let dPercRecebimento(ByVal vData As Double)
Attribute dPercRecebimento.VB_Description = "Percentual de Recebimento nessa parcela"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dPercRecebimento = 5
    mvardPercRecebimento = vData
End Property


Public Property Get dPercRecebimento() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dPercRecebimento
    dPercRecebimento = mvardPercRecebimento
End Property



Public Property Let iParcelamento(ByVal vData As Integer)
Attribute iParcelamento.VB_Description = "Código do Parcelamento (AdmMeioPagtoCondPagto)"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iParcelamento = 5
    mvariParcelamento = vData
End Property


Public Property Get iParcelamento() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iParcelamento
    iParcelamento = mvariParcelamento
End Property



Public Property Let iFilialEmpresa(ByVal vData As Integer)
Attribute iFilialEmpresa.VB_Description = "Filial da Empresa onde o Meio de Pagamento é aceito"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iFilialEmpresa = 5
    mvariFilialEmpresa = vData
End Property


Public Property Get iFilialEmpresa() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iFilialEmpresa
    iFilialEmpresa = mvariFilialEmpresa
End Property



Public Property Let iAdmMeioPagto(ByVal vData As Integer)
Attribute iAdmMeioPagto.VB_Description = "Código do Meio de Pagamento (AdmMeioPagto)"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iAdmMeioPagto = 5
    mvariAdmMeioPagto = vData
End Property


Public Property Get iAdmMeioPagto() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iAdmMeioPagto
    iAdmMeioPagto = mvariAdmMeioPagto
End Property




VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassValeTicket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlNumIntDoc As Long
Private mvariFilialEmpresa As Integer
Private mvariECF As Integer
Private mvariOperador As Integer
Private mvarsAdmMeioPagto As String
Private mvardtDataRecebimento As Date
Private mvardValor As Double
Private mvarlBordero As Long
Private mvariStatus As Integer
Private mvardValorRecebido As Double
Public Property Let dValorRecebido(ByVal vData As Double)
    mvardValorRecebido = vData
End Property


Public Property Get dValorRecebido() As Double
    dValorRecebido = mvardValorRecebido
End Property



Public Property Let iStatus(ByVal vData As Integer)
    mvariStatus = vData
End Property


Public Property Get iStatus() As Integer
    iStatus = mvariStatus
End Property



Public Property Let lBordero(ByVal vData As Long)
    mvarlBordero = vData
End Property


Public Property Get lBordero() As Long
    lBordero = mvarlBordero
End Property



Public Property Let dValor(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dValor = 5
    mvardValor = vData
End Property


Public Property Get dValor() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dValor
    dValor = mvardValor
End Property



Public Property Let dtDataRecebimento(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtDataRecebimento = 5
    mvardtDataRecebimento = vData
End Property


Public Property Get dtDataRecebimento() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtDataRecebimento
    dtDataRecebimento = mvardtDataRecebimento
End Property



Public Property Let sAdmMeioPagto(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sAdmMeioPagto = 5
    mvarsAdmMeioPagto = vData
End Property


Public Property Get sAdmMeioPagto() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sAdmMeioPagto
    sAdmMeioPagto = mvarsAdmMeioPagto
End Property



Public Property Let iOperador(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iOperador = 5
    mvariOperador = vData
End Property


Public Property Get iOperador() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iOperador
    iOperador = mvariOperador
End Property



Public Property Let iECF(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iECF = 5
    mvariECF = vData
End Property


Public Property Get iECF() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iECF
    iECF = mvariECF
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




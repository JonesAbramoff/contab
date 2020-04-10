VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassBaixaBoletoParcela"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlNumIntDoc As Long
Private mvarNumIntBaixaBoletosParcelas As Long
Private mvarlNumIntBoletoParcela As Long
Private mvariStatus As Integer
Private mvardtDataCancelamento As Date
Public Property Let dtDataCancelamento(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtDataCancelamento = 5
    mvardtDataCancelamento = vData
End Property


Public Property Get dtDataCancelamento() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtDataCancelamento
    dtDataCancelamento = mvardtDataCancelamento
End Property



Public Property Let iStatus(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iStatus = 5
    mvariStatus = vData
End Property


Public Property Get iStatus() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iStatus
    iStatus = mvariStatus
End Property



Public Property Let lNumIntBoletoParcela(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lNumIntBoletoParcela = 5
    mvarlNumIntBoletoParcela = vData
End Property


Public Property Get lNumIntBoletoParcela() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lNumIntBoletoParcela
    lNumIntBoletoParcela = mvarlNumIntBoletoParcela
End Property



Public Property Let NumIntBaixaBoletosParcelas(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NumIntBaixaBoletosParcelas = 5
    mvarNumIntBaixaBoletosParcelas = vData
End Property


Public Property Get NumIntBaixaBoletosParcelas() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NumIntBaixaBoletosParcelas
    NumIntBaixaBoletosParcelas = mvarNumIntBaixaBoletosParcelas
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




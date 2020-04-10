VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassMapaCotacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"ColItensMapaCotacao"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarlNumIntDoc As Long 'local copy
Private mvarlCodigo As Long 'local copy
Private mvariFilialEmpresa As Integer 'local copy
Private mvardtData As Date 'local copy
Private mvardTaxaFinanceira As Double 'local copy
Private mvarColItensMapaCotacao As New Collection

Public Property Set ColItensMapaCotacao(ByVal vData As Object)
    Set mvarColItensMapaCotacao = vData
End Property

Public Property Get ColItensMapaCotacao() As Collection
    Set ColItensMapaCotacao = mvarColItensMapaCotacao
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




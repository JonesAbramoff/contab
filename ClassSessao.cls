VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassSessao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlNumInt As Long
Private mvariCaixa As Integer
Private mvardtDataInicio As Date
Private mvariAbertura As Integer
Private mvardtDataFim As Date
Private mvariFilialEmpresa As Integer
Private mvariOperador As Integer
Private mvarlGerenteAbertura As Long
Private mvarlGerenteFechamento As Long
Private mvariStatus As Integer
Private mvariFechamento As Integer
Public Property Let iFechamento(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iFechamento = 5
    mvariFechamento = vData
End Property


Public Property Get iFechamento() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iFechamento
    iFechamento = mvariFechamento
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



Public Property Let lGerenteFechamento(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lGerenteFechamento = 5
    mvarlGerenteFechamento = vData
End Property


Public Property Get lGerenteFechamento() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lGerenteFechamento
    lGerenteFechamento = mvarlGerenteFechamento
End Property



Public Property Let lGerenteAbertura(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lGerenteAbertura = 5
    mvarlGerenteAbertura = vData
End Property


Public Property Get lGerenteAbertura() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lGerenteAbertura
    lGerenteAbertura = mvarlGerenteAbertura
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



Public Property Let dtDataFim(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtDataFim = 5
    mvardtDataFim = vData
End Property


Public Property Get dtDataFim() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtDataFim
    dtDataFim = mvardtDataFim
End Property



Public Property Let iAbertura(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iAbertura = 5
    mvariAbertura = vData
End Property


Public Property Get iAbertura() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iAbertura
    iAbertura = mvariAbertura
End Property



Public Property Let dtDataInicio(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtDataInicio = 5
    mvardtDataInicio = vData
End Property


Public Property Get dtDataInicio() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtDataInicio
    dtDataInicio = mvardtDataInicio
End Property



Public Property Let iCaixa(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iCaixa = 5
    mvariCaixa = vData
End Property


Public Property Get iCaixa() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iCaixa
    iCaixa = mvariCaixa
End Property



Public Property Let lNumInt(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lNumInt = 5
    mvarlNumInt = vData
End Property


Public Property Get lNumInt() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lNumInt
    lNumInt = mvarlNumInt
End Property




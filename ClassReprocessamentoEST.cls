VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassReprocessamentoEST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvardtDataInicio As Date 'local copy
Private mvardtDataFim As Date 'local copy
Private mvariOrdemReproc As Integer 'local copy
Private mvariLogReproc As Integer 'local copy
Private mvarsProdutoCodigo As String 'local copy
Private mvariPulaFaseDesfaz As Integer 'local copy
Private mvariReprocTestaInt As Integer 'local copy
'local variable(s) to hold property value(s)
Private mvariFilialEmpresa As Integer 'local copy
Private mvariAno As Integer 'utilizado na rotina de apuração de custo de produção para indicar o Ano que está sendo apurado
Private mvariMes As Integer 'utilizado na rotina de apuração de custo de produção para indicar o Mes que está sendo apurado
Private mvariAcertaEstProd As Integer 'utilizado na rotina de apuração de custo de produção para indicar o Mes que está sendo apurado

Public sFilialEmpGrupo As String
Private mvariApenasSaldoTerc As Integer

Public Property Let iApenasSaldoTerc(ByVal vData As Integer)
    mvariApenasSaldoTerc = vData
End Property

Public Property Get iApenasSaldoTerc() As Integer
    iApenasSaldoTerc = mvariApenasSaldoTerc
End Property

Public Property Let iAcertaEstProd(ByVal vData As Integer)
    mvariAcertaEstProd = vData
End Property

Public Property Get iAcertaEstProd() As Integer
    iAcertaEstProd = mvariAcertaEstProd
End Property

Public Property Let iMes(ByVal vData As Integer)
    mvariMes = vData
End Property

Public Property Get iMes() As Integer
    iMes = mvariMes
End Property

Public Property Let iAno(ByVal vData As Integer)
    mvariAno = vData
End Property

Public Property Get iAno() As Integer
    iAno = mvariAno
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



Public Property Let iReprocTestaInt(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iReprocTestaInt = 5
    mvariReprocTestaInt = vData
End Property


Public Property Get iReprocTestaInt() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iReprocTestaInt
    iReprocTestaInt = mvariReprocTestaInt
End Property



Public Property Let iPulaFaseDesfaz(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iPulaFaseDesfaz = 5
    mvariPulaFaseDesfaz = vData
End Property


Public Property Get iPulaFaseDesfaz() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iPulaFaseDesfaz
    iPulaFaseDesfaz = mvariPulaFaseDesfaz
End Property



Public Property Let sProdutoCodigo(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sProdutoCodigo = 5
    mvarsProdutoCodigo = vData
End Property


Public Property Get sProdutoCodigo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sProdutoCodigo
    sProdutoCodigo = mvarsProdutoCodigo
End Property



Public Property Let iLogReproc(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iLogReproc = 5
    mvariLogReproc = vData
End Property


Public Property Get iLogReproc() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iLogReproc
    iLogReproc = mvariLogReproc
End Property



Public Property Let iOrdemReproc(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iOrdemReproc = 5
    mvariOrdemReproc = vData
End Property


Public Property Get iOrdemReproc() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iOrdemReproc
    iOrdemReproc = mvariOrdemReproc
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




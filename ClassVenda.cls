VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassVenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarobjCupomFiscal As New ClassCupomFiscal  'local copy
Private mvarcolMovimentosCaixa As New Collection 'local copy
Private mvarcolCheques As New Collection 'local copy
Private mvarobjCarne As New ClassCarne   'local copy
Private mvariTipo As Integer
Private mvarcolTroca As New Collection 'local copy
Private mvardValorTEF As Double
Private mvardFalta As Double
Private mvarcolVinculado As New Collection 'local copy
Private mvarcolIndiceImpCheque As New Collection 'local copy
Private mvarobjTEF As ClassTEF

Public iCodModeloECF As Integer

Private mvarobjSATInfo As ClassSATInfo
Private mvarobjNFeInfo As ClassNFeInfo

Public bEntregaEmDomicilio As Boolean
Public iTipoForcado As Integer
Public iForcadoF5 As Integer

Public Property Get objSATInfo() As ClassSATInfo
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.objSATInfo
Dim Y As New ClassSATGlobal
    
    If mvarobjSATInfo Is Nothing Then Set mvarobjSATInfo = Y.gobjSATInfo
    
    Set objSATInfo = mvarobjSATInfo

End Property

Public Property Set objSATInfo(ByVal vData As ClassSATInfo)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.objSATInfo = Form1
    Set mvarobjSATInfo = vData
End Property

Public Property Get objNFeInfo() As ClassNFeInfo
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.objNFeInfo
Dim Y As New ClassNFeGlobal
    
    If mvarobjNFeInfo Is Nothing Then Set mvarobjNFeInfo = Y.gobjNFeInfo
    
    Set objNFeInfo = mvarobjNFeInfo

End Property

Public Property Set objNFeInfo(ByVal vData As ClassNFeInfo)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.objNFeInfo = Form1
    Set mvarobjNFeInfo = vData
End Property

Public Property Get objTEF() As ClassTEF
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.objTEF
    Set objTEF = mvarobjTEF
End Property

Public Property Set objTEF(ByVal vData As ClassTEF)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.objTEF = Form1
    Set mvarobjTEF = vData
End Property


Public Property Set colIndiceImpCheque(ByVal vData As Collection)
    Set mvarcolIndiceImpCheque = vData
End Property

Public Property Get colIndiceImpCheque() As Collection
    Set colIndiceImpCheque = mvarcolIndiceImpCheque
End Property

Public Property Set colVinculado(ByVal vData As Collection)
    Set mvarcolVinculado = vData
End Property

Public Property Get colVinculado() As Collection
    Set colVinculado = mvarcolVinculado
End Property

Public Property Let dFalta(ByVal vData As Double)
    mvardFalta = vData
End Property

Public Property Get dFalta() As Double
    dFalta = mvardFalta
End Property

Public Property Let dValorTEF(ByVal vData As Double)
    mvardValorTEF = vData
End Property

Public Property Get dValorTEF() As Double
    dValorTEF = mvardValorTEF
End Property

Public Property Let iTipo(ByVal vData As Integer)
    mvariTipo = vData
End Property

Public Property Get iTipo() As Integer
    iTipo = mvariTipo
End Property

Public Property Set colMovimentosCaixa(ByVal vData As Collection)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.colMovimentosCaixa = Form1
    Set mvarcolMovimentosCaixa = vData
End Property


Public Property Get colMovimentosCaixa() As Collection
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.colMovimentosCaixa
    Set colMovimentosCaixa = mvarcolMovimentosCaixa
End Property

Public Property Set colTroca(ByVal vData As Collection)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.colTroca = Form1
    Set mvarcolTroca = vData
End Property


Public Property Get colTroca() As Collection
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.colTroca
    Set colTroca = mvarcolTroca
End Property

Public Property Set colCheques(ByVal vData As Collection)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.colCheques = Form1
    Set mvarcolCheques = vData
End Property

Public Property Get colCheques() As Collection
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.colCheques
    Set colCheques = mvarcolCheques
End Property


Public Property Set objCupomFiscal(ByVal vData As ClassCupomFiscal)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.objCupomFiscal = Form1
    Set mvarobjCupomFiscal = vData
End Property


Public Property Get objCupomFiscal() As ClassCupomFiscal
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.objCupomFiscal
    Set objCupomFiscal = mvarobjCupomFiscal
End Property

Public Property Set objCarne(ByVal vData As ClassCarne)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.objCarne = Form1
    Set mvarobjCarne = vData
End Property


Public Property Get objCarne() As ClassCarne
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.objCarne
    Set objCarne = mvarobjCarne
End Property

Private Sub Class_Initialize()
    
    iCodModeloECF = 0
    bEntregaEmDomicilio = False
    
End Sub

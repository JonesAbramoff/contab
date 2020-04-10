VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassBaixaBoletosParcelas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit


Private mvarlNumIntBaixa As Long
Private mvariStatus As Integer
Private mvarsHistorico As String
Private mvariMotivo As Integer
Private mvardtData As Date
Private mvardtDataContabil As Date
Private mvardtDataRegistro As Date
Private mvarlNumMovCta As Long
Private mvarcolBaixaBoletoParcela As New Collection
Public Property Set colBaixaBoletoParcela(ByVal vData As Collection)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.colBaixaBoletoParcela = Form1
    Set mvarcolBaixaBoletoParcela = vData
End Property


Public Property Get colBaixaBoletoParcela() As Collection
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.colBaixaBoletoParcela
    Set colBaixaBoletoParcela = mvarcolBaixaBoletoParcela
End Property



Public Property Let lNumMovCta(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lNumMovCta = 5
    mvarlNumMovCta = vData
End Property


Public Property Get lNumMovCta() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lNumMovCta
    lNumMovCta = mvarlNumMovCta
End Property



Public Property Let dtDataRegistro(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtDataRegistro = 5
    mvardtDataRegistro = vData
End Property


Public Property Get dtDataRegistro() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtDataRegistro
    dtDataRegistro = mvardtDataRegistro
End Property



Public Property Let dtDataContabil(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtDataContabil = 5
    mvardtDataContabil = vData
End Property


Public Property Get dtDataContabil() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtDataContabil
    dtDataContabil = mvardtDataContabil
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



Public Property Let iMotivo(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iMotivo = 5
    mvariMotivo = vData
End Property


Public Property Get iMotivo() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iMotivo
    iMotivo = mvariMotivo
End Property



Public Property Let sHistorico(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sHistorico = 5
    mvarsHistorico = vData
End Property


Public Property Get sHistorico() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sHistorico
    sHistorico = mvarsHistorico
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



Public Property Let lNumIntBaixa(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lNumIntBaixa = 5
    mvarlNumIntBaixa = vData
End Property


Public Property Get lNumIntBaixa() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lNumIntBaixa
    lNumIntBaixa = mvarlNumIntBaixa
End Property




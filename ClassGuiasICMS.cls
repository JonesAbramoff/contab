VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassGuiasICMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvariFilialEmpresa As Integer 'local copy
Private mvardtData As Date 'local copy
Private mvarsNumero As String 'local copy
Private mvardValor As Double 'local copy
Private mvarsOrgaoArrecadador As String 'local copy
Private mvardtDataEntrega As Date 'local copy
Private mvarsLocalEntrega As String 'local copy
Private mvardtApuracaoDe As Date 'local copy
Private mvardtApuracaoAte As Date 'local copy
Private mvarsCodReceita As String 'local copy
Private mvarsCodObrigRecolher As String 'local copy
Private mvardtVencimento As Date
Private mvariICMSST As Integer
Private mvarsUF As String 'local copy

Public Property Let sUF(ByVal vData As String)
    mvarsUF = vData
End Property

Public Property Get sUF() As String
    sUF = mvarsUF
End Property

Public Property Let iICMSST(ByVal vData As Integer)
    mvariICMSST = vData
End Property

Public Property Get iICMSST() As Integer
    iICMSST = mvariICMSST
End Property

Public Property Let dtVencimento(ByVal vData As Date)
    mvardtVencimento = vData
End Property

Public Property Get dtVencimento() As Date
    dtVencimento = mvardtVencimento
End Property

Public Property Let sCodObrigRecolher(ByVal vData As String)
    mvarsCodObrigRecolher = vData
End Property

Public Property Get sCodObrigRecolher() As String
    sCodObrigRecolher = mvarsCodObrigRecolher
End Property

Public Property Let sCodReceita(ByVal vData As String)
    mvarsCodReceita = vData
End Property

Public Property Get sCodReceita() As String
    sCodReceita = mvarsCodReceita
End Property

Public Property Let sLocalEntrega(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sLocalEntrega = 5
    mvarsLocalEntrega = vData
End Property


Public Property Get sLocalEntrega() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sLocalEntrega
    sLocalEntrega = mvarsLocalEntrega
End Property

Public Property Let dtApuracaoDe(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtDataEntrega = 5
    mvardtApuracaoDe = vData
End Property


Public Property Get dtApuracaoDe() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtDataEntrega
    dtApuracaoDe = mvardtApuracaoDe
End Property


Public Property Let dtApuracaoAte(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtDataEntrega = 5
    mvardtApuracaoAte = vData
End Property


Public Property Get dtApuracaoAte() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtDataEntrega
    dtApuracaoAte = mvardtApuracaoAte
End Property

Public Property Let dtDataEntrega(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtDataEntrega = 5
    mvardtDataEntrega = vData
End Property


Public Property Get dtDataEntrega() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtDataEntrega
    dtDataEntrega = mvardtDataEntrega
End Property



Public Property Let sOrgaoArrecadador(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sOrgaoArrecadador = 5
    mvarsOrgaoArrecadador = vData
End Property


Public Property Get sOrgaoArrecadador() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sOrgaoArrecadador
    sOrgaoArrecadador = mvarsOrgaoArrecadador
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



Public Property Let sNumero(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sNumero = 5
    mvarsNumero = vData
End Property


Public Property Get sNumero() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sNumero
    sNumero = mvarsNumero
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




VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFornecedorHistorico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarlNumIntDoc As Long 'local copy
Private mvardtDataAtualizacao As Date 'local copy
Private mvarlCodigo As Long 'local copy
Private mvarsCGC As String 'local copy
Private mvarsInscricaoEstadual As String 'local copy
Private mvarsInscricaoMunicipal As String 'local copy
Private mvarsRazaoSocial As String 'local copy
Private mvarsEndereco As String 'local copy
Private mvarsBairro As String 'local copy
Private mvarsCidade As String 'local copy
Private mvarsUF As String 'local copy
Private mvarsPais As String 'local copy
Private mvarsCEP As String 'local copy
Private mvariFilialForn As Integer 'local copy
Private mvarsInscricaoSuframa As String

Public Property Let sInscricaoSuframa(ByVal vData As String)
    mvarsInscricaoSuframa = vData
End Property

Public Property Get sInscricaoSuframa() As String
    sInscricaoSuframa = mvarsInscricaoSuframa
End Property

Public Property Let iFilialForn(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iFilialForn = 5
    mvariFilialForn = vData
End Property


Public Property Get iFilialForn() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iFilialForn
    iFilialForn = mvariFilialForn
End Property



Public Property Let sCEP(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sCEP = 5
    mvarsCEP = vData
End Property


Public Property Get sCEP() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sCEP
    sCEP = mvarsCEP
End Property



Public Property Let sPais(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sPais = 5
    mvarsPais = vData
End Property


Public Property Get sPais() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sPais
    sPais = mvarsPais
End Property



Public Property Let sUF(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sUF = 5
    mvarsUF = vData
End Property


Public Property Get sUF() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sUF
    sUF = mvarsUF
End Property



Public Property Let sCidade(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sCidade = 5
    mvarsCidade = vData
End Property


Public Property Get sCidade() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sCidade
    sCidade = mvarsCidade
End Property



Public Property Let sBairro(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sBairro = 5
    mvarsBairro = vData
End Property


Public Property Get sBairro() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sBairro
    sBairro = mvarsBairro
End Property



Public Property Let sEndereco(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sEndereco = 5
    mvarsEndereco = vData
End Property


Public Property Get sEndereco() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sEndereco
    sEndereco = mvarsEndereco
End Property



Public Property Let sRazaoSocial(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sRazaoSocial = 5
    mvarsRazaoSocial = vData
End Property


Public Property Get sRazaoSocial() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sRazaoSocial
    sRazaoSocial = mvarsRazaoSocial
End Property



Public Property Let sInscricaoMunicipal(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sInscricaoMunicipal = 5
    mvarsInscricaoMunicipal = vData
End Property


Public Property Get sInscricaoMunicipal() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sInscricaoMunicipal
    sInscricaoMunicipal = mvarsInscricaoMunicipal
End Property



Public Property Let sInscricaoEstadual(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sInscricaoEstadual = 5
    mvarsInscricaoEstadual = vData
End Property


Public Property Get sInscricaoEstadual() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sInscricaoEstadual
    sInscricaoEstadual = mvarsInscricaoEstadual
End Property



Public Property Let sCgc(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sCGC = 5
    mvarsCGC = vData
End Property


Public Property Get sCgc() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sCGC
    sCgc = mvarsCGC
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


Public Property Let dtDataAtualizacao(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtDataAtualizacao = 5
    mvardtDataAtualizacao = vData
End Property


Public Property Get dtDataAtualizacao() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtDataAtualizacao
    dtDataAtualizacao = mvardtDataAtualizacao
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




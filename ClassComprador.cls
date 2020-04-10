VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassComprador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvariFilialEmpresa As Integer 'local copy
Private mvariCodigo As Integer 'local copy
Private mvarsCodUsuario As String 'local copy
Private mvariAumentaQuant As Integer 'local copy
Private mvarsNome As String 'local copy
Private mvarsNomeReduzido As String 'local copy
Private mvarsEmail As String

Public Property Let sEmail(ByVal vData As String)
    mvarsEmail = vData
End Property

Public Property Get sEmail() As String
    sEmail = mvarsEmail
End Property

Public Property Let sNomeReduzido(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sNomeReduzido = 5
    mvarsNomeReduzido = vData
End Property


Public Property Get sNomeReduzido() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sNomeReduzido
    sNomeReduzido = mvarsNomeReduzido
End Property



Public Property Let sNome(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sNome = 5
    mvarsNome = vData
End Property


Public Property Get sNome() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sNome
    sNome = mvarsNome
End Property



Public Property Let iAumentaQuant(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iAumentaQuant = 5
    mvariAumentaQuant = vData
End Property


Public Property Get iAumentaQuant() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iAumentaQuant
    iAumentaQuant = mvariAumentaQuant
End Property



Public Property Let sCodUsuario(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sCodUsuario = 5
    mvarsCodUsuario = vData
End Property


Public Property Get sCodUsuario() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sCodUsuario
    sCodUsuario = mvarsCodUsuario
End Property



Public Property Let iCodigo(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iCodigo = 5
    mvariCodigo = vData
End Property


Public Property Get iCodigo() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iCodigo
    iCodigo = mvariCodigo
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




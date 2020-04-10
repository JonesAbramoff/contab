VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRequisitante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarlCodigo As Long 'local copy
Private mvarsNome As String 'local copy
Private mvarsNomeReduzido As String 'local copy
Private mvarsCCL As String 'local copy
Private mvarsEmail As String
Private mvarsCodUsuario As String

Public Property Let sCodUsuario(ByVal vData As String)
    mvarsCodUsuario = vData
End Property

Public Property Get sCodUsuario() As String
    sCodUsuario = mvarsCodUsuario
End Property

Public Property Let sEmail(ByVal vData As String)
    mvarsEmail = vData
End Property

Public Property Get sEmail() As String
    sEmail = mvarsEmail
End Property

Public Property Let sCCL(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sCcl = 5
    mvarsCCL = vData
End Property


Public Property Get sCCL() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sCcl
    sCCL = mvarsCCL
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




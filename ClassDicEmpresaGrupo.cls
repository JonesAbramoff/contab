VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassDicEmpresaGrupo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Códigos de Empresa e Grupo (associados ao Usuário)."
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarCodEmpresa As Long 'local copy
Private mvarCodGrupo As String 'local copy
Public Property Let sCodGrupo(ByVal vData As String)
Attribute sCodGrupo.VB_Description = "Código de Grupo."
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CodGrupo = 5
    mvarCodGrupo = vData
End Property


Public Property Get sCodGrupo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CodGrupo
    sCodGrupo = mvarCodGrupo
End Property



Public Property Let lCodEmpresa(ByVal vData As Long)
Attribute lCodEmpresa.VB_Description = "Código de Empresa."
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CodEmpresa = 5
    mvarCodEmpresa = vData
End Property


Public Property Get lCodEmpresa() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CodEmpresa
    lCodEmpresa = mvarCodEmpresa
End Property




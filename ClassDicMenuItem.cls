VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassDicMenuItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Classe MenuItem."
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarsTitulo As String 'local copy
Private mvarsSiglaRotina As String 'local copy
Private mvarNomeTela As String 'local copy
'local variable(s) to hold property value(s)
Private mvariIdentificador As Integer 'local copy
Public Property Let iIdentificador(ByVal vData As Integer)
Attribute iIdentificador.VB_Description = "Identificador de Item de Menu."
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iIdentificador = 5
    mvariIdentificador = vData
End Property


Public Property Get iIdentificador() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iIdentificador
    iIdentificador = mvariIdentificador
End Property



Public Property Let sNomeTela(ByVal vData As String)
Attribute sNomeTela.VB_Description = "Nome de Tela."
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NomeTela = 5
    mvarNomeTela = vData
End Property


Public Property Get sNomeTela() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NomeTela
    sNomeTela = mvarNomeTela
End Property



Public Property Let sSiglaRotina(ByVal vData As String)
Attribute sSiglaRotina.VB_Description = "Sigla de Rotina."
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sSiglaRotina = 5
    mvarsSiglaRotina = vData
End Property


Public Property Get sSiglaRotina() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sSiglaRotina
    sSiglaRotina = mvarsSiglaRotina
End Property



Public Property Let sTitulo(ByVal vData As String)
Attribute sTitulo.VB_Description = "Titulo (caption) do ítem de menu."
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sTitulo = 5
    mvarsTitulo = vData
End Property


Public Property Get sTitulo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sTitulo
    sTitulo = mvarsTitulo
End Property




VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassDicDados"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Variáveis globais de Dicionário de Dados."
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit


Public Property Let gsGrupo(ByVal vData As String)
Attribute gsGrupo.VB_Description = "Código do Grupo corrente."
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.gsGrupo = 5
    DicGlob_sGrupo = vData
End Property


Public Property Get gsGrupo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.gsGrupo
    gsGrupo = DicGlob_sGrupo
End Property



Public Property Let gsTela(ByVal vData As String)
Attribute gsTela.VB_Description = "Nome da Tela corrente."
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.gsTela = 5
    DicGlob_sTela = vData
End Property


Public Property Get gsTela() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.gsTela
    gsTela = DicGlob_sTela
End Property



Public Property Let gsRotina(ByVal vData As String)
Attribute gsRotina.VB_Description = "Sigla da Rotina corrente."
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.gsRotina = 5
    DicGlob_sRotina = vData
End Property


Public Property Get gsRotina() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.gsRotina
    gsRotina = DicGlob_sRotina
End Property

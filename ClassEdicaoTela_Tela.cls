VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassEdicaoTela_Tela"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarcolEdicaoTela_Controle As Collection
Private mvarsNomeTela As String
Private mvarsTabIndex As String
'local variable(s) to hold property value(s)
Private mvarsTabIndexPadrao As String 'local copy
Private mvarsGrupoUsuarios As String

Public Property Get sGrupoUsuarios() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sGrupoUsuarios
    sGrupoUsuarios = mvarsGrupoUsuarios
End Property

Public Property Let sGrupoUsuarios(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sGrupoUsuarios = 5
    mvarsGrupoUsuarios = vData
End Property

Public Property Let sTabIndexPadrao(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sTabIndexPadrao = 5
    mvarsTabIndexPadrao = vData
End Property


Public Property Get sTabIndexPadrao() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sTabIndexPadrao
    sTabIndexPadrao = mvarsTabIndexPadrao
End Property

Public Property Let sTabIndex(ByVal vData As String)
Attribute sTabIndex.VB_Description = "Lista de Todos os Tab Index da tela\r\n"
    mvarsTabIndex = vData
End Property

Public Property Get sTabIndex() As String
    sTabIndex = mvarsTabIndex
End Property

Public Property Let sNomeTela(ByVal vData As String)
Attribute sNomeTela.VB_Description = "Nome da Tela onde contém o controle"
    mvarsNomeTela = vData
End Property

Public Property Get sNomeTela() As String
    sNomeTela = mvarsNomeTela
End Property

Public Property Set colEdicaoTela_Controle(ByVal vData As Collection)
Attribute colEdicaoTela_Controle.VB_Description = "Coleção que contém todos os controles"
    Set mvarcolEdicaoTela_Controle = vData
End Property

Public Property Get colEdicaoTela_Controle() As Collection
    Set colEdicaoTela_Controle = mvarcolEdicaoTela_Controle
End Property

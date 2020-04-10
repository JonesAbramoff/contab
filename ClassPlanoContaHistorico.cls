VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPlanoContaHistorico"
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
Private mvarsCodigo As String 'local copy
Private mvariTipo As Integer 'local copy
Private mvarsContaPai As String 'local copy
Private mvarsDescricao As String 'local copy
Public Property Let sDescricao(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sDescricao = 5
    mvarsDescricao = vData
End Property


Public Property Get sDescricao() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sDescricao
    sDescricao = mvarsDescricao
End Property



Public Property Let sContaPai(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sContaPai = 5
    mvarsContaPai = vData
End Property


Public Property Get sContaPai() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sContaPai
    sContaPai = mvarsContaPai
End Property



Public Property Let iTipo(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iTipo = 5
    mvariTipo = vData
End Property


Public Property Get iTipo() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iTipo
    iTipo = mvariTipo
End Property



Public Property Let sCodigo(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sCodigo = 5
    mvarsCodigo = vData
End Property


Public Property Get sCodigo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sCodigo
    sCodigo = mvarsCodigo
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




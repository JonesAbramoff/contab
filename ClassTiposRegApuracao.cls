VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTiposRegApuracao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariCodigo As Integer
Private mvarsDescricao As String
Private mvariSecao As Integer
Private mvariPreCadastrado As Integer
Private mvariLancamento As Integer
'local variable(s) to hold property value(s)
Private mvarsNomeTabela As String 'local copy
Public Property Let sNomeTabela(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sNomeTabela = 5
    mvarsNomeTabela = vData
End Property


Public Property Get sNomeTabela() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sNomeTabela
    sNomeTabela = mvarsNomeTabela
End Property




Public Property Let iLancamento(ByVal vData As Integer)
    mvariLancamento = vData
End Property

Public Property Get iLancamento() As Integer
    iLancamento = mvariLancamento
End Property

Public Property Let iPreCadastrado(ByVal vData As Integer)
    mvariPreCadastrado = vData
End Property

Public Property Get iPreCadastrado() As Integer
    iPreCadastrado = mvariPreCadastrado
End Property

Public Property Let iSecao(ByVal vData As Integer)
    mvariSecao = vData
End Property

Public Property Get iSecao() As Integer
    iSecao = mvariSecao
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let iCodigo(ByVal vData As Integer)
    mvariCodigo = vData
End Property

Public Property Get iCodigo() As Integer
    iCodigo = mvariCodigo
End Property

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassInvLote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvariFilialEmpresa As Integer 'local copy
Private mvariLote As Integer 'local copy
Private mvarsDescricao As String 'local copy
Private mvariNumItensInf As Integer 'local copy
Private mvariNumItensAtual As Integer 'local copy
Private mvariIdAtualizacao As Integer 'local copy

Public Property Let iIDAtualizacao(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iIdAtualizacao = 5
    mvariIdAtualizacao = vData
End Property

Public Property Get iIDAtualizacao() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iIdAtualizacao
    iIDAtualizacao = mvariIdAtualizacao
End Property

Public Property Let iNumItensAtual(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iNumItensAtual = 5
    mvariNumItensAtual = vData
End Property


Public Property Get iNumItensAtual() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iNumItensAtual
    iNumItensAtual = mvariNumItensAtual
End Property



Public Property Let iNumItensInf(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iNumItensInf = 5
    mvariNumItensInf = vData
End Property


Public Property Get iNumItensInf() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iNumItensInf
    iNumItensInf = mvariNumItensInf
End Property

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



Public Property Let iLote(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iLote = 5
    mvariLote = vData
End Property


Public Property Get iLote() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iLote
    iLote = mvariLote
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




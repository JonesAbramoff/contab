VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTributoDocItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'CLASSE ABSTRATA p/definir a interface a ser implementada por um item de docto que queira interfacear com as classes de tributacao

Public Property Get objTributacaoDocItem() As ClassTributacaoDocItemAux

End Property

Public Property Get objDoc() As ClassTributoDoc

End Property

Function Ativar(objDoc As ClassTributoDoc) As Long
'inicializacoes
End Function

Function Desativar() As Long
'finalizacoes
End Function

Function ObterVar(sIdentificador As String, objVariavelCalculo As ClassVariavelCalculo) As Long
'retorna o valor solicitado se este estiver "disponivel" em seu "contexto"
'objVariavelCalculo.sIdentificador tem que vir preenchida
End Function

Function GuardarVar(objVariavelCalculo As ClassVariavelCalculo) As Long
'armazena info de variavel guardando se esta foi alterada ou nao desde ...
End Function

Function GuardarVarTela(sIdentificador As String, vValor As Variant) As Long
'armazena info de variavel com os parametros passados um a um

End Function

Function Alterado() As Long
'finalizacoes
End Function

Function Calculado() As Long
'finalizacoes
End Function

Function iJaCalculado() As Integer
'finalizacoes
End Function

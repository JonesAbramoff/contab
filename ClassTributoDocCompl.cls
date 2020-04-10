VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassTributoDocCompl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'CLASSE ABSTRATA p/definir a interface a ser implementada por um complemento de docto que queira interfacear com as classes de tributacao
'utilizada p/tratar frete, seguro, etc

Public Property Get objDoc() As ClassTributoDoc

End Property

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




VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTributoDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'CLASSE ABSTRATA p/definir a interface a ser implementada por um docto que queira interfacear com as classes de tributacao

'exemplo de interface que terá que ser implementada pelos objetos que
'solicitarao o calculo de impostos

'tem que saber transferir obj <-> vars
'    ('a tela que edita o doc tem que saber transferir obj <-> tela)

Public Property Get objTributacaoDoc() As ClassTributacaoDocAux

End Property

Function Ativar() As Long
'inicializacoes
End Function

Function Alterado() As Long
    'marcar tendo ocorrido uma alteracao que impeça o aproveitamento do calculo anterior dos itens
End Function

Function Desativar() As Long
'finalizacoes
End Function

Function ObterColDocCompl(colDocCompl As Collection) As Long
'retorna colecao de complementos ( c/frete, seguro, desconto e outras despesas/receitas ) que implementam interface descrita em ClassTributoDocItem

End Function

Function ObterColDocItens(colDocItens As Collection) As Long
'retorna colecao de itens que implementam interface descrita em ClassTributoDocItem

End Function

Function ObterDocAux(objTributacaoDoc As ClassTributacaoDocAux) As Long
'retorna objeto auxiliar vinculado ao Doc

End Function

Function ObterVar(sIdentificador As String, objVariavelCalculo As ClassVariavelCalculo) As Long
'retorna o valor solicitado se este estiver "disponivel" em seu "contexto"
'objVariavelCalculo.sIdentificador tem que vir preenchida
    
End Function

Function GuardarVar(objVariavelCalculo As ClassVariavelCalculo) As Long
'armazena info de variavel
        
End Function

Function GuardarVarTela(sIdentificador As String, vValor As Variant) As Long
'armazena info de variavel com os parametros passados um a um

End Function

Function ObterNatOpPadrao(sNatOp As String) As Long
'retorna natureza de operacao padrao

End Function

Function ObterTribPadrao(iTipoTrib As Integer) As Long
'retorna o tipo de tributacao padrao

End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassTributoDriver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'CLASSE ABSTRATA p/definir a interface de um driver p/tratamento de impostos

Function AtualizarImposto(objDoc As ClassTributoDoc, iRecalculaAutomatica As Integer) As Long
'recalcula a tributacao de todos os itens de um doc e dele como um todo

End Function

'Function AtualizarImpostoItem(objDocItem As ClassTributoDocItem) As Long
''recalcula a tributacao do docitem e do doc como um todo
'
'End Function

Function Ativar() As Long
    'permite a inicializacao do driver
    'pode ser utilizado p/criar cache etc
End Function

Function Desativar() As Long
    'informa que o driver nao vai mais ser necessario
    'pode ser utilizado p/liberar cache etc
End Function


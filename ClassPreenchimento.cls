VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassPreenchimento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Se referencia ao preechimento dos segmentos"
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariPreenchimento As Integer
Private mvarsDescricao As String

Public Property Let sDescricao(ByVal vData As String)
Attribute sDescricao.VB_Description = "Descrição do preenchimento tal como é visto nas telas."
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let iPreenchimento(ByVal vData As Integer)
Attribute iPreenchimento.VB_Description = "1 = Com zeros a esquerda numerico/brancos a direita alfa. 0 = Espaços"
    mvariPreenchimento = vData
End Property

Public Property Get iPreenchimento() As Integer
    iPreenchimento = mvariPreenchimento
End Property




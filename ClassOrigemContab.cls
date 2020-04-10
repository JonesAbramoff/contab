VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassOrigemContab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarsDescricao As String
Private mvarsOrigem As String
Private mvariContabilidade As Integer

Public Property Let iContabilidade(ByVal vData As Integer)
Attribute iContabilidade.VB_Description = "1=indica que a origem pode ser alterada/criada pelo módulo de contabilidade. 0=não pode ser alterada/criada pelo modulo de contabilidade"
    mvariContabilidade = vData
End Property

Public Property Get iContabilidade() As Integer
    iContabilidade = mvariContabilidade
End Property

Public Property Let sOrigem(ByVal vData As String)
Attribute sOrigem.VB_Description = "CPR=Contas a Pagar/Receber, EST= Estoque, CTB=contabilidade etc."
    mvarsOrigem = vData
End Property

Public Property Get sOrigem() As String
    sOrigem = mvarsOrigem
End Property

Public Property Let sDescricao(ByVal vData As String)
Attribute sDescricao.VB_Description = "Nome descritivo da origem. Ex: Estoque, Contabilidade, etc."
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Private Sub Class_Initialize()
    mvarsDescricao = String(STRING_ORIGEM_DESCRICAO, 0)
    mvarsOrigem = String(STRING_ORIGEM, 0)
End Sub

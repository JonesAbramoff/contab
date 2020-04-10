VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassContaCategoria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarsNome As String
Private mvariApuracao As Integer
Private mvariCodigo As Integer

Public Property Let iCodigo(ByVal vData As Integer)
Attribute iCodigo.VB_Description = "Codigo identificador da categoria."
    mvariCodigo = vData
End Property

Public Property Get iCodigo() As Integer
    iCodigo = mvariCodigo
End Property

Public Property Let iApuracao(ByVal vData As Integer)
Attribute iApuracao.VB_Description = "1=Faz parte da apuração 0=não faz parte da apuração"
    mvariApuracao = vData
End Property

Public Property Get iApuracao() As Integer
    iApuracao = mvariApuracao
End Property

Public Property Let sNome(ByVal vData As String)
Attribute sNome.VB_Description = "Categoria da Conta. Ex: Ativo, Passivo, Receita, Despesa,..."
    mvarsNome = vData
End Property

Public Property Get sNome() As String
    sNome = mvarsNome
End Property




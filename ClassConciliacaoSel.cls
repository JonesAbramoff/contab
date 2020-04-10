VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassConciliacaoSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "guarda os dados da ultima selecao dos movimentos que devem ser exibidos para conciliacao"
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariCodCCI As Integer
Private mvardtDataDe As Date
Private mvardtDataAte As Date
Private mvardValorDe As Double
Private mvardValorAte As Double
Private mvariExibeNaoConciliados As Integer

Public Property Let iExibeNaoConciliados(ByVal vData As Integer)
Attribute iExibeNaoConciliados.VB_Description = "1= Indica que so exibe os movimentos não conciliados. 0 = Exibe todos os movimentos."
    mvariExibeNaoConciliados = vData
End Property

Public Property Get iExibeNaoConciliados() As Integer
    iExibeNaoConciliados = mvariExibeNaoConciliados
End Property

Public Property Let dValorAte(ByVal vData As Double)
Attribute dValorAte.VB_Description = "Valor Final"
    mvardValorAte = vData
End Property

Public Property Get dValorAte() As Double
    dValorAte = mvardValorAte
End Property

Public Property Let dValorDe(ByVal vData As Double)
Attribute dValorDe.VB_Description = "Valor Inicial "
    mvardValorDe = vData
End Property

Public Property Get dValorDe() As Double
    dValorDe = mvardValorDe
End Property

Public Property Let dtDataAte(ByVal vData As Date)
Attribute dtDataAte.VB_Description = "Data Final"
    mvardtDataAte = vData
End Property

Public Property Get dtDataAte() As Date
    dtDataAte = mvardtDataAte
End Property

Public Property Let dtDataDe(ByVal vData As Date)
Attribute dtDataDe.VB_Description = "Data inicial"
    mvardtDataDe = vData
End Property

Public Property Get dtDataDe() As Date
    dtDataDe = mvardtDataDe
End Property

Public Property Let iCodCCI(ByVal vData As Integer)
Attribute iCodCCI.VB_Description = "identificador da conta corrente interna"
    mvariCodCCI = vData
End Property

Public Property Get iCodCCI() As Integer
    iCodCCI = mvariCodCCI
End Property




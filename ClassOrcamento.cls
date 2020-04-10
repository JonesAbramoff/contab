VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassOrcamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariExercicio As Integer
Private mvarsConta As String
Private mvariPeriodo As Integer
Private mvarsCcl As String
Private mvardValor As Double
Private mvariFilialEmpresa As Integer

Public Property Let iFilialEmpresa(ByVal vData As Integer)
Attribute iFilialEmpresa.VB_Description = "identifica a filial ou estabelecimento da empresa do usuario(zero p/empresa toda)"
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let dValor(ByVal vData As Double)
Attribute dValor.VB_Description = "Valor orçado"
    mvardValor = vData
End Property

Public Property Get dValor() As Double
    dValor = mvardValor
End Property

Public Property Let sCcl(ByVal vData As String)
Attribute sCcl.VB_Description = "Centro de Custo/Centro de Lucro"
    mvarsCcl = vData
End Property

Public Property Get sCcl() As String
    sCcl = mvarsCcl
End Property

Public Property Let iPeriodo(ByVal vData As Integer)
    mvariPeriodo = vData
End Property

Public Property Get iPeriodo() As Integer
    iPeriodo = mvariPeriodo
End Property

Public Property Let sConta(ByVal vData As String)
Attribute sConta.VB_Description = "Conta contábil"
    mvarsConta = vData
End Property

Public Property Get sConta() As String
    sConta = mvarsConta
End Property

Public Property Let iExercicio(ByVal vData As Integer)
    mvariExercicio = vData
End Property

Public Property Get iExercicio() As Integer
    iExercicio = mvariExercicio
End Property



Private Sub Class_Initialize()
    mvarsConta = String(STRING_CONTA, 0)
    mvarsCcl = String(STRING_CCL, 0)
End Sub

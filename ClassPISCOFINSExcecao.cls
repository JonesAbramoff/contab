VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPISCOFINSExcecao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarsCategoriaCliente As String
Private mvarsCategoriaClienteItem As String
Private mvarsCategoriaProduto As String
Private mvarsCategoriaProdutoItem As String
Private mvariTipo As Integer
Private mvariTipoPIS As Integer
Private mvariTipoCOFINS As Integer
Private mvariPISTipoCalculo As Integer
Private mvariCOFINSTipoCalculo As Integer
Private mvardAliquotaPisRS As Double
Private mvardAliquotaPisPerc As Double
Private mvardAliquotaCofinsRS As Double
Private mvardAliquotaCofinsPerc As Double
Private mvarsFundamentacao As String
Private mvariPrioridade As Integer
Private mvariTipoCliForn As Integer
Private mvarsCategoriaFornecedor As String
Private mvarsCategoriaFornecedorItem As String
Private mvariTipoPISE As Integer
Private mvariTipoCOFINSE As Integer

Public Property Let iTipoPISE(ByVal vData As Integer)
    mvariTipoPISE = vData
End Property

Public Property Get iTipoPISE() As Integer
    iTipoPISE = mvariTipoPISE
End Property

Public Property Let iTipoCOFINSE(ByVal vData As Integer)
    mvariTipoCOFINSE = vData
End Property

Public Property Get iTipoCOFINSE() As Integer
    iTipoCOFINSE = mvariTipoCOFINSE
End Property

Public Property Let sCategoriaFornecedorItem(ByVal vData As String)
    mvarsCategoriaFornecedorItem = vData
End Property

Public Property Get sCategoriaFornecedorItem() As String
    sCategoriaFornecedorItem = mvarsCategoriaFornecedorItem
End Property

Public Property Let sCategoriaFornecedor(ByVal vData As String)
    mvarsCategoriaFornecedor = vData
End Property

Public Property Get sCategoriaFornecedor() As String
    sCategoriaFornecedor = mvarsCategoriaFornecedor
End Property

Public Property Let iTipoCliForn(ByVal vData As Integer)
    mvariTipoCliForn = vData
End Property

Public Property Get iTipoCliForn() As Integer
    iTipoCliForn = mvariTipoCliForn
End Property

Public Property Let sCategoriaCliente(ByVal vData As String)
    mvarsCategoriaCliente = vData
End Property

Public Property Get sCategoriaCliente() As String
    sCategoriaCliente = mvarsCategoriaCliente
End Property

Public Property Let sCategoriaClienteItem(ByVal vData As String)
    mvarsCategoriaClienteItem = vData
End Property

Public Property Get sCategoriaClienteItem() As String
    sCategoriaClienteItem = mvarsCategoriaClienteItem
End Property

Public Property Let sCategoriaProduto(ByVal vData As String)
    mvarsCategoriaProduto = vData
End Property

Public Property Get sCategoriaProduto() As String
    sCategoriaProduto = mvarsCategoriaProduto
End Property

Public Property Let sCategoriaProdutoItem(ByVal vData As String)
    mvarsCategoriaProdutoItem = vData
End Property

Public Property Get sCategoriaProdutoItem() As String
    sCategoriaProdutoItem = mvarsCategoriaProdutoItem
End Property

Public Property Let iTipo(ByVal vData As Integer)
    mvariTipo = vData
End Property

Public Property Get iTipo() As Integer
    iTipo = mvariTipo
End Property

Public Property Let iTipoPIS(ByVal vData As Integer)
    mvariTipoPIS = vData
End Property

Public Property Get iTipoPIS() As Integer
    iTipoPIS = mvariTipoPIS
End Property

Public Property Let iTipoCOFINS(ByVal vData As Integer)
    mvariTipoCOFINS = vData
End Property

Public Property Get iTipoCOFINS() As Integer
    iTipoCOFINS = mvariTipoCOFINS
End Property

Public Property Let iPISTipoCalculo(ByVal vData As Integer)
    mvariPISTipoCalculo = vData
End Property

Public Property Get iPISTipoCalculo() As Integer
    iPISTipoCalculo = mvariPISTipoCalculo
End Property

Public Property Let iCOFINSTipoCalculo(ByVal vData As Integer)
    mvariCOFINSTipoCalculo = vData
End Property

Public Property Get iCOFINSTipoCalculo() As Integer
    iCOFINSTipoCalculo = mvariCOFINSTipoCalculo
End Property

Public Property Let dAliquotaPisRS(ByVal vData As Double)
    mvardAliquotaPisRS = vData
End Property

Public Property Get dAliquotaPisRS() As Double
    dAliquotaPisRS = mvardAliquotaPisRS
End Property

Public Property Let dAliquotaPisPerc(ByVal vData As Double)
    mvardAliquotaPisPerc = vData
End Property

Public Property Get dAliquotaPisPerc() As Double
    dAliquotaPisPerc = mvardAliquotaPisPerc
End Property

Public Property Let dAliquotaCofinsRS(ByVal vData As Double)
    mvardAliquotaCofinsRS = vData
End Property

Public Property Get dAliquotaCofinsRS() As Double
    dAliquotaCofinsRS = mvardAliquotaCofinsRS
End Property

Public Property Let dAliquotaCofinsPerc(ByVal vData As Double)
    mvardAliquotaCofinsPerc = vData
End Property

Public Property Get dAliquotaCofinsPerc() As Double
    dAliquotaCofinsPerc = mvardAliquotaCofinsPerc
End Property

Public Property Let sFundamentacao(ByVal vData As String)
    mvarsFundamentacao = vData
End Property

Public Property Get sFundamentacao() As String
    sFundamentacao = mvarsFundamentacao
End Property

Public Property Let iPrioridade(ByVal vData As Integer)
    mvariPrioridade = vData
End Property

Public Property Get iPrioridade() As Integer
    iPrioridade = mvariPrioridade
End Property

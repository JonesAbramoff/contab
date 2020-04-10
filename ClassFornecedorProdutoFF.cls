VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFornecedorProdutoFF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Private mvariFilialEmpresa As Integer
Private mvarsProduto As String
Private mvarlFornecedor As Long
Private mvariFilialForn As Integer 'local copy
Private mvarsProdutoFornecedor As String 'local copy
Private mvardLoteMinimo As Double 'local copy
Private mvariNota As Integer 'local copy
Private mvardQuantPedAbertos As Double 'local copy
Private mvardtDataUltimaCompra As Date 'local copy
Private mvardTempoRessup As Double 'local copy
Private mvardQuantPedida As Double 'local copy
Private mvarsUMQuantPedida As String
Private mvardQuantRecebida As Double 'local copy
Private mvarsUMQuantRecebida As String
Private mvardtDataPedido As Date 'local copy
Private mvardtDataReceb As Date 'local copy
Private mvardPrecoTotal As Double 'local copy
Private mvardUltimaCotacao As Double 'local copy
Private mvardtDataUltimaCotacao As Date 'local copy
Private mvariTipoFreteUltimaCotacao As Integer 'local copy
Private mvardQuantUltimaCotacao As Double 'local copy
Private mvarsUMQuantUltimaCotacao As String
Private mvariPadrao As Integer 'local copy
Private mvariCondPagto As Integer 'local copy
Private mvarsCondPagto As String 'local copy
Private mvarlNumIntCotacaoProduto As Long 'local copy
Private mvariEscolhido As Integer 'local copy
Private mvarlNumIntDoc As Long 'local copy
Private mvariSelecionado As Integer 'local copy
Private mvarsFilialFornecedorNome As String
Private mvarsProdutoNomeReduzido As String
Private mvarsProdutoDescricao As String
Private mvarsFornecedorNomeReduzido As String
Private mvardtDataPrimeiraCompra As Date 'local copy
'local variable(s) to hold property value(s)
Private mvarsDescricao As String 'local copy
'local variable(s) to hold property value(s)
Private mvarsFilialEmpNomeReduzido As String 'local copy
Public Property Let sFilialEmpNomeReduzido(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sFilialEmpNomeReduzido = 5
    mvarsFilialEmpNomeReduzido = vData
End Property


Public Property Get sFilialEmpNomeReduzido() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sFilialEmpNomeReduzido
    sFilialEmpNomeReduzido = mvarsFilialEmpNomeReduzido
End Property



Public Property Let sDescricao(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sDescricao = 5
    mvarsDescricao = vData
End Property


Public Property Get sDescricao() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sDescricao
    sDescricao = mvarsDescricao
End Property




Public Property Let dtDataPrimeiraCompra(ByVal vData As Date)
    mvardtDataPrimeiraCompra = vData
End Property

Public Property Get dtDataPrimeiraCompra() As Date
    dtDataPrimeiraCompra = mvardtDataPrimeiraCompra
End Property

Public Property Let sFornecedorNomeReduzido(ByVal vData As String)
    mvarsFornecedorNomeReduzido = vData
End Property

Public Property Get sFornecedorNomeReduzido() As String
    sFornecedorNomeReduzido = mvarsFornecedorNomeReduzido
End Property

Public Property Let sProdutoDescricao(ByVal vData As String)
    mvarsProdutoDescricao = vData
End Property

Public Property Get sProdutoDescricao() As String
    sProdutoDescricao = mvarsProdutoDescricao
End Property

Public Property Let sProdutoNomeReduzido(ByVal vData As String)
    mvarsProdutoNomeReduzido = vData
End Property

Public Property Get sProdutoNomeReduzido() As String
    sProdutoNomeReduzido = mvarsProdutoNomeReduzido
End Property

Public Property Let sFilialFornecedorNome(ByVal vData As String)
    mvarsFilialFornecedorNome = vData
End Property

Public Property Get sFilialFornecedorNome() As String
    sFilialFornecedorNome = mvarsFilialFornecedorNome
End Property

Public Property Let iSelecionado(ByVal vData As Integer)
    mvariSelecionado = vData
End Property

Public Property Get iSelecionado() As Integer
    iSelecionado = mvariSelecionado
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let iEscolhido(ByVal vData As Integer)
    mvariEscolhido = vData
End Property

Public Property Get iEscolhido() As Integer
    iEscolhido = mvariEscolhido
End Property

Public Property Let lNumIntCotacaoProduto(ByVal vData As Long)
    mvarlNumIntCotacaoProduto = vData
End Property

Public Property Get lNumIntCotacaoProduto() As Long
    lNumIntCotacaoProduto = mvarlNumIntCotacaoProduto
End Property

Public Property Let sCondPagto(ByVal vData As String)
Attribute sCondPagto.VB_Description = "Descrição Reduzida da Cond. Pagto."
    mvarsCondPagto = vData
End Property

Public Property Get sCondPagto() As String
    sCondPagto = mvarsCondPagto
End Property

Public Property Let iCondPagto(ByVal vData As Integer)
Attribute iCondPagto.VB_Description = "Código da cond. pagto."
    mvariCondPagto = vData
End Property

Public Property Get iCondPagto() As Integer
    iCondPagto = mvariCondPagto
End Property

Public Property Let iPadrao(ByVal vData As Integer)
Attribute iPadrao.VB_Description = "1 ->Filial  Fonrnecedor é padrão para esse Produto nessa Filial da Empresa     0 -> Não é."
    mvariPadrao = vData
End Property

Public Property Get iPadrao() As Integer
    iPadrao = mvariPadrao
End Property

Public Property Let sUMQuantUltimaCotacao(ByVal vData As String)
    mvarsUMQuantUltimaCotacao = vData
End Property

Public Property Get sUMQuantUltimaCotacao() As String
    sUMQuantUltimaCotacao = mvarsUMQuantUltimaCotacao
End Property

Public Property Let dQuantUltimaCotacao(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dQuantUltimaCotacao = 5
    mvardQuantUltimaCotacao = vData
End Property

Public Property Get dQuantUltimaCotacao() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dQuantUltimaCotacao
    dQuantUltimaCotacao = mvardQuantUltimaCotacao
End Property

Public Property Let iTipoFreteUltimaCotacao(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iTipoFreteUltimaCotacao = 5
    mvariTipoFreteUltimaCotacao = vData
End Property


Public Property Get iTipoFreteUltimaCotacao() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iTipoFreteUltimaCotacao
    iTipoFreteUltimaCotacao = mvariTipoFreteUltimaCotacao
End Property



Public Property Let dtDataUltimaCotacao(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtDataUltimaCotacao = 5
    mvardtDataUltimaCotacao = vData
End Property


Public Property Get dtDataUltimaCotacao() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtDataUltimaCotacao
    dtDataUltimaCotacao = mvardtDataUltimaCotacao
End Property



Public Property Let dUltimaCotacao(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dUltimaCotacao = 5
    mvardUltimaCotacao = vData
End Property


Public Property Get dUltimaCotacao() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dUltimaCotacao
    dUltimaCotacao = mvardUltimaCotacao
End Property



Public Property Let dPrecoTotal(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dPrecoTotal = 5
    mvardPrecoTotal = vData
End Property


Public Property Get dPrecoTotal() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dPrecoTotal
    dPrecoTotal = mvardPrecoTotal
End Property



Public Property Let dtDataReceb(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtDataReceb = 5
    mvardtDataReceb = vData
End Property


Public Property Get dtDataReceb() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtDataReceb
    dtDataReceb = mvardtDataReceb
End Property



Public Property Let dtDataPedido(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtDataPedido = 5
    mvardtDataPedido = vData
End Property


Public Property Get dtDataPedido() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtDataPedido
    dtDataPedido = mvardtDataPedido
End Property

Public Property Let sUMQuantRecebida(ByVal vData As String)
    mvarsUMQuantRecebida = vData
End Property

Public Property Get sUMQuantRecebida() As String
    sUMQuantRecebida = mvarsUMQuantRecebida
End Property

Public Property Let dQuantRecebida(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dQuantRecebida = 5
    mvardQuantRecebida = vData
End Property


Public Property Get dQuantRecebida() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dQuantRecebida
    dQuantRecebida = mvardQuantRecebida
End Property

Public Property Let sUMQuantPedida(ByVal vData As String)
    mvarsUMQuantPedida = vData
End Property

Public Property Get sUMQuantPedida() As String
    sUMQuantPedida = mvarsUMQuantPedida
End Property

Public Property Let dQuantPedida(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dQuantPedida = 5
    mvardQuantPedida = vData
End Property


Public Property Get dQuantPedida() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dQuantPedida
    dQuantPedida = mvardQuantPedida
End Property

Public Property Let dTempoRessup(ByVal vData As Double)
    mvardTempoRessup = vData
End Property

Public Property Get dTempoRessup() As Double
    dTempoRessup = mvardTempoRessup
End Property

Public Property Let dtDataUltimaCompra(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtDataUltimaCompra = 5
    mvardtDataUltimaCompra = vData
End Property


Public Property Get dtDataUltimaCompra() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtDataUltimaCompra
    dtDataUltimaCompra = mvardtDataUltimaCompra
End Property



Public Property Let dQuantPedAbertos(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dQuantPedAbertos = 5
    mvardQuantPedAbertos = vData
End Property


Public Property Get dQuantPedAbertos() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dQuantPedAbertos
    dQuantPedAbertos = mvardQuantPedAbertos
End Property



Public Property Let iNota(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iNota = 5
    mvariNota = vData
End Property


Public Property Get iNota() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iNota
    iNota = mvariNota
End Property



Public Property Let dLoteMinimo(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dLoteMinimo = 5
    mvardLoteMinimo = vData
End Property


Public Property Get dLoteMinimo() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dLoteMinimo
    dLoteMinimo = mvardLoteMinimo
End Property



Public Property Let sProdutoFornecedor(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sProdutoFornecedor = 5
    mvarsProdutoFornecedor = vData
End Property


Public Property Get sProdutoFornecedor() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sProdutoFornecedor
    sProdutoFornecedor = mvarsProdutoFornecedor
End Property



Public Property Let iFilialForn(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iFilialForn = 5
    mvariFilialForn = vData
End Property


Public Property Get iFilialForn() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iFilialForn
    iFilialForn = mvariFilialForn
End Property



Public Property Let lFornecedor(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lFornecedor = 5
    mvarlFornecedor = vData
End Property


Public Property Get lFornecedor() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lFornecedor
    lFornecedor = mvarlFornecedor
End Property



Public Property Let sProduto(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sProduto = 5
    mvarsProduto = vData
End Property


Public Property Get sProduto() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sProduto
    sProduto = mvarsProduto
End Property



Public Property Let iFilialEmpresa(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iFilialEmpresa = 5
    mvariFilialEmpresa = vData
End Property


Public Property Get iFilialEmpresa() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iFilialEmpresa
    iFilialEmpresa = mvariFilialEmpresa
End Property




VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelDemonstrativo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarsProduto As String 'local copy
Private mvariMes As Integer 'local copy
Private mvariAno As Integer 'local copy
Private mvariFilialEmpresa As Integer 'local copy
Private mvarsDescricao As String 'local copy
Private mvarsUMEstoque As String 'local copy
Private mvarsUMVenda As String 'local copy
Private mvariClasseUM As Integer 'local copy
Private mvardConsumoInterno As Double 'local copy
Private mvardVendaMedia As Double 'local copy
Private mvardProducaoAcumulada As Double 'local copy
Private mvardProducaoDiaria As Double 'local copy
Private mvardEstoqueAtual As Double 'local copy
Private mvardQuantidadeFaturada As Double 'local copy
Private mvardProgramacaoVendas As Double 'local copy
Private mvardProdAnterior As Double
Private mvardQuantidadePrevVenda As Double
Private mvarsCodigoPrevVenda As String 'local copy
Private mvardEstoqueNaoConforme As Double 'local copy
Private mvardEstoqueEmRecup As Double 'local copy
Private mvardEstoqueForaValidade As Double 'local copy
Private mvardEstoqueElaboracao As Double 'local copy
Private mvardQtdPedVenda As Double 'local copy

Public Property Let dQtdPedVenda(ByVal vData As Double)
    mvardQtdPedVenda = vData
End Property

Public Property Get dQtdPedVenda() As Double
    dQtdPedVenda = mvardQtdPedVenda
End Property

Public Property Let dEstoqueElaboracao(ByVal vData As Double)
    mvardEstoqueElaboracao = vData
End Property

Public Property Get dEstoqueElaboracao() As Double
    dEstoqueElaboracao = mvardEstoqueElaboracao
End Property

Public Property Let dEstoqueForaValidade(ByVal vData As Double)
    mvardEstoqueForaValidade = vData
End Property

Public Property Get dEstoqueForaValidade() As Double
    dEstoqueForaValidade = mvardEstoqueForaValidade
End Property

Public Property Let dEstoqueEmRecup(ByVal vData As Double)
    mvardEstoqueEmRecup = vData
End Property

Public Property Get dEstoqueEmRecup() As Double
    dEstoqueEmRecup = mvardEstoqueEmRecup
End Property

Public Property Let dEstoqueNaoConforme(ByVal vData As Double)
    mvardEstoqueNaoConforme = vData
End Property

Public Property Get dEstoqueNaoConforme() As Double
    dEstoqueNaoConforme = mvardEstoqueNaoConforme
End Property

Public Property Let sCodigoPrevVenda(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sDescricao = 5
    mvarsCodigoPrevVenda = vData
End Property


Public Property Get sCodigoPrevVenda() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sDescricao
    sCodigoPrevVenda = mvarsCodigoPrevVenda
End Property


Public Property Let dQuantidadePrevVenda(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dProgramacaoVendas = 5
    mvardQuantidadePrevVenda = vData
End Property


Public Property Get dQuantidadePrevVenda() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dProgramacaoVendas
    dQuantidadePrevVenda = mvardQuantidadePrevVenda
End Property


Public Property Let dProdAnterior(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dProgramacaoVendas = 5
    mvardProdAnterior = vData
End Property


Public Property Get dProdAnterior() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dProgramacaoVendas
    dProdAnterior = mvardProdAnterior
End Property


Public Property Let dProgramacaoVendas(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dProgramacaoVendas = 5
    mvardProgramacaoVendas = vData
End Property


Public Property Get dProgramacaoVendas() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dProgramacaoVendas
    dProgramacaoVendas = mvardProgramacaoVendas
End Property



Public Property Let dQuantidadeFaturada(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dQuantidadeFaturada = 5
    mvardQuantidadeFaturada = vData
End Property


Public Property Get dQuantidadeFaturada() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dQuantidadeFaturada
    dQuantidadeFaturada = mvardQuantidadeFaturada
End Property



Public Property Let dEstoqueAtual(ByVal vData As Double)
    mvardEstoqueAtual = vData
End Property

Public Property Get dEstoqueAtual() As Double
    dEstoqueAtual = mvardEstoqueAtual
End Property



Public Property Let dProducaoDiaria(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dProducaoDiaria = 5
    mvardProducaoDiaria = vData
End Property


Public Property Get dProducaoDiaria() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dProducaoDiaria
    dProducaoDiaria = mvardProducaoDiaria
End Property



Public Property Let dProducaoAcumulada(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dProducaoAcumulada = 5
    mvardProducaoAcumulada = vData
End Property


Public Property Get dProducaoAcumulada() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dProducaoAcumulada
    dProducaoAcumulada = mvardProducaoAcumulada
End Property



Public Property Let dVendaMedia(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dVendaMedia = 5
    mvardVendaMedia = vData
End Property


Public Property Get dVendaMedia() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dVendaMedia
    dVendaMedia = mvardVendaMedia
End Property



Public Property Let dConsumoInterno(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dConsumoInterno = 5
    mvardConsumoInterno = vData
End Property


Public Property Get dConsumoInterno() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dConsumoInterno
    dConsumoInterno = mvardConsumoInterno
End Property



Public Property Let sUMEstoque(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sUMEstoque = 5
    mvarsUMEstoque = vData
End Property


Public Property Get sUMEstoque() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sUMEstoque
    sUMEstoque = mvarsUMEstoque
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



Public Property Let iAno(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iAno = 5
    mvariAno = vData
End Property


Public Property Get iAno() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iAno
    iAno = mvariAno
End Property



Public Property Let iMes(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iMes = 5
    mvariMes = vData
End Property


Public Property Get iMes() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iMes
    iMes = mvariMes
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

Public Property Let sUMVenda(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
    mvarsUMVenda = vData
End Property


Public Property Get sUMVenda() As String
'used when retrieving value of a property, on the right side of an assignment.
    sUMVenda = mvarsUMVenda
End Property

Public Property Let iClasseUM(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
    mvariClasseUM = vData
End Property


Public Property Get iClasseUM() As Integer
'used when retrieving value of a property, on the right side of an assignment.
    iClasseUM = mvariClasseUM
End Property



VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassEvolucaoFatTela"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvardtDataDe As Date
Private mvardtDataAte As Date
Private mvariDetalhado As Integer
Private mvariFilialEmpresa As Integer
Private mvariExibeNumVendas As Integer
Private mvariExibeNumProdutos As Integer
Private mvariExibeValorVendas As Integer
Private mvarsTituloGrafico As String
Private mvarsNomeGrafico As String
Private mvarsNomePlanilha As String

Public Property Let dtDataDe(ByVal vData As Date)
    mvardtDataDe = vData
End Property

Public Property Get dtDataDe() As Date
    dtDataDe = mvardtDataDe
End Property

Public Property Let dtDataAte(ByVal vData As Date)
    mvardtDataAte = vData
End Property

Public Property Get dtDataAte() As Date
    dtDataAte = mvardtDataAte
End Property

Public Property Let iDetalhado(ByVal vData As Integer)
    mvariDetalhado = vData
End Property

Public Property Get iDetalhado() As Integer
    iDetalhado = mvariDetalhado
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let iExibeNumVendas(ByVal vData As Integer)
    mvariExibeNumVendas = vData
End Property

Public Property Get iExibeNumVendas() As Integer
    iExibeNumVendas = mvariExibeNumVendas
End Property

Public Property Let iExibeNumProdutos(ByVal vData As Integer)
    mvariExibeNumProdutos = vData
End Property

Public Property Get iExibeNumProdutos() As Integer
    iExibeNumProdutos = mvariExibeNumProdutos
End Property

Public Property Let iExibeValorVendas(ByVal vData As Integer)
    mvariExibeValorVendas = vData
End Property

Public Property Get iExibeValorVendas() As Integer
    iExibeValorVendas = mvariExibeValorVendas
End Property

Public Property Let sTituloGrafico(ByVal vData As String)
    mvarsTituloGrafico = vData
End Property

Public Property Get sTituloGrafico() As String
    sTituloGrafico = mvarsTituloGrafico
End Property

Public Property Let sNomeGrafico(ByVal vData As String)
    mvarsNomeGrafico = vData
End Property

Public Property Get sNomeGrafico() As String
    sNomeGrafico = mvarsNomeGrafico
End Property

Public Property Let sNomePlanilha(ByVal vData As String)
    mvarsNomePlanilha = vData
End Property

Public Property Get sNomePlanilha() As String
    sNomePlanilha = mvarsNomePlanilha
End Property

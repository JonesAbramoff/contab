VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFormacaoPreco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariFilialEmpresa As Integer
Private mvariEscopo As Integer
Private mvarsItemCategoria As String
Private mvarsProduto As String
Private mvariTabelaPreco As Integer
Private mvarlNumIntDoc As Long
Private mvarsExpressao As String
Private mvariLinha As Integer
Private mvarsTitulo As String
Private mvardValor As Double

Public lCliente As Long
Public iFilialCli As Integer

Public Property Let dValor(ByVal vData As Double)
    mvardValor = vData
End Property

Public Property Get dValor() As Double
    dValor = mvardValor
End Property

Public Property Let sTitulo(ByVal vData As String)
    mvarsTitulo = vData
End Property

Public Property Get sTitulo() As String
    sTitulo = mvarsTitulo
End Property

Public Property Let iLinha(ByVal vData As Integer)
    mvariLinha = vData
End Property

Public Property Get iLinha() As Integer
    iLinha = mvariLinha
End Property

Public Property Let sExpressao(ByVal vData As String)
    mvarsExpressao = vData
End Property

Public Property Get sExpressao() As String
    sExpressao = mvarsExpressao
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let iTabelaPreco(ByVal vData As Integer)
    mvariTabelaPreco = vData
End Property

Public Property Get iTabelaPreco() As Integer
    iTabelaPreco = mvariTabelaPreco
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let sItemCategoria(ByVal vData As String)
    mvarsItemCategoria = vData
End Property

Public Property Get sItemCategoria() As String
    sItemCategoria = mvarsItemCategoria
End Property

Public Property Let iEscopo(ByVal vData As Integer)
    mvariEscopo = vData
End Property

Public Property Get iEscopo() As Integer
    iEscopo = mvariEscopo
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property




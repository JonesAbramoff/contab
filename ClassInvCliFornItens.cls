VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassInvCliFornItens"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntInvCliForn As Long
Private mvariSeq As Integer
Private mvarsProduto As String
Private mvardQtdData As Double 'Quantidade total para o produto\escaninho\data quando o inventário foi feito
Private mvardQtdEncontCliData As Double 'Quantidade que o usuário encontrou no cliente(fornecedor)\filial para produto\escaninho\data
Private mvardQtdCliData As Double 'Quantidade que o sistema encontrou no cliente(fornecedor)\filial para produto\escaninho\data
Private mvardQtdAcerto As Double 'Quantidade ser se considerada na movimentação para acertos na distribuição entre terceiros
Private mvarcolDistribuicao As New Collection 'Não é gravado, dado auxiliar, usado para ter a relação de clientes\fornecedores que possuem o produto em estoque no escaninho\data
Private mvardQtdDistrib As Double 'Não é gravado, dado auxiliar, quantidade que o sistema encontrou para todos clientes\fornedores para produto\escaninho\data
Private mvarsObs As String

Public Property Let sObs(ByVal vData As String)
    mvarsObs = vData
End Property

Public Property Get sObs() As String
    sObs = mvarsObs
End Property

Public Property Let dQtdDistrib(ByVal vData As Double)
    mvardQtdDistrib = vData
End Property

Public Property Get dQtdDistrib() As Double
    dQtdDistrib = mvardQtdDistrib
End Property

Public Property Set colDistribuicao(ByVal vData As Object)
    Set mvarcolDistribuicao = vData
End Property

Public Property Get colDistribuicao() As Collection
    Set colDistribuicao = mvarcolDistribuicao
End Property

Public Property Let lNumIntInvCliForn(ByVal vData As Long)
    mvarlNumIntInvCliForn = vData
End Property

Public Property Get lNumIntInvCliForn() As Long
    lNumIntInvCliForn = mvarlNumIntInvCliForn
End Property

Public Property Let iSeq(ByVal vData As Integer)
    mvariSeq = vData
End Property

Public Property Get iSeq() As Integer
    iSeq = mvariSeq
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let dQtdData(ByVal vData As Double)
    mvardQtdData = vData
End Property

Public Property Get dQtdData() As Double
    dQtdData = mvardQtdData
End Property

Public Property Let dQtdEncontCliData(ByVal vData As Double)
    mvardQtdEncontCliData = vData
End Property

Public Property Get dQtdEncontCliData() As Double
    dQtdEncontCliData = mvardQtdEncontCliData
End Property

Public Property Let dQtdCliData(ByVal vData As Double)
    mvardQtdCliData = vData
End Property

Public Property Get dQtdCliData() As Double
    dQtdCliData = mvardQtdCliData
End Property

Public Property Let dQtdAcerto(ByVal vData As Double)
    mvardQtdAcerto = vData
End Property

Public Property Get dQtdAcerto() As Double
    dQtdAcerto = mvardQtdAcerto
End Property

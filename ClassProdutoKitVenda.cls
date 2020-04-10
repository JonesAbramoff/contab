VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
End
Attribute VB_Name = "ClassProdutoKitVenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level", "; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarsProdutoKit AS String
Private mvarsProduto AS String
Private mvariSeq AS Integer
Private mvardQuantidade AS Double
Private mvarsUM AS String

Public Property Let sProdutoKit (ByVal vData As String)
    mvarsProdutoKit = vData
End Property

Public Property Get sProdutoKit () AS String
    sProdutoKit= mvarsProdutoKit
End Property

Public Property Let sProduto (ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto () AS String
    sProduto= mvarsProduto
End Property

Public Property Let iSeq (ByVal vData As Integer)
    mvariSeq = vData
End Property

Public Property Get iSeq () AS Integer
    iSeq= mvariSeq
End Property

Public Property Let dQuantidade (ByVal vData As Double)
    mvardQuantidade = vData
End Property

Public Property Get dQuantidade () AS Double
    dQuantidade= mvardQuantidade
End Property

Public Property Let sUM (ByVal vData As String)
    mvarsUM = vData
End Property

Public Property Get sUM () AS String
    sUM= mvarsUM
End Property

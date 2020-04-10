VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassEstoque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private mvarlNumIntDocRastroMovto As Long
Private mvarlNumIntDocRastroLote As Long
Private mvarlNumIntDocApropInsProd As Long
Private mvardtData As Long 'guarda a data do movimento mais antigo, se a data for mais antiga que a data do ultimo movimento de estoque
Private mvarcolEstoqueMesProduto As New Collection 'colecao de ClassEstoqueMesProduto

Public Property Set colEstoqueMesProduto(ByVal vData As Object)
    Set mvarcolEstoqueMesProduto = vData
End Property

Public Property Get colEstoqueMesProduto() As Collection
    Set colEstoqueMesProduto = mvarcolEstoqueMesProduto
End Property

Public Property Let dtData(ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let lNumIntDocApropInsProd(ByVal vData As Long)
    mvarlNumIntDocApropInsProd = vData
End Property

Public Property Get lNumIntDocApropInsProd() As Long
    lNumIntDocApropInsProd = mvarlNumIntDocApropInsProd
End Property

Public Property Let lNumIntDocRastroLote(ByVal vData As Long)
    mvarlNumIntDocRastroLote = vData
End Property

Public Property Get lNumIntDocRastroLote() As Long
    lNumIntDocRastroLote = mvarlNumIntDocRastroLote
End Property

Public Property Let lNumIntDocRastroMovto(ByVal vData As Long)
    mvarlNumIntDocRastroMovto = vData
End Property

Public Property Get lNumIntDocRastroMovto() As Long
    lNumIntDocRastroMovto = mvarlNumIntDocRastroMovto
End Property

Private Sub Class_Initialize()
    mvardtData = DATA_NULA
End Sub

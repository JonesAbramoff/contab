VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassEmbalagensItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarlNumIntDocItem As Long
Private mvariTipoDocOrigem As Integer
Private mvariEmbalagem As Integer
Private mvardQuantEmb As Double
Private mvardQuantProduto As Double
Private mvarsProduto As String
Private mvariItem As Integer

Public Property Let dQuantProduto(ByVal vData As Double)
    mvardQuantProduto = vData
End Property

Public Property Get dQuantProduto() As Double
    dQuantProduto = mvardQuantProduto
End Property

Public Property Let dQuantEmb(ByVal vData As Double)
    mvardQuantEmb = vData
End Property

Public Property Get dQuantEmb() As Double
    dQuantEmb = mvardQuantEmb
End Property

Public Property Let iEmbalagem(ByVal vData As Integer)
    mvariEmbalagem = vData
End Property

Public Property Get iEmbalagem() As Integer
    iEmbalagem = mvariEmbalagem
End Property

Public Property Let lNumIntDocItem(ByVal vData As Long)
    mvarlNumIntDocItem = vData
End Property

Public Property Get lNumIntDocItem() As Long
    lNumIntDocItem = mvarlNumIntDocItem
End Property

Public Property Let iTipoDocOrigem(ByVal vData As Integer)
    mvariTipoDocOrigem = vData
End Property

Public Property Get iTipoDocOrigem() As Integer
    iTipoDocOrigem = mvariTipoDocOrigem
End Property

Public Property Let iItem(ByVal vData As Integer)
    mvariItem = vData
End Property

Public Property Get iItem() As Integer
    iItem = mvariItem
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property



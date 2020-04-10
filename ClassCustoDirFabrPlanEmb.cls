VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCustoDirFabrPlanEmb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvariFilialEmpresa As Integer
Private mvariAno As Integer
Private mvarsProduto As String
Private mvariSeq As Integer
Private mvariEmbalagem As Integer
Private mvarsProdutoEmb As String
Private mvardQtde As Double
Private mvardCustoUnitario As Double

Public Property Let dCustoUnitario(ByVal vData As Double)
    mvardCustoUnitario = vData
End Property

Public Property Get dCustoUnitario() As Double
    dCustoUnitario = mvardCustoUnitario
End Property

Public Property Let iEmbalagem(ByVal vData As Integer)
    mvariEmbalagem = vData
End Property

Public Property Get iEmbalagem() As Integer
    iEmbalagem = mvariEmbalagem
End Property

Public Property Let dQtde(ByVal vData As Double)
    mvardQtde = vData
End Property

Public Property Get dQtde() As Double
    dQtde = mvardQtde
End Property

Public Property Let sProdutoEmb(ByVal vData As String)
    mvarsProdutoEmb = vData
End Property

Public Property Get sProdutoEmb() As String
    sProdutoEmb = mvarsProdutoEmb
End Property

Public Property Let iSeq(ByVal vData As Integer)
    mvariSeq = vData
End Property

Public Property Get iSeq() As Integer
    iSeq = mvariSeq
End Property

Public Property Let iAno(ByVal vData As Integer)
    mvariAno = vData
End Property

Public Property Get iAno() As Integer
    iAno = mvariAno
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property






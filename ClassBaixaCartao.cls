VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassBaixaCartao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvardtDataDe As Date
Private mvardtDataAte As Date
Private mvariFilialEmpresa As Integer
Private mvariBandeira As Integer
Private mvarcolArq As New Collection
Private mvariTotalReg As Integer

Public Property Let iTotalReg(ByVal vData As Integer)
    mvariTotalReg = vData
End Property

Public Property Get iTotalReg() As Integer
    iTotalReg = mvariTotalReg
End Property

Public Property Set colArq(ByVal vData As Collection)
    Set mvarcolArq = vData
End Property

Public Property Get colArq() As Collection
    Set colArq = mvarcolArq
End Property

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

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let iBandeira(ByVal vData As Integer)
    mvariBandeira = vData
End Property

Public Property Get iBandeira() As Integer
    iBandeira = mvariBandeira
End Property

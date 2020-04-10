VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassOSApDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlCodigoAp As Long
Private mvariFilialEmpresa As Integer
Private mvariSeq As Integer
Private mvardQuantidade As Double
Private mvarsOS As String
Private mvarsProdutoOS As String
Private mvariTipo As Integer
Private mvarsUM As String
Private mvarlNumIntDoc As Long
Private mvarsProduto As String
Private mvarsProdutoDesc As String
Private mvarsTipo As String
Private mvarsDet As String
Private mvariMarcado As Integer
Private mvariItemNF As Integer
Private mvariClasseUM As Integer
Private mvariPeca As Integer
Private mvarlNumIntDocAP As Long
Private mvarobjSol As New ClassSolicSRV
Private mvarobjItemSol As New ClassItensSolicSRV

Public Property Set objItemSol(ByVal vData As Object)
    Set mvarobjItemSol = vData
End Property

Public Property Get objItemSol() As ClassItensSolicSRV
    Set objItemSol = mvarobjItemSol
End Property

Public Property Set objSol(ByVal vData As Object)
    Set mvarobjSol = vData
End Property

Public Property Get objSol() As ClassSolicSRV
    Set objSol = mvarobjSol
End Property

Public Property Let lNumIntDocAP(ByVal vData As Long)
    mvarlNumIntDocAP = vData
End Property

Public Property Get lNumIntDocAP() As Long
    lNumIntDocAP = mvarlNumIntDocAP
End Property

Public Property Let iPeca(ByVal vData As Integer)
    mvariPeca = vData
End Property

Public Property Get iPeca() As Integer
    iPeca = mvariPeca
End Property

Public Property Let iClasseUM(ByVal vData As Integer)
    mvariClasseUM = vData
End Property

Public Property Get iClasseUM() As Integer
    iClasseUM = mvariClasseUM
End Property

Public Property Let iItemNF(ByVal vData As Integer)
    mvariItemNF = vData
End Property

Public Property Get iItemNF() As Integer
    iItemNF = mvariItemNF
End Property

Public Property Let iMarcado(ByVal vData As Integer)
    mvariMarcado = vData
End Property

Public Property Get iMarcado() As Integer
    iMarcado = mvariMarcado
End Property

Public Property Let sDet(ByVal vData As String)
    mvarsDet = vData
End Property

Public Property Get sDet() As String
    sDet = mvarsDet
End Property

Public Property Let sTipo(ByVal vData As String)
    mvarsTipo = vData
End Property

Public Property Get sTipo() As String
    sTipo = mvarsTipo
End Property

Public Property Let sProdutoDesc(ByVal vData As String)
    mvarsProdutoDesc = vData
End Property

Public Property Get sProdutoDesc() As String
    sProdutoDesc = mvarsProdutoDesc
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let sUM(ByVal vData As String)
    mvarsUM = vData
End Property

Public Property Get sUM() As String
    sUM = mvarsUM
End Property

Public Property Let iTipo(ByVal vData As Integer)
    mvariTipo = vData
End Property

Public Property Get iTipo() As Integer
    iTipo = mvariTipo
End Property

Public Property Let lCodigoAp(ByVal vData As Long)
    mvarlCodigoAp = vData
End Property

Public Property Get lCodigoAp() As Long
    lCodigoAp = mvarlCodigoAp
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let iSeq(ByVal vData As Integer)
    mvariSeq = vData
End Property

Public Property Get iSeq() As Integer
    iSeq = mvariSeq
End Property

Public Property Let dQuantidade(ByVal vData As Double)
    mvardQuantidade = vData
End Property

Public Property Get dQuantidade() As Double
    dQuantidade = mvardQuantidade
End Property

Public Property Let sOS(ByVal vData As String)
    mvarsOS = vData
End Property

Public Property Get sOS() As String
    sOS = mvarsOS
End Property

Public Property Let sProdutoOS(ByVal vData As String)
    mvarsProdutoOS = vData
End Property

Public Property Get sProdutoOS() As String
    sProdutoOS = mvarsProdutoOS
End Property


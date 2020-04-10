VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelMvPerCtaRef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntRel As Long
Private mvariSeqCtaPer As Integer
Private mvariExercicio As Integer
Private mvariPeriodo As Integer
Private mvarsContaRef As String
Private mvarsContaRefMask As String
Private mvarsDescricao As String
Private mvardSaldoAnterior As Double
Private mvardCredito As Double
Private mvardDebito As Double
Private mvardSaldoPeriodo As Double
Private mvardSaldoAtual As Double
Private mvariFilialEmpresa As Integer
Private mvarcolDia As New Collection

Public Property Set colDia(ByVal vData As Object)
    Set mvarcolDia = vData
End Property

Public Property Get colDia() As Collection
    Set colDia = mvarcolDia
End Property

Public Property Let lNumIntRel(ByVal vData As Long)
    mvarlNumIntRel = vData
End Property

Public Property Get lNumIntRel() As Long
    lNumIntRel = mvarlNumIntRel
End Property

Public Property Let iSeqCtaPer(ByVal vData As Integer)
    mvariSeqCtaPer = vData
End Property

Public Property Get iSeqCtaPer() As Integer
    iSeqCtaPer = mvariSeqCtaPer
End Property

Public Property Let iExercicio(ByVal vData As Integer)
    mvariExercicio = vData
End Property

Public Property Get iExercicio() As Integer
    iExercicio = mvariExercicio
End Property

Public Property Let iPeriodo(ByVal vData As Integer)
    mvariPeriodo = vData
End Property

Public Property Get iPeriodo() As Integer
    iPeriodo = mvariPeriodo
End Property

Public Property Let sContaRef(ByVal vData As String)
    mvarsContaRef = vData
End Property

Public Property Get sContaRef() As String
    sContaRef = mvarsContaRef
End Property

Public Property Let sContaRefMask(ByVal vData As String)
    mvarsContaRefMask = vData
End Property

Public Property Get sContaRefMask() As String
    sContaRefMask = mvarsContaRefMask
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let dSaldoAnterior(ByVal vData As Double)
    mvardSaldoAnterior = vData
End Property

Public Property Get dSaldoAnterior() As Double
    dSaldoAnterior = mvardSaldoAnterior
End Property

Public Property Let dCredito(ByVal vData As Double)
    mvardCredito = vData
End Property

Public Property Get dCredito() As Double
    dCredito = mvardCredito
End Property

Public Property Let dDebito(ByVal vData As Double)
    mvardDebito = vData
End Property

Public Property Get dDebito() As Double
    dDebito = mvardDebito
End Property

Public Property Let dSaldoPeriodo(ByVal vData As Double)
    mvardSaldoPeriodo = vData
End Property

Public Property Get dSaldoPeriodo() As Double
    dSaldoPeriodo = mvardSaldoPeriodo
End Property

Public Property Let dSaldoAtual(ByVal vData As Double)
    mvardSaldoAtual = vData
End Property

Public Property Get dSaldoAtual() As Double
    dSaldoAtual = mvardSaldoAtual
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

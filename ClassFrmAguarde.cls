VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFrmAguarde"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariCancelar As Integer
Private mvardTempoInicial As Double
Private mvariTotalItens As Long
Private mvariItensProcessados As Long
Private mvardMediaTempoItem As Double
Private mvardTempoEstimado As Double
Private mvardPercConcluido As Double

Public Property Let dPercConcluido(ByVal vData As Double)
    mvardPercConcluido = vData
End Property

Public Property Get dPercConcluido() As Double
    dPercConcluido = mvardPercConcluido
End Property

Public Property Let dTempoEstimado(ByVal vData As Double)
    mvardTempoEstimado = vData
End Property

Public Property Get dTempoEstimado() As Double
    dTempoEstimado = mvardTempoEstimado
End Property

Public Property Let dMediaTempoItem(ByVal vData As Double)
    mvardMediaTempoItem = vData
End Property

Public Property Get dMediaTempoItem() As Double
    dMediaTempoItem = mvardMediaTempoItem
End Property

Public Property Let iItensProcessados(ByVal vData As Long)
    mvariItensProcessados = vData
End Property

Public Property Get iItensProcessados() As Long
    iItensProcessados = mvariItensProcessados
End Property

Public Property Let iTotalItens(ByVal vData As Long)
    mvariTotalItens = vData
End Property

Public Property Get iTotalItens() As Long
    iTotalItens = mvariTotalItens
End Property

Public Property Let dTempoInicial(ByVal vData As Double)
    mvardTempoInicial = vData
End Property

Public Property Get dTempoInicial() As Double
    dTempoInicial = mvardTempoInicial
End Property

Public Property Let iCancelar(ByVal vData As Integer)
    mvariCancelar = vData
End Property

Public Property Get iCancelar() As Integer
    iCancelar = mvariCancelar
End Property




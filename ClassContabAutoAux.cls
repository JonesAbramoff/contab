VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassContabAutoAux"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariFilial As Integer
Private mvariLote As Integer
Private mvarlProxDoc As Long 'numero do proximo Doc contabil a ser utilizado
Private mvarlComandoProxDoc As Long
Private mvarsCtaTransferencia As String 'conta contabil p/transferencia de uma filial da empresa (p/outras filiais)
Private mvariExercicio As Integer
Private mvariPeriodo As Integer

Public Property Let iPeriodo(ByVal vData As Integer)
    mvariPeriodo = vData
End Property

Public Property Get iPeriodo() As Integer
    iPeriodo = mvariPeriodo
End Property

Public Property Let iExercicio(ByVal vData As Integer)
    mvariExercicio = vData
End Property

Public Property Get iExercicio() As Integer
    iExercicio = mvariExercicio
End Property

Public Property Let sCtaTransferencia(ByVal vData As String)
    mvarsCtaTransferencia = vData
End Property

Public Property Get sCtaTransferencia() As String
    sCtaTransferencia = mvarsCtaTransferencia
End Property

Public Property Let lProxDoc(ByVal vData As Long)
    mvarlProxDoc = vData
End Property

Public Property Get lProxDoc() As Long
    lProxDoc = mvarlProxDoc
End Property

Public Property Let iLote(ByVal vData As Integer)
    mvariLote = vData
End Property

Public Property Get iLote() As Integer
    iLote = mvariLote
End Property

Public Property Let iFilial(ByVal vData As Integer)
    mvariFilial = vData
End Property

Public Property Get iFilial() As Integer
    iFilial = mvariFilial
End Property

Public Property Let lComandoProxDoc(ByVal vData As Long)
    mvarlComandoProxDoc = vData
End Property

Public Property Get lComandoProxDoc() As Long
    lComandoProxDoc = mvarlComandoProxDoc
End Property




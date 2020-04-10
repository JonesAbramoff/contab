VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTransferencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariCodContaOrigem As Integer
Private mvariCodContaDestino As Integer
Private mvariTipoMeioPagto As Integer
Private mvarlNumero As Long
Private mvardtData As Date
Private mvardValor As Double
Private mvarsHistorico As String
Private mvarsNumRefExterna As String
Private mvarlNumIntDoc As String
Private mvariFavorecido As Integer
'local variable(s) to hold property value(s)
Private mvarlSeqDestino As Long 'local copy
Private mvarlSeqOrigem As Long 'local copy
Public Property Let lSeqOrigem(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lSeqOrigem = 5
    mvarlSeqOrigem = vData
End Property


Public Property Get lSeqOrigem() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lSeqOrigem
    lSeqOrigem = mvarlSeqOrigem
End Property



Public Property Let lSeqDestino(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lSeqDestino = 5
    mvarlSeqDestino = vData
End Property


Public Property Get lSeqDestino() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lSeqDestino
    lSeqDestino = mvarlSeqDestino
End Property




Public Property Let iFavorecido(ByVal vData As Integer)
    mvariFavorecido = vData
End Property

Public Property Get iFavorecido() As Integer
    iFavorecido = mvariFavorecido
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let sNumRefExterna(ByVal vData As String)
    mvarsNumRefExterna = vData
End Property

Public Property Get sNumRefExterna() As String
    sNumRefExterna = mvarsNumRefExterna
End Property

Public Property Let sHistorico(ByVal vData As String)
    mvarsHistorico = vData
End Property

Public Property Get sHistorico() As String
    sHistorico = mvarsHistorico
End Property

Public Property Let dValor(ByVal vData As Double)
    mvardValor = vData
End Property

Public Property Get dValor() As Double
    dValor = mvardValor
End Property

Public Property Let dtData(ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let lNumero(ByVal vData As Long)
    mvarlNumero = vData
End Property

Public Property Get lNumero() As Long
    lNumero = mvarlNumero
End Property

Public Property Let iTipoMeioPagto(ByVal vData As Integer)
    mvariTipoMeioPagto = vData
End Property

Public Property Get iTipoMeioPagto() As Integer
    iTipoMeioPagto = mvariTipoMeioPagto
End Property

Public Property Let iCodContaOrigem(ByVal vData As Integer)
    mvariCodContaOrigem = vData
End Property

Public Property Get iCodContaOrigem() As Integer
    iCodContaOrigem = mvariCodContaOrigem
End Property

Public Property Let iCodContaDestino(ByVal vData As Integer)
    mvariCodContaDestino = vData
End Property

Public Property Get iCodContaDestino() As Integer
    iCodContaDestino = mvariCodContaDestino
End Property


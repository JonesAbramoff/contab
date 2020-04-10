VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassLancamento_Cabecalho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public sOrigem As String
Public iExercicio As Integer
Public iPeriodoLan As Integer
Public lDoc As Long
Public iLote As Integer
Public iPeriodoLote As Integer
Public dtData As Date
Public iFilialEmpresa As Integer
Public iCustoRealProducao As Integer
Public lNumIntDoc As Long
Public iTransacao As Integer
Public sModelo As String 'modelo de contabilização associado aos lançamentos
Public iSubTipo As Integer

Private Sub Class_Initialize()
    sOrigem = String(STRING_ORIGEM, 0)
End Sub

Public Sub Limpa()
    
    iLote = 0
    iExercicio = 0
    sOrigem = ""
    iPeriodoLan = 0
    iPeriodoLote = 0
    lDoc = 0
    iSubTipo = 0
End Sub


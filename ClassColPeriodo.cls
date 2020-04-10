VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassColPeriodo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private colPeriodo As Collection


Public Property Get Item(vntIndexKey As Integer) As ClassPeriodo
  Set Item = colPeriodo(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = colPeriodo.Count
End Property

Public Function NomeExterno(iExercicio As Integer, iPeriodo As Integer) As String
'recebe o código do Exercicio e Periodo e retorna o nome externo
'retorna string vazia se não achou o periodo fornecido

Dim objPeriodo As ClassPeriodo
Dim iAchou As Integer
    
    iAchou = 0
    
    For Each objPeriodo In colPeriodo
    
        If objPeriodo.iExercicio = iExercicio And objPeriodo.iPeriodo = iPeriodo Then
            NomeExterno = objPeriodo.sNomeExterno
            iAchou = 1
            Exit For
        End If
    Next
    
    If iAchou = 0 Then
        NomeExterno = ""
    End If
    
End Function


Private Sub Class_Initialize()
'faz a carga da coleção com todos os periodos de todos os exercicios

    Set colPeriodo = New Collection
    
    Call CF("Periodo_Le_Todos", colPeriodo)
    
End Sub


Private Sub Class_Terminate()
    'destroi a colecao quando essa class termina
    Set colPeriodo = Nothing
End Sub



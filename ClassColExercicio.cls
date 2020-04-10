VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassColExercicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private colExercicio As Collection


Public Property Get Item(vntIndexKey As Integer) As ClassExercicio
  Set Item = colExercicio(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = colExercicio.Count
End Property
Public Function Exercicio(sNomeExterno As String) As Long
'recebe o nome externo de um exercicio e retorna o codigo do Exercicio (iExercicio)
'retorna 0 se não achou um codigo para o nome externo fornecido

Dim objExercicio As ClassExercicio
Dim iAchou As Integer
    
    iAchou = 0
    
    For Each objExercicio In colExercicio
    
        If objExercicio.sNomeExterno = sNomeExterno Then
            Exercicio = objExercicio.iExercicio
            iAchou = 1
            Exit For
        End If
    Next
    
    If iAchou = 0 Then
        Exercicio = 0
    End If
    
End Function

Public Function NomeExterno(iExercicio As Integer) As String
'recebe o código do Exercicio e retorna o nome externo
'retorna string vazia se não achou o exercicio fornecido

Dim objExercicio As ClassExercicio
Dim iAchou As Integer
    
    iAchou = 0
    
    For Each objExercicio In colExercicio
    
        If objExercicio.iExercicio = iExercicio Then
            NomeExterno = objExercicio.sNomeExterno
            iAchou = 1
            Exit For
        End If
    Next
    
    If iAchou = 0 Then
        NomeExterno = ""
    End If
    
End Function

Private Sub Class_Initialize()
'faz a carga da coleção com todos os exercicios

    Set colExercicio = New Collection
    
    Call CF("Exercicios_Le_Todos", colExercicio)
    
End Sub

Private Sub Class_Terminate()
    'destroi a colecao quando essa class termina
    Set colExercicio = Nothing
End Sub


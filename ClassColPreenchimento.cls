VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassColPreenchimento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private colPreenchimento As Collection

Public Property Get Item(vntIndexKey As Integer) As ClassPreenchimento
  Set Item = colPreenchimento(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = colPreenchimento.Count
End Property

Public Function Preenchimento(sDescricao As String) As Long
'recebe a descri��o do preenchimento e retorna o codigo do preenchimento
'retorna 0 se n�o achou um codigo de preenchimento para a descri��o fornecida

Dim objPreenchimento As ClassPreenchimento
Dim iAchou As Integer
    
    iAchou = 0
    
    For Each objPreenchimento In colPreenchimento
    
        If objPreenchimento.sDescricao = sDescricao Then
            Preenchimento = objPreenchimento.iPreenchimento
            iAchou = 1
            Exit For
        End If
    Next
    
    If iAchou = 0 Then
        Preenchimento = 0
    End If
    
End Function

Public Function Descricao(iPreenchimento As Integer) As String
'recebe o codigo do preenchimento e retorna a descri��o
'retorna string vazia se n�o achou uma descri��o para o c�digo fornecido

Dim objPreenchimento As ClassPreenchimento
Dim iAchou As Integer
    
    iAchou = 0
    
    For Each objPreenchimento In colPreenchimento
    
        If objPreenchimento.iPreenchimento = iPreenchimento Then
            Descricao = objPreenchimento.sDescricao
            iAchou = 1
            Exit For
        End If
    Next
    
    If iAchou = 0 Then
        Descricao = ""
    End If
    
End Function

Private Sub Class_Initialize()
'faz a carga da cole��o com os preenchimentos

Dim objPreenchimento As ClassPreenchimento

    Set colPreenchimento = New Collection
    Set objPreenchimento = New ClassPreenchimento
        
    objPreenchimento.iPreenchimento = ZEROS_ESPACOS
    objPreenchimento.sDescricao = "Num.-Zeros � Esq. Alfa-Espa�os � Dir."
        
    'armazena em mem�ria
    colPreenchimento.Add objPreenchimento
        
    Set objPreenchimento = New ClassPreenchimento
    
    objPreenchimento.iPreenchimento = ESPACOS
    objPreenchimento.sDescricao = "Espa�os"
        
    'armazena em mem�ria
    colPreenchimento.Add objPreenchimento
    
    Set objPreenchimento = New ClassPreenchimento
    
    objPreenchimento.iPreenchimento = PREENCH_LIMPA_BRANCOS
    objPreenchimento.sDescricao = "Num.-Sem zeros � Esq. Alfa- Sem Espa�os � Dir."
        
    'armazena em mem�ria
    colPreenchimento.Add objPreenchimento
        
End Sub


Private Sub Class_Terminate()
    'destroi a colecao quando essa classe termina
    Set colPreenchimento = Nothing
End Sub





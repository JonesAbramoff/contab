VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassColTipoConta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private colTipoConta As Collection


Public Property Get Item(vntIndexKey As Integer) As ClassTipoConta
  Set Item = colTipoConta(vntIndexKey)
End Property



Public Property Get Count() As Long
    Count = colTipoConta.Count
End Property
Public Function TipoConta(sDescricao As String) As Long
'recebe a descrição do tipo da conta e retorna o tipo da conta
'retorna 0 se não achou um tipo de conta para a descrição fornecida

    Dim objTipoConta As ClassTipoConta
    Dim iAchou As Integer
    
    iAchou = 0
    
    For Each objTipoConta In colTipoConta
    
        If objTipoConta.sDescricao = sDescricao Then
            TipoConta = objTipoConta.iTipo
            iAchou = 1
            Exit For
        End If
    Next
    
    If iAchou = 0 Then
        TipoConta = 0
    End If
    
End Function

Public Function Descricao(iTipo As Integer) As String
'recebe o tipo da conta e retorna a descrição
'retorna string vazia se não achou uma descrição para o código fornecido

    Dim objTipoConta As ClassTipoConta
    Dim iAchou As Integer
    
    iAchou = 0
    
    For Each objTipoConta In colTipoConta
    
        If objTipoConta.iTipo = iTipo Then
            Descricao = objTipoConta.sDescricao
            iAchou = 1
            Exit For
        End If
    Next
    
    If iAchou = 0 Then
        Descricao = ""
    End If
    
End Function



Private Sub Class_Initialize()
'faz a carga da coleção com os tipos de conta

Dim objTipoConta As ClassTipoConta

    Set colTipoConta = New Collection
    Set objTipoConta = New ClassTipoConta
        
    objTipoConta.iTipo = CONTA_SINTETICA
    objTipoConta.sDescricao = "Sintética"
        
    'armazena em memória
    colTipoConta.Add objTipoConta
        
    Set objTipoConta = New ClassTipoConta
    
    objTipoConta.iTipo = CONTA_ANALITICA
    objTipoConta.sDescricao = "Analítica"
        
    'armazena em memória
    colTipoConta.Add objTipoConta
    
    
End Sub


Private Sub Class_Terminate()
    'destroi a colecao quando essa class termina
    Set colTipoConta = Nothing
End Sub




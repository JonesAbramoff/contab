VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassColNaturezaConta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private colNaturezaConta As Collection


Public Property Get Item(vntIndexKey As Integer) As ClassNaturezaConta
  Set Item = colNaturezaConta(vntIndexKey)
End Property



Public Property Get Count() As Long
    Count = colNaturezaConta.Count
End Property
Public Function NaturezaConta(sDescricao As String) As Long
'recebe a descrição da natureza da conta e retorna o codigo da natureza
'retorna 0 se não achou um codigo para a descrição fornecida

    Dim objNaturezaConta As ClassNaturezaConta
    Dim iAchou As Integer
    
    iAchou = 0
    
    For Each objNaturezaConta In colNaturezaConta
    
        If objNaturezaConta.sDescricao = sDescricao Then
            NaturezaConta = objNaturezaConta.iNatureza
            iAchou = 1
            Exit For
        End If
    Next
    
    If iAchou = 0 Then
        NaturezaConta = 0
    End If
    
End Function

Public Function Descricao(iNatureza As Integer) As String
'recebe o código da natureza da conta e retorna a descrição
'retorna string vazia se não achou uma descrição para o código fornecido

    Dim objNaturezaConta As ClassNaturezaConta
    Dim iAchou As Integer
    
    iAchou = 0
    
    For Each objNaturezaConta In colNaturezaConta
    
        If objNaturezaConta.iNatureza = iNatureza Then
            Descricao = objNaturezaConta.sDescricao
            iAchou = 1
            Exit For
        End If
    Next
    
    If iAchou = 0 Then
        Descricao = ""
    End If
    
End Function



Private Sub Class_Initialize()
'faz a carga da coleção com as naturezas de conta

Dim objNaturezaConta As ClassNaturezaConta

    Set colNaturezaConta = New Collection
    Set objNaturezaConta = New ClassNaturezaConta
        
    objNaturezaConta.iNatureza = CONTA_CREDITO
    objNaturezaConta.sDescricao = "Crédito"
        
    'armazena em memória
    colNaturezaConta.Add objNaturezaConta
        
        
    Set objNaturezaConta = New ClassNaturezaConta
        
    objNaturezaConta.iNatureza = CONTA_DEBITO
    objNaturezaConta.sDescricao = "Débito"
        
    'armazena em memória
    colNaturezaConta.Add objNaturezaConta
    
    
End Sub


Private Sub Class_Terminate()
    'destroi a colecao quando essa class termina
    Set colNaturezaConta = Nothing
End Sub





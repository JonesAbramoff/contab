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
'recebe a descri��o da natureza da conta e retorna o codigo da natureza
'retorna 0 se n�o achou um codigo para a descri��o fornecida

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
'recebe o c�digo da natureza da conta e retorna a descri��o
'retorna string vazia se n�o achou uma descri��o para o c�digo fornecido

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
'faz a carga da cole��o com as naturezas de conta

Dim objNaturezaConta As ClassNaturezaConta

    Set colNaturezaConta = New Collection
    Set objNaturezaConta = New ClassNaturezaConta
        
    objNaturezaConta.iNatureza = CONTA_CREDITO
    objNaturezaConta.sDescricao = "Cr�dito"
        
    'armazena em mem�ria
    colNaturezaConta.Add objNaturezaConta
        
        
    Set objNaturezaConta = New ClassNaturezaConta
        
    objNaturezaConta.iNatureza = CONTA_DEBITO
    objNaturezaConta.sDescricao = "D�bito"
        
    'armazena em mem�ria
    colNaturezaConta.Add objNaturezaConta
    
    
End Sub


Private Sub Class_Terminate()
    'destroi a colecao quando essa class termina
    Set colNaturezaConta = Nothing
End Sub





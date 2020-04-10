VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassColOrigemContab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private colOrigem As Collection

Public Property Get Item(vntIndexKey As Integer) As ClassOrigemContab
    Set Item = colOrigem(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = colOrigem.Count
End Property

Public Function Origem(sDescricao As String) As String
'recebe a descrição da origem e retorna o código da origem
'retorna string vazia se não achou um código para a descrição fornecida

Dim objOrigem As ClassOrigemContab
Dim iAchou As Integer
    
    iAchou = 0
    
    For Each objOrigem In colOrigem
    
        If objOrigem.sDescricao = sDescricao Then
            Origem = objOrigem.sOrigem
            iAchou = 1
            Exit For
        End If
    Next
    
    If iAchou = 0 Then
        Origem = ""
    End If
    
End Function

Public Function Descricao(sOrigem As String) As String
'recebe o codigo da origem e retorna a descrição da origem
'retorna string vazia se não achou uma descrição para o código fornecido

Dim objOrigem As ClassOrigemContab
Dim iAchou As Integer
    
    iAchou = 0
    
    For Each objOrigem In colOrigem
    
        If objOrigem.sOrigem = sOrigem Then
            Descricao = objOrigem.sDescricao
            iAchou = 1
            Exit For
        End If
    Next
    
    If iAchou = 0 Then
        Descricao = ""
    End If
    
End Function

Private Sub Class_Initialize()

Dim lComando As Long
Dim lErro As Long
Dim iFim_de_Arquivo As Integer
Dim sOrigem As String
Dim sDescricao As String
Dim objOrigem As ClassOrigemContab

On Error GoTo Erro_Init_ClassColOrigemContab

    Set colOrigem = New Collection
    
    lErro = CF("Origem_Le_Todos", colOrigem)
    If lErro <> SUCESSO Then Error 9275
    
    Exit Sub
    
Erro_Init_ClassColOrigemContab:

    Select Case Err
    
        Case 9275
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147428)
        
    End Select

    Exit Sub
    
End Sub

Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set colOrigem = Nothing
End Sub



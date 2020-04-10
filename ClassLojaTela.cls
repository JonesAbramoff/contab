VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassLojaTela"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True


Function ComissoesCalcula_Cria(X As Object) As Long

Dim lErro As Long

On Error GoTo Erro_ComissoesCalcula_Cria

    Set X = New ComissoesCalcula
            
    ComissoesCalcula_Cria = SUCESSO
            
    Exit Function
    
Erro_ComissoesCalcula_Cria:

    ComissoesCalcula_Cria = gErr

    Select Case gErr
    
                   
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142281)
        
    End Select
        
    Exit Function

End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassLayoutDuplicata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function Obtem_ViaTransp_Descricao_Rel(sViaTransp As String, iViaTransp As Integer) As Long

Dim lErro As Long

On Error GoTo Erro_Obtem_ViaTransp_Descricao_Rel

    lErro = CF("Obtem_ViaTransp_Descricao", sViaTransp, iViaTransp)
    If lErro <> SUCESSO Then gError 94911
    
    Obtem_ViaTransp_Descricao_Rel = SUCESSO
    
    Exit Function

Erro_Obtem_ViaTransp_Descricao_Rel:
    
    Obtem_ViaTransp_Descricao_Rel = gErr
    
    Select Case gErr
        
        Case 94911
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150907)
            
    End Select
    
    Exit Function
    
End Function

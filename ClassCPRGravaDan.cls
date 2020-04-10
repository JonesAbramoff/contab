VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCPRGravaDan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function AntecipPag_Testa_Gravacao(ByVal objAntecipPag As ClassAntecipPag) As Long

Dim lErro As Long
Dim objUsuario As New ClassDicUsuario
Dim objAntecipPagBD As New ClassAntecipPag

On Error GoTo Erro_AntecipPag_Testa_Gravacao

    If objAntecipPag.lNumPedCompra = 0 Then gError 189409
    
    objAntecipPagBD.lNumMovto = objAntecipPag.lNumMovto
    
    lErro = CF("AntecipPag_Le_NumMovto", objAntecipPagBD)
    If lErro <> SUCESSO And lErro <> 42845 Then gError 189410
    
    If lErro = SUCESSO Then
    
        objUsuario.sCodUsuario = gsUsuario
    
        lErro = DicUsuario_Le(objUsuario)
        If lErro <> SUCESSO Then gError 189411
        
        If objAntecipPagBD.lNumPedCompra <> objAntecipPag.lNumPedCompra Then
        
            If UCase(objUsuario.sCodGrupo) <> "SUPERVISOR" Then gError 189412
        
        End If
    
    End If

    AntecipPag_Testa_Gravacao = SUCESSO

    Exit Function

Erro_AntecipPag_Testa_Gravacao:

    AntecipPag_Testa_Gravacao = gErr

    Select Case gErr
    
        Case 189409 'ERRO_NUMERO_PEDIDO_NAO_PREENCHIDO
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_PEDIDO_NAO_PREENCHIDO", gErr)
        
        Case 189410, 189411
        
        Case 189412
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_PEDIDO_NAO_ALTERAVEL", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 189408)

    End Select


    Exit Function

End Function


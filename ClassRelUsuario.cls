VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private sUltUsuario As String
Private iLigado_Salva As Integer

Public Function Obter_LigadoRegiaoVenda(iLigado As Integer) As Long
'Para obter se o usuario gsUsuario participa como Usuario cobrador de alguma regiao de venda

Dim lErro As Long

On Error GoTo Erro_Obter_LigadoRegiaoVenda
    
    If gsUsuario <> sUltUsuario Then
        
        lErro = CF("RegiaoVenda_Le_UsuarioCobrador", gsUsuario, iLigado_Salva)
        If lErro <> SUCESSO Then gError 210628
    
        sUltUsuario = gsUsuario
    
    End If
    
    iLigado = iLigado_Salva
    
    Obter_LigadoRegiaoVenda = SUCESSO
    
    Exit Function
    
Erro_Obter_LigadoRegiaoVenda:

    Obter_LigadoRegiaoVenda = gErr
    
    Select Case gErr

        Case 210628
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 210629)

    End Select

    Exit Function

End Function



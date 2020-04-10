VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassChamaFuncoesLoja"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function LojaConfig_Le(objLoja As ClassLoja) As Long
'Lê a alcada a partir do código do usuário.
'Devolve os dados em objLoja.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_LojaConfig_Le
    
    
    lErro = Rotina_ObterObjetoFuncao("LJ_LOJACONFIG_LE", X)
    If lErro <> SUCESSO Then gError 62822
    
    LojaConfig_Le = X.LojaConfig_Le(objLoja)

    Exit Function
    
Erro_LojaConfig_Le:
    
    LojaConfig_Le = gErr
    
    Select Case gErr

        Case 62822
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145895)
        
    End Select
    
    Exit Function

End Function


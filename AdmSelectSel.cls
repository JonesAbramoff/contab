VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AdmSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function FilialEmpresaAlmox_Customiza(iFilialEmpresa As Integer) As Long

Dim lErro As Long

On Error GoTo Erro_FilialEmpresaAlmox_Customiza

    iFilialEmpresa = 1

    FilialEmpresaAlmox_Customiza = SUCESSO
    
    Exit Function
    
Erro_FilialEmpresaAlmox_Customiza:

    FilialEmpresaAlmox_Customiza = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142613)
        
    End Select
    
    Exit Function

End Function
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTBordCobrX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get OptAceite(ByVal objCT As Object, ByVal Index As Integer) As Object
     Set OptAceite = objCT.objUserControl.Controls("OptAceite")(Index)
End Property

Function BC_Traz_Bordero_Tela(ByVal objCT As Object, ByVal objBorderoCobrancaEmissao As ClassBorderoCobrancaEmissao) As Long

Dim lErro As Long

On Error GoTo Erro_BC_Traz_Bordero_Tela

    OptAceite(objCT, objBorderoCobrancaEmissao.iTipoAceite) = True
    
    BC_Traz_Bordero_Tela = SUCESSO
     
    Exit Function
    
Erro_BC_Traz_Bordero_Tela:

    BC_Traz_Bordero_Tela = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179125)
     
    End Select
     
    Exit Function
    
End Function

Function BC_Move_Bordero_Memoria(ByVal objCT As Object, ByVal objBorderoCobrancaEmissao As ClassBorderoCobrancaEmissao) As Long

Dim lErro As Long

On Error GoTo Erro_BC_Move_Bordero_Memoria

    If OptAceite(objCT, TIPO_ACEITE_AMBOS) = True Then
        objBorderoCobrancaEmissao.iTipoAceite = TIPO_ACEITE_AMBOS
    ElseIf OptAceite(objCT, TIPO_ACEITE_COM) = True Then
        objBorderoCobrancaEmissao.iTipoAceite = TIPO_ACEITE_COM
    ElseIf OptAceite(objCT, TIPO_ACEITE_SEM) = True Then
        objBorderoCobrancaEmissao.iTipoAceite = TIPO_ACEITE_SEM
    Else
        gError 140292
    End If
    
    BC_Move_Bordero_Memoria = SUCESSO
     
    Exit Function
    
Erro_BC_Move_Bordero_Memoria:

    BC_Move_Bordero_Memoria = gErr
     
    Select Case gErr
    
        Case 140292
            Call Rotina_Erro(vbOKOnly, "ERRO_ACEITE_NAO_PREENCHIDO", gErr)
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179126)
     
    End Select
     
    Exit Function
    
End Function



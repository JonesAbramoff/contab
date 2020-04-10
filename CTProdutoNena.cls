VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTProdutoNena"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Property Get QtdePorEmb(ByVal objCT As Object) As Object
     Set QtdePorEmb = objCT.objUserControl.Controls("QtdePorEmb")
End Property

Function Prod_TrazCarac_Tela(ByVal objCT As Object, ByVal objProduto As ClassProduto) As Long
'em Traz_TabCaracteristicas_Tela

Dim lErro As Long

On Error GoTo Erro_Prod_TrazCarac_Tela

    QtdePorEmb(objCT).PromptInclude = False
    QtdePorEmb(objCT).Text = objProduto.dQtdeEmbBase
    QtdePorEmb(objCT).PromptInclude = True

    Prod_TrazCarac_Tela = SUCESSO
    
    Exit Function
    
Erro_Prod_TrazCarac_Tela:

    Prod_TrazCarac_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184271)

    End Select
    
    Exit Function

End Function

Function Prod_MoveCaracMemoria(ByVal objCT As Object, ByVal objProduto As ClassProduto) As Long
'em Move_TabCaracteristicas_Memoria

Dim lErro As Long

On Error GoTo Erro_Prod_MoveCaracMemoria

    objProduto.dQtdeEmbBase = StrParaDbl(QtdePorEmb(objCT).Text)
    
    Prod_MoveCaracMemoria = SUCESSO
    
    Exit Function
    
Erro_Prod_MoveCaracMemoria:

    Prod_MoveCaracMemoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184272)

    End Select
    
    Exit Function

End Function

Sub QtdePorEmb_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Sub QtdePorEmb_GotFocus(ByVal objCT As Object)
    Call MaskEdBox_TrataGotFocus(QtdePorEmb(objCT), objCT.iAlterado)
End Sub

Sub QtdePorEmb_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_QtdePorEmb_Validate

    'Veifica se Previsao está preenchida
    If Len(Trim(QtdePorEmb(objCT).Text)) <> 0 Then

       'Critica a Previsao
       lErro = Valor_Positivo_Critica(QtdePorEmb(objCT).Text)
       If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    End If
    
    Exit Sub

Erro_QtdePorEmb_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177553)

    End Select

End Sub

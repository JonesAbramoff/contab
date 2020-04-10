VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTNFPAGHic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'--- inicio dos properties get dos controles da tela
Public Property Get Situacao(ByVal objCT As Object) As Object
     Set Situacao = objCT.objUserControl.Controls("Situacao")
End Property

Public Property Get Historico(ByVal objCT As Object) As Object
     Set Historico = objCT.objUserControl.Controls("Historico")
End Property

Public Property Get SubConta(ByVal objCT As Object) As Object
     Set SubConta = objCT.objUserControl.Controls("SubConta")
End Property

Public Property Get Observacao(ByVal objCT As Object) As Object
     Set Observacao = objCT.objUserControl.Controls("Observacao")
End Property

Public Property Get Status(ByVal objCT As Object) As Object
     Set Status = objCT.objUserControl.Controls("Status")
End Property

Function NFPAG_Traz_NFPag_Tela(ByVal objCT As Object, objNFPag As ClassNFsPag) As Long

Dim lErro As Long

On Error GoTo Erro_NFPAG_Traz_NFPag_Tela
   
    Call Combo_Seleciona_ItemData(Situacao(objCT), objNFPag.objInfoUsu.lSituacao)
    Call Combo_Seleciona_ItemData(Historico(objCT), objNFPag.objInfoUsu.lHistorico)
    Call Combo_Seleciona_ItemData(SubConta(objCT), objNFPag.objInfoUsu.lSubConta)

    Observacao(objCT).Text = objNFPag.objInfoUsu.sObservacao
    
    If objNFPag.iStatus = 0 Then
        Status(objCT).Caption = ""
    Else
        If objNFPag.iStatus = STATUS_BAIXADO Then
            Status(objCT).Caption = objNFPag.iStatus & SEPARADOR & STRING_STATUS_BAIXADO
        Else
            Status(objCT).Caption = objNFPag.iStatus & SEPARADOR & "NORMAL"
        End If
    End If
    
    NFPAG_Traz_NFPag_Tela = SUCESSO
     
    Exit Function
    
Erro_NFPAG_Traz_NFPag_Tela:

    NFPAG_Traz_NFPag_Tela = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179309)
     
    End Select
     
    Exit Function

End Function

Function NFPAG_Limpa_Tela(ByVal objCT As Object) As Long

Dim lErro As Long

On Error GoTo Erro_NFPAG_Limpa_Tela

    Situacao(objCT).ListIndex = -1
    Historico(objCT).ListIndex = -1
    SubConta(objCT).ListIndex = -1
    
    NFPAG_Limpa_Tela = SUCESSO
     
    Exit Function
    
Erro_NFPAG_Limpa_Tela:

    NFPAG_Limpa_Tela = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179310)
     
    End Select
     
    Exit Function

End Function

Function NFPAG_Move_Tela_Memoria(ByVal objCT As Object, objNFPag As ClassNFsPag) As Long

Dim lErro As Long
Dim objNFPagTitPag As New ClassNfPagTitPagHic

On Error GoTo Erro_NFPAG_Move_Tela_Memoria

    objNFPagTitPag.sObservacao = Observacao(objCT).Text
    
    If Situacao(objCT).ListIndex <> -1 Then
        objNFPagTitPag.lSituacao = Situacao(objCT).ItemData(Situacao(objCT).ListIndex)
    End If
    
    If Historico(objCT).ListIndex <> -1 Then
        objNFPagTitPag.lHistorico = Historico(objCT).ItemData(Historico(objCT).ListIndex)
    End If
    
    If SubConta(objCT).ListIndex <> -1 Then
        objNFPagTitPag.lSubConta = SubConta(objCT).ItemData(SubConta(objCT).ListIndex)
    End If
    
    Set objNFPag.objInfoUsu = objNFPagTitPag
        
    NFPAG_Move_Tela_Memoria = SUCESSO
     
    Exit Function
    
Erro_NFPAG_Move_Tela_Memoria:

    NFPAG_Move_Tela_Memoria = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179311)
     
    End Select
     
    Exit Function

End Function

Private Function Carrega_ComboSituacao(ByVal objCT As Object) As Long
'Carrega a combo de Situacao

Dim lErro As Long

On Error GoTo Erro_Carrega_ComboSituacao

    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_SITUACAO, Situacao(objCT), True, False, True)
    If lErro <> SUCESSO Then gError 132228

    Carrega_ComboSituacao = SUCESSO

    Exit Function

Erro_Carrega_ComboSituacao:

    Carrega_ComboSituacao = gErr

    Select Case gErr
    
        Case 132228

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179312)

    End Select

    Exit Function

End Function

Private Function Carrega_ComboHistorico(ByVal objCT As Object) As Long
'Carrega a combo de Historico

Dim lErro As Long

On Error GoTo Erro_Carrega_ComboHistorico

    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_HISTORICO, Historico(objCT), True, False, True)
    If lErro <> SUCESSO Then gError 132229

    Carrega_ComboHistorico = SUCESSO

    Exit Function

Erro_Carrega_ComboHistorico:

    Carrega_ComboHistorico = gErr

    Select Case gErr
    
        Case 132229

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179313)

    End Select

    Exit Function

End Function

Private Function Carrega_ComboSubConta(ByVal objCT As Object, ByVal lHistorico As Long) As Long
'Carrega a combo de SubConta

Dim lErro As Long

On Error GoTo Erro_Carrega_ComboSubConta

    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_SUBCONTA, SubConta(objCT), True, False, True, CStr(lHistorico))
    If lErro <> SUCESSO Then gError 132230

    Carrega_ComboSubConta = SUCESSO

    Exit Function

Erro_Carrega_ComboSubConta:

    Carrega_ComboSubConta = gErr

    Select Case gErr
    
        Case 132230

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179314)

    End Select

    Exit Function

End Function

Function NFPAG_Form_Load(ByVal objCT As Object) As Long

Dim lErro As Long

On Error GoTo Erro_NFPAG_Form_Load

    Call Carrega_ComboSituacao(objCT)
    Call Carrega_ComboHistorico(objCT)
   
    NFPAG_Form_Load = SUCESSO
     
    Exit Function
    
Erro_NFPAG_Form_Load:

    NFPAG_Form_Load = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179315)
     
    End Select
     
    Exit Function

End Function

Public Sub Situacao_Change(ByVal objCT As Object)

    objCT.iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub Situacao_Click(ByVal objCT As Object)

    objCT.iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub Historico_Change(ByVal objCT As Object)

    objCT.iAlterado = REGISTRO_ALTERADO
    
    If Historico(objCT).ListIndex <> -1 Then
        Call Carrega_ComboSubConta(objCT, Historico(objCT).ItemData(Historico(objCT).ListIndex))
    End If
    
End Sub

Public Sub Historico_Click(ByVal objCT As Object)

    objCT.iAlterado = REGISTRO_ALTERADO
    
    If Historico(objCT).ListIndex <> -1 Then
        Call Carrega_ComboSubConta(objCT, Historico(objCT).ItemData(Historico(objCT).ListIndex))
    End If
    
End Sub

Public Sub SubConta_Change(ByVal objCT As Object)

    objCT.iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub SubConta_Click(ByVal objCT As Object)

    objCT.iAlterado = REGISTRO_ALTERADO
    
End Sub




VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTNFFatEntComHic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'--- inicio dos properties get dos controles da tela

'##############################################
'Inserido por Wagner
Public Property Get Historico(ByVal objCT As Object) As Object
     Set Historico = objCT.objUserControl.Controls("Historico")
End Property

Public Property Get SubConta(ByVal objCT As Object) As Object
     Set SubConta = objCT.objUserControl.Controls("SubConta")
End Property
'##############################################

Public Property Get TaxaConversao(ByVal objCT As Object) As Object
     Set TaxaConversao = objCT.objUserControl.Controls("TaxaConversao")
End Property

'--- fim dos properties get dos controles da tela

Public Sub TaxaConversao_Change(ByVal objCT As Object)

    objCT.iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub TaxaConversao_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_TaxaConversao_Validate

    'Verifica se algum valor foi digitado
    If Len(Trim(TaxaConversao(objCT).Text)) > 0 Then

        'Critica se é valor Positivo
        lErro = Valor_Positivo_Critica_Double(TaxaConversao(objCT).Text)
        If lErro <> SUCESSO Then gError 103378
    
        'Põe o valor formatado na tela
        TaxaConversao(objCT).Text = Format(TaxaConversao(objCT).Text, FORMATO_TAXA_CONVERSAO_MOEDA)
        
    End If
    
    Exit Sub
     
Erro_TaxaConversao_Validate:

    Cancel = True
    
    Select Case gErr
          
        Case 103378

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179254)
     
    End Select
     
    Exit Sub

End Sub

Function NFEntC_Traz_NFiscal_Tela(ByVal objCT As Object, objNFiscal As ClassNFiscal) As Long

Dim lErro As Long

On Error GoTo Erro_NFEntC_Traz_NFiscal_Tela

    '####################################
    'Inserido por Wagner
    lErro = CF("NFiscal_Le_InfoUsu", objNFiscal)
    If lErro <> SUCESSO Then gError 132361
    '####################################

    If objNFiscal.dTaxaMoeda <> 0 Then
        TaxaConversao(objCT).Text = Format(objNFiscal.dTaxaMoeda, FORMATO_TAXA_CONVERSAO_MOEDA)
    Else
        TaxaConversao(objCT).Text = ""
    End If
    
    '####################################
    'Inserido por Wagner
    Call Combo_Seleciona_ItemData(Historico(objCT), objNFiscal.objInfoUsu.lHistorico)
    Call Combo_Seleciona_ItemData(SubConta(objCT), objNFiscal.objInfoUsu.lSubConta)
    '####################################
    
    NFEntC_Traz_NFiscal_Tela = SUCESSO
     
    Exit Function
    
Erro_NFEntC_Traz_NFiscal_Tela:

    NFEntC_Traz_NFiscal_Tela = gErr
     
    Select Case gErr
          
        Case 132361
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179255)
     
    End Select
     
    Exit Function

End Function

Function NFEntC_Move_Tela_Memoria(ByVal objCT As Object, objNFiscal As ClassNFiscal) As Long

Dim lErro As Long
Dim objNFiscalHic As ClassNFiscalHic

On Error GoTo Erro_NFEntC_Move_Tela_Memoria

    Set objNFiscalHic = New ClassNFiscalHic

    objNFiscal.dTaxaMoeda = StrParaDbl(TaxaConversao(objCT).Text)
    
    '#############################
    'Inserido por Wagner
    If Historico(objCT).ListIndex <> -1 Then
        objNFiscalHic.lHistorico = Historico(objCT).ItemData(Historico(objCT).ListIndex)
    End If
    
    If SubConta(objCT).ListIndex <> -1 Then
        objNFiscalHic.lSubConta = SubConta(objCT).ItemData(SubConta(objCT).ListIndex)
    End If
    
    Set objNFiscal.objInfoUsu = objNFiscalHic
    '#############################
    
    NFEntC_Move_Tela_Memoria = SUCESSO
     
    Exit Function
    
Erro_NFEntC_Move_Tela_Memoria:

    NFEntC_Move_Tela_Memoria = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179256)
     
    End Select
     
    Exit Function

End Function

'##################################################
'Inserido por Wagner
Function NFEntC_Limpa_Tela(ByVal objCT As Object) As Long

Dim lErro As Long

On Error GoTo Erro_NFEntC_Limpa_Tela

    Historico(objCT).ListIndex = -1
    SubConta(objCT).ListIndex = -1
    
    NFEntC_Limpa_Tela = SUCESSO
     
    Exit Function
    
Erro_NFEntC_Limpa_Tela:

    NFEntC_Limpa_Tela = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179257)
     
    End Select
     
    Exit Function

End Function

Private Function Carrega_ComboHistorico(ByVal objCT As Object) As Long
'Carrega a combo de Historico

Dim lErro As Long

On Error GoTo Erro_Carrega_ComboHistorico

    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_HISTORICO, Historico(objCT), True, False, True)
    If lErro <> SUCESSO Then gError 132330

    Carrega_ComboHistorico = SUCESSO

    Exit Function

Erro_Carrega_ComboHistorico:

    Carrega_ComboHistorico = gErr

    Select Case gErr
    
        Case 132330

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179258)

    End Select

    Exit Function

End Function

Private Function Carrega_ComboSubConta(ByVal objCT As Object, ByVal lHistorico As Long) As Long
'Carrega a combo de SubConta

Dim lErro As Long

On Error GoTo Erro_Carrega_ComboSubConta

    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_SUBCONTA, SubConta(objCT), True, False, True, CStr(lHistorico))
    If lErro <> SUCESSO Then gError 132331

    Carrega_ComboSubConta = SUCESSO

    Exit Function

Erro_Carrega_ComboSubConta:

    Carrega_ComboSubConta = gErr

    Select Case gErr
    
        Case 132331

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179259)

    End Select

    Exit Function

End Function

Function NFEntC_Form_Load(ByVal objCT As Object) As Long

Dim lErro As Long

On Error GoTo Erro_NFEntC_Form_Load

    Call Carrega_ComboHistorico(objCT)
   
    NFEntC_Form_Load = SUCESSO
     
    Exit Function
    
Erro_NFEntC_Form_Load:

    NFEntC_Form_Load = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179260)
     
    End Select
     
    Exit Function

End Function

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

'##################################################



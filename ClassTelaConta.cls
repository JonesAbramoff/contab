VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTelaConta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Rotinas auxiliares para telas de entrada de dados que possuem campos do tipo "".Conta"""
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Function Conta_Perde_Foco(ByVal Conta As Object, ByVal Desc As Object) As Long
'recebe MaskEdBox da conta e o label da descrição

Dim lErro As Long
Dim objPlanoConta As New ClassPlanoConta
Dim sCtaFormatada As String
Dim iCtaPreenchida As Integer

On Error GoTo Erro_Conta_Perde_Foco

    sCtaFormatada = String(STRING_CONTA, 0)

    lErro = CF("Conta_Formata", Conta.Text, sCtaFormatada, iCtaPreenchida)
    If lErro Then Error 13053
    
    If iCtaPreenchida = CONTA_PREENCHIDA Then
    
        'verifica se a Conta Final existe
        lErro = CF("Conta_SelecionaUma", sCtaFormatada, objPlanoConta, MODULO_CONTABILIDADE)
        If lErro <> SUCESSO And lErro <> 6030 Then Error 13054

        If lErro = 6030 Then

            Conta.PromptInclude = False
            Conta.Text = ""
            Conta.PromptInclude = True
            Conta.SetFocus

            Error 13055

        End If

        Desc.Caption = objPlanoConta.sDescConta

    Else

        Desc.Caption = ""

    End If

    Conta_Perde_Foco = SUCESSO

    Exit Function

Erro_Conta_Perde_Foco:

    Conta_Perde_Foco = Err
    
    Select Case Err

        Case 13053

        Case 13054

        Case 13055

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153596)

    End Select

    Exit Function

End Function

Public Function Traz_Conta_MaskEd(ByVal sConta As String, ByVal MaskEdConta As Object, ByVal Label As Object) As Long

Dim objPlanoConta As New ClassPlanoConta
Dim lErro As Long
Dim sContaEnxuta As String

On Error GoTo Erro_Traz_Conta_MaskEd

    sContaEnxuta = String(STRING_CONTA, 0)

    If sConta <> "" Then
        lErro = Mascara_RetornaContaEnxuta(sConta, sContaEnxuta)
        If lErro Then Error 13048
    Else
        sContaEnxuta = "0"
    End If
    
    If sContaEnxuta <> "0" Then

        'le a conta
        lErro = CF("Conta_SelecionaUma", sConta, objPlanoConta, MODULO_CONTABILIDADE)
        If lErro <> SUCESSO And lErro <> 6030 Then Error 13046
        
        'nao encontrou a conta
        If lErro = 6030 Then Error 13047

        MaskEdConta.PromptInclude = False
        MaskEdConta.Text = sContaEnxuta
        MaskEdConta.PromptInclude = True
        Label.Caption = objPlanoConta.sDescConta
    Else
        MaskEdConta.PromptInclude = False
        MaskEdConta.Text = ""
        MaskEdConta.PromptInclude = True
        Label.Caption = ""
    End If

    Traz_Conta_MaskEd = SUCESSO

    Exit Function

Erro_Traz_Conta_MaskEd:

    Traz_Conta_MaskEd = Err

    Select Case Err

        Case 13047
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTA1", Err, gsConta)
        
        Case 13046
                    
        Case 13048
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, sConta)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153597)
        
    End Select

    Exit Function
        
End Function

Function Inicializa_Mascara_Conta_MaskEd(ByVal MaskEdConta As Object) As Long
'inicializa a mascara de conta

Dim sMascaraConta As String
Dim lErro As Long

On Error GoTo Erro_Inicializa_Mascara_Conta_MaskEd

    'Inicializa a máscara de Conta
    sMascaraConta = String(STRING_CONTA, 0)
    
    'le a mascara das contas
    lErro = MascaraConta(sMascaraConta)
    If lErro <> SUCESSO Then Error 13038
    
    MaskEdConta.Mask = sMascaraConta
    
    If MaskEdConta.BorderStyle = 0 Then
        If gobjCTB.iESTTamanhoMinimoContaGrid > MaskEdConta.Width Then
            MaskEdConta.Width = gobjCTB.iESTTamanhoMinimoContaGrid
        End If
    End If
    
    Inicializa_Mascara_Conta_MaskEd = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Mascara_Conta_MaskEd:

    Inicializa_Mascara_Conta_MaskEd = Err
    
    Select Case Err

        Case 13038

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153598)

    End Select

    Exit Function

End Function

Function Inicializa_Mascara_Ccl_MaskEd(ByVal objMaskEdCcl As Object) As Long
'inicializa a mascara de centro de custo

Dim sMascaraCcl As String
Dim lErro As Long

On Error GoTo Erro_Inicializa_Mascara_Ccl_MaskEd

    'Inicializa a máscara de Ccl
    sMascaraCcl = String(STRING_CCL, 0)
    
    'le a mascara das ccl
    lErro = MascaraCcl(sMascaraCcl)
    If lErro <> SUCESSO Then gError 89253
    
    objMaskEdCcl.Mask = sMascaraCcl
    
    Inicializa_Mascara_Ccl_MaskEd = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Mascara_Ccl_MaskEd:

    Inicializa_Mascara_Ccl_MaskEd = gErr
    
    Select Case gErr

        Case 89253

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153599)

    End Select

    Exit Function

End Function



VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTNFFATPAGHic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'--- inicio dos properties get dos controles da tela

Public Property Get TaxaConversao(ByVal objCT As Object) As Object
     Set TaxaConversao = objCT.objUserControl.Controls("TaxaConversao")
End Property

'#########################################################
'Inserido por Wagner
Public Property Get Situacao(ByVal objCT As Object) As Object
     Set Situacao = objCT.objUserControl.Controls("Situacao")
End Property

Public Property Get Historico(ByVal objCT As Object) As Object
     Set Historico = objCT.objUserControl.Controls("HistoricoCG")
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
'#########################################################

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179264)
     
    End Select
     
    Exit Sub

End Sub

Function NFFPAG_Traz_NFFatPag_Tela(ByVal objCT As Object, objTituloPagar As ClassTituloPagar) As Long

Dim lErro As Long

On Error GoTo Erro_NFFPAG_Traz_NFFatPag_Tela

    If objTituloPagar.dTaxaMoeda <> 0 Then
        TaxaConversao(objCT).Text = Format(objTituloPagar.dTaxaMoeda, FORMATO_TAXA_CONVERSAO_MOEDA)
    Else
        TaxaConversao(objCT).Text = ""
    End If
    
    '###########################################################
    'Inserido por Wagner
    Call Combo_Seleciona_ItemData(Situacao(objCT), objTituloPagar.objInfoUsu.lSituacao)
    Call Combo_Seleciona_ItemData(Historico(objCT), objTituloPagar.objInfoUsu.lHistorico)
    Call Combo_Seleciona_ItemData(SubConta(objCT), objTituloPagar.objInfoUsu.lSubConta)

    Observacao(objCT).Text = objTituloPagar.objInfoUsu.sObservacao

    If objTituloPagar.iStatus = 0 Then
        Status(objCT).Caption = ""
    Else
        If objTituloPagar.iStatus = STATUS_BAIXADO Then
            Status(objCT).Caption = objTituloPagar.iStatus & SEPARADOR & STRING_STATUS_BAIXADO
        Else
            Status(objCT).Caption = objTituloPagar.iStatus & SEPARADOR & "NORMAL"
        End If
    End If
    '###########################################################
    
    NFFPAG_Traz_NFFatPag_Tela = SUCESSO
     
    Exit Function
    
Erro_NFFPAG_Traz_NFFatPag_Tela:

    NFFPAG_Traz_NFFatPag_Tela = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179265)
     
    End Select
     
    Exit Function

End Function

Function NFFPAG_Move_Tela_Memoria(ByVal objCT As Object, objTituloPagar As ClassTituloPagar) As Long

Dim lErro As Long
Dim objNFPagTitPag As New ClassNfPagTitPagHic

On Error GoTo Erro_NFFPAG_Move_Tela_Memoria

    objTituloPagar.dTaxaMoeda = StrParaDbl(TaxaConversao(objCT).Text)
    
    '###########################################################
    'Inserido por Wagner
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
    Set objTituloPagar.objInfoUsu = objNFPagTitPag
    '###########################################################
    
    NFFPAG_Move_Tela_Memoria = SUCESSO
     
    Exit Function
    
Erro_NFFPAG_Move_Tela_Memoria:

    NFFPAG_Move_Tela_Memoria = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179266)
     
    End Select
     
    Exit Function

End Function

Function NFFPAG_Form_Unload(ByVal objCT As Object, Cancel As Integer) As Long

    'liberar objetos globais referentes a customizacao
    
    NFFPAG_Form_Unload = SUCESSO
    
End Function

'#############################################################
'Inserido por Wagner
Private Function Carrega_ComboSituacao(ByVal objCT As Object) As Long
'Carrega a combo de Situacao

Dim lErro As Long

On Error GoTo Erro_Carrega_ComboSituacao

    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_SITUACAO, Situacao(objCT), True, False, True)
    If lErro <> SUCESSO Then gError 132222

    Carrega_ComboSituacao = SUCESSO

    Exit Function

Erro_Carrega_ComboSituacao:

    Carrega_ComboSituacao = gErr

    Select Case gErr
    
        Case 132222

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179267)

    End Select

    Exit Function

End Function

Private Function Carrega_ComboHistorico(ByVal objCT As Object) As Long
'Carrega a combo de Historico

Dim lErro As Long

On Error GoTo Erro_Carrega_ComboHistorico

    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_HISTORICO, Historico(objCT), True, False, True)
    If lErro <> SUCESSO Then gError 132226

    Carrega_ComboHistorico = SUCESSO

    Exit Function

Erro_Carrega_ComboHistorico:

    Carrega_ComboHistorico = gErr

    Select Case gErr
    
        Case 132226

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179268)

    End Select

    Exit Function

End Function

Private Function Carrega_ComboSubConta(ByVal objCT As Object, ByVal lHistorico As Long) As Long
'Carrega a combo de SubConta

Dim lErro As Long

On Error GoTo Erro_Carrega_ComboSubConta

    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_SUBCONTA, SubConta(objCT), True, False, True, CStr(lHistorico))
    If lErro <> SUCESSO Then gError 132227

    Carrega_ComboSubConta = SUCESSO

    Exit Function

Erro_Carrega_ComboSubConta:

    Carrega_ComboSubConta = gErr

    Select Case gErr
    
        Case 132227

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179269)

    End Select

    Exit Function

End Function

Function NFFPAG_Form_Load(ByVal objCT As Object) As Long

Dim lErro As Long

On Error GoTo Erro_NFFPAG_Form_Load

    Call Carrega_ComboSituacao(objCT)
    Call Carrega_ComboHistorico(objCT)
    
    NFFPAG_Form_Load = SUCESSO
     
    Exit Function
    
Erro_NFFPAG_Form_Load:

    NFFPAG_Form_Load = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179270)
     
    End Select
     
    Exit Function

End Function

Function NFFPAG_Limpa_Tela(ByVal objCT As Object) As Long

Dim lErro As Long

On Error GoTo Erro_NFFPAG_Limpa_Tela

    Situacao(objCT).ListIndex = -1
    Historico(objCT).ListIndex = -1
    SubConta(objCT).ListIndex = -1
    
    Status(objCT).Caption = ""
    
    NFFPAG_Limpa_Tela = SUCESSO
     
    Exit Function
    
Erro_NFFPAG_Limpa_Tela:

    NFFPAG_Limpa_Tela = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179271)
     
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

Function NFFPAG_Trata_Parametros(ByVal objCT As Object, objTituloPagar As ClassTituloPagar) As Long

Dim lErro As Long

On Error GoTo Erro_NFFPAG_Trata_Parametros

    lErro = CF("TituloPagar_Le_Todos", objTituloPagar)
    If lErro <> SUCESSO Then gError 132236
    
    NFFPAG_Trata_Parametros = SUCESSO

    Exit Function

Erro_NFFPAG_Trata_Parametros:

    NFFPAG_Trata_Parametros = gErr

    Select Case gErr
    
        Case 132236

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179272)

    End Select

    Exit Function
    
End Function
'#############################################################




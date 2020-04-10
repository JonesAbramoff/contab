VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTOrcamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Dim m_objUserControl As Object

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim objGrid1 As AdmGrid
Dim iAlterado As Integer
Dim iExercicio1 As Integer

Private WithEvents objEventoOrcamento As AdmEvento
Attribute objEventoOrcamento.VB_VarHelpID = -1
Private WithEvents objEventoCcl As AdmEvento
Attribute objEventoCcl.VB_VarHelpID = -1
Private WithEvents objEventoConta As AdmEvento
Attribute objEventoConta.VB_VarHelpID = -1

'--- inicio dos properties get dos controles da tela

Public Property Get BotaoGravar() As Object
     Set BotaoGravar = objUserControl.Controls("BotaoGravar")
End Property

Public Property Get BotaoExcluir() As Object
     Set BotaoExcluir = objUserControl.Controls("BotaoExcluir")
End Property

Public Property Get BotaoLimpar() As Object
     Set BotaoLimpar = objUserControl.Controls("BotaoLimpar")
End Property

Public Property Get BotaoFechar() As Object
     Set BotaoFechar = objUserControl.Controls("BotaoFechar")
End Property

Public Property Get Exercicio() As Object
     Set Exercicio = objUserControl.Controls("Exercicio")
End Property

Public Property Get Valor() As Object
     Set Valor = objUserControl.Controls("Valor")
End Property

Public Property Get TvwContas() As Object
     Set TvwContas = objUserControl.Controls("TvwContas")
End Property

Public Property Get TvwCcls() As Object
     Set TvwCcls = objUserControl.Controls("TvwCcls")
End Property

Public Property Get GridOrcamento() As Object
     Set GridOrcamento = objUserControl.Controls("GridOrcamento")
End Property

Public Property Get Ccl() As Object
     Set Ccl = objUserControl.Controls("Ccl")
End Property

Public Property Get Conta() As Object
     Set Conta = objUserControl.Controls("Conta")
End Property

Public Property Get LabelTvwCcls() As Object
     Set LabelTvwCcls = objUserControl.Controls("LabelTvwCcls")
End Property

Public Property Get LabelTvwContas() As Object
     Set LabelTvwContas = objUserControl.Controls("LabelTvwContas")
End Property

Public Property Get LabelConta() As Object
     Set LabelConta = objUserControl.Controls("LabelConta")
End Property

Public Property Get LabelCCL() As Object
     Set LabelCCL = objUserControl.Controls("LabelCcl")
End Property

Public Property Get TotalOrcamento() As Object
     Set TotalOrcamento = objUserControl.Controls("TotalOrcamento")
End Property

Public Property Get LabelTotal() As Object
     Set LabelTotal = objUserControl.Controls("LabelTotal")
End Property

'--- fim dos properties get dos controles da tela

Public Sub BotaoExcluir_Click()

Dim lErro As Long
Dim iContaPreenchida As Integer
Dim iCclPreenchida As Integer
Dim sConta As String
Dim sCcl As String
Dim objOrcamento As New ClassOrcamento
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'critica o formato da conta
    lErro = CF("Conta_Formata", Conta.Text, sConta, iContaPreenchida)
    If lErro <> SUCESSO Then Error 16061

    'testa se a conta está preenchida
    If iContaPreenchida <> CONTA_PREENCHIDA Then Error 16062

    If giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then

        'critica o formato do centro de custo/lucro
        lErro = CF("Ccl_Formata", Ccl.Text, sCcl, iCclPreenchida)
        If lErro <> SUCESSO Then Error 10269

        'testa se o centro de custo/lucro está preenchido
        If iCclPreenchida <> CCL_PREENCHIDA Then sCcl = ""

    Else
        sCcl = ""
    End If

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_ORCAMENTO")

    If vbMsgRes = vbYes Then

        objOrcamento.iFilialEmpresa = giFilialEmpresa
        objOrcamento.sConta = sConta
        objOrcamento.iExercicio = Exercicio.ItemData(Exercicio.ListIndex)
        objOrcamento.sCcl = sCcl

        lErro = CF("Orcamento_Exclui", objOrcamento)
        If lErro <> SUCESSO Then Error 16054

        Call Limpa_Tela_Orcamento

        iAlterado = 0

    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 10269, 16054, 16061

        Case 16062
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_INFORMADA", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 157693)

    End Select

    Exit Sub

End Sub

Public Sub BotaoFechar_Click()

    Unload Me

End Sub

Public Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 10270

    Call Limpa_Tela_Orcamento

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 10270

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 157694)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim colOrcamento1 As New Collection
Dim objOrcamento As ClassOrcamento

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    lErro = Move_Dados_Tela(colOrcamento1)
    If lErro <> SUCESSO Then Error 16070

    Set objOrcamento = colOrcamento1.Item(1)

    lErro = Trata_Alteracao(objOrcamento, objOrcamento.iFilialEmpresa, objOrcamento.iExercicio, objOrcamento.sConta, objOrcamento.sCcl, objOrcamento.iPeriodo)
    If lErro <> SUCESSO Then Error 32305

    lErro = CF("Orcamento_Grava", colOrcamento1)
    If lErro <> SUCESSO Then Error 16047

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = Err

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 16047, 16070, 32305

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 157695)

    End Select

    Exit Function

End Function

Private Function Move_Dados_Tela(colOrcamento As Collection) As Long
'Move os dados da tela para colOrcamento

Dim lErro As Long
Dim iExercicio As Integer
Dim sConta As String
Dim iContaPreenchida As Integer
Dim iPeriodo As Integer
Dim objOrcamento As ClassOrcamento
Dim sCcl As String
Dim iCclPreenchida As Integer

On Error GoTo Erro_Move_Dados_Tela

    iExercicio = Exercicio.ItemData(Exercicio.ListIndex)

    'critica o formato da conta
    lErro = CF("Conta_Formata", Conta.Text, sConta, iContaPreenchida)
    If lErro <> SUCESSO Then Error 16033

    'testa se a conta está preenchida
    If iContaPreenchida <> CONTA_PREENCHIDA Then Error 16032

    If giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then

        'critica o formato do centro de custo/lucro
        lErro = CF("Ccl_Formata", Ccl.Text, sCcl, iCclPreenchida)
        If lErro <> SUCESSO Then Error 10272

        'testa se o centro de custo/lucro está preenchido
        If iCclPreenchida <> CCL_PREENCHIDA Then sCcl = ""

    Else
        sCcl = ""
    End If

    'numero de periodos = linhas de edicao do grid
    For iPeriodo = 1 To objGrid1.ilinhasExistentes

        Set objOrcamento = New ClassOrcamento

        objOrcamento.iFilialEmpresa = giFilialEmpresa
        objOrcamento.sConta = sConta
        objOrcamento.sCcl = sCcl
        objOrcamento.iExercicio = iExercicio
        objOrcamento.iPeriodo = iPeriodo

        'coloca zero como valor, caso o campo nao esteja preenchido
        If Len(GridOrcamento.TextMatrix(iPeriodo, 1)) <> 0 Then
            objOrcamento.dValor = CDbl(GridOrcamento.TextMatrix(iPeriodo, 1))
        Else
            objOrcamento.dValor = 0
        End If

        colOrcamento.Add objOrcamento

    Next

    Move_Dados_Tela = SUCESSO

    Exit Function

Erro_Move_Dados_Tela:

    Move_Dados_Tela = Err

    Select Case Err

        Case 10272, 16033

        Case 16032
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_INFORMADA", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 157696)

    End Select

    Exit Function

End Function

Private Sub Limpa_Tela_Orcamento()
'limpa os campos de input da tela, exceto o exercicio

Dim lErro As Long

    Call Limpa_Grid_Orcamento

    Conta.PromptInclude = False
    Conta.Text = ""
    Conta.PromptInclude = True

    Ccl.PromptInclude = False
    Ccl.Text = ""
    Ccl.PromptInclude = True

    TotalOrcamento.Caption = "0,00"

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

End Sub

Sub Limpa_Grid_Orcamento()

Dim iPeriodo As Integer
Dim ilinhasExistentes As Integer

    ilinhasExistentes = objGrid1.ilinhasExistentes

    Call Grid_Limpa(objGrid1)

    objGrid1.ilinhasExistentes = ilinhasExistentes

    'coloca zeros nos valores dos periodos
    For iPeriodo = 1 To objGrid1.ilinhasExistentes
        GridOrcamento.TextMatrix(iPeriodo, 1) = Format(0, "Standard")
    Next

End Sub

Public Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 16065

    Call Limpa_Tela_Orcamento

    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 16065

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 157697)

    End Select

    Exit Sub

End Sub

Public Sub Ccl_GotFocus()

'    TvwContas.Visible = False
'    LabelTvwContas.Visible = False
'    TvwCcls.Visible = True
'    LabelTvwCcls.Visible = True

End Sub

Public Sub Conta_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Ccl_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Conta_GotFocus()

'    TvwContas.Visible = True
'    LabelTvwContas.Visible = True
'    TvwCcls.Visible = False
'    LabelTvwCcls.Visible = False

End Sub

Public Sub Conta_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sContaFormatada As String
Dim iContaPreenchida As Integer
Dim objPlanoConta1 As New ClassPlanoConta
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Conta_Validate

    If Len(Conta.ClipText) > 0 Then

        sContaFormatada = String(STRING_CONTA, 0)

        'critica o formato da conta
        lErro = CF("Conta_Formata", Conta.Text, sContaFormatada, iContaPreenchida)
        If lErro <> SUCESSO Then Error 16026

        If iContaPreenchida = CONTA_PREENCHIDA Then

            lErro = CF("Conta_SelecionaUma", sContaFormatada, objPlanoConta1, MODULO_CONTABILIDADE)
            If lErro <> SUCESSO And lErro <> 6030 Then Error 16027

            'conta não existe
            If lErro = 6030 Then Error 16029

        End If

    End If

    Exit Sub

Erro_Conta_Validate:

    Cancel = True

    If Not (Parent Is GL_objMDIForm.ActiveForm) Then
        Me.Show
    End If

    Select Case Err

        Case 16026

        Case 16027

        Case 16029
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONTA_INEXISTENTE", objPlanoConta1.sConta)
            If vbMsgRes = vbYes Then

                objPlanoConta1.sConta = sContaFormatada

                'Usuário quer criar esta conta
                Call Chama_Tela("PlanoConta", objPlanoConta1)

            End If

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 157698)

    End Select

    Exit Sub

End Sub

Public Sub Ccl_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sCclFormatada As String
Dim iCclPreenchida As Integer
Dim objCcl As New ClassCcl
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Ccl_Validate

    If Len(Ccl.ClipText) > 0 Then

        sCclFormatada = String(STRING_CCL, 0)

        'critica o formato do centro de custo/lucro
        lErro = CF("Ccl_Formata", Ccl.Text, sCclFormatada, iCclPreenchida)
        If lErro <> SUCESSO Then Error 10282

        If iCclPreenchida = CCL_PREENCHIDA Then

            objCcl.sCcl = sCclFormatada

            lErro = CF("Ccl_Le", objCcl)
            If lErro <> SUCESSO And lErro <> 5599 Then Error 10280

            'verifica se o centro de custo/lucro está cadastrado
            If lErro = 5599 Then Error 10281
            
            If objCcl.iAtivo = 0 Then Error 10283

        End If

    End If

    Exit Sub

Erro_Ccl_Validate:

    Cancel = True

    If Not (Parent Is GL_objMDIForm.ActiveForm) Then
        Me.Show
    End If

    Select Case Err

        Case 10280

        Case 10281
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CCL_INEXISTENTE", objCcl.sCcl)

            If vbMsgRes = vbYes Then

                'Usuário quer criar este centro de custo/lucro
                Call Chama_Tela("CclTela", objCcl)

            End If

        Case 10282
        
        Case 10283
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_INATIVO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 157699)

    End Select

    Exit Sub

End Sub

Public Sub Exercicio_Click()

Dim lErro As Long
Dim objExercicio As New ClassExercicio
Dim iPeriodo As Integer
Dim iIndice As Integer

On Error GoTo Erro_Exercicio_Click

    'se trocou o exercicio selecionado
    If Exercicio.ItemData(Exercicio.ListIndex) <> iExercicio1 Then

        If iExercicio1 <> 0 Then

            'Se houve uma alteração no exercicio anterior, pergunta se quer salvar
            lErro = Teste_Salva(Me, iAlterado)
            If lErro <> SUCESSO Then Error 9636

        End If

        iExercicio1 = Exercicio.ItemData(Exercicio.ListIndex)

        lErro = CF("Exercicio_Le", Exercicio.ItemData(Exercicio.ListIndex), objExercicio)
        If lErro <> SUCESSO And lErro <> 10083 Then Error 9635

        'Se o exercicio não está cadastrado
        If lErro = 10083 Then Error 10094

        'linhas de edicao do grid
        objGrid1.ilinhasExistentes = objExercicio.iNumPeriodos

        Call Grid_Limpa_Orcamento(objGrid1)

        Conta.PromptInclude = False
        Conta.Text = ""
        Conta.PromptInclude = True

        Ccl.PromptInclude = False
        Ccl.Text = ""
        Ccl.PromptInclude = True

        'coloca zeros nos valores dos periodos
        For iPeriodo = 1 To objGrid1.ilinhasExistentes
            GridOrcamento.TextMatrix(iPeriodo, 1) = Format(0, "Standard")
        Next

    End If

    Exit Sub

Erro_Exercicio_Click:

    Select Case Err

        Case 9635

        Case 9636
            For iIndice = 0 To Exercicio.ListCount - 1
                If Exercicio.ItemData(iIndice) = iExercicio1 Then
                    Exercicio.ListIndex = iIndice
                    Exit For
                End If
            Next

        Case 10094
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_NAO_CADASTRADO", Err, Exercicio.ItemData(Exercicio.ListIndex))

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 157700)

    End Select

    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub LabelConta_Click()

Dim objOrcamento As New ClassOrcamento
Dim colSelecao As New Collection
Dim lErro As Long
Dim sContaFormatada As String
Dim sCclFormatada As String
Dim iContaPreenchida As Integer
Dim iCclPreenchida As Integer

On Error GoTo Erro_LabelConta_Click

    objOrcamento.iFilialEmpresa = giFilialEmpresa

    If Len(Exercicio.Text) = 0 Then
        objOrcamento.iExercicio = 0
    Else
        objOrcamento.iExercicio = Exercicio.ItemData(Exercicio.ListIndex)
    End If

    If Len(Conta.ClipText) = 0 Then

        objOrcamento.sConta = ""

    Else

        'critica o formato da conta
        lErro = CF("Conta_Formata", Conta.Text, sContaFormatada, iContaPreenchida)
        If lErro <> SUCESSO Then Error 16081

        objOrcamento.sConta = sContaFormatada

    End If


    If giSetupUsoCcl = CCL_USA_EXTRACONTABIL And Len(Ccl.ClipText) > 0 Then

        'critica o formato do centro de custo
        lErro = CF("Ccl_Formata", Ccl.Text, sCclFormatada, iCclPreenchida)
        If lErro <> SUCESSO Then Error 10283

        objOrcamento.sCcl = sCclFormatada

    Else
        objOrcamento.sCcl = ""

    End If

    Call Chama_Tela("OrcamentoLista", colSelecao, objOrcamento, objEventoOrcamento)

    Exit Sub

Erro_LabelConta_Click:

    Select Case Err

        Case 10283, 16081

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 157701)

    End Select

    Exit Sub

End Sub

Public Sub LabelCcl_Click()

Dim lErro As Long
Dim objOrcamento As New ClassOrcamento
Dim colSelecao As New Collection
Dim sContaFormatada As String
Dim sCclFormatada As String
Dim iContaPreenchida As Integer
Dim iCclPreenchida As Integer

On Error GoTo Erro_LabelCcl_Click

    objOrcamento.iFilialEmpresa = giFilialEmpresa

    If Len(Exercicio.Text) = 0 Then
        objOrcamento.iExercicio = 0
    Else
        objOrcamento.iExercicio = Exercicio.ItemData(Exercicio.ListIndex)
    End If

    If Len(Conta.ClipText) = 0 Then

        objOrcamento.sConta = ""

    Else

        'critica o formato da conta
        lErro = CF("Conta_Formata", Conta.Text, sContaFormatada, iContaPreenchida)
        If lErro <> SUCESSO Then Error 28829

        objOrcamento.sConta = sContaFormatada

    End If


    If giSetupUsoCcl = CCL_USA_EXTRACONTABIL And Len(Ccl.ClipText) > 0 Then

        'critica o formato do centro de custo
        lErro = CF("Ccl_Formata", Ccl.Text, sCclFormatada, iCclPreenchida)
        If lErro <> SUCESSO Then Error 28830

        objOrcamento.sCcl = sCclFormatada

    Else
        objOrcamento.sCcl = ""

    End If

    Call Chama_Tela("OrcamentoLista", colSelecao, objOrcamento, objEventoOrcamento)

    Exit Sub

Erro_LabelCcl_Click:

    Select Case Err

        Case 28829, 28830

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 157702)

    End Select

    Exit Sub

End Sub


Private Sub objEventoOrcamento_evSelecao(obj1 As Object)

Dim objOrcamento As New ClassOrcamento
Dim lErro As Long
Dim sContaEnxuta As String
Dim sCclMascarado As String
Dim iIndice As Integer

On Error GoTo Erro_objEventoOrcamento_evSelecao

    Set objOrcamento = obj1

    For iIndice = 0 To Exercicio.ListCount - 1
        If Exercicio.ItemData(iIndice) = objOrcamento.iExercicio Then
            Exercicio.ListIndex = iIndice
            Exit For
        End If
    Next

    sContaEnxuta = String(STRING_CONTA, 0)

    'coloca a conta no formato conta enxuta
    lErro = Mascara_RetornaContaEnxuta(objOrcamento.sConta, sContaEnxuta)
    If lErro <> SUCESSO Then Error 9646

    Conta.PromptInclude = False
    Conta.Text = sContaEnxuta
    Conta.PromptInclude = True

    If giSetupUsoCcl = CCL_USA_EXTRACONTABIL And Len(objOrcamento.sCcl) > 0 Then


        sCclMascarado = String(STRING_CCL, 0)

        'coloca a conta no formato conta enxuta
        lErro = Mascara_MascararCcl(objOrcamento.sCcl, sCclMascarado)
        If lErro <> SUCESSO Then Error 28834

    Else

        sCclMascarado = ""

    End If


    Ccl.PromptInclude = False
    Ccl.Text = sCclMascarado
    Ccl.PromptInclude = True

    objOrcamento.iFilialEmpresa = giFilialEmpresa

    lErro = Traz_Valores_Orcamento(objOrcamento)
    If lErro <> SUCESSO Then Error 16078

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoOrcamento_evSelecao:

    Select Case Err

        Case 9646
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objOrcamento.sConta)

        Case 16078

        Case 28834
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACCLENXUTA", Err, objOrcamento.sCcl)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 157704)

        End Select

    Exit Sub

End Sub

Private Function Traz_Valores_Orcamento(objOrcamento As ClassOrcamento) As Long
'recupera os valores de orcamento em cada periodo, colocando-os no grid orcamento

Dim lErro As Long
Dim iContaPreenchida As Integer
Dim colOrcamento As New Collection
Dim objOrcamento1 As ClassOrcamento
Dim dSoma As Double

On Error GoTo Erro_Traz_Valores_Orcamento

    dSoma = 0

    Call Limpa_Grid_Orcamento

    lErro = CF("Orcamento_Le", objOrcamento, colOrcamento)
    If lErro <> SUCESSO And lErro <> 16019 Then Error 16069

    If lErro = 16019 Then Error 20298

    'se existe orcamento para a conta x exercicio selecionados
    If lErro = SUCESSO Then

        For Each objOrcamento1 In colOrcamento

            'coloca os valores de orcamento para cada periodo
            GridOrcamento.TextMatrix(objOrcamento1.iPeriodo, 1) = Format(objOrcamento1.dValor, "Standard")

            dSoma = dSoma + objOrcamento1.dValor

        Next

    End If

    'preenche o total
    TotalOrcamento.Caption = Format(dSoma, "Standard")

    iAlterado = 0

    Traz_Valores_Orcamento = SUCESSO

    Exit Function

Erro_Traz_Valores_Orcamento:

    Traz_Valores_Orcamento = Err

    Select Case Err

        Case 16069

        Case 20298
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ORCAMENTO_NAO_CADASTRADO", Err, objOrcamento.iExercicio, objOrcamento.sConta)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 157705)

    End Select

    iAlterado = 0

    Exit Function

End Function

Public Sub TvwContas_Expand(ByVal objNode As MSComctlLib.Node)

Dim lErro As Long

On Error GoTo Erro_TvwContas_Expand

    If objNode.Tag <> NETOS_NA_ARVORE Then

        'move os dados do plano de contas do banco de dados para a arvore colNodes.
        lErro = CF("Carga_Arvore_Conta1", objNode, TvwContas.Nodes)
        If lErro <> SUCESSO Then Error 40817

    End If

    Exit Sub

Erro_TvwContas_Expand:

    Select Case Err

        Case 40817

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 157706)

    End Select

    Exit Sub

End Sub

Public Sub TvwContas_NodeClick(ByVal Node As MSComctlLib.Node)

Dim lErro As Long
Dim iPeriodo As Integer
Dim iContaPreenchida As Integer
Dim sContaEnxuta As String
Dim sConta As String
Dim colOrcamento As Collection
Dim objOrcamento As New ClassOrcamento
Dim sContaFormatada As String
Dim sCclFormatada As String
Dim iCclPreenchida As Integer

On Error GoTo Erro_TvwContas_NodeClick

    sConta = Right(Node.Key, Len(Node.Key) - 1)
    sContaEnxuta = String(STRING_CONTA, 0)

    'coloca a conta no formato conta enxuta
    lErro = Mascara_RetornaContaEnxuta(sConta, sContaEnxuta)
    If lErro <> SUCESSO Then Error 16031

    Conta.PromptInclude = False
    Conta.Text = sContaEnxuta
    Conta.PromptInclude = True

    Exit Sub

Erro_TvwContas_NodeClick:

    Select Case Err

        Case 16031
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, sConta)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 157707)

    End Select

    Exit Sub

End Sub

Public Sub TvwCcls_NodeClick(ByVal Node As MSComctlLib.Node)

Dim sCcl As String
Dim sCclMascarado As String
Dim lErro As Long
Dim objOrcamento As New ClassOrcamento
Dim iContaPreenchida As Integer
Dim sContaFormatada As String

On Error GoTo Erro_TvwCcls_NodeClick

    sCcl = Right(Node.Key, Len(Node.Key) - 1)

    sCclMascarado = String(STRING_CCL, 0)

    lErro = Mascara_MascararCcl(sCcl, sCclMascarado)
    If lErro <> SUCESSO Then Error 10289

    Ccl.PromptInclude = False
    Ccl.Text = sCclMascarado
    Ccl.PromptInclude = True

    Exit Sub

Erro_TvwCcls_NodeClick:

    Select Case Err

        Case 10289
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararCcl", Err, sCcl)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 157708)

    End Select

    Exit Sub

End Sub

Public Sub Valor_Change()
    
    iAlterado = REGISTRO_ALTERADO
    If Valor.Text = "0" Then Valor.Text = ""
    
End Sub

Public Sub Valor_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid1)

End Sub

Public Sub Valor_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)

End Sub

Public Sub Valor_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid1.objControle = Valor
    lErro = Grid_Campo_Libera_Foco(objGrid1)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub GridOrcamento_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGrid1, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid1, iAlterado)
    End If

End Sub

Public Sub GridOrcamento_GotFocus()

    Call Grid_Recebe_Foco(objGrid1)

End Sub

Public Sub GridOrcamento_EnterCell()

    Call Grid_Entrada_Celula(objGrid1, iAlterado)

End Sub

Public Sub GridOrcamento_LeaveCell()

    Call Saida_Celula(objGrid1)

End Sub

Public Sub GridOrcamento_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGrid1)

End Sub

Public Sub GridOrcamento_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGrid1, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid1, iAlterado)
    End If

End Sub

Public Sub GridOrcamento_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGrid1)

End Sub

Public Sub GridOrcamento_RowColChange()

    Call Grid_RowColChange(objGrid1)

End Sub

Public Sub GridOrcamento_Scroll()

    Call Grid_Scroll(objGrid1)

End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    Set objEventoOrcamento = Nothing
    Set objEventoCcl = Nothing
    Set objEventoConta = Nothing

    'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

    Set objGrid1 = Nothing

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_Load()

Dim iIndice As Integer
Dim lErro As Long
Dim objExercicio As ClassExercicio
Dim colExercicios As New Collection

On Error GoTo Erro_Form_Load

    Set objEventoOrcamento = New AdmEvento
    Set objEventoCcl = New AdmEvento

    Set objEventoConta = New AdmEvento

    'inicializa a mascara de conta
    lErro = Inicializa_Mascaras()
    If lErro <> SUCESSO Then Error 16028

    'ler os exercicios do banco de dados
    lErro = CF("Exercicios_Le_Todos", colExercicios)
    If lErro <> SUCESSO Then Error 16013

    'carrega combobox Exercicio
    For iIndice = 1 To colExercicios.Count
        Set objExercicio = colExercicios.Item(iIndice)
        Exercicio.AddItem objExercicio.sNomeExterno
        Exercicio.ItemData(Exercicio.NewIndex) = objExercicio.iExercicio
    Next
    
    
    TvwContas.Visible = False
    LabelTvwContas.Visible = False
    TvwCcls.Visible = False
    LabelTvwCcls.Visible = False
    
'    'Inicializa a Lista de Plano de Contas
'    lErro = CF("Carga_Arvore_Conta", TvwContas.Nodes)
'    If lErro <> SUCESSO Then Error 16014

    If giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then

'        'Inicializa a Lista de Centros de Custo
'        lErro = Carga_Arvore_Ccl(TvwCcls.Nodes)
'        If lErro <> SUCESSO Then Error 10287

        LabelCCL.Visible = True
        Ccl.Visible = True

    End If

    Set objGrid1 = New AdmGrid

    'tela em questão
    Set objGrid1.objForm = Me

    lErro = Inicializa_Grid_Orcamento(objGrid1)
    If lErro <> SUCESSO Then Error 16024

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 10287, 16013, 16014, 16024, 16028

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 157709)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Private Function Inicializa_Mascaras() As Long
'inicializa as mascaras de conta e centro de custo

Dim sMascaraConta As String
Dim sMascaraCcl As String
Dim lErro As Long

On Error GoTo Erro_Inicializa_Mascaras

    'Inicializa a máscara de Conta
    sMascaraConta = String(STRING_CONTA, 0)

    'le a mascara das contas
    lErro = MascaraConta(sMascaraConta)
    If lErro <> SUCESSO Then Error 10285

    Conta.Mask = sMascaraConta

    'Se usa centro de custo/lucro ==> inicializa mascara de centro de custo/lucro
    If giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then

        sMascaraCcl = String(STRING_CCL, 0)

        'le a mascara dos centros de custo/lucro
        lErro = MascaraCcl(sMascaraCcl)
        If lErro <> SUCESSO Then Error 10286

        Ccl.Mask = sMascaraCcl

    End If

    Inicializa_Mascaras = SUCESSO

    Exit Function

Erro_Inicializa_Mascaras:

    Inicializa_Mascaras = Err

    Select Case Err

        Case 10285, 10286

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 157710)

    End Select

    Exit Function

End Function

Function Inicializa_Grid_Orcamento(objGridInt As AdmGrid) As Long
'Seta os valores de inicializacao do grid

Dim lErro As Long
Dim iPeriodo As Integer
Dim iGrid_Valor_Col As Integer
Dim iIndice As Integer

On Error GoTo Erro_Inicializa_Grid_Orcamento

    'titulos do grid
    objGridInt.colColuna.Add ("Periodo")
    objGridInt.colColuna.Add ("Valor")

    'campos de edição do grid
    objGridInt.colCampo.Add (Valor.Name)

    objGridInt.objGrid = GridOrcamento

    'todas as linhas do grid
    GridOrcamento.Rows = 13

    'linhas visiveis do grid sem contar com as linhas fixas
    objGridInt.iLinhasVisiveis = 12

    GridOrcamento.ColWidth(0) = 1100

    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    objGridInt.iProibidoExcluir = PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = PROIBIDO_INCLUIR

    Call Grid_Inicializa(objGrid1)

    iGrid_Valor_Col = 1

    TotalOrcamento.Width = GridOrcamento.ColWidth(iGrid_Valor_Col)

    TotalOrcamento.Top = GridOrcamento.Top + GridOrcamento.Height
    TotalOrcamento.Left = GridOrcamento.Left
    For iIndice = 0 To iGrid_Valor_Col - 1
        TotalOrcamento.Left = TotalOrcamento.Left + GridOrcamento.ColWidth(iIndice) + GridOrcamento.GridLineWidth + 20
    Next

    TotalOrcamento.Caption = "0,00"

    Inicializa_Grid_Orcamento = SUCESSO

    Exit Function

Erro_Inicializa_Grid_Orcamento:

    Inicializa_Grid_Orcamento = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 157711)

    End Select

    Exit Function

End Function

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        Set objGridInt.objControle = Valor

        lErro = Saida_Celula_Valor(objGridInt)
        If lErro <> SUCESSO Then Error 16000

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 16001

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = Err

    Select Case Err

        Case 16000

        Case 16001
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 157712)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Valor(objGridInt As AdmGrid) As Long
'faz a critica da celula valor do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer
Dim dTotal As Double

On Error GoTo Erro_Saida_Celula_Valor

    Set objGridInt.objControle = Valor

    If Len(Valor.Text) > 0 Then

        lErro = Valor_Critica(Valor.Text)
        If lErro <> SUCESSO Then Error 16002

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 16003

    For iIndice = 1 To objGridInt.ilinhasExistentes

        dTotal = dTotal + StrParaDbl(GridOrcamento.TextMatrix(iIndice, 1))

    Next

    TotalOrcamento.Caption = Format(dTotal, "Standard")

    Saida_Celula_Valor = SUCESSO

    Exit Function

Erro_Saida_Celula_Valor:

    Saida_Celula_Valor = Err

    Select Case Err

        Case 16002, 16003
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 157713)

    End Select

    Exit Function

End Function

Private Function Carga_Arvore_Ccl(colNodes As Nodes) As Long
'move os dados de centro de custo/lucro do banco de dados para a arvore colNodes.

Dim objNode As Node
Dim colCcl As New Collection
Dim objCcl As ClassCcl
Dim lErro As Long
Dim sCclMascarado As String
Dim sCcl As String
Dim sCclPai As String

On Error GoTo Erro_Carga_Arvore_Ccl

    lErro = CF("Ccl_Le_Todos", colCcl)
    If lErro <> SUCESSO Then Error 24014

    For Each objCcl In colCcl

        sCclMascarado = String(STRING_CCL, 0)

        lErro = Mascara_MascararCcl(objCcl.sCcl, sCclMascarado)
        If lErro <> SUCESSO Then Error 24015

        If objCcl.iTipoCcl = CCL_ANALITICA Then
            sCcl = "A" & objCcl.sCcl
        Else
            sCcl = "S" & objCcl.sCcl
        End If

        sCclPai = String(STRING_CCL, 0)

        'retorna o centro de custo/lucro "pai" do centro de custo/lucro em questão, se houver
        lErro = Mascara_RetornaCclPai(objCcl.sCcl, sCclPai)
        If lErro <> SUCESSO Then Error 24016

        'se o centro de custo/lucro possui um centro de custo/lucro "pai"
        If Len(Trim(sCclPai)) > 0 Then

            sCclPai = "S" & sCclPai

            Set objNode = colNodes.Add(colNodes.Item(sCclPai), tvwChild, sCcl)

        Else
            'se o centro de custo/lucro não possui centro de custo/lucro "pai"
            Set objNode = colNodes.Add(, tvwLast, sCcl)

        End If

        objNode.Text = sCclMascarado & SEPARADOR & objCcl.sDescCcl

    Next

    Carga_Arvore_Ccl = SUCESSO

    Exit Function

Erro_Carga_Arvore_Ccl:

    Carga_Arvore_Ccl = Err

    Select Case Err

        Case 24014

        Case 24015
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararCcl", Err, objCcl.sCcl)

        Case 24016
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_RetornaCclPai", Err, objCcl.sCcl)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 157714)

    End Select

    Exit Function

End Function

Function Trata_Parametros(Optional objOrcamento As ClassOrcamento) As Long
'Trata os parametros que podem ser passados quando ocorre a chamada da tela de Orcamento

Dim iIndice As Integer
Dim iExercicio As Integer
Dim lErro As Long
Dim sContaEnxuta As String
Dim sCclMascarado As String

On Error GoTo Erro_Trata_Parametros

    'Verifica se houve passagem de parametro
    If objOrcamento Is Nothing Then

        Call Limpa_Tela_Orcamento
        
        If lErro_Chama_Tela = SUCESSO Then

            'mostra o exercicio atual
            For iIndice = 0 To Exercicio.ListCount - 1
                If Exercicio.ItemData(iIndice) = giExercicioAtual Then
                    Exercicio.ListIndex = iIndice
                    Exit For
                End If
            Next

        End If

    Else

        'mostra exercicio passado como parametro
        For iIndice = 0 To Exercicio.ListCount - 1
            If Exercicio.ItemData(iIndice) = objOrcamento.iExercicio Then
                Exercicio.ListIndex = iIndice
                Exit For
            End If
        Next

        'mostra conta passada como parametro
        sContaEnxuta = String(STRING_CONTA, 0)

        lErro = Mascara_RetornaContaEnxuta(objOrcamento.sConta, sContaEnxuta)
        If lErro <> SUCESSO Then Error 16020

        Conta.PromptInclude = False
        Conta.Text = sContaEnxuta
        Conta.PromptInclude = True

        If giSetupUsoCcl = CCL_USA_EXTRACONTABIL And Len(objOrcamento.sCcl) > 0 Then
    
            sCclMascarado = String(STRING_CCL, 0)
    
            'coloca a conta no formato conta enxuta
            lErro = Mascara_MascararCcl(objOrcamento.sCcl, sCclMascarado)
            If lErro <> SUCESSO Then Error 28834
    
        Else
    
            sCclMascarado = ""
    
        End If
    
        Ccl.PromptInclude = False
        Ccl.Text = sCclMascarado
        Ccl.PromptInclude = True
    
        lErro = Traz_Valores_Orcamento(objOrcamento)
        If lErro <> SUCESSO Then Error 16078

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 16020
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objOrcamento.sConta)

        Case 16021, 16078, 28834

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 157715)

    End Select

    iAlterado = 0

    Exit Function

End Function

Function Grid_Limpa_Orcamento(objGridInt As AdmGrid) As Long

Dim iLinha As Integer
Dim iColuna As Integer

    For iLinha = objGridInt.ilinhasExistentes + 1 To objGridInt.objGrid.Rows - 1

        For iColuna = objGridInt.objGrid.FixedCols To objGridInt.objGrid.Cols - 1

            objGridInt.objGrid.TextMatrix(iLinha, iColuna) = ""
        Next
    Next

    For iColuna = objGridInt.objGrid.FixedCols To objGridInt.objGrid.Cols - 1
        Call Grid_Limpa_Campo(objGridInt.objForm.Controls(objGridInt.colCampo(iColuna)))
    Next

    objGridInt.objGrid.Row = 0
    objGridInt.objGrid.Col = 0

    Grid_Limpa_Orcamento = SUCESSO

End Function

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim sConta As String
Dim sCcl As String
Dim iContaPreenchida As Integer
Dim iCclPreenchida As Integer

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "Orcamentos"
    
    sConta = String(STRING_CONTA, 0)
    
    'critica o formato da conta
    lErro = CF("Conta_Formata", Conta.Text, sConta, iContaPreenchida)
    If lErro <> SUCESSO Then Error 24134

    If giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then

        sCcl = String(STRING_CCL, 0)
        
        'critica o formato do centro de custo/lucro
        lErro = CF("Ccl_Formata", Ccl.Text, sCcl, iCclPreenchida)
        If lErro <> SUCESSO Then Error 55705

    End If

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "FilialEmpresa", giFilialEmpresa, 0, "FilialEmpresa"
    colCampoValor.Add "Exercicio", Exercicio.ItemData(Exercicio.ListIndex), 0, "Exercicio"
    colCampoValor.Add "Conta", sConta, STRING_CONTA, "Conta"
    colCampoValor.Add "Ccl", sCcl, STRING_CCL, "Ccl"

    'Exemplo de Filtro para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa

    Exit Sub

Erro_Tela_Extrai:

    Select Case Err

        Case 24134, 55705

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 157716)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objOrcamento As New ClassOrcamento
Dim iIndice As Integer
Dim sContaEnxuta As String
Dim sCclEnxuta As String

On Error GoTo Erro_Tela_Preenche

    objOrcamento.sConta = colCampoValor.Item("Conta").vValor

    If objOrcamento.sConta <> "" Then

        objOrcamento.iFilialEmpresa = colCampoValor.Item("FilialEmpresa").vValor
        objOrcamento.iExercicio = colCampoValor.Item("Exercicio").vValor
        objOrcamento.sCcl = colCampoValor.Item("Ccl").vValor


        For iIndice = 0 To Exercicio.ListCount - 1
            If Exercicio.ItemData(iIndice) = objOrcamento.iExercicio Then
                Exercicio.ListIndex = iIndice
                Exit For
            End If
        Next

        sContaEnxuta = String(STRING_CONTA, 0)
    
        lErro = Mascara_RetornaContaEnxuta(objOrcamento.sConta, sContaEnxuta)
        If lErro <> SUCESSO Then Error 57994

        Conta.PromptInclude = False
        Conta.Text = sContaEnxuta
        Conta.PromptInclude = True

        If Len(Trim(objOrcamento.sCcl)) > 0 Then
        
            sCclEnxuta = String(STRING_CCL, 0)
        
            lErro = Mascara_RetornaCclEnxuta(objOrcamento.sCcl, sCclEnxuta)
            If lErro <> SUCESSO Then Error 57995
    
            Ccl.PromptInclude = False
            Ccl.Text = sCclEnxuta
            Ccl.PromptInclude = True
        
        Else
        
            Ccl.PromptInclude = False
            Ccl.Text = ""
            Ccl.PromptInclude = True
        
        End If
        
        lErro = Traz_Valores_Orcamento(objOrcamento)
        If lErro <> SUCESSO Then Error 24135

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case Err

        Case 24135, 57994, 57995

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 157717)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_ORCAMENTO
    Set Form_Load_Ocx = Me
    Caption = "Orçamento"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "Orcamento"

End Function

Public Sub Show()
    Parent.Show
    Parent.SetFocus
End Sub

Private Sub Unload(objme As Object)
   ' Parent.UnloadDoFilho

   RaiseEvent Unload

End Sub

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Parent.Caption = New_Caption
    m_Caption = New_Caption
End Property

'***** fim do trecho a ser copiado ******

Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYCODE_BROWSER Then
        If Me.ActiveControl Is Conta Then
            Call LabelConta_Click
        End If
    End If
End Sub

Public Property Get objUserControl() As Object
    Set objUserControl = m_objUserControl
End Property

Public Property Set objUserControl(ByVal vData As Object)
    Set m_objUserControl = vData
End Property

'Devolve Parent do User Control
Public Property Get Parent() As Object
    Set Parent = objUserControl.Parent
End Property

Public Property Get Controls() As Object
    Set Controls = objUserControl.Controls
End Property

Public Property Get ActiveControl() As Object
    Set ActiveControl = objUserControl.ActiveControl
End Property

Public Property Get Enabled() As Boolean
    Enabled = objUserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    objUserControl.Enabled = New_Enabled
End Property

Public Sub BotaoConta_Click()

Dim objPlanoConta As New ClassPlanoConta
Dim colSelecao As New Collection
    
    'Se o Vendedor estiver preenchido move seu codigo para objVendedor
    If Len(Conta.Text) > 0 Then objPlanoConta.sConta = Conta.Text
    
    'Chama a tela que lista os vendedores
    Call Chama_Tela("PlanoContaLista", colSelecao, objPlanoConta, objEventoConta)

End Sub

Private Sub objEventoConta_evSelecao(obj1 As Object)
    
Dim lErro As Long
Dim objPlanoConta As ClassPlanoConta
Dim sConta As String
Dim sContaEnxuta As String

On Error GoTo Erro_objEventoConta_evSelecao
    
    Set objPlanoConta = obj1
    
    sConta = objPlanoConta.sConta
    
    'le a conta
    lErro = CF("PlanoConta_Le_Conta1", sConta, objPlanoConta)
    If lErro <> SUCESSO And lErro <> 6030 Then gError 197926
    
    If objPlanoConta.iAtivo <> CONTA_ATIVA Then gError 197927
    
    If objPlanoConta.iTipoConta <> CONTA_ANALITICA Then gError 197928
    
    sContaEnxuta = String(STRING_CONTA, 0)

    lErro = Mascara_RetornaContaEnxuta(sConta, sContaEnxuta)
    If lErro <> SUCESSO Then gError 197929

    Conta.PromptInclude = False
    Conta.Text = sContaEnxuta
    Conta.PromptInclude = True

    Me.Show
    
    Exit Sub
    
Erro_objEventoConta_evSelecao:

    Select Case gErr
    
        Case 197926
    
        Case 197927
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTA_INATIVA", gErr, sConta)
        
        Case 197928
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_ANALITICA", gErr, sConta)
    
        Case 197929
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, sConta)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 197930)
        
    End Select

    Exit Sub

End Sub

Public Sub BotaoCcl_Click()

Dim objCcl As New ClassCcl
Dim colSelecao As New Collection
    
    'Se o Vendedor estiver preenchido move seu codigo para objVendedor
    If Len(Ccl.Text) > 0 Then objCcl.sCcl = Ccl.Text
    
    'Chama a tela que lista os vendedores
    Call Chama_Tela("CclLista", colSelecao, objCcl, objEventoCcl)

End Sub

Private Sub objEventoCcl_evSelecao(obj1 As Object)
    
Dim lErro As Long
Dim objCcl As ClassCcl
Dim sConta As String
Dim sCclEnxuta As String

On Error GoTo Erro_objEventoCcl_evSelecao
    
    Set objCcl = obj1

    lErro = CF("Ccl_Le", objCcl)
    If lErro <> SUCESSO And lErro <> 5599 Then gError 197931

    If objCcl.iTipoCcl <> CCL_ANALITICA Then gError 197932
    
    If objCcl.iAtivo = 0 Then gError 197933
    
    sCclEnxuta = String(STRING_CONTA, 0)

    lErro = Mascara_RetornaCclEnxuta(objCcl.sCcl, sCclEnxuta)
    If lErro <> SUCESSO Then gError 197934

    Ccl.PromptInclude = False
    Ccl.Text = sCclEnxuta
    Ccl.PromptInclude = True

    Me.Show
    
    Exit Sub
    
Erro_objEventoCcl_evSelecao:

    Select Case gErr
    
        Case 197931

        Case 197932
            Call Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_ANALITICA1", gErr, objCcl.sCcl)
  
        Case 197933
            Call Rotina_Erro(vbOKOnly, "ERRO_CCL_INATIVO", gErr, objCcl.sCcl)

        Case 197934
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACCLENXUTA", gErr, objCcl.sCcl)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 197935)
        
    End Select

    Exit Sub

End Sub


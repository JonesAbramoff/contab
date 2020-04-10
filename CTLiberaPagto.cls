VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTLiberaPagto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim m_objUserControl As Object

Public gobjInfoUsu As Object

Public gcolTiposDeCobranca As New AdmColCodigoNome

'Property Variables:
Dim m_Caption As String
Event Unload()

'Indicação de seleção no grid
Const S_MARCADO As String = "1"
Const S_DESMARCADO As String = "0"

Public iOrdenacaoAnt As Integer

'Grid Parcelas:
Public objGridParcelas As AdmGrid
Public iFornecedorAlterado As Integer
Public iGrid_Selecionar_Col As Integer
Public iGrid_DataVencimento_Col As Integer
Public iGrid_Tipo_Col As Integer
Public iGrid_DataEmissaoTitulo_Col As Integer
Public iGrid_Numero_Col As Integer
Public iGrid_Parcela_Col As Integer
Public iGrid_Saldo_Col As Integer
Public iGrid_Cobranca_Col As Integer
Public iGrid_ValorParcela_Col As Integer
Public iGrid_FilialEmpresa_Col As Integer
Public iFramePrincipalAlterado As Integer
Public iGrid_Fornecedor_Col As Integer
Public iGrid_Filial_Col As Integer

Public iAlterado As Integer
Dim iGridAtual As Integer
Dim iFrameAtual As Integer
Dim iFramePagamentoAtual As Integer
Public gcolInfoParcPag As Collection

Private WithEvents objEventoFornecedor As AdmEvento
Attribute objEventoFornecedor.VB_VarHelpID = -1

'Indicação de número máximo de linhas dos Grids
Private Const NUM_MAX_PARCELAS_BAIXA = 300

'Constantes públicas dos tabs
Private Const TAB_Titulos = 1
Private Const TAB_Parcelas = 2

'--- inicio dos properties get dos controles da tela

Public Property Get Frame1() As Object
     Set Frame1 = objUserControl.Controls("Frame1")
End Property

Public Property Get FramePagamento() As Object
     Set FramePagamento = objUserControl.Controls("FramePagamento")
End Property

Public Property Get DataEmissao() As Object
     Set DataEmissao = objUserControl.Controls("DataEmissao")
End Property

Public Property Get FrameParcelas() As Object
     Set FrameParcelas = objUserControl.Controls("FrameParcelas")
End Property

Public Property Get Selecionada() As Object
     Set Selecionada = objUserControl.Controls("Selecionada")
End Property

Public Property Get DataVencimento() As Object
     Set DataVencimento = objUserControl.Controls("DataVencimento")
End Property

Public Property Get Saldo() As Object
     Set Saldo = objUserControl.Controls("Saldo")
End Property

Public Property Get Numero() As Object
     Set Numero = objUserControl.Controls("Numero")
End Property

Public Property Get Tipo() As Object
     Set Tipo = objUserControl.Controls("Tipo")
End Property

Public Property Get Parcela() As Object
     Set Parcela = objUserControl.Controls("Parcela")
End Property

Public Property Get GridParcelas() As Object
     Set GridParcelas = objUserControl.Controls("GridParcelas")
End Property

Public Property Get Frame2() As Object
     Set Frame2 = objUserControl.Controls("Frame2")
End Property

Public Property Get Pagamento() As Object
     Set Pagamento = objUserControl.Controls("Pagamento")
End Property

Public Property Get Cobranca() As Object
     Set Cobranca = objUserControl.Controls("Cobranca")
End Property

Public Property Get ValorParcela() As Object
     Set ValorParcela = objUserControl.Controls("ValorParcela")
End Property

Public Property Get FilialEmpresa() As Object
     Set FilialEmpresa = objUserControl.Controls("FilialEmpresa")
End Property

Public Property Get TotalBaixar() As Object
     Set TotalBaixar = objUserControl.Controls("TotalBaixar")
End Property

Public Property Get Frame8() As Object
     Set Frame8 = objUserControl.Controls("Frame8")
End Property

Public Property Get Filial() As Object
     Set Filial = objUserControl.Controls("Filial")
End Property

Public Property Get Fornecedor() As Object
     Set Fornecedor = objUserControl.Controls("Fornecedor")
End Property

Public Property Get FornecLabel() As Object
     Set FornecLabel = objUserControl.Controls("FornecLabel")
End Property

Public Property Get Frame9() As Object
     Set Frame9 = objUserControl.Controls("Frame9")
End Property

Public Property Get Frame4() As Object
     Set Frame4 = objUserControl.Controls("Frame4")
End Property

Public Property Get UpDownEmissaoInic() As Object
     Set UpDownEmissaoInic = objUserControl.Controls("UpDownEmissaoInic")
End Property

Public Property Get EmissaoInic() As Object
     Set EmissaoInic = objUserControl.Controls("EmissaoInic")
End Property

Public Property Get UpDownEmissaoFim() As Object
     Set UpDownEmissaoFim = objUserControl.Controls("UpDownEmissaoFim")
End Property

Public Property Get EmissaoFim() As Object
     Set EmissaoFim = objUserControl.Controls("EmissaoFim")
End Property

Public Property Get Frame5() As Object
     Set Frame5 = objUserControl.Controls("Frame5")
End Property

Public Property Get UpDownVencInic() As Object
     Set UpDownVencInic = objUserControl.Controls("UpDownVencInic")
End Property

Public Property Get VencInic() As Object
     Set VencInic = objUserControl.Controls("VencInic")
End Property

Public Property Get UpDownVencFim() As Object
     Set UpDownVencFim = objUserControl.Controls("UpDownVencFim")
End Property

Public Property Get VencFim() As Object
     Set VencFim = objUserControl.Controls("VencFim")
End Property

Public Property Get Frame6() As Object
     Set Frame6 = objUserControl.Controls("Frame6")
End Property

Public Property Get TituloInic() As Object
     Set TituloInic = objUserControl.Controls("TituloInic")
End Property

Public Property Get TituloFim() As Object
     Set TituloFim = objUserControl.Controls("TituloFim")
End Property

Public Property Get BotaoFechar() As Object
     Set BotaoFechar = objUserControl.Controls("BotaoFechar")
End Property

Public Property Get BotaoLimpar() As Object
     Set BotaoLimpar = objUserControl.Controls("BotaoLimpar")
End Property

Public Property Get BotaoGravar() As Object
     Set BotaoGravar = objUserControl.Controls("BotaoGravar")
End Property

Public Property Get Opcao() As Object
     Set Opcao = objUserControl.Controls("Opcao")
End Property

Public Property Get DataEmissaoTitulo() As Object
     Set DataEmissaoTitulo = objUserControl.Controls("DataEmissaoTitulo")
End Property

Public Property Get BotaoConsultaDocOriginal() As Object
     Set BotaoConsultaDocOriginal = objUserControl.Controls("BotaoConsultaDocOriginal")
End Property


'######################################################################
'Inserido por Wagner 31/07/2006
Public Property Get FornItem() As Object
     Set FornItem = objUserControl.Controls("FornItem")
End Property

Public Property Get FilialFornItem() As Object
     Set FilialFornItem = objUserControl.Controls("FilialFornItem")
End Property

Public Property Get Ordenacao() As Object
     Set Ordenacao = objUserControl.Controls("Ordenacao")
End Property
'######################################################################

Public Property Get TipoDocSeleciona() As Object
     Set TipoDocSeleciona = objUserControl.Controls("TipoDocSeleciona")
End Property

Public Property Get TipoDocTodos() As Object
     Set TipoDocTodos = objUserControl.Controls("TipoDocTodos")
End Property

Public Property Get TipoDocApenas() As Object
     Set TipoDocApenas = objUserControl.Controls("TipoDocApenas")
End Property

'--- fim dos properties get dos controles da tela


Public Sub BotaoFechar_Click()

    'Fecha a tela
    Unload Me

End Sub

Public Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Grava a Baixa
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 42769

    'Limpa a tela
    Call Limpa_Tela_BaixaPagar
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 42769

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155357)

    End Select

    Exit Sub

End Sub

Public Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Verifica se houve alterações e confirma se deseja salvá-las
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 42749

    'Limpa a tela
    Call Limpa_Tela_BaixaPagar
    
    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 42749

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155358)

    End Select

    Exit Sub

End Sub

Public Sub Cobranca_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Public Sub Cobranca_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Public Sub Cobranca_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = Cobranca
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub DataVencimento_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Public Sub DataVencimento_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Public Sub DataVencimento_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = DataVencimento
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub EmissaoFim_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO
    iFramePrincipalAlterado = REGISTRO_ALTERADO

End Sub

Public Sub EmissaoFim_GotFocus()
    
Dim iFrameAux As Integer
    
    iFrameAux = iFramePrincipalAlterado
    Call MaskEdBox_TrataGotFocus(EmissaoFim, iAlterado)
    iFramePrincipalAlterado = iFrameAux

End Sub

Public Sub EmissaoFim_Validate(Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_EmissaoFim_Validate

    'Se a data EmissaoFim está preenchida
    If Len(EmissaoFim.ClipText) > 0 Then

        'Verifica se a data EmissaoFim é válida
        lErro = Data_Critica(EmissaoFim.Text)
        If lErro <> SUCESSO Then Error 42768

        'Se a data EmissaoInic Está preenchida
        If Len(EmissaoInic.ClipText) > 0 Then

            'Verifica se a EmissaoFim é maior ou igual a EmissaoInic
            If CDate(EmissaoFim.Text) < CDate(EmissaoInic.Text) Then Error 15898

        End If

    End If

    Exit Sub

Erro_EmissaoFim_Validate:

    Cancel = True


    Select Case Err

        Case 42768

        Case 15898
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", Err)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155361)

    End Select

    Exit Sub

End Sub

Public Sub EmissaoInic_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO
    iFramePrincipalAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub EmissaoInic_GotFocus()
    
Dim iFrameAux As Integer

    iFrameAux = iFramePrincipalAlterado
    Call MaskEdBox_TrataGotFocus(EmissaoInic, iAlterado)
    iFramePrincipalAlterado = iFrameAux

End Sub

Public Sub EmissaoInic_Validate(Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_EmissaoInic_Validate

    'Se a data EmissaoInic está preenchida
    If Len(EmissaoInic.ClipText) > 0 Then

        'Verifica se a data EmissaoInic é válida
        lErro = Data_Critica(EmissaoInic.Text)
        If lErro <> SUCESSO Then Error 42770

    End If

    Exit Sub

Erro_EmissaoInic_Validate:

    Cancel = True


    Select Case Err

        Case 42770

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155362)

    End Select

    Exit Sub

End Sub

Public Sub Filial_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO
    iFramePrincipalAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Filial_Click()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO
    iFramePrincipalAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Filial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim iCodigo As Integer
Dim sNomeRed As String

On Error GoTo Erro_Filial_Validate

    'Verifica se foi preenchida a ComboBox Filial
    If Len(Trim(Filial.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o ítem selecionado na ComboBox Filial
    If Filial.Text = Filial.List(Filial.ListIndex) Then Exit Sub

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(Filial, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 42771

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        'Verifica se foi preenchido o Fornecedor
        If Len(Trim(Fornecedor.Text)) = 0 Then Error 42772

        'Lê o Fornecedor que está na tela
        sNomeRed = Trim(Fornecedor.Text)

        'Passa o Código da Filial que está na tela para o Obj
        objFilialFornecedor.iCodFilial = iCodigo

        'Lê Filial no BD a partir do NomeReduzido do Fornecedor e Código da Filial
        lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", sNomeRed, objFilialFornecedor)
        If lErro <> SUCESSO And lErro <> 18272 Then Error 42773

        'Se não existe a Filial
        If lErro = 18272 Then Error 42774

        'Encontrou Filial no BD, coloca no Text da Combo
        Filial.Text = CStr(objFilialFornecedor.iCodFilial) & SEPARADOR & objFilialFornecedor.sNome

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then Error 42775

    Exit Sub

Erro_Filial_Validate:

    Cancel = True


    Select Case Err

        Case 42771, 42773

        Case 42772
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", Err)

        Case 42774
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIAL_FORNECEDOR")

            If vbMsgRes = vbYes Then
                'Chama a tela de Filiais
                Call Chama_Tela("FiliaisFornecedores", objFilialFornecedor)
            Else
                'Segura o foco
            End If

        Case 42775
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_ENCONTRADA", Err, Filial.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155363)

    End Select

    Exit Sub

End Sub

Public Sub FilialEmpresa_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Public Sub FilialEmpresa_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Public Sub FilialEmpresa_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = FilialEmpresa
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long
        
    Set gcolInfoParcPag = Nothing
    Set gcolTiposDeCobranca = Nothing
    
    Set objEventoFornecedor = Nothing

    Set objGridParcelas = Nothing
    
End Sub

Public Sub Fornecedor_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO
    iFornecedorAlterado = REGISTRO_ALTERADO
    iFramePrincipalAlterado = REGISTRO_ALTERADO
    
    Call Fornecedor_Preenche
    
End Sub

Public Sub Fornecedor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Fornecedor_Validate

    If iFornecedorAlterado = 0 Then Exit Sub
    
    'Se Fornecedor está preenchido
    If Len(Trim(Fornecedor.Text)) > 0 Then

        'Tenta ler o Fornecedor (NomeReduzido ou Código ou CPF ou CGC)
        lErro = TP_Fornecedor_Le(Fornecedor, objFornecedor, iCodFilial)
        If lErro <> SUCESSO Then Error 15888

        'Lê coleção de códigos, nomes de Filiais do Fornecedor
        lErro = CF("FiliaisFornecedores_Le_Fornecedor", objFornecedor, colCodigoNome)
        If lErro <> SUCESSO Then Error 15889

        'Preenche ComboBox de Filiais
        Call CF("Filial_Preenche", Filial, colCodigoNome)
        
        'Seleciona filial na Combo Filial
        Call CF("Filial_Seleciona", Filial, iCodFilial)
        
    'Se Fornecedor não está preenchido
    ElseIf Len(Trim(Fornecedor.Text)) = 0 Then

        'Limpa a Combo de Filiais
        Filial.Clear

    End If

    iFornecedorAlterado = 0
    
    Exit Sub

Erro_Fornecedor_Validate:

    Cancel = True


    Select Case Err

        Case 15888, 15889

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155364)

    End Select

    Exit Sub

End Sub

Public Sub FornecLabel_Click()
'Chamada do Browse de Fornecedores

Dim colSelecao As Collection
Dim objFornecedor As New ClassFornecedor

    'Passa o Fornecedor que está na tela para o Obj
    objFornecedor.sNomeReduzido = Trim(Fornecedor.Text)

    'Chama a tela com a lista de Fornecedores
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornecedor)

    Exit Sub

End Sub

Public Sub GridParcelas_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridParcelas, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridParcelas, iAlterado)
    End If

End Sub

Public Sub GridParcelas_GotFocus()
    Call Grid_Recebe_Foco(objGridParcelas)
End Sub

Public Sub GridParcelas_EnterCell()
    
    Call Grid_Entrada_Celula(objGridParcelas, iAlterado)
    
End Sub

Public Sub GridParcelas_LeaveCell()
    
    Call Saida_Celula(objGridParcelas)
    
End Sub

Public Sub GridParcelas_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call Grid_Trata_Tecla1(KeyCode, objGridParcelas)
    
End Sub

Public Sub GridParcelas_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridParcelas, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridParcelas, iAlterado)
    End If

End Sub

Public Sub GridParcelas_Validate(Cancel As Boolean)
    
    Call Grid_Libera_Foco(objGridParcelas)
    
End Sub

Public Sub GridParcelas_RowColChange()
    
    Call Grid_RowColChange(objGridParcelas)
    
End Sub

Public Sub GridParcelas_Scroll()
    
    Call Grid_Scroll(objGridParcelas)
    
End Sub

Public Sub Numero_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Public Sub Numero_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Public Sub Numero_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = Numero
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub


Private Sub objEventoFornecedor_evSelecao(obj1 As Object)

Dim objFornecedor As ClassFornecedor

    Me.Show

    'Preenche Fornecedor na tela com NomeReduzido
    Set objFornecedor = obj1
    
    Fornecedor.Text = CStr(objFornecedor.sNomeReduzido)

    'Chama Validate de Fornecedor
    Call Fornecedor_Validate(bSGECancelDummy)
    
    Exit Sub

End Sub

Public Sub Opcao_Click()

Dim lErro As Long

On Error GoTo Erro_Opcao_Click

    'Se Frame atual não corresponde ao Tab clicado
    If Opcao.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, Opcao, Me) <> SUCESSO Then Exit Sub
        
        'Torna Frame selecionado visível
        Frame1(Opcao.SelectedItem.Index).Visible = True
        
        'Torna Frame atual invisível
        Frame1(iFrameAtual).Visible = False
        
        'Armazena novo valor de iFrameAtual
        iFrameAtual = Opcao.SelectedItem.Index
    
        'Se Frame selecionado foi o de Títulos
        If Opcao.SelectedItem.Index = TAB_Titulos Then
            
            Parent.HelpContextID = IDH_BAIXA_PARCELAS_PAGAR_TITULO
            iFramePrincipalAlterado = 0

        'Se Frame selecionado foi o de Parcelas
        ElseIf Opcao.SelectedItem.Index = TAB_Parcelas Then
            
            Parent.HelpContextID = IDH_BAIXA_PARCELAS_PAGAR_PARCELAS
            
            If iFramePrincipalAlterado <> 0 Then
                
                lErro = Carrega_Tab_Parcelas
                If lErro <> SUCESSO Then Error 42767
                
                iFramePrincipalAlterado = 0

            End If
    
        End If

    End If
            
    Exit Sub

Erro_Opcao_Click:

    Select Case Err

        Case 42767
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155365)
            
    End Select

    Exit Sub

End Sub

Public Sub Parcela_GotFocus()


    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Public Sub Parcela_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Public Sub Parcela_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = Parcela
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Saldo_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Public Sub Saldo_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Public Sub Saldo_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = Saldo
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Selecionada_Click()

Dim lErro As Long
Dim objInfoParcPag As ClassInfoParcPag

On Error GoTo Erro_Selecionar_Click

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO

    Call Calcula_Total_Pagar
    
    Exit Sub

Erro_Selecionar_Click:

    Select Case gErr
    
        Case 181429
        
        Case 181430
             Call Rotina_Erro(vbOKOnly, "ERRO_PARC_FORNECEDORES_DIFERENTES", gErr)
    
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Public Sub Tipo_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Public Sub Tipo_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Public Sub Tipo_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = Tipo
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub TituloFim_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO
    iFramePrincipalAlterado = REGISTRO_ALTERADO

End Sub

Public Sub TituloFim_GotFocus()
    
Dim iFrameAux As Integer

    iFrameAux = iFramePrincipalAlterado
    Call MaskEdBox_TrataGotFocus(TituloFim, iAlterado)
    iFramePrincipalAlterado = iFrameAux

End Sub

Public Sub TituloFim_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_TituloFim_Validate

    'Se TituloFim e TituloInic estão preenchidos
    If Len(Trim(TituloFim.Text)) And Len(Trim(TituloInic.Text)) > 0 Then

        'Verifica se TituloFim é maior ou igual que TituloInic
        If CLng(TituloFim.Text) < CLng(TituloInic.Text) Then Error 15909

    End If

    Exit Sub

Erro_TituloFim_Validate:

    Cancel = True


    Select Case Err

        Case 15909
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TITULOINIC_MAIOR_TITULOFIM", Err)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155367)

    End Select

    Exit Sub

End Sub

Public Sub TituloInic_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO
    iFramePrincipalAlterado = REGISTRO_ALTERADO

End Sub

Public Sub TituloInic_GotFocus()
    
Dim iFrameAux As Integer

    iFrameAux = iFramePrincipalAlterado
    Call MaskEdBox_TrataGotFocus(TituloInic, iAlterado)
    iFramePrincipalAlterado = iFrameAux

End Sub

Public Sub UpDownEmissaoFim_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoFim_DownClick

    'Diminui a data EmissaoFim em 1 dia
    lErro = Data_Up_Down_Click(EmissaoFim, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 15903

    Exit Sub

Erro_UpDownEmissaoFim_DownClick:

    Select Case Err

        Case 15903

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155370)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEmissaoFim_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoFim_UpClick

    'Aumenta a data EmissaoFim em 1 dia
    lErro = Data_Up_Down_Click(EmissaoFim, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 15908

    Exit Sub

Erro_UpDownEmissaoFim_UpClick:

    Select Case Err

        Case 15908

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155371)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEmissaoInic_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoInic_DownClick

    'Diminui a data EmissaoInic em 1 dia
    lErro = Data_Up_Down_Click(EmissaoInic, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 15901

    Exit Sub

Erro_UpDownEmissaoInic_DownClick:

    Select Case Err

        Case 15901

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155372)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEmissaoInic_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoInic_UpClick

    'Aumenta a data EmissaoInic em 1 dia
    lErro = Data_Up_Down_Click(EmissaoInic, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 15907

    Exit Sub

Erro_UpDownEmissaoInic_UpClick:

    Select Case Err

        Case 15907

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155373)

    End Select

    Exit Sub

End Sub

Public Sub UpDownVencFim_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownVencFim_DownClick

    'Diminui a data VencFim em 1 dia
    lErro = Data_Up_Down_Click(VencFim, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 15902

    Exit Sub

Erro_UpDownVencFim_DownClick:

    Select Case Err

        Case 15902

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155374)

    End Select

    Exit Sub

End Sub

Public Sub UpDownVencFim_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownVencFim_UpClick

    'Aumenta a data VencFim em 1 dia
    lErro = Data_Up_Down_Click(VencFim, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 15906

    Exit Sub

Erro_UpDownVencFim_UpClick:

    Select Case Err

        Case 15906

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155375)

    End Select

    Exit Sub

End Sub

Public Sub UpDownVencInic_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownVencInic_DownClick

    'Diminui a data VencInic em 1 dia
    lErro = Data_Up_Down_Click(VencInic, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 15904

    Exit Sub

Erro_UpDownVencInic_DownClick:

    Select Case Err

        Case 15904

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155376)

    End Select

    Exit Sub

End Sub

Public Sub UpDownVencInic_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownVencInic_UpClick

    'Aumenta a data VencInic em 1 dia
    lErro = Data_Up_Down_Click(VencInic, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 15905

    Exit Sub

Erro_UpDownVencInic_UpClick:

    Select Case Err

        Case 15905

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155377)

    End Select

    Exit Sub

End Sub

Public Sub Selecionada_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Public Sub Selecionada_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Public Sub Selecionada_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = Selecionada
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then
        
        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 15884

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = Err

    Select Case Err

        Case 15880, 15881, 15882, 15883, 15952, 15958

        Case 5898, 6666, 7777, 8888, 9999

        Case 15884
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 39677
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155378)

    End Select

    Exit Function

End Function

Public Sub Form_Load()

Dim lErro As Long
Dim iSubTipo As Integer

On Error GoTo Erro_Form_Load

    'Visibilidade para versão LIGHT
'    If giTipoVersao = VERSAO_LIGHT Then
'        FilialEmpresa.left = POSICAO_FORA_TELA
'        FilialEmpresa.TabStop = False
'    End If
    
    iFrameAtual = 1
    iFramePagamentoAtual = 0
    iOrdenacaoAnt = 0
    Ordenacao.ListIndex = 0
    
    iFramePrincipalAlterado = REGISTRO_ALTERADO
    
    Set objGridParcelas = New AdmGrid
    Set objEventoFornecedor = New AdmEvento

    Set gcolInfoParcPag = New Collection

    'Lê o código e a descrição de todos os Tipos de Cobrança
    lErro = CF("Cod_Nomes_Le", "TiposDeCobranca", "Codigo", "Descricao", STRING_TIPOSDECOBRANCA_DESCRICAO, gcolTiposDeCobranca)
    If lErro <> SUCESSO Then gError 57340

    lErro = Carrega_TipoDocumento(TipoDocSeleciona)
    If lErro <> SUCESSO Then gError 185035

    'Executa a Inicialização do grid Parcelas
    lErro = Inicializa_Grid_Parcelas(objGridParcelas)
    If lErro <> SUCESSO Then gError 15879

    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 6760, 6761, 15878, 15879, 15885, 15886, 39676, 57340, 79673, 185035

        Case 79674
            Call Rotina_Erro(vbOKOnly, "ERRO_SUBTIPOCONTABIL_TIPOBAIXA_NAO_ENCONTRADO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155381)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Private Function Inicializa_Grid_Parcelas(objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid Parcelas

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add ("  ")
    objGridInt.colColuna.Add ("Liberar")
    objGridInt.colColuna.Add ("Vencimento")
    objGridInt.colColuna.Add ("Tipo")
    objGridInt.colColuna.Add ("Emissão")
    objGridInt.colColuna.Add ("Número")
    objGridInt.colColuna.Add ("Parcela")
    objGridInt.colColuna.Add ("Saldo")
    objGridInt.colColuna.Add ("Fornecedor")
    objGridInt.colColuna.Add ("Filial")
    objGridInt.colColuna.Add ("Cobrança")
    objGridInt.colColuna.Add ("Valor Parcela")
    
'    If giTipoVersao = VERSAO_FULL Then
        objGridInt.colColuna.Add ("Filial Empresa")
'    End If
    
    'campos de edição do grid
    objGridInt.colCampo.Add (Selecionada.Name)
    objGridInt.colCampo.Add (DataVencimento.Name)
    objGridInt.colCampo.Add (Tipo.Name)
    objGridInt.colCampo.Add (DataEmissaoTitulo.Name)
    objGridInt.colCampo.Add (Numero.Name)
    objGridInt.colCampo.Add (Parcela.Name)
    objGridInt.colCampo.Add (Saldo.Name)
    objGridInt.colCampo.Add (FornItem.Name)
    objGridInt.colCampo.Add (FilialFornItem.Name)
    objGridInt.colCampo.Add (Cobranca.Name)
    objGridInt.colCampo.Add (ValorParcela.Name)
    objGridInt.colCampo.Add (FilialEmpresa.Name)
    
    iGrid_Selecionar_Col = 1
    iGrid_DataVencimento_Col = 2
    iGrid_Tipo_Col = 3
    iGrid_DataEmissaoTitulo_Col = 4
    iGrid_Numero_Col = 5
    iGrid_Parcela_Col = 6
    iGrid_Saldo_Col = 7
    iGrid_Fornecedor_Col = 8
    iGrid_Filial_Col = 9
    iGrid_Cobranca_Col = 10
    iGrid_ValorParcela_Col = 11
    iGrid_FilialEmpresa_Col = 12

    objGridInt.objGrid = GridParcelas

    'todas as linhas do grid
    objGridInt.objGrid.Rows = 20 + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 10

    'largura da primeira coluna
    GridParcelas.ColWidth(0) = 600

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'incluir barra de rolagem horizontal
    objGridInt.iIncluirHScroll = GRID_INCLUIR_HSCROLL

    'Não permite incluir novas linhas nem excluir as existentes
    objGridInt.iProibidoIncluir = 1
    objGridInt.iProibidoExcluir = 1

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Parcelas = SUCESSO
        
    Exit Function

End Function

Function Trata_Parametros() As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    iAlterado = 0
    
    Trata_Parametros = SUCESSO
    
    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155382)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Public Sub ValorPago_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ValorParcela_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Public Sub ValorParcela_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Public Sub ValorParcela_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = ValorParcela
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub VencFim_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO
    iFramePrincipalAlterado = REGISTRO_ALTERADO

End Sub

Public Sub VencFim_GotFocus()
    
Dim iFrameAux As Integer
    
    iFrameAux = iFramePrincipalAlterado
    Call MaskEdBox_TrataGotFocus(VencFim, iAlterado)
    iFramePrincipalAlterado = iFrameAux

End Sub

Public Sub VencFim_Validate(Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_VencFim_Validate

    'Se a data VencFim está preenchida
    If Len(VencFim.ClipText) > 0 Then

        'Verifica se a data VencFim é válida
        lErro = Data_Critica(VencFim.Text)
        If lErro <> SUCESSO Then Error 15899

        'Se a data vencInica está preenchida
        If Len(VencInic.ClipText) > 0 Then

            'Verifica se a VencFim é maior ou igual a VencInic
            If CDate(VencFim.Text) < CDate(VencInic.Text) Then Error 15900

        End If

    End If

    Exit Sub

Erro_VencFim_Validate:

    Cancel = True


    Select Case Err

        Case 15899

        Case 15900
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", Err)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155383)

    End Select

    Exit Sub

End Sub

Public Sub VencInic_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO
    iFramePrincipalAlterado = REGISTRO_ALTERADO

End Sub

Public Sub VencInic_GotFocus()
    
Dim iFrameAux As Integer
    
    iFrameAux = iFramePrincipalAlterado
    Call MaskEdBox_TrataGotFocus(VencInic, iAlterado)
    iFramePrincipalAlterado = iFrameAux

End Sub

Public Sub VencInic_Validate(Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_VencInic_Validate

    'Se a data VencInic está preenchida
    If Len(VencInic.ClipText) > 0 Then

        'Verifica se a data VencInic é válida
        lErro = Data_Critica(VencInic.Text)
        If lErro <> SUCESSO Then Error 15896

    End If

    Exit Sub

Erro_VencInic_Validate:

    Cancel = True


    Select Case Err

        Case 15896

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155384)

    End Select

    Exit Sub

End Sub

Public Sub Grid_Parcelas_Preenche(colInfoParcPag As Collection)
'Preenche o Grid Parcelas com os dados de gcolInfoParcPag

Dim lErro As Long
Dim iLinha As Integer
Dim objInfoParcPag As ClassInfoParcPag
Dim objFilialEmpresa As New AdmFiliais
Dim objCodDescricao As New AdmCodigoNome
'#########################
'Inserido por Wagner
Dim objFornecedor As ClassFornecedor
Dim objFilialForn As ClassFilialFornecedor
'#########################

On Error GoTo Erro_Grid_Parcelas_Preenche

    'Se o número de parcelas for maior que o número de linhas do Grid
    If colInfoParcPag.Count + 1 > GridParcelas.Rows Then
    
        'Altera o número de linhas do Grid de acordo com o número de parcelas
        GridParcelas.Rows = colInfoParcPag.Count + 1
        
        'Chama rotina de Inicialização do Grid
        Call Grid_Inicializa(objGridParcelas)

    End If

    iLinha = 0

    'Percorre todas as Parcelas da Coleção
    For Each objInfoParcPag In colInfoParcPag

        iLinha = iLinha + 1

        'Passa para a tela os dados da Parcela em questão
        GridParcelas.TextMatrix(iLinha, iGrid_Selecionar_Col) = S_DESMARCADO
        GridParcelas.TextMatrix(iLinha, iGrid_DataVencimento_Col) = Format(objInfoParcPag.dtDataVencimento, "dd/mm/yyyy")
        GridParcelas.TextMatrix(iLinha, iGrid_Tipo_Col) = objInfoParcPag.sSiglaDocumento
        If objInfoParcPag.dtDataEmissao <> DATA_NULA Then GridParcelas.TextMatrix(iLinha, iGrid_DataEmissaoTitulo_Col) = Format(objInfoParcPag.dtDataEmissao, "dd/mm/yyyy")
        GridParcelas.TextMatrix(iLinha, iGrid_Numero_Col) = objInfoParcPag.lNumTitulo
        GridParcelas.TextMatrix(iLinha, iGrid_Parcela_Col) = objInfoParcPag.iNumParcela
        GridParcelas.TextMatrix(iLinha, iGrid_Saldo_Col) = Format(objInfoParcPag.dValor, "Standard")
        
        '#################################################
        'Inserido por Wagner
        Set objFornecedor = New ClassFornecedor
        Set objFilialForn = New ClassFilialFornecedor
        
        objFornecedor.lCodigo = objInfoParcPag.lFornecedor
        
        objFilialForn.lCodFornecedor = objFornecedor.lCodigo
        objFilialForn.iCodFilial = objInfoParcPag.iFilialForn
            
        'le o nome reduzido do cliente
        lErro = CF("Fornecedor_Le", objFornecedor)
        If lErro <> SUCESSO And lErro <> 12729 And lErro <> 12732 Then gError 181431
            
        'le o nome reduzido da filial  cliente
        lErro = CF("FilialFornecedor_Le", objFilialForn)
        If lErro <> SUCESSO And lErro <> 12929 Then gError 181432
            
        GridParcelas.TextMatrix(iLinha, iGrid_Fornecedor_Col) = objFornecedor.lCodigo & SEPARADOR & objFornecedor.sNomeReduzido
        GridParcelas.TextMatrix(iLinha, iGrid_Filial_Col) = objFilialForn.iCodFilial & SEPARADOR & objFilialForn.sNome
        '#################################################
        
        If objInfoParcPag.iTipoCobranca <> 0 Then
        
            For Each objCodDescricao In gcolTiposDeCobranca
                If objCodDescricao.iCodigo = objInfoParcPag.iTipoCobranca Then
                
                    GridParcelas.TextMatrix(iLinha, iGrid_Cobranca_Col) = objInfoParcPag.iTipoCobranca & SEPARADOR & objCodDescricao.sNome
                    Exit For
                    
                End If
            Next
            
        End If
        
        If objFilialEmpresa.iCodFilial <> objInfoParcPag.iFilialEmpresa Then
        
            objFilialEmpresa.iCodFilial = objInfoParcPag.iFilialEmpresa
            'Lê a Filial Empresa
            lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
            If lErro <> SUCESSO Then gError 43367
        
        End If
        
'        If giTipoVersao = VERSAO_FULL Then
            GridParcelas.TextMatrix(iLinha, iGrid_FilialEmpresa_Col) = objInfoParcPag.iFilialEmpresa & SEPARADOR & objFilialEmpresa.sNome
'        End If
        
        GridParcelas.TextMatrix(iLinha, iGrid_ValorParcela_Col) = Format(objInfoParcPag.dValorOriginal, "Standard")
        
    Next
    
    'Passa para o Obj o número de Parcelas passadas pela Coleção
    objGridParcelas.iLinhasExistentes = colInfoParcPag.Count

    Call Grid_Refresh_Checkbox(objGridParcelas)
    
    Exit Sub

Erro_Grid_Parcelas_Preenche:

    Select Case gErr

        Case 43367, 181431, 181432 'Inserido por Wagner
        
        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155385)

    End Select

    Exit Sub

End Sub

Public Function Frame_Parcelas_Limpa()
'Limpa os comandos do Frame Parcelas

    'Limpa os Grids
    Call Grid_Limpa(objGridParcelas)
    
End Function

Public Function Gravar_Registro() As Long
'Verifica se os principais dados para gravação foram preenchidos

Dim lErro As Long
Dim iIndice As Integer
Dim iParcelaSelecionada As Integer
Dim dSomaPagar As Double
Dim dValorParcela As Double, dValorParcelaBaixa As Double
Dim colBaixaParcPagar As New colBaixaParcPagar
Dim objInfoParcPag As ClassInfoParcPag
Dim dtDataBaixa As Date, vbMsgRes As VbMsgBoxResult
Dim colInfoParcPag As New Collection

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    iParcelaSelecionada = 0

    'Percorre todas as linhas do Grid Parcelas
    For iIndice = 1 To objGridParcelas.iLinhasExistentes
    
        'Passa a linha do Grid para o Obj
        Set objInfoParcPag = gcolInfoParcPag.Item(iIndice)
        
        'Se a checkbox Selecionada está marcada
        If GridParcelas.TextMatrix(iIndice, iGrid_Selecionar_Col) = S_MARCADO Then

            'Conta quantas Parcelas estão selecionadas
            iParcelaSelecionada = iParcelaSelecionada + 1
            
            colInfoParcPag.Add objInfoParcPag
            
        End If
        
    Next

    'Verifica se alguma célula Selecionada do GridParcelas está marcada
    If iParcelaSelecionada < 1 Then Error 15987

    lErro = CF("LiberaPagto_Grava", colInfoParcPag)
    If lErro <> SUCESSO Then Error 15895
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = Err

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 59087, 59088 'desistiu por que descobriu que tinha digitado valor errado
        
        Case 15895, 15897, 15975, 20837
        
        Case 15987
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TITULO_NAO_MARCADO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155392)

    End Select

    Exit Function

End Function

Private Function Limpa_Tela_BaixaPagar()
'Limpa a tela

Dim lErro As Long

    'Limpa a coleção de Parcelas
    Set gcolInfoParcPag = New Collection
    
    'Limpa os comandos da tela
    Call Limpa_Tela(Me)
    
    'Limpa os Labels
    TotalBaixar.Caption = Format(0, "Fixed")
    
    'Limpa os Grid
    Call Grid_Limpa(objGridParcelas)
    
    'Torna Frame atual invisível
    Frame1(Opcao.SelectedItem.Index).Visible = False
    iFrameAtual = TAB_Titulos
    'Torna Frame atual visível
    Frame1(iFrameAtual).Visible = True
    Opcao.Tabs.Item(iFrameAtual).Selected = True
    
    Call Opcao_Click
    
    'Deixa as ComboBOx sem ítem selecionado
    Filial.ListIndex = -1
    
    TipoDocTodos.Value = True
    Call TipoDocTodos_Click

    'Zera iAlterado
    iAlterado = 0
    
    iOrdenacaoAnt = 0
    Ordenacao.ListIndex = 0

End Function

Private Function Carrega_Tab_Parcelas() As Long
'Carrega os dados das parcelas, créditos e pagamentos antecipados para tela

Dim lErro As Long
Dim iCodFilialFornecedor As Integer
Dim iFilialForn As Integer
Dim lCodForn As Long
Dim objFornecedor As New ClassFornecedor
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim colCreditos As New Collection
Dim colPagtoAntecipado As New Collection
Dim colChequePre As New Collection
Dim dtEmissaoInic As Date
Dim dtEmissaoFim As Date
Dim dtVencInic As Date
Dim dtVencFim As Date
Dim lTituloInic As Long
Dim lTituloFim As Long
Dim sTipoDoc As String
    
On Error GoTo Erro_Carrega_Tab_Parcelas
    
    'Verifica se Fornecedor e a filial estão preenchidos
'    If Len(Trim(Fornecedor.Text)) = 0 Then Error 15917
'    If Len(Trim(Filial.Text)) = 0 Then Error 15918

    If iFramePrincipalAlterado <> 0 Then

        'Limpa o frame Parcelas
        Call Frame_Parcelas_Limpa
        
        iOrdenacaoAnt = 0
        Ordenacao.ListIndex = 0

        'Lê os dados do Fornecedor
        If Len(Trim(Fornecedor.Text)) <> 0 Then
            lErro = TP_Fornecedor_Le(Fornecedor, objFornecedor, iCodFilialFornecedor)
            If lErro <> SUCESSO Then Error 29009
        End If
        
        'PassaFornecedor que está na tela para o Obj
        lCodForn = objFornecedor.lCodigo
        objFilialFornecedor.lCodFornecedor = objFornecedor.lCodigo
    
        'Passa a Filial que está na tela para o Obj
        iFilialForn = Codigo_Extrai(Filial.Text)
        objFilialFornecedor.iCodFilial = iFilialForn
    
        dtEmissaoInic = MaskedParaDate(EmissaoInic)
        dtEmissaoFim = MaskedParaDate(EmissaoFim)
        dtVencInic = MaskedParaDate(VencInic)
        dtVencFim = MaskedParaDate(VencFim)
            
        'Se as datas EmissaoFim e EmissaoInic estão preenchidas
        If dtEmissaoInic <> DATA_NULA And dtEmissaoFim <> DATA_NULA Then
            'Verifica se a EmissaoFim é maior ou igual a EmissaoInic
            If dtEmissaoFim < dtEmissaoInic Then Error 15919
        End If
    
        'Se as datas VencFim e VencInic estão preenchidas
        If dtVencInic <> DATA_NULA And dtVencFim <> DATA_NULA Then
            'Verifica se a EmissaoFim é maior ou igual a EmissaoInic
            If dtVencFim < dtVencInic Then Error 15920
        End If
    
        'Lê TituloInic e TituloFim que estão na tela
        lTituloInic = StrParaLong(TituloInic.Text)
        lTituloFim = StrParaLong(TituloFim.Text)
            
        'Verifica se TituloFim é maior ou igual que TituloInic
        If lTituloFim < lTituloInic Then Error 15921
        
        'Limpa gcolInfoParcPag antes de carregar as novas parcelas
        Set gcolInfoParcPag = New Collection

        If TipoDocApenas.Value = True Then
            If TipoDocSeleciona.ListIndex = -1 Then Error 15922
            sTipoDoc = SCodigo_Extrai(TipoDocSeleciona.Text)
        Else
            sTipoDoc = ""
        End If

        'Preenche a Coleção de Parcelas
        lErro = CF("ParcelasPagar_Le_BaixaManual", lCodForn, iFilialForn, dtEmissaoInic, dtEmissaoFim, dtVencInic, dtVencFim, lTituloInic, lTituloFim, gcolInfoParcPag, sTipoDoc, STATUS_SUSPENSO)
        If lErro <> SUCESSO Then Error 15922
    
        'Preenche o GridParcelas
        Call Grid_Parcelas_Preenche(gcolInfoParcPag)
    
    End If
        
    Carrega_Tab_Parcelas = SUCESSO
    
    Exit Function
    
Erro_Carrega_Tab_Parcelas:

    Carrega_Tab_Parcelas = Err
    
    Select Case Err
    
        Case 15917
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", Err)
            
        Case 15918
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", Err)
            
        Case 15919
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_INICIAL_MAIOR", Err)
            
        Case 15920
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_INICIAL_MAIOR", Err)
            
        Case 15921
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TITULOINIC_MAIOR_TITULOFIM", Err)
            
        Case 15922, 15923, 15924, 29009
            
        Case 43092
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PARCELAS_SUPERIOR_NUM_MAX_PARCELAS_BAIXA", Err)
            
        Case 43093
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CREDITOS_SUPERIOR_NUM_MAX_CREDITOS", Err)
            
        Case 43094
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PAGTOSANTECIPADOS_SUPERIOR_NUM_MAX_PAGTOS_ANTECIPADOS", Err)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155394)
            
    End Select
    
    Exit Function
        
End Function

Private Sub Calcula_Total_Pagar()

Dim iIndice As Integer
Dim dTotalSelecionado As Double

    dTotalSelecionado = 0
    
    For iIndice = 1 To objGridParcelas.iLinhasExistentes
        
        If StrParaInt(GridParcelas.TextMatrix(iIndice, iGrid_Selecionar_Col)) <> 0 Then
            
            dTotalSelecionado = dTotalSelecionado + StrParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_Saldo_Col))
            
        End If
        
    Next
    
    If dTotalSelecionado > 0 Then
        TotalBaixar.Caption = Format(dTotalSelecionado, "Standard")
    Else
        TotalBaixar.Caption = ""
    End If
    
    Exit Sub
    
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_BAIXA_PARCELAS_PAGAR_TITULO
    Set Form_Load_Ocx = Me
    Caption = "Liberação de Pagamento"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "LiberaPagto"
    
End Function

Public Sub Show()
    Parent.Show
    Parent.SetFocus
End Sub

Private Sub Unload(objme As Object)
    
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
        
        If Me.ActiveControl Is Fornecedor Then
            Call FornecLabel_Click
        End If
    
    End If
    
End Sub

Public Sub BotaoConsultaDocOriginal_Click()
'Exibe na tela o documento que originou a parcela, o adiantamento ou o crédito selecionado
'Esse botão não está mais visível na tela, mas esse código ainda é chamado a partir de mnuGridConsultaDocOriginal_Click

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim objFilialFornecedor As New ClassFilialFornecedor

On Error GoTo Erro_BotaoConsultaDocOriginal_Click
    
    'Se nenhum grid foi selecionado => erro
    If iGridAtual = NENHUM_GRID_SELECIONADO Then gError 79883
    
    'Verifica em qual grid está o documento que será consultado
    Select Case iGridAtual
    
        'Se for o grid parcelas
        Case GRID_PARCELAS
        
            'Abre a tela de consulta de títulos a pagar
            lErro = Consulta_TituloPagar()
            If lErro <> SUCESSO Then gError 79877
       
    End Select

    Exit Sub
    
Erro_BotaoConsultaDocOriginal_Click:

    Select Case gErr
    
        Case 79876 To 79879
        
        Case 79882
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objFornecedor.sNomeReduzido)
        
        Case 79883
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 79884
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

        Case 79891
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155399)
    
    End Select
    
    Exit Sub
    
End Sub

Private Function Consulta_TituloPagar() As Long
'Exibe o título referente à parcela selecionada no grid
            
Dim lErro As Long
Dim objTituloPagar As New ClassTituloPagar
            
On Error GoTo Erro_Consulta_TituloPagar

    'Se não há linha selecionada no grid parcelas => erro
    If GridParcelas.Row <= 0 Then gError 79874
    
    'Se a linha selecionada não contém dados => erro
    If Len(Trim(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Numero_Col))) = 0 Then gError 79875
    
    'Preenche o obj com os dados que serão utilizados para leitura do título que será consultado
    With objTituloPagar
        
        .lFornecedor = LCodigo_Extrai(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Fornecedor_Col))
        .iFilial = Codigo_Extrai(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Filial_Col))
        .sSiglaDocumento = GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Tipo_Col)
        .lNumTitulo = GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Numero_Col)
        .dtDataEmissao = StrParaDate(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_DataEmissaoTitulo_Col))
    
    End With
    
    Call Chama_Tela("TituloPagar_Consulta", objTituloPagar)
    
    Consulta_TituloPagar = SUCESSO
    
    Exit Function

Erro_Consulta_TituloPagar:

    Consulta_TituloPagar = gErr
    
    Select Case gErr
    
        Case 79874
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
    
        Case 79875
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_PREENCHIDA", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155400)
    
    End Select
    
    Exit Function
    
End Function

Public Sub GridParcelas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Faz com que apareca um PopupMenu o botao direito do mouse acionado sobre o grid

    'Seta o Grid Parcelas como grid atual
    iGridAtual = GRID_PARCELAS

    'Verifica se foi o botao direito do mouse que foi pressionado
    If Button = vbRightButton Then

        'Seta objTela como a Tela de Baixas a Receber
        Set PopUpMenuGridMD.objTela = Me

        'Chama o Menu PopUp
        PopUpMenuGridMD.PopupMenu PopUpMenuGridMD.mnuGrid, vbPopupMenuRightButton

        'Limpa o objTela
        Set PopUpMenuGridMD.objTela = Nothing

    End If

End Sub

Public Sub mnuGridConsultaDocOriginal_Click()
    Call BotaoConsultaDocOriginal_Click
End Sub

Public Sub mnuGridMarcarTodos_Click()
    Select Case iGridAtual
        'Se for o grid parcelas
        Case GRID_PARCELAS
            Call Grid_Marca_Desmarca(objGridParcelas, iGrid_Selecionar_Col, MARCADO)
        
            Call Calcula_Total_Pagar
        
    End Select
End Sub

Public Sub mnuGridDesmarcarTodos_Click()
    Select Case iGridAtual
        'Se for o grid parcelas
        Case GRID_PARCELAS
            Call Grid_Marca_Desmarca(objGridParcelas, iGrid_Selecionar_Col, DESMARCADO)
            TotalBaixar.Caption = ""
    End Select
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

Private Sub Fornecedor_Preenche()
'por Jorge Specian - Para localizar pela parte digitada do Nome
'Reduzido do Fornecedor através da CF Fornecedor_Pesquisa_NomeReduzido em RotinasCPR.ClassCPRSelect'

Static sNomeReduzidoParte As String '*** rotina para trazer cliente
Dim lErro As Long
Dim objFornecedor As Object
    
On Error GoTo Erro_Fornecedor_Preenche
    
    Set objFornecedor = Fornecedor
    
    lErro = CF("Fornecedor_Pesquisa_NomeReduzido", objFornecedor, sNomeReduzidoParte)
    If lErro <> SUCESSO Then gError 134053

    Exit Sub

Erro_Fornecedor_Preenche:

    Select Case gErr

        Case 134053

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155403)

    End Select
    
    Exit Sub

End Sub


Public Function Ordenacao_Click() As Long
    Call Ordenacao_Atualiza
End Function

Public Function Ordenacao_Change() As Long
    Call Ordenacao_Atualiza
End Function

Public Function Ordenacao_Atualiza() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objInfoParcPag As ClassInfoParcPag

On Error GoTo Erro_Ordenacao_Atualiza:

    If iOrdenacaoAnt <> Ordenacao.ListIndex Then

        iIndice = 0
        For Each objInfoParcPag In gcolInfoParcPag
            iIndice = iIndice + 1
            objInfoParcPag.iSelecionado = StrParaInt(GridParcelas.TextMatrix(iIndice, iGrid_Selecionar_Col))
        Next

        lErro = Ordena_gcolInfoParcPag(gcolInfoParcPag)
        If lErro <> SUCESSO Then gError 181421
        
        Call Grid_Limpa(objGridParcelas)
        
        Call Grid_Parcelas_Preenche(gcolInfoParcPag)
        'If lErro <> SUCESSO Then gError 181422
        
        lErro = Marca_Parcelas(gcolInfoParcPag)
        If lErro <> SUCESSO Then gError 181423
        
        iOrdenacaoAnt = Ordenacao.ListIndex
        
    End If

    Ordenacao_Atualiza = SUCESSO

    Exit Function

Erro_Ordenacao_Atualiza:

    Ordenacao_Atualiza = gErr

    Select Case gErr
    
        Case 181421 To 181423

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function
    
End Function

Public Function Ordena_gcolInfoParcPag(colParcelas As Collection) As Long

Dim lErro As Long
Dim colSaida As New Collection
Dim colCampos As New Collection

On Error GoTo Erro_Ordena_gcolInfoParcPag:

    Select Case Ordenacao.ListIndex
        Case 0
            colCampos.Add "lNumTitulo"
            colCampos.Add "iNumParcela"
            colCampos.Add "lFornecedor"
            colCampos.Add "iFilialForn"
        Case 1
            colCampos.Add "lFornecedor"
            colCampos.Add "iFilialForn"
            colCampos.Add "dtDataVencimento"
        Case 2
            colCampos.Add "lFornecedor"
            colCampos.Add "iFilialForn"
            colCampos.Add "lNumTitulo"
            colCampos.Add "iNumParcela"
        Case 3
            colCampos.Add "dtDataVencimento"
            colCampos.Add "lNumTitulo"
            colCampos.Add "iNumParcela"
    End Select
    
    lErro = Ordena_Colecao(colParcelas, colSaida, colCampos)
    If lErro <> SUCESSO Then gError 181420
    
    Set colParcelas = colSaida

    Ordena_gcolInfoParcPag = SUCESSO

    Exit Function

Erro_Ordena_gcolInfoParcPag:

    Ordena_gcolInfoParcPag = gErr

    Select Case gErr
    
        Case 181420

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function
    
End Function

Public Function Marca_Parcelas(ByVal colParcelas As Collection) As Long

Dim lErro As Long
Dim iLinha As Integer
Dim objInfoParcPag As ClassInfoParcPag

On Error GoTo Erro_Marca_Parcelas:

    'Percorre todas as Parcelas da Coleção
    iLinha = 0
    For Each objInfoParcPag In colParcelas

        iLinha = iLinha + 1

        'Passa para a tela os dados da Parcela em questão
        GridParcelas.TextMatrix(iLinha, iGrid_Selecionar_Col) = objInfoParcPag.iSelecionado

    Next

    Call Grid_Refresh_Checkbox(objGridParcelas)

    Marca_Parcelas = SUCESSO

    Exit Function

Erro_Marca_Parcelas:

    Marca_Parcelas = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function
    
End Function
'##############################################################################

Private Function Carrega_TipoDocumento(ByVal objComboBox As ComboBox)
'Carrega os Tipos de Documento

Dim lErro As Long
Dim iIndice As Integer
Dim colTipoDocumento As New Collection
Dim objTipoDocumento As ClassTipoDocumento

On Error GoTo Erro_Carrega_TipoDocumento

    objComboBox.Clear

    'Le os Tipos de Documentos utilizados em Titulos a Receber
    lErro = CF("TiposDocumento_Le_TituloPag", colTipoDocumento)
    If lErro <> SUCESSO Then gError 185032
    
    'Carrega a combobox com as Siglas  - DescricaoReduzida lidas
    For iIndice = 1 To colTipoDocumento.Count
        Set objTipoDocumento = colTipoDocumento.Item(iIndice)
                    
        objComboBox.AddItem objTipoDocumento.sSigla & SEPARADOR & objTipoDocumento.sDescricaoReduzida
    
    Next

    Carrega_TipoDocumento = SUCESSO

    Exit Function

Erro_Carrega_TipoDocumento:

    Carrega_TipoDocumento = gErr

    Select Case gErr

        Case 185032

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 185033)

    End Select

    Exit Function

End Function

Public Sub TipoDocApenas_Click()

    'Habilita a combo para a seleção da conta corrente
    TipoDocSeleciona.Enabled = True
    
    iFramePrincipalAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub TipoDocTodos_Click()

    'Desabilita a combo para a seleção da conta corrente
    TipoDocSeleciona.Enabled = False

    'Limpa a combo de seleção de conta corrente
    TipoDocSeleciona.ListIndex = COMBO_INDICE

    iFramePrincipalAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub TipoDocSeleciona_Change()

    iFramePrincipalAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub BotaoDesmarcarTodos_Click()
Dim iIndice As Integer
    
    For iIndice = 1 To objGridParcelas.iLinhasExistentes
        
        GridParcelas.TextMatrix(iIndice, iGrid_Selecionar_Col) = S_DESMARCADO
            
    Next
    
    Call Grid_Refresh_Checkbox(objGridParcelas)
    
    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO

    Call Calcula_Total_Pagar
    
    Exit Sub

End Sub

Public Sub BotaoMarcarTodos_Click()

Dim iIndice As Integer
    
    For iIndice = 1 To objGridParcelas.iLinhasExistentes
        
        GridParcelas.TextMatrix(iIndice, iGrid_Selecionar_Col) = S_MARCADO
            
    Next
    
    Call Grid_Refresh_Checkbox(objGridParcelas)
    
    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO

    Call Calcula_Total_Pagar
    
    Exit Sub

End Sub

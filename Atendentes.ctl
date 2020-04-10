VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl AtendentesOcx 
   ClientHeight    =   3990
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8625
   ScaleHeight     =   3990
   ScaleWidth      =   8625
   Begin VB.Frame FrameAtendente 
      Caption         =   "Atendente"
      Height          =   780
      Left            =   240
      TabIndex        =   12
      Top             =   960
      Width           =   5595
      Begin VB.CheckBox AumentaQuant 
         Caption         =   "Pode aumentar quantidades requisitadas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   870
         TabIndex        =   17
         Top             =   780
         Width           =   3855
      End
      Begin VB.CommandButton BotaoProxNum 
         Height          =   285
         Left            =   2280
         Picture         =   "Atendentes.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Numeração Automática"
         Top             =   330
         Width           =   300
      End
      Begin MSMask.MaskEdBox Codigo 
         Height          =   315
         Left            =   1635
         TabIndex        =   14
         Top             =   300
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin VB.Label LabelCodigo 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   870
         TabIndex        =   15
         Top             =   360
         Width           =   660
      End
   End
   Begin VB.Frame FrameUsuario 
      Caption         =   "Usuário"
      Height          =   1395
      Left            =   240
      TabIndex        =   7
      Top             =   1800
      Width           =   5595
      Begin VB.Label LabelCodUsuario 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   495
         TabIndex        =   11
         Top             =   360
         Width           =   660
      End
      Begin VB.Label CodUsuario 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1245
         TabIndex        =   10
         Top             =   330
         Width           =   1080
      End
      Begin VB.Label LabelNomeUsuario 
         AutoSize        =   -1  'True
         Caption         =   "Nome:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   600
         TabIndex        =   9
         Top             =   990
         Width           =   555
      End
      Begin VB.Label NomeUsuario 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1245
         TabIndex        =   8
         Top             =   930
         Width           =   3900
      End
   End
   Begin VB.CommandButton BotaoUsuarios 
      Caption         =   "Usuários"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   240
      Picture         =   "Atendentes.ctx":00EA
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3360
      Width           =   2055
   End
   Begin VB.ListBox Atendentes 
      Height          =   1815
      Left            =   6120
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   1320
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6135
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "Atendentes.ctx":0694
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "Atendentes.ctx":07EE
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "Atendentes.ctx":0978
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1590
         Picture         =   "Atendentes.ctx":0EAA
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Label LabelAtendentes 
      AutoSize        =   -1  'True
      Caption         =   "Atendentes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6105
      TabIndex        =   16
      Top             =   1080
      Width           =   975
   End
End
Attribute VB_Name = "AtendentesOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Public iAlterado As Integer

Private WithEvents objEventoUsuario As AdmEvento
Attribute objEventoUsuario.VB_VarHelpID = -1

'*** CARREGAMENTO DA TELA - INÍCIO ***
Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    iAlterado = 0

    'Inicializa o objEvento
    Set objEventoUsuario = New AdmEvento

    'Carrega a listbox de atendentes
    lErro = Carrega_Atendentes()
    If lErro <> SUCESSO Then gError 102718

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 102718

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143145)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Public Function Trata_Parametros(Optional objAtendente As ClassAtendentes) As Long

Dim lErro As Long
Dim objUsuario As ClassUsuarios

On Error GoTo Erro_Trata_Parametros

    'Se há um Atendente preenchido
    If Not (objAtendente Is Nothing) Then

        'Se o código do atendente foi passado dentro do obj
        If objAtendente.iCodigo > 0 Then

            'Lê o atendente a partir do código passado no obj
            lErro = CF("Atendentes_Le", objAtendente)
            If lErro <> SUCESSO And lErro <> 102752 Then gError 102719

            'Se o atendente existe
            If lErro = SUCESSO Then

                'Traz o atendente pra tela
                lErro = Traz_Atendente_Tela(objAtendente)
                If lErro <> SUCESSO Then gError 102720

            'Se o atendente não existe
            ElseIf objAtendente.iCodigo > 0 Then

                'Exibe o código passado como parâmetro
                Codigo.Text = CStr(objAtendente.iCodigo)

            End If
        
        'Se foi passado o nome reduzido
        ElseIf Len(Trim(objAtendente.sNomeReduzido)) > 0 Then
        
            'Instancia um objUsuario
            Set objUsuario = New ClassUsuarios
            
            'Guarda o nome reduzido no objusuario
            objUsuario.sNomeReduzido = objAtendente.sNomeReduzido
            
            'Le o Usuario
            lErro = CF("Usuarios_Le_NomeRed", objUsuario)
            If lErro <> SUCESSO And lErro <> 53206 Then gError 102821
        
            'Se não encontrou o usuário
            If lErro = 53206 Then gError 102822
            
            'Coloca o código do usuário na tela
            CodUsuario.Caption = objUsuario.sCodUsuario
            
            'Coloca o usuário na tela
            NomeUsuario.Caption = objUsuario.sNomeReduzido
            
            'Cria o código do atendente
            Call BotaoProxNum_Click

        End If

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 102719, 102720, 102821
        
        Case 102822
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_ENCONTRADO", gErr, objAtendente.sNomeReduzido)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143146)

    End Select

    iAlterado = 0

    Exit Function

End Function
'*** CARREGAMENTO DA TELA - FIM ***

'*** FECHAMENTO DA TELA - INÍCIO ***
Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Public Sub Form_Unload(Cancel As Integer)
    Set objEventoUsuario = Nothing
    
    Call ComandoSeta_Liberar(Me.Name)
End Sub
'*** FECHAMENTO DA TELA - FIM ***

'*** TRATAMENTO DOS CONTROLES DA TELA - INÍCIO****

'*** EVENTO GOTFOCUS DOS CONTROLES MASCARADOS - INÍCIO ***
Private Sub Codigo_GotFocus()
    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)
End Sub
'*** EVENTO GOTFOCUS DOS CONTROLES MASCARADOS - FIM ***

'*** EVENTO CLICK DOS CONTROLES - INÍCIO ***
Private Sub BotaoGravar_Click()

Dim lErro As Long
Dim objAtendente As New ClassAtendentes

On Error GoTo Erro_BotaoGravar_Click

    'Chama rotina de Gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 102721

    'Limpa a tela
    Call Limpa_Tela_Atendentes

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 102721

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143147)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objAtendente As New ClassAtendentes
Dim iIndice As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim objUsuarios As New ClassUsuarios

On Error GoTo Erro_BotaoExcluir_Click

    'Transforma o ponteiro do mouse em ampulheta
    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se o atendente está preenchido
    If Len(Trim(Codigo.ClipText)) = 0 Then gError 102722

    'Guarda no obj o código do atendente e a filial empresa
    objAtendente.iCodigo = StrParaInt(Codigo.ClipText)
    objAtendente.iFilialEmpresa = giFilialEmpresa

    'Lê os dados do atendente
    lErro = CF("Atendentes_Le", objAtendente)
    If lErro <> SUCESSO And lErro <> 102752 Then gError 102723

    'Se não encontrou o atendente => erro
    If lErro = 102752 Then gError 102724

    'Pede a confirmação da exclusão do comprador do usuário
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_ATENDENTE", objAtendente.sCodUsuario)

    'Se o usuário confirmou a exclusão
    If vbMsgRes = vbYes Then

        'Dispara a exclusão do atendente
        lErro = CF("Atendentes_Exclui", objAtendente)
        If lErro <> SUCESSO Then gError 102725

        'Limpa a tela de Atendentes
        Call Limpa_Tela_Atendentes

        'Para cada atendente na lista
        For iIndice = 0 To Atendentes.ListCount - 1

            'Se o conteúdo que está sendo exibido na list é igual ao nome reduzido do usuário
            If Atendentes.List(iIndice) = objAtendente.sNomeReduzido Then

                'Significa que encontrou o atendente correspondente e remove-o da lista
                Atendentes.RemoveItem iIndice

                'Sai do loop
                Exit For

            End If
        Next

        iAlterado = 0

    End If

    'Volta o ponteiro do mouse para o padrão
    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 102723, 102725, 102726
        
        Case 102722
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 102724
            Call Rotina_Erro(vbOKOnly, "ERRO_ATENDENTE_NAO_ENCONTRADO", gErr, objAtendente.iCodigo, giFilialEmpresa)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143148)

    End Select

    'Volta o ponteiro do mouse para o padrão
    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se há alterações e quer salvá-las
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 102727

    Call Limpa_Tela_Atendentes

    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 102727

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143149)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()
    Unload Me
End Sub

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_BotaoProxNum_Click

    'Obtém o próximo código de relacionamento para giFilialEmpresa
    lErro = CF("Config_Obter_Inteiro_Automatico", "CRFATConfig", "NUM_PROX_ATENDENTE", "Atendentes", "Codigo", iCodigo)
    If lErro <> SUCESSO Then gError 102728

    'Exibe o código obtido
    Codigo.PromptInclude = False
    Codigo.Text = iCodigo
    Codigo.PromptInclude = True

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 102728

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143150)

    End Select

    Exit Sub

End Sub

Private Sub BotaoUsuarios_Click()

Dim objUsuarios As New ClassUsuarios
Dim colSelecao As Collection

On Error GoTo Erro_BotaoUsuarios_Click

    'Guarda o Codigo do Usuario
    objUsuarios.sCodUsuario = CodUsuario.Caption

    'Chama a tela UsuarioLista
    Call Chama_Tela("UsuarioLista", colSelecao, objUsuarios, objEventoUsuario)

    Exit Sub

Erro_BotaoUsuarios_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143151)

    End Select

    Exit Sub

End Sub

Private Sub Atendentes_DblClick()

Dim lErro As Long
Dim objAtendente As New ClassAtendentes
Dim objUsuarios As New ClassUsuarios

On Error GoTo Erro_Atendentes_Click

    'Guarda no obj o nome do usuário clicado na lista
    objUsuarios.sNomeReduzido = Atendentes.List(Atendentes.ListIndex)

    'Le o Usuario
    lErro = CF("Usuarios_Le_NomeRed", objUsuarios)
    If lErro <> SUCESSO And lErro <> 53206 Then gError 102729

    'Se não encontrou o usuário
    If lErro = 53206 Then gError 102730

    'Guarda no obj o código do usuário lido e a filial empresa
    objAtendente.sCodUsuario = objUsuarios.sCodUsuario
    objAtendente.iFilialEmpresa = giFilialEmpresa

    'Lê os dados do atendente a partir do código do usuário
    lErro = CF("Atendentes_Le_Usuario", objAtendente)
    If lErro <> SUCESSO Then gError 102731

    'Traz o atendente para a tela
    lErro = Traz_Atendente_Tela(objAtendente)
    If lErro <> SUCESSO Then gError 102733

    Exit Sub

Erro_Atendentes_Click:

    Select Case gErr

        Case 102729, 102731, 102733

        Case 102730
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_CADASTRADO2", gErr, objUsuarios.sNomeReduzido)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143152)

    End Select

    Exit Sub

End Sub
'*** EVENTO CLICK DOS CONTROLES - FIM ***

'*** EVENTO CHANGE DOS CONTROLES - INÍCIO ***
Private Sub Codigo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub
'*** EVENTO CHANGE DOS CONTROLES - FIM ***

'*** EVENTO VALIDATE DOS CONTROLES - INÍCIO ***
Private Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objAtendente As New ClassAtendentes

On Error GoTo Erro_Codigo_Validate

    'Se o código não está preenchido => sai da função
    If Len(Trim(Codigo.ClipText)) = 0 Then Exit Sub

    'Guarda no obj o código utilizado e a filial empresa
    objAtendente.iCodigo = Codigo.Text
    objAtendente.iFilialEmpresa = giFilialEmpresa

    'Lê o atendente com os dados passados por parâmetros
    lErro = CF("Atendentes_Le", objAtendente)
    If lErro <> SUCESSO And lErro <> 102752 Then gError 102734

    'Se encontrou o atendente
    If lErro = SUCESSO Then
    
        'Exibe o atendente na tela
        lErro = Traz_Atendente_Tela(objAtendente)
        If lErro <> SUCESSO Then gError 102735
        
    End If

    'Fecha o comando de setas
    Call ComandoSeta_Fechar(Me.Name)

    Exit Sub

Erro_Codigo_Validate:

    Cancel = True

    Select Case gErr

        Case 102734, 102735

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143153)

    End Select

End Sub
'*** EVENTO VALIDATE DOS CONTROLES - FIM ***

'*** TRATAMENTO DO EVENTO KEYDOWN  - INÍCIO ***
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_PROXIMO_NUMERO Then
        Call BotaoProxNum_Click
    End If

End Sub
'*** TRATAMENTO DO EVENTO KEYDOWN  - FIM ***

'*** TRATAMENTO DOS EVENTOS DE BROWSER - INÍCIO ***
Private Sub ObjEventoUsuario_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objAtendente As New ClassAtendentes
Dim objUsuarios As ClassUsuarios
Dim iCodigo As Integer

On Error GoTo Erro_ObjEventoUsuario_evSelecao

    'Instancia um novo objUsuario
    Set objUsuarios = obj1

    'Guarda no obj o código do usuário e a filial empresa ativa
    objAtendente.sCodUsuario = objUsuarios.sCodUsuario
    objAtendente.iFilialEmpresa = giFilialEmpresa

    'Coloca os dados do usuário na tela
    CodUsuario.Caption = objUsuarios.sCodUsuario
    NomeUsuario.Caption = objUsuarios.sNomeReduzido

    'Lê o atendente correspondente ao usuario
    lErro = CF("Atendentes_Le_Usuario", objAtendente)
    If lErro <> SUCESSO And lErro <> 102756 Then gError 102736

    'Se encontrou o atendente
    If lErro = SUCESSO Then
        Codigo.PromptInclude = False
        Codigo.Text = objAtendente.iCodigo
        Codigo.PromptInclude = True
    End If

    'Fecha o comando de setas, se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    Me.Show

    Exit Sub

Erro_ObjEventoUsuario_evSelecao:

    Select Case gErr

        Case 102736

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143154)

    End Select

    Exit Sub

End Sub
'*** TRATAMENTO DOS EVENTOS DE BROWSER - FIM ***

'**** TRATAMENTO DO SISTEMA DE SETAS - INÍCIO ****
Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objAtendente As New ClassAtendentes

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à tela
    sTabela = "Atendentes"

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "CodUsuario", CodUsuario.Caption, STRING_USUARIO_CODIGO, "CodUsuario"
    colCampoValor.Add "Codigo", StrParaInt(Codigo.Text), 0, "Codigo"

    'Filtro
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa

    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143155)

    End Select

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim objAtendente As New ClassAtendentes
Dim lErro As Long

On Error GoTo Erro_Tela_Preenche

    'Passa os dados da colecao de campos-valores para o objAtendente
    objAtendente.sCodUsuario = colCampoValor.Item("CodUsuario").vValor
    objAtendente.iCodigo = colCampoValor.Item("Codigo").vValor

    'Se o código do atendente foi preenchido
    If objAtendente.iCodigo <> 0 Then

        'Exibe na tela os dados do atendente
        lErro = Traz_Atendente_Tela(objAtendente)
        If lErro <> SUCESSO Then gError 102737

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 102737

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143156)

    End Select

End Sub

Public Sub Form_Activate()
    Call TelaIndice_Preenche(Me)
End Sub

Public Sub Form_Deactivate()
    gi_ST_SetaIgnoraClick = 1
End Sub
'**** TRATAMENTO DO SISTEMA DE SETAS - FIM ****

'*** FUNÇÕES DE APOIO À TELA - INÍCIO ***
Private Function Traz_Atendente_Tela(objAtendente As ClassAtendentes) As Long
'Traz os dados do atendente para a tela
'objAtendente RECEBE(Input) os dados que servirão para identificar o atendente a ser trazido para a tela

Dim iIndice As Integer
Dim lErro As Long
Dim objUsuarios As New ClassUsuarios

On Error GoTo Erro_Traz_Atendente_Tela

    'Guarda no obj o código do usuário
    objUsuarios.sCodUsuario = objAtendente.sCodUsuario

    'Lê o usuário no BD
    lErro = CF("Usuarios_Le", objUsuarios)
    If lErro <> SUCESSO And lErro <> 40832 Then gError 102738
    
    'Se não encontrou => erro
    If lErro = 40832 Then gError 102739

    'Preenche a tela com os dados de objAtendente
    CodUsuario.Caption = objAtendente.sCodUsuario
    NomeUsuario.Caption = objUsuarios.sNomeReduzido
    
    Codigo.PromptInclude = False
    Codigo.Text = objAtendente.iCodigo
    Codigo.PromptInclude = True
    
    iAlterado = 0

    Traz_Atendente_Tela = SUCESSO

    Exit Function

Erro_Traz_Atendente_Tela:

    Traz_Atendente_Tela = gErr

    Select Case gErr

        Case 102738

        Case 102739
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_ENCONTRADO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143157)

    End Select

End Function

Private Sub Limpa_Tela_Atendentes()
'Limpa a tela

    NomeUsuario.Caption = ""
    CodUsuario.Caption = ""
    
    Codigo.PromptInclude = False
    Codigo.Text = ""
    Codigo.PromptInclude = True

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long
'Grava um registro no bd

Dim lErro As Long
Dim objAtendente As New ClassAtendentes
Dim objUsuarios As New ClassUsuarios

On Error GoTo Erro_Gravar_Registro

    'Transforma o ponteiro do mouse em ampulheta
    GL_objMDIForm.MousePointer = vbHourglass

    'Se não foi selecionado um usuário
    If Len(CodUsuario.Caption) = 0 Then gError 102740
    If Len(NomeUsuario.Caption) = 0 Then gError 102741
    
    'Se não foi informado o código do atendente
    If Len(Codigo.ClipText) = 0 Then gError 102742

    'Transfere os dados da tela para os obj's
    objAtendente.sCodUsuario = CodUsuario.Caption
    objAtendente.sNomeReduzido = NomeUsuario.Caption
    objAtendente.iCodigo = StrParaInt(Codigo.ClipText)
    objAtendente.iFilialEmpresa = giFilialEmpresa

    'Lê o atendente com o usuário da tela
    lErro = CF("Atendentes_Le_Usuario", objAtendente)
    If lErro <> SUCESSO And lErro <> 102756 Then gError 102743

    'Se encontrou
    If lErro = SUCESSO Then

        'Verifica se o codigo e o mesmo que o codigo da tela
        If (objAtendente.iCodigo <> StrParaInt(Codigo.ClipText)) Then gError 102744

    End If

    'Grava o atendente
    lErro = CF("Atendentes_Grava", objAtendente)
    If lErro <> SUCESSO Then gError 102746

    'Adiciona na listbox se necessário
    Call Adiciona_Lista_Atendentes(objAtendente)

    'Retorna o ponteiro padrão do mouse
    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    Select Case gErr

        Case 102743, 102746

        Case 102740, 102741
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_PREENCHIDO", gErr)

        Case 102742
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 102744
            Call Rotina_Erro(vbOKOnly, "ERRO_ATENDENTE_USUARIO", gErr, objAtendente.sCodUsuario, objAtendente.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143158)

    End Select

    'Retorna o ponteiro padrão do mouse
    GL_objMDIForm.MousePointer = vbDefault

    Exit Function

End Function

Private Sub Adiciona_Lista_Atendentes(objAtendente As ClassAtendentes)
'Adiciona um atendente na ListBox

    'Se o nome reduzido do atendente está preenchido no obj
    If Len(Trim(objAtendente.sNomeReduzido)) > 0 Then
        
        'Se ele é novo adiciona-o na lista
        Atendentes.AddItem objAtendente.sNomeReduzido
    
    End If

End Sub

Private Function Carrega_Atendentes() As Long
'Carrega a ListBox

Dim lErro As Long
Dim objAtendente As New ClassAtendentes
Dim colUsuarios As New Collection
Dim objUsuarios As New ClassUsuarios
Dim colAtendentes As New Collection

On Error GoTo Erro_Carrega_Atendentes

    'Le todos os Atendentes da Filial Empresa
    lErro = CF("Atendentes_Le_Todos", colAtendentes)
    If lErro <> SUCESSO And lErro <> 102761 Then gError 102748

    'Para cada atendente encontrado
    For Each objAtendente In colAtendentes
        'Adiciona o usuário na lista de atendentes
        Atendentes.AddItem objAtendente.sNomeReduzido
    Next

    Carrega_Atendentes = SUCESSO

    Exit Function

Erro_Carrega_Atendentes:

    Carrega_Atendentes = gErr

    Select Case gErr

        Case 102747, 102748

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143159)

    End Select

End Function

'***************************************************
'Início do trecho de codigo comum as telas
'***************************************************
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Atendentes"
    Call Form_Load

End Function

Public Function Name() As String
    Name = "Atendentes"
End Function

Public Sub Show()
    Parent.Show
    Parent.SetFocus
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Controls
Public Property Get Controls() As Object
    Set Controls = UserControl.Controls
End Property

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Property Get Height() As Long
    Height = UserControl.Height
End Property

Public Property Get Width() As Long
    Width = UserControl.Width
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ActiveControl
Public Property Get ActiveControl() As Object
    Set ActiveControl = UserControl.ActiveControl
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
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
'***************************************************
'Fim Trecho de codigo comum as telas
'***************************************************

'*** TRATAMENTO DE DRAG AND DROP / MOUSEDOWN DOS LABELS - INÍCIO ***
Private Sub LabelCodigo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodigo, Source, X, Y)
End Sub

Private Sub LabelCodigo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodigo, Button, Shift, X, Y)
End Sub

Private Sub LabelCodUsuario_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodUsuario, Source, X, Y)
End Sub

Private Sub LabelCodUsuario_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodUsuario, Button, Shift, X, Y)
End Sub

Private Sub LabelNomeUsuario_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNomeUsuario, Source, X, Y)
End Sub

Private Sub LabelNomeUsuario_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNomeUsuario, Button, Shift, X, Y)
End Sub

Private Sub LabelAtendentes_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelAtendentes, Source, X, Y)
End Sub

Private Sub LabelAtendentes_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelAtendentes, Button, Shift, X, Y)
End Sub

Private Sub CodUsuario_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CodUsuario, Source, X, Y)
End Sub

Private Sub CodUsuario_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CodUsuario, Button, Shift, X, Y)
End Sub

Private Sub NomeUsuario_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NomeUsuario, Source, X, Y)
End Sub

Private Sub NomeUsuario_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NomeUsuario, Button, Shift, X, Y)
End Sub
'*** TRATAMENTO DE DRAG AND DROP / MOUSEDOWN DOS LABELS - FIM ***

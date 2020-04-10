VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl RelOpClientesSemRelacOcx 
   ClientHeight    =   6255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7620
   ScaleHeight     =   6255
   ScaleWidth      =   7620
   Begin VB.CheckBox AnalisarFiliaisClientes 
      Caption         =   "Analisar por Filiais"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   5880
      Width           =   2055
   End
   Begin VB.Frame FrameCategoriaCliente 
      Caption         =   "Categoria"
      Height          =   1470
      Left            =   240
      TabIndex        =   25
      Top             =   4320
      Width           =   4980
      Begin VB.ComboBox CategoriaCliente 
         Height          =   315
         Left            =   1395
         TabIndex        =   9
         Top             =   540
         Width           =   2745
      End
      Begin VB.ComboBox CategoriaClienteDe 
         Height          =   315
         Left            =   705
         TabIndex        =   10
         Top             =   1020
         Width           =   1740
      End
      Begin VB.CheckBox CategoriaClienteTodas 
         Caption         =   "Todas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   300
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
      Begin VB.ComboBox CategoriaClienteAte 
         Height          =   315
         Left            =   3030
         TabIndex        =   11
         Top             =   1005
         Width           =   1740
      End
      Begin VB.Label Label5 
         Caption         =   "Label5"
         Height          =   15
         Left            =   360
         TabIndex        =   29
         Top             =   720
         Width           =   30
      End
      Begin VB.Label LabelCategoriaClienteAte 
         AutoSize        =   -1  'True
         Caption         =   "Até:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2595
         TabIndex        =   28
         Top             =   1065
         Width           =   360
      End
      Begin VB.Label LabelCategoriaClienteDe 
         AutoSize        =   -1  'True
         Caption         =   "De:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   330
         TabIndex        =   27
         Top             =   1065
         Width           =   315
      End
      Begin VB.Label LabelCategoriaCliente 
         Caption         =   "Categoria:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   480
         TabIndex        =   26
         Top             =   585
         Width           =   855
      End
   End
   Begin VB.Frame FrameTipoRelacionamento 
      Caption         =   "Tipo de Relacionamento"
      Height          =   1095
      Left            =   240
      TabIndex        =   20
      Top             =   1920
      Width           =   4965
      Begin VB.ComboBox TipoRelacionamento 
         Height          =   315
         Left            =   1245
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   600
         Width           =   2550
      End
      Begin VB.OptionButton TipoRelacApenas 
         Caption         =   "Apenas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   180
         TabIndex        =   3
         Top             =   615
         Width           =   1050
      End
      Begin VB.OptionButton TipoRelacTodos 
         Caption         =   "Todos os tipos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   195
         TabIndex        =   2
         Top             =   285
         Width           =   1620
      End
   End
   Begin VB.Frame FrameTipoCliente 
      Caption         =   "Tipo de Cliente"
      Height          =   1095
      Left            =   240
      TabIndex        =   24
      Top             =   3120
      Width           =   4965
      Begin VB.OptionButton TipoClienteTodos 
         Caption         =   "Todos os tipos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   195
         TabIndex        =   5
         Top             =   285
         Width           =   1620
      End
      Begin VB.OptionButton TipoClienteApenas 
         Caption         =   "Apenas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   180
         TabIndex        =   6
         Top             =   615
         Width           =   1050
      End
      Begin VB.ComboBox TipoCliente 
         Height          =   315
         Left            =   1245
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   585
         Width           =   2550
      End
   End
   Begin VB.Frame FrameDias 
      Caption         =   "Clientes sem relacionamentos há"
      Height          =   735
      Left            =   240
      TabIndex        =   21
      Top             =   1080
      Width           =   4935
      Begin MSMask.MaskEdBox NumDias 
         Height          =   315
         Left            =   3120
         TabIndex        =   1
         Top             =   300
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   2
         Mask            =   "##"
         PromptChar      =   " "
      End
      Begin VB.Label LabelNumDias 
         AutoSize        =   -1  'True
         Caption         =   "Clientes sem relacionamentos há:"
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
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   2850
      End
      Begin VB.Label LabelDias 
         AutoSize        =   -1  'True
         Caption         =   "dias"
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
         Left            =   3720
         TabIndex        =   22
         Top             =   360
         Width           =   360
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5250
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   165
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpClientesSemRelacOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpClientesSemRelacOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpClientesSemRelacOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpClientesSemRelacOcx.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.CommandButton BotaoExecutar 
      Caption         =   "Executar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   5445
      Picture         =   "RelOpClientesSemRelacOcx.ctx":0994
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   795
      Width           =   1815
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpClientesSemRelacOcx.ctx":0A96
      Left            =   1050
      List            =   "RelOpClientesSemRelacOcx.ctx":0A98
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   285
      Width           =   2730
   End
   Begin VB.Label Label1 
      Caption         =   "Opção:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   19
      Top             =   330
      Width           =   615
   End
End
Attribute VB_Name = "RelOpClientesSemRelacOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Se a opção AnalisarFiliaisClientes estiver MARCADA e houver uma categoria selecionada, chamar o relatório FCLSRLCT
'Se a opção AnalisarFiliaisClientes estiver MARCADA e NÃO houver uma categoria selecionada, chamar o relatório FCLSRL
'Se a opção AnalisarFiliaisClientes estiver DESMARCADA, chamar o relatório CLSRL'Verificar a tela RelacClientesOcx em RelatoriosFAT2 para fazer o tratamento do frame TipoRelacionamento
'Se a opção AnalisarFiliaisClientes estiver DESMARCADA, os controles do frame CategoriaCliente devem estar desabilitados. Ao marcar essa opção, os mesmos devem ser habilitados
'Verificar a tela RelOpTitRecOcx em RelatoriosCPR2 para fazer o tratamento do frame TipoCliente
'Verificar a tela RelOpEstoqueVendasOcx em RelatoriosFAT2 para fazer o tratamento do frame Categoria
'O campo NumDias, deve ser passado na expressão de seleção como DataBase<=
'Para encontrar essa data base, deve ser utilizada a data atual - o número de dias informado pelo usuário


Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    'Carrega a combo Tipo Relacionamento
    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_TIPORELACIONAMENTOCLIENTES, TipoRelacionamento)
    If lErro <> SUCESSO Then gError 131592

    'Carrega a combo Tipo Cliente
    Call Carrega_ComboTipoCliente(TipoCliente)
    
    Call Carrega_ComboCategoriaCliente(CategoriaCliente)

    'Define os Campos
    lErro = Define_Padrao()
    If lErro <> SUCESSO Then gError 131593

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case 131592 To 131593

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167593)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 131594

    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche a Combo Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 131595

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 131594
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)

        Case 131595

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167594)

    End Select

    Exit Function

End Function

Private Sub AnalisarFiliaisClientes_Click()

    If AnalisarFiliaisClientes.Value = vbChecked Then
        CategoriaClienteTodas.Enabled = True
    Else
        CategoriaClienteTodas.Enabled = False
    End If

    CategoriaClienteTodas.Value = vbChecked
    Call CategoriaClienteTodas_Click

End Sub

Private Sub CategoriaClienteAte_Validate(Cancel As Boolean)

    Call CategoriaClienteItem_Validate(Cancel, CategoriaClienteAte)

End Sub


Private Sub CategoriaClienteDe_Validate(Cancel As Boolean)
    
    Call CategoriaClienteItem_Validate(Cancel, CategoriaClienteDe)

End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Public Sub TipoRelacionamento_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_TipoRelacionamento_Validate

    'Valida o tipo de relacionamento selecionado pelo cliente
    lErro = CF("CamposGenericos_Validate", CAMPOSGENERICOS_TIPORELACIONAMENTOCLIENTES, TipoRelacionamento, "AVISO_CRIAR_TIPORELACIONAMENTOCLIENTES")
    If lErro <> SUCESSO Then gError 131596

    Exit Sub

Erro_TipoRelacionamento_Validate:

    Cancel = True

    Select Case gErr

        Case 131596

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 167595)

    End Select

End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)

End Sub

Private Sub CategoriaCliente_Validate(Cancel As Boolean)

Dim lErro As Long, iCodigo As Integer

On Error GoTo Erro_CategoriaCliente_Validate

    If Len(CategoriaCliente.Text) <> 0 And CategoriaCliente.ListIndex = -1 Then
    
        'pesquisa a categoria na lista
        lErro = Combo_Item_Igual(CategoriaCliente)
        If lErro <> SUCESSO And lErro <> 12253 Then gError 131623
        
        If lErro <> SUCESSO Then gError 131624
    
    End If
    
    'Se a CategoriaCliente estiver em branco desabilita e limpa a combo
    If Len(CategoriaCliente.Text) = 0 Then
        CategoriaClienteDe.Enabled = False
        CategoriaClienteDe.Clear
        CategoriaClienteAte.Enabled = False
        CategoriaClienteAte.Clear
    End If
    
    Exit Sub

Erro_CategoriaCliente_Validate:
    
    Cancel = True
    
    Select Case gErr

        Case 131623
         
        Case 131624
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIA_NAO_CADASTRADA", gErr, CategoriaCliente.Text)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 167596)

    End Select

    Exit Sub

End Sub

Private Sub CategoriaClienteItem_Validate(Cancel As Boolean, objCombo As ComboBox)

Dim lErro As Long

On Error GoTo Erro_CategoriaClienteItem_Validate

    If Len(objCombo.Text) <> 0 Then
    
        'pesquisa o item na lista
        lErro = Combo_Item_Igual(objCombo)
        If lErro <> SUCESSO And lErro <> 12253 Then gError 131625
        
        If lErro <> SUCESSO Then gError 131626
    
    End If

    Exit Sub

Erro_CategoriaClienteItem_Validate:

    Cancel = True

    Select Case gErr

        Case 131625
        
        Case 131626
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIACLIENTEITEM_INEXISTENTE", gErr, objCombo.Text)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 167597)

    End Select

    Exit Sub

End Sub

Private Sub TipoRelacTodos_Click()

Dim lErro As Long

On Error GoTo Erro_TipoRelacTodos_Click

    'Desabilita o combotipo
    TipoRelacionamento.ListIndex = -1
    TipoRelacionamento.Enabled = False

    Exit Sub

Erro_TipoRelacTodos_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167598)

    End Select

    Exit Sub

End Sub

Private Sub TipoRelacApenas_Click()

Dim lErro As Long

On Error GoTo Erro_TipoRelacApenas_Click

    'Habilita a ComboTipo
    TipoRelacionamento.Enabled = True

    Exit Sub

Erro_TipoRelacApenas_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167599)

    End Select

    Exit Sub

End Sub

Private Sub CategoriaClienteTodas_Click()

Dim lErro As Long

On Error GoTo Erro_CategoriaClienteTodas_Click

    If CategoriaClienteTodas.Value = vbChecked Then
        'Desabilita o combotipo
        CategoriaCliente.ListIndex = -1
        CategoriaCliente.Enabled = False
        CategoriaClienteDe.Clear
        CategoriaClienteAte.Clear
    Else
        CategoriaCliente.Enabled = True
    End If

    Call CategoriaCliente_Click

    Exit Sub

Erro_CategoriaClienteTodas_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167600)

    End Select

    Exit Sub

End Sub

Private Sub TipoClienteTodos_Click()

Dim lErro As Long

On Error GoTo Erro_TipoClienteTodos_Click

    'Desabilita o combotipo
    TipoCliente.ListIndex = -1
    TipoCliente.Enabled = False

    Exit Sub

Erro_TipoClienteTodos_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167601)

    End Select

    Exit Sub

End Sub

Private Sub TipoClienteApenas_Click()

Dim lErro As Long

On Error GoTo Erro_TipoClienteApenas_Click

    'Habilita a ComboTipo
    TipoCliente.Enabled = True

    Exit Sub

Erro_TipoClienteApenas_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167602)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    'Faz a chamada da função que irá realizar o preenchimento do objeto RelOpcoes
    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 131597

    If AnalisarFiliaisClientes.Value <> vbChecked Then
        gobjRelatorio.sNomeTsk = "CLSRL"
    Else
        If CategoriaClienteTodas.Value = vbChecked Then
            gobjRelatorio.sNomeTsk = "FCLSRL"
        Else
            gobjRelatorio.sNomeTsk = "FCLSRLCT"
        End If
    End If

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 131597

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167603)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

     'Limpa a tela
    Call LimpaRelatorioClientesSemRelac

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167604)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 131598

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_CLIENTESEMRELAC")

    If vbMsgRes = vbYes Then

        'Exclui o elemento do banco de dados
        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 131599

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'Limpa a tela
        lErro = LimpaRelatorioClientesSemRelac()
        If lErro <> SUCESSO Then gError 131600

    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 131598
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 131599, 131600

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167605)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 131601

    'Faz a chamada da função que irá realizar o preenchimento do objeto RelOpcoes
    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 131602

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    'Grava no banco de dados
    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 131603

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 131604

    'Limpa a tela
    lErro = LimpaRelatorioClientesSemRelac()
    If lErro <> SUCESSO Then gError 131605

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 131601
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 131602 To 131605

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167606)

    End Select

    Exit Sub

End Sub

'*** FUNÇÕES DE APOIO À TELA - INÍCIO ***
Private Function Define_Padrao() As Long
'Preenche as datas e carrega as combos da tela

Dim lErro As Long

On Error GoTo Erro_Define_Padrao

    'defina todos os tipos
    TipoRelacTodos.Value = True
    TipoRelacionamento.Enabled = False

    TipoClienteTodos.Value = True
    TipoCliente.Enabled = False
    
    CategoriaClienteTodas.Value = vbChecked
    CategoriaCliente.Enabled = False
    CategoriaClienteDe.Enabled = False
    CategoriaClienteAte.Enabled = False
    CategoriaClienteDe.ListIndex = -1
    CategoriaClienteAte.ListIndex = -1
    
    AnalisarFiliaisClientes.Value = vbChecked

    Define_Padrao = SUCESSO

    Exit Function

Erro_Define_Padrao:

    Define_Padrao = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167607)

    End Select

    Exit Function

End Function

Private Function LimpaRelatorioClientesSemRelac()
'Limpa a tela RelOpRelacClientes

Dim lErro As Long

On Error GoTo Erro_LimpaRelatorioClientesSemRelac

    'Limpa os Campos
    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 131606

    ComboOpcoes.Text = ""

    'Define os Campos
    lErro = Define_Padrao()
    If lErro <> SUCESSO Then gError 131607

    LimpaRelatorioClientesSemRelac = SUCESSO

    Exit Function

Erro_LimpaRelatorioClientesSemRelac:

    LimpaRelatorioClientesSemRelac = gErr

    Select Case gErr

        Case 131606, 131607

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167608)

    End Select

    Exit Function

End Function

Private Sub Carrega_ComboTipoCliente(ByVal objComboBox As ComboBox)

Dim lErro As Long
Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodigoDescricao As AdmCodigoNome

On Error GoTo Erro_Carrega_ComboTipoCliente

    'Lê cada código e descrição da tabela TiposDeCliente
    lErro = CF("Cod_Nomes_Le", "TiposDeCliente", "Codigo", "Descricao", STRING_TIPO_CLIENTE_DESCRICAO, colCodigoDescricao)
    If lErro <> SUCESSO Then gError 131608

    'Preenche a ComboBox Tipo com os objetos da colecao colCodigoDescricao
    For Each objCodigoDescricao In colCodigoDescricao
        objComboBox.AddItem CStr(objCodigoDescricao.iCodigo) & SEPARADOR & objCodigoDescricao.sNome
        objComboBox.ItemData(objComboBox.NewIndex) = objCodigoDescricao.iCodigo
    Next

    Exit Sub

Erro_Carrega_ComboTipoCliente:

    Select Case gErr

        Case 131608

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 167609)

    End Select

    Exit Sub

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o objRelOp com os dados fornecidos pelo usuário

Dim lErro As Long

Dim iIndice As Integer
Dim sDataBase As String
Dim sCategoria As String
Dim sCategoria_De As String
Dim sCategoria_Ate As String
Dim sTipoCliente As String
Dim sTipoRelac As String

On Error GoTo Erro_PreencherRelOp

    'Critica os valores preenchidos pelo usuário
    lErro = Formata_E_Critica_Parametros(sCategoria, sCategoria_De, sCategoria_Ate, sTipoCliente, sTipoRelac, sDataBase)
    If lErro <> SUCESSO Then gError 131609

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 131610

    'Inclui a data inicial
    lErro = objRelOpcoes.IncluirParametro("NDIAS", NumDias.Text)
    If lErro <> AD_BOOL_TRUE Then gError 131611

    'Inclui o tipo
    lErro = objRelOpcoes.IncluirParametro("TTIPORELACIONAMENTO", sTipoRelac)
    If lErro <> AD_BOOL_TRUE Then gError 131612

    lErro = objRelOpcoes.IncluirParametro("TTIPOCLIENTE", sTipoCliente)
    If lErro <> AD_BOOL_TRUE Then gError 131613

    lErro = objRelOpcoes.IncluirParametro("TCATEGORIA", sCategoria)
    If lErro <> AD_BOOL_TRUE Then gError 131627

    lErro = objRelOpcoes.IncluirParametro("TCATEGORIADE", sCategoria_De)
    If lErro <> AD_BOOL_TRUE Then gError 131628
    
    lErro = objRelOpcoes.IncluirParametro("TCATEGORIAATE", sCategoria_Ate)
    If lErro <> AD_BOOL_TRUE Then gError 131629

    lErro = objRelOpcoes.IncluirParametro("NANALFILCLI", AnalisarFiliaisClientes.Value)
    If lErro <> AD_BOOL_TRUE Then gError 131629

    'Faz a chamada da função que irá montar a expressão
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sCategoria, sCategoria_De, sCategoria_Ate, sTipoCliente, sTipoRelac, sDataBase)
    If lErro <> SUCESSO Then gError 131614

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 131609 To 131614, 131627 To 131629

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167610)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros(sCategoria As String, sCategoria_De As String, sCategoria_Ate As String, sTipoCliente As String, sTipoRelac As String, sDataBase As String) As Long
'Verifica se os parâmetros iniciais são maiores que os finais

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Formata_E_Critica_Parametros

    If Len(Trim(NumDias.ClipText)) = 0 Then gError 131630
    
    sDataBase = CStr(DateAdd("d", StrParaInt(NumDias.ClipText) * -1, gdtDataAtual))

    'Se a opção para todos os tipos estiver selecionada
    If TipoRelacTodos.Value = True Then
        sTipoRelac = ""
    Else
        If TipoRelacionamento.Text = "" Then gError 131615
        sTipoRelac = TipoRelacionamento.Text
    End If

    'Se a opção para todos os tipos estiver selecionada
    If TipoClienteTodos.Value = True Then
        sTipoCliente = ""
    Else
        If TipoCliente.Text = "" Then gError 131616
        sTipoCliente = TipoCliente.Text
    End If

    'Se a opção para todos os tipos estiver selecionada
    If CategoriaClienteTodas.Value = vbChecked Then
        sCategoria = ""
        sCategoria_De = ""
        sCategoria_Ate = ""
    Else
        If CategoriaCliente.Text = "" Then gError 131629
        sCategoria = CategoriaCliente.Text
        sCategoria_De = CategoriaClienteDe.Text
        sCategoria_Ate = CategoriaClienteAte.Text
    End If
    
    If sCategoria_De > sCategoria_Ate Then gError 131634

    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr

        Case 131615
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NAO_PREENCHIDO1", gErr)
            TipoRelacionamento.SetFocus

        Case 131616
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NAO_PREENCHIDO1", gErr)
            TipoCliente.SetFocus
            
        Case 131629
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIACLIENTE_NAO_INFORMADA", gErr)
            CategoriaCliente.SetFocus
            
        Case 131630
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMDIA_NAO_PREENCHIDO", gErr)
            NumDias.SetFocus
            
        Case 131634
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIACLIENTE_ITEM_INICIAL_MAIOR", gErr)
            CategoriaClienteDe.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167611)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sCategoria As String, sCategoria_De As String, sCategoria_Ate As String, sTipoCliente As String, sTipoRelac As String, sDataBase As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Monta_Expressao_Selecao

    'Verifica se o Cliente Inicial foi preenchido
    If sCategoria_De <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "CategoriaCliente >= " & Forprint_ConvTexto(sCategoria_De)

    End If

    'Verifica se o Cliente Final foi preenchido
    If sCategoria_Ate <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "CategoriaCliente<= " & Forprint_ConvTexto(sCategoria_Ate)

    End If

    'Verifica se a data inicial foi preenchida
    If sDataBase <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "DataBase <= " & Forprint_ConvData(StrParaDate(sDataBase))

    End If

    'Se a opção para apenas um tipo estiver selecionada
    If sTipoRelac <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "TipoRelacionamento = " & Forprint_ConvInt(Codigo_Extrai(sTipoRelac))

    End If

    'Se a opção para apenas um tipo estiver selecionada
    If sTipoCliente <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "TipoCliente = " & Forprint_ConvInt(Codigo_Extrai(sTipoCliente))

    End If

    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167612)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros armazenados no bd e exibe na tela

Dim iTipo As Integer
Dim lErro As Long
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then gError 131617

    'Preenche a data inicial
    lErro = objRelOpcoes.ObterParametro("NDIAS", sParam)
    If lErro <> SUCESSO Then gError 131618

    NumDias.Text = sParam

    lErro = objRelOpcoes.ObterParametro("TTIPORELACIONAMENTO", sParam)
    If lErro <> SUCESSO Then gError 131619

    'Preenche o tipo
    If sParam = "" Then
        TipoRelacionamento.ListIndex = -1
        TipoRelacionamento.Enabled = False
        TipoRelacTodos.Value = True
    Else
        TipoRelacApenas.Value = True
        TipoRelacionamento.Enabled = True
        Call Combo_Seleciona_ItemData(TipoRelacionamento, Codigo_Extrai(sParam))
    End If

    lErro = objRelOpcoes.ObterParametro("TTIPOCLIENTE", sParam)
    If lErro <> SUCESSO Then gError 131620

    'Preenche o tipo
    If sParam = "" Then
        TipoCliente.ListIndex = -1
        TipoCliente.Enabled = False
        TipoClienteTodos.Value = True
    Else
        TipoClienteApenas.Value = True
        TipoCliente.Enabled = True
        Call Combo_Seleciona_ItemData(TipoCliente, Codigo_Extrai((sParam)))
    End If
    
    'Prenche CheckBox Analisar Filiais
    lErro = objRelOpcoes.ObterParametro("NANALFILCLI", sParam)
    If lErro <> SUCESSO Then gError 131630
    
    AnalisarFiliaisClientes.Value = StrParaInt(sParam)
    
    Call AnalisarFiliaisClientes_Click
    
    'Prenche Categoria
    lErro = objRelOpcoes.ObterParametro("TCATEGORIA", sParam)
    If lErro <> SUCESSO Then gError 131631
    
    CategoriaCliente.Text = sParam
    Call CategoriaCliente_Validate(bSGECancelDummy)
    
    If Len(Trim(sParam)) > 0 Then
        CategoriaClienteTodas.Value = vbFalse
    Else
        CategoriaClienteTodas.Value = vbChecked
    End If
    
    'Prenche Categoria
    lErro = objRelOpcoes.ObterParametro("TCATEGORIAATE", sParam)
    If lErro <> SUCESSO Then gError 131632
    
    CategoriaClienteAte.Text = sParam
    Call CategoriaClienteAte_Validate(bSGECancelDummy)

    'Prenche Categoria
    lErro = objRelOpcoes.ObterParametro("TCATEGORIADE", sParam)
    If lErro <> SUCESSO Then gError 131633
    
    CategoriaClienteDe.Text = sParam
    Call CategoriaClienteDe_Validate(bSGECancelDummy)

    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 131617 To 131620, 131630 To 131633

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167613)

    End Select

    Exit Function

End Function
'*** FUNÇÕES DE APOIO À TELA - FIM ***

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing

End Sub

Private Sub CategoriaCliente_Click()

Dim lErro As Long

On Error GoTo Erro_CategoriaCliente_Click

    If Len(Trim(CategoriaCliente.Text)) > 0 Then
        CategoriaClienteDe.Enabled = True
        CategoriaClienteAte.Enabled = True
        Call Carrega_ComboCategoriaItens(CategoriaCliente, CategoriaClienteDe)
        Call Carrega_ComboCategoriaItens(CategoriaCliente, CategoriaClienteAte)
    Else
        CategoriaClienteDe.Enabled = False
        CategoriaClienteAte.Enabled = False
    End If


    Exit Sub

Erro_CategoriaCliente_Click:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167614)

    End Select

    Exit Sub

End Sub

Private Sub Carrega_ComboCategoriaCliente(ByVal objCombo As ComboBox)

Dim lErro As Long
Dim colCategoriaCliente As New Collection
Dim objCategoriaCliente As New ClassCategoriaCliente

On Error GoTo Erro_Carrega_ComboCategoriaCliente

    'Le as categorias de cliente
    lErro = CF("CategoriaCliente_Le_Todos", colCategoriaCliente)
    If lErro <> SUCESSO Then gError 131621

    'Preenche CategoriaCliente
    For Each objCategoriaCliente In colCategoriaCliente

        objCombo.AddItem objCategoriaCliente.sCategoria

    Next
    
    Exit Sub

Erro_Carrega_ComboCategoriaCliente:

    Select Case gErr
    
        Case 131621

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 167615)

    End Select

    Exit Sub

End Sub

Private Sub Carrega_ComboCategoriaItens(ByVal objComboCategoria As ComboBox, ByVal objComboItens As ComboBox)

Dim lErro As Long
Dim iIndice As Integer
Dim objCategoriaClienteItem As ClassCategoriaClienteItem
Dim objCategoriaCliente As New ClassCategoriaCliente
Dim colCategoria As New Collection

On Error GoTo Erro_Carrega_ComboCategoriaItens

    'Verifica se a CategoriaCliente foi preenchida
    If objComboCategoria.ListIndex <> -1 Then

        objCategoriaCliente.sCategoria = objComboCategoria.Text

        'Lê os dados de Itens da Categoria do Cliente
        lErro = CF("CategoriaCliente_Le_Itens", objCategoriaCliente, colCategoria)
        If lErro <> SUCESSO Then gError 131622

        objComboItens.Enabled = True

        'Limpa os dados de ItemCategoriaCliente
        objComboItens.Clear

        'Preenche ItemCategoriaCliente
        For Each objCategoriaClienteItem In colCategoria

            objComboItens.AddItem objCategoriaClienteItem.sItem

        Next
        
        CategoriaClienteTodas.Value = vbFalse
    
    Else
        
        'Senão Desablita ItemCategoriaCliente
        objComboItens.ListIndex = -1
        objComboItens.Enabled = False
    
    End If
    
    Exit Sub

Erro_Carrega_ComboCategoriaItens:

    Select Case gErr
    
        Case 131622

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 167616)

    End Select

    Exit Sub

End Sub
'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_TITPAG_L
    Set Form_Load_Ocx = Me
    Caption = "Clientes Sem Relacionamento"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "RelOpClientesSemRelac"

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

Public Sub Unload(objme As Object)

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




